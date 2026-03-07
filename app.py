from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Optional

import pandas as pd
from flask import Flask, render_template, request, send_file

app = Flask(__name__)


@dataclass
class ColumnMap:
    order_number: int = 4   # E
    order_date: int = 3     # D
    order_customer: int = 6 # G
    order_part: int = 14    # O
    order_net_sales: int = 20  # U

    quote_number: int = 0   # A
    quote_line: int = 1     # B
    quote_part: int = 2     # C
    quote_customer: int = 35  # AJ
    quote_date: int = 48      # AW
    quote_user: int = 61      # BJ


COLUMNS = ColumnMap()


def _safe_get(df: pd.DataFrame, idx: int, name: str) -> pd.Series:
    if idx >= df.shape[1]:
        raise ValueError(
            f"Expected column {name} at index {idx} (Excel letter position), "
            f"but file only has {df.shape[1]} columns."
        )
    return df.iloc[:, idx]


def load_orders(file_obj) -> pd.DataFrame:
    raw = pd.read_excel(file_obj)
    orders = pd.DataFrame(
        {
            "order_number": _safe_get(raw, COLUMNS.order_number, "Order Number (E)"),
            "order_date": pd.to_datetime(_safe_get(raw, COLUMNS.order_date, "Order Date (D)"), errors="coerce"),
            "customer_id": _safe_get(raw, COLUMNS.order_customer, "Customer ID (G)").astype(str).str.strip(),
            "part_number": _safe_get(raw, COLUMNS.order_part, "Part Number (O)").astype(str).str.strip().str.upper(),
            "net_sales": pd.to_numeric(_safe_get(raw, COLUMNS.order_net_sales, "Net Sales (U)"), errors="coerce").fillna(0.0),
        }
    )
    orders = orders.dropna(subset=["order_date"]).copy()
    orders = orders[orders["part_number"].ne("")]
    return orders


def load_quotes(file_obj) -> pd.DataFrame:
    raw = pd.read_excel(file_obj)
    quotes = pd.DataFrame(
        {
            "quote_number": _safe_get(raw, COLUMNS.quote_number, "Quote Number (A)").astype(str).str.strip(),
            "line_number": _safe_get(raw, COLUMNS.quote_line, "Line Number (B)"),
            "part_number": _safe_get(raw, COLUMNS.quote_part, "Part Number (C)").astype(str).str.strip().str.upper(),
            "customer_id": _safe_get(raw, COLUMNS.quote_customer, "Customer ID (AJ)").astype(str).str.strip(),
            "quote_date": pd.to_datetime(_safe_get(raw, COLUMNS.quote_date, "Date Quoted (AW)"), errors="coerce"),
            "user_id": _safe_get(raw, COLUMNS.quote_user, "User ID (BJ)").astype(str).str.strip(),
        }
    )
    quotes = quotes.dropna(subset=["quote_date"]).copy()
    quotes = quotes[quotes["part_number"].ne("")]
    return quotes


def run_conversion(
    orders: pd.DataFrame,
    quotes: pd.DataFrame,
    start_date: Optional[pd.Timestamp],
    end_date: Optional[pd.Timestamp],
    window_days: int,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if start_date is not None:
        orders = orders[orders["order_date"] >= start_date]
        quotes = quotes[quotes["quote_date"] >= start_date]
    if end_date is not None:
        orders = orders[orders["order_date"] <= end_date]
        quotes = quotes[quotes["quote_date"] <= end_date]

    if quotes.empty:
        return quotes.assign(converted=False), pd.DataFrame()

    merged = quotes.merge(
        orders,
        on=["customer_id", "part_number"],
        how="left",
        suffixes=("_quote", "_order"),
    )

    merged["days_to_convert"] = (merged["order_date"] - merged["quote_date"]).dt.days
    merged["valid_conversion"] = (
        merged["days_to_convert"].notna()
        & (merged["days_to_convert"] >= 0)
        & (merged["days_to_convert"] <= window_days)
    )

    # First conversion event per quote line
    converted_events = (
        merged[merged["valid_conversion"]]
        .sort_values("days_to_convert")
        .groupby(["quote_number", "line_number"], as_index=False)
        .first()
    )

    output = quotes.merge(
        converted_events[
            [
                "quote_number",
                "line_number",
                "order_number",
                "order_date",
                "net_sales",
                "days_to_convert",
            ]
        ],
        on=["quote_number", "line_number"],
        how="left",
    )

    output["converted"] = output["order_number"].notna()

    summary = (
        output.groupby("user_id", dropna=False)
        .agg(
            quote_lines=("quote_number", "count"),
            converted_lines=("converted", "sum"),
            converted_net_sales=("net_sales", "sum"),
        )
        .reset_index()
    )
    summary["conversion_rate"] = (summary["converted_lines"] / summary["quote_lines"]).fillna(0)

    return output, summary


@app.route("/", methods=["GET", "POST"])
def index():
    line_results_html = None
    rep_results_html = None
    error = None

    if request.method == "POST":
        try:
            order_file = request.files.get("order_file")
            quote_file = request.files.get("quote_file")
            start = request.form.get("start_date")
            end = request.form.get("end_date")
            window = int(request.form.get("window_days") or 90)

            if not order_file or not quote_file:
                raise ValueError("Please upload both an order log file and a quote summary file.")

            orders = load_orders(order_file)
            quotes = load_quotes(quote_file)

            start_date = pd.to_datetime(start) if start else None
            end_date = pd.to_datetime(end) if end else None

            line_results, rep_summary = run_conversion(orders, quotes, start_date, end_date, window)

            line_results_html = line_results.sort_values(["quote_date", "quote_number", "line_number"]).to_html(
                index=False, classes="results-table"
            )
            rep_results_html = rep_summary.sort_values("conversion_rate", ascending=False).to_html(
                index=False, classes="results-table"
            )

            app.config["last_line_results"] = line_results
            app.config["last_rep_results"] = rep_summary
        except Exception as exc:  # show user-friendly message
            error = str(exc)

    return render_template(
        "index.html",
        line_results_html=line_results_html,
        rep_results_html=rep_results_html,
        error=error,
    )


@app.route("/download/<report_type>")
def download(report_type: str):
    line_results: pd.DataFrame = app.config.get("last_line_results", pd.DataFrame())
    rep_results: pd.DataFrame = app.config.get("last_rep_results", pd.DataFrame())

    if report_type == "line":
        df = line_results
        filename = "quote_conversion_line_detail.xlsx"
    elif report_type == "rep":
        df = rep_results
        filename = "quote_conversion_rep_summary.xlsx"
    else:
        return "Unknown report type", 404

    if df.empty:
        return "No report available yet. Run an analysis first.", 400

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/favicon.ico")
def favicon():
    return send_file("assets/app.ico", mimetype="image/x-icon")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
