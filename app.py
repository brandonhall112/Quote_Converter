from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import re
from threading import Timer
from typing import Any
import sys
import webbrowser

import pandas as pd
from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook

BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))


@dataclass
class ColumnMap:
    order_number: int = 4  # E
    order_date: int = 3  # D
    order_customer: int = 6  # G
    order_part: int = 14  # O
    order_net_sales: int = 20  # U

    quote_number: int = 0  # A
    quote_line: int = 1  # B
    quote_part: int = 2  # C
    quote_customer: int = 35  # AJ
    quote_customer_name: int = 36  # AK
    quote_date: int = 48  # AW
    quote_user: int = 61  # BJ


COLUMNS = ColumnMap()


def _safe_get(df: pd.DataFrame, idx: int, name: str) -> pd.Series:
    if idx >= df.shape[1]:
        raise ValueError(
            f"Expected column {name} at index {idx} (Excel letter position), "
            f"but file only has {df.shape[1]} columns."
        )
    return df.iloc[:, idx]


def load_orders(file_obj: Any) -> pd.DataFrame:
    raw = pd.read_excel(file_obj)
    orders = pd.DataFrame(
        {
            "order_number": _safe_get(raw, COLUMNS.order_number, "Order Number (E)").astype(str).str.strip(),
            "order_date": pd.to_datetime(_safe_get(raw, COLUMNS.order_date, "Order Date (D)"), errors="coerce"),
            "customer_id": _safe_get(raw, COLUMNS.order_customer, "Customer ID (G)").astype(str).str.strip(),
            "part_number": _safe_get(raw, COLUMNS.order_part, "Part Number (O)").astype(str).str.strip().str.upper(),
            "net_sales": pd.to_numeric(_safe_get(raw, COLUMNS.order_net_sales, "Net Sales (U)"), errors="coerce").fillna(0.0),
        }
    )
    orders = orders.dropna(subset=["order_date"]).copy()
    return orders[(orders["part_number"].ne("")) & (orders["customer_id"].ne(""))]


def load_quotes(file_obj: Any) -> pd.DataFrame:
    raw = pd.read_excel(file_obj)
    quotes = pd.DataFrame(
        {
            "quote_number": _safe_get(raw, COLUMNS.quote_number, "Quote Number (A)").astype(str).str.strip(),
            "line_number": _safe_get(raw, COLUMNS.quote_line, "Line Number (B)"),
            "part_number": _safe_get(raw, COLUMNS.quote_part, "Part Number (C)").astype(str).str.strip().str.upper(),
            "customer_id": _safe_get(raw, COLUMNS.quote_customer, "Customer ID (AJ)").astype(str).str.strip(),
            "customer_name": _safe_get(raw, COLUMNS.quote_customer_name, "Customer Name (AK)").astype(str).str.strip(),
            "quote_date": pd.to_datetime(_safe_get(raw, COLUMNS.quote_date, "Date Quoted (AW)"), errors="coerce"),
            "user_id": _safe_get(raw, COLUMNS.quote_user, "User ID (BJ)").astype(str).str.strip(),
        }
    )
    quotes = quotes.dropna(subset=["quote_date"]).copy()
    return quotes[(quotes["part_number"].ne("")) & (quotes["quote_number"].ne(""))]


def run_conversion(orders: pd.DataFrame, quotes: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if quotes.empty:
        empty = quotes.assign(converted=False)
        return empty, pd.DataFrame(), pd.DataFrame()

    merged = quotes.merge(
        orders,
        on=["customer_id", "part_number"],
        how="left",
        suffixes=("_quote", "_order"),
    )

    merged["days_to_convert"] = (merged["order_date"] - merged["quote_date"]).dt.days
    merged["valid_conversion"] = merged["days_to_convert"].notna() & (merged["days_to_convert"] >= 0)

    converted_events = (
        merged[merged["valid_conversion"]]
        .sort_values("days_to_convert")
        .groupby(["quote_number", "line_number"], as_index=False)
        .first()
    )

    line_output = quotes.merge(
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
    line_output["converted"] = line_output["order_number"].notna()

    rep_summary = (
        line_output.groupby("user_id", dropna=False)
        .agg(
            quote_lines=("quote_number", "count"),
            converted_lines=("converted", "sum"),
            converted_net_sales=("net_sales", "sum"),
        )
        .reset_index()
    )
    rep_summary["conversion_rate"] = (rep_summary["converted_lines"] / rep_summary["quote_lines"]).fillna(0)

    def _first_non_empty(series: pd.Series) -> str:
        for value in series:
            if pd.notna(value) and str(value).strip() != "":
                return str(value).strip()
        return ""

    quote_output = (
        line_output.groupby(["quote_number"], dropna=False)
        .agg(
            customer_id=("customer_id", _first_non_empty),
            customer_name=("customer_name", _first_non_empty),
            user_id=("user_id", _first_non_empty),
            quote_date=("quote_date", "min"),
            parts_quoted=("part_number", "nunique"),
            matched_orders=("order_number", lambda s: ", ".join(sorted({v for v in s.dropna().astype(str) if v}))),
            converted_net_sales=("net_sales", "sum"),
            converted_lines=("converted", "sum"),
            total_lines=("line_number", "count"),
        )
        .reset_index()
    )
    quote_output["converted"] = quote_output["converted_lines"] > 0
    quote_output["follow_up_needed"] = ~quote_output["converted"]

    return line_output, rep_summary, quote_output


def _normalize_header(value: Any) -> str:
    return str(value or "").strip().lower().replace("_", " ")


def _find_header_row(ws) -> int:
    for row_num in range(1, min(ws.max_row, 30) + 1):
        labels = [_normalize_header(c.value) for c in ws[row_num] if c.value is not None]
        if any("quote" in v for v in labels) and (any("customer" in v for v in labels) or any("date" in v for v in labels)):
            return row_num
    return 1


def _column_map_from_header(ws, header_row: int) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        label = _normalize_header(ws.cell(row=header_row, column=col).value)
        if not label:
            continue
        if (("quote" in label and "#" in label) or "quote number" in label or label == "quote"):
            mapping.setdefault("quote_number", col)
        elif "customer" in label:
            mapping.setdefault("customer_id", col)
        elif "date" in label:
            mapping.setdefault("quote_date", col)
        elif "rep" in label or "user" in label or "owner" in label or "assignee" in label:
            mapping.setdefault("user_id", col)
        elif "quote amount" in label or ("amount" in label and "order" not in label):
            mapping.setdefault("quote_amount", col)
        elif "part" in label and ("count" in label or "qty" in label or "quoted" in label):
            mapping.setdefault("parts_quoted", col)
        elif "won/lost" in label or ("won" in label and "lost" in label):
            mapping.setdefault("follow_up_needed", col)
        elif "actual order" in label:
            mapping.setdefault("converted_net_sales", col)
        elif "order" in label:
            mapping.setdefault("matched_orders", col)
        elif "net" in label or "sales" in label:
            mapping.setdefault("converted_net_sales", col)
        elif "status" in label or "follow" in label:
            mapping.setdefault("follow_up_needed", col)
        elif "note" in label:
            mapping.setdefault("notes", col)
    return mapping


def _clear_sheet_data(ws, start_row: int) -> None:
    if ws.max_row <= start_row:
        return

    first_data_row = start_row + 1
    last_data_row = ws.max_row

    if ws._tables:
        table = next(iter(ws._tables.values()))
        _, last_cell = table.ref.split(":")
        last_data_row = ws[last_cell].row

    for row in range(first_data_row, last_data_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if cell.data_type == "f":
                continue
            cell.value = None


def _normalize_person(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (value or "").lower())


def _display_name_from_sheet(sheet_title: str) -> str:
    return sheet_title.split("(")[0].strip()


def _assign_sheet_for_rep(workbook, rep: str):
    non_summary = [ws for ws in workbook.worksheets if "summary" not in ws.title.lower()]

    rep_clean = _normalize_person(rep)
    if not rep_clean:
        return non_summary[0] if non_summary else workbook.worksheets[0]

    for ws in non_summary:
        if _normalize_person(ws.title) == rep_clean:
            return ws

    for ws in non_summary:
        ws_clean = _normalize_person(ws.title)
        if rep_clean in ws_clean or ws_clean in rep_clean:
            return ws

    # Match user IDs like bschrader to sheet names like Brent Schrader by last name.
    for ws in non_summary:
        ws_parts = [p for p in re.split(r"\s+", ws.title.strip()) if p]
        if ws_parts:
            last_name = _normalize_person(ws_parts[-1])
            if last_name and rep_clean.endswith(last_name):
                return ws

    # Master Allocation should be fallback only when no named tab could be identified.
    for ws in non_summary:
        if "master" not in ws.title.lower():
            return ws

    for ws in non_summary:
        if rep_clean and rep_clean in _normalize_person(ws.title):
            return ws
    return non_summary[0] if non_summary else workbook.worksheets[0]


def build_follow_up_workbook(template_bytes: bytes, quotes_for_followup: pd.DataFrame) -> bytes:
    wb = load_workbook(BytesIO(template_bytes))

    quotes_for_followup = quotes_for_followup.copy()
    if quotes_for_followup.empty:
        out = BytesIO()
        wb.save(out)
        return out.getvalue()

    target_sheets = {row.user_id: _assign_sheet_for_rep(wb, row.user_id) for row in quotes_for_followup[["user_id"]].drop_duplicates().itertuples()}

    for ws in set(target_sheets.values()):
        header_row = _find_header_row(ws)
        _clear_sheet_data(ws, header_row)

    for rep, rep_df in quotes_for_followup.groupby("user_id", dropna=False):
        ws = target_sheets.get(rep)
        if ws is None:
            continue

        header_row = _find_header_row(ws)
        col_map = _column_map_from_header(ws, header_row)

        if not col_map:
            continue

        row_num = header_row + 1

        table_last_row = ws.max_row
        if ws._tables:
            table = next(iter(ws._tables.values()))
            _, last_cell = table.ref.split(":")
            table_last_row = ws[last_cell].row

        for rec in rep_df.sort_values(["quote_date", "quote_number"]).to_dict("records"):
            if row_num > table_last_row:
                break

            assignee_name = _display_name_from_sheet(ws.title)
            if "quote_number" in col_map:
                ws.cell(row=row_num, column=col_map["quote_number"], value=rec.get("quote_number"))
            if "customer_id" in col_map:
                ws.cell(
                    row=row_num,
                    column=col_map["customer_id"],
                    value=rec.get("customer_name") or rec.get("customer_id"),
                )
            if "quote_date" in col_map:
                ws.cell(row=row_num, column=col_map["quote_date"], value=rec.get("quote_date"))
            if "user_id" in col_map:
                ws.cell(row=row_num, column=col_map["user_id"], value=assignee_name)
            if "quote_amount" in col_map:
                ws.cell(row=row_num, column=col_map["quote_amount"], value=None)
            if "parts_quoted" in col_map:
                ws.cell(row=row_num, column=col_map["parts_quoted"], value=int(rec.get("parts_quoted") or 0))
            if "converted_net_sales" in col_map:
                ws.cell(row=row_num, column=col_map["converted_net_sales"], value=float(rec.get("converted_net_sales") or 0))
            if "follow_up_needed" in col_map:
                ws.cell(
                    row=row_num,
                    column=col_map["follow_up_needed"],
                    value="Lost" if rec.get("follow_up_needed") else "Won",
                )
            if "matched_orders" in col_map:
                ws.cell(row=row_num, column=col_map["matched_orders"], value=rec.get("matched_orders"))
            if "notes" in col_map:
                ws.cell(row=row_num, column=col_map["notes"], value="")
            row_num += 1

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def _resolve_template_bytes(uploaded_template) -> bytes:
    if uploaded_template and uploaded_template.filename:
        return uploaded_template.read()

    local_template = BASE_DIR / "assets" / "Parts Follow Up Template.xlsx"
    if local_template.exists():
        return local_template.read_bytes()

    raise ValueError(
        "Template file is required. Upload 'Parts Follow Up Template.xlsx' in the form, "
        "or place it in assets/Parts Follow Up Template.xlsx."
    )


@app.route("/", methods=["GET", "POST"])
def index():
    quote_results_html = None
    rep_results_html = None
    error = None

    if request.method == "POST":
        try:
            order_file = request.files.get("order_file")
            quote_file = request.files.get("quote_file")
            template_file = request.files.get("template_file")

            if not order_file or not quote_file:
                raise ValueError("Please upload both an order log file and a quote summary file.")

            template_bytes = _resolve_template_bytes(template_file)
            orders = load_orders(order_file)
            quotes = load_quotes(quote_file)

            line_results, rep_summary, quote_results = run_conversion(orders, quotes)
            follow_up_quotes = quote_results[quote_results["follow_up_needed"]].copy()
            generated_report = build_follow_up_workbook(template_bytes, follow_up_quotes)

            quote_results_html = quote_results.sort_values(["quote_date", "quote_number"]).to_html(
                index=False,
                classes="results-table",
            )
            rep_results_html = rep_summary.sort_values("conversion_rate", ascending=False).to_html(
                index=False,
                classes="results-table",
            )

            app.config["last_quote_results"] = quote_results
            app.config["last_rep_results"] = rep_summary
            app.config["last_followup_workbook"] = generated_report
        except Exception as exc:
            error = str(exc)

    return render_template(
        "index.html",
        quote_results_html=quote_results_html,
        rep_results_html=rep_results_html,
        error=error,
    )


@app.route("/download/followup")
def download_followup():
    workbook_bytes = app.config.get("last_followup_workbook")
    if not workbook_bytes:
        return "No report available yet. Run an analysis first.", 400

    return send_file(
        BytesIO(workbook_bytes),
        as_attachment=True,
        download_name="Parts_Follow_Up_Output.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/favicon.ico")
def favicon():
    return send_file(BASE_DIR / "assets" / "app.ico", mimetype="image/x-icon")


def _open_browser() -> None:
    webbrowser.open("http://127.0.0.1:8000")


if __name__ == "__main__":
    Timer(1.0, _open_browser).start()
    app.run(host="127.0.0.1", port=8000, debug=False, use_reloader=False)
