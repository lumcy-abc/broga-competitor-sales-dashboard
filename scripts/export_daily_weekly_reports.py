#!/usr/bin/env python3
"""Generate long-term monthly daily trackers and monthly weekly reports.

Run after scripts/export_competitor_workbook.mjs. Reads the generated workbook's
Historical Records sheet; does not alter the TEDUH collector or its JSON history.
"""
from datetime import datetime, date, timedelta
from pathlib import Path
from collections import defaultdict
from calendar import monthrange
from zoneinfo import ZoneInfo
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

ROOT = Path.cwd()
SOURCE = ROOT / "Broga_Competitor_Sales_Analysis.xlsx"
DAILY_OUT = ROOT / "Broga_Daily_Tracking.xlsx"
WEEKLY_OUT = ROOT / "Broga_Weekly_Tracking.xlsx"
GREEN, BROWN, SAGE, CREAM = "526747", "8B6F4E", "E8EDE3", "F3EBDD"
WHITE, DARK, GRAY, LINE, PALE, AMBER, LIGHT = "FFFFFF", "263126", "666666", "D7DCCF", "DCE8D5", "F6E5C9", "F7F8F4"
TODAY = datetime.now(ZoneInfo("Asia/Kuala_Lumpur")).date()


def read_records():
    if not SOURCE.exists():
        raise FileNotFoundError(f"Missing {SOURCE.name}. Run the existing Excel export first.")
    book = load_workbook(SOURCE, read_only=True, data_only=True)
    if "Historical Records" not in book.sheetnames:
        raise ValueError("Source workbook has no Historical Records sheet.")
    sheet = book["Historical Records"]
    headers = [c.value for c in next(sheet.iter_rows(min_row=1, max_row=1))]
    out = {}
    for values in sheet.iter_rows(min_row=2, values_only=True):
        rec = dict(zip(headers, values))
        raw_date = rec.get("Record Date")
        try:
            dt = raw_date.date() if isinstance(raw_date, datetime) else datetime.strptime(str(raw_date), "%d/%m/%Y").date()
            signed, total = float(rec.get("Signed Units")), float(rec.get("Total Units"))
        except (ValueError, TypeError, AttributeError):
            continue
        if str(rec.get("Status Type", "")).lower() != "signed" or total <= 0:
            continue
        project = str(rec.get("Project") or "").strip()
        group = str(rec.get("Unit Group / Layout") or "").strip()
        if not project or not group:
            continue
        key = (dt, project, group)
        out[key] = {"date": dt, "project": project, "group": group, "signed": signed, "total": total}
    return sorted(out.values(), key=lambda r: (r["date"], r["project"], r["group"]))


def style_cell(cell, fill=WHITE, bold=False, color=DARK, horizontal="center", size=9):
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(name="Aptos", size=size, bold=bold, color=color)
    cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=True)
    thin = Side(style="thin", color=LINE)
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def title(ws, last_col, heading, subtitle):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    c = ws.cell(1, 1, heading); c.fill = PatternFill("solid", fgColor=GREEN)
    c.font = Font(name="Aptos Display", size=15, bold=True, color=WHITE)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 32
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    c = ws.cell(2, 1, subtitle); c.fill = PatternFill("solid", fgColor=CREAM)
    c.font = Font(italic=True, color="8B5E34", size=10)
    c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 27
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def status(r):
    return f"{round(r['signed'] / r['total'] * 100)}% Signed ({int(r['signed'])}/{int(r['total'])} units)"


def delta(cur, prev):
    """Describe the change against the previous valid observation, including its gap."""
    if not cur: return "No record"
    if not prev: return "No prior record"
    gap = (cur["date"] - prev["date"]).days
    gap_label = f"over {gap} day(s)" if gap > 1 else "over 1 day"
    if cur["total"] != prev["total"]:
        return f"Total changed — review ({prev['date']:%d %b} → {cur['date']:%d %b}; {gap_label})"
    n = int(cur["signed"] - prev["signed"])
    change = "No change" if n == 0 else f"{n:+d} signed"
    return f"{change} ({prev['date']:%d %b} → {cur['date']:%d %b}; {gap_label})"


def month_range(start, end):
