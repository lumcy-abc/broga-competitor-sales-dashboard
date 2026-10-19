#!/usr/bin/env python3
"""Generate long-term monthly daily trackers and monthly weekly reports.

Run after scripts/export_competitor_workbook.mjs. Reads the generated workbook's
Historical Records sheet; does not alter the TEDUH collector or its JSON history.
"""
from datetime import datetime, date, timedelta
from pathlib import Path
from collections import defaultdict
from calendar import monthrange
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

ROOT = Path.cwd()
SOURCE = ROOT / "Broga_Competitor_Sales_Analysis.xlsx"
DAILY_OUT = ROOT / "Broga_Daily_Tracking.xlsx"
WEEKLY_OUT = ROOT / "Broga_Weekly_Tracking.xlsx"
GREEN, BROWN, SAGE, CREAM = "526747", "8B6F4E", "E8EDE3", "F3EBDD"
WHITE, DARK, GRAY, LINE, PALE, AMBER, LIGHT = "FFFFFF", "263126", "666666", "D7DCCF", "DCE8D5", "F6E5C9", "F7F8F4"
TODAY = date.today()


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
    if not cur: return "No record"
    if not prev: return "No prior record"
    if cur["total"] != prev["total"]: return "Total changed — review"
    n = int(cur["signed"] - prev["signed"])
    return "No change" if n == 0 else f"{n:+d} signed"


def month_range(start, end):
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        yield y, m
        if m == 12: y, m = y + 1, 1
        else: m += 1


def month_name(y, m):
    return date(y, m, 1).strftime("%b %Y")


def group_records(records):
    groups = sorted({(r["project"], r["group"]) for r in records}, key=lambda x: (x[0].casefold(), x[1].casefold()))
    lookup = {(r["date"], (r["project"], r["group"])): r for r in records}
    by_group = defaultdict(list)
    for r in records: by_group[(r["project"], r["group"])].append(r)
    for g in by_group: by_group[g].sort(key=lambda r: r["date"])
    return groups, lookup, by_group


def write_group_headers(ws, groups, first_header_row=4):
    # Column A is explicitly labelled PROJECT NAMES per the requested design.
    ws.cell(first_header_row, 1, "PROJECT NAMES")
    ws.cell(first_header_row + 1, 1, "Unit group / layout")
    col = 2
    while col <= len(groups) + 1:
        project = groups[col - 2][0]
        end = col + 1
        while end <= len(groups) + 1 and groups[end - 2][0] == project:
            end += 1
        if end - 1 > col:
            ws.merge_cells(start_row=first_header_row, start_column=col, end_row=first_header_row, end_column=end - 1)
        ws.cell(first_header_row, col, project)
        for c in range(col, end):
            ws.cell(first_header_row + 1, c, groups[c - 2][1])
            ws.column_dimensions[get_column_letter(c)].width = 23
        col = end
    for c in range(1, len(groups) + 2):
        style_cell(ws.cell(first_header_row, c), GREEN, True, WHITE, size=9)
        style_cell(ws.cell(first_header_row + 1, c), BROWN, True, WHITE, size=9)
    ws.column_dimensions["A"].width = 24
    ws.row_dimensions[first_header_row].height = 35
    ws.row_dimensions[first_header_row + 1].height = 32


def build_previous_lookup(by_group, groups, month_start, month_end):
    # For each calendar date, cache the latest observation before that date.
    result = {}
    for g in groups:
        rows = by_group.get(g, [])
        idx = 0
        previous = None
        dt = month_start
        while dt <= month_end:
            while idx < len(rows) and rows[idx]["date"] < dt:
                previous = rows[idx]
                idx += 1
            result[(dt, g)] = previous
            dt += timedelta(days=1)
    return result



def outline_project_blocks(ws, groups, first_data_row, last_data_row, header_row=5):
    """Add a clear vertical boundary around each project's layout columns."""
    medium = Side(style="medium", color=GREEN)
    for idx, group in enumerate(groups, start=2):
        starts_project = idx == 2 or groups[idx - 3][0] != group[0]
        ends_project = idx == len(groups) + 1 or groups[idx - 1][0] != group[0]
        if starts_project or ends_project:
            col = idx
            for row in range(header_row, last_data_row + 1):
                cell = ws.cell(row, col)
                b = cell.border
                cell.border = Border(left=medium if starts_project else b.left,
                                     right=medium if ends_project else b.right,
                                     top=b.top, bottom=b.bottom)


def create_daily(records, groups, lookup, by_group):
    wb = Workbook()
    wb.remove(wb.active)
    # Daily tabs are created only for months with valid observations, plus the current month.
    # This preserves useful history without producing empty tabs for months with no records.
    data_months = {(r["date"].year, r["date"].month) for r in records}
    data_months.add((TODAY.year, TODAY.month))
    months = sorted(data_months)
    for y, m in months:
        ws = wb.create_sheet(month_name(y, m))
        last_day = TODAY.day if (y, m) == (TODAY.year, TODAY.month) else monthrange(y, m)[1] if date(y, m, 1) < TODAY.replace(day=1) else 0
        month_start = date(y, m, 1)
        month_end = date(y, m, last_day) if last_day else date(y, m, monthrange(y, m)[1])
        title(ws, len(groups) + 1, f"BROGA COMPETITOR SALES ANALYSIS | DAILY TRACKER | {month_name(y,m).upper()}",
              "Daily validation view · Sales status is combined with unit count and percentage. Daily Change compares with the previous available signed observation; it does not sum snapshots.")
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(groups)+1)
        ws.cell(3, 1, "A new monthly tab is generated automatically when this report runs in a new month. Historical month tabs are retained in the regenerated workbook.")
        ws.cell(3, 1).font = Font(italic=True, size=9, color=DARK); ws.cell(3,1).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[3].height = 29
        write_group_headers(ws, groups, 5)
        previous_lookup = build_previous_lookup(by_group, groups, month_start, month_end)
        rr = 7
        for day_num in range(1, month_end.day + 1):
            dt = date(y, m, day_num)
            ws.cell(rr, 1, dt.strftime("%d %b %Y") + " · Sales Status")
            ws.cell(rr + 1, 1, "Daily Change")
            style_cell(ws.cell(rr, 1), BROWN, True, WHITE, "left")
            style_cell(ws.cell(rr + 1, 1), CREAM, True, DARK, "left")
            for j, g in enumerate(groups, 2):
                cur = lookup.get((dt, g))
                prev = previous_lookup.get((dt, g))
                val = status(cur) if cur else ("No signed observation" if dt <= TODAY else "Not yet reached")
                change = delta(cur, prev)
                style_cell(ws.cell(rr, j, val), SAGE if cur else LIGHT)
                fill = PALE if change not in ("No change", "No prior record", "No record") and "review" not in change else AMBER if "review" in change else LIGHT
                style_cell(ws.cell(rr + 1, j, change if cur else ("No observation" if dt <= TODAY else "Not yet reached")), fill, True)
            ws.row_dimensions[rr].height = 32; ws.row_dimensions[rr+1].height = 27
            rr += 2
        ws.freeze_panes = "B7"
        ws.print_title_rows = "1:6"
        ws.auto_filter.ref = f"A6:{get_column_letter(len(groups)+1)}{rr-1}"
        outline_project_blocks(ws, groups, 7, rr - 1, header_row=5)
    wb.active = len(wb.sheetnames) - 1
    wb.save(DAILY_OUT)


def month_week_periods(y, m):
    first = date(y, m, 1)
    last = date(y, m, monthrange(y, m)[1])
    # Start with the first complete Monday–Sunday reporting week that begins
    # inside this calendar month (e.g. October 2026 starts on Monday 5 Oct).
    monday = first + timedelta(days=(7 - first.weekday()) % 7)
    periods = []
    start = monday
    while start <= last:
        end = start + timedelta(days=6)
        periods.append((start, end))
        start += timedelta(days=7)
    return periods


def create_weekly(records, groups, lookup, by_group):
    wb = Workbook(); wb.remove(wb.active)
    # Weekly comparisons begin in October 2026. Create a monthly tab only when
    # that month has valid records; future tabs appear automatically after data arrives.
    weekly_start = date(2026, 10, 1)
    months = sorted({(r["date"].year, r["date"].month) for r in records
                     if r["date"] >= weekly_start})
    if not months:
        raise ValueError("No valid signed observations on or after 1 October 2026; weekly report has not been created.")
    for y, m in months:
        ws = wb.create_sheet(month_name(y, m))
        title(ws, len(groups)+1, f"BROGA COMPETITOR SALES ANALYSIS | WEEKLY TRACKER | {month_name(y,m).upper()}",
              "Each week uses the latest valid signed observation in its Monday–Sunday period. Weekly Change appears inside every week block and compares against the previous week's snapshot. No Not Yet Signed row is included.")
        ws.merge_cells(start_row=3,start_column=1,end_row=3,end_column=len(groups)+1)
        ws.cell(3,1,"Project names are merged across their unit-group columns. A changed total denominator is flagged for review; missing records are not treated as zero.")
        ws.cell(3,1).font=Font(italic=True,size=9,color=DARK); ws.cell(3,1).alignment=Alignment(wrap_text=True,vertical="center")
        ws.row_dimensions[3].height=29
        write_group_headers(ws, groups, 5)
        rr=7
        periods=month_week_periods(y,m)
        for wi,(start,end) in enumerate(periods):
            period_end=min(end,TODAY)
            ws.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=len(groups)+1)
            c=ws.cell(rr,1,f"WEEK · {start.strftime('%-d %b %Y')} – {end.strftime('%-d %b %Y')}")
            c.fill=PatternFill("solid",fgColor=BROWN); c.font=Font(bold=True,color=WHITE,size=10); c.alignment=Alignment(vertical="center")
            ws.row_dimensions[rr].height=24; rr+=1
            for metric,kind in [("Sales Status","status"),("Weekly Change","change")]:
                style_cell(ws.cell(rr,1,metric), SAGE if kind=="status" else CREAM, True, DARK, "left")
                for j,g in enumerate(groups,2):
                    valid=[r for r in by_group.get(g,[]) if start<=r["date"]<=period_end] if start<=TODAY else []
                    cur=max(valid,key=lambda x:x["date"]) if valid else None
                    prev_start=start-timedelta(days=7); prev_end=start-timedelta(days=1)
                    prior=[r for r in by_group.get(g,[]) if prev_start<=r["date"]<=prev_end]
                    prev=max(prior,key=lambda x:x["date"]) if prior else None
                    if kind=="status":
                        val=status(cur) if cur else ("In progress" if start<=TODAY<=end else "Future week" if start>TODAY else "No record this week")
                        fill=SAGE if cur else CREAM
                    else:
                        if start>TODAY: val="Future week"
                        elif not cur: val="In progress" if start<=TODAY<=end else "No record this week"
                        elif not prev: val="Baseline / no prior-week record"
                        else: val=delta(cur,prev)
                        fill=PALE if (str(val).startswith("+") or val=="No change") else AMBER if "review" in str(val).lower() else CREAM
                    style_cell(ws.cell(rr,j,val),fill,kind=="change")
                ws.row_dimensions[rr].height=32 if kind=="status" else 28; rr+=1
        ws.merge_cells(start_row=rr+1,start_column=1,end_row=rr+1,end_column=len(groups)+1)
        c=ws.cell(rr+1,1,"Monthly tabs are generated from available history. Weeks are Monday–Sunday; weeks spanning month boundaries are shown in each month they overlap.")
        c.fill=PatternFill("solid",fgColor=CREAM); c.font=Font(italic=True,color="8B5E34",size=9); c.alignment=Alignment(wrap_text=True,vertical="center")
        ws.row_dimensions[rr+1].height=30
        ws.freeze_panes="B7"; ws.print_title_rows="1:6"
        outline_project_blocks(ws, groups, 7, rr + 1, header_row=5)
    wb.active=len(wb.sheetnames)-1
    wb.save(WEEKLY_OUT)


def main():
    records=read_records()
    if not records: raise ValueError("No valid signed records found in Historical Records.")
    groups, lookup, by_group = group_records(records)
    create_daily(records, groups, lookup, by_group)
    create_weekly(records, groups, lookup, by_group)
    daily_tabs = load_workbook(DAILY_OUT, read_only=True).sheetnames
    weekly_tabs = load_workbook(WEEKLY_OUT, read_only=True).sheetnames
    print(f"Created {DAILY_OUT.name} with {len(daily_tabs)} tabs ({', '.join(daily_tabs)})")
    print(f"Created {WEEKLY_OUT.name} with {len(weekly_tabs)} tabs ({', '.join(weekly_tabs)}), starting from October 2026")
    print(f"Source: {len(records)} valid signed observations across {len(groups)} project/layout groups.")

if __name__ == "__main__":
    main()
