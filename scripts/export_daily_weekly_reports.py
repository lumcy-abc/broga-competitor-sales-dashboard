#!/usr/bin/env python3
"""Generate long-term monthly daily trackers and monthly weekly reports.

Run after scripts/export_competitor_workbook.mjs. Reads the generated workbook's
Historical Records sheet; does not alter the TEDUH collector or its JSON history.
"""
from datetime import datetime, date, timedelta
from pathlib import Path
from collections import defaultdict
from copy import copy
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


def canonical_group(project, group):
    """Normalize legacy labels without merging unrelated unit groups."""
    project_key = str(project or "").strip().casefold()
    group_text = str(group or "").strip()
    if project_key == "springvale sempurna heights" and group_text.casefold() == "group 1":
        return "Group B · rear-kitchen layout"
    if project_key == "amani gallery homes" and group_text.casefold() in ("terrace houses", "terrace house"):
        return "20' x 70' / 20' x 72'"
    return group_text


def read_records():
    if not SOURCE.exists():
        raise FileNotFoundError(f"Missing {SOURCE.name}. Run the existing Excel export first.")
    book = load_workbook(SOURCE, read_only=True, data_only=True)
    if "Historical Records" not in book.sheetnames:
        raise ValueError("Source workbook has no Historical Records sheet.")
    sheet = book["Historical Records"]
    headers = [c.value for c in next(sheet.iter_rows(min_row=1, max_row=1))]
    out = {}
    skipped_dates = skipped_status = skipped_fields = duplicates = 0
    for row_num, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        rec = dict(zip(headers, values))
        raw_date = rec.get("Record Date")
        try:
            if isinstance(raw_date, datetime):
                dt = raw_date.date()
            elif isinstance(raw_date, date):
                dt = raw_date
            else:
                raw = str(raw_date).strip()
                try:
                    dt = datetime.strptime(raw, "%d/%m/%Y").date()
                except ValueError:
                    dt = datetime.strptime(raw[:10], "%Y-%m-%d").date()
        except (ValueError, TypeError, AttributeError):
            skipped_dates += 1
            continue
        status_type = str(rec.get("Status Type", "")).strip().lower()
        if status_type != "signed":
            if status_type not in ("booked", "yet to launch", "yet to sign", "unclassified", ""):
                print(f"WARNING: row {row_num}: unknown status type {status_type!r}; excluded from signed tracker.")
            skipped_status += 1
            continue
        try:
            signed, total = float(rec.get("Signed Units")), float(rec.get("Total Units"))
        except (ValueError, TypeError):
            skipped_fields += 1
            print(f"WARNING: row {row_num}: signed record has invalid unit counts; skipped.")
            continue
        if total <= 0 or signed < 0 or signed > total:
            skipped_fields += 1
            print(f"WARNING: row {row_num}: invalid counts signed={signed}, total={total}; skipped.")
            continue
        project = str(rec.get("Project") or "").strip()
        group = canonical_group(project, str(rec.get("Unit Group / Layout") or "").strip())
        if not project or not group:
            skipped_fields += 1
            print(f"WARNING: row {row_num}: missing project/layout; skipped.")
            continue
        key = (dt, project, group)
        candidate = {"date": dt, "project": project, "group": group, "signed": signed, "total": total}
        if key in out:
            duplicates += 1
            if out[key]["signed"] == signed and out[key]["total"] == total:
                print(f"WARNING: identical duplicate dated group row {key!r}; duplicate ignored (source row {row_num}).")
                continue
            raise ValueError(
                f"Conflicting duplicate dated group row {key!r}: existing signed/total="
                f"{out[key]['signed']}/{out[key]['total']}, new={signed}/{total} at source row {row_num}. "
                "Resolve the source conflict instead of silently choosing one value."
            )
        out[key] = candidate
    print(f"Source validation: {len(out)} unique signed rows; skipped {skipped_dates} invalid dates, {skipped_status} non-signed/unknown statuses, {skipped_fields} invalid signed rows; {duplicates} duplicate keys.")
    return sorted(out.values(), key=lambda r: (r["date"], r["project"], r["group"]))


def read_history_rows():
    """Read every dated row in Historical Records, including non-signed/unknown statuses.

    Trend calculations still use only valid signed observations, but Snapshot History
    must expose all available source dates/statuses for auditing and date filtering.
    """
    book = load_workbook(SOURCE, read_only=True, data_only=True)
    sheet = book["Historical Records"]
    headers = [c.value for c in next(sheet.iter_rows(min_row=1, max_row=1))]
    rows = []
    for row_num, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        rec = dict(zip(headers, values))
        raw_date = rec.get("Record Date")
        try:
            if isinstance(raw_date, datetime): dt = raw_date.date()
            elif isinstance(raw_date, date): dt = raw_date
            else:
                raw = str(raw_date).strip()
                try: dt = datetime.strptime(raw, "%d/%m/%Y").date()
                except ValueError: dt = datetime.strptime(raw[:10], "%Y-%m-%d").date()
        except (ValueError, TypeError, AttributeError):
            print(f"WARNING: history row {row_num} has an invalid date and cannot appear in date filtering: {raw_date!r}")
            continue
        project = str(rec.get("Project") or "").strip()
        group = canonical_group(project, str(rec.get("Unit Group / Layout") or "").strip())
        status_type = str(rec.get("Status Type") or "Unclassified").strip()
        if not project or not group:
            print(f"WARNING: history row {row_num} has no project/layout label; retained with a placeholder.")
            project = project or "Unidentified project"
            group = group or "Unidentified unit group"
        try: signed = float(rec.get("Signed Units"))
        except (ValueError, TypeError): signed = None
        try: total = float(rec.get("Total Units"))
        except (ValueError, TypeError): total = None
        rows.append({"date": dt, "project": project, "group": group, "status": status_type,
                     "signed": signed, "total": total,
                     "raw": str(rec.get("Raw Recorded Status") or ""),
                     "source": str(rec.get("Source") or "")})
        if status_type.lower() in ("", "unclassified"):
            print(f"WARNING: history row {row_num} has a blank/unclassified status ({project} / {group}, {dt}); retained for review.")
    book.close()
    return rows


def style_cell(cell, fill=WHITE, bold=False, color=DARK, horizontal="center", size=9):
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(name="Aptos", size=size, bold=bold, color=color)
    cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=True)
    thin = Side(style="thin", color=LINE)
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def title(ws, last_col, heading, subtitle=None):
    """Title-only top row; explanatory rows were removed by user request."""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    c = ws.cell(1, 1, heading); c.fill = PatternFill("solid", fgColor=GREEN)
    c.font = Font(name="Aptos Display", size=15, bold=True, color=WHITE)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 32
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


def daily_change_and_gap(cur, prev):
    if not cur:
        return "No observation", None
    if not prev:
        return "No prior observation", None
    gap = (cur["date"] - prev["date"]).days
    if cur["total"] != prev["total"]:
        return "Total changed — review scope/release", gap
    change = int(cur["signed"] - prev["signed"])
    return ("No change" if change == 0 else f"{change:+d} signed"), gap


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
    """Create a single current-month daily workbook, without explanatory rows."""
    wb = Workbook()
    ws = wb.active
    ws.title = month_name(TODAY.year, TODAY.month)
    y, m = TODAY.year, TODAY.month
    last_day = TODAY.day
    month_start = date(y, m, 1)
    month_end = TODAY
    title(ws, len(groups) + 1, f"BROGA COMPETITOR SALES ANALYSIS | DAILY TRACKER | {month_name(y,m).upper()}")
    write_group_headers(ws, groups, 2)
    previous_lookup = build_previous_lookup(by_group, groups, month_start, month_end)
    rr = 5
    for day_num in range(1, month_end.day + 1):
        dt = date(y, m, day_num)
        ws.cell(rr, 1, dt.strftime("%d %b %Y") + " · Sales Status")
        ws.cell(rr + 1, 1, "Change vs Previous Observation")
        ws.cell(rr + 2, 1, "Days Since Previous Observation")
        style_cell(ws.cell(rr, 1), BROWN, True, WHITE, "left")
        style_cell(ws.cell(rr + 1, 1), CREAM, True, DARK, "left")
        style_cell(ws.cell(rr + 2, 1), SAGE, True, DARK, "left")
        for j, g in enumerate(groups, 2):
            cur = lookup.get((dt, g))
            prev = previous_lookup.get((dt, g))
            val = status(cur) if cur else "No signed observation"
            change, gap = daily_change_and_gap(cur, prev)
            style_cell(ws.cell(rr, j, val), SAGE if cur else LIGHT)
            if not cur:
                change = "No observation"
            if "review" in change.lower() or "scope" in change.lower():
                fill = AMBER
            elif str(change).startswith("+"):
                fill = PALE
            elif str(change).startswith("-"):
                fill = "FCE4D6"
            else:
                fill = LIGHT
            style_cell(ws.cell(rr + 1, j, change), fill, True)
            style_cell(ws.cell(rr + 2, j, gap if gap is not None else "—"), LIGHT)
        ws.row_dimensions[rr].height = 32
        ws.row_dimensions[rr+1].height = 27
        ws.row_dimensions[rr+2].height = 22
        rr += 3
    ws.freeze_panes = "B4"
    ws.print_title_rows = "1:3"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(groups)+1)}{rr-1}"
    outline_project_blocks(ws, groups, 4, rr - 1, header_row=2)
    wb.save(DAILY_OUT)

def month_week_periods(y, m):
    first = date(y, m, 1)
    last = date(y, m, monthrange(y, m)[1])
    # Include every Monday–Sunday week that overlaps this month, even when the
    # Monday falls in the preceding month. The same boundary week may appear in both tabs.
    monday = first - timedelta(days=first.weekday())
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
        # Weekly tracker uses a compact one-row title, matching the daily tracker's
        # clear visual hierarchy while removing the two redundant explanatory rows.
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(groups)+1)
        c = ws.cell(1, 1, f"BROGA COMPETITOR SALES ANALYSIS | WEEKLY TRACKER | {month_name(y,m).upper()}")
        c.fill = PatternFill("solid", fgColor=GREEN)
        c.font = Font(name="Aptos Display", size=15, bold=True, color=WHITE)
        c.alignment = Alignment(vertical="center")
        ws.row_dimensions[1].height = 32
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        write_group_headers(ws, groups, 2)
        rr=4
        periods=month_week_periods(y,m)
        for wi,(start,end) in enumerate(periods):
            period_end=min(end,TODAY)
            ws.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=len(groups)+1)
            c=ws.cell(rr,1,f"WEEK · {start.day} {start:%b %Y} – {end.day} {end:%b %Y}")
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
        ws.freeze_panes="B4"; ws.print_title_rows="1:3"
        outline_project_blocks(ws, groups, 4, rr - 1, header_row=2)
    wb.active=len(wb.sheetnames)-1
    wb.save(WEEKLY_OUT)



def create_snapshot_history(records, by_group, history_rows):
    """Create a compact eight-column audit table with valid signed-trend changes."""
    wb = load_workbook(WEEKLY_OUT)
    sheet_name = "Snapshot History"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    if "Latest Snapshot Comparison" in wb.sheetnames:
        del wb["Latest Snapshot Comparison"]
    ws = wb.create_sheet(sheet_name, 0)
    headers = [
        "Observation Date", "Project", "Unit Group / Layout", "Status",
        "Units / Total", "Signed %", "Change vs Previous", "Source / Validation Notes"
    ]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    c = ws.cell(1, 1, "BROGA COMPETITOR SALES ANALYSIS | SNAPSHOT HISTORY")
    c.fill = PatternFill("solid", fgColor=GREEN)
    c.font = Font(name="Aptos Display", size=15, bold=True, color=WHITE)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 32
    header_row = 3
    for col, header in enumerate(headers, 1):
        style_cell(ws.cell(header_row, col, header), BROWN, True, WHITE, size=9)
    ws.row_dimensions[header_row].height = 34

    # One delta per project + unit group, calculated only from valid signed observations.
    trend = {}
    for group_key in sorted(by_group, key=lambda x: (x[0].casefold(), x[1].casefold())):
        observations = sorted(by_group[group_key], key=lambda x: x["date"])
        previous = None
        for current in observations:
            gap = (current["date"] - previous["date"]).days if previous else None
            signed_change = current["signed"] - previous["signed"] if previous else None
            current_pct = current["signed"] / current["total"] if current["total"] else None
            previous_pct = previous["signed"] / previous["total"] if previous and previous["total"] else None
            pp_change = ((current_pct - previous_pct) * 100
                         if current_pct is not None and previous_pct is not None else None)
            total_changed = bool(previous and current["total"] != previous["total"])
            pieces = []
            if previous is None:
                change_text = "Baseline"
                pieces.append("No earlier valid signed observation")
            else:
                change_parts = []
                if signed_change is not None:
                    change_parts.append(f"{signed_change:+g} signed")
                if pp_change is not None:
                    change_parts.append(f"{pp_change:+.1f} pp")
                change_text = " · ".join(change_parts) if change_parts else "Change unavailable"
                if total_changed:
                    diff = current["total"] - previous["total"]
                    pieces.append(f"TOTAL CHANGED: {previous['total']:g} → {current['total']:g} ({diff:+g}); percentage recalculated")
                if gap and gap > 1:
                    pieces.append(f"{gap}-day gap; not daily movement")
                if signed_change < 0:
                    pieces.append("Signed count decreased; verify source/status correction")
                if not pieces:
                    pieces.append("No scope change detected")
            trend[(current["date"], current["project"], current["group"])] = {
                "signed": current["signed"], "total": current["total"], "pct": current_pct,
                "change": change_text, "signed_change": signed_change, "pp_change": pp_change,
                "total_changed": total_changed, "notes": pieces,
            }
            previous = current

    rows_to_write = []
    written_signed_keys = set()
    for item in history_rows:
        key = (item["date"], item["project"], item["group"])
        status_lower = item["status"].strip().lower()
        valid_signed = (status_lower == "signed" and item["signed"] is not None and
                        item["total"] is not None and item["total"] > 0 and
                        0 <= item["signed"] <= item["total"])
        if valid_signed and key in trend:
            # Use the canonical signed observation once, not a duplicate source row.
            if key in written_signed_keys:
                continue
            written_signed_keys.add(key)
            t = trend[key]
            units_total = f"{t['signed']:g} / {t['total']:g}"
            signed_pct = t["pct"]
            change_text = t["change"]
            notes = [str(item.get("raw") or "Source status retained")]
            notes.extend(t["notes"])
            row_fill = AMBER if t["total_changed"] else (WHITE if len(rows_to_write) % 2 == 0 else LIGHT)
            change_color = ("16803D" if (t["signed_change"] or 0) > 0 else
                            "C00000" if (t["signed_change"] or 0) < 0 or t["total_changed"] else DARK)
            rows_to_write.append({"values": [item["date"], item["project"], item["group"], item["status"],
                                               units_total, signed_pct, change_text, "; ".join(x for x in notes if x)],
                                  "fill": row_fill, "change_color": change_color,
                                  "scope": t["total_changed"]})
        else:
            if status_lower in ("", "unclassified"):
                note = "Blank/unclassified status; verify source"
            elif status_lower != "signed":
                note = "Non-signed status retained for audit; excluded from signed trend"
            else:
                note = "Signed status has invalid/missing unit counts; verify source"
            units = item["signed"] if item["signed"] is not None else None
            total = item["total"] if item["total"] is not None else None
            unit_text = f"{units:g} / {total:g}" if units is not None and total is not None else (f"— / {total:g}" if total is not None else "—")
            pct = (units / total) if valid_signed and total else None
            rows_to_write.append({"values": [item["date"], item["project"], item["group"], item["status"],
                                               unit_text, pct, "Not applicable", f"{item.get('raw') or ''}; {note}".strip("; ")],
                                  "fill": AMBER if status_lower in ("", "unclassified") else (WHITE if len(rows_to_write) % 2 == 0 else LIGHT),
                                  "change_color": DARK, "scope": False})

    rows_to_write.sort(key=lambda x: (x["values"][0], str(x["values"][1]).casefold(),
                                      str(x["values"][2]).casefold(), str(x["values"][3]).casefold()))
    for rr, row in enumerate(rows_to_write, start=header_row + 1):
        for col, value in enumerate(row["values"], 1):
            cell = ws.cell(rr, col, value)
            style_cell(cell, row["fill"], color=DARK,
                       horizontal="left" if col in (2, 3, 8) else "center", size=9)
        ws.cell(rr, 1).number_format = "dd mmm yyyy"
        ws.cell(rr, 6).number_format = "0.0%"
        ws.cell(rr, 7).font = Font(name="Aptos", size=9, bold=True, color=row["change_color"])
        if row["scope"]:
            ws.cell(rr, 7).fill = PatternFill("solid", fgColor="F4CCCC")
            ws.cell(rr, 8).font = Font(name="Aptos", size=9, bold=True, color="C00000")
        elif row["values"][6] not in ("Baseline", "Not applicable"):
            ws.cell(rr, 7).fill = PatternFill("solid", fgColor="E2F0D9" if row["change_color"] == "16803D" else "FCE4D6" if row["change_color"] == "C00000" else "F7F8F4")
        ws.row_dimensions[rr].height = 30 if row["scope"] else 24

    last_row = max(header_row, header_row + len(rows_to_write))
    # The user requested a filter on Observation Date only to keep the table uncluttered.
    ws.auto_filter.ref = f"A{header_row}:A{last_row}"
    ws.freeze_panes = "A4"
    widths = [12, 22, 23, 10, 11, 9, 21, 30]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:3"
    wb.save(WEEKLY_OUT)

def main():
    records=read_records()
    if not records: raise ValueError("No valid signed records found in Historical Records.")
    history_rows=read_history_rows()
    groups, lookup, by_group = group_records(records)
    create_daily(records, groups, lookup, by_group)
    create_weekly(records, groups, lookup, by_group)
    create_snapshot_history(records, by_group, history_rows)
    daily_tabs = load_workbook(DAILY_OUT, read_only=True).sheetnames
    weekly_tabs = load_workbook(WEEKLY_OUT, read_only=True).sheetnames
    print(f"Created {DAILY_OUT.name} with {len(daily_tabs)} tabs ({', '.join(daily_tabs)})")
    print(f"Created {WEEKLY_OUT.name} with {len(weekly_tabs)} tabs ({', '.join(weekly_tabs)}), including date-filtered snapshot history and Monday–Sunday monthly tabs")
    print(f"Source: {len(records)} valid signed observations across {len(groups)} project/layout groups.")

if __name__ == "__main__":
    main()
