#!/usr/bin/env python3
"""Generate long-term monthly daily trackers and monthly weekly reports.

Run after scripts/export_competitor_workbook.mjs. Reads the generated workbook's
Historical Records sheet; does not alter the TEDUH collector or its JSON history.
"""
import json
import re
from datetime import datetime, date, timedelta
from pathlib import Path
from collections import defaultdict
from copy import copy
from calendar import monthrange
from zoneinfo import ZoneInfo
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

ROOT = Path.cwd()
SOURCE = ROOT / "Broga_Competitor_Sales_Analysis.xlsx"
DAILY_OUT = ROOT / "Broga_Daily_Tracking.xlsx"
WEEKLY_OUT = ROOT / "Broga_Weekly_Tracking.xlsx"
GREEN, BROWN, SAGE, CREAM = "526747", "8B6F4E", "E8EDE3", "F3EBDD"
WHITE, DARK, GRAY, LINE, PALE, AMBER, LIGHT = "FFFFFF", "263126", "666666", "D7DCCF", "DCE8D5", "F6E5C9", "F7F8F4"
UP_FILL, UP_FONT = "FFC7CE", "9C0006"      # weekly increase: light red fill, dark red text
DOWN_FILL, DOWN_FONT = "FFD8A8", "C65911"  # weekly decrease: light orange fill, dark orange text
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


def scope_change_text(cur, prev):
    """A total-unit change is not a sale: still report signed and pp change, flagged."""
    n = int(cur["signed"] - prev["signed"])
    pp = (cur["signed"] / cur["total"] - prev["signed"] / prev["total"]) * 100
    signed = "0 signed change" if n == 0 else f"{n:+d} signed"
    return f"⚠ Scope changed ({prev['total']:g} → {cur['total']:g}); {signed}; {pp:+.1f} pp"


def delta(cur, prev):
    """Describe the change against the previous week's latest observation.
    No date/gap suffix: the week header gives the period (user decision 2026-11-16)."""
    if not cur: return "No record"
    if not prev: return "No prior record"
    if cur["total"] != prev["total"]:
        return scope_change_text(cur, prev)
    n = int(cur["signed"] - prev["signed"])
    return "No change" if n == 0 else f"{n:+d} signed"


def daily_change_and_gap(cur, prev):
    if not cur:
        return "No observation", None
    if not prev:
        return "No prior observation", None
    gap = (cur["date"] - prev["date"]).days
    if cur["total"] != prev["total"]:
        return scope_change_text(cur, prev), gap
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
                        fill=AMBER if "scope" in str(val).lower() else PALE if str(val).startswith("No change") else CREAM
                    font=DARK
                    if kind=="change" and "scope" not in str(val).lower():
                        # User request 2026-11-16: increases red, decreases orange.
                        if str(val).startswith("+"): fill, font = UP_FILL, UP_FONT
                        elif str(val).startswith("-"): fill, font = DOWN_FILL, DOWN_FONT
                    style_cell(ws.cell(rr,j,val),fill,kind=="change",color=font)
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
        scope_since = None  # (date, old total, new total) of the latest total change in this group
        for current in observations:
            gap = (current["date"] - previous["date"]).days if previous else None
            signed_change = current["signed"] - previous["signed"] if previous else None
            current_pct = current["signed"] / current["total"] if current["total"] else None
            previous_pct = previous["signed"] / previous["total"] if previous and previous["total"] else None
            pp_change = ((current_pct - previous_pct) * 100
                         if current_pct is not None and previous_pct is not None else None)
            total_changed = bool(previous and current["total"] != previous["total"])
            under_revised_scope = False
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
                # Keep later observations under a changed total visibly flagged for review.
                under_revised_scope = bool(not total_changed and scope_since and current["total"] == scope_since[2])
                if under_revised_scope:
                    pieces.append(f"Under revised scope since {scope_since[0]:%d %b %Y} "
                                  f"({scope_since[1]:g} → {scope_since[2]:g}); review")
                if not pieces:
                    pieces.append("No scope change detected")
            if total_changed:
                scope_since = (current["date"], previous["total"], current["total"])
            trend[(current["date"], current["project"], current["group"])] = {
                "signed": current["signed"], "total": current["total"], "pct": current_pct,
                "change": change_text, "signed_change": signed_change, "pp_change": pp_change,
                "total_changed": total_changed, "notes": pieces,
                "revised_scope": under_revised_scope,
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
                                  "scope": t["total_changed"], "revised": t["revised_scope"]})
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
                                  "change_color": DARK, "scope": False, "revised": False})

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
        if row["revised"]:
            ws.cell(rr, 8).fill = PatternFill("solid", fgColor=AMBER)
            ws.cell(rr, 8).font = Font(name="Aptos", size=9, bold=True, color="B54708")
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


def dashboard_data():
    """The DATA object embedded in index.html (projects, endpoints, property details)."""
    html = (ROOT / "index.html").read_text(encoding="utf-8")
    start = html.index("const DATA=") + len("const DATA=")
    end = start + re.search(r";\r?\nconst \$", html[start:]).start()
    return json.loads(html[start:end])


def dashboard_endpoint_codes():
    """TEDUH codes the dashboard collects, read from the DATA object embedded in index.html."""
    return {e["project"]: e["id"] for e in dashboard_data().get("endpoints", [])}


# Master View: the original "Broga Competitor Sales Status" block layout (user request 2026-11-16).
MASTER_NAME = "Master View"
DETAILS_DATE = date(2026, 10, 13)       # property details in index.html come from the 13/10/2026 workbook
MASTER_TEDUH_FROM = date(2026, 10, 19)  # first Monday after the last manual snapshot
DETAIL_ROWS = [("Land Area", "land area"), ("BUA", "bua"), ("SPA Price", "spa price"),
               ("Nett Price", "nett price"), ("Rebate", "rebate")]
NOT_CHECKED = "EDEDED"


def master_columns(project, history_rows):
    """Columns for one project, newest first: manual snapshot dates plus one TEDUH reading per Monday.

    Each column is (date, [history rows]). A Monday with no reading uses the latest TEDUH
    reading in the 6 days before it; the header shows the date actually used.
    """
    rows = [r for r in history_rows if r["project"] == project]
    manual = defaultdict(list)
    teduh = defaultdict(list)
    for r in rows:
        (manual if r["source"].startswith("Embedded") else teduh)[r["date"]].append(r)
    cols = dict(manual)
    monday = MASTER_TEDUH_FROM
    while monday <= TODAY:
        used = max((d for d in teduh if monday - timedelta(days=6) <= d <= monday), default=None)
        if used and used not in cols:
            cols[used] = teduh[used]
        monday += timedelta(days=7)
    return sorted(cols.items(), key=lambda kv: kv[0], reverse=True)


def master_status_text(rows):
    """Sales Status cell: the recorded text for manual snapshots, the TEDUH count otherwise.
    Multi-layout projects list one line per layout, prefixed with the layout name."""
    rows = sorted(rows, key=lambda r: r["group"].casefold())
    lines = []
    for r in rows:
        if r["source"].startswith("Embedded") and r["raw"]:
            text = r["raw"].replace("\n", " ")
        elif r["signed"] is not None and r["total"]:
            text = status(r)
        else:
            text = r["raw"] or r["status"]
        lines.append(f"{r['group']}: {text}" if len(rows) > 1 else text.replace(" (", "\n(", 1))
    return "\n".join(lines)


def project_signed(rows):
    vals = [r for r in rows if str(r["status"]).lower() == "signed" and r["signed"] is not None]
    return sum(r["signed"] for r in vals) if vals else None


def project_total(rows):
    vals = [r["total"] for r in rows if r["total"] is not None]
    return sum(vals) if vals else None


MASTER_UP, MASTER_DOWN = "FF0000", DOWN_FONT  # red = more signed, orange = fewer (user 2026-11-16)


def master_marks(rows, older):
    """Font colour and total-change note for a Sales Status cell versus the older column."""
    now, before = project_signed(rows), project_signed(older) if older else None
    color = MASTER_UP if (now is not None and before is not None and now > before) else \
        MASTER_DOWN if (now is not None and before is not None and now < before) else None
    t_now, t_before = project_total(rows), project_total(older) if older else None
    note = (f"⚠ Total units changed ({t_before:g} → {t_now:g})"
            if t_now is not None and t_before is not None and t_now != t_before else None)
    return color, note


def pct_mismatches(rows):
    """Manual snapshots whose recorded % does not match their own units, e.g. '25% Signed (81/189 units)'."""
    out = []
    for r in rows:
        m = re.search(r"(\d+(?:\.\d+)?)%\s*\w+[^(]*\((\d+)/(\d+) units\)", str(r["raw"]).replace("\n", " "))
        if m and int(m.group(3)) > 0:
            pct, a, b = float(m.group(1)), int(m.group(2)), int(m.group(3))
            calc = a / b * 100
            if abs(round(calc) - pct) > 1:
                out.append(f"{r['group']}: recorded {pct:g}% but {a}/{b} = {calc:.0f}%")
    return out


def create_master_view(history_rows):
    """Add the Master View tab: one block per project, rows as in the original workbook."""
    data = dashboard_data()
    details = data.get("propertyDetails", {})
    projects = data.get("projects", [])
    wb = load_workbook(WEEKLY_OUT)
    if MASTER_NAME in wb.sheetnames:
        del wb[MASTER_NAME]
    ws = wb.create_sheet(MASTER_NAME, 0)
    blocks = [(p, master_columns(p, history_rows)) for p in projects]
    width = max((len(c) for _, c in blocks), default=1) + 1
    title(ws, width, "BROGA COMPETITOR SALES STATUS | MASTER VIEW")
    note = ws.cell(2, 1, f"Land Area, BUA, prices and rebate are as recorded on {DETAILS_DATE:%d/%m/%Y} and carried "
                         "forward unchanged (no updates since); grey '—' = earlier values not in saved data. "
                         "Sales Status after 13/10/2026 is the TEDUH "
                         "reading for each Monday. Red = more units signed than the column to its right; orange = fewer. "
                         "'⚠ Total units changed' = % not directly comparable. Amber cell with a comment = the "
                         "recorded % does not match its units in the source; check the original.")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
    note.font = Font(name="Aptos", size=9, italic=True, color=GRAY)
    note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30
    thin, thick = Side(style="thin", color="000000"), Side(style="medium", color="000000")
    rr = 4
    for project, cols in blocks:
        header = ws.cell(rr, 1, project)
        header.font = Font(name="Aptos", size=10, bold=True)
        for j, (d, _) in enumerate(cols, 2):
            ws.cell(rr, j, d.strftime("%d/%m/%Y")).font = Font(name="Aptos", size=10)
        for j in range(1, len(cols) + 2):
            ws.cell(rr, j).border = Border(left=thin, right=thin, top=thick, bottom=thin)
            ws.cell(rr, j).alignment = Alignment(vertical="top")
        for label, key in DETAIL_ROWS + [("Sales Status", None)]:
            rr += 1
            ws.cell(rr, 1, label).font = Font(name="Aptos", size=10)
            ws.cell(rr, 1).border = Border(left=thin, right=thin, top=thin, bottom=thin)
            ws.cell(rr, 1).alignment = Alignment(vertical="top")
            lines = 1
            for j, (d, rows) in enumerate(cols, 2):
                cell = ws.cell(rr, j)
                font = Font(name="Aptos", size=10)
                if key:
                    value = details.get(project, {}).get(key)
                    # User 2026-11-16: no price updates since 13/10/2026, so later columns copy it exactly.
                    if d >= DETAILS_DATE and value not in (None, ""):
                        value = str(value).strip()
                        # Same rule as the dashboard: rebates stored as fractions (0.1) show as 10%.
                        if key == "rebate" and re.fullmatch(r"0\.\d+", value):
                            value = f"{float(value) * 100:g}%"
                        cell.value = value
                    else:
                        cell.value = "—"
                        cell.fill = PatternFill("solid", fgColor=NOT_CHECKED)
                        font = Font(name="Aptos", size=10, color=GRAY)
                else:
                    older = cols[j - 1][1] if j - 1 < len(cols) else None
                    color, total_note = master_marks(rows, older)
                    text = master_status_text(rows)
                    cell.value = f"{text}\n{total_note}" if total_note else text
                    if color:
                        font = Font(name="Aptos", size=10, color=color)
                    mismatch = pct_mismatches(rows)
                    if mismatch:
                        cell.comment = Comment("Source % does not match units; check original workbook.\n"
                                               + "\n".join(mismatch), "Report check")
                        cell.fill = PatternFill("solid", fgColor=AMBER)
                cell.font = font
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                lines = max(lines, str(cell.value).count("\n") + 1)
            ws.row_dimensions[rr].height = max(15, 14 * lines)
        rr += 2
    ws.column_dimensions["A"].width = 30
    for j in range(2, width + 1):
        ws.column_dimensions[get_column_letter(j)].width = 30
    ws.freeze_panes = "B4"
    # Inserting at position 0 shifts the tabs; keep the newest monthly tab as the one that opens.
    months = [s for s in wb.sheetnames if re.fullmatch(r"[A-Z][a-z]{2} \d{4}", s)]
    for sheet in wb.worksheets:
        sheet.sheet_view.tabSelected = False
    if months:
        wb.active = wb.sheetnames.index(months[-1])
        wb[months[-1]].sheet_view.tabSelected = True
    wb.save(WEEKLY_OUT)
    return blocks


def validate_master_view(blocks, records):
    """Re-open Master View and reconcile it with the source data; stop on any mismatch."""
    problems = []
    data = dashboard_data()
    ws = load_workbook(WEEKLY_OUT)[MASTER_NAME]
    cells = [list(r) for r in ws.iter_rows()]
    sheet_rows = [[c.value for c in r] for r in cells]
    found = {}
    for i, values in enumerate(sheet_rows):
        if values and values[0] in data.get("projects", []):
            found[values[0]] = i
    missing = [p for p in data.get("projects", []) if p not in found]
    if missing:
        problems.append(f"Master View has no block for {missing}")
    signed_by = defaultdict(float)
    total_by = defaultdict(float)
    for r in records:
        signed_by[(r["date"], r["project"])] += r["signed"]
        total_by[(r["date"], r["project"])] += r["total"]
    checked = 0
    for project, cols in blocks:
        if project not in found:
            continue
        top = found[project]
        status_row = sheet_rows[top + len(DETAIL_ROWS) + 1]
        if status_row[0] != "Sales Status":
            problems.append(f"{project}: Sales Status row is not where expected")
            continue
        for j, (d, rows) in enumerate(cols, 1):
            header = sheet_rows[top][j]
            if header != d.strftime("%d/%m/%Y"):
                problems.append(f"{project}: column {j} header {header!r} != {d:%d/%m/%Y}")
            # Detail rows: 13/10/2026 and later columns copy the 13/10/2026 value exactly; older are "—".
            base_j = next((jj for jj, (dd, _) in enumerate(cols, 1) if dd == DETAILS_DATE), None)
            for k in range(1, len(DETAIL_ROWS) + 1):
                v = sheet_rows[top + k][j]
                if d < DETAILS_DATE and v != "—":
                    problems.append(f"{project} {d}: {sheet_rows[top + k][0]} has {v!r} before {DETAILS_DATE}")
                if d > DETAILS_DATE and base_j and v != sheet_rows[top + k][base_j]:
                    problems.append(f"{project} {d}: {sheet_rows[top + k][0]} {v!r} differs from {DETAILS_DATE} value")
            # Colour: red = more signed, orange = fewer, versus the older column.
            older = cols[j][1] if j < len(cols) else None
            want, _ = master_marks(rows, older)
            got = cells[top + len(DETAIL_ROWS) + 1][j].font.color
            got = str(got.rgb)[-6:] if got is not None and got.rgb else None
            if (want or None) != (got if got in (MASTER_UP, MASTER_DOWN) else None):
                problems.append(f"{project} {d}: font colour {got} != expected {want}")
            # Manual columns: every recorded unit count appears in the cell.
            if d <= DETAILS_DATE:
                for r in rows:
                    m = re.search(r"\((\d+)/(\d+) units\)", str(r["raw"]).replace("\n", " "))
                    if m and f"({m.group(1)}/{m.group(2)} units)" not in str(status_row[j]):
                        problems.append(f"{project} {d}: {r['group']} {m.group(0)} missing from Master View")
                checked_manual = True
            if d > DETAILS_DATE and (d, project) not in signed_by:
                problems.append(f"{project} {d}: TEDUH column has no signed TEDUH record to reconcile")
            elif d > DETAILS_DATE:
                nums = [tuple(map(int, m)) for m in re.findall(r"\((\d+)/(\d+) units\)", str(status_row[j]))]
                if (sum(a for a, _ in nums), sum(b for _, b in nums)) != (signed_by[(d, project)], total_by[(d, project)]):
                    problems.append(f"{project} {d}: Master View {nums} != TEDUH {signed_by[(d, project)]:g}/{total_by[(d, project)]:g}")
                checked += 1
    if problems:
        for p in problems[:30]:
            print(f"VALIDATION ERROR: {p}")
        raise ValueError(f"{len(problems)} Master View validation problem(s).")
    print(f"Validation: Master View has all {len(found)} projects; {checked} TEDUH Monday cells reconcile.")


def create_reference_sheet():
    """Add the TEDUH search reference (keyword, project code, data link) for checking projects in a browser."""
    reference = json.loads((ROOT / "data" / "teduh-reference.json").read_text(encoding="utf-8"))["projects"]
    collected = dashboard_endpoint_codes()
    for item in reference:
        code = collected.get(item["project"])
        if code and item["code"] and code != item["code"]:
            raise ValueError(f"TEDUH reference code for {item['project']} is {item['code']} but the dashboard collects {code}; "
                             "update data/teduh-reference.json.")
    wb = load_workbook(WEEKLY_OUT)
    name = "TEDUH Search Reference"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    headers = ["BROGA Developer", "Search Keyword", "Project Code and Name", "TEDUH Data Link"]
    title(ws, len(headers), "BROGA COMPETITOR SALES ANALYSIS | TEDUH SEARCH REFERENCE")
    for col, header in enumerate(headers, 1):
        style_cell(ws.cell(3, col, header), BROWN, True, WHITE, size=9)
    ws.row_dimensions[3].height = 28
    for rr, item in enumerate(reference, start=4):
        code_name = f"{item['code']}\n{item['teduhName']}" if item["code"] else "No TEDUH code recorded"
        values = [item["project"], item["keyword"] or "—", code_name, ""]
        fill = WHITE if rr % 2 == 0 else LIGHT
        for col, value in enumerate(values, 1):
            style_cell(ws.cell(rr, col, value), fill, col == 1, horizontal="left", size=9)
        if item["code"]:
            link = ws.cell(rr, 4, f"Open {item['code']}")
            link.hyperlink = f"https://teduh.kpkt.gov.my/api/unit-projek-swasta/{item['code']}"
            link.font = Font(name="Aptos", size=9, color="1769AA", underline="single")
        ws.row_dimensions[rr].height = 30
    for col, width in enumerate([32, 22, 30, 18], 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A4"
    # Keep the newest monthly tab as the one that opens.
    months = [s for s in wb.sheetnames if re.fullmatch(r"[A-Z][a-z]{2} \d{4}", s)]
    if months:
        wb.active = wb.sheetnames.index(months[-1])
    wb.save(WEEKLY_OUT)


def check_source_is_current(history_rows):
    """Refuse to build reports from an analysis workbook older than the saved TEDUH history."""
    teduh = ROOT / "data" / "teduh-history.json"
    if not teduh.exists():
        print("WARNING: data/teduh-history.json not found; freshness check skipped.")
        return
    saved = [str(r.get("date", ""))[:10] for r in json.loads(teduh.read_text(encoding="utf-8")).get("records", [])]
    saved = [d for d in saved if len(d) == 10]
    if not saved:
        return
    latest_saved = date.fromisoformat(max(saved))
    latest_workbook = max((r["date"] for r in history_rows), default=None)
    if latest_workbook is None or latest_workbook < latest_saved:
        raise ValueError(f"{SOURCE.name} is stale: latest record {latest_workbook} but data/teduh-history.json "
                         f"has {latest_saved}. Run scripts/export_competitor_workbook.mjs first.")


def validate_outputs(records, groups, lookup):
    """Re-open the written files and check daily, weekly and Snapshot History agree."""
    problems = []
    daily = load_workbook(DAILY_OUT, read_only=True)
    expected_tab = month_name(TODAY.year, TODAY.month)
    if daily.sheetnames != [expected_tab]:
        problems.append(f"Daily workbook tabs are {daily.sheetnames}, expected only [{expected_tab!r}]")
    ws = daily[daily.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    for i, values in enumerate(rows):
        label = str(values[0] or "")
        if not label.endswith("· Sales Status"):
            continue
        dt = datetime.strptime(label.split(" · ")[0], "%d %b %Y").date()
        for j, g in enumerate(groups, 1):
            cur = lookup.get((dt, g))
            want = status(cur) if cur else "No signed observation"
            if values[j] != want:
                problems.append(f"Daily {dt} {g}: {values[j]!r} != {want!r}")
    daily.close()

    weekly = load_workbook(WEEKLY_OUT, read_only=True)
    if "Snapshot History" not in weekly.sheetnames:
        problems.append("Weekly workbook has no Snapshot History tab")
    if "Latest Snapshot Comparison" in weekly.sheetnames:
        problems.append("Weekly workbook still has Latest Snapshot Comparison")
    snap = set()
    for values in weekly["Snapshot History"].iter_rows(min_row=4, values_only=True):
        if values[0] is None:
            continue
        d = values[0].date() if isinstance(values[0], datetime) else values[0]
        if str(values[3]).lower() == "signed":
            snap.add((d, values[1], values[2], values[4]))
    for r in records:
        if (r["date"], r["project"], r["group"], f"{r['signed']:g} / {r['total']:g}") not in snap:
            problems.append(f"Snapshot History is missing {r['date']} {r['project']} / {r['group']}")
    weekly.close()

    if problems:
        for p in problems[:30]:
            print(f"VALIDATION ERROR: {p}")
        raise ValueError(f"{len(problems)} output validation problem(s); reports are not consistent.")
    print(f"Validation: daily tab, Snapshot History and {len(records)} signed observations reconcile.")


def main():
    records=read_records()
    if not records: raise ValueError("No valid signed records found in Historical Records.")
    history_rows=read_history_rows()
    check_source_is_current(history_rows)
    groups, lookup, by_group = group_records(records)
    create_daily(records, groups, lookup, by_group)
    create_weekly(records, groups, lookup, by_group)
    create_snapshot_history(records, by_group, history_rows)
    create_reference_sheet()
    blocks = create_master_view(history_rows)
    validate_outputs(records, groups, lookup)
    validate_master_view(blocks, records)
    daily_tabs = load_workbook(DAILY_OUT, read_only=True).sheetnames
    weekly_tabs = load_workbook(WEEKLY_OUT, read_only=True).sheetnames
    print(f"Created {DAILY_OUT.name} with {len(daily_tabs)} tabs ({', '.join(daily_tabs)})")
    print(f"Created {WEEKLY_OUT.name} with {len(weekly_tabs)} tabs ({', '.join(weekly_tabs)}), including date-filtered snapshot history and Monday–Sunday monthly tabs")
    print(f"Source: {len(records)} valid signed observations across {len(groups)} project/layout groups.")

if __name__ == "__main__":
    main()
