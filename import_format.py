"""Copy the look of an existing department schedule into a template.

Reads one sheet laid out like the department's live schedule (a "Date" cell with dates to its
right, a "Day" row under it, pharmacists below, and an optional legend further down), then
rebuilds the template with:
  - each code's fill, text color, font and size (colors converted to exact RGB),
  - the schedule's fonts, title and PRN STAFF row style,
  - every non-role code found, added to the Other Codes table (meanings marked "assumed" where
    guessed - please review them),
  - role descriptions (shift times, units) taken from the legend,
  - optionally, PRN pharmacists.

Only the chosen sheet is read; other sheets in that workbook are ignored and nothing from them
is copied.

Usage:
    python import_format.py live.xlsx --sheet "Nov 2-Dec 13 FINAL LIVE" \
        --template "Pharmacist Schedule Template.xlsx" --out "New Template.xlsx" \
        --project-code PROJ --add-prn 9 --prn-roles CH,CP
"""
import argparse
import sys
from collections import Counter, defaultdict
from copy import copy
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

from scheduler import LoaderError, load_template
from scheduler.colors import portable_style, resolve, theme_colors
from scheduler.loader import Staff, style_signature
from scheduler.template_builder import build_template

# Best guesses for this department's non-role codes: (counts toward days, workday, covers, meaning).
# Anything marked "assumed" should be checked on the Requirements sheet's Other Codes table.
GUESSES = {
    "PTO": (True, False, None, "Paid time off"),
    "HPTO": (True, False, None, "Holiday PTO (assumed)"),
    "X": (False, False, None, "Off / unavailable (assumed)"),
    "OFF": (False, False, None, "Day off"),
    "A": (False, False, None, "Available - used in PRN rows (assumed)"),
    "?": (False, False, None, "Unconfirmed - treated as not working (assumed)"),
    "PROJ": (True, True, None, "Project day"),
    "MIDY": (True, False, None, "Midyear meeting (assumed)"),
    "MY": (True, False, None, "Midyear meeting (assumed)"),
    "JRY": (True, False, None, "Jury duty (assumed)"),
    "MGT": (True, True, None, "Management (assumed)"),
    "RECR": (True, True, None, "Recruitment (assumed)"),
    "ON8": (True, True, "ON", "8-hour ON shift (assumed; covers ON)"),
    "NW8": (True, True, "NW", "8-hour NW shift (assumed; covers NW)"),
    "ER8": (True, True, "ER", "8-hour ER shift (assumed; covers ER)"),
    "CH/ECH": (True, True, "CH", "Split CH/ECH shift (assumed; counted as covering CH)"),
    "WPC": (True, True, None, "12-hour weekend shift 0800-2030"),
}
TRAINING_SUFFIX = "t"   # e.g. TXt, SPdt: training in that role (assumed; doesn't cover it)


def find_cell(ws, text):
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.strip().lower() == text.lower():
                return c
    return None


def main() -> int:
    ap = argparse.ArgumentParser(description="Copy a live schedule's formatting into a template.")
    ap.add_argument("live")
    ap.add_argument("--sheet", help="sheet to read (default: the active sheet)")
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--project-code", help="code the program writes for project days (e.g. PROJ)")
    ap.add_argument("--add-prn", type=int, default=0, help="add this many PRN pharmacists (PRN1, PRN2, ...)")
    ap.add_argument("--prn-roles", default="", help="roles the added PRNs can work, e.g. CH,CP")
    ap.add_argument("--prn-score", type=int, default=2)
    args = ap.parse_args()

    try:
        t = load_template(args.template)
    except (LoaderError, FileNotFoundError) as e:
        print(f"Couldn't load the template: {e}")
        return 2

    wb = openpyxl.load_workbook(args.live)
    ws = wb[args.sheet] if args.sheet else wb.active
    theme = theme_colors(wb)
    date_cell = find_cell(ws, "Date")
    if date_cell is None:
        print(f"No 'Date' cell on sheet '{ws.title}'.")
        return 2
    r0, lc = date_cell.row, date_cell.column
    date_cols = []
    c = lc + 1
    while isinstance(ws.cell(r0, c).value, (datetime, date)):
        date_cols.append(c)
        c += 1
    day_row = r0 + 1
    roles = {r.upper(): r for r in t.roles}

    # ---- schedule rows: until the notes/legend start
    data_rows = []
    r = day_row + 1
    while r <= ws.max_row:
        lab = ws.cell(r, lc).value
        lab = lab.strip() if isinstance(lab, str) else lab
        if isinstance(lab, str) and (lab.startswith("**") or lab.upper() in roles):
            break
        data_rows.append(r)
        r += 1
    legend_start = r

    # ---- per-code formatting (most common), grouping case variants (x = X)
    seen = defaultdict(Counter)
    example = {}
    spelling = defaultdict(Counter)
    base_font = Counter()
    for r in data_rows:
        for c in date_cols:
            cell = ws.cell(r, c)
            v = cell.value
            if not isinstance(v, str) or not v.strip() or v.strip().startswith("="):
                continue
            code = v.strip()
            key = code.upper()
            spelling[key][code] += 1
            st = portable_style(cell, wb, theme)
            sig = style_signature(st)
            seen[key][sig] += 1
            example[sig] = st
            base_font[(cell.font.name, cell.font.sz)] += 1
    canonical = {k: (roles.get(k) or spelling[k].most_common(1)[0][0]) for k in seen}
    styles = {canonical[k]: example[c.most_common(1)[0][0]] for k, c in seen.items()}

    # ---- the sheet's look
    (bf_name, bf_size), _ = base_font.most_common(1)[0]
    def pfont(cell, **over):
        f = cell.font
        rgb = resolve(f.color, wb, theme)
        return Font(name=f.name, sz=f.sz, b=over.get("b", f.b), i=f.i,
                    color=("FF" + rgb) if rgb and rgb != "000000" else None)
    title_cell = next((ws.cell(rr, lc) for rr in range(r0 - 1, 0, -1)
                       if isinstance(ws.cell(rr, lc).value, str)), None)
    prn_cell = find_cell(ws, "PRN STAFF")
    first_name = ws.cell(data_rows[0], lc) if data_rows else date_cell
    look = {
        "title_font": pfont(title_cell) if title_cell else Font(name=bf_name, sz=20, b=True),
        "label_font": pfont(date_cell),
        "date_font": pfont(ws.cell(r0, date_cols[0])),
        "day_font": pfont(ws.cell(day_row, date_cols[0])),
        "date_format": ws.cell(r0, date_cols[0]).number_format or "m/d",
        "name_font": pfont(first_name, b=True),
        "pid_font": Font(name=bf_name, sz=bf_size, color="FF808080"),
        "cell_font": Font(name=bf_name, sz=bf_size),
        "section_font": pfont(prn_cell) if prn_cell else Font(name=bf_name, sz=bf_size, b=True),
        "section_fill": (PatternFill("solid", fgColor="FF" + (resolve(prn_cell.fill.fgColor, wb, theme) or "BFBFBF"))
                         if prn_cell and prn_cell.fill.fill_type == "solid" else PatternFill("solid", fgColor="BFBFBF")),
        "shade_weekend_headers": False,
        "date_col_width": 6.5,
    }

    # ---- legend: code cell followed by its description cells
    # legend may name a role without its trailing number (IM for IM1)
    short_names = {}
    for r_ in sorted(t.roles, key=lambda x: not x.endswith("1")):   # IM -> IM1, not IM2
        base = r_.rstrip("0123456789")
        if r_[-1].isdigit() and base.upper() not in roles and base not in short_names:
            short_names[base] = r_
    known = set(canonical.values()) | set(t.roles) | set(GUESSES) | set(short_names)
    legend = {}
    legend_rows = []
    for r in range(legend_start, min(ws.max_row, legend_start + 15) + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(isinstance(v, str) and v.startswith("=") for v in vals):
            break
        legend_rows.append(r)
    # columns that hold codes in the legend: a description stops at the next such column
    code_cols = {c for r in legend_rows for c in range(1, ws.max_column + 1)
                 if isinstance(ws.cell(r, c).value, str) and ws.cell(r, c).value.strip() in known}
    for r in legend_rows:
        cells = [(c, ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        cells = [(c, v.strip() if isinstance(v, str) else v) for c, v in cells if v not in (None, "")]
        if any(isinstance(v, str) and v.startswith("=") for _, v in cells):
            break  # formula area: stop
        i = 0
        while i < len(cells):
            c, v = cells[i]
            code = v if isinstance(v, str) and v in known else None
            if code:
                desc, j = [], i + 1
                while j < len(cells) and cells[j][0] - c <= 12 and cells[j][0] not in code_cols and not (
                        isinstance(cells[j][1], str) and cells[j][1] in known):
                    desc.append(str(cells[j][1]))
                    j += 1
                if desc and code not in legend:
                    legend[code] = " - ".join(desc)
                i = j
            else:
                i += 1
    for short, role in short_names.items():
        if short in legend and role not in legend:
            legend[role] = legend[short]
    for role in t.roles:
        if role in legend:
            have = t.role_descriptions.get(role, "")
            if legend[role] not in have:
                t.role_descriptions[role] = f"{have} - {legend[role]}" if have else legend[role]

    # ---- shift times from the legend ("0630 to 1500", "1130-2200")
    import re
    def legend_times(text):
        m = re.search(r"\b(\d{3,4})\s*(?:to|-)\s*(\d{3,4})\b", text or "")
        if not m:
            return None
        a, b = (int(x[:-2]) * 60 + int(x[-2:]) for x in m.groups())
        return (a, b if b > a else b + 1440)
    for code, text in legend.items():
        tm = legend_times(text)
        if tm and code not in t.shift_times:
            t.shift_times[code] = tm

    # ---- codes: roles keep their place; everything else goes to Other Codes
    if args.project_code:
        old = t.settings.project_code
        t.settings.project_code = args.project_code
        if old != args.project_code and old in t.codes and old.upper() not in seen:
            t.codes.pop(old)
    new_codes = []
    for key, code in sorted(canonical.items(), key=lambda kv: -sum(seen[kv[0]].values())):
        if key in roles:
            continue
        if code in t.codes:
            continue
        if code in GUESSES:
            pay, work, cov, meaning = GUESSES[code]
        elif code.endswith(TRAINING_SUFFIX) and code[:-1].upper() in roles:
            role = roles[code[:-1].upper()]
            pay, work, cov, meaning = True, True, None, f"Training in {role} (assumed; doesn't cover {role})"
        else:
            pay, work, cov, meaning = True, True, None, "Unknown - please describe and check the Yes/No columns"
        if code in legend and "assumed" not in meaning:
            meaning = f"{meaning} - {legend[code]}" if meaning else legend[code]
        t.codes[code] = {"counts_for_pay": pay, "on_duty": work, "covers": cov, "meaning": meaning}
        new_codes.append(code)
    if t.settings.project_code not in t.codes:
        t.codes[t.settings.project_code] = {"counts_for_pay": True, "on_duty": True, "covers": None,
                                            "meaning": "Project day"}
    t.code_styles = {**t.code_styles, **styles}
    t.style_book = None   # don't relearn from the old template's cells

    # ---- PRNs
    prn_roles = [roles[r.strip().upper()] for r in args.prn_roles.split(",") if r.strip()]
    for k in range(1, args.add_prn + 1):
        pid = f"PRN{k}"
        if pid in t.staff:
            continue
        t.staff[pid] = Staff(pid, {r: (args.prn_score if r in prn_roles else 0) for r in t.roles},
                             None, "", prn=True)
        t.schedule_order.append(pid)

    build_template(t, args.out, look=look)
    print(f"Wrote {args.out}")
    print(f"  Read sheet '{ws.title}': {len(data_rows)} rows x {len(date_cols)} dates; "
          f"formatting captured for {len(styles)} codes; base font {bf_name} {bf_size:g}")
    print(f"  Added to Other Codes: {', '.join(new_codes) or 'none'}")
    print(f"  Role descriptions from legend: {', '.join(r for r in t.roles if r in legend) or 'none'}")
    print(f"  Shift times: {', '.join(f'{c} {a // 60:02d}{a % 60:02d}-{(b % 1440) // 60:02d}{b % 60:02d}' for c, (a, b) in t.shift_times.items())}")
    missing = [r for r in t.requirements if r not in t.shift_times]
    if missing:
        print(f"  No times found for: {', '.join(missing)}")
    if args.add_prn:
        print(f"  Added {args.add_prn} PRN pharmacists able to work: {', '.join(prn_roles) or 'nothing'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
