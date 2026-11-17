"""Write a solved schedule to a new .xlsx.

If the template was an .xlsx, the output is a copy of it (formatting and other sheets kept)
with the open cells filled in. If it was a .csv, the grid is recreated in a new workbook.
Report sheets are added after the schedule.
"""
from __future__ import annotations

from copy import copy
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import config
from .loader import (Issue, Template, cell_str, cell_style, read_applied_styles,
                     legacy_style_signature, style_signature, write_applied_styles)
from .model import Solution
from .settings import SPEC as SETTINGS_SPEC, display_value

FILLS = {
    "apt4": PatternFill("solid", fgColor="A9D08E"),
    "apt3": PatternFill("solid", fgColor="C6EFCE"),
    "apt2": PatternFill("solid", fgColor="FFEB9C"),
    "apt1": PatternFill("solid", fgColor="F8CBAD"),
    "prj": PatternFill("solid", fgColor="DDEBF7"),
    "unfilled": PatternFill("solid", fgColor="FF7C80"),
    "header": PatternFill("solid", fgColor="D9D9D9"),
    # Shifts Fulfilled matrix
    "met": PatternFill("solid", fgColor="C6EFCE"),
    "short": PatternFill("solid", fgColor="FF0000"),
    "over": PatternFill("solid", fgColor="FFEB9C"),
    "optional_open": PatternFill("solid", fgColor="E7E6E6"),
    "prn": PatternFill("solid", fgColor="7030A0"),
}
PRN_FONT = Font(bold=True, color="FFFFFF")
PRN_ROW_LABEL = "PRN NEEDED"


def _prn_text(reqs: list[tuple[str, int]]) -> str:
    """'CH', 'CH, CP', '2 CH' for one day's PRN requests."""
    return ", ".join(f"{n} {r}" if n > 1 else r for r, n in reqs)


def _remove_old_fulfilled_block(ws) -> None:
    """If the template is a previous output, delete its Shifts Fulfilled block first."""
    target = config.LABEL_SHIFTS_FULFILLED.lower()
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip().lower() == target:
                r, c = cell.row, cell.column
                n = 0
                while ws.cell(r + n, c).value not in (None, ""):
                    n += 1
                if ws.cell(r + n, c).value in (None, ""):
                    n += 1  # trailing blank row we added
                ws.delete_rows(r, n)
                return


def _write_fulfilled_block(t: Template, sol: Solution, ws) -> tuple[int, int]:
    """Insert the Shifts Fulfilled matrix under the schedule. Returns (met, short) counts
    for required role-days."""
    label_col = min(t.col_of.values())  # dates start one column right of the label column
    label_col = label_col - 1
    last_sched = max(t.row_of.values()) + 1           # 1-based Excel row of last pharmacist
    roles = list(t.requirements)
    show_prn_row = bool(t.prn_pids) or bool(sol.prn_requests)
    # Layout: schedule / blank / title / [PRN NEEDED] / role rows / legend / blank / (input matrices)
    block = 1 + int(show_prn_row) + len(roles) + 1 + 1
    start = last_sched + 2
    if all(ws.cell(last_sched + 1, c + 1).value in (None, "") for c in [label_col, *t.col_of.values()]):
        ws.insert_rows(start, block)              # reuse the existing blank row above
    else:
        ws.insert_rows(last_sched + 1, block + 1)  # add the blank row above too

    counts: dict[tuple[str, date], int] = {}
    for (p, d), code in sol.assignment.items():
        role = t.covered_role(code)
        if role:
            counts[(role, d)] = counts.get((role, d), 0) + 1
    prn = {(r, d): n for r, d, n in sol.prn_requests}

    center = Alignment(horizontal="center")
    title = ws.cell(start, label_col + 1, config.LABEL_SHIFTS_FULFILLED)
    title.font = BOLD
    letter = get_column_letter(label_col + 1)
    ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0, 16)
    for d, c in t.col_of.items():
        h = ws.cell(start, c + 1, f"{d.month}/{d.day}")
        h.font = BOLD
        h.alignment = center
        h.fill = FILLS["header"]

    if show_prn_row:
        pr = start + 1
        lab = ws.cell(pr, label_col + 1, PRN_ROW_LABEL)
        lab.font, lab.fill = PRN_FONT, FILLS["prn"]
        for d, c in t.col_of.items():
            reqs = [(r_, n) for (r_, dd), n in sorted(prn.items()) if dd == d]
            cell = ws.cell(pr, c + 1)
            cell.alignment = center
            if reqs:
                cell.value = _prn_text(reqs)
                cell.fill, cell.font = FILLS["prn"], PRN_FONT

    met = short = prn_slots = 0
    first_role_row = start + 1 + int(show_prn_row)
    for i, role in enumerate(roles):
        r = first_role_row + i
        ws.cell(r, label_col + 1, role).font = BOLD
        optional = role in t.optional_roles
        for d, c in t.col_of.items():
            need, have = t.required(role, d), counts.get((role, d), 0)
            n_prn = prn.get((role, d), 0)
            cell = ws.cell(r, c + 1)
            cell.alignment = center
            if need == 0 and have == 0:
                continue
            cell.value = f"{have}/{need}"
            if n_prn:
                cell.value = f"{have}+{n_prn} PRN/{need}" if have else f"PRN {n_prn}/{need}"
                cell.fill, cell.font = FILLS["prn"], PRN_FONT
                prn_slots += n_prn
                if have + n_prn < need and role not in t.optional_roles:
                    cell.fill = FILLS["short"]
                    short += 1
                continue
            if have > need:
                cell.fill = FILLS["over"]
            elif have == need:
                cell.fill = FILLS["met"]
                met += 0 if optional else 1
            elif optional:
                cell.fill = FILLS["optional_open"]
            else:
                cell.fill = FILLS["short"]
                cell.font = Font(bold=True, color="FFFFFF")
                short += 1
    # match the schedule's font (name and size) in the whole block
    first = ws.cell(min(t.row_of.values()) + 1, min(t.col_of.values()) + 1).font
    for rr in range(start, first_role_row + len(roles)):
        for cc in [label_col, *t.col_of.values()]:
            cell = ws.cell(rr, cc + 1)
            f = copy(cell.font)
            f.name, f.sz = first.name or f.name, first.sz or f.sz
            cell.font = f
    ws.cell(first_role_row + len(roles), label_col + 1,
            "filled/required. Red = requirement not met; purple = a PRN is needed (see PRN NEEDED "
            "and the PRN Requests sheet); gray = optional role left open; yellow = more than "
            "required (filled in).").font = Font(italic=True, size=9)
    return met, short, prn_slots
BOLD = Font(bold=True)


def _fill_for(t: Template, pid: str, code: str) -> PatternFill:
    if code == t.settings.project_code:
        return FILLS["prj"]
    a = t.staff[pid].aptitude.get(code, 0)
    return FILLS.get(f"apt{a}", FILLS["apt1"])


def _apply_code_style(t: Template, cell, code: str) -> None:
    st = t.code_styles.get(code)
    if not st:
        return
    cell.fill = copy(st["fill"])
    f = copy(st["font"])
    cell.font = f


def _clear_right_block(t: Template, ws) -> None:
    """Remove a right-hand name / days block left by an earlier run (output fed back in)."""
    date_row = min(t.row_of.values()) - 1          # Excel row of the Date header
    first = max(t.col_of.values()) + 2            # Excel column right after the last date
    if str(ws.cell(date_row, first).value or "").strip().lower() != config.LABEL_NAME.lower():
        return
    last = first
    while ws.cell(date_row, last + 1).value not in (None, ""):
        last += 1
    for r in range(max(1, date_row - 1), max(t.row_of.values()) + 2):
        for c in range(first, last + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.font = Font()
            cell.border = Border()


def _write_right_block(t: Template, sol: Solution, ws) -> int:
    """Name again, then days given / target per pay period and in total, right of the dates.
    Returns how many pharmacist-periods are short of target."""
    first_staff = min(t.row_of.values()) + 1
    date_row, day_row = first_staff - 2, first_staff - 1
    name_col = max(t.col_of.values()) + 2
    base = ws.cell(first_staff, min(t.col_of.values()) + 1).font
    fname, fsize = base.name or "Calibri", base.sz or 11
    thin = Side(style="thin", color="BFBFBF")
    grid = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    periods = t.periods()

    head = ws.cell(date_row, name_col, config.LABEL_NAME)
    head.font = Font(name=fname, sz=fsize, b=True)
    sub = ws.cell(day_row, name_col, "")
    for k, per in enumerate(periods, 1):
        c = name_col + k
        lbl = f"{per['start'].month}/{per['start'].day}-{per['end'].month}/{per['end'].day}"
        h = ws.cell(date_row, c, lbl + ("" if per["complete"] else " (partial)"))
        h.font, h.alignment, h.fill = Font(name=fname, sz=fsize, b=True), center, FILLS["header"]
        d = ws.cell(day_row, c, "days/target")
        d.font, d.alignment = Font(name=fname, sz=max(7, fsize - 2), i=True), center
        ws.column_dimensions[get_column_letter(c)].width = 11
    tot_c = name_col + len(periods) + 1
    h = ws.cell(date_row, tot_c, "Total")
    h.font, h.alignment, h.fill = Font(name=fname, sz=fsize, b=True), center, FILLS["header"]
    d = ws.cell(day_row, tot_c, "days/target")
    d.font, d.alignment = Font(name=fname, sz=max(7, fsize - 2), i=True), center
    ws.column_dimensions[get_column_letter(tot_c)].width = 11
    ws.column_dimensions[get_column_letter(name_col)].width = 22

    # PRN STAFF separator rows: repeat the label and band on the right
    left_name_col = min(t.col_of.values()) - 1        # Excel column of the left Name column
    for r in range(first_staff, max(t.row_of.values()) + 2):
        if left_name_col < 1:
            break
        left = ws.cell(r, left_name_col)
        if str(left.value or "").strip().upper() == config.LABEL_PRN_SECTION:
            for c in range(name_col, tot_c + 1):
                ws.cell(r, c).fill = copy(left.fill)
            lab = ws.cell(r, name_col, config.LABEL_PRN_SECTION)
            lab.font = copy(left.font)

    short = 0
    for pid, r0 in t.row_of.items():
        r = r0 + 1
        s = t.staff.get(pid)
        nm = ws.cell(r, name_col, (s.name if s and s.name else pid))
        nm.font = Font(name=fname, sz=fsize, b=True)
        tot_n = tot_t = 0
        has_target = bool(s and s.target_days and not s.prn)
        for k, per in enumerate(periods, 1):
            n = sum(1 for d in per["dates"] if (pid, d) in sol.assignment
                    and t.code_counts_for_pay(sol.assignment[pid, d]))
            cell = ws.cell(r, name_col + k)
            cell.alignment, cell.border = center, grid
            if not has_target:
                cell.value = n if n else ""
                cell.font = Font(name=fname, sz=fsize, color="808080")
                continue
            target = s.target_days if per["complete"] else round(
                s.target_days * len(per["dates"]) / t.settings.pay_period_days)
            cell.value = f"{n}/{target}"
            cell.font = Font(name=fname, sz=fsize, b=True)
            if n == target:
                cell.fill = FILLS["met"]
            elif n < target:
                cell.fill, cell.font = FILLS["short"], Font(name=fname, sz=fsize, b=True, color="FFFFFF")
                short += 1
            else:
                cell.fill = FILLS["over"]
            tot_n += n
            tot_t += target
        cell = ws.cell(r, tot_c)
        cell.alignment, cell.border = center, grid
        if has_target:
            cell.value = f"{tot_n}/{tot_t}"
            cell.font = Font(name=fname, sz=fsize, b=True)
            cell.fill = FILLS["met"] if tot_n == tot_t else FILLS["short"] if tot_n < tot_t else FILLS["over"]
            if tot_n < tot_t:
                cell.font = Font(name=fname, sz=fsize, b=True, color="FFFFFF")
        else:
            tot_all = sum(1 for d in t.dates if (pid, d) in sol.assignment
                          and t.code_counts_for_pay(sol.assignment[pid, d]))
            cell.value = tot_all if tot_all else ""
            cell.font = Font(name=fname, sz=fsize, color="808080")
    if date_row > 1 and ws.cell(date_row - 1, name_col + 1).value in (None, ""):
        note = ws.cell(date_row - 1, name_col + 1,
                       "Days given / target (PTO counts). Green = target met, red = short, "
                       "yellow = over (typed-in days). Gray = PRN or no target.")
        note.font = Font(name=fname, sz=max(7, fsize - 2), i=True, color="595959")
    return short


def _write_aptitude_view(t: Template, sol: Solution, wb) -> None:
    """Review copy of the schedule: program-filled cells colored by aptitude, filled-in cells grey."""
    av = wb.create_sheet("Aptitude View")
    av.cell(1, 1, "Aptitude View").font = Font(bold=True, size=14)
    av.cell(1, 4, "Review copy for the scheduler. Program-filled cells are colored by aptitude: "
                  "dark green 4, green 3, yellow 2, orange 1, blue project day. Grey text = filled in "
                  "on the template. The Schedule sheet is the one to publish.").font = Font(italic=True, color="595959")
    av.cell(2, 1, "Name").font = BOLD
    av.cell(2, 2, "PID").font = BOLD
    center = Alignment(horizontal="center")
    for j, d in enumerate(t.dates):
        h = av.cell(2, 3 + j, f"{d.month}/{d.day}")
        h.font, h.alignment, h.fill = BOLD, center, FILLS["header"]
        w = av.cell(3, 3 + j, d.strftime("%a").upper())
        w.font, w.alignment = BOLD, center
    grey = Font(color="808080")
    for i, pid in enumerate(t.schedule_order):
        r = 4 + i
        av.cell(r, 1, t.staff[pid].name if pid in t.staff else "")
        av.cell(r, 2, pid).alignment = center
        for j, d in enumerate(t.dates):
            code = sol.assignment.get((pid, d))
            if not code:
                continue
            cell = av.cell(r, 3 + j, code)
            cell.alignment = center
            if (pid, d) in sol.solver_filled:
                cell.fill = _fill_for(t, pid, code)
            else:
                cell.font = grey
    av.column_dimensions["A"].width = 20
    av.column_dimensions["B"].width = 6
    for j in range(len(t.dates)):
        av.column_dimensions[get_column_letter(3 + j)].width = 6.5
    av.freeze_panes = "C4"


def _md(d: date) -> str:
    return f"{d.month}/{d.day}"


def _table(ws, headers: list[str], rows: list[list], start_row: int = 1) -> int:
    for c, h in enumerate(headers, 1):
        cell = ws.cell(start_row, c, h)
        cell.font = BOLD
        cell.fill = FILLS["header"]
    for r, row in enumerate(rows, start_row + 1):
        for c, v in enumerate(row, 1):
            ws.cell(r, c, v)
    for c, h in enumerate(headers, 1):
        width = max([len(str(h))] + [len(str(row[c - 1])) for row in rows if c - 1 < len(row)])
        ws.column_dimensions[get_column_letter(c)].width = min(max(width + 2, 8), 90)
    ws.freeze_panes = ws.cell(start_row + 1, 1)
    return start_row + len(rows) + 1


def write_output(t: Template, sol: Solution, issues: list[Issue], out_path: str | Path) -> Path:
    out_path = Path(out_path)

    # ---- schedule sheet
    if t.source.suffix.lower() in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(t.source)   # keeps formulas and formatting
        ws = wb[t.sheet_name] if t.sheet_name else wb.worksheets[0]
        _remove_old_fulfilled_block(ws)
        _clear_right_block(t, ws)
        for name in config.REPORT_SHEETS:          # output fed back in as a template
            if name in wb.sheetnames:
                del wb[name]
        # Header dates/days as plain values, so the output reloads even without Excel recalculating.
        date_row = min(t.row_of.values()) - 1      # Day row; Date row is one above
        for d, c in t.col_of.items():
            dc, wc = ws.cell(date_row - 1 + 1, c + 1), ws.cell(date_row + 1, c + 1)
            if isinstance(dc.value, str) and dc.value.startswith("="):
                dc.value = d
                dc.number_format = "d-mmm"
            if isinstance(wc.value, str) and wc.value.startswith("="):
                wc.value = d.strftime("%a").upper()
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Schedule"
        for r, row in enumerate(t.grid, 1):
            for c, v in enumerate(row, 1):
                s = cell_str(v)
                if s == "":
                    continue
                ws.cell(r, c, int(s) if s.lstrip("-").isdigit() else s)

    center = Alignment(horizontal="center")
    if t.source.suffix.lower() == ".csv":
        for pid, r in t.row_of.items():
            for c in t.col_of.values():
                ws.cell(r + 1, c + 1).alignment = center
    role_colors = bool(t.code_styles)
    previously_applied = read_applied_styles(wb)
    for (pid, d) in sorted(sol.solver_filled, key=lambda k: (t.row_of[k[0]], k[1])):
        code = sol.assignment[pid, d]
        cell = ws.cell(t.row_of[pid] + 1, t.col_of[d] + 1, code)
        if role_colors:
            _apply_code_style(t, cell, code)
        else:
            cell.fill = _fill_for(t, pid, code)
        cell.alignment = center
    if role_colors:
        # Filled-in cells get their code's colors too, so the schedule is consistent, unless
        # someone formatted the cell by hand. A cell counts as hand-formatted when its formatting
        # is neither blank nor what the program applied last time (recorded in a hidden sheet).
        for (pid, d), code in t.locked.items():
            cell = ws.cell(t.row_of[pid] + 1, t.col_of[d] + 1)
            own_st = cell_style(cell)
            own = style_signature(own_st)
            prev = previously_applied.get(code)
            if own == "" or prev in (own, legacy_style_signature(own_st)):
                if code in t.code_styles:
                    _apply_code_style(t, cell, code)
                elif own:                       # color removed from the Color cell: clear it
                    cell.fill = PatternFill(fill_type=None)
                    cell.font = Font()
        write_applied_styles(wb, t.code_styles)

    short_periods = _write_right_block(t, sol, ws)
    met, short, prn_slots = _write_fulfilled_block(t, sol, ws)
    _write_aptitude_view(t, sol, wb)
    if t.source.suffix.lower() == ".csv":
        first_staff_row = min(t.row_of.values()) + 1
        ws.freeze_panes = ws.cell(first_staff_row, min(t.col_of.values()) + 1)

    # ---- Summary
    sm = wb.create_sheet("Summary", 0 if False else None)
    gap = (sol.bound - sol.objective) if sol.objective is not None else None
    required_unfilled = sum(n for _, _, n in sol.unfilled)
    optional_open = {}
    locked_cov = {}
    for (p, d), code in sol.assignment.items():
        role = t.covered_role(code)
        if role:
            locked_cov[(role, d)] = locked_cov.get((role, d), 0) + 1
    prn_req = {(r, d): n for r, d, n in sol.prn_requests}
    for role in t.optional_roles:
        for d in t.dates:
            need = t.required(role, d)
            have = locked_cov.get((role, d), 0) + prn_req.get((role, d), 0)
            if need > have:
                optional_open[(role, d)] = need - have
    n_prj = sum(1 for k in sol.solver_filled if sol.assignment[k] == t.settings.project_code)
    deviations = [r for r in sol.target_rows if r["deviation"]]
    lines = [
        ("Template", t.source.name),
        ("Horizon", f"{t.dates[0]:%m/%d/%Y} - {t.dates[-1]:%m/%d/%Y} ({len(t.dates)} days)"),
        ("Solver status", f"{sol.status} in {sol.seconds:.0f}s"
         + (" (best possible found)" if sol.status == "OPTIMAL" else
            f" (time limit reached; within {gap:,.0f} points of the best possible)" if gap is not None else "")),
        ("Cells filled by solver", len(sol.solver_filled)),
        ("  of which project days", n_prj),
        ("Required slots left empty", required_unfilled),
        ("PRN shifts needed", f"{sum(n for _, _, n in sol.prn_requests)} (see PRN NEEDED under the schedule "
                              f"and the PRN Requests sheet)" if sol.prn_requests else 0),
        ("Shifts Fulfilled matrix", f"{met} required role-days met, {short} short (red)"),
        ("Optional slots left empty", sum(optional_open.values())),
        ("Pay-period target deviations", f"{len(deviations)} ({short_periods} short; see the days/target "
                                         f"columns right of the schedule)" if deviations else 0),
        ("Top-aptitude shift balance", (f"largest gap from fair share: "
                                        f"{max(abs(b['shifts'] - b['share']) for b in sol.balance):.1f} shifts "
                                        f"(see Shift Balance)") if sol.balance else "off or nothing to balance"),
        ("Shortest rest between shifts", (f"{min(g for *_, g, _ in sol.short_rests):g} h"
                                          if sol.short_rests else
                                          f"{t.settings.preferred_rest_hours} h or more everywhere")
         if t.shift_times else "not checked (no shift times)"),
        (f"Rests under {t.settings.preferred_rest_hours} h made by the program",
         sum(1 for *_, made in sol.short_rests if made)),
        ("Runs longer than preferred", len(sol.long_runs)),
        ("Weekend comp Mondays worked", sum(1 for _, _, lbl in sol.comp_misses if lbl == "Monday")),
        ("", ""),
        ("Settings used", f"from {t.settings_source}"),
        *[(f"  {lbl}", f"{getattr(t.settings, attr):%m/%d/%Y}" if kind == "date"
           else display_value(kind, getattr(t.settings, attr)))
          for lbl, attr, kind, _ in SETTINGS_SPEC],
        ("  Optional roles", ", ".join(sorted(t.optional_roles)) or "none"),
        ("", ""),
        *([("Schedule colors", "Role colors from the template's Color cells (Requirements / Other Codes). "
                                "See the Aptitude View sheet for aptitude coloring.")]
          if t.code_styles else
          [("Color key (solver-filled cells only)", ""),
           ("Aptitude 4", "apt4"), ("Aptitude 3", "apt3"), ("Aptitude 2", "apt2"), ("Aptitude 1", "apt1"),
           ("Project day (PRJ)", "prj"), ("Locked cells", "kept as entered, no color change")]),
    ]
    for r, (k, v) in enumerate(lines, 1):
        sm.cell(r, 1, k).font = BOLD if k and not k.startswith("  ") else Font()
        if v in FILLS:
            sm.cell(r, 2, "").fill = FILLS[v]
        else:
            sm.cell(r, 2, v)
    sm.column_dimensions["A"].width = 38
    sm.column_dimensions["B"].width = 70

    # ---- Shift balance among top-rated pharmacists
    if sol.balance:
        sb = wb.create_sheet("Shift Balance")
        rows = []
        for b in sol.balance:
            diff = b["shifts"] - b["share"]
            rows.append([b["role"], t.label(b["pid"]), b["shifts"], round(b["share"], 1), round(diff, 1)])
        last = _table(sb, ["Role", "Pharmacist", "Shifts in role", "Fair share", "Difference"], rows)
        for i, row in enumerate(rows, 2):
            if abs(row[4]) >= 2:
                sb.cell(i, 5).fill = FILLS["over"]
        sb.cell(last + 2, 1, f"Pharmacists with aptitude {t.settings.balance_aptitude} in each role, over "
                             f"the whole schedule. Fair share: the role's shifts among them split "
                             f"{'equally' if t.settings.balance_mode == 'Equal' else 'by days available (days per pay period minus PTO)'}. "
                             f"Yellow: 2 or more shifts from fair share (usually forced by filled-in "
                             f"cells, PTO or other rules).").font = Font(italic=True)

    # ---- PRN requests (contact list)
    if sol.prn_requests or t.prn_pids:
        pq = wb.create_sheet("PRN Requests")
        rows = []
        for role, d, n in sorted(sol.prn_requests, key=lambda k: (k[1], k[0])):
            able = [t.label(p) for p in t.prn_pids
                    if t.staff[p].aptitude.get(role, 0) > 0 and (p, d) not in t.locked]
            rows.append([d.strftime("%a %m/%d/%Y"), role, n, t.role_descriptions.get(role, ""),
                         ", ".join(able)])
        last = _table(pq, ["Date", "Role", "PRNs needed", "Shift / notes", "PRNs able to work it"],
                      rows or [["No PRN shifts needed", "", "", "", ""]])
        pq.cell(last + 2, 1, "When a PRN accepts, enter the role code in their row on the Schedule sheet "
                             "and run the program again: the request disappears and nothing else moves "
                             "unless it has to.").font = Font(italic=True)

    # ---- Pay periods
    pp = wb.create_sheet("Pay Periods")
    rows = [[r["pid"], t.staff[r["pid"]].name, f"{_md(r['start'])}-{_md(r['end'])}" + ("" if r["complete"] else " (partial)"),
             r["target"] if r["target"] is not None else "", r["actual"], r["deviation"], r["project_days"]]
            for r in sol.target_rows]
    _table(pp, ["PID", "Name", "Pay period", "Target", "Days (incl. PTO)", "Deviation", "Project days"], rows)
    for i, r in enumerate(sol.target_rows, 2):
        if r["deviation"]:
            pp.cell(i, 6).fill = FILLS["unfilled"]

    # ---- Role counts per pharmacist
    rc = wb.create_sheet("Role Counts")
    codes = t.roles + [t.settings.project_code] + [c for c in t.codes if c != t.settings.project_code]
    used = {c for c in sol.assignment.values()}
    codes = [c for c in codes if c in used or c in t.roles]
    rows = []
    for p in t.schedule_order:
        counts = [sum(1 for d in t.dates if sol.assignment.get((p, d)) == c) for c in codes]
        rows.append([p, t.staff[p].name if p in t.staff else ""] + [n or "" for n in counts] + [sum(counts)])
    _table(rc, ["PID", "Name"] + codes + ["Total"], rows)

    # ---- Coverage gaps
    cg = wb.create_sheet("Coverage Gaps")
    rows = [[_md(d), d.strftime("%a"), role, "required", n] for role, d, n in sorted(sol.unfilled, key=lambda k: (k[1], k[0]))]
    rows += [[_md(d), d.strftime("%a"), role, "optional", n] for (role, d), n in sorted(optional_open.items(), key=lambda k: (k[0][1], k[0][0]))]
    _table(cg, ["Date", "Day", "Role", "Type", "Empty slots"], rows or [["", "", "None", "", ""]])

    # ---- Rule check
    ck = wb.create_sheet("Rule Check")
    rows = [[i.level, i.category, i.message] for i in issues]
    rows += [["solver", "long run", f"{t.label(p)}: {n} on-duty days in a row starting {_md(s)} "
              f"(preferred max {t.settings.max_consecutive_soft})"] for p, s, n in sol.long_runs]
    rows += [["solver" if made else "input", "rest",
              f"{t.label(p)}: {a} {_md(d)} then {b} {_md(d + timedelta(days=1))}: {gap:g} h between shifts"
              + ("" if made else " (both filled in)")]
             for p, d, a, b, gap, made in sorted(sol.short_rests, key=lambda k: k[4])]
    rows += [["solver", "weekend comp", f"{t.label(p)} works {lbl} {_md(d)} after working the weekend"]
             for p, d, lbl in sol.comp_misses if lbl == "Monday"]
    _table(ck, ["Level", "Category", "Detail"], rows or [["", "", "No issues"]])

    wb.save(out_path)
    return out_path
