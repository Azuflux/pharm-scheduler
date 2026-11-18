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
from .live_blocks import add_live_blocks, inject_cached_values, remove_old_blocks
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
BOLD = Font(bold=True)
PRN_FONT = Font(bold=True, color="FFFFFF")
PRN_ROW_LABEL = "PRN NEEDED"


def _prn_text(reqs: list[tuple[str, int]]) -> str:
    """'CH', 'CH, CP', '2 CH' for one day's PRN requests."""
    return ", ".join(f"{n} {r}" if n > 1 else r for r, n in reqs)


def _coverage_counts(t: Template, sol: Solution) -> tuple[int, int, int]:
    """(required role-days met, required role-days short, PRN shifts requested)."""
    have: dict = {}
    for (p, d), code in sol.assignment.items():
        role = t.covered_role(code)
        if role:
            have[(role, d)] = have.get((role, d), 0) + 1
    prn = {(r, d): n for r, d, n in sol.prn_requests}
    met = short = 0
    for role in t.requirements:
        if role in t.optional_roles:
            continue
        for d in t.dates:
            need = t.required(role, d)
            if not need:
                continue
            if have.get((role, d), 0) >= need:
                met += 1
            elif have.get((role, d), 0) + prn.get((role, d), 0) < need:
                short += 1
    return met, short, sum(prn.values())


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
        removed = remove_old_blocks(t, ws)
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
        removed = (0, 0)
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

    cached = add_live_blocks(t, wb, ws, sol.assignment, sol.prn_requests,
                             with_prn_row=bool(t.prn_pids) or bool(sol.prn_requests), removed=removed)
    short_periods = sum(1 for r in sol.target_rows if r["deviation"] < 0)
    met, short, prn_slots = _coverage_counts(t, sol)
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
    inject_cached_values(out_path, ws.title, cached)
    return out_path
