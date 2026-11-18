"""Build the multi-sheet .xlsx template (Instructions, Schedule, Aptitude, Requirements, Settings)
from a loaded Template. Used to convert older single-sheet templates, and reusable by a web app."""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import config
from copy import copy
from collections import Counter

from .live_blocks import refresh_file
from .loader import Template, cell_style, style_signature, write_applied_styles
from .settings import SPEC as SETTINGS_SPEC, display_value

BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
NOTE = Font(italic=True, color="595959")
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
WEEKEND_FILL = PatternFill("solid", fgColor="EDEDED")
THIN = Side(style="thin", color="BFBFBF")
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
WRAP = Alignment(wrap_text=True, vertical="top")

# Headings on the Instructions sheet must never equal a block label (the loader skips that
# sheet anyway, but this keeps it unambiguous).
INSTRUCTIONS = [
    ("title", "Pharmacist Schedule Template"),
    ("p", "This workbook is the input for the scheduling program. You fill in what is already "
          "decided, and the program fills every blank cell on the Schedule sheet, then returns a new "
          "workbook. Your entries are never changed."),
    ("h", "How to use it"),
    ("n", "1. On the Schedule sheet, enter anything already settled: PTO, rotations (nights, ER, NW), "
          "weekend assignments, fixed weeks. Leave everything else blank."),
    ("n", "2. Check the Aptitude, Requirements and Settings sheets are current (who can do what, what "
          "must be covered each day, and department rules)."),
    ("n", "3. Run the program on this file. It returns a copy with the blanks filled in."),
    ("n", "4. Review the output. The schedule uses your role colors (see Colors below); the Aptitude "
          "View sheet shows the same schedule colored by aptitude. Under the schedule, the Shifts Fulfilled table shows "
          "each role's coverage per day: red means a required role was not covered. The Summary, Pay "
          "Periods, Role Counts, Coverage Gaps and Rule Check sheets explain the result."),
    ("n", "5. To adjust, edit cells in the output and run it again. Everything already in the "
          "schedule is kept, so you can lock in the parts you like and re-run the rest. The days/target "
          "columns and the Shifts Fulfilled table recalculate as you type, so you can see what an edit "
          "does before you re-run (see Live checks while you edit)."),
    ("h", "The sheets"),
    ("b", "Schedule"),
    ("p", "One row per pharmacist (Name, then PID), one column per date. Each cell holds a code: a role "
          "code from the Requirements sheet, or PTO (paid time off: counts toward the pay-period days, "
          "not a workday), OFF (a day off that does not count), or PRJ (project day: counts as a workday). "
          "Codes are not case-sensitive. Row 1 holds the dates; row 2 the day of the week."),
    ("b", "Aptitude"),
    ("p", "One row per pharmacist. Days Per Pay Period is how many days that person works in each pay "
          "period (PTO counts). 0 or blank means the program never adds shifts for them: only what you "
          "enter on the Schedule. Each role column scores how suitable the person is for that role: "
          "4 = first choice, 3, 2, 1 = last resort, 0 = never. The program never assigns a 0, but anything "
          "you type on the Schedule is kept even if it is a 0."),
    ("b", "Requirements"),
    ("p", "One row per role: its Color (see Colors below) and how many pharmacists that role needs on "
          "each day of the week. Optional = "
          "Yes means the role is nice to have but not required: it is filled only when someone would "
          "otherwise get a project day, and an unfilled optional role is not counted as a problem. The "
          "Description column is for your notes; the program ignores it."),
    ("b", "Settings"),
    ("p", "Department rules: pay-period dates, how many days in a row someone may work, weekend comp "
          "days, project-day spreading and how long the program searches. Change only the Value column. "
          "Each row explains what it does."),
    ("h", "Shift times and rest"),
    ("p", "Start and End on the Requirements sheet (and in Other Codes) give each shift's hours; an End "
          "earlier than Start means the shift ends the next morning (N: 2100 to 0700). The program never "
          "schedules a shift that starts within the Settings' minimum hours of the previous day's shift "
          "ending, and avoids short gaps up to the preferred hours. Leave Start/End blank for codes with "
          "no fixed hours; those aren't checked. Type times as 0630 or 6:30."),
    ("h", "Days per pay period"),
    ("p", "The program never schedules anyone for more days than their Days Per Pay Period. If someone "
          "asks for an extra day, type it on the Schedule: entries you type are always kept, and the "
          "program fills around them."),
    ("h", "Live checks while you edit"),
    ("p", "Right of the dates, each pharmacist's days given / target per pay period and in total; under "
          "the schedule, the Shifts Fulfilled table (filled / required for every role and day). Both are "
          "formulas: they update as you type codes on the Schedule or change targets (Aptitude) and "
          "requirements (Requirements), so you can see the effect of a change before running the program "
          "again. Green = exactly right, red = short, yellow = over; in Shifts Fulfilled, gray = an optional "
          "role left open, purple = a PRN was requested on the last run and hasn't been entered yet. Each "
          "pharmacist's row finds its target on Aptitude by the PID in that row (shown in the Target "
          "column), and each Shifts Fulfilled row finds its numbers on Requirements by role name, so "
          "you can move or sort rows on any sheet. The program rebuilds both blocks on every run, so "
          "don't type inside them."),
    ("h", "Sharing top-aptitude work"),
    ("p", "For each role, pharmacists scored at the Settings' balance level (4 by default) share that "
          "role's shifts as evenly as the other rules allow over the whole schedule; for example, two "
          "transplant pharmacists each get about half the TX shifts. 'By days available' adjusts for "
          "part-time hours and PTO; 'Equal' doesn't. The Shift Balance sheet in the output shows the "
          "result. It is a preference: coverage, pay-period limits and rest always come first."),
    ("h", "PRN pharmacists"),
    ("p", "PRN pharmacists are listed under the PRN STAFF row on the Schedule sheet and marked PRN = Yes "
          "on the Aptitude sheet, where their scores say which roles they can work. The program never "
          "assigns a particular PRN. When regular staff can't cover a role a PRN can work, it shows a PRN "
          "NEEDED line under the schedule and lists each shift on the PRN Requests sheet, with the PRNs "
          "who could take it. Once a PRN accepts, type the role code in their row and run the program "
          "again; the request disappears. Codes typed in PRN rows (availability, X, booked shifts) are "
          "respected: a PRN already entered on a day isn't counted as available that day."),
    ("h", "Other codes"),
    ("p", "Every code that isn't a role must be listed in the Other Codes table under the requirement "
          "matrix (Requirements sheet), with whether it counts toward days per pay period, whether it is "
          "a workday, and which role it also covers, if any (for example, ON8 covers ON). The program "
          "reports any code it doesn't recognize, with the cell, instead of guessing."),
    ("h", "Colors"),
    ("p", "The Color column on the Requirements sheet (and the Other Codes table under it, for PTO, OFF "
          "and PRJ) sets how each code looks on the finished schedule. Format the Color cell exactly as you "
          "want that role to appear: fill color, text color, bold. The program copies that formatting onto "
          "every cell with that code. A code with an unformatted Color cell is left plain. Cells you format "
          "by hand on the Schedule keep their own formatting."),
    ("p", "The output also has an Aptitude View sheet: the same schedule colored by aptitude score, for "
          "whoever builds the schedule to review. The Schedule sheet is the one to publish."),
    ("h", "Common changes"),
    ("b", "Add a pharmacist"),
    ("p", "Add a row at the bottom of the Schedule sheet (Name, PID) and a row at the bottom of the "
          "Aptitude sheet with the same PID, their days per pay period and a score for every role. PIDs "
          "must be unique; the PID is what connects the two sheets. On the Schedule sheet, insert a "
          "row (right-click a row number, Insert) rather than typing into the blank row above Shifts "
          "Fulfilled, so the table stays one row below the list. The live checks pick up the new person on the next run (or use 'Only update the live "
          "checks' in the web app)."),
    ("b", "Remove a pharmacist"),
    ("p", "Delete their row on both the Schedule and Aptitude sheets. To keep them listed but stop "
          "scheduling them, set Days Per Pay Period to 0."),
    ("b", "Add a role"),
    ("p", "On the Aptitude sheet, add a column at the right end with the new role code in row 2 and a "
          "score for every pharmacist. On the Requirements sheet, add a row with the same code and the "
          "number needed each day, and format its Color cell. Inserting whole rows or columns is safe on these sheets, as long as "
          "no blank row or column is left inside a table. The Shifts Fulfilled table gets a row for the new "
          "role on the next run (or use 'Only update the live checks')."),
    ("b", "Remove a role"),
    ("p", "Delete its column on the Aptitude sheet and its row on the Requirements sheet. Make sure the "
          "code no longer appears on the Schedule."),
    ("b", "Start a new schedule period"),
    ("p", "Put the first date in the first date cell of row 1 and fill the dates to the right (drag the "
          "fill handle). Update row 2 with the day of each date (MON, TUE...), or use a formula such as "
          "=UPPER(TEXT(C1,\"ddd\")). Clear the old codes and enter the new period's settled assignments."),
    ("h", "Rules the program follows"),
    ("p", "Always: keeps every cell you filled in; gives each person at most one code per day; never "
          "assigns a role scored 0; covers each required role exactly (a slot is left empty only if no "
          "one can fill it, and it then shows red); never exceeds an optional role's number; never "
          "schedules more consecutive workdays than the Settings limit."),
    ("p", "As far as possible: hits each person's days per pay period; prefers higher aptitude scores and "
          "uses 1s only as a last resort; fills optional roles before giving project days; spreads project "
          "days so no one is stuck with them; keeps runs to the preferred length; gives full-timers the "
          "Monday (and ideally Tuesday) off after a worked weekend; keeps people in the same role on "
          "back-to-back days."),
    ("h", "Please don't"),
    ("p", "Rename or move the labeled cells the program looks for: the table titles on each sheet, Date "
          "and Day on the Schedule sheet, the Name / PID / Days Per Pay Period headers, Role / Optional on "
          "Requirements, and Setting / Value on Settings. Leave blank rows or columns inside a table (the "
          "program stops at a gap and will tell you where it is)."),
]

SHEET_NOTES = {
    "Aptitude": "One row per pharmacist. Scores: 4 first choice, 3, 2, 1 last resort, 0 never. "
                "Days Per Pay Period: 0 or blank = only scheduled where you fill in the Schedule. "
                "New pharmacists and roles: add at the bottom / right end.",
    "Requirements": "Pharmacists needed per role per weekday. Color: format the cell as the role should "
                    "look on the schedule. Optional = Yes: filled only when someone is free.",
    "Settings": "Department rules. Change only the Value column; each row explains itself.",
}


def _style(cell, st: dict) -> None:
    cell.fill = copy(st["fill"])
    cell.font = copy(st["font"])


def _signature(st: dict) -> str:
    return style_signature(st)


def _source_style(t: Template, pid: str, d) -> dict | None:
    if t.style_book is None or t.sheet_name not in getattr(t.style_book, "sheetnames", []):
        return None
    return cell_style(t.style_book[t.sheet_name].cell(t.row_of[pid] + 1, t.col_of[d] + 1))


def learn_styles(t: Template) -> dict[str, dict]:
    """Most common formatting per code among the schedule's filled-in cells (when the source is a
    formatted .xlsx). Lets a color-coded schedule hand its colors to the new Color cells."""
    if t.style_book is None or t.sheet_name not in getattr(t.style_book, "sheetnames", []):
        return {}
    ws = t.style_book[t.sheet_name]
    seen: dict[str, Counter] = {}
    example: dict[tuple, dict] = {}
    for (pid, d), code in t.locked.items():
        st = cell_style(ws.cell(t.row_of[pid] + 1, t.col_of[d] + 1))
        if not st:
            continue
        sig = _signature(st)
        seen.setdefault(code, Counter())[sig] += 1
        example.setdefault(sig, st)
    return {code: example[c.most_common(1)[0][0]] for code, c in seen.items()}


def _write_times(ws, r: int, c: int, times) -> None:
    """Start/End as Excel times shown like 0630 (end past midnight shown as its clock time)."""
    from datetime import time as dtime
    for k, mins in enumerate(times or (None, None)):
        cell = ws.cell(r, c + k)
        cell.alignment = CENTER
        cell.number_format = "hhmm"
        if mins is not None:
            mins %= 1440
            cell.value = dtime(mins // 60, mins % 60)


def _widths(ws, widths: dict[int, float]) -> None:
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w


DEFAULT_LOOK = {
    "title_font": Font(bold=True, size=16),
    "label_font": BOLD,
    "date_font": BOLD,
    "day_font": BOLD,
    "date_format": "d-mmm",
    "name_font": Font(),
    "pid_font": Font(color="808080"),
    "cell_font": Font(),
    "section_font": BOLD,
    "section_fill": PatternFill("solid", fgColor="BFBFBF"),
    "shade_weekend_headers": True,
    "date_col_width": 7,
}


def write_instructions(ins) -> None:
    """(Re)write the Instructions sheet text."""
    for row in list(ins.iter_rows()):
        for cell in row:
            cell.value = None
            cell.font = Font()
            ins.row_dimensions[cell.row].height = None
    ins.column_dimensions["A"].width = 110
    r = 1
    for kind, text in INSTRUCTIONS:
        if kind == "h":
            r += 1
        cell = ins.cell(r, 1, text)
        cell.alignment = WRAP
        if kind == "title":
            cell.font = TITLE
        elif kind in ("h",):
            cell.font = Font(bold=True, size=12, color="1F4E78")
        elif kind == "b":
            cell.font = BOLD
        lines = max(1, len(text) // 105 + 1)
        ins.row_dimensions[r].height = 15 * lines + 2
        r += 1


def build_template(t: Template, out_path: str | Path, look: dict | None = None) -> Path:
    out_path = Path(out_path)
    wb = openpyxl.Workbook()
    styles = {**learn_styles(t), **{k: v for k, v in t.code_styles.items()}}  # Color cells win

    # ---------------- Instructions
    ins = wb.active
    ins.title = config.INSTRUCTIONS_SHEET
    write_instructions(ins)

    # ---------------- Schedule
    lk = {**DEFAULT_LOOK, **(look or {})}
    ws = wb.create_sheet("Schedule")
    NAME_C, PID_C, D0 = 1, 2, 3
    TITLE_R, DATE_R, DAY_R = 1, 2, 3
    title = ws.cell(TITLE_R, NAME_C, f"Pharmacist Schedule {t.dates[0]:%B} {t.dates[0].day} - "
                                     f"{t.dates[-1]:%B} {t.dates[-1].day}, {t.dates[-1].year}")
    title.font = copy(lk["title_font"])
    for r_, lbl in ((DATE_R, config.LABEL_SCHEDULE_DATES), (DAY_R, config.LABEL_SCHEDULE_DAYS)):
        cell = ws.cell(r_, PID_C, lbl)
        cell.font, cell.alignment = copy(lk["label_font"]), CENTER
    ws.cell(DATE_R, NAME_C, config.LABEL_NAME).font = copy(lk["label_font"])
    for j, d in enumerate(t.dates):
        c = D0 + j
        dc = ws.cell(DATE_R, c, d)
        dc.number_format = lk["date_format"]
        wc = ws.cell(DAY_R, c, d.strftime("%a").upper())
        dc.font, wc.font = copy(lk["date_font"]), copy(lk["day_font"])
        for cell in (dc, wc):
            cell.alignment = CENTER
            if lk["shade_weekend_headers"] and d.weekday() >= 5:
                cell.fill = WEEKEND_FILL
        dc.border = Border(bottom=Side(style="medium"))
        wc.border = Border(bottom=Side(style="thin"))
    order = [p for p in t.schedule_order if not t.is_prn(p)] + [p for p in t.schedule_order if t.is_prn(p)]
    r = DAY_R + 1
    for pid in order:
        if t.is_prn(pid) and pid == next(p for p in order if t.is_prn(p)):
            for c in range(NAME_C, D0 + len(t.dates)):      # "PRN STAFF" separator row
                sep = ws.cell(r, c)
                sep.fill = copy(lk["section_fill"])
                sep.border = Border(bottom=Side(style="medium"))
            sep = ws.cell(r, NAME_C, config.LABEL_PRN_SECTION)
            sep.font = copy(lk["section_font"])
            r += 1
        s = t.staff.get(pid)
        nm = ws.cell(r, NAME_C, s.name if s else "")
        nm.font, nm.alignment = copy(lk["name_font"]), Alignment(horizontal="right")
        pc = ws.cell(r, PID_C, int(pid) if pid.isdigit() else pid)
        pc.alignment, pc.font = CENTER, copy(lk["pid_font"])
        for j, d in enumerate(t.dates):
            code = t.locked.get((pid, d))
            cell = ws.cell(r, D0 + j, code)
            cell.alignment = CENTER
            cell.border = GRID
            cell.font = copy(lk["cell_font"])
            own = _source_style(t, pid, d) if code else None
            if own:
                _style(cell, own)          # keep the cell's own formatting from the source
            elif code and code in styles:
                _style(cell, styles[code])
        r += 1
    _widths(ws, {NAME_C: 22, PID_C: 7, **{D0 + j: lk["date_col_width"] for j in range(len(t.dates))}})
    ws.freeze_panes = ws.cell(DAY_R + 1, D0)

    # ---------------- Aptitude
    ap = wb.create_sheet("Aptitude")
    ap.cell(1, 1, config.LABEL_APTITUDE).font = TITLE
    ap.cell(1, 6, SHEET_NOTES["Aptitude"]).font = NOTE
    heads = [config.LABEL_NAME, config.LABEL_PID, config.LABEL_DAYS_PER_PERIOD, config.LABEL_PRN, *t.roles]
    R0 = 5  # first role column
    for c, h in enumerate(heads, 1):
        cell = ap.cell(2, c, h)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ap.row_dimensions[2].height = 30
    pids = [p for p in order if p in t.staff] + [p for p in t.staff if p not in order]
    for i, pid in enumerate(pids):
        r = 3 + i
        s = t.staff[pid]
        ap.cell(r, 1, s.name)
        ap.cell(r, 2, int(pid) if pid.isdigit() else pid).alignment = CENTER
        ap.cell(r, 3, s.target_days).alignment = CENTER
        ap.cell(r, 4, "Yes" if s.prn else "No").alignment = CENTER
        for k, role in enumerate(t.roles):
            cell = ap.cell(r, R0 + k, s.aptitude[role])
            cell.alignment = CENTER
            cell.border = GRID
    last = 2 + max(len(pids), 1)
    end_col = get_column_letter(R0 - 1 + len(t.roles))
    apt_range = f"{get_column_letter(R0)}3:{end_col}{last + 200}"
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="4", allow_blank=True,
                        showErrorMessage=True, errorTitle="Aptitude score",
                        error="Enter a whole number from 0 (never) to 4 (first choice).")
    ap.add_data_validation(dv)
    dv.add(apt_range)
    dv2 = DataValidation(type="whole", operator="between", formula1="0", formula2="31", allow_blank=True,
                         showErrorMessage=True, errorTitle="Days per pay period",
                         error="Enter a whole number of days (0-31), or leave blank.")
    ap.add_data_validation(dv2)
    dv2.add(f"C3:C{last + 200}")
    dv3 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True, showErrorMessage=True,
                         errorTitle="PRN", error="Choose Yes or No.")
    ap.add_data_validation(dv3)
    dv3.add(f"D3:D{last + 200}")
    ap.conditional_formatting.add(apt_range, ColorScaleRule(
        start_type="num", start_value=0, start_color="F2F2F2",
        mid_type="num", mid_value=2, mid_color="FFF2CC",
        end_type="num", end_value=4, end_color="63BE7B"))
    _widths(ap, {1: 22, 2: 7, 3: 12, 4: 7, **{R0 + k: 7 for k in range(len(t.roles))}})
    ap.freeze_panes = f"{get_column_letter(R0)}3"

    # ---------------- Requirements
    rq = wb.create_sheet("Requirements")
    rq.cell(1, 1, config.LABEL_REQUIREMENTS).font = TITLE
    rq.cell(1, 5, SHEET_NOTES["Requirements"]).font = NOTE
    day_heads = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    heads = [config.LABEL_ROLE, config.LABEL_COLOR, config.LABEL_OPTIONAL, config.LABEL_START, config.LABEL_END,
             *day_heads, config.LABEL_DESCRIPTION]
    D1 = 6  # first weekday column
    for c, h in enumerate(heads, 1):
        cell = rq.cell(2, c, h)
        cell.font = BOLD
        cell.fill = WEEKEND_FILL if h in ("Sat", "Sun") else HEADER_FILL
        cell.alignment = CENTER
    for i, role in enumerate(t.requirements):
        r = 3 + i
        rq.cell(r, 1, role).font = BOLD
        sw = rq.cell(r, 2, role)
        sw.alignment = CENTER
        sw.border = GRID
        sw.font = copy(lk["cell_font"])
        if role in styles:
            _style(sw, styles[role])
        rq.cell(r, 3, "Yes" if role in t.optional_roles else "No").alignment = CENTER
        _write_times(rq, r, 4, t.shift_times.get(role))
        for k, n in enumerate(t.requirements[role]):
            cell = rq.cell(r, D1 + k, n)
            cell.alignment = CENTER
            cell.border = GRID
        rq.cell(r, D1 + 7, t.role_descriptions.get(role, ""))
    last = 2 + max(len(t.requirements), 1)
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True, showErrorMessage=True,
                        errorTitle="Optional", error="Choose Yes or No.")
    rq.add_data_validation(dv)
    dv.add(f"C3:C{last}")
    dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                        showErrorMessage=True, errorTitle="Pharmacists needed",
                        error="Enter a whole number of pharmacists (0 or more).")
    rq.add_data_validation(dv)
    dv.add(f"{get_column_letter(D1)}3:{get_column_letter(D1 + 6)}{last}")
    # Other codes, two rows under the requirement table: every non-role code the schedule may use
    o = last + 3
    rq.cell(o, 1, config.LABEL_OTHER_CODES).font = Font(bold=True, size=12)
    rq.cell(o, 4, "Every code that isn't a role. Counts toward days: counts toward days per pay period. "
                  "Workday: counts as a day worked (for days-in-a-row). Covers role: also counts as "
                  "covering that role (e.g. ON8 covers ON).").font = NOTE
    oheads = [config.LABEL_CODE, config.LABEL_COLOR, config.LABEL_COUNTS, config.LABEL_WORKDAY,
              config.LABEL_COVERS, config.LABEL_START, config.LABEL_END, config.LABEL_MEANING]
    for c, h in enumerate(oheads, 1):
        cell = rq.cell(o + 1, c, h)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    rq.row_dimensions[o + 1].height = 30
    yn = lambda v: "Yes" if v else "No"
    for i, (code, info) in enumerate(t.codes.items(), 1):
        r = o + 1 + i
        rq.cell(r, 1, code).font = BOLD
        sw = rq.cell(r, 2, code)
        sw.alignment = CENTER
        sw.border = GRID
        sw.font = copy(lk["cell_font"])
        if code in styles:
            _style(sw, styles[code])
        for c, v in ((3, yn(info["counts_for_pay"])), (4, yn(info["on_duty"])), (5, info.get("covers") or "")):
            cell = rq.cell(r, c, v)
            cell.alignment = CENTER
        _write_times(rq, r, 6, t.shift_times.get(code))
        rq.cell(r, 8, info.get("meaning", ""))
    olast = o + 1 + max(len(t.codes), 1)
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False, showErrorMessage=True,
                        error="Choose Yes or No.")
    rq.add_data_validation(dv)
    dv.add(f"C{o + 2}:D{olast + 30}")
    roles_list = ",".join(t.roles)
    if len(roles_list) < 250:
        dv = DataValidation(type="list", formula1=f'"{roles_list}"', allow_blank=True, showErrorMessage=True,
                            error="Choose a role code, or leave blank.")
        rq.add_data_validation(dv)
        dv.add(f"E{o + 2}:E{olast + 30}")
    _widths(rq, {1: 12, 2: 9, 3: 10, 4: 9, 5: 9, 6: 7, 7: 7, 8: 7, 9: 7, 10: 7, 11: 7, 12: 7, 13: 60})
    rq.freeze_panes = "D3"

    # ---------------- Settings
    st = wb.create_sheet("Settings")
    st.cell(1, 1, config.LABEL_SETTINGS).font = TITLE
    st.cell(1, 3, SHEET_NOTES["Settings"]).font = NOTE
    for c, h in enumerate([config.LABEL_SETTING_NAME, config.LABEL_SETTING_VALUE, "What it does"], 1):
        cell = st.cell(2, c, h)
        cell.font = BOLD
        cell.fill = HEADER_FILL
    for i, (label, attr, kind, help_text) in enumerate(SETTINGS_SPEC):
        r = 3 + i
        st.cell(r, 1, label).alignment = WRAP
        v = st.cell(r, 2, display_value(kind, getattr(t.settings, attr)))
        v.alignment = CENTER
        v.fill = PatternFill("solid", fgColor="FFF2CC")
        v.border = GRID
        if kind == "date":
            v.number_format = "m/d/yyyy"
            dv = DataValidation(type="date", operator="greaterThan", formula1="36526",
                                showErrorMessage=True, error="Enter a date.")
        elif kind == "bool":
            dv = DataValidation(type="list", formula1='"Yes,No"', showErrorMessage=True,
                                error="Choose Yes or No.")
        elif kind.startswith("choice:"):
            opts = kind.split(":", 1)[1].split("|")
            dv = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', showErrorMessage=True,
                                error="Choose one of: " + ", ".join(opts))
        else:
            dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0",
                                allow_blank=(kind == "int_or_blank"), showErrorMessage=True,
                                error="Enter a whole number" + (" or leave blank." if kind == "int_or_blank" else "."))
        st.add_data_validation(dv)
        dv.add(v.coordinate)
        h = st.cell(r, 3, help_text)
        h.alignment = WRAP
        st.row_dimensions[r].height = 15 * max(1, len(help_text) // 80 + 1) + 2
    _widths(st, {1: 42, 2: 14, 3: 85})

    write_applied_styles(wb, styles)
    wb.active = wb.sheetnames.index("Schedule")
    wb.save(out_path)
    refresh_file(out_path)       # live days/target and Shifts Fulfilled blocks
    return out_path


def refresh_template(path: str | Path, out_path: str | Path | None = None) -> Path:
    """Bring an existing template (or schedule) up to date without solving: current Instructions
    text plus freshly built live check blocks. Keeps the sheet's look and every entry."""
    path = Path(path)
    out_path = Path(out_path or path)
    wb = openpyxl.load_workbook(path)
    if config.INSTRUCTIONS_SHEET in wb.sheetnames:
        write_instructions(wb[config.INSTRUCTIONS_SHEET])
    wb.save(out_path)
    return refresh_file(out_path)
