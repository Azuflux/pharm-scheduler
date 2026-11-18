"""Formula-driven blocks on the Schedule sheet, so they update as the schedule is edited in Excel:

  - right of the dates: the names again, then days given / target per pay period and in total
  - under the schedule: Shifts Fulfilled (filled / required per role per day), with an optional
    PRN NEEDED row from the last solver run

Counts are COUNTIF/SUMPRODUCT formulas over the schedule cells; targets and requirements point at
the Aptitude and Requirements cells, so editing those updates the blocks too. Colors are
conditional formatting on the "a/b" text. Used by the output writer and the template builder.
"""
from __future__ import annotations

import re
import zipfile
from copy import copy
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

from . import config
from .loader import Template

FILL = {
    "met": PatternFill("solid", fgColor="C6EFCE", bgColor="C6EFCE"),
    "short": PatternFill("solid", fgColor="FF0000", bgColor="FF0000"),
    "over": PatternFill("solid", fgColor="FFEB9C", bgColor="FFEB9C"),
    "optional_open": PatternFill("solid", fgColor="E7E6E6", bgColor="E7E6E6"),
    "prn": PatternFill("solid", fgColor="7030A0", bgColor="7030A0"),
    "header": PatternFill("solid", fgColor="D9D9D9"),
}
WHITE_BOLD = Font(bold=True, color="FFFFFF")
PRN_ROW_LABEL = "PRN NEEDED"
CENTER = Alignment(horizontal="center", vertical="center")
THIN = Side(style="thin", color="BFBFBF")
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


@dataclass
class Layout:
    """Where things are, in Excel's 1-based rows/columns."""
    ws: object                                   # the Schedule worksheet
    staff_rows: dict[str, int]                   # pid -> row
    date_cols: dict[date, int]                   # date -> column
    date_row: int
    day_row: int
    pid_col: int                                 # column holding PIDs ("Date"/"Day" labels)
    name_col: int | None                         # left Name column, if any
    dpp_ref: dict[str, str] = field(default_factory=dict)          # pid -> "'Aptitude'!$C$5"
    req_ref: dict[tuple[str, int], str] = field(default_factory=dict)  # (role, weekday) -> ref

    @property
    def first_row(self) -> int:
        return min(self.staff_rows.values())

    @property
    def last_row(self) -> int:
        return max(self.staff_rows.values())

    @property
    def first_col(self) -> int:
        return min(self.date_cols.values())

    @property
    def last_col(self) -> int:
        return max(self.date_cols.values())


def sheet_ref(sheet: str | None, row: int, col: int) -> str:
    """Absolute reference like 'Aptitude'!$C$5 (no sheet prefix when sheet is None)."""
    cell = f"${L(col)}${row}"
    if not sheet:
        return cell
    return "'" + sheet.replace("'", "''") + "'!" + cell


def _crit(code: str) -> str:
    """COUNTIF criterion for an exact code: escape wildcards and quotes."""
    esc = code.replace("~", "~~").replace("*", "~*").replace("?", "~?").replace('"', '""')
    return f'"{esc}"'


def _a(x: str) -> str:   # left number of "a/b"
    return f'VALUE(LEFT({x},FIND("/",{x})-1))'


def _b(x: str) -> str:   # right number of "a/b"
    return f'VALUE(MID({x},FIND("/",{x})+1,9))'


def _ratio_rules(ws, rng: str, top_left: str, short_fill=None, short_font=None, prefix_rules=()) -> None:
    """Green when a=b, yellow when a>b, short color when a<b, for cells holding "a/b"."""
    has = f'ISNUMBER(FIND("/",{top_left}))'
    for rule in prefix_rules:
        ws.conditional_formatting.add(rng, rule)
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND({has},{_a(top_left)}<{_b(top_left)})"],
        fill=short_fill or FILL["short"], font=short_font or WHITE_BOLD, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND({has},{_a(top_left)}={_b(top_left)})"], fill=FILL["met"], stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND({has},{_a(top_left)}>{_b(top_left)})"], fill=FILL["over"], stopIfTrue=True))


def clear_block_formatting(ws, below_row: int, from_col: int) -> None:
    """Drop conditional formatting on the output blocks (rows below the schedule, or columns at/after
    the right-hand Name column), so re-runs don't stack rules."""
    from openpyxl.formatting.formatting import ConditionalFormattingList
    keep = ConditionalFormattingList()
    for cf in ws.conditional_formatting:
        ranges = [r for r in cf.sqref.ranges if r.min_row <= below_row and r.min_col < from_col]
        for r in ranges:
            for rule in cf.rules:
                keep.add(r.coord, rule)
    ws.conditional_formatting = keep


# --------------------------------------------------------------------------- days / target

def write_days_block(t: Template, lay: Layout, base_font: Font) -> dict[str, object]:
    """Right-hand block. Returns {cell: expected value} for the cached results."""
    ws = lay.ws
    fname, fsize = base_font.name or "Calibri", base_font.sz or 11
    name_col = lay.last_col + 1
    periods = t.periods()
    tot_c = name_col + len(periods) + 1
    nonpay = [c for c, info in t.codes.items() if not info["counts_for_pay"]]
    cached: dict[str, object] = {}

    def count_expr(rng: str) -> str:
        return f"SUMPRODUCT(--(LEN(TRIM({rng}))>0))" + "".join(f"-COUNTIF({rng},{_crit(c)})" for c in nonpay)

    head = ws.cell(lay.date_row, name_col, config.LABEL_NAME)
    head.font = Font(name=fname, sz=fsize, b=True)
    ws.column_dimensions[L(name_col)].width = 22
    pcols = []
    for k, per in enumerate(periods, 1):
        c = name_col + k
        cols = [lay.date_cols[d] for d in per["dates"]]
        pcols.append((per, min(cols), max(cols)))
        lbl = f"{per['start'].month}/{per['start'].day}-{per['end'].month}/{per['end'].day}"
        h = ws.cell(lay.date_row, c, lbl + ("" if per["complete"] else " (partial)"))
        h.font, h.alignment, h.fill = Font(name=fname, sz=fsize, b=True), CENTER, FILL["header"]
        d = ws.cell(lay.day_row, c, "days/target")
        d.font, d.alignment = Font(name=fname, sz=max(7, fsize - 2), i=True), CENTER
        ws.column_dimensions[L(c)].width = 11
    h = ws.cell(lay.date_row, tot_c, "Total")
    h.font, h.alignment, h.fill = Font(name=fname, sz=fsize, b=True), CENTER, FILL["header"]
    d = ws.cell(lay.day_row, tot_c, "days/target")
    d.font, d.alignment = Font(name=fname, sz=max(7, fsize - 2), i=True), CENTER
    ws.column_dimensions[L(tot_c)].width = 11

    # PRN STAFF separator rows: carry the band across
    if lay.name_col:
        for r in range(lay.first_row, lay.last_row + 1):
            left = ws.cell(r, lay.name_col)
            if str(left.value or "").strip().upper() == config.LABEL_PRN_SECTION:
                for c in range(name_col, tot_c + 1):
                    ws.cell(r, c).fill = copy(left.fill)
                lab = ws.cell(r, name_col, config.LABEL_PRN_SECTION)
                lab.font = copy(left.font)

    ratio_cells = []
    for pid, r in lay.staff_rows.items():
        s = t.staff.get(pid)
        nm = ws.cell(r, name_col, (s.name if s and s.name else pid))
        nm.font = Font(name=fname, sz=fsize, b=True)
        has_target = bool(s and s.target_days and not s.prn and pid in lay.dpp_ref)
        tparts = []
        for k, (per, c1, c2) in enumerate(pcols, 1):
            cell = ws.cell(r, name_col + k)
            cell.alignment, cell.border = CENTER, GRID
            cnt = count_expr(f"{L(c1)}{r}:{L(c2)}{r}")
            n = sum(1 for d in per["dates"] if (pid, d) in _assign(t)
                    and t.code_counts_for_pay(_assign(t)[pid, d]))
            if not has_target:
                cell.value = f'=IF(({cnt})=0,"",{cnt})'
                cell.font = Font(name=fname, sz=fsize, color="808080")
                cached[cell.coordinate] = n if n else ""
                continue
            ref = lay.dpp_ref[pid]
            tgt = ref if per["complete"] else f"ROUND({ref}*{len(per['dates'])}/{t.settings.pay_period_days},0)"
            tparts.append(tgt)
            cell.value = f'={cnt}&"/"&{tgt}'
            cell.font = Font(name=fname, sz=fsize, b=True)
            target = s.target_days if per["complete"] else round(
                s.target_days * len(per["dates"]) / t.settings.pay_period_days)
            cached[cell.coordinate] = f"{n}/{target}"
            ratio_cells.append(cell)
        cell = ws.cell(r, tot_c)
        cell.alignment, cell.border = CENTER, GRID
        full = f"{L(lay.first_col)}{r}:{L(lay.last_col)}{r}"
        n_all = sum(1 for d in t.dates if (pid, d) in _assign(t) and t.code_counts_for_pay(_assign(t)[pid, d]))
        if has_target:
            cell.value = f'={count_expr(full)}&"/"&({"+".join(tparts)})'
            cell.font = Font(name=fname, sz=fsize, b=True)
            tot_t = sum(s.target_days if per["complete"] else round(
                s.target_days * len(per["dates"]) / t.settings.pay_period_days) for per, _, _ in pcols)
            cached[cell.coordinate] = f"{n_all}/{tot_t}"
        else:
            cell.value = f'=IF(({count_expr(full)})=0,"",{count_expr(full)})'
            cell.font = Font(name=fname, sz=fsize, color="808080")
            cached[cell.coordinate] = n_all if n_all else ""

    rng = f"{L(name_col + 1)}{lay.first_row}:{L(tot_c)}{lay.last_row}"
    _ratio_rules(ws, rng, f"{L(name_col + 1)}{lay.first_row}")
    if lay.date_row > 1 and ws.cell(lay.date_row - 1, name_col + 1).value in (None, ""):
        note = ws.cell(lay.date_row - 1, name_col + 1,
                       "Days given / target, updating as you edit (PTO counts). Green = met, red = short, "
                       "yellow = over. Gray = PRN or no target.")
        note.font = Font(name=fname, sz=max(7, fsize - 2), i=True, color="595959")
    return cached


# --------------------------------------------------------------------------- shifts fulfilled

def insert_fulfilled_rows(t: Template, lay: Layout, with_prn_row: bool) -> tuple[int, int, int]:
    """Make room under the schedule. Returns (title row, rows inserted, first row that moved)."""
    ws = lay.ws
    block = 1 + int(with_prn_row) + len(t.requirements) + 1 + 1
    start = lay.last_row + 2
    blank = all(ws.cell(lay.last_row + 1, c).value in (None, "")
                for c in [lay.pid_col, *lay.date_cols.values()])
    if blank:
        ws.insert_rows(start, block)
        return start, block, start
    ws.insert_rows(lay.last_row + 1, block + 1)
    return start, block + 1, lay.last_row + 1


def write_fulfilled_block(t: Template, lay: Layout, start: int, base_font: Font,
                          prn_requests: list[tuple[str, date, int]] = (),
                          with_prn_row: bool = False) -> dict[str, object]:
    """Shifts Fulfilled table at `start` (rows already inserted). Returns expected cached values."""
    ws = lay.ws
    fname, fsize = base_font.name or "Calibri", base_font.sz or 11
    roles = list(t.requirements)
    label_col = lay.pid_col
    cached: dict[str, object] = {}
    covers: dict[str, list[str]] = {r: [r] for r in roles}
    for code, info in t.codes.items():
        if info.get("covers") in covers:
            covers[info["covers"]].append(code)
    prn = {(r, d): n for r, d, n in prn_requests}

    title = ws.cell(start, label_col, config.LABEL_SHIFTS_FULFILLED)
    title.font = Font(name=fname, sz=fsize, b=True)
    ws.column_dimensions[L(label_col)].width = max(ws.column_dimensions[L(label_col)].width or 0, 16)
    for d, c in lay.date_cols.items():
        h = ws.cell(start, c, f"{d.month}/{d.day}")
        h.font, h.alignment, h.fill = Font(name=fname, sz=fsize, b=True), CENTER, FILL["header"]

    if with_prn_row:
        pr = start + 1
        lab = ws.cell(pr, label_col, PRN_ROW_LABEL)
        lab.font, lab.fill = Font(name=fname, sz=fsize, b=True, color="FFFFFF"), FILL["prn"]
        for d, c in lay.date_cols.items():
            reqs = [(r_, n) for (r_, dd), n in sorted(prn.items()) if dd == d]
            cell = ws.cell(pr, c)
            cell.alignment = CENTER
            if reqs:
                cell.value = ", ".join(f"{n} {r_}" if n > 1 else r_ for r_, n in reqs)
                cell.fill, cell.font = FILL["prn"], Font(name=fname, sz=fsize, b=True, color="FFFFFF")

    first_role = start + 1 + int(with_prn_row)
    r1, r2 = lay.first_row, lay.last_row
    assign = _assign(t)
    for i, role in enumerate(roles):
        r = first_role + i
        ws.cell(r, label_col, role).font = Font(name=fname, sz=fsize, b=True)
        optional = role in t.optional_roles
        prn_cells = []
        for d, c in lay.date_cols.items():
            col = L(c)
            have = "+".join(f"COUNTIF({col}${r1}:{col}${r2},{_crit(code)})" for code in covers[role])
            need = lay.req_ref.get((role, d.weekday()), "0")
            cell = ws.cell(r, c)
            cell.alignment = CENTER
            cell.font = Font(name=fname, sz=fsize)
            cell.value = f'=IF(AND(({have})=0,N({need})=0),"",({have})&"/"&N({need}))'
            n_have = sum(1 for (p, dd), code in assign.items() if dd == d and code in covers[role])
            n_need = t.required(role, d)
            cached[cell.coordinate] = "" if (n_have == 0 and n_need == 0) else f"{n_have}/{n_need}"
            if prn.get((role, d)):
                prn_cells.append(cell.coordinate)
        rng = f"{L(lay.first_col)}{r}:{L(lay.last_col)}{r}"
        tl = f"{L(lay.first_col)}{r}"
        for coord in prn_cells:   # purple until the PRN shift is entered
            ws.conditional_formatting.add(coord, FormulaRule(
                formula=[f'AND(ISNUMBER(FIND("/",{coord})),{_a(coord)}<{_b(coord)})'],
                fill=FILL["prn"], font=WHITE_BOLD, stopIfTrue=True))
        if optional:
            _ratio_rules(ws, rng, tl, short_fill=FILL["optional_open"], short_font=Font(color="595959"))
        else:
            _ratio_rules(ws, rng, tl)

    legend = ws.cell(first_role + len(roles), label_col,
                     "filled/required, updating as you edit. Red = requirement not met; purple = a PRN "
                     "was requested (PRN NEEDED, from the last run) and isn't entered yet; gray = "
                     "optional role open; yellow = more than required.")
    legend.font = Font(name=fname, sz=max(7, fsize - 2), i=True)
    return cached


def remove_old_blocks(t: Template, ws) -> tuple[int, int]:
    """Remove blocks left by an earlier run (an output or refreshed template fed back in).
    Returns (first deleted row, rows deleted) so same-sheet references can be adjusted."""
    clear_block_formatting(ws, max(t.row_of.values()) + 1, max(t.col_of.values()) + 2)
    # right-hand block: clear in place
    date_row = min(t.row_of.values()) - 1
    first = max(t.col_of.values()) + 2
    if str(ws.cell(date_row, first).value or "").strip().lower() == config.LABEL_NAME.lower():
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
    # Shifts Fulfilled block: delete its rows
    target = config.LABEL_SHIFTS_FULFILLED.lower()
    for row in ws.iter_rows(min_row=max(t.row_of.values()) + 2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip().lower() == target:
                r, c = cell.row, cell.column
                n = 0
                while ws.cell(r + n, c).value not in (None, ""):
                    n += 1
                if ws.cell(r + n, c).value in (None, ""):
                    n += 1  # trailing blank row we added
                ws.delete_rows(r, n)
                return r, n
    return 0, 0


def refresh_file(path, out_path=None) -> Path:
    """Rebuild the live blocks in a template or schedule workbook without solving."""
    import openpyxl
    from .loader import load_template
    path = Path(path)
    out_path = Path(out_path or path)
    t = load_template(path)
    wb = openpyxl.load_workbook(path)
    ws = wb[t.sheet_name] if t.sheet_name else wb.worksheets[0]
    removed = remove_old_blocks(t, ws)
    cached = add_live_blocks(t, wb, ws, dict(t.locked), [], with_prn_row=bool(t.prn_pids), removed=removed)
    wb.save(out_path)
    inject_cached_values(out_path, ws.title, cached)
    return out_path


def add_live_blocks(t: Template, wb, ws, assignment: dict, prn_requests=(), with_prn_row: bool = False,
                    removed: tuple[int, int] = (0, 0)) -> dict[str, object]:
    """Add both blocks to the Schedule sheet `ws` of `wb` (positions from the loaded Template `t`).
    Returns {cell: current result} for inject_cached_values."""
    t._live_assignment = assignment
    first = min(t.row_of.values()) + 1
    pid_col = min(t.col_of.values())
    lay = Layout(ws, staff_rows={p: r + 1 for p, r in t.row_of.items()},
                 date_cols={d: c + 1 for d, c in t.col_of.items()},
                 date_row=first - 2, day_row=first - 1, pid_col=pid_col,
                 name_col=(pid_col - 1) if pid_col - 1 >= 1 else None)
    base_font = copy(ws.cell(first, lay.first_col).font)
    clear_block_formatting(ws, lay.last_row, lay.last_col + 1)
    start, n, moved = insert_fulfilled_rows(t, lay, with_prn_row)

    def ref(sheet, r0, c0):
        same = sheet == t.sheet_name
        row = r0 + 1
        if same and removed[1] and row >= removed[0]:
            row -= removed[1]  # an old block above these inputs was deleted
        if same and row >= moved:
            row += n          # the inserted rows pushed same-sheet inputs down
        return sheet_ref(None if same else sheet, row, c0 + 1)

    if t.dpp_col is not None:
        lay.dpp_ref = {p: ref(t.apt_sheet, r, t.dpp_col) for p, r in t.apt_row.items()}
    lay.req_ref = {(role, wd): ref(t.req_sheet, r, c)
                   for role, r in t.req_row.items() for wd, c in t.req_day_col.items()}
    cached = {}
    cached.update(write_days_block(t, lay, base_font))
    cached.update(write_fulfilled_block(t, lay, start, base_font, prn_requests, with_prn_row))
    wb.calculation.fullCalcOnLoad = True     # Excel recalculates everything when the file opens
    return cached


def _assign(t: Template) -> dict:
    """Current cell contents used for the cached results: the solver's answer if attached, else the
    typed-in cells."""
    return getattr(t, "_live_assignment", None) or t.locked


# --------------------------------------------------------------------------- cached results

def inject_cached_values(path: str | Path, sheet_title: str, values: dict[str, object]) -> None:
    """Write each formula's current result into the saved file, so the numbers show even in viewers
    that don't recalculate, and so the program can read the file back without Excel saving it first.
    Excel recalculates everything on open anyway (fullCalcOnLoad)."""
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    wbx = files["xl/workbook.xml"].decode("utf-8")
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    m = re.search(r'<sheet[^>]*name="' + re.escape(_xml_escape(sheet_title)) + r'"[^>]*r:id="([^"]+)"', wbx) \
        or re.search(r'<sheet[^>]*r:id="([^"]+)"[^>]*name="' + re.escape(_xml_escape(sheet_title)) + '"', wbx)
    if not m:
        return
    rid = m.group(1)
    tm = re.search(r'<Relationship[^>]*Id="' + re.escape(rid) + r'"[^>]*Target="([^"]+)"', rels) \
        or re.search(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="' + re.escape(rid) + '"', rels)
    if not tm:
        return
    target = tm.group(1).lstrip("/")
    name = target if target.startswith("xl/") else "xl/" + target
    xml = files[name].decode("utf-8")

    def repl(mo):
        ref = mo.group(2)
        if ref not in values:
            return mo.group(0)
        v = values[ref]
        attrs = re.sub(r'\st="[^"]*"', "", mo.group(1))
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return f'<c{attrs}><f>{mo.group(3)}</f><v>{v}</v></c>'
        return f'<c{attrs} t="str"><f>{mo.group(3)}</f><v>{_xml_escape(str(v))}</v></c>'

    xml = re.sub(r'<c((?:\s[^>]*?)?\sr="([A-Z]+\d+)"[^>]*)><f>(.*?)</f><v\s*/?>(?:</v>)?</c>', repl, xml)
    files[name] = xml.encode("utf-8")
    tmp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in files.items():
            z.writestr(n, data)
    tmp.replace(path)


def _xml_escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
