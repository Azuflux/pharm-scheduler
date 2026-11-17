"""Read a schedule template (.xlsx or .csv) into a structured Template object.

Blocks are located by their header labels (see config.py), not by fixed row/column
positions, so rows and blocks can move around in the sheet without breaking the loader.
"""
from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

from . import config
from .settings import SPEC as SETTINGS_SPEC, Settings, check as check_settings, parse_value


class LoaderError(Exception):
    """The template's structure can't be read. Message says what and where."""


@dataclass
class Issue:
    level: str      # "error" (blocks solving), "warning" (will affect the result), "info"
    category: str
    message: str


@dataclass
class Staff:
    pid: str                         # key linking schedule rows to aptitude rows
    aptitude: dict[str, int]         # role -> 0..4
    target_days: int | None          # days per pay period; None = no target (e.g. PRN)
    name: str = ""                   # optional display name
    prn: bool = False                # PRN pharmacist: never scheduled individually


@dataclass
class Template:
    source: Path
    dates: list[date]
    staff: dict[str, Staff]                      # pid -> Staff
    schedule_order: list[str]                    # pids in schedule row order
    roles: list[str]                             # role codes, aptitude-matrix order
    requirements: dict[str, list[int]]           # role -> [Mon..Sun] headcount
    locked: dict[tuple[str, date], str]          # (pid, date) -> code, pre-filled cells
    year_note: str = ""
    issues: list[Issue] = field(default_factory=list)
    settings: Settings = field(default_factory=Settings)
    optional_roles: set[str] = field(default_factory=set)
    role_descriptions: dict[str, str] = field(default_factory=dict)
    settings_source: str = "defaults"
    # code -> {"fill": PatternFill, "font": Font} copied from the template's Color swatches
    code_styles: dict[str, dict] = field(default_factory=dict)
    # Non-role codes allowed in the schedule: code -> counts_for_pay, on_duty, covers, meaning
    codes: dict[str, dict] = field(default_factory=lambda: {k: dict(v) for k, v in config.NON_ROLE_CODES.items()})
    # code -> (start, end) in minutes from midnight of the shift's day; end > start (overnight adds 1440)
    shift_times: dict[str, tuple[int, int]] = field(default_factory=dict)
    style_book: object = field(default=None, repr=False)   # openpyxl workbook the styles came from
    # Where things are in the sheet (0-based grid indices), so output can be written in place.
    sheet_name: str | None = None
    row_of: dict[str, int] = field(default_factory=dict)
    col_of: dict[date, int] = field(default_factory=dict)
    grid: list[list] = field(default_factory=list, repr=False)

    # ---- helpers used by the validator and, later, the model ----
    def label(self, pid: str) -> str:
        """'Name (PID)' when a name is known, else the PID."""
        s = self.staff.get(pid)
        return f"{s.name} ({pid})" if s and s.name else pid

    def pay_period(self, d: date) -> int:
        """Index of the pay period containing d (0 = period starting at the anchor)."""
        return (d - self.settings.pay_period_anchor).days // self.settings.pay_period_days

    def period_start(self, idx: int) -> date:
        return self.settings.pay_period_anchor + timedelta(days=idx * self.settings.pay_period_days)

    def periods(self) -> list[dict]:
        """Pay periods touched by the horizon, with the horizon dates in each."""
        out: dict[int, list[date]] = {}
        for d in self.dates:
            out.setdefault(self.pay_period(d), []).append(d)
        result = []
        for idx, ds in sorted(out.items()):
            start = self.period_start(idx)
            end = start + timedelta(days=self.settings.pay_period_days - 1)
            result.append({"index": idx, "start": start, "end": end, "dates": ds,
                           "complete": len(ds) == self.settings.pay_period_days})
        return result

    def required(self, role: str, d: date) -> int:
        return self.requirements.get(role, [0] * 7)[d.weekday()]

    def code_counts_for_pay(self, code: str) -> bool:
        if code in self.codes:
            return self.codes[code]["counts_for_pay"]
        return True  # any role assignment is a worked day

    def code_on_duty(self, code: str) -> bool:
        if code in self.codes:
            return self.codes[code]["on_duty"]
        return True

    def covered_role(self, code: str) -> str | None:
        """The role a code counts toward for coverage: itself if it's a role, else its 'covers'."""
        if code in self.requirements or code in self.roles:
            return code
        return self.codes.get(code, {}).get("covers")

    def is_prn(self, pid: str) -> bool:
        s = self.staff.get(pid)
        return bool(s and s.prn)

    @property
    def prn_pids(self) -> list[str]:
        return [p for p in self.schedule_order if self.is_prn(p)]


# ---------------------------------------------------------------- reading the file

def read_grids(path: Path):
    """Every sheet as a rectangular list of rows, keyed by sheet name (None for a CSV), plus the
    openpyxl workbook (formatting included) for .xlsx, else None.
    Values are raw (str/int/float/datetime/None). grid[r][c] is Excel cell (r+1, c+1)."""
    suffix = path.suffix.lower()
    grids: dict[str | None, list[list]] = {}
    wf = None
    if suffix == ".csv":
        with open(path, encoding="utf-8-sig", newline="") as fh:
            grids[None] = [list(r) for r in csv.reader(fh)]
    elif suffix in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        wf = openpyxl.load_workbook(path, data_only=False)
        uncalculated = []
        for ws in wb.worksheets:
            if ws.title in config.IGNORED_SHEETS:
                continue
            rows = [list(r) for r in ws.iter_rows(min_row=1, min_col=1, values_only=True)]
            fws = wf[ws.title]
            for r, row in enumerate(rows):
                for c, v in enumerate(row):
                    if v is None:
                        f = fws.cell(r + 1, c + 1).value
                        if isinstance(f, str) and f.startswith("="):
                            uncalculated.append(f"{ws.title}!{cell_ref(r, c)}")
            grids[ws.title] = rows
        if uncalculated:
            shown = ", ".join(uncalculated[:5]) + (f" (+{len(uncalculated) - 5} more)" if len(uncalculated) > 5 else "")
            raise LoaderError("This file has formulas without saved results (it was last saved by a "
                              "program other than Excel). Open it in Excel, save, and try again. "
                              f"Cells: {shown}")
    else:
        raise LoaderError(f"Unsupported file type '{suffix}'. Use .xlsx or .csv.")
    for k, rows in grids.items():
        width = max((len(r) for r in rows), default=0)
        grids[k] = [r + [None] * (width - len(r)) for r in rows]
    return grids, wf


def cell_str(v) -> str:
    """Normalize a raw cell value to a stripped string ('' for empty)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    return str(v).strip()


def cell_ref(r: int, c: int) -> str:
    """Excel-style reference (A1) for messages."""
    letters = ""
    c += 1
    while c:
        c, rem = divmod(c - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{r + 1}"


def _where(sheet: str | None, r: int, c: int) -> str:
    return f"{sheet}!{cell_ref(r, c)}" if sheet else cell_ref(r, c)


def find_label(grids: dict, label: str, only_sheet: str | None = None,
               required: bool = True) -> tuple[str | None, int, int] | None:
    """Find a block label on any sheet. Returns (sheet, row, col)."""
    target = label.strip().lower()
    hits = [(name, r, c) for name, grid in grids.items()
            if only_sheet is None or name == only_sheet
            for r, row in enumerate(grid) for c, v in enumerate(row)
            if cell_str(v).lower() == target]
    if not hits:
        if not required:
            return None
        where = f" on sheet '{only_sheet}'" if only_sheet else ""
        raise LoaderError(f"Couldn't find a cell labeled '{label}'{where}.")
    if len(hits) > 1:
        refs = ", ".join(_where(*h) for h in hits)
        raise LoaderError(f"Label '{label}' appears more than once ({refs}); it must be unique.")
    return hits[0]


def to_int(v, where: str, issues: list[Issue], blank=0) -> int | None:
    s = cell_str(v)
    if s == "":
        return blank
    try:
        f = float(s)
    except ValueError:
        raise LoaderError(f"Expected a number at {where}, found '{s}'.")
    if not f.is_integer():
        raise LoaderError(f"Expected a whole number at {where}, found '{s}'.")
    return int(f)


def weekday_of(token, where: str) -> int:
    s = cell_str(token).upper().rstrip(".")
    if s not in config.WEEKDAY_TOKENS:
        raise LoaderError(f"Unrecognized weekday '{cell_str(token)}' at {where}.")
    return config.WEEKDAY_TOKENS[s]


# ---------------------------------------------------------------- dates

_MD = re.compile(r"^(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?$")                        # 2/23, 2/23/27
_DMON = re.compile(r"^(\d{1,2})[-\s]([A-Za-z]{3,9})\.?(?:[-\s,]+(\d{2,4}))?$")     # 23-Feb, 23 Feb 2027
_MOND = re.compile(r"^([A-Za-z]{3,9})\.?[-\s](\d{1,2})(?:[-\s,]+(\d{2,4}))?$")     # Feb-23, Feb 23, 2027
_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def _month(name: str) -> int | None:
    return _MONTHS.get(name[:3].lower())


def _parse_md(s: str) -> tuple[int, int, int | None] | None:
    """(month, day, year or None) from the common header formats, or None."""
    m = _MD.match(s)
    if m:
        mo, dy, yr = int(m.group(1)), int(m.group(2)), m.group(3)
    else:
        m = _DMON.match(s)
        if m and _month(m.group(2)):
            mo, dy, yr = _month(m.group(2)), int(m.group(1)), m.group(3)
        else:
            m = _MOND.match(s)
            if not (m and _month(m.group(1))):
                return None
            mo, dy, yr = _month(m.group(1)), int(m.group(2)), m.group(3)
    if yr is not None:
        yr = int(yr) + (2000 if len(yr) == 2 else 0)
    return mo, dy, yr


def parse_dates(raw: list, dow_row: list, cols: list[int], hdr_row: int, dow_row_idx: int,
                year: int | None) -> tuple[list[date], str]:
    """Turn header cells into real dates. Cells without a year get one inferred from the
    Day row (the year in which every listed weekday matches), closest to today."""
    parsed: list[tuple[int, int, int | None] | date] = []
    for v, c in zip(raw, cols):
        if isinstance(v, datetime):
            parsed.append(v.date())
            continue
        if isinstance(v, date):
            parsed.append(v)
            continue
        s = cell_str(v)
        md = _parse_md(s)
        if md is None:
            try:
                parsed.append(date.fromisoformat(s[:10]))
                continue
            except ValueError:
                raise LoaderError(f"Can't read date '{s}' at {cell_ref(hdr_row, c)}. "
                                  "Use a format like 2/23, 23-Feb, Feb 23, or a real Excel date.")
        parsed.append(md)

    expected_dow = [weekday_of(dow_row[c], cell_ref(dow_row_idx, c)) if cell_str(dow_row[c]) else None
                    for c in cols]

    def build(start_year: int) -> list[date] | None:
        out, y, prev_month = [], start_year, None
        for p in parsed:
            if isinstance(p, date):
                out.append(p); y = p.year; prev_month = p.month
                continue
            mo, dy, yr = p
            if yr is not None:
                y = yr
            elif prev_month is not None and mo < prev_month:
                y += 1  # crossed into January
            try:
                out.append(date(y, mo, dy))
            except ValueError:
                return None
            prev_month = mo
        return out

    if all(isinstance(p, date) or p[2] is not None for p in parsed):
        dates, note = build(1), "Dates include years."  # start year unused when every cell has one
        if dates is None:
            raise LoaderError("One of the header dates isn't a real calendar date.")
    elif year is not None:
        dates, note = build(year), f"Year {year} supplied on the command line."
        if dates is None:
            raise LoaderError(f"Dates aren't valid in year {year}.")
    else:
        today = date.today()
        candidates = []
        for y in range(today.year - 3, today.year + 4):
            ds = build(y)
            if ds and all(e is None or d.weekday() == e for d, e in zip(ds, expected_dow)):
                candidates.append(ds)
        if not candidates:
            raise LoaderError("Dates have no year, and no year between "
                              f"{today.year - 3} and {today.year + 3} matches the Day row. "
                              "Add years to the dates or pass --year.")
        dates = min(candidates, key=lambda ds: abs((ds[0] - today).days))
        note = (f"Dates have no year; inferred {dates[0].year} because it's the nearest year "
                f"where the dates fall on the weekdays in the Day row. Pass --year to override.")

    # Whatever the source, check weekdays and sequence.
    for d, e, c in zip(dates, expected_dow, cols):
        if e is not None and d.weekday() != e:
            raise LoaderError(f"{d:%m/%d/%Y} at {cell_ref(hdr_row, c)} is a {d:%A}, "
                              f"but the Day row says otherwise.")
    for a, b, c in zip(dates, dates[1:], cols[1:]):
        if (b - a).days != 1:
            raise LoaderError(f"Dates must be consecutive; {a:%m/%d} is followed by {b:%m/%d} "
                              f"at {cell_ref(hdr_row, c)}.")
    return dates, note


# ---------------------------------------------------------------- blocks

def _block_rows(grid, start_row: int, key_col: int, name_col: int | None = None) -> list[int]:
    """Rows from start_row down while the key column (or the name column) is non-empty.
    A row with a name but no key is an error, not the end of the block."""
    rows = []
    r = start_row
    while r < len(grid):
        key = cell_str(grid[r][key_col])
        name = cell_str(grid[r][name_col]) if name_col is not None else ""
        if not key and not name:
            break
        if not key and config.LABEL_PRN_SECTION.lower() in (name.lower(), ""):
            r += 1
            continue  # "PRN STAFF" separator row
        if key.lower() == config.LABEL_PRN_SECTION.lower() and not name:
            r += 1
            continue
        if not key:
            raise LoaderError(f"Row {r + 1} has the name '{name}' but no {config.LABEL_PID} "
                              f"(expected at {cell_ref(r, key_col)}).")
        rows.append(r)
        r += 1
    return rows


def _gap_after_rows(grid, last_row: int, col: int, looks_like_data, what: str, look: int = 3) -> None:
    """Raise if a data-looking row appears within `look` rows after a block ends at last_row.
    Catches a blank row accidentally inserted inside a block (which would cut it short)."""
    labels = {x.lower() for x in (config.LABEL_APTITUDE, config.LABEL_REQUIREMENTS,
                                  config.LABEL_DAYS_PER_PERIOD, config.LABEL_SHIFTS_FULFILLED,
                                  config.LABEL_SCHEDULE_DATES, config.LABEL_PID, config.LABEL_NAME,
                                  config.LABEL_OTHER_CODES, config.LABEL_SETTINGS,
                                  config.LABEL_PRN_SECTION)}
    for r in range(last_row + 2, min(last_row + 2 + look, len(grid))):
        if any(cell_str(v).lower() in labels for v in grid[r]):
            return  # reached the next block
        v = cell_str(grid[r][col])
        if v and looks_like_data(r, v):
            raise LoaderError(f"The {what} stops at row {last_row + 1}, but more entries continue at "
                              f"{cell_ref(r, col)} after a blank row. Remove the blank row so the "
                              f"block is continuous.")


def _find_in_row(grid, row: int, text: str, before_col: int | None = None) -> int | None:
    """Column of the cell in `row` equal to text (case-insensitive), optionally left of before_col."""
    end = before_col if before_col is not None else len(grid[row])
    hits = [c for c in range(end) if cell_str(grid[row][c]).lower() == text.lower()]
    return hits[0] if hits else None


def _header_cols(grid, row: int, start_col: int) -> list[int]:
    """Columns from start_col rightward while the header cell is non-empty."""
    cols = []
    c = start_col
    while c < len(grid[row]) and cell_str(grid[row][c]):
        cols.append(c)
        c += 1
    return cols




def cell_style(cell, any_font: bool = False) -> dict | None:
    """The fill and font of a cell. None if it has no color formatting, unless any_font is set
    (swatches: their font name and size count too). 'colored' says whether a fill or text color
    is present."""
    from copy import copy
    fill, font = cell.fill, cell.font
    has_fill = fill is not None and fill.fill_type not in (None, "none")
    has_font_color = font is not None and font.color is not None and \
        not (font.color.type == "theme" and font.color.theme == 1 and not font.color.tint) and \
        not (font.color.type == "rgb" and font.color.rgb in ("FF000000", "00000000"))
    if not has_fill and not has_font_color and not (font and (font.b or font.i)) and not any_font:
        return None
    f = copy(font)
    if f.name is None:
        f.name = "Calibri"
    if f.sz is None:
        f.sz = 11
    return {"fill": copy(fill), "font": f, "colored": bool(has_fill or has_font_color)}


def style_signature(st: dict | None) -> str:
    """Comparable text form of a style dict ('' for no style)."""
    if not st:
        return ""
    def col(c):
        return None if c is None else (c.type, c.value, round(c.tint or 0, 3))
    f, n = st["fill"], st["font"]
    return repr((f.fill_type if f.fill_type not in (None, "none") else None,
                 col(f.fgColor) if f.fill_type not in (None, "none") else None,
                 col(n.color), bool(n.b), bool(n.i), n.name, float(n.sz or 11)))


def legacy_style_signature(st: dict | None) -> str:
    """Signature format used by templates made before font name/size were tracked."""
    if not st:
        return ""
    def col(c):
        return None if c is None else (c.type, c.value, round(c.tint or 0, 3))
    f, n = st["fill"], st["font"]
    return repr((f.fill_type, col(f.fgColor), col(n.color), bool(n.b), bool(n.i)))


def read_applied_styles(wb) -> dict[str, str]:
    """code -> signature the program last applied (from the hidden sheet), if present."""
    if config.APPLIED_STYLES_SHEET not in wb.sheetnames:
        return {}
    ws = wb[config.APPLIED_STYLES_SHEET]
    return {str(r[0]): str(r[1] or "") for r in ws.iter_rows(min_row=2, values_only=True) if r[0]}


def write_applied_styles(wb, styles: dict[str, dict]) -> None:
    if config.APPLIED_STYLES_SHEET in wb.sheetnames:
        del wb[config.APPLIED_STYLES_SHEET]
    ws = wb.create_sheet(config.APPLIED_STYLES_SHEET)
    ws.append(["Code", "Applied style (managed by the program; do not edit)"])
    for code, st in styles.items():
        ws.append([code, style_signature(st)])
    ws.sheet_state = "hidden"


def parse_time(raw, where: str) -> int | None:
    """Minutes after midnight from an Excel time, 0630 / 630 / '6:30' / '6:30 PM'; None if blank."""
    from datetime import time as dtime
    if raw is None or (isinstance(raw, str) and not raw.strip()):
        return None
    if isinstance(raw, datetime):
        raw = raw.time()
    if isinstance(raw, dtime):
        return raw.hour * 60 + raw.minute
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        if 0 <= raw < 1:                       # Excel time as a fraction of a day
            return round(raw * 1440) % 1440
        if float(raw).is_integer() and 0 <= raw <= 2400:
            hh, mm = divmod(int(raw), 100)
            if hh <= 24 and mm < 60:
                return (hh * 60 + mm) % 1440
        raise LoaderError(f"Can't read the time {raw!r} at {where}. Use e.g. 0630 or 6:30.")
    s = str(raw).strip().upper().replace(".", "")
    m = re.fullmatch(r"(\d{1,2}):?(\d{2})\s*(AM|PM)?", s)
    if not m:
        raise LoaderError(f"Can't read the time '{raw}' at {where}. Use e.g. 0630 or 6:30.")
    hh, mm, ap = int(m.group(1)), int(m.group(2)), m.group(3)
    if ap == "PM" and hh < 12:
        hh += 12
    if ap == "AM" and hh == 12:
        hh = 0
    if hh > 24 or mm > 59:
        raise LoaderError(f"Can't read the time '{raw}' at {where}.")
    return (hh * 60 + mm) % 1440


def _shift(grid, r, sc, ec, sheet, code, out: dict) -> None:
    """Read a Start/End pair on row r into out[code]."""
    if sc is None or ec is None:
        return
    a = parse_time(grid[r][sc], _where(sheet, r, sc))
    b = parse_time(grid[r][ec], _where(sheet, r, ec))
    if a is None and b is None:
        return
    if a is None or b is None:
        raise LoaderError(f"'{code}' has a {'Start' if a is not None else 'End'} time but no "
                          f"{'End' if a is not None else 'Start'} ({_where(sheet, r, sc)}).")
    out[code] = (a, b if b > a else b + 1440)


def _yes_no(raw, where: str) -> bool:
    try:
        return bool(parse_value("bool", raw, where))
    except ValueError as e:
        raise LoaderError(str(e))


def load_template(path: str | Path, sheet: str | None = None, year: int | None = None) -> Template:
    """Load a template. Blocks may sit on one sheet or on separate sheets; each is found by
    its label. `sheet` limits where the schedule ('Date' label) is looked for."""
    path = Path(path)
    grids, book = read_grids(path)
    issues: list[Issue] = []
    code_styles: dict[str, dict] = {}

    # ================= aptitude matrix (+ days per pay period)
    a_sheet, ar, label_col = find_label(grids, config.LABEL_APTITUDE)
    g = grids[a_sheet]
    hdr = ar + 1
    if hdr >= len(g):
        raise LoaderError(f"Nothing under '{config.LABEL_APTITUDE}' ({_where(a_sheet, ar, label_col)}).")
    ac = _find_in_row(g, hdr, config.LABEL_PID)
    if ac is None:
        ac = label_col  # no "PID" header: IDs sit under the block label
    name_col_apt = _find_in_row(g, hdr, config.LABEL_NAME, before_col=ac)
    hdr_cols = _header_cols(g, hdr, ac + 1)
    dpp_col = next((c for c in hdr_cols
                    if cell_str(g[hdr][c]).lower() == config.LABEL_DAYS_PER_PERIOD.lower()), None)
    prn_col = next((c for c in hdr_cols if cell_str(g[hdr][c]).lower() == config.LABEL_PRN.lower()), None)
    role_cols = [c for c in hdr_cols if c not in (dpp_col, prn_col)]
    if dpp_col is None:
        # Older layout: "Days Per Pay Period" is its own label on the label row.
        d_sheet, dr, dpp_col = find_label(grids, config.LABEL_DAYS_PER_PERIOD)
        if d_sheet != a_sheet or dr != ar:
            issues.append(Issue("warning", "layout",
                                f"'{config.LABEL_DAYS_PER_PERIOD}' isn't in the aptitude header or on the "
                                f"'{config.LABEL_APTITUDE}' row; reading its column on the aptitude rows anyway."))
    roles = [cell_str(g[hdr][c]) for c in role_cols]
    if not roles:
        raise LoaderError(f"No role codes found in the aptitude header ({_where(a_sheet, hdr, ac + 1)}).")
    after = hdr_cols[-1] + 2
    if after < len(g[hdr]) and cell_str(g[hdr][after]) and \
            cell_str(g[hdr][after]).upper().rstrip(".") not in config.WEEKDAY_TOKENS:
        raise LoaderError(f"The aptitude header stops at {_where(a_sheet, hdr, hdr_cols[-1])}, but another "
                          f"code '{cell_str(g[hdr][after])}' follows at {_where(a_sheet, hdr, after)} after a "
                          f"blank column. Remove the blank column.")
    dupes = {r for r in roles if [x.upper() for x in roles].count(r.upper()) > 1}
    if dupes:
        raise LoaderError(f"Duplicate role codes in the aptitude header: {', '.join(sorted(dupes))}.")
    canon = {r.upper(): r for r in roles}

    staff: dict[str, Staff] = {}
    apt_rows = _block_rows(g, hdr + 1, ac, name_col_apt)
    if apt_rows:
        _gap_after_rows(g, apt_rows[-1], ac,
                        lambda r, v: all(cell_str(g[r][c]).isdigit() for c in role_cols[:3]),
                        "aptitude matrix")
    for r in apt_rows:
        pid = cell_str(g[r][ac])
        if pid in staff:
            raise LoaderError(f"PID '{pid}' appears twice in the aptitude matrix ({_where(a_sheet, r, ac)}).")
        apt = {}
        for role, c in zip(roles, role_cols):
            if cell_str(g[r][c]) == "":
                issues.append(Issue("warning", "aptitude",
                                    f"Blank aptitude at {_where(a_sheet, r, c)} ({pid}, {role}); treated as 0."))
            val = to_int(g[r][c], _where(a_sheet, r, c), issues)
            if not 0 <= val <= 4:
                raise LoaderError(f"Aptitude must be 0-4; found {val} at {_where(a_sheet, r, c)}.")
            apt[role] = val
        target = to_int(g[r][dpp_col], _where(a_sheet, r, dpp_col), issues, blank=None)
        name = cell_str(g[r][name_col_apt]) if name_col_apt is not None else ""
        prn = prn_col is not None and _yes_no(g[r][prn_col], _where(a_sheet, r, prn_col))
        staff[pid] = Staff(pid, apt, None if prn else target, name, prn)

    # ================= requirement matrix
    r_sheet, rr, rc = find_label(grids, config.LABEL_REQUIREMENTS)
    g = grids[r_sheet]
    rhdr = rr + 1
    role_col = _find_in_row(g, rhdr, config.LABEL_ROLE)
    if role_col is None or role_col < rc:
        role_col = rc
    opt_col = _find_in_row(g, rhdr, config.LABEL_OPTIONAL)
    color_col = _find_in_row(g, rhdr, config.LABEL_COLOR)
    if color_col is not None and color_col < role_col:
        color_col = None
    start_col = _find_in_row(g, rhdr, config.LABEL_START)
    end_col = _find_in_row(g, rhdr, config.LABEL_END)
    shift_times: dict[str, tuple[int, int]] = {}
    desc_col = _find_in_row(g, rhdr, config.LABEL_DESCRIPTION)
    if opt_col is not None and opt_col < role_col:
        opt_col = None
    if desc_col is not None and desc_col < role_col:
        desc_col = None
    day_cols: dict[int, int] = {}
    for c in range(role_col + 1, len(g[rhdr])):
        s = cell_str(g[rhdr][c])
        if s and s.upper().rstrip(".") in config.WEEKDAY_TOKENS:
            wd = config.WEEKDAY_TOKENS[s.upper().rstrip(".")]
            if wd in day_cols:
                break  # a second week of headers belongs to something else
            day_cols[wd] = c
    missing = [d for d in range(7) if d not in day_cols]
    if missing:
        names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        raise LoaderError("Requirement matrix header is missing: " + ", ".join(names[d] for d in missing))

    requirements: dict[str, list[int]] = {}
    optional_roles: set[str] = set()
    descriptions: dict[str, str] = {}
    req_rows = _block_rows(g, rhdr + 1, role_col)
    if req_rows:
        _gap_after_rows(g, req_rows[-1], role_col, lambda r, v: v.upper() in canon, "requirement matrix")
    for r in req_rows:
        raw_role = cell_str(g[r][role_col])
        role = canon.get(raw_role.upper(), raw_role)
        if role in requirements:
            raise LoaderError(f"Role '{role}' appears twice in the requirement matrix ({_where(r_sheet, r, role_col)}).")
        counts = []
        for wd in range(7):
            v = to_int(g[r][day_cols[wd]], _where(r_sheet, r, day_cols[wd]), issues)
            if v < 0:
                raise LoaderError(f"Negative requirement at {_where(r_sheet, r, day_cols[wd])}.")
            counts.append(v)
        requirements[role] = counts
        if opt_col is not None and _yes_no(g[r][opt_col], _where(r_sheet, r, opt_col)):
            optional_roles.add(role)
        if desc_col is not None and cell_str(g[r][desc_col]):
            descriptions[role] = cell_str(g[r][desc_col])
        _shift(g, r, start_col, end_col, r_sheet, role, shift_times)
        if color_col is not None and book is not None:
            st = cell_style(book[r_sheet].cell(r + 1, color_col + 1), any_font=True)
            if st:
                code_styles[role] = st
    if opt_col is None:
        optional_roles = {r for r in config.OPTIONAL_ROLES if r in requirements}
        if optional_roles:
            issues.append(Issue("info", "optional role",
                                f"No '{config.LABEL_OPTIONAL}' column in the requirement matrix; using "
                                f"config.py: {', '.join(sorted(optional_roles))} optional."))

    # ================= other codes (optional block): every non-role code, its color and meaning
    codes: dict[str, dict] = {k: dict(v) for k, v in config.NON_ROLE_CODES.items()}
    found = find_label(grids, config.LABEL_OTHER_CODES, required=False)
    if found:
        o_sheet, o_r, o_c = found
        g = grids[o_sheet]
        h = o_r + 1
        hdr_of = lambda lbl: _find_in_row(g, h, lbl) if h < len(g) else None
        code_c, col_c = hdr_of(config.LABEL_CODE), hdr_of(config.LABEL_COLOR)
        pay_c, work_c = hdr_of(config.LABEL_COUNTS), hdr_of(config.LABEL_WORKDAY)
        cov_c, mean_c = hdr_of(config.LABEL_COVERS), hdr_of(config.LABEL_MEANING)
        os_c, oe_c = hdr_of(config.LABEL_START), hdr_of(config.LABEL_END)
        if code_c is None:
            raise LoaderError(f"The '{config.LABEL_OTHER_CODES}' table needs a '{config.LABEL_CODE}' header "
                              f"in the row under it ({_where(o_sheet, h, o_c)}).")
        if pay_c is not None:
            codes = {}  # the table defines the full list of non-role codes
        seen_codes: set[str] = set()
        for r in range(h + 1, len(g)):
            raw = cell_str(g[r][code_c])
            if not raw:
                break
            if raw.upper() in canon:
                raise LoaderError(f"'{raw}' in {config.LABEL_OTHER_CODES} ({_where(o_sheet, r, code_c)}) is "
                                  f"also a role code. List roles only on the Aptitude/Requirements sheets.")
            base = codes.get(raw, config.NON_ROLE_CODES.get(raw.upper(),
                                                            {"counts_for_pay": True, "on_duty": True,
                                                             "covers": None, "meaning": ""}))
            entry = dict(base)
            if pay_c is not None:
                entry["counts_for_pay"] = _yes_no(g[r][pay_c], _where(o_sheet, r, pay_c))
            if work_c is not None:
                entry["on_duty"] = _yes_no(g[r][work_c], _where(o_sheet, r, work_c))
            if cov_c is not None and cell_str(g[r][cov_c]):
                cov = canon.get(cell_str(g[r][cov_c]).upper())
                if cov is None:
                    raise LoaderError(f"'{cell_str(g[r][cov_c])}' at {_where(o_sheet, r, cov_c)} isn't a role code.")
                entry["covers"] = cov
            if mean_c is not None:
                entry["meaning"] = cell_str(g[r][mean_c])
            if raw.upper() in seen_codes:
                raise LoaderError(f"Code '{raw}' is listed twice in {config.LABEL_OTHER_CODES}.")
            seen_codes.add(raw.upper())
            _shift(g, r, os_c, oe_c, o_sheet, raw, shift_times)
            codes.pop(next((k for k in codes if k.upper() == raw.upper()), None), None)
            codes[raw] = entry
            if col_c is not None and book is not None:
                st = cell_style(book[o_sheet].cell(r + 1, col_c + 1), any_font=True)
                if st:
                    code_styles[raw] = st

    # ================= settings (optional block)
    settings = Settings()
    settings_source = "defaults (no Settings sheet)"
    found = find_label(grids, config.LABEL_SETTINGS, required=False)
    if found:
        s_sheet, s_r, s_c = found
        g = grids[s_sheet]
        h = s_r + 1
        key_c = _find_in_row(g, h, config.LABEL_SETTING_NAME) if h < len(g) else None
        val_c = _find_in_row(g, h, config.LABEL_SETTING_VALUE) if h < len(g) else None
        if key_c is None or val_c is None:
            raise LoaderError(f"The Settings table needs '{config.LABEL_SETTING_NAME}' and "
                              f"'{config.LABEL_SETTING_VALUE}' headers in the row under "
                              f"'{config.LABEL_SETTINGS}' ({_where(s_sheet, h, s_c)}).")
        spec = {lbl.lower(): (lbl, attr, kind) for lbl, attr, kind, _ in SETTINGS_SPEC}
        seen = set()
        for r in range(h + 1, len(g)):
            key = cell_str(g[r][key_c])
            if not key:
                break
            if key.lower() not in spec:
                issues.append(Issue("warning", "settings",
                                    f"Unknown setting '{key}' at {_where(s_sheet, r, key_c)} ignored. "
                                    "Setting names must match exactly."))
                continue
            lbl, attr, kind = spec[key.lower()]
            try:
                setattr(settings, attr, parse_value(kind, g[r][val_c], _where(s_sheet, r, val_c)))
            except ValueError as e:
                raise LoaderError(f"Setting '{lbl}': {e}")
            seen.add(key.lower())
        for lbl, attr, kind, _ in SETTINGS_SPEC:
            if lbl.lower() not in seen:
                issues.append(Issue("info", "settings",
                                    f"'{lbl}' not on the Settings sheet; using the default "
                                    f"({getattr(settings, attr)})."))
        settings_source = f"sheet '{s_sheet}'" if s_sheet else "template"
    for msg in check_settings(settings):
        issues.append(Issue("warning", "settings", msg))
    if settings.project_code not in codes:
        if settings.project_code.upper() in canon:
            raise LoaderError(f"Project day code '{settings.project_code}' is a role code; choose another.")
        codes[settings.project_code] = {"counts_for_pay": True, "on_duty": True, "covers": None,
                                        "meaning": "Project day"}
        if found:
            issues.append(Issue("warning", "codes",
                                f"Project day code '{settings.project_code}' isn't in {config.LABEL_OTHER_CODES}; "
                                f"added as a paid workday."))

    # ================= schedule grid
    s_sheet, sr, sc = find_label(grids, config.LABEL_SCHEDULE_DATES, only_sheet=sheet)
    g = grids[s_sheet]
    day_row = sr + 1
    if day_row >= len(g) or cell_str(g[day_row][sc]).lower() != config.LABEL_SCHEDULE_DAYS.lower():
        raise LoaderError(f"Expected '{config.LABEL_SCHEDULE_DAYS}' at {_where(s_sheet, day_row, sc)}, "
                          f"directly under '{config.LABEL_SCHEDULE_DATES}'.")
    date_cols = _header_cols(g, sr, sc + 1)
    # A "Name" header after the dates starts the right-hand name / days-per-pay-period columns
    # that the output adds; the dates end there.
    right_name = next((i for i, c in enumerate(date_cols)
                       if cell_str(g[sr][c]).lower() == config.LABEL_NAME.lower()), None)
    if right_name is not None:
        date_cols = date_cols[:right_name]
    if not date_cols:
        raise LoaderError(f"No dates found to the right of {_where(s_sheet, sr, sc)}.")
    nxt = date_cols[-1] + 2
    if right_name is None and nxt < len(g[sr]) and cell_str(g[sr][nxt]) and cell_str(g[day_row][nxt]):
        raise LoaderError(f"Dates stop at {_where(s_sheet, sr, date_cols[-1])}, but more dates continue at "
                          f"{_where(s_sheet, sr, nxt)} after a blank column. Remove the blank column.")
    dates, year_note = parse_dates([g[sr][c] for c in date_cols], g[day_row], date_cols,
                                   sr, day_row, year)

    code_map = dict(canon)
    code_map.update({k.upper(): k for k in codes})

    locked: dict[tuple[str, date], str] = {}
    order: list[str] = []
    row_of: dict[str, int] = {}
    unknown: dict[str, list[str]] = {}
    whitespace: list[str] = []
    name_col_sched = _find_in_row(g, sr, config.LABEL_NAME, before_col=sc)
    if name_col_sched is None:
        name_col_sched = _find_in_row(g, day_row, config.LABEL_NAME, before_col=sc)
    sched_names: dict[str, str] = {}
    sched_rows = _block_rows(g, day_row + 1, sc, name_col_sched)
    if sched_rows:
        _gap_after_rows(g, sched_rows[-1], sc, lambda r, v: v in staff, "schedule")
    for r in sched_rows:
        pid = cell_str(g[r][sc])
        if pid in order:
            raise LoaderError(f"'{pid}' appears twice in the schedule ({_where(s_sheet, r, sc)}).")
        order.append(pid)
        row_of[pid] = r
        if name_col_sched is not None and cell_str(g[r][name_col_sched]):
            sched_names[pid] = cell_str(g[r][name_col_sched])
        for d, c in zip(dates, date_cols):
            raw = g[r][c]
            s = cell_str(raw)
            if s == "":
                if raw not in (None, ""):
                    whitespace.append(_where(s_sheet, r, c))
                continue
            code = code_map.get(s.upper())
            if code is None:
                unknown.setdefault(s, []).append(_where(s_sheet, r, c))
                continue
            locked[(pid, d)] = code

    for s, refs in unknown.items():
        shown = ", ".join(refs[:6]) + (f" (+{len(refs) - 6} more)" if len(refs) > 6 else "")
        issues.append(Issue("error", "unknown code",
                            f"'{s}' isn't a role or a code in {config.LABEL_OTHER_CODES}: {shown}. "
                            f"Add it to the {config.LABEL_OTHER_CODES} table (Requirements sheet) "
                            f"or correct the cell."))
    if whitespace:
        issues.append(Issue("info", "cleanup",
                            f"{len(whitespace)} cell(s) contain only spaces, treated as blank: "
                            + ", ".join(whitespace)))
    for pid, nm in sched_names.items():
        st = staff.get(pid)
        if st is None:
            continue
        if not st.name:
            st.name = nm
        elif st.name.strip().lower() != nm.strip().lower():
            issues.append(Issue("warning", "names",
                                f"PID {pid} is '{nm}' on the schedule but '{st.name}' in the aptitude "
                                f"matrix. Using '{st.name}'. (PID is what links the rows.)"))

    return Template(source=path, dates=dates, staff=staff, schedule_order=order, roles=roles,
                    requirements=requirements, locked=locked, year_note=year_note, issues=issues,
                    settings=settings, optional_roles=optional_roles, role_descriptions=descriptions,
                    settings_source=settings_source, code_styles=code_styles, style_book=book, codes=codes,
                    shift_times=shift_times,
                    sheet_name=s_sheet, row_of=row_of,
                    col_of=dict(zip(dates, date_cols)), grid=g)
