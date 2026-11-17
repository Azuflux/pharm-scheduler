"""Department policy settings, editable on the template's Settings sheet.

Each entry below is one row on that sheet. Anything missing from the sheet (or a template
without a Settings sheet, like a CSV) falls back to the default here. Solver scoring
weights are not here on purpose: they live in config.py as advanced tuning.
"""
from __future__ import annotations

from dataclasses import dataclass, fields
from datetime import date, datetime

from . import config


@dataclass
class Settings:
    pay_period_anchor: date = config.PAY_PERIOD_ANCHOR
    pay_period_days: int = config.PAY_PERIOD_DAYS
    max_consecutive_hard: int = config.MAX_CONSECUTIVE_HARD
    max_consecutive_soft: int = config.MAX_CONSECUTIVE_SOFT
    full_time_target: int = config.FULL_TIME_TARGET
    weekend_comp: bool = True
    project_weekends: bool = not config.PROJECT_WEEKDAYS_ONLY
    project_soft_cap: int | None = config.PROJECT_DAYS_SOFT_CAP
    time_limit: int = config.SOLVER_TIME_LIMIT_SECONDS
    min_rest_hours: int = 8
    preferred_rest_hours: int = 14
    balance_aptitude: int | None = 4
    balance_mode: str = "By days available"
    project_code: str = config.PROJECT_CODE
    prn_for_optional: bool = True


# (row label on the sheet, attribute, kind, help text)
SPEC = [
    ("Pay period start date", "pay_period_anchor", "date",
     "Any date a pay period begins (a Monday). Every other pay period is counted forward and "
     "backward from it."),
    ("Pay period length (days)", "pay_period_days", "int",
     "Length of a pay period. Days-per-pay-period targets on the Aptitude sheet apply to this span."),
    ("Max consecutive workdays", "max_consecutive_hard", "int",
     "Never schedule more than this many days in a row (filled-in cells can still exceed it)."),
    ("Preferred max consecutive workdays", "max_consecutive_soft", "int",
     "Try not to go past this many days in a row. Must be no more than the max above."),
    ("Minimum hours between shifts", "min_rest_hours", "int",
     "From the end of one shift to the start of the next day's shift. A gap of this many hours or "
     "fewer is never scheduled (cells you fill in yourself are kept)."),
    ("Preferred hours between shifts", "preferred_rest_hours", "int",
     "Gaps shorter than this are allowed but discouraged, more strongly the closer they get to the "
     "minimum. 14 leaves normal day-to-day patterns alone (about 15 hours between 0700-1530 shifts); "
     "raise it to push for longer rests everywhere."),
    ("Full-time days per pay period", "full_time_target", "int",
     "Anyone with at least this many days per pay period counts as full-time for comp days."),
    ("Weekend comp days for full-timers", "weekend_comp", "bool",
     "Yes: a full-timer who works Sat+Sun gets the next Monday off when possible, and ideally "
     "Tuesday. No: ignore."),
    ("Project days on weekends", "project_weekends", "bool",
     "Yes: project days (PRJ) may fall on Sat/Sun. No: weekdays only."),
    ("Project days per pay period before spreading", "project_soft_cap", "int_or_blank",
     "Above this many project days in one pay period, the program shares them out with others, "
     "even if that means a 2 covers a role instead of a 3. Leave blank to never spread."),
    ("Solver time limit (seconds)", "time_limit", "int",
     "How long the program searches. 60 is usually enough; use 180+ for a final schedule."),
    ("Balance shifts among aptitude", "balance_aptitude", "int_or_blank",
     "For each role, pharmacists with this aptitude score share that role's shifts as evenly as the "
     "other rules allow, over the whole schedule (e.g. both transplant pharmacists get about the "
     "same number of TX shifts). Leave blank to turn off."),
    ("Balance by", "balance_mode", "choice:By days available|Equal",
     "By days available: shares in proportion to each person's days per pay period minus PTO (equal "
     "when two people work the same days). Equal: the same number regardless of hours or PTO."),
    ("Project day code", "project_code", "text",
     "The code the program writes for a project day. It must be listed in the Other Codes table."),
    ("Request PRNs for optional roles", "prn_for_optional", "bool",
     "Yes: if an optional role (e.g. CP) can't be covered by regular staff and a PRN is able to "
     "work it, the program requests a PRN. No: optional roles are simply left open."),
]

_YES = {"yes", "y", "true", "1", "x", "on"}
_NO = {"no", "n", "false", "0", "", "off"}


def parse_value(kind: str, raw, where: str) -> object:
    """Convert a cell value; raise ValueError with a readable message."""
    s = "" if raw is None else str(raw).strip()
    if kind == "date":
        if isinstance(raw, datetime):
            return raw.date()
        if isinstance(raw, date):
            return raw
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
            try:
                return datetime.strptime(s[:10], fmt).date()
            except ValueError:
                pass
        raise ValueError(f"{where}: expected a date, found '{s}'.")
    if kind in ("int", "int_or_blank"):
        if s == "" and kind == "int_or_blank":
            return None
        try:
            f = float(s)
            if f.is_integer() and f >= 0:
                return int(f)
        except ValueError:
            pass
        raise ValueError(f"{where}: expected a whole number, found '{s}'.")
    if kind.startswith("choice:"):
        options = kind.split(":", 1)[1].split("|")
        for o in options:
            if s.lower() == o.lower():
                return o
        raise ValueError(f"{where}: expected one of {', '.join(options)}, found '{s}'.")
    if kind == "text":
        if not s:
            raise ValueError(f"{where}: expected text, found nothing.")
        return s
    if kind == "bool":
        if isinstance(raw, bool):
            return raw
        if s.lower() in _YES:
            return True
        if s.lower() in _NO:
            return False
        raise ValueError(f"{where}: expected Yes or No, found '{s}'.")
    raise ValueError(f"unknown kind {kind}")


def display_value(kind: str, v) -> object:
    """Value as it should appear on the Settings sheet."""
    if kind == "bool":
        return "Yes" if v else "No"
    if v is None:
        return ""
    return v


def check(s: Settings) -> list[str]:
    """Cross-field problems."""
    out = []
    if s.max_consecutive_soft > s.max_consecutive_hard:
        out.append("Preferred max consecutive workdays is larger than the max; using the max for both.")
        s.max_consecutive_soft = s.max_consecutive_hard
    if s.pay_period_days < 1:
        out.append("Pay period length must be at least 1 day; using 14.")
        s.pay_period_days = 14
    if s.preferred_rest_hours < s.min_rest_hours:
        out.append("Preferred hours between shifts is below the minimum; using the minimum for both.")
        s.preferred_rest_hours = s.min_rest_hours
    if s.time_limit < 1:
        s.time_limit = config.SOLVER_TIME_LIMIT_SECONDS
    return out


assert {a for _, a, _, _ in SPEC} == {f.name for f in fields(Settings)}
