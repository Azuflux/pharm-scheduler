"""Department-level settings. Edit these rather than the loader/validator code."""
from datetime import date

# Pay periods are 14 days, Monday through the second Sunday.
# Any known period start works as the anchor; earlier/later periods are computed from it.
PAY_PERIOD_ANCHOR = date(2026, 11, 16)
PAY_PERIOD_DAYS = 14

# Codes that may appear in the schedule grid but are not roles from the aptitude matrix.
#   counts_for_pay: counts toward the pharmacist's days-per-pay-period target
#   on_duty:        counts as a day physically worked (for consecutive-day limits)
#   covers:         role this code also counts as covering (e.g. an 8-hour ON8 covers ON)
# These are defaults for templates without an "Other Codes" table (e.g. older CSVs).
# Newer templates define every non-role code in that table instead.
NON_ROLE_CODES = {
    "PTO": {"counts_for_pay": True, "on_duty": False, "covers": None, "meaning": "Paid time off"},
    "OFF": {"counts_for_pay": False, "on_duty": False, "covers": None, "meaning": "Day off"},
    "PRJ": {"counts_for_pay": True, "on_duty": True, "covers": None, "meaning": "Project day"},
}
PROJECT_CODE = "PRJ"   # default; the Settings sheet's "Project day code" overrides it

# Project days are only given on weekdays.
PROJECT_WEEKDAYS_ONLY = True

# Roles that are nice to fill but not required. Their number in the requirement
# matrix is the most people the role can take that day. Only used for templates whose
# requirement matrix has no "Optional" column (e.g. older CSVs).
OPTIONAL_ROLES = {"CP"}

# Consecutive on-duty days: never more than HARD; prefer no more than SOFT.
MAX_CONSECUTIVE_HARD = 7
MAX_CONSECUTIVE_SOFT = 6

# A pharmacist with at least this many days per pay period is full-time
# (gets the weekend comp-day preference below).
FULL_TIME_TARGET = 10

# ---- Objective weights (the solver maximizes reward minus penalties) ----
# Reward for each solver-made assignment, by aptitude score. A 1 is negative: it's a last
# resort, used only when no combination of 2s-4s covers the slot. (It also means optional
# roles are never filled with a 1.)
APTITUDE_WEIGHT = {1: -60, 2: 40, 3: 100, 4: 160}
# Optional roles use their own (all-positive) aptitude weights, since any eligible
# pharmacist covering one beats a project day. Higher aptitude still preferred.
OPTIONAL_APTITUDE_WEIGHT = {1: 10, 2: 20, 3: 30, 4: 40}
# Extra reward for filling an optional role (on top of aptitude), so a pharmacist
# who would otherwise get a project day covers it instead.
OPTIONAL_FILL_BONUS = 20
# Penalty per project day. Small: project days are fine, roles are better.
PROJECT_DAY_PENALTY = 1
# Spread project days: penalty for each project day a person gets beyond the cap in one
# pay period. Compare with the aptitude gaps above: at 80, the solver will put a 2 in
# place of a 3 (a 60-point loss) to take a project day off someone over the cap, but
# won't put a 1 in place of a 2 (a 100-point loss). Set to 0 to ignore spread.
PROJECT_DAYS_SOFT_CAP = 4
PROJECT_SPREAD_PENALTY = 80
# Penalty per PRN shift requested. Large, so a PRN is requested only when no regular pharmacist
# can cover the slot within the rules. (Scheduling a regular past their pay-period target is
# never an option: that's a hard rule.)
PRN_REQUEST_PENALTY = 5_000
# Penalty per required slot left empty. Huge: only if truly impossible.
UNFILLED_PENALTY = 1_000_000
# Penalty per day a pharmacist falls short of their pay-period target. Above every other
# preference and above a PRN request, so someone ends up short only when the hard rules
# (rest, consecutive days, filled-in cells, aptitude 0) leave no possible day for them.
TARGET_DEVIATION_PENALTY = 200_000
# Short rest between shifts: a gap at the Settings minimum or less is forbidden. Between the
# minimum and the preferred rest, the penalty grows with the square of the shortfall, reaching
# about REST_PENALTY just above the minimum (e.g. min 8 h, preferred 14 h: a 9 h gap ~280,
# 11 h ~100, 13 h ~10). For scale: a 3 -> 2 aptitude swap costs 60, a PRN shift 5,000.
REST_PENALTY = 400
# Sharing a role's shifts among pharmacists with the balanced aptitude (Settings): penalty per
# shift of difference between two such pharmacists (after adjusting for days available).
# Below a 3 -> 2 aptitude swap (60), so balance never costs coverage quality much.
BALANCE_PENALTY = 40
# Penalty for each run longer than MAX_CONSECUTIVE_SOFT days.
CONSECUTIVE_SOFT_PENALTY = 200
# Full-timer who worked Sat+Sun: penalty for working the following Monday / Tuesday.
WEEKEND_COMP_MONDAY_PENALTY = 150
WEEKEND_COMP_TUESDAY_PENALTY = 30
# Reward when someone keeps the same role on back-to-back workdays (Fri->Mon counts).
CONTINUITY_BONUS = 15

# ---- Solver ----
SOLVER_TIME_LIMIT_SECONDS = 60
def _workers() -> int:
    import os
    v = os.environ.get("SCHEDULER_WORKERS", "")
    return max(1, int(v)) if v.isdigit() else 8


# Parallel search strategies. Keep 8 even on a 2-CPU machine: CP-SAT runs a portfolio of
# different strategies, and the variety matters more than one CPU per strategy (tested: on
# 2 CPUs, 8 workers reach a no-shortfall schedule in ~15 s; 2 workers still hadn't at 60 s).
# Override with the SCHEDULER_WORKERS environment variable.
SOLVER_WORKERS = _workers()

# Labels the loader searches for (case-insensitive) to find each block in the sheet.
LABEL_SCHEDULE_DATES = "Date"
LABEL_SCHEDULE_DAYS = "Day"
LABEL_APTITUDE = "Aptitude Matrix"
LABEL_DAYS_PER_PERIOD = "Days Per Pay Period"
LABEL_REQUIREMENTS = "Requirement Matrix"
LABEL_SHIFTS_FULFILLED = "Shifts Fulfilled"   # output-only block written under the schedule
# Column headers. PID is the key that links schedule rows to aptitude rows. The Name column
# is optional: put it anywhere left of PID, headed "Name" (in the aptitude header row, and in
# the schedule's Date or Day row).
LABEL_PID = "PID"
LABEL_NAME = "Name"
# Requirement matrix column headers (Optional and Description columns are optional).
LABEL_ROLE = "Role"
LABEL_OPTIONAL = "Optional"
LABEL_DESCRIPTION = "Description"
LABEL_COLOR = "Color"            # swatch column: format the cell as the code should look
LABEL_PRN = "PRN"                # Aptitude column: Yes = PRN pharmacist
LABEL_PRN_SECTION = "PRN STAFF"  # optional separator row in the schedule, above the PRN rows
LABEL_COUNTS = "Counts toward days"
LABEL_WORKDAY = "Workday"
LABEL_COVERS = "Covers role"
LABEL_MEANING = "Meaning"
LABEL_START = "Start"            # shift start time (Requirements and Other Codes)
LABEL_END = "End"                # shift end time; earlier than Start = ends the next day
# Table for non-role codes' colors (PTO, OFF, PRJ), under the requirement matrix.
LABEL_OTHER_CODES = "Other Codes"
LABEL_CODE = "Code"
# Settings table: a cell labeled "Settings", with "Setting" and "Value" headers under it.
LABEL_SETTINGS = "Settings"
LABEL_SETTING_NAME = "Setting"
LABEL_SETTING_VALUE = "Value"

# Sheets the loader never reads (instructions and the output's report sheets).
INSTRUCTIONS_SHEET = "Instructions"
REPORT_SHEETS = ["Shift Balance", "PRN Requests", "Aptitude View", "Summary", "Pay Periods", "Role Counts", "Coverage Gaps", "Rule Check"]
# Hidden sheet recording which formatting the program applied per code, so a later run can tell
# program-colored cells (recolor if the Color cell changed) from hand-formatted ones (keep).
APPLIED_STYLES_SHEET = "_Applied Colors"
IGNORED_SHEETS = {INSTRUCTIONS_SHEET, APPLIED_STYLES_SHEET, *REPORT_SHEETS}

# Weekday tokens accepted in the schedule's Day row and the requirement matrix header.
WEEKDAY_TOKENS = {
    "M": 0, "MO": 0, "MON": 0, "MONDAY": 0,
    "T": 1, "TU": 1, "TUE": 1, "TUES": 1, "TUESDAY": 1,
    "W": 2, "WE": 2, "WED": 2, "WEDNESDAY": 2,
    "TR": 3, "TH": 3, "R": 3, "THU": 3, "THUR": 3, "THURS": 3, "THURSDAY": 3,
    "F": 4, "FR": 4, "FRI": 4, "FRIDAY": 4,
    "S": 5, "SA": 5, "SAT": 5, "SATURDAY": 5,
    "SU": 6, "SUN": 6, "SUNDAY": 6,
}
