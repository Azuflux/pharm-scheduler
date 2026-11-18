# Pharmacist Scheduler

Fills the open cells of a schedule workbook with CP-SAT and writes a new .xlsx.

## Setup (once)
Copy these files into `~/pharm-scheduler` (next to your `.venv`), then:

```bash
cd ~/pharm-scheduler
source .venv/bin/activate
pip install ortools openpyxl
```

## Web app
`streamlit run app.py` opens the scheduler in your browser: upload a workbook, watch it build,
download the result. To host it online (privately), follow **DEPLOY.md**.

## Use
```bash
python check_template.py "Pharmacist Schedule Template.xlsx"    # report problems, no solving
python solve_schedule.py "Pharmacist Schedule Template.xlsx"    # writes ..._solved.xlsx next to it
python solve_schedule.py template.xlsx --time-limit 180 --out "March schedule.xlsx"
python convert_template.py old_single_sheet.csv                 # older layout -> new workbook
```
Tip: type the command and a space, then drag the file from Finder into Terminal to paste its path.

## The workbook
| Sheet | What it holds |
|---|---|
| Instructions | How to use and maintain the workbook (the program ignores this sheet) |
| Schedule | Name, PID, one column per date. Fill in what's settled; leave the rest blank |
| Aptitude | One row per pharmacist: Name, PID, Days Per Pay Period, a 0-4 score per role |
| Requirements | One row per role: Color, Optional (Yes/No), pharmacists needed Mon-Sun, Description; Other Codes colors below |
| Settings | Department rules (pay-period start/length, consecutive-day limits, comp days, project-day spreading, solver time) |

Cells have dropdowns/number limits where it helps. Each block is found by its label (the table
title, `Date`/`Day`, `Name`/`PID`/`Days Per Pay Period`, `Role`/`Color`/`Optional`, `Setting`/`Value`),
so tables can move, and rows/columns can be inserted freely as long as no blank row or column is
left inside a table. The loader stops with the exact cell if it finds one.

Schedule codes: any role code, `PTO` (counts toward days, not a workday), `OFF`, `PRJ` (project day).

## PRN pharmacists
Mark PRNs with PRN = Yes on the Aptitude sheet (their scores say which roles they can work) and
list them under a "PRN STAFF" row on the Schedule. The program never assigns a particular PRN:
when regular staff can't cover a role a PRN can work, it adds a request instead (PRN NEEDED row
under the schedule, purple cells in Shifts Fulfilled, and a PRN Requests sheet listing each
shift with the PRNs able to take it). Type the accepted shift into that PRN's row and re-run.
A PRN with anything typed on a day isn't counted as available that day. Setting "Request PRNs
for optional roles" decides whether optional roles (CP) can trigger requests.

## Codes
Every non-role code (PTO, X, HPTO, MIDY, training codes, ...) is listed in the Other Codes table
on the Requirements sheet with: Counts toward days, Workday, Covers role (e.g. ON8 covers ON),
Meaning, and its Color. Unknown codes are reported with their cell. The project-day code is a
setting (PROJ).

## Copying a live schedule's look
`import_format.py live.xlsx --sheet "<sheet>" --template template.xlsx --out new.xlsx
--project-code PROJ [--add-prn 9 --prn-roles CH,CP]` reads one sheet of the department's
workbook, captures each code's fill/text color/font (converted to exact RGB), the title/header/
PRN STAFF styling and the legend's shift times, adds unknown codes to Other Codes (meanings
marked "assumed"), and rebuilds the template. Only the named sheet is read.

## Colors
The **Color** column on Requirements (and the **Other Codes** table under it for PTO/OFF/PRJ) sets
how each code looks on the finished schedule: format the Color cell the way the role should look
(fill, text color, bold/italic) and the program copies it. `convert_template.py` run on an existing
color-coded .xlsx schedule fills these in automatically from the colors already used. Cells you
format by hand on the Schedule keep their own formatting; changing a Color cell recolors every
cell the program colored before (tracked in a hidden "_Applied Colors" sheet). Without any Color
formatting (e.g. a CSV), the Schedule is colored by aptitude as before.

## The output
A copy of the workbook with blanks filled and colored by role; to the right of the dates, the
names again plus days given / target for each pay period and in total (green met, red short,
yellow over from typed-in days); an **Aptitude View** sheet (same
schedule colored by aptitude: dark green 4, green 3, yellow 2, orange 1, blue project day), a
**Shifts Fulfilled** table under the schedule (filled/required per
role per day; red = required role not covered, gray = optional role open), and report sheets:
Summary (incl. settings used), Pay Periods, Role Counts, Coverage Gaps, Rule Check.
An output can be fed back in as the next input: everything in it is kept, and old report
sheets and the old Shifts Fulfilled table are replaced.

## Live checks
The days/target columns and the Shifts Fulfilled table are Excel formulas (in the template and
in every output), so editing a code on the Schedule, a target on Aptitude or a count on
Requirements updates them immediately: you can test a change before re-running. Inputs are
looked up by key: each pharmacist row finds its target on Aptitude by its own PID, and each
Shifts Fulfilled row finds its requirement by role name, so rows can be moved or sorted on any
sheet. Results are also stored in the file, so it previews correctly and can be re-uploaded without being opened
in Excel first. Purple in Shifts Fulfilled marks a PRN slot requested by the last run that hasn't
been entered yet; it turns green once a PRN's row has that code. The blocks are rebuilt on every
run; after adding a pharmacist or role, rebuild them without solving with
`python refresh_checks.py "<workbook>.xlsx"` or the web app's "Only update the live checks".

## Hard rules (never broken)
- Cells typed into the Schedule are kept exactly.
- No one is scheduled past their Days Per Pay Period. Extra days happen only when typed in.
- No new role where aptitude is 0; one code per person per day.
- No more consecutive workdays than the Settings maximum.
- No shift starting within the Settings' "Minimum hours between shifts" (default 8) of the previous
  day's shift ending. Shift hours come from Start/End on Requirements and Other Codes.
- Required roles covered exactly; if regular staff can't, a PRN is requested where PRNs can work
  the role, otherwise the slot is left empty (red).

## Preferences (traded off; weights in config.py)
Hitting each pay-period target exactly (falling short is weighted above everything except leaving a
required slot empty, so it happens only when the hard rules leave no possible day); higher aptitude (1 = last resort); rests shorter than "Preferred
hours between shifts" (default 14), penalized more the closer they get to the minimum; project
days spread out; each role's shifts shared evenly among its aptitude-4 pharmacists (Settings:
level, and Equal vs. by days available; see the Shift Balance sheet); optional roles before project days; runs within the preferred length; weekend
comp days; same role on back-to-back days.

## Where rules live
- **Settings sheet**: department policy, editable by managers.
- **`scheduler/config.py`**: solver scoring weights (how aptitude, spreading, comp days, continuity
  trade off) and the fixed codes. Technical tuning, not day-to-day policy.

## Files
- `scheduler/loader.py`: reads .xlsx (any sheets) or older single-sheet .csv into a `Template`
- `scheduler/settings.py`: the Settings sheet's rows, types, defaults and help text
- `scheduler/validate.py`: pre-solve checks
- `scheduler/model.py`: the CP-SAT model
- `scheduler/writer.py`: output workbook
- `scheduler/template_builder.py`: builds the multi-sheet template
- `scheduler/colors.py`: converts palette/theme colors to RGB
- `refresh_checks.py`: rebuilds the live check blocks and Instructions text without solving
- `import_format.py`: copies a live schedule's formatting, codes and legend into a template
- `scheduler/pipeline.py`: one run (load, check, solve, write) with progress reporting
- `app.py`: the web app; `requirements.txt` and `.streamlit/config.toml` for hosting
- `check_template.py`, `solve_schedule.py`, `convert_template.py`: command-line entry points
