"""Fill a schedule template and write the result to a new .xlsx.

Usage:
    python solve_schedule.py template.xlsx
    python solve_schedule.py template.xlsx --out solved.xlsx --time-limit 120
    python solve_schedule.py template.csv --year 2027 --sheet "Schedule" --log
"""
import argparse
import sys
from pathlib import Path

from scheduler.pipeline import run


def main() -> int:
    ap = argparse.ArgumentParser(description="Fill a pharmacist schedule template.")
    ap.add_argument("template", help=".xlsx or .csv template")
    ap.add_argument("--out", help="output .xlsx (default: <template>_solved.xlsx next to the template)")
    ap.add_argument("--sheet", help="worksheet name (default: first sheet)")
    ap.add_argument("--year", type=int, help="year for dates written without one")
    ap.add_argument("--time-limit", type=float, help="seconds the solver may search (default in config.py)")
    ap.add_argument("--log", action="store_true", help="show the solver's progress log")
    args = ap.parse_args()

    src = Path(args.template)
    out = Path(args.out) if args.out else src.with_name(src.stem + "_solved.xlsx")
    last = {"stage": None}

    def progress(stage, frac, msg):
        if stage != last["stage"] or stage == "done":
            print(msg + ("..." if stage not in ("done",) else ""))
            last["stage"] = stage

    res = run(src, out, time_limit=args.time_limit, on_progress=progress, sheet=args.sheet, year=args.year)
    if not res.ok:
        print(res.message)
        for i in res.errors:
            print(f"  [{i.category}] {i.message}")
        return 2 if res.template is None else 1
    s = res.summary
    print(f"Status: {s['status']} in {s['seconds']:.0f}s"
          + ("" if s["status"] == "OPTIMAL" else f" (within {s['gap']:,.0f} points of the best possible)"))
    print(f"Filled {s['filled']} cells; {s['required_empty']} required slots left empty; "
          f"{s['short_periods'] + s['over_periods']} pay-period target deviations.")
    print(f"Wrote {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
