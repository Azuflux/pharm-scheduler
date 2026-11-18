"""Rebuild the live check blocks (days/target per pay period and Shifts Fulfilled) and the
Instructions text in a template or schedule, without running the solver.

Use it after adding or removing a pharmacist or role, so the formulas cover the new rows.

    python refresh_checks.py "Pharmacist Schedule Template.xlsx"            # in place
    python refresh_checks.py in.xlsx --out refreshed.xlsx
"""
import argparse
import sys

from scheduler.loader import LoaderError
from scheduler.template_builder import refresh_template


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("workbook")
    ap.add_argument("--out", help="write here instead of updating the file in place")
    a = ap.parse_args()
    try:
        out = refresh_template(a.workbook, a.out)
    except LoaderError as e:
        print(f"Problem: {e}", file=sys.stderr)
        return 1
    print(f"Updated {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
