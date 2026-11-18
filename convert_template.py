"""Convert a template (older single-sheet .csv/.xlsx, or a current one) into the multi-sheet
.xlsx layout with Instructions, Schedule, Aptitude, Requirements and Settings sheets.

Usage:
    python convert_template.py old_template.csv
    python convert_template.py old_template.csv --out "Pharmacist Schedule Template.xlsx"
"""
import argparse
import sys
from pathlib import Path

from scheduler import LoaderError, load_template
from scheduler.template_builder import build_template


def main() -> int:
    ap = argparse.ArgumentParser(description="Convert a schedule template to the multi-sheet layout.")
    ap.add_argument("template")
    ap.add_argument("--out", help="output .xlsx (default: <template>_template.xlsx)")
    ap.add_argument("--year", type=int, help="year for dates written without one")
    args = ap.parse_args()
    try:
        t = load_template(args.template, year=args.year)
    except (LoaderError, FileNotFoundError) as e:
        print(f"Couldn't load the template: {e}")
        return 2
    src = Path(args.template)
    out = Path(args.out) if args.out else src.with_name(src.stem + "_template.xlsx")
    build_template(t, out)
    print(f"Wrote {out}")
    print(f"  {len(t.schedule_order)} pharmacists, {len(t.roles)} roles, {len(t.dates)} days, "
          f"{len(t.locked)} filled-in cells; optional roles: {', '.join(sorted(t.optional_roles)) or 'none'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
