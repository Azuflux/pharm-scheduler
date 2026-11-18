"""Load a schedule template and print a validation report.

Usage:
    python check_template.py template.xlsx
    python check_template.py template.xlsx --sheet "Schedule" --year 2027 --out report.txt
"""
import argparse
import sys

from scheduler import LoaderError, load_template, summary, validate

LEVELS = [("error", "ERRORS (must fix before solving)"),
          ("warning", "WARNINGS (the schedule can be built, but a rule will bend)"),
          ("info", "NOTES")]


def main() -> int:
    ap = argparse.ArgumentParser(description="Check a pharmacist schedule template.")
    ap.add_argument("template", help=".xlsx or .csv template")
    ap.add_argument("--sheet", help="worksheet name (default: first sheet)")
    ap.add_argument("--year", type=int, help="year for dates written without one")
    ap.add_argument("--out", help="also write the report to this text file")
    args = ap.parse_args()

    try:
        t = load_template(args.template, sheet=args.sheet, year=args.year)
    except (LoaderError, FileNotFoundError) as e:
        print(f"Couldn't load the template: {e}")
        return 2

    issues = t.issues + validate(t)
    lines = ["=" * 72, "TEMPLATE SUMMARY", "=" * 72, *summary(t), ""]
    for level, title in LEVELS:
        group = [i for i in issues if i.level == level]
        lines.append(f"{title}: {len(group)}")
        lines.append("-" * 72)
        for i in group:
            lines.append(f"[{i.category}] {i.message}")
        lines.append("")

    report = "\n".join(lines)
    print(report)
    if args.out:
        with open(args.out, "w") as fh:
            fh.write(report + "\n")
    return 1 if any(i.level == "error" for i in issues) else 0


if __name__ == "__main__":
    sys.exit(main())
