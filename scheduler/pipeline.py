"""One run of the scheduler: load -> check -> solve -> write, with progress reporting.

Used by the command line (solve_schedule.py) and the web app (app.py). The solver runs in a
background thread; progress callbacks are made from the calling thread, so a UI can update
safely while it waits.
"""
from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from .loader import Issue, LoaderError, Template, load_template
from .model import Solution, solve
from .validate import validate
from .writer import write_output

# on_progress(stage, fraction 0..1, message)
Progress = Callable[[str, float, str], None]


@dataclass
class RunResult:
    ok: bool
    message: str = ""
    template: Template | None = None
    issues: list[Issue] = field(default_factory=list)
    solution: Solution | None = None
    output: Path | None = None
    summary: dict = field(default_factory=dict)

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.level == "error"]

    @property
    def warnings(self) -> list[Issue]:
        return [i for i in self.issues if i.level == "warning"]


def summarize(t: Template, sol: Solution) -> dict:
    short = [r for r in sol.target_rows if r["deviation"] < 0]
    over = [r for r in sol.target_rows if r["deviation"] > 0]
    gap = (sol.bound - sol.objective) if sol.objective is not None else None
    return {
        "status": sol.status,
        "seconds": sol.seconds,
        "gap": gap,
        "pharmacists": len(t.schedule_order),
        "days": len(t.dates),
        "start": t.dates[0],
        "end": t.dates[-1],
        "filled": len(sol.solver_filled),
        "required_empty": sum(n for _, _, n in sol.unfilled),
        "prn_needed": sum(n for _, _, n in sol.prn_requests),
        "short_periods": len(short),
        "over_periods": len(over),
        "program_short_rests": sum(1 for *_, made in sol.short_rests if made),
    }


def run(input_path: str | Path, output_path: str | Path, time_limit: float | None = None,
        on_progress: Progress | None = None, sheet: str | None = None,
        year: int | None = None) -> RunResult:
    report = on_progress or (lambda *a: None)

    report("load", 0.02, "Reading the workbook")
    try:
        t = load_template(input_path, sheet=sheet, year=year)
    except (LoaderError, FileNotFoundError) as e:
        return RunResult(False, f"Couldn't read the template: {e}")
    except Exception as e:  # a damaged or non-template file
        return RunResult(False, f"Couldn't open this file as a schedule template ({type(e).__name__}: {e}).")

    report("check", 0.05, "Checking the template")
    issues = t.issues + validate(t)
    res = RunResult(False, template=t, issues=issues)
    if res.errors:
        res.message = "The template has problems to fix before a schedule can be built."
        return res

    limit = time_limit or t.settings.time_limit
    found = {"n": 0, "at": None}
    box: dict = {}

    def on_solution(n, objective, seconds):
        found["n"], found["at"] = n, seconds

    def work():
        try:
            box["sol"] = solve(t, time_limit=limit, on_solution=on_solution)
        except Exception as e:  # surfaced to the caller below
            box["err"] = e

    report("solve", 0.08, "Setting up the rules")
    th = threading.Thread(target=work, daemon=True)
    started = time.monotonic()
    th.start()
    while th.is_alive():
        th.join(0.5)
        elapsed = time.monotonic() - started
        frac = 0.08 + 0.84 * min(elapsed / max(limit, 1), 1.0)
        if found["n"]:
            msg = (f"Searching for the best schedule: {elapsed:.0f} of {limit:.0f} s "
                   f"({found['n']} improvement{'s' if found['n'] != 1 else ''} so far)")
        else:
            msg = f"Searching for a first schedule: {elapsed:.0f} of {limit:.0f} s"
        report("solve", frac, msg)
    if "err" in box:
        res.message = f"The solver stopped with an error: {box['err']}"
        return res
    sol: Solution = box["sol"]
    res.solution = sol
    if not sol.ok:
        res.message = f"No schedule could be found (solver status: {sol.status})."
        return res

    report("write", 0.94, "Writing the workbook")
    out = Path(output_path)
    write_output(t, sol, issues, out)
    res.ok, res.output, res.summary = True, out, summarize(t, sol)
    res.message = "Schedule built."
    report("done", 1.0, "Done")
    return res
