"""CP-SAT model: fill the open cells of a Template.

Hard rules (never broken):
  - Locked cells are kept exactly as entered.
  - At most one code per pharmacist per day.
  - New role assignments only where aptitude > 0, only for pharmacists with a day target.
  - Required roles get exactly their headcount (an "unfilled" slack absorbs truly
    impossible slots at a huge penalty, so you always get a schedule back).
  - Optional roles never exceed their headcount.
  - No more than MAX_CONSECUTIVE_HARD on-duty days in a row (unless locks already do it).

Soft rules (traded off in the objective):
  - Never schedule anyone past their pay-period day target (hard); falling short is a heavy penalty.
  - Prefer higher aptitude (nonlinear weights).
  - Fill optional roles instead of giving project days.
  - Avoid runs longer than MAX_CONSECUTIVE_SOFT.
  - Full-timers who work a weekend get the next Monday (and ideally Tuesday) off.
  - Keep people in the same role on back-to-back workdays.
"""
from __future__ import annotations

import math
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, timedelta

from ortools.sat.python import cp_model

from . import config
from .loader import Template
from .validate import balance_groups, coverage_counts, prn_capacity, prn_eligible


@dataclass
class Solution:
    status: str
    objective: float | None
    bound: float | None
    seconds: float
    assignment: dict[tuple[str, date], str] = field(default_factory=dict)   # every non-blank cell
    solver_filled: set[tuple[str, date]] = field(default_factory=set)       # cells the solver added
    unfilled: list[tuple[str, date, int]] = field(default_factory=list)     # required role, day, count
    target_rows: list[dict] = field(default_factory=list)                   # per person per period
    long_runs: list[tuple[str, date, int]] = field(default_factory=list)    # pid, start, length > soft max
    comp_misses: list[tuple[str, date, str]] = field(default_factory=list)  # pid, Monday/Tuesday worked
    prn_requests: list[tuple[str, date, int]] = field(default_factory=list) # role, day, PRNs needed
    balance: list[dict] = field(default_factory=list)  # role, pid, shifts, fair share
    # short rests in the result: pid, day, code, next code, gap hours, whether the program made it
    short_rests: list[tuple[str, date, str, str, float, bool]] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.status in ("OPTIMAL", "FEASIBLE")


def _next_workday_pairs(dates: list[date]) -> list[tuple[date, date]]:
    """(d, next) pairs for continuity: consecutive dates, plus Friday -> Monday."""
    ds = set(dates)
    pairs = []
    for d in dates:
        nxt = d + timedelta(days=1)
        if nxt in ds:
            pairs.append((d, nxt))
        if d.weekday() == 4 and d + timedelta(days=3) in ds:
            pairs.append((d, d + timedelta(days=3)))
    return pairs


class _Progress(cp_model.CpSolverSolutionCallback):
    """Reports each improved schedule the solver finds (count, score, seconds)."""
    def __init__(self, fn):
        super().__init__()
        self.fn, self.n = fn, 0

    def on_solution_callback(self):
        self.n += 1
        try:
            self.fn(self.n, self.ObjectiveValue(), self.WallTime())
        except Exception:
            pass


def solve(t: Template, time_limit: float | None = None, log: bool = False,
          on_solution=None) -> Solution:
    m = cp_model.CpModel()
    staff = [p for p in t.schedule_order if p in t.staff]
    prj = t.settings.project_code
    locked_cov = coverage_counts(t, t.locked)

    def open_count(role: str, d: date) -> int:
        return t.required(role, d) - locked_cov[(role, d)]

    # ---------------- decision variables
    # x[p, d, code] = 1 if pharmacist p has `code` on day d. Locked cells get a variable fixed to 1.
    x: dict[tuple[str, date, str], cp_model.IntVar] = {}
    for p in staff:
        s = t.staff[p]
        for d in t.dates:
            if (p, d) in t.locked:
                code = t.locked[(p, d)]
                v = m.NewBoolVar(f"L_{p}_{d}_{code}")
                m.Add(v == 1)
                x[p, d, code] = v
                continue
            if not s.target_days:
                continue  # target 0/blank: only works locked cells
            options = []
            for role in t.roles:
                if s.aptitude.get(role, 0) > 0 and open_count(role, d) > 0:
                    options.append(role)
            if t.settings.project_weekends or d.weekday() < 5:
                options.append(prj)
            day_vars = []
            for code in options:
                v = m.NewBoolVar(f"x_{p}_{d}_{code}")
                x[p, d, code] = v
                day_vars.append(v)
            if len(day_vars) > 1:
                m.AddAtMostOne(day_vars)

    # Index for speed
    by_pd: dict[tuple[str, date], list[tuple[str, cp_model.IntVar]]] = {}
    for (p, d, code), v in x.items():
        by_pd.setdefault((p, d), []).append((code, v))

    def on_duty(p: str, d: date):
        return sum(v for code, v in by_pd.get((p, d), []) if t.code_on_duty(code))

    def paid(p: str, d: date):
        return sum(v for code, v in by_pd.get((p, d), []) if t.code_counts_for_pay(code))

    def all_locked(p: str, days: list[date]) -> bool:
        return all((p, d) in t.locked for d in days)

    obj = []        # terms to maximize

    # ---------------- coverage (regular staff first, then PRN requests, then left empty)
    # PRNs are a pool: the program says how many PRN shifts each role needs each day, never who.
    unfilled_vars: dict[tuple[str, date], cp_model.IntVar] = {}
    prn_vars: dict[tuple[str, date], cp_model.IntVar] = {}
    prn_by_day: dict[date, list] = {}
    for role in t.requirements:
        use_prn = prn_eligible(t, role)
        for d in t.dates:
            n_open = open_count(role, d)
            if n_open <= 0:
                continue
            vs = [x[p, d, role] for p in staff if (p, d, role) in x and (p, d) not in t.locked]
            cap = min(n_open, prn_capacity(t, role, d)) if use_prn else 0
            q = None
            if cap > 0:
                q = m.NewIntVar(0, cap, f"prn_{role}_{d}")
                prn_vars[role, d] = q
                prn_by_day.setdefault(d, []).append(q)
                obj.append(-config.PRN_REQUEST_PENALTY * q)
            if role in t.optional_roles:
                if q is not None:
                    # optional + PRN allowed: fill it, by a PRN if no regular can
                    e = m.NewIntVar(0, n_open, f"open_{role}_{d}")
                    m.Add(sum(vs) + q + e == n_open)
                    obj.append(-(config.PRN_REQUEST_PENALTY + 1) * e)
                    obj.append(config.OPTIONAL_FILL_BONUS * sum(vs))
                elif vs:
                    m.Add(sum(vs) <= n_open)
                    obj.append(config.OPTIONAL_FILL_BONUS * sum(vs))
            else:
                u = m.NewIntVar(0, n_open, f"unfilled_{role}_{d}")
                m.Add(sum(vs) + (q if q is not None else 0) + u == n_open)
                unfilled_vars[role, d] = u
                obj.append(-config.UNFILLED_PENALTY * u)
    # each PRN works at most one shift a day
    for d, qs in prn_by_day.items():
        free = sum(1 for p in t.prn_pids if (p, d) not in t.locked)
        m.Add(sum(qs) <= free)

    # ---------------- aptitude reward, project-day penalty
    for (p, d, code), v in x.items():
        if (p, d) in t.locked:
            continue
        if code == prj:
            obj.append(-config.PROJECT_DAY_PENALTY * v)
        elif code in t.optional_roles:
            obj.append(config.OPTIONAL_APTITUDE_WEIGHT[t.staff[p].aptitude[code]] * v)
        else:
            obj.append(config.APTITUDE_WEIGHT[t.staff[p].aptitude[code]] * v)

    # ---------------- pay-period targets
    dev_vars = {}
    for per in t.periods():
        n = len(per["dates"])
        for p in staff:
            s = t.staff[p]
            if not s.target_days:
                continue  # 0/blank target: satisfied by whatever is locked
            frac = s.target_days * n / t.settings.pay_period_days
            target = s.target_days if per["complete"] else round(frac)
            # Hard cap: the program never schedules anyone past their target. Days typed into the
            # input are always kept, so if those alone exceed it, the program just adds nothing.
            locked_paid = sum(1 for d in per["dates"] if (p, d) in t.locked
                              and t.code_counts_for_pay(t.locked[p, d]))
            cap_days = max(s.target_days if per["complete"] else math.ceil(frac), locked_paid)
            days = sum(paid(p, d) for d in per["dates"])
            m.Add(days <= cap_days)
            under = m.NewIntVar(0, n, f"under_{p}_{per['index']}")
            m.Add(days + under >= target)
            obj.append(-config.TARGET_DEVIATION_PENALTY * under)
            dev_vars[p, per["index"]] = (target, None, under)
            # spread project days
            prj_vars = [x[p, d, prj] for d in per["dates"] if (p, d, prj) in x and (p, d) not in t.locked]
            cap = t.settings.project_soft_cap
            if cap is not None and len(prj_vars) > cap:
                excess = m.NewIntVar(0, n, f"prj_excess_{p}_{per['index']}")
                m.Add(sum(prj_vars) - cap <= excess)
                obj.append(-config.PROJECT_SPREAD_PENALTY * excess)

    # ---------------- consecutive days
    H, S = t.settings.max_consecutive_hard, t.settings.max_consecutive_soft
    for p in staff:
        for i in range(len(t.dates) - H):
            window = t.dates[i:i + H + 1]
            if not all_locked(p, window):
                m.Add(sum(on_duty(p, d) for d in window) <= H)
        for i in range(len(t.dates) - S):
            window = t.dates[i:i + S + 1]
            if all_locked(p, window):
                continue
            too_long = m.NewBoolVar(f"run_{p}_{window[0]}")
            m.Add(sum(on_duty(p, d) for d in window) <= S + too_long)
            obj.append(-config.CONSECUTIVE_SOFT_PENALTY * too_long)

    # ---------------- weekend comp days for full-timers
    date_set = set(t.dates)
    for p in staff:
        s = t.staff[p]
        if not t.settings.weekend_comp or not s.target_days or s.target_days < t.settings.full_time_target:
            continue
        for d in t.dates:
            if d.weekday() != 5:
                continue
            sun = d + timedelta(days=1)
            for offset, weight in ((2, config.WEEKEND_COMP_MONDAY_PENALTY),
                                   (3, config.WEEKEND_COMP_TUESDAY_PENALTY)):
                wd = d + timedelta(days=offset)
                if sun not in date_set or wd not in date_set or all_locked(p, [d, sun, wd]):
                    continue
                miss = m.NewBoolVar(f"comp_{p}_{wd}")
                m.Add(on_duty(p, d) + on_duty(p, sun) + on_duty(p, wd) - 2 <= miss)
                obj.append(-weight * miss)

    # ---------------- rest between shifts
    # gap = start of the next day's shift - end of this shift. <= minimum: forbidden.
    # Between minimum and preferred: penalty growing with the square of the shortfall.
    lo, hi = t.settings.min_rest_hours * 60, t.settings.preferred_rest_hours * 60
    def rest_penalty(gap: int) -> int:
        if gap >= hi or hi <= lo:
            return 0
        return max(1, round(config.REST_PENALTY * ((hi - gap) / (hi - lo)) ** 2))
    for p in staff:
        for d1 in t.dates:
            d2 = d1 + timedelta(days=1)
            if d2 not in date_set:
                continue
            for a, va in by_pd.get((p, d1), []):
                if a not in t.shift_times:
                    continue
                for b, vb in by_pd.get((p, d2), []):
                    if b not in t.shift_times:
                        continue
                    a_locked, b_locked = (p, d1) in t.locked, (p, d2) in t.locked
                    if a_locked and b_locked:
                        continue  # both typed in: kept as entered (the checker reports it)
                    gap = 1440 + t.shift_times[b][0] - t.shift_times[a][1]
                    if gap <= lo:
                        m.Add(va + vb <= 1)
                        continue
                    pen = rest_penalty(gap)
                    if not pen:
                        continue
                    if a_locked:
                        obj.append(-pen * vb)
                    elif b_locked:
                        obj.append(-pen * va)
                    else:
                        both = m.NewBoolVar(f"rest_{p}_{d1}_{a}_{b}")
                        m.Add(va + vb - 1 <= both)
                        obj.append(-pen * both)

    # ---------------- balance a role's shifts among its top-rated pharmacists
    groups = balance_groups(t)
    shift_count = {}
    for role, members in groups.items():
        counts = {}
        for p in members:
            counts[p] = sum(v for d in t.dates for code, v in by_pd.get((p, d), [])
                            if t.covered_role(code) == role)
            shift_count[role, p] = counts[p]
        w = {p: members[p] for p in members}
        W = sum(w.values())
        total = sum(counts.values())
        coef = max(1, round(config.BALANCE_PENALTY / W))
        for p in members:
            # |W*count_p - w_p*total| = W * (shifts away from p's fair share)
            dev = m.NewIntVar(0, W * 2 * len(t.dates), f"bal_{role}_{p}")
            m.Add(dev >= W * counts[p] - w[p] * total)
            m.Add(dev >= w[p] * total - W * counts[p])
            obj.append(-coef * dev)

    # ---------------- continuity
    for p in staff:
        for d1, d2 in _next_workday_pairs(t.dates):
            if (p, d1) in t.locked and (p, d2) in t.locked:
                continue
            for code, v1 in by_pd.get((p, d1), []):
                if code not in t.roles:
                    continue
                v2 = x.get((p, d2, code))
                if v2 is None:
                    continue
                same = m.NewBoolVar(f"same_{p}_{d1}_{code}")
                m.Add(same <= v1)
                m.Add(same <= v2)
                obj.append(config.CONTINUITY_BONUS * same)

    m.Maximize(sum(obj))

    # ---------------- solve
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit or t.settings.time_limit
    solver.parameters.num_workers = config.SOLVER_WORKERS
    solver.parameters.log_search_progress = log
    status = solver.Solve(m, _Progress(on_solution)) if on_solution else solver.Solve(m)
    name = solver.StatusName(status)
    sol = Solution(status=name, objective=None, bound=None, seconds=solver.WallTime())
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return sol
    sol.objective = solver.ObjectiveValue()
    sol.bound = solver.BestObjectiveBound()

    # ---------------- read the answer back
    for (p, d, code), v in x.items():
        if solver.Value(v):
            sol.assignment[p, d] = code
            if (p, d) not in t.locked:
                sol.solver_filled.add((p, d))
    for p in staff:
        for d1 in t.dates:
            d2 = d1 + timedelta(days=1)
            a, b = sol.assignment.get((p, d1)), sol.assignment.get((p, d2))
            if a in t.shift_times and b in t.shift_times:
                gap = 1440 + t.shift_times[b][0] - t.shift_times[a][1]
                if gap < hi:
                    made = (p, d1) in sol.solver_filled or (p, d2) in sol.solver_filled
                    sol.short_rests.append((p, d1, a, b, gap / 60, made))
    for role, members in groups.items():
        W = sum(members.values())
        got = {p: sum(1 for d in t.dates if t.covered_role(sol.assignment.get((p, d), "")) == role)
               for p in members}
        total = sum(got.values())
        for p in members:
            sol.balance.append({"role": role, "pid": p, "shifts": got[p],
                                "share": total * members[p] / W, "weight": members[p]})
    for (role, d), q in prn_vars.items():
        if solver.Value(q):
            sol.prn_requests.append((role, d, solver.Value(q)))
    for (role, d), u in unfilled_vars.items():
        if solver.Value(u):
            sol.unfilled.append((role, d, solver.Value(u)))
    for per in t.periods():
        for p in staff:
            s = t.staff[p]
            days = sum(1 for d in per["dates"] if (p, d) in sol.assignment
                       and t.code_counts_for_pay(sol.assignment[p, d]))
            target = dev_vars.get((p, per["index"]), (s.target_days, None, None))[0]
            sol.target_rows.append({
                "pid": p, "start": per["start"], "end": per["end"], "complete": per["complete"],
                "target": target, "actual": days,
                "deviation": 0 if not s.target_days else days - target,
                "project_days": sum(1 for d in per["dates"] if sol.assignment.get((p, d)) == prj),
            })
    for p in staff:
        run, start, touched = 0, None, False
        for d in t.dates + [None]:
            if d is not None and (p, d) in sol.assignment and t.code_on_duty(sol.assignment[p, d]):
                run += 1
                start = start or d
                touched = touched or (p, d) in sol.solver_filled
            else:
                if run > S and touched:   # skip runs made entirely of locked cells
                    sol.long_runs.append((p, start, run))
                run, start, touched = 0, None, False
        s = t.staff[p]
        if t.settings.weekend_comp and s.target_days and s.target_days >= t.settings.full_time_target:
            for d in t.dates:
                if d.weekday() != 5:
                    continue
                sun = d + timedelta(days=1)
                worked_weekend = all((p, dd) in sol.assignment and t.code_on_duty(sol.assignment[p, dd])
                                     for dd in (d, sun))
                if not worked_weekend:
                    continue
                for offset, label in ((2, "Monday"), (3, "Tuesday")):
                    wd = d + timedelta(days=offset)
                    if (p, wd) in sol.solver_filled and t.code_on_duty(sol.assignment[p, wd]):
                        sol.comp_misses.append((p, wd, label))
    return sol
