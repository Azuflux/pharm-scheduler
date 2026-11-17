"""Pre-solve checks on a loaded Template.

Locked (pre-filled) cells are always trusted. Checks here report where the locks make
the rest of the problem impossible or force a rule to bend, so you see it before solving.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, timedelta

from . import config
from .loader import Issue, Template


def _fmt(d: date) -> str:
    return f"{d.month}/{d.day}"


def _fmt_dates(ds: list[date], limit: int = 8) -> str:
    shown = ", ".join(_fmt(d) for d in ds[:limit])
    return shown + (f" (+{len(ds) - limit} more)" if len(ds) > limit else "")


def open_slots(t: Template) -> dict[tuple[str, date], int]:
    """Role-days still needing someone after locked cells are counted."""
    locked_count = coverage_counts(t, t.locked)
    out = {}
    for d in t.dates:
        for role in t.requirements:
            n = t.required(role, d) - locked_count[(role, d)]
            if n > 0:
                out[(role, d)] = n
    return out


def coverage_counts(t: Template, cells: dict) -> Counter:
    """(role, date) -> people covering it, counting codes that cover a role (e.g. ON8 -> ON)."""
    c = Counter()
    for (pid, d), code in cells.items():
        role = t.covered_role(code)
        if role:
            c[(role, d)] += 1
    return c


def schedulable(t: Template, pid: str) -> bool:
    """Whether the solver may give this person new assignments (not just locked ones)."""
    s = t.staff.get(pid)
    return bool(s and s.target_days and not s.prn)


def balance_groups(t: Template) -> dict[str, dict[str, int]]:
    """role -> {pid: weight} for roles whose shifts should be shared among pharmacists with the
    balanced aptitude. Only roles with open slots (something for the program to decide) and at
    least two such regular pharmacists with a day target. Weights: equal, or days available over
    the schedule (target days minus typed-in PTO-type days), reduced to small integers."""
    level = t.settings.balance_aptitude
    if not level:
        return {}
    import math
    open_roles = {r for (r, d) in open_slots(t)}
    groups = {}
    for role in t.requirements:
        if role not in open_roles:
            continue
        members = [p for p in t.schedule_order if schedulable(t, p)
                   and t.staff[p].aptitude.get(role, 0) == level]
        if len(members) < 2:
            continue
        if t.settings.balance_mode == "Equal":
            w = {p: 1 for p in members}
        else:
            w = {}
            for p in members:
                avail = 0
                for per in t.periods():
                    n = len(per["dates"])
                    avail += t.staff[p].target_days * n / t.settings.pay_period_days
                away = sum(1 for d in t.dates if (p, d) in t.locked
                           and t.code_counts_for_pay(t.locked[p, d]) and not t.code_on_duty(t.locked[p, d]))
                w[p] = max(1, round(avail - away))
            g = math.gcd(*w.values())
            w = {p: v // g for p, v in w.items()}
        groups[role] = w
    return groups


def prn_capacity(t: Template, role: str, d: date) -> int:
    """PRN pharmacists able to work `role` who aren't already entered on day d."""
    return sum(1 for p in t.prn_pids
               if t.staff[p].aptitude.get(role, 0) > 0 and (p, d) not in t.locked)


def prn_eligible(t: Template, role: str) -> bool:
    if role in t.optional_roles and not t.settings.prn_for_optional:
        return False
    return any(t.staff[p].aptitude.get(role, 0) > 0 for p in t.prn_pids)


def validate(t: Template) -> list[Issue]:
    issues: list[Issue] = []
    add = lambda lvl, cat, msg: issues.append(Issue(lvl, cat, msg))

    # ---- staff lists agree
    sched, apt = set(t.schedule_order), set(t.staff)
    for pid in [p for p in t.schedule_order if p not in apt]:
        add("error", "staff", f"'{pid}' is on the schedule but has no aptitude row.")
    for pid in [p for p in t.staff if p not in sched]:
        add("info", "staff", f"'{pid}' has an aptitude row but isn't on the schedule.")

    # ---- roles agree
    for role in t.requirements:
        if role not in t.roles:
            add("error", "roles", f"Requirement matrix lists '{role}', which isn't in the aptitude matrix.")
    for role in t.roles:
        if role not in t.requirements:
            add("warning", "roles", f"'{role}' has an aptitude column but no requirement row; it will never be scheduled.")

    # ---- locks that exceed the requirement (locks win; coverage will be higher than required)
    locked_count = coverage_counts(t, t.locked)
    over = defaultdict(list)
    for (code, d), n in locked_count.items():
        if code in t.requirements and n > t.required(code, d):
            over[code].append(d)
    for role, ds in over.items():
        add("warning", "over-coverage",
            f"{role}: more people locked than required on {_fmt_dates(sorted(ds))}. "
            f"Locks win, so these days will be over-covered.")

    # ---- slots nobody can fill
    slots = open_slots(t)
    locked_days = defaultdict(set)
    for (pid, d) in t.locked:
        locked_days[pid].add(d)
    short = defaultdict(lambda: [0, []])
    prn_needed = defaultdict(lambda: [0, []])
    for (role, d), need in slots.items():
        eligible = [p for p in t.schedule_order
                    if schedulable(t, p) and t.staff[p].aptitude.get(role, 0) > 0
                    and d not in locked_days[p]]
        gap = need - len(eligible)
        if gap > 0 and prn_eligible(t, role):
            by_prn = min(gap, prn_capacity(t, role, d))
            if by_prn:
                prn_needed[role][0] += by_prn
                prn_needed[role][1].append(d)
            gap -= by_prn
        if gap > 0:
            short[role][0] += gap
            short[role][1].append(d)
    for role, (n, ds) in sorted(prn_needed.items()):
        add("info", "PRN",
            f"{role}: no regular pharmacist is free for {n} slot(s); a PRN will be requested. "
            f"Days: {_fmt_dates(sorted(ds))}")
    for role, (n, ds) in sorted(short.items()):
        anyone = [p for p in t.staff if t.staff[p].aptitude.get(role, 0) > 0]
        why = ("no pharmacist has aptitude > 0 for this role" if not anyone
               else "everyone eligible, PRNs included, is already entered elsewhere or has no day target")
        if role in t.optional_roles:
            add("info", "optional role",
                f"{role} (optional): {n} slot(s) can't be filled ({why}).")
        else:
            add("warning", "unfillable",
                f"{role}: {n} required slot(s) can't be filled ({why}); they'll be left empty. "
                f"Days: {_fmt_dates(sorted(ds))}")

    # ---- locked cells that conflict with aptitude (locks win; reported for awareness)
    conflicts = defaultdict(Counter)   # role -> pid -> locked days
    for (pid, d), code in t.locked.items():
        s = t.staff.get(pid)
        if s and code in s.aptitude and s.aptitude[code] == 0:
            conflicts[code][pid] += 1
    for role in t.roles:
        if conflicts[role]:
            nobody = all(s.aptitude.get(role, 0) == 0 for s in t.staff.values())
            people = ", ".join(f"{t.label(p)} ({n}d)" for p, n in conflicts[role].items())
            add("info", "aptitude conflict",
                f"{role}: locked for pharmacists with aptitude 0: {people}. Kept as entered."
                + (" (Nobody has aptitude > 0 for this role.)" if nobody else ""))

    # ---- pay-period targets
    for per in t.periods():
        label = f"{_fmt(per['start'])}-{_fmt(per['end'])}"
        if not per["complete"]:
            add("info", "pay period",
                f"Pay period {label} is only partly inside the schedule "
                f"({len(per['dates'])} of {t.settings.pay_period_days} days). Targets can't be checked "
                f"for it unless the days outside the schedule are supplied.")
            continue
        for pid in t.schedule_order:
            s = t.staff.get(pid)
            if not s:
                continue
            paid = sum(1 for d in per["dates"] if (pid, d) in t.locked
                       and t.code_counts_for_pay(t.locked[(pid, d)]))
            if not s.target_days:
                continue  # target 0/blank: whatever is locked counts as satisfied
            target = s.target_days
            free = sum(1 for d in per["dates"] if (pid, d) not in t.locked)
            if paid > target:
                add("warning", "days per period",
                    f"{t.label(pid)}, {label}: {paid} locked days vs target {target}.")
            elif target - paid > free:
                add("warning", "days per period",
                    f"{t.label(pid)}, {label}: needs {target - paid} more days but only {free} unlocked days remain.")

    # ---- consecutive on-duty days within locked cells
    for pid in t.schedule_order:
        run, start = 0, None
        for d in t.dates + [None]:
            on = d is not None and (pid, d) in t.locked and t.code_on_duty(t.locked[(pid, d)])
            if on:
                run += 1
                start = start or d
            else:
                if run > t.settings.max_consecutive_hard:
                    add("warning", "consecutive days",
                        f"{t.label(pid)}: {run} consecutive locked days starting {_fmt(start)} "
                        f"(limit {t.settings.max_consecutive_hard}).")
                run, start = 0, None

    # ---- rest between shifts: filled-in pairs already under the minimum, and roles without times
    lo = t.settings.min_rest_hours * 60
    bad_rest = []
    for pid in t.schedule_order:
        for d in t.dates:
            a, b = t.locked.get((pid, d)), t.locked.get((pid, d + timedelta(days=1)))
            if a in t.shift_times and b in t.shift_times:
                gap = 1440 + t.shift_times[b][0] - t.shift_times[a][1]
                if gap <= lo:
                    bad_rest.append(f"{t.label(pid)} {a} {_fmt(d)} -> {b} {_fmt(d + timedelta(days=1))} "
                                    f"({gap / 60:g} h)")
    if bad_rest:
        add("warning", "rest",
            f"{len(bad_rest)} filled-in shift pair(s) leave {t.settings.min_rest_hours} hours or less "
            f"between shifts; kept as entered: " + "; ".join(bad_rest[:8])
            + (f" (+{len(bad_rest) - 8} more)" if len(bad_rest) > 8 else ""))
    no_times = [r for r in t.requirements if r not in t.shift_times]
    if no_times and t.shift_times:
        add("info", "rest", "No Start/End times for: " + ", ".join(no_times)
            + ". The rest-between-shifts rule can't check these roles.")
    elif not t.shift_times:
        add("info", "rest", "No shift times in the template, so rest between shifts isn't checked. "
                            "Add Start/End columns to the requirement matrix to turn it on.")

    # ---- shift balancing groups
    groups = balance_groups(t)
    if groups:
        add("info", "balance",
            f"Sharing shifts among aptitude-{t.settings.balance_aptitude} pharmacists "
            f"({t.settings.balance_mode.lower()}): "
            + "; ".join(f"{r}: {', '.join(t.label(p) for p in m)}" for r, m in groups.items()))

    # ---- capacity vs open slots, per complete pay period
    for per in t.periods():
        if not per["complete"]:
            continue
        label = f"{_fmt(per['start'])}-{_fmt(per['end'])}"
        pdates = set(per["dates"])
        need = sum(n for (role, d), n in slots.items()
                   if d in pdates and role not in t.optional_roles)
        remaining_total, usable = 0, 0
        for pid in t.schedule_order:
            if not schedulable(t, pid):
                continue
            s = t.staff[pid]
            paid = sum(1 for d in pdates if (pid, d) in t.locked
                       and t.code_counts_for_pay(t.locked[(pid, d)]))
            remaining = max(0, s.target_days - paid)
            useful_days = sum(1 for d in pdates if (pid, d) not in t.locked and any(
                s.aptitude.get(role, 0) > 0 for (role, dd) in slots if dd == d))
            remaining_total += remaining
            usable += min(remaining, useful_days)
        if usable < need:
            add("warning", "capacity",
                f"Pay period {label}: {need} required open slots but only about {usable} days "
                f"regular staff can still work. The rest will need PRNs where PRNs can work the "
                f"role, or be left empty (red in Shifts Fulfilled).")
        else:
            add("info", "capacity",
                f"Pay period {label}: {need} required open slots, {remaining_total} remaining "
                f"target days -> about {remaining_total - need} days for optional roles or project days.")

    return issues


def summary(t: Template) -> list[str]:
    slots = open_slots(t)
    kinds = Counter("role" if c in t.roles else "other codes" for c in t.locked.values())
    by_role = Counter()
    for (role, d), n in slots.items():
        by_role[role] += n
    targets = Counter(s.target_days for s in t.staff.values())
    wk = sum(v[0] for v in t.requirements.values())
    we = sum(v[5] for v in t.requirements.values())
    lines = [
        f"File:      {t.source.name}",
        f"Horizon:   {t.dates[0]:%a %m/%d/%Y} - {t.dates[-1]:%a %m/%d/%Y} ({len(t.dates)} days)",
        f"           {t.year_note}",
        "Pay periods:",
    ]
    for per in t.periods():
        flag = "complete" if per["complete"] else f"partial ({len(per['dates'])} days in schedule)"
        lines.append(f"           {per['start']:%m/%d}-{per['end']:%m/%d}  {flag}")
    lines += [
        f"Staff:     {len(t.schedule_order)} on schedule ({len(t.prn_pids)} PRN), {len(t.staff)} in aptitude matrix",
        "           days/period targets: " + ", ".join(
            f"{k if k is not None else 'blank'}: {v}" for k, v in sorted(targets.items(), key=lambda kv: (kv[0] is None, kv[0] or 0))),
        f"Roles:     {len(t.roles)}; daily headcount {wk} on Mondays, {we} on Saturdays; "
        f"optional: {', '.join(sorted(t.optional_roles)) or 'none'}",
        f"Settings:  from {t.settings_source}; pay periods start {t.settings.pay_period_anchor:%m/%d/%Y}, "
        f"{t.settings.pay_period_days} days",
        f"Locked:    {len(t.locked)} cells (" + ", ".join(f"{k}: {v}" for k, v in kinds.most_common()) + ")",
        f"Open:      {sum(n for (r, d), n in slots.items() if r not in t.optional_roles)} required slots, "
        f"{sum(n for (r, d), n in slots.items() if r in t.optional_roles)} optional",
    ]
    if by_role:
        lines.append("           " + ", ".join(f"{r} {n}" for r, n in sorted(by_role.items(), key=lambda kv: -kv[1])))
    return lines
