"""Checks a plant manager would otherwise do by eye, every morning.

Each check produces an Issue that points back to the exact file, sheet and
row, so the supervisor can go and fix the source sheet instead of someone
patching the summary by hand.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date

from .normalize import RawIssue

ERROR, WARNING, INFO = "error", "warning", "info"
SEVERITY_ORDER = {ERROR: 0, WARNING: 1, INFO: 2}
IMPOSSIBLE_CHECKS = {"over_capacity", "rejected_gt_produced"}


@dataclass
class Issue:
    severity: str
    check: str
    message: str
    date: date | None = None
    shift: str | None = None
    machine: str | None = None
    file: str | None = None
    sheet: str | None = None
    row: int | None = None

    def where(self) -> str:
        if not self.file:
            return "-"
        loc = f"{self.file} / {self.sheet}"
        return f"{loc} row {self.row}" if self.row else loc


DEFAULT_THRESHOLDS = {
    "min_target_achievement": 0.85,  # below 85% of target -> flag
    "max_rejection_rate": 0.05,      # above 5% rejection -> flag
    "max_downtime_min": 60,          # more than 60 min with no remark -> flag
}


def _issue(row: dict, severity: str, check: str, message: str) -> Issue:
    return Issue(severity, check, message, row.get("date"), row.get("shift"),
                 row.get("machine"), row.get("source_file"), row.get("source_sheet"),
                 row.get("source_row"))


def check_row(row: dict, config: dict) -> list[Issue]:
    """Checks that only need one row."""
    th = {**DEFAULT_THRESHOLDS, **config.get("thresholds", {})}
    machines = config.get("machines", {})
    out: list[Issue] = []

    missing = [k for k in ("date", "shift", "machine") if not row.get(k)]
    if missing:
        out.append(_issue(row, ERROR, "cannot_place",
                          f"Missing {', '.join(missing)} — this entry can't be placed in the summary"))

    produced, rejected = row.get("produced_qty"), row.get("rejected_qty")
    target, downtime = row.get("target_qty"), row.get("downtime_min")

    unreadable = row.get("_unreadable") or []
    for u in unreadable:
        out.append(_issue(row, ERROR, "unreadable",
                          f"Not a number: {u} — left out of totals until corrected"))

    if produced is None and not any(u.startswith("produced") for u in unreadable):
        out.append(_issue(row, ERROR, "blank_production", "Production quantity is blank"))

    for col, val in (("produced", produced), ("rejected", rejected),
                     ("target", target), ("downtime", downtime)):
        if val is not None and val < 0:
            out.append(_issue(row, ERROR, "negative", f"Negative {col} value ({val:g})"))

    machine = row.get("machine")
    if machine and machines and machine not in machines:
        out.append(_issue(row, WARNING, "unknown_machine",
                          f"Machine {machine} is not in the plant's machine list"))

    cap = machines.get(machine, {}).get("capacity_per_shift") if machine else None
    if cap and produced is not None and produced > cap:
        out.append(_issue(row, ERROR, "over_capacity",
                          f"Produced {produced:g} but {machine} can make at most {cap:g} per shift "
                          f"— likely a typo (extra zero?). Left out of totals"))

    if produced is not None and rejected is not None:
        if rejected > produced:
            out.append(_issue(row, ERROR, "rejected_gt_produced",
                              f"Rejected ({rejected:g}) is more than produced ({produced:g}). Left out of totals"))
        elif produced > 0 and rejected / produced > th["max_rejection_rate"]:
            out.append(_issue(row, WARNING, "high_rejection",
                              f"Rejection {rejected / produced:.1%} is above the "
                              f"{th['max_rejection_rate']:.0%} limit"))

    if target and produced is not None and produced < target * th["min_target_achievement"]:
        reason = f" (remark: {row['remarks']})" if row.get("remarks") else " — no remark given"
        out.append(_issue(row, WARNING, "below_target",
                          f"Produced {produced:g} vs target {target:g} "
                          f"({produced / target:.0%}){reason}"))

    if downtime and downtime > th["max_downtime_min"] and not row.get("remarks"):
        out.append(_issue(row, WARNING, "downtime_no_reason",
                          f"{downtime:g} min downtime with no reason written"))
    return out


REVISION_WORDS = re.compile(r"(?i)revised|corrected|updated|final|\bv2\b|\brev\b")


def _is_revision(row: dict) -> bool:
    return bool(REVISION_WORDS.search(row.get("source_file") or ""))


def _qty(v: float | None) -> str:
    return "blank" if v is None else f"{v:g}"


def find_duplicates(rows: list[dict]) -> tuple[list[dict], list[Issue]]:
    """Same date + shift + machine + part entered more than once.

    * Exact copies (same numbers) usually mean the sheet was forwarded twice on
      WhatsApp — count it once, note it.
    * Different numbers where one file is marked "revised"/"corrected"/"final"
      — the supervisor sent a correction, so use the revised figure and say so.
    * Different numbers with no revision marker — a real conflict. Counting
      both would inflate totals, so we keep the first one, leave the other out,
      and flag it as an error for someone to resolve.
    """
    chosen: dict[tuple, dict] = {}
    order: list[tuple | int] = []
    issues: list[Issue] = []
    for idx, row in enumerate(rows):
        key = (row.get("date"), row.get("shift"), row.get("machine"), row.get("part"))
        if None in key[:3]:
            order.append(idx)
            continue
        if key not in chosen:
            chosen[key] = row
            order.append(key)
            continue
        first = chosen[key]
        if all(first.get(c) == row.get(c) for c in ("produced_qty", "rejected_qty", "target_qty")):
            issues.append(_issue(row, INFO, "duplicate",
                                 f"Exact copy of {first['source_file']} row {first['source_row']} "
                                 f"— counted once"))
        elif _is_revision(row) and not _is_revision(first):
            chosen[key] = row
            issues.append(_issue(row, WARNING, "revised",
                                 f"Revised figure used: produced {_qty(row.get('produced_qty'))} replaces "
                                 f"{_qty(first.get('produced_qty'))} from {first['source_file']}"))
        elif _is_revision(first) and not _is_revision(row):
            issues.append(_issue(row, WARNING, "revised",
                                 f"Superseded by revised file {first['source_file']} "
                                 f"(produced {_qty(first.get('produced_qty'))} used instead of "
                                 f"{_qty(row.get('produced_qty'))})"))
        else:
            issues.append(_issue(row, ERROR, "conflict",
                                 f"Conflicts with {first['source_file']} row {first['source_row']} "
                                 f"(produced {_qty(first.get('produced_qty'))} vs "
                                 f"{_qty(row.get('produced_qty'))}). Using the first one until this is "
                                 f"resolved"))
    kept = [rows[k] if isinstance(k, int) else chosen[k] for k in order]
    return kept, issues


def find_missing_entries(rows: list[dict], config: dict) -> list[Issue]:
    """Every configured machine should report for every shift on every date seen."""
    machines = list(config.get("machines", {}))
    shifts = config.get("shifts", [])
    if not machines or not shifts:
        return []
    present = {(r["date"], r["shift"], r["machine"]) for r in rows
               if r.get("date") and r.get("shift") and r.get("machine")}
    dates = sorted({d for d, _, _ in present})
    out = []
    for d in dates:
        for s in shifts:
            for m in machines:
                if (d, s, m) not in present:
                    out.append(Issue(WARNING, "missing_entry",
                                     f"No entry for {m} in shift {s} — machine idle, or sheet not filled?",
                                     d, s, m))
    return out


def from_raw(raw: list[RawIssue]) -> list[Issue]:
    return [Issue(ERROR, "unreadable", r.message, file=r.file, sheet=r.sheet, row=r.row) for r in raw]


def run_all_checks(rows: list[dict], raw_issues: list[RawIssue], config: dict) -> tuple[list[dict], list[Issue]]:
    kept, issues = find_duplicates(rows)
    issues += from_raw(raw_issues)
    for row in kept:
        row_issues = check_row(row, config)
        issues += row_issues
        # Impossible figures would wreck the totals (one extra zero turns a normal
        # day into 200% achievement), so they are reported but left out of totals.
        # Rows with no usable production figure are left out too, otherwise their
        # target would count against achievement with nothing produced against it.
        impossible = any(i.check in IMPOSSIBLE_CHECKS for i in row_issues)
        row["counted"] = not impossible and row.get("produced_qty") is not None
    issues += find_missing_entries(kept, config)
    issues.sort(key=lambda i: (SEVERITY_ORDER[i.severity], i.date or date.min,
                               i.shift or "", i.machine or ""))
    return kept, issues
