"""Checks on stock registers, dispatch registers, and between sheets.

Same rules as the production checks: never silently fix a number, point to the
exact file / sheet / row, and keep a wrong figure out of the totals rather than
guess what it should have been.
"""

from __future__ import annotations

from collections import defaultdict
from datetime import date

import pandas as pd

from .checks import ERROR, INFO, SEVERITY_ORDER, WARNING, Issue
from .registers import code_key

EPS = 1e-6
DEFAULT_POD_DAYS = 7          # a truck without POD for more than a week gets flagged


def _issue(row: dict, severity: str, check: str, message: str, area: str) -> Issue:
    return Issue(severity, check, message, row.get("date"), None, None, row.get("source_file"),
                 row.get("source_sheet"), row.get("source_row"), area=area, item=row.get("item"))


def _n(x: float | None) -> str:
    return "0" if x is None else f"{x:,.3f}".rstrip("0").rstrip(".")


def _drop_exact_copies(rows: list[dict], key, area: str, check: str) -> tuple[list[dict], list[Issue]]:
    """The same line in a second file (a sheet forwarded twice) is counted once.
    Identical lines inside one file are kept: two equal issues on one day can be real."""
    seen: dict[tuple, dict] = {}
    kept, issues = [], []
    for row in rows:
        k = key(row)
        first = seen.get(k)
        if first is not None and first["source_file"] != row["source_file"]:
            issues.append(_issue(row, INFO, check, f"Exact copy of {first['source_file']} row "
                                                   f"{first['source_row']} — counted once", area))
            continue
        seen.setdefault(k, row)
        kept.append(row)
    return kept, issues


# --------------------------------------------------------------------------
# Stock register
# --------------------------------------------------------------------------

STOCK_COLUMNS = ["item", "item_name", "uom", "opening", "received", "issued", "closing",
                 "closing_as_written", "reorder_level", "status"]


def check_stock(rows: list[dict]) -> tuple[pd.DataFrame, list[Issue]]:
    """Replay each item's register line by line.

    Two running figures are kept: the true balance from the numbers (opening +
    received - issued) and the balance as the store keeper wrote it. A written
    balance that doesn't follow from the line above is flagged once, on the row
    where the mistake was made; later rows are checked against what was written,
    so one slip doesn't flag every row after it.
    """
    rows, issues = _drop_exact_copies(
        rows, lambda r: (r["item"], r["date"], r.get("ref"), r["received_qty"], r["issued_qty"],
                         r["balance_qty"]), "Stock", "stock_duplicate")
    by_item: dict[str, list[dict]] = defaultdict(list)
    for r in rows:
        if r["item"]:
            by_item[r["item"]].append(r)
        else:
            issues.append(_issue(r, ERROR, "stock_no_item", "Quantity entered with no item code", "Stock"))

    grn_seen: dict[str, dict] = {}
    summary = []
    for item, item_rows in by_item.items():
        item_rows.sort(key=lambda r: r["date"] or date.min)          # stable: file order within a day
        uom = next((r["uom"] for r in item_rows if r.get("uom")), "") or ""
        name = next((r["item_name"] for r in item_rows if r.get("item_name")), None)
        rol = next((r["reorder_level"] for r in item_rows if r.get("reorder_level") is not None), None)

        first = item_rows[0]
        if first["opening_qty"] is not None:
            opening, inferred = first["opening_qty"], False
        elif first["balance_qty"] is not None and not first["_unreadable"]:
            opening = first["balance_qty"] - (first["received_qty"] or 0) + (first["issued_qty"] or 0)
            inferred = True
        else:
            opening, inferred = 0.0, False

        calc = ref = opening
        received = issued = 0.0
        last_written = None
        for i, r in enumerate(item_rows):
            if r["_unreadable"]:
                issues.append(_issue(r, ERROR, "stock_unreadable",
                                     f"Not a number: {'; '.join(r['_unreadable'])} — left out of the balance "
                                     f"until corrected", "Stock"))
                continue
            rin, rout = r["received_qty"] or 0.0, r["issued_qty"] or 0.0
            received += rin
            issued += rout
            new_calc = calc + rin - rout
            if new_calc < -EPS <= calc:
                issues.append(_issue(r, ERROR, "stock_negative",
                                     f"Issued {_n(rout)} {uom} but only {_n(calc)} was in stock — a receipt "
                                     f"is missing or was entered late", "Stock"))
            expected = ref + rin - rout
            written = r["balance_qty"]
            if written is not None:
                if abs(written - expected) > 0.001 and not (i == 0 and inferred):
                    issues.append(_issue(r, ERROR, "stock_balance_mismatch",
                                         f"Balance written as {_n(written)}, but {_n(ref)} + {_n(rin)} received "
                                         f"− {_n(rout)} issued = {_n(expected)} (off by "
                                         f"{'+' if written > expected else '−'}{_n(abs(written - expected))})",
                                         "Stock"))
                ref = last_written = written
            else:
                ref = expected
            calc = new_calc

            ref_no = (r.get("ref") or "").upper().replace(" ", "")
            if rin > 0 and ref_no:
                prev = grn_seen.get(ref_no)
                if prev is not None:
                    issues.append(_issue(r, ERROR, "stock_duplicate_ref",
                                         f"{r['ref']} is entered twice (also {prev['source_file']} row "
                                         f"{prev['source_row']}). If it's the same delivery, stock is "
                                         f"overstated by {_n(rin)} {uom}", "Stock"))
                else:
                    grn_seen[ref_no] = r

        status = "negative" if calc < -EPS else "low" if rol and calc <= rol else "ok"
        if status == "low":
            issues.append(_issue(item_rows[-1], WARNING, "stock_low",
                                 f"{item} is at {_n(calc)} {uom}, at or below its reorder level of {_n(rol)}. "
                                 f"Raise a purchase order", "Stock"))
        summary.append({"item": item, "item_name": name, "uom": uom, "opening": opening,
                        "received": received, "issued": issued, "closing": round(calc, 3),
                        "closing_as_written": last_written, "reorder_level": rol, "status": status})
    return pd.DataFrame(summary, columns=STOCK_COLUMNS), issues


# --------------------------------------------------------------------------
# Dispatch register
# --------------------------------------------------------------------------

DISPATCH_COLUMNS = ["order_no", "customer", "item", "ordered", "dispatched", "pending", "due_date",
                    "status", "days_overdue", "trucks", "awaiting_pod"]


def check_dispatch(rows: list[dict], as_of: date, config: dict | None = None) -> tuple[pd.DataFrame, list[Issue]]:
    """Group trucks by order line (order no + item) and check them against the order."""
    cfg = (config or {}).get("dispatch", {})
    tol = cfg.get("tolerance_pct", 0) / 100
    pod_days = cfg.get("pod_days", DEFAULT_POD_DAYS)

    rows, issues = _drop_exact_copies(
        rows, lambda r: (r.get("invoice_no"), r["item"], r["dispatched_qty"], r.get("order_no"), r["date"]),
        "Dispatch", "dispatch_duplicate")

    by_invoice: dict[str, dict] = {}
    lines: dict[tuple, list[dict]] = defaultdict(list)
    for r in rows:
        if r["_unreadable"]:
            issues.append(_issue(r, ERROR, "dispatch_unreadable",
                                 f"Not a number: {'; '.join(r['_unreadable'])} — left out of totals until "
                                 f"corrected", "Dispatch"))
            continue
        inv = (r.get("invoice_no") or "").upper().replace(" ", "")
        if not inv:
            issues.append(_issue(r, WARNING, "dispatch_no_invoice", "Dispatch with no invoice number", "Dispatch"))
        else:
            prev = by_invoice.get(inv)
            if prev is not None and (prev.get("order_no") != r.get("order_no") or prev["item"] == r["item"]):
                where = f"{prev.get('order_no') or 'no order no'} / {prev['item']}"
                issues.append(_issue(r, ERROR, "dispatch_duplicate_invoice",
                                     f"Invoice {r['invoice_no']} is also used for {where} ({prev['source_file']} "
                                     f"row {prev['source_row']}). One invoice should be one truck for one order",
                                     "Dispatch"))
            by_invoice.setdefault(inv, r)
        lines[(r.get("order_no") or "(no order no)", r["item"])].append(r)

        if r["date"] and not r["pod_received"] and (as_of - r["date"]).days > pod_days:
            issues.append(_issue(r, WARNING, "dispatch_no_pod",
                                 f"No POD for invoice {r.get('invoice_no') or '?'} after "
                                 f"{(as_of - r['date']).days} days — the payment can't be chased", "Dispatch"))

    summary = []
    for (order_no, item), trucks in lines.items():
        trucks.sort(key=lambda r: r["date"] or date.min)
        ordered_values = [t["ordered_qty"] for t in trucks if t["ordered_qty"] is not None]
        ordered = ordered_values[-1] if ordered_values else None
        if len(set(ordered_values)) > 1:
            issues.append(_issue(trucks[-1], WARNING, "dispatch_order_qty_changed",
                                 f"Order quantity for {order_no} / {item} is written as "
                                 f"{' and '.join(_n(v) for v in dict.fromkeys(ordered_values))} on different "
                                 f"rows — using the latest ({_n(ordered)})", "Dispatch"))
        sent = sum(t["dispatched_qty"] or 0 for t in trucks)
        if ordered and sent > ordered * (1 + tol) + EPS:
            extra = f" (tolerance {tol:.0%})" if tol else ""
            issues.append(_issue(trucks[-1], ERROR, "dispatch_over",
                                 f"Sent {_n(sent)} against an order of {_n(ordered)}{extra} for {order_no} / "
                                 f"{item} — {_n(sent - ordered)} extra. A wrong entry, or does the customer "
                                 f"know?", "Dispatch"))
        pending = max(0.0, ordered - sent) if ordered else None
        due = next((t["due_date"] for t in reversed(trucks) if t["due_date"]), None)
        status = ("no order qty" if not ordered else "over-sent" if sent > ordered + EPS
                  else "complete" if sent >= ordered - EPS else "partial" if sent > 0 else "open")
        overdue = (as_of - due).days if due and pending and as_of > due else 0
        if overdue:
            issues.append(_issue(trucks[-1], WARNING, "dispatch_overdue",
                                 f"{order_no} / {item}: {_n(pending)} still to send, {overdue} day"
                                 f"{'s' if overdue != 1 else ''} past the due date ({due:%d %b})", "Dispatch"))
        summary.append({"order_no": order_no, "customer": trucks[0].get("customer"), "item": item,
                        "ordered": ordered, "dispatched": sent, "pending": pending, "due_date": due,
                        "status": status, "days_overdue": overdue, "trucks": len(trucks),
                        "awaiting_pod": sum(1 for t in trucks if not t["pod_received"])})
    return pd.DataFrame(summary, columns=DISPATCH_COLUMNS), issues


# --------------------------------------------------------------------------
# Across sheets
# --------------------------------------------------------------------------

def cross_checks(production_rows: list[dict], stock_rows: list[dict], dispatch_rows: list[dict],
                 stock_summary: pd.DataFrame) -> list[Issue]:
    """Problems no single sheet can show."""
    from .registers import normalize_code
    issues: list[Issue] = []
    if not dispatch_rows:
        return issues

    produced = defaultdict(float)
    prod_dates = []
    for r in production_rows:
        code = normalize_code(r.get("part"))
        if code and r.get("counted") and r.get("produced_qty") is not None:
            produced[code] += r["produced_qty"] - (r.get("rejected_qty") or 0)
            if r.get("date"):
                prod_dates.append(r["date"])
    stock_items = {r["item"] for r in stock_rows if r["item"]}
    known = set(produced) | stock_items

    # 1. A dispatched item code nobody else knows: usually a typo (BRK220 for BRK-220).
    # Only when something lists finished goods: production sheets, or a stock register that
    # shares items with dispatch. A raw-material register (sheet, nuts) can't vouch for products.
    dispatched_items = {r["item"] for r in dispatch_rows if r["item"]}
    if produced or (stock_items & dispatched_items):
        by_key = {code_key(k): k for k in known}
        flagged = set()
        for r in dispatch_rows:
            item = r["item"]
            if item and item not in known and item not in flagged:
                flagged.add(item)
                guess = by_key.get(code_key(item))
                hint = f" Did you mean {guess}?" if guess else ""
                issues.append(_issue(r, WARNING, "cross_unknown_item",
                                     f"{item} is dispatched but isn't in the production sheets or the stock "
                                     f"register.{hint}", "Cross-check"))

    # 2. More dispatched than made (on the days the production sheets cover).
    if prod_dates:
        start, end = min(prod_dates), max(prod_dates)
        opening = {}
        if not stock_summary.empty:
            opening = dict(zip(stock_summary["item"], stock_summary["opening"]))
        sent = defaultdict(float)
        last_row = {}
        for r in dispatch_rows:
            if r["item"] in produced and r["date"] and start <= r["date"] <= end and not r["_unreadable"]:
                sent[r["item"]] += r["dispatched_qty"] or 0
                last_row[r["item"]] = r
        days = f"{start:%d %b}" if start == end else f"{start:%d}–{end:%d %b}"
        for item, qty in sent.items():
            have = produced[item] + opening.get(item, 0)
            if qty > have + EPS:
                if item in opening:
                    msg = (f"Dispatched {_n(qty)} {item} on {days}, but only {_n(produced[item])} good pieces "
                           f"were made and {_n(opening[item])} were in stock. A dispatch or production entry "
                           f"is wrong")
                    sev = ERROR
                else:
                    msg = (f"Dispatched {_n(qty)} {item} on {days}, but the production sheets show only "
                           f"{_n(produced[item])} good pieces made on those days. Was the rest from earlier "
                           f"stock? Add finished goods to the stock register to check properly")
                    sev = WARNING
                issues.append(_issue(last_row[item], sev, "cross_dispatch_gt_production", msg, "Cross-check"))
    return issues


AREA_ORDER = {"Production": 0, "Stock": 1, "Dispatch": 2, "Cross-check": 3}


def sort_issues(issues: list[Issue]) -> list[Issue]:
    return sorted(issues, key=lambda i: (SEVERITY_ORDER[i.severity], AREA_ORDER.get(i.area, 9),
                                         i.date or date.min, i.shift or "", i.machine or "", i.item or ""))


def latest_date(*row_lists: list[dict]) -> date | None:
    dates = [r["date"] for rows in row_lists for r in rows if r.get("date")]
    return max(dates) if dates else None
