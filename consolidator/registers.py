"""Stock registers and dispatch registers: recognise them and read them into rows.

A factory doesn't only send production sheets. The store keeps a stock register
(what came in, what went out, the balance) and dispatch keeps a register of
trucks against customer orders. Each is kept in Excel in its own way, so the
same idea as production applies: find the header row by known column names,
whatever the banner above it, and turn every line into a clean row that points
back to its file, sheet and Excel row.

Only sheets that are *not* production sheets are tried here, so a production
sheet is never read twice.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import pandas as pd

from .normalize import _clean_text, _date_from_banner, _key, parse_date, parse_qty

# Canonical column -> the different ways people write it (compared after _key()).
STOCK_ALIASES: dict[str, list[str]] = {
    "date": ["date", "dt", "txn date", "entry date", "dated"],
    "item": ["item code", "item", "material code", "material", "part no", "part code", "code",
             "item no", "mat code"],
    "item_name": ["item name", "description", "item description", "material description",
                  "material name", "particulars"],
    "uom": ["uom", "unit", "units"],
    "opening_qty": ["opening", "opening stock", "opening balance", "op bal", "op stock", "opening qty"],
    "received_qty": ["received", "receipt", "receipts", "in", "qty in", "inward", "received qty",
                     "receipt qty", "grn qty", "purchase"],
    "issued_qty": ["issued", "issue", "issues", "out", "qty out", "outward", "issued qty",
                   "consumption", "consumed", "issue qty"],
    "balance_qty": ["balance", "closing", "closing stock", "closing balance", "bal", "balance qty",
                    "cl bal", "stock in hand"],
    "ref": ["grn no", "grn", "grn number", "grn challan no", "challan no", "ref", "ref no",
            "voucher no", "doc no", "mrn no"],
    "reorder_level": ["reorder level", "rol", "min level", "minimum stock", "min stock", "reorder"],
    "remarks": ["remarks", "remark", "comments", "comment", "notes", "issued to"],
}

DISPATCH_ALIASES: dict[str, list[str]] = {
    "date": ["date", "dispatch date", "desp date", "invoice date", "dispatched on", "dt"],
    "order_date": ["order date", "po date"],
    "due_date": ["due date", "delivery date", "promised date", "required by", "schedule date", "due",
                 "del date"],
    "order_no": ["po no", "po", "po number", "order no", "order", "so no", "sales order", "customer po",
                 "order ref"],
    "customer": ["customer", "party", "party name", "customer name", "buyer", "consignee"],
    "item": ["item code", "item", "part no", "part", "product", "material code", "part code"],
    "ordered_qty": ["order qty", "ordered qty", "po qty", "ordered", "order quantity", "qty ordered"],
    "dispatched_qty": ["dispatch qty", "dispatched qty", "dispatched", "qty dispatched", "desp qty",
                       "sent qty", "qty sent", "invoice qty", "supplied qty"],
    "invoice_no": ["invoice no", "invoice", "inv no", "invoice number", "bill no"],
    "pod": ["pod", "pod date", "pod received", "delivered on", "delivery confirmed", "pod status"],
    "vehicle_no": ["vehicle no", "vehicle", "truck no", "lorry no"],
    "remarks": ["remarks", "remark", "comments", "comment", "notes"],
}


def _lookup(aliases: dict[str, list[str]]) -> dict[str, str]:
    return {_key(a): canon for canon, names in aliases.items() for a in names}


_STOCK_LOOKUP = _lookup(STOCK_ALIASES)
_DISPATCH_LOOKUP = _lookup(DISPATCH_ALIASES)
_UNIT_SUFFIXES = (" nos", " pcs", " kg", " kgs", " in kg", " in nos")


def _match(cell: object, lookup: dict[str, str]) -> str | None:
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return None
    k = _key(cell)
    if not k:
        return None
    if k in lookup:
        return lookup[k]
    for suffix in _UNIT_SUFFIXES:              # "Received (kg)", "Dispatch Qty (Nos)"
        if k.endswith(suffix) and k[: -len(suffix)].strip() in lookup:
            return lookup[k[: -len(suffix)].strip()]
    return None


def _stock_ok(cols: set[str]) -> bool:
    return "item" in cols and bool({"received_qty", "issued_qty"} & cols)


def _dispatch_ok(cols: set[str]) -> bool:
    return "dispatched_qty" in cols and bool({"order_no", "invoice_no"} & cols)


def find_register_header(df: pd.DataFrame, lookup: dict[str, str], required: Callable[[set[str]], bool],
                         scan_rows: int = 15) -> tuple[int | None, dict[int, str]]:
    """Same idea as production: score the first rows, keep the best one that has the
    columns this kind of register can't do without (and at least 3 known columns)."""
    best_row, best_map = None, {}
    for r in range(min(scan_rows, len(df))):
        mapping: dict[int, str] = {}
        for c, cell in enumerate(df.iloc[r].tolist()):
            canon = _match(cell, lookup)
            if canon and canon not in mapping.values():
                mapping[c] = canon
        if len(mapping) > len(best_map) and required(set(mapping.values())):
            best_row, best_map = r, mapping
    if len(best_map) < 3:
        return None, {}
    return best_row, best_map


# --------------------------------------------------------------------------
# Value helpers
# --------------------------------------------------------------------------

def normalize_code(value: object) -> str | None:
    """Item codes: upper case, no spaces. 'brk-220 ' -> 'BRK-220'. Hyphens are kept,
    so 'BRK220' stays different and can be flagged as a likely typo."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = re.sub(r"\s+", "", str(value)).upper()
    return s or None


def code_key(code: str | None) -> str:
    """Letters and digits only, to spot 'BRK220' vs 'BRK-220'."""
    return re.sub(r"[^A-Z0-9]", "", (code or "").upper())


POD_YES = {"yes", "y", "received", "recd", "done", "ok", "delivered", "rcvd"}


def parse_pod(value: object) -> tuple[bool, object]:
    """POD column: a date, or yes/received = delivered. Blank, 'no', 'pending' = not yet."""
    d = parse_date(value)
    if d:
        return True, d
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False, None
    return _key(value) in POD_YES, None


# --------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------

@dataclass
class RegisterRead:
    stock_rows: list[dict] = field(default_factory=list)
    dispatch_rows: list[dict] = field(default_factory=list)
    sheets_read: list[str] = field(default_factory=list)


STOCK_QTY = ("opening_qty", "received_qty", "issued_qty", "balance_qty", "reorder_level")
DISPATCH_QTY = ("ordered_qty", "dispatched_qty")


def read_registers(path: str | Path, skip_sheets: set[str] | None = None) -> RegisterRead:
    """Read every stock or dispatch register sheet in one workbook."""
    path = Path(path)
    out = RegisterRead()
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
    except Exception:
        return out          # the production reader has already reported a file it can't open
    file_date = parse_date(path.stem)

    for sheet_name, df in sheets.items():
        if skip_sheets and str(sheet_name) in skip_sheets:
            continue
        df = df.dropna(how="all").reset_index(drop=False)
        orig_index = df.pop("index")
        if df.empty:
            continue
        kinds = []
        for kind, lookup, required in (("stock", _STOCK_LOOKUP, _stock_ok),
                                       ("dispatch", _DISPATCH_LOOKUP, _dispatch_ok)):
            row, mapping = find_register_header(df, lookup, required)
            if row is not None:
                kinds.append((len(mapping), kind, row, mapping))
        if not kinds:
            continue
        _, kind, header_row, mapping = max(kinds)
        out.sheets_read.append(str(sheet_name))
        banner_date = _date_from_banner(df, header_row)

        for r in range(header_row + 1, len(df)):
            values = df.iloc[r].tolist()
            raw = {canon: values[c] for c, canon in mapping.items()}
            first = _key(next((v for v in values if v is not None and not
                               (isinstance(v, float) and pd.isna(v))), ""))
            if first.startswith("total") or first.startswith("grand total"):
                continue
            row: dict = {"source_file": path.name, "source_sheet": str(sheet_name),
                         "source_row": int(orig_index.iloc[r]) + 1,
                         "date": parse_date(raw.get("date")) or banner_date or file_date,
                         "item": normalize_code(raw.get("item")),
                         "remarks": _clean_text(raw.get("remarks"))}
            unreadable = []
            for col in (STOCK_QTY if kind == "stock" else DISPATCH_QTY):
                val, ok = parse_qty(raw.get(col))
                row[col] = val
                if not ok:
                    unreadable.append(f"{col.replace('_', ' ')} = {raw.get(col)!r}")
            row["_unreadable"] = unreadable

            if kind == "stock":
                if row["item"] is None and not any(row[c] is not None for c in STOCK_QTY):
                    continue    # blank / spacer line
                row["item_name"] = _clean_text(raw.get("item_name"))
                row["uom"] = _clean_text(raw.get("uom"))
                row["ref"] = _clean_text(raw.get("ref"))
                out.stock_rows.append(row)
            else:
                if row["item"] is None and row["dispatched_qty"] is None:
                    continue
                row["order_no"] = _clean_text(raw.get("order_no"))
                row["customer"] = _clean_text(raw.get("customer"))
                row["invoice_no"] = _clean_text(raw.get("invoice_no"))
                row["order_date"] = parse_date(raw.get("order_date"))
                row["due_date"] = parse_date(raw.get("due_date"))
                row["pod_received"], row["pod_date"] = parse_pod(raw.get("pod"))
                plate = _clean_text(raw.get("vehicle_no"))
                row["vehicle_no"] = plate.upper().replace(" ", "") if plate else None
                out.dispatch_rows.append(row)
    return out
