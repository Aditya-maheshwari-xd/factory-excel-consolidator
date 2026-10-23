"""Turn messy, hand-maintained production sheets into clean rows.

Every supervisor keeps their sheet a little differently: different column
names, a title banner above the header, dates as text or as Excel serials,
quantities typed as "1,200 pcs", shifts called "A" in one plant and "Day" in
another. This module absorbs those differences so the checks downstream can
work with one shape of data.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

# Canonical column -> the different ways people actually write it.
# Keys are compared after lower-casing and stripping punctuation.
HEADER_ALIASES: dict[str, list[str]] = {
    "date": ["date", "dt", "prod date", "production date", "report date", "dated"],
    "shift": ["shift", "shift name", "shft", "pali"],
    "machine": ["machine", "machine no", "machine id", "m c", "mc", "mc no",
                "m c no", "press", "press no", "line", "line no", "station"],
    "operator": ["operator", "operator name", "worker", "name", "op name", "karigar"],
    "part": ["part", "part no", "part name", "item", "item code", "component",
             "product", "sku"],
    "target_qty": ["target", "target qty", "plan", "plan qty", "planned qty",
                   "planned", "tgt"],
    "produced_qty": ["produced", "produced qty", "production", "prod qty",
                     "output", "output pcs", "actual", "actual qty", "ok qty",
                     "qty produced", "total produced", "good qty"],
    "rejected_qty": ["rejected", "rejection", "rej", "rej qty", "rejected qty",
                     "ng qty", "scrap", "defects", "rejection qty"],
    "downtime_min": ["downtime", "downtime min", "downtime mins", "breakdown",
                     "breakdown min", "idle min", "stoppage min", "bd min"],
    "remarks": ["remarks", "remark", "comments", "comment", "notes", "reason"],
}

CANONICAL_COLUMNS = list(HEADER_ALIASES)


def _key(text: object) -> str:
    """Lower-case and strip punctuation/units so 'Qty. (Pcs)' ~ 'qty pcs'."""
    s = str(text).strip().lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


_ALIAS_LOOKUP = {_key(a): canon for canon, aliases in HEADER_ALIASES.items() for a in aliases}


def match_header(cell: object) -> str | None:
    """Map a header cell to a canonical column name, or None."""
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return None
    k = _key(cell)
    if not k:
        return None
    if k in _ALIAS_LOOKUP:
        return _ALIAS_LOOKUP[k]
    # Tolerate trailing units like "produced qty nos" or "downtime in min".
    for suffix in (" nos", " pcs", " in min", " in mins", " minutes", " min", " mins"):
        if k.endswith(suffix) and k[: -len(suffix)].strip() in _ALIAS_LOOKUP:
            return _ALIAS_LOOKUP[k[: -len(suffix)].strip()]
    return None


# --------------------------------------------------------------------------
# Value parsing
# --------------------------------------------------------------------------

ZERO_WORDS = {"-", "--", "nil", "none", "zero", "0"}
EMPTY_WORDS = {"", "na", "n a", "nan"}


def parse_qty(value: object) -> tuple[float | None, bool]:
    """Parse a quantity cell.

    Returns (number, ok). ok=False means the cell had something in it that we
    could not read as a number — the caller should flag it, not silently use 0.
    A blank cell returns (None, True): missing, but not garbage.
    """
    if value is None:
        return None, True
    if isinstance(value, bool):
        return None, False
    if isinstance(value, (int, float)):
        if isinstance(value, float) and pd.isna(value):
            return None, True
        return float(value), True
    s = str(value).strip().lower()
    if s in ZERO_WORDS:  # check before "empty": "-" has no letters but means zero
        return 0.0, True
    if _key(s) in EMPTY_WORDS:
        return None, True
    s = re.sub(r"(pcs|pc|nos|no|units|mins|min|kg)\.?$", "", s).strip()
    s = s.replace(",", "").replace(" ", "")
    try:
        return float(s), True
    except ValueError:
        return None, False


_DATE_TEXT = re.compile(r"(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})")
_DATE_ISO = re.compile(r"(\d{4})-(\d{1,2})-(\d{1,2})")


def parse_date(value: object) -> date | None:
    """Parse a date the way Indian shop floors write it: day first.

    Handles datetime objects, Excel serial numbers, '14/10/2026', '14-10-26',
    '14.10.2026' and ISO '2026-10-14'. Never guesses month-first.
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, pd.Timestamp):
        return None if pd.isna(value) else value.date()
    if isinstance(value, (int, float)):
        if pd.isna(value):
            return None
        # Excel serial date (1900 system). 40000 ~ 2009, 60000 ~ 2064.
        if 30000 < value < 80000:
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
        return None
    s = str(value)
    m = _DATE_ISO.search(s)
    if m:
        y, mo, d = map(int, m.groups())
        return _safe_date(y, mo, d)
    m = _DATE_TEXT.search(s)
    if m:
        d, mo, y = map(int, m.groups())
        if y < 100:
            y += 2000
        return _safe_date(y, mo, d)
    return None


def _safe_date(y: int, m: int, d: int) -> date | None:
    try:
        return date(y, m, d)
    except ValueError:
        return None


SHIFT_ALIASES = {
    "A": ["a", "1", "i", "first", "1st", "day", "morning", "general", "g", "shift a",
          "shift 1", "day shift", "morning shift"],
    "B": ["b", "2", "ii", "second", "2nd", "evening", "afternoon", "shift b",
          "shift 2", "evening shift"],
    "C": ["c", "3", "iii", "third", "3rd", "night", "shift c", "shift 3", "night shift"],
}
_SHIFT_LOOKUP = {_key(a): canon for canon, aliases in SHIFT_ALIASES.items() for a in aliases}


def normalize_shift(value: object) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        value = str(int(value))
    return _SHIFT_LOOKUP.get(_key(value))


def normalize_machine(value: object) -> str | None:
    """'M-1', 'm01', 'Machine 1', 'MC 1', 1  ->  'M01'."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return f"M{int(value):02d}"
    s = str(value).strip()
    m = re.search(r"(\d+)", s)
    if not m:
        return s.upper() or None
    return f"M{int(m.group(1)):02d}"


# --------------------------------------------------------------------------
# Reading a workbook
# --------------------------------------------------------------------------

@dataclass
class RawIssue:
    file: str
    sheet: str
    row: int | None  # 1-based Excel row number, as the supervisor sees it
    message: str


@dataclass
class ReadResult:
    rows: list[dict] = field(default_factory=list)
    issues: list[RawIssue] = field(default_factory=list)
    sheets_read: int = 0
    sheets_skipped: list[str] = field(default_factory=list)
    sheet_names_read: list[str] = field(default_factory=list)   # production sheets, by name


def find_header_row(df: pd.DataFrame, scan_rows: int = 15) -> tuple[int | None, dict[int, str]]:
    """Find the row that looks most like a header.

    Supervisors often put a company banner, a date line or a blank row above
    the real header, so we score the first few rows by how many cells match a
    known column name and take the best one (needs at least 3 matches and
    must include a quantity column).
    """
    best_row, best_map = None, {}
    for r in range(min(scan_rows, len(df))):
        mapping: dict[int, str] = {}
        for c, cell in enumerate(df.iloc[r].tolist()):
            canon = match_header(cell)
            if canon and canon not in mapping.values():
                mapping[c] = canon
        if len(mapping) > len(best_map) and "produced_qty" in mapping.values():
            best_row, best_map = r, mapping
    if len(best_map) < 3:
        return None, {}
    return best_row, best_map


def _date_from_banner(df: pd.DataFrame, header_row: int) -> date | None:
    for r in range(header_row):
        for cell in df.iloc[r].tolist():
            d = parse_date(cell) if not isinstance(cell, (int, float)) else None
            if d:
                return d
    return None


def read_workbook(path: str | Path) -> ReadResult:
    """Read every sheet of one workbook into canonical rows."""
    path = Path(path)
    result = ReadResult()
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
    except Exception as exc:  # corrupt file, password protected, not Excel...
        result.issues.append(RawIssue(path.name, "-", None, f"Could not open file: {exc}"))
        return result

    file_date = parse_date(path.stem)

    for sheet_name, df in sheets.items():
        df = df.dropna(how="all").reset_index(drop=False)  # keep original index for row numbers
        orig_index = df.pop("index")
        if df.empty:
            result.sheets_skipped.append(f"{path.name} / {sheet_name} (empty)")
            continue
        header_row, mapping = find_header_row(df)
        if header_row is None:
            result.sheets_skipped.append(f"{path.name} / {sheet_name} (no production table found)")
            continue
        result.sheets_read += 1
        result.sheet_names_read.append(str(sheet_name))

        banner_date = _date_from_banner(df, header_row)
        sheet_shift = normalize_shift(re.sub(r"(?i)shift", "", str(sheet_name)).strip()) \
            if re.search(r"(?i)shift|day|night", str(sheet_name)) else None

        for r in range(header_row + 1, len(df)):
            excel_row = int(orig_index.iloc[r]) + 1
            values = df.iloc[r].tolist()
            raw = {canon: values[c] for c, canon in mapping.items()}

            first = _key(next((v for v in values if v is not None and not
                               (isinstance(v, float) and pd.isna(v))), ""))
            if first.startswith("total") or first.startswith("grand total"):
                continue  # supervisor's own total row — we recompute totals ourselves

            row: dict = {"source_file": path.name, "source_sheet": str(sheet_name),
                         "source_row": excel_row}
            row["date"] = parse_date(raw.get("date")) or banner_date or file_date
            row["shift"] = normalize_shift(raw.get("shift")) or sheet_shift
            row["machine"] = normalize_machine(raw.get("machine"))
            row["operator"] = _clean_text(raw.get("operator"))
            row["part"] = _clean_text(raw.get("part"))
            row["remarks"] = _clean_text(raw.get("remarks"))

            unreadable = []
            for col in ("target_qty", "produced_qty", "rejected_qty", "downtime_min"):
                val, ok = parse_qty(raw.get(col))
                row[col] = val
                if not ok:
                    unreadable.append(f"{col.replace('_', ' ')} = {raw.get(col)!r}")
            if row["rejected_qty"] is None and "rejected_qty" in mapping.values():
                row["rejected_qty"] = 0.0 if raw.get("rejected_qty") is None else row["rejected_qty"]

            if row["machine"] is None and row["produced_qty"] is None:
                continue  # blank / spacer line

            # Kept on the row (not as a loose issue) so the checks can report it
            # with the date / shift / machine it belongs to.
            row["_unreadable"] = unreadable
            result.rows.append(row)
    return result


def _clean_text(value: object) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = re.sub(r"\s+", " ", str(value)).strip()
    return s or None
