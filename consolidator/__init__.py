"""Merge messy daily factory sheets (production, stock, dispatch) into one checked summary."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pandas as pd

from .checks import Issue, run_all_checks
from .normalize import read_workbook
from .register_checks import check_dispatch, check_stock, cross_checks, latest_date, sort_issues
from .registers import read_registers
from .report import build_summaries, plant_now, write_excel, write_html

__all__ = ["consolidate", "ConsolidationResult"]

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}


@dataclass
class ConsolidationResult:
    rows: list[dict]
    issues: list[Issue]
    summaries: dict
    files_read: list[str]
    sheets_skipped: list[str]
    stock_rows: list[dict] = field(default_factory=list)
    dispatch_rows: list[dict] = field(default_factory=list)
    stock_summary: pd.DataFrame = field(default_factory=pd.DataFrame)
    dispatch_summary: pd.DataFrame = field(default_factory=pd.DataFrame)
    as_of: date | None = None

    @property
    def has_data(self) -> bool:
        return bool(self.rows or self.stock_rows or self.dispatch_rows)


def load_config(path: str | Path | None) -> dict:
    if not path:
        return {}
    return json.loads(Path(path).read_text(encoding="utf-8"))


def consolidate(folder: str | Path, config: dict | None = None) -> ConsolidationResult:
    config = config or {}
    folder = Path(folder)
    # Originals before WhatsApp-style copies ("report (1).xlsx"), so a forwarded
    # copy is the one reported as the duplicate, not the original.
    files = sorted((p for p in folder.iterdir()
                    if p.suffix.lower() in EXCEL_SUFFIXES and not p.name.startswith("~$")),
                   key=lambda p: (re.sub(r"\s*\(\d+\)$", "", p.stem), p.stem))
    rows, raw_issues, skipped = [], [], []
    stock_rows, dispatch_rows = [], []
    for f in files:
        r = read_workbook(f)
        rows += r.rows
        raw_issues += r.issues
        # Sheets that aren't production may be a stock or dispatch register.
        reg = read_registers(f, skip_sheets=set(r.sheet_names_read))
        stock_rows += reg.stock_rows
        dispatch_rows += reg.dispatch_rows
        recognised = tuple(f"{f.name} / {name} (" for name in reg.sheets_read)
        skipped += [s for s in r.sheets_skipped if not s.startswith(recognised)]

    kept, issues = run_all_checks(rows, raw_issues, config)
    # "Today" for overdue and POD checks = the latest date in the sheets, so an old
    # set of sheets gives the same answer whenever it is checked.
    as_of = latest_date(kept, stock_rows, dispatch_rows) or plant_now().date()
    stock_summary, stock_issues = check_stock(stock_rows)
    dispatch_summary, dispatch_issues = check_dispatch(dispatch_rows, as_of, config)
    cross = cross_checks(kept, stock_rows, dispatch_rows, stock_summary)
    issues = sort_issues(issues + stock_issues + dispatch_issues + cross)
    return ConsolidationResult(kept, issues, build_summaries(kept), [f.name for f in files], skipped,
                               stock_rows, dispatch_rows, stock_summary, dispatch_summary, as_of)


def write_reports(result: ConsolidationResult, out_dir: str | Path, plant_name: str = "Plant",
                  header_html: str = "") -> tuple[Path, Path]:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx = write_excel(out_dir / "production_summary.xlsx", result.summaries, result.issues,
                       result.stock_summary, result.dispatch_summary)
    page = write_html(out_dir / "production_summary.html", result.summaries, result.issues,
                      plant_name, len(result.files_read), header_html,
                      result.stock_summary, result.dispatch_summary, result.as_of)
    return xlsx, page
