"""Merge messy daily production sheets into one checked summary."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path

from .checks import Issue, run_all_checks
from .normalize import read_workbook
from .report import build_summaries, write_excel, write_html

__all__ = ["consolidate", "ConsolidationResult"]

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}


@dataclass
class ConsolidationResult:
    rows: list[dict]
    issues: list[Issue]
    summaries: dict
    files_read: list[str]
    sheets_skipped: list[str]


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
    for f in files:
        r = read_workbook(f)
        rows += r.rows
        raw_issues += r.issues
        skipped += r.sheets_skipped
    kept, issues = run_all_checks(rows, raw_issues, config)
    return ConsolidationResult(kept, issues, build_summaries(kept), [f.name for f in files], skipped)


def write_reports(result: ConsolidationResult, out_dir: str | Path, plant_name: str = "Plant") -> tuple[Path, Path]:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx = write_excel(out_dir / "production_summary.xlsx", result.summaries, result.issues)
    page = write_html(out_dir / "production_summary.html", result.summaries, result.issues,
                      plant_name, len(result.files_read))
    return xlsx, page
