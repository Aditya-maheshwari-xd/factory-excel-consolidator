"""Command line: python -m consolidator <folder> [--config plant.json] [--out reports/]"""

from __future__ import annotations

import argparse
import sys
from collections import Counter

from . import consolidate, load_config, write_reports


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Merge daily production Excel sheets into one checked summary.")
    p.add_argument("folder", help="Folder containing the day's .xlsx files")
    p.add_argument("--config", help="Plant config JSON (machines, shifts, thresholds)")
    p.add_argument("--out", default="reports", help="Where to write the summary (default: reports/)")
    args = p.parse_args(argv)

    config = load_config(args.config)
    result = consolidate(args.folder, config)
    if not result.files_read:
        print(f"No Excel files found in {args.folder}", file=sys.stderr)
        return 1

    xlsx, page = write_reports(result, args.out, config.get("plant_name", "Plant"))
    counts = Counter(i.severity for i in result.issues)
    print(f"Read {len(result.files_read)} file(s), {len(result.rows)} entries.")
    for s in result.sheets_skipped:
        print(f"  skipped sheet: {s}")
    print(f"Issues: {counts['error']} error(s), {counts['warning']} warning(s), {counts['info']} note(s)")
    print(f"Wrote {xlsx}\nWrote {page}")
    return 2 if counts["error"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
