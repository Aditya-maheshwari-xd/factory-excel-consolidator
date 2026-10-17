# Factory Excel Consolidator

Every morning, a plant manager opens a pile of Excel sheets — one per supervisor, sometimes one per shift — adds them up by hand, and squints for mistakes. Each supervisor keeps their sheet a little differently, the same file gets forwarded twice on WhatsApp, and a single extra zero can make yesterday look like a record day.

This tool takes a folder of those sheets **as they actually are** and produces one summary the owner can trust, plus a list of exactly which cell in which file needs fixing.

```
python -m consolidator sample_data/daily_sheets --config plant_config.example.json
```

```
Read 6 file(s), 23 entries.
  skipped sheet: cnc_2026-10-14.xlsx / Instructions (no production table found)
  skipped sheet: cnc_2026-10-15.xlsx / Instructions (no production table found)
Issues: 2 error(s), 7 warning(s), 6 note(s)
Wrote reports/production_summary.xlsx
Wrote reports/production_summary.html
```

The Excel report has five tabs: **Daily Summary**, **By Shift**, **By Machine**, **Issues To Fix** (colour-coded, each pointing to file / sheet / row) and **All Entries (Clean)**. The HTML page shows the same thing for a phone.

## What it copes with

Real sheets don't follow a template. The reader handles:

| Mess in the sheet | How it's handled |
|---|---|
| Company banner and a "Date:" line above the real header | Scans the first rows and picks the one that looks most like a header |
| `Prod. Qty`, `Output (pcs)`, `Actual`, `OK Qty` … | Header aliases map them all to one column |
| `1,200`, `395 pcs`, `nil`, `-` | Parsed to numbers; `nil` / `-` mean zero |
| `approx 750` | **Not** guessed — flagged as unreadable and left out of totals |
| Dates as `14/10/2026`, `14-10-26`, Excel serials, or only in the file name | Always read **day first** (Indian format); impossible dates are rejected, not guessed |
| Shift written `1st`, `Day`, `II`, `Night`, or only as the sheet name `Shift B` | Normalised to A / B / C |
| Machine written `M-1`, `m01`, `Machine 1`, `MC 1` | Normalised to `M01` |
| Supervisor's own `TOTAL` row, blank spacer rows, an "Instructions" tab | Skipped (totals are recomputed) |

## What it checks

| Check | Severity | Why it matters |
|---|---|---|
| Produced more than the machine's capacity per shift | Error | Almost always a typo (extra zero). Left out of totals so one typo can't fake a record day |
| Rejected more than produced | Error | Impossible. Left out of totals |
| Production blank or not a number | Error | Left out of totals, so its target doesn't drag achievement down either |
| Same machine + shift + date entered twice with **different** numbers | Error | Uses the first, excludes the other, so totals aren't double-counted |
| Same entry with different numbers, one file named `…revised` / `corrected` / `final` | Warning | Uses the revised figure and says what it replaced |
| Exact copy of an entry (file forwarded twice) | Note | Counted once |
| A machine/shift with no entry at all | Warning | Machine idle, or nobody filled the sheet? Someone should know which |
| Below target (default < 85%) | Warning | Shows the supervisor's remark if they gave one |
| Rejection above limit (default 5%) | Warning | |
| Long downtime (default > 60 min) with no reason written | Warning | |
| Machine not in the plant's machine list | Warning | |

Thresholds, machines, capacities and shifts live in a small JSON file — see [`plant_config.example.json`](plant_config.example.json). Without a config the tool still consolidates; it just can't do the capacity and missing-entry checks.

## Design decisions, and why

- **Never silently turn bad data into a number.** `approx 750` could be read as 750, but if the summary is wrong the owner stops trusting it. Flag it, exclude it, and point to the cell.
- **Impossible figures stay out of totals.** In the sample data, one `11900` instead of `1190` would turn a 93% day into 250%. Errors are shown in the Issues tab; totals only use figures that make sense.
- **Point to the source, don't patch the summary.** Every issue lists file / sheet / Excel row number, so the supervisor fixes their own sheet and tomorrow's sheet is better.
- **Day-first dates only.** `03/04/2026` is 3 April. A tool that sometimes guesses month-first is worse than one that refuses.
- **Conflicts aren't averaged or summed.** When two files disagree and neither says it's a revision, a person has to decide. The tool keeps one, says which, and raises it as an error.

## Try it

```bash
pip install -r requirements.txt
python sample_data/generate.py          # writes messy sample sheets for a fictional plant
python -m consolidator sample_data/daily_sheets --config plant_config.example.json --out reports
python -m pytest -q                     # 49 tests
```

The sample data (two supervisors, two days) has these problems planted on purpose: a sheet forwarded twice, an extra-zero typo later fixed in a `_revised` file, rejected > produced, `approx 750`, 120 min downtime with no reason, one machine/shift never filled in, and a shift with 10.7% rejection. The end-to-end test checks every one of them is caught and that totals aren't inflated.

Exit code is `2` when there are errors, so this can run on a schedule and alert someone.

## Layout

```
consolidator/
  normalize.py   reading sheets: header detection, aliases, dates, shifts, quantities
  checks.py      row checks, duplicates/revisions/conflicts, missing entries
  report.py      Excel + HTML summary
  __main__.py    command line
sample_data/generate.py   messy sample sheets
tests/                    unit + end-to-end tests
```

## Next steps with a real plant

- Sit with the supervisors and collect their actual sheets — the alias list will grow on day one.
- Ask the owner which numbers they look at first; the summary should open with those.
- Watch a folder (or a WhatsApp export) and send the summary at 9 AM automatically.
- Track who fixes which issue, so recurring mistakes show up per sheet / per supervisor.


