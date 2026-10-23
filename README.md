# Factory Excel Consolidator

Every morning, a plant manager opens a pile of Excel sheets — one per supervisor, sometimes one per shift, plus the store's stock register and the dispatch register — adds them up by hand, and squints for mistakes. Each person keeps their sheet a little differently, the same file gets forwarded twice on WhatsApp, and a single extra zero can make yesterday look like a record day.

This tool takes a folder of those sheets **as they actually are**, works out which sheet is production, stock or dispatch from its column names, and produces one summary the owner can trust, plus a list of exactly which cell in which file needs fixing. It also compares the sheets with each other, which is what the owner is really doing when they open them all.

```
python -m consolidator sample_data/daily_sheets --config plant_config.example.json
```

```
Read 8 file(s): 23 production entries, 17 stock register lines, 11 dispatch lines.
  skipped sheet: cnc_2026-10-14.xlsx / Instructions (no production table found)
  skipped sheet: cnc_2026-10-15.xlsx / Instructions (no production table found)
Issues: 8 error(s), 12 warning(s), 6 note(s)
Wrote reports/production_summary.xlsx
Wrote reports/production_summary.html
```

The Excel report has tabs for **Daily Summary**, **By Shift**, **By Machine**, **Issues To Fix** (colour-coded, each pointing to file / sheet / row), **All Entries (Clean)**, and, when those sheets are present, **Stock Summary** and **Dispatch Summary**. The HTML page shows the same thing for a phone.

**[See the sample report live](https://aditya-maheshwari-xd.github.io/factory-excel-consolidator/sample_report.html)**, built from the messy sheets in `sample_data/`.

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

### Stock register

| Check | Severity | Why it matters |
|---|---|---|
| Balance written doesn't follow from the line above (opening + received − issued) | Error | Flagged once, on the row where the slip happened; later rows are checked against what was written, so one mistake doesn't flag every row after it. The summary shows both the true closing and the closing as written |
| Stock goes below zero | Error | Usually material issued before its GRN was entered. Flagged once, where it went negative |
| Same GRN / challan number entered twice | Error | If it's the same delivery, stock is overstated |
| Quantity that isn't a number (`approx 30`) | Error | Left out of the balance, not guessed |
| At or below the reorder level (ROL column) | Warning | Raise a purchase order |
| Exact copy of a line in another file | Note | Counted once |

If there's no Opening column, the opening stock is worked out from the first balance written.

### Dispatch register

| Check | Severity | Why it matters |
|---|---|---|
| More sent than ordered (order no + item, all trucks together) | Error | Per-plant tolerance in the config (`dispatch.tolerance_pct`) |
| Invoice number used for a different order, or twice for the same item | Error | One invoice is one truck for one order. One invoice covering several items of one order is fine |
| Order not fully sent and past its due date | Warning | Shows what's pending and how many days late |
| No POD more than 7 days after dispatch (`dispatch.pod_days`) | Warning | Without the POD the payment can't be chased |
| Order quantity written differently on different rows | Warning | Uses the latest |
| Dispatch with no invoice number | Warning | |

"Today" for overdue and POD checks is the **latest date in the sheets**, so a set of sheets gives the same answer whenever it is checked.

### Across sheets

| Check | Severity | Why it matters |
|---|---|---|
| Dispatched item code not in the production sheets or stock register, with a guess (`BRK220` → "Did you mean BRK-220?") | Warning | Only when production sheets, or a stock register listing finished goods, are there to compare with: a raw-material register can't vouch for product codes |
| More of an item dispatched than was made on the days the production sheets cover | Warning, or Error if the stock register's opening stock still doesn't cover it | Either a wrong entry, or stock from earlier that isn't recorded |

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
python -m pytest -q                     # 81 tests
```

The sample data has these problems planted on purpose. **Production** (two supervisors, two days): a sheet forwarded twice, an extra-zero typo later fixed in a `_revised` file, rejected > produced, `approx 750`, 120 min downtime with no reason, one machine/shift never filled in, and a shift with 10.7% rejection. **Stock register:** a balance carried forward 100 kg too high, nuts issued before their receipt was entered, a GRN entered twice, `approx 30`, an item below its reorder level. **Dispatch register:** an invoice number reused, 1,100 sent against an order of 1,000, an overdue order, a truck with no POD after 10 days, an item code typed without its hyphen, and more of one item sent than was made. The end-to-end test checks every one of them is caught and that totals aren't inflated.

Exit code is `2` when there are errors, so this can run on a schedule and alert someone.

### In the browser: upload page

```bash
uvicorn consolidator.web:app --reload   # open http://127.0.0.1:8000
```

Upload the day's sheets (production, stock and dispatch together) and the page shows the same summary, with a link to download the Excel report. It runs exactly the same checks as the command line; `web.py` only handles the upload. A **Run on the sample sheets** button and download links for the eight sample sheets let anyone try it without their own data: download one, change a number in Excel, upload it again and see what gets flagged.

Uploads are limited to 20 files of 5 MB each (.xlsx or .xlsm) and are deleted after an hour. File names are kept, because the checks read words like `revised` and `(1)` from them.

### Deploy (free)

`render.yaml` sets the upload page up as a free web service on [Render](https://render.com): in the Render dashboard choose **New → Blueprint** and pick this repo. Free services sleep after 15 minutes without visitors and take about a minute to wake up.

## Layout

```
consolidator/
  normalize.py   reading production sheets: header detection, aliases, dates, shifts, quantities
  checks.py      production checks: row checks, duplicates/revisions/conflicts, missing entries
  registers.py   recognising and reading stock and dispatch registers
  register_checks.py   stock, dispatch and cross-sheet checks
  report.py      Excel + HTML summary
  __main__.py    command line
  web.py         upload page (FastAPI)
sample_data/generate.py   messy sample sheets
tests/                    unit + end-to-end tests
```

## Next steps with a real plant

- Sit with the supervisors and collect their actual sheets — the alias list will grow on day one.
- Ask the owner which numbers they look at first; the summary should open with those.
- Watch a folder (or a WhatsApp export) and send the summary at 9 AM automatically.
- Track who fixes which issue, so recurring mistakes show up per sheet / per supervisor.


