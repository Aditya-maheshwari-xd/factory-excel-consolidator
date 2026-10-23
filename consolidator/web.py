"""A small web page in front of the consolidator: upload the day's sheets, get the report.

Production sheets, stock registers and dispatch registers can be uploaded together;
each sheet's type is worked out from its column names.

    uvicorn consolidator.web:app --reload        # open http://127.0.0.1:8000

The checks and the report are exactly the command-line tool's; this file only
handles the upload, runs `consolidate()` on a temporary folder, and serves the
result. Uploaded files and reports are kept for an hour, then deleted.
"""

from __future__ import annotations

import html
import json
import re
import shutil
import tempfile
import time
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

from . import consolidate, load_config, write_reports

ROOT = Path(__file__).resolve().parent.parent
SAMPLE_DIR = ROOT / "sample_data" / "daily_sheets"
SAMPLE_CONFIG = ROOT / "plant_config.example.json"
JOBS_DIR = Path(tempfile.gettempdir()) / "consolidator_jobs"

ALLOWED_SUFFIXES = {".xlsx", ".xlsm"}
MAX_FILES = 20
MAX_FILE_BYTES = 5 * 1024 * 1024   # 5 MB per sheet is far more than a daily report needs
KEEP_SECONDS = 60 * 60             # delete uploads and reports after an hour

app = FastAPI(title="Factory Excel Consolidator", version="1.0")

PAGE_STYLE = """
:root{--bg:#f6f7f9;--card:#fff;--ink:#1c2430;--muted:#5b6675;--line:#e3e6ea;--accent:#1f3a5f;--err:#a4262c}
@media (prefers-color-scheme:dark){:root{--bg:#12161c;--card:#1b212a;--ink:#e7ebf0;--muted:#9aa5b4;
--line:#2c3440;--accent:#8fb4e3;--err:#ff8a8f}}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.55 system-ui,sans-serif}
main{max-width:760px;margin:0 auto;padding:28px 16px}
h1{font-size:24px;margin:0 0 6px} h2{font-size:17px;margin:0 0 10px}
.muted{color:var(--muted)} .err{color:var(--err)}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:18px;margin-top:18px}
button,.btn{display:inline-block;background:var(--accent);color:var(--bg);border:0;border-radius:6px;
padding:9px 16px;font:inherit;cursor:pointer;text-decoration:none}
input[type=file]{display:block;margin:10px 0;max-width:100%}
label{display:block;margin:10px 0}
ul{padding-left:20px;margin:8px 0} a{color:var(--accent)}
"""

BAR_STYLE = ("display:flex;flex-wrap:wrap;gap:10px 18px;align-items:center;padding:12px 14px;"
             "margin-bottom:18px;border:1px solid var(--line);border-radius:8px;background:var(--card)")


def _page(title: str, body: str, status: int = 200) -> HTMLResponse:
    doc = (f"<!doctype html><html lang='en'><head><meta charset='utf-8'>"
           f"<meta name='viewport' content='width=device-width,initial-scale=1'>"
           f"<title>{html.escape(title)}</title><style>{PAGE_STYLE}</style></head>"
           f"<body><main>{body}</main></body></html>")
    return HTMLResponse(doc, status_code=status)


def _sample_files() -> list[str]:
    return sorted(p.name for p in SAMPLE_DIR.iterdir() if p.suffix.lower() in ALLOWED_SUFFIXES)


def _safe_name(name: str) -> str:
    """Keep the original file name (the checks read words like 'revised' and '(1)' from it),
    but drop any folder part and characters that don't belong in a file name."""
    name = Path(name or "").name
    name = re.sub(r"[^\w .()\-]", "_", name).strip(" .")
    return name or "upload.xlsx"


def _cleanup_old_jobs() -> None:
    if not JOBS_DIR.exists():
        return
    cutoff = time.time() - KEEP_SECONDS
    for job in JOBS_DIR.iterdir():
        if job.is_dir() and job.stat().st_mtime < cutoff:
            shutil.rmtree(job, ignore_errors=True)


def _new_job() -> tuple[str, Path]:
    _cleanup_old_jobs()
    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    (job_dir / "in").mkdir(parents=True)
    return job_id, job_dir


def _run(job_id: str, job_dir: Path, sheets_dir: Path, use_sample_config: bool, uploaded: bool) -> HTMLResponse:
    config = load_config(SAMPLE_CONFIG) if use_sample_config else {}
    result = consolidate(sheets_dir, config)
    if not result.has_data:
        skipped = "".join(f"<li>{html.escape(s)}</li>" for s in result.sheets_skipped)
        problems = "".join(f"<li>{html.escape(i.where())}: {html.escape(i.message)}</li>" for i in result.issues)
        return _page("No production data found", f"""
<h1>No factory data found</h1>
<p class="muted">Read {len(result.files_read)} file(s), but none had a table the tool recognises. It looks for a
header row with columns like these (many spellings are recognised):</p>
<ul class="muted"><li><b>Production:</b> Date, Shift, Machine, Produced, Rejected</li>
<li><b>Stock register:</b> Item Code, Received, Issued, Balance</li>
<li><b>Dispatch register:</b> PO No or Invoice No, Item, Order Qty, Dispatch Qty</li></ul>
{f'<div class="card"><h2>Sheets skipped</h2><ul>{skipped}</ul></div>' if skipped else ''}
{f'<div class="card"><h2>Files that could not be read</h2><ul>{problems}</ul></div>' if problems else ''}
<p><a class="btn" href="/">Try other files</a></p>""", status=422)

    files = ", ".join(html.escape(f) for f in result.files_read)
    skipped = (f"<span class='muted'>Skipped: {html.escape('; '.join(result.sheets_skipped))}</span>"
               if result.sheets_skipped else "")
    config_note = "sample plant's settings" if use_sample_config else "no machine list"
    found = ", ".join(f"{n} {what}" for n, what in ((len(result.rows), "production entries"),
                                                    (len(result.stock_rows), "stock lines"),
                                                    (len(result.dispatch_rows), "dispatch lines")) if n)
    bar = (f"<div style='{BAR_STYLE}'><a href='/'>&larr; Upload other sheets</a>"
           f"<a href='/download/{job_id}'><b>Download the Excel report</b></a>"
           f"<span class='muted'>{'Your' if uploaded else 'Sample'} files: {files} &middot; found {found} "
           f"&middot; checked with {config_note}</span>{skipped}</div>")
    plant = config.get("plant_name", "Uploaded sheets") if use_sample_config else "Uploaded sheets"
    _, page = write_reports(result, job_dir / "out", plant, header_html=bar)
    return HTMLResponse(page.read_text(encoding="utf-8"))


# ---- routes ----------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def home():
    samples = "".join(f"<li><a href='/samples/{html.escape(n)}'>{html.escape(n)}</a></li>" for n in _sample_files())
    return _page("Factory Excel Consolidator", f"""
<h1>Factory Excel Consolidator</h1>
<p class="muted">Upload the day's sheets together, as they are: supervisors' <b>production sheets</b>, the
store's <b>stock register</b> and the <b>dispatch register</b>. The tool works out which sheet is which, checks
each one, compares them with each other, and lists exactly which cell in which file needs fixing.</p>

<div class="card"><h2>Upload sheets</h2>
<form action="/consolidate" method="post" enctype="multipart/form-data">
<input type="file" name="files" accept=".xlsx,.xlsm" multiple required>
<label><input type="checkbox" name="use_sample_config" value="1" checked>
Use the sample plant's settings (machine capacities and shifts). Untick it for your own sheets:
the machine-capacity and missing-entry checks then switch off, everything else still runs.</label>
<button type="submit">Build the summary</button>
</form>
<p class="muted">Up to {MAX_FILES} files, {MAX_FILE_BYTES // (1024 * 1024)} MB each, .xlsx or .xlsm.
Files are deleted after an hour.</p></div>

<div class="card"><h2>No sheets handy? Use the sample data</h2>
<p>Two days of messy sheets from a fictional plant, with problems planted on purpose:</p>
<ul><li><b>Production (6 files):</b> a sheet forwarded twice, an extra-zero typo fixed later in a
<i>revised</i> file, rejected more than produced, "approx 750", downtime with no reason, a missing machine entry.</li>
<li><b>Stock register:</b> a balance carried forward wrongly, nuts issued before their receipt was entered,
a GRN entered twice, an item below its reorder level.</li>
<li><b>Dispatch register:</b> an invoice number reused, 1,100 sent against an order of 1,000, an overdue order,
a truck with no POD after 10 days, and an item code typed without its hyphen.</li></ul>
<p><a class="btn" href="/sample">Run on the sample sheets</a></p>
<p class="muted">Or download them, open them in Excel, change anything, and upload them above:</p>
<ul>{samples}</ul></div>

<p class="muted">Source code and how it works:
<a href="https://github.com/Aditya-maheshwari-xd/factory-excel-consolidator">GitHub</a></p>""")


@app.get("/sample", response_class=HTMLResponse)
def run_sample():
    job_id, job_dir = _new_job()
    for name in _sample_files():
        shutil.copy(SAMPLE_DIR / name, job_dir / "in" / name)
    return _run(job_id, job_dir, job_dir / "in", use_sample_config=True, uploaded=False)


@app.post("/consolidate", response_class=HTMLResponse)
async def run_upload(files: list[UploadFile] = File(...), use_sample_config: str | None = Form(None)):
    files = [f for f in files if f.filename]
    if not files:
        return _page("No files", "<h1>No files chosen</h1><p><a class='btn' href='/'>Back</a></p>", 400)
    if len(files) > MAX_FILES:
        return _page("Too many files", f"<h1>Too many files</h1><p>Upload up to {MAX_FILES} at a time.</p>"
                                       "<p><a class='btn' href='/'>Back</a></p>", 400)
    bad = [f.filename for f in files if Path(f.filename).suffix.lower() not in ALLOWED_SUFFIXES]
    if bad:
        names = ", ".join(html.escape(b) for b in bad)
        return _page("Not Excel files", f"<h1>Only Excel files, please</h1><p class='err'>Not .xlsx or .xlsm: "
                                        f"{names}</p><p>If it's an old .xls file, open it in Excel and "
                                        "save it as .xlsx.</p><p><a class='btn' href='/'>Back</a></p>", 400)

    job_id, job_dir = _new_job()
    used: set[str] = set()
    for f in files:
        data = await f.read(MAX_FILE_BYTES + 1)
        if len(data) > MAX_FILE_BYTES:
            shutil.rmtree(job_dir, ignore_errors=True)
            return _page("File too large", f"<h1>{html.escape(f.filename)} is larger than "
                                           f"{MAX_FILE_BYTES // (1024 * 1024)} MB</h1>"
                                           "<p><a class='btn' href='/'>Back</a></p>", 413)
        name = _safe_name(f.filename)
        stem, suffix, n = Path(name).stem, Path(name).suffix, 1
        while name.lower() in used:               # same name twice: keep both, like a downloaded copy
            name, n = f"{stem} ({n}){suffix}", n + 1
        used.add(name.lower())
        (job_dir / "in" / name).write_bytes(data)
    return _run(job_id, job_dir, job_dir / "in", use_sample_config=bool(use_sample_config), uploaded=True)


@app.get("/download/{job_id}")
def download(job_id: str):
    if not re.fullmatch(r"[0-9a-f]{32}", job_id):
        raise HTTPException(404, "Report not found")
    path = JOBS_DIR / job_id / "out" / "production_summary.xlsx"
    if not path.exists():
        raise HTTPException(404, "This report has expired (reports are kept for an hour). Build it again.")
    return FileResponse(path, filename="production_summary.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/samples/{name}")
def sample_file(name: str):
    if name not in _sample_files():                # only the bundled sample sheets, nothing else
        raise HTTPException(404, "No such sample file")
    return FileResponse(SAMPLE_DIR / name, filename=name)


@app.get("/health")
def health():
    return {"ok": True, "sample_files": len(_sample_files()),
            "sample_config": json.loads(SAMPLE_CONFIG.read_text(encoding="utf-8")).get("plant_name")}
