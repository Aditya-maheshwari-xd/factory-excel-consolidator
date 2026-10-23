"""Run the whole pipeline on the generated sample sheets and check the story it tells."""

import subprocess
import sys
from collections import Counter
from pathlib import Path

import pytest

from consolidator import consolidate, load_config, write_reports

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="module")
def result():
    subprocess.run([sys.executable, str(ROOT / "sample_data" / "generate.py")], check=True,
                   capture_output=True)
    return consolidate(ROOT / "sample_data" / "daily_sheets", load_config(ROOT / "plant_config.example.json"))


def test_every_planted_problem_is_found(result):
    found = Counter(i.check for i in result.issues)
    assert found["duplicate"] == 6          # WhatsApp re-send of a whole sheet
    assert found["revised"] == 1            # corrected file replaced the typo
    assert found["rejected_gt_produced"] == 1
    assert found["unreadable"] == 1         # "approx 750"
    assert found["downtime_no_reason"] == 1
    assert found["missing_entry"] == 1      # M05 shift B, day 2
    assert found["high_rejection"] == 1
    assert found["over_capacity"] == 0      # the 11900 typo was superseded by the revised file


    # Stock register
    assert found["stock_balance_mismatch"] == 1   # CRCA balance carried forward 100 kg too high
    assert found["stock_negative"] == 1           # nuts issued before their receipt was entered
    assert found["stock_duplicate_ref"] == 1      # GRN-1044 entered twice
    assert found["stock_unreadable"] == 1         # "approx 30"
    assert found["stock_low"] == 1                # MS rod below reorder level
    # Dispatch register
    assert found["dispatch_over"] == 1            # 1100 sent against 1000
    assert found["dispatch_duplicate_invoice"] == 1
    assert found["dispatch_overdue"] == 1         # REW/88, 4 days late
    assert found["dispatch_no_pod"] == 1          # INV-0388, 10 days
    # Across sheets
    assert found["cross_unknown_item"] == 1       # BRK220 for BRK-220
    assert found["cross_dispatch_gt_production"] == 1   # 1800 HNG-310 sent, 1584 made
    assert Counter(i.severity for i in result.issues) == {"error": 8, "warning": 12, "info": 6}


def test_register_summaries(result):
    stock = result.stock_summary.set_index("item")
    assert stock.loc["CRCA-12", "closing"] == 1930 and stock.loc["CRCA-12", "closing_as_written"] == 2030
    assert stock.loc["MSROD-8", "status"] == "low"
    orders = result.dispatch_summary.set_index(["order_no", "item"])
    assert orders.loc[("REW/88", "BRK-221"), "pending"] == 1500
    assert orders.loc[("CP-0091", "BSH-044"), "status"] == "over-sent"
    assert str(result.as_of) == "2026-10-15"


def test_totals_are_not_inflated(result):
    daily = result.summaries["daily"].set_index("date")
    day2 = daily.iloc[1]
    # The 11900 typo, if counted, would push achievement far above 100%.
    assert day2["achievement_pct"] < 1.0
    assert day2["rejection_pct"] < 0.05     # the impossible 820-rejected row is excluded


def test_reports_are_written(result, tmp_path):
    xlsx, page = write_reports(result, tmp_path, "Test Plant")
    assert xlsx.exists() and page.exists()
    html = page.read_text()
    assert "Test Plant" in html and "Issues to fix" in html
    assert "Stock register" in html and "Dispatch register" in html
