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
