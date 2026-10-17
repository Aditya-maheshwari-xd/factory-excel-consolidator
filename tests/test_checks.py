from datetime import date

from consolidator.checks import check_row, find_duplicates, find_missing_entries, run_all_checks

CONFIG = {
    "shifts": ["A", "B"],
    "machines": {"M01": {"capacity_per_shift": 1500}, "M02": {"capacity_per_shift": 1500}},
}
D = date(2026, 10, 15)


def row(**kw):
    base = {"date": D, "shift": "A", "machine": "M01", "part": "P1", "target_qty": 1000.0,
            "produced_qty": 1000.0, "rejected_qty": 10.0, "downtime_min": 0.0, "remarks": None,
            "source_file": "f.xlsx", "source_sheet": "S", "source_row": 5}
    base.update(kw)
    return base


def checks(r):
    return {i.check for i in check_row(r, CONFIG)}


def test_clean_row_has_no_issues():
    assert checks(row()) == set()


def test_extra_zero_typo_is_caught():
    assert "over_capacity" in checks(row(produced_qty=11900.0))


def test_rejected_more_than_produced():
    assert "rejected_gt_produced" in checks(row(produced_qty=500.0, rejected_qty=600.0))


def test_high_rejection_and_below_target():
    found = checks(row(produced_qty=700.0, rejected_qty=70.0))
    assert {"high_rejection", "below_target"} <= found


def test_downtime_needs_a_reason():
    assert "downtime_no_reason" in checks(row(downtime_min=120.0))
    assert "downtime_no_reason" not in checks(row(downtime_min=120.0, remarks="Die change"))


def test_entry_that_cannot_be_placed():
    assert "cannot_place" in checks(row(machine=None))


def test_unknown_machine():
    assert "unknown_machine" in checks(row(machine="M09"))


def test_exact_duplicate_counted_once():
    kept, issues = find_duplicates([row(), row(source_file="f (1).xlsx")])
    assert len(kept) == 1
    assert [i.check for i in issues] == ["duplicate"]


def test_revised_file_wins_over_original():
    original = row(produced_qty=11900.0)
    revised = row(produced_qty=1190.0, source_file="report_revised.xlsx")
    kept, issues = find_duplicates([original, revised])
    assert [r["produced_qty"] for r in kept] == [1190.0]
    assert issues[0].check == "revised"


def test_unmarked_conflict_keeps_first_and_flags_error():
    kept, issues = find_duplicates([row(produced_qty=900.0), row(produced_qty=950.0, source_file="g.xlsx")])
    assert [r["produced_qty"] for r in kept] == [900.0]  # totals are not double-counted
    assert issues[0].check == "conflict" and issues[0].severity == "error"


def test_missing_machine_shift_is_flagged():
    rows = [row(), row(shift="B"), row(machine="M02")]
    missing = find_missing_entries(rows, CONFIG)
    assert [(i.shift, i.machine) for i in missing] == [("B", "M02")]


def test_impossible_rows_are_left_out_of_totals():
    rows = [row(), row(machine="M02", produced_qty=15000.0), row(shift="B", produced_qty=None)]
    kept, _ = run_all_checks(rows, [], CONFIG)
    assert [r["counted"] for r in kept] == [True, False, False]
