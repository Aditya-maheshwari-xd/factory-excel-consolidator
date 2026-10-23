"""Stock and dispatch registers: recognising them, and every check on them."""

from collections import Counter
from datetime import date

import pytest
from openpyxl import Workbook

from consolidator import consolidate
from consolidator.register_checks import check_dispatch, check_stock, cross_checks
from consolidator.registers import code_key, normalize_code, parse_pod, read_registers


def save(path, header, rows, banner=True):
    wb = Workbook()
    ws = wb.active
    if banner:
        ws.append(["SOME COMPANY - REGISTER"])
        ws.append([])
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


def stock(item, date_, opening=None, rin=None, rout=None, bal=None, ref=None, rol=None, row=1,
          file="stock.xlsx", unreadable=()):
    return {"item": item, "date": date_, "opening_qty": opening, "received_qty": rin, "issued_qty": rout,
            "balance_qty": bal, "ref": ref, "reorder_level": rol, "uom": "kg", "item_name": None,
            "remarks": None, "source_file": file, "source_sheet": "S", "source_row": row,
            "_unreadable": list(unreadable)}


def disp(order, item, ordered, sent, d, inv, due=None, pod=False, row=1, file="disp.xlsx"):
    return {"order_no": order, "item": item, "ordered_qty": ordered, "dispatched_qty": sent, "date": d,
            "invoice_no": inv, "due_date": due, "pod_received": pod, "pod_date": None, "customer": "C",
            "source_file": file, "source_sheet": "S", "source_row": row, "_unreadable": []}


D13, D14, D15 = date(2026, 10, 13), date(2026, 10, 14), date(2026, 10, 15)


# ---- recognising sheets ---------------------------------------------------

def test_stock_register_is_recognised_under_a_banner(tmp_path):
    f = save(tmp_path / "store.xlsx", ["Dt.", "Material Code", "UOM", "Op. Bal", "Receipt Qty", "Issue Qty",
                                       "Cl. Bal", "GRN No"],
             [["14/10/2026", " crca-12 ", "kg", 100, "50 kg", None, 150, "GRN-1"]])
    r = read_registers(f)
    assert r.sheets_read == ["Sheet"] and not r.dispatch_rows
    row = r.stock_rows[0]
    assert row["item"] == "CRCA-12" and row["received_qty"] == 50 and row["date"] == D14
    assert row["source_row"] == 4                       # banner, blank, header, then this line


def test_dispatch_register_is_recognised(tmp_path):
    f = save(tmp_path / "desp.xlsx", ["Invoice Date", "Inv No", "Party Name", "PO No", "Part No",
                                      "PO Qty", "Desp. Qty", "POD"],
             [["15/10/2026", "INV-1", "Acme", "PO-9", "BRK-220", 100, "60 nos", "Yes"]])
    r = read_registers(f)
    row = r.dispatch_rows[0]
    assert row["dispatched_qty"] == 60 and row["order_no"] == "PO-9" and row["pod_received"] is True


def test_production_sheet_is_not_read_as_a_register(tmp_path):
    save(tmp_path / "prod.xlsx", ["Date", "Machine", "Shift", "Item Code", "Target", "Produced", "Rejected"],
         [["14/10/2026", "M-1", "A", "BRK-220", 100, 90, 1]], banner=False)
    result = consolidate(tmp_path)
    assert len(result.rows) == 1 and not result.stock_rows and not result.dispatch_rows


def test_helpers():
    assert normalize_code(" brk-220 ") == "BRK-220" and normalize_code(220.0) == "220"
    assert code_key("BRK220") == code_key("brk-220")
    assert parse_pod("12/10/2026") == (True, date(2026, 10, 12))
    assert parse_pod("pending") == (False, None) and parse_pod(None) == (False, None)


# ---- stock ----------------------------------------------------------------

def test_wrong_balance_is_flagged_once_where_it_happened():
    rows = [stock("A", D13, opening=100, bal=100, row=1), stock("A", D14, rout=20, bal=90, row=2),  # should be 80
            stock("A", D15, rout=10, bal=80, row=3)]                       # follows on from the wrong 90
    summary, issues = check_stock(rows)
    wrong = [i for i in issues if i.check == "stock_balance_mismatch"]
    assert [i.row for i in wrong] == [2] and "off by +10" in wrong[0].message
    s = summary.iloc[0]
    assert s["closing"] == 70 and s["closing_as_written"] == 80


def test_opening_is_inferred_from_first_balance():
    summary, issues = check_stock([stock("A", D14, rin=50, bal=150), stock("A", D15, rout=30, bal=120)])
    assert summary.iloc[0]["opening"] == 100 and not issues


def test_issue_before_receipt_goes_negative_once():
    rows = [stock("A", D13, opening=10), stock("A", D14, rout=30, row=2), stock("A", D14, rout=5, row=3),
            stock("A", D15, rin=100, row=4)]
    summary, issues = check_stock(rows)
    assert [i.row for i in issues if i.check == "stock_negative"] == [2]
    assert summary.iloc[0]["closing"] == 75 and summary.iloc[0]["status"] == "ok"


def test_same_grn_twice_and_reorder_level():
    rows = [stock("A", D13, opening=10, rol=50), stock("A", D14, rin=20, ref="GRN 7", row=2),
            stock("A", D15, rin=20, ref="grn7", row=3)]
    summary, issues = check_stock(rows)
    found = Counter(i.check for i in issues)
    assert found["stock_duplicate_ref"] == 1 and found["stock_low"] == 1
    assert summary.iloc[0]["status"] == "low"


def test_unreadable_quantity_is_left_out():
    rows = [stock("A", D13, opening=10), stock("A", D14, rout=None, row=2, unreadable=["issued qty = 'approx 3'"])]
    summary, issues = check_stock(rows)
    assert summary.iloc[0]["closing"] == 10 and issues[0].check == "stock_unreadable"


def test_forwarded_copy_of_register_counted_once():
    a = stock("A", D14, opening=10, rin=5, file="store.xlsx")
    b = {**a, "source_file": "store (1).xlsx"}
    summary, issues = check_stock([a, b])
    assert summary.iloc[0]["received"] == 5 and issues[0].check == "stock_duplicate"


# ---- dispatch -------------------------------------------------------------

def test_over_dispatch_respects_tolerance():
    rows = [disp("P1", "X", 1000, 600, D14, "I1", row=1), disp("P1", "X", 1000, 440, D15, "I2", row=2)]
    _, issues = check_dispatch(rows, D15)
    assert [i.check for i in issues] == ["dispatch_over"]
    _, issues = check_dispatch(rows, D15, {"dispatch": {"tolerance_pct": 5}})
    assert not issues


def test_invoice_reused_for_another_order_but_multi_item_invoice_is_fine():
    rows = [disp("P1", "X", 10, 10, D14, "I1"), disp("P1", "Y", 10, 10, D14, "I1"),   # one truck, two items
            disp("P2", "Z", 10, 10, D15, "I1", row=3)]
    _, issues = check_dispatch(rows, D15)
    assert [(i.check, i.row) for i in issues] == [("dispatch_duplicate_invoice", 3)]


def test_overdue_and_missing_pod():
    rows = [disp("P1", "X", 100, 40, date(2026, 10, 1), "I1", due=date(2026, 10, 10)),
            disp("P2", "Y", 100, 100, date(2026, 10, 1), "I2", pod=True)]
    summary, issues = check_dispatch(rows, D15)
    found = {i.check: i for i in issues}
    assert "5 days past the due date" in found["dispatch_overdue"].message
    assert "after 14 days" in found["dispatch_no_pod"].message
    assert list(summary["status"]) == ["partial", "complete"]


def test_changed_order_quantity_uses_latest():
    rows = [disp("P1", "X", 100, 40, D14, "I1"), disp("P1", "X", 80, 40, D15, "I2")]
    summary, issues = check_dispatch(rows, D15)
    assert summary.iloc[0]["ordered"] == 80 and summary.iloc[0]["status"] == "complete"
    assert issues[0].check == "dispatch_order_qty_changed"


# ---- across sheets --------------------------------------------------------

def prod(part, produced, rejected, d):
    return {"part": part, "produced_qty": produced, "rejected_qty": rejected, "date": d, "counted": True}


def test_unknown_item_code_suggests_the_real_one():
    issues = cross_checks([prod("BRK-220", 100, 0, D14)], [], [disp("P1", "BRK220", 10, 10, D14, "I1")],
                          check_stock([])[0])
    assert issues[0].check == "cross_unknown_item" and "Did you mean BRK-220?" in issues[0].message


def test_raw_material_register_alone_does_not_flag_products():
    # Found while testing: with only a raw-material stock register and a dispatch register,
    # every product was reported as unknown, because the register lists sheet and nuts, not products.
    raw = [stock("CRCA-12", D13, opening=100)]
    issues = cross_checks([], raw, [disp("P1", "BRK-220", 10, 10, D14, "I1")], check_stock(raw)[0])
    assert not issues
    # A finished-goods register that shares items with dispatch does count.
    fg = [stock("BRK-220", D13, opening=100)]
    issues = cross_checks([], fg, [disp("P1", "BRK-220", 10, 10, D14, "I1"),
                                   disp("P2", "BRK220", 10, 10, D14, "I2")], check_stock(fg)[0])
    assert [i.check for i in issues] == ["cross_unknown_item"] and "BRK-220?" in issues[0].message


def test_dispatched_more_than_made():
    production = [prod("X", 100, 10, D14)]                                # 90 good
    dispatch = [disp("P1", "X", 200, 120, D14, "I1")]
    issues = cross_checks(production, [], dispatch, check_stock([])[0])
    assert issues[0].severity == "warning" and "only 90 good pieces" in issues[0].message

    stock_rows = [stock("X", D13, opening=20)]                            # 90 made + 20 in stock < 120
    issues = cross_checks(production, stock_rows, dispatch, check_stock(stock_rows)[0])
    assert issues[0].severity == "error"
    stock_rows = [stock("X", D13, opening=40)]                            # 130 available: fine
    assert not cross_checks(production, stock_rows, dispatch, check_stock(stock_rows)[0])


def test_dispatches_outside_production_days_are_not_compared():
    production = [prod("X", 100, 0, D14)]
    dispatch = [disp("P1", "X", 500, 500, date(2026, 10, 1), "I1")]
    assert not cross_checks(production, [], dispatch, check_stock([])[0])


@pytest.mark.parametrize("missing", ["stock", "dispatch"])
def test_works_with_only_one_kind_of_register(tmp_path, missing):
    if missing != "stock":
        save(tmp_path / "s.xlsx", ["Item Code", "Received", "Issued", "Balance"], [["A", 10, None, 10]])
    if missing != "dispatch":
        save(tmp_path / "d.xlsx", ["PO No", "Item", "Order Qty", "Dispatch Qty"], [["P", "A", 5, 5]])
    result = consolidate(tmp_path)
    assert result.has_data and not result.rows
