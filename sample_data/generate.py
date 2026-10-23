"""Generate realistic, deliberately messy sample sheets for a fictional plant.

Production: two supervisors, two days, two different sheet styles, and the kind
of mistakes that really happen: a typo with an extra zero, a sheet re-sent on
WhatsApp, a "revised" file that disagrees with the original, a machine that
nobody filled in, quantities typed as text.

Stores and dispatch: the store keeper's stock register and the dispatch
register for the same days, with their own typical slips: a balance carried
forward wrongly, stock issued before its receipt was entered, a GRN entered
twice, an invoice number reused, an over-dispatch, a missing POD, and an item
code typed without its hyphen.

Run:  python sample_data/generate.py
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).parent
OUT = HERE / "daily_sheets"


def press_shop(day: datetime, rows: list[list], filename: str) -> None:
    """Supervisor 1 (press shop): banner on top, date in the banner, '1st'/'2nd' shifts,
    machines written 'M-1', numbers with commas, a TOTAL row at the bottom."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Production"
    ws.append(["SHREE BALAJI AUTO COMPONENTS (fictional) - PRESS SHOP DAILY REPORT"])
    ws.append([f"Date: {day:%d/%m/%Y}", None, None, "Prepared by: Ramesh"])
    ws.append([])
    ws.append(["Sr No", "M/C No", "Shift", "Operator", "Part No", "Target", "Prod. Qty",
               "Rej. Qty", "Breakdown (min)", "Remarks"])
    for i, r in enumerate(rows, start=1):
        ws.append([i, *r])
    ws.append(["TOTAL", None, None, None, None, None, "=SUM(G5:G10)", "=SUM(H5:H10)"])
    wb.save(OUT / filename)


def cnc_shop(day: datetime, shifts: dict[str, list[list]], filename: str) -> None:
    """Supervisor 2 (CNC): one sheet per shift, no date column (date is only in the
    file name), 'pcs' typed after numbers, 'nil' for zero, plus an instructions sheet."""
    wb = Workbook()
    info = wb.active
    info.title = "Instructions"
    info.append(["Fill one sheet per shift. Send on WhatsApp group by 9 AM."])
    for name, rows in shifts.items():
        ws = wb.create_sheet(name)
        ws.append(["Machine", "Item Code", "Operator Name", "Plan", "Output", "Rejection",
                   "Downtime", "Comments"])
        for r in rows:
            ws.append(r)
    wb.save(OUT / filename)


def stock_register(filename: str) -> None:
    """Store keeper: one running register for all raw materials, a balance after every line,
    the opening stock only on each item's first line, and a reorder level (ROL) column."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Stock Register"
    ws.append(["SHREE BALAJI AUTO COMPONENTS (fictional) - STORES STOCK REGISTER, OCT 2026"])
    ws.append([])
    ws.append(["Date", "Item Code", "Item Description", "UOM", "Opening", "Received", "Issued",
               "Balance", "GRN / Challan No", "ROL", "Remarks"])
    rows = [
        # CRCA sheet: the 15/10 balance is carried forward 100 kg too high (2330, should be 2230)
        ["13/10/2026", "CRCA-12", "CRCA sheet 1.2 mm", "kg", 1200, None, None, 1200, None, 500, None],
        ["14/10/2026", "CRCA-12", "CRCA sheet 1.2 mm", "kg", None, None, 450, 750, None, None, "Press shop"],
        ["14/10/2026", "CRCA-12", "CRCA sheet 1.2 mm", "kg", None, 2000, None, 2750, "GRN-1041", None, None],
        ["15/10/2026", "CRCA-12", "CRCA sheet 1.2 mm", "kg", None, None, 520, 2330, None, None, "Press shop"],
        ["15/10/2026", "CRCA-12", "CRCA sheet 1.2 mm", "kg", None, None, 300, 2030, None, None, "Press shop"],
        # MS rod: runs below its reorder level
        ["13/10/2026", "MSROD-8", "MS round rod 8 mm", "kg", 150, None, None, 150, None, 100, None],
        ["14/10/2026", "MSROD-8", "MS round rod 8 mm", "kg", None, None, 60, 90, None, None, "CNC"],
        ["15/10/2026", "MSROD-8", "MS round rod 8 mm", "kg", None, None, 40, 50, None, None, "CNC"],
        # Nuts: issued on 14/10 before the receipt was entered (on 15/10), balance left blank
        ["13/10/2026", "NUT-M8", "Hex nut M8", "nos", 1500, None, None, 1500, None, 2000, None],
        ["14/10/2026", "NUT-M8", "Hex nut M8", "nos", None, None, 1800, None, None, None, "Assembly"],
        ["15/10/2026", "NUT-M8", "Hex nut M8", "nos", None, "5,000", None, 4700, "GRN-1043", None, None],
        # Powder coat: the same GRN entered again the next day
        ["13/10/2026", "PWD-BLK", "Powder coat, black", "kg", 60, None, None, 60, None, 25, None],
        ["14/10/2026", "PWD-BLK", "Powder coat, black", "kg", None, 40, None, 100, "GRN-1044", None, None],
        ["15/10/2026", "PWD-BLK", "Powder coat, black", "kg", None, 40, None, 140, "GRN-1044", None, None],
        ["15/10/2026", "PWD-BLK", "Powder coat, black", "kg", None, None, 30, 110, None, None, "Paint shop"],
        # Boxes: a quantity written as text
        ["13/10/2026", "BOX-50", "Corrugated box (50 pcs)", "nos", 120, None, None, 120, None, 40, None],
        ["15/10/2026", "BOX-50", "Corrugated box (50 pcs)", "nos", None, None, "approx 30", None, None, None,
         "Dispatch"],
    ]
    for r in rows:
        ws.append(r)
    wb.save(OUT / filename)


def dispatch_register(filename: str) -> None:
    """Dispatch in-charge: one line per invoice (truck), the order quantity repeated on each line,
    POD written as a date or 'Yes' when it comes back."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Dispatch Register"
    ws.append(["SHREE BALAJI AUTO COMPONENTS (fictional) - DISPATCH REGISTER"])
    ws.append([])
    ws.append(["Date", "Invoice No", "Customer", "PO No", "Item Code", "Order Qty", "Dispatch Qty",
               "Due Date", "Vehicle No", "POD", "Remarks"])
    rows = [
        ["05/10/2026", "INV-0388", "Hosur Precision", "HP/7781", "SFT-118", 1500, 1500, "08/10/2026",
         "KA 51 MN 2290", None, None],                                            # POD never came back
        ["08/10/2026", "INV-0395", "Pune Auto Assemblies", "PAA-4471", "HNG-310", 2000, 1200, "12/10/2026",
         "MH 12 QT 9081", "Yes", None],
        ["10/10/2026", "INV-0401", "Sundaram Motors", "SM/1187", "BRK-220", 5000, 3000, "20/10/2026",
         "TN 09 BX 4421", "12/10/2026", None],
        ["13/10/2026", "INV-0409", "Rajkot Engineering Works", "REW/88", "BRK-221", 2500, 1000, "11/10/2026",
         "GJ 03 KL 7710", None, "Balance pending"],                              # overdue
        ["14/10/2026", "INV-0412", "Sundaram Motors", "SM/1187", "BRK-220", 5000, 2000, "20/10/2026",
         "TN 09 BX 4421", None, None],
        ["14/10/2026", "INV-0413", "Pune Auto Assemblies", "PAA-4471", "HNG-310", 2000, 800, "12/10/2026",
         "MH 12 QT 9081", None, None],
        ["14/10/2026", "INV-0414", "Coimbatore Pumps", "CP-0091", "BSH-044", 1000, 600, "18/10/2026",
         "TN 37 CK 1182", None, None],
        ["15/10/2026", "INV-0414", "Hosur Precision", "HP/7790", "SFT-119", 300, 300, "16/10/2026",
         "KA 51 MN 2290", None, "Invoice no. reused"],                            # same invoice, other order
        ["15/10/2026", "INV-0418", "Hosur Precision", "HP/7795", "HNG-310", 1000, 1000, "20/10/2026",
         "KA 51 MN 2290", None, None],                                            # more HNG-310 than made
        ["15/10/2026", "INV-0419", "Coimbatore Pumps", "CP-0091", "BSH-044", 1000, 500, "18/10/2026",
         "TN 37 CK 1182", None, None],                                            # 1100 against 1000
        ["15/10/2026", "INV-0420", "Pune Auto Assemblies", "PAA-4520", "BRK220", 1200, 400, "25/10/2026",
         "MH 12 QT 9081", None, None],                                            # code typed without hyphen
    ]
    for r in rows:
        ws.append(r)
    wb.save(OUT / filename)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for old in OUT.glob("*.xlsx"):
        old.unlink()

    d1, d2 = datetime(2026, 10, 14), datetime(2026, 10, 15)

    # ---- Day 1: a mostly clean day ---------------------------------------
    day1_press = [
        ["M-1", "1st", "Suresh", "BRK-220", "1,200", 1180, 12, 15, None],
        ["M-2", "1st", "Arjun", "BRK-221", 1000, 960, 9, 0, None],
        ["M-3", "1st", "Imran", "HNG-310", 800, 790, 6, None, None],
        ["M-1", "2nd", "Prakash", "BRK-220", "1,200", 1150, 20, 20, None],
        ["M-2", "2nd", "Vijay", "BRK-221", 1000, 610, 8, 150, "Die change + power cut"],
        ["M-3", "2nd", "Salim", "HNG-310", 800, 805, 5, 0, None],
    ]
    press_shop(d1, day1_press, "press_shop_14-10-2026.xlsx")
    # Same file forwarded again on WhatsApp -> exact duplicate.
    press_shop(d1, day1_press, "press_shop_14-10-2026 (1).xlsx")

    cnc_shop(d1, {
        "Shift A": [
            ["M04", "SFT-118", "Deepak", 400, "395 pcs", "nil", 10, None],
            ["M05", "SFT-119", "Anil", 400, "402 pcs", 3, None, None],
            ["M06", "BSH-044", "Kiran", 600, 588, 4, 0, None],
        ],
        "Shift B": [
            ["M04", "SFT-118", "Rahul", 400, 380, 6, 20, None],
            ["M05", "SFT-119", "Mohan", 400, 290, 31, 45, None],   # high rejection
            ["M06", "BSH-044", "Ganesh", 600, 575, 2, 0, None],
        ],
    }, "cnc_2026-10-14.xlsx")

    # ---- Day 2: the kind of mistakes that slip through ---------------------
    day2_press = [
        ["M-1", "1st", "Suresh", "BRK-220", "1,200", 11900, 10, 0, None],     # extra zero
        ["M-2", "1st", "Arjun", "BRK-221", 1000, 640, 7, 0, None],           # below target, no reason
        ["M-3", "1st", "Imran", "HNG-310", 800, 780, 820, 0, None],          # rejected > produced
        ["M-1", "2nd", "Prakash", "BRK-220", "1,200", 1175, 14, 0, None],
        ["M-2", "2nd", "Vijay", "BRK-221", 1000, 990, 11, 120, None],        # downtime, no reason
        ["M-3", "2nd", "Salim", "HNG-310", 800, "approx 750", 5, 0, None],  # not a number
    ]
    press_shop(d2, day2_press, "press_shop_15-10-2026.xlsx")
    # A "revised" sheet sent later where only M-1 1st shift was corrected.
    revised = [list(r) for r in day2_press]
    revised[0][5] = 1190
    press_shop(d2, revised[:1], "press_shop_15-10-2026_revised.xlsx")

    cnc_shop(d2, {
        "Shift A": [
            ["M04", "SFT-118", "Deepak", 400, 398, "nil", 0, None],
            ["M05", "SFT-119", "Anil", 400, 405, 2, 0, None],
            ["M06", "BSH-044", "Kiran", 600, 600, 3, 0, None],
        ],
        "Shift B": [
            ["M04", "SFT-118", "Rahul", 400, 390, 4, 0, None],
            # M05 shift B not filled in at all
            ["M06", "BSH-044", "Ganesh", 600, 560, 5, 30, None],
        ],
    }, "cnc_2026-10-15.xlsx")

    stock_register("stores_stock_register_oct-2026.xlsx")
    dispatch_register("dispatch_register_oct-2026.xlsx")

    print(f"Wrote sample sheets to {OUT}")


if __name__ == "__main__":
    main()
