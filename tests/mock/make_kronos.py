"""Write the mock Kronos report files (P9): the "Employee Transactions and Totals" export as ImportKronosData.bas reads
it (briefs/v27-kronos.md section 1), one .xlsx per report, into tests/mock/kronos/ (gitignored: *.xlsx).

    py -3.14 tests/mock/make_kronos.py

Files (copy each into the Kronos mock folder on the morning the plan names it):
    00 cumulative 2026-01-11 to 2026-10-30.xlsx     Day 0  - the year to date through Fri Oct 30 (tests/mock/kronos_totals.csv)
    01 day 2026-10-31.xlsx .. 14 day 2026-11-13.xlsx    one one-day report per morning for the day before (kronos_days.csv, Report = day)
    06b day 2026-11-03 repulled.xlsx                 Nov 6  - the Nov 3 day re-pulled with the edit (Report = re); copy it with 05's file
    15 cumulative 2026-01-11 to 2026-11-14.xlsx     the close-out - the year to date through Sat Nov 14 (the totals + every day + Report = cum)
Layout: rows 1-5 the title block (Time Period in A2 with the value in D2, Query in A3 / D3, Executed on, Printed for),
row 6 the Kronos error line, then Transactions: (heading: Global ID, Employee in D, Date in F, Day, Pay Code in H,
Earnings Amount, Hours in J, Wages, Edited By in L, DataSource), Totals: (the SHIFTED heading: Global ID, Employee in D,
Pay Code in F, Earnings Amount, Hours in H, Wages) and Summary Totals:. Pay codes carry the "CA " prefix as Kronos
prints them. Names are "Last, First" as Kronos prints them (the import matches by Payroll # first).
"""
from __future__ import annotations

import csv
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).resolve().parent
OUT = HERE / "kronos"
YEAR_START = date(2026, 1, 11)


def kronos_name(nm: str) -> str:
    first, last = nm.split(" ", 1)
    return f"{last}, {first}"


def mdy(d: date) -> str:
    return f"{d.month:02d}/{d.day:02d}/{d.year}"


def read_rows():
    totals = []
    with (HERE / "kronos_totals.csv").open(newline="", encoding="ascii") as f:
        for r in csv.DictReader(f):
            totals.append((r["Payroll #"], r["Employee"], r["Pay Code"], float(r["Hours"])))
    days = {}
    with (HERE / "kronos_days.csv").open(newline="", encoding="ascii") as f:
        for r in csv.DictReader(f):
            days.setdefault((r["Report"], date.fromisoformat(r["Date"])), []).append((r["Payroll #"], r["Employee"], r["Pay Code"], float(r["Hours"])))
    return totals, days


def write_report(path: Path, d_from: date, d_to: date, totals: list, transactions: list, executed: datetime) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws["A1"] = "Employee Transactions and Totals"
    ws["A2"] = "Time Period :"
    ws["D2"] = f"{mdy(d_from)} - {mdy(d_to)}"
    ws["A3"] = "Query :"
    ws["D3"] = "Location - Welland Hourly"
    ws["A4"] = "Executed on :"
    ws["D4"] = executed.strftime("%m/%d/%Y %I:%M %p")
    ws["A5"] = "Printed for :"
    ws["D5"] = "Mock, Report"
    ws["C6"] = "[WFM-COMMON-1234] Failed to retrieve some data from the providers"
    r = 7
    ws.cell(r, 1, "Transactions:")
    r += 1
    for c, h in ((1, "Global ID"), (4, "Employee"), (6, "Date"), (7, "Day"), (8, "Pay Code"), (9, "Earnings Amount"), (10, "Hours"), (11, "Wages"), (12, "Edited By"), (13, "DataSource")):
        ws.cell(r, c, h)
    r += 1
    for payroll, nm, code, hours, d, editor in transactions:
        ws.cell(r, 1, payroll)
        ws.cell(r, 4, kronos_name(nm))
        ws.cell(r, 6, datetime(d.year, d.month, d.day))
        ws.cell(r, 7, d.strftime("%a"))
        ws.cell(r, 8, "CA " + code)
        ws.cell(r, 9, 0)
        ws.cell(r, 10, hours)
        ws.cell(r, 11, 0)
        ws.cell(r, 12, editor)
        ws.cell(r, 13, "Timecard Editor")
        r += 1
    ws.cell(r, 1, "Totals:")
    r += 1
    for c, h in ((1, "Global ID"), (4, "Employee"), (6, "Pay Code"), (7, "Earnings Amount"), (8, "Hours"), (9, "Wages")):
        ws.cell(r, c, h)
    r += 1
    summary = {}
    for payroll, nm, code, hours in sorted(totals, key=lambda x: (kronos_name(x[1]), x[2])):
        ws.cell(r, 1, payroll)
        ws.cell(r, 4, kronos_name(nm))
        ws.cell(r, 6, "CA " + code)
        ws.cell(r, 7, 0)
        ws.cell(r, 8, hours)
        ws.cell(r, 9, 0)
        summary[code] = summary.get(code, 0.0) + hours
        r += 1
    ws.cell(r, 1, "Summary Totals:")
    r += 1
    for c, h in ((1, "Pay Code"), (6, "Earnings Amount"), (7, "Hours"), (8, "Wages")):
        ws.cell(r, c, h)
    r += 1
    for code, hours in sorted(summary.items()):
        ws.cell(r, 1, "CA " + code)
        ws.cell(r, 6, 0)
        ws.cell(r, 7, hours)
        ws.cell(r, 8, 0)
        r += 1
    wb.save(path)


def merge(rows: list) -> list:
    """Sum the hours per (payroll, name, code) - a cumulative's Totals block has one row per person per code."""
    acc = {}
    for payroll, nm, code, hours in rows:
        acc[(payroll, nm, code)] = acc.get((payroll, nm, code), 0.0) + hours
    return [(p, n, c, h) for (p, n, c), h in acc.items() if abs(h) > 0.0001]


def main() -> int:
    OUT.mkdir(exist_ok=True)
    totals, days = read_rows()
    day0 = date(2026, 10, 30)
    write_report(OUT / f"00 cumulative {YEAR_START.isoformat()} to {day0.isoformat()}.xlsx", YEAR_START, day0, totals, [],
                 datetime(2026, 10, 31, 7, 5))
    k = 0
    one_days = sorted(d for (rep, d) in days if rep == "day")
    for d in one_days:
        k += 1
        write_report(OUT / f"{k:02d} day {d.isoformat()}.xlsx", d, d, days[("day", d)], [], datetime.combine(d, datetime.min.time()).replace(hour=7, minute=10) + __import__("datetime").timedelta(days=1))
    for (rep, d), rows in days.items():
        if rep == "re":
            # the re-pulled day: the Transactions block carries the edit (the old figure reversed, the new one added)
            old = {(p, n, c): h for p, n, c, h in days[("day", d)]}
            tx = []
            for p, n, c, h in rows:
                h0 = old.get((p, n, c), 0.0)
                if abs(h - h0) > 0.0001:
                    tx.append((p, n, c, -h0, d, "Mock, Editor"))
                    tx.append((p, n, c, h, d, "Mock, Editor"))
            write_report(OUT / f"06b day {d.isoformat()} repulled.xlsx", d, d, rows, tx, datetime(2026, 11, 6, 7, 12))
    # the close-out cumulative: the opening totals + every day (the re-pulled figure wins) + the cum-only rows
    final = list(totals)
    for d in one_days:
        src = days.get(("re", d)) or days[("day", d)]
        final.extend(src)
    for (rep, d), rows in days.items():
        if rep == "cum":
            final.extend(rows)
    end = date(2026, 11, 14)
    write_report(OUT / f"15 cumulative {YEAR_START.isoformat()} to {end.isoformat()}.xlsx", YEAR_START, end, merge(final), [],
                 datetime(2026, 11, 16, 7, 0))
    print(f"{len(list(OUT.glob('*.xlsx')))} report files in {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
