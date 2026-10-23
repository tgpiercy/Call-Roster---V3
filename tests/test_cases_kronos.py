"""Known-answer tests for the Kronos ledger rule's pure cores (v27.1).

cases-kronos.csv has the same columns and conventions as cases.csv (suite,case,in1..in5,expected; no comma inside a
field; lists use ';'; dates yyyy-mm-dd; stamps yyyy-mm-dd-hh-nn as modSelfTest.ParseStampMin reads them; 'none' = a
date of 0). Integration folds the rows into cases.csv / modSelfTest; until then they run here against rules_kronos.py.

Suites:
  kledger  cum Through|none, roster year's first day, the ledger's days (;), the snapshot 'code=h;...', the ledger
           rows 'yyyy-mm-dd:code=h;...' (ONE person) -> 'run end|H' (run end yyyy-mm-dd or none; H = that person's
           Worked, '-' = no row in the pivot). Composed as PivotFromLedger composes LedgerRunCore, LedgerDayIn and
           LedgerPersonCore: the snapshot is read only when there is one this year, and a person is a row of the pivot
           only with a snapshot row or a day inside the run.
  kedit    the edit's date, the year's first day, the first one-day report|none, the day's newest re-import stamp|none,
           the edit row's When stamp (or 'notadate') -> counts / skip (EditCountsCore)
  kday     the ledger row's date (or a stamp, or 'notadate'), the report's day, the row's Employee, on the roster Y|N
           -> drop / keep (LedgerDayDrop)
  kname    the roster name the Payroll # gives, the one the name gives -> clash / number / name / none (RosterNameCore)
"""
from __future__ import annotations

import csv
import re
from datetime import date, datetime
from pathlib import Path

import pytest

import rules_kronos as rk

CASES = Path(__file__).with_name("cases-kronos.csv")
VBA = Path(__file__).resolve().parents[1] / "modules" / "ImportKronosData.bas"

# How each expected answer was derived from the VBA (ImportKronosData.bas, v27.1 commit 1). OT = CA OVERTIME x 1.75,
# LIEU ST x 1, DBLTIME not a roster code; rounding = RoundHalfUp once per person.
NOTES = {
    "kledger:unbroken-run": "run 11,12,13 after the Oct 10 Through -> ends Oct 13; (100+4+4+4) x 1.75 = 196",
    "kledger:days-in-any-order": "LedgerRunCore puts the serials in a dictionary: order does not matter -> as unbroken-run",
    "kledger:gap-of-one-day-stops-the-run": "Oct 12 missing: the Do While stops at Oct 11; the Oct 13 row is after the run; 104 x 1.75 = 182",
    "kledger:first-day-missing": "Oct 11 missing: the loop never runs, the run end stays the Through (Oct 10); snapshot only: 175",
    "kledger:quiet-day-keeps-the-run": "Oct 12 is in the day list (another person's row or the quiet marker) though this person has none: run to Oct 13; 108 x 1.75 = 189",
    "kledger:no-ledger-days": "empty serial list: run end = the Through; 8 x 1.75 = 14",
    "kledger:no-cumulative-yet": "Through 0: LedgerRunCore returns 0, PivotFromLedger reads neither table -> no row",
    "kledger:cumulative-from-last-year": "Dec 20 < the year's Dec 28: LedgerCumOf gives 0 (review H3) -> none, no row",
    "kledger:cumulative-on-the-years-first-day": "Through = the year's first day is not before it: kept; run to Dec 29; 12 x 1.75 = 21",
    "kledger:days-inside-the-totals-not-added": "Oct 9 / Oct 10 are not after the Through (LedgerDayIn d > cum): only Oct 11's 4 adds; 182",
    "kledger:only-in-the-ledger": "no snapshot row but a day inside the run: a row; 2 x 1.75 = 3.5 rounds half up to 4",
    "kledger:ledger-only-after-a-gap-gets-no-row": "Oct 11 missing: run end Oct 10; the person's rows are after it, so spec.Add never runs: no pivot row (SyncKronosData then takes the missing / zeroed path, not found)",
    "kledger:rounded-once-not-per-day": "0.2 + 0.2 = 0.4 x 1.75 = 0.7 -> 1 (rounded per day it would be 0 + 0)",
    "kledger:other-code-not-counted": "DBLTIME has no multiplier (-1): dropped; 10 x 1.75 + 3 x 1 = 20.5 -> 21",
    "kledger:reversal-in-a-day": "a negative day row adds as it is: 96 x 1.75 = 168",
    "kedit:inside-the-ledger": "Oct 5 >= the year, ledger from Oct 1 > 0, Oct 5 >= Oct 1, no re-import -> counts",
    "kedit:on-the-ledger-start-day": "d >= ledgerFrom is inclusive -> counts",
    "kedit:on-the-years-first-day": "d >= yearFrom is inclusive -> counts",
    "kedit:before-the-ledger-start": "Sep 20 < the first one-day report (Oct 1) -> skip",
    "kedit:no-ledger-yet": "ledgerFrom 0: 'ledgerFrom > 0' fails -> skip",
    "kedit:last-years-edit-after-rollover": "Dec 20 < the year's Dec 28 -> skip",
    "kedit:day-reimported-after-the-edit": "re-imported Oct 18 7:00 >= the edit's Oct 17 8:00: DayReimportedAfter -> skip",
    "kedit:day-reimported-the-same-minute": "the test is >=: the same moment counts as after -> skip",
    "kedit:day-reimported-before-the-edit": "re-imported Oct 16 < the edit's Oct 17 -> counts",
    "kedit:edit-when-not-a-date": "IsDate(When) False: DayReimportedAfter exits False even with a re-import -> counts",
    "kedit:pins-the-missed-morning": "PINS TODAY (fixed in v27.5): an edit dated at the year-to-date's Through (no earlier "
                                     "Transactions day) for a day whose one-day report arrives later as a FIRST import (no 'day' "
                                     "change row) still counts, beside that day's own ledger rows - Worked Periods counts it twice",
    "kday:same-day-roster-row-replaced": "Int(row date) = Int(d) and a roster person -> drop",
    "kday:same-day-quiet-marker-replaced": "a blank Employee on the day is the quiet-day marker -> drop",
    "kday:same-day-former-person-kept": "not on the roster (a Former person's row) -> keep",
    "kday:other-day-kept": "another day -> keep",
    "kday:other-day-quiet-marker-kept": "the marker of another day -> keep",
    "kday:row-date-not-a-date-kept": "IsDate False -> keep",
    "kday:row-date-with-a-time-same-day": "the days are compared by Int: 3:30 PM on the day -> drop",
    "kname:number-only": "byNum only -> number (RosterNameFor returns byNum)",
    "kname:name-only": "byName only -> name",
    "kname:both-agree-exactly": "StrComp = 0 -> not a clash; byNum <> '' -> number",
    "kname:both-agree-any-case": "vbTextCompare: 'Ann Lee' = 'ann lee' -> number (the number's spelling)",
    "kname:clash": "both found, they differ -> clash (nothing taken, both named)",
    "kname:neither": "both '' -> none (RosterNameFor returns '')",
}

# The VBA cores and their parameter names, as declared (a drift guard: the mirror's docstrings name these)
VBA_CORES = {
    "LedgerRunCore": ["cumThrough", "yearFrom", "daySerials"],
    "LedgerDayIn": ["d", "cumThrough", "runEnd"],
    "LedgerPersonCore": ["snapSpec", "daySpec"],
    "EditCountsCore": ["d", "yearFrom", "ledgerFrom", "reimportedAt", "editWhen"],
    "LedgerDayDrop": ["rowDate", "d", "emp", "onRoster"],
    "RosterNameCore": ["byNum", "byName"],
}


def d(s: str) -> date:
    return date(int(s[0:4]), int(s[5:7]), int(s[8:10]))


def d_or_none(s: str) -> date | None:
    return None if s == "none" else d(s)


def stamp(s: str) -> datetime:
    """yyyy-mm-dd-hh-nn (modSelfTest.ParseStampMin)."""
    return datetime(int(s[0:4]), int(s[5:7]), int(s[8:10]), int(s[11:13]), int(s[14:16]))


def date_value(s: str):
    """A cell's value: 'notadate' as text, a stamp as a date and time, else a date."""
    if s == "notadate":
        return s
    return stamp(s) if len(s) >= 16 else d(s)


def load_cases() -> list[dict]:
    with CASES.open(newline="") as f:
        rows = list(csv.DictReader(f))
    assert rows, "cases-kronos.csv is empty"
    return rows


def case_ledger(i1: str, i2: str, i3: str, i4: str, i5: str) -> str:
    """kledger, composed as PivotFromLedger composes the cores (see the module docstring)."""
    cum, year_from = d_or_none(i1), d(i2)
    serials = ";".join(str(rk.serial(d(x))) for x in i3.split(";") if x)
    run_end = rk.ledger_run(cum, year_from, serials)
    cum_adj = rk.ledger_cum_of(cum, year_from)
    snap = i4 if cum_adj is not None else ""            # the snapshot is read only when there is one this year
    pieces = []
    if cum_adj is not None:
        for item in i5.split(";"):
            if item:
                day_txt, piece = item.split(":", 1)
                if rk.ledger_day_in(d(day_txt), cum_adj, run_end):
                    pieces.append(piece)
    end_txt = run_end.isoformat() if run_end is not None else "none"
    if snap == "" and not pieces:
        return f"{end_txt}|-"                             # no snapshot row and no day inside the run: no pivot row
    h = rk.ledger_person(snap, ";".join(pieces))
    return f"{end_txt}|{int(h) if float(h).is_integer() else h}"


def evaluate(row: dict) -> str:
    suite = row["suite"]
    i1, i2, i3, i4, i5 = (row[k] for k in ("in1", "in2", "in3", "in4", "in5"))
    if suite == "kledger":
        return case_ledger(i1, i2, i3, i4, i5)
    if suite == "kedit":
        re_at = None if i4 == "none" else stamp(i4)
        return "counts" if rk.edit_counts(d(i1), d(i2), d_or_none(i3), re_at, date_value(i5)) else "skip"
    if suite == "kday":
        return "drop" if rk.ledger_day_drop(date_value(i1), d(i2), i3, i4 == "Y") else "keep"
    if suite == "kname":
        return rk.roster_name(i1, i2)
    raise ValueError(f"unknown suite {suite}")


@pytest.mark.parametrize("row", load_cases(), ids=lambda r: f"{r['suite']}:{r['case']}")
def test_case(row):
    assert evaluate(row) == row["expected"], f"{row['suite']} {row['case']}: {row}"


def test_every_case_has_a_note_and_names_are_unique():
    keys = [f"{r['suite']}:{r['case']}" for r in load_cases()]
    assert len(keys) == len(set(keys)), "duplicate case names"
    assert set(keys) == set(NOTES), f"cases without a note: {set(keys) - set(NOTES)}; notes without a case: {set(NOTES) - set(keys)}"


def test_no_comma_inside_a_field():
    with CASES.open(newline="") as f:
        for line in f.read().splitlines():
            assert line.count(",") == 7, line


def test_vba_cores_exist_with_these_parameters():
    src = re.sub(r"\s_[ \t]*\r?\n\s*", " ", VBA.read_text(encoding="ascii"))     # line continuations joined
    for name, params in VBA_CORES.items():
        m = re.search(r"^Public Function " + name + r"\((.*?)\) As ", src, re.M)
        assert m, f"{name} is not a Public Function in ImportKronosData.bas"
        got = [re.sub(r"^(ByVal|ByRef)\s+", "", p.strip()).split()[0] for p in m.group(1).split(",")]
        assert got == params, f"{name}: {got}"
