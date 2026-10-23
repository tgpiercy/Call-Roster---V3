"""The Kronos ledger rule's pure cores (v27.1, ImportKronosData.bas): every kledger / kedit / kday / kname case in
cases.csv carries a note saying how its answer was read from the VBA, and the VBA cores keep the parameters the mirror
(rules.py) and SelfTest compose. The cases themselves run in test_cases.py with every other suite."""
from __future__ import annotations

import csv
import re
from pathlib import Path

CASES = Path(__file__).with_name("cases.csv")
VBA = Path(__file__).resolve().parents[1] / "modules" / "ImportKronosData.bas"
SUITES = ("kledger", "kedit", "kday", "kname")

# How each expected answer was derived from the VBA (ImportKronosData.bas, v27.1). OT = CA OVERTIME x 1.75,
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

# The VBA cores and their parameter names, as declared (a drift guard)
VBA_CORES = {
    "LedgerRunCore": ["cumThrough", "yearFrom", "daySerials"],
    "LedgerDayIn": ["d", "cumThrough", "runEnd"],
    "LedgerPersonCore": ["snapSpec", "daySpec"],
    "EditCountsCore": ["d", "yearFrom", "ledgerFrom", "reimportedAt", "editWhen"],
    "LedgerDayDrop": ["rowDate", "d", "emp", "onRoster"],
    "RosterNameCore": ["byNum", "byName"],
}


def test_every_kronos_case_has_a_note():
    with CASES.open(newline="") as f:
        keys = [f"{r['suite']}:{r['case']}" for r in csv.DictReader(f) if r["suite"] in SUITES]
    assert len(keys) == len(set(keys)), "duplicate case names"
    assert set(keys) == set(NOTES), f"cases without a note: {set(keys) - set(NOTES)}; notes without a case: {set(NOTES) - set(keys)}"


def test_vba_cores_exist_with_these_parameters():
    src = re.sub(r"\s_[ \t]*\r?\n\s*", " ", VBA.read_text(encoding="ascii"))     # line continuations joined
    for name, params in VBA_CORES.items():
        m = re.search(r"^Public Function " + name + r"\((.*?)\) As ", src, re.M)
        assert m, f"{name} is not a Public Function in ImportKronosData.bas"
        got = [re.sub(r"^(ByVal|ByRef)\s+", "", p.strip()).split()[0] for p in m.group(1).split(",")]
        assert got == params, f"{name}: {got}"
