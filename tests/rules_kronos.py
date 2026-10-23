"""Python mirror of the Kronos ledger rule's pure cores (v27.1, ImportKronosData.bas).

Each function mirrors the VBA procedure named in its docstring; tests/cases-kronos.csv holds the known answers
(suites kledger, kedit, kday, kname) and tests/test_cases_kronos.py runs them. A VBA date of 0 ("none") is None here.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta

import rules

EPOCH = date(1899, 12, 30)          # VBA / Excel day 0


def serial(d: date) -> int:
    """A whole day's VBA serial (CLng(Int(CDbl(d))))."""
    if isinstance(d, datetime):
        d = d.date()
    return (d - EPOCH).days


def _ser(d: date | None) -> float:
    """A date as the number VBA compares (0 = none)."""
    if d is None:
        return 0.0
    if isinstance(d, datetime):
        return (d - datetime(1899, 12, 30)).total_seconds() / 86400.0
    return float(serial(d))


def ledger_cum_of(cum: date | None, year_from: date) -> date | None:
    """ImportKronosData.LedgerCumOf (Private): a Totals snapshot from before the roster year is none."""
    if cum is None or _ser(cum) < _ser(year_from):
        return None
    return cum


def ledger_run(cum: date | None, year_from: date, day_serials: str) -> date | None:
    """ImportKronosData.LedgerRunCore: the Totals' Through, then each next day the ledger holds along the unbroken run;
    only days after the Through and in the roster year count. None = no snapshot this year."""
    cum = ledger_cum_of(cum, year_from)
    out = cum
    days: set[int] = set()
    for part in day_serials.split(";"):
        if part.strip():
            k = round(rules._val(part))            # CLng(Val(...)): half to even, as Python's round
            if k > _ser(cum) and k >= _ser(year_from):
                days.add(k)
    if cum is not None:
        d = cum + timedelta(days=1)
        while serial(d) in days:
            out = d
            d = d + timedelta(days=1)
    return out


def ledger_day_in(d: date, cum: date | None, run_end: date | None) -> bool:
    """ImportKronosData.LedgerDayIn: after the Totals' Through and inside the run."""
    return _ser(d) > _ser(cum) and _ser(d) <= _ser(run_end)


def ledger_person(snap_spec: str, day_spec: str) -> float:
    """ImportKronosData.LedgerPersonCore: the snapshot's 'code=h;...' then the run's, weighted and rounded once
    (WeightedTotal), the codes not counted dropped."""
    s = snap_spec
    if day_spec != "":
        s = s + ("" if s == "" else ";") + day_spec
    w = rules.weighted_total(s)
    if "|" in w:
        w = w[:w.index("|")]
    return rules._val(w)


def day_reimported_after(reimported_at: datetime | None, when) -> bool:
    """ImportKronosData.DayReimportedAfter (Private): the day re-imported at or after the edit's When."""
    if reimported_at is None:
        return False
    if not isinstance(when, (date, datetime)):
        return False                                # IsDate False: never "after"
    return _ser(reimported_at) >= _ser(when)


def edit_counts(d: date, year_from: date, ledger_from: date | None, reimported_at: datetime | None, when) -> bool:
    """ImportKronosData.EditCountsCore: a timecard edit dated d counts in Worked Periods when it is this year, from the
    first one-day report on, and the day was not re-imported at or after the edit was logged."""
    return (_ser(d) >= _ser(year_from) and _ser(ledger_from) > 0 and _ser(d) >= _ser(ledger_from)
            and not day_reimported_after(reimported_at, when))


def ledger_day_drop(row_date, d: date, emp: str, on_roster: bool) -> bool:
    """ImportKronosData.LedgerDayDrop: a one-day report for d replaces a ledger row of that day that is a roster
    person's or the quiet-day marker (blank Employee)."""
    if not isinstance(row_date, (date, datetime)):
        return False
    if int(_ser(row_date) // 1) == int(_ser(d) // 1):
        return emp == "" or on_roster
    return False


def roster_name(by_num: str, by_name: str) -> str:
    """ImportKronosData.RosterNameCore: clash / number / name / none."""
    if by_num != "" and by_name != "" and by_num.lower() != by_name.lower():
        return "clash"
    if by_num != "":
        return "number"
    if by_name != "":
        return "name"
    return "none"
