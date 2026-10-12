"""Known-answer tests shared with the workbook's SelfTest.

cases.csv is the single source of expected answers. The same rows are
pasted into modSelfTest.bas by gen_selftest_cases.py, so the VBA and the
Python mirror are checked against identical values.

Row format: suite,case,in1,in2,in3,in4,in5,expected. No commas inside a
field (lists use '/', records use ';', fields inside a record use ':').
"""
from __future__ import annotations

import csv
from datetime import date
from pathlib import Path

import pytest

import rules

CASES = Path(__file__).with_name("cases.csv")
YEAR_FROM = date(2025, 12, 28)      # provcount suite: this roster's year
YEAR_TO = date(2026, 12, 26)


def d(s: str) -> date:
    return date(int(s[0:4]), int(s[5:7]), int(s[8:10]))


def iso(x: date) -> str:
    return x.isoformat()


def load_cases() -> list[dict]:
    with CASES.open(newline="") as f:
        rows = list(csv.DictReader(f))
    assert rows, "cases.csv is empty"
    return rows


def parse_people(spec: str) -> list[rules.Emp]:
    out = []
    for rec in spec.split(";"):
        name, grade, hrs, sen = rec.split(":")
        out.append(rules.Emp(name, float(grade), float(hrs), d(sen) if sen else None))
    return out


def parse_listed(spec: str) -> list[rules.Listed]:
    out = []
    for rec in spec.split(";"):
        rank, grade, signed, awarded = rec.split(":")
        out.append(rules.Listed(int(rank), float(grade), signed == "Y", awarded == "Y"))
    return out


def fmt_num(x: float) -> str:
    return str(int(x)) if float(x).is_integer() else repr(float(x))


def charged_text(lst: list[rules.Listed]) -> str:
    return ";".join(f"{p.rank}={fmt_num(p.charged)}" for p in sorted(lst, key=lambda p: p.rank))


def evaluate(row: dict) -> str:
    """One case -> the actual answer as text (same formats as modSelfTest)."""
    suite = row["suite"]
    i1, i2, i3, i4, i5 = (row[k] for k in ("in1", "in2", "in3", "in4", "in5"))
    if suite == "rotation":                      # date -> days crew / nights crew
        return f"{rules.days_crew(d(i1))}/{rules.nights_crew(d(i1))}"
    if suite == "crewshift":                     # code, date -> Days / Nights / off
        return rules.crew_shift_on(i1, d(i2)) or "off"
    if suite == "avail":                         # date, D|N -> crews eligible
        return rules.available_crew(d(i1), i2 == "N")
    if suite == "holiday":                       # name, year -> date
        return iso(rules.holiday_date(i1, int(i2)))
    if suite == "holoff":                        # date, fixture 'name|year|observed' -> off / on
        observed = {}
        if i2:
            nm, yy, obs = i2.split("|")
            observed[(int(yy), nm)] = d(obs)
        return "off" if rules.days_holiday_off(d(i1), observed) else "on"
    if suite == "dst":                           # year, fall|spring -> date
        return iso(rules.dst_fall_back(int(i1)) if i2 == "fall" else rules.dst_spring_forward(int(i1)))
    if suite == "dsthours":                      # date -> 13 / 11 / 12
        return fmt_num(rules.dst_night_hours(d(i1)))
    if suite == "callorder":                     # people -> names in order
        return "/".join(e.name for e in rules.call_order(parse_people(i1)))
    if suite == "callout":                       # outcome, grade, hours -> charged
        return fmt_num(rules.callout_charge(i1, float(i2), float(i3)))
    if suite == "posting":                       # mode, hours, list -> rank=charged;...
        lst = parse_listed(i3)
        hrs = float(i2)
        if i1 == "award":
            rules.compute_award(hrs, lst)
        elif i1 == "notakers":
            rules.compute_no_takers(hrs, lst)
        elif i1 == "outside":
            rules.compute_outside(hrs, lst)
        elif i1 == "cancel":
            rules.compute_cancel(lst)
        else:
            raise ValueError(i1)
        return charged_text(lst)
    if suite == "eligibility":                   # code, date, shift, fixture -> reason
        flags = {}
        if i4.startswith("off:"):
            flags["booked_off"] = i4[4:]
        elif i4 == "ss":
            flags["ss_working"] = True
        elif i4 == "ssoff":
            flags["ss_gave_up"] = True
        elif i4 == "ssrest":
            flags["ss_rest"] = True
        elif i4:
            raise ValueError(i4)
        return rules.not_free_reason(i1, d(i2), i3, **flags) or "free"
    if suite == "prov":                          # shift date, D|N, through -> covered / counts
        return "covered" if rules.provisional_covered(d(i1), i2 == "N", d(i3)) else "counts"
    if suite == "provcount":                     # outcome, status, shift date, D|N, through -> counts / no
        return "counts" if rules.provisional_counts(i1, i2, d(i3), i4 == "N", d(i5), YEAR_FROM, YEAR_TO) else "no"
    if suite == "sick":                          # sick date, D|N, candidate date, D|N -> blocked / free
        return "blocked" if rules.sick_blocks(d(i1), i2 == "N", d(i3), i4 == "N") else "free"
    if suite == "hoursasof":                     # hh 'through:hours;...', log 'date:outcome:status:hours:charged:D|N;...', as-of -> total|through|est
        hh = [(d(p.split(":")[0]), float(p.split(":")[1])) for p in i1.split(";") if p]
        lg = []
        for rec in (p for p in i2.split(";") if p):
            dt, outcome, status, hrs, chg, sh = rec.split(":")
            lg.append(dict(date=d(dt), outcome=outcome, status=status, hours=float(hrs), charged=float(chg), night=(sh == "N")))
        total, through, est = rules.hours_as_of(hh, lg, d(i3), YEAR_FROM, YEAR_TO)
        return f"{fmt_num(total)}|{iso(through) if through else 'none'}|{'yes' if est else 'no'}"
    if suite == "leavepeer":                     # grade, returner Y/N, overlapping leave Y/N, history before Y/N -> peer / excluded
        return "peer" if rules.peer_eligible(float(i1), i2 == "Y", i3 == "Y", i4 == "Y") else "excluded"
    if suite == "leaveavg":                      # 'grade:returner:onleave:hist:delta;...' -> average of the eligible
        deltas = []
        for rec in (p for p in i1.split(";") if p):
            g, r, o, h, delta = rec.split(":")
            if rules.peer_eligible(float(g), r == "Y", o == "Y", h == "Y"):
                deltas.append(float(delta))
        return fmt_num(rules.leave_average(deltas))
    if suite == "merge":                         # 'date;date;...' -> 'start..end;...'
        runs = rules.merge_runs([d(p) for p in i1.split(";") if p])
        return ";".join(f"{iso(a)}..{iso(b)}" for a, b in runs)
    raise ValueError(f"unknown suite {suite}")


@pytest.mark.parametrize("row", load_cases(), ids=lambda r: f"{r['suite']}:{r['case']}")
def test_case(row):
    assert evaluate(row) == row["expected"], f"{row['suite']} {row['case']}: {row}"


def test_known_facts():
    """A few anchors the brief and the plant schedule state outright."""
    assert rules.days_crew(date(2025, 12, 28)) == "B"
    assert rules.nights_crew(date(2025, 12, 28)) == "A"
    assert rules.dst_fall_back(2026) == date(2026, 10, 31)
    assert rules.dst_spring_forward(2027) == date(2027, 3, 13)
    assert rules.holiday_date("Thanksgiving", 2026) == date(2026, 10, 12)
    assert rules.callout_charge("Refused", 1, 13) == 22.75
