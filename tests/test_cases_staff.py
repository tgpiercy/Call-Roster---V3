"""Known-answer cases for suite "staffcount" (v27.1, R4): modMinStaff.StaffCountsCore and its mirror rules_staff.py.

Row format as cases.csv: suite,case,in1,in2,in3,in4,in5,expected
  in1 = the people tokens, ";" between them ("S" / "SA" / "SW"; anything else = not on the shift)
  in2 = the Covered-by texts, ";" between them
  in3 = the minimum
  expected = scheduled|off|covered|StaffShort ("ok" when it gives "")
Integration folds these rows into cases.csv (and so into SelfTest) at release prep.
"""
from __future__ import annotations

import csv
import re
from pathlib import Path

import pytest

import rules_staff

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
CASES = HERE / "cases-staff.csv"


def load_cases() -> list[dict]:
    with CASES.open(newline="") as f:
        rows = list(csv.DictReader(f))
    assert rows, "cases-staff.csv is empty"
    return rows


def run(row: dict) -> str:
    people = row["in1"].split(";") if row["in1"] != "" else []
    texts = row["in2"].split(";") if row["in2"] != "" else []
    s, o, c = rules_staff.staff_counts_core(people, texts)
    short = rules_staff.staff_short(s, o, c, int(row["in3"] or 0))
    return f"{s}|{o}|{c}|{short or 'ok'}"


@pytest.mark.parametrize("row", load_cases(), ids=lambda r: f"{r['suite']}-{r['case']}")
def test_case(row):
    assert row["suite"] == "staffcount"
    assert run(row) == row["expected"]


def test_rows_fit_selftest():
    raw = CASES.read_text(encoding="ascii").splitlines()
    assert raw[0] == "suite,case,in1,in2,in3,in4,in5,expected"
    names = set()
    for line in raw[1:]:
        assert line.count(",") == 7, line                    # no comma inside a field
        name = line.split(",")[1]
        assert name not in names, f"duplicate case {name}"
        names.add(name)
    assert 10 <= len(names) <= 20


def test_mirror_matches_vba():
    src = (ROOT / "modules" / "modMinStaff.bas").read_bytes().decode("ascii")
    assert re.search(r"^Public Sub StaffCountsCore\(ByRef people As Variant, ByRef rightTexts As Variant, _\r\n"
                     r"\s+ByRef scheduled As Long, ByRef off As Long, ByRef covered As Long\)", src, re.M)
    assert "StaffCountsCore people, texts, scheduled, off, covered" in src
    oe = (ROOT / "modules" / "modOvertimeEntry.bas").read_bytes().decode("ascii")
    m = re.search(r"Public Const USUAL_SHIFT_HOURS As Double = ([\d.]+)", oe)
    assert m and float(m.group(1)) == rules_staff.USUAL_SHIFT_HOURS
