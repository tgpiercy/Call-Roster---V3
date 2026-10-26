"""modMinStaff.StaffCountsCore (v27.1): the VBA core keeps the signature the mirror (rules.py, suite staffcount in
cases.csv) and SelfTest compose, StaffCounts calls it, and the shift length the mirror uses is the VBA's."""
from __future__ import annotations

import re
from pathlib import Path

import rules

ROOT = Path(__file__).resolve().parent.parent


def test_mirror_matches_vba():
    src = (ROOT / "modules" / "modMinStaff.bas").read_bytes().decode("ascii")
    assert re.search(r"^Public Sub StaffCountsCore\(ByRef people As Variant, ByRef rightTexts As Variant, _\r\n"
                     r"\s+ByRef scheduled As Long, ByRef off As Long, ByRef covered As Long, ByRef part As Boolean, _\r\n"
                     r"\s+Optional ByVal shiftHours As Double = USUAL_SHIFT_HOURS\)", src, re.M)
    assert ("StaffCountsCore people, texts, scheduled, off, covered, part, "
            "DefaultShiftHours(d, IsNightShift(shiftLbl), USUAL_SHIFT_HOURS)") in src
    # v27.5 (item 3): the shares summed once and rounded down; "; " between two covers of one line is a part break too
    assert "covered = CLng(Int(total + 0.000001))" in src
    assert 'Split(Replace(rt, "; ", " / "), " / ")' in src
    assert "CoveredCount" not in src
    oe = (ROOT / "modules" / "modOvertimeEntry.bas").read_bytes().decode("ascii")
    m = re.search(r"Public Const USUAL_SHIFT_HOURS As Double = ([\d.]+)", oe)
    assert m and float(m.group(1)) == rules.USUAL_SHIFT_HOURS
