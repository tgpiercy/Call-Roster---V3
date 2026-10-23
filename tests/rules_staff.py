"""Python mirror of modMinStaff.StaffCountsCore (v27.1, R4) - the staffing counts behind one unit / date / shift.

Kept beside rules.py until integration folds it in (suite "staffcount", cases in cases-staff.csv).

    staff_counts_core(people, right_texts) -> (scheduled, off, covered)

people      one token per code-holder: "S" on the shift by rotation, "SA" also booked off or sick for it,
            "SW" also gave it up in a switch; anything without "S" is not on the shift. (InStr, case-sensitive.)
right_texts the Covered-by texts of the unit's calendar lines on that shift.

The known edges stay as the VBA has them (Graeme: fixed in v27.5, not here): two partial lines round up
separately; an empty " / " part counts as one person.
"""
from __future__ import annotations

import math
import re

USUAL_SHIFT_HOURS = 12.0          # modOvertimeEntry.USUAL_SHIFT_HOURS

_NUMERIC = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?$")
_VAL = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?")


def vb_trim(s: str) -> str:
    """Trim$: spaces only."""
    return s.strip(" ")


def vb_isnumeric(s: str) -> bool:
    """IsNumeric for the text CoveredCount gives it: a run with no spaces between the last space and the final h.
    Plain decimal numbers only (hex, currency and thousands forms are out of scope; the cases avoid them)."""
    return bool(_NUMERIC.match(s))


def vb_val(s: str) -> float:
    """Val on a text with no spaces: the leading number, 0 when none."""
    m = _VAL.match(s)
    return float(m.group(0)) if m else 0.0


def covered_count(rt: str) -> int:
    """modMinStaff.CoveredCount: one person per " / " part; a part ending in "Nh" with 0 < N < 12 counts N / 12 of a
    person; the sum rounded up, at least 1."""
    total = 0.0
    for raw in rt.split(" / "):
        part = vb_trim(raw)
        hrs = 0.0
        if part[-1:].lower() == "h":
            p = part.rfind(" ") + 1                         # InStrRev, 1-based (0 = no space)
            if p > 0 and vb_isnumeric(part[p:len(part) - 1]):   # Mid$(part, p + 1, Len(part) - p - 1)
                hrs = vb_val(part[p:])                          # Val(Mid$(part, p + 1))
        if 0 < hrs < USUAL_SHIFT_HOURS:
            total += hrs / USUAL_SHIFT_HOURS
        else:
            total += 1
    return max(1, math.ceil(total))                         # CLng(-Int(-total)), at least 1


def staff_counts_core(people: list[str], right_texts: list[str]) -> tuple[int, int, int]:
    """modMinStaff.StaffCountsCore."""
    scheduled = off = covered = 0
    for tok in people:
        if "S" in tok:
            scheduled += 1
            if "A" in tok or "W" in tok:
                off += 1
    for raw in right_texts:
        rt = vb_trim(raw)
        if rt != "":
            up = rt.upper()
            if not (up.startswith("NO TAKERS") or up.startswith("POSTED") or up == "NCN"):
                covered += covered_count(rt)
    return scheduled, off, covered


def staff_short(scheduled: int, off: int, covered: int, minimum: int) -> str:
    """modMinStaff.StaffShort (rules.py has the same rule for suite minstaff)."""
    if minimum <= 0:
        return ""
    staffed = scheduled - off + covered
    return f"SHORT {staffed}/{minimum}" if staffed < minimum else ""
