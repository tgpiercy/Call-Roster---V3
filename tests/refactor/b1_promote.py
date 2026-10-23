"""v27.1 builder B1, commit 2: what the split needs to compile (Graeme's decision: state Public AS-IS).

Run from the repo root on the move commit (tests/refactor/b1_move.py already run):

    py -3.14 tests/refactor/b1_promote.py

1. Private -> Public on exactly the names movecheck's SCOPE listed after the move (a module-level Dim would become
   Public; there is none). Each name is found as ONE module-level declaration or procedure header in its module;
   anything else (missing, twice, inside a combined declaration) stops the script before any file is written.
2. The 7 macro strings whose target moved: "modCalendar.X" -> "modNeeds.X" / "modBoard.X" / "modCalImport.X".

No other text changes. Then:  py -3.14 tests/movecheck.py ab858ac  (0 failures).
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

PROMOTE = {
    "modCalendar": [
        # state
        "mAbs", "mCov", "mShiftMap", "mPostV", "mDrawing", "mTimeSlots",
        "CAL_FONT", "CAL_FONT_B", "FULL_SHIFT", "NA_KEY",
        # procedures
        "AbsActiveOn", "AbsenceOnShift", "AddToRange", "AppendTo", "AsDate", "AskNumber", "CalBtn", "CalDrop",
        "CalEditReasonId", "CalEditTextKey", "CalendarSheet", "CalGroup", "CalMonth", "CalPostKey",
        "CalSetCoverageKey", "Clip", "CovHasSourceCached", "DrawContext", "Dt2", "FormatCovDates", "HomeUnitCached",
        "HrsTag", "KeyDate", "KeyShift", "LineDetails", "LoadCache", "LoadPostCache", "MapSheet", "MapWrite",
        "NAStale", "NextId", "PostIdOfLine", "PostingDetails", "PostRow", "RowOf", "Secs", "SingleUnitOf",
        "SlotTexts", "SlotUnit", "StampTxt", "TableData", "Tint",
        "thBand", "thGreen", "thGrid", "thHead", "thMuted", "thNav", "thOpen", "thOpenFill", "thPosted",
        "thPostedFill", "thToday",
    ],
    "modMinStaff": ["StaffSweepBegin", "MinStaffCount", "StaffShortList"],
}

# "modCalendar.X" -> new module, for the macros that moved (callcheck MACROSTR)
MACROS = {
    "ShowNeedsAttention": "modNeeds",
    "BoardPrint": "modBoard",
    "ImportUnitCalendar": "modCalImport",
    "NARefresh": "modNeeds",
    "NAGoCalendar": "modNeeds",
    "NAUnitChanged": "modNeeds",
    "NARangeChanged": "modNeeds",
}

PROC = re.compile(r"^(Private)(\s+(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+)(\w+)\b", re.I)
VAR = re.compile(r"^(Private|Dim)(\s+(?:Const\s+)?)(\w+)\b", re.I)
START = re.compile(r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+", re.I)
END = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)


def load(path: Path) -> tuple[list[str], bool]:
    lines = path.read_bytes().decode("ascii").split("\r\n")
    trailing = bool(lines) and lines[-1] == ""
    if trailing:
        lines.pop()
    return lines, trailing


def save(path: Path, lines: list[str], trailing: bool) -> None:
    path.write_bytes(("\r\n".join(lines) + ("\r\n" if trailing else "")).encode("ascii"))


def main() -> int:
    root = Path(".")
    edits: dict[Path, tuple[list[str], bool]] = {}
    report = []
    for mod, names in PROMOTE.items():
        path = root / "modules" / f"{mod}.bas"
        lines, trailing = load(path)
        want = {n.lower(): n for n in names}
        found: dict[str, list[int]] = {k: [] for k in want}
        inproc = False
        for i, ln in enumerate(lines):
            if START.match(ln):
                m = PROC.match(ln)
                if m and m.group(3).lower() in want:
                    found[m.group(3).lower()].append(i)
                inproc = not re.search(r":\s*End\s+(?:Sub|Function|Property)\s*(?:'.*)?$", ln, re.I)
                continue
            if END.match(ln):
                inproc = False
                continue
            if inproc:
                continue
            m = VAR.match(ln)
            if m and m.group(3).lower() in want:
                code = ln.split("'")[0]
                if "," in code:
                    print(f"{mod}:{i + 1}: combined declaration, promote by hand: {ln.strip()}")
                    return 1
                found[m.group(3).lower()].append(i)
        for k, hits in found.items():
            if len(hits) != 1:
                print(f"{mod}: {want[k]} found {len(hits)} time(s) at module level: refused")
                return 1
        for k, (i,) in found.items():
            ln = lines[i]
            new = re.sub(r"^(Private|Dim)\b", "Public", ln, count=1)
            assert new != ln, ln
            lines[i] = new
            report.append((mod, want[k], i + 1))
        edits[path] = (lines, trailing)

    # the macro strings, anywhere in modules/ and paste-in/
    moved = 0
    files = sorted((root / "modules").glob("*.bas")) + sorted((root / "paste-in").glob("*.txt"))
    for path in files:
        lines, trailing = edits.get(path) or load(path)
        changed = False
        for i, ln in enumerate(lines):
            for name, dest in MACROS.items():
                old = f'"modCalendar.{name}"'
                if old in ln:
                    ln = ln.replace(old, f'"{dest}.{name}"')
                    changed = True
                    moved += 1
                    print(f"string {path.as_posix()}:{i + 1}: modCalendar.{name} -> {dest}.{name}")
            lines[i] = ln
        if changed:
            edits[path] = (lines, trailing)

    for path, (lines, trailing) in edits.items():
        save(path, lines, trailing)
    for mod, name, ln in sorted(report, key=lambda r: (r[0], r[2])):
        print(f"Public {mod}:{ln} {name}")
    print(f"{len(report)} promoted, {moved} macro string(s) changed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
