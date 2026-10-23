"""v27.1 builder B1, wave 2 commit 2: DrawCalendar in named steps, made by script (never retyped).

Run from the repo root on e854dfc (modCalendar.bas as merged at 39f66e4; DrawCalendar = lines 2371-2762):

    py -3.14 tests/refactor/b1w2_drawcalendar.py

The statements move unchanged and in order into four Private Subs in modCalendar, placed right after DrawCalendar:
  CalDrawWipe   2400-2414  the grid and the click map wiped, the columns sized
  CalDrawHeader 2417-2472  rows 1-4: the title band, the legend chips, the toolbar / hint row, the weekdays
  CalDrawWeeks  2476-2670  the week blocks and the closing border
  CalDrawPage   2675-2715  the page setup and the window (its own On Error Resume Next, as inline)
Every Timer reading, mDrawing / mTimeSlots, the unlock / relock and the handler stay in DrawCalendar, between the
same statements. The only edits besides the call lines: the Dims of the locals the steps share are split or
repeated (nLines / r in DrawCalendar; k, i, shortTxt, nShortD in the steps that use them) and "Dim r As Long:
r = 4" becomes "r = 4" (r is the header's ByRef parameter); CalDrawPage ends with "On Error GoTo 0" so its
Resume Next visibly stops where the inline one did (callcheck ORNSWALLOW; the handler state is per procedure, so
it changes nothing).

The script then expands every call back into its step's body and prints the difference from the original
procedure, which must be only those Dim lines and the call lines.
"""
from __future__ import annotations

import difflib
import sys
from pathlib import Path

PATH = Path("modules/modCalendar.bas")
FIRST, LAST = 2371, 2762

EXPECT = {
    2371: "Public Sub DrawCalendar()",
    2379: "    Dim nLines As Long, shortTxt As String, nShortD As Long   ' v26: the date's short shifts (minimum staffing)",
    2398: "    tA = Timer                                      ' caches loaded",
    2400: "    ' ---- wipe the grid and the click map (controls are rebuilt at the end) ----",
    2410: "    Dim k As Long",
    2414: "    ws.Columns(16).ColumnWidth = 1",
    2415: "    tB = Timer                                      ' grid wiped",
    2417: "    ' ---- row 1: title band (controls sit on it) ----",
    2429: "    Dim lg As Variant, i As Long, parts() As String",
    2459: "    Dim r As Long: r = 4",
    2472: "    r = r + 1",
    2473: "    tC = Timer                                      ' header rows written",
    2474: "    mTimeSlots = 0: mTimeNotes = 0                  ' the weeks' share starts here",
    2476: "    ' ---- weeks ----",
    2670: "    End With",
    2671: "    tD = Timer                                      ' weeks written",
    2675: "    ' ---- page ----",
    2676: "    On Error Resume Next",
    2715: "    End If",
    2716: "    On Error GoTo Failed",
    2762: "End Sub",
}


def main() -> int:
    raw = PATH.read_bytes().decode("ascii")
    lines = raw.split("\r\n")
    trailing = lines[-1] == ""
    if trailing:
        lines.pop()
    for n, txt in EXPECT.items():
        if lines[n - 1] != txt:
            print(f"line {n} is not what this script was written for:\n  want {txt!r}\n  have {lines[n - 1]!r}")
            return 1
    L = lambda a, b: lines[a - 1:b]           # noqa: E731  1-based, inclusive

    caller = (
        L(2371, 2378)
        + ["    Dim nLines As Long                              ' the lines drawn, counted by CalDrawWeeks (for Z14)"]
        + L(2380, 2399)
        + ["    CalDrawWipe ws, mp                              ' v27.1: the grid and the click map wiped"]
        + L(2415, 2416)
        + ["    Dim r As Long",
           "    CalDrawHeader ws, mp, r                         ' v27.1: rows 1-4; r = the first week row"]
        + L(2473, 2475)
        + ["    CalDrawWeeks ws, mp, m0, grp, r, nLines         ' v27.1: the weeks; r moves past the last one"]
        + L(2671, 2674)
        + ["    CalDrawPage ws, mp, r                           ' v27.1: the page setup and the window (Resume Next inside)"]
        + L(2716, 2762)
    )
    wipe = (
        ["' v27.1: DrawCalendar's steps. No handler of their own: an error rises to DrawCalendar's Failed, which",
         "' restores everything. Called only while that handler is active.",
         "Private Sub CalDrawWipe(ByVal ws As Worksheet, ByVal mp As Worksheet)"]
        + L(2400, 2414)
        + ["End Sub", ""]
    )
    hdr_body = L(2417, 2472)
    hdr_body[2459 - 2417] = "    r = 4"
    header = (
        ["Private Sub CalDrawHeader(ByVal ws As Worksheet, ByVal mp As Worksheet, ByRef r As Long)",
         "    Dim k As Long"]
        + hdr_body
        + ["End Sub", ""]
    )
    weeks = (
        ["Private Sub CalDrawWeeks(ByVal ws As Worksheet, ByVal mp As Worksheet, ByVal m0 As Date, ByVal grp As String, _",
         "                         ByRef r As Long, ByRef nLines As Long)",
         "    Dim k As Long, i As Long",
         "    Dim shortTxt As String, nShortD As Long          ' v26: the date's short shifts (minimum staffing)"]
        + L(2476, 2670)
        + ["End Sub", ""]
    )
    page = (
        ["' the page setup and the window: everything under On Error Resume Next, as it was inline (the caller's",
         "' On Error GoTo Failed comes back when this returns)",
         "Private Sub CalDrawPage(ByVal ws As Worksheet, ByVal mp As Worksheet, ByVal r As Long)"]
        + L(2675, 2715)
        + ["    On Error GoTo 0                                 ' v27.1: ends this step's Resume Next (DrawCalendar's GoTo Failed follows)",
           "End Sub"]
    )
    new_lines = lines[:FIRST - 1] + caller + [""] + wipe + header + weeks + page + lines[LAST:]

    # proof: expand every call into its step's body and compare with the original procedure
    bodies = {"CalDrawWipe": wipe[3:-2], "CalDrawHeader": header[2:-2], "CalDrawWeeks": weeks[4:-2],
              "CalDrawPage": page[3:-1]}
    expanded = []
    for ln in caller:
        word = ln.strip().split(" ")[0]
        expanded += bodies.get(word, [ln])
    diff = list(difflib.unified_diff(L(FIRST, LAST), expanded, "original", "expanded", lineterm="", n=0))
    print("\n".join(diff) if diff else "(no difference)")

    PATH.write_bytes(("\r\n".join(new_lines) + ("\r\n" if trailing else "")).encode("ascii"))
    for name, proc in (("DrawCalendar", caller), ("CalDrawWipe", wipe[2:-1]), ("CalDrawHeader", header[:-1]),
                       ("CalDrawWeeks", weeks[:-1]), ("CalDrawPage", page[2:])):
        print(f"{name}: {len(proc)} lines")
    return 0


if __name__ == "__main__":
    sys.exit(main())
