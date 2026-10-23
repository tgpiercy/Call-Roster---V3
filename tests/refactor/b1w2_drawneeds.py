"""v27.1 builder B1, wave 2 commit 3: DrawNeeds in named steps, made by script (never retyped).

Run from the repo root on 7396c80 (modNeeds.bas as merged at 39f66e4; DrawNeeds = lines 180-656):

    py -3.14 tests/refactor/b1w2_drawneeds.py

The statements move unchanged and in order into three Private Subs in modNeeds, placed right after DrawNeeds:
  NAGatherWork  205-340  A open postings, B open lines, B2 below minimum staffing, C paybacks owed, D leave,
                         E holidays to confirm
  NAGatherAdmin 342-496  H imported lines, G members' file, the errors / exceptions / disputes / restrictions /
                         feedback, F blank home units, D2 the data check, V planned moves, T attendance
  NADrawHead    499-538  the sheet wiped, the columns, the title and the subtitle rows
The row lists and their counts (and nReady, hLook) are declared in DrawNeeds and passed ByRef; each step ReDims
its own lists where the inline code did. Every Timer reading (t0, tA, tB, tC), mDrawing / mTimeSlots, the
unlock / relock and the handler stay in DrawNeeds between the same statements. The steps have no handler: an
error rises to DrawNeeds' Failed. Each inline "On Error GoTo Failed" that closed an On Error Resume Next inside
a step becomes "On Error GoTo 0" (B4's rule): the next error again rises to Failed. No step is called while
On Error Resume Next is active.

The only other edits: the Dims of the lists, counts, nReady and hLook move from the steps to DrawNeeds (the rest
of each Dim line stays in its step) and NAGatherAdmin declares its own loop counter i. The script expands every
call back into its step's body and prints the difference from the original procedure.
"""
from __future__ import annotations

import difflib
import sys
from pathlib import Path

PATH = Path("modules/modNeeds.bas")
FIRST, LAST = 180, 656

EXPECT = {
    180: "Public Sub DrawNeeds()",
    194: "    On Error GoTo Failed",
    198: "    Dim plV As Variant: plV = PostListData()",
    204: "    ReDim shortRows(1 To 50)",
    205: "    Dim i As Long, j As Long, m As Long, cl As Date, sd As Date, d As Date, sh As Variant, rt As String",
    206: "    Dim nReady As Long, tier As Long, st As String, stC As Long, stF As Long, id As String, txt As String",
    208: "    ' A. postings still open",
    274: "    On Error Resume Next",
    276: "    On Error GoTo Failed",
    329: "    Dim hr() As NARow, nH As Long, hv As Variant, hk As Variant, hd As Date, hLook As Long",
    340: "    End If",
    342: "    ' H. imported vacation lines not yet booked (v22) - they don't keep people off call lists",
    343: "    Dim ir() As NARow, nI As Long, iv As Variant",
    345: "    On Error Resume Next",
    347: "    On Error GoTo Failed",
    361: "    Dim mr() As NARow, nM As Long, pubAt As Date, pubTxt As String",
    375: "    Dim er() As NARow, nE As Long, nErrLines As Long",
    423: "    Dim ur() As NARow, nU As Long, uv As Variant, uk As Variant",
    425: "    On Error Resume Next",
    427: "    On Error GoTo Failed",
    436: "    Dim dcR() As NARow, nDc As Long, dcV As Variant, dcL As Variant, dcP As Variant",
    438: "    On Error Resume Next",
    440: "    On Error GoTo Failed",
    450: "    Dim vr() As NARow, nV As Long, pmv As Variant, pmD As Date, pmDue As Boolean, pmWho As String",
    452: "    On Error Resume Next",
    454: "    On Error GoTo Failed",
    470: "    Dim tr() As NARow, nT As Long, rn As Variant, ast As AttendanceStatus, lvl As Long, hu As String, nmT As String",
    472: "    On Error Resume Next",
    496: "    On Error GoTo Failed",
    498: "    tA = Timer                                      ' gathered",
    499: "    ' ---- draw ----",
    538: "    ws.Rows(6).RowHeight = 12",
    540: "    Dim r As Long: r = 7",
    656: "End Sub",
}

# line -> its replacement inside a step (None = the whole line moves to DrawNeeds' new Dims)
EDITS = {
    206: "    Dim tier As Long, st As String, stC As Long, stF As Long, id As String, txt As String",
    276: "    On Error GoTo 0",
    329: "    Dim hv As Variant, hk As Variant, hd As Date",
    343: "    Dim iv As Variant",
    347: "    On Error GoTo 0",
    361: "    Dim pubAt As Date, pubTxt As String",
    375: "    Dim nErrLines As Long",
    423: "    Dim uv As Variant, uk As Variant",
    427: "    On Error GoTo 0",
    436: "    Dim dcV As Variant, dcL As Variant, dcP As Variant",
    440: "    On Error GoTo 0",
    450: "    Dim pmv As Variant, pmD As Date, pmDue As Boolean, pmWho As String",
    454: "    On Error GoTo 0",
    470: "    Dim rn As Variant, ast As AttendanceStatus, lvl As Long, hu As String, nmT As String",
    496: "    On Error GoTo 0",
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

    def L(a: int, b: int) -> list[str]:
        """1-based, inclusive, with the step edits applied"""
        return [EDITS.get(n, lines[n - 1]) for n in range(a, b + 1)]

    caller = (
        L(180, 204)
        + ["    Dim nReady As Long                                ' v27.1: the lists the steps fill (and two counts) live here",
           "    Dim hr() As NARow, nH As Long, hLook As Long",
           "    Dim ir() As NARow, nI As Long, mr() As NARow, nM As Long, er() As NARow, nE As Long",
           "    Dim ur() As NARow, nU As Long, dcR() As NARow, nDc As Long, vr() As NARow, nV As Long",
           "    Dim tr() As NARow, nT As Long",
           "    NAGatherWork grp, nDays, plV, pr, nP, nReady, lr, nL, shortRows, nShort, sr, nS, br, nB, hr, nH, hLook",
           "    NAGatherAdmin grp, ir, nI, mr, nM, er, nE, ur, nU, dcR, nDc, vr, nV, tr, nT"]
        + L(497, 498)
        + ["    NADrawHead ws, grp, nDays                       ' v27.1: the sheet wiped, the columns, rows 1-6"]
        + L(539, 656)
    )
    work = (
        ["' v27.1: DrawNeeds' steps. No handler of their own: an error rises to DrawNeeds' Failed, which restores",
         "' everything. Called only while that handler is active. The lists and counts are DrawNeeds' own (ByRef).",
         "Private Sub NAGatherWork(ByVal grp As String, ByVal nDays As Long, ByRef plV As Variant, _",
         "                         ByRef pr() As NARow, ByRef nP As Long, ByRef nReady As Long, _",
         "                         ByRef lr() As NARow, ByRef nL As Long, ByRef shortRows() As NARow, ByRef nShort As Long, _",
         "                         ByRef sr() As NARow, ByRef nS As Long, ByRef br() As NARow, ByRef nB As Long, _",
         "                         ByRef hr() As NARow, ByRef nH As Long, ByRef hLook As Long)"]
        + L(205, 340)
        + ["End Sub", ""]
    )
    admin = (
        ["Private Sub NAGatherAdmin(ByVal grp As String, ByRef ir() As NARow, ByRef nI As Long, _",
         "                          ByRef mr() As NARow, ByRef nM As Long, ByRef er() As NARow, ByRef nE As Long, _",
         "                          ByRef ur() As NARow, ByRef nU As Long, ByRef dcR() As NARow, ByRef nDc As Long, _",
         "                          ByRef vr() As NARow, ByRef nV As Long, ByRef tr() As NARow, ByRef nT As Long)",
         "    Dim i As Long"]
        + L(342, 496)
        + ["End Sub", ""]
    )
    head = (
        ["Private Sub NADrawHead(ByVal ws As Worksheet, ByVal grp As String, ByVal nDays As Long)"]
        + L(499, 538)
        + ["End Sub"]
    )
    new_lines = lines[:FIRST - 1] + caller + [""] + work + admin + head + lines[LAST:]

    # proof: expand every call into its step's body (edits included) and compare with the original procedure
    bodies = {"NAGatherWork": work[7:-2], "NAGatherAdmin": admin[5:-2], "NADrawHead": head[1:-1]}
    expanded = []
    for ln in caller:
        word = ln.strip().split(" ")[0]
        expanded += bodies.get(word, [ln])
    diff = list(difflib.unified_diff(lines[FIRST - 1:LAST], expanded, "original", "expanded", lineterm="", n=0))
    print("\n".join(diff) if diff else "(no difference)")

    PATH.write_bytes(("\r\n".join(new_lines) + ("\r\n" if trailing else "")).encode("ascii"))
    for name, proc in (("DrawNeeds", caller), ("NAGatherWork", work[2:-1]), ("NAGatherAdmin", admin[:-1]),
                       ("NADrawHead", head)):
        print(f"{name}: {len(proc)} lines")
    return 0


if __name__ == "__main__":
    sys.exit(main())
