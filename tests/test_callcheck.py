"""The v25 callcheck classes catch what they claim to, and accept the patterns the roster uses.

A fixture module is written to a temp folder and checked with callcheck.check_project, so a
clean run on the real modules is backed by a run that is known to fire.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import callcheck  # noqa: E402

FIXTURE = """Attribute VB_Name = "modFixture"
Option Explicit

Private Function NeverAssigned(ByVal x As Long) As Boolean
    If x > 1 Then Exit Function
End Function

Private Function Assigned(ByVal x As Long) As Boolean
    If x > 1 Then Assigned = True
End Function

Private Function AssignedOnOneLineIf(ByVal x As Long) As String
    If x > 1 Then AssignedOnOneLineIf = "big" Else AssignedOnOneLineIf = "small"
End Function

Private Sub OffNoRestore()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableCancelKey = xlErrorHandler
End Sub

Private Sub OffRestoredHere()
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableCancelKey = xlErrorHandler
    Application.EnableCancelKey = xlInterrupt
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = evWas
    Application.ScreenUpdating = True
End Sub

Private Sub OffRestoredByHelper()
    Application.EnableEvents = False
    QuietOff
End Sub

Private Sub QuietOff()
    Application.EnableEvents = True
End Sub

Private Sub OffHandedOff()
    Application.ScreenUpdating = False      ' restored by Workbook_AfterSave
End Sub

Private Sub OffOnlyReportError()
    On Error GoTo Oops
    Application.EnableEvents = False
    Exit Sub
Oops:
    ReportError "OffOnlyReportError"
End Sub

Private Sub ReportError(ByVal who As String)
    Application.EnableEvents = True
End Sub

Private Sub UntitledBoxes()
    MsgBox "no title at all"
    MsgBox "no title", vbExclamation
    If MsgBox("no title either", vbYesNo + vbQuestion) = vbYes Then Exit Sub
    MsgBox "titled", vbInformation, "Call Out Roster"
    If MsgBox("titled too", vbYesNo, "Call Out Roster") = vbNo Then Exit Sub
    MsgBox "one, two, three", vbInformation, "a, title"
End Sub

Private Sub MissingSet()
    Dim ws As Worksheet, lo As ListObject, d As Object, r As Range, n As Long, v As Variant, m As VbMsgBoxResult
    ws = ActiveSheet
    Set lo = ws.ListObjects(1)
    If n = 0 Then d = CreateObject("Scripting.Dictionary")
    Set r = ws.Range("A1")
    n = 1
    v = ws
    m = vbYes
End Sub
"""


def run(tmp_path: Path, capsys) -> list[str]:
    f = tmp_path / "modFixture.bas"
    f.write_bytes(FIXTURE.replace("\n", "\r\n").encode("ascii"))
    callcheck.check_project([f], "fixture")
    return [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]


def test_noret_fires_only_on_the_unassigned_function(tmp_path, capsys):
    out = run(tmp_path, capsys)
    noret = [ln for ln in out if ln.startswith("NORET")]
    assert len(noret) == 1 and "NeverAssigned" in noret[0], noret


def test_restore_fires_on_the_bare_off_and_the_error_only_restore(tmp_path, capsys):
    out = run(tmp_path, capsys)
    restore = [ln for ln in out if ln.startswith("RESTORE")]
    procs = sorted({ln.split(": ")[1] for ln in restore})
    assert procs == ["OffNoRestore", "OffOnlyReportError"], restore
    assert sum("OffNoRestore" in ln for ln in restore) == 4          # events, screen, calculation, the cancel key (v26.5.5)


def test_set_fires_on_the_object_assignments_only(tmp_path, capsys):
    out = run(tmp_path, capsys)
    sets = [ln for ln in out if ln.startswith("SET")]
    names = sorted(ln.split("'")[1] for ln in sets)
    assert names == ["d", "ws"], sets


def test_msgtitle_fires_on_the_untitled_boxes_only(tmp_path, capsys):
    out = run(tmp_path, capsys)
    boxes = [ln for ln in out if ln.startswith("MSGTITLE")]
    assert len(boxes) == 3 and all("UntitledBoxes" in ln for ln in boxes), boxes


def test_nothing_else_fires(tmp_path, capsys):
    out = run(tmp_path, capsys)
    other = [ln for ln in out if not ln.startswith(("NORET", "RESTORE", "SET", "MSGTITLE"))]
    assert other == [], other
