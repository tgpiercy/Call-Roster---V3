"""vbacheck's module-level-declaration check fires on a Const after the first procedure (the v27 item 10 HIGH:
a Const between two procedures does not compile, and three rounds of reading had missed it) and stays quiet on a
declarations block in its place. vbacheck is a script, so it runs as one."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent

GOOD = """Attribute VB_Name = "modFixture"
Option Explicit

Public Const ABOVE As Long = 1

Private Sub First()
    Dim x As Long
    x = ABOVE
End Sub

Private Sub Second()
End Sub
"""

BAD = """Attribute VB_Name = "modFixture"
Option Explicit

Private Sub First()
    Dim x As Long
    x = 1
End Sub

Public Const BETWEEN As Long = 1

Private Sub Second()
End Sub
"""


def run(tmp_path: Path, text: str) -> str:
    (tmp_path / "modules").mkdir(exist_ok=True)          # vbacheck scans modules/ and paste-in/ under the root
    (tmp_path / "modules" / "modFixture.bas").write_bytes(text.replace("\n", "\r\n").encode("ascii"))
    out = subprocess.run([sys.executable, str(HERE / "vbacheck.py"), str(tmp_path)], capture_output=True, text=True)
    return out.stdout


def test_a_declaration_after_the_first_procedure_is_a_problem(tmp_path):
    out = run(tmp_path, BAD)
    assert "after the first procedure" in out, out


def test_a_declarations_block_above_the_procedures_is_clean(tmp_path):
    out = run(tmp_path, GOOD)
    assert "after the first procedure" not in out and "OK" in out, out


EXIT_BAD = """Attribute VB_Name = "modFixture"
Option Explicit

Private Function Pick(ByVal x As Long) As Boolean
    If x < 1 Then Exit Sub
    Pick = True
End Function
"""

EXIT_GOOD = """Attribute VB_Name = "modFixture"
Option Explicit

Private Function Pick(ByVal x As Long) As Boolean
    If x < 1 Then Exit Function
    Pick = True
End Function

Private Sub Go()
    If Pick(2) Then Exit Sub
End Sub
"""


def test_exit_sub_inside_a_function_is_a_problem(tmp_path):
    out = run(tmp_path, EXIT_BAD)
    assert "'Exit Sub' inside a Function" in out, out


def test_matching_exits_are_clean(tmp_path):
    out = run(tmp_path, EXIT_GOOD)
    assert "inside a" not in out and "OK" in out, out


PAREN_BAD = """Attribute VB_Name = "modFixture"
Option Explicit

Private Sub Show(ByVal a As String, ByVal b As String)
End Sub

Private Sub Caller()
    Show "one (in a string)", _
         "two" & Format$(Now, "yyyy")
    Show "a", _
         "b")
End Sub
"""


def test_a_stray_bracket_is_a_problem_and_brackets_in_strings_are_not(tmp_path):
    out = run(tmp_path, PAREN_BAD)
    hits = [ln for ln in out.splitlines() if "brackets don't balance" in ln]
    assert len(hits) == 1 and "stray" in hits[0], out


IIF_BAD = """Attribute VB_Name = "modFixture"
Option Explicit

Private Sub Caller(ByVal dc As Variant, ByVal o As Object)
    Dim s As String
    s = IIf(IsArray(dc), UBound(dc) & " lines", "none")
    s = IIf(o Is Nothing, "", o.Name)
    s = IIf(IsArray(dc), "some", "none")
End Sub
"""


def test_an_iif_that_uses_what_it_guards_is_a_problem(tmp_path):
    out = run(tmp_path, IIF_BAD)
    hits = [ln for ln in out.splitlines() if "IIf guards" in ln]
    assert len(hits) == 2, out



# ---- v28 (review): a one-line Sub closes on its own line, so the declaration rule reaches the rest of the file ----
ONELINER_THEN_CONST = """Attribute VB_Name = "modFixture"
Option Explicit

Public Sub GoHome(): ActiveSheet.Range("A1").Select: End Sub

Private Const LATE As Long = 1

Private Sub Second()
    Dim x As Long
    x = LATE
End Sub
"""

ONELINER_CLEAN = """Attribute VB_Name = "modFixture"
Option Explicit

Private Const EARLY As Long = 1

Public Sub GoHome(): ActiveSheet.Range("A1").Select: End Sub
Public Function Two() As Long: Two = EARLY + 1: End Function

Private Sub Second()
    Dim x As Long
    x = Two()
End Sub
"""


def test_a_one_line_sub_does_not_hide_a_late_declaration(tmp_path):
    out = run(tmp_path, ONELINER_THEN_CONST)
    assert "module-level declaration" in out and "LATE" in out, out
    assert "unclosed" not in out, out


def test_one_line_procedures_are_clean(tmp_path):
    out = run(tmp_path, ONELINER_CLEAN)
    assert "OK" in out and "problem" not in out, out
