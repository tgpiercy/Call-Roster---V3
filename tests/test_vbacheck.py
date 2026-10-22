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
    (tmp_path / "modFixture.bas").write_bytes(text.replace("\n", "\r\n").encode("ascii"))
    out = subprocess.run([sys.executable, str(HERE / "vbacheck.py"), str(tmp_path)], capture_output=True, text=True)
    return out.stdout


def test_a_declaration_after_the_first_procedure_is_a_problem(tmp_path):
    out = run(tmp_path, BAD)
    assert "after the first procedure" in out, out


def test_a_declarations_block_above_the_procedures_is_clean(tmp_path):
    out = run(tmp_path, GOOD)
    assert "after the first procedure" not in out and "OK" in out, out
