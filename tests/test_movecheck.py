"""movecheck proves a split is pure movement; splitmodule makes the split. Both are proved on fixtures here,
and the v27.1 callcheck classes (DUPPUBLIC, QUALREF, MACROSTR's missing module) fire on planted mistakes."""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import callcheck  # noqa: E402
import movecheck  # noqa: E402

HERE = Path(__file__).resolve().parent

OLD_A = """Attribute VB_Name = "modA"
Option Explicit

Private mCount As Long
Public Const A_VERSION As String = "modA v1 (Oct 1)"

' ==== CORE ====
Public Sub Core()
    mCount = mCount + 1
    Helper
End Sub

Private Sub Helper()
    mCount = 0
End Sub

' ==== BOARD ====
' The board's one job.
Public Sub BoardPrint()
    Helper
    Application.Run "modA.BoardPrint"
End Sub

Private Function BoardText() As String
    BoardText = "x"
End Function
"""

OTHER = """Attribute VB_Name = "modOther"
Option Explicit

Public Sub UsesIt()
    modA.BoardPrint
    modA.Core
End Sub
"""


def tree(tmp: Path, name: str, files: dict[str, str]) -> Path:
    root = tmp / name
    (root / "modules").mkdir(parents=True)
    for d in ("paste-in", "member", "updater"):
        (root / d).mkdir()
    for fn, text in files.items():
        (root / "modules" / fn).write_bytes(text.replace("\n", "\r\n").encode("ascii"))
    return root


def split_good() -> dict[str, str]:
    a = OLD_A.split("' ==== BOARD ====")[0].rstrip("\n") + "\n"
    a = a.replace("Private Sub Helper()", "Public Sub Helper()")
    b = ('Attribute VB_Name = "modBoard"\nOption Explicit\n\n' + "' ==== BOARD ====" + OLD_A.split("' ==== BOARD ====")[1])
    b = b.replace('"modA.BoardPrint"', '"modBoard.BoardPrint"')
    return {"modA.bas": a, "modBoard.bas": b,
            "modOther.bas": OTHER.replace("modA.BoardPrint", "modBoard.BoardPrint")}


def check(tmp: Path, new: dict[str, str]) -> movecheck.Report:
    old = tree(tmp, "old", {"modA.bas": OLD_A, "modOther.bas": OTHER})
    nw = tree(tmp, "new", new)
    return movecheck.compare(movecheck.Tree.from_dir(old, "old"), movecheck.Tree.from_dir(nw, "new"))


def tags(rep: movecheck.Report) -> list[str]:
    return [f.split()[0] for f in rep.fails]


def test_a_clean_split_passes_and_lists_the_promotion(tmp_path):
    rep = check(tmp_path, split_good())
    assert rep.fails == [], rep.fails
    assert any(n.startswith("PROMOTE") and "Helper" in n for n in rep.notes), rep.notes
    assert sum(n.startswith("MOVE") for n in rep.notes) == 2
    assert any(n.startswith("QUALIFY") and "modOther" in n for n in rep.notes), rep.notes


def test_a_changed_body_fails(tmp_path):
    new = split_good()
    new["modBoard.bas"] = new["modBoard.bas"].replace('BoardText = "x"', 'BoardText = "y"')
    assert tags(check(tmp_path, new)) == ["BODY"]


def test_a_lost_and_an_added_procedure_fail(tmp_path):
    new = split_good()
    new["modBoard.bas"] = new["modBoard.bas"].replace("Private Function BoardText", "Private Function BoardWords") \
        .replace("BoardText = ", "BoardWords = ").replace("End Function", "End Function")
    t = tags(check(tmp_path, new))
    assert "LOST" in t and "NEW" in t, t


def test_a_private_reached_across_the_cut_fails(tmp_path):
    new = split_good()
    new["modA.bas"] = new["modA.bas"].replace("Public Sub Helper()", "Private Sub Helper()")
    rep = check(tmp_path, new)
    assert "SCOPE" in tags(rep) and any("Helper" in f for f in rep.fails), rep.fails


def test_a_qualifier_left_on_the_old_module_fails(tmp_path):
    new = split_good()
    new["modOther.bas"] = OTHER                       # still says modA.BoardPrint
    rep = check(tmp_path, new)
    assert any("SCOPE" in f and "modA.BoardPrint" in f for f in rep.fails), rep.fails


def test_a_wrong_qualifier_in_another_file_fails(tmp_path):
    new = split_good()
    new["modOther.bas"] = OTHER.replace("modA.Core", "modBoard.Core")
    assert "OTHER" in tags(check(tmp_path, new))


def test_a_changed_declaration_fails_and_a_new_version_constant_does_not(tmp_path):
    new = split_good()
    new["modA.bas"] = new["modA.bas"].replace("Private mCount As Long", "Private mCount As Double")
    assert "DECL" in tags(check(tmp_path, new))
    new = split_good()
    new["modBoard.bas"] = new["modBoard.bas"].replace("Option Explicit\n", 'Option Explicit\nPublic Const BOARD_VERSION As String = "modBoard v1 (Oct 1)"\n')
    rep = check(tmp_path / "v", new)
    assert rep.fails == [] and any(n.startswith("VERSION") for n in rep.notes), rep.fails


def test_public_to_private_fails(tmp_path):
    new = split_good()
    new["modBoard.bas"] = new["modBoard.bas"].replace("Public Sub BoardPrint", "Private Sub BoardPrint")
    assert "BODY" in tags(check(tmp_path, new))


def test_splitmodule_makes_a_split_movecheck_passes(tmp_path):
    root = tree(tmp_path, "s", {"modA.bas": OLD_A, "modOther.bas": OTHER})
    lines = OLD_A.split("\n")
    first = lines.index("' ==== BOARD ====") + 1
    last = len(lines) - 1                              # the text ends with a newline
    r = subprocess.run([sys.executable, str(HERE / "splitmodule.py"), str(root), "modA", f"modBoard={first}-{last}"],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr
    after = movecheck.Tree.from_dir(root, "after")
    assert {f.name for f in after.master()} == {"modA", "modBoard", "modOther"}
    rep = movecheck.compare(movecheck.Tree.from_dir(tree(tmp_path, "o", {"modA.bas": OLD_A, "modOther.bas": OTHER}), "o"), after)
    t = tags(rep)
    assert "SCOPE" in t and not set(t) - {"SCOPE"}, rep.fails      # the move is clean; Helper and the qualifiers still to fix


def test_splitmodule_refuses_a_range_that_cuts_a_procedure(tmp_path):
    root = tree(tmp_path, "s", {"modA.bas": OLD_A})
    lines = OLD_A.split("\n")
    mid = lines.index("    Helper") + 1                # inside Core
    r = subprocess.run([sys.executable, str(HERE / "splitmodule.py"), str(root), "modA", f"modBoard={mid}-{len(lines) - 1}"],
                       capture_output=True, text=True)
    assert r.returncode == 1 and "cuts the procedure" in r.stdout, r.stdout
    assert not (root / "modules" / "modBoard.bas").exists()


# ---- callcheck v27.1 ----
FIX_A = """Attribute VB_Name = "modFixA"
Option Explicit

Public gShared As Long

Public Sub Twice()
End Sub

Private Sub Hidden()
End Sub

Public Sub Caller()
    modFixB.Here
    modFixB.Gone
    modFixB.Secret
    modFixA.Hidden
    Application.Run "modNowhere.Twice"
End Sub
"""
FIX_B = """Attribute VB_Name = "modFixB"
Option Explicit

Public gShared As Long

Public Sub Twice()
End Sub

Public Sub Here()
End Sub

Private Sub Secret()
End Sub
"""


def test_duppublic_qualref_and_a_missing_module(tmp_path, capsys):
    files = []
    for name, text in (("modFixA.bas", FIX_A), ("modFixB.bas", FIX_B)):
        f = tmp_path / name
        f.write_bytes(text.replace("\n", "\r\n").encode("ascii"))
        callcheck._RAW.pop(str(f), None)
        files.append(f)
    callcheck.check_project(files, "master")
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    dup = [ln for ln in out if ln.startswith("DUPPUBLIC")]
    assert sorted(ln.split(": ")[1] for ln in dup) == ["gshared", "twice"], dup
    qual = [ln for ln in out if ln.startswith("QUALREF")]
    assert len(qual) == 2 and any("Gone" in ln for ln in qual) and any("Secret" in ln for ln in qual), qual
    assert not any("modFixA.Hidden" in ln for ln in qual)          # its own module may qualify a Private
    assert any(ln.startswith("MACROSTR") and "modNowhere" in ln for ln in out), out
