"""The test-only snapshot module (tests/snapshot/modSnapshot.bas, v27.1 refactor proof) must import into the roster as
it stands: every name it uses declared, every call matching. Wave 2's draw breakups must keep DrawCalendar, DrawNeeds,
CalShowDate, BoardPrint, PostingOutput, Call_Out and DrawDesk Public under the same names, or step 4 of the test
cannot run on the new release."""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import callcheck  # noqa: E402
import undeclcheck  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SNAP = ROOT / "tests" / "snapshot" / "modSnapshot.bas"


def test_snapshot_compiles_against_the_modules(tmp_path, capsys):
    files = []
    for p in sorted((ROOT / "modules").glob("*.bas")) + sorted((ROOT / "paste-in").glob("*.txt")) + [SNAP]:
        q = tmp_path / p.name
        shutil.copyfile(p, q)
        files.append(q)
    undeclcheck.check_project(files, "master")
    callcheck.check_project(files, "master")
    out = [ln for ln in capsys.readouterr().out.splitlines() if "modSnapshot" in ln]
    assert out == [], out


def test_snapshot_is_never_shipped():
    assert not (ROOT / "modules" / "modSnapshot.bas").exists()
    raw = SNAP.read_bytes()
    assert b"\r\n" in raw and b"\n" not in raw.replace(b"\r\n", b""), "CRLF only"
    raw.decode("ascii")
