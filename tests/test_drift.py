"""Drift checks between the VBA and its Python mirror (v25, deep review B3 rank 5).

- The AddCase block pasted into modSelfTest.bas equals tests/cases.csv row for row.
- The constants rules.py copies from the VBA still hold the VBA's values.
- The rotation strings and anchor match modShiftCycle.
- The leave-word list matches modPublish.LeaveWordsIn.
- Every suite in cases.csv is handled by test_cases.py, and every public mirror
  function in rules.py is used somewhere (no dead mirror).
"""
from __future__ import annotations

import csv
import re
from datetime import date
from pathlib import Path

import rules

ROOT = Path(__file__).resolve().parent.parent
MODULES = ROOT / "modules"


def vba(name: str) -> str:
    return (MODULES / name).read_text(encoding="ascii")


def vba_const(module: str, const: str) -> str:
    m = re.search(r"Const\s+" + re.escape(const) + r"\s+As\s+\w+\s*=\s*([^'\r\n]+)", vba(module))
    assert m, f"{module}: Const {const} not found"
    return m.group(1).strip()


def test_selftest_block_matches_cases_csv():
    bas = vba("modSelfTest.bas").replace("\r\n", "\n")
    start = bas.index("' >>> CASES")
    end = bas.index("' <<< CASES")
    pasted = [ln.strip() for ln in bas[start:end].split("\n")[1:] if ln.strip().startswith("AddCase ")]
    pasted_rows = [ln[len('AddCase "'):-1] for ln in pasted]
    with (ROOT / "tests" / "cases.csv").open(newline="") as f:
        rows = list(csv.DictReader(f))
    fields = ["suite", "case", "in1", "in2", "in3", "in4", "in5", "expected"]
    csv_rows = [",".join(r[k] for k in fields) for r in rows]
    assert pasted_rows == csv_rows, "modSelfTest.bas is out of date: run python tests/gen_selftest_cases.py"


def test_charging_constants_match_vba():
    assert float(vba_const("modOvertimeEntry.bas", "CHARGE_MULTIPLIER")) == rules.CHARGE_MULTIPLIER
    assert float(vba_const("modOvertimeEntry.bas", "CHARGE_MAX_GRADE")) == rules.CHARGE_MAX_GRADE
    assert float(vba_const("modLeave.bas", "PEER_MAX_GRADE")) == rules.PEER_MAX_GRADE
    assert (vba_const("modOvertimeEntry.bas", "DST_USE_ACTUAL_HOURS") == "True") == rules.DST_USE_ACTUAL_HOURS
    assert float(vba_const("modOvertimeEntry.bas", "USUAL_SHIFT_HOURS")) == rules.USUAL_SHIFT_HOURS
    assert vba_const("modOvertimeEntry.bas", "BACKUP_BASENAME").strip('"') == rules.BACKUP_BASENAME


def test_attendance_constants_match_vba():
    for name in ("PAID_SICK_DAYS", "PAID_SICK_OCC", "PERSONAL_DAYS", "COMBINED_LIMIT"):
        assert int(vba_const("modAttendance.bas", name)) == getattr(rules, name), name


def test_rotation_matches_vba():
    assert vba_const("modShiftCycle.bas", "ROT_DAYS").strip('"') == rules.ROT_DAYS
    assert vba_const("modShiftCycle.bas", "ROT_NIGHTS").strip('"') == rules.ROT_NIGHTS
    assert int(vba_const("modShiftCycle.bas", "ROT_LEN")) == rules.ROT_LEN == len(rules.ROT_DAYS) == len(rules.ROT_NIGHTS)
    anchor = vba_const("modShiftCycle.bas", "ROT_ANCHOR").strip("#")       # #12/28/2025#
    m, d, y = (int(x) for x in anchor.split("/"))
    assert date(y, m, d) == rules.ROT_ANCHOR


def test_leave_words_match_vba():
    src = vba("modPublish.bas")
    m = re.search(r'Case ("STD"(?:,\s*"[A-Z ]+")+)', src)
    assert m, "modPublish.LeaveWordsIn: the Case line with the leave words wasn't found"
    words = tuple(w.strip().strip('"') for w in m.group(1).split(","))
    assert words == rules.LEAVE_WORDS


def test_every_suite_has_a_handler_and_no_dead_mirror():
    with (ROOT / "tests" / "cases.csv").open(newline="") as f:
        suites = sorted({r["suite"] for r in csv.DictReader(f)})
    tests_src = (ROOT / "tests" / "test_cases.py").read_text(encoding="utf-8")
    missing = [s for s in suites if f'"{s}"' not in tests_src and f"'{s}'" not in tests_src]
    assert not missing, f"cases.csv suites with no handler in test_cases.py: {missing}"
    rules_src = (ROOT / "tests" / "rules.py").read_text(encoding="utf-8")
    public = [m.group(1) for m in re.finditer(r"^def ([a-z][a-z0-9_]*)\(", rules_src, re.M)]
    used = tests_src + rules_src
    dead = [fn for fn in public if len(re.findall(r"\b" + fn + r"\b", used)) < 2]      # the def itself + one use
    assert not dead, f"rules.py functions used nowhere (dead mirrors): {dead}"


def test_only_unlocksheet_unprotects():
    """v25 (R4): code opens a roster sheet through UnlockSheet only (modPublish unprotects the MEMBERS' file)."""
    hits = []
    for folder in (MODULES, ROOT / "paste-in"):
        for p in sorted(folder.glob("*.bas")) + sorted(folder.glob("*.txt")):
            for i, ln in enumerate(p.read_text(encoding="ascii").splitlines(), 1):
                code = ln.split("'")[0]
                if re.search(r"\.Unprotect\b", code, re.I):
                    hits.append((p.name, i, code.strip()))
    roster = [h for h in hits if "SHEET_PASSWORD" in h[2]]
    other = [h for h in hits if "SHEET_PASSWORD" not in h[2]]
    # UnlockSheet (a sheet) and UnlockStructure (the workbook's tabs) are the two allowed roster lines
    assert len(roster) == 2 and all(h[0] == "modOvertimeEntry.bas" for h in roster), roster
    assert all(h[0] == "modPublish.bas" and "MEMBER_PASSWORD" in h[2] for h in other), other
    src = (MODULES / "modOvertimeEntry.bas").read_text(encoding="ascii")
    assert "Public Sub UnlockSheet(" in src and "Public Sub UnlockStructure(" in src


def test_versions_agree():
    """v25 review: SystemCheck's expectations equal what ThisWorkbook and the forms say about themselves."""
    sysrc = vba("modSystem.bas")
    thiswb = re.search(r'THISWB_RELEASE As String = "([^"]+)"', sysrc).group(1)
    wbtxt = (ROOT / "paste-in" / "ThisWorkbook.txt").read_text(encoding="ascii")
    assert re.search(r"^' @version " + re.escape(thiswb) + r" ", wbtxt, re.M), f"ThisWorkbook.txt @version is not {thiswb}"
    cv = re.search(r'CodeVersion = "([^"]+)"', wbtxt).group(1)
    assert thiswb in cv, f"ThisWorkbook.CodeVersion {cv!r} does not carry {thiswb} (SystemCheck would fail)"
    forms = re.search(r'FORM_RELEASES As String = ((?:"[^"]*"\s*&\s*_\s*)*"[^"]*")', sysrc).group(1)
    table = "".join(re.findall(r'"([^"]*)"', forms))
    expected = dict(x.split("|") for x in table.split(","))
    for name, rel_ in expected.items():
        txt = (ROOT / "paste-in" / f"{name}.txt").read_text(encoding="ascii")
        m = re.search(r"^' @version (\S+) ", txt, re.M)
        assert m and m.group(1) == rel_, f"{name}: FORM_RELEASES says {rel_}, the paste-in says {m.group(1) if m else None}"
