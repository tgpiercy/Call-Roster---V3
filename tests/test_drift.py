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


def test_loadcases_chunks_stay_under_the_procedure_limit():
    """v27 (review C4): VBA refuses a procedure over 64 KB; LoadCases reached 57 KB, so the generator writes
    LoadCases1..N of 200 cases each. Every chunk stays well under the limit, and LoadCases only calls them."""
    bas = vba("modSelfTest.bas")
    procs = re.findall(r"Private Sub (LoadCases\d*)\(\)(.*?)End Sub", bas, re.S)
    names = [p[0] for p in procs]
    assert "LoadCases" in names and "LoadCases1" in names
    for name, body in procs:
        assert len(body) < 56 * 1024, f"{name} is {len(body)} bytes: raise the chunk count in gen_selftest_cases.py"
        if name == "LoadCases":
            calls = [ln.strip() for ln in body.split("\n") if ln.strip() and not ln.strip().startswith("'")]
            assert calls == [f"LoadCases{i + 1}" for i in range(len(names) - 1)], "LoadCases must call every chunk in order"


def test_charging_constants_match_vba():
    assert float(vba_const("modOvertimeEntry.bas", "CHARGE_MULTIPLIER")) == rules.CHARGE_MULTIPLIER
    assert float(vba_const("modOvertimeEntry.bas", "CHARGE_MAX_GRADE")) == rules.CHARGE_MAX_GRADE
    assert float(vba_const("modLeave.bas", "PEER_MAX_GRADE")) == rules.PEER_MAX_GRADE
    assert (vba_const("modOvertimeEntry.bas", "DST_USE_ACTUAL_HOURS") == "True") == rules.DST_USE_ACTUAL_HOURS
    assert float(vba_const("modOvertimeEntry.bas", "USUAL_SHIFT_HOURS")) == rules.USUAL_SHIFT_HOURS
    assert vba_const("modOvertimeEntry.bas", "BACKUP_BASENAME").strip('"') == rules.BACKUP_BASENAME


def test_backup_retention_matches_vba():
    """v27 item 12: the three retention constants are one table in two places."""
    assert int(vba_const("modOvertimeEntry.bas", "BACKUP_KEEP_DAYS")) == rules.BACKUP_KEEP_DAYS
    assert int(vba_const("modOvertimeEntry.bas", "BACKUP_HOURLY_KEEP_DAYS")) == rules.BACKUP_HOURLY_KEEP_DAYS
    assert int(vba_const("modOvertimeEntry.bas", "BACKUP_PERIOD_KEEP_DAYS")) == rules.BACKUP_PERIOD_KEEP_DAYS


def test_kronos_code_table_matches_vba():
    """v27: the premium codes and their multipliers are one table in two places."""
    assert vba_const("modOvertimeEntry.bas", "PREMIUM_CODES").strip('"') == rules.PREMIUM_CODES
    assert vba_const("modOvertimeEntry.bas", "PREMIUM_MULTIPLIERS").strip('"') == rules.PREMIUM_MULTIPLIERS
    assert len(rules.PREMIUM_CODES.split(",")) == len(rules.PREMIUM_MULTIPLIERS.split(","))
    assert int(vba_const("ImportKronosData.bas", "KRONOS_YEAR_SLACK")) == rules.KRONOS_YEAR_SLACK


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


def test_module_stamps_carry_the_release():
    """v25.5 review: every stamped module carries EXPECTED_RELEASE, and N_MODULES equals the VerCheck count."""
    sysrc = vba("modSystem.bas")
    rel = re.search(r'EXPECTED_RELEASE As String = "([^"]+)"', sysrc).group(1)
    n_modules = int(re.search(r"N_MODULES As Long = (\d+)", sysrc).group(1))
    checks = re.findall(r"^\s*VerCheck\b", sysrc, re.M)
    assert len(checks) == n_modules, f"N_MODULES is {n_modules}, SystemCheck checks {len(checks)}"
    stamped = {}
    for p in sorted((ROOT / "modules").glob("*.bas")):
        m = re.search(r'_VERSION As String = "([^"]+)"', p.read_text(encoding="ascii"))
        if m:
            stamped[p.name] = m.group(1)
    assert len(stamped) == n_modules, f"{len(stamped)} stamped modules, N_MODULES {n_modules}"
    wrong = {k: v for k, v in stamped.items() if rel + " " not in v and not v.endswith(rel)}
    assert not wrong, f"stamps not on {rel}: {wrong}"


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


# ---------------------------------------------------------------- v27 (item 30, C8): more of the one-table-in-two-places rules
def test_holiday_names_match_vba():
    m = re.search(r"HolidayNames = Array\((.*?)\)\s*\n\s*End Function", vba("modHolidays.bas").replace(" _\r\n", " ").replace(" _\n", " "), re.S)
    assert m, "modHolidays.HolidayNames not found"
    names = re.findall(r'"([^"]*)"', m.group(1))
    assert names == rules.HOLIDAY_NAMES


def test_leave_codes_match_vba():
    m = re.search(r"Case\s+((?:\"[A-Z]+\",?\s*)+):\s*IsLeaveCode = True", vba("modCalendar.bas"))
    assert m, "modCalendar.IsLeaveCode not found"
    assert tuple(re.findall(r'"([^"]+)"', m.group(1))) == rules.LEAVE_CODES


def test_posting_constants_match_vba():
    close = vba_const("modPostings.bas", "POSTING_CLOSE_TIME").strip('"')
    hh = re.match(r"(\d+):(\d\d)\s*(AM|PM)", close, re.I)
    assert hh, close
    hour = int(hh.group(1)) % 12 + (12 if hh.group(3).upper() == "PM" else 0)
    assert hour == rules.POSTING_CLOSE_HOUR
    assert vba_const("modPostings.bas", "EXEMPT_TAG").strip('"') == rules.EXEMPT_TAG
    assert int(vba_const("modPostings.bas", "POSTING_CLOSE_DAYS")) == rules.POSTING_CLOSE_DAYS


def test_tunable_defaults_match_vba():
    assert int(vba_const("modOvertimeEntry.bas", "PAYBACK_WARN_DAYS")) == rules.PAYBACK_WARN_DAYS
    assert int(vba_const("modOvertimeEntry.bas", "SIGNUP_STALE_DAYS")) == rules.SIGNUP_STALE_DAYS


def test_sick_window_and_shift_starts_match_vba():
    """The sick shift + the 24 h after it, days from 07:00 and nights from 19:00: the same arithmetic both sides."""
    import inspect
    body = re.search(r"Public Function SickBlocks.*?End Function", vba("modLeave.bas"), re.S).group(0)
    for frag in ("19 / 24, 7 / 24", "sickStart + 0.5", "sickEnd + 1"):
        assert frag in body, f"modLeave.SickBlocks no longer has {frag!r}: update rules.sick_blocks with it"
    src = inspect.getsource(rules.sick_blocks)
    for frag in ("19 / 24 if sick_night else 7 / 24", "sick_start + 0.5", "sick_end + 1"):
        assert frag in src, f"rules.sick_blocks no longer has {frag!r}"

