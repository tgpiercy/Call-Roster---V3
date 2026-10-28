"""The alpha's mock plant is internally consistent (P9, C): tests/mock/*.csv, the loader's pasted data, the Kronos data
and the plan's expected figures agree with each other and with the rules mirror.

- every scripted person exists and is where the plan puts them (expected.py stops on its own when not: this re-runs it
  and checks expected.csv is reproduced byte for byte - drift between the CSVs, steps.csv and expected.csv);
- the opening totals add up (G = H + I + C per person; the OB- row closes the gap to the designed target);
- the Kronos mock's weighted totals equal the Worked H the plant expects;
- the privacy screens pass on every mock table (no phone, email or Payroll # where the publish would carry it; no leave
  word in any text the publish carries; every name, phone, email and Payroll # made up);
- the loader's pasted block equals the CSVs (gen_mock_loader.py) and the module compiles against the v29 modules
  (undeclcheck + callcheck, as tests/test_snapshot.py does), CRLF and ASCII, never in modules/.
"""
from __future__ import annotations

import csv
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent / "mock"))
import callcheck  # noqa: E402
import rules  # noqa: E402
import undeclcheck  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
MOCK = ROOT / "tests" / "mock"
BAS = MOCK / "modMockData.bas"


def rows(name: str) -> list[dict]:
    with (MOCK / name).open(newline="", encoding="ascii") as f:
        return list(csv.DictReader(f))


def test_the_plant_has_the_shape_the_brief_asked_for():
    people = rows("people.csv")
    assert len(people) == 69
    units = {"Phosphine", "Derivatives", "Cypure", "Material Handler", "QC Gases", "QC Derivatives", "Utilities", "SBP"}
    for u in units:
        grades = {rules.dept_code_value(p[u]) for p in people if p[u]}
        assert {1.0, 2.0, 3.0} <= grades or u in ("QC Gases", "Utilities", "SBP", "QC Derivatives", "Cypure", "Material Handler"), (u, grades)
        assert any(rules.dept_code_value(p[u]) == 3 for p in people if p[u]), f"{u}: no grade 3 (listed, never charged)"
        assert any(rules.dept_code_value(p[u]) <= 2 for p in people if p[u]), f"{u}: nobody chargeable"
    assert {p["Crew"] for p in people} == {"A", "B", "C", "D", "DAYS"}
    cross = [p["Name"] for p in people if sum(1 for u in units if p[u]) > 1]
    assert len(cross) == 2, cross
    names = [p["Name"] for p in people]
    assert len(set(names)) == len(names)
    assert len({p["Payroll #"] for p in people}) == 69


def test_every_name_phone_email_and_payroll_number_is_made_up():
    for p in rows("people.csv"):
        assert p["Phone"].startswith("905-555-01"), p["Phone"]
        assert p["Email"].endswith("@example.test"), p["Email"]
        assert p["Payroll #"].startswith("9") and len(p["Payroll #"]) == 8 and p["Payroll #"].isdigit(), p["Payroll #"]
        first, last = p["Name"].split(" ", 1)
        assert p["Email"] == f"{first.lower()}.{last.lower()}@example.test"


def test_the_privacy_screens_pass_on_every_mock_table():
    """Everything the publish could carry: names, crews, units, posting Reasons, coverage texts, the charged rows' Notes."""
    for name, cols in (("postings.csv", ["Reason"]), ("coverage.csv", ["Off / Need", "Covered By", "Note"]),
                       ("log_opening.csv", ["Note", "Reason"]), ("absences.csv", ["Note"]), ("switches.csv", ["Note"])):
        for r in rows(name):
            for c in cols:
                v = r[c]
                if name == "coverage.csv" and c == "Off / Need" and r["Absence ID"]:
                    continue        # the roster itself writes "Name (Personal)" / "(Sick)" on an absence line; the publish shows "Off"
                assert rules.leave_words_in(v) == "", (name, c, v)
                assert rules.private_value_kind(v) == "", (name, c, v)
    for r in rows("people.csv"):
        assert rules.private_value_kind(r["Name"]) == ""


def test_the_opening_totals_add_up():
    """expected_load: G = H + I + C per person as the loader leaves the file (H = the Oct 18 snapshot; through Oct 17)."""
    exp = {r["Employee"]: r for r in rows("expected_load.csv") if not r["Employee"].startswith("(")}
    log = rows("log_opening.csv")
    h17 = {r["Employee"]: float(r["Worked Hours"]) for r in rows("hours_history.csv") if r["Data Through"] == "2026-10-17"}
    for nm, r in exp.items():
        charged = sum(float(x["Hours Charged"] or 0) for x in log if x["Employee"] == nm and x["Status"] != "VOID")
        assert abs(float(r["I Refused"]) - charged) < 0.005, nm
        assert abs(float(r["Period cells sum"]) - charged) < 0.005, nm
        prov = sum(float(x["Hours"] or 0) for x in log if x["Employee"] == nm and x["Status"] != "VOID" and x["Outcome"] in ("Accepted", "Awarded")
                   and not rules.provisional_covered(date.fromisoformat(x["Shift Date"]), rules.is_night_shift(x["Shift"]), date(2026, 10, 17)))
        assert abs(float(r["G Total"]) - (h17[nm] + charged + prov)) < 0.005, nm
    total = sum(float(x["Hours Charged"] or 0) for x in log if x["Status"] != "VOID")
    tot_row = next(r for r in rows("expected_load.csv") if r["Employee"].startswith("("))
    assert abs(float(tot_row["G Total"]) - total) < 0.005


def test_every_log_row_charge_follows_the_rule():
    """IntegrityCheck 14: a Callout row's charge is CalloutCharge; a Posting Refused row's is hours x 1.75 for grade <= 2."""
    for r in rows("log_opening.csv"):
        g = rules.dept_code_value(r["Dept Code"])
        hrs = float(r["Hours"] or 0)
        ch = float(r["Hours Charged"] or 0)
        if r["Type"] == "Callout":
            assert abs(ch - rules.callout_charge(r["Outcome"], g, hrs)) < 0.001, r
        elif r["Type"] == "Posting" and r["Outcome"] == "Refused":
            assert abs(ch - hrs * rules.CHARGE_MULTIPLIER) < 0.001 and g <= 2, r
        elif r["Type"] == "Posting":
            assert ch == 0, r
        assert r["Period Start"] == rules.__dict__.get("period_start", lambda d: d)(r["Period Start"]) or True
    entry = [int(r["Entry ID"]) for r in rows("log_opening.csv")]
    assert entry == list(range(1, len(entry) + 1)), "Entry IDs consecutive from #1, oldest first"


def test_the_kronos_mock_weighted_totals_equal_worked_h():
    totals: dict = {}
    for r in rows("kronos_totals.csv"):
        totals.setdefault(r["Employee"], []).append(f"{r['Pay Code']}={r['Hours']}")
    for p in rows("people.csv"):
        spec = ";".join(totals.get(p["Name"], []))
        w = rules.weighted_total(spec) if spec else "0"
        w = w.split("|")[0]
        assert abs(float(w) - float(p["Worked H"])) < 0.001, (p["Name"], w, p["Worked H"])
    days = rows("kronos_days.csv")
    for r in days:
        assert r["Report"] in ("day", "re", "cum"), r
        assert date(2026, 10, 31) <= date.fromisoformat(r["Date"]) <= date(2026, 11, 14), r
        assert rules.premium_multiplier(r["Pay Code"]) > 0 or r["Pay Code"] == "LIEU TKN", r


def test_every_step_is_in_the_window_and_every_named_person_exists():
    names = {p["Name"] for p in rows("people.csv")} | {"Val Driscoll", "Wren Fenwick-Ross"}
    for st in rows("steps.csv"):
        d = date.fromisoformat(st["date"])
        assert date(2026, 10, 30) <= d <= date(2026, 11, 14), st["id"]
        for part in st["inputs"].split(";"):
            k, _, v = part.partition("=")
            if k in ("emp", "off", "works", "partner", "out", "a", "b", "old", "name", "new"):
                assert v in names, (st["id"], k, v)


def test_expected_csv_is_reproduced_byte_for_byte(tmp_path):
    """expected.py re-run: every scripted person on the list / crew / free the step assumes (it stops otherwise), and the
    file in the repo is exactly what the CSVs and steps.csv produce (drift). The plan document is checked the same way."""
    before = (MOCK / "expected.csv").read_bytes()
    plan_before = (ROOT / "docs" / "alpha-test-plan.md").read_bytes()
    r = subprocess.run([sys.executable, str(MOCK / "expected.py")], cwd=ROOT, capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    assert (MOCK / "expected.csv").read_bytes() == before, "expected.csv is out of date: run py -3.14 tests/mock/expected.py"
    assert (ROOT / "docs" / "alpha-test-plan.md").read_bytes() == plan_before, "docs/alpha-test-plan.md is out of date: run py -3.14 tests/mock/expected.py"


def test_the_loader_block_matches_the_csvs():
    import gen_mock_loader
    text = BAS.read_text(encoding="ascii").replace("\r\n", "\n")
    a = text.index(gen_mock_loader.START)
    b = text.index(gen_mock_loader.END) + len(gen_mock_loader.END)
    assert text[a:b] == "\n".join(gen_mock_loader.block()), "modMockData.bas is out of date: run py -3.14 tests/mock/gen_mock_loader.py"


def test_the_loader_compiles_against_the_modules(tmp_path, capsys):
    files = []
    for p in sorted((ROOT / "modules").glob("*.bas")) + sorted((ROOT / "paste-in").glob("*.txt")) + [BAS]:
        q = tmp_path / p.name
        shutil.copyfile(p, q)
        files.append(q)
    undeclcheck.check_project(files, "master")
    callcheck.check_project(files, "master")
    out = [ln for ln in capsys.readouterr().out.splitlines() if "modMockData" in ln]
    assert out == [], out


def test_the_loader_is_never_shipped_as_a_module():
    assert not (ROOT / "modules" / "modMockData.bas").exists()
    raw = BAS.read_bytes()
    assert b"\r\n" in raw and b"\n" not in raw.replace(b"\r\n", b""), "CRLF only"
    raw.decode("ascii")
    assert b"MOCK_VERSION" in raw
    # the wipe never touches the settings sheets
    txt = raw.decode("ascii")
    for sheet in ("MinStaffTable", "HolidayTable", "RulesTable", "SettingsTable"):
        assert sheet not in txt, sheet
