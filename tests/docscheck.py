"""The end-user docs say what the code says (v27 item 30, C9).

    python tests/docscheck.py .

For the four end-user docs (the supervisor quick guide, the system manual, the admin install guide, the
rollover one-pager):
  RELEASE   "Current as of vNN" names modSystem.EXPECTED_RELEASE
  MODULES   "N of N on vNN" (the SystemCheck line) uses N_MODULES and the release
  CASES     "N cases in M suites" is tests/cases.csv's count of cases and suites
  STEPS     "N steps" on a SmokeTest line is modSmokeTest.SMK_STEPS

The docs are brought up to the release in the release-prep commit, so this is run by tests/build_zip.py
(it refuses to build on a finding), not by every pytest run; tests/test_docscheck.py proves it fires.
"""
from __future__ import annotations

import csv
import re
import sys
from pathlib import Path

DOCS = ("docs/supervisor-quick-guide.md", "docs/system-manual.md", "docs/admin-install-guide.md", "docs/rollover-one-pager.md")
# v29.0.3 (review 2.7): the alpha's plan, its step table and its reset runbook name the alpha build (v29.0.x) - every
# three-part release they mention must be the code's (the real file's "v29" is left alone)
ALPHA_DOCS = ("docs/alpha-test-plan.md", "docs/alpha-data-reset.md", "tests/mock/steps.csv")


def code_facts(root: Path) -> dict:
    sys_txt = (root / "modules" / "modSystem.bas").read_text(encoding="ascii")
    smoke = (root / "modules" / "modSmokeTest.bas").read_text(encoding="ascii")
    release = re.search(r'EXPECTED_RELEASE As String = "([^"]+)"', sys_txt).group(1)
    n_mod = int(re.search(r"N_MODULES As Long = (\d+)", sys_txt).group(1))
    steps = int(re.search(r"SMK_STEPS As Long = (\d+)", smoke).group(1))
    with (root / "tests" / "cases.csv").open(newline="", encoding="utf-8") as f:
        rows = [r for r in csv.DictReader(f) if r.get("suite")]
    return {"release": release, "modules": n_mod, "steps": steps, "cases": len(rows), "suites": len({r["suite"] for r in rows})}


def doc_problems(text: str, name: str, facts: dict) -> list[str]:
    out = []
    for m in re.finditer(r"Current as of (v\d+(?:\.\d+)*)", text):
        if m.group(1) != facts["release"]:
            out.append(f"RELEASE {name}: 'Current as of {m.group(1)}' - the code is {facts['release']}")
    for m in re.finditer(r"\b(\d+) of (\d+) on (v\d+(?:\.\d+)*)", text):
        if int(m.group(1)) != facts["modules"] or int(m.group(2)) != facts["modules"] or m.group(3) != facts["release"]:
            out.append(f"MODULES {name}: '{m.group(0)}' - the code is {facts['modules']} of {facts['modules']} on {facts['release']}")
    for m in re.finditer(r"\b(\d+) cases in (\d+) suites", text):
        if int(m.group(1)) != facts["cases"] or int(m.group(2)) != facts["suites"]:
            out.append(f"CASES {name}: '{m.group(0)}' - the code has {facts['cases']} cases in {facts['suites']} suites")
    for line in text.splitlines():
        if re.search(r"smoke ?test", line, re.I):
            for m in re.finditer(r"\b(\d+) steps\b", line):
                if int(m.group(1)) != facts["steps"]:
                    out.append(f"STEPS {name}: '{m.group(0)}' on a SmokeTest line - SMK_STEPS is {facts['steps']}")
    return out


def alpha_problems(text: str, name: str, facts: dict) -> list[str]:
    """Every vNN.N.N token of the code's major.minor must be the release (v29.0.1 left in a plan step after v29.0.3)."""
    rel = facts["release"]
    parts = rel.split(".")
    if len(parts) < 3:
        return []
    out = []
    for m in re.finditer(r"\bv\d+\.\d+\.\d+\b", text):
        if m.group(0).split(".")[:2] == parts[:2] and m.group(0) != rel:
            out.append(f"ALPHA {name}: '{m.group(0)}' - the alpha build is {rel}")
    return out


def check(root: Path) -> list[str]:
    facts = code_facts(root)
    out = []
    for rel in DOCS:
        p = root / rel
        if p.is_file():
            out += doc_problems(p.read_text(encoding="utf-8"), rel, facts)
    for rel in ALPHA_DOCS:
        p = root / rel
        if p.is_file():
            out += alpha_problems(p.read_text(encoding="utf-8"), rel, facts)
    return out


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    probs = check(root)
    for p in probs:
        print(p)
    print(f"done {len(probs)} finding(s)")
    return 1 if probs else 0


if __name__ == "__main__":
    sys.exit(main())
