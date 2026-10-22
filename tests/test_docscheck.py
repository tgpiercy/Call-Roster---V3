"""docscheck fires on each kind of stale line and stays quiet on a current one (v27 item 30, C9)."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import docscheck  # noqa: E402

FACTS = {"release": "v27", "modules": 30, "steps": 22, "cases": 860, "suites": 98}

CURRENT = """Current as of v27 (Nov 20, 2026).
SystemCheck says 30 of 30 on v27.
SelfTest: 860 cases in 98 suites.
SmokeTest runs 22 steps.
"""

STALE = """Current as of v26.5.5 (Oct 21, 2026).
SystemCheck says 29 of 29 on v26.5.5.
SelfTest: 649 cases in 75 suites.
SmokeTest runs 20 steps (v26.5).
The calendar shows 14 steps ahead.
"""


def test_a_current_doc_is_clean():
    assert docscheck.doc_problems(CURRENT, "fixture", FACTS) == []


def test_each_stale_line_is_found_once():
    out = docscheck.doc_problems(STALE, "fixture", FACTS)
    kinds = sorted(p.split()[0] for p in out)
    assert kinds == ["CASES", "MODULES", "RELEASE", "STEPS"], out      # "14 steps" is not on a SmokeTest line


def test_the_facts_come_from_the_code():
    facts = docscheck.code_facts(Path(__file__).resolve().parent.parent)
    assert facts["steps"] >= 22 and facts["modules"] >= 30 and facts["cases"] > 800
