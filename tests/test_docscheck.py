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


def test_an_old_alpha_build_left_in_the_plan_is_found():
    """v29.1 (review 2.7): a three-part release of the code's major.minor that is not the code's; v29 alone is fine."""
    facts = dict(FACTS, release="v29.1")
    out = docscheck.alpha_problems("installed v29.0.1 through the panel; the real file stays on v29; v28.1 is later", "fixture", facts)
    assert [p.split()[0] for p in out] == ["ALPHA"] and "v29.0.1" in out[0], out
    assert docscheck.alpha_problems("v29.1 on the mock file, v29 on the real one", "fixture", facts) == []
    assert docscheck.alpha_problems("v29.0.1 anywhere", "fixture", FACTS) == []      # a one-part release (v27) checks nothing
    facts = dict(FACTS, release="v29.1")
    out = docscheck.alpha_problems("renamed v29.1 from v29.0.3; v29 on the real file; v28.1 later", "fixture", facts)
    assert [p.split()[0] for p in out] == ["ALPHA"] and "v29.0.3" in out[0], out
