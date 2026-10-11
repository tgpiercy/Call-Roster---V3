"""Print rotation / crewshift / avail / eligibility rows for cases.csv.

These rows are REGRESSION values derived from the mirror in rules.py (the
cycle strings come straight from modShiftCycle). The anchor facts that are
known independently - Dec 28, 2025 is B days / A nights - are asserted in
test_cases.py. Run and paste the output into cases.csv when adding dates.
"""
from __future__ import annotations

from datetime import date, timedelta

import rules

DATES = [
    date(2025, 12, 28), date(2025, 12, 29), date(2025, 12, 31), date(2026, 1, 1),
    date(2026, 1, 24), date(2026, 1, 25), date(2026, 2, 14), date(2026, 3, 7),
    date(2026, 5, 18), date(2026, 7, 1), date(2026, 9, 7), date(2026, 10, 12),
    date(2026, 10, 31), date(2026, 11, 1), date(2026, 12, 25), date(2026, 12, 26),
    date(2026, 12, 31), date(2027, 1, 1), date(2027, 3, 13), date(2027, 7, 1),
    date(2027, 12, 31), date(2028, 2, 29),
]


def main() -> None:
    for d in DATES:
        print(f"rotation,{d.isoformat()},{d.isoformat()},,,,,{rules.days_crew(d)}/{rules.nights_crew(d)}")
    for d in (date(2026, 10, 12), date(2026, 10, 10), date(2026, 10, 31)):
        for code in ("A", "B", "C", "D", "A/B", "C/D", "DAYS"):
            on = rules.crew_shift_on(code, d) or "off"
            print(f"crewshift,{code.replace('/', '-')}-{d.isoformat()},{code},{d.isoformat()},,,,{on}")
    for d in (date(2026, 10, 30), date(2026, 10, 31)):
        for sh in ("D", "N"):
            print(f"avail,{d.isoformat()}-{sh},{d.isoformat()},{sh},,,,{rules.available_crew(d, sh == 'N')}")
    # eligibility: one crew per rotation state on Oct 31, 2026 (B days / D nights, A/C off)
    d = date(2026, 10, 31)
    print("# Oct 31 2026:", rules.days_crew(d), "days", rules.nights_crew(d), "nights; Oct 30:",
          rules.days_crew(d - timedelta(1)), rules.nights_crew(d - timedelta(1)), "; Nov 1:",
          rules.days_crew(d + timedelta(1)), rules.nights_crew(d + timedelta(1)))


if __name__ == "__main__":
    main()
