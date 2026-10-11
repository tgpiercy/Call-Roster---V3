"""Paste cases.csv into modSelfTest.bas as VBA constants.

Run after editing cases.csv:
    python tests/gen_selftest_cases.py

It rewrites the block between the CASES markers in modules/modSelfTest.bas
(one AddCase line per CSV row, CRLF line endings, ASCII only) so the
workbook's SelfTest and the Python tests share one table.
"""
from __future__ import annotations

import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tests" / "cases.csv"
BAS = ROOT / "modules" / "modSelfTest.bas"
START = "' >>> CASES (generated from tests/cases.csv by tests/gen_selftest_cases.py - do not edit by hand)"
END = "' <<< CASES"
FIELDS = ["suite", "case", "in1", "in2", "in3", "in4", "in5", "expected"]


def main() -> int:
    with CSV.open(newline="") as f:
        rows = list(csv.DictReader(f))
    lines = []
    for r in rows:
        vals = [r[k] for k in FIELDS]
        for v in vals:
            if "," in v or '"' in v or any(ord(ch) > 126 for ch in v):
                raise SystemExit(f"bad character in case {r['suite']}:{r['case']}: {v!r}")
        lines.append('    AddCase "' + ",".join(vals) + '"')
    text = BAS.read_text(encoding="ascii").replace("\r\n", "\n")
    a = text.index(START)
    b = text.index(END)
    new = text[: a + len(START)] + "\n" + "\n".join(lines) + "\n" + text[b:]
    BAS.write_bytes(new.replace("\n", "\r\n").encode("ascii"))
    print(f"wrote {len(lines)} cases into {BAS.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
