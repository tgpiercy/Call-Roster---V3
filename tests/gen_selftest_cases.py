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
CHUNK = 200      # cases per LoadCasesN (about 16 KB each; the limit is 64 KB a procedure)


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
    # v27 (review C4): LoadCases reached 57 KB of VBA's 64 KB procedure limit, so the cases are emitted in
    # chunks of CHUNK - LoadCases calls LoadCases1..N, each one a Sub of its own - all inside the ONE marker pair
    # (tests/test_drift.py reads every AddCase line between the markers and ignores the Sub frames).
    chunks = [lines[i:i + CHUNK] for i in range(0, len(lines), CHUNK)] or [[]]
    body = [f"    LoadCases{i + 1}" for i in range(len(chunks))]
    for i, chunk in enumerate(chunks):
        body.append("End Sub")
        body.append("")
        body.append(f"Private Sub LoadCases{i + 1}()")
        body.extend(chunk)
    new = text[: a + len(START)] + "\n" + "\n".join(body) + "\n" + text[b:]
    BAS.write_bytes(new.replace("\n", "\r\n").encode("ascii"))
    print(f"wrote {len(lines)} cases into {BAS.relative_to(ROOT)} in {len(chunks)} chunk(s) of up to {CHUNK}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
