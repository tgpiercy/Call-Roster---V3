"""Paste tests/mock/*.csv into tests/mock/modMockData.bas as VBA data (P9, the loader's data block).

    py -3.14 tests/mock/gen_mock_loader.py

It rewrites the block between the MOCK DATA markers: one Private Sub per table (chunked so no procedure passes VBA's
64 KB), one MockRow "a|b|c" line per CSV row (fields joined with "|"), and the LoadRows dispatcher. CRLF, ASCII only; a
field holding "|" or a double quote, or a non-ASCII character, stops the run. tests/test_mock.py checks the pasted
block equals the CSVs (drift) and that the module compiles against the v29 modules (undeclcheck + callcheck).
"""
from __future__ import annotations

import csv
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
BAS = HERE / "modMockData.bas"
START = "' >>> MOCK DATA (generated from tests/mock/*.csv by tests/mock/gen_mock_loader.py - do not edit by hand)"
END = "' <<< MOCK DATA"
CHUNK = 120

# key -> (csv file, the columns in the order the loader's WriteTable types list them)
TABLES = {
    "people": ("people.csv", None),            # + the 17th field: the Oct 18 snapshot's Worked Hours (hours_history.csv)
    "log": ("log_opening.csv", None),
    "crew": ("crew_history.csv", None),
    "absences": ("absences.csv", None),
    "coverage": ("coverage.csv", None),
    "switches": ("switches.csv", None),
    "postings": ("postings.csv", None),
    "postinglist": ("posting_list.csv", None),
    "planned": ("planned_moves.csv", None),
    "restrictions": ("restrictions.csv", None),
    "hourshistory": ("hours_history.csv", None),
    "idmarks": ("idmarks.csv", None),
    "expected": ("expected_load.csv", None),
}


def rows_of(key: str) -> list[str]:
    name, _ = TABLES[key]
    with (HERE / name).open(newline="", encoding="ascii") as f:
        rows = [list(r.values()) for r in csv.DictReader(f)]
    if key == "people":
        h17 = {}
        with (HERE / "hours_history.csv").open(newline="", encoding="ascii") as f:
            for r in csv.DictReader(f):
                if r["Data Through"] == "2026-10-17":
                    h17[r["Employee"]] = r["Worked Hours"]
        rows = [r + [h17[r[2]]] for r in rows]
    out = []
    for r in rows:
        for v in r:
            if "|" in v or '"' in v or any(ord(ch) > 126 for ch in v):
                raise SystemExit(f"{key}: a field holds | or a quote or a non-ASCII character: {v!r}")
        out.append("|".join(r))
    return out


def block() -> list[str]:
    lines = [START, "Private Sub LoadRows(ByVal key As String)", "    Select Case key"]
    subs = []
    for key in TABLES:
        rows = rows_of(key)
        chunks = [rows[i:i + CHUNK] for i in range(0, len(rows), CHUNK)] or [[]]
        names = [f"Data_{key}_{k + 1}" for k in range(len(chunks))]
        lines.append(f'        Case "{key}": ' + ": ".join(names))
        for nm, ch in zip(names, chunks):
            subs.append(f"Private Sub {nm}()")
            for row in ch:
                subs.append(f'    MockRow "{row}"')
            subs.append("End Sub")
            subs.append("")
    lines.append("    End Select")
    lines.append("End Sub")
    lines.append("")
    lines.extend(subs)
    if lines[-1] == "":
        lines.pop()
    lines.append(END)
    return lines


def main() -> int:
    text = BAS.read_text(encoding="ascii").replace("\r\n", "\n")
    a = text.index(START)
    b = text.index(END) + len(END)
    new = text[:a] + "\n".join(block()) + text[b:]
    data = new.replace("\n", "\r\n").encode("ascii")
    BAS.write_bytes(data)
    n = sum(1 for ln in new.splitlines() if ln.startswith('    MockRow "'))
    print(f"{BAS.name}: {n} data rows, {len(data)} bytes")
    return 0


if __name__ == "__main__":
    sys.exit(main())
