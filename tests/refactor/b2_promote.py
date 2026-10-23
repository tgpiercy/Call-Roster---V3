"""v27.1 builder B2, commit 2: Private -> Public on exactly what movecheck's SCOPE listed after the move
(tests/refactor/b2_move.cmd). Decided by Graeme: state Public as-is, no setters, no renames.

    py -3.14 tests/refactor/b2_promote.py

Each target is found by its declaration (not by line number), must be found exactly once, and only the
leading keyword changes. Files stay CRLF + ASCII.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]

TARGETS = [
    ("modPeople", "mBusy", r"Private\s+mBusy\s+As\s+Boolean\b"),
    ("modPeople", "BodyOf", r"Private\s+Function\s+BodyOf\("),
    ("modPeople", "QuietOn", r"Private\s+Sub\s+QuietOn\("),
    ("modPeople", "QuietOff", r"Private\s+Sub\s+QuietOff\("),
    ("modPeople", "AfterPeopleChange", r"Private\s+Sub\s+AfterPeopleChange\("),
    ("modPeople", "AddRosterRow", r"Private\s+Function\s+AddRosterRow\("),
    ("modPlannedMoves", "CancelPlannedCore", r"Private\s+Function\s+CancelPlannedCore\("),
]


def main() -> int:
    files: dict[str, list[str]] = {}
    for mod, _, _ in TARGETS:
        if mod not in files:
            raw = (ROOT / "modules" / f"{mod}.bas").read_bytes().decode("ascii")
            files[mod] = raw.split("\r\n")
    for mod, name, pat in TARGETS:
        rx = re.compile(r"^" + pat, re.I)
        lines = files[mod]
        hits = [i for i, ln in enumerate(lines) if rx.match(ln)]
        if len(hits) != 1:
            print(f"{mod}: {name}: {len(hits)} declarations found, expected 1 - nothing written")
            return 1
        i = hits[0]
        lines[i] = "Public" + lines[i][len("Private"):]
        print(f"{mod}.bas:{i + 1}: {name} -> Public")
    for mod, lines in files.items():
        (ROOT / "modules" / f"{mod}.bas").write_bytes("\r\n".join(lines).encode("ascii"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
