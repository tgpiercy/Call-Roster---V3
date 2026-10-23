"""v27.1 B3 commit 2: Private -> Public on exactly the names movecheck's SCOPE listed after the modPostings split.

    py -3.14 tests/refactor/b3_promote.py .

Each name must be found exactly once as a Private procedure header in its module; only the keyword changes.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

PROMOTE = {
    "modPostings": ["GetTable", "ToDate", "SplitExempt", "AppendRows"],
    "modPostingSheet": ["PostingSheet", "RenderPostingSheet"],
    "modPostingSync": ["UpdatePostingRecords"],
    "modPostingMail": ["EmailAward"],
}


def main(root: Path) -> int:
    for mod, names in PROMOTE.items():
        path = root / "modules" / f"{mod}.bas"
        lines = path.read_bytes().decode("ascii").split("\r\n")
        for nm in names:
            pat = re.compile(r"^Private((?:\s+Static)?\s+(?:Sub|Function)\s+" + nm + r"\b)", re.I)
            hits = [i for i, ln in enumerate(lines) if pat.match(ln)]
            if len(hits) != 1:
                print(f"{mod}.{nm}: {len(hits)} Private header(s), expected 1 - refused")
                return 1
            i = hits[0]
            lines[i] = pat.sub(r"Public\1", lines[i], count=1)
            print(f"{mod}.bas:{i + 1}: {nm} -> Public")
        path.write_bytes("\r\n".join(lines).encode("ascii"))
    return 0


if __name__ == "__main__":
    sys.exit(main(Path(sys.argv[1] if len(sys.argv) > 1 else ".")))
