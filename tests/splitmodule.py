"""Cut whole sections of one module into new modules: the move of the v27.1 refactor, made by script, never retyped.

    py -3.14 tests/splitmodule.py ROOT FROM NEW=a-b[,c-d] [NEW2=e-f ...]
    py -3.14 tests/splitmodule.py . modCalendar modBoard=3520-3730 modNeeds=4319-5393

Line numbers are 1-based, inclusive, and ALL refer to FROM as it is before the run (the cuts are made together).
Each NEW module gets  Attribute VB_Name = "NEW" / Option Explicit / a blank line, then its ranges in file order.
A NEW file that already exists is refused. A range must start and end OUTSIDE a procedure (a banner or blank line is
fine) and must not overlap another range. The files are written CRLF, ASCII, as the VBA editor needs them.

Then prove it:  py -3.14 tests/movecheck.py <old ref>   (SCOPE lists the Private names the cut now reaches across)
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

HEAD = re.compile(r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+", re.I)
END = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)
ONE_LINE = re.compile(r":\s*End\s+(?:Sub|Function|Property)\s*(?:'.*)?$", re.I)


def proc_spans(lines: list[str]) -> list[tuple[int, int]]:
    """(first, last) 1-based line of every procedure."""
    out, start = [], 0
    for i, ln in enumerate(lines, 1):
        if not start and HEAD.match(ln) and not re.search(r"\bDeclare\b", ln, re.I):
            if ONE_LINE.search(ln):
                out.append((i, i))
            else:
                start = i
            continue
        if start and END.match(ln):
            out.append((start, i))
            start = 0
    return out


def main(argv: list[str]) -> int:
    if len(argv) < 3:
        print(__doc__)
        return 2
    root, src = Path(argv[0]), argv[1]
    path = root / "modules" / f"{src}.bas"
    raw = path.read_bytes().decode("ascii")
    lines = raw.split("\r\n")
    trailing = lines and lines[-1] == ""
    if trailing:
        lines.pop()
    spans = proc_spans(lines)
    plan: list[tuple[str, list[tuple[int, int]]]] = []
    taken: list[tuple[int, int, str]] = []
    for spec in argv[2:]:
        name, _, ranges = spec.partition("=")
        if not re.match(r"^[A-Za-z_]\w*$", name) or not ranges:
            print(f"bad spec {spec!r}: NEW=a-b[,c-d]")
            return 2
        if (root / "modules" / f"{name}.bas").exists():
            print(f"modules/{name}.bas already exists: refused")
            return 1
        rs = []
        for r in ranges.split(","):
            a, _, b = r.partition("-")
            a, b = int(a), int(b)
            if not (1 <= a <= b <= len(lines)):
                print(f"{name}: range {a}-{b} is outside 1-{len(lines)}")
                return 1
            for s, e in spans:
                if (s < a <= e) or (s <= b < e):
                    print(f"{name}: {a}-{b} cuts the procedure at lines {s}-{e}: a range holds whole procedures")
                    return 1
            for x, y, who in taken:
                if a <= y and x <= b:
                    print(f"{name}: {a}-{b} overlaps {who} {x}-{y}")
                    return 1
            taken.append((a, b, name))
            rs.append((a, b))
        plan.append((name, sorted(rs)))
    cut = set()
    for name, rs in plan:
        body = []
        for a, b in rs:
            body += lines[a - 1:b]
            cut |= set(range(a - 1, b))
        text = [f'Attribute VB_Name = "{name}"', "Option Explicit", ""] + body
        (root / "modules" / f"{name}.bas").write_bytes(("\r\n".join(text) + "\r\n").encode("ascii"))
        print(f"modules/{name}.bas: {len(body)} lines from {src} {', '.join(f'{a}-{b}' for a, b in rs)}")
    keep = [ln for i, ln in enumerate(lines) if i not in cut]
    path.write_bytes(("\r\n".join(keep) + ("\r\n" if trailing else "")).encode("ascii"))
    print(f"modules/{src}.bas: {len(lines)} -> {len(keep)} lines")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
