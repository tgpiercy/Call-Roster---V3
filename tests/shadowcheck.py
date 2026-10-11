"""Find locals / parameters that hide a callable procedure (VBA is case-insensitive).
Usage: python shadowcheck.py <repo root>"""
import re
import sys
from pathlib import Path

root = Path(sys.argv[1])
files = sorted(list((root / "modules").glob("*.bas")) + list((root / "paste-in").glob("*.txt")) + list((root / "member").glob("*.bas")) + list((root / "member").glob("*.txt")))
PROC = re.compile(r"^\s*(public |private |friend )?(static )?(sub|function|property (get|let|set))\s+(\w+)\s*\((.*)\)", re.I)
END = re.compile(r"^\s*end (sub|function|property)\b", re.I)
DIM = re.compile(r"\b(dim|const|static)\s+(.*)", re.I)

public = {}      # name.lower() -> (file, line)
private = {}     # file -> {name.lower(): line}
bodies = []      # (file, procname, startline, params, body_lines)
for f in files:
    text = f.read_bytes().decode("latin-1").replace("\r\n", "\n")
    lines = text.split("\n")
    # join continuations for header parsing only
    cur = None
    private[f.name] = {}
    i = 0
    while i < len(lines):
        ln = lines[i]
        full = ln
        while full.rstrip().endswith("_") and i + 1 < len(lines):
            i += 1
            full = full.rstrip()[:-1] + " " + lines[i]
        m = PROC.match(full)
        if m:
            name = m.group(5)
            scope_priv = (m.group(1) or "").strip().lower() == "private"
            if scope_priv:
                private[f.name][name.lower()] = i + 1
            else:
                public.setdefault(name.lower(), (f.name, i + 1))
            params = [p.strip().split(" As ")[0].split(" as ")[0].strip() for p in m.group(6).split(",") if p.strip()]
            params = [re.sub(r"^(optional|byval|byref|paramarray)\s+", "", p, flags=re.I) for p in params]
            params = [re.sub(r"^(optional|byval|byref|paramarray)\s+", "", p, flags=re.I).replace("()", "").strip() for p in params]
            cur = (f.name, name, i + 1, params, [])
            bodies.append(cur)
        elif END.match(full):
            cur = None
        elif cur:
            cur[4].append((i + 1, full))
        i += 1

problems = []
for fname, pname, start, params, body in bodies:
    callable_here = dict(public)
    for n, l in private[fname].items():
        callable_here[n] = (fname, l)
    locals_ = {p.lower(): start for p in params if p}
    for n, ln in body:
        code = ln.split("'")[0]
        m = DIM.search(code)
        if m and not code.strip().lower().startswith("redim"):
            for part in m.group(2).split(","):
                v = part.strip().split(" ")[0].replace("(", "").strip()
                if v:
                    locals_[v.lower()] = n
    for v, ln in locals_.items():
        if v in callable_here and v != pname.lower():
            where = callable_here[v]
            # only matters if the hidden procedure is actually used inside this body
            used = any(re.search(r"\b" + re.escape(v) + r"\s*\(", b.split("'")[0], re.I) for _, b in body)
            problems.append((fname, ln, pname, v, where, used))

for p in sorted(problems, key=lambda x: (not x[5], x[0], x[1])):
    fname, ln, pname, v, where, used = p
    print(("USED  " if used else "quiet ") + f"{fname}:{ln} in {pname}: local/param '{v}' hides {where[0]}:{where[1]}")
print("done", len(problems), "shadowings")
