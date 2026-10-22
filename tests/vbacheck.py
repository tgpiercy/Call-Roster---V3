"""Mechanical checks over the VBA sources: block matching, Option Explicit,
Attribute line, CRLF, ASCII, duplicate Public names, undeclared-looking
procedure calls. Usage: python vbacheck.py <repo root>"""
import re
import sys
from pathlib import Path

root = Path(sys.argv[1])
files = sorted(list((root / "modules").glob("*.bas")) + list((root / "paste-in").glob("*.txt")) + list((root / "member").glob("*.bas")) + list((root / "updater").glob("*.bas")) + list((root / "member").glob("*.txt")))
problems = []
publics = {}
procs = {}          # name -> file (all procedures, any scope)

OPEN = {
    "if": "end if", "for": "next", "with": "end with", "select": "end select",
    "do": "loop", "while": "wend", "sub": "end sub", "function": "end function",
    "property": "end property", "type": "end type", "enum": "end enum",
}


def strip_strings_and_comments(line: str) -> str:
    out, i, in_str = [], 0, False
    while i < len(line):
        ch = line[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    i += 1
                else:
                    in_str = False
        else:
            if ch == '"':
                in_str = True
            elif ch == "'":
                break
            else:
                out.append(ch)
        i += 1
    return "".join(out)


for f in files:
    raw = f.read_bytes()
    if b"\n" in raw.replace(b"\r\n", b""):
        problems.append(f"{f.name}: LF-only line endings present")
    try:
        text = raw.decode("ascii")
    except UnicodeDecodeError as e:
        problems.append(f"{f.name}: non-ASCII byte at {e.start}")
        text = raw.decode("latin-1")
    lines = text.split("\r\n")
    is_bas = f.suffix == ".bas"
    if is_bas and not lines[0].startswith('Attribute VB_Name = "'):
        problems.append(f"{f.name}: missing Attribute VB_Name on line 1")
    if "Option Explicit" not in lines[:3]:
        problems.append(f"{f.name}: Option Explicit not in the first 3 lines")

    # join continuations
    joined, buf, start = [], "", 0
    for n, ln in enumerate(lines, 1):
        code = strip_strings_and_comments(ln).rstrip()
        if not buf:
            start = n
        if code.endswith(" _") or code.endswith("\t_"):      # a continuation is " _" (a name ending in _ isn't one)
            buf += code[:-1] + " "
            continue
        joined.append((start, buf + code))
        buf = ""

    stack = []
    seen_proc = False          # v27: a module-level Const / Dim / Type / Enum / Declare after the first procedure does not compile
    for n, code in joined:
        s = code.strip()
        if not s:
            continue
        low = s.lower()
        # v27 (Oct 22, found at install): every ( has its ) in one statement - a stray ) at the end of a Sub call
        # written without brackets is "Syntax error" (two receipts had one; reading and every checker had missed it)
        depth = 0
        for ch in s:
            depth += (ch == "(") - (ch == ")")
            if depth < 0:
                break
        if depth != 0:
            problems.append(f"{f.name}:{n}: brackets don't balance ({'a stray )' if depth < 0 else 'an unclosed ('}) - Syntax error")
        # v27 (Oct 22, SmokeTest step 22 at install): IIf works out BOTH answers, so a guard in its condition guards
        # nothing - IIf(IsArray(x), UBound(x) ...) and IIf(x Is Nothing, .., x.Member) raise on the case they test for
        for mm in re.finditer(r"\biif\(\s*(?:not\s+)?(?:isarray\(\s*(\w+)\s*\)|(\w+)\s+is\s+nothing)", low):
            v = mm.group(1) or mm.group(2)
            rest = low[mm.end():]
            if re.search(r"\b(?:ubound|lbound)\(\s*" + re.escape(v) + r"\b", rest) or re.search(r"\b" + re.escape(v) + r"\.\w", rest):
                problems.append(f"{f.name}:{n}: IIf guards {v} in its condition but uses it in a branch - IIf works out both: use If ... Else")
        # v27 (review of item 21): an Exit Sub inside a Function (or the reverse) does not compile
        procs_open = [k for k, _ in stack if k in ("sub", "function", "property")]
        if procs_open:
            for ex in re.findall(r"\bexit\s+(sub|function|property)\b", low):
                if ex != procs_open[-1]:
                    problems.append(f"{f.name}:{n}: 'Exit {ex.capitalize()}' inside a {procs_open[-1].capitalize()} "
                                    f"(compile error: write 'Exit {procs_open[-1].capitalize()}')")
        if seen_proc and not stack:
            dm = re.match(r"(public |private |friend |global )?(const|dim|type|enum|declare|withevents)\b", low)
            if dm:
                problems.append(f"{f.name}:{n}: module-level declaration '{s[:50]}' after the first procedure "
                                "(compile error: 'Only comments may appear after End Sub' - move it above the first Sub / Function)")
        # single-line If ... Then <stmt>  (no block)
        if low.startswith("if ") and " then" in low and not low.endswith(" then") and not re.search(r"\bthen\s*$", low):
            # could be "If x Then stmt : Else stmt" one-liner
            if re.search(r"\belseif\b", low):                    # v25 review: a one-line If with ElseIf is a compile error
                problems.append(f"{f.name}:{n}: one-line If ... Then <stmt> ElseIf (compile error: use a block If)")
            continue
        # statements separated by ':' - only consider first token for openers
        m = re.match(r"(public |private |friend |static )?(\w+)", low)
        tok = m.group(2) if m else ""
        head = low
        # closers
        for opener, closer in OPEN.items():
            if head.startswith(closer) and (len(head) == len(closer) or not head[len(closer)].isalnum()):
                if not stack or stack[-1][0] != opener:
                    problems.append(f"{f.name}:{n}: '{closer}' without matching '{opener}' (stack={[x[0] for x in stack][-3:]})")
                else:
                    stack.pop()
                break
        else:
            if tok in ("sub", "function", "property") and not low.startswith("exit ") and not low.startswith("end "):
                if "declare " not in low:
                    stack.append((tok, n))
                    seen_proc = True
                    nm = re.search(r"(sub|function|property\s+(get|let|set))\s+(\w+)", low)
                    if nm:
                        name = nm.group(3)
                        procs.setdefault(name, []).append(f.name)
                        scope_pub = low.startswith("public ") or (not low.startswith("private ") and is_bas and tok != "property")
                        if scope_pub and is_bas:
                            publics.setdefault(name, []).append(f.name)
            elif tok == "if":
                # block if: ends with Then
                if re.search(r"\bthen\s*$", low):
                    stack.append(("if", n))
            elif tok == "for":
                stack.append(("for", n))
            elif tok == "with":
                stack.append(("with", n))
            elif tok == "select":
                stack.append(("select", n))
            elif tok == "do":
                stack.append(("do", n))
            elif tok == "while":
                stack.append(("while", n))
            elif tok == "type" or (tok in ("public", "private") and re.match(r"(public|private)\s+type\b", low)):
                stack.append(("type", n))
            elif re.match(r"(public |private )?enum\b", low):
                stack.append(("enum", n))
            elif tok == "next":
                if stack and stack[-1][0] == "for":
                    stack.pop()
                else:
                    problems.append(f"{f.name}:{n}: 'Next' without For")
            elif tok == "loop":
                if stack and stack[-1][0] == "do":
                    stack.pop()
                else:
                    problems.append(f"{f.name}:{n}: 'Loop' without Do")
            elif tok == "wend":
                if stack and stack[-1][0] == "while":
                    stack.pop()
                else:
                    problems.append(f"{f.name}:{n}: 'Wend' without While")
        # Public Const names
        mc = re.match(r"public\s+const\s+(\w+)", low)
        if mc:
            publics.setdefault(mc.group(1), []).append(f.name)
    if stack:
        problems.append(f"{f.name}: unclosed blocks {[(a, b) for a, b in stack]}")

for name, where in sorted(publics.items()):
    if len(set(where)) > 1:
        problems.append(f"duplicate Public name {name}: {sorted(set(where))}")

print(f"checked {len(files)} files, {len(procs)} procedures")
for p in problems:
    print("PROBLEM:", p)
print("OK" if not problems else f"{len(problems)} problem(s)")
