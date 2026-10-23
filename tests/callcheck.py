"""More of what the VBA compiler would refuse, found by reading the text.

    python tests/callcheck.py .

Checks, per project (master = modules/ + paste-in/, members = member/):
  DUPDECL   a name declared twice in one procedure, or a local with the
            same name as a parameter ("Duplicate declaration in current scope")
  LABEL     GoTo / Resume / On Error GoTo a label that isn't in the procedure
  DUPPROC   two procedures, or a procedure and a module-level name, with the
            same name in one module ("Ambiguous name detected")
  ARGS      a call to one of the project's own procedures with too few or
            too many arguments
  PARENS    a procedure called as a statement with several arguments inside
            parentheses:  Foo(a, b)   ("Expected: =")
  BYREF     a plain variable passed to a ByRef parameter of a different
            type ("ByRef argument type mismatch")
  FIELD     a field that the user-defined Type doesn't have
  NORET     a Function whose name is never assigned (it returns Empty /
            0 / "" / False on every path - "free", "0 charged")      (v25)
  RESTORE   Application.EnableEvents / ScreenUpdating set False, or
            Calculation set manual, in a procedure that never sets it
            back, itself or through a procedure it calls (ReportError
            doesn't count: it runs only on an error). A deliberate
            hand-off is marked on that line: ' restored by <who>       (v25)
  SET       an object variable (Worksheet, ListObject, Range, Object,
            Dictionary ...) assigned without Set ("Object variable not
            set" / "Invalid use of property" at run time)             (v25)
  MSGTITLE  a MsgBox (statement or function form) with fewer than three
            arguments: every box carries a title (v26.5.5, item 15; the
            default title is "Call Out Roster"; the members' file too since M4) (v26.5.5)
  MACROSTR  a "Module.Name" string (an OnAction, Application.Run, a Desk
            button) naming a procedure that module does not have, or a
            Private one; a bare OnAction / Application.Run "Name" that no
            module has as a Public procedure                          (v27)
  ORNSWALLOW an On Error Resume Next that runs to the end of the procedure
            (no later On Error GoTo, not inside an error handler) with no
            ' never raises marker in the procedure or the four lines above
            it: a quiet helper says so, anything else gets a handler (v27)
  PROCLEN   a procedure over 300 lines (KNOWN_LONG lists the three the
            v27.5 refactor splits)                                    (v27)
  DEADPUB   a Public procedure with parameters that nothing in the project
            names (an Alt+F8 macro has none, so it is never flagged)  (v27)
  DUPPUBLIC the same Public name (procedure, variable, Const, Type, Enum,
            Declare) in two standard modules of one project ("Ambiguous name
            detected" wherever it is called unqualified)              (v27.1)
  QUALREF   Module.Name in code where that standard module has no Name, or
            holds it Private ("Method or data member not found"): a call
            left pointing at the module a procedure moved out of      (v27.1)
            (MACROSTR, v27.1: also a "modX.Name" string whose module does
            not exist)

It is a text scan, not a compiler. Anything it can't parse with confidence
it skips, so a clean run is good evidence, not proof.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from undeclcheck import KEYWORDS, TOKEN, Module, split_statements  # noqa: E402

INTRINSIC = {"long", "integer", "string", "double", "single", "boolean", "date", "byte", "currency", "variant",
             "longlong", "longptr"}
# an assignment to NAME at the start of a statement, or after Then / Else on a one-line If
ASSIGN = r"(?:^\s*|\bThen\s+|\bElse\s+)(?:Set\s+)?{name}\s*=(?!=)"
PROC_MAX_LINES = 300
EXTERNAL_ENTRY = {"modsmoke"}     # v27: the Install panel (Roster Updater.xlsm, another project) runs these through Application.Run
KNOWN_LONG: set[str] = set()      # v27.1: empty - SyncKronosData, DrawCalendar and DrawNeeds are in named steps; every procedure is held to 300 lines


def swallowers(mod: Module) -> list[str]:
    """ORNSWALLOW: procedures whose last On Error Resume Next runs to the end, outside a handler, unmarked."""
    out = []
    for name, a, b in mod.procs:
        body = raw_proc_lines(mod, name)
        if not body:
            continue
        rn = [i for i, ln in enumerate(body) if re.search(r"(^|:)\s*On\s+Error\s+Resume\s+Next\b", ln.split("'")[0], re.I)]
        if not rn:
            continue
        last = rn[-1]
        on_err = {m.group(1).lower() for ln in body for m in [re.search(r"On\s+Error\s+GoTo\s+([A-Za-z_]\w*)", ln.split("'")[0], re.I)] if m}
        gotos = {m.group(1).lower() for ln in body for m in re.finditer(r"(?<![\w.])GoTo\s+([A-Za-z_]\w*)", ln.split("'")[0], re.I)}
        handler = False
        seen_exit = False
        for ln in body[1:last + 1]:
            if re.search(r"\bExit\s+(Sub|Function)\b", ln.split("'")[0], re.I):
                seen_exit = True
            mm = re.match(r"^\s*([A-Za-z_]\w*):\s*(?:'.*)?$", ln)
            if mm and (mm.group(1).lower() in on_err or (mm.group(1).lower() in gotos and seen_exit)):
                handler = True
        if handler:
            continue                                   # in an error handler's or a clean-up tail's code (review: a loop label is not one)
        if any(re.search(r"On\s+Error\s+GoTo\s", ln.split("'")[0], re.I) for ln in body[last + 1:]):
            continue
        key = str(mod.path)
        lines = _RAW.get(key) or mod.path.read_text(encoding="ascii", errors="replace").replace("\r\n", "\n").split("\n")
        head_i = next((i for i, ln in enumerate(lines) if ln == body[0]), 0)
        above = []                                     # the comment block right above the header, no further (review)
        k = head_i - 1
        while k >= 0 and lines[k].lstrip().startswith("'") and len(above) < 6:
            above.append(lines[k])
            k -= 1
        if any("never raises" in ln.lower() for ln in body + above):
            continue
        out.append(name)
    return out


OFF_ON = [  # (what, the 'off' value pattern, the 'on' value pattern)
    ("Application.EnableEvents", r"False", r"(?!False\b)\S"),
    ("Application.ScreenUpdating", r"False", r"(?!False\b)\S"),
    ("Application.Calculation", r"xlCalculationManual\b", r"(?!xlCalculationManual\b)\S"),
    # v26.5.5 (A3): Esc made trappable for a write must be handed back (xlInterrupt) on every exit path
    ("Application.EnableCancelKey", r"xlErrorHandler\b", r"(?!xlErrorHandler\b)\S"),
]


def msgbox_args(st: str) -> list[list[str]]:
    """The argument lists of every MsgBox call in one statement (strings masked by the caller's splitter):
    the statement form  MsgBox a, b  and the function form  x = MsgBox(a, b) ."""
    out: list[list[str]] = []
    masked = re.sub(r'"(?:[^"]|"")*"', lambda mm: "S" * len(mm.group(0)), st)
    for mm in re.finditer(r"(?<![\w.])MsgBox\b(\s*\(?)", masked, re.I):
        start = mm.end()
        if mm.group(1).strip() == "(":
            depth, i = 1, start
            while i < len(masked) and depth:
                depth += {"(": 1, ")": -1}.get(masked[i], 0)
                i += 1
            inner = st[start:i - 1]
        else:
            inner = st[start:]
            inner = re.split(r"\s+Then\b|\s+Else\b|:\s*(?=[A-Za-z_])", inner, maxsplit=1)[0]   # a one-line If ends the statement
        inner = inner.strip()
        out.append([] if inner == "" else [a for a in split_args(inner)])
    return out


def split_args(s: str) -> list[str]:
    parts, depth, cur = [], 0, []
    for ch in s:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    parts.append("".join(cur))
    if len(parts) == 1 and parts[0].strip() == "":
        return []
    return parts


def balanced(s: str, start: int) -> int:
    """Index just past the ')' matching the '(' at s[start] (-1 if none)."""
    depth = 0
    for j in range(start, len(s)):
        if s[j] == "(":
            depth += 1
        elif s[j] == ")":
            depth -= 1
            if depth == 0:
                return j + 1
    return -1


class Param:
    def __init__(self, text: str):
        t = text.strip()
        self.optional = bool(re.search(r"\bOptional\b", t, re.I))
        self.paramarray = bool(re.search(r"\bParamArray\b", t, re.I))
        self.byval = bool(re.search(r"\bByVal\b", t, re.I))
        t2 = re.sub(r"\b(Optional|ByVal|ByRef|ParamArray)\b", " ", t, flags=re.I)
        t2 = t2.split("=")[0]
        m = re.match(r"\s*([A-Za-z_]\w*)\s*(\(\s*\))?\s*(?:As\s+([A-Za-z_][\w.]*))?", t2, re.I)
        self.name = m.group(1) if m else "?"
        self.is_array = bool(m and m.group(2))
        self.type = (m.group(3) if m and m.group(3) else "Variant").lower()


class Proc:
    def __init__(self, mod: Module, name: str, a: int, b: int):
        self.mod, self.name, self.a, self.b = mod, name, a, b
        head = mod.lines[a]
        self.kind = "property" if re.search(r"\bProperty\b", head, re.I) else ("function" if re.search(r"\bFunction\b", head, re.I) else "sub")
        self.private = bool(re.match(r"^\s*Private\b", head, re.I))
        i = head.find("(")
        self.params: list[Param] = []
        if i >= 0:
            j = balanced(head, i)
            inner = head[i + 1:j - 1] if j > 0 else ""
            self.params = [Param(p) for p in split_args(inner)]
            self.body_first = head[j:] if j > 0 else ""
        else:
            self.body_first = ""
        self.required = sum(1 for p in self.params if not p.optional and not p.paramarray)
        self.maxargs = 10 ** 6 if any(p.paramarray for p in self.params) else len(self.params)


def decl_types(stmt: str) -> list[tuple[str, bool, str, str]]:
    """(name, is_array, type, kind) for each variable declared by a statement. kind: dim / redim / const."""
    m = re.match(r"^\s*(Public|Private|Global|Dim|Static|ReDim(?:\s+Preserve)?|Const)\b(.*)$", stmt, re.I)
    if not m:
        return []
    kw = m.group(1).lower()
    rest = m.group(2)
    kind = "redim" if kw.startswith("redim") else "dim"
    if re.match(r"^\s*Const\b", rest, re.I) or kw == "const":
        kind = "const"
        rest = re.sub(r"^\s*Const\b", "", rest, flags=re.I)
    rest = re.sub(r"^\s*WithEvents\b", "", rest, flags=re.I)
    out = []
    for piece in split_args(rest):
        mm = re.match(r"\s*([A-Za-z_]\w*)\s*(\(.*?\))?\s*(?:As\s+(?:New\s+)?([A-Za-z_][\w.]*))?", piece.split("=")[0] if kind != "redim" else piece, re.I)
        if mm:
            out.append((mm.group(1), bool(mm.group(2)), (mm.group(3) or "Variant").lower(), kind))
    return out


def proc_statements(mod: Module, p: Proc) -> tuple[list[str], set[str]]:
    stmts, labels = [], set()
    for k in range(p.a, p.b + 1):
        line = mod.lines[k]
        if k == p.a:
            line = re.sub(r"^\s*As\s+[A-Za-z_][\w.]*(\(\))?", "", p.body_first, flags=re.I).lstrip(": ")
        lab = re.match(r"^\s*([A-Za-z_]\w*):(?!=)\s*(.*)$", line)
        if lab and lab.group(1).lower() not in KEYWORDS:
            labels.add(lab.group(1).lower())
            line = lab.group(2)
        stmts.extend(st for st in split_statements(line))
    return stmts, labels


_RAW: dict[str, list[str]] = {}


def raw_proc_lines(mod: Module, name: str) -> list[str]:
    """The procedure's lines as written (comments kept), for markers such as ' restored by <who>."""
    key = str(mod.path)
    if key not in _RAW:
        _RAW[key] = mod.path.read_text(encoding="ascii", errors="replace").replace("\r\n", "\n").split("\n")
    lines = _RAW[key]
    head = re.compile(r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+" + re.escape(name) + r"\b", re.I)
    out, inside = [], False
    for ln in lines:
        if not inside and head.match(ln):
            inside = True
        if inside:
            out.append(ln)
            if re.match(r"^\s*End\s+(Sub|Function|Property)\b", ln, re.I) or re.search(r":\s*End\s+(Sub|Function)\s*(?:'.*)?$", ln, re.I):
                break
    return out


def collect_enums(mod: Module) -> set[str]:
    return {m.group(1).lower() for line in mod.lines
            for m in [re.match(r"^\s*(?:Public\s+|Private\s+)?Enum\s+([A-Za-z_]\w*)", line, re.I)] if m}


def is_object_type(ty: str, types: dict, enums: set[str]) -> bool:
    """A declared type that needs Set: not intrinsic, not a UDT, not an enum (project or Xl / Vb / Mso built-in)."""
    t = ty.lower()
    if t in INTRINSIC or t in types or t in enums or t == "variant":
        return False
    if re.match(r"^(xl|vb|mso|fm)[a-z]", t):            # XlCalculation, VbMsgBoxResult, MsoBarPosition, fmBorderStyle
        return False
    return True


def collect_types(mod: Module) -> dict[str, set[str]]:
    types, cur = {}, None
    for line in mod.lines:
        m = re.match(r"^\s*(?:Public\s+|Private\s+)?Type\s+([A-Za-z_]\w*)", line, re.I)
        if m:
            cur = m.group(1).lower()
            types[cur] = set()
            continue
        if cur:
            if re.match(r"^\s*End\s+Type\b", line, re.I):
                cur = None
            else:
                mm = TOKEN.search(line)
                if mm:
                    types[cur].add(mm.group(0).lower())
    return types


def check_project(files: list[Path], label: str) -> int:
    mods = [Module(p) for p in files]
    by_mod: dict[str, dict[str, Proc]] = {}
    pub: dict[str, Proc] = {}
    modnames = {m.name.lower() for m in mods}
    types: dict[str, set[str]] = {}
    enums: set[str] = set()
    mod_vars: dict[str, dict[str, tuple[bool, str]]] = {}
    restore: dict[tuple[str, str], tuple[set, set, set]] = {}     # (module, proc) -> (turned off, turned on, procs called)
    n = 0

    def say(tag, m, where, msg):
        nonlocal n
        print(f"{tag} {label} {m.path.name}: {where}: {msg}")
        n += 1

    for m in mods:
        types.update(collect_types(m))
        enums |= collect_enums(m)
        procs = {}
        for name, a, b in m.procs:
            p = Proc(m, name, a, b)
            if name.lower() in procs and p.kind != "property" and procs[name.lower()].kind != "property":
                say("DUPPROC", m, name, "defined twice in this module")
            procs[name.lower()] = p
            if not p.private and not m.is_class and p.kind != "property":
                pub.setdefault(name.lower(), p)
        by_mod[m.name.lower()] = procs
        mv = {}
        in_proc = False
        for line in m.lines:
            if re.match(r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property)\b", line, re.I):
                in_proc = True
            if in_proc:
                if re.match(r"^\s*End\s+(Sub|Function|Property)\b", line, re.I) or re.search(r":\s*End\s+(Sub|Function)\s*$", line, re.I):
                    in_proc = False
                continue
            if re.match(r"^\s*(Public|Private|Global|Dim|Const)\b", line, re.I) and not re.match(r"^\s*(Public|Private)\s+(Type|Enum|Declare)\b", line, re.I):
                for nm, arr, ty, kind in decl_types(line):
                    if nm.lower() in mv:
                        say("DUPPROC", m, nm, "module-level name declared twice")
                    mv[nm.lower()] = (arr, ty if kind != "const" else "const")
                    if nm.lower() in procs:
                        say("DUPPROC", m, nm, "is both a module-level name and a procedure in this module")
        mod_vars[m.name.lower()] = mv

    for m in mods:
        procs = by_mod[m.name.lower()]
        for p in procs.values():
            stmts, labels = proc_statements(m, p)
            local: dict[str, tuple[bool, str]] = {}
            for prm in p.params:
                local[prm.name.lower()] = (prm.is_array, prm.type)
            for st in stmts:
                for nm, arr, ty, kind in decl_types(st):
                    low = nm.lower()
                    if kind == "redim":
                        local.setdefault(low, (True, "variant"))
                        continue
                    if low in local:
                        say("DUPDECL", m, p.name, f"'{nm}' is declared twice (or is also a parameter)")
                    local[low] = (arr, ty if kind != "const" else "const")
            # labels
            for st in stmts:
                for mm in re.finditer(r"(?<![\w.])(?:GoTo|Resume)\s+([A-Za-z_]\w*)", st, re.I):     # not Application.Goto
                    tgt = mm.group(1).lower()
                    if tgt in ("next",) or tgt in labels:
                        continue
                    say("LABEL", m, p.name, f"label '{mm.group(1)}' doesn't exist in this procedure")

            def var_type(name: str):
                low = name.lower()
                if low in local:
                    return local[low]
                if low in mod_vars[m.name.lower()]:
                    return mod_vars[m.name.lower()][low]
                return None

            # ---- MSGTITLE (v26.5.5): every MsgBox names its title (the third argument) ----
            # (the members' workbook got its titles with Members M4: every project now)
            if True:
                for st in stmts:
                    for args in msgbox_args(st):
                        if len(args) < 3:
                            say("MSGTITLE", m, p.name, f"MsgBox with {len(args)} argument(s): give it a title (the third argument)")
            # ---- PROCLEN (v27): a procedure over PROC_MAX_LINES lines ----
            if p.b - p.a + 1 > PROC_MAX_LINES and p.name.lower() not in KNOWN_LONG:
                say("PROCLEN", m, p.name, f"{p.b - p.a + 1} lines (over {PROC_MAX_LINES}): split it")
            # ---- NORET: a Function that never assigns its own name (v25) ----
            if p.kind == "function":
                pat = re.compile(ASSIGN.format(name=re.escape(p.name)), re.I)
                if not any(pat.search(st) for st in stmts):
                    say("NORET", m, p.name, "Function never assigns its result (returns Empty / 0 / \"\" / False everywhere)")
            # ---- RESTORE (v25): what this procedure turns off / on, and whom it calls; judged after every procedure is read ----
            offs, ons = set(), set()
            raw_lines = raw_proc_lines(m, p.name)
            for what, off, on in OFF_ON:
                w = re.escape(what)
                off_pat = re.compile(ASSIGN.format(name=w) + r"\s*" + off, re.I)
                on_pat = re.compile(ASSIGN.format(name=w) + r"\s*" + on, re.I)
                handed_off = any(off_pat.search(ln) and re.search(r"'.*\brestored by\b", ln, re.I) for ln in raw_lines)
                if any(off_pat.search(st) for st in stmts) and not handed_off:
                    offs.add(what)
                if any(on_pat.search(st) for st in stmts):
                    ons.add(what)
            called = set()
            if offs:
                for st in stmts:
                    for mm in TOKEN.finditer(st):
                        nm = mm.group(0).lower()
                        if nm != p.name.lower() and nm not in ("reporterror", "logerror"):
                            callee = procs.get(nm) or pub.get(nm)
                            if callee is not None:
                                called.add((callee.mod.name.lower(), callee.name.lower()))
            restore[(m.name.lower(), p.name.lower())] = (offs, ons, called)
            # ---- SET: an object variable assigned without Set (v25) ----
            for st in stmts:
                if re.match(r"^\s*(Public|Private|Global|Dim|Static|Const|ReDim)\b", st, re.I):
                    continue
                for mm in re.finditer(r"(?:^\s*|\bThen\s+|\bElse\s+)([A-Za-z_]\w*)\s*=(?!=)", st, re.I):
                    nm = mm.group(1)
                    if nm.lower() in KEYWORDS or nm.lower() == p.name.lower():
                        continue
                    vt = var_type(nm)
                    if vt and not vt[0] and is_object_type(vt[1], types, enums):
                        say("SET", m, p.name, f"'{nm}' (As {vt[1]}) is assigned without Set")

            def resolve(name: str, qualifier: str | None):
                low = name.lower()
                if qualifier:
                    q = qualifier.lower()
                    if q in by_mod and not q.startswith("frm") and q != "thisworkbook":
                        return by_mod[q].get(low)
                    return None
                if low in local or low in mod_vars[m.name.lower()]:
                    return None                        # a variable, not a call
                if low in procs:
                    return procs[low]
                return pub.get(low)

            for st in stmts:
                if re.match(r"^\s*(Public|Private|Global|Dim|Static|Const|ReDim|End|Exit|On\s+Error|GoTo|Resume)\b", st, re.I):
                    if not re.match(r"^\s*(Dim|Static|ReDim)\b", st, re.I):
                        continue
                # ---- UDT fields ----
                for mm in re.finditer(r"(?<![\w.])([A-Za-z_]\w*)\s*(\([^()]*(?:\([^()]*\)[^()]*)*\))?\.([A-Za-z_]\w*)", st):
                    vt = var_type(mm.group(1))
                    if vt and vt[1] in types and mm.group(3).lower() not in types[vt[1]]:
                        say("FIELD", m, p.name, f"'{mm.group(1)}.{mm.group(3)}': Type {vt[1]} has no field '{mm.group(3)}'")
                # ---- calls ----
                for mm in TOKEN.finditer(st):
                    name = mm.group(0)
                    a, b = mm.start(), mm.end()
                    qualifier = None
                    if a > 0 and st[a - 1] == ".":
                        q = re.search(r"([A-Za-z_]\w*)\.$", st[:a])
                        if not q:
                            continue
                        qualifier = q.group(1)
                    if a > 0 and st[a - 1] == "!":
                        continue
                    callee = resolve(name, qualifier)
                    if callee is None or callee.kind == "property":
                        continue
                    after = st[b:]
                    prefix = st[:a if not qualifier else a - len(qualifier) - 1]
                    at_start = bool(re.match(r"^\s*(?:Call\s+)?$", prefix, re.I)) or bool(re.search(r"\b(?:Then|Else)\s+(?:Call\s+)?$", prefix, re.I))
                    args = None
                    if re.match(r"\s*\(", after):
                        i = b + after.find("(")
                        j = balanced(st, i)
                        if j < 0:
                            continue
                        inner = st[i + 1:j - 1]
                        args = split_args(inner)
                        tail = st[j:].strip()
                        is_call_kw = bool(re.search(r"\bCall\s+$", prefix, re.I))
                        if at_start and not is_call_kw and len(args) > 1 and tail == "":
                            say("PARENS", m, p.name, f"'{name}(...)' with {len(args)} arguments is called as a statement: drop the parentheses or use Call")
                        if at_start and not is_call_kw and tail.startswith(","):
                            continue                   # Foo (a), b  - rare, skip
                    elif at_start:
                        rest = re.split(r"\s+Else\b", after, maxsplit=1, flags=re.I)[0].strip()     # If c Then Foo Else Bar
                        if rest.startswith("=") or rest.startswith("."):
                            continue                   # the function's own return value (assigned, or used as an object)
                        args = split_args(rest)
                    else:
                        if callee.name.lower() == p.name.lower():
                            continue                   # reading its own return value
                        args = []
                    if any(re.match(r"\s*[A-Za-z_]\w*\s*:=", x) for x in args):
                        continue                       # named arguments: not position-checked
                    if not (callee.required <= len(args) <= callee.maxargs):
                        if callee.name.lower() == p.name.lower() and len(args) == 0:
                            continue
                        say("ARGS", m, p.name, f"'{name}' is called with {len(args)} argument(s); it takes "
                            f"{callee.required}" + ("" if callee.required == callee.maxargs else f" to {callee.maxargs if callee.maxargs < 10**6 else 'any number'}"))
                        continue
                    for k, arg in enumerate(args):
                        if k >= len(callee.params):
                            break
                        prm = callee.params[k]
                        if prm.byval or prm.paramarray:
                            continue
                        am = re.match(r"^\s*([A-Za-z_]\w*)\s*$", arg)
                        if not am:
                            continue
                        vt = var_type(am.group(1))
                        if vt is None or vt[1] == "const":
                            continue
                        v_arr, v_ty = vt
                        if prm.type == "variant" and not prm.is_array:
                            continue                   # a ByRef Variant takes anything
                        if prm.is_array and not v_arr and v_ty == "variant":
                            continue                   # a Variant holding an array
                        if prm.is_array != v_arr:
                            say("BYREF", m, p.name, f"'{am.group(1)}' passed to {callee.name}'s ByRef parameter '{prm.name}': one is an array, the other isn't")
                            continue
                        if v_ty != prm.type and (v_ty in INTRINSIC or prm.type in INTRINSIC or v_ty in types or prm.type in types):
                            say("BYREF", m, p.name, f"'{am.group(1)}' (As {v_ty}) passed to {callee.name}'s ByRef parameter '{prm.name}' (As {prm.type})")

    # ---- ORNSWALLOW (v27) ----
    for m in mods if label not in ("member", "members", "updater") else []:     # the master's rule (the other two install apart)
        for name in swallowers(m):
            say("ORNSWALLOW", m, name, "an On Error Resume Next runs to the end: add a handler, or mark it ' never raises")
    # ---- MACROSTR / DEADPUB (v27): every string naming a procedure, and every Public with parameters, is matched ----
    texts = {m.name.lower(): m.path.read_text(encoding="ascii", errors="replace") for m in mods}
    std_mods = {m.name.lower() for m in mods if m.path.suffix.lower() == ".bas" and not m.is_class}
    for m in mods:
        for mm in re.finditer(r'"([A-Za-z_]\w*)\.([A-Za-z_]\w*)"', texts[m.name.lower()]):
            mod_l, proc_l = mm.group(1).lower(), mm.group(2).lower()
            if proc_l in ("bas", "txt", "cls", "frm", "xlsm", "xlsx"):
                continue                               # a file name
            if mod_l not in std_mods:
                if re.match(r"^mod[a-z]", mod_l) and mod_l not in modnames:       # v27.1: a module that isn't there
                    say("MACROSTR", m, mm.group(0), f"there is no module {mm.group(1)} in this project")
                continue                               # "Scripting.Dictionary", "WScript.Shell", a form's handler label
            tgt = by_mod[mod_l].get(proc_l)
            if tgt is None:
                say("MACROSTR", m, mm.group(0), f"{mm.group(1)} has no procedure {mm.group(2)}")
            elif tgt.private:
                say("MACROSTR", m, mm.group(0), f"{mm.group(1)}.{mm.group(2)} is Private: a button or Application.Run can't reach it")
        for mm in re.finditer(r'(?:OnAction\s*:?=\s*|Application\.Run\s*\(?\s*)"([A-Za-z_]\w*)"', texts[m.name.lower()], re.I):
            nm = mm.group(1).lower()
            if nm not in pub:
                say("MACROSTR", m, mm.group(1), "no module has a Public procedure of that name")
    def no_comments(line: str) -> str:
        inside = False
        for i, ch in enumerate(line):
            if ch == '"':
                inside = not inside
            elif ch == "'" and not inside:
                return line[:i]
        return line

    # (review) words outside comments, and outside the procedure's own body (its return-value assignments)
    code_lines = {m.name.lower(): [no_comments(ln) for ln in m.path.read_text(encoding="ascii", errors="replace").replace("\r\n", "\n").split("\n")] for m in mods}
    for m in mods:
        if m.is_class or m.path.suffix.lower() != ".bas" or m.name.lower() in EXTERNAL_ENTRY:
            continue
        for p in by_mod[m.name.lower()].values():
            if p.private or p.kind == "property" or not p.params:
                continue
            own = {ln for ln in raw_proc_lines(m, p.name)}
            pat = re.compile(r"(?<![\w.])" + re.escape(p.name) + r"\b", re.I)
            used = False
            for mm in mods:
                for ln in code_lines[mm.name.lower()]:
                    if pat.search(ln) and not (mm is m and ln in {no_comments(x) for x in own}):
                        used = True
                        break
                if used:
                    break
            if not used:
                say("DEADPUB", m, p.name, "a Public procedure nothing calls: Private, or delete it")
    # ---- DUPPUBLIC (v27.1): one Public name in two standard modules ----
    owners: dict[str, list[str]] = {}
    for m in mods:
        if m.is_class or m.path.suffix.lower() != ".bas":
            continue
        for nm_ in m.public_names:
            owners.setdefault(nm_, []).append(m.name)
    for nm_, where in sorted(owners.items()):
        if len(where) > 1:
            first = next(x for x in mods if x.name == where[0])
            say("DUPPUBLIC", first, nm_, f"Public in {len(where)} modules ({', '.join(sorted(where))}): Ambiguous name")
    # ---- QUALREF (v27.1): Module.Name in code must name a member that module has and shows ----
    std_by = {x.name.lower(): x for x in mods if not x.is_class and x.path.suffix.lower() == ".bas"}
    for m in mods:
        for p in by_mod[m.name.lower()].values():
            stmts, _ = proc_statements(m, p)
            for st in stmts:
                code = re.sub(r'"(?:[^"]|"")*"', '""', st)
                for mm in re.finditer(r"(?<![\w.!])([A-Za-z_]\w*)\.([A-Za-z_]\w*)", code):
                    tgt = std_by.get(mm.group(1).lower())
                    if tgt is None or mm.group(1).lower() == mm.group(2).lower():
                        continue
                    mem = mm.group(2).lower()
                    if mem not in tgt.module_names:
                        say("QUALREF", m, p.name, f"{mm.group(0)}: {tgt.name} has no {mm.group(2)} (moved?)")
                    elif mem not in tgt.public_names and tgt is not m:
                        say("QUALREF", m, p.name, f"{mm.group(0)}: {mm.group(2)} is Private to {tgt.name}")
    # ---- RESTORE verdicts: a restore in a procedure this one calls (any depth) counts ----
    def restores(key, what, seen):
        if key in seen:
            return False
        seen.add(key)
        offs, ons, called = restore.get(key, (set(), set(), set()))
        if what in ons:
            return True
        return any(restores(c, what, seen) for c in called)

    for m in mods:
        for p in by_mod[m.name.lower()].values():
            key = (m.name.lower(), p.name.lower())
            offs = restore.get(key, (set(), set(), set()))[0]
            for what in sorted(offs):
                if not restores(key, what, set()):
                    say("RESTORE", m, p.name, f"{what} is turned off here and never turned back on (by this procedure or one it calls)")
    return n


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    master = sorted((root / "modules").glob("*.bas")) + sorted((root / "paste-in").glob("*.txt"))
    member = sorted((root / "member").glob("*"))
    updater = sorted((root / "updater").glob("*.bas"))      # v24.2: its own workbook, so its own project
    n = check_project(master, "master") + check_project(member, "member") + check_project(updater, "updater")
    print(f"done {n} finding(s)")
    return 1 if n else 0


if __name__ == "__main__":
    sys.exit(main())
