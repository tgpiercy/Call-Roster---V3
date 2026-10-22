"""Find identifiers used in a procedure that are declared nowhere, and
declared names that are VBA reserved words.

    python tests/undeclcheck.py .

Option Explicit makes every one of these a "Variable not defined" compile
error, and nothing else in tests/ catches it (vbacheck checks blocks,
shadowcheck checks locals hiding procedures). It is a text scan, not a
compiler: it knows VBA keywords, the common VBA / Excel built-ins and the
vb* / xl* / mso* / fm* constant families, and treats everything else that
is not declared in the procedure, the module or as a Public name in the
same project as a problem.

Two projects are checked separately: the master (modules/ + paste-in/)
and the members' workbook (member/).
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

KEYWORDS = set("""
and as boolean byref byval byte call case const currency date dim do double each else elseif empty end enum
erase error exit false for function get global goto if in integer is let like long longlong longptr loop me mod
new next not nothing null object on option optional or paramarray preserve private property public redim resume
select set single static step stop string sub then to true type typeof until variant wend while with withevents xor
explicit declare lib alias debug print open close input output append binary random access read write line name shared lock
attribute friend implements event raiseevent addressof gosub return rem
""".split())

BUILTINS = set(x.lower() for x in """
Abs Array Asc AscW CBool CByte CCur CDate CDbl CInt CLng CSng CStr CVar CVErr Choose Chr ChrW Cos CreateObject
GetSetting SaveSetting DeleteSetting GetAllSettings CurDir DateAdd DateDiff DatePart DateSerial DateValue Day Dir DoEvents Environ EOF Err Exp FileDateTime FileLen
Fix Format FreeFile GetAttr GetObject Hex Hour IIf InStr InStrRev InputBox Int IsArray IsDate IsEmpty IsError
IsMissing IsNull IsNumeric IsObject Join LBound LCase Left Len Log LTrim Mid Minute MkDir Month MonthName MsgBox
Now Replace RGB Right Rnd Round RTrim Second SetAttr Sgn Shell Sin Space Split Sqr Str StrComp StrConv Tan Time
Timer TimeSerial TimeValue Trim TypeName UBound UCase Val VarType Weekday WeekdayName Year Load Unload Beep
Randomize Kill FileCopy RmDir ChDir ChDrive LOF Loc Seek
Application ThisWorkbook ActiveWorkbook ActiveSheet ActiveCell ActiveWindow Selection Workbooks Worksheets Sheets
Range Cells Columns Rows Names Intersect Union Evaluate UserForms Windows WorksheetFunction MSForms VBA Excel
Count Item
""".split())

# Words VBA will not accept as the name of a variable, parameter, constant or procedure
# ("Dim single As Long" is a syntax error: Single is a data type).
RESERVED = set("""
call case close const declare dim do else elseif end endif enum erase event exit for friend function get global
gosub goto if implements input let lock loop lset next on open option print private public put raiseevent redim
resume return rset seek select set static stop sub type unlock wend while with write rem
any as byref byval each in new shared until withevents optional paramarray preserve spc tab then to
addressof and eqv imp is like mod not or typeof xor
abs cbool cbyte ccur cdate cdbl cdec cint clng clnglng clngptr csng cstr cvar cverr date debug doevents fix int
len lenb me pset scale sgn string array circle inputb lbound ubound
boolean byte currency double integer long longlong longptr single variant
true false nothing empty null cdecl decimal defbool defbyte defint deflng defcur defsng defdbl defdec defdate
defstr defobj defvar
""".split())

CONST_FAMILY = re.compile(r"^(vb|xl|mso|fm)[A-Z0-9]", re.I)
TOKEN = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
DESIGNER_CONTROL = re.compile(r"^(btn|lst|cbo|txt|lbl|chk|opt|fra|cmd)[A-Z]")
PROC = re.compile(r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\s*(\(.*)?$", re.I)
END_PROC = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
DECL = re.compile(r"^\s*(?:(Public|Private|Global|Dim|Static)\s+)?(?:(Const)\s+)?(?:WithEvents\s+)?(.*)$", re.I)


def strip_code(text: str) -> list[str]:
    """Logical lines with strings and comments removed (line continuations joined)."""
    out, buf = [], ""
    for raw in text.replace("\r\n", "\n").split("\n"):
        line, in_str, res = raw, False, []
        for ch in line:
            if ch == '"':
                in_str = not in_str
                res.append('"')
                continue
            if in_str:
                continue
            if ch == "'":
                break
            res.append(ch)
        s = "".join(res).rstrip()
        if re.match(r"^\s*Rem\b", s, re.I):
            s = ""
        if s.endswith(" _"):
            buf += s[:-1] + " "
            continue
        out.append(buf + s)
        buf = ""
    if buf:
        out.append(buf)
    return out


def split_statements(line: str) -> list[str]:
    """Split on ':' outside parentheses (labels are handled by the caller)."""
    parts, depth, cur = [], 0, []
    for j, ch in enumerate(line):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == ":" and depth == 0 and line[j + 1:j + 2] != "=":      # ":=" is a named argument, not a separator
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    parts.append("".join(cur))
    return parts


def declared_names(stmt: str) -> list[str]:
    """Names declared by a Dim / Static / Const / Public / Private / ReDim statement."""
    m = re.match(r"^\s*(Public|Private|Global|Dim|Static|ReDim(?:\s+Preserve)?|Const)\b(.*)$", stmt, re.I)
    if not m:
        return []
    rest = m.group(2)
    rest = re.sub(r"^\s*Const\b", "", rest, flags=re.I)
    rest = re.sub(r"^\s*WithEvents\b", "", rest, flags=re.I)
    names, depth, cur = [], 0, []
    pieces = []
    for ch in rest:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            pieces.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    pieces.append("".join(cur))
    for p in pieces:
        mm = TOKEN.search(p)
        if mm:
            names.append(mm.group(0))
    return names


def param_names(sig: str) -> list[str]:
    inner = sig[sig.find("(") + 1:] if "(" in sig else ""
    depth, cur, pieces = 0, [], []
    for ch in inner:
        if ch == "(":
            depth += 1
        elif ch == ")":
            if depth == 0:
                break
            depth -= 1
        if ch == "," and depth == 0:
            pieces.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    pieces.append("".join(cur))
    names = []
    for p in pieces:
        p = re.sub(r"\b(Optional|ByVal|ByRef|ParamArray)\b", " ", p, flags=re.I)
        mm = TOKEN.search(p)
        if mm:
            names.append(mm.group(0))
    return names


def used_names(stmt: str) -> list[str]:
    s = re.sub(r"\bAs\s+(New\s+)?[A-Za-z_][\w.]*", " ", stmt, flags=re.I)      # types
    s = re.sub(r"\bNew\s+[A-Za-z_][\w.]*", " ", s, flags=re.I)
    s = re.sub(r"\bTypeOf\s+(\S+)\s+Is\s+[A-Za-z_][\w.]*", r" \1 ", s, flags=re.I)
    s = re.sub(r"#[^#\n]*#", " ", s)                                           # date literals
    s = re.sub(r"&H[0-9A-Fa-f]+&?", " ", s)
    s = re.sub(r"\b\d+(\.\d+)?([eE][+-]?\d+)?[#!@&%]?", " ", s)
    out = []
    for m in TOKEN.finditer(s):
        a, b = m.start(), m.end()
        if a > 0 and s[a - 1] in ".!":
            continue                                   # member access
        if re.match(r"\s*:=", s[b:]):
            continue                                   # named argument
        out.append(m.group(0))
    return out


class Module:
    def __init__(self, path: Path):
        self.path = path
        self.name = path.stem
        self.is_class = path.suffix.lower() == ".txt"       # forms / ThisWorkbook
        self.lines = strip_code(path.read_text(encoding="ascii", errors="replace"))
        self.module_names: set[str] = set()
        self.module_level_decls: set[str] = set()      # variables / constants only (checked against RESERVED)
        self.public_names: set[str] = set()
        self.procs: list[tuple[str, int, int]] = []
        self._scan()

    def _scan(self):
        in_type = in_proc = False
        start = 0
        pname = ""
        for i, line in enumerate(self.lines):
            if re.match(r"^\s*Attribute\b", line, re.I) or re.match(r"^\s*Option\b", line, re.I):
                continue
            if not in_proc:
                m = re.match(r"^\s*(?:(Public|Private)\s+)?(Type|Enum)\s+([A-Za-z_]\w*)", line, re.I)
                if m:
                    in_type = m.group(2).lower()
                    self.module_names.add(m.group(3).lower())
                    if (m.group(1) or "public").lower() == "public":
                        self.public_names.add(m.group(3).lower())
                    self._type_public = (m.group(1) or "public").lower() == "public"
                    continue
                if in_type:
                    if re.match(r"^\s*End\s+(Type|Enum)\b", line, re.I):
                        in_type = False
                    elif in_type == "enum":
                        mm = TOKEN.search(line)
                        if mm:
                            self.module_names.add(mm.group(0).lower())
                            if self._type_public:
                                self.public_names.add(mm.group(0).lower())
                    continue
                m = re.match(r"^\s*(?:(Public|Private)\s+)?Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+([A-Za-z_]\w*)", line, re.I)
                if m:
                    self.module_names.add(m.group(2).lower())
                    if (m.group(1) or "public").lower() == "public":
                        self.public_names.add(m.group(2).lower())
                    continue
                m = PROC.match(line)
                if m:
                    in_proc, start, pname = True, i, m.group(3)
                    self.module_names.add(pname.lower())
                    if (m.group(1) or "public").lower() != "private" and not self.is_class:
                        self.public_names.add(pname.lower())
                    if END_PROC.search(line.split(":")[-1]) and ":" in line:      # one-line Sub X(): ...: End Sub
                        self.procs.append((pname, start, i))
                        in_proc = False
                    continue
                m = re.match(r"^\s*(Public|Private|Global|Dim|Const)\b", line, re.I)
                if m:
                    for st in split_statements(line):
                        for n in declared_names(st):
                            self.module_names.add(n.lower())
                            self.module_level_decls.add(n.lower())
                            if m.group(1).lower() in ("public", "global") and not self.is_class:
                                self.public_names.add(n.lower())
            else:
                if END_PROC.match(line):
                    self.procs.append((pname, start, i))
                    in_proc = False


def check_project(files: list[Path], label: str) -> int:
    mods = [Module(p) for p in files]
    globals_ = set()
    for m in mods:
        globals_ |= m.public_names
        globals_.add(m.name.lower())
    problems = 0
    for m in mods:
        for n in sorted(m.module_level_decls):
            if n in RESERVED:
                print(f"RESERVED {label} {m.path.name}: module-level name '{n}' is a reserved word")
                problems += 1
    for m in mods:
        for pname, a, b in m.procs:
            local = {pname.lower()}
            head = m.lines[a]
            local |= {n.lower() for n in param_names(head)}
            body = m.lines[a:b + 1]
            stmts = []
            for k, line in enumerate(body):
                if k == 0:
                    # a one-line procedure keeps its body after the signature's closing parenthesis
                    depth, cut = 0, len(line)
                    for j, ch in enumerate(line):
                        if ch == "(":
                            depth += 1
                        elif ch == ")":
                            depth -= 1
                            if depth == 0:
                                cut = j + 1
                                break
                    rest = re.sub(r"^\s*As\s+[A-Za-z_][\w.]*(\(\))?", "", line[cut:], flags=re.I)
                    line = rest.lstrip(": ")
                lab = re.match(r"^\s*([A-Za-z_]\w*):(?!=)\s*(.*)$", line)
                if lab and lab.group(1).lower() not in KEYWORDS:
                    local.add(lab.group(1).lower())
                    line = lab.group(2)
                stmts.extend(split_statements(line))
            for st in stmts:
                for n in declared_names(st):
                    local.add(n.lower())
            for n in sorted(local):
                if n in RESERVED and n != pname.lower():
                    print(f"RESERVED {label} {m.path.name}: {pname}: '{n}' is a reserved word and can't be a variable or parameter name")
                    problems += 1
            if pname.lower() in RESERVED:
                print(f"RESERVED {label} {m.path.name}: procedure name '{pname}' is a reserved word")
                problems += 1
            for st in stmts:
                if re.match(r"^\s*(On\s+Error\s+)?(GoTo|Resume)\s+", st, re.I) or re.match(r"^\s*End\b", st, re.I):
                    continue
                for n in used_names(st):
                    low = n.lower()
                    if low in KEYWORDS or low in BUILTINS or CONST_FAMILY.match(n):
                        continue
                    if low in local or low in m.module_names or low in globals_:
                        continue
                    if m.is_class and DESIGNER_CONTROL.match(n):
                        continue                       # drawn on a designer form (frmOvertimeEntry, frmPostingEntry)
                    print(f"UNDECLARED {label} {m.path.name}: {pname}: {n}")
                    problems += 1
    return problems


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    master = sorted((root / "modules").glob("*.bas")) + sorted((root / "paste-in").glob("*.txt"))
    member = sorted((root / "member").glob("*"))
    updater = sorted((root / "updater").glob("*.bas"))      # v24.2: its own workbook, so its own project
    n = check_project(master, "master") + check_project(member, "member") + check_project(updater, "updater")
    print(f"done {n} undeclared or reserved name(s)")
    return 1 if n else 0


if __name__ == "__main__":
    sys.exit(main())
