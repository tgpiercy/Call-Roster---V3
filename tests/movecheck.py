"""The proof that a module split is PURE MOVEMENT (v27.1 refactor).

    py -3.14 tests/movecheck.py v27                  old = the tag, new = the working tree
    py -3.14 tests/movecheck.py v27 HEAD             new = a commit
    py -3.14 tests/movecheck.py --dirs OLD NEW       two folders, each holding modules/ paste-in/ member/ updater/

What it proves, over the whole master project (modules/*.bas + paste-in/*.txt; member/ and updater/
are compared file for file):

  LOST / NEW / TWICE  every procedure of the old tree is in the new tree exactly once, and the new
                      tree has no procedure the old one lacked
  BODY                each procedure's text is identical, line for line, except:
                        - its header's Private changed to Public (a PROMOTION, listed), and
                        - a module qualifier changed to the module the name now lives in:
                          modCalendar.BoardPrint -> modBoard.BoardPrint, in code or in a string
  DECL                every module-level declaration (Dim / Private / Public / Const / Type / Declare)
                      is in the new tree exactly once, word for word, except Private / Dim -> Public;
                      the only new declarations allowed are the new modules' X_VERSION constants, and a
                      version constant's value may change (the release stamps)
  COMMENT             every comment line outside a procedure is still somewhere (a banner moves with its
                      section); a comment ADDED outside a procedure is listed, not failed (--strict fails it)
  OTHER               every file that is not part of the split (no procedure left it or reached it) is
                      identical, except a module qualifier changed to where the name now lives
  SCOPE               after the move nothing reaches across a module boundary to something it can't see:
                      a call / name that is Private in another module (and Public nowhere), a qualified
                      Module.Name where that module has no such member or holds it Private, a Public
                      procedure whose signature uses its module's Private Type (all compile errors)

Exit code 1 on any failure. It is a text comparison, not a compiler: run the four checkers too.
"""
from __future__ import annotations

import re
import subprocess
import sys
from collections import Counter, defaultdict
from pathlib import Path

PROJECT_DIRS = ("modules", "paste-in", "member", "updater")
HEAD_RE = re.compile(r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)", re.I)
END_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
ONE_LINE_END = re.compile(r":\s*End\s+(Sub|Function|Property)\s*$", re.I)
QUAL_RE = re.compile(r"(?<![\w.!])([A-Za-z_]\w*)\.([A-Za-z_]\w*)")
VERSION_CONST = re.compile(r"^Public\s+Const\s+[A-Z][A-Z0-9_]*_VERSION\s+As\s+String\s*=", re.I)
TOKEN = re.compile(r"[A-Za-z_]\w*")
KEYWORDS = set("""and as boolean byref byte byval call case const currency date debug declare dim do double each else
elseif empty end enum eqv erase error exit explicit false for friend function get global gosub goto if imp in integer
is let lib like long longlong longptr loop me mod new next not nothing null object on optional or paramarray preserve
print private property public raiseevent redim rem resume return select set single static step stop string sub then
to true type typeof until variant wend while with withevents xor attribute option compare binary text base module
ptrsafe""".split())


# ---------------------------------------------------------------- reading
def code_part(line: str) -> str:
    """The line with its comment removed (strings kept)."""
    inside = False
    for i, ch in enumerate(line):
        if ch == '"':
            inside = not inside
        elif ch == "'" and not inside:
            return line[:i]
    return line


def mask_strings(s: str) -> str:
    return re.sub(r'"(?:[^"]|"")*"', lambda m: '"' + " " * (len(m.group(0)) - 2) + '"', s)


class Proc:
    def __init__(self, file: str, name: str, kind: str, scope: str | None, lines: list[str], start: int):
        self.file, self.name, self.kind, self.scope, self.lines, self.start = file, name, kind, scope, lines, start
        self.key = name.lower() + ("|" + kind.split()[-1].lower() if kind.lower().startswith("property") else "")

    @property
    def private(self) -> bool:
        return (self.scope or "").lower() == "private"

    def where(self) -> str:
        return f"{self.file}:{self.start + 1}"


class Decl:
    def __init__(self, file: str, text: str, start: int, names: list[str], scope: str, kind: str):
        self.file, self.text, self.start, self.names, self.scope, self.kind = file, text, start, names, scope, kind
        self.norm = normalise_decl(text)


def normalise_decl(text: str) -> str:
    """The declaration without its scope word: 'Private x As Long' / 'Dim x As Long' / 'Public x As Long' -> 'x As Long';
    'Private Const A = 1' / 'Const A = 1' -> 'Const A = 1'; 'Private Type T' -> 'Type T'."""
    t = text.strip()
    t = re.sub(r"^(Public|Private|Global|Dim)\s+", "", t, flags=re.I)
    return t


class SourceFile:
    def __init__(self, path: str, text: str):
        self.path = path
        self.name = Path(path).stem
        self.is_std = path.lower().endswith(".bas")
        self.text = text
        self.lines = text.replace("\r\n", "\n").split("\n")
        if self.lines and self.lines[-1] == "":
            self.lines.pop()
        self.procs: list[Proc] = []
        self.decls: list[Decl] = []
        self.comments: list[tuple[int, str]] = []     # comment lines outside procedures
        self._parse()

    def _logical(self):
        """(first index, last index) of each logical line (continuations joined)."""
        i, n = 0, len(self.lines)
        while i < n:
            j = i
            while j < n and re.search(r"\s_\s*$", code_part(self.lines[j])) and j + 1 < n:
                j += 1
            yield i, j
            i = j + 1

    def _parse(self):
        in_proc = None
        block = None                   # a Type / Enum block outside procedures
        for a, b in self._logical():
            raw = self.lines[a:b + 1]
            joined = " ".join(code_part(x).rstrip().rstrip("_") for x in raw)
            code = code_part(joined).strip()
            if in_proc is not None:
                in_proc[1].extend(raw)
                if END_RE.match(code):
                    self._close(in_proc)
                    in_proc = None
                continue
            if block is not None:
                block[1].extend(raw)
                if re.match(r"^End\s+(Type|Enum)\b", code, re.I):
                    text = "\n".join(block[1])
                    self.decls.append(Decl(self.path, text, block[0], [block[2]], block[3], "type"))
                    block = None
                continue
            m = HEAD_RE.match(code)
            if m and not re.search(r"\bDeclare\b", code, re.I):
                in_proc = [a, list(raw), m.group(3), m.group(2), m.group(1)]
                if ONE_LINE_END.search(code):
                    self._close(in_proc)
                    in_proc = None
                continue
            if not code:
                for k, ln in enumerate(raw):
                    if ln.strip():
                        self.comments.append((a + k, ln.strip()))
                continue
            if re.match(r"^(Attribute|Option)\b", code, re.I):
                continue
            mt = re.match(r"^(?:(Public|Private)\s+)?(Type|Enum)\s+([A-Za-z_]\w*)", code, re.I)
            if mt:
                scope = (mt.group(1) or "Public").lower()
                block = [a, list(raw), mt.group(3), scope]
                continue
            md = re.match(r"^(Public|Private|Global|Dim|Const|Declare|#)", code, re.I)
            if md:
                self.decls.append(Decl(self.path, "\n".join(raw), a, decl_names(code), decl_scope(code), "decl"))
                continue
            self.decls.append(Decl(self.path, "\n".join(raw), a, [], "?", "odd"))

    def _close(self, rec):
        a, lines, name, kind, scope = rec
        self.procs.append(Proc(self.path, name, kind, scope, lines, a))


def decl_scope(code: str) -> str:
    w = code.split()[0].lower()
    if w in ("public", "global"):
        return "public"
    if w == "declare":
        return "public"
    return "private"                  # Private / Dim / a bare Const


def decl_names(code: str) -> list[str]:
    c = mask_strings(code)
    m = re.match(r"^(?:Public|Private|Global|Dim)?\s*(?:Const\s+)?(?:WithEvents\s+)?(.*)$", c, re.I)
    rest = m.group(1) if m else c
    dm = re.match(r"^(?:(?:Public|Private)\s+)?Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+([A-Za-z_]\w*)", c, re.I)
    if dm:
        return [dm.group(1)]
    names, depth, cur = [], 0, ""
    for ch in rest:
        depth += (ch == "(") - (ch == ")")
        if ch == "," and depth == 0:
            names.append(cur)
            cur = ""
        else:
            cur += ch
    names.append(cur)
    out = []
    for piece in names:
        mm = re.match(r"\s*([A-Za-z_]\w*)", piece)
        if mm:
            out.append(mm.group(1))
    return out


class Tree:
    def __init__(self, label: str, files: dict[str, str]):
        self.label = label
        self.files = {p: SourceFile(p, t) for p, t in files.items()}

    @staticmethod
    def from_dir(root: Path, label: str) -> "Tree":
        files = {}
        for d in PROJECT_DIRS:
            for p in sorted((root / d).glob("*")):
                if p.suffix.lower() in (".bas", ".txt") and p.is_file():
                    files[f"{d}/{p.name}"] = p.read_bytes().decode("ascii", errors="replace")
        return Tree(label, files)

    @staticmethod
    def from_git(repo: Path, ref: str) -> "Tree":
        names = subprocess.run(["git", "-C", str(repo), "ls-tree", "-r", "--name-only", ref, *PROJECT_DIRS],
                               capture_output=True, text=True, check=True).stdout.split()
        files = {}
        for n in names:
            if n.lower().endswith((".bas", ".txt")) and n.split("/")[0] in PROJECT_DIRS and n.count("/") == 1:
                raw = subprocess.run(["git", "-C", str(repo), "show", f"{ref}:{n}"], capture_output=True, check=True).stdout
                files[n] = raw.decode("ascii", errors="replace")
        return Tree(ref, files)

    def master(self) -> list[SourceFile]:
        return [f for p, f in self.files.items() if p.startswith(("modules/", "paste-in/"))]

    def std_names(self) -> dict[str, SourceFile]:
        return {f.name.lower(): f for f in self.master() if f.is_std}


# ---------------------------------------------------------------- comparing
class Report:
    def __init__(self):
        self.fails: list[str] = []
        self.notes: list[str] = []

    def fail(self, tag: str, msg: str):
        self.fails.append(f"{tag:8} {msg}")

    def note(self, tag: str, msg: str):
        self.notes.append(f"{tag:8} {msg}")


def canon_line(line: str, home: dict[str, str], std_old: set[str], std_new: set[str], side: str) -> str:
    """Module qualifiers rewritten to the OLD home of the name they qualify, so a qualifier changed to follow a
    move compares equal and nothing else does. home: member name (lower) -> old module (lower) for every name
    that moved; new: member name (lower) -> new module."""
    def sub(m):
        mod, mem = m.group(1), m.group(2)
        ml = mod.lower()
        if side == "new" and ml in std_new and mem.lower() in home.get("@new", {}) and home["@new"][mem.lower()] == ml:
            return home[mem.lower()] + "." + mem
        if ml in std_old or ml in std_new:
            return ml + "." + mem
        return m.group(0)
    return QUAL_RE.sub(sub, line)


def strip_scope(header: str) -> str:
    return re.sub(r"^(\s*)(?:(?:Public|Private|Friend)\s+)", r"\1", header, flags=re.I)


def compare(old: Tree, new: Tree, strict: bool = False) -> Report:
    rep = Report()
    om, nm = old.master(), new.master()
    std_old, std_new = set(old.std_names()), set(new.std_names())

    # ---- match procedures: same file first, then the leftovers by name ----
    old_by = {(f.path, p.key): p for f in om for p in f.procs}
    new_by = {(f.path, p.key): p for f in nm for p in f.procs}
    for f in om + nm:
        seen = Counter(p.key for p in f.procs)
        for k, c in seen.items():
            if c > 1:
                rep.fail("TWICE", f"{f.path}: procedure {k} is defined {c} times in one file")
    pairs: list[tuple[Proc, Proc]] = []
    left_old = [p for k, p in old_by.items() if k not in new_by]
    left_new = [p for k, p in new_by.items() if k not in old_by]
    pairs += [(old_by[k], new_by[k]) for k in old_by if k in new_by]
    by_name_new = defaultdict(list)
    for p in left_new:
        by_name_new[p.key].append(p)
    moved: list[tuple[Proc, Proc]] = []
    for p in left_old:
        cands = [q for q in by_name_new.get(p.key, []) if q.file.endswith(".bas") == p.file.endswith(".bas")]
        if not cands:
            rep.fail("LOST", f"{p.name} ({p.where()}) is in no file of {new.label}")
            continue
        if len(cands) > 1:
            rep.fail("TWICE", f"{p.name} ({p.where()}) is now in {len(cands)} files: {', '.join(q.file for q in cands)}")
        q = cands[0]
        by_name_new[p.key].remove(q)
        moved.append((p, q))
    for k, qs in by_name_new.items():
        for q in qs:
            rep.fail("NEW", f"{q.name} ({q.where()}) was not in {old.label}: a move adds no procedure")
    pairs += moved

    # ---- where every moved name lives now (procedures and module-level names) ----
    home: dict = {"@new": {}}
    for p, q in moved:
        if p.file.endswith(".bas"):
            home[p.name.lower()] = Path(p.file).stem.lower()
            home["@new"][p.name.lower()] = Path(q.file).stem.lower()

    # ---- module-level declarations ----
    def decl_bag(files):
        bag = defaultdict(list)
        for f in files:
            for d in f.decls:
                bag[d.norm].append(d)
        return bag
    split_old = {p.file for p, q in moved} | {q.file for p, q in moved}
    split_new = split_old | {f.path for f in nm if f.path not in old.files}
    old_d = decl_bag([f for f in om if f.path in split_old])
    new_d = decl_bag([f for f in nm if f.path in split_new])
    for norm in set(old_d) | set(new_d):
        o, n = old_d.get(norm, []), new_d.get(norm, [])
        if len(o) == len(n):
            for a, b in zip(sorted(o, key=lambda d: d.file), sorted(n, key=lambda d: d.file)):
                if a.scope == "public" and b.scope != "public":
                    rep.fail("DECL", f"{b.file}:{b.start + 1}: '{norm.splitlines()[0][:60]}' went from Public to Private")
                if a.file != b.file:
                    for nmn in a.names:
                        home[nmn.lower()] = Path(a.file).stem.lower()
                        home["@new"][nmn.lower()] = Path(b.file).stem.lower()
                    what = ", ".join(a.names) if a.names else norm[:40]
                    rep.note("DECL", f"{what}: {a.file} -> {b.file}" + (" (now Public)" if a.scope != b.scope else ""))
            continue
        for b in n[len(o):]:
            first = b.text.strip().splitlines()[0]
            if VERSION_CONST.match(first.strip()):
                rep.note("VERSION", f"{b.file}:{b.start + 1}: {first.strip()[:70]}")
                continue
            rep.fail("DECL", f"{b.file}:{b.start + 1}: '{first.strip()[:70]}' is new")
        for a in o[len(n):]:
            first = a.text.strip().splitlines()[0]
            if VERSION_CONST.match(first.strip()):
                vname = re.match(r"^Public\s+Const\s+(\w+)", first.strip(), re.I).group(1).lower()
                if any(vname in [x.lower() for x in d.names] for ds in new_d.values() for d in ds):
                    continue                                   # the same constant with a new value (a release stamp)
            rep.fail("DECL", f"{a.file}:{a.start + 1}: '{first.strip()[:70]}' is gone (or changed)")

    # ---- comments outside procedures ----
    oc = Counter(c for f in om if f.path in split_old for _, c in f.comments)
    nc = Counter(c for f in nm if f.path in split_new for _, c in f.comments)
    for c, k in (oc - nc).items():
        rep.fail("COMMENT", f"{k} x '{c[:70]}' is gone")
    added = nc - oc
    if added:
        msg = f"{sum(added.values())} comment line(s) added outside procedures (new modules' headers?): " + \
              "; ".join(f"'{c[:40]}'" for c in list(added)[:4])
        (rep.fail if strict else rep.note)("COMMENT", msg)

    # ---- bodies ----
    promotions = []
    for p, q in pairs:
        if len(p.lines) != len(q.lines):
            rep.fail("BODY", f"{p.name}: {len(p.lines)} lines in {p.where()}, {len(q.lines)} in {q.where()}")
            continue
        for i, (a, b) in enumerate(zip(p.lines, q.lines)):
            if i == 0:
                if p.private and not q.private:
                    if q.scope is None or q.scope.lower() == "public":
                        promotions.append(f"{q.name} ({Path(p.file).stem} -> {Path(q.file).stem})")
                    a, b = strip_scope(a), strip_scope(b)
                elif (p.scope or "").lower() != (q.scope or "").lower():
                    if not p.private and q.private:
                        rep.fail("BODY", f"{q.name} ({q.where()}): Public became Private")
                        break
                    a, b = strip_scope(a), strip_scope(b)
            ca = canon_line(a, home, std_old, std_new, "old")
            cb = canon_line(b, home, std_old, std_new, "new")
            if ca != cb:
                rep.fail("BODY", f"{q.name} ({q.where()} line {i + 1}): '{b.strip()[:60]}' was '{a.strip()[:60]}'")
                break
    for pr in promotions:
        rep.note("PROMOTE", pr)
    for p, q in moved:
        rep.note("MOVE", f"{p.name}: {Path(p.file).stem} -> {Path(q.file).stem}")

    # ---- every other file: identical but for qualifiers that follow a move ----
    for path, f in new.files.items():
        if path in split_new or path not in old.files:
            continue
        a_lines, b_lines = old.files[path].lines, f.lines
        if a_lines == b_lines:
            continue
        if len(a_lines) != len(b_lines):
            rep.fail("OTHER", f"{path}: {len(a_lines)} lines before, {len(b_lines)} after (only a qualifier may change)")
            continue
        for i, (a, b) in enumerate(zip(a_lines, b_lines)):
            if a != b and canon_line(a, home, std_old, std_new, "old") != canon_line(b, home, std_old, std_new, "new"):
                rep.fail("OTHER", f"{path}:{i + 1}: '{b.strip()[:60]}' was '{a.strip()[:60]}'")
                break
        else:
            rep.note("QUALIFY", f"{path}: module qualifiers follow the move")
    for path in old.files:
        if path not in new.files and path not in split_old:
            rep.fail("OTHER", f"{path} is gone")
    scope_check(new, rep)
    return rep


# ---------------------------------------------------------------- scope (the new tree on its own)
def scope_check(tree: Tree, rep: Report, only: set[str] | None = None):
    files = tree.master()
    std = tree.std_names()
    pub_any: set[str] = set()
    members: dict[str, dict[str, str]] = {}           # module -> member -> "public" / "private"
    priv_types: dict[str, set[str]] = defaultdict(set)
    for f in files:
        mm = {}
        for p in f.procs:
            sc = "private" if p.private else "public"
            mm[p.name.lower()] = sc
            if f.is_std and sc == "public":
                pub_any.add(p.name.lower())
        for d in f.decls:
            for n in d.names:
                mm[n.lower()] = d.scope
                if f.is_std and d.scope == "public":
                    pub_any.add(n.lower())
            if d.kind == "type" and d.scope == "private":
                priv_types[f.name.lower()].add(d.names[0].lower())
        members[f.name.lower()] = mm
    priv_owner: dict[str, list[str]] = defaultdict(list)
    for f in files:
        if not f.is_std:
            continue
        for n, sc in members[f.name.lower()].items():
            if sc == "private" and n not in pub_any:
                priv_owner[n].append(f.name)
    for f in files:
        if only and f.path not in only:
            continue
        own = members[f.name.lower()]
        for p in f.procs:
            body = [mask_strings(code_part(x)) for x in p.lines]
            text = " ".join(body)
            local = set()
            k = 0
            while k < len(body) - 1 and re.search(r"\s_\s*$", body[k]):
                k += 1
            hdr = " ".join(x.rstrip().rstrip("_") for x in body[:k + 1])      # the header statement, continuations joined
            for prm in re.findall(r"(?:\(|,)\s*(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(?:ParamArray\s+)?([A-Za-z_]\w*)", hdr, re.I):
                local.add(prm.lower())
            for ln in body:
                for st in ln.split(":"):
                    dm = re.match(r"^\s*(?:Dim|Static|Const|ReDim(?:\s+Preserve)?)\s+(.*)$", st, re.I)
                    if dm:
                        for n in decl_names("Dim " + dm.group(1)):
                            local.add(n.lower())
                    lab = re.match(r"^\s*([A-Za-z_]\w*):", ln)
                    if lab:
                        local.add(lab.group(1).lower())
            local.add(p.name.lower())
            # a Public procedure's signature using its module's Private Type
            if f.is_std and not p.private and priv_types.get(f.name.lower()):
                for t in re.findall(r"\bAs\s+([A-Za-z_]\w*)", hdr, re.I):
                    if t.lower() in priv_types[f.name.lower()]:
                        rep.fail("SCOPE", f"{p.where()}: Public {p.name} uses the Private Type {t} in its signature")
            for m in TOKEN.finditer(text):
                tok = m.group(0)
                low = tok.lower()
                a = m.start()
                if a > 0 and text[a - 1] == ".":
                    q = re.search(r"([A-Za-z_]\w*)\.$", text[:a])
                    if q and q.group(1).lower() in std and q.group(1).lower() != low:
                        qm = q.group(1).lower()
                        sc = members[qm].get(low)
                        if sc is None:
                            rep.fail("SCOPE", f"{p.where()} ({p.name}): {q.group(1)}.{tok} - {q.group(1)} has no {tok}")
                        elif sc == "private" and qm != f.name.lower():
                            rep.fail("SCOPE", f"{p.where()} ({p.name}): {q.group(1)}.{tok} is Private to {q.group(1)}")
                    continue
                if low in KEYWORDS or low in local or low in own:
                    continue
                if low in priv_owner:
                    rep.fail("SCOPE", f"{p.where()} ({p.name}): '{tok}' is Private to {', '.join(priv_owner[low])}")


def main(argv: list[str]) -> int:
    strict = "--strict" in argv
    args = [a for a in argv if a != "--strict"]
    repo = Path(__file__).resolve().parent.parent
    if args and args[0] == "--dirs":
        old, new = Tree.from_dir(Path(args[1]), args[1]), Tree.from_dir(Path(args[2]), args[2])
    elif args:
        old = Tree.from_git(repo, args[0])
        new = Tree.from_git(repo, args[1]) if len(args) > 1 else Tree.from_dir(repo, "working tree")
    else:
        print(__doc__)
        return 2
    rep = compare(old, new, strict)
    for n in rep.notes:
        print("note    ", n)
    for f in rep.fails:
        print("FAIL    ", f)
    moved = sum(1 for n in rep.notes if n.startswith("MOVE"))
    prom = sum(1 for n in rep.notes if n.startswith("PROMOTE"))
    print(f"done: {moved} procedure(s) moved, {prom} promoted, {len(rep.fails)} failure(s)")
    return 1 if rep.fails else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
