"""Build dist/Callout_Roster_v<N>.zip from the files that changed since the
previous tag, plus INSTALL.md.

    python tests/build_zip.py v21 v20

Modules and paste-ins are stored byte-for-byte (CRLF, ASCII). The install
notes come from docs/INSTALL-<tag>.md and are stored as INSTALL.md.
"""
from __future__ import annotations

import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    tag = sys.argv[1] if len(sys.argv) > 1 else "v21"
    base = sys.argv[2] if len(sys.argv) > 2 else "v20"
    out = ROOT / "dist" / f"Callout_Roster_{tag}.zip"
    changed = subprocess.run(
        ["git", "diff", "--name-only", f"{base}..HEAD", "--", "modules", "paste-in", "member"],
        cwd=ROOT, capture_output=True, text=True, check=True,
    ).stdout.split()
    # files in the working tree that are not committed yet also count
    pending = subprocess.run(
        ["git", "status", "--porcelain", "--", "modules", "paste-in", "member"],
        cwd=ROOT, capture_output=True, text=True, check=True,
    ).stdout.splitlines()
    for ln in pending:
        changed.append(ln[3:].strip())
    files = sorted({p for p in changed if (ROOT / p).is_file()})
    # from v22 on every zip is the full set, so it can serve as the master
    if "--full" in sys.argv or (tag.startswith("v") and tag[1:].isdigit() and int(tag[1:]) >= 22):
        files = sorted(
            [str(p.relative_to(ROOT)).replace("\\", "/") for p in ROOT.glob("modules/*.bas")]
            + [str(p.relative_to(ROOT)).replace("\\", "/") for p in ROOT.glob("paste-in/*.txt")]
            + [str(p.relative_to(ROOT)).replace("\\", "/") for p in ROOT.glob("member/*")]
        )
    install = ROOT / "docs" / f"INSTALL-{tag}.md"
    if not install.is_file():
        raise SystemExit(f"missing {install}")
    out.parent.mkdir(exist_ok=True)
    top = f"Callout_Roster_{tag}/"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for p in files:
            data = (ROOT / p).read_bytes()
            if b"\n" in data.replace(b"\r\n", b""):
                raise SystemExit(f"{p}: not CRLF")
            data.decode("ascii")
            z.writestr(top + p.replace("\\", "/"), data)
        z.writestr(top + "INSTALL.md", install.read_bytes())
    print(f"{out.relative_to(ROOT)}: {len(files)} code files + INSTALL.md")
    for p in files:
        print("  " + p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
