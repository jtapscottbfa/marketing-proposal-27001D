#!/usr/bin/env python3
"""scripts/publish/publish_outputs.py — versioned, never-overwrite output publisher.

Repo rule (see CLAUDE.md §6): every publish of pursuit outputs goes through
this script. It

  1. copies staged files into ``output/``:
       *.md                      -> output/artifacts/
       *.docx, *.xlsx, *.ics ... -> output/published/
     (the staging tree's ``artifacts/`` and ``published/`` sub-folders are
     honoured as-is, so ``published/calendar/x.ics`` lands in
     ``output/published/calendar/``);
  2. versions every file as ``<base>_vNN.<ext>``. A file is NEVER
     overwritten: if a current version exists it is moved to
     ``output/<area>/_superseded/vNN/`` and the new file takes vNN+1;
  3. saves state: commits the whole working tree on a short-lived
     ``publish/<timestamp>`` branch and fast-forward merges it into ``main``
     (or commits directly when ``main`` has no history yet), then optionally
     pushes ``main``.

Usage:
    python scripts/publish/publish_outputs.py --stage _scratch/stage \
        -m "docs(output): publish 27001D proposal outputs" [--push] [--no-git]
"""
from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts" / "lib"))
from guard_cwd import require_proposal_or_engine_root  # noqa: E402

OUTPUT = REPO_ROOT / "output"
VER_RE = re.compile(r"^(?P<base>.+?)_v(?P<n>\d{2,})$")
ATTRIBUTION = "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"


def split_version(stem: str) -> tuple[str, int | None]:
    m = VER_RE.match(stem)
    return (m.group("base"), int(m.group("n"))) if m else (stem, None)


def area_for(rel: Path) -> tuple[str, Path]:
    """Return (area, path-within-area) for a staged file."""
    parts = rel.parts
    if parts[0] in ("artifacts", "published"):
        return parts[0], Path(*parts[1:])
    return ("artifacts" if rel.suffix.lower() == ".md" else "published"), rel


def place(src: Path, area: str, inner: Path, log: list[str]) -> Path:
    dest_dir = OUTPUT / area / inner.parent
    dest_dir.mkdir(parents=True, exist_ok=True)
    base, _ = split_version(src.stem)
    ext = src.suffix
    current = []
    for f in dest_dir.glob(f"{base}_v*{ext}"):
        b, n = split_version(f.stem)
        if b == base and n is not None:
            current.append((n, f))
    nxt = max((n for n, _ in current), default=0) + 1
    for n, f in current:  # supersede, never delete
        sup = OUTPUT / area / "_superseded" / f"v{n:02d}" / inner.parent
        sup.mkdir(parents=True, exist_ok=True)
        target = sup / f.name
        if target.exists():
            sys.exit(f"refusing to overwrite superseded file: {target}")
        shutil.move(str(f), str(target))
        log.append(f"superseded  {f.relative_to(REPO_ROOT)} -> {target.relative_to(REPO_ROOT)}")
    dest = dest_dir / f"{base}_v{nxt:02d}{ext}"
    if dest.exists():
        sys.exit(f"refusing to overwrite: {dest}")
    shutil.copy2(src, dest)
    log.append(f"published   {dest.relative_to(REPO_ROOT)}")
    return dest


def retire(pattern: str, log: list[str]) -> None:
    """Move current versioned files matching ``pattern`` (relative to output/)
    to ``output/<area>/_superseded/vNN/`` with no replacement. Never deletes."""
    hits = [f for f in OUTPUT.glob(pattern)
            if f.is_file() and "_superseded" not in f.parts and split_version(f.stem)[1] is not None]
    if not hits:
        sys.exit(f"--retire matched no current files: {pattern}")
    for f in hits:
        rel = f.relative_to(OUTPUT)
        area, inner = rel.parts[0], Path(*rel.parts[1:])
        n = split_version(f.stem)[1]
        sup = OUTPUT / area / "_superseded" / f"v{n:02d}" / inner.parent
        sup.mkdir(parents=True, exist_ok=True)
        target = sup / f.name
        if target.exists():
            sys.exit(f"refusing to overwrite superseded file: {target}")
        shutil.move(str(f), str(target))
        log.append(f"retired     {f.relative_to(REPO_ROOT)} -> {target.relative_to(REPO_ROOT)}")


def git(*args: str, check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(["git", "-C", str(REPO_ROOT), *args], text=True,
                          capture_output=True, check=check)


def save_state(message: str, push: bool) -> None:
    has_main = git("rev-parse", "--verify", "--quiet", "refs/heads/main", check=False).returncode == 0
    msg = f"{message}\n\n{ATTRIBUTION}"
    if not has_main:
        git("checkout", "-B", "main")
        git("add", "-A")
        git("commit", "-m", msg)
        print("committed initial state directly on main")
    else:
        branch = f"publish/{datetime.now():%Y%m%d-%H%M%S}"
        git("checkout", "-b", branch)
        git("add", "-A")
        if git("diff", "--cached", "--quiet", check=False).returncode == 0:
            git("checkout", "main")
            git("branch", "-D", branch)
            print("nothing to commit; main unchanged")
            return
        git("commit", "-m", msg)
        git("checkout", "main")
        git("merge", "--ff-only", branch)
        git("branch", "-d", branch)
        print(f"merged {branch} into main (fast-forward)")
    if push:
        r = git("push", "-u", "origin", "main", check=False)
        print(r.stdout + r.stderr)


def main() -> None:
    require_proposal_or_engine_root(REPO_ROOT)
    ap = argparse.ArgumentParser()
    ap.add_argument("--stage", required=True, type=Path)
    ap.add_argument("-m", "--message", default="docs(output): publish proposal outputs")
    ap.add_argument("--push", action="store_true")
    ap.add_argument("--no-git", action="store_true")
    ap.add_argument("--retire", action="append", default=[], metavar="OUTPUT_GLOB",
                    help="move current files matching this glob (relative to output/) to _superseded without a replacement; repeatable")
    a = ap.parse_args()
    stage = (REPO_ROOT / a.stage).resolve() if not a.stage.is_absolute() else a.stage
    files = sorted(p for p in stage.rglob("*") if p.is_file())
    if not files:
        sys.exit(f"no files staged in {stage}")
    log: list[str] = []
    for f in files:
        area, inner = area_for(f.relative_to(stage))
        place(f, area, inner, log)
    for pattern in a.retire:
        retire(pattern, log)
    print("\n".join(log))
    if not a.no_git:
        save_state(a.message, a.push)


if __name__ == "__main__":
    main()
