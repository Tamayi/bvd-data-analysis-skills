#!/usr/bin/env python3
"""Package each skill under ``skills/`` into ``dist/<name>.zip``.

Each zip contains a single top-level folder ``<name>/`` holding that skill's
contents, ready to drop into any Agent-Skills runtime or attach to a GitHub
Release for the ChatGPT / offline zip-fallback path.

Constraints enforced:
- zip max 50 MB
- each uncompressed file max 25 MB
- excludes ``__pycache__``, ``*.pyc``, ``.DS_Store``, ``data/`` and anything
  matched by ``.gitignore``
- exactly one ``SKILL.md`` per zip, at ``<name>/SKILL.md``
"""
import fnmatch
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
SKILLS = REPO / "skills"
DIST = REPO / "dist"

MAX_ZIP_BYTES = 50 * 1024 * 1024          # 50 MB
MAX_FILE_BYTES = 25 * 1024 * 1024         # 25 MB per uncompressed file

# Always-excluded path fragments / patterns (in addition to .gitignore).
EXCLUDE_DIRS = {"__pycache__", "data"}
EXCLUDE_GLOBS = ["*.pyc", ".DS_Store"]


def load_gitignore_patterns() -> list[str]:
    """Read simple glob patterns from .gitignore (best-effort, not full spec)."""
    gi = REPO / ".gitignore"
    if not gi.exists():
        return []
    patterns = []
    for line in gi.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or line.startswith("!"):
            continue
        patterns.append(line.rstrip("/"))
    return patterns


GITIGNORE = load_gitignore_patterns()


def is_excluded(rel_to_repo: Path) -> bool:
    parts = rel_to_repo.parts
    name = rel_to_repo.name
    if any(p in EXCLUDE_DIRS for p in parts):
        return True
    if any(fnmatch.fnmatch(name, g) for g in EXCLUDE_GLOBS):
        return True
    posix = rel_to_repo.as_posix()
    for pat in GITIGNORE:
        pat = pat.lstrip("/")
        if (
            fnmatch.fnmatch(posix, pat)
            or fnmatch.fnmatch(posix, f"{pat}/*")
            or fnmatch.fnmatch(name, pat)
            or any(fnmatch.fnmatch(p, pat) for p in parts)
        ):
            return True
    return False


def collect_files(skill_dir: Path) -> list[Path]:
    files = []
    for path in sorted(skill_dir.rglob("*")):
        if not path.is_file():
            continue
        if is_excluded(path.relative_to(REPO)):
            continue
        files.append(path)
    return files


def package(skill_dir: Path) -> Path:
    name = skill_dir.name
    files = collect_files(skill_dir)

    skill_mds = [f for f in files if f.name == "SKILL.md" and f.parent == skill_dir]
    if len(skill_mds) != 1:
        sys.exit(
            f"ERROR: {name} must have exactly one top-level SKILL.md "
            f"(found {len(skill_mds)})"
        )

    for f in files:
        size = f.stat().st_size
        if size > MAX_FILE_BYTES:
            sys.exit(
                f"ERROR: {f.relative_to(REPO)} is {size / 1024 / 1024:.1f} MB, "
                f"exceeds the 25 MB per-file limit"
            )

    DIST.mkdir(parents=True, exist_ok=True)
    out = DIST / f"{name}.zip"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            # arcname: <name>/<path-relative-to-skill-dir>
            arcname = Path(name) / f.relative_to(skill_dir)
            zf.write(f, arcname.as_posix())

    zip_size = out.stat().st_size
    if zip_size > MAX_ZIP_BYTES:
        out.unlink()
        sys.exit(
            f"ERROR: {name}.zip is {zip_size / 1024 / 1024:.1f} MB, "
            f"exceeds the 50 MB limit"
        )

    # sanity: exactly one SKILL.md at <name>/SKILL.md
    with zipfile.ZipFile(out) as zf:
        names = zf.namelist()
    if names.count(f"{name}/SKILL.md") != 1:
        sys.exit(f"ERROR: {name}.zip does not contain exactly one {name}/SKILL.md")

    return out


def main() -> None:
    if not SKILLS.is_dir():
        sys.exit("ERROR: no skills/ directory found")
    skill_dirs = sorted(d for d in SKILLS.iterdir() if d.is_dir() and (d / "SKILL.md").exists())
    if not skill_dirs:
        sys.exit("ERROR: no skills with a SKILL.md found under skills/")

    outputs = [package(d) for d in skill_dirs]
    for out in outputs:
        print(out.relative_to(REPO).as_posix())


if __name__ == "__main__":
    main()
