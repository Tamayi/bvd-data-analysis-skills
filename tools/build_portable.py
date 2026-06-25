#!/usr/bin/env python3
"""Build a self-contained, portable copy of a skill script.

Skills are authored DRY: they import shared code from ``bvd_common``. For
offline / ChatGPT / "any LLM" distribution we need a single file with no
dependency on this repo. This tool inlines the ``bvd_common`` modules used by a
skill script directly into a standalone copy.

How it works
------------
Each skill script wraps its shared imports in a marker block::

    # >>> shared-imports ...
    sys.path.insert(...)
    from bvd_common.palette import ...
    from bvd_common.pptx import ...
    # <<< shared-imports

This tool replaces that whole block with the *source* of the referenced
``bvd_common`` modules (relative imports stripped, third-party imports kept —
duplicate imports are harmless in Python).

Usage
-----
    python tools/build_portable.py skills/bvd-sage-analysis/scripts/generate_slides.py
    python tools/build_portable.py skills/bvd-sage-analysis/scripts/generate_slides.py -o dist/foo.py
"""
import argparse
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
COMMON = REPO / "src" / "bvd_common"
START = "# >>> shared-imports"
END = "# <<< shared-imports"

# palette must precede modules that depend on it (e.g. pptx imports from palette)
ORDER = {"palette": 0}


def module_source(name: str) -> str:
    """Source of a bvd_common module with intra-package relative imports removed."""
    src = (COMMON / f"{name}.py").read_text(encoding="utf-8")
    kept = [ln for ln in src.splitlines() if not ln.lstrip().startswith("from .")]
    return f"# ── inlined: bvd_common.{name} " + "─" * 30 + "\n" + "\n".join(kept).strip() + "\n"


def build(script_path: Path, out_path: Path) -> Path:
    text = script_path.read_text(encoding="utf-8")
    if START not in text or END not in text:
        sys.exit(f"ERROR: no shared-imports marker block in {script_path}")

    block = text[text.index(START): text.index(END) + len(END)]
    modules = re.findall(r"from bvd_common\.(\w+) import", block)
    seen, ordered = set(), []
    for m in sorted(dict.fromkeys(modules), key=lambda m: (ORDER.get(m, 1), m)):
        if m not in seen:
            seen.add(m)
            ordered.append(m)

    inlined = (
        "# ============================================================\n"
        "# PORTABLE BUILD — bvd_common inlined. Do not edit by hand;\n"
        "# regenerate with tools/build_portable.py after editing the skill\n"
        "# script or the shared library.\n"
        "# ============================================================\n"
        + "\n".join(module_source(m) for m in ordered)
    )
    portable = text.replace(block, inlined)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(portable, encoding="utf-8")
    return out_path


def main():
    ap = argparse.ArgumentParser(description="Inline bvd_common into a standalone skill script")
    ap.add_argument("script", help="path to a skill script that imports bvd_common")
    ap.add_argument("-o", "--out", default=None, help="output path (default: dist/<name>.standalone.py)")
    args = ap.parse_args()

    script_path = Path(args.script).resolve()
    out_path = Path(args.out).resolve() if args.out else REPO / "dist" / f"{script_path.stem}.standalone.py"
    built = build(script_path, out_path)
    print(f"Built portable bundle: {built.relative_to(REPO) if REPO in built.parents else built}")


if __name__ == "__main__":
    main()
