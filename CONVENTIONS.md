# Repo conventions — how skills are organised

This repo holds **multiple data-analysis skills** that share code and sometimes
compose with each other. The design reconciles two goals the team chose:

- **DRY** — common logic lives once, in `src/bvd_common/`.
- **Portable** — any skill can be exported as a single self-contained file that
  runs offline / in ChatGPT / in any LLM with Python, with no dependency on this
  repo.

The trick: the shared library is the **source of truth**; portability is a
**build output** (`tools/build_portable.py` inlines the library into a standalone
script). You author DRY and ship portable.

## Layout

```
bvd-data-analysis-skills/
├── README.md
├── CONVENTIONS.md                  # this file
├── pyproject.toml                  # makes `bvd_common` importable (pip install -e .)
├── src/
│   └── bvd_common/                 # CANONICAL shared library
│       ├── palette.py              #   colours
│       ├── pptx.py                 #   python-pptx drawing primitives
│       └── stats.py                #   generic epi maths (cfr, pct)
├── skills/
│   └── <skill-name>/
│       ├── SKILL.md                #   workflow + rules the AI follows
│       ├── scripts/                #   runnable code; imports bvd_common
│       └── references/             #   deep-dive docs, loaded on demand
├── tools/
│   └── build_portable.py           # inline bvd_common -> dist/<name>.standalone.py
├── tests/fixtures/                 # small synthetic inputs
├── data/                           # real exports (git-ignored)
└── dist/                           # generated portable bundles (git-ignored)
```

## Adding a new skill

1. `skills/<skill-name>/SKILL.md` — frontmatter (`name`, `version`, `description`,
   `compatible_with`) plus the workflow, analytical rules, and quality gates the
   AI must follow. Keep verbose detail in `references/`, loaded only when needed.
2. `skills/<skill-name>/scripts/` — the runnable code. Import shared logic from
   `bvd_common` inside the marked block so it can be inlined:

   ```python
   # >>> shared-imports (this block is inlined by tools/build_portable.py)
   import sys, pathlib
   sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[3] / "src"))
   from bvd_common.pptx import txt, rect, chart
   # <<< shared-imports
   ```
   `parents[3]` = repo root from `skills/<name>/scripts/<file>.py`.
3. Build a portable copy when you need one:
   `python tools/build_portable.py skills/<skill-name>/scripts/<file>.py`

## The shared-code rule (avoid premature abstraction)

- **Generic and dataset-agnostic** (drawing primitives, colour palette, CFR maths)
  → `bvd_common` immediately.
- **Specific to one dataset/deck** (DHIS2 column names, the 12 reconciliation
  gates, EN/FR slide strings) → keep in the skill.
- **Promote** skill-local logic into `bvd_common` only when a *second* skill needs
  it. One skill is not enough evidence to abstract.

## How skills relate to each other

Two distinct mechanisms — don't conflate them:

| | Shared code | Workflow hand-off |
|---|---|---|
| **What** | Import the same functions | One skill triggers another |
| **How** | `from bvd_common... import ...` | SKILL.md instructs the AI to invoke skill B by name |
| **Coupling** | Compile-time (Python import) | Run-time (AI orchestration); hand off via files / the extracted data dict |

Claude Code resolves a skill invoked by name; the standalone/offline bundles do
not auto-chain, so a portable skill that hands off must say so in prose for the
human to run the next step.

## Verifying a change

```bash
pip install -e .                       # once: makes bvd_common importable
python -c "import bvd_common.pptx"      # shared lib imports
python tools/build_portable.py skills/<name>/scripts/<file>.py   # portable build succeeds
```
