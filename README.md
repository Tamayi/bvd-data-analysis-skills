# BVD Data Analysis Skills

A collection of data-analysis **skills** for the team's epidemiological datasets
(DHIS2 Tracker line-lists, MVE/Ebola outbreak, DRC). Each skill is a folder under
[`skills/`](skills/) with a `SKILL.md` an AI assistant follows, plus runnable
scripts. Common logic lives once in a shared library and can be inlined into
self-contained, portable copies.

See [CONVENTIONS.md](CONVENTIONS.md) for the full layout and rules.

## Skills

| Skill | What it does |
|---|---|
| [bvd-sex-age-analysis](skills/bvd-sex-age-analysis/) | Generate the standard Sex & Age PowerPoint deck (cases, deaths, CFR by age/gender/province/zone) from a DHIS2 line-list, with reconciliation quality gates. |

## Quick start

```bash
pip install -e .                       # makes bvd_common importable; pulls deps
pip install pandas openpyxl python-pptx

# run a skill in-repo
python skills/bvd-sex-age-analysis/scripts/generate_slides.py data.xlsx --lang en

# build a portable, self-contained copy (no repo / no bvd_common needed)
python tools/build_portable.py skills/bvd-sex-age-analysis/scripts/generate_slides.py
```

## How it's organised

- `src/bvd_common/` — canonical shared library (palette, pptx primitives, stats).
- `skills/<name>/` — each skill: `SKILL.md`, `scripts/`, `references/`.
- `tools/build_portable.py` — inline the shared library into one standalone file.
- `dist/` — generated portable bundles (git-ignored).
- `data/` — real data exports (git-ignored; never commit patient-level data).
