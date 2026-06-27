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
| [bvd-sitrep-review](skills/bvd-sitrep-review/) | Two-stage SitRep pipeline for the COUSP / WHO AFRO Bundibugyo (MVE-Bundibugyo) outbreak: enhance a raw SitRep with a day-over-day dashboard and figures, then produce a technical review & data-quality report. |

## Install for Claude (Cowork + Code)

One-click install via the plugin marketplace.

**Claude Code** — add the marketplace, then install the skills you want:

```
/plugin marketplace add Tamayi/bvd-data-analysis-skills
```

**Cowork** — go to **Customize → Plugins → Add marketplace** and paste the repo
URL (`https://github.com/Tamayi/bvd-data-analysis-skills`), then add the skills
from the list.

Both read [`.claude-plugin/marketplace.json`](.claude-plugin/marketplace.json),
which exposes every skill under `skills/` as an installable plugin.

## Use with ChatGPT

**Primary — GitHub connector.** This gives ChatGPT live read access to the repo
so it always sees the latest `SKILL.md`.

1. In ChatGPT, go to **Settings → Apps (Connectors) → GitHub** and authorize
   access to this repository.
2. Create a **ChatGPT Project** and paste the following into the Project
   instructions:

   ```
   This project uses the BVD Data Analysis Skills repo
   (Tamayi/bvd-data-analysis-skills) via the GitHub connector.

   When I request an analysis task, first read AGENTS.md to pick the matching
   skill, then open and follow that skill's instructions at
   skills/<name>/SKILL.md (and any references/ or scripts/ it points to).

   Available skills:
   - skills/bvd-sex-age-analysis/SKILL.md — Sex & Age PowerPoint deck from a
     DHIS2 line-list, with reconciliation quality gates.
   - skills/bvd-sitrep-review/SKILL.md — enhance and technically review a COUSP /
     WHO AFRO Bundibugyo (MVE-Bundibugyo) SitRep.

   Do not invent steps; follow the SKILL.md exactly.
   ```

**Fallback — Releases zip.** If you can't use the GitHub connector (e.g. ChatGPT
Plus without Deep Research / Agent Mode), download the skill zip from the repo's
[Releases](https://github.com/Tamayi/bvd-data-analysis-skills/releases) page and
upload it into the chat, then ask ChatGPT to follow the `SKILL.md` inside.

## Use with Gemini CLI, Codex, DeepCode, or other Agent-Skills runtimes

One path: `git clone` (or symlink) the repo into the agent's skills directory.
These runtimes read [`AGENTS.md`](AGENTS.md) to discover the skills, then load
`skills/<name>/SKILL.md` on demand.

```bash
# Gemini CLI (workspace)
git clone https://github.com/Tamayi/bvd-data-analysis-skills .agents
# or user-level
git clone https://github.com/Tamayi/bvd-data-analysis-skills ~/.agents/skills/bvd

# DeepCode / DeepSeek-TUI
git clone https://github.com/Tamayi/bvd-data-analysis-skills ~/.deepseek/skills/bvd
```

`.agents/skills/` is the interoperable alias recognized by Gemini CLI, DeepCode,
and other Agent Skills runtimes. Run `git pull` in the cloned directory to update.

## How it's organised

- `src/bvd_common/` — canonical shared library (palette, pptx primitives, stats).
- `skills/<name>/` — each skill: `SKILL.md`, `scripts/`, `references/`.
- `tools/build_portable.py` — inline the shared library into one standalone file.
- `tools/package_skills.py` — build `dist/<name>.zip` for each skill (Releases / ChatGPT fallback).
- `dist/` — generated portable bundles and zips (git-ignored).
- `data/` — real data exports (git-ignored; never commit patient-level data).

## Contributing

### Develop locally

```bash
pip install -e .                       # makes bvd_common importable; pulls deps
pip install pandas openpyxl python-pptx

# run a skill in-repo
python skills/bvd-sex-age-analysis/scripts/generate_slides.py data.xlsx --lang en

# build a portable, self-contained copy (no repo / no bvd_common needed)
python tools/build_portable.py skills/bvd-sex-age-analysis/scripts/generate_slides.py
```
