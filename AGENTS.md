# AGENTS.md

Cross-platform entry point for AI agents (Gemini CLI, Codex, DeepCode,
DeepSeek-TUI, and other Agent-Skills runtimes).

This repository is a collection of **skills**. Each skill lives in its own
folder under [`skills/`](skills/) and is described by a `SKILL.md` file
containing the workflow, analytical rules, and quality gates to follow.

## Available skills

| Skill | Load this file | What it does |
|---|---|---|
| `bvd-sex-age-analysis` | [`skills/bvd-sex-age-analysis/SKILL.md`](skills/bvd-sex-age-analysis/SKILL.md) | Generate the standard BVD Sex and Age epidemiological PowerPoint slide deck from a DHIS2 Tracker line-list export (Ebola/MVE outbreak, DRC). Runs all reconciliation quality gates before producing the deck. |
| `bvd-sitrep-review` | [`skills/bvd-sitrep-review/SKILL.md`](skills/bvd-sitrep-review/SKILL.md) | Two-stage pipeline for the COUSP / WHO AFRO SitRep of the Bundibugyo virus disease (MVE-Bundibugyo) outbreak in DRC: enhance a raw SitRep with a comparison dashboard and figures, then produce a technical review & data-quality report. |

## How to use

When the user requests a task that matches one of the skills above, **load
that skill's `SKILL.md`** and follow it. The `SKILL.md` points to deeper
references under `skills/<name>/references/` and runnable code under
`skills/<name>/scripts/` — load those on demand as the skill instructs.

Do not duplicate skill content here; always read the live `SKILL.md`.
