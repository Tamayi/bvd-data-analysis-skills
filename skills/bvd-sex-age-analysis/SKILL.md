---
name: bvd-sex-age-analysis
version: 4.0
description: >
  Generate the standard BVD Sex and Age epidemiological PowerPoint slide deck
  from a DHIS2 Tracker line-list export (Ebola/MVE outbreak, DRC). Runs all
  reconciliation quality gates before producing the deck.
compatible_with:
  - Claude (claude.ai, API) and Claude Code
  - Any LLM with a Python code interpreter or bash + python3
  - Offline use (run the portable bundle directly — no AI needed)
shared_library: bvd_common   # palette, pptx primitives, stats (see /src)
---

# BVD Sex and Age Analysis Skill

Generate the standard **Sex and Age** epidemiological deck (cases, deaths, CFR
by age group, gender, province and health zone) from a DHIS2 Tracker line-list.

---

## How this skill is organised

This skill is **DRY by default and portable on demand**:

| File | Role |
|---|---|
| `SKILL.md` (this file) | Workflow, analytical rules, and quality gates the AI must follow |
| `scripts/generate_slides.py` | The generator. Imports shared code from `bvd_common` |
| `../../src/bvd_common/` | Canonical shared library (pptx primitives, palette, stats) |
| `references/` | Deep-dive docs loaded only when needed (data dictionary, gate rationale) |
| `dist/<name>.standalone.py` | Generated portable copy — `bvd_common` inlined, no repo needed |

**In this repo (Claude Code, dev):** run `scripts/generate_slides.py` directly —
it adds `src/` to the path and imports `bvd_common`.

**Offline / ChatGPT / any LLM:** build a self-contained copy first, then ship or
run that single file:

```bash
python tools/build_portable.py skills/bvd-sex-age-analysis/scripts/generate_slides.py
# -> dist/generate_slides.standalone.py  (bvd_common inlined; runs anywhere)
```

> Always edit `scripts/generate_slides.py` and `src/bvd_common/`, never the
> generated `dist/` bundle. Regenerate the bundle after any change.

### What the AI actually does
1. Asks language (EN/FR) and confirms the uploaded line-list.
2. Runs `generate_slides.py <xlsx> --lang <en|fr> --out <file.pptx>`.
3. The script extracts data, **runs all quality gates**, and only then builds the deck.
4. The AI shows the gate report, sense-checks the printed numbers, and delivers the file.

---

## Workflow the AI must follow

### Step 1 — Ask language preference FIRST

> "Should the presentation be in **English** or **French**?
> _(La présentation doit-elle être en **anglais** ou en **français** ?)_"

### Step 2 — Confirm data upload

> **EN:** "Please confirm you have uploaded the latest DHIS2 Tracker line-list
> (`DHIS2_Tracker_Short_Line_List_YYYYMMDD.xlsx`). If not, please upload it now."
>
> **FR:** "Veuillez confirmer que vous avez téléversé le dernier export DHIS2
> (`DHIS2_Tracker_Short_Line_List_YYYYMMDD.xlsx`). Sinon, téléversez-le maintenant."

Do not proceed without a confirmed file.

### Step 3 — Install dependencies (if needed)

```bash
pip install pandas openpyxl python-pptx --quiet
```

### Step 4 — Run the script

```bash
python skills/bvd-sex-age-analysis/scripts/generate_slides.py <path_to_xlsx> \
  --lang <en|fr> --out BVD-Sex-Age-Analysis-Report-YYYYMMDD.pptx
```

The script **prints a quality-gate report** before building slides. If any gate
fails it prints `❌ QUALITY GATE FAILED` and exits without producing a file. The
AI must show the gate report to the user.

### Step 5 — Post-generation review (AI, not script)

Sense-check the numbers the script printed (no need to reopen the file):

**Totals reconciliation**
- Female + Male cases = Total confirmed cases
- Child + Adult + Unknown age = Total confirmed cases
- Province case / death totals sum to overall totals

**CFR signals**
- Flag overall CFR > 50% or < 1% (likely a filter/classification issue)
- Flag any single age-group CFR > 100%
- If a previous run was referenced, flag any age-group CFR that shifted > 5 pp

**Data quality signals**
- Flag if unknown age > 10% of total cases
- Flag if symptom-onset completeness < 80%
- Flag any new invalid age codes beyond {999, 1999, 9999}

**Geography**
- Flag any province with zero cases (possible filter issue)
- Flag if the number of health zones decreased since a previous run

Present the review as a short plain-language paragraph an epidemiologist would
use — not a checklist. Only flag genuinely unusual items; do not raise false alarms.

### Step 6 — Deliver

Present the `.pptx` and summarise: total cases, deaths, overall CFR, top zone by
caseload, child CFR vs adult CFR, and anything flagged in Step 5.

---

## Analytical rules (fixed — keep in sync with the script)

| Decision | Rule |
|---|---|
| **Confirmed cases only** | `classification_finale == "Cas confirmé"` |
| **Death definition** | `nature_alerte == "Décès"` OR `date_deces` not null OR `date_deces_alerte` not null |
| **Invalid age codes** | `age_ans` in `{9999, 1999, 999}` OR `NaN` → set to `NaN`, classified **Unknown** |
| **Child / Adult** | `0 ≤ age < 15` / `age ≥ 15` |
| **Unknown age** | Tracked explicitly; included in totals; shown in footnotes |
| **Age group bins** | `[0,4,9,14,17,24,34,44,54,64,200]`, `include_lowest=True` (captures age 0) |
| **Health zones shown** | Zones with ≥ 10 confirmed cases, top 9 by volume |
| **Rounding** | CFR and percentages rounded to 1 dp before display |
| **Data date** | Parsed from filename `..._YYYYMMDD.xlsx` |

---

## Quality gates (all must pass before slide generation)

Implemented in `scripts/generate_slides.py::run_gates`. They encode lessons from
known data-quality issues in this dataset. Detailed rationale: see
[references/quality-gates.md](references/quality-gates.md).

| # | Gate | Type |
|---|---|---|
| 1 | No unexpected invalid age codes (> 120, not in `INVALID_AGES`) | FAIL |
| 2 | Child + Adult + Unknown **cases** = total confirmed | FAIL |
| 3 | Child + Adult + Unknown **deaths** = total deaths | FAIL |
| 4 | Age-group cases + unknown-age = total confirmed | FAIL |
| 5 | Age-group deaths + unknown-age = total deaths | FAIL |
| 6 | Province case totals sum to overall total | FAIL |
| 7 | Per-province child + adult + unknown = province total (cases & deaths) | FAIL |
| 8 | Province death totals sum to overall total | FAIL |
| 9 | No CFR < 0 or > 120 | FAIL |
| 10 | All cases have a recognised sex value | WARN |
| 11 | F + M + Unknown **cases** = total | FAIL |
| 12 | F + M + Unknown **deaths** = total | FAIL |

---

## Slide structure (13 slides)

1 Title (stat + geo tiles) · 2 Data Quality Assessment · 3 Cases by age/gender ·
4 Deaths by age/gender · 5 CFR by age group · 6 Children vs Adults (divider) ·
7 Global cases & deaths · 8 Global key metrics · 9 Province cases · 10 Province
deaths · 11 Zone cases · 12 Zone deaths & CFR · 13 Zone female vs male.

Slides 7–10 footnote the number of unknown-age cases/deaths included in totals
but excluded from child/adult bars.

---

## Composing with other skills

Skills in this repo compose two ways (see [CONVENTIONS.md](../../CONVENTIONS.md)):

- **Shared code** — import reusable logic from `bvd_common` (this skill uses
  `palette`, `pptx`, `stats`). Promote skill-local logic into `bvd_common` once a
  second skill needs it.
- **Workflow hand-off** — to chain skills, this SKILL.md can instruct the AI to
  invoke another skill by name (e.g. "after generating the deck, run
  `bvd-trend-analysis` on the same line-list"). The AI orchestrates; outputs
  (the `.pptx`, the extracted `D` dict) are the hand-off.

---

## Expected DHIS2 column names

If a new export renames columns, update `extract()` accordingly. Full data
dictionary: [references/data-dictionary.md](references/data-dictionary.md).

| Column | Purpose |
|---|---|
| `classification_finale` | Filter: `== "Cas confirmé"` |
| `age_ans` | Age in years; `9999`, `1999`, `999` = unknown |
| `sexe` | `"Feminin"` or `"Masculin"` |
| `nature_alerte` | `"Décès"` = death |
| `date_deces`, `date_deces_alerte` | Death dates (supplementary) |
| `date_debut_symptomes_notification` | Onset date |
| `Province`, `Zone_sante`, `Aire_sante` | Geography |
