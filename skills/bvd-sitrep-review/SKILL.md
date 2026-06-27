---
name: bvd-sitrep-review
description: >
  Two-stage pipeline for the COUSP / WHO AFRO internal Situation Report (SitRep) of the
  Bundibugyo virus disease (MVE-Bundibugyo) outbreak in the Democratic Republic of the Congo.
  STAGE 1 (enhance): turn a raw COUSP SitRep into an enriched final by adding a colour-coded
  day-over-day comparison dashboard and six analytical figures, preserving all original
  formatting, logos and maps. STAGE 2 (review): produce a technical review & data-quality
  report that arithmetically verifies every table, reconciles deaths and cumulative deltas,
  flags impossible values and missing data, and lists pre-dissemination corrections.
  Use this skill whenever the user mentions a "SitRep MVE", "SitRep BVD/MVB", "COUSP",
  "rapport de situation Ebola", "Bundibugyo", a SitRep number like "N°041", or asks to
  enhance, finalise, review, fact-check, or QA an Ebola/MVE situation report for this event —
  even if they don't explicitly say "enhance" or "review". Do NOT use for other diseases or
  for generic Word-document tasks unrelated to this outbreak.
version: "1.0"
last_updated: "2026-06-25"
author: "BVD data-analysis skills — Tamayi"
---

# BVD SitRep — Enhance & Review

## What this skill is for

The COUSP (Centre d'Opérations d'Urgence de Santé Publique) publishes a daily SitRep for the
17th Ebola (MVE-Bundibugyo) epidemic in the DRC. The raw SitRep is data-rich but plain: tables
of numbers, narrative bullets, maps. Two things add value before it circulates:

1. **Enhancement** — a one-glance comparison dashboard (today vs yesterday) and a small set of
   figures that turn the tables into trends a decision-maker can read in seconds.
2. **Technical review** — an independent pass that checks the arithmetic, reconciles the death
   counts, catches impossible or missing values, and lists what to fix before dissemination.

These are **two stages of one process**. They can run together (enhance, then review the result)
or independently. Decide which the user wants from their request: "enrich / add figures / make
the final" → Stage 1; "review / QA / fact-check / data quality" → Stage 2; "do the full thing"
→ both, Stage 1 then Stage 2.

## Before doing anything: read the source and the previous day

- **Current SitRep** (the one to enhance/review): a `.docx` or `.pdf`. Read its full text and
  every table. Tables you will always need: Tableau 1 (province totals), Tableau 2 (health-zone
  detail), Tableau 3 (alerts), Tableau 4 (contacts), Tableau 5 (PoE/PoC), Tableau 6 (patient
  movement), Tableau 7 (SMSPS), plus the lab section (4.3) and Faits Saillants.
- **Previous SitRep (N-1)** — required for the dashboard deltas and to verify cumulative changes.
  If only the current file is provided, ask for N-1 or locate it (the COUSP/INSP site publishes
  each day; the previous "FINAL" in the project is ideal).
- **Earlier SitReps (N-2, N-3)** are valuable: they let the trend figures use *verified daily
  anchors* instead of interpolation. Pull them when available.

## Data governance (these rules win any conflict)

- **Never invent a number.** Every value must trace to a named table or SitRep. If a figure
  needs a value you don't have, interpolate *only* between real anchors and label it explicitly.
- **Never merge confirmed / probable / suspected** unless the source gives a combined category.
- **Deaths carry nuance.** The daily death count in Faits Saillants often counts only deaths
  *among new confirmed cases*; additional deaths occur in CTEs among already-confirmed patients.
  Always reconcile against the cumulative change (see review.md). Do not "simplify" this away.
- **Don't claim a vaccine or approved therapeutic for Bundibugyo exists** unless the source says so.
- **Everything produced is a draft for technical validation** by the surveillance/IM lead. Say so
  if asked; the published finals in this project omit a printed disclaimer, so only add one if the
  user asks (see "House style").

## Stage 1 — Enhancement

Goal: inject a comparison dashboard and six figures into a copy of the source DOCX, changing
nothing else (logos, maps, photos, existing tables and text stay byte-for-byte where possible).

Read **`references/enhancement.md`** for the full recipe: exact figure list, the data each one
needs, caption/source-note wording, the WHO colour palette, the dashboard row set and colour
rules, and the precise injection anchors inside the document. Use **`scripts/make_figures.py`**
to render the six figures from a small JSON of the day's numbers — it encodes the palette,
sizes and styling so every SitRep looks consistent, while you supply the data. Then inject with
`python-docx` (insert the table and images before their anchor paragraphs) and verify by
rendering to PDF and eyeballing the dashboard page and each figure.

## Stage 2 — Technical review & data-quality report

Goal: a short, skimmable Word report that tells the IM exactly what is correct, what is wrong,
and what to fix — in the same house format as prior reviews.

Read **`references/review.md`** for the full method: the verification families to run (province
sums, CFR, contact and alert rates, death reconciliation, isolation balance, weighted-average
sanity, backlog, cross-document consistency), the *reasoning* checks that matter most (novel
epidemiological signals, missing/ND data that breaks comparability, narrative-vs-table
mismatches), and the exact report structure (global result, errors & inconsistencies with
severity, full arithmetic table, suggested figures, pre-dissemination checklist).

**Verification is prose-and-judgment, not a rigid script.** Compute the arithmetic yourself and
show your working, but spend most of your attention on what the numbers *mean*: an internally
consistent table can still hide a misleading headline or a globally significant event. The
highest-value findings in past reviews were not arithmetic slips — they were context calls.

## House style (applies to both stages)

- **Language: French** (primary), WHO terminology. Numbers use a space as thousands separator and
  a comma as decimal (`1 155`, `26,3%`). Keep this in figures, tables and prose.
- **WHO palette**: blue `#0072BC`, dark `#1A3A5C`, red `#C8132A` (deaths / deterioration),
  green `#2E7D32` (recoveries / improvement), amber `#F4A81D` (caution / stable), grey `#D9D9D9`.
  Font Arial throughout.
- **Severity glyphs** in reviews/dashboards: ▼ red = deterioration, ▲ green = improvement,
  ▶ amber = stable; ⚠ for a data caveat, ❌ for an error, ℹ for a note.
- Match the look of the previous final/review in the project rather than inventing a new layout.

## Keeping the user informed (progress updates)

Several steps here run for a while with little or no streaming output — rendering the DOCX to PDF,
generating the figures, injecting images, extracting a large source file, or any subagent work. To
the user these look like a silent "Running command… · 11m 23s" spinner, and they cannot tell whether
anything is actually happening. That is unsettling and erodes trust. Keep them in the loop:

- **Announce long steps before you start them.** Before any step likely to take more than ~30 seconds,
  post one short line saying what you are about to do and that you will report back — e.g.
  "Rendering the enhanced DOCX to PDF to spot-check; this usually takes a minute or two."
- **Prefer several short commands over one long silent one.** Each command that returns is a visible
  heartbeat. Generate the figures, then inject, then render — as separate steps with a one-line result
  after each — rather than a single script that does everything in one multi-minute call.
- **Make long scripts print progress as they go, and flush.** Streamed output is itself a sign of life:
  print each figure as it is saved, each verification family as it is checked, each injection anchor as
  it is filled. `scripts/make_figures.py` already prints one line per figure for this reason.
- **Give a rough estimate when a step is unavoidably long** (a large LibreOffice render, a subagent
  run), then confirm when it finishes.
- **Status line after each stage.** One sentence: "Dashboard + 6 figures injected; rendering to check
  now." Don't disappear for long stretches.

The goal is simple: the user should always have a recent, plain-language sense of what is running and
roughly how long it will take. A steady trickle of short updates beats one perfect summary 11 minutes
later.

## Quick checklist

- [ ] Source SitRep and N-1 in hand (and N-2/N-3 if doing trend figures)?
- [ ] Stage decided (enhance / review / both)?
- [ ] Stage 1: dashboard + 6 figures injected, source formatting untouched, PDF spot-checked?
- [ ] Stage 2: all arithmetic families verified, reasoning checks done, report follows house format?
- [ ] Deaths reconciled (cumulative Δ = new-case deaths + CTE deaths), provincially and nationally?
- [ ] Any ND / missing province data flagged as breaking J-1 comparability?
- [ ] Output saved where the user can reach it; comparison to the prior final/review done?
