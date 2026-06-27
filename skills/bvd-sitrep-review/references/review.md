# Stage 2 — Technical review & data-quality method

Produce a short Word report that an Incident Manager can skim in two minutes and know exactly what
is correct, what is wrong, and what to fix before the SitRep circulates. Match the house format of
the previous review in the project.

This is **verification by reasoning**, not by running a fixed script. Do the arithmetic and show it,
but the point of the review is judgment: an arithmetically perfect SitRep can still mislead. The
most valuable findings to date were context calls, not failed sums. Approach each SitRep fresh and
ask "what here could be wrong, misread, or missing?" — the families below are a floor, not a ceiling.

## Contents
1. Arithmetic families to verify
2. Reasoning checks (where the real value is)
3. Severity scale
4. Report structure
5. Worked examples from real SitReps

---

## 1. Arithmetic families to verify

Recompute each and compare to the printed value; show the computation in the report's check table.

- **Province → national sums**: Ituri + NK + SK = national, for cases, deaths, and new cases (T1, T2).
- **Health-zone → province sums** (T2): the listed zones plus "autres non identifiées" equal the
  province total, for cases and deaths.
- **CFR (taux de létalité)** = deaths / cases × 100, per zone, per province, national. Verify to 1 dp.
- **Contact follow-up rate** (T4) = contacts vus / contacts sous suivi × 100, per province + national.
- **Alert investigation rate** (T3) = investiguées / (report J-1 + nouvelles) × 100; also check
  report + nouvelles = total alerts.
- **Isolation balance** (T6): end-of-day = J-1 + admissions − total exits; confirmed + suspect = total
  in isolation; total exits = deaths + non-cases + recoveries + escapes + transfers.
- **Bed occupancy** = in isolation / capacity (capacity is implied; sanity-check it's stable).
- **Lab positivity** = positives / analysed × 100; recompute and check the rounding matches.
- **Lab backlog** = samples received − samples analysed; confirm it's stated (it is often omitted).
- **Weighted-average sanity**: any "global/total" rate must lie between its component values. A total
  below the minimum component (or above the maximum) is arithmetically impossible — flag it and give
  the expected weighted value.
- **Banner / Faits Saillants deltas**: cumulative cases, deaths, recoveries vs the previous SitRep.

## 2. Reasoning checks (where the real value is)

- **Death reconciliation — do this every time.** The Faits Saillants daily death number frequently
  counts only deaths *among new confirmed cases*. The cumulative death total can rise by more because
  patients already confirmed die in CTEs. Reconcile:
  `Δ cumulative deaths = (deaths among new confirmed cases) + (deaths in CTE among already-confirmed)`,
  and check it provincially too. If the headline number ≠ the cumulative delta and the gap isn't
  explained, that's an error (see example A). If it *is* explained but the headline still understates
  the day's mortality, recommend clarifying wording (see example B).
- **Missing / ND data breaks comparability.** If a province's table is "ND", national totals that
  exclude it are not comparable to the previous day. Flag explicitly and warn against a naive J-1 delta
  (e.g. isolation total dropping only because one province is missing).
- **Narrative vs table.** Check that percentages and "concentration" claims in the prose (e.g. "the top
  three zones hold 75,5% of cases") actually recompute from the tables.
- **Novel epidemiological signals.** Watch for events that change the risk picture: a case exported to
  another country, a death outside the affected provinces, a newly affected health zone, a sudden
  positivity or alert-rate shift. Confirm each is correctly *included or excluded* from the counts, and
  surface it — these matter more than a rounding error.
- **Cross-document consistency.** If a Draft, an IM summary (Résumé IM) and a final PDF are all present,
  confirm the headline indicators agree across them; an error in one usually propagates.
- **Operational red flags.** A large single-day drop in alert investigation, a rebounding lab backlog,
  near-saturation bed occupancy — call these out even when the arithmetic is fine.

## 3. Severity scale

- ❌ **Erreur critique** — a wrong published number (e.g. an understated death count, an impossible
  rate). Must be fixed before dissemination.
- ⚠ **À corriger / à signaler** — not strictly wrong but misleading or incomplete (ND data not flagged,
  backlog not quantified, an operational drop not surfaced).
- ℹ **Note** — clarity or rounding improvements, denominator clarifications, items to confirm.

## 4. Report structure

A Word document in the house style (Arial; WHO palette; coloured status cells: green `CCFFCC` = OK,
red `FFCCCC` = error, amber `FFF3CD` = note; header fill `1A3A5C`). Sections:

1. **Title block** — "COUSP RDC — … / Revue Technique & Contrôle Qualité — SitRep MVE N°<N> / <date>"
   and a row of "Documents examinés | Date de revue".
2. **1. Résultat global** — a short shaded paragraph: overall verdict, count of errors/notes, and the
   one or two things that matter most. Then a small **delta-verification table** (#N-1 → #N: indicator,
   previous, current, Δ, status).
3. **2. Erreurs et incohérences** — one subsection per finding, ordered by severity, each with
   *Localisation*, *Problème*, and a *Correction / Texte suggéré*. Be specific and quote the source text.
4. **3. Vérifications arithmétiques — résultats complets** — a table of every check (Source |
   Vérification | Statut) with a one-line summary above it ("N checks, M correct, K issues").
5. **4. Figures et tableau de bord** — the enhancement items (dashboard + 6 figures) listed as the
   recommended/added value, noting they use only data already in the SitReps.
6. **5. Checklist de validation avant diffusion** — a Priorité | Action table (❌ urgent, ⚠ to-do,
   ℹ optional), then the italic line that this is an automated review to be validated by the
   surveillance/IM lead before dissemination.

Announce the build and the PDF render before running them (they are the slow steps; see "Keeping the
user informed" in SKILL.md), and report when each finishes. Build it with `python-docx` (shade cells via `w:shd`, set table borders, fixed column widths) and
render to PDF to confirm the colours and layout before delivering.

## 5. Worked examples from real SitReps

**Example A — critical death error (#040).** Faits Saillants said "24 nouveaux cas dont 12 décès",
but cumulative deaths went 277 → 291 (+14). The missing 2 were NK deaths (Kyondo, Musienene) visible
in Tableau 1's own footnote. Finding: ❌ correct the headline to 14 and add "Nord-Kivu (1 dont 2 décès)".

**Example B — death headline understates the day (#041).** Faits Saillants said "37 nouveaux cas dont
5 décès", and cumulative deaths rose 291 → 304 (+13). Here it *reconciles*: 5 deaths among new cases +
8 deaths in Ituri CTEs among already-confirmed patients = 13, and Tableau 1's note documents the 8.
So not an error — but the "5" understates the day's 13 deaths. Finding: ℹ recommend the FS state
"+13 au cumul (5 nouveaux cas + 8 en CTE)".

**Example C — impossible weighted average (#041).** Tableau 5 (PoE/PoC) "% voyageurs sensibilisés":
Ituri 94,7%, NK 97,0%, Global 89,3%. A global below both components is impossible; with the throughput
weights it should be ≈96,2%. The "screened" and "handwashing" rows on the same table were correct.
Finding: ❌/⚠ recompute the global.

**Example D — ND breaks comparability (#041).** NK patient-movement (Tableau 6) was entirely "ND", so
the isolation total (326) and occupancy (79,3%) covered only Ituri + SK and were not comparable to
#040's 408 / 81,0%. Finding: ⚠ add a Tableau 6 note; don't compute a J-1 isolation delta.

**Example E — novel signal (#041).** A confirmed case was exported to France (a clinician who had
worked in Bunia) and an INSP expert died in Kinshasa (MVE-negative). Both correctly excluded from RDC
totals. Finding: ℹ surface the France export as a cross-border signal to confirm via IHR/WHO; note the
Kinshasa death is correctly outside the MVE count.

These are illustrations of the *kinds* of things to look for — reason about the specific SitRep in
front of you rather than checking only for these exact cases.
