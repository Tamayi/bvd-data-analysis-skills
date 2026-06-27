# Stage 1 — Enhancement recipe

Produce an enriched copy of the source SitRep DOCX: a comparison dashboard plus six figures,
injected at fixed anchor points, with everything else preserved. Below is exactly what to build,
where it goes, and how to word it. Match the previous "FINAL" in the project for look and feel.

## Contents
1. The comparison dashboard
2. The six figures (data, captions, placement)
3. Rendering the figures (`scripts/make_figures.py`)
4. Injection into the DOCX
5. Verification

---

## 1. The comparison dashboard

A 5-column colour-coded table inserted **immediately after section "1. FAITS SAILLANTS" and
before "2. CONTEXTE"**. Title above it:
`Tableau de bord comparatif : SitRep #<N-1> (<date N-1>) vs SitRep #<N> (<date N>)`.

Columns: `Indicateur | SitRep #<N-1> | SitRep #<N> | Évolution | Évaluation`.
Header row fill `1A3A5C`, white bold Arial 8pt. Indicator column fill `F2F2F2`, dark bold.

Row colour by direction (fill / text):
- deterioration ▼ → `FFCCCC` / `C8132A`
- improvement ▲ → `CCFFCC` / `2E7D32`
- stable ▶ → `FFF3CD` / `1A3A5C`

Legend line below the table, italic grey 7.5pt:
`▼ Rouge = dégradation  |  ▲ Vert = amélioration  |  ▶ Ambre = stable  |  * <caveat>  |  Source : COUSP SitRep N°<N>, <date>`.

Indicator rows (use the ones the data supports; ~18–19 is typical). The Évaluation cell is a short
sentence of context, not just a number — that is what makes the dashboard useful:
1. Cas confirmés cumulés (RDC)
2. Décès confirmés (TL %) — note the new-case vs CTE split in the evaluation text
3. Guérisons cumulées
4. Patients en isolement (taux d'occupation) — if a province is ND, mark the total non-comparable
5. Taux de suivi des contacts (RDC) — show gap to the 95% target
6. Nouveaux cas (24h)
7. Nouveaux décès (24h) — spell out the new-case + CTE composition
8. Ituri : cas cumulés
9. Ituri : taux de suivi des contacts
10. Ituri : alertes investiguées (rate, n/n)
11. Ituri : taux d'occupation CTE/CT
12. Nord-Kivu : cas cumulés
13. Nord-Kivu : taux de suivi des contacts
14. Nord-Kivu : alertes investiguées
15. Lab Ituri : positivité
16. Lab Nord-Kivu : backlog (reçus − analysés)
17. Sud-Kivu : alertes investiguées
18. Alertes investiguées (national)
19. Zones de santé touchées (national)

**Critical:** pull each #<N-1> value from the real previous SitRep, not by assuming "same as today".
A prior final once mislabelled a Sud-Kivu alert rate as stable when the previous day was actually
0% (fuel shortage) — an improvement shown as no-change. Always diff against the true N-1 value.

---

## 2. The six figures

All in the WHO palette, French labels, comma decimals. Titles below each one are the **caption**
(bold, dark, centred, 8.5pt) and a **source note** (italic grey, centred, 8pt).

- **Figure 1 — Daily cases / deaths / recoveries** (grouped bars, last ~4 days).
  Cases blue, deaths red, recoveries green; annotate values; "ND" where a series is unavailable.
  Caption: `Figure 1. Cas confirmés, décès et guérisons par jour — RDC, <range>`.
  Note: `⚠ Variation journalière à interpréter avec prudence. Ancres vérifiées SitReps #<..>–#<N>. Source : COUSP SitRep N°<N>.`

- **Figure 2 — Cumulative trajectory + CFR** (lines: national total, Ituri filled, NK, SK; CFR on a
  right axis in orange `#E65100`). Annotate the final total and CFR.
  Caption: `Figure 2. Trajectoire cumulée des cas confirmés et taux de létalité — RDC, <range>`.

- **Figure 3 — Daily cases by health zone vs national total** (stacked bars of the top zones +
  "Autres ZS", with a dashed national-total line). Use the per-ZS daily counts from the Faits
  Saillants of the last few SitReps.
  Caption: `Figure 3. Cas confirmés journaliers par zone de santé vs total national — RDC, <range>`.

- **Figure 4 — Contact-tracing rate + contacts under follow-up** (province lines + national dashed,
  grey bars for contacts under follow-up on a second axis, dotted 95% target line, annotate gap).
  Caption: `Figure 4. Taux de suivi des contacts et contacts sous suivi — RDC, <range>`.

- **Figure 5 — Nord-Kivu profile by health zone** (cas + décès bars per NK zone, CFR markers, mean
  NK CFR line). This is the NK-by-zone profile — *not* an age/sex breakdown.
  Caption: `Figure 5. Profil Nord-Kivu : cas confirmés, décès et TL par zone de santé — <date>`.

- **Figure 6 — Laboratory capacity & backlog** (left: analysed / positives / backlog bars per
  province + positivity line; right: NK backlog over the last few days).
  Caption: `Figure 6. Capacité diagnostique et backlog laboratoire — RDC, <date>`.

**Placement anchors** (insert the caption + image + note paragraphs *before* the anchor paragraph):
- Figures 1, 2, 3 → before `Tableau 2.`
- Figure 4 → before `4.3. Laboratoire`
- Figure 6 → before `4.4. Prévention`
- Figure 5 → before `Tableau 7.`
- Dashboard → before `2. CONTEXTE`

Note the document already has its own "Figure 1. Chronologie" and "Figures 2 & 3. Spatialisation";
the injected figures reuse the numbers 1–6 as their own series, matching prior finals.

**Trend data — prefer verified anchors over interpolation.** For the daily/cumulative figures,
read the headline numbers from each of the last consecutive SitReps (e.g. #N-3..#N) and use those
real points. Only interpolate a gap between two real anchors, and when you do, say so in the note.
Daily new cases = difference of cumulative totals between consecutive days (cross-check against the
Faits Saillants "nouveaux cas" figure).

---

## 3. Rendering the figures

`scripts/make_figures.py` renders all six PNGs (300 DPI, WHO palette, French formatting) from one
JSON file. Write the day's numbers to a JSON, then run:

```bash
python scripts/make_figures.py --data day.json --outdir fig/
```

See the `--help` and the example JSON the script prints with `--example`. Supplying the data as
JSON keeps the styling identical across SitReps while you only change the numbers. If a figure
needs a layout the script doesn't cover, adapt the script rather than hand-rolling a new style —
consistency across days matters more than per-day cleverness.

---

## 4. Injection into the DOCX

Work on a copy of the source so the original is untouched. `python-docx` is the most reliable path
and preserves the source's logos, maps, photos, tables and text:

- Find an anchor paragraph by its text (e.g. the paragraph containing `Tableau 2.`).
- Build the dashboard with `add_table`, shade cells via `w:shd`, then move it before the anchor with
  `anchor._p.addprevious(table._tbl)`.
- For each figure, create caption / image / note paragraphs and move them before the anchor with
  `addprevious`. Add the picture with `run.add_picture(path, width=Inches(6.3))` (fits A4 margins).
- Save to a new filename like `SitRep_MVE_RDC_N<NN>_<DD>Mon2026_FINAL.docx`.

(If byte-exact preservation of every source element is required, the alternative is to unpack the
DOCX, edit the XML, and repack — see the docx skill. `python-docx` is usually enough.)

## 5. Verification

*(This render is the slowest step on a large SitRep — say so before you start it and confirm when
it finishes.)* Render to PDF (`soffice --headless --convert-to pdf`) and convert pages to images. Check: page 1
unchanged; the dashboard renders with correct colours; each figure sits before its anchor with the
caption and note; numbers in the dashboard match the source tables. Fix and re-render before
delivering.
