# bvd-sitrep-review

A two-stage skill for the COUSP / WHO AFRO daily SitRep of the 17th Ebola (MVE-Bundibugyo)
epidemic in the DRC.

- **Stage 1 — Enhance:** add a colour-coded day-over-day comparison dashboard and six analytical
  figures to a raw COUSP SitRep, preserving all original formatting.
- **Stage 2 — Review:** produce a technical review & data-quality report that verifies the
  arithmetic, reconciles deaths, flags impossible or missing values, and lists pre-dissemination
  corrections.

## Layout
```
bvd-sitrep-review/
├── SKILL.md                  # entry point: when to use, governance, the two-stage workflow
├── references/
│   ├── enhancement.md        # Stage 1 recipe: dashboard, figures, anchors, injection
│   └── review.md             # Stage 2 method: checks, reasoning, report structure, examples
└── scripts/
    └── make_figures.py       # renders the six figures from a JSON of the day's numbers
```

## Figure script
```bash
python scripts/make_figures.py --example > day.json   # template
# fill day.json with the SitRep's numbers, then:
python scripts/make_figures.py --data day.json --outdir fig/
```

Verification in Stage 2 is intentionally prose-and-judgment rather than a fixed script: the
highest-value findings are context calls (novel signals, missing data, misleading headlines), not
arithmetic slips. Compute the sums, but reason about what they mean.
