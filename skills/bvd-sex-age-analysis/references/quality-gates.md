# Quality gates — detailed rationale

These gates run automatically in `scripts/generate_slides.py::run_gates`. Each
encodes a lesson from a real data-quality issue in the DHIS2 MVE line-list. A
`FAIL` blocks slide generation; a `WARN` is footnoted but does not block.

### Gate 1 — Age code audit
```
FAIL if: any age_ans > 120 not in {9999, 1999, 999}
ACTION : add the new code to INVALID_AGES and re-run, or flag for manual review
```

### Gate 2 — Case total reconciliation
```
FAIL if: child_cases + adult_cases + unknown_cases ≠ total_confirmed_cases   (tolerance 0)
```

### Gate 3 — Death total reconciliation
```
FAIL if: child_deaths + adult_deaths + unknown_deaths ≠ total_deaths          (tolerance 0)
```

### Gate 4 — Age-group case reconciliation
```
FAIL if: sum(age_group_cases) + unknown_age_cases ≠ total_confirmed_cases     (tolerance 0)
```

### Gate 5 — Age-group death reconciliation
```
FAIL if: sum(age_group_deaths) + unknown_age_deaths ≠ total_deaths            (tolerance 0)
```

### Gate 6 — Province total reconciliation
```
FAIL if: sum(province_cases) ≠ total_confirmed_cases                          (tolerance 0)
```

### Gate 7 — Province child/adult reconciliation (cases & deaths)
```
FAIL if: any province where child + adult + unknown ≠ province_total          (tolerance 0)
```

### Gate 8 — Province death reconciliation
```
FAIL if: sum(province_deaths) ≠ total_deaths                                  (tolerance 0)
```

### Gate 9 — CFR sanity
```
FAIL if: any CFR < 0 or > 120
```

### Gate 10 — Gender completeness
```
WARN if: any confirmed case has a null/unexpected sexe value
ACTION : warning only; footnote added to affected slides
```

### Gate 11 / 12 — Gender-sum reconciliation
```
FAIL if: Female + Male + Unknown cases  ≠ total confirmed
FAIL if: Female + Male + Unknown deaths ≠ total deaths
```

---

## Known data issues (history)

| Issue | Root cause | Fix |
|---|---|---|
| Child + Adult ≠ total (15 cases missing) | `age_ans = 1999` and `999` not treated as invalid | added to `INVALID_AGES` |
| Deaths by age group ≠ total (15 missing) | invalid codes + `pd.cut` dropping age 0 | `include_lowest=True` + invalid-code fix |
| CFR displayed as `17.899999…` | float not rounded before display | `round(x, 1)` at extraction time |
| Zone narrative mismatched data | footnote hardcoded from an earlier dataset | all footnotes now computed from live data |
| Province child/adult subtotals ≠ province total | unknown-age cases not tracked per province | `Unknown` bucket tracked in `cag()` |
