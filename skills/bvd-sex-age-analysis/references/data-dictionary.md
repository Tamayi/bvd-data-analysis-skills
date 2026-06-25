# DHIS2 Tracker line-list — data dictionary

Columns consumed by `extract()`. If a new export renames a column, update
`extract()` in `scripts/generate_slides.py` (and promote the rename map to
`bvd_common` if other skills read the same export).

| Column | Type | Purpose |
|---|---|---|
| `classification_finale` | text | Case classification. Confirmed = `"Cas confirmé"` |
| `age_ans` | number | Age in years. Codes `9999`, `1999`, `999` and `NaN` → Unknown |
| `sexe` | text | `"Feminin"` / `"Masculin"`; anything else → Unknown sex |
| `nature_alerte` | text | `"Décès"` marks a death |
| `date_deces` | date | Date of death (supplementary death signal) |
| `date_deces_alerte` | date | Date of death recorded at alert (supplementary) |
| `date_debut_symptomes_notification` | date | Symptom onset (primary) |
| `date_debut_symptomes` | date | Symptom onset (fallback) |
| `date_heure_investigation` | datetime | Investigation timestamp (future-date check) |
| `date_prelevement` | date | Sample collection (future-date check) |
| `Province` | text | Province; expected {Ituri, Nord Kivu, Sud Kivu} |
| `Zone_sante` | text | Health zone |
| `Aire_sante` | text | Health area (optional) |
| `numero_identification_alerte` | text | Case ID (duplicate check) |

## Derived fields (added in `extract()`)

| Field | Definition |
|---|---|
| `age_clean` | `age_ans` with invalid codes → `NaN` |
| `sexe_clean` | `Feminin` / `Masculin` / `Unknown` |
| `is_death` | `nature_alerte == "Décès"` OR `date_deces` notna OR `date_deces_alerte` notna |
| `ag` | 10-bin age group (`pd.cut`, `include_lowest=True`) |
| `a2` | `Child` (<15) / `Adult` (≥15) / `Unknown` |
