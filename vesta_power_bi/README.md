# Vesta Document Power BI Dashboard

A Power BI project (`.pbip`) for the `vesta_Document` / `vesta_DocumentVersion` tables, with a
Wells Fargo colour theme (red `#D71E28`, gold `#FFCD41`).

> **The data in `data/` is synthetic.** It was generated from the reconstruction rules in
> [docs/SPEC_transcribed.md](docs/SPEC_transcribed.md) and contains no real users, models or IDs.
> **This repository is public — never commit the real export.** Put real files in `data_private/`
> (ignored by git) as described below.

## Pages

- **Model Explorer** – searchable single-select *Model ID* dropdown. Selecting a model shows its
  documents, versions, latest state and version per document, Vesta commits, publications,
  lifecycle days, versions by state (donut), weekly version timeline, workflow-field coverage,
  version history table and document table.
- **Portfolio Overview** – all models: KPI cards, lifecycle state by document type, versions per
  week, models by risk tier / model type / usage category, workflow-field coverage %.

## Open it

1. Requires Power BI Desktop (tested with 2.159, Oct 2026).
2. Point the dashboard at the CSV folder on this computer (the path is stored as an absolute path):
   ```bash
   python scripts/set_data_folder.py
   ```
3. Open `VestaDashboard.pbip` and click **Refresh**.

You can also change the folder inside Power BI: *Transform data → Edit parameters → DataFolder*.

## Use the real data

1. Create `data_private/` in this folder (git ignores it) and copy in the real exports, named exactly
   `vesta_Document.csv` and `vesta_DocumentVersion.csv`, with the same column headers.
2. Close Power BI Desktop, then run:
   ```bash
   python scripts/set_data_folder.py data_private
   ```
3. Open `VestaDashboard.pbip` and click **Refresh**.

Before committing anything, run `git status` and confirm that no real CSVs are listed. The
`DataFolder` path in `expressions.tmdl` changes when you run the script; it is safe to commit, but it
shows the local folder path.

## Files

| Path | What it is |
|---|---|
| `VestaDashboard.pbip` | Open this in Power BI Desktop |
| `VestaDashboard.SemanticModel/` | Model in TMDL: tables, relationships, measures, `DataFolder` parameter |
| `VestaDashboard.Report/` | Report pages and visuals (PBIR) plus the Wells Fargo theme |
| `data/` | Synthetic CSVs (131 documents, 145 versions) |
| `docs/SPEC_transcribed.md` | Data reconstruction rules the synthetic data follows |
| `scripts/generate_data.py` | Regenerates the synthetic CSVs (deterministic) |
| `scripts/validate_data.py` | Checks a pair of CSVs against the spec (151 checks) |
| `scripts/build_pbip.py` | Regenerates the whole Power BI project from code (overwrites manual edits) |
| `scripts/set_data_folder.py` | Points the `DataFolder` parameter at a CSV folder |

## Model notes

- Relationships: `Models[model_id]` → `vesta_Document[model_id]` → `vesta_DocumentVersion[document_id]`;
  `Calendar[Date]` → `vesta_DocumentVersion[Created Date]`.
- `Models` is built in Power Query from the documents (one row per model; type, risk and usage are parsed
  from the template description, e.g. `MDD | Quantitative Model | Moderate risk | General`). If the real
  descriptions use a different layout, those three columns will be blank, but everything else still works.
- Blank CSV cells are loaded as nulls.
- `Calendar` covers 2026-06-01 to 2026-10-31. If the real data falls outside that range, widen it in
  `VestaDashboard.SemanticModel/definition/tables/Calendar.tmdl`.
