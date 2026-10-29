# Vesta Document Power BI Dashboard

A Power BI project (`.pbip`) for the `vesta_Document` / `vesta_DocumentVersion` tables, with a
Wells Fargo colour theme (red `#D71E28`, gold `#FFCD41`).

> **The data in `data/` is synthetic.** It was generated from the reconstruction rules in
> [docs/SPEC_transcribed.md](docs/SPEC_transcribed.md) and contains no real users, models or IDs.
> **This repository is public — never commit the real export.** Put real files in `data_private/`
> (ignored by git) as described below.

**Quick start:** copy the two CSVs into `data_private/`, then double-click `Open Dashboard.bat`.

## Pages

- **Model Explorer** – searchable single-select *Model ID* dropdown. Selecting a model shows its
  documents, versions, latest state and version per document, Vesta commits, publications,
  lifecycle days, versions by state (donut), weekly version timeline, workflow-field coverage,
  version history table and document table.
- **Portfolio Overview** – all models: KPI cards, lifecycle state by document type, versions per
  week, models by risk tier / model type / usage category, workflow-field coverage %.

## Open it (one double-click)

Requires Power BI Desktop on Windows (tested with 2.159, Oct 2026). No Python needed.

1. Copy the real exports into **`data_private/`**, named exactly `vesta_Document.csv` and
   `vesta_DocumentVersion.csv` (same column headers as the files in `data/`).
2. Double-click **`Open Dashboard.bat`**.

The launcher points the dashboard at `data_private/` (or at the synthetic `data/` if
`data_private/` has no CSVs), opens `VestaDashboard.pbip` in Power BI Desktop, waits for it to load and
refreshes the data automatically. Close any open copy of the dashboard before running it. If the
automatic refresh ever fails, click **Refresh** on the Home ribbon.

Git ignores everything in `data_private/` except its README, so the real data is never committed.
Still, run `git status` before any commit and confirm that no real CSVs are listed. The launcher
rewrites the `DataFolder` path in `expressions.tmdl`; that change is safe to commit, but it shows the
local folder path.

You can also change the folder by hand: *Transform data → Edit parameters → DataFolder*, or
`python scripts/set_data_folder.py <folder>`.

## Files

| Path | What it is |
|---|---|
| `Open Dashboard.bat` | Double-click: picks the CSV folder, opens the dashboard, refreshes it |
| `data_private/` | Drop the real CSVs here (git-ignored) |
| `VestaDashboard.pbip` | The Power BI project (opened by the launcher) |
| `VestaDashboard.SemanticModel/` | Model in TMDL: tables, relationships, measures, `DataFolder` parameter |
| `VestaDashboard.Report/` | Report pages and visuals (PBIR) plus the Wells Fargo theme |
| `data/` | Synthetic CSVs (131 documents, 145 versions) |
| `docs/SPEC_transcribed.md` | Data reconstruction rules the synthetic data follows |
| `scripts/generate_data.py` | Regenerates the synthetic CSVs (deterministic) |
| `scripts/validate_data.py` | Checks a pair of CSVs against the spec (151 checks) |
| `scripts/build_pbip.py` | Regenerates the whole Power BI project from code (overwrites manual edits) |
| `scripts/open_dashboard.ps1` | What the launcher runs |
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
