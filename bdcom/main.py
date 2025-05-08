"""
dash_code.py
------------
Single-click a cell in the Summary table to

1.  Show only that field’s rows in the Value-Distribution
    and Population-Comparison tabs.
2.  Display the field’s `value_sql_logic` text under each table
    (the logic is assumed identical for all rows of the same field).

Requires: Python ≥3.8 · Dash ≥2.6 · Pandas ≥1.3
"""

import datetime as dt
import re

import pandas as pd
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output, State
import dash

# ────────────────────────────────────────────────────────────────
# 1.  Load the workbook
# ────────────────────────────────────────────────────────────────
INPUT_FILE = "input.xlsx"           # adjust path if needed
df_data    = pd.read_excel(INPUT_FILE, sheet_name="Data")

df_data["filemonth_dt"] = pd.to_datetime(
    df_data["filemonth_dt"], format="%m/%d/%Y"
)

# ────────────────────────────────────────────────────────────────
# 2.  Reference months
# ────────────────────────────────────────────────────────────────
DATE1 = dt.datetime(2025,  1, 1)
DATE2 = dt.datetime(2024, 12, 1)

# ────────────────────────────────────────────────────────────────
# 3.  Helper for pop-comp phrases
# ────────────────────────────────────────────────────────────────
PHRASES = [
    r"1\)\s*CF Loan - Both Pop, Diff Values",
    r"2\)\s*CF Loan - Prior Null, Current Pop",
    r"3\)\s*CF Loan - Prior Pop, Current Null",
]
def contains_phrase(x: str) -> bool:
    return any(re.search(p, str(x)) for p in PHRASES)

# ────────────────────────────────────────────────────────────────
# 4.  Build the Summary dataframe
# ────────────────────────────────────────────────────────────────
rows = []
for fld in sorted(df_data["field_name"].unique()):
    # Missing counts
    miss1 = df_data[
        (df_data["analysis_type"] == "value_dist") &
        (df_data["field_name"]   == fld)           &
        (df_data["filemonth_dt"] == DATE1)         &
        (df_data["value_label"].str.contains("Missing", case=False, na=False))
    ]["value_records"].sum()

    miss2 = df_data[
        (df_data["analysis_type"] == "value_dist") &
        (df_data["field_name"]   == fld)           &
        (df_data["filemonth_dt"] == DATE2)         &
        (df_data["value_label"].str.contains("Missing", case=False, na=False))
    ]["value_records"].sum()

    # Pop-comp diffs
    diff1 = df_data[
        (df_data["analysis_type"] == "pop_comp")   &
        (df_data["field_name"]   == fld)           &
        (df_data["filemonth_dt"] == DATE1)         &
        (df_data["value_label"].apply(contains_phrase))
    ]["value_records"].sum()

    diff2 = df_data[
        (df_data["analysis_type"] == "pop_comp")   &
        (df_data["field_name"]   == fld)           &
        (df_data["filemonth_dt"] == DATE2)         &
        (df_data["value_label"].apply(contains_phrase))
    ]["value_records"].sum()

    rows.append([fld, miss1, miss2, diff1, diff2])

df_summary = pd.DataFrame(
    rows,
    columns=[
        "Field Name",
        f"Missing {DATE1:%m/%d/%Y}",
        f"Missing {DATE2:%m/%d/%Y}",
        f"M2M Diff {DATE1:%m/%d/%Y}",
        f"M2M Diff {DATE2:%m/%d/%Y}",
    ],
)

# ────────────────────────────────────────────────────────────────
# 5.  Detail tables (only the two months)
# ────────────────────────────────────────────────────────────────
mask_two_months = df_data["filemonth_dt"].isin([DATE1, DATE2])

vd_all = df_data[
    mask_two_months & (df_data["analysis_type"] == "value_dist")
].copy()

pc_all = df_data[
    mask_two_months & (df_data["analysis_type"] == "pop_comp")
].loc[lambda d: d["value_label"].apply(contains_phrase)].copy()

# Display-friendly date strings
vd_all["filemonth_dt"] = vd_all["filemonth_dt"].dt.strftime("%m/%d/%Y")
pc_all["filemonth_dt"] = pc_all["filemonth_dt"].dt.strftime("%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 6.  Dash layout
# ────────────────────────────────────────────────────────────────
app = Dash(__name__)
app.layout = html.Div([
    html.H2("BDCOMM FRY14M Field Analysis"),
    dcc.Tabs([
        # ── Summary ──────────────────────────────────────────
        dcc.Tab(label="Summary", children=[
            dash_table.DataTable(
                id="summary",
                columns=[{"name": c, "id": c} for c in df_summary.columns],
                data=df_summary.to_dict("records"),
                page_size=20,
                style_header={"backgroundColor": "#4F81BD",
                              "color": "white", "fontWeight": "bold"},
                style_table={"overflowX": "auto"},
            )
        ]),
        # ── Value Distribution ──────────────────────────────
        dcc.Tab(label="Value Distribution", children=[
            dash_table.DataTable(
                id="vd",
                columns=[{"name": c, "id": c}
                         for c in vd_all.columns if c != "value_sql_logic"],
                data=vd_all.to_dict("records"),
                page_size=20,
                style_header={"backgroundColor": "#4F81BD",
                              "color": "white", "fontWeight": "bold"},
                style_table={"overflowX": "auto"},
            ),
            html.Pre(id="vd_sql", style={
                "whiteSpace": "pre-wrap",
                "backgroundColor": "#f3f3f3",
                "padding": "0.75rem",
                "border": "1px solid #ddd",
                "marginTop": "0.5rem",
                "fontFamily": "monospace",
                "fontSize": "0.9rem"
            })
        ]),
        # ── Population Comparison ───────────────────────────
        dcc.Tab(label="Population Comparison", children=[
            dash_table.DataTable(
                id="pc",
                columns=[{"name": c, "id": c}
                         for c in pc_all.columns if c != "value_sql_logic"],
                data=pc_all.to_dict("records"),
                page_size=20,
                style_header={"backgroundColor": "#4F81BD",
                              "color": "white", "fontWeight": "bold"},
                style_table={"overflowX": "auto"},
            ),
            html.Pre(id="pc_sql", style={
                "whiteSpace": "pre-wrap",
                "backgroundColor": "#f3f3f3",
                "padding": "0.75rem",
                "border": "1px solid #ddd",
                "marginTop": "0.5rem",
                "fontFamily": "monospace",
                "fontSize": "0.9rem"
            })
        ]),
    ])
])

# ────────────────────────────────────────────────────────────────
# 7.  Single callback – tables + SQL logic
# ────────────────────────────────────────────────────────────────
@app.callback(
    Output("vd",     "data"),
    Output("pc",     "data"),
    Output("vd_sql", "children"),
    Output("pc_sql", "children"),
    Input("summary", "active_cell"),    # single-click trigger
    State("summary", "data"),
    prevent_initial_call=True
)
def filter_and_show_sql(active, rows):
    """Filter tables and show value_sql_logic for the clicked field."""
    # Default: no filter, blank SQL boxes
    if active is None:
        return (vd_all.to_dict("records"),
                pc_all.to_dict("records"),
                "",
                "")

    field_clicked = rows[active["row"]]["Field Name"]

    vd_filtered = vd_all[vd_all["field_name"] == field_clicked]
    pc_filtered = pc_all[pc_all["field_name"] == field_clicked]

    # Grab the SQL logic (identical within each field). Use .get() in case
    # the column is missing.
    vd_sql_text = ""
    pc_sql_text = ""
    if "value_sql_logic" in vd_filtered.columns:
        vd_sql_text = vd_filtered["value_sql_logic"].iloc[0] or ""
    if "value_sql_logic" in pc_filtered.columns:
        pc_sql_text = pc_filtered["value_sql_logic"].iloc[0] or ""

    return (vd_filtered.to_dict("records"),
            pc_filtered.to_dict("records"),
            vd_sql_text,
            pc_sql_text)

# ────────────────────────────────────────────────────────────────
# 8.  Run the server
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)