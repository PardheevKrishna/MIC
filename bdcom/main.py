"""
dash_ag_grid_code.py
====================
Same functionality as before but with Dash AG Grid, so every column now
gets an Excel-style “value list” filter.

Requires
--------
pip install dash dash-ag-grid pandas xlrd  (plus openpyxl if your xlsx uses it)

Tested with Dash 2.17 · dash-ag-grid 30.x · Pandas 2.x
"""

import datetime as dt
import re
import pandas as pd
from dash import Dash, dcc, html
from dash.dependencies import Input, Output, State
import dash_ag_grid as dag
import dash  # for callback_context

# ────────────────────────────────────────────────────────────────
# 1.  Load workbook
# ────────────────────────────────────────────────────────────────
INPUT_FILE = "input.xlsx"
df_data = pd.read_excel(INPUT_FILE, sheet_name="Data")
df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"],
                                         format="%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 2.  Two months to compare
# ────────────────────────────────────────────────────────────────
DATE1 = dt.datetime(2025, 1, 1)
DATE2 = dt.datetime(2024, 12, 1)

# ────────────────────────────────────────────────────────────────
# 3.  Helper for pop-comp flag rows
# ────────────────────────────────────────────────────────────────
PHRASES = [
    r"1\)\s*CF Loan - Both Pop, Diff Values",
    r"2\)\s*CF Loan - Prior Null, Current Pop",
    r"3\)\s*CF Loan - Prior Pop, Current Null",
]
contains_phrase = lambda x: any(re.search(p, str(x)) for p in PHRASES)

# ────────────────────────────────────────────────────────────────
# 4.  Build Summary frame (adds comment cols)
# ────────────────────────────────────────────────────────────────
summary_rows = []
for fld in sorted(df_data["field_name"].unique()):
    miss1 = df_data[(df_data["analysis_type"] == "value_dist") &
                    (df_data["field_name"] == fld) &
                    (df_data["filemonth_dt"] == DATE1) &
                    (df_data["value_label"]
                     .str.contains("Missing", case=False, na=False))
                    ]["value_records"].sum()

    miss2 = df_data[(df_data["analysis_type"] == "value_dist") &
                    (df_data["field_name"] == fld) &
                    (df_data["filemonth_dt"] == DATE2) &
                    (df_data["value_label"]
                     .str.contains("Missing", case=False, na=False))
                    ]["value_records"].sum()

    diff1 = df_data[(df_data["analysis_type"] == "pop_comp") &
                    (df_data["field_name"] == fld) &
                    (df_data["filemonth_dt"] == DATE1) &
                    (df_data["value_label"].apply(contains_phrase))
                    ]["value_records"].sum()

    diff2 = df_data[(df_data["analysis_type"] == "pop_comp") &
                    (df_data["field_name"] == fld) &
                    (df_data["filemonth_dt"] == DATE2) &
                    (df_data["value_label"].apply(contains_phrase))
                    ]["value_records"].sum()

    summary_rows.append([fld, miss1, miss2, "", diff1, diff2, ""])

df_summary = pd.DataFrame(summary_rows, columns=[
    "Field Name",
    f"Missing {DATE1:%m/%d/%Y}",
    f"Missing {DATE2:%m/%d/%Y}",
    "Comment Missing",
    f"M2M Diff {DATE1:%m/%d/%Y}",
    f"M2M Diff {DATE2:%m/%d/%Y}",
    "Comment M2M",
])

# ────────────────────────────────────────────────────────────────
# 5.  Detail frames (only the two months)
# ────────────────────────────────────────────────────────────────
mask_2m = df_data["filemonth_dt"].isin([DATE1, DATE2])

vd_all = df_data[mask_2m & (df_data["analysis_type"] == "value_dist")].copy()
pc_all = df_data[mask_2m & (df_data["analysis_type"] == "pop_comp")
                 ].loc[lambda d: d["value_label"]
                       .apply(contains_phrase)].copy()

vd_all["filemonth_dt"] = vd_all["filemonth_dt"].dt.strftime("%m/%d/%Y")
pc_all["filemonth_dt"] = pc_all["filemonth_dt"].dt.strftime("%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 6.  Column definitions helper  (AG Grid)
# ────────────────────────────────────────────────────────────────
def make_col_defs(df, hide_sql_logic=False):
    """Return columnDefs with set filter & correct editability."""
    defs = []
    for c in df.columns:
        if hide_sql_logic and c == "value_sql_logic":
            continue
        defs.append({
            "headerName": c,
            "field":      c,
            "filter":     "agSetColumnFilter",  # checkbox list
            "floatingFilter": True,
            "sortable":   True,
            "editable":   c.startswith("Comment"),   # only comment cols editable
        })
    return defs

# ────────────────────────────────────────────────────────────────
# 7.  Dash layout
# ────────────────────────────────────────────────────────────────
def comment_box(text_id, btn_id):
    return html.Div([
        dcc.Textarea(id=text_id, placeholder="Add comment…",
                     style={"width": "100%", "height": "60px"}),
        html.Button("Submit", id=btn_id, n_clicks=0,
                    style={"marginTop": "0.25rem"})
    ], style={"marginTop": "0.5rem"})

app = Dash(__name__)
app.layout = html.Div([
    html.H2("BDCOMM FRY14M Field Analysis (AG Grid)"),
    dcc.Tabs([
        # ─── Summary ────────────────────────────────────────
        dcc.Tab(label="Summary", children=[
            dag.AgGrid(
                id="summary",
                columnDefs=make_col_defs(df_summary),
                rowData=df_summary.to_dict("records"),
                rowSelection="single",
                className="ag-theme-alpine",
                defaultColDef={"resizable": True},
                pagination=True, paginationPageSize=20,
            )
        ]),
        # ─── Value Distribution ─────────────────────────────
        dcc.Tab(label="Value Distribution", children=[
            dag.AgGrid(
                id="vd",
                columnDefs=make_col_defs(vd_all, hide_sql_logic=True),
                rowData=vd_all.to_dict("records"),
                className="ag-theme-alpine",
                defaultColDef={"resizable": True},
                pagination=True, paginationPageSize=20,
            ),
            comment_box("vd_comm_text", "vd_comm_btn"),
            html.Pre(id="vd_sql", style={"whiteSpace": "pre-wrap",
                                         "backgroundColor": "#f3f3f3",
                                         "padding": "0.75rem",
                                         "border": "1px solid #ddd",
                                         "marginTop": "0.5rem",
                                         "fontFamily": "monospace",
                                         "fontSize": "0.9rem"})
        ]),
        # ─── Population Comparison ─────────────────────────
        dcc.Tab(label="Population Comparison", children=[
            dag.AgGrid(
                id="pc",
                columnDefs=make_col_defs(pc_all, hide_sql_logic=True),
                rowData=pc_all.to_dict("records"),
                className="ag-theme-alpine",
                defaultColDef={"resizable": True},
                pagination=True, paginationPageSize=20,
            ),
            comment_box("pc_comm_text", "pc_comm_btn"),
            html.Pre(id="pc_sql", style={"whiteSpace": "pre-wrap",
                                         "backgroundColor": "#f3f3f3",
                                         "padding": "0.75rem",
                                         "border": "1px solid #ddd",
                                         "marginTop": "0.5rem",
                                         "fontFamily": "monospace",
                                         "fontSize": "0.9rem"})
        ]),
    ])
])

# ────────────────────────────────────────────────────────────────
# 8.  Master callback
# ────────────────────────────────────────────────────────────────
@app.callback(
    Output("vd", "rowData"),
    Output("pc", "rowData"),
    Output("vd_sql", "children"),
    Output("pc_sql", "children"),
    Output("summary", "rowData"),

    Input("summary", "cellClicked"),         # ← AG Grid click
    Input("vd_comm_btn", "n_clicks"),
    Input("pc_comm_btn", "n_clicks"),

    State("summary", "rowData"),
    State("vd", "rowData"),
    State("pc", "rowData"),
    State("vd_comm_text", "value"),
    State("pc_comm_text", "value"),

    prevent_initial_call=True
)
def master_callback(cell_click,
                    n_vd, n_pc,
                    s_rows, vd_rows, pc_rows,
                    vd_comment, pc_comment):

    s_df = pd.DataFrame(s_rows)
    trig = dash.callback_context.triggered[0]["prop_id"].split(".")[0]

    # ── A. Append comments ──────────────────────────────────
    if trig == "vd_comm_btn" and vd_comment:
        field = vd_rows[0]["field_name"] if vd_rows else None
        if field:
            m = s_df["Field Name"] == field
            old = s_df.loc[m, "Comment Missing"].iloc[0]
            s_df.loc[m, "Comment Missing"] = (old + "\n" if old else "") + vd_comment

    if trig == "pc_comm_btn" and pc_comment:
        field = pc_rows[0]["field_name"] if pc_rows else None
        if field:
            m = s_df["Field Name"] == field
            old = s_df.loc[m, "Comment M2M"].iloc[0]
            s_df.loc[m, "Comment M2M"] = (old + "\n" if old else "") + pc_comment

    # ── B. Which field is now active? ───────────────────────
    if trig == "summary" and cell_click:
        field_active = cell_click["data"]["Field Name"]
    else:
        field_active = vd_rows[0]["field_name"] if vd_rows else None

    # ── C. Filter detail grids & pick SQL ───────────────────
    if field_active:
        vd_filt = vd_all[vd_all["field_name"] == field_active]
        pc_filt = pc_all[pc_all["field_name"] == field_active]
        vd_sql = (vd_filt["value_sql_logic"].iloc[0]
                  if "value_sql_logic" in vd_filt else "")
        pc_sql = (pc_filt["value_sql_logic"].iloc[0]
                  if "value_sql_logic" in pc_filt else "")
    else:
        vd_filt, pc_filt = vd_all, pc_all
        vd_sql = pc_sql = ""

    # ── D. Return all outputs ───────────────────────────────
    return (vd_filt.to_dict("records"),
            pc_filt.to_dict("records"),
            vd_sql, pc_sql,
            s_df.to_dict("records"))

# ────────────────────────────────────────────────────────────────
# 9.  Run
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)