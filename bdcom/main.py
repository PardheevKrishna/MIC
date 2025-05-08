"""
dash_ag_grid_code.py
====================
AG Grid version — now without the floating filter row.

Python ≥3.8 · Dash ≥2.17 · dash-ag-grid ≥31 · Pandas ≥1.3
"""

import datetime as dt
import re
import pandas as pd
from dash import Dash, dcc, html
from dash.dependencies import Input, Output, State
import dash_ag_grid as dag
import dash   # for callback_context

# ────────────────────────────────────────────────────────────────
# 1.  Load workbook
# ────────────────────────────────────────────────────────────────
INPUT_FILE = "input.xlsx"
df_data = pd.read_excel(INPUT_FILE, sheet_name="Data")
df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"], format="%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 2.  Months to compare
# ────────────────────────────────────────────────────────────────
DATE1 = dt.datetime(2025, 1, 1)
DATE2 = dt.datetime(2024, 12, 1)

# ────────────────────────────────────────────────────────────────
# 3.  Helper for pop-comp rows
# ────────────────────────────────────────────────────────────────
_PHRASES = [
    r"1\)\s*CF Loan - Both Pop, Diff Values",
    r"2\)\s*CF Loan - Prior Null, Current Pop",
    r"3\)\s*CF Loan - Prior Pop, Current Null",
]
_contains = lambda x: any(re.search(p, str(x)) for p in _PHRASES)

# ────────────────────────────────────────────────────────────────
# 4.  Build Summary dataframe (adds comment cols)
# ────────────────────────────────────────────────────────────────
rows = []
for fld in sorted(df_data["field_name"].unique()):
    miss1 = df_data[(df_data["analysis_type"]=="value_dist")&(df_data["field_name"]==fld)&
                    (df_data["filemonth_dt"]==DATE1)&
                    (df_data["value_label"].str.contains("Missing",case=False,na=False))
                    ]["value_records"].sum()
    miss2 = df_data[(df_data["analysis_type"]=="value_dist")&(df_data["field_name"]==fld)&
                    (df_data["filemonth_dt"]==DATE2)&
                    (df_data["value_label"].str.contains("Missing",case=False,na=False))
                    ]["value_records"].sum()
    diff1 = df_data[(df_data["analysis_type"]=="pop_comp")&(df_data["field_name"]==fld)&
                    (df_data["filemonth_dt"]==DATE1)&
                    (df_data["value_label"].apply(_contains))
                    ]["value_records"].sum()
    diff2 = df_data[(df_data["analysis_type"]=="pop_comp")&(df_data["field_name"]==fld)&
                    (df_data["filemonth_dt"]==DATE2)&
                    (df_data["value_label"].apply(_contains))
                    ]["value_records"].sum()
    rows.append([fld, miss1, miss2, "", diff1, diff2, ""])

df_summary = pd.DataFrame(rows, columns=[
    "Field Name",
    f"Missing {DATE1:%m/%d/%Y}",
    f"Missing {DATE2:%m/%d/%Y}",
    "Comment Missing",
    f"M2M Diff {DATE1:%m/%d/%Y}",
    f"M2M Diff {DATE2:%m/%d/%Y}",
    "Comment M2M",
])

# ────────────────────────────────────────────────────────────────
# 5.  Detail frames (two months only)
# ────────────────────────────────────────────────────────────────
mask_2m = df_data["filemonth_dt"].isin([DATE1, DATE2])

vd_all = df_data[mask_2m & (df_data["analysis_type"]=="value_dist")].copy()
pc_all = df_data[mask_2m & (df_data["analysis_type"]=="pop_comp")
                 ].loc[lambda d: d["value_label"].apply(_contains)].copy()

vd_all["filemonth_dt"] = vd_all["filemonth_dt"].dt.strftime("%m/%d/%Y")
pc_all["filemonth_dt"] = pc_all["filemonth_dt"].dt.strftime("%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 6.  Column-def factory  (no floatingFilter)
# ────────────────────────────────────────────────────────────────
def make_col_defs(df, hide_sql=False):
    defs = []
    for c in df.columns:
        if hide_sql and c == "value_sql_logic":
            continue
        defs.append({
            "headerName": c,
            "field":      c,
            "filter":     "agSetColumnFilter",   # checkbox filter
            "sortable":   True,
            "editable":   c.startswith("Comment"),
            "resizable":  True,
        })
    return defs

# ────────────────────────────────────────────────────────────────
# 7.  Comment area helper
# ────────────────────────────────────────────────────────────────
def comment_box(text_id, btn_id):
    return html.Div([
        dcc.Textarea(id=text_id, placeholder="Add comment…",
                     style={"width":"100%","height":"60px"}),
        html.Button("Submit", id=btn_id, n_clicks=0,
                    style={"marginTop":"0.25rem"})
    ], style={"marginTop":"0.5rem"})

GRID_OPTS = {"pagination":True, "paginationPageSize":20,
             "rowSelection":"single", "domLayout":"normal"}

# ────────────────────────────────────────────────────────────────
# 8.  Layout
# ────────────────────────────────────────────────────────────────
app = Dash(__name__)
app.layout = html.Div([
    html.H2("BDCOMM FRY14M Field Analysis — AG Grid"),
    dcc.Tabs([
        dcc.Tab(label="Summary", children=[
            dag.AgGrid(
                id="summary",
                columnDefs=make_col_defs(df_summary),
                rowData=df_summary.to_dict("records"),
                className="ag-theme-alpine",
                dashGridOptions=GRID_OPTS,
            )
        ]),
        dcc.Tab(label="Value Distribution", children=[
            dag.AgGrid(
                id="vd",
                columnDefs=make_col_defs(vd_all, hide_sql=True),
                rowData=vd_all.to_dict("records"),
                className="ag-theme-alpine",
                dashGridOptions=GRID_OPTS,
            ),
            comment_box("vd_comm_text", "vd_comm_btn"),
            html.Pre(id="vd_sql", style={"whiteSpace":"pre-wrap",
                                         "backgroundColor":"#f3f3f3",
                                         "padding":"0.75rem",
                                         "border":"1px solid #ddd",
                                         "marginTop":"0.5rem",
                                         "fontFamily":"monospace",
                                         "fontSize":"0.9rem"})
        ]),
        dcc.Tab(label="Population Comparison", children=[
            dag.AgGrid(
                id="pc",
                columnDefs=make_col_defs(pc_all, hide_sql=True),
                rowData=pc_all.to_dict("records"),
                className="ag-theme-alpine",
                dashGridOptions=GRID_OPTS,
            ),
            comment_box("pc_comm_text", "pc_comm_btn"),
            html.Pre(id="pc_sql", style={"whiteSpace":"pre-wrap",
                                         "backgroundColor":"#f3f3f3",
                                         "padding":"0.75rem",
                                         "border":"1px solid #ddd",
                                         "marginTop":"0.5rem",
                                         "fontFamily":"monospace",
                                         "fontSize":"0.9rem"})
        ]),
    ])
])

# ────────────────────────────────────────────────────────────────
# 9.  Callback (unchanged except for variable names)
# ────────────────────────────────────────────────────────────────
@app.callback(
    Output("vd","rowData"), Output("pc","rowData"),
    Output("vd_sql","children"), Output("pc_sql","children"),
    Output("summary","rowData"),
    Input("summary","cellClicked"),
    Input("vd_comm_btn","n_clicks"), Input("pc_comm_btn","n_clicks"),
    State("summary","rowData"), State("vd","rowData"), State("pc","rowData"),
    State("vd_comm_text","value"), State("pc_comm_text","value"),
    prevent_initial_call=True
)
def master(click_evt, n_vd, n_pc,
           s_rows, vd_rows, pc_rows, vd_txt, pc_txt):

    s_df = pd.DataFrame(s_rows)
    trig = dash.callback_context.triggered[0]["prop_id"].split(".")[0]

    # append comments
    if trig=="vd_comm_btn" and vd_txt:
        fld = vd_rows[0]["field_name"] if vd_rows else None
        if fld:
            m = s_df["Field Name"]==fld
            s_df.loc[m,"Comment Missing"] = (
                (s_df.loc[m,"Comment Missing"].iloc[0]+"\n" if s_df.loc[m,"Comment Missing"].iloc[0] else "") + vd_txt
            )
    if trig=="pc_comm_btn" and pc_txt:
        fld = pc_rows[0]["field_name"] if pc_rows else None
        if fld:
            m = s_df["Field Name"]==fld
            s_df.loc[m,"Comment M2M"] = (
                (s_df.loc[m,"Comment M2M"].iloc[0]+"\n" if s_df.loc[m,"Comment M2M"].iloc[0] else "") + pc_txt
            )

    # active field
    if trig=="summary" and click_evt and "rowIndex" in click_evt:
        fld_active = s_df.iloc[click_evt["rowIndex"]]["Field Name"]
    else:
        fld_active = vd_rows[0]["field_name"] if vd_rows else None

    if fld_active:
        vd_f = vd_all[vd_all["field_name"]==fld_active]
        pc_f = pc_all[pc_all["field_name"]==fld_active]
        vd_sql = vd_f["value_sql_logic"].iloc[0] if "value_sql_logic" in vd_f else ""
        pc_sql = pc_f["value_sql_logic"].iloc[0] if "value_sql_logic" in pc_f else ""
    else:
        vd_f, pc_f, vd_sql, pc_sql = vd_all, pc_all, "", ""

    return (vd_f.to_dict("records"), pc_f.to_dict("records"),
            vd_sql, pc_sql, s_df.to_dict("records"))

# ────────────────────────────────────────────────────────────────
# 10.  Run
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)