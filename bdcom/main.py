"""
dash_ag_grid_code.py
====================
• Summary grid unchanged (scrollable, 600 px tall).
• Value-Distribution & Population-Comparison grids show:
    field_name, value_label, and 13-month (Sum, %) pairs in
    descending order starting Jan-25.
  + Pinned “Total” row under all *Sum* columns.
  + Columns auto-sized to fit.

Python ≥3.8 · Dash ≥2.17 · dash-ag-grid ≥31 · Pandas ≥1.3
"""

import datetime as dt
import re
import pandas as pd
from dash import Dash, dcc, html
from dash.dependencies import Input, Output, State
import dash_ag_grid as dag
import dash                    # for callback_context

# ────────────────────────────────────────────────────────────────
# 1.  Load workbook
# ────────────────────────────────────────────────────────────────
INPUT_FILE = "input.xlsx"
df_data = pd.read_excel(INPUT_FILE, sheet_name="Data")
df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"],
                                         format="%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 2.  Time window (13 months ending DATE1, descending)
# ────────────────────────────────────────────────────────────────
DATE1  = dt.datetime(2025, 1, 1)
MONTHS = pd.date_range(end=DATE1, periods=13, freq="MS")[::-1]  # DESC
fmt    = lambda d: d.strftime("%b-%Y")                          # Jan-2025 …

# ────────────────────────────────────────────────────────────────
# 3.  Helper for pop-comp rows
# ────────────────────────────────────────────────────────────────
_PHRASES  = [r"1\)\s*CF Loan - Both Pop, Diff Values",
             r"2\)\s*CF Loan - Prior Null, Current Pop",
             r"3\)\s*CF Loan - Prior Pop, Current Null"]
_contains = lambda x: any(re.search(p, str(x)) for p in _PHRASES)

# ────────────────────────────────────────────────────────────────
# 4.  Summary dataframe
# ────────────────────────────────────────────────────────────────
prev_month = MONTHS[1]                              # Dec-24
summary_rows = []
for fld in sorted(df_data["field_name"].unique()):
    miss1 = df_data[(df_data.analysis_type=="value_dist") & (df_data.field_name==fld) &
                    (df_data.filemonth_dt==DATE1) &
                    (df_data.value_label.str.contains("Missing",case=False,na=False))
                   ]["value_records"].sum()
    miss2 = df_data[(df_data.analysis_type=="value_dist") & (df_data.field_name==fld) &
                    (df_data.filemonth_dt==prev_month) &
                    (df_data.value_label.str.contains("Missing",case=False,na=False))
                   ]["value_records"].sum()
    diff1 = df_data[(df_data.analysis_type=="pop_comp") & (df_data.field_name==fld) &
                    (df_data.filemonth_dt==DATE1) &
                    (df_data.value_label.apply(_contains))
                   ]["value_records"].sum()
    diff2 = df_data[(df_data.analysis_type=="pop_comp") & (df_data.field_name==fld) &
                    (df_data.filemonth_dt==prev_month) &
                    (df_data.value_label.apply(_contains))
                   ]["value_records"].sum()
    summary_rows.append([fld, miss1, miss2, "", diff1, diff2, ""])

df_summary = pd.DataFrame(summary_rows, columns=[
    "Field Name",
    f"Missing {DATE1:%m/%d/%Y}",
    f"Missing {prev_month:%m/%d/%Y}",
    "Comment Missing",
    f"M2M Diff {DATE1:%m/%d/%Y}",
    f"M2M Diff {prev_month:%m/%d/%Y}",
    "Comment M2M",
])

# ────────────────────────────────────────────────────────────────
# 5.  Build 13-month wide frames (descending) + total rows
# ────────────────────────────────────────────────────────────────
def build_wide(df_src: pd.DataFrame):
    """Return (wide_df, total_row_dict)."""
    df_src = df_src[df_src.filemonth_dt.isin(MONTHS)].copy()

    # Denominator per field & month
    denom = (df_src.groupby(["field_name","filemonth_dt"], as_index=False)
                    .value_records.sum()
                    .rename(columns={"value_records":"_total"}))

    merged       = df_src.merge(denom, on=["field_name","filemonth_dt"])
    merged["_%"] = merged["value_records"]/merged["_total"]

    wide = (merged[["field_name","value_label"]]
            .drop_duplicates()
            .sort_values(["field_name","value_label"])
            .reset_index(drop=True))

    for m in MONTHS:           # already DESC
        mm = merged[merged.filemonth_dt==m][["field_name","value_label",
                                             "value_records","_%"]]
        wide = (wide.merge(mm, on=["field_name","value_label"], how="left")
                    .rename(columns={"value_records":f"{fmt(m)} Sum",
                                     "_%":f"{fmt(m)} %"}))

    # NaN -> 0
    num_cols = [c for c in wide.columns if c not in ("field_name","value_label")]
    wide[num_cols] = wide[num_cols].fillna(0)

    # build a totals row (Sum columns only)
    total_row = {"field_name":"Total", "value_label":""}
    for c in num_cols:
        total_row[c] = wide[c].sum() if c.endswith(" Sum") else ""
    return wide, total_row

vd_wide, vd_total = build_wide(df_data[df_data.analysis_type=="value_dist"])
pc_src            = df_data[(df_data.analysis_type=="pop_comp") & (df_data.value_label.apply(_contains))]
pc_wide, pc_total = build_wide(pc_src)

# ────────────────────────────────────────────────────────────────
# 6.  Column definitions (auto size, % formatter)
# ────────────────────────────────────────────────────────────────
def make_col_defs(df):
    defs = []
    for c in df.columns:
        col = {"headerName": c, "field": c,
               "filter":"agSetColumnFilter", "sortable":True,
               "resizable":True, "minWidth":90}
        if c.endswith(" %"):
            col["valueFormatter"] = {"function":"d3.format('.1%')(params.value)"}
        defs.append(col)
    return defs

# ────────────────────────────────────────────────────────────────
# 7.  Comment box helper
# ────────────────────────────────────────────────────────────────
def comment_box(text_id, btn_id):
    return html.Div([
        dcc.Textarea(id=text_id, placeholder="Add comment…",
                     style={"width":"100%","height":"60px"}),
        html.Button("Submit", id=btn_id, n_clicks=0,
                    style={"marginTop":"0.25rem"})
    ], style={"marginTop":"0.5rem"})

# grid options
GRID_SUMMARY = {"pagination":False, "rowSelection":"single", "domLayout":"normal"}
GRID_DETAIL  = {"pagination":True, "paginationPageSize":20,
                "rowSelection":"single", "domLayout":"normal"}

# ────────────────────────────────────────────────────────────────
# 8.  Dash layout
# ────────────────────────────────────────────────────────────────
app = Dash(__name__)
app.layout = html.Div([
    html.H2("BDCOMM FRY14M Field Analysis — 13-Month View"),
    dcc.Tabs([
        dcc.Tab(label="Summary", children=[
            dag.AgGrid(
                id="summary",
                columnDefs=make_col_defs(df_summary),
                rowData=df_summary.to_dict("records"),
                className="ag-theme-alpine",
                columnSize="sizeToFit",
                dashGridOptions=GRID_SUMMARY,
                style={"height":"600px","width":"100%"}
            )
        ]),
        dcc.Tab(label="Value Distribution", children=[
            dag.AgGrid(
                id="vd",
                columnDefs=make_col_defs(vd_wide),
                rowData=vd_wide.to_dict("records"),
                columnSize="sizeToFit",
                className="ag-theme-alpine",
                dashGridOptions={**GRID_DETAIL,
                                 "pinnedBottomRowData":[vd_total]},
                style={"height":"500px","width":"100%"}
            ),
            comment_box("vd_comm_text","vd_comm_btn"),
            html.Pre(id="vd_sql")
        ]),
        dcc.Tab(label="Population Comparison", children=[
            dag.AgGrid(
                id="pc",
                columnDefs=make_col_defs(pc_wide),
                rowData=pc_wide.to_dict("records"),
                columnSize="sizeToFit",
                className="ag-theme-alpine",
                dashGridOptions={**GRID_DETAIL,
                                 "pinnedBottomRowData":[pc_total]},
                style={"height":"500px","width":"100%"}
            ),
            comment_box("pc_comm_text","pc_comm_btn"),
            html.Pre(id="pc_sql")
        ]),
    ])
])

# ────────────────────────────────────────────────────────────────
# 9.  Callback (comments & filtering)
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
def master(evt, n_vd, n_pc,
           s_rows, vd_rows, pc_rows, vd_txt, pc_txt):

    s_df  = pd.DataFrame(s_rows)
    trig  = dash.callback_context.triggered[0]["prop_id"].split(".")[0]

    # -- append comments ----------------------------------------------------
    if trig=="vd_comm_btn" and vd_txt and vd_rows:
        fld = vd_rows[0]["field_name"]
        m   = s_df["Field Name"]==fld
        old = s_df.loc[m,"Comment Missing"].iloc[0]
        s_df.loc[m,"Comment Missing"] = (old+"\n" if old else "") + vd_txt
    if trig=="pc_comm_btn" and pc_txt and pc_rows:
        fld = pc_rows[0]["field_name"]
        m   = s_df["Field Name"]==fld
        old = s_df.loc[m,"Comment M2M"].iloc[0]
        s_df.loc[m,"Comment M2M"] = (old+"\n" if old else "") + pc_txt

    # -- which field is active? --------------------------------------------
    fld_active = None
    if trig=="summary" and evt and "rowIndex" in evt:
        fld_active = s_df.iloc[evt["rowIndex"]]["Field Name"]
    elif vd_rows:
        fld_active = vd_rows[0]["field_name"]

    if fld_active:
        vd_filtered = vd_wide[vd_wide["field_name"]==fld_active]
        pc_filtered = pc_wide[pc_wide["field_name"]==fld_active]
    else:
        vd_filtered, pc_filtered = vd_wide, pc_wide

    return (vd_filtered.to_dict("records"), pc_filtered.to_dict("records"),
            "", "", s_df.to_dict("records"))

# ────────────────────────────────────────────────────────────────
# 10.  Run
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)