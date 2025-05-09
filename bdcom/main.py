import datetime as dt, re, pandas as pd
from dash import Dash, dcc, html
from dash.dependencies import Input, Output, State
import dash_ag_grid as dag
import dash  # for callback_context

# ────────────────────────────────────────────────────────────────
# 1.  Load workbook
# ────────────────────────────────────────────────────────────────
INPUT_FILE = "input.xlsx"
df_data = pd.read_excel(INPUT_FILE, sheet_name="Data")
df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"], format="%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 2.  13-month window (DESC)
# ────────────────────────────────────────────────────────────────
DATE1 = dt.datetime(2025, 1, 1)
MONTHS = pd.date_range(end=DATE1, periods=13, freq="MS")[::-1]   # DESC
fmt = lambda d: d.strftime("%b-%Y")  # Jan-2025 …

# ────────────────────────────────────────────────────────────────
# 3.  Helper for pop-comp phrases
# ────────────────────────────────────────────────────────────────
_PHRASES = [r"1\)\s*CF Loan - Both Pop, Diff Values",
            r"2\)\s*CF Loan - Prior Null, Current Pop",
            r"3\)\s*CF Loan - Prior Pop, Current Null"]
_contains = lambda x: any(re.search(p, str(x)) for p in _PHRASES)

# ────────────────────────────────────────────────────────────────
# 4.  Summary dataframe (unchanged)
# ────────────────────────────────────────────────────────────────
prev_month = MONTHS[1]  # Dec-24
rows = []
for fld in sorted(df_data["field_name"].unique()):
    miss1 = df_data[(df_data.analysis_type == "value_dist") & (df_data.field_name == fld) &
                    (df_data.filemonth_dt == DATE1) &
                    (df_data.value_label.str.contains("Missing", case=False, na=False))]["value_records"].sum()
    miss2 = df_data[(df_data.analysis_type == "value_dist") & (df_data.field_name == fld) &
                    (df_data.filemonth_dt == prev_month) &
                    (df_data.value_label.str.contains("Missing", case=False, na=False))]["value_records"].sum()
    diff1 = df_data[(df_data.analysis_type == "pop_comp") & (df_data.field_name == fld) &
                    (df_data.filemonth_dt == DATE1) &
                    (df_data.value_label.apply(_contains))]["value_records"].sum()
    diff2 = df_data[(df_data.analysis_type == "pop_comp") & (df_data.field_name == fld) &
                    (df_data.filemonth_dt == prev_month) &
                    (df_data.value_label.apply(_contains))]["value_records"].sum()
    rows.append([fld, miss1, miss2, "", diff1, diff2, ""])

df_summary = pd.DataFrame(rows, columns=[
    "Field Name",
    f"Missing {DATE1:%m/%d/%Y}",
    f"Missing {prev_month:%m/%d/%Y}",
    "Comment Missing",
    f"M2M Diff {DATE1:%m/%d/%Y}",
    f"M2M Diff {prev_month:%m/%d/%Y}",
    "Comment M2M",
])

# ────────────────────────────────────────────────────────────────
# 5.  Build wide frames + totals for each field
# ────────────────────────────────────────────────────────────────
def wide(df_src, include_percentage=True):
    """Return (wide_df, total_row_dict), optionally excluding percentage columns."""
    df_src = df_src[df_src.filemonth_dt.isin(MONTHS)].copy()

    # Denominator per field & month for the sum and percentage
    denom = (df_src.groupby(["field_name", "filemonth_dt"], as_index=False)
             .value_records.sum()
             .rename(columns={"value_records": "_tot"}))

    # Merge to get percentages
    merged = df_src.merge(denom, on=["field_name", "filemonth_dt"])
    merged["_%"] = merged["value_records"] / merged["_tot"]

    # Start with a base frame with unique field_name and value_label
    base = merged[["field_name", "value_label"]].drop_duplicates() \
        .sort_values(["field_name", "value_label"]).reset_index(drop=True)

    for m in MONTHS:  # Loop over months in descending order
        mm = merged[merged.filemonth_dt == m][["field_name", "value_label", "value_records", "_%"]]
        base = (base.merge(mm, on=["field_name", "value_label"], how="left")
                .rename(columns={"value_records": f"{fmt(m)}", "_%": f"{fmt(m)} %"}))

    # Remove percentage columns if not needed
    if not include_percentage:
        base = base.drop(columns=[col for col in base.columns if col.endswith(" %")])

    # Fill NaNs with 0s
    num_cols = [c for c in base.columns if c not in ("field_name", "value_label")]
    base[num_cols] = base[num_cols].fillna(0)

    return base


vd_wide = wide(df_data[df_data.analysis_type == "value_dist"], include_percentage=False)
pc_src = df_data[(df_data.analysis_type == "pop_comp") & (df_data.value_label.apply(_contains))]
pc_wide = wide(pc_src, include_percentage=False)

# ────────────────────────────────────────────────────────────────
# 6.  Column defs
# ────────────────────────────────────────────────────────────────
def col_defs(df):
    out = []
    for c in df.columns:
        d = {"headerName": c, "field": c, "filter": "agSetColumnFilter",
             "sortable": True, "resizable": True, "minWidth": 90}
        if c.endswith(" %"):
            d["valueFormatter"] = {"function": "d3.format('.1%')(params.value)"}
        out.append(d)
    return out

# ────────────────────────────────────────────────────────────────
# 7.  Helpers for comment UI
# ────────────────────────────────────────────────────────────────
def comment_block(label_id, text_id, btn_id):
    """Value-label display + textarea + button."""
    return html.Div([
        dcc.Input(id=label_id, type="text", readOnly=True,
                  placeholder="value_label (select a row)",
                  style={"width": "100%", "marginBottom": "0.25rem"}),
        dcc.Textarea(id=text_id, placeholder="Add comment…",
                     style={"width": "100%", "height": "60px"}),
        html.Button("Submit", id=btn_id, n_clicks=0,
                    style={"marginTop": "0.25rem"})
    ], style={"marginTop": "0.5rem"})


# grid options
GRID_SUMMARY = {"pagination": False, "rowSelection": "single", "domLayout": "normal"}
GRID_DETAIL = {"pagination": True, "paginationPageSize": 20,
               "rowSelection": "single", "domLayout": "normal"}

# ────────────────────────────────────────────────────────────────
# 8.  Layout
# ────────────────────────────────────────────────────────────────
app = Dash(__name__)
app.layout = html.Div([
    html.H2("BDCOMM FRY14M Field Analysis — 13-Month View"),
    dcc.Tabs([
        # -------- Summary --------------------------------------------------
        dcc.Tab(label="Summary", children=[
            dag.AgGrid(id="summary", columnDefs=col_defs(df_summary),
                       rowData=df_summary.to_dict("records"),
                       className="ag-theme-alpine", columnSize="sizeToFit",
                       dashGridOptions=GRID_SUMMARY,
                       style={"height": "600px", "width": "100%"})
        ]),
        # -------- Value Distribution --------------------------------------
        dcc.Tab(label="Value Distribution", children=[
            dag.AgGrid(id="vd", columnDefs=col_defs(vd_wide),
                       rowData=vd_wide.to_dict("records"),
                       className="ag-theme-alpine", columnSize="sizeToFit",
                       dashGridOptions={**GRID_DETAIL},
                       style={"height": "500px", "width": "100%"}),
            comment_block("vd_val_lbl", "vd_comm_text", "vd_comm_btn"),
            html.Pre(id="vd_sql", style={"whiteSpace": "pre-wrap",
                                         "backgroundColor": "#f3f3f3",
                                         "padding": "0.75rem", "border": "1px solid #ddd",
                                         "marginTop": "0.5rem", "fontFamily": "monospace",
                                         "fontSize": "0.85rem"}),
            dcc.Clipboard(target_id="vd_sql", title="Copy SQL Logic", style={"marginTop": "0.5rem"})
        ]),
        # -------- Population Comparison -----------------------------------
        dcc.Tab(label="Population Comparison", children=[
            dag.AgGrid(id="pc", columnDefs=col_defs(pc_wide),
                       rowData=pc_wide.to_dict("records"),
                       className="ag-theme-alpine", columnSize="sizeToFit",
                       dashGridOptions={**GRID_DETAIL},
                       style={"height": "500px", "width": "100%"}),
            comment_block("pc_val_lbl", "pc_comm_text", "pc_comm_btn"),
            html.Pre(id="pc_sql", style={"whiteSpace": "pre-wrap",
                                         "backgroundColor": "#f3f3f3",
                                         "padding": "0.75rem", "border": "1px solid #ddd",
                                         "marginTop": "0.5rem", "fontFamily": "monospace",
                                         "fontSize": "0.85rem"}),
            dcc.Clipboard(target_id="pc_sql", title="Copy SQL Logic", style={"marginTop": "0.5rem"})
        ]),
    ])
])

# ────────────────────────────────────────────────────────────────
# 9-A.  Small callbacks: populate value-label boxes
# ────────────────────────────────────────────────────────────────
@app.callback(Output("vd_val_lbl", "value"),
              Input("vd", "cellClicked"), State("vd", "rowData"))
def vd_label(evt, rows):
    if evt and "rowIndex" in evt:
        return rows[evt["rowIndex"]]["value_label"]
    return ""


@app.callback(Output("pc_val_lbl", "value"),
              Input("pc", "cellClicked"), State("pc", "rowData"))
def pc_label(evt, rows):
    if evt and "rowIndex" in evt:
        return rows[evt["rowIndex"]]["value_label"]
    return ""


# ────────────────────────────────────────────────────────────────
# 9-B.  Master callback (comments, filtering, SQL logic)
# ────────────────────────────────────────────────────────────────
@app.callback(
    Output("vd", "rowData"), Output("pc", "rowData"),
    Output("vd_sql", "children"), Output("pc_sql", "children"),
    Output("summary", "rowData"),
    Input("summary", "cellClicked"),
    Input("vd_comm_btn", "n_clicks"), Input("pc_comm_btn", "n_clicks"),
    State("summary", "rowData"),
    State("vd", "rowData"), State("pc", "rowData"),
    State("vd_comm_text", "value"), State("pc_comm_text", "value"),
    State("vd_val_lbl", "value"), State("pc_val_lbl", "value"),
    prevent_initial_call=True
)
def master(evt, n_vd, n_pc,
           s_rows, vd_rows, pc_rows,
           vd_txt, pc_txt, vd_lbl, pc_lbl):

    s_df = pd.DataFrame(s_rows)
    trig = dash.callback_context.triggered[0]["prop_id"].split(".")[0]

    # ---- append comments --------------------------------------------------
    if trig == "vd_comm_btn" and vd_txt and vd_lbl:
        fld = vd_rows[0]["field_name"] if vd_rows else None
        if fld:
            m = s_df["Field Name"] == fld
            new_entry = f"{vd_lbl} - {vd_txt}"
            old = s_df.loc[m, "Comment Missing"].iloc[0]
            s_df.loc[m, "Comment Missing"] = (old + "\n" if old else "") + new_entry

    if trig == "pc_comm_btn" and pc_txt and pc_lbl:
        fld = pc_rows[0]["field_name"] if pc_rows else None
        if fld:
            m = s_df["Field Name"] == fld
            new_entry = f"{pc_lbl} - {pc_txt}"
            old = s_df.loc[m, "Comment M2M"].iloc[0]
            s_df.loc[m, "Comment M2M"] = (old + "\n" if old else "") + new_entry

    # ---- field active (for filtering & SQL) ------------------------------
    fld_active = None
    if trig == "summary" and evt and "rowIndex" in evt:
        fld_active = s_df.iloc[evt["rowIndex"]]["Field Name"]
    elif vd_rows:
        fld_active = vd_rows[0]["field_name"]

    if fld_active:
        vd_filtered = vd_wide[vd_wide["field_name"] == fld_active]
        pc_filtered = pc_wide[pc_wide["field_name"] == fld_active]

        # Compute totals only for the selected field
        vd_total = vd_filtered.sum(numeric_only=True)
        pc_total = pc_filtered.sum(numeric_only=True)

        # Add a "Total" row
        vd_total_row = {"field_name": "Total", "value_label": ""}
        for col in vd_filtered.columns:
            if col.endswith(" Sum"):
                vd_total_row[col] = vd_total[col] if col in vd_total else 0
            elif col.endswith(" %"):
                vd_total_row[col] = vd_filtered[col].sum() if col in vd_filtered else 0

        pc_total_row = {"field_name": "Total", "value_label": ""}
        for col in pc_filtered.columns:
            if col.endswith(" Sum"):
                pc_total_row[col] = pc_total[col] if col in pc_total else 0
            elif col.endswith(" %"):
                pc_total_row[col] = pc_filtered[col].sum() if col in pc_filtered else 0

    else:
        vd_filtered, pc_filtered = vd_wide, pc_wide
        vd_total_row = {}
        pc_total_row = {}

    # ---- SQL logic -------------------------------------------------------
    def sql_for(fld, analysis):
        if fld is None:
            return ""
        sub = df_data[(df_data.analysis_type == analysis) &
                      (df_data.field_name == fld) &
                      (df_data.value_sql_logic.notna())]
        if sub.empty:
            return ""
        return (sub.value_sql_logic.iloc[0]
                .replace("\\n", "\n").replace("\\t", "\t").replace("\\r", "\r"))

    vd_sql = sql_for(fld_active, "value_dist")
    pc_sql = sql_for(fld_active, "pop_comp")

    return (vd_filtered.to_dict("records"), pc_filtered.to_dict("records"),
            vd_sql, pc_sql, s_df.to_dict("records"))

# ────────────────────────────────────────────────────────────────
# 10.  Run
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)