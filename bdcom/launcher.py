import datetime as dt
import re
import pandas as pd
import numpy as np

from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output, State

# ────────────────────────────────────────────────────────────────
# 0. Sample lists for gating selections
# ────────────────────────────────────────────────────────────────
PORTFOLIOS    = ["Portfolio A", "Portfolio B", "Portfolio C"]
EXCEL_REPORTS = ["FieldAnalysis_v1.xlsx", "FieldAnalysis_v2.xlsx"]

# ────────────────────────────────────────────────────────────────
# 1. Load the standardized input.xlsx
# ────────────────────────────────────────────────────────────────
INPUT_FILE = "input.xlsx"
df_data = pd.read_excel(INPUT_FILE, sheet_name="Data")
df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"], format="%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 2. Prepare the 13-month window
# ────────────────────────────────────────────────────────────────
DATE1 = dt.datetime(2025, 1, 1)
MONTHS = pd.date_range(end=DATE1, periods=13, freq="MS")[::-1]
fmt = lambda d: d.strftime("%b-%Y")
MONTH_OPTIONS = [fmt(m) for m in MONTHS]
prev_month = MONTHS[1]

# ────────────────────────────────────────────────────────────────
# 3. Helper for pop-comp flags
# ────────────────────────────────────────────────────────────────
_PHRASES = [
    r"1\)\s*CF Loan - Both Pop, Diff Values",
    r"2\)\s*CF Loan - Prior Null, Current Pop",
    r"3\)\s*CF Loan - Prior Pop, Current Null",
]
_contains = lambda x: any(re.search(p, str(x)) for p in _PHRASES)

# ────────────────────────────────────────────────────────────────
# 4. Build the Summary DataFrame
# ────────────────────────────────────────────────────────────────
rows = []
for fld in sorted(df_data["field_name"].unique()):
    m1 = df_data[
        (df_data.analysis_type=="value_dist") &
        (df_data.field_name==fld) &
        (df_data.filemonth_dt==DATE1) &
        (df_data.value_label.str.contains("Missing", case=False, na=False))
    ]["value_records"].sum()
    m2 = df_data[
        (df_data.analysis_type=="value_dist") &
        (df_data.field_name==fld) &
        (df_data.filemonth_dt==prev_month) &
        (df_data.value_label.str.contains("Missing", case=False, na=False))
    ]["value_records"].sum()
    d1 = df_data[
        (df_data.analysis_type=="pop_comp") &
        (df_data.field_name==fld) &
        (df_data.filemonth_dt==DATE1) &
        (df_data.value_label.apply(_contains))
    ]["value_records"].sum()
    d2 = df_data[
        (df_data.analysis_type=="pop_comp") &
        (df_data.field_name==fld) &
        (df_data.filemonth_dt==prev_month) &
        (df_data.value_label.apply(_contains))
    ]["value_records"].sum()
    rows.append([fld, m1, m2, "", d1, d2, ""])

initial_summary = pd.DataFrame(
    rows,
    columns=[
        "Field Name",
        f"Missing {DATE1:%m/%d/%Y}",
        f"Missing {prev_month:%m/%d/%Y}",
        "Comment Missing",
        f"M2M Diff {DATE1:%m/%d/%Y}",
        f"M2M Diff {prev_month:%m/%d/%Y}",
        "Comment M2M",
    ],
)
FIELD_NAMES = initial_summary["Field Name"].tolist()

# ────────────────────────────────────────────────────────────────
# 5. Wide-format tables for Value Distribution & Pop-Comp
# ────────────────────────────────────────────────────────────────
def wide(df_src):
    src = df_src[df_src.filemonth_dt.isin(MONTHS)].copy()
    denom = src.groupby(["field_name","filemonth_dt"], as_index=False).value_records.sum().rename(columns={"value_records":"_tot"})
    merged = src.merge(denom, on=["field_name","filemonth_dt"])
    base = merged[["field_name","value_label"]].drop_duplicates().sort_values(["field_name","value_label"]).reset_index(drop=True)
    for m in MONTHS:
        mm = merged[merged.filemonth_dt==m][["field_name","value_label","value_records"]]
        base = base.merge(mm, on=["field_name","value_label"], how="left").rename(columns={"value_records":fmt(m)})
    cols = [c for c in base.columns if c not in ("field_name","value_label")]
    base[cols] = base[cols].fillna(0)
    return base

vd_wide = wide(df_data[df_data.analysis_type=="value_dist"])
pc_wide = wide(df_data[(df_data.analysis_type=="pop_comp") & (df_data.value_label.apply(_contains))])

def add_total_row(df):
    total = {"field_name":"Total","value_label":"Sum"}
    for c in df.columns:
        if c not in ("field_name","value_label"):
            total[c] = df[c].sum()
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)

def sql_for(fld, analysis):
    sub = df_data[
        (df_data.analysis_type==analysis) &
        (df_data.field_name==fld) &
        (df_data.value_sql_logic.notna())
    ]
    return sub.value_sql_logic.iloc[0].replace("\\n","\n").replace("\\t","\t") if not sub.empty else ""

# ────────────────────────────────────────────────────────────────
# 6. Load & pivot user comments
# ────────────────────────────────────────────────────────────────
prev_c = pd.read_csv("prev_comments.csv", parse_dates=["date"])
prev_c["month"] = prev_c.date.dt.to_period("M").dt.to_timestamp()
prev_c = prev_c[prev_c.month.isin(MONTHS[1:])]
miss = prev_c[prev_c.research=="Missing"]
m2m  = prev_c[prev_c.research=="M2M Diff"]

def pivot_comments(df, prefix):
    pts = df.groupby(["field_name","month"]).comment.agg(lambda x:"\n".join(x)).reset_index()
    pts["col"] = pts.month.apply(fmt)
    w = pts.pivot(index="field_name",columns="col",values="comment")
    w.columns = [f"{prefix} {c}" for c in w.columns]
    return w.reset_index()

pivot_miss = pivot_comments(miss, "Prev Missing")
pivot_m2m   = pivot_comments(m2m,   "Prev M2M")
prev_wide = pd.merge(pivot_miss, pivot_m2m, on="field_name", how="outer").fillna("")
cur = initial_summary[["Field Name","Comment Missing","Comment M2M"]].rename(
    columns={"Comment Missing":"Comment Missing This Month","Comment M2M":"Comment M2M This Month"}
)
prev_summary = pd.merge(cur, prev_wide, left_on="Field Name", right_on="field_name", how="left").drop(columns=["field_name"])
prev_cols = [c for c in prev_summary.columns if c not in ["Comment Missing This Month","Comment M2M This Month"]]
prev_display = prev_summary[prev_cols]
style_cond_prev = [{"if":{"column_id":c},"width":f"{max(len(c)*8,100)}px"} for c in prev_cols]

# ────────────────────────────────────────────────────────────────
# 7. Simulated SAS history script & pct output
# ────────────────────────────────────────────────────────────────
sas_scripts = [{
    "filename":"bref_14M_final.sas",
    "code":(
        "/* field-level pct check */\n"
        "proc sql;\n"
        "  select field_name, sum(value)/sum(_tot)*100 as pct\n"
        "  from analysis\n"
        "  group by field_name;\n"
        "quit;"
    )
}]
sas_history_df = pd.DataFrame({
    "filename":"bref_14M_final.sas",
    "field_name":FIELD_NAMES,
    "pct":np.random.uniform(0,100,len(FIELD_NAMES))
})

# ────────────────────────────────────────────────────────────────
# 8. Build Dash app & layout
# ────────────────────────────────────────────────────────────────
app = Dash(__name__, external_stylesheets=["https://cdn.jsdelivr.net/npm/bootswatch@4.5.2/dist/flatly/bootstrap.min.css"])
app.layout = html.Div([
    html.H2("BDCOMM FRY14M Field Analysis", className="text-center mb-4"),

    # gating controls
    html.Div(className="row mb-4", children=[
        html.Div(className="col-md-4", children=[
            html.Label("Select Portfolio:"), dcc.Dropdown(id="select-portfolio", options=[{"label":p,"value":p} for p in PORTFOLIOS], placeholder="Portfolio...")
        ]),
        html.Div(className="col-md-4", children=[
            html.Label("Select Excel Report:"), dcc.Dropdown(id="select-report", options=[{"label":r,"value":r} for r in EXCEL_REPORTS], placeholder="Report...")
        ]),
        html.Div(className="col-md-4", children=[
            html.Label("Select Month:"), dcc.Dropdown(id="select-month", options=[{"label":m,"value":m} for m in MONTH_OPTIONS], placeholder="Month...")
        ]),
    ]),

    html.Div(id="dashboard-container", style={"display":"none"}, children=[
        dcc.Store(id="summary-store", data=initial_summary.to_dict("records")),

        dcc.Tabs(id="main-tabs", children=[

            # Summary Tab
            dcc.Tab(label="Summary", className="p-3", children=[
                html.Div(className="row mb-4", children=[
                    *[
                        html.Div(className="col-md-3", children=[
                            html.Label(label),
                            dcc.Dropdown(
                                id=did, options=[{"label":i,"value":i} for i in sorted(initial_summary[col].unique())],
                                multi=True, value=sorted(initial_summary[col].unique()),
                                className="form-control"
                            )
                        ])
                        for label, col, did in [
                            (f"Missing {DATE1:%b-%Y}", f"Missing {DATE1:%m/%d/%Y}", "filter-miss1"),
                            (f"Missing {prev_month:%b-%Y}", f"Missing {prev_month:%m/%d/%Y}", "filter-miss2"),
                            (f"M2M Diff {DATE1:%b-%Y}", f"M2M Diff {DATE1:%m/%d/%Y}",   "filter-m2m1"),
                            (f"M2M Diff {prev_month:%b-%Y}", f"M2M Diff {prev_month:%m/%d/%Y}", "filter-m2m2"),
                        ]
                    ]
                ]),
                dash_table.DataTable(
                    id="summary-table",
                    columns=[{"name":c,"id":c,"editable":c in ["Comment Missing","Comment M2M"]} for c in initial_summary.columns],
                    data=[], editable=True, row_selectable="single", page_size=20,
                    style_table={"overflowX":"auto"}, style_cell={"textAlign":"left"}
                )
            ]),

            # Value Distribution Tab
            dcc.Tab(label="Value Distribution", className="p-3", children=[
                dash_table.DataTable(
                    id="vd-table",
                    columns=[{"name":c,"id":c} for c in vd_wide.columns],
                    data=[], filter_action="native", sort_action="native",
                    row_selectable="single", page_size=20,
                    style_table={"overflowX":"auto"}, style_cell={"textAlign":"left"}
                ),
                html.Div(className="mt-3", children=[
                    html.Label("Selected Value Label:"),
                    dcc.Input(id="vd-val-lbl", readOnly=True, className="form-control mb-2"),
                    dcc.Textarea(id="vd_comm_text", placeholder="Enter comment…", style={"width":"100%","height":"80px"}, className="form-control"),
                    html.Button("Add Comment", id="vd_comm_btn", className="btn btn-primary btn-sm mt-2")
                ]),
                html.Div(className="mt-3", children=[
                    html.H5("Value SQL Logic:"), html.Div(id="vd_sql", style={"whiteSpace":"pre-wrap","border":"1px solid #ced4da","padding":"0.5rem","borderRadius":"0.25rem"}),
                    dcc.Clipboard(target_id="vd_sql", title="Copy SQL", style={"marginTop":"0.5rem"})
                ])
            ]),

            # Population Comparison Tab
            dcc.Tab(label="Population Comparison", className="p-3", children=[
                dash_table.DataTable(
                    id="pc-table",
                    columns=[{"name":c,"id":c} for c in pc_wide.columns],
                    data=[], filter_action="native", sort_action="native",
                    row_selectable="single", page_size=20,
                    style_table={"overflowX":"auto"}, style_cell={"textAlign":"left"}
                ),
                html.Div(className="mt-3", children=[
                    html.Label("Selected Value Label:"),
                    dcc.Input(id="pc-val-lbl", readOnly=True, className="form-control mb-2"),
                    dcc.Textarea(id="pc_comm_text", placeholder="Enter comment…", style={"width":"100%","height":"80px"}, className="form-control"),
                    html.Button("Add Comment", id="pc_comm_btn", className="btn btn-primary btn-sm mt-2")
                ]),
                html.Div(className="mt-3", children=[
                    html.H5("Population-Comp SQL Logic:"), html.Div(id="pc_sql", style={"whiteSpace":"pre-wrap","border":"1px solid #ced4da","padding":"0.5rem","borderRadius":"0.25rem"}),
                    dcc.Clipboard(target_id="pc_sql", title="Copy SQL", style={"marginTop":"0.5rem"})
                ])
            ]),

            # Comments Tab
            dcc.Tab(label="Comments", className="p-3", children=[
                html.Button("Show All Fields", id="prev_show_all_btn", className="btn btn-secondary btn-sm mb-3"),
                dash_table.DataTable(
                    id="prev-comments-table",
                    columns=[{"name":c,"id":c} for c in prev_cols],
                    data=prev_display.to_dict("records"),
                    filter_action="native", sort_action="native", page_size=20,
                    style_table={"overflowX":"auto"}, style_cell_conditional=style_cond_prev, style_cell={"whiteSpace":"normal","textAlign":"left"}
                )
            ]),

            # SAS History Tab
            dcc.Tab(label="SAS History", className="p-3", children=[
                html.H4("bref_14M_final.sas"),
                html.Pre(sas_scripts[0]["code"], style={"whiteSpace":"pre-wrap","border":"1px solid #ddd","padding":"10px"}),
                html.Div(className="row my-3", children=[
                    html.Div(className="col-md-6", children=[
                        html.Label("Select Fields:"), dcc.Dropdown(id="sas-history-fields", options=[{"label":f,"value":f} for f in FIELD_NAMES], multi=True)
                    ]),
                    html.Div(className="col-md-6", children=[
                        html.Label("Threshold (%):"), dcc.Input(id="sas-threshold", type="number", min=0, max=100, step=1, value=50, className="form-control")
                    ])
                ]),
                html.Div(id="sas-history-results")
            ]),

            # SAS Ad-hoc Execution Tab
            dcc.Tab(label="SAS Ad-hoc", className="p-3", children=[
                html.Label("Enter SAS Code:"), 
                dcc.Textarea(id="sas-code-input", style={"width":"100%","height":"300px"}, placeholder="Paste SAS code..."),
                html.Button("Run SAS", id="run-sas-btn", className="btn btn-primary mt-2"),
                html.Div(className="mt-4", children=[html.H5("Log Output:"), html.Pre(id="sas-log-output", style={"whiteSpace":"pre-wrap","border":"1px solid #ccc","padding":"10px"})]),
                html.Div(className="mt-4", children=[html.H5("Data Output:"), dash_table.DataTable(id="sas-data-output", columns=[], data=[], page_size=5, style_table={"overflowX":"auto"}, style_cell={"textAlign":"left"})])
            ]),

        ])
    ]),
], className="container-fluid p-4", style={"backgroundColor":"#f8f9fa"})

# ────────────────────────────────────────────────────────────────
# 9. Callbacks
# ────────────────────────────────────────────────────────────────
@app.callback(
    Output("dashboard-container","style"),
    Input("select-portfolio","value"),
    Input("select-report","value"),
    Input("select-month","value"),
)
def toggle_dashboard(p, r, m):
    return {"display":"block"} if p and r and m else {"display":"none"}

@app.callback(
    Output("summary-table","data"),
    Input("summary-store","data"),
    Input("filter-miss1","value"), Input("filter-miss2","value"),
    Input("filter-m2m1","value"), Input("filter-m2m2","value"),
)
def filter_summary(store, m1, m2, d1, d2):
    df = pd.DataFrame(store)
    df = df[df[f"Missing {DATE1:%m/%d/%Y}"].isin(m1)]
    df = df[df[f"Missing {prev_month:%m/%d/%Y}"].isin(m2)]
    df = df[df[f"M2M Diff {DATE1:%m/%d/%Y}"].isin(d1)]
    df = df[df[f"M2M Diff {prev_month:%m/%d/%Y}"].isin(d2)]
    return df.to_dict("records")

@app.callback(
    Output("summary-store","data"),
    Input("vd_comm_btn","n_clicks"), Input("pc_comm_btn","n_clicks"),
    State("vd-table","active_cell"), State("vd-table","data"), State("vd_comm_text","value"),
    State("pc-table","active_cell"), State("pc-table","data"), State("pc_comm_text","value"),
    State("summary-store","data"), prevent_initial_call=True
)
def update_comments(n_vd, n_pc, vd_act, vd_data, vd_txt, pc_act, pc_data, pc_txt, store):
    df_sum = pd.DataFrame(store)
    trig = dash.callback_context.triggered[0]["prop_id"].split(".")[0]
    if trig=="vd_comm_btn" and vd_act and vd_txt:
        r=vd_act["row"]; fld=vd_data[r]["field_name"]; lbl=vd_data[r]["value_label"]
        ent=f"{lbl} - {vd_txt}"; m=df_sum["Field Name"]==fld
        old=df_sum.loc[m,"Comment Missing"].iloc[0]
        df_sum.loc[m,"Comment Missing"]=(old+"\n" if old else "")+ent
    if trig=="pc_comm_btn" and pc_act and pc_txt:
        r=pc_act["row"]; fld=pc_data[r]["field_name"]; lbl=pc_data[r]["value_label"]
        ent=f"{lbl} - {pc_txt}"; m=df_sum["Field Name"]==fld
        old=df_sum.loc[m,"Comment M2M"].iloc[0]
        df_sum.loc[m,"Comment M2M"]=(old+"\n" if old else "")+ent
    return df_sum.to_dict("records")

@app.callback(
    Output("vd-table","data"), Output("pc-table","data"),
    Output("vd_sql","children"), Output("pc_sql","children"),
    Input("summary-table","selected_rows"), State("summary-table","data")
)
def update_detail(selected, rows):
    if selected:
        fld=rows[selected[0]]["Field Name"]
        vd_df=vd_wide[vd_wide.field_name==fld]
        pc_df=pc_wide[pc_wide.field_name==fld]
        vd_sql=sql_for(fld,"value_dist"); pc_sql=sql_for(fld,"pop_comp")
    else:
        vd_df,pc_df,vd_sql,pc_sql=vd_wide,pc_wide,"",""
    return add_total_row(vd_df).to_dict("records"), add_total_row(pc_df).to_dict("records"), vd_sql, pc_sql

@app.callback(
    Output("vd-val-lbl","value"), Input("vd-table","active_cell"), State("vd-table","data")
)
def update_vd_label(active, rows):
    return rows[active["row"]]["value_label"] if active else ""

@app.callback(
    Output("pc-val-lbl","value"), Input("pc-table","active_cell"), State("pc-table","data")
)
def update_pc_label(active, rows):
    return rows[active["row"]]["value_label"] if active else ""

@app.callback(
    Output("prev-comments-table","data"),
    Input("summary-table","selected_rows"), Input("prev_show_all_btn","n_clicks"),
    State("summary-table","data")
)
def update_prev(sel, show, rows):
    trig = dash.callback_context.triggered[0]["prop_id"].split(".")[0]
    if trig=="prev_show_all_btn":
        filt=prev_display
    elif sel:
        fld=rows[sel[0]]["Field Name"]; filt=prev_display[prev_display["Field Name"]==fld]
    else:
        filt=prev_display
    return filt.to_dict("records")

@app.callback(
    Output("sas-history-results","children"),
    Input("sas-history-fields","value"), Input("sas-threshold","value")
)
def filter_sas_history(fields, threshold):
    df = sas_history_df.copy()
    if fields:
        df = df[df.field_name.isin(fields)]
    df = df[df.pct >= (threshold or 0)]
    results = []
    for fld in df.field_name.unique():
        labels = pc_wide[pc_wide.field_name==fld]["value_label"].unique().tolist()
        sample = np.random.choice(labels, size=min(5,len(labels)), replace=False)
        out = pd.DataFrame({"value_label": sample, "pct": np.round(np.random.uniform(0,100,len(sample)),2)})
        results.append(html.H5(fld))
        results.append(
            dash_table.DataTable(
                columns=[{"name":c,"id":c} for c in out.columns],
                data=out.to_dict("records"),
                page_size=5,
                style_table={"overflowX":"auto"},
                style_cell={"textAlign":"left"}
            )
        )
    return results

@app.callback(
    Output("sas-log-output","children"),
    Output("sas-data-output","data"), Output("sas-data-output","columns"),
    Input("run-sas-btn","n_clicks"), State("sas-code-input","value"),
    prevent_initial_call=True
)
def run_sas_code(n, code):
    log = "NOTE: SAS code executed (static placeholder)."
    sample_fields = np.random.choice(FIELD_NAMES, size=min(5,len(FIELD_NAMES)), replace=False)
    df = pd.DataFrame({"field_name": sample_fields, "pct": np.round(np.random.uniform(0,100,len(sample_fields)),2)})
    return log, df.to_dict("records"), [{"name":c,"id":c} for c in df.columns]

if __name__ == "__main__":
    app.run(debug=True)