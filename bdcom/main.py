import datetime as dt, re, pandas as pd
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output, State
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
# 4.  Summary dataframe
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
def wide(df_src):
    df_src = df_src[df_src.filemonth_dt.isin(MONTHS)].copy()
    denom = (df_src.groupby(["field_name", "filemonth_dt"], as_index=False)
             .value_records.sum()
             .rename(columns={"value_records": "_tot"}))
    merged = df_src.merge(denom, on=["field_name", "filemonth_dt"])
    base = merged[["field_name", "value_label"]].drop_duplicates() \
        .sort_values(["field_name", "value_label"]).reset_index(drop=True)
    for m in MONTHS:
        mm = merged[merged.filemonth_dt == m][["field_name", "value_label", "value_records"]]
        base = (base.merge(mm, on=["field_name", "value_label"], how="left")
                .rename(columns={"value_records": f"{fmt(m)}"}))
    num_cols = [c for c in base.columns if c not in ("field_name", "value_label")]
    base[num_cols] = base[num_cols].fillna(0)
    return base

vd_wide = wide(df_data[df_data.analysis_type == "value_dist"])
pc_src = df_data[(df_data.analysis_type == "pop_comp") & (df_data.value_label.apply(_contains))]
pc_wide = wide(pc_src)

def add_total_row(df):
    total = {"field_name": "Total", "value_label": "Sum"}
    for col in df.columns:
        if col not in ("field_name", "value_label"):
            total[col] = df[col].sum()
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)

def sql_for(fld, analysis):
    sub = df_data[(df_data.analysis_type == analysis) & (df_data.field_name == fld) &
                  (df_data.value_sql_logic.notna())]
    if sub.empty:
        return ""
    return sub.value_sql_logic.iloc[0].replace("\\n", "\n").replace("\\t", "\t")

# ────────────────────────────────────────────────────────────────
# 6.  Dash App
# ────────────────────────────────────────────────────────────────
app = Dash(__name__)
app.layout = html.Div([
    html.H2("BDCOMM FRY14M Field Analysis — 13-Month View"),
    dcc.Tabs([
        dcc.Tab(label="Summary", children=[
            dash_table.DataTable(
                id='summary-table',
                columns=[{"name": c, "id": c} for c in df_summary.columns],
                data=df_summary.to_dict('records'),
                filter_action='native',
                sort_action='native',
                row_selectable='single',
                selected_rows=[],
                page_size=20,
                style_table={'overflowX': 'auto'},
            )
        ]),
        dcc.Tab(label="Value Distribution", children=[
            html.Div(["Selected Value Label:", dcc.Input(id='vd_val_lbl', type='text', readOnly=True)]),
            dash_table.DataTable(
                id='vd-table',
                columns=[{"name": c, "id": c} for c in vd_wide.columns],
                data=vd_wide.to_dict('records'),
                filter_action='native',
                sort_action='native',
                row_selectable='single',
                selected_rows=[],
                page_size=20,
                style_table={'overflowX': 'auto'},
            ),
            dcc.Textarea(id='vd_comm_text', placeholder='Add comment…', style={'width': '100%', 'height': '60px'}),
            html.Button('Submit', id='vd_comm_btn', n_clicks=0),
            html.Pre(id='vd_sql', style={'whiteSpace': 'pre-wrap', 'backgroundColor': '#f3f3f3',
                                         'padding': '0.75rem', 'border': '1px solid #ddd', 'fontFamily': 'monospace',
                                         'fontSize': '0.85rem'})
        ]),
        dcc.Tab(label="Population Comparison", children=[
            html.Div(["Selected Value Label:", dcc.Input(id='pc_val_lbl', type='text', readOnly=True)]),
            dash_table.DataTable(
                id='pc-table',
                columns=[{"name": c, "id": c} for c in pc_wide.columns],
                data=pc_wide.to_dict('records'),
                filter_action='native',
                sort_action='native',
                row_selectable='single',
                selected_rows=[],
                page_size=20,
                style_table={'overflowX': 'auto'},
            ),
            dcc.Textarea(id='pc_comm_text', placeholder='Add comment…', style={'width': '100%', 'height': '60px'}),
            html.Button('Submit', id='pc_comm_btn', n_clicks=0),
            html.Pre(id='pc_sql', style={'whiteSpace': 'pre-wrap', 'backgroundColor': '#f3f3f3',
                                         'padding': '0.75rem', 'border': '1px solid #ddd', 'fontFamily': 'monospace',
                                         'fontSize': '0.85rem'})
        ])
    ])
])

# ────────────────────────────────────────────────────────────────
# 7.  Callbacks
# ────────────────────────────────────────────────────────────────
@app.callback(
    Output('vd-table', 'data'),
    Output('pc-table', 'data'),
    Output('vd_sql', 'children'),
    Output('pc_sql', 'children'),
    Input('summary-table', 'selected_rows'),
    State('summary-table', 'data')
)
def update_detail(selected_rows, summary_rows):
    if selected_rows:
        fld = summary_rows[selected_rows[0]]['Field Name']
        vd_df = vd_wide[vd_wide['field_name'] == fld]
        pc_df = pc_wide[pc_wide['field_name'] == fld]
        vd_sql = sql_for(fld, 'value_dist')
        pc_sql = sql_for(fld, 'pop_comp')
    else:
        vd_df = vd_wide
        pc_df = pc_wide
        vd_sql = ''
        pc_sql = ''
    return add_total_row(vd_df).to_dict('records'), add_total_row(pc_df).to_dict('records'), vd_sql, pc_sql

@app.callback(
    Output('summary-table', 'data'),
    Input('vd_comm_btn', 'n_clicks'),
    Input('pc_comm_btn', 'n_clicks'),
    State('vd-table', 'selected_rows'),
    State('vd-table', 'data'),
    State('vd_comm_text', 'value'),
    State('pc-table', 'selected_rows'),
    State('pc-table', 'data'),
    State('pc_comm_text', 'value'),
    State('summary-table', 'data'),
    prevent_initial_call=True
)
def update_comments(n_vd, n_pc, vd_sel, vd_data, vd_txt, pc_sel, pc_data, pc_txt, summary_data):
    df_sum = pd.DataFrame(summary_data)
    ctx = dash.callback_context
    if not ctx.triggered:
        return summary_data
    trig = ctx.triggered[0]['prop_id'].split('.')[0]
    if trig == 'vd_comm_btn' and vd_sel and vd_txt:
        idx = vd_sel[0]
        fld = vd_data[idx]['field_name']
        val_lbl = vd_data[idx]['value_label']
        new_entry = f"{val_lbl} - {vd_txt}"
        mask = df_sum['Field Name'] == fld
        old = df_sum.loc[mask, 'Comment Missing'].iloc[0]
        df_sum.loc[mask, 'Comment Missing'] = (old + '\n' if old else '') + new_entry
    if trig == 'pc_comm_btn' and pc_sel and pc_txt:
        idx = pc_sel[0]
        fld = pc_data[idx]['field_name']
        val_lbl = pc_data[idx]['value_label']
        new_entry = f"{val_lbl} - {pc_txt}"
        mask = df_sum['Field Name'] == fld
        old = df_sum.loc[mask, 'Comment M2M'].iloc[0]
        df_sum.loc[mask, 'Comment M2M'] = (old + '\n' if old else '') + new_entry
    return df_sum.to_dict('records')

# ────────────────────────────────────────────────────────────────
# 8.  Run
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)
