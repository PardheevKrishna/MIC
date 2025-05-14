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
MONTHS = pd.date_range(end=DATE1, periods=13, freq="MS")[::-1]
fmt = lambda d: d.strftime("%b-%Y")
prev_month = MONTHS[1]

# ────────────────────────────────────────────────────────────────
# 3.  Helper for pop-comp phrases
# ────────────────────────────────────────────────────────────────
_PHRASES = [
    r"1\)\s*CF Loan - Both Pop, Diff Values",
    r"2\)\s*CF Loan - Prior Null, Current Pop",
    r"3\)\s*CF Loan - Prior Pop, Current Null",
]
_contains = lambda x: any(re.search(p, str(x)) for p in _PHRASES)

# ────────────────────────────────────────────────────────────────
# 4.  Summary dataframe (current month)
# ────────────────────────────────────────────────────────────────
rows = []
for fld in sorted(df_data["field_name"].unique()):
    miss1 = df_data[
        (df_data.analysis_type == "value_dist")
        & (df_data.field_name == fld)
        & (df_data.filemonth_dt == DATE1)
        & (df_data.value_label.str.contains("Missing", case=False, na=False))
    ]["value_records"].sum()
    miss2 = df_data[
        (df_data.analysis_type == "value_dist")
        & (df_data.field_name == fld)
        & (df_data.filemonth_dt == prev_month)
        & (df_data.value_label.str.contains("Missing", case=False, na=False))
    ]["value_records"].sum()
    diff1 = df_data[
        (df_data.analysis_type == "pop_comp")
        & (df_data.field_name == fld)
        & (df_data.filemonth_dt == DATE1)
        & (df_data.value_label.apply(_contains))
    ]["value_records"].sum()
    diff2 = df_data[
        (df_data.analysis_type == "pop_comp")
        & (df_data.field_name == fld)
        & (df_data.filemonth_dt == prev_month)
        & (df_data.value_label.apply(_contains))
    ]["value_records"].sum()
    rows.append([fld, miss1, miss2, "", diff1, diff2, ""])

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

# ────────────────────────────────────────────────────────────────
# 5.  Build wide frames for Value Distribution and Population Comparison
# ────────────────────────────────────────────────────────────────
def wide(df_src):
    df_src = df_src[df_src.filemonth_dt.isin(MONTHS)].copy()
    denom = (
        df_src.groupby(["field_name", "filemonth_dt"], as_index=False)
        .value_records.sum()
        .rename(columns={"value_records": "_tot"})
    )
    merged = df_src.merge(denom, on=["field_name", "filemonth_dt"])
    base = (
        merged[["field_name", "value_label"]]
        .drop_duplicates()
        .sort_values(["field_name", "value_label"])
        .reset_index(drop=True)
    )
    for m in MONTHS:
        mm = merged[
            merged.filemonth_dt == m
        ][["field_name", "value_label", "value_records"]]
        base = base.merge(
            mm,
            on=["field_name", "value_label"],
            how="left",
        ).rename(columns={"value_records": fmt(m)})
    cols = [c for c in base.columns if c not in ("field_name", "value_label")]
    base[cols] = base[cols].fillna(0)
    return base

vd_wide = wide(df_data[df_data.analysis_type == "value_dist"])
pc_wide = wide(
    df_data[(df_data.analysis_type == "pop_comp") & (df_data.value_label.apply(_contains))]
)

def add_total_row(df):
    total = {"field_name": "Total", "value_label": "Sum"}
    for c in df.columns:
        if c not in ("field_name", "value_label"):
            total[c] = df[c].sum()
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)

# SQL logic extractor
def sql_for(fld, analysis):
    sub = df_data[
        (df_data.analysis_type == analysis)
        & (df_data.field_name == fld)
        & (df_data.value_sql_logic.notna())
    ]
    if not sub.empty:
        return (
            sub.value_sql_logic.iloc[0]
            .replace("\\n", "\n")
            .replace("\\t", "\t")
        )
    return ""

# ────────────────────────────────────────────────────────────────
# 6.  Load and pivot previous comments from CSV
# ────────────────────────────────────────────────────────────────
prev_comments = pd.read_csv("prev_comments.csv", parse_dates=["date"])
prev_months = MONTHS[1:]  # exclude current month
temp = prev_comments.copy()
temp['month'] = temp['date'].dt.to_period('M').dt.to_timestamp()
temp = temp[temp.month.isin(prev_months)]
miss = temp[temp.research == 'Missing']
m2m = temp[temp.research == 'M2M Diff']

def pivot_comments(df, prefix):
    pts = (
        df.groupby(['field_name', 'month'])['comment']
        .agg(lambda x: '\n'.join(x))
        .reset_index()
    )
    pts['col'] = pts['month'].apply(fmt)
    wide = pts.pivot(index='field_name', columns='col', values='comment')
    wide.columns = [f"{prefix} {c}" for c in wide.columns]
    return wide.reset_index()

pivot_miss = pivot_comments(miss, 'Prev Missing')
pivot_m2m = pivot_comments(m2m, 'Prev M2M')
prev_comments_wide = pd.merge(
    pivot_miss,
    pivot_m2m,
    on='field_name',
    how='outer'
).fillna('')

# Merge with current-month comments
cur = initial_summary[
    ['Field Name', 'Comment Missing', 'Comment M2M']
].rename(
    columns={
        'Comment Missing': 'Comment Missing This Month',
        'Comment M2M': 'Comment M2M This Month',
    }
)
prev_summary = pd.merge(
    cur,
    prev_comments_wide,
    left_on='Field Name',
    right_on='field_name',
    how='left'
)
prev_summary.drop(columns=['field_name'], inplace=True)

# ────────────────────────────────────────────────────────────────
# 6b.  Prepare display version for Previous Comments tab:
#        drop the "This Month" columns, and recalc widths based on header only
# ────────────────────────────────────────────────────────────────
# keep only the history columns
prev_cols = [
    c for c in prev_summary.columns
    if c not in ['Comment Missing This Month', 'Comment M2M This Month']
]
prev_summary_display = prev_summary[prev_cols]

# widths based solely on header length
style_cell_conditional_prev = []
for col in prev_cols:
    width_px = max(len(col) * 8, 100)
    style_cell_conditional_prev.append({
        'if': {'column_id': col},
        'width': f"{width_px}px"
    })

# ────────────────────────────────────────────────────────────────
# 7.  Dash App Layout
# ────────────────────────────────────────────────────────────────
app = Dash(__name__)
app.layout = html.Div([
    dcc.Store(id='summary-store', data=initial_summary.to_dict('records')),

    html.H2("BDCOMM FRY14M Field Analysis — 13-Month View"),
    dcc.Tabs(
        id='main-tabs',
        children=[

            # ────────────────── Summary ──────────────────
            dcc.Tab(label="Summary", children=[
                html.Div([
                    # filter dropdowns...
                ], style={'marginBottom':'1rem'}),
                dash_table.DataTable(
                    id='summary-table',
                    # make only the two comment columns editable
                    columns=[
                        {
                            "name": c,
                            "id": c,
                            "editable": True if c in ["Comment Missing", "Comment M2M"] else False
                        }
                        for c in initial_summary.columns
                    ],
                    data=[],
                    editable=True,
                    filter_action='none',
                    sort_action='native',
                    row_selectable='single',
                    selected_rows=[],
                    page_size=20,
                    style_table={'overflowX':'auto'}
                ),
            ]),

            # ─────────────── Value Distribution ───────────────
            dcc.Tab(label="Value Distribution", children=[
                # unchanged...
            ]),

            # ─────────── Population Comparison ───────────
            dcc.Tab(label="Population Comparison", children=[
                # unchanged...
            ]),

            # ─────────── Previous Comments ───────────
            dcc.Tab(label="Previous Comments", children=[
                dash_table.DataTable(
                    id='prev-comments-table',
                    # only the history columns
                    columns=[{"name": c, "id": c} for c in prev_cols],
                    data=prev_summary_display.to_dict('records'),
                    filter_action='native',
                    sort_action='native',
                    page_size=20,
                    style_table={'overflowX': 'auto'},
                    style_cell_conditional=style_cell_conditional_prev,
                    style_cell={'whiteSpace': 'normal'}
                )
            ]),

        ],
        style={'display': 'flex', 'flexWrap': 'nowrap'}
    )
])

# ────────────────────────────────────────────────────────────────
# 8.  Callbacks
# ────────────────────────────────────────────────────────────────
@app.callback(
    Output('summary-table','data'),
    Input('summary-store','data'),
    Input('filter-miss1','value'), Input('filter-miss2','value'),
    Input('filter-m2m1','value'), Input('filter-m2m2','value')
)
def filter_summary(store_data, m1, m2, d1, d2):
    df = pd.DataFrame(store_data)
    df = df[df[f"Missing {DATE1:%m/%d/%Y}"].isin(m1)]
    df = df[df[f"Missing {prev_month:%m/%d/%Y}"].isin(m2)]
    df = df[df[f"M2M Diff {DATE1:%m/%d/%Y}"].isin(d1)]
    df = df[df[f"M2M Diff {prev_month:%m/%d/%Y}"].isin(d2)]
    return df.to_dict('records')

@app.callback(
    Output('summary-store','data'),
    Input('vd_comm_btn','n_clicks'), Input('pc_comm_btn','n_clicks'),
    State('vd-table','active_cell'), State('vd-table','data'), State('vd_comm_text','value'),
    State('pc-table','active_cell'), State('pc-table','data'), State('pc_comm_text','value'),
    State('summary-store','data'), prevent_initial_call=True
)
def update_comments(n_vd, n_pc, vd_act, vd_data, vd_txt, pc_act, pc_data, pc_txt, store_data):
    df_sum = pd.DataFrame(store_data)
    trig = dash.callback_context.triggered[0]['prop_id'].split('.')[0]
    if trig == 'vd_comm_btn' and vd_act and vd_txt:
        r = vd_act['row']
        fld = vd_data[r]['field_name']
        lbl = vd_data[r]['value_label']
        ent = f"{lbl} - {vd_txt}"
        m = df_sum['Field Name'] == fld
        old = df_sum.loc[m, 'Comment Missing'].iloc[0]
        df_sum.loc[m, 'Comment Missing'] = (old + '\n' if old else '') + ent
    if trig == 'pc_comm_btn' and pc_act and pc_txt:
        r = pc_act['row']
        fld = pc_data[r]['field_name']
        lbl = pc_data[r]['value_label']
        ent = f"{lbl} - {pc_txt}"
        m = df_sum['Field Name'] == fld
        old = df_sum.loc[m, 'Comment M2M'].iloc[0]
        df_sum.loc[m, 'Comment M2M'] = (old + '\n' if old else '') + ent
    return df_sum.to_dict('records')

@app.callback(
    Output('vd-table','data'), Output('pc-table','data'),
    Output('vd_sql','children'), Output('pc_sql','children'),
    Input('summary-table','selected_rows'), State('summary-table','data')
)
def update_detail(selected, summary_rows):
    if selected:
        fld = summary_rows[selected[0]]['Field Name']
        vd_df = vd_wide[vd_wide['field_name'] == fld]
        pc_df = pc_wide[pc_wide['field_name'] == fld]
        vd_sql = sql_for(fld, 'value_dist')
        pc_sql = sql_for(fld, 'pop_comp')
    else:
        vd_df, pc_df, vd_sql, pc_sql = vd_wide, pc_wide, '', ''
    return (
        add_total_row(vd_df).to_dict('records'),
        add_total_row(pc_df).to_dict('records'),
        vd_sql,
        pc_sql,
    )

@app.callback(
    Output('vd-val-lbl','value'),
    Input('vd-table','active_cell'), State('vd-table','data')
)
def update_vd_label(active, rows):
    return rows[active['row']]['value_label'] if active else ''

@app.callback(
    Output('pc-val-lbl','value'),
    Input('pc-table','active_cell'), State('pc-table','data')
)
def update_pc_label(active, rows):
    return rows[active['row']]['value_label'] if active else ''

# ────────────────────────────────────────────────────────────────
# New callback: filter Previous Comments by selected field
# ────────────────────────────────────────────────────────────────
@app.callback(
    Output('prev-comments-table','data'),
    Input('summary-table','selected_rows'),
    State('summary-table','data')
)
def update_prev_comments(selected_rows, summary_rows):
    if selected_rows:
        fld = summary_rows[selected_rows[0]]['Field Name']
        filtered = prev_summary_display[prev_summary_display['Field Name'] == fld]
    else:
        filtered = prev_summary_display
    return filtered.to_dict('records')

# ────────────────────────────────────────────────────────────────
# 9.  Run
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)