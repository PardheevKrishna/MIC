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
        mm = merged[merged.filemonth_dt == m][["field_name", "value_label", "value_records"]]
        base = base.merge(mm, on=["field_name", "value_label"], how="left").rename(columns={"value_records": fmt(m)})
    cols = [c for c in base.columns if c not in ("field_name", "value_label")]
    base[cols] = base[cols].fillna(0)
    return base

vd_wide = wide(df_data[df_data.analysis_type == "value_dist"])
pc_wide = wide(df_data[(df_data.analysis_type == "pop_comp") & (df_data.value_label.apply(_contains))])

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
        return sub.value_sql_logic.iloc[0].replace("\\n", "\n").replace("\\t", "\t")
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
prev_comments_wide = pd.merge(pivot_miss, pivot_m2m, on='field_name', how='outer').fillna('')

# Merge with current-month comments and then drop them for previous tab
cur = initial_summary[['Field Name', 'Comment Missing', 'Comment M2M']].rename(
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
# Drop helper and this-month comment columns
prev_summary.drop(columns=['field_name', 'Comment Missing This Month', 'Comment M2M This Month'], inplace=True)

# Pre-calc style_cell_conditional based on header length only
style_cell_conditional = []
for col in prev_summary.columns:
    width_px = max(len(col) * 8, 100)
    style_cell_conditional.append(
        {'if': {'column_id': col}, 'width': f"{width_px}px"}
    )

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
            dcc.Tab(label="Summary", children=[
                html.Div([
                    html.Div([
                        html.Label(f"Missing {DATE1:%b-%Y}"),
                        dcc.Dropdown(
                            id='filter-miss1',
                            options=[{'label': v, 'value': v} for v in sorted(initial_summary[f"Missing {DATE1:%m/%d/%Y}"].unique())],
                            value=list(sorted(initial_summary[f"Missing {DATE1:%m/%d/%Y}"].unique())),
                            multi=True
                        ),
                    ], style={'width':'24%','display':'inline-block'}),
                    html.Div([
                        html.Label(f"Missing {prev_month:%b-%Y}"),
                        dcc.Dropdown(
                            id='filter-miss2',
                            options=[{'label': v, 'value': v} for v in sorted(initial_summary[f"Missing {prev_month:%m/%d/%Y}"].unique())],
                            value=list(sorted(initial_summary[f"Missing {prev_month:%m/%d/%Y}"].unique())),
                            multi=True
                        ),
                    ], style={'width':'24%','display':'inline-block'}),
                    html.Div([
                        html.Label(f"M2M Diff {DATE1:%b-%Y}"),
                        dcc.Dropdown(
                            id='filter-m2m1',
                            options=[{'label': v, 'value': v} for v in sorted(initial_summary[f"M2M Diff {DATE1:%m/%d/%Y}"].unique())],
                            value=list(sorted(initial_summary[f"M2M Diff {DATE1:%m/%d/%Y}"].unique())),
                            multi=True
                        ),
                    ], style={'width':'24%','display':'inline-block'}),
                    html.Div([
                        html.Label(f"M2M Diff {prev_month:%b-%Y}"),
                        dcc.Dropdown(
                            id='filter-m2m2',
                            options=[{'label': v, 'value': v} for v in sorted(initial_summary[f"M2M Diff {prev_month:%m/%d/%Y}"].unique())],
                            value=list(sorted(initial_summary[f"M2M Diff {prev_month:%m/%d/%Y}"].unique())),
                            multi=True
                        ),
                    ], style={'width':'24%','display':'inline-block'}),
                ], style={'marginBottom':'1rem'}),
                dash_table.DataTable(
                    id='summary-table',
                    columns=[
                        {"name": c, "id": c, "editable": c in ["Comment Missing", "Comment M2M"]}
                        for c in initial_summary.columns
                    ],
                    data=[],
                    filter_action='none',
                    sort_action='native',
                    row_selectable='single',
                    selected_rows=[],
                    page_size=20,
                    style_table={'overflowX':'auto'}
                ),
            ]),
            dcc.Tab(label="Value Distribution", children=[
                dash_table.DataTable(
                    id='vd-table',
                    columns=[{"name":c,"id":c} for c in vd_wide.columns],
                    data=add_total_row(vd_wide).to_dict('records'),
                    filter_action='native', sort_action='native',
                    page_size=20, style_table={'overflowX':'auto'}
                ),
                dcc.Input(id='vd_val_lbl', type='text', readOnly=True, placeholder='value_label (select any cell)', style={'width':'100%','marginTop':'0.5rem'}),
                dcc.Textarea(id='vd_comm_text', placeholder='Add comment…', style={'width':'100%','height':'60px','marginTop':'0.5rem'}),
                html.Button('Submit', id='vd_comm_btn', n_clicks=0, style={'marginTop':'0.25rem'}),
                html.Pre(id='vd_sql', style={'whiteSpace':'pre-wrap','backgroundColor':'#f3f3f3','padding':'0.75rem','border':'1px solid #ddd','fontFamily':'monospace','fontSize':'0.85rem','marginTop':'0.5rem'}),
                dcc.Clipboard(target_id='vd_sql', title='Copy SQL Logic', style={'marginTop':'0.5rem'})
            ]),
            dcc.Tab(label="Population Comparison", children=[
                dash_table.DataTable(
                    id='pc-table',
                    columns=[{"name":c,"id":c} for c in pc_wide.columns],
                    data=add_total_row(pc_wide).to_dict('records'),
                    filter_action='native', sort_action='native',
                    page_size=20, style_table={'overflowX':'auto'}
                ),
                dcc.Input(id='pc_val_lbl', type='text', readOnly=True, placeholder='value_label (select any cell)', style={'width':'100%','marginTop':'0.5rem'}),
                dcc.Textarea(id='pc_comm_text', placeholder='Add comment…', style={'width':'100%','height':'60px','marginTop':'0.5rem'}),
                html.Button('Submit', id='pc_comm_btn', n_clicks=0, style={'marginTop':'0.25rem'}),
                html.Pre(id='pc_sql', style={'whiteSpace':'pre-wrap','backgroundColor':'#f3f3f3','padding':'0.75rem','border':'1px solid #ddd','fontFamily':'monospace','fontSize':'0.85rem','marginTop':'0.5rem'}),
                dcc.Clipboard(target_id='pc_sql', title='Copy SQL Logic', style={'marginTop':'0.5rem'})
            ]),
            dcc.Tab(label="Previous Comments", children=[
                dash_table.DataTable(
                    id='prev-comments-table',
                    columns=[{"name": c, "id": c} for c in prev_summary.columns],
                    data=prev_summary.to_dict('records'),
                    filter_action='native',
                    sort_action='native',
                    page_size=20,
                    style_table={'overflowX': 'auto'},
                    style_cell_conditional=style_cell_conditional,
                    style_cell={'whiteSpace': 'normal'}
                )
            ])
        ],
        style={'display':'flex', 'flexWrap':'nowrap'}
    )
])

# ────────────────────────────────────────────────────────────────
# 8.  Callbacks (unchanged)
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
        lbl = pc_data[r]['
