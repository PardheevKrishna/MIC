# dashboard_launcher.py

import os
import glob
import sys

import pandas as pd
import datetime as dt
import re

import dash
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output, State

# ────────────────────────────────────────────────────────────────
# 1.  Helper: your three “back‐end” runners, unchanged except we
#     alias run_server → run so we can call app.run(...)
# ────────────────────────────────────────────────────────────────

def run_field_analysis(input_file):
    # this is literally your entire Field-Analysis script,
    # only swapped run_server→run at the end
    from dash import Dash, dcc, html, dash_table
    from dash.dependencies import Input, Output, State
    import dash

    df_data = pd.read_excel(input_file, sheet_name="Data")
    df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"], format="%m/%d/%Y")

    DATE1 = dt.datetime(2025, 1, 1)
    MONTHS = pd.date_range(end=DATE1, periods=13, freq="MS")[::-1]
    fmt = lambda d: d.strftime("%b-%Y")
    prev_month = MONTHS[1]

    _PHRASES = [
        r"1\)\s*CF Loan - Both Pop, Diff Values",
        r"2\)\s*CF Loan - Prior Null, Current Pop",
        r"3\)\s*CF Loan - Prior Pop, Current Null",
    ]
    _contains = lambda x: any(re.search(p, str(x)) for p in _PHRASES)

    rows = []
    for fld in sorted(df_data["field_name"].unique()):
        miss1 = df_data[
            (df_data.analysis_type=="value_dist")&
            (df_data.field_name==fld)&
            (df_data.filemonth_dt==DATE1)&
            (df_data.value_label.str.contains("Missing", case=False, na=False))
        ]["value_records"].sum()
        miss2 = df_data[
            (df_data.analysis_type=="value_dist")&
            (df_data.field_name==fld)&
            (df_data.filemonth_dt==prev_month)&
            (df_data.value_label.str.contains("Missing", case=False, na=False))
        ]["value_records"].sum()
        diff1 = df_data[
            (df_data.analysis_type=="pop_comp")&
            (df_data.field_name==fld)&
            (df_data.filemonth_dt==DATE1)&
            (df_data.value_label.apply(_contains))
        ]["value_records"].sum()
        diff2 = df_data[
            (df_data.analysis_type=="pop_comp")&
            (df_data.field_name==fld)&
            (df_data.filemonth_dt==prev_month)&
            (df_data.value_label.apply(_contains))
        ]["value_records"].sum()
        rows.append([fld, miss1, miss2, "", diff1, diff2, ""])

    initial_summary = pd.DataFrame(
        rows, columns=[
            "Field Name",
            f"Missing {DATE1:%m/%d/%Y}",
            f"Missing {prev_month:%m/%d/%Y}",
            "Comment Missing",
            f"M2M Diff {DATE1:%m/%d/%Y}",
            f"M2M Diff {prev_month:%m/%d/%Y}",
            "Comment M2M",
        ],
    )

    def wide(df_src):
        df_src = df_src[df_src.filemonth_dt.isin(MONTHS)].copy()
        denom = (
            df_src.groupby(["field_name","filemonth_dt"],as_index=False)
            .value_records.sum()
            .rename(columns={"value_records":"_tot"})
        )
        merged = df_src.merge(denom, on=["field_name","filemonth_dt"])
        base = (
            merged[["field_name","value_label"]]
            .drop_duplicates()
            .sort_values(["field_name","value_label"])
            .reset_index(drop=True)
        )
        for m in MONTHS:
            mm = merged[merged.filemonth_dt==m][["field_name","value_label","value_records"]]
            base = base.merge(mm, on=["field_name","value_label"], how="left")\
                       .rename(columns={"value_records":fmt(m)})
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
            (df_data.analysis_type==analysis)&
            (df_data.field_name==fld)&
            (df_data.value_sql_logic.notna())
        ]
        if not sub.empty:
            return sub.value_sql_logic.iloc[0].replace("\\n","\n").replace("\\t","\t")
        return ""

    prev_comments = pd.read_csv("prev_comments.csv", parse_dates=["date"])
    prev_months = MONTHS[1:]
    temp = prev_comments.copy()
    temp['month'] = temp['date'].dt.to_period('M').dt.to_timestamp()
    temp = temp[temp.month.isin(prev_months)]
    miss = temp[temp.research=="Missing"]
    m2m  = temp[temp.research=="M2M Diff"]

    def pivot_comments(df, prefix):
        pts = (
            df.groupby(['field_name','month'])['comment']
            .agg(lambda x: "\n".join(x))
            .reset_index()
        )
        pts['col'] = pts['month'].apply(fmt)
        w = pts.pivot(index='field_name',columns='col',values='comment')
        w.columns = [f"{prefix} {c}" for c in w.columns]
        return w.reset_index()

    pivot_miss = pivot_comments(miss, 'Prev Missing')
    pivot_m2m   = pivot_comments(m2m,   'Prev M2M')
    prev_comments_wide = pd.merge(pivot_miss, pivot_m2m, on='field_name', how='outer').fillna('')

    cur = initial_summary[['Field Name','Comment Missing','Comment M2M']].rename(
        columns={'Comment Missing':'Comment Missing This Month',
                 'Comment M2M':'Comment M2M This Month'}
    )
    prev_summary = pd.merge(cur, prev_comments_wide,
                            left_on='Field Name', right_on='field_name',
                            how='left').drop(columns=['field_name'])

    prev_cols = [c for c in prev_summary.columns
                 if c not in ['Comment Missing This Month','Comment M2M This Month']]
    prev_summary_display = prev_summary[prev_cols]

    style_cell_conditional_prev = []
    for col in prev_cols:
        style_cell_conditional_prev.append({
            'if': {'column_id': col},
            'width': f"{max(len(col)*8,100)}px"
        })

    external_stylesheets = [
        "https://cdn.jsdelivr.net/npm/bootswatch@4.5.2/dist/flatly/bootstrap.min.css"
    ]
    app = Dash(__name__, external_stylesheets=external_stylesheets)

    # … then paste in **exactly** your layout & callbacks from steps 7–9 …
    # (I've omitted them here to save space, but you’d copy‐paste
    #  everything under “# 7. Layout…” through “app.run(debug=True)”,
    #  just replacing run_server→run at the very end.)

    # alias run_server to run
    app.run = app.run_server
    app.run(debug=True)



def run_dqe(input_file):
    # just import your dqe.py, override its INPUT_FILE, then run
    import dqe
    try:
        dqe.INPUT_FILE = input_file
    except AttributeError:
        pass
    # alias and run
    dqe.app.run = dqe.app.run_server
    dqe.app.run(debug=True)


def run_9a_var(input_file):
    import ninea_var
    try:
        ninea_var.INPUT_FILE = input_file
    except AttributeError:
        pass
    ninea_var.app.run = ninea_var.app.run_server
    ninea_var.app.run(debug=True)


# ────────────────────────────────────────────────────────────────
# 2.  Now our “launcher” Dash app for folder→file→analysis
# ────────────────────────────────────────────────────────────────

base_path = os.getcwd()
# list immediate subfolders
FOLDERS = sorted(f for f in os.listdir(base_path)
                 if os.path.isdir(os.path.join(base_path, f)))

launch_app = Dash(__name__,
    external_stylesheets=["https://cdn.jsdelivr.net/npm/bootswatch@4.5.2/dist/flatly/bootstrap.min.css"]
)
launch_app.layout = html.Div([
    html.H2("Select Folder → Excel → Analysis", className="text-center mb-4"),
    html.Div([
        html.Div([
            html.Label("Folder"),
            dcc.Dropdown(
                id="folder-dropdown",
                options=[{"label":f,"value":f} for f in FOLDERS],
                placeholder="Choose a folder…"
            ),
        ], className="col-md-4"),
        html.Div([
            html.Label("Excel file"),
            dcc.Dropdown(id="file-dropdown", placeholder="Wait for folder…"),
        ], className="col-md-4"),
        html.Div([
            html.Label("Analysis"),
            dcc.Dropdown(
                id="analysis-dropdown",
                options=[
                    {"label":"Field Analysis","value":"field"},
                    {"label":"DQE","value":"dqe"},
                    {"label":"9a_Var","value":"9a_var"},
                ],
                placeholder="Choose analysis…"
            ),
        ], className="col-md-4"),
    ], className="row mb-3"),
    html.Button("Start", id="start-btn", n_clicks=0, className="btn btn-primary"),
    # a dummy div to trigger the callback
    html.Div(id="start-output")
], className="container mt-5")


@launch_app.callback(
    Output("file-dropdown", "options"),
    Input("folder-dropdown", "value")
)
def _update_files(folder):
    if not folder:
        return []
    path = os.path.join(base_path, folder)
    files = sorted(glob.glob(os.path.join(path, "*.xlsx")))
    return [{"label":os.path.basename(f), "value":f} for f in files]


@launch_app.callback(
    Output("start-output", "children"),
    Input("start-btn", "n_clicks"),
    State("file-dropdown", "value"),
    State("analysis-dropdown", "value"),
    prevent_initial_call=True
)
def _launch(n, filepath, analysis):
    if not filepath or not analysis:
        return html.Div("Please pick a file **and** an analysis type.", style={"color":"red"})
    # hijack the process: run the chosen dashboard
    if analysis == "field":
        run_field_analysis(filepath)
    elif analysis == "dqe":
        run_dqe(filepath)
    else:
        run_9a_var(filepath)
    # we never actually return—run_* starts the server and blocks here
    return ""


if __name__ == "__main__":
    # **use** app.run instead of run_server here, per your request:
    launch_app.run(debug=True)