import os
import datetime
from io import BytesIO

import pandas as pd
from dateutil.relativedelta import relativedelta

import dash
from dash import dcc, html, Input, Output, State, dash_table
import dash_bootstrap_components as dbc

# -----------------------------------------------------------------------------
# Helper functions (identical logic to your Streamlit app)
# -----------------------------------------------------------------------------

def normalize_columns(df, mapping=None):
    if mapping is None:
        mapping = {
            "edit_nbr": "edit_nbr",
            "edit_error_cnt": "edit_error_cnt",
            "edit_threshold_cnt": "edit_threshold_cnt"
        }
    df.columns = [str(col).strip() for col in df.columns]
    for orig, new in mapping.items():
        for col in df.columns:
            if col.lower() == orig.lower() and col != new:
                df.rename(columns={col: new}, inplace=True)
    return df

def format_date_to_ym(date_obj):
    return date_obj.strftime("%Y-%m")

def load_data_input(excel_path):
    df = pd.read_excel(excel_path, sheet_name="DATA_INPUT")
    df = normalize_columns(df)
    df["filemonth_dt"] = pd.to_datetime(df["filemonth_dt"])
    df["filemonth_str"] = df["filemonth_dt"].apply(format_date_to_ym)
    return df

def load_dqe_thresholds(csv_path):
    df = pd.read_csv(csv_path)
    df.columns = [str(col).strip() for col in df.columns]
    return df

def process_dqe_analysis(data_input_df, thresholds_df, date1):
    # build the list of 5 month‐strings YYYY-MM
    months = [date1 - relativedelta(months=i) for i in range(5)]
    month_strs = [format_date_to_ym(m) for m in months]

    grp = (
        data_input_df
        .groupby(["filemonth_str", "edit_nbr"])
        .agg(edit_error_cnt=("edit_error_cnt", "sum"),
             edit_threshold_cnt=("edit_threshold_cnt", "sum"))
        .reset_index()
    )

    results = []
    for _, thresh in thresholds_df.iterrows():
        edit_nbr_csv = thresh["Edit Nbr"]
        out = thresh.to_dict()
        for m in month_strs:
            filt = grp[(grp["filemonth_str"] == m) & (grp["edit_nbr"] == edit_nbr_csv)]
            err = int(filt["edit_error_cnt"].sum()) if not filt.empty else 0
            thrc = int(filt["edit_threshold_cnt"].sum()) if not filt.empty else 0
            pct = (err / thrc * 100) if thrc > 0 else 0
            try:
                thr_val = float(thresh["Threshold"])
            except:
                thr_val = 0
            status = "Pass" if pct <= thr_val else "Fail"
            out[f"{m} Errors"] = err
            out[f"{m} Error%"]  = round(pct, 2)
            out[f"{m} Status"] = status
            out[f"{m} Error Comments"] = ""
        results.append(out)

    df_out = pd.DataFrame(results)
    return df_out.replace("nan", "")

# -----------------------------------------------------------------------------
# Dash app setup
# -----------------------------------------------------------------------------

app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    title="DQE Analysis"
)
server = app.server  # for gunicorn, etc.

FOLDERS = ["BDCOM", "WFHMSA", "BCards"]

app.layout = dbc.Container(fluid=True, children=[
    dbc.Row(dbc.Col(html.H1("DQE Analysis Report"), className="my-4")),

    dbc.Row([
        dbc.Col([
            dbc.FormGroup([
                dbc.Label("Select Folder"),
                dcc.Dropdown(
                    id="folder-dropdown",
                    options=[{"label": f, "value": f} for f in FOLDERS],
                    value=FOLDERS[0]
                ),
            ]),
            html.Div(id="folder-path", className="text-muted mb-2"),

            dbc.FormGroup([
                dbc.Label("Select DATA_INPUT Excel File"),
                dcc.Dropdown(id="excel-dropdown"),
                html.Div(id="csv-warning", className="text-danger mt-1"),
            ]),

            dbc.FormGroup([
                dbc.Label("Select Analysis Date (Date1)"),
                dcc.DatePickerSingle(
                    id="date-picker",
                    date=datetime.date(2025, 1, 1),
                    display_format="YYYY-MM-DD"
                ),
            ]),
        ], width=3),

        dbc.Col([
            dash_table.DataTable(
                id="dqe-table",
                columns=[],
                data=[],
                editable=True,
                filter_action="native",
                sort_action="native",
                page_size=30,
                style_table={'overflowX': 'auto', 'maxHeight': '600px'},
                style_cell={
                    'whiteSpace': 'normal',
                    'height': 'auto',
                    'textAlign': 'left',
                    'padding': '5px'
                },
                style_header={
                    'backgroundColor': '#f8f9fa',
                    'fontWeight': 'bold'
                },
            ),
            html.Br(),
            dbc.Button(
                "Download DQE Analysis Report as Excel",
                id="download-button",
                color="primary"
            ),
            dcc.Download(id="download-report"),
        ], width=9),
    ]),
])

# -----------------------------------------------------------------------------
# Callbacks
# -----------------------------------------------------------------------------

@app.callback(
    Output("folder-path",         "children"),
    Output("excel-dropdown",      "options"),
    Output("excel-dropdown",      "value"),
    Output("csv-warning",         "children"),
    Input("folder-dropdown",      "value"),
)
def refresh_file_list(folder):
    base = os.path.join(os.getcwd(), folder)
    if not os.path.exists(base):
        return "", [], None, f"Folder '{folder}' not found."
    excels = [f for f in os.listdir(base) if f.lower().endswith((".xlsx"," .xlsb"))]
    opts = [{"label": f, "value": f} for f in excels]
    warn = ""
    if not excels:
        warn = f"No Excel files found in '{folder}'."
    if not os.path.exists(os.path.join(base, "dqe_thresholds.csv")):
        warn += "\nMissing 'dqe_thresholds.csv'."
    return f"Folder path: {base}", opts, (opts[0]["value"] if opts else None), warn

@app.callback(
    Output("dqe-table", "data"),
    Output("dqe-table", "columns"),
    Input("folder-dropdown", "value"),
    Input("excel-dropdown",  "value"),
    Input("date-picker",     "date"),
)
def update_table(folder, excel_file, date):
    if not (folder and excel_file and date):
        return [], []
    base = os.path.join(os.getcwd(), folder)
    data_input = load_data_input(os.path.join(base, excel_file))
    thresholds = load_dqe_thresholds(os.path.join(base, "dqe_thresholds.csv"))
    dt1 = datetime.datetime.fromisoformat(date)
    out_df = process_dqe_analysis(data_input, thresholds, dt1)

    cols = [
        {"name": c, "id": c, "editable": ("Error Comments" in c)}
        for c in out_df.columns
    ]
    return out_df.to_dict("records"), cols

@app.callback(
    Output("download-report", "data"),
    Input("download-button", "n_clicks"),
    State("dqe-table",    "data"),
    State("dqe-table",    "columns"),
    prevent_initial_call=True
)
def func_download(n, rows, cols):
    df = pd.DataFrame(rows, columns=[c["id"] for c in cols])
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DQE Analysis")
    buf.seek(0)
    return dcc.send_bytes(buf.read(), "DQE_Analysis_Report.xlsx")

# -----------------------------------------------------------------------------
if __name__ == "__main__":
    app.run_server(debug=True)