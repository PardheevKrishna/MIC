import os
import datetime
from io import BytesIO

import pandas as pd
from dateutil.relativedelta import relativedelta

import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
from dash_ag_grid import AgGrid

# -----------------------------------------------------------------------------
# Helper functions (identical to your Streamlit logic)
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
            out[f"{m} Errors"]         = err
            out[f"{m} Error%"]         = round(pct, 2)
            out[f"{m} Status"]         = status
            out[f"{m} Error Comments"] = ""
        results.append(out)

    df_out = pd.DataFrame(results)
    return df_out.replace("nan", "")

# -----------------------------------------------------------------------------
# Dash app setup
# -----------------------------------------------------------------------------

app = dash.Dash(
    __name__,
    external_stylesheets=[
        # Bootstrap for layout
        dbc.themes.BOOTSTRAP,
        # AG-Grid core + theme CSS
        "https://unpkg.com/ag-grid-community/dist/styles/ag-grid.css",
        "https://unpkg.com/ag-grid-community/dist/styles/ag-theme-alpine.css",
    ],
    title="DQE Analysis"
)
server = app.server

FOLDERS = ["BDCOM", "WFHMSA", "BCards"]

app.layout = dbc.Container(fluid=True, children=[

    # Title with top margin
    html.H1("DQE Analysis", className="text-center my-4"),

    # Single row of 3 selectors
    dbc.Row(
        [
            dbc.Col(
                html.Div([
                    html.Label("Folder", className="form-label"),
                    dcc.Dropdown(
                        id="folder-dropdown",
                        options=[{"label": f, "value": f} for f in FOLDERS],
                        value=FOLDERS[0],
                        className="form-select"
                    ),
                ]),
                width=3
            ),
            dbc.Col(
                html.Div([
                    html.Label("DATA_INPUT Excel File", className="form-label"),
                    dcc.Dropdown(id="excel-dropdown", className="form-select"),
                ]),
                width=6
            ),
            dbc.Col(
                html.Div([
                    html.Label("Analysis Date (Date1)", className="form-label"),
                    dcc.DatePickerSingle(
                        id="date-picker",
                        date=datetime.date(2025, 1, 1),
                        display_format="YYYY-MM-DD",
                        className="w-100"
                    ),
                ]),
                width=3
            ),
        ],
        className="mb-4 align-items-end"
    ),

    # AG-Grid & Download button
    dbc.Row(
        dbc.Col([
            AgGrid(
                id="dqe-grid",
                columnDefs=[],
                rowData=[],
                defaultColDef={
                    "filter": True,
                    "sortable": True,
                    "resizable": True,
                    "wrapText": True,
                    "autoHeight": True
                },
                className="ag-theme-alpine",
                style={"width": "100%", "height": "calc(100vh - 250px)"}
            ),
            html.Br(),
            dbc.Button("Download Excel Report", id="download-button", color="primary"),
            dcc.Download(id="download-report"),
        ], width=12)
    ),

])

# -----------------------------------------------------------------------------
# Callbacks
# -----------------------------------------------------------------------------

@app.callback(
    Output("excel-dropdown", "options"),
    Output("excel-dropdown", "value"),
    Input("folder-dropdown", "value"),
)
def update_file_list(folder):
    folder_path = os.path.join(os.getcwd(), folder)
    if not os.path.exists(folder_path):
        return [], None
    files = [f for f in os.listdir(folder_path) if f.lower().endswith((".xlsx", ".xlsb"))]
    opts = [{"label": f, "value": f} for f in files]
    return opts, (opts[0]["value"] if opts else None)

@app.callback(
    Output("dqe-grid", "columnDefs"),
    Output("dqe-grid", "rowData"),
    Input("folder-dropdown", "value"),
    Input("excel-dropdown",  "value"),
    Input("date-picker",     "date"),
)
def update_grid(folder, excel_file, date):
    if not (folder and excel_file and date):
        return [], []
    base = os.path.join(os.getcwd(), folder)
    data_input  = load_data_input(os.path.join(base, excel_file))
    thresholds  = load_dqe_thresholds(os.path.join(base, "dqe_thresholds.csv"))
    dt1         = datetime.datetime.fromisoformat(date)
    result_df   = process_dqe_analysis(data_input, thresholds, dt1)

    # Build AG-Grid column definitions
    col_defs = []
    for col in result_df.columns:
        col_defs.append({
            "headerName": col,
            "field": col,
            "filter": True,
            "sortable": True,
            "resizable": True,
            "editable": True if "Error Comments" in col else False,
            "minWidth": 150
        })

    return col_defs, result_df.to_dict("records")

@app.callback(
    Output("download-report", "data"),
    Input("download-button", "n_clicks"),
    State("dqe-grid", "rowData"),
    State("dqe-grid", "columnDefs"),
    prevent_initial_call=True
)
def download_excel(n_clicks, rows, cols):
    # Reconstruct DataFrame in original column order
    fields = [c["field"] for c in cols]
    df = pd.DataFrame(rows, columns=fields)

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DQE Analysis")
    buf.seek(0)
    return dcc.send_bytes(buf.read(), "DQE_Analysis_Report.xlsx")

# -----------------------------------------------------------------------------
if __name__ == "__main__":
    # Use server.run() instead of app.run_server()
    server.run(debug=True, port=8051)