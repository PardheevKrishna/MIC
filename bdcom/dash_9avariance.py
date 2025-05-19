import os
from io import BytesIO

import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta

import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
from dash_ag_grid import AgGrid

# -----------------------------------------------------------------------------
# Helper Functions (same logic as your Streamlit version)
# -----------------------------------------------------------------------------

def normalize_columns(df):
    df.columns = [str(col).strip() for col in df.columns]
    return df

def load_output_data(excel_path):
    # Load the OUTPUT sheet (header row 0)
    output_df = pd.read_excel(excel_path, sheet_name="OUTPUT", header=0)
    return normalize_columns(output_df)

def load_variance_analysis_sheet(excel_path):
    # Load the Variance_Analysis sheet (header is row 8 → header=7)
    var_df = pd.read_excel(excel_path, sheet_name="Variance_Analysis", header=7)
    return normalize_columns(var_df)

def process_variance_analysis(output_df, var_df):
    # Compute current and prior values
    current_values = []
    prior_values = []
    for _, row in var_df.iterrows():
        name = row["Field Name"]
        curr = output_df[output_df.iloc[:,1] == name].iloc[:,2].sum() if name in output_df.iloc[:,1].values else 0
        prior = output_df[output_df.iloc[:,4] == name].iloc[:,5].sum() if name in output_df.iloc[:,4].values else 0
        current_values.append(curr)
        prior_values.append(prior)

    var_df["Current Value"] = current_values
    var_df["Prior value"]  = prior_values
    var_df["$Variance"]    = var_df["Current Value"] - var_df["Prior value"]
    var_df["%Variance"]    = var_df.apply(
        lambda r: (r["$Variance"] / r["Prior value"] * 100) if r["Prior value"] != 0 else 0,
        axis=1
    )

    # Ensure editable columns exist
    for col in ["Comments", "Detail File Link"]:
        if col not in var_df.columns:
            var_df[col] = ""

    # Final column order
    final_cols = [
        "14M file", "Field No.", "Field Name",
        "Current Value", "Prior value", "$Variance", "%Variance",
        "$Tolerance", "%Tolerance", "Comments", "Detail File Link"
    ]
    for col in final_cols:
        if col not in var_df.columns:
            var_df[col] = ""
    return var_df[final_cols].replace("nan", "")

# -----------------------------------------------------------------------------
# Dash app setup
# -----------------------------------------------------------------------------

external_stylesheets = [
    dbc.themes.BOOTSTRAP,
    "https://unpkg.com/ag-grid-community/dist/styles/ag-grid.css",
    "https://unpkg.com/ag-grid-community/dist/styles/ag-theme-alpine.css",
]

app = dash.Dash(__name__, external_stylesheets=external_stylesheets, title="Variance Analysis")
server = app.server

FOLDERS = ["BDCOM", "WFHMSA", "BCards"]

app.layout = dbc.Container(fluid=True, children=[

    # Title
    html.H1("Variance Analysis Report", className="text-center my-4"),

    # Folder & File selectors
    dbc.Row(
        [
            dbc.Col([
                html.Label("Folder", className="form-label"),
                dcc.Dropdown(
                    id="folder-dropdown",
                    options=[{"label": f, "value": f} for f in FOLDERS],
                    value=FOLDERS[0],
                    className="form-select"
                ),
            ], width=4),

            dbc.Col([
                html.Label("Excel File", className="form-label"),
                dcc.Dropdown(id="file-dropdown", className="form-select"),
            ], width=8),
        ],
        className="mb-4 align-items-end"
    ),

    # AG-Grid table + Download
    dbc.Row(
        dbc.Col([
            AgGrid(
                id="variance-grid",
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
                style={"width": "100%", "height": "70vh"}
            ),
            html.Br(),
            dbc.Button("Download Report as Excel", id="download-btn", color="primary"),
            dcc.Download(id="download-report")
        ], width=12)
    ),

])

# -----------------------------------------------------------------------------
# Callbacks
# -----------------------------------------------------------------------------

@app.callback(
    Output("file-dropdown", "options"),
    Output("file-dropdown", "value"),
    Input("folder-dropdown", "value"),
)
def update_file_list(folder):
    folder_path = os.path.join(os.getcwd(), folder)
    if not os.path.isdir(folder_path):
        return [], None
    files = [
        f for f in os.listdir(folder_path)
        if f.lower().endswith((".xlsx", ".xlsb"))
    ]
    opts = [{"label": f, "value": f} for f in files]
    return opts, (opts[0]["value"] if opts else None)

@app.callback(
    Output("variance-grid", "columnDefs"),
    Output("variance-grid", "rowData"),
    Input("folder-dropdown", "value"),
    Input("file-dropdown",   "value"),
)
def update_grid(folder, filename):
    if not (folder and filename):
        return [], []
    path = os.path.join(os.getcwd(), folder, filename)
    output_df = load_output_data(path)
    var_df    = load_variance_analysis_sheet(path)
    result_df = process_variance_analysis(output_df, var_df)

    # Build column definitions
    col_defs = []
    for col in result_df.columns:
        cfg = {
            "headerName": col,
            "field": col,
            "filter": True,
            "sortable": True,
            "resizable": True,
            "minWidth": 120
        }
        if col in ["Comments", "Detail File Link"]:
            cfg["editable"] = True
        if col in ["14M file", "Field No.", "Field Name"]:
            cfg["pinned"] = "left"
        col_defs.append(cfg)

    return col_defs, result_df.to_dict("records")

@app.callback(
    Output("download-report", "data"),
    Input("download-btn",     "n_clicks"),
    State("variance-grid",    "rowData"),
    State("folder-dropdown",  "value"),
    State("file-dropdown",    "value"),
    prevent_initial_call=True
)
def download_excel(n_clicks, rows, folder, filename):
    # Reconstruct the two sheets
    df_updated = pd.DataFrame(rows)
    path       = os.path.join(os.getcwd(), folder, filename)
    output_df  = load_output_data(path)

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_updated.to_excel(writer, index=False, sheet_name="Variance Analysis")
        output_df.to_excel(writer, index=False, sheet_name="OUTPUT")
    buf.seek(0)
    return dcc.send_bytes(buf.read(), "Variance_Analysis_Report.xlsx")

# -----------------------------------------------------------------------------
if __name__ == "__main__":
    server.run(debug=True, port=8052)