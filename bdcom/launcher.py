# dashboard_launcher.py

import os
import glob
import sys
import subprocess

from dash import Dash, dcc, html
from dash.dependencies import Input, Output, State

# ────────────────────────────────────────────────────────────────
# 1.  Launcher Dash app: Folder → Analysis → Excel → Start
# ────────────────────────────────────────────────────────────────

BASE_PATH = os.getcwd()
FOLDERS = sorted(
    d for d in os.listdir(BASE_PATH)
    if os.path.isdir(os.path.join(BASE_PATH, d))
)

launcher = Dash(__name__,
    external_stylesheets=[
        "https://cdn.jsdelivr.net/npm/bootswatch@4.5.2/dist/flatly/bootstrap.min.css"
    ]
)

launcher.layout = html.Div([
    html.H2("📁 Folder → Analysis → Excel Launcher",
            className="text-center mb-4"),

    html.Div(className="row mb-3", children=[

        html.Div(className="col-md-4", children=[
            html.Label("Folder"),
            dcc.Dropdown(
                id="folder-dropdown",
                options=[{"label":f,"value":f} for f in FOLDERS],
                placeholder="Select folder…"
            ),
        ]),

        html.Div(className="col-md-4", children=[
            html.Label("Analysis"),
            dcc.Dropdown(
                id="analysis-dropdown",
                options=[
                    {"label":"Field Analysis","value":"field"},
                    {"label":"DQE","value":"dqe"},
                    {"label":"9a_Var","value":"9a_var"},
                ],
                placeholder="Select analysis…"
            ),
        ]),

        html.Div(className="col-md-4", children=[
            html.Label("Excel File"),
            dcc.Dropdown(
                id="file-dropdown",
                placeholder="Choose folder first…"
            ),
        ]),
    ]),

    html.Button("Start", id="start-btn", n_clicks=0,
                className="btn btn-primary"),
    html.Div(id="start-output", className="mt-3")
], className="container mt-5")


# ────────────────────────────────────────────────────────────────
# 2.  Populate Excel‐file dropdown when folder is chosen
# ────────────────────────────────────────────────────────────────
@launcher.callback(
    Output("file-dropdown","options"),
    Input("folder-dropdown","value")
)
def update_file_list(folder):
    if not folder:
        return []
    folder_path = os.path.join(BASE_PATH, folder)
    files = sorted(glob.glob(os.path.join(folder_path, "*.xlsx")))
    return [{"label":os.path.basename(f), "value":f} for f in files]


# ────────────────────────────────────────────────────────────────
# 3.  On Start: spawn the selected dash_xxx.py in its own process
# ────────────────────────────────────────────────────────────────
@launcher.callback(
    Output("start-output","children"),
    Input("start-btn","n_clicks"),
    State("analysis-dropdown","value"),
    State("file-dropdown","value"),
    prevent_initial_call=True
)
def launch_dashboard(n_clicks, analysis, filepath):
    if not analysis or not filepath:
        return html.Div(
            "⚠️ Please select an analysis type and an Excel file.",
            style={"color":"red"}
        )

    # map analysis → (module, port)
    mapping = {
        "field":   ("dash_fieldanalysis", 8051),
        "dqe":     ("dash_dqe",           8052),
        "9a_var":  ("dash_9avar",         8053),
    }
    module_name, port = mapping[analysis]

    # build command to run in a fresh interpreter:
    #   python -c "import MODULE; MODULE.INPUT_FILE=r'...'; MODULE.app.run_server(debug=True, port=PORT)"
    cmd = [
        sys.executable, "-c",
        (
            f"import {module_name}; "
            f"{module_name}.INPUT_FILE = r'{filepath}'; "
            f"{module_name}.app.run_server(debug=True, port={port})"
        )
    ]

    # spawn it (non-blocking)
    subprocess.Popen(cmd, cwd=BASE_PATH)

    link = f"http://127.0.0.1:{port}"
    return html.Div([
        f"🚀 Launched **{module_name}** on port {port}. ",
        html.A("Open it here", href=link, target="_blank")
    ])


# ────────────────────────────────────────────────────────────────
# 4.  Run the launcher itself (on port 8050)
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # alias so we can call .run(...)
    launcher.run = launcher.run_server
    launcher.run(debug=True, port=8050)