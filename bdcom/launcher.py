# dashboard_launcher.py

import os
import glob
import multiprocessing
import sys

from dash import Dash, dcc, html
from dash.dependencies import Input, Output, State

# ────────────────────────────────────────────────────────────────
# 1.  Wrappers that spawn each dashboard in its own process
# ────────────────────────────────────────────────────────────────

def _start_field(input_file, port=8051):
    # import and override, then run with no reloader
    import dash_fieldanalysis
    dash_fieldanalysis.INPUT_FILE = input_file
    dash_fieldanalysis.app.run = dash_fieldanalysis.app.run_server
    dash_fieldanalysis.app.run(debug=True, port=port, use_reloader=False)

def _start_dqe(input_file, port=8052):
    import dash_dqe
    dash_dqe.INPUT_FILE = input_file
    dash_dqe.app.run = dash_dqe.app.run_server
    dash_dqe.app.run(debug=True, port=port, use_reloader=False)

def _start_9avar(input_file, port=8053):
    import dash_9avar
    dash_9avar.INPUT_FILE = input_file
    dash_9avar.app.run = dash_9avar.app.run_server
    dash_9avar.app.run(debug=True, port=port, use_reloader=False)


# ────────────────────────────────────────────────────────────────
# 2.  Build the “launcher” Dash app
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
                placeholder="Choose folder & analysis first…"
            ),
        ]),
    ]),

    html.Button("Start", id="start-btn", n_clicks=0,
                className="btn btn-primary"),
    html.Div(id="start-output", className="mt-3")
], className="container mt-5")


# ────────────────────────────────────────────────────────────────
# 3.  Populate the Excel-file dropdown
# ────────────────────────────────────────────────────────────────
@launcher.callback(
    Output("file-dropdown", "options"),
    Input("folder-dropdown", "value"),
    Input("analysis-dropdown", "value")
)
def _update_file_list(folder, analysis):
    if not folder or not analysis:
        return []
    folder_path = os.path.join(BASE_PATH, folder)
    files = sorted(glob.glob(os.path.join(folder_path, "*.xlsx")))
    return [{"label": os.path.basename(f), "value": f} for f in files]


# ────────────────────────────────────────────────────────────────
# 4.  When Start is clicked → spawn the right dashboard
# ────────────────────────────────────────────────────────────────
@launcher.callback(
    Output("start-output", "children"),
    Input("start-btn", "n_clicks"),
    State("analysis-dropdown", "value"),
    State("file-dropdown", "value"),
    prevent_initial_call=True
)
def _launch(n_clicks, analysis, filepath):
    if not (analysis and filepath):
        return html.Div(
            "⚠️ Please select an Analysis type and an Excel file.",
            style={"color":"red"}
        )

    if analysis == "field":
        target, port, label = _start_field, 8051, "Field Analysis"
    elif analysis == "dqe":
        target, port, label = _start_dqe, 8052, "DQE"
    else:
        target, port, label = _start_9avar, 8053, "9a_Var"

    # spawn it in its own process
    proc = multiprocessing.Process(target=target, args=(filepath, port))
    proc.daemon = True
    proc.start()

    link = f"http://127.0.0.1:{port}"
    return html.Div([
        f"🚀 Launched **{label}** on port {port}. ",
        html.A("Open it here", href=link, target="_blank")
    ])


# ────────────────────────────────────────────────────────────────
# 5.  Run the launcher itself on 8050
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # alias for convenience
    launcher.run = launcher.run_server
    launcher.run(debug=True, port=8050)