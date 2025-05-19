# dashboard_launcher.py

import os
import glob
import multiprocessing
from dash import Dash, dcc, html
from dash.dependencies import Input, Output, State

# ────────────────────────────────────────────────────────────────
# 1.  Helper functions to spawn each dashboard in its own process
# ────────────────────────────────────────────────────────────────

def _start_field(filepath):
    import dash_fieldanalysis
    dash_fieldanalysis.INPUT_FILE = filepath
    # alias run → run_server
    dash_fieldanalysis.app.run = dash_fieldanalysis.app.run_server
    dash_fieldanalysis.app.run(debug=True, port=8051)

def _start_dqe(filepath):
    import dash_dqe
    dash_dqe.INPUT_FILE = filepath
    dash_dqe.app.run = dash_dqe.app.run_server
    dash_dqe.app.run(debug=True, port=8052)

def _start_9avar(filepath):
    import dash_9avar
    dash_9avar.INPUT_FILE = filepath
    dash_9avar.app.run = dash_9avar.app.run_server
    dash_9avar.app.run(debug=True, port=8053)


# ────────────────────────────────────────────────────────────────
# 2.  Build our “launcher” Dash app
# ────────────────────────────────────────────────────────────────

BASE_PATH = os.getcwd()
# look for subfolders
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

    html.Button("Start", id="start-btn", n_clicks=0, className="btn btn-primary"),
    html.Div(id="start-output", className="mt-3")
], className="container mt-5")


# ────────────────────────────────────────────────────────────────
# 3.  When folder or analysis changes, list all .xlsx in that folder
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
# 4.  Spawn the selected dashboard in a new process
# ────────────────────────────────────────────────────────────────
@launcher.callback(
    Output("start-output", "children"),
    Input("start-btn", "n_clicks"),
    State("folder-dropdown", "value"),
    State("analysis-dropdown", "value"),
    State("file-dropdown", "value"),
    prevent_initial_call=True
)
def _launch(n_clicks, folder, analysis, filepath):
    if not (folder and analysis and filepath):
        return html.Div(
            "⚠️ Please select a folder, analysis type, and Excel file.",
            style={"color": "red"}
        )

    if analysis == "field":
        proc = multiprocessing.Process(target=_start_field, args=(filepath,))
        link = "http://127.0.0.1:8051"
        label = "Field Analysis"
    elif analysis == "dqe":
        proc = multiprocessing.Process(target=_start_dqe, args=(filepath,))
        link = "http://127.0.0.1:8052"
        label = "DQE"
    else:  # "9a_var"
        proc = multiprocessing.Process(target=_start_9avar, args=(filepath,))
        link = "http://127.0.0.1:8053"
        label = "9a_Var"

    proc.daemon = True
    proc.start()

    return html.Div([
        f"🚀 Launched {label} on port {link.split(':')[-1]}. ",
        html.A("Click here to open", href=link, target="_blank")
    ])


# ────────────────────────────────────────────────────────────────
# 5.  Run the launcher itself (on port 8050)
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # alias .run → .run_server so we can call .run(...)
    launcher.run = launcher.run_server
    launcher.run(debug=True, port=8050)