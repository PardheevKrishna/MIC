# dashboard_launcher.py

import os, glob
from dash import Dash, dcc, html
from dash.dependencies import Input, Output, State

# ────────────────────────────────────────────────────────────────
# 1.  Discover subfolders in cwd
# ────────────────────────────────────────────────────────────────
BASE_PATH = os.getcwd()
FOLDERS = sorted(
    d for d in os.listdir(BASE_PATH)
    if os.path.isdir(os.path.join(BASE_PATH, d))
)

# ────────────────────────────────────────────────────────────────
# 2.  Launcher app layout
# ────────────────────────────────────────────────────────────────
launcher = Dash(__name__,
    external_stylesheets=["https://cdn.jsdelivr.net/npm/bootswatch@4.5.2/dist/flatly/bootstrap.min.css"]
)

launcher.layout = html.Div([
    html.H2("📁 Select Folder → Analysis → Excel", className="text-center mb-4"),

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

        html.Div([
            html.Label("Excel file"),
            dcc.Dropdown(
                id="file-dropdown",
                placeholder="Select folder & analysis first…"
            ),
        ], className="col-md-4"),
    ], className="row mb-3"),

    html.Button("Start", id="start-btn", n_clicks=0, className="btn btn-primary"),
    html.Div(id="start-output", className="mt-3")
], className="container mt-5")


# ────────────────────────────────────────────────────────────────
# 3.  Populate file‐dropdown once folder & analysis chosen
# ────────────────────────────────────────────────────────────────
@launcher.callback(
    Output("file-dropdown","options"),
    Input("folder-dropdown","value"),
    Input("analysis-dropdown","value"),
)
def update_file_list(folder, analysis):
    if not folder or not analysis:
        return []
    path = os.path.join(BASE_PATH, folder)
    # list all .xlsx in that folder
    files = sorted(glob.glob(os.path.join(path, "*.xlsx")))
    return [{"label":os.path.basename(f), "value":f} for f in files]


# ────────────────────────────────────────────────────────────────
# 4.  Start the selected dash app when button clicked
# ────────────────────────────────────────────────────────────────
@launcher.callback(
    Output("start-output","children"),
    Input("start-btn","n_clicks"),
    State("folder-dropdown","value"),
    State("analysis-dropdown","value"),
    State("file-dropdown","value"),
    prevent_initial_call=True
)
def launch_app(n, folder, analysis, filepath):
    if not (folder and analysis and filepath):
        return html.Div("Please select Folder, Analysis & Excel file.", style={"color":"red"})

    if analysis == "field":
        import dash_fieldanalysis
        # override the input file
        dash_fieldanalysis.INPUT_FILE = filepath
        # alias run_server → run
        dash_fieldanalysis.app.run = dash_fieldanalysis.app.run_server
        dash_fieldanalysis.app.run(debug=True)

    elif analysis == "dqe":
        import dash_dqe
        dash_dqe.INPUT_FILE = filepath
        dash_dqe.app.run = dash_dqe.app.run_server
        dash_dqe.app.run(debug=True)

    else:  # "9a_var"
        import dash_9avar
        dash_9avar.INPUT_FILE = filepath
        dash_9avar.app.run = dash_9avar.app.run_server
        dash_9avar.app.run(debug=True)

    return ""  # never reached, server blocks here


# ────────────────────────────────────────────────────────────────
# 5.  Run the launcher itself
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # per your request, use app.run not run_server
    launcher.run = launcher.run_server
    launcher.run(debug=True)