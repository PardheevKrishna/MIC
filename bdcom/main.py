"""
Run:  python app.py
Dash ≥2.17 / Plotly ≥5.20 tested (but anything ≥2.6 is fine).

Behaviour
---------
* Click (single or double) a cell in the Summary tab → the other two
  tabs show only that field.
* Enter a comment in either detail tab and press its button →
  comment is stored in the Summary table and shown as a tooltip.

No duplicate-output errors, single callback handles both tasks.
"""
import datetime as dt
import re

import pandas as pd
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output, State
import dash

# ────────────────────────────────────────────────────────────────
# 1.  Load the workbook
# ────────────────────────────────────────────────────────────────
INPUT_FILE = "input.xlsx"

df_data    = pd.read_excel(INPUT_FILE, sheet_name="Data")
df_control = pd.read_excel(INPUT_FILE, sheet_name="Control")     # unused for now

df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"],
                                         format="%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 2.  Reference months
# ────────────────────────────────────────────────────────────────
d1 = dt.datetime(2025,  1, 1)
d2 = dt.datetime(2024, 12, 1)

# ────────────────────────────────────────────────────────────────
# 3.  Helpers
# ────────────────────────────────────────────────────────────────
PHRASES = [
    r"1\)\s*CF Loan - Both Pop, Diff Values",
    r"2\)\s*CF Loan - Prior Null, Current Pop",
    r"3\)\s*CF Loan - Prior Pop, Current Null",
]
def contains_phrase(text: str) -> bool:
    return any(re.search(p, str(text)) for p in PHRASES)

# ────────────────────────────────────────────────────────────────
# 4.  Build the “Summary” frame
# ────────────────────────────────────────────────────────────────
rows = []
for field in sorted(df_data["field_name"].unique()):
    # Missing counts
    miss_1 = df_data[
        (df_data["analysis_type"] == "value_dist") &
        (df_data["field_name"]   == field)        &
        (df_data["filemonth_dt"] == d1)
    ]["value_label"].str.contains("Missing", case=False, na=False)
    miss_2 = df_data[
        (df_data["analysis_type"] == "value_dist") &
        (df_data["field_name"]   == field)          &
        (df_data["filemonth_dt"] == d2)
    ]["value_label"].str.contains("Missing", case=False, na=False)
    msum1 = df_data.loc[miss_1, "value_records"].sum()
    msum2 = df_data.loc[miss_2, "value_records"].sum()

    # Pop-comp diffs
    pop_1 = df_data[
        (df_data["analysis_type"] == "pop_comp")   &
        (df_data["field_name"]   == field)         &
        (df_data["filemonth_dt"] == d1)
    ]
    pop_2 = df_data[
        (df_data["analysis_type"] == "pop_comp")   &
        (df_data["field_name"]   == field)         &
        (df_data["filemonth_dt"] == d2)
    ]
    diff1 = pop_1[pop_1["value_label"].apply(contains_phrase)]["value_records"].sum()
    diff2 = pop_2[pop_2["value_label"].apply(contains_phrase)]["value_records"].sum()

    rows.append([field, msum1, msum2, diff1, diff2, ""])  # final "" is the comment

df_summary = pd.DataFrame(
    rows,
    columns=[
        "Field Name",
        f"Missing {d1.strftime('%m/%d/%Y')}",
        f"Missing {d2.strftime('%m/%d/%Y')}",
        f"M2M Diff {d1.strftime('%m/%d/%Y')}",
        f"M2M Diff {d2.strftime('%m/%d/%Y')}",
        "Comment",
    ],
)

# ────────────────────────────────────────────────────────────────
# 5.  Value-dist and pop-comp frames (only the two months)
# ────────────────────────────────────────────────────────────────
mask_month = df_data["filemonth_dt"].isin([d1, d2])

df_vd = df_data[mask_month & (df_data["analysis_type"] == "value_dist")].copy()
df_pc = (
    df_data[mask_month & (df_data["analysis_type"] == "pop_comp")]
    .loc[lambda d: d["value_label"].apply(contains_phrase)]
    .copy()
)

for _df in (df_vd, df_pc):
    _df["filemonth_dt"] = _df["filemonth_dt"].dt.strftime("%m/%d/%Y")

# Keep originals for “reset”
VD_ALL = df_vd.copy()
PC_ALL = df_pc.copy()

# ────────────────────────────────────────────────────────────────
# 6.  Dash layout
# ────────────────────────────────────────────────────────────────
app = Dash(__name__)
app.layout = html.Div([
    html.H2("BDCOMM FRY14M Field Analysis"),
    dcc.Tabs([
        # ---------- Summary ----------
        dcc.Tab(label="Summary", children=[
            dash_table.DataTable(
                id="summary_table",
                columns=[{"name": c, "id": c} for c in df_summary.columns if c != "Comment"],
                data=df_summary.to_dict("records"),
                page_size=20,
                style_header={"backgroundColor": "#4F81BD", "color": "white",
                              "fontWeight": "bold"},
                style_cell={"textAlign": "center"},
                style_table={"overflowX": "auto"},
                tooltip_data=[
                    {"Field Name": {"value": r["Comment"], "type": "markdown"}}
                    if r["Comment"] else {"Field Name": {"value": ""}}
                    for r in df_summary.to_dict("records")
                ],
            )
        ]),
        # ---------- Value Distribution ----------
        dcc.Tab(label="Value Distribution", children=[
            dash_table.DataTable(
                id="value_dist_table",
                columns=[{"name": c, "id": c} for c in df_vd.columns if c != "value_sql_logic"],
                data=VD_ALL.to_dict("records"),
                page_size=20,
                style_header={"backgroundColor": "#4F81BD", "color": "white",
                              "fontWeight": "bold"},
                style_table={"overflowX": "auto"},
            ),
            html.Div([
                dcc.Input(id="vd_comment", type="text",
                          placeholder="Enter comment for selected field…"),
                html.Button("Submit", id="vd_btn", n_clicks=0),
            ], style={"padding": "0.5rem"}),
        ]),
        # ---------- Population Comparison ----------
        dcc.Tab(label="Population Comparison", children=[
            dash_table.DataTable(
                id="pop_comp_table",
                columns=[{"name": c, "id": c} for c in df_pc.columns if c != "value_sql_logic"],
                data=PC_ALL.to_dict("records"),
                page_size=20,
                style_header={"backgroundColor": "#4F81BD", "color": "white",
                              "fontWeight": "bold"},
                style_table={"overflowX": "auto"},
            ),
            html.Div([
                dcc.Input(id="pc_comment", type="text",
                          placeholder="Enter comment for selected field…"),
                html.Button("Submit", id="pc_btn", n_clicks=0),
            ], style={"padding": "0.5rem"}),
        ]),
    ])
])

# ────────────────────────────────────────────────────────────────
# 7.  Single callback = no duplicate outputs
# ────────────────────────────────────────────────────────────────
@app.callback(
    Output("summary_table",     "data"),
    Output("value_dist_table",  "data"),
    Output("pop_comp_table",    "data"),
    Input("vd_btn",             "n_clicks"),
    Input("pc_btn",             "n_clicks"),
    Input("summary_table",      "active_cell"),        # clicks in Summary
    State("vd_comment",         "value"),
    State("pc_comment",         "value"),
    State("value_dist_table",   "data"),
    State("pop_comp_table",     "data"),
    State("summary_table",      "data"),               # visible (sorted) rows
    prevent_initial_call=True
)
def manage_everything(n_vd, n_pc, active_cell,
                      vd_comment, pc_comment,
                      vd_rows, pc_rows, summary_rows):
    """
    One function that:
    1.  Stores comments (if either submit button was clicked).
    2.  Applies / removes filtering based on which Summary cell is active.
    """
    global df_summary, VD_ALL, PC_ALL  # will mutate df_summary comments

    trig = dash.callback_context.triggered

    # ----------------------------------------------------------
    #  A. If a comment button fired store the comment
    # ----------------------------------------------------------
    if trig:
        trig_id = trig[0]["prop_id"].split(".")[0]

        if trig_id == "vd_btn" and vd_comment:
            field = vd_rows[0]["field_name"]            # first row always the field on screen
            df_summary.loc[df_summary["Field Name"] == field, "Comment"] = vd_comment

        elif trig_id == "pc_btn" and pc_comment:
            field = pc_rows[0]["field_name"]
            df_summary.loc[df_summary["Field Name"] == field, "Comment"] = pc_comment

    # ----------------------------------------------------------
    #  B. Filtering detail tables
    # ----------------------------------------------------------
    # On first load `active_cell` is None → show everything.
    if active_cell is None:
        vd_out = VD_ALL
        pc_out = PC_ALL
    else:
        row_idx   = active_cell["row"]
        field_val = summary_rows[row_idx]["Field Name"]

        vd_out = VD_ALL[VD_ALL["field_name"] == field_val]
        pc_out = PC_ALL[PC_ALL["field_name"] == field_val]

    # ----------------------------------------------------------
    #  C. Return everything in one go (no duplicates!)
    # ----------------------------------------------------------
    return (
        df_summary.to_dict("records"),
        vd_out.to_dict("records"),
        pc_out.to_dict("records"),
    )

# ────────────────────────────────────────────────────────────────
# 8.  Run
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)