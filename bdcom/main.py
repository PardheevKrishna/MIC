"""
Run with:  python app.py
Dash ≥2.6 required (for active_cell).  Tested with Dash 2.17 / Plotly 5.20.

Assumes `input.xlsx` has the sheets “Data” and “Control”.
"""
import datetime
import re

import pandas as pd
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output, State

# ---------------------------
# Constants & Input/Output
# ---------------------------
INPUT_FILE  = "input.xlsx"
OUTPUT_FILE = "output.xlsx"       #  still unused but kept for future work

# ---------------------------
# Step 1: Read the Excel File
# ---------------------------
df_data    = pd.read_excel(INPUT_FILE, sheet_name="Data")
df_control = pd.read_excel(INPUT_FILE, sheet_name="Control")  # if you need it later

# Parse the date column once
df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"], format="%m/%d/%Y")

# ---------------------------
# Step 2: Define the Two Dates
# ---------------------------
date1 = datetime.datetime(2025, 1, 1)
date2 = datetime.datetime(2024, 12, 1)

# ---------------------------
# Step 3: Unique Fields & Regex Phrases
# ---------------------------
fields  = sorted(df_data["field_name"].unique())
phrases = [
    r"1\)\s*CF Loan - Both Pop, Diff Values",
    r"2\)\s*CF Loan - Prior Null, Current Pop",
    r"3\)\s*CF Loan - Prior Pop, Current Null",
]

def contains_phrase(text: str) -> bool:
    text = str(text)
    return any(re.search(p, text) for p in phrases)

# ---------------------------
# Step 4: Compute Summary Data
# ---------------------------
summary_rows = []
for field in fields:
    # 1️⃣ Missing values
    m1 = df_data[
        (df_data["analysis_type"] == "value_dist")
        & (df_data["field_name"] == field)
        & (df_data["filemonth_dt"] == date1)
    ]
    missing_sum_date1 = m1[m1["value_label"].str.contains("Missing", case=False, na=False)][
        "value_records"
    ].sum()

    m2 = df_data[
        (df_data["analysis_type"] == "value_dist")
        & (df_data["field_name"] == field)
        & (df_data["filemonth_dt"] == date2)
    ]
    missing_sum_date2 = m2[m2["value_label"].str.contains("Missing", case=False, na=False)][
        "value_records"
    ].sum()

    # 2️⃣ Month-to-month population diffs
    p1 = df_data[
        (df_data["analysis_type"] == "pop_comp")
        & (df_data["field_name"] == field)
        & (df_data["filemonth_dt"] == date1)
    ]
    m2m_sum_date1 = p1[p1["value_label"].apply(contains_phrase)]["value_records"].sum()

    p2 = df_data[
        (df_data["analysis_type"] == "pop_comp")
        & (df_data["field_name"] == field)
        & (df_data["filemonth_dt"] == date2)
    ]
    m2m_sum_date2 = p2[p2["value_label"].apply(contains_phrase)]["value_records"].sum()

    summary_rows.append(
        [
            field,
            missing_sum_date1,
            missing_sum_date2,
            m2m_sum_date1,
            m2m_sum_date2,
            "",  # comment column – editable via other callbacks
        ]
    )

df_summary = pd.DataFrame(
    summary_rows,
    columns=[
        "Field Name",
        f"Missing {date1.strftime('%m/%d/%Y')}",
        f"Missing {date2.strftime('%m/%d/%Y')}",
        f"M2M Diff {date1.strftime('%m/%d/%Y')}",
        f"M2M Diff {date2.strftime('%m/%d/%Y')}",
        "Comment",
    ],
)

# ---------------------------
# Step 5: Build Value-Dist & Pop-Comp DataFrames
# ---------------------------
mask_months = df_data["filemonth_dt"].isin([date1, date2])

df_value_dist = df_data[mask_months & (df_data["analysis_type"] == "value_dist")].copy()
df_value_dist["filemonth_dt"] = df_value_dist["filemonth_dt"].dt.strftime("%m/%d/%Y")

df_pop_comp = (
    df_data[mask_months & (df_data["analysis_type"] == "pop_comp")]
    .loc[lambda d: d["value_label"].apply(contains_phrase)]
    .copy()
)
df_pop_comp["filemonth_dt"] = df_pop_comp["filemonth_dt"].dt.strftime("%m/%d/%Y")

# ---------------------------
# Step 6: Dash App with Three Tabs
# ---------------------------
app = Dash(__name__)

app.layout = html.Div(
    [
        html.H2("BDCOMM FRY14M Field Analysis"),
        dcc.Tabs(
            [
                # ────────────── Summary ──────────────
                dcc.Tab(
                    label="Summary",
                    children=[
                        dash_table.DataTable(
                            id="summary_table",
                            columns=[
                                {"name": c, "id": c}
                                for c in df_summary.columns
                                if c != "Comment"
                            ],
                            data=df_summary.to_dict("records"),
                            page_size=20,
                            style_header={
                                "backgroundColor": "#4F81BD",
                                "color": "white",
                                "fontWeight": "bold",
                            },
                            style_cell={"textAlign": "center"},
                            style_table={"overflowX": "auto"},
                            # 👉  we *don't* need `selected_cells`
                            editable=False,
                            tooltip_data=[
                                {
                                    "Field Name": {
                                        "value": row["Comment"],
                                        "type": "markdown",
                                    }
                                    if row["Comment"]
                                    else {"value": "", "type": "markdown"}
                                }
                                for row in df_summary.to_dict("records")
                            ],
                        )
                    ],
                ),
                # ───────────── Value Distribution ─────────────
                dcc.Tab(
                    label="Value Distribution",
                    children=[
                        html.Div(
                            id="value_sql_logic_box",
                            style={"padding": "10px", "backgroundColor": "#f5f5f5"},
                        ),
                        dash_table.DataTable(
                            id="value_dist_table",
                            columns=[
                                {"name": c, "id": c}
                                for c in df_value_dist.columns
                                if c != "value_sql_logic"
                            ],
                            data=df_value_dist.to_dict("records"),
                            page_size=20,
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#4F81BD",
                                "color": "white",
                            },
                            style_cell={"textAlign": "left"},
                            style_table={"overflowX": "auto"},
                            editable=True,
                            row_deletable=True,
                        ),
                        html.Div(
                            [
                                dcc.Input(
                                    id="value_dist_comment",
                                    type="text",
                                    placeholder="Enter comment here",
                                ),
                                html.Button(
                                    "Submit Comment",
                                    id="submit_value_dist_comment",
                                    n_clicks=0,
                                ),
                            ]
                        ),
                        html.Div(id="value_sql_logic", children="SQL logic goes here"),
                    ],
                ),
                # ───────────── Population Comparison ──────────
                dcc.Tab(
                    label="Population Comparison",
                    children=[
                        html.Div(
                            id="pop_sql_logic_box",
                            style={"padding": "10px", "backgroundColor": "#f5f5f5"},
                        ),
                        dash_table.DataTable(
                            id="pop_comp_table",
                            columns=[
                                {"name": c, "id": c}
                                for c in df_pop_comp.columns
                                if c != "value_sql_logic"
                            ],
                            data=df_pop_comp.to_dict("records"),
                            page_size=20,
                            style_header={
                                "fontWeight": "bold",
                                "backgroundColor": "#4F81BD",
                                "color": "white",
                            },
                            style_cell={"textAlign": "left"},
                            style_table={"overflowX": "auto"},
                            editable=True,
                            row_deletable=True,
                        ),
                        html.Div(
                            [
                                dcc.Input(
                                    id="pop_comp_comment",
                                    type="text",
                                    placeholder="Enter comment here",
                                ),
                                html.Button(
                                    "Submit Comment",
                                    id="submit_pop_comp_comment",
                                    n_clicks=0,
                                ),
                            ]
                        ),
                        html.Div(id="pop_sql_logic", children="SQL logic goes here"),
                    ],
                ),
            ]
        ),
    ]
)

# ───────────────────────────────────────────────────────────
#  Callbacks
# ───────────────────────────────────────────────────────────
@app.callback(
    Output("summary_table", "data"),
    Output("value_dist_table", "data"),
    Output("pop_comp_table", "data"),
    Input("submit_value_dist_comment", "n_clicks"),
    Input("submit_pop_comp_comment", "n_clicks"),
    State("value_dist_comment", "value"),
    State("pop_comp_comment", "value"),
    State("value_dist_table", "data"),
    State("pop_comp_table", "data"),
)
def update_comments(
    n_clicks_vd,
    n_clicks_pc,
    comment_vd,
    comment_pc,
    value_dist_data,
    pop_comp_data,
):
    """Store user comments in the Summary table."""
    global df_summary  # mutate in place so the other callback sees it

    # Value-distribution comment
    if n_clicks_vd and comment_vd:
        selected_field = value_dist_data[0]["field_name"]
        df_summary.loc[df_summary["Field Name"] == selected_field, "Comment"] = comment_vd

    # Population-comparison comment
    if n_clicks_pc and comment_pc:
        selected_field = pop_comp_data[0]["field_name"]
        df_summary.loc[df_summary["Field Name"] == selected_field, "Comment"] = comment_pc

    return (
        df_summary.to_dict("records"),
        value_dist_data,
        pop_comp_data,
    )

@app.callback(
    Output("value_dist_table", "data"),
    Output("pop_comp_table", "data"),
    Input("summary_table", "active_cell"),      # 🔑  cell click event
    State("summary_table", "data"),             # rows after sort/filter
)
def filter_detail_tables(active_cell, summary_rows):
    """
    When a cell is clicked in the Summary table, isolate that field name
    in the two detail tables. Click elsewhere to change the filter;
    `active_cell` is None on initial page load.
    """
    if active_cell is None:
        return (
            df_value_dist.to_dict("records"),
            df_pop_comp.to_dict("records"),
        )

    row_idx   = active_cell["row"]             # row within current sort
    field_val = summary_rows[row_idx]["Field Name"]

    vd_filtered = df_value_dist[df_value_dist["field_name"] == field_val]
    pc_filtered = df_pop_comp[df_pop_comp["field_name"] == field_val]

    return (
        vd_filtered.to_dict("records"),
        pc_filtered.to_dict("records"),
    )

# Run the app
if __name__ == "__main__":
    app.run(debug=True)