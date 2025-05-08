"""
dash_code.py
============
Single-click a cell in **Summary** to

1. Filter the rows in **Value Distribution** and **Population Comparison**.
2. Show the field’s `value_sql_logic` under each table.
3. Let you write a “Missing” comment (Value Distribution tab) or
   a “M2M” comment (Population Comparison tab).  
   The comment is **appended** to the corresponding column in Summary.
4. You can also edit either comment column directly in Summary.

Python ≥3.8 · Dash ≥2.6 · Pandas ≥1.3
"""

import datetime as dt
import re
import pandas as pd
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output, State
import dash  # only to read callback_context

# ────────────────────────────────────────────────────────────────
# 1.  Load workbook
# ────────────────────────────────────────────────────────────────
INPUT_FILE = "input.xlsx"
df_data    = pd.read_excel(INPUT_FILE, sheet_name="Data")
df_data["filemonth_dt"] = pd.to_datetime(df_data["filemonth_dt"], format="%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 2.  Reference months
# ────────────────────────────────────────────────────────────────
DATE1 = dt.datetime(2025,  1, 1)
DATE2 = dt.datetime(2024, 12, 1)

# ────────────────────────────────────────────────────────────────
# 3.  Helper for pop-comp phrases
# ────────────────────────────────────────────────────────────────
PHRASES = [
    r"1\)\s*CF Loan - Both Pop, Diff Values",
    r"2\)\s*CF Loan - Prior Null, Current Pop",
    r"3\)\s*CF Loan - Prior Pop, Current Null",
]
def contains_phrase(x: str) -> bool:
    return any(re.search(p, str(x)) for p in PHRASES)

# ────────────────────────────────────────────────────────────────
# 4.  Build Summary dataframe  (incl. new comment columns)
# ────────────────────────────────────────────────────────────────
summary_rows = []
for fld in sorted(df_data["field_name"].unique()):
    miss1 = df_data[(df_data["analysis_type"] == "value_dist") &
                    (df_data["field_name"]   == fld)           &
                    (df_data["filemonth_dt"] == DATE1)         &
                    (df_data["value_label"]
                     .str.contains("Missing", case=False, na=False)
                    )]["value_records"].sum()

    miss2 = df_data[(df_data["analysis_type"] == "value_dist") &
                    (df_data["field_name"]   == fld)           &
                    (df_data["filemonth_dt"] == DATE2)         &
                    (df_data["value_label"]
                     .str.contains("Missing", case=False, na=False)
                    )]["value_records"].sum()

    diff1 = df_data[(df_data["analysis_type"] == "pop_comp")   &
                    (df_data["field_name"]   == fld)           &
                    (df_data["filemonth_dt"] == DATE1)         &
                    (df_data["value_label"].apply(contains_phrase))
                    ]["value_records"].sum()

    diff2 = df_data[(df_data["analysis_type"] == "pop_comp")   &
                    (df_data["field_name"]   == fld)           &
                    (df_data["filemonth_dt"] == DATE2)         &
                    (df_data["value_label"].apply(contains_phrase))
                    ]["value_records"].sum()

    summary_rows.append([
        fld, miss1, miss2, "",      # “Comment Missing” will go here
        diff1, diff2, "",           # “Comment M2M” will go here
    ])

df_summary = pd.DataFrame(
    summary_rows,
    columns=[
        "Field Name",
        f"Missing {DATE1:%m/%d/%Y}",
        f"Missing {DATE2:%m/%d/%Y}",
        "Comment Missing",                # NEW
        f"M2M Diff {DATE1:%m/%d/%Y}",
        f"M2M Diff {DATE2:%m/%d/%Y}",
        "Comment M2M",                    # NEW
    ],
)

# ────────────────────────────────────────────────────────────────
# 5.  Detail tables (two months only)
# ────────────────────────────────────────────────────────────────
mask_2m = df_data["filemonth_dt"].isin([DATE1, DATE2])

vd_all = df_data[mask_2m & (df_data["analysis_type"] == "value_dist")].copy()
pc_all = df_data[mask_2m & (df_data["analysis_type"] == "pop_comp")
                 ].loc[lambda d: d["value_label"].apply(contains_phrase)].copy()

vd_all["filemonth_dt"] = vd_all["filemonth_dt"].dt.strftime("%m/%d/%Y")
pc_all["filemonth_dt"] = pc_all["filemonth_dt"].dt.strftime("%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 6.  Dash layout
# ────────────────────────────────────────────────────────────────
def comment_area(txt_id, btn_id):
    """Return a Textarea + button bundle."""
    return html.Div([
        dcc.Textarea(id=txt_id,
                     placeholder="Add your comment here…",
                     style={"width": "100%", "height": "60px"}),
        html.Button("Submit", id=btn_id, n_clicks=0,
                    style={"marginTop": "0.25rem"})
    ], style={"marginTop": "0.5rem"})

def summary_columns():
    """Create column definitions, making the two comment cols editable."""
    cols = []
    for c in df_summary.columns:
        editable = c.startswith("Comment")
        cols.append({"name": c, "id": c, "editable": editable})
    return cols

app = Dash(__name__)
app.layout = html.Div([
    html.H2("BDCOMM FRY14M Field Analysis"),
    dcc.Tabs([
        # ── Summary ──────────────────────────────────────────
        dcc.Tab(label="Summary", children=[
            dash_table.DataTable(
                id="summary",
                columns=summary_columns(),
                data=df_summary.to_dict("records"),
                page_size=20,
                editable=True,                           # allow direct edits
                style_header={"backgroundColor": "#4F81BD", "color": "white",
                              "fontWeight": "bold"},
                style_table={"overflowX": "auto"},
            )
        ]),
        # ── Value Distribution ──────────────────────────────
        dcc.Tab(label="Value Distribution", children=[
            dash_table.DataTable(
                id="vd",
                columns=[{"name": c, "id": c}
                         for c in vd_all.columns if c != "value_sql_logic"],
                data=vd_all.to_dict("records"),
                page_size=20,
                style_header={"backgroundColor": "#4F81BD", "color": "white",
                              "fontWeight": "bold"},
                style_table={"overflowX": "auto"},
            ),
            comment_area("vd_comm_text", "vd_comm_btn"),
            html.Pre(id="vd_sql", style={
                "whiteSpace": "pre-wrap",
                "backgroundColor": "#f3f3f3",
                "padding": "0.75rem",
                "border": "1px solid #ddd",
                "marginTop": "0.5rem",
                "fontFamily": "monospace",
                "fontSize": "0.9rem"
            })
        ]),
        # ── Population Comparison ───────────────────────────
        dcc.Tab(label="Population Comparison", children=[
            dash_table.DataTable(
                id="pc",
                columns=[{"name": c, "id": c}
                         for c in pc_all.columns if c != "value_sql_logic"],
                data=pc_all.to_dict("records"),
                page_size=20,
                style_header={"backgroundColor": "#4F81BD", "color": "white",
                              "fontWeight": "bold"},
                style_table={"overflowX": "auto"},
            ),
            comment_area("pc_comm_text", "pc_comm_btn"),
            html.Pre(id="pc_sql", style={
                "whiteSpace": "pre-wrap",
                "backgroundColor": "#f3f3f3",
                "padding": "0.75rem",
                "border": "1px solid #ddd",
                "marginTop": "0.5rem",
                "fontFamily": "monospace",
                "fontSize": "0.9rem"
            })
        ]),
    ])
])

# ────────────────────────────────────────────────────────────────
# 7.  Single callback  (filter, SQL logic, comments)
# ────────────────────────────────────────────────────────────────
@app.callback(
    # detail-tables & SQL logic
    Output("vd", "data"),
    Output("pc", "data"),
    Output("vd_sql", "children"),
    Output("pc_sql", "children"),
    # Summary table data (to reflect comment edits / appends)
    Output("summary", "data"),

    # Triggers
    Input("summary",      "active_cell"),   # click in Summary
    Input("vd_comm_btn",  "n_clicks"),      # submit Missing comment
    Input("pc_comm_btn",  "n_clicks"),      # submit M2M comment

    # State needed to operate
    State("summary",      "data"),
    State("vd",           "data"),
    State("pc",           "data"),
    State("vd_comm_text", "value"),
    State("pc_comm_text", "value"),

    prevent_initial_call=True
)
def master_callback(active_cell,
                    n_vd, n_pc,
                    summary_rows, vd_rows, pc_rows,
                    vd_cmt, pc_cmt):
    """
    • Filters detail tables & shows SQL when Summary cell clicked.
    • Appends user comments to the proper column in Summary.
    """

    # Work with current Summary as DataFrame
    s_df = pd.DataFrame(summary_rows)

    # Helper: which prop fired?
    trig_id = dash.callback_context.triggered[0]["prop_id"].split(".")[0]

    # ── A. Handle comment submissions ───────────────────────
    if trig_id == "vd_comm_btn" and vd_cmt:
        # first row of vd table has the current field
        field = vd_rows[0]["field_name"] if vd_rows else None
        if field:
            mask = s_df["Field Name"] == field
            old = s_df.loc[mask, "Comment Missing"].iloc[0]
            new = (old + "\n" if old else "") + vd_cmt
            s_df.loc[mask, "Comment Missing"] = new

    if trig_id == "pc_comm_btn" and pc_cmt:
        # first row of pc table has the current field
        field = pc_rows[0]["field_name"] if pc_rows else None
        if field:
            mask = s_df["Field Name"] == field
            old = s_df.loc[mask, "Comment M2M"].iloc[0]
            new = (old + "\n" if old else "") + pc_cmt
            s_df.loc[mask, "Comment M2M"] = new

    # ── B. Determine which field is “active” ────────────────
    if trig_id == "summary" and active_cell is not None:
        # User just clicked -> use that field
        field_active = s_df.iloc[active_cell["row"]]["Field Name"]
    else:
        # No new click; keep whatever field is shown in vd_rows (if any)
        field_active = vd_rows[0]["field_name"] if vd_rows else None

    # ── C. Build outputs: filtered tables + SQL boxes ───────
    if field_active:
        vd_filtered = vd_all[vd_all["field_name"] == field_active]
        pc_filtered = pc_all[pc_all["field_name"] == field_active]

        vd_sql = vd_filtered["value_sql_logic"].iloc[0] \
                 if "value_sql_logic" in vd_filtered.columns else ""
        pc_sql = pc_filtered["value_sql_logic"].iloc[0] \
                 if "value_sql_logic" in pc_filtered.columns else ""
    else:
        vd_filtered, pc_filtered = vd_all, pc_all
        vd_sql = pc_sql = ""

    # ── D. Return everything in one shot  (unique outputs) ──
    return (vd_filtered.to_dict("records"),
            pc_filtered.to_dict("records"),
            vd_sql, pc_sql,
            s_df.to_dict("records"))

# ────────────────────────────────────────────────────────────────
# 8.  Run the server
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)