"""
dash_code.py
============
Same functionality as before, plus:

• Column-level filter boxes in Summary, Value Distribution
  and Population Comparison tables (DataTable’s native UI).

Python ≥3.8 · Dash ≥2.6 · Pandas ≥1.3
"""

import datetime as dt
import re
import pandas as pd
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output, State
import dash

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
contains_phrase = lambda x: any(re.search(p, str(x)) for p in PHRASES)

# ────────────────────────────────────────────────────────────────
# 4.  Build Summary dataframe (with comment columns)
# ────────────────────────────────────────────────────────────────
summary_rows = []
for fld in sorted(df_data["field_name"].unique()):
    miss1 = df_data[(df_data["analysis_type"]=="value_dist")&(df_data["field_name"]==fld)&
                    (df_data["filemonth_dt"]==DATE1)&
                    (df_data["value_label"].str.contains("Missing",case=False,na=False))]["value_records"].sum()
    miss2 = df_data[(df_data["analysis_type"]=="value_dist")&(df_data["field_name"]==fld)&
                    (df_data["filemonth_dt"]==DATE2)&
                    (df_data["value_label"].str.contains("Missing",case=False,na=False))]["value_records"].sum()
    diff1 = df_data[(df_data["analysis_type"]=="pop_comp")&(df_data["field_name"]==fld)&
                    (df_data["filemonth_dt"]==DATE1)&
                    (df_data["value_label"].apply(contains_phrase))]["value_records"].sum()
    diff2 = df_data[(df_data["analysis_type"]=="pop_comp")&(df_data["field_name"]==fld)&
                    (df_data["filemonth_dt"]==DATE2)&
                    (df_data["value_label"].apply(contains_phrase))]["value_records"].sum()
    summary_rows.append([fld, miss1, miss2, "", diff1, diff2, ""])

df_summary = pd.DataFrame(summary_rows, columns=[
    "Field Name",
    f"Missing {DATE1:%m/%d/%Y}",
    f"Missing {DATE2:%m/%d/%Y}",
    "Comment Missing",
    f"M2M Diff {DATE1:%m/%d/%Y}",
    f"M2M Diff {DATE2:%m/%d/%Y}",
    "Comment M2M",
])

# ────────────────────────────────────────────────────────────────
# 5.  Detail tables (just the two months)
# ────────────────────────────────────────────────────────────────
mask_2m = df_data["filemonth_dt"].isin([DATE1, DATE2])

vd_all = df_data[mask_2m & (df_data["analysis_type"]=="value_dist")].copy()
pc_all = df_data[mask_2m & (df_data["analysis_type"]=="pop_comp")
                 ].loc[lambda d: d["value_label"].apply(contains_phrase)].copy()

vd_all["filemonth_dt"] = vd_all["filemonth_dt"].dt.strftime("%m/%d/%Y")
pc_all["filemonth_dt"] = pc_all["filemonth_dt"].dt.strftime("%m/%d/%Y")

# ────────────────────────────────────────────────────────────────
# 6.  Dash layout
# ────────────────────────────────────────────────────────────────
def comment_area(text_id, btn_id):
    return html.Div([
        dcc.Textarea(id=text_id, placeholder="Add comment…",
                     style={"width":"100%","height":"60px"}),
        html.Button("Submit", id=btn_id, n_clicks=0,
                    style={"marginTop":"0.25rem"})
    ], style={"marginTop":"0.5rem"})

def summary_cols():
    return [{"name":c,"id":c,"editable":c.startswith("Comment")}
            for c in df_summary.columns]

app = Dash(__name__)
app.layout = html.Div([
    html.H2("BDCOMM FRY14M Field Analysis"),
    dcc.Tabs([
        # ── Summary ──────────────────────────────────────────
        dcc.Tab(label="Summary", children=[
            dash_table.DataTable(
                id="summary",
                columns=summary_cols(),
                data=df_summary.to_dict("records"),
                page_size=20,
                editable=True,
                filter_action="native",       #  ←─  filter box
                sort_action="native",
                style_header={"backgroundColor":"#4F81BD","color":"white",
                              "fontWeight":"bold"},
                style_table={"overflowX":"auto"},
            )
        ]),
        # ── Value Distribution ──────────────────────────────
        dcc.Tab(label="Value Distribution", children=[
            dash_table.DataTable(
                id="vd",
                columns=[{"name":c,"id":c} for c in vd_all.columns
                         if c!="value_sql_logic"],
                data=vd_all.to_dict("records"),
                page_size=20,
                filter_action="native",       #  ←─  filter box
                sort_action="native",
                style_header={"backgroundColor":"#4F81BD","color":"white",
                              "fontWeight":"bold"},
                style_table={"overflowX":"auto"},
            ),
            comment_area("vd_comm_text","vd_comm_btn"),
            html.Pre(id="vd_sql",style={"whiteSpace":"pre-wrap",
                                        "backgroundColor":"#f3f3f3",
                                        "padding":"0.75rem",
                                        "border":"1px solid #ddd",
                                        "marginTop":"0.5rem",
                                        "fontFamily":"monospace",
                                        "fontSize":"0.9rem"})
        ]),
        # ── Population Comparison ───────────────────────────
        dcc.Tab(label="Population Comparison", children=[
            dash_table.DataTable(
                id="pc",
                columns=[{"name":c,"id":c} for c in pc_all.columns
                         if c!="value_sql_logic"],
                data=pc_all.to_dict("records"),
                page_size=20,
                filter_action="native",       #  ←─  filter box
                sort_action="native",
                style_header={"backgroundColor":"#4F81BD","color":"white",
                              "fontWeight":"bold"},
                style_table={"overflowX":"auto"},
            ),
            comment_area("pc_comm_text","pc_comm_btn"),
            html.Pre(id="pc_sql",style={"whiteSpace":"pre-wrap",
                                        "backgroundColor":"#f3f3f3",
                                        "padding":"0.75rem",
                                        "border":"1px solid #ddd",
                                        "marginTop":"0.5rem",
                                        "fontFamily":"monospace",
                                        "fontSize":"0.9rem"})
        ]),
    ])
])

# ────────────────────────────────────────────────────────────────
# 7.  Master callback  (unchanged logic, filters unaffected)
# ────────────────────────────────────────────────────────────────
@app.callback(
    Output("vd","data"), Output("pc","data"),
    Output("vd_sql","children"), Output("pc_sql","children"),
    Output("summary","data"),
    Input("summary","active_cell"),
    Input("vd_comm_btn","n_clicks"), Input("pc_comm_btn","n_clicks"),
    State("summary","data"), State("vd","data"), State("pc","data"),
    State("vd_comm_text","value"), State("pc_comm_text","value"),
    prevent_initial_call=True
)
def master_cb(active, n_vd, n_pc,
              s_rows, vd_rows, pc_rows, vd_txt, pc_txt):

    s_df = pd.DataFrame(s_rows)
    trig  = dash.callback_context.triggered[0]["prop_id"].split(".")[0]

    # ---- append comments ---------------------------------------------------
    if trig=="vd_comm_btn" and vd_txt:
        field = vd_rows[0]["field_name"] if vd_rows else None
        if field:
            m = s_df["Field Name"]==field
            old = s_df.loc[m,"Comment Missing"].iloc[0]
            s_df.loc[m,"Comment Missing"] = (old+"\n" if old else "")+vd_txt
    if trig=="pc_comm_btn" and pc_txt:
        field = pc_rows[0]["field_name"] if pc_rows else None
        if field:
            m = s_df["Field Name"]==field
            old = s_df.loc[m,"Comment M2M"].iloc[0]
            s_df.loc[m,"Comment M2M"] = (old+"\n" if old else "")+pc_txt

    # ---- which field is active? -------------------------------------------
    if trig=="summary" and active:
        field_active = s_df.iloc[active["row"]]["Field Name"]
    else:
        field_active = vd_rows[0]["field_name"] if vd_rows else None

    if field_active:
        vd_filt = vd_all[vd_all["field_name"]==field_active]
        pc_filt = pc_all[pc_all["field_name"]==field_active]
        vd_sql  = vd_filt["value_sql_logic"].iloc[0] \
                  if "value_sql_logic" in vd_filt else ""
        pc_sql  = pc_filt["value_sql_logic"].iloc[0] \
                  if "value_sql_logic" in pc_filt else ""
    else:
        vd_filt, pc_filt, vd_sql, pc_sql = vd_all, pc_all, "", ""

    return (vd_filt.to_dict("records"), pc_filt.to_dict("records"),
            vd_sql, pc_sql, s_df.to_dict("records"))

# ────────────────────────────────────────────────────────────────
# 8.  Run
# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True)