import pandas as pd
import datetime
import re

from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output, State

# ---------------------------
# Constants & Input/Output
# ---------------------------
INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "output.xlsx"

# ---------------------------
# Step 1: Read the Excel File
# ---------------------------
df_data = pd.read_excel(INPUT_FILE, sheet_name="Data")
df_control = pd.read_excel(INPUT_FILE, sheet_name="Control")  # if you use it elsewhere

# Parse the date column
df_data['filemonth_dt'] = pd.to_datetime(df_data['filemonth_dt'], format='%m/%d/%Y')

# ---------------------------
# Step 2: Define the Two Dates
# ---------------------------
date1 = datetime.datetime(2025, 1, 1)
date2 = datetime.datetime(2024, 12, 1)

# ---------------------------
# Step 3: Unique Fields & Regex Phrases
# ---------------------------
fields = sorted(df_data['field_name'].unique())
phrases = [
    r"1\)\s*F6CF Loan - Both Pop, Diff Values",
    r"2\)\s*CF Loan - Prior Null, Current Pop",
    r"3\)\s*CF Loan - Prior Pop, Current Null"
]

def contains_phrase(text):
    text = str(text)
    return any(re.search(p, text) for p in phrases)

# ---------------------------
# Step 4: Compute Summary Data
# ---------------------------
summary_data = []
for field in fields:
    # Missing values sums
    m1 = df_data[
        (df_data['analysis_type']=='value_dist') &
        (df_data['field_name']==field) &
        (df_data['filemonth_dt']==date1)
    ]
    missing_sum_date1 = m1[m1['value_label'].str.contains("Missing", case=False, na=False)]['value_records'].sum()

    m2 = df_data[
        (df_data['analysis_type']=='value_dist') &
        (df_data['field_name']==field) &
        (df_data['filemonth_dt']==date2)
    ]
    missing_sum_date2 = m2[m2['value_label'].str.contains("Missing", case=False, na=False)]['value_records'].sum()

    # Month-to-Month diffs
    p1 = df_data[
        (df_data['analysis_type']=='pop_comp') &
        (df_data['field_name']==field) &
        (df_data['filemonth_dt']==date1)
    ]
    m2m_sum_date1 = p1[p1['value_label'].apply(contains_phrase)]['value_records'].sum()

    p2 = df_data[
        (df_data['analysis_type']=='pop_comp') &
        (df_data['field_name']==field) &
        (df_data['filemonth_dt']==date2)
    ]
    m2m_sum_date2 = p2[p2['value_label'].apply(contains_phrase)]['value_records'].sum()

    summary_data.append([
        field,
        missing_sum_date1,
        missing_sum_date2,
        m2m_sum_date1,
        m2m_sum_date2,
        ""  # Initialize the comment column with empty strings (for user input in the tables)
    ])

# Build a Pandas DataFrame for Dash
df_summary = pd.DataFrame(summary_data, columns=[
    "Field Name",
    f"Missing {date1.strftime('%m/%d/%Y')}",
    f"Missing {date2.strftime('%m/%d/%Y')}",
    f"M2M Diff {date1.strftime('%m/%d/%Y')}",
    f"M2M Diff {date2.strftime('%m/%d/%Y')}",
    "Comment",  # New column for comment storage (for user input)
])

# ---------------------------
# Step 5: Build Value-Dist & Pop-Comp DataFrames
# ---------------------------
# filter to just the two months
mask_months = df_data['filemonth_dt'].isin([date1, date2])

df_value_dist = (
    df_data[mask_months & (df_data['analysis_type']=='value_dist')]
    .copy()
)
df_value_dist['filemonth_dt'] = df_value_dist['filemonth_dt'].dt.strftime('%m/%d/%Y')

df_pop_comp = (
    df_data[mask_months & (df_data['analysis_type']=='pop_comp')]
    .loc[lambda d: d['value_label'].apply(contains_phrase)]
    .copy()
)
df_pop_comp['filemonth_dt'] = df_pop_comp['filemonth_dt'].dt.strftime('%m/%d/%Y')

# ---------------------------
# Step 6: Dash App with Three Tabs
# ---------------------------
app = Dash(__name__)

app.layout = html.Div([
    html.H2("BDCOMM FRY14M Field Analysis"),
    dcc.Tabs([
        dcc.Tab(label="Summary", children=[
            dash_table.DataTable(
                id='summary_table',
                columns=[{"name": c, "id": c} for c in df_summary.columns if c != 'Comment'],  # Remove comment column
                data=df_summary.to_dict("records"),
                page_size=20,
                style_header={'backgroundColor': '#4F81BD', 'color': 'white', 'fontWeight': 'bold'},
                style_cell={'textAlign': 'center'},
                style_table={'overflowX': 'auto'},
                selected_cells=[],  # For capturing double-clicks
                editable=False,  # Summary table is not editable
                style_data_conditional=[
                    {
                        'if': {'column_id': c, 'filter_query': '{Comment} != ""'},
                        'backgroundColor': 'yellow',
                        'color': 'black',
                    }
                    for c in ['C', 'D', 'F', 'G']  # Highlight columns with comments (Missing, M2M, etc.)
                ],
                tooltip_data=[
                    {
                        'Field Name': {'value': row['Comment'], 'type': 'markdown'} if row['Comment'] else {'value': '', 'type': 'markdown'}
                    }
                    for row in df_summary.to_dict("records")
                ]
            ),
            html.Div(id='comment_display', style={'padding': '10px', 'backgroundColor': '#f5f5f5'})
        ]),
        dcc.Tab(label="Value Distribution", children=[
            html.Div(id='value_sql_logic_box', style={'padding': '10px', 'backgroundColor': '#f5f5f5'}),
            dash_table.DataTable(
                id='value_dist_table',
                columns=[{"name": c, "id": c} for c in df_value_dist.columns if c != 'value_sql_logic'],  # Remove the value_sql_logic column
                data=df_value_dist.to_dict("records"),
                page_size=20,
                style_header={'fontWeight': 'bold', 'backgroundColor': '#4F81BD', 'color': 'white'},
                style_cell={'textAlign': 'left'},
                style_table={'overflowX': 'auto'},
                editable=True,  # Enable comment editing
                row_deletable=True,
                selected_rows=[]  # To track selected rows for filtering
            ),
            html.Div([
                dcc.Input(id='value_dist_comment', type='text', placeholder='Enter comment here'),
                html.Button('Submit Comment', id='submit_value_dist_comment', n_clicks=0)
            ])
        ]),
        dcc.Tab(label="Population Comparison", children=[
            html.Div(id='pop_sql_logic_box', style={'padding': '10px', 'backgroundColor': '#f5f5f5'}),
            dash_table.DataTable(
                id='pop_comp_table',
                columns=[{"name": c, "id": c} for c in df_pop_comp.columns if c != 'value_sql_logic'],  # Remove the value_sql_logic column
                data=df_pop_comp.to_dict("records"),
                page_size=20,
                style_header={'fontWeight': 'bold', 'backgroundColor': '#4F81BD', 'color': 'white'},
                style_cell={'textAlign': 'left'},
                style_table={'overflowX': 'auto'},
                editable=True,  # Enable comment editing
                row_deletable=True,
                selected_rows=[]  # To track selected rows for filtering
            ),
            html.Div([
                dcc.Input(id='pop_comp_comment', type='text', placeholder='Enter comment here'),
                html.Button('Submit Comment', id='submit_pop_comp_comment', n_clicks=0)
            ])
        ]),
    ])
])


# Callback for handling double-click and switching to correct tab (Value Distribution / Population Comparison)
@app.callback(
    [Output('value_dist_table', 'data'),
     Output('pop_comp_table', 'data')],
    [Input('summary_table', 'selected_cells')]
)
def update_tables(selected_cells):
    if not selected_cells:
        return [df_value_dist.to_dict("records"), df_pop_comp.to_dict("records")]

    field_name = df_summary.iloc[selected_cells[0]['row']]["Field Name"]

    # Filter the tables to show only the selected field_name
    filtered_value_dist = df_value_dist[df_value_dist['field_name'] == field_name]
    filtered_pop_comp = df_pop_comp[df_pop_comp['field_name'] == field_name]

    return [
        filtered_value_dist.to_dict("records"),
        filtered_pop_comp.to_dict("records")
    ]


# Callback to handle comment input in Value Distribution
@app.callback(
    Output('summary_table', 'data'),
    [Input('submit_value_dist_comment', 'n_clicks')],
    [State('value_dist_comment', 'value'),
     State('value_dist_table', 'data')]
)
def submit_value_dist_comment(n_clicks, comment, table_data):
    if n_clicks > 0 and comment:
        selected_field = table_data[0]['field_name']
        df_summary.loc[df_summary['Field Name'] == selected_field, 'Comment'] = comment
        return df_summary.to_dict("records")
    return dash.no_update


# Callback to handle comment input in Population Comparison
@app.callback(
    Output('summary_table', 'data'),
    [Input('submit_pop_comp_comment', 'n_clicks')],
    [State('pop_comp_comment', 'value'),
     State('pop_comp_table', 'data')]
)
def submit_pop_comp_comment(n_clicks, comment, table_data):
    if n_clicks > 0 and comment:
        selected_field = table_data[0]['field_name']
        df_summary.loc[df_summary['Field Name'] == selected_field, 'Comment'] = comment
        return df_summary.to_dict("records")
    return dash.no_update


# Run the app
if __name__ == "__main__":
    app.run(debug=True)