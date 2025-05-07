import pandas as pd
import datetime
import re

from dash import Dash, dcc, html
from dash.dependencies import Input, Output, State
import dash_ag_grid as dag

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
            dag.AgGrid(
                id='summary_table',
                columnDefs=[
                    {'headerName': c, 'field': c} for c in df_summary.columns if c != 'Comment'
                ],
                rowData=df_summary.to_dict('records'),
                pagination=True,
                sortable=True,
                filter=True,
                enableRangeSelection=True,
            ),
        ]),
        dcc.Tab(label="Value Distribution", children=[
            dag.AgGrid(
                id='value_dist_table',
                columnDefs=[{'headerName': c, 'field': c} for c in df_value_dist.columns if c != 'value_sql_logic'],
                rowData=df_value_dist.to_dict('records'),
                pagination=True,
                sortable=True,
                filter=True,
                enableRangeSelection=True,
            ),
            html.Div([
                dcc.Input(id='value_dist_comment', type='text', placeholder='Enter comment here'),
                html.Button('Submit Comment', id='submit_value_dist_comment', n_clicks=0)
            ])
        ]),
        dcc.Tab(label="Population Comparison", children=[
            dag.AgGrid(
                id='pop_comp_table',
                columnDefs=[{'headerName': c, 'field': c} for c in df_pop_comp.columns if c != 'value_sql_logic'],
                rowData=df_pop_comp.to_dict('records'),
                pagination=True,
                sortable=True,
                filter=True,
                enableRangeSelection=True,
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
    [Output('value_dist_table', 'rowData'),
     Output('pop_comp_table', 'rowData')],
    [Input('summary_table', 'selectedRows')]
)
def update_tables(selected_rows):
    if not selected_rows:
        return [df_value_dist.to_dict('records'), df_pop_comp.to_dict('records')]

    field_name = selected_rows[0]['Field Name']

    # Filter the tables to show only the selected field_name
    filtered_value_dist = df_value_dist[df_value_dist['field_name'] == field_name]
    filtered_pop_comp = df_pop_comp[df_pop_comp['field_name'] == field_name]

    return [
        filtered_value_dist.to_dict('records'),
        filtered_pop_comp.to_dict('records')
    ]


# Callback to handle comment input in Value Distribution and Population Comparison
@app.callback(
    Output('summary_table', 'rowData'),
    [Input('submit_value_dist_comment', 'n_clicks'),
     Input('submit_pop_comp_comment', 'n_clicks')],
    [State('value_dist_comment', 'value'),
     State('pop_comp_comment', 'value'),
     State('value_dist_table', 'rowData'),
     State('pop_comp_table', 'rowData')]
)
def submit_comment(n_clicks_value_dist, n_clicks_pop_comp, comment_value_dist, comment_pop_comp, value_dist_data, pop_comp_data):
    if n_clicks_value_dist > 0 and comment_value_dist:
        selected_field = value_dist_data[0]['field_name']
        df_summary.loc[df_summary['Field Name'] == selected_field, 'Comment'] = comment_value_dist
    if n_clicks_pop_comp > 0 and comment_pop_comp:
        selected_field = pop_comp_data[0]['field_name']
        df_summary.loc[df_summary['Field Name'] == selected_field, 'Comment'] = comment_pop_comp

    return df_summary.to_dict('records')


# Run the app
if __name__ == "__main__":
    app.run(debug=True)