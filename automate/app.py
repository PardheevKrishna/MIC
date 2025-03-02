import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
from io import BytesIO
import os

# ---------- CONFIGURATION ----------
FILE_PATH = "input.xlsx"        # Path to your main Excel file
TEMP_JSON_FILE = "temp_update.json"  # Where we temporarily store user updates

# -------------- CUSTOM CSS & TITLE --------------
st.markdown("""
<style>
.mySkipButton { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.title("Team Report Dashboard")

# ===================== PROCESS FUNCTION =====================
def process_excel_file(file_path):
    """
    Reads each employee sheet (employee names from "Home") and returns two DataFrames:
      - working_details: row-level data for all employees
      - violations_df: flagged violations
    (You can keep your existing logic that checks allowed values, start date, etc.)
    """
    # For brevity, placeholders:
    try:
        # Suppose you read the 'Home' sheet, gather employee names, parse each sheet
        # We'll just return two empty DataFrames for the skeleton
        working_details = pd.DataFrame({
            "Main project": ["ProjectA", "ProjectA", "ProjectB"],
            "Month": ["2023-04", "2023-04", "2023-04"],
            "Start Date": ["04-01-2023", "04-05-2023", "04-03-2023"],
            "Completion Date": ["04-10-2023", "04-20-2023", "04-15-2023"],
            "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)": [
                "CRIT", "CRIT - Data Governance", "CRIT - Data Management"
            ],
            "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)": [
                "Data Infrastructure", "GDA Related", "Monitoring & Insights"
            ],
            "Complexity (H,M,L)": ["H", "M", "L"],
            "Novelity (BAU repetitive, One time repetitive, New one time)": [
                "BAU repetitive", "New one time", "One time repetitive"
            ],
            "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :": [
                "Core production work", "Ad-hoc short-term projects", "Business Management"
            ],
            "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)": [
                "Insights", "Risk reduction", "Customer Experience"
            ]
        })
        violations_df = pd.DataFrame()  # Placeholder
        return working_details, violations_df
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None, None

# ---- LOAD THE DATA ----
working_details, violations_df = process_excel_file(FILE_PATH)
if working_details is None:
    st.stop()
st.success("Data loaded (placeholder).")

# ========== TABS / RADIO SELECTOR ==========
tab_option = st.radio("Select a Tab", [
    "Team Monthly Summary",
    "Working Hours Summary",
    "Violations",
    "Update Data"  # <-- Our new tab
])

if tab_option == "Team Monthly Summary":
    st.subheader("Team Monthly Summary")
    st.write("Placeholder for summary logic or filtering.")
    # E.g. let user filter by Month, then show data, etc.

elif tab_option == "Working Hours Summary":
    st.subheader("Working Hours Summary")
    st.write("Placeholder for summary logic or filtering.")
    # E.g. let user filter by Month, show hours, etc.

elif tab_option == "Violations":
    st.subheader("Violations (Placeholder)")
    st.write("Placeholder for your existing violations logic (if any).")

else:
    # ========== UPDATE DATA TAB ==========
    st.subheader("Update Data")

    # 1. Let user pick filters for easier experience
    # For instance, filter by 'Main project', 'Month', etc.
    all_projects = sorted(working_details["Main project"].unique())
    all_months = sorted(working_details["Month"].unique())
    with st.form("update_data_filter_form"):
        sel_projects = st.multiselect("Select Main Project(s)", options=all_projects)
        sel_months = st.multiselect("Select Month(s)", options=all_months)
        filter_btn = st.form_submit_button("Apply Filters")

    if filter_btn:
        df_update = working_details.copy()
        if sel_projects:
            df_update = df_update[df_update["Main project"].isin(sel_projects)]
        if sel_months:
            df_update = df_update[df_update["Month"].isin(sel_months)]
        st.dataframe(df_update, use_container_width=True)

        # 2. Two modes: Automatic or Manual
        update_mode = st.radio("Select Mode", ["Automatic", "Manual"], index=0)

        if update_mode == "Automatic":
            st.markdown("**Automatic Mode**")
            st.write("For each project+month in the filtered data, we'll override Start Date with first occurrence, Completion Date with last occurrence.")
            # For each categorical column, the user picks 'First occurrence' or 'Most frequent'
            cat_columns = [
                "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                "Complexity (H,M,L)",
                "Novelity (BAU repetitive, One time repetitive, New one time)",
                "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"
            ]
            cat_choices = {}
            for col in cat_columns:
                choice = st.radio(
                    f"For {col}, choose how to override",
                    ["First occurrence within that month", "Most frequent within that month"],
                    key=f"choice_{col}"
                )
                cat_choices[col] = choice

            if st.button("Update (Automatic)"):
                # Step: for each (project, month) in df_update, override
                updated_data = []
                grouped = df_update.groupby(["Main project", "Month"])
                for (proj, mon), subdf in grouped:
                    # Start date = first occurrence
                    # e.g. parse all subdf["Start Date"] -> pick min
                    # Completion date = last occurrence
                    # Then for each cat col, either first occurrence or mode
                    # We'll store the updated row in a dictionary
                    # In real code, you might have multiple rows to override
                    # For simplicity, let's just do a placeholder
                    pass

                # Write the final updated data to text file
                with open(TEMP_JSON_FILE, "w", encoding="utf-8") as f:
                    # In real code, you'd store the updated_data dict
                    json.dump({"placeholder": "automatic updates here"}, f, indent=2)
                st.success(f"Automatic updates saved to {TEMP_JSON_FILE}. (Excel not updated in this skeleton.)")

        else:
            st.markdown("**Manual Mode**")
            st.write("User can override each field. We'll show suggestions (min Start, max Completion, etc.) but let them pick final values.")
            # We'll gather user inputs
            updated_data = []
            # For each row in df_update, compute suggestions, then show selectboxes, date inputs, etc.
            for idx, row in df_update.iterrows():
                with st.expander(f"Edit Row {idx}", expanded=False):
                    st.write(f"Project: {row['Main project']}, Month: {row['Month']}")
                    # Suggest min Start, max Completion in that project+month
                    # Actually you'd do group logic; for now just placeholders
                    suggested_start = row["Start Date"]
                    suggested_comp = row.get("Completion Date", "")
                    new_start = st.text_input("Start Date", value=suggested_start, key=f"start_{idx}")
                    new_comp = st.text_input("Completion Date", value=suggested_comp, key=f"comp_{idx}")
                    # For cat columns
                    cat_values = {}
                    for col in [
                        "Functional Area (CRIT, CRIT - Data Management, CRIT - Data Governance, CRIT - Regulatory Reporting, CRIT - Portfolio Reporting, CRIT - Transformation)",
                        "Project Category (Data Infrastructure, Monitoring & Insights, Analytics / Strategy Development, GDA Related, Trainings and Team Meeting)",
                        "Complexity (H,M,L)",
                        "Novelity (BAU repetitive, One time repetitive, New one time)",
                        "Output Type (Core production work, Ad-hoc long-term projects, Ad-hoc short-term projects, Business Management, Administration, Trainings/L&D activities, Others) :",
                        "Impact type (Customer Experience, Financial impact, Insights, Risk reduction, Others)"
                    ]:
                        cat_current = str(row.get(col, ""))
                        cat_values[col] = st.text_input(col, value=cat_current, key=f"{col}_{idx}")
                    updated_data.append({
                        "Index": idx,
                        "Project": row["Main project"],
                        "Month": row["Month"],
                        "New Start": new_start,
                        "New Comp": new_comp,
                        **cat_values
                    })

            if st.button("Update (Manual)"):
                # Save updated_data to text file
                with open(TEMP_JSON_FILE, "w", encoding="utf-8") as f:
                    json.dump(updated_data, f, indent=2)
                st.success(f"Manual updates saved to {TEMP_JSON_FILE}. (Excel not updated in this skeleton.)")

        st.info("In real code, you'd read the text file and update the Excel behind the scenes.")
    else:
        st.info("Apply filters to see data to update.")