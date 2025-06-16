#!/usr/bin/env python3
import os
import datetime
import tkinter as tk
from tkinter import messagebox
import pandas as pd
from openpyxl.utils import get_column_letter

def main():
    # ←— SET YOUR SOURCE .xlsm PATH HERE:
    src_file = r"C:\path\to\your\RegulatoryReporting.xlsm"
    
    # ←— YOUR EMPLOYEE LIST:
    employees = [
        "Alice Johnson",
        "Bob Smith",
        "Carol Lee",
        # …
    ]
    
    # ←— PREPARE default output path in your Downloads folder
    now = datetime.datetime.now()
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    filename = f"Filtered_Report_{now.strftime('%Y%m%d_%H%M%S')}.xlsx"
    dst_file = os.path.join(downloads, filename)
    
    # --- 0) Pre-scan all sheets for month-year options
    excel = pd.ExcelFile(src_file, engine="openpyxl")
    month_periods = set()
    for sheet in excel.sheet_names:
        df_tmp = pd.read_excel(src_file, sheet_name=sheet, usecols=[0], engine="openpyxl")
        dates = pd.to_datetime(df_tmp.iloc[:,0], errors="coerce")
        month_periods.update(dates.dropna().dt.to_period("M").unique())
    sorted_periods = sorted(month_periods)
    options = [p.strftime("%B %Y") for p in sorted_periods]
    period_map = {p.strftime("%B %Y"): p for p in sorted_periods}
    
    # --- 1) GUI
    root = tk.Tk()
    root.title("Filter by Month-Year & Employee")
    root.geometry("440x540")
    root.resizable(False, False)
    
    # Source file (predefined)
    tk.Label(root, text="Source file (predefined):").pack(anchor="w", padx=10, pady=(10,0))
    tk.Entry(root, width=60, state="readonly", 
             readonlybackground="white", fg="black",
             textvariable=tk.StringVar(value=src_file)
            ).pack(padx=10, pady=(0,5))
    
    # Month-Year multi-select
    tk.Label(root, text="Select month-year(s):").pack(anchor="w", padx=10, pady=(15,0))
    lb = tk.Listbox(root, selectmode="multiple", height=14, exportselection=False)
    for opt in options:
        lb.insert("end", opt)
    lb.pack(padx=10, pady=5, fill="x")
    
    # Output file (predefined in Downloads)
    tk.Label(root, text="Output file (predefined):").pack(anchor="w", padx=10, pady=(15,0))
    tk.Entry(root, width=60, state="readonly", 
             readonlybackground="white", fg="black",
             textvariable=tk.StringVar(value=dst_file)
            ).pack(padx=10, pady=(0,5))
    
    # Process button
    def on_submit():
        sel = lb.curselection()
        if not sel:
            messagebox.showerror("Error", "Please select at least one month-year.")
            return
        
        sel_periods = { period_map[options[i]] for i in sel }
        
        try:
            # Read & filter all sheets
            all_sheets = pd.read_excel(src_file, sheet_name=None, engine="openpyxl", header=0)
            filtered = {}
            present = set()
            for name, df in all_sheets.items():
                if df.shape[1] < 5: continue
                date_col, emp_col = df.columns[0], df.columns[4]
                
                df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
                periods = df[date_col].dt.to_period("M")
                mask = periods.isin(sel_periods) & df[emp_col].isin(employees)
                
                df2 = df.loc[mask].copy()
                if df2.empty: continue
                
                filtered[name] = df2
                present.update(df2[emp_col].dropna().unique())
            
            if not filtered:
                messagebox.showinfo("No Data", "No rows matched your criteria.")
                return
            
            # Compute missing employees
            missing = sorted(set(employees) - present)
            
            # Write & format
            with pd.ExcelWriter(dst_file, engine="openpyxl") as writer:
                for name, df2 in filtered.items():
                    safe = name[:31]
                    df2.to_excel(writer, sheet_name=safe, index=False)
                for name, df2 in filtered.items():
                    safe = name[:31]
                    ws = writer.sheets[safe]
                    # auto-fit columns
                    for idx, col in enumerate(df2.columns, 1):
                        col_letter = get_column_letter(idx)
                        maxlen = max(len(str(col)), *(df2[col].astype(str).map(len)))
                        ws.column_dimensions[col_letter].width = maxlen + 2
                    # enforce date format on A
                    date_letter = get_column_letter(1)
                    for row in range(2, len(df2)+2):
                        ws[f"{date_letter}{row}"].number_format = "M/D/YYYY"
                writer.save()
            
            # Done popup with missing list
            msg = f"Filtered {sum(len(df) for df in filtered.values())} row(s) " \
                  f"across {len(filtered)} sheet(s).\n\nSaved to:\n{dst_file}\n\n"
            if missing:
                msg += "Employees with NO matching rows:\n• " + "\n• ".join(missing)
            else:
                msg += "All employees had at least one matching row."
            messagebox.showinfo("Done", msg)
        
        except Exception as e:
            messagebox.showerror("Processing Error", str(e))
    
    tk.Button(
        root, text="Submit",
        command=on_submit,
        bg="#4CAF50", fg="white",
        font=("Segoe UI", 12, "bold")
    ).pack(pady=20, ipadx=10, ipady=5)
    
    root.mainloop()

if __name__ == "__main__":
    main()