#!/usr/bin/env python3
import tkinter as tk
from tkinter import filedialog, messagebox
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
    
    # --- 0) Pre-scan all sheets to build month-year options
    excel = pd.ExcelFile(src_file, engine="openpyxl")
    sheet_names = excel.sheet_names
    month_periods = set()
    for sheet in sheet_names:
        # read only column A
        df_tmp = pd.read_excel(
            src_file,
            sheet_name=sheet,
            usecols=[0],
            header=0,
            engine="openpyxl"
        )
        date_col = df_tmp.columns[0]
        dates = pd.to_datetime(df_tmp[date_col], errors="coerce")
        month_periods.update(dates.dropna().dt.to_period("M").unique())
    # sort and format
    sorted_periods = sorted(month_periods)
    options = [p.strftime("%B %Y") for p in sorted_periods]
    period_map = {p.strftime("%B %Y"): p for p in sorted_periods}
    
    # --- 1) Build the GUI
    root = tk.Tk()
    root.title("Filter All Sheets by Month-Year & Employee")
    root.geometry("420x520")
    root.resizable(False, False)
    
    # show source path
    tk.Label(root, text="Source file (predefined):").pack(anchor="w", padx=10, pady=(10,0))
    in_var = tk.StringVar(value=src_file)
    tk.Entry(root, textvariable=in_var, width=60, state="readonly").pack(padx=10, pady=(0,5))
    
    # month-year multi-select
    tk.Label(root, text="Select month-year(s):").pack(anchor="w", padx=10, pady=(15,0))
    lb = tk.Listbox(root, selectmode="multiple", height=14, exportselection=False)
    for opt in options:
        lb.insert("end", opt)
    lb.pack(padx=10, pady=5, fill="x")
    
    # output path
    tk.Label(root, text="Choose output Excel (.xlsx):").pack(anchor="w", padx=10, pady=(15,0))
    out_var = tk.StringVar()
    tk.Entry(root, textvariable=out_var, width=60, state="readonly").pack(padx=10)
    def select_output():
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook","*.xlsx"),("All files","*.*")],
            title="Save filtered data as…"
        )
        if path:
            out_var.set(path)
    tk.Button(root, text="Browse…", command=select_output).pack(pady=5)
    
    # process
    def on_submit():
        dst = out_var.get().strip()
        sel = lb.curselection()
        if not sel:
            messagebox.showerror("Error", "Please select at least one month-year.")
            return
        if not dst:
            messagebox.showerror("Error", "Please choose an output filename.")
            return
        
        # build set of Periods to match
        chosen = [options[i] for i in sel]
        sel_periods = {period_map[c] for c in chosen}
        
        try:
            # read all sheets
            all_sheets = pd.read_excel(
                src_file,
                sheet_name=None,
                engine="openpyxl",
                header=0
            )
            
            filtered = {}
            total = 0
            
            for name, df in all_sheets.items():
                if df.shape[1] < 5:
                    continue
                date_col = df.columns[0]
                emp_col  = df.columns[4]
                
                # ensure datetime
                df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
                periods = df[date_col].dt.to_period("M")
                
                # filter
                mask = periods.isin(sel_periods) & df[emp_col].isin(employees)
                df2 = df.loc[mask].copy()
                if df2.empty:
                    continue
                
                filtered[name] = df2
                total += len(df2)
            
            if not filtered:
                messagebox.showinfo("No Data", "No rows matched your criteria.")
                return
            
            # write & format
            with pd.ExcelWriter(dst, engine="openpyxl") as writer:
                for name, df2 in filtered.items():
                    safe = name[:31]
                    df2.to_excel(writer, sheet_name=safe, index=False)
                
                # adjust each sheet
                for name, df2 in filtered.items():
                    safe = name[:31]
                    ws = writer.sheets[safe]
                    
                    # auto-fit columns
                    for idx, col in enumerate(df2.columns, 1):
                        col_letter = get_column_letter(idx)
                        vals = [str(col)] + df2[col].astype(str).tolist()
                        maxlen = max(len(v) for v in vals)
                        ws.column_dimensions[col_letter].width = maxlen + 2
                    
                    # enforce date format M/D/YYYY on column A
                    date_col_letter = get_column_letter(1)
                    for row in range(2, len(df2) + 2):
                        cell = ws[f"{date_col_letter}{row}"]
                        cell.number_format = "M/D/YYYY"
                
                writer.save()
            
            messagebox.showinfo(
                "Done",
                f"Filtered {total} row(s) across {len(filtered)} sheet(s) →\n{dst}"
            )
        
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