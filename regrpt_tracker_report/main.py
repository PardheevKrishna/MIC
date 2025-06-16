#!/usr/bin/env python3
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def main():
    # ←— SET YOUR SOURCE .xlsm PATH HERE:
    src_file = r"C:\path\to\your\RegulatoryReporting.xlsm"
    
    # ←— REPLACE this with your actual employee names
    employees = [
        "Alice Johnson",
        "Bob Smith",
        "Carol Lee",
        # …
    ]
    
    # --- GUI setup
    root = tk.Tk()
    root.title("Filter Across All Sheets")
    root.geometry("400x500")
    root.resizable(False, False)
    
    # Show the predefined source path
    tk.Label(root, text="Source file (predefined):").pack(anchor="w", padx=10, pady=(10,0))
    in_var = tk.StringVar(value=src_file)
    tk.Entry(root, textvariable=in_var, width=60, state="readonly").pack(padx=10, pady=(0,5))
    
    # Month multi-select
    tk.Label(root, text="Select month(s):").pack(anchor="w", padx=10, pady=(15,0))
    lb = tk.Listbox(root, selectmode="multiple", height=12, exportselection=False)
    months = [
        "January","February","March","April","May","June",
        "July","August","September","October","November","December"
    ]
    for m in months:
        lb.insert("end", m)
    lb.pack(padx=10, pady=5, fill="x")
    
    # Output file selection
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
    
    # Process button
    def on_submit():
        dst = out_var.get().strip()
        sel = lb.curselection()
        if not sel:
            messagebox.showerror("Error", "Please select at least one month.")
            return
        if not dst:
            messagebox.showerror("Error", "Please choose an output filename.")
            return
        
        months_selected = [i+1 for i in sel]  # Listbox is 0-based
        
        try:
            # 1) Read all sheets
            all_sheets = pd.read_excel(
                src_file,
                sheet_name=None,
                engine="openpyxl",
                header=0
            )
            
            filtered_sheets = {}
            total_rows = 0
            
            # 2) Loop & filter each sheet
            for sheet_name, df in all_sheets.items():
                # skip if too few columns
                if df.shape[1] < 5:
                    continue
                date_col = df.columns[0]
                emp_col  = df.columns[4]
                
                # parse + month
                df[date_col] = pd.to_datetime(
                    df[date_col],
                    format="%m/%d/%Y",
                    errors="coerce"
                )
                df["_month"] = df[date_col].dt.month
                
                # filter
                df2 = df[
                    df["_month"].isin(months_selected) &
                    df[emp_col].isin(employees)
                ].copy()
                
                if not df2.empty:
                    df2.drop(columns=["_month"], inplace=True)
                    filtered_sheets[sheet_name] = df2
                    total_rows += len(df2)
            
            if not filtered_sheets:
                messagebox.showinfo("No Data", "No rows matched your criteria.")
                return
            
            # 3) Write out new .xlsx with one sheet per filtered set
            with pd.ExcelWriter(dst, engine="openpyxl") as writer:
                for sheet_name, df2 in filtered_sheets.items():
                    # truncate sheet names to 31 chars if needed
                    safe_name = sheet_name[:31]
                    df2.to_excel(writer, sheet_name=safe_name, index=False)
            
            messagebox.showinfo(
                "Done",
                f"Filtered {total_rows} row(s) across {len(filtered_sheets)} sheet(s) →\n{dst}"
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