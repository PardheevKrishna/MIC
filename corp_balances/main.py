#!/usr/bin/env python3
import os
import getpass
import time
import threading
import logging
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import pandas as pd
import numpy as np
from docx import Document
from tqdm import tqdm

# ─── Logging ────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s',
    datefmt='%H:%M:%S'
)

# ─── GUI SETUP ─────────────────────────────────────────────────────────────
root = tk.Tk()
root.title('Metrics Processor')

# Frame & widgets
frame = ttk.Frame(root, padding=10)
frame.grid()

btn_upload = ttk.Button(frame, text='Upload Excel File')
btn_upload.grid(row=0, column=0, sticky='w')
lbl_file = ttk.Label(frame, text='No file selected')
lbl_file.grid(row=0, column=1, padx=10)

lbl_uploaded_date = ttk.Label(frame, text='Uploaded date: N/A')
lbl_uploaded_date.grid(row=1, column=0, columnspan=2, sticky='w', pady=5)

lbl_user = ttk.Label(frame, text='Uploaded by: N/A')
lbl_user.grid(row=2, column=0, columnspan=2, sticky='w', pady=5)

lbl_process_time = ttk.Label(frame, text='Process time: 0.00s')
lbl_process_time.grid(row=3, column=0, columnspan=2, sticky='w', pady=5)

lbl_rows = ttk.Label(frame, text='Processed rows: 0')
lbl_rows.grid(row=4, column=0, columnspan=2, sticky='w', pady=5)

lbl_elapsed = ttk.Label(frame, text='Elapsed time: 0.00s')
lbl_elapsed.grid(row=5, column=0, columnspan=2, sticky='w', pady=5)

def update_progress(processed_rows, elapsed):
    """Called via root.after to update the GUI per‐row."""
    lbl_rows.config(text=f'Processed rows: {processed_rows}')
    lbl_elapsed.config(text=f'Elapsed time: {elapsed:.2f}s')

# ─── PROCESSING FUNCTION ───────────────────────────────────────────────────
def process_file(file_path):
    start_time = time.perf_counter()
    logging.info(f"Starting processing of '{file_path}'")

    # Phase 1: Read Excel
    logging.info("Reading Excel file…")
    df = pd.read_excel(file_path, engine='openpyxl')
    nrows, ncols = df.shape
    logging.info(f"Loaded {nrows:,} rows × {ncols} cols")

    # Column positions:
    #  A → df.iloc[:,0]
    #  B–K → df.iloc[:,1:11]
    #  blank col at 11
    #  M–V → df.iloc[:,12:22]
    dates = df.iloc[:, 0].astype(str).tolist()
    val_df = df.iloc[:, 1:11]
    var_df = df.iloc[:, 12:22]

    # Phase 2: Compute missing var metrics from val metrics
    logging.info("Computing var metrics from val metrics…")
    computed = val_df.subtract(val_df.shift(-1)).fillna(np.nan)
    var_final = var_df.where(var_df.notna(), computed)
    logging.info("Var metrics ready.")

    # Phase 3: Compute summary statistics for each var‐column
    logging.info("Calculating summary statistics…")
    summary = []
    # columns M–V correspond to letters M…V
    letters = [chr(ord('M') + i) for i in range(var_final.shape[1])]
    for idx, letter in enumerate(letters):
        series = var_final.iloc[:, idx].dropna().astype(float)
        if series.empty:
            continue
        q0, q1, q2, q3, q4 = (
            series.quantile(0.0),
            series.quantile(0.25),
            series.quantile(0.5),
            series.quantile(0.75),
            series.quantile(1.0),
        )
        iqr   = q3 - q1
        mean  = series.mean()
        std   = series.std()
        low2  = mean - 2*std; up2  = mean + 2*std
        low3  = mean - 3*std; up3  = mean + 3*std
        low4  = mean - 4*std; up4  = mean + 4*std
        summary.append({
            'Metric': letter,
            'Q0': q0, 'Q1': q1, 'Q2': q2, 'Q3': q3, 'Q4': q4,
            'IQR': iqr,
            'Lower 2SD': low2, 'Upper 2SD': up2,
            'Lower 3SD': low3, 'Upper 3SD': up3,
            'Lower 4SD': low4, 'Upper 4SD': up4,
            'Rec Lower (3SD)': round(low3, 3),
            'Rec Upper (3SD)': round(up3, 3),
        })
    logging.info("Summary statistics done.")

    # Phase 4: Build the Word document
    logging.info("Generating Word document…")
    doc = Document()
    user = getpass.getuser()
    uploaded_date = time.strftime('%m/%d/%Y')
    filename = os.path.basename(file_path)

    # Meta info
    doc.add_paragraph(f"Uploaded by: {user}")
    doc.add_paragraph(f"Uploaded date: {uploaded_date}")
    doc.add_paragraph(f"File name: {filename}")
    p_time = doc.add_paragraph("Process time: calculating…")

    # Summary table
    cols_sum = [
        'Metric','Q0','Q1','Q2','Q3','Q4','IQR',
        'Lower 2SD','Upper 2SD',
        'Lower 3SD','Upper 3SD',
        'Lower 4SD','Upper 4SD',
        'Rec Lower (3SD)','Rec Upper (3SD)'
    ]
    tbl = doc.add_table(rows=1, cols=len(cols_sum))
    for i, h in enumerate(cols_sum):
        tbl.rows[0].cells[i].text = h
    for row in summary:
        cells = tbl.add_row().cells
        for i, key in enumerate(cols_sum):
            cells[i].text = str(row[key])

    # Page break
    doc.add_page_break()

    # Variances table on second page
    cols_var = ['Date'] + letters
    tbl2 = doc.add_table(rows=1, cols=len(cols_var))
    for i, h in enumerate(cols_var):
        tbl2.rows[0].cells[i].text = h

    # Row-by-row insertion so we can update both tqdm and the Tk UI each time
    for i in tqdm(range(nrows), desc='Building variances table', unit='row'):
        cells = tbl2.add_row().cells
        cells[0].text = dates[i]
        for j in range(len(letters)):
            val = var_final.iat[i, j]
            cells[j+1].text = '' if pd.isna(val) else str(val)

        # Update the GUI
        elapsed = time.perf_counter() - start_time
        root.after(0, update_progress, i+1, elapsed)

    # Compute and patch in process time
    total_time = time.perf_counter() - start_time
    p_time.text = f"Process time: {total_time:.2f} seconds"

    # Save
    out_doc = 'output.docx'
    doc.save(out_doc)
    logging.info(f"Document saved as '{out_doc}'")

    # Notify user
    messagebox.showinfo('Done', f"Finished in {total_time:.2f}s\nSaved → {out_doc}")
    btn_upload.config(state='normal')


# ─── BUTTON CALLBACK ────────────────────────────────────────────────────────
def on_upload():
    fp = filedialog.askopenfilename(
        title='Select Excel file',
        filetypes=[('Excel', '*.xlsx *.xls')]
    )
    if not fp:
        return
    # Update UI meta
    lbl_file.config(text=os.path.basename(fp))
    lbl_uploaded_date.config(text=f"Uploaded date: {time.strftime('%m/%d/%Y')}")
    lbl_user.config(text=f"Uploaded by: {getpass.getuser()}")

    # Kick off processing in a background thread
    btn_upload.config(state='disabled')
    threading.Thread(target=process_file, args=(fp,), daemon=True).start()

btn_upload.config(command=on_upload)

# ─── MAINLOOP ───────────────────────────────────────────────────────────────
root.mainloop()