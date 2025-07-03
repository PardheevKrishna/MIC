#!/usr/bin/env python3
import os, getpass, time, threading, logging, datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import numpy as np
import pandas as pd

# ReportLab for PDF output
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer, PageBreak
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# ─── CONFIG ────────────────────────────────────────────────────────────────
# Toggle between FAST pandas.read_excel (no tqdm) vs. SLOW openpyxl+tqdm
FAST_READ = True

# ─── Logging ───────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    datefmt="%H:%M:%S"
)

# ─── Tkinter UI ────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("Metrics ⇒ PDF")
frame = ttk.Frame(root, padding=10)
frame.grid()

btn = ttk.Button(frame, text="Upload Excel")
btn.grid(row=0, column=0)
lbl_file = ttk.Label(frame, text="No file selected")
lbl_file.grid(row=0, column=1, padx=10)

lbl_by   = ttk.Label(frame, text="Uploaded by: N/A")
lbl_date = ttk.Label(frame, text="Uploaded date: N/A")
lbl_time = ttk.Label(frame, text="Process time: N/A")
lbl_rows = ttk.Label(frame, text="Processed rows: 0")

for i, w in enumerate((lbl_by, lbl_date, lbl_time, lbl_rows), start=1):
    w.grid(row=i, column=0, columnspan=2, sticky="w", pady=2)

def update_progress(rows, elapsed):
    lbl_rows.config(text=f"Processed rows: {rows}")
    lbl_time.config(text=f"Elapsed time: {elapsed:.2f}s")

# ─── Worker ─────────────────────────────────────────────────────────────────
def process_file(path):
    t0 = time.perf_counter()
    user = getpass.getuser()
    today = time.strftime("%m/%d/%Y")

    # immediate UI update
    root.after(0, lambda: lbl_by.config(text=f"Uploaded by: {user}"))
    root.after(0, lambda: lbl_date.config(text=f"Uploaded date: {today}"))

    # 1) Read headers to get column names
    logging.info("Reading header row to get column names…")
    header_df = pd.read_excel(path, nrows=0, engine='openpyxl')
    cols = list(header_df.columns)
    colA_name = cols[0]
    val_names = cols[1:11]
    var_names = cols[12:22]

    if FAST_READ:
        # 2a) FAST: atomic pandas.read_excel
        usecols = [colA_name] + val_names + var_names
        logging.info(f"Fast reading {len(usecols)} cols via pandas.read_excel…")
        df = pd.read_excel(path, usecols=usecols, engine='openpyxl')
        total_rows = len(df)
        logging.info(f"Loaded {total_rows:,} rows in {time.perf_counter()-t0:.2f}s")

        # Extract into numpy
        colA      = df[colA_name].astype(str).to_numpy()
        val_block = df[val_names].to_numpy(dtype=float)
        var_block = df[var_names].to_numpy(dtype=float)

        # update UI now that we know rows/time
        root.after(0, update_progress, total_rows, time.perf_counter()-t0)

    else:
        # 2b) SLOW: openpyxl + tqdm row-by-row (shows progress)
        from openpyxl import load_workbook
        from tqdm import tqdm

        logging.info("Opening workbook in read_only mode…")
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        total_rows = ws.max_row - 1

        # pre-alloc
        colA      = np.empty(total_rows, dtype="U50")
        val_block = np.full((total_rows, 10), np.nan)
        var_block = np.full((total_rows, 10), np.nan)

        logging.info("Streaming rows with tqdm…")
        it = ws.iter_rows(min_row=2, max_row=ws.max_row,
                          min_col=1, max_col=22,
                          values_only=True)
        for i, row in enumerate(tqdm(it, total=total_rows, desc="Reading Excel", unit="row")):
            a = row[0]
            if isinstance(a, (datetime.date, datetime.datetime)):
                colA[i] = a.strftime("%m/%d/%Y")
            else:
                colA[i] = str(a) if a is not None else ""
            val_block[i,:] = row[1:11]
            var_block[i,:] = row[12:22]

            # update UI each row
            elapsed = time.perf_counter()-t0
            root.after(0, update_progress, i+1, elapsed)

        wb.close()
        logging.info(f"Stream read {total_rows:,} rows in {time.perf_counter()-t0:.2f}s")

    # 3) Compute missing variances vectorized
    logging.info("Filling missing variances (vectorized)…")
    shifted    = np.vstack([val_block[1:], np.full((1,10), np.nan)])
    mask       = np.isnan(var_block)
    var_block[mask] = (val_block - shifted)[mask]

    # 4) Summary statistics pivoted
    logging.info("Computing summary statistics…")
    stats_idx = [
        "Q0","Q1","Q2","Q3","Q4","IQR",
        "Lower 2SD","Upper 2SD",
        "Lower 3SD","Upper 3SD",
        "Lower 4SD","Upper 4SD",
        "Rec Lower Thresh","Rec Upper Thresh"
    ]
    df_stats = pd.DataFrame(index=var_names, columns=stats_idx, dtype=float)

    for j, m in enumerate(var_names):
        col = var_block[:, j]
        col = col[~np.isnan(col)]
        if col.size == 0: continue
        q = np.percentile(col, [0,25,50,75,100])
        mean, std = col.mean(), col.std(ddof=1)
        df_stats.loc[m, ["Q0","Q1","Q2","Q3","Q4"]] = q
        df_stats.at[m, "IQR"] = q[3] - q[1]
        for k in (2,3,4):
            df_stats.at[m, f"Lower {k}SD"] = mean - k*std
            df_stats.at[m, f"Upper {k}SD"] = mean + k*std
        df_stats.at[m, "Rec Lower Thresh"] = round(mean - 3*std, -3)
        df_stats.at[m, "Rec Upper Thresh"] = round(mean + 3*std, -3)

    # 5) Write full variance table to CSV
    logging.info("Writing variance CSV…")
    df_var = pd.DataFrame(var_block, columns=var_names)
    df_var.insert(0, colA_name, colA)
    csv_out = "variances.csv"
    df_var.to_csv(csv_out, index=False)

    # 6) Build PDF with ReportLab
    logging.info("Building PDF…")
    pdf_out = "output.pdf"
    doc = SimpleDocTemplate(
        pdf_out,
        pagesize=landscape(A4),
        leftMargin=20, rightMargin=20,
        topMargin=20, bottomMargin=20
    )
    styles = getSampleStyleSheet()
    elems = []

    # Meta
    for k,v in [
        ("Uploaded by", user),
        ("Uploaded date", today),
        ("File name", os.path.basename(path)),
        ("Process time", f"{time.perf_counter()-t0:.2f}s")
    ]:
        elems.append(Paragraph(f"<b>{k}:</b> {v}", styles["Normal"]))
    elems.append(Spacer(1,12))

    # Summary‐stats as table: rows=stats, cols=metrics
    header_row = ["Statistic"] + var_names
    data = [header_row]
    for stat in stats_idx:
        row = [stat] + [f"{df_stats.at[m,stat]:,.2f}" for m in var_names]
        data.append(row)

    tbl = Table(data, repeatRows=1, hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE",  (0,0), (-1,0),    8),
        ("FONTSIZE",  (0,1), (-1,-1),   6),
        ("ALIGN",     (1,0), (-1,-1), "RIGHT"),
        ("GRID",      (0,0), (-1,-1), 0.25, colors.black),
        ("BACKGROUND",(0,0), (-1,0), colors.lightgrey),
    ]))
    elems.append(tbl)
    elems.append(PageBreak())

    # Full variance table (be careful—huge!)
    var_data = [[colA_name] + var_names]
    for i in range(total_rows):
        var_data.append([
            colA[i]
        ] + [
            f"{var_block[i,j]:,.2f}" if not np.isnan(var_block[i,j]) else ""
            for j in range(len(var_names))
        ])

    tbl2 = Table(var_data, repeatRows=1, hAlign="LEFT")
    tbl2.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE",  (0,0), (-1,0),    7),
        ("FONTSIZE",  (0,1), (-1,-1),   4),
        ("ALIGN",     (1,0), (-1,-1), "RIGHT"),
        ("GRID",      (0,0), (-1,-1), 0.25, colors.black),
        ("BACKGROUND",(0,0), (-1,0), colors.lightgrey),
    ]))
    elems.append(tbl2)

    try:
        doc.build(elems)
        logging.info("PDF build complete.")
    except Exception:
        logging.exception("PDF build error")
        root.after(0, lambda: messagebox.showerror("PDF Error","See console"))
        btn.config(state="normal")
        return

    # Final UI update & message
    def finish():
        lbl_time .config(text=f"Process time: {time.perf_counter()-t0:.2f}s")
        lbl_rows .config(text=f"Processed rows: {total_rows}")
        messagebox.showinfo("Done",
            f"Completed in {time.perf_counter()-t0:.2f}s\n"
            f"PDF → {pdf_out}\nCSV → {csv_out}"
        )
        btn.config(state="normal")

    root.after(0, finish)

# ─── Button hookup ─────────────────────────────────────────────────────────
def on_upload():
    fn = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel","*.xlsx *.xls")]
    )
    if not fn:
        return
    lbl_file.config(text=os.path.basename(fn))
    btn.config(state="disabled")
    threading.Thread(target=process_file, args=(fn,), daemon=True).start()

btn.config(command=on_upload)
root.mainloop()