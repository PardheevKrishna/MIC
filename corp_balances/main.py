#!/usr/bin/env python3
import os, getpass, time, threading, logging
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from docx import Document
from tqdm import tqdm

# ─── Logging ───────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    datefmt="%H:%M:%S"
)

# ─── TK UI SETUP ────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("Metrics Processor")
frame = ttk.Frame(root, padding=10)
frame.grid()

btn = ttk.Button(frame, text="Upload Excel")
btn.grid(row=0, column=0)
lbl_file = ttk.Label(frame, text="No file")
lbl_file.grid(row=0, column=1, padx=10)

lbl_by   = ttk.Label(frame, text="Uploaded by: N/A")
lbl_date = ttk.Label(frame, text="Uploaded date: N/A")
lbl_time = ttk.Label(frame, text="Process time: N/A")
lbl_rows = ttk.Label(frame, text="Processed rows: 0")

for i, w in enumerate((lbl_by, lbl_date, lbl_time, lbl_rows), start=1):
    w.grid(row=i, column=0, columnspan=2, sticky="w", pady=2)

def update_progress(r, elapsed):
    lbl_rows .config(text=f"Processed rows: {r}")
    lbl_time .config(text=f"Elapsed time: {elapsed:.2f}s")

# ─── WORKER ─────────────────────────────────────────────────────────────────
def process_file(path):
    start_all = time.perf_counter()
    logging.info(f"Loading (read_only) '{path}'…")

    # ── 1) STREAM‐READ only cols A,B–K,M–V via openpyxl read_only ─────────────
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    # we'll collect in numpy arrays for speed
    rows = ws.max_row - 1  # minus header
    dates = np.empty(rows, dtype="U10")
    val_block = np.empty((rows, 10), dtype="float64")
    var_block = np.empty((rows, 10), dtype="float64")
    val_block.fill(np.nan)
    var_block.fill(np.nan)

    logging.info("Reading rows…")
    it = ws.iter_rows(min_row=2, max_row=ws.max_row,
                      min_col=1, max_col=22, values_only=True)
    for i, row in enumerate(tqdm(it, total=rows, desc="Reading Excel", unit="row")):
        # row: (A,B…K,L,M…V,…) total 22 cols; we ignore L (idx=11)
        dates[i] = row[0].strftime("%m/%d/%Y")
        val_block[i,:] = row[1:11]
        var_block[i,:] = row[12:22]

    wb.close()
    logging.info(f"Read {rows:,} rows in {time.perf_counter()-start_all:.2f}s")

    # ── 2) COMPUTE missing var = val[i] - val[i+1] vectorized ────────────────
    logging.info("Computing missing variances…")
    # shift up by one for next‐row val
    shifted = np.vstack([val_block[1:], np.full((1,10), np.nan)])
    mask_missing = np.isnan(var_block)
    var_block[mask_missing] = (val_block - shifted)[mask_missing]

    # ── 3) SUMMARY STATS (quartiles, IQR, ±2/3/4 SD) ─────────────────────────
    logging.info("Calculating summary statistics…")
    summary = []
    for j in range(10):
        col = var_block[:, j]
        col = col[~np.isnan(col)]
        if col.size == 0:
            continue
        q0, q1, q2, q3, q4 = np.percentile(col, [0,25,50,75,100])
        iqr = q3 - q1
        mean = col.mean(); std = col.std(ddof=1)
        lows = [mean - k*std for k in (2,3,4)]
        ups  = [mean + k*std for k in (2,3,4)]
        summary.append({
            "Metric": f"Var_Metric{j+1}",
            "Q0": q0, "Q1": q1, "Q2": q2, "Q3": q3, "Q4": q4,
            "IQR": iqr,
            "Lower 2SD": lows[0],  "Upper 2SD": ups[0],
            "Lower 3SD": lows[1],  "Upper 3SD": ups[1],
            "Lower 4SD": lows[2],  "Upper 4SD": ups[2],
            "Rec Low (3SD)": round(lows[1],3),
            "Rec Up  (3SD)": round(ups[1],3)
        })

    # ── 4) DUMP variance table to CSV (blazing fast) ─────────────────────────
    var_df = pd.DataFrame(var_block, columns=[f"Var_Metric{i}" for i in range(1,11)])
    var_csv = "variances.csv"
    var_df.insert(0, "Date", dates)
    var_df.to_csv(var_csv, index=False)
    logging.info(f"Wrote variance CSV → {var_csv}")

    # ── 5) BUILD DOCX (meta + summary + link to CSV) ──────────────────────────
    logging.info("Building DOCX…")
    doc = Document()
    user = getpass.getuser()
    today = time.strftime("%m/%d/%Y")
    fname = os.path.basename(path)

    doc.add_paragraph(f"Uploaded by: {user}")
    doc.add_paragraph(f"Uploaded date: {today}")
    doc.add_paragraph(f"File name: {fname}")
    p_time = doc.add_paragraph("Process time: calculating…")

    # summary table
    cols = list(summary[0].keys())
    tbl = doc.add_table(rows=1, cols=len(cols))
    for i, h in enumerate(cols):
        tbl.rows[0].cells[i].text = h
    for row in summary:
        cells = tbl.add_row().cells
        for i, h in enumerate(cols):
            cells[i].text = str(row[h])

    # second page: link to CSV
    doc.add_page_break()
    para = doc.add_paragraph("Full variance table (1 M rows) → ")
    link = para.add_run(var_csv)
    link.font.underline = True

    total_t = time.perf_counter() - start_all
    p_time.clear().add_run(f"Process time: {total_t:.2f} seconds")

    out = "output.docx"
    doc.save(out)
    logging.info(f"Saved DOCX → {out}")

    # GUI updates
    root.after(0, lbl_by .config, {"text":f"Uploaded by: {user}"})
    root.after(0, lbl_date.config, {"text":f"Uploaded date: {today}"})
    root.after(0, lbl_time.config, {"text":f"Process time: {total_t:.2f}s"})
    root.after(0, lbl_rows.config, {"text":f"Processed rows: {rows}"})

    messagebox.showinfo("Done",
        f"Completed in {total_t:.2f}s\n"
        f"- Summary DOCX: {out}\n"
        f"- Variance CSV: {var_csv}"
    )
    btn.config(state="normal")

# ─── UPLOAD CALLBACK ────────────────────────────────────────────────────────
def on_upload():
    fn = filedialog.askopenfilename(
        title="Select Excel",
        filetypes=[("Excel", "*.xlsx *.xls")]
    )
    if not fn:
        return
    lbl_file.config(text=os.path.basename(fn))
    btn.config(state="disabled")
    threading.Thread(target=process_file, args=(fn,), daemon=True).start()

btn.config(command=on_upload)
root.mainloop()