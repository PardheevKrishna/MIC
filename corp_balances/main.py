#!/usr/bin/env python3
import os, getpass, time, threading, logging, datetime
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

# ─── Tkinter UI ────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("Metrics Processor")
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

def update_progress(r, elapsed):
    lbl_rows.config(text=f"Processed rows: {r}")
    lbl_time.config(text=f"Elapsed time: {elapsed:.2f}s")

# ─── Worker ─────────────────────────────────────────────────────────────────
def process_file(path):
    t0 = time.perf_counter()
    logging.info(f"Opening (read_only) '{path}'…")
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    total_rows = ws.max_row - 1  # exclude header

    # Pre-allocate storage
    colA = np.empty(total_rows, dtype="U50")       # for arbitrary text or date
    val_block = np.full((total_rows, 10), np.nan)  # B–K
    var_block = np.full((total_rows, 10), np.nan)  # M–V

    logging.info("Streaming rows from Excel…")
    it = ws.iter_rows(min_row=2, max_row=ws.max_row,
                      min_col=1, max_col=22, values_only=True)
    for i, row in enumerate(tqdm(it, total=total_rows, desc="Reading Excel", unit="row")):
        # Column A: could be date or name; stringify
        a = row[0]
        if isinstance(a, (datetime.date, datetime.datetime)):
            colA[i] = a.strftime("%m/%d/%Y")
        else:
            colA[i] = str(a) if a is not None else ""
        # Val metrics B–K:
        val_block[i, :] = row[1:11]
        # Var metrics M–V (skip the blank L at idx=11):
        var_block[i, :] = row[12:22]
        # GUI update
        elapsed = time.perf_counter() - t0
        root.after(0, update_progress, i+1, elapsed)

    wb.close()
    logging.info(f"Read {total_rows:,} rows in {time.perf_counter()-t0:.2f}s")

    # Fill missing variances: var = val[i] - val[i+1]
    logging.info("Vectorized fill of missing var metrics…")
    shifted = np.vstack([val_block[1:], np.full((1,10), np.nan)])
    mask = np.isnan(var_block)
    var_block[mask] = (val_block - shifted)[mask]

    # Summary stats per var‐column (M–V)
    logging.info("Computing summary statistics…")
    summary = []
    letters = [chr(ord('M') + j) for j in range(10)]
    for j, letter in enumerate(letters):
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
            "Metric": letter,
            "Q0": q0, "Q1": q1, "Q2": q2, "Q3": q3, "Q4": q4,
            "IQR": iqr,
            "Lower 2SD": lows[0],  "Upper 2SD": ups[0],
            "Lower 3SD": lows[1],  "Upper 3SD": ups[1],
            "Lower 4SD": lows[2],  "Upper 4SD": ups[2],
            "Rec Lower (3SD)": round(lows[1],3),
            "Rec Upper (3SD)": round(ups[1],3)
        })

    # Dump the full variance table to CSV
    logging.info("Writing variances.csv…")
    df_var = pd.DataFrame(var_block, columns=letters)
    df_var.insert(0, "A", colA)  # preserve col A as first column
    csv_out = "variances.csv"
    df_var.to_csv(csv_out, index=False)

    # Build the Word document
    logging.info("Generating output.docx…")
    doc = Document()
    user = getpass.getuser()
    today = time.strftime("%m/%d/%Y")
    filename = os.path.basename(path)

    doc.add_paragraph(f"Uploaded by: {user}")
    doc.add_paragraph(f"Uploaded date: {today}")
    doc.add_paragraph(f"File name: {filename}")
    ptime = doc.add_paragraph("Process time: calculating…")

    # Summary table
    cols = list(summary[0].keys())
    tbl = doc.add_table(rows=1, cols=len(cols))
    for idx, h in enumerate(cols):
        tbl.rows[0].cells[idx].text = h
    for row in summary:
        cells = tbl.add_row().cells
        for idx, h in enumerate(cols):
            cells[idx].text = str(row[h])

    doc.add_page_break()
    para = doc.add_paragraph("Full variance table saved as ")
    run = para.add_run(csv_out)
    run.font.underline = True

    total_t = time.perf_counter() - t0
    ptime.clear().add_run(f"Process time: {total_t:.2f} seconds")

    doc_out = "output.docx"
    doc.save(doc_out)

    # Final GUI update & notify
    root.after(0, lbl_by.config, {"text":f"Uploaded by: {user}"})
    root.after(0, lbl_date.config,{"text":f"Uploaded date: {today}"})
    root.after(0, lbl_time.config,{"text":f"Process time: {total_t:.2f}s"})
    root.after(0, lbl_rows.config,{"text":f"Processed rows: {total_rows}"})
    messagebox.showinfo("Done",
        f"Completed in {total_t:.2f}s\n"
        f"- Summary DOCX: {doc_out}\n"
        f"- Variances CSV: {csv_out}"
    )
    btn.config(state="normal")


# ─── Button callback ────────────────────────────────────────────────────────
def on_upload():
    path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel", "*.xlsx *.xls")]
    )
    if not path:
        return
    lbl_file.config(text=os.path.basename(path))
    lbl_by  .config(text="Uploaded by: N/A")
    lbl_date.config(text="Uploaded date: N/A")
    lbl_time.config(text="Process time: N/A")
    lbl_rows.config(text="Processed rows: 0")
    btn.config(state="disabled")
    threading.Thread(target=process_file, args=(path,), daemon=True).start()

btn.config(command=on_upload)
root.mainloop()