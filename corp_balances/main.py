#!/usr/bin/env python3
import os, getpass, time, threading, logging, datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import numpy as np
import pandas as pd
from openpyxl import load_workbook

# PDF generation via ReportLab
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer, PageBreak
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from tqdm import tqdm

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

def update_progress(r, elapsed):
    lbl_rows.config(text=f"Processed rows: {r}")
    lbl_time.config(text=f"Elapsed time: {elapsed:.2f}s")

# ─── Worker ─────────────────────────────────────────────────────────────────
def process_file(path):
    t0 = time.perf_counter()
    logging.info(f"Opening (read_only) '{path}'…")

    # 1) Stream‐read only cols A, B–K, M–V
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    total_rows = ws.max_row - 1  # minus header

    # grab header names from row 1
    header = [cell.value for cell in ws[1]][:22]
    colA_name    = header[0]
    val_names    = header[1:11]
    # skip header[11] (blank) then:
    var_names    = header[12:22]

    # preallocate
    colA      = np.empty(total_rows, dtype="U50")
    val_block = np.full((total_rows, 10), np.nan)
    var_block = np.full((total_rows, 10), np.nan)

    logging.info("Streaming rows from Excel…")
    it = ws.iter_rows(min_row=2, max_row=ws.max_row,
                      min_col=1, max_col=22, values_only=True)
    for i, row in enumerate(tqdm(it, total=total_rows, desc="Reading Excel", unit="row")):
        # Column A: date or text
        a = row[0]
        if isinstance(a, (datetime.date, datetime.datetime)):
            colA[i] = a.strftime("%m/%d/%Y")
        else:
            colA[i] = str(a) if a is not None else ""
        # val & var
        val_block[i, :] = row[1:11]
        var_block[i, :] = row[12:22]

        # GUI update
        elapsed = time.perf_counter() - t0
        root.after(0, update_progress, i+1, elapsed)

    wb.close()
    logging.info(f"Read {total_rows:,} rows in {time.perf_counter()-t0:.2f}s")

    # 2) Vectorized fill of missing variances
    logging.info("Computing missing variances…")
    shifted = np.vstack([val_block[1:], np.full((1,10), np.nan)])
    mask = np.isnan(var_block)
    var_block[mask] = (val_block - shifted)[mask]

    # 3) Summary‐stats pivoted
    logging.info("Computing summary statistics…")
    stats = [
        ("Q0",   0),
        ("Q1",  25),
        ("Q2",  50),
        ("Q3",  75),
        ("Q4", 100),
        ("IQR", None),
        ("Lower 2SD", None),
        ("Upper 2SD", None),
        ("Lower 3SD", None),
        ("Upper 3SD", None),
        ("Lower 4SD", None),
        ("Upper 4SD", None),
        ("Rec Lower Thresh", None),
        ("Rec Upper Thresh", None),
    ]
    # build DataFrame: index=metrics, columns=stat names
    df_stats = pd.DataFrame(index=var_names, columns=[s[0] for s in stats], dtype=float)

    for j, m in enumerate(var_names):
        col = var_block[:, j]
        col = col[~np.isnan(col)]
        if col.size == 0:
            continue
        # quartiles
        q = np.percentile(col, [0,25,50,75,100])
        mean = col.mean()
        std  = col.std(ddof=1)
        # fill quartiles
        df_stats.at[m, "Q0"] = q[0]
        df_stats.at[m, "Q1"] = q[1]
        df_stats.at[m, "Q2"] = q[2]
        df_stats.at[m, "Q3"] = q[3]
        df_stats.at[m, "Q4"] = q[4]
        df_stats.at[m, "IQR"] = q[3] - q[1]
        # ±2/3/4 SD
        df_stats.at[m, "Lower 2SD"] = mean - 2*std
        df_stats.at[m, "Upper 2SD"] = mean + 2*std
        df_stats.at[m, "Lower 3SD"] = mean - 3*std
        df_stats.at[m, "Upper 3SD"] = mean + 3*std
        df_stats.at[m, "Lower 4SD"] = mean - 4*std
        df_stats.at[m, "Upper 4SD"] = mean + 4*std
        # recommended rounded to nearest 1 000
        df_stats.at[m, "Rec Lower Thresh"] = round(mean - 3*std, -3)
        df_stats.at[m, "Rec Upper Thresh"] = round(mean + 3*std, -3)

    # 4) Save full variance table as CSV (for reference)
    df_var = pd.DataFrame(var_block, columns=var_names)
    df_var.insert(0, colA_name, colA)
    csv_out = "variances.csv"
    df_var.to_csv(csv_out, index=False)
    logging.info(f"Wrote variance CSV → {csv_out}")

    # 5) Build the PDF
    logging.info("Generating output.pdf…")
    pdf_out = "output.pdf"
    doc = SimpleDocTemplate(
        pdf_out,
        pagesize=landscape(A4),
        leftMargin=20, rightMargin=20,
        topMargin=20, bottomMargin=20
    )
    styles = getSampleStyleSheet()
    elems = []

    # Meta info
    user = getpass.getuser()
    today = time.strftime("%m/%d/%Y")
    filename = os.path.basename(path)
    for k,v in [
        ("Uploaded by", user),
        ("Uploaded date", today),
        ("File name", filename),
        ("Process time", f"{time.perf_counter()-t0:.2f} s")
    ]:
        elems.append(Paragraph(f"<b>{k}:</b> {v}", styles["Normal"]))
    elems.append(Spacer(1,12))

    # Summary table: stats as rows, metrics as columns
    # build data grid
    summary_data = [[ "Statistic" ] + var_names]
    for stat_name in df_stats.columns:
        row = [ stat_name ] + [
            f"{df_stats.at[m,stat_name]:,.2f}"
            for m in var_names
        ]
        summary_data.append(row)

    tbl = Table(summary_data, repeatRows=1, hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE",  (0,0), (-1,0),    8),
        ("FONTSIZE",  (0,1), (-1,-1),   6),
        ("ALIGN",     (1,0), (-1,-1), "RIGHT"),
        ("GRID",      (0,0), (-1,-1), 0.25, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
    ]))
    elems.append(tbl)
    elems.append(PageBreak())

    # Variance table: full data
    # **Warning**: this is huge for 1 M rows, may be very large.
    # We insert literally every row – adjust the font down if needed.
    var_data = [ [colA_name] + var_names ]
    # stream row-by-row to avoid intermediate giant list
    for i in tqdm(range(total_rows), desc="Building PDF table", unit="row"):
        row = [ colA[i] ] + [
            f"{var_block[i,j]:,.2f}" if not np.isnan(var_block[i,j]) else ""
            for j in range(len(var_names))
        ]
        var_data.append(row)

    tbl2 = Table(var_data, repeatRows=1, hAlign="LEFT")
    tbl2.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE",  (0,0), (-1,0),    7),
        ("FONTSIZE",  (0,1), (-1,-1),   4),
        ("ALIGN",     (1,0), (-1,-1), "RIGHT"),
        ("GRID",      (0,0), (-1,-1), 0.25, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
    ]))
    elems.append(tbl2)

    doc.build(elems)
    logging.info(f"PDF saved → {pdf_out}")

    # Final UI update & notify
    root.after(0, lbl_by.  config, {"text":f"Uploaded by: {user}"})
    root.after(0, lbl_date.config,  {"text":f"Uploaded date: {today}"})
    root.after(0, lbl_time.config,  {"text":f"Process time: {time.perf_counter()-t0:.2f}s"})
    root.after(0, lbl_rows.config,  {"text":f"Processed rows: {total_rows}"})
    messagebox.showinfo("Done",
        f"Completed in {time.perf_counter()-t0:.2f}s\n"
        f"PDF → {pdf_out}\nCSV → {csv_out}"
    )
    btn.config(state="normal")


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