#!/usr/bin/env python3
import os
import getpass
import time
import threading
import logging
import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from tqdm import tqdm

from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfgen.canvas import Canvas

# ─── Logging ───────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    datefmt="%H:%M:%S"
)

# ─── Tk UI ─────────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("Metrics ⇒ PDF + Excel")
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

# ─── Canvas for PDF page tqdm ───────────────────────────────────────────────
class ProgressCanvas(Canvas):
    def __init__(self, filename, progress, **kwargs):
        super().__init__(filename, **kwargs)
        self._pbar = progress
    def showPage(self):
        super().showPage()
        self._pbar.update(1)
    def save(self):
        super().save()
        self._pbar.close()

# ─── Main work ───────────────────────────────────────────────────────────────
def process_file(path):
    t0 = time.perf_counter()
    user  = getpass.getuser()
    today = time.strftime("%m/%d/%Y")

    # immediate UI update
    root.after(0, lambda: lbl_by  .config(text=f"Uploaded by: {user}"))
    root.after(0, lambda: lbl_date.config(text=f"Uploaded date: {today}"))

    # 1) read header row
    logging.info("Reading header row…")
    wb_h = load_workbook(path, read_only=True)
    ws_h = wb_h.active
    hdr = [c.value for c in ws_h[1]][:22]
    wb_h.close()
    colA_name = hdr[0]
    val_names = hdr[1:11]
    var_names = hdr[12:22]
    usecols   = [colA_name] + val_names + var_names

    # 2) fast read via pandas with no per-row tqdm
    logging.info("Loading Excel via pandas.read_excel…")
    df = pd.read_excel(path, usecols=usecols, engine='openpyxl')
    n  = len(df)
    elapsed = time.perf_counter() - t0
    logging.info(f"Loaded {n:,} rows in {elapsed:.2f}s")
    root.after(0, update_progress, n, elapsed)

    colA      = df[colA_name].astype(str).to_numpy()
    val_block = df[val_names].to_numpy(float)
    var_block = df[var_names].to_numpy(float)

    # 3) fill missing variances if any (dummy tqdm 1 step)
    if np.isnan(var_block).any():
        logging.info("Filling missing variances…")
        pbar_fill = tqdm(total=1, desc="Filling variances", unit="step")
        shifted = np.vstack([val_block[1:], np.full((1,10), np.nan)])
        m = np.isnan(var_block)
        var_block[m] = (val_block - shifted)[m]
        pbar_fill.update(1)
        pbar_fill.close()
    else:
        logging.info("All var metrics present; skipping fill.")

    # 4) summary statistics per metric
    logging.info("Computing summary statistics…")
    stats = [
        "Q0","Q1","Q2","Q3","Q4","IQR",
        "Lower 2SD","Upper 2SD",
        "Lower 3SD","Upper 3SD",
        "Lower 4SD","Upper 4SD",
        "Rec Lower Thresh","Rec Upper Thresh"
    ]
    df_stats = pd.DataFrame(index=var_names, columns=stats, dtype=float)
    for j, mn in tqdm(enumerate(var_names),
                      total=len(var_names),
                      desc="Summary stats",
                      unit="metric"):
        col = var_block[:,j]
        col = col[~np.isnan(col)]
        if col.size==0: continue
        q = np.percentile(col, [0,25,50,75,100])
        mean,std = col.mean(), col.std(ddof=1)
        df_stats.loc[mn, ["Q0","Q1","Q2","Q3","Q4"]] = q
        df_stats.at[mn,"IQR"] = q[3]-q[1]
        for k in (2,3,4):
            df_stats.at[mn,f"Lower {k}SD"] = mean - k*std
            df_stats.at[mn,f"Upper {k}SD"] = mean + k*std
        df_stats.at[mn,"Rec Lower Thresh"] = round(mean-3*std, -3)
        df_stats.at[mn,"Rec Upper Thresh"] = round(mean+3*std, -3)

    # 5) write variances.xlsx with a 1-step tqdm
    logging.info("Writing variances.xlsx…")
    pbar_xlsx = tqdm(total=1, desc="Writing Excel", unit="task")
    df_var = pd.DataFrame(var_block, columns=var_names)
    df_var.insert(0, colA_name, colA)
    var_xlsx = "variances.xlsx"
    with pd.ExcelWriter(var_xlsx, engine="xlsxwriter",
                        options={"constant_memory":True}) as writer:
        df_var.to_excel(writer, index=False, sheet_name="Sheet1")
    pbar_xlsx.update(1); pbar_xlsx.close()
    logging.info(f"Wrote '{var_xlsx}'")

    # 6) PDF summary + link
    logging.info("Building PDF summary…")
    pdf_out = "summary.pdf"
    doc = SimpleDocTemplate(
        pdf_out, pagesize=landscape(A4),
        leftMargin=20, rightMargin=20,
        topMargin=20, bottomMargin=20
    )
    styles = getSampleStyleSheet()
    elems = []

    for label,val in [
        ("Uploaded by", user),
        ("Uploaded date", today),
        ("Original file", os.path.basename(path)),
        ("Process time", f"{time.perf_counter()-t0:.2f}s")
    ]:
        elems.append(Paragraph(f"<b>{label}:</b> {val}",
                               styles["Normal"]))
    elems.append(Spacer(1,12))

    header_row = ["Statistic"] + var_names
    data = [header_row]
    for st in stats:
        data.append([st] + [f"{df_stats.at[m,st]:,.2f}" for m in var_names])

    elems.append(Table(
        data, repeatRows=1, hAlign="LEFT",
        style=TableStyle([
            ("GRID",(0,0),(-1,-1),0.25,colors.black),
            ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
            ("ALIGN",(1,0),(-1,-1),"RIGHT"),
            ("FONTSIZE",(0,0),(-1,0),10),
            ("FONTSIZE",(0,1),(-1,-1),8),
        ])
    ))
    elems.append(Spacer(1,12))

    link = f"file:///{os.path.abspath(var_xlsx)}"
    elems.append(Paragraph(
        f'<a href="{link}">Download full variances Excel</a>',
        styles["Normal"]
    ))

    pbar_pdf = tqdm(total=1, desc="PDF pages", unit="page")
    try:
        doc.build(
            elems,
            canvasmaker=lambda fn, **kw:
                ProgressCanvas(fn, progress=pbar_pdf, **kw)
        )
        logging.info(f"Built PDF '{pdf_out}'")
    except Exception:
        logging.exception("PDF build error")
        root.after(0,
            lambda: messagebox.showerror(
                "PDF Error","See console"))
        btn.config(state="normal")
        return

    # final UI update
    def finish():
        elapsed = time.perf_counter()-t0
        lbl_time .config(text=f"Process time: {elapsed:.2f}s")
        lbl_rows .config(text=f"Processed rows: {n}")
        messagebox.showinfo("Done",
            f"Completed in {elapsed:.2f}s\n"
            f"PDF → {pdf_out}\n"
            f"Excel → {var_xlsx}"
        )
        btn.config(state="normal")
    root.after(0, finish)

def on_upload():
    fn = filedialog.askopenfilename(
        title="Select Excel",
        filetypes=[("Excel","*.xlsx *.xls")]
    )
    if not fn:
        return
    lbl_file.config(text=os.path.basename(fn))
    btn.config(state="disabled")
    threading.Thread(target=process_file,
                     args=(fn,), daemon=True).start()

btn.config(command=on_upload)
root.mainloop()