#!/usr/bin/env python3
import os
import datetime
import threading
import queue
import logging
import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import DateEntry
import pandas as pd
import win32com.client as win32

# ─── CONFIG ───────────────────────────────────────────────────────────────────
SRC_FILE = r"C:\path\to\your\CRIT_Master.xlsm"

consumer_dataProvider   = ["Alice Johnson", "Bob Smith"]
commercial_dataProvider = ["Carol Lee",    "Dan Patel"]
scheduleOwner           = ["Eve Zhang",    "Frank Müller"]
CR360Transformation     = ["Grace Kim",    "Hiro Tanaka"]

SHEET_MAP = {
    "Reg_Reporting_DP": [
        ("consumer_dataProvider",   consumer_dataProvider),
        ("commercial_dataProvider", commercial_dataProvider),
    ],
    "Reg_Reporting_SO": [
        ("scheduleOwner", scheduleOwner),
    ],
    "CR360": [
        ("CR360Transformation", CR360Transformation),
    ],
}

EMAIL_TO      = ["to1@example.com"]
EMAIL_CC      = ["cc1@example.com"]
EMAIL_SUBJECT = "CRIT Automated Summary"
# ─── END CONFIG ───────────────────────────────────────────────────────────────

# ─── logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)

def send_via_outlook(to, cc, subject, html_body, attachments):
    logger.info("Creating Outlook email...")
    ol = win32.Dispatch('Outlook.Application')
    mail = ol.CreateItem(0)
    mail.To      = ";".join(to)
    mail.CC      = ";".join(cc)
    mail.Subject = subject
    mail.HTMLBody = html_body
    for f in attachments:
        logger.info(f" Attaching {os.path.basename(f)}")
        mail.Attachments.Add(f)
    mail.Send()
    logger.info("Email sent.")

def worker(start, end, q):
    """
    Puts:
      ('init', total_tasks)
      ('task', sheet→list)
      ('sheet_info', total_rows)
      ('filtered_info', filtered_rows)
      ('missing_info', [names])
      ('progress', tasks_done)
      ('done', attachments, missing_info_map)
      ('error', errmsg)
    """
    try:
        # build flat task list
        tasks = [
            (sheet, list_name, emp_list)
            for sheet, lists in SHEET_MAP.items()
            for list_name, emp_list in lists
        ]
        total_tasks = len(tasks)
        q.put(('init', total_tasks))
        logger.info(f"{total_tasks} tasks queued.")

        now       = datetime.datetime.now()
        ts        = now.strftime("%Y%m%d_%H%M%S")
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")

        # preload all sheets
        logger.info("Loading all sheets into pandas...")
        sheet_dfs = {
            sheet: pd.read_excel(
                SRC_FILE, sheet_name=sheet,
                engine="openpyxl", parse_dates=[0]
            )
            for sheet in SHEET_MAP
        }
        logger.info("All sheets loaded.")

        attachments = []
        missing_overall = {}

        done = 0
        for sheet, list_name, emp_list in tasks:
            desc = f"{sheet} → {list_name}"
            logger.info("Starting " + desc)
            q.put(('task', desc))

            df = sheet_dfs[sheet]
            date_col = df.columns[0]
            emp_col  = df.columns[4]

            total_rows = len(df)
            q.put(('sheet_info', total_rows))

            # vectorized filter
            mask = (
                (df[date_col].dt.date >= start) &
                (df[date_col].dt.date <= end) &
                (df[emp_col].isin(emp_list))
            )
            df_f = df.loc[mask].copy()
            filtered_rows = len(df_f)
            q.put(('filtered_info', filtered_rows))

            present = set(df_f[emp_col].dropna().unique())
            missing = sorted(set(emp_list) - present)
            q.put(('missing_info', missing))
            missing_overall[list_name] = missing

            # write out filtered `.xlsx`
            out_name = f"{list_name}_{ts}.xlsx"
            out_path = os.path.join(downloads, out_name)
            with pd.ExcelWriter(out_path, engine="openpyxl") as w:
                df_f.to_excel(w, sheet_name=list_name[:31], index=False)
            attachments.append(out_path)
            logger.info(f"  → Saved {out_name} ({filtered_rows}/{total_rows} rows)")

            done += 1
            q.put(('progress', done))

        # build email body
        logger.info("Building email body...")
        html = ['<html><body><h1>Missing Entries</h1>']
        for ln, miss in missing_overall.items():
            html.append(f"<h2>{ln}</h2>")
            if not miss:
                html.append("<p>All employees reported in range.</p>")
            else:
                html.append("<ul>")
                for name in miss:
                    html.append(f"<li>{name}</li>")
                html.append("</ul>")
        html.append("</body></html>")
        body = "\n".join(html)

        # send
        send_via_outlook(EMAIL_TO, EMAIL_CC, EMAIL_SUBJECT, body, attachments)
        q.put(('done', attachments, missing_overall))
    except Exception as e:
        logger.exception("Worker error")
        q.put(('error', str(e)))

def start_process():
    start = start_cal.get_date()
    end   = end_cal.get_date()
    if start > end:
        return messagebox.showerror("Error", "Start must be ≤ End.")

    btn_submit.config(state='disabled')
    progress_var.set("Initializing…")
    task_var.set(""); sheet_var.set(""); filt_var.set(""); miss_var.set("")
    prog_bar['value'] = 0

    threading.Thread(target=worker, args=(start, end, q), daemon=True).start()
    root.after(100, poll_queue)

def poll_queue():
    try:
        msg = q.get_nowait()
    except queue.Empty:
        return root.after(100, poll_queue)

    kind = msg[0]
    if kind == 'init':
        total = msg[1]
        prog_bar['maximum'] = total
        progress_var.set(f"0 of {total} tasks")

    elif kind == 'task':
        task_var.set(f"Task: {msg[1]}")

    elif kind == 'sheet_info':
        sheet_var.set(f"Total rows: {msg[1]}")

    elif kind == 'filtered_info':
        filt_var.set(f"Filtered rows: {msg[1]}")

    elif kind == 'missing_info':
        miss = msg[1]
        miss_var.set("Missing: " + (", ".join(miss) if miss else "None"))

    elif kind == 'progress':
        done = msg[1]
        prog_bar['value'] = done
        total = prog_bar['maximum']
        progress_var.set(f"{done} of {total} tasks")

    elif kind == 'done':
        btn_submit.config(state='normal')
        messagebox.showinfo("Finished", "Reports saved & email sent via Outlook.")

    elif kind == 'error':
        btn_submit.config(state='normal')
        messagebox.showerror("Error", msg[1])

    root.after(100, poll_queue)

# ─── BUILD UI ─────────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("CRIT Filter & Emailer (fast)")
root.geometry("480x360")
root.resizable(False, False)

q = queue.Queue()

padx, wrap = 10, 460

tk.Label(root, text="Source (fixed):").pack(anchor="w", padx=padx, pady=(10,0))
tk.Label(root, text=SRC_FILE, wraplength=wrap, fg="gray").pack(anchor="w", padx=padx)

tk.Label(root, text="Start Date:").pack(anchor="w", padx=padx, pady=(10,0))
start_cal = DateEntry(root, date_pattern="mm/dd/yyyy"); start_cal.pack(padx=padx)

tk.Label(root, text="End Date:").pack(anchor="w", padx=padx, pady=(10,0))
end_cal   = DateEntry(root, date_pattern="mm/dd/yyyy"); end_cal.pack(padx=padx)

btn_submit = tk.Button(
    root, text="Submit", command=start_process,
    bg="#4CAF50", fg="white",
    font=("Segoe UI", 12, "bold")
)
btn_submit.pack(pady=15, ipadx=10, ipady=5)

progress_var = tk.StringVar(value="Idle")
tk.Label(root, textvariable=progress_var).pack(pady=(5,0))
prog_bar = ttk.Progressbar(root, length=450, mode="determinate")
prog_bar.pack(pady=(2,10))

task_var = tk.StringVar()
tk.Label(root, textvariable=task_var, wraplength=wrap, fg="blue").pack()

sheet_var = tk.StringVar()
tk.Label(root, textvariable=sheet_var, wraplength=wrap).pack()

filt_var = tk.StringVar()
tk.Label(root, textvariable=filt_var, wraplength=wrap).pack()

miss_var = tk.StringVar()
tk.Label(root, textvariable=miss_var, wraplength=wrap).pack()

root.mainloop()