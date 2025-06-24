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

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger()

def send_via_outlook(to, cc, subject, html_body, attachments):
    ol   = win32.Dispatch('Outlook.Application')
    mail = ol.CreateItem(0)
    mail.To      = ";".join(to)
    mail.CC      = ";".join(cc)
    mail.Subject = subject
    mail.HTMLBody = html_body
    for f in attachments:
        mail.Attachments.Add(f)
    mail.Send()
    logger.info("Email sent via Outlook.")

def worker(start_date, end_date, q):
    try:
        # 1) Build tasks
        tasks = [
            (sheet, list_name, emp_list)
            for sheet, groups in SHEET_MAP.items()
            for list_name, emp_list in groups
        ]
        q.put(('init', len(tasks)))
        ts        = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")

        # 2) Pre‐load and parse date column, create date_only
        logger.info("Loading and parsing sheets...")
        sheet_dfs = {}
        for sheet in SHEET_MAP:
            df = pd.read_excel(SRC_FILE, sheet_name=sheet, engine="openpyxl")
            date_col = df.columns[0]
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            df['date_only'] = df[date_col].dt.date
            sheet_dfs[sheet] = df
        logger.info("Sheets ready.")

        attachments = []
        missing_map  = {}
        done = 0

        for sheet, list_name, emp_list in tasks:
            desc = f"{sheet} → {list_name}"
            q.put(('task', desc))
            logger.info(f"Starting {desc}")

            df = sheet_dfs[sheet]
            date_only = df['date_only']
            emp_col   = df.columns[4]

            q.put(('sheet_info', len(df)))

            # 3) Filter between start and end
            mask = (
                (date_only >= start_date) &
                (date_only <= end_date) &
                (df[emp_col].isin(emp_list))
            )
            df_f = df.loc[mask].copy()
            q.put(('filtered_info', len(df_f)))
            logger.info(f"  → {len(df_f)} rows in range")

            # 4) Identify missing employees
            present = set(df_f[emp_col].dropna())
            missing = sorted(set(emp_list) - present)

            # 5) For each missing, find last date before start
            df_before = df[df['date_only'] < start_date]
            info_list = []
            for emp in missing:
                dates = df_before.loc[df_before[emp_col]==emp, 'date_only'].dropna().tolist()
                logger.info(f"  → {emp} pre-start dates: {dates}")
                if dates:
                    lb = max(dates).strftime("%m/%d/%Y")
                else:
                    lb = "N/A"
                info_list.append({"employee": emp, "last_before": lb})
            missing_map[list_name] = info_list
            q.put(('missing_info', missing))

            # 6) Drop helper column before saving
            df_save = df_f.drop(columns=['date_only'])

            # 7) Write out .xlsx (always at least headers)
            out_name = f"{list_name}_{ts}.xlsx"
            out_path = os.path.join(downloads, out_name)
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                if df_save.empty:
                    pd.DataFrame(columns=df.columns[:-1]).to_excel(
                        writer, sheet_name=list_name[:31], index=False
                    )
                else:
                    df_save.to_excel(writer, sheet_name=list_name[:31], index=False)
            attachments.append(out_path)
            logger.info(f"Saved {out_name}")

            done += 1
            q.put(('progress', done))

        # 8) Build HTML email with last_before column
        html = ['<html><body><h1>Missing Entries</h1>']
        for ln, recs in missing_map.items():
            html.append(f"<h2>{ln}</h2>")
            if not recs:
                html.append("<p>All employees reported in range.</p>")
            else:
                html.append(
                    "<table border='1' cellpadding='4'>"
                    "<tr><th>Employee</th><th>Last Update Before Start</th></tr>"
                )
                for r in recs:
                    html.append(
                        f"<tr><td>{r['employee']}</td>"
                        f"<td>{r['last_before']}</td></tr>"
                    )
                html.append("</table>")
        html.append("</body></html>")
        body = "\n".join(html)

        # 9) Send via Outlook
        send_via_outlook(EMAIL_TO, EMAIL_CC, EMAIL_SUBJECT, body, attachments)
        q.put(('done', attachments, missing_map))
        logger.info("All tasks complete.")

    except Exception as e:
        logger.exception("Worker error")
        q.put(('error', str(e)))

def start_process():
    s = start_cal.get_date()
    e = end_cal.get_date()
    if s > e:
        return messagebox.showerror("Error", "Start must be on or before End.")
    btn_submit.config(state='disabled')
    threading.Thread(target=worker, args=(s, e, q), daemon=True).start()
    root.after(100, poll_queue)

def poll_queue():
    try:
        msg = q.get_nowait()
    except queue.Empty:
        return root.after(100, poll_queue)
    kind, *rest = msg
    if kind == 'init':
        prog_bar['maximum'] = rest[0]
        progress_var.set(f"0 of {rest[0]} tasks")
    elif kind == 'task':
        task_var.set(f"Task: {rest[0]}")
    elif kind == 'sheet_info':
        sheet_var.set(f"Total rows: {rest[0]}")
    elif kind == 'filtered_info':
        filt_var.set(f"Filtered rows: {rest[0]}")
    elif kind == 'missing_info':
        miss_var.set("Missing: " + (", ".join(rest[0]) if rest[0] else "None"))
    elif kind == 'progress':
        prog_bar['value'] = rest[0]
        progress_var.set(f"{rest[0]} of {prog_bar['maximum']} tasks")
    elif kind == 'done':
        btn_submit.config(state='normal')
        messagebox.showinfo("Finished", "Reports saved & email sent via Outlook.")
    elif kind == 'error':
        btn_submit.config(state='normal')
        messagebox.showerror("Error", rest[0])
    root.after(100, poll_queue)

# ─── BUILD UI ─────────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("CRIT Filter & Emailer")
root.geometry("480x360")
root.resizable(False, False)

q = queue.Queue()

# Correct StringVar initialization
progress_var = tk.StringVar(value="Idle")
task_var     = tk.StringVar(value="")
sheet_var    = tk.StringVar(value="")
filt_var     = tk.StringVar(value="")
miss_var     = tk.StringVar(value="")

padx, wrap = 10, 460
tk.Label(root, text="Start Date:").pack(anchor="w", padx=padx, pady=(10,0))
start_cal = DateEntry(root, date_pattern="mm/dd/yyyy"); start_cal.pack(padx=padx)
tk.Label(root, text="End Date:").pack(anchor="w", padx=padx, pady=(10,0))
end_cal = DateEntry(root, date_pattern="mm/dd/yyyy"); end_cal.pack(padx=padx)

btn_submit = tk.Button(
    root, text="Submit", command=start_process,
    bg="#4CAF50", fg="white", font=("Segoe UI",12,"bold")
)
btn_submit.pack(pady=15, ipadx=10, ipady=5)

tk.Label(root, textvariable=progress_var).pack(pady=(5,0))
prog_bar = ttk.Progressbar(root, length=450, mode="determinate")
prog_bar.pack(pady=(2,10))

tk.Label(root, textvariable=task_var, wraplength=wrap, fg="blue").pack()
tk.Label(root, textvariable=sheet_var, wraplength=wrap).pack()
tk.Label(root, textvariable=filt_var, wraplength=wrap).pack()
tk.Label(root, textvariable=miss_var, wraplength=wrap).pack()

root.mainloop()