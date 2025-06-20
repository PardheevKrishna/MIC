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
logger = logging.getLogger()

def send_via_outlook(to, cc, subject, html_body, attachments):
    ol = win32.Dispatch('Outlook.Application')
    mail = ol.CreateItem(0)
    mail.To      = ";".join(to)
    mail.CC      = ";".join(cc)
    mail.Subject = subject
    mail.HTMLBody = html_body
    for f in attachments:
        mail.Attachments.Add(f)
    mail.Send()
    logger.info("Email sent via Outlook.")

def worker(start, end, q):
    try:
        # build flat list of tasks
        tasks = [
            (sheet, list_name, emp_list)
            for sheet, lists in SHEET_MAP.items()
            for list_name, emp_list in lists
        ]
        q.put(('init', len(tasks)))
        logger.info(f"{len(tasks)} tasks queued.")

        now       = datetime.datetime.now()
        ts        = now.strftime("%Y%m%d_%H%M%S")
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")

        # preload all sheets into pandas
        logger.info("Loading sheets into pandas...")
        sheet_dfs = {
            sheet: pd.read_excel(
                SRC_FILE, sheet_name=sheet,
                engine="openpyxl", parse_dates=[0]
            )
            for sheet in SHEET_MAP
        }
        logger.info("All sheets loaded.")

        attachments   = []
        missing_map   = {}
        done          = 0

        for sheet, list_name, emp_list in tasks:
            desc = f"{sheet} → {list_name}"
            q.put(('task', desc))
            logger.info("Starting " + desc)

            df = sheet_dfs[sheet]
            date_col = df.columns[0]
            emp_col  = df.columns[4]

            total_rows = len(df)
            q.put(('sheet_info', total_rows))

            # filter rows in range
            mask = (
                (df[date_col].dt.date >= start) &
                (df[date_col].dt.date <= end) &
                (df[emp_col].isin(emp_list))
            )
            df_f = df.loc[mask].copy()
            q.put(('filtered_info', len(df_f)))

            # find missing employees
            present = set(df_f[emp_col].dropna())
            missing = sorted(set(emp_list) - present)

            # compute last update before start date for each missing
            df_before = df[df[date_col].dt.date < start]
            last_before = {}
            for emp in missing:
                df_e = df_before[df_before[emp_col] == emp]
                if df_e.empty:
                    last_before[emp] = "N/A"
                else:
                    d = df_e[date_col].max().date()
                    last_before[emp] = d.strftime("%m/%d/%Y")

            # store for email
            missing_map[list_name] = [
                {"employee": emp, "last_before": last_before[emp]}
                for emp in missing
            ]
            q.put(('missing_info', missing))

            # strip time from any datetime columns
            for col in df_f.columns:
                if pd.api.types.is_datetime64_any_dtype(df_f[col]):
                    df_f[col] = df_f[col].dt.date

            # write out filtered .xlsx (always at least headers)
            out_name = f"{list_name}_{ts}.xlsx"
            out_path = os.path.join(downloads, out_name)
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                if df_f.empty:
                    pd.DataFrame(columns=df.columns).to_excel(
                        writer,
                        sheet_name=list_name[:31],
                        index=False
                    )
                else:
                    df_f.to_excel(
                        writer,
                        sheet_name=list_name[:31],
                        index=False
                    )
            attachments.append(out_path)
            logger.info(f"Saved {out_name} ({len(df_f)}/{total_rows} rows)")

            done += 1
            q.put(('progress', done))

        # build email HTML with last-before column
        html = ['<html><body><h1>Missing Entries</h1>']
        for ln, items in missing_map.items():
            html.append(f"<h2>{ln}</h2>")
            if not items:
                html.append("<p>All employees reported in range.</p>")
            else:
                html.append(
                    "<table border='1' cellpadding='4'>"
                    "<tr><th>Employee</th><th>Last Update Before Start</th></tr>"
                )
                for it in items:
                    html.append(
                        f"<tr><td>{it['employee']}</td>"
                        f"<td>{it['last_before']}</td></tr>"
                    )
                html.append("</table>")
        html.append("</body></html>")
        body = "\n".join(html)

        # send via Outlook
        send_via_outlook(EMAIL_TO, EMAIL_CC, EMAIL_SUBJECT, body, attachments)

        q.put(('done', attachments, missing_map))
        logger.info("Worker finished successfully.")

    except Exception as e:
        logger.exception("Error in worker")
        q.put(('error', str(e)))

def start_process():
    start = start_cal.get_date()
    end   = end_cal.get_date()
    if start > end:
        return messagebox.showerror("Error", "Start must be on or before End.")
    btn_submit.config(state='disabled')
    progress_var.set("Initializing…")
    task_var.set(sheet_var.set(filt_var.set(miss_var.set(""))))
    prog_bar['value'] = 0
    threading.Thread(target=worker, args=(start, end, q), daemon=True).start()
    root.after(100, poll_queue)

def poll_queue():
    try:
        msg = q.get_nowait()
    except queue.Empty:
        return root.after(100, poll_queue)

    kind, *rest = msg

    if kind == 'init':
        total = rest[0]
        prog_bar['maximum'] = total
        progress_var.set(f"0 of {total} tasks")

    elif kind == 'task':
        task_var.set(f"Task: {rest[0]}")

    elif kind == 'sheet_info':
        sheet_var.set(f"Total rows: {rest[0]}")

    elif kind == 'filtered_info':
        filt_var.set(f"Filtered rows: {rest[0]}")

    elif kind == 'missing_info':
        lst = rest[0]
        miss_var.set("Missing: " + (", ".join(lst) if lst else "None"))

    elif kind == 'progress':
        done = rest[0]
        prog_bar['value'] = done
        progress_var.set(f"{done} of {prog_bar['maximum']} tasks")

    elif kind == 'done':
        btn_submit.config(state='normal')
        messagebox.showinfo("Finished", "Reports saved & email sent via Outlook.")

    elif kind == 'error':
        btn_submit.config(state='normal')
        messagebox.showerror("Error", rest[0])

    root.after(100, poll_queue)

# ─── BUILD UI ─────────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("CRIT Filter & Emailer (fast)")
root.geometry("480x360")
root.resizable(False, False)

q = queue.Queue()
padx, wrap = 10, 460

tk.Label(root, text="Start Date:").pack(anchor="w", padx=padx, pady=(10,0))
start_cal = DateEntry(root, date_pattern="mm/dd/yyyy"); start_cal.pack(padx=padx)

tk.Label(root, text="End Date:").pack(anchor="w", padx=padx, pady=(10,0))
end_cal   = DateEntry(root, date_pattern="mm/dd/yyyy"); end_cal.pack(padx=padx)

btn_submit = tk.Button(root, text="Submit", command=start_process,
                       bg="#4CAF50", fg="white",
                       font=("Segoe UI", 12, "bold"))
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