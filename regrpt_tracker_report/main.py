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
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

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
    ol = win32.Dispatch('Outlook.Application')
    mail = ol.CreateItem(0)
    mail.To      = ";".join(to)
    mail.CC      = ";".join(cc)
    mail.Subject = subject
    mail.HTMLBody = html_body
    for f in attachments:
        mail.Attachments.Add(f)
    mail.Send()

def worker(start, end, q):
    try:
        tasks = [
            (sheet, list_name, emp_list)
            for sheet, lists in SHEET_MAP.items()
            for list_name, emp_list in lists
        ]
        q.put(('init', len(tasks)))

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")

        # preload
        sheet_dfs = {
            sheet: pd.read_excel(
                SRC_FILE, sheet_name=sheet,
                engine="openpyxl", parse_dates=[0]
            )
            for sheet in SHEET_MAP
        }

        attachments = []
        missing_overall = {}
        done = 0

        for sheet, list_name, emp_list in tasks:
            desc = f"{sheet} → {list_name}"
            q.put(('task', desc))
            df = sheet_dfs[sheet]
            date_col = df.columns[0]
            emp_col  = df.columns[4]

            q.put(('sheet_info', len(df)))

            # filter
            mask = (
                (df[date_col].dt.date >= start) &
                (df[date_col].dt.date <= end) &
                (df[emp_col].isin(emp_list))
            )
            df_f = df.loc[mask].copy()
            q.put(('filtered_info', len(df_f)))

            present = set(df_f[emp_col].dropna().unique())
            missing = sorted(set(emp_list) - present)
            missing_overall[list_name] = missing
            q.put(('missing_info', missing))

            # strip time portion on date columns in df_f
            for idx in [0, 5, 6]:
                if idx < len(df_f.columns) and pd.api.types.is_datetime64_any_dtype(df_f[df_f.columns[idx]]):
                    df_f[df_f.columns[idx]] = df_f[df_f.columns[idx]].dt.date

            # write to .xlsx
            out_name = f"{list_name}_{ts}.xlsx"
            out_path = os.path.join(downloads, out_name)
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                df_f.to_excel(writer, sheet_name=list_name[:31], index=False)

            # openpyxl: format date-only and auto-fit columns
            wb2 = load_workbook(out_path)
            ws2 = wb2[list_name[:31]]
            for col_idx, col in enumerate(ws2.iter_cols(min_row=1, max_row=ws2.max_row), start=1):
                maxlen = 0
                col_letter = get_column_letter(col_idx)
                for cell in col:
                    val = cell.value
                    # if datetime or date, set to date-only number format
                    if isinstance(val, (datetime.datetime, datetime.date)):
                        cell.number_format = "M/D/YYYY"
                        val = val.date() if isinstance(val, datetime.datetime) else val
                    text = "" if val is None else str(val)
                    maxlen = max(maxlen, len(text))
                ws2.column_dimensions[col_letter].width = maxlen + 2
            wb2.save(out_path)

            attachments.append(out_path)
            done += 1
            q.put(('progress', done))

        # build email
        html = ['<html><body><h1>Missing Entries</h1>']
        for ln, miss in missing_overall.items():
            html.append(f"<h2>{ln}</h2>")
            if not miss:
                html.append("<p>All employees reported in range.</p>")
            else:
                html.append("<table border='1' cellpadding='4'>"
                            "<tr><th>Employee</th></tr>")
                for name in miss:
                    html.append(f"<tr><td>{name}</td></tr>")
                html.append("</table>")
        html.append("</body></html>")
        send_via_outlook(EMAIL_TO, EMAIL_CC, EMAIL_SUBJECT, "\n".join(html), attachments)

        q.put(('done', attachments, missing_overall))

    except Exception as e:
        q.put(('error', str(e)))

def start_process():
    start = start_cal.get_date()
    end   = end_cal.get_date()
    if start > end:
        return messagebox.showerror("Error", "Start must be ≤ End.")
    btn_submit.config(state='disabled')
    q.put(('init', 0))
    threading.Thread(target=worker, args=(start, end, q), daemon=True).start()
    root.after(100, poll_queue)

def poll_queue():
    try:
        kind, data = q.get_nowait()
    except queue.Empty:
        root.after(100, poll_queue)
        return

    if kind == 'init':
        total = data
        prog_bar['maximum'] = total
        progress_var.set(f"0 of {total} tasks")
    elif kind == 'task':
        task_var.set(f"Task: {data}")
    elif kind == 'sheet_info':
        sheet_var.set(f"Total rows: {data}")
    elif kind == 'filtered_info':
        filt_var.set(f"Filtered rows: {data}")
    elif kind == 'missing_info':
        miss_var.set("Missing: " + (", ".join(data) if data else "None"))
    elif kind == 'progress':
        prog_bar['value'] = data
        progress_var.set(f"{data} of {prog_bar['maximum']} tasks")
    elif kind == 'done':
        btn_submit.config(state='normal')
        messagebox.showinfo("Finished", "Reports saved & email sent.")
    elif kind == 'error':
        btn_submit.config(state='normal')
        messagebox.showerror("Error", data)

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