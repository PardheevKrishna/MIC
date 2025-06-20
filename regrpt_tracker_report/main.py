#!/usr/bin/env python3
import os
import datetime
import threading
import queue
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkcalendar import DateEntry
import pandas as pd
from openpyxl import load_workbook
import win32com.client as win32

# ─── USER CONFIG ───────────────────────────────────────────────────────────────
SRC_FILE = r"C:\path\to\your\CRIT_Master.xlsm"

consumer_dataProvider   = ["Alice Johnson", "Bob Smith"]
commercial_dataProvider = ["Carol Lee", "Dan Patel"]
scheduleOwner           = ["Eve Zhang", "Frank Müller"]
CR360Transformation     = ["Grace Kim", "Hiro Tanaka"]

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
# ─── END USER CONFIG ───────────────────────────────────────────────────────────

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
    """
    Runs in background thread. Reports:
      ('init', total_tasks)
      ('progress', done_count)
      ('done', attachments, missing_info)
      ('error', message)
    """
    try:
        # Build tasks (sheet_name, list_name, emp_list)
        tasks = []
        for sheet, lsts in SHEET_MAP.items():
            for list_name, emp_list in lsts:
                tasks.append((sheet, list_name, emp_list))
        total = len(tasks)
        q.put(('init', total))

        now = datetime.datetime.now()
        ts  = now.strftime("%Y%m%d_%H%M%S")
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")

        attachments = []
        missing_info = {}

        # Load each sheet once (pandas) into dict
        sheet_dfs = {}
        for sheet, _ in SHEET_MAP.items():
            sheet_dfs[sheet] = pd.read_excel(
                SRC_FILE, sheet_name=sheet,
                engine="openpyxl", header=0,
                parse_dates=[0]
            )

        done = 0
        for sheet, list_name, emp_list in tasks:
            df_full = sheet_dfs[sheet]
            date_col = df_full.columns[0]
            emp_col  = df_full.columns[4]

            # in-range mask
            mask = ((df_full[date_col].dt.date >= start) &
                    (df_full[date_col].dt.date <= end) &
                    (df_full[emp_col].isin(emp_list)))
            present = set(df_full.loc[mask, emp_col].dropna().unique())
            missing = sorted(set(emp_list) - present)

            info = []
            for emp in missing:
                df_emp = df_full[df_full[emp_col] == emp]
                if df_emp.empty:
                    info.append({"employee": emp,
                                 "last_date": "N/A",
                                 "last_row":  "No record"})
                else:
                    last_dt = df_emp[date_col].max().date()
                    idx0    = df_emp[df_emp[date_col].dt.date == last_dt].index[0]
                    excel_row = idx0 + 2
                    info.append({"employee": emp,
                                 "last_date": last_dt.strftime("%m/%d/%Y"),
                                 "last_row":  excel_row})
            missing_info[list_name] = info

            # Now openpyxl to preserve macros/formatting
            wb = load_workbook(SRC_FILE, keep_vba=True)
            # remove other sheets
            for s in wb.sheetnames:
                if s != sheet:
                    wb.remove(wb[s])
            ws = wb[sheet]
            # delete rows outside range or not in emp_list
            to_del = []
            for r in range(2, ws.max_row+1):
                cell = ws.cell(r,1).value
                try:
                    dt = pd.to_datetime(cell).date()
                except:
                    dt = None
                emp = ws.cell(r,5).value
                if dt is None or dt < start or dt > end or emp not in emp_list:
                    to_del.append(r)
            for r in reversed(to_del):
                ws.delete_rows(r)

            out_name = f"{list_name}_{ts}.xlsm"
            out_path = os.path.join(downloads, out_name)
            wb.save(out_path)
            attachments.append(out_path)

            done += 1
            q.put(('progress', done))

        # build HTML
        html = ['<html><body><h1>Missing Entries</h1>']
        for ln, rows in missing_info.items():
            html.append(f"<h2>{ln}</h2>")
            if not rows:
                html.append("<p>All employees had entries in range.</p>")
            else:
                html.append("<table border='1' cellpadding='4'>"
                            "<tr><th>Employee</th><th>Last Update</th><th>Last Row</th></tr>")
                for r in rows:
                    html.append(f"<tr><td>{r['employee']}</td>"
                                f"<td>{r['last_date']}</td>"
                                f"<td>{r['last_row']}</td></tr>")
                html.append("</table>")
        html.append("</body></html>")
        body = "\n".join(html)

        # send mail
        send_via_outlook(EMAIL_TO, EMAIL_CC, EMAIL_SUBJECT, body, attachments)
        q.put(('done', attachments, missing_info))

    except Exception as e:
        q.put(('error', str(e)))

def start_process():
    start = start_cal.get_date()
    end   = end_cal.get_date()
    if start > end:
        messagebox.showerror("Error", "Start must be ≤ End.")
        return

    # disable UI
    btn_submit.config(state='disabled')
    progress_var.set("Preparing…")
    prog_bar['value'] = 0

    # kick off worker
    threading.Thread(target=worker, args=(start, end, q), daemon=True).start()
    root.after(100, poll_queue)

def poll_queue():
    try:
        msg = q.get_nowait()
    except queue.Empty:
        root.after(100, poll_queue)
        return

    kind = msg[0]
    if kind == 'init':
        total = msg[1]
        prog_bar['maximum'] = total
        progress_var.set(f"0 of {total} done")
        root.after(100, poll_queue)

    elif kind == 'progress':
        done = msg[1]
        total = prog_bar['maximum']
        prog_bar['value'] = done
        progress_var.set(f"{done} of {total} done")
        root.after(100, poll_queue)

    elif kind == 'done':
        attachments, missing_info = msg[1], msg[2]
        btn_submit.config(state='normal')
        progress_var.set("All done!")
        messagebox.showinfo(
            "Finished",
            f"Saved {len(attachments)} reports to Downloads\n"
            f"Email sent via Outlook."
        )

    elif kind == 'error':
        err = msg[1]
        btn_submit.config(state='normal')
        progress_var.set("Error.")
        messagebox.showerror("Error in processing thread", err)

# ─── BUILD UI ─────────────────────────────────────────────────────────────────
root = tk.Tk()
root.title("CRIT Filter & Emailer")
root.geometry("400x300")
root.resizable(False, False)

q = queue.Queue()

tk.Label(root, text="Source (fixed):").pack(anchor="w", padx=10, pady=(10,0))
tk.Entry(root, width=50, state="readonly",
         textvariable=tk.StringVar(value=SRC_FILE)).pack(padx=10)

tk.Label(root, text="Start Date:").pack(anchor="w", padx=10, pady=(10,0))
start_cal = DateEntry(root, date_pattern="mm/dd/yyyy"); start_cal.pack(padx=10)

tk.Label(root, text="End Date:").pack(anchor="w", padx=10, pady=(10,0))
end_cal   = DateEntry(root, date_pattern="mm/dd/yyyy"); end_cal.pack(padx=10)

btn_submit = tk.Button(root, text="Submit", command=start_process,
                       bg="#4CAF50", fg="white",
                       font=("Segoe UI", 12, "bold"))
btn_submit.pack(pady=15)

progress_var = tk.StringVar(value="Idle")
tk.Label(root, textvariable=progress_var).pack(pady=(5,0))

prog_bar = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
prog_bar.pack(pady=(5,10))

root.mainloop()