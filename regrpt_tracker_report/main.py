#!/usr/bin/env python3
import os
import datetime
import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
import pandas as pd
from openpyxl import load_workbook
import win32com.client as win32

# ─── USER CONFIGURATION ───────────────────────────────────────────────────────
SRC_FILE = r"C:\path\to\your\RegulatoryReporting.xlsm"

consumer_dataProvider   = ["Alice Johnson", "Bob Smith",     # …
                          ]
commercial_dataProvider = ["Carol Lee",      "Dan Patel",    # …
                          ]
scheduleOwner           = ["Eve Zhang",      "Frank Müller", # …
                          ]
CR360Transformation     = ["Grace Kim",      "Hiro Tanaka",  # …
                          ]

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

# Predefine email fields
EMAIL_TO      = ["to1@example.com"]
EMAIL_CC      = ["cc1@example.com"]
EMAIL_SUBJECT = "Automated Regulatory Reporting Summary"
# ─── END USER CONFIG ──────────────────────────────────────────────────────────

def send_via_outlook(to, cc, subject, html_body, attachments):
    outlook = win32.Dispatch('Outlook.Application')
    mail    = outlook.CreateItem(0)  # olMailItem
    mail.To      = ";".join(to)
    mail.CC      = ";".join(cc)
    mail.Subject = subject
    mail.HTMLBody = html_body
    for path in attachments:
        mail.Attachments.Add(path)
    mail.Send()

def main():
    now       = datetime.datetime.now()
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    timestamp = now.strftime("%Y%m%d_%H%M%S")

    root = tk.Tk()
    root.title("Regulatory Reporting Filter & Emailer")
    root.geometry("380x240")
    root.resizable(False, False)

    tk.Label(root, text="Source (predefined):").pack(anchor="w", padx=10, pady=(10,0))
    tk.Entry(root, width=50, state="readonly",
             textvariable=tk.StringVar(value=SRC_FILE)).pack(padx=10, pady=(0,10))

    tk.Label(root, text="Start date:").pack(anchor="w", padx=10)
    start_cal = DateEntry(root, date_pattern="mm/dd/yyyy"); start_cal.pack(padx=10)

    tk.Label(root, text="End date:").pack(anchor="w", padx=10, pady=(10,0))
    end_cal   = DateEntry(root, date_pattern="mm/dd/yyyy");   end_cal.pack(padx=10)

    def on_submit():
        start = start_cal.get_date()
        end   = end_cal.get_date()
        if start > end:
            messagebox.showerror("Error", "Start must be on or before End.")
            return

        attachments = []
        missing_info = {}

        for sheet_name, lists in SHEET_MAP.items():
            df_full = pd.read_excel(
                SRC_FILE,
                sheet_name=sheet_name,
                header=0,
                engine="openpyxl",
                parse_dates=[0]
            )
            date_col = df_full.columns[0]
            emp_col  = df_full.columns[4]

            for list_name, emp_list in lists:
                mask_in = (
                    (df_full[date_col].dt.date >= start) &
                    (df_full[date_col].dt.date <= end) &
                    (df_full[emp_col].isin(emp_list))
                )
                present = set(df_full.loc[mask_in, emp_col].dropna().unique())
                missing = sorted(set(emp_list) - present)

                info = []
                for emp in missing:
                    df_emp = df_full[df_full[emp_col] == emp]
                    if df_emp.empty:
                        info.append({"employee": emp, "last_date":"N/A", "last_row":"No record"})
                    else:
                        last_dt = df_emp[date_col].max().date()
                        idx0    = df_emp[df_emp[date_col].dt.date == last_dt].index[0]
                        excel_row = idx0 + 2
                        info.append({
                            "employee": emp,
                            "last_date": last_dt.strftime("%m/%d/%Y"),
                            "last_row": excel_row
                        })
                missing_info[list_name] = info

                # build filtered .xlsm preserving formatting
                wb = load_workbook(SRC_FILE, keep_vba=True)
                for s in wb.sheetnames:
                    if s != sheet_name:
                        wb.remove(wb[s])
                ws = wb[sheet_name]

                to_delete = []
                for r in range(2, ws.max_row+1):
                    cell = ws.cell(r,1).value
                    try: dt = pd.to_datetime(cell).date()
                    except: dt = None
                    emp = ws.cell(r,5).value
                    if dt is None or dt<start or dt>end or emp not in emp_list:
                        to_delete.append(r)
                for r in reversed(to_delete):
                    ws.delete_rows(r)

                out_name = f"{list_name}_{timestamp}.xlsm"
                out_path = os.path.join(downloads, out_name)
                wb.save(out_path)
                attachments.append(out_path)

        # build HTML email body
        html = ['<html><body>','<h1>Missing Entries Report</h1>']
        for list_name, rows in missing_info.items():
            html.append(f"<h2>{list_name}</h2>")
            if not rows:
                html.append("<p>All employees had ≥1 entry in range.</p>")
            else:
                html.append("<table border='1' cellpadding='4'>"
                            "<tr><th>Employee</th><th>Last Update</th><th>Last Row #</th></tr>")
                for r in rows:
                    html.append(
                        f"<tr><td>{r['employee']}</td>"
                        f"<td>{r['last_date']}</td>"
                        f"<td>{r['last_row']}</td></tr>"
                    )
                html.append("</table>")
        html.append("</body></html>")
        body = "\n".join(html)

        try:
            send_via_outlook(EMAIL_TO, EMAIL_CC, EMAIL_SUBJECT, body, attachments)
            messagebox.showinfo(
                "Done",
                "Reports saved to Downloads and sent via Outlook."
            )
        except Exception as e:
            messagebox.showerror("Email Error", str(e))

    tk.Button(
        root, text="Submit",
        command=on_submit,
        bg="#4CAF50", fg="white",
        font=("Segoe UI", 12, "bold")
    ).pack(pady=20, ipadx=10, ipady=5)

    root.mainloop()

if __name__ == "__main__":
    main()