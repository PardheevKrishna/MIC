import os
import time
import datetime
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Requires: pandas, openpyxl, pywin32, tkcalendar, tqdm
try:
    import win32com.client as win32
except ImportError:
    win32 = None

from openpyxl import load_workbook
from tkcalendar import DateEntry
from tqdm import tqdm

class FileReportApp:
    def __init__(self, master):
        self.master = master
        master.title('Access Folder File Report')
        master.geometry('650x480')
        master.resizable(False, False)

        width_value = 40
        style = ttk.Style()
        style.theme_use('clam')

        main = ttk.Frame(master, padding=10)
        main.pack(fill='both', expand=True)

        # --- Excel selector ---
        ttk.Label(main, text='Access Excel File:').grid(row=0, column=0, sticky='w')
        self.excel_path = tk.StringVar()
        ttk.Entry(main, textvariable=self.excel_path, width=width_value).grid(row=0, column=1, sticky='w')
        ttk.Button(main, text='Browse…', command=self._browse_excel).grid(row=0, column=2, padx=5)

        # --- Person dropdown ---
        ttk.Label(main, text='Select Person:').grid(row=1, column=0, pady=10, sticky='w')
        self.person_var = tk.StringVar()
        self.person_cb = ttk.Combobox(main, textvariable=self.person_var,
                                      state='readonly', width=width_value)
        self.person_cb.grid(row=1, column=1, columnspan=2, sticky='w')
        self.person_cb.bind('<<ComboboxSelected>>', self._person_selected)

        # --- Entitlement Owner Email ---
        ttk.Label(main, text='To (Email):').grid(row=2, column=0, sticky='w')
        self.email_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.email_var, width=width_value).grid(row=2, column=1, columnspan=2, sticky='w')

        # --- Delegate Email CC ---
        ttk.Label(main, text='CC:').grid(row=3, column=0, pady=10, sticky='w')
        self.cc_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.cc_var, width=width_value).grid(row=3, column=1, columnspan=2, sticky='w')

        # --- Date selector ---
        ttk.Label(main, text='Select Start Date:').grid(row=4, column=0, sticky='w')
        self.date_entry = DateEntry(main, width=12, background='darkblue',
                                    foreground='white', borderwidth=2,
                                    year=datetime.datetime.now().year)
        self.date_entry.grid(row=4, column=1, sticky='w')

        # --- Run button ---
        self.run_btn = ttk.Button(main, text='Generate & Send', command=self._generate_and_send)
        self.run_btn.grid(row=5, column=1, pady=20)

        # --- Progress bar and time label ---
        self.progress = ttk.Progressbar(main, orient='horizontal', length=500, mode='determinate')
        self.progress.grid(row=6, column=0, columnspan=3, pady=(10, 0))
        self.time_var = tk.StringVar(value='Estimated time left: N/A')
        ttk.Label(main, textvariable=self.time_var).grid(row=7, column=0, columnspan=3, sticky='w')

        # --- Status ---
        self.status = ttk.Label(main, text='', foreground='green',
                                wraplength=600, justify='left')
        self.status.grid(row=8, column=0, columnspan=3, pady=(10, 0))

        self.df_access = None

    def _browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[('Excel files','*.xlsx *.xls')])
        if not path:
            return
        self.excel_path.set(path)
        try:
            df = pd.read_excel(path, sheet_name='Access Folder')
            self.df_access = df
            owners = sorted(df['Entitlement Owner'].dropna().unique())
            self.person_cb['values'] = owners
            if owners:
                self.person_cb.current(0)
                self._person_selected()
        except Exception as e:
            messagebox.showerror('Error', f'Could not read "Access Folder" sheet:\n{e}')

    def _person_selected(self, event=None):
        if self.df_access is None:
            return
        owner = self.person_var.get()
        df_person = self.df_access[self.df_access['Entitlement Owner'] == owner]
        if not df_person.empty:
            self.email_var.set(df_person['Entitlement Owner Email'].iloc[0] or '')
            self.cc_var.set(df_person.get('Delegate Email', pd.Series()).iloc[0] or '')

    def _generate_and_send(self):
        path = self.excel_path.get().strip()
        owner = self.person_var.get().strip()
        to_email = self.email_var.get().strip()
        cc_email = self.cc_var.get().strip()
        start_date = self.date_entry.get_date()
        if not path or not owner or not to_email:
            messagebox.showerror('Missing Data', 'Please select Excel, person, and To email.')
            return

        df = self.df_access
        df_person = df[df['Entitlement Owner'] == owner]
        base_paths = df_person['Full Path'].dropna().tolist()

        # collect all file paths
        file_list = []
        for base in base_paths:
            for root, dirs, files in os.walk(base or ''):
                for f in files:
                    file_list.append(os.path.join(root, f))

        total = len(file_list)
        if total == 0:
            messagebox.showinfo('No files', 'No files found under given paths.')
            return

        now = datetime.datetime.now()
        cutoff = datetime.datetime.combine(start_date, datetime.time.min)
        rows = []

        self.progress['maximum'] = total
        start_time = time.time()

        for idx, fp in enumerate(tqdm(file_list, desc='Processing files'), start=1):
            try:
                cts = datetime.datetime.fromtimestamp(os.path.getctime(fp))
                # filter by date range
                if not (cutoff <= cts <= now):
                    raise ValueError('out of range')
                mts = datetime.datetime.fromtimestamp(os.path.getmtime(fp))
                ats = datetime.datetime.fromtimestamp(os.path.getatime(fp))
                days_ago = (now - cts).days
                rows.append({
                    'Person':         owner,
                    'File Name':      os.path.basename(fp),
                    'File Path':      fp,
                    'Created':        cts.strftime('%Y-%m-%d %H:%M:%S'),
                    'Last Modified':  mts.strftime('%Y-%m-%d %H:%M:%S'),
                    'Last Accessed':  ats.strftime('%Y-%m-%d %H:%M:%S'),
                    'Days Ago':       days_ago
                })
            except Exception:
                # skip files with errors or out of date range
                pass

            # update progress & time estimate
            elapsed = time.time() - start_time
            avg = elapsed / idx
            rem = avg * (total - idx)
            self.time_var.set(f"Estimated time left: {int(rem)}s")
            self.progress['value'] = idx
            self.master.update_idletasks()

        if not rows:
            messagebox.showinfo('No Matches', 'No files matched the date criteria.')
            return

        # build and save report
        out_df = pd.DataFrame(rows)
        stamp = now.strftime('%Y%m%d_%H%M%S')
        safe = owner.replace(' ', '_')
        report_fn = f"{safe}_report_{stamp}.xlsx"

        out_df.to_excel(report_fn, index=False, engine='openpyxl')
        wb = load_workbook(report_fn)
        ws = wb.active
        for col in ws.columns:
            max_length = max((len(str(cell.value)) for cell in col if cell.value), default=0)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2
        wb.save(report_fn)

        # send email
        status_msg = f"Report saved as {report_fn}"
        if win32:
            try:
                outlook = win32.Dispatch('Outlook.Application')
                mail = outlook.CreateItem(0)
                mail.To = to_email
                mail.CC = cc_email
                mail.Subject = f"File Report for {owner}"
                html = out_df.to_html(index=False)
                mail.HTMLBody = (
                    f"<p>Dear {owner},</p>"
                    f"<p>Please find files from {start_date} to now:</p>"
                    f"{html}"
                    "<p>Regards,</p>"
                )
                mail.Attachments.Add(os.path.abspath(report_fn))
                mail.Send()
                status_msg += f"\nand emailed to {to_email}"
                if cc_email:
                    status_msg += f" (CC: {cc_email})"
            except Exception as e:
                messagebox.showerror('Email Error', str(e))
        else:
            messagebox.showwarning('Outlook Unavailable',
                                   'pywin32 not installed: report saved only.')

        self.status.config(text=status_msg)


if __name__ == '__main__':
    root = tk.Tk()
    FileReportApp(root)
    root.mainloop()