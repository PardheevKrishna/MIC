import os
import time
import datetime
import threading
import queue
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Requires: pandas, openpyxl, pywin32, tkcalendar
try:
    import win32com.client as win32
except ImportError:
    win32 = None

from openpyxl import load_workbook
from tkcalendar import DateEntry

class FileReportApp:
    def __init__(self, master):
        self.master = master
        master.title('Access Folder File Report')
        master.geometry('650x500')
        master.resizable(False, False)

        width_value = 40
        style = ttk.Style(master)
        style.theme_use('clam')

        # queue for cross-thread progress updates
        self.queue = queue.Queue()

        main = ttk.Frame(master, padding=10)
        main.pack(fill='both', expand=True)

        # Excel selector
        ttk.Label(main, text='Access Excel File:').grid(row=0, column=0, sticky='w')
        self.excel_path = tk.StringVar()
        ttk.Entry(main, textvariable=self.excel_path, width=width_value).grid(row=0, column=1, sticky='w')
        ttk.Button(main, text='Browse…', command=self._browse_excel).grid(row=0, column=2, padx=5)

        # Person dropdown
        ttk.Label(main, text='Select Person:').grid(row=1, column=0, pady=10, sticky='w')
        self.person_var = tk.StringVar()
        self.person_cb = ttk.Combobox(main, textvariable=self.person_var, state='readonly', width=width_value)
        self.person_cb.grid(row=1, column=1, columnspan=2, sticky='w')
        self.person_cb.bind('<<ComboboxSelected>>', self._person_selected)

        # To and CC
        ttk.Label(main, text='To (Email):').grid(row=2, column=0, sticky='w')
        self.email_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.email_var, width=width_value).grid(row=2, column=1, columnspan=2, sticky='w')
        ttk.Label(main, text='CC:').grid(row=3, column=0, pady=10, sticky='w')
        self.cc_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.cc_var, width=width_value).grid(row=3, column=1, columnspan=2, sticky='w')

        # Date selector
        ttk.Label(main, text='Select Start Date:').grid(row=4, column=0, sticky='w')
        self.date_entry = DateEntry(main, width=12, background='darkblue',
                                    foreground='white', borderwidth=2,
                                    year=datetime.datetime.now().year)
        self.date_entry.grid(row=4, column=1, sticky='w')

        # Run button
        self.run_btn = ttk.Button(main, text='Generate & Send', command=self._on_run)
        self.run_btn.grid(row=5, column=1, pady=20)

        # Progress bar & time estimate
        self.progress = ttk.Progressbar(main, orient='horizontal', length=500, mode='determinate')
        self.progress.grid(row=6, column=0, columnspan=3, pady=(10,0))
        self.time_var = tk.StringVar(value='Estimated time left: N/A')
        ttk.Label(main, textvariable=self.time_var).grid(row=7, column=0, columnspan=3, sticky='w')

        # Status message
        self.status = ttk.Label(main, text='', foreground='green', wraplength=600, justify='left')
        self.status.grid(row=8, column=0, columnspan=3, pady=(10,0))

        # data holders
        self.df_access = None

        # start the queue-poll loop
        self.master.after(100, self._process_queue)

    def _browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[('Excel files','*.xlsx *.xls')])
        if not path: return
        self.excel_path.set(path)

        try:
            df = pd.read_excel(path, sheet_name='Access Folder')
        except Exception as e:
            messagebox.showerror('Error', f'Could not read "Access Folder" sheet:\n{e}')
            return

        self.df_access = df
        owners = sorted(df['Entitlement Owner'].dropna().unique())
        self.person_cb['values'] = owners
        if owners:
            self.person_cb.current(0)
            self._person_selected()

    def _person_selected(self, event=None):
        if self.df_access is None: return
        owner = self.person_var.get()
        dfp = self.df_access[self.df_access['Entitlement Owner']==owner]
        if not dfp.empty:
            self.email_var.set(dfp['Entitlement Owner Email'].iloc[0] or '')
            self.cc_var.set(dfp.get('Delegate Email', pd.Series()).iloc[0] or '')

    def _on_run(self):
        # disable UI
        self.run_btn.config(state='disabled')
        self.status.config(text='')
        self.progress['value'] = 0
        self.time_var.set('Estimated time left: N/A')

        # fire background thread
        threading.Thread(target=self._process_files, daemon=True).start()

    def _process_files(self):
        """Runs in background thread: two-pass walk, build report, send email."""
        path = self.excel_path.get().strip()
        owner = self.person_var.get().strip()
        to_email = self.email_var.get().strip()
        cc_email = self.cc_var.get().strip()
        start_date = self.date_entry.get_date()

        dfp = self.df_access[self.df_access['Entitlement Owner']==owner]
        base_paths = dfp['Full Path'].dropna().tolist()

        # 1) count total files
        total = 0
        for base in base_paths:
            for _, _, files in os.walk(base or ''):
                total += len(files)
        self.queue.put(('init', total))

        # 2) process each file
        now = datetime.datetime.now()
        cutoff = datetime.datetime.combine(start_date, datetime.time.min)
        start_time = time.time()
        rows = []
        idx = 0

        for base in base_paths:
            for root, _, files in os.walk(base or ''):
                for fname in files:
                    idx += 1
                    fp = os.path.join(root, fname)
                    try:
                        cts = datetime.datetime.fromtimestamp(os.path.getctime(fp))
                        if not (cutoff <= cts <= now):
                            raise ValueError()
                        mts = datetime.datetime.fromtimestamp(os.path.getmtime(fp))
                        ats = datetime.datetime.fromtimestamp(os.path.getatime(fp))
                        days_ago = (now - cts).days
                        rows.append({
                            'Person':        owner,
                            'File Name':     fname,
                            'File Path':     fp,
                            'Created':       cts.strftime('%Y-%m-%d %H:%M:%S'),
                            'Last Modified': mts.strftime('%Y-%m-%d %H:%M:%S'),
                            'Last Accessed': ats.strftime('%Y-%m-%d %H:%M:%S'),
                            'Days Ago':      days_ago
                        })
                    except Exception:
                        # skip missing/permission/out-of-range
                        pass

                    elapsed = time.time() - start_time
                    avg = elapsed / idx
                    rem  = avg * (total - idx)
                    self.queue.put(('progress', idx, rem))

        # 3) build DataFrame & save
        out_df = pd.DataFrame(rows)
        stamp = now.strftime('%Y%m%d_%H%M%S')
        safe  = owner.replace(' ', '_')
        report_fn = f"{safe}_report_{stamp}.xlsx"

        out_df.to_excel(report_fn, index=False, engine='openpyxl')
        wb = load_workbook(report_fn)
        ws = wb.active
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=0)
            ws.column_dimensions[col[0].column_letter].width = max_len + 2
        wb.save(report_fn)

        # 4) send Outlook email
        status_msg = f"Report saved as {report_fn}"
        if win32:
            try:
                outlook = win32.Dispatch('Outlook.Application')
                m = outlook.CreateItem(0)
                m.To      = to_email
                m.CC      = cc_email
                m.Subject = f"File Report for {owner}"
                m.HTMLBody = (
                    f"<p>Dear {owner},</p>"
                    f"<p>Files from {start_date} to now:</p>"
                    f"{out_df.to_html(index=False)}"
                    "<p>Regards,</p>"
                )
                m.Attachments.Add(os.path.abspath(report_fn))
                m.Send()
                status_msg += f"\nand emailed to {to_email}"
                if cc_email:
                    status_msg += f" (CC: {cc_email})"
            except Exception as e:
                status_msg += f"\n(Email failed: {e})"

        # signal done
        self.queue.put(('done', status_msg))

    def _process_queue(self):
        """Poll for progress updates and apply them on the main thread."""
        try:
            while True:
                msg = self.queue.get_nowait()
                typ = msg[0]
                if typ == 'init':
                    _, total = msg
                    self.progress['maximum'] = total
                elif typ == 'progress':
                    _, idx, rem = msg
                    self.progress['value'] = idx
                    self.time_var.set(f"Estimated time left: {int(rem)}s")
                elif typ == 'done':
                    _, status_msg = msg
                    self.status.config(text=status_msg)
                    self.run_btn.config(state='normal')
        except queue.Empty:
            pass
        finally:
            self.master.after(100, self._process_queue)

if __name__ == '__main__':
    root = tk.Tk()
    FileReportApp(root)
    root.mainloop()