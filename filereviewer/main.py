import os
import time
import datetime
import threading
import queue
import logging
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

# --- Logging configuration ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

class FileReportApp:
    def __init__(self, master):
        logger.info("Starting GUI")
        self.master = master
        master.title('Access Folder File Report')
        master.geometry('650x540')
        master.resizable(False, False)

        self.queue = queue.Queue()
        self.df_access = None

        width = 40
        ttk.Style(master).theme_use('clam')

        frm = ttk.Frame(master, padding=10)
        frm.pack(fill='both', expand=True)

        # Excel picker
        ttk.Label(frm, text='Access Excel File:').grid(row=0, column=0, sticky='w')
        self.excel_path = tk.StringVar()
        ttk.Entry(frm, textvariable=self.excel_path, width=width).grid(row=0, column=1, sticky='w')
        ttk.Button(frm, text='Browse…', command=self._browse_excel).grid(row=0, column=2, padx=5)

        # Person selector
        ttk.Label(frm, text='Select Person:').grid(row=1, column=0, pady=10, sticky='w')
        self.person_var = tk.StringVar()
        self.person_cb = ttk.Combobox(frm, textvariable=self.person_var,
                                      state='readonly', width=width)
        self.person_cb.grid(row=1, column=1, columnspan=2, sticky='w')
        self.person_cb.bind('<<ComboboxSelected>>', self._person_selected)
        self.person_var.trace_add('write', lambda *a: self._person_selected())

        # To / CC
        ttk.Label(frm, text='To (Email):').grid(row=2, column=0, sticky='w')
        self.email_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.email_var, width=width).grid(row=2, column=1, columnspan=2, sticky='w')
        ttk.Label(frm, text='CC:').grid(row=3, column=0, pady=10, sticky='w')
        self.cc_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.cc_var, width=width).grid(row=3, column=1, columnspan=2, sticky='w')

        # Date picker + helper text
        ttk.Label(frm, text='Select Cutoff Date:').grid(row=4, column=0, sticky='w')
        self.date_entry = DateEntry(
            frm, width=12, background='darkblue', foreground='white', borderwidth=2,
            year=datetime.datetime.now().year
        )
        self.date_entry.grid(row=4, column=1, sticky='w')
        ttk.Label(
            frm,
            text='(Includes all files created on or before this date)',
            font=("TkDefaultFont", 8)
        ).grid(row=5, column=1, columnspan=2, sticky='w', pady=(0,10))

        # Run button
        self.run_btn = ttk.Button(frm, text='Generate & Send', command=self._on_run)
        self.run_btn.grid(row=6, column=1, pady=10)

        # Progress spinner & ETA
        self.progress = ttk.Progressbar(frm, mode='indeterminate', length=500)
        self.progress.grid(row=7, column=0, columnspan=3, pady=(10, 0))
        self.time_var = tk.StringVar(value='Processed: 0 files | Elapsed: 00:00:00')
        ttk.Label(frm, textvariable=self.time_var).grid(row=8, column=0, columnspan=3, sticky='w')

        # Status
        self.status = ttk.Label(frm, text='', foreground='green', wraplength=600, justify='left')
        self.status.grid(row=9, column=0, columnspan=3, pady=(10, 0))

        self.master.after(100, self._process_queue)

    def _browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[('Excel files','*.xlsx *.xls')])
        if not path:
            return
        logger.info("Excel selected: %s", path)
        self.excel_path.set(path)
        try:
            df = pd.read_excel(path, sheet_name='Access Folder')
        except Exception as e:
            logger.exception("Failed to read 'Access Folder' sheet")
            messagebox.showerror('Error', f'Could not read "Access Folder":\n{e}')
            return

        self.df_access = df
        owners = sorted(df['Entitlement Owner'].dropna().unique())
        logger.info("Owners found: %s", owners)
        self.person_cb['values'] = owners
        if owners:
            self.person_var.set(owners[0])
            self._person_selected()

    def _person_selected(self, event=None):
        if self.df_access is None:
            return
        owner = self.person_var.get()
        dfp = self.df_access[self.df_access['Entitlement Owner'] == owner]
        if dfp.empty:
            return
        row = dfp.iloc[0]
        self.email_var.set(row.get('Entitlement Owner Email','') or '')
        self.cc_var.set(row.get('Delegate Email','') or '')

    def _format_size(self, size_bytes):
        for unit in ['B','KB','MB','GB','TB']:
            if size_bytes < 1024:
                return f"{size_bytes:.2f} {unit}"
            size_bytes /= 1024
        return f"{size_bytes:.2f} PB"

    def _on_run(self):
        self.run_btn.config(state='disabled')
        self.status.config(text='')
        self.progress.start(10)
        threading.Thread(target=self._scan_and_report, daemon=True).start()

    def _scan_and_report(self):
        owner     = self.person_var.get().strip()
        to_mail   = self.email_var.get().strip()
        cc_mail   = self.cc_var.get().strip()
        cutoff_d  = self.date_entry.get_date()
        cutoff_ts = datetime.datetime.combine(cutoff_d, datetime.time.max).timestamp()
        now_ts    = time.time()

        bases      = self.df_access[self.df_access['Entitlement Owner']==owner]['Full Path'].dropna().tolist()
        rows       = []
        errors     = []
        idx        = 0
        start_time = time.time()
        total_size = 0

        def scan_dir(path):
            try:
                with os.scandir(path) as it:
                    for ent in it:
                        if ent.is_dir(follow_symlinks=False):
                            yield from scan_dir(ent.path)
                        elif ent.is_file(follow_symlinks=False):
                            yield ent
            except Exception as e:
                errors.append({
                    'Entitlement Owner': owner,
                    'File Path': path,
                    'Error Message': str(e)
                })

        for base in bases:
            for ent in scan_dir(base):
                idx += 1
                try:
                    st = ent.stat(follow_symlinks=False)
                    if st.st_ctime > cutoff_ts:
                        continue
                    size = st.st_size
                    total_size += size
                    days = int((now_ts - st.st_ctime)//86400)
                    rows.append({
                        'Entitlement Owner': owner,
                        'File Name':         ent.name,
                        'File Path':         ent.path,
                        'Created':           datetime.datetime.fromtimestamp(st.st_ctime)
                                                  .strftime('%Y-%m-%d %H:%M:%S'),
                        'Last Modified':     datetime.datetime.fromtimestamp(st.st_mtime)
                                                  .strftime('%Y-%m-%d %H:%M:%S'),
                        'Last Accessed':     datetime.datetime.fromtimestamp(st.st_atime)
                                                  .strftime('%Y-%m-%d %H:%M:%S'),
                        'Days Ago':          days,
                        'Size (MB)':         round(size / (1024*1024), 2)
                    })
                except Exception as e:
                    errors.append({
                        'Entitlement Owner': owner,
                        'File Path': ent.path,
                        'Error Message': str(e)
                    })

                # real-time update
                elapsed = time.time() - start_time
                elapsed_str = time.strftime('%H:%M:%S', time.gmtime(elapsed))
                self.queue.put(('update', idx, elapsed_str))

        total_elapsed = time.time() - start_time
        elapsed_str   = time.strftime('%H:%M:%S', time.gmtime(total_elapsed))
        total_files   = len(rows)

        # Write Excel with two sheets
        df_main = pd.DataFrame(rows)
        df_err  = pd.DataFrame(errors)
        stamp   = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        fn      = f"{owner.replace(' ','_')}_report_{stamp}.xlsx"

        with pd.ExcelWriter(fn, engine='openpyxl') as writer:
            df_main.to_excel(writer, sheet_name='Entitlement Files', index=False)
            if not df_err.empty:
                df_err.to_excel(writer, sheet_name='Access Errors', index=False)

        # auto-fit columns on both sheets
        wb = load_workbook(fn)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value), default=0)
                ws.column_dimensions[col[0].column_letter].width = max_len + 2
        wb.save(fn)

        # Email summary
        status = f"Report saved: {fn}"
        if win32:
            try:
                mail = win32.Dispatch('Outlook.Application').CreateItem(0)
                for addr in to_mail.split(';'):
                    r = mail.Recipients.Add(addr.strip()); r.Type = 1; r.Resolve()
                for addr in cc_mail.split(';'):
                    r = mail.Recipients.Add(addr.strip()); r.Type = 2; r.Resolve()
                mail.Subject = f"File Report for {owner}"
                summary = (
                    f"<p>Dear {owner},</p>"
                    "<p>Please find attached the detailed report.</p>"
                    "<p>Summary:</p><ul>"
                    f"<li>Base folders searched: {', '.join(bases)}</li>"
                    f"<li>Number of files found: {total_files}</li>"
                    f"<li>Total size: {self._format_size(total_size)}</li>"
                    f"<li>Cutoff date: {cutoff_d}</li>"
                    f"<li>Scan duration: {elapsed_str}</li>"
                    "</ul><p>Regards,</p>"
                )
                mail.HTMLBody = summary
                mail.Attachments.Add(os.path.abspath(fn))
                mail.Recipients.ResolveAll()
                mail.Send()
                status += " & emailed"
            except Exception as e:
                logger.error("Email failed: %s", e)
                status += " (email failed)"

        self.queue.put(('done', total_files, status, elapsed_str))

    def _process_queue(self):
        try:
            while True:
                typ, *data = self.queue.get_nowait()
                if typ == 'update':
                    count, elapsed = data
                    self.time_var.set(f"Processed: {count} files | Elapsed: {elapsed}")
                elif typ == 'done':
                    count, status, elapsed = data
                    self.progress.stop()
                    self.status.config(text=status)
                    self.run_btn.config(state='normal')
                    self.time_var.set(f"Total processed: {count} | Elapsed: {elapsed}")
        except queue.Empty:
            pass
        finally:
            self.master.after(100, self._process_queue)

if __name__ == '__main__':
    logger.info("Launching app")
    root = tk.Tk()
    FileReportApp(root)
    root.mainloop()
    logger.info("App closed")