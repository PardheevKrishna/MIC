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
        master.geometry('650x480')
        master.resizable(False, False)

        self.queue = queue.Queue()
        self.df_access = None

        width = 40
        style = ttk.Style(master)
        style.theme_use('clam')

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

        # Date picker
        ttk.Label(frm, text='Select Cutoff Date:').grid(row=4, column=0, sticky='w')
        self.date_entry = DateEntry(
            frm, width=12, background='darkblue', foreground='white', borderwidth=2,
            year=datetime.datetime.now().year
        )
        self.date_entry.grid(row=4, column=1, sticky='w')

        # Run button
        self.run_btn = ttk.Button(frm, text='Generate & Send', command=self._on_run)
        self.run_btn.grid(row=5, column=1, pady=20)

        # Progress spinner & ETA
        self.progress = ttk.Progressbar(frm, mode='indeterminate', length=500)
        self.progress.grid(row=6, column=0, columnspan=3, pady=(10, 0))
        self.time_var = tk.StringVar(value='Processed: 0 files | Elapsed: 00:00:00')
        ttk.Label(frm, textvariable=self.time_var).grid(row=7, column=0, columnspan=3, sticky='w')

        # Status
        self.status = ttk.Label(frm, text='', foreground='green', wraplength=600, justify='left')
        self.status.grid(row=8, column=0, columnspan=3, pady=(10, 0))

        # start queue polling
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
            return messagebox.showerror('Error', f'Could not read "Access Folder":\n{e}')

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
        logger.info("Owner: %s", owner)
        dfp = self.df_access[self.df_access['Entitlement Owner'] == owner]
        if dfp.empty:
            return
        row = dfp.iloc[0]
        to = row.get('Entitlement Owner Email', '') or ''
        cc = row.get('Delegate Email', '') or ''
        logger.debug("Setting To=%s, CC=%s", to, cc)
        self.email_var.set(to)
        self.cc_var.set(cc)

    def _on_run(self):
        self.run_btn.config(state='disabled')
        self.status.config(text='')
        self.progress.start(10)  # 10ms step
        logger.info("Starting scan thread")
        threading.Thread(target=self._scan_and_report, daemon=True).start()

    def _scan_and_report(self):
        logger.info("Background scan begins")
        owner   = self.person_var.get().strip()
        to_mail = self.email_var.get().strip()
        cc_mail = self.cc_var.get().strip()
        cutoff_d= self.date_entry.get_date()
        cutoff_ts = datetime.datetime.combine(cutoff_d, datetime.time.max).timestamp()
        now_ts    = time.time()

        dfp = self.df_access[self.df_access['Entitlement Owner'] == owner]
        bases = dfp['Full Path'].dropna().tolist()
        logger.info("Paths to scan: %s", bases)

        rows = []
        idx = 0
        start = time.time()

        def scan(p):
            try:
                with os.scandir(p) as it:
                    for ent in it:
                        if ent.is_dir(follow_symlinks=False):
                            yield from scan(ent.path)
                        elif ent.is_file(follow_symlinks=False):
                            yield ent
            except Exception as e:
                logger.warning("Skip dir %s: %s", p, e)

        for base in bases:
            for ent in scan(base):
                idx += 1
                try:
                    st = ent.stat(follow_symlinks=False)
                    if st.st_ctime > cutoff_ts:
                        continue
                    days = int((now_ts - st.st_ctime) // 86400)
                    rows.append({
                        'Person':        owner,
                        'File Name':     ent.name,
                        'File Path':     ent.path,
                        'Created':       datetime.datetime.fromtimestamp(st.st_ctime)
                                           .strftime('%Y-%m-%d %H:%M:%S'),
                        'Last Modified': datetime.datetime.fromtimestamp(st.st_mtime)
                                           .strftime('%Y-%m-%d %H:%M:%S'),
                        'Last Accessed': datetime.datetime.fromtimestamp(st.st_atime)
                                           .strftime('%Y-%m-%d %H:%M:%S'),
                        'Days Ago':      days
                    })
                except Exception as e:
                    logger.error("Error on %s: %s", ent.path, e)

                # real-time update for each file
                elapsed = time.time() - start
                elapsed_str = time.strftime('%H:%M:%S', time.gmtime(elapsed))
                self.queue.put(('update', idx, elapsed_str))

        # final elapsed
        total_elapsed = time.time() - start
        total_elapsed_str = time.strftime('%H:%M:%S', time.gmtime(total_elapsed))

        # Save Excel
        df_out = pd.DataFrame(rows)
        stamp  = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        fn     = f"{owner.replace(' ','_')}_report_{stamp}.xlsx"
        logger.info("Writing report %s (%d rows)", fn, len(rows))
        df_out.to_excel(fn, index=False, engine='openpyxl')
        wb = load_workbook(fn)
        ws = wb.active
        for col in ws.columns:
            ml = max((len(str(c.value)) for c in col if c.value), default=0)
            ws.column_dimensions[col[0].column_letter].width = ml + 2
        wb.save(fn)

        # Send Email
        status = f"Report saved: {fn}"
        if win32:
            try:
                logger.info("Emailing report to %s CC %s", to_mail, cc_mail)
                ol = win32.Dispatch('Outlook.Application')
                m  = ol.CreateItem(0)
                m.To      = to_mail
                m.CC      = cc_mail
                m.Subject = f"File Report for {owner}"
                m.HTMLBody= (
                    f"<p>Dear {owner},</p>"
                    f"<p>Files created on or before {cutoff_d}:</p>"
                    f"{df_out.to_html(index=False)}"
                    "<p>Regards,</p>"
                )
                m.Attachments.Add(os.path.abspath(fn))
                m.Send()
                status += f" & emailed to {to_mail}"
                logger.info("Email sent")
            except Exception as e:
                logger.exception("Email failure")
                status += f" (email failed: {e})"

        # signal done, include total elapsed
        self.queue.put(('done', idx, status, total_elapsed_str))
        logger.info("Scan complete, %d files matched", len(rows))

    def _process_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                typ = msg[0]
                if typ == 'update':
                    _, count, elapsed = msg
                    self.time_var.set(f"Processed: {count} files | Elapsed: {elapsed}")
                elif typ == 'done':
                    _, count, status, elapsed = msg
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