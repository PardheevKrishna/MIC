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
        logger.info("Initializing FileReportApp")
        self.master = master
        master.title('Access Folder File Report')
        master.geometry('650x500')
        master.resizable(False, False)

        self.queue = queue.Queue()
        self.df_access = None

        width_value = 40
        style = ttk.Style(master)
        style.theme_use('clam')

        main = ttk.Frame(master, padding=10)
        main.pack(fill='both', expand=True)

        # Excel file picker
        ttk.Label(main, text='Access Excel File:').grid(row=0, column=0, sticky='w')
        self.excel_path = tk.StringVar()
        ttk.Entry(main, textvariable=self.excel_path, width=width_value).grid(row=0, column=1, sticky='w')
        ttk.Button(main, text='Browse…', command=self._browse_excel).grid(row=0, column=2, padx=5)

        # Person dropdown
        ttk.Label(main, text='Select Person:').grid(row=1, column=0, pady=10, sticky='w')
        self.person_var = tk.StringVar()
        self.person_cb = ttk.Combobox(main, textvariable=self.person_var,
                                      state='readonly', width=width_value)
        self.person_cb.grid(row=1, column=1, columnspan=2, sticky='w')
        self.person_cb.bind('<<ComboboxSelected>>', self._person_selected)
        self.person_var.trace_add('write', lambda *args: self._person_selected())

        # To / CC
        ttk.Label(main, text='To (Email):').grid(row=2, column=0, sticky='w')
        self.email_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.email_var, width=width_value).grid(row=2, column=1, columnspan=2, sticky='w')
        ttk.Label(main, text='CC:').grid(row=3, column=0, pady=10, sticky='w')
        self.cc_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.cc_var, width=width_value).grid(row=3, column=1, columnspan=2, sticky='w')

        # Date picker
        ttk.Label(main, text='Select Cutoff Date:').grid(row=4, column=0, sticky='w')
        self.date_entry = DateEntry(main, width=12, background='darkblue',
                                    foreground='white', borderwidth=2,
                                    year=datetime.datetime.now().year)
        self.date_entry.grid(row=4, column=1, sticky='w')

        # Run
        self.run_btn = ttk.Button(main, text='Generate & Send', command=self._on_run)
        self.run_btn.grid(row=5, column=1, pady=20)

        # Progress & ETA
        self.progress = ttk.Progressbar(main, orient='horizontal', length=500, mode='determinate')
        self.progress.grid(row=6, column=0, columnspan=3, pady=(10, 0))
        self.time_var = tk.StringVar('Estimated time left: N/A')
        ttk.Label(main, textvariable=self.time_var).grid(row=7, column=0, columnspan=3, sticky='w')

        # Status
        self.status = ttk.Label(main, text='', foreground='green',
                                wraplength=600, justify='left')
        self.status.grid(row=8, column=0, columnspan=3, pady=(10, 0))

        # Start queue polling
        self.master.after(100, self._process_queue)

    def _browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx *.xls')])
        if not path:
            logger.debug("Excel browse canceled")
            return
        self.excel_path.set(path)
        logger.info("Selected Excel: %s", path)
        try:
            df = pd.read_excel(path, sheet_name='Access Folder')
            logger.debug("Loaded 'Access Folder' (%d rows)", len(df))
        except Exception as e:
            logger.exception("Failed reading 'Access Folder'")
            messagebox.showerror('Error', f'Could not read "Access Folder":\n{e}')
            return

        self.df_access = df
        owners = sorted(df['Entitlement Owner'].dropna().unique())
        logger.info("Owners found: %s", owners)
        self.person_cb['values'] = owners
        if owners:
            self.person_cb.current(0)
            self._person_selected()

    def _person_selected(self, event=None):
        if self.df_access is None:
            return
        owner = self.person_var.get()
        logger.info("Person selected: %s", owner)
        dfp = self.df_access[self.df_access['Entitlement Owner'] == owner]
        if dfp.empty:
            logger.warning("No entries for owner %s", owner)
            return
        row = dfp.iloc[0]
        to_email = row.get('Entitlement Owner Email', '') or ''
        cc_email = row.get('Delegate Email', '') or ''
        logger.debug("Setting To=%s, CC=%s", to_email, cc_email)
        self.email_var.set(to_email)
        self.cc_var.set(cc_email)

    def _on_run(self):
        logger.info("Run clicked")
        self.run_btn.config(state='disabled')
        self.status.config(text='')
        self.progress['value'] = 0
        self.time_var.set('Estimated time left: N/A')
        threading.Thread(target=self._process_files, daemon=True).start()

    def _process_files(self):
        logger.info("Background scan starting")
        owner = self.person_var.get().strip()
        to_email = self.email_var.get().strip()
        cc_email = self.cc_var.get().strip()
        cutoff_date = self.date_entry.get_date()
        cutoff_dt = datetime.datetime.combine(cutoff_date, datetime.time.max)
        cutoff_ts = cutoff_dt.timestamp()
        now_ts = time.time()

        dfp = self.df_access[self.df_access['Entitlement Owner'] == owner]
        base_paths = dfp['Full Path'].dropna().tolist()
        logger.debug("Base paths: %s", base_paths)

        # count files
        total = 0
        for base in base_paths:
            for _ in os.scandir(base or ''):
                total += 1  # scandir only counts top-level; deeper dirs handled below
            for root, dirs, _ in os.walk(base or ''):
                for d in dirs:
                    try:
                        total += len(os.listdir(os.path.join(root, d)))
                    except Exception:
                        pass
        logger.info("Total files (approx): %d", total)
        self.queue.put(('init', total))

        rows = []
        idx = 0
        start_time = time.time()

        def scan_dir(path):
            try:
                with os.scandir(path) as it:
                    for entry in it:
                        if entry.is_dir(follow_symlinks=False):
                            yield from scan_dir(entry.path)
                        elif entry.is_file(follow_symlinks=False):
                            yield entry
            except Exception as e:
                logger.error("Scan error in %s: %s", path, e)

        for base in base_paths:
            for entry in scan_dir(base):
                idx += 1
                try:
                    st = entry.stat(follow_symlinks=False)
                    cts = st.st_ctime
                    if cts > cutoff_ts:
                        continue
                    mts = st.st_mtime
                    ats = st.st_atime
                    days_ago = int((now_ts - cts) // 86400)
                    rows.append({
                        'Person':        owner,
                        'File Name':     entry.name,
                        'File Path':     entry.path,
                        'Created':       datetime.datetime.fromtimestamp(cts).strftime('%Y-%m-%d %H:%M:%S'),
                        'Last Modified': datetime.datetime.fromtimestamp(mts).strftime('%Y-%m-%d %H:%M:%S'),
                        'Last Accessed': datetime.datetime.fromtimestamp(ats).strftime('%Y-%m-%d %H:%M:%S'),
                        'Days Ago':      days_ago
                    })
                except Exception as e:
                    logger.error("Failed stat %s: %s", entry.path, e)
                    continue

                elapsed = time.time() - start_time
                avg = elapsed / idx
                rem = avg * (total - idx)
                self.queue.put(('progress', idx, rem))

        # write Excel
        now = datetime.datetime.now()
        stamp = now.strftime('%Y%m%d_%H%M%S')
        safe = owner.replace(' ', '_')
        report_fn = f"{safe}_report_{stamp}.xlsx"
        logger.info("Writing report to %s", report_fn)
        out_df = pd.DataFrame(rows)
        out_df.to_excel(report_fn, index=False, engine='openpyxl')

        wb = load_workbook(report_fn)
        ws = wb.active
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=0)
            ws.column_dimensions[col[0].column_letter].width = max_len + 2
        wb.save(report_fn)
        logger.info("Excel saved and columns auto-fitted")

        # send email
        status_msg = f"Report saved as {report_fn}"
        if win32:
            try:
                logger.info("Sending email To=%s CC=%s", to_email, cc_email)
                outlook = win32.Dispatch('Outlook.Application')
                m = outlook.CreateItem(0)
                m.To = to_email
                m.CC = cc_email
                m.Subject = f"File Report for {owner}"
                m.HTMLBody = (
                    f"<p>Dear {owner},</p>"
                    f"<p>Files created on or before {cutoff_date}:</p>"
                    f"{out_df.to_html(index=False)}"
                    "<p>Regards,</p>"
                )
                m.Attachments.Add(os.path.abspath(report_fn))
                m.Send()
                status_msg += f"\nand emailed to {to_email}"
                if cc_email:
                    status_msg += f" (CC: {cc_email})"
                logger.info("Email sent")
            except Exception as e:
                logger.exception("Email send failed")
                status_msg += f"\n(Email failed: {e})"
        else:
            logger.warning("pywin32 unavailable: skipping email")

        self.queue.put(('done', status_msg))
        logger.info("Background processing done")

    def _process_queue(self):
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
    logger.info("Starting application")
    root = tk.Tk()
    FileReportApp(root)
    root.mainloop()
    logger.info("Application exited")