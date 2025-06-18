import os
import datetime
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Requires: pandas, openpyxl, pywin32
try:
    import win32com.client as win32
except ImportError:
    win32 = None

from openpyxl import load_workbook

class FileReportApp:
    def __init__(self, master):
        self.master = master
        master.title('Access Folder File Report')
        master.geometry('520x300')
        master.resizable(False, False)

        # define a standard width for entry/combobox
        width_value = 40

        style = ttk.Style()
        style.theme_use('clam')

        main = ttk.Frame(master, padding=20)
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

        # --- Days threshold ---
        ttk.Label(main, text='Days Threshold:').grid(row=2, column=0, sticky='w')
        self.days_var = tk.IntVar(value=3)
        ttk.Spinbox(main, from_=0, to=3650, textvariable=self.days_var, width=5).grid(row=2, column=1, sticky='w')

        # --- Recipient email ---
        ttk.Label(main, text='Recipient Email:').grid(row=3, column=0, pady=10, sticky='w')
        self.email_var = tk.StringVar()
        ttk.Entry(main, textvariable=self.email_var, width=width_value).grid(row=3, column=1, columnspan=2, sticky='w')

        # --- Run button & status ---
        self.run_btn = ttk.Button(main, text='Generate & Send', command=self._generate_and_send)
        self.run_btn.grid(row=4, column=1, pady=20)
        # wraplength ensures long text wraps to next line
        self.status = ttk.Label(main, text='', foreground='green',
                                wraplength=480, justify='left')
        self.status.grid(row=5, column=0, columnspan=3)

        # Storage for email lookup
        self.email_map = {}

        for child in main.winfo_children():
            child.grid_configure(padx=5, pady=5)

    def _browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[('Excel files','*.xlsx *.xls')])
        if not path:
            return
        self.excel_path.set(path)

        try:
            # load headers for the first sheet
            df0 = pd.read_excel(path, nrows=0)
            headers = list(df0.columns)
            self.person_cb['values'] = headers
            if headers:
                self.person_cb.current(0)
        except Exception as e:
            messagebox.showerror('Error', f'Could not read Excel headers:\n{e}')
            return

        try:
            # load the emails sheet
            df_emails = pd.read_excel(path, sheet_name='emails')
            # expect columns 'TM' and 'Email'
            self.email_map = dict(zip(df_emails['TM'], df_emails['Email']))
            # populate the email field for the initial selection
            sel = self.person_var.get()
            if sel in self.email_map:
                self.email_var.set(self.email_map[sel])
        except Exception:
            # if no emails sheet or bad format, just skip
            self.email_map = {}

    def _person_selected(self, event):
        tm = self.person_var.get()
        email = self.email_map.get(tm, '')
        self.email_var.set(email)

    def _generate_and_send(self):
        excel_file = self.excel_path.get().strip()
        person = self.person_var.get().strip()
        days = self.days_var.get()
        email = self.email_var.get().strip()

        if not excel_file or not person or not email:
            messagebox.showerror('Missing Data', 'Please select an Excel file, a person, and enter an email address.')
            return

        # read folder paths for that person
        try:
            df_access = pd.read_excel(excel_file, header=0)
            paths = df_access[person].dropna().tolist()
        except Exception as e:
            messagebox.showerror('Error', f'Failed to load paths for "{person}":\n{e}')
            return

        now = datetime.datetime.now()
        cutoff = now - datetime.timedelta(days=days)
        rows = []

        for base in paths:
            if not os.path.isdir(base):
                continue
            for root, dirs, files in os.walk(base):
                for fname in files:
                    fp = os.path.join(root, fname)
                    cts = datetime.datetime.fromtimestamp(os.path.getctime(fp))
                    if cts <= cutoff:
                        mts = datetime.datetime.fromtimestamp(os.path.getmtime(fp))
                        age = (now - cts).days
                        rows.append({
                            'Person':        person,
                            'File Name':     fname,
                            'File Path':     fp,
                            'Created':       cts.strftime('%Y-%m-%d %H:%M:%S'),
                            'Last Modified': mts.strftime('%Y-%m-%d %H:%M:%S'),
                            'Days Ago':      age
                        })

        if not rows:
            messagebox.showinfo('No Matches', f'No files older than {days} days were found.')
            return

        out_df = pd.DataFrame(rows)
        stamp = now.strftime('%Y%m%d_%H%M%S')
        safe = person.replace(' ', '_')
        report_fn = f"{safe}_report_{stamp}.xlsx"

        # write to Excel then auto-adjust column widths
        out_df.to_excel(report_fn, index=False, engine='openpyxl')
        wb = load_workbook(report_fn)
        ws = wb.active
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                val = cell.value
                if val is None:
                    continue
                length = len(str(val))
                if length > max_length:
                    max_length = length
            ws.column_dimensions[col_letter].width = max_length + 2
        wb.save(report_fn)

        # send email via Outlook
        status_msg = f"Report saved as {report_fn}"
        if win32:
            try:
                outlook = win32.Dispatch('Outlook.Application')
                mail    = outlook.CreateItem(0)
                mail.To       = email
                mail.Subject  = f"File Report for {person}"
                html_tbl      = out_df.to_html(index=False)
                mail.HTMLBody = (
                    f"<p>Dear {person},</p>"
                    f"<p>Please find below the list of files older than {days} days:</p>"
                    f"{html_tbl}"
                    "<p>Best regards,</p>"
                )
                mail.Attachments.Add(os.path.abspath(report_fn))
                mail.Send()
                status_msg += f"\nand emailed to {email}"
            except Exception as e:
                messagebox.showerror('Email Error', f'Failed to send email:\n{e}')
        else:
            messagebox.showwarning('Outlook Unavailable',
                                   'pywin32 not installed – report saved but email not sent.')

        # update status label (will wrap if too long)
        self.status.config(text=status_msg)

if __name__ == '__main__':
    root = tk.Tk()
    FileReportApp(root)
    root.mainloop()