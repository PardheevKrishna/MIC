import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import queue
import time
import os
import glob
import pandas as pd

def format_time(seconds):
    mins, secs = divmod(int(seconds), 60)
    hours, mins = divmod(mins, 60)
    return f"{hours:02d}:{mins:02d}:{secs:02d}"

def process_files(folder, queue):
    start_time = time.time()
    combined_fp = []
    combined_pol = []
    env_map = {'DEV': '1-DEV', 'UAT': '2-UAT', 'PROD': '3-PROD'}

    for env_key in ['DEV', 'UAT', 'PROD']:
        # find matching CSV
        pattern = os.path.join(folder, f'*_{env_key}.csv')
        files = glob.glob(pattern)
        if not files:
            continue
        csv_file = files[0]
        env = env_map[env_key]

        # --- Folders-Permissions section ---
        df = pd.read_csv(csv_file, skiprows=13, header=0, dtype=str).fillna('')
        for _, row in df.iterrows():
            a = str(row.iat[0]).strip()
            if a == 'Policies':
                break  # stop when we hit the policies section
            b = str(row.iat[1]).strip()
            q = str(row.iat[16]).strip()  # Column Q

            # derive path segments
            segments = b.split('/')
            loc_path = '/'.join(segments[1:]) if len(segments) > 1 else ''
            top = segments[1] if len(segments) > 1 else ''
            second = '/'.join(segments[2:]) if len(segments) > 2 else ''

            # parse Cognos Group/Role and policies
            ppp = qqq = policies_str = ''
            if ' - ' in q:
                left, rest = q.split(' - ', 1)
                ppp = left.strip()
                if ':' in rest:
                    qqq, policies_str = rest.split(':', 1)
                    qqq = qqq.strip()
                    policies_str = policies_str.strip()
                else:
                    qqq = rest.strip()
            else:
                if ':' in q:
                    policies_str = q.split(':', 1)[1].strip()

            parts = policies_str.split('-') if policies_str else []
            flags = {letter.upper(): (1 if letter in parts else 0) for letter in ['x','r','p','w','t']}

            combined_fp.append({
                'Default Name': a,
                'Location': b,
                'Location / Path': loc_path,
                'Top Level Folder': top,
                'Second Level Folder': second,
                'Cognos Group/Role': ppp,
                'Security Group': qqq,
                'Policies[x=Execute, r=Read, p=Set Policies, w=Write, t=Traverse]': policies_str,
                'X': flags['X'], 'R': flags['R'], 'P': flags['P'],
                'W': flags['W'], 'T': flags['T'],
                'Environment': env
            })

            # update progress
            processed = len(combined_fp)
            elapsed = time.time() - start_time
            queue.put(('progress_fp', processed, elapsed))

        # --- Policies section (raw file parse) ---
        with open(csv_file, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()

        header_idx = None
        for i, ln in enumerate(lines):
            if ln.strip().startswith('Policies'):
                header_idx = i + 1
                break

        if header_idx is not None and header_idx < len(lines):
            for ln in lines[header_idx+1:]:
                if not ln.strip():
                    break
                cells = ln.strip().split(',')
                if len(cells) < 2:
                    continue
                combined_pol.append({
                    'Owner Default Name': cells[0].strip(),
                    'Object Default Name': cells[1].strip(),
                    'Environment': env
                })
                queue.put(('progress_pol', len(combined_pol), time.time() - start_time))

    # Build DataFrames
    df_fp = pd.DataFrame(combined_fp)
    df_pol = pd.DataFrame(combined_pol)

    # Write to Excel with formatting & autofit
    out_file = os.path.join(folder, 'output.xlsx')
    with pd.ExcelWriter(out_file, engine='xlsxwriter') as writer:
        # Folders-Permissions sheet
        df_fp.to_excel(writer, sheet_name='Folders-Permissions', index=False)
        fp_ws = writer.sheets['Folders-Permissions']

        # Policies sheet (with "Policies" in A1)
        df_pol.to_excel(writer, sheet_name='Policies', startrow=1, index=False, header=False)
        pol_ws = writer.sheets['Policies']
        pol_ws.write('A1', 'Policies')
        for col_num, value in enumerate(df_pol.columns):
            pol_ws.write(1, col_num, value)

        # Autofit columns on both sheets
        for ws, df in [(fp_ws, df_fp), (pol_ws, df_pol)]:
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                ws.set_column(idx, idx, max_len)

    queue.put(('done', out_file))

class App:
    def __init__(self, root):
        self.root = root
        root.title("CSV → Excel Processor")
        root.geometry("500x150")
        self.queue = queue.Queue()

        self.btn_select = ttk.Button(root, text="Select Folder", command=self.select_folder)
        self.btn_select.pack(pady=20)

        self.status_var = tk.StringVar("Waiting for folder selection…")
        ttk.Label(root, textvariable=self.status_var).pack()

    def select_folder(self):
        folder = filedialog.askdirectory(title="Select folder with your CSVs")
        if not folder:
            return
        self.btn_select.config(state=tk.DISABLED)
        self.status_var.set("Initializing…")
        threading.Thread(target=process_files, args=(folder, self.queue), daemon=True).start()
        self.root.after(100, self.update_gui)

    def update_gui(self):
        try:
            msg = self.queue.get_nowait()
            key, count, elapsed = msg[0], msg[1], msg[2]
            if key == 'progress_fp':
                self.status_var.set(f"Folders-Permissions rows: {count} | Elapsed: {format_time(elapsed)}")
            elif key == 'progress_pol':
                self.status_var.set(f"Policies rows: {count} | Elapsed: {format_time(elapsed)}")
            elif key == 'done':
                out_file = msg[1]
                self.status_var.set(f"Done! Saved to {out_file}")
                messagebox.showinfo("Completed", f"Excel generated at:\n{out_file}")
                self.btn_select.config(state=tk.NORMAL)
                return
        except queue.Empty:
            pass
        self.root.after(100, self.update_gui)

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()