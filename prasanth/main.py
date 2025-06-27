import os, glob, time, queue, threading, logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

# ─── CONFIGURE LOGGING ─────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)

# ─── TIME FORMATTING ───────────────────────────────────────────────────────────
def format_time(elapsed):
    hrs, rem = divmod(elapsed, 3600)
    mins, secs = divmod(rem, 60)
    return f"{int(hrs):02d}:{int(mins):02d}:{secs:05.2f}"

# ─── WORKER ───────────────────────────────────────────────────────────────────
def process_files(folder, gui_queue):
    start_time = time.time()
    env_map     = {'DEV': '1-DEV', 'UAT': '2-UAT', 'PROD': '3-PROD'}
    chunksize   = 200_000

    all_fp_chunks = []
    all_pol_rows  = []

    for env_key in ('DEV','UAT','PROD'):
        pattern = os.path.join(folder, f'*_{env_key}.csv')
        files   = glob.glob(pattern)
        if not files:
            logger.warning("No %s file found, skipping", env_key)
            continue
        csv_file = files[0]
        env_tag  = env_map[env_key]
        logger.info("=== Processing %s → %s ===", csv_file, env_tag)

        # ── PART 1: Folders-Permissions (chunked) ──────────────────────
        reader = pd.read_csv(
            csv_file,
            skiprows=13,
            header=0,
            dtype=str,
            chunksize=chunksize
        )
        processed_rows = 0
        hit_policies  = False

        for chunk in reader:
            chunk = chunk.fillna('')
            # stop at the 'Policies' sentinel in col A
            mask_p = chunk.iloc[:,0].eq('Policies')
            if mask_p.any():
                idx = mask_p.idxmax()
                chunk = chunk.loc[:idx-1]
                hit_policies = True

            n_chunk = len(chunk)
            if n_chunk == 0 and hit_policies:
                break

            # vectorized derivations
            default = chunk.iloc[:,0].str.strip()
            loc     = chunk.iloc[:,1].str.strip()

            parts1 = loc.str.split('/', n=1, expand=True)
            path   = parts1[1].fillna('')   # drop xxx/
            parts2 = loc.str.split('/', n=2, expand=True)
            top    = parts2[1].fillna('')
            second = parts2[2].fillna('')

            # ── FIXED PARSING FOR COLUMN Q ─────────────────────────────
            qcol         = chunk.iloc[:,16].str.strip().fillna('')
            # 1) split on colon
            split_colon  = qcol.str.split(':', n=1, expand=True)
            left_part    = split_colon[0].str.strip()
            pols         = split_colon[1].fillna('').str.strip()
            # 2) split left_part on ' - ' only when present
            lr           = left_part.str.split(' - ', n=1, expand=True)
            ppp          = lr[0].fillna('').str.strip()  # Cognos Group/Role
            qqq          = lr[1].fillna('').str.strip()  # Security Group or ''

            dfc = pd.DataFrame({
                'Default Name': default,
                'Location': loc,
                'Location / Path': path,
                'Top Level Folder': top,
                'Second Level Folder': second,
                'Cognos Group/Role': ppp,
                'Security Group': qqq,
                'Policies[x=Execute, r=Read, p=Set Policies, w=Write, t=Traverse]': pols,
            })

            # flags for each policy letter
            for letter in ('x','r','p','w','t'):
                pattern = fr'(^|-){letter}($|-)'
                dfc[letter.upper()] = pols.str.contains(pattern).astype(int)

            dfc['Environment'] = env_tag

            all_fp_chunks.append(dfc)
            processed_rows += n_chunk

            elapsed = time.time() - start_time
            gui_queue.put(('progress_fp', processed_rows, elapsed))
            logger.info("Processed %7d FP rows (chunk %d)", processed_rows, n_chunk)

            if hit_policies:
                break

        logger.info("Finished FP for %s: %d rows", env_key, processed_rows)

        # ── PART 2: Policies section ────────────────────────────────────
        logger.info("Parsing Policies section in %s", csv_file)
        with open(csv_file, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
        hdr_idx = None
        for i, ln in enumerate(lines):
            if ln.strip().startswith('Policies'):
                hdr_idx = i + 1
                break
        if hdr_idx is not None:
            for ln in lines[hdr_idx+1:]:
                if not ln.strip():
                    break
                cells = ln.rstrip('\n').split(',')
                if len(cells) < 2:
                    continue
                all_pol_rows.append({
                    'Owner Default Name': cells[0].strip(),
                    'Object Default Name': cells[1].strip(),
                    'Environment': env_tag
                })
            gui_queue.put(('progress_pol', len(all_pol_rows), time.time() - start_time))
            logger.info("Finished Policies for %s: %d rows", env_key, len(all_pol_rows))
        else:
            logger.warning("No Policies section found in %s", csv_file)

    # ── WRITE EXCEL ────────────────────────────────────────────────────────
    logger.info("Writing output.xlsx…")
    df_fp  = pd.concat(all_fp_chunks, ignore_index=True)
    df_pol = pd.DataFrame(all_pol_rows)
    out_path = os.path.join(folder, 'output.xlsx')
    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        df_fp.to_excel(writer, sheet_name='Folders-Permissions', index=False)
        ws_fp = writer.sheets['Folders-Permissions']

        df_pol.to_excel(writer, sheet_name='Policies',
                        startrow=1, index=False, header=False)
        ws_pol = writer.sheets['Policies']
        ws_pol.write('A1', 'Policies')
        for c, col in enumerate(df_pol.columns):
            ws_pol.write(1, c, col)

        # autofit columns
        for ws, df in ((ws_fp, df_fp), (ws_pol, df_pol)):
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                ws.set_column(idx, idx, max_len)

    gui_queue.put(('done', out_path))
    logger.info("All done! Excel saved to %s", out_path)

# ─── GUI APP ─────────────────────────────────────────────────────────────────
class App:
    def __init__(self, root):
        root.title("CSV → Excel Fast Processor")
        root.geometry("520x160")
        self.q = queue.Queue()

        self.btn = ttk.Button(root, text="Select Folder", command=self.choose)
        self.btn.pack(pady=15)

        self.status = tk.StringVar(value="Waiting for folder selection…")
        ttk.Label(root, textvariable=self.status).pack(pady=5)

        self.root = root

    def choose(self):
        folder = filedialog.askdirectory(title="Select folder with CSVs")
        if not folder:
            return
        self.btn.config(state=tk.DISABLED)
        self.status.set("Starting…")
        threading.Thread(target=process_files,
                         args=(folder, self.q),
                         daemon=True).start()
        self.root.after(100, self._poll)

    def _poll(self):
        try:
            key, a, b = self.q.get_nowait()
            if key == 'progress_fp':
                self.status.set(f"FP rows: {a:,} | {format_time(b)}")
            elif key == 'progress_pol':
                self.status.set(f"Policies rows: {a:,} | {format_time(b)}")
            elif key == 'done':
                self.status.set(f"Done! → {a}")
                messagebox.showinfo("Completed", f"Excel ready:\n{a}")
                self.btn.config(state=tk.NORMAL)
                return
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()