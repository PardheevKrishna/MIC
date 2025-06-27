import os, glob, time, queue, threading, logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

# ─── LOGGING ──────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)

# ─── TIME FORMAT ──────────────────────────────────────────────────────────
def format_time(elapsed):
    hrs, rem = divmod(elapsed, 3600)
    mins, secs = divmod(rem, 60)
    return f"{int(hrs):02d}:{int(mins):02d}:{secs:05.2f}"

# ─── WORKER ───────────────────────────────────────────────────────────────
def process_files(folder, gui_queue):
    start_time   = time.time()
    env_map      = {'DEV':'1-DEV','UAT':'2-UAT','PROD':'3-PROD'}
    chunksize    = 200_000

    all_fp       = []
    all_policies = []
    skipped_raw  = []   # collect malformed policy lines

    for env_key in ('DEV','UAT','PROD'):
        pattern  = os.path.join(folder, f'*_{env_key}.csv')
        files    = glob.glob(pattern)
        if not files:
            logger.warning("Skipping %s (not found)", env_key)
            continue
        path_csv = files[0]
        env_tag  = env_map[env_key]
        logger.info("→ %s [%s]", path_csv, env_tag)

        # read raw lines once for policies parsing
        with open(path_csv, 'r', encoding='utf-8', errors='ignore') as f:
            raw_lines = f.readlines()

        # find where "Policies" appears (0-based)
        header_idx    = 12   # row 13 = index12
        sentinel_idx  = None
        for i, ln in enumerate(raw_lines):
            if ln.strip().startswith('Policies'):
                sentinel_idx = i
                break
        if sentinel_idx is None:
            sentinel_idx = len(raw_lines)

        # ── 1) Folders-Permissions ──────────────────────────────────
        reader = pd.read_csv(
            path_csv,
            skiprows=header_idx,
            header=0,
            dtype=str,
            chunksize=chunksize
        )
        rows_fp, hit_pol = 0, False

        for chunk in reader:
            chunk = chunk.fillna('')
            # stop at "Policies" in first column
            mpol = chunk.iloc[:,0].eq('Policies')
            if mpol.any():
                idx = mpol.idxmax()
                chunk = chunk.loc[:idx-1]
                hit_pol = True

            n = len(chunk)
            if n == 0 and hit_pol:
                break

            # vectorized transforms
            df0     = chunk.iloc[:,0].str.strip().rename('Default Name')
            loc     = chunk.iloc[:,1].str.strip()

            parts1  = loc.str.split(pat='/', n=1, expand=True)
            pathp   = parts1[1].fillna('').rename('Location / Path')

            parts2  = loc.str.split(pat='/', n=2, expand=True)
            top     = parts2[1].fillna('').rename('Top Level Folder')
            second  = parts2[2].fillna('').rename('Second Level Folder')

            # parse Q-column
            qcol         = chunk.iloc[:,16].str.strip().fillna('')
            sc           = qcol.str.split(pat=':', n=1, expand=True)
            left_part    = sc[0].str.strip()
            policies_str = sc[1].fillna('').str.strip()
            lr           = left_part.str.split(pat=' - ', n=1, expand=True)
            ppp          = lr[0].fillna('').str.strip().rename('Cognos Group/Role')
            qqq          = lr[1].fillna('').str.strip().rename('Security Group')

            dfc = pd.concat([
                df0,
                chunk.iloc[:,1].str.strip().rename('Location'),
                pathp, top, second,
                ppp, qqq,
                policies_str.rename(
                  'Policies[x=Execute, r=Read, p=Set Policies, w=Write, t=Traverse]'
                )
            ], axis=1)

            # policy flags
            for let in ('x','r','p','w','t'):
                dfc[let.upper()] = policies_str.\
                    str.contains(fr'(^|-){let}($|-)').astype(int)

            dfc['Environment'] = env_tag
            all_fp.append(dfc)

            rows_fp += n
            gui_queue.put(('progress_fp', rows_fp, time.time()-start_time))
            logger.info(" FP rows: %7d (+%d)", rows_fp, n)

            if hit_pol:
                break

        # ── 2) Policies section ───────────────────────────────────
        logger.info(" Parsing Policies in %s", path_csv)
        # header row = raw_lines[sentinel_idx], data starts at sentinel_idx+1
        for ln in raw_lines[sentinel_idx+1:]:
            if not ln.strip():
                # once we hit a blank, everything below is "skipped"
                continue
            cells = ln.rstrip('\n').split(',')
            if len(cells) < 2:
                skipped_raw.append({
                    'Environment': env_tag,
                    'Section': 'Policies',
                    'RawLine': ln.rstrip('\n'),
                    'Reason': 'fewer than 2 columns'
                })
                continue
            all_policies.append({
                'Owner Default Name': cells[0].strip(),
                'Object Default Name':cells[1].strip(),
                'Environment': env_tag
            })

        gui_queue.put(('progress_pol', len(all_policies), time.time()-start_time))
        logger.info(" Policies rows: %d", len(all_policies))

    # ── WRITE EXCEL ───────────────────────────────────────────────────
    logger.info(" Writing output.xlsx…")
    df_fp  = pd.concat(all_fp, ignore_index=True)
    df_pol = pd.DataFrame(all_policies)
    df_sk  = pd.DataFrame(skipped_raw)
    out_xl = os.path.join(folder, 'output.xlsx')

    with pd.ExcelWriter(out_xl, engine='xlsxwriter') as writer:
        # Folders-Permissions
        df_fp.to_excel(
          writer,
          sheet_name='Folders-Permissions',
          index=False
        )
        ws1 = writer.sheets['Folders-Permissions']

        # Policies (A1="Policies", row2=header, row3+ data)
        df_pol.to_excel(
          writer,
          sheet_name='Policies',
          startrow=2, index=False, header=False
        )
        ws2 = writer.sheets['Policies']
        ws2.write('A1', 'Policies')
        for c, col in enumerate(df_pol.columns):
            ws2.write(1, c, col)

        # Skipped-Raw
        if not df_sk.empty:
            df_sk.to_excel(
              writer,
              sheet_name='Skipped-Raw',
              index=False
            )
            ws3 = writer.sheets['Skipped-Raw']
            for idx, col in enumerate(df_sk.columns):
                width = max(df_sk[col].astype(str).map(len).max(), len(col)) + 2
                ws3.set_column(idx, idx, width)

        # autofit for sheet1 & sheet2
        for ws, df in ((ws1,df_fp),(ws2,df_pol)):
            for idx, col in enumerate(df.columns):
                w = max(df[col].astype(str).map(len).max(), len(col)) + 2
                ws.set_column(idx, idx, w)

    gui_queue.put(('done', out_xl))
    logger.info(" All done! %s", out_xl)


# ─── GUI ────────────────────────────────────────────────────────────────
class App:
    def __init__(self, root):
        root.title("CSV → Excel Fast (with debug)")
        root.geometry("540x180")
        self.q = queue.Queue()

        ttk.Button(root, text="Select Folder", command=self.choose).pack(pady=15)
        self.status = tk.StringVar(value="Waiting for folder selection…")
        ttk.Label(root, textvariable=self.status).pack(pady=5)
        self.root = root

    def choose(self):
        folder = filedialog.askdirectory(title="Select folder with CSVs")
        if not folder:
            return
        self.status.set("Starting…")
        threading.Thread(
          target=process_files,
          args=(folder, self.q),
          daemon=True
        ).start()
        self.root.after(100, self._poll)

    def _poll(self):
        try:
            msg = self.q.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll)
            return

        key = msg[0]
        if key == 'progress_fp':
            cnt, elapsed = msg[1], msg[2]
            self.status.set(f"FP rows: {cnt:,} | {format_time(elapsed)}")
        elif key == 'progress_pol':
            cnt, elapsed = msg[1], msg[2]
            self.status.set(f"Policies rows: {cnt:,} | {format_time(elapsed)}")
        elif key == 'done':
            out = msg[1]
            self.status.set(f"Done! → {out}")
            messagebox.showinfo("Completed", f"Excel ready:\n{out}")
        else:
            logger.warning("Unknown message: %r", msg)

        if key != 'done':
            self.root.after(100, self._poll)

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()