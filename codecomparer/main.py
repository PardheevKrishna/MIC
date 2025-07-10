#!/usr/bin/env python3
import os
import re
import tempfile
import webbrowser
from pathlib import Path

import pandas as pd
import difflib
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- Helpers for code diff (unchanged) ---
def _code_diff_html(old_lines, new_lines, old_name, new_name):
    diff = difflib.HtmlDiff(tabsize=4, wrapcolumn=0)
    html = diff.make_file(
        old_lines, new_lines,
        fromdesc=old_name, todesc=new_name,
        context=False, numlines=0
    )
    # inject our CSS and strip out n…/t… anchors
    custom_css = """
    <style>
      html,body { margin:0; padding:0; overflow-x:hidden; }
      table.diff { width:100%!important; table-layout:fixed; border-collapse:collapse; }
      td,th { white-space:pre-wrap; word-wrap:break-word; font-family:monospace; padding:2px 4px; }
      td.diff_sub { background:#ffcccc!important; }
      td.diff_add { background:#ccffcc!important; }
      td.diff_chg { background:#ffff99!important; }
      th { background:#f0f0f0; padding:4px; }
    </style>
    </head>
    """
    html = html.replace("</head>", custom_css, 1)
    html = re.sub(r'<a name="[nt]\d+">(\d+)</a>', r'\1', html)
    html = re.sub(r'<a href="#[nt]\d+">(\d+)</a>', r'\1', html)
    return html

# --- Helpers for table diff ---
def _table_diff_html(old_df: pd.DataFrame, new_df: pd.DataFrame, old_name, new_name):
    # unify shape
    idx = old_df.index.union(new_df.index)
    cols = old_df.columns.union(new_df.columns)
    old = old_df.reindex(index=idx, columns=cols, fill_value="")
    new = new_df.reindex(index=idx, columns=cols, fill_value="")

    # build HTML tables with per-cell classes
    def df_to_html(df, other_df, side):
        # side = 'old' or 'new'
        rows = []
        # header
        header = "".join(f"<th>{col}</th>" for col in df.columns)
        rows.append(f"<tr><th></th>{header}</tr>")
        for i in df.index:
            cells = []
            for col in df.columns:
                a = df.at[i,col]
                b = other_df.at[i,col]
                if side=="old":
                    cls = "deleted" if (a!="" and b=="") else ("changed" if (a!="" and b!="" and a!=b) else "")
                else:
                    cls = "added"   if (b!="" and a=="") else ("changed" if (a!="" and b!="" and a!=b) else "")
                cell = f'<td class="{cls}">{a if side=="old" else b}</td>'
                cells.append(cell)
            rows.append(f"<tr><th>{i}</th>" + "".join(cells) + "</tr>")
        return "<table class='{}'><caption>{}</caption>{}</table>".format(
            f"table-{side}", old_name if side=="old" else new_name, "\n".join(rows)
        )

    old_table = df_to_html(old, new, "old")
    new_table = df_to_html(new, old, "new")

    css = """
    <style>
      html,body { margin:0; padding:0; overflow-x:hidden; }
      .container { display:flex; justify-content:space-between; }
      table { border-collapse:collapse; table-layout:fixed; width:48%; font-family:monospace; }
      th, td { border:1px solid #ccc; padding:4px; word-wrap:break-word; white-space:pre-wrap; }
      .added   { background:#ccffcc; }
      .deleted { background:#ffcccc; }
      .changed { background:#ffff99; }
      caption { font-weight:bold; text-align:left; padding:4px; }
      th { background:#f0f0f0; }
    </style>
    </head>
    """
    body = f"<div class='container'>\n{old_table}\n{new_table}\n</div>"
    return "<html><head>" + css + "</html><body>" + body + "</body>"

# --- Main dispatcher ---
def generate_diff_html(old_path, new_path):
    ext = (Path(old_path).suffix.lower(), Path(new_path).suffix.lower())
    tabular_exts = {'.csv', '.xls', '.xlsx', '.xlsm', '.ods'}
    if ext[0] in tabular_exts and ext[1] in tabular_exts:
        # read as DataFrame
        old_df = pd.read_excel(old_path) if ext[0] != '.csv' else pd.read_csv(old_path)
        new_df = pd.read_excel(new_path) if ext[1] != '.csv' else pd.read_csv(new_path)
        return _table_diff_html(old_df, new_df,
                                os.path.basename(old_path), os.path.basename(new_path))
    else:
        # fallback to code diff
        with open(old_path, 'r', encoding='utf-8', errors='ignore') as f:
            old_lines = f.read().splitlines()
        with open(new_path, 'r', encoding='utf-8', errors='ignore') as f:
            new_lines = f.read().splitlines()
        return _code_diff_html(old_lines, new_lines,
                               os.path.basename(old_path), os.path.basename(new_path))

# --- file & UI logic (unchanged) ---
def write_and_open(html):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode='w', encoding='utf-8')
    tmp.write(html)
    tmp.close()
    webbrowser.open(Path(tmp.name).absolute().as_uri())
    return tmp.name

def on_browse(entry, title):
    path = filedialog.askopenfilename(title=title, filetypes=[("All files","*.*")])
    if path:
        entry.delete(0, tk.END); entry.insert(0, path)

def on_compare(old_e, new_e):
    o, n = old_e.get().strip(), new_e.get().strip()
    if not o or not n or not os.path.isfile(o) or not os.path.isfile(n):
        messagebox.showerror("Error", "Select two valid files.")
        return
    try:
        html = generate_diff_html(o, n)
        out = write_and_open(html)
        messagebox.showinfo("Done", f"Report opened:\n{out}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def main():
    root = tk.Tk(); root.title("Universal Comparator"); root.geometry("600x200")
    style = ttk.Style(root); style.theme_use('clam')
    for w in ('TLabel','TButton','TEntry'):
        style.configure(w, font=('Segoe UI', 10))
    frame = ttk.Frame(root, padding=20); frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frame, text="Old File:").grid(row=0,column=0, sticky=tk.W)
    old_ent = ttk.Entry(frame, width=50); old_ent.grid(row=0,column=1,padx=5)
    ttk.Button(frame, text="Browse…", command=lambda: on_browse(old_ent,"Old")).grid(row=0,column=2)

    ttk.Label(frame, text="New File:").grid(row=1,column=0, sticky=tk.W, pady=(10,0))
    new_ent = ttk.Entry(frame, width=50); new_ent.grid(row=1,column=1,padx=5,pady=(10,0))
    ttk.Button(frame, text="Browse…", command=lambda: on_browse(new_ent,"New")).grid(row=1,column=2,pady=(10,0))

    ttk.Button(frame, text="Compare", command=lambda: on_compare(old_ent,new_ent))\
        .grid(row=2,column=0,columnspan=3,pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()