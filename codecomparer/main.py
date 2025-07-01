#!/usr/bin/env python3
import os
import re
import difflib
import tempfile
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

def generate_diff_html(old_path, new_path):
    # Read both files
    with open(old_path, 'r', encoding='utf-8', errors='ignore') as f:
        old_lines = f.read().splitlines()
    with open(new_path, 'r', encoding='utf-8', errors='ignore') as f:
        new_lines = f.read().splitlines()

    # Generate base HTML diff
    diff = difflib.HtmlDiff(tabsize=4, wrapcolumn=0)
    html = diff.make_file(
        old_lines, new_lines,
        fromdesc=os.path.basename(old_path),
        todesc=os.path.basename(new_path),
        context=False, numlines=0
    )

    # Inject custom CSS: fixed layout + wrapping + color codes
    custom_css = """
    <style type="text/css">
      html, body { margin:0; padding:0; width:100%; overflow-x:hidden; }
      table.diff { width:100% !important; table-layout:fixed; border-collapse:collapse; }
      td, th { white-space: pre-wrap; word-wrap: break-word; }
      td.diff_sub { background-color: #ffcccc !important; }
      td.diff_add { background-color: #ccffcc !important; }
      td.diff_chg { background-color: #ffff99 !important; }
      th { background-color: #f0f0f0; padding:4px; }
      td { padding:2px 4px; vertical-align:top; font-family: monospace; }
    </style>
    </head>
    """
    html = html.replace("</head>", custom_css, 1)

    # Strip out <a name="n...">, <a name="t..."> and their hrefs
    html = re.sub(r'<a name="[nt]\d+">(\d+)</a>',       r'\1', html)
    html = re.sub(r'<a href="#[nt]\d+">(\d+)</a>',       r'\1', html)

    return html

def write_and_open(html):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode='w', encoding='utf-8')
    tmp.write(html)
    tmp.close()
    webbrowser.open(Path(tmp.name).absolute().as_uri())
    return tmp.name

def on_browse(entry_widget, title):
    path = filedialog.askopenfilename(title=title, filetypes=[("All files","*.*")])
    if path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, path)

def on_compare(old_entry, new_entry):
    old_path = old_entry.get().strip()
    new_path = new_entry.get().strip()
    if not old_path or not new_path:
        messagebox.showerror("Error", "Please select both Old and New code files.")
        return
    if not os.path.isfile(old_path) or not os.path.isfile(new_path):
        messagebox.showerror("Error", "One or both paths are not valid files.")
        return

    try:
        html = generate_diff_html(old_path, new_path)
        report_path = write_and_open(html)
        messagebox.showinfo("Done", f"Report generated and opened in your browser:\n{report_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while generating diff:\n{e}")

def main():
    root = tk.Tk()
    root.title("Code Comparison Tool")
    root.geometry("600x200")
    root.resizable(False, False)

    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TLabel', font=('Segoe UI', 10))
    style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=6)
    style.configure('TEntry', font=('Consolas', 10))

    frame = ttk.Frame(root, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frame, text="Old Code File:").grid(row=0, column=0, sticky=tk.W)
    old_entry = ttk.Entry(frame, width=50)
    old_entry.grid(row=0, column=1, padx=5)
    ttk.Button(frame, text="Browse…", command=lambda: on_browse(old_entry, "Select Old Code File"))\
        .grid(row=0, column=2)

    ttk.Label(frame, text="New Code File:").grid(row=1, column=0, sticky=tk.W, pady=(10,0))
    new_entry = ttk.Entry(frame, width=50)
    new_entry.grid(row=1, column=1, padx=5, pady=(10,0))
    ttk.Button(frame, text="Browse…", command=lambda: on_browse(new_entry, "Select New Code File"))\
        .grid(row=1, column=2, pady=(10,0))

    ttk.Button(frame, text="Compare", command=lambda: on_compare(old_entry, new_entry))\
        .grid(row=2, column=0, columnspan=3, pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()