#!/usr/bin/env python3
import os
import sys
import difflib
import tempfile
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

def generate_diff_html(old_path, new_path):
    # Read files
    with open(old_path, 'r', encoding='utf-8', errors='ignore') as f:
        old_lines = f.read().splitlines()
    with open(new_path, 'r', encoding='utf-8', errors='ignore') as f:
        new_lines = f.read().splitlines()

    # Create the HtmlDiff and generate the full HTML
    diff = difflib.HtmlDiff(tabsize=4, wrapcolumn=0)
    html = diff.make_file(
        old_lines,
        new_lines,
        fromdesc=os.path.basename(old_path),
        todesc=os.path.basename(new_path),
        context=False,
        numlines=0
    )

    # Inject custom CSS to override colors and stretch to full width
    custom_css = """
    <style type="text/css">
      /* Make the diff table fill the window */
      table.diff { width: 100% !important; border-collapse: collapse; }
      /* Deleted lines: bright red on the old side */
      td.diff_sub { background-color: #ffcccc !important; }
      /* Inserted lines: bright green on the new side */
      td.diff_add { background-color: #ccffcc !important; }
      /* Changed lines: bright yellow on both sides */
      td.diff_chg { background-color: #ffff99 !important; }
      /* Optional: tweak header */
      th { background-color: #f0f0f0; padding: 4px; }
      td { padding: 2px 4px; vertical-align: top; font-family: monospace; }
    </style>
    </head>
    """
    # Replace only the first </head> with our CSS+</head>
    html = html.replace("</head>", custom_css, 1)
    return html

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
        # Write to temp file
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode='w', encoding='utf-8')
        tmp.write(html)
        tmp.close()
        # Open in browser
        uri = Path(tmp.name).absolute().as_uri()
        webbrowser.open(uri)
        messagebox.showinfo("Done", f"Report generated and opened in your browser:\n\n{tmp.name}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while generating diff:\n{e}")

def main():
    root = tk.Tk()
    root.title("Code Comparison Tool")
    root.geometry("600x200")
    root.resizable(False, False)

    # Use a nicer theme
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TLabel', font=('Segoe UI', 10))
    style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=6)
    style.configure('TEntry', font=('Consolas', 10))

    # Layout
    frame = ttk.Frame(root, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)

    # Old code
    ttk.Label(frame, text="Old Code File:").grid(row=0, column=0, sticky=tk.W)
    old_entry = ttk.Entry(frame, width= fifty  50)
    old_entry.grid(row=0, column=1, padx=5)
    ttk.Button(frame, text="Browse…", command=lambda: on_browse(old_entry, "Select Old Code File")).grid(row=0, column=2)

    # New code
    ttk.Label(frame, text="New Code File:").grid(row=1, column=0, sticky=tk.W, pady=(10,0))
    new_entry = ttk.Entry(frame, width=50)
    new_entry.grid(row=1, column=1, padx=5, pady=(10,0))
    ttk.Button(frame, text="Browse…", command=lambda: on_browse(new_entry, "Select New Code File")).grid(row=1, column=2, pady=(10,0))

    # Compare button
    compare_btn = ttk.Button(frame, text="Compare", command=lambda: on_compare(old_entry, new_entry))
    compare_btn.grid(row=2, column=0, columnspan=3, pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()