#!/usr/bin/env python3
"""
pdf_diff_global.py

0) Hard-codes old/new PDF paths and all parameters.
1) Uses pdf-diff’s compute_changes() to get a list of change‐boxes.
2) Writes those boxes out as diff.json.
3) Uses pdf-diff’s render_changes() to draw red outlines on a PNG.
"""

import json
from pdf_diff import command_line  # pdf-diff’s Python API  [oai_citation:0‡Stack Overflow](https://stackoverflow.com/questions/64951531/using-a-python-module-in-spyder-rather-than-command-line?utm_source=chatgpt.com)

# ─── Configuration ─────────────────────────────────────────────────────────────
OLD_PDF     = "old.pdf"
NEW_PDF     = "new.pdf"
JSON_OUT    = "diff.json"
PNG_OUT     = "comparison_output.png"

# pdf-diff parameters (hard-coded)
TOP_MARGIN    = 0      # percent
BOTTOM_MARGIN = 100    # percent
STYLE         = ["underline", "strike"]  # default styles
WIDTH         = 900    # output image width

# ─── 1) Compute the raw changes ─────────────────────────────────────────────────
# returns a mixed list of dicts (boxes) and "*" markers  [oai_citation:1‡GitHub](https://raw.githubusercontent.com/JoshData/pdf-diff/primary/pdf_diff/command_line.py)
changes = command_line.compute_changes(
    OLD_PDF, NEW_PDF,
    top_margin=TOP_MARGIN,
    bottom_margin=BOTTOM_MARGIN
)

# ─── 2) Write out JSON of just the box‐dicts ────────────────────────────────────
# filter out the "*" markers and serialize
boxes = [c for c in changes if isinstance(c, dict)]
with open(JSON_OUT, "w") as f:
    json.dump(boxes, f, indent=2)
print(f"Wrote {len(boxes)} change‐boxes to {JSON_OUT!r}")

# ─── 3) Render a side-by-side PNG with red outlines ─────────────────────────────
# note: render_changes() returns a PIL Image  [oai_citation:2‡GitHub](https://raw.githubusercontent.com/JoshData/pdf-diff/primary/pdf_diff/command_line.py)
img = command_line.render_changes(changes, STYLE, width=WIDTH)
img.save(PNG_OUT, "PNG")
print(f"Saved visual diff to {PNG_OUT!r}")