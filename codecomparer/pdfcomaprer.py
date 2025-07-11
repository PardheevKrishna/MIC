#!/usr/bin/env python3
"""
pdf_diff_global.py

Whole-document, order-independent text diff.  
Highlights only true deletions (red) and insertions (green).
"""

import fitz      # PyMuPDF
import logging
from collections import defaultdict, Counter
from tqdm import tqdm

# ─── CONFIG ─────────────────────────────────────────────────────────────────────
OLD_PATH       = "old.pdf"
NEW_PATH       = "new.pdf"
OLD_ANNOT      = "old_annotated.pdf"
NEW_ANNOT      = "new_annotated.pdf"
OUT_SIDEBYSIDE = "diff_output.pdf"

# highlight colors and opacity
COLOR_DEL = (1, 0, 0)    # red
COLOR_INS = (0, 1, 0)    # green
ALPHA     = 0.3

# ─── LOGGING ───────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


# ─── 1) EXTRACT ALL TEXT BLOCKS ─────────────────────────────────────────────────
def extract_blocks(path):
    """
    Returns a list of block-dicts for every non-empty text block:
      { 'page':int, 'rect':fitz.Rect, 'text':str }
    """
    logger.info("Extracting text blocks from %r…", path)
    doc = fitz.open(path)
    blocks = []
    for pno in tqdm(range(doc.page_count), desc="Extracting"):
        page = doc[pno]
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, text, *rest = b
            txt = text.strip()
            if not txt:
                continue
            blocks.append({
                "page": pno,
                "rect": fitz.Rect(x0, y0, x1, y1),
                "text": txt
            })
    doc.close()
    logger.info("→ %d blocks extracted", len(blocks))
    return blocks


# ─── 2) MATCH & DIFF GLOBALLY ───────────────────────────────────────────────────
def global_diff(old_blocks, new_blocks):
    """
    Match identical block-texts in a document-wide, order-independent way.
    Returns two dicts: delete_map and insert_map mapping page_no→[rect,...]
    """
    # Count occurrences of each text in both docs
    old_count = Counter(b["text"] for b in old_blocks)
    new_count = Counter(b["text"] for b in new_blocks)

    # For each text, the number of matched = min(old_count, new_count)
    match_count = {t: min(old_count[t], new_count[t]) for t in old_count}

    # Collect blocks by text
    old_by_text = defaultdict(list)
    for b in old_blocks:
        old_by_text[b["text"]].append(b)
    new_by_text = defaultdict(list)
    for b in new_blocks:
        new_by_text[b["text"]].append(b)

    delete_map = defaultdict(list)
    insert_map = defaultdict(list)

    # Handle each unique text
    for text, o_list in old_by_text.items():
        n_list = new_by_text.get(text, [])
        m = match_count.get(text, 0)
        # first m occurrences are "matched" → skip
        # any beyond m in old → deletions
        for b in o_list[m:]:
            delete_map[b["page"]].append(b["rect"])
    for text, n_list in new_by_text.items():
        o_list = old_by_text.get(text, [])
        m = match_count.get(text, 0)
        # any beyond m in new → insertions
        for b in n_list[m:]:
            insert_map[b["page"]].append(b["rect"])

    return delete_map, insert_map


# ─── 3) ANNOTATE PDFs ────────────────────────────────────────────────────────────
def annotate_and_save(delete_map, insert_map):
    # OLD: highlight deletions in RED
    logger.info("Annotating OLD PDF…")
    old_doc = fitz.open(OLD_PATH)
    for pno, page in enumerate(tqdm(old_doc, desc="OLD")):
        for r in delete_map.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_DEL)
            a.set_opacity(ALPHA)
            a.update()
    old_doc.save(OLD_ANNOT)
    logger.info("→ saved %r", OLD_ANNOT)

    # NEW: highlight insertions in GREEN
    logger.info("Annotating NEW PDF…")
    new_doc = fitz.open(NEW_PATH)
    for pno, page in enumerate(tqdm(new_doc, desc="NEW")):
        for r in insert_map.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_INS)
            a.set_opacity(ALPHA)
            a.update()
    new_doc.save(NEW_ANNOT)
    logger.info("→ saved %r", NEW_ANNOT)

    return fitz.open(OLD_ANNOT), fitz.open(NEW_ANNOT)


# ─── 4) SIDE-BY-SIDE RENDER ─────────────────────────────────────────────────────
def render_side_by_side(old_doc, new_doc):
    logger.info("Rendering side-by-side…")
    combo = fitz.open()
    pages = min(old_doc.page_count, new_doc.page_count)
    for i in tqdm(range(pages), desc="Render"):
        po, pn = old_doc[i], new_doc[i]
        w1, h1 = po.rect.width, po.rect.height
        w2, h2 = pn.rect.width, pn.rect.height
        H = max(h1, h2)

        newp = combo.new_page(width=w1 + w2, height=H)
        pix1 = po.get_pixmap(alpha=False)
        pix2 = pn.get_pixmap(alpha=False)
        newp.insert_image(fitz.Rect(0,    H-h1, w1,      H), pixmap=pix1)
        newp.insert_image(fitz.Rect(w1,   H-h2, w1+w2,   H), pixmap=pix2)
    combo.save(OUT_SIDEBYSIDE)
    logger.info("→ saved %r", OUT_SIDEBYSIDE)


def main():
    old_blocks = extract_blocks(OLD_PATH)
    new_blocks = extract_blocks(NEW_PATH)
    delete_map, insert_map = global_diff(old_blocks, new_blocks)
    old_doc, new_doc = annotate_and_save(delete_map, insert_map)
    render_side_by_side(old_doc, new_doc)
    logger.info("Done.")

if __name__ == "__main__":
    main()