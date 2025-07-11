#!/usr/bin/env python3
"""
pdf_diff_whole_doc.py

Whole-document text diff & highlight.  No CLI args.
"""

import fitz        # PyMuPDF
import difflib
import logging
from tqdm import tqdm

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
OLD_PATH       = "old.pdf"
NEW_PATH       = "new.pdf"
OLD_ANNOT      = "old_annotated.pdf"
NEW_ANNOT      = "new_annotated.pdf"
OUT_SIDEBYSIDE = "diff_output.pdf"

# highlight colors and opacity
COLOR_DEL    = (1, 0, 0)    # red
COLOR_INS    = (0, 1, 0)    # green
COLOR_REP    = (1, 1, 0)    # yellow
ALPHA        = 0.3

# ─── LOGGING ───────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


# ─── 1) EXTRACT ALL TEXT BLOCKS ─────────────────────────────────────────────────
def extract_all_blocks(path):
    """
    Open the PDF at `path` and return a list of all non-empty text-block dicts:
      { 'page': int, 'rect': fitz.Rect, 'text': str }
    """
    logger.info(f"Opening {path!r} and extracting text blocks…")
    doc = fitz.open(path)
    blocks = []
    for pno in tqdm(range(doc.page_count), desc="Extract blocks"):
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
    logger.info(f"Extracted {len(blocks)} blocks from {path!r}")
    return blocks


# ─── 2) COMPUTE DOCUMENT-LEVEL DIFF ─────────────────────────────────────────────
def compute_whole_doc_diff(old_blocks, new_blocks):
    """
    Given two lists of block-dicts, run SequenceMatcher on their 'text' fields.
    Returns four maps: deletions, replacements_old, insertions, replacements_new,
    each mapping page_no→[fitz.Rect,...].
    """
    logger.info("Running document-level diff…")
    old_texts = [b["text"] for b in old_blocks]
    new_texts = [b["text"] for b in new_blocks]
    sm = difflib.SequenceMatcher(None, old_texts, new_texts)
    
    del_map = {}
    ins_map = {}
    rep_o   = {}
    rep_n   = {}

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "delete":
            for idx in range(i1, i2):
                blk = old_blocks[idx]
                del_map.setdefault(blk["page"], []).append(blk["rect"])
        elif tag == "insert":
            for idx in range(j1, j2):
                blk = new_blocks[idx]
                ins_map.setdefault(blk["page"], []).append(blk["rect"])
        elif tag == "replace":
            for idx in range(i1, i2):
                blk = old_blocks[idx]
                rep_o.setdefault(blk["page"], []).append(blk["rect"])
            for idx in range(j1, j2):
                blk = new_blocks[idx]
                rep_n.setdefault(blk["page"], []).append(blk["rect"])
        # 'equal' we ignore entirely

    return del_map, rep_o, ins_map, rep_n


# ─── 3) ANNOTATE AND SAVE PDFs ──────────────────────────────────────────────────
def annotate_and_save(del_map, rep_o, ins_map, rep_n):
    # OLD PDF
    logger.info("Annotating OLD PDF (deletions & replacements)…")
    old_doc = fitz.open(OLD_PATH)
    for pno, page in enumerate(tqdm(old_doc, desc="Annotate OLD")):
        for r in del_map.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_DEL)
            a.set_opacity(ALPHA); a.update()
        for r in rep_o.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_REP)
            a.set_opacity(ALPHA); a.update()
    old_doc.save(OLD_ANNOT)
    logger.info(f"Saved {OLD_ANNOT!r}")

    # NEW PDF
    logger.info("Annotating NEW PDF (insertions & replacements)…")
    new_doc = fitz.open(NEW_PATH)
    for pno, page in enumerate(tqdm(new_doc, desc="Annotate NEW")):
        for r in ins_map.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_INS)
            a.set_opacity(ALPHA); a.update()
        for r in rep_n.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_REP)
            a.set_opacity(ALPHA); a.update()
    new_doc.save(NEW_ANNOT)
    logger.info(f"Saved {NEW_ANNOT!r}")

    return old_doc, new_doc


# ─── 4) RENDER SIDE-BY-SIDE ─────────────────────────────────────────────────────
def render_side_by_side(old_doc, new_doc):
    logger.info("Rendering side-by-side PDF…")
    out = fitz.open()
    pages = min(old_doc.page_count, new_doc.page_count)
    for i in tqdm(range(pages), desc="Render SxS"):
        p_old, p_new = old_doc[i], new_doc[i]
        w1, h1 = p_old.rect.width, p_old.rect.height
        w2, h2 = p_new.rect.width, p_new.rect.height
        H = max(h1, h2)

        newp = out.new_page(width=w1 + w2, height=H)
        pix1 = p_old.get_pixmap(alpha=False)
        pix2 = p_new.get_pixmap(alpha=False)
        newp.insert_image(fitz.Rect(0,    H - h1, w1,      H), pixmap=pix1)
        newp.insert_image(fitz.Rect(w1,   H - h2, w1 + w2, H), pixmap=pix2)

    out.save(OUT_SIDEBYSIDE)
    logger.info(f"Saved side-by-side PDF to {OUT_SIDEBYSIDE!r}")


# ─── MAIN WORKFLOW ─────────────────────────────────────────────────────────────
def main():
    old_blocks = extract_all_blocks(OLD_PATH)
    new_blocks = extract_all_blocks(NEW_PATH)

    del_map, rep_o, ins_map, rep_n = compute_whole_doc_diff(old_blocks, new_blocks)
    old_doc, new_doc = annotate_and_save(del_map, rep_o, ins_map, rep_n)
    render_side_by_side(old_doc, new_doc)
    logger.info("Done!")

if __name__ == "__main__":
    main()