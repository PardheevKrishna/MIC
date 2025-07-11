#!/usr/bin/env python3
"""
pdf_diff_flatten.py

Document-level PDF text diff & highlight.  No CLI args—all paths/colors hard-coded.
"""

import fitz        # PyMuPDF
import difflib
import logging
from tqdm import tqdm

# ─── CONFIG ─────────────────────────────────────────────────────────────────────
OLD_PATH       = "old.pdf"
NEW_PATH       = "new.pdf"
OLD_ANNOT      = "old_annotated.pdf"
NEW_ANNOT      = "new_annotated.pdf"
OUT_SIDEBYSIDE = "diff_output.pdf"

# RGB tuples in [0,1]
COLOR_DELETE  = (1, 0, 0)    # red
COLOR_INSERT  = (0, 1, 0)    # green
COLOR_REPLACE = (1, 1, 0)    # yellow
ALPHA         = 0.3          # transparency for all

# ─── LOGGING SETUP ──────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)


# ─── 1) FLATTEN TEXT BLOCKS ─────────────────────────────────────────────────────
def extract_text_blocks(path):
    """
    Flatten all non-empty text blocks across the document.
    Returns a list of dicts: {
      'page': int,
      'rect': fitz.Rect,
      'text': str
    }
    """
    logger.info("Opening %r and extracting text blocks…", path)
    doc = fitz.open(path)
    blocks = []
    for page_no in tqdm(range(doc.page_count), desc="Extracting blocks"):
        page = doc[page_no]
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, text, *rest = b
            txt = text.strip()
            if not txt:
                continue
            blocks.append({
                "page": page_no,
                "rect": fitz.Rect(x0, y0, x1, y1),
                "text": txt
            })
    doc.close()
    logger.info("Extracted %d text blocks from %r", len(blocks), path)
    return blocks


# ─── 2) RUN A DOCUMENT-LEVEL DIFF ───────────────────────────────────────────────
def compute_diff(old_blocks, new_blocks):
    """
    Given two lists of block-dicts, compute diff ops on their .['text'].
    Returns four dicts: delete_map, replace_old_map, insert_map, replace_new_map,
    each mapping page_no → list of fitz.Rect to highlight.
    """
    old_texts = [b["text"] for b in old_blocks]
    new_texts = [b["text"] for b in new_blocks]

    logger.info("Running SequenceMatcher on %d old vs %d new blocks…",
                len(old_texts), len(new_texts))
    sm = difflib.SequenceMatcher(None, old_texts, new_texts)
    ops = sm.get_opcodes()

    delete_map      = {}
    replace_old_map = {}
    insert_map      = {}
    replace_new_map = {}

    for tag, i1, i2, j1, j2 in ops:
        if tag == "delete":
            for idx in range(i1, i2):
                blk = old_blocks[idx]
                delete_map.setdefault(blk["page"], []).append(blk["rect"])
        elif tag == "insert":
            for idx in range(j1, j2):
                blk = new_blocks[idx]
                insert_map.setdefault(blk["page"], []).append(blk["rect"])
        elif tag == "replace":
            for idx in range(i1, i2):
                blk = old_blocks[idx]
                replace_old_map.setdefault(blk["page"], []).append(blk["rect"])
            for idx in range(j1, j2):
                blk = new_blocks[idx]
                replace_new_map.setdefault(blk["page"], []).append(blk["rect"])
        # “equal” we ignore

    return delete_map, replace_old_map, insert_map, replace_new_map


# ─── 3) ANNOTATE PDFs ────────────────────────────────────────────────────────────
def annotate_docs(delete_map, replace_old_map,
                  insert_map, replace_new_map):
    """
    Opens both PDFs, applies real highlight annotations,
    then saves out old_annotated.pdf and new_annotated.pdf.
    """
    logger.info("Annotating OLD PDF with deletions & replacements…")
    old_doc = fitz.open(OLD_PATH)
    for pno, page in enumerate(tqdm(old_doc, desc="Annotate OLD")):
        # deletions (red)
        for r in delete_map.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_DELETE)
            a.set_opacity(ALPHA); a.update()
        # replacements (yellow)
        for r in replace_old_map.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_REPLACE)
            a.set_opacity(ALPHA); a.update()
    old_doc.save(OLD_ANNOT)
    logger.info("Saved %r", OLD_ANNOT)

    logger.info("Annotating NEW PDF with insertions & replacements…")
    new_doc = fitz.open(NEW_PATH)
    for pno, page in enumerate(tqdm(new_doc, desc="Annotate NEW")):
        # insertions (green)
        for r in insert_map.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_INSERT)
            a.set_opacity(ALPHA); a.update()
        # replacements (yellow)
        for r in replace_new_map.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_REPLACE)
            a.set_opacity(ALPHA); a.update()
    new_doc.save(NEW_ANNOT)
    logger.info("Saved %r", NEW_ANNOT)

    return old_doc, new_doc


# ─── 4) RENDER SIDE-BY-SIDE ──────────────────────────────────────────────────────
def render_side_by_side(old_doc, new_doc):
    logger.info("Rendering side-by-side PDF…")
    out = fitz.open()
    pages = min(old_doc.page_count, new_doc.page_count)

    for i in tqdm(range(pages), desc="Render SxS"):
        po = old_doc[i]; pn = new_doc[i]
        w1, h1 = po.rect.width, po.rect.height
        w2, h2 = pn.rect.width, pn.rect.height
        H = max(h1, h2)

        newp = out.new_page(width=w1 + w2, height=H)
        pix1 = po.get_pixmap(alpha=False)
        pix2 = pn.get_pixmap(alpha=False)

        # old on left, new on right
        newp.insert_image(fitz.Rect(0,    H-h1,   w1,      H), pixmap=pix1)
        newp.insert_image(fitz.Rect(w1,   H-h2,   w1+w2,   H), pixmap=pix2)

    out.save(OUT_SIDEBYSIDE)
    logger.info("Saved side-by-side PDF to %r", OUT_SIDEBYSIDE)


# ─── MAIN WORKFLOW ──────────────────────────────────────────────────────────────
def main():
    old_blocks = extract_text_blocks(OLD_PATH)
    new_blocks = extract_text_blocks(NEW_PATH)

    delete_map, rep_o_map, insert_map, rep_n_map = compute_diff(old_blocks, new_blocks)
    old_doc, new_doc = annotate_docs(delete_map, rep_o_map, insert_map, rep_n_map)
    render_side_by_side(old_doc, new_doc)
    logger.info("All done!")

if __name__ == "__main__":
    main()