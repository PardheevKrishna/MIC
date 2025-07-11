#!/usr/bin/env python3
"""
pdf_diff_text_only.py

Compare two PDFs on text alone (including table text).
Highlight only the blocks that differ:
  • Red   = deletions (old.pdf)
  • Green = insertions (new.pdf)
  • Yellow= replacements (both)
Produce:
  - old_annotated.pdf
  - new_annotated.pdf
  - diff_output.pdf (side-by-side)
"""

import fitz                  # PyMuPDF
import difflib
import logging
from tqdm import tqdm

# ─── Configuration (hard-coded) ────────────────────────────────────────────────
OLD_PATH      = "old.pdf"
NEW_PATH      = "new.pdf"
OLD_ANNOT     = "old_annotated.pdf"
NEW_ANNOT     = "new_annotated.pdf"
OUT_SIDEBYSIDE= "diff_output.pdf"

# Highlight colors (r,g,b) in [0,1], and opacity
COLOR_DEL = (1, 0, 0)   # red for deletions
COLOR_INS = (0, 1, 0)   # green for insertions
COLOR_REP = (1, 1, 0)   # yellow for replacements
ALPHA     = 0.3

# ─── Logging setup ─────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)


# ─── 1) Extract text blocks per page ────────────────────────────────────────────
def extract_text_blocks(path):
    """
    Return a list-of-pages, each page a list of (Rect, text)
    for every non-empty text block in reading order.
    """
    doc = fitz.open(path)
    pages = []
    logger.info("Extracting text blocks from %r (%d pages)…", path, doc.page_count)
    for page in tqdm(doc, desc="Extract blocks"):
        blks = []
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, text, *rest = b
            if text.strip():
                blks.append((fitz.Rect(x0, y0, x1, y1), text.strip()))
        # sort top→bottom, left→right
        blks.sort(key=lambda bt: (bt[0].y0, bt[0].x0))
        pages.append(blks)
    doc.close()
    return pages


# ─── 2) Diff two pages' text blocks ─────────────────────────────────────────────
def diff_text_blocks(old_blks, new_blks):
    """
    Given lists of (Rect,text) for old & new, return four lists of Rect:
      - deleted (in old not in new)
      - inserted
      - replaced_old
      - replaced_new
    """
    old_texts = [t for _, t in old_blks]
    new_texts = [t for _, t in new_blks]
    sm = difflib.SequenceMatcher(None, old_texts, new_texts)

    deleted, inserted, rep_o, rep_n = [], [], [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if   tag == "delete":
            deleted.extend(old_blks[i][0] for i in range(i1, i2))
        elif tag == "insert":
            inserted.extend(new_blks[j][0] for j in range(j1, j2))
        elif tag == "replace":
            rep_o.extend(old_blks[i][0] for i in range(i1, i2))
            rep_n.extend(new_blks[j][0] for j in range(j1, j2))
    return deleted, inserted, rep_o, rep_n


# ─── 3) Annotate two PDFs in place ─────────────────────────────────────────────
def annotate_pdfs(old_doc, new_doc, deletions, insertions, rep_old, rep_new):
    """
    Apply highlight-annotations:
      • deletions (list-of-lists of Rect) on old_doc in RED  
      • insertions on new_doc in GREEN  
      • replacements on both in YELLOW  
    """
    logger.info("Annotating OLD PDF…")
    for page_no, page in enumerate(tqdm(old_doc, desc="Old")):
        # deletions + replacements
        for r in deletions[page_no] + rep_old[page_no]:
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_DEL if r in deletions[page_no] else COLOR_REP)
            a.set_opacity(ALPHA)
            a.update()
    old_doc.save(OLD_ANNOT)

    logger.info("Annotating NEW PDF…")
    for page_no, page in enumerate(tqdm(new_doc, desc="New")):
        for r in insertions[page_no] + rep_new[page_no]:
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_INS if r in insertions[page_no] else COLOR_REP)
            a.set_opacity(ALPHA)
            a.update()
    new_doc.save(NEW_ANNOT)


# ─── 4) Render side-by-side ────────────────────────────────────────────────────
def render_side_by_side(old_doc, new_doc):
    """
    Create OUT_SIDEBYSIDE with old|new pages next to each other.
    """
    logger.info("Rendering side-by-side PDF…")
    combo = fitz.open()
    pages = min(old_doc.page_count, new_doc.page_count)

    for i in tqdm(range(pages), desc="Render SxS"):
        p_old, p_new = old_doc[i], new_doc[i]
        w1, h1 = p_old.rect.width, p_old.rect.height
        w2, h2 = p_new.rect.width, p_new.rect.height
        H = max(h1, h2)

        newp = combo.new_page(width=w1 + w2, height=H)
        pix1 = p_old.get_pixmap(alpha=False)
        pix2 = p_new.get_pixmap(alpha=False)

        newp.insert_image(fitz.Rect(0,    H - h1, w1,      H), pixmap=pix1)
        newp.insert_image(fitz.Rect(w1,   H - h2, w1 + w2, H), pixmap=pix2)

    combo.save(OUT_SIDEBYSIDE)
    logger.info("Saved side-by-side PDF to %r", OUT_SIDEBYSIDE)


# ─── Main Workflow ─────────────────────────────────────────────────────────────
def main():
    # 1) extract text blocks
    old_pages = extract_text_blocks(OLD_PATH)
    new_pages = extract_text_blocks(NEW_PATH)
    assert len(old_pages) == len(new_pages), "Page counts differ!"

    pages = len(old_pages)
    del_ann, ins_ann, ro_ann, rn_ann = ([] for _ in range(4))

    # 2) diff each page
    logger.info("Diffing text blocks on %d pages…", pages)
    for i in tqdm(range(pages), desc="Diff pages"):
        d, ins, ro, rn = diff_text_blocks(old_pages[i], new_pages[i])
        del_ann.append(d)
        ins_ann.append(ins)
        ro_ann.append(ro)
        rn_ann.append(rn)

    # 3) open docs and annotate
    old_doc = fitz.open(OLD_PATH)
    new_doc = fitz.open(NEW_PATH)
    annotate_pdfs(old_doc, new_doc, del_ann, ins_ann, ro_ann, rn_ann)

    # 4) render side-by-side
    render_side_by_side(old_doc, new_doc)

    logger.info("All done! → %r, %r, %r",
                OLD_ANNOT, NEW_ANNOT, OUT_SIDEBYSIDE)


if __name__ == "__main__":
    main()