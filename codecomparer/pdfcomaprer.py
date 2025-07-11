#!/usr/bin/env python3
"""
pdf_diff_simple.py

Whole-document, block-level PDF diff & highlight.

Hard-coded paths/colors; no CLI args.
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

# RGB colors in [0,1]
COLOR_DELETE  = (1, 0, 0)    # red
COLOR_INSERT  = (0, 1, 0)    # green
COLOR_REPLACE = (1, 1, 0)    # yellow
ALPHA         = 0.3          # transparency

# ─── LOGGING ───────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


# ─── 1) FLATTEN ALL TEXT BLOCKS ─────────────────────────────────────────────────
def extract_blocks(path):
    """
    Read every non-empty text block from `path`, returning a list of dicts:
      { 'page': int, 'rect': fitz.Rect, 'text': str }
    """
    logger.info("Extracting text blocks from %r…", path)
    doc = fitz.open(path)
    blocks = []
    for pno in tqdm(range(doc.page_count), desc="Extract blocks"):
        page = doc[pno]
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, txt, *rest = b
            txt = txt.strip()
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


# ─── 2) DIFF AT BLOCK LEVEL ─────────────────────────────────────────────────────
def classify_blocks(old_blocks, new_blocks):
    """
    Run SequenceMatcher on the two block-lists’ .['text'] fields,
    and return four maps page→list[Rect]:
      - del_map    (deleted blocks)
      - ins_map    (inserted blocks)
      - rep_old    (replaced blocks on old)
      - rep_new    (replaced blocks on new)
    """
    old_texts = [b["text"] for b in old_blocks]
    new_texts = [b["text"] for b in new_blocks]

    logger.info("Running document-level diff…")
    sm = difflib.SequenceMatcher(None, old_texts, new_texts)
    del_map, ins_map, rep_old, rep_new = {}, {}, {}, {}

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
                rep_old.setdefault(blk["page"], []).append(blk["rect"])
            for idx in range(j1, j2):
                blk = new_blocks[idx]
                rep_new.setdefault(blk["page"], []).append(blk["rect"])
        # ‘equal’ → ignore

    return del_map, ins_map, rep_old, rep_new


# ─── 3) ANNOTATE & SAVE ─────────────────────────────────────────────────────────
def annotate_and_save(del_map, ins_map, rep_old, rep_new):
    # OLD PDF
    logger.info("Annotating OLD PDF…")
    old_doc = fitz.open(OLD_PATH)
    for pno, page in enumerate(tqdm(old_doc, desc="Old")):
        for r in del_map.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_DELETE)
            a.set_opacity(ALPHA); a.update()
        for r in rep_old.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_REPLACE)
            a.set_opacity(ALPHA); a.update()
    old_doc.save(OLD_ANNOT)
    logger.info("Saved %r", OLD_ANNOT)

    # NEW PDF
    logger.info("Annotating NEW PDF…")
    new_doc = fitz.open(NEW_PATH)
    for pno, page in enumerate(tqdm(new_doc, desc="New")):
        for r in ins_map.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_INSERT)
            a.set_opacity(ALPHA); a.update()
        for r in rep_new.get(pno, []):
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_REPLACE)
            a.set_opacity(ALPHA); a.update()
    new_doc.save(NEW_ANNOT)
    logger.info("Saved %r", NEW_ANNOT)

    return old_doc, new_doc


# ─── 4) SIDE-BY-SIDE RENDER ─────────────────────────────────────────────────────
def render_side_by_side(old_doc, new_doc):
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

        newp.insert_image(fitz.Rect(0,    H-h1,   w1,      H), pixmap=pix1)
        newp.insert_image(fitz.Rect(w1,   H-h2,   w1 + w2, H), pixmap=pix2)

    combo.save(OUT_SIDEBYSIDE)
    logger.info("Saved %r", OUT_SIDEBYSIDE)


# ─── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    old_blocks = extract_blocks(OLD_PATH)
    new_blocks = extract_blocks(NEW_PATH)

    del_map, ins_map, rep_old, rep_new = classify_blocks(old_blocks, new_blocks)
    old_doc, new_doc = annotate_and_save(del_map, ins_map, rep_old, rep_new)
    render_side_by_side(old_doc, new_doc)
    logger.info("Done!")

if __name__ == "__main__":
    main()