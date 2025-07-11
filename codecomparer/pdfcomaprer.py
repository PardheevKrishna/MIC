#!/usr/bin/env python3
"""
pdf_diff_highlight_fixed.py

Compare two PDF files in-place and highlight diffs, then render a side-by-side PDF.
All settings (file names, colors, thresholds) are hard-coded.
"""

import fitz                  # PyMuPDF
import difflib
from PIL import Image, ImageChops
import logging
from tqdm import tqdm

# ─── CONFIGURATION (hard-coded) ────────────────────────────────────────────────
OLD_PATH     = "old.pdf"
NEW_PATH     = "new.pdf"
OUTPUT_PATH  = "diff_output.pdf"

PIXEL_THRESH = 20           # sensitivity for image diffs
HIGHLIGHT_OP = 0.2          # opacity for all highlights

# Colors: (r, g, b) with values in [0,1]
COLORS_OLD   = {
    "delete": (1, 0, 0),    # red highlights on old = deletions
    "insert": (0, 0, 0),    # (unused on old)
    "change": (0, 0, 1),    # blue highlights on both = replacements
}
COLORS_NEW   = {
    "delete": (0, 0, 0),    # (unused on new)
    "insert": (0, 1, 0),    # green highlights on new = insertions
    "change": (0, 0, 1),    # blue for replacements
}

# ─── Logging setup ─────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)


def extract_text_blocks(doc):
    """Extract (Rect, text) for every non-empty text block, per page."""
    pages = []
    logger.info("Extracting text from %d pages…", doc.page_count)
    for page in tqdm(doc, desc="Extract text"):
        blocks = []
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, text, *rest = b
            if text.strip():
                blocks.append((fitz.Rect(x0, y0, x1, y1), text))
        pages.append(blocks)
    logger.info("Text extraction done.")
    return pages


def diff_text_blocks(old_blks, new_blks):
    """
    Diff two lists of (Rect, text), return dict of lists of Rects:
    'delete', 'insert', 'change'
    """
    old_txt = [t for _, t in old_blks]
    new_txt = [t for _, t in new_blks]
    sm = difflib.SequenceMatcher(None, old_txt, new_txt)

    ann = {"delete": [], "insert": [], "change": []}
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "delete":
            ann["delete"].extend(old_blks[i][0] for i in range(i1, i2))
        elif tag == "insert":
            ann["insert"].extend(new_blks[j][0] for j in range(j1, j2))
        elif tag == "replace":
            ann["change"].extend(old_blks[i][0] for i in range(i1, i2))
            ann["change"].extend(new_blks[j][0] for j in range(j1, j2))
    return ann


def diff_page_images(p_old, p_new):
    """Return list of Rects where rendered pages differ beyond PIXEL_THRESH."""
    pm1 = p_old.get_pixmap(alpha=False)
    pm2 = p_new.get_pixmap(alpha=False)
    im1 = Image.frombytes("RGB", [pm1.width, pm1.height], pm1.samples)
    im2 = Image.frombytes("RGB", [pm2.width, pm2.height], pm2.samples)

    diff = ImageChops.difference(im1, im2).convert("L")
    bw   = diff.point(lambda x: 255 if x > PIXEL_THRESH else 0)
    bbox = bw.getbbox()
    return [fitz.Rect(bbox)] if bbox else []


def annotate_doc(doc, delete_ann, insert_ann, change_ann, colors):
    """
    Apply real PDF highlight annots on each page of `doc`:
      • delete_ann[i] = list of Rects to highlight as deletions
      • insert_ann[i] = list for insertions
      • change_ann[i] = list for replacements
      • colors: dict mapping those tags → (r,g,b)
    """
    logger.info("Annotating %d-page PDF…", doc.page_count)
    for page_idx, page in enumerate(tqdm(doc, desc="Annotate")):
        for tag, rect_list in (
            ("delete", delete_ann),
            ("insert", insert_ann),
            ("change", change_ann),
        ):
            for r in rect_list[page_idx]:
                # highlight annotation preserves underlying content
                annot = page.add_highlight_annot(r)
                annot.set_colors(fill=colors[tag])
                annot.set_opacity(HIGHLIGHT_OP)
                annot.update()


def render_side_by_side(old_doc, new_doc):
    """
    Create a new PDF with each page laid out as [old | new] and save it.
    """
    logger.info("Rendering side-by-side PDF…")
    combo = fitz.open()
    page_count = min(old_doc.page_count, new_doc.page_count)

    for i in tqdm(range(page_count), desc="Render pages"):
        p1, p2 = old_doc[i], new_doc[i]
        w1, h1 = p1.rect.width, p1.rect.height
        w2, h2 = p2.rect.width, p2.rect.height
        H = max(h1, h2)

        newp = combo.new_page(width=w1 + w2, height=H)
        pix1 = p1.get_pixmap(alpha=False)
        pix2 = p2.get_pixmap(alpha=False)

        newp.insert_image(fitz.Rect(0,   H - h1, w1,       H), pixmap=pix1)
        newp.insert_image(fitz.Rect(w1,  H - h2, w1 + w2,  H), pixmap=pix2)

    combo.save(OUTPUT_PATH)
    logger.info("Saved combined PDF: %r", OUTPUT_PATH)


def main():
    logger.info("Opening PDFs…")
    old_doc = fitz.open(OLD_PATH)
    new_doc = fitz.open(NEW_PATH)
    pages   = min(old_doc.page_count, new_doc.page_count)

    # 1) extract text blocks
    old_blocks = extract_text_blocks(old_doc)
    new_blocks = extract_text_blocks(new_doc)

    # Prepare per-page lists
    del_ann = [[] for _ in range(pages)]
    ins_ann = [[] for _ in range(pages)]
    chg_ann = [[] for _ in range(pages)]

    # 2) compute diffs
    logger.info("Computing diffs on %d pages…", pages)
    for i in tqdm(range(pages), desc="Compute diffs"):
        txt = diff_text_blocks(old_blocks[i], new_blocks[i])
        img = diff_page_images(old_doc[i], new_doc[i])

        del_ann[i] = txt["delete"]
        ins_ann[i] = txt["insert"]
        chg_ann[i] = txt["change"] + img

    # 3) annotate PDFs
    annotate_doc(old_doc, del_ann, ins_ann, chg_ann, COLORS_OLD)
    annotate_doc(new_doc, del_ann, ins_ann, chg_ann, COLORS_NEW)

    # 4) render side-by-side and save
    render_side_by_side(old_doc, new_doc)


if __name__ == "__main__":
    main()