#!/usr/bin/env python3
"""
pdf_diff_with_ocr.py

OCR-based, whole-document PDF comparison & highlighting.
"""

import fitz                       # PyMuPDF
from pdf2image import convert_from_path
import pytesseract
import difflib
import tempfile
import os
import logging
from collections import defaultdict, Counter
from tqdm import tqdm

# ─── CONFIG ─────────────────────────────────────────────────────────────────────
OLD_PATH        = "old.pdf"
NEW_PATH        = "new.pdf"
OLD_OCR_ANN     = "old_ocr_annotated.pdf"
NEW_OCR_ANN     = "new_ocr_annotated.pdf"
OUT_SIDEBYSIDE  = "ocr_diff_output.pdf"

# Colors and opacity
COLOR_DEL = (1, 0, 0)     # red
COLOR_INS = (0, 1, 0)     # green
ALPHA     = 0.3

# DPI for rasterizing PDF → image
DPI = 300

# temporary directory for images
TMPDIR = tempfile.mkdtemp(prefix="pdf_ocr_diff_")

# logging
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


def ocr_extract_lines(pdf_path):
    """
    1) Convert each PDF page to an image.
    2) OCR via pytesseract.image_to_data to get per-word bboxes.
    3) Group words by line_num into lines, compute each line's bbox and text.
    Returns a flat list of dicts: {page, rect, text}.
    """
    logger.info(f"OCR-extracting lines from {pdf_path!r}")
    pages = convert_from_path(pdf_path, dpi=DPI, fmt="png", output_folder=TMPDIR)
    all_lines = []
    for pno, pil_im in enumerate(tqdm(pages, desc="OCR pages")):
        # get a pandas-like TSV of word-level boxes
        data = pytesseract.image_to_data(pil_im, output_type=pytesseract.Output.DICT)
        n = len(data["level"])
        # group words into lines by line_num
        lines = defaultdict(lambda: {"words": [], "bboxes": []})
        for i in range(n):
            txt = data["text"][i].strip()
            if not txt:
                continue
            ln = data["line_num"][i]
            x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
            lines[ln]["words"].append(txt)
            lines[ln]["bboxes"].append((x, y, x + w, y + h))

        # build line records
        for ln, info in lines.items():
            text = " ".join(info["words"])
            xs = [b[0] for b in info["bboxes"]]
            ys = [b[1] for b in info["bboxes"]]
            xe = [b[2] for b in info["bboxes"]]
            ye = [b[3] for b in info["bboxes"]]
            rect = fitz.Rect(min(xs), min(ys), max(xe), max(ye))
            all_lines.append({"page": pno, "rect": rect, "text": text})
    logger.info(f"OCR-extracted {len(all_lines)} lines from {pdf_path!r}")
    return all_lines


def global_diff_and_maps(old_lines, new_lines):
    """
    Perform an order-independent whole-document diff:
      - Count occurrences of each line-text in old vs new.
      - Matched = min(counts); extra in old are deletions, extra in new are insertions.
    Returns two maps: delete_map & insert_map mapping page_no→list of rects.
    """
    logger.info("Computing global diff over OCR lines…")
    old_count = Counter(l["text"] for l in old_lines)
    new_count = Counter(l["text"] for l in new_lines)
    match_count = {t: min(old_count[t], new_count[t]) for t in old_count}

    old_by_text = defaultdict(list)
    for l in old_lines:
        old_by_text[l["text"]].append(l)
    new_by_text = defaultdict(list)
    for l in new_lines:
        new_by_text[l["text"]].append(l)

    delete_map = defaultdict(list)
    insert_map = defaultdict(list)

    # any extra occurrences in old → deletions
    for text, lst in old_by_text.items():
        m = match_count.get(text, 0)
        for extra in lst[m:]:
            delete_map[extra["page"]].append(extra["rect"])

    # any extra in new → insertions
    for text, lst in new_by_text.items():
        m = match_count.get(text, 0)
        for extra in lst[m:]:
            insert_map[extra["page"]].append(extra["rect"])

    return delete_map, insert_map


def annotate_and_save(delete_map, insert_map):
    # OLD
    logger.info("Annotating OLD PDF (deletions)…")
    old_doc = fitz.open(OLD_PATH)
    for pno, page in enumerate(tqdm(old_doc, desc="OLD annotate")):
        for rect in delete_map.get(pno, []):
            a = page.add_rect_annot(rect)             # outline red rect
            a.set_colors(stroke=COLOR_DEL)
            a.set_border(width=1)
            a.set_opacity(ALPHA)
            a.update()
    old_doc.save(OLD_OCR_ANN)
    logger.info(f"Saved {OLD_OCR_ANN!r}")

    # NEW
    logger.info("Annotating NEW PDF (insertions)…")
    new_doc = fitz.open(NEW_PATH)
    for pno, page in enumerate(tqdm(new_doc, desc="NEW annotate")):
        for rect in insert_map.get(pno, []):
            a = page.add_rect_annot(rect)
            a.set_colors(stroke=COLOR_INS)
            a.set_border(width=1)
            a.set_opacity(ALPHA)
            a.update()
    new_doc.save(NEW_OCR_ANN)
    logger.info(f"Saved {NEW_OCR_ANN!r}")

    return fitz.open(OLD_OCR_ANN), fitz.open(NEW_OCR_ANN)


def render_side_by_side(old_doc, new_doc):
    logger.info("Rendering side-by-side PDF…")
    out = fitz.open()
    pages = min(old_doc.page_count, new_doc.page_count)
    for i in tqdm(range(pages), desc="Render SxS"):
        po, pn = old_doc[i], new_doc[i]
        w1, h1 = po.rect.width, po.rect.height
        w2, h2 = pn.rect.width, pn.rect.height
        H = max(h1, h2)
        newp = out.new_page(width=w1 + w2, height=H)
        pix1 = po.get_pixmap(alpha=False)
        pix2 = pn.get_pixmap(alpha=False)
        newp.insert_image(fitz.Rect(0,    H-h1,   w1,      H), pixmap=pix1)
        newp.insert_image(fitz.Rect(w1,   H-h2,   w1+w2,   H), pixmap=pix2)
    out.save(OUT_SIDEBYSIDE)
    logger.info(f"Saved {OUT_SIDEBYSIDE!r}")


def cleanup():
    for fn in os.listdir(TMPDIR):
        os.remove(os.path.join(TMPDIR, fn))
    os.rmdir(TMPDIR)


def main():
    old_lines = ocr_extract_lines(OLD_PATH)
    new_lines = ocr_extract_lines(NEW_PATH)
    delete_map, insert_map = global_diff_and_maps(old_lines, new_lines)
    old_doc, new_doc = annotate_and_save(delete_map, insert_map)
    render_side_by_side(old_doc, new_doc)
    cleanup()
    logger.info("Done.")

if __name__ == "__main__":
    main()