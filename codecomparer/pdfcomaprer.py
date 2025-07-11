#!/usr/bin/env python3
"""
pdf_diff_exact.py

Hard-coded paths; only true diffs get highlighted.
"""

import fitz                  # PyMuPDF
import difflib
from PIL import Image, ImageChops
import logging
from tqdm import tqdm

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
OLD_PATH    = "old.pdf"
NEW_PATH    = "new.pdf"
OUT_PATH    = "diff_output.pdf"

PIXEL_THRESH = 20    # image-diff sensitivity
ALPHA        = 0.2   # annotation opacity

COLORS_OLD = {
    "delete": (1, 0, 0),  # red
    "replace": (0, 0, 1)  # blue
}
COLORS_NEW = {
    "insert": (0, 1, 0),  # green
    "replace": (0, 0, 1)  # blue
}

# ─── LOGGER SETUP ───────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# ─── TEXT EXTRACTION ────────────────────────────────────────────────────────────
def extract_text_blocks(doc):
    pages = []
    logger.info("Extracting text from %d pages…", doc.page_count)
    for page in tqdm(doc, desc="Extract text"):
        blks = []
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, text, *rest = b
            if text.strip():
                blks.append((fitz.Rect(x0, y0, x1, y1), text))
        pages.append(blks)
    return pages

# ─── TEXT DIFF (separate old/new replace rects) ─────────────────────────────────
def diff_text_blocks(old_blks, new_blks):
    old_texts = [t for _, t in old_blks]
    new_texts = [t for _, t in new_blks]
    sm = difflib.SequenceMatcher(None, old_texts, new_texts)

    delete_rects = []
    insert_rects = []
    replace_old = []
    replace_new = []

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "delete":
            delete_rects.extend(old_blks[i][0] for i in range(i1, i2))
        elif tag == "insert":
            insert_rects.extend(new_blks[j][0] for j in range(j1, j2))
        elif tag == "replace":
            replace_old.extend(old_blks[i][0] for i in range(i1, i2))
            replace_new.extend(new_blks[j][0] for j in range(j1, j2))

    return delete_rects, insert_rects, replace_old, replace_new

# ─── IMAGE DIFF ─────────────────────────────────────────────────────────────────
def diff_page_images(old_page, new_page):
    pm1 = old_page.get_pixmap(alpha=False)
    pm2 = new_page.get_pixmap(alpha=False)
    im1 = Image.frombytes("RGB", [pm1.width, pm1.height], pm1.samples)
    im2 = Image.frombytes("RGB", [pm2.width, pm2.height], pm2.samples)

    diff = ImageChops.difference(im1, im2).convert("L")
    bw   = diff.point(lambda x: 255 if x > PIXEL_THRESH else 0)
    bbox = bw.getbbox()
    return [fitz.Rect(bbox)] if bbox else []

# ─── ANNOTATION ────────────────────────────────────────────────────────────────
def annotate_exact(old_doc, new_doc,
                   del_ann, ins_ann, rep_old, rep_new, img_boxes):
    """
    del_ann[i], rep_old[i]: old-page rects
    ins_ann[i], rep_new[i]: new-page rects
    img_boxes[i]: rects to mark on both
    """
    logger.info("Annotating old PDF…")
    for page, dels, reps, imgs in tqdm(zip(old_doc, del_ann, rep_old, img_boxes),
                                       total=len(old_doc), desc="Old annots"):
        for r in dels:
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLORS_OLD["delete"])
            a.set_opacity(ALPHA); a.update()
        for r in reps + imgs:
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLORS_OLD["replace"])
            a.set_opacity(ALPHA); a.update()

    logger.info("Annotating new PDF…")
    for page, ins, reps, imgs in tqdm(zip(new_doc, ins_ann, rep_new, img_boxes),
                                      total=len(new_doc), desc="New annots"):
        for r in ins:
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLORS_NEW["insert"])
            a.set_opacity(ALPHA); a.update()
        for r in reps + imgs:
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLORS_NEW["replace"])
            a.set_opacity(ALPHA); a.update()

# ─── SIDE-BY-SIDE RENDER ────────────────────────────────────────────────────────
def render_side_by_side(old_doc, new_doc):
    out = fitz.open()
    pages = min(old_doc.page_count, new_doc.page_count)
    logger.info("Rendering %d pages side-by-side…", pages)

    for i in tqdm(range(pages), desc="Render"):
        p1, p2 = old_doc[i], new_doc[i]
        w1,h1 = p1.rect.width, p1.rect.height
        w2,h2 = p2.rect.width, p2.rect.height
        H = max(h1,h2)

        newp = out.new_page(width=w1+w2, height=H)
        pix1 = p1.get_pixmap(alpha=False)
        pix2 = p2.get_pixmap(alpha=False)
        newp.insert_image(fitz.Rect(0,    H-h1,   w1,     H),    pixmap=pix1)
        newp.insert_image(fitz.Rect(w1,   H-h2,   w1+w2,  H),    pixmap=pix2)

    out.save(OUT_PATH)
    logger.info("Saved %r", OUT_PATH)

# ─── MAIN WORKFLOW ─────────────────────────────────────────────────────────────
def main():
    old = fitz.open(OLD_PATH)
    new = fitz.open(NEW_PATH)
    pages = min(old.page_count, new.page_count)

    # 1) extract
    old_blocks = extract_text_blocks(old)
    new_blocks = extract_text_blocks(new)

    # prepare per-page lists
    del_ann, ins_ann = [], []
    rep_old, rep_new = [], []
    img_boxes        = []

    # 2) diff each page
    logger.info("Computing diffs on %d pages…", pages)
    for i in tqdm(range(pages), desc="Diffing"):
        d, ins, ro, rn = diff_text_blocks(old_blocks[i], new_blocks[i])
        imgs = diff_page_images(old[i], new[i])
        del_ann.append(d)
        ins_ann.append(ins)
        rep_old.append(ro)
        rep_new.append(rn)
        img_boxes.append(imgs)

    # 3) annotate only those rects
    annotate_exact(old, new, del_ann, ins_ann, rep_old, rep_new, img_boxes)

    # 4) side-by-side
    render_side_by_side(old, new)


if __name__ == "__main__":
    main()