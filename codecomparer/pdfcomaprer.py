import fitz              # PyMuPDF
import difflib
from PIL import Image, ImageChops
import logging
from tqdm import tqdm

# ─── 1. SETUP LOGGER ─────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ─── 2. TEXT EXTRACTION ─────────────────────────────────────────────────────────
def extract_text_blocks(doc):
    """
    Returns a list (per page) of [(Rect, text), …] for all non-empty text blocks.
    """
    logger.info("Extracting text blocks from %d pages…", len(doc))
    all_blocks = []
    for page in tqdm(doc, desc="Extract text"):
        blocks = []
        for b in page.get_text("blocks"):
            # PyMuPDF v1.x/v2.x both work with star-unpack
            x0, y0, x1, y1, text, *rest = b
            if text.strip():
                blocks.append((fitz.Rect(x0, y0, x1, y1), text))
        all_blocks.append(blocks)
    logger.info("Done extracting text.")
    return all_blocks

# ─── 3. TEXT DIFF ────────────────────────────────────────────────────────────────
def diff_text_blocks(old_blocks, new_blocks):
    """
    Given two lists of (Rect, text), returns dict of 
    { 'delete': [Rects], 'insert': [Rects], 'change': [Rects] }.
    """
    logger.info("Diffing text blocks…")
    old_texts = [t for _, t in old_blocks]
    new_texts = [t for _, t in new_blocks]
    sm = difflib.SequenceMatcher(None, old_texts, new_texts)
    
    ann = {"delete": [], "insert": [], "change": []}
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "delete":
            ann["delete"].extend(old_blocks[i][0] for i in range(i1, i2))
        elif tag == "insert":
            ann["insert"].extend(new_blocks[j][0] for j in range(j1, j2))
        elif tag == "replace":
            ann["change"].extend(old_blocks[i][0] for i in range(i1, i2))
            ann["change"].extend(new_blocks[j][0] for j in range(j1, j2))
    logger.info("Text diff done.")
    return ann

# ─── 4. IMAGE DIFF ───────────────────────────────────────────────────────────────
def diff_page_images(old_page, new_page, thresh=10):
    """
    Returns a list of bounding‐boxes where the two page bitmaps differ.
    """
    logger.info("Image‐diffing a page…")
    pm1 = old_page.get_pixmap(alpha=False)
    pm2 = new_page.get_pixmap(alpha=False)
    im1 = Image.frombytes("RGB", [pm1.width, pm1.height], pm1.samples)
    im2 = Image.frombytes("RGB", [pm2.width, pm2.height], pm2.samples)

    diff = ImageChops.difference(im1, im2).convert("L")
    bw = diff.point(lambda x: 255 if x > thresh else 0)
    bbox = bw.getbbox()
    boxes = [bbox] if bbox else []
    logger.info("Image‐diff done.")
    return boxes

# ─── 5. ANNOTATION ───────────────────────────────────────────────────────────────
def annotate_pages(doc, per_page_rects, color, alpha=0.25):
    """
    Draw translucent rectangles (color as (r,g,b)) on each page.
    per_page_rects is a list of lists of Rects.
    """
    logger.info("Annotating %d pages…", len(doc))
    for page, rects in tqdm(zip(doc, per_page_rects),
                            total=len(doc),
                            desc="Annotate"):
        shape = page.new_shape()
        for r in rects:
            shape.draw_rect(r)
            shape.finish(fill=color + (alpha,))
        shape.commit()
    logger.info("Annotation complete.")

# ─── 6. SIDE-BY-SIDE RENDER ──────────────────────────────────────────────────────
def render_side_by_side(old_doc, new_doc, out_path):
    """
    Creates a new PDF where each page is old|new side by side.
    """
    logger.info("Starting side-by-side render…")
    out = fitz.open()
    n = min(len(old_doc), len(new_doc))
    for i in tqdm(range(n), desc="Rendering pages"):
        p1, p2 = old_doc[i], new_doc[i]
        w1, h1 = p1.rect.width, p1.rect.height
        w2, h2 = p2.rect.width, p2.rect.height
        H = max(h1, h2)
        newp = out.new_page(width=w1 + w2, height=H)

        pix1 = p1.get_pixmap(alpha=False)
        pix2 = p2.get_pixmap(alpha=False)
        # place old at left, new at right, aligned to bottom
        newp.insert_image(fitz.Rect(0, H - h1, w1, H), pixmap=pix1)
        newp.insert_image(fitz.Rect(w1, H - h2, w1 + w2, H), pixmap=pix2)

    out.save(out_path)
    logger.info(f"Side-by-side PDF written to {out_path!r}")

# ─── 7. PUTTING IT ALL TOGETHER ─────────────────────────────────────────────────
if __name__ == "__main__":
    old_pdf = fitz.open("old.pdf")
    new_pdf = fitz.open("new.pdf")

    # 1) extract
    old_blocks = extract_text_blocks(old_pdf)
    new_blocks = extract_text_blocks(new_pdf)

    # 2) per-page diffs
    per_page_rects = []
    for ob, nb, p1, p2 in zip(old_blocks, new_blocks, old_pdf, new_pdf):
        txt_ann = diff_text_blocks(ob, nb)
        img_boxes = diff_page_images(p1, p2)
        # merge all rects for this page
        rects = txt_ann["delete"] + txt_ann["insert"] + txt_ann["change"] + img_boxes
        per_page_rects.append(rects)

    # 3) annotate old vs. new (you can choose different colors or even 
    #    produce two separate docs if you like)
    annotate_pages(old_pdf, per_page_rects, color=(1,0,0))  # red = deletions/old
    annotate_pages(new_pdf, per_page_rects, color=(0,1,0))  # green = inserts/new

    # 4) render side-by-side
    render_side_by_side(old_pdf, new_pdf, "diff_output.pdf")