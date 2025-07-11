import fitz                  # PyMuPDF
import difflib
from PIL import Image, ImageChops, ImageDraw
import imagehash
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def extract_text_blocks(doc):
    """Return a list per page of (block_rect, text)."""
    pages = []
    for page in doc:
        blocks = []
        for b in page.get_text("blocks"):
            x0,y0,x1,y1, text, _, _, _ = b
            # filter out whitespace
            if text.strip():
                blocks.append(((x0,y0,x1,y1), text))
        pages.append(blocks)
    return pages

def diff_text_blocks(old_blocks, new_blocks):
    """Match blocks by order and generate diffs."""
    old_texts = [t for _,t in old_blocks]
    new_texts = [t for _,t in new_blocks]
    sm = difflib.SequenceMatcher(None, old_texts, new_texts)
    ops = sm.get_opcodes()
    # ops is list of (tag, i1,i2, j1,j2)
    annotations = {"delete": [], "insert": [], "change": []}
    for tag, i1, i2, j1, j2 in ops:
        if tag == "equal":
            continue
        if tag == "delete":
            for idx in range(i1,i2):
                annotations["delete"].append(old_blocks[idx][0])
        elif tag == "insert":
            for idx in range(j1,j2):
                annotations["insert"].append(new_blocks[idx][0])
        else:  # replace
            for idx in range(i1,i2):
                annotations["change"].append(old_blocks[idx][0])
            for idx in range(j1,j2):
                annotations["change"].append(new_blocks[idx][0])
    return annotations

def diff_page_images(old_page, new_page, thresh=10):
    """Return bounding boxes where the two page‐images differ."""
    pm_old = old_page.get_pixmap(alpha=False)
    pm_new = new_page.get_pixmap(alpha=False)
    im_old = Image.frombytes("RGB", [pm_old.width, pm_old.height], pm_old.samples)
    im_new = Image.frombytes("RGB", [pm_new.width, pm_new.height], pm_new.samples)
    # pixel diff
    diff = ImageChops.difference(im_old, im_new).convert("L")
    # threshold to B&W
    bw = diff.point(lambda x: 255 if x>thresh else 0)
    # find connected components or simply bounding box of all nonzero
    bbox = bw.getbbox()  # coarse single‐region box
    return [bbox] if bbox else []

def annotate_page(page, annots, color, alpha=0.25):
    """Draw translucent rects onto a page object in place."""
    # color as RGB tuple e.g. (1,0,0)
    shape = page.new_shape()
    for rect in annots:
        x0,y0,x1,y1 = rect
        shape.draw_rect(fitz.Rect(x0,y0,x1,y1))
        shape.finish(fill=color + (alpha,))
    shape.commit()

def render_side_by_side(old_doc, new_doc, out_path):
    c = canvas.Canvas(out_path, pagesize=letter)
    w, h = letter
    for pno in range(min(len(old_doc), len(new_doc))):
        # left: old
        c.saveState()
        c.translate(0,0)
        c.doForm(c.acroForm)  # not needed; we’ll embed images
        img_old = old_doc[pno].get_pixmap(matrix=fitz.Matrix(w/old_doc[pno].rect.width,
                                                            h/old_doc[pno].rect.height))
        c.drawInlineImage(Image.frombytes("RGB", [img_old.width, img_old.height], img_old.samples),
                          0, 0, width=w/2, height=h)
        # right: new
        img_new = new_doc[pno].get_pixmap(matrix=fitz.Matrix(w/2/new_doc[pno].rect.width * (w/2),
                                                            h/new_doc[pno].rect.height))
        c.drawInlineImage(Image.frombytes("RGB", [img_new.width, img_new.height], img_new.samples),
                          w/2, 0, width=w/2, height=h)
        c.showPage()
    c.save()

if __name__ == "__main__":
    old = fitz.open("old.pdf")
    new = fitz.open("new.pdf")
    # per‐page
    for p in range(len(old)):
        oblocks = extract_text_blocks(old)[p]
        nblocks = extract_text_blocks(new)[p]
        ann = diff_text_blocks(oblocks, nblocks)
        diff_boxes = diff_page_images(old[p], new[p])
        # annotate
        annotate_page(old[p], ann["delete"], color=(1,0,0))
        annotate_page(new[p], ann["insert"], color=(0,1,0))
        annotate_page(old[p], ann["change"] + diff_boxes, color=(0,0,1))
        annotate_page(new[p], ann["change"] + diff_boxes, color=(0,0,1))
    render_side_by_side(old, new, "diff_output.pdf")