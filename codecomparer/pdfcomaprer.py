#!/usr/bin/env python3
"""
pdf_diff_inline.py

Whole-doc, inline-diff PDF highlighter.
Hard-coded paths/colors; no CLI args.
"""

import fitz               # PyMuPDF
import difflib
import logging
from tqdm import tqdm

# ─── CONFIGURATION ─────────────────────────────────────────────────────────────
OLD_PATH       = "old.pdf"
NEW_PATH       = "new.pdf"
OLD_ANNOT      = "old_annotated.pdf"
NEW_ANNOT      = "new_annotated.pdf"
OUT_SIDEBYSIDE = "diff_output.pdf"

# Colors (r,g,b) in [0,1]
COLOR_DELETE  = (1, 0, 0)    # red
COLOR_INSERT  = (0, 1, 0)    # green
COLOR_REPLACE = (1, 1, 0)    # yellow
ALPHA         = 0.3          # transparency

# ─── LOGGING SETUP ─────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)


# ─── 1) EXTRACT & FLATTEN TEXT BLOCKS ───────────────────────────────────────────
def extract_blocks(path):
    """
    Returns a list of {'page':int, 'rect':Rect, 'text':str}
    for every non-empty text block in reading order.
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


# ─── 2) FIND RECTANGLES FOR A SUBSTRING WITHIN A BLOCK ─────────────────────────
def find_and_annotate(page, substring, clip_rect, color):
    """
    search_for substring clipped to block rect, highlight each hit.
    """
    if not substring:
        return
    # search_for may miss weird whitespace; strip substring
    hits = page.search_for(substring, clip=clip_rect)
    for r in hits:
        annot = page.add_highlight_annot(r)
        annot.set_colors(fill=color)
        annot.set_opacity(ALPHA)
        annot.update()


# ─── 3) MAIN DIFF & ANNOTATION ─────────────────────────────────────────────────
def diff_and_annotate():
    # 1) extract flattened blocks
    old_blocks = extract_blocks(OLD_PATH)
    new_blocks = extract_blocks(NEW_PATH)

    old_texts = [b["text"] for b in old_blocks]
    new_texts = [b["text"] for b in new_blocks]

    # 2) whole-doc diff
    logger.info("Running document-level diff…")
    sm = difflib.SequenceMatcher(None, old_texts, new_texts)
    ops = sm.get_opcodes()

    # 3) open documents for annotation
    old_doc = fitz.open(OLD_PATH)
    new_doc = fitz.open(NEW_PATH)

    # 4) iterate ops
    logger.info("Annotating differences…")
    for tag, i1, i2, j1, j2 in tqdm(ops, desc="Diff ops"):
        if tag == "delete":
            # pure deletions → highlight whole block red on OLD
            for idx in range(i1, i2):
                blk = old_blocks[idx]
                p = old_doc[blk["page"]]
                a = p.add_highlight_annot(blk["rect"])
                a.set_colors(fill=COLOR_DELETE)
                a.set_opacity(ALPHA); a.update()

        elif tag == "insert":
            # pure insertions → highlight whole block green on NEW
            for idx in range(j1, j2):
                blk = new_blocks[idx]
                p = new_doc[blk["page"]]
                a = p.add_highlight_annot(blk["rect"])
                a.set_colors(fill=COLOR_INSERT)
                a.set_opacity(ALPHA); a.update()

        elif tag == "replace":
            # 1:1 block mapping for inline diff
            count_old = i2 - i1
            count_new = j2 - j1
            common = min(count_old, count_new)

            # inline diff within each paired block
            for k in range(common):
                ob = old_blocks[i1 + k]
                nb = new_blocks[j1 + k]
                page_o = old_doc[ob["page"]]
                page_n = new_doc[nb["page"]]

                # character-level diff of the block texts
                sm2 = difflib.SequenceMatcher(None, ob["text"], nb["text"])
                for s_tag, a1, a2, b1, b2 in sm2.get_opcodes():
                    if s_tag == "equal":
                        continue
                    # deletions or replacements in old
                    if s_tag in ("delete", "replace"):
                        substr = ob["text"][a1:a2]
                        find_and_annotate(page_o, substr, ob["rect"],
                                          COLOR_DELETE if s_tag=="delete" else COLOR_REPLACE)
                    # insertions or replacements in new
                    if s_tag in ("insert", "replace"):
                        substr = nb["text"][b1:b2]
                        find_and_annotate(page_n, substr, nb["rect"],
                                          COLOR_INSERT if s_tag=="insert" else COLOR_REPLACE)

            # extra unmatched blocks
            if count_old > common:
                for k in range(common, count_old):
                    ob = old_blocks[i1 + k]
                    p = old_doc[ob["page"]]
                    a = p.add_highlight_annot(ob["rect"])
                    a.set_colors(fill=COLOR_DELETE)
                    a.set_opacity(ALPHA); a.update()
            if count_new > common:
                for k in range(common, count_new):
                    nb = new_blocks[j1 + k]
                    p = new_doc[nb["page"]]
                    a = p.add_highlight_annot(nb["rect"])
                    a.set_colors(fill=COLOR_INSERT)
                    a.set_opacity(ALPHA); a.update()

        # ‘equal’ → do nothing

    # 5) save annotated PDFs
    old_doc.save(OLD_ANNOT)
    logger.info(f"Saved {OLD_ANNOT!r}")
    new_doc.save(NEW_ANNOT)
    logger.info(f"Saved {NEW_ANNOT!r}")

    # 6) side-by-side render
    render_side_by_side(old_doc, new_doc)


# ─── 4) SIDE-BY-SIDE RENDER ─────────────────────────────────────────────────────
def render_side_by_side(old_doc, new_doc):
    logger.info("Rendering side-by-side PDF…")
    combo = fitz.open()
    count = min(old_doc.page_count, new_doc.page_count)

    for i in tqdm(range(count), desc="Render SxS"):
        p_old, p_new = old_doc[i], new_doc[i]
        w1, h1 = p_old.rect.width, p_old.rect.height
        w2, h2 = p_new.rect.width, p_new.rect.height
        H = max(h1, h2)

        newp = combo.new_page(width=w1 + w2, height=H)
        pix1 = p_old.get_pixmap(alpha=False)
        pix2 = p_new.get_pixmap(alpha=False)
        newp.insert_image(fitz.Rect(0,    H-h1,   w1,      H), pixmap=pix1)
        newp.insert_image(fitz.Rect(w1,   H-h2,   w1+w2,   H), pixmap=pix2)

    combo.save(OUT_SIDEBYSIDE)
    logger.info(f"Saved side-by-side PDF to {OUT_SIDEBYSIDE!r}")


if __name__ == "__main__":
    diff_and_annotate()