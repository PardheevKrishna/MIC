#!/usr/bin/env python3
"""
pdf_diff_text_tables.py

Compare two PDFs *only* on text (including tables).  
Highlight deleted lines/cells in RED on the OLD PDF,  
inserted in GREEN on the NEW PDF,  
and replaced lines/cells in YELLOW on *both*.

Produces:
  - old_annotated.pdf
  - new_annotated.pdf
  - diff_output.pdf   (side-by-side)
"""

import fitz                  # PyMuPDF
import pdfplumber
import difflib
import pandas as pd
import logging
from tqdm import tqdm

# ─── CONFIG ─────────────────────────────────────────────────────────────────────
OLD_PATH       = "old.pdf"
NEW_PATH       = "new.pdf"
OUT_SIDEBYSIDE = "diff_output.pdf"
OLD_ANNOT      = "old_annotated.pdf"
NEW_ANNOT      = "new_annotated.pdf"

ALPHA         = 0.3    # transparency of highlights
COLOR_DEL     = (1, 0, 0)   # red   = deletions
COLOR_INS     = (0, 1, 0)   # green = insertions
COLOR_REP     = (1, 1, 0)   # yellow= replacements

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


# ─── 1) EXTRACT TEXT LINES ─────────────────────────────────────────────────────
def extract_text_lines(path):
    """
    Returns list-of-pages, each a list of (Rect, text)
    for every non-empty line on that page.
    """
    doc = fitz.open(path)
    pages = []
    logger.info(f"Extracting text lines from {path!r} (%d pages)…", doc.page_count)
    for page in tqdm(doc, desc="Text→lines"):
        lines = []
        d = page.get_text("dict")
        for blk in d["blocks"]:
            if blk["type"] != 0:          # 0 = text block
                continue
            for ln in blk["lines"]:
                text = "".join(span["text"] for span in ln["spans"]).strip()
                if not text:
                    continue
                # compute the bounding box of the whole line
                x0 = min(span["bbox"][0] for span in ln["spans"])
                y0 = min(span["bbox"][1] for span in ln["spans"])
                x1 = max(span["bbox"][2] for span in ln["spans"])
                y1 = max(span["bbox"][3] for span in ln["spans"])
                lines.append((fitz.Rect(x0, y0, x1, y1), text))
        pages.append(lines)
    return pages


# ─── 2) EXTRACT TABLE CELLS ─────────────────────────────────────────────────────
def extract_table_cells(path):
    """
    Returns list-of-pages, each a list of tables;
    each table is (DataFrame, list-of-cell-dicts)
    where each cell-dict has keys: row_idx, col_idx, bbox (as fitz.Rect), text.
    """
    tables_per_page = []
    logger.info(f"Extracting tables from {path!r}…")
    with pdfplumber.open(path) as pdf:
        for page in tqdm(pdf.pages, desc="PDF→tables"):
            page_tables = []
            for tbl in page.find_tables():
                # 1) get the table data as DataFrame
                raw = tbl.extract_table()
                if not raw or len(raw) < 2:
                    continue
                df = pd.DataFrame(raw[1:], columns=raw[0])
                # 2) gather each cell’s bbox & its row/col index
                cells = []
                for cell in tbl.cells:
                    # pdfplumber Cell: cell.row_index, cell.col_index, cell.bbox, cell.text
                    cells.append({
                        "row":      cell["row_idx"],
                        "col":      cell["col_idx"],
                        "bbox":     fitz.Rect(cell["bbox"]),
                        "text":     cell["text"] or ""
                    })
                page_tables.append((df, cells))
            tables_per_page.append(page_tables)
    return tables_per_page


# ─── 3) DIFF TEXT LINES ─────────────────────────────────────────────────────────
def diff_lines(old_lines, new_lines):
    """
    Given two lists of (Rect, text), produce four lists of Rects:
      - deleted (in old not in new)
      - inserted
      - replaced_old
      - replaced_new
    """
    o_txt = [t for _, t in old_lines]
    n_txt = [t for _, t in new_lines]
    sm = difflib.SequenceMatcher(None, o_txt, n_txt)
    del_, ins, rep_o, rep_n = [], [], [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "delete":
            del_.extend(old_lines[i][0] for i in range(i1, i2))
        elif tag == "insert":
            ins .extend(new_lines[j][0] for j in range(j1, j2))
        elif tag == "replace":
            rep_o.extend(old_lines[i][0] for i in range(i1, i2))
            rep_n.extend(new_lines[j][0] for j in range(j1, j2))
    return del_, ins, rep_o, rep_n


# ─── 4) DIFF TABLE CELLS ─────────────────────────────────────────────────────────
def diff_table_cells(tbl_old, tbl_new):
    """
    tbl_old/tbl_new = (DataFrame, list-of-cell-dicts).
    Returns four lists of Rect:
      - deleted_cells (in old only)
      - inserted_cells
      - changed_old
      - changed_new
    """
    df_o, cells_o = tbl_old
    df_n, cells_n = tbl_new

    # unify shapes
    idxs = df_o.index.union(df_n.index)
    cols  = df_o.columns.union(df_n.columns)
    o2 = df_o.reindex(index=idxs, columns=cols, fill_value="")
    n2 = df_n.reindex(index=idxs, columns=cols, fill_value="")

    del_c, ins_c, co, cn = [], [], [], []
    # old cells
    for c in cells_o:
        ro, coo = c["row"], c["col"]
        t_o, t_n = o2.iat[ro, coo], n2.iat[ro, coo]
        if t_o and not t_n:
            del_c.append(c["bbox"])
        elif t_o and t_n and t_o != t_n:
            co.append(c["bbox"])
    # new cells
    for c in cells_n:
        ro, coo = c["row"], c["col"]
        t_o, t_n = o2.iat[ro, coo], n2.iat[ro, coo]
        if t_n and not t_o:
            ins_c.append(c["bbox"])
        elif t_o and t_n and t_o != t_n:
            cn.append(c["bbox"])
    return del_c, ins_c, co, cn


# ─── 5) ANNOTATE AND SAVE ───────────────────────────────────────────────────────
def annotate_and_save(old_path, new_path):
    # extract everything
    old_lines = extract_text_lines(old_path)
    new_lines = extract_text_lines(new_path)
    old_tabs  = extract_table_cells(old_path)
    new_tabs  = extract_table_cells(new_path)

    old_doc = fitz.open(old_path)
    new_doc = fitz.open(new_path)
    pages   = min(old_doc.page_count, new_doc.page_count)

    # prepare per‐page buckets
    del_l, ins_l, rep_lo, rep_ln = ([] for _ in range(4))
    del_c, ins_c, rep_co, rep_cn = ([] for _ in range(4))

    # diff page by page
    logger.info(f"Diffing text & tables across {pages} pages…")
    for i in tqdm(range(pages), desc="Diff pages"):
        # text
        dL, iL, rLo, rLn = diff_lines(old_lines[i], new_lines[i])
        del_l.append(dL); ins_l.append(iL)
        rep_lo.append(rLo); rep_ln.append(rLn)
        # tables: zip same‐index tables; extras ignored
        dC,iC,rCo,rCn = [], [], [], []
        for (to, co), (tn, cn) in zip(old_tabs[i], new_tabs[i]):
            x0,x1,x2,x3 = diff_table_cells((to, co), (tn, cn))
            dC.extend(x0); iC.extend(x1)
            rCo.extend(x2); rCn.extend(x3)
        del_c.append(dC); ins_c.append(iC)
        rep_co.append(rCo); rep_cn.append(rCn)

    # annotate OLD
    logger.info("Annotating OLD PDF…")
    for p in tqdm(range(pages), desc="Annot old"):
        page = old_doc[p]
        for r in del_l[p]   + del_c[p]:
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_DEL); a.set_opacity(ALPHA); a.update()
        for r in rep_lo[p]  + rep_co[p]:
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_REP); a.set_opacity(ALPHA); a.update()
    old_doc.save(OLD_ANNOT)

    # annotate NEW
    logger.info("Annotating NEW PDF…")
    for p in tqdm(range(pages), desc="Annot new"):
        page = new_doc[p]
        for r in ins_l[p]   + ins_c[p]:
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_INS); a.set_opacity(ALPHA); a.update()
        for r in rep_ln[p]  + rep_cn[p]:
            a = page.add_highlight_annot(r)
            a.set_colors(fill=COLOR_REP); a.set_opacity(ALPHA); a.update()
    new_doc.save(NEW_ANNOT)

    # side-by-side
    logger.info("Rendering side-by-side output…")
    combo = fitz.open()
    for i in tqdm(range(pages), desc="Render SxS"):
        p1, p2 = old_doc[i], new_doc[i]
        w1,h1 = p1.rect.width, p1.rect.height
        w2,h2 = p2.rect.width, p2.rect.height
        H = max(h1,h2)
        newp = combo.new_page(width=w1+w2, height=H)
        pix1 = p1.get_pixmap(alpha=False)
        pix2 = p2.get_pixmap(alpha=False)
        newp.insert_image(fitz.Rect(0,   H-h1,  w1,     H),    pixmap=pix1)
        newp.insert_image(fitz.Rect(w1,  H-h2,  w1+w2,  H),    pixmap=pix2)
    combo.save(OUT_SIDEBYSIDE)

    logger.info(f"Done. → {OLD_ANNOT}, {NEW_ANNOT}, {OUT_SIDEBYSIDE}")


if __name__ == "__main__":
    annotate_and_save(OLD_PATH, NEW_PATH)