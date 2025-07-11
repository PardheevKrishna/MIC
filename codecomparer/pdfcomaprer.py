import json
import fitz            # PyMuPDF
from tqdm import tqdm

# hard-coded paths
OLD, NEW = "old.pdf", "new.pdf"
OLD_OUT, NEW_OUT, SXS = "old_annotated.pdf", "new_annotated.pdf", "diff_output.pdf"

# highlight colors & opacity
COLORS = {
    "delete":  (1, 0, 0),   # red
    "insert":  (0, 1, 0),   # green
    "replace": (1, 1, 0),   # yellow
}
ALPHA = 0.3

# 1) Load the diff JSON
with open("diff.json") as f:
    data = json.load(f)

# 2) Open PDFs
old_doc = fitz.open(OLD)
new_doc = fitz.open(NEW)

# 3) Annotate each hunk
for h in tqdm(data["hunks"], desc="Applying highlights"):
    action = h["action"]
    pno     = h["page_num"]
    rect    = fitz.Rect(*h["bbox"])
    color   = COLORS[action]

    if action == "delete":
        page = old_doc[pno]
        a    = page.add_highlight_annot(rect)
        a.set_colors(fill=color); a.set_opacity(ALPHA); a.update()

    elif action == "insert":
        page = new_doc[pno]
        a    = page.add_highlight_annot(rect)
        a.set_colors(fill=color); a.set_opacity(ALPHA); a.update()

    else:  # replace
        for doc in (old_doc, new_doc):
            page = doc[pno]
            a    = page.add_highlight_annot(rect)
            a.set_colors(fill=color); a.set_opacity(ALPHA); a.update()

# 4) Save annotated PDFs
old_doc.save(OLD_OUT)
new_doc.save(NEW_OUT)

# 5) (Optional) render side-by-side
out = fitz.open()
pages = min(old_doc.page_count, new_doc.page_count)
for i in range(pages):
    po, pn = old_doc[i], new_doc[i]
    w1, h1 = po.rect.width, po.rect.height
    w2, h2 = pn.rect.width, pn.rect.height
    H = max(h1, h2)
    newp = out.new_page(width=w1 + w2, height=H)
    p1 = po.get_pixmap(alpha=False)
    p2 = pn.get_pixmap(alpha=False)
    newp.insert_image(fitz.Rect(0,    H-h1, w1,      H), pixmap=p1)
    newp.insert_image(fitz.Rect(w1,   H-h2, w1+w2,   H), pixmap=p2)
out.save(SXS)
print("Done →", OLD_OUT, NEW_OUT, SXS)