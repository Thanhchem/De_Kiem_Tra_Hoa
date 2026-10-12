import os
import pymupdf

OUT_DIR = "scratch/de24_pages"
os.makedirs(OUT_DIR, exist_ok=True)

doc = pymupdf.open("PDF_Goc/BO_24_DE_KIEM_TRA_HOA_12.pdf")
print("Total pages:", len(doc))

for p in range(102, len(doc)): # pages 103 to 108 (1-indexed)
    page = doc[p]
    pix = page.get_pixmap(dpi=300)
    out_path = os.path.join(OUT_DIR, f"page_{p+1}.png")
    pix.save(out_path)
    print(f"Rendered page {p+1} -> {out_path}")
