import os
import pymupdf

# Let's inspect text lines on all pages of BO_24_DE_KIEM_TRA_HOA_12.pdf
# Wait, PyMuPDF can get text with formatting, flags, font properties!
doc = pymupdf.open("PDF_Goc/BO_24_DE_KIEM_TRA_HOA_12.pdf")

for p_num in range(102, 108):
    page = doc[p_num]
    # search for drawings / underlines (rectangles with small height)
    drawings = page.get_drawings()
    print(f"Page {p_num+1}: {len(drawings)} drawings/lines found")
    for d in drawings:
        rect = d["rect"]
        # underline usually has height < 3 and width between 5 and 20
        if rect.height < 3 and rect.width < 50:
            print(f"  Underline at: {rect}")
