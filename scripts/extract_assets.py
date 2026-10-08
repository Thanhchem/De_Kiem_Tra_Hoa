import pymupdf
import os

pdf_path = r"C:\Users\Admin\.gemini\antigravity\brain\a6ca2d96-ae73-43ed-8b7f-ad044c4b259a\.user_uploaded\media_1791426343824.pdf"
out_dir = r"c:\Antigravity_Thanh\San_Pham\images_crop_101"
os.makedirs(out_dir, exist_ok=True)

doc = pymupdf.open(pdf_path)

for i, page in enumerate(doc):
    imgs = page.get_images()
    print(f"Page {i+1} has {len(imgs)} images")
    for idx, img in enumerate(imgs):
        xref = img[0]
        base_img = doc.extract_image(xref)
        ext = base_img["ext"]
        img_bytes = base_img["image"]
        w = base_img["width"]
        h = base_img["height"]
        fname = f"p{i+1}_img{idx}_{xref}.{ext}"
        target = os.path.join(out_dir, fname)
        with open(target, "wb") as f:
            f.write(img_bytes)
        print(f"  Saved {fname} ({w}x{h})")
