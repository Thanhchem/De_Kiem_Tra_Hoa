import os
from PIL import Image, ImageChops

CROP_DIR = "San_Pham/images_crop_de24"
os.makedirs(CROP_DIR, exist_ok=True)

scale_x = 5000.0 / 595.0
scale_y = 7069.0 / 841.0

def crop_and_clean(page_num, l, t, w, h, out_name):
    img_path = f"scratch/de24_pages/page_{page_num}.png"
    im = Image.open(img_path).convert("RGB")
    x1 = int(l * scale_x)
    y1 = int(t * scale_y)
    x2 = int((l + w) * scale_x)
    y2 = int((t + h) * scale_y)
    cropped = im.crop((x1, y1, x2, y2))
    
    # Auto-trim white borders
    bg = Image.new("RGB", cropped.size, (255, 255, 255))
    diff = ImageChops.difference(cropped, bg)
    bbox = diff.getbbox()
    if bbox:
        pad = 15
        bx1 = max(0, bbox[0] - pad)
        by1 = max(0, bbox[1] - pad)
        bx2 = min(cropped.size[0], bbox[2] + pad)
        by2 = min(cropped.size[1], bbox[3] + pad)
        final_img = cropped.crop((bx1, by1, bx2, by2))
    else:
        final_img = cropped
        
    out_path = os.path.join(CROP_DIR, out_name)
    final_img.save(out_path)
    print(f"Saved {out_name}: size={final_img.size}")

# 1. Page 105: Cau 16 Phuc chat [Fe(OH2)6]3+
# t=436 to 542, l=120 to 450
crop_and_clean(105, 120, 436, 350, 105, "p1_cau16_fe_oh2_6.png")

# 2. Page 106: Cau 1 Part II - Quy trinh Kroll
# t=158 to 395, l=55 to 540
crop_and_clean(106, 50, 158, 495, 238, "p2_cau1_quy_trinh_kroll.png")

# 3. Page 107: Cau 4 Part II - Phuc chat A va B
# t=95 to 258, l=55 to 540
crop_and_clean(107, 50, 95, 495, 162, "p2_cau4_phuc_A_B.png")

# 4. Page 107: Cau 1 Part III - Cau truc [Co(NH3)4Cl2]+
# t=376 to 480, l=80 to 500
crop_and_clean(107, 80, 376, 430, 104, "p3_cau1_co_nh3_4_cl2.png")
