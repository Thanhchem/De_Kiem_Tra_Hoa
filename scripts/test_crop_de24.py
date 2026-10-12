import os
from PIL import Image

CROP_DIR = "San_Pham/images_crop_de24"
os.makedirs(CROP_DIR, exist_ok=True)

# Page dimensions in XML: w=595, h=841
# Rendered image size: w=5000, h=7069
scale_x = 5000.0 / 595.0
scale_y = 7069.0 / 841.0

def crop_box(page_num, l, t, w, h, out_name):
    img_path = f"scratch/de24_pages/page_{page_num}.png"
    im = Image.open(img_path)
    x1 = int(l * scale_x)
    y1 = int(t * scale_y)
    x2 = int((l + w) * scale_x)
    y2 = int((t + h) * scale_y)
    cropped = im.crop((x1, y1, x2, y2))
    out_path = os.path.join(CROP_DIR, out_name)
    cropped.save(out_path)
    print(f"Cropped {out_name}: box=({x1},{y1},{x2},{y2}), size={cropped.size}")

# 1. Page 105: Cau 16 Phuc chat [Fe(OH2)6]3+
# t=421 to 543
crop_box(105, 150, 435, 290, 110, "p1_cau16_fe_oh2_6.png")

# 2. Page 106: Cau 1 Part II - Quy trinh Kroll
# t=143 to 395
crop_box(106, 50, 145, 500, 250, "p2_cau1_quy_trinh_kroll.png")

# 3. Page 107: Cau 4 Part II - Phuc chat A va B
# t=77 to 260
crop_box(107, 50, 95, 500, 160, "p2_cau4_phuc_A_B.png")

# 4. Page 107: Cau 1 Part III - Cau truc [Co(NH3)4Cl2]+
# t=360 to 480
crop_box(107, 120, 375, 360, 110, "p3_cau1_co_nh3_4_cl2.png")
