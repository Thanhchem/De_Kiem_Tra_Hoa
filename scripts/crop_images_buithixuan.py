# -*- coding: utf-8 -*-
import os
from PIL import Image

OUT_DIR = os.path.abspath("San_Pham/images_ck1_hoa11_buithixuan")
os.makedirs(OUT_DIR, exist_ok=True)

IMG1 = r"C:\Users\Admin\.gemini\antigravity\brain\d5f03847-7882-4778-b376-e98bb33a9fea\.user_uploaded\media_1795397850320_90d9eae8.jpg"
IMG2 = r"C:\Users\Admin\.gemini\antigravity\brain\d5f03847-7882-4778-b376-e98bb33a9fea\.user_uploaded\media_1795397856349_04557a03.jpg"
IMG3 = r"C:\Users\Admin\.gemini\antigravity\brain\d5f03847-7882-4778-b376-e98bb33a9fea\.user_uploaded\media_1795397862798_71585d40.jpg"

# 1. Câu 7 Trang 1: Phổ khối lượng
im1 = Image.open(IMG1)
# Bounding box x1, y1, x2, y2
crop_c7 = im1.crop((280, 620, 600, 755))
crop_c7.save(os.path.join(OUT_DIR, "ms_cau7.png"))

# 2. Câu 2 Phần II Trang 2: Phổ IR & Bảng tra
im2 = Image.open(IMG2)
crop_ir = im2.crop((190, 520, 735, 695))
crop_ir.save(os.path.join(OUT_DIR, "ir_cau2_p2.png"))

# 3. Câu 4 Phần III Trang 3: CTCT Geraniol
im3 = Image.open(IMG3)
crop_geraniol = im3.crop((400, 360, 525, 455))
crop_geraniol.save(os.path.join(OUT_DIR, "geraniol_cau4_p3.png"))

# 4. Câu 3 Tự luận Trang 3: Phổ khối lượng acetic acid
crop_ms_acetic = im3.crop((260, 650, 640, 780))
crop_ms_acetic.save(os.path.join(OUT_DIR, "ms_acetic_acid_cau3_tl.png"))

print("Đã cắt xong 4 ảnh!")
