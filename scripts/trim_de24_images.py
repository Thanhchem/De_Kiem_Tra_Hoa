import os
from PIL import Image, ImageChops

CROP_DIR = "San_Pham/images_crop_de24"

def auto_trim(img_name, padding=15):
    p = os.path.join(CROP_DIR, img_name)
    im = Image.open(p).convert("RGB")
    # find bounding box of non-white pixels
    # background is white (255, 255, 255)
    bg = Image.new(im.mode, im.size, (255, 255, 255))
    diff = ImageChops.difference(im, bg)
    bbox = diff.getbbox()
    if bbox:
        x1 = max(0, bbox[0] - padding)
        y1 = max(0, bbox[1] - padding)
        x2 = min(im.size[0], bbox[2] + padding)
        y2 = min(im.size[1], bbox[3] + padding)
        trimmed = im.crop((x1, y1, x2, y2))
        trimmed.save(p)
        print(f"Trimmed {img_name}: old={im.size} -> new={trimmed.size}, bbox={bbox}")
    else:
        print(f"No non-white pixels found for {img_name}")

for name in [
    "p1_cau16_fe_oh2_6.png",
    "p2_cau1_quy_trinh_kroll.png",
    "p2_cau4_phuc_A_B.png",
    "p3_cau1_co_nh3_4_cl2.png"
]:
    auto_trim(name)
