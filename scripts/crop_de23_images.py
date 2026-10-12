import sys
sys.stdout.reconfigure(encoding='utf-8')
from PIL import Image
import numpy as np

# Page 100 of PDF is Trang 99 of book
p100 = Image.open('scratch/de23_pages/page_100.png')
w100, h100 = p100.size

# Let's save vertical slices around Câu 14 and Câu 16 on page_100
# t=317 -> 317/841 * 7069 = ~2664
# t=454 -> 454/841 * 7069 = ~3815
c14_crop = p100.crop((1200, 2750, 3800, 3550))
c14_crop.save('San_Pham/images_crop_de23/p1_cau14_cis_pt_en.png')

# Câu 16:
# t=631 -> ~5300
c16_crop = p100.crop((450, 5380, 4550, 5750))
c16_crop.save('San_Pham/images_crop_de23/p1_cau16_so_do.png')

# Câu 17:
# Câu 17 text is at bottom of page 100, diagram is at top of page 101!
p101 = Image.open('scratch/de23_pages/page_101.png')
# Let's crop top of page 101 (below header)
c17_crop = p101.crop((450, 600, 4550, 1450))
c17_crop.save('San_Pham/images_crop_de23/p1_cau17_so_do.png')

def auto_trim(img_path, pad=15, threshold=240):
    img = Image.open(img_path).convert('RGB')
    arr = np.array(img)
    mask = np.any(arr < threshold, axis=2)
    rows = np.any(mask, axis=1)
    cols = np.any(mask, axis=0)
    if not np.any(rows):
        print(f"Warning: {img_path} is empty!")
        return
    rmin, rmax = np.where(rows)[0][[0, -1]]
    cmin, cmax = np.where(cols)[0][[0, -1]]
    
    rmin = max(0, rmin - pad)
    rmax = min(img.height, rmax + pad)
    cmin = max(0, cmin - pad)
    cmax = min(img.width, cmax + pad)
    
    trimmed = img.crop((cmin, rmin, cmax, rmax))
    trimmed.save(img_path)
    print(f"Trimmed {img_path}: {trimmed.size}")

auto_trim('San_Pham/images_crop_de23/p1_cau14_cis_pt_en.png')
auto_trim('San_Pham/images_crop_de23/p1_cau16_so_do.png')
auto_trim('San_Pham/images_crop_de23/p1_cau17_so_do.png')
