from PIL import Image
import numpy as np

p100 = Image.open('scratch/de23_pages/page_100.png')
# Câu 14:
# In scratch/c14_full.png (y was 2600 to 3600):
# Text "Câu 14... như hình sau:" ended at relative y=160 (i.e. 2600 + 160 = 2760)
# Text "Có bao nhiêu..." starts at relative y=870 (i.e. 2600 + 870 = 3470)
# So crop y: 2770 to 3460, x: 1300 to 3500
c14 = p100.crop((1400, 2770, 3400, 3460))
c14.save('San_Pham/images_crop_de23/p1_cau14_cis_pt_en.png')

# Câu 16:
# In scratch/c16_full.png (y was 5200 to 6300):
# Text "Câu 16... sau:" ended at relative y=140 (i.e. 5200 + 140 = 5340)
# Scheme starts around 5350, ends at relative y=810 (i.e. 5200 + 810 = 6010)
# Text "Những phản ứng..." starts at relative y=830 (i.e. 6030)
c16 = p100.crop((650, 5350, 4350, 6015))
c16.save('San_Pham/images_crop_de23/p1_cau16_so_do.png')

# Câu 17:
p101 = Image.open('scratch/de23_pages/page_101.png')
# In scratch/c17_full.png (y was 400 to 1600):
# Scheme starts around relative y=140 (i.e. 400 + 140 = 540)
# Scheme ends around relative y=870 (i.e. 400 + 870 = 1270)
# Text "Những phản ứng..." starts at relative y=900 (i.e. 1300)
c17 = p101.crop((700, 540, 4350, 1275))
c17.save('San_Pham/images_crop_de23/p1_cau17_so_do.png')

print("Saved cropped images. Auto trimming whitespace...")

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
