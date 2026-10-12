from PIL import Image
import os

os.makedirs("San_Pham/images_crop_de03", exist_ok=True)

# 1. Cau 9 reaction on page 11
im11 = Image.open("scratch/de03_pages/page_11.png")
w11, h11 = im11.size
# Top line was around 0.46*h11. Let's skip the text line "Cau 9. Xa phong..."
# In test_cau9, text ended at around y = 0.473*h11
crop9 = im11.crop((int(0.12 * w11), int(0.472 * h11), int(0.90 * w11), int(0.605 * h11)))
crop9.save("San_Pham/images_crop_de03/cau9_phan_ung_xa_phong_hoa.png")

# 2. Part II Cau 1 diagram on page 13
im13 = Image.open("scratch/de03_pages/page_13.png")
w13, h13 = im13.size
# The black box is from y = 0.302*h13 to 0.388*h13, x = 0.35*w13 to 0.77*w13
crop_p2_c1 = im13.crop((int(0.351 * w13), int(0.301 * h13), int(0.767 * w13), int(0.389 * h13)))
crop_p2_c1.save("San_Pham/images_crop_de03/cau1_hydrogen_hoa_dau_an.png")

# 3. Part II Cau 2 reaction on page 13
# In test_p2_c2, text was at top. Reaction starts around 0.590*h13
crop_p2_c2 = im13.crop((int(0.20 * w13), int(0.588 * h13), int(0.88 * w13), int(0.685 * h13)))
crop_p2_c2.save("San_Pham/images_crop_de03/cau2_thuy_phan_aspirin.png")

# 4. Part II Cau 3 structures on page 13
# Reaction starts around 0.816*h13
crop_p2_c3 = im13.crop((int(0.10 * w13), int(0.812 * h13), int(0.90 * w13), int(0.880 * h13)))
crop_p2_c3.save("San_Pham/images_crop_de03/cau3_sodium_palmitate_laurylsulfate.png")

# 5. Part III Cau 2 structures on page 14
im14 = Image.open("scratch/de03_pages/page_14.png")
w14, h14 = im14.size
# Cau 2 solution is between y = 0.78*h14 and 0.87*h14
crop_p3_c2 = im14.crop((int(0.09 * w14), int(0.795 * h14), int(0.93 * w14), int(0.865 * h14)))
crop_p3_c2.save("San_Pham/images_crop_de03/p3_cau2_dong_phan_triglyceride.png")

print("All final cropped images generated successfully.")
