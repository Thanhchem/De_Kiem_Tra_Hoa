from PIL import Image

p100 = Image.open('scratch/de23_pages/page_100.png')

# Câu 14:
c14 = p100.crop((1400, 2810, 3400, 3550))
c14.save('San_Pham/images_crop_de23/p1_cau14_cis_pt_en.png')

# Câu 16:
# Top has a tiny speck of previous line, let's start at 5450.
# Bottom text (màu xanh nhạt) (màu xanh lam) needs to be complete, let's end at 6150.
c16 = p100.crop((650, 5450, 4350, 6150))
c16.save('San_Pham/images_crop_de23/p1_cau16_so_do.png')

# Câu 17:
p101 = Image.open('scratch/de23_pages/page_101.png')
# Diagram text ends at 1175, question line begins at 1220.
c17 = p101.crop((700, 540, 4350, 1180))
c17.save('San_Pham/images_crop_de23/p1_cau17_so_do.png')

print("Refined crops saved!")
