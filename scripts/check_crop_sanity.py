from PIL import Image

for name in [
    "p1_cau16_fe_oh2_6.png",
    "p2_cau1_quy_trinh_kroll.png",
    "p2_cau4_phuc_A_B.png",
    "p3_cau1_co_nh3_4_cl2.png"
]:
    p = os.path.join("San_Pham/images_crop_de24", name)
    im = Image.open(p)
    print(f"=== {name} ({im.size}) ===")
