# -*- coding: utf-8 -*-
"""
Script sửa chữa triệt để cho câu hỏi C1-P1-TH-069:
1. Trích xuất và lưu ảnh công thức cấu tạo ester X vào images_gk1_hoa12.
2. Cập nhật cơ sở dữ liệu JSON master (Ngan_Hang_Hoa_12_Tron_Bo_9_Chuong.json & Ngan_Hang_Hoa_12_Chuong_1_2_3.json).
3. Tái tạo Đề 20 (Đề có Câu 1 là C1-P1-TH-069) gồm cả 3 bản: HocSinh, GiaoVien, Azota ở cả 2 thư mục.
4. Cập nhật trực tiếp file Word Ngân hàng Chương 1 (HocSinh và GiaoVien) ở cả 2 thư mục.
"""

import os
import re
import io
import sys
import json
import shutil
import docx
from PIL import Image
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

IMG_DIR = r"c:\Antigravity_Thanh\San_Pham\images_gk1_hoa12"
os.makedirs(IMG_DIR, exist_ok=True)
IMG_NAME = "C1_Q069_1.png"
IMG_PATH = os.path.join(IMG_DIR, IMG_NAME)

ORIG_DA_DOCX = r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\2026 - Phan dang de thi thu truong, so. TTB1-20260711T201135Z-2-001\2026 - Phan dang de thi thu truong, so. TTB1\PD 2026 - Hoa 12. CD1 - Ester - lipid. TTB. ok - DA chi tiet.docx"

# 1. Trích xuất và lưu ảnh công thức
if os.path.exists(ORIG_DA_DOCX):
    doc_orig = docx.Document(ORIG_DA_DOCX)
    p209 = doc_orig.paragraphs[209]
    v_imgs = p209._element.xpath('.//*[local-name()="imagedata"]')
    if v_imgs:
        rId = v_imgs[0].get(qn('r:id'))
        part = doc_orig.part.related_parts[rId]
        im = Image.open(io.BytesIO(part.blob))
        im.save(IMG_PATH, format="PNG")
        im.save(os.path.join(IMG_DIR, "c1_q069_img1.png"), format="PNG")
        print(f"[1] Đã trích xuất và lưu ảnh nét: {IMG_PATH}")

DETAILED_SOLUTION = (
    "Công thức cấu tạo của ester X là CH3COOCH2CH(CH3)CH2CH3 (2-methylbutyl acetate).\n"
    "Phân tích từng nhận định:\n"
    "(a) ĐÚNG. Phân tử X gồm 7C, 14H và 2O, mạch hở chứa 1 nhóm ester (-COO-) và gốc hydrocarbon no "
    "nên X là ester no, đơn chức, có công thức phân tử là C7H14O2.\n"
    "(b) ĐÚNG. Ester X bị thuỷ phân thuận nghịch trong môi trường acid (H2SO4 loãng, t°) "
    "và thuỷ phân một chiều (xà phòng hoá) trong môi trường base (dung dịch kiềm NaOH, t°).\n"
    "(c) SAI. Do phân tử không có liên kết O-H linh động nên X không tạo được liên kết hydrogen mạnh với nước, "
    "khiến ester rất ít tan trong nước.\n"
    "(d) ĐÚNG. Trong phòng thí nghiệm, X được tổng hợp bằng phản ứng ester hoá giữa acetic acid (CH3COOH) "
    "và 2-methylbutan-1-ol (CH3CH2CH(CH3)CH2OH) với xúc tác H2SO4 đặc, đun nóng:\n"
    "    CH3COOH + CH3CH2CH(CH3)CH2OH <=> CH3COOCH2CH(CH3)CH2CH3 + H2O.\n"
    "➔ Số nhận định đúng là 3 gồm (a), (b), (d).\n"
    "➔ Chọn đáp án A."
)

DETAILED_SOL_LINES = [
    "- Công thức cấu tạo của ester X là CH3COOCH2CH(CH3)CH2CH3 (2-methylbutyl acetate).",
    "- Phân tích từng nhận định:",
    "  • (a) ĐÚNG. Phân tử X gồm 7C, 14H, 2O, mạch hở chứa 1 nhóm -COO- và gốc no nên là ester no, đơn chức (C7H14O2).",
    "  • (b) ĐÚNG. Ester bị thuỷ phân trong cả môi trường acid (thuận nghịch) và môi trường base (xà phòng hoá một chiều).",
    "  • (c) SAI. Giữa các phân tử ester không tạo liên kết hydrogen mạnh với nước (do không có nguyên tử H linh động).",
    "  • (d) ĐÚNG. X được tổng hợp từ CH3COOH và 2-methylbutan-1-ol: CH3COOH + CH3CH2CH(CH3)CH2OH <=> X + H2O.",
    "➔ Số nhận định đúng là 3 gồm (a), (b), (d). Chọn đáp án A."
]

# 2. Cập nhật JSON Banks
JSON_FILES = [
    r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Tron_Bo_9_Chuong.json",
    r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Chuong_1_2_3.json"
]

for jf in JSON_FILES:
    if os.path.exists(jf):
        with open(jf, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        updated = False
        c1_qs = data.get("C1", {}).get("questions", [])
        for q in c1_qs:
            if q.get("id") == "C1-P1-TH-069":
                q["solution"] = DETAILED_SOLUTION
                q["sol_lines"] = DETAILED_SOL_LINES
                q["image_paths"] = [IMG_PATH]
                updated = True
                break
        
        if updated:
            with open(jf, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"[2] Đã cập nhật JSON: {jf}")

# 3. Chạy lại phát sinh Đề 20
sys.path.insert(0, r"c:\Antigravity_Thanh\scripts")
from generate_5_de_gk1_hoa12 import BANK_PATH, TESTS_SPEC, generate_single_test_docx
from generate_azota_gk1_hoa12 import generate_azota_docx
from extract_q_assets import extract_chapter_assets
from select_de11_to_de20 import new_tests

print("[3] Đang tái tạo lại bộ file Đề 20 (Học sinh, Giáo viên, Azota)...")
with open(BANK_PATH, "r", encoding="utf-8") as f:
    bank_data = json.load(f)

bank_map = {}
for ch_code in ["C1", "C2", "C3"]:
    for q in bank_data[ch_code]["questions"]:
        bank_map[(ch_code, q["q_num"])] = q

assets_map = {c: extract_chapter_assets(c) for c in [1, 2, 3]}
# Đảm bảo câu 69 chương 1 luôn có ảnh
assets_map[1][69] = {"images": [IMG_PATH], "tables": []}

de20_cfg = [t for t in new_tests if t["de_num"] == 20][0]
generate_single_test_docx(de20_cfg, bank_map, assets_map, is_teacher=False)
generate_single_test_docx(de20_cfg, bank_map, assets_map, is_teacher=True)
generate_azota_docx(de20_cfg, bank_map, assets_map)
print("    --> Đã tái tạo Đề 20 hoàn tất!")

# 4. Cập nhật trực tiếp file Word Ngân hàng Chương 1
print("[4] Đang cập nhật ngân hàng Chương 1 (Ngan_Hang_Hoa_12_Chuong_1_Ester_Lipid)...")
c1_gv_files = [
    r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Chuong_1_Ester_Lipid_GiaoVien.docx",
    r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12_Chuong_1_Ester_Lipid_GiaoVien.docx"
]
c1_hs_files = [
    r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Chuong_1_Ester_Lipid_HocSinh.docx",
    r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12_Chuong_1_Ester_Lipid_HocSinh.docx"
]

def update_c1_word_file(filepath, is_gv=True):
    if not os.path.exists(filepath):
        return
    doc = docx.Document(filepath)
    found_q = False
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if "C1-P1-TH-069" in txt or ("Câu 69." in txt and "cảnh báo các con ong khác" in txt):
            found_q = True
            # Kiểm tra đoạn tiếp theo có ảnh hay không
            next_p = doc.paragraphs[i+1]
            # Nếu next_p không có run hoặc trống, ta chèn ảnh rõ ràng vào next_p
            has_drawing = len(next_p._element.xpath('.//*[local-name()="blip" or local-name()="imagedata"]')) > 0
            if not has_drawing or len(next_p.runs) == 0:
                p_img = doc.paragraphs[i+1]
                p_img.text = ""
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(IMG_PATH, width=Cm(6.0))
            
            # Nếu là bản Giáo viên, sửa lại phần Lời giải chi tiết
            if is_gv:
                # Tìm dòng Hướng dẫn giải
                for j in range(i+2, min(len(doc.paragraphs), i+15)):
                    pj = doc.paragraphs[j]
                    if pj.text.strip().startswith("Hướng dẫn giải") or "Ester không tạo liên kết hydrogen" in pj.text:
                        # Thay thế toàn bộ đoạn giải thích
                        pj.text = ""
                        r_hd = pj.add_run("Hướng dẫn giải: ")
                        r_hd.bold = True
                        r_hd.font.name = "Times New Roman"
                        r_hd.font.size = Pt(11)
                        r_hd.font.color.rgb = RGBColor(21, 101, 192) # Blue
                        
                        r_ans = pj.add_run("Đáp án A\n")
                        r_ans.bold = True
                        r_ans.font.name = "Times New Roman"
                        r_ans.font.size = Pt(11)
                        r_ans.font.color.rgb = RGBColor(198, 40, 40) # Red
                        
                        # Thêm các đoạn phân tích
                        r_body = pj.add_run(DETAILED_SOLUTION)
                        r_body.font.name = "Times New Roman"
                        r_body.font.size = Pt(10.5)
                        pj.paragraph_format.left_indent = Cm(0.6)
                        break
            break
    if found_q:
        doc.save(filepath)
        print(f"    --> Đã cập nhật xong file: {os.path.basename(filepath)}")

for f in c1_gv_files:
    update_c1_word_file(f, is_gv=True)
for f in c1_hs_files:
    update_c1_word_file(f, is_gv=False)

print("\nTẤT CẢ CÁC BƯỚC ĐÃ HOÀN TẤT THÀNH CÔNG!")
