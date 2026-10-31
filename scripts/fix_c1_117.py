# -*- coding: utf-8 -*-
"""
Script sửa chữa triệt để cho câu hỏi C1-P2-TH-117:
1. Đảm bảo ảnh phương trình điều chế benzyl hexanoate (C1_Q117_1.png) sẵn sàng trong images_gk1_hoa12.
2. Nâng cấp xử lý chèn ảnh trong PHẦN II của generate_5_de_gk1_hoa12.py và generate_azota_gk1_hoa12.py:
   Nếu stem_lines[0] có các từ khóa như 'phương trình hoá học sau:', 'như sau:',... thì chèn ảnh phương trình ngay sau stem_lines[0], trước đoạn mô tả thực nghiệm tiếp theo.
3. Bổ sung Lời giải chi tiết hoàn chỉnh cho cả 4 ý a, b, c, d của C1-P2-TH-117.
4. Cập nhật cơ sở dữ liệu JSON master và tái tạo lại Đề 20 (cả 3 bản HocSinh, GiaoVien, Azota ở cả 2 thư mục).
"""

import os
import re
import io
import sys
import json
import docx
from PIL import Image

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

IMG_DIR = r"c:\Antigravity_Thanh\San_Pham\images_gk1_hoa12"
os.makedirs(IMG_DIR, exist_ok=True)
IMG_NAME = "C1_Q117_1.png"
IMG_PATH = os.path.join(IMG_DIR, IMG_NAME)

ORIG_DA_DOCX = r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\2026 - Phan dang de thi thu truong, so. TTB1-20260711T201135Z-2-001\2026 - Phan dang de thi thu truong, so. TTB1\PD 2026 - Hoa 12. CD1 - Ester - lipid. TTB. ok - DA chi tiet.docx"

# 1. Trích xuất ảnh phương trình nếu chưa có
if os.path.exists(ORIG_DA_DOCX):
    doc_orig = docx.Document(ORIG_DA_DOCX)
    p654 = doc_orig.paragraphs[654]
    v_imgs = p654._element.xpath('.//*[local-name()="imagedata"]')
    if v_imgs:
        rId = v_imgs[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        part = doc_orig.part.related_parts[rId]
        im = Image.open(io.BytesIO(part.blob))
        im.save(IMG_PATH, format="PNG")
        im.save(os.path.join(IMG_DIR, "c1_q117_img1.png"), format="PNG")
        print(f"[1] Đã trích xuất và lưu ảnh phương trình: {IMG_PATH} (kích thước {im.size})")

DETAILED_SOLUTION_117 = (
    "a) Sai vì liên kết O-H của carboxylic acid có số sóng hấp thụ đặc trưng trong khoảng 3300 – 2500 cm^-1 "
    "(tín hiệu ở 1721 cm^-1 là đặc trưng cho liên kết C=O của nhóm carbonyl trong carboxylic acid).\n"
    "b) Sai vì:\n"
    "   - n(benzyl alcohol) = 16,2 / 108 = 0,15 mol; n(hexanoic acid) = 23,2 / 116 = 0,20 mol.\n"
    "   - Do n(alcohol) < n(acid) (tỉ lệ mol phản ứng 1 : 1) nên lượng ester lí thuyết tính theo benzyl alcohol.\n"
    "   - Khối lượng benzyl hexanoate lí thuyết (H = 100%): m(lí thuyết) = 0,15 × 206 = 30,9 gam.\n"
    "   - Hiệu suất phản ứng: h% = (18,54 / 30,9) × 100% = 60% => h = 60 (không phải 75).\n"
    "c) Sai vì trong phản ứng ester hoá, benzyl alcohol đóng vai trò là tác nhân nucleophile "
    "tấn công vào carbon mang điện tích dương của nhóm carbonyl, không đóng vai trò acid theo thuyết Brønsted–Lowry.\n"
    "d) Sai vì phản ứng ester hoá là phản ứng thế nhóm acyl, số oxi hoá của các nguyên tố đều không thay đổi "
    "nên hexanoic acid không đóng vai trò là chất oxi hoá."
)

IR_TABLE_117 = [
    ["Liên kết", "O-H (alcohol)", "O-H (carboxylic acid)", "C=O (ester, carboxylic acid)"],
    ["Số sóng (cm-1)", "3650 - 3200", "3300 - 2500", "1780 - 1650"]
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
            if q.get("id") == "C1-P2-TH-117" or q.get("q_num") == 117:
                q["solution"] = DETAILED_SOLUTION_117
                q["image_paths"] = [IMG_PATH]
                q["tables"] = [IR_TABLE_117]
                updated = True
                break
        
        if updated:
            with open(jf, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"[2] Đã cập nhật JSON: {jf}")

# 3. Cập nhật generate_5_de_gk1_hoa12.py
GEN_FILE = r"c:\Antigravity_Thanh\scripts\generate_5_de_gk1_hoa12.py"
if os.path.exists(GEN_FILE):
    txt = open(GEN_FILE, "r", encoding="utf-8").read()
    
    # 3.1. Cập nhật SOLUTION_PATCHES cho ("C1", 117)
    if '("C1", 117)' not in txt:
        sol_entry = f'    ("C1", 117): """{DETAILED_SOLUTION_117}""",\n'
        txt = txt.replace("SOLUTION_PATCHES = {\n", "SOLUTION_PATCHES = {\n" + sol_entry)
        print("[3.1] Đã thêm SOLUTION_PATCHES cho ('C1', 117)")
    
    # 3.2. Cập nhật xử lý hình ảnh sớm trong PHẦN II
    old_p2_block = """        stem_lines = q["stem_lines"]
        if stem_lines:
            add_chem_runs(p_q, stem_lines[0], size=11.5)
            for extra_ln in stem_lines[1:]:
                p_ex = doc.add_paragraph()
                p_ex.paragraph_format.space_before = Pt(1)
                p_ex.paragraph_format.space_after = Pt(1)
                add_chem_runs(p_ex, extra_ln, size=11.5)

        c_int = int(ch_code[1])
        q_assets = assets_map.get(c_int, {}).get(orig_qnum, {"images": [], "tables": []})
        for img_path in q_assets["images"]:
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_path, width=Cm(8.5))
        for tbl_data in q_assets["tables"]:
            add_data_table_to_doc(doc, tbl_data)"""

    new_p2_block = """        stem_lines = q["stem_lines"]
        c_int = int(ch_code[1])
        q_assets = assets_map.get(c_int, {}).get(orig_qnum, {"images": [], "tables": []})
        img_placed_early = False
        img_triggers = ("như sau:", "dưới đây:", "hình vẽ sau:", "hình sau:", "sơ đồ sau:", "phương trình hoá học sau:", "phương trình sau:")

        if stem_lines:
            add_chem_runs(p_q, stem_lines[0], size=11.5)
            if any(trig in stem_lines[0].lower() for trig in img_triggers) and q_assets["images"]:
                for img_path in q_assets["images"]:
                    if os.path.exists(img_path):
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.paragraph_format.space_before = Pt(2)
                        p_img.paragraph_format.space_after = Pt(2)
                        p_img.add_run().add_picture(img_path, width=Cm(12.5) if "Q117" in img_path else Cm(8.5))
                img_placed_early = True
            for extra_ln in stem_lines[1:]:
                p_ex = doc.add_paragraph()
                p_ex.paragraph_format.space_before = Pt(1)
                p_ex.paragraph_format.space_after = Pt(1)
                add_chem_runs(p_ex, extra_ln, size=11.5)

        if not img_placed_early:
            for img_path in q_assets["images"]:
                if os.path.exists(img_path):
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.paragraph_format.space_before = Pt(2)
                    p_img.paragraph_format.space_after = Pt(2)
                    p_img.add_run().add_picture(img_path, width=Cm(8.5))
        for tbl_data in q_assets["tables"]:
            add_data_table_to_doc(doc, tbl_data)"""

    if old_p2_block in txt:
        txt = txt.replace(old_p2_block, new_p2_block)
        print("[3.2] Đã nâng cấp xử lý chèn ảnh sớm cho PHẦN II trong generate_5_de_gk1_hoa12.py")
    
    with open(GEN_FILE, "w", encoding="utf-8") as f:
        f.write(txt)

# 4. Cập nhật generate_azota_gk1_hoa12.py
AZOTA_FILE = r"c:\Antigravity_Thanh\scripts\generate_azota_gk1_hoa12.py"
if os.path.exists(AZOTA_FILE):
    txt_az = open(AZOTA_FILE, "r", encoding="utf-8").read()
    
    old_az_p2 = """        stem_lines = q["stem_lines"]
        if stem_lines:
            add_chem_runs(p_q, stem_lines[0], size=12)
            for extra_ln in stem_lines[1:]:
                p_ex = doc.add_paragraph()
                p_ex.paragraph_format.space_before = Pt(1)
                p_ex.paragraph_format.space_after = Pt(2)
                add_chem_runs(p_ex, extra_ln, size=12)

        c_int = int(ch_code[1])
        q_assets = assets_map.get(c_int, {}).get(orig_qnum, {"images": [], "tables": []})
        for img_path in q_assets["images"]:
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_path, width=Cm(8.5))
        for tbl_data in q_assets["tables"]:
            add_data_table_to_doc(doc, tbl_data)"""

    new_az_p2 = """        stem_lines = q["stem_lines"]
        c_int = int(ch_code[1])
        q_assets = assets_map.get(c_int, {}).get(orig_qnum, {"images": [], "tables": []})
        img_placed_early = False
        img_triggers = ("như sau:", "dưới đây:", "hình vẽ sau:", "hình sau:", "sơ đồ sau:", "phương trình hoá học sau:", "phương trình sau:")

        if stem_lines:
            add_chem_runs(p_q, stem_lines[0], size=12)
            if any(trig in stem_lines[0].lower() for trig in img_triggers) and q_assets["images"]:
                for img_path in q_assets["images"]:
                    if os.path.exists(img_path):
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.paragraph_format.space_before = Pt(2)
                        p_img.paragraph_format.space_after = Pt(2)
                        p_img.add_run().add_picture(img_path, width=Cm(12.5) if "Q117" in img_path else Cm(8.5))
                img_placed_early = True
            for extra_ln in stem_lines[1:]:
                p_ex = doc.add_paragraph()
                p_ex.paragraph_format.space_before = Pt(1)
                p_ex.paragraph_format.space_after = Pt(2)
                add_chem_runs(p_ex, extra_ln, size=12)

        if not img_placed_early:
            for img_path in q_assets["images"]:
                if os.path.exists(img_path):
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.paragraph_format.space_before = Pt(2)
                    p_img.paragraph_format.space_after = Pt(2)
                    p_img.add_run().add_picture(img_path, width=Cm(8.5))
        for tbl_data in q_assets["tables"]:
            add_data_table_to_doc(doc, tbl_data)"""

    if old_az_p2 in txt_az:
        txt_az = txt_az.replace(old_az_p2, new_az_p2)
        print("[4] Đã nâng cấp xử lý chèn ảnh sớm cho PHẦN II trong generate_azota_gk1_hoa12.py")
    
    with open(AZOTA_FILE, "w", encoding="utf-8") as f:
        f.write(txt_az)

# 5. Tái tạo lại Đề 20 (Đề GK1 có Câu 1 - Phần II là C1-P2-TH-117)
print("[5] Đang tái tạo lại trọn bộ 3 bản của Đề 20...")
sys.path.insert(0, r"c:\Antigravity_Thanh\scripts")
from generate_5_de_gk1_hoa12 import BANK_PATH, generate_single_test_docx
from generate_azota_gk1_hoa12 import generate_azota_docx
from extract_q_assets import extract_chapter_assets
from select_de11_to_de20 import new_tests

with open(BANK_PATH, "r", encoding="utf-8") as f:
    bank_data = json.load(f)

bank_map = {}
for ch_code in ["C1", "C2", "C3"]:
    for q in bank_data[ch_code]["questions"]:
        bank_map[(ch_code, q["q_num"])] = q

assets_map = {c: extract_chapter_assets(c) for c in [1, 2, 3]}
# Đảm bảo assets của câu 117 có đủ ảnh và bảng
assets_map[1][117] = {"images": [IMG_PATH], "tables": [IR_TABLE_117]}

de20_cfg = [t for t in new_tests if t["de_num"] == 20][0]
generate_single_test_docx(de20_cfg, bank_map, assets_map, is_teacher=False)
generate_single_test_docx(de20_cfg, bank_map, assets_map, is_teacher=True)
generate_azota_docx(de20_cfg, bank_map, assets_map)
print("    --> Đã hoàn tất tái tạo Đề 20!")

print("\nHOÀN TẤT BỔ SUNG PHƯƠNG TRÌNH ĐIỀU CHẾ VÀ LỜI GIẢI CHI TIẾT CHO C1-P2-TH-117!")
