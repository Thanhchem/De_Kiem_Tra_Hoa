# -*- coding: utf-8 -*-
"""
Script sửa chữa câu C2-P1-NB-220 theo đúng ảnh người dùng cung cấp:
- Stem: Tinh bột là sản phẩm quang hóa của cây xanh và được dự trữ trong các loại hạt, củ, quả,... Loại nông sản nào sau đây có chứa hàm lượng tinh bột cao nhất?
- Opts: A. Rau răm.          B. Hạt gạo.          C. Hạt óc chó.          D. Quả nho chín.
- ĐA: B
- Hướng dẫn giải: Chọn đáp án B (Hạt gạo).

Cập nhật vào:
1. JSON Banks: Ngan_Hang_Hoa_12_Tron_Bo_9_Chuong.json & Ngan_Hang_Hoa_12_Chuong_1_2_3.json
2. Scripts: generate_5_de_gk1_hoa12.py & generate_azota_gk1_hoa12.py (STEM_PATCHES, OPT_PATCHES, SOLUTION_PATCHES)
3. Đề 20 (Đề GK1 có Câu 13 là C2-P1-NB-220): cả 3 bản HocSinh, GiaoVien, Azota ở San_Pham và Ngan_Hang_Hoa_12
4. File Word Ngân hàng Chương 2: Ngan_Hang_Hoa_12_Chuong_2_Carbohydrate_GiaoVien.docx & HocSinh.docx
"""

import os
import re
import sys
import json
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

NEW_STEM = "Tinh bột là sản phẩm quang hóa của cây xanh và được dự trữ trong các loại hạt, củ, quả,... Loại nông sản nào sau đây có chứa hàm lượng tinh bột cao nhất?"
NEW_OPTS_STR = "A. Rau răm.\tB. Hạt gạo.\tC. Hạt óc chó.\tD. Quả nho chín."
NEW_SOLUTION = "Chọn đáp án B (Hạt gạo)."

# 1. Cập nhật JSON Banks
JSON_FILES = [
    r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Tron_Bo_9_Chuong.json",
    r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Chuong_1_2_3.json"
]

for jf in JSON_FILES:
    if os.path.exists(jf):
        with open(jf, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        updated = False
        c2_qs = data.get("C2", {}).get("questions", [])
        for q in c2_qs:
            if q.get("id") == "C2-P1-NB-220" or q.get("q_num") == 220:
                q["stem_lines"] = [NEW_STEM]
                q["opt_lines"] = [NEW_OPTS_STR]
                q["answer"] = "B"
                q["solution"] = NEW_SOLUTION
                q["sol_lines"] = [NEW_SOLUTION]
                updated = True
                break
        
        if updated:
            with open(jf, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"[1] Đã cập nhật JSON: {jf}")

# 2. Cập nhật PATCHES trong scripts
# Trong generate_5_de_gk1_hoa12.py: thêm STEM_PATCHES[("C2", 220)] và OPT_PATCHES[("C2", 220)]
GEN_SCRIPT = r"c:\Antigravity_Thanh\scripts\generate_5_de_gk1_hoa12.py"
if os.path.exists(GEN_SCRIPT):
    txt = open(GEN_SCRIPT, "r", encoding="utf-8").read()
    
    # Kiểm tra STEM_PATCHES
    patch_entry = '    ("C2", 220): ["Tinh bột là sản phẩm quang hóa của cây xanh và được dự trữ trong các loại hạt, củ, quả,... Loại nông sản nào sau đây có chứa hàm lượng tinh bột cao nhất?"],\n'
    if '("C2", 220)' not in txt:
        target = "STEM_PATCHES = {\n"
        if target in txt:
            txt = txt.replace(target, target + patch_entry)
            print("[2] Đã bổ sung STEM_PATCHES cho ('C2', 220) trong generate_5_de_gk1_hoa12.py")
    
    # Kiểm tra OPT_PATCHES
    if "OPT_PATCHES" not in txt:
        opt_def = '\nOPT_PATCHES = {\n    ("C2", 220): ["A. Rau răm.\\tB. Hạt gạo.\\tC. Hạt óc chó.\\tD. Quả nho chín."]\n}\n'
        # Chèn trước STEM_PATCHES
        txt = txt.replace("STEM_PATCHES = {", opt_def + "\nSTEM_PATCHES = {")
        # Thêm xử lý OPT_PATCHES trong generate_single_test_docx
        old_chk = '        if (ch_code, orig_qnum) in STEM_PATCHES:\n            q["stem_lines"] = STEM_PATCHES[(ch_code, orig_qnum)]'
        new_chk = '        if (ch_code, orig_qnum) in STEM_PATCHES:\n            q["stem_lines"] = STEM_PATCHES[(ch_code, orig_qnum)]\n        if (ch_code, orig_qnum) in OPT_PATCHES:\n            q["opt_lines"] = OPT_PATCHES[(ch_code, orig_qnum)]'
        txt = txt.replace(old_chk, new_chk, 1)
        print("[2] Đã bổ sung OPT_PATCHES vào generate_5_de_gk1_hoa12.py")
    else:
        if '("C2", 220)' not in txt:
            txt = txt.replace("OPT_PATCHES = {\n", 'OPT_PATCHES = {\n    ("C2", 220): ["A. Rau răm.\\tB. Hạt gạo.\\tC. Hạt óc chó.\\tD. Quả nho chín."],\n')
            
    with open(GEN_SCRIPT, "w", encoding="utf-8") as f:
        f.write(txt)

# Trong generate_azota_gk1_hoa12.py
AZOTA_SCRIPT = r"c:\Antigravity_Thanh\scripts\generate_azota_gk1_hoa12.py"
if os.path.exists(AZOTA_SCRIPT):
    txt_az = open(AZOTA_SCRIPT, "r", encoding="utf-8").read()
    if '("C2", 220)' not in txt_az:
        target_az = "EXTRA_OPT_PATCHES = {\n"
        if target_az in txt_az:
            entry_az = '    ("C2", 220): ["A. Rau răm.\\tB. Hạt gạo.\\tC. Hạt óc chó.\\tD. Quả nho chín."],\n'
            txt_az = txt_az.replace(target_az, target_az + entry_az)
            with open(AZOTA_SCRIPT, "w", encoding="utf-8") as f:
                f.write(txt_az)
            print("[2] Đã bổ sung EXTRA_OPT_PATCHES cho ('C2', 220) trong generate_azota_gk1_hoa12.py")

# 3. Cập nhật trực tiếp các file Word Đề 20
print("[3] Cập nhật các file Word Đề 20...")
de20_files = [
    (r"c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_GK1_Hoa_12_De20_GiaoVien.docx", "GV"),
    (r"c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_GK1_Hoa_12_De20_HocSinh.docx", "HS"),
    (r"c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_GK1_Hoa_12_De20_Azota.docx", "AZ"),
    (r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\De_Kiem_Tra_GK1_Hoa_12_De20_GiaoVien.docx", "GV"),
    (r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\De_Kiem_Tra_GK1_Hoa_12_De20_HocSinh.docx", "HS"),
    (r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\De_Kiem_Tra_GK1_Hoa_12_De20_Azota.docx", "AZ"),
]

def update_de20_docx(fpath, ftype):
    if not os.path.exists(fpath):
        return
    doc = docx.Document(fpath)
    for i, p in enumerate(doc.paragraphs):
        txt = p.text
        if ("Câu 13." in txt and "quang hóa của cây xanh" in txt) or "C2-P1-NB-220" in txt:
            # Sửa câu hỏi
            # Giữ nguyên tag [Mã NH: ...] nếu có
            if "[Mã NH:" in txt:
                tag_part = txt[txt.find("[Mã NH:"):txt.find("]")+1]
                p.text = f"Câu 13. {tag_part} {NEW_STEM}"
                # Format màu
                runs = p.runs
                if runs:
                    runs[0].font.name = "Times New Roman"
                    runs[0].font.size = Pt(11.5)
            elif "Câu 13." in txt:
                p.text = f"Câu 13. {NEW_STEM}"
                if p.runs:
                    p.runs[0].font.name = "Times New Roman"
                    p.runs[0].font.size = Pt(11.5 if ftype != "AZ" else 12)
            
            # Đoạn tiếp theo là phương án A, B, C, D
            p_opt = doc.paragraphs[i+1]
            p_opt.text = ""
            options_data = [
                ("A. ", "Rau răm.\t", False),
                ("B. ", "Hạt gạo.\t", ftype == "GV"),
                ("C. ", "Hạt óc chó.\t", False),
                ("D. ", "Quả nho chín.", False),
            ]
            for lbl, opt_text, is_ans in options_data:
                r_lbl = p_opt.add_run(lbl)
                r_lbl.bold = True
                r_lbl.font.name = "Times New Roman"
                r_lbl.font.size = Pt(11.5 if ftype != "AZ" else 12)
                if is_ans and ftype == "GV":
                    r_lbl.font.color.rgb = RGBColor(198, 40, 40)
                    r_lbl.underline = True
                else:
                    r_lbl.font.color.rgb = RGBColor(21, 101, 192) if ftype != "AZ" else RGBColor(0, 102, 204)
                
                r_txt = p_opt.add_run(opt_text)
                r_txt.font.name = "Times New Roman"
                r_txt.font.size = Pt(11.5 if ftype != "AZ" else 12)
                if is_ans and ftype == "GV":
                    r_txt.bold = True
                    r_txt.font.color.rgb = RGBColor(198, 40, 40)
                    r_txt.underline = True
            
            # Nếu là bản Giáo viên hoặc Azota: có dòng Hướng dẫn giải
            if ftype in ("GV", "AZ"):
                p_sol = doc.paragraphs[i+2]
                if "Hướng dẫn giải" in p_sol.text or "Lời giải" in p_sol.text:
                    p_sol.text = ""
                    p_sol.paragraph_format.left_indent = Cm(0.6)
                    r_h = p_sol.add_run("Hướng dẫn giải: " if ftype == "GV" else "Lời giải: ")
                    r_h.bold = True
                    r_h.font.name = "Times New Roman"
                    r_h.font.size = Pt(11 if ftype == "GV" else 12)
                    r_h.font.color.rgb = RGBColor(21, 101, 192)
                    
                    r_b = p_sol.add_run(NEW_SOLUTION)
                    r_b.font.name = "Times New Roman"
                    r_b.font.size = Pt(11 if ftype == "GV" else 12)
            break
    doc.save(fpath)
    print(f"    --> Đã cập nhật xong Đề 20: {os.path.basename(fpath)}")

for fp, ft in de20_files:
    update_de20_docx(fp, ft)

# 4. Cập nhật trực tiếp file Word Ngân hàng Chương 2
print("[4] Cập nhật file Word Ngân hàng Chương 2...")
c2_word_files = [
    (r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Chuong_2_Carbohydrate_GiaoVien.docx", True),
    (r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Chuong_2_Carbohydrate_HocSinh.docx", False),
    (r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12_Chuong_2_Carbohydrate_GiaoVien.docx", True),
    (r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12_Chuong_2_Carbohydrate_HocSinh.docx", False)
]

def update_c2_bank_docx(fpath, is_gv):
    if not os.path.exists(fpath):
        return
    doc = docx.Document(fpath)
    for i, p in enumerate(doc.paragraphs):
        txt = p.text
        if "C2-P1-NB-220" in txt or ("Câu 220." in txt and "quang hóa của cây xanh" in txt):
            # Cập nhật stem
            tag_part = txt[txt.find("[C2-P1-NB-220"):txt.find("]")+1] if "[C2-P1-NB-220" in txt else ""
            p.text = f"Câu 220. {tag_part} {NEW_STEM}".strip()
            if p.runs:
                p.runs[0].font.name = "Times New Roman"
                p.runs[0].font.size = Pt(11)
            
            # Cập nhật options
            p_opt = doc.paragraphs[i+1]
            p_opt.text = ""
            options_data = [
                ("A. ", "Rau răm.\t", False),
                ("B. ", "Hạt gạo.\t", is_gv),
                ("C. ", "Hạt óc chó.\t", False),
                ("D. ", "Quả nho chín.", False),
            ]
            for lbl, opt_text, is_ans in options_data:
                r_lbl = p_opt.add_run(lbl)
                r_lbl.bold = True
                r_lbl.font.name = "Times New Roman"
                r_lbl.font.size = Pt(11)
                if is_ans and is_gv:
                    r_lbl.font.color.rgb = RGBColor(198, 40, 40)
                    r_lbl.underline = True
                else:
                    r_lbl.font.color.rgb = RGBColor(21, 101, 192)
                
                r_txt = p_opt.add_run(opt_text)
                r_txt.font.name = "Times New Roman"
                r_txt.font.size = Pt(11)
                if is_ans and is_gv:
                    r_txt.bold = True
                    r_txt.font.color.rgb = RGBColor(198, 40, 40)
                    r_txt.underline = True
            
            # Lời giải cho GiaoVien
            if is_gv:
                p_sol = doc.paragraphs[i+2]
                if "Hướng dẫn giải" in p_sol.text:
                    p_sol.text = ""
                    p_sol.paragraph_format.left_indent = Cm(0.6)
                    r_h = p_sol.add_run("Hướng dẫn giải: ")
                    r_h.bold = True
                    r_h.font.name = "Times New Roman"
                    r_h.font.size = Pt(11)
                    r_h.font.color.rgb = RGBColor(21, 101, 192)
                    
                    r_b = p_sol.add_run(NEW_SOLUTION)
                    r_b.font.name = "Times New Roman"
                    r_b.font.size = Pt(11)
            break
    doc.save(fpath)
    print(f"    --> Đã cập nhật xong Ngân hàng C2: {os.path.basename(fpath)}")

for fp, ig in c2_word_files:
    update_c2_bank_docx(fp, ig)

print("\nHOÀN TẤT CẬP NHẬT CÂU C2-P1-NB-220 THEO ĐÚNG ẢNH CỦA THẦY!")
