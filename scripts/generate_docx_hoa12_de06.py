# -*- coding: utf-8 -*-
"""
Script tạo file Word (.docx) chuẩn "FORM ĐỀ KIỂM TRA" cho:
ĐỀ KIỂM TRA SỐ 6 - CHƯƠNG 2: CARBOHYDRATE (HÓA HỌC 12)
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De06.docx (Bản Học Sinh)
2. De_Kiem_Tra_Hoa_12_De06_GiaoVien.docx (Bản Giáo Viên có lời giải chi tiết)
Tuân thủ 100% quy cách: Header & Footer, Tab Stops (NO TABLE), Sub/Superscript công thức hóa học.
"""

import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
IMG_DIR = r"c:\Antigravity_Thanh\San_Pham\images_crop_de06"

COLOR_HF_RED = RGBColor(211, 47, 47)      # Coral Red (#D32F2F)
COLOR_HF_GREEN_HEX = "8BC390"             # Sage Green (#8BC390)
COLOR_PRIMARY = RGBColor(46, 125, 50)     # Dark Green (#2E7D32)
COLOR_CORRECT = RGBColor(198, 40, 40)     # Red for correct answers (#C62828)
COLOR_SOLUTION = RGBColor(21, 101, 192)   # Blue for solutions (#1565C0)
COLOR_GRAY = RGBColor(97, 97, 97)

def add_page_number_field(p, font_name="Times New Roman", font_size=10, bold=True, color=COLOR_HF_RED):
    run_txt = p.add_run("Trang ")
    set_run_style(run_txt, font_name=font_name, font_size=font_size, bold=bold, color=color)
    
    r = p.add_run()
    set_run_style(r, font_name=font_name, font_size=font_size, bold=bold, color=color)
    fld1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instr1 = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
    fld2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fld3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    r._r.append(fld1); r._r.append(instr1); r._r.append(fld2); r._r.append(fld3)

def set_run_style(run, font_name="Times New Roman", font_size=12, bold=False, italic=False, subscript=False, superscript=False, color=None):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold
    run.italic = italic
    if subscript:
        run.font.subscript = True
    if superscript:
        run.font.superscript = True
    if color:
        run.font.color.rgb = color

def add_chem_run(p, text, base_font="Times New Roman", base_size=12, bold=False, italic=False, color=None):
    tokens = re.split(r'(\^\{[^}]+\}|\^[0-9\+\-]+|~\{[^}]+\}|~[0-9]+)', text)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith('^'):
            content = tok[2:-1] if tok.startswith('^{') else tok[1:]
            r = p.add_run(content)
            set_run_style(r, font_name=base_font, font_size=base_size, bold=bold, italic=italic, superscript=True, color=color)
        elif tok.startswith('~'):
            content = tok[2:-1] if tok.startswith('~{') else tok[1:]
            r = p.add_run(content)
            set_run_style(r, font_name=base_font, font_size=base_size, bold=bold, italic=italic, subscript=True, color=color)
        else:
            r = p.add_run(tok)
            set_run_style(r, font_name=base_font, font_size=base_size, bold=bold, italic=italic, color=color)

def setup_page(doc):
    for sec in doc.sections:
        sec.top_margin = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
        sec.left_margin = Cm(1.5)
        sec.right_margin = Cm(1.5)
        sec.header_distance = Cm(0.6)
        sec.footer_distance = Cm(0.6)
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)

def set_cell_border_bottom(cell, color="8BC390", sz="6"):
    tcPr = cell._tc.get_or_add_tcPr()
    xml = f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/></w:tcBorders>'
    tcPr.append(parse_xml(xml))

def set_cell_border_top(cell, color="8BC390", sz="6"):
    tcPr = cell._tc.get_or_add_tcPr()
    xml = f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/></w:tcBorders>'
    tcPr.append(parse_xml(xml))

def set_cell_margins_zero(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    xml = f'<w:tcMar {nsdecls("w")}><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>'
    tcPr.append(parse_xml(xml))

def setup_header_footer(doc):
    sec = doc.sections[0]
    
    # 1. HEADER
    hdr = sec.header
    for p in list(hdr.paragraphs):
        p._p.getparent().remove(p._p)
    tbl_h = hdr.add_table(rows=1, cols=2, width=Cm(18.0))
    tbl_h.alignment = WD_TAB_ALIGNMENT.CENTER
    c_left = tbl_h.cell(0, 0)
    c_right = tbl_h.cell(0, 1)
    c_left.width = Cm(9.0)
    c_right.width = Cm(9.0)
    set_cell_margins_zero(c_left)
    set_cell_margins_zero(c_right)
    set_cell_border_bottom(c_left, color=COLOR_HF_GREEN_HEX, sz="6")
    set_cell_border_bottom(c_right, color=COLOR_HF_GREEN_HEX, sz="6")
    
    p_hl = c_left.paragraphs[0]
    p_hl.paragraph_format.space_before = Pt(0)
    p_hl.paragraph_format.space_after = Pt(2)
    p_hl.paragraph_format.line_spacing = 1.0
    r_hl = p_hl.add_run("Tài liệu lưu hành nội bộ")
    set_run_style(r_hl, font_name="Times New Roman", font_size=10, italic=True, color=COLOR_HF_RED)
    
    p_hr = c_right.paragraphs[0]
    p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_hr.paragraph_format.space_before = Pt(0)
    p_hr.paragraph_format.space_after = Pt(2)
    p_hr.paragraph_format.line_spacing = 1.0
    r_hr1 = p_hr.add_run("Thầy ")
    set_run_style(r_hr1, font_name="Times New Roman", font_size=10, color=COLOR_HF_RED)
    r_hr2 = p_hr.add_run("TRẦN VĂN THẠNH")
    set_run_style(r_hr2, font_name="Times New Roman", font_size=10, bold=True, color=COLOR_HF_RED)
    r_hr3 = p_hr.add_run(" - 0777.470.803")
    set_run_style(r_hr3, font_name="Times New Roman", font_size=10, color=COLOR_HF_RED)
    
    # 2. FOOTER
    ftr = sec.footer
    for p in list(ftr.paragraphs):
        p._p.getparent().remove(p._p)
    tbl_f = ftr.add_table(rows=1, cols=2, width=Cm(18.0))
    tbl_f.alignment = WD_TAB_ALIGNMENT.CENTER
    fc_left = tbl_f.cell(0, 0)
    fc_right = tbl_f.cell(0, 1)
    fc_left.width = Cm(15.0)
    fc_right.width = Cm(3.0)
    set_cell_margins_zero(fc_left)
    set_cell_margins_zero(fc_right)
    set_cell_border_top(fc_left, color=COLOR_HF_GREEN_HEX, sz="6")
    set_cell_border_top(fc_right, color=COLOR_HF_GREEN_HEX, sz="6")
    
    p_fl = fc_left.paragraphs[0]
    p_fl.paragraph_format.space_before = Pt(2)
    p_fl.paragraph_format.space_after = Pt(0)
    p_fl.paragraph_format.line_spacing = 1.0
    r_fl = p_fl.add_run("CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế   |   CS2: 24 Đặng Thái Thân, TP Huế")
    set_run_style(r_fl, font_name="Times New Roman", font_size=9.5, italic=True, color=COLOR_HF_RED)
    
    p_fr = fc_right.paragraphs[0]
    p_fr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_fr.paragraph_format.space_before = Pt(2)
    p_fr.paragraph_format.space_after = Pt(0)
    p_fr.paragraph_format.line_spacing = 1.0
    add_page_number_field(p_fr, font_name="Times New Roman", font_size=10, bold=True, color=COLOR_HF_RED)

def add_header_block(doc, is_teacher=False):
    def add_tab_line(p_left_text, p_right_text, bold_left=False, bold_right=False, italic_left=False, italic_right=False, sz_left=11, sz_right=11, color_left=None, color_right=None, space_after=2):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.tab_stops.add_tab_stop(Cm(18.0), WD_TAB_ALIGNMENT.RIGHT)
        
        r1 = p.add_run(p_left_text)
        set_run_style(r1, font_name="Times New Roman", font_size=sz_left, bold=bold_left, italic=italic_left, color=color_left)
        p.add_run("\t")
        r2 = p.add_run(p_right_text)
        set_run_style(r2, font_name="Times New Roman", font_size=sz_right, bold=bold_right, italic=italic_right, color=color_right)
        return p

    add_tab_line("SỞ GIÁO DỤC VÀ ĐÀO TẠO HUẾ", "ĐỀ KIỂM TRA CHUYÊN ĐỀ NĂM HỌC 2026 - 2027", bold_left=False, bold_right=True, sz_left=11, sz_right=11)
    add_tab_line("LỚP HOÁ THẦY TRẦN VĂN THẠNH", "MÔN: HOÁ HỌC - KHỐI 12", bold_left=True, bold_right=True, sz_left=11, sz_right=11, color_left=COLOR_PRIMARY)
    add_tab_line("ĐỀ CHÍNH THỨC", "CHỦ ĐỀ: CARBOHYDRATE (ĐỀ SỐ 06)", bold_left=True, bold_right=True, sz_left=11, sz_right=11)
    
    if is_teacher:
        p_note = doc.add_paragraph()
        p_note.paragraph_format.space_before = Pt(0)
        p_note.paragraph_format.space_after = Pt(2)
        p_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_nt = p_note.add_run("(BẢN GIÁO VIÊN — CÓ LỜI GIẢI CHI TIẾT)")
        set_run_style(r_nt, font_name="Times New Roman", font_size=11, bold=True, italic=True, color=COLOR_CORRECT)
        
        add_tab_line("(Đề thi có 04 trang)", "Thời gian làm bài: 50 phút (không kể thời gian phát đề)", italic_left=True, italic_right=True, sz_left=10.5, sz_right=10.5)
    else:
        add_tab_line("(Đề thi có 04 trang)", "Thời gian làm bài: 50 phút (không kể thời gian phát đề)", italic_left=True, italic_right=True, sz_left=10.5, sz_right=10.5)

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(3)
    p_info.paragraph_format.space_after = Pt(2)
    p_info.paragraph_format.line_spacing = 1.15
    r_info = p_info.add_run("Họ, tên thí sinh: ........................................................................ Lớp: ................. SBD: ................. Mã đề: 106")
    set_run_style(r_info, font_name="Times New Roman", font_size=11, italic=True)

    p_mass = doc.add_paragraph()
    p_mass.paragraph_format.space_before = Pt(1)
    p_mass.paragraph_format.space_after = Pt(4)
    p_mass.paragraph_format.line_spacing = 1.15
    r_m = p_mass.add_run("(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Ag = 108; Ba = 137)")
    set_run_style(r_m, font_name="Times New Roman", font_size=10, italic=True, color=COLOR_GRAY)

    p_sep = doc.add_paragraph()
    p_sep.paragraph_format.space_before = Pt(0)
    p_sep.paragraph_format.space_after = Pt(6)
    pPr = p_sep._element.get_or_add_pPr()
    xml = f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="8BC390"/></w:pBdr>'
    pPr.append(parse_xml(xml))

def add_section_heading(doc, title, subtitle=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(title)
    set_run_style(r, font_name="Times New Roman", font_size=12, bold=True, color=COLOR_PRIMARY)
    
    if subtitle:
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(4)
        p_sub.paragraph_format.line_spacing = 1.15
        r_sub = p_sub.add_run(subtitle)
        set_run_style(r_sub, font_name="Times New Roman", font_size=11, italic=True)

def add_solution_paragraph(doc, text_lines):
    p_head = doc.add_paragraph()
    p_head.paragraph_format.space_before = Pt(2)
    p_head.paragraph_format.space_after = Pt(1)
    p_head.paragraph_format.line_spacing = 1.15
    p_head.paragraph_format.left_indent = Cm(0.6)
    r_h = p_head.add_run("Lời giải chi tiết:")
    set_run_style(r_h, font_name="Times New Roman", font_size=11, bold=True, italic=True, color=COLOR_SOLUTION)

    for line in text_lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.left_indent = Cm(0.8)
        add_chem_run(p, line, base_font="Times New Roman", base_size=11, italic=False)

def add_question(doc, q_num, q_text, is_part2=False, is_part3=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    
    r_label = p.add_run(f"Câu {q_num}: ")
    set_run_style(r_label, font_name="Times New Roman", font_size=12, bold=True)
    add_chem_run(p, q_text, base_font="Times New Roman", base_size=12)
    return p

def add_options_1line(doc, a, b, c, d, correct=None, is_teacher=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.tab_stops.add_tab_stop(Cm(4.5), WD_TAB_ALIGNMENT.LEFT)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(13.5), WD_TAB_ALIGNMENT.LEFT)

    opts = [("A", a), ("B", b), ("C", c), ("D", d)]
    for i, (lbl, txt) in enumerate(opts):
        is_corr = is_teacher and (lbl == correct)
        col = COLOR_CORRECT if is_corr else None
        
        r_lbl = p.add_run(f"{lbl}. ")
        set_run_style(r_lbl, font_name="Times New Roman", font_size=12, bold=True, color=col)
        add_chem_run(p, txt, base_font="Times New Roman", base_size=12, bold=is_corr, color=col)
        
        if i < 3:
            p.add_run("\t")

def add_options_2lines(doc, a, b, c, d, correct=None, is_teacher=False):
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(1)
    p1.paragraph_format.space_after = Pt(1)
    p1.paragraph_format.line_spacing = 1.15
    p1.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)

    # Line 1: A and B
    for i, (lbl, txt) in enumerate([("A", a), ("B", b)]):
        is_corr = is_teacher and (lbl == correct)
        col = COLOR_CORRECT if is_corr else None
        r_lbl = p1.add_run(f"{lbl}. ")
        set_run_style(r_lbl, font_name="Times New Roman", font_size=12, bold=True, color=col)
        add_chem_run(p1, txt, base_font="Times New Roman", base_size=12, bold=is_corr, color=col)
        if i == 0:
            p1.add_run("\t")

    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(1)
    p2.paragraph_format.space_after = Pt(2)
    p2.paragraph_format.line_spacing = 1.15
    p2.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)

    # Line 2: C and D
    for i, (lbl, txt) in enumerate([("C", c), ("D", d)]):
        is_corr = is_teacher and (lbl == correct)
        col = COLOR_CORRECT if is_corr else None
        r_lbl = p2.add_run(f"{lbl}. ")
        set_run_style(r_lbl, font_name="Times New Roman", font_size=12, bold=True, color=col)
        add_chem_run(p2, txt, base_font="Times New Roman", base_size=12, bold=is_corr, color=col)
        if i == 0:
            p2.add_run("\t")

def add_options_4lines(doc, a, b, c, d, correct=None, is_teacher=False):
    opts = [("A", a), ("B", b), ("C", c), ("D", d)]
    for lbl, txt in opts:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.left_indent = Cm(0.5)

        is_corr = is_teacher and (lbl == correct)
        col = COLOR_CORRECT if is_corr else None
        r_lbl = p.add_run(f"{lbl}. ")
        set_run_style(r_lbl, font_name="Times New Roman", font_size=12, bold=True, color=col)
        add_chem_run(p, txt, base_font="Times New Roman", base_size=12, bold=is_corr, color=col)

def add_part2_items(doc, items, answers=None, is_teacher=False):
    for lbl, txt in items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.left_indent = Cm(0.5)

        r_lbl = p.add_run(f"{lbl}) ")
        set_run_style(r_lbl, font_name="Times New Roman", font_size=12, bold=True)
        add_chem_run(p, txt, base_font="Times New Roman", base_size=12)

        if is_teacher and answers and lbl in answers:
            ans_val = answers[lbl] # "Đúng" or "Sai"
            r_ans = p.add_run(f" ➔ [{ans_val.upper()}]")
            set_run_style(r_ans, font_name="Times New Roman", font_size=11, bold=True, color=COLOR_CORRECT)

def add_part3_answer_line(doc, ans_str=None, is_teacher=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.left_indent = Cm(0.5)
    
    if is_teacher:
        r_t = p.add_run(f"Đáp số: {ans_str}")
        set_run_style(r_t, font_name="Times New Roman", font_size=12, bold=True, color=COLOR_CORRECT)
    else:
        r_s = p.add_run("Đáp số: ......................................................................................................................................")
        set_run_style(r_s, font_name="Times New Roman", font_size=11, italic=True)

def add_image_centered(doc, img_name, width_cm=8.0, caption=None):
    img_path = os.path.join(IMG_DIR, img_name)
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(2)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(img_path, width=Cm(width_cm))
        if caption:
            p_cap = doc.add_paragraph()
            p_cap.paragraph_format.space_before = Pt(1)
            p_cap.paragraph_format.space_after = Pt(3)
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_cap = p_cap.add_run(caption)
            set_run_style(r_cap, font_name="Times New Roman", font_size=10.5, italic=True, bold=True, color=COLOR_PRIMARY)

def add_summary_table(doc):
    p_head = doc.add_paragraph()
    p_head.paragraph_format.space_before = Pt(8)
    p_head.paragraph_format.space_after = Pt(3)
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_h = p_head.add_run("BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ SỐ 6")
    set_run_style(r_h, font_name="Times New Roman", font_size=12, bold=True, color=COLOR_PRIMARY)

    # Phần I
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(2)
    p1.paragraph_format.space_after = Pt(2)
    r1 = p1.add_run("PHẦN I (Mỗi câu đúng 0,25 điểm):")
    set_run_style(r1, font_name="Times New Roman", font_size=11, bold=True)

    ans_p1 = [
        (1, "A"), (2, "A"), (3, "B"), (4, "B"), (5, "D"), (6, "C"),
        (7, "D"), (8, "A"), (9, "D"), (10, "D"), (11, "B"), (12, "A"),
        (13, "D"), (14, "A"), (15, "C"), (16, "B"), (17, "D"), (18, "C")
    ]
    p_p1_text = doc.add_paragraph()
    p_p1_text.paragraph_format.space_before = Pt(1)
    p_p1_text.paragraph_format.space_after = Pt(3)
    p_p1_text.paragraph_format.line_spacing = 1.15
    for i, (q, a) in enumerate(ans_p1):
        r_c = p_p1_text.add_run(f"{q}.{a}    ")
        set_run_style(r_c, font_name="Times New Roman", font_size=11, bold=True, color=COLOR_CORRECT)
        if (i + 1) % 6 == 0 and i < len(ans_p1) - 1:
            p_p1_text.add_run("\n")

    # Phần II
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(2)
    p2.paragraph_format.space_after = Pt(2)
    r2 = p2.add_run("PHẦN II (Mỗi ý đúng 0,25đ; 1 câu 1,0đ):")
    set_run_style(r2, font_name="Times New Roman", font_size=11, bold=True)

    ans_p2 = [
        "Câu 1: a-Đ, b-Đ, c-S, d-S",
        "Câu 2: a-Đ, b-S, c-S, d-Đ",
        "Câu 3: a-S, b-Đ, c-S, d-Đ",
        "Câu 4: a-Đ, b-S, c-S, d-Đ"
    ]
    for l in ans_p2:
        p_it = doc.add_paragraph()
        p_it.paragraph_format.space_before = Pt(1)
        p_it.paragraph_format.space_after = Pt(1)
        p_it.paragraph_format.left_indent = Cm(0.5)
        r_it = p_it.add_run(l)
        set_run_style(r_it, font_name="Times New Roman", font_size=11, bold=True, color=COLOR_CORRECT)

    # Phần III
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(2)
    p3.paragraph_format.space_after = Pt(2)
    r3 = p3.add_run("PHẦN III (Mỗi câu đúng 0,25 điểm):")
    set_run_style(r3, font_name="Times New Roman", font_size=11, bold=True)

    ans_p3 = [
        (1, "3"), (2, "5"), (3, "2"), (4, "397"), (5, "95"), (6, "12,5")
    ]
    p_p3_text = doc.add_paragraph()
    p_p3_text.paragraph_format.space_before = Pt(1)
    p_p3_text.paragraph_format.space_after = Pt(3)
    p_p3_text.paragraph_format.line_spacing = 1.15
    for q, a in ans_p3:
        r_c = p_p3_text.add_run(f"Câu {q}: {a}    |    ")
        set_run_style(r_c, font_name="Times New Roman", font_size=11, bold=True, color=COLOR_CORRECT)

def build_document(is_teacher=False):
    doc = docx.Document()
    setup_page(doc)
    setup_header_footer(doc)
    add_header_block(doc, is_teacher=is_teacher)

    # ==========================================
    # PHẦN I
    # ==========================================
    add_section_heading(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # Câu 1
    add_question(doc, 1, "Tinh bột và cellulose đều không thuộc loại")
    add_options_1line(doc, "monosaccharide.", "glucide.", "polysaccharide.", "carbohydrate.", correct="A", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Tinh bột và cellulose đều thuộc loại polysaccharide (carbohydrate phức tạp tạo bởi nhiều mắt xích glucose), không thuộc loại monosaccharide (đường đơn)."
        ])

    # Câu 2
    add_question(doc, 2, "Trong điều kiện thích hợp glucose lên men tạo thành khí CO~2~ và")
    add_options_1line(doc, "C~2~H~5~OH.", "CH~3~COOH.", "HCOOH.", "CH~3~CHO.", correct="A", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Phương trình lên men rượu: C~6~H~12~O~6~ -> 2C~2~H~5~OH + 2CO~2~ (xúc tác men rượu, 30 - 35°C)."
        ])

    # Câu 3
    add_question(doc, 3, "(SBT – KNTT) Saccharose thuộc loại carbohydrate nào sau đây?")
    add_options_1line(doc, "Monosaccharide.", "Disaccharide.", "Polysaccharide.", "Oligosaccharide.", correct="B", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Saccharose có công thức C~12~H~22~O~11~, được cấu tạo từ một gốc alpha-glucose và một gốc beta-fructose qua cầu nối oxygen, thuộc loại disaccharide."
        ])

    # Câu 4
    add_question(doc, 4, "(SBT – KNTT) Thuỷ phân một phân tử saccharose tạo thành")
    add_options_2lines(doc, 
        "hai phân tử glucose.", 
        "một phân tử glucose và một phân tử fructose.", 
        "hai phân tử fructose.", 
        "một phân tử galactose và một phân tử glucose.", 
        correct="B", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Phương trình thuỷ phân: C~12~H~22~O~11~ + H~2~O -> C~6~H~12~O~6~ (glucose) + C~6~H~12~O~6~ (fructose) (xúc tác acid hoặc enzyme sucrase)."
        ])

    # Câu 5
    add_question(doc, 5, "(SBT – KNTT) Phát biểu nào sau đây về cellulose không đúng?")
    add_options_2lines(doc,
        "Không tan trong nước.",
        "Là nguồn nguyên liệu sản xuất giấy.",
        "Có thể phản ứng với HNO~3~ tạo cellulose nitrate.",
        "Phản ứng màu với dung dịch iodine.",
        correct="D", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Chỉ có tinh bột mới có phản ứng màu xanh tím đặc trưng với dung dịch iodine; cellulose không có cấu trúc xoắn tạo lỗ rỗng hấp phụ iodine nên không có phản ứng này."
        ])

    # Câu 6
    add_question(doc, 6, "(SBT – CĐ) Tinh bột và cellulose là các polymer lần lượt tạo bởi các mắt xích")
    add_options_2lines(doc,
        "alpha-fructose và beta-glucose.",
        "beta-fructose và beta-glucose.",
        "alpha-glucose và beta-glucose.",
        "alpha-glucose và beta-fructose.",
        correct="C", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Tinh bột gồm các mắt xích alpha-glucose (amylose và amylopectin), còn cellulose cấu tạo từ các mắt xích beta-glucose liên kết beta-1,4-glycoside."
        ])

    # Câu 7
    add_question(doc, 7, "(SBT – CTST) Phát biểu nào sau đây không đúng về glucose và fructose?")
    add_options_2lines(doc,
        "Chúng đều có công thức phân tử C~6~H~12~O~6~.",
        "Chúng đều là các hợp chất carbohydrate.",
        "Chúng đều là các monosaccharide.",
        "Chúng có tính chất hoá học tương tự nhau.",
        correct="D", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Glucose là polyhydroxy aldehyde, còn fructose là polyhydroxy ketone. Do nhóm chức khác nhau nên chúng có tính chất hoá học khác biệt (ví dụ glucose làm mất màu nước bromine, fructose không phản ứng với nước bromine trong môi trường acid)."
        ])

    # Câu 8
    add_question(doc, 8, "(SBT – CĐ) Khi đun nóng dung dịch chứa carbohydrate X và Cu(OH)~2~ trong môi trường kiềm, X có phản ứng với Cu(OH)~2~ tạo kết tủa đỏ gạch. X không thể là")
    add_options_1line(doc, "saccharose.", "glucose.", "fructose.", "maltose.", correct="A", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Saccharose không có nhóm -OH hemiacetal hay hemiketal tự do (không mở vòng tạo nhóm aldehyde), do đó trong môi trường kiềm không khử được Cu(OH)~2~ thành Cu~2~O đỏ gạch."
        ])

    # Câu 9
    add_question(doc, 9, "(SBT – CĐ) Polymer là nguồn carbohydrate dự trữ có trong cơ thể động vật và được tạo thành từ các đơn vị glucose là")
    add_options_1line(doc, "cellulose.", "amylose.", "amylopectin.", "glycogen.", correct="D", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Glycogen (còn gọi là tinh bột động vật) là nguồn carbohydrate dự trữ năng lượng ở động vật và nấm, được cấu tạo từ các mắt xích alpha-glucose phân nhánh mạnh."
        ])

    # Câu 10
    add_question(doc, 10, "Trong công thức cấu tạo dạng mạch vòng alpha-glucose như sau:")
    add_image_centered(doc, "cau10_alpha_glucose.png", width_cm=4.5)
    p_q10 = doc.add_paragraph()
    p_q10.paragraph_format.space_before = Pt(1)
    p_q10.paragraph_format.space_after = Pt(2)
    p_q10.paragraph_format.line_spacing = 1.15
    r_q10 = p_q10.add_run("Nhóm -OH hemiacetal là -OH gắn ở carbon số mấy?")
    set_run_style(r_q10, font_name="Times New Roman", font_size=12)
    add_options_1line(doc, "2", "3", "4", "1", correct="D", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Trong dạng mạch vòng của glucose, carbon C1 liên kết trực tiếp với nguyên tử O vòng và nhóm -OH. Nhóm -OH ở vị trí C1 này được gọi là nhóm -OH hemiacetal (có khả năng mở vòng trở lại nhóm -CHO)."
        ])

    # Câu 11
    add_question(doc, 11, "Dung dịch saccharose tinh khiết không có tính khử, nhưng khi đun nóng với dung dịch H~2~SO~4~ lại có thể cho phản ứng tráng gương. Đó là do:")
    add_options_4lines(doc,
        "Đã có sự tạo thành aldehyde sau phản ứng.",
        "Saccharose bị thuỷ phân tạo thành glucose và fructose.",
        "Saccharose bị thuỷ phân tạo thành glucose.",
        "Saccharose bị thuỷ phân tạo thành fructose.",
        correct="B", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Khi đun nóng với dung dịch H~2~SO~4~ loãng (xúc tác acid), saccharose bị thuỷ phân thành glucose và fructose. Cả glucose và fructose sinh ra đều có khả năng tham gia phản ứng tráng gương."
        ])

    # Câu 12
    add_question(doc, 12, "Phát biểu nào sau đây sai?")
    add_options_2lines(doc,
        "Thủy phân saccharose chỉ thu được glucose.",
        "Glucose có khả năng tham gia phản ứng tráng bạc.",
        "Cellulose và tinh bột đều thuộc loại polysaccharide.",
        "Cellulose có cấu tạo mạch không phân nhánh.",
        correct="A", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Thuỷ phân saccharose thu được cả glucose và fructose với tỉ lệ mol 1 : 1, do đó phát biểu A sai."
        ])

    # Câu 13
    add_question(doc, 13, "(SBT – CTST) Mạch nha là tên gọi khác của carbohydrate nào sau đây?")
    add_options_1line(doc, "Glucose.", "Fructose.", "Saccharose.", "Maltose.", correct="D", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Maltose còn có tên gọi là đường mạch nha (malt sugar), có nhiều trong mầm lúa mạch, mầm thóc."
        ])

    # Câu 14
    add_question(doc, 14, "Polysaccharide X là chất rắn, màu trắng, dạng sợi. Trong bông nõn có gần 98% chất X. Thủy phân X, thu được monosaccharide Y. Phát biểu nào sau đây đúng?")
    add_options_2lines(doc,
        "Y có tính chất của alcohol đa chức.",
        "X có phản ứng tráng bạc.",
        "Phân tử khối của Y bằng 342.",
        "X dễ tan trong nước.",
        correct="A", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "X là cellulose, thuỷ phân cellulose thu được glucose (Y). Glucose có nhiều nhóm -OH liền kề nên mang đầy đủ tính chất của một alcohol đa chức (hoà tan Cu(OH)~2~ tạo phức xanh lam)."
        ])

    # Câu 15
    add_question(doc, 15, "(SBT – CTST) Chất nào sau đây không phải polymer sinh học?")
    add_options_1line(doc, "Tinh bột.", "Cellulose.", "Glucosamine.", "Chitin.", correct="C", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Glucosamine là monosaccharide dẫn xuất amin (phân tử đơn, M = 179), không phải là polymer. Tinh bột, cellulose và chitin đều là các polymer sinh học."
        ])

    # Câu 16
    add_question(doc, 16, "Cho các phát biểu sau về ứng dụng và trạng thái tự nhiên của glucose và fructose:\n(a) Cả glucose và fructose đều là chất rắn, dễ tan trong nước, có vị ngọt.\n(b) Glucose có nhiều trong quả chín (nhất là nho chín), fructose có nhiều trong mật ong.\n(c) Dung dịch truyền tĩnh mạch là glucose 5%.\n(d) Glucose và fructose chủ yếu đóng vai trò cung cấp năng lượng cho tế bào.\nSố phát biểu đúng là:")
    add_options_1line(doc, "3.", "4.", "1.", "2.", correct="B", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Cả 4 phát biểu (a), (b), (c), (d) đều hoàn toàn chính xác."
        ])

    # Câu 17
    add_question(doc, 17, "Cho các phát biểu sau:\n(a) Glucose được gọi là đường nho do có nhiều trong quả nho chín.\n(b) Chất béo là diester của glycerol với acid béo.\n(c) Phân tử amylopectin có cấu trúc mạch phân nhánh.\n(d) Ở nhiệt độ thường, triolein ở trạng thái rắn.\n(e) Trong mật ong chứa nhiều fructose.\n(f) Tinh bột là một trong những lương thực cơ bản của con người.\nSố phát biểu đúng là:")
    add_options_1line(doc, "5.", "6.", "3.", "4.", correct="D", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "(a) Đúng; (b) Sai (chất béo là triester); (c) Đúng; (d) Sai (triolein lỏng ở nhiệt độ thường); (e) Đúng (chiếm khoảng 40% trong mật ong); (f) Đúng.\n=> Có 4 phát biểu đúng: (a), (c), (e), (f)."
        ])

    # Câu 18
    add_question(doc, 18, "(SBT – CTST) Cho các phát biểu sau về carbohydrate:\n(a) Saccharose và fructose không phải là đường khử.\n(b) Amylopectin phân nhánh là do sự xuất hiện của liên kết alpha-1,6-glycoside.\n(c) Cellulose không phân nhánh do chỉ xuất hiện liên kết beta-1,4-glycoside.\n(d) Số nhóm -OH trong phân tử saccharose là 8.\nSố phát biểu đúng là:")
    add_options_1line(doc, "1.", "2.", "3.", "4.", correct="C", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "(a) Sai, fructose là đường khử vì trong môi trường base đồng phân hoá thành glucose.\n(b) Đúng, liên kết alpha-1,6-glycoside tạo nhánh.\n(c) Đúng, liên kết beta-1,4-glycoside tạo chuỗi mạch thẳng không nhánh.\n(d) Đúng, công thức saccharose có 8 nhóm -OH (C~12~H~14~O~3~(OH)~8~).\n=> Có 3 phát biểu đúng: (b), (c), (d)."
        ])

    # ==========================================
    # PHẦN II
    # ==========================================
    add_section_heading(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1 Phần II
    add_question(doc, 1, "(SBT – KNTT)", is_part2=True)
    c1_p2_items = [
        ("a", "Glucose và fructose đều có công thức phân tử là C~6~H~12~O~6~."),
        ("b", "Glucose và fructose đều tồn tại dạng mạch hở và mạch vòng."),
        ("c", "Glucose và fructose đều là pentahydroxy aldehyde."),
        ("d", "Dạng vòng của glucose và fructose đều là vòng sáu cạnh.")
    ]
    c1_p2_ans = {"a": "Đúng", "b": "Đúng", "c": "Sai", "d": "Sai"}
    add_part2_items(doc, c1_p2_items, c1_p2_ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "a) Đúng, cả 2 đều có CTPT C~6~H~12~O~6~.",
            "b) Đúng, trong dung dịch tồn tại cân bằng giữa dạng hở và các dạng vòng.",
            "c) Sai, glucose là pentahydroxy aldehyde, còn fructose là pentahydroxy ketone.",
            "d) Sai, dạng vòng của glucose chủ yếu là vòng 6 cạnh, còn fructose chủ yếu là vòng 5 cạnh."
        ])

    # Câu 2 Phần II
    add_question(doc, 2, "Cấu trúc phân tử saccharose như sau:", is_part2=True)
    add_image_centered(doc, "p2_cau2_saccharose.png", width_cm=7.5)
    c2_p2_items = [
        ("a", "Phân tử saccharose gồm một gốc alpha-glucose và một gốc beta-fructose."),
        ("b", "Phân tử saccharose có chứa nhóm -OH hemiacetal."),
        ("c", "Phân tử saccharose có chứa nhóm -OH hemiketal."),
        ("d", "Hai gốc monosaccharide trong saccharose được liên kết với nhau bởi một nguyên tử oxygen.")
    ]
    c2_p2_ans = {"a": "Đúng", "b": "Sai", "c": "Sai", "d": "Đúng"}
    add_part2_items(doc, c2_p2_items, c2_p2_ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "a) Đúng, saccharose gồm gốc alpha-glucose và beta-fructose.",
            "b) Sai, nhóm -OH hemiacetal ở C1 của gốc glucose đã tham gia liên kết glycoside.",
            "c) Sai, nhóm -OH hemiketal ở C2 của gốc fructose đã tham gia liên kết glycoside.",
            "d) Đúng, liên kết alpha-1,2-glycoside nối qua 1 nguyên tử oxygen (C1-O-C2)."
        ])

    # Câu 3 Phần II
    add_question(doc, 3, "Tiến hành thí nghiệm phản ứng của hồ tinh bột với iodine theo các bước sau đây:\nBước 1: Cho vài giọt dung dịch iodine vào ống nghiệm đựng sẵn 2 mL dung dịch hồ tinh bột.\nBước 2: Đun nóng dung dịch một lát, sau đó để nguội.", is_part2=True)
    add_image_centered(doc, "p2_cau3_tinh_bot_iodine.png", width_cm=5.5, caption="Mô hình tương tác giữa tinh bột & iodine")
    c3_p2_items = [
        ("a", "Sau bước 1, dung dịch thu được có màu xanh tím. Sau bước 2, dung dịch bị mất màu."),
        ("b", "Mạch tinh bột xoắn lại tạo các lỗ rỗng, các lỗ rỗng này hấp phụ I~2~ tạo nên hợp chất màu xanh tím."),
        ("c", "Nếu nhỏ vài giọt dung dịch iodine lên mặt cắt của quả chuối chín thì màu xanh tím cũng xuất hiện."),
        ("d", "Có thể dùng dung dịch iodine để phân biệt hai dung dịch riêng biệt gồm hồ tinh bột và saccharose.")
    ]
    c3_p2_ans = {"a": "Sai", "b": "Đúng", "c": "Sai", "d": "Đúng"}
    add_part2_items(doc, c3_p2_items, c3_p2_ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "a) Sai, sau khi đun nóng rồi để nguội thì màu xanh tím xuất hiện trở lại (chỉ mất màu khi đang đun nóng).",
            "b) Đúng, cấu trúc lò xo xoắn của amylose hấp phụ I~2~ tạo màu xanh tím đặc trưng.",
            "c) Sai, chuối chín tinh bột đã chuyển hoá hoàn toàn thành glucose và fructose nên không cho màu xanh tím.",
            "d) Đúng, hồ tinh bột chuyển màu xanh tím, saccharose không hiện tượng."
        ])

    # Câu 4 Phần II
    add_question(doc, 4, "Thí nghiệm phản ứng của cellulose với nitric acid\n"
        "Hoá chất: cellulose (bông), dung dịch HNO~3~ đặc, dung dịch H~2~SO~4~ đặc, dung dịch NaHCO~3~ loãng, quỳ tím.\n"
        "Dụng cụ: cốc thuỷ tinh 100 mL, chậu nước nóng, chậu nước đá, đũa thuỷ tinh, giấy lọc, đĩa sứ, đèn cồn.\n"
        "Tiến hành:\n"
        "Bước 1: Cho khoảng 5 mL dung dịch HNO~3~ đặc vào cốc thuỷ tinh (loại 100 mL) ngâm trong chậu nước đá. Thêm từ từ khoảng 10 mL dung dịch H~2~SO~4~ đặc vào cốc và khuấy đều.\n"
        "Bước 2: Sau đó, lấy cốc thuỷ tinh ra khỏi chậu nước đá, thêm tiếp một nhúm bông vào cốc và dùng đũa thuỷ tinh ấn bông ngập trong dung dịch. Ngâm cốc trong chậu nước nóng khoảng 10 phút.\n"
        "Bước 3: Để nguội, lấy sản phẩm thu được ra khỏi cốc, rửa nhiều lần với nước lạnh (đến khi nước rửa không làm đổi màu quỳ tím), sau đó rửa lại bằng dung dịch NaHCO~3~ loãng. Ép sản phẩm giữa hai miếng giấy lọc để hút nước và làm khô tự nhiên. Sau đó, để sản phẩm lên đĩa sứ rồi đốt cháy sản phẩm.", is_part2=True)
    add_image_centered(doc, "p2_cau4_cellulose_trinitrate.png", width_cm=8.5, caption="Cellulose trinitrate và Cellulose trinitrate cháy")
    c4_p2_items = [
        ("a", "Sau bước 3, sản phẩm thu được có màu vàng."),
        ("b", "Có thể thay thế nhúm bông bằng tinh bột."),
        ("c", "Sau bước 3, lấy sản phẩm thu được đốt cháy thấy có khói trắng xuất hiện."),
        ("d", "Thí nghiệm trên chứng minh mỗi mắt xích trong phân tử cellulose có 3 nhóm OH tự do.")
    ]
    c4_p2_ans = {"a": "Đúng", "b": "Sai", "c": "Sai", "d": "Đúng"}
    add_part2_items(doc, c4_p2_items, c4_p2_ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "a) Đúng, cellulose trinitrate có màu vàng ngà.",
            "b) Sai, tinh bột bị acid phân huỷ mạnh và tạo khối keo dính nhầy khó tách lọc, không thể thay thế bông.",
            "c) Sai, cellulose trinitrate là thuốc súng không khói, khi cháy bùng cháy rất nhanh không sinh khói.",
            "d) Đúng, phản ứng thế este hoá tạo [C~6~H~7~O~2~(ONO~2~)~3~]~n~ chứng minh có 3 nhóm -OH tự do trong mỗi mắt xích."
        ])

    # ==========================================
    # PHẦN III
    # ==========================================
    add_section_heading(
        doc,
        "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 6. Mỗi câu hỏi thí sinh điền câu trả lời vào ô tương ứng."
    )

    # Câu 1 Phần III
    add_question(doc, 1, "Cho các chất: saccharose, glucose, fructose, formic acid và acetic aldehyde. Trong các chất trên, số chất vừa có khả năng tham gia phản ứng tráng bạc vừa có khả năng phản ứng với Cu(OH)~2~ ở điều kiện thường là bao nhiêu?", is_part3=True)
    add_part3_answer_line(doc, ans_str="3", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "- Tác dụng với Cu(OH)~2~ ở đk thường: glucose, fructose, saccharose (polyalcohol có >= 2 nhóm -OH kề nhau) và formic acid (acid carboxylic).",
            "- Có phản ứng tráng bạc: glucose, fructose, formic acid, acetic aldehyde.",
            "=> Vừa tráng bạc vừa tác dụng với Cu(OH)~2~ ở đk thường gồm 3 chất: glucose, fructose, formic acid."
        ])

    # Câu 2 Phần III
    add_question(doc, 2, "Phân tử maltose ở dạng mạch vòng có bao nhiêu nhóm alcohol bậc 2?", is_part3=True)
    add_image_centered(doc, "p3_cau2_maltose.png", width_cm=7.0)
    add_part3_answer_line(doc, ans_str="5", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Quan sát công thức cấu tạo maltose mạch vòng:",
            "- Gốc alpha-glucose thứ nhất (bên trái): các nhóm -OH ở C2, C3, C4 đều gắn với carbon bậc 2 => 3 nhóm alcohol bậc 2.",
            "- Gốc alpha-glucose thứ hai (bên phải): các nhóm -OH ở C2, C3 gắn với carbon bậc 2 => 2 nhóm alcohol bậc 2 (nhóm -OH ở C1 là hemiacetal; nhóm ở C6 là bậc 1).",
            "=> Tổng số nhóm alcohol bậc 2 là 3 + 2 = 5."
        ])

    # Câu 3 Phần III
    add_question(doc, 3, "Cho các phản ứng sau: hoà tan Cu(OH)~2~ ở nhiệt độ thường tạo dung dịch xanh lam, phản ứng với thuốc thử Tollens, Cu(OH)~2~/OH^- (t°), thủy phân môi trường acid, phản ứng màu với dung dịch iodine, phản ứng với HNO~3~ đặc/H~2~SO~4~ đặc (t°). Số phản ứng mà cellulose có khả năng tham gia phản ứng là", is_part3=True)
    add_part3_answer_line(doc, ans_str="2", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Cellulose chỉ tham gia 2 phản ứng trong danh sách trên:",
            "1. Thủy phân trong môi trường acid tạo glucose.",
            "2. Phản ứng với HNO~3~ đặc / H~2~SO~4~ đặc (t°) tạo cellulose trinitrate."
        ])

    # Câu 4 Phần III
    add_question(doc, 4, "Trong y học, glucose làm thuốc tăng lực cho người bệnh, dễ hấp thu và cung cấp khá nhiều năng lượng. Phản ứng oxi hóa glucose tạo thành CO~2~ và H~2~O cung cấp năng lượng cho cơ thể. Một người bệnh được truyền hai chai chứa 250 mL dung dịch glucose 5%. Tính năng lượng tối đa bệnh nhân nhận được từ các chai truyền glucose đó. Biết nhiệt tạo thành của glucose C~6~H~12~O~6~, CO~2~ và H~2~O lần lượt là -1 271 kJ/mol, -393,5 kJ/mol và -285,8 kJ/mol; dung dịch glucose 5% có khối lượng riêng là 1,019 g/mL. (Làm tròn kết quả đến hàng đơn vị).", is_part3=True)
    add_part3_answer_line(doc, ans_str="397", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Phương trình nhiệt hoá học: C~6~H~12~O~6~ + 6O~2~ -> 6CO~2~ + 6H~2~O",
            "delta_r H°_298 = 6.(-393,5) + 6.(-285,8) - (-1271) = -2804,8 kJ.",
            "=> Oxi hoá 1 mol glucose toả ra 2804,8 kJ.",
            "Khối lượng dung dịch 2 chai: m_dd = 2 . 250 . 1,019 = 509,5 gam.",
            "Khối lượng glucose: m_glucose = 509,5 . 5% = 25,475 gam.",
            "Năng lượng tối đa: Q = (25,475 / 180) . 2804,8 = 396,96 kJ -> làm tròn: 397."
        ])

    # Câu 5 Phần III
    add_question(doc, 5, "Trong công nghiệp, saccharose là nguyên liệu để thủy phân thành glucose và fructose dùng trong kĩ thuật tráng gương, ruột phích. Để thu được 30 kg glucose cần thuỷ phân m kg saccharose với hiệu suất phản ứng là 60%. Giá trị của m là", is_part3=True)
    add_part3_answer_line(doc, ans_str="95", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "Phương trình thuỷ phân: C~12~H~22~O~11~ + H~2~O -> C~6~H~12~O~6~ (glucose) + C~6~H~12~O~6~ (fructose)",
            "Theo lí thuyết: 342 kg saccharose tạo ra 180 kg glucose.",
            "Khối lượng saccharose cần dùng theo lí thuyết: m_lt = 30 . 342 / 180 = 57 kg.",
            "Vì hiệu suất H = 60% nên thực tế cần: m = 57 . (100 / 60) = 95 kg."
        ])

    # Câu 6 Phần III
    add_question(doc, 6, "Một người nông dân có 5 sào ruộng để cấy lúa (3 sào lấy thóc dùng làm thức ăn và 2 sào dùng nấu rượu (ethyl alcohol). Biết năng suất của mỗi sào lúa là 180 kg/sào và một năm thu hoạch 2 vụ. Hàm lượng tinh bột trong gạo là 70%, biết 1 kg thóc sau khi xát sẽ thu được 0,7 kg gạo, ethyl alcohol được nấu là 30° với hiệu suất đạt 75% (khối lượng riêng của ethyl alcohol D = 0,8 gam/mL). Nếu được bán với giá 20.000 (đồng/lít) thì số tiền thu được từ việc bán rượu (ethyl alcohol) trong một năm của người nông dân là a triệu (làm tròn một chữ số thập phân). Tính a.", is_part3=True)
    add_part3_answer_line(doc, ans_str="12,5", is_teacher=is_teacher)
    if is_teacher:
        add_solution_paragraph(doc, [
            "- Khối lượng thóc dùng nấu rượu: m_thóc = 2 sào . 180 kg/sào . 2 vụ = 720 kg.",
            "- Khối lượng gạo: m_gạo = 720 . 0,7 = 504 kg.",
            "- Khối lượng tinh bột phản ứng (H = 75%): m_tinh_bột = 504 . 70% . 75% = 264,6 kg.",
            "- Số mol mắt xích tinh bột: n = 264,6 / 162 = 49/30 kmol.",
            "- Số mol ethanol sinh ra: n_ethanol = 2 . (49/30) = 49/15 kmol.",
            "- Khối lượng ethanol nguyên chất: m_ethanol = (49/15) . 46 = 2254/15 kg.",
            "- Thể tích ethanol nguyên chất: V_ethanol = (2254/15) : 0,8 = 1127/6 lít.",
            "- Thể tích rượu 30° thu được: V_rượu = (1127/6) : 0,3 = 626,11 lít.",
            "- Số tiền thu được: 626,11 . 20.000 = 12.522.222 đồng = 12,5 triệu đồng.",
            "=> Giá trị a = 12,5."
        ])

    if is_teacher:
        add_summary_table(doc)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(4)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("--- HẾT ---")
    set_run_style(r_end, font_name="Times New Roman", font_size=11, bold=True, italic=True)

    out_file = os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_12_De06_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_12_De06.docx")
    doc.save(out_file)
    print(f"Saved: {out_file}")

if __name__ == "__main__":
    build_document(is_teacher=False)
    build_document(is_teacher=True)
