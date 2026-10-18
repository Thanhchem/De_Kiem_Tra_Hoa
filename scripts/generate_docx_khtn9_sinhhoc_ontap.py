# -*- coding: utf-8 -*-
"""
Script tạo trọn bộ file Word (.docx) ĐỀ CƯƠNG ÔN TẬP KIỂM TRA ĐỊNH KÌ - KHTN 9 (PHÂN MÔN SINH HỌC)
Trường THCS Kim Long - Thầy TRẦN VĂN THẠNH - 0777.470.803

Tạo 3 phiên bản chuẩn hóa theo FORM ĐỀ KIỂM TRA:
1. Bản Học sinh: San_Pham/De_Kiem_Tra_KHTN_9_SinhHoc_HocSinh.docx
2. Bản Giáo viên: San_Pham/De_Kiem_Tra_KHTN_9_SinhHoc_GiaoVien.docx
3. Bản Dạy trên màn hình: San_Pham/De_Kiem_Tra_KHTN_9_SinhHoc_DayTrenManHinh.docx
"""

import os
import sys
import re
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

sys.path.append(os.path.dirname(__file__))
sys.stdout.reconfigure(encoding='utf-8')

from khtn9_sinhhoc_ontap_data import (
    INFO, TF_QUESTIONS,
    SHORT_QUESTIONS_A, SHORT_QUESTIONS_B, SHORT_QUESTIONS_C,
    ESSAY_QUESTIONS
)

ROOT_DIR = r"c:\Antigravity_Thanh"
OUT_DIR = os.path.join(ROOT_DIR, "San_Pham")
ONEDRIVE_DIRS = [
    r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9\KHTN9-DEKIEMTRADINHKI",
    r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9"
]

# Bảng màu chuẩn mực
COLOR_HF_RED = RGBColor(211, 47, 47)         # Đỏ san hô (#D32F2F)
COLOR_HF_GREEN_HEX = "8BC390"                # Xanh Sage Green
COLOR_PRIMARY = RGBColor(46, 125, 50)        # Xanh lá đậm (#2E7D32)
COLOR_CORRECT = RGBColor(198, 40, 40)        # Đỏ đậm nhấn đáp án đúng (#C62828)
COLOR_WRONG = RGBColor(0, 105, 92)           # Xanh tím than cho [SAI]
COLOR_SOLUTION = RGBColor(21, 101, 192)      # Xanh lam đậm (#1565C0) cho tiêu đề lời giải
COLOR_SCREEN_TITLE = RGBColor(106, 27, 154)  # Tím thanh lịch cho bản trình chiếu
COLOR_GRAY = RGBColor(97, 97, 97)            # Xám ghi chú
COLOR_BORDER_GREY_HEX = "CCCCCC"

def add_p_border_bottom(p, color="8BC390", sz="6"):
    """Thêm đường viền đáy cho đoạn văn bản."""
    pPr = p._element.get_or_add_pPr()
    xml = f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/></w:pBdr>'
    pPr.append(parse_xml(xml))

def add_p_border_top(p, color="8BC390", sz="6"):
    """Thêm đường viền trên cho đoạn văn bản."""
    pPr = p._element.get_or_add_pPr()
    xml = f'<w:pBdr {nsdecls("w")}><w:top w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/></w:pBdr>'
    pPr.append(parse_xml(xml))

def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    """Cài đặt lề đệm cho ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_shading(cell, color_hex="F1F8E9"):
    """Đặt màu nền cho ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_table_borders(table, color="CCCCCC", sz="4"):
    """Đặt viền thanh mảnh cho bảng."""
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'  <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'  <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

def set_run_style(run, font_name="Times New Roman", font_size=10.5, bold=False, italic=False, subscript=False, superscript=False, color=None):
    """Cài đặt style chi tiết cho từng run text."""
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

def add_page_number_field(p, font_name="Times New Roman", font_size=10, bold=True, color=COLOR_HF_RED):
    """Chèn trường số trang tự động PAGE vào chân trang."""
    run_txt = p.add_run("Trang ")
    set_run_style(run_txt, font_name=font_name, font_size=font_size, bold=bold, color=color)
    
    r = p.add_run()
    set_run_style(r, font_name=font_name, font_size=font_size, bold=bold, color=color)
    fld1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instr1 = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
    fld2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fld3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    r._r.append(fld1); r._r.append(instr1); r._r.append(fld2); r._r.append(fld3)

def render_rich_text(p, text, base_font="Times New Roman", base_size=10.5, base_bold=False, base_italic=False, base_color=None):
    """
    Sinh văn bản định dạng phong phú từ ký hiệu:
    - *bold*
    - _italic_
    - ~subscript~ (F~1~, F~2~, P~tc~)
    - ^superscript^ (2^n, 5', 3')
    """
    tokens = re.split(r'(~[^~]+~|\^[^\^]+\^|\*[^\*]+\*|_[^_]+_)', text)
    for token in tokens:
        if not token:
            continue

        is_bold = base_bold
        is_italic = base_italic
        is_sub = False
        is_sup = False
        content = token

        if token.startswith('*') and token.endswith('*'):
            is_bold = True
            content = token[1:-1]
        elif token.startswith('_') and token.endswith('_'):
            is_italic = True
            content = token[1:-1]
        elif token.startswith('~') and token.endswith('~'):
            is_sub = True
            content = token[1:-1]
        elif token.startswith('^') and token.endswith('^'):
            is_sup = True
            content = token[1:-1]

        run = p.add_run(content)
        run_color = COLOR_CORRECT if (is_bold and not base_bold and not base_color) else base_color
        set_run_style(run, font_name=base_font, font_size=base_size, bold=is_bold, italic=is_italic,
                      subscript=is_sub, superscript=is_sup, color=run_color)

def setup_header_footer(doc):
    """Cài đặt Header & Footer chuẩn theo FORM_DE_KIEM_TRA."""
    normal_style = doc.styles['Normal']
    normal_style.paragraph_format.tab_stops.clear_all()

    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    section.start_type = WD_SECTION_START.NEW_PAGE
    section.different_first_page_header_footer = False
    section.header_distance = Cm(0.6)
    section.footer_distance = Cm(0.6)

    # 1. Header (Bảng 1x2, rộng 18.0 cm)
    header = section.header
    for p in header.paragraphs:
        p.text = ""
    tbl_h = header.add_table(rows=1, cols=2, width=Cm(18.0))
    tbl_h.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_left = tbl_h.cell(0, 0)
    c_right = tbl_h.cell(0, 1)
    c_left.width = Cm(8.0)
    c_right.width = Cm(10.0)

    p_hl = c_left.paragraphs[0]
    p_hl.paragraph_format.space_before = Pt(0)
    p_hl.paragraph_format.space_after = Pt(2)
    p_hl.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_hl = p_hl.add_run("Tài liệu lưu hành nội bộ")
    set_run_style(r_hl, font_size=10, italic=True, color=COLOR_HF_RED)

    p_hr = c_right.paragraphs[0]
    p_hr.paragraph_format.space_before = Pt(0)
    p_hr.paragraph_format.space_after = Pt(2)
    p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1 = p_hr.add_run("Thầy ")
    set_run_style(r1, font_size=10, color=COLOR_HF_RED)
    r2 = p_hr.add_run("TRẦN VĂN THẠNH")
    set_run_style(r2, font_size=10, bold=True, color=COLOR_HF_RED)
    r3 = p_hr.add_run(" - 0777.470.803")
    set_run_style(r3, font_size=10, color=COLOR_HF_RED)

    add_p_border_bottom(p_hl, color=COLOR_HF_GREEN_HEX, sz="6")
    add_p_border_bottom(p_hr, color=COLOR_HF_GREEN_HEX, sz="6")

    # 2. Footer (Bảng 1x2, rộng 18.0 cm)
    footer = section.footer
    for p in footer.paragraphs:
        p.text = ""
    tbl_f = footer.add_table(rows=1, cols=2, width=Cm(18.0))
    tbl_f.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_fl = tbl_f.cell(0, 0)
    c_fr = tbl_f.cell(0, 1)
    c_fl.width = Cm(14.5)
    c_fr.width = Cm(3.5)

    p_fl = c_fl.paragraphs[0]
    p_fl.paragraph_format.space_before = Pt(2)
    p_fl.paragraph_format.space_after = Pt(0)
    p_fl.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_fl = p_fl.add_run("CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế   |   CS2: 24 Đặng Thái Thân, TP Huế")
    set_run_style(r_fl, font_size=9.5, italic=True, color=COLOR_HF_RED)

    p_fr = c_fr.paragraphs[0]
    p_fr.paragraph_format.space_before = Pt(2)
    p_fr.paragraph_format.space_after = Pt(0)
    p_fr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_page_number_field(p_fr, font_size=10, bold=True, color=COLOR_HF_RED)

    add_p_border_top(p_fl, color=COLOR_HF_GREEN_HEX, sz="6")
    add_p_border_top(p_fr, color=COLOR_HF_GREEN_HEX, sz="6")

def build_title_block(doc, is_teacher=False, is_screen=False, base_size=10.5):
    """Xây dựng Header Block không viền, chuẩn mực."""
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_l, cell_r = tbl.cell(0, 0), tbl.cell(0, 1)
    cell_l.width = Cm(8.0)
    cell_r.width = Cm(10.0)

    # Cột trái
    p_tl1 = cell_l.paragraphs[0]
    p_tl1.paragraph_format.space_before, p_tl1.paragraph_format.space_after = Pt(0), Pt(1)
    p_tl1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tl1, f"*{INFO['school']}*", base_size=10.5, base_bold=True)

    p_tl2 = cell_l.add_paragraph()
    p_tl2.paragraph_format.space_before, p_tl2.paragraph_format.space_after = Pt(0), Pt(1)
    p_tl2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tl2, f"*{INFO['year']}*", base_size=10)

    p_tl3 = cell_l.add_paragraph()
    p_tl3.paragraph_format.space_before, p_tl3.paragraph_format.space_after = Pt(0), Pt(0)
    p_tl3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_off = p_tl3.add_run("TÀI LIỆU ÔN TẬP ĐỘC QUYỀN")
    set_run_style(r_off, font_size=9, bold=True, color=COLOR_PRIMARY)

    # Cột phải
    p_tr1 = cell_r.paragraphs[0]
    p_tr1.paragraph_format.space_before, p_tr1.paragraph_format.space_after = Pt(0), Pt(1)
    p_tr1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tr1, f"*{INFO['title']}*", base_size=11.5, base_bold=True, base_color=COLOR_HF_RED)

    p_tr2 = cell_r.add_paragraph()
    p_tr2.paragraph_format.space_before, p_tr2.paragraph_format.space_after = Pt(0), Pt(1)
    p_tr2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tr2, f"*{INFO['subject']}*", base_size=11, base_bold=True, base_color=COLOR_PRIMARY)

    p_tr3 = cell_r.add_paragraph()
    p_tr3.paragraph_format.space_before, p_tr3.paragraph_format.space_after = Pt(0), Pt(0)
    p_tr3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tr3, f"_{INFO['time_str']}_", base_size=9.5, base_italic=True)

    if is_teacher:
        p_gv = doc.add_paragraph()
        p_gv.paragraph_format.space_before, p_gv.paragraph_format.space_after = Pt(4), Pt(2)
        p_gv.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_gv = p_gv.add_run("★ BẢN GIÁO VIÊN — CÓ ĐÁP ÁN & LỜI GIẢI CHI TIẾT ★")
        set_run_style(r_gv, font_size=11, bold=True, color=COLOR_CORRECT)
    elif is_screen:
        p_scr = doc.add_paragraph()
        p_scr.paragraph_format.space_before, p_scr.paragraph_format.space_after = Pt(4), Pt(2)
        p_scr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_scr = p_scr.add_run("🖥 BẢN TRÌNH CHIẾU GIẢNG DẠY TRÊN LỚP — TƯƠNG TÁC MÀN HÌNH 🖥")
        set_run_style(r_scr, font_size=11.5, bold=True, color=COLOR_SCREEN_TITLE)

    # Dòng thông tin thí sinh
    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(4)
    p_info.paragraph_format.space_after = Pt(4)
    p_info.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_info = p_info.add_run("Họ, tên học sinh: ................................................................ Lớp: ............. Số báo danh: .............")
    set_run_style(r_info, font_size=10, italic=True)

def add_major_section_header(doc, title, sub_title="", base_size=11.5):
    """Tiêu đề phần lớn với thanh nền nổi bật."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True

    pPr = p._element.get_or_add_pPr()
    shd_xml = f'<w:shd {nsdecls("w")} w:fill="E8F5E9"/>'
    pbdr_xml = f'<w:pBdr {nsdecls("w")}><w:left w:val="single" w:sz="18" w:space="8" w:color="2E7D32"/></w:pBdr>'
    pPr.append(parse_xml(shd_xml))
    pPr.append(parse_xml(pbdr_xml))

    r_title = p.add_run(f" {title.upper()} ")
    set_run_style(r_title, font_size=base_size, bold=True, color=COLOR_PRIMARY)
    if sub_title:
        p2 = doc.add_paragraph()
        p2.paragraph_format.left_indent = Cm(0.5)
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(3)
        p2.paragraph_format.keep_with_next = True
        r_sub = p2.add_run(f"({sub_title})")
        set_run_style(r_sub, font_size=base_size - 1.5, italic=True, color=COLOR_GRAY)

def add_section_title(doc, title, base_size=11.0):
    """Tiêu đề mục nhỏ."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    render_rich_text(p, f"*{title}*", base_size=base_size, base_bold=True, base_color=COLOR_PRIMARY)

def add_solution_block(doc, solution_text, base_size=10.0):
    """Hiển thị khối lời giải chi tiết (không dùng Table, thụt lề chuẩn)."""
    p_sol = doc.add_paragraph()
    p_sol.paragraph_format.left_indent = Cm(0.6)
    p_sol.paragraph_format.space_before = Pt(2)
    p_sol.paragraph_format.space_after = Pt(1)
    p_sol.paragraph_format.line_spacing = 1.15

    r_lbl = p_sol.add_run("► Lời giải chi tiết:\n")
    set_run_style(r_lbl, font_size=base_size, bold=True, color=COLOR_SOLUTION)

    lines = solution_text.split('\n')
    for idx, line in enumerate(lines):
        line_str = line.strip()
        if not line_str:
            continue
        if idx > 0:
            p_sol = doc.add_paragraph()
            p_sol.paragraph_format.left_indent = Cm(0.6)
            p_sol.paragraph_format.space_before = Pt(1)
            p_sol.paragraph_format.space_after = Pt(1)
            p_sol.paragraph_format.line_spacing = 1.15
        render_rich_text(p_sol, line_str, base_size=base_size)

def add_screen_teaching_space(doc, num_lines=3):
    """Chừa khoảng trống tương tác cho bản Dạy Trên Màn Hình."""
    p_box = doc.add_paragraph()
    p_box.paragraph_format.left_indent = Cm(0.5)
    p_box.paragraph_format.space_before = Pt(3)
    p_box.paragraph_format.space_after = Pt(5)
    
    # Viền nét đứt nhẹ, viền trái tím đậm
    pPr = p_box._element.get_or_add_pPr()
    xml = f'<w:pBdr {nsdecls("w")}><w:top w:val="dashed" w:sz="4" w:space="4" w:color="B0BEC5"/><w:bottom w:val="dashed" w:sz="4" w:space="4" w:color="B0BEC5"/><w:left w:val="single" w:sz="12" w:space="6" w:color="7E57C2"/></w:pBdr>'
    pPr.append(parse_xml(xml))
    
    r_hint = p_box.add_run("✎ Ghi chú & Lời giải trực tiếp trên màn hình:\n")
    set_run_style(r_hint, font_size=9.5, bold=True, color=COLOR_SCREEN_TITLE)
    for _ in range(num_lines):
        p_box.add_run("\n")

def add_student_writing_lines(doc, num_lines=4):
    """Thêm dòng chấm cho học sinh làm bài tự luận."""
    for _ in range(num_lines):
        p_line = doc.add_paragraph()
        p_line.paragraph_format.left_indent = Cm(0.5)
        p_line.paragraph_format.space_before = Pt(1)
        p_line.paragraph_format.space_after = Pt(2)
        r = p_line.add_run("...................................................................................................................................................................")
        set_run_style(r, font_size=9, color=RGBColor(189, 189, 189))

def render_tf_questions(doc, questions, is_teacher=False, is_screen=False, base_size=10.5):
    """Hiển thị các câu hỏi đúng / sai chuẩn mực."""
    for q in questions:
        qid = q["id"]
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        p_q.paragraph_format.keep_with_next = True
        render_rich_text(p_q, f"*Câu {qid}.* {q['content']}", base_size=base_size)

        for st in q["statements"]:
            sub_id = st["id"]
            st_text = st["text"]
            correct_val = "Đúng" if st["correct"] else "Sai"

            p_st = doc.add_paragraph()
            p_st.paragraph_format.left_indent = Cm(0.5)
            p_st.paragraph_format.space_before = Pt(1)
            p_st.paragraph_format.space_after = Pt(1)
            p_st.paragraph_format.line_spacing = 1.15

            if is_teacher:
                r_tag = p_st.add_run(f"{sub_id}) ")
                set_run_style(r_tag, font_size=base_size, bold=True)
                
                ans_color = COLOR_CORRECT if st["correct"] else COLOR_WRONG
                r_ans = p_st.add_run(f"[{correct_val.upper()}] ")
                set_run_style(r_ans, font_size=base_size, bold=True, color=ans_color)

                render_rich_text(p_st, st_text, base_size=base_size)
            else:
                r_tag = p_st.add_run(f"{sub_id}) ")
                set_run_style(r_tag, font_size=base_size, bold=True)
                render_rich_text(p_st, st_text, base_size=base_size)

        if is_teacher:
            add_solution_block(doc, q["explanation"], base_size=base_size - 0.5)
        elif is_screen:
            add_screen_teaching_space(doc, num_lines=3)

def render_short_questions(doc, questions, group_title="", is_teacher=False, is_screen=False, base_size=10.5):
    """Hiển thị các câu hỏi trả lời ngắn theo nhóm."""
    if group_title:
        p_grp = doc.add_paragraph()
        p_grp.paragraph_format.space_before = Pt(4)
        p_grp.paragraph_format.space_after = Pt(1)
        p_grp.paragraph_format.keep_with_next = True
        render_rich_text(p_grp, f"*{group_title}*", base_size=base_size, base_bold=True, base_color=COLOR_PRIMARY)

    for q in questions:
        qid = q["id"]
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        p_q.paragraph_format.keep_with_next = True
        render_rich_text(p_q, f"*Câu {qid}.* {q['content']}", base_size=base_size)

        if is_teacher:
            add_solution_block(doc, f"Đáp án: *{q['correct']}*\n" + q["explanation"], base_size=base_size - 0.5)
        elif is_screen:
            add_screen_teaching_space(doc, num_lines=3)
        else:
            p_ans_box = doc.add_paragraph()
            p_ans_box.paragraph_format.left_indent = Cm(0.5)
            p_ans_box.paragraph_format.space_before = Pt(1)
            p_ans_box.paragraph_format.space_after = Pt(2)
            r_box = p_ans_box.add_run("Trả lời: .............................................................................................................................................")
            set_run_style(r_box, font_size=10, italic=True, color=COLOR_GRAY)

def render_essay_questions(doc, questions, is_teacher=False, is_screen=False, base_size=10.5):
    """Hiển thị các câu hỏi tự luận."""
    for q in questions:
        qid = q["id"]
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        p_q.paragraph_format.keep_with_next = True
        render_rich_text(p_q, f"*Câu {qid}.* {q['content']}", base_size=base_size)

        # Chèn hình ảnh nếu có
        if q.get("has_image") and os.path.exists(q.get("image_path", "")):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(3)
            p_img.paragraph_format.space_after = Pt(1)
            run_img = p_img.add_run()
            run_img.add_picture(q["image_path"], width=Cm(6.5))

            if q.get("image_caption"):
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cap.paragraph_format.space_before = Pt(0)
                p_cap.paragraph_format.space_after = Pt(3)
                r_cap = p_cap.add_run(q["image_caption"])
                set_run_style(r_cap, font_size=9, italic=True, color=COLOR_GRAY)

        if is_teacher:
            add_solution_block(doc, q["solution"], base_size=base_size - 0.5)
        elif is_screen:
            add_screen_teaching_space(doc, num_lines=5)
        else:
            add_student_writing_lines(doc, num_lines=6)

def render_summary_answer_table(doc, base_size=10.0):
    """Bảng tổng hợp đáp án nhanh cho giáo viên ở cuối tài liệu."""
    p_t = doc.add_paragraph()
    p_t.paragraph_format.space_before = Pt(12)
    p_t.paragraph_format.space_after = Pt(4)
    p_t.paragraph_format.keep_with_next = True
    r_t = p_t.add_run("BẢNG TỔNG HỢP ĐÁP ÁN NHANH (GIÁO VIÊN)")
    set_run_style(r_t, font_size=11, bold=True, color=COLOR_PRIMARY)

    # 1. Bảng Đúng / Sai
    p_st = doc.add_paragraph()
    p_st.paragraph_format.space_before = Pt(2)
    p_st.paragraph_format.space_after = Pt(2)
    render_rich_text(p_st, "*1. Bảng đáp án Trắc nghiệm Đúng / Sai:*", base_size=10, base_bold=True)

    tbl_tf = doc.add_table(rows=6, cols=5)
    tbl_tf.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_tf, color="CCCCCC", sz="4")

    headers = ["Câu", "Ý a", "Ý b", "Ý c", "Ý d"]
    for c_idx, h in enumerate(headers):
        cell = tbl_tf.cell(0, c_idx)
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        set_cell_shading(cell, "E8F5E9")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_run_style(r, font_size=10, bold=True, color=COLOR_PRIMARY)

    for r_idx, q in enumerate(TF_QUESTIONS, start=1):
        cell_q = tbl_tf.cell(r_idx, 0)
        set_cell_margins(cell_q, top=50, bottom=50, left=80, right=80)
        p = cell_q.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"Câu {q['id']}")
        set_run_style(r, font_size=9.5, bold=True)

        for s_idx, st in enumerate(q["statements"], start=1):
            cell_ans = tbl_tf.cell(r_idx, s_idx)
            set_cell_margins(cell_ans, top=50, bottom=50, left=80, right=80)
            p_a = cell_ans.paragraphs[0]
            p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
            val_txt = "Đúng" if st["correct"] else "Sai"
            val_col = COLOR_CORRECT if st["correct"] else COLOR_WRONG
            r_a = p_a.add_run(val_txt)
            set_run_style(r_a, font_size=9.5, bold=True, color=val_col)

    # 2. Bảng Trả lời ngắn
    p_sn = doc.add_paragraph()
    p_sn.paragraph_format.space_before = Pt(6)
    p_sn.paragraph_format.space_after = Pt(2)
    render_rich_text(p_sn, "*2. Bảng đáp số Câu hỏi trả lời ngắn (15 câu):*", base_size=10, base_bold=True)

    all_short = SHORT_QUESTIONS_A + SHORT_QUESTIONS_B + SHORT_QUESTIONS_C
    tbl_sh = doc.add_table(rows=4, cols=10) # 2 hàng tiêu đề & đáp án chia đôi (8 cột)
    tbl_sh.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_sh, color="CCCCCC", sz="4")

    # Bảng 2 tầng: Hàng 1 (Câu 1-8), Hàng 2 (Đáp số), Hàng 3 (Câu 9-15), Hàng 4 (Đáp số)
    # Tầng 1: Câu 1 đến 8
    for i in range(8):
        c_h = tbl_sh.cell(0, i)
        set_cell_margins(c_h, top=50, bottom=50, left=60, right=60)
        set_cell_shading(c_h, "F1F8E9")
        p = c_h.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"C{i+1}")
        set_run_style(r, font_size=9.5, bold=True, color=COLOR_PRIMARY)

        c_v = tbl_sh.cell(1, i)
        set_cell_margins(c_v, top=50, bottom=50, left=60, right=60)
        p = c_v.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(all_short[i]["correct"])
        set_run_style(r, font_size=9.5, bold=True, color=COLOR_CORRECT)

    # Tầng 2: Câu 9 đến 15
    for i in range(7):
        c_h = tbl_sh.cell(2, i)
        set_cell_margins(c_h, top=50, bottom=50, left=60, right=60)
        set_cell_shading(c_h, "F1F8E9")
        p = c_h.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"C{i+9}")
        set_run_style(r, font_size=9.5, bold=True, color=COLOR_PRIMARY)

        c_v = tbl_sh.cell(3, i)
        set_cell_margins(c_v, top=50, bottom=50, left=60, right=60)
        p = c_v.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(all_short[i+8]["correct"])
        set_run_style(r, font_size=9.5, bold=True, color=COLOR_CORRECT)

def build_exam_docx(output_path, mode="student"):
    """
    Hàm chính tạo file Word đề kiểm tra / ôn tập
    mode: 'student' | 'teacher' | 'screen'
    """
    is_teacher = (mode == "teacher")
    is_screen = (mode == "screen")
    base_size = 11.5 if is_screen else 10.5

    doc = docx.Document()
    setup_header_footer(doc)
    build_title_block(doc, is_teacher=is_teacher, is_screen=is_screen, base_size=base_size)

    # =========================================================================
    # PHẦN I: TRẮC NGHIỆM ĐÚNG SAI (5 CÂU)
    # =========================================================================
    add_major_section_header(
        doc,
        "PHẦN I: TRẮC NGHIỆM ĐÚNG - SAI",
        "Thí sinh trả lời từ câu 1 đến câu 5. Trong mỗi ý a), b), c), d) ở mỗi câu, chọn Đúng hoặc Sai",
        base_size=base_size + 0.5
    )
    render_tf_questions(doc, TF_QUESTIONS, is_teacher=is_teacher, is_screen=is_screen, base_size=base_size)

    # =========================================================================
    # PHẦN II: CÂU HỎI TRẢ LỜI NGẮN (15 CÂU)
    # =========================================================================
    add_major_section_header(
        doc,
        "PHẦN II: CÂU HỎI TRẢ LỜI NGẮN",
        "Thí sinh trả lời từ câu 1 đến câu 15. Ghi đáp số hoặc kết luận ngắn gọn vào phần trả lời",
        base_size=base_size + 0.5
    )
    render_short_questions(doc, SHORT_QUESTIONS_A, group_title="A. Quy luật di truyền của Menđen", is_teacher=is_teacher, is_screen=is_screen, base_size=base_size)
    render_short_questions(doc, SHORT_QUESTIONS_B, group_title="B. Cấu trúc và số lượng nucleotide trong DNA", is_teacher=is_teacher, is_screen=is_screen, base_size=base_size)
    render_short_questions(doc, SHORT_QUESTIONS_C, group_title="C. RNA và gene", is_teacher=is_teacher, is_screen=is_screen, base_size=base_size)

    # =========================================================================
    # PHẦN III: TỰ LUẬN (4 CÂU)
    # =========================================================================
    add_major_section_header(
        doc,
        "PHẦN III: TỰ LUẬN",
        "Trình bày chi tiết các bước giải, quy ước gene, sơ đồ lai và lập luận khoa học",
        base_size=base_size + 0.5
    )
    render_essay_questions(doc, ESSAY_QUESTIONS, is_teacher=is_teacher, is_screen=is_screen, base_size=base_size)

    # =========================================================================
    # BẢNG TỔNG HỢP ĐÁP ÁN (DÀNH CHO BẢN GIÁO VIÊN)
    # =========================================================================
    if is_teacher:
        render_summary_answer_table(doc, base_size=base_size)

    # Kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before, p_end.paragraph_format.space_after = Pt(12), Pt(4)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("--- HẾT ---")
    set_run_style(r_end, font_size=11, bold=True)

    doc.save(output_path)
    print(f"Created: {output_path}")

def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    files = [
        ("De_Kiem_Tra_KHTN_9_SinhHoc_HocSinh.docx", "student"),
        ("De_Kiem_Tra_KHTN_9_SinhHoc_GiaoVien.docx", "teacher"),
        ("De_Kiem_Tra_KHTN_9_SinhHoc_DayTrenManHinh.docx", "screen"),
        # Bản sao với tên De_Cuong cho đồng bộ
        ("De_Cuong_KHTN_9_SinhHoc_HocSinh.docx", "student"),
        ("De_Cuong_KHTN_9_SinhHoc_GiaoVien.docx", "teacher"),
        ("De_Cuong_KHTN_9_SinhHoc_DayTrenManHinh.docx", "screen"),
    ]

    for filename, mode in files:
        out_file = os.path.join(OUT_DIR, filename)
        build_exam_docx(out_file, mode=mode)
        
        # Sao chép sang thư mục OneDrive của Thầy
        for onedrive_folder in ONEDRIVE_DIRS:
            if os.path.exists(onedrive_folder):
                try:
                    dest = os.path.join(onedrive_folder, filename)
                    shutil.copy2(out_file, dest)
                    print(f"Copied to OneDrive successfully: {dest}")
                except Exception as e:
                    print(f"Error copying to OneDrive: {e}")

if __name__ == "__main__":
    main()
