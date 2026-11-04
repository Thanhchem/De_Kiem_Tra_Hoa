# -*- coding: utf-8 -*-
"""
Script tạo 3 file Word (.docx) chuẩn "FORM ĐỀ KIỂM TRA" cho Đề Kiểm Tra Giữa Học Kỳ I - KHTN 9
Trường THCS Trần Cao Vân - Năm học 2026 - 2027
Thầy TRẦN VĂN THẠNH - 0777.470.803

Tạo 3 phiên bản:
1. Bản Học sinh: San_Pham/De_Kiem_Tra_KHTN_9_Giua_HK1_Tran_Cao_Van_HocSinh.docx
2. Bản Giáo viên: San_Pham/De_Kiem_Tra_KHTN_9_Giua_HK1_Tran_Cao_Van_GiaoVien.docx
3. Bản Dạy trên màn hình: San_Pham/De_Kiem_Tra_KHTN_9_Giua_HK1_Tran_Cao_Van_DayTrenManHinh.docx
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

from khtn9_de_kiem_tra_trancaovan_data import (
    EXAM_INFO, PART1_QUESTIONS, PART2_QUESTIONS, PART3_QUESTIONS, ESSAY_QUESTIONS
)

ROOT_DIR = r"c:\Antigravity_Thanh"
OUT_DIR = os.path.join(ROOT_DIR, "San_Pham")
ONEDRIVE_DIRS = [
    r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9\KHTN9-DEKIEMTRADINHKI",
    r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9"
]
os.makedirs(OUT_DIR, exist_ok=True)

# Bảng màu chuẩn quy định trong FORM_DE_KIEM_TRA
COLOR_HF_RED = RGBColor(211, 47, 47)        # Coral Red (#D32F2F) cho Header & Footer
COLOR_HF_GREEN_HEX = "8BC390"               # Sage Green (#8BC390) cho viền Header & Footer
COLOR_PRIMARY = RGBColor(46, 125, 50)       # Xanh lá đậm (#2E7D32) cho Tiêu đề Phần
COLOR_CORRECT = RGBColor(198, 40, 40)       # Đỏ (#C62828) cho đáp án đúng Bản Giáo Viên
COLOR_SOLUTION = RGBColor(21, 101, 192)     # Xanh lam (#1565C0) cho tiêu đề Lời giải
COLOR_SCREEN_TITLE = RGBColor(106, 27, 154) # Tím đậm (#6A1B9A) cho tiêu đề bản Dạy Màn Hình
COLOR_GRAY = RGBColor(97, 97, 97)
COLOR_BORDER_GREY_HEX = "CCCCCC"

def add_p_border_bottom(p, color="8BC390", sz="6"):
    """Thêm đường kẻ viền dưới cho đoạn văn bản."""
    pPr = p._element.get_or_add_pPr()
    xml = f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/></w:pBdr>'
    pPr.append(parse_xml(xml))

def add_p_border_top(p, color="8BC390", sz="6"):
    """Thêm đường kẻ viền trên cho đoạn văn bản."""
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
    - ~subscript~ (CO~2~, H~2~O, H~2~SO~4~)
    - ^superscript^ (W_đ, v^2^, Cu^2+^)
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
    """Xây dựng Header Block không viền, đẹp mắt, chuẩn mực."""
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_l, cell_r = tbl.cell(0, 0), tbl.cell(0, 1)
    cell_l.width = Cm(8.0)
    cell_r.width = Cm(10.0)

    # Cột trái
    p_tl1 = cell_l.paragraphs[0]
    p_tl1.paragraph_format.space_before = Pt(0)
    p_tl1.paragraph_format.space_after = Pt(1)
    p_tl1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tl1, f"*{EXAM_INFO['ubnd']}*", base_size=10, base_bold=True)

    p_tl2 = cell_l.add_paragraph()
    p_tl2.paragraph_format.space_before = Pt(0)
    p_tl2.paragraph_format.space_after = Pt(1)
    p_tl2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tl2, f"*{EXAM_INFO['school']}*", base_size=10.5, base_bold=True)

    p_tl3 = cell_l.add_paragraph()
    p_tl3.paragraph_format.space_before = Pt(0)
    p_tl3.paragraph_format.space_after = Pt(0)
    p_tl3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tl3, f"*{EXAM_INFO['official_label']}*", base_size=10, base_bold=True)

    p_tl4 = cell_l.add_paragraph()
    p_tl4.paragraph_format.space_before = Pt(0)
    p_tl4.paragraph_format.space_after = Pt(0)
    p_tl4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tl4, f"_{EXAM_INFO['pages_label']}_", base_size=9, base_italic=True)

    # Cột phải
    p_tr1 = cell_r.paragraphs[0]
    p_tr1.paragraph_format.space_before = Pt(0)
    p_tr1.paragraph_format.space_after = Pt(1)
    p_tr1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tr1, f"*{EXAM_INFO['exam_title']}*", base_size=11, base_bold=True)

    p_tr2 = cell_r.add_paragraph()
    p_tr2.paragraph_format.space_before = Pt(0)
    p_tr2.paragraph_format.space_after = Pt(1)
    p_tr2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tr2, f"*{EXAM_INFO['year']}*", base_size=10.5, base_bold=True)

    p_tr3 = cell_r.add_paragraph()
    p_tr3.paragraph_format.space_before = Pt(0)
    p_tr3.paragraph_format.space_after = Pt(1)
    p_tr3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tr3, f"*{EXAM_INFO['subject']}*", base_size=11.5, base_bold=True, base_color=COLOR_PRIMARY)

    p_tr4 = cell_r.add_paragraph()
    p_tr4.paragraph_format.space_before = Pt(0)
    p_tr4.paragraph_format.space_after = Pt(0)
    p_tr4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_tr4, f"_{EXAM_INFO['time_str']}_", base_size=9.5, base_italic=True)

    # Dòng tiêu đề phụ cho Giáo viên / Màn hình
    if is_teacher:
        p_gv = doc.add_paragraph()
        p_gv.paragraph_format.space_before = Pt(4)
        p_gv.paragraph_format.space_after = Pt(2)
        p_gv.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rgv = p_gv.add_run("(BẢN GIÁO VIÊN -- CÓ ĐÁP ÁN & LỜI GIẢI CHI TIẾT)")
        set_run_style(rgv, font_size=11.5, bold=True, color=COLOR_CORRECT)
    elif is_screen:
        p_sc = doc.add_paragraph()
        p_sc.paragraph_format.space_before = Pt(4)
        p_sc.paragraph_format.space_after = Pt(2)
        p_sc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rsc = p_sc.add_run("(BẢN TRÌNH CHIẾU / GIẢNG DẠY MÀN HÌNH -- CÓ KHUNG GHI CHÚ BÀI DẠY)")
        set_run_style(rsc, font_size=11.5, bold=True, color=COLOR_SCREEN_TITLE)

    # Dòng thông tin học sinh
    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(4)
    p_info.paragraph_format.space_after = Pt(2)
    render_rich_text(p_info, "*Họ, tên thí sinh:* .................................................... *Lớp:* ......... *Số báo danh:* .........", base_size=base_size)

    # Dòng nguyên tử khối
    p_atm = doc.add_paragraph()
    p_atm.paragraph_format.space_before = Pt(0)
    p_atm.paragraph_format.space_after = Pt(3)
    p_atm.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_atm, f"_({EXAM_INFO['atomic_mass']})_", base_size=9, base_italic=True)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(0)
    p_div.paragraph_format.space_after = Pt(4)
    add_p_border_bottom(p_div, color=COLOR_BORDER_GREY_HEX, sz="4")

def add_section_header(doc, title, subtitle=None, base_size=11):
    """Tiêu đề phần: PHẦN I. TRẮC NGHIỆM KHÁCH QUAN, PHẦN II. TỰ LUẬN, v.v."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    render_rich_text(p, f"*{title}*", base_size=base_size, base_bold=True, base_color=COLOR_PRIMARY)
    if subtitle:
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(2)
        render_rich_text(p_sub, f"_{subtitle}_", base_size=base_size * 0.95, base_italic=True)

def add_solution_block(doc, solution_text, base_size=10):
    """Trình bày Lời giải chi tiết theo chuẩn không dùng Table."""
    p_sol = doc.add_paragraph()
    p_sol.paragraph_format.left_indent = Cm(0.6)
    p_sol.paragraph_format.space_before = Pt(2)
    p_sol.paragraph_format.space_after = Pt(3)
    p_sol.paragraph_format.line_spacing = 1.15

    r_lbl = p_sol.add_run("Lời giải chi tiết:\n")
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
            p_sol.paragraph_format.space_after = Pt(2)
            p_sol.paragraph_format.line_spacing = 1.15
        render_rich_text(p_sol, line_str, base_size=base_size)

def add_screen_teaching_space(doc, num_lines=3):
    """Chừa khoảng trống rộng rãi cho bản Dạy Trên Màn Hình."""
    p_box = doc.add_paragraph()
    p_box.paragraph_format.left_indent = Cm(0.5)
    p_box.paragraph_format.space_before = Pt(3)
    p_box.paragraph_format.space_after = Pt(5)
    
    # Viền nét đứt màu xám nhẹ, viền trái nét liền tím
    pPr = p_box._element.get_or_add_pPr()
    xml = f'<w:pBdr {nsdecls("w")}><w:top w:val="dashed" w:sz="4" w:space="4" w:color="B0BEC5"/><w:bottom w:val="dashed" w:sz="4" w:space="4" w:color="B0BEC5"/><w:left w:val="single" w:sz="12" w:space="6" w:color="7E57C2"/></w:pBdr>'
    pPr.append(parse_xml(xml))
    
    r_hint = p_box.add_run("✎ Ghi chú & Lời giải trực tiếp trên màn hình:\n")
    set_run_style(r_hint, font_size=9.5, bold=True, color=COLOR_SCREEN_TITLE)
    for _ in range(num_lines):
        p_box.add_run("\n")

def add_writing_lines(doc, num_lines=4):
    """Thêm dòng chấm chấm cho học sinh làm bài tự luận."""
    for _ in range(num_lines):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("................................................................................................................................................................................................................")
        set_run_style(r, font_size=9, color=COLOR_GRAY)

def render_mcq_options(doc, opts, correct_letter, is_teacher, base_size=10.5):
    """Bố trí 4 phương án A, B, C, D bằng Tab Stops - Tuyệt đối không dùng Table."""
    lens = [len(opt) for opt in opts]
    max_len = max(lens)

    if max_len <= 18:
        # 4 phương án trên 1 dòng
        p_opt = doc.add_paragraph()
        p_opt.paragraph_format.space_before = Pt(1)
        p_opt.paragraph_format.space_after = Pt(2)
        p_opt.paragraph_format.tab_stops.clear_all()
        p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(0.5), WD_TAB_ALIGNMENT.LEFT)
        p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(5.0), WD_TAB_ALIGNMENT.LEFT)
        p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(9.5), WD_TAB_ALIGNMENT.LEFT)
        p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(14.0), WD_TAB_ALIGNMENT.LEFT)

        for opt in opts:
            p_opt.add_run("\t")
            letter = opt.strip()[0]
            if is_teacher and letter == correct_letter:
                render_rich_text(p_opt, f"*{opt}*", base_size=base_size, base_bold=True, base_color=COLOR_CORRECT)
            else:
                render_rich_text(p_opt, opt, base_size=base_size)

    elif max_len <= 45:
        # 2 dòng x 2 phương án
        p1 = doc.add_paragraph()
        p1.paragraph_format.space_before = Pt(1)
        p1.paragraph_format.space_after = Pt(1)
        p1.paragraph_format.tab_stops.clear_all()
        p1.paragraph_format.tab_stops.add_tab_stop(Cm(0.5), WD_TAB_ALIGNMENT.LEFT)
        p1.paragraph_format.tab_stops.add_tab_stop(Cm(9.2), WD_TAB_ALIGNMENT.LEFT)

        p1.add_run("\t")
        if is_teacher and opts[0].strip()[0] == correct_letter:
            render_rich_text(p1, f"*{opts[0]}*", base_size=base_size, base_bold=True, base_color=COLOR_CORRECT)
        else:
            render_rich_text(p1, opts[0], base_size=base_size)

        p1.add_run("\t")
        if is_teacher and opts[1].strip()[0] == correct_letter:
            render_rich_text(p1, f"*{opts[1]}*", base_size=base_size, base_bold=True, base_color=COLOR_CORRECT)
        else:
            render_rich_text(p1, opts[1], base_size=base_size)

        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(1)
        p2.paragraph_format.space_after = Pt(2)
        p2.paragraph_format.tab_stops.clear_all()
        p2.paragraph_format.tab_stops.add_tab_stop(Cm(0.5), WD_TAB_ALIGNMENT.LEFT)
        p2.paragraph_format.tab_stops.add_tab_stop(Cm(9.2), WD_TAB_ALIGNMENT.LEFT)

        p2.add_run("\t")
        if is_teacher and opts[2].strip()[0] == correct_letter:
            render_rich_text(p2, f"*{opts[2]}*", base_size=base_size, base_bold=True, base_color=COLOR_CORRECT)
        else:
            render_rich_text(p2, opts[2], base_size=base_size)

        p2.add_run("\t")
        if is_teacher and len(opts) > 3 and opts[3].strip()[0] == correct_letter:
            render_rich_text(p2, f"*{opts[3]}*", base_size=base_size, base_bold=True, base_color=COLOR_CORRECT)
        elif len(opts) > 3:
            render_rich_text(p2, opts[3], base_size=base_size)

    else:
        # Mỗi phương án 1 dòng
        for opt in opts:
            p_line = doc.add_paragraph()
            p_line.paragraph_format.left_indent = Cm(0.5)
            p_line.paragraph_format.space_before = Pt(1)
            p_line.paragraph_format.space_after = Pt(1)
            letter = opt.strip()[0]
            if is_teacher and letter == correct_letter:
                render_rich_text(p_line, f"*{opt}*", base_size=base_size, base_bold=True, base_color=COLOR_CORRECT)
            else:
                render_rich_text(p_line, opt, base_size=base_size)

def build_exam_docx(output_path, mode="student"):
    """
    Hàm chính tạo file Word đề kiểm tra
    mode: 'student' | 'teacher' | 'screen'
    """
    is_teacher = (mode == "teacher")
    is_screen = (mode == "screen")
    base_size = 11.5 if is_screen else 10.5

    doc = docx.Document()
    setup_header_footer(doc)
    build_title_block(doc, is_teacher=is_teacher, is_screen=is_screen, base_size=base_size)

    # =========================================================================
    # PHẦN I. TRẮC NGHIỆM KHÁCH QUAN (7,0 ĐIỂM)
    # =========================================================================
    add_section_header(doc, "PHẦN I. TRẮC NGHIỆM KHÁCH QUAN (7,0 điểm)", base_size=base_size + 0.5)

    # 1. Trắc nghiệm nhiều lựa chọn
    p_p1_title = doc.add_paragraph()
    p_p1_title.paragraph_format.space_before = Pt(4)
    p_p1_title.paragraph_format.space_after = Pt(1)
    render_rich_text(p_p1_title, "*1. Trắc nghiệm nhiều lựa chọn (3,0 điểm; câu 1 – câu 12) – Mỗi câu đúng 0,25 điểm*", base_size=base_size, base_bold=True, base_color=COLOR_PRIMARY)
    
    p_p1_sub = doc.add_paragraph()
    p_p1_sub.paragraph_format.space_before = Pt(0)
    p_p1_sub.paragraph_format.space_after = Pt(3)
    render_rich_text(p_p1_sub, "_Đối với mỗi câu trắc nghiệm, học sinh chỉ chọn một phương án trả lời đúng nhất._", base_size=base_size * 0.95, base_italic=True)

    for q in PART1_QUESTIONS:
        qid = q["id"]
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(1)
        p_q.paragraph_format.line_spacing = 1.15
        render_rich_text(p_q, f"*Câu {qid}.* {q['content']}", base_size=base_size)

        render_mcq_options(doc, q["options"], q["correct"], is_teacher, base_size=base_size)

        if is_teacher and "solution" in q:
            add_solution_block(doc, q["solution"], base_size=base_size * 0.95)

        if is_screen:
            add_screen_teaching_space(doc, num_lines=2)

    # 2. Trắc nghiệm Đúng - Sai
    p_p2_title = doc.add_paragraph()
    p_p2_title.paragraph_format.space_before = Pt(6)
    p_p2_title.paragraph_format.space_after = Pt(1)
    render_rich_text(p_p2_title, "*2. Trắc nghiệm Đúng – Sai (2,0 điểm; câu 13 – câu 14) – Mỗi câu 1,0 điểm; mỗi ý đúng 0,25 điểm*", base_size=base_size, base_bold=True, base_color=COLOR_PRIMARY)

    p_p2_sub = doc.add_paragraph()
    p_p2_sub.paragraph_format.space_before = Pt(0)
    p_p2_sub.paragraph_format.space_after = Pt(3)
    render_rich_text(p_p2_sub, "_Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn Đúng hoặc Sai._", base_size=base_size * 0.95, base_italic=True)

    for q in PART2_QUESTIONS:
        qid = q["id"]
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(1)
        p_q.paragraph_format.line_spacing = 1.15
        render_rich_text(p_q, f"*Câu {qid}.* {q['title']}", base_size=base_size)

        for st in q["statements"]:
            p_st = doc.add_paragraph()
            p_st.paragraph_format.left_indent = Cm(0.5)
            p_st.paragraph_format.space_before = Pt(1)
            p_st.paragraph_format.space_after = Pt(1)
            p_st.paragraph_format.line_spacing = 1.15

            lbl = st["label"]
            txt = st["text"]
            if is_teacher:
                ans_str = " (Đúng)" if st["correct"] else " (Sai)"
                ans_color = COLOR_CORRECT if st["correct"] else RGBColor(0, 105, 92)
                render_rich_text(p_st, f"*{lbl})* {txt}", base_size=base_size)
                r_ans = p_st.add_run(ans_str)
                set_run_style(r_ans, font_size=base_size, bold=True, color=ans_color)
            else:
                render_rich_text(p_st, f"*{lbl})* {txt}", base_size=base_size)

        if is_teacher:
            exp_text = "\n".join([f"- Ý {st['label']}: {st['explain']}" for st in q["statements"]])
            add_solution_block(doc, exp_text, base_size=base_size * 0.95)

        if is_screen:
            add_screen_teaching_space(doc, num_lines=3)

    # 3. Trắc nghiệm trả lời ngắn
    p_p3_title = doc.add_paragraph()
    p_p3_title.paragraph_format.space_before = Pt(6)
    p_p3_title.paragraph_format.space_after = Pt(1)
    render_rich_text(p_p3_title, "*3. Trắc nghiệm trả lời ngắn (2,0 điểm; câu 15 – câu 16) – Mỗi câu 1,0 điểm; mỗi ý đúng 0,5 điểm*", base_size=base_size, base_bold=True, base_color=COLOR_PRIMARY)

    p_p3_sub = doc.add_paragraph()
    p_p3_sub.paragraph_format.space_before = Pt(0)
    p_p3_sub.paragraph_format.space_after = Pt(3)
    render_rich_text(p_p3_sub, "_Mỗi ý chỉ ghi đáp án, không cần trình bày lời giải._", base_size=base_size * 0.95, base_italic=True)

    for q in PART3_QUESTIONS:
        qid = q["id"]
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(1)
        render_rich_text(p_q, f"*Câu {qid}.*", base_size=base_size, base_bold=True)

        for sq in q["sub_questions"]:
            p_sq = doc.add_paragraph()
            p_sq.paragraph_format.left_indent = Cm(0.5)
            p_sq.paragraph_format.space_before = Pt(2)
            p_sq.paragraph_format.space_after = Pt(1)
            p_sq.paragraph_format.line_spacing = 1.15
            render_rich_text(p_sq, f"*{sq['sub_id']})* {sq['content']}", base_size=base_size)

            if is_teacher:
                p_ans = doc.add_paragraph()
                p_ans.paragraph_format.left_indent = Cm(0.8)
                p_ans.paragraph_format.space_before = Pt(1)
                p_ans.paragraph_format.space_after = Pt(1)
                r_ans_lbl = p_ans.add_run("Đáp số: ")
                set_run_style(r_ans_lbl, font_size=base_size, bold=True, color=COLOR_CORRECT)
                r_ans_val = p_ans.add_run(sq["correct"])
                set_run_style(r_ans_val, font_size=base_size + 0.5, bold=True, color=COLOR_CORRECT)

                if "solution" in sq:
                    add_solution_block(doc, sq["solution"], base_size=base_size * 0.95)
            else:
                p_ans = doc.add_paragraph()
                p_ans.paragraph_format.left_indent = Cm(0.8)
                p_ans.paragraph_format.space_before = Pt(1)
                p_ans.paragraph_format.space_after = Pt(3)
                r_ans_lbl = p_ans.add_run("Đáp án: ............................................................")
                set_run_style(r_ans_lbl, font_size=base_size, italic=True)

            if is_screen:
                add_screen_teaching_space(doc, num_lines=3)

    # =========================================================================
    # PHẦN II. TỰ LUẬN (3,0 ĐIỂM)
    # =========================================================================
    add_section_header(doc, "PHẦN II. TỰ LUẬN (3,0 điểm)", base_size=base_size + 0.5)

    for q in ESSAY_QUESTIONS:
        qid = q["id"]
        q_score = q["score"]
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_rich_text(p_q, f"*Câu {qid} ({q_score}).*", base_size=base_size, base_bold=True, base_color=COLOR_PRIMARY)

        for sq in q["sub_questions"]:
            p_sq = doc.add_paragraph()
            p_sq.paragraph_format.left_indent = Cm(0.5)
            p_sq.paragraph_format.space_before = Pt(2)
            p_sq.paragraph_format.space_after = Pt(2)
            p_sq.paragraph_format.line_spacing = 1.15
            render_rich_text(p_sq, f"*{sq['sub_id']})* *({sq['score']})* {sq['content']}", base_size=base_size)

            if is_teacher:
                add_solution_block(doc, sq["solution"], base_size=base_size * 0.95)
            else:
                add_writing_lines(doc, num_lines=4)

            if is_screen:
                add_screen_teaching_space(doc, num_lines=4)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(6)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "---------------- *HẾT* ----------------", base_size=base_size, base_bold=True)
    
    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(0)
    p_note.paragraph_format.space_after = Pt(4)
    p_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_note, "_(Giám thị không giải thích gì thêm)_", base_size=base_size * 0.9, base_italic=True)

    # =========================================================================
    # BẢNG TỔNG HỢP ĐÁP ÁN (DÀNH CHO BẢN GIÁO VIÊN)
    # =========================================================================
    if is_teacher:
        p_ans_head = doc.add_paragraph()
        p_ans_head.paragraph_format.space_before = Pt(14)
        p_ans_head.paragraph_format.space_after = Pt(4)
        p_ans_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_ah = p_ans_head.add_run("BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ KIỂM TRA GIỮA KÌ I")
        set_run_style(r_ah, font_size=12, bold=True, color=COLOR_CORRECT)

        # 1. Bảng Phần 1: 12 câu trắc nghiệm nhiều lựa chọn
        p_t1 = doc.add_paragraph()
        p_t1.paragraph_format.space_before = Pt(4)
        p_t1.paragraph_format.space_after = Pt(2)
        render_rich_text(p_t1, "*1. Đáp án Phần I.1 (Trắc nghiệm nhiều lựa chọn - 12 câu):*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

        tbl_p1 = doc.add_table(rows=2, cols=12)
        tbl_p1.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl_p1, color="B0BEC5", sz="4")
        for i in range(12):
            c0 = tbl_p1.cell(0, i); c1 = tbl_p1.cell(1, i)
            c0.width = Cm(1.5); c1.width = Cm(1.5)
            set_cell_margins(c0, 30, 30, 20, 20); set_cell_margins(c1, 30, 30, 20, 20)
            set_cell_shading(c0, "E8F5E9")
            p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r0 = p0.add_run(f"Câu {i+1}")
            set_run_style(r0, font_size=9, bold=True)
            p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p1.add_run(PART1_QUESTIONS[i]["correct"])
            set_run_style(r1, font_size=11, bold=True, color=COLOR_CORRECT)

        # 2. Bảng Phần 2: 2 câu Đúng/Sai
        p_t2 = doc.add_paragraph()
        p_t2.paragraph_format.space_before = Pt(6)
        p_t2.paragraph_format.space_after = Pt(2)
        render_rich_text(p_t2, "*2. Đáp án Phần I.2 (Trắc nghiệm Đúng/Sai - 2 câu):*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

        tbl_p2 = doc.add_table(rows=3, cols=5)
        tbl_p2.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl_p2, color="B0BEC5", sz="4")
        headers_p2 = ["Câu hỏi", "Lệnh hỏi a", "Lệnh hỏi b", "Lệnh hỏi c", "Lệnh hỏi d"]
        for j, h in enumerate(headers_p2):
            cell = tbl_p2.cell(0, j)
            cell.width = Cm(3.6)
            set_cell_margins(cell, 30, 30, 40, 40)
            set_cell_shading(cell, "E8F5E9")
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(h)
            set_run_style(r, font_size=9.5, bold=True)

        for q_idx, q in enumerate(PART2_QUESTIONS):
            row_idx = q_idx + 1
            c_qid = tbl_p2.cell(row_idx, 0)
            c_qid.width = Cm(3.6)
            set_cell_margins(c_qid, 30, 30, 40, 40)
            p = c_qid.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(f"Câu {q['id']}")
            set_run_style(r, font_size=9.5, bold=True)
            for st_idx, st in enumerate(q["statements"]):
                c_st = tbl_p2.cell(row_idx, st_idx + 1)
                c_st.width = Cm(3.6)
                set_cell_margins(c_st, 30, 30, 40, 40)
                p_st = c_st.paragraphs[0]; p_st.alignment = WD_ALIGN_PARAGRAPH.CENTER
                val_str = "Đ" if st["correct"] else "S"
                r_st = p_st.add_run(val_str)
                color_st = COLOR_CORRECT if st["correct"] else RGBColor(0, 105, 92)
                set_run_style(r_st, font_size=10.5, bold=True, color=color_st)

        # 3. Bảng Phần 3: 2 câu Trả lời ngắn (4 ý)
        p_t3 = doc.add_paragraph()
        p_t3.paragraph_format.space_before = Pt(6)
        p_t3.paragraph_format.space_after = Pt(2)
        render_rich_text(p_t3, "*3. Đáp án Phần I.3 (Trắc nghiệm Trả lời ngắn - 4 ý):*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

        tbl_p3 = doc.add_table(rows=2, cols=4)
        tbl_p3.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(tbl_p3, color="B0BEC5", sz="4")
        p3_items = [
            ("Câu 15a", PART3_QUESTIONS[0]["sub_questions"][0]["correct"]),
            ("Câu 15b", PART3_QUESTIONS[0]["sub_questions"][1]["correct"]),
            ("Câu 16a", PART3_QUESTIONS[1]["sub_questions"][0]["correct"]),
            ("Câu 16b", PART3_QUESTIONS[1]["sub_questions"][1]["correct"]),
        ]
        for i, (lbl, ans) in enumerate(p3_items):
            c0 = tbl_p3.cell(0, i); c1 = tbl_p3.cell(1, i)
            c0.width = Cm(4.5); c1.width = Cm(4.5)
            set_cell_margins(c0, 30, 30, 40, 40); set_cell_margins(c1, 30, 30, 40, 40)
            set_cell_shading(c0, "E8F5E9")
            p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r0 = p0.add_run(lbl)
            set_run_style(r0, font_size=9.5, bold=True)
            p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p1.add_run(ans)
            set_run_style(r1, font_size=11, bold=True, color=COLOR_CORRECT)

    # Lưu tài liệu
    doc.save(output_path)
    print(f"Created: {output_path}")

def main():
    files = [
        ("De_Kiem_Tra_KHTN_9_Giua_HK1_Tran_Cao_Van_HocSinh.docx", "student"),
        ("De_Kiem_Tra_KHTN_9_Giua_HK1_Tran_Cao_Van_GiaoVien.docx", "teacher"),
        ("De_Kiem_Tra_KHTN_9_Giua_HK1_Tran_Cao_Van_DayTrenManHinh.docx", "screen"),
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
                    print(f"Copy to OneDrive error ({onedrive_folder}): {e}")

if __name__ == "__main__":
    main()
