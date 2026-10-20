# -*- coding: utf-8 -*-
"""
Script tạo trọn bộ 3 file Word (.docx) Đề kiểm tra / Ôn tập KHTN 9 Giữa HK1 (2026 - 2027)
Theo đúng quy chuẩn "FORM ĐỀ KIỂM TRA" của Thầy TRẦN VĂN THẠNH:
1. Bản Học sinh (De_Kiem_Tra_KHTN_9_Giua_HK1_HocSinh.docx)
2. Bản Giáo viên (De_Kiem_Tra_KHTN_9_Giua_HK1_GiaoVien.docx)
3. Bản Dạy trên màn hình (De_Kiem_Tra_KHTN_9_Giua_HK1_DayTrenManHinh.docx)
"""

import os
import sys
import re
import shutil

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from khtn9_ontap_data import (
    PART_1_QUESTIONS,
    PART_2_QUESTIONS,
    PART_3_QUESTIONS,
    ESSAY_QUESTIONS,
    IMG_DNA
)

ROOT_DIR = r"c:\Antigravity_Thanh"
OUT_DIR = os.path.join(ROOT_DIR, "San_Pham")
os.makedirs(OUT_DIR, exist_ok=True)

ONEDRIVE_DIRS = [
    r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9\KHTN9-DEKIEMTRADINHKI",
    r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9"
]

# Bảng màu quy chuẩn
COLOR_HF_RED = RGBColor(211, 47, 47)        # Coral Red (#D32F2F) cho Header & Footer
COLOR_HF_GREEN_HEX = "8BC390"               # Sage Green (#8BC390) cho viền Header & Footer
COLOR_PRIMARY = RGBColor(46, 125, 50)       # Xanh lá đậm (#2E7D32) cho Tiêu đề Phần
COLOR_CORRECT = RGBColor(198, 40, 40)       # Đỏ (#C62828) cho đáp án đúng Bản Giáo Viên
COLOR_SOLUTION = RGBColor(21, 101, 192)     # Xanh lam (#1565C0) cho tiêu đề Lời giải
COLOR_SCREEN_TITLE = RGBColor(106, 27, 154) # Tím đậm (#6A1B9A) cho tiêu đề bản Dạy Màn Hình
COLOR_GRAY = RGBColor(110, 110, 110)        # Xám
COLOR_LINE_DOT = RGBColor(185, 185, 185)    # Xám nhạt cho nét chấm ..........
COLOR_BORDER_GREY_HEX = "CCCCCC"            # Xám viền phân cách tiêu đề

def add_p_border_bottom(p, color="8BC390", sz="6"):
    """Thêm đường kẻ dưới cho đoạn văn bản (Header, phân cách)."""
    pPr = p._element.get_or_add_pPr()
    xml = f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/></w:pBdr>'
    pPr.append(parse_xml(xml))

def add_p_border_top(p, color="8BC390", sz="6"):
    """Thêm đường kẻ trên cho đoạn văn bản (Footer)."""
    pPr = p._element.get_or_add_pPr()
    xml = f'<w:pBdr {nsdecls("w")}><w:top w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/></w:pBdr>'
    pPr.append(parse_xml(xml))

def set_run_style(run, font_name="Times New Roman", font_size=11, bold=False, italic=False, subscript=False, superscript=False, color=None):
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

def render_rich_text(p, text, base_font="Times New Roman", base_size=11, base_bold=False, base_italic=False, base_color=None):
    """
    Sinh văn bản phong phú từ cú pháp markdown:
    - *text*: in đậm
    - _text_: in nghiêng
    - ~text~: chỉ số dưới (Subscript, ví dụ: H~2~O)
    - ^text^: chỉ số trên (Superscript, ví dụ: mv^2^)
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
    """
    Cài đặt chuẩn trang A4, lề 1.5 cm, Header & Footer theo form quy chuẩn:
    Header: Bảng 1x2, rộng 18.0cm, viền dưới #8BC390.
    Footer: Bảng 1x2, rộng 18.0cm, viền trên #8BC390.
    """
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

def build_title_block(doc, is_teacher=False, is_screen=False, base_size=11):
    """
    Xây dựng khối tiêu đề đầu đề thi dùng Tab Stop phải ở 18.0 cm.
    TUYỆT ĐỐI KHÔNG DÙNG BẢNG để tránh lỗi viền rườm rà.
    """
    headers_data = [
        ("SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ", "HƯỚNG DẪN ÔN TẬP KIỂM TRA GIỮA KÌ I", 10.5, True, 11.5, True),
        ("TRƯỜNG THCS KIM LONG", "MÔN: KHOA HỌC TỰ NHIÊN 9 - NĂM HỌC 2026 - 2027", 10.5, True, 11.5, True),
        ("ĐỀ CHÍNH THỨC", "Thời gian làm bài: 60 phút (không kể thời gian phát đề)", 10.0, False, 10.5, True)
    ]
    for l_text, r_text, l_size, l_bold, r_size, r_bold in headers_data:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.tab_stops.clear_all()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(18.0), WD_TAB_ALIGNMENT.RIGHT)

        rl = p.add_run(l_text)
        set_run_style(rl, font_size=l_size, bold=l_bold, italic=(not l_bold))

        p.add_run('\t')
        rr = p.add_run(r_text)
        set_run_style(rr, font_size=r_size, bold=r_bold, italic=(not r_bold))

    if is_teacher:
        p_gv = doc.add_paragraph()
        p_gv.paragraph_format.space_before = Pt(3)
        p_gv.paragraph_format.space_after = Pt(2)
        p_gv.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rgv = p_gv.add_run("(BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)")
        set_run_style(rgv, font_size=11.5, bold=True, color=COLOR_CORRECT)
    elif is_screen:
        p_sc = doc.add_paragraph()
        p_sc.paragraph_format.space_before = Pt(3)
        p_sc.paragraph_format.space_after = Pt(2)
        p_sc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rsc = p_sc.add_run("(BẢN TRÌNH CHIẾU / GIẢNG DẠY MÀN HÌNH -- CÓ CHỪA CHỖ TRỐNG VIẾT BÀI)")
        set_run_style(rsc, font_size=11.5, bold=True, color=COLOR_SCREEN_TITLE)

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(4)
    p_info.paragraph_format.space_after = Pt(2)
    render_rich_text(p_info, "*Họ, tên thí sinh:* ............................................................ *Lớp:* ............. *Số báo danh:* .............", base_size=base_size - 0.5)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(0)
    p_div.paragraph_format.space_after = Pt(4)
    add_p_border_bottom(p_div, color=COLOR_BORDER_GREY_HEX, sz="4")

def add_major_section_header(doc, title, subtitle=None, base_size=11):
    """Tiêu đề La Mã lớn: I. TRẮC NGHIỆM KHÁCH QUAN / II. TỰ LUẬN."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    render_rich_text(p, f"*{title}*", base_size=base_size + 1.0, base_bold=True, base_color=COLOR_PRIMARY)
    if subtitle:
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(3)
        render_rich_text(p2, f"_{subtitle}_", base_size=base_size - 0.5, base_italic=True)

def add_part_header(doc, title, subtitle=None, base_size=11):
    """Tiêu đề Phần (Phần 1, Phần 2, Phần 3)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    render_rich_text(p, f"*{title}*", base_size=base_size + 0.5, base_bold=True, base_color=COLOR_PRIMARY)
    if subtitle:
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(3)
        render_rich_text(p2, f"_{subtitle}_", base_size=base_size - 0.5, base_italic=True)

def render_single_option(p, opt_str, is_teacher=False, is_correct=False, base_size=11):
    """
    In đậm nhãn A., B., C., D. cho mọi phương án.
    Nếu là bản giáo viên và là phương án đúng: in đậm và tô màu đỏ #C62828.
    """
    clean_str = opt_str.strip()
    m = re.match(r'^([A-D]\.\s*)(.*)$', clean_str)
    if m:
        lbl = m.group(1)
        body = m.group(2)
    else:
        lbl = ""
        body = clean_str

    if is_teacher and is_correct:
        if lbl:
            r_lbl = p.add_run(lbl)
            set_run_style(r_lbl, font_size=base_size, bold=True, color=COLOR_CORRECT)
        render_rich_text(p, body, base_size=base_size, base_bold=True, base_color=COLOR_CORRECT)
    else:
        if lbl:
            r_lbl = p.add_run(lbl)
            set_run_style(r_lbl, font_size=base_size, bold=True, color=None)
        render_rich_text(p, body, base_size=base_size, base_bold=False, base_color=None)

def render_options_no_table(doc, opts, correct_letter, is_teacher=False, base_size=11):
    """
    Trình bày 4 phương án A, B, C, D hoàn toàn KHÔNG DÙNG BẢNG:
    - Chiều dài <= 20 ký tự: 1 dòng (Tab stops: 4.5 cm, 9.0 cm, 13.5 cm).
    - Chiều dài 21 - 42 ký tự: 2 dòng (Tab stop: 9.0 cm).
    - Chiều dài > 42 ký tự: 4 dòng riêng biệt (Left indent: 0.5 cm).
    """
    clean_opts = [re.sub(r'[*_~^]', '', opt) for opt in opts]
    max_len = max(len(o) for o in clean_opts)

    if max_len <= 20:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.tab_stops.clear_all()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(0.5), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(5.0), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(9.5), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(14.0), WD_TAB_ALIGNMENT.LEFT)

        p.add_run('\t')
        for idx, opt in enumerate(opts):
            opt_lbl = opt.strip()[:1]
            is_c = (opt_lbl == correct_letter)
            render_single_option(p, opt, is_teacher=is_teacher, is_correct=is_c, base_size=base_size)
            if idx < 3:
                p.add_run('\t')
    elif max_len <= 42:
        for row_idx in range(2):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1.5)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.tab_stops.clear_all()
            p.paragraph_format.tab_stops.add_tab_stop(Cm(0.5), WD_TAB_ALIGNMENT.LEFT)
            p.paragraph_format.tab_stops.add_tab_stop(Cm(9.2), WD_TAB_ALIGNMENT.LEFT)

            p.add_run('\t')
            opt1 = opts[row_idx * 2]
            is_c1 = (opt1.strip()[:1] == correct_letter)
            render_single_option(p, opt1, is_teacher=is_teacher, is_correct=is_c1, base_size=base_size)

            p.add_run('\t')
            opt2 = opts[row_idx * 2 + 1]
            is_c2 = (opt2.strip()[:1] == correct_letter)
            render_single_option(p, opt2, is_teacher=is_teacher, is_correct=is_c2, base_size=base_size)
    else:
        for opt in opts:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1.5)
            p.paragraph_format.line_spacing = 1.15
            opt_lbl = opt.strip()[:1]
            is_c = (opt_lbl == correct_letter)
            render_single_option(p, opt, is_teacher=is_teacher, is_correct=is_c, base_size=base_size)

def add_solution_block(doc, solution_text, base_size=10.5):
    """
    Trình bày lời giải chi tiết:
    - Tiêu đề '► Lời giải chi tiết:' màu xanh #1565C0, thụt lề 0.4 cm.
    - Thân bài thụt lề 0.8 cm, cỡ 10 pt.
    """
    lines = solution_text.strip().split("\n")
    p_head = doc.add_paragraph()
    p_head.paragraph_format.left_indent = Cm(0.4)
    p_head.paragraph_format.space_before = Pt(3)
    p_head.paragraph_format.space_after = Pt(1)
    render_rich_text(p_head, "► *Lời giải chi tiết:*", base_size=base_size, base_bold=True, base_color=COLOR_SOLUTION)

    for line in lines:
        if line.startswith("*Lời giải chi tiết:*") or line.startswith("►"):
            continue
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.15
        render_rich_text(p, line, base_size=base_size - 0.5)

def add_dotted_lines(doc, num_lines=3, base_size=10.5):
    """
    Thêm các dòng nét chấm .......... cho bản Dạy Trên Màn Hình hoặc làm bài tự luận.
    Đúng yêu cầu: có chừa chỗ trống (có nét ..........)
    """
    for _ in range(num_lines):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1.5)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run("." * 105)
        set_run_style(r, font_size=9.5, color=COLOR_LINE_DOT)

def set_cell_margins(cell, top=50, bottom=50, left=60, right=60):
    """Đặt lề trong cho ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_teacher_answer_key_tables(doc):
    """Tạo bảng tổng hợp đáp án cuối đề cho Bản Giáo Viên."""
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(4)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("--- HẾT ---")
    set_run_style(r_end, font_size=11, bold=True)

    # Tiêu đề bảng đáp án
    p_ah = doc.add_paragraph()
    p_ah.paragraph_format.space_before = Pt(10)
    p_ah.paragraph_format.space_after = Pt(3)
    p_ah.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ah = p_ah.add_run("BẢNG TỔNG HỢP ĐÁP ÁN VÀ BIỂU ĐIỂM")
    set_run_style(r_ah, font_size=12, bold=True, color=COLOR_CORRECT)

    # 1. Bảng Phần I: 24 câu trắc nghiệm (2 bảng con 12 cột)
    p_t1 = doc.add_paragraph()
    p_t1.paragraph_format.space_before = Pt(4)
    p_t1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_t1, "*1. Đáp án Phần 1 (Trắc nghiệm nhiều lựa chọn - 24 câu):*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl_p1 = doc.add_table(rows=4, cols=12)
    tbl_p1.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i in range(12):
        c0 = tbl_p1.cell(0, i); c1 = tbl_p1.cell(1, i)
        c0.width = Cm(1.5); c1.width = Cm(1.5)
        set_cell_margins(c0, 40, 40, 30, 30); set_cell_margins(c1, 40, 40, 30, 30)
        p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(f"C{i+1}")
        set_run_style(r0, font_size=9, bold=True)
        p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(PART_1_QUESTIONS[i]["correct"])
        set_run_style(r1, font_size=10.5, bold=True, color=COLOR_CORRECT)

    for i in range(12):
        c0 = tbl_p1.cell(2, i); c1 = tbl_p1.cell(3, i)
        c0.width = Cm(1.5); c1.width = Cm(1.5)
        set_cell_margins(c0, 40, 40, 30, 30); set_cell_margins(c1, 40, 40, 30, 30)
        p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(f"C{i+13}")
        set_run_style(r0, font_size=9, bold=True)
        p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(PART_1_QUESTIONS[i+12]["correct"])
        set_run_style(r1, font_size=10.5, bold=True, color=COLOR_CORRECT)

    # 2. Bảng Phần II: 4 câu đúng sai
    p_t2 = doc.add_paragraph()
    p_t2.paragraph_format.space_before = Pt(6)
    p_t2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_t2, "*2. Đáp án Phần 2 (Trắc nghiệm đúng / sai - 4 câu):*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl_p2 = doc.add_table(rows=5, cols=5)
    tbl_p2.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers_p2 = ["Câu hỏi", "Lệnh hỏi a", "Lệnh hỏi b", "Lệnh hỏi c", "Lệnh hỏi d"]
    for j, h in enumerate(headers_p2):
        cell = tbl_p2.cell(0, j)
        cell.width = Cm(3.6)
        set_cell_margins(cell, 40, 40, 40, 40)
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_run_style(r, font_size=9.5, bold=True)

    for q_idx, q in enumerate(PART_2_QUESTIONS):
        row_idx = q_idx + 1
        c_qid = tbl_p2.cell(row_idx, 0)
        c_qid.width = Cm(3.6)
        set_cell_margins(c_qid, 40, 40, 40, 40)
        p = c_qid.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"Câu {q['id']}")
        set_run_style(r, font_size=9.5, bold=True)
        for st_idx, st in enumerate(q["statements"]):
            c_st = tbl_p2.cell(row_idx, st_idx + 1)
            c_st.width = Cm(3.6)
            set_cell_margins(c_st, 40, 40, 40, 40)
            p_st = c_st.paragraphs[0]; p_st.alignment = WD_ALIGN_PARAGRAPH.CENTER
            val_str = "Đ" if st["correct"] else "S"
            r_st = p_st.add_run(val_str)
            color_st = COLOR_CORRECT if st["correct"] else RGBColor(0, 105, 92)
            set_run_style(r_st, font_size=10, bold=True, color=color_st)

    # 3. Bảng Phần III: 8 câu trả lời ngắn
    p_t3 = doc.add_paragraph()
    p_t3.paragraph_format.space_before = Pt(6)
    p_t3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_t3, "*3. Đáp án Phần 3 (Trắc nghiệm trả lời ngắn - 8 câu):*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl_p3 = doc.add_table(rows=2, cols=8)
    tbl_p3.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i in range(8):
        c0 = tbl_p3.cell(0, i); c1 = tbl_p3.cell(1, i)
        c0.width = Cm(2.25); c1.width = Cm(2.25)
        set_cell_margins(c0, 40, 40, 30, 30); set_cell_margins(c1, 40, 40, 30, 30)
        p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p0.add_run(f"Câu {i+1}")
        set_run_style(r0, font_size=9, bold=True)
        p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(PART_3_QUESTIONS[i]["correct"])
        set_run_style(r1, font_size=10.5, bold=True, color=COLOR_CORRECT)

def build_exam_docx(output_path, mode="student"):
    """
    Hàm tổng thể tạo file Word chuẩn:
    mode: 'student' (Học sinh) | 'teacher' (Giáo viên) | 'screen' (Dạy trên màn hình)
    """
    is_teacher = (mode == "teacher")
    is_screen = (mode == "screen")
    base_size = 11.5 if is_screen else 10.5

    doc = docx.Document()
    setup_header_footer(doc)
    build_title_block(doc, is_teacher=is_teacher, is_screen=is_screen, base_size=base_size)

    # =========================================================================
    # I. TRẮC NGHIỆM KHÁCH QUAN
    # =========================================================================
    add_major_section_header(
        doc,
        "I. TRẮC NGHIỆM KHÁCH QUAN.",
        base_size=base_size
    )

    # -------------------------------------------------------------------------
    # Phần 1: Trắc nghiệm nhiều lựa chọn (24 câu)
    # -------------------------------------------------------------------------
    add_part_header(
        doc,
        "Phần 1. Trắc nghiệm nhiều lựa chọn.",
        "Mỗi câu hỏi học sinh chỉ chọn một phương án.",
        base_size=base_size
    )

    for q in PART_1_QUESTIONS:
        qid = q["id"]
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3.5)
        p_q.paragraph_format.space_after = Pt(1.5)
        p_q.paragraph_format.line_spacing = 1.15
        p_q.paragraph_format.keep_with_next = True
        render_rich_text(p_q, f"*Câu {qid}.* {q['content']}", base_size=base_size)

        # 4 Phương án
        render_options_no_table(doc, q["options"], q["correct"], is_teacher=is_teacher, base_size=base_size)

        if is_teacher:
            add_solution_block(doc, f"Đáp án đúng: *{q['correct']}*\n" + q["solution"], base_size=base_size)
        elif is_screen:
            # Chừa 2 dòng nét chấm .......... để giáo viên viết nháp / ghi chú trên màn hình
            add_dotted_lines(doc, num_lines=2, base_size=base_size)

    # -------------------------------------------------------------------------
    # Phần 2: Trắc nghiệm chọn đúng - sai (4 câu)
    # -------------------------------------------------------------------------
    add_part_header(
        doc,
        "Phần 2. Trắc nghiệm chọn đúng - sai.",
        "Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai.",
        base_size=base_size
    )

    for q in PART_2_QUESTIONS:
        qid = q["id"]
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        p_q.paragraph_format.keep_with_next = True
        render_rich_text(p_q, f"*Câu {qid}.* {q['content']}", base_size=base_size)

        # Chèn ảnh minh họa nếu có (Câu 3: cấu trúc DNA)
        if q.get("has_image") and os.path.exists(q.get("image_path", "")):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(3)
            run_img = p_img.add_run()
            run_img.add_picture(q["image_path"], width=Cm(6.8))

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
                ans_color = COLOR_CORRECT if st["correct"] else RGBColor(0, 105, 92)
                r_ans = p_st.add_run(f"[{correct_val.upper()}] ")
                set_run_style(r_ans, font_size=base_size, bold=True, color=ans_color)
                render_rich_text(p_st, st_text, base_size=base_size)
            else:
                r_tag = p_st.add_run(f"{sub_id}) ")
                set_run_style(r_tag, font_size=base_size, bold=True)
                render_rich_text(p_st, st_text, base_size=base_size)

        if is_teacher:
            add_solution_block(doc, q["explanation"], base_size=base_size)
        elif is_screen:
            # Chừa 3 dòng nét chấm .......... cho bản dạy màn hình
            add_dotted_lines(doc, num_lines=3, base_size=base_size)

    # -------------------------------------------------------------------------
    # Phần 3: Trắc nghiệm trả lời ngắn (8 câu)
    # -------------------------------------------------------------------------
    add_part_header(
        doc,
        "Phần 3. Trắc nghiệm trả lời ngắn.",
        "Thí sinh viết câu trả lời ngắn hoặc kết quả số tính toán.",
        base_size=base_size
    )

    for q in PART_3_QUESTIONS:
        qid = q["id"]
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3.5)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        p_q.paragraph_format.keep_with_next = True
        render_rich_text(p_q, f"*Câu {qid}.* {q['content']}", base_size=base_size)

        if is_teacher:
            add_solution_block(doc, f"Đáp án: *{q['correct']}*\n" + q["solution"], base_size=base_size)
        elif is_screen:
            # Chừa 3 dòng nét chấm .......... để tính toán trực tiếp trên màn hình
            add_dotted_lines(doc, num_lines=3, base_size=base_size)
        else:
            # Học sinh: có dòng điền đáp án ngắn gọn
            p_ans_box = doc.add_paragraph()
            p_ans_box.paragraph_format.left_indent = Cm(0.5)
            p_ans_box.paragraph_format.space_before = Pt(1)
            p_ans_box.paragraph_format.space_after = Pt(2)
            r_box = p_ans_box.add_run("Trả lời: .............................................................................................................................................")
            set_run_style(r_box, font_size=10, italic=True, color=COLOR_GRAY)

    # =========================================================================
    # II. TỰ LUẬN (8 CÂU)
    # =========================================================================
    add_major_section_header(
        doc,
        "II. TỰ LUẬN",
        base_size=base_size
    )

    for q in ESSAY_QUESTIONS:
        qid = q["id"]
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        p_q.paragraph_format.keep_with_next = True
        render_rich_text(p_q, f"*Câu {qid}.* {q['content']}", base_size=base_size)

        if is_teacher:
            add_solution_block(doc, q["solution"], base_size=base_size)
        elif is_screen:
            # Chừa từ 5 đến 7 dòng nét chấm .......... rộng rãi để dạy viết bài trực tiếp trên màn hình
            num_dots = 7 if qid in [3, 4, 5, 6, 7, 8] else 5
            add_dotted_lines(doc, num_lines=num_dots, base_size=base_size)
        else:
            # Học sinh: chừa 3-4 dòng chấm làm bài
            num_dots = 4 if qid in [3, 4, 5, 6, 7, 8] else 3
            add_dotted_lines(doc, num_lines=num_dots, base_size=base_size)

    # Kết thúc và Bảng đáp án
    if is_teacher:
        add_teacher_answer_key_tables(doc)
    else:
        p_end = doc.add_paragraph()
        p_end.paragraph_format.space_before = Pt(10)
        p_end.paragraph_format.space_after = Pt(4)
        p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_end = p_end.add_run("------------------------ Hết -----------------------")
        set_run_style(r_end, font_size=11, bold=True)

    # Lưu tài liệu
    doc.save(output_path)
    print(f"-> Đã tạo thành công: {output_path}")

def main():
    files = [
        ("De_Kiem_Tra_KHTN_9_Giua_HK1_HocSinh.docx", "student"),
        ("De_Kiem_Tra_KHTN_9_Giua_HK1_GiaoVien.docx", "teacher"),
        ("De_Kiem_Tra_KHTN_9_Giua_HK1_DayTrenManHinh.docx", "screen"),
        # Tạo thêm bộ tên tương ứng theo tiêu đề Hướng dẫn ôn tập để thuận tiện tra cứu
        ("Huong_Dan_On_Tap_KHTN_9_GK1_HocSinh.docx", "student"),
        ("Huong_Dan_On_Tap_KHTN_9_GK1_GiaoVien.docx", "teacher"),
        ("Huong_Dan_On_Tap_KHTN_9_GK1_DayTrenManHinh.docx", "screen"),
    ]

    generated_paths = []
    for filename, mode in files:
        out_file = os.path.join(OUT_DIR, filename)
        build_exam_docx(out_file, mode=mode)
        generated_paths.append((filename, out_file))

    # Đồng bộ sang OneDrive nếu tồn tại thư mục
    for onedrive_folder in ONEDRIVE_DIRS:
        if os.path.exists(onedrive_folder):
            print(f"\nĐang sao chép file sang thư mục OneDrive: {onedrive_folder}")
            for filename, src_path in generated_paths:
                try:
                    dest = os.path.join(onedrive_folder, filename)
                    shutil.copy2(src_path, dest)
                    print(f"  [OneDrive] Đã chép: {filename}")
                except Exception as e:
                    print(f"  [OneDrive] Lỗi khi chép {filename}: {e}")

if __name__ == "__main__":
    main()
