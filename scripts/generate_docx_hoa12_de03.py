# -*- coding: utf-8 -*-
"""
Script tạo file Word (.docx) chuẩn "FORM ĐỀ KIỂM TRA" cho:
ĐỀ KIỂM TRA SỐ 3 - CHƯƠNG 1: ESTER - LIPID (HÓA HỌC 12)
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De03.docx (Bản Học Sinh)
2. De_Kiem_Tra_Hoa_12_De03_GiaoVien.docx (Bản Giáo Viên có lời giải chi tiết)
Tuân thủ 100% quy cách: Header & Footer, Tab Stops (NO TABLE), Sub/Superscript công thức hóa học.
"""

import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
IMG_DIR = r"c:\Antigravity_Thanh\San_Pham\images_crop_de03"

COLOR_HF_RED = RGBColor(211, 47, 47)      # Coral Red (#D32F2F)
COLOR_HF_GREEN_HEX = "8BC390"             # Sage Green (#8BC390)
COLOR_PRIMARY = RGBColor(46, 125, 50)     # Dark Green (#2E7D32)
COLOR_CORRECT = RGBColor(198, 40, 40)     # Red for correct answers (#C62828)
COLOR_SOLUTION = RGBColor(21, 101, 192)   # Blue for solutions (#1565C0)
COLOR_GRAY = RGBColor(97, 97, 97)

def add_p_border_bottom(p, color="8BC390", sz="6"):
    pPr = p._element.get_or_add_pPr()
    xml = f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/></w:pBdr>'
    pPr.append(parse_xml(xml))

def add_p_border_top(p, color="8BC390", sz="6"):
    pPr = p._element.get_or_add_pPr()
    xml = f'<w:pBdr {nsdecls("w")}><w:top w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/></w:pBdr>'
    pPr.append(parse_xml(xml))

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

def render_rich_text(p, text, base_font="Times New Roman", base_size=10.5, base_bold=False, base_italic=False, base_color=None):
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

    # 1. Header (bảng 1 hàng 2 cột, rộng 18.0 cm)
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

    # 2. Footer (bảng 1 hàng 2 cột, rộng 18.0 cm)
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

def build_title_block_no_table(doc, is_teacher=False):
    headers_data = [
        ("SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ", "ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 1 (ESTER - LIPID)", 10, True, 11, True),
        ("ĐỀ CHÍNH THỨC", "MÔN: HOÁ HỌC - LỚP 12", 10.5, True, 11, True),
        ("(Đề thi có 04 trang)", "Thời gian: 50 phút | MÃ ĐỀ THI: 103", 9.5, False, 10, True)
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
        p_gv.paragraph_format.space_before = Pt(2)
        p_gv.paragraph_format.space_after = Pt(2)
        p_gv.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rgv = p_gv.add_run("(BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)")
        set_run_style(rgv, font_size=11, bold=True, color=COLOR_CORRECT)

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(4)
    p_info.paragraph_format.space_after = Pt(2)
    render_rich_text(p_info, "*Họ, tên thí sinh:* ............................................................ *Lớp:* ............. *Số báo danh:* .............", base_size=10.5)

    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(0)
    p_note.paragraph_format.space_after = Pt(4)
    render_rich_text(p_note, "_(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; K = 39; Br = 80; Cl = 35,5)._", base_size=10, base_italic=True)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(0)
    p_div.paragraph_format.space_after = Pt(4)
    add_p_border_bottom(p_div, color="CCCCCC", sz="4")

def add_section_header(doc, title, subtitle=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    render_rich_text(p, f"*{title}*", base_size=11.5, base_bold=True, base_color=COLOR_PRIMARY)
    if subtitle:
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(4)
        render_rich_text(p2, f"_{subtitle}_", base_size=10.5, base_italic=True)

def render_question_prompt(p, text, base_size=10.5):
    m = re.match(r'^(Câu\s+\d+\.)\s*(.*)$', text.strip(), re.DOTALL)
    if m:
        lbl = m.group(1)
        rest = m.group(2)
        r_lbl = p.add_run(lbl + " ")
        set_run_style(r_lbl, font_size=base_size, bold=True)
        render_rich_text(p, rest, base_size=base_size, base_bold=False)
    else:
        render_rich_text(p, text, base_size=base_size, base_bold=False)

def render_single_option(p, opt_str, is_teacher=False, base_size=10.5):
    clean_str = opt_str.strip()
    is_correct = False
    if clean_str.startswith('*') and clean_str.endswith('*'):
        is_correct = True
        clean_str = clean_str[1:-1].strip()

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

def render_options_no_table(doc, opts, is_teacher=False):
    clean_opts = [re.sub(r'[*_~^]', '', opt) for opt in opts]
    max_len = max(len(o) for o in clean_opts)

    if max_len <= 18:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.tab_stops.clear_all()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(4.5), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(13.5), WD_TAB_ALIGNMENT.LEFT)

        for idx, opt in enumerate(opts):
            if idx > 0:
                p.add_run('\t')
            render_single_option(p, opt, is_teacher=is_teacher, base_size=10.5)
    elif max_len <= 38:
        p1 = doc.add_paragraph()
        p1.paragraph_format.space_before = Pt(1)
        p1.paragraph_format.space_after = Pt(1)
        p1.paragraph_format.tab_stops.clear_all()
        p1.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
        render_single_option(p1, opts[0], is_teacher=is_teacher, base_size=10.5)
        p1.add_run('\t')
        render_single_option(p1, opts[1], is_teacher=is_teacher, base_size=10.5)

        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(1)
        p2.paragraph_format.space_after = Pt(2)
        p2.paragraph_format.tab_stops.clear_all()
        p2.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
        render_single_option(p2, opts[2], is_teacher=is_teacher, base_size=10.5)
        p2.add_run('\t')
        render_single_option(p2, opts[3], is_teacher=is_teacher, base_size=10.5)
    else:
        for opt in opts:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            render_single_option(p, opt, is_teacher=is_teacher, base_size=10.5)

def render_statement(p, stmt, ans=None, is_teacher=False, base_size=10.5):
    m = re.match(r'^([a-d]\))\s*(.*)$', stmt.strip())
    if m:
        lbl = m.group(1)
        body = m.group(2)
        r_lbl = p.add_run(lbl + " ")
        set_run_style(r_lbl, font_size=base_size, bold=True)
        render_rich_text(p, body, base_size=base_size, base_bold=False)
    else:
        render_rich_text(p, stmt, base_size=base_size, base_bold=False)

    if is_teacher and ans:
        p.add_run("  ➔ ")
        r_ans = p.add_run(f"({ans})")
        set_run_style(r_ans, font_size=base_size, bold=True, color=COLOR_CORRECT)

def add_solution_no_table(doc, solution_text, sol_img=None, sol_img_width=12):
    lines = solution_text.strip().split("\n")
    p_head = doc.add_paragraph()
    p_head.paragraph_format.left_indent = Cm(0.4)
    p_head.paragraph_format.space_before = Pt(3)
    p_head.paragraph_format.space_after = Pt(1)
    render_rich_text(p_head, "► *Lời giải chi tiết:*", base_size=10.5, base_bold=True, base_color=COLOR_SOLUTION)

    if sol_img and os.path.exists(sol_img):
        p_img = doc.add_paragraph()
        p_img.paragraph_format.left_indent = Cm(0.8)
        p_img.paragraph_format.space_before = Pt(2)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.add_run().add_picture(sol_img, width=Cm(sol_img_width))

    for line in lines:
        if line.startswith("*Lời giải chi tiết:*") or line.startswith("►"):
            continue
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_rich_text(p, line, base_size=10)

def build_exam_de03(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)
    build_title_block_no_table(doc, is_teacher=is_teacher)

    # =========================================================================
    # PHẦN I
    # =========================================================================
    add_section_header(doc, "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
                       "Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")

    part1_questions = [
        {
            "num": 1,
            "q": "Câu 1 (SBT - KNTT). Chất nào sau đây thuộc loại ester?",
            "opts": ["*A. HCOOCH~3~.*", "B. CH~3~COOH.", "C. HCOOH.", "D. CH~3~COCH~3~."],
            "sol": "HCOOCH~3~ (methyl formate) là ester tạo từ formic acid và methanol.\nCH~3~COOH và HCOOH là carboxylic acid; CH~3~COCH~3~ là ketone (acetone)."
        },
        {
            "num": 2,
            "q": "Câu 2 (SBT - KNTT, CTST). Số đồng phân cấu tạo của ester có công thức C~4~H~8~O~2~ là",
            "opts": ["A. 2.", "B. 3.", "*C. 4.*", "D. 5."],
            "sol": "Các đồng phân cấu tạo ester ứng với CTPT C~4~H~8~O~2~ gồm:\n(1) HCOOCH~2~CH~2~CH~3~ (propyl formate)\n(2) HCOOCH(CH~3~)~2~ (isopropyl formate)\n(3) CH~3~COOCH~2~CH~3~ (ethyl acetate)\n(4) CH~3~CH~2~COOCH~3~ (methyl propionate)\n➔ Có tất cả 4 đồng phân cấu tạo ester."
        },
        {
            "num": 3,
            "q": "Câu 3 (SBT - KNTT). Ethyl propanoate là ester tạo nên mùi thơm đặc trưng của quả dứa. Công thức của ethyl propanoate là",
            "opts": ["A. CH~3~COOC~2~H~5~.", "B. C~2~H~5~COOCH~2~CH~2~CH~3~.", "C. CH~3~COOCH~2~CH~2~CH~3~.", "*D. C~2~H~5~COOC~2~H~5~.*"],
            "sol": "Ethyl propanoate được tạo thành từ propanoic acid (C~2~H~5~COOH) và alcohol ethylic (C~2~H~5~OH) nên có công thức cấu tạo là C~2~H~5~COOC~2~H~5~."
        },
        {
            "num": 4,
            "q": "Câu 4. Ester nào sau đây tác dụng với dung dịch NaOH thu được sodium formate?",
            "opts": ["A. C~2~H~5~COOC~2~H~5~.", "B. CH~3~COOC~2~H~5~.", "C. CH~3~COOCH~3~.", "*D. HCOOCH~3~.*"],
            "sol": "Sodium formate là HCOONa. Để thuỷ phân trong kiềm thu được HCOONa thì ester phải có dạng HCOOR'.\nPhương trình phản ứng: HCOOCH~3~ + NaOH ➔ HCOONa + CH~3~OH."
        },
        {
            "num": 5,
            "q": "Câu 5. Khi đun nóng chất X có công thức phân tử C~3~H~6~O~2~ với dung dịch NaOH thu được CH~3~COONa. Công thức cấu tạo của X là",
            "opts": ["A. CH~3~COOC~2~H~5~.", "B. HCOOC~2~H~5~.", "*C. CH~3~COOCH~3~.*", "D. C~2~H~5~COOH."],
            "sol": "X có CTPT C~3~H~6~O~2~ tác dụng với NaOH sinh ra muối CH~3~COONa nên X là ester của acetic acid (CH~3~COOR'). Vì có 3 nguyên tử C nên R' là nhóm methyl (-CH~3~).\nVậy công thức cấu tạo của X là CH~3~COOCH~3~ (methyl acetate)."
        },
        {
            "num": 6,
            "q": "Câu 6. Chất nào sau đây phản ứng với dung dịch NaOH tạo thành HCOONa và C~2~H~5~OH?",
            "opts": ["*A. HCOOC~2~H~5~.*", "B. CH~3~COOC~2~H~5~.", "C. CH~3~COOCH~3~.", "D. HCOOCH~3~."],
            "sol": "Phương trình phản ứng thuỷ phân:\nHCOOC~2~H~5~ + NaOH ➔ HCOONa + C~2~H~5~OH.\nChất thoả mãn là ethyl formate (HCOOC~2~H~5~)."
        },
        {
            "num": 7,
            "q": "Câu 7 (SBT - CTST). Tính chất nào sau đây không phải là tính chất thích hợp giúp ethyl methanoate (HCOOC~2~H~5~) được sử dụng trong sản xuất một số loại nước hoa?",
            "opts": ["*A. Khả năng dễ cháy.*", "B. Có mùi thơm dễ chịu.", "C. Không độc hại.", "D. Nhiệt độ sôi thấp."],
            "sol": "Ethyl methanoate được dùng trong công nghiệp nước hoa nhờ có mùi thơm dễ chịu, không độc hại và dễ bay hơi (nhiệt độ sôi thấp). Khả năng dễ cháy là đặc tính nguy hiểm cần lưu ý an toàn chứ không phải là tính chất thích hợp giúp nó được chọn ứng dụng làm nước hoa."
        },
        {
            "num": 8,
            "q": "Câu 8. Nguyên liệu nào sau đây là nguyên liệu vô cơ chủ yếu trong phản ứng xà phòng hoá?",
            "opts": ["A. Calcium hydroxide Ca(OH)~2~.", "*B. Sodium hydroxide NaOH.*", "C. Barium hydroxide Ba(OH)~2~.", "D. Iron(III) hydroxide Fe(OH)~3~."],
            "sol": "Trong công nghiệp sản xuất xà phòng, người ta thường dùng dung dịch base mạnh kiềm vô cơ như NaOH (để làm xà phòng bánh cứng) hoặc KOH (để làm xà phòng lỏng)."
        },
        {
            "num": 9,
            "q": "Câu 9. Xà phòng được tổng hợp theo phản ứng:",
            "img": os.path.join(IMG_DIR, "cau9_phan_ung_xa_phong_hoa.png"),
            "img_width": 12.0,
            "q_sub": "Phản ứng trên là ví dụ của",
            "opts": ["A. phản ứng ngưng tụ.", "B. phản ứng ester hóa.", "C. phản ứng oxi hóa.", "*D. phản ứng thủy phân.*"],
            "sol": "Phản ứng trên là phản ứng thuỷ phân chất béo (tristearin) trong môi trường kiềm (dung dịch NaOH), còn gọi là phản ứng xà phòng hoá."
        },
        {
            "num": 10,
            "q": "Câu 10. Phát biểu nào sau đây không đúng?",
            "opts": [
                "A. Chất béo không tan trong nước, nhẹ hơn nước nhưng tan nhiều trong dung môi hữu cơ.",
                "*B. Dầu ăn và mỡ bôi trơn có cùng thành phần nguyên tố.*",
                "C. Chất béo là triester của glycerol với các monocarboxylic acid có số chẵn nguyên tử carbon (12 - 24), mạch carbon dài, không phân nhánh.",
                "D. Phản ứng thủy phân chất béo trong môi trường base là phản ứng một chiều."
            ],
            "sol": "Dầu ăn là chất béo (chứa các nguyên tố C, H, O), còn mỡ bôi trơn máy móc là hỗn hợp các hydrocarbon (chỉ chứa các nguyên tố C, H). Do đó thành phần nguyên tố của chúng khác nhau ➔ Phát biểu B không đúng."
        },
        {
            "num": 11,
            "q": "Câu 11. Nguyên nhân nào làm cho bồ kết có khả năng giặt rửa:",
            "opts": [
                "*A. Vì bồ kết có những chất có cấu tạo kiểu đầu phân cực gắn với đuôi không phân cực.*",
                "B. Vì bồ kết có thành phần là este của glycerine.",
                "C. Vì trong bồ kết có những chất oxi hóa mạnh.",
                "D. Vì trong bồ kết có chất khử mạnh."
            ],
            "sol": "Quả bồ kết chứa saponin (chất hoạt động bề mặt tự nhiên). Phân tử có cấu tạo gồm một đầu ưa nước (phân cực) và một đuôi kị nước/ưa dầu (không phân cực), do đó có khả năng nhũ hoá dầu mỡ và tẩy rửa tương tự xà phòng."
        },
        {
            "num": 12,
            "q": "Câu 12. Cho phản ứng ester hóa: RCOOH + R'OH ⇌ RCOOR' + H~2~O. Để phản ứng chuyển dịch ưu tiên theo chiều thuận người ta thường:",
            "opts": [
                "A. Dùng H~2~SO~4~ đặc để hút nước và làm xúc tác.",
                "B. Chưng cất để tách ester ra khỏi hỗn hợp phản ứng.",
                "C. Tăng nồng độ của acid hoặc alcohol.",
                "*D. Tất cả đều đúng.*"
            ],
            "sol": "Theo nguyên lí chuyển dịch cân bằng Le Chatelier:\n• H~2~SO~4~ đặc hút nước làm giảm nồng độ H~2~O ➔ cân bằng chuyển dịch theo chiều thuận.\n• Chưng cất tách ester làm giảm nồng độ sản phẩm ➔ cân bằng chuyển dịch theo chiều thuận.\n• Tăng nồng độ chất phản ứng (acid hoặc alcohol) ➔ cân bằng dịch chuyển theo chiều thuận.\n➔ Tất cả các biện pháp đều đúng."
        },
        {
            "num": 13,
            "q": "Câu 13 (SBT-CD). Phát biểu nào sau đây là không đúng?",
            "opts": [
                "A. So với các đồng phân là carboxylic acid, ester luôn có nhiệt độ sôi thấp hơn.",
                "B. Phản ứng ester hoá là phản ứng thuận nghịch.",
                "*C. Phản ứng xà phòng hoá là phản ứng thuận nghịch.*",
                "D. Ester là những chất lỏng hoặc chất rắn ở nhiệt độ thường."
            ],
            "sol": "Phản ứng xà phòng hoá (thuỷ phân ester trong môi trường kiềm) là phản ứng một chiều (không thuận nghịch) vì anion carboxylate tạo thành không có khả năng phản ứng ngược lại với alcohol ở điều kiện thí nghiệm ➔ Phát biểu C không đúng."
        },
        {
            "num": 14,
            "q": "Câu 14. Thủy phân ester có công thức C~4~H~8~O~2~ (với xúc tác acid), thu được hai sản phẩm hữu cơ X và Y. Từ X có thể điều chế trực tiếp ra Y. Vậy chất X là:",
            "opts": ["A. Methyl alcohol.", "B. Ethyl acetate.", "C. Formic acid.", "*D. Ethyl alcohol.*"],
            "sol": "Thuỷ phân ester C~4~H~8~O~2~ trong acid thu được acid và alcohol. Phản ứng:\nCH~3~COOC~2~H~5~ + H~2~O ⇌ CH~3~COOH (Y) + C~2~H~5~OH (X).\nTừ X (C~2~H~5~OH) có thể lên men giấm trực tiếp tạo ra Y (CH~3~COOH):\nC~2~H~5~OH + O~2~ ➔ (men giấm) CH~3~COOH + H~2~O.\nVậy X là ethyl alcohol."
        },
        {
            "num": 15,
            "q": "Câu 15 (SBT-CD). Thuỷ phân ester E có công thức phân tử C~4~H~8~O~2~ với xúc tác acid vô cơ loãng, thu được hai sản phẩm hữu cơ X, Y (chỉ chứa các nguyên tử C, H, O). Từ X có thể điều chế trực tiếp ra Y bằng một phản ứng duy nhất. Chất X là chất nào sau đây?",
            "opts": ["A. Acetic acid.", "*B. Ethyl alcohol.*", "C. Ethyl acetate.", "D. Formic acid."],
            "sol": "Ester E là CH~3~COOC~2~H~5~ (ethyl acetate). Thuỷ phân:\nCH~3~COOC~2~H~5~ + H~2~O ⇌ C~2~H~5~OH (X) + CH~3~COOH (Y).\nPhản ứng oxi hoá ethyl alcohol tạo acetic acid:\nC~2~H~5~OH + O~2~ ➔ (men giấm) CH~3~COOH + H~2~O.\nVậy chất X là ethyl alcohol."
        },
        {
            "num": 16,
            "q": "Câu 16 (SGK - KNTT). Cho các phát biểu sau:\n(1) Một số ester có mùi thơm nên được dùng làm chất tạo hương trong công nghiệp thực phẩm và mĩ phẩm.\n(2) Chất béo là triester của glycerol với acid béo.\n(3) Chất béo tan tốt trong nước.\n(4) Mỡ động vật, dầu thực vật có thể được dùng làm nguyên liệu để sản xuất xà phòng.\n(5) Phản ứng thuỷ phân ester trong môi trường acid luôn là phản ứng một chiều.\nSố phát biểu đúng là",
            "opts": ["A. 2.", "*B. 3.*", "C. 4.", "D. 5."],
            "sol": "• (1) Đúng: Một số ester có mùi thơm hoa quả dùng làm hương liệu thực phẩm, mĩ phẩm.\n• (2) Đúng: Định nghĩa chất béo là triester của glycerol với các acid béo.\n• (3) Sai: Chất béo không tan trong nước, nhẹ hơn nước.\n• (4) Đúng: Dầu mỡ động thực vật là nguyên liệu chính để sản xuất xà phòng.\n• (5) Sai: Thuỷ phân ester trong môi trường acid thông thường là phản ứng thuận nghịch.\n➔ Có 3 phát biểu đúng: (1), (2), (4)."
        },
        {
            "num": 17,
            "q": "Câu 17 (SBT-CD). Tiến hành thí nghiệm điều chế isoamyl acetate (chất có mùi chuối chín) theo thứ tự các bước sau đây:\nBước 1: Cho khoảng 3 mL CH~3~CH(CH~3~)CH~2~CH~2~OH, 3 mL CH~3~COOH và vài giọt H~2~SO~4~ đặc vào ống nghiệm.\nBước 2: Lắc đều ống nghiệm rồi đun cách thuỷ (trong nồi nước nóng) khoảng 5 – 7 phút ở 65 – 70 °C.\nBước 3: Làm lạnh, sau đó thêm khoảng 5 mL dung dịch NaCl bão hoà vào ống nghiệm.\nCho các phát biểu sau:\n(a) H~2~SO~4~ đặc có vai trò xúc tác cho phản ứng tạo isoamyl acetate.\n(b) Thêm dung dịch NaCl bão hoà vào để tránh phân huỷ sản phẩm.\n(c) Sau bước 2, trong ống nghiệm vẫn còn CH~3~CH(CH~3~)CH~2~CH~2~OH và CH~3~COOH.\n(d) Sau bước 3, trong ống nghiệm thu được hỗn hợp chất lỏng đồng nhất.\nSố phát biểu đúng là",
            "opts": ["A. 4.", "B. 5.", "C. 3.", "*D. 2.*"],
            "sol": "• (a) Đúng: H~2~SO~4~ đặc làm xúc tác và hút nước để cân bằng dịch chuyển theo chiều thuận.\n• (b) Sai: Thêm dung dịch NaCl bão hoà để làm tăng khối lượng riêng của dung dịch và làm giảm độ tan của ester, giúp ester dễ tách lớp nổi lên trên.\n• (c) Đúng: Phản ứng ester hoá là phản ứng thuận nghịch nên sau bước 2 các chất phản ứng ban đầu vẫn còn.\n• (d) Sai: Isoamyl acetate không tan trong dung dịch NaCl và nhẹ hơn nên nổi lên trên mặt nước, tạo thành 2 lớp chất lỏng phân biệt (không đồng nhất).\n➔ Có 2 phát biểu đúng: (a), (c)."
        },
        {
            "num": 18,
            "q": "Câu 18. Một loại chất béo có chứa 65% tripalmitin về khối lượng. Để sản xuất 3 triệu hộp kem dưỡng da (có chứa chất dưỡng ẩm glycerol) cần dùng tối thiểu x tấn loại chất béo trên cho phản ứng với dung dịch NaOH, đun nóng. Biết hiệu suất phản ứng đạt 80%. Biết rằng trong mỗi hộp kem dưỡng da có chứa 5 gam glycerol. Giá trị của x là",
            "opts": ["A. 152,7.", "B. 525,4.", "*C. 252,7.*", "D. 175,7."],
            "sol": "• Tổng lượng glycerol cần cho 3 triệu hộp kem dưỡng da:\n  m~glycerol~ = 3 × 10^6 × 5 g = 15 × 10^6 g = 15 tấn.\n• Phản ứng xà phòng hoá tripalmitin:\n  (C~15~H~31~COO)~3~C~3~H~5~ + 3NaOH ➔ 3C~15~H~31~COONa + C~3~H~5~(OH)~3~\n  806 tấn ............................................................................ 92 tấn\n• Lượng tripalmitin lí thuyết cần:\n  m~tripalmitin (LT)~ = 15 × 806 / 92 ≈ 131,413 tấn.\n• Do hiệu suất phản ứng H = 80%:\n  m~tripalmitin (TT)~ = 131,413 / 0,8 = 164,266 tấn.\n• Do chất béo chỉ chứa 65% tripalmitin về khối lượng:\n  x = 164,266 / 0,65 ≈ 252,717 tấn ≈ 252,7 tấn.\n➔ Chọn C."
        }
    ]

    for q_data in part1_questions:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, q_data["q"])

        if "img" in q_data:
            p_img = doc.add_paragraph()
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(2)
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run().add_picture(q_data["img"], width=Cm(q_data.get("img_width", 12)))
            if "q_sub" in q_data:
                p_sub = doc.add_paragraph()
                p_sub.paragraph_format.space_before = Pt(1)
                p_sub.paragraph_format.space_after = Pt(1)
                render_rich_text(p_sub, q_data["q_sub"])

        render_options_no_table(doc, q_data["opts"], is_teacher=is_teacher)

        if is_teacher and "sol" in q_data:
            add_solution_no_table(doc, q_data["sol"])

    # =========================================================================
    # PHẦN II
    # =========================================================================
    add_section_header(doc, "PHẦN II. Câu trắc nghiệm đúng sai.",
                       "Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")

    part2_questions = [
        {
            "num": 1,
            "prompt": "Câu 1. Nhiều loại bơ thực vật (chất béo no ở dạng rắn) được tạo ra bởi quá trình hydrogen hóa một phần dầu thực vật (chất béo không no ở dạng lỏng).",
            "img": os.path.join(IMG_DIR, "cau1_hydrogen_hoa_dau_an.png"),
            "img_width": 11.0,
            "stmts": [
                ("a) Trong công nghiệp, người ta sử dụng phản ứng này để chuyển hóa chất béo lỏng thành chất béo rắn và ngược lại.", "Sai",
                 "Quá trình hydrogen hoá chỉ chuyển hoá chất béo lỏng không no thành chất béo rắn no; không có phản ứng ngược lại để biến chất béo rắn thành chất béo lỏng."),
                ("b) 1 mol triolein (chất béo dạng lỏng) phản ứng tối đa với 3 mol H~2~ (Ni, t°) tạo thành tristearin (chất béo dạng rắn).", "Đúng",
                 "Phương trình: (C~17~H~33~COO)~3~C~3~H~5~ + 3H~2~ ➔ (Ni, t°) (C~17~H~35~COO)~3~C~3~H~5~. 1 mol triolein tác dụng với 3 mol H~2~ tạo 1 mol tristearin."),
                ("c) Do chứa các liên kết đôi (C=C) trong phân tử, nên chất béo không no bị oxi hóa chậm bởi oxygen trong không khí tạo ra các chất có mùi khó chịu, làm cho dầu mỡ bị ôi.", "Đúng",
                 "Liên kết C=C trong gốc acid béo không no bị oxygen không khí oxi hoá chậm tạo thành peroxide, sau đó phân huỷ thành các hợp chất aldehyde, ketone có mùi ôi khó chịu."),
                ("d) Số nguyên tử carbon ở chất béo lỏng và rắn khác nhau.", "Sai",
                 "Ví dụ triolein (chất béo lỏng) và tristearin (chất béo rắn) đều có cùng số nguyên tử carbon trong phân tử là 57 carbon.")
            ]
        },
        {
            "num": 2,
            "prompt": "Câu 2. Aspirin được sử dụng làm thuốc giảm đau, hạ sốt. Sau khi uống, aspirin bị thuỷ phân trong cơ thể tạo thành salicylic acid. Salicylic acid ức chế quá trình sinh tổng hợp prostaglandin (chất gây đau, sốt và viêm khi nồng độ trong máu cao hơn mức bình thường).",
            "img": os.path.join(IMG_DIR, "cau2_thuy_phan_aspirin.png"),
            "img_width": 12.0,
            "stmts": [
                ("a) Aspirin có công thức phân tử C~9~H~8~O~4~.", "Đúng",
                 "Aspirin (acetylsalicylic acid: CH~3~COOC~6~H~4~COOH) có công thức phân tử là C~9~H~8~O~4~."),
                ("b) Salicylic acid có công thức phân tử C~7~H~6~O~3~.", "Đúng",
                 "Salicylic acid (2-hydroxybenzoic acid: HO-C~6~H~4~-COOH) có công thức phân tử là C~7~H~6~O~3~."),
                ("c) Aspirin có độ bất bão hòa k = 5.", "Sai",
                 "Độ bất bão hoà của aspirin: k = (2×9 + 2 - 8) / 2 = 6 (gồm 1 vòng benzen k = 4, 1 nhóm -COO- có k = 1, và 1 nhóm -COOH có k = 1)."),
                ("d) Salicylic acid là hợp chất hữu cơ đa chức.", "Sai",
                 "Salicylic acid chứa đồng thời nhóm chức hydroxy phenol (-OH) và nhóm carboxyl (-COOH) khác loại nhau nên là hợp chất hữu cơ tạp chức, không phải đa chức.")
            ]
        },
        {
            "num": 3,
            "prompt": "Câu 3 (SBT - CTST). Cho các chất sau:",
            "img": os.path.join(IMG_DIR, "cau3_sodium_palmitate_laurylsulfate.png"),
            "img_width": 13.5,
            "post_img": "Xét tính đúng/sai của các phát biểu sau:",
            "stmts": [
                ("a) Sodium palmitate và sodium laurylsulfate đều là thành phần chính của xà phòng.", "Sai",
                 "Sodium palmitate là muối của acid béo nên là thành phần chính của xà phòng; còn sodium laurylsulfate là muối alkyl sulfate – thành phần chính của chất giặt rửa tổng hợp."),
                ("b) Sodium palmitate và sodium laurylsulfate đều có tính năng giặt rửa.", "Đúng",
                 "Cả hai chất đều là chất hoạt động bề mặt có cấu tạo đầu ưa nước gắn đuôi kị nước nên đều có khả năng nhũ hoá dầu mỡ và tính năng giặt rửa tốt."),
                ("c) Sodium palmitate và sodium laurylsulfate đều tạo muối khó tan trong nước cứng.", "Sai",
                 "Sodium palmitate tạo kết tủa khó tan với Ca^2+^, Mg^2+^ trong nước cứng; nhưng chất giặt rửa tổng hợp sodium laurylsulfate không tạo kết tủa trong nước cứng nên vẫn tẩy rửa bình thường."),
                ("d) Sodium palmitate và sodium laurylsulfate đều có đầu ưa nước gắn với đuôi kị nước.", "Đúng",
                 "Cấu trúc phân tử của cả hai chất đều gồm chuỗi hydrocarbon dài kị nước liên kết với nhóm phân cực mang điện tích ưa nước (-COO^- hoặc -OSO~3~^-).")
            ]
        },
        {
            "num": 4,
            "prompt": "Câu 4. Các chất béo có thể là chất lỏng hoặc rắn, có trong một số loại động thực vật.",
            "stmts": [
                ("a) Tripalmitin là một loại chất béo không no có trong mỡ động vật và dầu cọ có công thức phân tử là C~51~H~98~O~6~.", "Sai",
                 "Tripalmitin (C~15~H~31~COO)~3~C~3~H~5~ có CTPT là C~51~H~98~O~6~ nhưng là chất béo NO, không phải chất béo không no."),
                ("b) Tristearin thu được từ mỡ động vật là một sản phẩm phụ của chế biến thịt bò có công thức phân tử là C~57~H~110~O~6~.", "Đúng",
                 "Tristearin có công thức (C~17~H~35~COO)~3~C~3~H~5~ = C~57~H~110~O~6~, là chất béo no thể rắn trong mỡ bò, mỡ cừu."),
                ("c) Triolein chiếm 4-30% trong dầu oliu là chất béo không no ở trạng thái lỏng có công thức là C~57~H~104~O~6~.", "Đúng",
                 "Triolein có công thức (C~17~H~33~COO)~3~C~3~H~5~ = C~57~H~104~O~6~, chứa 3 liên kết đôi C=C nên là chất béo không no ở trạng thái lỏng."),
                ("d) Glyceryl trilinoleate hay trilinolein là chất béo no ở trạng thái rắn có công thức phân tử là C~57~H~98~O~6~.", "Sai",
                 "Linoleic acid (C~17~H~31~COOH) có 2 nối đôi C=C nên trilinolein (C~17~H~31~COO)~3~C~3~H~5~ = C~57~H~98~O~6~ là chất béo KHÔNG NO ở trạng thái lỏng, không phải chất béo no thể rắn.")
            ]
        }
    ]

    for p2_data in part2_questions:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, p2_data["prompt"])

        if "img" in p2_data:
            p_img = doc.add_paragraph()
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(2)
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run().add_picture(p2_data["img"], width=Cm(p2_data.get("img_width", 11)))

        if "post_img" in p2_data:
            p_post = doc.add_paragraph()
            p_post.paragraph_format.space_before = Pt(1)
            p_post.paragraph_format.space_after = Pt(2)
            render_rich_text(p_post, p2_data["post_img"], base_italic=True)

        for stmt, ans, sol in p2_data["stmts"]:
            p_stmt = doc.add_paragraph()
            p_stmt.paragraph_format.left_indent = Cm(0.5)
            p_stmt.paragraph_format.space_before = Pt(1)
            p_stmt.paragraph_format.space_after = Pt(1)
            render_statement(p_stmt, stmt, ans=ans, is_teacher=is_teacher)

        if is_teacher:
            sol_full = "\n".join([f"• *{s[:2]}* là *{a}*: {sol}" for s, a, sol in p2_data["stmts"]])
            add_solution_no_table(doc, sol_full)

    # =========================================================================
    # PHẦN III
    # =========================================================================
    add_section_header(doc, "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
                       "Thí sinh trả lời từ câu 1 đến câu 6.")

    part3_questions = [
        {
            "num": 1,
            "q": "Câu 1 (SBT - CTST). Có bao nhiêu ester của methyl alcohol có công thức phân tử là C~5~H~10~O~2~?",
            "ans": "2",
            "sol": "Ester của methyl alcohol có dạng RCOOCH~3~.\nVới CTPT C~5~H~10~O~2~, gốc R là nhóm alkyl C~3~H~7~-.\nNhóm C~3~H~7~- có 2 đồng phân cấu tạo là propyl (CH~3~CH~2~CH~2~-) và isopropyl ((CH~3~)~2~CH-).\n➔ Có 2 ester thoả mãn:\n(1) CH~3~CH~2~CH~2~COOCH~3~ (methyl butyrate)\n(2) (CH~3~)~2~CHCOOCH~3~ (methyl isobutyrate)\n➔ Đáp án: 2."
        },
        {
            "num": 2,
            "q": "Câu 2 (SBT - CTST). Có bao nhiêu triglyceride mà khi xà phòng hoá hoàn toàn thu được glycerol cùng hỗn hợp chỉ gồm muối của palmitic acid và stearic acid?",
            "ans": "4",
            "sol_img": os.path.join(IMG_DIR, "p3_cau2_dong_phan_triglyceride.png"),
            "sol_img_width": 13.0,
            "sol": "Kí hiệu gốc palmitate (C~15~H~31~COO-) là P và gốc stearate (C~17~H~35~COO-) là S.\nĐể thu được cả 2 muối P và S thì triglyceride phải chứa cả hai loại gốc:\n• Trường hợp 1: Chứa 2 gốc P và 1 gốc S:\n  - Gốc S ở giữa (vị trí β): P-S-P (1 đồng phân)\n  - Gốc S ở đầu (vị trí α): S-P-P (1 đồng phân)\n• Trường hợp 2: Chứa 1 gốc P và 2 gốc S:\n  - Gốc P ở giữa (vị trí β): S-P-S (1 đồng phân)\n  - Gốc P ở đầu (vị trí α): P-S-S (1 đồng phân)\n➔ Tổng cộng có 1 + 1 + 1 + 1 = 4 triglyceride thoả mãn.\n➔ Đáp án: 4."
        },
        {
            "num": 3,
            "q": "Câu 3 (SBT - KNTT). Xà phòng hoá hoàn toàn 4,4 g ethyl acetate bằng dung dịch NaOH dư, thu được bao nhiêu gam muối sodium acetate?",
            "ans": "4,1",
            "sol": "• Khối lượng mol của ethyl acetate (CH~3~COOC~2~H~5~): M = 88 g/mol.\n  n~CH3COOC2H5~ = 4,4 / 88 = 0,05 mol.\n• Phương trình phản ứng:\n  CH~3~COOC~2~H~5~ + NaOH ➔ CH~3~COONa + C~2~H~5~OH\n  0,05 mol --------------------------➔ 0,05 mol\n• Khối lượng muối sodium acetate (M = 82 g/mol) thu được:\n  m~CH3COONa~ = 0,05 × 82 = 4,1 gam.\n➔ Đáp án: 4,1."
        },
        {
            "num": 4,
            "q": "Câu 4 (SBT - KNTT). Thực hiện phản ứng xà phòng hoá 100 kg chất béo chứa 80% tristearin, còn lại là các tạp chất không phản ứng, thu được bao nhiêu kilogam sodium stearate? (Biết hiệu suất phản ứng xà phòng hoá là 90%.) (Làm tròn kết quả đến hàng phần mười).",
            "ans": "74,3",
            "sol": "• Khối lượng tristearin có trong 100 kg chất béo:\n  m~tristearin~ = 100 × 80% = 80 kg = 80000 gam.\n• Khối lượng mol của tristearin (C~17~H~35~COO)~3~C~3~H~5~ là M = 890 g/mol.\n  n~tristearin~ = 80000 / 890 ≈ 89,888 mol.\n• Phương trình phản ứng:\n  (C~17~H~35~COO)~3~C~3~H~5~ + 3NaOH ➔ 3C~17~H~35~COONa + C~3~H~5~(OH)~3~\n  n~muối (LT)~ = 3 × n~tristearin~ = 3 × 89,888 ≈ 269,663 mol.\n• Khối lượng mol của sodium stearate (C~17~H~35~COONa) là M = 306 g/mol.\n• Do hiệu suất phản ứng là 90%, khối lượng muối thực tế thu được là:\n  m~muối~ = 269,663 × 306 × 90% ≈ 74265 gam ≈ 74,3 kg.\n➔ Đáp án: 74,3."
        },
        {
            "num": 5,
            "q": "Câu 5. Một loại chất béo có chứa 70% triolein về khối lượng. Xà phòng hóa hoàn toàn 24 kg chất béo này trong dung dịch NaOH, đun nóng thu được x bánh xà phòng. Biết rằng trong mỗi bánh xà phòng có chứa 60 gam sodium oleate. Giá trị của x là (lấy phần nguyên). Biết hiệu suất phản ứng đạt 90%.",
            "ans": "260",
            "sol": "• Khối lượng triolein trong 24 kg chất béo:\n  m~triolein~ = 24 × 70% = 16,8 kg = 16800 gam.\n• Khối lượng mol triolein (C~17~H~33~COO)~3~C~3~H~5~: M = 884 g/mol.\n  n~triolein~ = 16800 / 884 mol.\n• Phương trình phản ứng:\n  (C~17~H~33~COO)~3~C~3~H~5~ + 3NaOH ➔ 3C~17~H~33~COONa + C~3~H~5~(OH)~3~\n• Khối lượng sodium oleate (M = 304 g/mol) thu được với hiệu suất 90%:\n  m~sodium oleate~ = (16800 / 884) × 3 × 304 × 90% ≈ 15600 gam.\n• Số bánh xà phòng sản xuất được:\n  x = 15600 / 60 = 260 (bánh).\n➔ Đáp án: 260."
        },
        {
            "num": 6,
            "q": "Câu 6. Một loại chất béo có chứa 80% tristearin về khối lượng. Để sản xuất 10 triệu chai nước rửa tay (có chứa chất dưỡng ẩm glycerol) cần dùng tối thiểu x tấn loại chất béo trên cho phản ứng với dung dịch NaOH, đun nóng. Biết rằng trong mỗi chai nước rửa tay có chứa 6 gam glycerol. Giá trị của x là (làm tròn kết quả đến hàng phần mười). Biết hiệu suất phản ứng đạt 75%.",
            "ans": "967,4",
            "sol": "• Tổng lượng glycerol cần cho 10 triệu chai nước rửa tay:\n  m~glycerol~ = 10 × 10^6 × 6 g = 60 × 10^6 g = 60 tấn.\n• Phản ứng xà phòng hoá tristearin:\n  (C~17~H~35~COO)~3~C~3~H~5~ + 3NaOH ➔ 3C~17~H~35~COONa + C~3~H~5~(OH)~3~\n  890 tấn ............................................................................ 92 tấn\n• Lượng tristearin lí thuyết cần:\n  m~tristearin (LT)~ = 60 × 890 / 92 ≈ 580,435 tấn.\n• Do hiệu suất phản ứng đạt 75%:\n  m~tristearin (TT)~ = 580,435 / 0,75 ≈ 773,913 tấn.\n• Do chất béo chứa 80% tristearin về khối lượng:\n  x = 773,913 / 0,8 ≈ 967,39 tấn ≈ 967,4 tấn.\n➔ Đáp án: 967,4."
        }
    ]

    for p3_data in part3_questions:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, p3_data["q"])

        if is_teacher:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.4)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(1)
            render_rich_text(p_ans, f"► *Đáp án:* *{p3_data['ans']}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
            add_solution_no_table(doc, p3_data["sol"], sol_img=p3_data.get("sol_img"), sol_img_width=p3_data.get("sol_img_width", 12))
        else:
            p_box = doc.add_paragraph()
            p_box.paragraph_format.left_indent = Cm(0.5)
            p_box.paragraph_format.space_before = Pt(1)
            p_box.paragraph_format.space_after = Pt(3)
            r_box = p_box.add_run("Đáp số: ............................................................ ")
            set_run_style(r_box, font_size=10, italic=True, color=COLOR_GRAY)

    # Dòng kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(10)
    p_end.paragraph_format.space_after = Pt(10)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(r_end, font_size=11, bold=True, italic=True)

    out_filename = "De_Kiem_Tra_Hoa_12_De03_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_12_De03.docx"
    out_path = os.path.join(OUT_DIR, out_filename)
    doc.save(out_path)
    print(f"Saved: {out_path}")

if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)
    build_exam_de03(is_teacher=False)
    build_exam_de03(is_teacher=True)
