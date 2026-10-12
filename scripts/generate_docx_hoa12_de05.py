# -*- coding: utf-8 -*-
"""
Script tạo file Word (.docx) chuẩn "FORM ĐỀ KIỂM TRA" cho:
ĐỀ KIỂM TRA SỐ 5 - CHƯƠNG 2: CARBOHYDRATE (HÓA HỌC 12)
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De05.docx (Bản Học Sinh)
2. De_Kiem_Tra_Hoa_12_De05_GiaoVien.docx (Bản Giáo Viên có lời giải chi tiết)
Tuân thủ 100% quy cách: Header & Footer, Tab Stops (NO TABLE), Sub/Superscript công thức hóa học.
"""

import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
IMG_DIR = r"c:\Antigravity_Thanh\San_Pham\images_crop_de05"

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

def setup_page_layout(doc):
    for sec in doc.sections:
        sec.top_margin = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
        sec.left_margin = Cm(1.5)
        sec.right_margin = Cm(1.5)
        sec.header_distance = Cm(0.6)
        sec.footer_distance = Cm(0.6)
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.different_first_page_header_footer = False

def build_header_footer(doc):
    sec = doc.sections[0]
    
    # 1. Header
    header = sec.header
    header.is_linked_to_previous = False
    p_h_old = header.paragraphs[0]
    p_h_old.text = ""
    
    tbl_h = header.add_table(rows=1, cols=2, width=Cm(18.0))
    tbl_h.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_hl = tbl_h.cell(0, 0)
    c_hr = tbl_h.cell(0, 1)
    c_hl.width = Cm(9.0)
    c_hr.width = Cm(9.0)

    p_hl = c_hl.paragraphs[0]
    p_hl.paragraph_format.space_before = Pt(0)
    p_hl.paragraph_format.space_after = Pt(2)
    p_hl.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_hl = p_hl.add_run("Tài liệu lưu hành nội bộ")
    set_run_style(r_hl, font_size=10, italic=True, color=COLOR_HF_RED)

    p_hr = c_hr.paragraphs[0]
    p_hr.paragraph_format.space_before = Pt(0)
    p_hr.paragraph_format.space_after = Pt(2)
    p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_hr1 = p_hr.add_run("Thầy ")
    set_run_style(r_hr1, font_size=10, color=COLOR_HF_RED)
    r_hr2 = p_hr.add_run("TRẦN VĂN THẠNH")
    set_run_style(r_hr2, font_size=10, bold=True, color=COLOR_HF_RED)
    r_hr3 = p_hr.add_run(" - 0777.470.803")
    set_run_style(r_hr3, font_size=10, color=COLOR_HF_RED)

    add_p_border_bottom(p_hl, color=COLOR_HF_GREEN_HEX, sz="6")
    add_p_border_bottom(p_hr, color=COLOR_HF_GREEN_HEX, sz="6")

    # 2. Footer
    footer = sec.footer
    footer.is_linked_to_previous = False
    p_f_old = footer.paragraphs[0]
    p_f_old.text = ""

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
        ("SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ", "ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 2 (CARBOHYDRATE)", 10, True, 11, True),
        ("ĐỀ CHÍNH THỨC", "MÔN: HOÁ HỌC - LỚP 12", 10.5, True, 11, True),
        ("(Đề thi có 04 trang)", "Thời gian: 50 phút | MÃ ĐỀ THI: 105", 9.5, False, 10, True)
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
    render_rich_text(p_info, "*Họ, tên thí sinh:* .......................................................................... *Lớp:* ..................... *Số báo danh:* .....................", base_size=10)

    p_ntk = doc.add_paragraph()
    p_ntk.paragraph_format.space_before = Pt(1)
    p_ntk.paragraph_format.space_after = Pt(5)
    r_ntk = p_ntk.add_run("(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; K = 39; Cu = 64; Ag = 108).")
    set_run_style(r_ntk, font_size=9.5, italic=True)

def add_part_heading(doc, part_title, part_desc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    render_rich_text(p, part_title, base_size=11, base_bold=True, base_color=COLOR_PRIMARY)
    if part_desc:
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(4)
        render_rich_text(p_sub, part_desc, base_size=10, base_italic=True)

def add_question_title(doc, q_num, content, points=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    r_lbl = p.add_run(f"Câu {q_num}: ")
    set_run_style(r_lbl, font_size=10.5, bold=True)
    if points:
        r_pts = p.add_run(f"({points}) ")
        set_run_style(r_pts, font_size=9.5, italic=True, color=COLOR_GRAY)
    render_rich_text(p, content, base_size=10.5)
    return p

def render_single_option(p, opt_text, is_teacher=False, base_size=10.5):
    is_correct = opt_text.startswith('*') and opt_text.endswith('*')
    clean = opt_text[1:-1] if is_correct else opt_text
    m = re.match(r'^([A-D]\.\s*)(.*)$', clean)
    if m:
        lbl = m.group(1)
        body = m.group(2)
    else:
        lbl = ""
        body = clean

    if is_teacher and is_correct:
        if lbl:
            r_lbl = p.add_run(lbl)
            set_run_style(r_lbl, font_size=base_size, bold=True, color=COLOR_CORRECT)
        render_rich_text(p, body, base_size=base_size, base_bold=True, base_color=COLOR_CORRECT)
    else:
        if lbl:
            r_lbl = p.add_run(lbl)
            set_run_style(r_lbl, font_size=base_size, bold=False)
        render_rich_text(p, body, base_size=base_size, base_bold=False)

def add_options_no_table(doc, options, layout="1_line", is_teacher=False):
    tab_1line = [Cm(4.5), Cm(9.0), Cm(13.5)]
    tab_2lines = [Cm(9.0)]

    if layout == "1_line":
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.tab_stops.clear_all()
        for t in tab_1line:
            p.paragraph_format.tab_stops.add_tab_stop(t, WD_TAB_ALIGNMENT.LEFT)
        for i, opt in enumerate(options):
            if i > 0:
                p.add_run('\t')
            render_single_option(p, opt, is_teacher=is_teacher)

    elif layout == "2_lines":
        p1 = doc.add_paragraph()
        p1.paragraph_format.space_before = Pt(1)
        p1.paragraph_format.space_after = Pt(1)
        p1.paragraph_format.tab_stops.clear_all()
        p1.paragraph_format.tab_stops.add_tab_stop(tab_2lines[0], WD_TAB_ALIGNMENT.LEFT)
        render_single_option(p1, options[0], is_teacher=is_teacher)
        p1.add_run('\t')
        render_single_option(p1, options[1], is_teacher=is_teacher)

        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(1)
        p2.paragraph_format.space_after = Pt(2)
        p2.paragraph_format.tab_stops.clear_all()
        p2.paragraph_format.tab_stops.add_tab_stop(tab_2lines[0], WD_TAB_ALIGNMENT.LEFT)
        render_single_option(p2, options[2], is_teacher=is_teacher)
        p2.add_run('\t')
        render_single_option(p2, options[3], is_teacher=is_teacher)

    elif layout == "4_lines":
        for opt in options:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.left_indent = Cm(0.5)
            render_single_option(p, opt, is_teacher=is_teacher)

def render_statement(p, stmt_text, ans=None, is_teacher=False):
    render_rich_text(p, stmt_text, base_size=10.5)
    if is_teacher and ans:
        r_ans = p.add_run(f" -> {ans.upper()}")
        set_run_style(r_ans, font_size=10.5, bold=True, color=COLOR_CORRECT)

def add_solution_no_table(doc, sol_text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(0.5)
    r_title = p.add_run("Lời giải chi tiết: ")
    set_run_style(r_title, font_size=10, bold=True, italic=True, color=COLOR_SOLUTION)
    render_rich_text(p, sol_text, base_size=10, base_italic=True, base_color=COLOR_SOLUTION)

def add_student_lines(doc, num_lines=3):
    for _ in range(num_lines):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.left_indent = Cm(0.5)
        r = p.add_run("....................................................................................................................................................................")
        set_run_style(r, font_size=8.5, color=COLOR_GRAY)

def add_teacher_answer_summary_table(doc):
    p_sec = doc.add_paragraph()
    p_sec.paragraph_format.space_before = Pt(10)
    p_sec.paragraph_format.space_after = Pt(4)
    r = p_sec.add_run("BẢNG ĐÁP ÁN THAM KHẢO")
    set_run_style(r, font_size=11, bold=True, color=COLOR_CORRECT)

    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(2)
    p1.paragraph_format.space_after = Pt(2)
    r1 = p1.add_run("PHẦN I (Mỗi câu trả lời đúng được 0,25 điểm):")
    set_run_style(r1, font_size=10, bold=True, color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=3, cols=10)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    ans_p1 = [
        ("1", "D"), ("2", "D"), ("3", "D"), ("4", "C"), ("5", "C"),
        ("6", "C"), ("7", "D"), ("8", "C"), ("9", "C"), ("10", "D"),
        ("11", "A"), ("12", "A"), ("13", "C"), ("14", "B"), ("15", "D"),
        ("16", "A"), ("17", "D"), ("18", "B")
    ]
    # Row 0: 1 to 9
    for col_idx in range(9):
        c_q = tbl1.cell(0, col_idx)
        c_a = tbl1.cell(1, col_idx)
        c_q.text = f"Câu {ans_p1[col_idx][0]}"
        c_a.text = ans_p1[col_idx][1]
    # Row 1: 10 to 18
    # We can format row 0: headers 1-9, row 1: ans 1-9; row 2: headers 10-18, wait!
    # Let's recreate 2 rows of 9 cols:
    doc._body._element.remove(tbl1._element)
    
    tbl1 = doc.add_table(rows=4, cols=9)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    for col in range(9):
        tbl1.cell(0, col).text = f"Câu {ans_p1[col][0]}"
        tbl1.cell(1, col).text = ans_p1[col][1]
        tbl1.cell(2, col).text = f"Câu {ans_p1[col+9][0]}"
        tbl1.cell(3, col).text = ans_p1[col+9][1]
    
    for row in tbl1.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in cell.paragraphs[0].runs:
                set_run_style(run, font_size=9, bold=True)

    # Phần II
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(6)
    p2.paragraph_format.space_after = Pt(2)
    r2 = p2.add_run("PHẦN II (Điểm tối đa 4,0 điểm - Chọn đúng 1 ý: 0,1 đ; 2 ý: 0,25 đ; 3 ý: 0,5 đ; 4 ý: 1,0 đ):")
    set_run_style(r2, font_size=10, bold=True, color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=5, cols=5)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl2.cell(0, 0).text = "Câu / Lệnh"
    tbl2.cell(0, 1).text = "Câu 1"
    tbl2.cell(0, 2).text = "Câu 2"
    tbl2.cell(0, 3).text = "Câu 3"
    tbl2.cell(0, 4).text = "Câu 4"

    p2_data = [
        ("a)", "S", "Đ", "S", "S"),
        ("b)", "Đ", "Đ", "Đ", "S"),
        ("c)", "S", "S", "Đ", "Đ"),
        ("d)", "Đ", "S", "S", "Đ")
    ]
    for row_idx, row_vals in enumerate(p2_data):
        for col_idx, val in enumerate(row_vals):
            tbl2.cell(row_idx+1, col_idx).text = val

    for row in tbl2.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in cell.paragraphs[0].runs:
                set_run_style(run, font_size=9.5, bold=True)

    # Phần III
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(6)
    p3.paragraph_format.space_after = Pt(2)
    r3 = p3.add_run("PHẦN III (Mỗi câu trả lời đúng được 0,25 điểm):")
    set_run_style(r3, font_size=10, bold=True, color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=6)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    ans_p3 = [("1", "2"), ("2", "4"), ("3", "3"), ("4", "5832"), ("5", "21,2"), ("6", "9")]
    for col in range(6):
        tbl3.cell(0, col).text = f"Câu {ans_p3[col][0]}"
        tbl3.cell(1, col).text = ans_p3[col][1]

    for row in tbl3.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in cell.paragraphs[0].runs:
                set_run_style(run, font_size=9.5, bold=True)


def build_docx(is_teacher=False):
    doc = docx.Document()
    setup_page_layout(doc)
    build_header_footer(doc)
    build_title_block_no_table(doc, is_teacher=is_teacher)

    # =========================================================================
    # PHẦN I
    # =========================================================================
    add_part_heading(doc, "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
                     "Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")

    # Questions Part I
    p1_questions = [
        (1, "Saccharose và fructose đều thuộc loại",
         ["A. monosaccharide.", "B. disaccharide.", "C. polysaccharide.", "*D. carbohydrate.*"],
         "1_line",
         "Saccharose là disaccharide, fructose là monosaccharide; cả hai chất đều thuộc nhóm carbohydrate."),
        
        (2, "Một chất khi thủy phân trong môi trường acid, đun nóng không tạo ra glucose. Chất đó là",
         ["A. tinh bột.", "B. saccharose.", "C. cellulose.", "*D. tripalmitin.*"],
         "1_line",
         "Tripalmitin là triglyceride (chất béo), thủy phân trong acid tạo palmitic acid và glycerol, hoàn toàn không sinh ra glucose."),
        
        (3, "(SBT – CTST). Carbohydrate nào sau đây là thành phần chính của mật ong?",
         ["A. Glucose.", "B. Maltose.", "C. Saccharose.", "*D. Fructose.*"],
         "1_line",
         "Fructose chiếm tỉ lệ cao nhất trong mật ong (khoảng 40% khối lượng mật ong), glucose chiếm khoảng 30%."),
        
        (4, "(SBT – KNTT). Saccharose thường được tìm thấy trong loại thực vật nào sau đây?",
         ["A. Cây đậu nành.", "B. Cây lúa mì.", "*C. Cây mía.*", "D. Cây cà phê."],
         "1_line",
         "Saccharose có nhiều trong cây mía (nên còn gọi là đường mía), củ cải đường và hoa thốt nốt."),
        
        (5, "(SBT – KNTT). Công thức phân tử của saccharose là",
         ["A. C~5~H~10~O~5~.", "B. C~6~H~12~O~6~.", "*C. C~12~H~22~O~11~.*", "D. (C~6~H~10~O~5~)~n~."],
         "1_line",
         "Saccharose là disaccharide tạo bởi một gốc α-glucose và một gốc β-fructose, CTPT là C~12~H~22~O~11~."),
        
        (6, "(SBT – KNTT). Tinh bột và cellulose đều là",
         ["A. disaccharide.", "B. monosaccharide.", "*C. polysaccharide.*", "D. oligosaccharide."],
         "1_line",
         "Cả tinh bột và cellulose đều là các polymer tự nhiên thuộc nhóm polysaccharide có công thức chung (C~6~H~10~O~5~)~n~."),
        
        (7, "(SBT – KNTT). Phản ứng màu với dung dịch iodine là tính chất của chất nào sau đây?",
         ["A. Glucose.", "B. Fructose.", "C. Saccharose.", "*D. Tinh bột.*"],
         "1_line",
         "Cấu trúc xoắn dạng ống của amylose trong tinh bột hấp phụ các phân tử iodine tạo hợp chất bọc màu xanh tím đặc trưng."),
        
        (8, "Glucose có thể lên men tạo lactic acid có trong sữa chua qua phản ứng nào sau đây?",
         ["A. 2C~6~H~12~O~6~ + Cu(OH)~2~ -> (C~6~H~11~O~6~)~2~Cu + 2H~2~O",
          "B. C~6~H~12~O~6~ -(enzyme)-> 2C~2~H~5~OH + 2CO~2~",
          "*C. C~6~H~12~O~6~ -(enzyme)-> 2CH~3~CH(OH)COOH.*",
          "D. CH~2~OH-[CHOH]~4~-CHO + Br~2~ + H~2~O -(t°)-> CH~2~OH-[CHOH]~4~-COOH + 2HBr"],
         "4_lines",
         "Lên men lactic chuyển hóa glucose thành lactic acid (CH~3~CH(OH)COOH): C~6~H~12~O~6~ -(men lactic)-> 2CH~3~CH(OH)COOH."),
        
        (9, "(SGK 12 - CTST). Cho các carbohydrate sau: glucose, fructose, saccharose và maltose. Số carbohydrate có khả năng mở vòng trong dung dịch nước là",
         ["A. 1.", "B. 2.", "*C. 3.*", "D. 4."],
         "1_line",
         "Glucose, fructose và maltose đều có nhóm hemiacetal/hemiketal tự do nên có thể mở vòng trong dung dịch nước; saccharose bị khóa liên kết glycosidic ở cả 2 đầu nên không thể mở vòng -> Có 3 chất.")
    ]

    for q_num, title, opts, layout, sol in p1_questions:
        add_question_title(doc, q_num, title)
        add_options_no_table(doc, opts, layout=layout, is_teacher=is_teacher)
        if is_teacher:
            add_solution_no_table(doc, sol)

    # Câu 10 (Haworth formulas)
    add_question_title(doc, 10, "Công thức cấu tạo dạng mạch vòng β-fructose là:")
    # Render 4 images side-by-side using Tab stops
    p_img10 = doc.add_paragraph()
    p_img10.paragraph_format.space_before = Pt(2)
    p_img10.paragraph_format.space_after = Pt(2)
    p_img10.paragraph_format.tab_stops.clear_all()
    p_img10.paragraph_format.tab_stops.add_tab_stop(Cm(4.5), WD_TAB_ALIGNMENT.LEFT)
    p_img10.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
    p_img10.paragraph_format.tab_stops.add_tab_stop(Cm(13.5), WD_TAB_ALIGNMENT.LEFT)

    img_a = os.path.join(IMG_DIR, "cau10_opt_A.png")
    img_b = os.path.join(IMG_DIR, "cau10_opt_B.png")
    img_c = os.path.join(IMG_DIR, "cau10_opt_C.png")
    img_d = os.path.join(IMG_DIR, "cau10_opt_D.png")

    # Line of options text: A.   B.   C.   D.
    p_lbl10 = doc.add_paragraph()
    p_lbl10.paragraph_format.space_before = Pt(1)
    p_lbl10.paragraph_format.space_after = Pt(2)
    p_lbl10.paragraph_format.tab_stops.clear_all()
    p_lbl10.paragraph_format.tab_stops.add_tab_stop(Cm(4.5), WD_TAB_ALIGNMENT.LEFT)
    p_lbl10.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
    p_lbl10.paragraph_format.tab_stops.add_tab_stop(Cm(13.5), WD_TAB_ALIGNMENT.LEFT)

    c10_opts = ["A.", "B.", "C.", "*D.*"]
    for i, opt in enumerate(c10_opts):
        if i > 0:
            p_lbl10.add_run('\t')
        render_single_option(p_lbl10, opt, is_teacher=is_teacher)

    if os.path.exists(img_a) and os.path.exists(img_b) and os.path.exists(img_c) and os.path.exists(img_d):
        p_img10.add_run().add_picture(img_a, width=Cm(3.8))
        p_img10.add_run('\t')
        p_img10.add_run().add_picture(img_b, width=Cm(3.8))
        p_img10.add_run('\t')
        p_img10.add_run().add_picture(img_c, width=Cm(3.8))
        p_img10.add_run('\t')
        p_img10.add_run().add_picture(img_d, width=Cm(3.8))

    if is_teacher:
        add_solution_no_table(doc, "Fructose tạo vòng furanose 5 cạnh. Ở dạng β-fructose, nhóm hemiketal -OH tại vị trí C2 hướng lên trên (cùng phía với nhóm -CH~2~OH ở C5/C6) ứng với công thức ở đáp án D.")

    # Questions 11-18
    p1_part2 = [
        (11, "Mô tả nào dưới đây không đúng với glucose?",
         ["*A. Chất rắn, khó tan trong nước và có vị ngọt.*",
          "B. Có mặt trong hầu hết các bộ phận của cây, nhất là trong quả chín.",
          "C. Còn có tên gọi là đường nho.",
          "D. Có 80-130 mg/dL trong máu người trưởng thành, khỏe mạnh và lúc đói."],
         "4_lines",
         "Glucose là chất rắn màu trắng, *dễ tan trong nước*, không phải khó tan."),
        
        (12, "Cho các tính chất sau: (1) có vị ngọt, (2) dễ tan trong nước, (3) có phản ứng tráng bạc, (4) bị thủy phân trong môi trường acid, (5) hòa tan Cu(OH)~2~ thành dung dịch màu xanh lam.\nSố tính chất đúng với saccharose là",
         ["*A. 4.*", "B. 3.", "C. 5.", "D. 2."],
         "1_line",
         "Saccharose có các tính chất: (1), (2), (4), (5). Tính chất (3) không có vì saccharose không chứa nhóm aldehyde (-CHO) hay hemiacetal tự do để tráng bạc -> Có 4 tính chất đúng."),
        
        (13, "Nhận định sai là",
         ["A. Phân biệt glucose và saccharose bằng phản ứng tráng gương.",
          "B. Phân biệt tinh bột và cellulose bằng I~2~.",
          "*C. Phân biệt saccharose và glycerol bằng Cu(OH)~2~.*",
          "D. Phân biệt tinh bột và saccharose bằng I~2~."],
         "4_lines",
         "Cả saccharose và glycerol đều hòa tan Cu(OH)~2~ ở điều kiện thường tạo phức chất màu xanh lam thẫm, do đó không thể dùng Cu(OH)~2~ ở điều kiện thường để phân biệt hai chất này."),
        
        (14, "(SBT – CTST). Hầu hết những phản ứng sinh hoá xảy ra nhờ sự xúc tác của các enzyme. Biểu đồ dưới đây thể hiện tốc độ của một phản ứng sinh hoá do enzyme X làm chất tác theo nhiệt độ:",
         ["A. 0°C.", "*B. 50°C.*", "C. 60°C.", "D. 70°C."],
         "1_line",
         "Từ đồ thị thực nghiệm, đỉnh cao nhất của đường cong tốc độ phản ứng ứng với hoành độ 50°C, vậy 50°C là nhiệt độ tối ưu nhất cho enzyme X."),
        
        (15, "(SBT – CTST). Carbohydrate nào có cấu trúc phân tử được biểu diễn dưới đây?",
         ["A. Saccharose.", "B. Cellulose.", "C. Maltose.", "*D. Amylose.*"],
         "1_line",
         "Cấu trúc gồm các mắt xích α-glucose liên kết qua liên kết α-1,4-glycosidic tạo chuỗi polymer không phân nhánh dài khoảng 1000 mắt xích, đây chính là amylose của tinh bột."),
        
        (16, "Cho các phát biểu sau về carbohydrate:\n"
             "(a) Tất cả các carbohydrate đều có phản ứng thủy phân.\n"
             "(b) Thủy phân hoàn toàn tinh bột thu được glucose.\n"
             "(c) Glucose, fructose đều có phản ứng tráng bạc.\n"
             "(d) Glucose làm mất màu nước bromine.\n"
             "Số phát biểu đúng là:",
         ["*A. 3.*", "B. 4.", "C. 1.", "D. 2."],
         "1_line",
         "- (a) SAI: Các monosaccharide (glucose, fructose) không bị thủy phân.\n"
         "- (b) ĐÚNG: Tinh bột thủy phân hoàn toàn tạo glucose.\n"
         "- (c) ĐÚNG: Cả hai đều cho phản ứng tráng bạc (trong môi trường kiềm, fructose chuyển hóa thành glucose).\n"
         "- (d) ĐÚNG: Nhóm aldehyde trong glucose bị brom oxi hóa làm mất màu nước Br~2~.\n"
         "=> Có 3 phát biểu đúng: (b), (c), (d)."),
        
        (17, "Cho các phát biểu sau:\n"
             "(a) Đốt cháy hoàn toàn ester no, đơn chức, mạch hở luôn thu được số mol CO~2~ bằng số mol H~2~O.\n"
             "(b) Trong hợp chất hữu cơ nhất thiết phải có carbon và hydrogen.\n"
             "(c) Những hợp chất hữu cơ có thành phần nguyên tố giống nhau, thành phần phân tử hơn kém nhau một hay nhiều nhóm CH~2~ là đồng đẳng của nhau.\n"
             "(d) Dung dịch glucose bị khử bởi AgNO~3~ trong NH~3~ tạo ra Ag.\n"
             "(e) Saccharose chỉ có cấu tạo mạch vòng.\n"
             "Số phát biểu đúng là",
         ["A. 4.", "B. 5.", "C. 3.", "*D. 2.*"],
         "1_line",
         "- (a) ĐÚNG: C~n~H~2n~O~2~ cháy ra n CO~2~ và n H~2~O.\n"
         "- (b) SAI: Hợp chất hữu cơ nhất thiết phải có C, có thể không có H (vd: CCl~4~, CF~4~, C~2~F~4~).\n"
         "- (c) SAI: Đồng đẳng bắt buộc phải có tính chất hóa học tương tự nhau (cấu tạo tương tự nhau).\n"
         "- (d) SAI: Glucose bị *oxi hóa* thành muối gluconate; Ag^+ bị khử thành Ag.\n"
         "- (e) ĐÚNG: Saccharose không có nhóm hemiacetal nên không thể mở vòng, chỉ tồn tại ở dạng mạch vòng.\n"
         "=> Có 2 phát biểu đúng: (a), (e)."),
        
        (18, "Cho các phát biểu sau:\n"
             "(a) Hydrogen hóa hoàn toàn glucose tạo ra acid gluconic.\n"
             "(b) Ở điều kiện thường, glucose và saccharose đều là những chất rắn, dễ tan trong nước.\n"
             "(c) Cellulose trinitrate là nguyên liệu để sản xuất tơ nhân tạo và chế tạo thuốc súng không khói.\n"
             "(d) Amylopectin trong tinh bột chỉ có các liên kết α-1,4-glycoside.\n"
             "(e) Saccharose bị hóa đen trong H~2~SO~4~ đặc.\n"
             "(f) Trong công nghiệp dược phẩm, saccharose được dùng để pha chế thuốc.\n"
             "Trong các phát biểu trên, số phát biểu đúng là",
         ["A. 4.", "*B. 3.*", "C. 2.", "D. 5."],
         "1_line",
         "- (a) SAI: Hydrogen hóa glucose tạo sorbitol (alcohol đa chức), không phải gluconic acid.\n"
         "- (b) ĐÚNG: Đều là chất rắn kết tinh không màu, vị ngọt, dễ tan trong nước.\n"
         "- (c) SAI: Cellulose trinitrate dùng làm thuốc súng không khói, không dùng làm tơ nhân tạo.\n"
         "- (d) SAI: Amylopectin có mạch phân nhánh gồm cả liên kết α-1,4 và α-1,6-glycosidic.\n"
         "- (e) ĐÚNG: H~2~SO~4~ đặc háo nước hút nước làm saccharose bị than hóa thành C màu đen.\n"
         "- (f) ĐÚNG: Dùng làm siro, chất dẫn và bao viên thuốc.\n"
         "=> Có 3 phát biểu đúng: (b), (e), (f).")
    ]

    for q_num, title, opts, layout, sol in p1_part2:
        add_question_title(doc, q_num, title)
        if q_num == 14:
            img_c14 = os.path.join(IMG_DIR, "cau14_do_thi_toc_do.png")
            if os.path.exists(img_c14):
                p_im = doc.add_paragraph()
                p_im.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_im.paragraph_format.space_before = Pt(2)
                p_im.paragraph_format.space_after = Pt(2)
                p_im.add_run().add_picture(img_c14, width=Cm(9.5))
        elif q_num == 15:
            img_c15 = os.path.join(IMG_DIR, "cau15_cau_truc_amylose.png")
            if os.path.exists(img_c15):
                p_im = doc.add_paragraph()
                p_im.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_im.paragraph_format.space_before = Pt(2)
                p_im.paragraph_format.space_after = Pt(2)
                p_im.add_run().add_picture(img_c15, width=Cm(11.0))

        add_options_no_table(doc, opts, layout=layout, is_teacher=is_teacher)
        if is_teacher:
            add_solution_no_table(doc, sol)

    # =========================================================================
    # PHẦN II
    # =========================================================================
    add_part_heading(doc, "PHẦN II. Câu trắc nghiệm đúng sai.",
                     "Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")

    # Câu 1
    add_question_title(doc, 1, "(SBT – KNTT).")
    c1_stmts = [
        ("a) Tất cả các loại carbohydrate đều tan hoàn toàn trong nước.", "Sai"),
        ("b) Một số đường đơn (monosaccharide) có vị ngọt.", "Đúng"),
        ("c) Carbohydrate chỉ được tìm thấy trong thực phẩm có nguồn gốc thực vật.", "Sai"),
        ("d) Glucose và fructose đều là chất rắn, vị ngọt, tan tốt trong nước.", "Đúng")
    ]
    for stmt, ans in c1_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
                              "- a) SAI: Tinh bột không tan trong nước lạnh (tan một phần trong nước nóng tạo hồ), cellulose hoàn toàn không tan trong nước.\n"
                              "- b) ĐÚNG: Glucose, fructose có vị ngọt thanh/ngọt đậm.\n"
                              "- c) SAI: Carbohydrate có cả trong động vật, ví dụ glycogen dự trữ trong gan, lactose trong sữa mẹ và sữa động vật.\n"
                              "- d) ĐÚNG: Cả hai đều là chất kết tinh không màu, vị ngọt, dễ tan trong nước.")

    # Câu 2
    add_question_title(doc, 2, "Saccharose được tiêu thụ với lượng lớn trên toàn cầu.")
    img_mia = os.path.join(IMG_DIR, "p2_cau2_cay_mia_thot_not_cu_cai.png")
    if os.path.exists(img_mia):
        p_im = doc.add_paragraph()
        p_im.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_im.paragraph_format.space_before = Pt(2)
        p_im.paragraph_format.space_after = Pt(2)
        p_im.add_run().add_picture(img_mia, width=Cm(12.5))

    c2_stmts = [
        ("a) Saccharose còn được gọi là đường ăn, là chất rắn, vị ngọt, dễ tan trong nước.", "Đúng"),
        ("b) Saccharose có nhiều trong cây mía, hoa thốt nốt, củ cải đường.", "Đúng"),
        ("c) Thủy phân hoàn toàn saccharose trong môi trường acid chỉ thu được glucose.", "Sai"),
        ("d) Cấu tạo phân tử saccharose gồm 1 gốc α-glucose và 1 gốc β-fructose liên kết với nhau qua nguyên tử carbon.", "Sai")
    ]
    for stmt, ans in c2_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
                              "- a) ĐÚNG: Saccharose chính là đường kính/đường ăn thường ngày, có vị ngọt dễ chịu, tan rất tốt trong nước.\n"
                              "- b) ĐÚNG: Đây là 3 nguồn thực vật dồi dào saccharose nhất trong tự nhiên.\n"
                              "- c) SAI: Thủy phân saccharose thu được 1 phân tử α-glucose và 1 phân tử β-fructose.\n"
                              "- d) SAI: Hai gốc monosaccharide liên kết với nhau qua *nguyên tử oxygen* (cầu nối α-1,2-glycosidic), không phải qua nguyên tử carbon.")

    # Câu 3
    add_question_title(doc, 3, "Cellulose là một polysaccharide gồm nhiều gốc β-glucose liên kết với nhau tạo ra dạng mạch không nhánh, kéo dài thành dạng sợi có cấu tạo như sau:")
    img_cell = os.path.join(IMG_DIR, "p2_cau3_cau_truc_cellulose.png")
    if os.path.exists(img_cell):
        p_im = doc.add_paragraph()
        p_im.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_im.paragraph_format.space_before = Pt(2)
        p_im.paragraph_format.space_after = Pt(2)
        p_im.add_run().add_picture(img_cell, width=Cm(12.0))

    c3_stmts = [
        ("a) Đoạn mạch trên có cả hai loại liên kết là: liên kết β-1,4-glycoside và liên kết β-1,6-glycoside.", "Sai"),
        ("b) Các gốc β-glucose trong đoạn mạch trên liên kết với nhau bởi liên kết β-1,4-glycoside.", "Đúng"),
        ("c) Công thức mỗi gốc β-glucose là C~6~H~10~O~5~.", "Đúng"),
        ("d) Đoạn mạch trên có cấu tạo phân nhánh ở liên kết α-1,6-glycoside.", "Sai")
    ]
    for stmt, ans in c3_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
                              "- a) SAI: Phân tử cellulose mạch thẳng chỉ chứa duy nhất liên kết β-1,4-glycosidic.\n"
                              "- b) ĐÚNG: Các gốc β-glucose liên kết với nhau qua nguyên tử oxygen giữa C1 của gốc này và C4 của gốc kế tiếp (liên kết β-1,4-glycosidic).\n"
                              "- c) ĐÚNG: Công thức mỗi mắt xích là -[C~6~H~10~O~5~]-.\n"
                              "- d) SAI: Cellulose là phân tử dạng sợi mạch thẳng dài, không phân nhánh.")

    # Câu 4
    add_question_title(doc, 4,
                       "Phản ứng thuỷ phân cellulose trong môi trường acid\n"
                       "*Chuẩn bị:*\n"
                       "Hoá chất: cellulose (bông), dung dịch H~2~SO~4~ 70%, dung dịch NaOH 10%, dung dịch CuSO~4~ 5%, NaHCO~3~ rắn, nước nóng.\n"
                       "Dụng cụ: ống nghiệm, cốc thuỷ tinh 250 mL, đũa thuỷ tinh, đèn cồn, giá đựng ống nghiệm.\n"
                       "*Tiến hành:*\n"
                       "Bước 1: Cho 10 mL dung dịch H~2~SO~4~ 70% vào cốc thuỷ tinh, thêm một lượng nhỏ cellulose (bông) vào cốc và dùng đũa thuỷ tinh khuấy đều. Sau đó, đặt cốc thuỷ tinh vào cốc nước nóng và khuấy trong khoảng 3 phút để cellulose tan hết tạo dung dịch đồng nhất.\n"
                       "Bước 2: Trung hoà dung dịch bằng cách thêm từ từ NaHCO~3~ đến khi dừng sủi bọt khí, sau đó thêm tiếp 5 mL dung dịch NaOH 10%.\n"
                       "Bước 3: Cho 5 mL dung dịch thu được ở trên vào ống nghiệm chứa Cu(OH)~2~ (được điều chế bằng cách cho 0,5 mL dung dịch CuSO~4~ 5% vào 2 mL dung dịch NaOH 10%, lắc nhẹ). Đun nóng đều ống nghiệm khoảng 2 phút, sau đó để ống nghiệm trên giá khoảng 3 phút.")
    c4_stmts = [
        ("a) Sau bước 1, sản phẩm thu được có glucose.", "Sai"),
        ("b) Ở bước 2 thêm từ từ NaHCO~3~ để phản ứng diễn ra nhanh hơn.", "Sai"),
        ("c) Sau bước 3, lấy sản phẩm thu được có kết tủa màu đỏ gạch.", "Đúng"),
        ("d) Thí nghiệm trên chứng minh trong cellulose có thể bị thủy phân trong môi trường acid tạo glucose.", "Đúng")
    ]
    for stmt, ans in c4_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
                              "- a) SAI: Sau bước 1 trong nước nóng 3 phút cellulose mới chỉ bị thủy phân sơ bộ thành các đoạn dextrin mạch ngắn và dung dịch nhớt đồng nhất, chưa thu được hoàn toàn glucose tự do.\n"
                              "- b) SAI: Thêm NaHCO~3~ là để trung hòa H~2~SO~4~ dư (tránh làm tan kết tủa Cu(OH)~2~ ở bước sau).\n"
                              "- c) ĐÚNG: Sản phẩm thủy phân có glucose khử Cu(OH)~2~ trong môi trường kiềm khi đun nóng tạo kết tủa đỏ gạch Cu~2~O.\n"
                              "- d) ĐÚNG: Hiện tượng tạo kết tủa Cu~2~O chứng minh phản ứng thủy phân cellulose giải phóng monosaccharide có tính khử (glucose).")

    # =========================================================================
    # PHẦN III
    # =========================================================================
    add_part_heading(doc, "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
                     "Thí sinh trả lời từ câu 1 đến câu 6.")

    # Câu 1
    add_question_title(doc, 1, "Cho dãy các chất: tinh bột, saccharose, glucose, fructose, cellulose. Số chất trong dãy thuộc loại monosaccharide là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Có 2 chất là glucose và fructose. (Saccharose là disaccharide; tinh bột và cellulose là polysaccharide).\nĐáp số: 2.")
    else:
        add_student_lines(doc, 2)

    # Câu 2
    add_question_title(doc, 2, "Cho các tính chất sau: (1) có vị ngọt, (2) dễ tan trong nước, (3) có phản ứng tráng bạc, (4) bị thủy phân trong môi trường acid, (5) hòa tan Cu(OH)~2~ thành dung dịch màu xanh lam. Số tính chất đúng với saccharose là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Có 4 tính chất đúng gồm: (1), (2), (4), (5). Tính chất (3) sai vì saccharose không có tính khử tráng bạc.\nĐáp số: 4.")
    else:
        add_student_lines(doc, 2)

    # Câu 3
    add_question_title(doc, 3, "(SBT – KNTT). Trong số các chất saccharose, maltose, tinh bột và cellulose, có bao nhiêu chất khi thuỷ phân hoàn toàn sản phẩm thu được chỉ là glucose?")
    if is_teacher:
        add_solution_no_table(doc, "Có 3 chất là maltose, tinh bột và cellulose:\n"
                                   "- Maltose -> 2 glucose\n"
                                   "- Tinh bột -(+H~2~O, H^+)-> n glucose\n"
                                   "- Cellulose -(+H~2~O, H^+)-> n glucose\n"
                                   "(Saccharose khi thủy phân tạo 1 glucose và 1 fructose nên không thỏa mãn).\nĐáp số: 3.")
    else:
        add_student_lines(doc, 2)

    # Câu 4
    add_question_title(doc, 4,
                       "Giả thiết phản ứng tổng hợp glucose ở cây lương thực như sau:\n"
                       "6CO~2~ + 6H~2~O -(asmt, chlorophyll)-> C~6~H~12~O~6~ + 6O~2~     Δ~r~H°~298~ = 2800 kJ\n"
                       "Trung bình mỗi phút, 1 m² mặt đất nhận được từ Mặt Trời 60 kJ nhưng chỉ 1% được hấp thụ ở quá trình quang hợp. Khối lượng glucose tổng hợp được trong 7 giờ trên 1 sào (360 m²) trồng cây lương thực là bao nhiêu gam?")
    if is_teacher:
        add_solution_no_table(doc,
                              "- Tổng năng lượng ánh sáng cây hấp thụ trong 7 giờ trên 1 sào (360 m²):\n"
                              "  Q = 360 × (7 × 60) × 60 × 1% = 90.720 kJ.\n"
                              "- Số mol glucose tạo thành:\n"
                              "  n~glucose~ = 90.720 / 2.800 = 32,4 mol.\n"
                              "- Khối lượng glucose thu được:\n"
                              "  m~glucose~ = 32,4 × 180 = 5.832 gam.\n"
                              "Đáp số: 5832.")
    else:
        add_student_lines(doc, 3)

    # Câu 5
    add_question_title(doc, 5, "Để tráng bạc (silver) một số ruột phích, người ta phải thuỷ phân 34,2 gam saccharose rồi tiến hành phản ứng tráng gương. Có m gam Ag tạo thành sau phản ứng, biết hiệu suất mỗi quá trình là 70%. Xác định giá trị m? (Làm tròn kết quả đến chữ số thập phân thứ nhất).")
    if is_teacher:
        add_solution_no_table(doc,
                              "- n~saccharose~ = 34,2 / 342 = 0,1 mol.\n"
                              "- Sơ đồ: C~12~H~22~O~11~ -(thủy phân)-> 1 glucose + 1 fructose -(tráng bạc)-> 4 Ag.\n"
                              "- Vì hiệu suất mỗi quá trình là 70% (0,7):\n"
                              "  n~Ag~ = 0,1 × 4 × 0,7 × 0,7 = 0,196 mol.\n"
                              "- Khối lượng bạc tạo thành:\n"
                              "  m~Ag~ = 0,196 × 108 = 21,168 gam ≈ 21,2 gam.\n"
                              "Đáp số: 21,2.")
    else:
        add_student_lines(doc, 3)

    # Câu 6
    add_question_title(doc, 6,
                       "Xăng E5 được sản xuất bằng cách phối trộn xăng khoáng RON 92 với ethanol (D = 0,8 gam/mL) theo tỉ lệ thể tích tương ứng là 95 : 5. Ethanol được sản xuất từ tinh bột bằng phương pháp lên men. Tinh bột có nhiều trong sắn, ngô. Cho bảng thông tin sau:\n"
                       "• Sắn tươi (khoai mì): hàm lượng tinh bột 30%, giá 1.850 đồng/kg, hiệu suất toàn bộ quá trình đạt 79,83%.\n"
                       "• Ngô khô (bắp khô): hàm lượng tinh bột 75%, giá 6.500 đồng/kg, hiệu suất toàn bộ quá trình đạt 84,783%.\n"
                       "Để điều chế 50.000 lít xăng E5 từ sắn tươi cần chi x triệu đồng, từ ngô khô cần chi y triệu đồng. Tính chênh lệch tiền nguyên liệu là z (triệu đồng) khi dùng sắn tươi so với dùng ngô khô để sản xuất. Tính giá trị của z (làm tròn kết quả đến phần triệu).")
    img_san = os.path.join(IMG_DIR, "p3_cau6_san_va_ngo.png")
    if os.path.exists(img_san):
        p_im = doc.add_paragraph()
        p_im.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_im.paragraph_format.space_before = Pt(2)
        p_im.paragraph_format.space_after = Pt(2)
        p_im.add_run().add_picture(img_san, width=Cm(11.0))

    if is_teacher:
        add_solution_no_table(doc,
                              "- Thể tích ethanol cần dùng: V~ethanol~ = 50.000 × 5% = 2.500 lít = 2.500.000 mL.\n"
                              "- Khối lượng ethanol: m~ethanol~ = 2.500.000 × 0,8 = 2.000.000 gam = 2.000 kg.\n"
                              "- Phương trình lên men: (-C~6~H~10~O~5~-)~n~ -> n C~6~H~12~O~6~ -> 2n C~2~H~5~OH.\n"
                              "  Cứ 162 kg tinh bột lý thuyết cho 2 × 46 = 92 kg ethanol.\n"
                              "  => Khối lượng tinh bột lý thuyết: m~tinh bột (LT)~ = 2.000 × 162 / 92 ≈ 3.521,739 kg.\n"
                              "- Chi phí từ sắn tươi (hàm lượng 30%, H = 79,83%):\n"
                              "  x = 3.521,739 × (100 / 30) × (100 / 79,83) × 1.850 ≈ 27,203 × 10^6 đồng ≈ 27 triệu đồng.\n"
                              "- Chi phí từ ngô khô (hàm lượng 75%, H = 84,783%):\n"
                              "  y = 3.521,739 × (100 / 75) × (100 / 84,783) × 6.500 ≈ 36,000 × 10^6 đồng ≈ 36 triệu đồng.\n"
                              "- Chênh lệch tiền nguyên liệu: z = |y - x| = 36 - 27 = 9 triệu đồng.\n"
                              "Đáp số: 9.")
    else:
        add_student_lines(doc, 4)

    # Bảng đáp án cho giáo viên
    if is_teacher:
        add_teacher_answer_summary_table(doc)

    out_name = "De_Kiem_Tra_Hoa_12_De05_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_12_De05.docx"
    out_path = os.path.join(OUT_DIR, out_name)
    doc.save(out_path)
    print(f"File created successfully: {out_path}")

if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)
    build_docx(is_teacher=False)
    build_docx(is_teacher=True)
