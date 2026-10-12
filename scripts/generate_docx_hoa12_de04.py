# -*- coding: utf-8 -*-
"""
Script tạo file Word (.docx) chuẩn "FORM ĐỀ KIỂM TRA" cho:
ĐỀ KIỂM TRA SỐ 4 - CHƯƠNG 2: CARBOHYDRATE (HÓA HỌC 12)
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De04.docx (Bản Học Sinh)
2. De_Kiem_Tra_Hoa_12_De04_GiaoVien.docx (Bản Giáo Viên có lời giải chi tiết)
Tuân thủ 100% quy cách: Header & Footer, Tab Stops (NO TABLE), Sub/Superscript công thức hóa học.
"""

import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
IMG_DIR = r"c:\Antigravity_Thanh\San_Pham\images_crop_de04"

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
        ("SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ", "ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 2 (CARBOHYDRATE)", 10, True, 11, True),
        ("ĐỀ CHÍNH THỨC", "MÔN: HOÁ HỌC - LỚP 12", 10.5, True, 11, True),
        ("(Đề thi có 04 trang)", "Thời gian: 50 phút | MÃ ĐỀ THI: 104", 9.5, False, 10, True)
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

def build_exam_de04(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)
    build_title_block_no_table(doc, is_teacher=is_teacher)

    # =========================================================================
    # PHẦN I
    # =========================================================================
    add_part_heading(doc, "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
                     "Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")

    p1_data = [
        (1, "Đồng phân của glucose là",
         ["*A. fructose.*", "B. tinh bột.", "C. saccharose.", "D. cellulose."],
         "Glucose và fructose đều có cùng công thức phân tử là C~6~H~12~O~6~, chúng là đồng phân cấu tạo của nhau (glucose là polyhydroxyaldehyde, fructose là polyhydroxyketone)."),

        (2, "Ứng dụng nào sau đây *không* phải là ứng dụng của glucose?",
         ["A. Tráng gương, tráng phích.", "*B. Nguyên liệu sản xuất chất dẻo PVC.*",
          "C. Nguyên liệu sản xuất ethyl alcohol.", "D. Làm thực phẩm dinh dưỡng và thuốc tăng lực."],
         "Chất dẻo PVC (poly(vinyl chloride)) được sản xuất từ monomer vinyl chloride (CH~2~=CH-Cl), không phải từ glucose."),

        (3, "Glucose là chất dinh dưỡng và được dùng làm thuốc tăng lực cho người già, trẻ em và người ốm. Số nguyên tử carbon trong phân tử glucose là",
         ["*A. 6.*", "B. 11.", "C. 5.", "D. 12."],
         "Phân tử glucose có công thức phân tử là C~6~H~12~O~6~, chứa 6 nguyên tử carbon."),

        (4, "(SBT – CTST) Hormone nào sau đây làm giảm lượng glucose trong máu?",
         ["A. Adrenaline.", "B. Thyroxine.", "*C. Insuline.*", "D. Oxytocine."],
         "Insulin là hormone do tuyến tụy tiết ra có vai trò chuyển hóa glucose thành glycogen dự trữ trong gan và cơ, làm giảm nồng độ glucose trong máu."),

        (5, "(SBT – KNTT) Phản ứng đặc trưng của maltose là",
         ["A. Phản ứng với dung dịch NaOH.", "B. Phản ứng màu với iodine.",
          "*C. Phản ứng thuỷ phân tạo ra glucose.*", "D. Phản ứng lên men trực tiếp tạo ra ethanol."],
         "Maltose là disaccharide tạo bởi hai gốc α-glucose liên kết qua liên kết α-1,4-glycosidic. Phản ứng đặc trưng của disaccharide là phản ứng thuỷ phân tạo ra 2 phân tử glucose: C~12~H~22~O~11~ + H~2~O -(H^+, t°)-> 2 C~6~H~12~O~6~."),

        (6, "(SBT – KNTT) Saccharose và maltose đều tham gia phản ứng nào sau đây?",
         ["A. Phản ứng với thuốc thử Tollens.", "*B. Phản ứng thuỷ phân trong môi trường acid.*",
          "C. Phản ứng với dung dịch nước bromine.", "D. Phản ứng với Cu(OH)~2~ tạo kết tủa đỏ gạch."],
         "Cả saccharose và maltose đều là disaccharide nên đều tham gia phản ứng thuỷ phân trong môi trường acid khi đun nóng. Saccharose không có nhóm -OH hemiacetal nên không tráng bạc, không khử Cu(OH)~2~ thành Cu~2~O, không làm mất màu nước bromine."),

        (7, "(SBT – KNTT) Cellulose *không* được sử dụng trong ứng dụng nào sau đây?",
         ["*A. Sản xuất các thiết bị điện.*", "B. Nguyên liệu sản xuất ethanol và cellulose nitrate.",
          "C. Sản xuất giấy, tơ tự nhiên và sợi nhân tạo.", "D. Vật liệu gỗ xây dựng."],
         "Cellulose là chất không dẫn điện và dễ bắt cháy, không được dùng để chế tạo các thiết bị điện."),

        (8, "(SBT – KNTT) Nguyên liệu nào sau đây *không* phải là nguồn cung cấp tinh bột?",
         ["A. Củ và quả.", "B. Hạt ngũ cốc.", "*C. Sợi bông.*", "D. Gạo."],
         "Sợi bông có thành phần chủ yếu là cellulose (gần 98%), không phải nguồn cung cấp tinh bột."),

        (9, "(SGK – CD) Nhận xét nào dưới đây là *không đúng* khi nói về glucose và fructose?",
         ["A. Đều tạo được dung dịch màu xanh lam khi tác dụng với Cu(OH)~2~ trong môi trường kiềm.",
          "B. Đều tạo được kết tủa đỏ gạch Cu~2~O khi tác dụng với Cu(OH)~2~, đun nóng trong môi trường kiềm.",
          "*C. Đều làm mất màu nước bromine.*",
          "D. Đều xảy ra phản ứng tráng bạc khi tác dụng với thuốc thử Tollens."],
         "Nước bromine có tính acid nhẹ, fructose trong môi trường acid không chuyển hóa thành glucose nên không làm mất màu nước bromine. Chỉ có glucose chứa nhóm -CHO bị oxi hóa bởi nước bromine."),

        (10, "(SGK 12 – CTST) Carbohydrate nào dưới đây *không* có nhóm -OH hemiacetal hoặc nhóm -OH hemiketal?",
         ["A. Glucose.", "B. Fructose.", "*C. Saccharose.*", "D. Maltose."],
         "Phân tử saccharose liên kết bởi gốc α-glucose qua C1 và gốc β-fructose qua C2 (liên kết α-1,2-glycosidic), cả hai nhóm -OH hemiacetal và hemiketal đều đã tham gia liên kết nên saccharose không còn nhóm -OH hemiacetal/hemiketal tự do."),

        (11, "(SBT – CD) Polysaccharide mạch phân nhánh, có nhiều trong các loại ngũ cốc, thường được sử dụng làm lương thực là",
         ["A. cellulose.", "B. amylose.", "*C. amylopectin.*", "D. glycogen."],
         "Amylopectin là thành phần chính của tinh bột trong các loại ngũ cốc, có mạch phân nhánh nhờ liên kết α-1,6-glycosidic bên cạnh liên kết α-1,4-glycosidic."),

        (12, "Công thức cấu tạo dạng mạch vòng β-glucose là:",
         None, # Special image options
         "Trong dạng mạch vòng 6 cạnh của β-glucose (β-D-glucopyranose), nhóm -OH hemiacetal ở vị trí C1 hướng lên trên mặt phẳng vòng (cùng phía với nhóm -CH~2~OH ở C5)."),

        (13, "Chất T có các đặc điểm: (1) thuộc loại monosaccharide; (2) có nhiều trong quả nho chín; (3) tác dụng với nước bromine; (4) có phản ứng tráng gương. Chất T là",
         ["*A. glucose.*", "B. saccharose.", "C. fructose.", "D. cellulose."],
         "Chất T là monosaccharide có nhiều trong nho chín, tác dụng với nước bromine và tráng bạc được nên T là glucose."),

        (14, "Dãy các chất nào dưới đây đều phản ứng được với Cu(OH)~2~ ở điều kiện thường?",
         ["A. Ethylene glycol, glycerol và ethyl alcohol.",
          "*B. Glucose, glycerol và saccharose.*",
          "C. Glucose, glycerol và methyl acetate.",
          "D. Glycerol, glucose và ethyl acetate."],
         "Các chất polyalcohol có ít nhất hai nhóm -OH kề nhau phản ứng với Cu(OH)~2~ ở nhiệt độ thường tạo phức màu xanh lam gồm: glucose, glycerol, saccharose."),

        (15, "Phương trình: 6nCO~2~ + 5nH~2~O -(asmt, chlorophyll)-> (C~6~H~10~O~5~)~n~ + 6nO~2~ là phản ứng hoá học chính của quá trình nào sau đây?",
         ["A. quá trình hô hấp.", "*B. quá trình quang hợp.*", "C. quá trình khử.", "D. quá trình oxi hoá."],
         "Phương trình trên mô tả quá trình quang hợp của cây xanh dưới tác dụng của năng lượng ánh sáng mặt trời và diệp lục."),

        (16, "Thí nghiệm: Phản ứng của glucose với Cu(OH)~2~ khi đun nóng\n"
             "Bước 1: Cho khoảng 2 mL dung dịch NaOH 10% vào ống nghiệm.\n"
             "Bước 2: Sau đó, thêm khoảng 0,5 mL dung dịch CuSO~4~ 5% vào, lắc nhẹ.\n"
             "Bước 3: Cho tiếp khoảng 3 mL dung dịch glucose 2% vào ống nghiệm và lắc đều. Đun nóng ống nghiệm bằng ngọn lửa đèn cồn trong vài phút.\n"
             "Phát biểu nào sau đây *sai*?",
         ["A. Sản phẩm hữu cơ thu được sau bước 3 là sodium gluconate.",
          "B. Nếu thay dung dịch NaOH ở bước 2 bằng dung dịch KOH thì hiện tượng ở bước 3 vẫn tương tự.",
          "*C. Thí nghiệm trên chứng minh glucose có tính chất của polyalcohol.*",
          "D. Ở bước 3, nếu thay glucose bằng fructose thì hiện tượng xảy ra vẫn tương tự."],
         "Thí nghiệm đun nóng glucose với Cu(OH)~2~ trong môi trường kiềm tạo kết tủa Cu~2~O đỏ gạch chứng minh tính chất khử của nhóm aldehyde (-CHO), không phải tính chất của polyalcohol (tính chất polyalcohol thể hiện ở nhiệt độ thường tạo dd xanh lam)."),

        (17, "Cho các phát biểu sau:\n"
             "(a) Ở điều kiện thường, saccharose đều là chất rắn, dễ tan trong nước.\n"
             "(b) Saccharose bị hóa đen khi tiếp xúc với sulfuric acid đặc.\n"
             "(c) Trong công nghiệp dược phẩm, saccharose được dùng để pha chế thuốc.\n"
             "(d) Thủy phân hoàn toàn saccharose chỉ thu được glucose.\n"
             "Số phát biểu đúng là",
         ["*A. 3.*", "B. 1.", "C. 4.", "D. 2."],
         "Có 3 phát biểu đúng là (a), (b), (c).\n- (d) sai vì thủy phân hoàn toàn saccharose thu được cả glucose và fructose."),

        (18, "Cho các phát biểu sau:\n"
             "(a) Có thể dùng nước bromine để phân biệt glucose và fructose.\n"
             "(b) Trong môi trường acid, glucose và fructose có thể chuyển hoá lẫn nhau.\n"
             "(c) Có thể phân biệt glucose và fructose bằng phản ứng với dung dịch AgNO~3~ trong NH~3~.\n"
             "(d) Trong dung dịch, glucose và fructose đều hoà tan Cu(OH)~2~ ở nhiệt độ thường cho dung dịch màu xanh lam.\n"
             "(e) Trong dung dịch, fructose tồn tại chủ yếu ở dạng mạch hở.\n"
             "(g) Trong dung dịch, glucose tồn tại chủ yếu ở dạng vòng 6 cạnh (dạng α và β).\n"
             "Số phát biểu đúng là",
         ["A. 4.", "*B. 3.*", "C. 5.", "D. 2."],
         "Các phát biểu đúng gồm: (a), (d), (g) (tổng cộng 3 phát biểu đúng).\n"
         "- (b) sai: chuyển hóa lẫn nhau trong môi trường kiềm, không phải acid.\n"
         "- (c) sai: trong môi trường kiềm của AgNO~3~/NH~3~ fructose chuyển hóa thành glucose nên cả hai đều tráng bạc.\n"
         "- (e) sai: fructose tồn tại chủ yếu ở dạng mạch vòng (5 cạnh và 6 cạnh), dạng hở chiếm tỉ lệ rất nhỏ.")
    ]

    for item in p1_data:
        q_num, content, opts, sol = item
        add_question_title(doc, q_num, content)
        
        if q_num == 12:
            # Special image options layout for Cau 12
            img_a = os.path.join(IMG_DIR, "cau12_opt_A.png")
            img_b = os.path.join(IMG_DIR, "cau12_opt_B.png")
            img_c = os.path.join(IMG_DIR, "cau12_opt_C.png")
            img_d = os.path.join(IMG_DIR, "cau12_opt_D.png")

            p_row1 = doc.add_paragraph()
            p_row1.paragraph_format.space_before = Pt(2)
            p_row1.paragraph_format.space_after = Pt(2)
            p_row1.paragraph_format.tab_stops.clear_all()
            p_row1.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)

            r_a = p_row1.add_run("A. ")
            set_run_style(r_a, font_size=10.5, bold=True)
            if os.path.exists(img_a):
                p_row1.add_run().add_picture(img_a, width=Cm(3.8))
            
            p_row1.add_run("\t")
            r_b = p_row1.add_run("B. ")
            set_run_style(r_b, font_size=10.5, bold=True, color=COLOR_CORRECT if is_teacher else None)
            if os.path.exists(img_b):
                p_row1.add_run().add_picture(img_b, width=Cm(3.8))

            p_row2 = doc.add_paragraph()
            p_row2.paragraph_format.space_before = Pt(2)
            p_row2.paragraph_format.space_after = Pt(2)
            p_row2.paragraph_format.tab_stops.clear_all()
            p_row2.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)

            r_c = p_row2.add_run("C. ")
            set_run_style(r_c, font_size=10.5, bold=True)
            if os.path.exists(img_c):
                p_row2.add_run().add_picture(img_c, width=Cm(3.8))

            p_row2.add_run("\t")
            r_d = p_row2.add_run("D. ")
            set_run_style(r_d, font_size=10.5, bold=True)
            if os.path.exists(img_d):
                p_row2.add_run().add_picture(img_d, width=Cm(3.8))
        else:
            render_options_no_table(doc, opts, is_teacher=is_teacher)

        if is_teacher:
            add_solution_no_table(doc, sol)

    # =========================================================================
    # PHẦN II
    # =========================================================================
    add_part_heading(doc, "PHẦN II. Câu trắc nghiệm đúng sai.",
                     "Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")

    # Câu 1
    add_question_title(doc, 1,
                       "Một học sinh tiến hành làm thí nghiệm tráng bạc của glucose theo hai bước sau:\n"
                       "Bước 1: Cho vào một ống nghiệm sạch 1 mL dung dịch AgNO~3~ 1%, sau đó nhỏ tiếp từng giọt dung dịch NH~3~ vào cho đến khi kết tủa vừa xuất hiện lại tan hết.\n"
                       "Bước 2: Thêm tiếp 1 mL dung dịch glucose 1% vào ống nghiệm trên rồi đun nóng nhẹ.")
    c1_stmts = [
        ("a) Phản ứng xảy ra ở bước 2 là: CH~2~OH-[CHOH]~4~-CHO + 2AgNO~3~ + 3NH~3~ + H~2~O -(t°)-> CH~2~OH-[CHOH]~4~-COONH~4~ + 2Ag + 2NH~4~NO~3~.", "Đúng"),
        ("b) Hiện tượng xảy ra ở bước 2 là xuất hiện kết tủa màu đen.", "Sai"),
        ("c) Ở bước 2 chất oxi hóa là glucose, chất khử là AgNO~3~.", "Sai"),
        ("d) Trong công nghiệp, người ta dùng sucrose làm nguyên liệu để tráng ruột phích mà không dùng glucose là vì sucrose có lượng nhiều trong tự nhiên, giá thành thấp đồng thời khi thủy phân sinh ra glucose và fructose đều có phản ứng tráng bạc còn glucose có lượng ít trong tự nhiên, giá thành cao.", "Đúng")
    ]
    for stmt, ans in c1_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
                              "- a) ĐÚNG: Đây là phương trình phân tử biểu diễn phản ứng tráng bạc của glucose.\n"
                              "- b) SAI: Hiện tượng là tạo lớp bạc sáng bóng như gương bám vào thành ống nghiệm, không phải kết tủa đen.\n"
                              "- c) SAI: Ag^+ nhận electron giảm số oxi hóa nên AgNO~3~ là chất oxi hóa; glucose bị oxi hóa nên là chất khử.\n"
                              "- d) ĐÚNG: Trong công nghiệp tráng ruột phích dùng saccharose (sucrose) rẻ tiền, dễ kiếm và thủy phân cho cả 2 monosaccharide đều tráng gương được.")

    # Câu 2
    add_question_title(doc, 2, "Hợp chất saccharose và maltose là chất rắn có vị ngọt, dễ tan trong nước.")
    img_sugar = os.path.join(IMG_DIR, "p2_cau2_sugar.png")
    if os.path.exists(img_sugar):
        p_img = doc.add_paragraph()
        p_img.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(2)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.add_run().add_picture(img_sugar, width=Cm(9.5))

    c2_stmts = [
        ("a) Đều thuộc loại carbohydrate.", "Đúng"),
        ("b) Là hai chất đồng phân.", "Đúng"),
        ("c) Khi bị thủy phân hoàn toàn, mỗi phân tử đều tạo thành hai phân tử glucose.", "Sai"),
        ("d) Đều thuộc loại disaccharide.", "Đúng")
    ]
    for stmt, ans in c2_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
                              "- a) ĐÚNG: Cả saccharose và maltose đều thuộc nhóm disaccharide trong carbohydrate.\n"
                              "- b) ĐÚNG: Cùng có CTPT là C~12~H~22~O~11~ nhưng cấu tạo khác nhau nên là đồng phân.\n"
                              "- c) SAI: Maltose thủy phân tạo 2 glucose; còn saccharose thủy phân tạo 1 glucose và 1 fructose.\n"
                              "- d) ĐÚNG: Đều do 2 gốc monosaccharide liên kết tạo nên.")

    # Câu 3
    add_question_title(doc, 3, "Tinh bột là một polysaccharide gồm nhiều gốc α-glucose liên kết với nhau tạo ra hai dạng mạch amylose và amylopectin, dưới đây là cấu tạo của một trong hai dạng này.")
    img_amylose = os.path.join(IMG_DIR, "p2_cau3_amylose.png")
    if os.path.exists(img_amylose):
        p_img = doc.add_paragraph()
        p_img.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(2)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.add_run().add_picture(img_amylose, width=Cm(12.0))

    c3_stmts = [
        ("a) Đoạn mạch trên là cấu tạo của dạng amylose trong tinh bột.", "Đúng"),
        ("b) Các gốc α-glucose trong đoạn mạch trên liên kết với nhau bởi liên kết α-1,6-glycoside.", "Sai"),
        ("c) Công thức mỗi gốc α-glucose là C~6~H~10~O~5~.", "Đúng"),
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
                              "- a) ĐÚNG: Đoạn mạch thẳng gồm các gốc α-glucose nối với nhau là amylose.\n"
                              "- b) SAI: Các gốc α-glucose liên kết với nhau bằng liên kết α-1,4-glycosidic.\n"
                              "- c) ĐÚNG: Công thức mỗi mắt xích (gốc) là -C~6~H~10~O~5~-.\n"
                              "- d) SAI: Mạch amylose là mạch thẳng không phân nhánh.")

    # Câu 4
    add_question_title(doc, 4,
                       "Tính tan của cellulose trong nước Schweizer\n"
                       "*Chuẩn bị:*\n"
                       "Hoá chất: cellulose (bông), dung dịch CuSO~4~ 1 M, dung dịch NaOH 20%, dung dịch NH~3~ đặc.\n"
                       "Dụng cụ: giấy lọc, cốc thuỷ tinh 250 mL, đũa thuỷ tinh, ống hút.\n"
                       "*Tiến hành:*\n"
                       "Bước 1: Cho khoảng 50 mL dung dịch CuSO~4~ 1 M vào cốc 250 mL. Thêm 20 mL dung dịch NaOH 20% vào, khuấy đều.\n"
                       "Bước 2: Lọc tách kết tủa, cho vào cốc thuỷ tinh 250 mL. Thêm khoảng 50 mL dung dịch NH~3~ đặc, khuấy đều đến khi kết tủa tan hết thu được nước Schweizer.\n"
                       "Bước 3: Thêm một lượng nhỏ bông vào khoảng 30 mL nước Schweizer và khuấy đều trong khoảng 3 phút.")
    c4_stmts = [
        ("a) Sau bước 1, sản phẩm thu được là kết tủa Cu(OH)~2~ có màu xanh.", "Đúng"),
        ("b) Ở bước 2 xảy ra phương trình tạo nước Schweizer như sau:\n"
         "Cu(OH)~2~ + 4NH~3~ -> [Cu(NH~3~)~4~](OH)~2~ (Tetraamminecopper(II) hydroxide).", "Đúng"),
        ("c) Sau bước 3, dung dịch thu được có sự phân lớp.", "Sai"),
        ("d) Thí nghiệm trên chứng minh trong cellulose có 3 nhóm OH tự do.", "Sai")
    ]
    for stmt, ans in c4_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
                              "- a) ĐÚNG: CuSO~4~ + 2NaOH -> Cu(OH)~2~↓ (xanh lam) + Na~2~SO~4~.\n"
                              "- b) ĐÚNG: Cu(OH)~2~ tan trong dung dịch NH~3~ đặc tạo phức chất [Cu(NH~3~)~4~](OH)~2~ gọi là nước Schweizer.\n"
                              "- c) SAI: Cellulose tan trong nước Schweizer tạo dung dịch đồng nhất nhớt, không bị phân lớp.\n"
                              "- d) SAI: Thí nghiệm chứng minh tính tan của cellulose trong dung dịch Schweizer, không xác định số lượng 3 nhóm -OH tự do.")

    # =========================================================================
    # PHẦN III
    # =========================================================================
    add_part_heading(doc, "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
                     "Thí sinh trả lời từ câu 1 đến câu 6.")

    p3_data = [
        (1, "Cho dãy các chất sau: saccharose, glucose, fructose, cellulose, tristearin, methyl acetate. Số chất tham gia phản ứng thuỷ phân là bao nhiêu?",
         "4",
         "Các chất bị thủy phân gồm disaccharide, polysaccharide và ester:\n"
         "- saccharose -> glucose + fructose\n"
         "- cellulose -> glucose\n"
         "- tristearin -> glycerol + muối stearate\n"
         "- methyl acetate -> CH~3~COOH + CH~3~OH\n"
         "(Glucose và fructose là monosaccharide không bị thủy phân).\n"
         "➔ Có 4 chất tham gia phản ứng thủy phân."),

        (2, "Cho các chất: saccharose, glucose, fructose, ethyl formate, formic acid và acetic aldehyde. Trong các chất trên, số chất vừa có khả năng tham gia phản ứng tráng bạc vừa có khả năng phản ứng với Cu(OH)~2~ ở điều kiện thường là",
         "3",
         "- Vừa tráng bạc vừa hòa tan Cu(OH)~2~ ở nhiệt độ thường:\n"
         "  + Glucose: polyalcohol có -CHO\n"
         "  + Fructose: polyalcohol, chuyển thành glucose trong môi trường kiềm\n"
         "  + Formic acid: carboxylic acid hòa tan Cu(OH)~2~, có nhóm -CHO tráng bạc\n"
         "(Saccharose không tráng bạc; ethyl formate và acetic aldehyde không phản ứng Cu(OH)~2~ ở đk thường).\n"
         "➔ Đáp số: 3 chất (glucose, fructose, formic acid)."),

        (3, "Cho dãy các chất sau: glucose, fructose, saccharose, cellulose, tinh bột. Số chất khi bị thủy phân trong môi trường acid chỉ tạo glucose là",
         "2",
         "- Cellulose -(H^+, t°)-> n glucose\n"
         "- Tinh bột -(H^+, t°)-> n glucose\n"
         "- Saccharose thủy phân tạo glucose và fructose (loại)\n"
         "- Glucose, fructose không bị thủy phân (loại)\n"
         "➔ Đáp số: 2 chất (cellulose, tinh bột)."),

        (4, "Để tráng một số lượng gương soi có diện tích bề mặt 0,35 m^2 với độ dày 0,1 µm người ta đun nóng dung dịch chứa 26,77 gam glucose với một lượng dung dịch silver nitrate trong ammonia. Biết khối lượng riêng của silver là 10,49 g/cm^3, hiệu suất phản ứng tráng gương là 80% (tính theo glucose). Số lượng gương soi tối đa sản xuất được khoảng bao nhiêu?",
         "70",
         "- n~glucose~ = 26,77 / 180 mol.\n"
         "- Khối lượng Ag sinh ra: m~Ag~ = (2 · 26,77 / 180) · 108 · 80% = 25,7005 gam.\n"
         "- Thể tích Ag tạo thành: V~Ag~ = m~Ag~ / D = 25,7005 / 10,49 ≈ 2,45 cm^3.\n"
         "- Thể tích Ag cần cho 1 tấm gương:\n"
         "  S = 0,35 m^2 = 3500 cm^2; d = 0,1 µm = 10^-5 cm.\n"
         "  V~1 tấm~ = 3500 · 10^-5 = 0,035 cm^3.\n"
         "- Số lượng gương soi tối đa = 2,45 / 0,035 = 70 tấm.\n"
         "➔ Đáp số: 70."),

        (5, "Tại một nhà máy rượu, cứ 10 tấn tinh bột (chứa 6,85% tạp chất trơ) sẽ sản xuất được 7,21 m^3 ethanol 40° (cho khối lượng riêng của ethanol nguyên chất là 0,789 g/cm^3). Hiệu suất của quá trình sản xuất là H%. Xác định giá trị của H (làm tròn một chữ số thập phân).",
         "43",
         "- m~tinh bột thực tế~ = 10 · (100% - 6,85%) = 9,315 tấn.\n"
         "- V~ethanol nguyên chất~ = 7,21 · 0,40 = 2,884 m^3 = 2884 lít.\n"
         "- m~ethanol~ = 2,884 · 0,789 = 2,275476 tấn.\n"
         "- Sơ đồ: (C~6~H~10~O~5~)~n~ -> n C~6~H~12~O~6~ -> 2n C~2~H~5~OH\n"
         "  m~tinh bột lý thuyết~ = 2,275476 · 162 / 92 ≈ 4,006816 tấn.\n"
         "- Hiệu suất H = (4,006816 / 9,315) · 100% ≈ 43,01% ≈ 43% (hay 43,0%).\n"
         "➔ Đáp số: 43."),

        (6, "(SBT – CD) Theo Tiêu chuẩn Việt Nam TCVN 7624:2007, khi chế tạo gương, chiều dày lớp bạc phủ trên bề mặt tấm kính (quy ra tổng lượng bạc trên một đơn vị m^2 kính) phải đạt tối thiểu 0,7 g.m^-2. Một công ty cần sản xuất 10000 m^2 gương có độ dày lớp bạc phủ ở mức 0,72 g.m^-2. Biết rằng lớp bạc được tạo thành qua phản ứng giữa silver nitrate và glucose trong điều kiện thích hợp với hiệu suất phản ứng 90%. Khối lượng glucose cần sử dụng để sản xuất 10000 m^2 gương trên là bao nhiêu kg? (Làm tròn đến chữ số thập phân thứ hai).",
         "6,67",
         "- Khối lượng bạc cần phủ trên 10000 m^2 gương: m~Ag~ = 0,72 · 10000 = 7200 g = 7,2 kg.\n"
         "- Phản ứng: CH~2~OH[CHOH]~4~CHO + 2AgNO~3~ + 3NH~3~ + H~2~O -> ... + 2Ag + ...\n"
         "- Theo phương trình: n~glucose~ = 1/2 · n~Ag~ = 1/2 · (7,2 / 108) kmol.\n"
         "- Vì hiệu suất phản ứng 90%, lượng glucose cần dùng là:\n"
         "  m~glucose~ = (7,2 / 108) · (180 / 2) · (100 / 90) = 6,67 kg.\n"
         "  (Lượng AgNO~3~ cần dùng là: (7,2 / 108) · 170 · (100 / 90) = 12,59 kg).\n"
         "➔ Đáp số: 6,67.")
    ]

    for item in p3_data:
        q_num, content, ans, sol = item
        add_question_title(doc, q_num, content)

        if q_num == 5:
            img_wine = os.path.join(IMG_DIR, "p3_cau5_wine.png")
            if os.path.exists(img_wine):
                p_img = doc.add_paragraph()
                p_img.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_wine, width=Cm(9.0))
                p_cap = doc.add_paragraph()
                p_cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cap.paragraph_format.space_before = Pt(0)
                p_cap.paragraph_format.space_after = Pt(2)
                r_cap = p_cap.add_run("Dây chuyền sản xuất rượu (alcohol)")
                set_run_style(r_cap, font_size=9.5, italic=True, color=COLOR_HF_RED)

        if not is_teacher:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.5)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(3)
            r_ans_lbl = p_ans.add_run("Đáp số: .....................................................")
            set_run_style(r_ans_lbl, font_size=10, italic=True)
        else:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.5)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(1)
            r_ans_lbl = p_ans.add_run(f"Đáp số: {ans}")
            set_run_style(r_ans_lbl, font_size=10.5, bold=True, color=COLOR_CORRECT)
            add_solution_no_table(doc, sol)

    # Save
    out_name = "De_Kiem_Tra_Hoa_12_De04_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_12_De04.docx"
    out_path = os.path.join(OUT_DIR, out_name)
    doc.save(out_path)
    print(f"Saved: {out_path}")

if __name__ == "__main__":
    build_exam_de04(is_teacher=False)
    build_exam_de04(is_teacher=True)
    print("Done creating both Word documents for De 04!")
