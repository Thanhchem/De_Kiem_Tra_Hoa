# -*- coding: utf-8 -*-
"""
Script tạo 2 file Word (.docx) chuẩn Form Đề Kiểm Tra cho:
ĐỀ SỐ 16 - CHƯƠNG 6: ĐẠI CƯƠNG KIM LOẠI (HÓA HỌC 12)
Mã đề: 116
Bao gồm:
1. Bản Học sinh: De_Kiem_Tra_Hoa_12_De16.docx
2. Bản Giáo viên: De_Kiem_Tra_Hoa_12_De16_GiaoVien.docx (Có đáp án đỏ, lời giải chi tiết xanh, bảng đáp án)
"""

import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"

COLOR_CORAL_RED = RGBColor(211, 47, 47)      # #D32F2F
COLOR_SAGE_GREEN = "8BC390"                 # #8BC390
COLOR_DEEP_BLUE = RGBColor(21, 101, 192)    # #1565C0
COLOR_DARK_GREEN = RGBColor(46, 125, 50)    # #2E7D32
COLOR_BLACK = RGBColor(0, 0, 0)

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_header_footer(doc):
    for s in doc.sections:
        s.top_margin = Inches(0.59)     # 1.5 cm
        s.bottom_margin = Inches(0.59)  # 1.5 cm
        s.left_margin = Inches(0.59)    # 1.5 cm
        s.right_margin = Inches(0.59)   # 1.5 cm
        s.header_distance = Inches(0.24) # 0.6 cm
        s.footer_distance = Inches(0.24) # 0.6 cm

        # --- HEADER ---
        header = s.header
        header.is_linked_to_previous = False
        hp = header.paragraphs[0]
        hp.text = ""
        hp.paragraph_format.space_before = Pt(0)
        hp.paragraph_format.space_after = Pt(0)

        tbl_h = header.add_table(rows=1, cols=2, width=Inches(7.08)) # 18.0 cm
        tbl_h.alignment = WD_TABLE_ALIGNMENT.CENTER
        c0, c1 = tbl_h.rows[0].cells
        c0.width = Inches(3.54)
        c1.width = Inches(3.54)
        set_cell_margins(c0, 60, 60, 0, 0)
        set_cell_margins(c1, 60, 60, 0, 0)

        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run("Tài liệu lưu hành nội bộ")
        r0.font.name = "Times New Roman"
        r0.font.size = Pt(10)
        r0.font.italic = True
        r0.font.color.rgb = COLOR_CORAL_RED

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r1_a = p1.add_run("Thầy ")
        r1_a.font.name = "Times New Roman"
        r1_a.font.size = Pt(10)
        r1_a.font.color.rgb = COLOR_CORAL_RED
        r1_b = p1.add_run("TRẦN VĂN THẠNH")
        r1_b.font.name = "Times New Roman"
        r1_b.font.size = Pt(10)
        r1_b.font.bold = True
        r1_b.font.color.rgb = COLOR_CORAL_RED
        r1_c = p1.add_run(" - 0777.470.803")
        r1_c.font.name = "Times New Roman"
        r1_c.font.size = Pt(10)
        r1_c.font.color.rgb = COLOR_CORAL_RED

        tblPr_h = tbl_h._tbl.tblPr
        tblBorders_h = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="{COLOR_SAGE_GREEN}"/>'
            f'</w:tblBorders>'
        )
        tblPr_h.append(tblBorders_h)

        # --- FOOTER ---
        footer = s.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.text = ""
        fp.paragraph_format.space_before = Pt(0)
        fp.paragraph_format.space_after = Pt(0)

        tbl_f = footer.add_table(rows=1, cols=2, width=Inches(7.08))
        tbl_f.alignment = WD_TABLE_ALIGNMENT.CENTER
        fc0, fc1 = tbl_f.rows[0].cells
        fc0.width = Inches(5.31)
        fc1.width = Inches(1.77)
        set_cell_margins(fc0, 60, 60, 0, 0)
        set_cell_margins(fc1, 60, 60, 0, 0)

        fp0 = fc0.paragraphs[0]
        fp0.paragraph_format.space_before = Pt(0)
        fp0.paragraph_format.space_after = Pt(0)
        fr0 = fp0.add_run("CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế   |   CS2: 24 Đặng Thái Thân, TP Huế")
        fr0.font.name = "Times New Roman"
        fr0.font.size = Pt(9.5)
        fr0.font.italic = True
        fr0.font.color.rgb = COLOR_CORAL_RED

        fp1 = fc1.paragraphs[0]
        fp1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        fp1.paragraph_format.space_before = Pt(0)
        fp1.paragraph_format.space_after = Pt(0)
        fr1 = fp1.add_run("Trang ")
        fr1.font.name = "Times New Roman"
        fr1.font.size = Pt(10)
        fr1.font.bold = True
        fr1.font.color.rgb = COLOR_CORAL_RED

        fldSimple = OxmlElement('w:fldSimple')
        fldSimple.set(qn('w:instr'), 'PAGE')
        fp1._p.append(fldSimple)

        tblPr_f = tbl_f._tbl.tblPr
        tblBorders_f = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="6" w:space="0" w:color="{COLOR_SAGE_GREEN}"/>'
            f'</w:tblBorders>'
        )
        tblPr_f.append(tblBorders_f)

def add_chem_run(paragraph, text, bold=False, italic=False, color=None, font_size=11):
    tokens = re.split(r'(\~[0-9\+\-mnx]+\~|\^[0-9\+\-mnx]+\^)', text)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith('~') and tok.endswith('~'):
            content = tok[1:-1]
            r = paragraph.add_run(content)
            r.font.subscript = True
        elif tok.startswith('^') and tok.endswith('^'):
            content = tok[1:-1]
            r = paragraph.add_run(content)
            r.font.superscript = True
        else:
            r = paragraph.add_run(tok)
            
        r.font.name = "Times New Roman"
        r.font.size = Pt(font_size)
        r.font.bold = bold
        r.font.italic = italic
        if color:
            r.font.color.rgb = color

def create_title_block(doc, is_teacher=False):
    lines = [
        ("SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ", "ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 6 (ĐẠI CƯƠNG KIM LOẠI)", True, True),
        ("TRƯỜNG THPT HAI BÀ TRƯNG", "MÔN: HOÁ HỌC – LỚP 12", True, True),
        ("ĐỀ CHÍNH THỨC", "Thời gian làm bài: 50 phút (không kể thời gian phát đề)", True, False),
        ("(Đề thi có 04 trang)", "Mã đề thi: 116", False, True)
    ]
    for left_text, right_text, left_bold, right_bold in lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.tab_stops.add_tab_stop(Inches(7.08), WD_TAB_ALIGNMENT.RIGHT)

        r_left = p.add_run(left_text)
        r_left.font.name = "Times New Roman"
        r_left.font.size = Pt(10.5)
        r_left.font.bold = left_bold
        if not left_bold:
            r_left.font.italic = True

        p.add_run("\t")

        r_right = p.add_run(right_text)
        r_right.font.name = "Times New Roman"
        r_right.font.size = Pt(10.5)
        r_right.font.bold = right_bold
        if not right_bold:
            r_right.font.italic = True
        if "ĐỀ KIỂM TRA" in right_text:
            r_right.font.color.rgb = COLOR_DARK_GREEN

    if is_teacher:
        p_gv = doc.add_paragraph()
        p_gv.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_gv.paragraph_format.space_before = Pt(4)
        p_gv.paragraph_format.space_after = Pt(4)
        r_gv = p_gv.add_run("(BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)")
        r_gv.font.name = "Times New Roman"
        r_gv.font.size = Pt(12)
        r_gv.font.bold = True
        r_gv.font.color.rgb = COLOR_CORAL_RED

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(3)
    p_info.paragraph_format.space_after = Pt(2)
    r_info = p_info.add_run("Họ, tên thí sinh: ....................................................................................... Lớp: ........................ SBD: ....................")
    r_info.font.name = "Times New Roman"
    r_info.font.size = Pt(10)
    r_info.font.italic = True

    p_ntk = doc.add_paragraph()
    p_ntk.paragraph_format.space_before = Pt(0)
    p_ntk.paragraph_format.space_after = Pt(5)
    r_ntk = p_ntk.add_run("(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Cr = 52; Fe = 56; Cu = 64; Zn = 65; Ag = 108; Sn = 119; Pt = 195; Au = 197; Pb = 207; F = 96500 C/mol).")
    r_ntk.font.name = "Times New Roman"
    r_ntk.font.size = Pt(9.5)
    r_ntk.font.italic = True

def add_section_header(doc, title, subtitle):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(title)
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    r.font.bold = True

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(4)
    p_sub.paragraph_format.line_spacing = 1.15
    r_sub = p_sub.add_run(subtitle)
    r_sub.font.name = "Times New Roman"
    r_sub.font.size = Pt(10.5)
    r_sub.font.italic = True

def add_loigiai(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.28) # 0.7 cm
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15

    r_lbl = p.add_run("Lời giải chi tiết:\n")
    r_lbl.font.name = "Times New Roman"
    r_lbl.font.size = Pt(10)
    r_lbl.font.bold = True
    r_lbl.font.color.rgb = COLOR_DEEP_BLUE

    for line in text.strip().split("\n"):
        p_line = doc.add_paragraph()
        p_line.paragraph_format.left_indent = Inches(0.28)
        p_line.paragraph_format.space_before = Pt(0)
        p_line.paragraph_format.space_after = Pt(2)
        p_line.paragraph_format.line_spacing = 1.15
        
        is_bold = line.startswith("Chọn ") or line.startswith("Đáp số:") or line.startswith("=>")
        add_chem_run(p_line, line, bold=is_bold, color=COLOR_DEEP_BLUE, font_size=10)

def add_mcq_options(doc, options, correct_key=None, is_teacher=False, layout="1_row"):
    """
    Tạo các phương án A, B, C, D bằng Paragraph thuần tuý kết hợp Tab stops, TUYỆT ĐỐI KHÔNG DÙNG BẢNG.
    """
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15

    if layout == "1_row":
        p.paragraph_format.tab_stops.add_tab_stop(Inches(1.77), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(3.54), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(5.31), WD_TAB_ALIGNMENT.LEFT)
        for i, (k, txt) in enumerate(options):
            if i > 0:
                p.add_run("\t")
            is_c = (k == correct_key) and is_teacher
            col = COLOR_CORAL_RED if is_c else None
            add_chem_run(p, f"{k}. ", bold=True, color=col)
            add_chem_run(p, txt, bold=is_c, color=col)

    elif layout == "2_rows":
        p.paragraph_format.tab_stops.add_tab_stop(Inches(3.54), WD_TAB_ALIGNMENT.LEFT)
        for i in range(2):
            k, txt = options[i]
            if i > 0:
                p.add_run("\t")
            is_c = (k == correct_key) and is_teacher
            col = COLOR_CORAL_RED if is_c else None
            add_chem_run(p, f"{k}. ", bold=True, color=col)
            add_chem_run(p, txt, bold=is_c, color=col)

        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(3)
        p2.paragraph_format.line_spacing = 1.15
        p2.paragraph_format.tab_stops.add_tab_stop(Inches(3.54), WD_TAB_ALIGNMENT.LEFT)
        for i in range(2, 4):
            k, txt = options[i]
            if i > 2:
                p2.add_run("\t")
            is_c = (k == correct_key) and is_teacher
            col = COLOR_CORAL_RED if is_c else None
            add_chem_run(p2, f"{k}. ", bold=True, color=col)
            add_chem_run(p2, txt, bold=is_c, color=col)

    else: # 4_rows
        for i, (k, txt) in enumerate(options):
            if i > 0:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
            is_c = (k == correct_key) and is_teacher
            col = COLOR_CORAL_RED if is_c else None
            add_chem_run(p, f"{k}. ", bold=True, color=col)
            add_chem_run(p, txt, bold=is_c, color=col)

def generate_exam_docx(is_teacher=False):
    doc = docx.Document()
    add_header_footer(doc)
    create_title_block(doc, is_teacher=is_teacher)

    # ==============================================================
    # PHẦN I
    # ==============================================================
    add_section_header(
        doc,
        "PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    p1_data = [
        (
            1,
            "Phát biểu nào sau đây về tinh thể kim loại là đúng?",
            [
                ("A", "Chứa các cation kim loại và các electron hoá trị. Chúng cùng chuyển động tự do trong tinh thể."),
                ("B", "Chứa các cation kim loại và nguyên tử kim loại có tương tác tĩnh điện với nhau."),
                ("C", "Chứa các cation kim loại sắp xếp theo trật tự nhất định và các electron hoá trị chuyển động tự do."),
                ("D", "Bao gồm các cation kim loại và electron hoá trị tự do. Giữa chúng hình thành liên kết ion giúp tinh thể kim loại có cấu trúc bền vững.")
            ],
            "C",
            "4_rows",
            "Mạng tinh thể kim loại gồm các cation kim loại nằm ở các nút mạng tinh thể sắp xếp theo một trật tự xác định và các electron hoá trị tự do chuyển động trong toàn bộ mạng tinh thể.\n"
            "Chọn C."
        ),
        (
            2,
            "Lực liên kết kim loại gây ra bởi",
            [
                ("A", "tương tác tĩnh điện giữa cation kim loại và electron hoá trị tự do trong tinh thể."),
                ("B", "sự góp chung các electron hoá trị giữa các nguyên tử kim loại trong tinh thể."),
                ("C", "tương tác van der Waals giữa các nguyên tử kim loại trong tinh thể."),
                ("D", "sự hình thành liên kết cho - nhận giữa các nguyên tử kim loại trong tinh thể.")
            ],
            "A",
            "4_rows",
            "Liên kết kim loại là liên kết được hình thành do lực hút tĩnh điện giữa các cation kim loại ở nút mạng tinh thể và các electron hoá trị tự do chuyển động trong tinh thể.\n"
            "Chọn A."
        ),
        (
            3,
            "Tính chất vật lí chung của kim loại được quyết định bởi",
            [
                ("A", "bán kính nguyên tử kim loại."),
                ("B", "các electron hoá trị tự do trong tinh thể kim loại."),
                ("C", "kiểu tinh thể kim loại."),
                ("D", "độ âm điện của nguyên tử kim loại.")
            ],
            "B",
            "4_rows",
            "Các tính chất vật lí chung của kim loại (tính dẻo, tính dẫn điện, tính dẫn nhiệt và ánh kim) đều do các electron tự do trong mạng tinh thể kim loại gây ra.\n"
            "Chọn B."
        ),
        (
            4,
            "Hai kim loại thường được dùng làm dây dẫn điện là",
            [
                ("A", "đồng (copper) và nhôm (aluminium)."),
                ("B", "đồng (copper) và tungsten."),
                ("C", "nhôm (aluminium) và sắt (iron)."),
                ("D", "kẽm (zinc) và nhôm (aluminium).")
            ],
            "A",
            "2_rows",
            "Đồng và nhôm là hai kim loại có độ dẫn điện tốt, giá thành hợp lí và độ bền cao nên được sử dụng phổ biến nhất để chế tạo dây dẫn điện.\n"
            "Chọn A."
        ),
        (
            5,
            "Trong môi trường có pH = 7, các kim loại tan trong nước thường có giá trị thế điện cực chuẩn",
            [
                ("A", "lớn hơn -0,413 V."),
                ("B", "nhỏ hơn -0,413 V."),
                ("C", "lớn hơn -0,826 V."),
                ("D", "nhỏ hơn -0,826 V.")
            ],
            "B",
            "2_rows",
            "Ở pH = 7 (môi trường trung tính), thế điện cực của quá trình khử nước (2H~2~O + 2e ⇌ H~2~ + 2OH^-^) là E = -0,413 V.\n"
            "Để kim loại khử được nước giải phóng H~2~ ở điều kiện này, thế điện cực chuẩn của cặp oxi hoá - khử kim loại phải nhỏ hơn -0,413 V.\n"
            "Chọn B."
        ),
        (
            6,
            "Trong quá trình tái chế kim loại, ở công đoạn phân loại phế liệu, có thể dùng nam châm lớn để tách riêng phế liệu",
            [
                ("A", "nhôm (aluminium)."),
                ("B", "thép."),
                ("C", "đồng (copper)."),
                ("D", "kẽm (zinc).")
            ],
            "B",
            "2_rows",
            "Thép có thành phần chính là sắt (Fe) có tính nhiễm từ mạnh (bị nam châm hút), do đó có thể dùng nam châm điện lớn để tách riêng phế liệu thép khỏi các kim loại khác.\n"
            "Chọn B."
        ),
        (
            7,
            "Xỉ xuất hiện trong quá trình tách kim loại, luyện kim và tinh luyện để tái chế kim loại. Xỉ là hỗn hợp nhiều chất, hỗn hợp này thường",
            [
                ("A", "có khối lượng riêng nhỏ hơn so với khối lượng riêng của kim loại cần tách hoặc cần tái chế."),
                ("B", "dễ chảy lỏng."),
                ("C", "dễ bay hơi."),
                ("D", "có nhiều độc tính nên cần phải được chôn lấp sau khi được tháo ra khỏi lò.")
            ],
            "A",
            "4_rows",
            "Xỉ lò (như CaSiO~3~,...) có khối lượng riêng nhỏ hơn kim loại lỏng nên nổi lên bề mặt, giúp dễ dàng gạt bỏ và bảo vệ kim loại lỏng không bị oxi hoá bởi không khí.\n"
            "Chọn A."
        ),
        (
            8,
            "Hợp kim là",
            [
                ("A", "một kim loại tinh khiết."),
                ("B", "hỗn hợp các kim loại có thành phần tùy ý."),
                ("C", "hỗn hợp của kim loại nền với kim loại khác hoặc phi kim, có thành phần xác định."),
                ("D", "hỗn hợp hai phi kim.")
            ],
            "C",
            "4_rows",
            "Hợp kim là vật liệu kim loại chứa một kim loại cơ bản (kim loại nền) và một số kim loại khác hoặc phi kim, có thành phần xác định nhằm cải thiện cơ tính.\n"
            "Chọn C."
        ),
        (
            9,
            "Kim loại cơ bản trong dural và thép 304 lần lượt là",
            [
                ("A", "nhôm và magnesium."),
                ("B", "nhôm và sắt."),
                ("C", "magnesium và sắt."),
                ("D", "kẽm và sắt.")
            ],
            "B",
            "2_rows",
            "Dural (duralumin) có kim loại cơ bản là nhôm (Al > 90%); Thép 304 (inox 304) có kim loại cơ bản là sắt (Fe).\n"
            "Chọn B."
        ),
        (
            10,
            "Thành phần quan trọng của inox là",
            [
                ("A", "sắt (iron), chromium, carbon."),
                ("B", "nhôm (aluminium), magnesium."),
                ("C", "đồng (copper), kẽm (zinc)."),
                ("D", "sắt (iron), magnesium, carbon.")
            ],
            "A",
            "2_rows",
            "Inox (thép không gỉ) có thành phần chính là sắt kết hợp với chromium (tối thiểu 10,5% tạo lớp màng thụ động bảo vệ chống gỉ), carbon và nickel.\n"
            "Chọn A."
        ),
        (
            11,
            "Sự ăn mòn kim loại",
            [
                ("A", "là sự phá hủy kim loại hoặc hợp kim do các chất điện li trong môi trường."),
                ("B", "chính là sự ăn mòn điện hoá."),
                ("C", "diễn ra trong dung dịch chất điện li."),
                ("D", "là sự phá huỷ kim loại hoặc hợp kim do tác dụng của các chất trong môi trường, trong đó kim loại bị oxi hoá.")
            ],
            "D",
            "4_rows",
            "Định nghĩa: Sự ăn mòn kim loại là sự phá huỷ kim loại hoặc hợp kim do tác dụng hoá học hoặc điện hoá của các chất trong môi trường, trong đó kim loại bị oxi hoá.\n"
            "Chọn D."
        ),
        (
            12,
            "Phát biểu nào dưới đây là không đúng về chống ăn mòn kim loại?",
            [
                ("A", "Phủ lên bề mặt của kim loại một lớp sơn, dầu, mỡ, chất dẻo hoặc tráng, mạ bằng một kim loại khác là những cách chống ăn mòn kim loại theo phương pháp phủ bề mặt."),
                ("B", "Nối kim loại cần bảo vệ với một kim loại hoạt động hoá học mạnh hơn là cách chống ăn mòn kim loại theo phương pháp điện hoá."),
                ("C", "Quấn một dây kẽm quanh đinh sắt là cách để chống ăn mòn kẽm trong môi trường có chất điện li."),
                ("D", "Để tránh ăn mòn đồ trang sức bằng Ag thì người ta thường phải đánh sạch hoặc rửa sạch trong xà phòng.")
            ],
            "C",
            "4_rows",
            "Quấn dây kẽm quanh đinh sắt là phương pháp điện hoá để bảo vệ đinh sắt (kẽm bị ăn mòn thay cho sắt), không phải để chống ăn mòn kẽm. Nhận định C sai.\n"
            "Chọn C."
        ),
        (
            13,
            "Ở phản ứng tách kim loại từ zinc oxide bằng phương pháp nhiệt luyện tạo ra",
            [
                ("A", "kẽm (zinc) ở thể hơi."),
                ("B", "kẽm (zinc) và hơi nước."),
                ("C", "kẽm (zinc) ở thể rắn."),
                ("D", "hợp kim của zinc và carbon.")
            ],
            "A",
            "2_rows",
            "Phản ứng: ZnO + C → Zn + CO (ở nhiệt độ > 1000 °C). Do kẽm có nhiệt độ sôi 907 °C nên ở nhiệt độ này kẽm thoát ra ở thể hơi, sau đó được ngưng tụ thu hồi.\n"
            "Chọn A."
        ),
        (
            14,
            "Phát biểu nào sau đây là đúng khi nói về phương pháp tách kim loại ra khỏi hợp chất của chúng?",
            [
                ("A", "Đồng có thể được tách từ copper(II) oxide bằng cách nung nóng."),
                ("B", "Kim loại Zn được tách ra khỏi ZnO bằng phương pháp nhiệt luyện."),
                ("C", "Kim loại Mg được tách ra khỏi dung dịch MgCl~2~ bằng phương pháp điện phân."),
                ("D", "Kim loại Al được tách ra khỏi Al~2~O~3~ bởi CO bằng phương pháp nhiệt luyện.")
            ],
            "B",
            "4_rows",
            "ZnO bị khử bởi C hoặc CO ở nhiệt độ cao thu được kim loại Zn (phương pháp nhiệt luyện). Các phát biểu còn lại đều sai.\n"
            "Chọn B."
        ),
        (
            15,
            "Phát biểu nào sau đây là không đúng về điều kiện xảy ra ăn mòn điện hoá?",
            [
                ("A", "Hai điện cực phải tiếp xúc trực tiếp hoặc thông qua dây dẫn."),
                ("B", "Hai điện cực cùng tiếp xúc với dung dịch chất điện li."),
                ("C", "Hai điện cực phải là hai kim loại khác loại."),
                ("D", "Thiếu một trong các điều kiện về điện cực, tiếp xúc, chất điện li sẽ không xảy ra ăn mòn.")
            ],
            "C",
            "4_rows",
            "Hai điện cực có thể là một kim loại và một phi kim dẫn điện (ví dụ: Fe và C trong gang, thép), không bắt buộc phải là hai kim loại khác loại. Nhận định C sai.\n"
            "Chọn C."
        ),
        (
            16,
            "Trong lò luyện gang có phản ứng sau: Fe~2~O~3~ + 3CO → 2Fe + 3CO~2~. Để tăng hiệu suất phản ứng, có thể sử dụng các biện pháp nào sau đây?",
            [
                ("A", "Tăng diện tích bề mặt Fe~2~O~3~."),
                ("B", "Tăng lượng khí CO, giảm lượng khí CO~2~."),
                ("C", "Giảm áp suất chung của hệ phản ứng."),
                ("D", "Tăng áp suất chung của hệ phản ứng.")
            ],
            "B",
            "2_rows",
            "Tăng nồng độ chất phản ứng (CO) và liên tục loại bỏ sản phẩm (CO~2~) làm cân bằng chuyển dịch mạnh mẽ theo chiều thuận, làm tăng hiệu suất phản ứng.\n"
            "Chọn B."
        ),
        (
            17,
            "Dural là một loại hợp kim quan trọng của nhôm, có đặc điểm là nhẹ, cứng, bền cơ học phù hợp với các ứng dụng nào sau đây?\n"
            "(1) Chế tạo cánh máy bay. (2) Áo giáp, khiên bảo vệ. (3) Làm ống dẫn dầu, mỏ neo.",
            [
                ("A", "(1), (2)."),
                ("B", "(1)."),
                ("C", "(1), (2), (3)."),
                ("D", "(1), (3).")
            ],
            "B",
            "1_row",
            "Duralumin có độ bền cao và rất nhẹ, ứng dụng chủ yếu trong chế tạo vỏ, cánh máy bay và phương tiện giao thông. Ống dẫn dầu hay mỏ neo cần vật liệu chịu ăn mòn biển sâu hoặc nặng như thép đúc.\n"
            "Chọn B."
        ),
        (
            18,
            "Các tấm tôn (thép mạ kẽm) lợp nhà thường bị gỉ sét nhanh hơn tại các vị trí cố định bằng đinh thép. Nguyên nhân của hiện tượng này là do",
            [
                ("A", "việc đóng đinh tạo điều kiện cho nước mưa tiếp xúc với các kim loại trong tấm tôn."),
                ("B", "đinh thép đóng vai trò như anode làm cho việc ăn mòn diễn ra nhanh hơn."),
                ("C", "do trong đinh thép có carbon, đóng vai trò như cathode thúc đẩy quá trình ăn mòn."),
                ("D", "do chất lượng thép của đinh và trong tấm tôn là không đồng nhất.")
            ],
            "C",
            "4_rows",
            "Khi đóng đinh thép, carbon trong đinh thép đóng vai trò là cathode (cực dương), tiếp xúc với màng nước mưa làm tăng tốc độ ăn mòn điện hoá của lớp kẽm và lõi thép tôn.\n"
            "Chọn C."
        )
    ]

    for q_num, content, opts, ans_key, layout, sol in p1_data:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        add_chem_run(p_q, f"Câu {q_num}. ", bold=True)
        add_chem_run(p_q, content)

        add_mcq_options(doc, opts, correct_key=ans_key, is_teacher=is_teacher, layout=layout)
        if is_teacher:
            add_loigiai(doc, sol)

    # ==============================================================
    # PHẦN II
    # ==============================================================
    add_section_header(
        doc,
        "PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    p2_data = [
        (
            1,
            "Kim loại magnesium có khối lượng riêng là 1,74 g/cm^3^. Kim loại này phản ứng chậm với nước ở nhiệt độ thường tạo ra anion OH^-^ và khí H~2~, phản ứng nhanh với sulfuric acid loãng và làm nhạt màu dung dịch copper(II) sulfate. Mỗi phát biểu sau đây là đúng hay sai?",
            [
                ("a", "Những hợp kim mà magnesium đóng vai trò kim loại cơ bản là những hợp kim nặng.", False),
                ("b", "Giá trị thế điện cực chuẩn của Mg^2+^/Mg lớn hơn 0 V.", False),
                ("c", "Trong môi trường trung tính, có: 2H~2~O + 2e ⇌ 2OH^-^ + H~2~ với E = -0,413 V.", True),
                ("d", "Giá trị thế điện cực chuẩn của Cu^2+^/Cu lớn hơn giá trị thế điện cực chuẩn của Mg^2+^/Mg.", True)
            ],
            None,
            "a) Sai. Magnesium có khối lượng riêng nhỏ (1,74 g/cm^3^ < 5 g/cm^3^) nên là kim loại nhẹ, hợp kim của nó là hợp kim nhẹ.\n"
            "b) Sai. E°(Mg^2+^/Mg) = -2,37 V < 0 V.\n"
            "c) Đúng. Trong môi trường trung tính (pH = 7), thế khử của nước là -0,413 V.\n"
            "d) Đúng. E°(Cu^2+^/Cu) = +0,34 V > E°(Mg^2+^/Mg) = -2,37 V."
        ),
        (
            2,
            "Cho lá kim loại kẽm (zinc) vào ống nghiệm chứa lượng dư dung dịch copper(II) sulfate màu xanh, thấy:\n"
            "- Màu xanh của dung dịch nhạt dần.\n"
            "- Có hiện tượng sủi bọt khí quanh lá kim loại.\n"
            "Mỗi phát biểu sau đây là đúng hay sai?",
            [
                ("a", "Tính oxi hoá của Zn^2+^ mạnh hơn tính oxi hoá của Cu^2+^.", False),
                ("b", "Trong ống nghiệm có diễn ra quá trình ăn mòn điện hoá.", True),
                ("c", "Muối ZnSO~4~ không tan trong nước.", False),
                ("d", "Trong ống nghiệm có diễn ra quá trình các ion H^+^ nhận electron.", False)
            ],
            None,
            "a) Sai. Tính oxi hoá của Cu^2+^ mạnh hơn Zn^2+^.\n"
            "b) Đúng. Kim loại Cu sinh ra bám vào Zn tạo cặp điện cực Zn - Cu cùng tiếp xúc với dung dịch chất điện li => xảy ra ăn mòn điện hoá học.\n"
            "c) Sai. ZnSO~4~ là muối tan tốt trong nước.\n"
            "d) Sai. Quá trình nhận electron chủ đạo là của ion Cu^2+^ (Cu^2+^ + 2e → Cu)."
        ),
        (
            3,
            "Hình bên minh hoạ quá trình tinh chế đồng bằng phương pháp điện phân.",
            [
                ("a", "Thanh A là đồng nguyên chất, thanh B là đồng chứa tạp chất.", True),
                ("b", "C là dung dịch CuSO~4~.", True),
                ("c", "D là đồng tinh khiết.", False),
                ("d", "Thanh đồng B nối vào cực âm của nguồn điện.", False)
            ],
            r"c:\Antigravity_Thanh\San_Pham\images_crop_de16\p2_cau3_tinh_che_dong.png",
            "a) Đúng. Thanh A là cathode (đồng tinh khiết, đồng bám vào dày lên), thanh B là anode (đồng thô chứa tạp chất, tan dần).\n"
            "b) Đúng. Dung dịch C là chất điện li CuSO~4~.\n"
            "c) Sai. D là bùn anode lắng dưới đáy chứa các tạp chất không tan như Ag, Au, Pt.\n"
            "d) Sai. Thanh B là anode chứa tạp chất nên phải nối vào cực dương của nguồn điện."
        ),
        (
            4,
            "Trong vỏ Trái đất, sắt và nhôm là hai nguyên tố kim loại có hàm lượng cao hơn so với các nguyên tố kim loại khác.",
            [
                ("a", "Sắt hoặc nhôm đều được sử dụng với vai trò là kim loại cơ bản trong sản xuất các hợp kim nặng.", False),
                ("b", "Từ quặng bauxite sẽ tách được sắt bằng phương pháp nhiệt luyện, từ quặng hematite sẽ tách được nhôm bằng phương pháp điện phân.", False),
                ("c", "Khi tráng một lớp kẽm lên đinh thép sẽ hạn chế được sự ăn mòn sắt trong thép theo phương pháp điện hoá.", False),
                ("d", "Nhiệt độ cần để tái chế thép cao hơn nhiệt độ cần để tái chế nhôm.", True)
            ],
            None,
            "a) Sai. Nhôm là kim loại nhẹ dùng làm hợp kim nhẹ.\n"
            "b) Sai. Quặng bauxite dùng để điều chế nhôm, quặng hematite dùng để luyện gang, thép (sắt).\n"
            "c) Sai. Tráng kẽm là phương pháp bảo vệ bề mặt.\n"
            "d) Đúng. Nhiệt độ nóng chảy của Fe (1538 °C) cao hơn nhiều so với Al (660 °C)."
        )
    ]

    for q_num, content, sub_items, img_path, sol in p2_data:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        add_chem_run(p_q, f"Câu {q_num}. ", bold=True)
        add_chem_run(p_q, content)

        if img_path and os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(2)
            p_img.add_run().add_picture(img_path, width=Inches(2.5))

        for sub_lbl, sub_txt, is_true in sub_items:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.left_indent = Inches(0.2)
            p_sub.paragraph_format.space_before = Pt(0)
            p_sub.paragraph_format.space_after = Pt(1.5)
            p_sub.paragraph_format.line_spacing = 1.15

            add_chem_run(p_sub, f"{sub_lbl}) ", bold=True)
            add_chem_run(p_sub, sub_txt)

            if is_teacher:
                status_txt = " [ĐÚNG]" if is_true else " [SAI]"
                add_chem_run(p_sub, status_txt, bold=True, color=COLOR_CORAL_RED)

        if is_teacher:
            add_loigiai(doc, sol)

    # ==============================================================
    # PHẦN III
    # ==============================================================
    add_section_header(
        doc,
        "PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 6."
    )

    p3_data = [
        (
            1,
            "Cho quá trình: 4Al(s) + 3O~2~(g) → 2Al~2~O~3~(s), Δ_r H°_298 = -3351,4 kJ. Lượng nhiệt toả ra khi 1 gam nhôm (aluminium) bị oxi hoá hoàn toàn ở điều kiện chuẩn là bao nhiêu kJ? (Làm tròn kết quả đến hàng phần mười).",
            "31,1",
            "Theo phương trình nhiệt hoá học, 4 mol Al (tương ứng 4 * 27 = 108 g) toả ra 3351,4 kJ.\n"
            "Nhiệt lượng toả ra khi oxi hoá hoàn toàn 1 gam nhôm là:\n"
            "Q = 3351,4 / 108 ≈ 31,03 ≈ 31,1 kJ.\n"
            "Đáp số: 31,1."
        ),
        (
            2,
            "Cho các kim loại: Ag, Al, Au, Cu, Cr, Fe, Mg, Pt, Zn. Có bao nhiêu kim loại phản ứng được với dung dịch sulfuric acid đặc, nguội?",
            "4",
            "- Au, Pt là kim loại trơ không phản ứng với H~2~SO~4~ đặc, nguội.\n"
            "- Al, Fe, Cr bị thụ động hoá trong H~2~SO~4~ đặc, nguội.\n"
            "- Các kim loại phản ứng được gồm: Ag, Cu, Mg, Zn. Tổng cộng có 4 kim loại.\n"
            "Đáp số: 4."
        ),
        (
            3,
            "Cho các cặp oxi hoá - khử: Ag^+^/Ag, Al^3+^/Al, Au^3+^/Au, Cu^2+^/Cu, Fe^2+^/Fe, K^+^/K, Mg^2+^/Mg, Na^+^/Na, Pt^2+^/Pt, Zn^2+^/Zn và Pb^2+^/Pb. Có bao nhiêu cặp oxi hoá - khử có giá trị thế điện cực chuẩn lớn hơn 0 V?",
            "4",
            "Các cặp oxi hoá - khử có E° > 0 V đứng sau H^+^/H~2~ trong dãy điện hoá gồm:\n"
            "Cu^2+^/Cu (+0,34 V), Ag^+^/Ag (+0,80 V), Pt^2+^/Pt (+1,20 V), Au^3+^/Au (+1,52 V).\n"
            "Tổng cộng có 4 cặp.\n"
            "Đáp số: 4."
        ),
        (
            4,
            "Có bao nhiêu trường hợp xảy ra ăn mòn điện hoá trong các quá trình sau?\n"
            "(a) Vỏ tàu bằng thép có gắn các khối kẽm, neo đậu tại một cảng biển.\n"
            "(b) Vật dụng bằng bạc bị sẫm màu khi tiếp xúc với không khí có lẫn H~2~S.\n"
            "(c) Tấm tôn trầy xước trên mái nhà tiếp xúc với nước mưa.\n"
            "(d) Hợp kim Na - K nóng đỏ, bốc cháy khi tiếp xúc với khí chlorine.\n"
            "(e) Một cổ vật làm bằng đồng thau (hợp kim Cu - Zn) chìm trong nước biển.\n"
            "(g) Dây chảy cầu chì làm bằng hợp kim Pb - Sn đứt khi cường độ dòng điện tăng đột ngột.",
            "3",
            "Các trường hợp ăn mòn điện hoá gồm:\n"
            "- (a) Cặp Zn - Fe trong chất điện li nước biển.\n"
            "- (c) Cặp Zn - Fe bị trầy xước tiếp xúc với màng nước mưa.\n"
            "- (e) Cặp Cu - Zn trong hợp kim đồng thau tiếp xúc với nước biển.\n"
            "Các trường hợp (b), (d) là ăn mòn hoá học; (g) là hiện tượng vật lý.\n"
            "Tổng cộng có 3 trường hợp.\n"
            "Đáp số: 3."
        ),
        (
            5,
            "Trong quá trình sản xuất nhôm (aluminium) bằng phương pháp điện phân nóng chảy Al~2~O~3~, điện cực dương bằng than chì bị ăn mòn liên tục do phản ứng tạo thành hỗn hợp khí O~2~, CO, CO~2~ có tỉ lệ mol bằng nhau. Giả sử các phản ứng xảy ra hoàn toàn, mỗi kg nhôm sinh ra tương ứng với bao nhiêu kg than chì bị đốt cháy ở cực dương? (Làm tròn đến hàng phần trăm).",
            "0,27",
            "Gọi số mol các khí ở anode: n(O~2~) = x, n(CO) = x, n(CO~2~) = x (kmol).\n"
            "Tổng mol O = 2x + x + 2x = 5x kmol => n(Al) = (10x / 3) kmol.\n"
            "Khối lượng Al: m(Al) = (10x / 3) * 27 = 90x = 1 kg => x = 1/90 kmol.\n"
            "Khối lượng carbon bị đốt cháy:\n"
            "m(C) = 12 * (n_CO + n_CO2) = 12 * 2x = 24 * (1/90) ≈ 0,267 ≈ 0,27 kg.\n"
            "Đáp số: 0,27."
        ),
        (
            6,
            "Một dòng điện 1,04 A đi qua dung dịch sulfuric acid loãng trong 6,00 phút trong một thiết bị điện phân nước. Tổng số mol H~2~ và O~2~ thu được là n * 10^-3^ mol. Giá trị của n là bao nhiêu? (Làm tròn kết quả đến hàng phần trăm).",
            "2,91",
            "Thời gian: t = 6,00 * 60 = 360 giây.\n"
            "Số mol electron: n_e = (I * t) / F = (1,04 * 360) / 96500 ≈ 0,0038798 mol.\n"
            "Quá trình ở các điện cực:\n"
            "Cathode: 2H^+^ + 2e → H~2~ => n(H~2~) = n_e / 2.\n"
            "Anode: 2H~2~O → O~2~ + 4H^+^ + 4e => n(O~2~) = n_e / 4.\n"
            "Tổng số mol khí: n_khí = (3/4) * n_e = (3/4) * 0,0038798 ≈ 0,00291 mol = 2,91 * 10^-3^ mol.\n"
            "Giá trị của n là 2,91.\n"
            "Đáp số: 2,91."
        )
    ]

    for q_num, content, ans, sol in p3_data:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        add_chem_run(p_q, f"Câu {q_num}. ", bold=True)
        add_chem_run(p_q, content)

        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(3)
        p_ans.paragraph_format.line_spacing = 1.15
        if is_teacher:
            add_chem_run(p_ans, "Đáp số: ", bold=True)
            add_chem_run(p_ans, ans, bold=True, color=COLOR_CORAL_RED)
            add_loigiai(doc, sol)
        else:
            add_chem_run(p_ans, "Trả lời: ...........................................................................................................................................................", italic=True)

    # Summary table for teacher version
    if is_teacher:
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(12)
        p_title.paragraph_format.space_after = Pt(4)
        r_t = p_title.add_run("BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ SỐ 16")
        r_t.font.name = "Times New Roman"
        r_t.font.size = Pt(13)
        r_t.font.bold = True
        r_t.font.color.rgb = COLOR_DARK_GREEN

        p_p1 = doc.add_paragraph()
        p_p1.paragraph_format.space_before = Pt(4)
        p_p1.paragraph_format.space_after = Pt(2)
        r_p1 = p_p1.add_run("PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN")
        r_p1.font.name = "Times New Roman"
        r_p1.font.size = Pt(11)
        r_p1.font.bold = True
        r_p1.font.color.rgb = COLOR_DARK_GREEN

        t1 = doc.add_table(rows=2, cols=18)
        t1.alignment = WD_TABLE_ALIGNMENT.CENTER
        for j in range(18):
            c_h = t1.cell(0, j)
            c_h.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_h = c_h.paragraphs[0].add_run(str(j + 1))
            r_h.font.name = "Times New Roman"
            r_h.font.size = Pt(9.5)
            r_h.font.bold = True

            c_v = t1.cell(1, j)
            c_v.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_v = c_v.paragraphs[0].add_run(p1_data[j][3])
            r_v.font.name = "Times New Roman"
            r_v.font.size = Pt(9.5)
            r_v.font.bold = True
            r_v.font.color.rgb = COLOR_CORAL_RED

        p_p2 = doc.add_paragraph()
        p_p2.paragraph_format.space_before = Pt(6)
        p_p2.paragraph_format.space_after = Pt(2)
        r_p2 = p_p2.add_run("PHẦN II: TRẮC NGHIỆM ĐÚNG SAI")
        r_p2.font.name = "Times New Roman"
        r_p2.font.size = Pt(11)
        r_p2.font.bold = True
        r_p2.font.color.rgb = COLOR_DARK_GREEN

        t2 = doc.add_table(rows=5, cols=5)
        t2.alignment = WD_TABLE_ALIGNMENT.CENTER
        headers_p2 = ["Câu", "Lệnh hỏi a", "Lệnh hỏi b", "Lệnh hỏi c", "Lệnh hỏi d"]
        for j, h in enumerate(headers_p2):
            c = t2.cell(0, j)
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = c.paragraphs[0].add_run(h)
            r.font.name = "Times New Roman"
            r.font.size = Pt(10)
            r.font.bold = True

        for i in range(4):
            c_num = t2.cell(i + 1, 0)
            c_num.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = c_num.paragraphs[0].add_run(str(i + 1))
            r.font.name = "Times New Roman"
            r.font.size = Pt(10)
            r.font.bold = True
            for j in range(4):
                c_val = t2.cell(i + 1, j + 1)
                c_val.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                is_t = p2_data[i][2][j][2]
                val_str = "Đ" if is_t else "S"
                r_v = c_val.paragraphs[0].add_run(val_str)
                r_v.font.name = "Times New Roman"
                r_v.font.size = Pt(10)
                r_v.font.bold = True
                if is_t:
                    r_v.font.color.rgb = COLOR_CORAL_RED

        p_p3 = doc.add_paragraph()
        p_p3.paragraph_format.space_before = Pt(6)
        p_p3.paragraph_format.space_after = Pt(2)
        r_p3 = p_p3.add_run("PHẦN III: TRẢ LỜI NGẮN")
        r_p3.font.name = "Times New Roman"
        r_p3.font.size = Pt(11)
        r_p3.font.bold = True
        r_p3.font.color.rgb = COLOR_DARK_GREEN

        t3 = doc.add_table(rows=2, cols=6)
        t3.alignment = WD_TABLE_ALIGNMENT.CENTER
        for j in range(6):
            c_h = t3.cell(0, j)
            c_h.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = c_h.paragraphs[0].add_run(f"Câu {j + 1}")
            r.font.name = "Times New Roman"
            r.font.size = Pt(10)
            r.font.bold = True

            c_v = t3.cell(1, j)
            c_v.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_v = c_v.paragraphs[0].add_run(p3_data[j][2])
            r_v.font.name = "Times New Roman"
            r_v.font.size = Pt(10)
            r_v.font.bold = True
            r_v.font.color.rgb = COLOR_CORAL_RED

    file_name = "De_Kiem_Tra_Hoa_12_De16_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_12_De16.docx"
    out_path = os.path.join(OUT_DIR, file_name)
    doc.save(out_path)
    print(f"Saved: {out_path}")

if __name__ == "__main__":
    generate_exam_docx(is_teacher=False)
    generate_exam_docx(is_teacher=True)
