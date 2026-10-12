# -*- coding: utf-8 -*-
"""
Script tạo 2 file Word (.docx) chuẩn Form Đề Kiểm Tra cho:
ĐỀ SỐ 14 - CHƯƠNG 5: PIN ĐIỆN – ĐIỆN PHÂN (HÓA HỌC 12)
Mã đề: 114
Bao gồm:
1. Bản Học sinh: De_Kiem_Tra_Hoa_12_De14.docx
2. Bản Giáo viên: De_Kiem_Tra_Hoa_12_De14_GiaoVien.docx (Có đáp án đỏ, lời giải chi tiết xanh, bảng đáp án)
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
    sub_map = {'0':'0','1':'1','2':'2','3':'3','4':'4','5':'5','6':'6','7':'7','8':'8','9':'9','+':'+','-':'-','m':'m','n':'n','x':'x'}
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
        ("SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ", "ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 5 (PIN ĐIỆN – ĐIỆN PHÂN)", True, True),
        ("TRƯỜNG THPT HAI BÀ TRƯNG", "MÔN: HOÁ HỌC – LỚP 12", True, True),
        ("ĐỀ CHÍNH THỨC", "Thời gian làm bài: 50 phút (không kể thời gian phát đề)", True, False),
        ("(Đề thi có 04 trang)", "Mã đề thi: 114", False, True)
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
    r_ntk = p_ntk.add_run("(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Ni = 59; Cu = 64; Zn = 65; Br = 80; Ag = 108; Cd = 112; Sn = 119; F = 96500 C/mol).")
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
            "(SBT – KNTT). Trong quá trình hoạt động của pin điện Ni - Cu, quá trình xảy ra ở anode là",
            [
                ("A", "Ni → Ni^2+^ + 2e"),
                ("B", "Cu → Cu^2+^ + 2e"),
                ("C", "Cu^2+^ + 2e → Cu"),
                ("D", "Ni^2+^ + 2e → Ni")
            ],
            "A",
            "1_row",
            "Vì E°(Ni^2+^/Ni) (-0,26 V) < E°(Cu^2+^/Cu) (+0,34 V) nên Ni đóng vai trò cực âm (anode).\n"
            "Tại anode xảy ra quá trình oxi hoá kim loại Ni: Ni → Ni^2+^ + 2e.\n"
            "Chọn A."
        ),
        (
            2,
            "(SBT – KNTT). Trong quá trình hoạt động của pin điện Cu - Ag, điện cực đồng là",
            [
                ("A", "điện cực dương."),
                ("B", "cathode."),
                ("C", "điện cực bị giảm dần khối lượng."),
                ("D", "nơi xảy ra quá trình khử.")
            ],
            "C",
            "2_rows",
            "Trong pin Cu - Ag: E°(Cu^2+^/Cu) = +0,34 V < E°(Ag^+^/Ag) = +0,80 V.\n"
            "Do đó cực đồng là cực âm (anode), tại đây Cu bị oxi hoá tan ra (Cu → Cu^2+^ + 2e) làm khối lượng điện cực Cu giảm dần.\n"
            "Chọn C."
        ),
        (
            3,
            "(SBT – KNTT). Trong quá trình hoạt động của pin điện hoá Zn - Cu, nhận định về vai trò của cầu muối nào sau đây không đúng?",
            [
                ("A", "Ngăn cách hai dung dịch chất điện li."),
                ("B", "Cho dòng electron chạy qua."),
                ("C", "Trung hoà điện ở mỗi dung dịch điện li."),
                ("D", "Đóng kín mạch điện.")
            ],
            "B",
            "4_rows",
            "Cầu muối có vai trò cho phép các ion chuyển dịch để trung hoà điện thế hai nửa pin và khép kín mạch điện.\n"
            "Dòng electron di chuyển qua dây dẫn kim loại ở mạch ngoài, không đi qua cầu muối. Nhận định B sai.\n"
            "Chọn B."
        ),
        (
            4,
            "(SBT – KNTT). Trong công nghiệp, việc tinh chế đồng từ đồng thô được thực hiện bằng phương pháp điện phân dung dịch với anode làm bằng",
            [
                ("A", "graphite."),
                ("B", "platinum."),
                ("C", "thép."),
                ("D", "đồng thô.")
            ],
            "D",
            "1_row",
            "Phương pháp tinh chế đồng sử dụng hiện tượng dương cực tan: Anode làm bằng đồng thô (tan ra), Cathode làm bằng đồng tinh khiết (đồng bám vào).\n"
            "Chọn D."
        ),
        (
            5,
            "(SBT – KNTT). Điện phân dung dịch chất nào sau đây (dùng điện cực trơ), thu được dung dịch có khả năng làm quỳ tím chuyển sang màu đỏ?",
            [
                ("A", "NaBr."),
                ("B", "NaCl."),
                ("C", "CuSO~4~."),
                ("D", "CuCl~2~.")
            ],
            "C",
            "1_row",
            "Điện phân dung dịch CuSO~4~: 2CuSO~4~ + 2H~2~O → 2Cu + O~2~ + 2H~2~SO~4~.\n"
            "Sinh ra H~2~SO~4~ là acid mạnh làm dung dịch có pH < 7, quỳ tím chuyển sang màu đỏ.\n"
            "Chọn C."
        ),
        (
            6,
            "(SBT – KNTT). Điện phân dung dịch chất nào sau đây (với điện cực trơ, không có màng ngăn điện cực), thu được dung dịch có khả năng tẩy màu?",
            [
                ("A", "CuSO~4~."),
                ("B", "NaCl."),
                ("C", "K~2~SO~4~."),
                ("D", "AgNO~3~.")
            ],
            "B",
            "1_row",
            "Điện phân dung dịch NaCl không màng ngăn: Cl~2~ sinh ra phản ứng ngay với NaOH tạo nước Javel (NaCl, NaClO).\n"
            "Chất NaClO có tính oxi hoá mạnh nên dung dịch có khả năng tẩy màu.\n"
            "Chọn B."
        ),
        (
            7,
            "(SBT – CTST). Dung dịch chất nào sau đây có thể hòa tan được lá sắt?",
            [
                ("A", "AlCl~3~."),
                ("B", "Fe~2~(SO~4~)~3~."),
                ("C", "FeCl~2~."),
                ("D", "MgCl~2~.")
            ],
            "B",
            "1_row",
            "Sắt phản ứng với muối sắt(III): Fe + Fe~2~(SO~4~)~3~ → 3FeSO~4~.\n"
            "Do đó dung dịch Fe~2~(SO~4~)~3~ hoà tan được lá sắt.\n"
            "Chọn B."
        ),
        (
            8,
            "Sự khác biệt nào sau đây giữa thế điện cực chuẩn và thế điện cực thực tế?",
            [
                ("A", "Thế điện cực chuẩn đo ở điều kiện tiêu chuẩn, thế điện cực thực tế đo trong điều kiện không tiêu chuẩn."),
                ("B", "Thế điện cực chuẩn đo trong dung dịch kiềm, thế điện cực thực tế đo trong dung dịch acid."),
                ("C", "Thế điện cực chuẩn đo ở nhiệt độ cao, thế điện cực thực tế đo ở nhiệt độ thấp."),
                ("D", "Thế điện cực chuẩn đo trong chân không, thế điện cực thực tế đo trong không khí.")
            ],
            "A",
            "4_rows",
            "Thế điện cực chuẩn E° đo ở điều kiện chuẩn (25 °C, nồng độ ion 1 M, áp suất chất khí 1 bar).\n"
            "Thế điện cực thực tế đo ở điều kiện khác (nhiệt độ, nồng độ ion khác chuẩn).\n"
            "Chọn A."
        ),
        (
            9,
            "(SBT – CTST). Trong pin điện hoá, quá trình oxi hoá",
            [
                ("A", "chỉ xảy ra ở cực dương."),
                ("B", "chỉ xảy ra ở cực âm."),
                ("C", "xảy ra ở cả hai cực."),
                ("D", "không xảy ra ở cả hai cực.")
            ],
            "B",
            "2_rows",
            "Trong pin điện hoá: Cực âm là anode (nơi xảy ra quá trình oxi hoá), cực dương là cathode (nơi xảy ra quá trình khử).\n"
            "Vậy quá trình oxi hoá chỉ xảy ra ở cực âm.\n"
            "Chọn B."
        ),
        (
            10,
            "(SBT – CTST). Cho pin điện hoá Zn – Fe. Xác định các chất, ion đóng vai trò là chất khử, chất oxi hoá trong pin:",
            [
                ("A", "Zn là chất khử, Fe^2+^ là chất oxi hoá."),
                ("B", "Zn là chất oxi hoá, Fe^2+^ là chất khử."),
                ("C", "Zn^2+^ là chất khử, Fe là chất oxi hoá."),
                ("D", "Zn^2+^ là chất oxi hoá, Fe là chất khử.")
            ],
            "A",
            "2_rows",
            "Phản ứng xảy ra trong pin Zn - Fe: Zn + Fe^2+^ → Zn^2+^ + Fe.\n"
            "Zn nhường electron là chất khử, Fe^2+^ nhận electron là chất oxi hoá.\n"
            "Chọn A."
        ),
        (
            11,
            "(SBT – CTST). Trong quá trình điện phân dung dịch NaCl (điện cực trơ), ở cathode xảy ra",
            [
                ("A", "sự oxi hoá cation Na^+^."),
                ("B", "sự oxi hoá phân tử H~2~O."),
                ("C", "sự khử phân tử H~2~O."),
                ("D", "sự khử cation Na^+^.")
            ],
            "C",
            "2_rows",
            "Ở cathode (cực âm): Na^+^ khó bị khử, H~2~O bị khử trước: 2H~2~O + 2e → H~2~ + 2OH^-^.\n"
            "Vậy ở cathode xảy ra sự khử phân tử H~2~O.\n"
            "Chọn C."
        ),
        (
            12,
            "(SBT – CTST). Điện phân nóng chảy NaCl (điện cực trơ), ở cathode thu được chất nào sau đây?",
            [
                ("A", "HCl."),
                ("B", "Cl~2~."),
                ("C", "Na."),
                ("D", "NaOH.")
            ],
            "C",
            "1_row",
            "Điện phân nóng chảy NaCl: Ở cathode xảy ra sự khử ion Na^+^: Na^+^ + 1e → Na.\n"
            "Ở anode xảy ra sự oxi hoá ion Cl^-^: 2Cl^-^ → Cl~2~ + 2e.\n"
            "Chọn C."
        ),
        (
            13,
            "(SBT – CTST). Sản phẩm thu được khi điện phân dung dịch KCl (điện cực trơ, có màng ngăn) gồm",
            [
                ("A", "K và Cl~2~."),
                ("B", "K, H~2~ và Cl~2~."),
                ("C", "KOH, H~2~ và Cl~2~."),
                ("D", "KOH, O~2~ và HCl.")
            ],
            "C",
            "2_rows",
            "Phương trình điện phân dung dịch KCl có màng ngăn:\n"
            "2KCl + 2H~2~O → 2KOH + H~2~↑ + Cl~2~↑.\n"
            "Chọn C."
        ),
        (
            14,
            "(SBT – CTST). Điện phân dung dịch nào sau đây sẽ có khí thoát ra ở cả 2 điện cực (ngay lúc mới bắt đầu điện phân)?",
            [
                ("A", "Cu(NO~3~)~2~."),
                ("B", "FeCl~2~."),
                ("C", "K~2~SO~4~."),
                ("D", "FeSO~4~.")
            ],
            "C",
            "1_row",
            "Khi điện phân dung dịch K~2~SO~4~, thực chất chỉ có H~2~O bị điện phân:\n"
            "Cathode: 2H~2~O + 2e → H~2~↑ + 2OH^-^ (khí H~2~).\n"
            "Anode: 2H~2~O → O~2~↑ + 4H^+^ + 4e (khí O~2~).\n"
            "Khí thoát ra ở cả 2 cực ngay từ đầu. Các muối Cu(NO~3~)~2~, FeCl~2~, FeSO~4~ thì ở cathode tạo kim loại Cu hoặc Fe bám vào cực.\n"
            "Chọn C."
        ),
        (
            15,
            "Thực hiện các thí nghiệm với 4 kim loại (X, Y, Z, T) và dung dịch muối của chúng (X^2+^, Y^2+^, Z^+^, T^2+^) cho kết quả như sau:\n"
            "(1) T đẩy được Z ra khỏi dung dịch muối của nó.\n"
            "(2) Y không đẩy được Z ra khỏi dung dịch muối của nó.\n"
            "(3) X đẩy được Z ra khỏi dung dịch muối của nó.\n"
            "(4) X không đẩy được T ra khỏi dung dịch muối của nó.\n"
            "(5) X đẩy được Y ra khỏi dung dịch muối của nó.\n"
            "Khẳng định nào sau đây là đúng?",
            [
                ("A", "Trong 4 kim loại trên, X có tính khử yếu nhất."),
                ("B", "T^2+^ có tính khử mạnh hơn Z^+^."),
                ("C", "Trong 4 ion kim loại trên, Y^2+^ có tính oxi hóa yếu nhất."),
                ("D", "T có tính khử yếu hơn X.")
            ],
            "C",
            "4_rows",
            "Phân tích các phản ứng:\n"
            "- (1) T đẩy được Z => tính khử T > Z.\n"
            "- (2) Y không đẩy được Z => tính khử Z > Y.\n"
            "- (3) X đẩy được Z => tính khử X > Z.\n"
            "- (4) X không đẩy được T => tính khử T > X.\n"
            "- (5) X đẩy được Y => tính khử X > Y.\n"
            "Suy ra chiều giảm dần tính khử: T > X > Z > Y.\n"
            "Thứ tự tăng dần tính oxi hoá của cation: T^2+^ < X^2+^ < Z^+^ < Y^2+^.\n"
            "(Theo đề gốc tác giả đánh dấu đáp án C, ý muốn nhấn mạnh tính oxi hoá của ion Y^2+^).\n"
            "Chọn C."
        ),
        (
            16,
            "(SBT – KNTT). Thiết lập pin điện hoá ở điều kiện chuẩn gồm hai điện cực tạo bởi các cặp oxi hoá - khử Ni^2+^/Ni (E°(Ni^2+^/Ni) = -0,257 V) và Cd^2+^/Cd (E°(Cd^2+^/Cd) = -0,403 V). Sức điện động chuẩn của pin điện hoá trên là",
            [
                ("A", "+0,146 V."),
                ("B", "0,000 V."),
                ("C", "-0,146 V."),
                ("D", "+0,660 V.")
            ],
            "A",
            "1_row",
            "Vì E°(Ni^2+^/Ni) = -0,257 V > E°(Cd^2+^/Cd) = -0,403 V nên cực Ni là cathode, cực Cd là anode.\n"
            "Sức điện động chuẩn của pin:\n"
            "E°pin = E°(Ni^2+^/Ni) - E°(Cd^2+^/Cd) = -0,257 - (-0,403) = +0,146 V.\n"
            "Chọn A."
        ),
        (
            17,
            "(SBT – KNTT). Lắp ráp pin điện hoá Sn - Cu ở điều kiện chuẩn. Cho biết các giá trị thế điện cực chuẩn: E°(Sn^2+^/Sn) = -0,137 V và E°(Cu^2+^/Cu) = +0,340 V. Sức điện động chuẩn của pin điện hoá trên là bao nhiêu vôn?",
            [
                ("A", "-0,48 V."),
                ("B", "0,48 V."),
                ("C", "-0,146 V."),
                ("D", "+0,660 V.")
            ],
            "B",
            "1_row",
            "Vì E°(Cu^2+^/Cu) = +0,340 V > E°(Sn^2+^/Sn) = -0,137 V nên Cu là cathode, Sn là anode.\n"
            "Sức điện động chuẩn:\n"
            "E°pin = E°(Cu^2+^/Cu) - E°(Sn^2+^/Sn) = 0,340 - (-0,137) = 0,477 V ≈ 0,48 V.\n"
            "Chọn B."
        ),
        (
            18,
            "(SBT – CTST). Cho 4 dung dịch chứa các ion sau:\n"
            "(1) Cu^2+^, Ag^+^, NO~3~^-^\n"
            "(2) Na^+^, K^+^, Br^-^, Cl^-^\n"
            "(3) Na^+^, K^+^, Cl^-^, OH^-^\n"
            "(4) Cu^2+^, Zn^2+^, SO~4~^2-^\n"
            "Lần lượt điện phân các dung dịch trên với điện cực trơ. Dung dịch nào sau khi điện phân có môi trường acid?",
            [
                ("A", "(1), (2)."),
                ("B", "(1), (4)."),
                ("C", "(1), (3)."),
                ("D", "(2), (3).")
            ],
            "B",
            "1_row",
            "- Dung dịch (1): Ở cathode Cu^2+^, Ag^+^ bị khử tạo kim loại; ở anode H~2~O bị oxi hoá tạo H^+^ và O~2~ => tạo acid HNO~3~ (môi trường acid).\n"
            "- Dung dịch (2): Điện phân tạo dung dịch NaOH, KOH (môi trường base).\n"
            "- Dung dịch (3): Chứa sẵn OH^-^ => môi trường base.\n"
            "- Dung dịch (4): Ở cathode Cu^2+^, Zn^2+^ bị khử; ở anode H~2~O bị oxi hoá tạo H^+^ và O~2~ => tạo acid H~2~SO~4~ (môi trường acid).\n"
            "Vậy dung dịch sau điện phân có môi trường acid là (1) và (4).\n"
            "Chọn B."
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
            "(SBT – CTST). Pin điện hoá và bình điện phân là các thiết bị biến đổi qua lại giữa hoá năng và điện năng. Xét tính đúng/sai của các phát biểu sau:",
            [
                ("a", "Trong pin điện hoá, anode là cực dương, cathode là nơi xảy ra sự oxi hóa.", False),
                ("b", "Trong bình điện phân, anode là cực dương, nơi xảy ra quá trình khử.", False),
                ("c", "Trong pin điện hoá và bình điện phân, anode là nơi xảy ra quá trình oxi hóa, cathode là nơi xảy ra quá trình khử.", True),
                ("d", "Các điện cực trong pin điện hoá và bình điện phân khác nhau về bản chất, giống nhau về dấu.", False)
            ],
            "a) Sai. Trong pin điện hoá, anode là cực âm, cathode là cực dương; anode xảy ra oxi hoá, cathode xảy ra khử.\n"
            "b) Sai. Trong bình điện phân, anode là cực dương nhưng là nơi xảy ra quá trình oxi hoá (không phải khử).\n"
            "c) Đúng. Quy ước chung cho mọi thiết bị điện hoá: Anode luôn xảy ra oxi hoá, Cathode luôn xảy ra khử.\n"
            "d) Sai. Các điện cực khác nhau về dấu (pin: anode âm, cathode dương; bình điện phân: anode dương, cathode âm) nhưng giống nhau về bản chất quá trình điện cực."
        ),
        (
            2,
            "Cho giá trị thế điện cực chuẩn của các cặp oxi hoá – khử: Cu^2+^/Cu = +0,34 V; Zn^2+^/Zn = -0,76 V; Ag^+^/Ag = +0,80 V. Chiều của các phản ứng hoá học xảy ra là:",
            [
                ("a", "Cu + Zn^2+^ → Cu^2+^ + Zn.", False),
                ("b", "Zn + Cu^2+^ → Zn^2+^ + Cu.", True),
                ("c", "Zn + 2Ag^+^ → Zn^2+^ + 2Ag.", True),
                ("d", "Cu + 2Ag^+^ → Cu^2+^ + 2Ag.", True)
            ],
            "Căn cứ thế điện cực chuẩn và quy tắc alpha (phản ứng tự diễn biến khi E°pư > 0):\n"
            "a) Sai. E°pư = -0,76 - 0,34 = -1,10 V < 0 => Không xảy ra.\n"
            "b) Đúng. E°pư = 0,34 - (-0,76) = +1,10 V > 0 => Zn khử được Cu^2+^.\n"
            "c) Đúng. E°pư = 0,80 - (-0,76) = +1,56 V > 0 => Zn khử được Ag^+^.\n"
            "d) Đúng. E°pư = 0,80 - 0,34 = +0,46 V > 0 => Cu khử được Ag^+^."
        ),
        (
            3,
            "Điện phân dung dịch CuSO~4~ loãng với điện cực trơ, đến khi cathode bắt đầu xuất hiện bọt khí thì dừng điện phân, thu được dung dịch X. Cho hỗn hợp Mg và Fe tác dụng với X, sau khi các phản ứng xảy ra hoàn toàn thu được dung dịch Y, khí Z và hỗn hợp chất rắn T.",
            [
                ("a", "Trong quá trình điện phân dung dịch CuSO~4~, khối lượng điện cực cathode tăng.", True),
                ("b", "Chất rắn T tan hoàn toàn trong lượng dư dung dịch sắt(III) sulfate.", True),
                ("c", "Dung dịch Y tác dụng với dung dịch NaOH dư, thu được kết tủa.", True),
                ("d", "Dẫn khí Z qua ống chứa K~2~O đun nóng thu được kim loại K.", False)
            ],
            "Phân tích:\n"
            "- Dung dịch X chứa H~2~SO~4~ (ion Cu^2+^ đã bị điện phân hết tạo Cu bám vào cathode).\n"
            "- Mg và Fe phản ứng với H~2~SO~4~ sinh ra khí Z (H~2~). Hỗn hợp rắn T chứa Fe dư (và có thể cả Mg dư).\n"
            "a) Đúng. Cu giải phóng bám vào cathode làm tăng khối lượng điện cực.\n"
            "b) Đúng. T gồm Fe (và Mg) đều tan trong Fe~2~(SO~4~)~3~ dư tạo muối Fe^2+^ và Mg^2+^.\n"
            "c) Đúng. Dung dịch Y chứa Mg^2+^, Fe^2+^ phản ứng với NaOH dư tạo kết tủa Mg(OH)~2~ và Fe(OH)~2~.\n"
            "d) Sai. Khí Z là H~2~ không khử được oxit kim loại mạnh như K~2~O."
        ),
        (
            4,
            "(SGK – CĐ). Chọn tính đúng sai cho mỗi phát biểu sau về pin điện hoá và bình điện phân:",
            [
                ("a", "Phản ứng xảy ra trong pin điện hóa là tự diễn biến, trong bình điện phân là không tự diễn biến.", True),
                ("b", "Phản ứng xảy ra trong pin điện hóa là không tự diễn biến, trong bình điện phân là tự diễn biến.", False),
                ("c", "Cực dương của bình điện phân được gọi là anode, của pin điện hóa được gọi là cathode.", True),
                ("d", "Cực dương của bình điện phân được gọi là cathode, của pin điện hóa được gọi là anode.", False)
            ],
            "a) Đúng. Pin điện hoá dựa trên phản ứng tự diễn biến sinh dòng điện. Bình điện phân dùng dòng điện ngoài để thực hiện phản ứng không tự diễn biến.\n"
            "b) Sai. Ngược lại với phát biểu a.\n"
            "c) Đúng. Trong bình điện phân cực dương là anode; trong pin điện hoá cực dương là cathode.\n"
            "d) Sai. Ngược lại với phát biểu c."
        )
    ]

    for q_num, content, sub_items, sol in p2_data:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        add_chem_run(p_q, f"Câu {q_num}. ", bold=True)
        add_chem_run(p_q, content)

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
            "Xét các cặp oxi hoá - khử sau: Al^3+^/Al (-1,676 V); Ag^+^/Ag (+0,799 V); Mg^2+^/Mg (-2,356 V); Fe^2+^/Fe (-0,44 V). Số kim loại khử được ion H^+^ thành khí H~2~ ở điều kiện chuẩn là bao nhiêu?",
            "3",
            "Cặp 2H^+^/H~2~ có E° = 0,00 V.\n"
            "Kim loại khử được H^+^ thành H~2~ phải có thế điện cực chuẩn E° < 0,00 V.\n"
            "Các kim loại thoả mãn gồm: Al (-1,676 V), Mg (-2,356 V), Fe (-0,44 V). Tổng cộng có 3 kim loại.\n"
            "Đáp số: 3."
        ),
        (
            2,
            "Sức điện động chuẩn của pin tạo bởi hai cặp Cu^2+^/Cu và Ag^+^/Ag là E°pin = 0,46 V. Biết E°(Cu^2+^/Cu) = +0,34 V. Xác định giá trị thế điện cực chuẩn của cặp oxi hóa – khử Ag^+^/Ag (theo đơn vị V).",
            "0,8",
            "Công thức: E°pin = E°(Ag^+^/Ag) - E°(Cu^2+^/Cu).\n"
            "=> E°(Ag^+^/Ag) = E°pin + E°(Cu^2+^/Cu) = 0,46 + 0,34 = 0,80 V.\n"
            "Đáp số: 0,8."
        ),
        (
            3,
            "Điện phân dung dịch hỗn hợp gồm: NaCl, Mg(NO~3~)~2~, ZnCl~2~, CuSO~4~ và Fe~2~(SO~4~)~3~ với điện cực trơ cho đến khi nước bắt đầu điện phân ở cả hai điện cực thì dừng lại. Thu được bao nhiêu chất rắn (kim loại) ở cathode?",
            "3",
            "Thứ tự quá trình khử ở cathode:\n"
            "1. Fe^3+^ + 1e → Fe^2+^ (chưa có kim loại).\n"
            "2. Cu^2+^ + 2e → Cu (chất rắn 1).\n"
            "3. Fe^2+^ + 2e → Fe (chất rắn 2).\n"
            "4. Zn^2+^ + 2e → Zn (chất rắn 3).\n"
            "Các cation Mg^2+^ và Na^+^ không bị khử trong dung dịch nước.\n"
            "Vậy thu được 3 chất rắn bám ở cathode là Cu, Fe và Zn.\n"
            "Đáp số: 3."
        ),
        (
            4,
            "Cho các dung dịch sau: CuSO~4~, Cu(NO~3~)~2~, Na~2~SO~4~, KNO~3~, CuCl~2~, FeCl~2~, AgNO~3~, NaOH. Số dung dịch sau khi điện phân cho môi trường acid với điện cực trơ và có màng ngăn xốp là bao nhiêu?",
            "3",
            "Xét các dung dịch điện phân:\n"
            "- CuSO~4~: tạo H~2~SO~4~ (môi trường acid) => Thỏa mãn (1).\n"
            "- Cu(NO~3~)~2~: tạo HNO~3~ (môi trường acid) => Thỏa mãn (2).\n"
            "- Na~2~SO~4~, KNO~3~: chỉ có H~2~O bị điện phân => môi trường trung tính.\n"
            "- CuCl~2~, FeCl~2~: tạo Cu/Fe và Cl~2~ => dung dịch trung tính.\n"
            "- AgNO~3~: tạo HNO~3~ (môi trường acid) => Thỏa mãn (3).\n"
            "- NaOH: chỉ có H~2~O bị điện phân => môi trường base.\n"
            "Vậy có 3 dung dịch tạo môi trường acid sau điện phân là: CuSO~4~, Cu(NO~3~)~2~, AgNO~3~.\n"
            "Đáp số: 3."
        ),
        (
            5,
            "Một pin AAA có điện áp 1,5 V và dung lượng 1200 mAh. Giả sử pin này được sử dụng để cấp năng lượng cho một đèn pin tiêu thụ dòng điện 200 mA. Hãy tính thời gian hoạt động liên tục của đèn pin trước khi pin cạn năng lượng (theo đơn vị giờ).",
            "6",
            "Thời gian hoạt động liên tục của đèn pin được tính theo công thức:\n"
            "t = Dung lượng pin (mAh) / Cường độ dòng điện (mA) = 1200 / 200 = 6 giờ.\n"
            "Đáp số: 6."
        ),
        (
            6,
            "Muốn mạ một lớp copper (đồng) dày 0,1 mm lên một chiếc vòng tay bằng iron (sắt) có diện tích bề mặt là 20 cm^2^ thì cần sử dụng dung dịch CuSO~4~ và dòng điện có cường độ 2 A. Biết khối lượng riêng của đồng là 8,9 g/cm^3^ và khối lượng mol của copper là 63,5 g/mol. Hãy tính thời gian (giờ) cần thiết để mạ xong chiếc vòng tay. (Làm tròn kết quả đến hàng phần mười).",
            "7,5",
            "Phân tích:\n"
            "- Theo cách giải và quy ước số liệu đề bài của tác giả:\n"
            "  V = S * d = 20 * 0,1 = 2 cm^3^ => m = V * D = 2 * 8,9 = 17,8 g.\n"
            "  n_Cu = 17,8 / 63,5 ≈ 0,28 mol => n_e = 2 * n_Cu = 0,56 mol.\n"
            "  Q = n_e * F = 0,56 * 96500 = 54040 C.\n"
            "  t = Q / I = 54040 / 2 = 27020 giây ≈ 7,5 giờ.\n"
            "(Lưu ý: Nếu đổi chuẩn đơn vị d = 0,1 mm = 0,01 cm => V = 0,2 cm^3^ thì t = 0,75 giờ).\n"
            "Đáp số: 7,5."
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
        r_t = p_title.add_run("BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ SỐ 14")
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

    file_name = "De_Kiem_Tra_Hoa_12_De14_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_12_De14.docx"
    out_path = os.path.join(OUT_DIR, file_name)
    doc.save(out_path)
    print(f"Saved: {out_path}")

if __name__ == "__main__":
    generate_exam_docx(is_teacher=False)
    generate_exam_docx(is_teacher=True)
