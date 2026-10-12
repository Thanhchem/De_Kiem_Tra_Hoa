"""
Script tạo file Word (.docx) chuẩn "Form Đề Kiểm Tra" cho ĐỀ SỐ 24:
- Đề kiểm tra: Chương 8 - Sơ lược về dãy kim loại chuyển tiếp thứ nhất và Phức chất (Mã đề 124)
- Tạo 2 file:
  + San_Pham/De_Kiem_Tra_Hoa_12_De24.docx (Bản Học Sinh)
  + San_Pham/De_Kiem_Tra_Hoa_12_De24_GiaoVien.docx (Bản Giáo Viên)
"""

import os
import sys
import re
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
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
        hp.paragraph_format.space_after = Pt(0)
        hp.paragraph_format.space_before = Pt(0)

        tbl_h = header.add_table(rows=1, cols=2, width=Inches(7.08)) # 18.0 cm
        tbl_h.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_h.autofit = False

        c0 = tbl_h.cell(0, 0)
        c1 = tbl_h.cell(0, 1)
        c0.width = Inches(3.54)
        c1.width = Inches(3.54)
        set_cell_margins(c0, 40, 40, 60, 60)
        set_cell_margins(c1, 40, 40, 60, 60)

        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run("Tài liệu lưu hành nội bộ")
        r0.font.name = "Times New Roman"
        r0.font.size = Pt(10)
        r0.font.italic = True
        r0.font.color.rgb = RGBColor(211, 47, 47) # #D32F2F

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_after = Pt(0)
        r1_pre = p1.add_run("Thầy ")
        r1_pre.font.name = "Times New Roman"
        r1_pre.font.size = Pt(10)
        r1_pre.font.color.rgb = RGBColor(211, 47, 47)
        r1_bold = p1.add_run("TRẦN VĂN THẠNH")
        r1_bold.font.name = "Times New Roman"
        r1_bold.font.size = Pt(10)
        r1_bold.font.bold = True
        r1_bold.font.color.rgb = RGBColor(211, 47, 47)
        r1_post = p1.add_run(" - 0777.470.803")
        r1_post.font.name = "Times New Roman"
        r1_post.font.size = Pt(10)
        r1_post.font.color.rgb = RGBColor(211, 47, 47)

        # Border bottom: #8BC390 (sz=6)
        tblPr_h = tbl_h._tbl.tblPr
        tblBorders_h = parse_xml(
            r'<w:tblBorders %s><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="8BC390"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>' % nsdecls('w')
        )
        tblPr_h.append(tblBorders_h)

        # --- FOOTER ---
        footer = s.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.text = ""
        fp.paragraph_format.space_after = Pt(0)
        fp.paragraph_format.space_before = Pt(0)

        tbl_f = footer.add_table(rows=1, cols=2, width=Inches(7.08))
        tbl_f.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_f.autofit = False

        fc0 = tbl_f.cell(0, 0)
        fc1 = tbl_f.cell(0, 1)
        fc0.width = Inches(5.50)
        fc1.width = Inches(1.58)
        set_cell_margins(fc0, 40, 40, 60, 60)
        set_cell_margins(fc1, 40, 40, 60, 60)

        fp0 = fc0.paragraphs[0]
        fp0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fp0.paragraph_format.space_after = Pt(0)
        rf0 = fp0.add_run("CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế   |   CS2: 24 Đặng Thái Thân, TP Huế")
        rf0.font.name = "Times New Roman"
        rf0.font.size = Pt(9.5)
        rf0.font.italic = True
        rf0.font.color.rgb = RGBColor(211, 47, 47)

        fp1 = fc1.paragraphs[0]
        fp1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        fp1.paragraph_format.space_after = Pt(0)
        rf1 = fp1.add_run("Trang ")
        rf1.font.name = "Times New Roman"
        rf1.font.size = Pt(10)
        rf1.font.bold = True
        rf1.font.color.rgb = RGBColor(211, 47, 47)

        fldSimple = OxmlElement('w:fldSimple')
        fldSimple.set(qn('w:instr'), 'PAGE')
        r_fld = OxmlElement('w:r')
        rPr_fld = OxmlElement('w:rPr')
        rFonts_fld = OxmlElement('w:rFonts')
        rFonts_fld.set(qn('w:ascii'), 'Times New Roman')
        rFonts_fld.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts_fld.set(qn('w:cs'), 'Times New Roman')
        rPr_fld.append(rFonts_fld)
        b_fld = OxmlElement('w:b')
        rPr_fld.append(b_fld)
        col_fld = OxmlElement('w:color')
        col_fld.set(qn('w:val'), 'D32F2F')
        rPr_fld.append(col_fld)
        sz_fld = OxmlElement('w:sz')
        sz_fld.set(qn('w:val'), '20')
        rPr_fld.append(sz_fld)
        r_fld.append(rPr_fld)
        t_fld = OxmlElement('w:t')
        t_fld.text = "1"
        r_fld.append(t_fld)
        fldSimple.append(r_fld)
        fp1._p.append(fldSimple)

        # Border top: #8BC390 (sz=6)
        tblPr_f = tbl_f._tbl.tblPr
        tblBorders_f = parse_xml(
            r'<w:tblBorders %s><w:top w:val="single" w:sz="6" w:space="0" w:color="8BC390"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>' % nsdecls('w')
        )
        tblPr_f.append(tblBorders_f)

def add_exam_header(doc, is_teacher=False):
    lines = [
        ("SỞ GD&ĐT THỪA THIÊN HUẾ", "ĐỀ KIỂM TRA CHƯƠNG 8 -- NĂM HỌC 2026 - 2027", True, True, RGBColor(21, 101, 192)),
        ("TRƯỜNG THPT HAI BÀ TRƯNG", "MÔN: HOÁ HỌC -- KHỐI 12", True, True, RGBColor(211, 47, 47)),
        ("ĐỀ CHÍNH THỨC", "Thời gian làm bài: 50 phút (không kể thời gian phát đề)", False, False, None),
        ("(Đề thi có 04 trang)", "MÃ ĐỀ THI: 124", False, True, None)
    ]
    for left_txt, right_txt, left_bold, right_bold, right_color in lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.tab_stops.add_tab_stop(Inches(7.08), WD_TAB_ALIGNMENT.RIGHT)

        r_l = p.add_run(left_txt)
        r_l.font.name = "Times New Roman"
        r_l.font.size = Pt(10.5)
        r_l.bold = left_bold
        if not left_bold:
            r_l.italic = True

        p.add_run("\t")

        r_r = p.add_run(right_txt)
        r_r.font.name = "Times New Roman"
        r_r.font.size = Pt(10.5)
        r_r.bold = right_bold
        if right_color:
            r_r.font.color.rgb = right_color
        if not right_bold:
            r_r.italic = True

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(3)
    p_info.paragraph_format.space_after = Pt(3)
    p_info.paragraph_format.line_spacing = 1.15
    r_info = p_info.add_run("Họ, tên thí sinh: ..................................................................... Lớp: ................. SBD: .................")
    r_info.font.name = "Times New Roman"
    r_info.font.size = Pt(10)
    r_info.italic = True

    p_mass = doc.add_paragraph()
    p_mass.paragraph_format.space_before = Pt(0)
    p_mass.paragraph_format.space_after = Pt(6)
    p_mass.paragraph_format.line_spacing = 1.15
    r_mass = p_mass.add_run("(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Ti = 48; Cr = 52; Mn = 55; Fe = 56; Ni = 58,7; Cu = 64; Zn = 65; Ag = 108; Ba = 137; Pt = 195).")
    r_mass.font.name = "Times New Roman"
    r_mass.font.size = Pt(9.5)
    r_mass.italic = True

    if is_teacher:
        p_tea = doc.add_paragraph()
        p_tea.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_tea.paragraph_format.space_before = Pt(2)
        p_tea.paragraph_format.space_after = Pt(6)
        r_tea = p_tea.add_run("(BẢN GIÁO VIÊN -- CÓ ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI CHI TIẾT)")
        r_tea.font.name = "Times New Roman"
        r_tea.font.size = Pt(11)
        r_tea.bold = True
        r_tea.font.color.rgb = RGBColor(211, 47, 47)

def format_chemical_runs(paragraph, text, is_bold=False, is_italic=False, color_rgb=None):
    pattern = r'(\^\{[^\}]+\}|\^[0-9\+\-]+|\_\{[^\}]+\}|\_[0-9a-zA-Z\+\-]+)'
    tokens = re.split(pattern, text)
    for token in tokens:
        if not token:
            continue
        run = paragraph.add_run()
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)
        run.font.bold = is_bold
        run.font.italic = is_italic
        if color_rgb:
            run.font.color.rgb = color_rgb

        if token.startswith('^'):
            val = token[2:-1] if token.startswith('^{') else token[1:]
            run.text = val
            run.font.superscript = True
        elif token.startswith('_'):
            val = token[2:-1] if token.startswith('_{') else token[1:]
            run.text = val
            run.font.subscript = True
        else:
            run.text = token

def add_solution_paragraph(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.15
    
    r_lbl = p.add_run("Lời giải: ")
    r_lbl.font.name = "Times New Roman"
    r_lbl.font.size = Pt(10.5)
    r_lbl.bold = True
    r_lbl.italic = True
    r_lbl.font.color.rgb = RGBColor(21, 101, 192) # #1565C0
    
    format_chemical_runs(p, text, is_bold=False, is_italic=False, color_rgb=None)

def add_mcq_options(doc, opts, correct_idx=-1, is_teacher=False, cols=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    
    labels = ["A. ", "B. ", "C. ", "D. "]
    
    if cols == 4:
        p.paragraph_format.tab_stops.add_tab_stop(Inches(1.77), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(3.54), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(5.31), WD_TAB_ALIGNMENT.LEFT)
        for i, opt in enumerate(opts):
            if i > 0:
                p.add_run("\t")
            is_c = (i == correct_idx and is_teacher)
            col_rgb = RGBColor(211, 47, 47) if is_c else None
            format_chemical_runs(p, labels[i], is_bold=True, color_rgb=col_rgb)
            format_chemical_runs(p, opt, is_bold=is_c, color_rgb=col_rgb)
    elif cols == 2:
        p.paragraph_format.tab_stops.add_tab_stop(Inches(3.54), WD_TAB_ALIGNMENT.LEFT)
        for i in [0, 1]:
            if i == 1:
                p.add_run("\t")
            is_c = (i == correct_idx and is_teacher)
            col_rgb = RGBColor(211, 47, 47) if is_c else None
            format_chemical_runs(p, labels[i], is_bold=True, color_rgb=col_rgb)
            format_chemical_runs(p, opts[i], is_bold=is_c, color_rgb=col_rgb)
            
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(1)
        p2.paragraph_format.space_after = Pt(4)
        p2.paragraph_format.line_spacing = 1.15
        p2.paragraph_format.tab_stops.add_tab_stop(Inches(3.54), WD_TAB_ALIGNMENT.LEFT)
        for i in [2, 3]:
            if i == 3:
                p2.add_run("\t")
            is_c = (i == correct_idx and is_teacher)
            col_rgb = RGBColor(211, 47, 47) if is_c else None
            format_chemical_runs(p2, labels[i], is_bold=True, color_rgb=col_rgb)
            format_chemical_runs(p2, opts[i], is_bold=is_c, color_rgb=col_rgb)
    else:
        # 1 cột (4 dòng)
        for i in range(4):
            if i > 0:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
            is_c = (i == correct_idx and is_teacher)
            col_rgb = RGBColor(211, 47, 47) if is_c else None
            format_chemical_runs(p, labels[i], is_bold=True, color_rgb=col_rgb)
            format_chemical_runs(p, opts[i], is_bold=is_c, color_rgb=col_rgb)

def add_heading_section(doc, title, subtitle):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(title)
    r.font.name = "Times New Roman"
    r.font.size = Pt(11.5)
    r.bold = True
    r.font.color.rgb = RGBColor(21, 101, 192)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(4)
    p_sub.paragraph_format.line_spacing = 1.15
    r_sub = p_sub.add_run(subtitle)
    r_sub.font.name = "Times New Roman"
    r_sub.font.size = Pt(10.5)
    r_sub.italic = True

def build_exam_docx(is_teacher=False, output_path=""):
    doc = Document()
    add_header_footer(doc)
    add_exam_header(doc, is_teacher=is_teacher)

    # ------------------ PHẦN I ------------------
    add_heading_section(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    questions_p1 = [
        (
            "Câu 1. [KNTT - SBT] Nguyên tử Cr có cấu hình electron ở trạng thái cơ bản là [Ar]3d^5 4s^1. Trong phản ứng hoá học, khi nguyên tử Cr nhường đi 3 electron để tạo thành ion Cr^3+, số electron còn lại trên phân lớp 3d là",
            ["5.", "4.", "3.", "2."],
            2, 4,
            "Cấu hình electron của Cr: [Ar]3d^5 4s^1. Khi nhường 3 electron: nhường 1 electron ở phân lớp 4s và 2 electron ở phân lớp 3d, tạo cation Cr^3+ có cấu hình [Ar]3d^3. Số electron còn lại trên phân lớp 3d là 3.\n=> Chọn C."
        ),
        (
            "Câu 2. [KNTT - SBT] Dung dịch nào sau đây có màu vàng chanh?",
            ["CuSO_4.", "FeCl_3.", "KMnO_4.", "FeSO_4."],
            1, 4,
            "Dung dịch FeCl_3 có màu vàng chanh/vàng nâu đặc trưng (do sự thuỷ phân và tạo phức aqua của ion Fe^3+). Dung dịch CuSO_4 màu xanh lam, KMnO_4 màu tím, FeSO_4 màu xanh lục nhạt.\n=> Chọn B."
        ),
        (
            "Câu 3. [KNTT - SBT] Dãy kim loại nào sau đây sắp xếp theo thứ tự tăng dần nhiệt độ nóng chảy?",
            ["Na, Fe, Mg.", "Na, Mg, Fe.", "Fe, Mg, Na.", "Mg, Fe, Na."],
            1, 4,
            "Nhiệt độ nóng chảy: Na (97,8°C) < Mg (650°C) < Fe (1538°C). Dãy sắp xếp theo thứ tự tăng dần là Na, Mg, Fe.\n=> Chọn B."
        ),
        (
            "Câu 4. [CTST - SBT] Dung dịch muối chromium(III) sulfate có lẫn copper(II) sulfate, sử dụng dung dịch nào sau đây có thể loại bỏ nguyên tố copper ra khỏi muối chromium(III) sulfate?",
            ["Dung dịch ammonia.", "Dung dịch xút.", "Dung dịch barium chloride.", "Dung dịch soda."],
            0, 2,
            "Nhỏ dung dịch ammonia (NH_3) dư vào hỗn hợp: ion Cu^2+ tạo phức chất tan [Cu(NH_3)_4(H_2O)_2]^2+ có màu xanh thẫm, còn ion Cr^3+ tạo kết tủa Cr(OH)_3 không tan trong ammonia dư. Lọc lấy kết tủa Cr(OH)_3, sau đó cho tác dụng với lượng vừa đủ dung dịch H_2SO_4 loãng thu lại dung dịch Cr_2(SO_4)_3 tinh khiết.\n=> Chọn A."
        ),
        (
            "Câu 5. Phức chất là gì?",
            [
                "Phức chất là những hợp chất có cấu tạo rất phức tạp.",
                "Phức chất là những chất gồm nhiều phân tử chất đơn giản kết hợp lại.",
                "Phức chất là những chất được tạo thành từ hai hay nhiều phân tử muối.",
                "Phức chất là những chất khi phân li trong dung dịch tạo thành các ion phức tồn tại độc lập và có tính chất xác định."
            ],
            3, 1,
            "Theo định nghĩa SGK Hoá học 12, phức chất là những chất khi phân li trong dung dịch tạo thành các ion phức (hoặc bản thân là phân tử phức trung hoà) tồn tại độc lập và có tính chất hoá học xác định.\n=> Chọn D."
        ),
        (
            "Câu 6. Cho phát biểu sau: “Phức chất đơn giản thường có một.(1). liên kết với các phối tử bao quanh. Liên kết giữa nguyên tử trung tâm và phối tử trong phức chất là liên kết.(2).”. Cụm từ cần điền vào (1) và (2) lần lượt là",
            ["cation kim loại, ion.", "nguyên tử kim loại, cho - nhận.", "nguyên tử trung tâm, cho - nhận.", "phối tử, ion."],
            2, 2,
            "Phức chất đơn giản có một nguyên tử trung tâm liên kết với các phối tử bao quanh qua liên kết phối trí (liên kết cho - nhận).\n=> Chọn C."
        ),
        (
            "Câu 7. Phức chất nào sau đây chứa ion phức có số phối trí là 4?",
            ["[Zn(NH_3)_4]^2+.", "[Ag(NH_3)_2]^+.", "[Co(NH_3)_6]^3+.", "[Ni(CN)_4]^2-."],
            0, 4,
            "Trong ion phức [Zn(NH_3)_4]^2+, nguyên tử trung tâm Zn^2+ liên kết với 4 phối tử một càng NH_3 nên có số phối trí là 4 (cấu trúc tứ diện). (Lưu ý: ion phức [Ni(CN)_4]^2- cũng có số phối trí là 4 dạng vuông phẳng).\n=> Chọn A."
        ),
        (
            "Câu 8. Trong phức chất [Cr(H_2O)_6]Cl_3, ion phức có dạng hình học gì?",
            ["Hình vuông.", "Lập phương.", "Hình bát diện.", "Hình tứ diện."],
            2, 4,
            "Cation phức [Cr(H_2O)_6]^3+ có 6 phối tử nước liên kết với ion trung tâm Cr^3+, có số phối trí 6 nên có cấu trúc hình học bát diện (octahedron).\n=> Chọn C."
        ),
        (
            "Câu 9. Chọn phát biểu sai:",
            [
                "Phức chất có thể mang điện tích hoặc không mang điện tích.",
                "Phức chất mà nguyên tử trung tâm tạo 4 liên kết với các phối tử luôn có dạng hình học là tứ diện.",
                "Giống như phân tử ammonia, phân tử methylamine cũng có thể đóng vai trò là phối tử do có cặp electron chưa liên kết.",
                "Phức có 4 phối tử và nguyên tử trung tâm có AO d^8 sẽ có dạng hình học vuông phẳng."
            ],
            1, 1,
            "Phát biểu B sai vì phức chất có 4 phối tử có thể có cấu trúc tứ diện (như [Zn(NH_3)_4]^2+) hoặc vuông phẳng (như cisplatin [Pt(NH_3)_2Cl_2], [PtCl_4]^2-, [Ni(CN)_4]^2-).\n=> Chọn B."
        ),
        (
            "Câu 10. Trong phương trình hoá học dạng tổng quát của phức chất aqua: M^n+(aq) + mH_2O(l) -> [M(OH_2)_m]^n+(aq). Giá trị n là:",
            ["số electron hoá trị của cation kim loại M.", "giá trị điện tích của cation kim loại M.", "số electron độc thân của cation kim loại M.", "hoá trị của cation kim loại M."],
            1, 2,
            "Trong kí hiệu cation M^n+, n là độ lớn giá trị điện tích của cation kim loại M.\n=> Chọn B."
        ),
        (
            "Câu 11. Phản ứng tạo phức chất thường có một số dấu hiệu dễ quan sát như",
            ["sự hoà tan, sự ngưng tụ.", "sự biến đổi màu sắc, sự thay đổi nhiệt độ nóng chảy.", "sự biến đổi màu sắc, sự hoà tan, sự kết tủa.", "sự hoà tan, sự thay đổi nhiệt độ sôi."],
            2, 1,
            "Các dấu hiệu dễ quan sát nhất khi phản ứng tạo phức chất xảy ra trong dung dịch là sự đổi màu dung dịch, sự hoà tan kết tủa hoặc sự xuất hiện kết tủa màu đặc trưng.\n=> Chọn C."
        ),
        (
            "Câu 12. Khi cho dung dịch ammonia dư vào dung dịch chứa phức chất [Ni(OH_2)_6]^2+ và anion Cl^-, có phản ứng sau: [Ni(OH_2)_6]^2+(aq) + 6NH_3(aq) -> [Ni(NH_3)_6]^2+(aq) + 6H_2O(l) (*). Phát biểu nào dưới đây là không đúng?",
            [
                "Trong điều kiện của phản ứng (*), phức chất [Ni(NH_3)_6]^2+ kém bền hơn phức chất [Ni(OH_2)_6]^2+.",
                "Phản ứng (*) là phản ứng thế phối tử.",
                "Dung dịch sau phản ứng có pH > 7.",
                "Trong phản ứng không có sự thay đổi số oxi hoá của các nguyên tố."
            ],
            0, 1,
            "Phản ứng (*) diễn ra thuận lợi theo chiều thuận chứng tỏ phức chất [Ni(NH_3)_6]^2+ bền hơn phức chất aqua [Ni(OH_2)_6]^2+. Do đó phát biểu A nói phức amine kém bền hơn là sai.\n=> Chọn A."
        ),
        (
            "Câu 13. Nhỏ vài giọt dung dịch NH_3 vào dung dịch CuSO_4 loãng thấy xuất hiện kết tủa màu xanh nhạt, chứng tỏ phức chất [Cu(OH)_2(H_2O)_4] đã được tạo thành. Tiếp tục nhỏ vào ống nghiệm dung dịch NH_3 đến dư thấy kết tủa tan, đồng thời dung dịch chuyển sang màu xanh lam, chứng tỏ phức chất [Cu(NH_3)_4(H_2O)_2]^2+ đã được tạo thành. Dấu hiệu của phản ứng tạo phức trên dựa vào:",
            ["sự biến đổi màu sắc.", "sự biến đổi màu sắc và sự hoà tan.", "sự hoà tan.", "sự kết tủa."],
            1, 2,
            "Quá trình cho NH_3 đến dư có hiện tượng kết tủa tan dần và dung dịch đổi màu từ xanh nhạt sang xanh lam thẫm, do đó dựa vào sự biến đổi màu sắc và sự hoà tan kết tủa.\n=> Chọn B."
        ),
        (
            "Câu 14. Phản ứng chuẩn độ Fe^2+ trong dung dịch acid bằng dung dịch KMnO_4 được biểu diễn bởi phương trình ion rút gọn sau: MnO_4^-(aq) + 5Fe^2+(aq) + 8H^+(aq) -> Mn^2+(aq) + 5Fe^3+(aq) + 4H_2O(l). Chất oxi hoá trong phản ứng trên là",
            ["Fe^2+(aq).", "Mn^2+(aq).", "MnO_4^-(aq).", "H^+(aq)."],
            2, 4,
            "Trong phản ứng, ion MnO_4^- có số oxi hoá của Mn giảm từ +7 xuống +2 (nhận 5e) nên MnO_4^- là chất oxi hoá. Fe^2+ nhường 1e là chất khử.\n=> Chọn C."
        ),
        (
            "Câu 15. Có thể loại bỏ sắt trong nước ngầm bằng cách đưa nước ngầm lên các bể lắng, lọc và sục không khí vào. Khi đó Fe^2+ trong nước ngầm bị chuyển thành Fe(OH)_3 và lắng xuống. Chất oxi hoá Fe^2+ thành Fe(OH)_3 là",
            ["khí O_2.", "khí N_2.", "nước và khí O_2.", "khí O_2 và OH^-."],
            0, 4,
            "Khí oxygen (O_2) trong không khí sục vào oxi hoá Fe^2+ thành Fe(OH)_3 theo phương trình: 4Fe^2+ + O_2 + 10H_2O -> 4Fe(OH)_3 + 8H^+. Số oxi hoá của oxygen giảm từ 0 xuống -2, là chất oxi hoá.\n=> Chọn A."
        ),
        (
            "Câu 16. Phức chất [Fe(OH_2)_6]^3+ có cấu tạo như hình sau:\n[IMG:San_Pham/images_crop_de24/p1_cau16_fe_oh2_6.png]\nSố liên kết cho – nhận mà nguyên tử trung tâm iron tạo được với các phối tử OH_2 là",
            ["5.", "4.", "3.", "6."],
            3, 4,
            "Cation Fe^3+ ở tâm hình bát diện liên kết với 6 phối tử nước OH_2 qua 6 liên kết cho - nhận (phối trí).\n=> Chọn D."
        ),
        (
            "Câu 17. Trong các số oxi hoá của các kim loại chuyển tiếp dãy thứ nhất, trạng thái thường gặp nhất là +2. Điều này được giải thích là do đa số các kim loại chuyển tiếp dãy thứ nhất có độ âm điện không cao và",
            [
                "dễ nhường đi 2 electron ở phân lớp 3d.",
                "dễ nhường đi 1 electron ở phân lớp 3d và 1 electron ở phân lớp 4s.",
                "có khả năng nhận thêm 2 electron vào phân lớp 3d.",
                "có 2 electron lớp ngoài cùng."
            ],
            3, 1,
            "Đa số kim loại chuyển tiếp dãy thứ nhất có cấu hình electron hoá trị 3d^n 4s^2, lớp electron ngoài cùng có 2 electron ở phân lớp 4s dễ bị nhường đi để tạo cation mang số oxi hoá +2.\n=> Chọn D."
        ),
        (
            "Câu 18. Một học sinh tiến hành chuẩn độ dung dịch Fe^2+ bằng dung dịch KMnO_4 theo hai cách như sau:\nCách 1. Nhỏ từ từ dung dịch KMnO_4 vào dung dịch chứa Fe^2+ trong môi trường acid cho đến khi xuất hiện màu hồng nhạt bền trong khoảng 30 giây.\nCách 2. Nhỏ từ từ dung dịch chứa Fe^2+ trong môi trường acid vào dung dịch KMnO_4 cho đến khi màu hồng của dung dịch KMnO_4 biến mất.\nHãy cho biết cách tiến hành chuẩn độ nào là phù hợp.",
            ["Cách 1.", "Cách 2.", "Cả hai cách.", "Không có cách nào."],
            0, 2,
            "Cách 1 phù hợp vì dung dịch KMnO_4 đóng vai trò là chất chuẩn lẫn chất chỉ thị màu; khi đạt điểm tương đương, chỉ cần dư 1 giọt KMnO_4 sẽ làm dung dịch xuất hiện màu hồng nhạt rất dễ quan sát. Cách 2 quan sát màu hồng biến mất hoàn toàn rất khó chính xác.\n=> Chọn A."
        )
    ]

    for q_txt, opts, corr_idx, cols, sol_txt in questions_p1:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        
        parts = q_txt.split("\n")
        for idx_part, part in enumerate(parts):
            if part.startswith("[IMG:"):
                img_path = part[5:-1]
                if os.path.exists(img_path):
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.paragraph_format.space_before = Pt(2)
                    p_img.paragraph_format.space_after = Pt(2)
                    run_img = p_img.add_run()
                    run_img.add_picture(img_path, width=Inches(2.2))
            else:
                if idx_part > 0:
                    p_q = doc.add_paragraph()
                    p_q.paragraph_format.space_before = Pt(1)
                    p_q.paragraph_format.space_after = Pt(2)
                    p_q.paragraph_format.line_spacing = 1.15
                format_chemical_runs(p_q, part, is_bold=False)

        add_mcq_options(doc, opts, correct_idx=corr_idx, is_teacher=is_teacher, cols=cols)
        if is_teacher:
            add_solution_paragraph(doc, sol_txt)

    # ------------------ PHẦN II ------------------
    add_heading_section(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    questions_p2 = [
        (
            "Câu 1. Titanium (Ti) được sử dụng để chế tạo động cơ hoặc khung máy bay, còn titanium dioxide (TiO_2) được sử dụng trong sản xuất kem chống nắng. Quặng ilmenite (FeTiO_3) có giá trị quan trọng về mặt kinh tế để sản xuất kim loại titanium. TiO_2 sản xuất từ FeTiO_3 (phản ứng A), sau đó được chuyển hoá thành Ti theo quy trình Kroll (phản ứng B, C, D):\n[IMG:San_Pham/images_crop_de24/p2_cau1_quy_trinh_kroll.png]\nXét tính đúng/sai của các phát biểu sau:",
            [
                ("a) Phương trình hoá học của phản ứng A là: FeTiO_3 + C -> Fe + TiO_2 + CO.", True),
                ("b) Phương trình hoá học của phản ứng B là: TiO_2 + 2Cl_2 -> TiCl_4 + O_2.", False),
                ("c) Phương trình hoá học của phản ứng C là: TiCl_4 + 2Mg -> 2MgCl_2 + Ti.", True),
                ("d) Phương trình tổng cộng của 4 phản ứng là: FeTiO_3 + 3C -> Fe + Ti + 3CO.", True)
            ],
            "a) Đúng. Quá trình hoàn nguyên nhiệt cacbon có chọn lọc quặng ilmenite tạo kim loại Fe và TiO_2.\nb) Sai. Phản ứng chlorine hoá TiO_2 cần có mặt carbon ở nhiệt độ cao: TiO_2 + 2C + 2Cl_2 -> TiCl_4 + 2CO, không sinh ra O_2.\nc) Đúng. Khử TiCl_4 bằng kim loại Mg nóng chảy trong khí trơ argon ở khoảng 800 - 850°C.\nd) Đúng. Phản ứng D là điện phân nóng chảy tái chế MgCl_2 -> Mg + Cl_2; cộng cả 4 phản ứng thu được: FeTiO_3 + 3C -> Fe + Ti + 3CO."
        ),
        (
            "Câu 2. Phức chất cis-[Pt(NH_3)_2Cl_2] là một hợp chất vô cơ có khả năng ức chế các tế bào ung thư như ung thư tinh hoàn, buồng trứng, bàng quang và các khối u ở đầu và cổ. Mỗi phát biểu sau là đúng hay sai?",
            [
                ("a) Ion trung tâm là Pt.", False),
                ("b) Phối tử là NH_3.", False),
                ("c) Ion trung tâm là Pt^2+, phối tử là NH_3 và Cl^-.", True),
                ("d) Phức có dạng hình học là vuông phẳng.", True)
            ],
            "a) Sai. Cation trung tâm là ion Pt^2+ (nguyên tử Pt ở trạng thái oxi hoá +2).\nb) Sai. Phối tử gồm cả 2 phân tử NH_3 và 2 ion Cl^-; phát biểu chỉ nêu NH_3 là chưa đầy đủ.\nc) Đúng. Cation trung tâm là Pt^2+, các phối tử liên kết là NH_3 và Cl^-.\nd) Đúng. Phức chất cisplatin có dạng hình học vuông phẳng (square planar)."
        ),
        (
            "Câu 3. [KNTT - SBT] Trong dung dịch, ion Fe^3+ tồn tại dưới dạng phức chất aqua có sáu phối tử nước. Xét tính đúng/sai của các phát biểu sau:",
            [
                ("a) Phức chất aqua có công thức hoá học là [Fe(H_2O)_6]^3+.", True),
                ("b) Phức chất aqua có dạng hình học vuông phẳng.", False),
                ("c) 6 phối tử nước đã cho cặp electron chưa liên kết vào ion Fe^3+.", True),
                ("d) Nguyên tử trung tâm trong phức chất aqua là Fe^2+.", False)
            ],
            "a) Đúng. Phức chất aqua của sắt(III) có công thức là [Fe(H_2O)_6]^3+.\nb) Sai. Phức có 6 phối tử nước nên có dạng hình học bát diện (octahedral).\nc) Đúng. Mỗi phân tử nước dùng 1 cặp electron chưa liên kết của nguyên tử O để tạo liên kết phối trí với cation Fe^3+.\nd) Sai. Ion trung tâm là Fe^3+, không phải Fe^2+."
        ),
        (
            "Câu 4. Cho hai phức chất A và B có công thức cấu tạo lần lượt sau:\n[IMG:San_Pham/images_crop_de24/p2_cau4_phuc_A_B.png]\nMỗi phát biểu sau đây đúng hay sai?",
            [
                ("a) Nguyên tử trung tâm của hai phức chất đều là nguyên tố kim loại chuyển tiếp.", False),
                ("b) Trong phức chất B có 4 phối tử.", False),
                ("c) Hai phức chất A và B có dạng hình học khác nhau.", True),
                ("d) Trong A và trong B đều có hai loại phối tử.", False)
            ],
            "a) Sai. Phức A có nguyên tử trung tâm là Co (kim loại chuyển tiếp), còn phức B (chlorophyll) có nguyên tử trung tâm là Mg (kim loại kiềm thổ thuộc nhóm IIA - nhóm chính, không phải kim loại chuyển tiếp).\nb) Sai. Trong chlorophyll B chỉ có 1 phối tử vòng 4 càng porphyrin liên kết qua 4 nguyên tử N, không phải 4 phối tử độc lập.\nc) Đúng. Phức A có dạng hình học bát diện, còn phức B có dạng phẳng vòng càng (square planar quanh nguyên tử Mg).\nd) Sai. Phức A có 2 loại phối tử (NH_3 và NO_2^-), nhưng phức B chỉ có 1 loại phối tử là dẫn xuất chlorin/porphyrin."
        )
    ]

    for q_txt, items, sol_txt in questions_p2:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        
        parts = q_txt.split("\n")
        for idx_part, part in enumerate(parts):
            if part.startswith("[IMG:"):
                img_path = part[5:-1]
                if os.path.exists(img_path):
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.paragraph_format.space_before = Pt(2)
                    p_img.paragraph_format.space_after = Pt(2)
                    run_img = p_img.add_run()
                    w = Inches(2.2) if "p2_cau1" in img_path else Inches(4.5)
                    run_img.add_picture(img_path, width=w)
            else:
                if idx_part > 0:
                    p_q = doc.add_paragraph()
                    p_q.paragraph_format.space_before = Pt(1)
                    p_q.paragraph_format.space_after = Pt(2)
                    p_q.paragraph_format.line_spacing = 1.15
                format_chemical_runs(p_q, part, is_bold=False)

        for stmt, is_true in items:
            p_item = doc.add_paragraph()
            p_item.paragraph_format.left_indent = Inches(0.2)
            p_item.paragraph_format.space_before = Pt(1)
            p_item.paragraph_format.space_after = Pt(2)
            p_item.paragraph_format.line_spacing = 1.15
            format_chemical_runs(p_item, stmt, is_bold=False)
            if is_teacher:
                r_ans = p_item.add_run(" --> " + ("ĐÚNG" if is_true else "SAI"))
                r_ans.font.name = "Times New Roman"
                r_ans.font.size = Pt(10.5)
                r_ans.bold = True
                r_ans.font.color.rgb = RGBColor(211, 47, 47) if is_true else RGBColor(198, 40, 40)

        if is_teacher:
            add_solution_paragraph(doc, sol_txt)

    # ------------------ PHẦN III ------------------
    add_heading_section(
        doc,
        "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 6."
    )

    questions_p3 = [
        (
            "Câu 1. Phức chất [Co(NH_3)_4Cl_2]^+ có cấu trúc như hình sau:\n[IMG:San_Pham/images_crop_de24/p3_cau1_co_nh3_4_cl2.png]\nCó bao nhiêu liên kết sigma (σ) có trong phức chất đó?",
            "18",
            "• Trong 4 phối tử NH_3: mỗi phân tử NH_3 có 3 liên kết đơn N-H (liên kết σ), tổng cộng 4 × 3 = 12 liên kết σ N-H.\n• Cation trung tâm Co liên kết với 4 nguyên tử N qua 4 liên kết phối trí Co-N (đều là liên kết σ).\n• Cation trung tâm Co liên kết với 2 ion Cl^- qua 2 liên kết Co-Cl (đều là liên kết σ).\n• Tổng số liên kết σ trong phức chất = 12 + 4 + 2 = 18 liên kết σ.\n=> Đáp số: 18."
        ),
        (
            "Câu 2. [KNTT - SBT] Phức chất [MA_x B_2] có dạng hình học bát diện. Ở đó M là nguyên tử trung tâm, x là số phối tử của A. Giá trị của x là bao nhiêu?",
            "4",
            "Phức chất có dạng hình học bát diện nên có số phối trí là 6. Với các phối tử một càng A và B, tổng số phối tử là x + 2 = 6 => x = 4.\n=> Đáp số: 4."
        ),
        (
            "Câu 3. [KNTT - SBT] Phức chất [Co(NH_3)Cl_x]^y- có dạng hình học bát diện, nguyên tử trung tâm là Co^3+. Tổng giá trị của x và y là bao nhiêu?",
            "7",
            "• Phức bát diện có tổng số phối tử một càng là 1 + x = 6 => x = 5.\n• Điện tích ion phức: Co^3+ (+3), NH_3 (0), 5 phối tử Cl^- (-5) => điện tích ion phức là (+3) + 0 + 5×(-1) = -2 => y = 2.\n• Tổng giá trị của x và y = 5 + 2 = 7.\n=> Đáp số: 7."
        ),
        (
            "Câu 4. [KNTT - SBT] Cho dung dịch NH_3 đặc vào dung dịch phức chất [PtCl_4]^2- thu được phức chất có điện tích +1 là do một số phối tử Cl^- trong phức [PtCl_4]^2- bị thay thế bởi phối tử NH_3. Số lượng phối tử Cl^- đã bị thay thế là bao nhiêu?",
            "3",
            "• Trong [PtCl_4]^2-, nguyên tử trung tâm Pt có số oxi hoá +2.\n• Khi thay k phối tử Cl^- bằng k phân tử trung hoà NH_3, phức chất mới có dạng [Pt(NH_3)_k Cl_{4-k}]^{(k-2)+}.\n• Điện tích phức chất mới bằng +1 => k - 2 = +1 => k = 3.\n• Vậy số lượng phối tử Cl^- bị thay thế là 3 (tạo ra phức chất [Pt(NH_3)_3Cl]^+).\n=> Đáp số: 3."
        ),
        (
            "Câu 5. Sự gia tăng hàm lượng nitrate trong nước là một trong những nguyên nhân của hiện tượng phú dưỡng. Một trong những quy trình xác định hàm lượng nitrate trong nước được thực hiện như sau:\nThí nghiệm 1. Lấy 10,0 mL dung dịch muối Mohr [(NH_4)_2SO_4.FeSO_4.6H_2O], thêm H_2SO_4 1M vào và chuẩn độ bằng dung dịch KMnO_4.\nThí nghiệm 2. Lấy 10,0 mL dung dịch muối Mohr cho vào 100 mL nước chứa ion nitrate, sau đó thêm NaOH(s) vào để đạt khoảng 28%, khi đó muối Mohr khử nitrate thành ammonia theo các phản ứng:\nNO_3^-(aq) + 2Fe(OH)_2 + H_2O -> NO_2^-(aq) + 2Fe(OH)_3 (1)\nNO_2^-(aq) + 6Fe(OH)_2 + 5H_2O -> NH_3 + 6Fe(OH)_3 + OH^- (2)\nSau khi các phản ứng (1) và (2) hoàn toàn, lọc bỏ kết tủa, acid hoá dung dịch bằng dung dịch H_2SO_4 1M và chuẩn độ lượng Fe^2+ dư bằng dung dịch KMnO_4 (các chất và ion khác trong dung dịch không phản ứng với KMnO_4).\nHãy cho biết mỗi mmol KMnO_4 chênh lệch giữa thí nghiệm 1 với thí nghiệm 2 sẽ tương ứng với bao nhiêu mg NO_3^- trong nước? (Làm tròn kết quả đến hàng phần mười).",
            "38,8",
            "• Ở TN1: 5Fe^2+ + MnO_4^- + 8H^+ -> 5Fe^3+ + Mn^2+ + 4H_2O => 1 mmol KMnO_4 tương ứng 5 mmol Fe^2+.\n• Ở TN2: Cộng 2 phương trình khử NO_3^-:\nNO_3^- + 8Fe(OH)_2 + 6H_2O -> NH_3 + 8Fe(OH)_3 + OH^- => 1 mmol NO_3^- phản ứng với 8 mmol Fe^2+.\n• Lượng KMnO_4 chênh lệch:\nΔn_{KMnO_4} = n_{Fe^2+ (p/u)} / 5 = (8 / 5) n_{NO_3^-} = 1,6 n_{NO_3^-} => n_{NO_3^-} = Δn_{KMnO_4} / 1,6 = 0,625 Δn_{KMnO_4}.\n• Với mỗi 1 mmol KMnO_4 chênh lệch:\nm_{NO_3^-} = 0,625 × 62 = 38,75 mg ≈ 38,8 mg.\n=> Đáp số: 38,8."
        ),
        (
            "Câu 6. Chromium (Cr) có tính cứng cao. Nickel (Ni) có đặc tính cơ học là cứng, dễ dát mỏng, dễ uốn, dễ kéo sợi và đặc tính hoá học là trơ với không khí nên khi thêm vào inox để việc gia công dễ dàng hơn. Inox 18/10 có thành phần hoá học gồm 18% chromium, 10% nickel và tối đa 0,08% carbon (C). Xác định khối lượng iron (tính theo kg) tối đa trong 1,00 tấn inox 18/10.",
            "720",
            "• Trong 1,00 tấn (1000 kg) inox 18/10 có: 18% Cr, 10% Ni, tối đa 0,08% C và phần còn lại là Fe.\n• Để khối lượng iron (Fe) đạt tối đa thì hàm lượng carbon và các tạp chất khác phải ở mức tối thiểu bằng 0%.\n• Khi đó % khối lượng iron tối đa là: %m_{Fe (max)} = 100% - 18% - 10% = 72%.\n• Khối lượng iron tối đa thu được là: m_{Fe (max)} = 1000 × 72% = 720 kg.\n=> Đáp số: 720."
        )
    ]

    for q_txt, ans_val, sol_txt in questions_p3:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        
        parts = q_txt.split("\n")
        for idx_part, part in enumerate(parts):
            if part.startswith("[IMG:"):
                img_path = part[5:-1]
                if os.path.exists(img_path):
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.paragraph_format.space_before = Pt(2)
                    p_img.paragraph_format.space_after = Pt(2)
                    run_img = p_img.add_run()
                    run_img.add_picture(img_path, width=Inches(1.8))
            else:
                if idx_part > 0:
                    p_q = doc.add_paragraph()
                    p_q.paragraph_format.space_before = Pt(1)
                    p_q.paragraph_format.space_after = Pt(2)
                    p_q.paragraph_format.line_spacing = 1.15
                format_chemical_runs(p_q, part, is_bold=False)

        if not is_teacher:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Inches(0.2)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(4)
            r_ans_lbl = p_ans.add_run("Trả lời: .........................................................................................................................................................")
            r_ans_lbl.font.name = "Times New Roman"
            r_ans_lbl.font.size = Pt(10.5)
            r_ans_lbl.italic = True
        else:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Inches(0.2)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(2)
            r_ans_lbl = p_ans.add_run(f"Đáp số: {ans_val}")
            r_ans_lbl.font.name = "Times New Roman"
            r_ans_lbl.font.size = Pt(11)
            r_ans_lbl.bold = True
            r_ans_lbl.font.color.rgb = RGBColor(211, 47, 47)
            add_solution_paragraph(doc, sol_txt)

    # ------------------ BẢNG ĐÁP ÁN (BẢN GIÁO VIÊN) ------------------
    if is_teacher:
        p_tbl_title = doc.add_paragraph()
        p_tbl_title.paragraph_format.space_before = Pt(12)
        p_tbl_title.paragraph_format.space_after = Pt(4)
        p_tbl_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_tbl_title = p_tbl_title.add_run("BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ THI")
        r_tbl_title.font.name = "Times New Roman"
        r_tbl_title.font.size = Pt(12)
        r_tbl_title.bold = True
        r_tbl_title.font.color.rgb = RGBColor(21, 101, 192)

        # 1. Bảng Phần I
        p_p1 = doc.add_paragraph()
        p_p1.paragraph_format.space_before = Pt(4)
        p_p1.paragraph_format.space_after = Pt(2)
        r_p1 = p_p1.add_run("PHẦN I (Mỗi câu trả lời đúng được 0,25 điểm)")
        r_p1.font.name = "Times New Roman"
        r_p1.font.size = Pt(10.5)
        r_p1.bold = True

        tbl_p1 = doc.add_table(rows=3, cols=10)
        tbl_p1.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_p1.autofit = False
        tblPr1 = tbl_p1._tbl.tblPr
        tblBorders1 = parse_xml(
            r'<w:tblBorders %s><w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/></w:tblBorders>' % nsdecls('w')
        )
        tblPr1.append(tblBorders1)

        p1_keys = ["C", "B", "B", "A", "D", "C", "A", "C", "B", "B", "C", "A", "B", "C", "A", "D", "D", "A"]
        # Hàng 1: Câu 1-9
        for c in range(9):
            cell = tbl_p1.cell(0, c)
            cell.text = f"Câu {c+1}"
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].font.name = "Times New Roman"
            cell.paragraphs[0].runs[0].font.size = Pt(9.5)
            cell.paragraphs[0].runs[0].font.bold = True
            set_cell_margins(cell, 30, 30, 30, 30)

            cell_ans = tbl_p1.cell(1, c)
            cell_ans.text = p1_keys[c]
            cell_ans.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell_ans.paragraphs[0].runs[0].font.name = "Times New Roman"
            cell_ans.paragraphs[0].runs[0].font.size = Pt(10)
            cell_ans.paragraphs[0].runs[0].font.bold = True
            cell_ans.paragraphs[0].runs[0].font.color.rgb = RGBColor(211, 47, 47)
            set_cell_margins(cell_ans, 30, 30, 30, 30)

        # Câu 10-18
        tbl_p1_2 = doc.add_table(rows=2, cols=9)
        tbl_p1_2.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_p1_2.autofit = False
        tblPr1_2 = tbl_p1_2._tbl.tblPr
        tblBorders1_2 = parse_xml(
            r'<w:tblBorders %s><w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/></w:tblBorders>' % nsdecls('w')
        )
        tblPr1_2.append(tblBorders1_2)

        for c in range(9):
            idx = c + 9
            cell = tbl_p1_2.cell(0, c)
            cell.text = f"Câu {idx+1}"
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].font.name = "Times New Roman"
            cell.paragraphs[0].runs[0].font.size = Pt(9.5)
            cell.paragraphs[0].runs[0].font.bold = True
            set_cell_margins(cell, 30, 30, 30, 30)

            cell_ans = tbl_p1_2.cell(1, c)
            cell_ans.text = p1_keys[idx]
            cell_ans.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell_ans.paragraphs[0].runs[0].font.name = "Times New Roman"
            cell_ans.paragraphs[0].runs[0].font.size = Pt(10)
            cell_ans.paragraphs[0].runs[0].font.bold = True
            cell_ans.paragraphs[0].runs[0].font.color.rgb = RGBColor(211, 47, 47)
            set_cell_margins(cell_ans, 30, 30, 30, 30)

        # 2. Bảng Phần II
        p_p2 = doc.add_paragraph()
        p_p2.paragraph_format.space_before = Pt(6)
        p_p2.paragraph_format.space_after = Pt(2)
        r_p2 = p_p2.add_run("PHẦN II (Mỗi câu trả lời đúng 1 ý: 0,1 đ; 2 ý: 0,25 đ; 3 ý: 0,5 đ; 4 ý: 1,0 đ)")
        r_p2.font.name = "Times New Roman"
        r_p2.font.size = Pt(10.5)
        r_p2.bold = True

        tbl_p2 = doc.add_table(rows=5, cols=5)
        tbl_p2.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_p2.autofit = False
        tblPr2 = tbl_p2._tbl.tblPr
        tblBorders2 = parse_xml(
            r'<w:tblBorders %s><w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/></w:tblBorders>' % nsdecls('w')
        )
        tblPr2.append(tblBorders2)

        # Header cols
        tbl_p2.cell(0, 0).text = "Câu"
        tbl_p2.cell(0, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        tbl_p2.cell(0, 0).paragraphs[0].runs[0].font.bold = True
        for i in range(4):
            c = tbl_p2.cell(0, i+1)
            c.text = f"Câu {i+1}"
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            c.paragraphs[0].runs[0].font.bold = True

        p2_keys = [
            ("a", ["Đ", "S", "Đ", "S"]),
            ("b", ["S", "S", "S", "S"]),
            ("c", ["Đ", "Đ", "Đ", "Đ"]),
            ("d", ["Đ", "Đ", "S", "S"])
        ]
        for row_idx, (lbl, vals) in enumerate(p2_keys):
            c_lbl = tbl_p2.cell(row_idx+1, 0)
            c_lbl.text = f"Lệnh {lbl}"
            c_lbl.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            c_lbl.paragraphs[0].runs[0].font.bold = True
            for col_idx, val in enumerate(vals):
                cell_v = tbl_p2.cell(row_idx+1, col_idx+1)
                cell_v.text = val
                cell_v.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell_v.paragraphs[0].runs[0].font.bold = True
                cell_v.paragraphs[0].runs[0].font.color.rgb = RGBColor(211, 47, 47) if val == "Đ" else RGBColor(50, 50, 50)

        # 3. Bảng Phần III
        p_p3 = doc.add_paragraph()
        p_p3.paragraph_format.space_before = Pt(6)
        p_p3.paragraph_format.space_after = Pt(2)
        r_p3 = p_p3.add_run("PHẦN III (Mỗi câu trả lời đúng được 0,25 điểm)")
        r_p3.font.name = "Times New Roman"
        r_p3.font.size = Pt(10.5)
        r_p3.bold = True

        tbl_p3 = doc.add_table(rows=2, cols=6)
        tbl_p3.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_p3.autofit = False
        tblPr3 = tbl_p3._tbl.tblPr
        tblBorders3 = parse_xml(
            r'<w:tblBorders %s><w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/></w:tblBorders>' % nsdecls('w')
        )
        tblPr3.append(tblBorders3)

        p3_keys = ["18", "4", "7", "3", "38,8", "720"]
        for c in range(6):
            cell = tbl_p3.cell(0, c)
            cell.text = f"Câu {c+1}"
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].font.bold = True
            
            cell_ans = tbl_p3.cell(1, c)
            cell_ans.text = p3_keys[c]
            cell_ans.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell_ans.paragraphs[0].runs[0].font.bold = True
            cell_ans.paragraphs[0].runs[0].font.color.rgb = RGBColor(211, 47, 47)

    doc.save(output_path)
    print(f"Đã lưu: {output_path}")

def main():
    out_dir = r"c:\Antigravity_Thanh\San_Pham"
    os.makedirs(out_dir, exist_ok=True)
    
    # 1. Bản Học sinh
    hs_path = os.path.join(out_dir, "De_Kiem_Tra_Hoa_12_De24.docx")
    build_exam_docx(is_teacher=False, output_path=hs_path)
    
    # 2. Bản Giáo viên
    gv_path = os.path.join(out_dir, "De_Kiem_Tra_Hoa_12_De24_GiaoVien.docx")
    build_exam_docx(is_teacher=True, output_path=gv_path)

if __name__ == "__main__":
    main()
