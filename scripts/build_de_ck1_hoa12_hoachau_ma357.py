# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Học Kỳ I - Hóa học 12 - Trường THPT Hóa Châu (TP Huế)
Mã đề: 357 - Năm học 2025 - 2026
Chuẩn theo Quy chuẩn Form Đề Kiểm Tra của Thầy TRẦN VĂN THẠNH
Tạo 2 phiên bản: Bản Học Sinh & Bản Giáo Viên (có BẢNG ĐÁP ÁN RÚT GỌN BAO GỒM CẢ TỰ LUẬN & lời giải chi tiết)
"""

import os
import sys
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Import module form_de_kiem_tra
sys.path.append(os.path.join(os.path.dirname(__file__)))
from form_de_kiem_tra import (
    setup_header_footer,
    build_title_block,
    add_section_header,
    render_question_prompt,
    render_options_no_table,
    render_statement,
    add_solution_no_table,
    add_student_writing_lines,
    render_rich_text,
    set_run_style,
    add_p_border_bottom,
    COLOR_PRIMARY,
    COLOR_CORRECT,
    COLOR_SOLUTION,
    COLOR_GRAY
)

DIR_SAN_PHAM = os.path.abspath("San_Pham")
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_de357")
os.makedirs(DIR_SAN_PHAM, exist_ok=True)
os.makedirs(DIR_IMAGES, exist_ok=True)

def set_cell_border(cell, **kwargs):
    """Cài đặt viền cho ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            val, sz, space, color = edge_data
            tag = f'<w:{edge} {nsdecls("w")} w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'
            tcBorders.append(parse_xml(tag))
        else:
            tag = f'<w:{edge} {nsdecls("w")} w:val="none"/>'
            tcBorders.append(parse_xml(tag))
    tcPr.append(tcBorders)

def set_cell_shading(cell, color_hex):
    """Tô màu nền cho ô bảng."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def add_compact_summary_answers(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN đầy đủ: Phần I, Phần II, Phần III và Phần IV Tự luận cho Bản Giáo Viên.
    """
    p_box = doc.add_paragraph()
    p_box.paragraph_format.space_before = Pt(8)
    p_box.paragraph_format.space_after = Pt(2)
    p_box.paragraph_format.keep_with_next = True
    r_box = p_box.add_run("BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - BAO GỒM CẢ TỰ LUẬN)")
    set_run_style(r_box, font_size=11, bold=True, color=COLOR_CORRECT)

    # 1. BẢNG PHẦN I
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(2)
    p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p1, "*1. PHẦN 1: CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_p1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
    answers_p1 = ["Đ/A", "B", "D", "C", "A", "A", "D", "D", "B", "B", "A", "C", "C"]

    w1_c0 = Cm(1.8)
    w1_c = Cm(1.35)

    for r_idx, r_data in enumerate([headers_p1, answers_p1]):
        row = tbl1.rows[r_idx]
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = w1_c0 if c_idx == 0 else w1_c
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=9.5, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                is_label = (c_idx == 0)
                set_cell_shading(cell, "F5F5F5" if is_label else "FFFFFF")
                c_color = COLOR_PRIMARY if is_label else COLOR_CORRECT
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # 2. BẢNG PHẦN II
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(5)
    p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p2, "*2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM - 1 ý: 0,1đ; 2 ý: 0,25đ; 3 ý: 0,5đ; 4 ý: 1,0đ)*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl2.autofit = False

    tbl2_data = [
        ["Câu hỏi", "Nội dung câu hỏi", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Tính chất của polymer nhiệt dẻo", "ĐÚNG", "ĐÚNG", "ĐÚNG", "ĐÚNG"],
        ["Câu 2", "Pin điện hóa Galvani Zn – Cu", "SAI", "SAI", "SAI", "ĐÚNG"]
    ]

    w2_widths = [Cm(1.8), Cm(6.2), Cm(2.5), Cm(2.5), Cm(2.5), Cm(2.5)]

    for r_idx, row_vals in enumerate(tbl2_data):
        row = tbl2.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = w2_widths[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx != 1 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=9.0, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                is_label = (c_idx < 2)
                set_cell_shading(cell, "F5F5F5" if is_label else "FFFFFF")
                if is_label:
                    c_color = COLOR_PRIMARY if c_idx == 0 else RGBColor(0, 0, 0)
                else:
                    c_color = COLOR_CORRECT if val == "ĐÚNG" else COLOR_PRIMARY
                render_rich_text(p, f"*{val}*", base_size=9.5, base_bold=True, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # 3. BẢNG PHẦN III & PHẦN IV
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(5)
    p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p3, "*3. PHẦN III: TRẢ LỜI NGẮN (2,0 ĐIỂM - 0,25 Đ/CÂU) & PHẦN IV: TỰ LUẬN (3,0 ĐIỂM)*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=10)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False

    headers_p3 = ["P.III-C1", "P.III-C2", "P.III-C3", "P.III-C4", "P.III-C5", "P.III-C6", "P.III-C7", "P.III-C8", "P.IV-C1", "P.IV-C2"]
    answers_p3 = ["3", "10", "4", "2", "-1,66 V", "2", "2", "4958 m³", "k = 2", "M: Ag; E°=1,1V"]

    w3_c = Cm(1.8)

    for r_idx, r_data in enumerate([headers_p3, answers_p3]):
        row = tbl3.rows[r_idx]
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = w3_c
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=8.5, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                set_cell_shading(cell, "FFFFFF")
                render_rich_text(p, f"*{val}*", base_size=8.5, base_bold=True, base_color=COLOR_CORRECT)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # Đường kẻ phân cách Sage Green
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(8)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


DATA_EXAM_357 = {
    "school_line": "SỞ GD & ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT HOÁ CHÂU",
    "exam_title": "ĐỀ KIỂM TRA CUỐI HỌC KÌ I\nNĂM HỌC 2025-2026",
    "subject_line": "MÔN: HOÁ HỌC – Lớp 12",
    "num_pages": "03",
    "duration": "45 phút (Không kể thời gian phát đề)",
    "exam_code": "357",
    "atomic_masses": "H = 1; C = 12; O = 16; N = 14; Cl = 35,5; Ag = 108; Cu = 64; Zn = 65",

    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1: Cho các phát biểu sau:\n"
                      "(1) Protein bị thủy phân khi đun nóng với dung dịch acid.\n"
                      "(2) Tripeptide có khả năng tham gia phản ứng màu biuret.\n"
                      "(3) Trong phân tử Gly-Ala-Gly có chứa 3 liên kết peptide.\n"
                      "(4) Hợp chất H~2~N-CH~2~-CO-NH-CH~2~-CH~2~-COOH là dipeptide.\n"
                      "Có bao nhiêu phát biểu đúng trong các phát biểu trên?",
            "options": [
                "A. 1.",
                "*B. 2.*",
                "C. 3.",
                "D. 4."
            ],
            "solution": "Phân tích từng phát biểu:\n"
                       "- (1) ĐÚNG: Liên kết peptide trong protein bị thủy phân khi đun nóng với dung dịch acid hoặc dung dịch kiềm.\n"
                       "- (2) ĐÚNG: Tripeptide chứa 2 liên kết peptide nên có phản ứng màu biuret với Cu(OH)~2~ trong môi trường kiềm tạo phức chất màu tím.\n"
                       "- (3) SAI: Tripeptide Gly-Ala-Gly chỉ chứa (3 - 1) = 2 liên kết peptide.\n"
                       "- (4) SAI: Peptide là hợp chất chứa từ 2 đến 50 gốc α-amino acid. Gốc amino acid thứ hai là H~2~N-CH~2~-CH~2~-COOH (β-alanine, nhóm -NH~2~ ở vị trí β) nên phân tử này không phải là dipeptide.\n"
                       "=> Có 2 phát biểu đúng là (1) và (2)."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Ester là đồng phân với HCOOC~2~H~5~ được điều chế từ các hợp chất tương ứng nào sau đây?",
            "options": [
                "A. Acetic acid và ethyl alcohol.",
                "B. Formic acid và methyl alcohol.",
                "C. Formic acid và ethyl alcohol.",
                "*D. Acetic acid và methyl alcohol.*"
            ],
            "solution": "HCOOC~2~H~5~ có công thức phân tử là C~3~H~6~O~2~. Ester đồng phân của nó là methyl acetate (CH~3~COOCH~3~). Methyl acetate được điều chế từ acetic acid (CH~3~COOH) và methyl alcohol (CH~3~OH)."
        },
        {
            "id": 3,
            "prompt": "Câu 3: Cho pin điện hoá Zn-Cu. Quá trình xảy ra ở cực dương của pin là",
            "options": [
                "A. Zn^2+^ + 2e -> Zn.",
                "B. Zn -> Zn^2+^ + 2e.",
                "*C. Cu^2+^ + 2e -> Cu.*",
                "D. Cu -> Cu^2+^ + 2e."
            ],
            "solution": "Trong pin Galvani Zn-Cu, thanh kẽm là cực âm (anode) xảy ra quá trình oxi hóa Zn -> Zn^2+^ + 2e; thanh đồng là cực dương (cathode) xảy ra quá trình khử Cu^2+^ + 2e -> Cu."
        },
        {
            "id": 4,
            "prompt": "Câu 4: Tên gọi của polyme có công thức cho dưới đây là:\n"
                      "[-N(H)-(CH~2~)~6~-N(H)-C(=O)-(CH~2~)~4~-C(=O)-]~n~",
            "image": "ctct_nylon66_cau4.png",
            "image_width": Cm(8.5),
            "options": [
                "*A. nilon-6,6.*",
                "B. olon.",
                "C. nilon-7.",
                "D. nilon-6."
            ],
            "solution": "Polymer có công thức [-NH-(CH~2~)~6~-NH-CO-(CH~2~)~4~-CO-]~n~ được điều chế bằng phản ứng trùng ngưng giữa hexamethylenediamine và adipic acid, có tên gọi là poly(hexamethylene adipamide) hay nilon-6,6 (nylon-6,6)."
        },
        {
            "id": 5,
            "prompt": "Câu 5: Cho các phát biểu sau:\n"
                      "(a) Tinh bột trong gạo tẻ chứa nhiều amylose, trong gạo nếp chứa nhiều amylopectin.\n"
                      "(b) Amylose có mạch không phân nhánh do giữa các đơn vị \u03b1-glucose chỉ có liên kết \u03b1-1,4-glycoside.\n"
                      "(c) Ngoài liên kết \u03b1-1,4-glycoside, còn xuất hiện liên kết \u03b1-1,6-glycoside giữa các đơn vị \u03b1-glucose là nguyên nhân làm cho amylopectin có mạch phân nhánh.\n"
                      "(d) Amylose và amylopectin đều cho được phản ứng thuỷ phân.\n"
                      "Số phát biểu đúng là",
            "options": [
                "*A. 4.*",
                "B. 3.",
                "C. 2.",
                "D. 1."
            ],
            "solution": "Cả 4 phát biểu (a), (b), (c), (d) đều đúng về cấu tạo và tính chất của tinh bột (amylose và amylopectin)."
        },
        {
            "id": 6,
            "prompt": "Câu 6: Phân tử glucose có số nguyên tử oxygen là",
            "options": [
                "A. 11.",
                "B. 5.",
                "C. 12.",
                "*D. 6.*"
            ],
            "solution": "Công thức phân tử của glucose là C~6~H~12~O~6~ nên phân tử có 6 nguyên tử oxygen."
        },
        {
            "id": 7,
            "prompt": "Câu 7: Điện phân với điện cực trơ, có màng ngăn giữa các điện cực các dung dịch sau: KCl, CuSO~4~, AgNO~3~, CuCl~2~, NaCl, NiSO~4~, ZnCl~2~. Số dung dịch sau điện phân có khả năng hoà tan Al~2~O~3~ là",
            "options": [
                "A. 4.",
                "B. 3.",
                "C. 2.",
                "*D. 5.*"
            ],
            "solution": "Al~2~O~3~ là oxide lưỡng tính, tan được trong dung dịch acid mạnh và dung dịch base kiềm mạnh. Sau điện phân:\n"
                       "- Dung dịch KCl sinh ra KOH (base mạnh) -> hòa tan được Al~2~O~3~.\n"
                       "- Dung dịch CuSO~4~ sinh ra H~2~SO~4~ (acid mạnh) -> hòa tan được Al~2~O~3~.\n"
                       "- Dung dịch AgNO~3~ sinh ra HNO~3~ (acid mạnh) -> hòa tan được Al~2~O~3~.\n"
                       "- Dung dịch CuCl~2~ điện phân tạo Cu và Cl~2~ -> không tạo acid/kiềm mạnh.\n"
                       "- Dung dịch NaCl sinh ra NaOH (base mạnh) -> hòa tan được Al~2~O~3~.\n"
                       "- Dung dịch NiSO~4~ sinh ra H~2~SO~4~ (acid mạnh) -> hòa tan được Al~2~O~3~.\n"
                       "- Dung dịch ZnCl~2~ điện phân tạo Zn và Cl~2~ -> không tạo acid/kiềm mạnh.\n"
                       "=> Có 5 dung dịch sau điện phân hòa tan được Al~2~O~3~ là KCl, CuSO~4~, AgNO~3~, NaCl, NiSO~4~."
        },
        {
            "id": 8,
            "prompt": "Câu 8: Cho các phát biểu sau:\n"
                      "(a) Stearic acid là chất béo no, oleic acid là chất béo không no.\n"
                      "(b) Muối sodium hoặc potassium của stearic acid đều là thành phần chính của xà phòng.\n"
                      "(c) Hydrogen hoá oleic acid thu được stearic acid.\n"
                      "(d) Stearic acid có nhiệt độ nóng chảy cao hơn oleic acid.\n"
                      "Số phát biểu đúng là",
            "options": [
                "A. 4.",
                "*B. 3.*",
                "C. 1.",
                "D. 2."
            ],
            "solution": "Phân tích:\n"
                       "- (a) SAI: Stearic acid (C~17~H~35~COOH) và oleic acid (C~17~H~33~COOH) là các ACID BÉO, không phải chất béo (chất béo là triester của glycerol).\n"
                       "- (b) ĐÚNG: Muối sodium stearate hoặc potassium stearate là thành phần chính của xà phòng.\n"
                       "- (c) ĐÚNG: C~17~H~33~COOH + H~2~ --[Ni, t°]--> C~17~H~35~COOH.\n"
                       "- (d) ĐÚNG: Stearic acid là acid béo no, ở điều kiện thường ở trạng thái rắn (t°nc ≈ 69°C), cao hơn oleic acid là acid béo không no ở trạng thái lỏng (t°nc ≈ 13°C).\n"
                       "=> Có 3 phát biểu đúng là (b), (c), (d)."
        },
        {
            "id": 9,
            "prompt": "Câu 9: Trên một số vật dụng polymer có các ký hiệu sau:\n"
                      "Các kí hiệu này giúp người sử dụng và thu gom vật liệu polymer biết được thông tin gì?",
            "image": "ky_hieu_tai_che_6.png",
            "image_width": Cm(10.0),
            "options": [
                "A. Vật liệu độc hại, cần tránh xa tầm tay trẻ em.",
                "*B. Vật liệu có thể tái chế được.*",
                "C. Vật liệu dễ cháy, nên bảo quản ở nhiệt độ thấp.",
                "D. Vật liệu có thể đốt và không gây ô nhiễm môi trường."
            ],
            "solution": "Ký hiệu hình tam giác có mũi tên xoay vòng kèm chữ số từ 1 đến 6 là mã nhận diện nhựa (Resin Identification Code) giúp người dùng và các cơ sở thu gom biết loại nhựa đó có thể tái chế được và thuộc nhóm tái chế nào."
        },
        {
            "id": 10,
            "prompt": "Câu 10: Thuỷ phân 10 gam saccharose thu được dung dịch X chứa 10,3 gam hỗn hợp glucose, fructose và saccharose còn dư. Trung hoà dung dịch X rồi thực hiện phản ứng với lượng dư thuốc thử Tollens thì thu được tối đa bao nhiêu gam bạc?",
            "options": [
                "*A. 7,20 gam.*",
                "B. 8,64 gam.",
                "C. 3,24 gam.",
                "D. 6,48 gam."
            ],
            "solution": "1. Khối lượng tăng lên sau phản ứng thủy phân là do nước tham gia phản ứng:\n"
                       "   m(H~2~O) = 10,3 - 10 = 0,3 gam => n(H~2~O) = 0,3 / 18 = 1/60 mol.\n"
                       "2. Phương trình thủy phân: C~12~H~22~O~11~ + H~2~O -> C~6~H~12~O~6~ (glucose) + C~6~H~12~O~6~ (fructose)\n"
                       "   => n(glucose) = n(fructose) = n(H~2~O) = 1/60 mol.\n"
                       "3. Phản ứng tráng bạc với thuốc thử Tollens dư (trong môi trường kiềm NH~3~, cả glucose và fructose đều tráng bạc):\n"
                       "   1 glucose -> 2 Ag;  1 fructose -> 2 Ag\n"
                       "   => Tổng n(Ag) = 2 * (1/60) + 2 * (1/60) = 4/60 = 1/15 mol.\n"
                       "4. Khối lượng Ag thu được tối đa: m(Ag) = (1/15) * 108 = 7,20 gam."
        },
        {
            "id": 11,
            "prompt": "Câu 11: Phát biểu nào sau đây không đúng?",
            "options": [
                "A. Protein tác dụng với dung dịch nitric acid đặc tạo thành sản phẩm rắn có màu vàng.",
                "B. Thuỷ phân hoàn toàn polypeptide thu được các phân tử α-amino acid.",
                "*C. Protein tác dụng với Cu(OH)~2~ trong môi trường kiềm tạo dung dịch màu xanh lam.*",
                "D. Protein có thể bị đông tụ dưới tác dụng của nhiệt, acid hoặc base."
            ],
            "solution": "Phát biểu C không đúng: Protein chứa nhiều liên kết peptide tác dụng với Cu(OH)~2~ trong môi trường kiềm tạo hợp chất phức có màu TÍM đặc trưng (phản ứng màu biuret), không phải màu xanh lam."
        },
        {
            "id": 12,
            "prompt": "Câu 12: Trong các đồng phân cấu tạo của các amine có công thức C~3~H~9~N, số amine bậc một là",
            "options": [
                "A. 0.",
                "B. 3.",
                "*C. 2.*",
                "D. 1."
            ],
            "solution": "Các đồng phân amine bậc I có dạng R-NH~2~ ứng với C~3~H~9~N gồm 2 đồng phân:\n(1) CH~3~-CH~2~-CH~2~-NH~2~ (propan-1-amine)\n(2) CH~3~-CH(NH~2~)-CH~3~ (propan-2-amine)."
        }
    ],

    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1: Polymer nhiệt dẻo khi bị đun nóng đến nóng chảy thì trở nên mềm, dễ ăn khuôn và khi nguội thì đóng rắn lại.",
            "statements": [
                ("a) Vật liệu polymer nhiệt dẻo có thể đun nóng và tạo hình nhiều lần.", "ĐÚNG"),
                ("b) Vật liệu polymer nhiệt dẻo thích hợp cho việc tái chế.", "ĐÚNG"),
                ("c) PVC và PS là polymer nhiệt dẻo có nhiều ứng dụng.", "ĐÚNG"),
                ("d) Polymer nhiệt dẻo thường có các kí hiệu in trên bao bì, vỏ hộp, đồ dùng, ... để giúp nhận biết vật liệu polymer cũng như thuận lợi cho việc thu gom, phân loại, tái chế.", "ĐÚNG")
            ],
            "solution": "- a) ĐÚNG: Đặc tính cơ bản của polymer nhiệt dẻo là có thể nung nóng chảy rồi làm nguội đóng rắn để tạo hình nhiều lần mà không bị biến tính mạch polymer.\n"
                       "- b) ĐÚNG: Nhờ khả năng nấu chảy lại nên polymer nhiệt dẻo rất thuận tiện cho việc tái chế.\n"
                       "- c) ĐÚNG: Poly(vinyl chloride) (PVC) và polystyrene (PS) là các polymer nhiệt dẻo phổ biến trong đời sống.\n"
                       "- d) ĐÚNG: Các ký hiệu mã nhận diện nhựa giúp phân loại chính xác các loại nhựa nhiệt dẻo để tái chế."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Một pin điện hoá có điện cực kẽm nhúng trong dung dịch ZnSO~4~ và điện cực đồng nhúng trong dung dịch CuSO~4~. Em hãy cho biết các phát biểu sau đúng hay sai?",
            "statements": [
                ("a) Điện cực kẽm tăng còn khối lượng điện cực đồng giảm.", "SAI"),
                ("b) Cho E~pin(Zn - Cu)~ = 1,1 V và E°(Cu^2+^/Cu) = 0,34 V. Thế điện cực chuẩn của cặp oxi hoá - khử E°(Zn^2+^/Zn) là 0,76V.", "SAI"),
                ("c) Điện cực Zn xảy ra sự oxi hoá ion Zn^2+^.", "SAI"),
                ("d) Điện cực kẽm giảm còn khối lượng điện cực đồng tăng.", "ĐÚNG")
            ],
            "solution": "- a) SAI: Kẽm ở anode bị oxi hóa tan ra nên giảm khối lượng; Cu ở cathode được sinh ra bám vào nên tăng khối lượng.\n"
                       "- b) SAI: E°~pin~ = E°(Cu^2+^/Cu) - E°(Zn^2+^/Zn) => 1,1 = 0,34 - E°(Zn^2+^/Zn) => E°(Zn^2+^/Zn) = 0,34 - 1,1 = -0,76 V (không phải +0,76V).\n"
                       "- c) SAI: Tại điện cực Zn xảy ra quá trình oxi hóa kim loại Zn: Zn -> Zn^2+^ + 2e (không phải oxi hóa ion Zn^2+^).\n"
                       "- d) ĐÚNG: Khối lượng điện cực Zn giảm và khối lượng điện cực Cu tăng."
        }
    ],

    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1: Số nhóm -OH trong một mắt xích để tạo phân tử cellulose là bao nhiêu?",
            "answer": "3",
            "solution": "Mỗi mắt xích β-glucose trong phân tử cellulose có 3 nhóm hydroxy (-OH) tự do, công thức cấu tạo của một mắt xích là [C~6~H~7~O~2~(OH)~3~]."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Nicotine làm tăng huyết áp và nhịp tim, có khả năng gây xơ vữa động mạch vành và suy giảm trí nhớ. Trong cây thuốc lá và khói thuốc lá có chứa một amine rất độc, đó là nicotine với công thức cấu tạo như hình bên. Số nguyên tử carbon trong phân tử nicotine là bao nhiêu?",
            "image": "nicotine.png",
            "image_width": Cm(4.0),
            "answer": "10",
            "solution": "Phân tử nicotine gồm 1 vòng pyridine (5 nguyên tử C), 1 vòng pyrrolidine (4 nguyên tử C) và 1 nhóm methyl -CH~3~ đính trên nguyên tử nitrogen (1 nguyên tử C). Tổng số nguyên tử carbon là: 5 + 4 + 1 = 10 nguyên tử C (CTPT của nicotine là C~10~H~14~N~2~)."
        },
        {
            "id": 3,
            "prompt": "Câu 3: Có bao nhiêu phát biểu đúng khi nhận xét về acid béo?\n"
                      "(1). Là monocarboxylic acid, thường có số nguyên tử carbon chẵn và mạch carbon không phân nhánh.\n"
                      "(2). Có thể no hoặc không no.\n"
                      "(3). Tính chất vật lí phụ thuộc vào độ dài của mạch carbon và mức độ bão hoà.\n"
                      "(4). Muối của chúng được dùng làm xà phòng.",
            "answer": "4",
            "solution": "Cả 4 phát biểu trên đều hoàn toàn chính xác khi nhận xét về đặc điểm cấu tạo và tính chất của acid béo."
        },
        {
            "id": 4,
            "prompt": "Câu 4: Cho các phát biểu sau:\n"
                      "(a) Saccharose chỉ có cấu tạo dạng mạch vòng.\n"
                      "(b) Dung dịch saccharose hoà tan Cu(OH)~2~ tạo dung dịch xanh lam.\n"
                      "(c) Saccharose bị oxi hoá bởi thuốc thử Tollens.\n"
                      "(d) Saccharose có khả năng làm mất màu nước bromine.\n"
                      "(e) Saccharose cho được phản ứng thuỷ phân.\n"
                      "Có bao nhiêu phát biểu sai?",
            "answer": "2",
            "solution": "Các phát biểu sai là (c) và (d):\n"
                       "- (c) SAI vì saccharose không có nhóm -CHO hoặc -OH hemiaxetal tự do nên không phản ứng với thuốc thử Tollens.\n"
                       "- (d) SAI vì saccharose không làm mất màu nước bromine.\n"
                       "Các phát biểu đúng là (a), (b), (e).\n"
                       "=> Số phát biểu sai là 2."
        },
        {
            "id": 5,
            "prompt": "Câu 5: Cho pin điện hoá Al - Cu có E°~pin(Al-Cu)~ = 2,00 V; và E°(Cu^2+^/Cu) = 0,34V. Thế điện cực chuẩn của cặp oxi hoá khử Al^3+^/Al có giá trị là bao nhiêu volt? (Không làm tròn kết quả)",
            "answer": "-1,66 V",
            "solution": "Ta có công thức: E°~pin(Al-Cu)~ = E°(Cu^2+^/Cu) - E°(Al^3+^/Al)\n"
                       "=> 2,00 = 0,34 - E°(Al^3+^/Al)\n"
                       "=> E°(Al^3+^/Al) = 0,34 - 2,00 = -1,66 V."
        },
        {
            "id": 6,
            "prompt": "Câu 6: Có bao nhiêu amino acid ứng với công thức phân tử C~3~H~7~NO~2~?",
            "answer": "2",
            "solution": "Các amino acid ứng với CTPT C~3~H~7~NO~2~ (chứa 1 nhóm -NH~2~ và 1 nhóm -COOH) gồm 2 đồng phân:\n"
                       "(1) CH~3~-CH(NH~2~)-COOH (alanine hay α-aminopropionic acid)\n"
                       "(2) H~2~N-CH~2~-CH~2~-COOH (β-alanine hay β-aminopropionic acid)."
        },
        {
            "id": 7,
            "prompt": "Câu 7: Kevlar là một polymer có độ bền kéo rất cao. Loại vật liệu này được dùng để sản xuất áo chống đạn và mũ bảo hiểm cho quân đội. Kevlar được điều chế bằng phản ứng của hai chất: terephthalic acid và 1,4-diaminebenzene.\n"
                      "Cho các phát biểu sau:\n"
                      "(1) Kevlar thuộc loại polyamide.\n"
                      "(2) Phản ứng tổng hợp Kevlar thuộc phản ứng trùng ngưng.\n"
                      "(3) Mỗi phân tử terephthalic acid và 1,2-diaminebenzene đều có chứa 3 liên kết pi.\n"
                      "(4) Monomer để tổng hợp Kevlar là terephthalic acid và 1,2-diaminebenzene.\n"
                      "(5) Cả hai chất terephthalic acid và 1,2-diaminebenzene đều là hợp chất tạp chức.\n"
                      "Có bao nhiêu phát biểu đúng?",
            "image": "kevlar_scheme.png",
            "image_width": Cm(10.0),
            "answer": "2",
            "solution": "Phân tích:\n"
                       "- (1) ĐÚNG: Kevlar chứa liên kết amide -CO-NH- nên thuộc loại polyamide.\n"
                       "- (2) ĐÚNG: Phản ứng đồng trùng ngưng giải phóng phân tử H~2~O.\n"
                       "- (3) SAI: Terephthalic acid chứa 5 liên kết pi (3 trong vòng benzen và 2 trong hai nhóm C=O).\n"
                       "- (4) SAI: Monomer là 1,4-diaminebenzene (đồng phân para), không phải 1,2-diaminebenzene (ortho).\n"
                       "- (5) SAI: Cả hai chất đều là hợp chất đa chức (chứa 2 nhóm -COOH hoặc 2 nhóm -NH~2~ giống nhau).\n"
                       "=> Có 2 phát biểu đúng là (1) và (2)."
        },
        {
            "id": 8,
            "prompt": "Câu 8: Da nhân tạo (PVC) được điều chế từ khí thiên nhiên (CH~4~). Nếu hiệu suất của toàn bộ quá trình là 20% thì để điều chế 1 tấn PVC phải cần một thể tích khí thiên nhiên ở điều kiện chuẩn (xem khí thiên nhiên chứa 80% methane) là bao nhiêu m^3^?",
            "answer": "4958 m³",
            "solution": "Sơ đồ bảo toàn nguyên tố C: 2 CH~4~ -> C~2~H~3~Cl -> PVC\n"
                       "1. Lượng PVC cần điều chế: m = 1 tấn = 10^6^ gam => n(mắt xích PVC) = 10^6^ / 62,5 = 16000 mol.\n"
                       "2. Số mol CH~4~ lý thuyết: n(CH~4~) = 2 * 16000 = 32000 mol.\n"
                       "3. Thể tích CH~4~ ở đkc (24,79 L/mol): V_LT(CH~4~) = 32000 * 24,79 * 10^-3^ = 793,28 m^3^.\n"
                       "4. Do hiệu suất là 20% nên thể tích CH~4~ thực tế: V_TT(CH~4~) = 793,28 / 20% = 3966,4 m^3^.\n"
                       "5. Thể tích khí thiên nhiên cần dùng (chứa 80% methane):\n"
                       "   V(khí thiên nhiên) = 3966,4 / 80% = 4958 m^3^.\n"
                       "*(Ghi chú: Nếu tính theo đktc cũ 22,4 L/mol thì V = (32000 * 22,4 * 10^-3^) / (0,2 * 0,8) = 4480 m^3^).*"
        }
    ],

    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1: (1,5 điểm).\n"
                      "a) Viết phương trình điều chế các polymer sau: Cao su Buna, PVC, nylon – 6,6.\n"
                      "b) Chloro hoá PVC thu được một polymer chứa 66,77% chlorine về khối lượng, trung bình 1 phân tử chlorine phản ứng với k mắt xích trong mạch PVC. Hãy xác định k? (cho H = 1, C = 12, Cl = 35,5).",
            "num_lines": 8,
            "solution": "a) Viết phương trình phản ứng hóa học điều chế các polymer:\n"
                       "1. Cao su Buna (trùng hợp buta-1,3-diene):\n"
                       "   n CH~2~=CH-CH=CH~2~ --[t°, Na, p]--> [-CH~2~-CH=CH-CH~2~-]~n~\n"
                       "2. Poly(vinyl chloride) (trùng hợp vinyl chloride):\n"
                       "   n CH~2~=CH-Cl --[t°, xt, p]--> [-CH~2~-CH(Cl)-]~n~\n"
                       "3. Tơ nylon-6,6 (đồng trùng ngưng hexamethylenediamine và adipic acid):\n"
                       "   n H~2~N-(CH~2~)~6~-NH~2~ + n HOOC-(CH~2~)~4~-COOH --[t°, p]--> [-NH-(CH~2~)~6~-NH-CO-(CH~2~)~4~-CO-]~n~ + 2n H~2~O\n"
                       "b) Xác định hệ số k trong phản ứng chloro hóa PVC:\n"
                       "- Công thức đoạn mạch chứa k mắt xích PVC là (C~2~H~3~Cl)~k~ hay C~2k~H~3k~Cl~k~.\n"
                       "- Phản ứng thế 1 phân tử Cl~2~ vào k mắt xích PVC:\n"
                       "  C~2k~H~3k~Cl~k~ + Cl~2~ -> C~2k~H~3k-1~Cl~k+1~ + HCl\n"
                       "- Khối lượng phân tử đoạn polymer clo hóa:\n"
                       "  M = 12 * 2k + 1 * (3k - 1) + 35,5 * (k + 1) = 62,5k + 34,5\n"
                       "- Phần trăm khối lượng của chlorine trong polymer:\n"
                       "  %Cl = [35,5 * (k + 1)] / (62,5k + 34,5) * 100% = 66,77%\n"
                       "  => 35,5k + 35,5 = 0,6677 * (62,5k + 34,5)\n"
                       "  => 35,5k + 35,5 = 41,73125k + 23,03565\n"
                       "  => 6,23125k = 12,46435  =>  k = 2.\n"
                       "Kết luận: k = 2 (trung bình cứ 2 mắt xích PVC thì thế 1 nguyên tử chlorine)."
        },
        {
            "id": 2,
            "prompt": "Câu 2: (1,5 điểm).\n"
                      "a) Pin Galvani Zn - Cu gồm điện cực Zn nhúng trong dung dịch ZnSO~4~ và điện cực Cu nhúng trong dung dịch CuSO~4~ được nối với nhau bởi cầu muối (thường chứa dung dịch KCl bão hoà) như hình sau:\n"
                      "Em hãy:\n"
                      "- Nêu nguyên tắc hoạt động của pin Galvani Zn-Cu.\n"
                      "- Sau quá trình hoạt động, khối lượng của hai điện cực thay đổi như thế nào? Giải thích.\n"
                      "- Tính sức điện động chuẩn của pin Zn-Cu, biết E°(Cu^2+^/Cu) = 0,34V; E°(Zn^2+^/Zn) = -0,76V.\n"
                      "b) Điện phân dung dịch M(NO~3~)~n~ (điện cực trơ, cường độ dòng điện không đổi), ở cathode chỉ thu được 5,4 gam kim loại M và ở anode thu được 0,31 lít khí (đkc). Viết các quá trình xảy ra ở các điện cực và phương trình điện phân; Xác định kim loại M; Ở điện cực nào có pH thay đổi? Giải thích?",
            "image": "pin_galvani_de357.png",
            "image_width": Cm(8.0),
            "num_lines": 9,
            "solution": "a) Pin Galvani Zn - Cu:\n"
                       "1. Nguyên tắc hoạt động:\n"
                       "   Dựa trên phản ứng oxi hóa - khử tự diễn biến giữa kim loại kẽm và ion đồng: Zn + Cu^2+^ -> Zn^2+^ + Cu. Hóa năng chuyển thành điện năng. Các electron di chuyển từ thanh Zn (cực âm anode) qua dây dẫn ngoài đến thanh Cu (cực dương cathode). Cầu muối cho các ion khuếch tán để trung hòa điện tích hai dung dịch.\n"
                       "2. Sự thay đổi khối lượng hai điện cực:\n"
                       "   - Tại anode (cực âm Zn): Zn bị oxi hóa tan thành ion Zn^2+^ (Zn -> Zn^2+^ + 2e) => Khối lượng thanh Zn giảm.\n"
                       "   - Tại cathode (cực dương Cu): Cu^2+^ bị khử thành kim loại Cu bám vào điện cực (Cu^2+^ + 2e -> Cu) => Khối lượng thanh Cu tăng.\n"
                       "3. Sức điện động chuẩn của pin Zn-Cu:\n"
                       "   E°~pin~ = E°(cathode) - E°(anode) = E°(Cu^2+^/Cu) - E°(Zn^2+^/Zn) = 0,34 - (-0,76) = 1,10 V.\n"
                       "\n"
                       "b) Bài toán điện phân dung dịch M(NO~3~)~n~:\n"
                       "1. Khí thu được ở anode (cực dương) là O~2~ do nước bị oxi hóa:\n"
                       "   2H~2~O -> O~2~ + 4H^+^ + 4e\n"
                       "   n(O~2~) = 0,31 / 24,79 ≈ 0,0125 mol => n_e trao đổi = 4 * 0,0125 = 0,05 mol.\n"
                       "2. Quá trình xảy ra ở cathode (cực âm):\n"
                       "   M^n+^ + ne -> M\n"
                       "   n(M) = n_e / n = 0,05 / n mol.\n"
                       "3. Xác định kim loại M:\n"
                       "   M_M = m(M) / n(M) = 5,4 / (0,05 / n) = 108n.\n"
                       "   - Với n = 1 => M_M = 108 (g/mol): Kim loại Bạc (Ag).\n"
                       "   - Với n = 2, 3 => Không có kim loại phù hợp.\n"
                       "   Vậy kim loại M là Bạc (Ag), muối là AgNO~3~.\n"
                       "4. Phương trình điện phân dung dịch:\n"
                       "   4AgNO~3~ + 2H~2~O --[đpdd]--> 4Ag + O~2~↑ + 4HNO~3~\n"
                       "5. Điện cực có pH thay đổi:\n"
                       "   Ở ANODE (cực dương) sinh ra ion H^+^ (2H~2~O -> O~2~ + 4H^+^ + 4e) làm nồng độ H^+^ tăng lên, do đó pH ở khu vực điện cực này GIẢM (pH < 7, môi trường acid)."
        }
    ]
}


def generate_docx(is_teacher=False):
    """Tạo file docx theo chuẩn Thầy TRẦN VĂN THẠNH."""
    doc = Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line=DATA_EXAM_357["school_line"],
        exam_title=DATA_EXAM_357["exam_title"],
        subject_line=DATA_EXAM_357["subject_line"],
        num_pages=DATA_EXAM_357["num_pages"],
        duration=DATA_EXAM_357["duration"],
        exam_code=DATA_EXAM_357["exam_code"],
        atomic_masses=DATA_EXAM_357["atomic_masses"],
        is_teacher=is_teacher
    )

    if is_teacher:
        add_compact_summary_answers(doc)

    # PHẦN I
    add_section_header(
        doc,
        title="A. TRẮC NGHIỆM KHÁCH QUAN",
        subtitle="PHẦN 1: Câu trắc nghiệm nhiều phương án lựa chọn (3 điểm - 12 câu). Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    for item in DATA_EXAM_357["part1"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        if "image" in item:
            img_path = os.path.join(DIR_IMAGES, item["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(3)
                p_img.add_run().add_picture(img_path, width=item.get("image_width", Cm(8.0)))

        render_options_no_table(doc, item["options"], is_teacher=is_teacher)
        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN II
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai: 2 điểm",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    for item in DATA_EXAM_357["part2"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        for stmt, ans in item["statements"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt, ans=ans, is_teacher=is_teacher, base_size=10.5)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN III
    add_section_header(
        doc,
        title="PHẦN III. (8 câu: 2 điểm) Câu hỏi trắc nghiệm yêu cầu trả lời ngắn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 8."
    )

    for item in DATA_EXAM_357["part3"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        if "image" in item:
            img_path = os.path.join(DIR_IMAGES, item["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(3)
                p_img.add_run().add_picture(img_path, width=item.get("image_width", Cm(8.0)))

        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(2)
        if is_teacher:
            r_ans_lbl = p_ans.add_run("Trả lời: ")
            set_run_style(r_ans_lbl, font_size=10.5, bold=True)
            r_ans_val = p_ans.add_run(item["answer"])
            set_run_style(r_ans_val, font_size=10.5, bold=True, color=COLOR_CORRECT)
        else:
            r_ans = p_ans.add_run("Trả lời: ............................................................ ")
            set_run_style(r_ans, font_size=10.5, italic=True)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN IV
    add_section_header(
        doc,
        title="B. PHẦN IV - TỰ LUẬN: 3 điểm",
        subtitle="Thí sinh trình bày chi tiết lời giải từ câu 1 đến câu 2 vào giấy làm bài thi."
    )

    for item in DATA_EXAM_357["part4"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        if "image" in item:
            img_path = os.path.join(DIR_IMAGES, item["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(3)
                p_img.add_run().add_picture(img_path, width=item.get("image_width", Cm(8.0)))

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])
        elif not is_teacher:
            num_lines = item.get("num_lines", 8)
            add_student_writing_lines(doc, num_lines=num_lines)

    # Kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(0)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(r_end, font_size=11, bold=True)

    suffix = "_GiaoVien.docx" if is_teacher else "_HocSinh.docx"
    filename = f"De_Kiem_Tra_CK1_Hoa_12_HoaChau_Ma357{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
