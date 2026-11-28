# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Kỳ I - Hóa học 12 - Trường THPT Đặng Huy Trứ (TP Huế)
Mã đề: 104 - Năm học 2025 - 2026
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
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_de104")
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
    render_rich_text(p1, "*1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (4,0 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=17)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_p1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"]
    answers_p1 = ["Đ/A", "D", "D", "B", "D", "B", "B", "C", "A", "C", "C", "B", "A", "A", "B", "C", "C"]

    w1_c0 = Cm(1.2)
    w1_c = Cm(1.05)

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
                render_rich_text(p, f"*{val}*", base_size=9.0, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                is_label = (c_idx == 0)
                set_cell_shading(cell, "F5F5F5" if is_label else "FFFFFF")
                c_color = COLOR_PRIMARY if is_label else COLOR_CORRECT
                render_rich_text(p, f"*{val}*", base_size=9.5, base_bold=True, base_color=c_color)

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
        ["Câu 1", "Polymer Poly(ethylene terephthalate) (PET)", "ĐÚNG", "SAI", "ĐÚNG", "ĐÚNG"],
        ["Câu 2", "Phản ứng thanh đồng Cu trong dung dịch AgNO~3~", "SAI", "SAI", "ĐÚNG", "SAI"]
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
    render_rich_text(p3, "*3. PHẦN III: TRẢ LỜI NGẮN (1,0 ĐIỂM) & PHẦN IV: TỰ LUẬN (3,0 ĐIỂM)*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=7)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False

    headers_p34 = ["P.III - C1", "P.III - C2", "P.III - C3", "P.III - C4", "P.IV - C1", "P.IV - C2", "P.IV - C3"]
    answers_p34 = ["5", "3", "4", "3", "m = 193,2 g", "4 PTHH đúng", "E°(Zn-Sn) = 0,62 V"]

    w3_c = Cm(2.57)

    for r_idx, r_data in enumerate([headers_p34, answers_p34]):
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
                render_rich_text(p, f"*{val}*", base_size=9.0, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                set_cell_shading(cell, "FFFFFF")
                render_rich_text(p, f"*{val}*", base_size=9.0, base_bold=True, base_color=COLOR_CORRECT)

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


DATA_EXAM_104 = {
    "school_line": "SỞ GD - ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT ĐẶNG HUY TRỨ",
    "exam_title": "KIỂM TRA CUỐI KỲ I – NĂM HỌC 2025-2026",
    "subject_line": "MÔN: HÓA HỌC – Lớp 12",
    "num_pages": "03",
    "duration": "45 phút (không kể thời gian phát đề)",
    "exam_code": "104",
    "atomic_masses": "Cu: 64; Ag: 108; C: 12; H: 1; O: 16; N: 14; K: 39; Na: 23; Zn: 65; Ni: 59; Sn: 119; Mg: 24",

    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1: Cho các thế điện cực chuẩn của các cặp oxi hóa – khử của kim loại như sau:\n(1) Ag^+^/Ag (= 0,80V); (2) Cu^2+^/Cu (= 0,34V); (3) Na^+^/Na (-2,71V); (4) Mg^2+^/Mg (-2,37V)\nIon kim loại có tính oxi hoá mạnh nhất là ion kim loại trong cặp oxi hóa khử",
            "options": [
                "A. (2)",
                "B. (4)",
                "C. (3)",
                "*D. (1)*"
            ],
            "solution": "Căn cứ vào dãy thế điện cực chuẩn: Cặp oxi hóa - khử nào có giá trị E° càng lớn thì dạng oxi hóa càng có tính oxi hóa mạnh. Ở đây, cặp (1) Ag^+^/Ag có E° = +0,80V lớn nhất, do đó ion Ag^+^ có tính oxi hóa mạnh nhất."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Chất nào dưới đây không phải là amino acid?",
            "options": [
                "A. Glycine.",
                "B. Lysine.",
                "C. Glutamic acid",
                "*D. Aniline.*"
            ],
            "solution": "Aniline (C~6~H~5~NH~2~) là một amine thơm đơn chức bậc I, không chứa nhóm carboxyl (-COOH) nên không phải là amino acid. Glycine, lysine và glutamic acid đều là các amino acid."
        },
        {
            "id": 3,
            "prompt": "Câu 3: Chất nào sau đây có thể tham gia phản ứng trùng ngưng?",
            "options": [
                "A. CH~2~=CHC~6~H~5~.",
                "*B. H~2~N[CH~2~]~5~COOH.*",
                "C. CH~2~=CHCOOH.",
                "D. CH~2~=CH(CH~3~)."
            ],
            "solution": "H~2~N[CH~2~]~5~COOH (6-aminohexanoic acid hay ε-aminocaproic acid) chứa đồng thời nhóm amino (-NH~2~) và nhóm carboxyl (-COOH) nên có thể tham gia phản ứng trùng ngưng tạo tơ capron (nylon-6). Các chất còn lại chỉ có liên kết đôi C=C nên tham gia phản ứng trùng hợp."
        },
        {
            "id": 4,
            "prompt": "Câu 4: Một pin điện hoá có điện cực là Zn nhúng trong dung dịch ZnSO~4~ và điện cực là Cu nhúng trong dung dịch CuSO~4~. Sau một thời gian pin đó phóng điện thì",
            "options": [
                "A. khối lượng cả hai điện cực Zn và Cu đều giảm.",
                "B. khối lượng cả hai điện cực Zn và Cu đều tăng.",
                "C. khối lượng điện cực Zn tăng còn khối lượng điện cực Cu giảm.",
                "*D. khối lượng điện cực Zn giảm còn khối lượng điện cực Cu tăng.*"
            ],
            "solution": "Khi pin Galvani Zn-Cu hoạt động:\n- Ở cực âm (anode Zn): Zn bị oxi hóa tan vào dung dịch: Zn -> Zn^2+^ + 2e => Khối lượng thanh Zn giảm.\n- Ở cực dương (cathode Cu): Cu^2+^ bị khử bám vào điện cực: Cu^2+^ + 2e -> Cu => Khối lượng thanh Cu tăng."
        },
        {
            "id": 5,
            "prompt": "Câu 5: Chất nào dưới đây không tan trong nước nhưng tan được trong dung dịch Schweizer?",
            "options": [
                "A. Saccharose.",
                "*B. Cellulose.*",
                "C. Maltose.",
                "D. Fructose."
            ],
            "solution": "Cellulose có cấu trúc phân tử kéo dài, nhiều liên kết hydrogen nội phân tử và liên phân tử chặt chẽ nên không tan trong nước và các dung môi hữu cơ thông thường, nhưng tan được trong dung dịch Schweizer (dung dịch [Cu(NH~3~)~4~](OH)~2~)."
        },
        {
            "id": 6,
            "prompt": "Câu 6: Cho tetrapeptide X: Gly-Ala-Glu-Val. Amino acid đầu N là",
            "options": [
                "A. Val.",
                "*B. Gly.*",
                "C. Glu.",
                "D. Ala."
            ],
            "solution": "Theo quy ước quốc tế về biểu diễn chuỗi peptide, amino acid đầu N (còn nhóm -NH~2~ tự do) được viết ở vị trí đầu tiên bên trái, đó là Gly. Amino acid đầu C (còn nhóm -COOH tự do) được viết ở cuối cùng bên phải (Val)."
        },
        {
            "id": 7,
            "prompt": "Câu 7: Hình dưới đây là ký hiệu của 6 polymer nhiệt dẻo phổ biến có thể tái chế:\n"
                      "Các ký hiệu này thường được in trên bao bì, vỏ hộp, đồ dùng,... để giúp nhận biết vật liệu polymer cũng như thuận lợi cho việc thu gom, tái chế. Polymer có ký hiệu số 5 được điều chế bằng phản ứng trùng hợp monomer nào dưới đây?",
            "image": "ky_hieu_tai_che_6.png",
            "image_width": Cm(9.0),
            "options": [
                "A. CH~2~=CH~2~.",
                "B. CH~2~=CH-Cl.",
                "*C. CH~2~=CH-CH~3~.*",
                "D. CH~2~=CH-C~6~H~5~."
            ],
            "solution": "Ký hiệu số 5 là PP (polypropylene), được điều chế bằng phản ứng trùng hợp propylene (propene, CH~2~=CH-CH~3~). (Số 1 là PET, số 2 là HDPE, số 3 là PVC, số 4 là LDPE, số 6 là PS)."
        },
        {
            "id": 8,
            "prompt": "Câu 8: Cho các hợp chất có công thức cấu tạo dưới đây:\n"
                      "Những hợp chất nào trong số các chất trên thuộc loại \u03b1 – amino acid?",
            "image": "ctct_amino_acid_cau8.png",
            "image_width": Cm(8.0),
            "options": [
                "*A. Chất (1), chất (3).* ",
                "B. Chất (1) và chất (2).",
                "C. Chất (1), chất (2) và chất (3).",
                "D. Chất (2) và chất (3)."
            ],
            "solution": "\u03b1-amino acid là amino acid có nhóm -NH~2~ đính trực tiếp vào nguyên tử carbon số 2 (carbon \u03b1 kề nhóm -COOH):\n- Chất (1): CH~3~-CH~2~-CH(NH~2~)-COOH (\u03b1-aminobutyric acid, nhóm -NH~2~ ở vị trí \u03b1).\n- Chất (2): H~2~N-CH~2~-CH~2~-COOH (\u03b2-alanine, nhóm -NH~2~ ở vị trí \u03b2).\n- Chất (3): HOOC-CH(NH~2~)-CH~3~ (alanine, nhóm -NH~2~ ở vị trí \u03b1).\nVậy chất (1) và chất (3) thuộc loại \u03b1-amino acid."
        },
        {
            "id": 9,
            "prompt": "Câu 9: Khẳng định nào sau đây đúng?",
            "options": [
                "A. Chất béo là triester của glycerol với acid vô cơ.",
                "B. Chất béo là chất rắn tan tốt trong nước.",
                "*C. Chất béo không tan trong nước nhưng tan nhiều trong dung môi hữu cơ.*",
                "D. Dầu ăn và mỡ bôi trơn có cùng thành phần nguyên tố."
            ],
            "solution": "Chất béo nhẹ hơn nước, không tan trong nước nhưng tan nhiều trong các dung môi hữu cơ không phân cực như benzene, ether, chloroform, xăng,... A sai vì chất béo là triester của glycerol với các acid béo hữu cơ; B sai vì chất béo không tan trong nước; D sai vì dầu ăn chứa C, H, O còn mỡ bôi trơn là hydrocarbon (chỉ chứa C, H)."
        },
        {
            "id": 10,
            "prompt": "Câu 10: Xà phòng có thành phần chính là",
            "options": [
                "A. muối sodium hoặc potassium của acid bất kì.",
                "B. glycerol và sodium chloride hoặc potassium chloride.",
                "*C. muối sodium hoặc potassium của acid béo.*",
                "D. muối sodium hoặc potassium của carboxylic acid."
            ],
            "solution": "Xà phòng thông thường là hỗn hợp muối sodium hoặc potassium của các acid béo cao (như sodium stearate C~17~H~35~COONa, sodium palmitate C~15~H~31~COONa, sodium oleate C~17~H~33~COONa,...)."
        },
        {
            "id": 11,
            "prompt": "Câu 11: Nhận định nào sau đây đúng khi nói về amino acid?",
            "options": [
                "A. dễ nóng chảy và ít tan trong nước.",
                "*B. có nhiệt độ nóng chảy cao và tan tốt trong nước.*",
                "C. có nhiệt độ nóng chảy cao và ít tan trong nước.",
                "D. dễ nóng chảy và tan tốt trong nước."
            ],
            "solution": "Ở trạng thái tinh thể, phân tử amino acid tồn tại dưới dạng ion lưỡng cực H~3~N^+^-CHR-COO^-^. Lực hút tĩnh điện giữa các ion rất mạnh nên amino acid là chất rắn có nhiệt độ nóng chảy cao (thường bị phân hủy khi nóng chảy) và tan tốt trong nước."
        },
        {
            "id": 12,
            "prompt": "Câu 12: Phát biểu nào sau đây đúng?",
            "options": [
                "*A. PVC được điều chế bằng phản ứng trùng hợp.*",
                "B. Tơ visco thuộc loại tơ tổng hợp.",
                "C. Amylose có cấu trúc mạch phân nhánh.",
                "D. Tơ tằm thuộc loại tơ nhân tạo."
            ],
            "solution": "A đúng: Poly(vinyl chloride) (PVC) được điều chế bằng phản ứng trùng hợp vinyl chloride (CH~2~=CHCl). B sai vì tơ visco là tơ bán tổng hợp (nhân tạo); C sai vì amylose có mạch không phân nhánh; D sai vì tơ tằm là tơ thiên nhiên."
        },
        {
            "id": 13,
            "prompt": "Câu 13: Một tripeptide X được cấu thành từ 2 phân tử Ala và 1 phân tử Gly. Công thức cấu tạo của X không thể là",
            "options": [
                "*A. Gly-Ala-Gly.*",
                "B. Gly-Ala-Ala.",
                "C. Ala-Gly-Ala.",
                "D. Ala-Ala-Gly."
            ],
            "solution": "Tripeptide X tạo bởi 2 phân tử Ala và 1 phân tử Gly nên thành phần của X bắt buộc phải có đúng 2 gốc Ala và 1 gốc Gly. Gly-Ala-Gly được tạo thành từ 2 Gly và 1 Ala nên không thể là công thức của X."
        },
        {
            "id": 14,
            "prompt": "Câu 14: Có thể phân biệt glucose và fructose bằng cách cho từng mẫu chất tác dụng với",
            "options": [
                "A. thuốc thử Tollens.",
                "*B. nước bromine.*",
                "C. dung dịch NH~3~.",
                "D. Cu(OH)~2~ trong môi trường kiềm, đun nóng."
            ],
            "solution": "Dùng nước bromine để phân biệt: Glucose chứa nhóm aldehyde (-CHO) làm mất màu nâu đỏ của nước bromine tạo gluconic acid; fructose chứa nhóm ketone (-CO-) không làm mất màu nước bromine (vì môi trường nước bromine có tính acid yếu, fructose không bị chuyển hóa thành glucose). Trong khi đó, thuốc thử Tollens và Cu(OH)~2~/OH^-^ đun nóng đều có môi trường base kiềm làm fructose chuyển hóa thành glucose nên cả hai đều phản ứng tương tự nhau."
        },
        {
            "id": 15,
            "prompt": "Câu 15: Cao su buna – N (hay còn gọi là cao su nitrile, có kí hiệu là NBR) là loại cao su tổng hợp có khả năng chịu dầu mỡ tốt nên dùng làm ống dẫn nhiên liệu, gioăng làm kín trong các máy móc. Thực hiện phản ứng trùng hợp các chất nào dưới đây thu được sản phẩm cao su buna – N?",
            "options": [
                "A. CH~2~=CHCH=CH~2~ và C~6~H~5~CH = CH~2~.",
                "B. CH~2~=C(CH~3~)CH=CH~2~ và CH~2~=CHCN.",
                "*C. CH~2~=CHCH=CH~2~ và CH~2~=CHCN.*",
                "D. CH~2~=CHCH=CH~2~ và N~2~."
            ],
            "solution": "Cao su buna-N được sản xuất bằng phản ứng đồng trùng hợp giữa buta-1,3-diene (CH~2~=CH-CH=CH~2~) và acrylonitrile (vinyl cyanide, CH~2~=CH-CN) có xúc tác."
        },
        {
            "id": 16,
            "prompt": "Câu 16: Tên gọi của HCOOC~2~H~5~ là",
            "options": [
                "A. methyl acetate.",
                "B. ethyl acetate.",
                "*C. ethyl formate.*",
                "D. methyl formate."
            ],
            "solution": "Gốc alcohol -C~2~H~5~ là ethyl; gốc acid HCOO- là formate. Tên gọi của ester là ethyl formate."
        }
    ],

    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1: Poly(ethylene terephthalate) (viết tắt là PET) là một polymer được điều chế từ terephthalic acid và ethylene glycol theo phương trình:\n"
                      "PET được sử dụng để sản xuất tơ, chai đựng nước uống, hộp đựng thực phẩm. Để thuận lợi cho việc nhận biết, sử dụng và tái chế thì các đồ nhựa làm từ vật liệu chứa PET thường được in kí hiệu như hình trên.",
            "image": "phan_ung_pet.png",
            "image_width": Cm(12.0),
            "statements": [
                ("a) Trong một mắt xích PET, phần trăm khối lượng carbon là 62,5%.", "ĐÚNG"),
                ("b) Phản ứng tổng hợp PET từ terephthalic acid và ethylene glycol thuộc loại phản ứng trùng hợp.", "SAI"),
                ("c) Sản phẩm làm bằng nhựa số 01 (PET) có thể được tái chế nhưng không nên tái sử dụng để đựng thực phẩm, nước uống.", "ĐÚNG"),
                ("d) Tơ được chế tạo từ PET thuộc loại tơ tổng hợp.", "ĐÚNG")
            ],
            "solution": "- a) ĐÚNG: Một mắt xích PET có cấu trúc -O-CH~2~-CH~2~-OOC-C~6~H~4~-CO- tương ứng với công thức phân tử C~10~H~8~O~4~. Phân tử khối: M = 10*12 + 8*1 + 4*16 = 120 + 8 + 64 = 192 (g/mol). Phần trăm khối lượng carbon là: %C = (120 / 192) * 100% = 62,5%.\n"
                       "- b) SAI: Phản ứng kết hợp các monomer đồng thời giải phóng phân tử nhỏ H~2~O nên thuộc loại phản ứng TRÙNG NGƯNG (đồng trùng ngưng), không phải trùng hợp.\n"
                       "- c) ĐÚNG: Nhựa PET (số 01) được sản xuất cho mục đích sử dụng một lần. Khi tái sử dụng nhiều lần, cấu trúc bề mặt bị xốp, dễ tích tụ vi khuẩn và thôi nhiễm các hợp chất hóa học/chất phụ gia vào nước uống, thực phẩm nên được khuyến cáo không nên tái sử dụng để đựng thức ăn, nước uống.\n"
                       "- d) ĐÚNG: Tơ kéo từ PET (tơ lapsan/dacron) là loại polymer tổng hợp hoàn toàn nên thuộc loại tơ tổng hợp."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Trong phòng thí nghiệm, một học sinh nhúng thanh đồng có khối lượng 12,340 g vào 255 mL dung dịch AgNO~3~ 0,125M. Bằng quan sát, học sinh đó đã khẳng định có phản ứng xảy ra.",
            "statements": [
                ("a) Tính oxi hóa của Cu^2+^ mạnh hơn Ag^+^.", "SAI"),
                ("b) Dung dịch từ không màu chuyển sang màu nâu đỏ.", "SAI"),
                ("c) Phương trình hóa học của phản ứng: Cu + 2AgNO~3~ -> Cu(NO~3~)~2~ + 2Ag", "ĐÚNG"),
                ("d) Giả thuyết toàn bộ lượng Ag giải phóng đều bám vào thanh đồng thì khối lượng thanh đồng sau khi phản ứng kết thúc là 13,35 gam.", "SAI")
            ],
            "solution": "- a) SAI: Thế điện cực chuẩn E°(Ag^+^/Ag) = +0,80V > E°(Cu^2+^/Cu) = +0,34V nên tính oxi hóa của Ag^+^ mạnh hơn Cu^2+^ (Ag^+^ có khả năng oxi hóa Cu thành Cu^2+^).\n"
                       "- b) SAI: Dung dịch AgNO~3~ ban đầu không màu, sau phản ứng tạo muối Cu(NO~3~)~2~ làm dung dịch chuyển sang màu xanh lam đặc trưng của ion Cu^2+^, không phải màu nâu đỏ.\n"
                       "- c) ĐÚNG: Phương trình hóa học đúng là Cu + 2AgNO~3~ -> Cu(NO~3~)~2~ + 2Ag.\n"
                       "- d) SAI: Tính toán:\n"
                       "  + n(AgNO~3~) = 0,255 * 0,125 = 0,031875 mol.\n"
                       "  + Theo PTHH: n(Cu tan) = n(AgNO~3~) / 2 = 0,0159375 mol; n(Ag bám) = 0,031875 mol.\n"
                       "  + Độ tăng khối lượng của thanh kim loại: Δm = m(Ag bám) - m(Cu tan) = (0,031875 * 108) - (0,0159375 * 64) = 3,4425 - 1,02 = 2,4225 gam.\n"
                       "  + Khối lượng thanh đồng sau phản ứng: m_sau = 12,340 + 2,4225 = 14,7625 gam (khác giá trị 13,35 gam)."
        }
    ],

    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1: Cho các phát biểu:\n"
                      "(a) Oxi hóa glucose bằng thuốc thử Tollens tạo ra gluconic acid.\n"
                      "(b) Dung dịch của glucose hay saccharose đều có thể hòa tan Cu(OH)~2~ ở điều kiện thường.\n"
                      "(c) Cellulose trinitrate là nguyên liệu để sản xuất thuốc súng không khói.\n"
                      "(d) Trong cellulose, các gốc \u03b2-glucose chỉ liên kết với nhau bằng liên kết \u03b2-1,4-glycoside.\n"
                      "(e) Glucose và fructose có thể chuyển hóa lẫn nhau trong môi trường kiềm.\n"
                      "(f) Glucose là hợp chất thuộc loại monosaccharide.\n"
                      "Số phát biểu đúng là bao nhiêu?",
            "answer": "5",
            "solution": "Phân tích từng phát biểu:\n"
                       "- (a) SAI: Thuốc thử Tollens (AgNO~3~/NH~3~) oxi hóa glucose trong môi trường base kiềm tạo thành muối ammonium gluconate (CH~2~OH[CHOH]~4~COONH~4~), không tạo trực tiếp gluconic acid tự do (muốn tạo gluconic acid người ta dùng nước bromine).\n"
                       "- (b) ĐÚNG: Cả glucose và saccharose đều là polyalcohol có nhiều nhóm -OH liền kề nên hòa tan Cu(OH)~2~ ở nhiệt độ thường tạo phức màu xanh lam.\n"
                       "- (c) ĐÚNG: Cellulose trinitrate [C~6~H~7~O~2~(ONO~2~)~3~]~n~ rất dễ cháy và nổ mạnh, là nguyên liệu sản xuất thuốc súng không khói.\n"
                       "- (d) ĐÚNG: Phân tử cellulose mạch thẳng không nhánh, các gốc \u03b2-glucose liên kết với nhau duy nhất qua liên kết \u03b2-1,4-glycosidic.\n"
                       "- (e) ĐÚNG: Trong môi trường kiềm, glucose và fructose chuyển hóa thuận nghịch lẫn nhau qua dạng trung gian enediol.\n"
                       "- (f) ĐÚNG: Glucose là đường đơn (monosaccharide) không bị thủy phân.\n"
                       "-> Có 5 phát biểu đúng là (b), (c), (d), (e), (f).\n"
                       "*(Ghi chú: Nếu quan niệm rộng phản ứng Tollens oxi hóa nhóm aldehyde tạo gốc gluconate/acid thì có 6 phát biểu).* "
        },
        {
            "id": 2,
            "prompt": "Câu 2: Cho các loại tơ: bông, tơ capron, tơ cellulose acetate, tơ tằm, len, tơ nitron, nylon-6,6. Có bao nhiêu tơ thuộc loại tơ tổng hợp?",
            "answer": "3",
            "solution": "Phân loại các loại tơ sợi:\n"
                       "- Tơ thiên nhiên: bông, tơ tằm, len (3 loại).\n"
                       "- Tơ bán tổng hợp (nhân tạo): tơ cellulose acetate (1 loại).\n"
                       "- Tơ tổng hợp: tơ capron, tơ nitron (olon), nylon-6,6 (3 loại).\n"
                       "-> Có 3 loại tơ thuộc loại tơ tổng hợp."
        },
        {
            "id": 3,
            "prompt": "Câu 3: Xét pin Galvani hoạt động với phương trình tương ứng: Zn(s) + Cu^2+^(aq) -> Cu(s) + Zn^2+^(aq)\n"
                      "Cho các phát biểu:\n"
                      "(a) Điện cực Copper giảm khối lượng và điện cực Copper là cực dương.\n"
                      "(b) Sức điện động của pin E°_pin = E°(+) - E°(-).\n"
                      "(c) Điện cực Zinc giảm khối lượng và điện cực Zinc là cực âm.\n"
                      "(d) Khi pin hoạt động dòng điện có chiều từ Copper sang Zinc.\n"
                      "(e) Thế điện cực chuẩn của cặp Zn^2+^/Zn nhỏ hơn của cặp Cu^2+^/Cu.\n"
                      "Có bao nhiêu phát biểu đúng trong các phát biểu trên?",
            "answer": "4",
            "solution": "Phân tích các phát biểu:\n"
                       "- (a) SAI: Tại cathode (cực dương Cu), Cu^2+^ bị khử bám vào điện cực nên khối lượng điện cực Cu tăng, không phải giảm.\n"
                       "- (b) ĐÚNG: Sức điện động chuẩn của pin: E°_pin = E°(cathode) - E°(anode) = E°(+) - E°(-).\n"
                       "- (c) ĐÚNG: Tại anode (cực âm Zn), Zn bị oxi hóa tan ra nên điện cực Zn giảm khối lượng.\n"
                       "- (d) ĐÚNG: Dòng electron chạy từ cực âm Zn sang cực dương Cu, do đó chiều dòng điện quy ước chạy từ cực dương (Cu) sang cực âm (Zn).\n"
                       "- (e) ĐÚNG: E°(Zn^2+^/Zn) = -0,76V < E°(Cu^2+^/Cu) = +0,34V.\n"
                       "-> Có 4 phát biểu đúng là (b), (c), (d), (e)."
        },
        {
            "id": 4,
            "prompt": "Câu 4: Có bao nhiêu hợp chất đơn chức ứng với công thức phân tử C~3~H~6~O~2~ đều tác dụng được với dung dịch NaOH?",
            "answer": "3",
            "solution": "Hợp chất đơn chức mạch hở có CTPT C~3~H~6~O~2~ (độ bất bão hòa k = 1, có 2 nguyên tử O) tác dụng được với dung dịch NaOH gồm:\n"
                       "1. Carboxylic acid (tác dụng với NaOH theo phản ứng trung hòa): CH~3~-CH~2~-COOH (propanoic acid) (1 chất).\n"
                       "2. Ester (tác dụng với NaOH theo phản ứng xà phòng hóa):\n"
                       "   - HCOO-CH~2~-CH~3~ (ethyl formate)\n"
                       "   - CH~3~-COO-CH~3~ (methyl acetate)\n"
                       "   (2 chất).\n"
                       "-> Tổng số hợp chất đơn chức thỏa mãn là: 1 + 2 = 3 chất."
        }
    ],

    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1: (1,0 điểm). Mỡ bò có hàm lượng tristearin 80% về khối lượng, các loại chất béo khác có hàm lượng không đáng kể và một số tạp chất khác. Xà phòng hóa hoàn toàn 222,5 gam mỡ bò trong dung dịch KOH, thu được m gam potassium stearate.\n"
                      "a) Viết phương trình hóa học?\n"
                      "b) Tính giá trị của m? Cho rằng các tạp chất không tham gia phản ứng với KOH.",
            "num_lines": 7,
            "solution": "a) Phương trình phản ứng xà phòng hóa tristearin bằng dung dịch KOH:\n"
                       "   (C~17~H~35~COO)~3~C~3~H~5~ + 3KOH --[t°]--> 3C~17~H~35~COOK + C~3~H~5~(OH)~3~\n"
                       "b) Tính giá trị của m:\n"
                       "   - Khối lượng tristearin có trong 222,5 g mỡ bò: m(tristearin) = 222,5 * 80% = 178 gam.\n"
                       "   - Phân tử khối tristearin (C~17~H~35~COO)~3~C~3~H~5~: M = 3 * 283 + 41 = 890 (g/mol).\n"
                       "   - Số mol tristearin: n = 178 / 890 = 0,2 mol.\n"
                       "   - Theo phương trình hóa học: n(potassium stearate) = 3 * n(tristearin) = 3 * 0,2 = 0,6 mol.\n"
                       "   - Phân tử khối C~17~H~35~COOK: M = 283 + 39 = 322 (g/mol).\n"
                       "   - Khối lượng muối potassium stearate thu được: m = 0,6 * 322 = 193,2 gam.\n"
                       "Đáp số: m = 193,2 gam."
        },
        {
            "id": 2,
            "prompt": "Câu 2: (1,0 điểm). Hoàn thành các phương trình phản ứng sau:\n"
                      "a) Glucose + dung dịch Br~2~\n"
                      "b) Glycine + NaOH\n"
                      "c) Propan-1-amine + HCl\n"
                      "d) Methyl acetate + KOH (t°)",
            "num_lines": 7,
            "solution": "Các phương trình phản ứng hóa học:\n"
                       "a) Oxi hóa glucose bằng nước bromine:\n"
                       "   CH~2~OH[CHOH]~4~CHO + Br~2~ + H~2~O -> CH~2~OH[CHOH]~4~COOH + 2HBr\n"
                       "b) Phản ứng của amino acid glycine với dung dịch kiềm:\n"
                       "   H~2~N-CH~2~-COOH + NaOH -> H~2~N-CH~2~-COONa + H~2~O\n"
                       "c) Tính base của propan-1-amine tác dụng với acid mạnh:\n"
                       "   CH~3~CH~2~CH~2~NH~2~ + HCl -> CH~3~CH~2~CH~2~NH~3~Cl\n"
                       "d) Thủy phân ester methyl acetate trong môi trường kiềm KOH:\n"
                       "   CH~3~COOCH~3~ + KOH --[t°]--> CH~3~COOK + CH~3~OH"
        },
        {
            "id": 3,
            "prompt": "Câu 3: (1,0 điểm). Cho các pin điện hoá và sức điện động chuẩn tương ứng:\n"
                      "Bảng dữ liệu: Pin Ni - Cu có E° = 0,59V; Pin Zn - Cu có E° = 1,10V; Pin Ni - Sn có E° = 0,11V.\n"
                      "a) Viết phương trình ion diễn ra khi pin Zn - Sn hoạt động.\n"
                      "b) Sức điện động chuẩn của pin Zn- Sn bằng bao nhiêu V(volt)?",
            "num_lines": 8,
            "solution": "a) Viết phương trình ion khi pin Zn - Sn hoạt động:\n"
                       "- Tính khử của Zn mạnh hơn Sn (E°(Zn^2+^/Zn) < E°(Sn^2+^/Sn)), do đó:\n"
                       "  + Tại anode (cực âm): Zn(s) -> Zn^2+^(aq) + 2e\n"
                       "  + Tại cathode (cực dương): Sn^2+^(aq) + 2e -> Sn(s)\n"
                       "- Phương trình ion rút gọn của phản ứng diễn ra trong pin:\n"
                       "  Zn(s) + Sn^2+^(aq) -> Zn^2+^(aq) + Sn(s)\n"
                       "b) Tính sức điện động chuẩn của pin Zn - Sn:\n"
                       "- Theo công thức sức điện động chuẩn của pin Galvani:\n"
                       "  (1) E°_pin(Zn - Cu) = E°(Cu^2+^/Cu) - E°(Zn^2+^/Zn) = 1,10 V\n"
                       "  (2) E°_pin(Ni - Cu) = E°(Cu^2+^/Cu) - E°(Ni^2+^/Ni) = 0,59 V\n"
                       "  (3) E°_pin(Ni - Sn) = E°(Sn^2+^/Sn) - E°(Ni^2+^/Ni) = 0,11 V\n"
                       "- Lấy phương trình (1) trừ phương trình (2):\n"
                       "  E°(Ni^2+^/Ni) - E°(Zn^2+^/Zn) = 1,10 - 0,59 = 0,51 V  (chính là E°_pin(Zn - Ni))\n"
                       "- Cộng tiếp với phương trình (3):\n"
                       "  E°_pin(Zn - Sn) = E°(Sn^2+^/Sn) - E°(Zn^2+^/Zn)\n"
                       "                  = [E°(Sn^2+^/Sn) - E°(Ni^2+^/Ni)] + [E°(Ni^2+^/Ni) - E°(Zn^2+^/Zn)]\n"
                       "                  = 0,11 + 0,51 = 0,62 (V).\n"
                       "Kết luận: Sức điện động chuẩn của pin Zn - Sn là 0,62 V."
        }
    ]
}


def generate_docx(is_teacher=False):
    """Tạo file docx theo chuẩn Thầy TRẦN VĂN THẠNH."""
    doc = Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line=DATA_EXAM_104["school_line"],
        exam_title=DATA_EXAM_104["exam_title"],
        subject_line=DATA_EXAM_104["subject_line"],
        num_pages=DATA_EXAM_104["num_pages"],
        duration=DATA_EXAM_104["duration"],
        exam_code=DATA_EXAM_104["exam_code"],
        atomic_masses=DATA_EXAM_104["atomic_masses"],
        is_teacher=is_teacher
    )

    if is_teacher:
        add_compact_summary_answers(doc)

    # PHẦN I
    add_section_header(
        doc,
        title="A. TRẮC NGHIỆM: (7,0 điểm)",
        subtitle="Phần I: Trắc nghiệm nhiều lựa chọn (4,0 điểm). Học sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi học sinh chỉ chọn một phương án."
    )

    for item in DATA_EXAM_104["part1"]:
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
        title="Phần II: Trắc nghiệm dạng Đúng/Sai (2,0 điểm).",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai."
    )

    for item in DATA_EXAM_104["part2"]:
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
        title="Phần III: Trắc nghiệm trả lời ngắn (1,0 điểm).",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4."
    )

    for item in DATA_EXAM_104["part3"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

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
        title="B. TỰ LUẬN (3,0 điểm).",
        subtitle="Thí sinh trình bày chi tiết lời giải từ câu 1 đến câu 3 vào giấy làm bài thi."
    )

    for item in DATA_EXAM_104["part4"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        if item["id"] == 3:
            # Chèn bảng pin điện hóa cho Câu 3 Tự luận
            tbl_c3 = doc.add_table(rows=2, cols=4)
            tbl_c3.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl_c3.autofit = False
            headers_c3 = ["Pin điện hóa", "Ni - Cu", "Zn - Cu", "Ni - Sn"]
            values_c3 = ["Sức điện động chuẩn (V)", "0,59", "1,10", "0,11"]
            for r_idx, r_data in enumerate([headers_c3, values_c3]):
                row = tbl_c3.rows[r_idx]
                for c_idx, val in enumerate(r_data):
                    cell = row.cells[c_idx]
                    cell.width = Cm(4.5) if c_idx == 0 else Cm(3.2)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    p_c = cell.paragraphs[0]
                    p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_c.paragraph_format.space_before = Pt(1)
                    p_c.paragraph_format.space_after = Pt(1)
                    if r_idx == 0:
                        set_cell_shading(cell, "E8F5E9")
                        render_rich_text(p_c, f"*{val}*", base_size=9.5, base_bold=True, base_color=COLOR_PRIMARY)
                    else:
                        set_cell_shading(cell, "FFFFFF")
                        render_rich_text(p_c, f"*{val}*", base_size=9.5, base_bold=(c_idx > 0))
                    set_cell_border(cell,
                                    top=("single", "4", "0", "8BC390"),
                                    bottom=("single", "4", "0", "8BC390"),
                                    left=("single", "4", "0", "8BC390"),
                                    right=("single", "4", "0", "8BC390"))
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(2)
            p_sub.paragraph_format.space_after = Pt(2)
            p_sub.paragraph_format.left_indent = Cm(0.5)
            render_rich_text(p_sub, "a) Viết phương trình ion diễn ra khi pin Zn - Sn hoạt động.\nb) Sức điện động chuẩn của pin Zn - Sn bằng bao nhiêu V(volt)?", base_size=10.5)

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
    filename = f"De_Kiem_Tra_CK1_Hoa_12_DangHuyTru_Ma104{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
