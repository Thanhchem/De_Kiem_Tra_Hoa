# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì I - Hóa học 10 - Trường THPT Gia Hội (TP Huế)
Mã đề: 2101 - Năm học 2026 - 2027
Chuẩn theo Quy chuẩn Form Đề Kiểm Tra của Thầy TRẦN VĂN THẠNH
Tạo 2 phiên bản: Bản Học Sinh & Bản Giáo Viên (có đáp án đậm đỏ & lời giải chi tiết)
"""

import os
import sys
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
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
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_gk1_giahoi_2101")
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
    tcPr.append(tcBorders)

def set_cell_shading(cell, color_hex):
    """Tô màu nền cho ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def add_quick_answer_table_2101(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Giáo viên ở đầu tài liệu Mã đề 2101.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 2101)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_p1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
    answers_p1 = ["Đ/A", "C", "D", "A", "A", "B", "D", "D", "C", "B", "D", "A", "A"]

    w1_c0 = Cm(1.8)
    w1_c = Cm(1.35)

    for r_idx, r_data in enumerate([headers_p1, answers_p1]):
        row = tbl1.rows[r_idx]
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = w1_c0 if c_idx == 0 else w1_c
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

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

    # 2. PHẦN II
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(5)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM - MỖI CÂU 1,0 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl2.autofit = False

    tbl2_data = [
        ["Câu hỏi", "Nội dung câu hỏi", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Phổ khối lượng của nguyên tố R", "ĐÚNG", "SAI", "SAI", "ĐÚNG"],
        ["Câu 2", "Sơ đồ cấu tạo nguyên tử X (Lithium-7)", "ĐÚNG", "SAI", "SAI", "ĐÚNG"]
    ]

    w2_widths = [Cm(2.0), Cm(6.0), Cm(2.5), Cm(2.5), Cm(2.5), Cm(2.5)]

    for r_idx, row_vals in enumerate(tbl2_data):
        row = tbl2.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = w2_widths[c_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx != 1 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=9.5, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                is_label = (c_idx < 2)
                set_cell_shading(cell, "F5F5F5" if is_label else "FFFFFF")
                if is_label:
                    c_color = COLOR_PRIMARY if c_idx == 0 else RGBColor(0, 0, 0)
                else:
                    c_color = COLOR_CORRECT if val == "ĐÚNG" else COLOR_GRAY
                render_rich_text(p, f"*{val}*", base_size=9.5, base_bold=True, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # 3. PHẦN III
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(5)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM - MỖI CÂU 0,5 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False

    tbl3_headers = ["Câu", "Câu 1", "Câu 2", "Câu 3", "Câu 4"]
    tbl3_answers = ["Đáp án", "81", "35", "2", "16"]

    w3_c0 = Cm(2.4)
    w3_c = Cm(3.9)

    for r_idx, row_vals in enumerate([tbl3_headers, tbl3_answers]):
        row = tbl3.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = w3_c0 if c_idx == 0 else w3_c
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

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

    # 4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN
    p_tl = doc.add_paragraph()
    p_tl.paragraph_format.space_before = Pt(5)
    p_tl.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl, "*4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN (3,0 ĐIỂM - MỖI CÂU 1,0 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl4 = doc.add_table(rows=4, cols=3)
    tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl4.autofit = False

    tbl4_data = [
        ("Câu", "Điểm", "Đáp số cốt lõi & Hướng dẫn chấm"),
        ("Câu 1", "1,0 đ", "• a) Nguyên tử khối trung bình: Ā~Ne~ = 20,19 (amu).\n• b) Số mol neon: n~Ne~ = 6,057 / 20,19 = 0,3 mol ⇒ n~^20Ne~ = 0,3 × 90% = 0,27 mol ⇒ Khối lượng: m~^20Ne~ = 0,27 × 20 = 5,4 gam."),
        ("Câu 2", "1,0 đ", "• a) Loại hạt mang điện âm là electron (e); điện tích quy ước: -1 (tuyệt đối: -1,602.10^-19^ C).\n• b) Điền bảng: Cột 1 (^28_13^T): Z=13, +13, N=15, A=28; Cột 2 (^31_15^P): Z=15, +15, N=16, A=31; Cột 3: ^13_6^C, +6, A=13; Cột 4: ^24_12^Mg, Z=12, N=12."),
        ("Câu 3", "1,0 đ", "• a) Sulfur (Z=16): 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~4~; có 2 electron độc thân ở phân lớp 3p ([↑↓][↑ ][↑ ]).\n• b) Manganese (Z=25): 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~ 3d~5~ 4s~2~ ([Ar]3d~5~4s~2~); Vị trí: Ô 25, chu kì 4, nhóm VIIB.")
    ]

    w4_widths = [Cm(2.0), Cm(1.5), Cm(14.5)]

    for r_idx, (c1, c2, c3) in enumerate(tbl4_data):
        row = tbl4.rows[r_idx]
        for c_idx, val in enumerate([c1, c2, c3]):
            cell = row.cells[c_idx]
            cell.width = w4_widths[c_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx < 2 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=9.5, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                set_cell_shading(cell, "F5F5F5" if c_idx < 2 else "FFFFFF")
                c_color = COLOR_PRIMARY if c_idx == 0 else (COLOR_CORRECT if c_idx == 1 else None)
                render_rich_text(p, f"*{val}*" if c_idx < 2 else val, base_size=9.5, base_bold=(c_idx < 2), base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # Đường phân cách
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(6)
    p_div.paragraph_format.space_after = Pt(6)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


# Dữ liệu Đề thi mã 2101 - THPT Gia Hội
DATA_EXAM = {
    "school_line": "SỞ GD&ĐT THÀNH PHỐ HUẾ - TRƯỜNG THPT GIA HỘI",
    "exam_title": "KIỂM TRA GIỮA KỲ I NĂM HỌC 2026 - 2027",
    "subject_line": "MÔN: HOÁ HỌC 10",
    "duration": "45 phút (không kể thời gian phát đề)",
    "num_pages": "03",
    "exam_code": "2101",
    "atomic_masses": "H: 1; C: 12; O: 16; N: 14; S: 32.",

    # PHẦN I. (3,0 điểm) Câu trắc nghiệm nhiều phương án lựa chọn.
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1. Phát biểu nào sau đây đúng?",
            "options": [
                "A. Proton và neutron là các hạt mang điện, electron là hạt không mang điện.",
                "B. Proton tạo nên lớp vỏ nguyên tử.",
                "*C. Số lượng proton và electron trong nguyên tử là bằng nhau.*",
                "D. Khối lượng của proton lớn hơn rất nhiều so với khối lượng của neutron."
            ],
            "solution": "Trong nguyên tử trung hòa về điện:\n- Hạt nhân chứa proton mang điện dương (+1) và neutron không mang điện. Lớp vỏ chứa electron mang điện âm (-1).\n- Số đơn vị điện tích hạt nhân bằng số proton và bằng số electron trong vỏ: p = e.\n- Khối lượng của proton xấp xỉ khối lượng của neutron (m~p~ ≈ m~n~ ≈ 1 amu), và đều lớn hơn khối lượng electron rất nhiều (khoảng 1836 lần).\n*Đáp án đúng: C.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Nguyên tử của nguyên tố X có tổng số electron trong các phân lớp p là 9. Nguyên tố X là nguyên tố",
            "options": [
                "A. khí hiếm.",
                "B. kim loại.",
                "C. kim loại hoặc phi kim.",
                "*D. phi kim.*"
            ],
            "solution": "Cấu hình electron nguyên tử của X được điền lần lượt theo mức năng lượng:\n  1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~3~\nTổng số electron trên các phân lớp p là 6 + 3 = 9 (thỏa mãn đề bài).\nSố electron lớp ngoài cùng (lớp n = 3) là 2 + 3 = 5 electron. Với 5 electron ở lớp ngoài cùng, X là nguyên tố phi kim (Phosphorus - P, Z = 15).\n*Đáp án đúng: D.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Nguyên tố X có số hiệu nguyên tử bằng 9. Vị trí của X trong bảng tuần hoàn các nguyên tố hoá học là",
            "options": [
                "*A. Ô 9, chu kì 2, nhóm VIIA.*",
                "B. Ô 9, chu kì 5, nhóm VA.",
                "C. Ô 9, chu kì 2, nhóm VIIB.",
                "D. Ô 9, chu kì 3, nhóm VA."
            ],
            "solution": "Nguyên tử X có Z = 9. Cấu hình electron: 1s~2~ 2s~2~ 2p~5~.\n- Số hiệu nguyên tử Z = 9 ⇒ Ô số 9 trong bảng tuần hoàn.\n- Có 2 lớp electron ⇒ Chu kì 2.\n- Là nguyên tố p, có 7 electron ở lớp ngoài cùng (2s~2~ 2p~5~) ⇒ Nhóm VIIA (nguyên tố Fluorine, F).\n*Đáp án đúng: A.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Orbital p có dạng hình học là",
            "options": [
                "*A. hình số tám nổi.*",
                "B. hình số tám.",
                "C. hình tròn.",
                "D. hình cầu."
            ],
            "solution": "Theo SGK Hóa học 10 (Chương trình GDPT 2018):\n- Orbital s có dạng hình cầu.\n- Orbital p có dạng hình số tám nổi (gồm 3 orbital p~x~, p~y~, p~z~ định hướng theo 3 trục không gian tương ứng x, y, z).\n*Đáp án đúng: A.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5. Hình vẽ sau mô tả thí nghiệm tìm ra một hạt cấu tạo nên nguyên tử.",
            "image": "San_Pham/images_gk1_giahoi_2101/cau5_ong_tia_am_cuc_clean.png",
            "image_width": 7.5,
            "sub_prompt": "Đó là",
            "options": [
                "A. thí nghiệm tìm ra hạt nhân.",
                "*B. thí nghiệm tìm ra electron.*",
                "C. thí nghiệm tìm ra neutron.",
                "D. thí nghiệm tìm ra proton."
            ],
            "solution": "Hình vẽ mô tả thí nghiệm phóng điện qua chất khí trong ống tia catot (ống Crookers) của nhà bác học J.J. Thomson (năm 1897). Chùm tia phát ra từ catot bị lệch về phía cực dương của điện trường chứng tỏ chùm tia gồm các hạt mang điện tích âm, từ đó phát hiện ra hạt electron.\n*Đáp án đúng: B.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6. Khí chlorine (Cl~2~) được dùng phổ biến để diệt trùng nước sinh hoạt. Kí hiệu của nguyên tử chlorine có 17 proton và 20 neutron là",
            "options": [
                "A. ^17_37^Cl.",
                "B. ^35_17^Cl.",
                "C. ^17_35^Cl.",
                "*D. ^37_17^Cl.*"
            ],
            "solution": "Kí hiệu nguyên tử có dạng ^A_Z^X:\n- Số hiệu nguyên tử Z = p = 17.\n- Số khối A = Z + N = 17 + 20 = 37.\nDo đó, kí hiệu nguyên tử chlorine là ^37_17^Cl.\n*Đáp án đúng: D.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7. Biết điện tích của 1 proton là +1,602.10^-19^ C. Nguyên tử A có điện tích hạt nhân là +12,816.10^-19^ C, trong hạt nhân của A chứa số hạt mang điện là",
            "options": [
                "A. 6.",
                "B. 15.",
                "C. 13.",
                "*D. 8.*"
            ],
            "solution": "Trong hạt nhân nguyên tử, chỉ có proton là hạt mang điện dương (+1,602.10^-19^ C).\nSố hạt mang điện trong hạt nhân của nguyên tử A là:\n  p = (+12,816.10^-19^ C) / (+1,602.10^-19^ C) = 8 hạt proton.\n*Đáp án đúng: D.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8. Cho nguyên tử có ký hiệu ^39_19^Z, điều khẳng định nào sau đây không đúng?",
            "options": [
                "A. Nguyên tử khối của ^39_19^Z là 39 amu.",
                "B. Nguyên tử ^39_19^Z có 20 neutron.",
                "*C. Nguyên tử ^39_19^Z có 20 electron.*",
                "D. Nguyên tử ^39_19^Z có 19 proton."
            ],
            "solution": "Từ kí hiệu ^39_19^Z ta có:\n- Số proton p = Z = 19.\n- Trong nguyên tử trung hòa, số electron e = p = 19 (khẳng định C ghi 20 electron là SAI).\n- Số khối A = 39 ⇒ Nguyên tử khối xấp xỉ 39 amu.\n- Số neutron N = A - Z = 39 - 19 = 20.\n*Đáp án đúng: C.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9. Trong bảng tuần hoàn hiện nay, các nguyên tố hoá học thuộc cùng một chu kì có cùng",
            "options": [
                "A. số electron hóa trị.",
                "*B. số lớp electron.*",
                "C. số proton trong nguyên tử.",
                "D. số electron lớp ngoài cùng."
            ],
            "solution": "Chu kì là dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được sắp xếp theo chiều tăng dần của điện tích hạt nhân. Số thứ tự của chu kì bằng số lớp electron trong nguyên tử.\n*Đáp án đúng: B.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10. Để học tốt môn hoá học, theo em cần làm những gì sau đây?",
            "options": [
                "A. Vận dụng kiến thức để giải quyết một số tình huống trong thực tế.",
                "B. Đặt giả thuyết khoa học, xây dựng thí nghiệm để chứng minh, phân tích.",
                "C. Chịu khó quan sát và đặt câu hỏi.",
                "*D. Tất cả các phương án trên.*"
            ],
            "solution": "Phương pháp học tập và nghiên cứu Hóa học bao gồm:\n- Quan sát các hiện tượng tự nhiên, đặt câu hỏi nghiên cứu.\n- Đặt ra giả thuyết khoa học, thiết kế thí nghiệm kiểm chứng và phân tích kết quả.\n- Vận dụng kiến thức hóa học đã học để giải thích và giải quyết các bài toán, tình huống thực tế đời sống.\nDo đó, cần kết hợp tất cả các phương án trên.\n*Đáp án đúng: D.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11. Cấu hình electron của nguyên tử Al (Z = 13) ở trạng thái cơ bản là",
            "options": [
                "*A. 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~1~.*",
                "B. 1s~2~ 2s~2~ 2p~6~ 3s~1~ 3p~2~.",
                "C. 1s~2~ 2s~2~ 2p~6~ 3s~2~.",
                "D. 1s~2~ 2s~2~ 2p~6~ 3s~3~."
            ],
            "solution": "Nguyên tử Aluminium (Al) có Z = 13. Phân bố 13 electron vào các phân lớp theo mức năng lượng tăng dần:\n  1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~1~ ([Ne]3s~2~3p~1~).\n*Đáp án đúng: A.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12. Cho 4 nguyên tử: ^16_8^X, ^16_7^Y, ^18_9^Z, ^19_9^T. Cặp nguyên tử nào sau đây là đồng vị của nhau?",
            "options": [
                "*A. Z, T.*",
                "B. Y, Z.",
                "C. X, Z.",
                "D. X, Y."
            ],
            "solution": "Các đồng vị của cùng một nguyên tố hóa học là những nguyên tử có cùng số proton (cùng số hiệu nguyên tử Z) nhưng khác nhau về số neutron (khác nhau về số khối A).\nTrong 4 nguyên tử trên, Z và T đều có số hiệu nguyên tử bằng 9 (cùng có 9 proton, nguyên tố Fluorine) nhưng có số khối khác nhau (A~Z~ = 18, A~T~ = 19). Do đó Z và T là hai đồng vị của nhau.\n*Đáp án đúng: A.*"
        }
    ],

    # PHẦN II. (2,0 điểm) Câu trắc nghiệm đúng sai.
    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1. (1,0 điểm) Cho phổ khối lượng của nguyên tố R ở hình dưới đây, biết nguyên tử khối trung bình của R là 6,92.",
            "image": "San_Pham/images_gk1_giahoi_2101/pho_khoi_r.png",
            "image_width": 8.0,
            "statements": [
                ("a) Đồng vị ^x^R phổ biến hơn so với đồng vị còn lại.", "Đ"),
                ("b) Trong tự nhiên nguyên tố R có ba đồng vị bền.", "S"),
                ("c) Hai đồng vị bền trên có cùng số neutron.", "S"),
                ("d) Số khối x của đồng vị còn lại là 7.", "Đ")
            ],
            "solution": "Quan sát phổ khối lượng của nguyên tố R:\n- Phổ khối lượng chỉ có 2 pic tín hiệu: pic 1 tại m/z = 6 chiếm 7,59% và pic 2 tại m/z = x chiếm 92,41%.\n• a) Đúng. Đồng vị ^x^R chiếm 92,41% số nguyên tử trong tự nhiên, phổ biến hơn rất nhiều so với đồng vị ^6^R (chỉ chiếm 7,59%).\n• b) Sai. Phổ khối lượng chỉ có 2 pic, nghĩa là trong tự nhiên nguyên tố R chỉ có 2 đồng vị bền.\n• c) Sai. Các đồng vị của cùng một nguyên tố có cùng số proton nhưng khác nhau về số neutron.\n• d) Đúng. Áp dụng công thức tính nguyên tử khối trung bình:\n    Ā~R~ = (6 × 7,59 + x × 92,41) / 100 = 6,92\n    ⇔ 45,54 + 92,41x = 692\n    ⇔ 92,41x = 646,46 ⇒ x ≈ 7.\n    Số khối của đồng vị còn lại là 7 (nguyên tố Lithium: ^6^Li và ^7^Li)."
        },
        {
            "id": 2,
            "prompt": "Câu 2. (1,0 điểm) Cho sơ đồ của một nguyên tử X được biểu diễn như sau:",
            "image": "San_Pham/images_gk1_giahoi_2101/so_do_nguyen_tu_x.png",
            "image_width": 7.0,
            "statements": [
                ("a) X có cấu hình electron là 1s~2~ 2s~1~.", "Đ"),
                ("b) Số proton của X bằng 4.", "S"),
                ("c) Trong X, điện tích hạt nhân là +7.", "S"),
                ("d) X có số khối là 7.", "Đ")
            ],
            "solution": "Quan sát sơ đồ cấu tạo nguyên tử X:\n- Lớp vỏ có 2 lớp electron: lớp trong cùng (K) có 2 electron, lớp ngoài cùng (L) có 1 electron ⇒ Tổng số electron là e = 3. Cấu hình electron: 1s~2~ 2s~1~.\n- Trong hạt nhân: có 3 proton (mang điện dương) và 4 neutron (không mang điện).\n• a) Đúng. X có 3 electron phân bố vào 2 lớp: 1s~2~ 2s~1~.\n• b) Sai. Số proton của X là 3 (4 là số neutron).\n• c) Sai. Điện tích hạt nhân của X là +3 (hoặc +3e, +4,806.10^-19^ C). Số 7 là số khối A của nguyên tử.\n• d) Đúng. Số khối của X là A = p + n = 3 + 4 = 7."
        }
    ],

    # PHẦN III. (2,0 điểm) Câu trắc nghiệm trả lời ngắn.
    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1. (0,5 điểm) Boron là nguyên tố có nhiều tác dụng đối với cơ thể người như: làm lành vết thương, điều hòa nội tiết sinh dục, chống viêm khớp,... Do ngọn lửa cháy có màu đặc biệt nên boron vô định hình được dùng làm pháo hoa. Boron có hai đồng vị là ^10^B và ^11^B, nguyên tử khối trung bình là 10,81. Phần trăm số nguyên tử của đồng vị ^11^B là bao nhiêu? (Làm tròn đến phần nguyên)",
            "answer": "81",
            "solution": "Gọi phần trăm số nguyên tử của đồng vị ^11^B là x% (0 < x < 100).\nKhi đó phần trăm số nguyên tử của đồng vị ^10^B là (100 - x)%.\nNguyên tử khối trung bình của Boron là:\n  Ā = [11 × x + 10 × (100 - x)] / 100 = 10,81\n  ⇔ 11x + 1000 - 10x = 1081\n  ⇔ x = 81 (%).\nLàm tròn đến phần nguyên: 81.\n*Đáp án: 81*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. (0,5 điểm) Tổng số hạt cơ bản trong nguyên tử của nguyên tố X là 114, trong đó số hạt mang điện nhiều hơn số hạt không mang điện là 26. Xác định số đơn vị điện tích hạt nhân của X.",
            "answer": "35",
            "solution": "Gọi số proton, neutron, electron của nguyên tử X lần lượt là p, n, e. Trong nguyên tử trung hòa: p = e.\n- Tổng số hạt cơ bản: 2p + n = 114  (1)\n- Số hạt mang điện (2p) nhiều hơn số hạt không mang điện (n) là 26: 2p - n = 26  (2)\nCộng vế với vế (1) và (2):\n  4p = 114 + 26 = 140 ⇒ p = 35.\nSố đơn vị điện tích hạt nhân của X là Z = p = 35 (nguyên tố Bromine, Br).\n*Đáp án: 35*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. (0,5 điểm) Cho 5 nguyên tố có số hiệu nguyên tử lần lượt là 5, 8, 11, 15, 19. Trong số các nguyên tố trên có bao nhiêu nguyên tố ở chu kì 3 trong Bảng tuần hoàn các nguyên tố hoá học?",
            "answer": "2",
            "solution": "Viết cấu hình electron của 5 nguyên tố:\n- Z = 5 (B): 1s~2~ 2s~2~ 2p~1~ ⇒ 2 lớp electron ⇒ Chu kì 2.\n- Z = 8 (O): 1s~2~ 2s~2~ 2p~4~ ⇒ 2 lớp electron ⇒ Chu kì 2.\n- Z = 11 (Na): 1s~2~ 2s~2~ 2p~6~ 3s~1~ ⇒ 3 lớp electron ⇒ Chu kì 3.\n- Z = 15 (P): 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~3~ ⇒ 3 lớp electron ⇒ Chu kì 3.\n- Z = 19 (K): 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~ 4s~1~ ⇒ 4 lớp electron ⇒ Chu kì 4.\nTrong 5 nguyên tố trên, có 2 nguyên tố ở chu kì 3 (Z = 11 và Z = 15).\n*Đáp án: 2*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. (0,5 điểm) Sulfur (S) là chất rắn, xốp, màu vàng hơi nhạt ở điều kiện thường. Sulfur và hợp chất của nó được sử dụng trong acquy, bột giặt, thuốc diệt nấm; do dễ cháy nên S còn được dùng để sản xuất các loại diêm, thuốc súng, pháo hoa,... Trong bảng tuần hoàn, nguyên tố S nằm ở chu kì 3, nhóm VIA. Nguyên tử của nguyên tố S nằm ở ô thứ bao nhiêu của bảng tuần hoàn các nguyên tố hóa học?",
            "answer": "16",
            "solution": "Nguyên tố S nằm ở chu kì 3 nên có 3 lớp electron. Nằm ở nhóm VIA (nhóm A) nên có 6 electron ở lớp ngoài cùng.\nCấu hình electron nguyên tử của S là: 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~4~.\nTổng số electron là 2 + 2 + 6 + 2 + 4 = 16 ⇒ Số hiệu nguyên tử Z = 16.\nSố thứ tự của ô nguyên tố bằng số hiệu nguyên tử Z ⇒ S nằm ở ô thứ 16.\n*Đáp án: 16*"
        }
    ],

    # PHẦN IV. (3,0 điểm) Tự luận.
    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1. (1,0 điểm) Neon là nguyên tố khí hiếm được dùng nhiều trong các thiết bị điện, điện tử, bảng quảng cáo,... Nguyên tố Neon có 3 đồng vị cho ở bảng sau:",
            "table_type": "neon",
            "sub_prompts": [
                "a) Tính nguyên tử khối trung bình của Neon.",
                "b) Tính khối lượng của đồng vị ^20^Ne có trong 6,057 gam neon nguyên chất."
            ],
            "num_lines": 8,
            "solution": "a) Nguyên tử khối trung bình của Neon là:\n  Ā~Ne~ = (20 × 90,0 + 21 × 1,0 + 22 × 9,0) / 100\n  Ā~Ne~ = (1800 + 21 + 198) / 100 = 2019 / 100 = 20,19 (amu).\n\nb) Tính khối lượng của đồng vị ^20^Ne trong 6,057 gam neon nguyên chất:\n- Số mol nguyên tử neon là:\n  n~Ne~ = m~Ne~ / Ā~Ne~ = 6,057 / 20,19 = 0,3 (mol).\n- Vì đồng vị ^20^Ne chiếm 90,0% tổng số nguyên tử nên:\n  n~^20Ne~ = 0,3 × 90,0% = 0,27 (mol).\n- Khối lượng của đồng vị ^20^Ne là:\n  m~^20Ne~ = 0,27 × 20 = 5,4 gam.\n(Hoặc tính trực tiếp: m~^20Ne~ = 6,057 × (20 × 90%) / 20,19 = 5,4 gam)."
        },
        {
            "id": 2,
            "prompt": "Câu 2. (1,0 điểm)\na) Trong nguyên tử, cho biết loại hạt nào mang điện âm và điện tích của nó?\nb) Hoàn thành bảng sau:",
            "table_type": "atom_properties",
            "num_lines": 6,
            "solution": "a) Trong nguyên tử, hạt mang điện âm là hạt electron (kí hiệu: e).\n- Điện tích tương đối (quy ước) là -1.\n- Điện tích tuyệt đối của 1 electron là q~e~ = -1,602.10^-19^ C.\n\nb) Hoàn thành bảng:\n- Cột 1 (^28_13^T): Số hiệu Z = 13; Điện tích hạt nhân = +13; Số neutron N = 28 - 13 = 15; Số khối A = 28.\n- Cột 2 (^31_15^P): Số hiệu Z = 15; Điện tích hạt nhân = +15; Số neutron N = 31 - 15 = 16; Số khối A = 31.\n- Cột 3 (Z = 6, N = 7): Ký hiệu nguyên tử là ^13_6^C (Carbon); Điện tích hạt nhân = +6; Số khối A = 6 + 7 = 13.\n- Cột 4 (Điện tích hạt nhân +12, A = 24): Ký hiệu nguyên tử là ^24_12^Mg (Magnesium); Số hiệu Z = 12; Số neutron N = 24 - 12 = 12."
        },
        {
            "id": 3,
            "prompt": "Câu 3. (1,0 điểm) Sulfur (Z = 16) là nguyên tố phổ biến thứ 5 trên trái đất, thường xuất hiện ở các vùng suối nước nóng, mỏ núi lửa... Manganese (Z = 25) là nguyên tố được ứng dụng rộng rãi trong công nghiệp và đời sống.\na) Viết cấu hình electron của nguyên tử Sulfur, phân bố các electron trên các orbital.\nb) Viết cấu hình electron của nguyên tử Manganese. Xác định vị trí (ô, chu kì, nhóm) của nguyên tử Manganese trong bảng tuần hoàn.",
            "num_lines": 8,
            "solution": "a) Nguyên tử Sulfur (Z = 16):\n- Cấu hình electron: 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~4~ (hoặc [Ne]3s~2~3p~4~).\n- Sự phân bố electron vào các orbital (ô lượng tử):\n  + Lớp 1 (n = 1): Phân lớp 1s có 1 ô orbital chứa 2 electron ghép đôi [↑↓].\n  + Lớp 2 (n = 2): Phân lớp 2s có 1 ô orbital chứa 2 electron ghép đôi [↑↓]; phân lớp 2p có 3 ô orbital bão hòa chứa 6 electron ghép đôi [↑↓][↑↓][↑↓].\n  + Lớp 3 (n = 3 - lớp ngoài cùng): Phân lớp 3s có 1 ô orbital chứa 2 electron ghép đôi [↑↓]; phân lớp 3p có 3 ô orbital chứa 4 electron gồm 1 ô ghép đôi và 2 ô chứa electron độc thân [↑↓][↑ ][↑ ].\n  ⇒ Nguyên tử sulfur có 2 electron độc thân ở phân lớp 3p.\n\nb) Nguyên tử Manganese (Z = 25):\n- Phân bố theo mức năng lượng: 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~ 4s~2~ 3d~5~.\n- Cấu hình electron theo thứ tự lớp: 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~ 3d~5~ 4s~2~ (hoặc [Ar]3d~5~4s~2~).\n- Vị trí của Manganese trong bảng tuần hoàn:\n  + Ô nguyên tố: Ô số 25 (vì có số hiệu nguyên tử Z = 25).\n  + Chu kì: Chu kì 4 (vì nguyên tử có 4 lớp electron).\n  + Nhóm: Nhóm VIIB (vì là nguyên tố kim loại chuyển tiếp họ d, electron hóa trị nằm ở 3d và 4s với tổng số là 5 + 2 = 7)."
        }
    ]
}

def add_table_borders(table, color="CCCCCC", sz="4"):
    """Thêm viền cho bảng dữ liệu đề bài tự luận."""
    tblPr = table._tbl.tblPr
    borders_xml = f'''
    <w:tblBorders {nsdecls("w")}>
        <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        <w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
    </w:tblBorders>
    '''
    tblPr.append(parse_xml(borders_xml))

def render_table_neon(doc):
    """Bảng đề bài Neon ở Câu 1 Phần IV."""
    t = doc.add_table(rows=2, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_table_borders(t, color="888888", sz="6")
    
    headers = [("Nguyên tử", True), ("^20^Ne", False), ("^21^Ne", False), ("^22^Ne", False)]
    row0 = t.rows[0]
    for i, (txt, bold) in enumerate(headers):
        cell = row0.cells[i]
        cell.width = Cm(4.0 if i == 0 else 3.5)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        render_rich_text(p, txt, base_size=10.5, base_bold=bold)

    vals = ["% số nguyên tử", "90,0%", "1,0%", "9,0%"]
    row1 = t.rows[1]
    for i, txt in enumerate(vals):
        cell = row1.cells[i]
        cell.width = Cm(4.0 if i == 0 else 3.5)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        render_rich_text(p, txt, base_size=10.5, base_bold=(i == 0))

def render_table_atom_props(doc, is_teacher=False):
    """Bảng đặc trưng nguyên tử ở Câu 2b Phần IV."""
    t = doc.add_table(rows=5, cols=5)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_table_borders(t, color="888888", sz="6")

    # Dữ liệu bảng: (nhãn, col1, col2, col3, col4)
    # col1: ^28_13^T | col2: ^31_15^P | col3: Z=6, N=7 -> ^13_6^C | col4: +12, A=24 -> ^24_12^Mg
    rows_data = [
        ("Ký hiệu nguyên tử", "^28_13^T", "^31_15^P", ("^13_6^C" if is_teacher else ""), ("^24_12^Mg" if is_teacher else "")),
        ("Số hiệu nguyên tử", ("13" if is_teacher else ""), ("15" if is_teacher else ""), "6", ("12" if is_teacher else "")),
        ("Điện tích hạt nhân", ("+13" if is_teacher else ""), ("+15" if is_teacher else ""), ("+6" if is_teacher else ""), "+12"),
        ("Số neutron", ("15" if is_teacher else ""), ("16" if is_teacher else ""), "7", ("12" if is_teacher else "")),
        ("Số khối", ("28" if is_teacher else ""), ("31" if is_teacher else ""), ("13" if is_teacher else ""), "24")
    ]

    for r_idx, r_data in enumerate(rows_data):
        row = t.rows[r_idx]
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = Cm(4.2 if c_idx == 0 else 3.2)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            
            # Kiểm tra xem ô này là đáp án điền vào ở bản giáo viên
            is_filled_ans = False
            if is_teacher and c_idx > 0:
                # Các ô điền vào:
                # r=0: c=3, c=4
                # r=1: c=1, c=2, c=4
                # r=2: c=1, c=2, c=3
                # r=3: c=1, c=2, c=4
                # r=4: c=1, c=2, c=3
                if (r_idx == 0 and c_idx in [3, 4]) or \
                   (r_idx == 1 and c_idx in [1, 2, 4]) or \
                   (r_idx == 2 and c_idx in [1, 2, 3]) or \
                   (r_idx == 3 and c_idx in [1, 2, 4]) or \
                   (r_idx == 4 and c_idx in [1, 2, 3]):
                    is_filled_ans = True

            color = COLOR_CORRECT if is_filled_ans else None
            bold = True if (c_idx == 0 or is_filled_ans or (r_idx == 0 and c_idx in [1, 2])) else False
            render_rich_text(p, val, base_size=10.5, base_bold=bold, base_color=color)

def generate_docx(is_teacher=False):
    doc = Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line=DATA_EXAM["school_line"],
        exam_title=DATA_EXAM["exam_title"],
        subject_line=DATA_EXAM["subject_line"],
        num_pages=DATA_EXAM["num_pages"],
        duration=DATA_EXAM["duration"],
        exam_code=DATA_EXAM["exam_code"],
        atomic_masses=DATA_EXAM["atomic_masses"],
        is_teacher=is_teacher
    )

    # Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Bản Giáo Viên
    if is_teacher:
        add_quick_answer_table_2101(doc)

    # I. TRẮC NGHIỆM (7,0 ĐIỂM)
    add_section_header(
        doc,
        title="I. TRẮC NGHIỆM (7,0 ĐIỂM)",
        subtitle="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. (12 câu - 3,0 điểm)\nMỗi câu hỏi học sinh chỉ chọn một phương án. Mỗi phương án đúng được 0,25 điểm."
    )

    for item in DATA_EXAM["part1"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        # Chèn ảnh nếu có (như Câu 5)
        if "image" in item and os.path.exists(item["image"]):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(2)
            r_img = p_img.add_run()
            r_img.add_picture(item["image"], width=Cm(item.get("image_width", 7.5)))

        if "sub_prompt" in item:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(1)
            p_sub.paragraph_format.space_after = Pt(1)
            render_question_prompt(p_sub, item["sub_prompt"], base_size=10.5)

        render_options_no_table(doc, item["options"], is_teacher=is_teacher)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN II. Câu trắc nghiệm đúng sai. (2 câu - 2,0 điểm)
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai. (2 câu - 2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    for item in DATA_EXAM["part2"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        # Chèn ảnh nếu có (như phổ khối lượng hay sơ đồ nguyên tử)
        if "image" in item and os.path.exists(item["image"]):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(2)
            r_img = p_img.add_run()
            r_img.add_picture(item["image"], width=Cm(item.get("image_width", 7.5)))

        for stmt, ans in item["statements"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt, ans=ans, is_teacher=is_teacher, base_size=10.5)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN III. Câu trắc nghiệm trả lời ngắn. (4 câu - 2,0 điểm)
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm trả lời ngắn. (4 câu - 2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4. Thí sinh điền kết quả vào ô trả lời tương ứng."
    )

    for item in DATA_EXAM["part3"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(2)
        r_lbl = p_ans.add_run("Trả lời: ")
        set_run_style(r_lbl, bold=True, color=RGBColor(0, 0, 0))

        if is_teacher:
            r_val = p_ans.add_run(f"{item['answer']}")
            set_run_style(r_val, bold=True, color=COLOR_CORRECT)
        else:
            r_line = p_ans.add_run("..................................................")
            set_run_style(r_line, color=RGBColor(150, 150, 150))

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # II. TỰ LUẬN (3,0 ĐIỂM)
    add_section_header(
        doc,
        title="II. TỰ LUẬN (3,0 ĐIỂM)",
        subtitle="PHẦN IV. Tự luận. (3 câu - 3,0 điểm)\nHọc sinh làm bài theo hình thức tự luận."
    )

    for item in DATA_EXAM["part4"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        # Nếu có bảng ở đề bài
        if item.get("table_type") == "neon":
            render_table_neon(doc)
            for sub_p in item.get("sub_prompts", []):
                p_sub = doc.add_paragraph()
                p_sub.paragraph_format.left_indent = Cm(0.5)
                p_sub.paragraph_format.space_before = Pt(1)
                p_sub.paragraph_format.space_after = Pt(1)
                render_question_prompt(p_sub, sub_p, base_size=10.5)
        elif item.get("table_type") == "atom_properties":
            render_table_atom_props(doc, is_teacher=is_teacher)

        if not is_teacher:
            add_student_writing_lines(doc, num_lines=item.get("num_lines", 6))

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    filename = f"De_Kiem_Tra_GK1_Hoa_10_GiaHoi_Ma2101_{'GiaoVien' if is_teacher else 'HocSinh'}.docx"
    filepath = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(filepath)
    print(f"Đã tạo file Word: {filepath}")
    return filepath

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
