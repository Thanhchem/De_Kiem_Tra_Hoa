# -*- coding: utf-8 -*-
"""
Tạo file BẢNG ĐÁP ÁN RÚT GỌN chuẩn đẹp 100% theo mẫu ảnh người dùng gửi:
- Bảng viền xanh (#4CAF50), nền tiêu đề xanh nhạt (#E8F5E9)
- Đáp án in đậm màu đỏ (#C00000)
- Lệnh a, b, c, d có chữ ĐÚNG in đậm đỏ, SAI in bình thường
- Tóm tắt đáp số tự luận chi tiết, rõ ràng
Áp dụng cho cả 4 đề kiểm tra Hóa 10.
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

sys.path.append(os.path.abspath('scripts'))
from form_de_kiem_tra import (
    setup_header_footer,
    set_run_style,
    render_rich_text,
    COLOR_CORRECT,
    COLOR_PRIMARY
)

COLOR_TITLE_RED = RGBColor(192, 0, 0)       # #C00000
COLOR_SECTION_GREEN = RGBColor(46, 125, 50) # #2E7D32
COLOR_BORDER_HEX = "4CAF50"                 # #4CAF50
COLOR_HEADER_BG_HEX = "E8F5E9"              # #E8F5E9
COLOR_TEXT_DARK = RGBColor(33, 33, 33)

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_table_borders(tbl, color="4CAF50", sz="6"):
    tblPr = tbl._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            </w:tblBorders>
        ''')
        tblPr[0].append(borders)

def set_cell_padding(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def add_styled_answer_block(doc, exam_data):
    """
    Thêm toàn bộ khối BẢNG ĐÁP ÁN RÚT GỌN chuẩn 4 phần theo mẫu ảnh:
    exam_data: dict chứa thông tin chi tiết từng phần.
    """
    # 0. Tên đề / Tiêu đề chính
    p_main_title = doc.add_paragraph()
    p_main_title.paragraph_format.space_before = Pt(6)
    p_main_title.paragraph_format.space_after = Pt(2)
    p_main_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_main_title = p_main_title.add_run(f"BẢNG ĐÁP ÁN RÚT GỌN - {exam_data['exam_name'].upper()}")
    set_run_style(r_main_title, font_size=13, bold=True, color=COLOR_TITLE_RED)

    # 1. PHẦN I
    p_s1 = doc.add_paragraph()
    p_s1.paragraph_format.space_before = Pt(4)
    p_s1.paragraph_format.space_after = Pt(2)
    r_s1 = p_s1.add_run(exam_data['p1_title'])
    set_run_style(r_s1, font_size=10.5, bold=True, color=COLOR_SECTION_GREEN)

    p1_answers = exam_data['p1_answers']  # list of tuples (q_num, ans)
    # Tùy số lượng câu: nếu <= 10 câu 1 hàng; nếu 12 câu -> 2 hàng x 6; nếu 16 câu -> 2 hàng x 8
    n_p1 = len(p1_answers)
    if n_p1 <= 12:
        cols_per_row = (n_p1 + 1) // 2 if n_p1 > 9 else n_p1
        chunk1 = p1_answers[:cols_per_row]
        chunk2 = p1_answers[cols_per_row:]
        rows_data = [chunk1]
        if chunk2:
            rows_data.append(chunk2)
    else:
        cols_per_row = (n_p1 + 1) // 2
        chunk1 = p1_answers[:cols_per_row]
        chunk2 = p1_answers[cols_per_row:]
        rows_data = [chunk1, chunk2]

    # Tạo bảng Phần I
    # Mỗi hàng câu hỏi có: hàng 1 là Câu, hàng 2 là Đ/A
    total_table_rows = len(rows_data) * 2
    max_cols = max(len(c) for c in rows_data) + 1  # +1 cho cột tiêu đề 'Câu' / 'Đ/A'

    tbl1 = doc.add_table(rows=total_table_rows, cols=max_cols)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl1, color=COLOR_BORDER_HEX, sz="6")

    for r_idx, chunk in enumerate(rows_data):
        row_q = r_idx * 2
        row_a = r_idx * 2 + 1

        # Cột đầu
        c_q0 = tbl1.cell(row_q, 0)
        c_q0.text = "Câu"
        c_q0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_style(c_q0.paragraphs[0].runs[0], font_size=9.5, bold=True, color=COLOR_SECTION_GREEN)
        set_cell_background(c_q0, COLOR_HEADER_BG_HEX)
        set_cell_padding(c_q0)

        c_a0 = tbl1.cell(row_a, 0)
        c_a0.text = "Đ/A"
        c_a0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_style(c_a0.paragraphs[0].runs[0], font_size=9.5, bold=True, color=COLOR_SECTION_GREEN)
        set_cell_padding(c_a0)

        for c_idx, (q_num, ans) in enumerate(chunk):
            col = c_idx + 1
            cell_q = tbl1.cell(row_q, col)
            cell_q.text = str(q_num)
            cell_q.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run_style(cell_q.paragraphs[0].runs[0], font_size=9.5, bold=True, color=COLOR_SECTION_GREEN)
            set_cell_background(cell_q, COLOR_HEADER_BG_HEX)
            set_cell_padding(cell_q)

            cell_a = tbl1.cell(row_a, col)
            cell_a.text = ans
            cell_a.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run_style(cell_a.paragraphs[0].runs[0], font_size=10, bold=True, color=COLOR_TITLE_RED)
            set_cell_padding(cell_a)

    # 2. PHẦN II: ĐÚNG / SAI
    p_s2 = doc.add_paragraph()
    p_s2.paragraph_format.space_before = Pt(4)
    p_s2.paragraph_format.space_after = Pt(2)
    r_s2 = p_s2.add_run(exam_data['p2_title'])
    set_run_style(r_s2, font_size=10.5, bold=True, color=COLOR_SECTION_GREEN)

    p2_data = exam_data['p2_data']  # list of dict: {"q": "Câu 1", "content": "...", "a": "ĐÚNG", "b": "SAI", ...}
    tbl2 = doc.add_table(rows=1 + len(p2_data), cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl2, color=COLOR_BORDER_HEX, sz="6")

    # Header row
    headers_p2 = ["Câu hỏi", "Nội dung câu hỏi", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"]
    widths_p2 = [Cm(1.8), Cm(8.2), Cm(2.0), Cm(2.0), Cm(2.0), Cm(2.0)]
    for c_i, h_text in enumerate(headers_p2):
        cell = tbl2.cell(0, c_i)
        cell.width = widths_p2[c_i]
        cell.text = h_text
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_style(cell.paragraphs[0].runs[0], font_size=9.5, bold=True, color=COLOR_SECTION_GREEN)
        set_cell_background(cell, COLOR_HEADER_BG_HEX)
        set_cell_padding(cell)

    for r_i, q_item in enumerate(p2_data):
        row = r_i + 1
        # Cột câu
        c_q = tbl2.cell(row, 0)
        c_q.width = widths_p2[0]
        c_q.text = q_item["q"]
        c_q.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_style(c_q.paragraphs[0].runs[0], font_size=9.5, bold=True)
        set_cell_padding(c_q)

        # Cột nội dung
        c_desc = tbl2.cell(row, 1)
        c_desc.width = widths_p2[1]
        c_desc.text = q_item["content"]
        c_desc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_run_style(c_desc.paragraphs[0].runs[0], font_size=9.5)
        set_cell_padding(c_desc)

        # 4 lệnh a, b, c, d
        for opt_idx, opt_key in enumerate(["a", "b", "c", "d"]):
            c_opt = tbl2.cell(row, 2 + opt_idx)
            c_opt.width = widths_p2[2 + opt_idx]
            val = q_item[opt_key].strip()
            c_opt.text = val
            c_opt.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if val.upper() in ["ĐÚNG", "Đ"]:
                set_run_style(c_opt.paragraphs[0].runs[0], font_size=9.5, bold=True, color=COLOR_TITLE_RED)
            else:
                set_run_style(c_opt.paragraphs[0].runs[0], font_size=9.5, bold=False, color=COLOR_TEXT_DARK)
            set_cell_padding(c_opt)

    # 3. PHẦN III: TRẢ LỜI NGẮN
    p_s3 = doc.add_paragraph()
    p_s3.paragraph_format.space_before = Pt(4)
    p_s3.paragraph_format.space_after = Pt(2)
    r_s3 = p_s3.add_run(exam_data['p3_title'])
    set_run_style(r_s3, font_size=10.5, bold=True, color=COLOR_SECTION_GREEN)

    p3_answers = exam_data['p3_answers']  # list of (col_header, ans_str)
    tbl3 = doc.add_table(rows=2, cols=len(p3_answers) + 1)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl3, color=COLOR_BORDER_HEX, sz="6")

    # Header cột 0
    c_q0 = tbl3.cell(0, 0)
    c_q0.text = "Câu"
    c_q0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_style(c_q0.paragraphs[0].runs[0], font_size=9.5, bold=True, color=COLOR_SECTION_GREEN)
    set_cell_background(c_q0, COLOR_HEADER_BG_HEX)
    set_cell_padding(c_q0)

    c_a0 = tbl3.cell(1, 0)
    c_a0.text = "Đáp án"
    c_a0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_style(c_a0.paragraphs[0].runs[0], font_size=9.5, bold=True, color=COLOR_SECTION_GREEN)
    set_cell_padding(c_a0)

    for idx, (col_head, ans_val) in enumerate(p3_answers):
        col = idx + 1
        c_head = tbl3.cell(0, col)
        c_head.text = col_head
        c_head.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_style(c_head.paragraphs[0].runs[0], font_size=9.5, bold=True, color=COLOR_SECTION_GREEN)
        set_cell_background(c_head, COLOR_HEADER_BG_HEX)
        set_cell_padding(c_head)

        c_val = tbl3.cell(1, col)
        c_val.text = str(ans_val)
        c_val.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_style(c_val.paragraphs[0].runs[0], font_size=10, bold=True, color=COLOR_TITLE_RED)
        set_cell_padding(c_val)

    # 4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN
    p_s4 = doc.add_paragraph()
    p_s4.paragraph_format.space_before = Pt(4)
    p_s4.paragraph_format.space_after = Pt(2)
    r_s4 = p_s4.add_run(exam_data['p4_title'])
    set_run_style(r_s4, font_size=10.5, bold=True, color=COLOR_SECTION_GREEN)

    p4_data = exam_data['p4_data']  # list of {"q": "Câu 1", "score": "1,0 đ", "bullets": [...]}
    tbl4 = doc.add_table(rows=1 + len(p4_data), cols=3)
    tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl4, color=COLOR_BORDER_HEX, sz="6")

    headers_p4 = ["Câu", "Điểm", "Đáp số cốt lõi & Hướng dẫn chấm"]
    widths_p4 = [Cm(1.8), Cm(1.5), Cm(14.7)]
    for c_i, h_text in enumerate(headers_p4):
        cell = tbl4.cell(0, c_i)
        cell.width = widths_p4[c_i]
        cell.text = h_text
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_style(cell.paragraphs[0].runs[0], font_size=9.5, bold=True, color=COLOR_SECTION_GREEN)
        set_cell_background(cell, COLOR_HEADER_BG_HEX)
        set_cell_padding(cell)

    for r_i, q_item in enumerate(p4_data):
        row = r_i + 1
        c_q = tbl4.cell(row, 0)
        c_q.width = widths_p4[0]
        c_q.text = q_item["q"]
        c_q.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_style(c_q.paragraphs[0].runs[0], font_size=9.5, bold=True)
        set_cell_padding(c_q)

        c_sc = tbl4.cell(row, 1)
        c_sc.width = widths_p4[1]
        c_sc.text = q_item["score"]
        c_sc.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_style(c_sc.paragraphs[0].runs[0], font_size=9.5, bold=True, color=COLOR_SECTION_GREEN)
        set_cell_padding(c_sc)

        c_ans = tbl4.cell(row, 2)
        c_ans.width = widths_p4[2]
        c_ans.paragraphs[0].text = ""  # clear default
        for b_idx, bullet in enumerate(q_item["bullets"]):
            p = c_ans.paragraphs[0] if b_idx == 0 else c_ans.add_paragraph()
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            render_rich_text(p, f"• {bullet}", base_size=9.5)
        set_cell_padding(c_ans)

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(4)
    p_spacer.paragraph_format.space_after = Pt(4)

# DỮ LIỆU CỦA 4 ĐỀ THI
EXAMS_DATA = [
    {
        "exam_name": "Đề 1: THPT Nguyễn Đình Chiểu - Mã đề 1002",
        "p1_title": "1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 ĐIỂM - MỖI CÂU 0,25 Đ)",
        "p1_answers": [(1, "C"), (2, "A"), (3, "A"), (4, "D"), (5, "C"), (6, "D"), 
                       (7, "A"), (8, "B"), (9, "B"), (10, "B"), (11, "C"), (12, "D")],
        "p2_title": "2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM - MỖI CÂU 1,0 Đ)",
        "p2_data": [
            {
                "q": "Câu 1",
                "content": "Hợp chất Potassium iodide (KI), thuốc long đờm, bảo vệ tuyến giáp",
                "a": "ĐÚNG", "b": "ĐÚNG", "c": "SAI", "d": "ĐÚNG"
            },
            {
                "q": "Câu 2",
                "content": "Cấu hình electron lớp ngoài cùng 3s^x và 3p^y (x + y = 8), Y là khí hiếm",
                "a": "ĐÚNG", "b": "SAI", "c": "ĐÚNG", "d": "SAI"
            }
        ],
        "p3_title": "3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM - MỖI CÂU 0,5 Đ)",
        "p3_answers": [("Câu 1", "6 (hoặc 15)"), ("Câu 2", "4"), ("Câu 3", "14"), ("Câu 4", "17")],
        "p4_title": "4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN (3,0 ĐIỂM)",
        "p4_data": [
            {
                "q": "Câu 1",
                "score": "1,0 đ",
                "bullets": [
                    "Hệ phương trình: 2Z + n = 40 và 2Z - n = 12 ➔ Z = 13, n = 14.",
                    "a) Số hạt mỗi loại: *p = 13, e = 13, n = 14*.",
                    "b) Số khối: *A = Z + n = 13 + 14 = 27* (Nguyên tố Aluminium - Al)."
                ]
            },
            {
                "q": "Câu 2",
                "score": "1,0 đ",
                "bullets": [
                    "Phân lớp 3s hơn kém nhau 1e và x + y = 6 ➔ x = 1, y = 5.",
                    "a) Cấu hình electron: *X (Z=11): [Ne]3s^1^*; *Y (Z=17): [Ne]3s^2^3p^5^*.",
                    "b) Xu hướng tạo ion: X nhường 1e ➔ *Na^+^ ([Ne])*; Y nhận 1e ➔ *Cl^-^ ([Ar])*."
                ]
            },
            {
                "q": "Câu 3",
                "score": "1,0 đ",
                "bullets": [
                    "Tổng số hiệu: Z~X~ + Z~Y~ + Z~Z~ = 39 và Z~Y~ = (Z~X~ + Z~Z~)/2 ➔ Z~Y~ = 13 (Al).",
                    "Cùng chu kì 3, không tác dụng với H~2~O ➔ *X là Mg (Z = 12)*, *Y là Al (Z = 13)*, *Z là Si (Z = 14)*.",
                    "a) Vị trí: *Mg (Ô 12, CK 3, nhóm IIA)*; *Al (Ô 13, CK 3, nhóm IIIA)*; *Si (Ô 14, CK 3, nhóm IVA)*.",
                    "b) Công thức hydroxide: *Mg(OH)~2~, Al(OH)~3~, H~2~SiO~3~*.",
                    "So sánh tính base giảm dần: *Mg(OH)~2~ > Al(OH)~3~ > H~2~SiO~3~*."
                ]
            }
        ]
    },
    {
        "exam_name": "Đề 2: THPT Hoá Châu - Mã đề 132",
        "p1_title": "1. PHẦN 1: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 ĐIỂM - MỖI CÂU 0,25 Đ)",
        "p1_answers": [(1, "B"), (2, "C"), (3, "D"), (4, "B"), (5, "C"), (6, "D"), 
                       (7, "C"), (8, "B"), (9, "D"), (10, "D"), (11, "A"), (12, "A")],
        "p2_title": "2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM - MỖI CÂU 1,0 Đ)",
        "p2_data": [
            {
                "q": "Câu 1",
                "content": "Đồng vị của Sulfur, tính chất oxide, Fe tác dụng với H~2~SO~4~",
                "a": "SAI", "b": "ĐÚNG", "c": "SAI", "d": "SAI"
            },
            {
                "q": "Câu 2",
                "content": "Sodium chloride (NaCl), liên kết hóa học, điều chế NaOH",
                "a": "ĐÚNG", "b": "SAI", "c": "ĐÚNG", "d": "SAI"
            }
        ],
        "p3_title": "3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM - MỖI CÂU 0,25 Đ)",
        "p3_answers": [("Câu 1", "2"), ("Câu 2", "9"), ("Câu 3", "108"), ("Câu 4", "8"),
                       ("Câu 5", "3"), ("Câu 6", "14"), ("Câu 7", "6"), ("Câu 8", "16")],
        "p4_title": "4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN (3,0 ĐIỂM)",
        "p4_data": [
            {
                "q": "Câu 1",
                "score": "1,5 đ",
                "bullets": [
                    "a) Công thức Lewis & CTCT: NH~3~ (H-N̈(H)-H, H-N(H)-H); H~2~S (H-S̈̈-H, H-S-H); Cl~2~ (:C̈l̈-C̈l̈:, Cl-Cl); CO (:C≡O:, C≡O).",
                    "b) Loại liên kết (theo Δχ): *NH~3~ (CHT phân cực)*; *H~2~S (CHT không phân cực)*; *Cl~2~ (CHT không phân cực)*; *CO (CHT phân cực + 1 liên kết phối trí)*."
                ]
            },
            {
                "q": "Câu 2",
                "score": "1,5 đ",
                "bullets": [
                    "n~H2~ = 2,2311 / 24,79 = 0,09 mol ➔ n~M~ = 0,18/n ➔ M = 20n. Với n = 2 ➔ *M = 40 (Calcium - Ca)*.",
                    "a) Vị trí Ca (Z = 20): *Ô số 20, chu kì 4, nhóm IIA*.",
                    "b) So sánh tính kim loại: *K > Ca > Mg*."
                ]
            }
        ]
    },
    {
        "exam_name": "Đề 3: THPT Đặng Huy Trứ - Mã đề 104",
        "p1_title": "1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (4,0 ĐIỂM - MỖI CÂU 0,25 Đ)",
        "p1_answers": [(1, "C"), (2, "D"), (3, "A"), (4, "D"), (5, "C"), (6, "C"), (7, "B"), (8, "A"),
                       (9, "A"), (10, "B"), (11, "A"), (12, "A"), (13, "A"), (14, "B"), (15, "B"), (16, "C")],
        "p2_title": "2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM - MỖI CÂU 1,0 Đ)",
        "p2_data": [
            {
                "q": "Câu 1",
                "content": "Phân tử nước (H~2~O), liên kết CHT, độ phân cực, cặp e tự do",
                "a": "ĐÚNG", "b": "SAI", "c": "SAI", "d": "ĐÚNG"
            },
            {
                "q": "Câu 2",
                "content": "Nhóm halogen (F, Cl, Br), tính phi kim, hợp chất HX",
                "a": "ĐÚNG", "b": "SAI", "c": "ĐÚNG", "d": "SAI"
            }
        ],
        "p3_title": "3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (1,0 ĐIỂM - MỖI CÂU 0,25 Đ)",
        "p3_answers": [("Câu 1", "8"), ("Câu 2", "3"), ("Câu 3", "5"), ("Câu 4", "4")],
        "p4_title": "4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN (3,0 ĐIỂM)",
        "p4_data": [
            {
                "q": "Câu 1",
                "score": "1,25 đ",
                "bullets": [
                    "Cùng nhóm A liên tiếp: Z~X~ + Z~Y~ = 32 và Z~Y~ - Z~X~ = 8 ➔ *X là Mg (Z = 12)*, *Y là Ca (Z = 20)*.",
                    "a) So sánh bán kính: Ca có 4 lớp e > Mg có 3 lớp e ➔ *r~Ca~ > r~Mg~*.",
                    "b) n~H2~ = 18,5925 / 24,79 = 0,75 mol. Hệ: 24x + 40y = 22 và x + y = 0,75 ➔ x = 0,5 mol; y = 0,25 mol.",
                    "Khối lượng từng kim loại: *m~Mg~ = 12 gam*; *m~Ca~ = 10 gam*."
                ]
            },
            {
                "q": "Câu 2",
                "score": "1,0 đ",
                "bullets": [
                    "X: [Ne]3s^1^ ➔ *Z = 11 (Sodium - Na)*.",
                    "a) Tính chất cơ bản: *Tính kim loại mạnh*. Oxide cao nhất: *Na~2~O* (tính base); Hydroxide: *NaOH* (base mạnh).",
                    "b) PTHH: Na~2~O + H~2~O ➔ 2NaOH; Na~2~O + 2HCl ➔ 2NaCl + H~2~O; NaOH + HCl ➔ NaCl + H~2~O."
                ]
            },
            {
                "q": "Câu 3",
                "score": "0,75 đ",
                "bullets": [
                    "a) %m_R = [2R / (2R + 80)] × 100% = 25,93% ➔ *R = 14 (Nitrogen - N)*. Hợp chất khí A là NH~3~.",
                    "b) Liên kết giữa N và H là *liên kết cộng hóa trị phân cực* (Δχ = 0,84).",
                    "Công thức Lewis: H - N̈(H) - H; Công thức cấu tạo: H - N(H) - H."
                ]
            }
        ]
    },
    {
        "exam_name": "Đề 4: THPT Nguyễn Chí Thanh - Mã đề 1104",
        "p1_title": "1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (4,0 ĐIỂM - MỖI CÂU 0,25 Đ)",
        "p1_answers": [(1, "B"), (2, "C"), (3, "A"), (4, "D"), (5, "B"), (6, "C"), (7, "A"), (8, "C"),
                       (9, "D"), (10, "A"), (11, "C"), (12, "D"), (13, "B"), (14, "A"), (15, "B"), (16, "C")],
        "p2_title": "2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM - MỖI CÂU 1,0 Đ)",
        "p2_data": [
            {
                "q": "Câu 1",
                "content": "Phân tử ammonia (NH~3~), hạt mang điện, liên kết CHT, LK hydrogen",
                "a": "SAI", "b": "SAI", "c": "ĐÚNG", "d": "ĐÚNG"
            },
            {
                "q": "Câu 2",
                "content": "Ba nguyên tố X, Y, Z kế tiếp trong chu kì (Z~X~ + Z~Y~ + Z~Z~ = 36)",
                "a": "ĐÚNG", "b": "SAI", "c": "ĐÚNG", "d": "SAI"
            }
        ],
        "p3_title": "3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (1,0 ĐIỂM - MỖI CÂU 0,25 Đ)",
        "p3_answers": [("Câu 1", "3"), ("Câu 2", "3"), ("Câu 3", "98 (hoặc 50)"), ("Câu 4", "9")],
        "p4_title": "4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN (3,0 ĐIỂM)",
        "p4_data": [
            {
                "q": "Câu 1",
                "score": "1,0 đ",
                "bullets": [
                    "a) Cấu hình electron của S (Z = 16): *1s^2^2s^2^2p^6^3s^2^3p^4^*. Vị trí: *Ô số 16, chu kì 3, nhóm VIA*.",
                    "b) Phân tử H~2~S: CT Lewis: *H - S̈̈ - H*; CTCT: *H - S - H*."
                ]
            },
            {
                "q": "Câu 2",
                "score": "1,0 đ",
                "bullets": [
                    "Hợp chất khí XH~3~: %m_X = [M_X / (M_X + 3)] × 100% = 91,18% ➔ *M_X = 31 (Phosphorus - P)*.",
                    "Oxide cao nhất là P~2~O~5~ ➔ %m_P = [62 / 142] × 100% ≈ *43,7%*."
                ]
            },
            {
                "q": "Câu 3",
                "score": "1,0 đ",
                "bullets": [
                    "a) Phần trăm đồng vị: x + y = 100 và 63x + 65y = 6354 ➔ *%^63^Cu = 73%*; *%^65^Cu = 27%*.",
                    "b) M(Cu~2~S) = 2 × 63,54 + 32 = 159,08 g/mol; m(^65^Cu) = 2 × 0,27 × 65 = 35,1 g.",
                    "Thành phần phần trăm: *%m(^65^Cu) = (35,1 / 159,08) × 100% ≈ 22,06%*."
                ]
            }
        ]
    }
]

def generate_master_answer_file():
    """Tạo 1 file Word độc lập chứa toàn bộ BẢNG ĐÁP ÁN RÚT GỌN của cả 4 đề."""
    doc = docx.Document()
    setup_header_footer(doc)

    p_header = doc.add_paragraph()
    p_header.paragraph_format.space_before = Pt(0)
    p_header.paragraph_format.space_after = Pt(4)
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_header.add_run("TỔNG HỢP BẢNG ĐÁP ÁN RÚT GỌN CÁC ĐỀ KIỂM TRA HÓA HỌC 10\nCUỐI HỌC KÌ I (NĂM HỌC 2025 - 2026)")
    set_run_style(r1, font_size=14, bold=True, color=COLOR_TITLE_RED)

    for idx, exam_data in enumerate(EXAMS_DATA):
        if idx > 0:
            doc.add_page_break()
        add_styled_answer_block(doc, exam_data)

    out_path = os.path.join("San_Pham", "BANG_DAP_AN_RUT_GON_CAC_DE_HOA_10.docx")
    doc.save(out_path)
    print(f"Đã tạo file tổng hợp BẢNG ĐÁP ÁN: {out_path}")

if __name__ == "__main__":
    generate_master_answer_file()
