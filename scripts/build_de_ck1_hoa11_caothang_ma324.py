# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Học Kì I - Hóa học 11 - Trường THPT Cao Thắng (TP Huế)
Mã đề: 324 - Năm học 2024 - 2025
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
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_de324")
os.makedirs(DIR_SAN_PHAM, exist_ok=True)

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

def add_quick_answer_table_324(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Giáo viên ở đầu tài liệu Mã đề 324.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 324)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,5 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=15)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_p1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14"]
    answers_p1 = ["Đ/A", "B", "A", "B", "A", "D", "B", "D", "D", "B", "A", "A", "C", "C", "B"]

    w1_c0 = Cm(1.6)
    w1_c = Cm(1.17)

    for r_idx, r_data in enumerate([headers_p1, answers_p1]):
        row = tbl1.rows[r_idx]
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = w1_c0 if c_idx == 0 else w1_c
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
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
    render_rich_text(p_p2, "*2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (3,0 ĐIỂM - MỖI CÂU 1,0 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=4, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl2.autofit = False

    tbl2_data = [
        ["Câu hỏi", "Nội dung câu hỏi", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Hợp chất hữu cơ X (CH~2~O, M=60)", "ĐÚNG", "ĐÚNG", "SAI", "SAI"],
        ["Câu 2", "Sản xuất Sulfuric acid công nghiệp", "SAI", "SAI", "SAI", "ĐÚNG"],
        ["Câu 3", "Tách β-carotene từ nước ép cà rốt", "ĐÚNG", "ĐÚNG", "ĐÚNG", "ĐÚNG"]
    ]

    w2_widths = [Cm(2.0), Cm(6.0), Cm(2.5), Cm(2.5), Cm(2.5), Cm(2.5)]

    for r_idx, row_vals in enumerate(tbl2_data):
        row = tbl2.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = w2_widths[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
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
    render_rich_text(p_p3, "*3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (1,5 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=7)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False

    tbl3_headers = ["Câu", "Câu 1", "Câu 2", "Câu 3", "Câu 4", "Câu 5", "Câu 6"]
    tbl3_answers = ["Đáp án", "0,87", "31", "3", "6 (hoặc 5)", "2", "3"]

    w3_c0 = Cm(2.1)
    w3_c = Cm(2.65)

    for r_idx, row_vals in enumerate([tbl3_headers, tbl3_answers]):
        row = tbl3.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = w3_c0 if c_idx == 0 else w3_c
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
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
    render_rich_text(p_tl, "*4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN (2,0 ĐIỂM - MỖI CÂU 1,0 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl4 = doc.add_table(rows=3, cols=3)
    tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl4.autofit = False

    tbl4_data = [
        ("Câu", "Điểm", "Đáp số cốt lõi & Hướng dẫn chấm"),
        ("Câu 1", "1,0 đ", "• n[Ca(H~2~PO~4~)~2~] = 8 000 kmol ⇒ n(H~2~SO~4~) = 2 × 8 000 = 16 000 kmol.\n• m(H~2~SO~4~) = 1 568 tấn ⇒ m(dd) = 2 240 tấn ⇒ V(dd H~2~SO~4~ 70%) = 1 400 m^3^."),
        ("Câu 2", "1,0 đ", "• a) CTPT của X: C~4~H~10~O (M = 74 amu).\n• b) X là alcohol (vùng 3500 - 3200 cm^-1^), có 4 đồng phân cấu tạo: butan-1-ol, butan-2-ol, 2-methylpropan-1-ol, 2-methylpropan-2-ol.")
    ]

    w4_widths = [Cm(2.0), Cm(1.5), Cm(14.5)]

    for r_idx, (c1, c2, c3) in enumerate(tbl4_data):
        row = tbl4.rows[r_idx]
        for c_idx, val in enumerate([c1, c2, c3]):
            cell = row.cells[c_idx]
            cell.width = w4_widths[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
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


# Dữ liệu Đề thi mã 324
DATA_EXAM_324 = {
    "school_line": "SỞ GIÁO DỤC & ĐÀO TẠO - TRƯỜNG THPT CAO THẮNG",
    "exam_title": "ĐỀ KIỂM TRA CUỐI KÌ 1 (2024-2025)",
    "subject_line": "MÔN: HÓA HỌC 11",
    "duration": "45 phút (không kể thời gian giao đề)",
    "num_pages": "04",
    "exam_code": "324",
    "atomic_masses": "H = 1; O = 16; C = 12; N = 14; Ca = 40; Fe = 56; S = 32; P = 31",

    # A. TRẮC NGHIỆM KHÁCH QUAN (8,0 ĐIỂM)
    # PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,5 điểm - 14 câu)
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1. Ứng dụng nào sau đây không phải của sulfur dioxide?",
            "options": [
                "A. Sản xuất sulfuric acid.",
                "*B. Sản xuất phân bón.*",
                "C. Chống nấm mốc cho lương thực, thực phẩm.",
                "D. Tẩy trắng giấy, bột giấy."
            ],
            "solution": "Sulfur dioxide (SO~2~) được dùng làm chất trung gian quan trọng trong sản xuất sulfuric acid, làm chất tẩy trắng giấy và bột giấy, chống nấm mốc cho lương thực, thực phẩm và quả khô (nhờ tính kháng khuẩn, diệt vi nấm). Bản thân SO~2~ không được sử dụng trực tiếp làm phân bón (phân bón thường là các muối amoni, nitrate, phosphate hoặc kali).\n*Đáp án đúng: B.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Phân tử ammonia có dạng hình học nào sau đây?",
            "options": [
                "*A. Chóp tam giác.*",
                "B. Tam giác đều.",
                "C. Chữ T.",
                "D. Chóp tứ giác."
            ],
            "solution": "Trong phân tử ammonia (NH~3~), nguyên tử trung tâm nitrogen ở trạng thái lai hóa sp^3^ liên kết với 3 nguyên tử hydrogen và còn lại 1 cặp electron chưa liên kết (cặp electron tự do). Lực đẩy của cặp electron tự do làm biến dạng góc liên kết H-N-H xuống khoảng 107°, tạo nên cấu trúc hình học dạng chóp tam giác (trigonal pyramidal).\n*Đáp án đúng: A.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Công thức phân tử cho biết thông tin nào sau đây về phân tử hợp chất hữu cơ?",
            "options": [
                "A. Thành phần nguyên tố và tỉ lệ số lượng nguyên tử của mỗi nguyên tố.",
                "*B. Thành phần nguyên tố và số lượng nguyên tử của mỗi nguyên tố.*",
                "C. Số lượng nguyên tử mỗi nguyên tố và trật tự liên kết giữa các nguyên tử.",
                "D. Tỉ lệ số lượng nguyên tử của mỗi nguyên tố và trật tự liên kết giữa các nguyên tử."
            ],
            "solution": "Theo định nghĩa chuẩn SGK Hóa học 11:\n- Công thức đơn giản nhất: Cho biết thành phần nguyên tố và tỉ lệ tối giản số lượng nguyên tử của các nguyên tố trong phân tử.\n- Công thức phân tử: Cho biết thành phần nguyên tố và số lượng nguyên tử của mỗi nguyên tố trong phân tử.\n- Công thức cấu tạo: Cho biết thành phần nguyên tố, số lượng nguyên tử và trật tự liên kết giữa các nguyên tử trong phân tử.\n*Đáp án đúng: B.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Tại nhiệt độ không đổi, ở trạng thái cân bằng,",
            "options": [
                "*A. nồng độ của các chất trong phản ứng không thay đổi.*",
                "B. phản ứng hoá học không xảy ra.",
                "C. tốc độ phản ứng hoá học xảy ra chậm dần.",
                "D. nồng độ của các chất trong phản ứng vẫn liên tục thay đổi."
            ],
            "solution": "Cân bằng hóa học là một cân bằng động: ở trạng thái cân bằng, phản ứng thuận và phản ứng nghịch vẫn liên tục diễn ra với tốc độ bằng nhau (v~thuận~ = v~nghịch~ > 0), do đó nồng độ của các chất tham gia và chất tạo thành trong hệ phản ứng giữ nguyên không đổi theo thời gian.\n*Đáp án đúng: A.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5. Nitric acid là một acid có tính oxi hóa mạnh. Do đó, nitric acid oxi hóa được hầu hết kim loại, trừ vàng (Au), platinum (Pt),... Tuy nhiên, Au và Pt có thể bị hòa tan trong nước cường toan, là hỗn hợp nitric acid đặc và hydrochloric acid đặc có tỉ lệ thể tích là",
            "options": [
                "A. 2:3.",
                "B. 3:1.",
                "C. 1:1.",
                "*D. 1:3.*"
            ],
            "solution": "Nước cường toan (còn gọi là nước cường thủy hay aqua regia) là hỗn hợp gồm 1 thể tích HNO~3~ đặc và 3 thể tích HCl đặc (tỉ lệ V(HNO~3~) : V(HCl) = 1 : 3). Hỗn hợp này có tính oxi hóa cực mạnh, có khả năng hòa tan cả vàng (Au) và bạch kim (Pt).\n*Đáp án đúng: D.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6. Tính chất nào sau đây không phải tính chất vật lí của sulfur?",
            "options": [
                "A. Thể rắn ở điều kiện thường.",
                "*B. Không tan trong benzene.*",
                "C. Màu vàng ở điều kiện thường.",
                "D. Không tan trong nước."
            ],
            "solution": "Ở điều kiện thường, sulfur (lưu huỳnh) là chất rắn màu vàng, không tan trong nước, nhưng tan tốt trong các dung môi hữu cơ không phân cực hoặc ít phân cực như benzene, carbon disulfide (CS~2~), toluene. Do đó phát biểu \"Không tan trong benzene\" là không đúng.\n*Đáp án đúng: B.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7. \"Đất phèn\" là một khái niệm dân gian để chỉ loại đất chứa nhiều ion sulfate, có pH < 7. Môi trường của đất phèn là",
            "options": [
                "A. môi trường trung hòa.",
                "B. môi trường base.",
                "C. môi trường trung tính.",
                "*D. môi trường acid.*"
            ],
            "solution": "Theo thang đo pH chuẩn: dung dịch hoặc đất có pH < 7 mang môi trường acid (chứa nhiều ion H^+^ do quá trình thủy phân các ion Al^3+^, Fe^3+^, SO~4~^2-^).\n*Đáp án đúng: D.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8. Chất nào sau đây có chứa nhóm chức alcohol?",
            "options": [
                "A. CH~3~CHO.",
                "B. CH~3~COCH~3~.",
                "C. CH~3~NH~2~.",
                "*D. C~2~H~5~OH.*"
            ],
            "solution": "- CH~3~CHO: Chứa nhóm aldehyde (-CH=O).\n- CH~3~COCH~3~: Chứa nhóm ketone (>C=O).\n- CH~3~NH~2~: Chứa nhóm amine (-NH~2~).\n- C~2~H~5~OH (ethanol): Chứa nhóm chức hydroxy (-OH) liên kết trực tiếp với nguyên tử carbon no của gốc hydrocarbon, là hợp chất alcohol.\n*Đáp án đúng: D.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9. Trong tự nhiên, nguyên tố nitrogen tồn tại chủ yếu ở dạng đồng vị nào sau đây?",
            "options": [
                "A. ^15^N.",
                "*B. ^14^N.*",
                "C. ^13^N.",
                "D. ^12^N."
            ],
            "solution": "Trong tự nhiên, nguyên tố nitrogen có hai đồng vị bền là ^14^N và ^15^N, trong đó đồng vị ^14^N chiếm khoảng 99,63% tổng số nguyên tử nitrogen, còn ^15^N chỉ chiếm khoảng 0,37%.\n*Đáp án đúng: B.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10. Cặp chất nào sau đây là đồng đẳng của nhau?",
            "options": [
                "*A. CH~3~CH~2~CH~3~ và CH~3~CH~3~.*",
                "B. CH~3~CH~2~CH~2~CH~3~ và (CH~3~)~2~CHCH~3~.",
                "C. CH~3~CH=CH~2~ và CH~3~CH~2~CH~2~CH~3~.",
                "D. CH~2~=CH-CH=CH~2~ và CH~3~C≡CH."
            ],
            "solution": "- Đồng đẳng là hiện tượng các hợp chất hữu cơ có cấu tạo và tính chất hóa học tương tự nhau, nhưng thành phần phân tử hơn kém nhau một hay nhiều nhóm methylene (-CH~2~-).\n- A: CH~3~CH~2~CH~3~ (C~3~H~8~) và CH~3~CH~3~ (C~2~H~6~) đều là alkane mạch hở, hơn kém nhau 1 nhóm -CH~2~-, là đồng đẳng của nhau.\n- B: Cùng công thức C~4~H~10~, là đồng phân của nhau.\n- C: Một chất là alkene (C~3~H~6~), một chất là alkane (C~4~H~10~), không cùng dãy đồng đẳng.\n- D: Cùng C~4~H~6~ nhưng thuộc hai loại mạch khác nhau (alkadiene và alkyne), là đồng phân của nhau.\n*Đáp án đúng: A.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11. Cho iron (III) oxide tác dụng với nitric acid thì sản phẩm thu được là:",
            "options": [
                "*A. Fe(NO~3~)~3~ và H~2~O.*",
                "B. Fe(NO~3~)~3~, NO~2~ và H~2~O.",
                "C. Fe(NO~3~)~3~, N~2~ và H~2~O.",
                "D. Fe(NO~3~)~3~, NO và H~2~O."
            ],
            "solution": "Trong hợp chất Fe~2~O~3~, nguyên tố iron (Fe) đã ở trạng thái số oxi hóa cao nhất là +3, do đó phản ứng giữa Fe~2~O~3~ và dung dịch HNO~3~ chỉ là phản ứng trao đổi (acid - basic oxide), không phải phản ứng oxi hóa - khử và không sinh ra sản phẩm khử của nitrogen:\n  Fe~2~O~3~ + 6 HNO~3~ → 2 Fe(NO~3~)~3~ + 3 H~2~O.\n*Đáp án đúng: A.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12. Trong các hợp chất sau, chất nào là dẫn xuất của hydrocarbon?",
            "options": [
                "A. C~7~H~8~.",
                "B. C~4~H~4~.",
                "*C. CH~3~NH~2~.*",
                "D. C~2~H~2~."
            ],
            "solution": "Dẫn xuất của hydrocarbon là hợp chất hữu cơ mà trong phân tử ngoài nguyên tố C và H còn có chứa các nguyên tố khác như O, N, halogen, S,... CH~3~NH~2~ có chứa nguyên tố N nên là dẫn xuất hydrocarbon; các chất C~7~H~8~, C~4~H~4~, C~2~H~2~ chỉ chứa C và H nên là hydrocarbon.\n*Đáp án đúng: C.*"
        },
        {
            "id": 13,
            "prompt": "Câu 13. Dãy gồm các chất đều tác dụng với sulfur (trong điều kiện thích hợp của phản ứng) là:",
            "options": [
                "A. Au, Pt, Ag.",
                "B. Ag, O~2~, Cl~2~.",
                "*C. Fe, Hg, O~2~.*",
                "D. Fe, Hg, C."
            ],
            "solution": "Sulfur có khả năng tác dụng với nhiều kim loại và phi kim ở điều kiện thích hợp:\n- Fe + S --t°--> FeS\n- Hg + S → HgS (xảy ra ngay ở nhiệt độ thường, ứng dụng khử độc thủy ngân rơi vãi)\n- S + O~2~ --t°--> SO~2~\n(Au, Pt không phản ứng với S ở mọi nhiệt độ).\n*Đáp án đúng: C.*"
        },
        {
            "id": 14,
            "prompt": "Câu 14. Tín hiệu phổ hồng ngoại của một số nhóm chức cơ bản được cho ở Bảng (a) và Phổ hồng ngoại (IR) của hợp chất hữu cơ X được cho ở Hình (b).\n\nBảng (a): Số sóng hấp thụ đặc trưng của một số nhóm chức:\n• Alcohol: 3600 - 3300 cm^-1^ (O-H)\n• Amine: 3500 - 3300 cm^-1^ (N-H)\n• Aldehyde: 2900 - 2700 cm^-1^ (C-H); 1740 - 1720 cm^-1^ (C=O)\n• Ketone: 1725 - 1700 cm^-1^ (C=O)\n• Carboxylic acid: 3300 - 2500 cm^-1^ (O-H); 1725 - 1700 cm^-1^ (C=O)\n• Ester: 1750 - 1735 cm^-1^ (C=O); 1300 - 1000 cm^-1^ (C-O)",
            "image": "hinh_b_ir_crop.png",
            "image_width": Cm(8.5),
            "subprompt": "Phát biểu nào sau đây là đúng?",
            "options": [
                "A. Peak A và peak D cho thấy X là một carboxylic acid.",
                "*B. Peak A là tín hiệu của nhóm OH.*",
                "C. Peak C là tín hiệu của nhóm chức C=O.",
                "D. Peak B và peak D cho thấy X là một aldehyde."
            ],
            "solution": "Quan sát phổ IR ở Hình (b):\n- Peak A là một dải hấp thụ rất mạnh và tù, mở rộng ở vùng số sóng khoảng 3350 cm^-1^. Đối chiếu với Bảng (a), đây là tín hiệu đặc trưng của liên kết O-H trong nhóm alcohol (vùng 3600 - 3300 cm^-1^).\n- Trên phổ hoàn toàn không có peak hấp thụ nhọn và mạnh ở vùng 1700 - 1750 cm^-1^ (vùng đặc trưng của nhóm carbonyl C=O), do đó hợp chất X không thể là carboxylic acid, aldehyde hay ester.\n*Đáp án đúng: B.*"
        }
    ],

    # PHẦN II. Câu trắc nghiệm đúng sai (3,0 điểm - 3 câu)
    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1. Một hợp chất hữu cơ X được xác định có công thức đơn giản nhất là CH~2~O. Bằng phổ khối lượng MS, người ta xác định được phân tử khối của X là 60.",
            "statements": [
                ("a. X là dẫn xuất hydrocarbon.", "Đ"),
                ("b. Công thức phân tử của X là C~2~H~4~O~2~.", "Đ"),
                ("c. Thành phần phần trăm theo khối lượng của nguyên tố carbon trong X là 44%.", "S"),
                ("d. Nếu X chỉ chứa một nhóm chức thì X có 3 đồng phân cấu tạo.", "S")
            ],
            "solution": "Công thức phân tử của X có dạng (CH~2~O)~n~.\nTheo phổ khối lượng MS, phân tử khối M(X) = 60 ⇒ (12 + 2 + 16) × n = 60 ⇒ 30n = 60 ⇒ n = 2.\nCông thức phân tử của X là C~2~H~4~O~2~.\n• a) Đúng. Phân tử X chứa nguyên tố C, H và O nên X là dẫn xuất của hydrocarbon.\n• b) Đúng. Công thức phân tử là C~2~H~4~O~2~.\n• c) Sai. Phần trăm khối lượng của carbon trong X là:\n  %C = (2 × 12 / 60) × 100% = (24 / 60) × 100% = 40% (khác 44%).\n• d) Sai. Với công thức C~2~H~4~O~2~ (độ bất bão hòa k = 1), các đồng phân mạch hở chỉ chứa một nhóm chức gồm có:\n  1. Carboxylic acid: CH~3~-COOH (acetic acid) - 1 đồng phân.\n  2. Ester: H-COO-CH~3~ (methyl formate) - 1 đồng phân.\n  Tổng cộng chỉ có 2 đồng phân đơn chức (nếu kể HO-CH~2~-CHO là hợp chất tạp chức chứa 2 nhóm chức khác nhau: -OH và -CHO). Do đó nói \"có 3 đồng phân cấu tạo\" là sai."
        },
        {
            "id": 2,
            "prompt": "Câu 2. Sulfuric acid là một trong những hoá chất quan trọng nhất được sử dụng trong công nghiệp; được sản xuất hàng trăm triệu tấn mỗi năm, chiếm nhiều nhất trong ngành công nghiệp hoá chất. Phương pháp sản xuất sulfuric acid phổ biến nhất là phương pháp tiếp xúc, theo đó acid có thể được điều chế qua các giai đoạn sau:\n(1) 4 FeS~2~ + 11 O~2~ --t°--> 2 Fe~2~O~3~ + 8 SO~2~\n(2) 2 SO~2~(g) + O~2~(g) <=V~2~O~5~, 450°C=> 2 SO~3~(g); Δ~r~H°~298~ = - 196 kJ\n(3) H~2~SO~4~ + n SO~3~ → H~2~SO~4~.nSO~3~\n(4) H~2~SO~4~.nSO~3~ + n H~2~O → (n+1) H~2~SO~4~",
            "statements": [
                ("a. Có ba phản ứng oxi hóa – khử đã xảy ra trong các giai đoạn sản xuất sulfuric acid trên.", "S"),
                ("b. Để làm tăng hiệu suất tổng hợp SO~3~ ở phản ứng (2) cần tăng nhiệt độ của phản ứng trên 450°C.", "S"),
                ("c. Ở phản ứng (3), có thể dùng nước hấp thụ trực tiếp SO~3~ để tạo thành H~2~SO~4~ với nồng độ cao.", "S"),
                ("d. Để sản xuất được 1 tấn dung dịch H~2~SO~4~ 98% cần 1 tấn quặng pyrite chứa 75% FeS~2~. Biết hiệu suất của cả quá trình điều chế đạt 80% và các tạp chất trong quặng không chứa sulfur.", "Đ")
            ],
            "solution": "• a) Sai. Chỉ có 2 phản ứng oxi hóa – khử là phản ứng (1) và (2) (có sự thay đổi số oxi hóa của Fe, S, O). Phản ứng (3) và (4) không có sự thay đổi số oxi hóa (S vẫn giữ số oxi hóa +6) nên không phải phản ứng oxi hóa – khử.\n• b) Sai. Phản ứng (2) có Δ~r~H°~298~ = -196 kJ < 0 là phản ứng tỏa nhiệt. Theo nguyên lí chuyển dịch cân bằng Le Chatelier, khi tăng nhiệt độ thì cân bằng chuyển dịch theo chiều nghịch (chiều thu nhiệt), làm giảm hiệu suất tạo thành SO~3~. Nhiệt độ 450°C là nhiệt độ tối ưu được lựa chọn để đảm bảo tốc độ phản ứng đủ nhanh và chất xúc tác V~2~O~5~ hoạt động tốt nhất.\n• c) Sai. Không dùng nước để hấp thụ trực tiếp SO~3~ vì phản ứng SO~3~ + H~2~O → H~2~SO~4~ tỏa nhiệt rất mãnh liệt, làm sôi dung dịch và tạo ra các hạt sương mù H~2~SO~4~ lơ lửng rất khó ngưng tụ, bay theo khí thải gây ăn mòn đường ống và ô nhiễm môi trường. Do đó người ta dùng dung dịch H~2~SO~4~ đặc (98%) để hấp thụ SO~3~ tạo thành oleum.\n• d) Đúng. Tính toán chi tiết:\n  - Khối lượng H~2~SO~4~ nguyên chất có trong 1 tấn dung dịch 98%: m(H~2~SO~4~) = 1 × 98% = 0,98 tấn.\n  - Sơ đồ bảo toàn nguyên tố S: FeS~2~ → 2 H~2~SO~4~\n    120 tấn FeS~2~ → 2 × 98 = 196 tấn H~2~SO~4~\n    Khối lượng FeS~2~ lí thuyết: m(FeS~2~)~lt~ = (0,98 × 120) / 196 = 0,6 tấn.\n  - Vì hiệu suất cả quá trình là H = 80% = 0,8 nên khối lượng FeS~2~ thực tế cần lấy:\n    m(FeS~2~)~tt~ = 0,6 / 0,8 = 0,75 tấn.\n  - Do quặng pyrite chỉ chứa 75% FeS~2~ nên khối lượng quặng pyrite cần dùng là:\n    m(quặng) = m(FeS~2~)~tt~ / 75% = 0,75 / 0,75 = 1,0 tấn (hoàn toàn chính xác!)."
        },
        {
            "id": 3,
            "prompt": "Câu 3. Thí nghiệm tách β-carotene từ nước ép cà rốt được thực hiện như sau:\n- Bước 1: Cho khoảng 20 mL nước ép cà rốt vào phễu chiết. Thêm tiếp khoảng 20 mL hexane, lắc đều khoảng 2 phút.\n- Bước 2: Để yên phễu chiết trên giá thí nghiệm khoảng 5 phút để chất lỏng tách thành hai lớp.\n- Bước 3: Mở khoá phễu chiết cho phần nước chảy xuống, còn lại phần dung dịch β-carotene hoà tan trong hexane.",
            "statements": [
                ("a. Phễu chiết tách thành hai lớp, lớp bên trên là β-carotene hoà tan trong hexane, lớp dưới là nước.", "Đ"),
                ("b. Thí nghiệm tách β-carotene từ nước ép cà rốt dựa theo nguyên tắc chiết lỏng – lỏng.", "Đ"),
                ("c. Dùng dung môi là hexane có khả năng hoà tan β-carotene nhưng không tan trong nước và có nhiệt độ sôi thấp để chiết.", "Đ"),
                ("d. Sau bước 3, để thu được β-carotene người ta đun nóng để loại bỏ dung môi hexane.", "Đ")
            ],
            "solution": "• a) Đúng. Hexane là dung môi hữu cơ không phân cực, có khối lượng riêng d ≈ 0,66 g/mL nhẹ hơn nước và không tan trong nước. β-carotene tan tốt trong hexane nên dung dịch β-carotene trong hexane sẽ nổi lên thành lớp chất lỏng màu vàng cam ở phía trên, còn nước ở phía dưới.\n• b) Đúng. Thí nghiệm thực hiện chuyển chất tan (β-carotene) từ một dung dịch lỏng (nước ép cà rốt) sang một dung môi lỏng khác không trộn lẫn (hexane), đây là nguyên tắc chuẩn của phương pháp chiết lỏng - lỏng.\n• c) Đúng. Hexane đáp ứng hoàn hảo các tiêu chí của dung môi chiết: hòa tan tốt chất cần chiết (β-carotene), không trộn lẫn với dung môi ban đầu (nước) và có nhiệt độ sôi thấp (khoảng 68,7°C) giúp dễ dàng thu hồi sau này.\n• d) Đúng. Sau khi tách lấy lớp dịch chiết hexane chứa β-carotene, người ta làm bay hơi (đun nhẹ/cô quay chân không) để loại bỏ hoàn toàn dung môi hexane thu hồi β-carotene tinh khiết."
        }
    ],

    # PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn (1,5 điểm - 6 câu)
    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1. Cho cân bằng hóa học: N~2~O~4~(g) ⇄ 2 NO~2~(g). Ban đầu có 0,02 mol N~2~O~4~ trong bình kín có thể tích 500 mL. Khi phản ứng đạt trạng thái cân bằng, nồng độ của N~2~O~4~ là 0,0055 M. Tính hằng số cân bằng K~c~. (Kết quả làm tròn đến hàng phần trăm).",
            "answer": "0,87",
            "solution": "- Nồng độ ban đầu của N~2~O~4~:\n  C~0~(N~2~O~4~) = n / V = 0,02 / 0,5 = 0,04 M.\n- Ở trạng thái cân bằng: [N~2~O~4~] = 0,0055 M.\n  ⇒ Lượng N~2~O~4~ đã phân li: ΔC(N~2~O~4~) = 0,04 - 0,0055 = 0,0345 M.\n- Theo phương trình: N~2~O~4~(g) ⇄ 2 NO~2~(g)\n  Nồng độ NO~2~ tạo thành ở trạng thái cân bằng là:\n  [NO~2~] = 2 × ΔC(N~2~O~4~) = 2 × 0,0345 = 0,069 M.\n- Hằng số cân bằng K~c~ của phản ứng:\n  K~c~ = [NO~2~]^2^ / [N~2~O~4~] = (0,069)^2^ / 0,0055 = 0,004761 / 0,0055 ≈ 0,865636.\nLàm tròn đến hàng phần trăm: 0,87.\n*Đáp án: 0,87*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Menthol là hợp chất hữu cơ được tìm thấy trong cây bạc hà. Nó được sử dụng làm hương liệu trong nhiều chất như: kem đánh răng, kẹo cao su và xi-rô ho,... menthol có công thức cấu tạo là:\nTổng số nguyên tử có trong phân tử menthol là bao nhiêu?",
            "image": "menthol.png",
            "image_width": Cm(4.5),
            "answer": "31",
            "solution": "Xác định công thức phân tử của menthol (2-isopropyl-5-methylcyclohexanol):\n- Khung carbon gồm: 1 vòng cyclohexane 6 cạnh + 1 nhóm methyl (-CH~3~) + 1 nhóm isopropyl (-CH(CH~3~)~2~) ⇒ Tổng số carbon = 6 + 1 + 3 = 10 nguyên tử C.\n- Phân tử no, có 1 vòng duy nhất (độ bất bão hòa k = 1) và chứa 1 nhóm -OH ⇒ Công thức phân tử là C~10~H~20~O.\n- Tổng số nguyên tử trong một phân tử menthol là:\n  N = 10 (C) + 20 (H) + 1 (O) = 31 nguyên tử.\n*Đáp án: 31*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Cho các chất khí: NO, NO~2~, SO~2~, N~2~O và N~2~. Có bao nhiêu chất khí là nguyên nhân góp phần gây ra hiện tượng mưa acid?",
            "answer": "3",
            "solution": "Các khí chính gây ra hiện tượng mưa acid khi phản ứng với oxygen và hơi nước trong khí quyển tạo thành acid nitric và acid sulfuric là:\n1. NO: 2 NO + O~2~ → 2 NO~2~\n2. NO~2~: 4 NO~2~ + O~2~ + 2 H~2~O → 4 HNO~3~\n3. SO~2~: 2 SO~2~ + O~2~ + 2 H~2~O → 2 H~2~SO~4~\n(N~2~O là khí gây hiệu ứng nhà kính, N~2~ là khí trơ trong không khí, cả hai không gây mưa acid).\n⇒ Có tất cả 3 chất khí là NO, NO~2~ và SO~2~.\n*Đáp án: 3*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Cho dãy các chất sau: HClO~4~, HClO, Mg(OH)~2~, HNO~3~, KH~2~PO~4~, Ca(OH)~2~, NaCl, CuSO~4~, CH~3~COOH. Có bao nhiêu chất trong dãy thuộc loại chất điện li mạnh?",
            "answer": "6 (hoặc 5)",
            "solution": "Phân loại các chất trong dung dịch nước:\n- Chất điện li mạnh gồm:\n  + Các acid mạnh: HClO~4~, HNO~3~ (2 chất).\n  + Các base mạnh: Ca(OH)~2~ (1 chất).\n  + Hầu hết các muối tan: KH~2~PO~4~, NaCl, CuSO~4~ (3 chất).\n  ⇒ Tổng cộng có 6 chất điện li mạnh: HClO~4~, HNO~3~, Ca(OH)~2~, KH~2~PO~4~, NaCl, CuSO~4~.\n- Chất điện li yếu gồm: HClO, Mg(OH)~2~ (base ít tan, điện li yếu), CH~3~COOH (acid yếu).\n(Ghi chú: Nếu quy ước trường phổ thông không tính Ca(OH)~2~ do ít tan thì đáp số là 5 chất).\n*Đáp án chuẩn: 6 (chấp nhận 5 theo thang điểm trường)*"
        },
        {
            "id": 5,
            "prompt": "Câu 5. Cho các chất có công thức cấu tạo sau:\nCó bao nhiêu chất trong các chất trên có tín hiệu trên phổ hồng ngoại ở khoảng tần số 1750 – 1680 cm^-1^? Xem bảng tín hiệu phổ hồng ngoại của một số nhóm chức cơ bản ở phần I-Trắc nghiệm nhiều phương án lựa chọn.",
            "image": "cau5_4chat.png",
            "image_width": Cm(14.0),
            "answer": "2",
            "solution": "Theo Bảng tín hiệu phổ hồng ngoại, vùng số sóng từ 1750 đến 1680 cm^-1^ là vùng hấp thụ đặc trưng của nhóm carbonyl C=O (trong acid, aldehyde, ketone, ester):\n- Chất (1) H~2~N-CH~2~-CH=CH-CH~3~: Chứa nhóm amine và liên kết đôi C=C, không có liên kết C=O ⇒ Không có tín hiệu ở vùng này.\n- Chất (2) CH~3~-CH=CH-CH(OH)-CH~3~: Chứa nhóm alcohol (-OH) và liên kết C=C, không có liên kết C=O ⇒ Không có tín hiệu ở vùng này.\n- Chất (3) CH~3~-C(CH~3~)=CH-CH~2~-COOH: Chứa nhóm carboxylic acid (-COOH) có liên kết C=O hấp thụ mạnh ở 1725 - 1700 cm^-1^ ⇒ Thỏa mãn.\n- Chất (4) CH~3~-C(CH~3~)=CH-CH=O: Chứa nhóm aldehyde (-CH=O) có liên kết C=O hấp thụ mạnh ở 1740 - 1720 cm^-1^ ⇒ Thỏa mãn.\n⇒ Có 2 chất thỏa mãn là chất (3) và chất (4).\n*Đáp án: 2*"
        },
        {
            "id": 6,
            "prompt": "Câu 6. Cho phương trình hóa học của các phản ứng sau:\n(1) 8 NH~3~ + 3 Cl~2~ → 6 NH~4~Cl + N~2~\n(2) 4 NH~3~ + 5 O~2~ --t°, Pt--> 4 NO + 6 H~2~O\n(3) 2 NH~3~ + H~2~SO~4~ → (NH~4~)~2~SO~4~\n(4) 2 NH~3~ + 3 CuO --t°--> N~2~ + 3 Cu + 3 H~2~O\n(5) MgCl~2~ + 2 NH~3~ + 2 H~2~O → Mg(OH)~2~↓ + 2 NH~4~Cl\nCó bao nhiêu phản ứng hóa học trong đó ammonia đóng vai trò là chất khử?",
            "answer": "3",
            "solution": "Xét sự thay đổi số oxi hóa của nguyên tử nitrogen trong NH~3~ (N^-3^):\n- Phản ứng (1): N^-3^ (trong NH~3~) → N~2~^0^ (nhường electron, số oxi hóa tăng) ⇒ NH~3~ là chất khử.\n- Phản ứng (2): N^-3^ (trong NH~3~) → N^+2^O (nhường electron, số oxi hóa tăng) ⇒ NH~3~ là chất khử.\n- Phản ứng (3): N^-3^ không thay đổi số oxi hóa (trong (NH~4~)~2~SO~4~ vẫn là N^-3^), NH~3~ đóng vai trò là base.\n- Phản ứng (4): N^-3^ (trong NH~3~) → N~2~^0^ (nhường electron, số oxi hóa tăng) ⇒ NH~3~ là chất khử.\n- Phản ứng (5): N^-3^ không thay đổi số oxi hóa, phản ứng tạo kết tủa hydroxide, NH~3~ đóng vai trò là base.\n⇒ Có 3 phản ứng trong đó NH~3~ đóng vai trò là chất khử: (1), (2), (4).\n*Đáp án: 3*"
        }
    ],

    # B. PHẦN TỰ LUẬN (2,0 ĐIỂM)
    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1. (1,0 điểm)\nTại một công ty superphosphate và hóa chất, superphosphate kép được sản xuất từ quặng phosphorite theo hai giai đoạn sau:\nCa~3~(PO~4~)~2~ + 3 H~2~SO~4~ → 2 H~3~PO~4~ + 3 CaSO~4~\nCa~3~(PO~4~)~2~ + 4 H~3~PO~4~ → 3 Ca(H~2~PO~4~)~2~\nTrung bình mỗi ngày, công ty sản xuất được 1872 tấn Ca(H~2~PO~4~)~2~, giả thiết hiệu suất của mỗi giai đoạn đạt 100%. Tính thể tích dung dịch H~2~SO~4~ 70% (D = 1,6 tấn/m^3^) mà công ty đã dùng trong một ngày.",
            "num_lines": 10,
            "solution": "1. Tính số mol Ca(H~2~PO~4~)~2~ sản xuất được trong một ngày:\n- Khối lượng mol của Ca(H~2~PO~4~)~2~: M = 40 + (2 + 31 + 64) × 2 = 234 g/mol.\n- Số mol Ca(H~2~PO~4~)~2~:\n  n[Ca(H~2~PO~4~)~2~] = 1872 tấn / 234 = 8 000 kmol.\n\n2. Thiết lập mối quan hệ số mol giữa H~2~SO~4~ và Ca(H~2~PO~4~)~2~:\n- Phương trình giai đoạn (1): Nhân với 2 ta được: 2 Ca~3~(PO~4~)~2~ + 6 H~2~SO~4~ → 4 H~3~PO~4~ + 6 CaSO~4~\n- Phương trình giai đoạn (2): Ca~3~(PO~4~)~2~ + 4 H~3~PO~4~ → 3 Ca(H~2~PO~4~)~2~\n- Cộng hai phương trình lại (loại H~3~PO~4~):\n  3 Ca~3~(PO~4~)~2~ + 6 H~2~SO~4~ → 3 Ca(H~2~PO~4~)~2~ + 6 CaSO~4~\n  Rút gọn hệ số: 1 Ca~3~(PO~4~)~2~ + 2 H~2~SO~4~ → 1 Ca(H~2~PO~4~)~2~ + 2 CaSO~4~\n⇒ Tỉ lệ mol: n(H~2~SO~4~) = 2 × n[Ca(H~2~PO~4~)~2~] = 2 × 8 000 = 16 000 kmol.\n\n3. Tính khối lượng và thể tích dung dịch H~2~SO~4~ 70%:\n- Khối lượng H~2~SO~4~ nguyên chất:\n  m(H~2~SO~4~) = 16 000 kmol × 98 kg/kmol = 1 568 000 kg = 1 568 tấn.\n- Khối lượng dung dịch H~2~SO~4~ 70% cần dùng:\n  m(dd) = m(H~2~SO~4~) / 70% = 1 568 / 0,7 = 2 240 tấn.\n- Thể tích dung dịch H~2~SO~4~ 70% (với khối lượng riêng D = 1,6 tấn/m^3^):\n  V(dd) = m(dd) / D = 2 240 / 1,6 = 1 400 m^3^.\n*Kết luận: Thể tích dung dịch H~2~SO~4~ 70% công ty đã dùng trong một ngày là 1 400 m^3^.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. (1,0 điểm)\nHợp chất hữu cơ X (mạch hở) được dùng làm dung môi trong quá trình tổng hợp dược phẩm. Kết quả phân tích nguyên tố của X như sau: 64,864%C; 13,514%H về khối lượng và còn lại là oxygen. Phân tử khối của X được xác định thông qua kết quả phổ khối lượng với peak ion phân tử có giá trị m/z lớn nhất bằng 74.\na. Hãy xác định công thức phân tử chất hữu cơ X.\nb. Trong phổ hồng ngoại (IR), chất X có số sóng hấp thụ đặc trưng trong khoảng 3500 – 3200 cm^-1^. Hãy viết các đồng phân cấu tạo của X thỏa mãn điều kiện trên. Xem bảng tín hiệu phổ hồng ngoại của một số nhóm chức cơ bản ở phần I-Trắc nghiệm nhiều phương án lựa chọn.",
            "num_lines": 10,
            "solution": "a. Xác định công thức phân tử của X:\n- Thành phần phần trăm khối lượng của oxygen trong X:\n  %O = 100% - 64,864% - 13,514% = 21,622%.\n- Theo phổ khối lượng, giá trị m/z lớn nhất của peak ion phân tử bằng 74 ⇒ Phân tử khối M(X) = 74 amu.\n- Đặt công thức phân tử của X là C~x~H~y~O~z~. Ta có:\n  x = (%C × M) / (12 × 100) = (64,864 × 74) / 1200 ≈ 4\n  y = (%H × M) / (1 × 100) = (13,514 × 74) / 100 ≈ 10\n  z = (%O × M) / (16 × 100) = (21,622 × 74) / 1600 ≈ 1\n⇒ Công thức phân tử của chất hữu cơ X là C~4~H~10~O.\n\nb. Viết các đồng phân cấu tạo thỏa mãn phổ IR:\n- Trong phổ hồng ngoại (IR), chất X có peak hấp thụ đặc trưng trong khoảng 3500 - 3200 cm^-1^. Đối chiếu với Bảng (a) tín hiệu phổ hồng ngoại, đây là vùng hấp thụ đặc trưng của nhóm -OH trong alcohol (ancol).\n- X là alcohol mạch hở, no, đơn chức có công thức C~4~H~10~O. Có 4 đồng phân cấu tạo alcohol thỏa mãn:\n  1. CH~3~-CH~2~-CH~2~-CH~2~-OH   (butan-1-ol)\n  2. CH~3~-CH~2~-CH(OH)-CH~3~     (butan-2-ol)\n  3. (CH~3~)~2~CH-CH~2~-OH         (2-methylpropan-1-ol hay isobutyl alcohol)\n  4. (CH~3~)~3~C-OH                (2-methylpropan-2-ol hay tert-butyl alcohol)."
        }
    ]
}

def generate_docx(is_teacher=False):
    doc = Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line=DATA_EXAM_324["school_line"],
        exam_title=DATA_EXAM_324["exam_title"],
        subject_line=DATA_EXAM_324["subject_line"],
        num_pages=DATA_EXAM_324["num_pages"],
        duration=DATA_EXAM_324["duration"],
        exam_code=DATA_EXAM_324["exam_code"],
        atomic_masses=DATA_EXAM_324["atomic_masses"],
        is_teacher=is_teacher
    )

    # Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Bản Giáo Viên
    if is_teacher:
        add_quick_answer_table_324(doc)

    # A. TRẮC NGHIỆM KHÁCH QUAN (8,0 ĐIỂM)
    add_section_header(
        doc,
        title="A. TRẮC NGHIỆM KHÁCH QUAN (8,0 ĐIỂM)",
        subtitle="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,5 điểm). Thí sinh trả lời từ câu 1 đến câu 14. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    for item in DATA_EXAM_324["part1"]:
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

        if "subprompt" in item:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(1)
            p_sub.paragraph_format.space_after = Pt(2)
            render_rich_text(p_sub, item["subprompt"], base_size=10.5)

        render_options_no_table(doc, item["options"], is_teacher=is_teacher)
        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN II. Câu trắc nghiệm đúng sai.
    add_section_header(
        doc,
        title="PHẦN II: Câu trắc nghiệm đúng sai (3,0 điểm).",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 3. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    for item in DATA_EXAM_324["part2"]:
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

    # PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.
    add_section_header(
        doc,
        title="PHẦN III: Câu trắc nghiệm yêu cầu trả lời ngắn (1,5 điểm).",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 6. Thí sinh điền kết quả vào ô trả lời tương ứng."
    )

    for item in DATA_EXAM_324["part3"]:
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

    # B. PHẦN TỰ LUẬN (2,0 ĐIỂM)
    add_section_header(
        doc,
        title="B. PHẦN TỰ LUẬN (2,0 ĐIỂM)",
        subtitle="Thí sinh trình bày chi tiết lời giải vào giấy làm bài thi."
    )

    for item in DATA_EXAM_324["part4"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])
        elif not is_teacher:
            num_lines = item.get("num_lines", 6)
            add_student_writing_lines(doc, num_lines=num_lines)

    # Phần kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(0)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(r_end, font_size=11, bold=True)

    suffix = "_GiaoVien.docx" if is_teacher else "_HocSinh.docx"
    filename = f"De_Kiem_Tra_CK1_Hoa_11_CaoThang_Ma324{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
