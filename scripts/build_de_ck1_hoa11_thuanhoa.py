# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Kì 1 - Hóa học 11 - Trường THPT Thuận Hóa (Năm học 2025-2026)
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

# Thư mục lưu trữ hình ảnh và sản phẩm
DIR_IMAGES = os.path.abspath("San_Pham/images_ck1_hoa11_thuanhoa")
DIR_SAN_PHAM = os.path.abspath("San_Pham")
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

def add_quick_answer_table_thuanhoa(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Giáo viên ở đầu tài liệu Đề THPT Thuận Hóa.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - THPT THUẬN HÓA)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I (18 câu - chia 2 bảng 9 câu để bố cục thoáng đẹp)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (4,5 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=4, cols=10)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_row1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9"]
    answers_row1 = ["Đ/A", "A", "D", "B", "B", "D", "C", "D", "D", "B"]
    headers_row2 = ["Câu", "10", "11", "12", "13", "14", "15", "16", "17", "18"]
    answers_row2 = ["Đ/A", "C", "C", "D", "B", "C", "B", "B", "B", "D"]

    w1_c0 = Cm(1.8)
    w1_c = Cm(1.8)

    for r_idx, r_data in enumerate([headers_row1, answers_row1, headers_row2, answers_row2]):
        row = tbl1.rows[r_idx]
        is_header = (r_idx in [0, 2])
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = w1_c0 if c_idx == 0 else w1_c
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if is_header:
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

    # 2. PHẦN II (2 câu - 2,0 điểm)
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(5)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM - MỖI CÂU 1,0 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl2.autofit = False

    tbl2_data = [
        ["Câu hỏi", "Nội dung câu hỏi", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Tẩy rửa vết sơn móng tay và mực bút bi trên áo", "ĐÚNG", "SAI", "ĐÚNG", "SAI"],
        ["Câu 2", "Thí nghiệm độ dẫn điện (Ethanol, KCl, Acetic acid)", "ĐÚNG", "ĐÚNG", "SAI", "SAI"]
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

    # 3. PHẦN III (4 câu - 2,0 điểm)
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(5)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM - MỖI CÂU 0,5 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False

    tbl3_headers = ["Câu", "Câu 1", "Câu 2", "Câu 3", "Câu 4"]
    tbl3_answers = ["Đáp án", "19", "6", "3", "4"]

    w3_c0 = Cm(2.4)
    w3_c = Cm(3.9)

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

    # 4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN (1,5 điểm)
    p_tl = doc.add_paragraph()
    p_tl.paragraph_format.space_before = Pt(5)
    p_tl.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl, "*4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN (1,5 ĐIỂM)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl4 = doc.add_table(rows=4, cols=3)
    tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl4.autofit = False

    tbl4_data = [
        ("Câu", "Điểm", "Đáp số cốt lõi & Hướng dẫn chấm"),
        ("Câu 1", "0,5 đ", "• %O = 26,7% ⇒ Tỉ lệ C : H : O = 3 : 8 : 1 ⇒ Công thức đơn giản nhất: C~3~H~8~O.\n• Pic ion phân tử [M^+^] trên phổ khối lượng MS có m/z = 60 ⇒ M = 60.\n• Công thức phân tử của hợp chất A là C~3~H~8~O."),
        ("Câu 2", "0,5 đ", "• Thể tích nước mưa rơi xuống KCN: V = 10 km^2^ × 80 mm = 800.000 m^3^ = 8.10^8^ lít.\n• Số mol H~2~SO~4~: n = 2.10^-5^ × 8.10^8^ = 16.000 mol.\n• Khối lượng H~2~SO~4~ trong nước mưa: m = 16.000 × 98 = 1.568.000 g = 1.568 kg (hay 1,568 tấn)."),
        ("Câu 3", "0,5 đ", "• Hằng số cân bằng: K~c~ = ([HbCO] × [O~2~]) / ([HbO~2~] × [CO]) = 170.\n• Tỉ lệ hòa tan trong máu: [O~2~] / [CO] = 20,0% / 0,1% = 200.\n• Tỉ lệ giữa HbCO và HbO~2~ trong máu: [HbCO] / [HbO~2~] = 170 / 200 = 0,85 (hay 85%).")
    ]

    w4_widths = [Cm(2.0), Cm(1.8), Cm(14.2)]

    for r_idx, row_vals in enumerate(tbl4_data):
        row = tbl4.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
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
                is_label = (c_idx < 2)
                set_cell_shading(cell, "F5F5F5" if is_label else "FFFFFF")
                c_color = COLOR_PRIMARY if is_label else RGBColor(0, 0, 0)
                render_rich_text(p, val, base_size=9.5, base_bold=is_label, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(6)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


# Dữ liệu Đề thi
DATA_EXAM = {
    "school_line": "TRƯỜNG ĐHSP, ĐHH - TRƯỜNG THPT THUẬN HÓA",
    "exam_title": "ĐỀ KIỂM TRA CUỐI KÌ 1 NĂM HỌC 2025 - 2026",
    "subject_line": "MÔN: HOÁ HỌC - LỚP 11",
    "duration": "45 phút",
    "num_pages": "03",
    "exam_code": "101",
    "atomic_masses": "H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; P = 31; Ca = 40; Cu = 64; Ag = 108",
    
    # PHẦN I: 18 câu trắc nghiệm nhiều phương án lựa chọn (4,5 điểm)
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1: Cho cân bằng hoá học: H~2~(g) + I~2~(g) ⇌ 2HI(g) (ở 450 °C, Δ~r~H > 0). Cân bằng không bị chuyển dịch khi",
            "options": [
                "*A. giảm áp suất chung của hệ.*",
                "B. tăng nhiệt độ của hệ.",
                "C. tăng nồng độ H~2~.",
                "D. giảm nồng độ HI."
            ],
            "solution": "Phản ứng có tổng số mol khí ở hai vế bằng nhau (1 mol H~2~ + 1 mol I~2~ = 2 mol HI). Do đó, sự thay đổi áp suất chung của hệ không làm thay đổi tốc độ phản ứng thuận và nghịch theo các tỉ lệ khác nhau, nên cân bằng không bị chuyển dịch.\n*Đáp án đúng: A.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2: Đặc điểm dễ dàng nhận biết hiện tượng phú dưỡng ở các ao hồ là",
            "options": [
                "A. Nhiều loài cá sống nổi bềnh lên mặt nước.",
                "B. Nước ao màu vàng của tảo phát triển.",
                "C. Nước ao màu đen của tảo phát triển.",
                "*D. Nước ao màu xanh của tảo phát triển.*"
            ],
            "solution": "Hiện tượng phú dưỡng (eutrophication) xảy ra khi nguồn nước tích tụ quá mức các nguyên tố dinh dưỡng (N, P), kích thích thực vật thủy sinh và tảo sinh sôi nảy nở bùng phát (\"nở hoa nước\"), tạo thành một lớp váng màu xanh đặc trưng trên mặt hồ.\n*Đáp án đúng: D.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3: Sử dụng phương pháp kết tinh lại để tinh chế chất rắn. Hợp chất cần kết tinh lại cần có tính chất nào dưới đây để việc kết tinh lại được thuận lợi?",
            "options": [
                "A. Ít tan trong cả dung dịch nóng và lạnh.",
                "*B. Tan tốt trong dung dịch nóng, ít tan trong dung dịch lạnh.*",
                "C. Tan trong dung môi phân cực, không tan trong dung môi không phân cực.",
                "D. Tan tốt trong cả dung dịch nóng và lạnh."
            ],
            "solution": "Nguyên tắc chọn dung môi cho quá trình kết tinh lại chất rắn là: Hợp chất cần tinh chế phải tan tốt trong dung môi ở nhiệt độ cao (dung dịch nóng) và tan kém hoặc hầu như không tan ở nhiệt độ thấp (dung dịch lạnh). Khi làm nguội dung dịch bão hòa nóng, chất rắn sẽ kết tinh tách ra ngoài dạng tinh thể tinh khiết.\n*Đáp án đúng: B.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4: Phát biểu nào sau đây không đúng?",
            "options": [
                "A. Hai chất có cùng công thức thực nghiệm thì thành phần các nguyên tố trong phân tử của chúng là giống nhau.",
                "*B. Hai chất có cùng công thức thực nghiệm luôn có cùng công thức phân tử.*",
                "C. Hai chất có cùng công thức thực nghiệm có thể có phân tử khối khác nhau.",
                "D. Hai chất có cùng công thức thực nghiệm có phần trăm khối lượng các nguyên tố trong phân tử của chúng là như nhau."
            ],
            "solution": "Công thức phân tử bằng (CTTN)~n~ với n ≥ 1. Hai chất có cùng công thức thực nghiệm có thể có hệ số n khác nhau dẫn tới công thức phân tử và phân tử khối khác nhau (Ví dụ: C~2~H~2~ và C~6~H~6~ cùng có CTTN là CH nhưng CTPT khác nhau). Do đó phát biểu B \"luôn có cùng công thức phân tử\" là sai.\n*Đáp án đúng: B.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5: Cấu tạo hoá học là",
            "options": [
                "A. tỉ lệ số nguyên tử trong phân tử.",
                "B. liên kết giữa các nguyên tử trong phân tử.",
                "C. tính chất hoá học của phân tử đó.",
                "*D. thứ tự liên kết giữa các nguyên tử trong phân tử.*"
            ],
            "solution": "Theo thuyết cấu tạo hoá học của A. M. Butlerov: Thứ tự liên kết giữa các nguyên tử trong phân tử được gọi là cấu tạo hoá học.\n*Đáp án đúng: D.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6: Cho các phát biểu sau:\n(1) Sulfuric acid là chất lỏng, sánh như dầu, không màu, không bay hơi.\n(2) Sulfuric acid tan vô hạn trong nước, và tỏa rất nhiều nhiệt.\n(3) Khi pha loãng sulfuric acid đặc, ta cho từ từ acid vào nước và khuấy nhẹ.\n(4) Sulfuric acid đặc có tính háo nước, da thịt tiếp xúc với nó sẽ gây bỏng nặng.\nSố phát biểu đúng là:",
            "options": [
                "A. 3.",
                "B. 1.",
                "*C. 4.*",
                "D. 2."
            ],
            "solution": "Cả 4 phát biểu (1), (2), (3), (4) đều hoàn toàn chính xác theo kiến thức về tính chất vật lý và an toàn thực hành với sulfuric acid đặc (H~2~SO~4~).\n*Đáp án đúng: C.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7: Cho hỗn hợp các alkane sau: pentane (sôi ở 36 °C), heptane (sôi ở 98 °C), octane (sôi ở 126 °C), nonane (sôi ở 151 °C). Có thể tách riêng các chất đó bằng cách nào sau đây?",
            "options": [
                "A. Chiết.",
                "B. Thăng hoa.",
                "C. Kết tinh.",
                "*D. Chưng cất.*"
            ],
            "solution": "Các alkane trên đều là chất lỏng tan lẫn vào nhau đồng nhất và có nhiệt độ sôi cách biệt nhau đáng kể (từ 25 °C trở lên). Do đó phương pháp phù hợp nhất để tách riêng từng chất là chưng cất phân đoạn.\n*Đáp án đúng: D.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8: Trong phân tử hợp chất hữu cơ, carbon có hóa trị mấy?",
            "options": [
                "A. I.",
                "B. II.",
                "C. III.",
                "*D. IV.*"
            ],
            "solution": "Theo thuyết cấu tạo hoá học, trong tất cả các hợp chất hữu cơ, nguyên tử carbon luôn luôn có hóa trị IV.\n*Đáp án đúng: D.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9: Trong những cặp chất sau đây, cặp gồm các chất là đồng phân của nhau?",
            "options": [
                "A. C~4~H~10~, C~6~H~6~.",
                "*B. C~2~H~5~OH, CH~3~OCH~3~.*",
                "C. CH~3~CH~2~CH~2~OH, C~2~H~5~OH.",
                "D. CH~3~OCH~3~, CH~3~CHO."
            ],
            "solution": "Các chất đồng phân là những chất có cùng công thức phân tử nhưng khác nhau về cấu tạo hoá học. C~2~H~5~OH (ethanol) và CH~3~OCH~3~ (dimethyl ether) cùng có chung công thức phân tử là C~2~H~6~O.\n*Đáp án đúng: B.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10: Để phân biệt hai acid H~2~SO~4~ loãng và acid HCl có thể dùng chất nào sau đây?",
            "options": [
                "A. MgCl~2~.",
                "B. Na~2~CO~3~.",
                "*C. BaCl~2~.*",
                "D. NaNO~3~."
            ],
            "solution": "Sử dụng dung dịch BaCl~2~:\n- Dung dịch H~2~SO~4~ tạo kết tủa trắng không tan trong acid: BaCl~2~ + H~2~SO~4~ → BaSO~4~↓ + 2HCl.\n- Dung dịch HCl không có hiện tượng phản ứng.\n*Đáp án đúng: C.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11: Dung dịch của một base ở 25 °C có:",
            "options": [
                "A. [H^+^] > 1,0.10^-7^ M",
                "B. [H^+^].[OH^-^] > 1,0.10^-14^ M",
                "*C. [H^+^] < 1,0.10^-7^ M*",
                "D. [H^+^] = 1,0.10^-7^ M"
            ],
            "solution": "Ở 25 °C, tích số ion của nước là K~w~ = [H^+^][OH^-^] = 1,0.10^-14^. Trong dung dịch base có [OH^-^] > 1,0.10^-7^ M, do đó [H^+^] = 1,0.10^-14^ / [OH^-^] < 1,0.10^-7^ M (tương đương pH > 7).\n*Đáp án đúng: C.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12: Phản ứng nào sau đây là phản ứng một chiều?",
            "options": [
                "A. Cl~2~ + H~2~O ⇌ HCl + HClO.",
                "B. C + CO~2~ ⇌ 2CO.",
                "C. CH~3~COONa + H~2~O ⇌ CH~3~COOH + NaOH.",
                "*D. CH~4~ + 2O~2~ → CO~2~ + 2H~2~O (t°).*"
            ],
            "solution": "Phản ứng đốt cháy hydrocarbon (CH~4~ + 2O~2~ → CO~2~ + 2H~2~O) là phản ứng oxi hóa hoàn toàn diễn ra theo một chiều không thuận nghịch.\n*Đáp án đúng: D.*"
        },
        {
            "id": 13,
            "prompt": "Câu 13: Khí nào phổ biến nhất trong khí quyển Trái Đất?",
            "options": [
                "A. Argon.",
                "*B. Nitrogen.*",
                "C. Oxygen.",
                "D. Ozone."
            ],
            "solution": "Khí nitrogen (N~2~) chiếm khoảng 78% thể tích không khí quyển Trái Đất, là khí dồi dào và phổ biến nhất.\n*Đáp án đúng: B.*"
        },
        {
            "id": 14,
            "prompt": "Câu 14: Chất Y có công thức cấu tạo dạng khung phân tử như sau:",
            "image": "cau14_khung.png",
            "image_width": 2.2,
            "prompt_after_img": "Công thức phân tử của Y là:",
            "options": [
                "A. C~3~H~8~O.",
                "B. C~3~H~6~O.",
                "*C. C~4~H~10~O.*",
                "D. C~4~H~8~O."
            ],
            "solution": "Dạng khung phân tử thể hiện mạch gồm 4 nguyên tử carbon no mạch hở liên kết với nhóm –OH ở đầu mạch: CH~3~–CH~2~–CH~2~–CH~2~–OH (butan-1-ol). Công thức phân tử tương ứng là C~4~H~10~O.\n*Đáp án đúng: C.*"
        },
        {
            "id": 15,
            "prompt": "Câu 15: Trộn khí X vào khí Y thấy xuất hiện khói trắng. Các khí X và Y có thể là các khí nào sau đây?",
            "options": [
                "A. N~2~ và O~2~.",
                "*B. NH~3~ và HCl.*",
                "C. Cl~2~ và H~2~.",
                "D. N~2~ và H~2~."
            ],
            "solution": "Khi hai chất khí NH~3~ và HCl gặp nhau, chúng phản ứng nhanh tạo thành các hạt tinh thể rắn muối ammonium chloride (NH~4~Cl) lơ lửng như khói trắng: NH~3~(g) + HCl(g) → NH~4~Cl(s).\n*Đáp án đúng: B.*"
        },
        {
            "id": 16,
            "prompt": "Câu 16: Dung dịch nào sau đây có pH < 7?",
            "options": [
                "A. MgCl~2~.",
                "*B. HCl.*",
                "C. C~2~H~5~OH.",
                "D. Na~2~CO~3~."
            ],
            "solution": "Hydrochloric acid (HCl) là acid mạnh, trong nước phân li hoàn toàn sinh ra H^+^ làm cho [H^+^] > 10^-7^ M, do đó pH < 7.\n*Đáp án đúng: B.*"
        },
        {
            "id": 17,
            "prompt": "Câu 17: Các tính chất hoá học của nitric acid (HNO~3~) là:",
            "options": [
                "A. tính oxi hóa mạnh và tính base mạnh.",
                "*B. tính acid mạnh và tính oxi hóa mạnh.*",
                "C. tính acid mạnh và tính khử mạnh.",
                "D. tính oxi hóa mạnh và tính acid yếu."
            ],
            "solution": "Nitric acid (HNO~3~) là một acid vô cơ mạnh và đồng thời là chất oxi hóa rất mạnh trong hầu hết các phản ứng hoá học do nguyên tử nitrogen có số oxi hóa cực đại là +5.\n*Đáp án đúng: B.*"
        },
        {
            "id": 18,
            "prompt": "Câu 18: Trong phản ứng: SO~2~ + 2H~2~S → 3S + 2H~2~O. Câu nào diễn tả đúng tính chất của chất?",
            "options": [
                "A. SO~2~ là chất khử và H~2~S là chất oxi hóa.",
                "B. H~2~S vừa là chất khử vừa là chất oxi hóa.",
                "C. SO~2~ bị oxi hóa và H~2~S bị khử.",
                "*D. SO~2~ bị khử, H~2~S bị oxi hóa đều tạo thành S.*"
            ],
            "solution": "Số oxi hóa của S trong SO~2~ giảm từ +4 xuống 0 (nhận 4e) nên SO~2~ là chất oxi hóa (bị khử). Số oxi hóa của S trong H~2~S tăng từ -2 lên 0 (nhường 2e) nên H~2~S là chất khử (bị oxi hóa). Cả hai chất cùng chuyển hóa thành đơn chất sulfur (S).\n*Đáp án đúng: D.*"
        }
    ],

    # PHẦN II: 2 câu đúng sai (2,0 điểm)
    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1: Sơn móng tay dùng để làm đẹp hay vết mực bút bi dây trên áo, thì rất khó để tẩy, phải dùng loại hóa chất thích hợp mới tẩy được.",
            "image": "p2_c1_sonmong.png",
            "image_width": 4.5,
            "statements": [
                {"lbl": "a)", "text": "Để loại bỏ vết sơn móng tay thì thường dùng acetone (chất hữu cơ).", "ans": "Đúng"},
                {"lbl": "b)", "text": "Có thể dùng nước để tẩy vết sơn móng tay.", "ans": "Sai"},
                {"lbl": "c)", "text": "Vết sơn móng tay hay vết mực bút bi là các chất hữu cơ.", "ans": "Đúng"},
                {"lbl": "d)", "text": "Vết sơn móng tay hay vết mực bút bi dây trên áo dễ hòa tan trong nước.", "ans": "Sai"}
            ],
            "solution": "- Ý a) ĐÚNG: Acetone ((CH~3~)~2~CO) là dung môi hữu cơ phân cực hòa tan rất tốt chất kết dính polyme nitrocellulose có trong thành phần sơn móng tay.\n- Ý b) SAI: Nước là dung môi phân cực mạnh nhưng màng sơn móng tay kỵ nước nên không tan trong nước, không thể tẩy sạch bằng nước.\n- Ý c) ĐÚNG: Sơn móng tay chứa polyme, chất hóa dẻo, dung môi este hữu cơ; mực bút bi chứa các hạt nhựa và chất nhuộm hữu cơ.\n- Ý d) SAI: Các chất này đều kỵ nước, hầu như không tan trong nước nên không thể dễ hòa tan trong nước."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Thực hiện thí nghiệm với 3 chất tan khác nhau là ethanol, potassium chloride và acid acetic như hình dưới đây.",
            "image": "p2_c2_dandien.png",
            "image_width": 5.2,
            "statements": [
                {"lbl": "a)", "text": "Dung dịch KCl và acetic acid có khả năng dẫn điện.", "ans": "Đúng"},
                {"lbl": "b)", "text": "Dung dịch hòa tan ethanol không dẫn điện.", "ans": "Đúng"},
                {"lbl": "c)", "text": "Các phân tử acetic acid có 1 đầu mang điện dương và một đầu mang điện âm, chuyển động trong dung dịch tạo thành dòng điện.", "ans": "Sai"},
                {"lbl": "d)", "text": "Ethanol là chất điện li.", "ans": "Sai"}
            ],
            "solution": "- Ý a) ĐÚNG: KCl là chất điện li mạnh phân li hoàn toàn thành K^+^ và Cl^-^; acetic acid là chất điện li yếu phân li một phần thành CH~3~COO^-^ và H^+^. Sự có mặt của các ion chuyển động tự do giúp dung dịch dẫn điện.\n- Ý b) ĐÚNG: Dung dịch ethanol (C~2~H~5~OH) tồn tại ở dạng phân tử trung hòa, không phân li ra ion tự do nên không dẫn điện (bóng đèn không sáng).\n- Ý c) SAI: Dòng điện trong dung dịch dẫn điện là do dòng chuyển dịch có hướng của các cation và anion tự do dưới tác dụng của điện trường, không phải do phân tử phân cực chuyển động.\n- Ý d) SAI: Ethanol khi tan trong nước không phân li thành ion, do đó ethanol là chất không điện li."
        }
    ],

    # PHẦN III: 4 câu trả lời ngắn (2,0 điểm)
    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1: Cho sơ đồ phản ứng sau: NH~3~ + O~2~ → NO + H~2~O. Sau khi cân bằng với hệ số nguyên tối giản, tổng hệ số của tất cả các chất trong phản ứng là bao nhiêu?",
            "ans": "19",
            "solution": "Phương trình phản ứng oxi hóa NH~3~ bằng O~2~ với xúc tác Pt ở 850 - 900 °C:\n4NH~3~ + 5O~2~ → 4NO + 6H~2~O\nTổng hệ số của tất cả các chất trong phản ứng = 4 + 5 + 4 + 6 = 19.\n*Trả lời: 19.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2: Chất hữu cơ X có công thức đơn giản nhất CH~3~O và có phân tử khối bằng 62. Công thức phân tử hợp chất hữu cơ X có bao nhiêu nguyên tử hydrogen?",
            "ans": "6",
            "solution": "Công thức phân tử của hợp chất hữu cơ X có dạng: (CH~3~O)~n~ (với n nguyên dương).\nPhân tử khối: M = (12.1 + 1.3 + 16.1).n = 31n = 62 ⇒ n = 2.\nCông thức phân tử của X là C~2~H~6~O~2~ (ethylene glycol).\nSố nguyên tử hydrogen (H) trong phân tử X là 6.\n*Trả lời: 6.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3: Cho cân bằng hóa học sau: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g); Δ~r~H°~298~ < 0. Cho các biện pháp:\n(a) tăng nhiệt độ,\n(b) tăng áp suất chung của hệ phản ứng,\n(c) hạ nhiệt độ,\n(d) dùng thêm chất xúc tác V~2~O~5~,\n(đ) giảm nồng độ SO~3~,\n(e) giảm áp suất chung của hệ phản ứng.\nCó bao nhiêu biện pháp làm cân bằng trên chuyển dịch theo chiều thuận?",
            "ans": "3",
            "solution": "Phân tích tác động theo nguyên lí chuyển dịch cân bằng Le Chatelier:\n- Chiều thuận là chiều tỏa nhiệt (Δ~r~H°~298~ < 0) ⇒ Khi hạ nhiệt độ (c), cân bằng chuyển dịch theo chiều tỏa nhiệt để chống lại sự giảm nhiệt (chiều thuận).\n- Chiều thuận có số mol khí giảm (3 mol khí → 2 mol khí) ⇒ Khi tăng áp suất (b), cân bằng chuyển dịch theo chiều giảm số mol khí (chiều thuận).\n- Dùng chất xúc tác V~2~O~5~ (d) chỉ làm tăng tốc độ đạt tới trạng thái cân bằng, không làm chuyển dịch vị trí cân bằng.\n- Giảm nồng độ SO~3~ (đ): Cân bằng chuyển dịch theo chiều tạo thêm SO~3~ để bù lại lượng bị giảm (chiều thuận).\nNhư vậy có 3 biện pháp làm cân bằng chuyển dịch theo chiều thuận: (b), (c), (đ).\n*Trả lời: 3.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4: Cho các chất sau: Na~2~CO~3~, BaCl~2~, MgSO~4~, CH~3~COONa, C~2~H~5~Br, CaO, CHCl~3~, HCOOH. Số hợp chất hữu cơ trong các chất trên là bao nhiêu?",
            "ans": "4",
            "solution": "Hợp chất hữu cơ là hợp chất của carbon (loại trừ CO, CO~2~, muối carbonate, carbide, cyanide,...).\n- Các chất vô cơ: Na~2~CO~3~ (muối carbonate), BaCl~2~, MgSO~4~, CaO.\n- Các hợp chất hữu cơ gồm 4 chất:\n  1. CH~3~COONa (muối của acid hữu cơ acetic acid);\n  2. C~2~H~5~Br (bromoethane - dẫn xuất halogen);\n  3. CHCl~3~ (chloroform - dẫn xuất halogen);\n  4. HCOOH (formic acid - carboxylic acid).\nVậy số hợp chất hữu cơ là 4.\n*Trả lời: 4.*"
        }
    ],

    # PHẦN IV: 3 câu tự luận (1,5 điểm)
    "part4": [
        {
            "id": 1,
            "score": "0,5 điểm",
            "prompt": "Câu 1 (0,5 điểm): Hợp chất hữu cơ A có phần trăm về khối lượng của C là 60%; H là 13,3% còn lại là oxygen. Phổ khối lượng của hợp chất hữu cơ A được cho như hình vẽ bên dưới. Hãy xác định công thức phân tử của A.",
            "image": "p4_c1_ms.png",
            "image_width": 5.0,
            "solution": "• Tính phần trăm khối lượng của oxygen trong A:\n  %O = 100% - 60% - 13,3% = 26,7%.\n• Gọi công thức đơn giản nhất của A là C~x~H~y~O~z~. Lập tỉ lệ số mol nguyên tử:\n  x : y : z = (%C / 12) : (%H / 1) : (%O / 16) = (60 / 12) : (13,3 / 1) : (26,7 / 16)\n  x : y : z = 5 : 13,3 : 1,67 = 3 : 8 : 1.\n  ⇒ Công thức đơn giản nhất của A là C~3~H~8~O.\n• Công thức phân tử của A có dạng (C~3~H~8~O)~n~.\n• Quan sát phổ khối lượng (MS) của hợp chất A, pic ion phân tử [M^+^] có giá trị m/z = 60.\n  ⇒ Phân tử khối của A là M~A~ = 60.\n  Ta có: 60.n = 60 ⇒ n = 1.\nVậy công thức phân tử của hợp chất A là C~3~H~8~O."
        },
        {
            "id": 2,
            "score": "0,5 điểm",
            "prompt": "Câu 2 (0,5 điểm): Hiện nay, mưa acid, hiệu ứng nhà kính và thủng tầng ozone là ba thảm họa môi trường toàn cầu. Mưa acid tàn phá nhiều rừng cây, các công trình kiến trúc bằng đá và kim loại. Tác nhân chủ yếu gây mưa acid là sulfur dioxide. Trong khí quyển, SO~2~ chuyển hóa thành H~2~SO~4~ trong nước mưa theo sơ đồ sau:\nSO~2~ ⎯⎯⎯⎯⎯⎯→ SO~3~ ⎯⎯⎯⎯→ H~2~SO~4~\nMột cơn mưa acid xuất hiện tại một khu công nghiệp diện tích 10 km^2^ với lượng mưa trung bình 80 mm. Hãy tính:\n- Thể tích nước mưa đã rơi xuống khu công nghiệp.\n- Khối lượng H~2~SO~4~ trong lượng nước mưa, biết nồng độ của H~2~SO~4~ trong nước mưa là 2.10^-5^ M.",
            "solution": "1. Thể tích nước mưa rơi xuống khu công nghiệp:\n- Đổi đơn vị diện tích: S = 10 km^2^ = 10 × 10^6^ m^2^ = 10^7^ m^2^.\n- Đổi lượng mưa trung bình: h = 80 mm = 0,08 m.\n- Thể tích nước mưa rơi xuống:\n  V = S × h = 10^7^ m^2^ × 0,08 m = 800.000 m^3^ = 8.10^5^ m^3^.\n- Đổi ra lít (1 m^3^ = 10^3^ lít):\n  V = 8.10^5^ × 10^3^ = 8.10^8^ lít (800 triệu lít nước mưa).\n\n2. Khối lượng H~2~SO~4~ trong lượng nước mưa:\n- Số mol H~2~SO~4~ có trong lượng nước mưa là:\n  n(H~2~SO~4~) = C~M~ × V = 2.10^-5^ mol/L × 8.10^8^ L = 16.000 mol.\n- Khối lượng H~2~SO~4~ tương ứng:\n  m(H~2~SO~4~) = n × M = 16.000 mol × 98 g/mol = 1.568.000 gam = 1.568 kg (tương đương 1,568 tấn)."
        },
        {
            "id": 3,
            "score": "0,5 điểm",
            "prompt": "Câu 3 (0,5 điểm): Carbon monoxide thay thế oxygen trong hemoglobin đã bị oxi hóa theo phản ứng:\nHbO~2~(aq) + CO(aq) ⇌ HbCO(aq) + O~2~(aq).\nTại nhiệt độ trung bình trong cơ thể, hằng số cân bằng của phản ứng trên là K~c~ = 170. Giả sử một hỗn hợp không khí bị ô nhiễm carbon monoxide ở mức 0,1% (theo thể tích). Coi không khí chứa 20,0% oxygen về thể tích, tỉ lệ oxygen và carbon monoxide hòa tan trong máu giống tỉ lệ của chúng trong không khí. Cho biết tỉ lệ HbCO so với HbO~2~ trong máu là bao nhiêu.",
            "solution": "• Biểu thức hằng số cân bằng của phản ứng là:\n  K~c~ = ([HbCO] × [O~2~]) / ([HbO~2~] × [CO]) = 170\n• Theo đề bài, tỉ lệ nồng độ hòa tan trong máu bằng tỉ lệ phần trăm thể tích của các khí trong không khí hít vào:\n  [O~2~] / [CO] = %V(O~2~) / %V(CO) = 20,0% / 0,1% = 200\n• Thay vào biểu thức hằng số cân bằng:\n  ([HbCO] / [HbO~2~]) × ([O~2~] / [CO]) = 170\n  ⇒ ([HbCO] / [HbO~2~]) × 200 = 170\n  ⇒ [HbCO] / [HbO~2~] = 170 / 200 = 0,85 (hay 85%).\nVậy tỉ lệ HbCO so với HbO~2~ trong máu là 0,85."
        }
    ]
}

def generate_exam_doc(is_teacher=False):
    """Tạo tài liệu Word cho đề kiểm tra (Bản Học Sinh hoặc Bản Giáo Viên)."""
    doc = Document()
    setup_header_footer(doc)
    
    # 1. Khối tiêu đề đầu trang
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

    # 2. BẢNG ĐÁP ÁN RÚT GỌN (Chỉ hiển thị ở Bản Giáo Viên)
    if is_teacher:
        add_quick_answer_table_thuanhoa(doc)

    # 3. PHẦN I
    add_section_header(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án (4,5 điểm)."
    )
    for q in DATA_EXAM["part1"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"])

        # Nếu có ảnh đính kèm (như Câu 14)
        if "image" in q:
            img_path = os.path.join(DIR_IMAGES, q["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_path, width=Cm(q.get("image_width", 3.0)))
            if "prompt_after_img" in q:
                p_post = doc.add_paragraph()
                p_post.paragraph_format.space_before = Pt(1)
                p_post.paragraph_format.space_after = Pt(2)
                render_rich_text(p_post, q["prompt_after_img"])

        # Các phương án
        render_options_no_table(doc, q["options"], is_teacher=is_teacher)

        # Lời giải chi tiết
        if is_teacher and "solution" in q:
            add_solution_no_table(doc, q["solution"])

    # 4. PHẦN II
    add_section_header(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai (2,0 điểm)."
    )
    for q in DATA_EXAM["part2"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"])

        # Ảnh minh họa
        if "image" in q:
            img_path = os.path.join(DIR_IMAGES, q["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_path, width=Cm(q.get("image_width", 5.0)))

        for stmt in q["statements"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, f"{stmt['lbl']} {stmt['text']}", ans=stmt["ans"], is_teacher=is_teacher)

        if is_teacher and "solution" in q:
            add_solution_no_table(doc, q["solution"])

    # 5. PHẦN III
    add_section_header(
        doc,
        "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 4 (2,0 điểm)."
    )
    for q in DATA_EXAM["part3"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"])

        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(2)
        if is_teacher:
            render_rich_text(p_ans, f"*Trả lời:* *{q['ans']}*", base_color=COLOR_CORRECT)
            add_solution_no_table(doc, q["solution"])
        else:
            render_rich_text(p_ans, "*Trả lời:* ....................................................................")

    # 6. PHẦN IV
    add_section_header(
        doc,
        "PHẦN IV. Tự luận.",
        "Thí sinh giải chi tiết các bài toán sau (1,5 điểm)."
    )
    for q in DATA_EXAM["part4"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"])

        if "image" in q:
            img_path = os.path.join(DIR_IMAGES, q["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_path, width=Cm(q.get("image_width", 5.0)))

        if is_teacher:
            add_solution_no_table(doc, q["solution"])
        else:
            add_student_writing_lines(doc, num_lines=6)

    # Dòng kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(10)
    p_end.paragraph_format.space_after = Pt(10)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "---------- *HẾT* ----------", base_bold=True)

    # Lưu file
    suffix = "GiaoVien" if is_teacher else "HocSinh"
    out_file = os.path.join(DIR_SAN_PHAM, f"De_Kiem_Tra_CK1_Hoa_11_Thuan_Hoa_{suffix}.docx")
    try:
        doc.save(out_file)
        print(f"Đã tạo thành công: {out_file}")
    except PermissionError:
        print(f"File {out_file} đang mở trong ứng dụng khác, bỏ qua cập nhật file này.")
    return out_file

if __name__ == "__main__":
    generate_exam_doc(is_teacher=False)
    generate_exam_doc(is_teacher=True)
