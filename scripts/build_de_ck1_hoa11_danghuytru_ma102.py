# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Kì I - Hóa học 11 - Trường THPT Đặng Huy Trứ (TP Huế)
Năm học 2025 - 2026 - Mã đề 102
Quy chuẩn FORM ĐỀ KIỂM TRA của Thầy TRẦN VĂN THẠNH
Tạo 2 phiên bản:
- Bản Học Sinh: De_Kiem_Tra_CK1_Hoa_11_DangHuyTru_Ma102.docx
- Bản Giáo Viên: De_Kiem_Tra_CK1_Hoa_11_DangHuyTru_Ma102_GiaoVien.docx (có BẢNG ĐÁP ÁN RÚT GỌN CẢ TỰ LUẬN & lời giải chi tiết)
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
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_danghuytru_ma102")
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

def add_quick_answer_table_102(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Giáo viên ở đầu tài liệu Mã đề 102.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 102)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I (16 câu chia làm 2 hàng x 8 cột)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (4,0 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=4, cols=9)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    # Hàng 1 & 2: Câu 1 -> 8
    h1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8"]
    a1 = ["Đ/A", "C", "C", "B", "C", "B", "A", "B", "D"]
    # Hàng 3 & 4: Câu 9 -> 16
    h2 = ["Câu", "9", "10", "11", "12", "13", "14", "15", "16"]
    a2 = ["Đ/A", "A", "D", "D", "A", "D", "A", "C", "D"]

    w1_c0 = Cm(1.6)
    w1_c = Cm(2.05)

    all_rows = [h1, a1, h2, a2]
    for r_idx, r_data in enumerate(all_rows):
        row = tbl1.rows[r_idx]
        is_header_row = (r_idx in [0, 2])
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = w1_c0 if c_idx == 0 else w1_c
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if is_header_row:
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

    # 2. PHẦN II (2 câu đúng/sai)
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(5)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM - MỖI CÂU 1,0 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl2.autofit = False

    tbl2_data = [
        ["Câu hỏi", "Nội dung câu hỏi", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Sulfur dioxide (SO~2~) và các vấn đề môi trường", "ĐÚNG", "ĐÚNG", "ĐÚNG", "SAI"],
        ["Câu 2", "Cấu tạo và tính chất các hợp chất hữu cơ (1) - (5)", "SAI", "ĐÚNG", "ĐÚNG", "SAI"]
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

    # 3. PHẦN III (4 câu trả lời ngắn)
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(5)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (1,0 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False

    tbl3_headers = ["Câu", "Câu 1", "Câu 2", "Câu 3", "Câu 4"]
    tbl3_answers = ["Đáp án", "3", "0,15", "3", "11,9"]

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

    # 4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN
    p_tl = doc.add_paragraph()
    p_tl.paragraph_format.space_before = Pt(5)
    p_tl.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl, "*4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN (3,0 ĐIỂM)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl4 = doc.add_table(rows=4, cols=3)
    tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl4.autofit = False

    tbl4_data = [
        ("Câu", "Điểm", "Đáp số cốt lõi & Hướng dẫn chấm"),
        ("Câu 1", "1,0 đ", "a) 2KOH + H~2~SO~4~ → K~2~SO~4~ + 2H~2~O (0,25 đ)\n"
                           "b) 3H~2~SO~4~(loãng) + Fe~2~O~3~ → Fe~2~(SO~4~)~3~ + 3H~2~O (0,25 đ)\n"
                           "c) 4NO~2~ + O~2~ + 2H~2~O → 4HNO~3~ (0,25 đ)\n"
                           "d) NH~4~Cl + NaOH → NaCl + NH~3~↑ + H~2~O (0,25 đ)"),
        ("Câu 2", "1,0 đ", "1) x : y : z = (78,94/12) : (10,53/1) : (10,53/16) = 6,578 : 10,53 : 0,658 ≈ 10 : 16 : 1.\n"
                           "   ⇒ Công thức đơn giản nhất của camphor là C~10~H~16~O. (0,5 đ)\n"
                           "b) CTPT có dạng (C~10~H~16~O)~n~ ⇒ M = 152n.\n"
                           "   Phổ MS có peak ion phân tử [M^+^] tại m/z = 152 ⇒ M = 152 ⇒ n = 1.\n"
                           "   ⇒ Công thức phân tử của camphor là C~10~H~16~O. (0,5 đ)"),
        ("Câu 3", "1,0 đ", "• Phương trình: 5SO~2~ + 2KMnO~4~ + 2H~2~O → K~2~SO~4~ + 2MnSO~4~ + 2H~2~SO~4~\n"
                           "• n(KMnO~4~) = 0,0125 × 5,00.10^-3 = 6,25.10^-5 mol.\n"
                           "• n(SO~2~ trong 10 mL) = (5/2) × 6,25.10^-5 = 1,5625.10^-4 mol.\n"
                           "• n(SO~2~ trong 500 mL) = 1,5625.10^-4 × 50 = 7,8125.10^-3 mol.\n"
                           "• Bảo toàn nguyên tố S: n(S) = n(SO~2~) = 7,8125.10^-3 mol ⇒ m(S) = 7,8125.10^-3 × 32 = 0,25 g.\n"
                           "• Phần trăm khối lượng S trong nhiên liệu: %m(S) = (0,25 / 100,0) × 100% = 0,25% (Đạt giới hạn ≤ 0,3%).")
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


# Cấu trúc dữ liệu Đề thi mã 102
DATA_EXAM = {
    "school_line": "SỞ GD - ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT ĐẶNG HUY TRỨ",
    "exam_title": "KIỂM TRA CUỐI KỲ I – NĂM HỌC 2025-2026",
    "subject_line": "MÔN: HÓA HỌC – LỚP 11",
    "duration": "45 phút (không kể thời gian phát đề)",
    "num_pages": "03",
    "exam_code": "102",
    "atomic_masses": "C: 12, H: 1, O: 16, Ba: 137, S: 32",

    # A. PHẦN TRẮC NGHIỆM (7,0 điểm)
    # Phần I: Trắc nghiệm nhiều lựa chọn (4,0 điểm) - 16 câu
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1. Công thức phân tử chất hữu cơ không thể cho ta biết:",
            "options": [
                "A. Tỉ lệ giữa các nguyên tố trong hợp chất",
                "B. Số lượng các nguyên tố trong hợp chất.",
                "*C. Cấu trúc phân tử hợp chất hữu cơ.*",
                "D. Hàm lượng mỗi nguyên tố trong hợp chất."
            ],
            "solution": "• Công thức phân tử cho ta biết số lượng nguyên tử của mỗi nguyên tố có trong một phân tử hợp chất hữu cơ, từ đó xác định được tỉ lệ nguyên tử giữa các nguyên tố, phân tử khối và hàm lượng phần trăm khối lượng của mỗi nguyên tố.\n• Công thức phân tử KHÔNG cho biết thứ tự liên kết, loại liên kết và sự phân bố trong không gian của các nguyên tử (cấu trúc phân tử cần phải xác định qua công thức cấu tạo).\n*Đáp án đúng: C.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Cho phổ khối lượng của một hợp chất hữu cơ A như hình vẽ:",
            "image": "ms_cau2_clean.png",
            "image_width": 5.5,
            "subprompt": "Hợp chất hữu cơ A có thể là",
            "options": [
                "A. C~4~H~8~O~2~.",
                "B. CH~2~Cl~2~.",
                "*C. C~4~H~6~O~2~.*",
                "D. C~7~H~8~."
            ],
            "solution": "• Trên phổ khối lượng (MS) của chất A, peak ion phân tử [M^+^] có giá trị m/z lớn nhất ở cuối phổ là m/z = 86 ⇒ Phân tử khối của A là M~A~ = 86 g/mol.\n• Kiểm tra phân tử khối của các phương án:\n  - A. C~4~H~8~O~2~: M = 12×4 + 8 + 16×2 = 88.\n  - B. CH~2~Cl~2~: M = 12 + 2 + 35,5×2 = 85.\n  - C. C~4~H~6~O~2~: M = 12×4 + 6 + 16×2 = 86 (thỏa mãn).\n  - D. C~7~H~8~: M = 12×7 + 8 = 92.\n⇒ Hợp chất hữu cơ A có thể là C~4~H~6~O~2~.\n*Đáp án đúng: C.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Dung dịch nào sau đây có pH ~ 7?",
            "options": [
                "A. NaOH.",
                "*B. NaCl.*",
                "C. HNO~3~.",
                "D. H~2~SO~4~."
            ],
            "solution": "• NaCl là muối trung hòa tạo bởi base mạnh (NaOH) và acid mạnh (HCl), các ion Na^+^ và Cl^-^ đều không bị thủy phân trong nước nên dung dịch có môi trường trung tính, pH ~ 7.\n• Dung dịch NaOH là base mạnh có pH > 7; các dung dịch HNO~3~ và H~2~SO~4~ là acid mạnh có pH < 7.\n*Đáp án đúng: B.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Khi phản ứng với kim loại, Sulfur thể hiện tính chất gì?",
            "options": [
                "A. tính khử.",
                "B. tính lưỡng tính.",
                "*C. tính oxi hóa.*",
                "D. vừa tính oxi hóa, vừa tính khử."
            ],
            "solution": "• Khi phản ứng với các kim loại (ở nhiệt độ thích hợp), sulfur nhận 2 electron để tạo ion sulfide S^2-^:\n  Fe + S —(t°)→ FeS\n  2Na + S → Na~2~S\n• Số oxi hóa của sulfur giảm từ 0 xuống -2, do đó sulfur thể hiện tính oxi hóa.\n*Đáp án đúng: C.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5. Phú dưỡng là hiện tượng môi trường nước (ao, hồ, sông) bị dư thừa quá mức các chất dinh dưỡng, chủ yếu là Nitơ (N) và Photpho (P), do nước thải nông nghiệp, sinh hoạt, công nghiệp đổ vào, dẫn đến sự phát triển bùng nổ của tảo (nở hoa) và vi sinh vật, gây thiếu oxy, chết cá, làm suy giảm chất lượng nước và mất cân bằng sinh thái.\nCách nào sau đây không đúng khi xử lí hiện tượng phú dưỡng trong các ao hồ?",
            "options": [
                "A. Sử dụng phân bón hợp lý trong nông nghiệp.",
                "*B. Phun thuốc diệt cỏ để diệt tảo.*",
                "C. Xử lý nước thải trước khi xả ra môi trường.",
                "D. Nạo vét bùn, khơi thông dòng chảy cho ao, hồ, sông."
            ],
            "solution": "• Phun thuốc diệt cỏ vào nguồn nước ao hồ là biện pháp hoàn toàn sai và nguy hiểm, vì các hóa chất độc hại trong thuốc diệt cỏ sẽ tiêu diệt hệ sinh thái thủy sinh, gây độc hại cho cá và các sinh vật khác, ngấm vào chuỗi thức ăn và gây ô nhiễm nguồn nước nghiêm trọng hơn.\n• Các biện pháp đúng: sử dụng phân bón hợp lý (tránh dư thừa N, P), xử lý nước thải sinh hoạt và công nghiệp đạt chuẩn trước khi xả thải, nạo vét bùn đáy và tăng cường lưu thông dòng chảy.\n*Đáp án đúng: B.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6. Để tách các chất lỏng ra khỏi hỗn hợp các chất có nhiệt độ sôi khác nhau, nhằm thu được chất lỏng tinh khiết hơn là phương pháp nào sau đây?",
            "options": [
                "*A. Phương pháp chưng cất.*",
                "B. Phương pháp chiết",
                "C. Sắc kí cột.",
                "D. Phương pháp kết tinh."
            ],
            "solution": "• Phương pháp chưng cất là phương pháp tách các chất lỏng dựa trên sự khác nhau về nhiệt độ sôi của các chất trong hỗn hợp ở một áp suất nhất định.\n• Phương pháp chiết dựa trên độ tan khác nhau trong hai dung môi không hòa tan vào nhau; phương pháp kết tinh dựa trên độ tan thay đổi theo nhiệt độ; sắc kí cột dựa trên sự phân bố khác nhau giữa pha tĩnh và pha động.\n*Đáp án đúng: A.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7. HNO~3~ tinh khiết là chất lỏng không màu, nhưng dung dịch HNO~3~ để lâu thường chuyển sang màu vàng là do",
            "options": [
                "A. khi để lâu thì HNO~3~ bị khử bởi các chất của môi trường",
                "*B. dung dịch HNO~3~ có hoà tan một lượng nhỏ NO~2~.*",
                "C. dung dịch HNO~3~ có tính oxi hoá mạnh.",
                "D. HNO~3~ tan nhiều trong nước."
            ],
            "solution": "• Dung dịch HNO~3~ bị phân hủy một phần khi có ánh sáng hoặc ở nhiệt độ thường theo phản ứng:\n  4HNO~3~ → 4NO~2~↑ + O~2~↑ + 2H~2~O\n• Khí nitrogen dioxide (NO~2~) sinh ra có màu nâu đỏ, hòa tan một lượng nhỏ vào dung dịch HNO~3~ làm cho dung dịch có màu vàng.\n*Đáp án đúng: B.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8. Phản ứng hóa học của các hợp chất hữu cơ có đặc điểm là:",
            "options": [
                "A. thường xảy ra rất nhanh và cho một sản phẩm duy nhất.",
                "B. thường xảy ra rất chậm, nhưng hoàn toàn, không theo một hướng xác định.",
                "C. thường xảy ra rất nhanh, không hoàn toàn, không theo một hướng nhất định.",
                "*D. thường xảy ra chậm, không hoàn toàn, không theo một hướng nhất định.*"
            ],
            "solution": "• Do liên kết cộng hóa trị giữa các nguyên tử trong phân tử chất hữu cơ tương đối bền vững và kém phân cực, nên các phản ứng của hợp chất hữu cơ thường xảy ra chậm (thường cần đun nóng hoặc có xúc tác), không hoàn toàn (thuận nghịch) và sinh ra hỗn hợp nhiều sản phẩm phụ (không theo một hướng nhất định).\n*Đáp án đúng: D.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9. Trong các phát biểu sau đây, phát biểu nào đúng khi một hệ ở trạng thái cân bằng?",
            "options": [
                "*A. Nồng độ của các chất trong hệ không đổi.*",
                "B. Phản ứng thuận đã dừng.",
                "C. Nồng độ chất tham gia và sản phẩm bằng nhau.",
                "D. Phản ứng nghịch đã dừng."
            ],
            "solution": "• Cân bằng hóa học là một cân bằng động: ở trạng thái cân bằng, cả phản ứng thuận và phản ứng nghịch vẫn tiếp tục diễn ra với tốc độ bằng nhau (v~thuận~ = v~nghịch~ > 0), do đó nồng độ của các chất trong hệ phản ứng không đổi theo thời gian.\n*Đáp án đúng: A.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10. Tính chất nào sau đây không phải tính chất của dung dịch sulfuric acid đặc?",
            "options": [
                "A. Tính háo nước.",
                "B. Tính oxi hóa.",
                "C. Tính acid.",
                "*D. Tính khử.*"
            ],
            "solution": "• Trong sulfuric acid (H~2~SO~4~), nguyên tố sulfur có số oxi hóa cao nhất là +6, do đó H~2~SO~4~ chỉ có thể nhận electron để giảm số oxi hóa xuống (+4, 0, -2) mà không thể nhường thêm electron ⇒ H~2~SO~4~ đặc có tính oxi hóa rất mạnh, tính acid mạnh và tính háo nước mạnh; KHÔNG có tính khử.\n*Đáp án đúng: D.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11. Phát biểu nào sau đây đúng khi nói về sự điện li?",
            "options": [
                "A. Sự điện li thực chất là quá trình oxi hoá – khử.",
                "B. Sự điện li quá trình hòa tan một chất vào nước tạo thành dung dịch.",
                "C. Sự điện li quá trình phân li một chất dưới tác dụng của dòng điện.",
                "*D. Sự điện li là quá trình phân li một chất trong nước thành ion.*"
            ],
            "solution": "• Theo định nghĩa chuẩn: Sự điện li là quá trình phân li các chất khi tan trong nước thành các ion mang điện tích.\n*Đáp án đúng: D.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12. Xác định công thức cấu tạo thu gọn của hợp chất sau:",
            "image": "ctct_cau12_clean.png",
            "image_width": 2.6,
            "options": [
                "*A. CH~3~CH~2~CH~2~COOH.*",
                "B. CH~3~CH~2~CH~2~OH.",
                "C. CH~3~CH~2~CHOHCHO.",
                "D. CH~3~CH~2~COOH."
            ],
            "solution": "• Quan sát công thức cấu tạo đầy đủ của hợp chất:\n  - Nguyên tử C thứ 1 (bên trái) liên kết với 3 nguyên tử H ⇒ nhóm CH~3~-.\n  - Nguyên tử C thứ 2 liên kết với 2 nguyên tử H ⇒ nhóm -CH~2~-.\n  - Nguyên tử C thứ 3 liên kết với 2 nguyên tử H ⇒ nhóm -CH~2~-.\n  - Nguyên tử C thứ 4 liên kết đôi với O và liên kết đơn với nhóm -OH ⇒ nhóm carboxyl -COOH.\n⇒ Công thức cấu tạo thu gọn là: CH~3~CH~2~CH~2~COOH (butanoic acid).\n*Đáp án đúng: A.*"
        },
        {
            "id": 13,
            "prompt": "Câu 13. Cho chất X có công thức cấu tạo thu gọn như sau: CH~3~–COOH. Phân tử chất X chứa nhóm chức gì?",
            "options": [
                "A. -COO-",
                "B. -CO-",
                "C. -CHO",
                "*D. -COOH*"
            ],
            "solution": "• Phân tử acetic acid CH~3~COOH chứa nhóm chức carboxyl (-COOH) đặc trưng cho hợp chất carboxylic acid.\n*Đáp án đúng: D.*"
        },
        {
            "id": 14,
            "prompt": "Câu 14. Liên kết hoá học trong phân tử NH~3~ là liên kết",
            "options": [
                "*A. cộng hoá trị có cực.*",
                "B. phối trí.",
                "C. cộng hoá trị không cực.",
                "D. ion."
            ],
            "solution": "• Trong phân tử ammonia (NH~3~), nguyên tử N có độ âm điện lớn hơn nguyên tử H (χ~N~ = 3,04 > χ~H~ = 2,20), cặp electron dùng chung giữa N và H bị lệch về phía nguyên tử nitrogen ⇒ Liên kết N-H là liên kết cộng hóa trị có cực.\n*Đáp án đúng: A.*"
        },
        {
            "id": 15,
            "prompt": "Câu 15. Cho dung dịch HNO~3~ tác dụng với các chất sau: NH~3~, CaCO~3~, Ag, NaOH. Theo Bronsted, số phản ứng trong đó HNO~3~ đóng vai trò acid là",
            "options": [
                "A. 1.",
                "B. 4.",
                "*C. 3.*",
                "D. 2."
            ],
            "solution": "• Theo thuyết acid - base của Brønsted - Lowry, acid là chất có khả năng cho proton (H^+^).\n• Xét các phản ứng của dung dịch HNO~3~:\n  (1) HNO~3~ + NH~3~ → NH~4~^+^ + NO~3~^-^ : HNO~3~ nhường H^+^ cho NH~3~ ⇒ HNO~3~ là acid.\n  (2) 2HNO~3~ + CaCO~3~ → Ca(NO~3~)~2~ + CO~2~ + H~2~O : HNO~3~ nhường H^+^ cho CO~3~^2-^ ⇒ HNO~3~ là acid.\n  (3) 4HNO~3~ + Ag → AgNO~3~ + NO~2~ + 2H~2~O : Đây là phản ứng oxi hóa - khử, HNO~3~ đóng vai trò chất oxi hóa và môi trường (không phải phản ứng acid - base theo Brønsted).\n  (4) HNO~3~ + NaOH → NaNO~3~ + H~2~O : HNO~3~ nhường H^+^ cho OH^-^ ⇒ HNO~3~ là acid.\n⇒ Có 3 phản ứng HNO~3~ đóng vai trò là acid theo Brønsted (phản ứng với NH~3~, CaCO~3~, NaOH).\n*Đáp án đúng: C.*"
        },
        {
            "id": 16,
            "prompt": "Câu 16. Glucose là hợp chất hữu cơ có nhiều trong các loại quả chín, đặc biệt là quả nho. Công thức phân tử của glucose là C~6~H~12~O~6~. Công thức đơn giản nhất của glucose là",
            "options": [
                "A. CHO~2~.",
                "B. CH~3~O.",
                "C. C~3~H~4~O~3~.",
                "*D. CH~2~O.*"
            ],
            "solution": "• Công thức phân tử của glucose là C~6~H~12~O~6~.\n• Tỉ lệ số nguyên tử các nguyên tố: C : H : O = 6 : 12 : 6 = 1 : 2 : 1.\n⇒ Công thức đơn giản nhất của glucose là CH~2~O.\n*Đáp án đúng: D.*"
        }
    ],

    # Phần II: Trắc nghiệm dạng Đúng/Sai (2,0 điểm) - 2 câu
    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1. Sulfur dioxide (SO~2~) là một chất khí không màu, có mùi hăng đặc trưng, nặng hơn không khí và có mặt trong không khí từ các nguồn khác nhau.",
            "statements": [
                ("a. Khí SO~2~ là một trong những tác nhân gây ô nhiễm môi trường và viêm đường hô hấp ở người, ...", "ĐÚNG"),
                ("b. Để giảm thải SO~2~ vào khí quyển cần xử lí khí thải công nghiệp, sử dụng các nhiên liệu sinh học và năng lượng tái tạo thay cho nhiên liệu hóa thạch.", "ĐÚNG"),
                ("c. Cùng với khí NO~2~, khí SO~2~ là nguyên nhân chính gây nên hiện tượng mưa acid.", "ĐÚNG"),
                ("d. Nguồn phát sinh SO~2~ hiện tượng sấm sét trong khí quyển tạo nên.", "SAI")
            ],
            "solution": "• Ý a: ĐÚNG. Khí SO~2~ là một trong các khí độc gây ô nhiễm không khí phổ biến, kích ứng mạnh niêm mạc mắt và đường hô hấp, gây khó thở, viêm phế quản và các bệnh phổi mãn tính.\n• Ý b: ĐÚNG. Để hạn chế phát thải SO~2~, cần xử lý khí thải bằng các phương pháp hấp thụ (sữa vôi, đá vôi...), đồng thời chuyển đổi từ nhiên liệu hóa thạch chứa sulfur (than đá, dầu mỏ) sang các nguồn năng lượng sạch, năng lượng tái tạo (mặt trời, gió, sinh học).\n• Ý c: ĐÚNG. Khí SO~2~ và NO~2~ trong khí quyển bị oxi hóa bởi oxy và hơi nước dưới xúc tác của ion kim loại hoặc ánh sáng mặt trời tạo thành sulfuric acid (H~2~SO~4~) và nitric acid (HNO~3~), hòa tan trong nước mưa gây ra mưa acid.\n• Ý d: SAI. Hiện tượng sấm sét tạo nhiệt độ rất cao (hàng nghìn độ C) kích thích phản ứng giữa N~2~ và O~2~ tạo khí NO (N~2~ + O~2~ ⇌ 2NO), đây là nguồn tự nhiên sinh ra các oxide của nitrogen (NO~x~), KHÔNG PHẢI nguồn phát sinh SO~2~. Nguồn tự nhiên phát sinh SO~2~ chủ yếu là do núi lửa phun trào."
        },
        {
            "id": 2,
            "prompt": "Câu 2. Cho một số công thức cấu tạo của các chất như sau:",
            "image": "structures_p2_c2_clean.png",
            "image_width": 5.8,
            "statements": [
                ("a. Công thức cấu tạo (2) và (4) biểu diễn cùng một chất.", "SAI"),
                ("b. Chất (2) và (4) có tính chất hoá học tương tự nhau.", "ĐÚNG"),
                ("c. Chất (4) và (5) là đồng phân của nhau.", "ĐÚNG"),
                ("d. Chất (1) và (3) là hydrocarbon mạch hở không phân nhánh..", "SAI")
            ],
            "solution": "• Xác định công thức cấu tạo và phân tử của từng chất:\n  (1): butane (CH~3~-CH~2~-CH~2~-CH~3~, C~4~H~10~) - hydrocarbon mạch hở, không phân nhánh.\n  (2): propan-1-ol (CH~3~-CH~2~-CH~2~-OH, C~3~H~8~O) - alcohol no, đơn chức, mạch hở.\n  (3): isobutane hay 2-methylpropane (CH~3~-CH(CH~3~)-CH~3~, C~4~H~10~) - hydrocarbon mạch hở, CÓ PHÂN NHÁNH.\n  (4): ethanol (CH~3~-CH~2~-OH, C~2~H~6~O) - alcohol no, đơn chức, mạch hở.\n  (5): dimethyl ether (CH~3~-O-CH~3~, C~2~H~6~O) - ether no, đơn chức, mạch hở.\n\n• Ý a: SAI. Chất (2) có 3 carbon (propan-1-ol) và chất (4) có 2 carbon (ethanol) là hai chất khác nhau (đồng đẳng của nhau), không phải cùng một chất.\n• Ý b: ĐÚNG. Chất (2) và (4) đều là các alcohol no, đơn chức, mạch hở thuộc cùng một dãy đồng đẳng nên có tính chất hóa học tương tự nhau (tác dụng với Na giải phóng H~2~, phản ứng oxi hóa tạo aldehyde/acid...).\n• Ý c: ĐÚNG. Chất (4) và (5) có cùng công thức phân tử là C~2~H~6~O nhưng khác nhau về công thức cấu tạo (ethanol chứa nhóm -OH, dimethyl ether chứa liên kết -O-), nên chúng là đồng phân cấu tạo (khác loại nhóm chức) của nhau.\n• Ý d: SAI. Chất (1) là hydrocarbon mạch hở không phân nhánh, nhưng chất (3) là hydrocarbon mạch hở CÓ PHÂN NHÁNH (nhánh methyl tại vị trí carbon số 2)."
        }
    ],

    # Phần III: Trắc nghiệm trả lời ngắn (1,0 điểm) - 4 câu
    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1. Cho 6 hợp chất hữu cơ có công thức phân tử lần lượt là: C~2~H~6~O, C~2~H~4~O~2~, C~3~H~6~O~3~, C~3~H~6~O~2~, C~4~H~8~O~2~, C~6~H~12~O~6~. Có bao nhiêu hợp chất hữu cơ không có công thức đơn giản nhất là CH~2~O?",
            "answer": "3",
            "solution": "• Tìm công thức đơn giản nhất của từng hợp chất:\n  1) C~2~H~6~O: C : H : O = 2 : 6 : 1 ⇒ CTĐGN là C~2~H~6~O (khác CH~2~O) → Thỏa mãn (1).\n  2) C~2~H~4~O~2~: C : H : O = 2 : 4 : 2 = 1 : 2 : 1 ⇒ CTĐGN là CH~2~O.\n  3) C~3~H~6~O~3~: C : H : O = 3 : 6 : 3 = 1 : 2 : 1 ⇒ CTĐGN là CH~2~O.\n  4) C~3~H~6~O~2~: C : H : O = 3 : 6 : 2 ⇒ CTĐGN là C~3~H~6~O~2~ (khác CH~2~O) → Thỏa mãn (2).\n  5) C~4~H~8~O~2~: C : H : O = 4 : 8 : 2 = 2 : 4 : 1 ⇒ CTĐGN là C~2~H~4~O (khác CH~2~O) → Thỏa mãn (3).\n  6) C~6~H~12~O~6~: C : H : O = 6 : 12 : 6 = 1 : 2 : 1 ⇒ CTĐGN là CH~2~O.\n⇒ Có 3 hợp chất không có công thức đơn giản nhất là CH~2~O (gồm C~2~H~6~O, C~3~H~6~O~2~, C~4~H~8~O~2~).\n*Đáp án: 3*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Trộn 200 mL dung dịch gồm HCl 0,1M và H~2~SO~4~ 0,05M vào 300 mL dung dịch Ba(OH)~2~ có nồng độ a mol/L thu được có pH = 13. Giá trị của a là (Kết quả làm tròn đến hàng phần trăm).",
            "answer": "0,15",
            "solution": "• Thể tích dung dịch sau khi trộn: V = 200 + 300 = 500 mL = 0,5 L.\n• Dung dịch sau phản ứng có pH = 13 > 7 ⇒ môi trường base (OH^-^ dư).\n  pOH = 14 - pH = 14 - 13 = 1 ⇒ [OH^-^] dư = 10^-1 = 0,1 M.\n• Số mol OH^-^ dư: n(OH^-^ dư) = [OH^-^] dư × V = 0,1 × 0,5 = 0,05 mol.\n• Dung dịch acid ban đầu:\n  n(HCl) = 0,2 × 0,1 = 0,02 mol.\n  n(H~2~SO~4~) = 0,2 × 0,05 = 0,01 mol.\n  ⇒ Tổng số mol H^+^ ban đầu: n(H^+^) = n(HCl) + 2×n(H~2~SO~4~) = 0,02 + 2×0,01 = 0,04 mol.\n• Dung dịch base Ba(OH)~2~ ban đầu:\n  n(Ba(OH)~2~) = 0,3 × a mol ⇒ n(OH^-^ ban đầu) = 2 × 0,3a = 0,6a mol.\n• Phản ứng trung hòa: H^+^ + OH^-^ → H~2~O\n  n(OH^-^ phản ứng) = n(H^+^) = 0,04 mol.\n• Ta có:\n  n(OH^-^ dư) = n(OH^-^ ban đầu) - n(OH^-^ phản ứng)\n  ⇔ 0,05 = 0,6a - 0,04\n  ⇔ 0,6a = 0,09\n  ⇔ a = 0,09 / 0,6 = 0,15 M.\n*Đáp án: 0,15*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Các cách làm sau đây, cách nào sử dụng phương pháp chiết?\n(1) Giã lá cây chàm, cho vào nước, lọc lấy dung dịch màu để nhuộm sợi, vải.\n(2) Dùng nước sôi để pha trà, cà phê.\n(3) Ngâm rượu thuốc.\n(4) Làm đường cát, đường phèn từ nước mía.",
            "answer": "3",
            "solution": "• Phương pháp chiết (cụ thể là chiết lỏng - rắn) là phương pháp dùng một dung môi thích hợp để hòa tan và tách lấy chất cần tách ra khỏi hỗn hợp chất rắn:\n  + (1) Giã lá chàm rồi cho vào nước để hòa tan chất màu vào nước là phương pháp chiết lỏng - rắn.\n  + (2) Dùng nước sôi để chiết các chất tan (caffeine, tannin, tinh dầu thơm...) từ lá chè, bột cà phê vào nước là phương pháp chiết lỏng - rắn.\n  + (3) Ngâm các loại dược liệu trong cồn/rượu để hòa tan các hoạt chất có lợi vào rượu là phương pháp chiết lỏng - rắn.\n  + (4) Đun cô đặc nước mía để đường kết tinh tạo thành đường cát, đường phèn là phương pháp kết tinh (không phải chiết).\n⇒ Có 3 cách sử dụng phương pháp chiết là: (1), (2), (3).\n*Đáp án: 3*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Có một loại quặng pyrite chứa 84% FeS~2~. Nếu mỗi ngày nhà máy sản xuất 10 tấn sulfuric acid 98% thì cần m tấn quặng pyrite trên và biết hiệu suất của cả quá trình sản xuất H~2~SO~4~ là 60%. Giá trị của m là (Kết quả làm tròn đến hàng phần mười).",
            "answer": "11,9",
            "solution": "• Khối lượng H~2~SO~4~ nguyên chất cần sản xuất mỗi ngày:\n  m(H~2~SO~4~) = 10 × 98% = 9,8 tấn.\n• Số mol H~2~SO~4~: n(H~2~SO~4~) = 9,8 / 98 = 0,1 (triệu mol hay Mmol).\n• Sơ đồ bảo toàn nguyên tố sulfur trong toàn bộ quá trình:\n  FeS~2~ ———→ 2 H~2~SO~4~\n  1 mol            2 mol\n• Số mol FeS~2~ cần dùng theo lý thuyết:\n  n(FeS~2~ lt) = 1/2 × n(H~2~SO~4~) = 0,1 / 2 = 0,05 (Mmol).\n• Do hiệu suất của cả quá trình là H = 60%, lượng FeS~2~ thực tế cần dùng là:\n  n(FeS~2~ tt) = n(FeS~2~ lt) / H = 0,05 / 0,60 = 1/12 (Mmol).\n• Khối lượng FeS~2~ nguyên chất thực tế cần dùng:\n  M(FeS~2~) = 56 + 32×2 = 120 g/mol.\n  m(FeS~2~ tt) = (1/12) × 120 = 10 tấn.\n• Quặng pyrite chỉ chứa 84% FeS~2~, do đó khối lượng quặng pyrite cần dùng là:\n  m = m(FeS~2~ tt) / 84% = 10 / 0,84 ≈ 11,90476... tấn.\n• Làm tròn đến hàng phần mười theo yêu cầu đề bài: m ≈ 11,9 tấn.\n*Đáp án: 11,9*"
        }
    ],

    # B. PHẦN TỰ LUẬN (3,0 điểm)
    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1 (1,0 điểm): Hoàn thành phương trình hóa học của các phản ứng sau?\n"
                      "a) KOH + H~2~SO~4~ →\n"
                      "b) H~2~SO~4~ loãng + Fe~2~O~3~ →\n"
                      "c) NO~2~ + O~2~ + H~2~O →\n"
                      "d) NH~4~Cl + NaOH →",
            "num_lines": 8,
            "solution": "Phương trình hóa học của các phản ứng xảy ra:\n"
                       "a) 2KOH + H~2~SO~4~ → K~2~SO~4~ + 2H~2~O\n"
                       "b) 3H~2~SO~4~(loãng) + Fe~2~O~3~ → Fe~2~(SO~4~)~3~ + 3H~2~O\n"
                       "c) 4NO~2~ + O~2~ + 2H~2~O → 4HNO~3~\n"
                       "d) NH~4~Cl + NaOH —(t°)→ NaCl + NH~3~↑ + H~2~O\n"
                       "   (ở điều kiện thường: NH~4~Cl + NaOH → NaCl + NH~3~.H~2~O hoặc NaCl + NH~3~↑ + H~2~O)"
        },
        {
            "id": 2,
            "prompt": "Câu 2 (1,0 điểm): Camphor (có trong cây long não) là một chất rắn kết tinh màu trắng hay trong suốt giống như sáp với mùi thơm đặc trưng, thường dùng trong y học. Phần trăm khối lượng các nguyên tố trong camphor lần lượt là 78,94% carbon, 10,53% hydrogen và 10,53% oxygen.\n"
                      "1) Lập công thức đơn giản nhất của camphor.\n"
                      "b) Xác định công thức phân tử của camphor biết phổ MS của camphor được cho như hình bên dưới.",
            "image": "ms_camphor_clean.png",
            "image_width": 3.8,
            "num_lines": 10,
            "solution": "1) Lập công thức đơn giản nhất của camphor:\n"
                       "• Đặt công thức đơn giản nhất của camphor là C~x~H~y~O~z~.\n"
                       "• Tỉ lệ số mol giữa các nguyên tố:\n"
                       "  x : y : z = (%C / 12) : (%H / 1) : (%O / 16)\n"
                       "            = (78,94 / 12) : (10,53 / 1) : (10,53 / 16)\n"
                       "            = 6,578 : 10,53 : 0,658\n"
                       "• Chia tất cả cho số nhỏ nhất (0,658):\n"
                       "  x : y : z ≈ (6,578 / 0,658) : (10,53 / 0,658) : (0,658 / 0,658)\n"
                       "            ≈ 10 : 16 : 1\n"
                       "⇒ Công thức đơn giản nhất của camphor là C~10~H~16~O.\n\n"
                       "b) Xác định công thức phân tử của camphor:\n"
                       "• Công thức phân tử của camphor có dạng: (C~10~H~16~O)~n~ (với n là số nguyên dương).\n"
                       "• Phân tử khối tương ứng: M = (12×10 + 1×16 + 16)×n = 152n.\n"
                       "• Quan sát phổ khối lượng (MS) của camphor: peak ion phân tử [M^+^] có giá trị m/z lớn nhất là 152 ⇒ Phân tử khối của camphor M = 152 g/mol.\n"
                       "• Ta có: 152n = 152 ⇒ n = 1.\n"
                       "⇒ Công thức phân tử của camphor là C~10~H~16~O."
        },
        {
            "id": 3,
            "prompt": "Câu 3 (1,0 điểm): Hàm lượng cho phép của sulfur trong nhiên liệu là 0,3% về khối lượng. Để xác định hàm lượng sulfur trong một loại nhiên liệu người ta lấy 100,0 gam nhiên liệu đó và đốt cháy hoàn toàn. Khí tạo thành chỉ chứa carbon dioxide, sulfur dioxide và hơi nước được dẫn vào nước tạo ra 500,0 ml dung dịch. Biết rằng tất cả sulfur dioxide đã tan vào dung dịch. Lấy 10,0 ml dung dịch này cho tác dụng với dung dịch KMnO~4~ 5,00.10^-3 mol/L thì thể tích dung dịch KMnO~4~ cần dùng là 12,5 mL. Tính phần trăm khối lượng của sulfur trong nhiên liệu trên?\n"
                      "(Cho NTK các nguyên tố: C: 12, H: 1, O: 16, Ba: 137, S: 32)",
            "num_lines": 12,
            "solution": "• Phương trình hóa học của phản ứng chuẩn độ dung dịch SO~2~ bằng dung dịch KMnO~4~:\n"
                       "  5SO~2~ + 2KMnO~4~ + 2H~2~O → K~2~SO~4~ + 2MnSO~4~ + 2H~2~SO~4~\n"
                       "• Số mol KMnO~4~ đã phản ứng trong 10,0 mL dung dịch SO~2~:\n"
                       "  n(KMnO~4~) = V × C~M~ = (12,5 / 1000) × 5,00.10^-3 = 6,25.10^-5 mol.\n"
                       "• Theo phương trình phản ứng hóa học:\n"
                       "  n(SO~2~ trong 10,0 mL) = (5/2) × n(KMnO~4~) = 2,5 × 6,25.10^-5 = 1,5625.10^-4 mol.\n"
                       "• Thể tích toàn bộ dung dịch hấp thụ khí là 500,0 mL (gấp 500,0 / 10,0 = 50 lần):\n"
                       "  n(SO~2~ tổng) = 1,5625.10^-4 × (500,0 / 10,0) = 7,8125.10^-3 mol.\n"
                       "• Theo định luật bảo toàn nguyên tố sulfur:\n"
                       "  n(S trong nhiên liệu) = n(SO~2~ tổng) = 7,8125.10^-3 mol.\n"
                       "• Khối lượng nguyên tố sulfur có trong 100,0 gam nhiên liệu:\n"
                       "  m(S) = n(S) × M(S) = 7,8125.10^-3 × 32 = 0,25 gam.\n"
                       "• Phần trăm khối lượng của sulfur trong loại nhiên liệu trên là:\n"
                       "  %m(S) = (m(S) / m(nhiên liệu)) × 100% = (0,25 / 100,0) × 100% = 0,25%.\n"
                       "• Nhận xét: Hàm lượng sulfur thực tế trong nhiên liệu là 0,25% ≤ 0,3% nên đạt tiêu chuẩn an toàn cho phép."
        }
    ]
}

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

    # Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Bản Giáo Viên
    if is_teacher:
        add_quick_answer_table_102(doc)

    # A. PHẦN TRẮC NGHIỆM (7,0 điểm)
    add_section_header(
        doc,
        title="A. TRẮC NGHIỆM (7,0 điểm)",
        subtitle="Phần I: Trắc nghiệm nhiều lựa chọn (4,0 điểm). Học sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi học sinh chỉ chọn một phương án."
    )

    for item in DATA_EXAM["part1"]:
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
                p_img.paragraph_format.space_after = Pt(2)
                w = item.get("image_width", 4.0)
                p_img.add_run().add_picture(img_path, width=Inches(w))

        if "subprompt" in item:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(1)
            p_sub.paragraph_format.space_after = Pt(1)
            render_rich_text(p_sub, item["subprompt"], base_size=10.5)

        render_options_no_table(doc, item["options"], is_teacher=is_teacher)
        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # Phần II: Trắc nghiệm dạng Đúng/Sai (2,0 điểm)
    add_section_header(
        doc,
        title="Phần II: Trắc nghiệm dạng Đúng/Sai (2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai."
    )

    for item in DATA_EXAM["part2"]:
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
                p_img.paragraph_format.space_after = Pt(2)
                w = item.get("image_width", 4.5)
                p_img.add_run().add_picture(img_path, width=Inches(w))

        for stmt, ans in item["statements"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt, ans=ans, is_teacher=is_teacher, base_size=10.5)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # Phần III: Trắc nghiệm trả lời ngắn (1,0 điểm)
    add_section_header(
        doc,
        title="Phần III: Trắc nghiệm trả lời ngắn (1,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4."
    )

    for item in DATA_EXAM["part3"]:
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
                p_img.paragraph_format.space_after = Pt(2)
                w = item.get("image_width", 3.5)
                p_img.add_run().add_picture(img_path, width=Inches(w))

        if "subprompt" in item:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(1)
            p_sub.paragraph_format.space_after = Pt(1)
            render_rich_text(p_sub, item["subprompt"], base_size=10.5)

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

    # B. PHẦN TỰ LUẬN (3,0 điểm)
    add_section_header(
        doc,
        title="B. TỰ LUẬN (3,0 điểm)",
        subtitle="Học sinh trình bày lời giải chi tiết vào phần giấy thi."
    )

    for item in DATA_EXAM["part4"]:
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
                p_img.paragraph_format.space_after = Pt(2)
                w = item.get("image_width", 3.8)
                p_img.add_run().add_picture(img_path, width=Inches(w))

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])
        elif not is_teacher:
            num_lines = item.get("num_lines", 6)
            add_student_writing_lines(doc, num_lines=num_lines)

    # Dòng kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(0)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "---------- *HẾT* ----------", base_size=10.5, base_bold=True)

    # Đặt tên file xuất
    suffix = "_GiaoVien.docx" if is_teacher else ".docx"
    filename = f"De_Kiem_Tra_CK1_Hoa_11_DangHuyTru_Ma102{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo thành công: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
