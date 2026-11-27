# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Học Kì I - Hóa học 11 - Trường THPT Nguyễn Chí Thanh (TP Huế)
Năm học 2025 - 2026 - Mã đề 1101
Quy chuẩn FORM ĐỀ KIỂM TRA của Thầy TRẦN VĂN THẠNH
Tạo 2 phiên bản:
- Bản Học Sinh: De_Kiem_Tra_CK1_Hoa_11_NguyenChiThanh_Ma1101.docx
- Bản Giáo Viên: De_Kiem_Tra_CK1_Hoa_11_NguyenChiThanh_Ma1101_GiaoVien.docx (có BẢNG ĐÁP ÁN RÚT GỌN CẢ TỰ LUẬN & lời giải chi tiết)
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
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_nguyenchithanh_ma1101")
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

def add_quick_answer_table_1101(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Giáo viên ở đầu tài liệu Mã đề 1101.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 1101)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I (16 câu chia làm 2 hàng x 8 cột)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (4,0 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=4, cols=9)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    h1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8"]
    a1 = ["Đ/A", "D", "A", "D", "C", "B", "B", "A", "B"]
    h2 = ["Câu", "9", "10", "11", "12", "13", "14", "15", "16"]
    a2 = ["Đ/A", "D", "D", "C", "D", "B", "C", "B", "C"]

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
        ["Câu 1", "Chuẩn độ dung dịch H~2~SO~4~ bằng dung dịch NaOH", "ĐÚNG", "ĐÚNG", "SAI", "ĐÚNG"],
        ["Câu 2", "Đốt cháy hoàn toàn hợp chất hữu cơ X (C~4~H~10~)", "SAI", "ĐÚNG", "SAI", "SAI"]
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
    tbl3_answers = ["Đáp án", "6,72", "3", "3", "4"]

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
        ("Câu 1", "0,5 đ", "• Butane: CTCT thu gọn CH~3~-CH~2~-CH~2~-CH~3~; Khung mạch 4 đỉnh dích dắc.\n"
                           "• Isobutane: CTCT thu gọn CH~3~-CH(CH~3~)-CH~3~; Khung chữ Y 3 nhánh."),
        ("Câu 2", "1,0 đ", "a) Peak ion phân tử [M^+^] tại m/z = 194 ⇒ Phân tử khối M = 194 g/mol. (0,25 đ)\n"
                           "b) %O = 100% - 49,48% - 5,15% - 28,87% = 16,50%.\n"
                           "   Số nguyên tử: C = 8; H = 10; N = 4; O = 2 ⇒ CTPT: C~8~H~10~N~4~O~2~. (0,75 đ)"),
        ("Câu 3", "1,5 đ", "a) 4 phương trình sản xuất H~2~SO~4~ từ FeS~2~: (1) Đốt quặng; (2) Oxi hóa SO~2~; (3) Hấp thụ vào H~2~SO~4~ tạo oleum; (4) Pha loãng oleum. (0,5 đ)\n"
                           "b) Trong 50,7 g X: n(S) = n(BaSO~4~) = 139,8 / 233 = 0,6 mol ⇒ (98 + 80n)/(n + 1) = 84,5 ⇒ n = 3 (oleum H~2~SO~4~.3SO~3~).\n"
                           "   Trong 1000 g dd H~2~SO~4~ 98%: n(H~2~SO~4~) = 10 mol; n(H~2~O) = 10/9 mol.\n"
                           "   n(SO~3~ tổng) = 10/9 + 3 × (10 + 10/9) = 310/9 mol ⇒ n(FeS~2~) = 155/9 mol.\n"
                           "   m(FeS~2~) = (155/9) × 120 = 2066,67 g ⇒ m ≈ 2,067 kg (hoặc 2,07 kg). (1,0 đ)")
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


# Dữ liệu Đề thi mã 1101
DATA_EXAM = {
    "school_line": "SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HUẾ\nTRƯỜNG THPT NGUYỄN CHÍ THANH",
    "exam_title": "ĐỀ KIỂM TRA CUỐI HỌC KÌ 1 – NĂM HỌC 2025-2026",
    "subject_line": "MÔN: HÓA HỌC - LỚP 11",
    "duration": "45 phút (không kể thời gian phát đề)",
    "num_pages": "04",
    "exam_code": "1101",
    "atomic_masses": "Cu = 64; H = 1; O = 16; N = 14; Fe = 56; Al = 27; S = 32; C = 12; Ba = 137",

    # A. PHẦN TRẮC NGHIỆM (7,0 điểm)
    # Phần 1: Trắc nghiệm nhiều lựa chọn (4,0 điểm) - 16 câu
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1. Cho cân bằng hóa học: N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)   Δ~r~H = -92 kJ\nCân bằng trên chuyển dịch theo chiều thuận khi",
            "options": [
                "A. tăng nhiệt độ của hệ phản ứng.",
                "B. thêm chất xúc tác vào hệ phản ứng.",
                "C. tăng nồng độ NH~3~.",
                "*D. tăng áp suất của hệ phản ứng.*"
            ],
            "solution": "• Phản ứng thuận làm giảm số mol khí (từ 4 mol N~2~ + H~2~ thành 2 mol NH~3~), do đó khi tăng áp suất, cân bằng sẽ chuyển dịch theo chiều làm giảm áp suất (chiều thuận).\n• Tăng nhiệt độ cân bằng dịch theo chiều thu nhiệt (chiều nghịch); thêm chất xúc tác không làm chuyển dịch cân bằng; tăng nồng độ NH~3~ cân bằng dịch chuyển theo chiều làm giảm nồng độ NH~3~ (chiều nghịch).\n*Đáp án đúng: D.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Hỗn hợp X gồm các alkane: pentane (t~s~ = 36,1 °C), heptane (t~s~ = 98,4 °C), octane (t~s~ = 125,7 °C) và nonane (t~s~ = 150,8 °C). Có thể tách riêng các chất đó một cách thuận lợi bằng phương pháp nào sau đây?",
            "options": [
                "*A. Chưng cất.*",
                "B. Chiết.",
                "C. Kết tinh.",
                "D. Sắc kí."
            ],
            "solution": "• Các alkane trong hỗn hợp X đều là các chất lỏng hòa tan vào nhau và có nhiệt độ sôi (t~s~) khác nhau rõ rệt, do đó phương pháp thuận lợi và tối ưu nhất để tách riêng từng chất là phương pháp chưng cất phân đoạn.\n*Đáp án đúng: A.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Để tạo độ xốp cho một số loại bánh, có thể dùng chất nào sau đây?",
            "options": [
                "A. NaCl.",
                "B. (NH~4~)~3~PO~4~.",
                "C. CaCO~3~.",
                "*D. NH~4~HCO~3~.*"
            ],
            "solution": "• Muối ammonium hydrogencarbonate (NH~4~HCO~3~, còn gọi là bột nở) dễ bị nhiệt phân hủy khi nướng bánh tạo thành các chất khí: NH~4~HCO~3~ —(t°)→ NH~3~↑ + CO~2~↑ + H~2~O↑. Các bọt khí này thoát ra làm cho bánh nở xốp và mềm.\n*Đáp án đúng: D.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Sulfur là chất oxi hóa trong phản ứng nào sau đây?",
            "options": [
                "A. S + 3F~2~ → SF~6~",
                "B. S + O~2~ —(t°)→ SO~2~",
                "*C. S + Fe —(t°)→ FeS*",
                "D. S + 2H~2~SO~4~ —(t°)→ 3SO~2~ + 2H~2~O"
            ],
            "solution": "• Trong phản ứng C: S + Fe —(t°)→ FeS, số oxi hóa của sulfur giảm từ 0 xuống -2 ⇒ Sulfur đóng vai trò là chất oxi hóa.\n• Trong các phản ứng A, B, D: số oxi hóa của sulfur tăng lên (+4, +6) nên sulfur đóng vai trò là chất khử.\n*Đáp án đúng: C.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5. Nhóm chức là",
            "options": [
                "A. Là một nguyên tử bất kì trong phân tử chất hữu cơ",
                "*B. Là một nguyên tử (hoặc nhóm nguyên tử) gây ra những phản ứng hoá học đặc trưng của phân tử hợp chất hữu cơ*",
                "C. Là một nhóm nguyên tử khác biệt trong chất hữu cơ",
                "D. Là một nhóm nguyên tử có cấu trúc không gian đặc biệt mà trong đó các nguyên tử liên kết với nhau không theo quy tắc hoá trị nào"
            ],
            "solution": "• Định nghĩa chuẩn: Nhóm chức là nguyên tử hoặc nhóm nguyên tử gây ra những tính chất và phản ứng hóa học đặc trưng của phân tử hợp chất hữu cơ.\n*Đáp án đúng: B.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6. Chất nào sau đây là chất điện li?",
            "options": [
                "A. saccharose",
                "*B. sulfuric acid*",
                "C. glucose",
                "D. ethyl alcohol"
            ],
            "solution": "• Sulfuric acid (H~2~SO~4~) khi tan trong nước phân li hoàn toàn thành các ion (H^+^ và SO~4~^2-^) nên là chất điện li mạnh.\n• Saccharose, glucose và ethyl alcohol khi tan trong nước chỉ phân tán dưới dạng phân tử trung hòa, không phân li ra ion nên là chất không điện li.\n*Đáp án đúng: B.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7. Hợp chất X có công thức khung phân tử như sau:",
            "image": "khung_c7_clean.png",
            "image_width": 2.2,
            "subprompt": "Tổng số nguyên tử hydrogen trong X là",
            "options": [
                "*A. 10*",
                "B. 12",
                "C. 6",
                "D. 8"
            ],
            "solution": "• Quan sát công thức khung phân tử của X:\n  - Có liên kết đôi C=C ở đầu mạch gắn với 1 nhánh methyl: CH~2~=C(CH~3~)-\n  - Mạch tiếp tục nối với nhóm -CH~2~- và cuối cùng là nhóm -CH~3~.\n  ⇒ Công thức cấu tạo đầy đủ: CH~2~=C(CH~3~)–CH~2~–CH~3~ (2-methylbut-1-ene).\n  ⇒ Công thức phân tử: C~5~H~10~.\n• Tổng số nguyên tử hydrogen trong phân tử X là 10.\n*Đáp án đúng: A.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8. Cho iron (III) oxide tác dụng với nitric acid loãng thì sản phẩm thu được là",
            "options": [
                "A. Fe(NO~3~)~3~, NO~2~ và H~2~O.",
                "*B. Fe(NO~3~)~3~ và H~2~O.*",
                "C. Fe(NO~3~)~3~, NO và H~2~O.",
                "D. Fe(NO~3~)~3~, N~2~ và H~2~O."
            ],
            "solution": "• Trong iron (III) oxide (Fe~2~O~3~), nguyên tố Fe đã ở trạng thái số oxi hóa cao nhất là +3, do đó khi tác dụng với dung dịch HNO~3~ loãng chỉ xảy ra phản ứng trao đổi acid - base tạo muối và nước (không xảy ra phản ứng oxi hóa - khử, không giải phóng khí):\n  Fe~2~O~3~ + 6HNO~3~ → 2Fe(NO~3~)~3~ + 3H~2~O\n*Đáp án đúng: B.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9. Hợp chất X tan trong nước tạo dung dịch không màu. Dung dịch này không tạo kết tủa với dung dịch BaCl~2~, khi phản ứng với NaOH tạo ra khí mùi khai, khi phản ứng với dung dịch HCl tạo ra khí làm đục nước vôi trong và làm mất màu dung dịch nước bromine. Chất X là",
            "options": [
                "A. (NH~4~)~2~SO~3~",
                "B. (NH~4~)~2~CO~3~",
                "C. NH~4~HCO~3~",
                "*D. NH~4~HSO~3~*"
            ],
            "solution": "• Phân tích các dữ kiện thực nghiệm:\n  - Tác dụng với NaOH tạo khí mùi khai: chứng tỏ phân tử có chứa cation ammonium (NH~4~^+^).\n  - Không tạo kết tủa với dung dịch BaCl~2~: loại (NH~4~)~2~SO~3~ (tạo BaSO~3~↓ trắng) và loại (NH~4~)~2~CO~3~ (tạo BaCO~3~↓ trắng). Muối acid NH~4~HCO~3~ và NH~4~HSO~3~ không tạo kết tủa với BaCl~2~.\n  - Tác dụng với dung dịch HCl tạo khí làm đục nước vôi trong và làm mất màu nước bromine: Khí sinh ra phải là SO~2~ (SO~2~ + Ca(OH)~2~ → CaSO~3~↓ + H~2~O và SO~2~ + Br~2~ + 2H~2~O → 2HBr + H~2~SO~4~). Khí CO~2~ (từ gốc HCO~3~^-^) làm đục nước vôi trong nhưng KHÔNG làm mất màu nước bromine.\n  ⇒ Gốc acid là HSO~3~^-^, chất X là NH~4~HSO~3~.\n*Đáp án đúng: D.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10. Chất nào dưới đây là đồng đẳng của ethanol CH~3~CH~2~OH ?",
            "options": [
                "A. CH~3~OCH~2~CH~3~.",
                "B. CH~2~=CHCH~2~OH.",
                "C. CH~3~OCH~3~",
                "*D. CH~3~CH~2~CH~2~CH~2~OH.*"
            ],
            "solution": "• Đồng đẳng của ethanol là các alcohol no, đơn chức, mạch hở có công thức chung C~n~H~2n+1~OH (n ≥ 1).\n• CH~3~CH~2~CH~2~CH~2~OH (butan-1-ol) có cấu tạo và tính chất tương tự ethanol, hơn kém 2 nhóm -CH~2~-, nên thuộc cùng dãy đồng đẳng với ethanol.\n*Đáp án đúng: D.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11. Một số cơ sở sản xuất thuốc Bắc thường đốt một chất bột rắn X màu vàng (là một đơn chất) để tạo ra khí Y nhằm mục đích tẩy trắng, chống mốc. Tuy nhiên, theo các nhà khoa học thì khí Y có ảnh hưởng không tốt đến cơ quan nội tạng và khí Y cũng là một trong những nguyên nhân gây ra “mưa acid”. Chất rắn X là",
            "options": [
                "A. iodine.",
                "B. carbon.",
                "*C. sulfur.*",
                "D. phosphorus."
            ],
            "solution": "• Đơn chất bột rắn màu vàng là sulfur (S). Khi đốt cháy sulfur sinh ra khí sulfur dioxide (SO~2~): S + O~2~ —(t°)→ SO~2~. Khí SO~2~ có tính khử màu, diệt nấm mốc nhưng rất độc đối với hệ hô hấp và nội tạng của con người, đồng thời là nguyên nhân chính gây ra mưa acid.\n*Đáp án đúng: C.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12. Dụng cụ này có thể được sử dụng để tách riêng các chất trong hỗn hợp nào dưới đây?",
            "image": "pheu_chiet_c12_clean.png",
            "image_width": 1.2,
            "options": [
                "A. Xăng và dầu ăn.",
                "B. Nước và rượu.",
                "C. Nước và muối ăn.",
                "*D. Dầu ăn và nước.*"
            ],
            "solution": "• Dụng cụ trong hình là phễu chiết quả lê (có khóa), dùng để thực hiện phương pháp chiết lỏng - lỏng nhằm tách riêng hai chất lỏng không hòa tan vào nhau (tách thành hai lớp lỏng riêng biệt).\n• Dầu ăn không tan trong nước và nhẹ hơn nước nên nổi lên trên tạo thành 2 lớp chất lỏng phân biệt, dễ dàng tách riêng bằng phễu chiết.\n*Đáp án đúng: D.*"
        },
        {
            "id": 13,
            "prompt": "Câu 13. Chất nào sau đây không thuộc loại chất hữu cơ?",
            "options": [
                "A. CH~3~COONa",
                "*B. Al~4~C~3~*",
                "C. CH~3~Cl",
                "D. CCl~4~"
            ],
            "solution": "• Aluminium carbide (Al~4~C~3~) là hợp chất carbide kim loại, thuộc nhóm các hợp chất vô cơ của carbon.\n*Đáp án đúng: B.*"
        },
        {
            "id": 14,
            "prompt": "Câu 14. Khí thải công nghiệp, khí thải động cơ (ô tô, xe máy ...) là nguyên nhân chủ yếu gây mưa acid. Dãy nào dưới đây gồm các chất có trong các loại khí thải này và là nguyên nhân trực tiếp gây mưa acid?",
            "options": [
                "A. SO~2~, CO, NO.",
                "B. SO~2~, CO, NO~2~.",
                "*C. NO, NO~2~, SO~2~.*",
                "D. NO~2~, CO~2~, CO."
            ],
            "solution": "• Các chất khí là tác nhân trực tiếp gây nên hiện tượng mưa acid là các oxide của nitrogen (NO, NO~2~) và sulfur dioxide (SO~2~). Chúng bị oxi hóa trong khí quyển kết hợp với nước mưa tạo ra các acid mạnh tương ứng (HNO~3~ và H~2~SO~4~).\n*Đáp án đúng: C.*"
        },
        {
            "id": 15,
            "prompt": "Câu 15. Khi cho saccharose (C~12~H~22~O~11~) tác dụng với dung dịch H~2~SO~4~ đặc trong cốc thủy tinh thấy có bọt khí đẩy carbon trào lên khỏi cốc. Thí nghiệm trên chứng minh được tính chất nào sau đây của H~2~SO~4~ đặc?",
            "options": [
                "A. chỉ có tính oxi hóa mạnh.",
                "*B. tính háo nước và tính oxi hóa mạnh.*",
                "C. chỉ có tính háo nước.",
                "D. tính háo nước và tính khử mạnh."
            ],
            "solution": "• Ban đầu, H~2~SO~4~ đặc hút nước từ phân tử saccharose hóa than màu đen: C~12~H~22~O~11~ —(H~2~SO~4~ đặc)→ 12C + 11H~2~O (thể hiện tính háo nước mạnh).\n• Sau đó, H~2~SO~4~ đặc tiếp tục oxi hóa carbon thành các khí CO~2~ và SO~2~: C + 2H~2~SO~4~(đặc) —(t°)→ CO~2~↑ + 2SO~2~↑ + 2H~2~O (thể hiện tính oxi hóa mạnh). Các bọt khí sinh ra đẩy khối carbon xốp trào dâng lên miệng cốc.\n*Đáp án đúng: B.*"
        },
        {
            "id": 16,
            "prompt": "Câu 16. Kim loại nào sau đây không tác dụng với dung dịch acid H~2~SO~4~ loãng nguội?",
            "options": [
                "A. Na",
                "B. Zn",
                "*C. Cu*",
                "D. Al"
            ],
            "solution": "• Kim loại đồng (Cu) đứng sau hydrogen trong dãy hoạt động hóa học của kim loại nên không tác dụng với dung dịch H~2~SO~4~ loãng để giải phóng khí H~2~.\n*Đáp án đúng: C.*"
        }
    ],

    # Phần 2: Trắc nghiệm dạng Đúng/Sai (2,0 điểm) - 2 câu
    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1. Pha 1,0 L dung dịch sulfuric acid. Sau đó lấy 10,0 mL dung dịch acid cho vào bình tam giác, thêm vài giọt dung dịch phenolphthalein. Nhỏ từ từ dung dịch NaOH 0,10 M chứa trên burette vào bình tam giác đến khi dung dịch xuất hiện màu hồng nhạt, đọc thể tích NaOH đã dùng trên burette. Lặp lại thí nghiệm 3 lần, các giá trị thể tích NaOH đọc được lần lượt là 17,9 mL; 18,0 mL; 18,1 mL.",
            "statements": [
                ("a) Khi chuẩn độ xảy ra phản ứng hoá học H~2~SO~4~ + 2 NaOH → Na~2~SO~4~ + 2 H~2~O.", "ĐÚNG"),
                ("b) pH của dung dịch acid gần bằng 1,05.", "ĐÚNG"),
                ("c) Để pha 1 lít dung dịch sulfuric acid trên cần hoà 19,6 ml (làm tròn đến phần chục) dung dịch H~2~SO~4~ 98% (d=1,84g/ml) vào 980,4 ml nước.", "SAI"),
                ("d) Phương pháp trong thí nghiệm trên là phương pháp chuẩn độ acid-base.", "ĐÚNG")
            ],
            "solution": "• Thể tích dung dịch NaOH trung bình: V~NaOH~ = (17,9 + 18,0 + 18,1) / 3 = 18,0 mL = 0,018 L.\n• Số mol NaOH phản ứng: n(NaOH) = 0,018 × 0,10 = 0,0018 mol.\n• Ý a: ĐÚNG. Phản ứng trung hòa khi chuẩn độ: H~2~SO~4~ + 2NaOH → Na~2~SO~4~ + 2H~2~O.\n• Ý b: ĐÚNG. Theo phương trình: n(H~2~SO~4~ trong 10 mL) = 0,0018 / 2 = 0,0009 mol ⇒ Nồng độ mol của acid là C~M~ = 0,0009 / 0,01 = 0,09 M. Theo cách tính chuẩn độ đơn nấc [H^+^] = 0,09 M ⇒ pH = -log(0,09) ≈ 1,046 ≈ 1,05.\n• Ý c: SAI. Để pha 1 L dung dịch H~2~SO~4~ 0,09 M cần n(H~2~SO~4~) = 0,09 mol ⇒ m(H~2~SO~4~) = 0,09 × 98 = 8,82 gam ⇒ m(dd 98%) = 8,82 / 0,98 = 9,0 gam ⇒ V(dd H~2~SO~4~ 98%) = 9,0 / 1,84 ≈ 4,89 mL (không phải 19,6 mL).\n• Ý d: ĐÚNG. Thí nghiệm xác định nồng độ acid bằng dung dịch base chuẩn NaOH với chỉ thị phenolphthalein là phương pháp chuẩn độ acid - base."
        },
        {
            "id": 2,
            "prompt": "Câu 2. Đốt cháy hoàn toàn 7,25 gam hợp chất hữu cơ X thu được 11,2 lít CO~2~ (đktc) và 11,25 gam H~2~O. Biết tỉ khối của X so với khí nitrogen bằng 29/14.",
            "statements": [
                ("a) Phần trăm khối lượng của nguyên tố hydrogen trong X xấp xỉ 8,62%", "SAI"),
                ("b) X có chứa 4 nguyên tử carbon.", "ĐÚNG"),
                ("c) Phân tử khối của X bằng 29.", "SAI"),
                ("d) Phân tử X chứa 3 loại nguyên tố là C, H, O.", "SAI")
            ],
            "solution": "• Tính toán các đại lượng:\n  - n(CO~2~) = 11,2 / 22,4 = 0,5 mol ⇒ n(C) = 0,5 mol ⇒ m(C) = 0,5 × 12 = 6,0 gam.\n  - n(H~2~O) = 11,25 / 18 = 0,625 mol ⇒ n(H) = 0,625 × 2 = 1,25 mol ⇒ m(H) = 1,25 × 1 = 1,25 gam.\n  - Nhận thấy: m(C) + m(H) = 6,0 + 1,25 = 7,25 gam = m(X) ⇒ X là hydrocarbon, KHÔNG chứa nguyên tố oxygen.\n  - Phân tử khối của X: M~X~ = 28 × (29 / 14) = 58 g/mol.\n  - n(X) = 7,25 / 58 = 0,125 mol ⇒ Số nguyên tử C = 0,5 / 0,125 = 4; Số nguyên tử H = 1,25 / 0,125 = 10 ⇒ CTPT: C~4~H~10~.\n• Ý a: SAI. %m(H) = (1,25 / 7,25) × 100% ≈ 17,24% (không phải 8,62%).\n• Ý b: ĐÚNG. Phân tử X chứa 4 nguyên tử carbon (C~4~H~10~).\n• Ý c: SAI. Phân tử khối của X là M = 58 g/mol (không phải 29).\n• Ý d: SAI. Phân tử X chỉ chứa 2 loại nguyên tố là C và H (không có O)."
        }
    ],

    # Phần 3: Trắc nghiệm trả lời ngắn (1,0 điểm) - 4 câu
    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1. Hoà tan 5,4 gam aluminium (Al) bằng acid H~2~SO~4~ loãng dư sinh ra V lít khí (ở đktc). Tính V.",
            "answer": "6,72",
            "solution": "• Số mol Al: n(Al) = 5,4 / 27 = 0,2 mol.\n• Phương trình phản ứng:\n  2Al + 3H~2~SO~4~ → Al~2~(SO~4~)~3~ + 3H~2~↑\n• Theo phương trình: n(H~2~) = 3/2 × n(Al) = 1,5 × 0,2 = 0,3 mol.\n• Thể tích khí H~2~ ở đktc:\n  V = 0,3 × 22,4 = 6,72 lít.\n  (Nếu tính theo điều kiện chuẩn 25 °C, 1 bar: V = 0,3 × 24,79 = 7,437 L).\n*Đáp án: 6,72*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Khi cho dung dịch HNO~3~ đặc, nóng lần lượt tác dụng với các chất sau: Cu, Fe~2~O~3~, C, dung dịch Ba(OH)~2~, dung dịch Na~2~SO~3~. Số phản ứng mà HNO~3~ đóng vai trò là chất oxi hóa là bao nhiêu?",
            "answer": "3",
            "solution": "• Dung dịch HNO~3~ đóng vai trò là chất oxi hóa khi tác dụng với các chất có tính khử (làm số oxi hóa của N giảm xuống):\n  1) Cu + 4HNO~3~(đặc) —(t°)→ Cu(NO~3~)~2~ + 2NO~2~↑ + 2H~2~O (Oxi hóa - khử).\n  2) Fe~2~O~3~: Fe đã có số oxi hóa cao nhất (+3) ⇒ Phản ứng trao đổi tạo muối và nước.\n  3) C + 4HNO~3~(đặc) —(t°)→ CO~2~↑ + 4NO~2~↑ + 2H~2~O (Oxi hóa - khử).\n  4) Ba(OH)~2~: Phản ứng trung hòa acid - base.\n  5) Na~2~SO~3~ + 2HNO~3~(đặc) —(t°)→ Na~2~SO~4~ + 2NO~2~↑ + H~2~O (S^+4^ bị oxi hóa lên S^+6^ ⇒ Oxi hóa - khử).\n⇒ Có 3 phản ứng HNO~3~ đóng vai trò là chất oxi hóa (khi tác dụng với Cu, C, Na~2~SO~3~).\n*Đáp án: 3*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Cho các cân bằng hoá học:\n"
                      "(1) N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)\n"
                      "(2) H~2~(g) + I~2~(g) ⇌ 2HI(g)\n"
                      "(3) 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g)\n"
                      "(4) 2NO~2~(g) ⇌ N~2~O~4~(g)\n"
                      "Khi thay đổi áp suất có bao nhiêu cân bằng hóa học bị chuyển dịch?",
            "answer": "3",
            "solution": "• Cân bằng hóa học chỉ bị chuyển dịch khi thay đổi áp suất nếu tổng số mol khí ở hai vế phương trình là khác nhau (Δn~khí~ ≠ 0):\n  (1) Δn = 2 - (1 + 3) = -2 ≠ 0 ⇒ Cân bằng bị chuyển dịch.\n  (2) Δn = 2 - (1 + 1) = 0 ⇒ Cân bằng KHÔNG bị chuyển dịch khi thay đổi áp suất.\n  (3) Δn = 2 - (2 + 1) = -1 ≠ 0 ⇒ Cân bằng bị chuyển dịch.\n  (4) Δn = 1 - 2 = -1 ≠ 0 ⇒ Cân bằng bị chuyển dịch.\n⇒ Có 3 cân bằng hóa học bị chuyển dịch khi thay đổi áp suất là: (1), (3), (4).\n*Đáp án: 3*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Ứng với CTPT C~3~H~8~O có bao nhiêu đồng phân cấu tạo thuộc loại alcohol?",
            "answer": "2",
            "solution": "• CTPT C~3~H~8~O có độ bất bão hòa k = 0, các đồng phân cấu tạo alcohol no, đơn chức, mạch hở gồm:\n  1) CH~3~–CH~2~–CH~2~–OH (propan-1-ol)\n  2) CH~3~–CH(OH)–CH~3~ (propan-2-ol)\n⇒ Có 2 đồng phân cấu tạo thuộc loại alcohol.\n*Đáp án: 2*"
        }
    ],

    # B. PHẦN TỰ LUẬN (3,0 điểm) - 3 câu
    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1. Viết công thức cấu tạo thu gọn và công thức khung phân tử các đồng phân cấu tạo ứng với công thức phân tử C~4~H~10~.",
            "num_lines": 6,
            "solution": "Các đồng phân cấu tạo ứng với C~4~H~10~:\n"
                       "1. Butane (mạch thẳng 4C):\n"
                       "   • Công thức cấu tạo thu gọn: CH~3~–CH~2~–CH~2~–CH~3~\n"
                       "   • Công thức khung phân tử: Mạch dích dắc 4 đỉnh (\\/\\)\n"
                       "2. 2-Methylpropane / Isobutane (mạch nhánh):\n"
                       "   • Công thức cấu tạo thu gọn: CH~3~–CH(CH~3~)–CH~3~\n"
                       "   • Công thức khung phân tử: Hình chữ Y 3 nhánh (>–)"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Caffeine là chất kích thích tự nhiên được tìm thấy trong cây trà, cà phê và ca cao. Chúng hoạt động bằng cách kích thích não và hệ thần kinh trung ương, giúp con người tỉnh táo và ngăn ngừa sự mệt mỏi. Thành phần phần trăm về khối lượng của các nguyên tố trong phân tử caffeine như sau: 49,48% carbon; 5,15% hydrogen; 28,87% nitrogen còn lại là oxygen. Phổ MS của caffeine được cho như hình dưới đây:\n"
                      "a. Dựa vào phổ khối lượng phân tử (MS) hãy xác định phân tử khối của caffeine ?\n"
                      "b. Xác định công thức phân tử của caffeine.",
            "image": "ms_caffeine_clean.png",
            "image_width": 4.5,
            "num_lines": 10,
            "solution": "a. Xác định phân tử khối của caffeine:\n"
                       "• Trên phổ khối lượng (MS) của caffeine, peak ion phân tử [M^+^] có giá trị m/z lớn nhất là 194.\n"
                       "⇒ Phân tử khối của caffeine là M = 194 g/mol.\n\n"
                       "b. Xác định công thức phân tử của caffeine:\n"
                       "• Phần trăm khối lượng của nguyên tố oxygen trong phân tử:\n"
                       "  %O = 100% - 49,48% - 5,15% - 28,87% = 16,50%.\n"
                       "• Đặt công thức phân tử của caffeine là C~x~H~y~N~z~O~t~.\n"
                       "• Số nguyên tử của mỗi nguyên tố trong phân tử:\n"
                       "  x = (194 × 49,48%) / 12 = 8,00 ⇒ x = 8\n"
                       "  y = (194 × 5,15%) / 1 = 9,99 ⇒ y = 10\n"
                       "  z = (194 × 28,87%) / 14 = 4,00 ⇒ z = 4\n"
                       "  t = (194 × 16,50%) / 16 = 2,00 ⇒ t = 2\n"
                       "⇒ Công thức phân tử của caffeine là C~8~H~10~N~4~O~2~."
        },
        {
            "id": 3,
            "prompt": "Câu 3. Sulfuric acid là một trong những hoá chất quan trọng nhất được sử dụng trong công nghiệp; được sản xuất hàng trăm triệu tấn mỗi năm, chiếm nhiều nhất trong ngành công nghiệp hoá chất. Sulfuric acid được sản xuất từ quặng pyrite theo sơ đồ sau:\n"
                      "FeS~2~ —(1)→ SO~2~ —(2)→ SO~3~ —(3)→ oleum —(4)→ H~2~SO~4~.\n"
                      "a. Hoàn thành sơ đồ phản ứng trên, ghi rõ điều kiện phản ứng nếu có.\n"
                      "b. Hấp thụ toàn bộ lượng SO~3~ điều chế được từ m kg FeS~2~ vào 1000 gam dung dịch H~2~SO~4~ 98% thu được oleum X. Hoà tan 50,7 gam X vào nước được dung dịch Y. Cho dung dịch BaCl~2~ dư vào Y thu được 139,8 gam kết tủa. Giả sử các phản ứng xảy ra hoàn toàn. Tìm m.",
            "num_lines": 14,
            "solution": "a. Hoàn thành sơ đồ phản ứng:\n"
                       "(1) 4FeS~2~ + 11O~2~ —(t°)→ 2Fe~2~O~3~ + 8SO~2~↑\n"
                       "(2) 2SO~2~ + O~2~ ⇌ (V~2~O~5~, 450 °C) 2SO~3~\n"
                       "(3) H~2~SO~4~ + nSO~3~ → H~2~SO~4~.nSO~3~ (oleum)\n"
                       "(4) H~2~SO~4~.nSO~3~ + nH~2~O → (n+1) H~2~SO~4~\n\n"
                       "b. Tìm giá trị m (kg quặng FeS~2~):\n"
                       "• Trong 50,7 gam oleum X:\n"
                       "  - Cho BaCl~2~ dư vào dung dịch Y thu được kết tủa BaSO~4~:\n"
                       "    n(BaSO~4~) = 139,8 / 233 = 0,6 mol.\n"
                       "  - Theo bảo toàn nguyên tố S: Tổng số mol S trong 50,7 gam X = n(BaSO~4~) = 0,6 mol.\n"
                       "  - Khối lượng phân tử trung bình tương ứng cho 1 mol nguyên tử S trong oleum H~2~SO~4~.nSO~3~ là:\n"
                       "    M(oleum) / (n + 1) = 50,7 / 0,6 = 84,5 g/mol\n"
                       "    ⇔ (98 + 80n) / (n + 1) = 84,5 ⇔ 98 + 80n = 84,5n + 84,5 ⇔ 4,5n = 13,5 ⇔ n = 3.\n"
                       "  ⇒ Công thức của oleum X là H~2~SO~4~.3SO~3~.\n\n"
                       "• Xét quá trình tạo oleum X từ 1000 gam dung dịch H~2~SO~4~ 98%:\n"
                       "  - Khối lượng H~2~SO~4~ nguyên chất: m(H~2~SO~4~) = 1000 × 98% = 980 gam ⇒ n(H~2~SO~4~) = 980 / 98 = 10 mol.\n"
                       "  - Khối lượng nước: m(H~2~O) = 1000 - 980 = 20 gam ⇒ n(H~2~O) = 20 / 18 = 10/9 mol.\n"
                       "  - Khi hấp thụ SO~3~ vào dung dịch:\n"
                       "    + Giai đoạn 1: SO~3~ phản ứng với H~2~O tạo H~2~SO~4~:\n"
                       "      SO~3~ + H~2~O → H~2~SO~4~\n"
                       "      n(SO~3~ pư H~2~O) = n(H~2~O) = 10/9 mol.\n"
                       "      Lượng H~2~SO~4~ mới tạo thành: 10/9 mol ⇒ Tổng n(H~2~SO~4~) = 10 + 10/9 = 100/9 mol.\n"
                       "    + Giai đoạn 2: SO~3~ kết hợp với toàn bộ H~2~SO~4~ tạo oleum H~2~SO~4~.3SO~3~:\n"
                       "      H~2~SO~4~ + 3SO~3~ → H~2~SO~4~.3SO~3~\n"
                       "      n(SO~3~ tạo oleum) = 3 × n(H~2~SO~4~) = 3 × (100/9) = 100/3 mol = 300/9 mol.\n"
                       "  - Tổng số mol SO~3~ đã được hấp thụ là:\n"
                       "    n(SO~3~ tổng) = 10/9 + 300/9 = 310/9 mol.\n\n"
                       "• Sơ đồ bảo toàn nguyên tố S từ FeS~2~:\n"
                       "  FeS~2~ ———→ 2 SO~3~\n"
                       "  ⇒ n(FeS~2~) = 1/2 × n(SO~3~ tổng) = (1/2) × (310/9) = 155/9 mol.\n"
                       "• Khối lượng FeS~2~ cần dùng:\n"
                       "  m(FeS~2~) = n(FeS~2~) × M(FeS~2~) = (155/9) × 120 = 6200 / 3 ≈ 2066,67 gam.\n"
                       "• Đổi ra đơn vị kg theo yêu cầu đề bài: m ≈ 2,067 kg (hoặc m ≈ 2,07 kg)."
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
        add_quick_answer_table_1101(doc)

    # A. PHẦN TRẮC NGHIỆM
    add_section_header(
        doc,
        title="A. PHẦN TRẮC NGHIỆM",
        subtitle="PHẦN 1: TRẮC NGHIỆM NHIỀU LỰA CHỌN (16 câu – 4 điểm). Thí sinh trả lời câu hỏi từ câu 1 đến câu 16, mỗi câu hỏi thí sinh chỉ chọn một phương án."
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

    # PHẦN 2: TRẮC NGHIỆM ĐÚNG – SAI
    add_section_header(
        doc,
        title="PHẦN 2: TRẮC NGHIỆM ĐÚNG – SAI (2 câu – 2 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
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

    # PHẦN 3: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN
    add_section_header(
        doc,
        title="PHẦN 3: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (4 câu - 1 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4."
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

    # B. PHẦN TỰ LUẬN
    add_section_header(
        doc,
        title="B. PHẦN TỰ LUẬN ( 3 câu – 3 điểm)",
        subtitle="Thí sinh trình bày lời giải chi tiết vào phần giấy làm bài thi."
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
                w = item.get("image_width", 4.5)
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
    filename = f"De_Kiem_Tra_CK1_Hoa_11_NguyenChiThanh_Ma1101{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo thành công: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
