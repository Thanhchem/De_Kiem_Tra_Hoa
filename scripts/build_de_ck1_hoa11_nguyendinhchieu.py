# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Học Kì I - Hóa học 11 - Trường THPT Nguyễn Đình Chiểu (TP Huế)
Mã đề: 1103 - Năm học 2025 - 2026
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

def add_quick_answer_table_1103(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Giáo viên ở đầu tài liệu Mã đề 1103.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 1103)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_p1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
    answers_p1 = ["Đ/A", "A", "B", "B", "D", "D", "B", "A", "C", "D", "C", "C", "C"]

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
        ["Câu 1", "Phương pháp tách và tinh chế hợp chất hữu cơ", "ĐÚNG", "SAI", "SAI", "SAI"],
        ["Câu 2", "Sơ đồ chuyển hoá của hợp chất chứa nitrogen", "ĐÚNG", "ĐÚNG", "ĐÚNG", "SAI"]
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
    render_rich_text(p_p3, "*3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM - MỖI CÂU 0,5 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False

    tbl3_headers = ["Câu", "Câu 1", "Câu 2", "Câu 3", "Câu 4"]
    tbl3_answers = ["Đáp án", "311", "4", "24", "0,8"]

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
        ("Câu 1", "1,0 đ", "• a) Tăng nhiệt độ: Cân bằng chuyển dịch theo chiều nghịch (chiều thu nhiệt).\n• b) Tăng áp suất chung: Cân bằng chuyển dịch theo chiều thuận (chiều giảm số mol khí 3 → 2).\n• c) Thêm xúc tác V~2~O~5~: Cân bằng không bị chuyển dịch.\n• d) Tăng nồng độ O~2~: Cân bằng chuyển dịch theo chiều thuận (tiêu hao bớt O~2~)."),
        ("Câu 2", "1,0 đ", "• a) Thành phần % khối lượng: %m(Fe) ≈ 53,85%; %m(Cu) ≈ 46,15%.\n• b) Thể tích dung dịch HNO~3~ 1M đã dùng: V = 1,2 lít."),
        ("Câu 3", "1,0 đ", "• Thể tích dung dịch H~2~SO~4~ 0,3M: V = 500 mL (hay 0,5 lít).\n• Khối lượng kết tủa BaSO~4~ thu được: m = 11,65 gam.")
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


# Dữ liệu Đề thi mã 1103
DATA_EXAM = {
    "school_line": "SỞ GD&ĐT THÀNH PHỐ HUẾ - TRƯỜNG THPT NGUYỄN ĐÌNH CHIỂU",
    "exam_title": "ĐỀ KIỂM TRA CUỐI HỌC KÌ I - NĂM HỌC 2025 - 2026",
    "subject_line": "MÔN: HOÁ HỌC - KHỐI 11",
    "duration": "45 phút (không kể thời gian giao đề)",
    "num_pages": "03",
    "exam_code": "1103",
    "atomic_masses": "H = 1; C = 12; N = 14; O = 16; Cl = 35,5; Ca = 40; S = 32; Fe = 56; Cu = 64; Br = 80; Mn = 55; F = 19; I = 127; Al = 27",

    # A. PHẦN TRẮC NGHIỆM: (7 điểm)
    # PHẦN I. (12 câu – 3 điểm) Câu trắc nghiệm nhiều phương án lựa chọn.
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1. Bước sơ cứu đầu tiên cần làm ngay khi một người bị bỏng sulfuric acid là:",
            "options": [
                "*A. rửa với nước lạnh nhiều lần.*",
                "B. băng bó tạm thời vết bỏng.",
                "C. trung hoà acid bằng NaHCO~3~.",
                "D. đưa đến cơ sở y tế gần nhất."
            ],
            "solution": "Khi bị bỏng sulfuric acid (đặc biệt là acid đặc), bước sơ cứu đầu tiên và quan trọng nhất cần thực hiện ngay lập tức là xối rửa vết bỏng liên tục dưới vòi nước lạnh sạch nhiều lần trong ít nhất 15 – 20 phút để nhanh chóng làm loãng, cuốn trôi acid và giảm nhiệt độ tổn thương bề mặt da trước khi tiến hành các bước băng bó hoặc chuyển tới cơ sở y tế.\n*Đáp án đúng: A.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Sự có mặt của khí SO~2~ trong không khí là nguyên nhân chính gây ra hiện tượng mưa acid. Nồng độ của SO~2~ có thể xác định bằng cách chuẩn độ với dung dịch KMnO~4~ theo phản ứng sau:\n5SO~2~ + 2KMnO~4~ + 2H~2~O → K~2~SO~4~ + 2MnSO~4~ + 2H~2~SO~4~.\nBiết một mẫu không khí phản ứng vừa đủ với 150 mL dung dịch KMnO~4~ 8,0.10^-3^ M. Khối lượng (gam) của SO~2~ có trong mẫu không khí đó là:",
            "options": [
                "A. 13,6.",
                "*B. 0,192.*",
                "C. 19,2.",
                "D. 0,136."
            ],
            "solution": "• Số mol KMnO~4~ phản ứng: n(KMnO~4~) = 0,150 L × 8,0.10^-3^ M = 1,2.10^-3^ mol.\n• Theo phương trình hoá học:\n  5SO~2~ + 2KMnO~4~ + 2H~2~O → K~2~SO~4~ + 2MnSO~4~ + 2H~2~SO~4~\n  Tỉ lệ mol: n(SO~2~) = (5/2) × n(KMnO~4~) = 2,5 × 1,2.10^-3^ = 3,0.10^-3^ mol.\n• Khối lượng SO~2~ có trong mẫu không khí:\n  m(SO~2~) = 3,0.10^-3^ mol × 64 g/mol = 0,192 gam.\n*Đáp án đúng: B.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Trộn 300 ml dung dịch H~2~SO~4~ 1M với 100ml dung dịch BaCl~2~ 2M. Khối lượng kết tủa thu được là:",
            "options": [
                "A. 6,6 gam.",
                "*B. 46,6 gam.*",
                "C. 11,65 gam.",
                "D. 23,3 gam."
            ],
            "solution": "• Số mol các chất phản ứng:\n  n(H~2~SO~4~) = 0,3 L × 1 M = 0,3 mol.\n  n(BaCl~2~) = 0,1 L × 2 M = 0,2 mol.\n• Phương trình hoá học: H~2~SO~4~ + BaCl~2~ → BaSO~4~↓ + 2HCl.\n  So sánh tỉ lệ mol: n(BaCl~2~) = 0,2 mol < n(H~2~SO~4~) = 0,3 mol ⇒ H~2~SO~4~ dư, BaCl~2~ phản ứng hết.\n• Số mol kết tủa BaSO~4~ tính theo BaCl~2~: n(BaSO~4~) = 0,2 mol.\n• Khối lượng kết tủa: m(BaSO~4~) = 0,2 mol × 233 g/mol = 46,6 gam.\n*Đáp án đúng: B.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Cho các dung dịch có cùng nồng độ: Na~2~CO~3~ (1), H~2~SO~4~ (2), HCl (3), KNO~3~ (4). Giá trị pH của các dung dịch được sắp xếp theo chiều tăng dần từ trái sang phải là:",
            "options": [
                "A. (4), (1), (2), (3).",
                "B. (1), (2), (3), (4).",
                "C. (3), (2), (4), (1).",
                "*D. (2), (3), (4), (1).*"
            ],
            "solution": "Xét môi trường và nồng độ H^+^ của các dung dịch ở cùng nồng độ mol C (M):\n- H~2~SO~4~ (2): acid mạnh 2 nấc, [H^+^] = 2C ⇒ pH nhỏ nhất.\n- HCl (3): acid mạnh 1 nấc, [H^+^] = C ⇒ pH nhỏ thứ hai.\n- KNO~3~ (4): muối trung tính tạo bởi cation base mạnh (K^+^) và anion acid mạnh (NO~3~^-^), không bị thủy phân ⇒ môi trường trung tính, pH = 7.\n- Na~2~CO~3~ (1): muối tạo bởi cation base mạnh (Na^+^) và anion acid yếu (CO~3~^2-^), gốc CO~3~^2-^ thủy phân sinh ra ion OH^-^ ⇒ môi trường base, pH > 7 lớn nhất.\n⇒ Thứ tự pH tăng dần: (2) < (3) < (4) < (1).\n*Đáp án đúng: D.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5. Phương pháp chiết dùng để tách chất trong hỗn hợp nào sau đây?",
            "options": [
                "A. Dùng nam châm hút.",
                "B. Nước và rượu.",
                "C. Cát và nước.",
                "*D. Nước và dầu ăn.*"
            ],
            "solution": "Phương pháp chiết lỏng – lỏng dùng để tách các chất lỏng không hòa tan vào nhau (tách thành hai lớp chất lỏng riêng biệt trong phễu chiết). Hỗn hợp nước và dầu ăn không đồng tan, dầu ăn nhẹ hơn nổi lên trên, nước ở dưới nên dễ dàng tách riêng bằng phễu chiết.\n*Đáp án đúng: D.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6. Chất nào sau đây là chất điện li?",
            "options": [
                "A. MgO.",
                "*B. H~3~PO~4~.*",
                "C. C~2~H~6~.",
                "D. Cl~2~."
            ],
            "solution": "Chất điện li là chất khi tan trong nước phân li ra các ion. Phosphoric acid (H~3~PO~4~) tan trong nước và phân li từng nấc ra các ion H^+^, H~2~PO~4~^-^, HPO~4~^2-^, PO~4~^3-^ nên là chất điện li (chất điện li yếu). C~2~H~6~ là hydrocarbon không điện li, Cl~2~ là đơn chất, MgO là oxit rắn không tan.\n*Đáp án đúng: B.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7. Cách xử lý thủy ngân khi nhiệt kế thủy ngân không may bị vỡ là:",
            "options": [
                "*A. rắc bột sulfur lên thủy ngân rồi gom lại.*",
                "B. rắc muối ăn lên thủy ngân rồi gom lại.",
                "C. rắc bột sắt lên thủy ngân rồi gom lại.",
                "D. rắc đường lên thủy ngân rồi gom lại."
            ],
            "solution": "Thủy ngân (Hg) rất độc và dễ bay hơi ở nhiệt độ phòng. Khi nhiệt kế vỡ, người ta rắc bột sulfur (lưu huỳnh) lên thủy ngân vì sulfur phản ứng êm dịu ngay ở nhiệt độ thường tạo thành mercury(II) sulfide (HgS) dạng rắn, không bay hơi và không độc, dễ dàng thu gom: Hg + S → HgS.\n*Đáp án đúng: A.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8. Cặp chất nào dưới đây là đồng đẳng của nhau?",
            "options": [
                "A. CH~3~CH~2~CH~2~CH~3~ và (CH~3~)~2~CHCH~3~.",
                "B. CH~2~=CH–CH=CH~2~ và CH~3~C≡C–CH~3~.",
                "*C. CH~2~=CH–CH=CH~2~ và CH~2~=C(CH~3~)–CH=CH~2~.*",
                "D. CH~3~CH=CH~2~ và CH~3~–CH~2~–CH~2~–CH~3~."
            ],
            "solution": "Đồng đẳng là những hợp chất có cấu tạo hoá học tương tự nhau và có công thức phân tử hơn kém nhau một hay nhiều nhóm methylene (–CH~2~–).\n- CH~2~=CH–CH=CH~2~ (C~4~H~6~, buta-1,3-diene) và CH~2~=C(CH~3~)–CH=CH~2~ (C~5~H~8~, isoprene / 2-methylbuta-1,3-diene) đều là alkadiene liên hợp mạch hở, hơn kém nhau 1 nhóm –CH~2~– nên là đồng đẳng của nhau.\n- Cặp A và B cùng CTPT nên là đồng phân; cặp D thuộc hai dãy đồng đẳng khác nhau (alkene và alkane).\n*Đáp án đúng: C.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9. Để chuẩn độ 40 ml dung dịch HCl chưa biết nồng độ đã dùng trung bình hết 34 ml dung dịch NaOH 0,12 M. Nồng độ mol của dung dịch HCl là:",
            "options": [
                "A. 1,2.",
                "B. 0,012.",
                "C. 0,12.",
                "*D. 0,102.*"
            ],
            "solution": "• Phản ứng chuẩn độ: HCl + NaOH → NaCl + H~2~O.\n• Tại điểm tương đương: n(HCl) = n(NaOH)\n  ⇒ C~M~(HCl) × V(HCl) = C~M~(NaOH) × V(NaOH)\n  ⇒ C~M~(HCl) × 40 mL = 0,12 M × 34 mL\n  ⇒ C~M~(HCl) = (0,12 × 34) / 40 = 0,102 M.\n*Đáp án đúng: D.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10. Đặc điểm chung của các phân tử hợp chất hữu cơ là:\n(1) thường chứa nguyên tố C.\n(2) có thể chứa nguyên tố khác như Cl, N, P, O.\n(3) liên kết hóa học chủ yếu là liên kết cộng hoá trị.\n(4) liên kết hoá học chủ yếu là liên kết ion.\n(5) khó bay hơi, khó cháy.\n(6) phản ứng hoá học xảy ra nhanh.\n(7) nhiệt độ nóng chảy, nhiệt độ sôi thấp.\nCác phát biểu đúng là:",
            "options": [
                "A. 2, 3, 7.",
                "B. 1, 2, 3, 6.",
                "*C. 1, 2, 3, 7.*",
                "D. 4, 5, 6, 7."
            ],
            "solution": "Phân tích các đặc điểm chung của hợp chất hữu cơ:\n- (1) Đúng: Hợp chất hữu cơ luôn chứa carbon.\n- (2) Đúng: Thường chứa thêm H, O, N, halogen (Cl, Br), S, P,...\n- (3) Đúng: Liên kết hoá học chủ yếu trong phân tử là liên kết cộng hoá trị.\n- (4) Sai: Không phải liên kết ion.\n- (5) Sai: Hợp chất hữu cơ đa phần dễ bay hơi và dễ cháy.\n- (6) Sai: Phản ứng hữu cơ thường xảy ra chậm, không hoàn toàn, tạo hỗn hợp nhiều sản phẩm.\n- (7) Đúng: Nhiệt độ nóng chảy và nhiệt độ sôi nói chung thấp hơn hợp chất vô cơ.\n⇒ Các phát biểu đúng là: (1), (2), (3), (7).\n*Đáp án đúng: C.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11. Để xác định nhóm chức cho phân tử hợp chất hữu cơ, người ta dùng phương pháp:",
            "options": [
                "A. phổ gamma.",
                "B. phổ khối lượng MS.",
                "*C. phổ hồng ngoại IR.*",
                "D. phổ cực tím."
            ],
            "solution": "Phương pháp phổ hồng ngoại (IR - Infrared Spectroscopy) dựa trên sự hấp thụ bức xạ hồng ngoại bởi các dao động liên kết đặc trưng, được sử dụng để xác định sự có mặt của các nhóm chức trong phân tử hợp chất hữu cơ (Phổ MS dùng xác định phân tử khối).\n*Đáp án đúng: C.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12. Phú dưỡng là hiện tượng xảy ra do sự gia tăng hàm lượng của nguyên tố nào trong nước?",
            "options": [
                "A. Ca, Mg.",
                "B. Cl, F.",
                "*C. N, P.*",
                "D. Fe, Mn."
            ],
            "solution": "Hiện tượng phú dưỡng (eutrophication) là hậu quả của việc tích tụ dư thừa quá mức các chất dinh dưỡng chứa nguyên tố nitrogen (N) và phosphorus (P) trong các thủy vực ao, hồ.\n*Đáp án đúng: C.*"
        }
    ],

    # PHẦN II. (2 câu – 2 điểm) Câu trắc nghiệm đúng sai.
    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1. Để tách và tinh chế các chất trong hỗn hợp người ta dùng các phương pháp tách và tinh chế hợp chất hữu cơ: chưng cất, chiết, kết tinh, sắc kí, ...",
            "statements": [
                {"lbl": "a)", "text": "Phân tích thổ nhưỡng thuộc phương pháp chiết lỏng - rắn.", "ans": "Đúng"},
                {"lbl": "b)", "text": "Ngâm hoa quả làm siro thuộc phương pháp chiết lỏng- lỏng.", "ans": "Sai"},
                {"lbl": "c)", "text": "Nấu rượu uống thuộc phương pháp sắc kí.", "ans": "Sai"},
                {"lbl": "d)", "text": "Làm đường từ mía thuộc phương pháp chưng cất.", "ans": "Sai"}
            ],
            "solution": "- Ý a) ĐÚNG: Phân tích thổ nhưỡng dùng dung môi lỏng để hòa tan và rút các chất tan từ mẫu đất rắn, đây là phương pháp chiết lỏng – rắn.\n- Ý b) SAI: Ngâm hoa quả tươi (pha rắn) với đường để rút dịch quả ra ngoài là phương pháp chiết lỏng – rắn (hòa tan khuếch tán), không phải chiết lỏng – lỏng.\n- Ý c) SAI: Nấu rượu thủ công là quá trình đun sôi hỗn hợp rượu lên men và ngưng tụ hơi cồn, thuộc phương pháp chưng cất, không phải sắc kí.\n- Ý d) SAI: Sản xuất đường saccharose từ nước mía là quá trình làm sạch, cô đặc và kết tinh tinh thể đường, thuộc phương pháp kết tinh, không phải chưng cất."
        },
        {
            "id": 2,
            "prompt": "Câu 2. Cho sơ đồ chuyển hoá sau:\nNH~3~ ⎯⎯⎯⎯⎯⎯→ X ⎯⎯⎯⎯→ Y ⎯⎯⎯⎯⎯⎯⎯⎯→ Z ⎯⎯⎯⎯⎯→ T ⎯⎯⎯→ P.\nBiết X, Y, Z, T, P là các hợp chất chứa nitrogen. Hãy cho biết những phát biểu sau là đúng hay sai?",
            "statements": [
                {"lbl": "a)", "text": "NH~3~ thể hiện tính khử.", "ans": "Đúng"},
                {"lbl": "b)", "text": "Z và P lần lượt là HNO~3~ và NaNO~2~.", "ans": "Đúng"},
                {"lbl": "c)", "text": "Y và Z lần lượt là NO~2~ và HNO~3~.", "ans": "Đúng"},
                {"lbl": "d)", "text": "X và T lần lượt N~2~ và NaNO~3~.", "ans": "Sai"}
            ],
            "solution": "Xác định các chất trong chuỗi sản xuất nitric acid và muối nitrit từ ammonia:\n• (1) 4NH~3~ + 5O~2~ ⎯⎯⎯⎯⎯⎯→ 4NO + 6H~2~O  ⇒ X là NO (nitrogen monoxide).\n• (2) 2NO + O~2~ → 2NO~2~  ⇒ Y là NO~2~ (nitrogen dioxide).\n• (3) 4NO~2~ + O~2~ + 2H~2~O → 4HNO~3~  ⇒ Z là HNO~3~ (nitric acid).\n• (4) HNO~3~ + NaOH → NaNO~3~ + H~2~O  ⇒ T là NaNO~3~ (sodium nitrate).\n• (5) 2NaNO~3~ ⎯⎯→ 2NaNO~2~ + O~2~  ⇒ P là NaNO~2~ (sodium nitrite).\nPhân tích các phát biểu:\n- Ý a) ĐÚNG: Trong phản ứng (1), số oxi hóa của N tăng từ -3 (trong NH~3~) lên +2 (trong NO), NH~3~ là chất khử (thể hiện tính khử).\n- Ý b) ĐÚNG: Z là HNO~3~ và P là NaNO~2~.\n- Ý c) ĐÚNG: Y là NO~2~ và Z là HNO~3~.\n- Ý d) SAI: X là NO (không phải N~2~), T là NaNO~3~."
        }
    ],

    # PHẦN III. (4 câu – 2 điểm) Câu trắc nghiệm trả lời ngắn.
    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1. Ammonia được điều chế bằng phản ứng: N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)\nỞ t °C, nồng độ các chất ở trạng thái cân bằng là: [N~2~] = 0,45 M, [H~2~] = 0,14 M, [NH~3~] = 0,62 M.\nTính hằng số cân bằng K~c~ của phản ứng trên tại t °C? (kết quả làm tròn tới hàng đơn vị)",
            "ans": "311",
            "solution": "• Biểu thức hằng số cân bằng của phản ứng:\n  K~c~ = [NH~3~]^2^ / ([N~2~] × [H~2~]^3^)\n• Thay số nồng độ cân bằng:\n  K~c~ = (0,62)^2^ / (0,45 × 0,14^3^)\n       = 0,3844 / (0,45 × 0,002744)\n       = 0,3844 / 0,0012348 ≈ 311,305\n• Làm tròn tới hàng đơn vị: K~c~ = 311.\n*Trả lời: 311.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Cho dung dịch NH~3~ lần lượt vào các dung dịch sau: Ca(OH)~2~, HCl, BaCl~2~, CuSO~4~, CH~3~COOH, Al(NO~3~)~3~. Có bao nhiêu trường hợp có phản ứng xảy ra?",
            "ans": "4",
            "solution": "Xét tương tác của dung dịch NH~3~ với từng chất:\n1. HCl: NH~3~ + HCl → NH~4~Cl (phản ứng acid - base) ⇒ CÓ PHẢN ỨNG.\n2. CuSO~4~: 2NH~3~ + 2H~2~O + CuSO~4~ → Cu(OH)~2~↓ + (NH~4~)~2~SO~4~, sau đó Cu(OH)~2~ tan trong NH~3~ dư tạo phức chất [Cu(NH~3~)~4~](OH)~2~ màu xanh lam thẫm ⇒ CÓ PHẢN ỨNG.\n3. CH~3~COOH: NH~3~ + CH~3~COOH → CH~3~COONH~4~ (phản ứng acid - base) ⇒ CÓ PHẢN ỨNG.\n4. Al(NO~3~)~3~: Al(NO~3~)~3~ + 3NH~3~ + 3H~2~O → Al(OH)~3~↓ + 3NH~4~NO~3~ (Al(OH)~3~ không tan trong NH~3~ dư) ⇒ CÓ PHẢN ỨNG.\n5. Ca(OH)~2~ và BaCl~2~: Không xảy ra phản ứng.\n⇒ Có 4 trường hợp xảy ra phản ứng.\n*Trả lời: 4.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Eugenol là thành phần chính trong tinh dầu đinh hương hoặc tinh dầu hương nhu. Chất này được sử dụng là chất diệt nấm, dẫn dụ côn trùng. Phân tích phần trăm khối lượng các nguyên tố cho thấy eugenol có 73,17% carbon, 7,31% hydrogen, còn lại là oxygen. Lập công thức phân tử của eugenol, biết rằng kết quả phân tích phổ khối lượng cho thấy phân tử khối của eugenol là 164. Tổng số nguyên tử trong 1 phân tử eugenol là bao nhiêu?",
            "ans": "24",
            "solution": "• Phần trăm khối lượng oxygen trong eugenol:\n  %O = 100% - 73,17% - 7,31% = 19,52%.\n• Đặt công thức phân tử của eugenol là C~x~H~y~O~z~ với phân tử khối M = 164 g/mol.\n  x = (164 × 73,17%) / (12 × 100%) = 120 / 12 = 10 nguyên tử C.\n  y = (164 × 7,31%) / (1 × 100%) = 12 / 1 = 12 nguyên tử H.\n  z = (164 × 19,52%) / (16 × 100%) = 32 / 16 = 2 nguyên tử O.\n• Công thức phân tử của eugenol là C~10~H~12~O~2~.\n• Tổng số nguyên tử trong 1 phân tử eugenol = 10 + 12 + 2 = 24.\n*Trả lời: 24.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Trong công nghiệp, sulfuric acid được sản xuất từ quặng pirite sắt có thành phần chính là FeS~2~, theo sơ đồ sau: FeS~2~ → SO~2~ → SO~3~ → H~2~SO~4~.\nTính khối lượng (theo đơn vị tấn) H~2~SO~4~ 98% điều chế được từ 1 tấn quặng chứa 60% FeS~2~ (biết hiệu suất cả quá trình là 80%)?",
            "ans": "0,8",
            "solution": "• Khối lượng FeS~2~ nguyên chất trong 1 tấn quặng:\n  m(FeS~2~) = 1 tấn × 60% = 0,6 tấn.\n• Sơ đồ bảo toàn nguyên tố sulfur (S):\n  FeS~2~ ⎯⎯→ 2H~2~SO~4~\n  120 tấn        2 × 98 = 196 tấn (lý thuyết)\n  0,6 tấn        m(H~2~SO~4~ lí thuyết) = (0,6 × 196) / 120 = 0,98 tấn.\n• Khối lượng H~2~SO~4~ thực tế thu được với hiệu suất H = 80%:\n  m(H~2~SO~4~ thực tế) = 0,98 tấn × 80% = 0,784 tấn.\n• Khối lượng dung dịch H~2~SO~4~ 98% điều chế được:\n  m(dd H~2~SO~4~ 98%) = 0,784 / 98% = 0,8 tấn.\n*Trả lời: 0,8.*"
        }
    ],

    # B. PHẦN TỰ LUẬN: (3 câu – 3 điểm)
    "part4": [
        {
            "id": 1,
            "score": "1,0 điểm",
            "prompt": "Câu 1: (1 điểm) Cho cân bằng hoá học sau: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g)   Δ~r~H°~298~ < 0\nHãy cho biết cân bằng hoá học sẽ chuyển dịch theo chiều nào và giải thích khi tác động vào cân bằng hoá học một trong các yếu tố sau:\na) tăng nhiệt độ.\nb) tăng áp suất chung của hệ phản ứng.\nc) dùng thêm chất xúc tác V~2~O~5~.\nd) tăng nồng độ O~2~.",
            "solution": "Theo nguyên lí chuyển dịch cân bằng Le Chatelier:\na) Tăng nhiệt độ:\n- Chiều thuận là chiều tỏa nhiệt (Δ~r~H°~298~ < 0), chiều nghịch là chiều thu nhiệt (Δ~r~H°~298~ > 0).\n- Khi tăng nhiệt độ, cân bằng chuyển dịch theo chiều thu nhiệt nhằm làm giảm nhiệt độ của hệ, tức là chuyển dịch theo chiều nghịch.\nb) Tăng áp suất chung của hệ:\n- Tổng số mol khí ở vế trái là 3 mol (2 mol SO~2~ + 1 mol O~2~), vế phải là 2 mol SO~3~.\n- Khi tăng áp suất chung, cân bằng chuyển dịch theo chiều làm giảm số mol phân tử khí nhằm làm giảm áp suất của hệ, tức là chuyển dịch theo chiều thuận.\nc) Dùng thêm chất xúc tác V~2~O~5~:\n- Chất xúc tác làm tăng đồng thời tốc độ phản ứng thuận và tốc độ phản ứng nghịch với cùng một tỉ lệ, không làm thay đổi hằng số cân bằng.\n- Do đó cân bằng không bị chuyển dịch (chỉ giúp hệ phản ứng nhanh chóng đạt tới trạng thái cân bằng hơn).\nd) Tăng nồng độ O~2~:\n- Khi tăng nồng độ chất phản ứng O~2~, cân bằng chuyển dịch theo chiều làm giảm bớt nồng độ O~2~ (tiêu thụ O~2~), tức là chuyển dịch theo chiều thuận."
        },
        {
            "id": 2,
            "score": "1,0 điểm",
            "prompt": "Câu 2: (1 điểm) Hoà tan hoàn toàn 20,8 gam hỗn hợp X gồm Fe và Cu trong V lít dung dịch HNO~3~ 1M (vừa đủ) thu được 7,437 lít khí NO (sản phẩm khử duy nhất, ở đkc).\na) Tính phần trăm khối lượng các kim loại trong hỗn hợp X?\nb) Tính V?",
            "solution": "• Số mol khí NO ở điều kiện chuẩn (đkc: 25 °C, 1 bar):\n  n(NO) = 7,437 / 24,79 = 0,3 mol.\n• Gọi x, y lần lượt là số mol của Fe và Cu trong hỗn hợp X (x, y > 0).\n- Khối lượng hỗn hợp kim loại:\n  56x + 64y = 20,8  (1)\n- Quá trình oxi hóa - khử:\n  Fe → Fe^3+^ + 3e                 N^+5^ + 3e → N^+2^ (NO)\n   x     →       3x mol                       0,9 ← 0,3 mol\n  Cu → Cu^2+^ + 2e\n   y     →       2y mol\n  Bảo toàn electron: 3x + 2y = 3 × n(NO) = 3 × 0,3 = 0,9  (2)\n- Giải hệ phương trình (1) và (2):\n  Từ (2) ⇒ y = 0,45 - 1,5x. Thay vào (1):\n  56x + 64(0,45 - 1,5x) = 20,8 ⇒ -40x = -8 ⇒ x = 0,2 mol (Fe).\n  y = 0,45 - 1,5 × 0,2 = 0,15 mol (Cu).\n\na) Tính phần trăm khối lượng các kim loại:\n• m(Fe) = 0,2 mol × 56 g/mol = 11,2 gam.\n  %m(Fe) = (11,2 / 20,8) × 100% ≈ 53,85%.\n• m(Cu) = 0,15 mol × 64 g/mol = 9,6 gam.\n  %m(Cu) = 100% - 53,85% = 46,15%.\n\nb) Tính giá trị V:\n• Bảo toàn nguyên tố nitrogen (N):\n  n(HNO~3~ pư) = n(NO~3~^-^ trong muối) + n(NO)\n  Trong đó: n(NO~3~^-^ trong muối) = 3n(Fe) + 2n(Cu) = 3(0,2) + 2(0,15) = 0,9 mol.\n  ⇒ n(HNO~3~ pư) = 0,9 + 0,3 = 1,2 mol.\n  (Hoặc áp dụng công thức: n(HNO~3~) = 4n(NO) = 4 × 0,3 = 1,2 mol).\n• Thể tích dung dịch HNO~3~ 1M đã dùng:\n  V = n / C~M~ = 1,2 mol / 1 M = 1,2 lít."
        },
        {
            "id": 3,
            "score": "1,0 điểm",
            "prompt": "Câu 3: (1 điểm) Trộn 500ml dung dịch hỗn hợp NaOH 0,2M và Ba(OH)~2~ 0,1M với V ml dung dịch H~2~SO~4~ 0,3M. Sau phản ứng thu được m gam kết tủa và dung dịch A có pH = 1. Coi thể tích dung dịch thay đổi không đáng kể. Tính giá trị V và m?",
            "solution": "• Đổi 500 mL = 0,5 L.\n  Số mol các chất trong 500 mL dung dịch kiềm:\n  n(NaOH) = 0,5 × 0,2 = 0,1 mol.\n  n(Ba(OH)~2~) = 0,5 × 0,1 = 0,05 mol ⇒ n(Ba^2+^) = 0,05 mol.\n  Tổng số mol ion OH^-^ ban đầu:\n  n(OH^-^) = n(NaOH) + 2n(Ba(OH)~2~) = 0,1 + 2 × 0,05 = 0,2 mol.\n\n• Gọi thể tích dung dịch H~2~SO~4~ 0,3M là V~L~ (lít) ⇒ V (mL) = 1000 × V~L~.\n  n(H~2~SO~4~) = 0,3V~L~ (mol) ⇒ n(H^+^ ban đầu) = 2 × 0,3V~L~ = 0,6V~L~ (mol).\n  n(SO~4~^2-^ ban đầu) = 0,3V~L~ (mol).\n\n• Phản ứng trung hòa: H^+^ + OH^-^ → H~2~O.\n  Dung dịch A thu được có pH = 1 < 7 ⇒ Dung dịch có môi trường acid, H^+^ dư, OH^-^ phản ứng hết.\n  Nồng độ ion H^+^ dư: [H^+^ dư] = 10^-pH^ = 10^-1^ = 0,1 M.\n  Tổng thể tích dung dịch sau khi trộn: V~dd~ = 0,5 + V~L~ (lít).\n  Số mol H^+^ dư trong dung dịch A:\n  n(H^+^ dư) = [H^+^ dư] × V~dd~ = 0,1 × (0,5 + V~L~) = 0,05 + 0,1V~L~ (mol).\n\n• Ta có phương trình cân bằng số mol H^+^:\n  n(H^+^ dư) = n(H^+^ ban đầu) - n(OH^-^ đã phản ứng)\n  ⇒ 0,05 + 0,1V~L~ = 0,6V~L~ - 0,2\n  ⇒ 0,5V~L~ = 0,25 ⇒ V~L~ = 0,5 lít.\n  Vậy giá trị thể tích V = 0,5 × 1000 = 500 mL.\n\n• Tính khối lượng kết tủa BaSO~4~:\n  Phản ứng tạo kết tủa: Ba^2+^ + SO~4~^2-^ → BaSO~4~↓\n  Số mol Ba^2+^ ban đầu = 0,05 mol.\n  Số mol SO~4~^2-^ ban đầu = 0,3 × V~L~ = 0,3 × 0,5 = 0,15 mol.\n  So sánh: n(Ba^2+^) = 0,05 mol < n(SO~4~^2-^) = 0,15 mol ⇒ SO~4~^2-^ dư, Ba^2+^ kết tủa hết.\n  Số mol kết tủa BaSO~4~: n(BaSO~4~) = 0,05 mol.\n  Khối lượng kết tủa thu được:\n  m = m(BaSO~4~) = 0,05 mol × 233 g/mol = 11,65 gam.\n\n⇒ Kết luận: V = 500 mL và m = 11,65 gam."
        }
    ]
}

def generate_exam_doc(is_teacher=False):
    """Tạo tài liệu Word cho đề kiểm tra mã 1103 (Bản Học Sinh hoặc Bản Giáo Viên)."""
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
        add_quick_answer_table_1103(doc)

    # 3. PHẦN TRẮC NGHIỆM: (7 điểm)
    add_section_header(
        doc,
        "A. PHẦN TRẮC NGHIỆM: (7,0 điểm)",
        "PHẦN I. (12 câu – 3,0 điểm) Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi chỉ chọn một phương án."
    )
    for q in DATA_EXAM["part1"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"])

        # Các phương án
        render_options_no_table(doc, q["options"], is_teacher=is_teacher)

        # Lời giải chi tiết
        if is_teacher and "solution" in q:
            add_solution_no_table(doc, q["solution"])

    # 4. PHẦN II
    add_section_header(
        doc,
        "PHẦN II. (2 câu – 2,0 điểm) Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )
    for q in DATA_EXAM["part2"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"])

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
        "PHẦN III. (4 câu – 2,0 điểm) Câu trắc nghiệm trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 4."
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

    # 6. B. PHẦN TỰ LUẬN: (3 câu – 3 điểm)
    add_section_header(
        doc,
        "B. PHẦN TỰ LUẬN: (3 câu – 3,0 điểm)",
        "Thí sinh trình bày chi tiết lời giải các bài toán sau vào giấy làm bài."
    )
    for q in DATA_EXAM["part4"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"])

        if is_teacher:
            add_solution_no_table(doc, q["solution"])
        else:
            add_student_writing_lines(doc, num_lines=8)

    # Dòng kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(10)
    p_end.paragraph_format.space_after = Pt(10)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "---------- *HẾT* ----------", base_bold=True)

    # Lưu file
    suffix = "GiaoVien" if is_teacher else "HocSinh"
    out_file = os.path.join(DIR_SAN_PHAM, f"De_Kiem_Tra_CK1_Hoa_11_NguyenDinhChieu_Ma1103_{suffix}.docx")
    try:
        doc.save(out_file)
        print(f"Đã tạo thành công: {out_file}")
    except PermissionError:
        print(f"File {out_file} đang mở trong ứng dụng khác, bỏ qua cập nhật file này.")
    return out_file

if __name__ == "__main__":
    generate_exam_doc(is_teacher=False)
    generate_exam_doc(is_teacher=True)
