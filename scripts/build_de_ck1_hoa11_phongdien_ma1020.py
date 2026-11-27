# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Kì I - Hóa học 11 - Trường THPT Phong Điền (TP Huế)
Năm học 2025 - 2026 - Mã đề 1020
Quy chuẩn FORM ĐỀ KIỂM TRA của Thầy TRẦN VĂN THẠNH
Tạo 2 phiên bản:
- Bản Học Sinh: De_Kiem_Tra_CK1_Hoa_11_PhongDien_Ma1020.docx
- Bản Giáo Viên: De_Kiem_Tra_CK1_Hoa_11_PhongDien_Ma1020_GiaoVien.docx (có BẢNG ĐÁP ÁN RÚT GỌN CẢ TỰ LUẬN & lời giải chi tiết)
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
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_phongdien_ma1020")
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

def add_quick_answer_table_1020(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Giáo viên ở đầu tài liệu Mã đề 1020.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 1020)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I (12 câu x 1 hàng)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    h1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
    a1 = ["Đ/A", "A", "C", "C", "A", "C", "A", "A", "B", "C", "B", "D", "B"]

    w1_c0 = Cm(1.8)
    w1_c = Cm(1.35)

    for r_idx, r_data in enumerate([h1, a1]):
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
        ["Câu 1", "Thí nghiệm đốt cháy hexane trong bát sứ", "ĐÚNG", "ĐÚNG", "SAI", "SAI"],
        ["Câu 2", "Thí nghiệm chưng cất tách ethanol và nước", "SAI", "ĐÚNG", "SAI", "ĐÚNG"]
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
    render_rich_text(p_p3, "*3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM - MỖI CÂU 0,5 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False

    tbl3_headers = ["Câu", "Câu 1", "Câu 2", "Câu 3", "Câu 4"]
    tbl3_answers = ["Đáp án", "2 - 3 - 1", "5", "6, 8, 6", "4"]

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

    tbl4 = doc.add_table(rows=6, cols=3)
    tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl4.autofit = False

    tbl4_data = [
        ("Câu", "Điểm", "Đáp số cốt lõi & Hướng dẫn chấm"),
        ("Câu 1", "0,5 đ", "• Luận điểm 1: Các nguyên tử liên kết theo đúng hóa trị và trật tự nhất định (cấu tạo hóa học). Thay đổi trật tự tạo chất mới.\n"
                           "• Luận điểm 2: Carbon có hóa trị 4, các nguyên tử C liên kết trực tiếp với nhau tạo mạch carbon (thẳng, nhánh, vòng).\n"
                           "• Luận điểm 3: Tính chất phụ thuộc vào thành phần và cấu tạo hóa học, cũng như ảnh hưởng qua lại giữa các nguyên tử."),
        ("Câu 2", "0,5 đ", "• B1: Chọn mạch C dài nhất, nhiều nhánh nhất làm mạch chính.\n"
                           "• B2: Đánh số C mạch chính từ đầu gần nhánh hơn (tổng chỉ số vị trí nhánh nhỏ nhất).\n"
                           "• B3: Tên = Vị trí nhánh - Tên nhánh + Tên alkane mạch chính."),
        ("Câu 3", "0,5 đ", "a) Cu(OH)~2~ + H~2~SO~4~(đặc, nóng) → CuSO~4~ + 2H~2~O (0,25 đ)\n"
                           "b) 2Fe + 6H~2~SO~4~(đặc, nóng) —(t°)→ Fe~2~(SO~4~)~3~ + 3SO~2~↑ + 6H~2~O (0,25 đ)"),
        ("Câu 4", "0,5 đ", "• Dựa vào: Sự đổi màu đột ngột của chất chỉ thị màu acid - base (như phenolphthalein hoặc methyl da cam).\n"
                           "• Vì sao: Tại điểm tương đương, phản ứng vừa đủ xảy ra bước nhảy pH đột ngột, làm chất chỉ thị biến đổi cấu trúc và đổi màu rõ rệt."),
        ("Câu 5", "1,0 đ", "1. a) Phản ứng tỏa nhiệt (vì ΔH < 0). (0,25 đ)\n"
                           "   b) Không dùng nhiệt độ quá cao vì phản ứng tỏa nhiệt, nhiệt độ cao làm cân bằng dịch chuyển theo chiều nghịch (giảm hiệu suất SO~3~), đồng thời giảm tuổi thọ xúc tác V~2~O~5~. (0,25 đ)\n"
                           "2. Điều chỉnh để tăng hiệu suất: Hạ nhiệt độ vừa phải (~450 °C); Tăng áp suất; Tăng nồng độ SO~2~ và O~2~ (dùng dư không khí); Liên tục tách SO~3~ ra khỏi hệ. (0,5 đ)")
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


def add_ir_table_header(doc):
    """Thêm Bảng tín hiệu phổ hồng ngoại của một số nhóm chức cơ bản ở đầu đề thi."""
    p_ir_lbl = doc.add_paragraph()
    p_ir_lbl.paragraph_format.space_before = Pt(4)
    p_ir_lbl.paragraph_format.space_after = Pt(2)
    p_ir_lbl.paragraph_format.left_indent = Cm(0.0)
    render_rich_text(p_ir_lbl, "_Bảng tín hiệu phổ hồng ngoại của một số nhóm chức cơ bản_", base_size=10, base_italic=True)

    tbl = doc.add_table(rows=5, cols=6)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False

    headers = ["Hợp chất", "Liên kết", "Số sóng (cm^-1^)", "Hợp chất", "Liên kết", "Số sóng (cm^-1^)"]
    row_data = [
        ["Alcohol", "O – H", "3500 – 3200", "Carboxylic acid", "C = O", "1760 – 1690"],
        ["Ketone", "C = O", "1715 – 1666", "", "O – H", "3300 – 2500"],
        ["Aldehyde", "C – H", "2830 – 2695", "Ester", "C = O", "1750 – 1715"],
        ["", "C = O", "1740 – 1685", "", "C – O", "1300 – 1000"]
    ]

    col_widths = [Cm(3.0), Cm(2.2), Cm(3.8), Cm(3.2), Cm(2.2), Cm(3.6)]

    # Header row
    hdr_cells = tbl.rows[0].cells
    for c_idx, text in enumerate(headers):
        cell = hdr_cells[c_idx]
        cell.width = col_widths[c_idx]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        set_cell_shading(cell, "E8F5E9")
        render_rich_text(p, f"*{text}*", base_size=9.5, base_bold=True, base_color=COLOR_PRIMARY)
        set_cell_border(cell,
                        top=("single", "4", "0", "8BC390"),
                        bottom=("single", "4", "0", "8BC390"),
                        left=("single", "4", "0", "8BC390"),
                        right=("single", "4", "0", "8BC390"))

    # Data rows
    for r_idx, vals in enumerate(row_data):
        row = tbl.rows[r_idx + 1]
        for c_idx, val in enumerate(vals):
            cell = row.cells[c_idx]
            cell.width = col_widths[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx in [1, 2, 4, 5] else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.space_after = Pt(1.5)
            render_rich_text(p, val, base_size=9.5)
            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))


# Cấu trúc dữ liệu Đề thi mã 1020
DATA_EXAM = {
    "school_line": "SỞ GD&ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT PHONG ĐIỀN",
    "exam_title": "ĐỀ KIỂM TRA CUỐI KÌ I – NĂM HỌC 2025-2026",
    "subject_line": "MÔN: HÓA HỌC LỚP 11",
    "duration": "45 phút (không kể thời gian phát đề)",
    "num_pages": "04",
    "exam_code": "1020",
    "atomic_masses": "Na = 23; Cu = 64; Ag = 108; Fe = 56; S = 32; N = 14; O = 16; H = 1; Cl = 35,5; Br = 80",

    # A. PHẦN TRẮC NGHIỆM (7,0 điểm)
    # Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm) - 12 câu
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1: Để pha loãng H~2~SO~4~ đặc cách làm nào sau đây đúng?",
            "image": "pha_loang_h2so4_clean.png",
            "image_width": 4.5,
            "options": [
                "*A. cách 1.*",
                "B. cách 2.",
                "C. cách 3.",
                "D. cách 1 và 2."
            ],
            "solution": "• Quy tắc pha loãng sulfuric acid đặc: Phải rót từ từ acid vào nước và khuấy nhẹ bằng đũa thủy tinh (cách 1), tuyệt đối không làm ngược lại (rót nước vào acid) vì acid đặc háo nước tỏa nhiệt rất mạnh làm nước sôi đột ngột và bắn acid ra ngoài gây bỏng nguy hiểm.\n*Đáp án đúng: A.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2: Hợp chất hữu cơ nào sau đây có mạch carbon phân nhánh?",
            "options": [
                "A. CH~3~–CH(OH)–CH~3~",
                "B. cyclo-C~4~H~8~ (mạch vòng)",
                "*C. CH~3~–CH(CH~3~)–CH~3~*",
                "D. CH~3~–CH~2~–CH~2~–CH~2~–CH~3~"
            ],
            "solution": "• Mạch carbon phân nhánh là mạch có nguyên tử carbon liên kết với từ 3 nguyên tử carbon khác trở lên (carbon bậc 3, 4 làm điểm rẽ nhánh).\n  - A: CH~3~-CH(OH)-CH~3~ có mạch chính 3 carbon không phân nhánh (nhóm -OH là nhóm chức).\n  - B: Mạch vòng cyclobutane.\n  - C: CH~3~-CH(CH~3~)-CH~3~ (isobutane) có 1 nhánh methyl (-CH~3~) gắn vào C số 2 ⇒ mạch phân nhánh.\n  - D: Pentane mạch thẳng không phân nhánh.\n*Đáp án đúng: C.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3: Kết tinh là phương pháp tách biệt và tinh chế hỗn hợp các chất ....(1) dựa vào độ tan khác nhau và sự thay đổi độ tan của chúng theo .....(2). (1) và (2) lần lượt là",
            "options": [
                "A. rắn – thời gian.",
                "B. lỏng – nhiệt độ.",
                "*C. rắn – nhiệt độ.*",
                "D. lỏng – thời gian."
            ],
            "solution": "• Theo định nghĩa SGK: Kết tinh là phương pháp tách biệt và tinh chế hỗn hợp các chất rắn (1) dựa vào độ tan khác nhau và sự thay đổi độ tan của chúng theo nhiệt độ (2).\n*Đáp án đúng: C.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4: Trước kia, “phẩm đỏ” dùng để nhuộm áo choàng cho các Hồng y giáo chủ được tách chiết từ một loài ốc biển. Đó là một hợp chất có thành phần nguyên tố như sau: 45,70 %C; 1,90 %H; 7,60 %O; 6,70 %N; 38,10 %Br. Công thức đơn giản của phẩm đỏ là",
            "options": [
                "*A. C~8~H~4~ONBr.*",
                "B. C~4~H~2~ONBr.",
                "C. C~4~H~8~O~2~NBr~2~.",
                "D. C~2~H~4~ONBr."
            ],
            "solution": "• Đặt công thức đơn giản nhất của phẩm đỏ là C~x~H~y~O~z~N~t~Br~u~.\n• Tỉ lệ mol giữa các nguyên tố:\n  x : y : z : t : u = (45,70 / 12) : (1,90 / 1) : (7,60 / 16) : (6,70 / 14) : (38,10 / 80)\n                   = 3,808 : 1,900 : 0,475 : 0,479 : 0,476\n• Chia cho số nhỏ nhất (0,475):\n  x : y : z : t : u ≈ 8 : 4 : 1 : 1 : 1\n⇒ Công thức đơn giản nhất là C~8~H~4~ONBr.\n*Đáp án đúng: A.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5: Điền vào khoảng trống bằng cụm từ thích hợp: “Cân bằng hóa học là cân bằng ...(1)... vì tại cân bằng phản ứng ...(2)...”",
            "options": [
                "A. (1) tĩnh; (2) dừng lại.",
                "B. (1) tính; (2) tiếp tục xảy ra.",
                "*C. (1) động; (2) tiếp tục xảy ra.*",
                "D. (1) động; (2) dừng lại."
            ],
            "solution": "• Cân bằng hóa học là cân bằng động vì tại trạng thái cân bằng, phản ứng thuận và phản ứng nghịch vẫn tiếp tục xảy ra với tốc độ bằng nhau (v~thuận~ = v~nghịch~ > 0).\n*Đáp án đúng: C.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6: Liên kết hóa học trong phân tử HNO~3~ là liên kết",
            "options": [
                "*A. cộng hóa trị có cực.*",
                "B. hydrogen.",
                "C. ion.",
                "D. cộng hóa trị không cực."
            ],
            "solution": "• Trong phân tử HNO~3~, giữa các nguyên tử N, O và H liên kết với nhau bằng các cặp electron dùng chung bị lệch về phía nguyên tử có độ âm điện lớn hơn (O) ⇒ Liên kết cộng hóa trị có cực.\n*Đáp án đúng: A.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7: Phát biểu nào sau đây đúng khi nói về sự điện li?",
            "options": [
                "*A. Sự điện li là quá trình phân li một chất trong nước thành ion.*",
                "B. Sự điện li là quá trình phân li một chất dưới tác dụng của dòng điện.",
                "C. Sự điện li quá trình hoà tan một chất vào nước tạo thành dung dịch.",
                "D. Sự điện li thực chất là quá trình oxi hoá - khử."
            ],
            "solution": "• Định nghĩa: Sự điện li là quá trình phân li các chất khi tan trong nước thành các ion.\n*Đáp án đúng: A.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8: Cho hình vẽ mô tả thí nghiệm như sau:\n(Hình vẽ mô tả thí nghiệm đài phun nước của ammonia với dung dịch phenolphthalein)\nHình vẽ mô tả thí nghiệm để chứng minh",
            "image": "dai_phun_nuoc_nh3_clean.png",
            "image_width": 2.8,
            "options": [
                "A. tính base của NH~3~.",
                "*B. tính tan nhiều trong nước và tính base của NH~3~.*",
                "C. tính tan nhiều trong nước của NH~3~.",
                "D. tính khử của NH~3~."
            ],
            "solution": "• Thí nghiệm đài phun nước của NH~3~:\n  - Khí NH~3~ tan rất nhiều trong nước làm giảm áp suất trong bình cầu, hút nước từ chậu dưới phun mạnh thành tia vào bình cầu ⇒ chứng minh tính tan nhiều trong nước của NH~3~.\n  - Nước phun vào bình chuyển sang màu hồng/đỏ do NH~3~ tan trong nước tạo môi trường kiềm (NH~3~ + H~2~O ⇌ NH~4~^+^ + OH^-^) làm đổi màu phenolphthalein ⇒ chứng minh tính base của NH~3~.\n*Đáp án đúng: B.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9: Dãy nào sau đây gồm các chất đều thuộc dãy đồng đẳng của methane?",
            "options": [
                "A. C~2~H~2~, C~3~H~4~, C~4~H~6~.",
                "B. CH~4~, C~2~H~2~, C~3~H~4~.",
                "*C. C~2~H~6~, C~3~H~8~, C~5~H~12~.*",
                "D. CH~4~, C~2~H~6~, C~4~H~14~."
            ],
            "solution": "• Dãy đồng đẳng của methane là các alkane có công thức chung C~n~H~2n+2~ (n ≥ 1):\n  C~2~H~6~ (n=2), C~3~H~8~ (n=3), C~5~H~12~ (n=5) đều thỏa mãn công thức chung C~n~H~2n+2~.\n*Đáp án đúng: C.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10: Tác nhân chính gây ra hiện tượng mưa acid là",
            "options": [
                "A. NH~3~, NO~2~.",
                "*B. NO~x~, SO~2~.*",
                "C. CO, SO~2~.",
                "D. CO, NH~3~."
            ],
            "solution": "• Các oxide của sulfur (SO~2~) và nitrogen (NO~x~ gồm NO, NO~2~) trong khí quyển bị oxi hóa và kết hợp với nước mưa tạo ra H~2~SO~4~ và HNO~3~, là nguyên nhân chính gây nên hiện tượng mưa acid.\n*Đáp án đúng: B.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11: Ngoài đơn chất nitrogen thì nguyên tố nitrogen còn tồn tại dưới dạng nào?",
            "options": [
                "A. Khoáng vật sodium nitrate (NaNO~3~) gọi là diêm tiêu sodium",
                "B. Có trong protein",
                "C. Có trong nucleic acid",
                "*D. Có trong diêm tiêu sodium, protein, nucleic acid...*"
            ],
            "solution": "• Trong tự nhiên, nguyên tố nitrogen tồn tại ở cả dạng đơn chất (chiếm 78% thể tích không khí) và dạng hợp chất: khoáng vật diêm tiêu (NaNO~3~), cũng như là thành phần cấu tạo thiết yếu của các hợp chất sinh học hữu cơ như amino acid, protein, nucleic acid (DNA, RNA)...\n*Đáp án đúng: D.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12: Hợp chất hữu cơ là",
            "options": [
                "A. Hợp chất khó tan trong nước.",
                "*B. Hợp chất của cacbon trừ CO, CO~2~, H~2~CO~3~, muối cacbonat kim loại...*",
                "C. Hợp chất của cacbon và một số nguyên tố khác trừ N, Cl, O.",
                "D. Hợp chất có nhiệt độ sôi cao."
            ],
            "solution": "• Định nghĩa: Hợp chất hữu cơ là hợp chất của carbon (trừ các hợp chất vô cơ như CO, CO~2~, H~2~CO~3~, muối carbonate kim loại, carbide, cyanide...).\n*Đáp án đúng: B.*"
        }
    ],

    # Phần II: Trắc nghiệm dạng Đúng/Sai (2,0 điểm) - 2 câu
    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1: Cho khoảng 1 mL hexane vào bát sứ nhỏ, cẩn thận đưa que đóm đang cháy vào bề mặt chất lỏng, hexane bốc cháy cho ngọn lửa màu vàng.",
            "statements": [
                ("a) Đốt hexane trong điều kiện thiếu oxygen sinh ra carbon monooxide.", "ĐÚNG"),
                ("b) Không nên đốt hexane trong cốc thuỷ tinh do phản ứng này toả ra lượng nhiệt lớn có thể làm vỡ cốc thuỷ tinh.", "ĐÚNG"),
                ("c) Phản ứng trên là phản ứng thu nhiệt.", "SAI"),
                ("d) Hexane thành phần chủ yếu của khí dầu mỏ hóa lỏng (LPG).", "SAI")
            ],
            "solution": "• Ý a: ĐÚNG. Khi đốt cháy hydrocarbon trong điều kiện thiếu oxygen sẽ xảy ra quá trình cháy không hoàn toàn, tạo ra khí độc carbon monoxide (CO) và muội than (C).\n• Ý b: ĐÚNG. Phản ứng cháy của hexane tỏa nhiệt lượng rất lớn, nếu đốt trong cốc thủy tinh thông thường có thể gây giãn nở nhiệt không đều làm vỡ cốc thủy tinh nguy hiểm.\n• Ý c: SAI. Phản ứng cháy của các hydrocarbon luôn là phản ứng tỏa nhiệt mạnh (ΔH < 0).\n• Ý d: SAI. Khí dầu mỏ hóa lỏng (LPG) có thành phần chủ yếu là propane (C~3~H~8~) và butane (C~4~H~10~). Hexane (C~6~H~14~) ở điều kiện thường là chất lỏng, không phải thành phần của LPG."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Thực hiện thí nghiệm tách cồn (ethanol) ra khỏi hỗn hợp cồn – nước, được mô phỏng như hình sau:",
            "image": "thi_nghiem_chung_cat_clean.png",
            "image_width": 4.5,
            "statements": [
                ("a) Đầu nước vào và nước ra của ống sinh hàn có thể được đặt ngược lại.", "SAI"),
                ("b) Thí nghiệm đã sử dụng phương pháp chưng cất.", "ĐÚNG"),
                ("c) Phương pháp tách ethanol và nước dựa trên sự khác nhau về khối lượng riêng.", "SAI"),
                ("d) Độ cồn của sản phẩm sẽ lớn hơn so với độ cồn ban đầu do sản phẩm thu được tinh khiết hơn lẫn ít nước hơn cồn ban đầu.", "ĐÚNG")
            ],
            "solution": "• Ý a: SAI. Nước làm lạnh bắt buộc phải đi vào từ đầu thấp phía dưới và đi ra ở đầu cao phía trên để lòng ống sinh hàn luôn ngập đầy nước làm lạnh và tránh sốc nhiệt làm nứt ống sinh hàn.\n• Ý b: ĐÚNG. Thí nghiệm dùng nhiệt để làm bay hơi và ngưng tụ chất lỏng, đó chính là phương pháp chưng cất (chưng cất phân đoạn).\n• Ý c: SAI. Phương pháp chưng cất dựa trên sự khác nhau về nhiệt độ sôi của các chất (ethanol sôi ở 78,3 °C, nước sôi ở 100 °C), không phải dựa vào khối lượng riêng (chiết mới dựa vào khối lượng riêng và độ tan).\n• Ý d: ĐÚNG. Do ethanol có nhiệt độ sôi thấp hơn nước nên bay hơi trước và ngưng tụ vào bình hứng, làm cho phần dịch ngưng thu được có hàm lượng ethanol cao hơn hỗn hợp ban đầu, tức độ cồn tăng lên."
        }
    ],

    # Phần III: Trắc nghiệm trả lời ngắn (2,0 điểm) - 4 câu
    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1: Một cơ sở nhỏ sản xuất tinh dầu sả từ thân và lá sả tươi. Sau khi đun sôi hỗn hợp sả với nước, người ta thu được một hỗn hợp gồm nước và tinh dầu sả, trong đó tinh dầu sả không tan trong nước, nhẹ hơn nước và còn lẫn một lượng nhỏ tạp chất rắn từ nguyên liệu thực vật. Để tách riêng và tinh chế tinh dầu sả đạt độ tinh khiết cao, cách lựa chọn và ưu tiên thứ tự sử dụng các phương pháp hóa học nào sau đây là hợp lí nhất?\n"
                      "1. Làm khô và tinh chế bằng kết tinh nhằm loại bỏ tạp chất còn sót lại.\n"
                      "2. Chưng cất thường để thu hỗn hợp\n"
                      "3. chiết để tách tinh dầu khỏi nước",
            "answer": "2 - 3 - 1",
            "solution": "• Trình tự các bước thực hiện tách và tinh chế tinh dầu sả:\n  - Bước 1 (2): Tiến hành chưng cất lôi cuốn hơi nước để đưa hỗn hợp tinh dầu sả và hơi nước sang bộ làm lạnh ngưng tụ.\n  - Bước 2 (3): Hỗn hợp ngưng tụ phân thành 2 lớp lỏng, dùng phễu chiết để tách riêng lớp tinh dầu sả (nhẹ hơn, nổi lên trên) ra khỏi lớp nước.\n  - Bước 3 (1): Dùng chất hút ẩm (như Na~2~SO~4~ khan) để làm khô và loại bỏ tạp chất còn sót lại thu tinh dầu tinh khiết.\n⇒ Thứ tự hợp lý nhất là: 2 - 3 - 1 (hoặc 231).\n*Đáp án: 2 - 3 - 1*"
        },
        {
            "id": 2,
            "prompt": "Câu 2: Cho lần lượt các chất sau FeS, Fe~3~O~4~, Fe~2~O~3~, FeO, Fe(OH)~2~, Fe(OH)~3~, FeSO~4~, Fe~2~(SO~4~)~3~ tác dụng với dung dịch H~2~SO~4~ đặc nóng. Số phản ứng oxi hóa – khử là?",
            "answer": "5",
            "solution": "• Dung dịch H~2~SO~4~ đặc, nóng là chất oxi hóa rất mạnh, chỉ tham gia phản ứng oxi hóa – khử với các chất mà nguyên tố kim loại (Fe) hoặc phi kim (S) chưa đạt số oxi hóa cao nhất (+3 đối với Fe):\n  1) FeS: Fe(+2) và S(-2) đều có tính khử ⇒ Oxi hóa - khử.\n  2) Fe~3~O~4~: Có chứa Fe(+2) ⇒ Oxi hóa - khử.\n  3) Fe~2~O~3~: Fe đã đạt số oxi hóa cao nhất (+3) ⇒ Phản ứng trao đổi tạo muối và nước.\n  4) FeO: Fe(+2) có tính khử ⇒ Oxi hóa - khử.\n  5) Fe(OH)~2~: Fe(+2) có tính khử ⇒ Oxi hóa - khử.\n  6) Fe(OH)~3~: Fe(+3) ⇒ Phản ứng trao đổi.\n  7) FeSO~4~: Fe(+2) có tính khử ⇒ Oxi hóa - khử.\n  8) Fe~2~(SO~4~)~3~: Fe(+3) ⇒ Không phản ứng.\n⇒ Có 5 chất tham gia phản ứng oxi hóa - khử là: FeS, Fe~3~O~4~, FeO, Fe(OH)~2~, FeSO~4~.\n*Đáp án: 5*"
        },
        {
            "id": 3,
            "prompt": "Câu 3: Người ta xác định được % khối lượng các nguyên tố trong vitamin C: %C = 40,91%; %H = 4,545%; %O = 54,545%. Số nguyên tử của C, H, O trong CTPT vitamin C. biết M~vitamin C~ = 176.",
            "answer": "6, 8, 6",
            "solution": "• Đặt công thức phân tử của vitamin C là C~x~H~y~O~z~.\n• Số nguyên tử của từng nguyên tố trong phân tử:\n  x = (176 × 40,91%) / 12 = 72,00 / 12 = 6\n  y = (176 × 4,545%) / 1 = 8,00 / 1 = 8\n  z = (176 × 54,545%) / 16 = 96,00 / 16 = 6\n⇒ Công thức phân tử của vitamin C là C~6~H~8~O~6~.\n⇒ Số nguyên tử của C, H, O lần lượt là 6, 8, 6 (tổng số nguyên tử là 20).\n*Đáp án: 6, 8, 6*"
        },
        {
            "id": 4,
            "prompt": "Câu 4: Hợp chất 3-methylpentane có công thức cấu tạo dưới đây:\n"
                      "CH~3~ — CH~2~ — CH(CH~3~) — CH~2~ — CH~3~\n"
                      "Khi cho 3-methylpentane tác dụng với khí chlorine thu được số đồng phân cấu tạo là dẫn xuất monochloro tối đa là",
            "answer": "4",
            "solution": "• Phân tích cấu trúc đối xứng của 3-methylpentane:\n  C^1^H~3~ — C^2^H~2~ — C^3^H(C^6^H~3~) — C^4^H~2~ — C^5^H~3~\n• Phân tử có trục đối xứng đi qua C3 và nhóm methyl nhánh (C6):\n  - Vị trí C1 và C5 tương đương nhau ⇒ thế 1 Cl vào C1 (hoặc C5) cho sản phẩm: 1-chloro-3-methylpentane (1).\n  - Vị trí C2 và C4 tương đương nhau ⇒ thế 1 Cl vào C2 (hoặc C4) cho sản phẩm: 2-chloro-3-methylpentane (2).\n  - Thế 1 Cl vào C3 (nguyên tử C bậc 3) cho sản phẩm: 3-chloro-3-methylpentane (3).\n  - Thế 1 Cl vào C6 (nhóm methyl nhánh tại C3) cho sản phẩm: 1-chloro-2-ethylbutane (hay 3-(chloromethyl)pentane) (4).\n⇒ Tổng số đồng phân cấu tạo dẫn xuất monochloro tối đa thu được là 4.\n*Đáp án: 4*"
        }
    ],

    # B. PHẦN TỰ LUẬN (3,0 điểm) - 5 câu
    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1: (0,5 điểm) Trình bày các luận điểm chính của Thuyết cấu tạo hóa học?",
            "num_lines": 7,
            "solution": "Các luận điểm chính của Thuyết cấu tạo hóa học (Butlerov):\n"
                       "1. Trong phân tử hợp chất hữu cơ, các nguyên tử liên kết với nhau theo đúng hóa trị và theo một trật tự nhất định. Trật tự liên kết đó được gọi là cấu tạo hóa học. Sự thay đổi trật tự liên kết đó sẽ tạo ra chất mới.\n"
                       "2. Trong phân tử hợp chất hữu cơ, nguyên tố carbon luôn có hóa trị 4. Các nguyên tử carbon không những có thể liên kết với nguyên tử của các nguyên tố khác mà còn liên kết trực tiếp với nhau tạo thành các mạch carbon khác nhau (mạch hở không phân nhánh, mạch hở phân nhánh và mạch vòng).\n"
                       "3. Tính chất của các hợp chất hữu cơ phụ thuộc vào thành phần phân tử (bản chất các nguyên tố và số lượng nguyên tử) và cấu tạo hóa học (trật tự liên kết giữa các nguyên tử), đồng thời phụ thuộc vào sự ảnh hưởng qua lại giữa các nguyên tử trong phân tử."
        },
        {
            "id": 2,
            "prompt": "Câu 2: (0,5 điểm) Trình bày các bước gọi tên thay thế của Alkane phân nhánh?",
            "num_lines": 6,
            "solution": "Các bước gọi tên thay thế của alkane phân nhánh theo IUPAC:\n"
                       "• Bước 1 (Chọn mạch chính): Chọn mạch carbon dài nhất và có nhiều nhánh nhất làm mạch carbon chính.\n"
                       "• Bước 2 (Đánh số mạch chính): Đánh số thứ tự các nguyên tử carbon trên mạch chính từ đầu gần nhánh hơn, sao cho tổng chỉ số vị trí các nhánh là nhỏ nhất.\n"
                       "• Bước 3 (Gọi tên): Tên alkane = [Vị trí nhánh] - [Tên nhánh] + [Tên alkane mạch chính].\n"
                       "  (Lưu ý: Nếu có nhiều nhánh giống nhau dùng tiền tố di, tri, tetra...; nếu có nhiều nhánh khác nhau sắp xếp tên nhánh theo thứ tự chữ cái; giữa số và chữ dùng dấu gạch ngang '-', giữa các số dùng dấu phẩy ',')."
        },
        {
            "id": 3,
            "prompt": "Câu 3: (0,5 điểm) Hoàn thành các phương trình phản ứng sau:\n"
                      "a) Cu(OH)~2~ + H~2~SO~4~ (đặc, nóng) →\n"
                      "b) Fe + H~2~SO~4~ (đặc, nóng) →",
            "num_lines": 5,
            "solution": "Phương trình phản ứng xảy ra:\n"
                       "a) Cu(OH)~2~ + H~2~SO~4~(đặc, nóng) → CuSO~4~ + 2H~2~O\n"
                       "   (Đồng trong Cu(OH)~2~ đã có số oxi hóa cao nhất +2 nên đây là phản ứng acid - base tạo muối và nước).\n"
                       "b) 2Fe + 6H~2~SO~4~(đặc, nóng) —(t°)→ Fe~2~(SO~4~)~3~ + 3SO~2~↑ + 6H~2~O"
        },
        {
            "id": 4,
            "prompt": "Câu 4: (0,5 điểm) Chuẩn độ acid- base là phương pháp được sử dụng để xác định nồng độ acid hoặc base chưa biết nồng độ bằng dung dịch acid hoặc dung dịch base đã biết chính xác nồng độ. Khi đó dung dịch đã biết chính xác nồng độ gọi là dung dịch chuẩn. Khi chuẩn độ, người ta thêm từ từ dung dịch đựng trong burette vào dung dịch đựng trong bình tam giác. Tại điểm mà hai chất tác dụng vừa đủ với nhau gọi là điểm tương đương. Người ta thường dựa vào đặc điểm nào để nhận biết điểm tương đương trong phép chuẩn độ trên? Vì sao?",
            "num_lines": 7,
            "solution": "• Đặc điểm nhận biết điểm tương đương:\n  Người ta thường nhận biết điểm tương đương (hoặc điểm dừng chuẩn độ sát với điểm tương đương) dựa vào sự đổi màu đột ngột của chất chỉ thị màu acid - base (như phenolphthalein hoặc methyl da cam).\n"
                       "• Giải thích (Vì sao):\n  Tại điểm tương đương, lượng acid và base tác dụng vừa đủ với nhau, trong dung dịch xảy ra sự nhảy vọt pH rất mạnh (bước nhảy pH quanh điểm tương đương). Sự thay đổi pH đột ngột này làm cho chất chỉ thị màu acid - base bị biến đổi cấu trúc phân tử và đổi màu rõ rệt, giúp người làm thí nghiệm nhận biết chính xác thời điểm dừng chuẩn độ."
        },
        {
            "id": 5,
            "prompt": "Câu 5: (1,0 điểm) Một nhà máy sản xuất sulfuric acid áp dụng phương pháp tiếp xúc với các giai đoạn chính sau:\n"
                      "- Đốt Sulfur hoặc quặng pyrite tạo SO~2~\n"
                      "- Oxi hóa SO~2~ thành SO~3~ trên xúc tác V~2~O~5~\n"
                      "- Hấp thụ SO~3~ để thu sulfuric acid\n"
                      "Trong quá trình vận hành, nhà máy phải đồng thời đảm bảo hiệu suất cao, tiết kiệm năng lượng và giảm thiểu ô nhiễm môi trường.\n"
                      "1. Xét phản ứng oxi hóa: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g), (ΔH < 0)\n"
                      "a) Cho biết phản ứng trên là thu nhiệt hay tỏa nhiệt.\n"
                      "b) Giải thích vì sao trong thực tế không tiến hành phản ứng ở nhiệt độ quá cao, mặc dù tốc độ phản ứng sẽ tăng.\n"
                      "2. Để tăng hiệu suất tổng hợp SO~3~ người ta phải điều chỉnh các yếu tố: Nhiệt độ, Áp suất, Nồng độ SO~2~ và O~2~ như thế nào?",
            "num_lines": 12,
            "solution": "1.\n"
                       "a) Phản ứng có ΔH < 0 (Δ~r~H~298~^o^ = -198,4 kJ) nên đây là phản ứng tỏa nhiệt.\n"
                       "b) Giải thích: Theo nguyên lí chuyển dịch cân bằng Le Chatelier, vì phản ứng thuận tỏa nhiệt nên nếu tăng nhiệt độ lên quá cao, cân bằng sẽ chuyển dịch theo chiều nghịch (chiều thu nhiệt), làm giảm hiệu suất tạo thành SO~3~. Mặt khác, nhiệt độ quá cao sẽ làm giảm hoạt tính và độ bền của chất xúc tác V~2~O~5~, đồng thời gây lãng phí năng lượng. (Trong thực tế người ta chọn nhiệt độ khoảng 450 °C để vừa đảm bảo tốc độ phản ứng đủ nhanh vừa giữ hiệu suất chuyển hóa cao).\n\n"
                       "2. Điều chỉnh các yếu tố để tăng hiệu suất tổng hợp SO~3~:\n"
                       "• Nhiệt độ: Giảm nhiệt độ (vừa phải, duy trì nhiệt độ tối ưu khoảng 450 °C) để cân bằng chuyển dịch theo chiều thuận (chiều tỏa nhiệt).\n"
                       "• Áp suất: Tăng áp suất, vì phản ứng thuận làm giảm số mol khí (từ 3 mol khí phản ứng tạo 2 mol khí sản phẩm), tăng áp suất cân bằng chuyển dịch sang chiều làm giảm áp suất (chiều thuận) ⇒ tăng hiệu suất.\n"
                       "• Nồng độ SO~2~ và O~2~: Tăng nồng độ các chất phản ứng (thực tế người ta dùng dư không khí/dư O~2~) để cân bằng chuyển dịch theo chiều thuận tạo thêm SO~3~; đồng thời liên tục tách SO~3~ ra khỏi hỗn hợp (bằng cách hấp thụ vào H~2~SO~4~ đặc) để cân bằng tiếp tục dịch chuyển sang phải."
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
        add_quick_answer_table_1020(doc)

    # Thêm Bảng tín hiệu phổ hồng ngoại của một số nhóm chức cơ bản
    add_ir_table_header(doc)

    # A. PHẦN TRẮC NGHIỆM (7,0 điểm)
    add_section_header(
        doc,
        title="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,0 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
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
        title="PHẦN II. Câu trắc nghiệm đúng sai (2,0 điểm)",
        subtitle="Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai. Điểm tối đa của 01 câu hỏi là 1 điểm: 01 ý được 0,1đ; 02 ý được 0,25đ; 03 ý được 0,5đ; 04 ý được 1đ."
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

    # Phần III: Trắc nghiệm trả lời ngắn (2,0 điểm)
    add_section_header(
        doc,
        title="PHẦN III: Câu trắc nghiệm yêu cầu trả lời ngắn (2,0 điểm)",
        subtitle="Thí sinh làm tròn kết quả đến hàng phần mười."
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
        title="PHẦN IV: Tự luận (3,0 điểm)",
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
    filename = f"De_Kiem_Tra_CK1_Hoa_11_PhongDien_Ma1020{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo thành công: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
