# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Học Kỳ I - Hóa học 12 - Trường THPT Nguyễn Chí Thanh (TP Huế)
Mã đề: 1001 - Năm học 2025 - 2026
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
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_de1001")
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
    answers_p1 = ["Đ/A", "C", "C", "C", "A", "D", "C", "D", "C", "C", "D", "C", "B", "C", "D", "D", "B"]

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
        ["Câu 1", "Thế điện cực chuẩn & pin Fe - Cu", "SAI", "SAI", "ĐÚNG", "ĐÚNG"],
        ["Câu 2", "Polymer tơ capron từ caprolactam", "SAI", "ĐÚNG", "SAI", "SAI"]
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
    answers_p34 = ["74,1%", "4", "3", "165", "aminoethanoic acid", "1043,8 m³", "b = 14 bánh"]

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


DATA_EXAM_1001 = {
    "school_line": "SỞ GD VÀ ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT NGUYỄN CHÍ THANH",
    "exam_title": "ĐỀ KIỂM TRA CUỐI HỌC KÌ I\nNĂM HỌC 2025-2026",
    "subject_line": "MÔN: HÓA HỌC - LỚP 12",
    "num_pages": "04",
    "duration": "45 phút (không kể thời gian phát đề)",
    "exam_code": "1001",
    "atomic_masses": "C=12, H=1, O=16, Cu=64, Ag=108, Cl=35.5, Na=23, N=14",

    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1: Methyl acetate có công thức cấu tạo nào sau đây?",
            "options": [
                "A. CH~3~COOC~3~H~7~.",
                "B. HCOOC~2~H~5~.",
                "*C. CH~3~COOCH~3~.*",
                "D. CH~3~COOC~2~H~5~."
            ],
            "solution": "Methyl acetate được tạo từ acetic acid (CH~3~COOH) và methyl alcohol (CH~3~OH), có công thức cấu tạo là CH~3~COOCH~3~."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Amine nào sau đây là amine bậc hai?",
            "options": [
                "A. (CH~3~)~3~N.",
                "B. CH~5~N.",
                "*C. (CH~3~)~2~NH.*",
                "D. C~2~H~7~NH~2~."
            ],
            "solution": "Amine bậc hai là amine có nguyên tử nitrogen liên kết trực tiếp với hai gốc hydrocarbon: (CH~3~)~2~NH (dimethylamine). Các chất còn lại: (CH~3~)~3~N là amine bậc III; CH~5~N (tức CH~3~NH~2~) là amine bậc I."
        },
        {
            "id": 3,
            "prompt": "Câu 3: Khi nấu món canh từ thịt cua, tôm, tép có nhiều mảng thịt đóng rắn lại. Hiện tượng trên gây ra bởi tính chất nào sau đây?",
            "options": [
                "A. Sự thủy phân protein bởi nhiệt độ.",
                "B. Sự đông tụ protein bởi sự thay đổi pH.",
                "*C. Sự đông tụ protein bởi nhiệt độ.*",
                "D. Kết tủa carbonate của các chất khoáng có trong vỏ."
            ],
            "solution": "Dưới tác dụng của nhiệt độ khi đun sôi, cấu trúc không gian bậc cao của các phân tử protein hòa tan trong nước thịt cua, tôm tép bị biến tính và vón kết lại thành các mảng kết tủa đóng rắn nổi lên (hiện tượng đông tụ protein bởi nhiệt độ)."
        },
        {
            "id": 4,
            "prompt": "Câu 4: Cho sức điện động chuẩn của các pin điện hoá sau:\n"
                      "E°~pin(T-X)~ = 2,46 V; E°~pin (T – Y)~ = 2,00 V (với X, Y, T là ba kim loại). Dãy các kim loại được sắp xếp theo chiều tăng dần tính khử là:",
            "options": [
                "*A. X < Y < T*",
                "B. Y < X < T",
                "C. T < Y < X",
                "D. T < X < Y"
            ],
            "solution": "Trong pin điện hóa T-X và T-Y, kim loại T đóng vai trò là anode (cực âm) có tính khử mạnh nhất. Ta có công thức sức điện động chuẩn:\n"
                       "E°~pin(T-X)~ = E°(X^n+^/X) - E°(T^m+^/T) = 2,46 V  =>  E°(X^n+^/X) = E°(T^m+^/T) + 2,46 V\n"
                       "E°~pin(T-Y)~ = E°(Y^k+^/Y) - E°(T^m+^/T) = 2,00 V  =>  E°(Y^k+^/Y) = E°(T^m+^/T) + 2,00 V\n"
                       "=> E°(X^n+^/X) > E°(Y^k+^/Y) > E°(T^m+^/T).\n"
                       "Thế điện cực chuẩn E° càng lớn thì tính oxi hóa của cation càng mạnh và tính khử của kim loại càng yếu. Do đó chiều tăng dần tính khử của các kim loại là: X < Y < T."
        },
        {
            "id": 5,
            "prompt": "Câu 5: Chất nào sau đây thuộc loại monosaccharide?",
            "options": [
                "A. Cellulose.",
                "B. Saccharose.",
                "C. Tinh bột.",
                "*D. Glucose.*"
            ],
            "solution": "Glucose (và fructose) là các monosaccharide (đường đơn không bị thủy phân). Saccharose là disaccharide; tinh bột và cellulose là polysaccharide."
        },
        {
            "id": 6,
            "prompt": "Câu 6: Tinh bột thuộc loại polysaccharide, có nhiều trong các loại ngũ cốc như hạt lúa, hạt ngô, củ khoai, củ sắn, .... Công thức phân tử của tinh bột là:",
            "options": [
                "A. C~12~H~22~O~11~.",
                "B. C~6~H~12~O~6~.",
                "*C. (C~6~H~10~O~5~)~n~.*",
                "D. C~2~H~4~O~2~."
            ],
            "solution": "Tinh bột là polysaccharide được tạo thành từ nhiều mắt xích α-glucose, có công thức phân tử chung là (C~6~H~10~O~5~)~n~."
        },
        {
            "id": 7,
            "prompt": "Câu 7: Dạng ion chủ yếu nào của amino acid tồn tại trong môi trường acid mạnh (pH khoảng 1-2)?",
            "image": "cau7_amino_acid_forms.png",
            "image_width": Cm(11.0),
            "options": [
                "A. Dạng anion H~2~N-CH~2~-COO^-^.",
                "B. Dạng ion lưỡng cực ^+^H~3~N-CH~2~-COO^-^.",
                "C. Dạng anion CH~3~-CH(NH~2~)-COO^-^.",
                "*D. Dạng cation CH~3~-CH(^+^NH~3~)-COOH.*"
            ],
            "solution": "Trong dung dịch acid mạnh (pH khoảng 1 - 2), nồng độ H^+^ rất cao. Nhóm carboxylate -COO^-^ nhận proton H^+^ trở thành nhóm -COOH dạng phân tử, còn nhóm amino -NH~2~ nhận H^+^ tạo thành cation -NH~3~^+^. Do đó, trong môi trường acid mạnh, amino acid tồn tại chủ yếu ở dạng CATION mang điện tích dương: ^+^H~3~N-CH(R)-COOH (tương ứng với phương án D)."
        },
        {
            "id": 8,
            "prompt": "Câu 8: Sản phẩm thu được khi điện phân dung dịch KCl (điện cực trơ, có màng ngăn) gồm:",
            "options": [
                "A. K và Cl~2~.",
                "B. K, H~2~ và Cl~2~.",
                "*C. KOH, H~2~ và Cl~2~.*",
                "D. KOH, O~2~ và HCl."
            ],
            "solution": "Phương trình điện phân dung dịch KCl có màng ngăn xốp:\n2KCl + 2H~2~O --[đpdd, có màng ngăn]--> 2KOH + H~2~ + Cl~2~."
        },
        {
            "id": 9,
            "prompt": "Câu 9: Saccharose được cấu tạo từ:",
            "options": [
                "A. hai đơn vị \u03b1-glucose qua liên kết \u03b1-1,4-glycoside.",
                "B. một đơn vị \u03b1-glucose và một đơn vị \u03b2-fructose qua liên kết \u03b1-1,4-glycoside.",
                "*C. một đơn vị \u03b1-glucose và một đơn vị \u03b2-fructose qua liên kết \u03b1-1,2-glycoside.*",
                "D. hai đơn vị \u03b2-fructose qua liên kết \u03b2-1,4-glycoside."
            ],
            "solution": "Phân tử saccharose được cấu tạo từ một gốc \u03b1-glucose và một gốc \u03b2-fructose liên kết với nhau qua nguyên tử oxygen giữa C1 của gốc \u03b1-glucose và C2 của gốc \u03b2-fructose (gọi là liên kết \u03b1-1,2-glycosidic)."
        },
        {
            "id": 10,
            "prompt": "Câu 10: Một pin điện hoá có điện cực Zn nhúng trong dung dịch ZnSO~4~ và điện cực Cu nhúng trong dung dịch CuSO~4~. Sau một thời gian pin đó phóng điện thì:",
            "image": "pin_galvani_de1001.png",
            "image_width": Cm(8.0),
            "options": [
                "A. khối lượng cả hai điện cực Zn và Cu đều tăng.",
                "B. khối lượng cả hai điện cực Zn và Cu đều giảm.",
                "C. khối lượng điện cực Zn tăng còn khối lượng điện cực Cu giảm.",
                "*D. khối lượng điện cực Zn giảm còn khối lượng điện cực Cu tăng.*"
            ],
            "solution": "Khi pin phóng điện:\n- Ở cực âm (anode Zn): Zn bị oxi hóa tan ra: Zn -> Zn^2+^ + 2e => khối lượng thanh Zn giảm.\n- Ở cực dương (cathode Cu): Cu^2+^ bị khử bám vào: Cu^2+^ + 2e -> Cu => khối lượng thanh Cu tăng."
        },
        {
            "id": 11,
            "prompt": "Câu 11: Kết quả phân tích nguyên tố của một \u03b1-amino acid X như sau: %C=46,60%; %H=8,74%; %N=13,59% (về khối lượng); còn lại là oxygen. Bằng phổ khối lượng (MS), xác định được phân tử khối của X bằng 103. Số công thức cấu tạo thỏa mãn của X là:",
            "options": [
                "A. 4.",
                "B. 3.",
                "*C. 2.*",
                "D. 1."
            ],
            "solution": "1. Xác định CTPT của X:\n- %O = 100% - (46,60% + 8,74% + 13,59%) = 31,07%.\n- Số nguyên tử C: n_C = (103 * 46,60%) / 12 = 4.\n- Số nguyên tử H: n_H = (103 * 8,74%) / 1 = 9.\n- Số nguyên tử N: n_N = (103 * 13,59%) / 14 = 1.\n- Số nguyên tử O: n_O = (103 * 31,07%) / 16 = 2.\n=> CTPT của X là C~4~H~9~NO~2~ (M = 103).\n2. Số công thức cấu tạo là \u03b1-amino acid (nhóm -NH~2~ đính ở C số 2 kề nhóm -COOH):\n(1) Mạch không nhánh: CH~3~-CH~2~-CH(NH~2~)-COOH (2-aminobutanoic acid)\n(2) Mạch phân nhánh: (CH~3~)~2~C(NH~2~)-COOH (2-amino-2-methylpropanoic acid)\n=> Có 2 công thức cấu tạo thỏa mãn."
        },
        {
            "id": 12,
            "prompt": "Câu 12: Linoleic acid có công thức khung phân tử như bên dưới thuộc loại acid béo nào sau đây?",
            "image": "linoleic_acid.png",
            "image_width": Cm(9.5),
            "options": [
                "A. Omega-3.",
                "*B. Omega-6.*",
                "C. Omega-9.",
                "D. Omega-12."
            ],
            "solution": "Đếm chỉ số carbon từ nguyên tử carbon ở nhóm methyl cuối mạch (-CH~3~, gọi là đầu \u03c9, đối diện nhóm carboxyl -COOH): Liên kết đôi C=C đầu tiên xuất hiện ở vị trí carbon số 6 tính từ đầu \u03c9. Do đó, linoleic acid thuộc loại acid béo Omega-6."
        },
        {
            "id": 13,
            "prompt": "Câu 13: Trong quá trình điện phân MgCl~2~ nóng chảy với các điện cực trơ, ở cathode xảy ra quá trình:",
            "options": [
                "A. khử ion Cl^-^.",
                "B. oxi hóa ion Cl^-^.",
                "*C. khử ion Mg^2+^.*",
                "D. oxi hóa ion Mg^2+^."
            ],
            "solution": "Ở cathode (cực âm) xảy ra quá trình khử cation kim loại: Mg^2+^ + 2e -> Mg."
        },
        {
            "id": 14,
            "prompt": "Câu 14: Polymer nào sau đây được tổng hợp bằng phản ứng trùng hợp?",
            "options": [
                "A. tinh bột.",
                "B. tơ visco.",
                "C. tơ nylon-6.",
                "*D. cao su buna.*"
            ],
            "solution": "Cao su buna được tổng hợp bằng phản ứng trùng hợp buta-1,3-diene: n CH~2~=CH-CH=CH~2~ -> [-CH~2~-CH=CH-CH~2~-]~n~. Tinh bột là polymer thiên nhiên; tơ visco là bán tổng hợp; tơ nylon-6 là sản phẩm của phản ứng trùng ngưng (hoặc trùng hợp mở vòng)."
        },
        {
            "id": 15,
            "prompt": "Câu 15: Số liên kết peptide trong phân tử peptide Gly-Ala-Val-Tyr-Phe là:",
            "options": [
                "A. 2.",
                "B. 1.",
                "C. 3.",
                "*D. 4.*"
            ],
            "solution": "Phân tử pentapeptide gồm 5 gốc α-amino acid liên kết với nhau bằng 4 liên kết peptide (-CO-NH-)."
        },
        {
            "id": 16,
            "prompt": "Câu 16: Trong các loại tơ: nylon-6,6; tơ visco; nylon-6; tơ capron; sợi bông; tơ cellulose acetate. Số loại tơ có nguồn gốc từ cellulose là:",
            "options": [
                "A. 5.",
                "*B. 3.*",
                "C. 2.",
                "D. 4."
            ],
            "solution": "Các loại tơ có nguồn gốc từ cellulose gồm: tơ visco, sợi bông, tơ cellulose acetate (3 loại). Các loại tơ còn lại (nylon-6,6; nylon-6; tơ capron) là các polyamide tổng hợp."
        }
    ],

    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1: Cho thế điện cực chuẩn của các cặp oxi hóa - khử như sau:\n"
                      "Cặp Fe^2+^/Fe có E° = -0,04V; Cặp Cu^2+^/Cu có E° = 0,34V; Cặp Fe^3+^/Fe^2+^ có E° = 0,77V; Cặp Ag^+^/Ag có E° = 0,80V.",
            "statements": [
                ("a) Các kim loại được xếp theo chiều tính khử tăng dần: Ag < Fe < Cu.", "SAI"),
                ("b) Cho Fe tác dụng với dung dịch AgNO~3~ dư thu được dung dịch Y chứa 2 muối là Fe(NO~3~)~2~ và AgNO~3~.", "SAI"),
                ("c) Sức điện động chuẩn của pin điện hóa Fe-Cu là: 0,38V.", "ĐÚNG"),
                ("d) Khi nhúng thanh đồng có khối lượng 6,4 gam vào 400 mL dung dịch AgNO~3~ 0,2M. Giả thiết toàn bộ lượng Ag giải phóng đều bám vào thanh đồng thì khối lượng thanh đồng sau khi phản ứng kết thúc là 12,48 gam.", "ĐÚNG")
            ],
            "solution": "- a) SAI: Căn cứ vào thế điện cực chuẩn, tính khử của Fe mạnh hơn Cu: Chiều tăng dần tính khử đúng là Ag < Cu < Fe.\n"
                       "- b) SAI: Vì dung dịch AgNO~3~ DƯ nên Fe bị oxi hóa lên số oxi hóa cao nhất (+3): Fe + 3Ag^+^ -> Fe^3+^ + 3Ag. Dung dịch thu được gồm Fe(NO~3~)~3~ và AgNO~3~ dư, không chứa Fe(NO~3~)~2~.\n"
                       "- c) ĐÚNG: Theo bảng số liệu đề bài cho: E°~pin(Fe-Cu)~ = E°(Cu^2+^/Cu) - E°(Fe^2+^/Fe) = 0,34 - (-0,04) = 0,38 V.\n"
                       "- d) ĐÚNG: Tính toán:\n"
                       "  + n(AgNO~3~) = 0,4 * 0,2 = 0,08 mol.\n"
                       "  + PTHH: Cu + 2AgNO~3~ -> Cu(NO~3~)~2~ + 2Ag\n"
                       "  + n(Cu tan) = 0,08 / 2 = 0,04 mol; n(Ag bám) = 0,08 mol.\n"
                       "  + Độ tăng khối lượng: Δm = m(Ag bám) - m(Cu tan) = (0,08 * 108) - (0,04 * 64) = 8,64 - 2,56 = 6,08 gam.\n"
                       "  + Khối lượng thanh đồng sau phản ứng: m_sau = 6,4 + 6,08 = 12,48 gam."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Tơ capron là loại tơ có tính dai, độ đàn hồi và độ bóng cao, ít thấm nước và được sử dụng để dệt vải may mặc. Tơ capron được tổng hợp từ caprolactam theo phương trình hóa học sau:",
            "image": "caprolactam_scheme.png",
            "image_width": Cm(10.0),
            "statements": [
                ("a) Phản ứng (1) thuộc loại phản ứng trùng ngưng.", "SAI"),
                ("b) Tơ capron là tơ tổng hợp.", "ĐÚNG"),
                ("c) Do tơ capron có các nhóm –CO–NH– nên tơ này bền với dung dịch kiềm mạnh.", "SAI"),
                ("d) Do tơ capron có các nhóm –CO–NH– nên tơ capron là loại polypeptide.", "SAI")
            ],
            "solution": "- a) SAI: Phản ứng mở vòng caprolactam tạo tơ capron không giải phóng phân tử nhỏ nên thuộc loại phản ứng TRÙNG HỢP (trùng hợp mở vòng).\n"
                       "- b) ĐÚNG: Tơ capron được sản xuất bằng con đường tổng hợp hóa học nên thuộc loại tơ tổng hợp.\n"
                       "- c) SAI: Liên kết amide -CO-NH- rất dễ bị thủy phân trong môi trường acid và kiềm mạnh, do đó tơ capron kém bền với dung dịch kiềm mạnh và acid mạnh.\n"
                       "- d) SAI: Polypeptide là các polymer tạo nên từ các mắt xích α-amino acid. Tơ capron tạo thành từ amino acid có 6C (ε-aminocaproic acid) nên thuộc loại polyamide, không phải polypeptide."
        }
    ],

    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1: Isoamyl acetate có mùi thơm đặc trưng của chuối chín nên còn được gọi là dầu chuối. Khi đun nóng hỗn hợp gồm 16,2 gam acetic acid và 26,4 gam isoamyl alcohol ((CH~3~)~2~CHCH~2~CH~2~OH) với xúc tác là H~2~SO~4~ đặc, đun nóng thu được 26 gam dầu chuối. Tính hiệu suất của phản ứng điều chế dầu chuối trên. (Kết quả làm tròn đến hàng phần mười).",
            "answer": "74,1%",
            "solution": "1. Tính số mol các chất tham gia:\n"
                       "   - n(CH~3~COOH) = 16,2 / 60 = 0,27 mol.\n"
                       "   - M(isoamyl alcohol) = 88 g/mol => n(isoamyl alcohol) = 26,4 / 88 = 0,30 mol.\n"
                       "   - Vì 0,27 < 0,30 nên hiệu suất phản ứng tính theo acetic acid.\n"
                       "2. PTHH: CH~3~COOH + (CH~3~)~2~CHCH~2~CH~2~OH <--> CH~3~COOCH~2~CH~2~CH(CH~3~)~2~ + H~2~O\n"
                       "3. Tính khối lượng ester lý thuyết:\n"
                       "   - M(isoamyl acetate) = 130 g/mol.\n"
                       "   - m_LT = 0,27 * 130 = 35,1 gam.\n"
                       "4. Hiệu suất của phản ứng:\n"
                       "   H = (26 / 35,1) * 100% ≈ 74,074% ≈ 74,1%."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Cho các chất: glucose, amylose, cellulose, maltose, fructose, saccharose. Có bao nhiêu chất tham gia phản ứng thuỷ phân trong môi trường acid?",
            "answer": "4",
            "solution": "Các carbohydrate tham gia phản ứng thủy phân trong môi trường acid là các disaccharide và polysaccharide:\n"
                       "- Disaccharide: maltose, saccharose (2 chất).\n"
                       "- Polysaccharide: amylose, cellulose (2 chất).\n"
                       "Glucose và fructose là monosaccharide không bị thủy phân.\n"
                       "=> Có 4 chất tham gia phản ứng thủy phân trong môi trường acid."
        },
        {
            "id": 3,
            "prompt": "Câu 3: Thủy phân một tripeptide X (mạch hở) thu được hỗn hợp gồm glycine, alanine (có tỉ lệ mol 2:1). Có bao nhiêu đồng phân cấu tạo của X thỏa mãn tính chất trên?",
            "answer": "3",
            "solution": "Tripeptide X tạo bởi 2 gốc glycine (Gly) và 1 gốc alanine (Ala).\n"
                       "Các đồng phân cấu tạo thỏa mãn của X là các hoán vị vị trí của 1 gốc Ala và 2 gốc Gly:\n"
                       "(1) Gly-Gly-Ala\n"
                       "(2) Gly-Ala-Gly\n"
                       "(3) Ala-Gly-Gly\n"
                       "=> Có 3 đồng phân cấu tạo."
        },
        {
            "id": 4,
            "prompt": "Câu 4: Poly(hexamethylene adipamide) còn được gọi là tơ nylon-6,6 được điều chế từ phản ứng trùng ngưng của adipic acid và hexamethylene diamine. Tơ nylon-6,6 có tính dai, mềm, ít thấm nước nên được dùng để dệt vải may mặc, bện dây cáp, dây dù, đan lưới. Một loại tơ nylon-6,6 có khối lượng phân tử là 37290 đvC. Có bao nhiêu đơn vị mắt xích có trong phân tử loại tơ sợi này?",
            "answer": "165",
            "solution": "Công thức một mắt xích nylon-6,6 là [-NH-(CH~2~)~6~-NH-CO-(CH~2~)~4~-CO-].\n"
                       "Khối lượng phân tử của một mắt xích:\n"
                       "M_mắt xích = 14*12 + 26*1 + 2*14 + 2*16 = 168 + 26 + 28 + 32 = 226 (đvC).\n"
                       "Số lượng mắt xích có trong phân tử tơ:\n"
                       "n = 37290 / 226 = 165 mắt xích."
        }
    ],

    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1: (1,0 điểm). Glycine là một amino acid có khối lượng phân tử nhỏ nhất trong số 20 amino acid thường được tìm thấy trong protein. Glycine tham gia vào nhiều chức năng của cơ thể, thúc đẩy quá trình phát triển cơ, cải thiện hệ tiêu hóa, làm chậm quá trình lão hóa,...\n"
                      "a. Hãy viết công thức cấu tạo và gọi tên hệ thống của Glycine.\n"
                      "b. Viết phương trình hóa học chứng minh tính lưỡng tính của Glycine.",
            "num_lines": 8,
            "solution": "a. Công thức cấu tạo và tên hệ thống của Glycine:\n"
                       "- Công thức cấu tạo: H~2~N-CH~2~-COOH\n"
                       "- Tên hệ thống (IUPAC): aminoethanoic acid (hoặc 2-aminoethanoic acid).\n"
                       "b. Phương trình hóa học chứng minh tính lưỡng tính của Glycine:\n"
                       "- Tính base (tác dụng với acid mạnh do nhóm -NH~2~ nhận proton):\n"
                       "  H~2~N-CH~2~-COOH + HCl -> ClH~3~N-CH~2~-COOH (glycine hydrochloride)\n"
                       "- Tính acid (tác dụng với base mạnh do nhóm -COOH nhường proton):\n"
                       "  H~2~N-CH~2~-COOH + NaOH -> H~2~N-CH~2~-COONa + H~2~O (sodium glycinate)"
        },
        {
            "id": 2,
            "prompt": "Câu 2: (1,0 điểm). Poly(vinyl chloride) hay còn được gọi là PVC được điều chế từ khí thiên nhiên theo sơ đồ:\n"
                      "CH~4~ -> C~2~H~2~ -> CH~2~=CHCl -> PVC\n"
                      "Để sản xuất 1 tấn PVC theo sơ đồ trên cần bao nhiêu m^3^ khí thiên nhiên (ở đkc), giả sử hiệu suất toàn bộ quá trình là 80%? Biết khí thiên nhiên chứa 95% methane về thể tích. (Kết quả làm tròn đến hàng phần mười).",
            "num_lines": 8,
            "solution": "Sơ đồ chuyển hóa theo bảo toàn nguyên tố C:\n"
                       "2 CH~4~  ------>  [-CH~2~-CH(Cl)-]~n~\n"
                       "2 mol               62,5 gam\n"
                       "1. Khối lượng PVC cần sản xuất: m(PVC) = 1 tấn = 10^6^ gam.\n"
                       "   Số mol mắt xích PVC: n(PVC) = 10^6^ / 62,5 = 16000 mol.\n"
                       "2. Số mol CH~4~ cần dùng theo lý thuyết: n(CH~4~) = 2 * 16000 = 32000 mol.\n"
                       "3. Thể tích CH~4~ ở điều kiện chuẩn (đkc, 1 mol = 24,79 L = 0,02479 m^3^):\n"
                       "   V_LT(CH~4~) = 32000 * 24,79 * 10^-3^ = 793,28 m^3^.\n"
                       "4. Do hiệu suất toàn bộ quá trình là 80% nên thể tích CH~4~ thực tế cần lấy là:\n"
                       "   V_TT(CH~4~) = 793,28 / 80% = 991,6 m^3^.\n"
                       "5. Thể tích khí thiên nhiên cần dùng (chứa 95% thể tích methane):\n"
                       "   V(khí thiên nhiên) = 991,6 / 95% ≈ 1043,789 m^3^ ≈ 1043,8 m^3^.\n"
                       "Đáp số: 1043,8 m^3^ khí thiên nhiên."
        },
        {
            "id": 3,
            "prompt": "Câu 3: (1,0 điểm).\n"
                      "a. Xà phòng (xà bông) là một chất giặt rửa các vết bẩn, vết dầu mỡ,... Hãy nêu thành phần chính của xà phòng?\n"
                      "b. Hỗn hợp X gồm tristearin, tripalmitin, stearic acid và palmitic acid (tỉ lệ mol n~tristearin~ : n~tripalmitin~ = 1 : 1 và n~palmitic acid~ : n~stearic acid~ = 2 : 1). Để sản xuất b bánh xà phòng (mỗi bánh đều nặng 75 gam và chứa 75% muối của các acid béo theo khối lượng), người ta cho 863,92 gam hỗn hợp X tác dụng vừa đủ với 306 gam dung dịch NaOH 40%, đun nóng. Biết hiệu suất toàn bộ quá trình sản xuất xà phòng là 88%. Tính giá trị của b (Kết quả làm tròn đến hàng đơn vị).",
            "image": "banh_xa_bong.png",
            "image_width": Cm(4.5),
            "num_lines": 9,
            "solution": "a. Thành phần chính của xà phòng:\n"
                       "- Thành phần chính của xà phòng là muối sodium hoặc potassium của các acid béo cao (như sodium stearate C~17~H~35~COONa, sodium palmitate C~15~H~31~COONa,...).\n"
                       "b. Tính giá trị của b:\n"
                       "1. Số mol NaOH: n(NaOH) = (306 * 40%) / 40 = 122,4 / 40 = 3,06 mol.\n"
                       "2. Đặt ẩn số mol các chất trong hỗn hợp X:\n"
                       "   - Gọi n(tristearin) = n(tripalmitin) = x mol.\n"
                       "   - Gọi n(stearic acid) = y mol => n(palmitic acid) = 2y mol.\n"
                       "3. Lập hệ phương trình:\n"
                       "   - Theo số mol NaOH phản ứng: 3x + 3x + y + 2y = 6x + 3y = 3,06  =>  2x + y = 1,02  (1)\n"
                       "   - Theo khối lượng hỗn hợp X:\n"
                       "     m_X = 890x + 806x + 284y + 256*(2y) = 1696x + 796y = 863,92  (2)\n"
                       "   - Giải hệ (1) và (2) ta được: x = 0,5 mol; y = 0,02 mol.\n"
                       "4. Tính khối lượng muối acid béo tạo thành bằng bảo toàn khối lượng:\n"
                       "   - n(glycerol) = 2x = 1,0 mol => m(glycerol) = 1,0 * 92 = 92 gam.\n"
                       "   - n(H~2~O) = 3y = 0,06 mol => m(H~2~O) = 0,06 * 18 = 1,08 gam.\n"
                       "   - m(muối acid béo) = m_X + m(NaOH) - m(glycerol) - m(H~2~O)\n"
                       "                      = 863,92 + 122,4 - 92 - 1,08 = 893,24 gam.\n"
                       "5. Tính số bánh xà phòng b:\n"
                       "   - Lượng muối acid béo thực tế thu được (H = 88%): m_muối_TT = 893,24 * 88% = 786,0512 gam.\n"
                       "   - Khối lượng muối có trong 1 bánh xà phòng: m_muối_1_bánh = 75 * 75% = 56,25 gam.\n"
                       "   - Số bánh xà phòng: b = 786,0512 / 56,25 ≈ 13,974 ≈ 14 (bánh).\n"
                       "Đáp số: b = 14 bánh."
        }
    ]
}


def generate_docx(is_teacher=False):
    """Tạo file docx theo chuẩn Thầy TRẦN VĂN THẠNH."""
    doc = Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line=DATA_EXAM_1001["school_line"],
        exam_title=DATA_EXAM_1001["exam_title"],
        subject_line=DATA_EXAM_1001["subject_line"],
        num_pages=DATA_EXAM_1001["num_pages"],
        duration=DATA_EXAM_1001["duration"],
        exam_code=DATA_EXAM_1001["exam_code"],
        atomic_masses=DATA_EXAM_1001["atomic_masses"],
        is_teacher=is_teacher
    )

    if is_teacher:
        add_compact_summary_answers(doc)

    # PHẦN I
    add_section_header(
        doc,
        title="A. PHẦN TRẮC NGHIỆM",
        subtitle="PHẦN 1: TRẮC NGHIỆM NHIỀU LỰA CHỌN (4 điểm – 16 câu). Thí sinh trả lời câu hỏi từ câu 1 đến câu 16, mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    for item in DATA_EXAM_1001["part1"]:
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
        title="PHẦN 2: TRẮC NGHIỆM ĐÚNG – SAI (2 điểm – 2 câu).",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    for item in DATA_EXAM_1001["part2"]:
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
        title="PHẦN 3: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (1 điểm - 4 câu).",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4."
    )

    for item in DATA_EXAM_1001["part3"]:
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
        title="B. PHẦN TỰ LUẬN (3 câu – 3 điểm)",
        subtitle="Thí sinh trình bày chi tiết lời giải từ câu 1 đến câu 3 vào giấy làm bài thi."
    )

    for item in DATA_EXAM_1001["part4"]:
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
                p_img.add_run().add_picture(img_path, width=item.get("image_width", Cm(4.5)))

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
    filename = f"De_Kiem_Tra_CK1_Hoa_12_NguyenChiThanh_Ma1001{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
