# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Học Kì I - Hóa học 12 - Trường THPT Cao Thắng (TP Huế)
Mã đề: 1203 - Năm học 2025 - 2026
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
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_de1203")
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

def add_quick_answer_table_1203(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Giáo viên ở đầu tài liệu Mã đề 1203.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 1203)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_p1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
    answers_p1 = ["Đ/A", "B", "B", "D", "D", "D", "A", "C", "A", "A", "D", "C", "C"]

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
        ["Câu 1", "Thế điện cực chuẩn & pin điện hóa", "SAI", "ĐÚNG", "ĐÚNG", "SAI"],
        ["Câu 2", "Poly(vinyl chloride) (PVC)", "SAI", "ĐÚNG", "ĐÚNG", "SAI"]
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
    tbl3_answers = ["Đáp án", "14,5", "4 (hoặc 3)", "60", "200"]

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
    render_rich_text(p_tl, "*4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN (3,0 ĐIỂM - MỖI CÂU 1,0 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl4 = doc.add_table(rows=4, cols=3)
    tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl4.autofit = False

    tbl4_data = [
        ("Câu", "Điểm", "Đáp số cốt lõi & Hướng dẫn chấm"),
        ("Câu 1", "1,0 đ", "• a) Phương trình: o-HOC~6~H~4~COOH + C~6~H~5~CH~2~OH ⇄ o-HOC~6~H~4~COOCH~2~C~6~H~5~ + H~2~O (H~2~SO~4~ đặc, t°).\n• b) m(ester) = 468 kg ⇒ n = 2,0526 kmol ⇒ V(alcohol) = 266,45 lít."),
        ("Câu 2", "1,0 đ", "• a) Phân loại: Tơ tự nhiên (tơ tằm, len); Tơ nhân tạo (tơ visco, tơ cellulose acetate); Tơ tổng hợp (tơ nylon-6,6, tơ capron).\n• b) Phương trình trùng ngưng tơ nylon-6,6 và tơ capron (hoặc trùng hợp mở vòng caprolactam)."),
        ("Câu 3", "1,0 đ", "• a) Tại anode (cực -): Cu → Cu^2+^ + 2e; Tại cathode (cực +): Ag^+^ + 1e → Ag.\n• b) E°(Cu^2+^/Cu) = 0,34 V ⇒ E°_pin(Cu-Ag) = 0,80 - 0,34 = 0,46 V.")
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


# Dữ liệu Đề thi mã 1203
DATA_EXAM_1203 = {
    "school_line": "SỞ GIÁO DỤC VÀ ĐÀO TẠO - TRƯỜNG THPT CAO THẮNG",
    "exam_title": "ĐỀ KIỂM TRA CUỐI KÌ 1 - NĂM HỌC 2025 - 2026",
    "subject_line": "MÔN: HÓA HỌC 12",
    "duration": "45 phút (không kể thời gian phát đề)",
    "num_pages": "03",
    "exam_code": "1203",
    "atomic_masses": "H = 1; C = 12; N = 14; Cl = 35,5; O = 16",

    # I. TRẮC NGHIỆM (7,0 ĐIỂM)
    # PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. (12 câu - 3,0 điểm)
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1. Trong pin điện hoá, quá trình khử",
            "options": [
                "A. không xảy ra ở cả cực âm và cực dương.",
                "*B. xảy ra ở cực dương.*",
                "C. xảy ra ở cực âm.",
                "D. xảy ra ở cực âm và cực dương."
            ],
            "solution": "Trong pin điện hóa (Galvanic cell):\n- Cực âm (anode): Xảy ra quá trình oxi hóa (Anode Oxidation).\n- Cực dương (cathode): Xảy ra quá trình khử (Cathode Reduction).\n*Đáp án đúng: B.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Glucose là chất dinh dưỡng và được dùng làm thuốc tăng lực cho người già, trẻ em và người ốm. Số nguyên tử oxygen trong phân tử glucose là",
            "options": [
                "A. 12.",
                "*B. 6.*",
                "C. 11.",
                "D. 5."
            ],
            "solution": "Công thức phân tử của glucose là C~6~H~12~O~6~. Do đó, trong một phân tử glucose có chứa 6 nguyên tử oxygen.\n*Đáp án đúng: B.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Quá trình cộng hợp liên tiếp nhiều phân tử nhỏ giống nhau hoặc tương tự nhau (monomer) tạo thành phân tử lớn (polymer) được gọi là phản ứng",
            "options": [
                "A. xà phòng hoá.",
                "B. thuỷ phân.",
                "C. trùng ngưng.",
                "*D. trùng hợp.*"
            ],
            "solution": "Theo định nghĩa SGK Hóa học 12: Phản ứng trùng hợp là quá trình cộng hợp liên tiếp nhiều phân tử nhỏ giống nhau hoặc tương tự nhau (monomer) tạo thành phân tử lớn (polymer).\n*Đáp án đúng: D.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Poly(methyl methacrylate) là chất rắn trong suốt, có khả năng cho ánh sáng truyền qua tốt nên được dùng để chế tạo thuỷ tinh hữu cơ. Poly(methyl methacrylate) được điều chế từ monomer nào sau đây?",
            "options": [
                "A. CH~3~COOC(CH~3~)=CH~2~.",
                "B. CH~3~COOCH=CH~2~.",
                "C. CH~2~=CHCOOCH~3~.",
                "*D. CH~2~=C(CH~3~)COOCH~3~.*"
            ],
            "solution": "Monomer tương ứng để trùng hợp tạo poly(methyl methacrylate) (plexiglas) là methyl methacrylate có công thức cấu tạo là CH~2~=C(CH~3~)COOCH~3~.\n*Đáp án đúng: D.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5. Polymer nào sau đây trong thành phần chỉ gồm hai nguyên tố C và H?",
            "options": [
                "A. Tơ nylon-6,6.",
                "B. Poly(methyl methacrylate).",
                "C. Poly(phenol formaldehyde).",
                "*D. Polybuta-1,3-diene.*"
            ],
            "solution": "- Tơ nylon-6,6: [-NH-(CH~2~)~6~-NH-CO-(CH~2~)~4~-CO-]~n~ chứa 4 nguyên tố: C, H, N, O.\n- Poly(methyl methacrylate): [-CH~2~-C(CH~3~)(COOCH~3~)-]~n~ chứa 3 nguyên tố: C, H, O.\n- Poly(phenol formaldehyde): chứa 3 nguyên tố: C, H, O.\n- Polybuta-1,3-diene: [-CH~2~-CH=CH-CH~2~-]~n~ là hydrocarbon polymer, chỉ gồm hai nguyên tố C và H.\n*Đáp án đúng: D.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6. Phát biểu nào sau đây là bản chất của sự lưu hoá cao su?",
            "options": [
                "*A. Tạo cầu nối disulfide giữa các mạch phân tử cao su.*",
                "B. Tạo loại cao su nhẹ hơn.",
                "C. Giảm giá thành cao su.",
                "D. Làm cao su dễ ăn khuôn."
            ],
            "solution": "Bản chất của sự lưu hóa cao su là quá trình tạo thành các cầu nối disulfide (-S-S-) nối các mạch cao su riêng rẽ thành mạng không gian 3 chiều vững chắc, giúp cao su có độ đàn hồi tốt hơn, chịu nhiệt, lâu mòn và khó hòa tan trong dung môi hữu cơ hơn cao su thô.\n*Đáp án đúng: A.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7. Trong các chất sau, chất nào là amin bậc II?",
            "options": [
                "A. CH~3~-CH(CH~3~)-NH~2~.",
                "B. C~6~H~5~NH~2~.",
                "*C. CH~3~-NH-CH~3~.*",
                "D. H~2~N-[CH~2~]~6~-NH~2~."
            ],
            "solution": "Bậc của amine bằng số nguyên tử H trong phân tử NH~3~ được thay thế bởi gốc hydrocarbon:\n- CH~3~-CH(CH~3~)-NH~2~, C~6~H~5~NH~2~, H~2~N-[CH~2~]~6~-NH~2~: đều chứa nhóm -NH~2~ nên là amine bậc I.\n- CH~3~-NH-CH~3~ (dimethylamine): có 2 gốc methyl gắn vào nguyên tử N (nhóm -NH-), là amine bậc II.\n*Đáp án đúng: C.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8. Ở điều kiện chuẩn, thiết lập được điện cực zinc (Zn) bằng cách nhúng thanh Zn vào dung dịch",
            "options": [
                "*A. ZnSO~4~ 1 M.*",
                "B. H~2~SO~4~ 1 M.",
                "C. NaCl 1 M.",
                "D. HCl 1 M."
            ],
            "solution": "Điện cực chuẩn zinc (Zn^2+^/Zn) được thiết lập bằng cách nhúng thanh kim loại Zn nguyên chất vào dung dịch chứa ion Zn^2+^ nồng độ 1 M (ví dụ dung dịch ZnSO~4~ 1 M) ở điều kiện chuẩn (nhiệt độ 25 °C hay 298 K, áp suất 1 bar).\n*Đáp án đúng: A.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9. Trùng hợp propylene thu được polymer có tên gọi là",
            "options": [
                "*A. polypropylene.*",
                "B. polyethylene.",
                "C. polystyrene.",
                "D. poly(vinyl chloride)."
            ],
            "solution": "Trùng hợp monomer propylene (propene: CH~2~=CH-CH~3~) thu được polypropylene (kí hiệu là PP):\n  n CH~2~=CH(CH~3~) --t°, p, xt--> [-CH~2~-CH(CH~3~)-]~n~.\n*Đáp án đúng: A.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10. Chất nào sau đây thuộc loại α-amino acid?",
            "options": [
                "A. H~2~NCH~2~CH~2~COOH.",
                "B. HOCH~2~COOH.",
                "C. H~2~NCH(CH~3~)NH~2~.",
                "*D. H~2~NCH(CH~3~)COOH.*"
            ],
            "solution": "α-amino acid là hợp chất hữu cơ tạp chức chứa đồng thời nhóm amino (-NH~2~) và nhóm carboxyl (-COOH), trong đó nhóm -NH~2~ đính trực tiếp vào nguyên tử carbon số 2 (carbon α ngay cạnh nhóm -COOH).\nChất H~2~NCH(CH~3~)COOH (alanine, 2-aminopropanoic acid) là một α-amino acid tiêu biểu.\n*Đáp án đúng: D.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11. Chất hữu cơ X có nhiều trong cây mía, hoa thốt nốt, củ cải đường,... Nồng độ của X trong nước mía có thể đạt tới 13%. Chất X là",
            "options": [
                "A. fructose.",
                "B. maltose.",
                "*C. saccharose.*",
                "D. glucose."
            ],
            "solution": "Saccharose (sucrose, đường kính mía) là chất hữu cơ có nhiều trong cây mía (chiếm khoảng 13% trong nước mía ép), củ cải đường và hoa thốt nốt.\n*Đáp án đúng: C.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12. Phát biểu nào sau đây không đúng?",
            "options": [
                "A. Tơ nylon-6,6 được điều chế bằng phản ứng trùng ngưng.",
                "B. Tơ nitron được điều chế bằng phản ứng trùng hợp.",
                "*C. Tơ tằm thuộc loại tơ bán tổng hợp.*",
                "D. Cao su lưu hóa có cấu trúc mạng không gian."
            ],
            "solution": "- A đúng: Tơ nylon-6,6 được điều chế bằng phản ứng trùng ngưng giữa hexamethylenediamine và adipic acid.\n- B đúng: Tơ nitron (olon) được điều chế bằng phản ứng trùng hợp vinyl cyanide (acrylonitrile).\n- C sai: Tơ tằm là tơ thiên nhiên (protein sợi từ tằm), không phải tơ bán tổng hợp (tơ bán tổng hợp là tơ visco, tơ cellulose acetate sản xuất từ cellulose).\n- D đúng: Cao su lưu hóa có các cầu nối disulfide liên kết các mạch tạo mạng không gian 3 chiều.\n*Đáp án đúng: C.*"
        }
    ],

    # PHẦN II. Câu trắc nghiệm đúng sai. (2 câu - 2,0 điểm)
    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1. Cho các cặp oxi hóa - khử và thế điện cực chuẩn sau:",
            "table_data": [
                ["Cặp oxi hóa - khử", "Zn^2+^/Zn", "Cu^2+^/Cu", "Fe^2+^/Fe", "Pb^2+^/Pb", "Fe^3+^/Fe^2+^"],
                ["E° (V)", "- 0,76", "0,34", "- 0,44", "- 0,13", "0,77"]
            ],
            "statements": [
                ("a) Cho Zn dư tác dụng với dung dịch FeCl~3~ đến khi phản ứng hoàn toàn thu được dung dịch gồm FeCl~2~ và ZnCl~2~.", "S"),
                ("b) Tính oxi hoá của Fe^3+^ > Cu^2+^ > Pb^2+^ > Fe^2+^ > Zn^2+^.", "Đ"),
                ("c) Sức điện động chuẩn của pin Fe-Cu có giá trị là 0,78 V.", "Đ"),
                ("d) Cho kim loại Fe tác dụng với dung dịch Cu(NO~3~)~2~ dư, đến khi phản ứng xảy ra hoàn toàn, thu được dung dịch gồm Cu(NO~3~)~2~ và Fe(NO~3~)~3~.", "S")
            ],
            "solution": "• a) Sai. Do kim loại Zn dùng dư nên quá trình khử ion Fe^3+^ diễn ra qua 2 giai đoạn:\n    Giai đoạn 1: Zn + 2Fe^3+^ → Zn^2+^ + 2Fe^2+^\n    Giai đoạn 2: Vì Zn dư nên Zn tiếp tục khử Fe^2+^: Zn + Fe^2+^ → Zn^2+^ + Fe↓\n  Đến khi phản ứng kết thúc hoàn toàn, ion Fe^2+^ bị khử hết thành kim loại Fe. Dung dịch thu được chỉ chứa muối ZnCl~2~, không còn FeCl~2~.\n• b) Đúng. Dựa vào dãy thế điện cực chuẩn: E° càng lớn thì dạng oxi hóa có tính oxi hóa càng mạnh:\n  E°(Fe^3+^/Fe^2+^) = 0,77 V > E°(Cu^2+^/Cu) = 0,34 V > E°(Pb^2+^/Pb) = -0,13 V > E°(Fe^2+^/Fe) = -0,44 V > E°(Zn^2+^/Zn) = -0,76 V.\n  ⇒ Thứ tự giảm dần tính oxi hóa: Fe^3+^ > Cu^2+^ > Pb^2+^ > Fe^2+^ > Zn^2+^.\n• c) Đúng. Trong pin điện hóa Fe-Cu, cặp có thế điện cực chuẩn nhỏ hơn là cực âm (anode: Fe^2+^/Fe), cặp có thế lớn hơn là cực dương (cathode: Cu^2+^/Cu).\n  Sức điện động chuẩn của pin là:\n  E°_pin = E°_cathode - E°_anode = E°(Cu^2+^/Cu) - E°(Fe^2+^/Fe) = 0,34 - (-0,44) = 0,78 V.\n• d) Sai. Phản ứng xảy ra: Fe + Cu(NO~3~)~2~ → Fe(NO~3~)~2~ + Cu↓ (do E°(Fe^2+^/Fe) = -0,44 V < E°(Cu^2+^/Cu) = 0,34 V). Vì E°(Fe^3+^/Fe^2+^) = 0,77 V > E°(Cu^2+^/Cu) = 0,34 V nên ion Cu^2+^ không thể oxi hóa tiếp Fe^2+^ lên Fe^3+^. Do Cu(NO~3~)~2~ dư nên dung dịch thu được chứa Cu(NO~3~)~2~ dư và Fe(NO~3~)~2~ (không tạo Fe(NO~3~)~3~)."
        },
        {
            "id": 2,
            "prompt": "Câu 2. Poly(vinyl chloride) (PVC) là chất dẻo có tính cách điện tốt, bền với acid, dùng phổ biến để sản xuất vật liệu cách điện, ống dẫn nước, áo mưa.",
            "image": "pvc_logo.png",
            "image_width": Cm(2.2),
            "statements": [
                ("a) Poly(vinyl chloride) được điều chế từ vinyl chloride bằng phản ứng trùng ngưng.", "S"),
                ("b) Cho sơ đồ chuyển hóa: CH~4~ → C~2~H~2~ → C~2~H~3~Cl → PVC. Để tổng hợp 250 kg PVC theo sơ đồ trên thì cần 396,64 m^3^ khí CH~4~ (hiệu suất của cả quá trình là 50%).", "Đ"),
                ("c) Công thức của vinyl chloride là CH~2~=CHCl.", "Đ"),
                ("d) X kém bền với acid.", "S")
            ],
            "solution": "• a) Sai. Poly(vinyl chloride) được điều chế bằng phản ứng TRÙNG HỢP vinyl chloride (chứa liên kết đôi C=C), không phải phản ứng trùng ngưng:\n  n CH~2~=CH-Cl --t°, p, xt--> [-CH~2~-CH(Cl)-]~n~.\n• b) Đúng. Sơ đồ bảo toàn nguyên tố C và tính toán:\n  2 CH~4~ → C~2~H~2~ → C~2~H~3~Cl → 1/n (-C~2~H~3~Cl-)~n~\n  Khối lượng mol của 1 mắt xích PVC là M(C~2~H~3~Cl) = 2 × 12 + 3 + 35,5 = 62,5 g/mol.\n  Số mol mắt xích PVC cần tổng hợp: n_PVC = 250 / 62,5 = 4 kmol.\n  Theo lí thuyết: n_CH~4~(lt) = 2 × n_PVC = 2 × 4 = 8 kmol.\n  Vì hiệu suất toàn bộ quá trình là H = 50% nên số mol CH~4~ thực tế cần lấy là:\n  n_CH~4~(tt) = 8 / 0,5 = 16 kmol.\n  Thể tích khí CH~4~ ở điều kiện chuẩn (25 °C, 1 bar, 24,79 L/mol): \n  V_CH~4~ = 16 × 24,79 = 396,64 m^3^ (hoàn toàn chính xác!).\n• c) Đúng. Vinyl chloride là dẫn xuất monochloro của ethylene, có công thức cấu tạo là CH~2~=CH-Cl hay CH~2~=CHCl.\n• d) Sai. Theo thông tin đề bài đã nêu rõ: \"Poly(vinyl chloride) (PVC) là chất dẻo có tính cách điện tốt, bền với acid...\", do đó nhận định X kém bền với acid là sai sự thật."
        }
    ],

    # PHẦN III. Câu trắc nghiệm trả lời ngắn. (4 câu - 2,0 điểm)
    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1. Caprolactam được tổng hợp từ cuối thế kỉ XIX. Hiện nay, nhu cầu sản xuất caprolactam trên thế giới khoảng 10 triệu tấn/năm. 90% trong đó dùng để tổng hợp tơ capron. Trong công nghiệp, caprolactam được điều chế theo sơ đồ sau:",
            "image": "so_do_caprolactam.png",
            "image_width": Cm(14.0),
            "subprompt": "Để sản xuất 10 triệu tấn caprolactam, cần sử dụng bao nhiêu triệu tấn cyclohexanone? Giả sử hiệu suất trung bình của cả quá trình trên là 60%. (Kết quả làm tròn đến hàng phần mười).",
            "answer": "14,5",
            "solution": "Xác định công thức phân tử và khối lượng mol của các chất trong sơ đồ:\n- Cyclohexanone: C~6~H~10~O ⇒ M = 6 × 12 + 10 × 1 + 16 = 98 g/mol.\n- Caprolactam: C~6~H~11~NO ⇒ M = 6 × 12 + 11 × 1 + 14 + 16 = 113 g/mol.\nSơ đồ phản ứng bảo toàn khung carbon 6 cạnh:\n  C~6~H~10~O (cyclohexanone) → C~6~H~11~NO (caprolactam)\n       98 tấn                               113 tấn\n       m_lt                                10 triệu tấn\nKhối lượng cyclohexanone theo lí thuyết:\n  m_lt = (10 × 98) / 113 ≈ 8,672566 triệu tấn.\nVì hiệu suất trung bình của cả quá trình là H = 60% = 0,6 nên khối lượng cyclohexanone thực tế cần dùng là:\n  m_tt = m_lt / 0,6 = (10 × 98) / (113 × 0,6) = 980 / 67,8 ≈ 14,45428 triệu tấn.\nLàm tròn đến hàng phần mười (1 chữ số thập phân): 14,5.\n*Đáp án: 14,5*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Cho các nguyên liệu: dầu mỏ, acid béo, dầu thực vật, sodium hydroxide, sulfuric acid, mỡ động vật. Có bao nhiêu chất sử dụng được để điều chế xà phòng trực tiếp bằng phản ứng xà phòng hóa?",
            "answer": "4 (hoặc 3)",
            "solution": "Phân tích bản chất của phản ứng xà phòng hóa:\n- Phản ứng xà phòng hóa là phản ứng thủy phân chất béo trong dung dịch kiềm (hoặc phản ứng trung hòa acid béo bằng dung dịch kiềm) để trực tiếp tạo ra muối của acid béo (xà phòng).\n- Xét từng nguyên liệu trong danh sách:\n  1. Dầu thực vật: Chứa các triglyceride (chất béo) ⇒ Tham gia trực tiếp phản ứng xà phòng hóa với kiềm tạo xà phòng và glycerol.\n  2. Mỡ động vật: Chứa các triglyceride (chất béo) ⇒ Tham gia trực tiếp phản ứng xà phòng hóa với kiềm tạo xà phòng và glycerol.\n  3. Acid béo: Tác dụng trực tiếp với kiềm (NaOH) theo phản ứng xà phòng hóa (trung hòa) tạo muối acid béo (xà phòng).\n  4. Sodium hydroxide (NaOH): Là hóa chất kiềm bắt buộc tham gia trực tiếp vào phản ứng xà phòng hóa.\n  - Dầu mỏ: Dùng để tổng hợp chất giặt rửa tổng hợp, không điều chế xà phòng bằng phản ứng xà phòng hóa.\n  - Sulfuric acid (H~2~SO~4~): Là xúc tác cho phản ứng thủy phân trong môi trường acid (không phải xà phòng hóa).\n⇒ Đánh giá đáp số:\n• Nếu tính tất cả các chất tham gia trực tiếp phản ứng xà phòng hóa: có 4 chất (dầu thực vật, mỡ động vật, acid béo, sodium hydroxide).\n• Nếu câu hỏi nhấn mạnh vào các nguyên liệu hữu cơ cung cấp gốc acid béo: có 3 chất (dầu thực vật, mỡ động vật, acid béo).\n*Đáp án chuẩn: 4 (chấp nhận 3 theo quy ước đề)*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Cho sơ đồ chuyển hóa sau (mỗi mũi tên là một phản ứng):\nTinh bột → X → Y → Z → methyl acetate. Xác định phân tử khối của Z.",
            "answer": "60",
            "solution": "Lập luận xác định các chất trong chuỗi chuyển hóa:\n- Tinh bột (-C~6~H~10~O~5~-)~n~ thủy phân hoàn toàn tạo monosaccharide X là Glucose (C~6~H~12~O~6~):\n  (-C~6~H~10~O~5~-)~n~ + n H~2~O --enzyme/t°--> n C~6~H~12~O~6~\n- Lên men rượu glucose thu được ethanol Y (C~2~H~5~OH):\n  C~6~H~12~O~6~ --men rượu, 30-32°C--> 2 C~2~H~5~OH + 2 CO~2~↑\n- Lên men giấm ethanol thu được acetic acid Z (CH~3~COOH):\n  C~2~H~5~OH + O~2~ --men giấm--> CH~3~COOH + H~2~O\n- Acetic acid Z thực hiện phản ứng este hóa với methanol (CH~3~OH) tạo methyl acetate:\n  CH~3~COOH + CH~3~OH ⇄ CH~3~COOCH~3~ + H~2~O (H~2~SO~4~ đặc, t°)\n⇒ Chất Z là Acetic acid (CH~3~COOH).\nPhân tử khối của Z: M(CH~3~COOH) = 12 × 2 + 4 × 1 + 16 × 2 = 60 amu.\n*Đáp án: 60*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Khi thuỷ phân hoàn toàn 500 gam protein X thì thu được 178 gam alanine. Nếu phân tử khối của X là 50000 thì số mắt xích alanine trong phân tử X là bao nhiêu?",
            "answer": "200",
            "solution": "Công thức phân tử alanine là C~3~H~7~NO~2~ ⇒ Phân tử khối M(alanine) = 89 amu.\nSố mol alanine thu được khi thủy phân hoàn toàn:\n  n(Ala) = 178 / 89 = 2,0 mol.\nSố mol phân tử protein X đem thủy phân:\n  n(X) = 500 / 50 000 = 0,01 mol.\nGọi k là số mắt xích alanine có trong một phân tử protein X:\n  Protein X + ... → k Alanine + ...\n  Theo phương trình: n(Ala) = k × n(X)\n  ⇒ k = n(Ala) / n(X) = 2,0 / 0,01 = 200.\nVậy trong một phân tử protein X có chứa 200 mắt xích alanine.\n*Đáp án: 200*"
        }
    ],

    # II. PHẦN TỰ LUẬN (3,0 ĐIỂM)
    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1. (1,0 điểm)\nBenzyl salicylate được sử dụng làm chất cố định hương nước hoa, các sản phẩm khử mùi, kem chống nắng... Chất này được điều chế bằng phản ứng giữa salicylic acid (o-hydroxy benzoic acid) và benzyl alcohol.\na. Viết phương trình hoá học của các phản ứng điều chế ester trên.\nb. Để sản xuất 400 lít benzyl salicylate nguyên chất (D = 1,17 g/mL) thì cần tối thiểu bao nhiêu lít benzyl alcohol nguyên chất (D = 1,04 g/mL), biết hiệu suất phản ứng là 80%.",
            "image": "benzyl_salicylate.png",
            "image_width": Cm(15.0),
            "num_lines": 10,
            "solution": "a. Phương trình hóa học của phản ứng este hóa điều chế benzyl salicylate:\nNhóm -OH gắn trực tiếp vào vòng benzen (OH phenol) không tham gia phản ứng este hóa với alcohol thông thường. Chỉ có nhóm carboxyl (-COOH) của salicylic acid phản ứng este hóa với benzyl alcohol (C~6~H~5~CH~2~OH):\no-HO-C~6~H~4~-COOH + C~6~H~5~CH~2~OH ⇄ o-HO-C~6~H~4~-COOCH~2~C~6~H~5~ + H~2~O   (xúc tác H~2~SO~4~ đặc, đun nóng)\n\nb. Tính thể tích benzyl alcohol cần dùng:\n1. Khối lượng benzyl salicylate cần sản xuất:\n   m(ester) = V × D = 400 lít × 1,17 kg/lít = 468 kg.\n2. Phân tử khối của benzyl salicylate (C~14~H~12~O~3~):\n   M(ester) = 14 × 12 + 12 × 1 + 3 × 16 = 228 g/mol.\n   ⇒ Số mol ester: n(ester) = 468 / 228 = 39/19 ≈ 2,05263 kmol.\n3. Theo phương trình phản ứng: tỉ lệ mol n(benzyl alcohol) : n(ester) = 1 : 1.\n   Do hiệu suất phản ứng là H = 80% = 0,8 nên số mol benzyl alcohol thực tế cần dùng là:\n   n(alcohol) = n(ester) / 0,8 = (39/19) / 0,8 = 39 / 15,2 ≈ 2,56579 kmol.\n4. Phân tử khối của benzyl alcohol (C~7~H~8~O = C~6~H~5~CH~2~OH):\n   M(alcohol) = 7 × 12 + 8 × 1 + 16 = 108 g/mol.\n   ⇒ Khối lượng benzyl alcohol thực tế cần dùng:\n   m(alcohol) = 2,56579 × 108 ≈ 277,1053 kg.\n5. Thể tích benzyl alcohol nguyên chất (D = 1,04 g/mL = 1,04 kg/lít):\n   V(alcohol) = m(alcohol) / D = 277,1053 / 1,04 ≈ 266,45 lít.\n*Kết luận: Cần tối thiểu khoảng 266,45 lít benzyl alcohol nguyên chất.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. (1,0 điểm)\nTrong số các loại tơ sau: tơ tằm, tơ visco, tơ nylon-6,6, len, tơ cellulose acetate, tơ capron.\na. Tơ nào là tơ tự nhiên, tơ nhân tạo và tơ tổng hợp?\nb. Viết phương trình hoá học của phản ứng điều chế các loại tơ tổng hợp ở trên từ các monomer tương ứng (ghi rõ điều kiện nếu có).",
            "num_lines": 10,
            "solution": "a. Phân loại các loại tơ:\n- Tơ tự nhiên (có nguồn gốc sẵn từ thiên nhiên): tơ tằm, len (đều có bản chất là protein tự nhiên).\n- Tơ nhân tạo (tơ bán tổng hợp, chế hóa từ polymer tự nhiên cellulose): tơ visco, tơ cellulose acetate.\n- Tơ tổng hợp (chế tạo hoàn toàn bằng con đường tổng hợp hóa học): tơ nylon-6,6, tơ capron.\n\nb. Phương trình hóa học điều chế các tơ tổng hợp:\n1. Điều chế tơ nylon-6,6 bằng phản ứng trùng ngưng hexamethylenediamine và adipic acid (hexanedioic acid):\n   n H~2~N-[CH~2~]~6~-NH~2~ + n HOOC-[CH~2~]~4~-COOH --t°, p, xt--> [-NH-[CH~2~]~6~-NH-CO-[CH~2~]~4~-CO-]~n~ + 2n H~2~O\n2. Điều chế tơ capron (nylon-6):\n   • Cách 1: Trùng ngưng ε-aminocaproic acid (6-aminohexanoic acid):\n     n H~2~N-[CH~2~]~5~-COOH --t°, xt--> [-NH-[CH~2~]~5~-CO-]~n~ + n H~2~O\n   • Cách 2: Trùng hợp mở vòng monomer caprolactam:\n     n (Caprolactam) --t°, xt--> [-NH-[CH~2~]~5~-CO-]~n~"
        },
        {
            "id": 3,
            "prompt": "Câu 3. (1,0 điểm)\nBiết phản ứng oxi hoá - khử xảy ra trong một pin điện hoá: Cu + 2Ag^+^ → Cu^2+^ + 2Ag.\na. Viết các quá trình xảy ra tại mỗi điện cực.\nb. Cho thế điện cực chuẩn của các cặp Zn^2+^/Zn và Ag^+^/Ag lần lượt là - 0,76 V và 0,8 V. Sức điện động chuẩn của pin điện hoá Zn-Cu là 1,10 V. Tính sức điện động chuẩn của pin điện hoá Cu-Ag.",
            "num_lines": 10,
            "solution": "a. Các quá trình xảy ra tại mỗi điện cực của pin điện hóa Cu - Ag:\n- Ở cực âm (anode): Xảy ra quá trình oxi hóa kim loại Cu:\n  Cu → Cu^2+^ + 2e\n- Ở cực dương (cathode): Xảy ra quá trình khử ion Ag^+^:\n  Ag^+^ + 1e → Ag   (hoặc 2Ag^+^ + 2e → 2Ag)\n\nb. Tính sức điện động chuẩn của pin điện hóa Cu-Ag:\n1. Dựa vào sức điện động chuẩn của pin điện hóa Zn - Cu:\n   Ta có: E°_pin(Zn-Cu) = E°(cathode) - E°(anode) = E°(Cu^2+^/Cu) - E°(Zn^2+^/Zn)\n   Thay số: 1,10 = E°(Cu^2+^/Cu) - (- 0,76)\n   ⇒ E°(Cu^2+^/Cu) = 1,10 - 0,76 = + 0,34 V.\n2. Sức điện động chuẩn của pin điện hóa Cu - Ag:\n   Trong pin Cu - Ag, Cu đóng vai trò là cực âm (anode, E° = +0,34 V) và Ag đóng vai trò là cực dương (cathode, E° = +0,80 V).\n   Áp dụng công thức tính sức điện động chuẩn:\n   E°_pin(Cu-Ag) = E°(cathode) - E°(anode) = E°(Ag^+^/Ag) - E°(Cu^2+^/Cu)\n   Thay số: E°_pin(Cu-Ag) = 0,80 - 0,34 = 0,46 V.\n*Kết luận: Sức điện động chuẩn của pin điện hoá Cu-Ag là 0,46 V.*"
        }
    ]
}

def generate_docx(is_teacher=False):
    doc = Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line=DATA_EXAM_1203["school_line"],
        exam_title=DATA_EXAM_1203["exam_title"],
        subject_line=DATA_EXAM_1203["subject_line"],
        num_pages=DATA_EXAM_1203["num_pages"],
        duration=DATA_EXAM_1203["duration"],
        exam_code=DATA_EXAM_1203["exam_code"],
        atomic_masses=DATA_EXAM_1203["atomic_masses"],
        is_teacher=is_teacher
    )

    # Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Bản Giáo Viên
    if is_teacher:
        add_quick_answer_table_1203(doc)

    # I. TRẮC NGHIỆM (7,0 ĐIỂM)
    add_section_header(
        doc,
        title="I. PHẦN TRẮC NGHIỆM (7,0 ĐIỂM)",
        subtitle="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    for item in DATA_EXAM_1203["part1"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)
        render_options_no_table(doc, item["options"], is_teacher=is_teacher)
        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN II. Câu trắc nghiệm đúng sai.
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    for item in DATA_EXAM_1203["part2"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        # Nếu có bảng số liệu (như câu 1 thế điện cực)
        if "table_data" in item:
            t_data = item["table_data"]
            tbl = doc.add_table(rows=len(t_data), cols=len(t_data[0]))
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl.autofit = False
            for r_idx, row_vals in enumerate(t_data):
                row = tbl.rows[r_idx]
                for c_idx, val in enumerate(row_vals):
                    cell = row.cells[c_idx]
                    cell.width = Cm(3.2) if c_idx == 0 else Cm(2.4)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    p_c = cell.paragraphs[0]
                    p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_c.paragraph_format.space_before = Pt(2)
                    p_c.paragraph_format.space_after = Pt(2)
                    if r_idx == 0:
                        set_cell_shading(cell, "E8F5E9")
                        render_rich_text(p_c, f"*{val}*", base_size=9.5, base_bold=True, base_color=COLOR_PRIMARY)
                    else:
                        set_cell_shading(cell, "F5F5F5" if c_idx == 0 else "FFFFFF")
                        render_rich_text(p_c, val, base_size=9.5, base_bold=(c_idx == 0))
                    set_cell_border(cell,
                                    top=("single", "4", "0", "8BC390"),
                                    bottom=("single", "4", "0", "8BC390"),
                                    left=("single", "4", "0", "8BC390"),
                                    right=("single", "4", "0", "8BC390"))
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.space_before = Pt(2)
            p_sp.paragraph_format.space_after = Pt(2)

        # Nếu có ảnh (như logo PVC)
        if "image" in item:
            img_path = os.path.join(DIR_IMAGES, item["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(3)
                p_img.add_run().add_picture(img_path, width=item.get("image_width", Cm(3.0)))

        for stmt, ans in item["statements"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt, ans=ans, is_teacher=is_teacher, base_size=10.5)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN III. Câu trắc nghiệm trả lời ngắn.
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm trả lời ngắn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Thí sinh điền kết quả vào ô trả lời tương ứng."
    )

    for item in DATA_EXAM_1203["part3"]:
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
                p_img.add_run().add_picture(img_path, width=item.get("image_width", Cm(12.0)))

        if "subprompt" in item:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(1)
            p_sub.paragraph_format.space_after = Pt(2)
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

    # II. TỰ LUẬN (3,0 ĐIỂM)
    add_section_header(
        doc,
        title="II. PHẦN TỰ LUẬN (3,0 ĐIỂM)",
        subtitle="Thí sinh trình bày chi tiết lời giải vào giấy làm bài thi."
    )

    for item in DATA_EXAM_1203["part4"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        if is_teacher:
            if "image" in item:
                img_path = os.path.join(DIR_IMAGES, item["image"])
                if os.path.exists(img_path):
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.paragraph_format.space_before = Pt(2)
                    p_img.paragraph_format.space_after = Pt(3)
                    p_img.add_run().add_picture(img_path, width=item.get("image_width", Cm(13.0)))
            if "solution" in item:
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
    filename = f"De_Kiem_Tra_CK1_Hoa_12_CaoThang_Ma1203{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
