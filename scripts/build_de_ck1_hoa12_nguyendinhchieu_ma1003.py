# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Học Kì I - Hóa học 12 - Trường THPT Nguyễn Đình Chiểu (TP Huế)
Mã đề: 1003 - Năm học 2025 - 2026
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
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_de1003")
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

def add_quick_answer_table_1003(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Giáo viên ở đầu tài liệu Mã đề 1003.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 1003)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_p1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
    answers_p1 = ["Đ/A", "C", "A", "C", "B", "C", "B", "A", "C", "A", "C", "B", "B"]

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
        ["Câu 1", "Sơ đồ chuyển hóa X (C~4~H~8~O~2~)", "SAI", "SAI", "ĐÚNG", "ĐÚNG"],
        ["Câu 2", "Tính chất amino acid và amine", "ĐÚNG", "SAI", "ĐÚNG", "ĐÚNG"]
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
    tbl3_answers = ["Đáp án", "46", "6", "2", "4,5"]

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
        ("Câu 1", "1,0 đ", "• CTPT của X: C~3~H~7~NO~2~ (M = 89).\n• M(muối) = 4,85 / 0,05 = 97 g/mol (H~2~N-CH~2~-COONa) ⇒ CTCT: H~2~N-CH~2~-COOCH~3~ (methyl glycinate)."),
        ("Câu 2", "1,0 đ", "• n(CuSO~4~ pư) = n(Zn pư) = 0,01 / (65 - 64) = 0,01 mol.\n• Cu(OH)~2~ kết tủa = 0,49 / 98 = 0,005 mol ⇒ n(CuSO~4~ ban đầu) = 0,01 + 0,005 = 0,015 mol ⇒ C~M~ = 0,15 M."),
        ("Câu 3", "1,0 đ", "• V(ethanol) = 7 L = 7 000 mL ⇒ m = 5 523 g ⇒ n = 120,065 mol.\n• n(glucose tt) = (120,065 / 2) / 0,8 = 75,041 mol ⇒ m = 75,041 × 180 ≈ 13 507,34 gam (≈ 13,51 kg).")
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


# Dữ liệu Đề thi mã 1003
DATA_EXAM_1003 = {
    "school_line": "SỞ GD&ĐT TP HUẾ - TRƯỜNG THPT NGUYỄN ĐÌNH CHIỂU",
    "exam_title": "ĐỀ KIỂM TRA CUỐI HỌC KÌ I - NĂM HỌC: 2025 - 2026",
    "subject_line": "MÔN: HÓA HỌC – Khối 12",
    "duration": "45 phút (không kể thời gian giao đề)",
    "num_pages": "03",
    "exam_code": "1003",
    "atomic_masses": "H = 1; Cu = 64; Zn = 65; O = 16; C = 12; Na = 23; Ag = 108; N = 14; S = 32; Cl = 35,5; Br = 80",

    # A. PHẦN TRẮC NGHIỆM: (7 điểm)
    # PHẦN I: (3,0 điểm) Câu trắc nghiệm nhiều phương án lựa chọn. (12 câu)
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1: Công thức cấu tạo dạng mạch hở của glucose là",
            "options": [
                "A. [C~6~H~7~O~2~(OH)~3~]~n~.",
                "B. CH~2~OH[CHOH]~3~COCH~2~OH.",
                "*C. CH~2~OH[CHOH]~4~CHO.*",
                "D. CH~2~OH[CHOH]~4~CH~2~OH."
            ],
            "solution": "Glucose là một hợp chất polyhydroxy aldehyde tạp chức, có công thức cấu tạo mạch hở dạng thẳng là CH~2~OH[CHOH]~4~CHO (gồm 5 nhóm -OH kề nhau và 1 nhóm -CHO ở đầu mạch).\n(B là fructose dạng mạch hở, D là sorbitol).\n*Đáp án đúng: C.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2: Ester X có công thức phân tử C~4~H~8~O~2~. Thủy phân X trong dung dịch NaOH dư, thu được sản phẩm gồm sodium propionate và alcohol Y. Công thức của Y là",
            "options": [
                "*A. CH~3~OH.*",
                "B. C~3~H~5~(OH)~3~.",
                "C. C~3~H~7~OH.",
                "D. C~2~H~5~OH."
            ],
            "solution": "Muối sodium propionate có công thức là C~2~H~5~COONa (chứa 3 nguyên tử carbon trong gốc acid).\nEster X có công thức phân tử C~4~H~8~O~2~ (chứa 4 nguyên tử carbon) nên công thức cấu tạo của ester X là C~2~H~5~COOCH~3~ (methyl propionate).\nPhương trình phản ứng thủy phân trong dung dịch NaOH:\n  C~2~H~5~COOCH~3~ + NaOH → C~2~H~5~COONa + CH~3~OH\nVậy alcohol Y là methyl alcohol (CH~3~OH).\n*Đáp án đúng: A.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3: Một học sinh thực hiện ba thí nghiệm ở điều kiện chuẩn và quan sát được các hiện tượng sau:\n(1) Copper (Cu) kim loại không phản ứng với dung dịch Pb(NO~3~)~2~ 1 M.\n(2) Lead (Pb) kim loại tan trong dung dịch AgNO~3~ 1 M và xuất hiện tinh thể Ag.\n(3) Silver (Ag) kim loại không phản ứng với dung dịch Cu(NO~3~)~2~ 1 M.\nTrật tự nào sau đây thể hiện đúng mức độ khử của ba kim loại?",
            "options": [
                "A. Cu > Pb > Ag.",
                "B. Cu > Ag > Pb.",
                "*C. Pb > Cu > Ag.*",
                "D. Pb > Ag > Cu."
            ],
            "solution": "- Từ thí nghiệm (1): Cu không khử được ion Pb^2+^ trong dung dịch ⇒ Tính khử của Cu yếu hơn Pb (Pb > Cu).\n- Từ thí nghiệm (2): Pb khử được ion Ag^+^ tạo kim loại Ag kết tinh (Pb + 2 Ag^+^ → Pb^2+^ + 2 Ag) ⇒ Tính khử của Pb mạnh hơn Ag (Pb > Ag).\n- Từ thí nghiệm (3): Ag không khử được ion Cu^2+^ trong dung dịch (ngược lại Cu khử được Ag^+^) ⇒ Tính khử của Cu mạnh hơn Ag (Cu > Ag).\nKết hợp lại, trật tự giảm dần mức độ khử của ba kim loại là: Pb > Cu > Ag.\n*Đáp án đúng: C.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4: Trong điều kiện thích hợp, glucose lên men tạo thành khí CO~2~ và chất X. Tên gọi của X là",
            "options": [
                "A. methyl alcohol.",
                "*B. ethyl alcohol.*",
                "C. lactic acid.",
                "D. acetic acid."
            ],
            "solution": "Dưới tác dụng của enzyme trong men rượu ở nhiệt độ khoảng 30 - 32 °C, glucose bị lên men tạo thành ethyl alcohol (ethanol) và khí carbon dioxide:\n  C~6~H~12~O~6~ --men rượu--> 2 C~2~H~5~OH + 2 CO~2~↑\nVậy chất X là ethyl alcohol (ethanol).\n*Đáp án đúng: B.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5: Công thức cấu tạo dưới đây là của carbohydrate nào?",
            "image": "saccharose_haworth.png",
            "image_width": Cm(8.5),
            "options": [
                "A. glucose.",
                "B. fructose.",
                "*C. saccharose.*",
                "D. maltose."
            ],
            "solution": "Cấu trúc Haworth trên gồm một gốc α-glucopyranose liên kết qua cầu oxy giữa C~1~ của gốc α-glucose với C~2~ của gốc β-fructofuranose (liên kết α-1,2-glycosidic). Đây chính là cấu trúc phân tử của saccharose (sucrose).\n*Đáp án đúng: C.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6: Ester có mùi đặc trưng giống mùi táo và có công thức phân tử C~5~H~10~O~2~. Thủy phân X trong dung dịch NaOH, thu được muối sodium ethanoate và alcohol Y. Oxi hóa không hoàn toàn Y thu được chất Z có phản ứng iodoform. Công thức cấu tạo của X là",
            "options": [
                "A. CH~3~CH~2~CH~2~COOCH~3~.",
                "*B. CH~3~COOCH(CH~3~)~2~.*",
                "C. CH~3~COOCH~2~CH~2~CH~3~.",
                "D. CH~3~CH~2~COOCH~2~CH~3~."
            ],
            "solution": "- Muối sodium ethanoate (sodium acetate) có công thức là CH~3~COONa.\n- Ester X có CTPT C~5~H~10~O~2~ tạo CH~3~COONa nên có dạng CH~3~COOR', với gốc R' có 3 nguyên tử carbon (-C~3~H~7~).\n- Alcohol Y là C~3~H~7~OH:\n  + Nếu Y là propan-1-ol (CH~3~CH~2~CH~2~OH): oxi hóa tạo propanal (CH~3~CH~2~CHO) không cho phản ứng iodoform.\n  + Nếu Y là propan-2-ol (CH~3~-CH(OH)-CH~3~): oxi hóa không hoàn toàn tạo acetone (CH~3~-CO-CH~3~), chứa nhóm methyl ketone CH~3~-C(=O)-, phản ứng với I~2~/NaOH tạo kết tủa vàng CHI~3~ (phản ứng iodoform).\n⇒ Gốc R' là isopropyl: -CH(CH~3~)~2~.\nCông thức cấu tạo của X là CH~3~COOCH(CH~3~)~2~ (isopropyl acetate).\n*Đáp án đúng: B.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7: Dãy chỉ gồm các amino acid có số nhóm amino (-NH~2~) bằng số nhóm carboxyl (-COOH)?",
            "options": [
                "*A. Gly, Val, Ala.*",
                "B. Val, Lys, Ala.",
                "C. Gly, Ala, Lys.",
                "D. Gly, Ala, Glu."
            ],
            "solution": "- Glycine (Gly), Alanine (Ala), Valine (Val): đều chứa 1 nhóm -NH~2~ và 1 nhóm -COOH (số nhóm amino bằng số nhóm carboxyl, làm quỳ tím không đổi màu).\n- Lysine (Lys): Chứa 2 nhóm -NH~2~ và 1 nhóm -COOH.\n- Glutamic acid (Glu): Chứa 1 nhóm -NH~2~ và 2 nhóm -COOH.\n*Đáp án đúng: A.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8: Chất nào sau đây khi đun nóng với dung dịch H~2~SO~4~ loãng thu được sản phẩm có aldehyde?",
            "options": [
                "A. CH~3~-COO-CH~2~-CH=CH~2~.",
                "B. CH~2~=CH-COO-CH~2~-CH~3~.",
                "*C. CH~3~-COO-CH=CH-CH~3~.*",
                "D. CH~3~-COO-C(CH~3~)=CH~2~."
            ],
            "solution": "Thủy phân ester trong dung dịch acid:\n- A: CH~3~COOCH~2~CH=CH~2~ + H~2~O ⇄ CH~3~COOH + CH~2~=CH-CH~2~OH (alcohol allyl).\n- B: CH~2~=CHCOOC~2~H~5~ + H~2~O ⇄ CH~2~=CHCOOH + C~2~H~5~OH (alcohol ethyl).\n- C: CH~3~COOCH=CH-CH~3~ + H~2~O ⇄ CH~3~COOH + [CH~3~-CH=CH-OH]. Hợp chất enol kém bền chuyển vị thành aldehyde propanal: CH~3~-CH~2~-CH=O (sinh ra aldehyde!).\n- D: CH~3~COOC(CH~3~)=CH~2~ + H~2~O ⇄ CH~3~COOH + CH~3~-CO-CH~3~ (ketone acetone).\n*Đáp án đúng: C.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9: Từ amino acid X và methyl alcohol điều chế được ester Y có công thức phân tử C~3~H~7~O~2~N. Công thức cấu tạo của amino acid X là",
            "options": [
                "*A. H~2~NCH~2~COOH.*",
                "B. CH~3~CH~2~COOH.",
                "C. H~2~NCH~2~COOCH~3~.",
                "D. CH~3~CH(NH~2~)COOH."
            ],
            "solution": "Ester Y được tạo thành từ amino acid X và methyl alcohol (CH~3~OH) nên có dạng H~2~N-R-COOCH~3~.\nPhân tử Y có công thức C~3~H~7~O~2~N (có 3 nguyên tử carbon), trừ đi 1 carbon của gốc methyl alcohol (-CH~3~) thì gốc amino acid còn lại 2 carbon.\n⇒ Công thức của amino acid X là H~2~N-CH~2~-COOH (glycine).\n*Đáp án đúng: A.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10: Chất nào dưới đây không phải là polymer?",
            "options": [
                "A. Tinh bột.",
                "B. Cellulose.",
                "*C. Lipid.*",
                "D. Protein."
            ],
            "solution": "- Tinh bột, cellulose là các polysaccharide thiên nhiên có phân tử khối rất lớn (polymer thiên nhiên).\n- Protein là polypeptide cao phân tử (polymer thiên nhiên).\n- Lipid (chất béo như triglyceride) có phân tử khối khoảng vài trăm đến dưới một nghìn amu, không có cấu trúc mắt xích lặp lại nên không phải là polymer.\n*Đáp án đúng: C.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11: Quá trình cộng hợp liên tiếp nhiều phân tử nhỏ giống nhau hoặc tương tự nhau (monomer) tạo thành phân tử lớn (polymer) được gọi là phản ứng",
            "options": [
                "A. trùng ngưng.",
                "*B. trùng hợp.*",
                "C. tách.",
                "D. thế."
            ],
            "solution": "Theo định nghĩa chuẩn SGK Hóa học 12: Phản ứng trùng hợp là quá trình kết hợp liên tiếp nhiều phân tử nhỏ giống nhau hoặc tương tự nhau tạo thành phân tử lớn (polymer).\n*Đáp án đúng: B.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12: Trong các nhận xét dưới đây, nhận xét nào đúng?",
            "options": [
                "A. Tất cả các chất có công thức C~n~(H~2~O)~m~ đều là carbohydrate.",
                "*B. Đa số các carbohydrate có công thức chung C~n~(H~2~O)~m~.*",
                "C. Phân tử các carbohydrate đều có ít nhất 6 nguyên tử carbon.",
                "D. Tất cả các carbohydrate đều có công thức chung C~n~(H~2~O)~m~."
            ],
            "solution": "- B đúng: Đa số các carbohydrate quen thuộc đều có công thức chung dạng C~n~(H~2~O)~m~ (như glucose, fructose: C~6~(H~2~O)~6~; saccharose: C~12~(H~2~O)~11~; tinh bột, cellulose: [C~6~(H~2~O)~5~]~n~).\n- A sai: Một số chất có dạng C~n~(H~2~O)~m~ nhưng không phải carbohydrate (ví dụ formaldehyde HCHO = CH~2~O, acetic acid CH~3~COOH = C~2~(H~2~O)~2~, lactic acid C~3~H~6~O~3~ = C~3~(H~2~O)~3~).\n- C sai: Có carbohydrate chỉ có 3 carbon (triose), 5 carbon (như ribose C~5~H~10~O~5~).\n- D sai: Một số carbohydrate không có dạng C~n~(H~2~O)~m~ (ví dụ rhamnose C~6~H~12~O~5~, deoxyribose C~5~H~10~O~4~).\n*Đáp án đúng: B.*"
        }
    ],

    # PHẦN II: (2,0 điểm) Câu trắc nghiệm đúng sai. (2 câu)
    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1. Cho sơ đồ phản ứng sau:\n• X (C~4~H~8~O~2~) + H~2~O --H^+, t°--> Y + Z\n• Y + C~3~H~5~(OH)~3~ --H~2~SO~4~ đặc, t°--> T\n• Z --men giấm--> Y\n• Z + HCOOH --H~2~SO~4~ đặc, t°--> W\nBiết rằng X, Y, Z, T, W là các hợp chất hữu cơ khác nhau; T chỉ chứa một loại nhóm chức.",
            "statements": [
                ("a. Y có trong nọc độc của kiến.", "S"),
                ("b. Phần trăm khối lượng của nguyên tố oxygen trong T là 36,36%.", "S"),
                ("c. Công thức cấu tạo của W là HCOOC~2~H~5~.", "Đ"),
                ("d. Z có trong thành phần của nước rửa tay khô có tác dụng diệt khuẩn.", "Đ")
            ],
            "solution": "Lập luận giải mã sơ đồ chuyển hóa:\n- Z lên men giấm tạo thành Y ⇒ Z là ethanol (C~2~H~5~OH) và Y là acetic acid (CH~3~COOH).\n- X (C~4~H~8~O~2~) thủy phân tạo Y (CH~3~COOH) và Z (C~2~H~5~OH) ⇒ X là ethyl acetate (CH~3~COOC~2~H~5~).\n- Y (CH~3~COOH) phản ứng với glycerol (C~3~H~5~(OH)~3~) tạo ester T chỉ chứa một loại nhóm chức ⇒ T là triacetin: (CH~3~COO)~3~C~3~H~5~ (C~9~H~14~O~6~).\n- Z (C~2~H~5~OH) phản ứng với formic acid (HCOOH) tạo ester W ⇒ W là ethyl formate (HCOOC~2~H~5~).\n• a) Sai. Y là acetic acid (CH~3~COOH) có trong giấm ăn. Acid có trong nọc độc của kiến là formic acid (HCOOH).\n• b) Sai. T là triacetin (C~9~H~14~O~6~, M = 218 g/mol). Phần trăm khối lượng oxygen trong T là:\n  %O = (6 × 16 / 218) × 100% = (96 / 218) × 100% ≈ 44,04% (khác 36,36%).\n• c) Đúng. Ester W được tạo thành từ HCOOH và C~2~H~5~OH nên có công thức là HCOOC~2~H~5~.\n• d) Đúng. Z là ethanol (C~2~H~5~OH), cồn y tế 70 - 75° là thành phần sát khuẩn chính trong các dung dịch nước rửa tay khô."
        },
        {
            "id": 2,
            "prompt": "Câu 2. Cho các chất có công thức cấu tạo sau:\n(1) H~2~NCH~2~COOH (glycine)\n(2) C~2~H~5~COOH (propanoic acid)\n(3) C~2~H~5~NH~2~ (ethylamine)\n(4) H~2~NCH~2~CH~2~CH(NH~2~)COOH (2,4-diaminobutanoic acid)\n(5) C~6~H~5~NH~2~ (aniline)",
            "statements": [
                ("a. Có 4 chất tác dụng được với dung dịch HCl.", "Đ"),
                ("b. Có 1 chất tác dụng được với ethanol khi có mặt xúc tác H~2~SO~4~ đặc, đun nóng.", "S"),
                ("c. Có 3 chất tác dụng được với dung dịch NaOH.", "Đ"),
                ("d. Có 2 chất vừa tác dụng được với acid vừa tác dụng được với base.", "Đ")
            ],
            "solution": "Phân tích tính chất hóa học của 5 chất:\n• a) Đúng. Tác dụng với dung dịch HCl gồm các chất có nhóm amino (-NH~2~) có tính base hoặc amino acid lưỡng tính: (1) glycine, (3) ethylamine, (4) diamino acid, (5) aniline. Chỉ có (2) carboxylic acid không tác dụng với HCl. Vậy có 4 chất.\n• b) Sai. Tác dụng với ethanol (C~2~H~5~OH, xúc tác H~2~SO~4~ đặc, t°) qua phản ứng este hóa gồm các chất chứa nhóm carboxyl (-COOH): (1) H~2~NCH~2~COOH, (2) C~2~H~5~COOH và (4) diamino monocarboxylic acid. Có 3 chất tham gia phản ứng este hóa với ethanol (không phải 1 chất).\n• c) Đúng. Tác dụng với dung dịch NaOH gồm các chất có nhóm carboxyl (-COOH) có tính acid hoặc lưỡng tính: (1) glycine, (2) propanoic acid, (4) diamino acid. Các chất (3) và (5) chỉ có tính base, không tác dụng với NaOH. Vậy có 3 chất.\n• d) Đúng. Các chất vừa tác dụng được với dung dịch acid vừa tác dụng được với dung dịch base là các amino acid lưỡng tính: chất (1) H~2~NCH~2~COOH và chất (4) H~2~NCH~2~CH~2~CH(NH~2~)COOH (có cả nhóm -NH~2~ và nhóm -COOH). Vậy có 2 chất."
        }
    ],

    # PHẦN III: (2,0 điểm) Câu trắc nghiệm yêu cầu trả lời ngắn. (4 câu)
    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1. Cao su lưu hóa có 2% sulfur (S) về khối lượng. Có bao nhiêu mắt xích isoprene có một cầu disulfide (-S-S-)? Giả thiết rằng S đã thay thế cho H ở cầu methylene (-CH~2~-) trong mạch cao su.",
            "answer": "46",
            "solution": "Gọi n là số mắt xích isoprene (-C~5~H~8~-) ứng với 1 cầu disulfide (-S-S-):\n- Cứ một cầu disulfide (-S-S-) nối 2 mạch cao su, chứa 2 nguyên tử S và đã thay thế 2 nguyên tử H ở các nhóm methylene.\n- Khối lượng mol của 1 mắt xích isoprene: M(C~5~H~8~) = 5 × 12 + 8 × 1 = 68 g/mol.\n- Khối lượng đoạn mạch chứa n mắt xích và 1 cầu -S-S- là:\n  M = 68n - 2 × 1 + 2 × 32 = 68n + 62 g/mol.\n- Phần trăm khối lượng của sulfur trong cao su lưu hóa là 2%:\n  %S = [64 / (68n + 62)] × 100% = 2%\n  ⇔ 68n + 62 = 64 / 0,02 = 3 200\n  ⇔ 68n = 3 138\n  ⇔ n = 3 138 / 68 ≈ 46,147.\nLàm tròn số nguyên mắt xích: 46 mắt xích.\n*Đáp án: 46*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Thủy phân hoàn toàn 1 mol peptide mạch hở X, thu được 2 mol Gly, 2 mol Ala và 1 mol Val. Mặt khác, thủy phân không hoàn toàn X, thu được hỗn hợp các amino acid và các peptide (trong đó có Gly-Ala-Val). Số công thức cấu tạo phù hợp với tính chất của X là bao nhiêu?",
            "answer": "6",
            "solution": "- Thủy phân hoàn toàn 1 mol X thu được 2 mol Gly, 2 mol Ala, 1 mol Val ⇒ X là pentapeptide tạo bởi (Gly)~2~(Ala)~2~(Val).\n- Thủy phân không hoàn toàn X có chứa đoạn tripeptide: Gly-Ala-Val.\n  Coi cụm (Gly-Ala-Val) là một khối thống nhất Z.\n  Các amino acid còn lại chưa ghép là: 1 Gly và 1 Ala.\n- Cấu trúc của pentapeptide X là sự sắp xếp của 3 khối: {Gly, Ala, Z}.\n- Số đồng phân cấu tạo tương ứng với số hoán vị của 3 khối này:\n  P~3~ = 3! = 6 công thức cấu tạo.\nCụ thể các chuỗi peptide gồm:\n  1. Gly - Ala - [Gly - Ala - Val]\n  2. Ala - Gly - [Gly - Ala - Val]\n  3. Gly - [Gly - Ala - Val] - Ala\n  4. Ala - [Gly - Ala - Val] - Gly\n  5. [Gly - Ala - Val] - Gly - Ala\n  6. [Gly - Ala - Val] - Ala - Gly.\nCả 6 cấu tạo đều khác biệt và đều thỏa mãn.\n*Đáp án: 6*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Cho các polymer sau: polybuta-1,3-diene, poly(vinyl acetate), polystyrene, nylon-6,6 và poly(phenol formaldehyde). Số polymer được điều chế bằng phản ứng trùng ngưng là bao nhiêu?",
            "answer": "2",
            "solution": "Phân loại phương pháp tổng hợp từng polymer:\n1. Polybuta-1,3-diene: Điều chế bằng phản ứng trùng hợp buta-1,3-diene.\n2. Poly(vinyl acetate) (PVA): Điều chế bằng phản ứng trùng hợp vinyl acetate.\n3. Polystyrene (PS): Điều chế bằng phản ứng trùng hợp styrene.\n4. Nylon-6,6: Điều chế bằng phản ứng trùng ngưng giữa hexamethylenediamine và adipic acid.\n5. Poly(phenol formaldehyde) (nhựa PPF): Điều chế bằng phản ứng trùng ngưng giữa phenol và formaldehyde.\n⇒ Có 2 polymer được điều chế bằng phản ứng trùng ngưng: nylon-6,6 và poly(phenol formaldehyde).\n*Đáp án: 2*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Thuỷ phân 4,4 gam ethyl acetate bằng 300 mL dung dịch NaOH 0,2M. Sau khi phản ứng xảy ra hoàn toàn, cô cạn dung dịch, thu được m gam chất rắn khan. Giá trị của m là bao nhiêu?",
            "answer": "4,5",
            "solution": "- Số mol ethyl acetate (CH~3~COOC~2~H~5~): n(ester) = 4,4 / 88 = 0,05 mol.\n- Số mol NaOH: n(NaOH) = 0,3 × 0,2 = 0,06 mol.\n- Phương trình phản ứng xà phòng hóa:\n  CH~3~COOC~2~H~5~ + NaOH → CH~3~COONa + C~2~H~5~OH\n  0,05 mol            0,05 mol    0,05 mol      0,05 mol\n- Vì n(NaOH) = 0,06 mol > 0,05 mol nên NaOH còn dư: n(NaOH dư) = 0,06 - 0,05 = 0,01 mol.\n- Khi cô cạn dung dịch, ethanol (C~2~H~5~OH) và nước bay hơi hoàn toàn, chất rắn khan còn lại gồm:\n  + Muối CH~3~COONa: 0,05 mol ⇒ m = 0,05 × 82 = 4,1 gam.\n  + NaOH dư: 0,01 mol ⇒ m = 0,01 × 40 = 0,4 gam.\n⇒ Khối lượng chất rắn khan thu được là:\n  m = 4,1 + 0,4 = 4,5 gam.\n*Đáp án: 4,5*"
        }
    ],

    # B. PHẦN TỰ LUẬN: (3 điểm)
    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1. (1,0 điểm)\nHợp chất X có công thức phân tử trùng với công thức đơn giản nhất, vừa tác dụng được với acid vừa tác dụng được với base trong điều kiện thích hợp. Trong phân tử X, thành phần phần trăm khối lượng của các nguyên tố C, H, N lần lượt bằng 40,449%; 7,865% và 15,73%; còn lại là oxygen. Khi cho 4,45 gam X phản ứng hoàn toàn với một lượng vừa đủ dung dịch NaOH (đun nóng) thu được 4,85 gam muối khan. Xác định công thức cấu tạo của X.",
            "num_lines": 10,
            "solution": "1. Xác định công thức phân tử của X:\n- Phần trăm khối lượng của oxygen trong X là:\n  %O = 100% - 40,449% - 7,865% - 15,73% = 35,956%.\n- Tỉ lệ số mol các nguyên tố C, H, N, O trong X:\n  nC : nH : nN : nO = (40,449 / 12) : (7,865 / 1) : (15,73 / 14) : (35,956 / 16)\n                   = 3,37075 : 7,865 : 1,12357 : 2,24725\n                   = 3 : 7 : 1 : 2.\n⇒ Công thức đơn giản nhất của X là C~3~H~7~O~2~N.\nVì công thức phân tử trùng với công thức đơn giản nhất nên CTPT của X là C~3~H~7~O~2~N (M = 89 g/mol).\n\n2. Xác định công thức cấu tạo của X:\n- Số mol X đem phản ứng: n(X) = 4,45 / 89 = 0,05 mol.\n- X phản ứng vừa đủ với NaOH tạo muối khan. Do X chứa 1 nguyên tử N và 2 nguyên tử O nên số mol muối thu được bằng số mol X:\n  n(muối) = n(X) = 0,05 mol.\n- Khối lượng mol của muối khan:\n  M(muối) = 4,85 / 0,05 = 97 g/mol.\n- Khối lượng mol muối (97 g/mol) lớn hơn phân tử khối của X (89 g/mol) đúng bằng: 97 - 89 = 8 g/mol.\n  Điều này chứng tỏ X là ester của amino acid: H~2~N-CH~2~-COOCH~3~ (M = 89 g/mol).\n  Phương trình phản ứng thủy phân trong NaOH:\n  H~2~N-CH~2~-COOCH~3~ + NaOH --t°--> H~2~N-CH~2~-COONa + CH~3~OH\n  Kiểm tra muối H~2~N-CH~2~-COONa: M = 16 + 14 + 44 + 23 = 97 g/mol (thỏa mãn 100%).\n- X vừa có nhóm -NH~2~ (tính base, tác dụng với acid) vừa có liên kết ester (thủy phân được trong môi trường base).\n*Kết luận: Công thức cấu tạo của X là H~2~N-CH~2~-COOCH~3~ (methyl glycinate).*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. (1,0 điểm)\nNhúng một thanh Zn vào 100 mL dung dịch CuSO~4~, sau một thời gian phản ứng lấy thanh Zn ra khỏi dung dịch, làm khô và đem cân thấy khối lượng thanh Zn giảm 0,01 gam. Cho dung dịch loãng NaOH dư vào dung dịch sau phản ứng thu được 0,49 gam kết tủa. Tính nồng độ mol của dung dịch CuSO~4~ ban đầu.",
            "num_lines": 10,
            "solution": "1. Xét phản ứng kim loại Zn đẩy Cu ra khỏi dung dịch CuSO~4~:\n  Zn + CuSO~4~ → ZnSO~4~ + Cu↓ (bám vào thanh Zn)\n- Khi 1 mol Zn tan ra (65 g) và 1 mol Cu giải phóng bám vào (64 g) thì khối lượng thanh Zn giảm đi: 65 - 64 = 1 gam.\n- Theo đề bài, khối lượng thanh Zn giảm 0,01 gam, do đó số mol Zn đã phản ứng là:\n  n(Zn pư) = n(CuSO~4~ pư) = 0,01 / (65 - 64) = 0,01 mol.\n\n2. Xét phản ứng của dung dịch sau phản ứng với dung dịch NaOH dư:\n- Dung dịch sau phản ứng gồm có muối ZnSO~4~ vừa tạo thành (0,01 mol) và CuSO~4~ còn dư.\n- Khi cho dung dịch NaOH loãng dư vào:\n  ZnSO~4~ + 2 NaOH → Zn(OH)~2~↓ + Na~2~SO~4~\n  Vì NaOH dùng dư nên kết tủa lưỡng tính Zn(OH)~2~ tan hoàn toàn:\n  Zn(OH)~2~ + 2 NaOH → Na~2~[Zn(OH)~4~]  (hoặc Na~2~ZnO~2~ + 2 H~2~O)\n- Kết tủa thu được chỉ là Cu(OH)~2~ do CuSO~4~ dư tạo thành:\n  CuSO~4~(dư) + 2 NaOH → Cu(OH)~2~↓ + Na~2~SO~4~\n- Số mol kết tủa Cu(OH)~2~ (M = 98 g/mol):\n  n[Cu(OH)~2~] = 0,49 / 98 = 0,005 mol.\n  ⇒ n(CuSO~4~ dư) = 0,005 mol.\n\n3. Tính nồng độ mol của dung dịch CuSO~4~ ban đầu:\n- Tổng số mol CuSO~4~ có trong dung dịch ban đầu là:\n  n(CuSO~4~ bđ) = n(CuSO~4~ pư) + n(CuSO~4~ dư) = 0,01 + 0,005 = 0,015 mol.\n- Thể tích dung dịch V = 100 mL = 0,1 L.\n- Nồng độ mol của dung dịch CuSO~4~ ban đầu là:\n  C~M~(CuSO~4~) = 0,015 / 0,1 = 0,15 M.\n*Kết luận: Nồng độ mol của dung dịch CuSO~4~ ban đầu là 0,15 M.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. (1,0 điểm)\nĐể sản xuất 10 L cồn y tế 70° người ta cần lên men tối thiểu m gam glucose. Biết hiệu suất của quá trình lên men là 80%, khối lượng riêng của ethanol là 0,789 g/mL. Giá trị của m là bao nhiêu?",
            "num_lines": 8,
            "solution": "1. Tính khối lượng và số mol ethanol cần có trong 10 L cồn y tế 70°:\n- Thể tích ethanol nguyên chất có trong 10 lít cồn 70°:\n  V(ethanol) = 10 L × (70 / 100) = 7 L = 7 000 mL.\n- Khối lượng ethanol nguyên chất (với D = 0,789 g/mL):\n  m(ethanol) = V × D = 7 000 mL × 0,789 g/mL = 5 523 gam.\n- Khối lượng mol của ethanol (C~2~H~5~OH): M = 46 g/mol.\n  ⇒ Số mol ethanol: n(ethanol) = 5 523 / 46 = 120,0652 mol.\n\n2. Tính khối lượng glucose tối thiểu cần dùng:\n- Phương trình phản ứng lên men rượu:\n  C~6~H~12~O~6~ --men rượu--> 2 C~2~H~5~OH + 2 CO~2~↑\n- Theo phương trình, số mol glucose cần theo lí thuyết là:\n  n(glucose lt) = n(ethanol) / 2 = 120,0652 / 2 = 60,0326 mol.\n- Vì hiệu suất phản ứng lên men là H = 80% = 0,8 nên số mol glucose thực tế cần lấy là:\n  n(glucose tt) = n(glucose lt) / 0,8 = 60,0326 / 0,8 ≈ 75,04076 mol.\n- Khối lượng mol của glucose (C~6~H~12~O~6~): M = 180 g/mol.\n  ⇒ Khối lượng glucose tối thiểu cần dùng là:\n  m = n(glucose tt) × 180 = 75,04076 × 180 ≈ 13 507,34 gam (hay khoảng 13,51 kg).\n*Kết luận: Giá trị của m là khoảng 13 507,34 gam (hoặc 13,51 kg).* "
        }
    ]
}

def generate_docx(is_teacher=False):
    doc = Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line=DATA_EXAM_1003["school_line"],
        exam_title=DATA_EXAM_1003["exam_title"],
        subject_line=DATA_EXAM_1003["subject_line"],
        num_pages=DATA_EXAM_1003["num_pages"],
        duration=DATA_EXAM_1003["duration"],
        exam_code=DATA_EXAM_1003["exam_code"],
        atomic_masses=DATA_EXAM_1003["atomic_masses"],
        is_teacher=is_teacher
    )

    # Thêm BẢNG ĐÁP ÁN RÚT GỌN (BAO GỒM CẢ TỰ LUẬN) cho Bản Giáo Viên
    if is_teacher:
        add_quick_answer_table_1003(doc)

    # A. PHẦN TRẮC NGHIỆM: (7 điểm)
    add_section_header(
        doc,
        title="A. PHẦN TRẮC NGHIỆM: (7 điểm)",
        subtitle="PHẦN I: (3,0 điểm) Câu trắc nghiệm nhiều phương án lựa chọn. (Thí sinh trả lời từ câu 1 đến câu 12, mỗi câu hỏi thí sinh chỉ chọn một phương án)"
    )

    for item in DATA_EXAM_1003["part1"]:
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

    # PHẦN II: (2,0 điểm) Câu trắc nghiệm đúng sai.
    add_section_header(
        doc,
        title="PHẦN II: (2,0 điểm) Câu trắc nghiệm đúng sai.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    for item in DATA_EXAM_1003["part2"]:
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

    # PHẦN III: (2,0 điểm) Câu trắc nghiệm yêu cầu trả lời ngắn.
    add_section_header(
        doc,
        title="PHẦN III: (2,0 điểm) Câu trắc nghiệm yêu cầu trả lời ngắn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Thí sinh điền kết quả vào ô trả lời tương ứng."
    )

    for item in DATA_EXAM_1003["part3"]:
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

    # B. PHẦN TỰ LUẬN: (3 điểm)
    add_section_header(
        doc,
        title="B. PHẦN TỰ LUẬN: (3 điểm)",
        subtitle="Thí sinh trình bày chi tiết lời giải vào giấy làm bài thi."
    )

    for item in DATA_EXAM_1003["part4"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])
        elif not is_teacher:
            num_lines = item.get("num_lines", 8)
            add_student_writing_lines(doc, num_lines=num_lines)

    # Phần kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(0)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(r_end, font_size=11, bold=True)

    suffix = "_GiaoVien.docx" if is_teacher else "_HocSinh.docx"
    filename = f"De_Kiem_Tra_CK1_Hoa_12_NguyenDinhChieu_Ma1003{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
