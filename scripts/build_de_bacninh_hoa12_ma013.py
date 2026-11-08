# -*- coding: utf-8 -*-
"""
Script tạo Đề Thi Thử Tốt Nghiệp THPT Quốc Gia Lần 1 (Năm Học 2026 - 2027)
Sở Giáo Dục và Đào Tạo Bắc Ninh - Môn: Hoá Học - Mã Đề 013
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)

Xuất bản 2 file Word:
1. Bản Học sinh: De_Kiem_Tra_Hoa_12_BacNinh_Ma013.docx
2. Bản Giáo viên: De_Kiem_Tra_Hoa_12_BacNinh_Ma013_GiaoVien.docx
"""

import os
import sys

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

sys.path.append(os.path.abspath(r"C:\Antigravity_Thanh\scripts"))
from form_de_kiem_tra import (
    setup_header_footer,
    build_title_block,
    add_section_header,
    render_question_prompt,
    render_options_no_table,
    render_statement,
    add_solution_no_table,
    render_rich_text,
    set_run_style,
    add_p_border_bottom,
    COLOR_PRIMARY,
    COLOR_CORRECT,
    COLOR_SOLUTION,
    COLOR_GRAY
)

OUTPUT_DIR = r"C:\Antigravity_Thanh\San_Pham"
IMG_DIR = os.path.join(OUTPUT_DIR, "images_bacninh_12_ma013")

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

def add_quick_answer_table_013(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu Mã đề 013.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 013)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (18 CÂU)*", base_size=11, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=4, cols=10)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_p1_1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9"]
    answers_p1_1 = ["Đ/A", "B", "C", "D", "C", "A", "B", "A", "C", "B"]
    headers_p1_2 = ["Câu", "10", "11", "12", "13", "14", "15", "16", "17", "18"]
    answers_p1_2 = ["Đ/A", "D", "A", "B", "B", "A", "A", "D", "D", "B"]

    col_w_0 = Cm(1.8)
    col_w = Cm(1.8)

    for r_idx, r_data in enumerate([headers_p1_1, answers_p1_1, headers_p1_2, answers_p1_2]):
        row = tbl1.rows[r_idx]
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = col_w_0 if c_idx == 0 else col_w
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            is_header_row = (r_idx in [0, 2])
            is_first_col = (c_idx == 0)

            if is_header_row:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                set_cell_shading(cell, "FFFFFF" if not is_first_col else "F5F5F5")
                c_color = COLOR_CORRECT if not is_first_col else COLOR_PRIMARY
                render_rich_text(p, f"*{val}*", base_size=10.5, base_bold=True, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # 2. PHẦN II
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(6)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (4 CÂU)*", base_size=11, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=5, cols=5)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl2.autofit = False

    tbl2_data = [
        ["Câu hỏi", "Lệnh hỏi a)", "Lệnh hỏi b)", "Lệnh hỏi c)", "Lệnh hỏi d)"],
        ["Câu 19", "Đúng", "Sai", "Sai", "Đúng"],
        ["Câu 20", "Đúng", "Sai", "Đúng", "Đúng"],
        ["Câu 21", "Đúng", "Sai", "Đúng", "Sai"],
        ["Câu 22", "Sai", "Đúng", "Sai", "Đúng"]
    ]

    w_col0 = Cm(3.0)
    w_col_item = Cm(3.75)

    for r_idx, row_vals in enumerate(tbl2_data):
        row = tbl2.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = w_col0 if c_idx == 0 else w_col_item
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                is_label = (c_idx == 0)
                set_cell_shading(cell, "F5F5F5" if is_label else "FFFFFF")
                c_color = COLOR_PRIMARY if is_label else (COLOR_CORRECT if val == "Đúng" else COLOR_GRAY)
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # 3. PHẦN III
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(6)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (6 CÂU)*", base_size=11, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=7)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False

    tbl3_headers = ["Câu", "Câu 23", "Câu 24", "Câu 25", "Câu 26", "Câu 27", "Câu 28"]
    tbl3_answers = ["Đáp án", "1256", "3241", "42", "1686", "29,9", "3"]

    w3_c0 = Cm(2.4)
    w3_c = Cm(2.6)

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
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                is_label = (c_idx == 0)
                set_cell_shading(cell, "F5F5F5" if is_label else "FFFFFF")
                c_color = COLOR_PRIMARY if is_label else COLOR_CORRECT
                render_rich_text(p, f"*{val}*", base_size=10.5, base_bold=True, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # Đường phân cách
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(6)
    p_div.paragraph_format.space_after = Pt(8)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


def build_exam_bacninh_013(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="SỞ GIÁO DỤC VÀ ĐÀO TẠO BẮC NINH",
        exam_title="THI THỬ TN THPT QUỐC GIA LẦN 1 - NĂM HỌC 2026-2027",
        subject_line="MÔN: HOÁ HỌC - LỚP 12",
        num_pages="04",
        duration="50 phút (không tính thời gian phát đề)",
        exam_code="013",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; K = 39; S = 32; Cl = 35,5; Ba = 137",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_013(doc)

    # ==========================================
    # PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN
    # ==========================================
    add_section_header(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # CÂU 1
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(4)
    p1.paragraph_format.space_after = Pt(2)
    render_question_prompt(p1, "Câu 1: Chất nào sau đây được sử dụng làm chất giặt rửa tổng hợp?")
    opts1 = [
        "A. CH~3~COOK.",
        "*B. CH~3~[CH~2~]~11~OSO~3~Na.*" if is_teacher else "B. CH~3~[CH~2~]~11~OSO~3~Na.",
        "C. C~15~H~31~COONa.",
        "D. C~15~H~31~COOCH~3~."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        sol1 = (
            "• CH~3~[CH~2~]~11~OSO~3~Na (sodium lauryl sulfate) là chất giặt rửa tổng hợp gốc sulfate.\n"
            "• CH~3~COOK và C~15~H~31~COONa là các muối carboxylate (thành phần của xà phòng).\n"
            "• C~15~H~31~COOCH~3~ là ester (methyl palmitate), không có tính chất giặt rửa.\n"
            "➔ *Đáp án đúng: B*"
        )
        add_solution_no_table(doc, sol1)

    # CÂU 2
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(4)
    p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p2, "Câu 2: Cho dung dịch H~2~SO~4~ loãng vào dung dịch X thấy sủi bọt khí và sinh ra kết tủa. Dung dịch X là")
    opts2 = [
        "A. Na~2~SO~4~.",
        "B. KNO~3~.",
        "*C. Ba(HCO~3~)~2~.*" if is_teacher else "C. Ba(HCO~3~)~2~.",
        "D. BaCl~2~."
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        sol2 = (
            "• Khi cho H~2~SO~4~ vào dung dịch Ba(HCO~3~)~2~ xảy ra phản ứng đồng thời tạo kết tủa trắng và sủi bọt khí:\n"
            "  Ba(HCO~3~)~2~ + H~2~SO~4~ → BaSO~4~↓ (trắng) + 2CO~2~↑ + 2H~2~O\n"
            "• Dung dịch BaCl~2~ chỉ cho kết tủa BaSO~4~, không có khí thoát ra.\n"
            "• Dung dịch Na~2~SO~4~ và KNO~3~ không phản ứng.\n"
            "➔ *Đáp án đúng: C*"
        )
        add_solution_no_table(doc, sol2)

    # CÂU 3
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(4)
    p3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p3, "Câu 3: Một dung dịch baking soda có pH = 8,3. Nồng độ ion OH^-^ trong dung dịch trên là")
    opts3 = [
        "A. 5,7 M.",
        "B. 8,3 M.",
        "C. 10^-8,3^ M.",
        "*D. 10^-5,7^ M.*" if is_teacher else "D. 10^-5,7^ M."
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        sol3 = (
            "• Ở 25°C, tích số ion của nước là K~w~ = [H^+^].[OH^-^] = 10^-14^, do đó:\n"
            "  pH + pOH = 14 ➔ pOH = 14 - pH = 14 - 8,3 = 5,7.\n"
            "• Nồng độ ion OH^-^ trong dung dịch là: [OH^-^] = 10^-pOH^ = 10^-5,7^ M.\n"
            "➔ *Đáp án đúng: D*"
        )
        add_solution_no_table(doc, sol3)

    # CÂU 4
    p4 = doc.add_paragraph()
    p4.paragraph_format.space_before = Pt(4)
    p4.paragraph_format.space_after = Pt(2)
    render_question_prompt(p4, "Câu 4: Phát biểu nào sau đây đúng?")
    opts4 = [
        "A. Amylose và amylopectin đều có cấu trúc mạch phân nhánh.",
        "B. Trong phân tử glucose chỉ có 4 nhóm alcohol (OH).",
        "*C. Saccharose không tham gia phản ứng tráng bạc.*" if is_teacher else "C. Saccharose không tham gia phản ứng tráng bạc.",
        "D. Glucose và fructose đều làm mất màu nước bromine."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        sol4 = (
            "• Phát biểu A sai vì amylose có cấu trúc mạch không phân nhánh, xoắn lò xo.\n"
            "• Phát biểu B sai vì trong phân tử glucose dạng mạch hở có 5 nhóm alcohol (-OH) kề nhau.\n"
            "• Phát biểu C đúng vì phân tử saccharose được tạo bởi gốc α-glucose và β-fructose qua liên kết α-1,2-glycosidic, không còn nhóm -OH hemiacetal hay hemiketal tự do để mở vòng tạo nhóm aldehyde nên không tráng bạc.\n"
            "• Phát biểu D sai vì fructose không làm mất màu nước bromine (nước bromine dùng phân biệt glucose và fructose).\n"
            "➔ *Đáp án đúng: C*"
        )
        add_solution_no_table(doc, sol4)

    # CÂU 5
    p5 = doc.add_paragraph()
    p5.paragraph_format.space_before = Pt(4)
    p5.paragraph_format.space_after = Pt(2)
    render_question_prompt(p5, "Câu 5: Cho 4 dung dịch đựng trong 4 lọ mất nhãn sau: HCl, Na~2~SO~4~, NaCl, Ba(OH)~2~. Chỉ dùng thêm 1 thuốc thử nào dưới đây để nhận biết các dung dịch trên?")
    opts5 = [
        "*A. quỳ tím.*" if is_teacher else "A. quỳ tím.",
        "B. H~2~SO~4~.",
        "C. AgNO~3~.",
        "D. BaCl~2~."
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        sol5 = (
            "• Dùng quỳ tím nhúng vào các mẫu thử:\n"
            "  + Mẫu làm quỳ tím chuyển sang màu đỏ: HCl.\n"
            "  + Mẫu làm quỳ tím chuyển sang màu xanh: Ba(OH)~2~.\n"
            "  + Hai mẫu không làm đổi màu quỳ tím: Na~2~SO~4~ và NaCl (nhóm I).\n"
            "• Lấy dung dịch Ba(OH)~2~ vừa nhận biết nhỏ lần lượt vào 2 mẫu nhóm I:\n"
            "  + Xuất hiện kết tủa trắng không tan: Na~2~SO~4~ (Ba(OH)~2~ + Na~2~SO~4~ → BaSO~4~↓ + 2NaOH).\n"
            "  + Không có hiện tượng gì: NaCl.\n"
            "➔ *Đáp án đúng: A*"
        )
        add_solution_no_table(doc, sol5)

    # CÂU 6
    p6 = doc.add_paragraph()
    p6.paragraph_format.space_before = Pt(4)
    p6.paragraph_format.space_after = Pt(2)
    render_question_prompt(p6, "Câu 6: Chất nào sau đây thuộc loại acid béo omega-3?")

    # Chèn hình vẽ 4 chất acid béo
    p6_img = doc.add_paragraph()
    p6_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p6_img.paragraph_format.space_before = Pt(2)
    p6_img.paragraph_format.space_after = Pt(2)
    img_c6 = os.path.join(IMG_DIR, "cau6_acid_beo.png")
    if os.path.exists(img_c6):
        p6_img.add_run().add_picture(img_c6, width=Cm(14.0))

    opts6 = [
        "A. 3.",
        "*B. 4.*" if is_teacher else "B. 4.",
        "C. 2.",
        "D. 1."
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        sol6 = (
            "• Acid béo omega-3 (ω-3) là acid béo không bão hòa đa có liên kết đôi C=C đầu tiên bắt đầu từ nguyên tử carbon số 3 tính từ đầu methyl (-CH~3~, tức vị trí carbon ω-1).\n"
            "• Phân tích từng chất từ hình vẽ:\n"
            "  + Chất (1): Có liên kết đôi C=C đầu tiên ở vị trí C9 tính từ đầu methyl ➔ acid béo omega-9 (oleic acid).\n"
            "  + Chất (2): Có 4 liên kết đôi C=C, liên kết đôi đầu tiên ở vị trí C6 ➔ acid béo omega-6 (arachidonic acid).\n"
            "  + Chất (3): Có 2 liên kết đôi C=C, liên kết đôi đầu tiên ở vị trí C6 ➔ acid béo omega-6 (linoleic acid).\n"
            "  + Chất (4): Có 3 liên kết đôi C=C, liên kết đôi đầu tiên xuất hiện ở vị trí carbon số 3 tính từ đầu methyl (C1-C2-C3=C4...) ➔ acid béo omega-3 (alpha-linolenic acid - ALA).\n"
            "➔ *Đáp án đúng: B*"
        )
        add_solution_no_table(doc, sol6)

    # CÂU 7
    p7 = doc.add_paragraph()
    p7.paragraph_format.space_before = Pt(4)
    p7.paragraph_format.space_after = Pt(2)
    render_question_prompt(p7, "Câu 7: Thủy phân hoàn toàn hỗn hợp phenyl acetate và ethyl acetate trong dung dịch NaOH, thu được sản phẩm gồm")
    opts7 = [
        "*A. 2 muối và 1 alcohol.*" if is_teacher else "A. 2 muối và 1 alcohol.",
        "B. 2 muối và 2 alcohol.",
        "C. 1 muối và 1 alcohol.",
        "D. 1 muối và 2 alcohol."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        sol7 = (
            "• Phương trình hóa học thủy phân trong dung dịch NaOH:\n"
            "  CH~3~COOC~6~H~5~ + 2NaOH → CH~3~COONa + C~6~H~5~ONa + H~2~O (thu được 2 muối, không có alcohol).\n"
            "  CH~3~COOC~2~H~5~ + NaOH → CH~3~COONa + C~2~H~5~OH (thu được 1 muối và 1 alcohol).\n"
            "• Tổng sản phẩm hữu cơ gồm: 2 muối (CH~3~COONa, C~6~H~5~ONa) và 1 alcohol (C~2~H~5~OH).\n"
            "➔ *Đáp án đúng: A*"
        )
        add_solution_no_table(doc, sol7)

    # CÂU 8
    p8 = doc.add_paragraph()
    p8.paragraph_format.space_before = Pt(4)
    p8.paragraph_format.space_after = Pt(2)
    render_question_prompt(p8, "Câu 8: Ester X có công thức phân tử là C~4~H~8~O~2~, X có khả năng tham gia phản ứng với dung dịch AgNO~3~ trong NH~3~. Số công thức cấu tạo của ester X thỏa mãn điều kiện trên là")
    opts8 = [
        "A. 3.",
        "B. 4.",
        "*C. 2.*" if is_teacher else "C. 2.",
        "D. 1."
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        sol8 = (
            "• Ester X (C~4~H~8~O~2~) có khả năng tham gia phản ứng tráng bạc phải là ester của formic acid (dạng HCOOR').\n"
            "• Gốc R' có 3 nguyên tử carbon: -C~3~H~7~. Gốc này có 2 đồng phân mạch carbon:\n"
            "  (1) HCOOCH~2~CH~2~CH~3~ (propyl formate)\n"
            "  (2) HCOOCH(CH~3~)~2~ (isopropyl formate)\n"
            "• Vậy có đúng 2 công thức cấu tạo thỏa mãn.\n"
            "➔ *Đáp án đúng: C*"
        )
        add_solution_no_table(doc, sol8)

    # CÂU 9
    p9 = doc.add_paragraph()
    p9.paragraph_format.space_before = Pt(4)
    p9.paragraph_format.space_after = Pt(2)
    render_question_prompt(p9, "Câu 9: Các khí sinh ra trong thí nghiệm phản ứng của saccharose với dung dịch H~2~SO~4~ đặc bao gồm")
    opts9 = [
        "A. SO~2~, H~2~S.",
        "*B. CO~2~, SO~2~.*" if is_teacher else "B. CO~2~, SO~2~.",
        "C. CO~2~, SO~3~.",
        "D. CO~2~, H~2~S."
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        sol9 = (
            "• Ban đầu H~2~SO~4~ đặc hút nước của saccharose làm đường bị than hóa tạo carbon:\n"
            "  C~12~H~22~O~11~ ➔ 12C + 11H~2~O\n"
            "• Một phần carbon sinh ra bị H~2~SO~4~ đặc nóng oxi hóa thành CO~2~ và SO~2~:\n"
            "  C + 2H~2~SO~4~ (đặc) ➔ CO~2~↑ + 2SO~2~↑ + 2H~2~O\n"
            "• Vậy các khí sinh ra bao gồm CO~2~ và SO~2~.\n"
            "➔ *Đáp án đúng: B*"
        )
        add_solution_no_table(doc, sol9)

    # CÂU 10
    p10 = doc.add_paragraph()
    p10.paragraph_format.space_before = Pt(4)
    p10.paragraph_format.space_after = Pt(2)
    render_question_prompt(p10, "Câu 10: Để tráng một lớp bạc lên ruột phích, người ta cho chất X phản ứng với lượng dư dung dịch AgNO~3~ trong NH~3~, đun nóng nhẹ 60-70°C. Chất X là")
    opts10 = [
        "A. tinh bột.",
        "B. ethyl acetate.",
        "C. saccharose.",
        "*D. glucose.*" if is_teacher else "D. glucose."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        sol10 = (
            "• Trong công nghiệp sản xuất ruột phích (bình thủy), người ta dùng glucose để tráng gương vì glucose an toàn, không độc hại và giá thành rẻ hơn formaldehyde.\n"
            "• Tinh bột, saccharose và ethyl acetate không có phản ứng tráng bạc.\n"
            "➔ *Đáp án đúng: D*"
        )
        add_solution_no_table(doc, sol10)

    # CÂU 11
    p11 = doc.add_paragraph()
    p11.paragraph_format.space_before = Pt(4)
    p11.paragraph_format.space_after = Pt(2)
    render_question_prompt(p11, "Câu 11: Trong quá trình trồng trọt, người nông dân được khuyến cáo không bón vôi sống (thành phần chính là CaO) cùng với phân đạm ammonium. Nguyên nhân của khuyến cáo này là")
    opts11 = [
        "*A. thất thoát đạm vì giải phóng ammonia.*" if is_teacher else "A. thất thoát đạm vì giải phóng ammonia.",
        "B. tạo thành hỗn hợp gây cháy nổ.",
        "C. tạo acid làm ảnh hưởng tới cây trồng.",
        "D. làm tăng độ chua của đất."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        sol11 = (
            "• Vôi sống khi gặp nước trong đất chuyển thành base Ca(OH)~2~:\n"
            "  CaO + H~2~O → Ca(OH)~2~\n"
            "• Môi trường kiềm OH^-^ sẽ tác dụng với ion NH~4~^+^ của phân đạm giải phóng khí ammonia (NH~3~) bay mất:\n"
            "  Ca(OH)~2~ + 2NH~4~^+^ → Ca^2+^ + 2NH~3~↑ + 2H~2~O\n"
            "• Hiện tượng này làm thất thoát lượng nitơ cung cấp cho cây trồng.\n"
            "➔ *Đáp án đúng: A*"
        )
        add_solution_no_table(doc, sol11)

    # CÂU 12
    p12 = doc.add_paragraph()
    p12.paragraph_format.space_before = Pt(4)
    p12.paragraph_format.space_after = Pt(2)
    render_question_prompt(p12, "Câu 12: Công thức phân tử nào sau đây không thể của ester?")
    opts12 = [
        "A. C~3~H~6~O~2~.",
        "*B. C~3~H~8~O~2~.*" if is_teacher else "B. C~3~H~8~O~2~.",
        "C. C~3~H~4~O~2~.",
        "D. C~4~H~6~O~2~."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        sol12 = (
            "• Phân tử ester chứa ít nhất 1 nhóm chức ester (-COO-), do đó độ bất bão hòa k ≥ 1 (chứa tối thiểu 1 liên kết π trong nhóm C=O).\n"
            "• Xét C~3~H~8~O~2~: k = (2×3 + 2 - 8)/2 = 0 ➔ Hợp chất no mạch hở, không có liên kết π hay vòng nên không thể chứa nhóm chức ester (chỉ có thể là diether, diol no...).\n"
            "➔ *Đáp án đúng: B*"
        )
        add_solution_no_table(doc, sol12)

    # CÂU 13
    p13 = doc.add_paragraph()
    p13.paragraph_format.space_before = Pt(4)
    p13.paragraph_format.space_after = Pt(2)
    render_question_prompt(p13, "Câu 13: Polysaccharide X là chất rắn, màu trắng, dạng sợi, trong bông nõn có gần 98% chất X. Chất X là")
    opts13 = [
        "A. glucose.",
        "*B. cellulose.*" if is_teacher else "B. cellulose.",
        "C. maltose.",
        "D. saccharose."
    ]
    render_options_no_table(doc, opts13, is_teacher=is_teacher)
    if is_teacher:
        sol13 = (
            "• Cellulose là polysaccharide mạch kéo dài dạng sợi, là thành phần cấu tạo màng tế bào thực vật, chiếm gần 98% trong sợi bông nõn và khoảng 50% trong gỗ.\n"
            "➔ *Đáp án đúng: B*"
        )
        add_solution_no_table(doc, sol13)

    # CÂU 14
    p14 = doc.add_paragraph()
    p14.paragraph_format.space_before = Pt(4)
    p14.paragraph_format.space_after = Pt(2)
    render_question_prompt(p14, "Câu 14: Có nhiều vụ tai nạn giao thông xảy ra do người lái xe uống rượu. Hàm lượng alcohol ethanol trong máu người lái xe không được vượt quá 0,02% theo khối lượng. Để xác định hàm lượng đó ta chuẩn độ alcohol bằng K~2~Cr~2~O~7~ trong môi trường acid theo phản ứng:\n3CH~3~CH~2~OH + K~2~Cr~2~O~7~ + 4H~2~SO~4~ → 3CH~3~CHO + Cr~2~(SO~4~)~3~ + K~2~SO~4~ + 7H~2~O\nKhi chuẩn độ 25,0 gam huyết tương máu của một người lái xe cần dùng 20,0 mL dung dịch K~2~Cr~2~O~7~ 0,010M. Nồng độ phần trăm ethanol trong 25 gam huyết tương máu của một người lái xe là")
    opts14 = [
        "*A. 0,11%.*" if is_teacher else "A. 0,11%.",
        "B. 0,03%.",
        "C. 0,05%.",
        "D. 0,027%."
    ]
    render_options_no_table(doc, opts14, is_teacher=is_teacher)
    if is_teacher:
        sol14 = (
            "• Số mol K~2~Cr~2~O~7~ tiêu tốn:\n"
            "  n(K~2~Cr~2~O~7~) = 0,020 L × 0,010 mol/L = 2,0 × 10^-4^ mol.\n"
            "• Theo phương trình chuẩn độ:\n"
            "  n(C~2~H~5~OH) = 3 × n(K~2~Cr~2~O~7~) = 3 × (2,0 × 10^-4^) = 6,0 × 10^-4^ mol.\n"
            "• Khối lượng ethanol có trong mẫu huyết tương:\n"
            "  m(C~2~H~5~OH) = 6,0 × 10^-4^ × 46 = 0,0276 gam.\n"
            "• Nồng độ phần trăm của ethanol trong máu:\n"
            "  C% = (0,0276 / 25,0) × 100% = 0,1104% ≈ 0,11%.\n"
            "➔ *Đáp án đúng: A*"
        )
        add_solution_no_table(doc, sol14)

    # CÂU 15
    p15 = doc.add_paragraph()
    p15.paragraph_format.space_before = Pt(4)
    p15.paragraph_format.space_after = Pt(2)
    render_question_prompt(p15, "Câu 15: Cho phản ứng hóa học sau: CH~3~COOH(l) + CH~3~OH(l) ⇌ CH~3~COOCH~3~(l) + H~2~O(l). Biểu thức tính hằng số cân bằng nồng độ K~C~ của phản ứng trên là")
    opts15 = [
        "*A. K~C~ = [CH~3~COOCH~3~].[H~2~O]/([CH~3~COOH].[CH~3~OH]).*" if is_teacher else "A. K~C~ = [CH~3~COOCH~3~].[H~2~O]/([CH~3~COOH].[CH~3~OH]).",
        "B. K~C~ = [CH~3~COOCH~3~]/([CH~3~COOH].[CH~3~OH]).",
        "C. K~C~ = [CH~3~COOH].[CH~3~OH]/([CH~3~COOCH~3~].[H~2~O]).",
        "D. K~C~ = [CH~3~COOH].[CH~3~OH]/[CH~3~COOCH~3~]."
    ]
    render_options_no_table(doc, opts15, is_teacher=is_teacher)
    if is_teacher:
        sol15 = (
            "• Trong hệ phản ứng ester hóa đồng thể lỏng không dùng dung môi nước, tất cả các chất kể cả H~2~O(l) đều tham gia vào trạng thái cân bằng với nồng độ biến thiên.\n"
            "• Theo định luật tác dụng khối lượng:\n"
            "  K~C~ = ([CH~3~COOCH~3~].[H~2~O]) / ([CH~3~COOH].[CH~3~OH]).\n"
            "➔ *Đáp án đúng: A*"
        )
        add_solution_no_table(doc, sol15)

    # CÂU 16
    p16 = doc.add_paragraph()
    p16.paragraph_format.space_before = Pt(4)
    p16.paragraph_format.space_after = Pt(2)
    render_question_prompt(p16, "Câu 16: Methyl acrylate là một chất kích thích mạnh, có thể gây chóng mặt, đau đầu, hoa mắt và khó thở khi tiếp xúc với da hoặc hít phải, ester này có công thức cấu tạo thu gọn là")
    opts16 = [
        "A. CH~3~COOC~2~H~5~.",
        "B. CH~3~COOCH~3~.",
        "C. C~2~H~5~COOCH~3~.",
        "*D. CH~2~=CHCOOCH~3~.*" if is_teacher else "D. CH~2~=CHCOOCH~3~."
    ]
    render_options_no_table(doc, opts16, is_teacher=is_teacher)
    if is_teacher:
        sol16 = (
            "• Gốc acrylate có công thức CH~2~=CH-COO-.\n"
            "• Gốc methyl có công thức -CH~3~.\n"
            "• Công thức cấu tạo của methyl acrylate là CH~2~=CHCOOCH~3~.\n"
            "➔ *Đáp án đúng: D*"
        )
        add_solution_no_table(doc, sol16)

    # CÂU 17
    p17 = doc.add_paragraph()
    p17.paragraph_format.space_before = Pt(4)
    p17.paragraph_format.space_after = Pt(2)
    render_question_prompt(p17, "Câu 17: Phương trình hoá học của phản ứng hydrate hóa ethylene để điều chế ethanol là: CH~2~=CH~2~ + H~2~O → CH~3~CH~2~OH. Giai đoạn (1) trong cơ chế của phản ứng trên xảy ra như sau:")

    # Hình giai đoạn cơ chế
    p17_img = doc.add_paragraph()
    p17_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p17_img.paragraph_format.space_before = Pt(2)
    p17_img.paragraph_format.space_after = Pt(2)
    img_c17 = os.path.join(IMG_DIR, "cau17_co_che.png")
    if os.path.exists(img_c17):
        p17_img.add_run().add_picture(img_c17, width=Cm(8.0))

    p17_lead = doc.add_paragraph()
    p17_lead.paragraph_format.space_before = Pt(2)
    p17_lead.paragraph_format.space_after = Pt(2)
    render_rich_text(p17_lead, "Nhận định nào sau đây *không đúng*?")

    opts17 = [
        "A. Phản ứng hydrate hóa ethylene là phản ứng cộng.",
        "B. Trong giai đoạn (1) có sự phân cắt liên kết π.",
        "C. Trong giai đoạn (1) có sự hình thành liên kết σ.",
        "*D. Trong phân tử ethylene có 6 liên kết σ.*" if is_teacher else "D. Trong phân tử ethylene có 6 liên kết σ."
    ]
    render_options_no_table(doc, opts17, is_teacher=is_teacher)
    if is_teacher:
        sol17 = (
            "• Cấu tạo phân tử ethylene (H~2~C=CH~2~): gồm 4 liên kết đơn C-H (4 liên kết σ) và 1 liên kết đôi C=C (gồm 1 liên kết σ và 1 liên kết π). Tổng cộng trong phân tử ethylene chỉ có 5 liên kết σ, không phải 6 liên kết σ.\n"
            "• Phát biểu A, B, C đều đúng theo cơ chế electrophilic addition (A~E~).\n"
            "➔ *Đáp án đúng: D*"
        )
        add_solution_no_table(doc, sol17)

    # CÂU 18
    p18 = doc.add_paragraph()
    p18.paragraph_format.space_before = Pt(4)
    p18.paragraph_format.space_after = Pt(2)
    render_question_prompt(p18, "Câu 18: Tiến hành thí nghiệm theo các bước sau:\n• Bước 1: Cho khoảng 2mL dung dịch AgNO~3~ 1% vào ống nghiệm sạch.\n• Bước 2: Thêm từ từ từng giọt dung dịch ammonia 5%, lắc đều cho đến khi kết tủa tan hết. Dung dịch thu được là thuốc thử Tollens.\n• Bước 3: Thêm tiếp khoảng 2mL dung dịch glucose 2% lắc đều. Sau đó, ngâm ống nghiệm vào cốc thuỷ tinh chứa nước nóng trong vài phút.\nPhát biểu nào sau đây sai?")
    opts18 = [
        "A. Sản phẩm hữu cơ thu được sau bước 3 là ammonium gluconate.",
        "*B. Thí nghiệm trên chứng minh glucose có tính chất của polyalcohol.*" if is_teacher else "B. Thí nghiệm trên chứng minh glucose có tính chất của polyalcohol.",
        "C. Sau bước 3, có lớp bạc (silver) kim loại bám trên thành ống nghiệm.",
        "D. Trong phản ứng ở bước 3, glucose đóng vai trò là chất khử."
    ]
    render_options_no_table(doc, opts18, is_teacher=is_teacher)
    if is_teacher:
        sol18 = (
            "• Phản ứng tráng bạc của glucose chứng minh phân tử glucose có nhóm chức aldehyde (-CHO), có tính khử (bị oxi hóa thành ammonium gluconate và khử Ag^+^ thành Ag kim loại bám vào thành ống nghiệm).\n"
            "• Tính chất polyalcohol (chứa nhiều nhóm -OH kề nhau) của glucose được chứng minh bằng phản ứng tạo dung dịch xanh lam thẫm với Cu(OH)~2~ ở nhiệt độ thường, không phải bằng phản ứng tráng bạc.\n"
            "➔ *Đáp án đúng: B*"
        )
        add_solution_no_table(doc, sol18)

    # ==========================================
    # PHẦN II: TRẮC NGHIỆM ĐÚNG SAI
    # ==========================================
    add_section_header(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 19 đến câu 22. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # CÂU 19
    p19 = doc.add_paragraph()
    p19.paragraph_format.space_before = Pt(4)
    p19.paragraph_format.space_after = Pt(2)
    render_question_prompt(p19, "Câu 19: Methyl cinnamate là một ester có công thức phân tử C~10~H~10~O~2~ và có mùi thơm của dâu tây (strawberry) được sử dụng trong ngành công nghiệp hương liệu và nước hoa.")

    # Hình cấu trúc methyl cinnamate
    p19_img = doc.add_paragraph()
    p19_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p19_img.paragraph_format.space_before = Pt(2)
    p19_img.paragraph_format.space_after = Pt(2)
    img_c19 = os.path.join(IMG_DIR, "cau19_methyl_cinnamate.png")
    if os.path.exists(img_c19):
        p19_img.add_run().add_picture(img_c19, width=Cm(6.0))

    stmts19 = [
        ("a) Methyl cinnamate có đồng phân hình học.", "Đúng"),
        ("b) Methyl cinnamate có công thức cấu tạo là CH~3~COO-CH=CH-C~6~H~5~.", "Sai"),
        ("c) Methyl cinnamate phản ứng với NaOH với tỉ lệ 1 : 2.", "Sai"),
        ("d) Để điều chế 24,3 gam ester methyl cinnamate, người ta cho 29,6 gam cinnamic acid phản ứng với lượng dư methyl alcohol. Hiệu suất phản ứng ester hóa trong trường hợp này là 75%.", "Đúng")
    ]
    for s_text, s_ans in stmts19:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, s_text, ans=s_ans, is_teacher=is_teacher)

    if is_teacher:
        sol19 = (
            "• a) Đúng: Phân tử methyl cinnamate có liên kết đôi -CH=CH- liên kết với 2 nhóm khác nhau ở mỗi carbon: C thứ nhất liên kết với -H và -C~6~H~5~; C thứ hai liên kết với -H và -COOCH~3~ nên tồn tại đồng phân hình học cis - trans.\n"
            "• b) Sai: Công thức cấu tạo đúng của methyl cinnamate là C~6~H~5~-CH=CH-COOCH~3~.\n"
            "• c) Sai: Methyl cinnamate là ester của alcohol đơn chức (methanol), khi phản ứng với NaOH chỉ xảy ra theo tỉ lệ mol 1 : 1 (C~6~H~5~-CH=CH-COOCH~3~ + NaOH → C~6~H~5~-CH=CH-COONa + CH~3~OH).\n"
            "• d) Đúng: Phương trình: C~6~H~5~-CH=CH-COOH (M = 148) + CH~3~OH ⇌ C~6~H~5~-CH=CH-COOCH~3~ (M = 162) + H~2~O.\n"
            "  n(acid) = 29,6 / 148 = 0,2 mol ➔ m(ester lý thuyết) = 0,2 × 162 = 32,4 gam.\n"
            "  Hiệu suất: H = (24,3 / 32,4) × 100% = 75%."
        )
        add_solution_no_table(doc, sol19)

    # CÂU 20
    p20 = doc.add_paragraph()
    p20.paragraph_format.space_before = Pt(4)
    p20.paragraph_format.space_after = Pt(2)
    render_question_prompt(p20, "Câu 20: Tiến hành thí nghiệm điều chế ethyl acetate theo các bước sau đây:\n• Bước 1: Cho 1mL C~2~H~5~OH, 1mL CH~3~COOH và vài giọt dung dịch H~2~SO~4~ đặc vào ống nghiệm.\n• Bước 2: Lắc đều ống nghiệm, đun cách thủy (trong nồi nước nóng) khoảng 5 - 6 phút ở 60 - 70°C.\n• Bước 3: Làm lạnh, sau đó rót 2mL dung dịch NaCl bão hòa vào ống nghiệm.")

    stmts20 = [
        ("a) Ở bước 1, có thể thay bằng dung dịch H~3~PO~4~ đậm đặc.", "Đúng"),
        ("b) Mục đích chính của việc thêm dung dịch NaCl bão hòa là để tránh phân hủy sản phẩm.", "Sai"),
        ("c) Sau bước 3, chất lỏng trong ống nghiệm tách thành hai lớp.", "Đúng"),
        ("d) H~2~SO~4~ đặc có vai trò vừa làm chất xúc tác vừa làm tăng hiệu suất tạo sản phẩm.", "Đúng")
    ]
    for s_text, s_ans in stmts20:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, s_text, ans=s_ans, is_teacher=is_teacher)

    if is_teacher:
        sol20 = (
            "• a) Đúng: H~3~PO~4~ đậm đặc là acid vô cơ mạnh có tính hút nước, có thể dùng làm xúc tác acid cho phản ứng ester hóa thay cho H~2~SO~4~ đặc (thậm chí an toàn hơn vì không gây hiện tượng than hóa).\n"
            "• b) Sai: Thêm dung dịch NaCl bão hòa nhằm tăng khối lượng riêng của lớp chất lỏng bên dưới và làm giảm độ tan của ethyl acetate, giúp ester nổi lên và tách lớp rõ rệt hơn.\n"
            "• c) Đúng: Ethyl acetate ít tan trong nước, nhẹ hơn dung dịch NaCl nên nổi lên trên tạo thành hai lớp chất lỏng phân cách rõ rệt.\n"
            "• d) Đúng: H~2~SO~4~ đặc cung cấp proton H^+^ làm xúc tác phản ứng ester hóa, đồng thời hút nước sinh ra làm chuyển dịch cân bằng theo chiều thuận, nâng cao hiệu suất."
        )
        add_solution_no_table(doc, sol20)

    # CÂU 21
    p21 = doc.add_paragraph()
    p21.paragraph_format.space_before = Pt(4)
    p21.paragraph_format.space_after = Pt(2)
    render_question_prompt(p21, "Câu 21: Hợp chất hữu cơ mạch hở X (C~7~H~10~O~5~) tác dụng với lượng dư dung dịch NaOH đun nóng thu được glycerol và hỗn hợp 2 muối carboxylate Y và Z (M~Y~ < M~Z~).")

    stmts21 = [
        ("a) Y tham gia phản ứng tráng bạc.", "Đúng"),
        ("b) Có 2 công thức cấu tạo thỏa mãn tính chất của X.", "Sai"),
        ("c) X là hợp chất hữu cơ tạp chức.", "Đúng"),
        ("d) Carboxylic acid của muối Z có đồng phân hình học.", "Sai")
    ]
    for s_text, s_ans in stmts21:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, s_text, ans=s_ans, is_teacher=is_teacher)

    if is_teacher:
        sol21 = (
            "• Phân tích cấu tạo chất X (C~7~H~10~O~5~):\n"
            "  + Độ bất bão hòa: k = (2×7 + 2 - 10)/2 = 3.\n"
            "  + X tác dụng NaOH sinh ra glycerol (C~3~H~5~(OH)~3~) và 2 muối Y, Z. Do X có 5 nguyên tử O nên X là ester 2 chức của glycerol, còn lại 1 nhóm -OH tự do (2 nhóm -COO- chứa 4O và 1 nhóm -OH chứa 1O).\n"
            "  + Tổng số C trong 2 gốc acid = 7 - 3 = 4C. Vì sinh ra 2 muối khác nhau có M~Y~ < M~Z~ và k = 3 (gồm 2 liên kết π trong 2 nhóm -COO- và 1 liên kết π trong gốc acid) nên:\n"
            "    * Muối Y là HCOONa (M~Y~ = 68).\n"
            "    * Muối Z là CH~2~=CH-COONa (M~Z~ = 94).\n"
            "• a) Đúng: Muối Y (HCOONa) là muối formate có nhóm -CHO nên tham gia phản ứng tráng bạc.\n"
            "• b) Sai: Khung carbon glycerol có 3 vị trí (1, 2, 3), trong đó C1 và C3 đối xứng. Các cấu tạo thỏa mãn gồm 3 CTCT:\n"
            "  (1) CH~2~(OOCH)-CH(OOCCH=CH~2~)-CH~2~OH\n"
            "  (2) CH~2~(OOCCH=CH~2~)-CH(OOCH)-CH~2~OH\n"
            "  (3) CH~2~(OOCH)-CH(OH)-CH~2~(OOCCH=CH~2~)\n"
            "• c) Đúng: X chứa đồng thời nhóm chức ester (-COO-) và nhóm alcohol (-OH) nên là hợp chất hữu cơ tạp chức.\n"
            "• d) Sai: Carboxylic acid của muối Z là CH~2~=CH-COOH (acrylic acid), có nguyên tử carbon ở liên kết đôi liên kết với 2 nguyên tử H giống nhau nên không có đồng phân hình học."
        )
        add_solution_no_table(doc, sol21)

    # CÂU 22
    p22 = doc.add_paragraph()
    p22.paragraph_format.space_before = Pt(4)
    p22.paragraph_format.space_after = Pt(2)
    render_question_prompt(p22, "Câu 22: Thực vật sử dụng quá trình quang hợp để chuyển đổi năng lượng ánh sáng từ mặt trời thành năng lượng hóa học. Trong quá trình quang hợp xảy ra phản ứng giữa khí carbonic và nước theo phương trình trình hóa học:\n6CO~2~(g) + 6H~2~O(l) → C~6~H~12~O~6~(aq) + 6O~2~(g)\nCho enthalpy tạo thành chuẩn của CO~2~(g) = -393,5 kJ/mol; H~2~O(l) = -285,8 kJ/mol; C~6~H~12~O~6~(aq) = -1271 kJ/mol.")

    stmts22 = [
        ("a) Phản ứng quang hợp của cây xanh là phản ứng tỏa nhiệt.", "Sai"),
        ("b) Nhiệt tạo thành chuẩn của O~2~(g) bằng 0.", "Đúng"),
        ("c) Biến thiên enthalpy chuẩn của phản ứng trên có giá trị là -2804,8 kJ.", "Sai"),
        ("d) Trồng nhiều cây xanh sẽ làm giảm được nhiệt độ của môi trường xung quanh.", "Đúng")
    ]
    for s_text, s_ans in stmts22:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, s_text, ans=s_ans, is_teacher=is_teacher)

    if is_teacher:
        sol22 = (
            "• a) Sai: Quá trình quang hợp hấp thụ năng lượng mặt trời (thu nhiệt) để tổng hợp glucose, Δ~r~H°~298~ > 0.\n"
            "• b) Đúng: Theo quy ước nhiệt động hóa học, nhiệt tạo thành chuẩn của đơn chất bền nhất ở điều kiện chuẩn (như O~2~(g)) bằng 0 kJ/mol.\n"
            "• c) Sai: Tính biến thiên enthalpy chuẩn của phản ứng:\n"
            "  Δ~r~H°~298~ = Δ~f~H°~298~(C~6~H~12~O~6~) + 6×Δ~f~H°~298~(O~2~) - [6×Δ~f~H°~298~(CO~2~) + 6×Δ~f~H°~298~(H~2~O)]\n"
            "  Δ~r~H°~298~ = -1271 + 0 - [6×(-393,5) + 6×(-285,8)] = -1271 - (-4075,8) = +2804,8 kJ.\n"
            "  (Giá trị mang dấu dương biểu thị phản ứng thu nhiệt, đề cho mang dấu âm là sai).\n"
            "• d) Đúng: Phản ứng quang hợp thu nhiệt từ môi trường cùng với quá trình thoát hơi nước qua khí khổng giúp điều hòa và làm mát không khí xung quanh."
        )
        add_solution_no_table(doc, sol22)

    # ==========================================
    # PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN
    # ==========================================
    add_section_header(
        doc,
        "PHẦN III. Câu trắc nghiệm trả lời ngắn.",
        "Thí sinh trả lời từ câu 23 đến câu 28."
    )

    # CÂU 23
    p23 = doc.add_paragraph()
    p23.paragraph_format.space_before = Pt(4)
    p23.paragraph_format.space_after = Pt(2)
    render_question_prompt(p23, "Câu 23: Có các phát biểu sau:\n(1) Thủy phân chất béo trong môi trường kiềm luôn thu được glycerol.\n(2) Triolein làm mất màu nước bromine.\n(3) Các ester chỉ được điều chế từ carboxylic acid và alcohol.\n(4) Benzyl acetate là ester có mùi chuối chín.\n(5) Đốt cháy ethyl acetate thu được số mol nước bằng số mol khí carbonic.\n(6) Hydrogen hoá hoàn toàn chất béo lỏng (xúc tác Ni, t°) thu được chất béo rắn.\nLiệt kê theo trình tự từ nhỏ đến lớn các số ứng với các phát biểu đúng trên (ví dụ: 56, 124, 3456,...).")

    p23_ans = doc.add_paragraph()
    p23_ans.paragraph_format.space_before = Pt(2)
    p23_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p23_ans, "► *Đáp án:* *1256*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol23 = (
            "• (1) Đúng: Chất béo là triester của glycerol và các acid béo nên thủy phân trong môi trường kiềm (xà phòng hóa) luôn thu được glycerol.\n"
            "• (2) Đúng: Triolein ((C~17~H~33~COO)~3~C~3~H~5~) chứa 3 liên kết đôi C=C nên phản ứng cộng làm mất màu nước bromine.\n"
            "• (3) Sai: Nhiều ester được điều chế bằng con đường khác, ví dụ vinyl acetate điều chế từ acetic acid tác dụng với acetylene (CH~3~COOH + CH≡CH), ester của phenol điều chế từ acid anhydride hoặc acid chloride tác dụng với phenol...\n"
            "• (4) Sai: Benzyl acetate có mùi thơm của hoa nhài; isoamyl acetate mới có mùi thơm của chuối chín.\n"
            "• (5) Đúng: Ethyl acetate (CH~3~COOC~2~H~5~ - C~4~H~8~O~2~) là ester no đơn chức mạch hở (C~n~H~2n~O~2~), khi đốt cháy luôn cho n(CO~2~) = n(H~2~O).\n"
            "• (6) Đúng: Phản ứng hydrogen hóa chuyển chất béo lỏng không no thành chất béo no rắn (margarine nhân tạo).\n"
            "➔ Các phát biểu đúng gồm: (1), (2), (5), (6). Dãy số theo thứ tự: *1256*"
        )
        add_solution_no_table(doc, sol23)
    else:
        render_rich_text(p23_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # CÂU 24
    p24 = doc.add_paragraph()
    p24.paragraph_format.space_before = Pt(4)
    p24.paragraph_format.space_after = Pt(2)
    render_question_prompt(p24, "Câu 24: Cho các phản ứng được đánh số thứ tự như sau:\n(1) 6nCO~2~ + 5nH~2~O (ánh sáng) → (C~6~H~10~O~5~)~n~ + 6nO~2~\n(2) (C~6~H~10~O~5~)~n~ + nH~2~O (enzyme) → nC~6~H~12~O~6~\n(3) C~6~H~12~O~6~ (enzyme) → 2C~2~H~5~OH + 2CO~2~\n(4) C~2~H~5~OH + 3O~2~ → 2CO~2~ + 3H~2~O\nGán số thứ tự phương trình hoá học của các phản ứng theo tên gọi: lên men rượu, thủy phân, cháy, quang hợp và sắp xếp theo trình tự thành dãy bốn số.")

    p24_ans = doc.add_paragraph()
    p24_ans.paragraph_format.space_before = Pt(2)
    p24_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p24_ans, "► *Đáp án:* *3241*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol24 = (
            "• Phân loại các phản ứng hóa học:\n"
            "  + Lên men rượu: Phản ứng (3) (C~6~H~12~O~6~ → 2C~2~H~5~OH + 2CO~2~).\n"
            "  + Thủy phân: Phản ứng (2) ((C~6~H~10~O~5~)~n~ + nH~2~O → nC~6~H~12~O~6~).\n"
            "  + Cháy: Phản ứng (4) (C~2~H~5~OH + 3O~2~ → 2CO~2~ + 3H~2~O).\n"
            "  + Quang hợp: Phản ứng (1) (6nCO~2~ + 5nH~2~O → (C~6~H~10~O~5~)~n~ + 6nO~2~).\n"
            "➔ Thứ tự các phản ứng: lên men rượu (3) - thủy phân (2) - cháy (4) - quang hợp (1).\n"
            "➔ Dãy bốn số cần điền là: *3241*"
        )
        add_solution_no_table(doc, sol24)
    else:
        render_rich_text(p24_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # CÂU 25
    p25 = doc.add_paragraph()
    p25.paragraph_format.space_before = Pt(4)
    p25.paragraph_format.space_after = Pt(2)
    render_question_prompt(p25, "Câu 25: Một nhà máy đường trung bình mỗi ngày sử dụng 3000 tấn mía để sản xuất. Sau khi nước mía được cô đặc và kết tinh để lấy đường cát, phần chất lỏng màu nâu sẫm, đặc sệt còn lại chính là rỉ đường. Tỷ lệ rỉ đường bằng 4% khối lượng mía ban đầu. Thành phần chính của rỉ đường gồm: Saccharose (khoảng 40%), monosaccharide (tổng glucose và fructose khoảng 15%), nước và các khoáng chất. Nhằm tận dụng nguồn nguyên liệu, nhà máy dùng rỉ đường để sản xuất ethanol, người ta pha loãng rỉ đường và tiến hành lên men bằng nấm men (Saccharomyces cerevisiae). Quá trình gồm thuỷ phân saccharose và lên men monosaccharide (cả glucose và fructose đều lên men tạo ethanol). Hiệu suất thuỷ phân Saccharose và lên men đều đạt 90%, các thành phần khác không tham gia quá trình tạo ethanol. Thể tích dung dịch ethanol 88° trung bình nhà máy sản xuất được trong 1 ngày là V (m^3^), biết khối lượng riêng của ethanol nguyên chất là 0,79 g/mL. Giá trị của V bằng bao nhiêu? (Không làm tròn các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng đơn vị).")

    p25_ans = doc.add_paragraph()
    p25_ans.paragraph_format.space_before = Pt(2)
    p25_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p25_ans, "► *Đáp án:* *42*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol25 = (
            "• Bước 1: Tính khối lượng các thành phần trong rỉ đường thu được trong 1 ngày:\n"
            "  m(rỉ đường) = 3000 tấn × 4% = 120 tấn.\n"
            "  m(saccharose) = 120 × 40% = 48 tấn.\n"
            "  m(monosaccharide có sẵn) = 120 × 15% = 18 tấn.\n"
            "• Bước 2: Thủy phân saccharose:\n"
            "  C~12~H~22~O~11~ + H~2~O → C~6~H~12~O~6~ (glucose) + C~6~H~12~O~6~ (fructose)\n"
            "  n(saccharose) = 48 / 342 tấn mol.\n"
            "  Do hiệu suất thủy phân đạt 90% (0,9):\n"
            "  n(mono sinh ra) = 2 × n(saccharose) × 0,9 = 2 × (48 / 342) × 0,9 = 86,4 / 342 tấn mol.\n"
            "• Bước 3: Tổng lượng monosaccharide đem lên men:\n"
            "  n(mono có sẵn) = 18 / 180 = 0,1 tấn mol.\n"
            "  n(mono tổng) = 0,1 + 86,4 / 342 = 0,35263158 tấn mol.\n"
            "• Bước 4: Lên men tạo ethanol (hiệu suất 90%):\n"
            "  C~6~H~12~O~6~ → 2C~2~H~5~OH + 2CO~2~\n"
            "  n(C~2~H~5~OH) = 2 × n(mono tổng) × 0,9 = 1,8 × (0,1 + 86,4 / 342) ≈ 0,63473684 tấn mol.\n"
            "• Bước 5: Tính khối lượng và thể tích dung dịch ethanol 88°:\n"
            "  m(C~2~H~5~OH nguyên chất) = 0,63473684 × 46 = 29,1978947 tấn = 29.197.894,7 gam.\n"
            "  V(C~2~H~5~OH nguyên chất) = m / D = 29.197.894,7 / 0,79 = 36.959.360,4 mL = 36,95936 m^3^.\n"
            "  V(dung dịch ethanol 88°) = V(nguyên chất) × 100 / 88 = 36,95936 × 100 / 88 = 41,9993 m^3^ ≈ 42 m^3^.\n"
            "➔ *Đáp án: 42*"
        )
        add_solution_no_table(doc, sol25)
    else:
        render_rich_text(p25_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # CÂU 26
    p26 = doc.add_paragraph()
    p26.paragraph_format.space_before = Pt(4)
    p26.paragraph_format.space_after = Pt(2)
    render_question_prompt(p26, "Câu 26: Xác định giá trị số sóng ν (cm^-1^) hấp thụ đặc trưng của nhóm C=O (ketone) trên phổ hồng ngoại IR của methyl phenyl ketone cho dưới đây:")

    # Hình phổ hồng ngoại IR
    p26_img = doc.add_paragraph()
    p26_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p26_img.paragraph_format.space_before = Pt(2)
    p26_img.paragraph_format.space_after = Pt(2)
    img_c26 = os.path.join(IMG_DIR, "cau26_pho_ir.png")
    if os.path.exists(img_c26):
        p26_img.add_run().add_picture(img_c26, width=Cm(11.0))

    p26_ans = doc.add_paragraph()
    p26_ans.paragraph_format.space_before = Pt(2)
    p26_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p26_ans, "► *Đáp án:* *1686*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol26 = (
            "• Trên phổ hồng ngoại (IR), tín hiệu hấp thụ đặc trưng của liên kết đôi C=O (nhóm carbonyl) là mũi hấp thụ có cường độ rất mạnh, nhọn và sâu nhất nằm trong vùng 1670 - 1750 cm^-1^.\n"
            "• Đối với hợp chất methyl phenyl ketone (acetophenone), do nhóm C=O liên hợp với hệ electron π của vòng benzene nên số sóng hấp thụ dịch chuyển về vùng tần số thấp hơn một chút, thể hiện rõ pic sâu nhất tại số sóng 1686 cm^-1^.\n"
            "➔ *Đáp án: 1686*"
        )
        add_solution_no_table(doc, sol26)
    else:
        render_rich_text(p26_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # CÂU 27
    p27 = doc.add_paragraph()
    p27.paragraph_format.space_before = Pt(4)
    p27.paragraph_format.space_after = Pt(2)
    render_question_prompt(p27, "Câu 27: Chỉ số xà phòng hoá của chất béo là số mg KOH cần dùng để xà phòng hoá hoàn toàn 1 gam chất béo (bao gồm cả trung hoà acid béo tự do có trong chất béo). Cho bảng số liệu:")

    # Bảng số liệu Câu 27
    tbl27 = doc.add_table(rows=5, cols=4)
    tbl27.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl27.autofit = False

    tbl27_data = [
        ["Chất béo", "Chỉ số xà phòng hoá", "Chất béo", "Chỉ số xà phòng hoá"],
        ["Dầu dừa", "257", "Dầu olive", "188"],
        ["Dầu cọ", "199", "Mỡ vịt", "194"],
        ["Dầu lạc", "192", "Mỡ gà", "195"],
        ["Dầu mè", "188", "Mỡ lợn", "198"]
    ]

    col_w27_cb = Cm(3.5)
    col_w27_val = Cm(3.5)

    for r_idx, r_vals in enumerate(tbl27_data):
        row = tbl27.rows[r_idx]
        for c_idx, val in enumerate(r_vals):
            cell = row.cells[c_idx]
            cell.width = col_w27_cb if c_idx in [0, 2] else col_w27_val
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                set_cell_shading(cell, "FFFFFF" if c_idx in [1, 3] else "FAFAFA")
                is_bold = (val == "199" or val == "Dầu cọ")
                render_rich_text(p, f"*{val}*" if is_bold else val, base_size=10, base_bold=is_bold, base_color=COLOR_CORRECT if is_bold else None)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    p27_tail = doc.add_paragraph()
    p27_tail.paragraph_format.space_before = Pt(3)
    p27_tail.paragraph_format.space_after = Pt(2)
    render_rich_text(p27_tail, "Giả sử quá trình thực hành điều chế xà phòng, học sinh được yêu cầu điều chế xà phòng từ 200 gam dầu cọ và NaOH (rắn). Biết xà phòng được điều chế bằng cách cho chất béo tác dụng vừa đủ với dung dịch NaOH. Cần cân m (gam) hoá chất từ lọ đựng NaOH rắn trong phòng thí nghiệm, biết rằng loại NaOH được dùng trong phòng thí nghiệm đã hút ẩm nên lượng NaOH chiếm 95%. Giá trị của m bằng bao nhiêu? (Các quá trình tính toán không làm tròn, kết quả cuối cùng làm tròn đến hàng phần mười).", base_size=10.5)

    p27_ans = doc.add_paragraph()
    p27_ans.paragraph_format.space_before = Pt(2)
    p27_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p27_ans, "► *Đáp án:* *29,9*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol27 = (
            "• Tra bảng: Chỉ số xà phòng hoá của dầu cọ là 199 (cần 199 mg KOH để xà phòng hoá 1 gam dầu cọ).\n"
            "• Khối lượng KOH cần dùng để xà phòng hoá 200 gam dầu cọ là:\n"
            "  m(KOH) = 200 × 199 mg = 39800 mg = 39,8 gam.\n"
            "• Số mol KOH tương ứng:\n"
            "  n(KOH) = 39,8 / 56 mol.\n"
            "• Vì phản ứng xà phòng hoá của NaOH và KOH có tỉ lệ mol như nhau (cùng là base đơn hóa trị):\n"
            "  n(NaOH nguyên chất) = n(KOH) = 39,8 / 56 mol.\n"
            "• Khối lượng NaOH nguyên chất cần dùng:\n"
            "  m(NaOH nguyên chất) = (39,8 / 56) × 40 = 39,8 × (5 / 7) ≈ 28,42857 gam.\n"
            "• Do hóa chất NaOH rắn trong phòng thí nghiệm đã hút ẩm nên NaOH chỉ chiếm 95% khối lượng:\n"
            "  m(hóa chất cần cân) = m(NaOH nguyên chất) / 0,95 = 28,42857 / 0,95 ≈ 29,9248 gam.\n"
            "• Làm tròn đến hàng phần mười theo yêu cầu đề bài: m = 29,9 gam.\n"
            "➔ *Đáp án: 29,9*"
        )
        add_solution_no_table(doc, sol27)
    else:
        render_rich_text(p27_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # CÂU 28
    p28 = doc.add_paragraph()
    p28.paragraph_format.space_before = Pt(4)
    p28.paragraph_format.space_after = Pt(2)
    render_question_prompt(p28, "Câu 28: Cho các chất: CH~3~[CH~2~]~14~COONa, CH~3~[CH~2~]~10~CH~2~OSO~3~Na, CH~3~[CH~2~]~16~COOK, CH~3~[CH~2~]~7~CH=CH[CH~2~]~7~COONa, CH~3~COONa, (C~15~H~31~COO)~3~C~3~H~5~. Có bao nhiêu chất có thể là thành phần chính của xà phòng?")

    p28_ans = doc.add_paragraph()
    p28_ans.paragraph_format.space_before = Pt(2)
    p28_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p28_ans, "► *Đáp án:* *3*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol28 = (
            "• Xà phòng là hỗn hợp muối sodium hoặc potassium của các acid béo (thường có mạch hydrocarbon dài từ 12 đến 18 nguyên tử C):\n"
            "  (1) CH~3~[CH~2~]~14~COONa (sodium palmitate - muối của acid béo no C16) ➔ Là thành phần của xà phòng.\n"
            "  (2) CH~3~[CH~2~]~10~CH~2~OSO~3~Na (sodium lauryl sulfate - muối alkyl sulfate) ➔ Là chất giặt rửa tổng hợp, không phải xà phòng.\n"
            "  (3) CH~3~[CH~2~]~16~COOK (potassium stearate - muối của acid béo no C18) ➔ Là thành phần của xà phòng mềm/lỏng.\n"
            "  (4) CH~3~[CH~2~]~7~CH=CH[CH~2~]~7~COONa (sodium oleate - muối của acid béo không no C18) ➔ Là thành phần của xà phòng.\n"
            "  (5) CH~3~COONa (sodium acetate) ➔ Muối của acid mạch ngắn, không có tính năng tạo bọt và giặt rửa.\n"
            "  (6) (C~15~H~31~COO)~3~C~3~H~5~ (tripalmitin) ➔ Chất béo (nguyên liệu sản xuất xà phòng, không phải thành phần xà phòng).\n"
            "• Vậy có đúng 3 chất thỏa mãn là: CH~3~[CH~2~]~14~COONa, CH~3~[CH~2~]~16~COOK, CH~3~[CH~2~]~7~CH=CH[CH~2~]~7~COONa.\n"
            "➔ *Đáp án: 3*"
        )
        add_solution_no_table(doc, sol28)
    else:
        render_rich_text(p28_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(12)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "*------------------- HẾT -------------------*", base_size=11, base_bold=True)

    p_subend = doc.add_paragraph()
    p_subend.paragraph_format.space_before = Pt(2)
    p_subend.paragraph_format.space_after = Pt(2)
    p_subend.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_subend, "_(Học sinh không được sử dụng tài liệu. Cán bộ coi thi không giải thích gì thêm)_", base_size=9.5, base_italic=True)

    suffix = "_GiaoVien.docx" if is_teacher else ".docx"
    filename = f"De_Kiem_Tra_Hoa_12_BacNinh_Ma013{suffix}"
    out_path = os.path.join(OUTPUT_DIR, filename)
    doc.save(out_path)
    print(f"Đã tạo thành công file: {out_path}")
    return out_path

if __name__ == "__main__":
    build_exam_bacninh_013(is_teacher=False)
    build_exam_bacninh_013(is_teacher=True)
