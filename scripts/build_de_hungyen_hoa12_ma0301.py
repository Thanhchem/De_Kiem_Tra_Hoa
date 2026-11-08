# -*- coding: utf-8 -*-
"""
Script tạo Đề Khảo Sát Chất Lượng Giữa Học Kỳ I Lớp 12 (Năm Học 2026 - 2027)
Sở Giáo Dục và Đào Tạo Hưng Yên - Môn: Hoá Học - Mã Đề 0301
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)

Xuất bản 2 file Word:
1. Bản Học sinh: De_Kiem_Tra_Hoa_12_HungYen_Ma0301.docx
2. Bản Giáo viên: De_Kiem_Tra_Hoa_12_HungYen_Ma0301_GiaoVien.docx
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
IMG_DIR = os.path.join(OUTPUT_DIR, "images_hungyen_12_ma0301")

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

def add_quick_answer_table_0301(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu Mã đề 0301.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 0301)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (18 CÂU)*", base_size=11, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=4, cols=10)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_p1_1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9"]
    answers_p1_1 = ["Đ/A", "C", "A", "D", "B", "A", "B", "B", "D", "C"]
    headers_p1_2 = ["Câu", "10", "11", "12", "13", "14", "15", "16", "17", "18"]
    answers_p1_2 = ["Đ/A", "B", "A", "A", "A", "D", "B", "D", "C", "C"]

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
        ["Câu 1", "Sai", "Đúng", "Sai", "Đúng"],
        ["Câu 2", "Đúng", "Sai", "Đúng", "Đúng"],
        ["Câu 3", "Sai", "Sai", "Sai", "Đúng"],
        ["Câu 4", "Đúng", "Đúng", "Đúng", "Sai"]
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

    tbl3_headers = ["Câu", "Câu 1", "Câu 2", "Câu 3", "Câu 4", "Câu 5", "Câu 6"]
    tbl3_answers = ["Đáp án", "134", "196", "2314", "1", "297", "3"]

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


def build_exam_hungyen_0301(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="SỞ GIÁO DỤC VÀ ĐÀO TẠO HƯNG YÊN",
        exam_title="KHẢO SÁT CHẤT LƯỢNG GIỮA HỌC KỲ I LỚP 12 - NĂM HỌC 2026-2027",
        subject_line="MÔN: HOÁ HỌC - LỚP 12",
        num_pages="04",
        duration="50 phút, không kể thời gian phát đề",
        exam_code="0301",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Cl = 35,5; S = 32; Na = 23; K = 39; Ca = 40; Fe = 56",
        is_teacher=is_teacher
    )

    # Thêm bảng thông số phổ hồng ngoại (IR) đã cho trong đề
    p_ir_head = doc.add_paragraph()
    p_ir_head.paragraph_format.space_before = Pt(2)
    p_ir_head.paragraph_format.space_after = Pt(2)
    render_rich_text(p_ir_head, "_(Cho biết số sóng hấp thụ đặc trưng của một số liên kết trên phổ hồng ngoại như sau: )_", base_size=9.5, base_italic=True)

    tbl_ir = doc.add_table(rows=2, cols=4)
    tbl_ir.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_ir.autofit = False

    ir_headers = ["Liên kết", "O–H (alcohol)", "O–H (carboxylic acid)", "C=O (ester, carboxylic acid)", "C–O (ester)"]
    ir_vals = ["Số sóng (cm^-1^)", "3650 – 3200", "3300 – 2500", "1780 – 1650", "1300 – 1000"]

    tbl_ir = doc.add_table(rows=2, cols=5)
    tbl_ir.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_ir.autofit = False

    col_ir_0 = Cm(3.2)
    col_ir = Cm(3.7)

    for r_idx, r_data in enumerate([ir_headers, ir_vals]):
        row = tbl_ir.rows[r_idx]
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = col_ir_0 if c_idx == 0 else col_ir
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.space_after = Pt(1.5)

            is_header_row = (r_idx == 0)
            if is_header_row:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=9.5, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                set_cell_shading(cell, "FFFFFF")
                render_rich_text(p, val, base_size=9.5)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    p_div2 = doc.add_paragraph()
    p_div2.paragraph_format.space_before = Pt(2)
    p_div2.paragraph_format.space_after = Pt(4)
    add_p_border_bottom(p_div2, color="8BC390", sz="4")

    if is_teacher:
        add_quick_answer_table_0301(doc)

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
    render_question_prompt(p1, "Câu 1: Cho cấu tạo dạng mạch vòng của saccharose và maltose như sau:")

    p1_img = doc.add_paragraph()
    p1_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1_img.paragraph_format.space_before = Pt(2)
    p1_img.paragraph_format.space_after = Pt(2)
    img_c1 = os.path.join(IMG_DIR, "cau1_saccharose_maltose.png")
    if os.path.exists(img_c1):
        p1_img.add_run().add_picture(img_c1, width=Cm(11.0))

    p1_lead = doc.add_paragraph()
    p1_lead.paragraph_format.space_before = Pt(2)
    p1_lead.paragraph_format.space_after = Pt(2)
    render_rich_text(p1_lead, "Nhận định nào sau đây đúng?")

    opts1 = [
        "A. Saccharose và maltose đều có 7 nhóm OH.",
        "B. Maltose và saccharose đều có dạng mở vòng.",
        "*C. Nhóm OH (X) trong công thức maltose là nhóm OH hemiacetal.*" if is_teacher else "C. Nhóm OH (X) trong công thức maltose là nhóm OH hemiacetal.",
        "D. Nhóm OH (Y) trong công thức saccharose là nhóm OH hemiketal."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        sol1 = (
            "• Trong phân tử maltose, gốc glucose bên phải vẫn còn nguyên nhóm -OH ở vị trí carbon C1 anomeric, đây chính là nhóm -OH hemiacetal (được ký hiệu là X). Nhóm này giúp gốc glucose mở vòng để tạo nhóm aldehyde -CHO, tham gia phản ứng tráng bạc.\n"
            "• Saccharose và maltose (C~12~H~22~O~11~) đều có 8 nhóm -OH (không phải 7).\n"
            "• Saccharose được tạo bởi liên kết α-1,2-glycosidic giữa C1 (glucose) và C2 (fructose), không còn nhóm -OH hemiacetal hay hemiketal tự do nên không thể mở vòng.\n"
            "➔ *Đáp án đúng: C*"
        )
        add_solution_no_table(doc, sol1)

    # CÂU 2
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(4)
    p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p2, "Câu 2: Thuỷ phân hoàn toàn saccharose thu được X và Y. Để phân biệt X và Y, người ta sử dụng")
    opts2 = [
        "*A. nước bromine.*" if is_teacher else "A. nước bromine.",
        "B. thuốc thử Tollens.",
        "C. Cu(OH)~2~/OH^-^ đun nóng.",
        "D. Cu(OH)~2~ ở nhiệt độ thường."
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        sol2 = (
            "• Thủy phân saccharose thu được X (glucose) và Y (fructose).\n"
            "• Để phân biệt glucose và fructose, người ta sử dụng nước bromine:\n"
            "  + Glucose làm mất màu nước bromine: CH~2~OH[CHOH]~4~CHO + Br~2~ + H~2~O → CH~2~OH[CHOH]~4~COOH + 2HBr\n"
            "  + Fructose không làm mất màu nước bromine (ở môi trường trung tính/acid không có sự chuyển hóa giữa fructose và glucose).\n"
            "• Các thuốc thử Tollens, Cu(OH)~2~ đều ở môi trường kiềm làm fructose đồng phân hóa thành glucose nên không dùng để phân biệt được.\n"
            "➔ *Đáp án đúng: A*"
        )
        add_solution_no_table(doc, sol2)

    # CÂU 3
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(4)
    p3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p3, "Câu 3: Cho bảng nhiệt độ sôi của các chất sau ở 25°C và 1 bar:")

    p3_img = doc.add_paragraph()
    p3_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3_img.paragraph_format.space_before = Pt(2)
    p3_img.paragraph_format.space_after = Pt(2)
    img_c3 = os.path.join(IMG_DIR, "cau3_bieu_do_ts.png")
    if os.path.exists(img_c3):
        p3_img.add_run().add_picture(img_c3, width=Cm(9.0))

    p3_lead = doc.add_paragraph()
    p3_lead.paragraph_format.space_before = Pt(2)
    p3_lead.paragraph_format.space_after = Pt(2)
    render_rich_text(p3_lead, "Có bao nhiêu chất ở trạng thái khí?")

    opts3 = [
        "A. 3.",
        "B. 1.",
        "C. 2.",
        "*D. 4.*" if is_teacher else "D. 4."
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        sol3 = (
            "• Một chất tồn tại ở trạng thái khí ở 25°C và 1 bar khi nhiệt độ sôi của nó nhỏ hơn 25°C (t~sôi~ < 25°C).\n"
            "• Nhìn biểu đồ nhiệt độ sôi:\n"
            "  + CH~3~NH~2~: -6,4°C < 25°C ➔ chất khí.\n"
            "  + CH~3~CH~2~NH~2~: 16,6°C < 25°C ➔ chất khí.\n"
            "  + CH~3~NHCH~3~: 7,3°C < 25°C ➔ chất khí.\n"
            "  + (CH~3~)~3~N: 2,8°C < 25°C ➔ chất khí.\n"
            "• Cả 4 chất amine đều có nhiệt độ sôi thấp hơn 25°C nên đều ở trạng thái khí ở điều kiện này.\n"
            "➔ *Đáp án đúng: D*"
        )
        add_solution_no_table(doc, sol3)

    # CÂU 4
    p4 = doc.add_paragraph()
    p4.paragraph_format.space_before = Pt(4)
    p4.paragraph_format.space_after = Pt(2)
    render_question_prompt(p4, "Câu 4: Một acid béo được gọi là omega-n (khi đó n là vị trí của liên kết đôi đầu tiên tính từ nhóm methyl). Cho công thức cấu tạo của một loại acid béo X như sau:")

    p4_img = doc.add_paragraph()
    p4_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4_img.paragraph_format.space_before = Pt(2)
    p4_img.paragraph_format.space_after = Pt(2)
    img_c4 = os.path.join(IMG_DIR, "cau4_acid_beo.png")
    if os.path.exists(img_c4):
        p4_img.add_run().add_picture(img_c4, width=Cm(12.0))

    p4_lead = doc.add_paragraph()
    p4_lead.paragraph_format.space_before = Pt(2)
    p4_lead.paragraph_format.space_after = Pt(2)
    render_rich_text(p4_lead, "Acid béo X thuộc loại omega nào sau đây?")

    opts4 = [
        "A. Omega-3.",
        "*B. Omega-9.*" if is_teacher else "B. Omega-9.",
        "C. Omega-6.",
        "D. Omega-12."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        sol4 = (
            "• Đếm số nguyên tử carbon bắt đầu từ nhóm methyl (đầu đối diện nhóm -COOH):\n"
            "  + Carbon số 1: Nhóm -CH~3~ ở mút trái (vị trí ω-1).\n"
            "  + Các liên kết đơn nối tiếp: C1 - C2 - C3 - C4 - C5 - C6 - C7 - C8 - C9.\n"
            "  + Liên kết đôi C=C đầu tiên xuất hiện ở vị trí giữa carbon C9 và C10 (C9=C10).\n"
            "• Do liên kết đôi đầu tiên bắt đầu ở vị trí carbon số 9 tính từ nhóm methyl nên X thuộc loại acid béo Omega-9 (oleic acid: cis-octadec-9-enoic acid).\n"
            "➔ *Đáp án đúng: B*"
        )
        add_solution_no_table(doc, sol4)

    # CÂU 5
    p5 = doc.add_paragraph()
    p5.paragraph_format.space_before = Pt(4)
    p5.paragraph_format.space_after = Pt(2)
    render_question_prompt(p5, "Câu 5: Cho sơ đồ của phản ứng: (C~17~H~33~COO)~3~C~3~H~5~ + 3NaOH ⎯⎯(t°)→ 3Y (muối) + C~3~H~5~(OH)~3~. Công thức của muối Y là")
    opts5 = [
        "*A. C~17~H~33~COONa.*" if is_teacher else "A. C~17~H~33~COONa.",
        "B. C~17~H~35~COONa.",
        "C. C~17~H~31~COONa.",
        "D. C~15~H~31~COONa."
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        sol5 = (
            "• Chất béo (C~17~H~33~COO)~3~C~3~H~5~ (triolein) khi xà phòng hóa trong dung dịch NaOH tạo thành glycerol và muối sodium oleate:\n"
            "  (C~17~H~33~COO)~3~C~3~H~5~ + 3NaOH ⎯⎯(t°)→ 3C~17~H~33~COONa + C~3~H~5~(OH)~3~\n"
            "• Vậy muối Y là C~17~H~33~COONa.\n"
            "➔ *Đáp án đúng: A*"
        )
        add_solution_no_table(doc, sol5)

    # CÂU 6
    p6 = doc.add_paragraph()
    p6.paragraph_format.space_before = Pt(4)
    p6.paragraph_format.space_after = Pt(2)
    render_question_prompt(p6, "Câu 6: Ở các giá trị pH khác nhau, valine có thể tồn tại chủ yếu ở các dạng khác nhau. Trong điện trường, các dạng này có khả năng dịch chuyển về các điện cực khác nhau. Cho sơ đồ biểu diễn các dạng tồn tại của valine dưới đây:")

    p6_img = doc.add_paragraph()
    p6_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p6_img.paragraph_format.space_before = Pt(2)
    p6_img.paragraph_format.space_after = Pt(2)
    img_c6 = os.path.join(IMG_DIR, "cau6_valine.png")
    if os.path.exists(img_c6):
        p6_img.add_run().add_picture(img_c6, width=Cm(12.0))

    p6_lead = doc.add_paragraph()
    p6_lead.paragraph_format.space_before = Pt(2)
    p6_lead.paragraph_format.space_after = Pt(2)
    render_rich_text(p6_lead, "Phát biểu nào sau đây là sai?")

    opts6 = [
        "A. Thứ tự giá trị pH là a < b < c.",
        "*B. Trong dung dịch nước, cả ba dạng (1), (2) và (3) đều có tính chất lưỡng tính.*" if is_teacher else "B. Trong dung dịch nước, cả ba dạng (1), (2) và (3) đều có tính chất lưỡng tính.",
        "C. Ở pH = a, khi đặt trong điện trường, dạng (1) sẽ dịch chuyển sang cực âm.",
        "D. Dạng (1) là cation, dạng (2) là ion lưỡng cực, dạng (3) là anion."
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        sol6 = (
            "• Phát biểu B sai vì trong dung dịch nước chỉ có dạng (2) (ion lưỡng cực - zwitterion) là có tính lưỡng tính vì vừa có nhóm -COO^-^ có khả năng nhận proton H^+^, vừa có nhóm -NH~3~^+^ có khả năng nhường proton H^+^. Dạng (1) chỉ thể hiện tính acid (nhường H^+^), dạng (3) chỉ thể hiện tính base (nhận H^+^).\n"
            "• Phát biểu A đúng vì dạng cation (1) bền trong môi trường acid mạnh (pH = a thấp nhất), dạng lưỡng cực (2) ở điểm đẳng điện (pH = b ≈ 6), dạng anion (3) bền trong môi trường kiềm (pH = c cao nhất) ➔ a < b < c.\n"
            "• Phát biểu C đúng vì dạng (1) mang điện tích dương +1 nên dịch chuyển về cực âm (cathode).\n"
            "➔ *Đáp án đúng: B*"
        )
        add_solution_no_table(doc, sol6)

    # CÂU 7
    p7 = doc.add_paragraph()
    p7.paragraph_format.space_before = Pt(4)
    p7.paragraph_format.space_after = Pt(2)
    render_question_prompt(p7, "Câu 7: Ester X được tạo bởi ethyl alcohol và acetic acid có công thức cấu tạo là")
    opts7 = [
        "A. CH~3~COOCH~3~.",
        "*B. CH~3~COOC~2~H~5~.*" if is_teacher else "B. CH~3~COOC~2~H~5~.",
        "C. HCOOCH~3~.",
        "D. C~2~H~5~COOC~2~H~5~."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        sol7 = (
            "• Phản ứng ester hóa: CH~3~COOH + C~2~H~5~OH ⇌ CH~3~COOC~2~H~5~ + H~2~O.\n"
            "• Ester X là ethyl acetate: CH~3~COOC~2~H~5~.\n"
            "➔ *Đáp án đúng: B*"
        )
        add_solution_no_table(doc, sol7)

    # CÂU 8
    p8 = doc.add_paragraph()
    p8.paragraph_format.space_before = Pt(4)
    p8.paragraph_format.space_after = Pt(2)
    render_question_prompt(p8, "Câu 8: Cho sơ đồ: Disaccharide ⎯⎯(+H~2~O, acid, t°)→ Hỗn hợp monosaccharide ⎯⎯(+[Ag(NH~3~)~2~]OH, t°)→ Bạc kim loại. Trong quá trình trên, từ 1 mol saccharose tạo được tối đa bao nhiêu mol bạc kim loại?")
    opts8 = [
        "A. 3 mol.",
        "B. 2 mol.",
        "C. 1 mol.",
        "*D. 4 mol.*" if is_teacher else "D. 4 mol."
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        sol8 = (
            "• 1 mol saccharose (C~12~H~22~O~11~) thủy phân tạo ra: 1 mol glucose + 1 mol fructose.\n"
            "• Cả glucose và fructose trong môi trường kiềm của thuốc thử Tollens đều tham gia phản ứng tráng bạc với tỉ lệ 1 mol monosaccharide tạo 2 mol Ag:\n"
            "  1 mol glucose ➔ 2 mol Ag\n"
            "  1 mol fructose ➔ 2 mol Ag\n"
            "• Tổng số mol Ag tối đa thu được = 2 + 2 = 4 mol Ag.\n"
            "➔ *Đáp án đúng: D*"
        )
        add_solution_no_table(doc, sol8)

    # CÂU 9
    p9 = doc.add_paragraph()
    p9.paragraph_format.space_before = Pt(4)
    p9.paragraph_format.space_after = Pt(2)
    render_question_prompt(p9, "Câu 9: Khi trộn dầu ăn với nước rồi để yên, người ta thấy dầu ăn tách thành một lớp nổi trên mặt nước. Nhận định nào sau đây phù hợp với hiện tượng trên?")
    opts9 = [
        "A. Dầu ăn tan trong nước và có khối lượng riêng nhỏ hơn nước.",
        "B. Dầu ăn không tan trong nước và có khối lượng riêng lớn hơn nước.",
        "*C. Dầu ăn không tan trong nước và có khối lượng riêng nhỏ hơn nước.*" if is_teacher else "C. Dầu ăn không tan trong nước và có khối lượng riêng nhỏ hơn nước.",
        "D. Dầu ăn tan trong nước và có khối lượng riêng lớn hơn nước."
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        sol9 = (
            "• Dầu ăn (chất béo) không phân cực nên không tan trong dung môi phân cực như nước, đồng thời khối lượng riêng của dầu ăn (D ≈ 0,8 – 0,9 g/mL) nhỏ hơn nước (D = 1,0 g/mL) nên dầu ăn nổi lên trên mặt nước và tách thành lớp riêng biệt.\n"
            "➔ *Đáp án đúng: C*"
        )
        add_solution_no_table(doc, sol9)

    # CÂU 10
    p10 = doc.add_paragraph()
    p10.paragraph_format.space_before = Pt(4)
    p10.paragraph_format.space_after = Pt(2)
    render_question_prompt(p10, "Câu 10: Chất nào sau đây là thành phần chủ yếu của xà phòng?")
    opts10 = [
        "A. CH~3~COOH.",
        "*B. CH~3~[CH~2~]~14~COONa.*" if is_teacher else "B. CH~3~[CH~2~]~14~COONa.",
        "C. CH~3~CH~2~OH.",
        "D. CH~3~[CH~2~]~16~COOH."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        sol10 = (
            "• Xà phòng là hỗn hợp các muối sodium hoặc potassium của acid béo (mạch C dài từ 12-18 C).\n"
            "• CH~3~[CH~2~]~14~COONa (sodium palmitate - C~15~H~31~COONa) là muối sodium của acid béo, là thành phần chủ yếu của xà phòng.\n"
            "➔ *Đáp án đúng: B*"
        )
        add_solution_no_table(doc, sol10)

    # CÂU 11
    p11 = doc.add_paragraph()
    p11.paragraph_format.space_before = Pt(4)
    p11.paragraph_format.space_after = Pt(2)
    render_question_prompt(p11, "Câu 11: Dung dịch nào sau đây không có khả năng hòa tan Cu(OH)~2~?")
    opts11 = [
        "*A. Dung dịch ethanol.*" if is_teacher else "A. Dung dịch ethanol.",
        "B. Nước mía.",
        "C. Nước ép quả nho chín.",
        "D. Nước ép củ cải đường."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        sol11 = (
            "• Nước mía và nước ép củ cải đường chứa hàm lượng lớn saccharose; nước ép quả nho chín chứa nhiều glucose. Cả hai đều là polyalcohol có nhiều nhóm -OH kề nhau nên hòa tan Cu(OH)~2~ tạo dung dịch phức màu xanh lam.\n"
            "• Dung dịch ethanol (C~2~H~5~OH) là alcohol đơn chức, không có khả năng hòa tan Cu(OH)~2~ ở nhiệt độ thường.\n"
            "➔ *Đáp án đúng: A*"
        )
        add_solution_no_table(doc, sol11)

    # CÂU 12
    p12 = doc.add_paragraph()
    p12.paragraph_format.space_before = Pt(4)
    p12.paragraph_format.space_after = Pt(2)
    render_question_prompt(p12, "Câu 12: Thông thường, nếu mạch carbon của chất giặt rửa tổng hợp không phân nhánh thì chất giặt rửa đó dễ phân huỷ sinh học hơn so với mạch carbon phân nhánh. Chất giặt rửa nào sau đây dễ phân huỷ sinh học nhất?")

    p12_img = doc.add_paragraph()
    p12_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p12_img.paragraph_format.space_before = Pt(2)
    p12_img.paragraph_format.space_after = Pt(2)
    img_c12 = os.path.join(IMG_DIR, "cau12_chat_giat_rua.png")
    if os.path.exists(img_c12):
        p12_img.add_run().add_picture(img_c12, width=Cm(13.0))

    opts12 = [
        "*A. Chất A.*" if is_teacher else "A. Chất A.",
        "B. Chất B.",
        "C. Chất C.",
        "D. Chất D."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        sol12 = (
            "• Chất giặt rửa ở phương án A có mạch carbon dài dạng mạch thẳng không phân nhánh (Linear Alkylbenzene Sulfonate - LAS). Nhờ cấu trúc không phân nhánh, vi sinh vật trong tự nhiên có thể dễ dàng tiếp cận và phân giải hoàn toàn chuỗi alkyl qua con đường β-oxi hóa, do đó dễ phân hủy sinh học nhất và thân thiện với môi trường.\n"
            "• Các chất ở B, C, D đều có mạch alkyl phân nhánh (chứa các nhóm methyl, isopropyl hoặc gem-dimethyl cản trở không gian) nên khó phân hủy sinh học hơn nhiều.\n"
            "➔ *Đáp án đúng: A*"
        )
        add_solution_no_table(doc, sol12)

    # CÂU 13
    p13 = doc.add_paragraph()
    p13.paragraph_format.space_before = Pt(4)
    p13.paragraph_format.space_after = Pt(2)
    render_question_prompt(p13, "Câu 13: Nhỏ dung dịch methylamine đến dư vào dung dịch nào sau đây thu được kết tủa sau khi kết thúc phản ứng?")
    opts13 = [
        "*A. FeCl~3~.*" if is_teacher else "A. FeCl~3~.",
        "B. HCl.",
        "C. NaOH.",
        "D. CuSO~4~."
    ]
    render_options_no_table(doc, opts13, is_teacher=is_teacher)
    if is_teacher:
        sol13 = (
            "• Khi nhỏ dung dịch methylamine dư vào FeCl~3~:\n"
            "  FeCl~3~ + 3CH~3~NH~2~ + 3H~2~O → Fe(OH)~3~↓ (nâu đỏ) + 3CH~3~NH~3~Cl\n"
            "  Kết tủa Fe(OH)~3~ không tan trong dung dịch methylamine dư nên thu được kết tủa bền sau phản ứng.\n"
            "• Đối với CuSO~4~, ban đầu xuất hiện kết tủa xanh nhạt Cu(OH)~2~, nhưng methylamine dư sẽ tạo phức chất tan màu xanh lam thẫm [Cu(CH~3~NH~2~)~4~](OH)~2~ làm kết tủa tan hết.\n"
            "➔ *Đáp án đúng: A*"
        )
        add_solution_no_table(doc, sol13)

    # CÂU 14
    p14 = doc.add_paragraph()
    p14.paragraph_format.space_before = Pt(4)
    p14.paragraph_format.space_after = Pt(2)
    render_question_prompt(p14, "Câu 14: Sắp xếp theo chiều giảm dần nhiệt độ sôi của các chất sau: methyl formate (1); acetic acid (2); ethanol (3)?")
    opts14 = [
        "A. (1) > (2) > (3).",
        "B. (2) > (1) > (3).",
        "C. (1) > (3) > (2).",
        "*D. (2) > (3) > (1).*" if is_teacher else "D. (2) > (3) > (1)."
    ]
    render_options_no_table(doc, opts14, is_teacher=is_teacher)
    if is_teacher:
        sol14 = (
            "• Acetic acid (CH~3~COOH) có liên kết hydrogen liên phân tử dạng dimer rất bền ➔ nhiệt độ sôi cao nhất (118,1°C).\n"
            "• Ethanol (C~2~H~5~OH) có liên kết hydrogen liên phân tử ➔ nhiệt độ sôi trung bình (78,3°C).\n"
            "• Methyl formate (HCOOCH~3~) là ester, giữa các phân tử không có liên kết hydrogen ➔ nhiệt độ sôi thấp nhất (31,5°C).\n"
            "• Thứ tự giảm dần nhiệt độ sôi: (2) > (3) > (1).\n"
            "➔ *Đáp án đúng: D*"
        )
        add_solution_no_table(doc, sol14)

    # CÂU 15
    p15 = doc.add_paragraph()
    p15.paragraph_format.space_before = Pt(4)
    p15.paragraph_format.space_after = Pt(2)
    render_question_prompt(p15, "Câu 15: Trong công nghiệp sản xuất đường saccharose, người ta sử dụng phương pháp nào để tách saccharose ra khỏi dung dịch nước mía?")
    opts15 = [
        "A. Sắc kí.",
        "*B. Kết tinh.*" if is_teacher else "B. Kết tinh.",
        "C. Chiết.",
        "D. Chưng cất."
    ]
    render_options_no_table(doc, opts15, is_teacher=is_teacher)
    if is_teacher:
        sol15 = (
            "• Quy trình sản xuất đường saccharose từ nước mía: nước mía sau khi được ép, tẩy màu và làm sạch tạp chất sẽ được cô đặc dưới áp suất giảm đến trạng thái bão hòa, sau đó làm lạnh để đường saccharose kết tinh tạo thành đường cát trắng.\n"
            "➔ *Đáp án đúng: B*"
        )
        add_solution_no_table(doc, sol15)

    # CÂU 16
    p16 = doc.add_paragraph()
    p16.paragraph_format.space_before = Pt(4)
    p16.paragraph_format.space_after = Pt(2)
    render_question_prompt(p16, "Câu 16: Chất nào sau đây thuộc loại polysaccharide?")
    opts16 = [
        "A. Maltose.",
        "B. Fructose.",
        "C. Glucose.",
        "*D. Cellulose.*" if is_teacher else "D. Cellulose."
    ]
    render_options_no_table(doc, opts16, is_teacher=is_teacher)
    if is_teacher:
        sol16 = (
            "• Carbohydrate được chia thành 3 nhóm chính:\n"
            "  + Monosaccharide: glucose, fructose.\n"
            "  + Disaccharide: saccharose, maltose.\n"
            "  + Polysaccharide: tinh bột, cellulose.\n"
            "➔ *Đáp án đúng: D*"
        )
        add_solution_no_table(doc, sol16)

    # CÂU 17
    p17 = doc.add_paragraph()
    p17.paragraph_format.space_before = Pt(4)
    p17.paragraph_format.space_after = Pt(2)
    render_question_prompt(p17, "Câu 17: Cho dãy các chất có công thức cấu tạo như sau:")

    p17_img = doc.add_paragraph()
    p17_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p17_img.paragraph_format.space_before = Pt(2)
    p17_img.paragraph_format.space_after = Pt(2)
    img_c17 = os.path.join(IMG_DIR, "cau17_amino_acid.png")
    if os.path.exists(img_c17):
        p17_img.add_run().add_picture(img_c17, width=Cm(12.0))

    p17_lead = doc.add_paragraph()
    p17_lead.paragraph_format.space_before = Pt(2)
    p17_lead.paragraph_format.space_after = Pt(2)
    render_rich_text(p17_lead, "Có bao nhiêu chất trong dãy là α - amino acid ?")

    opts17 = [
        "A. 3.",
        "B. 4.",
        "*C. 2.*" if is_teacher else "C. 2.",
        "D. 1."
    ]
    render_options_no_table(doc, opts17, is_teacher=is_teacher)
    if is_teacher:
        sol17 = (
            "• α-amino acid là amino acid có nhóm amino (-NH~2~) gắn trực tiếp vào nguyên tử carbon số 2 (carbon α ngay cạnh nhóm -COOH).\n"
            "• Quan sát từng chất:\n"
            "  + Chất (1): Nhóm -NH~2~ gắn ở carbon α kề nhóm -COOH ➔ là α-amino acid (2-aminobutanoic acid).\n"
            "  + Chất (2): Nhóm -NH~2~ gắn ở carbon số 4 (carbon γ) ➔ là γ-amino acid.\n"
            "  + Chất (3): Nhóm -NH~2~ gắn ở carbon số 3 (carbon β) ➔ là β-amino acid.\n"
            "  + Chất (4): Nhóm -NH~2~ gắn ở carbon α kề nhóm -COOH ➔ là α-amino acid (leucine / 2-amino-4-methylpentanoic acid).\n"
            "• Vậy trong dãy có đúng 2 chất là α-amino acid: chất (1) và chất (4).\n"
            "➔ *Đáp án đúng: C*"
        )
        add_solution_no_table(doc, sol17)

    # CÂU 18
    p18 = doc.add_paragraph()
    p18.paragraph_format.space_before = Pt(4)
    p18.paragraph_format.space_after = Pt(2)
    render_question_prompt(p18, "Câu 18: Phản ứng nào sau đây thuộc loại phản ứng trùng ngưng?")

    p18_c = doc.add_paragraph()
    p18_c.paragraph_format.left_indent = Cm(0.5)
    p18_c.paragraph_format.space_before = Pt(2)
    p18_c.paragraph_format.space_after = Pt(2)
    img_c18 = os.path.join(IMG_DIR, "cau18_opt_c.png")
    if os.path.exists(img_c18):
        p18_c.add_run().add_picture(img_c18, width=Cm(10.0))

    opts18 = [
        "A. HCl + H~2~NCH(CH~3~)COOH → ClH~3~NCH(CH~3~)COOH.",
        "B. H~2~NCH~2~COOH + CH~3~OH ⎯⎯(H~2~SO~4~, t°)→ H~2~NCH~2~COOCH~3~ + H~2~O.",
        "*C. Phản ứng C.*" if is_teacher else "C. Phản ứng C.",
        "D. H~2~NCH(CH~3~)COOH + NaOH → H~2~NCH(CH~3~)COONa + H~2~O."
    ]
    render_options_no_table(doc, opts18, is_teacher=is_teacher)
    if is_teacher:
        sol18 = (
            "• Phản ứng trùng ngưng là quá trình kết hợp nhiều phân tử nhỏ (monomer) thành phân tử lớn (polymer) đồng thời giải phóng các phân tử nhỏ khác (như H~2~O).\n"
            "• Phương án C biểu diễn phản ứng trùng ngưng của 6-aminohexanoic acid (ε-aminocaproic acid) tạo thành policaproamide (nylon-6) và giải phóng nước:\n"
            "  n H~2~N-[CH~2~]~5~-COOH ⎯⎯(t°)→ -(-HN-[CH~2~]~5~-CO-)-~n~ + n H~2~O\n"
            "• A là phản ứng tạo muối; B là phản ứng ester hóa thông thường; D là phản ứng trung hòa.\n"
            "➔ *Đáp án đúng: C*"
        )
        add_solution_no_table(doc, sol18)

    # ==========================================
    # PHẦN II: TRẮC NGHIỆM ĐÚNG SAI
    # ==========================================
    add_section_header(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # CÂU 1
    p2_1 = doc.add_paragraph()
    p2_1.paragraph_format.space_before = Pt(4)
    p2_1.paragraph_format.space_after = Pt(2)
    render_question_prompt(p2_1, "Câu 1: Cho công thức cấu tạo của ester X như sau:")

    p2_1_img = doc.add_paragraph()
    p2_1_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2_1_img.paragraph_format.space_before = Pt(2)
    p2_1_img.paragraph_format.space_after = Pt(2)
    img_p2_1 = os.path.join(IMG_DIR, "p2_cau1_ester.png")
    if os.path.exists(img_p2_1):
        p2_1_img.add_run().add_picture(img_p2_1, width=Cm(5.0))

    stmts1 = [
        ("a) X có tên gọi là methyl butanoate.", "Sai"),
        ("b) Phần trăm khối lượng của carbon trong ester X là 62,07% (Kết quả làm tròn đến hàng phần trăm).", "Đúng"),
        ("c) Phản ứng thủy phân ester X trong môi trường acid là phản ứng một chiều.", "Sai"),
        ("d) Xà phòng hoá ester X bằng dung dịch NaOH, đun nóng thu được alcohol bậc một.", "Đúng")
    ]
    for s_text, s_ans in stmts1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, s_text, ans=s_ans, is_teacher=is_teacher)

    if is_teacher:
        sol2_1 = (
            "• Công thức cấu tạo của ester X là CH~3~CH~2~CH~2~COOCH~2~CH~3~ (ethyl butanoate, C~6~H~12~O~2~, M = 116).\n"
            "• a) Sai: Ester X có tên gọi là ethyl butanoate (không phải methyl butanoate có CTPT là C~5~H~10~O~2~).\n"
            "• b) Đúng: Phần trăm khối lượng carbon trong X là: %C = (6 × 12 / 116) × 100% = 62,0689% ≈ 62,07%.\n"
            "• c) Sai: Phản ứng thủy phân ester trong môi trường acid là phản ứng thuận nghịch (hai chiều).\n"
            "• d) Đúng: Thủy phân X trong NaOH sinh ra ethanol C~2~H~5~OH (CH~3~CH~2~OH), đây là alcohol bậc một."
        )
        add_solution_no_table(doc, sol2_1)

    # CÂU 2
    p2_2 = doc.add_paragraph()
    p2_2.paragraph_format.space_before = Pt(4)
    p2_2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p2_2, "Câu 2: Levodopa (L-dopa) là tiền chất của dopamine, thường được sử dụng như một chất thay thế dopamine để điều trị bệnh Parkinson. L-dopa có công thức cấu tạo như sau:")

    p2_2_img = doc.add_paragraph()
    p2_2_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2_2_img.paragraph_format.space_before = Pt(2)
    p2_2_img.paragraph_format.space_after = Pt(2)
    img_p2_2 = os.path.join(IMG_DIR, "p2_cau2_ldopa.png")
    if os.path.exists(img_p2_2):
        p2_2_img.add_run().add_picture(img_p2_2, width=Cm(5.5))

    stmts2 = [
        ("a) Khi cho L-dopa phản ứng với dung dịch HCl tạo ra hợp chất hữu cơ X, trong X nguyên tố carbon chiếm 46,25% về khối lượng (Kết quả làm tròn đến hàng phần trăm).", "Đúng"),
        ("b) 1 mol L-dopa phản ứng tối đa với 1 mol KOH.", "Sai"),
        ("c) Trong phân tử L-dopa chứa 3 loại nhóm chức.", "Đúng"),
        ("d) Công thức phân tử của L-dopa là C~9~H~11~NO~4~.", "Đúng")
    ]
    for s_text, s_ans in stmts2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, s_text, ans=s_ans, is_teacher=is_teacher)

    if is_teacher:
        sol2_2 = (
            "• Công thức phân tử của L-dopa là C~9~H~11~NO~4~ (M = 197). Cấu tạo gồm 2 nhóm -OH phenol, 1 nhóm -NH~2~ và 1 nhóm -COOH.\n"
            "• a) Đúng: Khi phản ứng với HCl, nhóm -NH~2~ nhận HCl tạo muối C~9~H~11~NO~4~·HCl (C~9~H~12~ClNO~4~, M = 233,5).\n"
            "  %C = (9 × 12 / 233,5) × 100% = 46,2526% ≈ 46,25%.\n"
            "• b) Sai: Phân tử chứa 2 nhóm -OH phenol và 1 nhóm -COOH nên 1 mol L-dopa phản ứng tối đa với 3 mol KOH.\n"
            "• c) Đúng: Chứa 3 loại nhóm chức là: phenol (-OH), amine (-NH~2~) và carboxylic acid (-COOH).\n"
            "• d) Đúng: Công thức phân tử là C~9~H~11~NO~4~."
        )
        add_solution_no_table(doc, sol2_2)

    # CÂU 3
    p2_3 = doc.add_paragraph()
    p2_3.paragraph_format.space_before = Pt(4)
    p2_3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p2_3, "Câu 3: Benzocain là chất gây tê cục bộ. Nó là thành phần hoạt chất trong nhiều loại thuốc mỡ gây tê không kê đơn như các sản phẩm trị loét miệng thương hiệu Orajel. Benzocain được tổng hợp từ nguyên liệu ban đầu là toluidin qua nhiều phản ứng, đơn giản hoá quá trình theo sơ đồ sau:")

    p2_3_img = doc.add_paragraph()
    p2_3_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2_3_img.paragraph_format.space_before = Pt(2)
    p2_3_img.paragraph_format.space_after = Pt(2)
    img_p2_3 = os.path.join(IMG_DIR, "p2_cau3_benzocain.png")
    if os.path.exists(img_p2_3):
        p2_3_img.add_run().add_picture(img_p2_3, width=Cm(12.0))

    stmts3 = [
        ("a) Cả toluidin và benzocain đều tan tốt trong nước.", "Sai"),
        ("b) Benzocain có tính lưỡng tính vì vừa tác dụng với dung dịch HCl và dung dịch NaOH.", "Sai"),
        ("c) Trên phổ hồng ngoại của benzocain, tín hiệu (peak) ở 1700 cm^-1^ đặc trưng cho liên kết C-O.", "Sai"),
        ("d) Để sản xuất một lô thuốc benzocain gồm 2000 lọ (mỗi lọ chứa 25 gam benzocain) thì cần dùng 43,2 kg toluidin. Biết hiệu suất cả quá trình là 75%. (Kết quả cuối cùng được làm tròn đến hàng phần mười).", "Đúng")
    ]
    for s_text, s_ans in stmts3:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, s_text, ans=s_ans, is_teacher=is_teacher)

    if is_teacher:
        sol2_3 = (
            "• a) Sai: Toluidin (amine thơm) và benzocain (ester thơm) đều có gốc hydrocarbon lớn không phân cực nên rất ít tan trong nước.\n"
            "• b) Sai: Benzocain có nhóm -NH~2~ thể hiện tính base và nhóm ester bị thủy phân trong dung dịch NaOH, đây không phải tính chất của chất lưỡng tính.\n"
            "• c) Sai: Tín hiệu ở vùng 1700 cm^-1^ đặc trưng cho liên kết đôi C=O của nhóm ester (liên kết C-O hấp thụ ở vùng 1300 – 1000 cm^-1^).\n"
            "• d) Đúng: m(benzocain cần tạo) = 2000 × 25 g = 50 kg = 50 000 gam.\n"
            "  Sơ đồ: C~7~H~9~N (107) ➔ C~9~H~11~NO~2~ (165)\n"
            "  m(toluidin lý thuyết) = 50 × (107 / 165) = 32,4242 kg.\n"
            "  Do hiệu suất H = 75% nên m(toluidin thực tế) = 32,4242 / 0,75 = 43,232 kg ≈ 43,2 kg."
        )
        add_solution_no_table(doc, sol2_3)

    # CÂU 4
    p2_4 = doc.add_paragraph()
    p2_4.paragraph_format.space_before = Pt(4)
    p2_4.paragraph_format.space_after = Pt(2)
    render_question_prompt(p2_4, "Câu 4: Việc pha ethanol vào xăng góp phần hạn chế sự phụ thuộc vào nhiên liệu hóa thạch và giảm phát thải CO~2~. Sơ đồ sau đây tóm tắt quá trình sản xuất và tiêu thụ xăng E5 (chứa 5% ethanol về thể tích):")

    p2_4_img = doc.add_paragraph()
    p2_4_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2_4_img.paragraph_format.space_before = Pt(2)
    p2_4_img.paragraph_format.space_after = Pt(2)
    img_p2_4 = os.path.join(IMG_DIR, "p2_cau4_xang_e5.png")
    if os.path.exists(img_p2_4):
        p2_4_img.add_run().add_picture(img_p2_4, width=Cm(13.0))

    stmts4 = [
        ("a) Giai đoạn (1) có xảy ra phản ứng thủy phân và lên men rượu.", "Đúng"),
        ("b) Ở giai đoạn (2) xăng E5 được tạo ra bằng cách phối trộn 95% thể tích xăng truyền thống (thường là xăng RON 92) với 5% thể tích ethanol.", "Đúng"),
        ("c) Phản ứng trong giai đoạn (4) xảy ra nhờ ánh sáng mặt trời và chất diệp lục trong lá cây.", "Đúng"),
        ("d) Lên men 400 kg sắn khô (chứa 35% khối lượng là tinh bột, phần còn lại là các chất không có khả năng lên men thành ethanol) với hiệu suất quá trình (1) là 72%. Toàn bộ lượng ethanol sinh ra pha chế được tối đa 1435 lít xăng E5. Biết khối lượng riêng của ethanol nguyên chất là 0,8 g/mL. (Kết quả cuối cùng làm tròn đến hàng đơn vị).", "Sai")
    ]
    for s_text, s_ans in stmts4:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, s_text, ans=s_ans, is_teacher=is_teacher)

    if is_teacher:
        sol2_4 = (
            "• a) Đúng: Quá trình (1) gồm thủy phân tinh bột thành glucose và lên men glucose thành ethanol.\n"
            "• b) Đúng: Xăng E5 chứa 5% thể tích bio-ethanol và 95% thể tích xăng khoáng.\n"
            "• c) Đúng: Quá trình (4) là phản ứng quang hợp tạo tinh bột nhờ năng lượng ánh sáng mặt trời và chất diệp lục.\n"
            "• d) Sai: m(tinh bột) = 400 × 35% = 140 kg = 140 000 gam.\n"
            "  Sơ đồ: (C~6~H~10~O~5~)~n~ (162n) ➔ 2n C~2~H~5~OH (2n × 46 = 92n)\n"
            "  m(ethanol lý thuyết) = 140 × (92 / 162) = 79,506 kg.\n"
            "  Với hiệu suất H = 72%: m(ethanol thực tế) = 79,506 × 0,72 = 57,244 kg = 57 244 gam.\n"
            "  V(ethanol nguyên chất) = m / D = 57 244 / 0,8 = 71 555 mL = 71,555 lít.\n"
            "  Thể tích xăng E5 pha chế được: V(E5) = 71,555 / 0,05 = 1431,1 lít ≈ 1431 lít (khác 1435 lít)."
        )
        add_solution_no_table(doc, sol2_4)

    # ==========================================
    # PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN
    # ==========================================
    add_section_header(
        doc,
        "PHẦN III. Câu trắc nghiệm trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 6."
    )

    # CÂU 1
    p3_1 = doc.add_paragraph()
    p3_1.paragraph_format.space_before = Pt(4)
    p3_1.paragraph_format.space_after = Pt(2)
    render_question_prompt(p3_1, "Câu 1: Terephthalic acid có thể được tổng hợp từ arene X có công thức phân tử C~8~H~10~, bằng cách oxi hóa X theo sơ đồ gồm hai giai đoạn:\n(a) x C~8~H~10~ + y KMnO~4~ (đun nóng) → z C~6~H~4~(COOK)~2~ + t MnO~2~ + g KOH + i H~2~O\n(b) C~6~H~4~(COOK)~2~ + HCl → C~6~H~4~(COOH)~2~ + KCl\nTrong một thí nghiệm, từ 2,0 mL X (D = 0,861 g/mL) thu được 1,6 gam terephthalic acid. Biết KMnO~4~ và HCl được lấy dư. Hiệu suất của quá trình tổng hợp terephthalic acid từ X là h%.\nCho các phát biểu sau:\n(1) Giá trị của h là 59,3 (Kết quả cuối cùng làm tròn đến hàng phần mười).\n(2) Tỉ lệ x : g là 1:6.\n(3) X có tên gọi là 1,4 - dimethylbenzene.\n(4) Anion ^-^OOCC~6~H~4~COO^-^ là một base theo thuyết Brønsted – Lowry.\nHãy liệt kê các phát biểu đúng theo thứ tự tăng dần. (Ví dụ: 12, 124...).")

    p3_1_ans = doc.add_paragraph()
    p3_1_ans.paragraph_format.space_before = Pt(2)
    p3_1_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p3_1_ans, "► *Đáp án:* *134*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol3_1 = (
            "• X là p-xylene (1,4-dimethylbenzene): p-CH~3~-C~6~H~4~-CH~3~ (C~8~H~10~) ➔ (3) Đúng.\n"
            "• Cân bằng phản ứng oxi hóa bằng phương pháp thăng bằng electron:\n"
            "  1 C~8~H~10~ + 4 KMnO~4~ ⎯⎯(t°)→ 1 C~6~H~4~(COOK)~2~ + 4 MnO~2~ + 2 KOH + 2 H~2~O\n"
            "  Hệ số: x = 1, y = 4, z = 1, t = 4, g = 2, i = 2 ➔ Tỉ lệ x : g = 1 : 2 ➔ (2) Sai.\n"
            "• Tính hiệu suất h%:\n"
            "  m(X) = 2,0 mL × 0,861 g/mL = 1,722 gam ➔ n(X) = 1,722 / 106 mol.\n"
            "  m(terephthalic acid lý thuyết) = (1,722 / 106) × 166 = 2,69675 gam.\n"
            "  h = (1,6 / 2,69675) × 100% = 59,33% ≈ 59,3 ➔ (1) Đúng.\n"
            "• Anion ^-^OOCC~6~H~4~COO^-^ có khả năng nhận proton H^+^ nên là base theo Brønsted – Lowry ➔ (4) Đúng.\n"
            "➔ Các phát biểu đúng là (1), (3), (4). Trình tự tăng dần: *134*"
        )
        add_solution_no_table(doc, sol3_1)
    else:
        render_rich_text(p3_1_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # CÂU 2
    p3_2 = doc.add_paragraph()
    p3_2.paragraph_format.space_before = Pt(4)
    p3_2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p3_2, "Câu 2: Chỉ số xà phòng hóa là số mg potassium hydroxide cần dùng để xà phòng hóa hoàn toàn các triglyceride và trung hòa các acid béo tự do trong 1 gam chất béo. Trong thực nghiệm, một kỹ thuật viên thực hiện thí nghiệm xác định chỉ số xà phòng hóa của một mẫu mỡ cừu theo quy trình sau:\nBước 1: Cân chính xác 0,50 gam mẫu mỡ cừu cho vào bình nón khô.\nBước 2: Thêm vào 25,00 mL dung dịch KOH x M trong cồn tuyệt đối, đậy kín và đun hồi lưu trong 30 phút để đảm bảo phản ứng xảy ra hoàn toàn.\nBước 3: Sau khi làm nguội, thêm vài giọt phenolphthalein và chuẩn độ lượng KOH dư bằng dung dịch HCl 0,50 M. Lượng HCl cần dùng vừa đủ là 19,20 mL.\nBước 4: Tiến hành mẫu trắng song song (không có mỡ), cũng với 25,00 mL dung dịch KOH x M trên, cần 22,70 mL dung dịch HCl 0,50 M để chuẩn độ vừa đủ KOH.\nChỉ số xà phòng hóa của mẫu mỡ cừu trên là bao nhiêu?")

    p3_2_ans = doc.add_paragraph()
    p3_2_ans.paragraph_format.space_before = Pt(2)
    p3_2_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p3_2_ans, "► *Đáp án:* *196*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol3_2 = (
            "• Lượng KOH ban đầu trong 25,00 mL dung dịch (chuẩn độ qua mẫu trắng):\n"
            "  n(KOH ban đầu) = n(HCl mẫu trắng) = 0,02270 L × 0,50 M = 0,01135 mol.\n"
            "• Lượng KOH dư sau khi phản ứng với 0,50 gam mỡ cừu (mẫu thật):\n"
            "  n(KOH dư) = n(HCl mẫu thật) = 0,01920 L × 0,50 M = 0,00960 mol.\n"
            "• Lượng KOH đã phản ứng với 0,50 gam mỡ cừu:\n"
            "  n(KOH pư) = 0,01135 - 0,00960 = 0,00175 mol.\n"
            "• Khối lượng KOH tương ứng:\n"
            "  m(KOH) = 0,00175 mol × 56 g/mol = 0,098 gam = 98 mg.\n"
            "• Chỉ số xà phòng hóa (tính cho 1,0 gam mỡ cừu):\n"
            "  Chỉ số xà phòng hóa = 98 mg / 0,50 g = 196 (mg KOH/g chất béo).\n"
            "➔ *Đáp án: 196*"
        )
        add_solution_no_table(doc, sol3_2)
    else:
        render_rich_text(p3_2_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # CÂU 3
    p3_3 = doc.add_paragraph()
    p3_3.paragraph_format.space_before = Pt(4)
    p3_3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p3_3, "Câu 3: Cho phương trình hóa học của các phản ứng được gán số thứ tự dưới đây:\n(1) (C~15~H~31~COO)~3~C~3~H~5~ + 3NaOH ⎯⎯(t°)→ 3C~15~H~31~COONa + C~3~H~5~(OH)~3~\n(2) CH~3~COOH + C~2~H~5~OH ⎯⎯(H~2~SO~4~ đặc, t°)⇌ CH~3~COOC~2~H~5~ + H~2~O\n(3) (C~17~H~33~COO)~3~C~3~H~5~ + 3H~2~ ⎯⎯(xt, t°, p)→ (C~17~H~35~COO)~3~C~3~H~5~\n(4) C~6~H~12~O~6~ ⎯⎯(enzyme)→ 2CH~3~CH(OH)COOH\nGán số thứ tự phương trình hóa học của các phản ứng theo thứ tự tên gọi: ester hóa, hydrogen hóa, xà phòng hóa, lên men và sắp xếp theo trình tự thành dãy bốn số (ví dụ: 1234, 2341, 3421, …)")

    p3_3_ans = doc.add_paragraph()
    p3_3_ans.paragraph_format.space_before = Pt(2)
    p3_3_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p3_3_ans, "► *Đáp án:* *2314*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol3_3 = (
            "• Phân loại các phản ứng hóa học:\n"
            "  + Ester hóa: Phản ứng (2) giữa acid và alcohol.\n"
            "  + Hydrogen hóa: Phản ứng (3) cộng H~2~ vào nối đôi C=C của chất béo lỏng.\n"
            "  + Xà phòng hóa: Phản ứng (1) thủy phân chất béo trong dung dịch kiềm NaOH.\n"
            "  + Lên men: Phản ứng (4) lên men glucose thành lactic acid (lên men lactic).\n"
            "➔ Thứ tự: ester hóa (2) - hydrogen hóa (3) - xà phòng hóa (1) - lên men (4).\n"
            "➔ Dãy bốn số cần tìm là: *2314*"
        )
        add_solution_no_table(doc, sol3_3)
    else:
        render_rich_text(p3_3_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # CÂU 4
    p3_4 = doc.add_paragraph()
    p3_4.paragraph_format.space_before = Pt(4)
    p3_4.paragraph_format.space_after = Pt(2)
    render_question_prompt(p3_4, "Câu 4: Kết quả phân tích nguyên tố của hợp chất amine đơn chức X có phần trăm khối lượng các nguyên tố như sau: %C = 65,75%; %H = 15,07%; %N = 19,18%. Từ chất X thực hiện sơ đồ chuyển hóa sau:\nX ⎯⎯(+HNO~2~)→ Y ⎯⎯(+CuO, t°)→ Z ⎯⎯(+I~2~/NaOH)→ T (kết tủa màu vàng)\nCho các phát biểu sau đây:\n(a) Ở điều kiện thường, chất X là chất khí có mùi xốc và tan tốt trong nước.\n(b) Chất Y là alcohol bậc một, có tên thay thế là butan-2-ol.\n(c) Chất X là amine bậc một, có tên thay thế là butan-1-amine.\n(d) Khử chất Z bằng NaBH~4~ thì thu được chất Y.\nXác định số phát biểu đúng trong các phát biểu trên.")

    p3_4_ans = doc.add_paragraph()
    p3_4_ans.paragraph_format.space_before = Pt(2)
    p3_4_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p3_4_ans, "► *Đáp án:* *1*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol3_4 = (
            "• Tìm CTPT của amine X: n~C~ : n~H~ : n~N~ = (65,75/12) : (15,07/1) : (19,18/14) = 5,479 : 15,07 : 1,370 = 4 : 11 : 1.\n"
            "  Vì amine đơn chức nên CTPT của X là C~4~H~11~N.\n"
            "• Xác định cấu tạo: Z tạo kết tủa vàng CHI~3~ với I~2~/NaOH nên Z có nhóm CH~3~-C(=O)- (butan-2-one: CH~3~COCH~2~CH~3~).\n"
            "  Y bị oxi hóa tạo Z nên Y là alcohol bậc hai: butan-2-ol (CH~3~CH(OH)CH~2~CH~3~).\n"
            "  X tác dụng HNO~2~ sinh ra alcohol bậc 1 hoặc 2. Để tạo Y thì X là butan-2-amine: CH~3~CH(NH~2~)CH~2~CH~3~.\n"
            "• Đánh giá các phát biểu:\n"
            "  (a) Sai: X (butan-2-amine) là chất lỏng ở điều kiện thường.\n"
            "  (b) Sai: Y là butan-2-ol, là alcohol bậc hai (không phải bậc một).\n"
            "  (c) Sai: X có tên thay thế là butan-2-amine (không phải butan-1-amine).\n"
            "  (d) Đúng: Khử ketone Z (butan-2-one) bằng NaBH~4~ thu được alcohol bậc hai Y (butan-2-ol).\n"
            "➔ Số phát biểu đúng: *1*"
        )
        add_solution_no_table(doc, sol3_4)
    else:
        render_rich_text(p3_4_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # CÂU 5
    p3_5 = doc.add_paragraph()
    p3_5.paragraph_format.space_before = Pt(4)
    p3_5.paragraph_format.space_after = Pt(2)
    render_question_prompt(p3_5, "Câu 5: Cho sơ đồ của phản ứng: Cellulose ⎯⎯(+HNO~3~/H~2~SO~4~ đặc, t°)→ Cellulose trinitrate. Phân tử khối của cellulose trinitrate là a.n (với n là số mắt xích trong phân tử cellulose trinitrate). Xác định giá trị của a.")

    p3_5_ans = doc.add_paragraph()
    p3_5_ans.paragraph_format.space_before = Pt(2)
    p3_5_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p3_5_ans, "► *Đáp án:* *297*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol3_5 = (
            "• Phương trình hóa học của phản ứng:\n"
            "  [C~6~H~7~O~2~(OH)~3~]~n~ + 3n HNO~3~ ⎯⎯(H~2~SO~4~ đặc, t°)→ [C~6~H~7~O~2~(ONO~2~)~3~]~n~ + 3n H~2~O\n"
            "• Công thức một mắt xích của cellulose trinitrate là C~6~H~7~O~2~(ONO~2~)~3~ hay C~6~H~7~N~3~O~11~.\n"
            "• Phân tử khối một mắt xích:\n"
            "  a = 12 × 6 + 1 × 7 + 14 × 3 + 16 × 11 = 72 + 7 + 42 + 176 = 297.\n"
            "➔ *Đáp án: 297*"
        )
        add_solution_no_table(doc, sol3_5)
    else:
        render_rich_text(p3_5_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # CÂU 6
    p3_6 = doc.add_paragraph()
    p3_6.paragraph_format.space_before = Pt(4)
    p3_6.paragraph_format.space_after = Pt(2)
    render_question_prompt(p3_6, "Câu 6: Cho dãy các chất: tinh bột, cellulose, saccharose, glucose. Xác định số chất trong dãy có chứa liên kết glycoside trong phân tử.")

    p3_6_ans = doc.add_paragraph()
    p3_6_ans.paragraph_format.space_before = Pt(2)
    p3_6_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p3_6_ans, "► *Đáp án:* *3*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)
        sol3_6 = (
            "• Liên kết glycoside là liên kết giữa nhóm -OH hemiacetal/hemiketal của một đơn vị monosaccharide với nhóm -OH của một phân tử khác:\n"
            "  + Tinh bột: gồm các gốc α-glucose nối với nhau bằng liên kết α-1,4-glycosidic và α-1,6-glycosidic ➔ có liên kết glycoside.\n"
            "  + Cellulose: gồm các gốc β-glucose nối với nhau bằng liên kết β-1,4-glycosidic ➔ có liên kết glycoside.\n"
            "  + Saccharose: gồm gốc α-glucose và β-fructose nối với nhau bằng liên kết α-1,2-glycosidic ➔ có liên kết glycoside.\n"
            "  + Glucose: là monosaccharide dạng đơn phân tự do, không chứa liên kết glycoside.\n"
            "• Vậy có đúng 3 chất trong dãy chứa liên kết glycoside.\n"
            "➔ *Đáp án: 3*"
        )
        add_solution_no_table(doc, sol3_6)
    else:
        render_rich_text(p3_6_ans, "*Trả lời:* ............................................................", base_size=10.5)

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
    filename = f"De_Kiem_Tra_Hoa_12_HungYen_Ma0301{suffix}"
    out_path = os.path.join(OUTPUT_DIR, filename)
    doc.save(out_path)
    print(f"Đã tạo thành công file: {out_path}")
    return out_path

if __name__ == "__main__":
    build_exam_hungyen_0301(is_teacher=False)
    build_exam_hungyen_0301(is_teacher=True)
