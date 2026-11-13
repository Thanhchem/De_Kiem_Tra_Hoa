# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì I Hóa Học 10 - THPT Nguyễn Trường Tộ - Mã đề 104
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_10_NguyenTruongTo_Ma104_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_10_NguyenTruongTo_Ma104_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
    add_student_writing_lines,
    render_rich_text,
    set_run_style,
    add_p_border_bottom,
    COLOR_PRIMARY,
    COLOR_CORRECT,
    COLOR_SOLUTION,
    COLOR_GRAY
)

OUTPUT_DIR = r"C:\Antigravity_Thanh\San_Pham"
IMG_AG_PATH = r"C:\Antigravity_Thanh\San_Pham\images_hoa10\pho_khoi_luong_silver_ag.png"
IMG_ORBITALS_PATH = r"C:\Antigravity_Thanh\San_Pham\images_hoa10\hinh_orbital_s_va_p.png"

def add_quick_answer_table(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (16 CÂU)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (16 CÂU)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=16)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    p1_ans = ["D", "B", "C", "B", "D", "C", "C", "C", "A", "C", "D", "C", "A", "A", "A", "D"]
    for i in range(16):
        cell_q = tbl1.cell(0, i)
        cell_a = tbl1.cell(1, i)
        cell_q.width = Cm(1.1)
        cell_a.width = Cm(1.1)

        pq = cell_q.paragraphs[0]
        pq.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pq.paragraph_format.space_before = Pt(1)
        pq.paragraph_format.space_after = Pt(1)
        rq = pq.add_run(f"C{i+1}")
        set_run_style(rq, font_size=8.5, bold=True)

        pa = cell_a.paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pa.paragraph_format.space_before = Pt(1)
        pa.paragraph_format.space_after = Pt(1)
        ra = pa.add_run(p1_ans[i])
        set_run_style(ra, font_size=9.5, bold=True, color=COLOR_CORRECT)

        shd_q = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F4F8"/>')
        cell_q._tc.get_or_add_tcPr().append(shd_q)

    tblBdr1 = parse_xml(f'<w:tblBdr {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBdr>')
    tbl1._tbl.tblPr.append(tblBdr1)

    # 2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2 CÂU)
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(6)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (2 CÂU)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=5)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    cols_w = [Cm(2.8), Cm(3.8), Cm(3.8), Cm(3.8), Cm(3.8)]
    headers2 = ["Câu", "Lệnh a)", "Lệnh b)", "Lệnh c)", "Lệnh d)"]
    for j, h in enumerate(headers2):
        cell = tbl2.cell(0, j)
        cell.width = cols_w[j]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(h)
        set_run_style(r, font_size=9.5, bold=True)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F4F8"/>')
        cell._tc.get_or_add_tcPr().append(shd)

    p2_data = [
        ("Câu 1", ["ĐÚNG", "SAI", "ĐÚNG", "SAI"]),
        ("Câu 2", ["ĐÚNG", "ĐÚNG", "SAI", "SAI"])
    ]
    for row_idx, (q_name, answers) in enumerate(p2_data, start=1):
        cell_lbl = tbl2.cell(row_idx, 0)
        cell_lbl.width = cols_w[0]
        p = cell_lbl.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(q_name)
        set_run_style(r, font_size=9.5, bold=True)

        for col_idx, ans in enumerate(answers, start=1):
            cell_ans = tbl2.cell(row_idx, col_idx)
            cell_ans.width = cols_w[col_idx]
            p = cell_ans.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(ans)
            ans_color = COLOR_CORRECT if ans == "ĐÚNG" else COLOR_PRIMARY
            set_run_style(r, font_size=9.5, bold=True, color=ans_color)

    tblBdr2 = parse_xml(f'<w:tblBdr {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBdr>')
    tbl2._tbl.tblPr.append(tblBdr2)

    # 3. PHẦN III: TRẢ LỜI NGẮN (4 CÂU)
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(6)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*PHẦN III: TRẢ LỜI NGẮN (4 CÂU)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=4)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    p3_cols_w = [Cm(4.5), Cm(4.5), Cm(4.5), Cm(4.5)]
    p3_answers = ["12", "46", "24", "3"]
    for k in range(4):
        cell_q = tbl3.cell(0, k)
        cell_a = tbl3.cell(1, k)
        cell_q.width = p3_cols_w[k]
        cell_a.width = p3_cols_w[k]

        pq = cell_q.paragraphs[0]
        pq.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pq.paragraph_format.space_before = Pt(1)
        pq.paragraph_format.space_after = Pt(1)
        rq = pq.add_run(f"Câu {k+1}")
        set_run_style(rq, font_size=9.5, bold=True)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F4F8"/>')
        cell_q._tc.get_or_add_tcPr().append(shd)

        pa = cell_a.paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pa.paragraph_format.space_before = Pt(1)
        pa.paragraph_format.space_after = Pt(1)
        ra = pa.add_run(p3_answers[k])
        set_run_style(ra, font_size=10.5, bold=True, color=COLOR_CORRECT)

    tblBdr3 = parse_xml(f'<w:tblBdr {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBdr>')
    tbl3._tbl.tblPr.append(tblBdr3)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(8)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


def build_exam_doc(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    # 1. Khối tiêu đề đầu trang
    build_title_block(
        doc,
        school_line="TRƯỜNG THPT NGUYỄN TRƯỜNG TỘ\nTỔ HÓA - TIN",
        exam_title="KIỂM TRA GIỮA KÌ I - NĂM HỌC 2026 - 2027",
        subject_line="MÔN: HÓA HỌC, LỚP 10",
        num_pages="03",
        duration="45 phút",
        exam_code="104",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; Si = 28; S = 32; Cl = 35,5; K = 39; Ca = 40; Ag = 107,964",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table(doc)

    # ==========================================
    # A. PHẦN TRẮC NGHIỆM (7 ĐIỂM)
    # ==========================================
    p_a = doc.add_paragraph()
    p_a.paragraph_format.space_before = Pt(4)
    p_a.paragraph_format.space_after = Pt(2)
    render_rich_text(p_a, "*A. PHẦN TRẮC NGHIỆM (7 ĐIỂM)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # PHẦN I
    add_section_header(
        doc,
        title="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn",
        subtitle="Học sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi học sinh chỉ chọn một phương án."
    )

    # Câu 1
    p1 = doc.add_paragraph()
    render_question_prompt(p1, "Câu 1. Tổng số electron tối đa có trong mỗi lớp L và M lần lượt là")
    render_options_no_table(doc, [
        "A. 2 và 8.",
        "B. 8 và 32.",
        "C. 18 và 32.",
        "*D. 8 và 18.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Số electron tối đa trong lớp thứ n là 2n^2^ (với n <= 4):\n"
            "- Lớp L (n = 2): chứa tối đa 2*(2^2^) = 8 electron.\n"
            "- Lớp M (n = 3): chứa tối đa 2*(3^2^) = 18 electron.\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 2
    p2 = doc.add_paragraph()
    render_question_prompt(p2, "Câu 2. Cho phổ khối lượng MS của silver (Ag) ở hình dưới đây:")
    if os.path.exists(IMG_AG_PATH):
        p_img_ag = doc.add_paragraph()
        p_img_ag.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img_ag.paragraph_format.space_before = Pt(2)
        p_img_ag.paragraph_format.space_after = Pt(2)
        r_ag = p_img_ag.add_run()
        r_ag.add_picture(IMG_AG_PATH, width=Inches(3.4))

    p2_sub = doc.add_paragraph()
    render_rich_text(p2_sub, "Nguyên tử khối trung bình của nguyên tố silver là", base_size=10.5)
    render_options_no_table(doc, [
        "A. 108,012.",
        "*B. 107,964.*",
        "C. 108,000.",
        "D. 107,987."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Từ phổ khối lượng ta có:\n"
            "- Đồng vị ^107^Ag chiếm 51,8%.\n"
            "- Đồng vị ^109^Ag chiếm 48,2%.\n"
            "Nguyên tử khối trung bình của Ag là:\n"
            "A_tb = (107 * 51,8 + 109 * 48,2) / 100 = (5542,6 + 5253,8) / 100 = 10796,4 / 100 = 107,964.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 3
    p3 = doc.add_paragraph()
    render_question_prompt(p3, "Câu 3. Cho các phát biểu sau:\n"
                                "(a) Phân lớp 4s có mức năng lượng cao hơn phân lớp 3d.\n"
                                "(b) Trong mỗi orbital chỉ chứa tối đa 2 electron có chiều tự quay ngược nhau.\n"
                                "(c) Lớp M có 4 phân lớp.\n"
                                "(d) Số electron tối đa trong phân lớp 3d là 18.\n"
                                "(e) Số orbital của phân lớp s, p, d lần lượt là 1, 3, 5.\n"
                                "(f) Những electron ở gần hạt nhân có mức năng lượng cao nhất.\n"
                                "Số phát biểu đúng là")
    render_options_no_table(doc, [
        "A. 5.",
        "B. 3.",
        "*C. 2.*",
        "D. 4."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- (a) Sai: Theo thứ tự năng lượng 1s 2s 2p 3s 3p 4s 3d..., phân lớp 4s có mức năng lượng thấp hơn 3d.\n"
            "- (b) Đúng: Theo nguyên lí Pauli, mỗi orbital chứa tối đa 2e có chiều tự quay ngược nhau.\n"
            "- (c) Sai: Lớp M (n = 3) có 3 phân lớp (3s, 3p, 3d).\n"
            "- (d) Sai: Phân lớp 3d chứa tối đa 10 electron.\n"
            "- (e) Đúng: Phân lớp s có 1 AO, p có 3 AO, d có 5 AO.\n"
            "- (f) Sai: Electron càng gần hạt nhân liên kết càng bền, có mức năng lượng thấp nhất.\n"
            "Vậy có 2 phát biểu đúng là (b) và (e).\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 4
    p4 = doc.add_paragraph()
    render_question_prompt(p4, "Câu 4. Trong hạt nhân nguyên tử, hạt mang điện là")
    render_options_no_table(doc, [
        "A. proton và neutron.",
        "*B. proton.*",
        "C. electron.",
        "D. proton và electron."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Hạt nhân nguyên tử chỉ chứa proton (mang điện dương) và neutron (không mang điện). Do đó hạt mang điện trong hạt nhân chỉ có proton.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 5
    p5 = doc.add_paragraph()
    render_question_prompt(p5, "Câu 5. Dãy nào dưới đây gồm các đồng vị của cùng một nguyên tố hóa học?")
    render_options_no_table(doc, [
        "A. ^40^~18~X, ^40^~19~Y.",
        "B. ^19^~9~X, ^20^~10~Y.",
        "C. ^14^~6~X, ^14^~7~Y.",
        "*D. ^28^~14~X, ^29^~14~Y.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Đồng vị là những nguyên tử có cùng số proton (cùng số hiệu nguyên tử Z). Dãy D có cùng Z = 14 (nguyên tố Silicon) nên là các đồng vị của nhau.\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 6
    p6 = doc.add_paragraph()
    render_question_prompt(p6, "Câu 6. Nguyên tố hóa học là tập hợp các nguyên tử có cùng")
    render_options_no_table(doc, [
        "A. số hạt neutron.",
        "B. tổng số hạt proton và neutron.",
        "*C. điện tích hạt nhân.*",
        "D. số khối."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Theo định nghĩa, nguyên tố hóa học là tập hợp những nguyên tử có cùng điện tích hạt nhân (cùng số proton Z).\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 7
    p7 = doc.add_paragraph()
    render_question_prompt(p7, "Câu 7. Sự phân bố electron theo ô orbital nào dưới đây là đúng?")
    render_options_no_table(doc, [
        "A. [↑↑]",
        "B. [↑↑][↑][ ]",
        "*C. [↑][↑][↑]*",
        "D. [↑↓][↑][ ]"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Theo quy tắc Hund, trong cùng một phân lớp, các electron có xu hướng phân bố đều vào các orbital sao cho số electron độc thân là tối đa và có chiều tự quay song song với nhau. Do đó với 3 electron trong 3 AO p thì cách biểu diễn [↑][↑][↑] là đúng quy tắc Hund và nguyên lí Pauli.\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 8
    p8 = doc.add_paragraph()
    render_question_prompt(p8, "Câu 8. Cho cấu hình electron nguyên tử của một số nguyên tố như sau: [Ne]3s^2^, [Ar]3d^10^ 4s^1^, [Ar]3d^10^ 4s^2^ 4p^5^, [He] 2s^2^ 2p^6^, [Ar]3d^7^ 4s^2^. Số nguyên tố thuộc nhóm A và nhóm B lần lượt là")
    render_options_no_table(doc, [
        "A. 4; 1.",
        "B. 2; 2.",
        "*C. 3; 2.*",
        "D. 2; 3."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Phân loại các nguyên tố:\n"
            "1. [Ne]3s^2^: nguyên tố s -> nhóm IIA (nhóm A).\n"
            "2. [Ar]3d^10^ 4s^1^: nguyên tố d -> nhóm IB (nhóm B).\n"
            "3. [Ar]3d^10^ 4s^2^ 4p^5^: nguyên tố p -> nhóm VIIA (nhóm A).\n"
            "4. [He] 2s^2^ 2p^6^: nguyên tố p -> nhóm VIIIA (nhóm A).\n"
            "5. [Ar]3d^7^ 4s^2^: nguyên tố d -> nhóm VIIIB (nhóm B).\n"
            "Vậy có 3 nguyên tố nhóm A và 2 nguyên tố nhóm B.\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 9
    p9 = doc.add_paragraph()
    render_question_prompt(p9, "Câu 9. Nguyên tử X có 56 electron và 81 neutron. Kí hiệu nguyên tử của X là")
    render_options_no_table(doc, [
        "*A. ^137^~56~X.*",
        "B. ^56^~81~X.",
        "C. ^81^~56~X.",
        "D. ^112^~56~X."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Nguyên tử X có Z = e = 56, N = 81.\n"
            "Số khối A = Z + N = 56 + 81 = 137.\n"
            "Kí hiệu nguyên tử là: ^137^~56~X (Barium - Ba).\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 10
    p10 = doc.add_paragraph()
    render_question_prompt(p10, "Câu 10. Số nguyên tố trong chu kì 3 và chu kì 5 lần lượt là")
    render_options_no_table(doc, [
        "A. 18 và 18.",
        "B. 18 và 8.",
        "*C. 8 và 18.*",
        "D. 8 và 8."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Trong bảng tuần hoàn:\n"
            "- Chu kì 3 là chu kì nhỏ, gồm 8 nguyên tố (từ Na đến Ar).\n"
            "- Chu kì 5 là chu kì lớn, gồm 18 nguyên tố (từ Rb đến Xe).\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 11
    p11 = doc.add_paragraph()
    render_question_prompt(p11, "Câu 11. Biết rằng một loại nguyên tử Sodium (Na) có 11 proton và 12 neutron. Phát biểu nào sau đây không đúng?")
    render_options_no_table(doc, [
        "A. Số khối của hạt nhân nguyên tử Sodium là 23.",
        "B. Số hạt mang điện trong nguyên tử Sodium là 22.",
        "C. Vỏ nguyên tử Sodium có 11 electron.",
        "*D. Điện tích hạt nhân nguyên tử Sodium là 11.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Điện tích hạt nhân phải kèm dấu dương hoặc đơn vị: +11 (hoặc +11e hay +1,762.10^-18^ C). Con số 11 chỉ là số đơn vị điện tích hạt nhân (Z = 11). Do đó phát biểu D không đúng.\n"
            "- Các phát biểu A, B, C đều hoàn toàn chính xác (A = 11+12 = 23; hạt mang điện = 2*11 = 22; electron = 11).\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 12
    p12 = doc.add_paragraph()
    render_question_prompt(p12, "Câu 12. Nguyên tử của nguyên tố X có cấu hình electron là 1s^2^ 2s^1^. Nguyên tố hóa học có cấu hình electron nào sau đây thuộc cùng một nhóm với nguyên tố X?")
    render_options_no_table(doc, [
        "A. 1s^2^ 2s^2^ 2p^1^.",
        "B. 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^6^ 3d^5^ 4s^1^.",
        "*C. 1s^2^ 2s^2^ 2p^6^ 3s^1^.*",
        "D. 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^1^."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Nguyên tố X (Lithium) có 1 electron ở lớp ngoài cùng dạng ns^1^, thuộc nhóm IA.\n"
            "Cấu hình ở C: 1s^2^ 2s^2^ 2p^6^ 3s^1^ (Sodium - Na) cũng có 1 electron lớp ngoài cùng dạng ns^1^, thuộc nhóm IA.\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 13
    p13 = doc.add_paragraph()
    render_question_prompt(p13, "Câu 13. Nguyên tử của nguyên tố X có electron ở mức năng lượng cao nhất là 3p. Nguyên tử của nguyên tố Y có một electron ở lớp N. Nguyên tử X và Y có số electron hơn kém nhau là 3. Nguyên tử Y, X lần lượt là")
    render_options_no_table(doc, [
        "*A. kim loại, phi kim.*",
        "B. kim loại, kim loại.",
        "C. khí hiếm, kim loại.",
        "D. phi kim, kim loại."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Nguyên tử Y có 1 electron ở lớp N (lớp 4), có cấu hình kết thúc ở 4s^1^ => Y là Potassium (K, Z = 19), là một kim loại kiềm.\n"
            "- Số electron của X và Y hơn kém nhau 3: Z_X = 19 - 3 = 16 (Sulfur - S, cấu hình: [Ne]3s^2^ 3p^4^, có electron mức năng lượng cao nhất là 3p, thỏa mãn đề bài). S là một phi kim.\n"
            "- Thứ tự hỏi: 'Nguyên tử Y, X lần lượt là' => Y là kim loại, X là phi kim.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 14
    p14 = doc.add_paragraph()
    render_question_prompt(p14, "Câu 14. Nguyên tố X được sử dụng như một nguyên liệu cơ bản trong sản xuất phân bón. X là chất vô cùng cần thiết cho sự tăng trưởng và phát triển của các cây xanh. Ngoài những công dụng trên, X còn là một khoáng chất có vai trò rất quan trọng đối với sức khỏe con người. Nguyên tử của nguyên tố X có tổng số hạt (proton, electron, neutron) là 58, trong đó tổng số hạt mang điện nhiều hơn số hạt không mang điện là 18. Nguyên tử X có số hạt proton là")
    render_options_no_table(doc, [
        "*A. 19.*",
        "B. 20.",
        "C. 38.",
        "D. 39."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Ta có hệ phương trình:\n"
            "2Z + N = 58\n"
            "2Z - N = 18\n"
            "=> 4Z = 76 => Z = 19 (Potassium - K), N = 20.\n"
            "Số hạt proton của X là 19.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 15
    p15 = doc.add_paragraph()
    render_question_prompt(p15, "Câu 15. Năm 1911, Rutherford cho chùm hạt alpha bắn vào một lá vàng rất mỏng và quan sát thấy hầu hết các hạt alpha đi thẳng qua lá vàng. Kết quả này chứng tỏ")
    render_options_no_table(doc, [
        "*A. nguyên tử có cấu tạo rỗng.*",
        "B. hạt nhân nguyên tử mang điện tích âm.",
        "C. electron có khối lượng rất lớn.",
        "D. nguyên tử không chứa hạt mang điện."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Hầu hết các hạt alpha đi thẳng chứng tỏ phần lớn không gian trong nguyên tử là rỗng, hạt nhân mang điện dương có kích thước rất nhỏ nằm ở tâm.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 16
    p16 = doc.add_paragraph()
    render_question_prompt(p16, "Câu 16. Trong bảng tuần hoàn, chu kì gồm các nguyên tố mà nguyên tử của chúng có cùng")
    render_options_no_table(doc, [
        "A. số phân lớp electron.",
        "B. cấu hình electron.",
        "C. số electron hóa trị.",
        "*D. số lớp electron.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Chu kì là dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được sắp xếp theo chiều tăng dần của điện tích hạt nhân.\n"
            "-> *Đáp án đúng: D*"
        )

    # ==========================================
    # PHẦN II. TRẮC NGHIỆM ĐÚNG / SAI
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai."
    )

    # Câu 1 Phần II
    p_ii_1 = doc.add_paragraph()
    render_question_prompt(p_ii_1, "Câu 1. Sulfur dạng kem bôi được sử dụng để điều trị mụn trứng cá. Nguyên tử sulfur có phân lớp electron ngoài cùng là 3p^4^.")

    stmts_1 = [
        ("a) Sulfur thuộc chu kì 3, nhóm VIA và là nguyên tố phi kim.", "ĐÚNG"),
        ("b) Sulfur có số hiệu nguyên tử là 14.", "SAI"),
        ("c) Nguyên tử ^34^~16~X là đồng vị của sulfur.", "ĐÚNG"),
        ("d) Ở trạng thái cơ bản, nguyên tử sulfur có 4 electron hóa trị.", "SAI")
    ]
    for s_text, ans in stmts_1:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, s_text, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(doc,
            "Lời giải chi tiết:\n"
            "Cấu hình electron đầy đủ của Sulfur: 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^4^ (Z = 16).\n"
            "- a) *Đúng*: Có 3 lớp electron -> chu kì 3; 6 electron lớp ngoài cùng (3s^2^ 3p^4^), nguyên tố p -> nhóm VIA; là nguyên tố phi kim.\n"
            "- b) *Sai*: Sulfur có số hiệu nguyên tử Z = 16 (không phải 14).\n"
            "- c) *Đúng*: Nguyên tử ^34^~16~X có cùng số proton Z = 16 với sulfur nên là đồng vị của sulfur.\n"
            "- d) *Sai*: Sulfur thuộc nhóm VIA nên có 6 electron hóa trị (ở phân lớp 3s^2^ 3p^4^), không phải 4."
        )

    # Câu 2 Phần II
    p_ii_2 = doc.add_paragraph()
    render_question_prompt(p_ii_2, "Câu 2. Hình dưới mô tả orbital (a) và orbital (b) chứa electron trong nguyên tử Sodium (Na) ở trạng thái cơ bản. Mức năng lượng của orbital (a) cao hơn orbital (b).")

    if os.path.exists(IMG_ORBITALS_PATH):
        p_img_orb = doc.add_paragraph()
        p_img_orb.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img_orb.paragraph_format.space_before = Pt(2)
        p_img_orb.paragraph_format.space_after = Pt(2)
        r_orb = p_img_orb.add_run()
        r_orb.add_picture(IMG_ORBITALS_PATH, width=Inches(3.8))

    stmts_2 = [
        ("a) Số electron trong orbital (b) gấp đôi số electron trong orbital (a).", "ĐÚNG"),
        ("b) Số tám nổi (b) là hình dạng của AO p.", "ĐÚNG"),
        ("c) Electron trên orbital (a) nằm gần hạt nhân hơn electron trên orbital (b).", "SAI"),
        ("d) Electron trong orbital (a) và (b) cùng thuộc 1 lớp electron.", "SAI")
    ]
    for s_text, ans in stmts_2:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, s_text, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(doc,
            "Lời giải chi tiết:\n"
            "Cấu hình electron của Sodium (Na, Z = 11): 1s^2^ 2s^2^ 2p^6^ 3s^1^.\n"
            "- Hình (a) là orbital s (dạng hình cầu đối xứng tâm).\n"
            "- Hình (b) là orbital p (dạng số tám nổi).\n"
            "- Theo đề bài: Mức năng lượng của AO (a) cao hơn AO (b). Vì Na chỉ có các AO 1s, 2s, 2p, 3s và năng lượng 2p < 3s, nên orbital (a) phải là 3s và orbital (b) là 2p.\n"
            "- a) *Đúng*: Orbital 3s chỉ có 1 electron (3s^1^); mỗi orbital 2p bão hòa có 2 electron (2p^6^). Do đó số electron trong AO (b) gấp đôi số electron trong AO (a).\n"
            "- b) *Đúng*: Hình số tám nổi là hình dạng đặc trưng của orbital nguyên tử p.\n"
            "- c) *Sai*: Orbital 3s (lớp M, n=3) ở xa hạt nhân hơn so với orbital 2p (lớp L, n=2).\n"
            "- d) *Sai*: Orbital 3s thuộc lớp thứ 3, orbital 2p thuộc lớp thứ 2, thuộc 2 lớp khác nhau."
        )

    # ==========================================
    # PHẦN III. CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm trả lời ngắn",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4."
    )

    # Câu 1 (Phần III)
    p_iii_1 = doc.add_paragraph()
    render_question_prompt(p_iii_1, "Câu 1. Khí carbon dioxide (CO~2~) là nguyên nhân chính gây nên hiệu ứng nhà kính. Carbon có 2 đồng vị ^12^~6~C, ^13^~6~C; oxygen có 3 đồng vị ^16^~8~O, ^17^~8~O, ^18^~8~O. Số loại phân tử CO~2~ tạo thành là bao nhiêu?")
    if is_teacher:
        p_ans1 = doc.add_paragraph()
        p_ans1.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans1, "*Trả lời:* *12*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Cấu tạo phân tử CO~2~ có dạng O - C - O:\n"
            "- Số cách chọn 1 nguyên tử C từ 2 đồng vị: 2 cách.\n"
            "- Số cặp 2 nguyên tử O chọn từ 3 đồng vị (kể cả giống nhau):\n"
            "  + 3 cặp gồm 2 đồng vị O giống nhau: (^16^O, ^16^O), (^17^O, ^17^O), (^18^O, ^18^O).\n"
            "  + 3 cặp gồm 2 đồng vị O khác nhau: (^16^O, ^17^O), (^16^O, ^18^O), (^17^O, ^18^O).\n"
            "  => Có 3 + 3 = 6 cặp nguyên tử O.\n"
            "- Tổng số loại phân tử CO~2~ tạo thành là: 2 * 6 = 12 loại.\n"
            "-> *Đáp án: 12*"
        )
    else:
        p_ans1 = doc.add_paragraph()
        p_ans1.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans1, "*Trả lời:* ............................................................", base_size=10.5)

    # Câu 2 (Phần III)
    p_iii_2 = doc.add_paragraph()
    render_question_prompt(p_iii_2, "Câu 2. Nguyên tố X thuộc chu kì 3, nhóm IIA. Nguyên tố Y thuộc cùng chu kì với X và thuộc nhóm VIIA. X và Y kết hợp với nhau tạo thành hợp chất có công thức XY~2~. Tổng số proton có trong một phân tử XY~2~ là bao nhiêu?")
    if is_teacher:
        p_ans2 = doc.add_paragraph()
        p_ans2.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans2, "*Trả lời:* *46*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "- X thuộc chu kì 3, nhóm IIA => X là Magnesium (Mg, Z_X = 12).\n"
            "- Y thuộc chu kì 3, nhóm VIIA => Y là Chlorine (Cl, Z_Y = 17).\n"
            "- Hợp chất XY~2~ là MgCl~2~.\n"
            "Tổng số proton trong 1 phân tử XY~2~ là:\n"
            "Z_tổng = Z_X + 2*Z_Y = 12 + 2*17 = 12 + 34 = 46.\n"
            "-> *Đáp án: 46*"
        )
    else:
        p_ans2 = doc.add_paragraph()
        p_ans2.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans2, "*Trả lời:* ............................................................", base_size=10.5)

    # Câu 3 (Phần III)
    p_iii_3 = doc.add_paragraph()
    render_question_prompt(p_iii_3, "Câu 3. Cho các phát biểu sau:\n"
                                    "(1) Nguyên tử có kích thước vô cùng nhỏ so với kích thước hạt nhân nguyên tử.\n"
                                    "(2) Số khối của nguyên tử bằng tổng số proton và số neutron.\n"
                                    "(3) Tất cả các nguyên tử đều chứa proton, electron và neutron.\n"
                                    "(4) Khối lượng của nguyên tử tập trung hầu hết ở hạt nhân nguyên tử.\n"
                                    "(5) Đồng vị là các nguyên tố có cùng số proton nhưng khác nhau số neutron.\n"
                                    "Viết số thứ tự của các phát biểu đúng theo thứ tự tăng dần từ trái sang phải (ví dụ: 12 hoặc 123 hoặc 23...).")
    if is_teacher:
        p_ans3 = doc.add_paragraph()
        p_ans3.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans3, "*Trả lời:* *24* _(hoặc 245)_", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Phân tích từng phát biểu:\n"
            "- (1) Sai: Nguyên tử có kích thước vô cùng LỚN so với hạt nhân nguyên tử (lớn gấp khoảng 10 000 lần).\n"
            "- (2) Đúng: Số khối A = Z + N (tổng proton và neutron).\n"
            "- (3) Sai: Đồng vị ^1^~1~H chỉ gồm 1 proton và 1 electron, không có neutron.\n"
            "- (4) Đúng: Vì khối lượng electron không đáng kể nên khối lượng nguyên tử tập trung hầu hết ở hạt nhân.\n"
            "- (5) Sai theo chuẩn SGK Hóa học: 'Đồng vị là các NGUYÊN TỬ có cùng số proton nhưng khác nhau về số neutron' (không phải 'các nguyên tố').\n"
            "Do đó các phát biểu đúng là (2) và (4) => Viết: 24.\n"
            "*(Lưu ý: Nếu người ra đề chấp nhận phát biểu (5) theo cách diễn đạt thông tục thì đáp án là 245).*\n"
            "-> *Đáp án: 24*"
        )
    else:
        p_ans3 = doc.add_paragraph()
        p_ans3.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans3, "*Trả lời:* ............................................................", base_size=10.5)

    # Câu 4 (Phần III)
    p_iii_4 = doc.add_paragraph()
    render_question_prompt(p_iii_4, "Câu 4. Cho 7 nguyên tố có số hiệu nguyên tử lần lượt là: 5, 18, 11, 13, 16, 19, 2. Trong các nguyên tố trên, có bao nhiêu nguyên tố là kim loại?")
    if is_teacher:
        p_ans4 = doc.add_paragraph()
        p_ans4.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans4, "*Trả lời:* *3*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Xét 7 nguyên tố:\n"
            "- Z = 5 (Boron - B): phi kim.\n"
            "- Z = 18 (Argon - Ar): khí hiếm.\n"
            "- Z = 11 (Sodium - Na): kim loại kiềm.\n"
            "- Z = 13 (Aluminium - Al): kim loại.\n"
            "- Z = 16 (Sulfur - S): phi kim.\n"
            "- Z = 19 (Potassium - K): kim loại kiềm.\n"
            "- Z = 2 (Helium - He): khí hiếm.\n"
            "Vậy có 3 nguyên tố là kim loại (Z = 11, 13, 19).\n"
            "-> *Đáp án: 3*"
        )
    else:
        p_ans4 = doc.add_paragraph()
        p_ans4.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans4, "*Trả lời:* ............................................................", base_size=10.5)

    # ==========================================
    # B. PHẦN TỰ LUẬN (3 ĐIỂM)
    # ==========================================
    p_b = doc.add_paragraph()
    p_b.paragraph_format.space_before = Pt(8)
    p_b.paragraph_format.space_after = Pt(2)
    render_rich_text(p_b, "*B. PHẦN TỰ LUẬN (3 ĐIỂM)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # Câu 1 Tự luận
    p_tl1 = doc.add_paragraph()
    render_question_prompt(p_tl1, "Câu 1. Silicon (Si) được sử dụng trong nhiều ngành công nghiệp: gốm, men sứ, thủy tinh, luyện thép, vật liệu bán dẫn ...\n"
                                  "a) Viết cấu hình electron nguyên tử của nguyên tố silicon (Z = 14).\n"
                                  "b) Xác định vị trí của silicon (ô nguyên tố, chu kì, nhóm) trong bảng tuần hoàn và giải thích.")
    if is_teacher:
        add_solution_no_table(doc,
            "Lời giải chi tiết:\n"
            "a) Cấu hình electron nguyên tử Silicon (Z = 14):\n"
            "   1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^2^ (hoặc [Ne] 3s^2^ 3p^2^). (0,5 điểm)\n"
            "b) Vị trí của silicon trong bảng tuần hoàn:\n"
            "   - Ô nguyên tố: số 14 (vì số hiệu nguyên tử Z = 14, có 14 proton).\n"
            "   - Chu kì: Chu kì 3 (vì nguyên tử có 3 lớp electron).\n"
            "   - Nhóm: Nhóm IVA (vì có 4 electron hóa trị ở lớp ngoài cùng 3s^2^ 3p^2^, electron cuối cùng điền vào phân lớp p). (0,5 điểm)"
        )
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Câu 2 Tự luận
    p_tl2 = doc.add_paragraph()
    render_question_prompt(p_tl2, "Câu 2. Cho X, Y, T là ba nguyên tố liên tiếp nhau trong một chu kì của bảng tuần hoàn các nguyên tố hóa học. Tổng số các hạt mang điện trong thành phần cấu tạo nguyên tử của X, Y, T bằng 72. Xác định số hiệu nguyên tử của các nguyên tố X, Y, T?")
    if is_teacher:
        add_solution_no_table(doc,
            "Lời giải chi tiết:\n"
            "Vì X, Y, T là 3 nguyên tố liên tiếp nhau trong cùng một chu kì nên số hiệu nguyên tử của chúng hơn kém nhau 1 đơn vị.\n"
            "Giả sử Z_X < Z_Y < Z_T, ta có: Z_X = Z_Y - 1 và Z_T = Z_Y + 1.\n"
            "Trong mỗi nguyên tử, số hạt mang điện là 2Z (gồm Z proton và Z electron).\n"
            "Tổng số hạt mang điện của 3 nguyên tử là 72:\n"
            "2*Z_X + 2*Z_Y + 2*Z_T = 72 => Z_X + Z_Y + Z_T = 36\n"
            "<=> (Z_Y - 1) + Z_Y + (Z_Y + 1) = 36\n"
            "<=> 3*Z_Y = 36 => Z_Y = 12 (Magnesium - Mg).\n"
            "=> Z_X = 12 - 1 = 11 (Sodium - Na).\n"
            "=> Z_T = 12 + 1 = 13 (Aluminium - Al).\n"
            "Kiểm tra: Na (Z=11), Mg (Z=12), Al (Z=13) đều thuộc cùng chu kì 3 của bảng tuần hoàn (thỏa mãn).\n"
            "Vậy số hiệu nguyên tử của các nguyên tố là:\n"
            "Z_X = 11; Z_Y = 12; Z_T = 13."
        )
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 3 Tự luận
    p_tl3 = doc.add_paragraph()
    render_question_prompt(p_tl3, "Câu 3. Một nguyên tố R có 3 đồng vị X, Y, Z, tổng số hạt electron, proton và neutron của 3 đồng vị bằng 111. Số neutron của đồng vị X bằng số proton, số neutron của đồng vị Z hơn đồng vị Y 1 hạt.\n"
                                  "a) Xác định số khối của 3 đồng vị.\n"
                                  "b) Cho biết tỷ lệ số nguyên tử của các đồng vị như sau: X : Y = 393 : 50 và Y : Z = 50 : 57, xác định nguyên tử khối trung bình của R và khối lượng của 9,033.10^22^ nguyên tử R.\n"
                                  "(Biết số Avogadro = 6,022.10^23^)")
    if is_teacher:
        add_solution_no_table(doc,
            "Lời giải chi tiết:\n"
            "a) Xác định số khối của 3 đồng vị:\n"
            "- Gọi Z là số proton của nguyên tố R (cả 3 đồng vị đều có Z proton và Z electron).\n"
            "- Gọi N_X, N_Y, N_Z lần lượt là số neutron của các đồng vị X, Y, Z.\n"
            "Tổng số hạt của 3 đồng vị: (2Z + N_X) + (2Z + N_Y) + (2Z + N_Z) = 111\n"
            "=> 6Z + N_X + N_Y + N_Z = 111. (1)\n"
            "Theo đề:\n"
            "+ N_X = Z. (2)\n"
            "+ N_Z = N_Y + 1. (3)\n"
            "Thay (2) và (3) vào (1): 6Z + Z + N_Y + (N_Y + 1) = 111 => 7Z + 2*N_Y = 110.\n"
            "=> 2*N_Y = 110 - 7Z.\n"
            "Vì 2*N_Y chẵn và 110 chẵn nên Z phải là số nguyên chẵn (Z < 110/7 ≈ 15,7).\n"
            "Mặt khác đối với các đồng vị bền ta có 1 <= N_Y/Z <= 1,5:\n"
            "- Với Z = 12 (Mg): 2*N_Y = 110 - 84 = 26 => N_Y = 13 => N_Y/Z = 13/12 ≈ 1,083 (thỏa mãn rất bền).\n"
            "  + N_X = Z = 12 => Số khối A_X = 12 + 12 = 24.\n"
            "  + N_Y = 13 => Số khối A_Y = 12 + 13 = 25.\n"
            "  + N_Z = 13 + 1 = 14 => Số khối A_Z = 12 + 14 = 26.\n"
            "(Với Z = 14: 2*N_Y = 12 => N_Y = 6 => N_Y/Z = 6/14 < 1 loại; Z <= 10: N_Y/Z > 1,5 loại).\n"
            "Vậy số khối của 3 đồng vị là: A_X = 24, A_Y = 25, A_Z = 26 (các đồng vị ^24^Mg, ^25^Mg, ^26^Mg). (0,5 điểm)\n\n"
            "b) Xác định nguyên tử khối trung bình và khối lượng:\n"
            "- Tỉ lệ số nguyên tử: X : Y : Z = 393 : 50 : 57. Tổng số phần: 393 + 50 + 57 = 500 phần.\n"
            "- Nguyên tử khối trung bình của R (Mg):\n"
            "  A_tb = (24 * 393 + 25 * 50 + 26 * 57) / 500 = (9432 + 1250 + 1482) / 500 = 12164 / 500 = 24,328.\n"
            "- Số mol của 9,033.10^22^ nguyên tử R:\n"
            "  n = (9,033.10^22^) / (6,022.10^23^) = 0,15 mol.\n"
            "- Khối lượng của 9,033.10^22^ nguyên tử R là:\n"
            "  m = n * A_tb = 0,15 * 24,328 = 3,6492 gam (hoặc xấp xỉ 3,65 gam). (0,5 điểm)"
        )
    else:
        add_student_writing_lines(doc, num_lines=12)

    # Chân đề thi
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(10)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "---------- *HẾT* ----------", base_size=10.5, base_bold=True)

    p_footer_sub = doc.add_paragraph()
    p_footer_sub.paragraph_format.space_before = Pt(0)
    p_footer_sub.paragraph_format.space_after = Pt(2)
    p_footer_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_footer_sub, "_Học sinh không được sử dụng tài liệu. Cán bộ coi thi không giải thích gì thêm._", base_size=9.5, base_italic=True)

    return doc

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 1. Bản Học sinh
    doc_hs = build_exam_doc(is_teacher=False)
    hs_path = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_NguyenTruongTo_Ma104_HocSinh.docx")
    doc_hs.save(hs_path)
    print(f"Đã lưu Bản Học sinh: {hs_path}")

    # 2. Bản Giáo viên
    doc_gv = build_exam_doc(is_teacher=True)
    gv_path = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_NguyenTruongTo_Ma104_GiaoVien.docx")
    doc_gv.save(gv_path)
    print(f"Đã lưu Bản Giáo viên: {gv_path}")

if __name__ == "__main__":
    main()
