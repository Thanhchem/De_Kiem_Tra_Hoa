# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì I Hóa Học 10 - THPT Hai Bà Trưng (2024 - 2025)
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_10_HaiBaTrung_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_10_HaiBaTrung_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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

def add_quick_answer_table(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (12 CÂU)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - mỗi câu đúng 0,25 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_1 = [2.4] + [1.3] * 12

    p1_answers = ["D", "B", "D", "C", "D", "C", "D", "B", "B", "C", "D", "D"]

    for row_idx, row in enumerate(tbl1.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for col_idx, cell in enumerate(row.cells):
            cell.width = Cm(col_widths_1[col_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tcPr = cell._tc.get_or_add_tcPr()
            
            borders_xml = f'''<w:tcBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
            </w:tcBorders>'''
            tcPr.append(parse_xml(borders_xml))

            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if row_idx == 0:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8F5E9"/>')
                tcPr.append(shd)
                txt = "Câu" if col_idx == 0 else str(col_idx)
                render_rich_text(p, f"*{txt}*", base_size=10, base_bold=True)
            else:
                txt = "Đáp án" if col_idx == 0 else p1_answers[col_idx - 1]
                if col_idx == 0:
                    render_rich_text(p, f"*{txt}*", base_size=10, base_bold=True)
                else:
                    render_rich_text(p, f"*{txt}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)

    # 2. PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (2 CÂU)
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(6)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm - 1 ý: 0,1đ; 2 ý: 0,25đ; 3 ý: 0,5đ; 4 ý: 1,0đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_2 = [2.4, 2.9, 3.175, 3.175, 3.175, 3.175]
    
    p2_data = [
        ["Câu", "Nội dung", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Cấu tạo nguyên tử", "SAI", "ĐÚNG", "SAI", "SAI"],
        ["Câu 2", "Nguyên tử P (Z=15)", "SAI", "ĐÚNG", "SAI", "ĐÚNG"]
    ]

    for row_idx, row_data in enumerate(p2_data):
        row = tbl2.rows[row_idx]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for col_idx, cell in enumerate(row.cells):
            cell.width = Cm(col_widths_2[col_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tcPr = cell._tc.get_or_add_tcPr()
            
            borders_xml = f'''<w:tcBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
            </w:tcBorders>'''
            tcPr.append(parse_xml(borders_xml))

            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            val = row_data[col_idx]
            if row_idx == 0:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8F5E9"/>')
                tcPr.append(shd)
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True)
            else:
                if col_idx >= 2:
                    color = COLOR_CORRECT if val == "ĐÚNG" else RGBColor(0, 102, 204)
                    render_rich_text(p, f"*{val}*", base_size=10, base_bold=True, base_color=color)
                else:
                    render_rich_text(p, f"*{val}*" if col_idx == 0 else val, base_size=10)

    # 3. PHẦN III: TRẢ LỜI NGẮN (4 CÂU)
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(6)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*3. Phần III: Trả lời ngắn (2,0 điểm - mỗi câu đúng 0,5 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_3 = [2.4, 3.9, 3.9, 3.9, 3.9]
    p3_answers = ["3, 2, 1, 4", "2", "0,1%", "17"]

    for row_idx, row in enumerate(tbl3.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for col_idx, cell in enumerate(row.cells):
            cell.width = Cm(col_widths_3[col_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tcPr = cell._tc.get_or_add_tcPr()
            
            borders_xml = f'''<w:tcBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
            </w:tcBorders>'''
            tcPr.append(parse_xml(borders_xml))

            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if row_idx == 0:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8F5E9"/>')
                tcPr.append(shd)
                txt = "Câu" if col_idx == 0 else f"Câu {col_idx}"
                render_rich_text(p, f"*{txt}*", base_size=10, base_bold=True)
            else:
                txt = "Đáp án" if col_idx == 0 else p3_answers[col_idx - 1]
                if col_idx == 0:
                    render_rich_text(p, f"*{txt}*", base_size=10, base_bold=True)
                else:
                    render_rich_text(p, f"*{txt}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(6)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


def build_exam_docx(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    # Title block
    build_title_block(
        doc,
        school_line="SỞ GIÁO DỤC & ĐÀO TẠO TT HUẾ\nTRƯỜNG THPT HAI BÀ TRƯNG",
        exam_title="ĐỀ KIỂM TRA GIỮA KÌ I - NĂM HỌC 2024 - 2025",
        subject_line="MÔN: HOÁ HỌC - KHỐI 10",
        num_pages="03",
        duration="45 phút",
        exam_code="GỐC",
        atomic_masses="Na = 23; K = 39; Mg = 24; Ca = 40; Ba = 137; Al = 27; Cl = 35,5; H = 1; C = 12; O = 16; N = 14",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table(doc)

    # =========================================================================
    # A. PHẦN TRẮC NGHIỆM
    # =========================================================================
    p_an = doc.add_paragraph()
    p_an.paragraph_format.space_before = Pt(6)
    p_an.paragraph_format.space_after = Pt(2)
    render_rich_text(p_an, "*A. PHẦN TRẮC NGHIỆM: (7,0 điểm)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # -------------------------------------------------------------------------
    # PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (12 CÂU - 3,0 ĐIỂM)
    # -------------------------------------------------------------------------
    add_section_header(
        doc,
        "PHẦN I. (3,0 điểm)",
        "Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # Câu 1
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Nội dung nào dưới đây không phải là đối tượng nghiên cứu của hóa học ?")
    opts_c1 = [
        "A. Thành phần, cấu trúc của chất.",
        "B. Tính chất và sự biến đổi của chất.",
        "C. Ứng dụng của chất.",
        "*D. Sự lớn lên và sinh sản của tế bào.*"
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        sol_c1 = (
            "Đối tượng nghiên cứu của hóa học là chất và sự biến đổi của chất (bao gồm thành phần, cấu trúc, tính chất và ứng dụng của chất).\n"
            "\"Sự lớn lên và sinh sản của tế bào\" là đối tượng nghiên cứu của sinh học."
        )
        add_solution_no_table(doc, sol_c1)

    # Câu 2
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Nguyên tử chứa những hạt mang điện là")
    opts_c2 = [
        "A. proton và neutron.",
        "*B. proton và electron.*",
        "C. electron và neutron.",
        "D. electron và hạt alpha."
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        sol_c2 = (
            "Hạt proton mang điện tích dương (+1).\n"
            "Hạt electron mang điện tích âm (-1).\n"
            "Hạt neutron không mang điện. Do đó, các hạt mang điện trong nguyên tử gồm proton và electron."
        )
        add_solution_no_table(doc, sol_c2)

    # Câu 3
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Chọn đáp án sai")
    opts_c3 = [
        "A. Proton, m ~ 1 amu, q = +1.",
        "B. Electron, m ~ 0,00055 amu, q = -1.",
        "C. Neutron, m ~ 1 amu, q = 0.",
        "*D. Electron, m ~ 1 amu, q = -1.*"
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        sol_c3 = (
            "Khối lượng của hạt electron rất nhỏ, xấp xỉ 0,00055 amu (khoảng 1/1836 amu), thực tế có thể bỏ qua khi tính khối lượng nguyên tử.\n"
            "Nhận định D ghi electron có khối lượng m ~ 1 amu là sai."
        )
        add_solution_no_table(doc, sol_c3)

    # Câu 4
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Thông tin nào sau đây không đúng về ^28^_14_N (hoặc ^28^_14_Si)?")
    opts_c4 = [
        "A. Số khối là 28.",
        "B. Tổng số proton và neutron là 28.",
        "*C. Số neuton là 28.*",
        "D. Số đơn vị điện tích hạt nhân là 14."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        sol_c4 = (
            "Kí hiệu nguyên tử có dạng ^A^_Z_X, trong đó Z = 14 (số đơn vị điện tích hạt nhân) và A = 28 (số khối).\n"
            "Tổng số proton và neutron là số khối: Z + N = A = 28.\n"
            "Số neutron là: N = A - Z = 28 - 14 = 14 (không phải 28). Do đó nhận định C là sai."
        )
        add_solution_no_table(doc, sol_c4)

    # Câu 5
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 5. Nguyên tử của nguyên tố X có 56 electron trong lớp vỏ, 82 neutron trong hạt nhân. Kí hiệu của nguyên tử nguyên tố X là")
    opts_c5 = [
        "A. ^56^_82_A.",
        "B. ^82^_56_A.",
        "C. ^56^_138_A.",
        "*D. ^138^_56_A.*"
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        sol_c5 = (
            "Nguyên tử trung hòa về điện nên số proton Z = số electron = 56.\n"
            "Số khối: A = Z + N = 56 + 82 = 138.\n"
            "Kí hiệu nguyên tử: ^A^_Z_X hay ^138^_56_A (đây là nguyên tố Barium - Ba)."
        )
        add_solution_no_table(doc, sol_c5)

    # Câu 6
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 6. Phát biểu nào sau đây không đúng?")
    opts_c6 = [
        "A. Orbital nguyên tử được kí hiệu là AO (Atomic Orbital).",
        "B. Theo mô hình hiện đại, trong nguyên tử, electron chuyển động rất nhanh, không theo quỹ đạo xác định.",
        "*C. Có 3 loại orbital nguyên tử: orbital s, orbital p và orbital d.*",
        "D. Các orbital s có dạng hình cầu và orbital p có dạng hình số 8 nổi."
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        sol_c6 = (
            "Các loại orbital nguyên tử gồm có orbital s, p, d, f... (không chỉ giới hạn ở 3 loại s, p, d).\n"
            "Do đó nhận định C là phát biểu không đúng."
        )
        add_solution_no_table(doc, sol_c6)

    # Câu 7
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 7. Lớp N có số phân lớp electron bằng")
    opts_c7 = [
        "A. 1.",
        "B. 2.",
        "C. 3.",
        "*D. 4.*"
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        sol_c7 = (
            "Lớp N là lớp electron thứ 4 (n = 4: K, L, M, N).\n"
            "Với n <= 4, lớp thứ n có đúng n phân lớp electron. Vậy lớp N (n = 4) có 4 phân lớp gồm: 4s, 4p, 4d, 4f."
        )
        add_solution_no_table(doc, sol_c7)

    # Câu 8
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 8. Nguyên tử nguyên tố M có số hiệu nguyên tử bằng 20. Cấu hình electron của ion M^2+^ là")
    opts_c8 = [
        "A. 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~ 4s~2~.",
        "*B. 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~.*",
        "C. 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~ 3d~2~ 4s~2~.",
        "D. 1s~2~ 2s~2~ 2p~8~ 3s~2~ 3p~6~."
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        sol_c8 = (
            "Cấu hình electron của nguyên tử M (Z = 20): 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~ 4s~2~ (hay [Ar]4s~2~).\n"
            "Khi tạo ion M^2+^, nguyên tử M nhường 2 electron ở phân lớp ngoài cùng (4s): M -> M^2+^ + 2e.\n"
            "Cấu hình electron của M^2+^ là: 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~ (cấu hình bền vững của khí hiếm Ar)."
        )
        add_solution_no_table(doc, sol_c8)

    # Câu 9
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 9. Chu kì là dãy các nguyên tố được xếp theo chiều điện tích hạt nhân tăng dần, nguyên tử của chúng có cùng")
    opts_c9 = [
        "A. số electron ở lớp ngoài cùng.",
        "*B. số lớp electron.*",
        "C. số electron hóa trị.",
        "D. số electron."
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        sol_c9 = (
            "Theo định nghĩa trong bảng tuần hoàn: Chu kì là dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được sắp xếp theo chiều điện tích hạt nhân tăng dần."
        )
        add_solution_no_table(doc, sol_c9)

    # Câu 10
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 10. Nguyên tố Y thuộc chu kì 4, nhóm IA của bảng tuần hoàn. Phát biểu nào sau đây về Y là đúng ?")
    opts_c10 = [
        "A. Y có độ âm điện lớn nhất và bán kính nguyên tử nhỏ nhất trong chu kì 4.",
        "B. Y có độ âm điện lớn nhất và bán kính nguyên tử lớn nhất trong chu kì 4.",
        "*C. Y có độ âm điện nhỏ nhất và bán kính nguyên tử lớn nhất trong chu kì 4.*",
        "D. Y có độ âm điện nhỏ nhất và bán kính nguyên tử nhỏ nhất trong chu kì 4."
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        sol_c10 = (
            "Trong cùng một chu kì, đi từ trái sang phải theo chiều tăng của điện tích hạt nhân: bán kính nguyên tử giảm dần, độ âm điện tăng dần.\n"
            "Nguyên tố Y thuộc nhóm IA đứng đầu chu kì 4 nên có độ âm điện nhỏ nhất và bán kính nguyên tử lớn nhất trong chu kì 4."
        )
        add_solution_no_table(doc, sol_c10)

    # Câu 11
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 11. Trong một chu kì, bán kính nguyên tử")
    opts_c11 = [
        "A. xấp xỉ bằng nhau.",
        "B. biến đổi ngẫu nhiên.",
        "C. tăng theo chiều tăng dần của điện tích hạt nhân.",
        "*D. giảm theo chiều tăng dần của điện tích hạt nhân.*"
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        sol_c11 = (
            "Trong một chu kì, đi từ trái sang phải theo chiều tăng của điện tích hạt nhân, số lớp electron không đổi nhưng điện tích hạt nhân tăng làm tăng lực hút tĩnh điện giữa hạt nhân với electron, dẫn đến bán kính nguyên tử giảm dần."
        )
        add_solution_no_table(doc, sol_c11)

    # Câu 12
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 12. Cho các nguyên tố sau: Li (Z=3), O (Z=8), F (Z=9), Na (Z=11). Dãy gồm các nguyên tố được sắp xếp theo chiều tăng dần bán kính nguyên tử từ trái sang phải là")
    opts_c12 = [
        "A. F, Li, O, Na.",
        "B. Li, Na, O, F.",
        "C. F, Na, O, Li.",
        "*D. F, O, Li, Na.*"
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        sol_c12 = (
            "Các nguyên tố Li (Z=3), O (Z=8), F (Z=9) cùng thuộc chu kì 2. Trong một chu kì, bán kính giảm dần từ trái sang phải: r_F < r_O < r_Li.\n"
            "Li (Z=3) và Na (Z=11) cùng thuộc nhóm IA. Trong một nhóm A, bán kính tăng dần từ trên xuống dưới: r_Li < r_Na.\n"
            "Thứ tự tăng dần bán kính nguyên tử: r_F < r_O < r_Li < r_Na (tức F, O, Li, Na)."
        )
        add_solution_no_table(doc, sol_c12)

    # -------------------------------------------------------------------------
    # PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (2 CÂU - 2,0 ĐIỂM)
    # -------------------------------------------------------------------------
    add_section_header(
        doc,
        "PHẦN II. (2,0 điểm)",
        "Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng (Đ) hoặc sai (S)."
    )

    # Câu 1 Phần II
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Các phát biểu sau đây là đúng hay sai?")
    stmts_c1 = [
        ("a) Nguyên tử của nguyên tố hydrogen luôn có 2 proton.", "SAI"),
        ("b) Nguyên tử của nguyên tố hydrogen có thể không có neutron.", "ĐÚNG"),
        ("c) Nếu đường kính của nguyên tử khoảng 10^2^ pm thì đường kính của hạt nhân khoảng 10^-4^ pm.", "SAI"),
        ("d) Nguyên tố hóa học là tập hợp các nguyên tử có cùng số neutron trong hạt nhân.", "SAI")
    ]
    for stmt_text, ans_str in stmts_c1:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.left_indent = Cm(0.5)
        p_stmt.paragraph_format.space_before = Pt(1)
        p_stmt.paragraph_format.space_after = Pt(1)
        render_statement(p_stmt, stmt_text, ans=ans_str if is_teacher else None, is_teacher=is_teacher)

    if is_teacher:
        sol_p2_c1 = (
            "a) SAI: Nguyên tử hydrogen có số hiệu nguyên tử Z = 1 nên luôn chỉ có 1 proton trong hạt nhân.\n"
            "b) ĐÚNG: Đồng vị phổ biến nhất của hydrogen là protium (^1^_1_H) có số khối A = 1, suy ra số neutron N = 1 - 1 = 0 (không có neutron).\n"
            "c) SAI: Kích thước hạt nhân nhỏ hơn kích thước nguyên tử khoảng 10^4^ lần (10.000 lần). Nếu đường kính nguyên tử khoảng 10^2^ pm thì đường kính hạt nhân vào khoảng 10^2^ x 10^-4^ = 10^-2^ pm (chứ không phải 10^-4^ pm).\n"
            "d) SAI: Nguyên tố hóa học là tập hợp các nguyên tử có cùng số đơn vị điện tích hạt nhân (cùng số proton), không phải cùng số neutron."
        )
        add_solution_no_table(doc, sol_p2_c1)

    # Câu 2 Phần II
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Cho sơ đồ của một nguyên tử X được biểu diễn như hình bên:")
    
    # Chèn ảnh sơ đồ nguyên tử P
    img_path = os.path.join(r"C:\Antigravity_Thanh\San_Pham\images_hoa10", "so_do_nguyen_tu_p_15.png")
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(3)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=Cm(4.5))

    stmts_c2 = [
        ("a) Các electron trong nguyên tử X chỉ được phân bố trên 2 lớp K và lớp L.", "SAI"),
        ("b) Ở trạng thái cơ bản, nguyên tử X có 3 electron độc thân.", "ĐÚNG"),
        ("c) Nguyên tử nguyên tố X có tổng số hạt cơ bản là 50 hạt. Số hạt mang điện nhiều hơn số hạt không mang điện là 10 hạt. Số hạt neutron của X là 35.", "SAI"),
        ("d) X thuộc ô 15, chu kỳ 3, nhóm VA.", "ĐÚNG")
    ]
    for stmt_text, ans_str in stmts_c2:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.left_indent = Cm(0.5)
        p_stmt.paragraph_format.space_before = Pt(1)
        p_stmt.paragraph_format.space_after = Pt(1)
        render_statement(p_stmt, stmt_text, ans=ans_str if is_teacher else None, is_teacher=is_teacher)

    if is_teacher:
        sol_p2_c2 = (
            "a) SAI: Theo sơ đồ, electron phân bố trên 3 lớp: lớp 1 (K có 2e), lớp 2 (L có 8e), lớp 3 (M có 5e).\n"
            "b) ĐÚNG: Cấu hình electron của X (Z = 15): 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~3~. Phân lớp 3p có 3 orbital, theo quy tắc Hund mỗi orbital có 1 electron độc thân -> có 3 electron độc thân.\n"
            "c) SAI: Ta có hệ phương trình: 2Z + N = 50 và 2Z - N = 10 -> 4Z = 60 -> Z = 15; suy ra N = 50 - 2 x 15 = 20 (chứ không phải 35).\n"
            "d) ĐÚNG: Z = 15 (ô 15); có 3 lớp e (chu kì 3); có 5e lớp ngoài cùng và là nguyên tố p -> nhóm VA (nguyên tố Phosphorus - P)."
        )
        add_solution_no_table(doc, sol_p2_c2)

    # -------------------------------------------------------------------------
    # PHẦN III: TRẢ LỜI NGẮN (4 CÂU - 2,0 ĐIỂM)
    # -------------------------------------------------------------------------
    add_section_header(
        doc,
        "PHẦN III. (2,0 điểm)",
        "Câu trắc nghiệm yêu cầu trả lời ngắn. Thí sinh trả lời từ câu 1 đến câu 4."
    )

    # Câu 1 Phần III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Cho dãy nguyên tố sau: Na (1), K (2), Rb (3), Li (4). Hãy sắp xếp theo chiều giảm dần bán kính nguyên tử? (ghi thứ tự các số)")
    if not is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* .....................................................................................................................", base_size=10.5)
    else:
        sol_p3_c1 = (
            "Các nguyên tố Li, Na, K, Rb đều thuộc nhóm IA.\n"
            "Trong cùng một nhóm A, theo chiều tăng của điện tích hạt nhân (từ trên xuống dưới: Li -> Na -> K -> Rb), số lớp electron tăng dần nên bán kính nguyên tử tăng dần: r_Li < r_Na < r_K < r_Rb.\n"
            "Sắp xếp theo chiều GIẢM DẦN bán kính nguyên tử: Rb > K > Na > Li, tương ứng dãy số: 3, 2, 1, 4."
        )
        add_solution_no_table(doc, sol_p3_c1)

    # Câu 2 Phần III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Cho số hiệu nguyên tử của Li (Z=3), Na (Z=11), P (Z=15), Ar (Z=18). Có bao nhiêu nguyên tố p trong các nguyên tố trên?")
    if not is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* .....................................................................................................................", base_size=10.5)
    else:
        sol_p3_c2 = (
            "Cấu hình electron của từng nguyên tố:\n"
            "+ Li (Z = 3): 1s~2~ 2s~1~ -> electron cuối cùng điền vào phân lớp s -> nguyên tố s.\n"
            "+ Na (Z = 11): [Ne] 3s~1~ -> electron cuối cùng điền vào phân lớp s -> nguyên tố s.\n"
            "+ P (Z = 15): [Ne] 3s~2~ 3p~3~ -> electron cuối cùng điền vào phân lớp p -> nguyên tố p.\n"
            "+ Ar (Z = 18): [Ne] 3s~2~ 3p~6~ -> electron cuối cùng điền vào phân lớp p -> nguyên tố p.\n"
            "Vậy có 2 nguyên tố p là P và Ar."
        )
        add_solution_no_table(doc, sol_p3_c2)

    # Câu 3 Phần III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Đồng vị phóng xạ cobalt (^60^_27_Co) phát ra tia gamma có khả năng đâm xuyên mạnh, dùng điều trị các khối u ở sâu trong cơ thể. Cobalt có ba đồng vị: ^59^_27_Co (chiếm 98%), ^58^_27_Co và ^60^_27_Co; nguyên tử khối trung bình là 58,982. Xác định hàm lượng % của đồng vị phóng xạ ^60^_27_Co.")
    if not is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* .....................................................................................................................", base_size=10.5)
    else:
        sol_p3_c3 = (
            "Gọi % số nguyên tử của ^58^_27_Co là x% và của ^60^_27_Co là y%.\n"
            "Tổng % các đồng vị: 98% + x% + y% = 100% -> x + y = 2 -> x = 2 - y.\n"
            "Nguyên tử khối trung bình: A_tb = [59 x 98 + 58 x x + 60 x y] / 100 = 58,982\n"
            "-> 5782 + 58(2 - y) + 60y = 5898,2\n"
            "-> 5782 + 116 + 2y = 5898,2 -> 5898 + 2y = 5898,2\n"
            "-> 2y = 0,2 -> y = 0,1%.\n"
            "Vậy hàm lượng % của đồng vị phóng xạ ^60^_27_Co là 0,1%."
        )
        add_solution_no_table(doc, sol_p3_c3)

    # Câu 4 Phần III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Nguyên tử X có tổng số hạt cơ bản là 49, số hạt không mang điện bằng 53,125% số hạt mang điện. Số neutron của nguyên tử X là bao nhiêu?")
    if not is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* .....................................................................................................................", base_size=10.5)
    else:
        sol_p3_c4 = (
            "Tổng số hạt cơ bản: 2Z + N = 49 (1).\n"
            "Số hạt không mang điện (N) bằng 53,125% số hạt mang điện (2Z):\n"
            "N = 0,53125 x 2Z = 1,0625 Z (2).\n"
            "Thay (2) vào (1): 2Z + 1,0625 Z = 49 -> 3,0625 Z = 49 -> Z = 16.\n"
            "Số neutron của X: N = 49 - 2 x 16 = 17 (X là nguyên tố Sulfur - S)."
        )
        add_solution_no_table(doc, sol_p3_c4)

    # =========================================================================
    # B. PHẦN TỰ LUẬN
    # =========================================================================
    p_tl = doc.add_paragraph()
    p_tl.paragraph_format.space_before = Pt(8)
    p_tl.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl, "*B. PHẦN TỰ LUẬN: (3,0 điểm)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # Câu 1 Tự luận
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1 (1,0 điểm): Cho nguyên tố A có số hiệu nguyên tử Z = 13.")
    p_a = doc.add_paragraph()
    p_a.paragraph_format.left_indent = Cm(0.5)
    render_rich_text(p_a, "a) Viết cấu hình electron nguyên tử của nguyên tố A. Cho biết A là kim loại, phi kim hay khí hiếm?", base_size=10.5)
    p_b = doc.add_paragraph()
    p_b.paragraph_format.left_indent = Cm(0.5)
    render_rich_text(p_b, "b) Biểu diễn cấu hình electron của A theo ô orbital. Cho biết số electron độc thân của nguyên tử A?", base_size=10.5)

    if not is_teacher:
        add_student_writing_lines(doc, num_lines=7)
    else:
        sol_tl1 = (
            "a) Cấu hình electron của nguyên tử A (Z = 13): 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~1~ (hoặc [Ne] 3s~2~ 3p~1~).\n"
            "   Nguyên tử A có 3 electron ở lớp ngoài cùng (lớp M, n = 3) nên A là nguyên tố kim loại (Aluminium - Al).\n"
            "b) Biểu diễn cấu hình electron theo ô orbital:\n"
            "   + Lớp 1 (1s): 1 ô chứa 2 electron ghép đôi ([UP-DOWN]).\n"
            "   + Lớp 2 (2s, 2p): ô 2s chứa 2 electron ghép đôi ([UP-DOWN]); 3 ô 2p chứa 6 electron ghép đôi ([UP-DOWN][UP-DOWN][UP-DOWN]).\n"
            "   + Lớp 3 (3s, 3p): ô 3s chứa 2 electron ghép đôi ([UP-DOWN]); 3 ô 3p chứa 1 electron độc thân ([UP][ ][ ]).\n"
            "   Kết luận: Số electron độc thân của nguyên tử A là 1 electron."
        )
        add_solution_no_table(doc, sol_tl1)

    # Câu 2 Tự luận
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2 (1,0 điểm): Kim loại M thuộc nhóm IIA của bảng tuần hoàn, là một thành phần dinh dưỡng quan trọng. Sự thiếu hụt rất nhỏ của nó đã ảnh hưởng tới sự hình thành và phát triển của xương, răng. Thừa M có thể dẫn đến sỏi thận. Cho 1,2 gam M tác dụng hết với dung dịch HCl, thu được 0,7437 L khí (đkc).")
    p_a2 = doc.add_paragraph()
    p_a2.paragraph_format.left_indent = Cm(0.5)
    render_rich_text(p_a2, "a) Xác định M và viết cấu hình electron của ion M^2+^.", base_size=10.5)
    p_b2 = doc.add_paragraph()
    p_b2.paragraph_format.left_indent = Cm(0.5)
    render_rich_text(p_b2, "b) Sắp xếp theo thứ tự tăng dần về bán kính nguyên tử của nguyên tố M với _19_K và _12_Mg.", base_size=10.5)

    if not is_teacher:
        add_student_writing_lines(doc, num_lines=8)
    else:
        sol_tl2 = (
            "a) Xác định M:\n"
            "   + Số mol khí H~2~ ở điều kiện chuẩn (đkc, 25 °C, 1 bar): n_H2 = 0,7437 / 24,79 = 0,03 mol.\n"
            "   + Kim loại M thuộc nhóm IIA nên có hóa trị II. Phương trình phản ứng: M + 2HCl -> MCl~2~ + H~2~\n"
            "   + Theo phương trình: n_M = n_H2 = 0,03 mol.\n"
            "   + Khối lượng mol của M: M_M = 1,2 / 0,03 = 40 (g/mol). Vậy M là kim loại Calcium (Ca, Z = 20).\n"
            "   + Cấu hình electron của Ca (Z = 20): 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~ 4s~2~ hay [Ar]4s~2~.\n"
            "   + Cấu hình electron của ion Ca^2+^: 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~ hay [Ar].\n"
            "b) Sắp xếp bán kính nguyên tử giữa Mg (Z = 12), Ca (Z = 20), K (Z = 19):\n"
            "   + Mg và Ca cùng thuộc nhóm IIA: Trong nhóm A, từ trên xuống dưới số lớp e tăng -> bán kính tăng: r_Mg < r_Ca.\n"
            "   + K và Ca cùng thuộc chu kì 4: Trong chu kì, từ trái sang phải Z tăng -> bán kính giảm: r_Ca < r_K.\n"
            "   + Thứ tự tăng dần bán kính nguyên tử: r_Mg < r_Ca < r_K (Mg < Ca < K)."
        )
        add_solution_no_table(doc, sol_tl2)

    # Câu 3 Tự luận
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3 (1,0 điểm): Trong tự nhiên nguyên tố X có hai đồng vị là ^35^_17_X và đồng vị thứ hai ^37^_17_X. Biết nguyên tử khối trung bình của X là 35,48.")
    p_a3 = doc.add_paragraph()
    p_a3.paragraph_format.left_indent = Cm(0.5)
    render_rich_text(p_a3, "a) Xác định vị trí của X trong bảng tuần hoàn (ô, chu kì, nhóm)?", base_size=10.5)
    p_b3 = doc.add_paragraph()
    p_b3.paragraph_format.left_indent = Cm(0.5)
    render_rich_text(p_b3, "b) Tính % số nguyên tử của đồng vị ^35^_17_X.", base_size=10.5)

    if not is_teacher:
        add_student_writing_lines(doc, num_lines=7)
    else:
        sol_tl3 = (
            "a) Vị trí của X trong bảng tuần hoàn:\n"
            "   + Z = 17 -> Cấu hình electron của X: 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~5~ (hoặc [Ne] 3s~2~ 3p~5~).\n"
            "   + Ô nguyên tố: 17 (do số hiệu nguyên tử Z = 17).\n"
            "   + Chu kì: 3 (do có 3 lớp electron).\n"
            "   + Nhóm: VIIA (do có 7 electron lớp ngoài cùng và là nguyên tố p).\n"
            "   + (X là nguyên tố Chlorine - Cl).\n"
            "b) Tính % số nguyên tử của đồng vị ^35^_17_X:\n"
            "   + Gọi phần trăm số nguyên tử của đồng vị ^35^_17_X là a%, suy ra % của ^37^_17_X là (100 - a)%.\n"
            "   + Áp dụng công thức nguyên tử khối trung bình:\n"
            "     A_tb = [35 x a + 37 x (100 - a)] / 100 = 35,48\n"
            "     <-> 35a + 3700 - 37a = 3548\n"
            "     <-> -2a = 3548 - 3700 = -152  ->  a = 76%.\n"
            "   + Vậy hàm lượng % số nguyên tử của đồng vị ^35^_17_X là 76% (và của ^37^_17_X là 24%)."
        )
        add_solution_no_table(doc, sol_tl3)

    # Footer note
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(10)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "*---------- HẾT ----------*", base_size=10.5, base_bold=True)
    p_end_note = doc.add_paragraph()
    p_end_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end_note, "_(Thí sinh không được sử dụng tài liệu. Cán bộ coi thi không giải thích gì thêm)_", base_size=9.5, base_italic=True)

    filename = "De_Kiem_Tra_GK1_Hoa_10_HaiBaTrung_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_GK1_Hoa_10_HaiBaTrung_HocSinh.docx"
    output_path = os.path.join(OUTPUT_DIR, filename)
    doc.save(output_path)
    print(f"[OK] Đã tạo thành công: {output_path}")

def main():
    print("Đang tạo Bản Học sinh...")
    build_exam_docx(is_teacher=False)
    print("Đang tạo Bản Giáo viên...")
    build_exam_docx(is_teacher=True)
    print("[HOÀN TẤT] Cả 2 bản Word đã được tạo thành công!")

if __name__ == "__main__":
    main()
