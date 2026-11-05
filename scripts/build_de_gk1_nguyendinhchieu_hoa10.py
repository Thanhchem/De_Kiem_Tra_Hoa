# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì 1 Hóa Học 10 - THPT Nguyễn Đình Chiểu - Huế
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_10_NguyenDinhChieu.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_10_NguyenDinhChieu_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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

def add_quick_answer_table_ndc(doc):
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
    p1_answers = ["A", "B", "A", "A", "C", "D", "B", "B", "D", "C", "A", "A"]

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

    # 2. PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (3 CÂU)
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(4)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*2. Phần II: Trắc nghiệm Đúng / Sai (3,0 điểm - 1 ý: 0,1đ; 2 ý: 0,25đ; 3 ý: 0,5đ; 4 ý: 1,0đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=4, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_2 = [2.4, 3.6, 3.0, 3.0, 3.0, 3.0]
    
    p2_data = [
        ["Câu", "Nội dung", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Đồng vị Cl và Cu", "ĐÚNG", "ĐÚNG", "ĐÚNG", "SAI"],
        ["Câu 2", "Nguyên tố chu kì 3", "ĐÚNG", "SAI", "ĐÚNG", "ĐÚNG"],
        ["Câu 3", "Cấu hình e & AO", "ĐÚNG", "ĐÚNG", "SAI", "ĐÚNG"]
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
    render_rich_text(p_p3, "*3. Phần III: Trả lời ngắn (1,0 điểm - mỗi câu đúng 0,25 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_3 = [2.4, 3.9, 3.9, 3.9, 3.9]
    p3_answers = ["137", "124", "11", "2"]

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

    # 4. TÓM TẮT ĐÁP SỐ TỰ LUẬN
    p_tl = doc.add_paragraph()
    p_tl.paragraph_format.space_before = Pt(6)
    p_tl.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl, "*4. Tóm tắt đáp số Phần Tự luận (3,0 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl_tl = doc.add_table(rows=4, cols=3)
    tbl_tl.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_tl = [2.4, 2.6, 13.0]
    
    tl_summary = [
        ["Câu hỏi", "Điểm số", "Đáp số cốt lõi"],
        ["Câu 1 (TL)", "1,0 điểm", "a) p = e = 12; n = 12.  b) Cấu hình: 1s^2 2s^2 2p^6 3s^2; là kim loại."],
        ["Câu 2 (TL)", "1,0 điểm", "a) X (Z=11, CK 3, nhóm IA); Y (Z=12, CK 3, nhóm IIA).  b) Base: NaOH > Mg(OH)~2~."],
        ["Câu 3 (TL)", "1,0 điểm", "a) Tên: Sulfur (Lưu huỳnh), KHHH: S.  b) Giá trị m = 9,8 gam acid (H~2~SO~4~)."]
    ]

    for row_idx, row_data in enumerate(tl_summary):
        row = tbl_tl.rows[row_idx]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for col_idx, cell in enumerate(row.cells):
            cell.width = Cm(col_widths_tl[col_idx])
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
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx < 2 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            val = row_data[col_idx]
            if row_idx == 0:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8F5E9"/>')
                tcPr.append(shd)
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True)
            else:
                if col_idx == 0:
                    render_rich_text(p, f"*{val}*", base_size=10, base_bold=True)
                elif col_idx == 1:
                    render_rich_text(p, f"_{val}_", base_size=9.5, base_italic=True, base_color=COLOR_PRIMARY)
                else:
                    render_rich_text(p, val, base_size=10)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(8)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


def build_exam_ndc(is_teacher=False, output_path=None):
    doc = docx.Document()
    setup_header_footer(doc)

    # 1. Khối tiêu đề đầu đề thi
    build_title_block(
        doc,
        school_line="SỞ GD&ĐT THỪA THIÊN HUẾ\nTRƯỜNG THPT NGUYỄN ĐÌNH CHIỂU",
        exam_title="KIỂM TRA GIỮA HỌC KỲ I",
        subject_line="MÔN: HOÁ HỌC - LỚP 10",
        num_pages="02",
        duration="45 phút",
        exam_code="ĐỀ GỐC",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Cr = 24; Fe = 55,85; Cu = 64; Br = 80; Ba = 137. Số Avogadro N~A~ = 6,022 × 10^23^ mol^-1.",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_ndc(doc)

    # =========================================================================
    # A. PHẦN TRẮC NGHIỆM: (7,0 điểm)
    # =========================================================================
    p_sec_a = doc.add_paragraph()
    p_sec_a.paragraph_format.space_before = Pt(6)
    p_sec_a.paragraph_format.space_after = Pt(2)
    render_rich_text(p_sec_a, "*A. PHẦN TRẮC NGHIỆM (7 điểm)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # PHẦN I
    add_section_header(
        doc,
        "PHẦN I (3 điểm). CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn 1 phương án."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Nguyên tử của nguyên tố X có tổng số hạt cơ bản (p, n, e) là 52, trong đó số hạt mang điện nhiều hơn số hạt không mang điện là 16. Số khối của nguyên tử X là")
    opts1 = [
        "*A. 35.*" if is_teacher else "A. 35.",
        "B. 36.",
        "C. 34.",
        "D. 37."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Gọi số proton, neutron, electron của nguyên tử X lần lượt là p, n, e (với p = e = Z).\nTheo đề bài ta có hệ phương trình:\n- Tổng số hạt cơ bản: 2Z + N = 52  (1)\n- Số hạt mang điện nhiều hơn số hạt không mang điện: 2Z - N = 16  (2)\nCộng (1) và (2) theo vế: 4Z = 68 => Z = 17.\nThay vào (1): N = 52 - 2 × 17 = 18.\nSố khối của nguyên tử X: A = Z + N = 17 + 18 = 35 (nguyên tố Chlorine: ^35_17Cl).\n➔ Đáp án đúng: A.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Hợp chất khí với hydrogen của nguyên tố R có dạng RH~3~. Trong oxide cao nhất của R, oxygen chiếm 74,07% về khối lượng. Phát biểu nào sau đây sai?")
    opts2 = [
        "A. Nguyên tố R thuộc chu kì 2, nhóm VA.",
        "*B. Phân tử oxide cao nhất của R có công thức R~2~O~3~.*" if is_teacher else "B. Phân tử oxide cao nhất của R có công thức R~2~O~3~.",
        "C. Độ âm điện của R lớn hơn của phosphorus.",
        "D. Hydroxide ứng với hóa trị cao nhất của R là HNO~3~."
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "- Hợp chất khí với hydrogen có dạng RH~3~ => Hóa trị của R với hydrogen là 3.\n  Suy ra R thuộc nhóm VA, hóa trị cao nhất của R với oxygen là 8 - 3 = 5.\n  Công thức oxide cao nhất của R là R~2~O~5~.\n- Trong R~2~O~5~, oxygen chiếm 74,07% về khối lượng:\n  %O = (16 × 5) / (2M_R + 16 × 5) × 100% = 74,07%\n  => 80 / (2M_R + 80) = 0,7407 => 2M_R + 80 ≈ 108,006 => M_R = 14 (Nitrogen - N).\n- Xét các phát biểu:\n  + A đúng: N có Z = 7 (1s^2 2s^2 2p^3) thuộc chu kì 2, nhóm VA.\n  + B sai: Oxide cao nhất có công thức R~2~O~5~ (N~2~O~5~), không phải R~2~O~3~.\n  + C đúng: N và P cùng thuộc nhóm VA, trong một nhóm A đi từ trên xuống dưới độ âm điện giảm dần => độ âm điện của N lớn hơn P.\n  + D đúng: Hydroxide ứng với hóa trị cao nhất của N là HNO~3~.\n➔ Đáp án đúng: B.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Nguyên tử của nguyên tố chromium (Cr) có số hiệu nguyên tử Z = 24. Cấu hình electron ở trạng thái cơ bản của nguyên tử Cr là")
    opts3 = [
        "*A. [Ar] 3d^5^ 4s^1^.*" if is_teacher else "A. [Ar] 3d^5^ 4s^1^.",
        "B. [Ar] 3d^5^ 4s^2^.",
        "C. [Ar] 3d^4^ 4s^2^.",
        "D. [Ar] 3d^6^."
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo thứ tự mức năng lượng: 1s^2 2s^2 2p^6 3s^2 3p^6 4s^2 3d^4.\nDo cấu hình bán bão hòa phân lớp d (d^5) bền vững hơn nên 1 electron từ phân lớp 4s chuyển sang phân lớp 3d.\nCấu hình electron ở trạng thái cơ bản của Cr (Z = 24) là: 1s^2 2s^2 2p^6 3s^2 3p^6 3d^5 4s^1 hay [Ar] 3d^5 4s^1.\n➔ Đáp án đúng: A.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Ở trạng thái cơ bản, phân lớp 3d chứa tối đa bao nhiêu electron?")
    opts4 = [
        "*A. 10.*" if is_teacher else "A. 10.",
        "B. 6.",
        "C. 18.",
        "D. 14."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân lớp d có 5 orbital (AO). Mỗi orbital chứa tối đa 2 electron có spin trái dấu.\nDo đó, phân lớp d nói chung và phân lớp 3d nói riêng chứa tối đa 2 × 5 = 10 electron.\n➔ Đáp án đúng: A.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 5. Cho các nguyên tố: C (Z = 6), N (Z = 7), Si (Z = 14), P (Z = 15). Chiều bán kính nguyên tử tăng dần là")
    opts5 = [
        "A. Si < P < C < N.",
        "B. N < P < C < Si.",
        "*C. N < C < P < Si.*" if is_teacher else "C. N < C < P < Si.",
        "D. C < N < Si < P."
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Vị trí trong bảng tuần hoàn:\n- Chu kì 2: C (Z = 6, nhóm IVA), N (Z = 7, nhóm VA).\n- Chu kì 3: Si (Z = 14, nhóm IVA), P (Z = 15, nhóm VA).\nQuy luật biến đổi bán kính nguyên tử:\n- Trong cùng một chu kì, bán kính giảm theo chiều tăng của Z: N < C và P < Si.\n- Trong cùng một nhóm A, bán kính tăng theo chiều tăng của Z: C < Si và N < P.\nKết hợp lại ta có thứ tự bán kính tăng dần: N < C < P < Si.\n➔ Đáp án đúng: C.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 6. Cho các nguyên tố Na (Z = 11), Mg (Z = 12), Al (Z = 13), K (Z = 19). Chiều tính base của các hydroxide tăng dần là")
    opts6 = [
        "A. Al(OH)~3~ < Mg(OH)~2~ < KOH < NaOH.",
        "B. Mg(OH)~2~ < Al(OH)~3~ < NaOH < KOH.",
        "C. KOH < NaOH < Mg(OH)~2~ < Al(OH)~3~.",
        "*D. Al(OH)~3~ < Mg(OH)~2~ < NaOH < KOH.*" if is_teacher else "D. Al(OH)~3~ < Mg(OH)~2~ < NaOH < KOH."
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "- Na (Z = 11), Mg (Z = 12), Al (Z = 13) cùng thuộc chu kì 3. Trong cùng chu kì, theo chiều tăng của điện tích hạt nhân, tính kim loại giảm dần => Tính base của các hydroxide giảm dần: NaOH > Mg(OH)~2~ > Al(OH)~3~.\n- K (Z = 19) và Na (Z = 11) cùng thuộc nhóm IA, K ở chu kì 4 dưới Na chu kì 3. Trong cùng một nhóm A, tính kim loại tăng dần từ trên xuống dưới => Tính base: KOH > NaOH.\nVậy chiều tính base tăng dần là: Al(OH)~3~ < Mg(OH)~2~ < NaOH < KOH.\n➔ Đáp án đúng: D.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 7. Trong tự nhiên, nguyên tố bromine (Br) có hai đồng vị bền là ^79^Br và ^81^Br. Biết nguyên tử khối trung bình của bromine là 79,91. Phần trăm số nguyên tử của đồng vị ^79^Br là")
    opts7 = [
        "A. 49,5%.",
        "*B. 54,5%.*" if is_teacher else "B. 54,5%.",
        "C. 50,5%.",
        "D. 45,5%."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Gọi phần trăm số nguyên tử của đồng vị ^79^Br là x (%) (0 < x < 100).\nKhi đó phần trăm số nguyên tử của đồng vị ^81^Br là (100 - x) (%).\nTheo công thức nguyên tử khối trung bình:\nA_tb = [79x + 81(100 - x)] / 100 = 79,91\n<=> 79x + 8100 - 81x = 7991\n<=> -2x = -109  <=>  x = 54,5%.\nVậy phần trăm số nguyên tử của đồng vị ^79^Br là 54,5%.\n➔ Đáp án đúng: B.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 8. Theo chiều tăng dần của điện tích hạt nhân trong một chu kì, xu hướng nào sau đây là đúng đối với các nguyên tố nhóm A?")
    opts8 = [
        "A. Độ âm điện giảm dần.",
        "*B. Tính kim loại giảm dần, tính phi kim tăng dần.*" if is_teacher else "B. Tính kim loại giảm dần, tính phi kim tăng dần.",
        "C. Bán kính nguyên tử tăng dần.",
        "D. Tính base tăng dần."
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong một chu kì, theo chiều tăng dần của điện tích hạt nhân (từ trái sang phải):\n- Bán kính nguyên tử giảm dần.\n- Độ âm điện tăng dần.\n- Tính kim loại giảm dần, tính phi kim tăng dần.\n- Tính base của oxide và hydroxide giảm dần, tính acid tăng dần.\nVậy phát biểu đúng là B.\n➔ Đáp án đúng: B.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 9. Nguyên tử của nguyên tố Y ở trạng thái cơ bản có tổng số electron ở các phân lớp p là 11. Vị trí của Y trong bảng tuần hoàn là")
    opts9 = [
        "A. chu kì 3, nhóm VIA.",
        "B. chu kì 4, nhóm IA.",
        "C. chu kì 3, nhóm VA.",
        "*D. chu kì 3, nhóm VIIA.*" if is_teacher else "D. chu kì 3, nhóm VIIA."
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Tổng số electron ở các phân lớp p là 11:\n- Phân lớp 2p bão hòa chứa 6 electron: 2p^6^.\n- Phân lớp 3p chứa 11 - 6 = 5 electron: 3p^5^.\nCấu hình electron đầy đủ của Y: 1s^2 2s^2 2p^6 3s^2 3p^5 (Z = 17 - Chlorine).\n- Có 3 lớp electron => Y thuộc chu kì 3.\n- Lớp ngoài cùng (lớp 3) có 2 + 5 = 7 electron, e cuối cùng điền vào phân lớp p => Y thuộc nhóm VIIA.\n➔ Đáp án đúng: D.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 10. Hạt nhân của hầu hết các nguyên tử được cấu tạo từ các loại hạt nào sau đây?")
    opts10 = [
        "A. Electron và neutron.",
        "B. Proton và electron.",
        "*C. Proton và neutron.*" if is_teacher else "C. Proton và neutron.",
        "D. Proton, neutron và electron."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hạt nhân của hầu hết các nguyên tử được cấu tạo từ hai loại hạt cơ bản là proton (mang điện tích dương) và neutron (không mang điện). Duy nhất đồng vị ^1_1H có hạt nhân chỉ chứa 1 proton, không có neutron. Electron nằm ở lớp vỏ bao quanh hạt nhân.\n➔ Đáp án đúng: C.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 11. Trong bảng tuần hoàn các nguyên tố hóa học, số thứ tự của nhóm A bằng")
    opts11 = [
        "*A. số electron ở lớp ngoài cùng.*" if is_teacher else "A. số electron ở lớp ngoài cùng.",
        "B. tổng số hạt proton và neutron.",
        "C. số electron ở phân lớp ngoài cùng.",
        "D. số lớp electron trong nguyên tử."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo quy tắc cấu tạo bảng tuần hoàn:\n- Đối với các nguyên tố nhóm A, số thứ tự nhóm (I A -> VIIIA) bằng số electron ở lớp ngoài cùng (cũng chính là số electron hóa trị).\n- Số lớp electron trong nguyên tử quyết định số thứ tự chu kì.\n➔ Đáp án đúng: A.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 12. Cặp kí hiệu nguyên tử nào sau đây là đồng vị của cùng một nguyên tố hóa học?")
    opts12 = [
        "*A. ^35_17X và ^37_17Y.*" if is_teacher else "A. ^35_17X và ^37_17Y.",
        "B. ^40_19X và ^40_20Y.",
        "C. ^16_8X và ^14_7Y.",
        "D. ^23_11X và ^24_12Y."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Đồng vị là các nguyên tử của cùng một nguyên tố hóa học, có cùng số proton (cùng số hiệu nguyên tử Z) nhưng có số neutron khác nhau (dẫn tới số khối A khác nhau).\nTrong các cặp đã cho, cặp ^35_17X và ^37_17Y đều có Z = 17, số khối A lần lượt là 35 và 37 => là đồng vị của cùng một nguyên tố (Chlorine).\n➔ Đáp án đúng: A.")

    # =========================================================================
    # PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN II (3 điểm). CÂU TRẮC NGHIỆM ĐÚNG / SAI",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 3. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Trong tự nhiên, nguyên tố chlorine (Cl) có hai đồng vị bền là ^35^Cl (75,77%) và ^37^Cl (24,23%). Nguyên tố copper (Cu) có hai đồng vị bền là ^63^Cu và ^65^Cu với nguyên tử khối trung bình là 63,54. (Khi tính khối lượng mol của CuCl~2~, dùng nguyên tử khối trung bình của Cl tính từ phần trăm các đồng vị ở trên.)")
    stmts1 = [
        ("a", "Hạt nhân nguyên tử đồng vị ^37^Cl có 17 proton và 20 neutron.", True),
        ("b", "Trong hợp chất CuCl~2~, phần trăm khối lượng của đồng vị ^63^Cu xấp xỉ bằng 34,19%.", True),
        ("c", "Nguyên tử khối trung bình của chlorine là 35,4846.", True),
        ("d", "Nếu lấy 13,45092 gam CuCl~2~ thì tổng số hạt neutron trong các nguyên tử ^37^Cl là 2,91826 × 10^23^ hạt.", False)
    ]
    for label, text, ans_bool in stmts1:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 1:\n- Ý a) ĐÚNG: Đồng vị ^37_17Cl có Z = 17 (17 proton) và số neutron N = 37 - 17 = 20.\n- Ý c) ĐÚNG: Nguyên tử khối trung bình của Cl là:\n  A_tb(Cl) = (35 × 75,77 + 37 × 24,23) / 100 = 3548,46 / 100 = 35,4846 amu.\n- Ý b) ĐÚNG: Gọi % số nguyên tử của ^63Cu là x: A_tb(Cu) = [63x + 65(100 - x)] / 100 = 63,54 => x = 73%.\n  Phân tử khối trung bình của CuCl~2~ = 63,54 + 2 × 35,4846 = 134,5092 amu.\n  Trong 1 mol CuCl~2~ có 1 mol Cu, trong đó số mol đồng vị ^63Cu là 1 × 73% = 0,73 mol.\n  Khối lượng của ^63Cu = 0,73 × 63 = 45,99 gam.\n  Phần trăm khối lượng của ^63Cu trong CuCl~2~ = (45,99 / 134,5092) × 100% ≈ 34,19%.\n- Ý d) SAI: Số mol CuCl~2~ = 13,45092 / 134,5092 = 0,1 mol.\n  Vì trong 1 phân tử CuCl~2~ có 2 nguyên tử Cl nên số mol nguyên tử Cl = 0,1 × 2 = 0,2 mol.\n  Số mol đồng vị ^37Cl = 0,2 × 24,23% = 0,04846 mol.\n  Mỗi nguyên tử ^37Cl có 20 neutron => Tổng số mol neutron của ^37Cl = 0,04846 × 20 = 0,9692 mol.\n  Tổng số hạt neutron = 0,9692 × 6,022 × 10^23 ≈ 5,8365 × 10^23 hạt (mệnh đề đưa ra 2,91826 × 10^23 là do quên nhân 2 nguyên tử Cl trong CuCl~2~).\n➔ Kết luận: a - Đ, b - Đ, c - Đ, d - S.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Cho các nguyên tố thuộc chu kì 3 trong bảng tuần hoàn: Na, Mg, Al, P, S, Cl.")
    stmts2 = [
        ("a", "Trong số các nguyên tố trên, Na là nguyên tố có bán kính nguyên tử lớn nhất và Cl có độ âm điện lớn nhất.", True),
        ("b", "Các oxide cao nhất gồm Na~2~O, MgO, Al~2~O~3~ đều là các basic oxide, còn P~2~O~5~, SO~3~, Cl~2~O~7~ đều là các acidic oxide.", False),
        ("c", "Theo chiều tăng dần của điện tích hạt nhân, hóa trị cao nhất của các nguyên tố với oxygen tăng dần từ I đến VII, đồng thời hóa trị với hydrogen của P, S, Cl giảm dần từ III xuống I.", True),
        ("d", "Cho một lượng dư dung dịch NaOH vào hỗn hợp gồm Al(OH)~3~ và Mg(OH)~2~, sau phản ứng chỉ có Al(OH)~3~ bị hòa tan.", True)
    ]
    for label, text, ans_bool in stmts2:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 2:\n- Ý a) ĐÚNG: Trong chu kì 3, từ Na đến Cl, điện tích hạt nhân tăng dần => Bán kính nguyên tử giảm dần (Na lớn nhất) và độ âm điện tăng dần (Cl lớn nhất).\n- Ý b) SAI: Al~2~O~3~ là oxide lưỡng tính (amphoteric oxide), không phải basic oxide.\n- Ý c) ĐÚNG: Hóa trị cao nhất với oxygen của các nguyên tố chu kì 3 tăng dần từ I đến VII: Na~2~O (I), MgO (II), Al~2~O~3~ (III), P~2~O~5~ (V), SO~3~ (VI), Cl~2~O~7~ (VII). Với hydrogen: PH~3~ (III), H~2~S (II), HCl (I) giảm dần từ III xuống I.\n- Ý d) ĐÚNG: Al(OH)~3~ là hydroxide lưỡng tính, tan được trong dung dịch NaOH dư: Al(OH)~3~ + NaOH -> NaAlO~2~ + 2H~2~O. Còn Mg(OH)~2~ là base không tan, không phản ứng với NaOH, nên chỉ có Al(OH)~3~ bị hòa tan.\n➔ Kết luận: a - Đ, b - S, c - Đ, d - Đ.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Cho bốn nguyên tố X, Y, Z, T có số hiệu nguyên tử lần lượt là 7, 14, 24, 26.")
    stmts3 = [
        ("a", "Khi biểu diễn cấu hình electron theo ô orbital ở trạng thái cơ bản, nguyên tử T (Z = 26) có tổng cộng 15 orbital chứa electron, trong đó có 4 AO chứa 1 electron độc thân.", True),
        ("b", "Nguyên tử Y ở trạng thái cơ bản có 4 electron ở lớp ngoài cùng và thuộc chu kì 3, nhóm IVA.", True),
        ("c", "Ở trạng thái cơ bản, nguyên tử Z (Z = 24) có 6 electron độc thân và có 2 electron ở lớp ngoài cùng.", False),
        ("d", "Ở trạng thái cơ bản, nguyên tử X có 3 electron độc thân ở phân lớp 2p.", True)
    ]
    for label, text, ans_bool in stmts3:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 3:\n- Cấu hình electron của các nguyên tử:\n  + X (Z = 7): 1s^2 2s^2 2p^3\n  + Y (Z = 14): 1s^2 2s^2 2p^6 3s^2 3p^2\n  + Z (Z = 24): [Ar] 3d^5 4s^1\n  + T (Z = 26): [Ar] 3d^6 4s^2\n- Ý a) ĐÚNG: Với T (Z = 26): các AO chứa electron gồm 1s (1 AO), 2s (1 AO), 2p (3 AO), 3s (1 AO), 3p (3 AO), 4s (1 AO), 3d (5 AO) => Tổng cộng 1 + 1 + 3 + 1 + 3 + 1 + 5 = 15 orbital. Phân lớp 3d^6 có 5 AO chứa 6 electron: gồm 1 AO có 2e ghép đôi và 4 AO có 1e độc thân.\n- Ý b) ĐÚNG: Y (Z = 14) có 3 lớp electron => chu kì 3; lớp ngoài cùng là lớp 3 có 4 electron (3s^2 3p^2), là nguyên tố p => nhóm IVA.\n- Ý c) SAI: Z (Z = 24) có cấu hình [Ar] 3d^5 4s^1: có 5e độc thân ở 3d và 1e độc thân ở 4s (tổng 6e độc thân). Tuy nhiên lớp ngoài cùng (lớp 4) chỉ có 1 electron (4s^1), mệnh đề nói \"có 2 electron ở lớp ngoài cùng\" là sai.\n- Ý d) ĐÚNG: X (Z = 7) có phân lớp 2p^3, theo quy tắc Hund 3 electron phân bố đều vào 3 AO 2p nên có 3 electron độc thân.\n➔ Kết luận: a - Đ, b - Đ, c - S, d - Đ.")

    # =========================================================================
    # PHẦN III: TRẢ LỜI NGẮN
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN III (1 điểm). CÂU TRẮC NGHIỆM YÊU CẦU TRẢ LỜI NGẮN",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Hòa tan hoàn toàn 6,85 gam một kim loại nhóm IIA (M) vào 200 gam dung dịch HCl 3,65%, cô cạn dung dịch thu được 10,4 gam chất rắn khan. Xác định nguyên tử khối của M.")
    if is_teacher:
        add_solution_no_table(doc, "Kim loại M thuộc nhóm IIA có hóa trị II trong hợp chất.\nPhương trình hóa học: M + 2HCl -> MCl~2~ + H~2~↑\nSố mol HCl ban đầu = (200 × 3,65%) / 36,5 = 0,2 mol.\nSố mol kim loại M là n_M = 6,85 / M_M.\nKhi cô cạn dung dịch, axit HCl dư và H~2~O bay hơi hết, chất rắn khan thu được chỉ là muối MCl~2~.\nKhối lượng muối MCl~2~: m_MCl2 = n_M × (M_M + 71) = 10,4 gam.\n<=> (6,85 / M_M) × (M_M + 71) = 10,4\n<=> 6,85 + 71 × (6,85 / M_M) = 10,4\n<=> 71 × (6,85 / M_M) = 3,55  =>  6,85 / M_M = 0,05 mol.\n=> M_M = 6,85 / 0,05 = 137 (Barium - Ba).\nKiểm tra lại: n_HCl phản ứng = 2 × 0,05 = 0,1 mol < 0,2 mol (HCl dư, thỏa mãn).\n➔ Đáp số: 137")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Ở 20 °C, khối lượng riêng của Fe (M = 55,85 g/mol) là 7,87 g/cm^3^. Giả thiết nguyên tử Fe là hình cầu chiếm 68% thể tích tinh thể. Tính bán kính nguyên tử Fe theo đơn vị pm (làm tròn kết quả đến hàng đơn vị).")
    if is_teacher:
        add_solution_no_table(doc, "Xét 1 mol nguyên tử Fe:\n- Thể tích 1 mol tinh thể Fe: V_tinh_thể = M / D = 55,85 / 7,87 ≈ 7,09657 cm^3^.\n- Thể tích thực của 1 mol nguyên tử Fe (chiếm 68% thể tích tinh thể):\n  V_thực = 7,09657 × 68% ≈ 4,82567 cm^3^.\n- Thể tích của 1 nguyên tử Fe:\n  V_1_nguyên_tử = V_thực / N_A = 4,82567 / (6,022 × 10^23^) ≈ 8,0134 × 10^-24^ cm^3^.\n- Giả thiết nguyên tử là hình cầu có thể tích V = (4/3)πr^3^:\n  r^3^ = (3 × V_1_nguyên_tử) / (4π) = (3 × 8,0134 × 10^-24^) / (4 × 3,14159) ≈ 1,9129 × 10^-24^ cm^3^.\n  => r ≈ 1,2413 × 10^-8^ cm = 1,2413 × 10^-10^ m = 124,13 pm ≈ 124 pm.\n➔ Đáp số: 124")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Trong tự nhiên, H có 2 đồng vị bền (^1^H, ^2^H), S có 1 đồng vị bền (^32^S), O có 3 đồng vị bền (^16^O, ^17^O, ^18^O). Có tối đa bao nhiêu loại phân tử H~2~SO~4~ khác nhau về phân tử khối được tạo thành từ các đồng vị trên?")
    if is_teacher:
        add_solution_no_table(doc, "Phân tử H~2~SO~4~ gồm: 2 nguyên tử H, 1 nguyên tử S, 4 nguyên tử O.\nPhân tử khối = m(H~2~) + m(S) + m(O~4~).\n- Phần S chỉ có đồng vị ^32S nên m(S) luôn bằng 32.\n- Phần 2 nguyên tử H tạo từ các đồng vị (^1H, ^2H):\n  Tổng số khối m(H~2~) có thể nhận các giá trị:\n  1 + 1 = 2;  1 + 2 = 3;  2 + 2 = 4  => m(H~2~) ∈ {2, 3, 4}.\n- Phần 4 nguyên tử O tạo từ các đồng vị (^16O, ^17O, ^18O):\n  + Giá trị nhỏ nhất: 16 × 4 = 64.\n  + Giá trị lớn nhất: 18 × 4 = 72.\n  Do số khối các đồng vị O là 3 số nguyên liên tiếp nên tổng số khối của 4 nguyên tử O nhận mọi giá trị nguyên liên tiếp từ 64 đến 72: {64, 65, 66, 67, 68, 69, 70, 71, 72} (tổng cộng 9 giá trị).\n- Phân tử khối H~2~SO~4~ = m(H~2~) + 32 + m(O~4~):\n  + Giá trị nhỏ nhất: 2 + 32 + 64 = 98.\n  + Giá trị lớn nhất: 4 + 32 + 72 = 108.\n  Vì m(H~2~) quét qua các số nguyên liên tiếp {2, 3, 4} và m(O~4~) quét qua các số nguyên liên tiếp từ 64 đến 72 nên tổng nhận tất cả các giá trị nguyên liên tiếp từ 98 đến 108.\n  Số loại phân tử khối khác nhau = 108 - 98 + 1 = 11 loại.\n➔ Đáp số: 11")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Hai nguyên tố X, Y thuộc nhóm A, kế tiếp nhau trong một chu kì, tổng số proton trong hai hạt nhân là 31. Ở trạng thái cơ bản, nguyên tử có độ âm điện lớn hơn có bao nhiêu electron độc thân?")
    if is_teacher:
        add_solution_no_table(doc, "- Hai nguyên tố X, Y thuộc nhóm A, kế tiếp nhau trong một chu kì nên hiệu số proton là:\n  |Z_Y - Z_X| = 1.\n  Theo đề bài: Z_X + Z_Y = 31.\n  Giả sử Z_X < Z_Y => Z_X = (31 - 1) / 2 = 15 (Phosphorus - P) và Z_Y = 16 (Sulfur - S).\n- Trong cùng một chu kì, đi từ trái sang phải theo chiều tăng Z, độ âm điện tăng dần => Độ âm điện của S (Z = 16) lớn hơn P (Z = 15).\n- Cấu hình electron của S (Z = 16): 1s^2 2s^2 2p^6 3s^2 3p^4.\n- Phân bố electron vào các orbital ở phân lớp 3p^4:\n  Có 3 AO 3p, trong đó 1 AO chứa 2 electron ghép đôi và 2 AO mỗi AO chứa 1 electron độc thân.\n  Vậy nguyên tử có độ âm điện lớn hơn (S) có 2 electron độc thân.\n➔ Đáp số: 2")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # =========================================================================
    # B. PHẦN TỰ LUẬN: (3,0 điểm)
    # =========================================================================
    p_sec_b = doc.add_paragraph()
    p_sec_b.paragraph_format.space_before = Pt(8)
    p_sec_b.paragraph_format.space_after = Pt(2)
    render_rich_text(p_sec_b, "*B. PHẦN TỰ LUẬN (3,0 điểm)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # Câu 1 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1 (1,0 điểm): Nguyên tử của nguyên tố X có tổng số các loại hạt cơ bản (proton, neutron, electron) là 36. Trong đó, số hạt mang điện gấp đôi số hạt không mang điện.\na) Tính số hạt proton, neutron và electron của nguyên tử X.\nb) Viết cấu hình electron của nguyên tử X và cho biết X là kim loại, phi kim hay khí hiếm? Vì sao?")
    if is_teacher:
        sol_tl1 = (
            "Hướng dẫn giải chi tiết Câu 1 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "- Gọi số hạt proton, neutron, electron của nguyên tử X lần lượt là p, n, e.\n"
            "- Do nguyên tử trung hòa về điện nên p = e.\n"
            "- Theo đề bài ta có hệ phương trình:\n"
            "    + Tổng số hạt: 2p + n = 36  (1)\n"
            "    + Số hạt mang điện gấp đôi số hạt không mang điện: 2p = 2n <=> p = n  (2) (0,25đ)\n"
            "- Thay (2) vào (1) ta được: 2n + n = 36 <=> 3n = 36 => n = 12.\n"
            "- Suy ra: p = e = 12.\n"
            "Vậy số hạt proton là 12, electron là 12 và neutron là 12. (0,25đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Với Z = 12, cấu hình electron của nguyên tử X là: 1s^2 2s^2 2p^6 3s^2 (hoặc [Ne] 3s^2). (0,25đ)\n"
            "- Nguyên tử X có 2 electron ở lớp ngoài cùng (lớp 3) nên X là nguyên tố kim loại (cụ thể là Magnesium - Mg). (0,25đ)"
        )
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 2 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2 (1,0 điểm): Cho hai nguyên tố X có số hiệu nguyên tử Z = 11 (Sodium) và Y có số hiệu nguyên tử Z = 12 (Magnesium).\na) Viết cấu hình electron của X, Y và xác định vị trí của chúng (chu kì, nhóm) trong bảng tuần hoàn.\nb) Viết công thức hydroxide tương ứng của X, Y. Theo quy luật tuần hoàn, hãy so sánh tính base của hai hydroxide này.")
    if is_teacher:
        sol_tl2 = (
            "Hướng dẫn giải chi tiết Câu 2 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "- Nguyên tố X (Z = 11):\n"
            "    + Cấu hình electron: 1s^2 2s^2 2p^6 3s^1 (hoặc [Ne] 3s^1).\n"
            "    + Vị trí: Ô số 11, chu kì 3 (có 3 lớp e), nhóm IA (có 1 electron hóa trị ở phân lớp s). (0,25đ)\n"
            "- Nguyên tố Y (Z = 12):\n"
            "    + Cấu hình electron: 1s^2 2s^2 2p^6 3s^2 (hoặc [Ne] 3s^2).\n"
            "    + Vị trí: Ô số 12, chu kì 3 (có 3 lớp e), nhóm IIA (có 2 electron hóa trị ở phân lớp s). (0,25đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Công thức hydroxide tương ứng:\n"
            "    + Của X: NaOH.\n"
            "    + Của Y: Mg(OH)~2~. (0,25đ)\n"
            "- So sánh tính base: Hai nguyên tố X (Na) và Y (Mg) cùng thuộc chu kì 3. Trong cùng một chu kì, theo chiều tăng dần của điện tích hạt nhân (từ trái sang phải), tính kim loại giảm dần kéo theo tính base của các hydroxide giảm dần.\n"
            "Do Z_Na (11) < Z_Mg (12) nên tính base của NaOH mạnh hơn Mg(OH)~2~ (NaOH > Mg(OH)~2~). (0,25đ)"
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 3 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3 (1,0 điểm): Nguyên tố R thuộc nhóm VIA trong bảng tuần hoàn. Trong phân tử oxide cao nhất của R, nguyên tố oxygen chiếm 60% về khối lượng.\na) Xác định tên và kí hiệu hóa học của nguyên tố R.\nb) Hòa tan hoàn toàn 8,0 gam oxide cao nhất của R vào nước dư, thu được dung dịch chứa m gam acid. Tính giá trị của m.")
    if is_teacher:
        sol_tl3 = (
            "Hướng dẫn giải chi tiết Câu 3 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "- Vì R thuộc nhóm VIA nên hóa trị cao nhất của R với oxygen là VI.\n"
            "- Công thức oxide cao nhất của R có dạng RO~3~. (0,25đ)\n"
            "- Trong phân tử RO~3~, oxygen chiếm 60% về khối lượng:\n"
            "    %m_O = (3 × 16) / (M_R + 3 × 16) × 100% = 60%\n"
            "    <=> 48 / (M_R + 48) = 0,60\n"
            "    <=> M_R + 48 = 48 / 0,60 = 80\n"
            "    <=> M_R = 80 - 48 = 32 (g/mol).\n"
            "- Vậy nguyên tố R là Sulfur (lưu huỳnh), kí hiệu hóa học là S. (0,25đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Oxide cao nhất của R là SO~3~ (M = 80 g/mol).\n"
            "- Số mol SO~3~ là: n_SO3 = 8,0 / 80 = 0,1 mol. (0,15đ)\n"
            "- Phương trình phản ứng khi cho SO~3~ tác dụng với nước:\n"
            "    SO~3~ + H~2~O -> H~2~SO~4~ (0,15đ)\n"
            "- Theo phương trình: n_H2SO4 = n_SO3 = 0,1 mol.\n"
            "- Khối lượng acid H~2~SO~4~ (M = 98 g/mol) thu được là:\n"
            "    m = 0,1 × 98 = 9,8 gam.\n"
            "Vậy giá trị của m là 9,8. (0,2đ)"
        )
        add_solution_no_table(doc, sol_tl3)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Cuối đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "*---------- HẾT ----------*", base_size=10.5, base_bold=True)

    if output_path:
        doc.save(output_path)
        print(f"[OK] Đã lưu thành công: {output_path}")

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    file_student = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_NguyenDinhChieu.docx")
    file_student_alias = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_NguyenDinhChieu_HocSinh.docx")
    file_teacher = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_NguyenDinhChieu_GiaoVien.docx")

    print("Đang tạo Bản Học sinh THPT Nguyễn Đình Chiểu...")
    build_exam_ndc(is_teacher=False, output_path=file_student)
    build_exam_ndc(is_teacher=False, output_path=file_student_alias)

    print("Đang tạo Bản Giáo viên THPT Nguyễn Đình Chiểu...")
    build_exam_ndc(is_teacher=True, output_path=file_teacher)

    print("\n[HOÀN TẤT] File Word THPT Nguyễn Đình Chiểu đã được tạo chuẩn 100% FORM ĐỀ KIỂM TRA!")

if __name__ == "__main__":
    main()
