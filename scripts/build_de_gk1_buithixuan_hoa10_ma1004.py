# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì 1 Hóa Học 10 - THPT Bùi Thị Xuân (Huế) - MÃ ĐỀ 1004
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản Word:
1. Bản Học sinh: De_Kiem_Tra_Hoa_10_Ma1004.docx
2. Bản Giáo viên: De_Kiem_Tra_Hoa_10_Ma1004_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết từng câu)
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

def add_quick_answer_table_1004(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu MÃ ĐỀ 1004.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 1004)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (12 CÂU)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - mỗi câu đúng 0,25 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    col_widths_1 = [2.4] + [1.3] * 12
    p1_answers = ["C", "D", "D", "B", "C", "B", "C", "A", "B", "B", "A", "D"]

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
    p_p2.paragraph_format.space_before = Pt(4)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm - 1 ý: 0,1đ; 2 ý: 0,25đ; 3 ý: 0,5đ; 4 ý: 1,0đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_2 = [2.4, 2.9, 3.175, 3.175, 3.175, 3.175]
    
    p2_data = [
        ["Câu", "Nội dung", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Mg và Ca", "SAI", "ĐÚNG", "ĐÚNG", "SAI"],
        ["Câu 2", "Nguyên tử X và Y", "ĐÚNG", "ĐÚNG", "SAI", "SAI"]
    ]

    for row_idx, row in enumerate(tbl2.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        row_data = p2_data[row_idx]
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
                if col_idx == 0:
                    render_rich_text(p, f"*{val}*", base_size=10, base_bold=True)
                elif col_idx == 1:
                    render_rich_text(p, f"_{val}_", base_size=9.5, base_italic=True, base_color=COLOR_PRIMARY)
                else:
                    color = COLOR_CORRECT if val in ["ĐÚNG", "SAI"] else None
                    render_rich_text(p, f"*{val}*", base_size=10, base_bold=True, base_color=color)

    # 3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (4 CÂU)
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(4)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*3. Phần III: Trắc nghiệm trả lời ngắn (2,0 điểm - mỗi câu đúng 0,5 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_3 = [3.6, 3.6, 3.6, 3.6, 3.6]
    p3_answers = ["24", "3", "21", "27"]

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
                txt = "Mục" if col_idx == 0 else f"Câu {col_idx}"
                render_rich_text(p, f"*{txt}*", base_size=10, base_bold=True)
            else:
                txt = "Đáp án" if col_idx == 0 else p3_answers[col_idx - 1]
                if col_idx == 0:
                    render_rich_text(p, f"*{txt}*", base_size=10, base_bold=True)
                else:
                    render_rich_text(p, f"*{txt}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)

    # 4. PHẦN IV: TỰ LUẬN (3 CÂU)
    p_p4 = doc.add_paragraph()
    p_p4.paragraph_format.space_before = Pt(4)
    p_p4.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p4, "*4. Phần IV: Tự luận (3,0 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl4 = doc.add_table(rows=4, cols=3)
    tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_4 = [2.4, 2.9, 12.7]
    p4_data = [
        ["Câu", "Điểm", "Đáp án vắn tắt"],
        ["Câu 1", "1,0 điểm", "a) p = 3, e = 3, n = 4; Kí hiệu: ^7_3Y (hoặc ^7_3Li). b) Y và T là đồng vị vì cùng Z = 3."],
        ["Câu 2", "1,0 điểm", "a) Z = 16, [Ne]3s^2 3p^4 (2, 8, 6). b) 2e độc thân; X + 2e -> X^2- giống khí hiếm Argon (Ar)."],
        ["Câu 3", "1,0 điểm", "a) Z_X = 12 (Mg), Z_Y = 13 (Al). b) Vị trí: X ô 12, CK 3, IIA; Y ô 13, CK 3, IIIA. c) Cả 2 là KL; X nguyên tố s, Y nguyên tố p."]
    ]

    for row_idx, row in enumerate(tbl4.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        row_data = p4_data[row_idx]
        for col_idx, cell in enumerate(row.cells):
            cell.width = Cm(col_widths_4[col_idx])
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


def build_exam_buithixuan_1004(is_teacher=False, output_path=None):
    doc = docx.Document()
    setup_header_footer(doc)

    # 1. Khối tiêu đề đầu đề thi
    build_title_block(
        doc,
        school_line="SỞ GD&ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT BÙI THỊ XUÂN",
        exam_title="KIỂM TRA ĐÁNH GIÁ GIỮA KÌ\nHỌC KÌ I, NĂM HỌC 2024 - 2025",
        subject_line="MÔN: HOÁ HỌC_LỚP 10",
        num_pages="02",
        duration="45 phút",
        exam_code="1004",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_1004(doc)

    # =========================================================================
    # PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (3,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều lựa chọn (3,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi học sinh chỉ chọn một phương án."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Phát biểu nào sau đây không đúng?")
    opts1 = [
        "A. Nguyên tử được cấu thành từ các hạt cơ bản là proton, neutron và electron.",
        "B. Hạt nhân nguyên tử cấu thành từ các hạt proton và neutron.",
        "*C. Nguyên tử có cấu trúc đặc khít, gồm vỏ nguyên tử và hạt nhân nguyên tử.*" if is_teacher else "C. Nguyên tử có cấu trúc đặc khít, gồm vỏ nguyên tử và hạt nhân nguyên tử.",
        "D. Vỏ nguyên tử cấu thành từ các hạt electron."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nguyên tử có cấu trúc rỗng (đường kính nguyên tử lớn hơn đường kính hạt nhân khoảng 10.000 lần), ở giữa là hạt nhân mang điện tích dương và vỏ electron mang điện tích âm, không phải cấu trúc đặc khít.\n=> Đáp án đúng: C.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Hạt có khối lượng xấp xỉ 1 amu trong hạt nhân nguyên tử là")
    opts2 = [
        "A. proton và electron.",
        "B. neutron và electron.",
        "C. electron.",
        "*D. proton và neutron.*" if is_teacher else "D. proton và neutron."
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong hạt nhân nguyên tử chỉ chứa các hạt proton và neutron. Khối lượng của hạt proton m_p ≈ 1,0073 amu và hạt neutron m_n ≈ 1,0087 amu đều xấp xỉ 1 amu (trong khi electron ở vỏ có khối lượng rất nhỏ m_e ≈ 0,00055 amu).\n=> Đáp án đúng: D.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Đường kính nguyên tử khoảng 10^-1 nm, đường kính hạt nhân khoảng 10^-5 nm. Đường kính nguyên tử gấp khoảng bao nhiêu lần đường kính hạt nhân?")
    opts3 = [
        "A. 10 lần.",
        "B. 100 lần.",
        "C. 1 000 lần.",
        "*D. 10 000 lần.*" if is_teacher else "D. 10 000 lần."
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Tỉ số kích thước giữa đường kính nguyên tử và đường kính hạt nhân:\n    d_nguyên tử / d_hạt nhân = 10^-1 / 10^-5 = 10^4 = 10 000 (lần).\n=> Đáp án đúng: D.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Hoạt động nào sau đây thuộc phương pháp thực nghiệm trong học tập và nghiên cứu hóa học?")
    opts4 = [
        "A. Đọc tài liệu và ghi nhớ nội dung lí thuyết.",
        "*B. Tiến hành thí nghiệm, quan sát hiện tượng và rút ra kết luận.*" if is_teacher else "B. Tiến hành thí nghiệm, quan sát hiện tượng và rút ra kết luận.",
        "C. Xem video bài giảng của giáo viên.",
        "D. Trao đổi nhóm để tóm tắt nội dung bài học."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phương pháp thực nghiệm trong nghiên cứu và học tập hóa học là phương pháp thu thập dữ liệu khách quan thông qua việc trực tiếp thực hiện thí nghiệm, quan sát hiện tượng xảy ra và rút ra kết luận khoa học.\n=> Đáp án đúng: B.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 5. Để kiểm tra một giả thuyết (dự đoán) trong nghiên cứu hóa học, người ta thường")
    opts5 = [
        "A. chỉ dựa vào ý kiến của số đông.",
        "B. chỉ đọc lại kết luận có sẵn trong sách.",
        "*C. tiến hành thí nghiệm, quan sát hiện tượng, thu thập số liệu rồi đối chiếu với dự đoán.*" if is_teacher else "C. tiến hành thí nghiệm, quan sát hiện tượng, thu thập số liệu rồi đối chiếu với dự đoán.",
        "D. chọn những kết quả phù hợp với dự đoán và bỏ qua những kết quả khác."
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Để kiểm tra một giả thuyết (dự đoán khoa học), quy trình chuẩn là thiết kế và tiến hành thí nghiệm kiểm chứng, quan sát hiện tượng, ghi chép số liệu thực nghiệm rồi so sánh, đối chiếu với giả thuyết ban đầu.\n=> Đáp án đúng: C.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 6. Cho ba nguyên tử: ^12_6X, ^14_6Y và ^14_7Z. Nhận định nào sau đây đúng?")
    opts6 = [
        "A. Y và Z là hai đồng vị của cùng một nguyên tố vì có cùng số khối.",
        "*B. X và Y là hai đồng vị của cùng một nguyên tố hóa học.*" if is_teacher else "B. X và Y là hai đồng vị của cùng một nguyên tố hóa học.",
        "C. X và Z có cùng số neutron.",
        "D. X và Z có cùng số khối."
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích cấu tạo 3 nguyên tử:\n- ^12_6X: Z = 6, số khối A = 12 => Số neutron N = 12 - 6 = 6.\n- ^14_6Y: Z = 6, số khối A = 14 => Số neutron N = 14 - 6 = 8.\n- ^14_7Z: Z = 7, số khối A = 14 => Số neutron N = 14 - 7 = 7.\nX và Y có cùng số proton (Z = 6) nhưng khác số neutron (6 và 8) nên X và Y là hai đồng vị của cùng một nguyên tố hóa học (carbon).\n=> Đáp án đúng: B.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 7. Mô hình nguyên tử hiện đại mô tả chuyển động của electron là")
    opts7 = [
        "A. theo quỹ đạo hình tròn hay bầu dục.",
        "B. theo quỹ đạo đường thẳng.",
        "*C. không theo quỹ đạo xác định.*" if is_teacher else "C. không theo quỹ đạo xác định.",
        "D. theo quỹ đạo hình cầu."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo mô hình hiện đại (cơ học lượng tử), trong nguyên tử, các electron chuyển động rất nhanh trong khu vực xung quanh hạt nhân và không theo những quỹ đạo xác định. Vùng không gian xung quanh hạt nhân có xác suất tìm thấy electron lớn nhất gọi là orbital nguyên tử (AO).\n=> Đáp án đúng: C.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 8. Nhà khoa học nào sau đây đã công bố bảng tuần hoàn các nguyên tố hóa học đầu tiên vào năm 1869?")
    opts8 = [
        "*A. D. I. Mendeleev.*" if is_teacher else "A. D. I. Mendeleev.",
        "B. E. Rutherford.",
        "C. J. J. Thomson.",
        "D. N. Bohr."
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Năm 1869, nhà hóa học người Nga D. I. Mendeleev đã công bố bảng tuần hoàn các nguyên tố hóa học đầu tiên, sắp xếp các nguyên tố theo chiều tăng dần của khối lượng nguyên tử.\n=> Đáp án đúng: A.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 9. Nguyên tử của nguyên tố X có tổng số electron trong các phân lớp p là 10. Vị trí của X trong bảng tuần hoàn là")
    opts9 = [
        "A. ô số 10, chu kì 2, nhóm VIIIA.",
        "*B. ô số 16, chu kì 3, nhóm VIA.*" if is_teacher else "B. ô số 16, chu kì 3, nhóm VIA.",
        "C. ô số 14, chu kì 3, nhóm IVA.",
        "D. ô số 16, chu kì 3, nhóm IVA."
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "X có tổng số electron trên các phân lớp p là 10: gồm 6e ở phân lớp 2p và 4e ở phân lớp 3p.\n- Cấu hình electron của X: 1s^2 2s^2 2p^6 3s^2 3p^4.\n- Tổng số electron = 2 + 2 + 6 + 2 + 4 = 16 => Z = 16.\n- Vị trí của X trong bảng tuần hoàn:\n  + Ô số 16 (do Z = 16).\n  + Chu kì 3 (do có 3 lớp electron).\n  + Nhóm VIA (nguyên tố p, có 2 + 4 = 6 electron ở lớp ngoài cùng).\n=> Đáp án đúng: B.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 10. Cho các nguyên tố X (Z = 13), Y (Z = 17) và T (Z = 20). Phát biểu nào sau đây đúng?")
    opts10 = [
        "A. X, Y đều là phi kim.",
        "*B. X, T là kim loại; Y là phi kim.*" if is_teacher else "B. X, T là kim loại; Y là phi kim.",
        "C. X là phi kim; Y, T là kim loại.",
        "D. X, Y, T đều là kim loại."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Xét cấu hình electron và tính chất:\n- X (Z = 13): 1s^2 2s^2 2p^6 3s^2 3p^1 => Có 3e ở lớp ngoài cùng => Kim loại (Aluminium).\n- Y (Z = 17): 1s^2 2s^2 2p^6 3s^2 3p^5 => Có 7e ở lớp ngoài cùng => Phi kim (Chlorine).\n- T (Z = 20): 1s^2 2s^2 2p^6 3s^2 3p^6 4s^2 => Có 2e ở lớp ngoài cùng => Kim loại (Calcium).\nVậy X, T là kim loại; Y là phi kim.\n=> Đáp án đúng: B.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 11. Số hiệu nguyên tử của nguyên tố hóa học bằng")
    opts11 = [
        "*A. số thứ tự của ô nguyên tố.*" if is_teacher else "A. số thứ tự của ô nguyên tố.",
        "B. số thứ tự của chu kì.",
        "C. số thứ tự của nhóm.",
        "D. số electron lớp ngoài cùng của nguyên tử."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong bảng tuần hoàn, số thứ tự ô nguyên tố = Số hiệu nguyên tử Z = Số đơn vị điện tích hạt nhân = Số proton = Số electron.\n=> Đáp án đúng: A.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 12. Nhóm A bao gồm các nguyên tố")
    opts12 = [
        "A. s.",
        "B. p.",
        "C. d và f.",
        "*D. s và p.*" if is_teacher else "D. s và p."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nhóm A (nhóm chính) bao gồm các nguyên tố s (nhóm IA, IIA) và các nguyên tố p (từ nhóm IIIA đến nhóm VIIIA). Các nguyên tố d và f thuộc nhóm B.\n=> Đáp án đúng: D.")

    # =========================================================================
    # PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (2,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng – sai (2,0 điểm)",
        subtitle="Học sinh trả lời câu 1 và câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai."
    )

    # Câu 1 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Magnesium được dùng để chế tạo hợp kim nhẹ, calcium được dùng trong sản xuất vôi và xi măng. Cho nguyên tử magnesium (Mg, Z = 12) và calcium (Ca, Z = 20).")
    stmts1 = [
        ("a", "Mg là nguyên tố s, còn Ca là nguyên tố p.", False),
        ("b", "Từ cấu hình electron, xác định được Ca nằm ở ô số 20, chu kì 4, nhóm IIA.", True),
        ("c", "Bảng tuần hoàn hiện nay có 8 nhóm A và 8 nhóm B.", True),
        ("d", "Mg và Ca thuộc cùng một chu kì vì nguyên tử của chúng đều có 2 electron ở lớp ngoài cùng.", False)
    ]
    for label, text, ans_bool in stmts1:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 1:\n- Cấu hình electron của Mg (Z = 12): 1s^2 2s^2 2p^6 3s^2 => Nguyên tố s (e cuối cùng điền vào phân lớp s).\n- Cấu hình electron của Ca (Z = 20): 1s^2 2s^2 2p^6 3s^2 3p^6 4s^2 => Nguyên tố s.\n+ Ý a) SAI: Cả Mg và Ca đều là nguyên tố s.\n+ Ý b) ĐÚNG: Ca có Z = 20 => Ô 20; có 4 lớp electron => Chu kì 4; có 2 electron lớp ngoài cùng và là nguyên tố s => Nhóm IIA.\n+ Ý c) ĐÚNG: Bảng tuần hoàn hiện nay có 18 cột, gồm 8 nhóm A (IA đến VIIIA) và 8 nhóm B (IB đến VIIIB).\n+ Ý d) SAI: Mg thuộc chu kì 3, Ca thuộc chu kì 4. Mg và Ca thuộc cùng một NHÓM (nhóm IIA) do cùng có 2 electron lớp ngoài cùng, không phải cùng một chu kì.\n=> Kết luận: a - SAI, b - ĐÚNG, c - ĐÚNG, d - SAI.")

    # Câu 2 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Nguyên tử của nguyên tố X có cấu hình electron kết thúc bằng 3p^2. Nguyên tử của nguyên tố Y có cấu hình electron kết thúc bằng 3p^5.")
    stmts2 = [
        ("a", "Nguyên tử Y có 1 electron độc thân ở trạng thái cơ bản.", True),
        ("b", "Nguyên tử Y có 7 electron ở lớp ngoài cùng nên Y là phi kim.", True),
        ("c", "Phân lớp 3p có 3 orbital và chứa tối đa 3 electron.", False),
        ("d", "Cấu hình electron theo lớp của X là (2,8,2).", False)
    ]
    for label, text, ans_bool in stmts2:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 2:\n- X: Cấu hình kết thúc bằng 3p^2 => 1s^2 2s^2 2p^6 3s^2 3p^2 (Z = 14, Silicon). Cấu hình theo lớp là (2, 8, 4).\n- Y: Cấu hình kết thúc bằng 3p^5 => 1s^2 2s^2 2p^6 3s^2 3p^5 (Z = 17, Chlorine).\n+ Ý a) ĐÚNG: Phân lớp 3p^5 có 3 ô orbital: [↑↓][↑↓][↑], do đó Y có 1 electron độc thân.\n+ Ý b) ĐÚNG: Y có 7 electron ở lớp ngoài cùng (3s^2 3p^5) nên Y là phi kim điển hình.\n+ Ý c) SAI: Phân lớp 3p có 3 orbital và chứa tối đa 6 electron (mỗi orbital chứa tối đa 2 electron ghép đôi).\n+ Ý d) SAI: Cấu hình theo lớp của X là (2, 8, 4), không phải (2, 8, 2).\n=> Kết luận: a - ĐÚNG, b - ĐÚNG, c - SAI, d - SAI.")

    # =========================================================================
    # PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN III. Câu trắc nghiệm trả lời ngắn (2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4."
    )

    # Câu 1 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Cho các nguyên tố có số hiệu nguyên tử lần lượt là 8, 12, 14, 16 và 19. Trong số đó có hai nguyên tố thuộc cùng một nhóm. Tổng số hiệu nguyên tử của hai nguyên tố này là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Xác định nhóm của từng nguyên tố:\n- Z = 8: 1s^2 2s^2 2p^4 => Nhóm VIA (chu kì 2).\n- Z = 12: 1s^2 2s^2 2p^6 3s^2 => Nhóm IIA (chu kì 3).\n- Z = 14: 1s^2 2s^2 2p^6 3s^2 3p^2 => Nhóm IVA (chu kì 3).\n- Z = 16: 1s^2 2s^2 2p^6 3s^2 3p^4 => Nhóm VIA (chu kì 3).\n- Z = 19: [Ar]4s^1 => Nhóm IA (chu kì 4).\nHai nguyên tố cùng nhóm là Z = 8 (Oxygen) và Z = 16 (Sulfur) đều thuộc nhóm VIA.\nTổng số hiệu nguyên tử của hai nguyên tố này là: 8 + 16 = 24.\n=> Đáp số: 24")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 2 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Nguyên tử Y có lớp electron thứ nhất đã bão hòa và lớp electron thứ hai chứa 5 electron. Số electron độc thân của Y là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "- Lớp thứ nhất (n = 1) đã bão hòa: 1s^2 (2 electron).\n- Lớp thứ hai (n = 2) chứa 5 electron: 2s^2 2p^3.\n=> Cấu hình electron của Y là 1s^2 2s^2 2p^3 (Z = 7, Nitrogen).\n- Ở phân lớp 2p^3, theo quy tắc Hund, 3 electron chiếm 3 orbital riêng biệt với spin song song: [↑][↑][↑].\nDo đó, nguyên tử Y có 3 electron độc thân.\n=> Đáp số: 3")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 3 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Nguyên tử R có điện tích lớp vỏ là -33,642.10^-19 C. Biết điện tích của một electron là -1,602.10^-19 C. Nguyên tử R có bao nhiêu electron?")
    if is_teacher:
        add_solution_no_table(doc, "Điện tích của lớp vỏ nguyên tử bằng tổng điện tích của các electron trong vỏ:\n    q_vỏ = N_e . q_e\nSố electron trong nguyên tử R là:\n    N_e = (-33,642 . 10^-19 C) / (-1,602 . 10^-19 C) = 21 electron.\n=> Đáp số: 21")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 4 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Trong tự nhiên, đồng có hai đồng vị bền là ^63Cu và ^65Cu, nguyên tử khối trung bình của đồng là 63,54. Phần trăm số nguyên tử của đồng vị ^65Cu là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Gọi phần trăm số nguyên tử của đồng vị ^65Cu là x (%), suy ra phần trăm của ^63Cu là (100 - x)%.\nÁp dụng công thức tính nguyên tử khối trung bình:\n    A_tb = [63.(100 - x) + 65.x] / 100 = 63,54\n    <=> 6300 + 2x = 6354\n    <=> 2x = 54  <=>  x = 27%.\nVậy phần trăm số nguyên tử của đồng vị ^65Cu là 27% (hoặc 27).\n=> Đáp số: 27")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # =========================================================================
    # PHẦN IV: TỰ LUẬN (3,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN IV. Tự luận (3,0 điểm)",
        subtitle="Học sinh trình bày lời giải chi tiết cho các câu hỏi sau."
    )

    # Câu 1 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. (1,0 điểm) Nguyên tử của nguyên tố Y có tổng số hạt là 10 và số khối bằng 7.\na) Xác định số proton, số neutron, số electron của Y và viết kí hiệu nguyên tử của Y.\nb) Cho thêm hai nguyên tố T và V có kí hiệu lần lượt là ^6_3T và ^7_4V. Trong ba nguyên tử Y, T, V, những nguyên tử nào là đồng vị của cùng một nguyên tố? Giải thích.")
    if is_teacher:
        sol_tl1 = (
            "Hướng dẫn giải chi tiết Câu 1 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "- Trong nguyên tử Y, gọi p, n, e lần lượt là số hạt proton, neutron và electron (p = e).\n"
            "- Theo đề bài:\n"
            "    + Tổng số hạt: p + n + e = 10  <=>  2p + n = 10 (1) (0,15đ)\n"
            "    + Số khối: A = p + n = 7 (2) (0,15đ)\n"
            "- Lấy (1) trừ (2): p = 10 - 7 = 3 => e = p = 3. (0,1đ)\n"
            "- Từ (2) suy ra: n = 7 - 3 = 4.\n"
            "- Vậy nguyên tử Y có: 3 proton, 3 electron và 4 neutron.\n"
            "- Kí hiệu nguyên tử của Y là: ^7_3Y (hoặc ^7_3Li). (0,1đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Ta có ba nguyên tử:\n"
            "    + Y có: Z = 3, số khối A = 7 (^7_3Y).\n"
            "    + T có: Z = 3, số khối A = 6 (^6_3T).\n"
            "    + V có: Z = 4, số khối A = 7 (^7_4V). (0,15đ)\n"
            "- Trong ba nguyên tử trên, hai nguyên tử Y và T là đồng vị của cùng một nguyên tố hóa học (Lithium). (0,2đ)\n"
            "- Giải thích: Vì Y và T có cùng số proton (Z = 3) nhưng có số neutron khác nhau (Y có 4 neutron, T có 3 neutron) dẫn đến số khối khác nhau (A_Y = 7, A_T = 6). (0,15đ)"
        )
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. (1,0 điểm) Nguyên tử của nguyên tố X có 3 lớp electron và lớp thứ ba chứa 6 electron.\na) Xác định số hiệu nguyên tử của X và viết cấu hình electron nguyên tử của X (theo lớp và phân lớp).\nb) Biểu diễn sự phân bố electron của X theo ô orbital và cho biết X có bao nhiêu electron độc thân. Nếu nguyên tử X nhận thêm electron để đạt cấu hình bền thì cấu hình electron của ion tạo thành giống cấu hình của nguyên tử khí hiếm nào?")
    if is_teacher:
        sol_tl2 = (
            "Hướng dẫn giải chi tiết Câu 2 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "- Nguyên tử X có 3 lớp electron và lớp thứ ba có 6 electron:\n"
            "    + Lớp 1 (n = 1): 2 electron (1s^2).\n"
            "    + Lớp 2 (n = 2): 8 electron (2s^2 2p^6).\n"
            "    + Lớp 3 (n = 3): 6 electron (3s^2 3p^4). (0,2đ)\n"
            "- Cấu hình electron theo phân lớp: 1s^2 2s^2 2p^6 3s^2 3p^4 (hoặc [Ne] 3s^2 3p^4). (0,15đ)\n"
            "- Cấu hình electron theo lớp của X: (2, 8, 6). (0,1đ)\n"
            "- Tổng số electron = 2 + 8 + 6 = 16 => Số hiệu nguyên tử của X là Z = 16 (Sulfur - S). (0,05đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Biểu diễn sự phân bố electron theo ô orbital ở lớp ngoài cùng (3s^2 3p^4):\n"
            "    + Phân lớp 3s có 1 ô orbital chứa 2 electron ghép đôi: [↑↓]_3s.\n"
            "    + Phân lớp 3p có 3 ô orbital chứa 4 electron: [↑↓][↑ ][↑ ]_3p. (0,2đ)\n"
            "- Dựa vào sơ đồ ô orbital, nguyên tử X có 2 electron độc thân. (0,15đ)\n"
            "- Nguyên tử X có 6 electron lớp ngoài cùng, xu hướng nhận thêm 2 electron để đạt cấu hình bền vững của khí hiếm:\n"
            "    X + 2e -> X^2-   (hoặc S + 2e -> S^2-)\n"
            "    Cấu hình electron của ion X^2-: 1s^2 2s^2 2p^6 3s^2 3p^6 (có 18 electron).\n"
            "- Cấu hình này giống cấu hình electron của nguyên tử khí hiếm Argon (^18Ar). (0,15đ)"
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=9)

    # Câu 3 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. (1,0 điểm) X và Y là hai nguyên tố thuộc cùng một chu kì và thuộc hai nhóm A liên tiếp trong bảng tuần hoàn, Z_X < Z_Y. Tổng số hiệu nguyên tử của X và Y bằng 25.\na) Hãy xác định số hiệu nguyên tử của X và Y.\nb) Viết cấu hình electron của X, Y và xác định vị trí của mỗi nguyên tố trong bảng tuần hoàn.\nc) Dựa vào cấu hình electron lớp ngoài cùng, cho biết X, Y là kim loại hay phi kim và thuộc loại nguyên tố s, p, d hay f.")
    if is_teacher:
        sol_tl3 = (
            "Hướng dẫn giải chi tiết Câu 3 (1,0 điểm):\n"
            "a) (0,3 điểm)\n"
            "- Vì X và Y thuộc cùng một chu kì và ở hai nhóm A liên tiếp trong BTH, với Z_X < Z_Y, nên số hiệu nguyên tử hơn kém nhau 1 đơn vị:\n"
            "    Z_Y - Z_X = 1  <=>  Z_Y = Z_X + 1 (1)\n"
            "- Tổng số hiệu nguyên tử của X và Y bằng 25:\n"
            "    Z_X + Z_Y = 25 (2)\n"
            "- Thay (1) vào (2): Z_X + (Z_X + 1) = 25  <=>  2Z_X = 24  <=>  Z_X = 12.\n"
            "- Suy ra: Z_Y = 12 + 1 = 13.\n"
            "- Vậy số hiệu nguyên tử của X là 12 (Magnesium - Mg), của Y là 13 (Aluminium - Al). (0,3đ)\n\n"
            "b) (0,4 điểm)\n"
            "- Nguyên tố X (Z_X = 12):\n"
            "    + Cấu hình electron: 1s^2 2s^2 2p^6 3s^2 (hoặc [Ne]3s^2).\n"
            "    + Vị trí trong bảng tuần hoàn: Ô số 12 (Z = 12); Chu kì 3 (có 3 lớp e); Nhóm IIA (nguyên tố s, có 2 electron hóa trị ở lớp ngoài cùng). (0,2đ)\n"
            "- Nguyên tố Y (Z_Y = 13):\n"
            "    + Cấu hình electron: 1s^2 2s^2 2p^6 3s^2 3p^1 (hoặc [Ne]3s^2 3p^1).\n"
            "    + Vị trí trong bảng tuần hoàn: Ô số 13 (Z = 13); Chu kì 3 (có 3 lớp e); Nhóm IIIA (nguyên tố p, có 3 electron hóa trị ở lớp ngoài cùng). (0,2đ)\n\n"
            "c) (0,3 điểm)\n"
            "- Nguyên tố X (Z = 12): Có 2 electron ở lớp ngoài cùng (3s^2) => X là kim loại. Electron cuối cùng điền vào phân lớp s => X là nguyên tố s. (0,15đ)\n"
            "- Nguyên tố Y (Z = 13): Có 3 electron ở lớp ngoài cùng (3s^2 3p^1) => Y là kim loại. Electron cuối cùng điền vào phân lớp p => Y là nguyên tố p. (0,15đ)"
        )
        add_solution_no_table(doc, sol_tl3)
    else:
        add_student_writing_lines(doc, num_lines=11)

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
    
    file_student = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_Hoa_10_Ma1004.docx")
    file_teacher = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_Hoa_10_Ma1004_GiaoVien.docx")

    print("Đang tạo Bản Học sinh MÃ ĐỀ 1004...")
    build_exam_buithixuan_1004(is_teacher=False, output_path=file_student)

    print("Đang tạo Bản Giáo viên MÃ ĐỀ 1004...")
    build_exam_buithixuan_1004(is_teacher=True, output_path=file_teacher)

    print("\n[HOÀN TẤT] Cả 2 file Word MÃ ĐỀ 1004 đã được tạo chuẩn 100% FORM ĐỀ KIỂM TRA!")

if __name__ == "__main__":
    main()
