# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì 1 Hóa Học 10 - THPT Hai Bà Trưng (2026 - 2027) - Mã Đề 485
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_Hoa_10_Ma485.docx
2. Bản Giáo viên: De_Kiem_Tra_Hoa_10_Ma485_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
    
    # Chiều rộng: cột đầu 2.4 cm, 12 cột sau mỗi cột 1.3 cm -> Tổng = 18.0 cm
    col_widths_1 = [2.4] + [1.3] * 12
    p1_answers = ["C", "A", "B", "B", "A", "A", "C", "B", "C", "B", "D", "B"]

    for row_idx, row in enumerate(tbl1.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for col_idx, cell in enumerate(row.cells):
            cell.width = Cm(col_widths_1[col_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tcPr = cell._tc.get_or_add_tcPr()
            
            # Viền bảng mỏng màu 8BC390 / CCCCCC
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
                # Nền xanh nhạt
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
    col_widths_2 = [2.4, 3.9, 3.9, 3.9, 3.9, 0.0]  # Tổng 18.0 cm (2.4 + 3.9*4 = 18.0)
    # Chia lại: 6 cột: Câu (2.4cm), Lệnh hỏi (4.0cm), Ý a (2.9cm), Ý b (2.9cm), Ý c (2.9cm), Ý d (2.9cm)
    col_widths_2 = [2.4, 2.9, 3.175, 3.175, 3.175, 3.175]
    
    p2_data = [
        ["Câu", "Nội dung", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Hạt & Kí hiệu", "ĐÚNG", "SAI", "SAI", "SAI"],
        ["Câu 2", "Nguyên tử Cr", "SAI", "ĐÚNG", "ĐÚNG", "ĐÚNG"]
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

    # 3. PHẦN III: TRẢ LỜI NGẮN (4 CÂU) & TỰ LUẬN
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(6)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*3. Phần III: Trả lời ngắn (2,0 điểm - mỗi câu đúng 0,5 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_3 = [2.4, 3.9, 3.9, 3.9, 3.9]
    p3_answers = ["4 (chu kì 4)", "2", "3", "17 (amu)"]

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
    col_widths_tl = [2.4, 3.6, 12.0]
    
    tl_summary = [
        ["Câu hỏi", "Điểm số", "Đáp số cốt lõi"],
        ["Câu 1 (TL)", "1,0 điểm", "a) 1s^2 2s^2 2p^6 3s^2 3p^3, phi kim (5 e lớp ngoài cùng).  b) Có 3 electron độc thân."],
        ["Câu 2 (TL)", "1,0 điểm", "a) Nguyên tử khối trung bình K = 39,135 amu.  b) m dd HCl 10% = 3,65 gam."],
        ["Câu 3 (TL)", "1,0 điểm", "a) Hai kim loại nhóm IIA là Mg (24) và Ca (40).  b) V dd HCl 30% = 3,04 mL."]
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

    # Đường phân cách sau bảng đáp án
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(8)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


def build_exam_haibatrung_485(is_teacher=False, output_path=None):
    doc = docx.Document()
    setup_header_footer(doc)

    # 1. Khối tiêu đề đầu đề thi
    build_title_block(
        doc,
        school_line="SỞ GD&ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT HAI BÀ TRƯNG",
        exam_title="KIỂM TRA GIỮA KÌ 1 NĂM HỌC 2026 - 2027",
        subject_line="MÔN HÓA HỌC – Khối lớp 10",
        num_pages="03",
        duration="45 phút",
        exam_code="485",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Si = 28; P = 31; S = 32; Li = 7; Na = 23; K = 39; Be = 9; Mg = 24; Ca = 40; Ba = 137\n(Cho số hiệu nguyên tử: H: 1; N: 7; C: 6; Cl: 17; O: 8; S: 16; Li: 3; Na: 11; K: 19; Be: 4; Mg: 12; Ca: 20. Học sinh không sử dụng bảng hệ thống tuần hoàn các nguyên tố hóa học)",
        is_teacher=is_teacher
    )

    # Nếu là bản giáo viên: Chèn BẢNG ĐÁP ÁN RÚT GỌN ngay trên đầu
    if is_teacher:
        add_quick_answer_table(doc)

    # =========================================================================
    # A. PHẦN TRẮC NGHIỆM: (7 điểm)
    # =========================================================================
    p_sec_a = doc.add_paragraph()
    p_sec_a.paragraph_format.space_before = Pt(6)
    p_sec_a.paragraph_format.space_after = Pt(2)
    render_rich_text(p_sec_a, "*A. PHẦN TRẮC NGHIỆM: (7 điểm)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # PHẦN I
    add_section_header(
        doc,
        "PHẦN I (3 điểm). TRẮC NGHIỆM NHIỀU LỰA CHỌN",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Đồng (Copper) là kim loại dễ dát mỏng và uốn lượn, có khả năng dẫn điện và dẫn nhiệt tốt. Vì thế đồng được sử dụng rất nhiều trong sản xuất các nguyên liệu. Các đồ dùng từ đồng như: dây điện, que hàn đồng, tay cầm và các đồ dùng nội thất trong nhà, các tượng đúc, nam châm điện từ, các động cơ máy móc, ... Trong tự nhiên, nguyên tố Copper có hai đồng vị là ^63^Cu và ^65^Cu. Nguyên tử khối trung bình của Copper là 63,54. Thành phần % số nguyên tử của đồng vị ^65^Cu là")
    opts1 = [
        "A. 50%.",
        "B. 73%.",
        "*C. 27%.*" if is_teacher else "C. 27%.",
        "D. 54%."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Gọi % số nguyên tử của đồng vị ^65^Cu là x (%), suy ra % số nguyên tử của ^63^Cu là (100 - x)%.\nÁp dụng công thức tính nguyên tử khối trung bình:\n    A_tb = [65.x + 63.(100 - x)] / 100 = 63,54\n    <=> 6300 + 2x = 6354  <=>  2x = 54  <=>  x = 27%.\nVậy thành phần % số nguyên tử của ^65^Cu là 27%.\n➔ Đáp án đúng: C.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Nguyên tố X và hợp kim của nó nhẹ và bền, dùng chế tạo vỏ máy bay, tên lửa, còn được sử dụng trong xây dựng, ngành điện và đồ gia dụng. Nguyên tử của nguyên tố X có tổng số hạt proton, electron, neutron là 40, trong đó số hạt mang điện nhiều hơn tổng số hạt không mang điện là 12. Tổng số electron trên các orbital d của X là")
    opts2 = [
        "*A. 0.*" if is_teacher else "A. 0.",
        "B. 6.",
        "C. 3.",
        "D. 4."
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo giả thiết ta có hệ phương trình về số hạt trong nguyên tử X:\n    { 2p + n = 40\n    { 2p - n = 12\nCộng vế với vế: 4p = 52 => p = 13 (nguyên tố Aluminium - Al).\nCấu hình electron của Al (Z = 13): 1s^2 2s^2 2p^6 3s^2 3p^1.\nNguyên tử Al không có electron nào phân bố trên phân lớp d (orbital d).\nVậy tổng số electron trên các orbital d của X là 0.\n➔ Đáp án đúng: A.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Trong số những quá trình kể dưới đây, cho biết đâu là hiện tượng hóa học?")
    opts3 = [
        "A. Hòa tan muối ăn vào nước thu được dung dịch nước muối.",
        "*B. Khử mùi tanh của cá bằng nước giấm ăn.*" if is_teacher else "B. Khử mùi tanh của cá bằng nước giấm ăn.",
        "C. Mở nút chai nước giải khát loại có gas thấy có bọt sủi lên.",
        "D. Sự đông đặc ở mỡ động vật."
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích từng hiện tượng:\n- A, C, D: Đều là các hiện tượng vật lí (hòa tan, giảm áp suất làm khí CO~2~ thoát ra, sự chuyển thể lỏng sang rắn).\n- B: Mùi tanh của cá chủ yếu do các amine gây ra (chất mang tính base yếu). Nước giấm ăn chứa acetic acid (CH~3~COOH) đã xảy ra phản ứng hóa học tạo muối làm mất mùi tanh -> Là hiện tượng hóa học.\n➔ Đáp án đúng: B.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Aluminium là kim loại phổ biến nhất trong vỏ Trái Đất và có trong thành phần của đất sét, khoáng vật cryolit,... Trong bảng tuần hoàn, aluminium thuộc chu kì 3, nhóm IIIA. Cấu hình electron lớp ngoài cùng của nguyên tử aluminium ở trạng thái cơ bản là")
    opts4 = [
        "A. 3s^2 3p^2.",
        "*B. 3s^2 3p^1.*" if is_teacher else "B. 3s^2 3p^1.",
        "C. 3s^1 3p^2.",
        "D. 3s^2 3p^5."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Aluminium thuộc chu kì 3 (có 3 lớp electron) và nhóm IIIA (nguyên tố p, có 3 electron hóa trị ở lớp ngoài cùng).\nCấu hình electron đầy đủ của Al: 1s^2 2s^2 2p^6 3s^2 3p^1.\nDo đó cấu hình electron lớp ngoài cùng (lớp 3) là 3s^2 3p^1.\n➔ Đáp án đúng: B.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 5. Số electron tối đa có trong lớp M là")
    opts5 = [
        "*A. 18.*" if is_teacher else "A. 18.",
        "B. 28.",
        "C. 32.",
        "D. 8."
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Các lớp electron theo thứ tự n = 1, 2, 3, 4 tương ứng là K, L, M, N.\nLớp M có n = 3. Số electron tối đa trong lớp thứ n tính theo công thức 2n^2 = 2 × 3^2 = 18 electron (gồm phân lớp 3s^2, 3p^6, 3d^10).\n➔ Đáp án đúng: A.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 6. X là nguyên tố hóa học có trong thành phần của chất có tác dụng oxi hóa và sát khuẩn cực mạnh, thường được sử dụng với mục đích khử trùng và tẩy trắng trong lĩnh vực thủy sản, dệt nhuộm, xử lí nước cấp, nước thải, nước bể bơi. Nguyên tử X có số hiệu nguyên tử Z = 17. Số electron độc thân có trong nguyên tử X ở trạng thái cơ bản là")
    opts6 = [
        "*A. 1.*" if is_teacher else "A. 1.",
        "B. 2.",
        "C. 3.",
        "D. 4."
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nguyên tử X có Z = 17 (Chlorine - Cl).\nCấu hình electron: 1s^2 2s^2 2p^6 3s^2 3p^5.\nPhân lớp 3p có 5 electron phân bố trên 3 orbital:\n  [↑↓] [↑↓] [↑ ]  => Có 2 cặp electron ghép đôi và 1 electron độc thân.\n➔ Đáp án đúng: A.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 7. Phát biểu nào sau đây đúng?")
    opts7 = [
        "A. Đồng vị là các nguyên tố có cùng số proton nhưng khác nhau số neutron.",
        "B. Số hiệu nguyên tử bằng điện tích hạt nhân nguyên tử.",
        "*C. Nguyên tố hoá học là những nguyên tử có cùng số đơn vị điện tích hạt nhân.*" if is_teacher else "C. Nguyên tố hoá học là những nguyên tử có cùng số đơn vị điện tích hạt nhân.",
        "D. Số khối của nguyên tử bằng tổng số proton và số electron."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích các phát biểu:\n- A sai vì đồng vị là tập hợp các NGUYÊN TỬ (không phải các nguyên tố) của cùng một nguyên tố.\n- B sai vì số hiệu nguyên tử Z bằng SỐ ĐƠN VỊ điện tích hạt nhân (Z là số nguyên dương không thứ nguyên, còn điện tích mang đơn vị Culông).\n- C đúng: Nguyên tố hoá học là tập hợp các nguyên tử có cùng số đơn vị điện tích hạt nhân (cùng Z).\n- D sai vì số khối A = số proton + số neutron (A = Z + N).\n➔ Đáp án đúng: C.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 8. Nội dung nào dưới đây thuộc đối tượng nghiên cứu của hóa học?")
    opts8 = [
        "A. Sự chuyển động các hành tinh trong thái dương hệ.",
        "*B. Tính chất và sự biến đổi của chất.*" if is_teacher else "B. Tính chất và sự biến đổi của chất.",
        "C. Sự vận chuyển của máu trong hệ tuần hoàn.",
        "D. Sự lớn lên và sinh sản của tế bào."
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Đối tượng nghiên cứu của hóa học là các chất, thành phần, cấu trúc, tính chất và sự biến đổi của chất cùng ứng dụng của chúng.\n(A thuộc Thiên văn/Vật lí, C và D thuộc Sinh học/Y học).\n➔ Đáp án đúng: B.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 9. Phát biểu nào dưới đây không đúng?")
    opts9 = [
        "A. Hạt nhân nguyên tử có điện tích dương.",
        "B. Tia âm cực là chùm tia mang điện tích âm chứa các hạt electron.",
        "*C. Tất cả các hạt nhân nguyên tử của mọi nguyên tố đều chứa proton và neutron.*" if is_teacher else "C. Tất cả các hạt nhân nguyên tử của mọi nguyên tố đều chứa proton và neutron.",
        "D. Nguyên tử có cấu trúc rỗng, gồm vỏ nguyên tử và hạt nhân nguyên tử."
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hạt nhân của đồng vị phổ biến nhất của hydrogen là Protium (^1_1H) chỉ có 1 proton và KHÔNG CÓ neutron.\nDo đó phát biểu \"Tất cả các hạt nhân nguyên tử của mọi nguyên tố đều chứa proton và neutron\" là không đúng.\n➔ Đáp án đúng: C.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 10. Cho dãy chất được biểu diễn bằng công thức hóa học như sau: H~2~O, O~3~, KOH, He, Cl~2~, CuO, Fe, NaOH, O~2~, NaBr, Zn, CH~4~. Số đơn chất trong dãy trên là")
    opts10 = [
        "A. 4.",
        "*B. 6.*" if is_teacher else "B. 6.",
        "C. 5.",
        "D. 3."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Đơn chất là chất được tạo nên từ một nguyên tố hóa học duy nhất.\nTrong dãy gồm 12 chất:\n- Đơn chất (6 chất): O~3~, He, Cl~2~, Fe, O~2~, Zn.\n- Hợp chất (6 chất): H~2~O, KOH, CuO, NaOH, NaBr, CH~4~.\n➔ Đáp án đúng: B.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 11. Các hạt nhân của hầu hết các nguyên tử được tạo nên từ hạt")
    opts11 = [
        "A. electron, proton và neutron.",
        "B. electron và neutron.",
        "C. electron và proton.",
        "*D. proton và neutron.*" if is_teacher else "D. proton và neutron."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hạt nhân của hầu hết các nguyên tử được tạo nên từ 2 loại hạt: proton (mang điện tích dương) và neutron (không mang điện). Hạt electron nằm ở lớp vỏ nguyên tử.\n➔ Đáp án đúng: D.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 12. Khái niệm nào sau đây đúng: Nhóm nguyên tố là")
    opts12 = [
        "A. tập hợp các nguyên tố mà nguyên tử có cấu hình electron gần giống nhau, do đó có tính chất hóa học giống nhau và được xếp thành một cột.",
        "*B. tập hợp các nguyên tố mà nguyên tử có cấu hình electron tương tự nhau, do đó có tính chất hóa học gần giống nhau và được xếp thành một cột.*" if is_teacher else "B. tập hợp các nguyên tố mà nguyên tử có cấu hình electron tương tự nhau, do đó có tính chất hóa học gần giống nhau và được xếp thành một cột.",
        "C. tập hợp các nguyên tố mà nguyên tử có cấu hình electron giống nhau, được xếp ở cùng một cột.",
        "D. tập hợp các nguyên tố mà nguyên tử có tính chất hóa học giống nhau và được xếp thành một cột."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo định nghĩa chuẩn SGK Hóa học 10:\n\"Nhóm nguyên tố là tập hợp các nguyên tố mà nguyên tử có cấu hình electron tương tự nhau, do đó có tính chất hóa học gần giống nhau và được xếp thành một cột.\"\n➔ Đáp án đúng: B.")

    # =========================================================================
    # PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN II (2 điểm). CÂU TRẮC NGHIỆM ĐÚNG / SAI",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng (Đ) hoặc sai (S)."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1: Các phát biểu sau đây là đúng hay sai?")
    stmts1 = [
        ("a", "Nguyên tử có kích thước lớn hơn nhiều so với hạt nhân, nhưng có khối lượng gần bằng khối lượng hạt nhân.", True),
        ("b", "Proton là hạt mang mang điện tích âm, có khối lượng gần bằng 1 amu.", False),
        ("c", "Nguyên tử R có điện tích hạt nhân là +41,6.10^-19 C thì hạt nhân nguyên tử R có 24 proton.", False),
        ("d", "Nguyên tử X có 16 electron và 16 neutron, kí hiệu nguyên tử X là ^16~32~S.", False)
    ]
    for label, text, ans_bool in stmts1:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 1:\n- Ý a) ĐÚNG: Đường kính nguyên tử lớn hơn hạt nhân khoảng 10 000 lần, nhưng toàn bộ khối lượng nguyên tử tập trung hầu như ở hạt nhân do khối lượng của electron rất nhỏ.\n- Ý b) SAI: Proton là hạt mang điện tích dương (+1), không phải điện tích âm.\n- Ý c) SAI: Số proton = (41,6.10^-19) / (1,6.10^-19) = 26 proton (đây là sắt - Fe), không phải 24 proton.\n- Ý d) SAI: Số electron = 16 => Z = 16. Số khối A = Z + N = 16 + 16 = 32. Kí hiệu nguyên tử chuẩn là ^32_16S (số khối ở trên, số hiệu Z ở dưới). Kí hiệu trong đề ^16_32S bị đảo ngược.\n➔ Kết luận: a - Đ, b - S, c - S, d - S.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2: Nguyên tử Chromium (Cr) có kí hiệu là ^52~24~Cr. Cho các phát biểu sau:")
    stmts2 = [
        ("a", "Cr là nguyên tố s.", False),
        ("b", "Nguyên tử Cr có 6 electron hoá trị.", True),
        ("c", "Trong bảng tuần hoàn Cr thuộc nhóm VIB.", True),
        ("d", "Trong nguyên tử Cr số hạt không mang điện chiếm khoảng 53,85% về khối lượng nguyên tử.", True)
    ]
    for label, text, ans_bool in stmts2:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 2:\n- Cr có Z = 24, A = 52 => Số p = 24; e = 24; n = 52 - 24 = 28.\n- Cấu hình electron của Cr: 1s^2 2s^2 2p^6 3s^2 3p^6 3d^5 4s^1 ([Ar] 3d^5 4s^1).\n- Ý a) SAI: Cr có electron điền cuối cùng vào phân lớp 3d nên Cr là nguyên tố d.\n- Ý b) ĐÚNG: Số electron hóa trị của Cr = số e ở 3d + số e ở 4s = 5 + 1 = 6 electron.\n- Ý c) ĐÚNG: Cr là nguyên tố d có 6 electron hóa trị và cấu hình (n-1)d^5 ns^1 nên thuộc nhóm VIB, chu kì 4.\n- Ý d) ĐÚNG: Hạt không mang điện là neutron (28 hạt). Khối lượng nguyên tử xấp xỉ bằng số khối A = 52 amu.\n  Tỉ lệ % về khối lượng của hạt không mang điện: %m_n = (28 / 52) × 100% ≈ 53,846% ≈ 53,85%.\n➔ Kết luận: a - S, b - Đ, c - Đ, d - Đ.")

    # =========================================================================
    # PHẦN III: TRẢ LỜI NGẮN
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN III (2 điểm). TRẮC NGHIỆM TRẢ LỜI NGẮN",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Viết câu trả lời vào ô quy định."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1: Nguyên tử của nguyên tố calcium nằm ở ô thứ 20 trong bảng HTTH các nguyên tố. Calcium thuộc chu kì ?")
    if is_teacher:
        add_solution_no_table(doc, "Calcium có Z = 20. Cấu hình electron: 1s^2 2s^2 2p^6 3s^2 3p^6 4s^2 (có 4 lớp electron).\nSố lớp electron bằng số thứ tự chu kì => Calcium thuộc chu kì 4.\n➔ Đáp số: 4")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2: Các electron của nguyên tử nguyên tố X được phân bố trên ba lớp, lớp thứ ba có 6 electron. Số electron độc thân của nguyên tử X là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Nguyên tử X có 3 lớp electron và lớp thứ ba có 6 electron:\n- Lớp 1 (n=1): 2 electron (1s^2)\n- Lớp 2 (n=2): 8 electron (2s^2 2p^6)\n- Lớp 3 (n=3): 6 electron (3s^2 3p^4)\nCấu hình electron đầy đủ: 1s^2 2s^2 2p^6 3s^2 3p^4 (Z = 16, Sulfur).\nPhân lớp 3p^4 có 3 orbital: [↑↓] [↑ ] [↑ ] => có 2 electron độc thân.\n➔ Đáp số: 2")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3: Cho các phát biểu sau:\n(1) Ở trạng thái cơ bản, trong một nguyên tử luôn có số proton bằng số electron.\n(2) Tổng số proton và số neutron trong hạt nhân được gọi là số khối.\n(3) Các đồng vị có tính chất vật lí giống nhau nhưng tính chất hoá học khác nhau.\n(4) Số neutron bằng số đơn vị điện tích hạt nhân nguyên tử.\n(5) Trong nguyên tử, hạt electron có khối lượng không đáng kể so với các hạt còn lại.\nTrong các phát biểu trên có bao nhiêu phát biểu đúng?")
    if is_teacher:
        add_solution_no_table(doc, "Xét từng phát biểu:\n(1) ĐÚNG: Nguyên tử luôn trung hòa về điện nên số p = số e.\n(2) ĐÚNG: Số khối A = Z + N (tổng số proton và neutron trong hạt nhân).\n(3) SAI: Các đồng vị có cùng số proton và cấu hình electron nên tính chất hóa học tương tự nhau, nhưng có tính chất vật lí khác nhau do khối lượng khác nhau.\n(4) SAI: Số neutron N không nhất thiết bằng Z (ví dụ ^1_1H có N=0, các nguyên tố nặng có N > Z).\n(5) ĐÚNG: Khối lượng electron m_e ≈ 0,00055 amu rất nhỏ so với proton và neutron (~1 amu).\nCác phát biểu đúng gồm: (1), (2), (5) => Tổng cộng có 3 phát biểu đúng.\n➔ Đáp số: 3")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4: Một nguyên tố tạo hợp chất khí với hydrogen có công thức RH~3~, được sử dụng để trung hòa các thành phần acid của dầu thô, bảo vệ thiết bị không bị ăn mòn trong ngành công nghiệp dầu khí. Nguyên tố này chiếm 25,93% về khối lượng trong oxide cao nhất. Phân tử khối của RH~3~ có giá trị bao nhiêu amu?")
    if is_teacher:
        add_solution_no_table(doc, "Hợp chất khí với hydrogen là RH~3~ => Hóa trị của R trong hợp chất với hydrogen là 3.\nDo tổng hóa trị trong hợp chất với H và oxide cao nhất bằng 8, nên hóa trị của R trong oxide cao nhất là 8 - 3 = 5.\nCông thức oxide cao nhất là R~2~O~5~.\nTheo đề bài, % khối lượng của R trong R~2~O~5~ là 25,93%:\n    %m_R = [2R / (2R + 16 × 5)] × 100% = 25,93%\n    <=> 2R = 0,2593 × (2R + 80)\n    <=> 2R - 0,5186R = 20,744\n    <=> 1,4814R = 20,744  <=>  R ≈ 14 (amu).\nVậy R là nguyên tố Nitrogen (N = 14).\nHợp chất khí RH~3~ là NH~3~ (ammonia, được dùng trung hòa acid chống ăn mòn trong dầu thô).\nPhân tử khối của NH~3~: M = 14 + 1 × 3 = 17 amu.\n➔ Đáp số: 17")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # =========================================================================
    # B. PHẦN TỰ LUẬN: (3 điểm)
    # =========================================================================
    p_sec_b = doc.add_paragraph()
    p_sec_b.paragraph_format.space_before = Pt(8)
    p_sec_b.paragraph_format.space_after = Pt(2)
    render_rich_text(p_sec_b, "*B. PHẦN TỰ LUẬN: (3 điểm)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # Câu 1 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1: (1,0 điểm) Cho nguyên tử của nguyên tố phosphorus có Z = 15.\na) Viết cấu hình electron nguyên tử phosphorus; cho biết phosphorus là kim loại, phi kim hay khí hiếm. Giải thích?\nb) Biểu diễn cấu hình electron của phosphorus theo ô orbital. Cho biết số electron độc thân của nguyên tử phosphorus.")
    if is_teacher:
        sol_tl1 = (
            "Hướng dẫn giải chi tiết Câu 1 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "- Phosphorus có Z = 15, cấu hình electron: 1s^2 2s^2 2p^6 3s^2 3p^3 (hoặc [Ne]3s^2 3p^3). (0,25đ)\n"
            "- Lớp ngoài cùng (lớp 3) có 2 + 3 = 5 electron.\n"
            "- Vì có 5 electron ở lớp ngoài cùng nên phosphorus là phi kim. (0,25đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Biểu diễn electron theo ô orbital:\n"
            "    Lớp 1: 1s [↑↓]\n"
            "    Lớp 2: 2s [↑↓]   2p [↑↓][↑↓][↑↓]\n"
            "    Lớp 3: 3s [↑↓]   3p [↑ ][↑ ][↑ ] (0,25đ)\n"
            "- Nguyên tử phosphorus có 3 electron độc thân (nằm ở 3 orbital 3p). (0,25đ)"
        )
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 2 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2: (1,0 điểm)\na. Nguyên tố potassium (kí hiệu: K) có các đồng vị: ^39~19~K (chiếm 93,258%), ^40~19~K (chiếm 0,012%) và ^41~19~K (chiếm 6,730%). Tính nguyên tử khối trung bình của potassium.\nb. Cần bao nhiêu gam dung dịch HCl 10% để trung hoà 100 ml dung dịch có chứa 0,56 gam potassium hydroxide?")
    if is_teacher:
        sol_tl2 = (
            "Hướng dẫn giải chi tiết Câu 2 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "Nguyên tử khối trung bình của potassium:\n"
            "    A_tb = (39 × 93,258 + 40 × 0,012 + 41 × 6,730) / 100\n"
            "         = (3637,062 + 0,480 + 275,930) / 100 = 3913,472 / 100 = 39,135 amu (hoặc ≈ 39,13 amu). (0,5đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Phương trình phản ứng trung hòa:\n"
            "    KOH + HCl -> KCl + H~2~O (1) (0,15đ)\n"
            "- Số mol KOH: n_KOH = 0,56 / 56 = 0,01 mol. (0,1đ)\n"
            "- Theo (1): n_HCl = n_KOH = 0,01 mol => m_HCl = 0,01 × 36,5 = 0,365 gam. (0,1đ)\n"
            "- Khối lượng dung dịch HCl 10% cần dùng:\n"
            "    m_dd HCl = (0,365 × 100) / 10 = 3,65 gam. (0,15đ)"
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 3 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3: (1,0 điểm) Hoà tan hoàn toàn 6,9081 g hỗn hợp muối cacbonat của 2 kim loại kế tiếp nhau trong nhóm IIA vào dd HCl dư. Lượng khí sinh ra đem sục từ từ vào dung dịch nước vôi trong (lấy dư) sau kết thúc thấy có 7,5 gam chất kết tủa trắng.\na. Xác định Hai kim loại trong muối.\nb. Cần dùng bao nhiêu ml dung dịch HCl 30%, d = 1,5 gam/ml để phản ứng hết 1,727 gam hỗn hợp muối trên.")
    if is_teacher:
        sol_tl3 = (
            "Hướng dẫn giải chi tiết Câu 3 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "- Gọi công thức chung của 2 muối cacbonat là M_tbCO~3~.\n"
            "- Phương trình phản ứng:\n"
            "    M_tbCO~3~ + 2HCl -> M_tbCl~2~ + CO~2~ + H~2~O (1)\n"
            "    CO~2~ + Ca(OH)~2~ dư -> CaCO~3~↓ + H~2~O (2) (0,15đ)\n"
            "- Kết tủa trắng là CaCO~3~: n_CaCO3 = 7,5 / 100 = 0,075 mol.\n"
            "- Theo (1) và (2): n_CO2 = n_MtbCO3 = n_CaCO3 = 0,075 mol. (0,1đ)\n"
            "- Khối lượng mol trung bình của muối:\n"
            "    M_tb(muối) = M_tb + 60 = 6,9081 / 0,075 = 92,108 g/mol\n"
            "    => M_tb = 92,108 - 60 = 32,108 g/mol. (0,15đ)\n"
            "- Do 2 kim loại thuộc nhóm IIA và ở hai chu kì liên tiếp (kế tiếp nhau):\n"
            "    M_A < 32,108 < M_B\n"
            "    Tra bảng các kim loại nhóm IIA: Be (9) < Mg (24) < Ca (40) < Sr (88) < Ba (137).\n"
            "    => 2 kim loại kế tiếp là Magnesium (Mg, M=24) và Calcium (Ca, M=40). (0,1đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Trong 1,727 gam hỗn hợp muối ban đầu (nhận thấy 1,727 = 6,9081 / 4):\n"
            "    Số mol muối = 1,727 / 92,108 = 0,01875 mol. (0,15đ)\n"
            "- Theo PTHH (1): n_HCl = 2 × n_muối = 2 × 0,01875 = 0,0375 mol. (0,1đ)\n"
            "- Khối lượng chất tan HCl: m_HCl = 0,0375 × 36,5 = 1,36875 gam.\n"
            "- Khối lượng dung dịch HCl 30%:\n"
            "    m_dd HCl = (1,36875 × 100) / 30 = 4,5625 gam. (0,15đ)\n"
            "- Thể tích dung dịch HCl 30% cần dùng:\n"
            "    V_dd HCl = m_dd / d = 4,5625 / 1,5 ≈ 3,0417 ml ≈ 3,04 ml. (0,1đ)"
        )
        add_solution_no_table(doc, sol_tl3)
    else:
        add_student_writing_lines(doc, num_lines=10)

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
    
    file_student = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_Hoa_10_Ma485.docx")
    file_teacher = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_Hoa_10_Ma485_GiaoVien.docx")

    print("Đang tạo Bản Học sinh...")
    build_exam_haibatrung_485(is_teacher=False, output_path=file_student)

    print("Đang tạo Bản Giáo viên...")
    build_exam_haibatrung_485(is_teacher=True, output_path=file_teacher)

    print("\n[HOÀN TẤT] Cả 2 file Word đã được tạo chuẩn 100% FORM ĐỀ KIỂM TRA!")

if __name__ == "__main__":
    main()
