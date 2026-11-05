# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì 1 Hóa Học 10 - THPT Hai Bà Trưng (2026 - 2027) - Mã Đề 132
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_Hoa_10_Ma132.docx
2. Bản Giáo viên: De_Kiem_Tra_Hoa_10_Ma132_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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

def add_quick_answer_table_132(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu Mã đề 132.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 132)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (12 CÂU)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - mỗi câu đúng 0,25 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    col_widths_1 = [2.4] + [1.3] * 12
    p1_answers = ["A", "C", "C", "D", "D", "C/D*", "C", "B", "B", "A", "D", "A"]

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

    p_note_c6 = doc.add_paragraph()
    p_note_c6.paragraph_format.space_before = Pt(1)
    p_note_c6.paragraph_format.space_after = Pt(2)
    render_rich_text(p_note_c6, "_(Ghi chú Câu 6: Xem phân tích chi tiết trong phần Lời giải chi tiết bên dưới)._", base_size=9, base_italic=True, base_color=COLOR_GRAY)

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
        ["Câu 1", "Hạt & Kí hiệu", "ĐÚNG", "ĐÚNG", "SAI", "SAI"],
        ["Câu 2", "Nguyên tử Fe", "ĐÚNG", "ĐÚNG", "SAI", "SAI"]
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
    p3_answers = ["4 (phát biểu đúng)", "20", "15", "28 (hạt)"]

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
        ["Câu 1 (TL)", "1,0 điểm", "a) 1s^2 2s^2 2p^6 3s^2 3p^5.  b) Oxide cao nhất: X~2~O~7~ (Cl~2~O~7~); Hợp chất khí: HX (HCl)."],
        ["Câu 2 (TL)", "1,0 điểm", "Nguyên tử khối của đồng vị thứ hai là A~2~ = 81 (amu)."],
        ["Câu 3 (TL)", "1,0 điểm", "a) Hai muối cacbonat là MgCO~3~ và CaCO~3~.  b) V dd HCl 30% = 9,73 mL."]
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


def build_exam_haibatrung_132(is_teacher=False, output_path=None):
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
        exam_code="132",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Si = 28; P = 31; S = 32; Li = 7; Na = 23; K = 39; Be = 9; Mg = 24; Ca = 40; Ba = 137\n(Cho số hiệu nguyên tử: H: 1; N: 7; C: 6; Cl: 17; O: 8; S: 16; Li: 3; Na: 11; K: 19; Be: 4; Mg: 12; Ca: 20. Học sinh không sử dụng bảng hệ thống tuần hoàn các nguyên tố hóa học)",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_132(doc)

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
    render_question_prompt(p, "Câu 1. Silicon được dùng trong công nghệ sản xuất chip máy tính hiện đại.. Vị trí của nguyên tố silicon trong bảng tuần hoàn thuộc ô 14, chu kì 3, nhóm IVA. Hãy chọn nhận đúng về silicon.")
    opts1 = [
        "*A. Cấu hình electron nguyên tử của nguyên tố silicon là [Ne]3s^2 3p^2.*" if is_teacher else "A. Cấu hình electron nguyên tử của nguyên tố silicon là [Ne]3s^2 3p^2.",
        "B. Lớp N có 4 electron hoá trị.",
        "C. Số điện tích hạt nhân của silicon bằng 28.",
        "D. Silic nằm ở hàng thứ 4."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Silicon ở ô 14 (Z = 14) => Cấu hình electron: 1s^2 2s^2 2p^6 3s^2 3p^2 = [Ne]3s^2 3p^2 (A đúng).\n- B sai vì 4 electron hóa trị nằm ở lớp M (lớp ngoài cùng n = 3), không phải lớp N (n = 4).\n- C sai vì số đơn vị điện tích hạt nhân bằng Z = 14 (hoặc điện tích hạt nhân là +14), số khối mới xấp xỉ 28.\n- D sai vì chu kì 3 tương ứng với hàng thứ 3 của bảng tuần hoàn.\n➔ Đáp án đúng: A.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Hầu hết các hạt nhân nguyên tử được tạo nên từ các hạt")
    opts2 = [
        "A. electron, proton và neutron.",
        "B. electron và neutron.",
        "*C. proton và neutron.*" if is_teacher else "C. proton và neutron.",
        "D. electron và proton."
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hạt nhân của hầu hết các nguyên tử được tạo nên từ hai loại hạt cơ bản: proton (mang điện tích dương) và neutron (không mang điện). Hạt electron nằm ở lớp vỏ nguyên tử.\n➔ Đáp án đúng: C.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Trong số những quá trình kể dưới đây, cho biết đâu là hiện tượng vật lí ?")
    opts3 = [
        "A. Dẫn khí carbon dioxide vào nước vôi trong, làm nước vôi trong vẩn đục.",
        "B. Cho vôi sống (CaO) vào nước được Ca(OH)~2~.",
        "*C. Nước bị hoá hơi khi đun nóng.*" if is_teacher else "C. Nước bị hoá hơi khi đun nóng.",
        "D. Đốt cháy đường mía tạo thành chất màu đen và có mùi khét."
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích các quá trình:\n- A: CO~2~ + Ca(OH)~2~ -> CaCO~3~↓ + H~2~O (phản ứng hóa học).\n- B: CaO + H~2~O -> Ca(OH)~2~ (phản ứng hóa học).\n- C: Nước lỏng chuyển sang trạng thái hơi (sự biến đổi trạng thái, không tạo chất mới) -> Là hiện tượng vật lí.\n- D: C~12~H~22~O~11~ -> 12C + 11H~2~O (phản ứng phân hủy hóa học).\n➔ Đáp án đúng: C.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Số phân lớp electron có trong lớp O là")
    opts4 = [
        "A. 2.",
        "B. 4.",
        "C. 3.",
        "*D. 5.*" if is_teacher else "D. 5."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Tên các lớp electron theo số lượng n tăng dần: K (n=1), L (n=2), M (n=3), N (n=4), O (n=5).\nTheo lí thuyết cấu tạo nguyên tử, lớp thứ n có n phân lớp electron. Do đó lớp O (n = 5) có 5 phân lớp electron (5s, 5p, 5d, 5f, 5g).\n➔ Đáp án đúng: D.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 5. Phát biểu nào sau đây không đúng?")
    opts5 = [
        "A. Số khối của nguyên tử bằng tổng số proton và số neutron.",
        "B. Nguyên tố hoá học là những nguyên tử có cùng số đơn vị điện tích hạt nhân.",
        "C. Số hiệu nguyên tử bằng số đơn vị điện tích hạt nhân nguyên tử.",
        "*D. Đồng vị là các nguyên tố có cùng số proton nhưng khác nhau số neutron.*" if is_teacher else "D. Đồng vị là các nguyên tố có cùng số proton nhưng khác nhau số neutron."
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích các phát biểu:\n- A, B, C đều là các định nghĩa chuẩn xác trong SGK Hóa học 10.\n- D không đúng vì đồng vị là tập hợp các NGUYÊN TỬ (không phải \"các nguyên tố\") của cùng một nguyên tố hóa học, có cùng số proton nhưng khác nhau số neutron.\n➔ Đáp án đúng: D.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 6. Phát biểu nào sau đây đúng?")
    opts6 = [
        "A. Electron ở orbital 3p có mức năng lượng thấp hơn electron ở orbital 3s.",
        "B. Những electron ở gần hạt nhân có mức năng lượng cao nhất.",
        "*C. Những electron ở lớp K có mức năng lượng cao nhất.*" if is_teacher else "C. Những electron ở lớp K có mức năng lượng cao nhất.",
        "D. Các electron trong cùng một lớp có năng lượng bằng nhau."
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích sư phạm chi tiết Câu 6:\n- Theo chuẩn kiến thức SGK Hóa học 10:\n  + A sai: Thứ tự mức năng lượng là 1s < 2s < 2p < 3s < 3p => electron ở 3p có năng lượng cao hơn 3s.\n  + B & C: Các electron càng gần hạt nhân liên kết càng bền chặt và có mức năng lượng THẤP NHẤT (lớp K có n = 1 có năng lượng thấp nhất). Do đó phát biểu B và C trong đề in \"cao nhất\" là lỗi in ấn của đề gốc (đúng ra phải là \"thấp nhất\").\n  + D: Theo định nghĩa SGK, các electron trên cùng một PHÂN LỚP có mức năng lượng bằng nhau; còn trên cùng một LỚP có mức năng lượng gần bằng nhau (không hoàn toàn bằng nhau).\n* Lưu ý cho giáo viên chấm thi: Nếu đề sửa \"thấp nhất\" ở phương án C thì chọn C; nếu đề tham khảo theo quan niệm đơn giản coi các e cùng lớp có năng lượng tương đương nhau thì chọn D. Học sinh được linh hoạt chấm điểm theo quy định của tổ chuyên môn.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 7. Tổng số các hạt proton, nơtron và electron của một nguyên tử nguyên tố X là 28. Trong đó số hạt mang điện nhiều hơn số hạt không mang điện là 8. X là:")
    opts7 = [
        "A. Cl.",
        "B. Al.",
        "*C. F.*" if is_teacher else "C. F.",
        "D. Si."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo giả thiết bài toán:\n    { 2p + n = 28\n    { 2p - n = 8\nCộng hai phương trình: 4p = 36 => p = 9 (Fluorine - F).\nSố neutron: n = 28 - 18 = 10.\nVậy X là nguyên tố Fluorine (F).\n➔ Đáp án đúng: C.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 8. Đồng (Copper) là kim loại dễ dát mỏng và uốn lượn, có khả năng dẫn điện và dẫn nhiệt tốt. Vì thế đồng được sử dụng rất nhiều trong sản xuất các nguyên liệu. Các đồ dùng từ đồng như: dây điện, que hàn đồng, tay cầm và các đồ dùng nội thất trong nhà, các tượng đúc, nam châm điện từ, các động cơ máy móc, ....Trong tự nhiên, nguyên tố Copper có hai đồng vị là ^63^Cu và ^65^Cu. Nguyên tử khối trung bình của Copper là 63,54. Thành phần % số nguyên tử của đồng vị ^63^Cu là")
    opts8 = [
        "A. 27%.",
        "*B. 73%.*" if is_teacher else "B. 73%.",
        "C. 50%.",
        "D. 54%."
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Gọi % số nguyên tử của đồng vị ^63^Cu là x (%), suy ra % của ^65^Cu là (100 - x)%.\nÁp dụng công thức tính nguyên tử khối trung bình:\n    A_tb = [63.x + 65.(100 - x)] / 100 = 63,54\n    <=> 6500 - 2x = 6354  <=>  2x = 146  <=>  x = 73%.\nVậy thành phần % số nguyên tử của ^63^Cu là 73%.\n➔ Đáp án đúng: B.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 9. Phát biểu nào sau đây về chu kì là đúng?")
    opts9 = [
        "A. Dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được xếp theo chiều số neutron tăng dần.",
        "*B. Mở đầu mỗi chu kì là kim loại điển hình và kết thúc là một khí hiếm (trừ chu kì 1).*" if is_teacher else "B. Mở đầu mỗi chu kì là kim loại điển hình và kết thúc là một khí hiếm (trừ chu kì 1).",
        "C. Dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được xếp theo chiều khối lượng nguyên tử tăng dần.",
        "D. Dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được xếp theo chiều số khối tăng dần."
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Quy luật biến đổi trong chu kì theo SGK Hóa học 10:\n\"Mở đầu mỗi chu kì là một kim loại kiềm (kim loại điển hình) và kết thúc là một khí hiếm (trừ chu kì 1 bắt đầu bằng phi kim H)\".\nCác phát biểu A, C, D đều sai vì các nguyên tố trong cùng một chu kì được sắp xếp theo chiều TĂNG DẦN CỦA ĐIỆN TÍCH HẠT NHÂN (Z), không phải theo chiều tăng của neutron, khối lượng hay số khối.\n➔ Đáp án đúng: B.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 10. X là nguyên tố hóa học có trong thành phần của chất có tác dụng oxi hóa và sát khuẩn cực mạnh, thường được sử dụng với mục đích khử trùng và tẩy trắng trong lĩnh vực thủy sản, dệt nhuộm, xử lí nước cấp, nước thải, nước bể bơi. Nguyên tử X có số hiệu nguyên tử Z = 17. Số electron lớp hoá trị trong nguyên tử X là")
    opts10 = [
        "*A. 7.*" if is_teacher else "A. 7.",
        "B. 6.",
        "C. 5.",
        "D. 4."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nguyên tử X có Z = 17 (Chlorine - Cl).\nCấu hình electron: 1s^2 2s^2 2p^6 3s^2 3p^5.\nX là nguyên tố p thuộc nhóm VIIA. Đối với nguyên tố nhóm A, số electron hóa trị bằng số electron lớp ngoài cùng: 2 + 5 = 7 electron hóa trị.\n➔ Đáp án đúng: A.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 11. Nội dung nào dưới đây không phải là đối tượng nghiên cứu của hóa học?")
    opts11 = [
        "A. Ứng dụng của chất.",
        "B. Tính chất và sự biến đổi của chất.",
        "C. Thành phần cấu trúc của chất.",
        "*D. Sự lớn lên và sinh sản của tế bào.*" if is_teacher else "D. Sự lớn lên và sinh sản của tế bào."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Đối tượng nghiên cứu của hóa học là chất và sự biến đổi của chất, thành phần, cấu trúc và ứng dụng của chúng.\nSự lớn lên và sinh sản của tế bào là đối tượng nghiên cứu của Sinh học.\n➔ Đáp án đúng: D.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 12. Cho dãy chất được biểu diễn bằng công thức hóa học như sau: HNO~3~, O~3~, NaOH, He, Cl~2~, CuO, H~2~SO~4~, NaBr, Zn, CH~4~, MnO~2~. Số hợp chất trong dãy trên là")
    opts12 = [
        "*A. 7.*" if is_teacher else "A. 7.",
        "B. 5.",
        "C. 4.",
        "D. 6."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hợp chất là chất được cấu tạo từ hai nguyên tố hóa học trở lên.\nTrong dãy gồm 11 chất:\n- Hợp chất (7 chất): HNO~3~, NaOH, CuO, H~2~SO~4~, NaBr, CH~4~, MnO~2~.\n- Đơn chất (4 chất): O~3~, He, Cl~2~, Zn.\nVậy số hợp chất trong dãy là 7.\n➔ Đáp án đúng: A.")

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
        ("a", "Proton là hạt mang điện tích dương, nằm trong hạt nhân, khối lượng gần bằng 1 amu.", True),
        ("b", "Nguyên tử trung hòa điện, có kích thước lớn hơn nhiều so với hạt nhân, nhưng có khối lượng gần bằng khối lượng hạt nhân.", True),
        ("c", "Nguyên tử R có điện tích lớp vỏ là -40,6.10^-19 C thì hạt nhân nguyên tử R có 26 proton.", False),
        ("d", "Nguyên tử X có 32 electron và 35 neutron, kí hiệu nguyên tử X là ^16~32~S.", False)
    ]
    for label, text, ans_bool in stmts1:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 1:\n- Ý a) ĐÚNG: Hạt proton mang điện tích dương (+1), nằm trong hạt nhân nguyên tử và có khối lượng m_p ≈ 1,673.10^-27 kg ≈ 1 amu.\n- Ý b) ĐÚNG: Nguyên tử trung hòa về điện (số p = số e), kích thước lớn gấp khoảng 10.000 lần hạt nhân nhưng khối lượng tập trung chủ yếu ở hạt nhân.\n- Ý c) SAI: Điện tích của 1 electron là -1,602.10^-19 C. Nếu R có 26 proton thì lớp vỏ phải có 26 electron và điện tích lớp vỏ là -26 × 1,602.10^-19 C ≈ -41,6.10^-19 C (không phải -40,6.10^-19 C).\n- Ý d) SAI: Nguyên tử X có 32 electron => Z = 32 (Germanium - Ge), số khối A = 32 + 35 = 67. Kí hiệu đúng là ^67_32Ge, không phải ^16_32S.\n➔ Kết luận: a - Đ, b - Đ, c - S, d - S.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2: Nguyên tử iron (Fe) có kí hiệu là ^56~26~Fe. Cho các phát biểu sau:")
    stmts2 = [
        ("a", "Nguyên tử Fe có 2 electron lớp ngoài cùng.", True),
        ("b", "Fe là nguyên tố d.", True),
        ("c", "Trong bảng tuần hoàn Fe thuộc nhóm IIA.", False),
        ("d", "Trong nguyên tử Fe số hạt không mang điện chiếm 39,6%. Về khối lượng nguyên tử.", False)
    ]
    for label, text, ans_bool in stmts2:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 2:\n- Nguyên tử Fe có Z = 26, số khối A = 56 => Số p = 26, e = 26, n = 56 - 26 = 30 hạt.\n- Cấu hình electron: 1s^2 2s^2 2p^6 3s^2 3p^6 3d^6 4s^2 ([Ar] 3d^6 4s^2).\n- Ý a) ĐÚNG: Lớp ngoài cùng là lớp thứ 4 (4s^2) có đúng 2 electron.\n- Ý b) ĐÚNG: Electron cuối cùng điền vào phân lớp 3d nên Fe là nguyên tố d.\n- Ý c) SAI: Fe có cấu hình (n-1)d^6 ns^2, tổng số e hóa trị là 6 + 2 = 8 => Fe thuộc nhóm VIIIB, chu kì 4 (không phải nhóm IIA).\n- Ý d) SAI: Hạt không mang điện là neutron (30 hạt). Khối lượng nguyên tử xấp xỉ bằng số khối A = 56 amu.\n  Tỉ lệ % về khối lượng của hạt không mang điện: %m_n = (30 / 56) × 100% ≈ 53,57% (không phải 39,6%).\n➔ Kết luận: a - Đ, b - Đ, c - S, d - S.")

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
    render_question_prompt(p, "Câu 1: Cho các phát biểu sau:\n(1) Tổng số proton và số neutron trong hạt nhân bằng giá trị số khối.\n(2) Trong một nguyên tử luôn có số proton bằng số neutron.\n(3) Các đồng vị có tính chất vật lí khác nhau nhưng tính chất hoá học giống nhau.\n(4) Số proton bằng số đơn vị điện tích hạt nhân nguyên tử.\n(5) Trong nguyên tử, hạt electron có khối lượng không đáng kể so với các hạt còn lại.\nTrong các phát biểu trên có bao nhiêu phát biểu đúng?")
    if is_teacher:
        add_solution_no_table(doc, "Phân tích từng phát biểu:\n- (1) ĐÚNG: Số khối A = Z + N (tổng số proton và neutron trong hạt nhân).\n- (2) SAI: Trong nguyên tử luôn có số proton bằng số electron; còn số neutron thường lớn hơn hoặc bằng số proton (trừ đồng vị ^1_1H có n = 0 < p = 1).\n- (3) ĐÚNG: Các đồng vị có cùng số proton và cấu hình electron nên tính chất hóa học giống nhau/tương tự nhau, nhưng có số khối khác nhau nên tính chất vật lí khác nhau.\n- (4) ĐÚNG: Số đơn vị điện tích hạt nhân Z bằng đúng số hạt proton.\n- (5) ĐÚNG: Khối lượng hạt electron m_e ≈ 0,00055 amu rất nhỏ so với proton và neutron (~1 amu) nên được coi là không đáng kể.\nCó 4 phát biểu đúng: (1), (3), (4), (5).\n➔ Đáp số: 4")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2: Trong bảng tuần hoàn nguyên tố X thuộc chu kì 4, nhóm IIA. Số hiệu nguyên tử của X là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Nguyên tố X thuộc chu kì 4 (có 4 lớp electron), nhóm IIA (nguyên tố s, có 2 electron hóa trị ở lớp ngoài cùng: 4s^2).\nCấu hình electron đầy đủ: 1s^2 2s^2 2p^6 3s^2 3p^6 4s^2.\nTổng số electron = 2 + 2 + 6 + 2 + 6 + 2 = 20.\nSố hiệu nguyên tử Z = 20 (Calcium - Ca).\n➔ Đáp số: 20")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3: Các electron của nguyên tử nguyên tố X được phân bố trên ba lớp, lớp thứ ba có 5 electron. Số đơn vị điện tích hạt nhân của nguyên tử nguyên tố X là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Nguyên tử X có 3 lớp electron:\n- Lớp 1 (n=1): 2 electron (1s^2)\n- Lớp 2 (n=2): 8 electron (2s^2 2p^6)\n- Lớp 3 (n=3): 5 electron (3s^2 3p^3)\nCấu hình electron: 1s^2 2s^2 2p^6 3s^2 3p^3.\nTổng số electron = 2 + 8 + 5 = 15.\nDo nguyên tử trung hòa về điện nên số proton = 15 => Số đơn vị điện tích hạt nhân Z = 15 (Phosphorus - P).\n➔ Đáp số: 15")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4: Nước cất (H~2~O) là nước tinh khiết, nguyên chất, được điều chế bằng cách chưng cất và thường được sử dụng trong y tế như pha chế thuốc tiêm, thuốc uống, biệt dược, rửa dụng cụ y tế, rửa vết thương,... Biết phân tử này chỉ được tạo nên từ các đồng vị ^1_1H và ^16_8O. Tổng số hạt electron, proton và neutron trong một phân tử H~2~O là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Xét từng đồng vị tạo nên một phân tử H~2~O:\n- 2 nguyên tử đồng vị ^1_1H: mỗi nguyên tử có 1 proton, 1 electron, 0 neutron.\n  => Trong 2H có: 2 proton, 2 electron, 0 neutron.\n- 1 nguyên tử đồng vị ^16_8O: có 8 proton, 8 electron, (16 - 8) = 8 neutron.\nTổng số hạt proton = 2 + 8 = 10 hạt.\nTổng số hạt electron = 2 + 8 = 10 hạt.\nTổng số hạt neutron = 0 + 8 = 8 hạt.\nTổng số hạt (p + e + n) trong một phân tử H~2~O = 10 + 10 + 8 = 28 hạt.\n➔ Đáp số: 28")
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
    render_question_prompt(p, "Câu 1: (1,0 điểm) Nguyên tử của nguyên tố X nằm ở ô thứ 17, chu kì 3 nhóm VIIA trong bảng HTTH các nguyên tố hoá học.\na) Viết cấu hình electron nguyên tử X.\nb) Viết công thức oxide cao nhất đối với oxigene và hợp chất với Hidrogene.")
    if is_teacher:
        sol_tl1 = (
            "Hướng dẫn giải chi tiết Câu 1 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "- X ở ô thứ 17 => Z = 17.\n"
            "- Cấu hình electron của X: 1s^2 2s^2 2p^6 3s^2 3p^5 (hoặc [Ne]3s^2 3p^5). (0,5đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Do X thuộc nhóm VIIA, hóa trị cao nhất của X với oxygen là VII => Công thức oxide cao nhất là X~2~O~7~ (hoặc Cl~2~O~7~). (0,25đ)\n"
            "- Hóa trị của X trong hợp chất với hydrogen là 8 - 7 = I => Công thức hợp chất với hydrogen là HX (hoặc HCl). (0,25đ)"
        )
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 2 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2: (1,0 điểm) Nguyên tử khối trung bình của nguyên tố R là 79,91. R có hai đồng vị. Biết ^79R chiếm 54,5%. Nguyên tử khối của đồng vị thứ 2 bằng bao nhiêu?")
    if is_teacher:
        sol_tl2 = (
            "Hướng dẫn giải chi tiết Câu 2 (1,0 điểm):\n"
            "- Gọi số khối của đồng vị thứ hai là A~2~.\n"
            "- Vì R có hai đồng vị và đồng vị ^79R chiếm 54,5% nên thành phần % của đồng vị thứ hai là:\n"
            "    %A~2~ = 100% - 54,5% = 45,5%. (0,25đ)\n"
            "- Áp dụng công thức tính nguyên tử khối trung bình:\n"
            "    A_tb = [79 × 54,5 + A~2~ × 45,5] / 100 = 79,91 (0,25đ)\n"
            "    <=> 4305,5 + 45,5 × A~2~ = 7991 (0,25đ)\n"
            "    <=> 45,5 × A~2~ = 3685,5  <=>  A~2~ = 81 (amu). (0,25đ)\n"
            "Vậy nguyên tử khối của đồng vị thứ hai là 81 (đây là đồng vị ^81Br)."
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 3 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3: (1,0 điểm) Một loại quặng chứa 2 muối cacbonate của 2 kim loại kế tiếp nhau trong nhóm IIA được dùng trong vật liệu xây dựng, luyện kim, cải tạo đất chua... Cho 10,80 g quặng trên tác dụng với dd H~2~SO~4~ loãng dư. Chất khí thu được cho hấp thụ hoàn toàn vào dd Ba(OH)~2~ dư thì thu được 23,64 g kết tủa.\na. Viết các phương trình phản ứng hoá học xảy ra. Tìm công thức 2 muối trên.\nb. Cần dùng bao nhiêu ml dung dịch HCl 30%, d = 1,5 gam/ml để phản ứng hết 5,4 gam quặng trên.")
    if is_teacher:
        sol_tl3 = (
            "Hướng dẫn giải chi tiết Câu 3 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "- Gọi công thức chung 2 muối cacbonat của 2 kim loại nhóm IIA là M_tbCO~3~.\n"
            "- Các phương trình hóa học xảy ra:\n"
            "    M_tbCO~3~ + H~2~SO~4~ -> M_tbSO~4~ + CO~2~ + H~2~O (1)\n"
            "    CO~2~ + Ba(OH)~2~ dư -> BaCO~3~↓ + H~2~O (2) (0,15đ)\n"
            "- Kết tủa thu được là BaCO~3~ (M = 197 g/mol):\n"
            "    n_BaCO3 = 23,64 / 197 = 0,12 mol.\n"
            "- Theo (1) và (2): n_CO2 = n_MtbCO3 = n_BaCO3 = 0,12 mol. (0,1đ)\n"
            "- Khối lượng mol trung bình của muối cacbonat:\n"
            "    M_tb(muối) = M_tb + 60 = 10,80 / 0,12 = 90 g/mol\n"
            "    => M_tb = 90 - 60 = 30 g/mol. (0,15đ)\n"
            "- Do 2 kim loại thuộc nhóm IIA và ở hai chu kì liên tiếp (kế tiếp nhau):\n"
            "    M_A < 30 < M_B\n"
            "    Các kim loại nhóm IIA gồm: Be (9) < Mg (24) < Ca (40) < Sr (88) < Ba (137).\n"
            "    => Hai kim loại đó là Magnesium (Mg, M=24) và Calcium (Ca, M=40).\n"
            "    Công thức của 2 muối cacbonat là: MgCO~3~ và CaCO~3~. (0,1đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Phương trình phản ứng với dung dịch HCl:\n"
            "    M_tbCO~3~ + 2HCl -> M_tbCl~2~ + CO~2~ + H~2~O (3)\n"
            "- Số mol muối trong 5,4 gam quặng (5,4 gam = 10,80 / 2):\n"
            "    n_muối = 5,4 / 90 = 0,06 mol. (0,15đ)\n"
            "- Theo PTHH (3): n_HCl = 2 × n_muối = 2 × 0,06 = 0,12 mol. (0,1đ)\n"
            "- Khối lượng chất tan HCl: m_HCl = 0,12 × 36,5 = 4,38 gam.\n"
            "- Khối lượng dung dịch HCl 30% cần dùng:\n"
            "    m_dd HCl = (4,38 × 100) / 30 = 14,6 gam. (0,15đ)\n"
            "- Thể tích dung dịch HCl 30% cần dùng:\n"
            "    V_dd HCl = m_dd / d = 14,6 / 1,5 ≈ 9,733 ml ≈ 9,73 ml. (0,1đ)"
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
    
    file_student = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_Hoa_10_Ma132.docx")
    file_teacher = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_Hoa_10_Ma132_GiaoVien.docx")

    print("Đang tạo Bản Học sinh Mã đề 132...")
    build_exam_haibatrung_132(is_teacher=False, output_path=file_student)

    print("Đang tạo Bản Giáo viên Mã đề 132...")
    build_exam_haibatrung_132(is_teacher=True, output_path=file_teacher)

    print("\n[HOÀN TẤT] Cả 2 file Word Mã đề 132 đã được tạo chuẩn 100% FORM ĐỀ KIỂM TRA!")

if __name__ == "__main__":
    main()
