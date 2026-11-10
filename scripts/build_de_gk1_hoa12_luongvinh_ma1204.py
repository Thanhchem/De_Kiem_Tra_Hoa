# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Kỳ I Hóa Học 12 - THPT Lương Vinh - Mã Đề 1204
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản Word:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_12_LuongVinh_Ma1204_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_12_LuongVinh_Ma1204_GiaoVien.docx
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


def add_solution(doc, sol):
    if isinstance(sol, list):
        sol = "\n".join(sol)
    add_solution_no_table(doc, sol)


def add_quick_answer_table_lv1204(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu MÃ ĐỀ 1204.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 1204)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I (12 câu)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (12 câu)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_1 = [2.4] + [1.3] * 12
    p1_answers = ["C", "A", "D", "C", "D", "B", "C", "B", "D", "C", "B", "D"]

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

    # 2. PHẦN II (2 câu Đúng/Sai)
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(6)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*2. Phần II: Trắc nghiệm Đúng / Sai*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_2 = [3.0, 3.0, 3.0, 3.0, 3.0, 3.0]
    p2_answers = [
        ("Câu 1", "Đ", "Đ", "Đ", "Đ"),
        ("Câu 2", "Đ", "S", "Đ", "S"),
    ]

    for row_idx, row in enumerate(tbl2.rows):
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

            if row_idx == 0:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8F5E9"/>')
                tcPr.append(shd)
                headers = ["Câu", "Lệnh hỏi", "Ý a)", "Ý b)", "Ý c)", "Ý d)"]
                render_rich_text(p, f"*{headers[col_idx]}*", base_size=10, base_bold=True)
            else:
                data = p2_answers[row_idx - 1]
                if col_idx == 0:
                    render_rich_text(p, f"*{data[0]}*", base_size=10, base_bold=True)
                elif col_idx == 1:
                    render_rich_text(p, "Đúng/Sai", base_size=9.5, base_italic=True)
                else:
                    ans_val = data[col_idx - 1]
                    ans_col = COLOR_CORRECT if ans_val == "Đ" else COLOR_PRIMARY
                    render_rich_text(p, f"*{ans_val}*", base_size=10.5, base_bold=True, base_color=ans_col)

    # 3. PHẦN III (4 câu trả lời ngắn)
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(6)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*3. Phần III: Trắc nghiệm Trả lời ngắn*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_3 = [3.6, 3.6, 3.6, 3.6, 3.6]
    p3_answers = ["25", "4", "2", "200"]

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
                    render_rich_text(p, f"*{txt}*", base_size=11, base_bold=True, base_color=COLOR_CORRECT)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(6)
    p_div.paragraph_format.space_after = Pt(6)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


def build_exam_document_lv1204(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="TRƯỜNG THPT LƯƠNG VINH",
        exam_title="ĐỀ KIỂM TRA GIỮA HỌC KỲ I",
        subject_line="MÔN: HOÁ HỌC 12",
        num_pages="04",
        duration="45 phút",
        exam_code="1204",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; S = 32; Cu = 64; Ag = 108",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_lv1204(doc)

    # ========================== PHẦN I ==========================
    add_section_header(
        doc,
        "PHẦN I. Câu hỏi trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # CÂU 1
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Glucose và fructose đều")
    opts_1 = [
        "A. có nhóm -CH=O trong phân tử.",
        "B. thuộc loại disaccharide.",
        "*C. có phản ứng tráng bạc.*",
        "D. có công thức phân tử C~6~H~10~O~5~."
    ]
    render_options_no_table(doc, opts_1, is_teacher=is_teacher)
    if is_teacher:
        sol_1 = [
            "*Chọn C.*",
            "Glucose có nhóm aldehyde (-CH=O) nên tham gia phản ứng tráng bạc. Fructose có nhóm ketone, nhưng trong môi trường kiềm (dung dịch AgNO~3~/NH~3~), fructose chuyển hóa một phần thành glucose nên cũng tham gia phản ứng tráng bạc. Cả hai đều có CTPT C~6~H~12~O~6~ và thuộc loại monosaccharide."
        ]
        add_solution(doc, sol_1)

    # CÂU 2
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Hợp chất C~2~H~5~NHC~2~H~5~ có tên là")
    opts_2 = [
        "*A. diethylamine.*",
        "B. dimethylamine.",
        "C. ethylmethylamine.",
        "D. propylamine."
    ]
    render_options_no_table(doc, opts_2, is_teacher=is_teacher)
    if is_teacher:
        sol_2 = [
            "*Chọn A.*",
            "Hợp chất C~2~H~5~NHC~2~H~5~ là amine bậc 2 chứa hai gốc ethyl liên kết với nguyên tử nitrogen, có tên gốc - chức là diethylamine."
        ]
        add_solution(doc, sol_2)

    # CÂU 3
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Tinh bột thuộc loại polysaccharide, có nhiều trong các loại ngũ cốc như hạt lúa, hạt ngô, củ khoai, củ sắn, .... Công thức của tinh bột là")
    opts_3 = [
        "A. C~12~H~22~O~11~.",
        "B. C~2~H~4~O~2~.",
        "C. C~6~H~12~O~6~.",
        "*D. (C~6~H~10~O~5~)~n~.*"
    ]
    render_options_no_table(doc, opts_3, is_teacher=is_teacher)
    if is_teacher:
        sol_3 = [
            "*Chọn D.*",
            "Tinh bột là một polysaccharide có công thức phân tử tổng quát là (C~6~H~10~O~5~)~n~."
        ]
        add_solution(doc, sol_3)

    # CÂU 4
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Ester HCOOCH~3~ phản ứng với dung dịch NaOH (đun nóng), sinh ra các sản phẩm hữu cơ là")
    opts_4 = [
        "A. HCOOH và CH~3~ONa.",
        "B. CH~3~ONa và HCOONa.",
        "*C. HCOONa và CH~3~OH.*",
        "D. CH~3~COOH và CH~3~OH."
    ]
    render_options_no_table(doc, opts_4, is_teacher=is_teacher)
    if is_teacher:
        sol_4 = [
            "*Chọn C.*",
            "Phương trình phản ứng xà phòng hóa methyl formate:",
            "HCOOCH~3~ + NaOH -(t^o^)-> HCOONa (sodium formate) + CH~3~OH (methanol)."
        ]
        add_solution(doc, sol_4)

    # CÂU 5
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 5. Cho các chất có công thức cấu tạo sau: H~2~NCH~2~COOH (1); C~2~H~5~COOH (2); C~2~H~5~NH~2~ (3); H~2~NCH~2~CH~2~CH(NH~2~)COOH (4); C~6~H~5~NH~2~ (5). Những chất vừa phản ứng được với acid vừa phản ứng được với base là")
    opts_5 = [
        "A. (1), (2).",
        "B. (2), (3).",
        "C. (4), (5).",
        "*D. (1), (4).*"
    ]
    render_options_no_table(doc, opts_5, is_teacher=is_teacher)
    if is_teacher:
        sol_5 = [
            "*Chọn D.*",
            "Các hợp chất amino acid có nhóm -NH~2~ (tính base) và nhóm -COOH (tính acid) nên thể hiện tính chất lưỡng tính, vừa tác dụng được với dung dịch acid mạnh vừa tác dụng được với dung dịch base mạnh. Đó là chất (1) glycine và (4) ornithine."
        ]
        add_solution(doc, sol_5)

    # CÂU 6
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 6. Hợp chất nào sau đây được sử dụng làm chất giặt rửa tổng hợp?")
    opts_6 = [
        "A. CH~3~[CH~2~]~12~COONa.",
        "*B. CH~3~[CH~2~]~11~OSO~3~Na.*",
        "C. CH~3~COONa.",
        "D. CH~3~[CH~2~]~12~COOCH~3~."
    ]
    render_options_no_table(doc, opts_6, is_teacher=is_teacher)
    if is_teacher:
        sol_6 = [
            "*Chọn B.*",
            "- CH~3~[CH~2~]~12~COONa là muối của acid béo (thành phần của xà phòng).",
            "- CH~3~[CH~2~]~11~OSO~3~Na (sodium lauryl sulfate) là muối alkyl sulfate mạch dài, được sản xuất từ dầu mỏ và được dùng làm chất giặt rửa tổng hợp."
        ]
        add_solution(doc, sol_6)

    # CÂU 7
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 7. Số nguyên tử hydrogen trong phân tử methyl formate là")
    opts_7 = [
        "A. 2.",
        "B. 6.",
        "*C. 4.*",
        "D. 8."
    ]
    render_options_no_table(doc, opts_7, is_teacher=is_teacher)
    if is_teacher:
        sol_7 = [
            "*Chọn C.*",
            "Methyl formate có công thức phân tử là C~2~H~4~O~2~ (CTCT: HCOOCH~3~), chứa đúng 4 nguyên tử hydrogen."
        ]
        add_solution(doc, sol_7)

    # CÂU 8
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 8. Phần trăm khối lượng nitrogen trong phân tử aniline bằng")
    opts_8 = [
        "A. 15,73%.",
        "*B. 15,05%.*",
        "C. 18,67%.",
        "D. 12,96%."
    ]
    render_options_no_table(doc, opts_8, is_teacher=is_teacher)
    if is_teacher:
        sol_8 = [
            "*Chọn B.*",
            "Aniline có công thức phân tử là C~6~H~7~N. Phân tử khối: M = 12 x 6 + 7 + 14 = 93 g/mol.",
            "%N = (14 / 93) x 100% = 15,05%."
        ]
        add_solution(doc, sol_8)

    # CÂU 9
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 9. Dung dịch chất nào sau đây hòa tan Cu(OH)~2~, thu được dung dịch có màu xanh lam?")
    opts_9 = [
        "A. Propyl alcohol.",
        "B. Propan-1,3-diol.",
        "C. Albumin.",
        "*D. Fructose.*"
    ]
    render_options_no_table(doc, opts_9, is_teacher=is_teacher)
    if is_teacher:
        sol_9 = [
            "*Chọn D.*",
            "Fructose là một monosaccharide chứa nhiều nhóm -OH kề nhau, ở nhiệt độ thường hòa tan Cu(OH)~2~ tạo dung dịch phức màu xanh lam đặc trưng. (Albumin tạo phức màu tím - biuret; Propan-1,3-diol có các nhóm -OH cách nhau nên không phản ứng hòa tan Cu(OH)~2~)."
        ]
        add_solution(doc, sol_9)

    # CÂU 10
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 10. Một phân tử saccharose có")
    opts_10 = [
        "A. một đơn vị \\beta-glucose và một đơn vị \\alpha-fructose.",
        "B. hai đơn vị \\alpha-glucose.",
        "*C. một đơn vị \\alpha-glucose và một đơn vị \\beta-fructose.*",
        "D. một đơn vị \\beta-glucose và một đơn vị \\beta-fructose."
    ]
    render_options_no_table(doc, opts_10, is_teacher=is_teacher)
    if is_teacher:
        sol_10 = [
            "*Chọn C.*",
            "Phân tử saccharose được tạo nên từ một gốc \\alpha-glucose và một gốc \\beta-fructose liên kết với nhau qua liên kết \\alpha-1,2-glycoside."
        ]
        add_solution(doc, sol_10)

    # CÂU 11
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 11. Phát biểu nào sau đây đúng?")
    opts_11 = [
        "A. Hai đơn vị glucose trong maltose liên kết với nhau bằng liên kết \\alpha-1,2-glycoside.",
        "*B. Saccharose và maltose có cùng công thức phân tử.*",
        "C. Một phân tử saccharose gồm hai đơn vị \\alpha-glucose.",
        "D. Maltose không có nhóm -OH hemiacetal."
    ]
    render_options_no_table(doc, opts_11, is_teacher=is_teacher)
    if is_teacher:
        sol_11 = [
            "*Chọn B.*",
            "- A sai vì maltose gồm 2 gốc \\alpha-glucose liên kết bằng liên kết \\alpha-1,4-glycosidic.",
            "- B đúng vì cả saccharose và maltose đều là disaccharide có cùng CTPT C~12~H~22~O~11~.",
            "- C sai vì saccharose gồm 1 gốc \\alpha-glucose và 1 gốc \\beta-fructose.",
            "- D sai vì maltose còn 1 nhóm -OH hemiacetal tự do nên có thể mở vòng và có tính khử (tráng bạc)."
        ]
        add_solution(doc, sol_11)

    # CÂU 12
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 12. Phát biểu nào sau đây là sai?")
    opts_12 = [
        "A. Dung dịch lysine làm xanh quỳ tím.",
        "B. Aniline tác dụng với nước bromine tạo thành kết tủa trắng.",
        "C. Dung dịch glycine không làm đổi màu quỳ tím.",
        "*D. Cho Cu(OH)~2~ vào dung dịch lòng trắng trứng thấy xuất hiện màu vàng.*"
    ]
    render_options_no_table(doc, opts_12, is_teacher=is_teacher)
    if is_teacher:
        sol_12 = [
            "*Chọn D.*",
            "Lòng trắng trứng (albumin) chứa các liên kết peptide, khi phản ứng với Cu(OH)~2~ trong môi trường kiềm (phản ứng biuret) cho dung dịch phức chất có *màu tím đặc trưng*, không phải màu vàng."
        ]
        add_solution(doc, sol_12)

    # -------------------------- PHẦN II --------------------------
    add_section_header(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # CÂU 1 PHẦN II
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Cho các carbohydrate: glucose, fructose, saccharose, maltose.")
    stmts_p2_1 = [
        ("a) Cả 4 chất ở điều kiện thường đều là chất rắn, vị ngọt, tan tốt trong nước.", "Đúng"),
        ("b) Có 2 chất có nhóm -OH hemiacetal.", "Đúng"),
        ("c) Có 2 cặp chất là đồng phân của nhau.", "Đúng"),
        ("d) Có 3 chất vừa có cấu tạo dạng mạch hở, vừa có cấu tạo dạng mạch vòng.", "Đúng")
    ]
    for lbl, ans_val in stmts_p2_1:
        p_stmt = doc.add_paragraph()
        render_statement(p_stmt, lbl, ans=ans_val, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_1 = [
            "*Lời giải chi tiết:*",
            "- *Ý a) ĐÚNG.* Cả 4 chất monosaccharide và disaccharide này ở điều kiện thường đều là chất rắn kết tinh màu trắng, có vị ngọt và tan tốt trong nước.",
            "- *Ý b) ĐÚNG.* Glucose có nhóm -OH hemiacetal ở C1; maltose có 1 nhóm -OH hemiacetal ở gốc glucose thứ hai. (Fructose chứa nhóm -OH hemiketal, còn saccharose không còn nhóm hemiacetal lẫn hemiketal tự do). Do đó có đúng 2 chất chứa nhóm -OH hemiacetal.",
            "- *Ý c) ĐÚNG.* Cặp 1: glucose và fructose (cùng C~6~H~12~O~6~); Cặp 2: saccharose và maltose (cùng C~12~H~22~O~11~).",
            "- *Ý d) ĐÚNG.* Glucose, fructose và maltose đều có khả năng chuyển hóa cân bằng giữa dạng mạch vòng và mạch hở (do có nhóm -OH hemiacetal/hemiketal tự do có thể mở vòng). Riêng saccharose liên kết glycosidic khóa cả 2 vị trí nên chỉ tồn tại ở dạng mạch vòng. Vậy có đúng 3 chất vừa có cấu tạo mạch hở vừa có mạch vòng."
        ]
        add_solution(doc, sol_p2_1)

    # CÂU 2 PHẦN II
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Cho công thức cấu tạo của peptide Y như sau:\nHOOC-CH~2~-CH~2~-CH(NH~2~)-CO-NH-CH~2~-CO-NH-CH(CH~3~)-COOH")
    stmts_p2_2 = [
        ("a) Tên gọi của Y là Glu-Gly-Ala.", "Đúng"),
        ("b) 1 mol Y phản ứng tối đa được với 3 mol NaOH.", "Sai"),
        ("c) Thủy phân hoàn toàn Y thu được glutamic acid, glycine và alanine.", "Đúng"),
        ("d) Phần trăm khối lượng của nitrogen (N) trong Y là 15,38%.", "Sai")
    ]
    for lbl, ans_val in stmts_p2_2:
        p_stmt = doc.add_paragraph()
        render_statement(p_stmt, lbl, ans=ans_val, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_2 = [
            "*Lời giải chi tiết:*",
            "- *Ý a) ĐÚNG.* Gốc đầu N là glutamic acid (Glu), giữa là glycine (Gly), đầu C là alanine (Ala) => tên gọi viết tắt là Glu-Gly-Ala.",
            "- *Ý b) SAI.* Y có 2 liên kết peptide và 2 nhóm -COOH tự do (1 nhóm ở đầu C của Ala và 1 nhóm ở mạch nhánh của Glu). Khi phản ứng với NaOH nóng: cần 2 mol NaOH để thủy phân 2 liên kết peptide và 2 mol NaOH để trung hòa 2 nhóm -COOH tự do => 1 mol Y phản ứng tối đa với *4 mol NaOH*.",
            "- *Ý c) ĐÚNG.* Y được cấu thành từ Glu, Gly, Ala nên khi thủy phân hoàn toàn trong acid hoặc enzyme sẽ thu được 3 amino acid tương ứng.",
            "- *Ý d) SAI.*",
            "  Khối lượng mol của Y: M_Y = M_Glu + M_Gly + M_Ala - 2 x 18 = 147 + 75 + 89 - 36 = 275 g/mol.",
            "  Y chứa 3 nguyên tử N: %N = (3 x 14 / 275) x 100% = 15,27% (khác 15,38%)."
        ]
        add_solution(doc, sol_p2_2)

    # -------------------------- PHẦN III --------------------------
    add_section_header(
        doc,
        "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 4."
    )

    # CÂU 1 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Tiến hành thí nghiệm sau:\n- Bước 1: Cho vào ống nghiệm 2 mL nước cất.\n- Bước 2: Nhỏ tiếp vài giọt aniline vào ống nghiệm, sau đó nhúng giấy quỳ tím vào dung dịch trong ống nghiệm.\n- Bước 3: Nhỏ tiếp 1 mL dung dịch HCl đặc vào ống nghiệm.\nCho các phát biểu sau:\n(1) Sau bước 2, chất lỏng bị phân thành 2 lớp, lớp dưới là aniline.\n(2) Sau bước 2, giấy quỳ tím chuyển thành màu xanh.\n(3) Sau bước 3, dung dịch thu được trong suốt.\n(4) Sau bước 3, trong dung dịch có chứa muối phenylammonium chloride tan tốt trong nước.\n(5) Ở bước 3, nếu thay HCl bằng Br~2~ thì sẽ thấy xuất hiện kết tủa màu vàng.\nLiệt kê các phát biểu sai theo thứ tự tăng dần (ví dụ 12, 13, ...).")
    if is_teacher:
        sol_p3_1 = [
            "*Trả lời:* *25*",
            "*Lời giải:*",
            "- (1) ĐÚNG: Aniline rất ít tan trong nước và nặng hơn nước (d = 1,02 g/mL) nên lắng xuống dưới đáy tạo 2 lớp phân cách.",
            "- (2) SAI: Aniline có tính base rất yếu, không làm đổi màu quỳ tím.",
            "- (3) ĐÚNG: Aniline phản ứng với HCl tạo muối C~6~H~5~NH~3~Cl tan tốt trong nước làm dung dịch đồng nhất, trong suốt.",
            "- (4) ĐÚNG: Muối sinh ra là phenylammonium chloride: C~6~H~5~NH~2~ + HCl -> C~6~H~5~NH~3~Cl.",
            "- (5) SAI: Aniline phản ứng với nước bromine sinh ra kết tủa trắng (2,4,6-tribromoaniline), không phải màu vàng.",
            "=> Các phát biểu sai là (2) và (5) -> Ghép thứ tự tăng dần: *25*."
        ]
        add_solution(doc, sol_p3_1)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 2 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Cho các chất: methylamine, glycine, aniline, acetic acid, lysine, valine. Có bao nhiêu chất phản ứng với dung dịch NaOH tạo muối?")
    if is_teacher:
        sol_p3_2 = [
            "*Trả lời:* *4*",
            "*Lời giải:* Các chất tác dụng với dung dịch NaOH tạo muối gồm:",
            "1. Glycine (H~2~NCH~2~COOH) -> H~2~NCH~2~COONa.",
            "2. Acetic acid (CH~3~COOH) -> CH~3~COONa.",
            "3. Lysine (H~2~N-[CH~2~]~4~-CH(NH~2~)-COOH) -> H~2~N-[CH~2~]~4~-CH(NH~2~)-COONa.",
            "4. Valine ((CH~3~)~2~CH-CH(NH~2~)-COOH) -> (CH~3~)~2~CH-CH(NH~2~)-COONa.",
            "(Methylamine và aniline là amine có tính base nên không phản ứng với dung dịch base NaOH).",
            "=> Tổng cộng có 4 chất."
        ]
        add_solution(doc, sol_p3_2)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 3 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Cho các chất: CH~3~[CH~2~]~14~COONa, CH~3~[CH~2~]~10~CH~2~OSO~3~Na, C~2~H~5~OSO~3~Na, CH~3~[CH~2~]~16~COOK, CH~3~[CH~2~]~11~C~6~H~4~SO~3~Na. Có bao nhiêu chất là chất giặt rửa tổng hợp?")
    if is_teacher:
        sol_p3_3 = [
            "*Trả lời:* *2*",
            "*Lời giải:*",
            "- CH~3~[CH~2~]~14~COONa và CH~3~[CH~2~]~16~COOK là muối sodium/potassium của acid béo -> xà phòng.",
            "- C~2~H~5~OSO~3~Na có gốc alkyl quá ngắn (chỉ 2C), không có đầu kị nước đủ dài nên không có hoạt tính giặt rửa.",
            "- Hai chất giặt rửa tổng hợp gồm:",
            "  1. CH~3~[CH~2~]~10~CH~2~OSO~3~Na (sodium lauryl sulfate - muối alkyl sulfate).",
            "  2. CH~3~[CH~2~]~11~C~6~H~4~SO~3~Na (sodium linear alkylbenzene sulfonate - LAS).",
            "=> Có 2 chất là chất giặt rửa tổng hợp."
        ]
        add_solution(doc, sol_p3_3)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 4 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Từ m kg mùn cưa chứa 60% cellulose (còn lại là tạp chất trơ) sản xuất được 100 kg glucose với hiệu suất toàn bộ quá trình là 75%. Giá trị của m là bao nhiêu?")
    if is_teacher:
        sol_p3_4 = [
            "*Trả lời:* *200*",
            "*Lời giải:*",
            "- Phương trình thủy phân: (C~6~H~10~O~5~)~n~ + n H~2~O -(t^o^, H^+^)-> n C~6~H~12~O~6~.",
            "  Cứ 162 kg cellulose theo lí thuyết thu được 180 kg glucose.",
            "- Khối lượng cellulose lí thuyết cần dùng:",
            "  m_cellulose_lt = 100 x (162 / 180) = 90 kg.",
            "- Khối lượng cellulose thực tế (với hiệu suất H = 75%):",
            "  m_cellulose_tt = 90 / 75% = 120 kg.",
            "- Khối lượng mùn cưa cần dùng (chứa 60% cellulose):",
            "  m = 120 / 60% = *200* kg."
        ]
        add_solution(doc, sol_p3_4)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # ========================== PHẦN IV. TỰ LUẬN ==========================
    p_pb = doc.add_paragraph()
    p_pb.paragraph_format.space_before = Pt(8)
    p_pb.paragraph_format.space_after = Pt(4)
    render_rich_text(p_pb, "*PHẦN IV. TỰ LUẬN*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # CÂU 1 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Hoàn thành dãy chuyển hóa sau (ghi rõ điều kiện nếu có):\nTinh bột -> Glucose -> Ethanol -> Acetic acid -> Methyl acetate")
    if is_teacher:
        sol_tl_1 = [
            "*Lời giải:*",
            "Các phương trình hóa học theo sơ đồ chuyển hóa:",
            "1. (C~6~H~10~O~5~)~n~ + n H~2~O -(acid hoặc enzyme, t^o^)-> n C~6~H~12~O~6~ (glucose)",
            "2. C~6~H~12~O~6~ -(men rượu, 30-35 ^o^C)-> 2C~2~H~5~OH + 2CO~2~^",
            "3. C~2~H~5~OH + O~2~ -(men giấm, 25-30 ^o^C)-> CH~3~COOH + H~2~O",
            "4. CH~3~COOH + CH~3~OH <=(H~2~SO~4~ đặc, t^o^)=> CH~3~COOCH~3~ (methyl acetate) + H~2~O"
        ]
        add_solution(doc, sol_tl_1)
    else:
        add_student_writing_lines(doc, num_lines=5)

    # CÂU 2 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Phản ứng tổng hợp glucose trong cây xanh cần được cung cấp năng lượng từ ánh sáng mặt trời:\n6CO~2~ + 6H~2~O + 673 kcal -(as, diệp lục)-> C~6~H~12~O~6~ + 6O~2~\nCứ trong một phút, mỗi cm^2^ lá xanh nhận được 0,5 cal năng lượng mặt trời, nhưng chỉ có 10% được sử dụng vào phản ứng tổng hợp glucose. Thời gian để một cây có 1000 lá xanh (diện tích mỗi lá 10 cm^2^) sản sinh được 18 gam glucose là bao lâu?")
    if is_teacher:
        sol_tl_2 = [
            "*Lời giải chi tiết:*",
            "- Số mol glucose cần sản sinh: n_glucose = 18 / 180 = 0,1 mol.",
            "- Theo phương trình: 1 mol glucose cần 673 kcal = 673 000 cal.",
            "  Năng lượng cần để tổng hợp 0,1 mol glucose:",
            "  E_cần = 0,1 x 673 000 cal = 67 300 cal.",
            "- Tổng diện tích bề mặt của 1000 lá xanh:",
            "  S = 1000 lá x 10 cm^2^/lá = 10 000 cm^2^.",
            "- Năng lượng mặt trời mà toàn bộ lá cây nhận được trong 1 phút:",
            "  E_nhận_1_phút = 10 000 cm^2^ x 0,5 cal/cm^2^ = 5 000 cal/phút.",
            "- Năng lượng hữu ích thực tế dùng cho quang hợp trong 1 phút (hiệu suất 10%):",
            "  E_hữu_ích = 5 000 x 10% = 500 cal/phút.",
            "- Thời gian cần thiết để cây quang hợp tạo ra 18 gam glucose:",
            "  t = E_cần / E_hữu_ích = 67 300 / 500 = *134,6 phút* (tương đương 2 giờ 14 phút 36 giây)."
        ]
        add_solution(doc, sol_tl_2)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # CÂU 3 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Cho X là hexapeptide Ala-Gly-Ala-Val-Gly-Val và Y là tetrapeptide Gly-Ala-Gly-Glu. Thủy phân hoàn toàn m gam hỗn hợp gồm X và Y thu được 4 amino acid, trong đó có 30 gam glycine và 28,48 gam alanine. Giá trị của m là bao nhiêu?")
    if is_teacher:
        sol_tl_3 = [
            "*Lời giải chi tiết:*",
            "- Số mol các amino acid thu được:",
            "  n_Gly = 30 / 75 = 0,4 mol;   n_Ala = 28,48 / 89 = 0,32 mol.",
            "- Phân tích công thức của các peptide:",
            "  + X: Ala-Gly-Ala-Val-Gly-Val có cấu tạo Ala~2~Gly~2~Val~2~ (hexapeptide).",
            "    Khối lượng mol của X: M_X = 2 x 89 + 2 x 75 + 2 x 117 - 5 x 18 = 472 g/mol.",
            "  + Y: Gly-Ala-Gly-Glu có cấu tạo Gly~2~AlaGlu (tetrapeptide).",
            "    Khối lượng mol của Y: M_Y = 2 x 75 + 89 + 147 - 3 x 18 = 332 g/mol.",
            "- Gọi số mol của X là x mol, số mol của Y là y mol trong m gam hỗn hợp:",
            "  + Số mol Gly: 2x + 2y = 0,4  => x + y = 0,2    (1)",
            "  + Số mol Ala: 2x + 1y = 0,32                   (2)",
            "- Giải hệ phương trình (1) và (2):",
            "  x = 0,32 - 0,2 = 0,12 mol;",
            "  y = 0,2 - 0,12 = 0,08 mol.",
            "- Khối lượng m của hỗn hợp:",
            "  m = 0,12 x 472 + 0,08 x 332 = 56,64 + 26,56 = *83,2 gam*."
        ]
        add_solution(doc, sol_tl_3)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(0)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "--- *HẾT* ---", base_size=11, base_bold=True)

    return doc


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 1. Bản Học sinh
    doc_hs = build_exam_document_lv1204(is_teacher=False)
    file_hs = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_12_LuongVinh_Ma1204_HocSinh.docx")
    doc_hs.save(file_hs)
    print(f"[OK] Đã tạo thành công Bản Học sinh: {file_hs}")

    file_hs_std = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_12_LuongVinh_Ma1204.docx")
    doc_hs.save(file_hs_std)
    print(f"[OK] Đã tạo thành công Bản Đề Chuẩn: {file_hs_std}")

    # 2. Bản Giáo viên
    doc_gv = build_exam_document_lv1204(is_teacher=True)
    file_gv = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_12_LuongVinh_Ma1204_GiaoVien.docx")
    doc_gv.save(file_gv)
    print(f"[OK] Đã tạo thành công Bản Giáo viên: {file_gv}")


if __name__ == "__main__":
    main()
