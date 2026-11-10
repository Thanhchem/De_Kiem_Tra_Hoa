# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Kỳ I Hóa Học 12 - THPT Nguyễn Trường Tộ - Mã Đề 103
Năm học 2026 - 2027
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản Word:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma103_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma103_GiaoVien.docx
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


def add_quick_answer_table_ntt103(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu MÃ ĐỀ 103.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 103)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I (16 câu)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (16 câu)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=17)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_1 = [2.0] + [1.0] * 16
    p1_answers = ["B", "B", "C", "C", "D", "C", "A", "C", "D", "A", "D", "C", "C", "B", "C", "C"]

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
                render_rich_text(p, f"*{txt}*", base_size=9.5, base_bold=True)
            else:
                txt = "Đáp án" if col_idx == 0 else p1_answers[col_idx - 1]
                if col_idx == 0:
                    render_rich_text(p, f"*{txt}*", base_size=9.5, base_bold=True)
                else:
                    render_rich_text(p, f"*{txt}*", base_size=10, base_bold=True, base_color=COLOR_CORRECT)

    # 2. PHẦN II (2 câu Đúng/Sai)
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(6)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*2. Phần II: Trắc nghiệm Đúng / Sai*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_2 = [3.0, 3.0, 3.0, 3.0, 3.0, 3.0]
    p2_answers = [
        ("Câu 1", "Đ", "S", "Đ", "S"),
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
    p3_answers = ["5", "4", "4402", "8"]

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


def build_exam_document_ntt103(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="TRƯỜNG THPT NGUYỄN TRƯỜNG TỘ\nTỔ HÓA - TIN",
        exam_title="ĐỀ KIỂM TRA GIỮA KỲ I - NĂM HỌC 2026 - 2027",
        subject_line="MÔN: HÓA HỌC LỚP 12",
        num_pages="04",
        duration="45 phút",
        exam_code="103",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Br = 80; Cu = 64; Ag = 108",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_ntt103(doc)

    # ========================== PHẦN A. TRẮC NGHIỆM ==========================
    p_pa = doc.add_paragraph()
    p_pa.paragraph_format.space_before = Pt(6)
    p_pa.paragraph_format.space_after = Pt(2)
    render_rich_text(p_pa, "*TRẮC NGHIỆM KHÁCH QUAN (7,0 ĐIỂM)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # -------------------------- PHẦN I --------------------------
    add_section_header(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # CÂU 1
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Dipeptide có công thức: H~2~NCH~2~CONHCH(CH~3~)COOH. Tên viết tắt của X là")
    opts_1 = [
        "A. Ala-Val.",
        "*B. Gly-Ala.*",
        "C. Ala-Gly.",
        "D. Gly-Val."
    ]
    render_options_no_table(doc, opts_1, is_teacher=is_teacher)
    if is_teacher:
        sol_1 = [
            "*Chọn B.*",
            "Đầu N của dipeptide chứa gốc glycine (H~2~NCH~2~CO-), đầu C chứa gốc alanine (-NHCH(CH~3~)COOH). Do đó tên viết tắt của dipeptide là Gly-Ala."
        ]
        add_solution(doc, sol_1)

    # CÂU 2
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Chất nào sau đây là thành phần chủ yếu của xà phòng?")
    opts_2 = [
        "A. CH~3~COONa.",
        "*B. C~15~H~31~COOK.*",
        "C. CH~2~=CHCOONa.",
        "D. CH~3~(CH~2~)~3~COOK."
    ]
    render_options_no_table(doc, opts_2, is_teacher=is_teacher)
    if is_teacher:
        sol_2 = [
            "*Chọn B.*",
            "Xà phòng là hỗn hợp muối sodium hoặc potassium của các acid béo mạch dài (thường có từ 12 đến 18 nguyên tử carbon). C~15~H~31~COOK là potassium palmitate, thành phần chính của xà phòng thơm/xà phòng mềm."
        ]
        add_solution(doc, sol_2)

    # CÂU 3
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Công thức cấu tạo của ethylamine là")
    opts_3 = [
        "A. CH~3~NH~2~.",
        "B. CH~3~NHCH~3~.",
        "*C. C~2~H~5~NH~2~.*",
        "D. (CH~3~)~3~N."
    ]
    render_options_no_table(doc, opts_3, is_teacher=is_teacher)
    if is_teacher:
        sol_3 = [
            "*Chọn C.*",
            "Ethylamine gồm nhóm ethyl (-C~2~H~5~) liên kết với nhóm amino (-NH~2~), có công thức là C~2~H~5~NH~2~."
        ]
        add_solution(doc, sol_3)

    # CÂU 4
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Cơ thể người mã hoá được loại amino acid nào sau đây trong tổng hợp protein cho cơ thể?")
    opts_4 = [
        "A. H~2~N-CH~2~-CH~2~-OH.",
        "B. ^+^NH~3~-CH~2~-CH~2~-COO^-^.",
        "*C. ^+^NH~3~-CH~2~-COO^-^.*",
        "D. ^+^NH~3~-(CH~2~)~3~-COO^-^."
    ]
    render_options_no_table(doc, opts_4, is_teacher=is_teacher)
    if is_teacher:
        sol_4 = [
            "*Chọn C.*",
            "Các amino acid tham gia tổng hợp protein trong cơ thể sinh vật đều là các \\alpha-amino acid. Trong các chất trên:",
            "- A không phải amino acid (là 2-aminoethanol).",
            "- B là \\beta-alanine (\\beta-amino acid).",
            "- C là ion lưỡng cực của glycine (\\alpha-amino acid duy nhất trong các phương án).",
            "- D là GABA (\\gamma-aminobutyric acid)."
        ]
        add_solution(doc, sol_4)

    # CÂU 5
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 5. Carbohydrate nào dưới đây không có nhóm -OH hemiacetal hoặc nhóm -OH hemiketal?")
    opts_5 = [
        "A. Maltose.",
        "B. Fructose.",
        "C. Glucose.",
        "*D. Saccharose.*"
    ]
    render_options_no_table(doc, opts_5, is_teacher=is_teacher)
    if is_teacher:
        sol_5 = [
            "*Chọn D.*",
            "Trong phân tử saccharose, nhóm -OH hemiacetal của C1(\\alpha-glucose) và nhóm -OH hemiketal của C2(\\beta-fructose) đều đã tham gia liên kết glycoside (C1-O-C2). Do đó saccharose không còn nhóm -OH hemiacetal hoặc nhóm -OH hemiketal tự do."
        ]
        add_solution(doc, sol_5)

    # CÂU 6
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 6. Giá trị pH mà khi đó amino acid có nồng độ ion lưỡng cực là cực đại được gọi là điểm đẳng điện (kí hiệu là pI). Khi pH < pI thì amino acid đó tồn tại chủ yếu ở dạng cation, còn khi pH > pI thì amino acid đó tồn tại chủ yếu ở dạng anion. Khi đặt trong một điện trường dạng anion sẽ di chuyển về cực (+) còn dạng cation sẽ di chuyển về cực (-). Tính chất này được gọi là tính điện di và được dùng để tách, tinh chế amino acid ra khỏi hỗn hợp của chúng. Cho các giá trị pI của các chất sau:\n- Glycine (H~2~NCH~2~COOH): pI = 6,0\n- Glutamic acid (HOOCCH~2~CH~2~CH(NH~2~)COOH): pI = 3,2\n- Lysine (H~2~N[CH~2~]~4~CH(NH~2~)COOH): pI = 9,7\nTrong các giá trị pH cho dưới đây, giá trị nào là tối ưu nhất để tách ba chất trên ra khỏi dung dịch hỗn hợp của chúng?")
    opts_6 = [
        "A. pH = 3,2.",
        "B. pH = 9,7.",
        "*C. pH = 6,0.*",
        "D. pH = 14,0."
    ]
    render_options_no_table(doc, opts_6, is_teacher=is_teacher)
    if is_teacher:
        sol_6 = [
            "*Chọn C.*",
            "Ở pH = 6,0:",
            "- Glycine có pH = pI = 6,0 nên tồn tại ở dạng ion lưỡng cực trung hòa điện tích (điện tích tổng bằng 0), đứng yên không di chuyển.",
            "- Glutamic acid có pI = 3,2 < pH (6,0) nên tồn tại ở dạng anion, di chuyển về cực dương (+).",
            "- Lysine có pI = 9,7 > pH (6,0) nên tồn tại ở dạng cation, di chuyển về cực âm (-).",
            "Do ba chất di chuyển theo 3 hướng hoàn toàn khác nhau nên pH = 6,0 là điều kiện tối ưu nhất để phân tách."
        ]
        add_solution(doc, sol_6)

    # CÂU 7
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 7. Polymer thiên nhiên X được sinh ra trong quá trình quang hợp của cây xanh. Ở nhiệt độ thường, X tạo với dung dịch iodine hợp chất có màu xanh tím. Polymer X là")
    opts_7 = [
        "*A. tinh bột.*",
        "B. cellulose.",
        "C. saccharose.",
        "D. glycogen."
    ]
    render_options_no_table(doc, opts_7, is_teacher=is_teacher)
    if is_teacher:
        sol_7 = [
            "*Chọn A.*",
            "Tinh bột là sản phẩm quang hợp của cây xanh và có phản ứng màu đặc trưng với dung dịch iodine tạo hợp chất màu xanh tím do cấu trúc mạch xoắn giữ các phân tử I~2~."
        ]
        add_solution(doc, sol_7)

    # CÂU 8
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 8. Cho các chất: glycerol, glucose, maltose, acetic acid, fructose và cellulose. Có bao nhiêu chất thuộc loại carbohydrate?")
    opts_8 = [
        "A. 2.",
        "B. 3.",
        "*C. 4.*",
        "D. 5."
    ]
    render_options_no_table(doc, opts_8, is_teacher=is_teacher)
    if is_teacher:
        sol_8 = [
            "*Chọn C.*",
            "Các chất thuộc loại carbohydrate gồm 4 chất: glucose (monosaccharide), maltose (disaccharide), fructose (monosaccharide), cellulose (polysaccharide).",
            "(Glycerol là polyalcohol, acetic acid là carboxylic acid)."
        ]
        add_solution(doc, sol_8)

    # CÂU 9
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 9. Một ester có công thức phân tử là C~3~H~6~O~2~, có tham gia phản ứng tráng bạc. Công thức cấu tạo thu gọn của ester đó là")
    opts_9 = [
        "A. HCOOC~3~H~7~.",
        "B. CH~3~COOCH~3~.",
        "C. C~2~H~5~COOCH~3~.",
        "*D. HCOOC~2~H~5~.*"
    ]
    render_options_no_table(doc, opts_9, is_teacher=is_teacher)
    if is_teacher:
        sol_9 = [
            "*Chọn D.*",
            "Ester tham gia phản ứng tráng bạc phải có gốc formate (HCOOR'). Với công thức C~3~H~6~O~2~, ester đó là ethyl formate: HCOOC~2~H~5~."
        ]
        add_solution(doc, sol_9)

    # CÂU 10
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 10. Để chứng minh trong phân tử của glucose có nhiều nhóm hydroxyl kề nhau, người ta cho dung dịch glucose phản ứng với")
    opts_10 = [
        "*A. Cu(OH)~2~ ở nhiệt độ thường.*",
        "B. kim loại Na.",
        "C. AgNO~3~ trong dung dịch NH~3~, đun nóng.",
        "D. Cu(OH)~2~ trong NaOH, đun nóng."
    ]
    render_options_no_table(doc, opts_10, is_teacher=is_teacher)
    if is_teacher:
        sol_10 = [
            "*Chọn A.*",
            "Ở nhiệt độ thường, các chất có từ 2 nhóm -OH kề nhau trở lên hòa tan Cu(OH)~2~ tạo dung dịch phức màu xanh lam đặc trưng."
        ]
        add_solution(doc, sol_10)

    # CÂU 11
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 11. Polysaccharide X là chất rắn, màu trắng, dạng sợi. Trong bông nõn có gần 98% chất X. Thủy phân X, thu được monosaccharide Y. Phát biểu nào sau đây là đúng?")
    opts_11 = [
        "A. X có tính chất của aldehyde.",
        "B. X dễ tan trong nước và có vị ngọt.",
        "C. Y là nguồn lương thực chính của con người và một số động vật.",
        "*D. X được sử dụng làm vật liệu xây dựng, sản xuất giấy và sợi nhân tạo.*"
    ]
    render_options_no_table(doc, opts_11, is_teacher=is_teacher)
    if is_teacher:
        sol_11 = [
            "*Chọn D.*",
            "X là cellulose (chiếm 98% bông nõn), Y là glucose. Cellulose được dùng làm vật liệu xây dựng (gỗ), sản xuất giấy, tơ sợi nhân tạo (viscose, cellulose acetate).",
            "- A sai vì cellulose không có nhóm aldehyde.",
            "- B sai vì cellulose không tan trong nước và không ngọt.",
            "- C sai vì nguồn lương thực chính là tinh bột, không phải glucose."
        ]
        add_solution(doc, sol_11)

    # CÂU 12
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 12. Glucose là một loại monosaccharide có nhiều trong quả nho chín. Công thức phân tử của glucose là")
    opts_12 = [
        "A. C~12~H~22~O~11~.",
        "B. C~2~H~4~O~2~.",
        "*C. C~6~H~12~O~6~.*",
        "D. (C~6~H~10~O~5~)~n~."
    ]
    render_options_no_table(doc, opts_12, is_teacher=is_teacher)
    if is_teacher:
        sol_12 = [
            "*Chọn C.*",
            "Glucose có công thức phân tử là C~6~H~12~O~6~."
        ]
        add_solution(doc, sol_12)

    # CÂU 13
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 13. Methyl acetate có công thức cấu tạo là")
    opts_13 = [
        "A. HCOOCH~3~.",
        "B. HCOOC~2~H~5~.",
        "*C. CH~3~COOCH~3~.*",
        "D. CH~3~COOC~2~H~5~."
    ]
    render_options_no_table(doc, opts_13, is_teacher=is_teacher)
    if is_teacher:
        sol_13 = [
            "*Chọn C.*",
            "Methyl acetate gồm gốc acid acetic (CH~3~COO-) và gốc methyl (-CH~3~): CH~3~COOCH~3~."
        ]
        add_solution(doc, sol_13)

    # CÂU 14
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 14. Mùi tanh của cá là hỗn hợp các amine và một số tạp chất khác. Để khử mùi tanh của cá trước khi chế biến thực phẩm, nên áp dụng cách nào sau đây?")
    opts_14 = [
        "A. Ngâm cá trong nước để amine tan vào nước.",
        "*B. Rửa cá bằng giấm ăn.*",
        "C. Rửa cá bằng dung dịch nước muối.",
        "D. Rửa cá bằng dung dịch soda."
    ]
    render_options_no_table(doc, opts_14, is_teacher=is_teacher)
    if is_teacher:
        sol_14 = [
            "*Chọn B.*",
            "Giấm ăn chứa acetic acid (CH~3~COOH) phản ứng trung hòa các amine có tính base yếu tạo thành muối tan không mùi và dễ rửa trôi bằng nước."
        ]
        add_solution(doc, sol_14)

    # CÂU 15
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 15. Cho các chất sau: CH~3~COOCH~3~, HCOOCH~3~, HCOOC~6~H~5~, CH~3~COOC~2~H~5~. Chất có nhiệt độ sôi thấp nhất là")
    opts_15 = [
        "A. HCOOC~6~H~5~.",
        "B. CH~3~COOC~2~H~5~.",
        "*C. HCOOCH~3~.*",
        "D. CH~3~COOCH~3~."
    ]
    render_options_no_table(doc, opts_15, is_teacher=is_teacher)
    if is_teacher:
        sol_15 = [
            "*Chọn C.*",
            "Các ester no đơn chức không tạo liên kết hydrogen liên phân tử. Chất có phân tử khối nhỏ nhất sẽ có nhiệt độ sôi thấp nhất: HCOOCH~3~ (M = 60, t~s~ ~ 31,5 ^o^C)."
        ]
        add_solution(doc, sol_15)

    # CÂU 16
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 16. Khi nấu canh cua thấy các mảng 'gạch cua' nổi lên là do")
    opts_16 = [
        "A. sự đông tụ lipid.",
        "B. phản ứng thủy phân protein.",
        "*C. sự đông tụ protein.*",
        "D. phản ứng màu biuret."
    ]
    render_options_no_table(doc, opts_16, is_teacher=is_teacher)
    if is_teacher:
        sol_16 = [
            "*Chọn C.*",
            "Dưới tác dụng của nhiệt độ khi đun sôi, protein trong thịt cua bị biến tính và đông tụ tạo thành các mảng nổi lên trên bề mặt nước."
        ]
        add_solution(doc, sol_16)

    # -------------------------- PHẦN II --------------------------
    add_section_header(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # CÂU 1 PHẦN II
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Ethyl acetate là một chất lỏng không màu, dễ bay hơi, ít độc hại, không hút ẩm, có mùi ngọt và dễ chịu, thường được sử dụng trong việc sản xuất chất tẩy sơn móng tay, mực in, keo dán, thuốc lá, nước hoa.")
    stmts_p2_1 = [
        ("a) Đun nóng hỗn hợp gồm với 9,0 gam acetic acid và 9,2 gam ethyl alcohol có H~2~SO~4~ đặc làm chất xúc tác, thu được 10,56 gam este. Biết hiệu suất phản ứng đạt 80%.", "Đúng"),
        ("b) Ethyl acetate có khả năng tham gia phản ứng tráng bạc.", "Sai"),
        ("c) Ethyl acetate thuộc loại ester no, mạch hở, đơn chức.", "Đúng"),
        ("d) Công thức cấu tạo của ethyl acetate là CH~3~CH~2~COOCH~3~.", "Sai")
    ]
    for lbl, ans_val in stmts_p2_1:
        p_stmt = doc.add_paragraph()
        render_statement(p_stmt, lbl, ans=ans_val, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_1 = [
            "*Lời giải chi tiết:*",
            "- *Ý a) ĐÚNG.*",
            "  n_CH~3~COOH = 9,0 / 60 = 0,15 mol; n_C~2~H~5~OH = 9,2 / 46 = 0,2 mol.",
            "  Tính theo chất thiếu CH~3~COOH (0,15 mol): m_ester lí thuyết = 0,15 x 88 = 13,2 gam.",
            "  Hiệu suất: H = (10,56 / 13,2) x 100% = 80%.",
            "- *Ý b) SAI.* Ethyl acetate là CH~3~COOC~2~H~5~ không chứa nhóm aldehyde (-CHO) nên không tráng bạc.",
            "- *Ý c) ĐÚNG.* Ethyl acetate có CTPT C~4~H~8~O~2~ (C~n~H~2n~O~2~) thuộc loại ester no, đơn chức, mạch hở.",
            "- *Ý d) SAI.* Công thức cấu tạo của ethyl acetate là CH~3~COOCH~2~CH~3~. Còn CH~3~CH~2~COOCH~3~ là methyl propionate."
        ]
        add_solution(doc, sol_p2_1)

    # CÂU 2 PHẦN II
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Trong công nghiệp sản xuất bia có các bước chính sau:\nHạt đại mạch -(H~2~O, t^o^, enzyme)-> Dịch nấu (maltose, glucose) -(enzyme)-> Ethanol")
    stmts_p2_2 = [
        ("a) Để biết được thời điểm kết thúc quá trình thủy phân tinh bột ta có thể kiểm tra bằng thuốc thử là dung dịch I~2~ trong KI.", "Đúng"),
        ("b) Maltose và glucose thuộc nhóm monosaccharide, khi lên men thu được ethanol.", "Sai"),
        ("c) Thành phần trong hạt đại mạch bị thủy phân tạo ra maltose, glucose là tinh bột.", "Đúng"),
        ("d) Sản phẩm thủy phân của tinh bột (hạt đại mạch), ngoài maltose và glucose còn thu được sản phẩm phụ là fructose.", "Sai")
    ]
    for lbl, ans_val in stmts_p2_2:
        p_stmt = doc.add_paragraph()
        render_statement(p_stmt, lbl, ans=ans_val, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_2 = [
            "*Lời giải chi tiết:*",
            "- *Ý a) ĐÚNG.* Dung dịch I~2~ tạo màu xanh tím với tinh bột. Khi tinh bột đã bị thủy phân hoàn toàn thì dung dịch không còn chuyển màu xanh tím nữa.",
            "- *Ý b) SAI.* Maltose là disaccharide, glucose mới là monosaccharide.",
            "- *Ý c) ĐÚNG.* Tinh bột dự trữ trong hạt đại mạch nảy mầm bị enzyme amylase thủy phân sinh ra maltose và glucose.",
            "- *Ý d) SAI.* Tinh bột được cấu tạo hoàn toàn từ các mắt xích \\alpha-glucose, khi thủy phân không bao giờ sinh ra fructose."
        ]
        add_solution(doc, sol_p2_2)

    # -------------------------- PHẦN III --------------------------
    add_section_header(
        doc,
        "PHẦN III: Câu trắc nghiệm yêu cầu trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 4."
    )

    # CÂU 1 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Cho các chất sau: ethyl acetate, tripalmitin, saccharose, ethylamine, Gly-Ala, fibroin, Ala-Ala-Val. Số chất tham gia phản ứng thủy phân trong môi trường kiềm là bao nhiêu?")
    if is_teacher:
        sol_p3_1 = [
            "*Trả lời:* *5*",
            "*Lời giải:* Các chất tham gia phản ứng thủy phân trong môi trường kiềm gồm 5 chất:",
            "1. Ethyl acetate (ester).",
            "2. Tripalmitin (triglyceride/chất béo).",
            "3. Gly-Ala (dipeptide).",
            "4. Fibroin (protein trong tơ tằm).",
            "5. Ala-Ala-Val (tripeptide).",
            "(Saccharose không bị thủy phân trong kiềm; ethylamine là amine không bị thủy phân)."
        ]
        add_solution(doc, sol_p3_1)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 2 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Có bao nhiêu đồng phân cấu tạo amine bậc 1 ứng với công thức phân tử C~4~H~11~N?")
    if is_teacher:
        sol_p3_2 = [
            "*Trả lời:* *4*",
            "*Lời giải:* Amine bậc 1 có dạng R-NH~2~. Với CTPT C~4~H~11~N, gốc butyl (C~4~H~9~-) có 4 đồng phân:",
            "1. CH~3~-CH~2~-CH~2~-CH~2~-NH~2~ (butan-1-amine)",
            "2. CH~3~-CH~2~-CH(NH~2~)-CH~3~ (butan-2-amine)",
            "3. (CH~3~)~2~CH-CH~2~-NH~2~ (2-methylpropan-1-amine / isobutylamine)",
            "4. (CH~3~)~3~C-NH~2~ (2-methylpropan-2-amine / tert-butylamine)"
        ]
        add_solution(doc, sol_p3_2)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 3 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Để sản xuất ethanol, người ta dùng nguyên liệu là mùn cưa và vỏ bao từ gỗ chứa 50% cellulose. Nếu muốn điều chế 1 tấn ethanol, hiệu suất cả quá trình là 80% thì khối lượng nguyên liệu cần dùng là bao nhiêu kg? (Kết quả làm tròn đến hàng đơn vị).")
    if is_teacher:
        sol_p3_3 = [
            "*Trả lời:* *4402*",
            "*Lời giải:*",
            "- Khối lượng ethanol cần điều chế: m_ethanol = 1 tấn = 1 000 kg.",
            "- Sơ đồ: (C~6~H~10~O~5~)~n~ -> n C~6~H~12~O~6~ -> 2n C~2~H~5~OH",
            "  162 kg ----------------------------------> 2 x 46 = 92 kg",
            "- Khối lượng cellulose theo lí thuyết: m_lt = 1 000 x (162 / 92) = 1760,87 kg.",
            "- Khối lượng cellulose thực tế (hiệu suất H = 80%):",
            "  m_tt = 1760,87 / 80% = 2201,087 kg.",
            "- Khối lượng nguyên liệu cần dùng (chứa 50% cellulose):",
            "  m_nguyên liệu = 2201,087 / 50% = 4402,17 kg ≈ *4402* kg."
        ]
        add_solution(doc, sol_p3_3)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 4 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Saccharose có công thức phân tử C~12~H~22~O~11~, cấu tạo từ một đơn vị \\alpha-glucose và một đơn vị \\beta-fructose qua liên kết \\alpha-1,2-glycoside. Tổng số nhóm -OH trong phân tử saccharose là bao nhiêu?")
    if is_teacher:
        sol_p3_4 = [
            "*Trả lời:* *8*",
            "*Lời giải:* Trong phân tử saccharose:",
            "- Gốc \\alpha-glucose còn lại 4 nhóm -OH tự do (nhóm -OH ở C1 đã tham gia liên kết glycoside).",
            "- Gốc \\beta-fructose còn lại 4 nhóm -OH tự do (nhóm -OH ở C2 đã tham gia liên kết glycoside).",
            "=> Tổng số nhóm -OH trong phân tử saccharose = 4 + 4 = *8* nhóm."
        ]
        add_solution(doc, sol_p3_4)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # ========================== PHẦN B. TỰ LUẬN ==========================
    p_pb = doc.add_paragraph()
    p_pb.paragraph_format.space_before = Pt(8)
    p_pb.paragraph_format.space_after = Pt(4)
    render_rich_text(p_pb, "*B. PHẦN TỰ LUẬN (3,0 ĐIỂM)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # CÂU 1 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Một loại chất béo có chứa 80% tristearin về khối lượng. Để sản xuất 9,2 triệu chai nước rửa tay (có chứa chất dưỡng ẩm glycerol) cần dùng tối thiểu x tấn loại chất béo trên cho phản ứng với dung dịch NaOH, đun nóng. Biết rằng trong mỗi chai nước rửa tay có chứa 6 gam glycerol. Tính x?")
    if is_teacher:
        sol_tl_1 = [
            "*Lời giải chi tiết:*",
            "- Tổng khối lượng glycerol cần cho 9,2 triệu chai nước rửa tay:",
            "  m_glycerol = 9 200 000 x 6 gam = 55 200 000 gam = 55,2 tấn.",
            "- Số mol glycerol:",
            "  n_glycerol = 55 200 000 / 92 = 600 000 mol = 600 kmol.",
            "- Phản ứng xà phòng hóa tristearin:",
            "  (C~17~H~35~COO)~3~C~3~H~5~ + 3NaOH -(t^o^)-> 3C~17~H~35~COONa + C~3~H~5~(OH)~3~",
            "  Khối lượng mol của tristearin: M = 890 g/mol.",
            "- Theo phương trình: n_tristearin = n_glycerol = 600 kmol.",
            "  Khối lượng tristearin nguyên chất cần dùng:",
            "  m_tristearin = 600 x 890 = 534 000 kg = 534 tấn.",
            "- Do chất béo chỉ chứa 80% tristearin về khối lượng:",
            "  x = m_chất béo = 534 / 80% = *667,5* tấn."
        ]
        add_solution(doc, sol_tl_1)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # CÂU 2 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Glucose tồn tại đồng thời dạng mạch hở và mạch vòng (\\alpha và \\beta). Glucose có tính chất của aldehyde và của polyalcohol. Cho các chất sau: Cu(OH)~2~, nước bromine, dung dịch I~2~ trong KI, thuốc thử Tollens và CH~3~OH/HCl khan. Ở điều kiện thích hợp, liệt kê các chất tác dụng được với glucose?")
    if is_teacher:
        sol_tl_2 = [
            "*Lời giải:*",
            "Các chất tác dụng được với glucose gồm:",
            "1. *Cu(OH)~2~:* Tác dụng ở nhiệt độ thường tạo dung dịch phức màu xanh lam (tính polyalcohol); tác dụng khi đun nóng tạo kết tủa đỏ gạch Cu~2~O (tính aldehyde).",
            "2. *Nước bromine:* Oxi hóa nhóm -CHO làm mất màu dung dịch bromine.",
            "3. *Thuốc thử Tollens ([Ag(NH~3~)~2~]OH):* Oxi hóa nhóm -CHO tạo kết tủa bạc Ag trắng sáng.",
            "4. *CH~3~OH / HCl khan:* Phản ứng với nhóm -OH hemiacetal ở dạng vòng của glucose tạo methyl glucoside.",
            "(Dung dịch I~2~ trong KI không tác dụng với glucose).",
            "=> Danh sách: *Cu(OH)~2~, nước bromine, thuốc thử Tollens, CH~3~OH/HCl khan*."
        ]
        add_solution(doc, sol_tl_2)
    else:
        add_student_writing_lines(doc, num_lines=5)

    # CÂU 3 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Thực hiện phản ứng theo sơ đồ sau:\nAla + NaOH -> X + H~2~O.\nViết công thức cấu tạo và tính phân tử khối của X?")
    if is_teacher:
        sol_tl_3 = [
            "*Lời giải:*",
            "- Alanine là CH~3~-CH(NH~2~)-COOH (M = 89).",
            "- Phương trình phản ứng:",
            "  CH~3~-CH(NH~2~)-COOH + NaOH -> CH~3~-CH(NH~2~)-COONa + H~2~O",
            "- Công thức cấu tạo của X: *CH~3~-CH(NH~2~)-COONa* (sodium alaninate).",
            "- Phân tử khối của X: M_X = 89 + 23 - 1 = *111* g/mol."
        ]
        add_solution(doc, sol_tl_3)
    else:
        add_student_writing_lines(doc, num_lines=4)

    # CÂU 4 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Cây xanh nhả khí oxygen sau khi hoàn tất quá trình quang hợp, con người và nhiều loài động vật khác cần khí oxygen để duy trì sự sống. Phản ứng quang hợp được mô tả theo phương trình sau:\nn CO~2~ + m H~2~O -(ánh sáng, diệp lục)-> C~n~(H~2~O)~m~ + n O~2~\nGiả sử một cây xanh trưởng thành tiêu thụ lượng CO~2~ trung bình khoảng 22 kg/năm để thực hiện quá trình quang hợp và trung bình một năm mỗi người cần tối thiểu 1,2 tấn oxygen để duy trì sự sống. Để cung cấp đủ lượng oxygen cho 8 tỉ người trên Trái đất trong một năm thì cần bao nhiêu tỉ cây xanh như trên?")
    if is_teacher:
        sol_tl_4 = [
            "*Lời giải chi tiết:*",
            "- Tổng lượng oxygen cần cho 8 tỉ người trong 1 năm:",
            "  m_O2_cần = 8 tỉ người x 1,2 tấn = 9,6 tỉ tấn = 9,6 x 10^12^ kg.",
            "- Theo phương trình quang hợp: n mol CO~2~ tạo n mol O~2~.",
            "  Tỉ lệ khối lượng: Cứ 44 kg CO~2~ quang hợp sinh ra 32 kg O~2~.",
            "- Lượng O~2~ do 1 cây xanh trưởng thành sinh ra trong 1 năm:",
            "  m_O2_1_cây = 22 x (32 / 44) = 16 kg O~2~/năm.",
            "- Số cây xanh cần thiết để cung cấp đủ lượng O~2~ cho 8 tỉ người:",
            "  Số cây = (9,6 x 10^12^ kg) / (16 kg/cây) = 0,6 x 10^12^ cây = *600 tỉ cây xanh*."
        ]
        add_solution(doc, sol_tl_4)
    else:
        add_student_writing_lines(doc, num_lines=6)

    # CÂU 5 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 5. Tiến hành thí nghiệm theo các bước sau:\n- Bước 1: Cho khoảng 1 mL nước bromine vào ống nghiệm.\n- Bước 2: Thêm từ từ vài giọt dung dịch aniline loãng vào ống nghiệm.\nNêu hiện tượng và viết phương trình phản ứng?")
    if is_teacher:
        sol_tl_5 = [
            "*Lời giải chi tiết:*",
            "- *Hiện tượng:* Dung dịch nước bromine bị mất màu (nhạt màu vàng cam dần), đồng thời xuất hiện kết tủa màu trắng.",
            "- *Phương trình phản ứng:*",
            "  C~6~H~5~NH~2~ + 3Br~2~ -> C~6~H~2~Br~3~NH~2~ v (trắng) + 3HBr",
            "  (Sản phẩm kết tủa trắng là 2,4,6-tribromoaniline)."
        ]
        add_solution(doc, sol_tl_5)
    else:
        add_student_writing_lines(doc, num_lines=5)

    # CÂU 6 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 6. Tiến hành thí nghiệm với các chất X, Y, Z là một trong các chất sau: hồ tinh bột, albumin, triolein, vinyl acetate. Kết quả được ghi ở bảng sau:")

    tbl_exp = doc.add_table(rows=4, cols=3)
    tbl_exp.alignment = WD_TABLE_ALIGNMENT.CENTER
    exp_widths = [2.2, 9.8, 6.0]
    exp_data = [
        ("Mẫu thử", "Thí nghiệm", "Hiện tượng"),
        ("X", "Đun nóng với dung dịch NaOH (loãng, dư), để nguội.\nThêm tiếp vài giọt dung dịch CuSO~4~", "Tạo dung dịch màu xanh lam"),
        ("Y", "Đun nóng với dung dịch NaOH loãng (vừa đủ).\nThêm tiếp dung dịch AgNO~3~ trong NH~3~, đun nóng", "Tạo kết tủa Ag"),
        ("Z", "Tác dụng với dung dịch I~2~ loãng", "Có màu xanh tím")
    ]
    for r_idx, row in enumerate(tbl_exp.rows):
        for c_idx, cell in enumerate(row.cells):
            cell.width = Cm(exp_widths[c_idx])
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
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            if r_idx == 0:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8F5E9"/>')
                tcPr.append(shd)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                render_rich_text(p, f"*{exp_data[r_idx][c_idx]}*", base_size=10, base_bold=True)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
                render_rich_text(p, exp_data[r_idx][c_idx], base_size=10)

    p_req = doc.add_paragraph()
    p_req.paragraph_format.space_before = Pt(3)
    p_req.paragraph_format.space_after = Pt(2)
    render_rich_text(p_req, "Tìm các chất X, Y, Z.", base_size=10.5)

    if is_teacher:
        sol_tl_6 = [
            "*Lời giải chi tiết:*",
            "- Mẫu Z tác dụng với I~2~ tạo màu xanh tím đặc trưng => *Z là hồ tinh bột*.",
            "- Mẫu Y thủy phân trong NaOH tạo aldehyde acetic CH~3~CHO có phản ứng tráng bạc sinh kết tủa Ag:",
            "  CH~3~COOCH=CH~2~ + NaOH -(t^o^)-> CH~3~COONa + CH~3~CHO.",
            "  CH~3~CHO + 2[Ag(NH~3~)~2~]OH -(t^o^)-> CH~3~COONH~4~ + 2Ag v + 3NH~3~ + H~2~O.",
            "  => *Y là vinyl acetate* (CH~3~COOCH=CH~2~).",
            "- Mẫu X thủy phân trong NaOH dư tạo glycerol; glycerol hòa tan Cu(OH)~2~ trong môi trường kiềm tạo phức màu xanh lam: 2C~3~H~5~(OH)~3~ + Cu(OH)~2~ -> [C~3~H~5~(OH)~2~O]~2~Cu + 2H~2~O => *X là triolein* ((C~17~H~33~COO)~3~C~3~H~5~).",
            "(Albumin tạo phức màu tím với Cu(OH)~2~ do phản ứng biuret).",
            "- *Kết luận:* X là triolein; Y là vinyl acetate; Z là hồ tinh bột."
        ]
        add_solution(doc, sol_tl_6)
    else:
        add_student_writing_lines(doc, num_lines=5)

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
    doc_hs = build_exam_document_ntt103(is_teacher=False)
    file_hs = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma103_HocSinh.docx")
    doc_hs.save(file_hs)
    print(f"[OK] Đã tạo thành công Bản Học sinh: {file_hs}")

    file_hs_std = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma103.docx")
    doc_hs.save(file_hs_std)
    print(f"[OK] Đã tạo thành công Bản Đề Chuẩn: {file_hs_std}")

    # 2. Bản Giáo viên
    doc_gv = build_exam_document_ntt103(is_teacher=True)
    file_gv = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma103_GiaoVien.docx")
    doc_gv.save(file_gv)
    print(f"[OK] Đã tạo thành công Bản Giáo viên: {file_gv}")


if __name__ == "__main__":
    main()
