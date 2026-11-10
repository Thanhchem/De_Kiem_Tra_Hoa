# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Kỳ I Hóa Học 12 - THPT Nguyễn Trường Tộ - Mã Đề 106
Năm học 2026 - 2027
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản Word:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma106_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma106_GiaoVien.docx
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


def add_quick_answer_table_ntt106(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu MÃ ĐỀ 106.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 106)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I (16 câu)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (16 câu)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    # Chia bảng 16 câu thành 2 hàng cho đẹp mắt
    tbl1 = doc.add_table(rows=2, cols=17)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_1 = [2.0] + [1.0] * 16
    p1_answers = ["D", "B", "B", "B", "D", "D", "D", "B", "B", "B", "A", "B", "C", "C", "A", "A"]

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
        ("Câu 1", "S", "S", "Đ", "Đ"),
        ("Câu 2", "S", "S", "Đ", "S"),
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
    p3_answers = ["2159", "5", "3", "22"]

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


def build_exam_document_ntt106(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="TRƯỜNG THPT NGUYỄN TRƯỜNG TỘ\nTỔ HÓA - TIN",
        exam_title="ĐỀ KIỂM TRA GIỮA KỲ I - NĂM HỌC 2026 - 2027",
        subject_line="MÔN: HÓA HỌC LỚP 12",
        num_pages="04",
        duration="45 phút",
        exam_code="106",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; S = 32; Cu = 64; Ag = 108",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_ntt106(doc)

    # ========================== PHẦN A. TRẮC NGHIỆM ==========================
    p_pa = doc.add_paragraph()
    p_pa.paragraph_format.space_before = Pt(6)
    p_pa.paragraph_format.space_after = Pt(2)
    render_rich_text(p_pa, "*A. PHẦN TRẮC NGHIỆM*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # -------------------------- PHẦN I --------------------------
    add_section_header(
        doc,
        "Phần I: Câu trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # CÂU 1
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Khi nấu canh cua thấy các mảng 'gạch cua' nổi lên là do")
    opts_1 = [
        "A. sự đông tụ lipid.",
        "B. phản ứng thủy phân protein.",
        "C. phản ứng màu biuret.",
        "*D. sự đông tụ protein.*"
    ]
    render_options_no_table(doc, opts_1, is_teacher=is_teacher)
    if is_teacher:
        sol_1 = [
            "*Chọn D.*",
            "Trong thịt cua chứa lượng lớn protein hòa tan trong nước. Khi đun sôi (dưới tác dụng của nhiệt độ), protein bị biến tính và đông tụ lại thành từng mảng nổi lên mặt nước ('gạch cua')."
        ]
        add_solution(doc, sol_1)

    # CÂU 2
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Công thức cấu tạo của methylamine là")
    opts_2 = [
        "A. C~2~H~5~NH~2~.",
        "*B. CH~3~NH~2~.*",
        "C. CH~3~NHCH~3~.",
        "D. (CH~3~)~3~N."
    ]
    render_options_no_table(doc, opts_2, is_teacher=is_teacher)
    if is_teacher:
        sol_2 = [
            "*Chọn B.*",
            "Methylamine gồm gốc methyl (-CH~3~) liên kết với nhóm amino (-NH~2~), có công thức phân tử và cấu tạo là CH~3~NH~2~."
        ]
        add_solution(doc, sol_2)

    # CÂU 3
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Dung dịch chất X hòa tan Cu(OH)~2~, thu được dung dịch màu xanh lam. Mặt khác, X bị thủy phân khi đun nóng trong môi trường acid. Chất X là")
    opts_3 = [
        "A. glucose.",
        "*B. saccharose.*",
        "C. amylose.",
        "D. fructose."
    ]
    render_options_no_table(doc, opts_3, is_teacher=is_teacher)
    if is_teacher:
        sol_3 = [
            "*Chọn B.*",
            "- Glucose và fructose là các monosaccharide nên không bị thủy phân.",
            "- Amylose không tan nhiều trong nước lạnh và không tạo dung dịch xanh lam đặc trưng với Cu(OH)~2~.",
            "- Saccharose là disaccharide có nhiều nhóm -OH kề nhau nên hòa tan Cu(OH)~2~ tạo dung dịch phức màu xanh lam, đồng thời bị thủy phân trong môi trường acid tạo glucose và fructose."
        ]
        add_solution(doc, sol_3)

    # CÂU 4
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Ester X có công thức phân tử C~4~H~8~O~2~. Thủy phân X trong dung dịch H~2~SO~4~ loãng, đun nóng, thu được sản phẩm gồm acetic acid và chất hữu cơ Y. Công thức của Y là")
    opts_4 = [
        "A. HCOOH.",
        "*B. C~2~H~5~OH.*",
        "C. CH~3~COOH.",
        "D. CH~3~OH."
    ]
    render_options_no_table(doc, opts_4, is_teacher=is_teacher)
    if is_teacher:
        sol_4 = [
            "*Chọn B.*",
            "Acetic acid là CH~3~COOH. Ester X có công thức C~4~H~8~O~2~ tạo ra acetic acid thì X phải là ethyl acetate: CH~3~COOC~2~H~5~.",
            "Phương trình: CH~3~COOC~2~H~5~ + H~2~O <=(H~2~SO~4~, t^o^)=> CH~3~COOH + C~2~H~5~OH.",
            "Vậy chất hữu cơ Y là alcohol C~2~H~5~OH (ethanol)."
        ]
        add_solution(doc, sol_4)

    # CÂU 5
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 5. Mùi tanh của cá là hỗn hợp các amine và một số tạp chất khác. Để khử mùi tanh của cá trước khi chế biến thực phẩm, nên áp dụng cách nào sau đây?")
    opts_5 = [
        "A. Rửa cá bằng dung dịch nước muối.",
        "B. Ngâm cá trong nước để amine tan vào nước.",
        "C. Rửa cá bằng dung dịch soda.",
        "*D. Rửa cá bằng giấm ăn.*"
    ]
    render_options_no_table(doc, opts_5, is_teacher=is_teacher)
    if is_teacher:
        sol_5 = [
            "*Chọn D.*",
            "Mùi tanh của cá do các amine (như trimethylamine) gây ra. Amine có tính base yếu, khi dùng giấm ăn (chứa acetic acid CH~3~COOH) sẽ xảy ra phản ứng trung hòa tạo thành muối tan không mùi và dễ bị rửa trôi bằng nước."
        ]
        add_solution(doc, sol_5)

    # CÂU 6
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 6. Khi thuỷ phân lipid trong môi trường base thì thu được muối của acid béo và")
    opts_6 = [
        "A. ester đơn chức.",
        "B. phenol.",
        "C. alcohol đơn chức.",
        "*D. glycerol.*"
    ]
    render_options_no_table(doc, opts_6, is_teacher=is_teacher)
    if is_teacher:
        sol_6 = [
            "*Chọn D.*",
            "Lipid (chất béo/triglyceride) là triester của glycerol với các acid béo. Khi thủy phân trong dung dịch kiềm (phản ứng xà phòng hóa) luôn thu được xà phòng (muối của acid béo) và glycerol: (RCOO)~3~C~3~H~5~ + 3NaOH -> 3RCOONa + C~3~H~5~(OH)~3~."
        ]
        add_solution(doc, sol_6)

    # CÂU 7
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 7. Giá trị pH mà khi đó amino acid có nồng độ ion lưỡng cực là cực đại được gọi là điểm đẳng điện (kí hiệu là pI). Khi pH < pI thì amino acid đó tồn tại chủ yếu ở dạng cation, còn khi pH > pI thì amino acid đó tồn tại chủ yếu ở dạng anion. Khi đặt trong một điện trường dạng anion sẽ di chuyển về cực (+) còn dạng cation sẽ di chuyển về cực (-). Tính chất này được gọi là tính điện di và được dùng để tách, tinh chế amino acid ra khỏi hỗn hợp của chúng. Cho các giá trị pI của các chất sau:\n- Glycine (H~2~NCH~2~COOH): pI = 6,0\n- Glutamic acid (HOOCCH~2~CH~2~CH(NH~2~)COOH): pI = 3,2\n- Lysine (H~2~N[CH~2~]~4~CH(NH~2~)COOH): pI = 9,7\nTrong các giá trị pH cho dưới đây, giá trị nào là tối ưu nhất để tách ba chất trên ra khỏi dung dịch hỗn hợp của chúng?")
    opts_7 = [
        "A. pH = 14,0.",
        "B. pH = 9,7.",
        "C. pH = 3,2.",
        "*D. pH = 6,0.*"
    ]
    render_options_no_table(doc, opts_7, is_teacher=is_teacher)
    if is_teacher:
        sol_7 = [
            "*Chọn D.*",
            "Tại pH = 6,0:",
            "- Với glycine: pH = pI = 6,0 => glycine tồn tại chủ yếu ở dạng ion lưỡng cực có điện tích tổng cộng bằng 0, không di chuyển trong điện trường (đứng yên tại chỗ).",
            "- Với glutamic acid: pI = 3,2 < pH (6,0) => tồn tại ở dạng anion mang điện tích âm, di chuyển về cực dương (+).",
            "- Với lysine: pI = 9,7 > pH (6,0) => tồn tại ở dạng cation mang điện tích dương, di chuyển về cực âm (-).",
            "Như vậy ở pH = 6,0, ba chất di chuyển theo 3 hướng khác nhau (1 chất đứng yên, 1 chất về cực +, 1 chất về cực -), giúp tách triệt để hỗn hợp ba chất này."
        ]
        add_solution(doc, sol_7)

    # CÂU 8
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 8. Tripeptide là hợp chất mà phân tử có")
    opts_8 = [
        "A. ba liên kết peptide, hai gốc \\alpha-amino acid.",
        "*B. hai liên kết peptide, ba gốc \\alpha-amino acid.*",
        "C. ba liên kết peptide, ba gốc \\alpha-amino acid.",
        "D. hai liên kết peptide, ba gốc \\beta-amino acid."
    ]
    render_options_no_table(doc, opts_8, is_teacher=is_teacher)
    if is_teacher:
        sol_8 = [
            "*Chọn B.*",
            "Peptide là hợp chất chứa từ 2 đến 50 gốc \\alpha-amino acid liên kết với nhau bằng các liên kết peptide. Phân tử tripeptide được tạo nên từ 3 gốc \\alpha-amino acid và có chứa 2 liên kết peptide (-CO-NH-)."
        ]
        add_solution(doc, sol_8)

    # CÂU 9
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 9. Chất nào sau đây vừa tác dụng được với H~2~NCH~2~COOH, vừa tác dụng được với CH~3~NH~2~?")
    opts_9 = [
        "A. NaOH.",
        "*B. HCl.*",
        "C. NaCl.",
        "D. CH~3~OH."
    ]
    render_options_no_table(doc, opts_9, is_teacher=is_teacher)
    if is_teacher:
        sol_9 = [
            "*Chọn B.*",
            "- Methylamine (CH~3~NH~2~) có tính base nên tác dụng được với acid HCl tạo muối methylammonium chloride: CH~3~NH~2~ + HCl -> CH~3~NH~3~Cl. (CH~3~NH~2~ không phản ứng với base NaOH).",
            "- Glycine (H~2~NCH~2~COOH) có tính chất lưỡng tính nên tác dụng được với dung dịch acid HCl: H~2~NCH~2~COOH + HCl -> ClH~3~NCH~2~COOH."
        ]
        add_solution(doc, sol_9)

    # CÂU 10
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 10. Phần không phân cực ('đuôi' kị nước) trong xà phòng và chất giặt rửa là")
    opts_10 = [
        "A. nhóm sulfonate.",
        "*B. gốc hydrocarbon dài.*",
        "C. nhóm sulfate.",
        "D. Nhóm carboxylate."
    ]
    render_options_no_table(doc, opts_10, is_teacher=is_teacher)
    if is_teacher:
        sol_10 = [
            "*Chọn B.*",
            "Phân tử xà phòng và chất giặt rửa tổng hợp gồm 2 phần cấu trúc:",
            "- Phần 'đầu' phân cực (ưa nước): là các nhóm ion như carboxylate (-COO^-^), sulfonate (-SO~3~^-^), hoặc sulfate (-OSO~3~^-^).",
            "- Phần 'đuôi' không phân cực (kị nước, ưa dầu mỡ): là gốc hydrocarbon mạch dài (R-)."
        ]
        add_solution(doc, sol_10)

    # CÂU 11
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 11. Glucose là một loại monosaccharide có nhiều trong quả nho chín. Công thức phân tử của glucose là")
    opts_11 = [
        "*A. C~6~H~12~O~6~.*",
        "B. (C~6~H~10~O~5~)~n~.",
        "C. C~2~H~4~O~2~.",
        "D. C~12~H~22~O~11~."
    ]
    render_options_no_table(doc, opts_11, is_teacher=is_teacher)
    if is_teacher:
        sol_11 = [
            "*Chọn A.*",
            "Glucose có công thức phân tử là C~6~H~12~O~6~ (M = 180 g/mol)."
        ]
        add_solution(doc, sol_11)

    # CÂU 12
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 12. Nhóm chức nào sau đây không có trong cấu tạo của glucose?")
    opts_12 = [
        "A. Aldehyde.",
        "*B. Ketone.*",
        "C. Hemiacetal.",
        "D. Hydroxy."
    ]
    render_options_no_table(doc, opts_12, is_teacher=is_teacher)
    if is_teacher:
        sol_12 = [
            "*Chọn B.*",
            "Glucose ở dạng mạch hở chứa 1 nhóm chức aldehyde (-CHO) và 5 nhóm hydroxy (-OH). Ở dạng mạch vòng, nhóm -CHO tạo nhóm -OH hemiacetal. Glucose không chứa nhóm chức ketone (nhóm C=O dạng ketone có trong fructose)."
        ]
        add_solution(doc, sol_12)

    # CÂU 13
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 13. Ở động vật, tinh bột được dự trữ dưới dạng glycogen ở đâu?")
    opts_13 = [
        "A. Ruột",
        "B. Dạ dày",
        "*C. Gan*",
        "D. Máu"
    ]
    render_options_no_table(doc, opts_13, is_teacher=is_teacher)
    if is_teacher:
        sol_13 = [
            "*Chọn C.*",
            "Glycogen (được ví là tinh bột động vật) được tổng hợp và tích trữ chủ yếu ở gan và các cơ bắp để điều hòa lượng đường huyết."
        ]
        add_solution(doc, sol_13)

    # CÂU 14
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 14. Số ester có cùng công thức phân tử C~3~H~6~O~2~ là")
    opts_14 = [
        "A. 4.",
        "B. 3.",
        "*C. 2.*",
        "D. 5."
    ]
    render_options_no_table(doc, opts_14, is_teacher=is_teacher)
    if is_teacher:
        sol_14 = [
            "*Chọn C.*",
            "Các ester đồng phân của C~3~H~6~O~2~ gồm có 2 chất:",
            "1. HCOOC~2~H~5~ (ethyl formate)",
            "2. CH~3~COOCH~3~ (methyl acetate)"
        ]
        add_solution(doc, sol_14)

    # CÂU 15
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 15. Monosaccharide X được dùng trong công nghiệp để tráng bạc lên bề mặt thuỷ tinh trong sản xuất ruột phích. Cùng với Ag, sản phẩm hữu cơ được tạo thành khi cho X tác dụng với lượng dư dung dịch AgNO~3~ trong NH~3~ là")
    opts_15 = [
        "*A. ammonium gluconate.*",
        "B. ammonium carbonate.",
        "C. gluconic acid.",
        "D. khí carbon dioxide."
    ]
    render_options_no_table(doc, opts_15, is_teacher=is_teacher)
    if is_teacher:
        sol_15 = [
            "*Chọn A.*",
            "Monosaccharide X là glucose. Khi phản ứng với thuốc thử Tollens [Ag(NH~3~)~2~]OH trong môi trường kiềm ammonia, nhóm aldehyde bị oxi hóa thành muối ammonium carboxylate tương ứng là ammonium gluconate:",
            "CH~2~OH[CHOH]~4~CHO + 2[Ag(NH~3~)~2~]OH -(t^o^)-> CH~2~OH[CHOH]~4~COONH~4~ + 2Ag v + 3NH~3~ + H~2~O."
        ]
        add_solution(doc, sol_15)

    # CÂU 16
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 16. Chất có công thức phân tử C~12~H~22~O~11~, được tạo thành trong quá trình thuỷ phân không hoàn toàn amylose có trong tinh bột là")
    opts_16 = [
        "*A. maltose.*",
        "B. glucose.",
        "C. saccharose.",
        "D. fructose."
    ]
    render_options_no_table(doc, opts_16, is_teacher=is_teacher)
    if is_teacher:
        sol_16 = [
            "*Chọn A.*",
            "Amylose là mạch polysaccharide gồm các mắt xích \\alpha-glucose liên kết qua liên kết \\alpha-1,4-glycosidic. Khi thủy phân không hoàn toàn tinh bột (bằng enzyme amylase), thu được disaccharide maltose (C~12~H~22~O~11~)."
        ]
        add_solution(doc, sol_16)

    # -------------------------- PHẦN II --------------------------
    add_section_header(
        doc,
        "Phần II: Câu hỏi trả lời đúng sai.",
        "Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # CÂU 1 PHẦN II
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Tiến hành thí nghiệm thủy phân ester theo các bước sau:\n- Bước 1: Cho vào hai ống nghiệm mỗi ống 2 mL ethyl acetate.\n- Bước 2: Thêm 2 mL dung dịch H~2~SO~4~ 20% vào ống thứ nhất; 4 mL dung dịch NaOH 30% vào ống thứ hai.\n- Bước 3: Lắc đều cả hai ống nghiệm, lắp ống sinh hàn, đun sôi nhẹ trong khoảng 5 phút, để nguội.")
    stmts_p2_1 = [
        ("a) Phản ứng xảy ra trong ống nghiệm 1 nhanh hơn phản ứng trong ống 2.", "Sai"),
        ("b) Vai trò của dung dịch H~2~SO~4~ 20% và dung dịch NaOH 30% trong thí nghiệm trên là chất xúc tác và không bị tiêu hao sau phản ứng.", "Sai"),
        ("c) Sau bước 2, chất lỏng trong cả hai ống nghiệm đều phân thành hai lớp.", "Đúng"),
        ("d) Sau bước 3: ống nghiệm (1) thể tích lớp chất lỏng phía trên giảm. Ống nghiệm (2) tạo thành hỗn hợp đồng nhất.", "Đúng")
    ]
    for lbl, ans_val in stmts_p2_1:
        p_stmt = doc.add_paragraph()
        render_statement(p_stmt, lbl, ans=ans_val, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_1 = [
            "*Lời giải chi tiết:*",
            "- *Ý a) SAI.* Phản ứng trong ống 2 là phản ứng thủy phân trong môi trường kiềm (phản ứng một chiều, xà phòng hóa) xảy ra nhanh hơn và triệt để hơn phản ứng thủy phân thuận nghịch trong môi trường acid ở ống 1.",
            "- *Ý b) SAI.* H~2~SO~4~ đóng vai trò là chất xúc tác (không bị tiêu hao), còn NaOH đóng vai trò là chất phản ứng và bị tiêu hao sau phản ứng tạo muối sodium acetate: CH~3~COOC~2~H~5~ + NaOH -> CH~3~COONa + C~2~H~5~OH.",
            "- *Ý c) ĐÚNG.* Ở bước 2 chưa đun nóng nên phản ứng thủy phân chưa xảy ra đáng kể. Do ethyl acetate ít tan trong nước và có khối lượng riêng nhỏ hơn nước (d ~ 0,9 g/mL) nên nổi lên trên, hình thành 2 lớp chất lỏng ở cả hai ống nghiệm.",
            "- *Ý d) ĐÚNG.* Ở ống 1, phản ứng thuận nghịch làm tiêu hao một phần ester nên lớp chất lỏng nổi bên trên giảm thể tích. Ở ống 2, phản ứng một chiều xảy ra hoàn toàn tạo CH~3~COONa và C~2~H~5~OH đều tan tốt trong nước tạo thành một dung dịch đồng nhất."
        ]
        add_solution(doc, sol_p2_1)

    # CÂU 2 PHẦN II
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Cho cấu tạo phân tử saccharose gồm một gốc \\alpha-glucose và một gốc \\beta-fructose liên kết với nhau qua liên kết \\alpha-1,2-glycoside (nguyên tử C1 của gốc glucose liên kết qua cầu oxygen với nguyên tử C2 của gốc fructose).")
    stmts_p2_2 = [
        ("a) Ngoài cấu tạo dạng mạch vòng, saccharose còn có cả cấu tạo ở dạng mở vòng của đơn vị cấu tạo glucose và fructose.", "Sai"),
        ("b) Mỗi phân tử saccharose gồm một đơn vị \\alpha-glucose và một đơn vị \\alpha-fructose liên kết với nhau qua liên kết \\alpha-1,2-glycoside.", "Sai"),
        ("c) Saccharose và maltose đều có 42,1% C về khối lượng.", "Đúng"),
        ("d) Phân tử saccharose không có nhóm -OH hemiacetal, nhưng có nhóm -OH hemiketal.", "Sai")
    ]
    for lbl, ans_val in stmts_p2_2:
        p_stmt = doc.add_paragraph()
        render_statement(p_stmt, lbl, ans=ans_val, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_2 = [
            "*Lời giải chi tiết:*",
            "- *Ý a) SAI.* Trong phân tử saccharose, cả nhóm -OH hemiacetal (ở C1 của glucose) và -OH hemiketal (ở C2 của fructose) đều đã tham gia tạo liên kết glycosidic. Do đó các vòng này không thể mở vòng được, saccharose chỉ tồn tại ở dạng mạch vòng, không có dạng mở vòng.",
            "- *Ý b) SAI.* Saccharose được tạo nên từ một gốc *\\alpha-glucose* và một gốc *\\beta-fructose* (đề ghi '\\alpha-fructose' là sai).",
            "- *Ý c) ĐÚNG.* Cả saccharose và maltose đều là disaccharide có cùng công thức phân tử C~12~H~22~O~11~ (M = 342 g/mol). Phần trăm khối lượng carbon: %C = (12 x 12 / 342) x 100% = 42,1%.",
            "- *Ý d) SAI.* Do cả hai nhóm -OH hemiacetal và -OH hemiketal đều đã liên kết với nhau tạo cầu nối -O- nên phân tử saccharose *không còn cả nhóm -OH hemiacetal lẫn nhóm -OH hemiketal*."
        ]
        add_solution(doc, sol_p2_2)

    # -------------------------- PHẦN III --------------------------
    add_section_header(
        doc,
        "Phần III: Câu trắc nghiệm trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 4."
    )

    # CÂU 1 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Ethanol có thể được sản xuất từ cellulose hoặc tinh bột. Loại ethanol này được dùng để sản xuất xăng E10 (xăng chứa 10% ethanol về thể tích). Lượng ethanol thu được từ 1 tấn mùn cưa (chứa 50% cellulose, phần còn lại là chất trơ) có thể dùng để pha chế bao nhiêu L xăng E10? Biết hiệu suất quá trình sản xuất ethanol từ cellulose là 60% và ethanol có khối lượng riêng là 0,789 g/mL. (Kết quả làm tròn đến hàng đơn vị).")
    if is_teacher:
        sol_p3_1 = [
            "*Trả lời:* *2159*",
            "*Lời giải:*",
            "- Khối lượng cellulose nguyên chất có trong 1 tấn mùn cưa:",
            "  m_cellulose = 1 000 x 50% = 500 kg = 500 000 gam.",
            "- Sơ đồ phản ứng sản xuất ethanol:",
            "  (C~6~H~10~O~5~)~n~ -> n C~6~H~12~O~6~ -> 2n C~2~H~5~OH",
            "  162n gam ----------------------------> 2n x 46 = 92n gam",
            "- Khối lượng ethanol thu được trên thực tế với hiệu suất H = 60%:",
            "  m_ethanol = 500 000 x (92 / 162) x 60% = 170 370,37 gam.",
            "- Thể tích ethanol nguyên chất tương ứng:",
            "  V_ethanol = m / D = 170 370,37 / 0,789 = 215 932,03 mL = 215,932 L.",
            "- Thể tích xăng E10 pha chế chế tạo được (chứa 10% ethanol về thể tích):",
            "  V_xăng E10 = V_ethanol / 10% = 215,932 / 0,10 = 2159,32 L.",
            "- Làm tròn đến hàng đơn vị: *2159*."
        ]
        add_solution(doc, sol_p3_1)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 2 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Cho các chất sau: ethyl acetate, tripalmitin, saccharose, ethylamine, Gly-Ala, albumin, Ala-Ala-Val. Số chất tham gia phản ứng thủy phân trong môi trường kiềm là bao nhiêu?")
    if is_teacher:
        sol_p3_2 = [
            "*Trả lời:* *5*",
            "*Lời giải:* Các chất có phản ứng thủy phân trong môi trường kiềm gồm:",
            "1. Ethyl acetate (ester): CH~3~COOC~2~H~5~ + NaOH -(t^o^)-> CH~3~COONa + C~2~H~5~OH. -> *(1)*",
            "2. Tripalmitin (chất béo): (C~15~H~31~COO)~3~C~3~H~5~ + 3NaOH -(t^o^)-> 3C~15~H~31~COONa + C~3~H~5~(OH)~3~. -> *(2)*",
            "3. Saccharose: disaccharide, chỉ bị thủy phân trong môi trường acid hoặc xúc tác enzyme, bền và không bị thủy phân trong môi trường kiềm.",
            "4. Ethylamine: amine đơn chức, không có liên kết có thể thủy phân.",
            "5. Gly-Ala (dipeptide): thủy phân liên kết peptide tạo muối của glycine và alanine trong NaOH. -> *(3)*",
            "6. Albumin (protein): protein bị thủy phân hoàn toàn trong dung dịch kiềm nóng tạo các muối của \\alpha-amino acid. -> *(4)*",
            "7. Ala-Ala-Val (tripeptide): thủy phân trong dung dịch kiềm nóng. -> *(5)*",
            "=> Tổng cộng có 5 chất tham gia phản ứng thủy phân trong môi trường kiềm."
        ]
        add_solution(doc, sol_p3_2)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 3 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Cho các dung dịch sau: HCl, NaOH, CH~3~COOH, CuSO~4~. Có bao nhiêu dung dịch tác dụng được với ethylamine?")
    if is_teacher:
        sol_p3_3 = [
            "*Trả lời:* *3*",
            "*Lời giải:* Ethylamine (C~2~H~5~NH~2~) tác dụng được với 3 dung dịch gồm:",
            "1. Dung dịch HCl: C~2~H~5~NH~2~ + HCl -> C~2~H~5~NH~3~Cl. -> *(1)*",
            "2. Dung dịch CH~3~COOH: C~2~H~5~NH~2~ + CH~3~COOH -> CH~3~COONH~3~C~2~H~5~. -> *(2)*",
            "3. Dung dịch CuSO~4~: Ban đầu tạo kết tủa Cu(OH)~2~ xanh nhạt, sau đó ethylamine dư tạo phức chất tan màu xanh lam đậm: [Cu(C~2~H~5~NH~2~)~4~]^2+^. -> *(3)*",
            "(Dung dịch NaOH không tác dụng vì cả hai đều có tính base)."
        ]
        add_solution(doc, sol_p3_3)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 4 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Một phân tử saccharose có bao nhiêu nguyên tử hydrogen?")
    if is_teacher:
        sol_p3_4 = [
            "*Trả lời:* *22*",
            "*Lời giải:* Công thức phân tử của saccharose là C~12~H~22~O~11~. Do đó một phân tử saccharose chứa đúng 22 nguyên tử hydrogen."
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
    render_rich_text(p_pb, "*B. PHẦN TỰ LUẬN*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # CÂU 1 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Hoàn thành phương trình phản ứng sau:\nCH~3~COOC~6~H~5~ (phenyl acetate) + NaOH dư (dung dịch, đun nóng).")
    if is_teacher:
        sol_tl_1 = [
            "*Lời giải:*",
            "Phenyl acetate là ester của phenol. Khi thủy phân trong dung dịch kiềm đun nóng, phenol sinh ra có tính acid tiếp tục tác dụng với lượng dư NaOH tạo muối phenolate:",
            "CH~3~COOC~6~H~5~ + 2NaOH -(t^o^)-> CH~3~COONa + C~6~H~5~ONa + H~2~O"
        ]
        add_solution(doc, sol_tl_1)
    else:
        add_student_writing_lines(doc, num_lines=4)

    # CÂU 2 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Glucose tồn tại đồng thời dạng mạch hở và mạch vòng (\\alpha và \\beta). Glucose có tính chất của aldehyde và của polyalcohol. Cho các chất sau: Cu(OH)~2~, nước bromine, dung dịch I~2~ trong KI, thuốc thử Tollens và CH~3~OH/HCl khan. Ở điều kiện thích hợp, liệt kê các chất tác dụng được với glucose?")
    if is_teacher:
        sol_tl_2 = [
            "*Lời giải:*",
            "Các chất tác dụng được với glucose gồm:",
            "1. *Cu(OH)~2~:* Tác dụng ở nhiệt độ thường tạo phức xanh lam (tính chất polyalcohol); tác dụng khi đun nóng tạo kết tủa đỏ gạch Cu~2~O (tính chất aldehyde).",
            "2. *Nước bromine:* Nhóm -CHO bị oxi hóa làm mất màu dung dịch bromine tạo gluconic acid.",
            "3. *Thuốc thử Tollens ([Ag(NH~3~)~2~]OH):* Oxi hóa nhóm -CHO tạo kết tủa Ag trắng sáng (phản ứng tráng bạc).",
            "4. *CH~3~OH / HCl khan:* Nhóm -OH hemiacetal ở dạng vòng của glucose phản ứng tạo liên kết methyl glucoside.",
            "(Dung dịch I~2~ trong KI chỉ dùng nhận biết hồ tinh bột, không phản ứng với glucose).",
            "=> Danh sách các chất tác dụng được: *Cu(OH)~2~, nước bromine, thuốc thử Tollens, CH~3~OH/HCl khan*."
        ]
        add_solution(doc, sol_tl_2)
    else:
        add_student_writing_lines(doc, num_lines=5)

    # CÂU 3 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Viết công thức cấu tạo của Alanine.")
    if is_teacher:
        sol_tl_3 = [
            "*Lời giải:*",
            "Alanine là một \\alpha-amino acid có tên thay thế là 2-aminopropanoic acid.",
            "Công thức cấu tạo: *CH~3~-CH(NH~2~)-COOH* (hoặc H~2~N-CH(CH~3~)-COOH)."
        ]
        add_solution(doc, sol_tl_3)
    else:
        add_student_writing_lines(doc, num_lines=3)

    # CÂU 4 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Trong cơ thể, glucose là sản phẩm của quá trình thủy phân tinh bột dưới tác dụng của enzyme. Trong quá trình hô hấp tế bào, 1 mol glucose bị oxi hóa hoàn toàn qua các phản ứng sinh hóa tạo thành carbon dioxide, nước và giải phóng năng lượng là 2880 kJ. Trong khẩu phần ăn hằng ngày của một người trưởng thành cần 300 gam gạo. Hàm lượng tinh bột trong gạo là 80%. Giả sử 60% lượng tinh bột trong gạo chuyển hóa thành glucose và toàn bộ lượng glucose này tham gia vào quá trình hô hấp tế bào. Năng lượng giải phóng từ sự oxi hóa hoàn toàn glucose chiếm bao nhiêu % so với nhu cầu năng lượng trung bình (2000 Calori/ngày) của một người trưởng thành? (Kết quả làm tròn đến hàng phần mười).\nCho biết: 1 Calori = 1 kcal = 1000 cal = 4184 J.")
    if is_teacher:
        sol_tl_4 = [
            "*Lời giải chi tiết:*",
            "- Đổi nhu cầu năng lượng trung bình mỗi ngày sang đơn vị kJ:",
            "  E_nhucau = 2000 Calori = 2000 x 4184 J = 8 368 000 J = 8368 kJ.",
            "- Khối lượng tinh bột có trong 300 gam gạo:",
            "  m_tinhbot = 300 x 80% = 240 gam.",
            "- Khối lượng tinh bột thực tế chuyển hóa thành glucose (hiệu suất 60%):",
            "  m_tinhbot_phanung = 240 x 60% = 144 gam.",
            "- Phương trình thủy phân tinh bột:",
            "  (C~6~H~10~O~5~)~n~ + n H~2~O -(enzyme)-> n C~6~H~12~O~6~",
            "  Cứ 162 gam tinh bột chuyển hóa thành 1 mol glucose.",
            "- Số mol glucose sinh ra và tham gia hô hấp tế bào:",
            "  n_glucose = 144 / 162 = 8/9 mol ≈ 0,8889 mol.",
            "- Năng lượng giải phóng khi oxi hóa hoàn toàn lượng glucose này:",
            "  E_giaiphong = (8/9) x 2880 kJ = 2560 kJ.",
            "- Phần trăm năng lượng giải phóng so với nhu cầu năng lượng trung bình:",
            "  % = (2560 / 8368) x 100% = *30,6%*."
        ]
        add_solution(doc, sol_tl_4)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # CÂU 5 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 5. Amine X có phân tử khối nhỏ hơn 80. Trong phân tử X, nitrogen chiếm 19,18% về khối lượng. Cho X tác dụng với nitrous acid (HNO~2~) ở nhiệt độ thường thu được alcohol Y. Oxi hóa không hoàn toàn Y thu được ketone Z. Viết công thức cấu tạo của Y.")
    if is_teacher:
        sol_tl_5 = [
            "*Lời giải chi tiết:*",
            "- Gọi công thức của amine X là R(NH~2~)~t~.",
            "  Vì M_X < 80 và %N = 19,18% nên số nguyên tử N trong phân tử là: t = (80 x 19,18%) / (14 x 100%) = 1,095 => t = 1.",
            "- Khối lượng mol phân tử của amine X:",
            "  M_X = 14 / 19,18% = 73 g/mol.",
            "- Amine X tác dụng với HNO~2~ ở nhiệt độ thường giải phóng khí N~2~ tạo alcohol Y => X là amine bậc 1, no, đơn chức, mạch hở C~n~H~2n+1~NH~2~.",
            "  Ta có: 14n + 17 = 73 => 14n = 56 => n = 4.",
            "  Công thức phân tử của X là C~4~H~11~N.",
            "- Alcohol Y thu được là C~4~H~9~OH.",
            "- Oxi hóa không hoàn toàn Y thu được ketone Z => Y phải là alcohol bậc 2.",
            "  Trong các đồng phân alcohol C~4~H~9~OH, chỉ có butan-2-ol là alcohol bậc 2: CH~3~-CH(OH)-CH~2~-CH~3~.",
            "- Công thức cấu tạo của alcohol Y là: *CH~3~-CH(OH)-CH~2~-CH~3~* (butan-2-ol)."
        ]
        add_solution(doc, sol_tl_5)
    else:
        add_student_writing_lines(doc, num_lines=6)

    # CÂU 6 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 6. Tiến hành thí nghiệm với các chất X, Y, Z là một trong các chất sau: hồ tinh bột, albumin, ethyl acetate, vinyl acetate. Kết quả được ghi ở bảng sau:")

    # Thêm bảng thí nghiệm đề bài (đây là bảng dữ liệu đề bài, được phép theo chuẩn đề thi)
    tbl_exp = doc.add_table(rows=4, cols=3)
    tbl_exp.alignment = WD_TABLE_ALIGNMENT.CENTER
    exp_widths = [2.2, 9.8, 6.0]
    exp_data = [
        ("Mẫu thử", "Thí nghiệm", "Hiện tượng"),
        ("X", "Tác dụng với Cu(OH)~2~ trong môi trường kiềm", "Có màu tím"),
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
            "- Mẫu thử X tác dụng với Cu(OH)~2~ trong môi trường kiềm cho phức chất màu tím đặc trưng => phản ứng màu biuret của protein => *X là albumin*.",
            "- Mẫu thử Z tác dụng với dung dịch I~2~ loãng cho dung dịch màu xanh tím => phản ứng đặc trưng của tinh bột => *Z là hồ tinh bột*.",
            "- Mẫu thử Y thủy phân trong dung dịch NaOH tạo sản phẩm có khả năng tham gia phản ứng tráng bạc với thuốc thử Tollens [Ag(NH~3~)~2~]OH:",
            "  + Ethyl acetate (CH~3~COOC~2~H~5~) thủy phân tạo CH~3~COONa và C~2~H~5~OH (cả 2 đều không tráng bạc).",
            "  + Vinyl acetate (CH~3~COOCH=CH~2~) thủy phân tạo CH~3~COONa và aldehyde acetic (CH~3~CHO):",
            "    CH~3~COOCH=CH~2~ + NaOH -(t^o^)-> CH~3~COONa + CH~3~CHO.",
            "    CH~3~CHO + 2[Ag(NH~3~)~2~]OH -(t^o^)-> CH~3~COONH~4~ + 2Ag v + 3NH~3~ + H~2~O.",
            "  => *Y là vinyl acetate* (CH~3~COOCH=CH~2~).",
            "- *Kết luận:* X là albumin; Y là vinyl acetate; Z là hồ tinh bột."
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
    doc_hs = build_exam_document_ntt106(is_teacher=False)
    file_hs = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma106_HocSinh.docx")
    doc_hs.save(file_hs)
    print(f"[OK] Đã tạo thành công Bản Học sinh: {file_hs}")

    file_hs_std = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma106.docx")
    doc_hs.save(file_hs_std)
    print(f"[OK] Đã tạo thành công Bản Đề Chuẩn: {file_hs_std}")

    # 2. Bản Giáo viên
    doc_gv = build_exam_document_ntt106(is_teacher=True)
    file_gv = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma106_GiaoVien.docx")
    doc_gv.save(file_gv)
    print(f"[OK] Đã tạo thành công Bản Giáo viên: {file_gv}")


if __name__ == "__main__":
    main()
