# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Kỳ I Hóa Học 12 - THPT Nguyễn Huệ - Mã Đề 1202
Năm học 2026 - 2027
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản Word:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_12_NguyenHue_Ma1202_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_12_NguyenHue_Ma1202_GiaoVien.docx
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


def add_quick_answer_table_nh1202(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu MÃ ĐỀ 1202.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 1202)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I (12 câu)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (12 câu)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_1 = [2.4] + [1.3] * 12
    p1_answers = ["B", "C", "B", "A", "B", "C", "B", "A", "A", "A", "D", "B"]

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
        ("Câu 1", "S", "Đ", "Đ", "S"),
        ("Câu 2", "S", "Đ", "Đ", "Đ"),
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
    p3_answers = ["2", "3", "26,5", "145"]

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
                render_rich_text(p, f"*{txt}*", base_size=9.5, base_bold=True)
            else:
                txt = "Đáp án" if col_idx == 0 else p3_answers[col_idx - 1]
                if col_idx == 0:
                    render_rich_text(p, f"*{txt}*", base_size=9.5, base_bold=True)
                else:
                    render_rich_text(p, f"*{txt}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(8)
    add_p_border_bottom(p_div, sz="6", color="8BC390")


def build_exam_document_nh1202(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="TRƯỜNG THPT NGUYỄN HUỆ\nTỔ LÍ - HÓA - CÔNG NGHỆ",
        exam_title="KIỂM TRA GIỮA KỲ I - NĂM HỌC 2026 - 2027",
        subject_line="MÔN: HÓA HỌC, LỚP 12",
        num_pages="04",
        duration="45 phút",
        exam_code="1202",
        atomic_masses="H = 1; C = 12; N = 14; O = 16",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_nh1202(doc)

    # ========================== PHẦN A. TRẮC NGHIỆM ==========================
    p_pa = doc.add_paragraph()
    p_pa.paragraph_format.space_before = Pt(6)
    p_pa.paragraph_format.space_after = Pt(2)
    render_rich_text(p_pa, "*A. PHẦN TRẮC NGHIỆM (7,0 ĐIỂM)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # -------------------------- PHẦN I --------------------------
    add_section_header(
        doc,
        "PHẦN I. Câu hỏi trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # CÂU 1
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Chất nào sau đây thuộc loại disaccharide?")
    opts_1 = [
        "A. Amylopectin.",
        "*B. Maltose.*",
        "C. Cellulose.",
        "D. Amylose."
    ]
    render_options_no_table(doc, opts_1, is_teacher=is_teacher)
    if is_teacher:
        sol_1 = [
            "*Chọn B.*",
            "Carbohydrate được phân thành 3 loại chính: monosaccharide (glucose, fructose), disaccharide (saccharose, maltose) và polysaccharide (tinh bột, cellulose). Vậy maltose là disaccharide."
        ]
        add_solution(doc, sol_1)

    # CÂU 2
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Số liên kết peptide có trong một phân tử Ala-Gly-Val-Gly-Ala là")
    opts_2 = [
        "A. 2.",
        "B. 5.",
        "*C. 4.*",
        "D. 3."
    ]
    render_options_no_table(doc, opts_2, is_teacher=is_teacher)
    if is_teacher:
        sol_2 = [
            "*Chọn C.*",
            "Phân tử Ala-Gly-Val-Gly-Ala là pentapeptide gồm 5 gốc alpha-amino acid. Số liên kết peptide = 5 - 1 = 4."
        ]
        add_solution(doc, sol_2)

    # CÂU 3
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Tiến hành thí nghiệm tìm hiểu một số tính chất hóa học của protein theo các bước sau:\n- Bước 1: Cho vào hai ống nghiệm (1) và (2), mỗi ống 2 mL dung dịch lòng trắng trứng.\n- Bước 2: Đun nóng ống nghiệm (1) trên ngọn lửa đèn cồn trong 2 - 3 phút; thêm vài giọt dung dịch HNO~3~ đặc vào ống nghiệm (2).\nPhát biểu nào sau đây không đúng ?")
    opts_3 = [
        "A. Sau khi đun nóng, albumin trong ống nghiệm (1) bị đông tụ, tạo thành chất rắn màu trắng.",
        "*B. Hiện tượng ở ống nghiệm (1) chứng tỏ protein đã bị thủy phân hoàn toàn thành các alpha-amino acid.*",
        "C. Sau khi thêm HNO~3~ đặc, trong ống nghiệm (2) xuất hiện chất rắn màu vàng.",
        "D. Nếu ngâm ống nghiệm (1) trong nước nóng ở khoảng 80 °C trong thời gian thích hợp thì hiện tượng đông tụ protein vẫn xảy ra."
    ]
    render_options_no_table(doc, opts_3, is_teacher=is_teacher)
    if is_teacher:
        sol_3 = [
            "*Chọn B.*",
            "Ở ống nghiệm (1), khi đun nóng albumin bị đông tụ do biến tính nhiệt, tạo khối kết tủa màu trắng. Đây là sự biến tính và đông tụ protein, không phải là phản ứng thủy phân hoàn toàn thành alpha-amino acid."
        ]
        add_solution(doc, sol_3)

    # CÂU 4
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Aniline có thể được điều chế bằng cách khử chất X bằng hỗn hợp Fe và dung dịch HCl. Chất X là")
    opts_4 = [
        "*A. C~6~H~5~NO~2~.*",
        "B. C~6~H~5~NH~2~.",
        "C. C~6~H~5~NO~3~.",
        "D. C~6~H~5~N~2~Cl."
    ]
    render_options_no_table(doc, opts_4, is_teacher=is_teacher)
    if is_teacher:
        sol_4 = [
            "*Chọn A.*",
            "Aniline được điều chế bằng phản ứng khử nitrobenzene (C~6~H~5~NO~2~) bằng nguyên tử H mới sinh từ hỗn hợp Fe và dung dịch HCl:\nC~6~H~5~NO~2~ + 6[H] -(Fe + HCl)-> C~6~H~5~NH~2~ + 2H~2~O."
        ]
        add_solution(doc, sol_4)

    # CÂU 5
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 5. Lipid là những hợp chất ...(1)... có trong tế bào sống, không tan trong nước nhưng tan trong dung môi hữu cơ ...(2)...\nNội dung phù hợp trong các ô trống (1), (2) lần lượt là:")
    opts_5 = [
        "A. vô cơ; phân cực.",
        "*B. hữu cơ; không phân cực.*",
        "C. vô cơ; không phân cực.",
        "D. hữu cơ; phân cực."
    ]
    render_options_no_table(doc, opts_5, is_teacher=is_teacher)
    if is_teacher:
        sol_5 = [
            "*Chọn B.*",
            "Lipid là những hợp chất hữu cơ có trong tế bào sống, không tan trong nước nhưng tan nhiều trong các dung môi hữu cơ không phân cực hoặc kém phân cực như ether, chloroform, benzene, hexane,..."
        ]
        add_solution(doc, sol_5)

    # CÂU 6
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 6. Loại protein nào sau đây tan được trong nước?")
    opts_6 = [
        "A. Keratin (có ở tóc, móng).",
        "B. Collagen (có ở da, sụn).",
        "*C. Hemoglobin (có ở máu).*",
        "D. Myosin (có ở cơ bắp)."
    ]
    render_options_no_table(doc, opts_6, is_teacher=is_teacher)
    if is_teacher:
        sol_6 = [
            "*Chọn C.*",
            "Protein dạng hình cầu (như albumin lòng trắng trứng, hemoglobin trong máu) tan được trong nước tạo dung dịch keo. Protein dạng hình sợi (keratin, collagen, myosin, fibroin) không tan trong nước."
        ]
        add_solution(doc, sol_6)

    # CÂU 7
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 7. Số nhóm amino (-NH~2~) có trong một phân tử Lysine là")
    opts_7 = [
        "A. 3.",
        "*B. 2.*",
        "C. 1.",
        "D. 4."
    ]
    render_options_no_table(doc, opts_7, is_teacher=is_teacher)
    if is_teacher:
        sol_7 = [
            "*Chọn B.*",
            "Công thức cấu tạo của lysine: H~2~N-[CH~2~]~4~-CH(NH~2~)-COOH. Phân tử lysine có 2 nhóm amino (-NH~2~) và 1 nhóm carboxyl (-COOH)."
        ]
        add_solution(doc, sol_7)

    # CÂU 8
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 8. Cho các chất hữu cơ sau: ethyl acetate (CH~3~COOC~2~H~5~), methyl formate (HCOOCH~3~), propanoic acid (C~2~H~5~COOH) và propan-1-ol (CH~3~CH~2~CH~2~OH). Biết nhiệt độ sôi của các chất trên ở 1 atm (không theo thứ tự) là: 97,2 °C; 77,1 °C; 32,0 °C và 141 °C. Nhiệt độ sôi của ethyl acetate là bao nhiêu?")
    opts_8 = [
        "*A. 77,1 °C.*",
        "B. 32,0 °C.",
        "C. 141 °C.",
        "D. 97,2 °C."
    ]
    render_options_no_table(doc, opts_8, is_teacher=is_teacher)
    if is_teacher:
        sol_8 = [
            "*Chọn A.*",
            "- Acid carboxylic (C~2~H~5~COOH) có liên kết hydrogen liên phân tử bền nhất nên nhiệt độ sôi cao nhất: 141 °C.\n- Alcohol (CH~3~CH~2~CH~2~OH) có liên kết hydrogen liên phân tử nên nhiệt độ sôi cao thứ hai: 97,2 °C.\n- Hai ester không có liên kết hydrogen: ethyl acetate (CH~3~COOC~2~H~5~, M = 88) có phân tử khối và tương tác van der Waals lớn hơn methyl formate (HCOOCH~3~, M = 60), nên nhiệt độ sôi của ethyl acetate là 77,1 °C (methyl formate là 32,0 °C)."
        ]
        add_solution(doc, sol_8)

    # CÂU 9
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 9. Saccharose có nhiều trong nguyên liệu thực phẩm nào sau đây?")
    opts_9 = [
        "*A. Thân cây mía.*",
        "B. Gạo.",
        "C. Mật ong.",
        "D. Lúa mạch."
    ]
    render_options_no_table(doc, opts_9, is_teacher=is_teacher)
    if is_teacher:
        sol_9 = [
            "*Chọn A.*",
            "Saccharose (đường mía) có nhiều nhất trong thân cây mía (13% - 15%), củ cải đường và hoa thốt nốt. Gạo chứa nhiều tinh bột; mật ong chứa nhiều fructose; mầm lúa mạch chứa nhiều maltose."
        ]
        add_solution(doc, sol_9)

    # CÂU 10
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 10. Tiến hành thí nghiệm theo các bước sau:\n- Bước 1: Cho vào cốc thuỷ tinh chịu nhiệt khoảng 5 gam dầu dừa và 10 ml dung dịch NaOH 40%.\n- Bước 2: Đun hỗn hợp trong khoảng 10 phút, liên tục khuấy bằng đũa thuỷ tinh và thỉnh thoảng thêm nước.\n- Bước 3: Rót hỗn hợp vào cốc thuỷ tinh chứa khoảng 70 mL dung dịch NaCl bão hoà, khuấy nhẹ rồi để yên.\nPhát biểu nào sau đây về thí nghiệm trên không đúng ?")
    opts_10 = [
        "*A. Ở bước 1, nếu thay dầu dừa bằng mỡ lợn thì sản phẩm thu được sau bước 3 hoàn toàn giống nhau.*",
        "B. Sau bước 3, xuất hiện lớp chất rắn màu trắng, chủ yếu là xà phòng, nổi lên trên bề mặt.",
        "C. Trong thí nghiệm trên, có xảy ra phản ứng xà phòng hóa chất béo.",
        "D. Ở bước 3, thêm dung dịch NaCl bão hòa là để tách muối của acid béo ra khỏi hỗn hợp."
    ]
    render_options_no_table(doc, opts_10, is_teacher=is_teacher)
    if is_teacher:
        sol_10 = [
            "*Chọn A.*",
            "Dầu dừa chứa chủ yếu gốc acid lauric (C~11~H~23~COOH), còn mỡ lợn chứa chủ yếu gốc acid palmitic, stearic, oleic... Muối sodium của acid béo (xà phòng) thu được từ hai nguồn chất béo trên là khác nhau, không thể hoàn toàn giống nhau."
        ]
        add_solution(doc, sol_10)

    # CÂU 11
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 11. Protein là hợp chất cao phân tử được cấu tạo từ một hay nhiều chuỗi")
    opts_11 = [
        "A. polysaccharide.",
        "B. lipid.",
        "C. nucleic acid.",
        "*D. polypeptide.*"
    ]
    render_options_no_table(doc, opts_11, is_teacher=is_teacher)
    if is_teacher:
        sol_11 = [
            "*Chọn D.*",
            "Protein là những hợp chất cao phân tử được cấu tạo từ một hay nhiều chuỗi polypeptide."
        ]
        add_solution(doc, sol_11)

    # CÂU 12
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 12. Monosaccharide X có trong nhiều loại trái cây (táo, lựu, nho, lê,...), trong một số loại rau củ (cà rốt, củ cải đường,...) và có nhiều nhất trong mật ong. Tên gọi của X là")
    opts_12 = [
        "A. maltose.",
        "*B. fructose.*",
        "C. saccharose.",
        "D. glucose."
    ]
    render_options_no_table(doc, opts_12, is_teacher=is_teacher)
    if is_teacher:
        sol_12 = [
            "*Chọn B.*",
            "Fructose là monosaccharide có vị ngọt sắc nhất, chiếm khoảng 40% khối lượng trong mật ong và có nhiều trong các loại trái cây chín."
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
    render_question_prompt(p, "Câu 1. Chất X có hai thành phần chính, cấu tạo một đoạn mạch của mỗi thành phần được biểu diễn như hình sau:")

    # Chèn hình ảnh cấu trúc Tinh bột
    img_path = os.path.join(r"C:\Antigravity_Thanh\San_Pham\images_gk1_hoa12_nguyenhue_1202\cau1_p2_tinh_bot.png")
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(4)
        run = p_img.add_run()
        run.add_picture(img_path, width=Cm(12.5))

    statements_p2_1 = [
        ("a) Thành phần (2) có cấu trúc phân nhánh do sự xuất hiện của các liên kết beta-1,6-glycoside.", "Sai"),
        ("b) X là một polysaccharide.", "Đúng"),
        ("c) Trong thực vật, X có chủ yếu trong củ, quả hay hạt.", "Đúng"),
        ("d) Tiến hành lần lượt các thí nghiệm của X (ở dạng dung dịch keo nhớt) với: dung dịch AgNO~3~/NH~3~, đun nóng (1); dung dịch I~2~ (2); Cu(OH)~2~ (3); dung dịch HCl, đun nóng (4); dung dịch NaOH (5). Số thí nghiệm X có phản ứng là 3.", "Sai"),
    ]
    for text, ans_val in statements_p2_1:
        p_stmt = doc.add_paragraph()
        render_statement(p_stmt, text, ans=ans_val, is_teacher=is_teacher)

    if is_teacher:
        sol_p2_1 = [
            "*Lời giải chi tiết:*",
            "Chất X là tinh bột, gồm 2 thành phần:",
            "- Thành phần (1) là *amylose*: mạch không phân nhánh, các gốc alpha-D-glucose liên kết qua liên kết alpha-1,4-glycosidic.",
            "- Thành phần (2) là *amylopectin*: mạch phân nhánh, gồm liên kết alpha-1,4-glycosidic và liên kết alpha-1,6-glycosidic tại điểm phân nhánh.",
            "Xét các phát biểu:",
            "a) *Sai.* Điểm phân nhánh của amylopectin là liên kết *alpha-1,6-glycosidic*, không phải beta-1,6-glycoside.",
            "b) *Đúng.* Tinh bột là polymer thiên nhiên thuộc loại polysaccharide có công thức (C~6~H~10~O~5~)~n~.",
            "c) *Đúng.* Tinh bột là chất dinh dưỡng dự trữ của thực vật, tích tụ nhiều trong hạt ngũ cốc, các loại củ và quả chín/xanh.",
            "d) *Sai.* Xét 5 thí nghiệm với dung dịch keo hồ tinh bột:",
            "   (1) AgNO~3~/NH~3~, t°: Không phản ứng (tinh bột không có nhóm -CHO tự do).",
            "   (2) Dung dịch I~2~: Phân tử tinh bột hấp phụ I~2~ tạo phức chất màu xanh tím -> CÓ phản ứng màu.",
            "   (3) Cu(OH)~2~: Không phản ứng ở nhiệt độ thường.",
            "   (4) HCl, t°: Tinh bột bị thủy phân thành alpha-glucose -> CÓ phản ứng hóa học.",
            "   (5) Dung dịch NaOH: Tinh bột không bị thủy phân trong môi trường kiềm -> Không phản ứng.",
            "   Vậy chỉ có 2 thí nghiệm xảy ra phản ứng/biến đổi đặc trưng là (2) và (4), không phải 3."
        ]
        add_solution(doc, sol_p2_1)

    # CÂU 2 PHẦN II
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Cho các phát biểu sau về ba chất: aniline, glycine và valine.")
    statements_p2_2 = [
        ("a) Tương tự glycine, ở pH = 11, valine di chuyển về cực âm (-) khi đặt trong điện trường.", "Sai"),
        ("b) Có 2 chất là amino acid thiên nhiên.", "Đúng"),
        ("c) Ở điều kiện thường, aniline là chất lỏng, ít tan trong nước.", "Đúng"),
        ("d) Cả glycine và valine đều tác dụng được với dung dịch HCl và dung dịch NaOH.", "Đúng"),
    ]
    for text, ans_val in statements_p2_2:
        p_stmt = doc.add_paragraph()
        render_statement(p_stmt, text, ans=ans_val, is_teacher=is_teacher)

    if is_teacher:
        sol_p2_2 = [
            "*Lời giải chi tiết:*",
            "a) *Sai.* Điểm đẳng điện (pI) của glycine và valine khoảng 6,0. Ở môi trường base pH = 11 (pH > pI), nhóm carboxyl nhường H^+ tạo anion carboxylate: H~2~N-CH(R)-COO^-. Do mang điện tích âm, khi đặt trong điện trường chúng di chuyển về *cực dương (+)*, không phải cực âm.",
            "b) *Đúng.* Glycine và valine là 2 alpha-amino acid thiên nhiên cấu tạo nên protein của cơ thể. Aniline là amine thơm.",
            "c) *Đúng.* Ở điều kiện thường, aniline là chất lỏng không màu, mùi hôi đặc trưng, độc, nặng hơn nước và rất ít tan trong nước.",
            "d) *Đúng.* Cả glycine và valine đều là alpha-amino acid có tính chất lưỡng tính nên đều phản ứng được với acid mạnh (HCl) và base mạnh (NaOH)."
        ]
        add_solution(doc, sol_p2_2)

    # -------------------------- PHẦN III --------------------------
    add_section_header(
        doc,
        "PHẦN III. Câu trắc nghiệm trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 4."
    )

    # CÂU 1 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Trong số các chất sau: oleic acid, sodium alkylsulfate, triolein, ethyl acetate, sodium stearate, glucose, có bao nhiêu chất có tính chất giặt rửa?")
    if is_teacher:
        sol_p3_1 = [
            "*Trả lời: 2*",
            "*Lời giải chi tiết:*",
            "Các chất có tính chất giặt rửa gồm:",
            "- Xà phòng: sodium stearate (C~17~H~35~COONa).",
            "- Chất giặt rửa tổng hợp: sodium alkylsulfate (ROSO~3~Na).",
            "Oleic acid (acid béo không tan), triolein (chất béo), ethyl acetate (ester) và glucose (đường) không có tính giặt rửa.",
            "=> Có *2* chất có tính chất giặt rửa."
        ]
        add_solution(doc, sol_p3_1)

    # CÂU 2 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Cho các chất sau: CH~3~-NH~2~, H~2~N-CH~2~-COOH, CH~3~-CH~2~-NO~2~, CH~3~-NH-CH~3~, C~6~H~5~-CH~2~-NH~2~, CH~3~-NH~3~Cl. Số chất thuộc loại amine là bao nhiêu?")
    if is_teacher:
        sol_p3_2 = [
            "*Trả lời: 3*",
            "*Lời giải chi tiết:*",
            "Amine là hợp chất thu được khi thay thế nguyên tử H trong NH~3~ bằng gốc hydrocarbon.",
            "- CH~3~-NH~2~ (methylamine): amine bậc 1.",
            "- H~2~N-CH~2~-COOH (glycine): amino acid.",
            "- CH~3~-CH~2~-NO~2~ (nitroethane): hợp chất nitro.",
            "- CH~3~-NH-CH~3~ (dimethylamine): amine bậc 2.",
            "- C~6~H~5~-CH~2~-NH~2~ (benzylamine): amine bậc 1.",
            "- CH~3~-NH~3~Cl (methylammonium chloride): muối ammonium.",
            "=> Có *3* chất thuộc loại amine là: CH~3~-NH~2~, CH~3~-NH-CH~3~, C~6~H~5~-CH~2~-NH~2~."
        ]
        add_solution(doc, sol_p3_2)

    # CÂU 3 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Thủy phân hoàn toàn tetrapeptide mạch hở X thu được hỗn hợp gồm: glycine, alanine và valine với tỉ lệ mol tương ứng là 2:1:1. Phần trăm khối lượng của oxygen có trong X là bao nhiêu (làm tròn kết quả đến hàng phần mười)?")
    if is_teacher:
        sol_p3_3 = [
            "*Trả lời: 26,5*",
            "*Lời giải chi tiết:*",
            "- Tetrapeptide X được cấu tạo từ 2 gốc Gly, 1 gốc Ala và 1 gốc Val.",
            "- Công thức phân tử của X: C~n~H~m~N~4~O~5~.",
            "  + Số nguyên tử C = 2*2 + 3 + 5 = 12.",
            "  + Số nguyên tử H = (2*5 + 7 + 11) - 3*2 = 28 - 6 = 22.",
            "  + Số nguyên tử N = 4.",
            "  + Số nguyên tử O = 4 + 1 = 5.",
            "  => CTPT của X là C~12~H~22~N~4~O~5~.",
            "- Phân tử khối của X: M_X = 12*12 + 22*1 + 14*4 + 16*5 = 302 g/mol.",
            "- Phần trăm khối lượng của oxygen trong X:",
            "  %m_O = (5 * 16 / 302) * 100% = (80 / 302) * 100% = 26,490% -> làm tròn đến hàng phần mười là *26,5%*."
        ]
        add_solution(doc, sol_p3_3)

    # CÂU 4 PHẦN III
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Học sinh tiến hành các thí nghiệm về tính chất hóa học của amine và nêu hiện tượng xảy ra tương ứng như sau:\n- Thí nghiệm 1: Cho từ từ nước bromine vào ống nghiệm đựng dung dịch aniline và lắc đều thấy xuất hiện chất rắn màu trắng.\n- Thí nghiệm 2: Cho dung dịch methylamine vào ống nghiệm đựng dung dịch iron (III) chloride thấy xuất hiện chất rắn màu trắng hơi xanh sau đó hóa nâu đỏ.\n- Thí nghiệm 3: Nhỏ một giọt dung dịch aniline lên mẩu giấy quỳ tím đặt trên mặt kính đồng hồ thấy giấy quỳ chuyển sang màu xanh.\n- Thí nghiệm 4: Nhỏ từ từ dung dịch methylamine vào ống nghiệm đựng kết tủa copper (II) hydroxide thấy kết tủa tan ra tạo dung dịch phức chất có màu xanh lam.\n- Thí nghiệm 5: Cho từ từ dung dịch ethylamine vào ống nghiệm đựng dung dịch nitrous acid thấy có bọt khí không màu thoát ra.\nLiệt kê các thí nghiệm có hiện tượng phù hợp tương ứng theo thứ tự tăng dần của số thứ tự (ví dụ: 134, 23, ...).")
    if is_teacher:
        sol_p3_4 = [
            "*Trả lời: 145*",
            "*Lời giải chi tiết:*",
            "- Thí nghiệm 1: Phù hợp. Aniline phản ứng với nước bromine tạo kết tủa trắng 2,4,6-tribromoaniline.",
            "- Thí nghiệm 2: Không phù hợp. Dung dịch methylamine tác dụng với FeCl~3~ tạo ngay kết tủa nâu đỏ Fe(OH)~3~ (kết tủa trắng hơi xanh hóa nâu đỏ là Fe(OH)~2~ của Fe^2+^).",
            "- Thí nghiệm 3: Không phù hợp. Aniline có tính base rất yếu không làm đổi màu quỳ tím.",
            "- Thí nghiệm 4: Phù hợp. Dung dịch methylamine hòa tan Cu(OH)~2~ tạo dung dịch phức chất [Cu(CH~3~NH~2~)~4~](OH)~2~ có màu xanh lam thẫm.",
            "- Thí nghiệm 5: Phù hợp. Ethylamine phản ứng với HNO~2~ sinh ra bọt khí N~2~ không màu: C~2~H~5~NH~2~ + HNO~2~ -> C~2~H~5~OH + N~2~↑ + H~2~O.",
            "=> Các thí nghiệm có hiện tượng phù hợp: 1, 4, 5. Ghi kết quả: *145*."
        ]
        add_solution(doc, sol_p3_4)

    # ========================== PHẦN B. TỰ LUẬN ==========================
    p_pb = doc.add_paragraph()
    p_pb.paragraph_format.space_before = Pt(8)
    p_pb.paragraph_format.space_after = Pt(4)
    render_rich_text(p_pb, "*B. PHẦN TỰ LUẬN (3,0 ĐIỂM)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # CÂU 1 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. (1,0 điểm) Ester X có công thức phân tử là C~4~H~8~O~2~. Xà phòng hóa hoàn toàn X thu được muối Y và alcohol Z. Biết số nguyên tử carbon trong một phân tử Z lớn hơn số nguyên tử carbon trong một phân tử Y. Viết tất cả các công thức cấu tạo thu gọn có thể có của X và tên gọi tương ứng (mỗi công thức cấu tạo chỉ ghi một tên gọi).")
    if is_teacher:
        sol_tl_1 = [
            "*Lời giải chi tiết:*",
            "- Ester X đơn chức, mạch hở no có CTPT C~4~H~8~O~2~, dạng RCOOR'.",
            "  Phương trình xà phòng hóa: RCOOR' + NaOH -> RCOONa (muối Y) + R'OH (alcohol Z).",
            "- Gọi C_Y và C_Z lần lượt là số nguyên tử C trong Y và Z. Ta có: C_Y + C_Z = 4.",
            "- Theo đề bài: C_Z > C_Y => C_Z > 2, nên C_Z = 3 và C_Y = 1.",
            "- Muối Y có 1 nguyên tử C là HCOONa (sodium formate).",
            "- Alcohol Z có 3 nguyên tử C là C~3~H~7~OH, có 2 gốc R' tương ứng là: propyl (-CH~2~CH~2~CH~3~) và isopropyl (-CH(CH~3~)~2~).",
            "- Các công thức cấu tạo thu gọn và tên gọi của X:",
            "  1. *HCOOCH~2~CH~2~CH~3~*: propyl formate",
            "  2. *HCOOCH(CH~3~)~2~*: isopropyl formate"
        ]
        add_solution(doc, sol_tl_1)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # CÂU 2 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. (1,0 điểm) Cellulose nitrate (nitrocellulose) có nhiều ứng dụng trong đời sống. Khi cho cellulose tác dụng với hỗn hợp HNO~3~ đặc và H~2~SO~4~ đặc trong điều kiện thích hợp, tùy theo mức độ phản ứng, số lượng nhóm nitrate (-ONO~2~) được tạo thành ở các mắt xích có thể khác nhau.\na) Giải thích vì sao mỗi mắt xích C~6~H~10~O~5~ của cellulose có thể phản ứng với tối đa 3 phân tử HNO~3~.\nb) Viết phương trình hóa học của phản ứng tạo thành cellulose trinitrate, một chất được sử dụng trong sản xuất thuốc súng không khói.")
    if is_teacher:
        sol_tl_2 = [
            "*Lời giải chi tiết:*",
            "a) *Giải thích:*",
            "Mỗi mắt xích C~6~H~10~O~5~ trong phân tử cellulose có 3 nhóm hydroxyl (-OH) tự do, có công thức dạng [C~6~H~7~O~2~(OH)~3~]. Khi phản ứng với HNO~3~ đặc (xúc tác H~2~SO~4~ đặc, đun nóng), mỗi nhóm -OH bị ester hóa bởi một phân tử HNO~3~ để chuyển thành nhóm nitrate (-ONO~2~). Do chỉ có tối đa 3 nhóm -OH tự do nên mỗi mắt xích chỉ phản ứng được với tối đa 3 phân tử HNO~3~.",
            "b) *Phương trình hóa học điều chế cellulose trinitrate:*",
            "[C~6~H~7~O~2~(OH)~3~]~n~ + 3n HNO~3~ (đặc) -(H~2~SO~4~ đặc, t°)-> [C~6~H~7~O~2~(ONO~2~)~3~]~n~ + 3n H~2~O"
        ]
        add_solution(doc, sol_tl_2)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # CÂU 3 TỰ LUẬN
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. (1,0 điểm) Một học sinh tiến hành thí nghiệm thử tính chất của carbohydrate như sau:\nBước 1: Cho khoảng 5 mL dung dịch hồ tinh bột 1% vào ống nghiệm (1). Sau đó thêm khoảng 1 mL dung dịch HCl 1 M vào, lắc đều.\nBước 2: Đặt ống nghiệm (1) trong một cốc thuỷ tinh chứa nước nóng, đun cách thuỷ trong 10 phút. Sau đó để nguội.\nBước 3: Thêm từ từ NaHCO~3~ vào ống nghiệm (1) đến khi ngừng sủi bọt khí.\nBước 4: Chia dung dịch thu được ở ống nghiệm (1) thành hai phần gần bằng nhau, cho vào ống nghiệm (2) và (3).\nBước 5: Cho 0,5 mL dung dịch CuSO~4~ 5% vào ống nghiệm chứa 2 mL dung dịch NaOH 10%, lắc nhẹ.\nBước 6: Đổ dung dịch trong ống nghiệm (2) vào ống nghiệm chứa Cu(OH)~2~ (thu được ở bước 5) và lắc đều.\nBước 7: Đổ dung dịch trong ống nghiệm (3) vào ống nghiệm chứa nước bromine loãng.\nViết phương trình hóa học của các phản ứng xảy ra ở các bước 2, 3, 6 và 7 trong thí nghiệm trên (ghi rõ điều kiện phản ứng nếu có).")
    if is_teacher:
        sol_tl_3 = [
            "*Lời giải chi tiết:*",
            "- *Bước 2 (Thủy phân tinh bột):*",
            "  (C~6~H~10~O~5~)~n~ + n H~2~O -(H^+, t°)-> n C~6~H~12~O~6~ (glucose)",
            "- *Bước 3 (Trung hòa acid HCl dư bằng NaHCO~3~):*",
            "  HCl + NaHCO~3~ -> NaCl + CO~2~↑ + H~2~O",
            "- *Bước 6 (Phản ứng hòa tan Cu(OH)~2~ tạo phức xanh lam):*",
            "  2 C~6~H~12~O~6~ + Cu(OH)~2~ -> (C~6~H~11~O~6~)~2~Cu + 2 H~2~O",
            "- *Bước 7 (Phản ứng oxi hóa glucose bằng nước bromine):*",
            "  CH~2~OH[CHOH]~4~CHO + Br~2~ + H~2~O -> CH~2~OH[CHOH]~4~COOH + 2 HBr",
            "  (hoặc viết: C~6~H~12~O~6~ + Br~2~ + H~2~O -> C~6~H~12~O~7~ + 2 HBr)"
        ]
        add_solution(doc, sol_tl_3)
    else:
        add_student_writing_lines(doc, num_lines=10)

    # Kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(12)
    p_end.paragraph_format.space_after = Pt(4)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "*--- HẾT ---*", base_size=11, base_bold=True)

    p_note = doc.add_paragraph()
    p_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_note, "*Thí sinh không được sử dụng tài liệu; giám thị không giải thích gì thêm.*", base_size=9.5, base_italic=True)

    filename = "De_Kiem_Tra_GK1_Hoa_12_NguyenHue_Ma1202_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_GK1_Hoa_12_NguyenHue_Ma1202_HocSinh.docx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    doc.save(filepath)
    print(f"Đã lưu thành công: {filepath}")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print("--- Đang tạo Đề Kiểm Tra Hóa 12 - THPT Nguyễn Huệ - Mã đề 1202 ---")
    build_exam_document_nh1202(is_teacher=False)
    build_exam_document_nh1202(is_teacher=True)
    print("--- Hoàn tất tạo 2 file Word! ---")


if __name__ == "__main__":
    main()
