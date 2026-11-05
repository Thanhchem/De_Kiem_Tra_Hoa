# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì 1 Hóa Học 12 - THPT Chuyên Khoa Học Huế (2026 - 2027) - Mã Đề 102
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản Word:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_12_ChuyenKHHue_Ma102.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_12_ChuyenKHHue_Ma102_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
IMG_DIR = os.path.join(OUTPUT_DIR, "images_chuyen_khoa_hoc_hue_102")

def add_quick_answer_table_102(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu Mã đề 102.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 102)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (12 CÂU)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - mỗi câu đúng 0,25 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    col_widths_1 = [2.4] + [1.3] * 12
    p1_answers = ["A", "C", "B", "B", "A", "A", "B", "B", "D", "A", "A", "A"]

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
    col_widths_2 = [2.4, 3.2, 3.1, 3.1, 3.1, 3.1]
    
    p2_data = [
        ["Câu", "Chủ đề", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Ester mùi táo tây", "ĐÚNG", "SAI", "SAI", "SAI"],
        ["Câu 2", "Cellulose & Ethanol", "ĐÚNG", "SAI", "SAI", "ĐÚNG"]
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
    render_rich_text(p_p3, "*3. Phần III: Trắc nghiệm Trả lời ngắn (2,0 điểm - mỗi câu đúng 0,5 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_3 = [2.4, 3.9, 3.9, 3.9, 3.9]
    p3_answers = ["2806", "12,1", "322", "13 (hoặc 12,5)"]

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
                txt = "Đáp số" if col_idx == 0 else p3_answers[col_idx - 1]
                if col_idx == 0:
                    render_rich_text(p, f"*{txt}*", base_size=10, base_bold=True)
                else:
                    render_rich_text(p, f"*{txt}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)

    # 4. PHẦN B: TỰ LUẬN (3 CÂU)
    p_p4 = doc.add_paragraph()
    p_p4.paragraph_format.space_before = Pt(6)
    p_p4.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p4, "*4. Phần B: Tự luận (3,0 điểm - mỗi câu 1,0 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl4 = doc.add_table(rows=2, cols=4)
    tbl4.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_4 = [2.4, 5.2, 5.2, 5.2]
    p4_answers = ["585 tỉ cây xanh", "4 đồng phân ester C~4~H~8~O~2~", "9,03 tấn mía tươi"]

    for row_idx, row in enumerate(tbl4.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
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
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if row_idx == 0:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8F5E9"/>')
                tcPr.append(shd)
                txt = "Câu" if col_idx == 0 else f"Câu {col_idx}"
                render_rich_text(p, f"*{txt}*", base_size=10, base_bold=True)
            else:
                txt = "Kết quả" if col_idx == 0 else p4_answers[col_idx - 1]
                if col_idx == 0:
                    render_rich_text(p, f"*{txt}*", base_size=10, base_bold=True)
                else:
                    render_rich_text(p, f"*{txt}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)

    p_sep = doc.add_paragraph()
    p_sep.paragraph_format.space_before = Pt(6)
    p_sep.paragraph_format.space_after = Pt(10)
    add_p_border_bottom(p_sep, color="CCCCCC", sz="6")


def build_word_document(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="TRƯỜNG ĐẠI HỌC KHOA HỌC\nTRƯỜNG THPT CHUYÊN KHOA HỌC HUẾ",
        exam_title="ĐỀ KIỂM TRA GIỮA HỌC KỲ I - NĂM HỌC 2026-2027",
        subject_line="Môn: Hóa học – Lớp 12 (CHUNG)",
        num_pages="04",
        duration="45 phút (không kể thời gian giao đề)",
        exam_code="102",
        atomic_masses="H = 1; C = 12; O = 16; Na = 23; Ag = 108",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_102(doc)

    # ==========================================
    # A. TRẮC NGHIỆM (7,0 ĐIỂM)
    # ==========================================
    p_sec_a = doc.add_paragraph()
    p_sec_a.paragraph_format.space_before = Pt(8)
    p_sec_a.paragraph_format.space_after = Pt(3)
    render_rich_text(p_sec_a, "*A. TRẮC NGHIỆM (7,0 điểm)*", base_size=11.5, base_bold=True, base_color=COLOR_PRIMARY)

    # ------------------------------------------
    # PHẦN I (3,0 ĐIỂM)
    # ------------------------------------------
    add_section_header(
        doc,
        "PHẦN I (3,0 điểm). Câu trắc nghiệm nhiều phương án lựa chọn.",
        subtitle="Học sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi chỉ chọn một phương án."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Maltose là một loại disaccharide có nhiều trong mạch nha. Công thức phân tử của maltose là:")
    opts1 = [
        "*A. C~12~H~22~O~11~.*" if is_teacher else "A. C~12~H~22~O~11~.",
        "B. C~6~H~12~O~6~.",
        "C. (C~6~H~10~O~5~)~n~",
        "D. C~3~H~6~O~2~."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Maltose là một loại disaccharide được tạo nên từ hai gốc glucose liên kết với nhau qua liên kết alpha-1,4-glycosidic.\nCông thức phân tử của maltose là C~12~H~22~O~11~. Chất này có nhiều trong mầm lúa mạch, mạch nha.\n➔ Đáp án đúng: A.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Tiến hành thí nghiệm glucose bị oxi hóa bởi thuốc thử Tollens theo các bước sau:\n"
                              "Bước 1: Cho khoảng 2 mL dung dịch AgNO~3~ 1% vào ống nghiệm sạch.\n"
                              "Bước 2: Thêm từ từ từng giọt dung dịch ammonia 5%, lắc đều cho đến khi kết tủa tan hết. Dung dịch thu được là thuốc thử Tollens.\n"
                              "Bước 3: Thêm tiếp khoảng 2 mL dung dịch glucose 2% lắc đều. Sau đó, ngâm ống nghiệm vào cốc thuỷ tinh chứa nước nóng trong vài phút.\n"
                              "Phát biểu nào sau đây sai?")
    opts2 = [
        "A. Sản phẩm hữu cơ thu được sau bước 3 là ammonium gluconate.",
        "B. Sau bước 3, có lớp bạc (silver) kim loại bám trên thành ống nghiệm.",
        "*C. Thí nghiệm trên chứng minh glucose có tính chất của polyalcohol.*" if is_teacher else "C. Thí nghiệm trên chứng minh glucose có tính chất của polyalcohol.",
        "D. Trong phản ứng ở bước 3, glucose đóng vai trò là chất khử."
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Thí nghiệm tráng bạc (thuốc thử Tollens) chứng minh glucose chứa nhóm chức aldehyde (-CHO) có tính khử.\nTính chất của polyalcohol được chứng minh qua phản ứng hòa tan Cu(OH)~2~ ở nhiệt độ thường tạo phức dung dịch màu xanh lam.\nDo đó, phát biểu C sai.\n➔ Đáp án đúng: C.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Ethyl acetate được viết tắt là EtOAc, được sản xuất ở quy mô khá lớn để làm dung môi, là một loại hoá chất được sử dụng rất nhiều trong nước hoa, trong các loại sơn móng tay, hóa chất này rất phổ biến trong đời sống sinh hoạt hằng ngày. Ethyl acetate có công thức thu gọn nào sau đây?")
    opts3 = [
        "A. CH~3~COOC~6~H~5~",
        "*B. CH~3~COOC~2~H~5~*" if is_teacher else "B. CH~3~COOC~2~H~5~",
        "C. CH~3~COOCH~2~CH~2~CH~3~",
        "D. CH~3~COOCH=CH~2~"
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Ethyl acetate là ester tạo bởi acetic acid (CH~3~COOH) và ethyl alcohol (C~2~H~5~OH).\nCông thức thu gọn của ethyl acetate là CH~3~COOC~2~H~5~.\n➔ Đáp án đúng: B.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Dầu mỡ để lâu bị ôi là do")
    opts4 = [
        "A. chất béo bị rữa ra.",
        "*B. liên kết đôi C=C trong chất béo bị oxi hóa chậm bởi oxygen không khí tạo chất có mùi khó chịu.*" if is_teacher else "B. liên kết đôi C=C trong chất béo bị oxi hóa chậm bởi oxygen không khí tạo chất có mùi khó chịu.",
        "C. chất béo bị oxygen hoá chậm bởi oxygen không khí.",
        "D. chất béo bị thủy phân với nước trong không khí."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Dầu mỡ để lâu trong không khí bị ôi do các gốc acid béo không bão hòa (chứa liên kết đôi C=C) bị oxi hóa chậm bởi oxygen không khí tạo thành hydroperoxide, sau đó bị phân hủy thành các aldehyde và ketone có mùi khó chịu và độc hại.\n➔ Đáp án đúng: B.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 5. Saccharose tham gia phản ứng nào sau đây?")
    opts5 = [
        "*A. Phản ứng với Cu(OH)~2~ tạo dung dịch màu xanh lam.*" if is_teacher else "A. Phản ứng với Cu(OH)~2~ tạo dung dịch màu xanh lam.",
        "B. Phản ứng với thuốc thử Tollens.",
        "C. Phản ứng với nước bromine.",
        "D. Phản ứng thuỷ phân trong môi trường kiềm."
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tử saccharose có nhiều nhóm -OH liền kề (polyalcohol) nên có khả năng hòa tan Cu(OH)~2~ ở nhiệt độ thường tạo phức màu xanh lam.\nSaccharose không có nhóm -OH hemiacetal/hemiketal tự do nên không mở vòng để có nhóm aldehyde, do đó không tham gia phản ứng tráng bạc và không làm mất màu nước bromine. Saccharose chỉ bị thủy phân trong môi trường acid hoặc xúc tác enzyme.\n➔ Đáp án đúng: A.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 6. Thủy phân hoàn toàn 34,2 g saccharose trong môi trường acid. Khối lượng tổng glucose và fructose tạo thành (coi hiệu suất 100%) là")
    opts6 = [
        "*A. 36,0 g*" if is_teacher else "A. 36,0 g",
        "B. 18,0 g",
        "C. 34,2 g",
        "D. 68,4 g"
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phương trình phản ứng thủy phân:\nC~12~H~22~O~11~ + H~2~O -(H^+, t^o)-> C~6~H~12~O~6~ (glucose) + C~6~H~12~O~6~ (fructose)\nSố mol saccharose: n = 34,2 / 342 = 0,1 mol.\nSố mol H~2~O phản ứng: n_H2O = 0,1 mol => m_H2O = 0,1 x 18 = 1,8 g.\nTheo định luật bảo toàn khối lượng:\nm_(sản phẩm) = m_saccharose + m_H2O = 34,2 + 1,8 = 36,0 gam.\n➔ Đáp án đúng: A.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 7. Khi cho xà phòng vào nước cứng chứa ion Ca^2+^ và Mg^2+^, hiện tượng nào thường xảy ra?")
    opts7 = [
        "A. Nước cứng chuyển thành nước mềm ngay lập tức.",
        "*B. Xà phòng tạo kết tủa với các ion Ca^2+^, Mg^2+^.*" if is_teacher else "B. Xà phòng tạo kết tủa với các ion Ca^2+^, Mg^2+^.",
        "C. Xà phòng bị oxi hóa thành chất béo.",
        "D. Xà phòng phân hủy hoàn toàn thành glycerol."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Các muối sodium/potassium của acid béo trong xà phòng tác dụng với ion Ca^2+^ và Mg^2+^ trong nước cứng tạo thành các muối calcium/magnesium stearate, palmitate kết tủa không tan:\n2RCOONa + Ca^2+^ -> (RCOO)~2~Ca↓ + 2Na^+^\nHiện tượng này làm kết tủa lắng xuống, làm giảm bọt xà phòng và giảm hiệu quả giặt rửa.\n➔ Đáp án đúng: B.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 8. Chất nào sau đây là chất béo bão hòa (ở trạng thái rắn ở nhiệt độ phòng)?")
    opts8 = [
        "A. Triolein.",
        "*B. Tristearin.*" if is_teacher else "B. Tristearin.",
        "C. Trilinolein.",
        "D. Trioleate."
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Tristearin có công thức (C~17~H~35~COO)~3~C~3~H~5~, là chất béo bão hòa (no), có nhiệt độ nóng chảy cao nên ở trạng thái rắn ở nhiệt độ phòng.\nTriolein và trilinolein là các triglyceride không no (không bão hòa), ở trạng thái lỏng ở nhiệt độ phòng.\n➔ Đáp án đúng: B.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 9. Đun nóng 8,8 g ethyl acetate với 100 mL NaOH 1,00 M, phản ứng hoàn toàn. Khối lượng muối tạo thành là")
    opts9 = [
        "A. 4,1 g",
        "B. 10,0 g",
        "C. 16,4 g",
        "*D. 8,2 g*" if is_teacher else "D. 8,2 g"
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Số mol ethyl acetate: n_ester = 8,8 / 88 = 0,1 mol.\nSố mol NaOH: n_NaOH = 0,100 x 1,00 = 0,1 mol.\nPhương trình phản ứng: CH~3~COOC~2~H~5~ + NaOH -> CH~3~COONa + C~2~H~5~OH\nPhản ứng xảy ra vừa đủ, muối tạo thành là CH~3~COONa (0,1 mol).\nKhối lượng muối: m = 0,1 x 82 = 8,2 gam.\n➔ Đáp án đúng: D.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 10. Hợp chất nào sau đây được sử dụng làm chất giặt rửa tổng hợp?")
    opts10 = [
        "*A. CH~3~[CH~2~]~11~OSO~3~Na.*" if is_teacher else "A. CH~3~[CH~2~]~11~OSO~3~Na.",
        "B. CH~3~COONa.",
        "C. CH~3~[CH~2~]~12~COOCH~3~.",
        "D. CH~3~[CH~2~]~12~COONa."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "CH~3~[CH~2~]~11~OSO~3~Na (Sodium lauryl sulfate / Sodium dodecyl sulfate) là chất hoạt động bề mặt tổng hợp phổ biến, được sử dụng làm chất giặt rửa tổng hợp.\nCH~3~[CH~2~]~12~COONa là thành phần của xà phòng thông thường.\n➔ Đáp án đúng: A.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 11. Tinh thể chất rắn X không màu, vị ngọt, dễ tan trong nước. X có nhiều trong cây mía, củ cải đường và hoa thốt nốt. Trong công nghiệp, X được chuyển hóa thành chất Y dùng để tráng gương, tráng ruột phích. Tên gọi của X và Y lần lượt là")
    opts11 = [
        "*A. Saccharose và glucose.*" if is_teacher else "A. Saccharose và glucose.",
        "B. Glucose và saccharose.",
        "C. Saccharose và cellulose.",
        "D. Glucose và fructose."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Chất X có nhiều trong mía, củ cải đường, hoa thốt nốt là saccharose.\nTrong công nghiệp, saccharose được thủy phân tạo thành glucose (và fructose), sau đó dịch thủy phân chứa glucose được dùng làm chất khử để tráng gương, tráng ruột phích. Vậy X là saccharose và Y là glucose.\n➔ Đáp án đúng: A.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 12. Trong thí nghiệm điều chế ethyl acetate trong phòng thí nghiệm (cho ethanol, acetic acid và H~2~SO~4~ đặc vào ống nghiệm rồi đun nhẹ), vai trò của H~2~SO~4~ đặc là:")
    opts12 = [
        "*A. chất xúc tác và hút nước để chuyển dịch cân bằng.*" if is_teacher else "A. chất xúc tác và hút nước để chuyển dịch cân bằng.",
        "B. chất môi trường kiềm để hòa tan ester.",
        "C. chất chỉ thị màu để nhận biết ester.",
        "D. chất oxi hóa alcohol thành aldehyde."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phản ứng ester hóa là phản ứng thuận nghịch: CH~3~COOH + C~2~H~5~OH <=(H~2~SO~4~, t^o)=> CH~3~COOC~2~H~5~ + H~2~O\nH~2~SO~4~ đặc đóng vai trò vừa là chất xúc tác làm tăng tốc độ phản ứng, vừa hút nước làm giảm nồng độ H~2~O, giúp cân bằng chuyển dịch theo chiều thuận làm tăng hiệu suất tạo ester.\n➔ Đáp án đúng: A.")

    # ------------------------------------------
    # PHẦN II (2,0 ĐIỂM)
    # ------------------------------------------
    add_section_header(
        doc,
        "PHẦN II (2,0 điểm). Câu trắc nghiệm đúng sai.",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn Đúng hoặc Sai."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Một ester no, đơn chức, mạch hở X có nguồn gốc từ quả táo tây có mùi thơm dễ chịu. Kết quả phân tích nguyên tố cho thấy X chứa 54,55 % carbon, 9,09 % hydrogen và 36,36 % oxygen về khối lượng. Từ phổ khối lượng (MS), xác định được tín hiệu ion phân tử [M^+^] có m/z = 88. Cho 8,8 gam X tác dụng vừa đủ với 100 mL dung dịch NaOH 1 M, thu được muối Y và alcohol Z. Cho các nhận định sau:")
    stmts_c1 = [
        ("a", "Alcohol Z sinh ra là ethanol (C~2~H~5~OH), có nhiệt độ sôi cao hơn ester X.", True),
        ("b", "Công thức phân tử của ester X là C~4~H~6~O~2~.", False),
        ("c", "Tên gọi của X là methyl acetate.", False),
        ("d", "Đun nóng alcohol Z với H~2~SO~4~ đặc ở 140 °C thu được alkene làm mất màu dung dịch nước bromine.", False)
    ]
    for lbl, text, ans in stmts_c1:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans else "Sai"
        render_statement(p_stmt, f"{lbl}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "- Xác định CTPT của X:\n  n_C : n_H : n_O = (54,55/12) : (9,09/1) : (36,36/16) = 4,546 : 9,09 : 2,2725 = 2 : 4 : 1 => Công thức thực nghiệm (C~2~H~4~O)~n~.\n  Tín hiệu [M^+^] có m/z = 88 => M_X = 88 => 44n = 88 => n = 2. CTPT của X là C~4~H~8~O~2~.\n- Số mol X: n_X = 8,8 / 88 = 0,1 mol. n_NaOH = 0,100 x 1 = 0,1 mol. X phản ứng vừa đủ với NaOH cho muối Y và alcohol Z.\n  Ester X có mùi thơm quả táo tây là ethyl acetate (CH~3~COOC~2~H~5~). Muối Y là CH~3~COONa và alcohol Z là ethanol (C~2~H~5~OH).\n- Đánh giá các mệnh đề:\n  + a) ĐÚNG. Alcohol Z là C~2~H~5~OH. Phân tử C~2~H~5~OH có liên kết hydrogen liên phân tử nên có nhiệt độ sôi (78,3 °C) cao hơn ester X (77,1 °C).\n  + b) SAI. Công thức phân tử của X là C~4~H~8~O~2~, không phải C~4~H~6~O~2~.\n  + c) SAI. Tên gọi của X là ethyl acetate, không phải methyl acetate (C~3~H~6~O~2~).\n  + d) SAI. Đun nóng ethanol với H~2~SO~4~ đặc ở 140 °C thu được diethyl ether (C~2~H~5~OC~2~H~5~), ở 170 °C mới thu được alkene (ethylene C~2~H~4~).\n➔ Kết luận: a - ĐÚNG, b - SAI, c - SAI, d - SAI.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Ngô và mía là hai nguyên liệu chính được sử dụng trong sản xuất ethanol. Tuy nhiên, chúng là những loại cây lương thực quan trọng. Trong khi đó, cellulose – một nguồn nguyên liệu dồi dào và dễ tìm – cũng có thể dùng để sản xuất ethanol. Tuy giá thành sản xuất ethanol từ cellulose còn cao, do quá trình nuôi cấy nấm sản sinh enzyme cellulase để thuỷ phân cellulose thành glucose tốn kém nhiều năng lượng, nhưng đây vẫn là một hướng đi đầy hứa hẹn trong tương lai.")
    
    # Chèn ảnh cấu trúc cellulose
    p_img_cell = doc.add_paragraph()
    p_img_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img_cell.paragraph_format.space_before = Pt(3)
    p_img_cell.paragraph_format.space_after = Pt(1)
    cell_img_path = os.path.join(IMG_DIR, "cellulose_clean.png")
    if os.path.exists(cell_img_path):
        p_img_cell.add_run().add_picture(cell_img_path, width=Cm(6.5))
    p_cap_cell = doc.add_paragraph()
    p_cap_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_cell.paragraph_format.space_before = Pt(1)
    p_cap_cell.paragraph_format.space_after = Pt(3)
    render_rich_text(p_cap_cell, "_Cấu trúc cellulose_", base_size=9.5, base_italic=True)

    stmts_c2 = [
        ("a", "Phân tử cellulose được tạo thành từ nhiều đơn vị beta-glucose, nối với nhau qua liên kết beta-1,4-glycoside.", True),
        ("b", "Với những tiến bộ công nghệ đạt được, người ta tin rằng 0,81 tấn cellulose có thể sản xuất được 345 (L) ethanol. Khi đó, hiệu suất của quá trình đạt 70 % (cho biết khối lượng riêng của ethanol là 0,8 g/mL).", False),
        ("c", "Cellulose có nhiều trong trái cây chín.", False),
        ("d", "Cellulose là một polymer mạch không phân nhánh.", True)
    ]
    for lbl, text, ans in stmts_c2:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans else "Sai"
        render_statement(p_stmt, f"{lbl}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "- Đánh giá các mệnh đề:\n  + a) ĐÚNG. Cellulose là polysaccharide gồm nhiều đơn vị beta-glucose liên kết với nhau bằng liên kết beta-1,4-glycosidic.\n  + b) SAI. Sơ đồ chuyển hóa: (C~6~H~10~O~5~)~n~ -> n C~6~H~12~O~6~ -> 2n C~2~H~5~OH.\n    Cứ 162 gam cellulose theo lý thuyết sinh ra 92 gam C~2~H~5~OH.\n    Từ 0,81 tấn = 810 kg cellulose => m_(ethanol lý thuyết) = (810 x 92) / 162 = 460 kg.\n    V_(ethanol lý thuyết) = 460 kg / (0,8 kg/L) = 575 L.\n    Hiệu suất thực tế: H = (345 / 575) x 100% = 60%, không phải 70%.\n  + c) SAI. Trái cây chín chứa nhiều glucose và fructose; cellulose là thành phần cấu tạo nên thành tế bào thực vật trong gỗ, bông, nứa...\n  + d) ĐÚNG. Phân tử cellulose có mạch kéo dài, không phân nhánh.\n➔ Kết luận: a - ĐÚNG, b - SAI, c - SAI, d - ĐÚNG.")

    # ------------------------------------------
    # PHẦN III (2,0 ĐIỂM)
    # ------------------------------------------
    add_section_header(
        doc,
        "PHẦN III (2,0 điểm). Câu trắc nghiệm trả lời ngắn.",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4. Viết đáp số tương ứng vào bài làm."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Acid béo omega-3 và omega-6 là các acid béo không no với liên kết đôi đầu tiên ở vị trí số 3 và 6 khi đánh số từ nhóm methyl. Acid béo omega-3 và omega-6 đều có vai trò quan trọng đối với cơ thể, giúp phòng ngừa nhiều loại bệnh. Linoleic acid có công thức cấu tạo như sau:")
    p_img_lin = doc.add_paragraph()
    p_img_lin.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img_lin.paragraph_format.space_before = Pt(3)
    p_img_lin.paragraph_format.space_after = Pt(2)
    lin_img_path = os.path.join(IMG_DIR, "linoleic_clean.png")
    if os.path.exists(lin_img_path):
        p_img_lin.add_run().add_picture(lin_img_path, width=Cm(9.0))
    p_sub_c1 = doc.add_paragraph()
    p_sub_c1.paragraph_format.space_before = Pt(2)
    p_sub_c1.paragraph_format.space_after = Pt(3)
    render_rich_text(p_sub_c1, "Linoleic acid thuộc loại acid béo omega-x và có phân tử khối là y. Số có giá trị 10y+x là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Linoleic acid có công thức thu gọn: CH~3~-(CH~2~)~4~-CH=CH-CH~2~-CH=CH-(CH~2~)~7~-COOH.\nĐánh số các nguyên tử carbon từ đầu nhóm methyl (-CH~3~ là vị trí số 1):\nLiên kết đôi đầu tiên xuất hiện ở giữa C số 6 và C số 7 => linoleic acid là acid béo omega-6 => x = 6.\nCông thức phân tử của linoleic acid là C~18~H~32~O~2~ (hoặc C~17~H~31~COOH).\nPhân tử khối: y = 18 x 12 + 32 x 1 + 2 x 16 = 280.\nGiá trị của biểu thức: 10y + x = 10 x 280 + 6 = 2806.\n➔ Đáp số: 2806.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Methyl salicylate dùng làm thuốc xoa bóp giảm đau, được điều chế theo phản ứng sau:\n"
                              "HOC~6~H~4~COOH + CH~3~OH <=(H~2~SO~4~, t^o)=> HOC~6~H~4~COOCH~3~ + H~2~O\n"
                              "Để sản xuất 3,8 triệu tuýp thuốc xoa bóp giảm đau cần tối thiểu m tấn axit salixylic. Biết mỗi tuýp thuốc chứa 2,8 gam methyl salicylate và hiệu suất phản ứng tính theo axit salixylic là 80 %. Tính giá trị của m ? (Kết quả làm tròn đến con số thứ nhất sau dấu phẩy).")
    p_img_metho = doc.add_paragraph()
    p_img_metho.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img_metho.paragraph_format.space_before = Pt(2)
    p_img_metho.paragraph_format.space_after = Pt(2)
    metho_img_path = os.path.join(IMG_DIR, "methocylat.png")
    if os.path.exists(metho_img_path):
        p_img_metho.add_run().add_picture(metho_img_path, width=Cm(4.5))
    if is_teacher:
        add_solution_no_table(doc, "Khối lượng methyl salicylate cần dùng cho 3,8 triệu tuýp:\nm_este = 3,8 x 10^6 x 2,8 g = 10.640.000 gam = 10,64 tấn.\nSố mol methyl salicylate (C~8~H~8~O~3~, M = 152): n_este = 10.640.000 / 152 = 70.000 mol = 70 kmol.\nAxit salixylic (HOC~6~H~4~COOH = C~7~H~6~O~3~) có M = 138 g/mol.\nVới hiệu suất phản ứng H = 80%, số mol axit salixylic tối thiểu cần dùng là:\nn_acid = 70.000 / 0,8 = 87.500 mol = 87,5 kmol.\nKhối lượng axit salixylic tối thiểu m:\nm = 87,5 kmol x 138 kg/kmol = 12.075 kg = 12,075 tấn.\nLàm tròn đến chữ số thập phân thứ nhất: m = 12,1 tấn.\n➔ Đáp số: 12,1.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Đun nóng một loại mỡ động vật với dung dịch KOH, sản phẩm thu được có chứa muối potassium stearate (C~17~H~35~COOK). Phân tử khối của potassium stearate là bao nhiêu ?")
    if is_teacher:
        add_solution_no_table(doc, "Potassium stearate có công thức C~17~H~35~COOK.\nPhân tử khối: M = 17 x 12 + 35 x 1 + 12 + 16 x 2 + 39 = 204 + 35 + 12 + 32 + 39 = 322.\n➔ Đáp số: 322.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Một xưởng sản xuất gương phẳng cần tráng bạc lên 15 000 m^2^ bề mặt kính. Biết khối lượng bạc cần phủ trên mỗi mét vuông bề mặt kính là 0,8 gam. Lớp bạc được tạo thành bằng phản ứng tráng bạc của glucose với thuốc thử Tollens (được điều chế từ dung dịch AgNO~3~ và NH~3~). Giả sử AgNO~3~ được sử dụng dư và hiệu suất tạo bạc bám trên bề mặt kính tính theo glucose đạt 80 %. Tính khối lượng glucose tối thiểu (theo kg) cần dùng để sản xuất số gương trên (làm tròn đến hàng đơn vị).")
    if is_teacher:
        add_solution_no_table(doc, "Khối lượng bạc cần phủ lên 15 000 m^2^ gương:\nm_Ag = 15 000 x 0,8 g = 12 000 gam = 12 kg.\nSố mol Ag: n_Ag = 12 000 / 108 = 1000 / 9 mol.\nPhương trình phản ứng tráng bạc của glucose:\nCH~2~OH[CHOH]~4~CHO + 2[Ag(NH~3~)~2~]OH -> CH~2~OH[CHOH]~4~COONH~4~ + 2Ag + 3NH~3~ + H~2~O\nTheo lý thuyết: n_glucose = 1/2 n_Ag = 500 / 9 mol.\nVì hiệu suất bám bạc đạt 80%, số mol glucose thực tế tối thiểu cần dùng là:\nn_glucose (thực tế) = (500 / 9) / 0,8 = 625 / 9 mol.\nKhối lượng glucose cần dùng:\nm_glucose = (625 / 9) x 180 = 625 x 20 = 12 500 gam = 12,5 kg.\nLàm tròn đến hàng đơn vị theo yêu cầu: 13 kg (hoặc 12,5 kg).\n➔ Đáp số: 13 (hoặc 12,5).")

    # ==========================================
    # B. TỰ LUẬN (3,0 ĐIỂM)
    # ==========================================
    p_sec_b = doc.add_paragraph()
    p_sec_b.paragraph_format.space_before = Pt(8)
    p_sec_b.paragraph_format.space_after = Pt(3)
    render_rich_text(p_sec_b, "*B. TỰ LUẬN (3,0 điểm)*", base_size=11.5, base_bold=True, base_color=COLOR_PRIMARY)

    # Câu 1 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. (1,0 điểm)\n"
                              "Cây xanh nhả khí oxygen sau khi hoàn tất quá trình quang hợp, con người và nhiều loài động vật khác cần khí oxygen để duy trì sự sống. Phản ứng quang hợp được mô tả theo phương trình sau:\n"
                              "nCO~2~ + mH~2~O -(ánh sáng / chất diệp lục)-> C~n~(H~2~O)~m~ + nO~2~\n"
                              "Giả sử một cây xanh trưởng thành tiêu thụ lượng CO~2~ trung bình khoảng 22 kg/năm để thực hiện quá trình quang hợp và trung bình một năm mỗi người cần tối thiểu 1,2 tấn oxygen để duy trì sự sống. Để cung cấp đủ lượng oxygen cho 7,8 tỉ người trên Trái đất trong một năm thì cần bao nhiêu tỉ cây xanh như trên?")
    if is_teacher:
        add_solution_no_table(doc, "1. Theo phương trình quang hợp: nCO~2~ + mH~2~O -> C~n~(H~2~O)~m~ + nO~2~\nTỉ lệ mol giữa CO~2~ tiêu thụ và O~2~ giải phóng là 1 : 1 (n_O2 = n_CO2).\n2. Khối lượng CO~2~ một cây xanh tiêu thụ trong một năm là 22 kg.\nSố mol CO~2~ cây tiêu thụ: n_CO2 = 22 / 44 = 0,5 kmol = 500 mol.\nSố mol O~2~ một cây giải phóng trong một năm: n_O2 = 0,5 kmol.\nKhối lượng O~2~ một cây giải phóng trong một năm: m_O2 = 0,5 x 32 = 16 kg/cây/năm.\n3. Lượng oxygen cần thiết cho 7,8 tỉ người trong một năm:\nm_(tổng O2) = 7,8 x 10^9 người x 1,2 tấn/người = 9,36 x 10^9 tấn = 9,36 x 10^12 kg.\n4. Số lượng cây xanh trưởng thành cần thiết:\nSố cây = (9,36 x 10^12 kg) / (16 kg/cây) = 5,85 x 10^11 cây = 585 tỉ cây xanh.\n➔ Đáp số: Cần khoảng 585 tỉ cây xanh.")
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 2 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. (1,0 điểm)\n"
                              "Viết công thức cấu tạo và gọi tên các đồng phân ester có công thức phân tử C~4~H~8~O~2~.")
    if is_teacher:
        add_solution_no_table(doc, "Ester no, đơn chức, mạch hở C~4~H~8~O~2~ có dạng RCOOR' (với R >= 0; R' >= 1, tổng số nguyên tử carbon bằng 4).\nCó tất cả 4 đồng phân ester:\n1) HCOOCH~2~CH~2~CH~3~ : propyl formate (hoặc propyl methanoate)\n2) HCOOCH(CH~3~)~2~ : isopropyl formate (hoặc 1-methylethyl methanoate)\n3) CH~3~COOCH~2~CH~3~ : ethyl acetate (hoặc ethyl ethanoate)\n4) CH~3~CH~2~COOCH~3~ : methyl propionate (hoặc methyl propanoate)")
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Câu 3 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. (1,0 điểm)\n"
                              "Một nhà máy đường sản xuất saccharose tinh khiết từ cây mía tươi. Biết rằng trong mía tươi chứa 13,5 % saccharose về khối lượng và hiệu suất thu hồi saccharose của toàn bộ quy trình sản xuất đạt 82 %. Tính khối lượng mía tươi (theo tấn) tối thiểu cần cung cấp cho nhà máy để sản xuất được 1,0 tấn đường saccharose tinh khiết (làm tròn đến hàng phần trăm).")
    if is_teacher:
        add_solution_no_table(doc, "1. Lượng saccharose có trong lượng mía tươi cần dùng với hiệu suất thu hồi 82% là:\nm_saccharose (trong mía) = 1,0 / 82% = 1,0 / 0,82 ≈ 1,2195 tấn.\n2. Khối lượng mía tươi tối thiểu cần dùng (biết trong mía chứa 13,5% saccharose):\nm_mía = 1,2195 / 13,5% = 1,0 / (0,82 x 0,135) = 1,0 / 0,1107 ≈ 9,0334 tấn.\n3. Làm tròn đến hàng phần trăm theo yêu cầu của bài toán:\nm_mía ≈ 9,03 tấn.\n➔ Đáp số: 9,03 tấn mía tươi.")
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Dòng kết thúc đề thi
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(12)
    p_end.paragraph_format.space_after = Pt(3)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "-------------------- HẾT --------------------", base_size=10, base_bold=True)

    p_note1 = doc.add_paragraph()
    p_note1.paragraph_format.space_before = Pt(1)
    p_note1.paragraph_format.space_after = Pt(1)
    p_note1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_note1, "_(Thí sinh không được sử dụng tài liệu - Giám thị không giải thích gì thêm)_", base_size=9.5, base_italic=True, base_color=COLOR_GRAY)

    filename = "De_Kiem_Tra_GK1_Hoa_12_ChuyenKHHue_Ma102_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_GK1_Hoa_12_ChuyenKHHue_Ma102.docx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    doc.save(filepath)
    print(f"Đã tạo file Word: {filepath}")
    return filepath

if __name__ == "__main__":
    build_word_document(is_teacher=False)
    build_word_document(is_teacher=True)
