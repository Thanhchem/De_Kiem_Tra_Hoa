# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kỳ 1 Hóa Học 11 - THPT Chuyên Khoa Học Huế - Mã Đề 135
Năm học 2026 - 2027
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản Word:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_11_ChuyenKhoaHocHue_Ma135_HocSinh.docx (và De_Kiem_Tra_GK1_Hoa_11_ChuyenKhoaHocHue_Ma135.docx)
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_11_ChuyenKhoaHocHue_Ma135_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết từng câu)
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


def add_quick_answer_table_ckh135(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu MÃ ĐỀ 135.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 135)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (12 CÂU - 3,0 ĐIỂM)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - mỗi câu đúng 0,25 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_1 = [2.4] + [1.3] * 12
    p1_answers = ["C", "A", "C", "A", "C", "C", "B", "D", "A", "D", "C", "A"]

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

    # 2. PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (2 CÂU - 2,0 ĐIỂM)
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(6)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm - 1 ý: 0,1đ; 2 ý: 0,25đ; 3 ý: 0,5đ; 4 ý: 1,0đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_2 = [2.4, 3.6, 3.0, 3.0, 3.0, 3.0]
    
    p2_data = [
        ["Câu", "Nội dung", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Khí thải NH~3~ & muối ammonium", "ĐÚNG", "SAI", "SAI", "ĐÚNG"],
        ["Câu 2", "Hệ cân bằng CO~2~ + H~2~", "ĐÚNG", "SAI", "SAI", "SAI"]
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

            val = p2_data[row_idx][col_idx]
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

    # 3. PHẦN III: TRẢ LỜI NGẮN (4 CÂU - 2,0 ĐIỂM)
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(6)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*3. Phần III: Trả lời ngắn (2,0 điểm - mỗi câu đúng 0,5 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_3 = [2.4, 3.9, 3.9, 3.9, 3.9]
    p3_answers = ["6,12", "5", "0,75", "2"]

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
        ["Câu 1 (TL)", "1,0 điểm", "Phương trình điện li đúng: KOH, Al~2~(SO~4~)~3~, (NH~4~)~3~PO~4~ mũi tên một chiều (->); CH~3~COOH mũi tên thuận nghịch (<=>)."],
        ["Câu 2 (TL)", "1,0 điểm", "a) 3 PTHH điều chế HNO~3~ từ NH~3~.  b) m(NH~3~ thực tế) = 33 660,03 tấn (hoặc ~33 660 tấn)."],
        ["Câu 3 (TL)", "1,0 điểm", "n(OH^- dư) = 0,032 mol => [OH^- dư] = 0,16 M => pOH ≈ 0,796 => pH = 13,2."]
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


def build_exam_chuyenkhoahoc_135(is_teacher=False, output_path=None):
    doc = docx.Document()
    setup_header_footer(doc)

    # 1. Khối tiêu đề đầu đề thi
    build_title_block(
        doc,
        school_line="TRƯỜNG ĐẠI HỌC KHOA HỌC\nTRƯỜNG THPT CHUYÊN KHOA HỌC HUẾ",
        exam_title="ĐỀ KIỂM TRA GIỮA HỌC KỲ I - NĂM HỌC 2026 - 2027",
        subject_line="Môn: Hóa học – Lớp 11 (CHUNG)",
        num_pages="03",
        duration="45 phút",
        exam_code="135",
        atomic_masses="H=1; Li=7; Be=9; N=14; O=16; Na=23; Mg=24; Al=27; P=31; S=32; Cl=35,5; K=39; Ca=40; Ba=137.",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_ckh135(doc)

    # =========================================================================
    # A. TRẮC NGHIỆM (7 điểm)
    # =========================================================================
    p_sec_a = doc.add_paragraph()
    p_sec_a.paragraph_format.space_before = Pt(6)
    p_sec_a.paragraph_format.space_after = Pt(2)
    render_rich_text(p_sec_a, "*A. TRẮC NGHIỆM (7 điểm)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # PHẦN I
    add_section_header(
        doc,
        "PHẦN I. (3 điểm) Câu trắc nghiệm nhiều phương án lựa chọn.",
        subtitle="Học sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi học sinh chỉ chọn một phương án."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Có thể nhận biết muối ammonium bằng cách cho muối tác dụng với dung dịch kiềm thấy thoát ra một chất khí làm xanh quỳ tím ẩm. Chất khí đó là")
    opts1 = [
        "A. NO~2~.",
        "B. NO.",
        "*C. NH~3~.*" if is_teacher else "C. NH~3~.",
        "D. H~2~."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khi cho muối ammonium tác dụng với dung dịch kiềm (chứa ion OH^-), phản ứng tạo thành khí ammonia (NH~3~) có mùi khai đặc trưng, làm xanh quỳ tím ẩm:\nNH~4~^+^ + OH^-^ -> NH~3~↑ + H~2~O.\n=> Đáp án đúng: C.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Trong các dung dịch có cùng nồng độ 0,10 mol/l dưới đây, dung dịch nào dẫn điện kém nhất?")
    opts2 = [
        "*A. HF*" if is_teacher else "A. HF",
        "B. HI",
        "C. HCl",
        "D. HBr"
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "HI, HCl, HBr là các acid mạnh, trong nước phân li hoàn toàn thành ion (chất điện li mạnh).\nHF là một acid yếu, trong nước chỉ phân li một phần thành ion (chất điện li yếu):\nHF <=> H^+^ + F^-^.\nVì vậy, ở cùng nồng độ 0,10 mol/L, dung dịch HF có nồng độ ion tự do nhỏ nhất nên dẫn điện kém nhất.\n=> Đáp án đúng: A.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Acid trong dạ dày giúp kích hoạt enzyme pepsin, hỗ trợ tiêu hóa protein và bảo vệ cơ thể khỏi vi khuẩn có hại. Nếu trong dạ dày có pH = 2 thì nồng độ H^+^ bằng bao nhiêu?")
    opts3 = [
        "A. 10 M.",
        "B. 10^2^ M.",
        "*C. 10^-2^ M.*" if is_teacher else "C. 10^-2^ M.",
        "D. 2 M."
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo định nghĩa pH:\npH = -log[H^+^]  =>  [H^+^] = 10^-pH^ = 10^-2^ M.\n=> Đáp án đúng: C.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Nhận xét nào dưới đây về nitrogen không đúng?")
    opts4 = [
        "*A. Khí nitrogen duy trì sự cháy.*" if is_teacher else "A. Khí nitrogen duy trì sự cháy.",
        "B. N~2~ phản ứng với O~2~ ở nhiệt độ rất cao.",
        "C. Khí nitrogen rất ít tan trong nước.",
        "D. N~2~ vừa có tính oxi hóa vừa có tính khử."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khí nitrogen (N~2~) không duy trì sự cháy và không duy trì sự hô hấp (sinh vật sẽ chết ngạt trong môi trường khí nitrogen tinh khiết). Vì vậy nhận xét A là không đúng.\n(Các nhận xét B, C, D đều đúng: N~2~ phản ứng với O~2~ ở khoảng 3000°C hoặc tia lửa điện; rất ít tan trong nước; số oxi hóa 0 có thể giảm xuống -3 hoặc tăng lên +1, +2, +4, +5 nên vừa có tính oxi hóa vừa có tính khử).\n=> Đáp án đúng: A.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 5. Phản ứng thuận nghịch là phản ứng")
    opts5 = [
        "A. có phương trình hoá học được biểu diễn bằng mũi tên một chiều.",
        "B. chỉ xảy ra theo một chiều nhất định.",
        "*C. trong cùng điều kiện, phản ứng xảy ra theo hai chiều trái ngược nhau.*" if is_teacher else "C. trong cùng điều kiện, phản ứng xảy ra theo hai chiều trái ngược nhau.",
        "D. xảy ra giữa hai chất khí."
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo định nghĩa SGK Hóa học 11 (Cánh Diều / Kết Nối Tri Thức / Chân Trời Sáng Tạo):\nPhản ứng thuận nghịch là phản ứng trong cùng điều kiện, xảy ra đồng thời theo hai chiều trái ngược nhau (chiều thuận và chiều nghịch).\n=> Đáp án đúng: C.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 6. Các yếu tố ảnh hưởng đến cân bằng hoá học là :")
    opts6 = [
        "A. nồng độ, nhiệt độ và chất xúc tác.",
        "B. nồng độ, áp suất và diện tích bề mặt.",
        "*C. nồng độ, nhiệt độ và áp suất.*" if is_teacher else "C. nồng độ, nhiệt độ và áp suất.",
        "D. áp suất, nhiệt độ và chất xúc tác."
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo nguyên lý chuyển dịch cân bằng Le Chatelier:\nCác yếu tố làm chuyển dịch cân bằng hoá học là: nồng độ, nhiệt độ và áp suất (đối với hệ có chất khí).\nChất xúc tác và diện tích bề mặt tiếp xúc chỉ làm thay đổi tốc độ phản ứng (giúp cân bằng được thiết lập nhanh hơn) mà không làm chuyển dịch cân bằng.\n=> Đáp án đúng: C.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 7. Hành động nào của con người không phải nguyên nhân chính gây nên hiện tượng phú dưỡng?")
    opts7 = [
        "A. Sử dụng dư thừa lượng phân bón chứa ion PO~4~^3-^.",
        "*B. Quá trình đánh bắt cá bằng chích điện và thuốc nổ của con người.*" if is_teacher else "B. Quá trình đánh bắt cá bằng chích điện và thuốc nổ của con người.",
        "C. Sử dụng dư thừa lượng phân bón chứa ion NO~2~^-, NO~3~^-.",
        "D. Nước thải chăn nuôi, nước thải công nghiệp chứa nhiều các nguyên tố dinh dưỡng như N, P."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hiện tượng phú dưỡng (eutrophication) là sự tích tụ lượng lớn các chất dinh dưỡng (đặc biệt là hợp chất chứa nitrogen và phosphorus) trong các nguồn nước mặt (ao, hồ, sông), làm tảo và vi khuẩn phát triển bùng phát.\n- A, C, D là các nguyên nhân chính đưa lượng lớn ion N và P vào nguồn nước.\n- B (đánh bắt cá bằng xung điện, thuốc nổ) là hành vi hủy hoại hệ sinh thái thủy sản, nhưng không phải là nguồn bổ sung dinh dưỡng N, P gây phú dưỡng.\n=> Đáp án đúng: B.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 8. Trong phản ứng sau đây, những chất, ion nào đóng vai trò là acid theo thuyết Bronsted-Lowry?\nH~2~S(aq) + H~2~O <=> HS^-^(aq) + H~3~O^+^(aq)")
    opts8 = [
        "A. H~2~S và HS^-^",
        "B. H~2~O và H~3~O^+^",
        "C. H~2~S và H~2~O",
        "*D. H~2~S và H~3~O^+^*" if is_teacher else "D. H~2~S và H~3~O^+^"
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo thuyết Brønsted – Lowry:\n- Acid là chất cho proton (H^+^).\n- Base là chất nhận proton (H^+^).\nXét phản ứng: H~2~S(aq) + H~2~O <=> HS^-^(aq) + H~3~O^+^(aq)\n- Ở chiều thuận: H~2~S nhường H^+^ cho H~2~O tạo HS^-^ => H~2~S là acid.\n- Ở chiều nghịch: H~3~O^+^ nhường H^+^ cho HS^-^ tạo H~2~O => H~3~O^+^ là acid.\nDo đó, các chất/ion đóng vai trò là acid là H~2~S và H~3~O^+^.\n=> Đáp án đúng: D.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 9. Nitrogen tương đối trơ về mặt hoá học ở nhiệt độ thường là do")
    opts9 = [
        "*A. phân tử N~2~ có liên kết ba với năng lượng liên kết lớn.*" if is_teacher else "A. phân tử N~2~ có liên kết ba với năng lượng liên kết lớn.",
        "B. nitrogen có độ âm điện lớn.",
        "C. phân tử N~2~ có liên kết cộng hoá trị không phân cực.",
        "D. phân tử N~2~ có liên kết ion."
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong phân tử N~2~, hai nguyên tử nitrogen liên kết với nhau bằng một liên kết ba rất bền vững (:N≡N:). Năng lượng liên kết này rất lớn (E_b = 945 kJ/mol), đòi hỏi năng lượng rất cao để phá vỡ nên ở nhiệt độ thường N~2~ tương đối trơ về mặt hóa học.\n=> Đáp án đúng: A.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 10. Khi tăng áp suất, những cân bằng hóa học nào dưới đây chuyển dịch theo chiều thuận?\n(1) H~2~(g) + Cl~2~(g) <=> 2HCl(g)\n(2) NH~4~Cl(s) <=> NH~3~(g) + HCl(g)\n(3) 2H~2~O~2~(aq) <=> 2H~2~O(l) + O~2~(g)\n(4) 3O~2~(g) <=> 2O~3~(g)")
    opts10 = [
        "A. (1) và (2).",
        "B. (1) và (4).",
        "C. (2) và (3).",
        "*D. Chỉ có (4).*" if is_teacher else "D. Chỉ có (4)."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khi tăng áp suất chung của hệ, cân bằng chuyển dịch theo chiều làm giảm số mol khí (chống lại sự tăng áp suất):\n- (1) Vế trái có 1 + 1 = 2 mol khí, vế phải có 2 mol khí (Δn_khí = 0). Khi tăng áp suất, cân bằng không chuyển dịch.\n- (2) NH~4~Cl(s) là chất rắn (0 mol khí); vế phải có 1 + 1 = 2 mol khí. Tăng áp suất chuyển dịch theo chiều nghịch.\n- (3) Vế trái là dung dịch (0 mol khí); vế phải có 1 mol khí O~2~. Tăng áp suất chuyển dịch theo chiều nghịch.\n- (4) Vế trái có 3 mol O~2~; vế phải có 2 mol O~3~ (số mol khí giảm từ 3 xuống 2). Tăng áp suất chuyển dịch theo chiều thuận.\nVậy chỉ có cân bằng (4) chuyển dịch theo chiều thuận khi tăng áp suất.\n=> Đáp án đúng: D.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 11. Dãy chất nào sau đây đều là chất điện li yếu?")
    opts11 = [
        "A. H~2~S, H~2~SO~3~, H~2~SO~4~.",
        "B. H~2~CO~3~, H~3~PO~4~, CH~3~COOH, Ba(OH)~2~.",
        "*C. H~2~S, CH~3~COOH, HClO.*" if is_teacher else "C. H~2~S, CH~3~COOH, HClO.",
        "D. H~2~CO~3~, H~2~SO~3~, HClO, Al~2~(SO~4~)~3~."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân loại các chất điện li:\n- A loại H~2~SO~4~ (acid mạnh, chất điện li mạnh).\n- B loại Ba(OH)~2~ (base mạnh, tan tốt, chất điện li mạnh).\n- D loại Al~2~(SO~4~)~3~ (muối tan, chất điện li mạnh).\n- C gồm H~2~S, CH~3~COOH, HClO đều là các acid yếu nên đều là chất điện li yếu.\n=> Đáp án đúng: C.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 12. Dung dịch nào sau đây có pH > 7?")
    opts12 = [
        "*A. NaOH.*" if is_teacher else "A. NaOH.",
        "B. H~2~SO~4~.",
        "C. NaCl.",
        "D. HNO~3~."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Xét môi trường dung dịch ở 25°C:\n- NaOH là base mạnh => môi trường base => [H^+^] < 10^-7^ M => pH > 7.\n- H~2~SO~4~, HNO~3~ là các acid mạnh => môi trường acid => pH < 7.\n- NaCl là muối tạo bởi base mạnh và acid mạnh => môi trường trung tính => pH = 7.\n=> Đáp án đúng: A.")

    # =========================================================================
    # PHẦN II. TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN II (2 điểm). Câu trắc nghiệm đúng sai.",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai."
    )

    # Câu 1 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Trong khí thải của quy trình sản xuất thuốc trừ sâu, phân bón hóa học có lẫn khí NH~3~. Khí này rất độc đối với sức khoẻ của con người và gây ô nhiễm môi trường. Con người hít phải khí này với lượng lớn sẽ gây ngộ độc: ho, đau ngực (nặng: đau thắt ngực, khó thở, thở nhanh, thở khò khè; chảy nước mắt và bỏng mắt, mù mắt, đau họng nặng, đau miệng; mạch nhanh, yếu, sốc; lẫn lộn, đi lại khó khăn, chóng mặt, thiếu sự phối hợp, bồn chồn, ngẩn ngơ).")
    stmts1 = [
        ("a", "Để xử lí NH~3~ lẫn trong khí thải, người ta có thể dẫn khí thải qua một bể lọc chứa dung dịch HCl.", True),
        ("b", "Tất cả các muối của NH~3~ đều bền với nhiệt.", False),
        ("c", "NH~3~ có tính base mạnh và tính khử mạnh.", False),
        ("d", "Đưa hai đũa thuỷ tinh có nhúng NH~3~ đặc và HCl đặc lại gần thấy khói màu trắng xuất hiện.", True)
    ]
    for label, text, ans_bool in stmts1:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_1 = (
            "Giải thích chi tiết Câu 1:\n"
            "- Ý a) ĐÚNG: Khí ammonia có tính base, tan tốt trong nước và phản ứng dễ dàng với dung dịch acid như HCl để tạo muối ammonium chloride (NH~4~Cl) không bay hơi, do đó dẫn qua dung dịch HCl là biện pháp xử lý khí thải hiệu quả.\n"
            "- Ý b) SAI: Tất cả các muối ammonium (muối của NH~3~) đều kém bền với nhiệt và dễ bị nhiệt phân hủy khi đun nóng (ví dụ: NH~4~Cl -> NH~3~ + HCl; NH~4~HCO~3~ -> NH~3~ + CO~2~ + H~2~O; NH~4~NO~3~ -> N~2~O + 2H~2~O).\n"
            "- Ý c) SAI: Ammonia chỉ có tính base yếu (trong nước phân li thuận nghịch tạo ion OH^- với hằng số K_b bé) và có tính khử (do nitrogen có số oxi hóa thấp nhất là -3).\n"
            "- Ý d) ĐÚNG: Khí NH~3~ và HCl đều dễ bay hơi từ dung dịch đặc. Khi đưa lại gần nhau, chúng phản ứng trong pha khí tạo thành các hạt tinh thể muối ammonium chloride (NH~4~Cl) lơ lửng như khói trắng: NH~3~(g) + HCl(g) -> NH~4~Cl(s).\n"
            "=> Kết luận: a - Đ, b - S, c - S, d - Đ."
        )
        add_solution_no_table(doc, sol_p2_1)

    # Câu 2 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Trong bình kín có hệ cân bằng hóa học sau:\nCO~2~(g) + H~2~(g) <=> CO(g) + H~2~O(g); Δ~r~H°~298~ > 0\nCho các phát biểu về sự tác động đến hệ cân bằng dưới đây, hãy xét tính đúng sai của các phát biểu đó.")
    stmts2 = [
        ("a", "Tăng nhiệt độ của hệ thì cân bằng chuyển dịch theo chiều thuận.", True),
        ("b", "Thêm chất xúc tác thì cân bằng chuyển dịch theo chiều thuận.", False),
        ("c", "Giảm áp suất chung của hệ thì cân bằng chuyển dịch theo chiều nghịch.", False),
        ("d", "Nếu nồng độ ban đầu CO~2~ và H~2~ đều là 0,03 M; tại cân bằng [CO] = 0,018 M thì giá trị K~c~ là 1,5.", False)
    ]
    for label, text, ans_bool in stmts2:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}) {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_2 = (
            "Giải thích chi tiết Câu 2:\n"
            "- Ý a) ĐÚNG: Phản ứng có Δ~r~H°~298~ > 0 nên chiều thuận là chiều thu nhiệt. Khi tăng nhiệt độ, theo nguyên lý Le Chatelier, cân bằng chuyển dịch theo chiều chống lại sự tăng nhiệt độ, tức chiều thu nhiệt (chiều thuận).\n"
            "- Ý b) SAI: Chất xúc tác chỉ làm tăng đồng thời tốc độ phản ứng thuận và phản ứng nghịch với số lần bằng nhau, giúp hệ nhanh đạt tới trạng thái cân bằng chứ không làm chuyển dịch cân bằng.\n"
            "- Ý c) SAI: Số mol khí ở hai vế: vế trái có 1 mol CO~2~ + 1 mol H~2~ = 2 mol khí; vế phải có 1 mol CO + 1 mol H~2~O = 2 mol khí (tổng số mol khí 2 vế bằng nhau). Do đó việc tăng hay giảm áp suất chung không làm chuyển dịch cân bằng.\n"
            "- Ý d) SAI:\n"
            "  Xét cân bằng:      CO~2~(g) + H~2~(g) <=> CO(g) + H~2~O(g)\n"
            "  Ban đầu (M):       0,03       0,03          0        0\n"
            "  Phản ứng (M):      0,018      0,018       0,018    0,018   (do [CO] = 0,018 M)\n"
            "  Cân bằng (M):      0,012      0,012       0,018    0,018\n"
            "  Hằng số cân bằng K~c~ = ([CO] . [H~2~O]) / ([CO~2~] . [H~2~]) = (0,018 . 0,018) / (0,012 . 0,012) = (1,5)^2^ = 2,25.\n"
            "  Giá trị đúng là K~c~ = 2,25 (1,5 chỉ là tỉ số nồng độ bậc một [CO]/[CO~2~]), do đó nhận định K~c~ = 1,5 là SAI.\n"
            "=> Kết luận: a - Đ, b - S, c - S, d - S."
        )
        add_solution_no_table(doc, sol_p2_2)

    # =========================================================================
    # PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN III (2 điểm). Câu trắc nghiệm trả lời ngắn.",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4."
    )

    # Câu 1 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Trong công nghiệp ammonia được tổng hợp từ nitrogen và hydrogen. Cho 14,874 lít N~2~ (đkc) tác dụng với lượng dư khí H~2~. Biết hiệu suất của phản ứng là 30%, khối lượng NH~3~ tạo thành là bao nhiêu gam? (Kết quả làm tròn đến hàng phần trăm)")
    if is_teacher:
        sol_p3_1 = (
            "Hướng dẫn giải Câu 1:\n"
            "- Ở điều kiện chuẩn (đkc: 25°C, 1 bar), 1 mol khí chiếm 24,79 lít:\n"
            "  n(N~2~) = 14,874 / 24,79 = 0,6 mol.\n"
            "- Phương trình phản ứng tổng hợp ammonia:\n"
            "  N~2~(g) + 3H~2~(g) <=> 2NH~3~(g)\n"
            "- Theo lí thuyết, n(NH~3~ lí thuyết) = 2 . n(N~2~) = 2 . 0,6 = 1,2 mol.\n"
            "- Vì hiệu suất phản ứng H = 30%, lượng NH~3~ thực tế thu được là:\n"
            "  n(NH~3~ thực tế) = 1,2 . 30% = 0,36 mol.\n"
            "- Khối lượng NH~3~ tạo thành:\n"
            "  m(NH~3~) = 0,36 . 17 = 6,12 gam.\n"
            "=> Đáp số: 6,12"
        )
        add_solution_no_table(doc, sol_p3_1)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 2 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Cho các chất dưới đây: HClO~4~, HClO, HF, HNO~3~, H~2~S, H~2~SO~3~, NaOH, NaCl, CuSO~4~, CH~3~COOH. Số chất thuộc loại chất điện li mạnh là bao nhiêu?")
    if is_teacher:
        sol_p3_2 = (
            "Hướng dẫn giải Câu 2:\n"
            "- Chất điện li mạnh gồm các acid mạnh, base mạnh và hầu hết các muối tan:\n"
            "  (1) HClO~4~: Acid rất mạnh (perchloric acid) -> điện li mạnh.\n"
            "  (2) HNO~3~: Acid mạnh -> điện li mạnh.\n"
            "  (3) NaOH: Base mạnh -> điện li mạnh.\n"
            "  (4) NaCl: Muối tan -> điện li mạnh.\n"
            "  (5) CuSO~4~: Muối tan -> điện li mạnh.\n"
            "- Các chất còn lại (HClO, HF, H~2~S, H~2~SO~3~, CH~3~COOH) đều là các acid yếu nên là chất điện li yếu.\n"
            "=> Tổng cộng có 5 chất thuộc loại chất điện li mạnh.\n"
            "=> Đáp số: 5"
        )
        add_solution_no_table(doc, sol_p3_2)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 3 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Thực hiện thí nghiệm chuẩn độ dung dịch HCl chưa biết nồng độ bằng dung dịch chuẩn NaOH 0,5M và thuốc thử phenolphthalein. Cho 10,00 mL dung dịch HCl vào bình tam giác cùng với 1 - 2 giọt thuốc thử. Thể tích dung dịch NaOH ở các lần chuẩn độ được ghi lại như sau:")
    
    # Bảng số liệu chuẩn độ Câu 3
    tbl_cb = doc.add_table(rows=2, cols=3)
    tbl_cb.alignment = WD_TABLE_ALIGNMENT.CENTER
    cb_widths = [4.0, 4.0, 4.0]
    cb_data = [
        ["Lần 1", "Lần 2", "Lần 3"],
        ["15,1 mL", "14,7 mL", "15,2 mL"]
    ]
    for r_idx, row in enumerate(tbl_cb.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for c_idx, cell in enumerate(row.cells):
            cell.width = Cm(cb_widths[c_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tcPr = cell._tc.get_or_add_tcPr()
            borders_xml = f'''<w:tcBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
            </w:tcBorders>'''
            tcPr.append(parse_xml(borders_xml))
            p_cell = cell.paragraphs[0]
            p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cell.paragraph_format.space_before = Pt(2)
            p_cell.paragraph_format.space_after = Pt(2)
            if r_idx == 0:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8F5E9"/>')
                tcPr.append(shd)
                render_rich_text(p_cell, f"*{cb_data[r_idx][c_idx]}*", base_size=10, base_bold=True)
            else:
                render_rich_text(p_cell, cb_data[r_idx][c_idx], base_size=10)

    p_post = doc.add_paragraph()
    p_post.paragraph_format.space_before = Pt(2)
    p_post.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_post, "Dung dịch HCl có nồng độ bằng bao nhiêu mol/L? (Kết quả làm tròn đến hàng phần trăm)")

    if is_teacher:
        sol_p3_3 = (
            "Hướng dẫn giải Câu 3:\n"
            "- Thể tích dung dịch NaOH trung bình sau 3 lần chuẩn độ là:\n"
            "  V_tb(NaOH) = (15,1 + 14,7 + 15,2) / 3 = 45,0 / 3 = 15,00 mL.\n"
            "- Phương trình phản ứng chuẩn độ:\n"
            "  HCl + NaOH -> NaCl + H~2~O\n"
            "- Tại điểm tương đương, n(HCl) = n(NaOH):\n"
            "  C_M(HCl) . V(HCl) = C_M(NaOH) . V_tb(NaOH)\n"
            "  => C_M(HCl) . 10,00 = 0,5 . 15,00\n"
            "  => C_M(HCl) = (0,5 . 15,00) / 10,00 = 7,50 / 10,00 = 0,75 M (hay mol/L).\n"
            "=> Đáp số: 0,75"
        )
        add_solution_no_table(doc, sol_p3_3)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 4 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Cho phản ứng: CO(g) + H~2~O(g) <=> CO~2~(g) + H~2~(g)   Δ~r~H°~298~ = -41 kJ\nCác phát biểu sau đây, có bao nhiêu phát biểu đúng về phản ứng trên?\n(1) Phản ứng theo chiều thuận là phản ứng thu nhiệt.\n(2) Tăng nhiệt độ cân bằng chuyển dịch theo chiều nghịch.\n(3) Thêm khí H~2~ vào hệ cân bằng chuyển dịch theo chiều nghịch.\n(4) Tăng áp suất chung bằng cách nén cho thể tích của hệ giảm xuống cân bằng chuyển dịch theo chiều thuận.")
    if is_teacher:
        sol_p3_4 = (
            "Hướng dẫn giải Câu 4:\n"
            "- (1) SAI: Phản ứng có Δ~r~H°~298~ = -41 kJ < 0 nên phản ứng theo chiều thuận là phản ứng tỏa nhiệt.\n"
            "- (2) ĐÚNG: Vì chiều thuận tỏa nhiệt, khi tăng nhiệt độ thì cân bằng chuyển dịch theo chiều thu nhiệt (chiều nghịch).\n"
            "- (3) ĐÚNG: H~2~ là chất sản phẩm ở vế phải. Khi thêm H~2~, theo nguyên lý Le Chatelier, cân bằng chuyển dịch theo chiều làm giảm nồng độ H~2~, tức chiều nghịch.\n"
            "- (4) SAI: Vế trái có 1 mol CO + 1 mol H~2~O = 2 mol khí; vế phải có 1 mol CO~2~ + 1 mol H~2~ = 2 mol khí. Tổng số mol khí ở 2 vế bằng nhau nên việc thay đổi áp suất chung không làm chuyển dịch vị trí cân bằng.\n"
            "=> Có 2 phát biểu đúng là (2) và (3).\n"
            "=> Đáp số: 2"
        )
        add_solution_no_table(doc, sol_p3_4)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # =========================================================================
    # B. TỰ LUẬN (3 ĐIỂM)
    # =========================================================================
    p_sec_b = doc.add_paragraph()
    p_sec_b.paragraph_format.space_before = Pt(8)
    p_sec_b.paragraph_format.space_after = Pt(2)
    render_rich_text(p_sec_b, "*B. TỰ LUẬN (3 điểm)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # Câu 1 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1 (1,0 điểm). Viết phương trình điện li của các chất sau trong nước: KOH, Al~2~(SO~4~)~3~, CH~3~COOH, (NH~4~)~3~PO~4~.")
    if is_teacher:
        sol_tl1 = (
            "Hướng dẫn giải Câu 1 (1,0 điểm):\n"
            "- KOH là base mạnh, chất điện li mạnh, phân li hoàn toàn:\n"
            "  KOH -> K^+^ + OH^-^\n"
            "- Al~2~(SO~4~)~3~ là muối tan, chất điện li mạnh, phân li hoàn toàn:\n"
            "  Al~2~(SO~4~)~3~ -> 2Al^3+^ + 3SO~4~^2-^\n"
            "- CH~3~COOH là carboxylic acid yếu, chất điện li yếu, phân li thuận nghịch:\n"
            "  CH~3~COOH <=> CH~3~COO^-^ + H^+^\n"
            "  (hoặc: CH~3~COOH + H~2~O <=> CH~3~COO^-^ + H~3~O^+^)\n"
            "- (NH~4~)~3~PO~4~ là muối tan, chất điện li mạnh, phân li hoàn toàn:\n"
            "  (NH~4~)~3~PO~4~ -> 3NH~4~^+^ + PO~4~^3-^"
        )
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2 (1,0 điểm). Trong công nghiệp, người ta sản xuất nitric acid (HNO~3~) từ ammonia theo sơ đồ chuyển hoá sau:\nNH~3~ ──(+O~2~, t°, xt)──> NO ──(+O~2~)──> NO~2~ ──(+O~2~ + H~2~O)──> HNO~3~\na) Viết các phương trình hoá học xảy ra.\nb) Để điều chế 200 000 tấn nitric acid có nồng độ 60% cần dùng bao nhiêu tấn ammonia? Biết rằng hiệu suất của quá trình sản xuất nitric acid theo sơ đồ trên là 96,2%.")
    if is_teacher:
        sol_tl2 = (
            "Hướng dẫn giải Câu 2 (1,0 điểm):\n"
            "a) (0,5 điểm) Các phương trình hoá học xảy ra theo chu trình Ostwald:\n"
            "  (1) Oxi hoá ammonia bằng oxygen có xúc tác Pt ở nhiệt độ cao (850 - 900°C):\n"
            "      4NH~3~ + 5O~2~ ──(xt Pt, t°)──> 4NO + 6H~2~O\n"
            "  (2) Oxi hoá khí NO thành NO~2~ ở nhiệt độ thường:\n"
            "      2NO + O~2~ -> 2NO~2~\n"
            "  (3) Hấp thụ NO~2~ cùng với O~2~ và nước tạo thành dung dịch nitric acid:\n"
            "      4NO~2~ + O~2~ + 2H~2~O -> 4HNO~3~\n\n"
            "b) (0,5 điểm) Tính khối lượng ammonia cần dùng:\n"
            "- Khối lượng chất tan HNO~3~ nguyên chất cần điều chế:\n"
            "  m(HNO~3~) = 200 000 . 60% = 120 000 tấn.\n"
            "- Sơ đồ bảo toàn nguyên tố Nitrogen (N):\n"
            "  NH~3~ (17 tấn)  ──>  HNO~3~ (63 tấn)\n"
            "- Theo lí thuyết (hiệu suất 100%), khối lượng NH~3~ cần dùng là:\n"
            "  m_lt(NH~3~) = (120 000 . 17) / 63 = 2 040 000 / 63 ≈ 32 380,95 tấn.\n"
            "- Do hiệu suất của toàn bộ quá trình sản xuất là H = 96,2%, nên khối lượng NH~3~ thực tế cần lấy là:\n"
            "  m_tt(NH~3~) = m_lt(NH~3~) / H = 32 380,9524 / 0,962 ≈ 33 660,03 tấn (hoặc ~33 660 tấn).\n"
            "=> Đáp số: 33 660,03 tấn ammonia."
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=10)

    # Câu 3 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3 (1,0 điểm). Trộn 80 mL dung dịch HCl 0,5M với 120 mL dung dịch NaOH 0,6M. Tính pH của dung dịch thu được. (Kết quả làm tròn đến hàng phần mười).")
    if is_teacher:
        sol_tl3 = (
            "Hướng dẫn giải Câu 3 (1,0 điểm):\n"
            "- Thể tích dung dịch sau khi trộn:\n"
            "  V_dd = 80 + 120 = 200 mL = 0,2 lít.\n"
            "- Số mol các chất ban đầu:\n"
            "  n(H^+^) = n(HCl) = 0,08 . 0,5 = 0,04 mol.\n"
            "  n(OH^-^) = n(NaOH) = 0,12 . 0,6 = 0,072 mol.\n"
            "- Phương trình ion thu gọn:\n"
            "             H^+^   +   OH^-^   ->   H~2~O\n"
            "  Ban đầu:   0,04       0,072               (mol)\n"
            "  Phản ứng:  0,04       0,04                (mol)\n"
            "  Sau pư:    0          0,032               (mol)\n"
            "- Sau phản ứng, ion OH^-^ còn dư, n(OH^-^ dư) = 0,072 - 0,04 = 0,032 mol.\n"
            "- Nồng độ ion OH^-^ trong dung dịch thu được:\n"
            "  [OH^-^] = 0,032 / 0,2 = 0,16 M.\n"
            "- Tính pOH và pH:\n"
            "  pOH = -log[OH^-^] = -log(0,16) ≈ 0,796\n"
            "  => pH = 14 - pOH = 14 - 0,796 = 13,204 ≈ 13,2.\n"
            "  (Hoặc: [H^+^] = 10^-14^ / 0,16 = 6,25 . 10^-14^ M => pH = -log(6,25 . 10^-14^) ≈ 13,2).\n"
            "=> Đáp số: pH = 13,2."
        )
        add_solution_no_table(doc, sol_tl3)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Cuối đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "-------------------- *HẾT* --------------------", base_size=10.5, base_bold=True)

    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(2)
    p_note.paragraph_format.space_after = Pt(2)
    p_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_note, "_- Học sinh không được sử dụng tài liệu;_\n_- Giám thị không giải thích gì thêm._", base_size=9.5, base_italic=True)

    if output_path:
        doc.save(output_path)
        print(f"Đã lưu thành công: {output_path}")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 1. Bản Học sinh
    hs_path_std = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_11_ChuyenKhoaHocHue_Ma135.docx")
    hs_path_alt = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_11_ChuyenKhoaHocHue_Ma135_HocSinh.docx")
    build_exam_chuyenkhoahoc_135(is_teacher=False, output_path=hs_path_std)
    build_exam_chuyenkhoahoc_135(is_teacher=False, output_path=hs_path_alt)

    # 2. Bản Giáo viên
    gv_path = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_11_ChuyenKhoaHocHue_Ma135_GiaoVien.docx")
    build_exam_chuyenkhoahoc_135(is_teacher=True, output_path=gv_path)

    print("\n>>> Hoàn thành xuất 2 file Word theo Form Đề Kiểm Tra chuẩn!")

if __name__ == "__main__":
    main()
