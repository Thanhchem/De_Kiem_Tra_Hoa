# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì I Hóa Học 11 - Mã Đề 102
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản Word:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_11_Ma102_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_11_Ma102_GiaoVien.docx
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


def add_quick_answer_table_ma102(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu MÃ ĐỀ 102.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 102)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I (12 câu - 3.0 điểm)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - mỗi câu đúng 0,25 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_1 = [2.4] + [1.3] * 12
    p1_answers = ["B", "A", "B", "B", "D", "A", "D", "D", "B", "A", "B", "C"]

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

    # 2. PHẦN II (2 câu Đúng/Sai - 2.0 điểm)
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(6)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm - 1 ý: 0,1đ; 2 ý: 0,25đ; 3 ý: 0,5đ; 4 ý: 1,0đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_2 = [3.0, 3.0, 3.0, 3.0, 3.0, 3.0]
    p2_answers = [
        ("Câu 13", "Đ", "S", "S", "Đ"),
        ("Câu 14", "S", "Đ", "S", "Đ"),
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

    # 3. PHẦN III (4 câu trả lời ngắn - 2.0 điểm)
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(6)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*3. Phần III: Trắc nghiệm Trả lời ngắn (2,0 điểm - mỗi câu đúng 0,5 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_3 = [3.6, 3.6, 3.6, 3.6, 3.6]
    p3_answers = ["1", "64,0", "3", "123"]

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
                txt = "Câu" if col_idx == 0 else f"Câu {col_idx + 14}"
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


def build_exam_document(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="TRƯỜNG THPT CHUYÊN / THPT TỈNH",
        exam_title="ĐỀ KIỂM TRA GIỮA HỌC KÌ I",
        subject_line="MÔN: HOÁ HỌC - LỚP 11",
        num_pages="04",
        duration="50 phút",
        exam_code="102",
        atomic_masses="Cu = 64; Fe = 56; N = 14; O = 16; S = 32; H = 1; K = 39; Cl = 35,5; C = 12; Ca = 40",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_ma102(doc)

    # ========================== PHẦN A. TRẮC NGHIỆM ==========================
    p_pa = doc.add_paragraph()
    p_pa.paragraph_format.space_before = Pt(6)
    p_pa.paragraph_format.space_after = Pt(2)
    render_rich_text(p_pa, "*A. PHẦN TRẮC NGHIỆM (7,0 ĐIỂM)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # -------------------------- PHẦN I --------------------------
    add_section_header(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án. (3,0 điểm)"
    )

    # CÂU 1
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 1. Sự chuyển dịch cân bằng hóa học là sự dịch chuyển từ trạng thái cân bằng hóa học này")
    opts_1 = [
        "A. sang trạng thái không cân bằng do tác động của các yếu tố từ bên ngoài tác động lên cân bằng.",
        "*B. sang trạng thái cân bằng hóa học khác do tác động của các yếu tố từ bên ngoài tác động lên cân bằng.*",
        "C. sang trạng thái cân bằng hóa học khác không cần có tác động của các yếu tố từ bên ngoài tác động lên cân bằng.",
        "D. sang trạng thái cân bằng hóa học khác do cân bằng hóa học tác động lên các yếu tố bên ngoài."
    ]
    render_options_no_table(doc, opts_1, is_teacher=is_teacher)
    if is_teacher:
        sol_1 = [
            "*Chọn B.*",
            "Sự chuyển dịch cân bằng hóa học là sự chuyển dịch từ trạng thái cân bằng này sang trạng thái cân bằng khác do tác động của các yếu tố từ bên ngoài (như nhiệt độ, nồng độ, áp suất) lên cân bằng."
        ]
        add_solution(doc, sol_1)

    # CÂU 2
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 2. Cho Cu phản ứng với dung dịch HNO~3~ loãng, nóng thu được một chất khí không màu hóa nâu trong không khí, khí đó là")
    opts_2 = [
        "*A. NO.*",
        "B. NH~3~.",
        "C. N~2~O.",
        "D. N~2~."
    ]
    render_options_no_table(doc, opts_2, is_teacher=is_teacher)
    if is_teacher:
        sol_2 = [
            "*Chọn A.*",
            "Phản ứng: 3Cu + 8HNO~3~(loãng) -> 3Cu(NO~3~)~2~ + 2NO^ + 4H~2~O.",
            "Khí NO là khí không màu, khi tiếp xúc với oxygen trong không khí dễ dàng bị oxi hóa thành khí nitrogen dioxide (NO~2~) có màu nâu đỏ: 2NO + O~2~ -> 2NO~2~."
        ]
        add_solution(doc, sol_2)

    # CÂU 3
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 3. Trong thực tế, khi bảo quản dung dịch muối sắt(III) chloride (FeCl~3~) trong phòng thí nghiệm, người ta thường phải cho thêm vài giọt dung dịch HCl đặc vào lọ đựng. Việc làm này nhằm mục đích chủ yếu là:")
    opts_3 = [
        "A. Ngăn cản quá trình bay hơi của dung dịch FeCl~3~, làm cho dung dịch trong suốt, không màu.",
        "*B. Ức chế sự thủy phân của ion Fe^3+^ tạo kết tủa Fe(OH)~3~ và ngăn sự oxi hóa Fe^2+^ (nếu có).*",
        "C. Làm tăng màu vàng nâu đặc trưng của dung dịch FeCl~3~.",
        "D. Tăng tốc độ điện li của muối FeCl~3~."
    ]
    render_options_no_table(doc, opts_3, is_teacher=is_teacher)
    if is_teacher:
        sol_3 = [
            "*Chọn B.*",
            "Trong dung dịch nước, ion Fe^3+^ bị thủy phân tạo môi trường acid và có thể tạo kết tủa hydroxide: Fe^3+^ + 3H~2~O <=> Fe(OH)~3~ v + 3H^+^.",
            "Khi thêm vài giọt HCl đặc (tăng nồng độ ion H^+^), theo nguyên lí Le Chatelier, cân bằng sẽ dịch chuyển sang chiều nghịch, giúp ức chế sự thủy phân của Fe^3+^, giữ cho dung dịch trong và ổn định."
        ]
        add_solution(doc, sol_3)

    # CÂU 4
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 4. Ở điều kiện thường, sulfur tồn tại ở dạng tinh thể, được tạo nên từ các phân tử sulfur. Số nguyên tử trong mỗi phân tử sulfur là")
    opts_4 = [
        "A. 4.",
        "*B. 8.*",
        "C. 6.",
        "D. 2."
    ]
    render_options_no_table(doc, opts_4, is_teacher=is_teacher)
    if is_teacher:
        sol_4 = [
            "*Chọn B.*",
            "Ở điều kiện thường, sulfur tồn tại ở trạng thái rắn dạng tinh thể, gồm các phân tử S~8~ có cấu tạo mạch vòng 8 nguyên tử liên kết cộng hóa trị với nhau."
        ]
        add_solution(doc, sol_4)

    # CÂU 5
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 5. Với xúc tác của các ion kim loại trong khói bụi, các oxide của sulfur và nitrogen bị oxi hóa bởi oxygen, ozone, hydrogen peroxide, gốc tự do,... rồi hòa tan vào nước tạo thành các acid tương ứng. Hai acid tạo thành từ quá trình trên là")
    opts_5 = [
        "A. H~2~S và H~2~SO~4~.",
        "B. H~2~CO~3~ và HNO~3~.",
        "C. H~2~S và HNO~3~.",
        "*D. H~2~SO~4~ và HNO~3~.*"
    ]
    render_options_no_table(doc, opts_5, is_teacher=is_teacher)
    if is_teacher:
        sol_5 = [
            "*Chọn D.*",
            "Đây là cơ chế chính tạo nên hiện tượng mưa acid:",
            "+ Khí SO~2~ bị oxi hóa thành SO~3~ rồi hòa tan tạo H~2~SO~4~: 2SO~2~ + O~2~ + 2H~2~O -(xt)-> 2H~2~SO~4~.",
            "+ Các oxide của nitrogen (NO, NO~2~) bị oxi hóa và hòa tan tạo HNO~3~: 4NO~2~ + O~2~ + 2H~2~O -> 4HNO~3~."
        ]
        add_solution(doc, sol_5)

    # CÂU 6
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 6. Trong khí quyển, nguyên tố nitrogen tồn tại chủ yếu dưới dạng")
    opts_6 = [
        "*A. N~2~.*",
        "B. NH~4~^+^.",
        "C. NO~3~^-^.",
        "D. NO~2~^-^."
    ]
    render_options_no_table(doc, opts_6, is_teacher=is_teacher)
    if is_teacher:
        sol_6 = [
            "*Chọn A.*",
            "Trong khí quyển, nitrogen tồn tại chủ yếu ở dạng đơn chất khí N~2~ (chiếm khoảng 78% thể tích không khí)."
        ]
        add_solution(doc, sol_6)

    # CÂU 7
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 7. Hiện tượng phú dưỡng có thể quan sát được thông qua sự xuất hiện dày đặc của tảo xanh trong nước. Khi đó, lượng oxygen trong nước sẽ nhanh chóng giảm đi. Điều này là do:\n(a) Quá trình phân hủy tảo chết bởi vi khuẩn đã tiêu tốn lượng lớn oxygen trong nước.\n(b) Sự phát triển của tảo xanh đã ngăn cản ánh sáng và không khí chứa oxygen khuếch tán vào nước.\n(c) Sự hoạt động của lượng lớn vi khuẩn đã hấp thụ đáng kể oxygen hòa tan trong nước.\nSố nhận định đúng là")
    opts_7 = [
        "A. 1.",
        "B. 0.",
        "C. 2.",
        "*D. 3.*"
    ]
    render_options_no_table(doc, opts_7, is_teacher=is_teacher)
    if is_teacher:
        sol_7 = [
            "*Chọn D.*",
            "Cả 3 nhận định (a), (b), (c) đều đúng và mô tả chính xác quá trình suy giảm lượng dissolved oxygen (DO) khi ao hồ bị phú dưỡng:",
            "- (a) Đúng: Tảo chết lắng xuống đáy, vi khuẩn hiếu khí phân hủy tiêu tốn lượng lớn O~2~.",
            "- (b) Đúng: Lớp tảo nở hoa dày đặc che chắn ánh sáng mặt trời làm ngừng quang hợp ở tầng sâu, đồng thời cản trở O~2~ khuếch tán từ không khí vào nước.",
            "- (c) Đúng: Lượng vi sinh vật phân hủy tăng vọt tiêu thụ đáng kể O~2~ hòa tan."
        ]
        add_solution(doc, sol_7)

    # CÂU 8
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 8. Chất nào sau đây thuộc loại chất điện li yếu?")
    opts_8 = [
        "A. NH~4~Cl.",
        "B. KCl.",
        "C. HNO~3~.",
        "*D. HF.*"
    ]
    render_options_no_table(doc, opts_8, is_teacher=is_teacher)
    if is_teacher:
        sol_8 = [
            "*Chọn D.*",
            "- NH~4~Cl, KCl là các muối tan -> chất điện li mạnh.",
            "- HNO~3~ là acid mạnh -> chất điện li mạnh.",
            "- HF là acid yếu, trong nước chỉ phân li một phần ra ion: HF <=> H^+^ + F^-^ -> chất điện li yếu."
        ]
        add_solution(doc, sol_8)

    # CÂU 9
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 9. Khí nitrogen ít tan trong nước là do")
    opts_9 = [
        "A. nitrogen có độ âm điện lớn.",
        "*B. phân tử nitrogen không phân cực.*",
        "C. nitrogen có bán kính nguyên tử nhỏ.",
        "D. phân tử nitrogen có liên kết ba bền vững."
    ]
    render_options_no_table(doc, opts_9, is_teacher=is_teacher)
    if is_teacher:
        sol_9 = [
            "*Chọn B.*",
            "Nước (H~2~O) là dung môi phân cực mạnh. Phân tử nitrogen (N#N) gồm hai nguyên tử giống nhau nên có cấu tạo hoàn toàn đối xứng và không phân cực. Theo quy tắc độ tan 'chất phân cực tan trong dung môi phân cực, chất không phân cực khó tan trong dung môi phân cực', phân tử N~2~ rất ít tan trong nước."
        ]
        add_solution(doc, sol_9)

    # CÂU 10
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 10. Phân tử ammonia có dạng hình học nào sau đây?")
    opts_10 = [
        "*A. Chóp tam giác.*",
        "B. Chóp tứ giác.",
        "C. Chữ T.",
        "D. Tam giác đều."
    ]
    render_options_no_table(doc, opts_10, is_teacher=is_teacher)
    if is_teacher:
        sol_10 = [
            "*Chọn A.*",
            "Trong phân tử NH~3~, nguyên tử trung tâm N ở trạng thái lai hóa sp^3^ liên kết với 3 nguyên tử H và còn 1 cặp electron tự do (chưa tham gia liên kết), tạo nên dạng hình học phân tử là chóp tam giác (pyramidal)."
        ]
        add_solution(doc, sol_10)

    # CÂU 11
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 11. Chất chỉ thị acid - base là những chất")
    opts_11 = [
        "A. Chỉ có khả năng đổi màu trong môi trường base mạnh.",
        "*B. Có màu sắc thay đổi tùy thuộc vào giá trị pH của dung dịch.*",
        "C. Chỉ có khả năng đổi màu trong môi trường acid mạnh.",
        "D. Có màu sắc không thay đổi khi tiếp xúc với các acid hoặc base."
    ]
    render_options_no_table(doc, opts_11, is_teacher=is_teacher)
    if is_teacher:
        sol_11 = [
            "*Chọn B.*",
            "Chất chỉ thị acid - base là chất có màu sắc biến đổi theo giá trị pH của dung dịch (như quỳ tím, phenolphthalein, chất chỉ thị vạn năng,...), được dùng để nhận biết khoảng giá trị pH hoặc xác định điểm tương đương trong chuẩn độ."
        ]
        add_solution(doc, sol_11)

    # CÂU 12
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 12. Khi tiến hành các thí nghiệm khảo sát tính háo nước và tính oxi hóa mạnh của H~2~SO~4~ đặc trên các mẫu vật liệu: tờ giấy trắng, hạt gạo, và đường mía; hiện tượng chung quan sát được trên các mẫu vật này là gì và giải thích đúng nhất là?")
    opts_12 = [
        "A. Cả ba mẫu vật chỉ bị mất nước tạm thời, khi cho rửa lại bằng nước sạch thì các mẫu vật sẽ trở lại trạng thái ban đầu.",
        "B. Giấy và đường mía bị hóa đen do bị cháy do nhiệt độ cao, còn gạo không có hiện tượng gì.",
        "*C. Tất cả các mẫu vật đều bị hóa đen; nguyên nhân là do H~2~SO~4~ đặc chiếm giữ các nguyên tố H và O (thành phần cấu tạo nên carbohydrate) để lại khung carbon.*",
        "D. H~2~SO~4~ đặc thủy phân hoàn toàn các chất này thành dung dịch acid hữu cơ làm tờ giấy trắng, hạt gạo, và đường mía sủi bọt khí hydrogen H~2~."
    ]
    render_options_no_table(doc, opts_12, is_teacher=is_teacher)
    if is_teacher:
        sol_12 = [
            "*Chọn C.*",
            "Tờ giấy trắng (cellulose), hạt gạo (tinh bột) và đường mía (sucrose) đều thuộc nhóm carbohydrate có công thức chung C~n~(H~2~O)~m~.",
            "H~2~SO~4~ đặc có tính háo nước cực mạnh, nó đã chiếm đoạt các nguyên tố H và O theo tỉ lệ của nước ra khỏi phân tử carbohydrate, biến chúng thành than carbon (C) có màu đen: C~n~(H~2~O)~m~ -(H~2~SO~4~ đ)-> nC (đen) + mH~2~O."
        ]
        add_solution(doc, sol_12)

    # -------------------------- PHẦN II --------------------------
    add_section_header(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 13 đến câu 14. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai. (2,0 điểm)"
    )

    # CÂU 13
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 13. Cho phản ứng hóa học sau xảy ra trong bình kín:\nN~2~(g) + 3H~2~(g) <=> 2NH~3~(g)   \\Delta_r H^0_298 = -92 kJ")
    stmts_13 = [
        ("a) Ở một nhiệt độ xác định, biểu thức xác định hằng số cân bằng K~c~ = [NH~3~]^2^ / ([N~2~][H~2~]^3^)", "Đúng"),
        ("b) Vì phản ứng thuận là phản ứng tỏa nhiệt, nên để nâng cao hiệu suất tổng hợp NH~3~ trong công nghiệp, người ta thường thực hiện phản ứng ở nhiệt độ vừa phải 500 ^o^C.", "Sai"),
        ("c) NH~3~ một chất khí không màu, không mùi, làm xanh giấy quỳ tím ẩm.", "Sai"),
        ("d) Biến thiên enthalpy của phản ứng có giá trị âm \\Delta_r H^0_298 < 0 chứng tỏ hệ tỏa nhiệt ra môi trường khi phản ứng thuận xảy ra.", "Đúng")
    ]
    for lbl, ans_val in stmts_13:
        p_stmt = doc.add_paragraph()
        render_statement(p_stmt, lbl, ans=ans_val, is_teacher=is_teacher)
    if is_teacher:
        sol_13 = [
            "*Lời giải chi tiết:*",
            "- *Ý a) ĐÚNG.* Theo định luật tác dụng khối lượng, biểu thức hằng số cân bằng nồng độ của phản ứng là: K~c~ = [NH~3~]^2^ / ([N~2~].[H~2~]^3^).",
            "- *Ý b) SAI.* Phản ứng thuận tỏa nhiệt (\\Delta H < 0), theo nguyên lí Le Chatelier, để tăng hiệu suất (chuyển dịch theo chiều thuận) thì về mặt lý thuyết phải *giảm nhiệt độ*. Tuy nhiên ở nhiệt độ thấp thì tốc độ phản ứng quá chậm. Con số khoảng 450 - 500 ^o^C là nhiệt độ tối ưu cân đối để phản ứng đạt vận tốc đủ nhanh với xúc tác bột sắt, chứ ở nhiệt độ này hiệu suất thấp hơn so với nhiệt độ thấp. Do đó câu khẳng định 'để nâng cao hiệu suất... nên thực hiện ở 500 ^o^C' là sai về lập luận nhiệt động học.",
            "- *Ý c) SAI.* NH~3~ là khí không màu, nhưng có *mùi khai xốc đặc trưng*, rất độc và kích ứng đường hô hấp, không phải 'không mùi'.",
            "- *Ý d) ĐÚNG.* Giá trị biến thiên enthalpy chuẩn âm (\\Delta_r H^0_298 = -92 kJ < 0) theo quy ước nhiệt động học biểu thị quá trình thuận là phản ứng tỏa nhiệt ra môi trường xung quanh."
        ]
        add_solution(doc, sol_13)

    # CÂU 14
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 14. Sulfur dioxide (SO~2~) là một trong các khí gây ô nhiễm môi trường, là nguyên nhân chính gây ra mưa acid và nhiều tác hại đối với sức khỏe con người. Trong công nghiệp, khí này được tạo ra từ phản ứng đốt cháy sulfur theo cân bằng sau: S(s) + O~2~(g) -(t^o^)-> SO~2~(g)")
    stmts_14 = [
        ("a) Nước mưa bình thường không ô nhiễm có pH khoảng 5,6 do có CO~2~ hòa tan. Nếu nước mưa tại vùng công nghiệp nhiễm SO~2~, tạo ra acid H~2~SO~4~ làm pH giảm xuống dưới 5,0 thì quỳ tím sẽ chuyển sang màu xanh.", "Sai"),
        ("b) Trong tự nhiên, nguyên tố sulfur (lưu huỳnh) tồn tại phổ biến ở dạng đơn chất trong lòng đất (mỏ lưu huỳnh) và dạng hợp chất trong các khoáng vật như quặng pyrite và thạch cao.", "Đúng"),
        ("c) Để giảm thiểu tác hại của SO~2~, các nhà máy được khuyến khích xây dựng các ống khói thật cao để khí SO~2~ khuếch tán nhanh lên tầng cao khí quyển, chuyển hóa triệt để thành chất an toàn mà không gây mưa acid.", "Sai"),
        ("d) Để xác định hàm lượng SO~2~ trong không khí, người ta dẫn lượng khí này qua dung dịch KMnO~4~. Một học sinh tiến hành chuẩn độ 50 mL dung dịch acid H~2~SO~4~ (được tạo ra sau khi thu khí SO~2~ bằng dung dịch NaOH 0,0001 M). Thể tích NaOH tiêu tốn trung bình là 15 mL khi dung dịch chuyển sang màu hồng nhạt (chỉ thị phenolphthalein). Nồng độ mol của dung dịch H~2~SO~4~ ban đầu thu được là 1,5.10^-5^ M.", "Đúng")
    ]
    for lbl, ans_val in stmts_14:
        p_stmt = doc.add_paragraph()
        render_statement(p_stmt, lbl, ans=ans_val, is_teacher=is_teacher)
    if is_teacher:
        sol_14 = [
            "*Lời giải chi tiết:*",
            "- *Ý a) SAI.* Khi nước mưa nhiễm SO~2~ tạo ra acid H~2~SO~4~ làm pH < 5,0 (môi trường acid), giấy quỳ tím phải chuyển sang *màu đỏ* (hoặc hồng), không thể chuyển sang 'màu xanh'.",
            "- *Ý b) ĐÚNG.* Sulfur phân bố rộng rãi trong vỏ Trái Đất: dạng đơn chất ở các mỏ lưu huỳnh tự do, dạng hợp chất gồm sulfide (pyrite FeS~2~, chu sa HgS, galena PbS) và sulfate (thạch cao CaSO~4~.2H~2~O, barite BaSO~4~).",
            "- *Ý c) SAI.* Việc xây ống khói cao chỉ làm phân tán khí SO~2~ ra phạm vi xa hơn xung quanh nhà máy, hoàn toàn không chuyển hóa SO~2~ thành chất vô hại mà ngược lại còn làm phát tán khí lên cao, gây mưa acid ở các vùng lân cận rộng lớn hơn. Biện pháp đúng là lắp hệ thống xử lý khí thải khử SO~2~ (bằng nước vôi hoặc đá vôi).",
            "- *Ý d) ĐÚNG.*",
            "  + Phản ứng chuẩn độ: H~2~SO~4~ + 2NaOH -> Na~2~SO~4~ + 2H~2~O.",
            "  + Số mol NaOH đã phản ứng: n_NaOH = 0,0001 x 0,015 = 1,5.10^-6^ mol.",
            "  + Số mol H~2~SO~4~: n_H~2~SO~4~ = 1/2 n_NaOH = 0,75.10^-6^ mol.",
            "  + Nồng độ mol của dung dịch H~2~SO~4~: C_M = (0,75.10^-6^) / 0,050 = 1,5.10^-5^ M."
        ]
        add_solution(doc, sol_14)

    # -------------------------- PHẦN III --------------------------
    add_section_header(
        doc,
        "PHẦN III. Câu hỏi trắc nghiệm yêu cầu trả lời ngắn.",
        "Thí sinh trả lời từ câu 15 đến câu 18. (2,0 điểm)"
    )

    # CÂU 15
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 15. Giá trị pH của dung dịch H~2~SO~4~ 0,05 M là bao nhiêu?")
    if is_teacher:
        sol_15 = [
            "*Trả lời:* *1*",
            "*Lời giải:* H~2~SO~4~ là acid mạnh phân li hoàn toàn: H~2~SO~4~ -> 2H^+^ + SO~4~^2-^.",
            "=> [H^+^] = 2 x C_M(H~2~SO~4~) = 2 x 0,05 = 0,1 M = 10^-1^ M.",
            "=> pH = -log[H^+^] = -log(0,1) = 1."
        ]
        add_solution(doc, sol_15)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 16
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 16. Cho a gam Cu phản ứng hoàn toàn với dung dịch H~2~SO~4~ đặc, nóng, dư thu được 24,79 lít khí SO~2~ (ở đkc, sản phẩm khử duy nhất). Giá trị của a? (Làm tròn kết quả đến hàng phần mười)")
    if is_teacher:
        sol_16 = [
            "*Trả lời:* *64,0*",
            "*Lời giải:*",
            "Ở điều kiện chuẩn (đkc: 25 ^o^C, 1 bar), thể tích mol phân tử chất khí là 24,79 L/mol.",
            "=> n_SO~2~ = 24,79 / 24,79 = 1,0 mol.",
            "Phương trình phản ứng: Cu + 2H~2~SO~4~(đặc) -(t^o^)-> CuSO~4~ + SO~2~^ + 2H~2~O.",
            "Theo phương trình: n_Cu = n_SO~2~ = 1,0 mol.",
            "Khối lượng kim loại Cu đã phản ứng: a = 1,0 x 64 = 64,0 gam."
        ]
        add_solution(doc, sol_16)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 17
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 17. Cho các chất: Au, Al, FeO, Fe(OH)~2~, CaCO~3~, K~2~SO~4~ lần lượt tác dụng với dung dịch HNO~3~ loãng. Có bao nhiêu trường hợp xảy ra phản ứng oxi hóa - khử?")
    if is_teacher:
        sol_17 = [
            "*Trả lời:* *3*",
            "*Lời giải:* Phản ứng oxi hóa - khử là phản ứng có sự thay đổi số oxi hóa của các nguyên tố:",
            "1. Au: Kim loại quý kém hoạt động, không phản ứng với HNO~3~ loãng.",
            "2. Al: Kim loại có tính khử mạnh, phản ứng oxi hóa - khử với HNO~3~ loãng giải phóng NO (hoặc N~2~O, N~2~, NH~4~NO~3~). -> *(1)*",
            "3. FeO: Nguyên tử Fe ở mức oxi hóa +2 chưa cực đại, bị oxi hóa lên Fe(+3): 3FeO + 10HNO~3~ -> 3Fe(NO~3~)~3~ + NO^ + 5H~2~O. -> *(2)*",
            "4. Fe(OH)~2~: Nguyên tử Fe ở mức +2, bị oxi hóa lên Fe(+3): 3Fe(OH)~2~ + 10HNO~3~ -> 3Fe(NO~3~)~3~ + NO^ + 8H~2~O. -> *(3)*",
            "5. CaCO~3~: Phản ứng trao đổi tạo CO~2~: CaCO~3~ + 2HNO~3~ -> Ca(NO~3~)~2~ + CO~2~^ + H~2~O (không thay đổi số oxi hóa).",
            "6. K~2~SO~4~: Muối của acid mạnh không phản ứng với HNO~3~ loãng.",
            "=> Tổng cộng có 3 trường hợp xảy ra phản ứng oxi hóa - khử: Al, FeO, Fe(OH)~2~."
        ]
        add_solution(doc, sol_17)
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ......................................................................................................................................................", base_size=10.5)

    # CÂU 18
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 18. Xét trạng thái tự nhiên của sulfur:\n1) Trong tự nhiên sulfur có cả ở dạng đơn chất và hợp chất.\n2) Trong cơ thể người, sulfur có trong thành phần của nhiều protein và enzyme.\n3) Dạng hợp chất sulfur thường có ở các khoáng vật như pyrite (FeS~2~), chu sa (HgS), ...\n4) Dạng đơn chất sulfur thường có ở gần các hồ nước mặn.\nHãy sắp xếp các mệnh đề đúng theo số thứ tự tăng dần (ví dụ 12, 123, ...)")
    if is_teacher:
        sol_18 = [
            "*Trả lời:* *123*",
            "*Lời giải:*",
            "- Mệnh đề (1) ĐÚNG: Sulfur tồn tại trong tự nhiên ở cả hai dạng đơn chất và hợp chất.",
            "- Mệnh đề (2) ĐÚNG: Sulfur là thành phần cấu tạo thiết yếu trong các amino acid (methionine, cysteine), từ đó cấu tạo nên nhiều loại protein và enzyme sinh học trong cơ thể người và động thực vật.",
            "- Mệnh đề (3) ĐÚNG: Dạng hợp chất của sulfur phổ biến trong quặng pyrite (FeS~2~), chu sa (HgS), chalcopyrite (CuFeS~2~), thạch cao (CaSO~4~.2H~2~O)...",
            "- Mệnh đề (4) SAI: Đơn chất sulfur thường có ở vùng đất đá gần miệng núi lửa hoặc các mỏ lưu huỳnh ngầm trong lòng đất; ở hồ nước mặn chủ yếu chứa các muối sulfate hòa tan, không tồn tại đơn chất sulfur.",
            "=> Các mệnh đề đúng gồm: 1, 2, 3 -> Ghép số thứ tự tăng dần là: 123."
        ]
        add_solution(doc, sol_18)
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

    # CÂU 19
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 19. (1,0 điểm) Trình bày cách pha loãng dung dịch sulfuric acid đặc an toàn và nguyên tắc xử lí sơ bộ khi bị bỏng sulfuric acid.")
    if is_teacher:
        sol_19 = [
            "*Lời giải chi tiết:*",
            "1. *Quy tắc pha loãng dung dịch H~2~SO~4~ đặc an toàn:*",
            "   - Rót từ từ acid sulfuric đặc vào cốc đựng sẵn nước dọc theo thân đũa thủy tinh và khuấy đều nhẹ nhàng.",
            "   - *Tuyệt đối KHÔNG làm ngược lại* (không được rót nước vào acid đặc), vì acid sulfuric đặc hấp thụ nước và tỏa một lượng nhiệt rất lớn; khối lượng riêng của nước nhẹ hơn acid nên nước sẽ sôi đột ngột và bắn các giọt acid nóng ra ngoài gây bỏng cực kì nguy hiểm.",
            "2. *Nguyên tắc xử lí sơ bộ khi bị bỏng sulfuric acid:*",
            "   - *Bước 1:* Lập tức cởi bỏ phần quần áo dính acid để tránh acid ngấm sâu vào da thịt.",
            "   - *Bước 2:* Nhanh chóng rửa ngay vùng da bị bỏng dưới vòi nước sạch chảy liên tục, nhẹ nhàng trong thời gian ít nhất 15 - 20 phút để xối rửa trôi acid và làm giảm nhiệt độ bề mặt.",
            "   - *Bước 3:* Rửa lại nhẹ bằng dung dịch NaHCO~3~ loãng (nồng độ khoảng 2% - 5%) nhằm trung hòa lượng acid còn sót lại.",
            "   - *Bước 4:* Che phủ vết bỏng bằng gạc sạch, khô, vô trùng và khẩn trương chuyển nạn nhân đến cơ sở y tế gần nhất để được chữa trị chuyên khoa."
        ]
        add_solution(doc, sol_19)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # CÂU 20
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 20. (1,0 điểm) Cho 100 mL dung dịch KOH x M vào dung dịch (NH~4~)~2~SO~4~ dư, đun nóng nhẹ thu được 2,479 lít khí (đo ở điều kiện chuẩn đkc).\na) Khí thu được có thể nhận biết bằng cách nào?\nb) Tính x?")
    if is_teacher:
        sol_20 = [
            "*Lời giải chi tiết:*",
            "*Phương trình phản ứng:*",
            "2KOH + (NH~4~)~2~SO~4~ -(t^o^)-> K~2~SO~4~ + 2NH~3~^ + 2H~2~O",
            "(Phương trình ion rút gọn: OH^-^ + NH~4~^+^ -(t^o^)-> NH~3~^ + H~2~O)",
            "*a) Nhận biết khí NH~3~ sinh ra:*",
            "- Khí sinh ra là ammonia (NH~3~), có mùi khai xốc rất đặc trưng.",
            "- Cách 1: Đặt mẩu giấy quỳ tím ẩm lên miệng ống nghiệm -> Quỳ tím ẩm chuyển sang *màu xanh*.",
            "- Cách 2: Dùng đũa thủy tinh nhúng vào dung dịch acid HCl đặc rồi đưa lại gần miệng ống nghiệm -> Xuất hiện làn *khói trắng* dày đặc (hạt tinh thể muối NH~4~Cl: NH~3~ + HCl -> NH~4~Cl).",
            "*b) Tính giá trị x:*",
            "- Ở điều kiện chuẩn (đkc): n_NH~3~ = 2,479 / 24,79 = 0,1 mol.",
            "- Theo phương trình hóa học: n_KOH = n_NH~3~ = 0,1 mol.",
            "- Đổi thể tích dung dịch: V_dd = 100 mL = 0,1 L.",
            "- Nồng độ mol của dung dịch KOH: x = C_M(KOH) = n / V = 0,1 / 0,1 = *1 M* (hoặc 1,0 M)."
        ]
        add_solution(doc, sol_20)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # CÂU 21
    p = doc.add_paragraph()
    render_question_prompt(p, "Câu 21. (1,0 điểm) Cho cân bằng hóa học sau trong pha khí:\n2SO~2~(g) + O~2~(g) <=(xt, t^o^)=> 2SO~3~(g)   \\Delta_r H^0_298 < 0\nCân bằng trên dịch chuyển theo chiều nào khi thay đổi các điều kiện sau (có giải thích ngắn gọn)?\na) Tăng nhiệt độ.\nb) Tăng áp suất chung của hệ phản ứng.\nc) Dùng thêm chất xúc tác V~2~O~5~.\nd) Tăng nồng độ SO~2~.")
    if is_teacher:
        sol_21 = [
            "*Lời giải chi tiết:*",
            "Áp dụng nguyên lí chuyển dịch cân bằng Le Chatelier:",
            "*a) Khi tăng nhiệt độ:*",
            "Phản ứng thuận tỏa nhiệt (\\Delta H < 0). Khi tăng nhiệt độ, cân bằng sẽ dịch chuyển theo chiều chống lại sự tăng nhiệt độ (chiều thu nhiệt, \\Delta H > 0) => Cân bằng dịch chuyển theo *chiều nghịch*.",
            "*b) Khi tăng áp suất chung của hệ:*",
            "- Số mol khí ở vế trái = 2 + 1 = 3 mol.",
            "- Số mol khí ở vế phải = 2 mol.",
            "Khi tăng áp suất, cân bằng dịch chuyển theo chiều làm giảm số mol khí để làm giảm áp suất của hệ => Cân bằng dịch chuyển theo *chiều thuận*.",
            "*c) Khi dùng thêm chất xúc tác V~2~O~5~:*",
            "Chất xúc tác làm tăng đồng thời tốc độ của cả phản ứng thuận và phản ứng nghịch với mức độ như nhau, giúp hệ mau đạt trạng thái cân bằng chứ không làm nồng độ các chất ở trạng thái cân bằng thay đổi => *Cân bằng không dịch chuyển*.",
            "*d) Khi tăng nồng độ khí SO~2~:*",
            "Khi tăng nồng độ chất phản ứng (SO~2~), cân bằng dịch chuyển theo chiều làm giảm bớt nồng độ SO~2~ (tiêu thụ SO~2~) => Cân bằng dịch chuyển theo *chiều thuận*."
        ]
        add_solution(doc, sol_21)
    else:
        add_student_writing_lines(doc, num_lines=8)

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
    doc_hs = build_exam_document(is_teacher=False)
    file_hs = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_11_Ma102_HocSinh.docx")
    doc_hs.save(file_hs)
    print(f"[OK] Đã tạo thành công Bản Học sinh: {file_hs}")

    # Tạo thêm bản không có hậu tố _HocSinh theo quy cách
    file_hs_std = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_11_Ma102.docx")
    doc_hs.save(file_hs_std)
    print(f"[OK] Đã tạo thành công Bản Đề Chuẩn: {file_hs_std}")

    # 2. Bản Giáo viên
    doc_gv = build_exam_document(is_teacher=True)
    file_gv = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_11_Ma102_GiaoVien.docx")
    doc_gv.save(file_gv)
    print(f"[OK] Đã tạo thành công Bản Giáo viên: {file_gv}")


if __name__ == "__main__":
    main()
