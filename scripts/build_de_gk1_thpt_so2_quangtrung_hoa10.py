# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Kì 1 Hóa Học 10 - THPT Số 2 Quang Trung (Gia Lai) - Mã đề 1343 (Năm học 2026 - 2027)
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_10_THPT_So2_QuangTrung_Ma1343.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_10_THPT_So2_QuangTrung_Ma1343_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
IMG_DIR = r"C:\Antigravity_Thanh\San_Pham\images_thpt_so2_quangtrung"

def add_centered_picture(doc, img_name, width_cm=6.5):
    """Chèn hình ảnh căn giữa trang nhã vào tài liệu."""
    img_path = os.path.join(IMG_DIR, img_name)
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run()
        run.add_picture(img_path, width=Cm(width_cm))

def add_quick_answer_table_qt1343(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 1343)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (16 CÂU - 4,0 ĐIỂM)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (4,0 điểm - mỗi câu đúng 0,25 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    # Chia 16 câu thành bảng 2 hàng câu và 2 hàng đáp án (8 câu mỗi bảng để vừa 18cm)
    tbl1 = doc.add_table(rows=4, cols=9)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_1 = [2.0] + [2.0] * 8
    p1_ans_row1 = ["A", "D", "C", "A", "C", "D", "B", "B"]
    p1_ans_row2 = ["A", "B", "D", "C", "B", "D", "C", "A"]

    table_data = [
        ["Câu"] + [str(i) for i in range(1, 9)],
        ["Đ/A"] + p1_ans_row1,
        ["Câu"] + [str(i) for i in range(9, 17)],
        ["Đ/A"] + p1_ans_row2
    ]

    for row_idx, row_content in enumerate(table_data):
        row = tbl1.rows[row_idx]
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

            val = row_content[col_idx]
            if row_idx in [0, 2]:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E8F5E9"/>')
                tcPr.append(shd)
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True)
            else:
                if col_idx == 0:
                    render_rich_text(p, f"*{val}*", base_size=10, base_bold=True)
                else:
                    render_rich_text(p, f"*{val}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)

    # 2. PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (2 CÂU - 2,0 ĐIỂM)
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(4)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm - 1 ý: 0,1đ; 2 ý: 0,25đ; 3 ý: 0,5đ; 4 ý: 1,0đ)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_2 = [2.4, 4.4, 2.8, 2.8, 2.8, 2.8]
    
    p2_data = [
        ["Câu", "Nội dung câu hỏi", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Cấu tạo nguyên tử Lithium", "SAI", "ĐÚNG", "SAI", "SAI"],
        ["Câu 2", "Bảng tuần hoàn & Bán kính (Na, Mg, K)", "ĐÚNG", "ĐÚNG", "SAI", "SAI"]
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

    # 3. PHẦN III: TRẢ LỜI NGẮN (4 CÂU - 2,0 ĐIỂM)
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(6)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*3. Phần III: Trả lời ngắn (2,0 điểm - mỗi câu đúng 0,5 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_3 = [2.4, 3.9, 3.9, 3.9, 3.9]
    p3_answers = ["35,5", "3 (hoặc 2*)", "2", "2"]

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
    render_rich_text(p_tl, "*4. Tóm tắt đáp số Phần Tự luận (2,0 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl_tl = doc.add_table(rows=3, cols=3)
    tbl_tl.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_tl = [2.4, 2.6, 13.0]
    
    tl_summary = [
        ["Câu hỏi", "Điểm số", "Đáp số cốt lõi"],
        ["Câu 1 (TL)", "1,0 điểm", "a) Lớp ngoài cùng gồm phân lớp 2s và 2p; tổng số AO = 4.  b) Ô orbital F: 1s^2 2s^2 2p^5."],
        ["Câu 2 (TL)", "1,0 điểm", "a) % số nguyên tử ^14N (X~1~) = 99,63%.  b) Khối lượng đồng vị X~2~ (^15N) = 3,81 gam."]
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


def build_exam_quangtrung_1343(is_teacher=False, output_path=None):
    doc = docx.Document()
    setup_header_footer(doc)

    # 1. Khối tiêu đề đầu đề thi
    build_title_block(
        doc,
        school_line="SỞ GD & ĐT GIA LAI\nTRƯỜNG THPT SỐ 2 QUANG TRUNG",
        exam_title="ĐỀ KIỂM TRA GIỮA KỲ I, NĂM HỌC 2026 - 2027",
        subject_line="MÔN HÓA HỌC - KHỐI 10",
        num_pages="04",
        duration="45 phút",
        exam_code="1343",
        atomic_masses="C = 12; H = 1; O = 16; Na = 23; Br = 80; N = 14.",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_qt1343(doc)

    # =========================================================================
    # PHẦN I. TRẮC NGHIỆM NHIỀU LỰA CHỌN (4,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN I. TRẮC NGHIỆM NHIỀU LỰA CHỌN (4,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 16. Mỗi câu học sinh chỉ chọn một phương án đúng."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Nguyên tử của một nguyên tố có bốn lớp electron như hình vẽ dưới đây:")
    add_centered_picture(doc, "hinh_cau1_lop_e.png", width_cm=7.5)
    p_q1 = doc.add_paragraph()
    p_q1.paragraph_format.space_before = Pt(1)
    p_q1.paragraph_format.space_after = Pt(1)
    render_rich_text(p_q1, "Trong nguyên tử đã cho, electron thuộc lớp nào có mức năng lượng thấp nhất?")
    opts1 = [
        "*A. Lớp K.*" if is_teacher else "A. Lớp K.",
        "B. Lớp M.",
        "C. Lớp L.",
        "D. Lớp N."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong nguyên tử, các electron ở lớp càng gần hạt nhân liên kết càng chặt chẽ với hạt nhân và có mức năng lượng càng thấp. Lớp K (n = 1) là lớp trong cùng, gần hạt nhân nhất nên các electron thuộc lớp K có mức năng lượng thấp nhất.\n➔ Đáp án đúng: A.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Để so sánh độ acid của nước ép, một nhóm học sinh đưa ra giả thuyết: _\"Nước cam tươi có độ acid cao hơn nước táo\"_. Để kiểm chứng giả thuyết, học sinh đo pH để xác định độ acid của hai mẫu nước ép bằng cùng một thiết bị đã hiệu chuẩn, trong cùng điều kiện thí nghiệm. Kết quả cho thấy pH của hai mẫu bằng nhau trong phạm vi độ chính xác của phép đo. Dựa trên kết quả thực nghiệm, nhận định nào sau đây là phù hợp nhất?")
    opts2 = [
        "A. Cần điều chỉnh kết quả đo để chứng minh nước cam có độ acid cao hơn nước táo.",
        "B. Hai mẫu nước ép có thành phần hóa học và hàm lượng acid hoàn toàn giống nhau.",
        "C. Giả thuyết ban đầu được xác nhận vì nước cam thường có vị chua.",
        "*D. Kết quả chưa ủng hộ giả thuyết ban đầu; cần đánh giá độ tin cậy của phép đo và tiến hành kiểm chứng thêm nếu cần.*" if is_teacher else "D. Kết quả chưa ủng hộ giả thuyết ban đầu; cần đánh giá độ tin cậy của phép đo và tiến hành kiểm chứng thêm nếu cần."
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong phương pháp nghiên cứu khoa học:\n- Không được tự ý thay đổi, chỉnh sửa số liệu thực nghiệm (loại A).\n- Không thể suy diễn hai mẫu có thành phần hoàn toàn giống nhau chỉ qua một phép đo pH (loại B).\n- Không thể dùng cảm tính vị giác để bác bỏ kết quả thực nghiệm khoa học (loại C).\n- Khi kết quả thực nghiệm không phù hợp với giả thuyết ban đầu, nhà khoa học cần đánh giá lại độ tin cậy của phép đo, thiết bị, quy trình và tiến hành thêm các thí nghiệm kiểm chứng độc lập nếu cần.\n➔ Đáp án đúng: D.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Nội dung nào dưới đây thuộc đối tượng nghiên cứu của hóa học?")
    opts3 = [
        "A. Quá trình phát triển của loài người.",
        "B. Tốc độ ánh sáng trong chân không.",
        "*C. Cấu tạo của chất và sự biến đổi của chất.*" if is_teacher else "C. Cấu tạo của chất và sự biến đổi của chất.",
        "D. Sự hình thành hệ Mặt Trời."
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hóa học là ngành khoa học tự nhiên nghiên cứu về thành phần, cấu trúc, tính chất và sự biến đổi của các chất cũng như ứng dụng của chúng.\n- A thuộc Lịch sử / Nhân chủng học / Sinh học.\n- B thuộc Vật lí học.\n- D thuộc Thiên văn học / Vật lí thiên văn.\n➔ Đáp án đúng: C.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Các hạt cấu tạo nên hầu hết hạt nhân của nguyên tử (trừ ^1H) là:")
    opts4 = [
        "*A. Proton và neutron.*" if is_teacher else "A. Proton và neutron.",
        "B. Electron và neutron.",
        "C. Electron, proton và neutron.",
        "D. Electron và proton."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hạt nhân của hầu hết các nguyên tử (trừ đồng vị ^1_1H chỉ chứa 1 proton) được cấu tạo từ hai loại hạt cơ bản: proton (mang điện tích dương +1) và neutron (không mang điện).\n➔ Đáp án đúng: A.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 5. Cơ sở tái chế chai nhựa tạo nguyên liệu cho sản xuất. Vai trò trực tiếp của hóa học trong hoạt động này là")
    opts5 = [
        "A. xác định tuổi cây.",
        "B. đo độ cao địa hình.",
        "*C. xử lí và tái sử dụng vật liệu.*" if is_teacher else "C. xử lí và tái sử dụng vật liệu.",
        "D. dự báo lượng mưa."
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Tái chế chai nhựa là quá trình thu gom, phân loại, làm sạch, khử trùng, biến đổi hóa học/nhiệt dẻo polymer để tái sinh thành vật liệu mới hoặc nguyên liệu sản xuất, thể hiện trực tiếp vai trò của hóa học trong xử lí và tái sử dụng vật liệu bảo vệ môi trường.\n➔ Đáp án đúng: C.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 6. Cho các nguyên tử ~11~X, ~13~Y, ~15~Z và ~17~T có giá trị độ âm điện theo biểu đồ dưới:")
    add_centered_picture(doc, "hinh_cau6_do_am_dien.png", width_cm=9.5)
    p_q6 = doc.add_paragraph()
    p_q6.paragraph_format.space_before = Pt(1)
    p_q6.paragraph_format.space_after = Pt(1)
    render_rich_text(p_q6, "Nhận định nào sau đây đúng?")
    opts6 = [
        "A. X có tính kim loại yếu nhất vì có độ âm điện nhỏ nhất.",
        "B. Y có khả năng hút electron trong liên kết hóa học mạnh hơn Z.",
        "C. T có xu hướng nhường electron mạnh hơn X vì có độ âm điện lớn hơn.",
        "*D. T có tính phi kim mạnh nhất trong 4 nguyên tố.*" if is_teacher else "D. T có tính phi kim mạnh nhất trong 4 nguyên tố."
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích giá trị độ âm điện:\n- Cả 4 nguyên tố đều thuộc chu kì 3: ~11~X (Na: 0,93), ~13~Y (Al: 1,61), ~15~Z (P: 2,19), ~17~T (Cl: 3,16).\n- Độ âm điện đặc trưng cho khả năng hút electron khi tạo liên kết hóa học. Độ âm điện càng lớn thì tính phi kim càng mạnh, tính kim loại càng yếu.\n- T có độ âm điện lớn nhất (3,16) => T có tính phi kim mạnh nhất trong 4 nguyên tố (D đúng).\n- X có độ âm điện nhỏ nhất (0,93) => X có tính kim loại mạnh nhất (A sai).\n- Y (1,61) có độ âm điện nhỏ hơn Z (2,19) nên khả năng hút e của Y kém hơn Z (B sai).\n- Độ âm điện lớn thì có xu hướng nhận e mạnh, không phải nhường e (C sai).\n➔ Đáp án đúng: D.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 7. Việc sản xuất phân bón hóa học (phân đạm, phân lân, phân kali) thể hiện vai trò nào sau đây của hóa học đối với đời sống và sản xuất?")
    opts7 = [
        "A. Chế tạo thiết bị chẩn đoán hình ảnh trong y học.",
        "*B. Góp phần tăng năng suất cây trồng, đảm bảo nguồn lương thực.*" if is_teacher else "B. Góp phần tăng năng suất cây trồng, đảm bảo nguồn lương thực.",
        "C. Dự báo chính xác thời tiết cho mùa vụ.",
        "D. Tạo ra nguồn điện cung cấp cho các nhà máy."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân bón hóa học (phân đạm cung cấp N, phân lân cung cấp P, phân kali cung cấp K) là nguồn dinh dưỡng thiết yếu giúp cây trồng sinh trưởng khỏe mạnh, tăng năng suất nông nghiệp, góp phần đảm bảo an ninh lương thực toàn cầu.\n➔ Đáp án đúng: B.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 8. Một cách tính tương đối khối lượng nguyên tử gần bằng khối lượng hạt nhân là do")
    opts8 = [
        "A. khối lượng hạt nhân bằng khối lượng proton.",
        "*B. Khối lượng electron rất nhỏ nên bỏ qua.*" if is_teacher else "B. Khối lượng electron rất nhỏ nên bỏ qua.",
        "C. Hạt nhân không mang điện.",
        "D. Khối lượng proton gần bằng khối lượng neutron."
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khối lượng của electron rất nhỏ (m_e ≈ 0,00055 amu) so với khối lượng proton (≈ 1 amu) và neutron (≈ 1 amu). Khối lượng toàn bộ lớp vỏ electron chỉ chiếm chưa đầy 0,05% khối lượng nguyên tử nên có thể bỏ qua khi tính toán gần đúng, do đó khối lượng nguyên tử tập trung chủ yếu ở hạt nhân.\n➔ Đáp án đúng: B.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 9. Hình dạng sau là của")
    add_centered_picture(doc, "hinh_cau9_orbital_p.png", width_cm=11.0)
    opts9 = [
        "*A. Orbital p.*" if is_teacher else "A. Orbital p.",
        "B. Orbital f.",
        "C. Orbital d.",
        "D. Orbital s."
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hình vẽ mô tả các orbital nguyên tử gồm hai thùy đối xứng hình số 8 nổi (hình quả tạ đôi) định hướng dọc theo ba trục tọa độ Ox, Oy, Oz tương ứng là các orbital p_x, p_y, p_z.\n➔ Đáp án đúng: A.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 10. Calcium (Ca) là nguyên tố khoáng đa lượng cần thiết cho cơ thể, góp phần hình thành và duy trì độ chắc khỏe của xương, răng. Nguyên tử calcium có 20 electron. Vị trí của calcium trong bảng tuần hoàn là")
    opts10 = [
        "A. ô số 20, chu kì 4, nhóm IIB.",
        "*B. ô số 20, chu kì 4, nhóm IIA.*" if is_teacher else "B. ô số 20, chu kì 4, nhóm IIA.",
        "C. ô số 20, chu kì 3, nhóm VIA.",
        "D. ô số 20, chu kì 4, nhóm VIB."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nguyên tử Ca có Z = 20: Cấu hình electron là 1s^2 2s^2 2p^6 3s^2 3p^6 4s^2 (hay [Ar] 4s^2).\n- Số electron = 20 => Ô số 20.\n- Có 4 lớp electron => Chu kì 4.\n- Có 2 electron hóa trị ở phân lớp s (4s^2), nguyên tố s => Nhóm IIA.\n➔ Đáp án đúng: B.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 11. Phát biểu nào sau đây không đúng ?")
    opts11 = [
        "A. Trong nguyên tử, các electron được sắp xếp thành từng lớp, lớp lại được chia thành các phân lớp.",
        "B. Các electron trong cùng một lớp có mức năng lượng gần bằng nhau.",
        "C. Các phân lớp được kí hiệu bằng chữ cái viết thường s, p, d, f.",
        "*D. Các electron trong cùng một phân lớp có mức năng lượng khác nhau.*" if is_teacher else "D. Các electron trong cùng một phân lớp có mức năng lượng khác nhau."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo lí thuyết cấu tạo vỏ nguyên tử:\n- Các electron trên cùng một lớp có mức năng lượng gần bằng nhau.\n- Các electron trên cùng một phân lớp có mức năng lượng BẰNG NHAU (chứ không phải khác nhau). Do đó phát biểu D là không đúng.\n➔ Đáp án đúng: D.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 12. Magnesium (Mg) là một trong những nguyên tố đa lượng đóng vai trò quan trọng của cơ thể, giúp xương chắc khỏe, tim khỏe mạnh và lượng đường trong máu bình thường. Tỉ lệ phần trăm số nguyên tử các đồng vị của magnesium được xác định theo phổ khối lượng như hình dưới đây (biết rằng điện tích z của các ion đồng vị của magnesium đều bằng +2 và quy ước nguyên tử khối bằng số khối).")
    add_centered_picture(doc, "hinh_cau12_pho_khoi_mg.png", width_cm=9.5)
    p_q12 = doc.add_paragraph()
    p_q12.paragraph_format.space_before = Pt(1)
    p_q12.paragraph_format.space_after = Pt(1)
    render_rich_text(p_q12, "Nhận định nào sau đây là sai")
    opts12 = [
        "A. Dựa vào phương pháp phổ khối lượng trên để xác định được phần trăm số nguyên tử đồng vị trong tự nhiên của magnesium có 3 đồng vị bền.",
        "B. Ba đồng vị của magnesium trên phổ có số khối lần lượt là 24, 25 và 26.",
        "*C. Nguyên tử khối trung bình của magnesium bằng 27,31.*" if is_teacher else "C. Nguyên tử khối trung bình của magnesium bằng 27,31.",
        "D. Số hạt neutron có trong 3 đồng vị lần lượt là 12, 13, 14."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích phổ khối lượng:\n- Do điện tích ion z = +2 nên số khối m = (m/z) × 2:\n  + Peak m/z = 12 => Số khối = 24 (^24Mg) chiếm 78,99%.\n  + Peak m/z = 12,5 => Số khối = 25 (^25Mg) chiếm 10%.\n  + Peak m/z = 13 => Số khối = 26 (^26Mg) chiếm 11,01%.\n- Mg có Z = 12 => Số neutron của 3 đồng vị: 24 - 12 = 12; 25 - 12 = 13; 26 - 12 = 14 (D đúng).\n- Nguyên tử khối trung bình của Mg:\n  A_tb = (24 × 78,99 + 25 × 10 + 26 × 11,01) / 100 = 24,32 amu (không phải 27,31).\n➔ Nhận định sai là: C.")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 13. Nguyên tử M cho 3 electron để tạo ion M^3+^. Ion M^3+^ có cấu hình electron được biểu diễn dưới dạng orbital (AO) như sau :")
    add_centered_picture(doc, "hinh_cau13_orbital_m3.png", width_cm=11.5)
    p_q13 = doc.add_paragraph()
    p_q13.paragraph_format.space_before = Pt(1)
    p_q13.paragraph_format.space_after = Pt(1)
    render_rich_text(p_q13, "Nhận định nào sau đây là đúng")
    opts13 = [
        "A. Ở dạng bền nhất và trạng thái cơ bản, nguyên tử M có 5 electron độc thân.",
        "*B. M là một kim loại.*" if is_teacher else "B. M là một kim loại.",
        "C. Trong ion M^3+^ có số proton bằng số electron.",
        "D. Ở dạng bền nhất, nguyên tử M có 23 electron ở lớp vỏ."
    ]
    render_options_no_table(doc, opts13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Cấu hình ô orbital của ion M^3+^ gồm 9 ô bão hòa (18e) và 5 ô có 1e (5e) => Tổng cộng 23 electron.\nCấu hình electron của M^3+^: 1s^2 2s^2 2p^6 3s^2 3p^6 3d^5 ([Ar] 3d^5).\nNguyên tử M nhường 3e (2e ở 4s và 1e ở 3d) để tạo M^3+^ nên số electron của nguyên tử M là: 23 + 3 = 26 (Iron - Fe).\n- Cấu hình electron của M (Z = 26): [Ar] 3d^6 4s^2.\n- M là nguyên tố d, kim loại chuyển tiếp (Fe) => B đúng.\n- M có cấu hình 3d^6 4s^2 nên có 4 electron độc thân (A sai).\n- Ion M^3+^ có 26 proton và 23 electron (C sai).\n- Nguyên tử M có 26 electron ở lớp vỏ (D sai).\n➔ Đáp án đúng: B.")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 14. Quá trình nào sau đây là sự biến đổi hóa học?")
    opts14 = [
        "A. Đường tan vào nước.",
        "B. Cồn bay hơi trong không khí.",
        "C. Nước đá tan chảy thành nước lỏng.",
        "*D. Đinh sắt để lâu ngày trong không khí ẩm bị gỉ.*" if is_teacher else "D. Đinh sắt để lâu ngày trong không khí ẩm bị gỉ."
    ]
    render_options_no_table(doc, opts14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích các quá trình:\n- A, B, C: Hòa tan, bay hơi, nóng chảy chỉ là sự biến đổi trạng thái vật lí, không sinh ra chất mới.\n- D: Đinh sắt bị oxi hóa bởi O~2~ và H~2~O trong không khí ẩm tạo gỉ sắt (oxit và hydroxide của sắt) => Sinh ra chất mới, là biến đổi hóa học.\n➔ Đáp án đúng: D.")

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 15. Những nguyên tử có cùng số proton nhưng khác nhau về số neutron được gọi là")
    opts15 = [
        "A. các đồng vị của cùng nhiều nguyên tố hóa học.",
        "B. các đồng vị của cùng một nguyên tử hóa học.",
        "*C. các đồng vị của cùng một nguyên tố hóa học.*" if is_teacher else "C. các đồng vị của cùng một nguyên tố hóa học.",
        "D. các đồng phân của cùng một nguyên tố hóa học."
    ]
    render_options_no_table(doc, opts15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Đồng vị là các nguyên tử của cùng một nguyên tố hóa học, có cùng số proton (cùng số hiệu nguyên tử Z) nhưng khác nhau về số neutron (dẫn đến số khối A khác nhau).\n➔ Đáp án đúng: C.")

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 16. Orbital nguyên tử là")
    opts16 = [
        "*A. khu vực không gian xung quanh hạt nhân mà tại đó xác suất có mặt electron lớn nhất.*" if is_teacher else "A. khu vực không gian xung quanh hạt nhân mà tại đó xác suất có mặt electron lớn nhất.",
        "B. đám mây chứa electron có dạng hình cầu.",
        "C. đám mây chứa electron có dạng hình số 8 nổi.",
        "D. quỹ đạo chuyển động của electron quay quanh hạt nhân có kích thước năng lượng xác định."
    ]
    render_options_no_table(doc, opts16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo mô hình hiện đại về cấu tạo nguyên tử: Orbital nguyên tử (AO) là khu vực không gian xung quanh hạt nhân mà tại đó xác suất tìm thấy (xác suất có mặt) electron là lớn nhất (khoảng 90%).\n➔ Đáp án đúng: A.")

    # =========================================================================
    # PHẦN II. TRẮC NGHIỆM ĐÚNG - SAI (2,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN II. TRẮC NGHIỆM ĐÚNG - SAI (2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Lithium (Li) và các hợp chất của nó được dùng trong pin, sản xuất thủy tinh, gốm và một số hợp kim nhẹ. Hình dưới đây mô phỏng cấu tạo của nguyên tử Li (hạt đỏ là proton, hạt xám là neutron, hạt xanh là electron).")
    add_centered_picture(doc, "hinh_p2_cau1_lithium.png", width_cm=7.5)
    p_q1_p2 = doc.add_paragraph()
    p_q1_p2.paragraph_format.space_before = Pt(1)
    p_q1_p2.paragraph_format.space_after = Pt(1)
    render_rich_text(p_q1_p2, "Biết m_proton ≈ 1 amu, m_neutron ≈ 1 amu và 1 amu = 1,661.10^-24^ (g). Hãy cho biết những phát biểu sau là đúng hay sai?")
    stmts1 = [
        ("a", "Điện tích hạt nhân của nguyên tử Lithium bằng -3.", False),
        ("b", "Trong hạt nhân nguyên tử Lithium có 3 proton.", True),
        ("c", "Khối lượng hạt nhân bằng 1,1627.10^-26^ kg và nhỏ hơn rất nhiều khối lượng nguyên tử.", False),
        ("d", "Lithium là nguyên tố kim loại vì nguyên tử có 3 electron lớp ngoài cùng.", False)
    ]
    for label, text, ans_bool in stmts1:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}. {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 1:\n- Ý a) SAI: Hạt nhân mang điện tích dương, gồm 3 hạt proton nên điện tích hạt nhân là +3 (hoặc +3e = +4,806.10^-19 C), không phải -3.\n- Ý b) ĐÚNG: Hạt nhân của nguyên tử Li chứa đúng 3 hạt proton (hạt màu đỏ có dấu +).\n- Ý c) SAI: Khối lượng hạt nhân: m = (3 + 4) amu = 7 × 1,661.10^-27 kg = 1,1627.10^-26 kg. Tuy nhiên, khối lượng nguyên tử tập trung hầu hết ở hạt nhân (khối lượng vỏ electron rất nhỏ không đáng kể), nên khối lượng hạt nhân XẤP XỈ bằng khối lượng nguyên tử chứ không thể \"nhỏ hơn rất nhiều khối lượng nguyên tử\".\n- Ý d) SAI: Cấu hình electron của Li (Z = 3) là 1s^2 2s^1. Lớp ngoài cùng (lớp 2) chỉ có 1 electron (chứ không phải 3 electron lớp ngoài cùng).\n➔ Kết luận: a - S, b - Đ, c - S, d - S.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Bảng tuần hoàn các nguyên tố hóa học được xây dựng dựa trên cấu tạo nguyên tử, giúp hệ thống hóa các nguyên tố và dự đoán sự biến đổi tính chất của chúng. Cho ba nguyên tố Na (Z = 11), Mg (Z = 12) và K (Z = 19).\nHãy xác định các phát biểu sau đây là đúng hay sai.")
    stmts2 = [
        ("a", "Trong bảng tuần hoàn, mỗi nguyên tố được xếp vào một ô nguyên tố theo chiều tăng dần của điện tích hạt nhân.", True),
        ("b", "Trong một chu kì, theo chiều tăng dần của điện tích hạt nhân, bán kính nguyên tử nhìn chung giảm vì số lớp electron không đổi, lực hút của hạt nhân đối với electron ngoài cùng tăng.", True),
        ("c", "Các nguyên tố được sắp xếp theo chiều tăng dần của điện tích hạt nhân, các nguyên tố có cùng số lớp electron được xếp vào cùng một nhóm nguyên tố.", False),
        ("d", "Bán kính nguyên tử tăng dần theo thứ tự K < Na < Mg.", False)
    ]
    for label, text, ans_bool in stmts2:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}. {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 2:\n- Ý a) ĐÚNG: Nguyên tắc sắp xếp: các nguyên tố xếp theo chiều tăng dần của điện tích hạt nhân, mỗi nguyên tố chiếm 1 ô.\n- Ý b) ĐÚNG: Trong cùng chu kì, số lớp e không đổi nhưng điện tích hạt nhân tăng làm tăng lực hút tĩnh điện với electron lớp ngoài cùng, kéo các electron lại gần hạt nhân hơn => bán kính nguyên tử giảm dần.\n- Ý c) SAI: Các nguyên tố có cùng số lớp electron được xếp vào cùng một CHU KÌ (hàng ngang), không phải cùng một nhóm.\n- Ý d) SAI: Na (Z=11) và Mg (Z=12) cùng thuộc chu kì 3: bán kính Mg < Na. K (Z=19, chu kì 4) có 4 lớp electron nên bán kính lớn hơn Na (chu kì 3, 3 lớp electron): Na < K. Do đó thứ tự tăng dần bán kính đúng phải là: Mg < Na < K.\n➔ Kết luận: a - Đ, b - Đ, c - S, d - S.")

    # =========================================================================
    # PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Trong tự nhiên, chlorine có hai đồng vị bền ^35~17~Cl (chiếm 75,5% số nguyên tử) và ^37~17~Cl (chiếm 24,5% số nguyên tử). Tính nguyên tử khối trung bình của chlorine (làm tròn kết quả đến hàng phần mười).")
    if is_teacher:
        add_solution_no_table(doc, "Áp dụng công thức tính nguyên tử khối trung bình:\nA_tb = (35 × 75,5 + 37 × 24,5) / 100 = (2642,5 + 906,5) / 100 = 3549 / 100 = 35,49 amu.\nLàm tròn đến hàng phần mười theo yêu cầu đề bài: 35,5.\n➔ Đáp số: 35,5")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Cho các phát biểu sau:\n(1) Nguyên tố hóa học là tập hợp các nguyên tử có cùng số proton trong hạt nhân.\n(2) Số hiệu nguyên tử bằng tổng số proton và neutron trong hạt nhân.\n(3) Trong kí hiệu nguyên tử ^A~Z~X trong đó Z là số hiệu nguyên tử, A là số khối.\n(4) Nguyên tử khối cho biết khối lượng của nguyên tử đó nặng gấp bao nhiêu lần đơn vị khối lượng nguyên tử.\n(5) Hai nguyên tử có cùng số khối luôn thuộc cùng một nguyên tố hóa học.\nCó bao nhiêu phát biểu đúng?")
    if is_teacher:
        add_solution_no_table(doc, "Phân tích từng phát biểu:\n- (1) ĐÚNG: Định nghĩa chuẩn xác của nguyên tố hóa học.\n- (2) SAI: Số hiệu nguyên tử Z bằng số proton (hoặc số đơn vị điện tích hạt nhân), tổng p + n là số khối A.\n- (3) ĐÚNG: Trong kí hiệu ^A_ZX, Z là số hiệu nguyên tử, A là số khối.\n- (4) ĐÚNG: Nguyên tử khối là khối lượng tương đối của nguyên tử, cho biết khối lượng của một nguyên tử nặng gấp bao nhiêu lần đơn vị khối lượng nguyên tử (1 amu).\n- (5) SAI: Hai nguyên tử có cùng số khối nhưng số proton khác nhau là hai nguyên tố khác nhau (ví dụ: ^40_18Ar và ^40_20Ca đều có A = 40 nhưng là hai nguyên tố khác nhau - đồng lượng tử).\nVậy có 3 phát biểu đúng: (1), (3), (4). (Lưu ý: Nếu một số tài liệu cũ không coi phát biểu (4) là chuẩn do cách định nghĩa 1/12 khối lượng C-12 thì có 2 phát biểu đúng).\n➔ Đáp số: 3")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Cho bốn nguyên tố X, Y, Z và T thuộc các nhóm A:\n- X thuộc chu kì 3, có 1 electron ở lớp ngoài cùng.\n- Y có số hiệu nguyên tử lớn hơn X hai đơn vị.\n- Z cùng nhóm với X và có nhiều hơn X một lớp electron.\n- T cùng chu kì với X, có 5 electron ở lớp ngoài cùng.\nXét các nhận định sau:\n(1) Bán kính nguyên tử giảm dần theo thứ tự Z > X > Y > T.\n(2) Độ âm điện tăng dần theo thứ tự Z < X < Y < T.\n(3) Y có tính kim loại mạnh hơn X vì có điện tích hạt nhân lớn hơn.\nCó bao nhiêu nhận định đúng?")
    if is_teacher:
        add_solution_no_table(doc, "Xác định các nguyên tố:\n- X: Chu kì 3, nhóm IA => X là Na (Z = 11: [Ne] 3s^1).\n- Y: Z_Y = 11 + 2 = 13 => Y là Al (Chu kì 3, nhóm IIIA: [Ne] 3s^2 3p^1).\n- Z: Cùng nhóm IA với X, nhiều hơn 1 lớp electron => Chu kì 4, nhóm IA => Z là K (Z = 19).\n- T: Cùng chu kì 3 với X, có 5e lớp ngoài cùng => Chu kì 3, nhóm VA => T là P (Z = 15: [Ne] 3s^2 3p^3).\nKiểm tra các nhận định:\n- (1) ĐÚNG: Trong chu kì 3: bán kính giảm từ trái sang phải: Na (X) > Al (Y) > P (T). Nhóm IA: K (Z) ở chu kì 4 có bán kính lớn hơn Na (X) ở chu kì 3. Vậy thứ tự giảm dần: Z > X > Y > T.\n- (2) ĐÚNG: Trong chu kì 3: độ âm điện tăng: Na (X) < Al (Y) < P (T). Nhóm IA: độ âm điện giảm từ trên xuống: K (Z) < Na (X). Vậy thứ tự tăng dần: Z < X < Y < T.\n- (3) SAI: X (Na) và Y (Al) cùng chu kì 3, từ trái sang phải tính kim loại giảm dần => X có tính kim loại mạnh hơn Y.\nVậy có 2 nhận định đúng: (1) và (2).\n➔ Đáp số: 2")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Có bao nhiêu nguyên tố là kim loại trong dãy sau: (1) X (Z = 7); (2) Y (Z = 11); (3) T (Z = 13); (4) L (Z = 9); (5) M (Z = 10)?")
    if is_teacher:
        add_solution_no_table(doc, "Xét cấu hình electron và tính chất của từng nguyên tố:\n- (1) X (Z = 7): 1s^2 2s^2 2p^3 => Lớp ngoài cùng có 5e, là phi kim (Nitrogen - N).\n- (2) Y (Z = 11): 1s^2 2s^2 2p^6 3s^1 => Lớp ngoài cùng có 1e, là kim loại (Sodium - Na).\n- (3) T (Z = 13): 1s^2 2s^2 2p^6 3s^2 3p^1 => Lớp ngoài cùng có 3e, là kim loại (Aluminium - Al).\n- (4) L (Z = 9): 1s^2 2s^2 2p^5 => Lớp ngoài cùng có 7e, là phi kim (Fluorine - F).\n- (5) M (Z = 10): 1s^2 2s^2 2p^6 => Lớp ngoài cùng có 8e, là khí hiếm (Neon - Ne).\nTrong dãy có 2 nguyên tố kim loại là Y (Z = 11) và T (Z = 13).\n➔ Đáp số: 2")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # =========================================================================
    # PHẦN IV. TỰ LUẬN (2,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN IV. TỰ LUẬN (2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Học sinh trình bày lời giải trên giấy bài thi."
    )

    # Câu 1 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1 (1,0 điểm): Fluorine (F) là nguyên tố halogen có độ âm điện lớn nhất trong bảng tuần hoàn. Một số hợp chất chứa fluorine được sử dụng trong kem đánh răng để hỗ trợ phòng ngừa sâu răng. Nguyên tử fluorine có số hiệu nguyên tử Z = 9.\na. Lớp electron ngoài cùng của nguyên tử fluorine gồm những phân lớp nào và có tổng số AO là bao nhiêu?\nb. Viết cấu hình electron bằng ô orbital của nguyên tử fluorine ở trạng thái cơ bản.")
    if is_teacher:
        sol_tl1 = (
            "Hướng dẫn giải chi tiết Câu 1 (1,0 điểm):\n"
            "a. (0,5 điểm)\n"
            "- Cấu hình electron của fluorine (Z = 9): 1s^2 2s^2 2p^5. (0,25đ)\n"
            "- Lớp ngoài cùng là lớp thứ 2 (n = 2), gồm 2 phân lớp là: phân lớp 2s và phân lớp 2p. (0,15đ)\n"
            "- Phân lớp 2s có 1 AO, phân lớp 2p có 3 AO. Vậy tổng số AO ở lớp ngoài cùng là: 1 + 3 = 4 AO. (0,1đ)\n\n"
            "b. (0,5 điểm)\n"
            "- Biểu diễn cấu hình electron theo ô orbital ở trạng thái cơ bản:\n"
            "  + Ô 1s chứa 2 electron có spin ngược chiều: [↑↓]\n"
            "  + Ô 2s chứa 2 electron có spin ngược chiều: [↑↓]\n"
            "  + 3 ô 2p chứa 5 electron gồm 2 ô ghép đôi và 1 ô chứa electron độc thân: [↑↓][↑↓][↑ ] (0,5đ)"
        )
        add_solution_no_table(doc, sol_tl1)
        add_centered_picture(doc, "hinh_tl_cau1b_fluorine.png", width_cm=7.0)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 2 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2 (1,0 điểm): Nitrogen (N) là nguyên tố dinh dưỡng thiết yếu đối với thực vật, tham gia cấu tạo protein và nucleic acid. Trong tự nhiên, nitrogen có hai đồng vị bền X~1~ và X~2~. Nguyên tử X~1~ có số proton, neutron và electron đều bằng 7, nguyên tử X~2~ có nhiều hơn nguyên tử X~1~ một neutron. Coi nguyên tử khối của mỗi đồng vị bằng số khối.\na) Tính phần trăm số nguyên tử đồng vị X~1~, biết nguyên tử khối trung bình của nitrogen bằng 14,0037.\nb) Để nghiên cứu khả năng hấp thụ phân đạm của cây trồng, người ta chuẩn bị 100 gam ammonium nitrate (NH~4~NO~3~) tinh khiết. Nitrogen trong ion NH~4~^+^ có thành phần đồng vị giống mẫu nitrogen ở ý a; nitrogen trong ion NO~3~^-^ có đồng vị X~2~ chiếm 20% số nguyên tử. Tính khối lượng đồng vị X~2~ trong 100 gam mẫu phân bón này. Biết hydrogen chỉ gồm đồng vị ^1H, oxygen chỉ gồm đồng vị ^16O. (Làm tròn kết quả theo đơn vị gam đến hàng phần trăm).")
    if is_teacher:
        sol_tl2 = (
            "Hướng dẫn giải chi tiết Câu 2 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "- Đồng vị X~1~ có p = 7, n = 7 => Số khối A~1~ = 7 + 7 = 14 (đồng vị ^14N). (0,1đ)\n"
            "- Đồng vị X~2~ có nhiều hơn X~1~ một neutron => n = 7 + 1 = 8 => Số khối A~2~ = 7 + 8 = 15 (đồng vị ^15N). (0,1đ)\n"
            "- Gọi x là phần trăm số nguyên tử của đồng vị X~1~ (^14N), phần trăm của X~2~ (^15N) là (100 - x)%.\n"
            "- Theo công thức nguyên tử khối trung bình:\n"
            "    A_tb = [14x + 15(100 - x)] / 100 = 14,0037 (0,15đ)\n"
            "    <=> 1500 - x = 1400,37  <=>  x = 99,63%.\n"
            "Vậy phần trăm số nguyên tử của đồng vị X~1~ là 99,63% (và đồng vị X~2~ chiếm 0,37%). (0,15đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Khối lượng mol trung bình của từng ion trong mẫu NH~4~NO~3~:\n"
            "    + Trong ion NH~4~^+^: A_tb(N) = 14,0037 => M(NH~4~^+^) = 14,0037 + 4 × 1 = 18,0037 g/mol.\n"
            "    + Trong ion NO~3~^-^: đồng vị ^15N chiếm 20%, ^14N chiếm 80% => A_tb(N) = 14 × 0,8 + 15 × 0,2 = 14,2 g/mol.\n"
            "      => M(NO~3~^-^) = 14,2 + 3 × 16 = 62,2 g/mol. (0,15đ)\n"
            "- Khối lượng mol phân tử của mẫu phân đạm NH~4~NO~3~:\n"
            "    M(NH~4~NO~3~) = 18,0037 + 62,2 = 80,2037 g/mol.\n"
            "- Số mol NH~4~NO~3~ trong 100 gam mẫu: n = 100 / 80,2037 ≈ 1,246825 mol. (0,1đ)\n"
            "- Tổng số mol đồng vị X~2~ (^15N) trong 100 gam mẫu:\n"
            "    n(X~2~) = n_NH4+ × 0,37% + n_NO3- × 20%\n"
            "            = 1,246825 × (0,0037 + 0,20) = 1,246825 × 0,2037 ≈ 0,253978 mol. (0,15đ)\n"
            "- Khối lượng đồng vị X~2~ trong mẫu:\n"
            "    m(X~2~) = n(X~2~) × 15 = 0,253978 × 15 ≈ 3,80967 gam ≈ 3,81 gam. (0,1đ)\n"
            "Vậy khối lượng đồng vị X~2~ trong 100 gam mẫu phân đạm là 3,81 gam."
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=9)

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
    
    file_student = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_THPT_So2_QuangTrung_Ma1343.docx")
    file_student_alias = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_THPT_So2_QuangTrung_Ma1343_HocSinh.docx")
    file_teacher = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_THPT_So2_QuangTrung_Ma1343_GiaoVien.docx")

    print("Đang tạo Bản Học sinh THPT Số 2 Quang Trung...")
    build_exam_quangtrung_1343(is_teacher=False, output_path=file_student)
    build_exam_quangtrung_1343(is_teacher=False, output_path=file_student_alias)

    print("Đang tạo Bản Giáo viên THPT Số 2 Quang Trung...")
    build_exam_quangtrung_1343(is_teacher=True, output_path=file_teacher)

    print("\n[HOÀN TẤT] File Word THPT Số 2 Quang Trung đã được tạo chuẩn 100% FORM ĐỀ KIỂM TRA!")

if __name__ == "__main__":
    main()
