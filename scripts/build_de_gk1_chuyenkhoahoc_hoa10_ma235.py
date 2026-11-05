# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Kì 1 Hóa Học 10 - THPT Chuyên Khoa Học Huế - Mã đề 235 (Năm học 2026 - 2027)
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_10_ChuyenKhoaHocHue_Ma235.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_10_ChuyenKhoaHocHue_Ma235_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
IMG_DIR = r"C:\Antigravity_Thanh\San_Pham\images_chuyenkhoahoc_hoa10_235"

def add_centered_picture(doc, img_name, width_cm=7.5):
    """Chèn hình ảnh căn giữa trang nhã vào tài liệu."""
    img_path = os.path.join(IMG_DIR, img_name)
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run()
        run.add_picture(img_path, width=Cm(width_cm))

def add_quick_answer_table_ckh235(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu Mã đề 235.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 235)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (12 CÂU - 3,0 ĐIỂM)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*1. Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - mỗi câu đúng 0,25 điểm)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=13)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths_1 = [2.4] + [1.3] * 12
    p1_answers = ["B", "C", "D", "B", "D", "B", "C", "A*", "C", "D", "C", "A"]

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

    p_note_c8 = doc.add_paragraph()
    p_note_c8.paragraph_format.space_before = Pt(1)
    p_note_c8.paragraph_format.space_after = Pt(2)
    render_rich_text(p_note_c8, "_(Ghi chú Câu 8: Đề gốc in nhầm 'Sai' thay vì 'Đúng'; A là nhận xét đúng duy nhất, xem phân tích chi tiết bên dưới)._", base_size=9, base_italic=True, base_color=COLOR_GRAY)

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
        ["Câu 1", "Lớp và phân lớp electron", "ĐÚNG", "SAI", "ĐÚNG", "ĐÚNG"],
        ["Câu 2", "Nguyên tử khoáng chất Y (Z = 15)", "SAI", "ĐÚNG", "ĐÚNG", "ĐÚNG"]
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
    p3_answers = ["9", "12", "19% (19)", "3"]

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
        ["Câu 1 (TL)", "1,0 điểm", "a) X: 12e; Y: 9e.  b) Z_X = 12; Z_Y = 9.  c) X: lớp M; Y: lớp L.  d) X: Kim loại; Y: Phi kim."],
        ["Câu 2 (TL)", "1,0 điểm", "a) A_tb(S) = 32,092 amu.  b) Có 1000 nguyên tử ^32S thì có ≈ 44 nguyên tử ^34S."],
        ["Câu 3 (TL)", "1,0 điểm", "Có nghi ngờ. Do A_tb(mẫu) = 12,0098 < 12,0102 (chuẩn) => % ^13C giảm, đặc trưng của doping tổng hợp."]
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


def build_exam_chuyenkhoahoc_235(is_teacher=False, output_path=None):
    doc = docx.Document()
    setup_header_footer(doc)

    # 1. Khối tiêu đề đầu đề thi
    build_title_block(
        doc,
        school_line="TRƯỜNG ĐẠI HỌC KHOA HỌC\nTRƯỜNG THPT CHUYÊN KHOA HỌC HUẾ",
        exam_title="ĐỀ KIỂM TRA GIỮA HỌC KỲ I, NĂM HỌC 2026 - 2027",
        subject_line="Môn: Hóa học – Lớp 10 (CHUNG)",
        num_pages="04",
        duration="45 phút",
        exam_code="235",
        atomic_masses="H=1; Li=7; Be=9; N=14; O=16; Na=23; Mg=24; Al=27; P=31; S=32; Cl=35,5; K=39; Ca=40; Ba=137.",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_ckh235(doc)

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
    render_question_prompt(p, "Câu 1. Biết rằng nguyên tố argon có ba đồng vị khác nhau, ứng với số khối 36; 38 và A. Phần trăm các đồng vị tương ứng lần lượt bằng: 0,34%; 0,06% và 99,6%. Số khối của đồng vị A của nguyên tố argon là bao nhiêu? (cho biết nguyên tử khối trung bình của argon bằng 39,98).")
    opts1 = [
        "A. 39.",
        "*B. 40.*" if is_teacher else "B. 40.",
        "C. 41.",
        "D. 42."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Áp dụng công thức nguyên tử khối trung bình:\nA_tb = (36 × 0,34 + 38 × 0,06 + A × 99,6) / 100 = 39,98\n<=> 12,24 + 2,28 + 99,6 × A = 3998\n<=> 14,52 + 99,6 × A = 3998\n<=> 99,6 × A = 3983,48  <=>  A = 40 (đồng vị ^40Ar).\n➔ Đáp án đúng: B.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Sắp xếp các orbital sau 3s, 3p, 3d, 4s theo thứ tự mức năng lượng tăng dần:")
    opts2 = [
        "A. 3s < 3p < 3d < 4s",
        "B. 3p < 3s < 3d < 4s",
        "*C. 3s < 3p < 4s < 3d*" if is_teacher else "C. 3s < 3p < 4s < 3d",
        "D. 3s < 4s < 3p < 3d"
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo quy tắc Klechkowski, trật tự các mức năng lượng của các orbital nguyên tử sắp xếp theo thứ tự tăng dần là:\n1s < 2s < 2p < 3s < 3p < 4s < 3d < 4p < 5s < 4d...\nDo đó thứ tự đúng là: 3s < 3p < 4s < 3d.\n➔ Đáp án đúng: C.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Hạt nhân của nguyên tử nguyên tố X có 35 hạt, trong đó số hạt không mang điện là 18. Số hạt electron trong nguyên tử nguyên tố X là")
    opts3 = [
        "A. 34.",
        "B. 18.",
        "C. 35.",
        "*D. 17.*" if is_teacher else "D. 17."
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong hạt nhân nguyên tử X:\n- Tổng số hạt: p + n = 35.\n- Số hạt không mang điện là neutron: n = 18.\n=> Số proton: p = 35 - 18 = 17.\nVì nguyên tử trung hòa về điện nên số hạt electron = số hạt proton = 17.\n➔ Đáp án đúng: D.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Nhà khoa học nào đã phát hiện ra neutron?")
    opts4 = [
        "A. Thomson.",
        "*B. Chadwick.*" if is_teacher else "B. Chadwick.",
        "C. Rutherford.",
        "D. Bohr."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Lịch sử tìm ra các hạt cấu tạo nguyên tử:\n- J.J. Thomson phát hiện ra hạt electron (1897).\n- E. Rutherford phát hiện ra hạt nhân (1911) và hạt proton (1918).\n- James Chadwick dùng hạt alpha bắn phá berylium phát hiện ra hạt neutron (1932).\n- Niels Bohr đề xuất mô hình cấu tạo nguyên tử Bohr (1913).\n➔ Đáp án đúng: B.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 5. Nguyên tử M có cấu hình electron 1s^2^2s^2^2p^4^. Phân bố electron trên các orbital là :")
    add_centered_picture(doc, "hinh_cau5_orbital_abcd.png", width_cm=11.5)
    opts5 = [
        "A. [↑↓] [↑↓] [↑↓][↑↓][  ]",
        "B. [↑↓] [↑↓] [↑↓][↑ ][↓ ]",
        "C. [↑↓] [↑↑] [↑↓][↑ ][↑ ]",
        "*D. [↑↓] [↑↓] [↑↓][↑ ][↑ ]*" if is_teacher else "D. [↑↓] [↑↓] [↑↓][↑ ][↑ ]"
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo nguyên lí Pauli và quy tắc Hund:\n- Phân lớp 1s^2 và 2s^2 bão hòa, mỗi AO chứa 1 cặp electron có spin ngược chiều [↑↓].\n- Phân lớp 2p có 4 electron phân bố trên 3 orbital 2p: đầu tiên 3 electron điền vào 3 AO với spin cùng chiều [↑][↑][↑], electron thứ 4 ghép đôi vào orbital đầu tiên tạo thành [↑↓][↑ ][↑ ].\n➔ Đáp án đúng: D.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 6. Cấu hình electron ở trạng thái cơ bản nào sau đây được viết đúng?")
    opts6 = [
        "A. 1s^2^2s^3^2p^6^.",
        "*B. 1s^2^2s^2^2p^6^3s^1^.*" if is_teacher else "B. 1s^2^2s^2^2p^6^3s^1^.",
        "C. 1s^2^2s^2^2p^6^2d^1^.",
        "D. 1s^2^1p^6^2s^2^."
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích các cấu hình:\n- A sai: Phân lớp s chỉ chứa tối đa 2 electron (2s^3 là sai).\n- C sai: Lớp thứ 2 (n = 2) chỉ có 2 phân lớp là 2s và 2p, không có phân lớp 2d.\n- D sai: Lớp thứ 1 (n = 1) chỉ có phân lớp 1s, không có phân lớp 1p.\n- B đúng: 1s^2 2s^2 2p^6 3s^1 là cấu hình electron trạng thái cơ bản của Sodium (Na, Z = 11).\n➔ Đáp án đúng: B.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 7. X là nguyên tố hoá học có trong thành phần của chất có tác dụng oxi hoá và sát khuẩn cực mạnh, thường được sử dụng với mục đích khử trùng và tẩy trắng trong lĩnh vực thuỷ sản, dệt nhuộm, xử lí nước cấp, nước thải, nước bể bơi. Nguyên tử X có tổng số các loại hạt bằng 52, trong đó số hạt mang điện nhiều hơn số hạt không mang điện là 16 hạt. Số khối của nguyên tử X là:")
    opts7 = [
        "A. 31.",
        "B. 32.",
        "*C. 35.*" if is_teacher else "C. 35.",
        "D. 40."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Gọi Z là số proton và N là số neutron của nguyên tử X.\nTheo đề bài ta có hệ phương trình:\n- Tổng số hạt: 2Z + N = 52  (1)\n- Số hạt mang điện nhiều hơn không mang điện: 2Z - N = 16  (2)\nCộng (1) và (2): 4Z = 68 => Z = 17 (Chlorine - Cl).\nThay vào (1): N = 52 - 2 × 17 = 18.\nSố khối của nguyên tử X: A = Z + N = 17 + 18 = 35 (đồng vị ^35_17Cl).\n➔ Đáp án đúng: C.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 8. Cho hình vẽ mô phỏng các nguyên tử với số liệu như sau:")
    add_centered_picture(doc, "hinh_cau8_mo_phong_nguyen_tu.png", width_cm=11.0)
    p_q8 = doc.add_paragraph()
    p_q8.paragraph_format.space_before = Pt(1)
    p_q8.paragraph_format.space_after = Pt(1)
    render_rich_text(p_q8, "Nhận xét nào sau đây Sai? _(Xem phân tích sư phạm)_")
    opts8 = [
        "*A. 1 và 2 là các đồng vị của cùng một nguyên tố hóa học.*" if is_teacher else "A. 1 và 2 là các đồng vị của cùng một nguyên tố hóa học.",
        "B. 1 và 3 là các đồng vị của cùng một nguyên tố hóa học.",
        "C. 1 và 2 là nguyên tử của hai nguyên tố hóa học khác nhau.",
        "D. 1 và 3 có cùng số proton trong hạt nhân."
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích cấu tạo 3 nguyên tử trên hình vẽ:\n- Nguyên tử 1: Hạt nhân có 8 proton, 8 neutron (Oxygen: ^16_8O).\n- Nguyên tử 2: Hạt nhân có 8 proton, 9 neutron (Oxygen: ^17_8O).\n- Nguyên tử 3: Hạt nhân có 9 proton, 10 neutron (Fluorine: ^19_9F).\nĐánh giá các nhận xét:\n+ A: \"1 và 2 là các đồng vị của cùng một nguyên tố hóa học\" là nhận xét ĐÚNG (cùng 8 proton, khác số neutron).\n+ B: \"1 và 3 là các đồng vị...\" là SAI (vì 1 có 8p, 3 có 9p).\n+ C: \"1 và 2 là nguyên tử của hai nguyên tố khác nhau\" là SAI (cùng 8p là cùng nguyên tố Oxygen).\n+ D: \"1 và 3 có cùng số proton\" là SAI (8p khác 9p).\n*Phân tích sư phạm:* Đề gốc in nhầm câu hỏi 'Nhận xét nào sau đây Sai?' thay vì 'Nhận xét nào sau đây Đúng?'. Trong 4 phương án thì A là mệnh đề đúng duy nhất về mặt khoa học, học sinh khoanh chọn A.\n➔ Đáp án chọn: A.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 9. Để đo kích thước của hạt nhân, nguyên tử... hay các hệ vi mô khác, người ta không dùng các đơn vị đo phổ biến đối với các hệ vĩ mô như cm, m, km... mà thường dùng đơn vị đo nanomet (nm) hay angstrom (Å). Cách đổi đơn vị đúng là:")
    opts9 = [
        "A. 1 nm = 10^-10^ m.",
        "B. 1 Å = 10^-9^ m.",
        "*C. 1 nm = 10^-7^ cm.*" if is_teacher else "C. 1 nm = 10^-7^ cm.",
        "D. 1 Å = 10 nm."
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Quy đổi các đơn vị đo kích thước vi mô:\n- 1 nm = 10^-9 m = 10^-9 × 10^2 cm = 10^-7 cm (C đúng).\n- 1 Å = 10^-10 m = 10^-8 cm = 0,1 nm.\n- A sai vì 1 nm = 10^-9 m.\n- B sai vì 1 Å = 10^-10 m.\n- D sai vì 1 Å = 0,1 nm (hay 1 nm = 10 Å).\n➔ Đáp án đúng: C.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 10. Trong các nguyên tử có kí hiệu như sau : ^50~23~X, ^50~22~Y, ^114~48~Z, ^48~22~T, ^116~50~R, các nguyên tử thuộc cùng một nguyên tố hoá học là")
    opts10 = [
        "A. Z và T.",
        "B. X và Y.",
        "C. X, Y và R.",
        "*D. Y và T.*" if is_teacher else "D. Y và T."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Các nguyên tử thuộc cùng một nguyên tố hóa học khi và chỉ khi chúng có cùng số proton (cùng số hiệu nguyên tử Z ở chỉ số dưới bên trái).\nTrong dãy trên: ^50~22~Y và ^48~22~T đều có Z = 22 (Titanium - Ti) nên thuộc cùng một nguyên tố hóa học.\n➔ Đáp án đúng: D.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 11. Nội dung nào dưới đây không thuộc đối tượng nghiên cứu của Hóa học?")
    opts11 = [
        "A. Sự chuyển hóa thức ăn trong hệ tiêu hóa.",
        "B. Giải pháp nhằm giảm thiểu tác hại của mưa acid.",
        "*C. Sự vận chuyển của máu trong hệ tuần hoàn.*" if is_teacher else "C. Sự vận chuyển của máu trong hệ tuần hoàn.",
        "D. Cấu tạo của chất và sự biến đổi của chất."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích đối tượng nghiên cứu:\n- A: Chuyển hóa thức ăn liên quan đến các phản ứng hóa sinh thủy phân glucid, lipid, protein.\n- B: Mưa acid và trung hòa acid là các quá trình biến đổi hóa học và bảo vệ môi trường.\n- D: Cấu tạo và biến đổi của chất là đối tượng nghiên cứu cốt lõi của Hóa học.\n- C: Sự vận chuyển của máu trong hệ tuần hoàn (lực đẩy của tim, huyết áp, cơ chế động lực học chất lưu) là hiện tượng sinh lí - vật lí học, thuộc đối tượng nghiên cứu của Sinh học và Vật lí học.\n➔ Đáp án đúng: C.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 12. Hình dạng orbital s là")
    add_centered_picture(doc, "hinh_cau12_orbital_abcd.png", width_cm=10.5)
    opts12 = [
        "*A. Hình A (dạng hình cầu đối xứng tâm).*" if is_teacher else "A. Hình A (dạng hình cầu đối xứng tâm).",
        "B. Hình B (orbital p_x).",
        "C. Hình C (orbital p_y).",
        "D. Hình D (orbital p_z)."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Orbital s có dạng hình cầu đối xứng trong không gian 3 chiều quanh hạt nhân (Hình A). Các hình B, C, D là các orbital p_x, p_y, p_z có dạng hình số 8 nổi.\n➔ Đáp án đúng: A.")

    # =========================================================================
    # PHẦN II. TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN II (2 điểm). Câu trắc nghiệm đúng sai.",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Trong nguyên tử, các electron được sắp xếp thành từng lớp. Mỗi lớp electron được chia thành các phân lớp.")
    stmts1 = [
        ("a", "Các electron trên cùng một phân lớp có năng lượng bằng nhau.", True),
        ("b", "Lớp K có 1 orbital, lớp L có tổng số orbital là 3.", False),
        ("c", "Lớp L có 2 phân lớp, được kí hiệu là 2s, 2p nên tổng số electron tối đa có trong lớp L là 8.", True),
        ("d", "Các phân lớp được kí hiệu bằng chữ cái viết thường s, p, d, f.", True)
    ]
    for label, text, ans_bool in stmts1:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}. {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 1:\n- Ý a) ĐÚNG: Theo định nghĩa, các electron trên cùng một phân lớp có mức năng lượng bằng nhau.\n- Ý b) SAI: Lớp K (n = 1) có 1 orbital (1s). Lớp L (n = 2) có 1 orbital 2s và 3 orbital 2p => Tổng số orbital của lớp L là 1 + 3 = 4 AO (công thức n^2 = 2^2 = 4), không phải 3.\n- Ý c) ĐÚNG: Lớp L (n = 2) có 2 phân lớp là 2s (chứa tối đa 2e) và 2p (chứa tối đa 6e), tổng số electron tối đa của lớp L là 2 + 6 = 8 electron (công thức 2n^2 = 2 × 2^2 = 8).\n- Ý d) ĐÚNG: Các phân lớp electron được kí hiệu bằng các chữ cái in thường s, p, d, f.\n➔ Kết luận: a - Đ, b - S, c - Đ, d - Đ.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Nguyên tố Y là một khoáng chất thiết yếu đối với sự phát triển của xương và răng, trong cơ thể Y tham gia vào quá trình co cơ, đông máu, điều hòa chức năng thần kinh, thân nhiệt, tái tạo mô, trẻ bảo đảm vận hành của quá trình hoạt động của tim. Trong nguyên tử của nguyên tố Y có tổng số các hạt là 46. Biết số hạt mang điện dương ít hơn số hạt không mang điện là 1. Biết điện tích của 1 hạt electron bằng -1,6.10^-19^ (C).")
    stmts2 = [
        ("a", "Nguyên tử Y có 30 hạt mang điện và 1 electron độc thân ở trạng thái cơ bản.", False),
        ("b", "Nguyên tử Y có 30 hạt mang điện và 3 electron độc thân ở trạng thái cơ bản.", True),
        ("c", "Y là nguyên tố phi kim, có số hiệu nguyên tử là 15, có 9 electron ở phân lớp p.", True),
        ("d", "Nguyên tử Y có điện tích hạt nhân của Y là +2,4.10^-18^ C.", True)
    ]
    for label, text, ans_bool in stmts2:
        p_stmt = doc.add_paragraph()
        p_stmt.paragraph_format.space_before = Pt(1.5)
        p_stmt.paragraph_format.space_after = Pt(1.5)
        ans_str = "Đúng" if ans_bool else "Sai"
        render_statement(p_stmt, f"{label}. {text}", ans=ans_str, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Giải thích chi tiết Câu 2:\n- Gọi Z là số proton, N là số neutron của Y.\n  + Tổng số hạt: 2Z + N = 46  (1)\n  + Số hạt mang điện dương (proton) ít hơn hạt không mang điện (neutron) là 1: N - Z = 1  (2)\n  Lấy (1) trừ (2): 3Z = 45 => Z = 15 (Phosphorus - P). N = 15 + 1 = 16.\n- Tổng số hạt mang điện: p + e = 15 + 15 = 30 hạt.\n- Cấu hình electron của Y (Z = 15): 1s^2 2s^2 2p^6 3s^2 3p^3.\n  Phân lớp 3p^3 có 3 electron độc thân nằm trên 3 AO 3p.\n- Ý a) SAI: Y có 3 electron độc thân (không phải 1).\n- Ý b) ĐÚNG: Y có 30 hạt mang điện và 3 electron độc thân ở trạng thái cơ bản.\n- Ý c) ĐÚNG: Z = 15 có 5e lớp ngoài cùng => phi kim. Tổng số electron ở phân lớp p: 2p^6 + 3p^3 = 6 + 3 = 9 electron.\n- Ý d) ĐÚNG: Điện tích hạt nhân của Y: q = +15 × 1,6.10^-19 C = +2,4.10^-18 C.\n➔ Kết luận: a - S, b - Đ, c - Đ, d - Đ.")

    # =========================================================================
    # PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM)
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN III (2 điểm). Câu trắc nghiệm trả lời ngắn.",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Cho nguyên tố X có 3 lớp electron, lớp thứ 3 có 5 electron. Nguyên tử X có bao nhiêu electron ở phân lớp p?")
    if is_teacher:
        add_solution_no_table(doc, "Nguyên tử X có 3 lớp electron và lớp thứ 3 có 5 electron:\n- Lớp 1 (n = 1): 1s^2\n- Lớp 2 (n = 2): 2s^2 2p^6\n- Lớp 3 (n = 3): 3s^2 3p^3\nCấu hình electron đầy đủ: 1s^2 2s^2 2p^6 3s^2 3p^3.\nCác electron ở phân lớp p gồm: phân lớp 2p (6e) và 3p (3e).\nTổng số electron ở phân lớp p = 6 + 3 = 9.\n➔ Đáp số: 9")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Khí carbon dioxide (CO~2~) là nguyên nhân chính gây nên hiệu ứng nhà kính. Carbon có 2 đồng vị: ^12C, ^13C; oxygen có 3 đồng vị: ^16O, ^17O, ^18O. Số loại phân tử CO~2~ được tạo thành từ các đồng vị trên là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Phân tử CO~2~ có cấu trúc dạng O(1) - C - O(2):\n- Có 2 cách chọn nguyên tử carbon (^12C hoặc ^13C).\n- Phần chứa 2 nguyên tử oxygen (O~2~) chọn từ 3 đồng vị (^16O, ^17O, ^18O):\n  + Hai nguyên tử O cùng đồng vị: 3 cách (16-16, 17-17, 18-18).\n  + Hai nguyên tử O khác đồng vị: C(3, 2) = 3 cách (16-17, 16-18, 17-18).\n  => Tổng số tổ hợp phân nhóm O~2~ = 3 + 3 = 6 cách.\n- Tổng số loại phân tử CO~2~ tạo thành = 2 × 6 = 12 loại phân tử.\n➔ Đáp số: 12")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Boron là nguyên tố có nhiều tác dụng đối với cơ thể người như: làm lành vết thương, điều hòa nội tiết sinh dục, chống viêm khớp,... Do ngọn lửa cháy có màu đặc biệt nên boron vô định hình được dùng làm pháo hoa. Boron có hai đồng vị là ^10B và ^11B, nguyên tử khối trung bình là 10,81. Tính phần trăm đồng vị ^10B?")
    if is_teacher:
        add_solution_no_table(doc, "Gọi x là phần trăm số nguyên tử của đồng vị ^10B, phần trăm của ^11B là (100 - x)%.\nTheo công thức nguyên tử khối trung bình:\nA_tb = [10x + 11(100 - x)] / 100 = 10,81\n<=> 10x + 1100 - 11x = 1081\n<=> 1100 - x = 1081  <=>  x = 1100 - 1081 = 19%.\nVậy phần trăm số nguyên tử của đồng vị ^10B là 19%.\n➔ Đáp số: 19% (hoặc 19)")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ........................................................................................................................", base_size=10.5)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Cho các cấu hình electron của một số nguyên tử nguyên tố như sau:\n(1) 1s^2^2s^2^2p^6^\n(2) 1s^2^2s^2^2p^6^3s^2^\n(3) 1s^2^2s^2^2p^6^3s^2^3p^6^3d^6^4s^2^\n(4) 1s^2^2s^2^2p^6^3s^2^3p^6^3d^10^4s^2^\n(5) 1s^2^2s^2^2p^6^3s^2^3p^4^\n(6) 1s^2^2s^2^2p^6^3s^2^3p^5^\nSố lượng các nguyên tố kim loại trong số các nguyên tố ở trên là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Xác định tính chất kim loại, phi kim, khí hiếm:\n- (1) 1s^2 2s^2 2p^6: Lớp ngoài cùng có 8e => Khí hiếm (Neon - Ne).\n- (2) 1s^2 2s^2 2p^6 3s^2: Lớp ngoài cùng có 2e, nguyên tố s => Kim loại (Magnesium - Mg).\n- (3) 1s^2 2s^2 2p^6 3s^2 3p^6 3d^6 4s^2: Nguyên tố d, kim loại chuyển tiếp (Iron - Fe).\n- (4) 1s^2 2s^2 2p^6 3s^2 3p^6 3d^10 4s^2: Nguyên tố d, kim loại chuyển tiếp (Zinc - Zn).\n- (5) 1s^2 2s^2 2p^6 3s^2 3p^4: Lớp ngoài cùng có 6e, nguyên tố p => Phi kim (Sulfur - S).\n- (6) 1s^2 2s^2 2p^6 3s^2 3p^5: Lớp ngoài cùng có 7e, nguyên tố p => Phi kim (Chlorine - Cl).\nCác nguyên tố kim loại gồm: (2), (3), (4) => Tổng cộng có 3 nguyên tố kim loại.\n➔ Đáp số: 3")
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
    render_question_prompt(p, "Câu 1 (1,0 điểm): Cho cấu hình electron của:\n- Nguyên tử X: 1s^2^2s^2^2p^6^3s^2^\n- Nguyên tử Y: 1s^2^2s^2^2p^5^\na. Mỗi nguyên tử X và Y chứa bao nhiêu electron?\nb. Hãy cho biết số hiệu nguyên tử của X và Y.\nc. Lớp electron nào trong nguyên tử X và Y có mức năng lượng cao nhất?\nd. X và Y là nguyên tố kim loại, phi kim hay khí hiếm?")
    if is_teacher:
        sol_tl1 = (
            "Hướng dẫn giải chi tiết Câu 1 (1,0 điểm):\n"
            "a. (0,25 điểm)\n"
            "- Số electron của nguyên tử X = 2 + 2 + 6 + 2 = 12 electron.\n"
            "- Số electron của nguyên tử Y = 2 + 2 + 5 = 9 electron.\n\n"
            "b. (0,25 điểm)\n"
            "- Do nguyên tử trung hòa về điện nên số hiệu nguyên tử Z = số electron:\n"
            "  + Số hiệu nguyên tử của X là Z_X = 12 (Magnesium - Mg).\n"
            "  + Số hiệu nguyên tử của Y là Z_Y = 9 (Fluorine - F).\n\n"
            "c. (0,25 điểm)\n"
            "- Trong nguyên tử, electron ở lớp càng xa hạt nhân có mức năng lượng càng cao:\n"
            "  + Nguyên tử X có 3 lớp electron (K, L, M) => Lớp có mức năng lượng cao nhất là lớp thứ 3 (lớp M).\n"
            "  + Nguyên tử Y có 2 lớp electron (K, L) => Lớp có mức năng lượng cao nhất là lớp thứ 2 (lớp L).\n\n"
            "d. (0,25 điểm)\n"
            "- Nguyên tử X có 2 electron ở lớp ngoài cùng (3s^2) => X là nguyên tố kim loại.\n"
            "- Nguyên tử Y có 7 electron ở lớp ngoài cùng (2s^2 2p^5) => Y là nguyên tố phi kim."
        )
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2 (1,0 điểm): Một mẫu sulfur có ba đồng vị: ^32S, ^33S và ^34S với tỷ lệ phần trăm các đồng vị như sau:")
    
    # Bảng số liệu đồng vị sulfur
    tbl_s = doc.add_table(rows=2, cols=4)
    tbl_s.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_col_w = [4.5, 4.5, 4.5, 4.5]
    s_data = [
        ["Tỷ lệ", "95,02%", "0,76%", "4,22%"],
        ["Khối lượng đồng vị", "32", "33", "34"]
    ]
    for r_idx, r_val in enumerate(s_data):
        row = tbl_s.rows[r_idx]
        for c_idx, val in enumerate(r_val):
            cell = row.cells[c_idx]
            cell.width = Cm(s_col_w[c_idx])
            tcPr = cell._tc.get_or_add_tcPr()
            borders_xml = f'''<w:tcBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>
            </w:tcBorders>'''
            tcPr.append(parse_xml(borders_xml))
            cp = cell.paragraphs[0]
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if c_idx == 0:
                render_rich_text(cp, f"*{val}*", base_size=10, base_bold=True)
            else:
                render_rich_text(cp, val, base_size=10)

    p_q2_sub = doc.add_paragraph()
    p_q2_sub.paragraph_format.space_before = Pt(3)
    p_q2_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_q2_sub, "a) Tính nguyên tử khối trung bình của sulfur (S).\nb) Nếu có 1000 nguyên tử ^32S thì số nguyên tử ^34S là bao nhiêu? (làm tròn đến hàng đơn vị)")
    if is_teacher:
        sol_tl2 = (
            "Hướng dẫn giải chi tiết Câu 2 (1,0 điểm):\n"
            "a) (0,5 điểm)\n"
            "- Áp dụng công thức tính nguyên tử khối trung bình của sulfur:\n"
            "    A_tb = (32 × 95,02 + 33 × 0,76 + 34 × 4,22) / 100\n"
            "         = (3040,64 + 25,08 + 143,48) / 100\n"
            "         = 3209,20 / 100 = 32,092 amu (xấp xỉ 32,09 amu). (0,5đ)\n\n"
            "b) (0,5 điểm)\n"
            "- Tỉ lệ số nguyên tử của đồng vị ^34S so với đồng vị ^32S là:\n"
            "    N(^34S) / N(^32S) = 4,22% / 95,02% = 4,22 / 95,02.\n"
            "- Khi có 1000 nguyên tử ^32S thì số nguyên tử ^34S tương ứng là:\n"
            "    N(^34S) = 1000 × (4,22 / 95,02) ≈ 44,41 ≈ 44 nguyên tử. (0,5đ)"
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 3 Tự luận
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3 (1,0 điểm): Trong thi đấu thể dục thể thao, có một số vận động viên sử dụng các loại chất kích thích trong thi đấu, gọi là doping, dẫn đến thành tích đạt được của họ không thật sự phù hợp với năng lực vốn có. Một trong các loại doping thường gặp nhất là testosterone tổng hợp. Tỉ lệ giữa hai đồng vị ^12C (98,98%) và ^13C (1,02%) là không đổi với testosterone tự nhiên trong cơ thể. Trong khi testosterone tổng hợp (doping) có phần trăm số nguyên tử đồng vị ^13C ít hơn testosterone tự nhiên. Đây chính là mấu chốt xác định xem vận động viên có sử dụng doping hay không.\nTrong một xét nghiệm với mục đích xác định xem vận động viên có sử dụng doping hay không, xét nghiệm phân tích theo CIR đối với một vận động viên thu được kết quả phần trăm số nguyên tử đồng vị ^12C là x% và ^13C là y%. Từ tỉ lệ này, người ta tính ra được khối lượng nguyên tử trung bình của carbon trong testosterone tổng hợp là 12,0098. Với kết quả nguyên tử khối trung bình của carbon trong mẫu phân tích có giá trị là 12,0098 thu được, em có nghi ngờ vận động viên này sử dụng doping không? Vì sao?")
    if is_teacher:
        sol_tl3 = (
            "Hướng dẫn giải chi tiết Câu 3 (1,0 điểm):\n"
            "- Tính nguyên tử khối trung bình của carbon trong testosterone tự nhiên của cơ thể:\n"
            "    A_tb(tự nhiên) = (12 × 98,98 + 13 × 1,02) / 100\n"
            "                   = (1187,76 + 13,26) / 100 = 1201,02 / 100 = 12,0102 amu. (0,3đ)\n"
            "- So sánh với mẫu xét nghiệm của vận động viên:\n"
            "    A_tb(mẫu) = 12,0098 < A_tb(tự nhiên) = 12,0102. (0,2đ)\n"
            "- Phân tích bản chất hóa học:\n"
            "    Do đồng vị ^13C nặng hơn đồng vị ^12C, nên khi nguyên tử khối trung bình của carbon giảm từ 12,0102 xuống 12,0098 chứng tỏ tỉ lệ phần trăm của đồng vị ^13C trong mẫu xét nghiệm đã bị giảm:\n"
            "    % ^13C(mẫu) = (12,0098 - 12) / (13 - 12) × 100% = 0,98% < 1,02%. (0,2đ)\n"
            "- Theo đề bài, testosterone tổng hợp (doping) có hàm lượng đồng vị ^13C ít hơn testosterone tự nhiên.\n"
            "- Kết luận: CÓ NGHI NGỜ vận động viên này sử dụng doping, vì tỉ lệ đồng vị ^13C trong mẫu phân tích thấp hơn mức tự nhiên trong cơ thể, làm giảm nguyên tử khối trung bình của carbon xuống 12,0098, phù hợp với dấu hiệu sử dụng testosterone ngoại sinh (doping). (0,3đ)"
        )
        add_solution_no_table(doc, sol_tl3)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Cuối đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "*---------- HẾT ----------*", base_size=10.5, base_bold=True)
    p_note = doc.add_paragraph()
    p_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_note, "_- Học sinh không được sử dụng tài liệu; Giám thị không giải thích gì thêm._", base_size=9.5, base_italic=True)

    if output_path:
        doc.save(output_path)
        print(f"[OK] Đã lưu thành công: {output_path}")

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    file_student = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_ChuyenKhoaHocHue_Ma235.docx")
    file_student_alias = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_ChuyenKhoaHocHue_Ma235_HocSinh.docx")
    file_teacher = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_ChuyenKhoaHocHue_Ma235_GiaoVien.docx")

    print("Đang tạo Bản Học sinh THPT Chuyên Khoa Học Huế (Mã đề 235)...")
    build_exam_chuyenkhoahoc_235(is_teacher=False, output_path=file_student)
    build_exam_chuyenkhoahoc_235(is_teacher=False, output_path=file_student_alias)

    print("Đang tạo Bản Giáo viên THPT Chuyên Khoa Học Huế (Mã đề 235)...")
    build_exam_chuyenkhoahoc_235(is_teacher=True, output_path=file_teacher)

    print("\n[HOÀN TẤT] File Word THPT Chuyên Khoa Học Huế (Mã đề 235) đã được tạo chuẩn 100% FORM ĐỀ KIỂM TRA!")

if __name__ == "__main__":
    main()
