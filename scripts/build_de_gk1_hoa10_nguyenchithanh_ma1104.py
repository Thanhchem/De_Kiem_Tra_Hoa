# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì I Hóa Học 10 - THPT Nguyễn Chí Thanh - Mã đề 1104
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản Word:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_10_NguyenChiThanh_Ma1104_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_10_NguyenChiThanh_Ma1104_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
IMG_PHO_KHOI = r"C:\Antigravity_Thanh\San_Pham\images_hoa10\pho_khoi_cu_ma1104.png"

def add_quick_answer_table(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (16 CÂU)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (16 CÂU - 4,0 ĐIỂM)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=16)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    p1_ans = ["D", "C", "A", "B", "A", "B", "C", "D", "B", "C", "A", "D", "A", "C", "A", "C"]
    col_w = Cm(1.12)
    for i in range(16):
        cell_q = tbl1.cell(0, i)
        cell_a = tbl1.cell(1, i)
        cell_q.width = col_w
        cell_a.width = col_w

        pq = cell_q.paragraphs[0]
        pq.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pq.paragraph_format.space_before = Pt(1)
        pq.paragraph_format.space_after = Pt(1)
        rq = pq.add_run(f"C{i+1}")
        set_run_style(rq, font_size=9, bold=True)

        pa = cell_a.paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pa.paragraph_format.space_before = Pt(1)
        pa.paragraph_format.space_after = Pt(1)
        ra = pa.add_run(p1_ans[i])
        set_run_style(ra, font_size=9.5, bold=True, color=COLOR_CORRECT)

        shd_q = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F4F8"/>')
        cell_q._tc.get_or_add_tcPr().append(shd_q)

    tblBdr1 = parse_xml(f'<w:tblBdr {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBdr>')
    tbl1._tbl.tblPr.append(tblBdr1)

    # 2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2 CÂU)
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(6)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (2 CÂU - 2,0 ĐIỂM)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=5)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    cols_w = [Cm(2.8), Cm(3.8), Cm(3.8), Cm(3.8), Cm(3.8)]
    headers2 = ["Câu", "Lệnh a)", "Lệnh b)", "Lệnh c)", "Lệnh d)"]
    for j, h in enumerate(headers2):
        cell = tbl2.cell(0, j)
        cell.width = cols_w[j]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(h)
        set_run_style(r, font_size=9.5, bold=True)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F4F8"/>')
        cell._tc.get_or_add_tcPr().append(shd)

    p2_data = [
        ("Câu 1", ["ĐÚNG", "ĐÚNG", "SAI", "SAI"]),
        ("Câu 2", ["ĐÚNG", "ĐÚNG", "ĐÚNG", "ĐÚNG"])
    ]
    for row_idx, (q_name, answers) in enumerate(p2_data, start=1):
        cell_lbl = tbl2.cell(row_idx, 0)
        cell_lbl.width = cols_w[0]
        p = cell_lbl.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(q_name)
        set_run_style(r, font_size=9.5, bold=True)

        for col_idx, ans in enumerate(answers, start=1):
            cell_ans = tbl2.cell(row_idx, col_idx)
            cell_ans.width = cols_w[col_idx]
            p = cell_ans.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(ans)
            ans_color = COLOR_CORRECT if ans == "ĐÚNG" else COLOR_PRIMARY
            set_run_style(r, font_size=9.5, bold=True, color=ans_color)

    tblBdr2 = parse_xml(f'<w:tblBdr {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBdr>')
    tbl2._tbl.tblPr.append(tblBdr2)

    # 3. PHẦN III: TRẢ LỜI NGẮN (4 CÂU)
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(6)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*PHẦN III: TRẢ LỜI NGẮN (4 CÂU - 1,0 ĐIỂM)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=4)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    p3_cols_w = [Cm(4.5), Cm(4.5), Cm(4.5), Cm(4.5)]
    p3_answers = ["14", "3", "10", "108"]
    for k in range(4):
        cell_q = tbl3.cell(0, k)
        cell_a = tbl3.cell(1, k)
        cell_q.width = p3_cols_w[k]
        cell_a.width = p3_cols_w[k]

        pq = cell_q.paragraphs[0]
        pq.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pq.paragraph_format.space_before = Pt(1)
        pq.paragraph_format.space_after = Pt(1)
        rq = pq.add_run(f"Câu {k+1}")
        set_run_style(rq, font_size=9.5, bold=True)

        pa = cell_a.paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pa.paragraph_format.space_before = Pt(1)
        pa.paragraph_format.space_after = Pt(1)
        ra = pa.add_run(p3_answers[k])
        set_run_style(ra, font_size=10, bold=True, color=COLOR_CORRECT)

        shd_q = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F4F8"/>')
        cell_q._tc.get_or_add_tcPr().append(shd_q)

    tblBdr3 = parse_xml(f'<w:tblBdr {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBdr>')
    tbl3._tbl.tblPr.append(tblBdr3)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(6)
    p_div.paragraph_format.space_after = Pt(6)
    add_p_border_bottom(p_div, color="8BC390", sz="8")


def generate_document(is_teacher=False):
    doc = docx.Document()

    # Thiết lập lề A4
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        section.header_distance = Cm(0.6)
        section.footer_distance = Cm(0.6)

    # Header & Footer
    setup_header_footer(doc)

    # Khung tiêu đề
    build_title_block(
        doc,
        school_line="SỞ GD&ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT NGUYỄN CHÍ THANH",
        exam_title="ĐỀ KIỂM TRA GIỮA HỌC KỲ I – NĂM HỌC 2026 – 2027",
        subject_line="MÔN HOÁ HỌC - LỚP 10",
        num_pages="04",
        duration="45 phút (không kể thời gian phát đề)",
        exam_code="1104",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Br = 80.\n"
                      "Số hiệu nguyên tử: H = 1; Li = 3; C = 6; N = 7; O = 8; F = 9; Ne = 10; Na = 11; Mg = 12; Al = 13; Si = 14; P = 15; S = 16; Cl = 17; Ar = 18; K = 19; Ca = 20; Fe = 26; Cu = 29",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table(doc)

    # ==========================================
    # PHẦN I. TRẮC NGHIỆM NHIỀU LỰA CHỌN (4,0 ĐIỂM)
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn: (16 câu – 4 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 16. Mỗi câu thí sinh chỉ chọn một phương án."
    )

    # Câu 1
    p1 = doc.add_paragraph()
    render_question_prompt(p1, "Câu 1: Các đồng vị của cùng một nguyên tố hóa học là những nguyên tử có")
    render_options_no_table(doc, [
        "A. cùng số neutron nhưng khác nhau về số proton.",
        "B. cùng số khối A nhưng khác nhau về số proton.",
        "C. cùng tổng số hạt proton và neutron nhưng khác nhau về số electron.",
        "*D. cùng số proton nhưng khác nhau về số neutron.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Đồng vị của cùng một nguyên tố hóa học là những nguyên tử có cùng số proton (cùng điện tích hạt nhân, cùng số hiệu nguyên tử Z) nhưng khác nhau về số neutron, do đó số khối A khác nhau.\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 2
    p2 = doc.add_paragraph()
    render_question_prompt(p2, "Câu 2: Cấu hình electron nguyên tử của potassium (Z = 19) là 1s^2^2s^2^2p^6^3s^2^3p^6^4s^1^. Vị trí của potassium trong bảng tuần hoàn là")
    render_options_no_table(doc, [
        "A. ô số 19, chu kì 4, nhóm IB.",
        "B. ô số 19, chu kì 3, nhóm IA.",
        "*C. ô số 19, chu kì 4, nhóm IA.*",
        "D. ô số 19, chu kì 1, nhóm IVA."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Ô số 19: do số hiệu nguyên tử Z = 19 (có 19 electron).\n"
            "- Chu kì 4: do nguyên tử có 4 lớp electron (lớp ngoài cùng n = 4).\n"
            "- Nhóm IA: do là nguyên tố s (electron cuối cùng điền vào phân lớp 4s) và có 1 electron ở lớp ngoài cùng (4s^1^).\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 3
    p3 = doc.add_paragraph()
    render_question_prompt(p3, "Câu 3: Nguyên tử của nguyên tố phosphorus có Z = 15 và A = 31. Số hạt neutron trong nguyên tử phosphorus là")
    render_options_no_table(doc, [
        "*A. 16.*",
        "B. 31.",
        "C. 46.",
        "D. 15."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Số hạt neutron N = A - Z = 31 - 15 = 16.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 4
    p4 = doc.add_paragraph()
    render_question_prompt(p4, "Câu 4: Cho các nguyên tử sau: ^12_6X, ^14_7Y, ^13_6Z, ^16_8T. Hai nguyên tử nào là đồng vị của cùng một nguyên tố hóa học?")
    render_options_no_table(doc, [
        "A. X và Y.",
        "*B. X và Z.*",
        "C. Z và T.",
        "D. Y và T."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Hai nguyên tử ^12_6X và ^13_6Z đều có cùng số proton Z = 6 (cùng là nguyên tố Carbon), khác nhau về số neutron (6 và 7), do đó X và Z là hai đồng vị của cùng một nguyên tố hóa học.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 5
    p5 = doc.add_paragraph()
    render_question_prompt(p5, "Câu 5: Số orbital trong các phân lớp s, p, d lần lượt bằng")
    render_options_no_table(doc, [
        "*A. 1, 3, 5.*",
        "B. 1, 2, 3.",
        "C. 1, 2, 4.",
        "D. 3, 5, 7."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Số orbital (AO) trong các phân lớp:\n"
            "- Phân lớp s có 1 orbital.\n"
            "- Phân lớp p có 3 orbital.\n"
            "- Phân lớp d có 5 orbital.\n"
            "- Phân lớp f có 7 orbital.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 6
    p6 = doc.add_paragraph()
    render_question_prompt(p6, "Câu 6: Số electron tối đa có thể chứa trong lớp thứ n (với n ≤ 4) được tính theo công thức là")
    render_options_no_table(doc, [
        "A. 2n + 1.",
        "*B. 2n^2^.*",
        "C. 2n.",
        "D. n^2^."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Trong lớp thứ n (với n ≤ 4), số orbital tối đa là n^2^, mỗi orbital chứa tối đa 2 electron nên số electron tối đa của lớp n là 2n^2^.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 7
    p7 = doc.add_paragraph()
    render_question_prompt(p7, "Câu 7: Kí hiệu chung của nguyên tử là ^A_ZX, trong đó A và Z lần lượt biểu diễn cho")
    render_options_no_table(doc, [
        "A. số hiệu nguyên tử và số khối.",
        "B. số khối và số neutron.",
        "*C. số khối và số hiệu nguyên tử.*",
        "D. số neutron và số hiệu nguyên tử."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Trong kí hiệu nguyên tử ^A_ZX, A nằm ở góc trên bên trái biểu diễn cho số khối, Z nằm ở góc dưới bên trái biểu diễn cho số hiệu nguyên tử (số proton).\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 8
    p8 = doc.add_paragraph()
    render_question_prompt(p8, "Câu 8: Bảng tuần hoàn các nguyên tố hoá học được sắp xếp không tuân theo nguyên tắc nào sau đây?")
    render_options_no_table(doc, [
        "A. Các nguyên tố mà nguyên tử có cùng số lớp electron được xếp vào một hàng.",
        "B. Các nguyên tố mà nguyên tử có cùng số electron hoá trị được xếp vào một cột.",
        "C. Các nguyên tố được sắp xếp theo chiều tăng dần của điện tích hạt nhân nguyên tử.",
        "*D. Các nguyên tố được sắp xếp theo chiều tăng dần khối lượng nguyên tử.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Bảng tuần hoàn hiện đại được sắp xếp theo chiều tăng dần của điện tích hạt nhân nguyên tử (Z), KHÔNG PHẢI theo chiều tăng dần khối lượng nguyên tử (nguyên tắc cũ thời Mendeleev). Trong bảng tuần hoàn hiện đại có một số ngoại lệ khối lượng nguyên tử không tăng dần như Ar (Z=18, M=39,95) đứng trước K (Z=19, M=39,10), Te (Z=52, M=127,6) đứng trước I (Z=53, M=126,9).\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 9
    p9 = doc.add_paragraph()
    render_question_prompt(p9, "Câu 9: Orbital nguyên tử (AO) là")
    render_options_no_table(doc, [
        "A. quỹ đạo chuyển động có hình tròn xác định của electron xung quanh hạt nhân.",
        "*B. khu vực không gian xung quanh hạt nhân mà tại đó xác suất tìm thấy electron là lớn nhất.*",
        "C. đám mây hình bầu dục chứa tất cả các electron của nguyên tử.",
        "D. đường tròn trên đó các electron chuyển động với tốc độ không đổi."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Theo mô hình hiện đại, orbital nguyên tử (AO) là khu vực không gian xung quanh hạt nhân nguyên tử mà tại đó xác suất tìm thấy (xác suất có mặt) electron là lớn nhất (khoảng 90%).\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 10
    p10 = doc.add_paragraph()
    render_question_prompt(p10, "Câu 10: Nguyên tố hóa học là tập hợp các nguyên tử có cùng")
    render_options_no_table(doc, [
        "A. số khối A.",
        "B. tổng số proton và neutron.",
        "*C. số proton.*",
        "D. số neutron."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Nguyên tố hóa học là tập hợp các nguyên tử có cùng điện tích hạt nhân (cùng số proton trong hạt nhân, cùng số hiệu nguyên tử Z).\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 11
    p11 = doc.add_paragraph()
    render_question_prompt(p11, "Câu 11: Đồng có hai đồng vị bền trong tự nhiên là ^63_29Cu và ^65_29Cu. Nguyên tử khối trung bình của đồng là 63,55 (điện tích z của các ion đồng vị bền đều bằng +1). Hình vẽ phổ khối nào dưới đây là đúng?")
    if os.path.exists(IMG_PHO_KHOI):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(3)
        p_img.paragraph_format.space_after = Pt(3)
        p_img.add_run().add_picture(IMG_PHO_KHOI, width=Inches(5.6))
    render_options_no_table(doc, [
        "*A. Hình A.*",
        "B. Hình B.",
        "C. Hình C.",
        "D. Hình D."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Gọi x là phần trăm số nguyên tử của đồng vị ^63Cu, ta có phần trăm của ^65Cu là (100 - x)%.\n"
            "- Áp dụng công thức nguyên tử khối trung bình:\n"
            "  Ā = [63x + 65(100 - x)] / 100 = 63,55\n"
            "  ➔ 6500 - 2x = 6355 ➔ 2x = 145 ➔ x = 72,5%.\n"
            "- Vậy đồng vị ^63Cu chiếm 72,5% và đồng vị ^65Cu chiếm 27,5%.\n"
            "- Với z = +1 thì m/z lần lượt là 63/1 = 63 và 65/1 = 65.\n"
            "- Trên phổ khối lượng: tại m/z = 63 có pic cao 72,5% (trên 70%), tại m/z = 65 có pic cao 27,5% (dưới 30%).\n"
            "➔ Hình A thể hiện chính xác tỉ lệ này.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 12
    p12 = doc.add_paragraph()
    render_question_prompt(p12, "Câu 12: Trong nguyên tử, loại hạt nào có khối lượng không đáng kể so với các hạt còn lại?")
    render_options_no_table(doc, [
        "A. Neutron.",
        "B. Proton.",
        "C. Neutron và electron.",
        "*D. Electron.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Khối lượng của electron rất nhỏ (m~e~ ≈ 0,00055 amu ≈ 9,109 × 10^-31^ kg), nhỏ hơn khối lượng của proton và neutron khoảng 1836 lần, nên khối lượng electron được coi là không đáng kể so với khối lượng hạt nhân.\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 13
    p13 = doc.add_paragraph()
    render_question_prompt(p13, "Câu 13: Bảng tuần hoàn hiện nay gồm có bao nhiêu chu kì?")
    render_options_no_table(doc, [
        "*A. 7.*",
        "B. 18.",
        "C. 8.",
        "D. 6."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Bảng tuần hoàn các nguyên tố hóa học hiện nay gồm có 7 chu kì, được đánh số từ 1 đến 7 tương ứng với số lớp electron trong nguyên tử.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 14
    p14 = doc.add_paragraph()
    render_question_prompt(p14, "Câu 14: Khối lượng của nguyên tử magnesium là 39,8271·10^-27^ kg. Khối lượng của magnesium theo amu là")
    render_options_no_table(doc, [
        "A. 24,000.",
        "B. 23,985·10^-3^.",
        "*C. 23,978.*",
        "D. 66,133·10^-51^."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Theo quy đổi đơn vị: 1 amu ≈ 1,661 × 10^-27^ kg.\n"
            "Khối lượng của nguyên tử Mg theo amu là:\n"
            "m(Mg) = (39,8271 × 10^-27^) / (1,661 × 10^-27^) ≈ 23,9777 ≈ 23,978 amu.\n"
            "(Lưu ý: phương án B là 23,985·10^-3^ bị thừa bậc số mũ 10^-3^ nên không chính xác).\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 15
    p15 = doc.add_paragraph()
    render_question_prompt(p15, "Câu 15: Hạt nhân của hầu hết các nguyên tử đều được tạo nên từ các loại hạt là")
    render_options_no_table(doc, [
        "*A. proton và neutron.*",
        "B. neutron và electron.",
        "C. electron, proton và neutron.",
        "D. electron và proton."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Hạt nhân của hầu hết các nguyên tử được tạo thành từ 2 loại hạt là proton (mang điện tích dương) và neutron (không mang điện), ngoại trừ đồng vị ^1_1H chỉ chứa 1 proton mà không có neutron.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 16
    p16 = doc.add_paragraph()
    render_question_prompt(p16, "Câu 16: Cấu hình electron của nguyên tử Ca (Z = 20) ở trạng thái cơ bản là")
    render_options_no_table(doc, [
        "A. 1s^2^2s^2^2p^6^3s^2^3p^6^4p^2^.",
        "B. 1s^2^2s^2^2p^6^3s^2^3p^6^4s^2^4p^1^.",
        "*C. 1s^2^2s^2^2p^6^3s^2^3p^6^4s^2^.*",
        "D. 1s^2^2s^2^2p^6^3s^2^3p^6^4s^1^."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Với Z = 20, các electron được phân bố lần lượt vào các mức năng lượng theo thứ tự tăng dần:\n"
            "1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^6^ 4s^2^ (viết gọn là [Ar]4s^2^).\n"
            "-> *Đáp án đúng: C*"
        )

    # ==========================================
    # PHẦN II. TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM)
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai: (2 câu – 2 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu thí sinh chọn đúng hoặc sai."
    )

    # Câu 1 Phần II
    p2_1 = doc.add_paragraph()
    render_question_prompt(p2_1, "Câu 1: Cho nguyên tử iron (Fe) có kí hiệu nguyên tử là ^56_26Fe. Biết nguyên tử Fe có cấu hình electron viết gọn là [Ar]3d^6^4s^2^.")
    p2_1_stmts = [
        ("a) Nguyên tử Fe có 26 hạt proton, 26 hạt electron và 30 hạt neutron.", "Đúng"),
        ("b) Nguyên tố Fe thuộc chu kì 4, nhóm VIIIB trong bảng tuần hoàn.", "Đúng"),
        ("c) Lớp vỏ nguyên tử Fe có 8 electron ở lớp ngoài cùng.", "Sai"),
        ("d) Cation Fe^2+^ được tạo thành khi nguyên tử Fe nhường đi 2 electron ở phân lớp 3d.", "Sai")
    ]
    for stmt_text, ans in p2_1_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt_text, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Ý a) ĐÚNG: Từ kí hiệu ^56_26Fe ta có: số proton p = 26; số electron e = p = 26; số neutron n = A - Z = 56 - 26 = 30 hạt.\n"
            "- Ý b) ĐÚNG: Cấu hình electron của Fe: 1s^2^2s^2^2p^6^3s^2^3p^6^3d^6^4s^2^. Có 4 lớp electron ➔ thuộc chu kì 4. Là nguyên tố d có tổng số electron hóa trị ở phân lớp 3d và 4s là 6 + 2 = 8 ➔ thuộc nhóm VIIIB.\n"
            "- Ý c) SAI: Lớp ngoài cùng của Fe là lớp thứ 4 (n = 4) chỉ có phân lớp 4s^2^, tức là chỉ có 2 electron ở lớp ngoài cùng (8 là tổng số electron hóa trị ở 3d và 4s).\n"
            "- Ý d) SAI: Khi Fe nhường electron tạo cation Fe^2+^, nó sẽ nhường 2 electron ở lớp ngoài cùng trước (phân lớp 4s) chứ không phải phân lớp 3d: Fe ([Ar]3d^6^4s^2^) ➔ Fe^2+^ ([Ar]3d^6^) + 2e.\n"
            "-> *Đáp án: a) ĐÚNG - b) ĐÚNG - c) SAI - d) SAI*"
        )

    # Câu 2 Phần II
    p2_2 = doc.add_paragraph()
    render_question_prompt(p2_2, "Câu 2: Nguyên tố X được dùng để chế tạo hợp kim nhẹ, bền, dùng trong nhiều lĩnh vực: hàng không, ô tô, xây dựng, hàng tiêu dùng... Nguyên tố Y ở dạng YO~4~^3-^, đóng vai trò quan trọng trong các phân tử sinh học như DNA và RNA. Các tế bào sống sử dụng YO~4~^3-^ để vận chuyển năng lượng. Nguyên tử của nguyên tố X có cấu hình electron kết thúc ở phân lớp 3p^1^. Nguyên tử của nguyên tố Y có cấu hình electron kết thúc ở phân lớp 3p^3^.")
    p2_2_stmts = [
        ("a) Cấu hình electron của X là [Ne]3s^2^3p^1^.", "Đúng"),
        ("b) Cấu hình electron của Y là [Ne]3s^2^3p^3^.", "Đúng"),
        ("c) Y là phi kim.", "Đúng"),
        ("d) X là kim loại.", "Đúng")
    ]
    for stmt_text, ans in p2_2_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt_text, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Ý a) ĐÚNG: Nguyên tố X có cấu hình electron kết thúc ở 3p^1^ nên cấu hình đầy đủ là 1s^2^2s^2^2p^6^3s^2^3p^1^ hay [Ne]3s^2^3p^1^ (Aluminium, Z = 13).\n"
            "- Ý b) ĐÚNG: Nguyên tố Y có cấu hình electron kết thúc ở 3p^3^ nên cấu hình đầy đủ là 1s^2^2s^2^2p^6^3s^2^3p^3^ hay [Ne]3s^2^3p^3^ (Phosphorus, Z = 15; anion PO~4~^3-^ trong DNA/RNA/ATP).\n"
            "- Ý c) ĐÚNG: Y (Phosphorus, Z = 15) có 5 electron ở lớp ngoài cùng (3s^2^3p^3^) nên Y là một phi kim.\n"
            "- Ý d) ĐÚNG: X (Aluminium, Z = 13) có 3 electron ở lớp ngoài cùng (3s^2^3p^1^) nên X là một kim loại.\n"
            "-> *Đáp án: a) ĐÚNG - b) ĐÚNG - c) ĐÚNG - d) ĐÚNG*"
        )

    # ==========================================
    # PHẦN III. TRẢ LỜI NGẮN (1,0 ĐIỂM)
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm trả lời ngắn: (4 câu – 1 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4."
    )

    # Câu 1 Phần III
    p3_1 = doc.add_paragraph()
    render_question_prompt(p3_1, "Câu 1: Aluminium là kim loại phổ biến nhất trên vỏ Trái Đất, được sử dụng trong các ngành xây dựng, ngành điện hoặc sản xuất đồ gia dụng. Hạt nhân của nguyên tử aluminium có điện tích bằng +13 và số khối bằng 27. Số neutron có trong nguyên tử aluminium là bao nhiêu?")
    p3_1_ans = doc.add_paragraph()
    if is_teacher:
        render_rich_text(p3_1_ans, "*Trả lời: 14*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "- Điện tích hạt nhân bằng +13 ➔ Số proton Z = 13.\n"
            "- Số khối A = 27.\n"
            "- Số hạt neutron: N = A - Z = 27 - 13 = 14.\n"
            "-> *Đáp số: 14*"
        )
    else:
        render_rich_text(p3_1_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # Câu 2 Phần III
    p3_2 = doc.add_paragraph()
    render_question_prompt(p3_2, "Câu 2: Cho các nguyên tố hóa học có số hiệu nguyên tử lần lượt là: Z~1~ = 7; Z~2~ = 11; Z~3~ = 12; Z~4~ = 16; Z~5~ = 17; Z~6~ = 18. Trong số các nguyên tố trên, có bao nhiêu nguyên tố là phi kim?")
    p3_2_ans = doc.add_paragraph()
    if is_teacher:
        render_rich_text(p3_2_ans, "*Trả lời: 3*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Viết cấu hình electron và xác định tính chất nguyên tố:\n"
            "- Z~1~ = 7 (N): 1s^2^2s^2^2p^3^ (5e lớp ngoài cùng) ➔ phi kim. (1)\n"
            "- Z~2~ = 11 (Na): [Ne]3s^1^ (1e lớp ngoài cùng) ➔ kim loại.\n"
            "- Z~3~ = 12 (Mg): [Ne]3s^2^ (2e lớp ngoài cùng) ➔ kim loại.\n"
            "- Z~4~ = 16 (S): [Ne]3s^2^3p^4^ (6e lớp ngoài cùng) ➔ phi kim. (2)\n"
            "- Z~5~ = 17 (Cl): [Ne]3s^2^3p^5^ (7e lớp ngoài cùng) ➔ phi kim. (3)\n"
            "- Z~6~ = 18 (Ar): [Ne]3s^2^3p^6^ (8e lớp ngoài cùng) ➔ khí hiếm.\n"
            "Vậy có tất cả 3 nguyên tố là phi kim (Z~1~, Z~4~, Z~5~).\n"
            "-> *Đáp số: 3*"
        )
    else:
        render_rich_text(p3_2_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # Câu 3 Phần III
    p3_3 = doc.add_paragraph()
    render_question_prompt(p3_3, "Câu 3: Nguyên tử của nguyên tố calcium (Z = 20) ở trạng thái cơ bản có bao nhiêu orbital nguyên tử chứa electron?")
    p3_3_ans = doc.add_paragraph()
    if is_teacher:
        render_rich_text(p3_3_ans, "*Trả lời: 10*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "- Cấu hình electron của Ca (Z = 20): 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^6^ 4s^2^.\n"
            "- Đếm số orbital chứa electron ở từng phân lớp:\n"
            "  + Phân lớp 1s^2^: 1 orbital.\n"
            "  + Phân lớp 2s^2^: 1 orbital.\n"
            "  + Phân lớp 2p^6^: 3 orbital.\n"
            "  + Phân lớp 3s^2^: 1 orbital.\n"
            "  + Phân lớp 3p^6^: 3 orbital.\n"
            "  + Phân lớp 4s^2^: 1 orbital.\n"
            "➔ Tổng số orbital nguyên tử chứa electron: 1 + 1 + 3 + 1 + 3 + 1 = 10 orbital.\n"
            "-> *Đáp số: 10*"
        )
    else:
        render_rich_text(p3_3_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # Câu 4 Phần III
    p3_4 = doc.add_paragraph()
    render_question_prompt(p3_4, "Câu 4: Nguyên tố X được sử dụng rộng rãi trong đời sống: đúc tiền, làm đồ trang sức, làm răng giả,... Muối iodide của X được sử dụng nhằm tụ mây tạo ra mưa nhân tạo. Tổng số hạt cơ bản trong nguyên tử nguyên tố X là 155, số hạt mang điện nhiều hơn số hạt không mang điện là 33 hạt. Số khối của nguyên tử X là bao nhiêu?")
    p3_4_ans = doc.add_paragraph()
    if is_teacher:
        render_rich_text(p3_4_ans, "*Trả lời: 108*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "- Gọi số hạt proton, neutron, electron của nguyên tử X lần lượt là p, n, e (với p = e).\n"
            "- Theo đề bài ta có hệ phương trình:\n"
            "  2p + n = 155  (1)\n"
            "  2p - n = 33   (2)\n"
            "- Giải hệ phương trình:\n"
            "  Lấy (1) + (2) ➔ 4p = 188 ➔ p = 47 (X là Bạc - Silver, Ag; muối AgI làm mưa nhân tạo).\n"
            "  Thay p = 47 vào (1) ➔ n = 155 - 2 × 47 = 61.\n"
            "- Số khối của nguyên tử X: A = p + n = 47 + 61 = 108.\n"
            "-> *Đáp số: 108*"
        )
    else:
        render_rich_text(p3_4_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # ==========================================
    # PHẦN IV. TỰ LUẬN (3,0 ĐIỂM)
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN IV. TỰ LUẬN: (3 câu – 3 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 3."
    )

    # Câu 1 Tự luận
    p_tl1 = doc.add_paragraph()
    render_question_prompt(p_tl1, "Câu 1: (1,0 điểm). Nguyên tố boron (B) có nguyên tử khối trung bình là 10,81. Trong tự nhiên, boron có hai đồng vị bền là ^10B và ^11B.\n"
                                   "a) Tính phần trăm số nguyên tử của mỗi đồng vị trong tự nhiên.\n"
                                   "b) Tính phần trăm khối lượng của đồng vị ^11B có trong hợp chất boric acid H~3~BO~3~ (cho biết nguyên tử khối của H = 1; O = 16).")
    if is_teacher:
        add_solution_no_table(doc,
            "a) (0,5 điểm)\n"
            "- Gọi phần trăm số nguyên tử của đồng vị ^10B là x (%) và ^11B là y (%).\n"
            "- Ta có hệ phương trình:\n"
            "  x + y = 100\n"
            "  (10x + 11y) / 100 = 10,81 ➔ 10x + 11y = 1081\n"
            "- Giải hệ: x = 19% và y = 81%.\n"
            "➔ Phần trăm số nguyên tử của ^10B là 19%; của ^11B là 81%.\n\n"
            "b) (0,5 điểm)\n"
            "- Khối lượng mol phân tử của hợp chất H~3~BO~3~ là:\n"
            "  M(H~3~BO~3~) = 3 × 1 + 10,81 + 3 × 16 = 61,81 g/mol.\n"
            "- Trong 1 mol phân tử H~3~BO~3~ có 1 mol nguyên tử B.\n"
            "  Khối lượng của đồng vị ^11B trong 1 mol H~3~BO~3~ là:\n"
            "  m(^11B) = 1 × 81% × 11 = 8,91 gam.\n"
            "- Phần trăm khối lượng của đồng vị ^11B trong H~3~BO~3~:\n"
            "  %m(^11B) = (8,91 / 61,81) × 100% ≈ 14,41%.\n"
            "➔ Đáp số: 14,41%."
        )
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Tự luận
    p_tl2 = doc.add_paragraph()
    render_question_prompt(p_tl2, "Câu 2: (1,0 điểm). Nguyên tử của nguyên tố X có tổng số electron trên các phân lớp p là 8.\n"
                                   "a) Viết cấu hình electron đầy đủ và cấu hình electron theo ô orbital của nguyên tử X.\n"
                                   "b) Xác định vị trí của X trong bảng tuần hoàn (ô số, chu kì, nhóm). Giải thích ngắn gọn.")
    if is_teacher:
        add_solution_no_table(doc,
            "a) (0,5 điểm)\n"
            "- Các phân lớp p được điền electron theo thứ tự năng lượng là 2p (tối đa 6e) rồi đến 3p.\n"
            "- Tổng số electron trên các phân lớp p là 8 ➔ Phân lớp 2p chứa đầy 6e (2p^6^) và phân lớp 3p chứa 8 - 6 = 2e (3p^2^).\n"
            "- Cấu hình electron đầy đủ của X: 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^2^ (tổng Z = 14, Silicon - Si).\n"
            "- Cấu hình electron theo ô orbital (ô lượng tử):\n"
            "  + Ô 1s: [↑↓] (2e ghép đôi)\n"
            "  + Ô 2s: [↑↓] (2e ghép đôi)\n"
            "  + Ba ô 2p: [↑↓][↑↓][↑↓] (6e ghép đôi hoàn toàn)\n"
            "  + Ô 3s: [↑↓] (2e ghép đôi)\n"
            "  + Ba ô 3p: [↑ ][↑ ][  ] (2e độc thân ở 2 ô đầu tiên, 1 ô trống theo quy tắc Hund).\n\n"
            "b) (0,5 điểm)\n"
            "Vị trí của nguyên tố X trong bảng tuần hoàn:\n"
            "- Ô số 14: vì nguyên tử X có Z = 14 (có 14 electron trong vỏ nguyên tử).\n"
            "- Chu kì 3: vì nguyên tử X có 3 lớp electron.\n"
            "- Nhóm IVA: vì X là nguyên tố p (electron cuối cùng điền vào phân lớp 3p) và có 4 electron ở lớp ngoài cùng (3s^2^3p^2^)."
        )
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 3 Tự luận
    p_tl3 = doc.add_paragraph()
    render_question_prompt(p_tl3, "Câu 3: (1,0 điểm). Hợp chất có công thức M~2~X (được sử dụng trong sản xuất xi măng, phân bón) có tổng số hạt là 140. Trong đó số hạt mang điện nhiều hơn số hạt không mang điện là 44. Số khối của M lớn hơn số khối của X là 23. Tổng số hạt trong nguyên tử M nhiều hơn trong nguyên tử X là 34. Xác định công thức phân tử của hợp chất M~2~X, gọi tên tiếng anh và phân loại các nguyên tố M và X theo cấu hình electron (nguyên tố s, p, d, f).")
    if is_teacher:
        add_solution_no_table(doc,
            "- Gọi số hạt proton, neutron, electron của M là p~M~, n~M~, e~M~ (p~M~ = e~M~); của X là p~X~, n~X~, e~X~ (p~X~ = e~X~).\n"
            "- Trong phân tử M~2~X:\n"
            "  + Tổng số hạt: 2(2p~M~ + n~M~) + (2p~X~ + n~X~) = 140  ➔  (4p~M~ + 2p~X~) + (2n~M~ + n~X~) = 140  (1)\n"
            "  + Số hạt mang điện nhiều hơn không mang điện là 44:  (4p~M~ + 2p~X~) - (2n~M~ + n~X~) = 44  (2)\n"
            "  Từ (1) và (2) suy ra:  4p~M~ + 2p~X~ = 92  ➔  2p~M~ + p~X~ = 46  (3)\n"
            "                         2n~M~ + n~X~ = 48  (4)\n\n"
            "- Theo mối liên hệ giữa hai nguyên tử M và X:\n"
            "  + Số khối của M lớn hơn X là 23: (p~M~ + n~M~) - (p~X~ + n~X~) = 23  (5)\n"
            "  + Tổng số hạt trong M nhiều hơn X là 34: (2p~M~ + n~M~) - (2p~X~ + n~X~) = 34  (6)\n"
            "  Lấy (6) trừ (5) ta được:  p~M~ - p~X~ = 11  (7)\n\n"
            "- Kết hợp (3) và (7):\n"
            "  { 2p~M~ + p~X~ = 46\n"
            "  { p~M~ - p~X~ = 11\n"
            "  ➔ 3p~M~ = 57 ➔ p~M~ = 19 (M là Potassium, K).\n"
            "  ➔ p~X~ = 19 - 11 = 8 (X là Oxygen, O).\n\n"
            "- Thay vào tính neutron và số khối:\n"
            "  + Từ (5): n~M~ - n~X~ = 23 - (19 - 8) = 12. Kết hợp với (4): 2n~M~ + n~X~ = 48 ➔ 3n~M~ = 60 ➔ n~M~ = 20, n~X~ = 8.\n"
            "  + Số khối A~M~ = 19 + 20 = 39 (K); A~X~ = 8 + 8 = 16 (O).\n\n"
            "- Kết luận:\n"
            "  1. Công thức phân tử hợp chất M~2~X: K~2~O.\n"
            "  2. Tên tiếng Anh của M: Potassium; Tên tiếng Anh của X: Oxygen.\n"
            "  3. Cấu hình electron và phân loại nguyên tố:\n"
            "     + M (Potassium, Z = 19): 1s^2^2s^2^2p^6^3s^2^3p^6^4s^1^ (electron cuối cùng điền vào phân lớp 4s) ➔ là NGUYÊN TỐ s.\n"
            "     + X (Oxygen, Z = 8): 1s^2^2s^2^2p^4^ (electron cuối cùng điền vào phân lớp 2p) ➔ là NGUYÊN TỐ p."
        )
    else:
        add_student_writing_lines(doc, num_lines=12)

    # Chân trang thông báo kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "*---------------------- HẾT ----------------------*", base_size=10.5, base_bold=True)

    p_note_end = doc.add_paragraph()
    p_note_end.paragraph_format.space_before = Pt(2)
    p_note_end.paragraph_format.space_after = Pt(2)
    p_note_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_note_end, "- Thí sinh không được sử dụng tài liệu.\n- Cán bộ coi thi không giải thích gì thêm.", base_size=9.5, base_italic=True)

    # Lưu file
    suffix = "GiaoVien" if is_teacher else "HocSinh"
    out_filename = f"De_Kiem_Tra_GK1_Hoa_10_NguyenChiThanh_Ma1104_{suffix}.docx"
    out_path = os.path.join(OUTPUT_DIR, out_filename)
    doc.save(out_path)
    print(f"Đã tạo thành công file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print("Bắt đầu tạo bộ tài liệu Đề Kiểm Tra Giữa Kì 1 Hóa Học 10 - THPT Nguyễn Chí Thanh - Mã Đề 1104...")
    generate_document(is_teacher=False)
    generate_document(is_teacher=True)
    print("Hoàn thành xuất sắc cả 2 phiên bản Word!")
