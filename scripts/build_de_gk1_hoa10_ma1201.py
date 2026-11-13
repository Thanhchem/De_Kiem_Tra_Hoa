# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì I Hóa Học 10 - Mã đề 1201
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_10_Ma1201_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_10_Ma1201_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
IMG_NEON_PATH = r"C:\Antigravity_Thanh\San_Pham\images_hoa10\pho_khoi_luong_neon_1201.png"

def add_quick_answer_table(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu.
    Bố trí khoa học, trang nhã, đúng kích thước lề 18.0 cm.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (12 CÂU)
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (12 CÂU - 3,0 ĐIỂM)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=12)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    p1_ans = ["B", "A", "A", "C", "D", "C", "C", "C", "B", "C", "B", "A"]
    for i in range(12):
        cell_q = tbl1.cell(0, i)
        cell_a = tbl1.cell(1, i)
        cell_q.width = Cm(1.5)
        cell_a.width = Cm(1.5)

        pq = cell_q.paragraphs[0]
        pq.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pq.paragraph_format.space_before = Pt(1)
        pq.paragraph_format.space_after = Pt(1)
        rq = pq.add_run(f"C{i+1}")
        set_run_style(rq, font_size=9.5, bold=True)

        pa = cell_a.paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pa.paragraph_format.space_before = Pt(1)
        pa.paragraph_format.space_after = Pt(1)
        ra = pa.add_run(p1_ans[i])
        set_run_style(ra, font_size=10, bold=True, color=COLOR_CORRECT)

        # Viền nhẹ và màu nền ô
        shd_q = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F4F8"/>')
        cell_q._tc.get_or_add_tcPr().append(shd_q)

    # Viền bảng 1
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
        ("Câu 1", ["ĐÚNG", "SAI", "SAI", "ĐÚNG"]),
        ("Câu 2", ["ĐÚNG", "SAI", "SAI", "SAI"])
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
    render_rich_text(p_p3, "*PHẦN III: TRẢ LỜI NGẮN (4 CÂU - 2,0 ĐIỂM)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=4)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    p3_cols_w = [Cm(4.5), Cm(4.5), Cm(4.5), Cm(4.5)]
    p3_answers = ["58", "4", "2", "1"]
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
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F4F8"/>')
        cell_q._tc.get_or_add_tcPr().append(shd)

        pa = cell_a.paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pa.paragraph_format.space_before = Pt(1)
        pa.paragraph_format.space_after = Pt(1)
        ra = pa.add_run(p3_answers[k])
        set_run_style(ra, font_size=10.5, bold=True, color=COLOR_CORRECT)

    tblBdr3 = parse_xml(f'<w:tblBdr {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBdr>')
    tbl3._tbl.tblPr.append(tblBdr3)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(8)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


def build_exam_doc(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    # 1. Khối tiêu đề đầu trang
    build_title_block(
        doc,
        school_line="TỔ LÍ - HÓA",
        exam_title="ĐỀ KIỂM TRA GIỮA HỌC KÌ I",
        subject_line="MÔN: HOÁ HỌC - LỚP 10",
        num_pages="04",
        duration="45 phút",
        exam_code="1201",
        atomic_masses="H = 1; B = 10,81; C = 12; N = 14; O = 16; F = 19; Ne = 20,19; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Br = 79,91",
        is_teacher=is_teacher
    )

    # Dòng ghi chú không được dùng BTH
    p_note_bth = doc.add_paragraph()
    p_note_bth.paragraph_format.space_before = Pt(2)
    p_note_bth.paragraph_format.space_after = Pt(4)
    p_note_bth.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_note_bth, "*CHÚ Ý: Học sinh không được sử dụng bảng tuần hoàn các nguyên tố hóa học.*", base_size=10, base_bold=True)

    # Thêm bảng đáp án rút gọn cho giáo viên
    if is_teacher:
        add_quick_answer_table(doc)

    # ==========================================
    # A. PHẦN TRẮC NGHIỆM (7 ĐIỂM)
    # ==========================================
    p_a = doc.add_paragraph()
    p_a.paragraph_format.space_before = Pt(4)
    p_a.paragraph_format.space_after = Pt(2)
    render_rich_text(p_a, "*A. PHẦN TRẮC NGHIỆM (7 ĐIỂM)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # PHẦN I
    add_section_header(
        doc,
        title="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,0 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # Câu 1
    p1 = doc.add_paragraph()
    render_question_prompt(p1, "Câu 1: Đường kính của nguyên tử khoảng 10^2^ pm, trong khi đường kính của hạt nhân khoảng 10^-2^ pm. Đường kính nguyên tử lớn gấp bao nhiêu lần so với đường kính hạt nhân?")
    render_options_no_table(doc, [
        "A. 1000.",
        "*B. 10000.*",
        "C. 100.",
        "D. 10."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, 
            "Tỉ số giữa đường kính nguyên tử và đường kính hạt nhân là:\n"
            "d_nguyên tử / d_hạt nhân = 10^2^ / 10^-2^ = 10^(2 - (-2))^ = 10^4^ = 10 000 (lần).\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 2
    p2 = doc.add_paragraph()
    render_question_prompt(p2, "Câu 2: Số electron tối đa có trong 1 AO p là")
    render_options_no_table(doc, [
        "*A. 2.*",
        "B. 6.",
        "C. 3.",
        "D. 1."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Theo nguyên lí Pauli, trong mỗi một orbital nguyên tử (AO) chỉ chứa tối đa 2 electron có chiều tự quay ngược nhau.\n"
            "Chú ý: Phân lớp p gồm 3 AO p nên chứa tối đa 6 electron, còn mỗi một AO p đơn lẻ chỉ chứa tối đa 2 electron.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 3
    p3 = doc.add_paragraph()
    render_question_prompt(p3, "Câu 3: Phát biểu nào sau đây không đúng khi nói về vai trò của hóa học đối với đời sống và sản xuất?")
    render_options_no_table(doc, [
        "*A. Nghiên cứu quy luật hình thành, vận động của các thiên thể và địa tầng Trái Đất.*",
        "B. Cung cấp phân bón hóa học và thuốc bảo vệ thực vật nhằm tăng năng suất nông nghiệp.",
        "C. Nghiên cứu và sản xuất các loại thuốc chữa bệnh, vắc-xin và vật tư y tế.",
        "D. Tạo ra các sản phẩm tiêu dùng thiết yếu như xà phòng, chất tẩy rửa và chất dẻo."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Nghiên cứu quy luật hình thành, chuyển động của các thiên thể là nhiệm vụ của Thiên văn học; nghiên cứu địa tầng Trái Đất là nhiệm vụ của Địa chất học.\n"
            "- Các phát biểu B, C, D đều thể hiện đúng vai trò thiết yếu của Hóa học đối với đời sống, nông nghiệp và y tế.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 4
    p4 = doc.add_paragraph()
    render_question_prompt(p4, "Câu 4: X là nguyên tố s. Trong bảng tuần hoàn, X có thể thuộc nhóm")
    render_options_no_table(doc, [
        "A. IB.",
        "B. VIA.",
        "*C. IA.*",
        "D. VIB."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Nguyên tố s là các nguyên tố có electron cuối cùng được điền vào phân lớp s. Trong bảng tuần hoàn, các nguyên tố s thuộc nhóm IA hoặc nhóm IIA (và Helium ở nhóm VIIIA). Do đó, X có thể thuộc nhóm IA.\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 5
    p5 = doc.add_paragraph()
    render_question_prompt(p5, "Câu 5: Trong bảng tuần hoàn các nguyên tố hóa học do nhà hóa học người Nga D. I. Mendeleev (1834 – 1907) đề xuất năm 1869, các nguyên tố được sắp xếp theo chiều tăng dần của")
    render_options_no_table(doc, [
        "A. số hiệu nguyên tử.",
        "B. số electron hóa trị.",
        "C. bán kính nguyên tử.",
        "*D. khối lượng nguyên tử.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Năm 1869, Mendeleev sắp xếp các nguyên tố theo chiều tăng dần của khối lượng nguyên tử (nguyên tử khối). Bảng tuần hoàn hiện đại ngày nay mới sắp xếp theo chiều tăng dần của số hiệu nguyên tử (điện tích hạt nhân).\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 6
    p6 = doc.add_paragraph()
    render_question_prompt(p6, "Câu 6: Bảng tuần hoàn các nguyên tố hóa học hiện nay gồm có")
    render_options_no_table(doc, [
        "A. 6 chu kì và 16 nhóm.",
        "B. 8 chu kì và 18 nhóm.",
        "*C. 7 chu kì và 18 cột.*",
        "D. 7 chu kì và 8 nhóm."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Bảng tuần hoàn hiện nay gồm 7 chu kì (chu kì 1 đến chu kì 7) và 18 cột (tương ứng với 8 nhóm A và 8 nhóm B gồm 16 nhóm).\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 7
    p7 = doc.add_paragraph()
    render_question_prompt(p7, "Câu 7: Hạt mang điện trong hạt nhân nguyên tử là")
    render_options_no_table(doc, [
        "A. neutron và electron.",
        "B. electron.",
        "*C. proton.*",
        "D. neutron."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Hạt nhân nguyên tử gồm 2 loại hạt: proton (mang điện tích dương +1) và neutron (không mang điện). Vì vậy, hạt mang điện trong hạt nhân chỉ có proton.\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 8
    p8 = doc.add_paragraph()
    render_question_prompt(p8, "Câu 8: Cho các phát biểu về sự chuyển động của electron trong nguyên tử như sau:\n"
                                "(a) Theo mô hình hành tinh nguyên tử, các electron chuyển động xung quanh hạt nhân theo những quỹ đạo tròn hay bầu dục.\n"
                                "(b) Theo mô hình hiện đại, các electron chuyển động rất nhanh xung quanh hạt nhân theo một quỹ đạo xác định tạo thành đám mây electron.\n"
                                "(c) Theo mô hình hiện đại, trong nguyên tử, electron chuyển động rất nhanh, không theo quỹ đạo xác định.\n"
                                "(d) Theo mô hình hiện đại, xác suất tìm thấy electron tại mỗi điểm trong không gian của AO là 90%.\n"
                                "Số phát biểu đúng là")
    render_options_no_table(doc, [
        "A. 1.",
        "B. 4.",
        "*C. 2.*",
        "D. 3."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- (a) Đúng: Mô hình hành tinh nguyên tử (Rutherford - Bohr) cho rằng electron quay quanh hạt nhân theo các quỹ đạo tròn hoặc elip xác định.\n"
            "- (b) Sai: Theo mô hình hiện đại, electron chuyển động không theo quỹ đạo xác định.\n"
            "- (c) Đúng: Electron chuyển động vô cùng nhanh trong khu vực xung quanh hạt nhân và không theo những quỹ đạo xác định nào.\n"
            "- (d) Sai: Orbital nguyên tử (AO) là khu vực không gian bao quanh hạt nhân mà tại đó xác suất tìm thấy electron là khoảng 90%, chứ không phải tại mỗi điểm xác suất đều là 90%.\n"
            "Vậy có 2 phát biểu đúng là (a) và (c).\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 9
    p9 = doc.add_paragraph()
    render_question_prompt(p9, "Câu 9: Trong nguyên tử, vùng không gian bao quanh hạt nhân và mang điện tích âm được tạo nên bởi các hạt")
    render_options_no_table(doc, [
        "A. proton.",
        "*B. electron.*",
        "C. neutron.",
        "D. alpha."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Lớp vỏ nguyên tử mang điện tích âm bao quanh hạt nhân được cấu tạo từ các hạt electron (mỗi hạt electron mang 1 điện tích âm e^-^).\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 10
    p10 = doc.add_paragraph()
    render_question_prompt(p10, "Câu 10: Kí hiệu phân lớp electron nào sau đây không đúng?")
    render_options_no_table(doc, [
        "A. 3p.",
        "B. 3s.",
        "*C. 3f.*",
        "D. 3d."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Lớp thứ 3 (n = 3) chỉ có tối đa 3 phân lớp là 3s, 3p, 3d. Phân lớp f chỉ bắt đầu xuất hiện từ lớp thứ 4 trở đi (4f, 5f,...). Do đó, kí hiệu phân lớp 3f là không đúng.\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 11
    p11 = doc.add_paragraph()
    render_question_prompt(p11, "Câu 11: Phát biểu nào sau đây không đúng?")
    render_options_no_table(doc, [
        "A. Khối lượng hạt electron nhỏ hơn nhiều so với khối lượng hạt proton.",
        "*B. Khối lượng nguyên tử tập trung chủ yếu ở lớp vỏ.*",
        "C. Trong nguyên tử, khối lượng tập trung chủ yếu ở hạt nhân.",
        "D. Hạt proton và hạt neutron có khối lượng gần bằng nhau."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Khối lượng của electron rất bé (khoảng 0,00055 amu), không đáng kể so với proton (~ 1 amu) và neutron (~ 1 amu). Vì vậy khối lượng của nguyên tử hầu như tập trung toàn bộ ở hạt nhân. Do đó phát biểu B sai.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 12
    p12 = doc.add_paragraph()
    render_question_prompt(p12, "Câu 12: Cho các nguyên tử: ^12^~6~X, ^14^~7~Y, ^14^~6~Z, ^24^~12~T. Các nguyên tử thuộc cùng một nguyên tố hóa học là")
    render_options_no_table(doc, [
        "*A. X và Z.*",
        "B. X và T.",
        "C. Y và Z.",
        "D. Z và T."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Các nguyên tử thuộc cùng một nguyên tố hóa học khi và chỉ khi chúng có cùng số hiệu nguyên tử (cùng số proton Z). Ở đây, ^12^~6~X và ^14^~6~Z đều có Z = 6 (nguyên tố Carbon - C) nên thuộc cùng một nguyên tố hóa học.\n"
            "-> *Đáp án đúng: A*"
        )

    # ==========================================
    # PHẦN II. TRẮC NGHIỆM ĐÚNG / SAI
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai (2,0 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1 (Phần II)
    p_ii_1 = doc.add_paragraph()
    render_question_prompt(p_ii_1, "Câu 1: Fluorine và hợp chất của nó được sử dụng làm chất chống sâu răng, chất cách điện, chất làm lạnh, vật liệu chống dính, ... Nguyên tử fluorine (F) chứa 9 electron và có số khối là 19.")

    stmts_1 = [
        ("a) Hạt nhân nguyên tử F có 9 proton và 10 neutron.", "ĐÚNG"),
        ("b) Nguyên tử F có điện tích lớp vỏ là +9.", "SAI"),
        ("c) Khối lượng của một nguyên tử F xấp xỉ 28 amu.", "SAI"),
        ("d) Phân tử RF~2~ có tổng số hạt mang điện là 60. Nguyên tử của nguyên tố R có 2 electron thuộc lớp M và R là nguyên tố kim loại.", "ĐÚNG")
    ]
    for s_text, ans in stmts_1:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, s_text, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(doc,
            "Lời giải chi tiết:\n"
            "- a) *Đúng*: Nguyên tử F có Z = 9 electron => p = 9 proton; số neutron n = A - p = 19 - 9 = 10.\n"
            "- b) *Sai*: Điện tích của lớp vỏ nguyên tử F là do 9 electron tạo nên, mang điện tích âm là -9 (hoặc -9e = -14,418.10^-19^ C), không phải +9.\n"
            "- c) *Sai*: Khối lượng nguyên tử xấp xỉ bằng số khối A, tức là xấp xỉ 19 amu (không phải 28 amu).\n"
            "- d) *Đúng*: Trong phân tử RF~2~:\n"
            "  + Mỗi nguyên tử F có 2*Z_F = 2*9 = 18 hạt mang điện. Hai nguyên tử F có 36 hạt mang điện.\n"
            "  + Tổng số hạt mang điện: 2*Z_R + 36 = 60 => 2*Z_R = 24 => Z_R = 12 (Magnesium - Mg).\n"
            "  + Cấu hình electron của R (Z = 12): 1s^2^ 2s^2^ 2p^6^ 3s^2^.\n"
            "  + Lớp M (lớp 3) có 2 electron (phân lớp 3s^2^). R có 2 electron lớp ngoài cùng nên là kim loại."
        )

    # Câu 2 (Phần II)
    p_ii_2 = doc.add_paragraph()
    render_question_prompt(p_ii_2, "Câu 2: Boron (B) là nguyên tố có nhiều tác dụng đối với cơ thể người như: làm lành vết thương, điều hoà nội tiết sinh dục, chống viêm khớp, ... Do ngọn lửa cháy có màu đặc biệt nên boron vô định hình được dùng làm pháo hoa. Boron có hai đồng vị là ^10^~5~B và ^11^~5~B, nguyên tử khối trung bình là 10,81.")

    stmts_2 = [
        ("a) Hai đồng vị của boron có cùng số proton nhưng khác nhau về số neutron.", "ĐÚNG"),
        ("b) Nguyên tố B thuộc chu kì 3 trong bảng tuần hoàn.", "SAI"),
        ("c) Nguyên tử của nguyên tố B có 4 electron s và có 3 electron ở phân lớp ngoài cùng.", "SAI"),
        ("d) Phần trăm số nguyên tử của đồng vị ^10^~5~B trong tự nhiên là 81%.", "SAI")
    ]
    for s_text, ans in stmts_2:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, s_text, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(doc,
            "Lời giải chi tiết:\n"
            "- a) *Đúng*: Các đồng vị của cùng một nguyên tố có cùng số proton (Z = 5) nhưng khác nhau về số neutron (^10^~5~B có 5 neutron; ^11^~5~B có 6 neutron).\n"
            "- b) *Sai*: Cấu hình electron của B (Z = 5) là 1s^2^ 2s^2^ 2p^1^. Nguyên tử có 2 lớp electron nên B thuộc chu kì 2 (không phải chu kì 3).\n"
            "- c) *Sai*: Nguyên tử B có 4 electron s (2e ở 1s và 2e ở 2s). Nhưng phân lớp ngoài cùng là 2p chỉ có 1 electron (chứ không phải 3 electron ở phân lớp ngoài cùng; lớp ngoài cùng mới có 3 electron).\n"
            "- d) *Sai*: Gọi x là phần trăm số nguyên tử của đồng vị ^10^~5~B, ta có:\n"
            "  [10*x + 11*(100 - x)] / 100 = 10,81 => 1100 - x = 1081 => x = 19%.\n"
            "  Vậy đồng vị ^10^~5~B chiếm 19%, đồng vị ^11^~5~B chiếm 81%."
        )

    # ==========================================
    # PHẦN III. CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn (2,0 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4."
    )

    # Câu 1 (Phần III)
    p_iii_1 = doc.add_paragraph()
    render_question_prompt(p_iii_1, "Câu 1: Cho biết tổng số hạt proton, neutron và electron trong một nguyên tử của nguyên tố ^39^~19~K là bao nhiêu?")
    if is_teacher:
        p_ans1 = doc.add_paragraph()
        p_ans1.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans1, "*Trả lời:* *58*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Trong nguyên tử ^39^~19~K:\n"
            "- Số proton p = 19, số electron e = p = 19.\n"
            "- Số neutron n = A - p = 39 - 19 = 20.\n"
            "- Tổng số hạt: S = p + e + n = 19 + 19 + 20 = 58.\n"
            "-> *Đáp án: 58*"
        )
    else:
        p_ans1 = doc.add_paragraph()
        p_ans1.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans1, "*Trả lời:* ............................................................", base_size=10.5)

    # Câu 2 (Phần III)
    p_iii_2 = doc.add_paragraph()
    render_question_prompt(p_iii_2, "Câu 2: Nguyên tử của nguyên tố calcium (Z = 20) có bao nhiêu lớp electron?")
    if is_teacher:
        p_ans2 = doc.add_paragraph()
        p_ans2.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans2, "*Trả lời:* *4*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Cấu hình electron của Ca (Z = 20): 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^6^ 4s^2^.\n"
            "Nguyên tử có 4 lớp electron (lớp 1: K, lớp 2: L, lớp 3: M, lớp 4: N).\n"
            "-> *Đáp án: 4*"
        )
    else:
        p_ans2 = doc.add_paragraph()
        p_ans2.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans2, "*Trả lời:* ............................................................", base_size=10.5)

    # Câu 3 (Phần III)
    p_iii_3 = doc.add_paragraph()
    render_question_prompt(p_iii_3, "Câu 3: Cho các nguyên tử sau: X (Z = 11); Y (Z = 17); Z (Z = 20); T (Z = 26). Có bao nhiêu nguyên tố thuộc loại nguyên tố s?")
    if is_teacher:
        p_ans3 = doc.add_paragraph()
        p_ans3.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans3, "*Trả lời:* *2*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Viết cấu hình electron của các nguyên tử:\n"
            "- X (Z = 11): 1s^2^ 2s^2^ 2p^6^ 3s^1^ -> electron cuối cùng điền vào phân lớp s -> nguyên tố s.\n"
            "- Y (Z = 17): 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^5^ -> electron cuối cùng điền vào phân lớp p -> nguyên tố p.\n"
            "- Z (Z = 20): 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^6^ 4s^2^ -> electron cuối cùng điền vào phân lớp s -> nguyên tố s.\n"
            "- T (Z = 26): 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^6^ 3d^6^ 4s^2^ -> electron cuối cùng điền vào phân lớp d -> nguyên tố d.\n"
            "Vậy có 2 nguyên tố thuộc loại nguyên tố s là X và Z.\n"
            "-> *Đáp án: 2*"
        )
    else:
        p_ans3 = doc.add_paragraph()
        p_ans3.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans3, "*Trả lời:* ............................................................", base_size=10.5)

    # Câu 4 (Phần III)
    p_iii_4 = doc.add_paragraph()
    render_question_prompt(p_iii_4, "Câu 4: Trong tự nhiên, nguyên tố bromine có 2 đồng vị là ^79^~35~Br và ^81^~35~Br. Nguyên tử khối trung bình của bromine là 79,91. Trong hợp chất NaBrO~x~, nguyên tử đồng vị ^81^Br chiếm 30,994% về khối lượng. Giá trị của x là bao nhiêu (với O và Na là các đồng vị ^16^~8~O, ^23^~11~Na)?")
    if is_teacher:
        p_ans4 = doc.add_paragraph()
        p_ans4.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans4, "*Trả lời:* *1*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "- Gọi y là phần trăm số nguyên tử của đồng vị ^79^Br trong tự nhiên:\n"
            "  79*y + 81*(1 - y) = 79,91 => 81 - 2y = 79,91 => 2y = 1,09 => y = 0,545 (54,5%).\n"
            "  => Tỉ lệ phần trăm số nguyên tử của đồng vị ^81^Br là 100% - 54,5% = 45,5% (ứng với phân số 0,455).\n"
            "- Xét 1 mol hợp chất NaBrO~x~:\n"
            "  + Khối lượng mol của phân tử: M = 23 + 79,91 + 16x = 102,91 + 16x (g/mol).\n"
            "  + Số mol đồng vị ^81^Br có trong 1 mol phân tử là 0,455 mol.\n"
            "  + Khối lượng đồng vị ^81^Br: m(^81^Br) = 0,455 * 81 = 36,855 gam.\n"
            "- Theo đề bài:\n"
            "  %m(^81^Br) = [36,855 / (102,91 + 16x)] * 100% = 30,994%\n"
            "  => 36,855 / (102,91 + 16x) = 0,30994\n"
            "  => 102,91 + 16x = 36,855 / 0,30994 ≈ 118,91\n"
            "  => 16x = 118,91 - 102,91 = 16 => x = 1.\n"
            "-> *Đáp án: 1*"
        )
    else:
        p_ans4 = doc.add_paragraph()
        p_ans4.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans4, "*Trả lời:* ............................................................", base_size=10.5)

    # ==========================================
    # B. PHẦN TỰ LUẬN (3 ĐIỂM)
    # ==========================================
    p_b = doc.add_paragraph()
    p_b.paragraph_format.space_before = Pt(8)
    p_b.paragraph_format.space_after = Pt(2)
    render_rich_text(p_b, "*B. PHẦN TỰ LUẬN (3 ĐIỂM)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # Câu 1 Tự luận
    p_tl1 = doc.add_paragraph()
    render_question_prompt(p_tl1, "Câu 1: (1,0 điểm) Phosphorus (P) là một nguyên tố dinh dưỡng khoáng thiết yếu cho sự phát triển của cây trồng, thường có mặt trong các loại phân bón NPK. Trong bảng tuần hoàn, nguyên tố phosphorus có số hiệu nguyên tử là 15.\n"
                                  "a) Viết cấu hình electron nguyên tử của P ở trạng thái cơ bản.\n"
                                  "b) Xác định vị trí của P (chu kì và nhóm) trong bảng tuần hoàn, cho biết P thuộc loại nguyên tố nào (s, p hay d)?")
    if is_teacher:
        add_solution_no_table(doc,
            "Lời giải chi tiết:\n"
            "a) Cấu hình electron nguyên tử của P (Z = 15):\n"
            "   1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^3^ (hoặc [Ne] 3s^2^ 3p^3^). (0,5 điểm)\n"
            "b) Vị trí của P trong bảng tuần hoàn:\n"
            "   - Ô nguyên tố: số 15 (vì Z = 15).\n"
            "   - Chu kì: Chu kì 3 (vì nguyên tử có 3 lớp electron).\n"
            "   - Nhóm: Nhóm VA (vì có 5 electron hóa trị ở lớp ngoài cùng 3s^2^ 3p^3^, electron cuối cùng điền vào phân lớp p).\n"
            "   - Loại nguyên tố: P là nguyên tố p (vì electron cuối cùng được điền vào phân lớp p). (0,5 điểm)"
        )
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Câu 2 Tự luận
    p_tl2 = doc.add_paragraph()
    render_question_prompt(p_tl2, "Câu 2: (1,0 điểm) Neon (Ne) là một nguyên tố khí hiếm, phát ra ánh sáng đỏ cam đặc trưng khi có phóng điện qua khí ở áp suất thấp và được ứng dụng trong một số loại biển hiệu phát sáng. Neon có ba đồng vị bền và tỉ lệ phần trăm số nguyên tử các đồng vị của Ne được xác định theo phổ khối lượng như hình sau:")

    # Chèn ảnh phổ khối lượng Neon
    if os.path.exists(IMG_NEON_PATH):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(3)
        p_img.paragraph_format.space_after = Pt(3)
        run_img = p_img.add_run()
        run_img.add_picture(IMG_NEON_PATH, width=Inches(3.8))

    p_tl2_sub = doc.add_paragraph()
    p_tl2_sub.paragraph_format.space_before = Pt(2)
    p_tl2_sub.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl2_sub,
        "a) Tính nguyên tử khối trung bình của Neon (kết quả làm tròn đến hàng phần trăm).\n"
        "b) Tính số nguyên tử của mỗi loại đồng vị còn lại khi có 1 nguyên tử ^21^~10~Ne (kết quả làm tròn đến hàng đơn vị).\n"
        "c) Tính thể tích của 10 gam Ne ở điều kiện chuẩn (25°C, 1 bar) (kết quả làm tròn đến hàng phần mười).",
        base_size=10.5
    )

    if is_teacher:
        add_solution_no_table(doc,
            "Lời giải chi tiết:\n"
            "Từ phổ khối lượng ta có thành phần phần trăm các đồng vị Neon:\n"
            "- ^20^~10~Ne chiếm 90,0%.\n"
            "- ^21^~10~Ne chiếm 1,0%.\n"
            "- ^22^~10~Ne chiếm 9,0%.\n"
            "a) Nguyên tử khối trung bình của Neon:\n"
            "   A_Ne = (20 * 90,0 + 21 * 1,0 + 22 * 9,0) / 100 = 2019 / 100 = 20,19. (0,4 điểm)\n"
            "b) Tỉ lệ số nguyên tử của 3 đồng vị là:\n"
            "   N(^20^Ne) : N(^21^Ne) : N(^22^Ne) = 90,0 : 1,0 : 9,0 = 90 : 1 : 9.\n"
            "   Khi có 1 nguyên tử ^21^~10~Ne thì tương ứng sẽ có:\n"
            "   - 90 nguyên tử ^20^~10~Ne.\n"
            "   - 9 nguyên tử ^22^~10~Ne. (0,3 điểm)\n"
            "c) Số mol của 10 gam Ne:\n"
            "   n_Ne = m / A_Ne = 10 / 20,19 ≈ 0,4953 mol.\n"
            "   Ở điều kiện chuẩn (25°C, 1 bar), 1 mol khí chiếm 24,79 lít:\n"
            "   V_Ne = 0,4953 * 24,79 ≈ 12,278 ≈ 12,3 (lít). (0,3 điểm)"
        )
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 3 Tự luận
    p_tl3 = doc.add_paragraph()
    render_question_prompt(p_tl3, "Câu 3: (1,0 điểm) Nguyên tố R thuộc nhóm A trong bảng tuần hoàn. Nguyên tử R có tổng số electron ở các phân lớp s là 7 và tổng số electron ở lớp ngoài cùng là 1.\n"
                                  "a) Viết cấu hình electron đầy đủ của nguyên tử R. Dự đoán tính chất hóa học cơ bản của nguyên tố R (kim loại hay phi kim) và giải thích.\n"
                                  "b) Phân tử R~2~XY~4~ có tổng số hạt proton là 86. Biết nguyên tố X và Y thuộc cùng một nhóm A và ở hai chu kì liên tiếp trong bảng tuần hoàn. Hãy xác định số hiệu nguyên tử của X và Y.")
    if is_teacher:
        add_solution_no_table(doc,
            "Lời giải chi tiết:\n"
            "a) Xác định nguyên tố R:\n"
            "- R thuộc nhóm A, có tổng số electron ở các phân lớp s là 7: 1s^2^ 2s^2^ 3s^2^ 4s^1^ (2+2+2+1 = 7 electron).\n"
            "- Lớp ngoài cùng (lớp 4) có 1 electron (4s^1^).\n"
            "=> Cấu hình electron đầy đủ của R là: 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^6^ 4s^1^ (hay [Ar] 4s^1^). Nguyên tố R có Z = 19 (Potassium - K).\n"
            "- Dự đoán tính chất: R là một kim loại mạnh (kim loại kiềm điển hình).\n"
            "  Giải thích: Do nguyên tử R có 1 electron ở lớp ngoài cùng, bán kính nguyên tử lớn, năng lượng ion hóa nhỏ nên rất dễ nhường đi 1 electron này để đạt cấu hình electron bền vững của khí hiếm Ar (tạo thành ion R^+^): R -> R^+^ + 1e. (0,5 điểm)\n"
            "b) Xác định số hiệu nguyên tử của X và Y:\n"
            "- Trong phân tử R~2~XY~4~: tổng số proton là 86\n"
            "  2 * Z_R + Z_X + 4 * Z_Y = 86 => 2 * 19 + Z_X + 4 * Z_Y = 86\n"
            "  => Z_X + 4 * Z_Y = 48. (*)\n"
            "- Do X và Y thuộc cùng một nhóm A và ở hai chu kì liên tiếp trong bảng tuần hoàn:\n"
            "  Vì Z_X + 4*Z_Y = 48 nên Z_Y < 12 => Y thuộc chu kì nhỏ (chu kì 2 hoặc 3). Hai nguyên tố cùng nhóm A ở hai chu kì liên tiếp thuộc chu kì nhỏ có số proton hơn kém nhau 8 đơn vị: |Z_X - Z_Y| = 8.\n"
            "  + Trường hợp 1: Z_Y = Z_X + 8 => Z_X + 4*(Z_X + 8) = 48 => 5*Z_X = 16 => Z_X = 3,2 (loại).\n"
            "  + Trường hợp 2: Z_X = Z_Y + 8 => (Z_Y + 8) + 4*Z_Y = 48 => 5*Z_Y = 40 => Z_Y = 8.\n"
            "    => Z_X = 8 + 8 = 16.\n"
            "- Kiểm tra cấu hình và vị trí:\n"
            "  + Y (Z = 8): 1s^2^ 2s^2^ 2p^4^ -> chu kì 2, nhóm VIA (Oxygen - O).\n"
            "  + X (Z = 16): 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^4^ -> chu kì 3, nhóm VIA (Sulfur - S).\n"
            "  X và Y cùng thuộc nhóm VIA, ở hai chu kì liên tiếp 2 và 3 (thỏa mãn hoàn toàn). Phân tử tương ứng là K~2~SO~4~.\n"
            "Vậy: Số hiệu nguyên tử của X là 16 (Sulfur) và số hiệu nguyên tử của Y là 8 (Oxygen). (0,5 điểm)"
        )
    else:
        add_student_writing_lines(doc, num_lines=10)

    # Chân đề thi
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(10)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "---------- *HẾT* ----------", base_size=10.5, base_bold=True)

    p_footer_sub = doc.add_paragraph()
    p_footer_sub.paragraph_format.space_before = Pt(0)
    p_footer_sub.paragraph_format.space_after = Pt(2)
    p_footer_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_footer_sub, "_Thí sinh không được sử dụng tài liệu. Cán bộ coi thi không giải thích gì thêm._", base_size=9.5, base_italic=True)

    return doc

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 1. Bản Học sinh
    doc_hs = build_exam_doc(is_teacher=False)
    hs_path = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_Ma1201_HocSinh.docx")
    doc_hs.save(hs_path)
    print(f"Đã lưu Bản Học sinh: {hs_path}")

    # 2. Bản Giáo viên
    doc_gv = build_exam_doc(is_teacher=True)
    gv_path = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_10_Ma1201_GiaoVien.docx")
    doc_gv.save(gv_path)
    print(f"Đã lưu Bản Giáo viên: {gv_path}")

if __name__ == "__main__":
    main()
