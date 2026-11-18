# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì I Hóa Học 11 - THPT Nguyễn Chí Thanh - Mã đề 1002
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản Word:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_11_NguyenChiThanh_Ma1002_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_11_NguyenChiThanh_Ma1002_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
    p1_ans = ["A", "A", "D", "C", "C", "D", "B", "B", "B", "A", "D", "D", "D", "C", "D", "D"]
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
        ("Câu 1", ["ĐÚNG", "ĐÚNG", "SAI", "ĐÚNG"]),
        ("Câu 2", ["ĐÚNG", "SAI", "ĐÚNG", "SAI"])
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
    p3_answers = ["1,7", "12", "4", "3"]
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
        school_line="SỞ GIÁO DỤC VÀ ĐÀO TẠO TP. HUẾ\nTRƯỜNG THPT NGUYỄN CHÍ THANH",
        exam_title="ĐỀ KIỂM TRA GIỮA HỌC KỲ I – NĂM HỌC 2026 – 2027",
        subject_line="MÔN: HOÁ HỌC. LỚP 11",
        num_pages="04",
        duration="45 phút (không kể thời gian phát đề)",
        exam_code="1002",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; S = 32; Cl = 35,5; Zn = 65; Ba = 137",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table(doc)

    # ==========================================
    # PHẦN I. TRẮC NGHIỆM NHIỀU LỰA CHỌN (4,0 ĐIỂM)
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn: (16 câu - 4 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 16. Mỗi câu thí sinh chỉ chọn một phương án."
    )

    # Câu 1
    p1 = doc.add_paragraph()
    render_question_prompt(p1, "Câu 1. Hiện tượng nào sau đây không phải là tác hại chính của mưa acid?")
    render_options_no_table(doc, [
        "*A. Làm gia tăng hiện tượng nóng lên toàn cầu*",
        "B. Làm hủy hoại rừng và thảm thực vật cây cối",
        "C. Gây ăn mòn và hủy hoại các công trình kiến trúc",
        "D. Gây hại và làm suy giảm sinh vật dưới nước"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Hiện tượng nóng lên toàn cầu là do sự gia tăng nồng độ của các khí nhà kính (chủ yếu là CO~2~, CH~4~, N~2~O,...) giữ nhiệt trong khí quyển, không phải là tác hại chính của hiện tượng mưa acid. Các tác hại chính của mưa acid gồm: làm chua đất, hủy hoại rừng và mùa màng, làm chết sinh vật thủy sinh, và ăn mòn công trình, di tích lịch sử.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 2
    p2 = doc.add_paragraph()
    render_question_prompt(p2, "Câu 2. pH của dung dịch nào sau đây có giá trị nhỏ nhất?")
    render_options_no_table(doc, [
        "*A. Dung dịch HCl 0,1M.*",
        "B. Dung dịch CH~3~COOH 0,1M.",
        "C. Dung dịch NaCl 0,1M.",
        "D. Dung dịch NaOH 0,01M."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Xét độ pH của các dung dịch:\n"
            "- Dung dịch HCl 0,1 M: là acid mạnh phân li hoàn toàn: HCl ➔ H^+^ + Cl^-^ ➔ [H^+^] = 0,1 M = 10^-1^ M ➔ pH = 1.\n"
            "- Dung dịch CH~3~COOH 0,1 M: là acid yếu phân li thuận nghịch không hoàn toàn: CH~3~COOH ⇌ CH~3~COO^-^ + H^+^ ➔ [H^+^] < 0,1 M ➔ pH > 1 (thực tế pH ≈ 2,88).\n"
            "- Dung dịch NaCl 0,1 M: muối tạo bởi acid mạnh và base mạnh, trung tính ➔ pH = 7.\n"
            "- Dung dịch NaOH 0,01 M: base mạnh phân li hoàn toàn ➔ [OH^-^] = 0,01 M = 10^-2^ M ➔ pOH = 2 ➔ pH = 14 - 2 = 12.\n"
            "Vậy dung dịch có pH nhỏ nhất là dung dịch HCl 0,1 M (pH = 1).\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 3
    p3 = doc.add_paragraph()
    render_question_prompt(p3, "Câu 3. Chuẩn độ dung dịch NaOH chưa biết chính xác nồng độ (biết nồng độ trong khoảng gần với 0,1 M) bằng dung dịch chuẩn HCl 0,1 M với chỉ thị phenolphthalein. Phương trình thể hiện bản chất của phản ứng của quá trình chuẩn độ là ?")
    render_options_no_table(doc, [
        "A. NaCl ➔ Na^+^ + Cl^-^",
        "B. Na^+^ + Cl^-^ ➔ NaCl",
        "C. H~2~O ⇌ H^+^ + OH^-^",
        "*D. H^+^ + OH^-^ ➔ H~2~O*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Bản chất của phản ứng chuẩn độ acid mạnh (HCl) bằng base mạnh (NaOH) là phản ứng trung hòa giữa ion H^+^ của acid và ion OH^-^ của base để tạo thành phân tử nước: H^+^ + OH^-^ ➔ H~2~O.\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 4
    p4 = doc.add_paragraph()
    render_question_prompt(p4, "Câu 4. Cho m gam Al phản ứng hoàn toàn với dung dịch HNO~3~ loãng (dư), thu được 4,958 lít khí NO (đkc, sản phẩm khử duy nhất). Giá trị của m là")
    render_options_no_table(doc, [
        "A. 4,5.",
        "B. 8,1.",
        "*C. 5,4.*",
        "D. 2,7."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Số mol khí NO thu được (ở đkc): n~NO~ = 4,958 / 24,79 = 0,2 mol.\n"
            "- Quá trình nhường - nhận electron:\n"
            "  Al ➔ Al^3+^ + 3e\n"
            "  N^+5^ + 3e ➔ N^+2^ (NO)\n"
            "- Áp dụng định luật bảo toàn electron:\n"
            "  3 × n~Al~ = 3 × n~NO~ ➔ n~Al~ = n~NO~ = 0,2 mol.\n"
            "- Giá trị của m là khối lượng Al phản ứng:\n"
            "  m = 0,2 × 27 = 5,4 gam.\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 5
    p5 = doc.add_paragraph()
    render_question_prompt(p5, "Câu 5. Nitrogen (N~2~) thể hiện tính khử trong phản ứng với chất nào trong các chất sau?")
    render_options_no_table(doc, [
        "A. H~2~.",
        "B. Li.",
        "*C. O~2~.*",
        "D. Mg."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Trong phản ứng với O~2~ (ở khoảng 3000 °C hoặc tia lửa điện):\n"
            "  N~2~ + O~2~ ⇌ 2NO\n"
            "  Số oxi hóa của N tăng từ 0 lên +2 ➔ N~2~ đóng vai trò là chất khử.\n"
            "- Trong các phản ứng với kim loại mạnh (Li, Mg) và H~2~, số oxi hóa của N giảm từ 0 xuống -3 ➔ N~2~ đóng vai trò là chất oxi hóa.\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 6
    p6 = doc.add_paragraph()
    render_question_prompt(p6, "Câu 6. Tính chất hóa học cơ bản của ammonia (NH~3~) là")
    render_options_no_table(doc, [
        "A. tính base mạnh, tính khử.",
        "B. tính base yếu, tính oxi hóa.",
        "C. tính base mạnh, tính oxi hóa.",
        "*D. tính khử mạnh, tính base yếu.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Trên nguyên tử N trong phân tử NH~3~ còn một cặp electron chưa liên kết, có khả năng nhận proton H^+^, do đó NH~3~ có tính base yếu.\n"
            "- Nguyên tử N trong NH~3~ có số oxi hóa là -3 (trạng thái oxi hóa thấp nhất của nitrogen), nên trong các phản ứng oxi hóa - khử, NH~3~ dễ nhường electron thể hiện tính khử mạnh.\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 7
    p7 = doc.add_paragraph()
    render_question_prompt(p7, "Câu 7. Các yếu tố có thể ảnh hưởng đến sự chuyển dịch cân bằng hoá học là :")
    render_options_no_table(doc, [
        "A. nồng độ, áp suất và diện tích bề mặt.",
        "*B. nồng độ, nhiệt độ và áp suất.*",
        "C. áp suất, nhiệt độ và chất xúc tác.",
        "D. nồng độ, nhiệt độ và chất xúc tác."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Theo nguyên lí chuyển dịch cân bằng Le Chatelier, các yếu tố ảnh hưởng đến sự chuyển dịch cân bằng hóa học là: nồng độ, nhiệt độ và áp suất (đối với hệ có chất khí). Chất xúc tác và diện tích bề mặt tiếp xúc chỉ làm thay đổi tốc độ phản ứng (giúp cân bằng nhanh được thiết lập hơn) chứ không làm chuyển dịch cân bằng.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 8
    p8 = doc.add_paragraph()
    render_question_prompt(p8, "Câu 8. Số oxi hoá của nitrogen (N) trong phân tử nitric acid (HNO~3~) là")
    render_options_no_table(doc, [
        "A. +4.",
        "*B. +5.*",
        "C. +2.",
        "D. +3."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Trong phân tử HNO~3~, nguyên tử H có số oxi hóa là +1, nguyên tử O có số oxi hóa là -2. Gọi x là số oxi hóa của nguyên tử N, ta có:\n"
            "(+1) × 1 + x × 1 + (-2) × 3 = 0 ➔ x - 5 = 0 ➔ x = +5.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 9
    p9 = doc.add_paragraph()
    render_question_prompt(p9, "Câu 9. Hành động nào sau đây của con người không phải là nguyên nhân chính gây ra hiện tượng phú dưỡng?")
    render_options_no_table(doc, [
        "A. Xả thải nước thải chăn nuôi, công nghiệp chứa hàm lượng cao các nguyên tố dinh dưỡng (N, P).",
        "*B. Đánh bắt thuỷ hải sản bằng chích điện và chất nổ.*",
        "C. Sử dụng dư thừa lượng phân bón chứa ion PO~4~^3-^ ngấm vào nguồn nước mặt.",
        "D. Sử dụng dư thừa lượng phân bón chứa các ion NO~2~^-^, NO~3~^-^ trong sản xuất nông nghiệp."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Hiện tượng phú dưỡng (eutrophication) là hiện tượng ao, hồ tiếp nhận một lượng quá mức các chất dinh dưỡng (chủ yếu là các hợp chất của nitrogen và phosphorus), làm cho tảo, rong rêu bùng phát mạnh mẽ.\n"
            "- Các hành động A, C, D trực tiếp đưa lượng lớn N và P vào nguồn nước, là các nguyên nhân chính gây phú dưỡng.\n"
            "- Hành vi đánh bắt thủy sản bằng xung điện và chất nổ là hành vi phá hủy môi sinh và tận diệt nguồn lợi thủy sản, không phải nguyên nhân chính tích tụ dinh dưỡng gây hiện tượng phú dưỡng.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 10
    p10 = doc.add_paragraph()
    render_question_prompt(p10, "Câu 10. Phát biểu nào sau đây về acid – base theo thuyết Bronsted – Lowry là đúng?")
    render_options_no_table(doc, [
        "*A. Acid là chất cho proton (H^+^).*",
        "B. Base là chất nhận electron.",
        "C. Base là chất cho proton (H^+^).",
        "D. Acid là chất nhận proton (H^+^)."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Theo thuyết acid - base của Brønsted – Lowry:\n"
            "- Acid là chất cho proton (H^+^).\n"
            "- Base là chất nhận proton (H^+^).\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 11
    p11 = doc.add_paragraph()
    render_question_prompt(p11, "Câu 11. Cho dung dịch nitric acid (HNO~3~) tác dụng với các chất sau: Cu, NaOH, C, FeO, Na~2~CO~3~. Khi phản ứng xảy ra, có bao nhiêu phản ứng trong đó nitric acid (HNO~3~) đóng vai trò oxi hóa ?")
    render_options_no_table(doc, [
        "A. 4.",
        "B. 5.",
        "C. 2.",
        "*D. 3.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "HNO~3~ đóng vai trò là chất oxi hóa khi số oxi hóa của nitrogen giảm từ +5 xuống các mức thấp hơn (như +4, +2, +1, 0, -3). Xét từng chất:\n"
            "1. Cu: Cu + 4HNO~3~ ➔ Cu(NO~3~)~2~ + 2NO~2~ + 2H~2~O (N^+5^ nhận e ➔ HNO~3~ là chất oxi hóa). -> (1)\n"
            "2. NaOH: NaOH + HNO~3~ ➔ NaNO~3~ + H~2~O (phản ứng trung hòa trao đổi, không đổi số oxi hóa ➔ HNO~3~ chỉ đóng vai trò là acid).\n"
            "3. C: C + 4HNO~3~ ➔ CO~2~ + 4NO~2~ + 2H~2~O (N^+5^ oxi hóa C^0^ lên C^+4^ ➔ HNO~3~ là chất oxi hóa). -> (2)\n"
            "4. FeO: 3FeO + 10HNO~3~ ➔ 3Fe(NO~3~)~3~ + NO + 5H~2~O (Fe^+2^ bị oxi hóa lên Fe^+3^, N^+5^ nhận e ➔ HNO~3~ là chất oxi hóa). -> (3)\n"
            "5. Na~2~CO~3~: Na~2~CO~3~ + 2HNO~3~ ➔ 2NaNO~3~ + CO~2~ + H~2~O (phản ứng trao đổi ion, không đổi số oxi hóa ➔ HNO~3~ chỉ đóng vai trò là acid).\n"
            "Như vậy có 3 phản ứng HNO~3~ đóng vai trò là chất oxi hóa (với Cu, C, FeO).\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 12
    p12 = doc.add_paragraph()
    render_question_prompt(p12, "Câu 12. Các dung dịch của muối, acid, base dẫn điện là do")
    render_options_no_table(doc, [
        "A. phân tử của chúng dẫn được điện.",
        "B. có sự di chuyển của electron tạo thành dòng electron.",
        "C. các ion hợp phần có khả năng dẫn điện.",
        "*D. muối, acid, base có khả năng phân li ra ion khi tan trong nước.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Sự dẫn điện của các dung dịch acid, base, muối là do khi hòa tan trong nước, các phân tử acid, base và muối phân li tạo thành các ion dương (cation) và ion âm (anion) tự do chuyển động trong dung dịch (chất điện li).\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 13
    p13 = doc.add_paragraph()
    render_question_prompt(p13, "Câu 13. Hệ phản ứng sau ở trạng thái cân bằng : H~2~(g) + I~2~(g) ⇌ 2HI(g)\n"
                                 "Biểu thức hằng số cân bằng của phản ứng trên là")
    render_options_no_table(doc, [
        "A. K~c~ = [2HI] / ([H~2~].[I~2~])",
        "B. K~c~ = ([H~2~].[I~2~]) / (2[HI])",
        "C. K~c~ = ([H~2~].[I~2~]) / [HI]^2^",
        "*D. K~c~ = [HI]^2^ / ([H~2~].[I~2~])*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Theo định luật tác dụng khối lượng, biểu thức hằng số cân bằng nồng độ K~c~ của phản ứng aA + bB ⇌ cC + dD là K~c~ = ([C]^c^.[D]^d^) / ([A]^a^.[B]^b^).\n"
            "Đối với phản ứng H~2~(g) + I~2~(g) ⇌ 2HI(g), ta có: K~c~ = [HI]^2^ / ([H~2~].[I~2~]).\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 14
    p14 = doc.add_paragraph()
    render_question_prompt(p14, "Câu 14. Ứng dụng phổ biến và quan trọng nhất của muối ammonium trong đời sống và sản xuất là")
    render_options_no_table(doc, [
        "A. sản xuất nhiên liệu xăng sinh học và dầu diesel.",
        "B. sản xuất xi măng và vật liệu chịu lửa trong xây dựng.",
        "*C. sản xuất phân bón hoá học, cung cấp nguyên tố nitrogen dinh dưỡng cho cây trồng.*",
        "D. điều chế kim loại nguyên chất bằng phương pháp nhiệt luyện."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Muối ammonium (như NH~4~NO~3~, NH~4~Cl, (NH~4~)~2~SO~4~, (NH~4~)~2~HPO~4~,...) được ứng dụng chủ yếu và quan trọng nhất để làm phân bón hóa học (phân đạm, phân phức hợp), cung cấp dưỡng chất nitrogen cần thiết cho cây trồng sinh trưởng và phát triển.\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 15
    p15 = doc.add_paragraph()
    render_question_prompt(p15, "Câu 15. Bóng cười (hay \"funky ball\") là một quả bóng được bơm đầy khí X. Khí X có vị ngọt nhẹ, không màu, gây tác động hưng phấn, ảo giác cười cho người hít phải. Lạm dụng bóng cười gây tác hại nghiêm trọng đến hệ thần kinh, tim mạch, tụt huyết áp, co giật và nguy cơ tử vong. Công thức hoá học của khí X là")
    render_options_no_table(doc, [
        "A. O~2~.",
        "B. N~2~.",
        "C. NO.",
        "*D. N~2~O.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Khí gây cười (laughing gas) là dinitrogen monoxide có công thức hóa học là N~2~O. Khí này có vị ngọt nhẹ, không màu, có tác dụng kích thích thần kinh tạo ảo giác hưng phấn cười đùa. Lạm dụng N~2~O gây tổn thương tủy sống và hệ thần kinh ngoại biên nghiêm trọng.\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 16
    p16 = doc.add_paragraph()
    render_question_prompt(p16, "Câu 16. Phản ứng nào sau đây là phản ứng thuận nghịch ?")
    render_options_no_table(doc, [
        "A. Fe + 2HCl ➔ FeCl~2~ + H~2~.",
        "B. H~2~ + Cl~2~ ➔ 2HCl.",
        "C. 2H~2~ + O~2~ ➔ 2H~2~O.",
        "*D. N~2~ + 3H~2~ ⇌ 2NH~3~.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Phản ứng tổng hợp ammonia từ nitrogen và hydrogen trong điều kiện nhiệt độ cao, áp suất cao và có xúc tác sắt là một phản ứng thuận nghịch (biểu thị bằng mũi tên hai chiều ⇌):\n"
            "N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g).\n"
            "Các phản ứng ở phương án A, B, C đều là các phản ứng một chiều trong điều kiện thông thường.\n"
            "-> *Đáp án đúng: D*"
        )

    # ==========================================
    # PHẦN II. TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM)
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai: (2 câu - 2 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu thí sinh chọn đúng hoặc sai."
    )

    # Câu 1 Phần II
    p2_1 = doc.add_paragraph()
    render_question_prompt(p2_1, "Câu 1. Trong dung dịch muối Fe^3+^ tồn tại cân bằng hóa học sau:\n"
                                  "Fe^3+^ + 3H~2~O ⇌ Fe(OH)~3~↓ + 3H^+^")
    p2_1_stmts = [
        ("a) Trong phòng thí nghiệm, để bảo quản dung dịch Fe^3+^, người ta thường thêm vào bình đựng vài giọt dung dịch acid HCl để cân bằng chuyển dịch theo chiều nghịch.", "Đúng"),
        ("b) Dựa vào thuyết acid – base của Brønsted – Lowry, Fe^3+^ đóng vai trò acid.", "Đúng"),
        ("c) Thêm vào dung dịch muối Fe^3+^ vài giọt dung dịch NaOH thì cân bằng chuyển dịch theo chiều nghịch.", "Sai"),
        ("d) Các loại đất có chứa nhiều ion Fe^3+^ có giá trị pH thấp hay còn gọi là đất chua.", "Đúng")
    ]
    for stmt_text, ans in p2_1_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt_text, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Ý a) ĐÚNG: Khi thêm vài giọt acid HCl, nồng độ H^+^ tăng lên. Theo nguyên lí Le Chatelier, cân bằng sẽ chuyển dịch theo chiều làm giảm nồng độ H^+^, tức là chiều nghịch, ngăn cản sự thủy phân tạo kết tủa Fe(OH)~3~ và giữ cho dung dịch trong suốt, không bị vẩn đục.\n"
            "- Ý b) ĐÚNG: Trong dung dịch nước, cation Fe^3+^ liên kết với các phân tử nước tạo phức chất [Fe(H~2~O)~6~]^3+^ và cho proton H^+^ vào nước: [Fe(H~2~O)~6~]^3+^ + H~2~O ⇌ [Fe(H~2~O)~5~(OH)]^2+^ + H~3~O^+^. Vì có khả năng cho proton H^+^ nên theo Brønsted – Lowry, Fe^3+^ đóng vai trò là một acid.\n"
            "- Ý c) SAI: Khi thêm dung dịch NaOH, ion OH^-^ sẽ phản ứng với H^+^ (H^+^ + OH^-^ ➔ H~2~O) làm giảm nồng độ H^+^ trong dung dịch. Cân bằng sẽ chuyển dịch theo chiều tạo thêm H^+^, tức là CHIỀU THUẬN (tạo thêm kết tủa Fe(OH)~3~), chứ không phải chiều nghịch.\n"
            "- Ý d) ĐÚNG: Quá trình thủy phân của ion Fe^3+^ sản sinh ra nhiều ion H^+^ làm tăng nồng độ H^+^ và hạ thấp độ pH của môi trường đất (pH < 7), do đó các loại đất có chứa nhiều khoáng chất sắt Fe^3+^ thường có tính acid và được gọi là đất chua.\n"
            "-> *Đáp án: a) ĐÚNG - b) ĐÚNG - c) SAI - d) ĐÚNG*"
        )

    # Câu 2 Phần II
    p2_2 = doc.add_paragraph()
    render_question_prompt(p2_2, "Câu 2. Trong những cơn mưa dông kèm sấm sét, sự chuyển hoá của nitrogen đơn chất xảy ra theo sơ đồ sau:\n"
                                  "N~2~ ==(+O~2~, (1))==> NO ==(+O~2~, (2))==> NO~2~ ==(+O~2~+H~2~O, (3))==> HNO~3~ ==((4))==> H^+^ + NO~3~^-^\n"
                                  "Ion nitrate (NO~3~^-^) là một dạng đạm cây trồng hấp thụ được để sinh trưởng, phát triển.")
    p2_2_stmts = [
        ("a) Trong phản ứng (1) N~2~ đóng vai trò là chất khử.", "Đúng"),
        ("b) Để tạo ra được 124 kg ion nitrate cần dùng ở điều kiện chuẩn là 991,6 m^3^ khí N~2~. Biết hiệu suất của toàn bộ quá trình là 5%.", "Sai"),
        ("c) Sơ đồ cũng áp dụng để giải thích sự tạo thành HNO~3~ từ hoạt động đốt cháy nhiên liệu bằng không khí của phương tiện giao thông vận tải.", "Đúng"),
        ("d) Năng lượng trong sấm sét gây ra sự phân li của HNO~3~ trong quá trình (4).", "Sai")
    ]
    for stmt_text, ans in p2_2_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt_text, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Ý a) ĐÚNG: Ở phản ứng (1): N~2~ + O~2~ ⇌ 2NO (ở 3000 °C hoặc tia lửa điện). Số oxi hóa của N tăng từ 0 lên +2, do đó N~2~ đóng vai trò là chất khử.\n"
            "- Ý b) SAI: Tính toán lượng khí N~2~ cần dùng:\n"
            "  + Khối lượng mol ion nitrate NO~3~^-^ là M = 62 g/mol. Số mol ion nitrate cần tạo: n~NO3-~ = 124 000 / 62 = 2000 mol = 2 kmol.\n"
            "  + Sơ đồ bảo toàn nguyên tố N: N~2~ ➔ 2NO~3~^-^ (1 mol N~2~ sinh ra 2 mol NO~3~^-^).\n"
            "  + Số mol N~2~ theo lí thuyết (H = 100%): n~N2 lt~ = 2000 / 2 = 1000 mol = 1 kmol.\n"
            "  + Với hiệu suất H = 5% = 0,05: n~N2 tt~ = 1000 / 0,05 = 20 000 mol = 20 kmol.\n"
            "  + Thể tích khí N~2~ ở điều kiện chuẩn (đkc, 25 °C, 1 bar): V~N2~ = 20 000 × 24,79 = 495 800 lít = 495,8 m^3^ (chứ không phải 991,6 m^3^; 991,6 m^3^ là giá trị sai khi quên chia 2 cho hệ số của N trong phân tử N~2~).\n"
            "- Ý c) ĐÚNG: Trong động cơ đốt trong của ô tô, xe máy, nhiệt độ buồng đốt rất cao khiến N~2~ và O~2~ trong không khí phản ứng tạo ra NO, tiếp tục bị oxi hóa thành NO~2~ và kết hợp với hơi nước tạo thành acid HNO~3~ gây mưa acid.\n"
            "- Ý d) SAI: Quá trình (4) là sự điện li của acid mạnh HNO~3~ trong dung dịch nước mưa: HNO~3~ ➔ H^+^ + NO~3~^-^. Phản ứng này tự động xảy ra hoàn toàn trong dung dịch nước mà KHÔNG cần năng lượng của tia sét. Năng lượng cực lớn của sấm sét chỉ cần thiết cho phản ứng (1) để bẻ gãy liên kết ba N≡N rất bền vững của phân tử N~2~.\n"
            "-> *Đáp án: a) ĐÚNG - b) SAI - c) ĐÚNG - d) SAI*"
        )

    # ==========================================
    # PHẦN III. TRẢ LỜI NGẮN (1,0 ĐIỂM)
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm trả lời ngắn: (4 câu - 1 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4."
    )

    # Câu 1 Phần III
    p3_1 = doc.add_paragraph()
    render_question_prompt(p3_1, "Câu 1. Trong công nghiệp, người ta sản xuất nitric acid (HNO~3~) từ ammonia theo sơ đồ chuyển hoá sau:\n"
                                  "NH~3~ ==(+O~2~, t°, xt)==> NO ==(+O~2~)==> NO~2~ ==(+O~2~+H~2~O)==> HNO~3~\n"
                                  "Để điều chế 8,4 tấn dung dịch nitric acid có nồng độ 60% cần dùng bao nhiêu tấn ammonia? Biết rằng hiệu suất của toàn bộ quá trình sản xuất nitric acid theo sơ đồ trên là 80%.")
    p3_1_ans = doc.add_paragraph()
    if is_teacher:
        render_rich_text(p3_1_ans, "*Trả lời: 1,7*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "- Khối lượng chất tan HNO~3~ nguyên chất cần điều chế:\n"
            "  m~HNO3~ = 8,4 × 60% = 5,04 tấn.\n"
            "- Số mol HNO~3~ cần điều chế:\n"
            "  n~HNO3~ = 5,04 / 63 = 0,08 kmol (triệu mol).\n"
            "- Sơ đồ phản ứng bảo toàn nguyên tố N: NH~3~ ➔ HNO~3~\n"
            "  Theo lý thuyết: n~NH3 lt~ = n~HNO3~ = 0,08 kmol.\n"
            "- Vì hiệu suất toàn bộ quá trình là H = 80% = 0,8 nên lượng ammonia thực tế cần dùng là:\n"
            "  n~NH3 tt~ = 0,08 / 0,8 = 0,1 kmol.\n"
            "- Khối lượng ammonia thực tế cần dùng:\n"
            "  m~NH3~ = 0,1 × 17 = 1,7 tấn.\n"
            "-> *Đáp số: 1,7*"
        )
    else:
        render_rich_text(p3_1_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # Câu 2 Phần III
    p3_2 = doc.add_paragraph()
    render_question_prompt(p3_2, "Câu 2. Hòa tan hết 0,8 gam NaOH vào nước để được 2000 mL dung dịch ở 25 °C. pH của dung dịch thu được bằng mấy?")
    p3_2_ans = doc.add_paragraph()
    if is_teacher:
        render_rich_text(p3_2_ans, "*Trả lời: 12*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "- Số mol NaOH hòa tan: n~NaOH~ = 0,8 / 40 = 0,02 mol.\n"
            "- Thể tích dung dịch: V = 2000 mL = 2 lít.\n"
            "- Nồng độ mol của dung dịch NaOH:\n"
            "  C~M(NaOH)~ = 0,02 / 2 = 0,01 M = 10^-2^ M.\n"
            "- NaOH là base mạnh phân li hoàn toàn: NaOH ➔ Na^+^ + OH^-^\n"
            "  ➔ [OH^-^] = 0,01 M = 10^-2^ M.\n"
            "- pOH = -lg[OH^-^] = -lg(10^-2^) = 2.\n"
            "- Ở 25 °C: pH = 14 - pOH = 14 - 2 = 12.\n"
            "-> *Đáp số: 12*"
        )
    else:
        render_rich_text(p3_2_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # Câu 3 Phần III
    p3_3 = doc.add_paragraph()
    render_question_prompt(p3_3, "Câu 3. Cho các chất (phân tử và ion): NaOH, H~2~SO~4~, NH~3~, CO~3~^2-^, S^2-^, NH~4~^+^. Theo quan điểm của Bronsted – Lowry có bao nhiêu chất (phân tử và ion) có thể là base?")
    p3_3_ans = doc.add_paragraph()
    if is_teacher:
        render_rich_text(p3_3_ans, "*Trả lời: 4*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Theo thuyết acid - base của Brønsted – Lowry, base là chất hoặc ion có khả năng nhận proton (H^+^). Xét từng tiểu phân:\n"
            "1. NaOH: trong dung dịch phân li tạo ion OH^-^, ion OH^-^ nhận proton (OH^-^ + H^+^ ➔ H~2~O) ➔ là base. (1)\n"
            "2. H~2~SO~4~: là acid cho proton (H~2~SO~4~ ➔ H^+^ + HSO~4~^-^).\n"
            "3. NH~3~: có cặp electron tự do nhận proton (NH~3~ + H^+^ ⇌ NH~4~^+^) ➔ là base. (2)\n"
            "4. CO~3~^2-^: nhận proton (CO~3~^2-^ + H^+^ ⇌ HCO~3~^-^) ➔ là base. (3)\n"
            "5. S^2-^: nhận proton (S^2-^ + H^+^ ⇌ HS^-^) ➔ là base. (4)\n"
            "6. NH~4~^+^: là acid nhường proton (NH~4~^+^ ⇌ NH~3~ + H^+^).\n"
            "Vậy có tất cả 4 chất và ion đóng vai trò là base: NaOH, NH~3~, CO~3~^2-^, S^2-^.\n"
            "-> *Đáp số: 4*"
        )
    else:
        render_rich_text(p3_3_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # Câu 4 Phần III
    p3_4 = doc.add_paragraph()
    render_question_prompt(p3_4, "Câu 4. Cho cân bằng hoá học sau: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g); Δ~r~H°~298~ < 0\n"
                                  "Người ta thực hiện các biện pháp sau:\n"
                                  "(1) Giảm nhiệt độ,\n"
                                  "(2) Tăng áp suất chung của hệ phản ứng,\n"
                                  "(3) Dùng thêm chất xúc tác V~2~O~5~,\n"
                                  "(4) Tăng nồng độ SO~3~,\n"
                                  "(5) Tăng nồng độ SO~2~.\n"
                                  "Có bao nhiêu biện pháp làm cân bằng trên chuyển dịch theo chiều thuận?")
    p3_4_ans = doc.add_paragraph()
    if is_teacher:
        render_rich_text(p3_4_ans, "*Trả lời: 3*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Xét sự chuyển dịch cân bằng theo nguyên lí Le Chatelier:\n"
            "- Phản ứng thuận tỏa nhiệt (Δ~r~H°~298~ < 0) và làm giảm số mol khí (từ 3 mol khí về 2 mol khí).\n"
            "(1) Giảm nhiệt độ: Cân bằng chuyển dịch theo chiều tỏa nhiệt để làm tăng nhiệt độ, tức là CHIỀU THUẬN. ➔ (Thỏa mãn)\n"
            "(2) Tăng áp suất chung của hệ: Cân bằng chuyển dịch theo chiều làm giảm số mol khí, tức là CHIỀU THUẬN. ➔ (Thỏa mãn)\n"
            "(3) Thêm xúc tác V~2~O~5~: Chất xúc tác chỉ làm tăng tốc độ đạt tới trạng thái cân bằng, KHÔNG làm chuyển dịch cân bằng. ➔ (Không thỏa mãn)\n"
            "(4) Tăng nồng độ SO~3~ (sản phẩm): Cân bằng chuyển dịch theo chiều làm giảm SO~3~, tức là CHIỀU NGHỊCH. ➔ (Không thỏa mãn)\n"
            "(5) Tăng nồng độ SO~2~ (chất đầu): Cân bằng chuyển dịch theo chiều làm giảm SO~2~, tức là CHIỀU THUẬN. ➔ (Thỏa mãn)\n"
            "Vậy có 3 biện pháp làm cân bằng chuyển dịch theo chiều thuận là: (1), (2), (5).\n"
            "-> *Đáp số: 3*"
        )
    else:
        render_rich_text(p3_4_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # ==========================================
    # PHẦN IV. TỰ LUẬN (3,0 ĐIỂM)
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN IV. TỰ LUẬN: (3 câu - 3 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 3."
    )

    # Câu 1 Tự luận
    p_tl1 = doc.add_paragraph()
    render_question_prompt(p_tl1, "Câu 1. (1,0 điểm) Hoàn thành các phương trình phản ứng sau (ghi rõ điều kiện phản ứng)\n"
                                   "a) N~2~ + O~2~ ➔\n"
                                   "b) NH~3~ + HCl ➔\n"
                                   "c) HNO~3 đặc~ + FeO ➔\n"
                                   "d) Nhiệt phân NH~4~HCO~3~ ➔")
    if is_teacher:
        add_solution_no_table(doc,
            "Các phương trình phản ứng hoàn chỉnh (mỗi phương trình 0,25 điểm):\n"
            "a) N~2~(g) + O~2~(g) ⇌ 2NO(g)   (Điều kiện: ≈ 3000 °C hoặc tia lửa điện)\n"
            "b) NH~3~(g) + HCl(g) ➔ NH~4~Cl(s)   (Hiện tượng sinh ra khói trắng ammonium chloride)\n"
            "c) 4HNO~3~(đặc) + FeO ➔ Fe(NO~3~)~3~ + NO~2~↑ + 2H~2~O   (Có gia nhiệt t°)\n"
            "d) NH~4~HCO~3~(s) ➔ NH~3~↑ + CO~2~↑ + H~2~O   (Điều kiện: nhiệt phân t°)"
        )
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Tự luận
    p_tl2 = doc.add_paragraph()
    render_question_prompt(p_tl2, "Câu 2. (1,0 điểm) Cho phản ứng tổng hợp ammonia ở trạng thái cân bằng tại 430 °C.\n"
                                   "N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g) .\n"
                                   "Ở trạng thái cân bằng, nồng độ của NH~3~, N~2~ và H~2~ lần lượt là 0,30 M; 0,05 M và 0,10 M.\n"
                                   "Tính hằng số cân bằng của phản ứng ở nhiệt độ trên.")
    if is_teacher:
        add_solution_no_table(doc,
            "- Biểu thức hằng số cân bằng nồng độ K~c~ của phản ứng tổng hợp ammonia là:\n"
            "  K~c~ = [NH~3~]^2^ / ([N~2~] × [H~2~]^3^)   (0,5 điểm)\n"
            "- Thay nồng độ ở trạng thái cân bằng của các chất vào biểu thức:\n"
            "  [NH~3~] = 0,30 M;  [N~2~] = 0,05 M;  [H~2~] = 0,10 M.\n"
            "  K~c~ = (0,30)^2^ / [0,05 × (0,10)^3^]\n"
            "       = 0,09 / [0,05 × 0,001]\n"
            "       = 0,09 / 0,00005 = 1800.   (0,5 điểm)\n"
            "➔ *Đáp số: Hằng số cân bằng của phản ứng ở 430 °C là K~c~ = 1800.*"
        )
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 3 Tự luận
    p_tl3 = doc.add_paragraph()
    render_question_prompt(p_tl3, "Câu 3. (1,0 điểm) Hydrochloric acid (HCl) là một chất được phát hiện trong dịch vị dạ dày có nồng độ 10^-4^ – 10^-3^ mol/L và độ pH duy trì ở mức 3 – 4 đối với người bình thường. Nếu thiếu acid trong dạ dày thì thức ăn không chuyển hóa được lâu dần gây suy nhược cơ thể, nếu dư acid lâu ngày thì nó sẽ phá hủy đường ruột gây viêm loét dạ dày. Khi cơ thể dư acid HCl, người bệnh có thể sử dụng thuốc giảm đau dạ dày có tên gọi là Nabica (thành phần chính là NaHCO~3~), được dùng để trung hòa bớt lượng acid dư trong dạ dày.\n"
                                   "Giả sử dịch vị dạ dày người bệnh chứa 1,5 lít dung dịch hỗn hợp thức ăn lỏng, trong đó chứa 0,09125 gam HCl. Nếu khả năng tiêu thụ NaHCO~3~ của cơ thể người bệnh là 70% thì khối lượng NaHCO~3~ người đó cần đưa vào cơ thể để duy trì độ pH trong dạ dày ở mức 3 là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc,
            "- Lượng HCl ban đầu có trong 1,5 lít dịch vị dạ dày người bệnh:\n"
            "  n~HCl ban đầu~ = 0,09125 / 36,5 = 0,0025 mol.   (0,25 điểm)\n\n"
            "- Để duy trì độ pH trong dạ dày ở mức 3:\n"
            "  [H^+^]~cần duy trì~ = 10^-3^ M = 0,001 M.\n"
            "  Số mol ion H^+^ (tương ứng HCl) còn lại trong 1,5 lít dịch vị là:\n"
            "  n~HCl còn lại~ = [H^+^] × V = 0,001 × 1,5 = 0,0015 mol.   (0,25 điểm)\n\n"
            "- Lượng acid HCl đã được trung hòa bởi thuốc NaHCO~3~:\n"
            "  n~HCl phản ứng~ = n~HCl ban đầu~ - n~HCl còn lại~ = 0,0025 - 0,0015 = 0,001 mol.\n\n"
            "- Phương trình hóa học của phản ứng trung hòa dạ dày:\n"
            "  NaHCO~3~ + HCl ➔ NaCl + CO~2~↑ + H~2~O\n"
            "  Theo phương trình: n~NaHCO3 phản ứng~ = n~HCl phản ứng~ = 0,001 mol.\n"
            "  Khối lượng NaHCO~3~ thực tế đã phản ứng là:\n"
            "  m~NaHCO3 phản ứng~ = 0,001 × 84 = 0,084 gam.   (0,25 điểm)\n\n"
            "- Vì khả năng tiêu thụ NaHCO~3~ của cơ thể người bệnh là 70% (hiệu suất H = 70% = 0,7), khối lượng NaHCO~3~ người bệnh cần đưa vào cơ thể là:\n"
            "  m~NaHCO3 cần đưa vào~ = m~NaHCO3 phản ứng~ / 70% = 0,084 / 0,7 = 0,12 gam.   (0,25 điểm)\n\n"
            "➔ *Đáp số: Cần đưa vào cơ thể 0,12 gam NaHCO~3~.*"
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
    out_filename = f"De_Kiem_Tra_GK1_Hoa_11_NguyenChiThanh_Ma1002_{suffix}.docx"
    out_path = os.path.join(OUTPUT_DIR, out_filename)
    doc.save(out_path)
    print(f"Đã tạo thành công file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print("Bắt đầu tạo bộ tài liệu Đề Kiểm Tra Giữa Kì 1 Hóa Học 11 - THPT Nguyễn Chí Thanh - Mã Đề 1002...")
    generate_document(is_teacher=False)
    generate_document(is_teacher=True)
    print("Hoàn thành xuất sắc cả 2 phiên bản Word!")
