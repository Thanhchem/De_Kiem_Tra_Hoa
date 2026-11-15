# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì I Hóa Học 11 - THPT Nguyễn Huệ - Mã đề 1204
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Tạo 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_11_NguyenHue_Ma1204_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_11_NguyenHue_Ma1204_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
IMG_CHUANDO_PATH = r"C:\Antigravity_Thanh\San_Pham\images_gk1_hoa11_ma1204\hinh_chuan_do_c5_hd.png"
IMG_DUNGCU_PATH = r"C:\Antigravity_Thanh\San_Pham\images_gk1_hoa11_ma1204\dung_cu_thi_nghiem_hd.png"

def add_quick_answer_table(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu.
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
    p1_ans = ["B", "C", "D", "B", "A", "B", "C", "A", "B", "A", "B", "B"]
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
        ("Câu 1", ["ĐÚNG", "SAI", "SAI", "ĐÚNG"]),
        ("Câu 2", ["ĐÚNG", "ĐÚNG", "SAI", "SAI"])
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
    p3_answers = ["3", "124", "51,5", "4312"]
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
        school_line="TRƯỜNG THPT NGUYỄN HUỆ\nTỔ LÝ – HOÁ – CN",
        exam_title="KIỂM TRA GIỮA KỲ I – NĂM HỌC 2026 – 2027",
        subject_line="Môn: HOÁ HỌC, Lớp 11",
        num_pages="04",
        duration="45 phút (Không kể thời gian phát đề)",
        exam_code="1204",
        atomic_masses="H = 1; N = 14; O = 16; K = 39; Ca = 40",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table(doc)

    # ==========================================
    # A. PHẦN TRẮC NGHIỆM (7 ĐIỂM)
    # ==========================================
    p_a = doc.add_paragraph()
    p_a.paragraph_format.space_before = Pt(4)
    p_a.paragraph_format.space_after = Pt(2)
    render_rich_text(p_a, "*A. PHẦN TRẮC NGHIỆM (7 ĐIỂM)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    add_section_header(
        doc,
        title="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi học sinh chỉ chọn một phương án."
    )

    # Câu 1
    p1 = doc.add_paragraph()
    render_question_prompt(p1, "Câu 1. Mưa acid là hiện tượng nước mưa có pH thấp hơn 5,6. Hai acid chủ yếu được hình thành trong khí quyển và hòa tan vào nước mưa, gây ra mưa acid là")
    render_options_no_table(doc, [
        "A. HCl, HNO~3~.",
        "*B. H~2~SO~4~, HNO~3~.*",
        "C. SO~2~, NO~x~.",
        "D. H~2~SO~4~, H~2~S."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Khí SO~2~ và các oxit của nitơ (NO~x~) phát thải vào khí quyển bị oxi hóa bởi oxygen với sự xúc tác của các ion kim loại hoặc bụi mịn, sau đó hòa tan trong nước mưa tạo thành hai acid chính gây mưa acid là H~2~SO~4~ và HNO~3~.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 2
    p2 = doc.add_paragraph()
    render_question_prompt(p2, "Câu 2. Cho phương trình: CO~3~^2-^ + H~2~O ⇌ HCO~3~^-^ + OH^-^\n"
                                "Trong phản ứng thuận, theo thuyết Brønsted – Lowry chất nào là base?")
    render_options_no_table(doc, [
        "A. HCO~3~^-^.",
        "B. OH^-^.",
        "*C. CO~3~^2-^.*",
        "D. H~2~O."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Theo thuyết Brønsted – Lowry, base là chất nhận proton (H^+^). Trong phản ứng thuận: CO~3~^2-^ nhận H^+^ từ phân tử H~2~O để biến đổi thành HCO~3~^-^, do đó CO~3~^2-^ đóng vai trò là một base (H~2~O nhường H^+^ nên là acid).\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 3
    p3 = doc.add_paragraph()
    render_question_prompt(p3, "Câu 3. Dung dịch X có [OH^-^] = 10^-6^ M. pH của dung dịch X là")
    render_options_no_table(doc, [
        "A. 5.",
        "B. 9.",
        "C. 6.",
        "*D. 8.*"
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Ta có: pOH = -lg[OH^-^] = -lg(10^-6^) = 6.\n"
            "Do pH + pOH = 14 nên pH = 14 - pOH = 14 - 6 = 8.\n"
            "-> *Đáp án đúng: D*"
        )

    # Câu 4
    p4 = doc.add_paragraph()
    render_question_prompt(p4, "Câu 4. Tại nhiệt độ không đổi, khi một phản ứng thuận nghịch đạt trạng thái cân bằng thì")
    render_options_no_table(doc, [
        "A. tốc độ phản ứng hoá học xảy ra chậm dần.",
        "*B. nồng độ của các chất trong hỗn hợp phản ứng không thay đổi.*",
        "C. phản ứng hoá học không xảy ra.",
        "D. nồng độ của các chất trong hỗn hợp phản ứng vẫn liên tục thay đổi."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Tại trạng thái cân bằng hóa học, tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch (v~t~ = v~n~ > 0). Đây là cân bằng động, các phản ứng thuận và nghịch vẫn diễn ra liên tục với tốc độ bằng nhau nên nồng độ các chất trong hệ phản ứng được giữ nguyên không thay đổi.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 5
    p5 = doc.add_paragraph()
    render_question_prompt(p5, "Câu 5. Chuẩn độ 10 mL dung dịch NaOH x M bằng dung dịch chuẩn HCl 0,1 M với chất chỉ thị phenolphthalein như hình vẽ dưới đây:")
    
    # Chèn hình vẽ chuẩn độ
    if os.path.exists(IMG_CHUANDO_PATH):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(4)
        r_img = p_img.add_run()
        r_img.add_picture(IMG_CHUANDO_PATH, width=Inches(2.5))

    p5_sub = doc.add_paragraph()
    render_rich_text(p5_sub, "Sau ba lần chuẩn độ, thể tích dung dịch HCl 0,1 M đã dùng có giá trị trung bình là 15 mL. Giá trị của x là", base_size=10.5)

    render_options_no_table(doc, [
        "*A. 0,15 M.*",
        "B. 0,10 M.",
        "C. 0,20 M.",
        "D. 0,30 M."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Phương trình phản ứng chuẩn độ: NaOH + HCl ➔ NaCl + H~2~O\n"
            "Tại điểm tương đương (chỉ thị đổi màu): n~NaOH~ = n~HCl~\n"
            "➔ C~NaOH~ × V~NaOH~ = C~HCl~ × V~HCl~\n"
            "➔ x × 10 = 0,1 × 15 ➔ x = (0,1 × 15) / 10 = 0,15 M.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 6
    p6 = doc.add_paragraph()
    render_question_prompt(p6, "Câu 6. Ở trạng thái tự nhiên, nitrogen")
    render_options_no_table(doc, [
        "A. tự do chiếm khoảng 20% thể tích không khí.",
        "*B. tồn tại ở dạng đơn chất và hợp chất.*",
        "C. chỉ tồn tại ở dạng đơn chất.",
        "D. chỉ tồn tại ở dạng hợp chất."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Trong tự nhiên, nitrogen tồn tại ở cả dạng đơn chất (khí N~2~ chiếm khoảng 78% thể tích khí quyển) và dạng hợp chất (như khoáng vật diêm tiêu NaNO~3~, protein, nucleic acid trong cơ thể động thực vật,...).\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 7
    p7 = doc.add_paragraph()
    render_question_prompt(p7, "Câu 7. Các gói bim bim được làm phồng lên bằng khí X không màu, không mùi, trơ về mặt hóa học giúp tăng thời gian bảo quản và giữ được hình dạng sản phẩm không bị vỡ vụn. X là khí nào sau đây?")
    render_options_no_table(doc, [
        "A. NH~3~.",
        "B. O~2~.",
        "*C. N~2~.*",
        "D. O~3~."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Khí nitrogen (N~2~) có liên kết ba N≡N với năng lượng liên kết rất lớn (945 kJ/mol) nên trơ về mặt hóa học ở nhiệt độ thường. N~2~ không màu, không mùi, không độc, được bơm vào bao bì thực phẩm như bim bim, đồ ăn vặt để tạo môi trường trơ chống oxi hóa chất béo và bảo vệ cơ học tránh vỡ nát.\n"
            "-> *Đáp án đúng: C*"
        )

    # Câu 8
    p8 = doc.add_paragraph()
    render_question_prompt(p8, "Câu 8. Phản ứng nào sau đây là phản ứng thuận nghịch?")
    render_options_no_table(doc, [
        "*A. N~2~ + O~2~ ⇌ 2NO (t°).* ",
        "B. C~2~H~5~OH + 3O~2~ ➔ 2CO~2~ + 3H~2~O (t°).",
        "C. 2KClO~3~ ➔ 2KCl + 3O~2~ (t°).",
        "D. Mg + 2HCl ➔ MgCl~2~ + H~2~."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Phản ứng giữa nitrogen và oxygen xảy ra ở khoảng 3000°C hoặc khi có tia lửa điện là phản ứng thuận nghịch: N~2~(g) + O~2~(g) ⇌ 2NO(g). Các phản ứng cháy của cồn, nhiệt phân muối hay kim loại tác dụng với acid đều là phản ứng một chiều.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 9
    p9 = doc.add_paragraph()
    render_question_prompt(p9, "Câu 9. Hòa tan một ít tinh thể CH~3~COONa vào ống nghiệm chứa nước, có tồn tại cân bằng sau:\n"
                                "CH~3~COONa + H~2~O ⇌ CH~3~COOH + NaOH (1).\n"
                                "Thêm vài giọt phenolphthalein vào ống nghiệm thấy dung dịch chuyển sang màu hồng nhạt, tiến hành ngâm ống nghiệm trong nước nóng màu dung dịch đậm dần. Kết luận nào sau đây chính xác?")
    render_options_no_table(doc, [
        "A. Khi tăng nhiệt độ, cân bằng (1) chuyển dịch theo chiều nghịch.",
        "*B. Nếu ngâm ống nghiệm trong nước đá, màu dung dịch nhạt dần.*",
        "C. Chiều thuận của phản ứng (1) là quá trình tỏa nhiệt.",
        "D. Cân bằng (1) không phụ thuộc vào nhiệt độ."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "- Màu hồng của phenolphthalein do môi trường kiềm (ion OH^-^ / NaOH) tạo ra.\n"
            "- Khi ngâm trong nước nóng (tăng nhiệt độ), màu dung dịch đậm dần chứng tỏ nồng độ OH^-^ tăng lên ➔ cân bằng (1) chuyển dịch theo chiều thuận.\n"
            "- Theo nguyên lí Le Chatelier, khi tăng nhiệt độ cân bằng chuyển dịch theo chiều thu nhiệt ➔ Chiều thuận là quá trình thu nhiệt (ΔH > 0), chiều nghịch là quá trình tỏa nhiệt.\n"
            "- Do đó, khi ngâm trong nước đá (hạ nhiệt độ), cân bằng chuyển dịch theo chiều tỏa nhiệt (chiều nghịch), làm nồng độ OH^-^ giảm ➔ màu hồng của dung dịch nhạt dần. Kết luận B hoàn toàn chính xác.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 10
    p10 = doc.add_paragraph()
    render_question_prompt(p10, "Câu 10. Phân tử ammonia được tạo thành bởi một nguyên tử nitrogen liên kết với ba nguyên tử hydrogen. Dạng hình học của phân tử ammonia là")
    render_options_no_table(doc, [
        "*A. chóp tam giác.*",
        "B. tứ giác.",
        "C. tam giác đều.",
        "D. vuông phẳng."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "Nguyên tử nitrogen trong NH~3~ ở trạng thái lai hóa sp^3^, liên kết với 3 nguyên tử H và còn lại một cặp electron tự do không liên kết. Lực đẩy của cặp electron này đẩy các liên kết N-H làm cho phân tử NH~3~ có dạng hình chóp tam giác (tháp tam giác) với góc liên kết H-N-H khoảng 107°.\n"
            "-> *Đáp án đúng: A*"
        )

    # Câu 11
    p11 = doc.add_paragraph()
    render_question_prompt(p11, "Câu 11. Chất nào sau đây khi hòa tan vào nước không phân li ra ion ?")
    render_options_no_table(doc, [
        "A. CH~3~COOH.",
        "*B. C~2~H~5~OH.*",
        "C. HCl.",
        "D. NaCl."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "C~2~H~5~OH (ethanol) là chất không điện li, khi tan vào nước tồn tại dưới dạng phân tử ngậm nước, không phân li ra các ion mang điện tích. CH~3~COOH là chất điện li yếu; HCl và NaCl là các chất điện li mạnh.\n"
            "-> *Đáp án đúng: B*"
        )

    # Câu 12
    p12 = doc.add_paragraph()
    render_question_prompt(p12, "Câu 12. Trong quá trình đốt cháy nhiên liệu hóa thạch như than đá và dầu nặng, thường sản sinh các khí NO, NO~2~... Tên gọi của NO~2~ là")
    render_options_no_table(doc, [
        "A. ammonia.",
        "*B. nitrogen dioxide.*",
        "C. dinitrogen oxide.",
        "D. nitrogen monoxide."
    ], is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc,
            "NO~2~ có tên gọi theo danh pháp quốc tế là nitrogen dioxide. (NO là nitrogen monoxide; N~2~O là dinitrogen oxide; NH~3~ là ammonia).\n"
            "-> *Đáp án đúng: B*"
        )

    # ==========================================
    # PHẦN II. TRẮC NGHIỆM ĐÚNG SAI (2 CÂU)
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai (2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai."
    )

    # Câu 1 Phần II
    p2_1 = doc.add_paragraph()
    render_question_prompt(p2_1, "Câu 1. Phương trình hóa học của phản ứng tổng hợp ammonia từ nitrogen và hydrogen bằng quá trình Haber như sau:\n"
                                  "N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)   Δ~r~H°~298~ = - 91,8 kJ (*)\n"
                                  "(điều kiện: t°, p, xúc tác Fe)")

    p2_1_stmts = [
        ("a) Liên kết hoá học trong phân tử NH~3~ là liên kết cộng hoá trị có cực.", "Đúng"),
        ("b) Do có hàm lượng nitrogen cao (82,35% theo khối lượng) nên ammonia được sử dụng trực tiếp làm phân đạm rất hiệu quả.", "Sai"),
        ("c) Có thể nhận biết muối ammonium bằng cách cho muối tác dụng với dung dịch kiềm thấy thoát ra chất khí làm quỳ tím ẩm hoá đỏ.", "Sai"),
        ("d) Phản ứng (*) là phản ứng tỏa nhiệt. Để tăng hiệu suất tổng hợp NH~3~ người ta cần thực hiện ở nhiệt độ thấp nhưng không quá thấp để tốc độ phản ứng không quá chậm.", "Đúng")
    ]
    for stmt_text, ans in p2_1_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt_text, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(doc,
            "- Ý a) ĐÚNG: Trong phân tử NH~3~, liên kết giữa N và H là liên kết cộng hóa trị có cực, cặp electron dùng chung lệch về phía nguyên tử nitrogen có độ âm điện lớn hơn.\n"
            "- Ý b) SAI: Ammonia ở điều kiện thường là chất khí có tính ăn mòn, mùi khai xốc độc hại và dễ bay hơi gây thất thoát N, nên ammonia KHÔNG được sử dụng trực tiếp làm phân đạm mà được dùng làm nguyên liệu để sản xuất các loại phân đạm như urê, đạm amoni, đạm nitrat.\n"
            "- Ý c) SAI: Khí thoát ra là NH~3~ có tính base, làm quỳ tím ẩm HÓA XANH (chứ không phải hóa đỏ):\n"
            "  NH~4~^+^ + OH^-^ ➔ NH~3~↑ + H~2~O.\n"
            "- Ý d) ĐÚNG: Phản ứng có Δ~r~H°~298~ = -91,8 kJ < 0 nên là phản ứng tỏa nhiệt. Theo nguyên lí Le Chatelier, hạ nhiệt độ giúp cân bằng chuyển dịch theo chiều thuận (tăng hiệu suất). Tuy nhiên nếu nhiệt độ quá thấp tốc độ phản ứng rất chậm, do đó thực tế người ta chọn nhiệt độ tối ưu khoảng 400 - 450°C kết hợp xúc tác Fe."
        )

    # Câu 2 Phần II
    p2_2 = doc.add_paragraph()
    render_question_prompt(p2_2, "Câu 2. Ở các vùng quê, người dân thường dùng phèn chua (K~2~SO~4~.Al~2~(SO~4~)~3~.24H~2~O hay KAl(SO~4~)~2~.12H~2~O) để làm trong nước. Khi phèn chua tan vào nước thì ion Al^3+^ bị thủy phân theo phản ứng:\n"
                                  "Al^3+^ + 3H~2~O ⇌ Al(OH)~3~ + 3H^+^ (*)")

    p2_2_stmts = [
        ("a) Phèn chua là chất điện li.", "Đúng"),
        ("b) Dung dịch sau khi hoà tan phèn chua vào nước có pH < 7.", "Đúng"),
        ("c) Dung dịch phèn chua làm phenolphthalein chuyển sang màu hồng.", "Sai"),
        ("d) Phèn chua được dùng để làm trong nước nhờ ứng dụng của phản ứng thuỷ phân ion Al^3+^. Các bụi bẩn sẽ bị cuốn theo kết tủa keo trắng Al(OH)~3~ lắng xuống đáy nên nước sẽ trong lại. Trong phản ứng (*) Al^3+^ là base và H~2~O là acid.", "Sai")
    ]
    for stmt_text, ans in p2_2_stmts:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        render_statement(p, stmt_text, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(doc,
            "- Ý a) ĐÚNG: Phèn chua là muối kép của kim loại, khi tan vào nước phân li hoàn toàn thành các ion K^+^, Al^3+^, SO~4~^2-^ nên là chất điện li mạnh.\n"
            "- Ý b) ĐÚNG: Do phản ứng thủy phân của ion Al^3+^ tạo ra môi trường chứa ion H^+^ (Al^3+^ + 3H~2~O ⇌ Al(OH)~3~ + 3H^+^), nồng độ [H^+^] > 10^-7^ M nên dung dịch có tính acid, pH < 7.\n"
            "- Ý c) SAI: Phenolphthalein chỉ chuyển màu hồng trong môi trường base (pH > 8,3). Dung dịch phèn chua có tính acid nên phenolphthalein không đổi màu.\n"
            "- Ý d) SAI: Mặc dù giải thích cơ chế làm trong nước bằng kết tủa keo Al(OH)~3~ lắng cặn là đúng, nhưng nhận định cuối cùng 'Al^3+^ là base và H~2~O là acid' là SAI! Trong phản ứng thủy phân theo thuyết Brønsted – Lowry, ion Al^3+^ (dưới dạng phức aqua [Al(H~2~O)~6~]^3+^) nhường proton H^+^ nên đóng vai trò là một ACID, còn H~2~O nhận proton H^+^ nên đóng vai trò là BASE."
        )

    # ==========================================
    # PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN (4 CÂU)
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm trả lời ngắn (2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4."
    )

    # Câu 1 Phần III
    p3_1 = doc.add_paragraph()
    render_question_prompt(p3_1, "Câu 1. Trong thí nghiệm chuẩn độ dung dịch NaOH bằng dung dịch HCl, người ta cho từ từ dung dịch NaOH từ burette vào bình tam giác chứa dung dịch HCl và vài giọt dung dịch phenolphthalein. Trong các dụng cụ thí nghiệm được đánh số dưới đây, burette là dụng cụ số mấy?")
    
    # Chèn hình ảnh 5 dụng cụ
    if os.path.exists(IMG_DUNGCU_PATH):
        p_img3 = doc.add_paragraph()
        p_img3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img3.paragraph_format.space_before = Pt(4)
        p_img3.paragraph_format.space_after = Pt(4)
        r_img3 = p_img3.add_run()
        r_img3.add_picture(IMG_DUNGCU_PATH, width=Inches(4.5))

    p3_1_ans = doc.add_paragraph()
    p3_1_ans.paragraph_format.space_before = Pt(2)
    p3_1_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p3_1_ans, "*Trả lời: 3*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Nhận diện các dụng cụ thí nghiệm trên hình:\n"
            "  • Dụng cụ (1): Ống đong (Measuring cylinder).\n"
            "  • Dụng cụ (2): Ống hút nhỏ giọt / pipet hút (Dropper pipette).\n"
            "  • Dụng cụ (3): Burette (buret kẹp trên giá sắt dùng để đo chính xác thể tích dung dịch chuẩn khi chuẩn độ).\n"
            "  • Dụng cụ (4): Bình tam giác / bình nón (Erlenmeyer flask).\n"
            "  • Dụng cụ (5): Cốc thủy tinh (Beaker).\n"
            "Do đó burette là dụng cụ số 3.\n"
            "-> *Đáp án: 3*"
        )
    else:
        render_rich_text(p3_1_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # Câu 2 Phần III
    p3_2 = doc.add_paragraph()
    render_question_prompt(p3_2, "Câu 2. Cho các ứng dụng sau:\n"
                                  "(1) Sản xuất phân đạm urea ((NH~2~)~2~CO).\n"
                                  "(2) Được sử dụng làm chất làm lạnh trong thiết bị lạnh.\n"
                                  "(3) Sản xuất nitrogen.\n"
                                  "(4) Sản xuất nitric acid.\n"
                                  "Hãy liệt kê các ứng dụng đúng của ammonia theo thứ tự tăng dần? (Ví dụ: 123, 234...)")
    p3_2_ans = doc.add_paragraph()
    p3_2_ans.paragraph_format.space_before = Pt(2)
    p3_2_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p3_2_ans, "*Trả lời: 124*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "- Ứng dụng (1) ĐÚNG: Ammonia dùng để sản xuất phân urê: 2NH~3~ + CO~2~ ➔ (NH~2~)~2~CO + H~2~O.\n"
            "- Ứng dụng (2) ĐÚNG: Ammonia dễ hóa lỏng, có nhiệt hóa hơi lớn nên được ứng dụng làm môi chất lạnh (chất sinh hàn) trong các hệ thống làm lạnh công nghiệp.\n"
            "- Ứng dụng (3) SAI: Trong công nghiệp người ta sản xuất nitrogen bằng cách chưng cất phân đoạn không khí lỏng, không sử dụng ammonia để sản xuất nitrogen.\n"
            "- Ứng dụng (4) ĐÚNG: Ammonia là nguyên liệu chủ yếu để điều chế nitric acid qua quá trình oxi hóa Ostwald (NH~3~ ➔ NO ➔ NO~2~ ➔ HNO~3~).\n"
            "Các ứng dụng đúng gồm: (1), (2), (4). Viết theo thứ tự tăng dần là 124.\n"
            "-> *Đáp án: 124*"
        )
    else:
        render_rich_text(p3_2_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # Câu 3 Phần III
    p3_3 = doc.add_paragraph()
    render_question_prompt(p3_3, "Câu 3. Cho phản ứng thuận nghịch sau ở nhiệt độ 430°C: H~2~(g) + I~2~(g) ⇌ 2HI(g).\n"
                                  "Nồng độ mol các chất ở thời điểm cân bằng như sau: [H~2~] = [I~2~] = 0,107 M; [HI] = 0,768 M.\n"
                                  "Tính hằng số cân bằng K~c~ của phản ứng ở 430°C (làm tròn kết quả đến hàng phần mười).")
    p3_3_ans = doc.add_paragraph()
    p3_3_ans.paragraph_format.space_before = Pt(2)
    p3_3_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p3_3_ans, "*Trả lời: 51,5*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Biểu thức tính hằng số cân bằng của phản ứng:\n"
            "  K~c~ = [HI]^2^ / ([H~2~] × [I~2~])\n"
            "Thay số vào biểu thức:\n"
            "  K~c~ = (0,768)^2^ / (0,107 × 0,107) = 0,589824 / 0,011449 ≈ 51,5175...\n"
            "Làm tròn kết quả đến hàng phần mười ta được: 51,5.\n"
            "-> *Đáp án: 51,5*"
        )
    else:
        render_rich_text(p3_3_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # Câu 4 Phần III
    p3_4 = doc.add_paragraph()
    render_question_prompt(p3_4, "Câu 4. Cho các phát biểu sau:\n"
                                  "(1) Ánh sáng mặt trời bị cản trở làm ảnh hưởng đến quá trình quang hợp, gây thiếu oxygen làm cho thực vật và động vật chết.\n"
                                  "(2) Sự phân hủy xác động thực vật bởi vi khuẩn sử dụng nhiều oxygen trong nước gây nên tình trạng thiếu oxygen nghiêm trọng, làm chết cả hệ sinh thái.\n"
                                  "(3) Chất dinh dưỡng giúp thực vật và tảo sống trong nước phát triển ồ ạt.\n"
                                  "(4) Phân bón và chất dinh dưỡng bị rửa trôi xuống sông, ao, hồ.\n"
                                  "Hãy sắp xếp các phát biểu trên theo đúng trình tự diễn biến quá trình hình thành hiện tượng phú dưỡng thành bộ bốn số (ví dụ: 2143, 4213, ...)?")
    p3_4_ans = doc.add_paragraph()
    p3_4_ans.paragraph_format.space_before = Pt(2)
    p3_4_ans.paragraph_format.space_after = Pt(2)
    if is_teacher:
        render_rich_text(p3_4_ans, "*Trả lời: 4312*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(doc,
            "Trình tự diễn biến khoa học chuẩn của hiện tượng phú dưỡng (Eutrophication):\n"
            "  • Bước 1 - (4): Nguồn phân bón dư thừa và chất thải chứa nhiều hợp chất của nitrogen, phosphorus bị nước mưa rửa trôi xuống sông, ao, hồ.\n"
            "  • Bước 2 - (3): Lượng lớn chất dinh dưỡng kích thích thực vật thủy sinh và tảo sống trong nước phát triển bùng phát, ồ ạt.\n"
            "  • Bước 3 - (1): Lớp tảo dày đặc che phủ mặt nước làm ánh sáng mặt trời bị cản trở, cản quang hợp và ức chế khuếch tán không khí gây thiếu oxygen làm một phần sinh vật bắt đầu chết.\n"
            "  • Bước 4 - (2): Quá trình phân hủy một lượng khổng lồ xác động thực vật bởi vi khuẩn tiêu tốn gần như cạn kiệt nguồn oxygen hòa tan, dẫn đến thiếu oxygen nghiêm trọng và làm chết suy thoái cả hệ sinh thái nước.\n"
            "Vậy trình tự diễn biến đúng là: (4) ➔ (3) ➔ (1) ➔ (2). Bộ bốn số là 4312.\n"
            "-> *Đáp án: 4312*"
        )
    else:
        render_rich_text(p3_4_ans, "Trả lời: .....................................................................................................................................", base_size=10)

    # ==========================================
    # B. PHẦN TỰ LUẬN (3 ĐIỂM)
    # ==========================================
    p_b = doc.add_paragraph()
    p_b.paragraph_format.space_before = Pt(6)
    p_b.paragraph_format.space_after = Pt(2)
    render_rich_text(p_b, "*B. PHẦN TỰ LUẬN (3 ĐIỂM)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # Câu 1 Tự luận
    p_tl1 = doc.add_paragraph()
    render_question_prompt(p_tl1, "Câu 1: (1,0 điểm) Cho cân bằng hoá học (trong bình kín) sau:\n"
                                   "2NO~2~(g) (màu nâu đỏ) ⇌ N~2~O~4~(g) (không màu) (*)\n"
                                   "a) Cân bằng (*) chuyển dịch theo chiều nào (thuận hay nghịch) khi tăng áp suất của hỗn hợp ở điều kiện nhiệt độ không đổi? Giải thích ngắn gọn.\n"
                                   "b) Khi hạ nhiệt độ của bình thì màu nâu đỏ nhạt dần. Phản ứng thuận là toả nhiệt hay thu nhiệt? (không giải thích).")
    if is_teacher:
        add_solution_no_table(doc,
            "a) (0,5 điểm)\n"
            "- Khi tăng áp suất của hệ, theo nguyên lí Le Chatelier, cân bằng sẽ chuyển dịch theo chiều làm giảm áp suất, tức là chiều làm giảm số mol khí.\n"
            "- Xét phản ứng: vế trái có 2 mol khí (2NO~2~), vế phải có 1 mol khí (1N~2~O~4~). Số mol khí giảm theo chiều thuận (từ 2 mol về 1 mol).\n"
            "➔ Do đó, khi tăng áp suất, cân bằng (*) chuyển dịch theo CHIỀU THUẬN.\n\n"
            "b) (0,5 điểm)\n"
            "- Khi hạ nhiệt độ, màu nâu đỏ nhạt dần chứng tỏ nồng độ NO~2~ giảm và nồng độ N~2~O~4~ tăng, nghĩa là cân bằng chuyển dịch theo chiều thuận.\n"
            "- Theo nguyên lí Le Chatelier, khi hạ nhiệt độ, cân bằng chuyển dịch theo chiều tỏa nhiệt để bù đắp nhiệt lượng bị mất đi.\n"
            "➔ Do đó, phản ứng thuận là phản ứng TỎA NHIỆT."
        )
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Tự luận
    p_tl2 = doc.add_paragraph()
    render_question_prompt(p_tl2, "Câu 2: (1,0 điểm) Nước thải công nghiệp trước khi xả ra môi trường cần được thu gom và xử lý đạt tiêu chuẩn theo quy định. Theo quy chuẩn kỹ thuật quốc gia về nước thải công nghiệp, giá trị pH của nước thải công nghiệp khi xả vào nguồn tiếp nhận phải nằm trong khoảng từ 5,5 đến 9,0. Một mẫu nước thải của nhà máy X có pH = 4. Để đảm bảo đạt tiêu chuẩn xả thải, nhà máy sử dụng vôi sống (CaO) cho vào nước thải để nâng pH từ 4 lên 7.\n"
                                   "a) Hãy cho biết mẫu nước thải của nhà máy X có môi trường là acid, base hay trung tính?\n"
                                   "b) Giải thích tác dụng của vôi sống trong việc làm tăng pH của nước thải bằng phương trình hoá học?\n"
                                   "c) Tính khối lượng vôi sống (kilogam) cần sử dụng để xử lý 6000 m^3^ nước thải nói trên? (Bỏ qua sự thủy phân của các muối nếu có). (Không làm tròn kết quả các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng phần mười).")
    if is_teacher:
        add_solution_no_table(doc,
            "a) (0,25 điểm)\n"
            "Mẫu nước thải có pH = 4 < 7 nên mẫu nước thải có MÔI TRƯỜNG ACID.\n\n"
            "b) (0,25 điểm)\n"
            "- Khi cho vôi sống (CaO) vào nước thải, CaO phản ứng với nước tạo thành Ca(OH)~2~ tan phân li giải phóng ion OH^-^:\n"
            "  CaO + H~2~O ➔ Ca(OH)~2~\n"
            "  Ca(OH)~2~ ➔ Ca^2+^ + 2OH^-^\n"
            "- Ion OH^-^ sinh ra sẽ phản ứng trung hòa ion H^+^ dư trong nước thải acid:\n"
            "  H^+^ + OH^-^ ➔ H~2~O   (hoặc viết: CaO + 2H^+^ ➔ Ca^2+^ + H~2~O)\n"
            "- Sự tiêu thụ bớt ion H^+^ làm giảm nồng độ [H^+^], do đó giá trị pH = -lg[H^+^] của nước thải tăng lên.\n\n"
            "c) (0,5 điểm)\n"
            "- Đổi thể tích nước thải: V = 6000 m^3^ = 6 000 000 L = 6 × 10^6^ L.\n"
            "- Ở pH = 4 ➔ [H^+^]~ban đầu~ = 10^-4^ M.\n"
            "- Khi nâng pH lên 7 (môi trường trung tính) ➔ [H^+^]~lúc sau~ = 10^-7^ M.\n"
            "- Lượng ion H^+^ cần được trung hòa là:\n"
            "  n~H+ cần trung hòa~ = ([H^+^]~ban đầu~ - [H^+^]~lúc sau~) × V\n"
            "  n~H+~ = (10^-4^ - 10^-7^) × 6 000 000 = 600 - 0,6 = 599,4 mol.\n"
            "  (Nếu tính gần đúng bỏ qua 10^-7^ rất bé so với 10^-4^ thì n~H+~ ≈ 10^-4^ × 6 000 000 = 600 mol).\n"
            "- Theo phương trình: CaO + 2H^+^ ➔ Ca^2+^ + H~2~O\n"
            "  ➔ n~CaO~ = 1/2 × n~H+~ = 599,4 / 2 = 299,7 mol.\n"
            "- Khối lượng vôi sống (CaO với M = 40 + 16 = 56 g/mol) cần dùng:\n"
            "  m~CaO~ = 299,7 × 56 = 16 783,2 gam = 16,7832 kg.\n"
            "  (Nếu dùng n~CaO~ = 300 mol thì m~CaO~ = 300 × 56 = 16 800 g = 16,8 kg).\n"
            "- Làm tròn đến hàng phần mười theo yêu cầu đề bài: m~CaO~ = 16,8 kg.\n"
            "➔ Đáp số: 16,8 kg."
        )
    else:
        add_student_writing_lines(doc, num_lines=12)

    # Câu 3 Tự luận
    p_tl3 = doc.add_paragraph()
    render_question_prompt(p_tl3, "Câu 3: (1,0 điểm) Quá trình hình thành đạm nitrate cho đất từ nước mưa được biểu diễn theo sơ đồ sau:\n"
                                   "N~2~ ==(1)==> NO ==(2)==> NO~2~ ==(3)==> HNO~3~ ==(4)==> H^+^ + NO~3~^-^\n"
                                   "a) Viết phương trình hóa học của các phản ứng xảy ra (ghi rõ điều kiện nếu có, mỗi mũi tên là một phản ứng hóa học).\n"
                                   "b) Một cơn mưa rơi xuống một thửa ruộng với lượng mưa trung bình là 5 m^3^/giờ. Hãy tính khối lượng nguyên tố nitrogen (tính bằng gam) mà thửa ruộng nhận được khi cơn mưa đó kéo dài 3 giờ? (Giả sử hàm lượng ion nitrate trong nước mưa khoảng 62 mg/L và hàm lượng nitrate trong nước mưa không đổi trong suốt thời gian mưa).")
    if is_teacher:
        add_solution_no_table(doc,
            "a) (0,5 điểm)\n"
            "Các phương trình hóa học xảy ra theo sơ đồ:\n"
            "  (1) N~2~(g) + O~2~(g) ⇌ 2NO(g)   (ở nhiệt độ ≈ 3000°C hoặc có tia lửa điện do sấm sét)\n"
            "  (2) 2NO(g) + O~2~(g) ➔ 2NO~2~(g)   (xảy ra ngay ở điều kiện thường)\n"
            "  (3) 4NO~2~(g) + O~2~(g) + 2H~2~O(l) ➔ 4HNO~3~(aq)\n"
            "  (4) HNO~3~(aq) ➔ H^+^(aq) + NO~3~^-^(aq)   (acid mạnh phân li hoàn toàn trong nước)\n\n"
            "b) (0,5 điểm)\n"
            "- Thể tích nước mưa rơi xuống thửa ruộng trong 3 giờ là:\n"
            "  V = 5 m^3^/giờ × 3 giờ = 15 m^3^ = 15 000 L.\n"
            "- Tổng khối lượng ion nitrate (NO~3~^-^) có trong lượng nước mưa đó là:\n"
            "  m~NO3-~ = 62 mg/L × 15 000 L = 930 000 mg = 930 gam.\n"
            "- Khối lượng mol: M(NO~3~^-^) = 14 + 16 × 3 = 62 g/mol; M(N) = 14 g/mol.\n"
            "- Áp dụng định luật bảo toàn nguyên tố đối với nitrogen (trong 1 mol NO~3~^-^ có 1 mol nguyên tử N):\n"
            "  m~N~ = m~NO3-~ × (M~N~ / M~NO3-~) = 930 × (14 / 62) = 210 gam.\n"
            "➔ Khối lượng nguyên tố nitrogen mà thửa ruộng nhận được là 210 gam."
        )
    else:
        add_student_writing_lines(doc, num_lines=12)

    # Chân trang thông báo kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "*---------- HẾT ----------*", base_size=10.5, base_bold=True)

    p_note_end = doc.add_paragraph()
    p_note_end.paragraph_format.space_before = Pt(2)
    p_note_end.paragraph_format.space_after = Pt(2)
    p_note_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_note_end, "- Thí sinh không được sử dụng tài liệu.\n- Giám thị không giải thích gì thêm.", base_size=9.5, base_italic=True)

    # Lưu file
    suffix = "GiaoVien" if is_teacher else "HocSinh"
    out_filename = f"De_Kiem_Tra_GK1_Hoa_11_NguyenHue_Ma1204_{suffix}.docx"
    out_path = os.path.join(OUTPUT_DIR, out_filename)
    doc.save(out_path)
    print(f"Đã tạo thành công file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print("Bắt đầu tạo bộ tài liệu Đề Kiểm Tra Giữa Kì 1 Hóa Học 11 - THPT Nguyễn Huệ - Mã Đề 1204...")
    generate_document(is_teacher=False)
    generate_document(is_teacher=True)
    print("Hoàn thành xuất sắc cả 2 phiên bản Word!")
