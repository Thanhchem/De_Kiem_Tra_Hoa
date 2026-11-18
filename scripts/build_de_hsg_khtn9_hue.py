# -*- coding: utf-8 -*-
"""
Script tạo Đề thi Học sinh giỏi Cấp Thành phố THCS - TP Huế 2025 - 2026
Môn: Khoa học tự nhiên - Mã đề thi: 1203
Gồm 2 phiên bản:
1. Bản Học sinh (De_Thi_HSG_KHTN9_TPHue_2025_2026_HocSinh.docx)
2. Bản Giáo viên (De_Thi_HSG_KHTN9_TPHue_2025_2026_GiaoVien.docx)
Áp dụng 100% chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
"""

import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

import sys
sys.path.append(r'c:\Antigravity_Thanh\scripts')
from form_de_kiem_tra import (
    COLOR_HF_RED, COLOR_HF_GREEN_HEX, COLOR_PRIMARY, COLOR_CORRECT,
    COLOR_SOLUTION, COLOR_GRAY, COLOR_BORDER_GREY_HEX,
    add_p_border_bottom, add_p_border_top, set_run_style,
    add_page_number_field, render_rich_text, setup_header_footer,
    render_question_prompt, render_single_option, render_options_no_table,
    render_statement, add_solution_no_table, add_student_writing_lines
)

IMG_DIR = r'c:\Antigravity_Thanh\San_Pham\images_khtn9_hue'
OUTPUT_DIR = r'c:\Antigravity_Thanh\San_Pham'

def add_centered_image(doc, img_path, width_cm=6.5):
    """Chèn hình ảnh căn giữa, dãn dòng hợp lý."""
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run()
        run.add_picture(img_path, width=Cm(width_cm))
    else:
        print(f"Warning: Image not found: {img_path}")

def build_exam_title_block(doc, is_teacher=False):
    """Khối tiêu đề chính xác theo đề thi TP Huế."""
    headers_data = [
        ("UBND THÀNH PHỐ HUẾ", "KỲ THI CHỌN HỌC SINH GIỎI THÀNH PHỐ", 10, True, 10.5, True),
        ("SỞ GIÁO DỤC VÀ ĐÀO TẠO", "TRUNG HỌC CƠ SỞ NĂM HỌC 2025 - 2026", 10, True, 10.5, True),
        ("ĐỀ THI CHÍNH THỨC", "MÔN: KHOA HỌC TỰ NHIÊN", 10.5, True, 11, True),
        ("(Đề thi gồm 04 trang)", "Thời gian: 120 phút (không kể thời gian giao đề) | Mã đề: 1203", 9.5, False, 10, True)
    ]
    for l_text, r_text, l_size, l_bold, r_size, r_bold in headers_data:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.tab_stops.clear_all()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(18.0), WD_TAB_ALIGNMENT.RIGHT)

        rl = p.add_run(l_text)
        set_run_style(rl, font_size=l_size, bold=l_bold, italic=(not l_bold))

        p.add_run('\t')
        rr = p.add_run(r_text)
        set_run_style(rr, font_size=r_size, bold=r_bold, italic=(not r_bold))

    if is_teacher:
        p_gv = doc.add_paragraph()
        p_gv.paragraph_format.space_before = Pt(3)
        p_gv.paragraph_format.space_after = Pt(3)
        p_gv.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rgv = p_gv.add_run("(BẢN GIÁO VIÊN -- CÓ ĐÁP ÁN VÀ LỜI GIẢI CHI TIẾT)")
        set_run_style(rgv, font_size=11, bold=True, color=COLOR_CORRECT)

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(4)
    p_info.paragraph_format.space_after = Pt(2)
    render_rich_text(p_info, "*Họ và tên học sinh:* ............................................................ *Số báo danh:* .........................", base_size=10.5)

    p_note1 = doc.add_paragraph()
    p_note1.paragraph_format.space_before = Pt(1)
    p_note1.paragraph_format.space_after = Pt(1)
    render_rich_text(p_note1, "*(Cho biết khối lượng nguyên tử (amu) của các nguyên tố: H = 1; He = 4; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Mn = 55; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108; Ba = 137).*", base_size=9.5, base_italic=True)

    p_note2 = doc.add_paragraph()
    p_note2.paragraph_format.space_before = Pt(0)
    p_note2.paragraph_format.space_after = Pt(4)
    render_rich_text(p_note2, "*(Các thể tích khí đều đo ở điều kiện chuẩn (đkc), giả thiết các khí sinh ra không tan trong nước).*", base_size=9.5, base_italic=True)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(0)
    p_div.paragraph_format.space_after = Pt(4)
    add_p_border_bottom(p_div, color=COLOR_BORDER_GREY_HEX, sz="4")

def build_data_table_cau1_p3(doc):
    """Bảng số liệu Câu 1 Phần III (ADN hai mạch)."""
    tbl = doc.add_table(rows=3, cols=5)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Cm(2.5), Cm(2.0), Cm(2.0), Cm(2.0), Cm(2.0)]
    
    headers = ["Mạch", "A", "T", "G", "C"]
    row1 = ["Mạch 1", "0,28", "0,12", "–", "–"]
    row2 = ["Mạch 2", "–", "–", "0,38", "–"]
    
    data = [headers, row1, row2]
    for r_idx, row in enumerate(tbl.rows):
        for c_idx, cell in enumerate(row.cells):
            cell.width = widths[c_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            txt = data[r_idx][c_idx]
            is_bold = (r_idx == 0 or c_idx == 0)
            r = p.add_run(txt)
            set_run_style(r, font_size=10, bold=is_bold)
            # cell borders
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                                  f'<w:top w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
                                  f'<w:left w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
                                  f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
                                  f'<w:right w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
                                  f'</w:tcBorders>')
            tcPr.append(tcBorders)
    
    # spacing after table
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(0)
    p_sp.paragraph_format.space_after = Pt(2)

def generate_exam_docx(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)
    build_exam_title_block(doc, is_teacher=is_teacher)

    # ==========================================
    # I. PHẦN CHUNG (6,0 ĐIỂM)
    # ==========================================
    p_sec_main = doc.add_paragraph()
    p_sec_main.paragraph_format.space_before = Pt(6)
    p_sec_main.paragraph_format.space_after = Pt(2)
    render_rich_text(p_sec_main, "*I. PHẦN CHUNG (6,0 điểm)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # ------------------------------------------
    # PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN
    # ------------------------------------------
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.*", base_size=11, base_bold=True, base_color=COLOR_PRIMARY)
    
    p_p1_sub = doc.add_paragraph()
    p_p1_sub.paragraph_format.space_before = Pt(0)
    p_p1_sub.paragraph_format.space_after = Pt(4)
    render_rich_text(p_p1_sub, "_Thí sinh trả lời từ Câu 1 đến Câu 6. Mỗi câu hỏi thí sinh chỉ chọn một phương án._", base_size=10.5, base_italic=True)

    # --- Câu 1 ---
    p_q1 = doc.add_paragraph()
    p_q1.paragraph_format.space_before = Pt(3)
    p_q1.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q1, "Câu 1. Cho các phát biểu sau:")
    
    q1_stmts = [
        "(a) Trong công nghiệp, aluminium được điều chế bằng cách nhiệt phân aluminium oxide.",
        "(b) Nước vôi trong là dung dịch calcium hydroxide.",
        "(c) Sodium hydrogencarbonate (NaHCO~3~) được dùng làm bột nở, thuốc giảm đau dạ dày.",
        "(d) Trong dãy các kim loại aluminium, copper, gold (Au) thì copper là kim loại có tính dẫn điện tốt nhất."
    ]
    for s in q1_stmts:
        p_s = doc.add_paragraph()
        p_s.paragraph_format.left_indent = Cm(0.5)
        p_s.paragraph_format.space_before = Pt(1)
        p_s.paragraph_format.space_after = Pt(1)
        render_rich_text(p_s, s, base_size=10.5)
    
    p_q1_end = doc.add_paragraph()
    p_q1_end.paragraph_format.space_before = Pt(2)
    p_q1_end.paragraph_format.space_after = Pt(2)
    render_rich_text(p_q1_end, "Số phát biểu đúng là", base_size=10.5)

    opts_c1 = ["A. 4.", "B. 1.", "*C. 3.*", "D. 2."] if is_teacher else ["A. 4.", "B. 1.", "C. 3.", "D. 2."]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)

    if is_teacher:
        sol_c1 = (
            "- Phát biểu (a) sai: Trong công nghiệp, aluminium (nhôm) được điều chế bằng cách điện phân nóng chảy aluminium oxide (Al~2~O~3~) với xúc tác criolit (Na~3~AlF~6~): 2Al~2~O~3~ -(đpnc, criolit)-> 4Al + 3O~2~↑ (Al~2~O~3~ rất bền nhiệt, không bị nhiệt phân).\n"
            "- Phát biểu (b) đúng: Nước vôi trong là dung dịch calcium hydroxide Ca(OH)~2~ trong nước.\n"
            "- Phát biểu (c) đúng: NaHCO~3~ dùng làm bột nở (bị nhiệt phân giải phóng CO~2~ làm xốp bánh) và dùng làm thuốc giảm đau dạ dày do trung hòa bớt acid dịch vị (HCl dư).\n"
            "- Phát biểu (d) đúng: Khả năng dẫn điện giảm dần theo thứ tự: Ag > Cu > Au > Al > Fe. Trong dãy aluminium (Al), copper (Cu) và gold (Au) thì copper (Cu) dẫn điện tốt nhất.\n"
            "=> Có 3 phát biểu đúng là (b), (c), (d)."
        )
        add_solution_no_table(doc, sol_c1)

    # --- Câu 2 ---
    p_q2 = doc.add_paragraph()
    p_q2.paragraph_format.space_before = Pt(4)
    p_q2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q2, "Câu 2. Cho các phát biểu sau:")

    q2_stmts = [
        "(a) Khí methane có trong thành phần của khí bùn ao, khí mỏ dầu, khí thiên nhiên.",
        "(b) Thủy phân hoàn toàn tinh bột và saccharose trong môi trường acid chỉ sinh ra glucose.",
        "(c) Trong thành phần của xăng sinh học E5 có ethanol.",
        "(d) Chất béo là triester của glycerol với các acid."
    ]
    for s in q2_stmts:
        p_s = doc.add_paragraph()
        p_s.paragraph_format.left_indent = Cm(0.5)
        p_s.paragraph_format.space_before = Pt(1)
        p_s.paragraph_format.space_after = Pt(1)
        render_rich_text(p_s, s, base_size=10.5)

    p_q2_end = doc.add_paragraph()
    p_q2_end.paragraph_format.space_before = Pt(2)
    p_q2_end.paragraph_format.space_after = Pt(2)
    render_rich_text(p_q2_end, "Số phát biểu đúng là", base_size=10.5)

    opts_c2 = ["A. 3.", "*B. 2.*", "C. 4.", "D. 1."] if is_teacher else ["A. 3.", "B. 2.", "C. 4.", "D. 1."]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)

    if is_teacher:
        sol_c2 = (
            "- Phát biểu (a) đúng: Khí methane (CH~4~) có trong thành phần khí thiên nhiên (~95%), khí mỏ dầu, khí bùn ao và biogas.\n"
            "- Phát biểu (b) sai: Thủy phân tinh bột chỉ tạo glucose, nhưng thủy phân saccharose thu được cả glucose và fructose.\n"
            "- Phát biểu (c) đúng: Xăng E5 chứa 5% ethanol (C~2~H~5~OH) về thể tích và 95% xăng khoáng truyền thống.\n"
            "- Phát biểu (d) sai: Theo SGK KHTN 9, chất béo là triester của glycerol với các *acid béo* (acid monocarboxylic mạch dài, không phân nhánh), không phải với *các acid* nói chung.\n"
            "=> Có 2 phát biểu đúng là (a) và (c)."
        )
        add_solution_no_table(doc, sol_c2)

    # --- Câu 3 ---
    p_q3 = doc.add_paragraph()
    p_q3.paragraph_format.space_before = Pt(4)
    p_q3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q3, "Câu 3. Chuyển động của một vật trên trục Ox được mô tả bằng đồ thị phụ thuộc quãng đường s theo thời gian t như hình vẽ bên. Trong những khoảng thời gian nào, vật chuyển động thẳng đều?")
    add_centered_image(doc, os.path.join(IMG_DIR, "cau3_dothi_st.png"), width_cm=5.5)

    opts_c3 = [
        "A. Trong khoảng thời gian từ 0 đến t~1~ và từ t~1~ đến t~2~.",
        "B. Trong khoảng thời gian từ t~1~ đến t~3~.",
        "*C. Trong khoảng thời gian từ 0 đến t~1~ và từ t~2~ đến t~3~.*",
        "D. Trong khoảng thời gian từ t~1~ đến t~2~."
    ] if is_teacher else [
        "A. Trong khoảng thời gian từ 0 đến t~1~ và từ t~1~ đến t~2~.",
        "B. Trong khoảng thời gian từ t~1~ đến t~3~.",
        "C. Trong khoảng thời gian từ 0 đến t~1~ và từ t~2~ đến t~3~.",
        "D. Trong khoảng thời gian từ t~1~ đến t~2~."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)

    if is_teacher:
        sol_c3 = (
            "- Từ 0 đến t~1~: Đồ thị là đường thẳng xiên dốc lên (s tăng tuyến tính theo thời gian t) => Vật chuyển động thẳng đều với vận tốc không đổi v~1~ > 0.\n"
            "- Từ t~1~ đến t~2~: Đồ thị là đoạn thẳng nằm ngang song song trục Ot (s không đổi) => Vật đứng yên (v = 0).\n"
            "- Từ t~2~ đến t~3~: Đồ thị là đoạn thẳng có độ dốc không đổi => Vật chuyển động thẳng đều.\n"
            "=> Vật chuyển động thẳng đều trong các khoảng thời gian từ 0 đến t~1~ và từ t~2~ đến t~3~."
        )
        add_solution_no_table(doc, sol_c3)

    # --- Câu 4 ---
    p_q4 = doc.add_paragraph()
    p_q4.paragraph_format.space_before = Pt(4)
    p_q4.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q4, "Câu 4. Dạng đột biến gene nào sau đây khi xảy ra có thể làm thay đổi số liên kết hydrogen nhưng không làm thay đổi số lượng nucleotide của gene?")

    opts_c4 = [
        "A. Đảo vị trí một số cặp nucleotide.",
        "B. Thêm một cặp nucleotide.",
        "C. Mất một cặp nucleotide.",
        "*D. Thay thế một cặp nucleotide khác.*"
    ] if is_teacher else [
        "A. Đảo vị trí một số cặp nucleotide.",
        "B. Thêm một cặp nucleotide.",
        "C. Mất một cặp nucleotide.",
        "D. Thay thế một cặp nucleotide khác."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)

    if is_teacher:
        sol_c4 = (
            "- Đột biến thêm hoặc mất cặp nucleotide làm tăng hoặc giảm số lượng nucleotide của gene.\n"
            "- Đột biến đảo vị trí các cặp nucleotide không làm thay đổi số lượng nucleotide và không làm đổi tổng số liên kết hydrogen của gene.\n"
            "- Đột biến thay thế một cặp nucleotide này bằng một cặp nucleotide khác (ví dụ: thay thế cặp A-T có 2 liên kết H bằng cặp G-C có 3 liên kết H hoặc ngược lại) làm thay đổi số liên kết hydrogen nhưng không làm thay đổi số lượng nucleotide của gene."
        )
        add_solution_no_table(doc, sol_c4)

    # --- Câu 5 ---
    p_q5 = doc.add_paragraph()
    p_q5.paragraph_format.space_before = Pt(4)
    p_q5.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q5, "Câu 5. Nam và Tuấn chơi trò \"điện thoại dây\" bằng hai ống nghe nối với nhau bởi một sợi chỉ bông dài 268 m. Khi Tuấn nói một từ, Nam nghe được âm thanh đó hai lần: lần thứ nhất do âm truyền qua sợi chỉ, và một lúc sau nghe lại lần thứ hai do âm truyền qua không khí. Biết rằng tốc độ truyền âm trong không khí là 340 m/s và tốc độ truyền âm trong sợi chỉ là 2 680 m/s. Khoảng thời gian giữa hai lần Nam nghe được từ đó là bao nhiêu?")
    add_centered_image(doc, os.path.join(IMG_DIR, "cau5_dienthoai.png"), width_cm=5.5)

    opts_c5 = ["A. 1,11 s.", "B. 0,79 s.", "*C. 0,69 s.*", "D. 1,50 s."] if is_teacher else ["A. 1,11 s.", "B. 0,79 s.", "C. 0,69 s.", "D. 1,50 s."]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)

    if is_teacher:
        sol_c5 = (
            "- Thời gian âm truyền qua sợi chỉ bông: t~1~ = s / v~chỉ~ = 268 / 2680 = 0,10 (s).\n"
            "- Thời gian âm truyền qua không khí: t~2~ = s / v~kk~ = 268 / 340 ≈ 0,7882 (s).\n"
            "- Khoảng thời gian giữa hai lần nghe thấy: Δt = t~2~ - t~1~ = 0,7882 - 0,10 = 0,6882 s ≈ 0,69 s."
        )
        add_solution_no_table(doc, sol_c5)

    # --- Câu 6 ---
    p_q6 = doc.add_paragraph()
    p_q6.paragraph_format.space_before = Pt(4)
    p_q6.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q6, "Câu 6. Hình bên mô tả số lượng NST ở các tế bào sinh dưỡng A và B của hai cơ thể thuộc một loài sinh vật lưỡng bội. Nếu tế bào A là tế bào của cơ thể bình thường, thì tế bào B thuộc thể đột biến nào sau đây?")
    add_centered_image(doc, os.path.join(IMG_DIR, "cau6_nst.png"), width_cm=6.0)

    opts_c6 = ["*A. Thể ba nhiễm.*", "B. Thể đa bội.", "C. Thể tam bội.", "D. Thể một nhiễm."] if is_teacher else ["A. Thể ba nhiễm.", "B. Thể đa bội.", "C. Thể tam bội.", "D. Thể một nhiễm."]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)

    if is_teacher:
        sol_c6 = (
            "- Tế bào A là tế bào bình thường có 3 cặp NST (mỗi cặp gồm 2 chiếc tương đồng) => Bộ NST lưỡng bội 2n = 6.\n"
            "- Tế bào B: có hai cặp bình thường (mỗi cặp 2 chiếc), riêng một cặp NST (dạng chữ V/mũi tên) có 3 chiếc => Thể đột biến có dạng (2n + 1) = 7 chiếc, đây là thể ba nhiễm (trisomy)."
        )
        add_solution_no_table(doc, sol_c6)

    # ------------------------------------------
    # PHẦN II. TRẮC NGHIỆM ĐÚNG SAI
    # ------------------------------------------
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(6)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*PHẦN II. Câu trắc nghiệm đúng sai.*", base_size=11, base_bold=True, base_color=COLOR_PRIMARY)

    p_p2_sub = doc.add_paragraph()
    p_p2_sub.paragraph_format.space_before = Pt(0)
    p_p2_sub.paragraph_format.space_after = Pt(4)
    render_rich_text(p_p2_sub, "_Thí sinh trả lời từ Câu 1 đến Câu 3. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai._", base_size=10.5, base_italic=True)

    # --- Câu 1 Phần II ---
    p_q1_p2 = doc.add_paragraph()
    p_q1_p2.paragraph_format.space_before = Pt(3)
    p_q1_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q1_p2, "Câu 1. Để phù hợp với các loại vải khác nhau, hệ thống điều chỉnh nhiệt của bàn là điện được thiết kế với nhiều mức nhiệt khác nhau. Sơ đồ mạch điện của hệ thống điều chỉnh nhiệt của bàn là được cho ở hình vẽ bên. Trên mạch có các chốt (1), (2), (3), (4) là những tiếp điểm dùng để nối các thanh dẫn có điện trở không đáng kể, nhằm thay đổi cách mắc các điện trở trong mạch và từ đó điều chỉnh công suất tỏa nhiệt của bàn là.")
    add_centered_image(doc, os.path.join(IMG_DIR, "cau1_p2_mach_ban_la.png"), width_cm=6.0)

    p_c1_sub = doc.add_paragraph()
    p_c1_sub.paragraph_format.space_before = Pt(1)
    p_c1_sub.paragraph_format.space_after = Pt(2)
    render_rich_text(p_c1_sub, "Các chế độ hoạt động của bàn là tương ứng với các loại vải thường gặp được thiết lập như sau:\n"
                               "+ Chế độ là vải lụa: nối tắt chốt (3) với chốt (4).\n"
                               "+ Chế độ là vải len: nối tắt chốt (1) với chốt (2), đồng thời nối chung chốt (3) với chốt (4).\n"
                               "+ Chế độ là vải bông: nối tắt chốt (1) với chốt (4) và chốt (1) với chốt (3).\n"
                               "+ Chế độ là vải dày: nối tắt chốt (2) với chốt (4) và chốt (1) với chốt (3).\n"
                               "Khi bàn là cắm vào nguồn điện có hiệu điện thế U = 220 V và hoạt động lần lượt ở các chế độ trên thì công suất tiêu thụ tương ứng của bàn là sẽ bằng bao nhiêu?", base_size=10.5)

    q1_p2_items = [
        ("a) Ở chế độ là vải lụa, công suất tiêu thụ của bàn là là 403,3 W.", "Đúng"),
        ("b) Ở chế độ là vải len, công suất tiêu thụ của bàn là là 1 210 W.", "Sai"),
        ("c) Ở chế độ là vải bông, công suất tiêu thụ của bàn là là 605 W.", "Sai"),
        ("d) Ở chế độ là vải dày, công suất tiêu thụ của bàn là là 1 815 W.", "Đúng")
    ]
    for stmt, ans in q1_p2_items:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.left_indent = Cm(0.5)
        p_item.paragraph_format.space_before = Pt(1)
        p_item.paragraph_format.space_after = Pt(1)
        render_statement(p_item, stmt, ans=ans, is_teacher=is_teacher, base_size=10.5)

    if is_teacher:
        sol_q1_p2 = (
            "- Chế độ vải lụa: Nối tắt (3) với (4) => Chốt (3) nối cực 1 nguồn điện, chốt (4) nối đầu trái R~1~ (80 Ω). Hai điện trở R~1~ và R~2~ (40 Ω) mắc nối tiếp: R~tđ~ = 80 + 40 = 120 Ω => P = U^2 / R~tđ~ = 220^2 / 120 ≈ 403,3 W => a) Đúng.\n"
            "- Chế độ vải len: Nối tắt (1) với (2) và (3) với (4) => Hai đầu R~2~ (40 Ω) bị nối tắt (đoản mạch). Mạch chỉ có R~1~ = 80 Ω hoạt động dưới U = 220 V => P = U^2 / 80 = 220^2 / 80 = 605 W (không phải 1 210 W) => b) Sai.\n"
            "- Chế độ vải bông: Nối tắt (1) với (4) và (1) với (3) => Cả hai đầu của R~1~ (80 Ω) đều nối cực 1 nguồn điện nên R~1~ bị đoản mạch. Mạch chỉ có R~2~ = 40 Ω hoạt động dưới U = 220 V => P = U^2 / 40 = 220^2 / 40 = 1 210 W (không phải 605 W) => c) Sai.\n"
            "- Chế độ vải dày: Nối tắt (2) với (4) và (1) với (3) => R~1~ và R~2~ mắc song song với nhau: R~tđ~ = (80 × 40)/(80 + 40) = 80/3 Ω => P = U^2 / R~tđ~ = 605 + 1210 = 1 815 W => d) Đúng."
        )
        add_solution_no_table(doc, sol_q1_p2)

    # --- Câu 2 Phần II ---
    p_q2_p2 = doc.add_paragraph()
    p_q2_p2.paragraph_format.space_before = Pt(4)
    p_q2_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q2_p2, "Câu 2. Cho ester X có công thức: CH~3~COOC~2~H~5~.")

    q2_p2_items = [
        ("a) Nhiệt độ sôi của X cao hơn so với butane.", "Đúng"),
        ("b) Phần trăm khối lượng nguyên tố oxygen trong X là 54,55%.", "Sai"),
        ("c) Thủy phân hoàn toàn X trong môi trường acid thu được hai chất Y và Z (M~Y~ < M~Z~), từ Y có thể điều chế trực tiếp được Z.", "Đúng"),
        ("d) Thủy phân hoàn toàn m gam X bằng dung dịch sodium hydroxide, đun nóng thu được 28,70 gam muối. Giá trị của m là 30,8 gam.", "Đúng")
    ]
    for stmt, ans in q2_p2_items:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.left_indent = Cm(0.5)
        p_item.paragraph_format.space_before = Pt(1)
        p_item.paragraph_format.space_after = Pt(1)
        render_statement(p_item, stmt, ans=ans, is_teacher=is_teacher, base_size=10.5)

    if is_teacher:
        sol_q2_p2 = (
            "- Ý a) Đúng: Butane (C~4~H~10~, M = 58) là chất khí ở điều kiện thường (t~sôi~ ≈ -0,5 °C), trong khi ethyl acetate X (C~4~H~8~O~2~, M = 88) là chất lỏng có nhóm ester phân cực nên nhiệt độ sôi cao hơn nhiều (t~sôi~ ≈ 77,1 °C).\n"
            "- Ý b) Sai: C~4~H~8~O~2~ có M = 88 g/mol; %m~O~ = (32 / 88) × 100% ≈ 36,36% (54,55% là %m~C~ = 48/88 × 100%).\n"
            "- Ý c) Đúng: Thủy phân X trong môi trường acid sinh ra ethanol C~2~H~5~OH (M~Y~ = 46) và acetic acid CH~3~COOH (M~Z~ = 60). Từ ethanol có thể lên men giấm trực tiếp tạo acetic acid: C~2~H~5~OH + O~2~ -(men giấm)-> CH~3~COOH + H~2~O.\n"
            "- Ý d) Đúng: Phản ứng xà phòng hóa: CH~3~COOC~2~H~5~ + NaOH -> CH~3~COONa + C~2~H~5~OH. Muối là CH~3~COONa (M = 82 g/mol) => n~muối~ = 28,70 / 82 = 0,35 mol => n~X~ = 0,35 mol => m = 0,35 × 88 = 30,8 gam."
        )
        add_solution_no_table(doc, sol_q2_p2)

    # --- Câu 3 Phần II ---
    p_q3_p2 = doc.add_paragraph()
    p_q3_p2.paragraph_format.space_before = Pt(4)
    p_q3_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q3_p2, "Câu 3. Ở thực vật, xét hai cặp gene A, a và B, b quy định hai tính trạng khác nhau, các allele trội hoàn toàn, các gene phân li độc lập. Khi giao phấn giữa hai cây dị hợp hai cặp gene (P) thu được F~1~.")

    q3_p2_items = [
        ("a) Khi (P) giao phấn thì F~1~ có 8 loại kiểu gene.", "Sai"),
        ("b) Các cây chứa hai tính trạng trội ở F~1~ có tỉ lệ tối đa là 50%.", "Sai"),
        ("c) Tỉ lệ kiểu gene dị hợp ở F~1~ có tỉ lệ 75%.", "Đúng"),
        ("d) Cho các cây trội hai tính trạng ở F~1~ giao phấn ngẫu nhiên có thể thu được F~2~ mang hai tính trạng lặn là 1/81.", "Đúng")
    ]
    for stmt, ans in q3_p2_items:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.left_indent = Cm(0.5)
        p_item.paragraph_format.space_before = Pt(1)
        p_item.paragraph_format.space_after = Pt(1)
        render_statement(p_item, stmt, ans=ans, is_teacher=is_teacher, base_size=10.5)

    if is_teacher:
        sol_q3_p2 = (
            "- Ý a) Sai: P: AaBb × AaBb => Số loại kiểu gene ở F~1~ là 3 × 3 = 9 loại kiểu gene (không phải 8).\n"
            "- Ý b) Sai: Tỉ lệ kiểu hình mang 2 tính trạng trội A-B- ở F~1~ là (3/4) × (3/4) = 9/16 = 56,25% > 50%.\n"
            "- Ý c) Đúng: Tỉ lệ đồng hợp cả 2 cặp gene (AABB, AAbb, aaBB, aabb) là 4 × (1/16) = 25% => Tỉ lệ dị hợp ở F~1~ = 100% - 25% = 75%.\n"
            "- Ý d) Đúng: Các cây mang 2 tính trạng trội A-B- ở F~1~ gồm (1/3 AA : 2/3 Aa)(1/3 BB : 2/3 Bb). Tần số allele a = 1/3 và tần số allele b = 1/3. Giao phấn ngẫu nhiên cho đời F~2~ có tỉ lệ kiểu hình 2 tính trạng lặn (aabb) = (1/3)^2 × (1/3)^2 = 1/9 × 1/9 = 1/81."
        )
        add_solution_no_table(doc, sol_q3_p2)

    # ------------------------------------------
    # PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN
    # ------------------------------------------
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(6)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*PHẦN III. Câu trắc nghiệm trả lời ngắn.*", base_size=11, base_bold=True, base_color=COLOR_PRIMARY)

    p_p3_sub = doc.add_paragraph()
    p_p3_sub.paragraph_format.space_before = Pt(0)
    p_p3_sub.paragraph_format.space_after = Pt(4)
    render_rich_text(p_p3_sub, "_Thí sinh trả lời từ Câu 1 đến Câu 6._", base_size=10.5, base_italic=True)

    # --- Câu 1 Phần III ---
    p_c1_p3 = doc.add_paragraph()
    p_c1_p3.paragraph_format.space_before = Pt(3)
    p_c1_p3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_c1_p3, "Câu 1. Một phân tử nucleic acid mạch kép có tỉ lệ từng loại nucleotide của từng mạch được thể hiện ở bảng sau (dấu \"–\" thể hiện chưa xác định số liệu). Biết rằng tổng tỉ lệ 4 loại nucleotide trên mỗi mạch đơn là 1,00.")
    build_data_table_cau1_p3(doc)

    p_c1_p3_end = doc.add_paragraph()
    p_c1_p3_end.paragraph_format.space_before = Pt(1)
    p_c1_p3_end.paragraph_format.space_after = Pt(2)
    render_rich_text(p_c1_p3_end, "Theo lí thuyết, cytosine (C) ở mạch 2 chiếm tỉ lệ bao nhiêu? (kết quả làm tròn đến chữ số hàng phần trăm)", base_size=10.5)

    if not is_teacher:
        p_ans_box = doc.add_paragraph()
        p_ans_box.paragraph_format.space_before = Pt(1)
        p_ans_box.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans_box, "*Trả lời:* .....................................................................................................", base_size=10.5)
    else:
        sol_c1_p3 = (
            "**Đáp số: 0,22**\n"
            "- Theo NTBS: A~1~ = T~2~; T~1~ = A~2~; G~1~ = C~2~; C~1~ = G~2~.\n"
            "- Ta có: G~2~ = 0,38 => C~1~ = 0,38.\n"
            "- Trên mạch 1: A~1~ + T~1~ + G~1~ + C~1~ = 1,00 => 0,28 + 0,12 + G~1~ + 0,38 = 1,00 => G~1~ = 1,00 - 0,78 = 0,22.\n"
            "- Vậy tỉ lệ cytosine ở mạch 2 là C~2~ = G~1~ = 0,22 (hay 22%)."
        )
        add_solution_no_table(doc, sol_c1_p3)

    # --- Câu 2 Phần III ---
    p_c2_p3 = doc.add_paragraph()
    p_c2_p3.paragraph_format.space_before = Pt(4)
    p_c2_p3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_c2_p3, "Câu 2. Cho dung dịch barium hydroxide đến dư lần lượt vào các dung dịch sau: NaHCO~3~, AlCl~3~, NaHSO~4~, NH~4~Cl, FeCl~3~, Na~2~SO~4~ và Na~3~PO~4~. Có bao nhiêu trường hợp thu được kết tủa sau khi kết thúc thí nghiệm?")

    if not is_teacher:
        p_ans_box = doc.add_paragraph()
        p_ans_box.paragraph_format.space_before = Pt(1)
        p_ans_box.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans_box, "*Trả lời:* .....................................................................................................", base_size=10.5)
    else:
        sol_c2_p3 = (
            "**Đáp số: 5**\n"
            "- Dung dịch NaHCO~3~: Ba(OH)~2~ + NaHCO~3~ -> BaCO~3~↓ (trắng) + NaOH + H~2~O. (Có kết tủa)\n"
            "- Dung dịch AlCl~3~: Al(OH)~3~ ban đầu sinh ra bị Ba(OH)~2~ dư hòa tan hết: 2Al(OH)~3~ + Ba(OH)~2~ -> Ba[Al(OH)~4~]~2~ (tan). (Không có kết tủa)\n"
            "- Dung dịch NaHSO~4~: Ba(OH)~2~ + NaHSO~4~ -> BaSO~4~↓ (trắng) + NaOH + H~2~O. (Có kết tủa)\n"
            "- Dung dịch NH~4~Cl: Ba(OH)~2~ + 2NH~4~Cl -> BaCl~2~ + 2NH~3~↑ + 2H~2~O. (Không có kết tủa)\n"
            "- Dung dịch FeCl~3~: 3Ba(OH)~2~ + 2FeCl~3~ -> 2Fe(OH)~3~↓ (nâu đỏ) + 3BaCl~2~. (Có kết tủa)\n"
            "- Dung dịch Na~2~SO~4~: Ba(OH)~2~ + Na~2~SO~4~ -> BaSO~4~↓ (trắng) + 2NaOH. (Có kết tủa)\n"
            "- Dung dịch Na~3~PO~4~: 3Ba(OH)~2~ + 2Na~3~PO~4~ -> Ba~3~(PO~4~)~2~↓ (trắng) + 6NaOH. (Có kết tủa)\n"
            "=> Có 5 trường hợp thu được kết tủa (NaHCO~3~, NaHSO~4~, FeCl~3~, Na~2~SO~4~, Na~3~PO~4~)."
        )
        add_solution_no_table(doc, sol_c2_p3)

    # --- Câu 3 Phần III ---
    p_c3_p3 = doc.add_paragraph()
    p_c3_p3.paragraph_format.space_before = Pt(4)
    p_c3_p3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_c3_p3, "Câu 3. Đun nóng 54,0 gam dung dịch glucose nồng độ a% với lượng dư dung dịch silver nitrate trong dung dịch ammonia. Sau khi phản ứng xảy ra hoàn toàn thu được 6,48 gam silver (Ag). Giá trị của a là bao nhiêu?")

    if not is_teacher:
        p_ans_box = doc.add_paragraph()
        p_ans_box.paragraph_format.space_before = Pt(1)
        p_ans_box.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans_box, "*Trả lời:* .....................................................................................................", base_size=10.5)
    else:
        sol_c3_p3 = (
            "**Đáp số: 10**\n"
            "- Số mol Ag: n~Ag~ = 6,48 / 108 = 0,06 mol.\n"
            "- 1 mol glucose (C~6~H~12~O~6~) tráng bạc sinh ra 2 mol Ag: n~glucose~ = n~Ag~ / 2 = 0,06 / 2 = 0,03 mol.\n"
            "- Khối lượng glucose nguyên chất: m~glucose~ = 0,03 × 180 = 5,4 gam.\n"
            "- Nồng độ a% = (5,4 / 54,0) × 100% = 10% => a = 10."
        )
        add_solution_no_table(doc, sol_c3_p3)

    # --- Câu 4 Phần III ---
    p_c4_p3 = doc.add_paragraph()
    p_c4_p3.paragraph_format.space_before = Pt(4)
    p_c4_p3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_c4_p3, "Câu 4. Một vật bắt đầu trượt không ma sát từ đầu đỉnh một máng nghiêng. Chọn mốc thế năng tại chân máng nghiêng B, bỏ qua mọi lực cản. Tại vị trí thế năng bằng 2 lần động năng của vật trên máng nghiêng, tốc độ của vật là 240 cm/s. Khi vật đến chân của máng nghiêng, vật ngay lập tức đi vào một đoạn đường tròn bán kính R = 40 cm nằm trong mặt phẳng thẳng đứng và trượt lên trên như hình vẽ bên. Tốc độ của vật tại vị trí M trên đoạn đường tròn khi α = 30° bằng bao nhiêu m/s? (kết quả làm tròn đến chữ số hàng phần trăm)")
    add_centered_image(doc, os.path.join(IMG_DIR, "cau4_p3_mang_nghieng.png"), width_cm=6.5)

    if not is_teacher:
        p_ans_box = doc.add_paragraph()
        p_ans_box.paragraph_format.space_before = Pt(1)
        p_ans_box.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans_box, "*Trả lời:* .....................................................................................................", base_size=10.5)
    else:
        sol_c4_p3 = (
            "**Đáp số: 4,03**\n"
            "- Chọn mốc thế năng tại chân máng B. Bỏ qua ma sát và lực cản, cơ năng W bảo toàn.\n"
            "- Tại vị trí W~t1~ = 2 W~đ1~: W = W~đ1~ + W~t1~ = 3 W~đ1~ = 3 × [ 0,5 m v~1~^2 ]. Với v~1~ = 2,4 m/s => W = 1,5 × m × (2,4)^2 = 8,64m (J).\n"
            "- Tại vị trí M trên đường tròn bán kính R = 0,4 m ứng với góc α = 30° so với phương thẳng đứng OB:\n"
            "  Độ cao điểm M: z~M~ = R(1 - cos α) = 0,4 × (1 - cos 30°) ≈ 0,05359 m.\n"
            "  Thế năng tại M: W~tM~ = m × g × z~M~ ≈ 0,53m (J) (với g = 9,8 - 10 m/s^2).\n"
            "- Bảo toàn cơ năng: W~đM~ = W - W~tM~ => 0,5 m v~M~^2 = 8,64m - 0,53m = 8,11m => v~M~^2 ≈ 16,22 => v~M~ ≈ 4,03 m/s."
        )
        add_solution_no_table(doc, sol_c4_p3)

    # --- Câu 5 Phần III ---
    p_c5_p3 = doc.add_paragraph()
    p_c5_p3.paragraph_format.space_before = Pt(4)
    p_c5_p3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_c5_p3, "Câu 5. Hình bên là tế bào một loài động vật đang phân bào bình thường. Một nhóm gồm 10 tế bào sinh dục sơ khai loài trên tiến hành nguyên phân 4 lần, các tế bào con tạo ra đều qua vùng chín giảm phân. Tổng số NST môi trường cung cấp cho nhóm tế bào sinh dục sơ khai thực hiện quá trình tạo giao tử là bao nhiêu?")
    add_centered_image(doc, os.path.join(IMG_DIR, "cau5_p3_te_bao.png"), width_cm=6.0)

    if not is_teacher:
        p_ans_box = doc.add_paragraph()
        p_ans_box.paragraph_format.space_before = Pt(1)
        p_ans_box.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans_box, "*Trả lời:* .....................................................................................................", base_size=10.5)
    else:
        sol_c5_p3 = (
            "**Đáp số: 1 860** (hoặc **960** nếu chỉ tính riêng giảm phân)\n"
            "- Tế bào đang ở kì sau nguyên phân, mỗi cực có 6 NST đơn gồm các cặp gen alen đi về 2 cực => Bộ NST lưỡng bội 2n = 6.\n"
            "- Nhóm a = 10 tế bào sinh dục sơ khai nguyên phân k = 4 lần:\n"
            "  + NST môi trường cung cấp cho nguyên phân: NST~mt(NP)~ = a × 2n × (2^k - 1) = 10 × 6 × (16 - 1) = 900 NST.\n"
            "  + Số tế bào con giảm phân: 10 × 2^4 = 160 tế bào.\n"
            "  + NST môi trường cung cấp cho giảm phân: NST~mt(GP)~ = 160 × 2n = 160 × 6 = 960 NST.\n"
            "- Tổng số NST môi trường cung cấp cho nhóm tế bào sinh dục sơ khai tạo giao tử (NP + GP) = 900 + 960 = 1 860 NST."
        )
        add_solution_no_table(doc, sol_c5_p3)

    # --- Câu 6 Phần III ---
    p_c6_p3 = doc.add_paragraph()
    p_c6_p3.paragraph_format.space_before = Pt(4)
    p_c6_p3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_c6_p3, "Câu 6. Một hộ gia đình sử dụng 8 bóng đèn loại 220 V - 40 W. Vào buổi tối, do sụt áp nên hiệu điện thế của mạng điện gia đình chỉ còn 193,6 V. Năng lượng điện năng mà hộ gia đình này tiêu thụ trong mỗi buổi tối hoạt động liên tục 4h bằng bao nhiêu kWh? (kết quả làm tròn đến chữ số hàng phần trăm)")

    if not is_teacher:
        p_ans_box = doc.add_paragraph()
        p_ans_box.paragraph_format.space_before = Pt(1)
        p_ans_box.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans_box, "*Trả lời:* .....................................................................................................", base_size=10.5)
    else:
        sol_c6_p3 = (
            "**Đáp số: 0,99**\n"
            "- Điện trở mỗi bóng đèn: R~đ~ = U~đm~^2 / P~đm~ = 220^2 / 40 = 1 210 Ω.\n"
            "- Khi U = 193,6 V, công suất thực tế của 1 bóng: P~1~ = U^2 / R~đ~ = (193,6)^2 / 1210 = 30,976 W.\n"
            "- Tổng công suất 8 bóng đèn: P~tổng~ = 8 × 30,976 = 247,808 W = 0,247808 kW.\n"
            "- Điện năng tiêu thụ trong 4 giờ: A = P~tổng~ × t = 0,247808 kW × 4 h = 0,991232 kWh ≈ 0,99 kWh."
        )
        add_solution_no_table(doc, sol_c6_p3)

    # ==========================================
    # II. PHẦN NĂNG LƯỢNG VÀ SỰ BIẾN ĐỔI (14,0 ĐIỂM)
    # ==========================================
    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_before = Pt(8)
    p_sec2.paragraph_format.space_after = Pt(3)
    render_rich_text(p_sec2, "*II. PHẦN NĂNG LƯỢNG VÀ SỰ BIẾN ĐỔI (14,0 điểm)*", base_size=12, base_bold=True, base_color=COLOR_PRIMARY)

    # --- Câu 1 Tự luận ---
    p_tl1 = doc.add_paragraph()
    p_tl1.paragraph_format.space_before = Pt(4)
    p_tl1.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_tl1, "Câu 1. (2,5 điểm)")

    p_tl1_body = doc.add_paragraph()
    p_tl1_body.paragraph_format.space_before = Pt(1)
    p_tl1_body.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl1_body, "Trong giải chạy bán Marathon 21 km do thành phố Huế tổ chức, ban kỹ thuật muốn phân tích mô hình chiến thuật tốc độ của hai vận động viên chuyên nghiệp A và B nhằm phục vụ việc xây dựng chiến thuật chạy cho các vận động viên trong các giải chạy sau này.\n"
                                 "- Dữ liệu về vận động viên A: Các thống kê từ đồng hồ GPS cho thấy vận động viên A duy trì chiến thuật giảm tốc theo thời gian:\n"
                                 "  + Trong 1/3 thời gian đầu, A chạy với tốc độ v~A~ (không đổi).\n"
                                 "  + Trong 1/3 thời gian tiếp theo, tốc độ giảm còn 96% so với ban đầu.\n"
                                 "  + Trong 1/3 thời gian cuối, tốc độ chạy lại bé hơn tốc độ trong khoảng thời gian trước đó là 4%.\n"
                                 "- Dữ liệu về vận động viên B: Theo chip gắn trên giày, vận động viên B áp dụng chiến thuật tăng tốc theo quãng đường:\n"
                                 "  + Trên 1/3 quãng đường đầu, B chạy với tốc độ v~B~ (không đổi).\n"
                                 "  + Trên 1/3 quãng đường tiếp theo, tốc độ lớn hơn 4% so với v~B~.\n"
                                 "  + Trên 1/3 quãng đường cuối, tốc độ lớn hơn tốc độ trong 1/3 quãng đường trước đó 4%.\n"
                                 "Hai vận động viên này cùng xuất phát lúc và về đích cùng lúc.\n"
                                 "*1.1.* Tính tỉ số v~A~ / v~B~.\n"
                                 "*1.2.* Thực tế, lúc sắp về đích vận động viên B có dừng lại 07 phút để giúp sơ cứu giúp một vận động viên khác bị thương, nhưng cuối cùng vận động viên B vẫn về đích cùng lúc với với vận động viên A. Biết tốc độ ban đầu của A là 11,75 km/h.\n"
                                 "  a) Tính v~B~.\n"
                                 "  b) Lúc vận động viên B chạy được một nửa chặng đường, tính khoảng cách của hai vận động viên lúc này.", base_size=10.5)

    if not is_teacher:
        add_student_writing_lines(doc, num_lines=8)
    else:
        sol_tl1 = (
            "**1.1. Tính tỉ số v~A~ / v~B~:**\n"
            "- Gọi S = 21 km là chiều dài quãng đường. Giả sử thời gian chạy lý thuyết của hai VĐV là T.\n"
            "- Đối với VĐV A (chia thời gian thành 3 khoảng T/3):\n"
            "  + Giai đoạn 1: tốc độ v~A1~ = v~A~ => s~A1~ = v~A~ × (T/3).\n"
            "  + Giai đoạn 2: tốc độ v~A2~ = 0,96 v~A~ => s~A2~ = 0,96 v~A~ × (T/3).\n"
            "  + Giai đoạn 3: v~A3~ = 0,96 v~A~ × (1 - 0,04) = 0,9216 v~A~ => s~A3~ = 0,9216 v~A~ × (T/3).\n"
            "  => Tổng quãng đường S = (T/3) × v~A~ × (1 + 0,96 + 0,9216) = (2,8816 / 3) × v~A~ × T.\n"
            "  => T = 3S / (2,8816 v~A~).\n"
            "- Đối với VĐV B (chia quãng đường thành 3 đoạn s = S/3 = 7 km):\n"
            "  + Đoạn 1: tốc độ v~B~ => t~B1~ = (S/3) / v~B~.\n"
            "  + Đoạn 2: tốc độ 1,04 v~B~ => t~B2~ = (S/3) / (1,04 v~B~).\n"
            "  + Đoạn 3: tốc độ (1,04)^2 v~B~ = 1,0816 v~B~ => t~B3~ = (S/3) / (1,0816 v~B~).\n"
            "  => Tổng thời gian T~B~ = (S / 3v~B~) × [ 1 + 1/1,04 + 1/1,0816 ] = (S / 3v~B~) × (1951 / 676) ≈ 2,88609 × (S / 3v~B~).\n"
            "- Do hai VĐV về đích cùng lúc nên T~A~ = T~B~:\n"
            "  3S / (2,8816 v~A~) = (1951 S) / (2028 v~B~) => v~A~ / v~B~ = (3 × 2028) / (2,8816 × 1951) ≈ **1,082** (hay 3 802 500 / 3 513 751).\n\n"
            "**1.2. Trường hợp thực tế B dừng 7 phút:**\n"
            "a) Tính v~B~:\n"
            "- Vận tốc ban đầu của A: v~A~ = 11,75 km/h.\n"
            "  Tốc độ trung bình của A: v~tbA~ = (2,8816 / 3) × 11,75 ≈ 11,2863 km/h.\n"
            "  Thời gian A chạy toàn bộ chặng đường: T~A~ = 21 / 11,2863 ≈ 1,86067 giờ (≈ 111,64 phút).\n"
            "- B dừng 7 phút = 7/60 giờ và về đích cùng lúc với A, nên thời gian B thực chạy là:\n"
            "  T~B,chạy~ = T~A~ - 7/60 = 1,86067 - 0,11667 = 1,74400 giờ (≈ 104,64 phút).\n"
            "- Ta có: T~B,chạy~ = (1951 × S) / (2028 × v~B~) => v~B~ = (1951 × 21) / (2028 × 1,74400) ≈ **11,58 km/h**.\n"
            "b) Khoảng cách hai VĐV khi B chạy được nửa chặng đường (10,5 km):\n"
            "- Thời gian B chạy 7 km đầu: t~B1~ = 7 / 11,584 ≈ 0,60428 giờ.\n"
            "- Thời gian B chạy 3,5 km tiếp theo với tốc độ 1,04 v~B~ = 12,047 km/h: t~B2~ = 3,5 / 12,047 ≈ 0,29053 giờ.\n"
            "=> Thời điểm B đến nửa chặng đường: t~0~ = 0,60428 + 0,29053 = 0,89481 giờ (≈ 53,69 phút).\n"
            "- Lúc này A đang chạy trong khoảng thời gian thứ hai (giai đoạn từ 37,21 phút đến 74,43 phút):\n"
            "  Quãng đường A chạy được trong T/3 đầu: s~A1~ = 11,75 × (1,86067 / 3) ≈ 7,2876 km.\n"
            "  Thời gian A chạy ở giai đoạn hai: Δt = 0,89481 - 0,62022 = 0,27459 giờ.\n"
            "  Quãng đường chạy thêm với vận tốc 0,96 v~A~ = 11,28 km/h: s~A2~ = 11,28 × 0,27459 ≈ 3,0974 km.\n"
            "  Tổng quãng đường A đã chạy: s~A~ = 7,2876 + 3,0974 = 10,385 km.\n"
            "- Khoảng cách hai VĐV lúc này: d = 10,500 - 10,385 = 0,115 km = **115 m**."
        )
        add_solution_no_table(doc, sol_tl1)

    # --- Câu 2 Tự luận ---
    p_tl2 = doc.add_paragraph()
    p_tl2.paragraph_format.space_before = Pt(4)
    p_tl2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_tl2, "Câu 2. (2,5 điểm)")

    p_tl2_body = doc.add_paragraph()
    p_tl2_body.paragraph_format.space_before = Pt(1)
    p_tl2_body.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl2_body, "Trong một cốc hình trụ đặt thẳng đứng, hai vật m~1~ và m~2~ có dạng hình hộp chữ nhật và có cùng khối lượng riêng được treo bằng các sợi chỉ vào lực kế gắn trên một giá đỡ cố định. Ban đầu cốc chưa chứa chất lỏng. Sau đó, người ta rót từ từ chất lỏng vào cốc, làm cho mực chất lỏng tăng dần theo độ cao h. Trong quá trình đó, lực kế đo lực căng của sợi chỉ treo vật m~1~. Kết quả thí nghiệm cho thấy lực căng của sợi chỉ T~1~ phụ thuộc vào độ cao mực chất lỏng h theo đồ thị biểu diễn trong hình vẽ bên.\n"
                                 "a) Xác định độ cao mực nước trong cốc khi vật m~2~ vừa ngập trong chất lỏng và vật m~1~ vừa ngập trong chất lỏng.\n"
                                 "b) Xác định lực căng dây T~2~ của sợi chỉ treo m~2~ tại thời điểm chất lỏng vừa ngập vật m~2~.", base_size=10.5)
    add_centered_image(doc, os.path.join(IMG_DIR, "cau2_tl_luc_ke_dothi.png"), width_cm=8.5)

    if not is_teacher:
        add_student_writing_lines(doc, num_lines=8)
    else:
        sol_tl2 = (
            "**a) Xác định độ cao mực chất lỏng khi các vật vừa ngập hoàn toàn:**\n"
            "- Phân tích đồ thị lực căng T~1~ theo độ cao h:\n"
            "  + Từ h = 0 đến h = 1 cm: T~1~ = 12 N (không đổi) => Mực chất lỏng chưa chạm tới đáy vật m~2~ (đáy m~2~ cách đáy cốc 1 cm).\n"
            "  + Từ h = 1 cm đến h = 2 cm: T~1~ giảm từ 12 N xuống 6 N. Tại h = 2 cm, đồ thị chuyển sang đoạn nằm ngang (T~1~ không đổi = 6 N) => Tại **h = 2 cm**, vật m~2~ vừa ngập hoàn toàn trong chất lỏng. Chiều cao vật m~2~ là 2 - 1 = 1 cm.\n"
            "  + Từ h = 2 cm đến h = 3 cm: T~1~ = 6 N (không đổi) => Mực chất lỏng dâng trong khoảng giữa hai vật, m~1~ chưa chạm nước.\n"
            "  + Từ h = 3 cm đến h = 4 cm: T~1~ giảm từ 6 N xuống 3 N. Tại h = 4 cm, đồ thị lại chuyển sang đoạn nằm ngang (T~1~ không đổi = 3 N) => Tại **h = 4 cm**, vật m~1~ vừa ngập hoàn toàn trong chất lỏng. Chiều cao vật m~1~ là 4 - 3 = 1 cm.\n"
            "**Kết luận a:** Vật m~2~ vừa ngập tại **h = 2 cm**; vật m~1~ vừa ngập tại **h = 4 cm**.\n\n"
            "**b) Xác định lực căng dây T~2~ treo m~2~ khi chất lỏng vừa ngập vật m~2~ (h = 2 cm):**\n"
            "- Khi chưa có chất lỏng (h = 0): T~1~ = P~1~ + P~2~ = 12 N.\n"
            "- Lực đẩy Ác-si-mét cực đại tác dụng lên m~2~ (khi ngập hoàn toàn): F~A2~ = 12 - 6 = 6 N.\n"
            "- Lực đẩy Ác-si-mét cực đại tác dụng lên m~1~ (khi ngập hoàn toàn): F~A1~ = 6 - 3 = 3 N.\n"
            "- Vì hai vật có cùng khối lượng riêng và F~A~ tỉ lệ thuận với thể tích V: F~A2~ / F~A1~ = 6 / 3 = 2 => V~2~ = 2 V~1~.\n"
            "- Cùng khối lượng riêng nên trọng lượng P tỉ lệ thuận với V => P~2~ = 2 P~1~.\n"
            "  Mà P~1~ + P~2~ = 12 N => 3 P~1~ = 12 N => P~1~ = 4 N và P~2~ = 8 N.\n"
            "- Xét riêng vật m~2~ tại thời điểm vừa ngập hoàn toàn (h = 2 cm):\n"
            "  Vật m~2~ cân bằng dưới tác dụng của 3 lực thẳng đứng: trọng lực P~2~ = 8 N (hướng xuống), lực đẩy Ác-si-mét F~A2~ = 6 N (hướng lên) và lực căng dây T~2~ (hướng lên).\n"
            "  Điều kiện cân bằng: T~2~ + F~A2~ = P~2~ => T~2~ = P~2~ - F~A2~ = 8 - 6 = **2 N**.\n"
            "*(Kiểm tra lại: Với m~1~, T~1~ = P~1~ + T~2~ = 4 + 2 = 6 N, hoàn toàn khớp đồ thị).* "
        )
        add_solution_no_table(doc, sol_tl2)

    # --- Câu 3 Tự luận ---
    p_tl3 = doc.add_paragraph()
    p_tl3.paragraph_format.space_before = Pt(4)
    p_tl3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_tl3, "Câu 3. (2,5 điểm)")

    p_tl3_body = doc.add_paragraph()
    p_tl3_body.paragraph_format.space_before = Pt(1)
    p_tl3_body.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl3_body, "Làn đường thoát hiểm là khu vực khẩn cấp được đặt bên cạnh đoạn đường xuống dốc. Nó cho phép các phương tiện mất phanh giảm tốc độ và dừng lại mà không ảnh hưởng đến các phương tiện tham gia giao thông khác. Khi xe mất phanh thay vì đi trên đường cao tốc xuống dốc sẽ rẽ vào làn đường thoát hiểm. Làn đường thoát hiểm dài 200 m và nằm bên phải đường cao tốc có độ dốc i = 6%. Chọn mốc thế năng tại mặt đường cao tốc ngay trước làn đường thoát hiểm, biết độ dốc i (%) có công thức: i = (h / d) × 100 (%).\n"
                                 "Trong đó: i là độ dốc của đường (%); h là chênh lệch độ cao giữa điểm đầu và cuối đoạn đường (m); d là chiều dài đoạn đường theo phương nằm ngang (m).\n"
                                 "Một tài xế có khối lượng 50 kg đang lái xe ô tô 04 chỗ của hãng X có khối lượng 1 450 kg đang chuyển động trên đường cao tốc và đang xuống dốc. Bất ngờ, xe gặp sự cố kĩ thuật mất phanh, tài xế cho xe rẽ chạy vào làn đường thoát hiểm trên.\n"
                                 "a) Nêu sự chuyển hóa năng lượng khi xe chuyển động từ mặt đường cao tốc lên làn đường thoát hiểm.\n"
                                 "b) Tính cơ năng của ô tô khi bắt đầu chuyển động vào làn đường thoát hiểm với tốc độ 72 km/h.\n"
                                 "c) Giả sử xe dừng lại tại làn đường thoát hiểm cách đỉnh làn 50 m, hãy tính phần năng lượng chuyển hóa thành do ma sát trên đoạn đường thoát hiểm.", base_size=10.5)

    if not is_teacher:
        add_student_writing_lines(doc, num_lines=8)
    else:
        sol_tl3 = (
            "**a) Sự chuyển hóa năng lượng:**\n"
            "- Khi xe chuyển động lên làn đường thoát hiểm, động năng ban đầu của xe chuyển hóa dần thành:\n"
            "  + Thế năng trọng trường (do độ cao của xe tăng dần khi lên dốc).\n"
            "  + Nhiệt năng làm nóng bánh xe, mặt đường, lớp sỏi cát giảm tốc và không khí do công cản của lực ma sát và lực cản.\n"
            "- Khi xe dừng lại hoàn toàn, toàn bộ động năng ban đầu đã chuyển hóa thành thế năng và nhiệt năng tiêu hao.\n\n"
            "**b) Tính cơ năng khi bắt đầu vào làn thoát hiểm:**\n"
            "- Tổng khối lượng: M = 1 450 + 50 = 1 500 kg.\n"
            "- Tốc độ ban đầu: v~0~ = 72 km/h = 20 m/s.\n"
            "- Chọn mốc thế năng tại mặt đường cao tốc trước làn thoát hiểm: W~t0~ = 0.\n"
            "- Động năng ban đầu: W~đ0~ = (1/2) × M × v~0~^2 = 0,5 × 1 500 × (20)^2 = 300 000 J = 300 kJ.\n"
            "=> Cơ năng khi bắt đầu vào làn: W~0~ = W~đ0~ + W~t0~ = **300 000 J = 300 kJ**.\n\n"
            "**c) Năng lượng chuyển hóa do ma sát khi dừng cách đỉnh làn 50 m:**\n"
            "- Chiều dài làn L = 200 m; xe dừng cách đỉnh 50 m nên quãng đường xe đi trên dốc là: s = 200 - 50 = 150 m.\n"
            "- Độ dốc i = 6% => tan α = 0,06. Vì góc α nhỏ: sin α ≈ tan α = 0,06 (hoặc sin α = 0,06 / √(1 + 0,06^2) ≈ 0,05989).\n"
            "- Độ cao của xe khi dừng: h = s × sin α ≈ 150 × 0,06 = 9,0 m (hoặc 150 × 0,05989 = 8,984 m).\n"
            "- Khi dừng lại, v = 0 => W~đ~ = 0. Cơ năng lúc này là thế năng:\n"
            "  + Lấy g = 10 m/s^2: W~1~ = W~t~ = M × g × h = 1 500 × 10 × 9,0 = 135 000 J = 135 kJ.\n"
            "  + Lấy g = 9,8 m/s^2: W~1~ = 1 500 × 9,8 × 8,984 ≈ 132 065 J ≈ 132,1 kJ.\n"
            "- Phần năng lượng chuyển hóa thành do ma sát:\n"
            "  + Với g = 10 m/s^2: ΔE~ms~ = W~0~ - W~1~ = 300 000 - 135 000 = **165 000 J = 165 kJ**.\n"
            "  + Với g = 9,8 m/s^2: ΔE~ms~ = 300 000 - 132 065 = **167 935 J ≈ 167,9 kJ**."
        )
        add_solution_no_table(doc, sol_tl3)

    # --- Câu 4 Tự luận ---
    p_tl4 = doc.add_paragraph()
    p_tl4.paragraph_format.space_before = Pt(4)
    p_tl4.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_tl4, "Câu 4. (2,5 điểm)")

    p_tl4_body = doc.add_paragraph()
    p_tl4_body.paragraph_format.space_before = Pt(1)
    p_tl4_body.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl4_body, "*4.1.* Cho lăng kính làm bằng chất có chiết suất 1,74 đặt trong không khí, có tiết diện thẳng là tam giác ABC cân tại A, góc chiết quang A = 36°. Chiếu chùm tia sáng hẹp SI đến mặt bên AB của lăng kính theo phương vuông góc với AB như hình vẽ bên. Hãy vẽ đường truyền của tia sáng qua lăng kính và giải thích cách vẽ.\n"
                                 "*4.2.* Theo thứ tự có 3 điểm A, B, C nằm trên quang trục chính xy của một thấu kính, biết AB = 15 cm, AC = 45 cm.\n"
                                 "Khi đặt điểm sáng tại A thì ta thu được ảnh thật của nó tạo bởi thấu kính ở tại C.\n"
                                 "Khi đặt điểm sáng tại B thì ta thu được ảnh ảo của nó tạo bởi thấu kính cũng ở tại C.\n"
                                 "a) Đây là thấu kính gì? Hãy xác định vị trí đặt thấu kính, giải thích.\n"
                                 "b) Tính khoảng cách từ thấu kính đến điểm A và tiêu cự của thấu kính.\n"
                                 "Cho công thức thấu kính: 1/f = 1/d + 1/d'; trong đó f là tiêu cự của thấu kính; d và d' lần lượt là khoảng cách từ thấu kính đến vật và ảnh của vật qua thấu kính, d' > 0 là ảnh thật, d' < 0 là ảnh ảo.", base_size=10.5)
    add_centered_image(doc, os.path.join(IMG_DIR, "cau41_tl_lang_kinh.png"), width_cm=4.5)

    if not is_teacher:
        add_student_writing_lines(doc, num_lines=8)
    else:
        sol_tl4 = (
            "**4.1. Đường truyền tia sáng qua lăng kính:**\n"
            "- Góc giới hạn phản xạ toàn phần: sin i~gh~ = 1 / n = 1 / 1,74 ≈ 0,5747 => i~gh~ ≈ 35,08°.\n"
            "- Tam giác ABC cân tại A, góc A = 36° => góc đáy B = C = (180° - 36°) / 2 = 72°.\n"
            "- Xét đường truyền của tia sáng:\n"
            "  + Tại mặt AB: Tia tới SI vuông góc với AB => góc tới i~1~ = 0°, tia sáng truyền thẳng không bị khúc xạ, tới gặp mặt AC tại điểm J.\n"
            "  + Tại mặt AC (điểm J): Trong tam giác AIJ vuông tại I có góc A = 36° => góc AJI = 54°. Góc tới tại J là i~2~ = 90° - 54° = 36°. Vì i~2~ = 36° > i~gh~ (35,08°) nên tại J xảy ra phản xạ toàn phần với góc phản xạ i~2~' = 36°, tạo tia phản xạ JK hướng xuống đáy BC.\n"
            "  + Tại đáy BC (điểm K): Trong tam giác CJK có góc C = 72°, góc CJK = 180° - (54° + 72°) = 54° => góc CKJ = 54°. Góc tới tại K là i~3~ = 90° - 54° = 36° > i~gh~ nên tại K lại tiếp tục phản xạ toàn phần, tạo tia phản xạ KL hướng về mặt AB.\n"
            "  + Tia sáng tiếp tục phản xạ tuần hoàn khép kín theo hình ngũ giác sao đối xứng trong lăng kính.\n\n"
            "**4.2. Bài toán thấu kính:**\n"
            "a) Xác định loại thấu kính và vị trí:\n"
            "- Thấu kính phân kì luôn cho ảnh ảo đối với vật thật. Ở đây, khi đặt vật tại A thu được ảnh THẬT tại C => **Đây là thấu kính hội tụ**.\n"
            "- Đối với thấu kính hội tụ:\n"
            "  + Vật thật A cho ảnh thật C => A và C nằm ở hai phía so với quang tâm O (O nằm giữa A và C).\n"
            "  + Vật thật B cho ảnh ảo C => B và C nằm cùng một phía so với quang tâm O (O không nằm giữa B và C).\n"
            "  Thứ tự các điểm trên trục chính là A, B, C (AB = 15 cm, AC = 45 cm). Kết hợp hai điều kiện trên suy ra **quang tâm O phải nằm giữa A và B** (thứ tự: A --- O --- B --- C).\n"
            "b) Tính khoảng cách từ thấu kính đến A và tiêu cự f:\n"
            "- Đặt khoảng cách OA = d~A~ (0 < d~A~ < 15 cm) => OB = d~B~ = 15 - d~A~; OC = 45 - d~A~.\n"
            "- Vật tại A cho ảnh thật C: d = d~A~; d' = 45 - d~A~ => 1/f = 1/d~A~ + 1/(45 - d~A~)  (1)\n"
            "- Vật tại B cho ảnh ảo C: d = 15 - d~A~; d' = -(45 - d~A~) => 1/f = 1/(15 - d~A~) - 1/(45 - d~A~)  (2)\n"
            "- Từ (1) và (2) suy ra: 1/d~A~ + 2/(45 - d~A~) = 1/(15 - d~A~)\n"
            "  <=> (45 - d~A~)(15 - d~A~) + 2d~A~(15 - d~A~) = d~A~(45 - d~A~)\n"
            "  <=> 675 - 30d~A~ - d~A~^2 = 45d~A~ - d~A~^2 <=> 75d~A~ = 675 => **d~A~ = 9 cm**.\n"
            "- Thay d~A~ = 9 cm vào (1): 1/f = 1/9 + 1/36 = 5/36 => **f = 36/5 = 7,2 cm**."
        )
        add_solution_no_table(doc, sol_tl4)

    # --- Câu 5 Tự luận ---
    p_tl5 = doc.add_paragraph()
    p_tl5.paragraph_format.space_before = Pt(4)
    p_tl5.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_tl5, "Câu 5. (3,0 điểm)")

    p_tl5_body = doc.add_paragraph()
    p_tl5_body.paragraph_format.space_before = Pt(1)
    p_tl5_body.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl5_body, "Cho mạch điện như hình vẽ bên. Điện trở R = 50 Ω, ampe kế được coi là lí tưởng nên có điện trở không đáng kể. Nguồn điện có hiệu điện thế U~AB~ có thể thay đổi được. Các chốt (1) và (2) dùng để cắm dây nối vào mạch. Bỏ qua điện trở của các dây dẫn.\n"
                                 "a) Điều chỉnh để U~AB~ = U~1~ = 220 V: nối một bóng đèn dây tóc có ghi 200 V – 80 W vào giữa hai chốt (1) và (2). Tính số chỉ ampe kế; nhận xét độ sáng của bóng đèn lúc này.\n"
                                 "b) Thay bóng đèn nối vào giữa hai chốt (1) và (2) ở câu 5a) bằng một hộp X chứa n bóng đèn giống hệt nhau mắc song song. Biết trên mỗi bóng đèn dây tóc có ghi 220 V– 20 W. Điều chỉnh hiệu điện thế của nguồn điện đến khi U~AB~ = U~2~ thì các đèn trong hộp X đều sáng bình thường, khi đó công suất tiêu thụ của toàn mạch là 460 W. Tính giá trị của U~2~ và n.", base_size=10.5)
    add_centered_image(doc, os.path.join(IMG_DIR, "cau5_tl_mach_dien.png"), width_cm=8.5)

    if not is_teacher:
        add_student_writing_lines(doc, num_lines=8)
    else:
        sol_tl5 = (
            "**a) Tính số chỉ ampe kế và nhận xét độ sáng của đèn:**\n"
            "- Bóng đèn Đ (200 V - 80 W) có:\n"
            "  + Điện trở: R~Đ~ = U~đm~^2 / P~đm~ = 200^2 / 80 = 500 Ω.\n"
            "  + Cường độ định mức: I~đm~ = P~đm~ / U~đm~ = 80 / 200 = 0,4 A.\n"
            "- Đèn Đ mắc nối tiếp với R = 50 Ω và ampe kế: R~tđ~ = R + R~Đ~ = 50 + 500 = 550 Ω.\n"
            "- Cường độ dòng điện trong mạch (số chỉ ampe kế): I~A~ = I = U~1~ / R~tđ~ = 220 / 550 = **0,4 A**.\n"
            "- Độ sáng: Ta thấy dòng điện thực tế qua đèn I = I~đm~ = 0,4 A (hay U~Đ~ = I × R~Đ~ = 200 V = U~đm~), do đó bóng đèn **sáng bình thường**.\n\n"
            "**b) Xác định n và U~2~:**\n"
            "- Hộp X gồm n bóng đèn (220 V - 20 W) mắc song song. Khi các đèn sáng bình thường:\n"
            "  + Hiệu điện thế hai đầu hộp X: U~X~ = 220 V.\n"
            "  + Dòng điện định mức mỗi đèn: I~0~ = 20 / 220 = 1/11 A.\n"
            "  + Cường độ dòng điện qua hộp X (dòng mạch chính): I~mạch~ = n × (1/11) = n / 11 (A).\n"
            "  + Công suất tiêu thụ của hộp X: P~X~ = n × 20 = 20n (W).\n"
            "- Công suất tỏa nhiệt trên điện trở R: P~R~ = I~mạch~^2 × R = (n / 11)^2 × 50 = (50 / 121) × n^2 (W).\n"
            "- Công suất toàn mạch: P~toàn mạch~ = P~R~ + P~X~ = (50 / 121) × n^2 + 20n = 460 W\n"
            "  <=> 50n^2 + 2 420n - 55 660 = 0 <=> 5n^2 + 242n - 5 566 = 0.\n"
            "  Giải phương trình bậc hai: Δ' = (121)^2 - 5 × (-5 566) = 42 471 => √Δ' ≈ 206,085.\n"
            "  Nghiệm dương: n = (-121 + 206,085) / 5 ≈ 17,017.\n"
            "  Do n là số nguyên dương (số bóng đèn) nên **n = 17** bóng.\n"
            "- Tính hiệu điện thế nguồn U~2~:\n"
            "  Dòng điện mạch chính: I = 17 / 11 A.\n"
            "  Độ giảm thế trên R: U~R~ = I × R = (17 / 11) × 50 = 850 / 11 ≈ 77,27 V.\n"
            "  Hiệu điện thế toàn mạch: U~2~ = U~R~ + U~X~ = 850 / 11 + 220 = 3 270 / 11 ≈ **297,27 V** (hay 297,3 V)."
        )
        add_solution_no_table(doc, sol_tl5)

    # --- Câu 6 Tự luận ---
    p_tl6 = doc.add_paragraph()
    p_tl6.paragraph_format.space_before = Pt(4)
    p_tl6.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_tl6, "Câu 6. (1,0 điểm)")

    p_tl6_body = doc.add_paragraph()
    p_tl6_body.paragraph_format.space_before = Pt(1)
    p_tl6_body.paragraph_format.space_after = Pt(2)
    render_rich_text(p_tl6_body, "Cho các dụng cụ: 01 thước dài tiết diện đều có vạch chia nhỏ nhất đến milimet, 01 giá đỡ, dây treo nhẹ, 01 cốc nước đã biết khối lượng riêng D~n~, 01 cốc có chất lỏng cần xác định khối lượng riêng D~x~, 02 vật rắn khối lượng m~1~, m~2~ khác nhau có thể chìm trong nước và chất lỏng nói trên.\n"
                                 "Hãy trình bày phương án xác định khối lượng riêng của một chất lỏng trên.", base_size=10.5)

    if not is_teacher:
        add_student_writing_lines(doc, num_lines=8)
    else:
        sol_tl6 = (
            "**1. Cơ sở lý thuyết:**\n"
            "- Sử dụng thước dài đồng chất, tiết diện đều như một đòn bẩy quay quanh trục nằm ngang qua điểm tựa O (với O là trọng tâm của thước, trọng lượng của thước không gây ra mômen làm quay).\n"
            "- Theo quy tắc mômen lực: Khi cân bằng trong không khí: P~1~ × l~1~ = P~2~ × l~2~.\n"
            "- Khi nhúng vật m~2~ (thể tích V~2~) vào chất lỏng có khối lượng riêng D, vật chịu lực đẩy Ác-si-mét F~A~ = 10 × D × V~2~ hướng lên. Trọng lượng biểu kiến giảm còn P~2~' = P~2~ - F~A~.\n"
            "- Lập tỉ số biến thiên mômen khi nhúng trong chất lỏng cần đo D~x~ và trong nước D~n~ sẽ triệt tiêu được các đại lượng chưa biết m~1~, m~2~, V~2~ và xác định được D~x~.\n\n"
            "**2. Các bước tiến hành thí nghiệm:**\n"
            "- *Bước 1 (Xác định trọng tâm O):* Treo thước lên giá đỡ bằng sợi dây nhẹ, dịch chuyển điểm treo cho đến khi thước thăng bằng nằm ngang. Đánh dấu vị trí O (trọng tâm của thước) và giữ cố định điểm tựa O này.\n"
            "- *Bước 2 (Cân bằng trong không khí):* Treo vật m~1~ vào bên trái cách O một khoảng OA = l~1~; treo vật m~2~ vào bên phải cách O một khoảng OB = l~2~ sao cho thước cân bằng nằm ngang. Ghi nhận l~1~, l~2~ trên thước.\n"
            "  Ta có: P~1~ × l~1~ = P~2~ × l~2~   (1)\n"
            "- *Bước 3 (Nhúng m~2~ chìm trong nước):* Giữ nguyên vị trí treo m~2~ (khoảng cách l~2~). Đặt cốc nước có khối lượng riêng D~n~ cho vật m~2~ chìm hoàn toàn trong nước (không chạm đáy và thành cốc). Dịch chuyển điểm treo m~1~ đến vị trí A~1~ cách O khoảng l~1~' để thước cân bằng nằm ngang trở lại. Ghi nhận l~1~'.\n"
            "  Ta có: P~1~ × l~1~' = (P~2~ - 10 × D~n~ × V~2~) × l~2~   (2)\n"
            "  Lấy (1) trừ (2): P~1~ × (l~1~ - l~1~') = 10 × D~n~ × V~2~ × l~2~   (3)\n"
            "- *Bước 4 (Nhúng m~2~ chìm trong chất lỏng cần đo):* Nhấc m~2~ ra khỏi cốc nước, dùng khăn lau thật khô. Nhúng chìm hoàn toàn m~2~ vào cốc chứa chất lỏng D~x~. Dịch chuyển điểm treo m~1~ đến vị trí A~2~ cách O khoảng l~1~'' để thước lại cân bằng nằm ngang. Ghi nhận l~1~''.\n"
            "  Ta có: P~1~ × l~1~'' = (P~2~ - 10 × D~x~ × V~2~) × l~2~   (4)\n"
            "  Lấy (1) trừ (4): P~1~ × (l~1~ - l~1~'') = 10 × D~x~ × V~2~ × l~2~   (5)\n\n"
            "**3. Công thức tính khối lượng riêng D~x~:**\n"
            "- Lấy (5) chia cho (3) vế theo vế, ta được:\n"
            "  (l~1~ - l~1~'') / (l~1~ - l~1~') = D~x~ / D~n~\n"
            "  => **D~x~ = D~n~ × (l~1~ - l~1~'') / (l~1~ - l~1~')**.\n\n"
            "**4. Biện pháp giảm sai số:**\n"
            "- Thao tác lặp lại 3 - 5 lần với các vị trí ban đầu l~1~, l~2~ khác nhau và lấy giá trị trung bình.\n"
            "- Đảm bảo vật ngập hoàn toàn trong chất lỏng, không chạm thành và đáy cốc; dây treo thẳng đứng vuông góc với thước.\n"
            "- Luôn lau khô vật m~2~ trước khi nhúng sang cốc chất lỏng cần đo."
        )
        add_solution_no_table(doc, sol_tl6)

    # --- Footer block ---
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "---------- *HẾT* ----------", base_size=10.5, base_bold=True)

    p_rule1 = doc.add_paragraph()
    p_rule1.paragraph_format.space_before = Pt(2)
    p_rule1.paragraph_format.space_after = Pt(1)
    p_rule1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_rule1, "_Thí sinh không được sử dụng tài liệu; Giám thị không giải thích gì thêm._", base_size=9.5, base_italic=True)

    # Save
    filename = "De_Thi_HSG_KHTN9_TPHue_2025_2026_GiaoVien.docx" if is_teacher else "De_Thi_HSG_KHTN9_TPHue_2025_2026_HocSinh.docx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    doc.save(filepath)
    print(f"File created successfully: {filepath}")
    return filepath

if __name__ == '__main__':
    generate_exam_docx(is_teacher=False)
    generate_exam_docx(is_teacher=True)
