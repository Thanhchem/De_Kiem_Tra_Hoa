# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file Word cho ĐỀ SỐ 1 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De01.docx (Bản Học Sinh)
- De_Kiem_Tra_Hoa_11_De01_GiaoVien.docx (Bản Giáo Viên)
Tuân thủ 100% Form Đề Kiểm Tra trong GEMINI.md & form_de_kiem_tra.py:
- Tuyệt đối không dùng Table trong nội dung câu hỏi/đáp án/lời giải.
- Header & Footer chuẩn: #D32F2F, viền #8BC390, số trang tự động PAGE.
- Căn lề A4 1.5 cm.
- Chỉ số trên/dưới hóa học triệt để.
"""

import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Thêm đường dẫn scripts để import module dùng chung
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
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
    COLOR_CORRECT,
    COLOR_SOLUTION,
    COLOR_PRIMARY
)

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de01")

def add_centered_picture(doc, img_path, width_cm=8.0):
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run()
        run.add_picture(img_path, width=Cm(width_cm))

def build_exam_document(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ",
        exam_title="ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 1 (CÂN BẰNG HÓA HỌC)",
        subject_line="MÔN: HOÁ HỌC - LỚP 11",
        num_pages="04",
        duration="45 phút",
        exam_code="101",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Ba = 137",
        is_teacher=is_teacher
    )

    # ==========================================
    # PHẦN I
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. Trong cùng điều kiện xác định, kết luận nào sau phản ứng xảy ra từ chất tham gia tạo thành chất sản phẩm. Kết luận nào sau đây là đúng?")
    opts_c1 = [
        "A. Nếu chất sản phẩm không thể tác dụng với nhau để tạo lại chất ban đầu thì đó là phản ứng thuận nghịch.",
        "B. Nếu chất sản phẩm có thể tác dụng với nhau để tạo lại chất ban đầu thì đó là phản ứng một chiều.",
        "C. Nếu chất sản phẩm có thể tác dụng với nhau để tạo lại chất khác thì đó là phản ứng thuận nghịch.",
        "*D. Nếu chất sản phẩm có thể tác dụng với nhau để tạo lại chất ban đầu thì đó là phản ứng thuận nghịch.*"
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phản ứng thuận nghịch là phản ứng trong cùng điều kiện xác định, chất phản ứng tạo thành chất sản phẩm và các chất sản phẩm cũng có thể tác dụng với nhau để tạo lại chất phản ứng ban đầu.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Trong phản ứng thuận nghịch, kết luận nào sau đây là đúng tại thời điểm ban đầu?")
    opts_c2 = [
        "A. Tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch.",
        "B. Nồng độ các chất trong hệ phản ứng không đổi.",
        "C. Tốc độ phản ứng nghịch bằng 0 sau đó giảm dần.",
        "*D. Tốc độ phản ứng thuận đạt lớn nhất sau đó giảm dần.*"
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Tại thời điểm ban đầu, nồng độ các chất tham gia lớn nhất nên tốc độ phản ứng thuận lớn nhất, sau đó giảm dần. Nồng độ chất sản phẩm bằng 0 nên tốc độ phản ứng nghịch bằng 0, sau đó tăng dần theo thời gian.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3 (SGK - CD). Phát biểu nào sau đây về một phản ứng thuận nghịch tại trạng thái cân bằng là sai?")
    opts_c3 = [
        "A. Tốc độ của phản ứng thuận bằng tốc độ của phản ứng nghịch.",
        "B. Nồng độ của tất cả các chất trong hỗn hợp phản ứng là không đổi.",
        "*C. Nồng độ mol của chất phản ứng luôn bằng nồng độ mol của chất sản phẩm phản ứng.*",
        "D. Phản ứng thuận và phản ứng nghịch vẫn diễn ra."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Tại trạng thái cân bằng, tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch và nồng độ của các chất không thay đổi theo thời gian, nhưng không nhất thiết nồng độ chất phản ứng phải bằng nồng độ chất sản phẩm.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Phản ứng nào sau đây là phản ứng thuận nghịch?")
    opts_c4 = [
        "A. Mg + 2HCl ➔ MgCl~2~ + H~2~.",
        "*B. H~2~ + I~2~ ⇌ 2HI.*",
        "C. C~2~H~5~OH + 3O~2~ ➔ 2CO~2~ + 3H~2~O.",
        "D. 2KClO~3~ ➔ 2KCl + 3O~2~."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phản ứng H~2~(g) + I~2~(g) ⇌ 2HI(g) là phản ứng thuận nghịch điển hình. Các phản ứng còn lại là phản ứng một chiều.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Trong số các dung dịch có cùng nồng độ mol sau đây, dung dịch nào có độ dẫn điện yếu nhất?")
    opts_c5 = [
        "A. NaCl",
        "B. CH~3~COONa",
        "*C. CH~3~COOH*",
        "D. H~2~SO~4~"
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "NaCl, CH~3~COONa là các chất điện li mạnh (muối tan hoàn toàn thành ion). H~2~SO~4~ là acid mạnh điện li mạnh. CH~3~COOH là acid yếu, chất điện li yếu (chỉ phân li một phần nhỏ thành ion) nên nồng độ ion tự do nhỏ nhất ➔ dẫn điện yếu nhất.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Cho các chất dưới đây: HClO~4~, HClO, HF, HNO~3~, H~2~S, H~2~SO~3~, NaOH, NaCl, CuSO~4~, CH~3~COOH. Số chất thuộc loại chất điện li mạnh là")
    opts_c6 = [
        "*A. 5.*",
        "B. 6.",
        "C. 7.",
        "D. 4."
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Các chất điện li mạnh gồm: acid mạnh (HClO~4~, HNO~3~), base mạnh (NaOH), muối tan (NaCl, CuSO~4~) ➔ Tổng cộng 5 chất. Các chất còn lại (HClO, HF, H~2~S, H~2~SO~3~, CH~3~COOH) là chất điện li yếu.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 7. Trong dung dịch loãng của các chất khác nhau, tích số ion của nước chỉ phụ thuộc vào yếu tố nào sau đây?")
    opts_c7 = [
        "A. nồng độ",
        "B. thể tích",
        "*C. nhiệt độ*",
        "D. áp suất"
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Tích số ion của nước K~w~ = [H^+].[OH^-] là hằng số cân bằng, do đó chỉ phụ thuộc vào nhiệt độ. Ở 25°C, K~w~ = 1,0.10^-14^.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 8. Dung dịch nào dưới đây dẫn điện tốt nhất?")
    opts_c8 = [
        "A. NaI 0,002M",
        "B. NaI 0,010M",
        "*C. NaI 0,100M*",
        "D. NaI 0,001M"
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "NaI là chất điện li mạnh. Cùng một chất điện li mạnh thì dung dịch có nồng độ mol càng lớn sẽ có mật độ ion tự do dẫn điện càng cao ➔ NaI 0,100M dẫn điện tốt nhất.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 9 [CD - SBT]. Đối với phản ứng sau, cân bằng sẽ bị ảnh hưởng như thế nào khi tăng nhiệt độ (các điều kiện khác giữ không đổi)?\n H~2~(g) + 1/2 O~2~(g) ⇌ H~2~O(l)    Δ~r~H°~298~ = -296 kJ")
    opts_c9 = [
        "A. Cân bằng chuyển dịch sang phải.",
        "*B. Cân bằng chuyển dịch sang trái.*",
        "C. Không thay đổi.",
        "D. Không dự đoán được sự chuyển dịch cân bằng."
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Δ~r~H°~298~ = -296 kJ < 0 ➔ phản ứng thuận tỏa nhiệt. Theo nguyên lí Le Chatelier, khi tăng nhiệt độ cân bằng chuyển dịch theo chiều phản ứng thu nhiệt (chiều nghịch, tức chuyển dịch sang trái).")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 10 [CD - SBT]. Trong phản ứng nào sau đây sự tăng áp suất sẽ dẫn tới cân bằng chuyển dịch sang trái (các điều kiện khác coi như không thay đổi)?")
    opts_c10 = [
        "*A. CaCO~3~(s) ⇌ CaO(s) + CO~2~(g)*",
        "B. CO(g) + H~2~O(g) ⇌ H~2~(g) + CO~2~(g)",
        "C. 2H~2~(g) + O~2~(g) ⇌ 2H~2~O(l)",
        "D. C(s) + O~2~(g) ⇌ CO~2~(g)"
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khi tăng áp suất, cân bằng chuyển dịch theo chiều giảm số mol khí.\n- Phản ứng A: vế trái có 0 mol khí, vế phải có 1 mol khí CO~2~. Chiều nghịch (sang trái) là chiều giảm số mol khí (từ 1 xuống 0) ➔ cân bằng chuyển dịch sang trái.\n- B và D có số mol khí 2 vế bằng nhau (áp suất không ảnh hưởng).\n- C có 3 mol khí vế trái và 0 mol khí vế phải (tăng áp suất chuyển dịch sang phải).")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 11. Cho cân bằng hóa học: CaCO~3~(rắn) ⇌ CaO(rắn) + CO~2~(khí). Biết phản ứng thuận là phản ứng thu nhiệt. Tác động nào sau đây vào hệ cân bằng để cân bằng đã cho chuyển dịch theo chiều thuận?")
    opts_c11 = [
        "A. Giảm nhiệt độ.",
        "B. Tăng áp suất.",
        "C. Tăng nồng độ khí CO~2~.",
        "*D. Tăng nhiệt độ.*"
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phản ứng thuận là phản ứng thu nhiệt (Δ~r~H > 0). Theo nguyên lí Le Chatelier, khi tăng nhiệt độ cân bằng sẽ chuyển dịch theo chiều thu nhiệt, tức chiều thuận.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 12. Cho cân bằng hóa học: PCl~5~(k) ⇌ PCl~3~(k) + Cl~2~(k); Δ~r~H°~298~ > 0. Cân bằng chuyển dịch theo chiều thuận khi")
    opts_c12 = [
        "A. Thêm PCl~3~ vào hệ phản ứng.",
        "B. Tăng áp suất của hệ phản ứng.",
        "*C. Tăng nhiệt độ của hệ phản ứng.*",
        "D. Thêm Cl~2~ vào hệ phản ứng."
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Δ~r~H°~298~ > 0 ➔ phản ứng thuận thu nhiệt. Khi tăng nhiệt độ, cân bằng chuyển dịch theo chiều thu nhiệt, tức chiều thuận.")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 13. Cho các cân bằng sau:\n(1) 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g)\n(2) N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)\n(3) CO~2~(g) + H~2~(g) ⇌ CO(g) + H~2~O(g)\n(4) 2HI(g) ⇌ H~2~(g) + I~2~(g)\nKhi thay đổi áp suất, nhóm gồm các cân bằng hoá học đều không bị chuyển dịch là")
    opts_c13 = [
        "A. (1) và (2).",
        "B. (1) và (3).",
        "*C. (3) và (4).*",
        "D. (2) và (4)."
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Áp suất không ảnh hưởng đến cân bằng của phản ứng có tổng số mol khí ở hai vế bằng nhau.\n- Cân bằng (3): 1 + 1 = 1 + 1 = 2 mol khí.\n- Cân bằng (4): 2 = 1 + 1 = 2 mol khí.\nVậy nhóm (3) và (4) không bị chuyển dịch khi thay đổi áp suất.")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 14. Cho cân bằng hoá học: N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g); phản ứng thuận là phản ứng tỏa nhiệt. Cân bằng hoá học không bị chuyển dịch khi")
    opts_c14 = [
        "A. Thay đổi nồng độ N~2~",
        "*B. Thêm chất xúc tác Fe.*",
        "C. Thay đổi nhiệt độ.",
        "D. Thay đổi áp suất của hệ."
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Chất xúc tác làm tăng tốc độ phản ứng thuận và tốc độ phản ứng nghịch với số lần bằng nhau, giúp hệ nhanh đạt tới trạng thái cân bằng chứ không làm chuyển dịch cân bằng hoá học.")

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 15 (SGK-KNTT-Tr21). pH của dung dịch nào sau đây có giá trị nhỏ nhất?")
    opts_c15 = [
        "*A. Dung dịch HCl 0,1 M.*",
        "B. Dung dịch CH~3~COOH 0,1 M.",
        "C. Dung dịch NaCl 0,1 M.",
        "D. Dung dịch NaOH 0,01 M."
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Dung dịch acid có pH < 7, acid càng mạnh thì pH càng nhỏ.\n- HCl là acid mạnh phân li hoàn toàn: [H^+] = 0,1 M ➔ pH = 1.\n- CH~3~COOH là acid yếu phân li một phần: [H^+] < 0,1 M ➔ pH > 1.\n- NaCl trung tính có pH = 7; NaOH có pH = 12.\nVậy HCl 0,1 M có pH nhỏ nhất.")

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 16 (SGK-KNTT-Tr22). Đo pH của một cốc nước chanh được giá trị pH bằng 2,4. Nhận định nào sau đây không đúng?")
    opts_c16 = [
        "A. Nước chanh có môi trường acid.",
        "B. Nồng độ ion [H^+] của nước chanh là 10^-2,4^ mol/L.",
        "*C. Nồng độ ion [H^+] của nước chanh là 0,24 mol/L.*",
        "D. Nồng độ của ion [OH^-] của nước chanh nhỏ hơn 10^-7^ mol/L."
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "- A đúng vì pH = 2,4 < 7 ➔ môi trường acid.\n- B đúng vì [H^+] = 10^-pH^ = 10^-2,4^ mol/L.\n- C sai vì [H^+] = 10^-2,4^ mol/L ≈ 3,98.10^-3^ mol/L ≠ 0,24 mol/L.\n- D đúng vì [OH^-] = 10^-14^ / 10^-2,4^ = 10^-11,6^ < 10^-7^ mol/L.")

    # Câu 17
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 17. Cho các phát biểu sau:\n(1) Khi phản ứng thuận nghịch đạt trạng thái cân bằng hóa học, phản ứng dừng lại.\n(2) Cân bằng hóa học là cân bằng động.\n(3) Khi thay đổi trạng thái cân bằng của phản ứng thuận nghịch, cân bằng sẽ chuyển dịch về phía chống lại sự thay đổi đó.\n(4) Các yếu tố luôn ảnh hưởng đến cân bằng hoá học là: Nhiệt độ, nồng độ, áp suất.\n(5) Phản ứng thuận nghịch là phản ứng xảy ra theo 2 chiều ngược nhau trong cùng một điều kiện.\n(6) Phản ứng bất thuận nghịch là phản ứng xảy ra theo 1 chiều xác định.\n(7) Cân bằng hóa học là trạng thái mà phản ứng đã xảy ra hoàn toàn.\n(8) Khi phản ứng thuận nghịch đạt trạng thái cân bằng hóa học, lượng các chất sẽ không đổi.\nSố phát biểu đúng là")
    opts_c17 = [
        "A. 7",
        "B. 8",
        "C. 6",
        "*D. 5*"
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Có 5 phát biểu đúng gồm: (2), (3), (5), (6), (8).\n- (1) Sai vì tại trạng thái cân bằng phản ứng thuận và nghịch vẫn diễn ra.\n- (4) Sai vì áp suất chỉ ảnh hưởng khi hệ có chất khí và có biến thiên số mol khí.\n- (7) Sai vì phản ứng thuận nghịch không xảy ra hoàn toàn.")

    # Câu 18
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 18 [CD - SBT].\na) Cốc A chứa 50 ml dung dịch KOH 0,10M được chuẩn độ với dung dịch HNO~3~ 0,10M. Sau khi thêm 52 ml dung dịch HNO~3~ vào, pH của dung dịch trong cốc A là")
    opts_c18a = [
        "A. 2,80.",
        "*B. 2,71.*",
        "C. 2,40.",
        "D. 3,00."
    ]
    render_options_no_table(doc, opts_c18a, is_teacher=is_teacher)
    
    p_18b = doc.add_paragraph()
    p_18b.paragraph_format.space_before = Pt(3)
    p_18b.paragraph_format.space_after = Pt(2)
    render_rich_text(p_18b, "*b) Chuẩn độ 100,0 ml dung dịch NaOH 0,1M bằng dung dịch HCl 1,0 M. Thể tích dung dịch HCl cần thêm để dung dịch thu được có pH = 12 là*", base_size=10.5)
    opts_c18b = [
        "*A. 8,91 ml.*",
        "B. 8,52 ml.",
        "C. 9,01 ml.",
        "D. 8,72 ml."
    ]
    render_options_no_table(doc, opts_c18b, is_teacher=is_teacher)
    if is_teacher:
        sol_c18 = (
            "a) n(KOH) = 0,050 . 0,10 = 0,0050 mol; n(HNO~3~) = 0,052 . 0,10 = 0,0052 mol.\n"
            "   H^+ + OH^- ➔ H~2~O  ➔  n(H^+ dư) = 0,0052 - 0,0050 = 0,0002 mol.\n"
            "   V~dd~ = 50 + 52 = 102 mL = 0,102 L.\n"
            "   [H^+] = 0,0002 / 0,102 ≈ 1,961.10^-3^ M ➔ pH = -log(1,961.10^-3^) ≈ 2,71.\n"
            "b) n(NaOH) = 0,100 . 0,1 = 0,01 mol. Gọi thể tích HCl 1,0 M là V (L) ➔ n(HCl) = 1,0.V (mol).\n"
            "   pH = 12 ➔ pOH = 2 ➔ [OH^-] = 0,01 M.\n"
            "   [OH^-] = (0,01 - V) / (0,1 + V) = 0,01  ➔  0,01 - V = 0,001 + 0,01V  ➔  1,01V = 0,009  ➔  V ≈ 0,00891 L = 8,91 mL."
        )
        add_solution_no_table(doc, sol_c18)

    # ==========================================
    # PHẦN II
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. Trong bình kín có hệ cân bằng hóa học sau: CO(g) + H~2~O(g) ⇌ CO~2~(g) + H~2~(g); Δ~r~H°~298~ > 0")
    stmts_p2_c1 = [
        ("a) Thêm một lượng hơi nước cân bằng chuyển dịch theo chiều thuận.", "ĐÚNG"),
        ("b) Thêm một lượng CO~2~ cân bằng chuyển dịch theo chiều thuận.", "SAI"),
        ("c) Giảm áp suất chung của hệ cân bằng chuyển dịch theo chiều thuận.", "SAI"),
        ("d) Tăng nhiệt độ cân bằng chuyển dịch theo chiều thuận.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c1 = (
            "a) Đúng vì thêm hơi nước (chất tham gia) làm cân bằng chuyển dịch theo chiều giảm hơi nước (chiều thuận).\n"
            "b) Sai vì thêm CO~2~ (sản phẩm) làm cân bằng chuyển dịch theo chiều nghịch.\n"
            "c) Sai vì tổng số mol khí 2 vế bằng nhau (1+1 = 1+1 = 2) nên thay đổi áp suất không làm chuyển dịch cân bằng.\n"
            "d) Đúng vì Δ~r~H° > 0 (thu nhiệt), khi tăng nhiệt độ cân bằng chuyển dịch theo chiều thu nhiệt (chiều thuận)."
        )
        add_solution_no_table(doc, sol_p2_c1)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Xét cân bằng trong bình kín có dung tích không đổi: X(g) ⇌ 2Y(g). Ban đầu cho 1 mol khí X vào bình, khi đạt đến trạng thái cân bằng thì thấy: Tại thời điểm ở 35°C trong bình có 0,730 mol X; Tại thời điểm ở 45°C trong bình có 0,623 mol X.")
    stmts_p2_c2 = [
        ("a) Phản ứng thuận là phản ứng thu nhiệt.", "ĐÚNG"),
        ("b) Khi tăng áp suất, cân bằng chuyển dịch theo chiều nghịch.", "ĐÚNG"),
        ("c) Thêm tiếp Y vào hỗn hợp cân bằng thì làm cho cân bằng chuyển dịch theo chiều nghịch.", "ĐÚNG"),
        ("d) Thêm xúc tác thích hợp vào hỗn hợp cân bằng thì cân bằng vẫn không chuyển dịch.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c2 = (
            "a) Đúng vì khi tăng nhiệt độ từ 35°C lên 45°C, lượng X giảm từ 0,730 mol xuống 0,623 mol chứng tỏ cân bằng chuyển dịch theo chiều thuận ➔ chiều thuận là chiều thu nhiệt.\n"
            "b) Đúng vì phản ứng thuận làm tăng số mol khí (từ 1 mol lên 2 mol). Khi tăng áp suất, cân bằng chuyển dịch theo chiều giảm số mol khí (chiều nghịch).\n"
            "c) Đúng vì thêm sản phẩm Y làm cân bằng chuyển dịch theo chiều nghịch để tiêu thụ bớt Y.\n"
            "d) Đúng vì chất xúc tác làm tăng tốc độ thuận và nghịch cùng một tỉ lệ, không làm chuyển dịch cân bằng."
        )
        add_solution_no_table(doc, sol_p2_c2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Độ acid và độ kiềm của dung dịch có thể được đánh giá bằng nồng độ H^+ (Nồng độ H^+ càng cao thì pH càng nhỏ) hoặc quy về một giá trị gọi là pH (pH là chỉ số đánh giá độ acid hay base của một dung dịch).")
    stmts_p2_c3 = [
        ("a) Để so sánh mức độ acid giữa các dung dịch có thể dựa vào nồng độ: dung dịch acid nào có nồng độ mol/lít lớn hơn sẽ có tính acid mạnh hơn.", "SAI"),
        ("b) Cho các dung dịch có cùng nồng độ: K~2~SO~3~ (1), NaClO~4~ (2), HNO~3~ (3), Ca(OH)~2~ (4). Chất có giá trị pH cao nhất là (3).", "SAI"),
        ("c) Cho ba dung dịch có cùng nồng độ: NH~3~ (1), Ca(OH)~2~ (2), KOH (3). Giá trị pH các dung dịch được sắp xếp theo thứ tự giảm dần là: (2), (3), (1).", "ĐÚNG"),
        ("d) Trong các dung dịch có cùng nồng độ, dung dịch có nồng độ ion H^+ nhỏ hơn và pH cao sẽ có tính acid yếu hơn.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c3:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c3 = (
            "a) Sai vì độ mạnh của acid phụ thuộc vào khả năng phân li ra H^+ (hằng số K~a~), không chỉ dựa vào nồng độ mol.\n"
            "b) Sai vì HNO~3~ (3) là acid mạnh nên pH nhỏ nhất; Ca(OH)~2~ (4) là base mạnh nên pH cao nhất.\n"
            "c) Đúng vì Ca(OH)~2~ phân li ra 2 OH^- ➔ [OH^-] lớn nhất ➔ pH lớn nhất; KOH phân li ra 1 OH^-; NH~3~ là base yếu phân li một phần ➔ pH giảm dần: (2) > (3) > (1).\n"
            "d) Đúng vì [H^+] càng nhỏ (pH càng lớn) thì tính acid càng yếu."
        )
        add_solution_no_table(doc, sol_p2_c3)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Bản chất acid – base của nước được thể hiện trong hình sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau4_mo_hinh_nuoc.png"), width_cm=7.5)
    stmts_p2_c4 = [
        ("a) Hai phân tử nước trong phương trình điện li trên là acid.", "SAI"),
        ("b) Hai phân tử nước trong phương trình điện li trên là base.", "SAI"),
        ("c) Nước là một chất lưỡng tính vì nó có thể đóng vai trò là một acid hoặc một base.", "ĐÚNG"),
        ("d) Hai phân tử nước trong phương trình điện li trên một phân tử đóng vai trò acid và một phân tử đóng vai trò base.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c4:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c4 = (
            "Trong quá trình tự ion hóa của nước: 2H~2~O ⇌ H~3~O^+ + OH^-:\n"
            "- Một phân tử H~2~O nhường proton H^+ đóng vai trò là acid.\n"
            "- Phân tử H~2~O còn lại nhận proton H^+ đóng vai trò là base.\n"
            "➔ Do đó nước là chất lưỡng tính theo thuyết Brønsted – Lowry. Ý a và b sai, ý c và d đúng."
        )
        add_solution_no_table(doc, sol_p2_c4)

    # ==========================================
    # PHẦN III
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 6."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. Nồng độ mol của ion NO~3~^- trong dung dịch Ba(NO~3~)~2~ 0,1 M là x M. Xác định giá trị của x.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *0,2*", base_size=10.5)
        add_solution_no_table(doc, "Ba(NO~3~)~2~ ➔ Ba^2+^ + 2NO~3~^-\n0,1 M ──────────➔ 0,2 M\nVậy [NO~3~^-] = 0,2 M ➔ x = 0,2.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 (SBT-CTST). Để chuẩn độ 40 mL dung dịch HCl chưa biết nồng độ đã dùng trung bình hết 34 mL dung dịch NaOH 0,12 M. Tính nồng độ mol của dung dịch HCl.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *0,102*", base_size=10.5)
        add_solution_no_table(doc, "n(NaOH) = 0,034 . 0,12 = 0,00408 mol.\nHCl + NaOH ➔ NaCl + H~2~O  ➔  n(HCl) = n(NaOH) = 0,00408 mol.\nC~M(HCl)~ = 0,00408 / 0,040 = 0,102 M.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3 [CD - SBT]. Xác định pH của dung dịch thu được sau khi thêm 25,0 ml dung dịch NaOH 0,1M vào 50,0 ml dung dịch HCl 0,1 M.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *1,48*", base_size=10.5)
        add_solution_no_table(doc, "n(NaOH) = 0,025 . 0,1 = 2,5.10^-3^ mol; n(HCl) = 0,050 . 0,1 = 5,0.10^-3^ mol.\nH^+ + OH^- ➔ H~2~O  ➔  n(H^+ dư) = 5,0.10^-3^ - 2,5.10^-3^ = 2,5.10^-3^ mol.\nV~dd~ = 25,0 + 50,0 = 75,0 mL = 0,075 L.\n[H^+] = 2,5.10^-3^ / 0,075 = 1/30 M ≈ 0,0333 M ➔ pH = -log(0,0333) ≈ 1,48.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4 [CD - SBT]. Có bao nhiêu phát biểu nào dưới đây là đúng trong số các phát biểu đã cho dưới đây?\n(a) Để so sánh mức độ acid giữa các dung dịch có thể dựa vào nồng độ: dung dịch acid nào có nồng độ lớn hơn sẽ có tính acid mạnh hơn.\n(b) Trong các dung dịch có cùng nồng độ, dung dịch nào có tính acid mạnh hơn sẽ có nồng độ ion H^+ lớn hơn và pH lớn hơn.\n(c) Trong các dung dịch có cùng nồng độ, dung dịch nào có nồng độ ion OH^- lớn hơn và pH nhỏ hơn sẽ có tính base lớn hơn.\n(d) Trong các dung dịch có cùng nồng độ, dung dịch nào có tính acid mạnh hơn sẽ có nồng độ ion H^+ lớn hơn và pH nhỏ hơn.\n(e) Trong các dung dịch có cùng nồng độ, dung dịch có nồng độ ion H^+ nhỏ hơn và pH cao sẽ có tính acid yếu hơn.\n(g) Trong một dãy các dung dịch có cùng nồng độ được sắp xếp theo tính acid tăng dần thì nồng độ ion OH^- sẽ giảm dần và K~a~ tăng dần.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *3*", base_size=10.5)
        add_solution_no_table(doc, "Có 3 phát biểu đúng là: (d), (e), (g).\n- (a) Sai vì độ mạnh acid phụ thuộc khả năng phân li ra H^+.\n- (b) Sai vì tính acid mạnh hơn thì pH phải nhỏ hơn.\n- (c) Sai vì tính base lớn hơn thì pH phải lớn hơn.\n- (d) Đúng vì tính acid mạnh hơn thì [H^+] lớn hơn và pH nhỏ hơn.\n- (e) Đúng vì [H^+] nhỏ hơn (pH cao hơn) thì tính acid yếu hơn.\n- (g) Đúng vì khi tính acid tăng dần thì [H^+] tăng ➔ [OH^-] giảm và K~a~ tăng.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5 [CD - SGK]. Để trung hoà 10mL dung dịch HCl nồng độ x M cần 50 ml dung dịch NaOH 0,5 M. Xác định giá trị của x.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *2,5*", base_size=10.5)
        add_solution_no_table(doc, "n(NaOH) = 0,050 . 0,5 = 0,025 mol.\nHCl + NaOH ➔ NaCl + H~2~O  ➔  n(HCl) = n(NaOH) = 0,025 mol.\nx = C~M(HCl)~ = 0,025 / 0,010 = 2,5 M.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6 (SBT - KNTT). Vỏ trứng có chứa calcium ở dạng CaCO~3~. Để xác định hàm lượng CaCO~3~ trong vỏ trứng, trong phòng thí nghiệm người ta có thể làm như sau:\nLấy 1,0 g vỏ trứng khô, đã được làm sạch, hoà tan hoàn toàn trong 50 mL dung dịch HCl 0,4 M. Lọc dung dịch sau phản ứng thu được 50 mL dung dịch A. Lấy 10,0 mL dung dịch A chuẩn độ với dung dịch NaOH 0,1 M thấy hết 5,6 mL. Xác định hàm lượng CaCO~3~ trong vỏ trứng (giả thiết các tạp chất khác trong vỏ trứng không phản ứng với HCl). Kết quả làm tròn đến hàng đơn vị phần trăm.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *86%*", base_size=10.5)
        add_solution_no_table(doc, "n(HCl ban đầu) = 0,050 . 0,4 = 0,02 mol.\nChuẩn độ 10,0 mL A: n(NaOH) = 0,0056 . 0,1 = 5,6.10^-4^ mol ➔ n(HCl trong 10 mL A) = 5,6.10^-4^ mol.\nTrong 50 mL dung dịch A (HCl dư): n(HCl dư) = 5,6.10^-4^ . (50 / 10) = 2,8.10^-3^ mol.\nn(HCl phản ứng) = 0,02 - 2,8.10^-3^ = 0,0172 mol.\nCaCO~3~ + 2HCl ➔ CaCl~2~ + CO~2~ + H~2~O  ➔  n(CaCO~3~) = 0,0172 / 2 = 8,6.10^-3^ mol.\nm(CaCO~3~) = 8,6.10^-3^ . 100 = 0,86 g.\nHàm lượng CaCO~3~ = (0,86 / 1,0) . 100% = 86%.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # ==========================================
    # PHẦN IV
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN IV. Tự luận",
        subtitle="Học sinh giải chi tiết các bài toán sau vào giấy làm bài."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1 [CD - SBT]. Điền từ ngữ thích hợp vào các chỗ trống trong mỗi phát biểu sau:\na) Phản ứng thuận nghịch là phản ứng hóa học trong đó ở cùng điều kiện, xảy ra ...(1)... sự chuyển chất phản ứng thành sản phẩm và sự chuyển ...(2)... thành ...(3)...\nb) Trạng thái cân bằng của mọi phản ứng thuận nghịch luôn có tốc độ phản ứng thuận ...(1)... tốc độ phản ứng nghịch, các phản ứng thuận và nghịch luôn diễn ra. Như vậy, cân bằng hóa học là ...(2)...\nc) Với một phản ứng hóa học, khi hằng số cân bằng rất lớn so với 1 thì ở trạng thái cân bằng, nồng độ các chất sản phẩm ...(1)... nồng độ ...(2)...")
    if is_teacher:
        sol_p4_c1 = (
            "a) (1) đồng thời; (2) chất sản phẩm; (3) chất phản ứng.\n"
            "b) (1) bằng; (2) cân bằng động.\n"
            "c) (1) lớn hơn; (2) các chất phản ứng."
        )
        add_solution_no_table(doc, sol_p4_c1)
    else:
        add_student_writing_lines(doc, num_lines=5)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 [CTST - SGK]. Ở các vùng quê, người dân thường dùng phèn chua để làm trong nước nhờ ứng dụng của phản ứng thuỷ phân ion Al^3+^? Giải thích? Chất hay ion nào là acid, là base trong phản ứng thuỷ phân Al^3+^?")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau2_tuan_phen_chua.png"), width_cm=7.5)
    if is_teacher:
        sol_p4_c2 = (
            "- Khi hòa tan phèn chua (KAl(SO~4~)~2~.12H~2~O) vào nước, ion Al^3+^ bị thuỷ phân tạo thành dạng phức [Al(H~2~O)~6~]^3+^.\n"
            "  Ion này nhường proton H^+ cho nước theo cân bằng:\n"
            "  [Al(H~2~O)~6~]^3+^ + H~2~O ⇌ [Al(OH)(H~2~O)~5~]^2+^ + H~3~O^+\n"
            "  hoặc viết gọn: Al^3+^ + 3H~2~O ⇌ Al(OH)~3~↓ + 3H^+\n"
            "- Kết tủa Al(OH)~3~ keo xốp, lơ lửng trong nước, có bề mặt hấp phụ rất lớn nên kéo dính các hạt bụi bẩn lơ lửng lại và lắng xuống đáy, làm nước trong sạch.\n"
            "- Trong phản ứng thủy phân trên: Al^3+^ (dạng [Al(H~2~O)~6~]^3+^) nhường H^+ nên là acid; H~2~O nhận H^+ nên là base."
        )
        add_solution_no_table(doc, sol_p4_c2)
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Lưu file
    filename = "De_Kiem_Tra_Hoa_11_De01_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_11_De01.docx"
    filepath = os.path.join(OUT_DIR, filename)
    doc.save(filepath)
    print(f"Saved: {filepath}")

if __name__ == "__main__":
    print("Generating De_Kiem_Tra_Hoa_11_De01.docx (Học Sinh)...")
    build_exam_document(is_teacher=False)
    print("Generating De_Kiem_Tra_Hoa_11_De01_GiaoVien.docx (Giáo Viên)...")
    build_exam_document(is_teacher=True)
    print("Done!")
