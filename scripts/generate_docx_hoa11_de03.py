# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file Word cho ĐỀ SỐ 3 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De03.docx (Bản Học Sinh)
- De_Kiem_Tra_Hoa_11_De03_GiaoVien.docx (Bản Giáo Viên)
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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de03")

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
        exam_code="103",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; I = 127; Ba = 137",
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
    render_question_prompt(p, "Câu 1. Cho cân bằng hoá học: H~2~(g) + I~2~(g) ⇌ 2HI(g); Δ~r~H°~298~ > 0. Cân bằng không bị chuyển dịch khi")
    opts_c1 = [
        "A. Tăng nhiệt độ của hệ.",
        "B. Giảm nồng độ HI.",
        "C. Tăng nồng độ H~2~.",
        "*D. Giảm áp suất chung của hệ.*"
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phản ứng có tổng số mol khí ở 2 vế bằng nhau (1 + 1 = 2) nên sự biến đổi áp suất không làm chuyển dịch cân bằng hóa học.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Cho cân bằng hóa học: N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g). Phản ứng thuận là phản ứng tỏa nhiệt. Cân bằng hóa học không bị dịch chuyển khi")
    opts_c2 = [
        "A. Thay đổi áp suất của hệ.",
        "B. Thay đổi nồng độ N~2~.",
        "C. Thay đổi nhiệt độ.",
        "*D. Thêm chất xúc tác.*"
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Chất xúc tác làm tăng tốc độ phản ứng thuận và tốc độ phản ứng nghịch với số lần bằng nhau, giúp hệ nhanh đạt trạng thái cân bằng chứ không làm dịch chuyển cân bằng.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Cho cân bằng hoá học: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g); phản ứng thuận là phản ứng tỏa nhiệt. Phát biểu đúng là:")
    opts_c3 = [
        "A. Cân bằng chuyển dịch theo chiều thuận khi tăng nhiệt độ.",
        "*B. Cân bằng chuyển dịch theo chiều nghịch khi giảm nồng độ O~2~.*",
        "C. Cân bằng chuyển dịch theo chiều thuận khi giảm áp suất hệ phản ứng.",
        "D. Cân bằng chuyển dịch theo chiều nghịch khi giảm nồng độ SO~3~."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khi giảm nồng độ O~2~, theo nguyên lí Le Chatelier cân bằng chuyển dịch theo chiều tạo thêm O~2~ (chiều nghịch).")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Các yếu tố ảnh hưởng đến cân bằng hoá học là")
    opts_c4 = [
        "A. Nồng độ, nhiệt độ và chất xúc tác.",
        "B. Nồng độ, áp suất và diện tích bề mặt.",
        "*C. Nồng độ, nhiệt độ và áp suất.*",
        "D. Áp suất, nhiệt độ và chất xúc tác."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Các yếu tố làm chuyển dịch cân bằng hóa học là nồng độ, nhiệt độ và áp suất (với hệ có chất khí và có biến thiên số mol khí). Chất xúc tác và diện tích bề mặt không làm chuyển dịch cân bằng.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Dãy các chất gồm những chất điện li mạnh là:")
    opts_c5 = [
        "*A. NaF, NaOH, KCl, BaCl~2~*",
        "B. HCl, NaCl, Na~2~CO~3~, Hg(CN)~2~",
        "C. KNO~3~, MgCl~2~, HNO~3~, HF",
        "D. NaOH, KCl, H~2~SO~4~, KOH, HClO"
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Dãy A gồm muối tan NaF, KCl, BaCl~2~ và base mạnh NaOH đều là các chất điện li mạnh. (HF, HClO là acid yếu; Hg(CN)~2~ là chất điện li rất yếu).")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Nhúng giấy quỳ vào dung dịch có pH = 4 giấy quỳ chuyển thành màu:")
    opts_c6 = [
        "*A. Đỏ*",
        "B. Xanh",
        "C. Không đổi màu",
        "D. Chưa xác định được"
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Dung dịch có pH = 4 < 7 có môi trường acid nên làm quỳ tím chuyển sang màu đỏ.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 7. Dung dịch chất nào sau đây làm quỳ tím chuyển thành màu đỏ?")
    opts_c7 = [
        "*A. H~2~SO~4~.*",
        "B. KOH.",
        "C. NaCl.",
        "D. C~2~H~5~OH."
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "H~2~SO~4~ là acid mạnh làm quỳ tím hóa đỏ. KOH làm quỳ hóa xanh, NaCl và C~2~H~5~OH không làm đổi màu quỳ.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 8. Dung dịch của một base ở 25°C có:")
    opts_c8 = [
        "A. [H^+] = 1,0.10^-7^ M",
        "*B. [H^+] < 1,0.10^-7^ M*",
        "C. [H^+] > 1,0.10^-7^ M",
        "D. [H^+].[OH^-] > 1,0.10^-14^"
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Dung dịch base có [OH^-] > 1,0.10^-7^ M ➔ [H^+] = 10^-14^ / [OH^-] < 1,0.10^-7^ M.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 9. Dung dịch acid mạnh H~2~SO~4~ 0,10M có:")
    opts_c9 = [
        "A. pH = 1,00",
        "*B. pH < 1,00*",
        "C. pH > 1,00",
        "D. [H^+] > 0,20M"
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "H~2~SO~4~ nấc 1 phân li hoàn toàn tạo 0,10 M H^+: H~2~SO~4~ ➔ H^+ + HSO~4~^-. Nấc 2 tiếp tục phân li một phần: HSO~4~^- ⇌ H^+ + SO~4~^2-. Do đó tổng [H^+] > 0,10 M ➔ pH = -log[H^+] < -log(0,10) = 1,00.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 10. Xét các hệ cân bằng sau đây trong một bình kín:\n(1) 2NaHCO~3~(s) ⇌ Na~2~CO~3~(s) + H~2~O(g) + CO~2~(g)\n(2) CO~2~(g) + CaO(s) ⇌ CaCO~3~(s)\n(3) C(s) + CO~2~(g) ⇌ 2CO(g)\n(4) CO(g) + H~2~O(g) ⇌ CO~2~(g) + H~2~(g)\nKhi thêm CO~2~ vào hệ thì số cân bằng chuyển dịch theo chiều thuận là")
    opts_c10 = [
        "*A. 2.*",
        "B. 1.",
        "C. 4.",
        "D. 3."
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khi thêm CO~2~, cân bằng chuyển dịch theo chiều làm giảm nồng độ CO~2~ (chiều tiêu thụ CO~2~).\n- Phản ứng (2) và (3) có CO~2~ là chất tham gia nên chuyển dịch theo chiều thuận.\n- Phản ứng (1) và (4) có CO~2~ là sản phẩm nên chuyển dịch theo chiều nghịch.\n➔ Có 2 cân bằng chuyển dịch theo chiều thuận.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 11. Cho cân bằng hoá học sau: 2NH~3~(g) ⇌ N~2~(g) + 3H~2~(g). Khi tăng nhiệt độ của hệ thì tỉ khối của hỗn hợp so với hydrogen giảm. Nhận xét nào sau đây là đúng?")
    opts_c11 = [
        "*A. Khi tăng nhiệt độ của hệ, cân bằng chuyển dịch theo chiều thuận.*",
        "B. Phản ứng thuận là phản ứng toả nhiệt.",
        "C. Khi tăng áp suất của hệ, cân bằng chuyển dịch theo chiều thuận.",
        "D. Khi tăng nồng độ của NH~3~, cân bằng chuyển dịch theo chiều nghịch."
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Tỉ khối so với H~2~ giảm ➔ M~hh~ giảm. Khối lượng hỗn hợp không đổi (m~hh~ = const) nên số mol khí n~hh~ tăng. Chiều tăng số mol khí là chiều thuận (từ 2 mol lên 4 mol khí). Vậy tăng nhiệt độ làm cân bằng chuyển dịch theo chiều thuận ➔ Phản ứng thuận là phản ứng thu nhiệt (Δ~r~H > 0).")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 12. Sự dịch chuyển cân bằng hoá học là sự di chuyển từ trạng thái cân bằng hoá học này")
    opts_c12 = [
        "A. Sang trạng thái cân bằng hóa học khác không cần có tác động của các yếu tố từ bên ngoài tác động lên cân bằng.",
        "B. Sang trạng thái không cân bằng do tác động của các yếu tố từ bên ngoài tác động lên cân bằng.",
        "*C. Sang trạng thái cân bằng hoá học khác do tác động của các yếu tố từ bên ngoài tác động lên cân bằng.*",
        "D. Sang trạng thái cân bằng hoá học khác do cân bằng hóa học tác động lên các yếu tố bên ngoài."
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Định nghĩa SGK: Sự chuyển dịch cân bằng hoá học là sự dịch chuyển từ trạng thái cân bằng này sang trạng thái cân bằng khác do các tác động từ bên ngoài tác động lên cân bằng.")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 13 (SBT - KNTT). Cho phản ứng hoá học sau: PCl~3~(g) + Cl~2~(g) ⇌ PCl~5~(g) ở T°C, nồng độ các chất ở trạng thái cân bằng như sau: [PCl~5~] = 0,059 mol/L; [PCl~3~] = [Cl~2~] = 0,035 mol/L. Hằng số cân bằng (K~C~) của phản ứng tại T°C là")
    opts_c13 = [
        "A. 1,68.",
        "*B. 48,16.*",
        "C. 0,02.",
        "D. 16,95."
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "K~C~ = [PCl~5~] / ([PCl~3~].[Cl~2~]) = 0,059 / (0,035 . 0,035) ≈ 48,16.")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 14. Trong một bình kín có cân bằng hóa học sau: 2NO~2~(g) ⇌ N~2~O~4~(g). Tỉ khối hơi của hỗn hợp khí trong bình so với H~2~ ở nhiệt độ T~1~ bằng 27,6 và ở nhiệt độ T~2~ bằng 34,5. Biết T~1~ > T~2~. Phát biểu nào sau đây về cân bằng trên là đúng?")
    opts_c14 = [
        "*A. Phản ứng thuận là phản ứng tỏa nhiệt.*",
        "B. Khi tăng nhiệt độ, áp suất chung của hệ cân bằng giảm.",
        "C. Khi giảm nhiệt độ, áp suất chung của hệ cân bằng tăng.",
        "D. Phản ứng nghịch là phản ứng tỏa nhiệt."
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "T~1~ > T~2~ (nhiệt độ giảm). Khi giảm nhiệt độ, tỉ khối hỗn hợp tăng ➔ M~hh~ tăng ➔ số mol khí giảm (chiều thuận từ 2 mol NO~2~ thành 1 mol N~2~O~4~). Vậy hạ nhiệt độ làm cân bằng dịch chuyển theo chiều thuận ➔ Phản ứng thuận là phản ứng tỏa nhiệt (Δ~r~H < 0).")

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 15. Khi hoà tan SO~2~ vào nước có cân bằng sau: SO~2~ + H~2~O ⇌ HSO~3~^- + H^+. Khi cho thêm NaOH và khi cho thêm H~2~SO~4~ loãng vào dung dịch trên thì cân bằng sẽ chuyển dịch tương ứng là")
    opts_c15 = [
        "A. Thuận và thuận.",
        "*B. Thuận và nghịch.*",
        "C. Nghịch và thuận.",
        "D. Nghịch và nghịch."
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "- Thêm NaOH (chứa OH^-), OH^- trung hòa H^+ làm nồng độ H^+ giảm ➔ cân bằng chuyển dịch theo chiều thuận tạo thêm H^+.\n- Thêm H~2~SO~4~ (chứa H^+), nồng độ H^+ tăng ➔ cân bằng chuyển dịch theo chiều nghịch làm giảm bớt H^+.")

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 16. Cho các cân bằng sau:\n(1) N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)\n(2) H~2~(g) + I~2~(g) ⇌ 2HI(g)\n(3) CaCO~3~(s) ⇌ CaO(s) + CO~2~(g)\nKhi tăng áp suất chung của hệ thì cân bằng:")
    opts_c16 = [
        "A. (1) và (3) dịch chuyển theo chiều thuận; (2) không dịch chuyển.",
        "*B. (1) dịch chuyển theo chiều thuận; (2) không dịch chuyển; (3) dịch chuyển theo chiều nghịch.*",
        "C. (1) và (3) dịch chuyển theo chiều nghịch; (2) không dịch chuyển.",
        "D. (1) dịch chuyển theo chiều nghịch; (2) không dịch chuyển; (3) dịch chuyển theo chiều thuận."
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khi tăng áp suất chung, cân bằng chuyển dịch theo chiều làm giảm số mol khí:\n- (1): 4 mol khí ⇌ 2 mol khí ➔ chuyển dịch theo chiều thuận.\n- (2): 2 mol khí ⇌ 2 mol khí ➔ không chuyển dịch.\n- (3): 0 mol khí ⇌ 1 mol khí ➔ chuyển dịch theo chiều nghịch.")

    # Câu 17
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 17. Trộn hai thể tích bằng nhau của dung dịch H~2~SO~4~ và dung dịch NaOH có cùng nồng độ mol, pH của dung dịch sau phản ứng sẽ như thế nào?")
    opts_c17 = [
        "A. pH = 7",
        "B. pH > 7",
        "*C. pH < 7*",
        "D. Không xác định được"
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "H~2~SO~4~ + 2NaOH ➔ Na~2~SO~4~ + 2H~2~O. Vì trộn 2 thể tích bằng nhau cùng nồng độ mol nên n(H~2~SO~4~) = n(NaOH). Theo tỉ lệ phản ứng 1:2 thì H~2~SO~4~ dư ➔ dung dịch có môi trường acid ➔ pH < 7.")

    # Câu 18
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 18. Cho phản ứng: H~2~(g) + I~2~(g) ⇌ 2HI(g). Ở nhiệt độ 430°C, hằng số cân bằng K~C~ của phản ứng trên bằng 53,96. Đun nóng một bình kín dung tích không đổi 10 lít chứa 4,0 gam H~2~ và 406,4 gam I~2~. Khi hệ phản ứng đạt trạng thái cân bằng ở 430°C, nồng độ của HI là")
    opts_c18 = [
        "*A. 0,275M.*",
        "B. 0,225M.",
        "C. 0,151M.",
        "D. 0,320M."
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "n(H~2~) = 4/2 = 2 mol; n(I~2~) = 406,4/254 = 1,6 mol.\nGọi x là số mol H~2~ phản ứng. Tại cân bằng: n(H~2~) = 2 - x; n(I~2~) = 1,6 - x; n(HI) = 2x.\nK~C~ = (2x/10)^2^ / [((2 - x)/10)((1,6 - x)/10)] = 4x^2^ / [(2 - x)(1,6 - x)] = 53,96\n➔ 49,96x^2^ - 194,256x + 172,672 = 0 ➔ x ≈ 1,375 mol (nhận vì x < 1,6).\n➔ [HI] = 2x / 10 = 2 . 1,375 / 10 = 0,275 M.")

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
    render_question_prompt(p, "Câu 1. Cho các phát biểu sau về trạng thái cân bằng hóa học:")
    stmts_p2_c1 = [
        ("a) Ở trạng thái cân bằng, tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch.", "ĐÚNG"),
        ("b) Ở trạng thái cân bằng, các chất không phản ứng với nhau.", "SAI"),
        ("c) Ở trạng thái cân bằng, nồng độ các chất sản phẩm luôn lớn hơn nồng độ các chất đầu.", "SAI"),
        ("d) Ở trạng thái cân bằng, nồng độ các chất không thay đổi.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c1 = (
            "a) Đúng theo bản chất cân bằng động.\n"
            "b) Sai vì ở trạng thái cân bằng phản ứng thuận và phản ứng nghịch vẫn tiếp diễn với tốc độ bằng nhau.\n"
            "c) Sai vì nồng độ các chất ở trạng thái cân bằng giữ nguyên không đổi chứ không nhất thiết chất sản phẩm phải lớn hơn chất đầu.\n"
            "d) Đúng vì tốc độ sinh ra bằng tốc độ tiêu hao của mỗi chất."
        )
        add_solution_no_table(doc, sol_p2_c1)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Khi hoà tan khí chlorine vào nước tạo thành dung dịch có màu vàng nhạt gọi là nước chlorine. Trong nước chlorine xảy ra cân bằng hoá học sau:\nCl~2~ + H~2~O ⇌ HClO + HCl (1)\nAcid HClO sinh ra không bền, dễ bị phân huỷ theo phản ứng:\nHClO ➔ HCl + O (2)")
    stmts_p2_c2 = [
        ("a) Nước chlorine sẽ nhạt màu dần theo thời gian, không bảo quản được lâu.", "ĐÚNG"),
        ("b) Phản ứng phân huỷ acid HClO làm giảm nồng độ của chất này, cân bằng (1) chuyển dịch theo chiều thuận.", "ĐÚNG"),
        ("c) Nước chlorine sẽ đậm màu dần theo thời gian, nên bảo quản được lâu.", "SAI"),
        ("d) Phản ứng phân huỷ acid HClO làm tăng nồng độ của chất này, cân bằng (1) chuyển dịch theo chiều nghịch.", "SAI")
    ]
    for stmt, ans in stmts_p2_c2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c2 = (
            "a) Đúng vì HClO phân hủy liên tục làm tiêu hao dần Cl~2~ tan trong nước, khiến dung dịch mất dần màu vàng lục nhạt.\n"
            "b) Đúng vì HClO bị phân hủy làm giảm nồng độ sản phẩm, theo Le Chatelier cân bằng (1) chuyển dịch theo chiều thuận.\n"
            "c) Sai vì nước chlorine nhạt màu dần và mất tính tẩy màu theo thời gian.\n"
            "d) Sai vì phản ứng (2) làm giảm nồng độ HClO."
        )
        add_solution_no_table(doc, sol_p2_c2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Hoa đậu biếc tên tiếng Anh là butterfly pea, là một loại hoa được trồng phổ biến tại các quốc gia Nam Á và Đông Nam Á như Ấn Độ, Thái Lan, Singapore, Việt Nam... Hoa có tên gọi như vậy vì cây đậu biếc thuộc họ đậu, sống dây leo và cánh hoa có màu xanh tím giống hình con bướm. Khi sử dụng đậu biếc làm chất tạo màu tự nhiên, người ta thường đun sôi hoa đậu biếc với nước, hoặc ngâm cánh hoa trong nước sôi từ 10 đến 15 phút, màu hoa sẽ được trích ly tạo thành màu xanh biếc. Sắc tố tạo màu đặc trưng cho đậu biếc là các hợp chất thuộc nhóm anthocyanin, một trong những chất chống oxi hóa tự nhiên. Điểm đặc biệt của nhóm anthocyanin là màu của chúng thay đổi dưới tác dụng pH của môi trường. Ở môi trường pH < 7 (môi trường Acid), anthocyanin chuyển sang đỏ tím, ngược lại chúng chuyển sang màu xanh sẫm khi môi trường pH ≥ 7.")
    stmts_p2_c3 = [
        ("a) Vắt nước cốt chanh vào cốc nước hoa đậu biếc đã được ngâm nóng trong nước sôi thì ta thấy màu của cốc nước chuyển từ màu xanh biếc sang màu đỏ tím.", "ĐÚNG"),
        ("b) Cho nước rửa bát vào cốc nước hoa đậu biếc đã được ngâm nóng trong nước sôi thì ta thấy màu của cốc nước chuyển từ màu xanh biếc sang màu xanh sẫm ngay lập tức.", "ĐÚNG"),
        ("c) Cho nước muối sinh lí vào cốc nước hoa đậu biếc đã được ngâm nóng trong nước sôi thì ta thấy màu của cốc nước không thay đổi.", "ĐÚNG"),
        ("d) Hoa đậu biếc là một chất chỉ thị màu trong tự nhiên.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c3:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c3 = (
            "a) Đúng vì nước chanh chứa citric acid có pH < 7, làm anthocyanin chuyển sang màu đỏ tím.\n"
            "b) Đúng vì nước rửa bát có tính base nhẹ (pH ≥ 7), làm anthocyanin chuyển sang màu xanh sẫm.\n"
            "c) Đúng vì nước muối sinh lí NaCl 0,9% có môi trường trung tính (pH ≈ 7) nên không làm biến đổi màu hoa đậu biếc.\n"
            "d) Đúng vì màu sắc của nước hoa đậu biếc biến đổi nhạy bén theo pH của dung dịch."
        )
        add_solution_no_table(doc, sol_p2_c3)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. NaCl được gọi là chất điện li, tan vào nước tạo ra ion và thu được dung dịch dẫn điện, gọi là dung dịch chất điện li.")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau4_p2_phan_li_nacl.png"), width_cm=11.5)
    stmts_p2_c4 = [
        ("a) Các phân tử nước kéo các ion sodium và ion chloride ra xa nhau, phá vỡ liên kết ion giữa hai ion sodium và ion chloride.", "ĐÚNG"),
        ("b) Sau khi các hợp chất muối được tách ra, các ion sodium và ion chloride được bao quanh bởi các phân tử nước. Lúc này, muối sẽ bị hòa tan, tạo thành dung dịch đồng nhất (dung dịch chất dẫn điện).", "ĐÚNG"),
        ("c) Đầu âm của các phân tử nước bị thu hút bởi các ion chloride tích điện dương, còn đầu dương của các phân tử nước bị thu hút bởi các ion sodium tích điện âm.", "SAI"),
        ("d) Sự điện li của NaCl trong nước có thể được biểu diễn bằng phương trình điện li như sau: NaCl + (a+b)H~2~O ➔ Na^+.aH~2~O + Cl^-.bH~2~O.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c4:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c4 = (
            "a) Đúng theo cơ chế solvate hoá của phân tử nước phân cực.\n"
            "b) Đúng vì các ion hydrate hoá phân tán đồng đều tạo thành dung dịch điện li dẫn điện.\n"
            "c) Sai vì phân tử nước có cực âm ở nguyên tử oxygen (hút ion dương Na^+) và cực dương ở nguyên tử hydrogen (hút ion âm Cl^-).\n"
            "d) Đúng, đây là phương trình biểu diễn chính xác sự điện li kèm theo sự hydrate hoá ion."
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
    render_question_prompt(p, "Câu 1. Độ pH của da thường từ 4,5 - 5,5, có tính Acid nhẹ. Điều này giúp da của chúng ta chống lại vi khuẩn trên bề mặt. Đối với da đầu, độ pH của da đầu thường từ 4 - 7. Sản phẩm dành cho tóc dầu sẽ có độ pH cao hơn 5,5, chứa chất tẩy rửa mạnh giúp loại bỏ chất dầu nhờn trên da đầu. Bên cạnh đó, nếu bạn thường xuyên uốn nhuộm thì những dòng sản phẩm dành cho tóc hư tổn với khả năng bổ sung tính Acid giúp cân bằng độ pH trên da đầu và tóc sẽ rất hữu ích.")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau1_p3_dau_goi.png"), width_cm=7.5)
    p_q1 = doc.add_paragraph()
    p_q1.paragraph_format.space_before = Pt(2)
    p_q1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_q1, "Một loại dầu gội đầu có nồng độ ion OH^- là 10^-5,17^ mol/L. Tính pH của loại dầu gội nói trên (kết quả làm tròn đến hai chữ số phần thập phân).", base_size=11)
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *8,83*", base_size=10.5)
        add_solution_no_table(doc, "pOH = -log[OH^-] = 5,17 ➔ pH = 14 - pOH = 14 - 5,17 = 8,83.\n(Hoặc [H^+] = 10^-14^ / 10^-5,17^ = 10^-8,83^ M ➔ pH = 8,83).")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 (SBT - KNTT). Aspirin là một loại thuốc có thành phần chính là acetylsalicylic acid. Nếu hoà tan thuốc này vào nước, người ta xác định được pH của dung dịch tạo thành là 2,8. Nồng độ OH^- của dung dịch tạo thành là 10^x^. Xác định x?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *-11,2*", base_size=10.5)
        add_solution_no_table(doc, "[H^+] = 10^-2,8^ M ➔ [OH^-] = 10^-14^ / 10^-2,8^ = 10^-11,2^ M ➔ x = -11,2.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3 [CD - SGK]. Chuẩn độ dung dịch NaOH chưa biết chính xác nồng độ (biết nồng độ trong khoảng gần với 0,1 M) bằng dung dịch chuẩn HCl 0,1 M với chất chỉ thị phenolphthalein hết 10 mL. Giả sử khi kết thúc chuẩn độ, thể tích dung dịch NaOH (tính trung bình sau 3 lần chuẩn độ) đã sử dụng ở burette là 10,27 mL. Nồng độ của dung dịch NaOH là x (M). Xác định giá trị của x (kết quả làm tròn đến hàng phần mười).")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *0,1*", base_size=10.5)
        add_solution_no_table(doc, "n(HCl) = 0,010 . 0,1 = 0,001 mol.\nHCl + NaOH ➔ NaCl + H~2~O ➔ n(NaOH) = 0,001 mol.\nC~M~(NaOH) = 0,001 / (10,27 . 10^-3^) ≈ 0,0974 M ≈ 0,1 M ➔ x = 0,1.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Cần bao nhiêu gam NaOH để pha chế 250,0 mL dung dịch có pH = 12?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *0,1*", base_size=10.5)
        add_solution_no_table(doc, "pH = 12 ➔ pOH = 2 ➔ [OH^-] = 10^-2^ M = 0,01 M.\nn(NaOH) = 0,01 . 0,250 = 0,0025 mol ➔ m(NaOH) = 0,0025 . 40 = 0,1 gam.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5 [CD - SBT]. Xét phản ứng: H~2~(g) + I~2~(g) ⇌ 2HI(g). Một hỗn hợp phản ứng chứa trong bình dung tích 3,67 lít ở một nhiệt độ xác định, ban đầu chứa 0,763 gam H~2~ và 96,6 gam I~2~. Ở trạng thái cân bằng, bình chứa 90,4 gam HI. Tính hằng số cân bằng (K~C~) cho phản ứng ở nhiệt độ này (làm tròn kết quả đến phần nguyên).")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *608*", base_size=10.5)
        add_solution_no_table(doc, "[H~2~] ban đầu = 0,763 / (2 . 3,67) = 0,1040 M.\n[I~2~] ban đầu = 96,6 / (254 . 3,67) = 0,1040 M.\n[HI] cân bằng = 90,4 / (128 . 3,67) = 0,1924 M ➔ 2x = 0,1924 ➔ x = 0,0962 M.\n[H~2~] = [I~2~] = 0,1040 - 0,0962 = 0,0078 M.\nK~C~ = (0,1924)^2^ / (0,0078)^2^ ≈ 608,4 ≈ 608.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Trong công nghiệp, hydrogen được sản xuất từ phản ứng:\nCH~4~(g) + H~2~O(g) ⇌ 3H~2~(g) + CO(g)\nTính hằng số cân bằng K~C~ của phản ứng trên ở 760°C. Biết ở nhiệt độ này, tất cả các chất đều ở thể khí và nồng độ mol của CH~4~, H~2~O, H~2~ và CO ở trạng thái cân bằng lần lượt là 0,126 M; 0,242 M; 1,150 M và 0,126 M (làm tròn kết quả đến hàng phần trăm).")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *6,28*", base_size=10.5)
        add_solution_no_table(doc, "K~C~ = [H~2~]^3^ . [CO] / ([CH~4~] . [H~2~O]) = (1,150^3^ . 0,126) / (0,126 . 0,242) = 1,150^3^ / 0,242 ≈ 6,28.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # ==========================================
    # PHẦN IV
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN IV. Tự luận",
        subtitle="Học sinh giải chi tiết các câu hỏi sau vào giấy làm bài."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. [CD - SBT] Điền từ ngữ thích hợp vào chỗ trống trong mỗi phát biểu sau:\na) Quá trình phân li của các chất khi tan trong nước thành các ion được gọi là .....(1)..... Chất điện li là chất khi tan trong nước phân li thành các .....(2)..... .....(3)..... là chất khi tan trong nước không phân li thành các ion.\nb) Theo thuyết Brønsted – Lowry, .....(1)..... là những chất có khả năng cho H^+, .....(2)..... là những chất có khả năng nhận H^+. Acid mạnh và base mạnh phân li .....(3)..... trong nước; acid yếu và base yếu phân li .....(4)..... trong nước.")
    if is_teacher:
        sol_p4_c1 = (
            "Điền từ thích hợp:\n"
            "a) (1) sự điện li;   (2) ion;   (3) Chất không điện li.\n"
            "b) (1) acid;   (2) base;   (3) hoàn toàn;   (4) một phần (hoặc không hoàn toàn)."
        )
        add_solution_no_table(doc, sol_p4_c1)
    else:
        add_student_writing_lines(doc, num_lines=5)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. (SBT - KNTT)\na) CH~3~COOH (có trong giấm ăn) là một acid yếu. Tính pH của dung dịch CH~3~COOH 0,1 M (biết hằng số cân bằng của sự phân li CH~3~COOH là 1,8.10^-5^, bỏ qua sự phân li của nước).\nb) Trong dung dịch nước ion CH~3~COO^- nhận proton của nước. Viết phương trình thuỷ phân và cho biết môi trường của dung dịch CH~3~COONa.\nc) Cho 10 mL dung dịch NaOH 0,1 M vào 10 mL dung dịch CH~3~COOH 0,2 M thu được 20 mL dung dịch A. Tính pH của dung dịch A.")
    if is_teacher:
        sol_p4_c2 = (
            "a) Phương trình phân li:\n"
            "   CH~3~COOH ⇌ CH~3~COO^- + H^+    K~a~ = 1,8.10^-5^\n"
            "   K~a~ = x^2^ / (0,1 - x) = 1,8.10^-5^ ➔ x ≈ 1,33.10^-3^ M ➔ pH = -log(1,33.10^-3^) ≈ 2,88.\n"
            "b) Muối CH~3~COONa phân li hoàn toàn: CH~3~COONa ➔ CH~3~COO^- + Na^+\n"
            "   Ion CH~3~COO^- bị thuỷ phân:\n"
            "   CH~3~COO^- + H~2~O ⇌ CH~3~COOH + OH^-\n"
            "   Do tạo ra ion OH^- nên dung dịch CH~3~COONa có môi trường base (kiềm, pH > 7).\n"
            "c) n(NaOH) = 0,010 . 0,1 = 1,0.10^-3^ mol; n(CH~3~COOH) = 0,010 . 0,2 = 2,0.10^-3^ mol.\n"
            "   CH~3~COOH + NaOH ➔ CH~3~COONa + H~2~O\n"
            "   Sau phản ứng: n(CH~3~COOH dư) = 1,0.10^-3^ mol; n(CH~3~COONa) = 1,0.10^-3^ mol.\n"
            "   V~dd~ = 20 mL = 0,020 L ➔ [CH~3~COOH] = [CH~3~COO^-] = 0,05 M.\n"
            "   Cân bằng: CH~3~COOH ⇌ CH~3~COO^- + H^+\n"
            "   [H^+] = K~a~ . [CH~3~COOH] / [CH~3~COO^-] = 1,8.10^-5^ . (0,05 / 0,05) = 1,8.10^-5^ M.\n"
            "   ➔ pH = -log(1,8.10^-5^) ≈ 4,74."
        )
        add_solution_no_table(doc, sol_p4_c2)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Lưu file
    filename = "De_Kiem_Tra_Hoa_11_De03_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_11_De03.docx"
    filepath = os.path.join(OUT_DIR, filename)
    doc.save(filepath)
    print(f"Saved: {filepath}")

if __name__ == "__main__":
    print("Generating De_Kiem_Tra_Hoa_11_De03.docx (Học Sinh)...")
    build_exam_document(is_teacher=False)
    print("Generating De_Kiem_Tra_Hoa_11_De03_GiaoVien.docx (Giáo Viên)...")
    build_exam_document(is_teacher=True)
    print("Done!")
