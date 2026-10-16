# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file Word cho ĐỀ SỐ 9 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De09.docx (Bản Học Sinh)
- De_Kiem_Tra_Hoa_11_De09_GiaoVien.docx (Bản Giáo Viên)
Chương 3: ĐẠI CƯƠNG HÓA HỌC HỮU CƠ
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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de09")

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
        exam_title="ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 3 (ĐẠI CƯƠNG HÓA HỌC HỮU CƠ)",
        subject_line="MÔN: HOÁ HỌC - LỚP 11",
        num_pages="04",
        duration="45 phút",
        exam_code="109",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; F = 19; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Br = 80; Ag = 108",
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
    render_question_prompt(p, "Câu 1. Cho các chất sau: NaHCO~3~, CH~4~, C~2~H~6~O, CH~3~COONa, CaC~2~, Al~4~C~3~. Hợp chất nào thuộc loại hợp chất hữu cơ?")
    opts_c1 = [
        "*A. CH~4~, C~2~H~6~O, CH~3~COONa.*",
        "B. NaHCO~3~, CaC~2~, Al~4~C~3~.",
        "C. NaHCO~3~, CH~4~, C~2~H~6~O.",
        "D. CH~4~, C~2~H~6~O, Al~4~C~3~."
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hợp chất hữu cơ là hợp chất của carbon (trừ CO, CO~2~, muối carbonate, cyanide, carbide,...). Do đó CH~4~, C~2~H~6~O, CH~3~COONa là hợp chất hữu cơ; NaHCO~3~, CaC~2~, Al~4~C~3~ là các hợp chất vô cơ.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Trong thành phần phân tử hợp chất hữu cơ nhất thiết phải có nguyên tố nào sau đây?")
    opts_c2 = [
        "*A. carbon.*",
        "B. hydrogen.",
        "C. oxygen.",
        "D. nitrogen."
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hợp chất hữu cơ là hợp chất của carbon (trừ một số chất vô cơ đơn giản như CO, CO~2~, muối carbonate,...). Do đó trong thành phần phân tử hợp chất hữu cơ nhất thiết phải có nguyên tố carbon.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Phát biểu nào sau đây không đúng khi nói về phương pháp chưng cất?")
    opts_c3 = [
        "A. Phương pháp chưng cất dùng để tách các chất lỏng có nhiệt độ sôi khác nhau nhiều.",
        "*B. Chất nào có nhiệt độ sôi thấp sẽ chuyển thành hơi muộn hơn và ít hơn.*",
        "C. Chất có nhiệt độ sôi thấp hơn sẽ chuyển thành hơi sớm hơn và nhiều hơn.",
        "D. Hơi bay ra được dẫn qua ống sinh hàn làm lạnh để ngưng tụ thành chất lỏng."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Chất có nhiệt độ sôi thấp hơn sẽ chuyển thành hơi sớm hơn và nhiều hơn chất có nhiệt độ sôi cao, do đó phát biểu B sai.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Chiết là phương pháp tách chất dựa vào sự hòa tan khác nhau của ... trong hai môi trường không tan vào nhau. Cụm từ thích hợp điền vào chỗ trống là")
    opts_c4 = [
        "*A. chất cần tách.*",
        "B. chất lỏng.",
        "C. dung môi.",
        "D. chất tan."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phương pháp chiết dựa trên sự hòa tan khác nhau của chất cần tách trong hai pha không trộn lẫn (thường là hai chất lỏng không tan vào nhau).")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Hợp chất hữu cơ X có công thức phân tử C~x~H~y~O~z~, phân tử khối là M. Công thức tính giá trị của z là")
    opts_c5 = [
        "A. z = (%m~O~ / 16) \u00d7 100.",
        "B. z = (%m~O~ / M) \u00d7 100.",
        "C. z = (%m~O~ / M) \u00d7 (16 / 100).",
        "*D. z = (%m~O~ / 16) \u00d7 (M / 100).*"
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Ta có %m~O~ = (16z / M) \u00d7 100% \u21d2 z = (%m~O~ \u00d7 M) / (16 \u00d7 100) = (%m~O~ / 16) \u00d7 (M / 100).")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6 [SBT - KNTT]. Trong phương pháp phổ khối lượng, đối với các hợp chất đơn giản, thường mảnh có giá trị m/z lớn nhất ứng với mảnh ion phân tử [M^+^] và giá trị này bằng giá trị ....... của chất nghiên cứu. Cụm từ thích hợp điền vào chỗ trống là")
    opts_c6 = [
        "*A. phân tử khối.*",
        "B. nguyên tử khối.",
        "C. điện tích ion.",
        "D. khối lượng."
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Mảnh ion phân tử [M^+^] có điện tích z = +1 nên giá trị m/z của nó bằng đúng phân tử khối của chất nghiên cứu.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 7 [KNTT - SGK]. Cặp chất nào sau đây là đồng phân của nhau?")
    opts_c7 = [
        "A. C~2~H~5~OH và CH~3~-O-C~2~H~5~.",
        "B. CH~3~-O-CH~3~ và CH~3~CHO.",
        "*C. CH~3~-CH~2~-CH~2~-OH và CH~3~-CH(OH)-CH~3~.*",
        "D. CH~3~-CH~2~-CH~2~-CH~3~ và CH~3~-CH~2~-CH=CH~2~."
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "CH~3~-CH~2~-CH~2~-OH (propan-1-ol) và CH~3~-CH(OH)-CH~3~ (propan-2-ol) có cùng công thức phân tử C~3~H~8~O nhưng khác nhau về vị trí nhóm -OH, là đồng phân vị trí nhóm chức của nhau.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 8 [KNTT - SGK]. Trong các dãy chất sau đây, dãy nào gồm các chất là đồng đẳng của nhau?")
    opts_c8 = [
        "*A. CH~3~-CH~2~-OH và CH~3~-CH~2~-CH~2~-OH.*",
        "B. CH~3~-O-CH~3~ và CH~3~-CH~2~-OH.",
        "C. CH~4~, C~2~H~6~ và C~4~H~8~.",
        "D. CH~4~ và C~3~H~6~."
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "CH~3~-CH~2~-OH (ethanol) và CH~3~-CH~2~-CH~2~-OH (propan-1-ol) đều là alcohol no, đơn chức, mạch hở, phân tử hơn kém nhau 1 nhóm CH~2~ nên thuộc cùng dãy đồng đẳng.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 9. Ethanol có công thức cấu tạo là C~2~H~5~OH, tính chất đặc trưng của ethanol là tác dụng với kim loại hoạt động như Na, K, không tác dụng với dung dịch NaOH, KOH. Hợp chất hữu cơ X có cùng nhóm chức với ethanol, X không thể là chất nào sau đây?")
    opts_c9 = [
        "A. CH~3~OH.",
        "B. C~3~H~7~OH.",
        "C. CH~2~=CH-CH~2~OH.",
        "*D. CH~3~CHO.*"
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Ethanol chứa nhóm chức alcohol (-OH). CH~3~CHO chứa nhóm chức aldehyde (-CHO), không cùng loại nhóm chức với ethanol.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 10. Sơn, nhựa cao su tan ít hoặc không tan trong dung môi nào sau đây?")
    opts_c10 = [
        "A. Acetone.",
        "B. Xăng.",
        "C. Chloroform.",
        "*D. Nước.*"
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Sơn, nhựa cao su là các hợp chất hữu cơ không phân cực hoặc kém phân cực, dễ tan trong dung môi hữu cơ không phân cực (xăng, acetone, chloroform) và hầu như không tan trong nước (dung môi phân cực mạnh).")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 11. Phát biểu nào sau đây không chính xác?")
    opts_c11 = [
        "A. Chưng cất là phương pháp thuận tiện để tinh chế các chất lỏng có nhiệt độ sôi khác nhau.",
        "B. Chưng cất là phương pháp tinh chế, tách dựa trên sự khác nhau về thành phần của hỗn hợp hơi và thành phần hỗn hợp lỏng nằm cân bằng với nhau.",
        "C. Chưng cất là phương pháp tinh chế, tách dựa trên sự khác nhau về nhiệt độ sôi của các chất.",
        "*D. Chưng cất là phương pháp tinh chế, tách dựa trên sự khác nhau về tỉ khối hơi của các chất.*"
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phương pháp chưng cất dựa trên sự khác nhau về nhiệt độ sôi và độ bay hơi của các chất lỏng, không phải dựa trên tỉ khối hơi của các chất.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 12. Để tách tinh dầu sả (có trong thân, lá, rễ,... cây sả) trong công nghiệp người ta sử dụng phương pháp")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau12_p1_tinh_dau_sa.png"), width_cm=6.5)
    opts_c12 = [
        "A. Chưng cất bằng hơi nước và chiết lại bằng nước lạnh.",
        "*B. Chưng cất bằng hơi nước và chiết tinh dầu ra khỏi hỗn hợp sản phẩm.*",
        "C. Chiết tinh dầu sả sau đó chưng cất bằng hơi nước.",
        "D. Kết tinh tinh dầu sả."
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Tinh dầu sả dễ bay hơi theo hơi nước và không tan trong nước, nhẹ hơn nước. Do đó dùng phương pháp chưng cất lôi cuốn hơi nước, sau đó ngưng tụ thu được hỗn hợp gồm nước và tinh dầu tách lớp, rồi dùng phễu chiết để tách tinh dầu ra khỏi nước.")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 13 [SBT - Cánh Diều]. Acetic acid có công thức phân tử là C~2~H~4~O~2~. Kết luận nào sau đây là đúng?")
    opts_c13 = [
        "A. Acetic acid có công thức thực nghiệm là CH~2~O và có khối lượng riêng lớn hơn gấp 30 lần so với hydrogen ở cùng điều kiện (nhiệt độ áp suất).",
        "B. Acetic acid có công thức thực nghiệm là CH~2~O và có tỉ khối hơi so với hydrogen ở cùng điều kiện (nhiệt độ, áp suất) là 30.",
        "*C. Acetic acid có công thức thực nghiệm là CH~2~O và có phân tử khối là 60.*",
        "D. Acetic acid có công thức thực nghiệm là (CH~2~O)~2~ và có phân tử khối là 60."
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "C~2~H~4~O~2~ có tỉ lệ số nguyên tử C : H : O = 2 : 4 : 2 = 1 : 2 : 1 nên công thức thực nghiệm (công thức đơn giản nhất) là CH~2~O. Phân tử khối M = 12 \u00d7 2 + 4 \u00d7 1 + 16 \u00d7 2 = 60.")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 14. Trước kia, \"phẩm đỏ\" dùng để nhuộm áo choàng cho các Hồng y giáo chủ được tách chiết từ một loài ốc biển. Đó là một hợp chất có thành phần nguyên tố gồm: 45,7% C; 1,90% H; 7,60% O; 6,70% N và 38,10% Br. Công thức phân tử (CTPT) của phẩm đỏ là: (Biết bằng phương pháp phổ khối lượng, người ta xác định được rằng trong phân tử \"phẩm đỏ\" có hai nguyên tử Br):")
    opts_c14 = [
        "*A. C~16~H~8~O~2~N~2~Br~2~.*",
        "B. C~8~H~6~ONBr.",
        "C. C~6~H~8~ONBr.",
        "D. C~8~H~4~ONBr~2~."
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Ta có tỉ lệ: C : H : O : N : Br = (45,7/12) : (1,9/1) : (7,6/16) : (6,7/14) : (38,1/80) = 3,808 : 1,90 : 0,475 : 0,478 : 0,476 = 8 : 4 : 1 : 1 : 1.\nDo đó công thức đơn giản nhất là C~8~H~4~ONBr. Vì phân tử có 2 nguyên tử Br nên công thức phân tử là (C~8~H~4~ONBr)~2~ \u2192 C~16~H~8~O~2~N~2~Br~2~.")

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 15. Cho các chất sau: CH~3~OH (methanol), CH~3~CH~2~OH (a), CH~3~CH~2~CH~2~OH (b), (CH~3~)~2~CHOH (c), (CH~3~)~2~CHCH~2~OH (d), (CH~3~)~2~CHCH~2~CH~2~OH (e), (CH~3~)~3~COH (g), HOCH~2~CH~2~OH (h). Phát biểu nào sau đây là đúng?")
    opts_c15 = [
        "*A. Các chất (a), (b), (c), (d), (e), (g) thuộc dãy đồng đẳng của methanol.*",
        "B. Các chất (a), (b), (c), (d), (g), (h) thuộc dãy đồng đẳng của methanol.",
        "C. Các chất (a), (b), (d), (e), (g), (h) thuộc dãy đồng đẳng của methanol.",
        "D. Tất cả các chất đều thuộc dãy đồng đẳng của methanol."
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Methanol là alcohol no, đơn chức, mạch hở. Các chất (a), (b), (c), (d), (e), (g) đều là alcohol no, đơn chức, mạch hở có công thức chung C~n~H~2n+1~OH nên thuộc cùng dãy đồng đẳng của methanol. Chất (h) là alcohol hai chức (glycol) nên không cùng dãy đồng đẳng.")

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 16 [SBT - KNTT]. Cho các phát biểu sau:\n(1) Công thức cấu tạo biểu diễn kiểu liên kết và trật tự liên kết giữa các nguyên tử trong phân tử;\n(2) Chất đồng phân có cùng công thức phân tử nhưng có thể khác nhau về loại nhóm chức, mạch carbon, vị trí liên kết pi (\u03c0) hoặc vị trí nhóm chức;\n(3) Chất đồng đẳng có cấu tạo và tính chất tương tự, nhưng thành phần phân tử khác nhau một hay nhiều nhóm CH~2~.\nSố phát biểu đúng là")
    opts_c16 = [
        "A. 0.",
        "B. 1.",
        "C. 2.",
        "*D. 3.*"
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Cả 3 phát biểu (1), (2), (3) đều hoàn toàn chính xác theo định nghĩa về công thức cấu tạo, đồng phân và đồng đẳng.")

    # Câu 17
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 17. Ngâm rượu dược liệu: Cách tiến hành: Cho dược liệu vào trong lọ, bình hoặc hũ. Đổ một lượng rượu phù hợp rồi bịt kín lại đặt ở nơi tối, mát. Ngâm từ 10 \u2013 15 ngày, mùa đông có thể ngâm lâu hơn. Cho các phát biểu sau:\n(1) Ngâm dược liệu áp dụng phương pháp chiết lỏng \u2013 rắn.\n(2) Ngâm dược liệu áp dụng phương pháp chiết lỏng \u2013 lỏng.\n(3) Tách lấy chất hữu cơ ra khỏi một hỗn hợp ở thể rắn.\n(4) Tách lấy chất hữu cơ ra khỏi một hỗn hợp ở thể lỏng.\nSố phát biểu đúng là")
    opts_c17 = [
        "*A. 2.*",
        "B. 3.",
        "C. 4.",
        "D. 1."
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Ngâm rượu dược liệu là dùng dung môi ethanol hòa tan các hoạt chất hữu cơ có trong dược liệu thể rắn, đây là phương pháp chiết lỏng \u2013 rắn (tách lấy chất hữu cơ ra khỏi hỗn hợp ở thể rắn). Vậy có 2 phát biểu đúng là (1) và (3).")

    # Câu 18
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 18. Cho các phát biểu:\n(a) Để tách ethanol và nước người ta sử dụng phương pháp chưng cất.\n(b) Để tinh luyện đường đỏ thành đường trắng người ta sử dụng phương pháp kết tinh.\n(c) Không thể tách tinh dầu chanh bằng phương pháp chưng cất và chiết.\n(d) Phương pháp cột sắc kí thường dùng để tách các chất hữu cơ có hàm lượng nhỏ và khó tách.\n(e) Phương pháp kết tinh được dùng để tách và tinh chế các chất rắn.\nSố phát biểu đúng là")
    opts_c18 = [
        "*A. 4.*",
        "B. 3.",
        "C. 2.",
        "D. 1."
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Các phát biểu đúng gồm: (a), (b), (d), (e). Phát biểu (c) sai vì tinh dầu chanh hoàn toàn có thể tách được bằng phương pháp chưng cất lôi cuốn hơi nước kết hợp với chiết. Vậy có 4 phát biểu đúng.")

    # ==========================================
    # PHẦN II
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. Trong bài báo khoa học \u201cBioactive-Guided Phytochemical Investigations, In Vitro and In Silico Alpha-Glucosidase Inhibition of Two Vietnamese Medicinal Plants Dicranopteris linearis and Smilax glabra\u201d, các nhà khoa học đã phân lập được hợp chất có công thức cấu tạo như sau từ loài dương xỉ Dicranopteris linearis:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau1_p2_dicranopteris.png"), width_cm=6.5)
    
    stmts_c1_p2 = [
        ("a) Hợp chất trên chứa nhóm chức alcohol.", "SAI"),
        ("b) Hợp chất trên chứa nhóm chức carboxylic acid.", "SAI"),
        ("c) Hợp chất trên chứa nhóm chức ketone.", "ĐÚNG"),
        ("d) Hợp chất trên là hợp chất hữu cơ tổng hợp.", "SAI")
    ]
    for stmt, ans in stmts_c1_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c1_p2 = (
            "a) Sai: Các nhóm -OH đều liên kết trực tiếp với vòng thơm benzene nên thuộc loại nhóm chức phenol, không phải alcohol.\n"
            "b) Sai: Trong cấu trúc không có nhóm chức carboxyl (-COOH).\n"
            "c) Đúng: Nhóm C=O nằm trong vòng dị tố liên kết với hai nguyên tử carbon tạo thành nhóm chức ketone dạng cyclic enone.\n"
            "d) Sai: Hợp chất này được chiết xuất và phân lập từ thực vật tự nhiên (cây dương xỉ Dicranopteris linearis), là hợp chất thiên nhiên."
        )
        add_solution_no_table(doc, sol_c1_p2)

    # Câu 2 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Nấu rượu thủ công là phương pháp truyền thống đã có từ lâu đời tại Việt Nam:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau2_p2_nau_ruou.png"), width_cm=7.5)

    stmts_c2_p2 = [
        ("a) Rượu nấu xong chảy ra trước do nước có nhiệt độ sôi thấp hơn ethanol.", "SAI"),
        ("b) Ống tre dẫn hơi rượu đi qua chậu nước lạnh nhằm mục đích ngưng tụ hơi ethanol thành thể lỏng.", "ĐÚNG"),
        ("c) Phương pháp nấu rượu thủ công dựa trên sự khác nhau về nhiệt độ sôi của các chất trong hỗn hợp.", "ĐÚNG"),
        ("d) Phương pháp tách chất trong quá trình nấu rượu thủ công thuộc phương pháp chưng cất phân đoạn.", "SAI")
    ]
    for stmt, ans in stmts_c2_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c2_p2 = (
            "a) Sai: Ethanol có nhiệt độ sôi (78,3 \u00b0C) thấp hơn nước (100 \u00b0C) nên ethanol bay hơi trước và ngưng tụ chảy ra trước.\n"
            "b) Đúng: Hệ thống làm lạnh bằng nước giúp làm lạnh và ngưng tụ hơi ethanol cùng hơi nước thành chất lỏng (rượu).\n"
            "c) Đúng: Đây là nguyên lý của phương pháp chưng cất, tách các chất lỏng dựa vào độ bay hơi và nhiệt độ sôi khác nhau.\n"
            "d) Sai: Nấu rượu thủ công sử dụng thiết bị ngưng tụ đơn giản, thuộc phương pháp chưng cất thường (chưng cất đơn giản), không có tháp phân đoạn."
        )
        add_solution_no_table(doc, sol_c2_p2)

    # Câu 3 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Cho phổ khối lượng của hai hợp chất aniline (C~6~H~7~N) và 2-aminopyridine (C~5~H~6~N~2~) như sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau3_p2_ms_aniline_aminopyridine.png"), width_cm=11.0)

    stmts_c3_p2 = [
        ("a) Phổ khối lượng ở hình (a) tương ứng với phân tử aniline.", "ĐÚNG"),
        ("b) Mảnh ion phân tử ở hình (b) có giá trị m/z là 67.", "SAI"),
        ("c) Phổ khối lượng ở hình (b) tương ứng với phân tử 2-aminopyridine.", "ĐÚNG"),
        ("d) Phân tử khối của hai hợp chất hữu cơ A và B bằng nhau.", "SAI")
    ]
    for stmt, ans in stmts_c3_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c3_p2 = (
            "a) Đúng: Aniline có công thức phân tử C~6~H~7~N, phân tử khối M = 12\u00d76 + 7 + 14 = 93. Trên phổ (a), pic ion phân tử [M^+^] có m/z = 93, tương ứng với aniline.\n"
            "b) Sai: Mảnh ion phân tử [M^+^] ở hình (b) có giá trị m/z = 94. Đỉnh tại m/z = 67 là mảnh ion phân đoạn (fragment ion) có cường độ tương đối cao nhất.\n"
            "c) Đúng: 2-Aminopyridine có công thức C~5~H~6~N~2~, phân tử khối M = 12\u00d75 + 6 + 14\u00d72 = 94, ứng với đỉnh ion phân tử [M^+^] có m/z = 94 ở hình (b).\n"
            "d) Sai: Phân tử khối của aniline là 93 và của 2-aminopyridine là 94, hai giá trị phân tử khối này khác nhau."
        )
        add_solution_no_table(doc, sol_c3_p2)

    # Câu 4 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Citronellol là hợp chất được sử dụng tạo mùi hương tự nhiên có nguồn gốc từ các loại thực vật như hoa hồng, phong lữ hoặc sả, có công thức cấu tạo đầy đủ như sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau4_p2_citronellol_full.png"), width_cm=4.5)
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    render_rich_text(p, "Cho 3 công thức cấu tạo sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau4_p2_citronellol_ABC.png"), width_cm=13.5)

    stmts_c4_p2 = [
        ("a) Trong các công thức (A), (B), (C) có đúng 1 công thức biểu diễn đúng citronellol.", "ĐÚNG"),
        ("b) Citronellol có đồng phân hình học.", "SAI"),
        ("c) Citronellol có khả năng phản ứng với Na giải phóng khí H~2~.", "ĐÚNG"),
        ("d) Citronellol thuộc loại hợp chất hữu cơ đa chức.", "SAI")
    ]
    for stmt, ans in stmts_c4_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c4_p2 = (
            "a) Đúng: Chỉ có công thức (B) biểu diễn đúng cấu trúc mạch hở của citronellol: (CH~3~)~2~C=CH-CH~2~-CH~2~-CH(CH~3~)-CH~2~-CH~2~-OH. Công thức (A) là menthol (vòng no 6 cạnh), công thức (C) là geraniol (chứa 2 liên kết đôi C=C).\n"
            "b) Sai: Tại liên kết đôi C=C, nguyên tử carbon số 7 liên kết với hai nhóm methyl (-CH~3~) giống hệt nhau nên không thỏa mãn điều kiện có đồng phân hình học cis-trans.\n"
            "c) Đúng: Citronellol chứa nhóm chức alcohol bậc một (-OH) có nguyên tử hydrogen linh động, tác dụng được với kim loại Na giải phóng H~2~.\n"
            "d) Sai: Phân tử citronellol chỉ chứa duy nhất một nhóm chức alcohol (-OH), do đó thuộc loại hợp chất đơn chức."
        )
        add_solution_no_table(doc, sol_c4_p2)

    # ==========================================
    # PHẦN III
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 6."
    )

    # Câu 1 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. Cho các chất sau: CH~4~, C~2~H~2~, CCl~4~, C~2~H~4~, C~2~H~5~OH, CH~3~CHO. Có bao nhiêu chất là dẫn xuất của hydrocarbon?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: 3")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        add_solution_no_table(doc, "Hydrocarbon chỉ chứa hai nguyên tố C và H: CH~4~, C~2~H~2~, C~2~H~4~.\nDẫn xuất hydrocarbon ngoài C, H còn chứa các nguyên tố khác: CCl~4~, C~2~H~5~OH, CH~3~CHO \u2192 Có 3 chất.")

    # Câu 2 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Cho các đặc điểm sau của sắc kí cột:\n(1) Pha tĩnh là chất rắn có diện tích bề mặt lớn, có khả năng hấp phụ khác nhau các chất trong hỗn hợp.\n(2) Pha động là dung môi hữu cơ đi qua pha tĩnh mà không hòa tan các chất trong hỗn hợp.\n(3) Các chất trong hỗn hợp được tách ra khỏi cột theo thứ tự khả năng bị hấp phụ tăng dần.\n(4) Dung môi rửa giải chỉ được dùng một loại dung môi duy nhất trong suốt quá trình sắc kí.\n(5) Sắc kí cột dùng để tách và tinh chế các chất hữu cơ có hàm lượng nhỏ hoặc khó tách bằng các phương pháp khác.\nCó bao nhiêu đặc điểm đúng?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: 3")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        add_solution_no_table(doc, "Có 3 đặc điểm đúng là (1), (3), (5).\n- Đặc điểm (2) sai vì pha động phải hòa tan và lôi cuốn các chất trong hỗn hợp di chuyển qua pha tĩnh.\n- Đặc điểm (4) sai vì trong quá trình rửa giải có thể thay đổi dung môi hoặc tăng độ phân cực của dung môi để rửa giải các chất hấp phụ mạnh ra khỏi cột.")

    # Câu 3 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Cho các chất sau: C~2~H~4~O~2~, C~6~H~12~O~6~, C~2~H~6~O, HCHO, C~3~H~6~O~2~. Có bao nhiêu chất có cùng công thức đơn giản nhất là CH~2~O?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: 3")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        add_solution_no_table(doc, "Xét tỉ lệ C : H : O của từng chất:\n- C~2~H~4~O~2~: 2 : 4 : 2 = 1 : 2 : 1 \u2192 CH~2~O (thỏa mãn)\n- C~6~H~12~O~6~: 6 : 12 : 6 = 1 : 2 : 1 \u2192 CH~2~O (thỏa mãn)\n- C~2~H~6~O: 2 : 6 : 1 \u2192 C~2~H~6~O\n- HCHO: 1 : 2 : 1 \u2192 CH~2~O (thỏa mãn)\n- C~3~H~6~O~2~: 3 : 6 : 2 \u2192 C~3~H~6~O~2~\nVậy có 3 chất thỏa mãn.")

    # Câu 4 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Trong cà rốt chứa hàm lượng cao \u03b2-caroten (tiền chất vitamin A), có công thức khung phân tử như sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau4_p3_beta_caroten.png"), width_cm=11.5)
    render_rich_text(p, "\nSố nguyên tử carbon trong một phân tử \u03b2-caroten là bao nhiêu?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: 40 (hoặc 39 theo bản in gốc)")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        add_solution_no_table(doc, "\u03b2-Caroten là hợp chất tetraterpenoid đối xứng gồm 8 đơn vị isoprene (mỗi đơn vị 5C), công thức phân tử chuẩn là C~40~H~56~, do đó số nguyên tử carbon chính xác trong phân tử là 40.\n(Ghi chú: Trong đề in gốc có thể đếm thiếu 1 đỉnh carbon dẫn đến đáp số 39; đáp án chuẩn hóa học là 40).")

    # Câu 5 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Cây mía chứa khoảng 70% nước ép, nước ép chứa khoảng 20% saccharose (C~12~H~22~O~11~). Từ 1 tấn mía, lượng saccharose kết tinh thu được là bao nhiêu kg (biết hiệu suất kết tinh đạt 80%)?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: 112")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        add_solution_no_table(doc, "Khối lượng nước ép mía thu được từ 1 tấn (1000 kg) mía là: m_{nước ép} = 1000 \u00d7 70% = 700 kg.\nKhối lượng saccharose có trong nước ép là: m_{saccharose} = 700 \u00d7 20% = 140 kg.\nKhối lượng saccharose kết tinh thực tế thu được (H = 80%) là:\nm_{kết tinh} = 140 \u00d7 80% = 112 kg.")

    # Câu 6 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Acetone là dung môi hữu cơ quan trọng, thường dùng tẩy sơn móng tay. Phổ khối lượng của acetone cho thấy ion phân tử [M^+^] có m/z = 58 như hình vẽ:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau6_p3_ms_acetone.png"), width_cm=7.5)
    render_rich_text(p, "\nBiết công thức đơn giản nhất của acetone là C~3~H~6~O. Tổng số nguyên tử trong một phân tử acetone là bao nhiêu?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: 10")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        add_solution_no_table(doc, "Phổ khối lượng cho pic ion phân tử [M^+^] tại m/z = 58 \u21d2 Phân tử khối M = 58.\nCông thức phân tử có dạng (C~3~H~6~O)~n~ \u21d2 (12\u00d73 + 1\u00d76 + 16)\u00d7n = 58 \u21d2 58n = 58 \u21d2 n = 1.\nVậy CTPT của acetone là C~3~H~6~O. Tổng số nguyên tử = 3 (C) + 6 (H) + 1 (O) = 10.")

    # ==========================================
    # PHẦN IV
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN IV. Tự luận.",
        subtitle="Thí sinh trình bày lời giải chi tiết cho các câu hỏi sau."
    )

    # Câu 1 Phần IV
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1 (1,0 điểm). Cho các công thức khung cấu trúc sau đây:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau1_p4_5khung_cau_truc.png"), width_cm=11.5)
    render_rich_text(p, "\na) Hai công thức nào cùng biểu diễn một chất?\nb) Hai công thức nào là đồng phân cấu tạo vị trí nhóm chức của nhau?")

    if is_teacher:
        add_solution_no_table(doc, "a) Hai công thức cùng biểu diễn một chất:\n- Công thức (3) và (5) cùng biểu diễn hydrocarbon có mạch chính 4 carbon, có 1 liên kết đôi ở đầu mạch và 1 nhóm methyl ở C2: 2-methylbut-1-ene (CH~2~=C(CH~3~)-CH~2~-CH~3~).\nb) Hai công thức là đồng phân vị trí nhóm chức của nhau:\n- Công thức (2) và (4) đều có cùng công thức phân tử C~5~H~12~O và cùng khung carbon 2-methylbutane, nhưng khác nhau về vị trí của nhóm chức alcohol (-OH):\n  + Công thức (2): 3-methylbutan-2-ol (nhóm -OH ở carbon số 2).\n  + Công thức (4): 2-methylbutan-2-ol (nhóm -OH ở carbon số 2 nhánh methyl).\nDo đó, (2) và (4) là đồng phân vị trí nhóm chức của nhau.")
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Phần IV
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 (1,0 điểm). Một hợp chất hữu cơ X chứa 85,7% C và 14,3% H về khối lượng:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau2_p4_cau_truc_c4h8.png"), width_cm=6.5)
    render_rich_text(p, "\nPhổ khối lượng cho thấy phân tử khối của X là 56.\na) Xác định công thức thực nghiệm (công thức đơn giản nhất) của X.\nb) Xác định công thức phân tử của X.\nc) Viết công thức cấu tạo của X biết:\n   c~1~. X có mạch carbon không phân nhánh.\n   c~2~. X có mạch hở phân nhánh.")

    if is_teacher:
        add_solution_no_table(doc, "a) Đặt công thức đơn giản nhất của X là C~x~H~y~.\nTa có tỉ lệ số mol các nguyên tử: x : y = (%m~C~ / 12) : (%m~H~ / 1) = (85,7 / 12) : (14,3 / 1) = 7,14 : 14,3 = 1 : 2.\nVậy công thức thực nghiệm (công thức đơn giản nhất) của X là CH~2~.\n\nb) Công thức phân tử của X có dạng (CH~2~)~n~.\nTheo phổ khối lượng, phân tử khối M = 56 \u21d2 14n = 56 \u21d2 n = 4.\nVậy công thức phân tử của X là C~4~H~8~.\n\nc) Các công thức cấu tạo thỏa mãn:\n- Trường hợp c~1~: Mạch carbon không phân nhánh:\n  + Alkene mạch hở: CH~2~=CH-CH~2~-CH~3~ (but-1-ene) hoặc CH~3~-CH=CH-CH~3~ (but-2-ene, gồm đồng phân hình học cis và trans).\n  + Cycloalkane vòng không phân nhánh: cyclobutane (vòng no 4 cạnh).\n- Trường hợp c~2~: Mạch hở phân nhánh:\n  + 2-methylpropene (isobutylene): CH~2~=C(CH~3~)~2~ hay (CH~3~)~2~C=CH~2~.")
    else:
        add_student_writing_lines(doc, num_lines=10)

    # Lưu tài liệu
    fname = "De_Kiem_Tra_Hoa_11_De09_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_11_De09.docx"
    fpath = os.path.join(OUT_DIR, fname)
    doc.save(fpath)
    print(f"[\u2713] Đã tạo thành công: {fpath}")

def main():
    print("--- Bắt đầu tạo 2 file Word Đề số 9 Hóa học 11 ---")
    build_exam_document(is_teacher=False)
    build_exam_document(is_teacher=True)
    print("--- Hoàn thành tạo 2 file Word Đề số 9! ---")

if __name__ == "__main__":
    main()
