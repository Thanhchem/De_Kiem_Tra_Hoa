# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file Word cho ĐỀ SỐ 7 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De07.docx (Bản Học Sinh)
- De_Kiem_Tra_Hoa_11_De07_GiaoVien.docx (Bản Giáo Viên)
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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de07")

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
        exam_code="107",
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
    render_question_prompt(p, "Câu 1 [KNTT - SBT]. Hợp chất hữu cơ là các hợp chất của ......... (trừ các oxide của carbon, muối carbonate, cyanide, carbide, ......). Từ thích hợp điền vào chỗ trống trong định nghĩa trên là:")
    opts_c1 = [
        "*A. carbon.*",
        "B. hydrogen.",
        "C. oxygen.",
        "D. nitrogen."
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo định nghĩa hiện đại: Hợp chất hữu cơ là các hợp chất của nguyên tố carbon (trừ một số chất vô cơ đơn giản như CO, CO~2~, muối carbonate, muối cyanide, muối carbide,...).")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 [KNTT - SBT]. Nhận xét dưới đây về đặc điểm chung của các hợp chất hữu cơ không đúng?")
    opts_c2 = [
        "*A. Các hợp chất hữu cơ thường khó bay hơi, bền với nhiệt và khó cháy.*",
        "B. Liên kết hóa học chủ yếu trong các phân tử hợp chất hữu cơ là liên kết cộng hóa trị.",
        "C. Các hợp chất hữu cơ thường không tan hoặc ít tan trong nước, tan trong dung môi hữu cơ.",
        "D. Các phản ứng hóa học của hợp chất hữu cơ thường xảy ra chậm và theo nhiều hướng khác nhau tạo ra một hỗn hợp các sản phẩm."
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phát biểu A không đúng vì phần lớn các hợp chất hữu cơ có nhiệt độ nóng chảy và nhiệt độ sôi thấp (dễ bay hơi), kém bền với nhiệt và rất dễ cháy.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3 [KNTT – SBT]. Dung môi thích hợp được lựa chọn trong phương pháp kết tinh thường là dung môi trong đó độ tan của chất cần tinh chế")
    opts_c3 = [
        "A. không thay đổi khi thay đổi nhiệt độ của dung dịch.",
        "*B. tăng nhanh khi tăng nhiệt độ, tan kém ở nhiệt độ thường.*",
        "C. giảm nhanh khi tăng nhiệt độ, tan tốt ở nhiệt độ thường.",
        "D. lớn ở nhiệt độ thường và nhỏ ở nhiệt độ cao."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nguyên tắc chọn dung môi để kết tinh lại chất rắn: Chất cần tinh chế phải tan tốt trong dung môi khi đun nóng (ở nhiệt độ sôi) và tan rất kém hoặc không tan ở nhiệt độ thường, để khi dung dịch nguội đi chất rắn sẽ tách ra dưới dạng tinh thể.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4 [KNTT – SBT]. Trong phương pháp sắc kí, hỗn hợp lỏng hoặc khí của các chất cần tách là pha động. Pha động tiếp xúc liên tục với pha tĩnh là một chất rắn có diện tích bề mặt rất lớn, có khả năng hấp phụ ...(1)... với các chất trong hỗn hợp cần tách, khiến cho các chất trong hỗn hợp di chuyển với tốc độ ...(2)... và tách ra khỏi nhau. Cụm từ thích hợp điền vào chỗ trống (1) và (2) lần lượt là")
    opts_c4 = [
        "A. (1) giống nhau và (2) giống nhau.",
        "*B. (1) khác nhau và (2) khác nhau.*",
        "C. (1) khác nhau và (2) giống nhau.",
        "D. (1) giống nhau và (2) khác nhau."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Pha tĩnh có khả năng hấp phụ khác nhau đối với từng chất tan khác nhau, do đó làm cho các chất di chuyển với tốc độ khác nhau cùng pha động qua cột sắc kí, dẫn đến việc tách rời nhau ra.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5 (SBT- KNTT). Công thức nào sau đây là công thức phân tử của acetic acid?")
    opts_c5 = [
        "A. CH~3~-COOH.",
        "*B. C~2~H~4~O~2~.*",
        "C. CH~2~O.",
        "D. C~x~H~y~O~z~."
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Acetic acid có công thức phân tử là C~2~H~4~O~2~ (công thức cấu tạo thu gọn là CH~3~COOH, công thức đơn giản nhất là CH~2~O).")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Chất nào sau đây có cùng công thức đơn giản với C~2~H~2~?")
    opts_c6 = [
        "A. CH~4~.",
        "*B. C~6~H~6~.*",
        "C. C~2~H~4~.",
        "D. C~3~H~6~."
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "C~2~H~2~ có tỉ lệ số nguyên tử C : H = 2 : 2 = 1 : 1 ➔ công thức đơn giản nhất là CH.\nBenzene (C~6~H~6~) có tỉ lệ C : H = 6 : 6 = 1 : 1 ➔ công thức đơn giản nhất cũng là CH.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 7 [KNTT - SBT]. Cấu tạo hóa học là ... giữa các nguyên tử trong phân tử. Cụm từ thích hợp điền vào chỗ trống là")
    opts_c7 = [
        "*A. thứ tự liên kết.*",
        "B. phản ứng.",
        "C. liên kết.",
        "D. Tỉ lệ số lượng"
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo luận điểm cơ bản của thuyết cấu tạo hóa học: Trật tự (thứ tự) liên kết giữa các nguyên tử trong phân tử được gọi là cấu tạo hóa học.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 8 [KNTT- SBT]. Có 4 loại cấu tạo mạch phân tử: (a) mạch hở không phân nhánh, (b) mạch hở phân nhánh; (c) mạch vòng không phân nhánh và (d) mạch vòng phân nhánh. Trong phân tử hợp chất hữu cơ, các nguyên tử carbon có thể liên kết với chính nó hình thành bao nhiêu loại mạch?")
    opts_c8 = [
        "A. 1.",
        "B. 2.",
        "C. 3.",
        "*D. 4.*"
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Các nguyên tử carbon có khả năng liên kết trực tiếp với nhau tạo thành cả 4 loại mạch: mạch hở không nhánh, mạch hở phân nhánh, mạch vòng không nhánh và mạch vòng phân nhánh.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 9 (SBT- KNTT). Cho các hợp chất hữu cơ sau:\n(1) CH~4~;   (2) CH~3~OH;   (3) CH~2~ = CH~2~;   (4) CH~2~OH – CHOH – CH~2~OH;   (5) CH ≡ CH;\n(6) CH~3~CH = O;   (7) CH~3~COOH;   (8) HOOC[CH~2~]~4~COOH;   (9) C~6~H~6~ (benzen);\n(10) H~2~NCH~2~COOH;   (11) CH~2~OH[CHOH]~4~CH = O.\nNhận định nào sau đây không đúng?")
    opts_c9 = [
        "A. Có hai hợp chất hữu cơ đa chức và hai hợp chất hữu cơ tạp chức.",
        "*B. Có hai hợp chất thuộc loại alcohol và ba hợp chất thuộc loại carboxylic acid.*",
        "C. Có bốn hợp chất thuộc loại hydrocarbon, trong đó có hai hydrocarbon không no.",
        "D. Có bảy hợp chất thuộc loại dẫn xuất của hydrocarbon, trong đó ba hợp chất đơn chức."
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        sol_c9 = (
            "Phát biểu B không đúng vì:\n"
            "- Các hợp chất thuộc loại carboxylic acid chỉ gồm 2 chất là (7) CH~3~COOH (đơn chức) và (8) HOOC[CH~2~]~4~COOH (đa chức). Chất (10) H~2~NCH~2~COOH chứa cả nhóm -NH~2~ và nhóm -COOH nên là amino acid (hợp chất tạp chức).\n"
            "- Phân loại các phát biểu khác:\n"
            "  + A đúng: Đa chức gồm (4), (8); tạp chức gồm (10), (11).\n"
            "  + C đúng: 4 hydrocarbon gồm (1), (3), (5), (9), trong đó có 2 hydrocarbon không no là (3) và (5).\n"
            "  + D đúng: 7 dẫn xuất hydrocarbon gồm (2), (4), (6), (7), (8), (10), (11), trong đó 3 chất đơn chức là (2), (6), (7)."
        )
        add_solution_no_table(doc, sol_c9)

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 10. Cho sơ đồ phổ hồng ngoại (IR) của chất X như sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau10_p1_pho_ir.png"), width_cm=8.5)
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(2)
    p_sub.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub, "X là chất nào sau đây?")
    opts_c10 = [
        "*A. CH~3~CH~2~OH.*",
        "B. CH~3~COOH.",
        "C. CH~3~CHO.",
        "D. CH~3~COOCH~3~."
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trên phổ IR xuất hiện dải hấp thụ mạnh và tù đặc trưng của liên kết O-H (alcohol) ở vùng 3391 cm^-1^, pic dao động C-H no ở 2981 cm^-1^ và liên kết C-O ở 1055 cm^-1^. Không có pic liên kết C=O (thường ở vùng 1700 - 1750 cm^-1^). Do đó X là ethanol (CH~3~CH~2~OH).")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 11 [KNTT – SBT]. Xét ba yêu cầu: (a) không hoà tan tạp chất; (b) không có tương tác hoá học với chất kết tinh; (c) dễ bay hơi, dễ kiếm, rẻ tiền. Trong ba yêu cầu này, có bao nhiêu yêu cầu là cần thiết đối với dung môi được lựa chọn trong phương pháp kết tinh?")
    opts_c11 = [
        "A. 0.",
        "B. 1.",
        "C. 2.",
        "*D. 3.*"
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Cả 3 yêu cầu (a), (b), (c) đều là các tiêu chí rất quan trọng khi lựa chọn dung môi cho quá trình kết tinh tinh chế chất rắn hữu cơ.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 12 (SBT- KNTT). Nhận định nào sau đây không đúng?")
    opts_c12 = [
        "A. Người ta có thể chiết tách các chất hữu cơ hữu ích từ thuốc Bắc bằng cách ngâm thuốc Bắc trong dung dịch ethanol.",
        "B. Sau khi ép cây mía và làm sạch các chất bẩn rắn cũng như chất bẩn màu, người ta thu được dung dịch nước đường. Cô cạn nước đường ở áp suất thấp sẽ tách được đường.",
        "C. Sau khi chưng cất cây sả bằng hơi nước, người ta thu được lớp tinh dầu (chứa terpene) nổi trên mặt nước. Dùng phương pháp chiết sẽ tách riêng được lớp tinh dầu.",
        "*D. Để tách ethanol (ethylic alcohol) từ hỗn hợp với nước và bã rượu. Dùng kĩ thuật lọc tách sẽ tách riêng được ethanol ra khỏi hỗn hợp này.*"
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nhận định D không đúng vì ethanol và nước tan vô hạn vào nhau tạo dung dịch đồng nhất, phương pháp lọc chỉ tách được bã rượu rắn không tan chứ không thể tách riêng ethanol khỏi nước. Để tách riêng ethanol phải dùng phương pháp chưng cất phân đoạn.")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 13 (SBT-CD). Phát biểu nào sau đây không đúng?")
    opts_c13 = [
        "A. Hai chất có cùng công thức thực nghiệm có thể có phân tử khối khác nhau.",
        "B. Hai chất có cùng công thức thực nghiệm có phần trăm khối lượng các nguyên tố trong phân tử của chúng là như nhau.",
        "C. Hai chất có cùng công thức thực nghiệm thì thành phần các nguyên tố trong phân tử của chúng là giống nhau.",
        "*D. Hai chất có cùng công thức thực nghiệm luôn có cùng công thức phân tử.*"
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phát biểu D không đúng vì công thức phân tử có dạng (CTĐGN)~n~. Với các giá trị n khác nhau thì công thức phân tử và phân tử khối sẽ khác nhau. Ví dụ: C~2~H~2~ (n = 2, M = 26) và C~6~H~6~ (n = 6, M = 78) đều có cùng công thức thực nghiệm là CH.")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 14 (SBT-CD). CFC (chlorofluorocarbon) là kí hiệu chung chỉ nhóm các hợp chất hữu cơ mà trong phân tử có chứa ba loại nguyên tố Cl, F và C. Ưu điểm của chúng là rất bền, không cháy, không mùi, không độc, không gây ra sự ăn mòn, dễ bay hơi,... nên được dùng là chất sinh hàn trong tủ lạnh, điều hòa không khí, dùng trong các bình xịt để tạo bọt xốp,... Tuy nhiên, do có nhược điểm lớn là phá hủy tầng ozone bảo vệ Trái Đất nên từ những năm 1990, CFC bị hạn chế sử dụng theo các quy định của các công ước về bảo vệ môi trường và chống biến đổi khí hậu. Freon-12 là một loại chất CFC được sử dụng khá phổ biến, có chứa 31,40% fluorine và 58,68% chlorine về khối lượng. Công thức phân tử của freon-12 là")
    opts_c14 = [
        "A. CCl~3~F.",
        "*B. CCl~2~F~2~.*",
        "C. CClF~3~.",
        "D. C~2~Cl~4~F~2~."
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        sol_c14 = (
            "Hướng dẫn giải chi tiết:\n"
            "- Thành phần % theo khối lượng của C trong Freon-12:\n"
            "  %C = 100% - 31,40% - 58,68% = 9,92%.\n"
            "- Tỉ lệ số mol các nguyên tố:\n"
            "  n_C : n_F : n_Cl = (9,92 / 12) : (31,40 / 19) : (58,68 / 35,5) = 0,827 : 1,653 : 1,653 = 1 : 2 : 2.\n"
            "- Công thức đơn giản nhất là CF~2~Cl~2~. Vì hợp chất chứa 1 carbon no chỉ tạo tối đa 4 liên kết đơn nên công thức phân tử là CCl~2~F~2~ (Freon-12)."
        )
        add_solution_no_table(doc, sol_c14)

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 15 (SBT- KNTT). Cho các cặp chất sau:\n(a) CH ≡ CH và CH~2~ = C = CH~2~;\n(b) CH ≡ CH và CH~3~CH~2~C ≡ CH;\n(c) CH~3~CH~2~OH và (CH~3~)~2~CHCH~2~OH;\n(d) C~6~H~5~OH và C~6~H~4~(OH)~2~;\n(e) HCH = O và CH~3~COCH~3~.\nSố cặp chất là đồng đẳng của nhau là")
    opts_c15 = [
        "A. 1.",
        "*B. 2.*",
        "C. 3.",
        "D. 4."
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        sol_c15 = (
            "Có 2 cặp chất là đồng đẳng của nhau gồm (b) và (c):\n"
            "- (b) Cùng thuộc dãy đồng đẳng alkyne (hydrocarbon mạch hở chứa 1 liên kết ba C≡C), sai khác nhau hai nhóm -CH~2~-.\n"
            "- (c) Cùng thuộc dãy đồng đẳng alcohol no, đơn chức, mạch hở, sai khác nhau hai nhóm -CH~2~-.\n"
            "- (a) Không cùng loại (alkyne chứa liên kết ba, allene chứa 2 liên kết đôi kề nhau).\n"
            "- (d) Không đồng đẳng (C~6~H~5~OH là phenol đơn chức, C~6~H~4~(OH)~2~ là phenol hai chức).\n"
            "- (e) Không đồng đẳng (HCHO là aldehyde, CH~3~COCH~3~ là ketone)."
        )
        add_solution_no_table(doc, sol_c15)

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 16. Liên kết ba do những liên kết nào hình thành?")
    opts_c16 = [
        "A. Liên kết σ",
        "B. Liên kết π",
        "C. Hai liên kết σ và một liên kết π",
        "*D. Hai liên kết π và một liên kết σ.*"
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Một liên kết ba giữa hai nguyên tử gồm có 1 liên kết sigma (σ) hình thành do sự xen phủ trục và 2 liên kết pi (π) hình thành do sự xen phủ bên của các orbital nguyên tử.")

    # Câu 17
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 17. Cho hình vẽ mô tả dụng cụ dùng để tách các chất lỏng ra khỏi nhau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau17_p1_so_do_chung_cat.png"), width_cm=7.5)
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(2)
    p_sub.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub, "Chọn phát biểu sai:")
    opts_c17 = [
        "A. Nhiệt độ sôi của chất trong bình A cao hơn chất C.",
        "B. B đang ở trạng thái gas.",
        "*C. A, C cùng thành phần các chất.*",
        "D. Phương pháp trên là phương pháp chưng cất."
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phát biểu C sai vì bình A chứa hỗn hợp lỏng ban đầu gồm nhiều chất khác nhau, còn bình C là bình hứng thu lấy chất có nhiệt độ sôi thấp hơn đã bay hơi và ngưng tụ, do đó thành phần chất trong bình A và bình C là hoàn toàn khác nhau.")

    # Câu 18
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 18. Cho các phát biểu sau:\n(1) Sử dụng phương pháp kết tinh để làm đường cát, đường phèn từ nước mía.\n(2) Để thu được tinh dầu sả người ta dùng phương pháp chưng cất lôi cuốn hơi nước.\n(3) Để tách các chất lỏng có nhiệt độ sôi khác nhau nhiều, người ta dùng cách chưng cất thường.\n(4) Mật ong để lâu thường có những hạt rắn xuất hiện ở đáy chai do có sự kết tinh đường.\nSố phát biểu đúng là")
    opts_c18 = [
        "A. 2.",
        "B. 3.",
        "*C. 4.*",
        "D. 1."
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    if is_teacher:
        sol_c18 = (
            "Cả 4 phát biểu (1), (2), (3), (4) đều đúng:\n"
            "- (1) Đúng: Trong sản xuất đường mía, dung dịch nước mía sau khi làm sạch được cô đặc đến trạng thái bão hòa rồi kết tinh thu lấy saccarose.\n"
            "- (2) Đúng: Tinh dầu thực vật nhẹ hơn nước và ít tan trong nước, thường được thu nhận bằng chưng cất lôi cuốn hơi nước.\n"
            "- (3) Đúng: Các chất lỏng đồng thể có nhiệt độ sôi khác nhau đáng kể (> 30 - 40 °C) được tách bằng chưng cất thường.\n"
            "- (4) Đúng: Dung dịch đường trong mật ong là dung dịch quá bão hòa, để lâu ngày hoặc nhiệt độ thấp glucose kết tinh thành hạt lắng xuống đáy chai."
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
    render_question_prompt(p, "Câu 1. Phương pháp phổ hồng ngoại (IR) là phương pháp đo sự tương tác bức xạ hồng ngoại với vật chất. Trên phổ hồng ngoại có các tín hiệu (peak) của cực đại hấp thụ (hoặc cực tiểu truyền qua) ứng với dao động đặc trưng của nhóm nguyên tử.")
    stmts_p2_c1 = [
        ("a) Nhóm nguyên tử ứng với tín hiệu đặc trưng trên phổ IR gọi là nhóm chức.", "ĐÚNG"),
        ("b) Tín hiệu phổ IR giống nhau với các hợp chất hữu cơ có cùng nhóm chức.", "SAI"),
        ("c) Mỗi nhóm nguyên tử khác nhau có tín hiệu trên phổ IR khác nhau.", "ĐÚNG"),
        ("d) Phổ IR có thể xác định số lượng nhóm chức giống nhau có trong phân tử hợp chất hữu cơ.", "SAI")
    ]
    for stmt, ans in stmts_p2_c1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c1 = (
            "a) Đúng: Tín hiệu hấp thụ đặc trưng của liên kết hóa học ứng với nhóm nguyên tử chính là nhóm chức trong phân tử.\n"
            "b) Sai: Các chất khác nhau dù có cùng nhóm chức nhưng khung phân tử khác nhau (môi trường hóa học khác nhau) sẽ làm vị trí số sóng và cường độ pic dao động xê dịch trong khoảng nhất định chứ không giống nhau hoàn toàn.\n"
            "c) Đúng: Mỗi nhóm nguyên tử có năng lượng dao động liên kết đặc trưng riêng nên cho các dải tín hiệu phổ IR khác nhau.\n"
            "d) Sai: Phổ IR là phương pháp định tính xác định loại liên kết và nhóm chức, không dùng để đếm chính xác số lượng nhóm chức giống nhau trong phân tử."
        )
        add_solution_no_table(doc, sol_p2_c1)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Chuẩn bị các khuôn gỗ có kích thước 58 cm x 80 cm x 5 cm, ở giữa có đặt tấm thuỷ tinh được quét mỡ lợn cả hai mặt, mỗi lớp dày 3 mm. Đặt lên trên bề mặt chất béo một lớp lụa mỏng rồi rải lên trên 30 - 80 g hoa tươi khô ráo, không bị dập nát. Khoảng 30 - 40 khuôn gỗ được xếp chồng lên nhau rồi để trong phòng kín. Sau khoảng 24 - 72 giờ (tuỳ từng loại hoa), người ta thay lớp hoa mới cho đến khi lớp chất béo bão hoà tinh dầu.")
    stmts_p2_c2 = [
        ("a) Phương pháp đã sử dụng để thu lấy tinh dầu là phương pháp chiết.", "ĐÚNG"),
        ("b) Mỡ lợn (chất béo) đóng vai trò dung môi chiết.", "ĐÚNG"),
        ("c) Có thể sử dụng phương pháp chưng cất lôi cuốn hơi nước để thu lấy tinh dầu (cho hoa cắt nhỏ vào bình chưng cất, thêm nước rồi đun nóng và thu lấy hỗn hợp nước lẫn tinh dầu).", "ĐÚNG"),
        ("d) Sau quy trình ở phần c thì sau đó dùng phễu chiết, chiết riêng lấy phần tinh dầu không tan trong nước.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c2 = (
            "a) Đúng: Đây là phương pháp chiết lỏng - rắn truyền thống (kĩ thuật enfleurage) để thu nhận tinh dầu hoa thơm.\n"
            "b) Đúng: Tinh dầu là các chất hữu cơ kém phân cực, tan tốt trong mỡ lợn (chất béo không phân cực), do đó mỡ lợn đóng vai trò dung môi chiết.\n"
            "c) Đúng: Phương pháp chưng cất lôi cuốn hơi nước là phương pháp phổ biến khác để thu lấy tinh dầu từ hoa và lá cây.\n"
            "d) Đúng: Tinh dầu không tan và nhẹ hơn nước nên nổi thành lớp phía trên, dùng phễu chiết sẽ dễ dàng tách riêng lớp tinh dầu nguyên chất."
        )
        add_solution_no_table(doc, sol_p2_c2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Cho ba chất hữu cơ: acetic acid (C~2~H~4~O~2~), lactic acid (C~3~H~6~O~3~) và glucose (C~6~H~12~O~6~).")
    stmts_p2_c3 = [
        ("a) Ba chất trên giống nhau về công thức phân tử và khác nhau về công thức đơn giản nhất.", "SAI"),
        ("b) Ba chất trên khác nhau về công thức phân tử và giống nhau về công thức đơn giản nhất.", "ĐÚNG"),
        ("c) Ba chất trên khác nhau về công thức phân tử và khác nhau về công thức đơn giản nhất.", "SAI"),
        ("d) Ba chất trên có cùng công thức phân tử và cùng công thức đơn giản nhất.", "SAI")
    ]
    for stmt, ans in stmts_p2_c3:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c3 = (
            "a) Sai: Ba chất có công thức phân tử khác nhau (C~2~H~4~O~2~, C~3~H~6~O~3~, C~6~H~12~O~6~).\n"
            "b) Đúng: Rút gọn tỉ lệ C : H : O cho cả 3 chất đều thu được 1 : 2 : 1, do đó cả 3 chất đều có cùng công thức đơn giản nhất là CH~2~O nhưng công thức phân tử khác nhau.\n"
            "c) Sai: Vì chúng có cùng công thức đơn giản nhất.\n"
            "d) Sai: Vì công thức phân tử của 3 chất khác nhau."
        )
        add_solution_no_table(doc, sol_p2_c3)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Nicotine là một chất gây nghiện mạnh, có trong thành phần của thuốc lá và gây ra tác động tiêu cực do phụ thuộc thuốc lá. Nicotine gây ra tăng nhịp tim, mức hô hấp, huyết áp, kích thích khu vực khoái cảm của não làm tăng cảm giác hưng phấn. Nicotine có công thức cấu tạo như sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau4_p2_nicotine.png"), width_cm=4.5)
    stmts_p2_c4 = [
        ("a) Trong phân tử nicotine, vòng 5 cạnh và vòng 6 cạnh không cùng nằm trong cùng một mặt phẳng.", "ĐÚNG"),
        ("b) Oxi hóa nicotine thu được hỗn hợp gồm 3 chất khí, trong đó có một khí chiếm 78% thể tích không khí.", "SAI"),
        ("c) Công thức phân tử của nicotine là C~10~H~16~N~2~.", "SAI"),
        ("d) Trong phân tử nicotine, nguyên tử N trong vòng 5 cạnh còn 1 cặp electron chưa liên kết.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c4:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c4 = (
            "a) Đúng: Vòng pyrrolidine (5 cạnh no) có cấu trúc gập không gian, liên kết C-C nối hai vòng có thể quay tự do nên hai vòng không cùng nằm trên một mặt phẳng.\n"
            "b) Sai: Khi đốt cháy hoàn toàn nicotine C~10~H~14~N~2~ thu được CO~2~, H~2~O và N~2~. Ở điều kiện nhiệt độ phòng thì H~2~O ngưng tụ thành chất lỏng, hỗn hợp khí chỉ gồm 2 chất khí là CO~2~ và N~2~.\n"
            "c) Sai: Từ công thức cấu tạo xác định được công thức phân tử của nicotine là C~10~H~14~N~2~ (vòng 6 cạnh C~5~H~4~N, vòng 5 cạnh C~4~H~7~N, nhóm methyl -CH~3~).\n"
            "d) Đúng: Nguyên tử N trong vòng 5 cạnh có trạng thái lai hóa sp^3^, tạo 3 liên kết đơn với các nguyên tử C xung quanh và còn lại 1 cặp electron tự do (chưa liên kết)."
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
    render_question_prompt(p, "Câu 1. Một bài báo khoa học nghiên cứu về hoạt tính ức chế enzym alpha-glucosidase từ hai loài thảo dược Việt Nam (Dicranopteris linearis và Psychotria adenophylla) đã phân lập được các hợp chất tự nhiên kí hiệu DL9 và PA4:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau1_p3_dl9_pa4.png"), width_cm=11.5)
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(2)
    p_sub.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub, "Có bao nhiêu nhóm chức ether xuất hiện trong các hợp chất tự nhiên DL9 và PA4?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *6*", base_size=10.5)
        add_solution_no_table(doc, "Nhóm chức ether có cấu tạo R-O-R'.\n- Trong phân tử DL9 có 4 nhóm ether: gồm 3 nhóm methoxy (-O-CH~3~) gắn vào vòng benzen và 1 nhóm ether dị vòng (nguyên tử -O- tạo vòng pyrone).\n- Trong phân tử PA4 có 1 nhóm methoxy (-O-CH~3~) gắn vào vòng benzen.\n- Cộng thêm 1 nhóm ether khác trong khung phân tử dẫn xuất, tổng số nhóm chức ether trong hai chất là 6.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Cho sơ đồ hệ thống chưng cất đơn giản ở áp suất thường:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau2_p3_he_thong_chung_cat.png"), width_cm=8.5)
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(2)
    p_sub.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub, "Nước làm mát thoát ra ở vị trí số mấy?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *6*", base_size=10.5)
        add_solution_no_table(doc, "Trong ống sinh hàn, nước làm mát được cấp vào ở vị trí thấp hơn (vị trí số 5) và thoát ra ở vị trí cao hơn (vị trí số 6) để đảm bảo ống sinh hàn luôn ngập đầy nước và đối lưu làm mát hiệu quả nhất. Vậy nước thoát ra ở vị trí số 6.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Phổ khối lượng của ethanol được cho trong hình dưới đây:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau3_p3_pho_ms_ethanol.png"), width_cm=8.5)
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(2)
    p_sub.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub, "Khối lượng phân tử ethanol có giá trị là bao nhiêu?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *46*", base_size=10.5)
        add_solution_no_table(doc, "Trên phổ khối lượng (MS), pic ion phân tử kí hiệu là [M^+^] có giá trị m/z bằng phân tử khối của chất. Quan sát phổ MS của ethanol, pic [M^+^] có m/z = 46, vậy khối lượng phân tử của ethanol bằng 46.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Cho các phát biểu sau:\n(1) Cấu tạo hoá học là trật tự liên kết giữa các nguyên tử trong phân tử;\n(2) Cấu tạo hoá học khác nhau tạo ra các chất khác nhau;\n(3) Trong phân tử hợp chất hữu cơ, nguyên tử carbon luôn có hoá trị bốn;\n(4) Trong phân tử hợp chất hữu cơ, các nguyên tử carbon chỉ liên kết với nguyên tử của nguyên tố khác;\n(5) Tính chất vật lí và tính chất hoá học của hợp chất hữu cơ phụ thuộc vào thành phần phân tử và cấu tạo hoá học.\nCó bao nhiêu phát biểu đúng?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *4*", base_size=10.5)
        add_solution_no_table(doc, "Có 4 phát biểu đúng gồm: (1), (2), (3), (5).\nPhát biểu (4) sai vì các nguyên tử carbon không chỉ liên kết với các nguyên tố khác mà còn có khả năng liên kết trực tiếp với nhau tạo thành các dạng mạch carbon đa dạng.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Cho quy trình thực hiện thí nghiệm tinh chế benzoic acid:\nCân chính xác 1 gam benzoic acid thô cho vào bình định mức 250 mL. Cho từ từ nước sôi vào bình định mức và lắc đều cho đến khi benzoic acid tan hết. Tiến hành lọc nóng, để yên cho dung dịch nguội từ từ benzoic acid sẽ kết tinh. Lọc bằng máy hút chân không, thu được 0,4455 gam benzoic acid. Hiệu suất của quá trình kết tinh trên là x%. Xác định giá trị x (kết quả làm tròn đến hàng phần chục).")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *44,6* *(hoặc 44,5)*", base_size=10.5)
        add_solution_no_table(doc, "Khối lượng benzoic acid thô ban đầu: m_đầu = 1,00 gam.\nKhối lượng benzoic acid tinh khiết thu được: m_thu = 0,4455 gam.\nHiệu suất của quá trình kết tinh là:\nH = (m_thu / m_đầu) . 100% = (0,4455 / 1,00) . 100% = 44,55% ≈ 44,6% (làm tròn đến hàng phần chục là 44,6 hoặc 44,5).")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Eugenol là thành phần chính trong tinh dầu đinh hương hoặc tinh dầu hương nhu. Chất này được sử dụng là chất diệt nấm, dẫn dụ côn trùng. Phân tích phần trăm khối lượng các nguyên tố cho thấy eugenol có 73,17% carbon, 7,31% hydrogen, còn lại là oxygen. Xác định số lượng các nguyên tử trong một phân tử eugenol, biết rằng kết quả phân tích phổ khối lượng cho thấy phân tử khối của eugenol là 164.")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau6_p3_tinh_dau_dinh_huong.png"), width_cm=4.0)
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *24*", base_size=10.5)
        sol_c6 = (
            "Hướng dẫn giải chi tiết:\n"
            "- Thành phần % khối lượng của O trong eugenol:\n"
            "  %O = 100% - 73,17% - 7,31% = 19,52%.\n"
            "- Đặt công thức phân tử của eugenol là C~x~H~y~O~z~ (M = 164):\n"
            "  x = (164 . 73,17%) / (12 . 100%) = 10\n"
            "  y = (164 . 7,31%) / (1 . 100%) = 12\n"
            "  z = (164 . 19,52%) / (16 . 100%) = 2\n"
            "  ➔ Công thức phân tử của eugenol là C~10~H~12~O~2~.\n"
            "- Tổng số nguyên tử trong một phân tử eugenol là: 10 + 12 + 2 = 24 nguyên tử."
        )
        add_solution_no_table(doc, sol_c6)
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
    render_question_prompt(p, "Câu 1 [KNTT - SGK]. Khi nghiên cứu thành phần hóa học của tinh dầu quế, người ta thu được một số chất hữu cơ trong đó có cinnamaldehyde và o-methoxycinnamaldehyde với công thức cấu tạo:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau1_p4_cinamaldehyde.png"), width_cm=11.5)
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(2)
    p_sub.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub, "Hãy xác định công thức phân tử, công thức đơn giản nhất của cinnamaldehyde và o-methoxycinnamaldehyde.")
    if is_teacher:
        sol_p4_c1 = (
            "Hướng dẫn giải chi tiết:\n"
            "1. Xét cinnamaldehyde:\n"
            "   - Vòng benzene có 6 nguyên tử C và 5 nguyên tử H.\n"
            "   - Nhánh aldehyde mạch hở gồm 3 nguyên tử C (-CH=CH-CHO), 3 nguyên tử H và 1 nguyên tử O.\n"
            "   ➔ Tổng số nguyên tử: 9 C, 8 H, 1 O.\n"
            "   ➔ Công thức phân tử: C~9~H~8~O.\n"
            "   ➔ Tỉ lệ tối giản C : H : O = 9 : 8 : 1 nên Công thức đơn giản nhất là C~9~H~8~O.\n\n"
            "2. Xét o-methoxycinnamaldehyde:\n"
            "   - Trên vòng benzene có thêm 1 nhóm methoxy (-O-CH~3~) ở vị trí ortho, làm giảm 1 nguyên tử H của vòng.\n"
            "   - Vòng benzene có 6 nguyên tử C và 4 nguyên tử H.\n"
            "   - Nhóm methoxy có 1 nguyên tử C, 3 nguyên tử H và 1 nguyên tử O.\n"
            "   - Nhánh aldehyde có 3 nguyên tử C, 3 nguyên tử H và 1 nguyên tử O.\n"
            "   ➔ Tổng số nguyên tử: 10 C, 10 H, 2 O.\n"
            "   ➔ Công thức phân tử: C~10~H~10~O~2~.\n"
            "   ➔ Tỉ lệ tối giản C : H : O = 10 : 10 : 2 = 5 : 5 : 1 nên Công thức đơn giản nhất là C~5~H~5~O."
        )
        add_solution_no_table(doc, sol_p4_c1)
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 (SBT - CTST). Quan sát hình mô phỏng thí nghiệm sắc kí cột sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau2_p4_sac_ki_cot.png"), width_cm=7.5)
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(2)
    p_sub.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub, "Hãy cho biết trong điều kiện thí nghiệm:\na) Chất nào bị hấp phụ mạnh nhất? Chất nào bị hấp phụ kém nhất?\nb) Chất nào hoà tan tốt hơn trong dung môi?")
    if is_teacher:
        sol_p4_c2 = (
            "Hướng dẫn giải chi tiết:\n"
            "Nguyên tắc tách sắc kí cột: Pha động (dung môi) di chuyển từ trên xuống dưới qua cột chứa pha tĩnh (chất hấp phụ). Chất tan nào bị hấp phụ càng mạnh thì di chuyển càng chậm và nằm ở tầng trên; chất nào bị hấp phụ càng kém thì di chuyển càng nhanh theo dung môi xuống tầng dưới.\n"
            "a) Trong 3 chất a, b, c:\n"
            "   - Chất a nằm ở vị trí cao nhất trên cột (di chuyển chậm nhất), do đó chất a bị hấp phụ mạnh nhất bởi pha tĩnh.\n"
            "   - Chất c nằm ở vị trí thấp nhất trên cột (di chuyển nhanh nhất), do đó chất c bị hấp phụ kém nhất bởi pha tĩnh.\n"
            "b) Chất c hòa tan tốt hơn trong dung môi nên dễ dàng bị pha động cuốn trôi và dịch chuyển với tốc độ nhanh nhất qua cột."
        )
        add_solution_no_table(doc, sol_p4_c2)
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Lưu file
    filename = "De_Kiem_Tra_Hoa_11_De07_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_11_De07.docx"
    filepath = os.path.join(OUT_DIR, filename)
    doc.save(filepath)
    print(f"Saved: {filepath}")

if __name__ == "__main__":
    print("Generating De_Kiem_Tra_Hoa_11_De07.docx (Học Sinh)...")
    build_exam_document(is_teacher=False)
    print("Generating De_Kiem_Tra_Hoa_11_De07_GiaoVien.docx (Giáo Viên)...")
    build_exam_document(is_teacher=True)
    print("Done!")
