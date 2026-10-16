# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file Word cho Đề Số 11 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De11.docx (Bản Học Sinh)
- De_Kiem_Tra_Hoa_11_De11_GiaoVien.docx (Bản Giáo Viên)
Chương 4: HYDROCARBON
Tuân thủ 100% Form Đề Kiểm Tra trong GEMINI.md & form_de_kiem_tra.py:
- Tuyệt đối không dùng Table trong nội dung câu hỏi/đáp án/lời giải.
- Header & Footer chuẩn: #D32F2F, viền #8BC390, số trang tự động PAGE.
- Căn lề A4 1.5 cm.
- Chỉ số trên/dưới hóa học triệt để.
"""

import os
import sys
if sys.stdout.encoding != 'utf-8':
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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de11")

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
        exam_title="ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 4 (HYDROCARBON)",
        subject_line="MÔN: HOÁ HỌC - LỚP 11",
        num_pages="04",
        duration="50 phút",
        exam_code="111",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Mn = 55; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108. (Thể tích khí đo ở đkc 25 °C, 1 bar là 24,79 L/mol)",
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
    render_question_prompt(p, "Câu 1. [KNTT - SBT] (CH~3~)~2~CH-CH~3~ có tên theo danh pháp thay thế là")
    opts_c1 = [
        "*A. 2-methylpropane.*",
        "B. isobutane.",
        "C. butane.",
        "D. 2-methylbutane."
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Mạch chính dài nhất gồm 3 nguyên tử carbon (propane), nhóm methyl thế ở vị trí carbon số 2. Danh pháp thay thế chuẩn là 2-methylpropane. (Isobutane là tên thông thường).")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 (SBT - CTST). Đồ thị dưới đây thể hiện mối tương quan giữa nhiệt độ sôi và số nguyên tử carbon trong phân tử alkane không phân nhánh được biểu diễn như sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau2_dothi_alkane.png"), width_cm=8.5)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    render_rich_text(p, "Dựa vào đồ thị đã cho, số phân tử alkane không phân nhánh ở thể khí trong điều kiện thường (25 °C) là:")
    opts_c2 = [
        "*A. 4.*",
        "B. 3.",
        "C. 2.",
        "D. 1."
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Ở điều kiện thường (khoảng 25 °C), các chất có nhiệt độ sôi thấp hơn nhiệt độ phòng sẽ tồn tại ở thể khí. Quan sát đồ thị ta thấy các alkane không phân nhánh từ 1 đến 4 nguyên tử carbon (CH~4~, C~2~H~6~, C~3~H~8~, C~4~H~10~) đều có nhiệt độ sôi < 25 °C (nhỏ hơn 0 °C), do đó có 4 alkane ở thể khí.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. [KNTT - SBT] Phát biểu nào sau đây không đúng về phản ứng reforming alkane?")
    opts_c3 = [
        "A. Chuyển alkane mạch không phân nhánh thành các alkane mạch phân nhánh.",
        "B. Chuyển alkane mạch không phân nhánh thành các hydrocarbon mạch vòng.",
        "C. Số nguyên tử carbon của chất tham gia và của sản phẩm bằng nhau.",
        "*D. Nhiệt độ sôi của sản phẩm lớn hơn nhiều so với alkane tham gia phản ứng.*"
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Reforming biến đổi alkane mạch không phân nhánh thành alkane mạch phân nhánh hoặc arene có cùng số C. Cấu trúc mạch phân nhánh có diện tích tiếp xúc phân tử nhỏ hơn làm giảm tương tác van der Waals, dẫn tới nhiệt độ sôi thường thấp hơn alkane mạch thẳng tương ứng, không phải lớn hơn nhiều. Do đó phát biểu D không đúng.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4 (SBT - CD). Phát biểu nào sau đây về alkane là không đúng?")
    opts_c4 = [
        "A. Trong phân tử alkane chỉ có liên kết đơn.",
        "*B. Chỉ các alkane là chất khí ở điều kiện thường được dùng làm nhiên liệu.*",
        "C. Các alkane lỏng được dùng sản xuất xăng, dầu và làm dung môi.",
        "D. Các alkane rắn được dùng làm nhựa đường, nguyên liệu cho quá trình cracking."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Cả alkane thể khí (khí đốt hóa lỏng LPG), alkane thể lỏng (xăng, dầu diesel, dầu hoả) và thể rắn (parafin) đều được dùng làm nhiên liệu hoặc chất đốt. Do đó phát biểu B \"Chỉ các alkane là chất khí... được dùng làm nhiên liệu\" là không đúng.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. [KNTT – SBT] Hợp chất nào sau đây là một alkene?")
    opts_c5 = [
        "A. CH~3~CH~2~CH~3~.",
        "*B. CH~3~CH=CH~2~.*",
        "C. CH~3~C≡CH.",
        "D. CH~2~=C=CH~2~."
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Alkene là hydrocarbon không no mạch hở trong phân tử có chứa duy nhất 1 liên kết đôi C=C (công thức tổng quát C~n~H~2n~, n ≥ 2). Chất CH~3~CH=CH~2~ (propene) là một alkene. (CH~3~C≡CH là alkyne, CH~2~=C=CH~2~ là alkadiene).")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6 (SBT - CTST). Biểu đồ dưới đây thể hiện mối tương quan giữa nhiệt độ sôi và một số alkene:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau6_bieudo_alkene.png"), width_cm=8.5)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    render_rich_text(p, "Có bao nhiêu alkene trong biểu đồ ở thể khí trong điều kiện thường (25 °C)?")
    opts_c6 = [
        "A. 4.",
        "B. 2.",
        "C. 3.",
        "*D. 5.*"
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Đường màu xanh trên biểu đồ đánh dấu mốc 25 °C. Cả 5 alkene trong biểu đồ (ethene, propene, 2-methylpropene, but-1-ene, cis-but-2-ene) đều có nhiệt độ sôi nằm dưới mức 25 °C, nên ở điều kiện thường cả 5 alkene này đều tồn tại ở thể khí.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 7. Trùng hợp ethylene, sản phẩm thu được có công thức cấu tạo là:")
    opts_c7 = [
        "A. [-CH~2~=CH~2~-]~n~.",
        "*B. [-CH~2~-CH~2~-]~n~.*",
        "C. [-CH-CH-]~n~.",
        "D. [-CH~3~-CH~3~-]~n~."
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phản ứng trùng hợp ethylene làm đứt liên kết π tạo poly(ethylene) (PE): nCH~2~=CH~2~ -(t°, p, xt)-> [-CH~2~-CH~2~-]~n~.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 8 (SBT - CD). Phát biểu nào sau đây là không đúng?")
    opts_c8 = [
        "A. Trong phòng thí nghiệm, người ta điều chế ethene bằng cách tách nước ethanol và thu bằng cách dời chỗ của nước.",
        "B. Một ứng dụng quan trọng của acetylene là làm nhiên liệu trong đèn xì oxygen - acetylene.",
        "*C. Trong công nghiệp, người ta điều chế acetylene bằng cách nhiệt phân nhanh methane có xúc tác hoặc cho calcium carbide (thành phần chính của đất đèn) tác dụng với nước.*",
        "D. Một ứng dụng quan trọng của acetylene là làm nguyên liệu tổng hợp ethylene."
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong công nghiệp hiện đại, acetylene được điều chế bằng cách nhiệt phân nhanh methane ở khoảng 1500 °C rồi làm lạnh nhanh, phản ứng không dùng xúc tác: 2CH~4~ -(1500 °C, làm lạnh nhanh)-> C~2~H~2~ + 3H~2~. Do đó phát biểu C \"có xúc tác\" là không đúng.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 9 (SBT - CTST). Công thức phân tử nào dưới đây không thể là của một arene?")
    opts_c9 = [
        "A. C~7~H~8~.",
        "B. C~10~H~8~.",
        "*C. C~11~H~18~.*",
        "D. C~8~H~8~."
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Arene phải chứa ít nhất 1 vòng benzene (3 liên kết π và 1 vòng) nên có độ bất bão hòa k ≥ 4. Với C~11~H~18~: k = (2×11 + 2 - 18) / 2 = 3 < 4, do đó không thể là một arene.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 10 (SBT - CTST). Biết độ dài liên kết C=C là 134 pm, liên kết C-C là 154 pm. Thực tế, 3 liên kết π trong vòng benzene không cố định mà trải đều trên toàn bộ vòng benzene. Giá trị nào dưới đây phù hợp với độ dài liên kết giữa carbon và carbon trong phân tử benzene?")
    opts_c10 = [
        "A. 125 pm.",
        "B. 132 pm.",
        "C. 160 pm.",
        "*D. 139 pm.*"
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hệ electron π liên hợp trải đều toàn bộ vòng benzene khiến cho 6 liên kết C-C hoàn toàn tương đương nhau và có tính chất trung gian giữa liên kết đơn (154 pm) và liên kết đôi (134 pm). Giá trị thực tế đo được là 139 pm.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 11 (SBT – KNTT). Nhận xét nào sau đây không đúng đối với phản ứng cộng chlorine vào benzene?")
    opts_c11 = [
        "A. Khó hơn phản ứng cộng chlorine vào ethylene.",
        "B. Xảy ra với điều kiện ánh sáng tử ngoại và đun nóng.",
        "C. Sản phẩm thu được là 1,2,3,4,5,6-hexachlorocyclohexane.",
        "*D. Tỉ lệ mol của các chất tham gia phản ứng là 1 : 1.*"
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phương trình phản ứng: C~6~H~6~ + 3Cl~2~ -(as)-> C~6~H~6~Cl~6~ (hexachlorocyclohexane). Tỉ lệ mol giữa benzene và chlorine tham gia phản ứng là 1 : 3, không phải 1 : 1.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 12 (SBT - CD). Một trong những ứng dụng của toluene là")
    opts_c12 = [
        "*A. làm phụ gia để tăng chỉ số octane của nhiên liệu.*",
        "B. làm chất đầu để sản xuất methylcyclohexane.",
        "C. làm chất đầu để điều chế phenol.",
        "D. làm chất đầu để sản xuất polystyrene."
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Toluene có chỉ số octane cao (~120) nên được dùng làm chất phụ gia pha trộn vào nhiên liệu (xăng) để nâng cao khả năng chống kích nổ. (Polystyrene được sản xuất từ styrene, phenol điều chế chủ yếu từ cumene).")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 13 (SBT - CD). Cho butane phản ứng với chlorine (chiếu sáng, tỉ lệ 1 : 1) thu được sản phẩm chính là")
    opts_c13 = [
        "*A. 2-chlorobutane.*",
        "B. 1-chlorobutane.",
        "C. 3-chlorobutane.",
        "D. 4-chlorobutane."
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Butane (CH~3~-CH~2~-CH~2~-CH~3~) có nguyên tử C bậc II ở vị trí 2 và 3, C bậc I ở vị trí 1 và 4. Theo quy tắc thế halogen, phản ứng ưu tiên thế vào nguyên tử H gắn với C bậc cao hơn, do đó 2-chlorobutane là sản phẩm chính.")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 14 (SBT - CTST). Alkane (A) có công thức phân tử C~8~H~18~. (A) tác dụng với chlorine đun nóng chỉ tạo một dẫn xuất monochloro duy nhất. Tên gọi của (A) là")
    opts_c14 = [
        "A. octane.",
        "B. 2-methylheptane.",
        "C. 2,2-dimethylhexane.",
        "*D. 2,2,3,3-tetramethylbutane.*"
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Chất 2,2,3,3-tetramethylbutane có cấu tạo (CH~3~)~3~C-C(CH~3~)~3~ đối xứng hoàn toàn, tất cả 18 nguyên tử H thuộc 6 nhóm -CH~3~ đều có vị trí tương đương nhau, do đó khi clo hóa chỉ tạo duy nhất 1 dẫn xuất monochloro là (CH~3~)~3~C-C(CH~3~)~2~CH~2~Cl.")

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 15 (SBT - CD). Dẫn dòng khí gồm acetylene và ethylene lần lượt đi vào ống nghiệm (1) đựng dung dịch AgNO~3~/NH~3~ ở điều kiện thường, sau đó dẫn tiếp qua ống nghiệm (2) đựng nước bromine. Hiện tượng thí nghiệm nào sau đây là không đúng?")
    opts_c15 = [
        "A. Ở ống nghiệm (1) có kết tủa màu vàng nhạt.",
        "B. Ở ống nghiệm (2) màu của nước bromine nhạt dần.",
        "C. Ở ống nghiệm (2) chất lỏng chia thành hai lớp.",
        "*D. Ở ống nghiệm (2) thu được chất lỏng đồng nhất.*"
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khí acetylene bị giữ lại hoàn toàn ở ống (1) do phản ứng tạo kết tủa vàng nhạt AgC≡CAg. Khí ethylene thoát ra đi vào ống nghiệm (2) phản ứng với dung dịch bromine: CH~2~=CH~2~ + Br~2~ -> CH~2~Br-CH~2~Br. Sản phẩm 1,2-dibromoethane là chất lỏng không tan trong nước và nặng hơn nước nên chìm xuống dưới, tách thành 2 lớp chất lỏng phân biệt. Do đó phát biểu D \"thu được chất lỏng đồng nhất\" là không đúng.")

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 16 (SBT – KNTT). Tên gọi của chất nào sau đây không đúng?")
    
    # Render 4 options for Cau 16 with cropped formula images
    # Option A
    p_opt = doc.add_paragraph()
    p_opt.paragraph_format.space_before = Pt(2)
    p_opt.paragraph_format.space_after = Pt(2)
    p_opt.paragraph_format.left_indent = Cm(0.5)
    run_lbl = p_opt.add_run("A. ")
    if is_teacher:
        set_run_style(run_lbl, bold=True, color=COLOR_CORRECT)
    if os.path.exists(os.path.join(IMG_DIR, "cau16_ct_A_clean.png")):
        run_img = p_opt.add_run()
        run_img.add_picture(os.path.join(IMG_DIR, "cau16_ct_A_clean.png"), width=Cm(2.2))
    run_txt = p_opt.add_run("   but-2-ene")
    if is_teacher:
        set_run_style(run_txt, bold=True, color=COLOR_CORRECT)

    # Option B
    p_opt = doc.add_paragraph()
    p_opt.paragraph_format.space_before = Pt(2)
    p_opt.paragraph_format.space_after = Pt(2)
    p_opt.paragraph_format.left_indent = Cm(0.5)
    p_opt.add_run("B. ")
    if os.path.exists(os.path.join(IMG_DIR, "cau16_ct_B_clean.png")):
        run_img = p_opt.add_run()
        run_img.add_picture(os.path.join(IMG_DIR, "cau16_ct_B_clean.png"), width=Cm(2.0))
    p_opt.add_run("   3-methylbut-1-yne")

    # Option C
    p_opt = doc.add_paragraph()
    p_opt.paragraph_format.space_before = Pt(2)
    p_opt.paragraph_format.space_after = Pt(2)
    p_opt.paragraph_format.left_indent = Cm(0.5)
    p_opt.add_run("C. ")
    if os.path.exists(os.path.join(IMG_DIR, "cau16_ct_C_clean.png")):
        run_img = p_opt.add_run()
        run_img.add_picture(os.path.join(IMG_DIR, "cau16_ct_C_clean.png"), width=Cm(2.2))
    p_opt.add_run("   2,2,4-trimethylpentane")

    # Option D
    p_opt = doc.add_paragraph()
    p_opt.paragraph_format.space_before = Pt(2)
    p_opt.paragraph_format.space_after = Pt(2)
    p_opt.paragraph_format.left_indent = Cm(0.5)
    p_opt.add_run("D. ")
    if os.path.exists(os.path.join(IMG_DIR, "cau16_ct_D_clean.png")):
        run_img = p_opt.add_run()
        run_img.add_picture(os.path.join(IMG_DIR, "cau16_ct_D_clean.png"), width=Cm(1.8))
    p_opt.add_run("   1-ethyl-2-methylbenzene")

    if is_teacher:
        add_solution_no_table(doc, "Công thức ở đáp án A có mạch chính 5 carbon, chứa liên kết đôi ở vị trí C số 2 (CH~3~-CH=CH-CH~2~-CH~3~), tên gọi thay thế đúng phải là pent-2-ene. Tên gọi \"but-2-ene\" là sai.")

    # Câu 17
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 17 (SBT – KNTT). Chất nào sau đây khi tác dụng với hỗn hợp HNO~3~ và H~2~SO~4~ đặc nóng tạo một sản phẩm mononitro hoá duy nhất?")
    opts_c17 = [
        "*A. Benzene.*",
        "B. Toluene.",
        "C. o-xylene.",
        "D. Naphthalene."
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tử benzene có cấu trúc đối xứng phẳng hoàn hảo, 6 nguyên tử hydrogen gắn vào vòng đều tương đương nhau nên khi nitro hóa ở tỉ lệ 1 : 1 chỉ tạo ra duy nhất 1 sản phẩm mononitro hóa là nitrobenzene (C~6~H~5~NO~2~). Toluene tạo hỗn hợp o- và p-nitrotoluene; o-xylene tạo hỗn hợp 2 dẫn xuất; naphthalene tạo 1-nitronaphthalene và 2-nitronaphthalene.")

    # Câu 18
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 18. Toluene có giá trị thương mại thấp hơn so với benzene. Người ta loại bỏ nhóm methyl khỏi toluene bằng quá trình dealkyl hóa: toluene được trộn với hydrogen ở 550 - 650 °C, 30 - 50 atm cùng xúc tác SiO~2~ và Al~2~O~3~. Phương trình hóa học của phản ứng là:")
    
    opts_c18 = [
        "A. C~6~H~5~-CH~3~ + 3H~2~ -(xt, t°)-> C~6~H~11~-CH~3~.",
        "B. C~6~H~5~-CH~3~ + H~2~ -(xt, t°)-> C~6~H~6~ + C~2~H~6~.",
        "*C. C~6~H~5~-CH~3~ + H~2~ -(xt, t°)-> C~6~H~6~ + CH~4~.*",
        "D. Sau phản ứng vẫn còn toluene và hydrogen dư."
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau18_pt_dealkyl.png"), width_cm=7.5)
    if is_teacher:
        add_solution_no_table(doc, "Phản ứng dealkyl hóa (hydrodealkylation) cắt bỏ nhóm methyl (-CH~3~) khỏi nhân thơm tạo ra benzene và methane theo phương trình: C~6~H~5~-CH~3~ + H~2~ -(xt, t°)-> C~6~H~6~ + CH~4~.")

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
    render_question_prompt(p, "Câu 1. Cho các alkane: Butane, hexane và decane; 2-methylheptane, octane và 2,2,4-trimethylpentane.")
    stmts_c1_p2 = [
        ("a) Nhiệt độ sôi của các alkane mạch không phân nhánh theo thứ tự: Butane < hexane < decane.", "ĐÚNG"),
        ("b) Nhiệt độ sôi của các alkane mạch không phân nhánh theo thứ tự: Butane > hexane > decane.", "SAI"),
        ("c) Nhiệt độ sôi của các đồng phân C~8~H~18~ theo thứ tự: 2,2,4-trimethylpentane < 2-methylheptane < octane.", "ĐÚNG"),
        ("d) Nhiệt độ sôi của các đồng phân C~8~H~18~ theo thứ tự: octane < 2-methylheptane < 2,2,4-trimethylpentane.", "SAI")
    ]
    for stmt, ans in stmts_c1_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Lời giải chi tiết:\n- a) ĐÚNG, b) SAI: Với các alkane mạch thẳng, khi phân tử khối tăng, diện tích tiếp xúc phân tử tăng làm tăng tương tác van der Waals giữa các phân tử, dẫn đến nhiệt độ sôi tăng dần: butane (C~4~, -0,5 °C) < hexane (C~6~, 69 °C) < decane (C~10~, 174 °C).\n- c) ĐÚNG, d) SAI: Các chất 2,2,4-trimethylpentane, 2-methylheptane và octane đều là đồng phân có cùng CTPT C~8~H~18~. Phân tử càng nhiều nhánh thì cấu trúc càng gọn cầu, diện tích tiếp xúc phân tử giảm làm tương tác van der Waals yếu hơn, nhiệt độ sôi thấp hơn: 2,2,4-trimethylpentane (3 nhánh, 99 °C) < 2-methylheptane (1 nhánh, 117,6 °C) < octane (không nhánh, 125,7 °C).")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Trong thí nghiệm điều chế và thử tính chất của acetylene trong phòng thí nghiệm:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "p2_cau2_tn_acetylene.png"), width_cm=8.5)
    stmts_c2_p2 = [
        ("a) Phương trình hóa học của phản ứng điều chế C~2~H~2~ trong thí nghiệm là: CaC~2~ + 2H~2~O -> C~2~H~2~ + Ca(OH)~2~.", "ĐÚNG"),
        ("b) Dung dịch bromine ban đầu có màu vàng cam, sau phản ứng sục khí acetylene thì dung dịch bromine mất màu.", "ĐÚNG"),
        ("c) Dung dịch bromine ban đầu không màu, sau phản ứng sục khí acetylene thì dung dịch chuyển sang màu nâu đỏ.", "SAI"),
        ("d) Trong phân tử C~2~H~2~, liên kết π ở liên kết ba kém bền vững, dễ bị đứt tạo thành liên kết σ với các nguyên tử bromine khi xảy ra phản ứng cộng.", "ĐÚNG")
    ]
    for stmt, ans in stmts_c2_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Lời giải chi tiết:\n- a) ĐÚNG: Khí acetylene được điều chế bằng phản ứng thủy phân calcium carbide: CaC~2~ + 2H~2~O -> C~2~H~2~↑ + Ca(OH)~2~. Khí sinh ra được dẫn qua bình chứa dung dịch NaOH đặc để giữ lại các tạp chất có tính acid (H~2~S, SO~2~) trước khi thử tính chất.\n- b) ĐÚNG, c) SAI: Nước bromine có màu nâu đỏ hoặc vàng da cam, khi sục C~2~H~2~ vào phản ứng cộng xảy ra: CH≡CH + 2Br~2~ -> CHBr~2~-CHBr~2~ làm dung dịch mất màu. Nước bromine không phải là chất không màu.\n- d) ĐÚNG: Liên kết ba -C≡C- gồm 1 liên kết σ bền và 2 liên kết π kém bền, dễ bị bẻ gãy khi tham gia phản ứng cộng để tạo thành liên kết σ bền vững hơn.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Terephthalic acid là nguyên liệu để tổng hợp nhựa poly(ethylene terephthalate) (PET) dùng để sản xuất tơ sợi, chai nhựa. Terephthalic acid có thể được tổng hợp từ arene X có công thức phân tử C~8~H~10~ bằng cách oxi hóa X bởi dung dịch thuốc tím:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "p2_cau3_sodo_pet.png"), width_cm=7.5)
    stmts_c3_p2 = [
        ("a) Chất X có độ bất bão hòa k = 4.", "ĐÚNG"),
        ("b) Chất X có tên là 1,4-dimethylbenzene hay p-xylene.", "ĐÚNG"),
        ("c) Chất X có tên là 1,3-dimethylbenzene hay m-xylene.", "SAI"),
        ("d) Chất X có tên là 1,2-dimethylbenzene hay o-xylene.", "SAI")
    ]
    for stmt, ans in stmts_c3_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Lời giải chi tiết:\n- a) ĐÚNG: Với arene C~8~H~10~, độ bất bão hòa k = (2×8 + 2 - 10) / 2 = 4 (tương ứng 1 vòng benzene gồm 3 liên kết π và 1 vòng).\n- b) ĐÚNG, c) SAI, d) SAI: Terephthalic acid là benzene-1,4-dicarboxylic acid (hai nhóm -COOH ở vị trí para đối diện nhau). Quá trình oxi hóa nhóm alkyl gắn trên nhân thơm bởi KMnO~4~ giữ nguyên vị trí khung vòng, nên X phải là 1,4-dimethylbenzene (p-xylene). Oxi hóa o-xylene tạo phthalic acid (vị trí 1,2) và m-xylene tạo isophthalic acid (vị trí 1,3).")
        add_centered_picture(doc, os.path.join(IMG_DIR, "p2_cau3_pxylene.png"), width_cm=2.5)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Có một số loại trái cây chưa chín mà chúng ta lại muốn được sớm thưởng thức chúng, chẳng hạn một quả bơ, xoài, ...")
    stmts_c4_p2 = [
        ("a) Có một cách giải quyết đơn giản là cho quả bơ vào túi giấy cùng với vài quả chuối, bơ sẽ chín nhanh hơn nhiều.", "ĐÚNG"),
        ("b) Chuối khi chín sẽ giải phóng một loại hormone thực vật ở thể khí là ethylene.", "ĐÚNG"),
        ("c) Chuối khi chín sẽ giải phóng một loại hormone thực vật ở thể khí là acetylene.", "SAI"),
        ("d) Khí ethylene chỉ giúp đẩy nhanh quá trình chín của chuối, còn nhiều loại quả khác thì không có tác dụng.", "SAI")
    ]
    for stmt, ans in stmts_c4_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Lời giải chi tiết:\n- a) ĐÚNG, b) ĐÚNG: Chuối chín sản sinh ra lượng lớn hormone thực vật tự nhiên ở thể khí là ethylene (C~2~H~4~). Khi giữ trong túi giấy kín, khí ethylene tích tụ thúc đẩy các enzyme xúc tác quá trình hô hấp và phân giải tinh bột, giúp bơ và xoài nhanh chín mềm.\n- c) SAI: Hormone thực vật tự nhiên kích thích chín là ethylene, không phải acetylene.\n- d) SAI: Khí ethylene có tác động thúc đẩy chín đối với hầu hết các loại trái cây hô hấp đột biến (climacteric fruits) như bơ, chuối, xoài, cà chua, đu đủ, táo, lê,...")

    # ==========================================
    # PHẦN III
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 6. Điền kết quả vào ô trả lời."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. Khí đốt hoá lỏng (Liquified Petroleum Gas, viết tắt là LPG) hay còn gọi là gas, là hỗn hợp khí chủ yếu gồm propane (C~3~H~8~) và alkane Y đã được hoá lỏng. Alkane Y có chứa 4 nguyên tử carbon trong phân tử. Tỉ lệ số nguyên tử hydrogen : số nguyên tử carbon trong Y là y. Giá trị của y bằng bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Đáp án: 2,5\nLời giải: Alkane Y chứa 4 nguyên tử carbon nên công thức phân tử của Y là C~4~H~10~. Tỉ lệ số nguyên tử H : C là y = 10 / 4 = 2,5.")
    else:
        pass

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Alkene C~4~H~8~ có bao nhiêu đồng phân cấu tạo mà khi tác dụng với dung dịch HCl chỉ cho một sản phẩm hữu cơ duy nhất?")
    if is_teacher:
        add_solution_no_table(doc, "Đáp án: 1\nLời giải: Để khi phản ứng cộng HCl chỉ cho duy nhất 1 sản phẩm hữu cơ thì phân tử alkene phải có cấu trúc đối xứng xứng qua liên kết đôi C=C. Các đồng phân cấu tạo alkene C~4~H~8~ gồm:\n- but-1-ene: CH~2~=CH-CH~2~-CH~3~ (không đối xứng -> cho 2 sản phẩm: 2-chlorobutane và 1-chlorobutane).\n- 2-methylpropene: CH~2~=C(CH~3~)~2~ (không đối xứng -> cho 2 sản phẩm).\n- but-2-ene: CH~3~-CH=CH-CH~3~ (đối xứng hoàn toàn -> chỉ cho 1 sản phẩm duy nhất là 2-chlorobutane).\nDo đề bài yêu cầu số đồng phân cấu tạo, đáp án là 1.")
    else:
        pass

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Cho các phát biểu sau về hydrocarbon thơm:\n(1) Arene còn gọi là hydrocarbon thơm là những hydrocarbon trong phân tử có chứa một hay nhiều vòng benzene.\n(2) Arene còn gọi là hydrocarbon thơm là những hydrocarbon trong phân tử có chứa một hay hai vòng benzene.\n(3) Hydrocarbon thơm gồm alkylbenzene và hydrocarbon thơm khác.\n(4) Alkylbenzene có công thức chung là C~n~H~2n-6~ (n ≥ 6).\n(5) Hydrocarbon thơm khác gồm styrene và naphthalene có công thức chung là C~n~H~2n-6~ (n ≥ 6).\nSố phát biểu đúng là bao nhiêu?")
    if is_teacher:
        add_solution_no_table(doc, "Đáp án: 3\nLời giải: Phân tích các phát biểu theo tài liệu gốc:\n- (2) đúng (theo phạm vi SGK xét 1 hoặc 2 vòng benzene; định nghĩa tổng quát (1) cũng chuẩn).\n- (3) đúng: Arene được phân thành alkylbenzene và các hydrocarbon thơm khác (như styrene, naphthalene).\n- (4) đúng: Dãy đồng đẳng alkylbenzene có CTTQ là C~n~H~2n-6~ (n ≥ 6).\n- (5) sai vì styrene là C~8~H~8~ (C~n~H~2n-8~), naphthalene là C~10~H~8~ (C~n~H~2n-12~), không tuân theo C~n~H~2n-6~.\nSố phát biểu đúng là 3.")
    else:
        pass

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Ethylene là một hormone thực vật có vai trò kích thích quá trình tạo mủ và làm cho mủ cao su chậm đông, chảy dai hơn. Trong nông nghiệp hiện đại, người ta áp khí ethylene trực tiếp vào vỏ thân cây để kích thích tiết mủ.")
    add_centered_picture(doc, os.path.join(IMG_DIR, "p3_cau4_cao_su.png"), width_cm=7.5)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    render_rich_text(p, "Trong công nghiệp, ethylene được điều chế từ phản ứng dehydro hóa ethane ở 500 °C có xúc tác: C~2~H~6~ -(xt, 500 °C)-> C~2~H~4~ + H~2~. Để bơm vào túi khí 60 cc (đkc: 25 °C, 1 bar) khí ethylene thì cần dùng bao nhiêu miligam ethane? (Làm tròn đến hàng đơn vị).")
    if is_teacher:
        add_solution_no_table(doc, "Đáp án: 73 (mg)\nLời giải chi tiết:\n- Thể tích ethylene cần dùng: V = 60 cc = 60 mL = 0,06 L.\n- Số mol C~2~H~4~ (ở đkc: 24,79 L/mol): n = 0,06 / 24,79 mol.\n- Phương trình phản ứng: C~2~H~6~ -(xt, t°)-> C~2~H~4~ + H~2~.\n- Số mol ethane cần dùng: n(C~2~H~6~) = n(C~2~H~4~) = 0,06 / 24,79 mol.\n- Khối lượng ethane: m = (0,06 / 24,79) × 30 ≈ 0,07261 g = 72,61 mg ≈ 73 mg.")
    else:
        pass

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Đất đèn có thành phần chính là calcium carbide (CaC~2~). Cho 20 gam một mẫu đất đèn tác dụng với nước dư thu được 7,437 lít khí acetylene (ở đkc). Cho rằng các tạp chất trong mẫu trơ với nước và phản ứng xảy ra hoàn toàn. Độ tinh khiết của mẫu CaC~2~ trên là bao nhiêu phần trăm?")
    if is_teacher:
        add_solution_no_table(doc, "Đáp án: 96 (%)\nLời giải chi tiết:\n- Số mol khí acetylene: n(C~2~H~2~) = 7,437 / 24,79 = 0,3 mol.\n- Phương trình phản ứng: CaC~2~ + 2H~2~O -> Ca(OH)~2~ + C~2~H~2~↑\n- Theo phương trình: n(CaC~2~) = n(C~2~H~2~) = 0,3 mol.\n- Khối lượng CaC~2~ nguyên chất: m = 0,3 × 64 = 19,2 gam.\n- Độ tinh khiết của mẫu: % CaC~2~ = (19,2 / 20) × 100% = 96%.")
    else:
        pass

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Butane là một trong hai thành phần chính của khí đốt hóa lỏng (LPG). Khi đốt cháy hoàn toàn 1 mol butane tỏa ra lượng nhiệt là 2497 kJ. Để đun nóng 1 gam nước tăng thêm 1 °C cần cung cấp nhiệt lượng là 4,18 J. Tính khối lượng butane (theo gam, làm tròn đến một chữ số thập phân) cần đốt để đun sôi 2 lít nước từ 25 °C lên 100 °C. Biết khối lượng riêng của nước là 1 g/mL và hiệu suất hấp thụ nhiệt của nước là 60%.")
    if is_teacher:
        add_solution_no_table(doc, "Đáp án: 24,3 (gam)\nLời giải chi tiết:\n- Khối lượng 2 lít nước: m(H~2~O) = 2000 mL × 1 g/mL = 2000 gam.\n- Nhiệt lượng cần thiết để đun nóng nước từ 25 °C lên 100 °C:\n  Q(thu) = m × c × Δt = 2000 × 4,18 × (100 - 25) = 627 000 J = 627 kJ.\n- Vì hiệu suất hấp thụ nhiệt là 60%, nhiệt lượng thực tế cần đốt cháy tỏa ra là:\n  Q(tỏa) = Q(thu) / 60% = 627 / 0,60 = 1045 kJ.\n- Số mol butane cần đốt: n(C~4~H~10~) = 1045 / 2497 ≈ 0,4185 mol.\n- Khối lượng butane cần dùng: m = 0,4185 × 58 ≈ 24,27 g ≈ 24,3 gam.")
    else:
        pass

    # ==========================================
    # PHẦN IV
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN IV. Tự luận.",
        subtitle="Thí sinh trình bày chi tiết lời giải vào bài làm."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1 (SBT - CTST). Viết công thức khung phân tử của các chất sau:\na) propene.\nb) pent-1-ene.\nc) 3-methylpent-1-yne.\nd) cis-pent-2-ene.\ne) trans-pent-2-ene.")
    if is_teacher:
        add_solution_no_table(doc, "Lời giải chi tiết:\nCông thức khung phân tử tương ứng của các hợp chất:")
        add_centered_picture(doc, os.path.join(IMG_DIR, "p4_cau1_khung_pt.png"), width_cm=13.0)
        add_solution_no_table(doc, "- a) propene: Khung 3 nguyên tử carbon với liên kết đôi ở đầu mạch.\n- b) pent-1-ene: Mạch hở 5 nguyên tử carbon dạng gấp khúc, liên kết đôi tại C số 1.\n- c) 3-methylpent-1-yne: Mạch chính 5 carbon chứa liên kết ba ở C số 1, nhánh methyl ở vị trí C số 3.\n- d) cis-pent-2-ene: Đồng phân hình học có 2 mạch nhánh alkyl (-CH~3~ và -C~2~H~5~) cùng nằm về một phía của mặt phẳng liên kết đôi C=C.\n- e) trans-pent-2-ene: Đồng phân hình học có 2 mạch nhánh alkyl nằm về hai phía đối diện của mặt phẳng liên kết đôi C=C.")
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Trinitrotoluene (TNT) là một hợp chất hữu cơ quan trọng có công thức C~6~H~2~(NO~2~)~3~CH~3~, được tổng hợp bằng phản ứng trinitro hóa toluene.\na) Viết phương trình hóa học điều chế 2,4,6-trinitrotoluene từ toluene và nitric acid đặc (xúc tác sulfuric acid đặc, đun nóng).\nb) Tính khối lượng toluene (theo tấn, làm tròn đến một chữ số thập phân) cần dùng để điều chế 36 000 tấn trinitrotoluene, biết hiệu suất phản ứng của toàn bộ quá trình đạt 70%.")
    if is_teacher:
        add_solution_no_table(doc, "Lời giải chi tiết:\na) Phương trình hóa học điều chế TNT:\n   C~6~H~5~-CH~3~ + 3HNO~3~ -(H~2~SO~4~ đặc, t°)-> C~6~H~2~(NO~2~)~3~-CH~3~ + 3H~2~O\nb) Phân tử khối:\n   M(toluene, C~7~H~8~) = 92 g/mol.\n   M(TNT, C~7~H~5~N~3~O~6~) = 12×7 + 5 + 14×3 + 16×6 = 227 g/mol.\n- Theo phương trình hóa học:\n   Cứ 92 tấn toluene phản ứng 100% sẽ tạo thành 227 tấn TNT.\n- Khối lượng toluene lý thuyết cần dùng để tạo 36 000 tấn TNT:\n   m(lt) = (36 000 × 92) / 227 ≈ 14 590,31 tấn.\n- Do hiệu suất của cả quá trình chỉ đạt 70%, lượng toluene thực tế cần lấy là:\n   m(tt) = m(lt) / H = 14 590,31 / 70% = 14 590,31 / 0,70 ≈ 20 843,3 tấn.")
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Lưu tài liệu
    fname = "De_Kiem_Tra_Hoa_11_De11_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_11_De11.docx"
    fpath = os.path.join(OUT_DIR, fname)
    doc.save(fpath)
    print(f" Đã lưu thành công: {fpath}")

if __name__ == "__main__":
    print(" Bắt đầu tạo 2 file Word cho Đề 11...")
    build_exam_document(is_teacher=False)
    build_exam_document(is_teacher=True)
    print(" Hoàn thành xuất 2 file Word chuẩn Form Đề Kiểm Tra!")
