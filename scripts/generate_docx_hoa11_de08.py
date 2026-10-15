# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file Word cho ĐỀ SỐ 8 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De08.docx (Bản Học Sinh)
- De_Kiem_Tra_Hoa_11_De08_GiaoVien.docx (Bản Giáo Viên)
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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de08")

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
        exam_code="108",
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
    render_question_prompt(p, "Câu 1 [KNTT - SBT]. Xét phản ứng quang hợp: 6CO~2~ + 6H~2~O \u2192 C~6~H~12~O~6~ + 6O~2~. Chất nào trong phản ứng này thuộc loại hợp chất hữu cơ?")
    opts_c1 = [
        "A. CO~2~.",
        "B. H~2~O.",
        "*C. C~6~H~12~O~6~.*",
        "D. O~2~."
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "C~6~H~12~O~6~ (glucose) là hợp chất hữu cơ; các chất CO~2~, H~2~O, O~2~ đều là các chất vô cơ.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 [KNTT - SBT]. Hydrocarbon là loại hợp chất hữu cơ mà thành phần phân tử có các nguyên tố nào sau đây?")
    opts_c2 = [
        "*A. C và H.*",
        "B. C, H và O.",
        "C. C, H và N.",
        "D. C, H, O và N."
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hydrocarbon là loại hợp chất hữu cơ mà thành phần phân tử chỉ chứa hai nguyên tố carbon (C) và hydrogen (H).")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3 [SBT - CTST]. Phương pháp chưng cất dùng để tách các chất")
    opts_c3 = [
        "*A. có nhiệt độ sôi khác nhau.*",
        "B. có nhiệt độ nóng chảy khác nhau.",
        "C. có độ tan khác nhau.",
        "D. có khối lượng riêng khác nhau."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phương pháp chưng cất dùng để tách các chất lỏng có nhiệt độ sôi khác nhau ở cùng áp suất ra khỏi hỗn hợp.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Chưng cất gồm bao nhiêu giai đoạn?")
    opts_c4 = [
        "A. 1 giai đoạn: thay đổi điều kiện hoà tan.",
        "*B. 2 giai đoạn: bay hơi và ngưng tụ.*",
        "C. 3 giai đoạn: đun nóng, bay hơi, ngưng tụ.",
        "D. 4 giai đoạn: hoà tan, lọc nóng, để nguội, lọc chất kết tinh."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Quá trình chưng cất gồm hai giai đoạn liên tiếp nhau: bay hơi (chuyển chất lỏng thành hơi khi sôi) và ngưng tụ (làm lạnh hơi chuyển lại thành chất lỏng).")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Glucose là hợp chất hữu cơ có nhiều trong các loại quả chín, đặc biệt là quả nho. Công thức phân tử của glucose là C~6~H~12~O~6~. Công thức đơn giản nhất của glucose là")
    opts_c5 = [
        "A. C~1,5~H~3~O~1,5~.",
        "*B. CH~2~O.*",
        "C. C~3~H~4~O~3~.",
        "D. CHO~2~."
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Tỉ lệ số nguyên tử trong C~6~H~12~O~6~ là N(C) : N(H) : N(O) = 6 : 12 : 6 = 1 : 2 : 1 \u2192 Công thức đơn giản nhất là CH~2~O.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Phân tử khối của chất hữu cơ nào sau đây là 30?")
    opts_c6 = [
        "*A. HCHO.*",
        "B. C~3~H~8~.",
        "C. C~2~H~4~.",
        "D. C~3~H~4~."
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "M(HCHO) = 1 + 12 + 1 + 16 = 30. (Các chất khác: M(C~3~H~8~) = 44, M(C~2~H~4~) = 28, M(C~3~H~4~) = 40).")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 7. Hợp chất nào sau đây là đồng đẳng của acetic acid CH~3~COOH?")
    opts_c7 = [
        "*A. HCOOH.*",
        "B. CH~3~COOCH~3~.",
        "C. HOCH~2~COOH.",
        "D. HOOC\u2013COOH."
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Đồng đẳng của acetic acid (carboxylic acid no, đơn chức, mạch hở C~n~H~2n~O~2~) là formic acid (methanoic acid) HCOOH.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 8 [KNTT - SBT]. Những hợp chất hữu cơ khác nhau nhưng có cùng công thức phân tử được gọi là các chất")
    opts_c8 = [
        "*A. đồng phân của nhau.*",
        "B. đồng đẳng của nhau.",
        "C. đồng vị của nhau.",
        "D. đồng khối của nhau."
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Các hợp chất khác nhau nhưng có cùng công thức phân tử được gọi là các đồng phân của nhau.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 9 [KNTT - SBT]. Biết rằng hydrocarbon no chỉ chứa liên kết đơn, hydrocarbon không no có chứa liên kết bội và hydrocarbon thơm có chứa vòng benzene. Xét các chất sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau9_p1_cau_truc_8chat.png"), width_cm=13.5)
    p_sub9 = doc.add_paragraph()
    p_sub9.paragraph_format.space_before = Pt(2)
    p_sub9.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub9, "Nhận định nào sau đây không đúng?")
    opts_c9 = [
        "A. số hydrocarbon bằng 5.",
        "B. số dẫn xuất hydrocarbon bằng 3.",
        "C. số hydrocarbon no bằng 2.",
        "*D. số hydrocarbon không no bằng 3.*"
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích 8 chất trên:\n- Dẫn xuất hydrocarbon gồm 3 chất: CH~3~Cl, CH~3~CH~2~OH, CH~3~COOH \u2192 B đúng.\n- Hydrocarbon gồm 5 chất: CH~4~, C~3~H~6~ (cyclopropane), C~2~H~4~, C~2~H~2~, C~6~H~6~ \u2192 A đúng.\n- Hydrocarbon no gồm 2 chất: CH~4~ và cyclopropane \u2192 C đúng.\n- Hydrocarbon không no chỉ có 2 chất: C~2~H~4~ (chứa liên kết đôi) và C~2~H~2~ (chứa liên kết ba); còn C~6~H~6~ là hydrocarbon thơm \u2192 D sai.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 10 [KNTT - SBT]. Nhận định nào sau đây không đúng?")
    opts_c10 = [
        "A. CH~4~, CH~2~=CH~2~ và CH\u2261CH là những hydrocarbon.",
        "B. CH~3~OH và HOCH~2~\u2013CH~2~OH là những alcohol.",
        "C. CH~3~COOH và CH~2~(COOH)~2~ là những carboxylic acid.",
        "*D. CH~3~CH=O và CH~3~COCH~3~ là những aldehyde.*"
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "CH~3~COCH~3~ (acetone) là hợp chất ketone (nhóm carbonyl gắn với hai gốc hydrocarbon), không phải aldehyde. Do đó D không đúng.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 11 [KNTT \u2013 SBT]. Trong quá trình chưng cất dầu thô, người ta thu được nhiều phân đoạn dầu mỏ, trong đó có xăng (thành phần chính là hỗn hợp các hydrocarbon có số nguyên tử C từ 4 đến 12, nhiệt độ sôi khoảng từ 40 °C đến 200 °C) và dầu hoả (thành phần chính là hỗn hợp các hydrocarbon có số nguyên tử C từ 12 đến 16, nhiệt độ sôi khoảng từ 200 °C đến 250 °C). Sản phẩm thu được ở 150 °C đến 200 °C là")
    opts_c11 = [
        "A. xăng.",
        "B. dầu hoả.",
        "*C. xăng và dầu hoả.*",
        "D. dầu hoả và xăng."
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Ở khoảng 150 °C đến 200 °C, phân đoạn chưng cất có sự gối đầu giữa các hydrocarbon mạch dài của xăng và các hydrocarbon đầu của dầu hoả, do đó sản phẩm thu được gồm cả xăng và dầu hoả.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 12 [KNTT \u2013 SBT]. Thêm benzene vào ống nghiệm đựng dung dịch nước bromine. Sau một thời gian quan sát thấy màu đỏ nâu của bromine")
    opts_c12 = [
        "A. chủ yếu trong lớp nước.",
        "*B. chủ yếu trong lớp benzene.*",
        "C. phân bố đồng đều ở hai lớp.",
        "D. bị mất màu hoàn toàn."
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Bromine tan tốt trong dung môi không phân cực (benzene) hơn rất nhiều so với trong dung môi phân cực (nước). Do đó bromine bị chiết vào lớp benzene ở phía trên làm lớp benzene có màu đỏ nâu đậm.")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 13 [SBT-CĐ]. Sau khi biết công thức thực nghiệm, có thể xác định công thức phân tử của hợp chất hữu cơ dựa trên đặc điểm nào sau đây?")
    opts_c13 = [
        "*A. Phân tử khối của chất.*",
        "B. Thành phần phần trăm về khối lượng của các nguyên tố có trong phân tử chất.",
        "C. Khối lượng các sản phẩm thu được khi đốt cháy hoàn toàn một lượng chất xác định.",
        "D. Các hấp thụ đặc trưng trên phổ IR của chất."
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Công thức phân tử = (Công thức thực nghiệm)~n~. Khi biết phân tử khối M của chất, ta xác định được hệ số n = M / M~CTTN~, từ đó suy ra công thức phân tử.")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 14 [SBT-KNTT]. Hình sau đây là phổ khối lượng của phân tử acetic acid.")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau14_p1_ms_acetic_acid.png"), width_cm=10.0)
    p_sub14 = doc.add_paragraph()
    p_sub14.paragraph_format.space_before = Pt(2)
    p_sub14.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub14, "Phân tử khối của acetic acid bằng")
    opts_c14 = [
        "A. 43.",
        "B. 45.",
        "*C. 60.*",
        "D. 29."
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trên phổ khối lượng MS, peak ion phân tử [M^+^] có giá trị m/z lớn nhất là 60. Do đó phân tử khối của acetic acid bằng 60.")

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 15. Cho các cặp chất:\n(1) CH~3~CH~2~OH và CH~3~OCH~3~\n(2) CH~3~CH~2~Br và BrCH~2~CH~3~\n(3) CH~2~=CH\u2013CH~2~OH và CH~3~CH~2~CHO\n(4) (CH~3~)~2~NH và CH~3~CH~2~NH~2~\nCó bao nhiêu cặp là đồng phân cấu tạo?")
    opts_c15 = [
        "A. 1.",
        "B. 2.",
        "*C. 3.*",
        "D. 4."
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "- Cặp (1): C~2~H~6~O (alcohol và ether) \u2192 là đồng phân cấu tạo.\n- Cặp (2): CH~3~CH~2~Br và BrCH~2~CH~3~ cùng là bromoethane \u2192 cùng một chất, không phải đồng phân.\n- Cặp (3): C~3~H~6~O (alcohol không no và aldehyde) \u2192 là đồng phân cấu tạo.\n- Cặp (4): C~2~H~7~N (amine bậc 2 và amine bậc 1) \u2192 là đồng phân cấu tạo.\nVậy có 3 cặp là đồng phân cấu tạo (1, 3, 4).")

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 16. Cho chất acetylene (C~2~H~2~) và benzene (C~6~H~6~), hãy chọn nhận xét đúng trong các nhận xét sau đây:")
    opts_c16 = [
        "A. Hai chất đó giống nhau về công thức phân tử và khác nhau về công thức đơn giản nhất.",
        "*B. Hai chất đó khác nhau về công thức phân tử và giống nhau về công thức đơn giản nhất.*",
        "C. Hai chất đó khác nhau về công thức phân tử và khác nhau về công thức đơn giản nhất.",
        "D. Hai chất đó có cùng công thức phân tử và cùng công thức đơn giản nhất."
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "C~2~H~2~ và C~6~H~6~ có công thức phân tử khác nhau, nhưng đều có tỉ lệ nguyên tử C : H = 1 : 1 nên có cùng công thức đơn giản nhất là CH.")

    # Câu 17
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 17. Methadone là một loại thuốc dùng trong cai nghiện ma túy, nó thực chất cũng là một loại chất gây nghiện nhưng \"nhẹ\" hơn các loại ma túy thông thường và dễ kiểm soát hơn. Công thức cấu tạo của methadone như hình dưới:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau17_p1_methadone.png"), width_cm=11.0)
    p_sub17 = doc.add_paragraph()
    p_sub17.paragraph_format.space_before = Pt(2)
    p_sub17.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub17, "Công thức phân tử của methadone là")
    opts_c17 = [
        "A. C~17~H~27~NO.",
        "B. C~17~H~22~NO.",
        "C. C~21~H~29~NO.",
        "*D. C~21~H~27~NO.*"
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Đếm số nguyên tử:\n- Số nguyên tử C: 2 vòng benzene (12 C) + C bậc 4 (1 C) + nhóm -CO-CH~2~CH~3~ (3 C) + mạch -CH~2~-CH(CH~3~)-N(CH~3~)~2~ (5 C) \u2192 Tổng = 21 C.\n- Độ bất bão hòa k = 2 \u00d7 4 (vòng benzen) + 1 (C=O) = 9 \u2192 Số H = 2 \u00d7 21 + 2 + 1 - 2 \u00d7 9 = 27 H.\n- Có 1 nguyên tử N và 1 nguyên tử O \u2192 CTPT là C~21~H~27~NO.")

    # Câu 18
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 18. Cho các phát biểu về đặc điểm chung của các phân tử hợp chất hữu cơ:\n(a) Thành phần nguyên tố chủ yếu là C và H.\n(b) Có thể chứa nguyên tố khác như Cl, N, P, O.\n(c) Liên kết hóa học chủ yếu là liên kết cộng hoá trị.\n(d) Liên kết hoá học chủ yếu là liên kết ion.\n(e) Dễ bay hơi, khó cháy.\n(f) Phản ứng hoá học xảy ra nhanh.\nCác phát biểu đúng là")
    opts_c18 = [
        "A. 4",
        "*B. 3*",
        "C. 2",
        "D. 1"
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Các phát biểu đúng gồm: (a), (b), (c) \u2192 Có 3 phát biểu đúng.\n- (d) sai vì liên kết chủ yếu là liên kết cộng hóa trị.\n- (e) sai vì hợp chất hữu cơ thường dễ cháy.\n- (f) sai vì phản ứng của hợp chất hữu cơ thường xảy ra chậm và theo nhiều hướng.")

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
    render_question_prompt(p, "Câu 1. Đặc điểm chung của các phân tử hợp chất hữu cơ:")
    stmts_c1_p2 = [
        ("a) Thành phần nguyên tố chủ yếu là C và H.", "ĐÚNG"),
        ("b) Liên kết hóa học chủ yếu là liên kết cộng hoá trị.", "ĐÚNG"),
        ("c) Liên kết hoá học chủ yếu là liên kết ion.", "SAI"),
        ("d) Dễ bay hơi, khó cháy.", "SAI")
    ]
    for stmt, ans in stmts_c1_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c1_p2 = (
            "a) Đúng: Thành phần phân tử hợp chất hữu cơ bắt buộc có carbon, thường có hydrogen.\n"
            "b) Đúng: Liên kết hóa học chủ yếu trong hợp chất hữu cơ là liên kết cộng hóa trị.\n"
            "c) Sai: Hợp chất hữu cơ chủ yếu chứa liên kết cộng hóa trị, liên kết ion chỉ xuất hiện trong muối hữu cơ.\n"
            "d) Sai: Các hợp chất hữu cơ thường kém bền nhiệt và dễ cháy."
        )
        add_solution_no_table(doc, sol_c1_p2)

    # Câu 2 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Một mẫu hoa hoè được xác định có hàm lượng rutin là 26%. Người ta đun sôi hoa hoè với nước (100 °C) để chiết lấy rutin. Biết độ tan của rutin là 5,2 gam trong 1 lít nước ở 100 °C và là 0,125 gam trong 1 lít nước ở 25 °C.")
    stmts_c2_p2 = [
        ("a) Trong 100 gam hoa hòe có 30 gam rutin.", "SAI"),
        ("b) Cần dùng tối thiểu 5 lít nước để chiết được lượng rutin có trong 100 gam hoa hoè.", "ĐÚNG"),
        ("c) Giả thiết rằng toàn bộ lượng rutin trong hoa hoè đã tan vào nước khi chiết. Làm nguội dung dịch chiết 100 gam hoa hoè ở trên từ 100 °C xuống 25 °C thì thu được 28,0 gam rutin kết tinh.", "SAI"),
        ("d) Khi tăng lượng nước, lượng rutin hoà tan trong dung dịch ở 25 °C tăng lên nên lượng rutin kết tinh bị giảm đi.", "ĐÚNG")
    ]
    for stmt, ans in stmts_c2_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c2_p2 = (
            "a) Sai: Khối lượng rutin có trong 100 g hoa hoè là: 100 × 26% = 26 gam (không phải 30 gam).\n"
            "b) Đúng: Thể tích nước sôi ở 100 °C tối thiểu cần dùng là: V = 26 / 5,2 = 5 (lít).\n"
            "c) Sai: Lượng rutin còn tan lại trong 5 lít nước ở 25 °C là: 5 × 0,125 = 0,625 g → Lượng rutin kết tinh tách ra là: 26 − 0,625 = 25,375 gam (khác 28,0 gam).\n"
            "d) Đúng: Vì m~tan(25 °C)~ = V × 0,125 g, nếu tăng V thì lượng rutin tan lại tăng, dẫn tới lượng rutin kết tinh m~kt~ = m~ban đầu~ − m~tan~ bị giảm."
        )
        add_solution_no_table(doc, sol_c2_p2)

    # Câu 3 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Khi nghiên cứu thành phần hoá học của tinh dầu quế, người ta thu được nhiều hợp chất hữu cơ trong đó có cinnamaldehyde và o–methoxycinnamaldehyde với công thức cấu tạo:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau3_p2_cinnamaldehyde.png"), width_cm=11.5)
    p_sub3_p2 = doc.add_paragraph()
    p_sub3_p2.paragraph_format.space_before = Pt(2)
    p_sub3_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub3_p2, "Hãy chọn phát biểu Đúng hoặc Sai:")
    stmts_c3_p2 = [
        ("a) Công thức phân tử cinnamaldehyde: C~9~H~8~O.", "ĐÚNG"),
        ("b) Công thức đơn giản nhất cinnamaldehyde: C~9~H~8~O.", "ĐÚNG"),
        ("c) Công thức phân tử o–methoxycinnamaldehyde: C~5~H~5~O.", "SAI"),
        ("d) Công thức đơn giản nhất: C~10~H~10~O~2~.", "SAI")
    ]
    for stmt, ans in stmts_c3_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c3_p2 = (
            "a) Đúng: Cinnamaldehyde có gốc C~6~H~5~CH=CHCHO gồm 9 C, 8 H, 1 O → CTPT là C~9~H~8~O.\n"
            "b) Đúng: Tỉ lệ 9 : 8 : 1 không thể tối giản thêm, do đó CTĐGN trùng CTPT là C~9~H~8~O.\n"
            "c) Sai: o-methoxycinnamaldehyde có công thức phân tử là C~10~H~10~O~2~ (C~5~H~5~O là công thức đơn giản nhất).\n"
            "d) Sai: Tỉ lệ số nguyên tử là 10 : 10 : 2 = 5 : 5 : 1 → Công thức đơn giản nhất là C~5~H~5~O, không phải C~10~H~10~O~2~."
        )
        add_solution_no_table(doc, sol_c3_p2)

    # Câu 4 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Vanilin là hợp chất thiên nhiên, được sử dụng rộng rãi với chức năng phụ gia bổ sung hương thơm trong các loại đồ ăn, đồ uống, bánh kẹo, nước hoa,... Vanilin có cấu tạo như sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau4_p2_vanilin.png"), width_cm=4.5)
    stmts_c4_p2 = [
        ("a) Phân tử vanilin có nhóm chức aldehyde.", "ĐÚNG"),
        ("b) Phân tử vanilin có chứa nhóm chức alcohol.", "SAI"),
        ("c) Vanilin có công thức phân tử C~8~H~8~O~3~.", "ĐÚNG"),
        ("d) Vanilin thuộc loại hợp chất hữu cơ tạp chức.", "ĐÚNG")
    ]
    for stmt, ans in stmts_c4_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c4_p2 = (
            "a) Đúng: Phân tử vanilin có nhóm chức –CHO (aldehyde) gắn vào vòng thơm.\n"
            "b) Sai: Nhóm –OH gắn trực tiếp vào nguyên tử C của vòng thơm nên là nhóm chức phenol, không phải alcohol.\n"
            "c) Đúng: Đếm số nguyên tử: vòng C~6~H~3~ + –CHO (1 C, 1 H, 1 O) + –OCH~3~ (1 C, 3 H, 1 O) + –OH (1 H, 1 O) → CTPT là C~8~H~8~O~3~.\n"
            "d) Đúng: Vanilin chứa đồng thời các nhóm chức khác nhau: aldehyde (–CHO), ether (–OCH~3~), phenol (–OH) nên là hợp chất hữu cơ tạp chức."
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
    render_question_prompt(p, "Câu 1. Có bao nhiêu hydrocarbon trong số các chất dưới đây:\nCH~4~, NH~3~, C~2~H~2~, CCl~4~, C~2~H~4~, C~6~H~6~.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *4*", base_size=10.5)
        add_solution_no_table(doc, "Các hydrocarbon (hợp chất chỉ tạo bởi C và H) gồm 4 chất: CH~4~, C~2~H~2~, C~2~H~4~, C~6~H~6~. (NH~3~ là chất vô cơ, CCl~4~ là dẫn xuất halogen).")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 2 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Cây neem (cây thường xanh Ấn Độ) hay còn gọi là cây nim, sầu đâu (tên khoa học Azadirachta indica) được người Ấn Độ dùng hơn 4000 năm nay để làm đẹp và chữa bệnh. Các chiết xuất từ lá neem có thể ức chế sự sao chép của virus Dengue, kháng virus nên có tác dụng điều trị bệnh sốt xuất huyết hiệu quả. Người ta hái một ít lá neem tươi, rửa sạch, cho vào nồi đun kĩ, lọc lấy nước để sử dụng. Hãy cho biết cách làm trên thuộc loại phương pháp tách và tinh chế nào (số thứ tự nào) trong các phương pháp sau:\n1) Chưng cất; 2) Chiết; 3) Kết tinh; 4) Sắc kí")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *2*", base_size=10.5)
        add_solution_no_table(doc, "Cách đun lá neem với nước sôi để hòa tan các dược chất từ mô thực vật vào nước rồi lọc lấy dịch chiết là phương pháp chiết lỏng – rắn (phương pháp số 2).")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 3 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Phổ khối lượng của hợp chất hữu cơ X thu được như hình vẽ:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau3_p3_ms_benzene.png"), width_cm=11.0)
    p_sub3_p3 = doc.add_paragraph()
    p_sub3_p3.paragraph_format.space_before = Pt(2)
    p_sub3_p3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub3_p3, "Phân tử khối của hợp chất hữu cơ X là?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *78*", base_size=10.5)
        add_solution_no_table(doc, "Trên phổ MS, peak ion phân tử [M^+^] có giá trị m/z lớn nhất bằng 78. Do đó phân tử khối của hợp chất X là 78 (chính là benzene C~6~H~6~).")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 4 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Cho công thức sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau4_p3_khung_alkene.png"), width_cm=6.0)
    p_sub4_p3 = doc.add_paragraph()
    p_sub4_p3.paragraph_format.space_before = Pt(2)
    p_sub4_p3.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub4_p3, "Hợp chất này có tổng số nguyên tử Carbon và Hydrogen là bao nhiêu?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *21*", base_size=10.5)
        add_solution_no_table(doc, "Công thức khung phân tử biểu diễn một alkene mạch hở có 1 liên kết đôi C=C:\n- Mạch chính có 5 C, có 2 nhóm thế methyl (-CH~3~) → Tổng số nguyên tử C = 7.\n- Hợp chất là alkene mạch hở có CTPT là C~7~H~14~.\n→ Tổng số nguyên tử C và H = 7 + 14 = 21.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 5 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Cho quy trình thực hiện thí nghiệm tinh chế benzoic acid:\nCân chính xác 1 gam benzoic acid thô cho vào bình định mức 250 mL. Cho từ từ nước sôi vào bình định mức và lắc đều cho đến khi benzoic acid tan hết. Tiến hành lọc nóng, để yên cho dung dịch nguội từ từ benzoic acid sẽ kết tinh. Lọc bằng máy hút chân không, thu được m gam benzoic acid. Hiệu suất của quá trình kết tinh trên là 40%. Giá trị của m là bao nhiêu? (kết quả làm tròn đến hàng phần chục).")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *0,4*", base_size=10.5)
        add_solution_no_table(doc, "Khối lượng benzoic acid tinh khiết kết tinh thu được là:\nm = 1 × 40% = 0,4 (g).")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 6 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Camphor (có trong cây long não) là một chất rắn kết tinh màu trắng hay trong suốt giống như sáp với mùi thơm đặc trưng, thường dùng trong y học. Phần trăm khối lượng các nguyên tố trong camphor lần lượt là 78,94% carbon, 10,53% hydrogen và 10,53% oxygen. Từ phổ khối lượng của camphor xác định được giá trị m/z của peak [M^+^] bằng 152. Xác định số lượng nguyên tử của camphor.")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau6_p3_cay_long_nao.png"), width_cm=5.0)
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau6_p3_ms_camphor.png"), width_cm=9.5)
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *27*", base_size=10.5)
        add_solution_no_table(doc, "Đặt CTPT của camphor là C~x~H~y~O~z~ với M = 152:\n- x = (152 × 78,94%) / (12 × 100%) ≈ 10\n- y = (152 × 10,53%) / (1 × 100%) ≈ 16\n- z = (152 × 10,53%) / (16 × 100%) ≈ 1\n→ CTPT của camphor là C~10~H~16~O.\n→ Tổng số lượng nguyên tử trong phân tử camphor là: 10 + 16 + 1 = 27.")
    else:
        add_student_writing_lines(doc, num_lines=2)

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
    render_question_prompt(p, "Câu 1 [CĐ-SGK]. Hãy xác định công thức đơn giản nhất của các hợp chất có công thức dưới đây:\na) C~4~H~10~ (butane)\nb) HOCH~2~CH~2~OH (ethane-1,2-diol)\nc) C~6~H~6~ (benzene)\nd) CHCl~2~COOH (dichloroethanoic acid)")
    if is_teacher:
        sol_c1_p4 = (
            "Hướng dẫn giải:\n"
            "• a) Butane: CTPT là C~4~H~10~. Tỉ lệ C : H = 4 : 10 = 2 : 5 \u2192 Công thức đơn giản nhất là C~2~H~5~.\n"
            "• b) Ethane-1,2-diol: CTPT là C~2~H~6~O~2~. Tỉ lệ C : H : O = 2 : 6 : 2 = 1 : 3 : 1 \u2192 Công thức đơn giản nhất là CH~3~O.\n"
            "• c) Benzene: CTPT là C~6~H~6~. Tỉ lệ C : H = 6 : 6 = 1 : 1 \u2192 Công thức đơn giản nhất là CH.\n"
            "• d) Dichloroethanoic acid: CTPT là C~2~H~2~Cl~2~O~2~. Tỉ lệ C : H : Cl : O = 2 : 2 : 2 : 2 = 1 : 1 : 1 : 1 \u2192 Công thức đơn giản nhất là CHClO (theo cách viết gom nhóm trong tài liệu gốc là C~2~H~2~O~2~Cl)."
        )
        add_solution_no_table(doc, sol_c1_p4)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Phần IV
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 [SBT - CTST]. Acetic acid được sử dụng rộng rãi trên thế giới trong nhiều ngành công nghiệp khác nhau như tạo ra polymer ứng dụng trong sơn, chất kết dính, là dung môi hoà tan các chất hoá học, sản xuất và bảo quản thực phẩm đặc biệt dùng để sản xuất giấm.\na) Lập công thức phân tử của acetic acid, biết kết quả phân tích nguyên tố của acetic acid có 40% C; 53,33% O về khối lượng; còn lại là H. Phân tử khối của acetic acid được xác định trên phổ khối lượng tương ứng với peak có giá trị m/z lớn nhất.")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau2_p4_ms_acetic_acid.png"), width_cm=10.0)
    p_sub2b = doc.add_paragraph()
    p_sub2b.paragraph_format.space_before = Pt(3)
    p_sub2b.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_sub2b, "b) Dựa vào phổ IR bên, hãy cho biết có thể xác định được nhóm chức carboxyl có trong acetic acid từ peak nào.")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau2_p4_ir_acetic_acid.png"), width_cm=11.5)

    if is_teacher:
        sol_c2_p4 = (
            "Hướng dẫn giải:\n"
            "a) Tính thành phần %H:\n"
            "   %m~H~ = 100% \u2212 40% \u2212 53,33% = 6,67%.\n"
            "   Từ phổ khối lượng (MS), peak ion phân tử [M^+^] có m/z lớn nhất bằng 60 \u2192 M~acetic acid~ = 60.\n"
            "   Đặt công thức phân tử của acetic acid là C~x~H~y~O~z~:\n"
            "   • x = (60 \u00d7 40%) / (12 \u00d7 100%) = 2\n"
            "   • y = (60 \u00d7 6,67%) / (1 \u00d7 100%) \u2248 4\n"
            "   • z = (60 \u00d7 53,33%) / (16 \u00d7 100%) \u2248 2\n"
            "   \u2192 Công thức phân tử của acetic acid là C~2~H~4~O~2~.\n"
            "b) Dựa vào phổ hồng ngoại (IR):\n"
            "   - Peak A (trong khoảng 3300 \u2013 3000 cm^-1^, tù và rộng) đặc trưng cho dao động hóa trị của nhóm O\u2013H liên kết trong nhóm carboxyl.\n"
            "   - Peak C (ở khoảng 1700 cm^-1^, cường độ mạnh) đặc trưng cho dao động hóa trị của liên kết C=O.\n"
            "   Sự hiện diện đồng thời của peak A và peak C giúp xác định sự có mặt của nhóm chức carboxyl (\u2013COOH) trong phân tử acetic acid."
        )
        add_solution_no_table(doc, sol_c2_p4)
    else:
        add_student_writing_lines(doc, num_lines=12)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(12)
    run_end = p_end.add_run("---------- HẾT ----------")
    run_end.bold = True
    run_end.font.color.rgb = COLOR_PRIMARY

    filename = "De_Kiem_Tra_Hoa_11_De08_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_11_De08.docx"
    filepath = os.path.join(OUT_DIR, filename)
    doc.save(filepath)
    print(f" Đã tạo thành công: {filepath}")

if __name__ == "__main__":
    build_exam_document(is_teacher=False)
    build_exam_document(is_teacher=True)
