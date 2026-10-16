# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file Word cho Đề Số 10 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De10.docx (Bản Học Sinh)
- De_Kiem_Tra_Hoa_11_De10_GiaoVien.docx (Bản Giáo Viên)
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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de10")

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
        duration="45 phút",
        exam_code="110",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; F = 19; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Br = 80; Ag = 108; Ba = 137",
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
    render_question_prompt(p, "Câu 1. Chất nào sau đây là pentane?")
    opts_c1 = [
        "A. CH~4~.",
        "*B. CH~3~[CH~2~]~3~CH~3~.*",
        "C. CH~3~[CH~2~]~2~CH~3~.",
        "D. CH~3~CH~2~CH~3~."
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Pentane là alkane no, mạch hở chứa 5 nguyên tử carbon (C~5~H~12~), có công thức cấu tạo thu gọn là CH~3~[CH~2~]~3~CH~3~ (hay CH~3~-CH~2~-CH~2~-CH~2~-CH~3~).")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Phát biểu nào sau đây không đúng khi nói về tính chất vật lý của alkane?")
    opts_c2 = [
        "A. Các alkane đều nhẹ hơn nước.",
        "B. Các alkane không tan trong nước.",
        "C. Ở điều kiện thường, các alkane từ C~1~ đến C~4~ ở thể khí, từ C~5~ đến khoảng C~17~ ở thể lỏng, từ khoảng C~18~ trở lên ở thể rắn.",
        "*D. Các alkane không tan hoặc tan rất ít trong các dung môi hữu cơ.*"
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Alkane là những phân tử không phân cực nên tan tốt trong các dung môi hữu cơ không phân cực hoặc kém phân cực như benzene, ether, chloroform,... Do đó, phát biểu D \"không tan hoặc tan rất ít trong các dung môi hữu cơ\" là không đúng.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Độ dài liên kết bội giữa hai nguyên tử carbon luôn ... độ dài liên kết đơn giữa hai nguyên tử carbon đó. Cụm từ thích hợp điền vào chỗ trống là")
    opts_c3 = [
        "*A. ngắn hơn.*",
        "B. dài hơn.",
        "C. bằng nhau.",
        "D. gấp đôi."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong liên kết bội (liên kết đôi C=C hoặc liên kết ba C≡C), mật độ electron xen phủ giữa hai nguyên tử carbon lớn hơn liên kết đơn C-C, lực hút tĩnh điện giữa hạt nhân và mây electron liên kết mạnh hơn kéo hai nguyên tử carbon lại gần nhau hơn, nên độ dài liên kết bội luôn ngắn hơn độ dài liên kết đơn.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Thành phần chính của khí biogas là")
    opts_c4 = [
        "A. ethane.",
        "B. propane.",
        "*C. methane.*",
        "D. butane."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Methane (CH~4~) chiếm khoảng 50 - 70% thể tích của khí sinh học (biogas), được sinh ra từ quá trình phân hủy kị khí các chất hữu cơ, chất thải chăn nuôi.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Alkene, alkyne là những hydrocarbon không no, mạch hở, trong phân tử có chứa")
    opts_c5 = [
        "A. liên kết π.",
        "B. liên kết đơn.",
        "*C. liên kết bội.*",
        "D. vòng benzene."
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Alkene chứa 1 liên kết đôi C=C, alkyne chứa 1 liên kết ba C≡C trong phân tử. Các liên kết đôi và liên kết ba này được gọi chung là liên kết bội.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Bảng dưới đây thể hiện nhiệt độ sôi của một số chất:")
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.left_indent = Cm(0.6)
    p_sub.paragraph_format.space_before = Pt(1)
    p_sub.paragraph_format.space_after = Pt(2)
    render_rich_text(p_sub, "+ Ethene (chất Y): nhiệt độ sôi là -103,7 °C\n+ Propene (chất X): nhiệt độ sôi là -47,7 °C\n+ But-1-ene (chất Z): nhiệt độ sôi là -6,3 °C\n+ Pent-1-ene (chất T): nhiệt độ sôi là 30 °C\nỞ điều kiện thường (25 °C), chất tồn tại ở trạng thái lỏng là")
    opts_c6 = [
        "A. Y.",
        "B. Z.",
        "C. X.",
        "*D. T.*"
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Ở điều kiện thường (25 °C), chất T (pent-1-ene) có nhiệt độ sôi bằng 30 °C > 25 °C nên tồn tại ở thể lỏng. Các chất X, Y, Z có nhiệt độ sôi thấp hơn 25 °C nên ở thể khí.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 7. Chất nào sau đây làm mất màu dung dịch bromine ở điều kiện thường?")
    opts_c7 = [
        "A. Butane.",
        "*B. But-1-ene.*",
        "C. Isobutane.",
        "D. Propane."
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "But-1-ene là alkene chứa liên kết đôi C=C kém bền nên dễ dàng tham gia phản ứng cộng làm mất màu dung dịch nước bromine ở nhiệt độ phòng:\nCH~2~=CH-CH~2~-CH~3~ + Br~2~ → CH~2~Br-CHBr-CH~2~-CH~3~.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 8. Khi cho carbide kim loại tác dụng với nước thì carbide kim loại nào sau đây không thể tạo ra acetylene (C~2~H~2~)?")
    opts_c8 = [
        "A. CaC~2~.",
        "B. Mg~2~C~3~.",
        "*C. Al~4~C~3~.*",
        "D. Na~2~C~2~."
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nhôm carbide khi tác dụng với nước sinh ra khí methane (CH~4~), không tạo ra acetylene (C~2~H~2~):\nAl~4~C~3~ + 12H~2~O → 4Al(OH)~3~ + 3CH~4~↑.\nTrong khi đó, calci carbide (CaC~2~) và natri acetylide (Na~2~C~2~) thủy phân tạo acetylene (C~2~H~2~): CaC~2~ + 2H~2~O → Ca(OH)~2~ + C~2~H~2~↑.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 9. Chất nào sau đây thuộc dãy đồng đẳng của benzene?")
    opts_c9 = [
        "A. C~6~H~5~CH=CH~2~.",
        "B. C~10~H~8~.",
        "C. C~6~H~5~-C~6~H~5~.",
        "*D. C~8~H~10~.*"
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Dãy đồng đẳng của benzene (alkylbenzene) có công thức chung là C~n~H~2n-6~ (n ≥ 6). Với n = 8 ta có C~8~H~10~ (ethylbenzene hoặc các đồng phân dimethylbenzene/xylene). Còn C~6~H~5~CH=CH~2~ (styrene) có nhánh không no; C~10~H~8~ (naphthalene) gồm 2 vòng ngưng tụ.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 10. Cho các hydrocarbon sau: benzene, toluene, naphthalene, ethylene, acetylene, methane. Số chất ở thể lỏng ở điều kiện thường là")
    opts_c10 = [
        "A. 1.",
        "*B. 2.*",
        "C. 3.",
        "D. 4."
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Ở điều kiện thường:\n- Thể khí: methane, ethylene, acetylene.\n- Thể lỏng: benzene, toluene.\n- Thể rắn: naphthalene.\nVậy có 2 chất ở thể lỏng (benzene và toluene).")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 11 (SBT - KNTT). Nitro hoá benzene bằng hỗn hợp HNO~3~ đặc và H~2~SO~4~ đặc ở nhiệt độ ≤ 50 °C, tạo thành chất hữu cơ X. Phát biểu nào sau đây về X không đúng?")
    opts_c11 = [
        "A. Tên của X là nitrobenzene.",
        "B. X là chất lỏng, sánh như dầu.",
        "C. X có màu vàng.",
        "*D. X tan tốt trong nước.*"
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Chất X là nitrobenzene (C~6~H~5~NO~2~). Nitrobenzene là chất lỏng sánh màu vàng nhạt, mùi hạnh nhân đặc trưng, không tan trong nước và nặng hơn nước (chìm xuống đáy ống nghiệm). Do đó phát biểu D \"X tan tốt trong nước\" là sai.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 12 (SBT - CD). Nhỏ 1 mL nước bromine vào ống nghiệm đựng 1 mL hexane, chiếu sáng và lắc đều. Hiện tượng quan sát được là")
    opts_c12 = [
        "A. trong ống nghiệm có chất lỏng đồng nhất.",
        "*B. màu của nước bromine bị mất.*",
        "C. màu của bromine không thay đổi.",
        "D. trong ống nghiệm xuất hiện kết tủa."
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Dưới tác dụng của ánh sáng, alkane (hexane) tham gia phản ứng thế halogen với bromine: C~6~H~14~ + Br~2~ —as→ C~6~H~13~Br + HBr. Phản ứng làm tiêu hao bromine nên màu vàng cam của nước bromine bị mất.")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 13 (SBT - KNTT). Arene hay còn gọi là hydrocarbon thơm là những hydrocarbon trong phân tử có chứa một hay nhiều")
    opts_c13 = [
        "*A. vòng benzene.*",
        "B. liên kết đơn.",
        "C. liên kết đôi.",
        "D. liên kết ba."
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo định nghĩa, hydrocarbon thơm (arene) là những hydrocarbon trong phân tử chứa một hay nhiều vòng benzene.")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 14 (SBT - CTST). Cho alkane có công thức cấu tạo:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "hoa11_de10_cau14_alkane.png"), width_cm=4.2)
    p_c14 = doc.add_paragraph()
    p_c14.paragraph_format.space_before = Pt(2)
    p_c14.paragraph_format.space_after = Pt(2)
    render_rich_text(p_c14, "Danh pháp thay thế của alkane trên là")
    opts_c14 = [
        "A. 2-ethyl-3-methylbutane.",
        "B. 2-methyl-3-ethylbutane.",
        "C. 3,4-dimethylpentane.",
        "*D. 2,3-dimethylpentane.*"
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Mạch chính dài nhất chứa 5 nguyên tử carbon (pentane). Đánh số mạch carbon từ đầu gần nhánh hơn (từ phải qua trái và đi xuống nhánh ethyl):\nC^5^H~3~-C^4^H~2~-C^3^H(CH~3~)-C^2^H(CH~3~)-C^1^H~3~.\nTại vị trí C2 và C3 có hai nhóm thế methyl (-CH~3~). Vậy tên thay thế đúng là 2,3-dimethylpentane.")

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 15 (SBT - CD). Cho pent-2-ene phản ứng với dung dịch KMnO~4~ ở nhiệt độ phòng có thể thu được sản phẩm hữu cơ có công thức cấu tạo nào sau đây?")
    opts_c15 = [
        "*A. CH~3~CH~2~CH(OH)CH(OH)CH~3~.*",
        "B. CH~3~CH~2~CH~2~CH(OH)CH~3~.",
        "C. CH~3~CH(OH)CH~2~CH(OH)CH~3~.",
        "D. CH~3~CH~2~CH(OH)CH~2~CH~3~."
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Dung dịch KMnO~4~ ở nhiệt độ phòng oxi hóa nhẹ alkene pent-2-ene thành diol liền kề (glycol), gắn 2 nhóm -OH vào hai nguyên tử C mang nối đôi C=C:\n3CH~3~-CH=CH-CH~2~-CH~3~ + 2KMnO~4~ + 4H~2~O → 3CH~3~-CH(OH)-CH(OH)-CH~2~-CH~3~ + 2MnO~2~↓ + 2KOH.")

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 16. Tính chất nào không phải của benzene?")
    opts_c16 = [
        "A. Tác dụng với Br~2~ (t°, Fe).",
        "B. Tác dụng với HNO~3~ (đ) / H~2~SO~4~ (đ).",
        "*C. Tác dụng với dung dịch KMnO~4~.*",
        "D. Tác dụng với Cl~2~ (as)."
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Vòng benzene có hệ liên hợp electron π khép kín rất bền vững nên không làm mất màu dung dịch KMnO~4~ ngay cả khi đun nóng. Các alkylbenzene như toluene mới làm mất màu dung dịch KMnO~4~ khi đun sôi.")

    # Câu 17
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 17 (SBT - KNTT). Cho các hydrocarbon sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "hoa11_de10_cau17_hydrocarbons.png"), width_cm=11.5)
    p_c17 = doc.add_paragraph()
    p_c17.paragraph_format.space_before = Pt(2)
    p_c17.paragraph_format.space_after = Pt(2)
    render_rich_text(p_c17, "Một số nhận định về các hydrocarbon trên là:\n(1) Số phân tử hydrocarbon không no bằng 5;\n(2) Số phân tử alkene bằng 3;\n(3) Số phân tử alkyne bằng 2;\n(4) Số phân tử thuộc dãy đồng đẳng của benzene bằng 3.\nTrong các nhận định này, số nhận định đúng bằng")
    opts_c17 = [
        "A. 1.",
        "B. 2.",
        "C. 3.",
        "*D. 4.*"
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tích 11 chất:\n- Hàng 1 (mạch hở): but-2-ene (alkene), 2-methylbutane (alkane no), 2-methylbut-2-ene (alkene), but-1-yne (alkyne), but-1-ene (alkene), but-2-yne (alkyne).\n- Hàng 2 (vòng thơm): benzene, toluene, p-xylene (3 chất đồng đẳng benzene C~n~H~2n-6~); styrene (arene có nhánh alkene); naphthalene (arene 2 vòng ngưng tụ).\nKiểm tra các nhận định:\n(1) Hydrocarbon không no gồm 5 chất: but-2-ene, 2-methylbut-2-ene, but-1-yne, but-1-ene, but-2-yne → Đúng.\n(2) Alkene gồm 3 chất: but-2-ene, 2-methylbut-2-ene, but-1-ene → Đúng.\n(3) Alkyne gồm 2 chất: but-1-yne, but-2-yne → Đúng.\n(4) Đồng đẳng benzene gồm 3 chất: benzene, toluene, p-xylene → Đúng.\nCả 4 nhận định đều đúng → Chọn D.")

    # Câu 18
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 18. Benzene gây tác hại lên tủy xương và làm giảm lượng hồng cầu, dẫn đến thiếu máu. Benzene cũng có thể ảnh hưởng đến hệ thống miễn dịch, làm tăng nguy cơ nhiễm trùng. Khi hút mỗi điếu thuốc lá, người hút đưa vào cơ thể 50 µg benzene. Nếu một người hút 15 điếu thuốc lá mỗi ngày thì lượng benzene người đó hấp thụ vào cơ thể là bao nhiêu g?")
    opts_c18 = [
        "A. 7,5.10^-5^.",
        "B. 15.10^-5^.",
        "*C. 7,5.10^-4^.*",
        "D. 15.10^-4^."
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Đổi: 1 µg = 10^-6^ g.\nLượng benzene người đó hấp thụ mỗi ngày từ 15 điếu thuốc lá là:\nm = 15 × 50 µg = 750 µg = 750 × 10^-6^ g = 7,5.10^-4^ g.")

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
    render_question_prompt(p, "Câu 1. Dựa vào đặc điểm cấu tạo phân tử của các alkane thì:")
    stmts_c1_p2 = [
        ("a) Alkane là những hydrocarbon mạch hở chỉ chứa liên kết đơn trong phân tử nên được gọi là hydrocarbon bão hoà và có công thức chung C~n~H~2n+2~ (n ≥ 1).", "ĐÚNG"),
        ("b) Alkane là những hydrocarbon chỉ chứa liên kết đơn trong phân tử nên được gọi là hydrocarbon bão hoà và có công thức chung C~n~H~2n+2~ (n ≥ 2).", "SAI"),
        ("c) Phân tử của một alkane trong sáp nến có 52 nguyên tử hydrogen, thì số nguyên tử carbon trong phân tử alkane là 25.", "ĐÚNG"),
        ("d) Cho hai chất: (1) CH~3~-CH~2~-CH~3~ ; (2) CH~3~-CH(CH~3~)-CH~3~. Alkane (1) là alkane có mạch phân nhánh.", "SAI")
    ]
    for stmt, ans in stmts_c1_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c1_p2 = (
            "a) Đúng: Đây là định nghĩa và công thức phân tử chung chuẩn xác của alkane.\n"
            "b) Sai: Điều kiện n ≥ 1 (chất đầu dãy là CH~4~ có n = 1) và định nghĩa phải nêu rõ là hydrocarbon mạch hở.\n"
            "c) Đúng: Với alkane C~n~H~2n+2~, số H = 2n + 2 = 52 ⇒ 2n = 50 ⇒ n = 25 carbon.\n"
            "d) Sai: Alkane (1) propane là alkane mạch không phân nhánh; alkane (2) isobutane mới là alkane có mạch nhánh."
        )
        add_solution_no_table(doc, sol_c1_p2)

    # Câu 2 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Vận dụng các kiến thức về alkene, alkyne để giải thích được các vai trò quan trọng của alkene và alkyne trong công nghiệp hóa học, công nghiệp dược phẩm, vật liệu...")
    stmts_c2_p2 = [
        ("a) Alkene là nguyên liệu để tổng hợp một số polymer như polyethylene (PE); polypropylene (PP)... đây là các polymer được dùng để làm bao bì và các đồ dùng như ly, cốc, tủ chứa đồ...", "ĐÚNG"),
        ("b) Trong công nghiệp dược phẩm, các alkyne là nguyên liệu sản xuất một số dược phẩm như thuốc kháng virus Efavirenz, thuốc kháng nấm Terbinafine,...", "ĐÚNG"),
        ("c) Alkyne được dùng làm nguyên liệu trong tổng hợp hữu cơ.", "ĐÚNG"),
        ("d) Acetylene là một alkyne được dùng để tổng hợp các vật liệu polymer như poly(vinyl alcohol), poly(vinyl acetate),... dùng làm bao bì, keo dán, màng đệm,...", "ĐÚNG")
    ]
    for stmt, ans in stmts_c2_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c2_p2 = (
            "a) Đúng: Trùng hợp ethylene tạo PE, trùng hợp propylene tạo PP dùng làm màng bọc, chai nhựa, đồ gia dụng.\n"
            "b) Đúng: Cấu trúc liên kết ba alkyne xuất hiện trong hoạt chất của thuốc chống HIV Efavirenz và thuốc trị nấm Terbinafine.\n"
            "c) Đúng: Alkyne là tiền chất tổng hợp nhiều hợp chất hữu cơ quan trọng trong công nghiệp hóa chất.\n"
            "d) Đúng: Từ acetylene điều chế vinyl acetate, rồi trùng hợp thành poly(vinyl acetate) (PVAc) và thủy phân thành poly(vinyl alcohol) (PVA)."
        )
        add_solution_no_table(doc, sol_c2_p2)

    # Câu 3 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Ở điều kiện thường: benzene, toluene, xylene là chất lỏng không màu, trong suốt, dễ cháy, có mùi đặc trưng; naphthalene là chất rắn màu trắng, có mùi đặc trưng.")
    stmts_c3_p2 = [
        ("a) Toluene (C~6~H~5~CH~3~) không tác dụng được với nước bromine và dung dịch thuốc tím ở điều kiện thường.", "ĐÚNG"),
        ("b) Styrene (C~6~H~5~CH=CH~2~) tác dụng được với nước bromine, làm mất màu dung dịch thuốc tím ở điều kiện thường.", "ĐÚNG"),
        ("c) Naphthalene (C~10~H~8~) tác dụng được với nước bromine, làm mất màu dung dịch thuốc tím ở điều kiện thường.", "SAI"),
        ("d) Nitro hoá toluene tạo thành hỗn hợp hai sản phẩm chính là ortho và meta – nitrotoluene.", "SAI")
    ]
    for stmt, ans in stmts_c3_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c3_p2 = (
            "a) Đúng: Toluene không làm mất màu nước bromine và dung dịch KMnO~4~ ở điều kiện thường, chỉ bị oxi hóa nhóm methyl khi đun nóng.\n"
            "b) Đúng: Nhánh vinyl -CH=CH~2~ không no có phản ứng cộng làm mất màu nước bromine và bị KMnO~4~ oxi hóa ngay ở nhiệt độ thường.\n"
            "c) Sai: Naphthalene có hệ vòng thơm bền vững, không phản ứng với nước bromine và dung dịch KMnO~4~ ở điều kiện thường.\n"
            "d) Sai: Nhóm methyl (-CH~3~) là nhóm đẩy electron định hướng thế ưu tiên vào vị trí ortho và para, tạo hai sản phẩm chính là ortho-nitrotoluene và para-nitrotoluene."
        )
        add_solution_no_table(doc, sol_c3_p2)

    # Câu 4 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Trong một phương pháp tổng hợp polyethylene (PE), các phân tử ethylene đã được hòa tan trong dung môi phản ứng với nhau để tạo thành polymer.")
    stmts_c4_p2 = [
        ("a) Không thể sử dụng methyl alcohol, nước, cyclohexane hay hex-1-ene làm dung môi cho phản ứng trùng hợp PE. Vì ở nhiệt độ cao, PE có thể hòa tan trong nước, trong các loại alcohol, cyclohexane hay hex-1-ene.", "SAI"),
        ("b) Có thể sử dụng methyl alcohol, nước, cyclohexane hay hex-1-ene làm dung môi cho phản ứng trùng hợp PE. Vì ở nhiệt độ cao, PE cũng không thể hòa tan trong nước, trong các loại rượu, cyclohexane hay hex-1-ene.", "ĐÚNG"),
        ("c) Có thể sử dụng methyl alcohol, nước, nhưng không dùng cyclohexane hay hex-1-ene làm dung môi cho phản ứng trùng hợp PE. Vì ở nhiệt độ cao, PE cũng không thể hòa tan trong nước, trong các loại alcohol, nhưng hòa tan trong cyclohexane hay hex-1-ene.", "SAI"),
        ("d) Không thể sử dụng methyl alcohol, nước, cyclohexane hay hex-1-ene làm dung môi cho phản ứng trùng hợp PE. Vì ở nhiệt độ cao, PE cũng không thể hòa tan trong nước, trong các loại rượu, cyclohexane hay hex-1-ene.", "SAI")
    ]
    for stmt, ans in stmts_c4_p2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_c4_p2 = (
            "Theo đáp án chính thức của đề kiểm tra:\n"
            "- Phát biểu b là phương án ĐÚNG: Có thể sử dụng các dung môi methyl alcohol, nước, cyclohexane hoặc hex-1-ene cho phản ứng trùng hợp ethylene tạo PE.\n"
            "- Các nhận định a, c, d là SAI."
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
    render_question_prompt(p, "Câu 1. Cho các chất sau: C~2~H~2~; C~2~H~6~; C~3~H~6~; C~3~H~8~; C~4~H~8~; C~4~H~9~Cl; C~9~H~12~; C~5~H~12~. Trong số các chất trên, số chất thuộc dãy đồng đẳng alkane là bao nhiêu?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: 3")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        add_solution_no_table(doc, "Dãy đồng đẳng alkane có công thức phân tử chung là C~n~H~2n+2~ (n ≥ 1). Trong các chất đã cho, các chất thỏa mãn công thức này gồm:\n- C~2~H~6~ (ethane, n = 2)\n- C~3~H~8~ (propane, n = 3)\n- C~5~H~12~ (pentane, n = 5)\nVậy có 3 chất thuộc dãy đồng đẳng alkane.")

    # Câu 2 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Có bao nhiêu alkyne ứng với công thức phân tử C~5~H~8~?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: 3")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        add_solution_no_table(doc, "Các đồng phân alkyne mạch hở chứa 1 liên kết ba C≡C của C~5~H~8~ là:\n(1) CH≡C-CH~2~-CH~2~-CH~3~ : pent-1-yne\n(2) CH~3~-C≡C-CH~2~-CH~3~ : pent-2-yne\n(3) CH≡C-CH(CH~3~)-CH~3~ : 3-methylbut-1-yne\nVậy có 3 đồng phân alkyne.")

    # Câu 3 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Cho tên các hydrocarbon thơm: benzene, 1,2-dimethylbenzene, o-dimethylbenzene, o-xylene, ethylbenzene. Có bao nhiêu chất trong các tên này?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: 3")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        add_solution_no_table(doc, "Xác định các chất tương ứng:\n- benzene: 1 chất (C~6~H~6~).\n- 1,2-dimethylbenzene, o-dimethylbenzene, o-xylene: cùng là các tên gọi khác nhau của 1 chất duy nhất (1,2-(CH~3~)~2~C~6~H~4~).\n- ethylbenzene: 1 chất (C~6~H~5~-CH~2~CH~3~).\nVậy có tất cả 3 chất phân biệt.")

    # Câu 4 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Propene là nguyên liệu cho sản xuất nhựa polypropylene (PP). PP được sử dụng để sản xuất các sản phẩm ống, màng, dây cách điện, kéo sợi, đồ gia dụng và các sản phẩm tạo hình khác:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "hoa11_de10_cau4_pp_products.png"), width_cm=10.5)
    p_c4_sub = doc.add_paragraph()
    p_c4_sub.paragraph_format.space_before = Pt(2)
    p_c4_sub.paragraph_format.space_after = Pt(2)
    render_rich_text(p_c4_sub, "Phản ứng tạo thành propene từ propyne:\nCH~3~-C≡CH(g) + H~2~(g) —t°, Pd/PbCO~3~→ CH~3~-CH=CH~2~(g)\nTừ năng lượng liên kết (E~b~, kJ/mol) của H-H là 432; của C-C là 347; của C-H là 413; của C≡C là 839; của C=C là 614, hãy tính biến thiên enthalpy chuẩn (Δ~r~H°~298~) của phản ứng tạo thành propene trên theo đơn vị kJ.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: -169 (hoặc -169 kJ)")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        sol_c4_p3 = (
            "Áp dụng công thức tính biến thiên enthalpy chuẩn theo năng lượng liên kết:\n"
            "Δ~r~H°~298~ = ∑E~b~(chất đầu) - ∑E~b~(sản phẩm)\n"
            "= [E~b~(C≡C) + E~b~(C-C) + 4E~b~(C-H) + E~b~(H-H)] - [E~b~(C=C) + E~b~(C-C) + 6E~b~(C-H)]\n"
            "= [E~b~(C≡C) + E~b~(H-H)] - [E~b~(C=C) + 2E~b~(C-H)]\n"
            "= (839 + 432) - (614 + 2 × 413)\n"
            "= 1271 - 1440 = -169 kJ."
        )
        add_solution_no_table(doc, sol_c4_p3)

    # Câu 5 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5 (SBT - CTST). Cục Quản Lí Thực Phẩm và Dược Phẩm Hoa Kì (FDA) đã công nhận ethylene là an toàn trong việc kích thích trái cây mau chín. Tuy nhiên khi vượt quá nồng độ cho phép, ví dụ đối với nồng độ 27 000 ppm, tức gấp khoảng 200 lần mức cần thiết để kích thích quá trình chín, một tia lửa điện có thể đốt cháy ethylene và gây ra vụ nổ chết người.\nTrong phòng ủ chín, ethylene được sử dụng ở nồng độ 100 ppm – 150 ppm. Khối lượng ethylene cần thiết (tính theo gam, làm tròn đến một chữ số thập phân) sử dụng để phòng ủ chín có thể tích 50 m³ đạt nồng độ 140 ppm ở 25 °C và 1 bar là bao nhiêu?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: 7,9 (hoặc 7,9 gam)")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        sol_c5_p3 = (
            "Thể tích phòng ủ: V = 50 m³ = 50 000 L.\n"
            "Thể tích khí ethylene ở nồng độ 140 ppm (1 ppm = 1/1 000 000) là:\n"
            "V_{C~2~H~4~} = (50 000 × 140) / 1 000 000 = 7 L.\n"
            "Ở điều kiện chuẩn (25 °C, 1 bar), thể tích 1 mol khí là 24,79 L/mol:\n"
            "Số mol ethylene là: n_{C~2~H~4~} = 7 / 24,79 mol.\n"
            "Khối lượng ethylene cần dùng là:\n"
            "m = (7 / 24,79) × 28 ≈ 7,9 g."
        )
        add_solution_no_table(doc, sol_c5_p3)

    # Câu 6 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Khí biogas (giả thiết chỉ chứa CH~4~) và khí gas (chứa 40% C~3~H~8~ và 60% C~4~H~10~ về thể tích) được dùng phổ biến làm nhiên liệu và đun nấu. Nhiệt lượng tỏa ra khi đốt cháy hoàn toàn 1 mol các chất như sau:\n+ CH~4~: 890 kJ\n+ C~3~H~8~: 2220 kJ\n+ C~4~H~10~: 2850 kJ\nNếu nhu cầu về năng lượng không đổi, hiệu suất sử dụng các loại nhiên liệu như nhau, khi dùng khí biogas để thay thế khí gas để làm nhiên liệu đốt cháy thì lượng khí CO~2~ thải ra môi trường sẽ giảm bao nhiêu phần trăm (làm tròn đến hai chữ số thập phân)?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.6)
        r = p_ans.add_run("Đáp số: 18,91% (hoặc 18,91)")
        set_run_style(r, bold=True, color=COLOR_CORRECT)
        sol_c6_p3 = (
            "Gọi Q (kJ) là nhiệt lượng cần cung cấp cho quá trình đun nấu.\n"
            "+ Nếu dùng biogas (CH~4~):\n"
            "  n_{CH~4~} = Q / 890 mol ⇒ n_{CO~2~ phát thải} = n_{CH~4~} = Q / 890 mol.\n"
            "+ Nếu dùng khí gas:\n"
            "  Tỉ lệ thể tích (tỉ lệ số mol) C~3~H~8~ và C~4~H~10~ là 40% : 60% = 2 : 3.\n"
            "  Đặt n_{C~3~H~8~} = 2x mol và n_{C~4~H~10~} = 3x mol.\n"
            "  Q = 2220 × (2x) + 2850 × (3x) = 4440x + 8550x = 12990x ⇒ x = Q / 12990 mol.\n"
            "  n_{CO~2~ phát thải} = 3 × (2x) + 4 × (3x) = 18x = 18 × (Q / 12990) = 3Q / 2165 mol.\n"
            "+ So sánh:\n"
            "  Vì 3Q / 2165 ≈ Q / 721,67 > Q / 890 nên dùng biogas phát thải ít CO~2~ hơn gas.\n"
            "  Phần trăm lượng CO~2~ giảm đi là:\n"
            "  % giảm = [(3Q / 2165 - Q / 890) / (3Q / 2165)] × 100% ≈ 18,91%."
        )
        add_solution_no_table(doc, sol_c6_p3)

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
    render_question_prompt(p, "Câu 1 (1,0 điểm) [CD - SGK]. Cho biết công thức cấu tạo và gọi tên của alkene, alkyne mà khi hydrogen hóa tạo thành butane.")

    if is_teacher:
        sol_c1_p4 = (
            "Butane (C~4~H~10~) có mạch carbon không phân nhánh gồm 4 nguyên tử carbon: CH~3~-CH~2~-CH~2~-CH~3~.\n"
            "Để khi hydrogen hóa hoàn toàn thu được butane, alkene và alkyne tham gia phản ứng phải có cùng khung carbon không phân nhánh 4C:\n"
            "1. Các alkene:\n"
            "- CH~2~=CH-CH~2~-CH~3~: but-1-ene\n"
            "  PTHH: CH~2~=CH-CH~2~-CH~3~ + H~2~ —Ni, t°→ CH~3~-CH~2~-CH~2~-CH~3~\n"
            "- CH~3~-CH=CH-CH~3~: but-2-ene\n"
            "  PTHH: CH~3~-CH=CH-CH~3~ + H~2~ —Ni, t°→ CH~3~-CH~2~-CH~2~-CH~3~\n"
            "2. Các alkyne:\n"
            "- CH≡C-CH~2~-CH~3~: but-1-yne\n"
            "  PTHH: CH≡C-CH~2~-CH~3~ + 2H~2~ —Ni, t°→ CH~3~-CH~2~-CH~2~-CH~3~\n"
            "- CH~3~-C≡C-CH~3~: but-2-yne\n"
            "  PTHH: CH~3~-C≡C-CH~3~ + 2H~2~ —Ni, t°→ CH~3~-CH~2~-CH~2~-CH~3~\n"
            "(Lưu ý: 2-methylpropene CH~2~=C(CH~3~)-CH~3~ khi hydrogen hóa tạo isobutane, không tạo butane)."
        )
        add_solution_no_table(doc, sol_c1_p4)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Phần IV
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 (1,0 điểm). Trong ngành công nghệ lọc hóa dầu, các alkane thường được loại bỏ hydrogen trong các phản ứng dehydro hóa để tạo ra những sản phẩm hydrocarbon không no có nhiều ứng dụng trong công nghiệp:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "hoa11_de10_cau2_oil_refinery.png"), width_cm=10.5)

    p_c2_sub = doc.add_paragraph()
    p_c2_sub.paragraph_format.space_before = Pt(2)
    p_c2_sub.paragraph_format.space_after = Pt(2)
    render_rich_text(p_c2_sub, "Hãy tính biến thiên enthalpy chuẩn của các phản ứng sau dựa vào năng lượng liên kết:\n+ Cho năng lượng liên kết E~b~ (kJ/mol): C-H: 414; C-C: 347 (hoặc 237 theo số liệu đề); C=C: 611; H-H: 436.\na) H~3~C-CH~2~-CH~2~-CH~3~ → CH~2~=CH-CH=CH~2~ + 2H~2~\nb) 6CH~4~ → C~6~H~6~ (1,3,5-cyclohexatriene) + 9H~2~\nCho biết công thức cấu tạo của 1,3,5-cyclohexatriene như sau:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "hoa11_de10_cau2_cyclohexatriene.png"), width_cm=3.5)

    p_c2_ask = doc.add_paragraph()
    p_c2_ask.paragraph_format.space_before = Pt(2)
    p_c2_ask.paragraph_format.space_after = Pt(2)
    render_rich_text(p_c2_ask, "Các phản ứng trên có thuận lợi về phương diện nhiệt hay không? Phản ứng theo chiều ngược lại có biến thiên enthalpy bằng bao nhiêu?")

    if is_teacher:
        sol_c2_p4 = (
            "a) H~3~C-CH~2~-CH~2~-CH~3~ → CH~2~=CH-CH=CH~2~ + 2H~2~\n"
            "Δ~r~H°~298~ = ∑E~b~(chất đầu) - ∑E~b~(sản phẩm)\n"
            "= [10E~b~(C-H) + 3E~b~(C-C)] - [6E~b~(C-H) + 2E~b~(C=C) + E~b~(C-C) + 2E~b~(H-H)]\n"
            "= 4E~b~(C-H) + 2E~b~(C-C) - 2E~b~(C=C) - 2E~b~(H-H)\n"
            "- Với E~b~(C-C) = 347 kJ/mol:\n"
            "  Δ~r~H°~298~ = 4 × 414 + 2 × 347 - 2 × 611 - 2 × 436 = 1656 + 694 - 1222 - 872 = 256 kJ.\n"
            "(Ghi chú: Nếu dùng số liệu in trong bảng đề E~b~(C-C) = 237 kJ/mol thì Δ~r~H°~298~ = 36 kJ, hướng dẫn giải gốc sử dụng giá trị tính ra 256 kJ).\n\n"
            "b) 6CH~4~ → C~6~H~6~ (1,3,5-cyclohexatriene) + 9H~2~\n"
            "Trong 6 phân tử CH~4~ có 24 liên kết C-H.\n"
            "Trong 1,3,5-cyclohexatriene có 6 liên kết C-H, 3 liên kết đơn C-C và 3 liên kết đôi C=C.\n"
            "Trong 9 phân tử H~2~ có 9 liên kết H-H.\n"
            "Δ~r~H°~298~ = 24E~b~(C-H) - [6E~b~(C-H) + 3E~b~(C-C) + 3E~b~(C=C) + 9E~b~(H-H)]\n"
            "= 18E~b~(C-H) - 3E~b~(C-C) - 3E~b~(C=C) - 9E~b~(H-H)\n"
            "= 18 × 414 - 3 × 347 - 3 × 611 - 9 × 436\n"
            "= 7452 - 1041 - 1833 - 3924 = 654 kJ.\n\n"
            "Đánh giá về phương diện nhiệt:\n"
            "- Cả hai phản ứng trên đều có Δ~r~H°~298~ > 0 (phản ứng thu nhiệt mạnh) nên không thuận lợi về phương diện nhiệt ở điều kiện thường.\n"
            "- Phản ứng theo chiều ngược lại (phản ứng tỏa nhiệt) thuận lợi về mặt nhiệt động học, với biến thiên enthalpy tương ứng bằng:\n"
            "  + CH~2~=CH-CH=CH~2~ + 2H~2~ → H~3~C-CH~2~-CH~2~-CH~3~ : Δ~r~H°~298~ = -256 kJ\n"
            "  + C~6~H~6~ + 9H~2~ → 6CH~4~ : Δ~r~H°~298~ = -654 kJ."
        )
        add_solution_no_table(doc, sol_c2_p4)
    else:
        add_student_writing_lines(doc, num_lines=10)

    # Save
    filename = "De_Kiem_Tra_Hoa_11_De10_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_11_De10.docx"
    filepath = os.path.join(OUT_DIR, filename)
    doc.save(filepath)
    print(f"Saved: {filepath}")

if __name__ == "__main__":
    build_exam_document(is_teacher=False)
    build_exam_document(is_teacher=True)
    print("All Docx files for Đề 10 created successfully!")
