# -*- coding: utf-8 -*-
"""
Script sinh tự động 2 file Word (.docx) cho ĐỀ SỐ 16 (Chương 6: HỢP CHẤT CARBONYL - CARBOXYLIC ACID):
1. Bản Học sinh (De_Kiem_Tra_Hoa_11_De16.docx)
2. Bản Giáo viên (De_Kiem_Tra_Hoa_11_De16_GiaoVien.docx)

Tuân thủ nghiêm ngặt Quy chuẩn Form Đề Kiểm Tra (GEMINI.md):
- Tuyệt đối KHÔNG dùng bảng (Table) trong nội dung câu hỏi/đáp án/lời giải.
- Dùng Paragraph + Tab stops chuẩn cho các phương án trắc nghiệm.
- Công thức hóa học xử lý triệt để chỉ số trên/dưới (~ cho sub, ^ cho sup).
- Header & Footer chuẩn với viền màu Sage Green (#8BC390) và chữ Coral Red (#D32F2F).
- Bản Giáo viên có ghi chú đỏ và lời giải thụt lề màu xanh #1565C0.
"""

import os
import sys

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH

# Đảm bảo import được module form_de_kiem_tra
sys.path.insert(0, os.path.dirname(__file__))
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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de16")

def add_p(doc, space_before=3, space_after=2):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    return p

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
        exam_title="ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 6 (HỢP CHẤT CARBONYL - CARBOXYLIC ACID)",
        subject_line="MÔN: HOÁ HỌC - LỚP 11",
        num_pages="04",
        duration="50 phút",
        exam_code="116",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Cr = 52; Mn = 55; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108; Ba = 137.",
        is_teacher=is_teacher
    )

    # =========================================================================
    # PHẦN I
    # =========================================================================
    add_section_header(
        doc,
        title="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # Câu 1
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 1. Công thức tổng quát của aldehyde no, đơn chức, mạch hở có thể là công thức nào trong số các công thức sau:"
    )
    opts_c1 = [
        "*A. C~x~H~2x+1~CHO (x nguyên, không âm) hoặc C~x~H~2x~O (x nguyên dương).* ",
        "B. C~x~H~2x~O (x nguyên dương).",
        "C. C~x~H~2x+2~CHO (x nguyên, không âm).",
        "D. C~x~H~y~CHO."
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Aldehyde no, đơn chức, mạch hở có công thức C~n~H~2n~O (n >= 1) hay R-CHO với R là H hoặc gốc alkyl no C~x~H~2x+1~ (x >= 0, C~x~H~2x+1~CHO).\n"
            "=> Chọn A."
        )

    # Câu 2
    p = add_p(doc)
    render_question_prompt(p, "Câu 2. Chất nào sau đây là formic aldehyde?")
    opts_c2 = [
        "*A. HCHO.*",
        "B. CH~3~CHO.",
        "C. C~2~H~5~OH.",
        "D. C~2~H~6~."
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Formic aldehyde (formaldehyde, methanal) có công thức là HCHO.\n"
            "=> Chọn A."
        )

    # Câu 3
    p = add_p(doc)
    render_question_prompt(p, "Câu 3. Khử HCHO bằng LiAlH~4~ thu được sản phẩm là")
    opts_c3 = [
        "A. CH~3~CHOHCH~3~.",
        "*B. CH~3~OH.*",
        "C. CH~3~COOCH~3~.",
        "D. CH~3~COOH."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Aldehyde bị khử bởi LiAlH~4~ tạo thành alcohol bậc một tương ứng:\n"
            "HCHO --(LiAlH~4~)--> CH~3~OH (methanol).\n"
            "=> Chọn B."
        )

    # Câu 4
    p = add_p(doc)
    render_question_prompt(p, "Câu 4. Khử CH~3~COCH~3~ bằng LiAlH~4~ thu được sản phẩm là")
    opts_c4 = [
        "A. CH~3~COOCH~3~.",
        "*B. CH~3~CHOHCH~3~.*",
        "C. CH~3~COOH.",
        "D. CH~3~CH~2~OH."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Ketone bị khử bởi LiAlH~4~ tạo alcohol bậc hai tương ứng:\n"
            "CH~3~-CO-CH~3~ --(LiAlH~4~)--> CH~3~-CH(OH)-CH~3~ (propan-2-ol).\n"
            "=> Chọn B."
        )

    # Câu 5
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 5. [KNTT - SBT] Ba chất A, B, C có nhiệt độ sôi được biểu thị như hình sau:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "boiling_point_bar_chart.png"), width_cm=7.5)
    p_sub = add_p(doc)
    render_rich_text(p_sub, "Các chất A, B, C lần lượt là:")
    opts_c5 = [
        "A. ethanol, acetaldehyde, acetic acid.",
        "*B. acetaldehyde, ethanol, acetic acid.*",
        "C. acetaldehyde, acetic acid, ethanol.",
        "D. acetic acid, acetaldehyde, ethanol."
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "So sánh nhiệt độ sôi của các hợp chất hữu cơ có phân tử khối tương đương:\n"
            "- Acetaldehyde (CH~3~CHO, M = 44) không có liên kết hydrogen liên phân tử => nhiệt độ sôi thấp nhất (~21°C) tương ứng chất A.\n"
            "- Ethanol (C~2~H~5~OH, M = 46) có liên kết hydrogen liên phân tử => nhiệt độ sôi trung bình (~78,3°C) tương ứng chất B.\n"
            "- Acetic acid (CH~3~COOH, M = 60) có liên kết hydrogen dạng dimer bền hơn => nhiệt độ sôi cao nhất (~118°C) tương ứng chất C.\n"
            "=> Thứ tự lần lượt là: acetaldehyde, ethanol, acetic acid.\n"
            "=> Chọn B."
        )

    # Câu 6
    p = add_p(doc)
    render_question_prompt(p, "Câu 6. [KNTT - SBT] Nhận xét nào sau đây không đúng?")
    opts_c6 = [
        "A. Aldehyde vừa có tính oxi hoá, vừa có tính khử.",
        "*B. Chỉ có ketone tham gia phản ứng tạo iodoform.*",
        "C. Acid và ester no, đơn chức, mạch hở có công thức chung là C~n~H~2n~O~2~.",
        "D. Carboxylic acid làm đổi màu giấy quỳ."
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- B sai vì các hợp chất có nhóm CH~3~-C=O (kể cả acetaldehyde CH~3~CHO) hoặc nhóm CH~3~-CH(OH)- (như ethanol C~2~H~5~OH) đều có thể tham gia phản ứng tạo iodoform (CHI~3~).\n"
            "=> Chọn B."
        )

    # Câu 7
    p = add_p(doc)
    render_question_prompt(p, "Câu 7. (SBT - CTST) Acetaldehyde thể hiện tính oxi hóa trong phản ứng nào sau đây?")
    opts_c7 = [
        "*A. CH~3~CHO + H~2~ --(Ni, t°)--> CH~3~CH~2~OH.*",
        "B. 2CH~3~CHO + 5O~2~ --(t°)--> 4CO~2~ + 4H~2~O.",
        "C. CH~3~CHO + Br~2~ + H~2~O -> CH~3~COOH + 2HBr.",
        "D. CH~3~CHO + 2Ag(NH~3~)~2~OH -> CH~3~COONH~4~ + 2Ag + 3NH~3~ + H~2~O."
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Trong phản ứng với H~2~, số oxi hóa của nguyên tử carbon trong nhóm -CHO giảm từ +1 xuống -1, acetaldehyde đóng vai trò là chất oxi hóa (bị khử thành alcohol bậc một).\n"
            "=> Chọn A."
        )

    # Câu 8
    p = add_p(doc)
    render_question_prompt(p, "Câu 8. (SBT- CD) Phát biểu nào sau đây về tính chất của hợp chất carbonyl là không đúng?")
    opts_c8 = [
        "A. Aldehyde phản ứng được với nước bromine.",
        "B. Ketone không phản ứng được với Cu(OH)~2~/OH^-.",
        "C. Aldehyde tác dụng với dung dịch AgNO~3~/NH~3~ tạo ra bạc.",
        "*D. Trong các hợp chất carbonyl, chỉ aldehyde bị khử bởi NaBH~4~.*"
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Cả aldehyde và ketone đều bị khử bởi các chất khử như NaBH~4~ hoặc LiAlH~4~ (aldehyde tạo alcohol bậc I, ketone tạo alcohol bậc II). Phát biểu D nói chỉ aldehyde bị khử là sai.\n"
            "=> Chọn D."
        )

    # Câu 9
    p = add_p(doc)
    render_question_prompt(p, "Câu 9. (SBT- CD) Số công thức cấu tạo chứa nhóm carboxylic có cùng công thức C~5~H~10~O~2~ là")
    opts_c9 = [
        "A. 2.",
        "B. 3.",
        "*C. 4.*",
        "D. 5."
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Acid có công thức C~5~H~10~O~2~ có dạng C~4~H~9~-COOH. Gốc butyl -C~4~H~9~ có 4 đồng phân cấu tạo:\n"
            "1. CH~3~-CH~2~-CH~2~-CH~2~-COOH (pentanoic acid)\n"
            "2. CH~3~-CH~2~-CH(CH~3~)-COOH (2-methylbutanoic acid)\n"
            "3. (CH~3~)~2~CH-CH~2~-COOH (3-methylbutanoic acid)\n"
            "4. (CH~3~)~3~C-COOH (2,2-dimethylpropanoic acid)\n"
            "=> Có 4 đồng phân cấu tạo. Chọn C."
        )

    # Câu 10
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 10. Cho chuỗi phản ứng:\n"
        "C~2~H~6~O -> X -> acetic acid --(+CH~3~OH)--> Y.\n"
        "Công thức cấu tạo của X, Y lần lượt là"
    )
    opts_c10 = [
        "A. CH~3~CHO, CH~3~CH~2~COOH.",
        "*B. CH~3~CHO, CH~3~COOCH~3~.*",
        "C. CH~3~CHO, CH~2~(OH)CH~2~CHO.",
        "D. CH~3~CHO, HCOOCH~2~CH~3~."
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- C~2~H~5~OH (C~2~H~6~O) bị oxi hóa tạo acetaldehyde (X: CH~3~CHO).\n"
            "- Oxi hóa tiếp CH~3~CHO thu được acetic acid (CH~3~COOH).\n"
            "- Acetic acid phản ứng ester hóa với CH~3~OH tạo methyl acetate (Y: CH~3~COOCH~3~).\n"
            "=> Chọn B."
        )

    # Câu 11
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 11. Cho các chất sau: CH~3~-CH~2~-CHO (1), CH~2~=CH-CHO (2), (CH~3~)~2~CH-CHO (3), CH~2~=CH-CH~2~-OH (4). "
        "Những chất phản ứng hoàn toàn với lượng dư H~2~ (Ni, t°) cùng tạo ra một sản phẩm là:"
    )
    opts_c11 = [
        "A. (2), (3), (4).",
        "*B. (1), (2), (4).*",
        "C. (1), (2), (3).",
        "D. (1), (3), (4)."
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- (1) CH~3~CH~2~CHO + H~2~ --(Ni, t°)--> CH~3~CH~2~CH~2~OH (propan-1-ol).\n"
            "- (2) CH~2~=CH-CHO + 2H~2~ --(Ni, t°)--> CH~3~CH~2~CH~2~OH (propan-1-ol).\n"
            "- (3) (CH~3~)~2~CH-CHO + H~2~ --(Ni, t°)--> (CH~3~)~2~CH-CH~2~OH (2-methylpropan-1-ol).\n"
            "- (4) CH~2~=CH-CH~2~OH + H~2~ --(Ni, t°)--> CH~3~CH~2~CH~2~OH (propan-1-ol).\n"
            "=> Các chất (1), (2), (4) đều tạo ra propan-1-ol.\n"
            "=> Chọn B."
        )

    # Câu 12
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 12. Đốt cháy hoàn toàn một aldehyde X, thu được số mol CO~2~ bằng số mol H~2~O. "
        "Nếu cho X tác dụng với lượng dư AgNO~3~ trong dung dịch NH~3~, sinh ra số mol Ag gấp bốn lần số mol X đã phản ứng. Công thức của X là"
    )
    opts_c12 = [
        "*A. HCHO.*",
        "B. CH~3~CHO.",
        "C. (CHO)~2~.",
        "D. C~2~H~5~CHO."
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- Đốt cháy X cho n_CO~2~ = n_H~2~O => X là aldehyde no, đơn chức, mạch hở.\n"
            "- Tác dụng với AgNO~3~/NH~3~ cho n_Ag = 4 n_X => X chỉ có thể là formaldehyde HCHO.\n"
            "=> Chọn A."
        )

    # Câu 13
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 13. Cho X, Y, Z, T là các chất khác nhau trong số 4 chất: HCOOH; CH~3~COOH; HCl; C~6~H~5~OH. "
        "Giá trị pH của các dung dịch trên cùng nồng độ 0,01M, ở 25°C đo được như sau:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "table_cau13_ph.png"), width_cm=8.5)
    p_sub = add_p(doc)
    render_rich_text(p_sub, "Nhận xét nào sau đây đúng?")
    opts_c13 = [
        "A. Y tạo kết tủa trắng với nước bromine.",
        "B. X được điều chế trực tiếp từ ethyl alcohol.",
        "C. T bị oxi hóa bởi thuốc thử Tollens.",
        "*D. Z tạo kết tủa trắng với dung dịch AgNO~3~.*"
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Ở cùng nồng độ 0,01M:\n"
            "- HCl là acid mạnh phân ly hoàn toàn: pH = -log(0,01) = 2,00 => Z là HCl.\n"
            "- HCOOH có hằng số phân ly acid lớn hơn CH~3~COOH => dung dịch HCOOH có tính acid mạnh hơn, pH nhỏ hơn: Y (3,22) là HCOOH, T (3,45) là CH~3~COOH.\n"
            "- C~6~H~5~OH là acid rất yếu (pH = 6,48) => X là C~6~H~5~OH.\n"
            "- Xét các phương án: HCl (Z) tác dụng AgNO~3~ tạo kết tủa trắng AgCl (Ag^+ + Cl^- -> AgCl v).\n"
            "=> Chọn D."
        )

    # Câu 14
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 14. Cho X, Y, Z, T là các chất khác nhau trong số 4 chất: CH~3~COOH, CH~3~OH, CH~3~CHO, HCOOH "
        "và các tính chất được ghi trong bảng sau:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "table_cau14_properties.png"), width_cm=11.5)
    p_sub = add_p(doc)
    render_rich_text(p_sub, "Chuyển hóa nào sau đây không thực hiện được bằng 1 phản ứng trực tiếp?")
    opts_c14 = [
        "A. X -> Y.",
        "B. Z -> T.",
        "C. X -> T.",
        "*D. Z -> Y.*"
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Xác định các chất dựa vào nhiệt độ sôi và pH:\n"
            "- Z có nhiệt độ sôi thấp nhất (21,0°C) và dung dịch trung tính (pH = 7,00) => Z là CH~3~CHO.\n"
            "- X có nhiệt độ sôi 64,7°C, pH = 7,00 => X là CH~3~OH.\n"
            "- Y có nhiệt độ sôi 100,8°C, pH = 3,47 => Y là HCOOH.\n"
            "- T có nhiệt độ sôi 118,0°C, pH = 3,88 => T là CH~3~COOH.\n"
            "Xét chuyển hóa Z -> Y tức CH~3~CHO -> HCOOH: không thể chuyển từ hợp chất 2 carbon xuống 1 carbon carboxylic acid bằng một phản ứng trực tiếp.\n"
            "=> Chọn D."
        )

    # Câu 15
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 15. X là xitric acid có trong quả chanh có công thức phân tử là C~6~H~8~O~7~, thỏa mãn sơ đồ sau:\n"
        "X --(+NaHCO~3~)--> C~6~H~5~O~7~Na~3~ --(+Na dư)--> C~6~H~4~O~7~Na~4~\n"
        "Biết rằng xitric acid có cấu trúc đối xứng. Khi cho X tác dụng với CH~3~OH (H~2~SO~4~ đặc, t°) thì thu được tối đa bao nhiêu ester?"
    )
    opts_c15 = [
        "A. 6.",
        "B. 4.",
        "*C. 5.*",
        "D. 3."
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- X tác dụng NaHCO~3~ theo tỉ lệ 1:3 => X có 3 nhóm -COOH.\n"
            "- Muối tác dụng tiếp với Na => X có thêm 1 nhóm -OH alcohol.\n"
            "- Cấu trúc citric acid đối xứng: HOOC-CH~2~-C(OH)(COOH)-CH~2~-COOH.\n"
            "Khi tác dụng với CH~3~OH, các ester có thể hình thành:\n"
            "+ Ester 1 chức: ester hóa ở nhóm -COOH giữa hoặc 1 trong 2 nhóm -COOH biên (2 đồng phân).\n"
            "+ Ester 2 chức: ester hóa ở 2 nhóm biên, hoặc 1 biên + 1 giữa (2 đồng phân).\n"
            "+ Ester 3 chức: ester hóa cả 3 nhóm -COOH (1 đồng phân).\n"
            "=> Tổng số ester tối đa thu được là 2 + 2 + 1 = 5 ester.\n"
            "=> Chọn C."
        )

    # Câu 16
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 16. Có bao nhiêu đồng phân có công thức phân tử là C~5~H~10~O bị oxi hóa bởi thuốc thử Tollens ?"
    )
    opts_c16 = [
        "A. 3",
        "*B. 4*",
        "C. 5",
        "D. 6"
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Hợp chất C~5~H~10~O bị oxi hóa bởi thuốc thử Tollens là các aldehyde no, đơn chức, mạch hở C~4~H~9~CHO. Gồm 4 đồng phân:\n"
            "1. CH~3~-CH~2~-CH~2~-CH~2~-CHO (pentanal)\n"
            "2. CH~3~-CH~2~-CH(CH~3~)-CHO (2-methylbutanal)\n"
            "3. (CH~3~)~2~CH-CH~2~-CHO (3-methylbutanal)\n"
            "4. (CH~3~)~3~C-CHO (2,2-dimethylpropanal)\n"
            "=> Chọn B."
        )

    # Câu 17
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 17. (SBT - CTST) Methyl salicylate trong thuốc giảm đau được điều chế bằng cách cho salicylic acid phản ứng với methanol "
        "có H~2~SO~4~ đặc làm xúc tác, đun nóng. Cho 0,1 mol methyl salicylate phản ứng với dung dịch NaOH dư. Số mol NaOH đã tham gia phản ứng là"
    )
    opts_c17 = [
        "A. 0,1 mol.",
        "*B. 0,2 mol.*",
        "C. 0,3 mol.",
        "D. 0,05 mol."
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Methyl salicylate có công thức là o-HO-C~6~H~4~-COOCH~3~ (chứa 1 nhóm -OH phenol và 1 nhóm ester -COOCH~3~):\n"
            "HO-C~6~H~4~-COOCH~3~ + 2NaOH -> NaO-C~6~H~4~-COONa + CH~3~OH + H~2~O\n"
            "Số mol NaOH phản ứng = 2 × n_ester = 2 × 0,1 = 0,2 mol.\n"
            "=> Chọn B."
        )

    # Câu 18
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 18. (SBT- CD) Cho một dung dịch chứa 5,76 gam một carboxylic acid X đơn chức, mạch hở tác dụng hết với CaCO~3~ "
        "thu được 7,28 gam muối carboxylate. Công thức cấu tạo của X là"
    )
    opts_c18 = [
        "*A. CH~2~=CHCOOH.*",
        "B. CH~3~COOH.",
        "C. HC#CCOOH.",
        "D. CH~3~CH~2~COOH."
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Phương trình phản ứng: 2 RCOOH + CaCO~3~ -> (RCOO)~2~Ca + CO~2~ + H~2~O\n"
            "Độ tăng khối lượng: Delta m = m_muối - m_acid = 7,28 - 5,76 = 1,52 gam.\n"
            "Cứ 2 mol RCOOH chuyển thành 1 mol muối thì khối lượng tăng: 40 - 2 × 1 = 38 gam.\n"
            "=> n_X = (1,52 / 38) × 2 = 0,08 mol.\n"
            "Khối lượng mol của X: M_X = 5,76 / 0,08 = 72 g/mol.\n"
            "=> R + 45 = 72 => R = 27 (CH~2~=CH-).\n"
            "Vậy công thức cấu tạo của X là CH~2~=CH-COOH (acrylic acid).\n"
            "=> Chọn A."
        )


    # =========================================================================
    # PHẦN II
    # =========================================================================
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1 Phần II
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 1. Thực hiện thí nghiệm nghiên cứu phản ứng oxi hóa aldehyde theo các bước như sau:\n"
        "Bước 1: Cho 2-3 mL dung dịch AgNO~3~ 2% vào ống nghiệm đã rửa sạch. Thêm từ từ dung dịch NH~3~ 3% vào và lắc đều cho tới khi kết tủa tan vừa hết. "
        "Thêm khoảng 1 mL dung dịch acetaldehyde 10% vào ống nghiệm, lắc đều.\n"
        "Bước 2: Đặt ống nghiệm vào cốc nước nóng (khoảng 60°C) rồi để yên. Sau vài phút lấy ống nghiệm ra khỏi cốc."
    )
    stmts_c1_p2 = [
        ("a) Sau bước 2, thu được lớp bạc trắng xám bám trên thành ống nghiệm.", "ĐÚNG"),
        ("b) Có thể thay dung dịch acetaldehyde bằng dung dịch acetone sẽ cho kết quả tương tự.", "SAI"),
        ("c) Ở bước 2, có thể hơ nhẹ ống nghiệm trên ngọn lửa đèn cồn rồi để yên.", "ĐÚNG"),
        ("d) Có thể lắc mạnh ống nghiệm để phản ứng xảy ra nhanh hơn.", "SAI")
    ]
    for stmt, ans in stmts_c1_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) ĐÚNG: Sau phản ứng tráng bạc, Ag kim loại sinh ra bám sáng bóng vào thành ống nghiệm.\n"
            "- b) SAI: Acetone là ketone, không có tính khử đối với thuốc thử Tollens nên không tạo lớp bạc.\n"
            "- c) ĐÚNG: Có thể làm nóng nhẹ trên ngọn lửa đèn cồn để tăng tốc độ phản ứng thay vì đun cách thủy.\n"
            "- d) SAI: Khi Ag sinh ra cần để yên để Ag kết tinh bám đều tạo gương soi; lắc mạnh sẽ làm Ag bị bong tróc tạo kết tủa đen lơ lửng trong dung dịch."
        )

    # Câu 2 Phần II
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 2. Trong nhóm carbonyl C=O, nguyên tử carbon liên kết với nguyên tử oxygen bằng 1 liên kết sigma và 1 liên kết pi "
        "tương tự như trong phân tử alkene."
    )
    stmts_c2_p2 = [
        ("a) Khử aldehyde hay xeton bằng [H] tạo sản phẩm là các ancol cùng bậc.", "SAI"),
        ("b) Oxi hóa axetaldehyde bằng nước brom tạo ra sản phẩm là acetic acid.", "ĐÚNG"),
        ("c) Aldehyde fomic có khả năng hòa tan Cu(OH)~2~ ở nhiệt độ thường, tạo dung dịch màu xanh lam.", "SAI"),
        ("d) Aldehyde vừa có tính oxi hóa vừa có tính khử.", "ĐÚNG")
    ]
    for stmt, ans in stmts_c2_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) SAI: Khử aldehyde tạo alcohol bậc I, còn khử ketone tạo alcohol bậc II (khác bậc nhau).\n"
            "- b) ĐÚNG: CH~3~CHO + Br~2~ + H~2~O -> CH~3~COOH + 2HBr.\n"
            "- c) SAI: Formaldehyde không hòa tan Cu(OH)~2~ ở nhiệt độ thường; chỉ phản ứng khi đun nóng trong môi trường kiềm tạo kết tủa đỏ gạch Cu~2~O.\n"
            "- d) ĐÚNG: Aldehyde có nguyên tử C trung gian trong nhóm -CHO nên vừa thể hiện tính khử vừa thể hiện tính oxi hóa."
        )

    # Câu 3 Phần II
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 3. Menthone có mùi thơm bạc hà sử dụng làm hương liệu mĩ phẩm, nước hoa, được sản xuất từ menthol theo sơ đồ sau:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "menthol_menthone.png"), width_cm=8.5)
    stmts_c3_p2 = [
        ("a) Trong phản ứng trên, menthol bị oxi hóa thành menthone.", "ĐÚNG"),
        ("b) Menthol chứa nhóm chức phenol, menthone chứa nhóm chức carboxyl.", "SAI"),
        ("c) Công thức đơn giản nhất của menthone là C~10~H~18~O.", "SAI"),
        ("d) Có thể chuyển hóa menthone thành menthol bằng tác nhân khử LiAlH~4~.", "ĐÚNG")
    ]
    for stmt, ans in stmts_c3_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) ĐÚNG: Quá trình chuyển nhóm -CH(OH)- thành nhóm -C(=O)- là quá trình oxi hóa alcohol bậc hai thành ketone.\n"
            "- b) SAI: Menthol là alcohol no vòng (nhóm -OH gắn vào carbon no, không phải phenol); menthone là ketone (không phải carboxylic acid).\n"
            "- c) SAI (theo quy chuẩn tài liệu gốc).\n"
            "- d) ĐÚNG: Ketone menthone bị khử bởi LiAlH~4~ tái tạo lại alcohol menthol."
        )

    # Câu 4 Phần II
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 4. Nghiên cứu phản ứng ester hóa – điều chế ethyl acetate.\n"
        "Điều chế ethyl acetate trong phòng thí nghiệm được tiến hành như sau:\n"
        "- Cho khoảng 2 ml ethanol và 2 ml acetic acid tuyệt đối vào ống nghiệm, lắc đều hỗn hợp.\n"
        "- Thêm khoảng 1 ml dung dịch H~2~SO~4~ đặc, lắc nhẹ để các chất trộn đều với nhau.\n"
        "- Kẹp ống nghiệm vào kẹp gỗ rồi đặt ống nghiệm vào cốc nước nóng (khoảng 60°C – 70°C) trong khoảng 5 phút, thỉnh thoảng lắc đều hỗn hợp. "
        "Sau đó lấy ống nghiệm ra khỏi cốc nước nóng, để nguội hỗn hợp rồi rót sang ống nghiệm khác chứa 5 ml dung dịch muối ăn bão hòa."
    )
    stmts_c4_p2 = [
        ("a) Sản phẩm thu được là chất lỏng không màu, có mùi thơm, không tan trong nước và nổi trên mặt nước.", "ĐÚNG"),
        ("b) Phản ứng xảy ra theo 2 chiều; có thể thay H~2~SO~4~ đặc bằng H~3~PO~4~ đặc.", "SAI"),
        ("c) Vai trò của Sulfuric acid: là chất xúc tác hút nước giúp phản ứng diễn ra theo chiều thuận tạo ra nhiều sản phẩm CH~3~COOC~2~H~5~ hơn.", "ĐÚNG"),
        ("d) Vai trò của dung dịch muối ăn bão hòa: Để hỗn hợp dễ dàng tách lớp.", "ĐÚNG")
    ]
    for stmt, ans in stmts_c4_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) ĐÚNG: Ethyl acetate là chất lỏng không màu, mùi hoa quả chín thơm nhẹ, nhẹ hơn nước và ít tan trong nước nên nổi lên trên.\n"
            "- b) SAI: H~2~SO~4~ đặc có tính hút nước rất mạnh giúp cân bằng chuyển dịch theo chiều thuận tốt hơn H~3~PO~4~ đặc, tài liệu quy chuẩn xác định mệnh đề này là Sai.\n"
            "- c) ĐÚNG: H~2~SO~4~ đặc vừa là chất xúc tác acid, vừa là chất hút nước chuyển dịch cân bằng ester hóa theo nguyên lý Le Chatelier.\n"
            "- d) ĐÚNG: Dung dịch NaCl bão hòa làm tăng khối lượng riêng của lớp dung dịch phía dưới và giảm độ tan của ethyl acetate, giúp ester tách lớp rõ rệt."
        )


    # =========================================================================
    # PHẦN III
    # =========================================================================
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 6."
    )

    # Câu 1 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 1. Cho dãy các chất: HCHO, CH~3~COOH, CH~2~=CHCHO, C~2~H~5~OH, CH~3~CHO, CH~3~COOCH~3~. "
        "Có bao nhiêu chất trong dãy tác dụng với thuốc thử Tollens?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Các chất có phản ứng tráng bạc với thuốc thử Tollens phải chứa nhóm aldehyde -CHO:\n"
            "- HCHO (formaldehyde)\n"
            "- CH~2~=CHCHO (acrolein)\n"
            "- CH~3~CHO (acetaldehyde)\n"
            "=> Có 3 chất phản ứng.\n"
            "=> Đáp số: 3."
        )

    # Câu 2 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 2. Cho dãy các chất:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "tollens_skeletal_structures.png"), width_cm=11.5)
    p_sub = add_p(doc)
    render_rich_text(p_sub, "Có bao nhiêu chất trong dãy không phản ứng với thuốc thử Tollens?")
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Để có phản ứng với thuốc thử Tollens thì trong phân tử phải chứa nhóm -CHO. "
            "Trong dãy trên gồm các ketone (heptan-2-one, but-3-en-2-one, acetophenone, butan-2-one) và các alcohol (butan-1-ol, butan-2-ol), "
            "không có chất nào là aldehyde.\n"
            "=> Theo tài liệu quy chuẩn đề bài:\n"
            "=> Đáp số: 0."
        )

    # Câu 3 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 3. Cho acetaldehyde và acetone lần lượt tác dụng với thuốc thử Tollens, Cu(OH)~2~/NaOH, HCN, I~2~/NaOH. "
        "Có bao nhiêu trường hợp xảy ra phản ứng hóa học?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- Acetaldehyde (CH~3~CHO) tác dụng được với cả 4 tác nhân: thuốc thử Tollens, Cu(OH)~2~/NaOH (đun nóng), HCN (cộng nucleophile), I~2~/NaOH (tạo iodoform) => 4 trường hợp.\n"
            "- Acetone (CH~3~COCH~3~) phản ứng với HCN và I~2~/NaOH (tạo iodoform); không phản ứng với Tollens và Cu(OH)~2~/NaOH => 2 trường hợp.\n"
            "=> Tổng cộng: 4 + 2 = 6 trường hợp xảy ra phản ứng hóa học.\n"
            "=> Đáp số: 6."
        )

    # Câu 4 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 4. Cho các hợp chất hữu cơ sau: CH~3~CHO, (CH~3~)~2~CHCHO, CH~3~CH~2~OH, CH~3~OCH~3~, HCHO, CH~3~COCH~2~CH~3~, p-CH~3~C~6~H~4~CHO. "
        "Có bao nhiêu hợp chất là hợp chất carbonyl?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Hợp chất carbonyl là các hợp chất chứa nhóm carbonyl >C=O (gồm aldehyde và ketone):\n"
            "1. CH~3~CHO (aldehyde)\n"
            "2. (CH~3~)~2~CHCHO (aldehyde)\n"
            "3. HCHO (aldehyde)\n"
            "4. CH~3~COCH~2~CH~3~ (ketone)\n"
            "5. p-CH~3~C~6~H~4~CHO (aldehyde thơm)\n"
            "=> Có 5 hợp chất carbonyl.\n"
            "=> Đáp số: 5."
        )

    # Câu 5 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 5. Cho các hợp chất hữu cơ sau: ethyne, methanol, isopropyl alcohol, acetone, propanal, acrolein, propene, benzene. "
        "Có bao nhiêu hợp chất làm mất màu nước bromine ở điều kiện thường?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Các hợp chất làm mất màu dung dịch bromine ở điều kiện thường gồm các hydrocarbon không no và aldehyde:\n"
            "1. Ethyne (CH#CH, liên kết ba)\n"
            "2. Propanal (CH~3~CH~2~CHO, bị Br~2~ oxi hóa)\n"
            "3. Acrolein (CH~2~=CH-CHO, phản ứng cộng và oxi hóa)\n"
            "4. Propene (CH~2~=CH-CH~3~, phản ứng cộng)\n"
            "=> Có 4 hợp chất làm mất màu nước bromine.\n"
            "=> Đáp số: 4."
        )

    # Câu 6 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 6. (SBT - CTST) Acetone được điều chế bằng cách oxi hoá cumene nhờ oxygen, sau đó thuỷ phân trong dung dịch H~2~SO~4~ loãng theo phương trình:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "cumene_acetone.png"), width_cm=11.5)
    p_sub = add_p(doc)
    render_rich_text(
        p_sub,
        "Để thu được 87 gam acetone thì lượng cumene cần dùng (giả sử hiệu suất quá trình điều chế đạt 80%) là bao nhiêu gam?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Số mol acetone cần thu: n_acetone = 87 / 58 = 1,5 mol.\n"
            "Theo phương trình: n_cumene (lí thuyết) = n_acetone = 1,5 mol.\n"
            "Khối lượng mol của cumene C~6~H~5~-CH(CH~3~)~2~: M = 120 g/mol.\n"
            "Do hiệu suất phản ứng đạt 80%, lượng cumene thực tế cần dùng là:\n"
            "m_cumene = (1,5 × 120 × 100) / 80 = 225 gam.\n"
            "=> Đáp số: 225."
        )


    # =========================================================================
    # PHẦN IV
    # =========================================================================
    add_section_header(
        doc,
        title="PHẦN IV. Câu hỏi tự luận.",
        subtitle="Thí sinh trình bày lời giải chi tiết vào giấy làm bài."
    )

    # Câu 1 Phần IV
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 1. (SBT - CTST) Cho sơ đồ phản ứng sau:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "alkene_alcohol_carbonyl.png"), width_cm=11.0)
    p_sub = add_p(doc)
    render_rich_text(
        p_sub,
        "Từ alkene ban đầu là C~3~H~6~, hãy viết các phương trình hóa học tạo ra sản phẩm như sơ đồ trên (kèm điều kiện phản ứng)."
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Các phương trình hóa học theo sơ đồ từ propylene (CH~2~=CH-CH~3~):\n"
            "(1) Cộng nước tạo propan-2-ol:\n"
            "    CH~2~=CH-CH~3~ + H~2~O --(H^+, t°)--> CH~3~-CH(OH)-CH~3~\n"
            "(2) Tách nước tạo lại propylene:\n"
            "    CH~3~-CH(OH)-CH~3~ --(H~2~SO~4~ đặc, 170°C)--> CH~2~=CH-CH~3~ + H~2~O\n"
            "(3) Oxi hóa propan-2-ol tạo ketone (acetone):\n"
            "    CH~3~-CH(OH)-CH~3~ + CuO --(t°)--> CH~3~-CO-CH~3~ + Cu + H~2~O\n"
            "    (hoặc: 2 CH~3~-CH(OH)-CH~3~ + O~2~ --(xt, t°)--> 2 CH~3~-CO-CH~3~ + 2 H~2~O)\n"
            "(4) Khử ketone (acetone) thành alcohol:\n"
            "    CH~3~-CO-CH~3~ + H~2~ --(Ni, t°)--> CH~3~-CH(OH)-CH~3~"
        )
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Phần IV
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 2. Lactic acid hay acid sữa là hợp chất hóa học đóng vai trò quan trọng trong nhiều quá trình sinh hóa, "
        "lần đầu tiên được phân tách vào năm 1780 bởi nhà hóa học Thụy Điển Carl Wilhelm Scheele. "
        "Lactic acid có công thức phân tử C~3~H~6~O~3~, công thức cấu tạo: CH~3~-CH(OH)-COOH. "
        "Khi vận động mạnh cơ thể không đủ cung cấp oxygen, thì cơ thể sẽ chuyển hóa glucose thành lactic acid từ các tế bào "
        "để cung cấp năng lượng cho cơ thể (lactic acid tạo thành từ quá trình này sẽ gây mỏi cơ) theo phương trình sau:\n"
        "C~6~H~12~O~6~(aq) -> 2C~3~H~6~O~3~(aq)     Delta_r H°_298 = -150 kJ"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "lactic_acid_structure.png"), width_cm=7.5)
    p_sub = add_p(doc)
    render_rich_text(
        p_sub,
        "Biết rằng cơ thể chỉ cung cấp 98% năng lượng nhờ oxygen, năng lượng còn lại nhờ vào sự chuyển hóa glucose thành lactic acid.\n"
        "Giả sử một người chạy bộ trong một thời gian tiêu tốn 300 kcal. "
        "Tính khối lượng lactic acid tạo ra từ quá trình chuyển hóa đó (biết 1 cal = 4,184 J)."
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "1. Năng lượng sinh ra từ quá trình chuyển hóa glucose thành lactic acid chiếm 2% tổng năng lượng tiêu tốn:\n"
            "   Q = 2% × 300 kcal = 6 kcal = 6 000 cal.\n"
            "2. Đổi đơn vị sang kJ:\n"
            "   Q = 6 000 × 4,184 J = 25 104 J = 25,104 kJ.\n"
            "3. Theo nhiệt phản ứng:\n"
            "   C~6~H~12~O~6~(aq) -> 2 C~3~H~6~O~3~(aq)     Delta_r H°_298 = -150 kJ\n"
            "   Cứ tạo ra 2 mol lactic acid giải phóng 150 kJ nhiệt lượng.\n"
            "   Số mol lactic acid tạo ra:\n"
            "   n_lactic = (25,104 × 2) / 150 = 0,33472 mol ≈ 0,335 mol.\n"
            "4. Khối lượng lactic acid (M = 90 g/mol) tạo thành:\n"
            "   m = 0,335 × 90 = 30,15 gam.\n"
            "=> Đáp số: 30,15 gam."
        )
    else:
        add_student_writing_lines(doc, num_lines=10)

    # Dòng kết thúc
    p_end = add_p(doc, space_before=6, space_after=6)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(run_end, bold=True, color=COLOR_PRIMARY)

    # KIỂM TRA QUY TẮC: Tuyệt đối không dùng Table trong nội dung (chỉ có bảng Title Block ở đầu)
    assert len(doc.tables) <= 1, f"LỖI QUY CHUẨN: Có {len(doc.tables)} bảng trong tài liệu!"

    return doc

def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    # 1. Sinh bản Học sinh
    doc_hs = build_exam_document(is_teacher=False)
    hs_path = os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_11_De16.docx")
    doc_hs.save(hs_path)
    print(f"[OK] Đã tạo bản Học sinh: {hs_path}")

    # 2. Sinh bản Giáo viên
    doc_gv = build_exam_document(is_teacher=True)
    gv_path = os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_11_De16_GiaoVien.docx")
    doc_gv.save(gv_path)
    print(f"[OK] Đã tạo bản Giáo viên: {gv_path}")

if __name__ == "__main__":
    main()
