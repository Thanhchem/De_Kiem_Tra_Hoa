# -*- coding: utf-8 -*-
"""
Script tạo 2 file Word chuẩn Form Đề Kiểm Tra Hóa học (Bản Học Sinh & Bản Giáo Viên)
Đề: TRƯỜNG THPT HOÁ CHÂU - MÃ ĐỀ 132 - CUỐI HỌC KÌ I - LỚP 10 (2025-2026)
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

sys.path.append(os.path.abspath('scripts'))
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
    COLOR_PRIMARY,
    COLOR_SOLUTION
)

def create_exam_file(is_teacher=False, output_path=""):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="SỞ GD & ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT HOÁ CHÂU",
        exam_title="ĐỀ KIỂM TRA CUỐI HỌC KÌ I - NĂM HỌC 2025 - 2026",
        subject_line="MÔN: HOÁ HỌC – LỚP 10",
        num_pages="02",
        duration="45 phút",
        exam_code="132",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Si = 28; P = 31; S = 32; Cl = 35,5; Mg = 24; Ca = 40; Fe = 56; Ag = 108",
        is_teacher=is_teacher
    )

    # -------------------------------------------------------------
    # BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO BẢN GIÁO VIÊN THEO MẪU ẢNH CHUẨN)
    # -------------------------------------------------------------
    if is_teacher:
        from build_bang_dap_an_rut_gon import add_styled_answer_block, EXAMS_DATA
        add_styled_answer_block(doc, EXAMS_DATA[1])

    # -------------------------------------------------------------
    # PHẦN I
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="A. TRẮC NGHIỆM KHÁCH QUAN",
        subtitle="PHẦN 1: Câu trắc nghiệm nhiều phương án lựa chọn (3 điểm). Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    questions_p1 = [
        {
            "num": 1,
            "prompt": "Câu 1: Bản chất của liên kết ion là lực hút tĩnh điện giữa",
            "opts": [
                "A. 2 ion.",
                "*B. các ion mang điện tích trái dấu.*",
                "C. các hạt mang điện trái dấu.",
                "D. hạt nhân và các electron hóa trị."
            ],
            "sol": "Liên kết ion là liên kết hóa học được hình thành bởi lực hút tĩnh điện giữa các ion mang điện tích trái dấu (cation và anion).\n➔ Chọn đáp án B."
        },
        {
            "num": 2,
            "prompt": "Câu 2: Ion X^2-^ có cấu hình e của phân lớp ngoài cùng là 2p^6^. Cấu hình electron của nguyên tử X là:",
            "opts": [
                "A. [Ne]2p^6^3s^2^.",
                "B. [Ne]2p^4^.",
                "*C. 1s^2^2s^2^2p^4^.*",
                "D. 1s^2^2s^2^2p^6^."
            ],
            "sol": "Ion X^2-^ được hình thành khi nguyên tử X nhận 2 electron: X + 2e ➔ X^2-^.\nTheo bài ra, cấu hình của X^2-^ là 1s^2^2s^2^2p^6^ ➔ Cấu hình electron của nguyên tử X là: 1s^2^2s^2^2p^4^ (Oxygen - O, Z = 8).\n➔ Chọn đáp án C."
        },
        {
            "num": 3,
            "prompt": "Câu 3: Điều nào dưới đây không đúng khi nói về hợp chất sodium oxide (Na~2~O)?",
            "opts": [
                "A. Trong phân tử Na~2~O, các ion sodium Na^+^ và ion oxide O^2-^ đều đạt cấu hình electron bền vững của khí hiếm neon.",
                "B. Phân tử Na~2~O tạo bởi lực hút tĩnh điện giữa hai ion Na^+^ và một ion O^2-^.",
                "C. Là chất rắn trong điều kiện thường.",
                "*D. Không tan trong nước, chỉ tan trong dung môi không phân cực như benzene, carbon tetrachloride,...*"
            ],
            "sol": "Na~2~O là oxide base tan rất mạnh trong nước tạo dung dịch kiềm: Na~2~O + H~2~O ➔ 2NaOH.\nDo đó phát biểu D 'Không tan trong nước...' là SAI.\n➔ Chọn đáp án D."
        },
        {
            "num": 4,
            "prompt": "Câu 4: Trong bảng hệ thống tuần hoàn các nguyên tố, số chu kì nhỏ và chu kì lớn là",
            "opts": [
                "A. 4 và 3",
                "*B. 3 và 4*",
                "C. 3 và 3",
                "D. 4 và 4"
            ],
            "sol": "Bảng tuần hoàn gồm 7 chu kì:\n- Chu kì nhỏ gồm 3 chu kì (chu kì 1, 2, 3).\n- Chu kì lớn gồm 4 chu kì (chu kì 4, 5, 6, 7).\n➔ Chọn đáp án B."
        },
        {
            "num": 5,
            "prompt": "Câu 5: Trong phân tử ammonia (NH~3~), số cặp electron chung giữa nguyên tử nitrogen và các nguyên tử hydrogen là",
            "opts": [
                "A. 1.",
                "B. 2.",
                "*C. 3.*",
                "D. 4."
            ],
            "sol": "Trong phân tử NH~3~, nguyên tử N liên kết với 3 nguyên tử H bằng 3 liên kết đơn N-H. Mỗi liên kết đơn là một cặp electron dùng chung ➔ Có 3 cặp electron dùng chung.\n➔ Chọn đáp án C."
        },
        {
            "num": 6,
            "prompt": "Câu 6: Hợp chất khí với hydrogen của nguyên tố R là RH~4~. Oxide cao nhất của R chứa 53,3% oxygen về khối lượng. Oxide này được sử dụng trong ngành xây dựng, như sản xuất bê tông. Nguyên tử khối của R là",
            "opts": [
                "A. 32",
                "B. 12",
                "C. 16",
                "*D. 28*"
            ],
            "sol": "Hợp chất khí với H là RH~4~ ➔ R thuộc nhóm IVA ➔ Oxide cao nhất là RO~2~.\n%m_O = [2 × 16 / (R + 32)] × 100% = 53,3% ➔ R + 32 = 60 ➔ R = 28 (Silicon - Si, SiO~2~ được dùng trong cát xây dựng, xi măng, bê tông).\n➔ Chọn đáp án D."
        },
        {
            "num": 7,
            "prompt": "Câu 7: Trong phân tử HCl, nguyên tử hydrogen và chlorine đã lần lượt đạt cấu hình electron bền của các khí hiếm nào dưới đây?",
            "opts": [
                "A. Helium và krypton.",
                "B. Helium và radon.",
                "*C. Helium và argon.*",
                "D. Helium và xenon."
            ],
            "sol": "Khi hình thành liên kết CHT trong HCl:\n- H (Z=1) góp chung 1e để có 2e ở lớp ngoài cùng, đạt cấu hình bền của khí hiếm Helium (He).\n- Cl (Z=17) góp chung 1e để có 8e ở lớp ngoài cùng, đạt cấu hình octet bền của khí hiếm Argon (Ar).\n➔ Chọn đáp án C."
        },
        {
            "num": 8,
            "prompt": "Câu 8: Sulfur (S) là nguyên tố thuộc nhóm VIA, chu kì 3 của bảng tuần hoàn. Trong các phát biểu sau:\n(1) Nguyên tử S có 3 lớp electron và có 10 electron p.\n(2) Nguyên tử S có 5 electron hóa trị và 6 electron s.\n(3) Công thức oxide cao nhất của S có dạng SO~3~ và là acidic oxide.\n(4) Nguyên tố S có tính phi kim mạnh hơn so với nguyên tố có số hiệu nguyên tử là 8.\n(5) Hydroxide cao nhất của S có dạng H~2~SO~4~ và có tính acid.\nSố phát biểu đúng là",
            "opts": [
                "A. 4.",
                "*B. 3.*",
                "C. 5.",
                "D. 2."
            ],
            "sol": "Cấu hình electron của S (Z=16): 1s^2^2s^2^2p^6^3s^2^3p^4^.\n- (1) ĐÚNG: S có 3 lớp electron; tổng electron p = 6 (2p) + 4 (3p) = 10 electron.\n- (2) SAI: S có 6 electron hóa trị (lớp ngoài cùng 3s^2^3p^4^).\n- (3) ĐÚNG: Oxide cao nhất SO~3~ là acidic oxide.\n- (4) SAI: Nguyên tố Z=8 là Oxygen (O), cùng nhóm VIA với S. Trong một nhóm A, từ trên xuống dưới tính phi kim giảm dần nên tính phi kim của O mạnh hơn S (S yếu hơn O).\n- (5) ĐÚNG: Hydroxide cao nhất là H~2~SO~4~, là acid mạnh.\n➔ Có 3 phát biểu đúng: (1), (3), (5).\n➔ Chọn đáp án B."
        },
        {
            "num": 9,
            "prompt": "Câu 9: Hợp chất nào sau đây chứa cả liên kết cộng hóa trị và liên kết ion?",
            "opts": [
                "A. CH~2~O.",
                "B. CH~4~.",
                "C. Na~2~O.",
                "*D. KOH.*"
            ],
            "sol": "Trong hợp chất KOH:\n- Giữa cation K^+^ và anion OH^-^ là liên kết ion.\n- Trong anion đa nguyên tử OH^-^, giữa nguyên tử O và H là liên kết cộng hóa trị phân cực.\n➔ KOH chứa cả liên kết ion và liên kết cộng hóa trị.\n➔ Chọn đáp án D."
        },
        {
            "num": 10,
            "prompt": "Câu 10: Cho giá trị độ âm điện của các nguyên tố như sau: Ca (1,00); Cl (3,16); N (3,04); H (2,20); C (2,55); O (3,44); Al (1,61). Dựa vào giá trị hiệu độ âm điện của hai nguyên tử trong phân tử, hãy cho biết trong các hợp chất sau: CaCl~2~, NH~3~, AlCl~3~, CH~4~, CO~2~. Số hợp chất có liên kết ion và số hợp chất có liên kết cộng hóa trị là",
            "opts": [
                "A. 2 và 3",
                "B. 3 và 2",
                "C. 4 và 1",
                "*D. 1 và 4*"
            ],
            "sol": "Xét hiệu độ âm điện Δχ của liên kết trong từng hợp chất:\n- CaCl~2~: Δχ(Cl-Ca) = 3,16 - 1,00 = 2,16 ≥ 1,7 ➔ Liên kết ion (1 hợp chất).\n- NH~3~: Δχ(N-H) = 3,04 - 2,20 = 0,84 (0,4 ≤ Δχ < 1,7) ➔ Liên kết cộng hóa trị.\n- AlCl~3~: Δχ(Cl-Al) = 3,16 - 1,61 = 1,55 (0,4 ≤ Δχ < 1,7) ➔ Liên kết cộng hóa trị.\n- CH~4~: Δχ(C-H) = 2,55 - 2,20 = 0,35 (< 0,4) ➔ Liên kết cộng hóa trị.\n- CO~2~: Δχ(O-C) = 3,44 - 2,55 = 0,89 (0,4 ≤ Δχ < 1,7) ➔ Liên kết cộng hóa trị.\n➔ Số hợp chất có liên kết ion là 1 (CaCl~2~) và số hợp chất có liên kết cộng hóa trị là 4.\n➔ Chọn đáp án D."
        },
        {
            "num": 11,
            "prompt": "Câu 11: Các nguyên tố trong bảng tuần hoàn được sắp xếp theo chiều tăng dần",
            "opts": [
                "*A. số hiệu nguyên tử.*",
                "B. khối lượng nguyên tử.",
                "C. bán kính nguyên tử.",
                "D. độ âm điện của nguyên tử."
            ],
            "sol": "Theo nguyên tắc sắp xếp của Bảng tuần hoàn: Các nguyên tố được sắp xếp theo chiều tăng dần của điện tích hạt nhân (chính là số hiệu nguyên tử Z).\n➔ Chọn đáp án A."
        },
        {
            "num": 12,
            "prompt": "Câu 12: Hạt mang điện trong hạt nhân nguyên tử là",
            "opts": [
                "*A. Proton.*",
                "B. Neutron.",
                "C. Electron.",
                "D. Neutron và electron."
            ],
            "sol": "Hạt nhân nguyên tử gồm proton (mang điện tích dương +1) và neutron (không mang điện). Do đó, hạt mang điện trong hạt nhân chỉ có proton (electron ở lớp vỏ).\n➔ Chọn đáp án A."
        }
    ]

    for q in questions_p1:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(1)
        render_question_prompt(p_q, q["prompt"])
        render_options_no_table(doc, q["opts"], is_teacher=is_teacher)
        if is_teacher:
            add_solution_no_table(doc, q["sol"])

    # -------------------------------------------------------------
    # PHẦN II
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai: 2 điểm",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1 phần II
    p_q1_p2 = doc.add_paragraph()
    p_q1_p2.paragraph_format.space_before = Pt(3)
    p_q1_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q1_p2, "Câu 1: Sulfur dạng kem bôi được sử dụng để điều trị mụn trứng cá. Trong tự nhiên, sulfur có 4 đồng vị bền là: ^32^_16_S (95,02%), ^33^_16_S (0,75%), ^34^_16_S (4,21%) và ^36^_16_S (0,02%).")

    stmts_c1 = [
        ("a) Hạt nhân nguyên tử sulfur có 16 electron.", "Sai"),
        ("b) Cấu hình electron của ion S^2-^ giống với cấu hình e nguyên tử Ar.", "Đúng"),
        ("c) Sulfur là một phi kim thuộc chu kì 3, nhóm VIA của bảng tuần hoàn. Nguyên tử khối trung bình của sulfur có giá trị là 32,0891.", "Sai"),
        ("d) Oxide cao nhất có dạng SO~3~ và hydroxide có dạng H~2~SO~4~ và chúng đều có tính acid. Cho 5,6 gam Fe tác dụng hết với dung dịch H~2~SO~4~ loãng dư, thu được 2,24 lít khí hydrogen ở điều kiện chuẩn.", "Sai")
    ]
    for stmt, ans in stmts_c1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        sol_c1_p2 = (
            "a) SAI: Hạt nhân nguyên tử chỉ chứa proton và neutron, không chứa electron (16 electron ở lớp vỏ nguyên tử).\n"
            "b) ĐÚNG: Nguyên tử S (Z=16) nhận thêm 2e tạo ion S^2-^ có 18e: 1s^2^2s^2^2p^6^3s^2^3p^6^ (giống cấu hình nguyên tử khí hiếm Ar).\n"
            "c) SAI: Tính nguyên tử khối trung bình theo số khối các đồng vị: Ā = (32 × 95,02 + 33 × 0,75 + 34 × 4,21 + 36 × 0,02) / 100 = 32,0925 (khác giá trị 32,0891 trong đề).\n"
            "d) SAI: n~Fe~ = 5,6 / 56 = 0,1 mol ➔ n~H2~ = 0,1 mol. Ở điều kiện chuẩn (đkc: 25°C, 1 bar), V~H2~ = 0,1 × 24,79 = 2,479 lít (thể tích 2,24 lít là tính theo đktc 0°C cũ)."
        )
        add_solution_no_table(doc, sol_c1_p2)

    # Câu 2 phần II
    p_q2_p2 = doc.add_paragraph()
    p_q2_p2.paragraph_format.space_before = Pt(4)
    p_q2_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q2_p2, "Câu 2: Hợp chất sodium chloride (NaCl) là thành phần chính muối ăn")

    stmts_c2 = [
        ("a) Nguyên tố Na, Cl đều thuộc chu kì 3 trong trong bảng tuần hoàn.", "Đúng"),
        ("b) Để đạt đến cấu hình bền vững của khí hiếm gần nó nhất, nguyên tử Na nhận 1 electron và nguyên tử Cl nhường 7electron.", "Sai"),
        ("c) Liên kết trong hợp chất NaCl là liên kết ion.", "Đúng"),
        ("d) NaOH được điều chế bằng cách điện phân dung dịch NaCl bão hoà. Trong phân tử NaOH chỉ hình thành liên kết cộng hoá trị phân cực.", "Sai")
    ]
    for stmt, ans in stmts_c2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        sol_c2_p2 = (
            "a) ĐÚNG: Na (Z=11): [Ne]3s^1^ và Cl (Z=17): [Ne]3s^2^3p^5^ đều có 3 lớp electron ➔ Cùng thuộc chu kì 3.\n"
            "b) SAI: Nguyên tử Na có 1e ngoài cùng nên nhường 1 electron (thành Na^+^); nguyên tử Cl có 7e ngoài cùng nên nhận 1 electron (thành Cl^-^).\n"
            "c) ĐÚNG: Lực hút tĩnh điện giữa cation Na^+^ và anion Cl^-^ là liên kết ion.\n"
            "d) SAI: Trong tinh thể NaOH, có liên kết ion giữa Na^+^ và OH^-^; trong ion OH^-^ có liên kết CHT phân cực giữa O và H (không phải 'chỉ có liên kết CHT')."
        )
        add_solution_no_table(doc, sol_c2_p2)

    # -------------------------------------------------------------
    # PHẦN III
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="PHẦN III. (8 câu: 2 điểm) Câu hỏi trắc nghiệm yêu cầu trả lời ngắn",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 8. Mỗi câu đúng được 0,25 điểm."
    )

    questions_p3 = [
        {
            "num": 1,
            "prompt": "Câu 1: Nguyên tử X có 26 proton trong hạt nhân. Cho các phát biểu sau về X:\n(1) X có 26 neutron trong hạt nhân.\n(2) X có 26 electron ở vỏ nguyên tử.\n(3) X có điện tích hạt nhân là +26.\n(4) Khối lượng nguyên tử X là 26 amu.\nCó bao nhiêu phát biểu đúng?",
            "ans": "2",
            "sol": "- (1) SAI: X có Z = 26 (nguyên tố Fe), số neutron thường là 30 (chưa đủ dữ kiện kết luận N = 26).\n- (2) ĐÚNG: Nguyên tử trung hòa về điện ➔ số electron = số proton = 26.\n- (3) ĐÚNG: Điện tích hạt nhân là +Ze = +26.\n- (4) SAI: Khối lượng nguyên tử xấp xỉ bằng số khối A = Z + N = 26 + N > 26 amu.\n➔ Có 2 phát biểu đúng: (2) và (3)."
        },
        {
            "num": 2,
            "prompt": "Câu 2: Oxi có 3 đồng vị ^16^_8_O, ^17^_8_O, ^18^_8_O. Lithium có hai đồng vị bền là: ^6^_3_Li, ^7^_3_Li. Có thể có bao nhiêu loại phân tử Li~2~O được tạo thành giữa lithium và oxygen?",
            "ans": "9",
            "sol": "Phân tử Li~2~O gồm 2 nguyên tử Li và 1 nguyên tử O.\n- Các tổ hợp cặp 2 nguyên tử Li: (^6^Li, ^6^Li), (^7^Li, ^7^Li), (^6^Li, ^7^Li) ➔ Có 3 kiểu cặp đồng vị Li.\n- Kết hợp với 3 đồng vị của Oxygen (^16^O, ^17^O, ^18^O) ➔ Số loại phân tử Li~2~O = 3 × 3 = 9 loại phân tử."
        },
        {
            "num": 3,
            "prompt": "Câu 3: Bạc có hai đồng vị bền trong tự nhiên: ^107^Ag có hàm lượng tương đối là 51,8%; ^109^Ag có hàm lượng tương đối là 48,2%. Tính nguyên tử khối trung bình của Ag. (Kết quả làm tròn đến hàng đơn vị).",
            "ans": "108",
            "sol": "Nguyên tử khối trung bình của Bạc (Ag):\nĀ~Ag~ = (107 × 51,8 + 109 × 48,2) / 100 = (5542,6 + 5253,8) / 100 = 107,964.\nLàm tròn đến hàng đơn vị: 108."
        },
        {
            "num": 4,
            "prompt": "Câu 4: Nguyên tử của nguyên tố Fe có 26e, nguyên tử Fe có bao nhiêu electron hoá trị.",
            "ans": "8",
            "sol": "Cấu hình electron của Fe (Z = 26): 1s^2^2s^2^2p^6^3s^2^3p^6^3d^6^4s^2^.\nVì Fe là nguyên tố d (nhóm VIIIB), phân lớp 3d chưa bão hòa ➔ Số electron hóa trị = số e ở phân lớp 3d + số e ở phân lớp ngoài cùng 4s = 6 + 2 = 8 electron."
        },
        {
            "num": 5,
            "prompt": "Câu 5: Xét bốn nguyên tố có cấu hình electron nguyên tử như sau :\nX: 1s^2^2s^2^2p^6^3s^1^; Y: 1s^2^2s^2^2p^6^3s^2^; Z: 1s^2^2s^2^2p^6^3s^2^3p^1^; Q: 1s^2^2s^2^2p^6^3s^2^3p^5^; P: 1s^2^.\nCó bao nhiêu nguyên tố kim loại?",
            "ans": "3",
            "sol": "Xét số electron ở lớp ngoài cùng:\n- X: có 1e lớp ngoài cùng (3s^1^) ➔ Kim loại (Na).\n- Y: có 2e lớp ngoài cùng (3s^2^) ➔ Kim loại (Mg).\n- Z: có 3e lớp ngoài cùng (3s^2^3p^1^) ➔ Kim loại (Al).\n- Q: có 7e lớp ngoài cùng (3s^2^3p^5^) ➔ Phi kim (Cl).\n- P: có cấu hình 1s^2^ bão hòa ➔ Khí hiếm (He).\n➔ Có 3 nguyên tố kim loại là X, Y, Z."
        },
        {
            "num": 6,
            "prompt": "Câu 6: Một nguyên tố tạo hợp chất khí với hydrogen có công thức RH~3~, được sử dụng để trung hoà các thành phần acid của dầu thô, bảo vệ thiết bị không bị ăn mòn trong ngành công nghiệp dầu khí. Nguyên tố này chiếm 25,93% về khối lượng trong oxide cao nhất. Nguyên tử khối của nguyên tố R là bao nhiêu?",
            "ans": "14",
            "sol": "Hợp chất khí với H là RH~3~ ➔ R thuộc nhóm VA ➔ Oxide cao nhất là R~2~O~5~.\nTheo bài ra: %m_R = [2R / (2R + 5 × 16)] × 100% = 25,93%\n➔ 2R / (2R + 80) = 0,2593 ➔ 2R = 0,5186 R + 20,744 ➔ 1,4814 R = 20,744 ➔ R ≈ 14.\nVậy R là Nitrogen (N, nguyên tử khối = 14, hợp chất khí là NH~3~)."
        },
        {
            "num": 7,
            "prompt": "Câu 7: Cho dãy các ion: K^+^, S^2-^, Al^3+^, SO~4~^2-^, NH~4~^+^, CO~3~^2-^, Na^+^, NO~3~^-^, Cl^-^, Cu^2+^, Fe^3+^. Hỏi có bao nhiêu cation trong dãy trên?",
            "ans": "6",
            "sol": "Cation là các ion mang điện tích dương:\nCác cation trong dãy gồm: K^+^, Al^3+^, NH~4~^+^, Na^+^, Cu^2+^, Fe^3+^.\n➔ Có tất cả 6 cation."
        },
        {
            "num": 8,
            "prompt": "Câu 8: Hai nguyên tố X và Y (Z~X~ < Z~Y~) đứng kế tiếp nhau trong cùng một chu kì, có tổng số điện tích hạt nhân bằng 33. Số hiệu nguyên tử của X là bao nhiêu?",
            "ans": "16",
            "sol": "Vì X và Y kế tiếp nhau trong cùng một chu kì nên: Z~Y~ - Z~X~ = 1.\nTổng số điện tích hạt nhân: Z~X~ + Z~Y~ = 33.\nGiải hệ phương trình: 2 Z~X~ = 33 - 1 = 32 ➔ Z~X~ = 16 (Sulfur - S).\n(Z~Y~ = 17 là Chlorine - Cl)."
        }
    ]

    for q in questions_p3:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(1)
        render_question_prompt(p_q, q["prompt"])
        if is_teacher:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.5)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(1)
            render_rich_text(p_ans, f"➔ *Đáp án:* *{q['ans']}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
            add_solution_no_table(doc, q["sol"])
        else:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.5)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(2)
            render_rich_text(p_ans, "*Trả lời:* .....................................................................", base_size=10.5)

    # -------------------------------------------------------------
    # PHẦN B: TỰ LUẬN
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="B. PHẦN IV- TỰ LUẬN: 3 điểm",
        subtitle="Thí sinh trình bày lời giải chi tiết cho các câu hỏi sau:"
    )

    # Câu 1 Tự luận
    p_tl1 = doc.add_paragraph()
    p_tl1.paragraph_format.space_before = Pt(3)
    p_tl1.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_tl1, "Câu 1 (1,5 điểm):")

    p_tl1_sub = doc.add_paragraph()
    p_tl1_sub.paragraph_format.left_indent = Cm(0.5)
    p_tl1_sub.paragraph_format.space_before = Pt(1)
    p_tl1_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_tl1_sub, "a. Hãy viết công thức electron, công thức Lewis, công thức cấu tạo, của các chất có công thức phân tử sau: NH~3~; H~2~S; Cl~2~; CO.\nb. Cho biết loại liên kết giữa các nguyên tử trong phân tử là loại liên kết gì?(ion, cộng hoá trị phân cực hay cộng hoá trị không phân cực, liên kết cho nhận). (Biết độ âm điện của H=2,20; N = 3,04; S = 2,58; O =3,44; Cl =3,16). Giải thích.")

    if is_teacher:
        sol_tl1 = (
            "a. Viết công thức electron, Lewis và cấu tạo:\n"
            "  1. NH~3~: N có 5e hóa trị, mỗi H có 1e hóa trị.\n"
            "     - Công thức electron: H : N̈(H) : H\n"
            "     - Công thức Lewis: H - N̈(H) - H (N liên kết đơn với 3 H và còn 1 cặp e tự do).\n"
            "     - Công thức cấu tạo: H - N(H) - H\n"
            "  2. H~2~S: S có 6e hóa trị, mỗi H có 1e hóa trị.\n"
            "     - Công thức electron: H : S̈̈ : H\n"
            "     - Công thức Lewis: H - S̈̈ - H (S liên kết đơn với 2 H và còn 2 cặp e tự do).\n"
            "     - Công thức cấu tạo: H - S - H\n"
            "  3. Cl~2~: mỗi Cl có 7e hóa trị, góp chung 1 cặp e.\n"
            "     - Công thức electron: :C̈l̈ : C̈l̈:\n"
            "     - Công thức Lewis: :C̈l̈ - C̈l̈: (mỗi Cl còn 3 cặp e tự do).\n"
            "     - Công thức cấu tạo: Cl - Cl\n"
            "  4. CO: C có 4e hóa trị, O có 6e hóa trị. Hai nguyên tử góp chung 2 cặp e, nguyên tử O cho cặp e tự do vào orbital trống của C tạo liên kết phối trí.\n"
            "     - Công thức electron: :C ::: O:\n"
            "     - Công thức Lewis: :C ≡ O: (hoặc :C ⇆ O:)\n"
            "     - Công thức cấu tạo: C ≡ O (hoặc C ⇆ O)\n"
            "b. Xác định loại liên kết và giải thích (dựa vào hiệu độ âm điện Δχ):\n"
            "  - Trong NH~3~: Δχ(N-H) = 3,04 - 2,20 = 0,84 (0,4 ≤ Δχ < 1,7) ➔ Liên kết cộng hóa trị phân cực.\n"
            "  - Trong H~2~S: Δχ(S-H) = 2,58 - 2,20 = 0,38 (< 0,4) ➔ Liên kết cộng hóa trị không phân cực (hoặc phân cực rất yếu).\n"
            "  - Trong Cl~2~: Δχ(Cl-Cl) = 3,16 - 3,16 = 0 (< 0,4) ➔ Liên kết cộng hóa trị không phân cực.\n"
            "  - Trong CO: Liên kết C-O có Δχ = 3,44 - 2,55 = 0,89 (0,4 ≤ Δχ < 1,7) là liên kết cộng hóa trị phân cực, đồng thời có 1 liên kết cho - nhận (phối trí) do cặp electron tự do của O đóng góp vào orbital trống của C."
        )
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=10)

    # Câu 2 Tự luận
    p_tl2 = doc.add_paragraph()
    p_tl2.paragraph_format.space_before = Pt(4)
    p_tl2.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_tl2, "Câu 2 (1,5 điểm): Kim loại M thuộc nhóm A trong bảng tuần hoàn có hoá trị n chưa biết, là một thành phần dinh dưỡng quan trọng . Sự thiếu hụt rất nhỏ của nó ảnh hưởng tới sự hình thành và phát triển của xương và răng. Thừa M có thể dẫn tới sỏi thận. Cho 3,6 gam M tác dụng hết với HCl, thu được 2,2311 lít khí (đo ở điều kiện chuẩn).")

    p_tl2_sub = doc.add_paragraph()
    p_tl2_sub.paragraph_format.left_indent = Cm(0.5)
    p_tl2_sub.paragraph_format.space_before = Pt(1)
    p_tl2_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_tl2_sub, "a. Xác định M và cho biết vị trí của M trong bảng tuần hoàn.\nb. So sánh tính kim loại của M với _19_K và _12_Mg. Giải thích.")

    if is_teacher:
        sol_tl2 = (
            "a. Xác định M và vị trí trong bảng tuần hoàn:\n"
            "- Ở điều kiện chuẩn (25°C, 1 bar): n~H2~ = 2,2311 / 24,79 = 0,09 mol.\n"
            "- Phương trình phản ứng:\n"
            "  2M + 2nHCl ➔ 2MCl~n~ + nH~2~ ↑\n"
            "- Theo phương trình: n~M~ = (2 / n) × n~H2~ = 0,18 / n mol.\n"
            "- Khối lượng mol của M: M = 3,6 / (0,18 / n) = 20n (g/mol).\n"
            "- Vì M là kim loại thuộc nhóm A ➔ n ∈ {1, 2, 3}:\n"
            "  + n = 1 ➔ M = 20 (loại)\n"
            "  + n = 2 ➔ M = 40 (g/mol) ➔ M là Calcium (Ca, Z = 20). Thỏa mãn là thành phần của xương, răng và thừa gây sỏi thận.\n"
            "  + n = 3 ➔ M = 60 (loại)\n"
            "- Vị trí của Ca (Z = 20) trong bảng tuần hoàn:\n"
            "  + Cấu hình electron: 1s^2^2s^2^2p^6^3s^2^3p^6^4s^2^ (hoặc [Ar]4s^2^).\n"
            "  + Ô số 20 (vì Z = 20).\n"
            "  + Chu kì 4 (vì có 4 lớp electron).\n"
            "  + Nhóm IIA (vì là nguyên tố s, có 2 electron hóa trị ở lớp ngoài cùng 4s^2^).\n"
            "b. So sánh tính kim loại của Ca với _19_K và _12_Mg:\n"
            "- Vị trí: Mg (Chu kì 3, nhóm IIA); Ca (Chu kì 4, nhóm IIA); K (Chu kì 4, nhóm IA).\n"
            "- So sánh:\n"
            "  + Trong nhóm IIA, theo chiều tăng điện tích hạt nhân (Mg ➔ Ca), bán kính tăng nên tính kim loại tăng: Ca > Mg.\n"
            "  + Trong chu kì 4, theo chiều tăng điện tích hạt nhân (K ➔ Ca), tính kim loại giảm: K > Ca.\n"
            "➔ Thứ tự tính kim loại: K > Ca > Mg (tính kim loại của Ca mạnh hơn Mg và yếu hơn K)."
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=10)

    # Dòng kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "*---------- HẾT ----------*", base_size=10.5, base_bold=True)

    doc.save(output_path)
    print(f"Đã lưu thành công: {output_path}")

if __name__ == "__main__":
    out_dir = "San_Pham"
    os.makedirs(out_dir, exist_ok=True)
    f_hs = os.path.join(out_dir, "De_Kiem_Tra_CK1_Hoa_10_HoaChau_Ma132_HocSinh.docx")
    f_gv = os.path.join(out_dir, "De_Kiem_Tra_CK1_Hoa_10_HoaChau_Ma132_GiaoVien.docx")
    create_exam_file(is_teacher=False, output_path=f_hs)
    create_exam_file(is_teacher=True, output_path=f_gv)
