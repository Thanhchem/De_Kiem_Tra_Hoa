# -*- coding: utf-8 -*-
"""
Script tạo 2 file Word chuẩn Form Đề Kiểm Tra Hóa học (Bản Học Sinh & Bản Giáo Viên)
Đề: TRƯỜNG THPT ĐẶNG HUY TRỨ - MÃ ĐỀ 104 - CUỐI HỌC KÌ I - LỚP 10 (2025-2026)
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
        school_line="SỞ GD - ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT ĐẶNG HUY TRỨ",
        exam_title="KIỂM TRA CUỐI KỲ I – NĂM HỌC 2025-2026",
        subject_line="Môn: HÓA HỌC – Lớp 10",
        num_pages="03",
        duration="45 phút",
        exam_code="104",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; Si = 28; P = 31; S = 32; Cl = 35,5; K = 19; Ca = 40; Fe = 56; Br = 80; Ba = 137",
        is_teacher=is_teacher
    )

    # -------------------------------------------------------------
    # BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO BẢN GIÁO VIÊN THEO MẪU ẢNH CHUẨN)
    # -------------------------------------------------------------
    if is_teacher:
        from build_bang_dap_an_rut_gon import add_styled_answer_block, EXAMS_DATA
        add_styled_answer_block(doc, EXAMS_DATA[2])

    # -------------------------------------------------------------
    # PHẦN I
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="A. TRẮC NGHIỆM (7,0 điểm)",
        subtitle="Phần I: Trắc nghiệm nhiều lựa chọn (4,0 điểm). Học sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi học sinh chỉ chọn một phương án."
    )

    questions_p1 = [
        {
            "num": 1,
            "prompt": "Câu 1. Nguyên tử nguyên tố X có phân lớp electron ngoài cùng là 3p^4^. Công thức oxide, hydroxide ứng với hóa trị cao nhất của X, và tính acid – base tương ứng của chúng là",
            "opts": [
                "A. XO, X(OH)~2~, tính base.",
                "B. XO~2~, H~2~XO~3~, tính acid.",
                "*C. XO~3~, H~2~XO~4~, tính acid.*",
                "D. X~2~O~3~, X(OH)~3~, tính lưỡng tính."
            ],
            "sol": "Cấu hình electron của X: 1s^2^2s^2^2p^6^3s^2^3p^4^ (Sulfur - S, Z = 16). X thuộc nhóm VIA nên có hóa trị cao nhất với oxygen là VI.\n- Công thức oxide cao nhất là XO~3~ (SO~3~), thuộc loại acidic oxide.\n- Công thức hydroxide cao nhất là H~2~XO~4~ (H~2~SO~4~), có tính acid mạnh.\n➔ Chọn đáp án C."
        },
        {
            "num": 2,
            "prompt": "Câu 2. Hợp chất nào sau đây có liên kết cộng hóa trị không phân cực?",
            "opts": [
                "A. CF~2~Cl~2~.",
                "B. LiCl.",
                "C. CHCl~3~.",
                "*D. N~2~.*"
            ],
            "sol": "Phân tử N~2~ được tạo thành từ 2 nguyên tử Nitrogen giống nhau nên hiệu độ âm điện Δχ = 0 ➔ Liên kết giữa hai nguyên tử N là liên kết cộng hóa trị không phân cực.\n(CF~2~Cl~2~ và CHCl~3~ có liên kết CHT phân cực; LiCl có liên kết ion).\n➔ Chọn đáp án D."
        },
        {
            "num": 3,
            "prompt": "Câu 3. Hạt nhân nguyên tử của hầu hết các nguyên tử được hạt cấu tạo bởi các loại hạt",
            "opts": [
                "*A. Neutron và proton.*",
                "B. Electron, neutron và proton.",
                "C. Electron và proton.",
                "D. Electron và neutron."
            ],
            "sol": "Hạt nhân của hầu hết các nguyên tử được cấu tạo từ hai loại hạt là proton (mang điện tích dương) và neutron (không mang điện). (Trừ nguyên tử ^1^H hạt nhân chỉ có 1 proton).\n➔ Chọn đáp án A."
        },
        {
            "num": 4,
            "prompt": "Câu 4. Phát biểu nào sau đây không đúng khi nói về mô hình Rutherford – Bohr?",
            "opts": [
                "A. Electron ở lớp M xa hạt nhân hơn so với electron ở lớp L.",
                "B. Electron ở lớp K gần hạt nhân hơn so với electron ở lớp L.",
                "C. Electron trên lớp M có năng lượng cao hơn trên lớp K.",
                "*D. Electron trên lớp K có năng lượng cao hơn trên lớp L.*"
            ],
            "sol": "Theo mô hình Rutherford - Bohr, các electron chuyển động trên các lớp electron xác định (K, L, M, N... tương ứng với n = 1, 2, 3, 4...). Lớp K (n = 1) ở gần hạt nhân nhất và có mức năng lượng thấp nhất. Càng xa hạt nhân năng lượng của electron càng cao (E~K~ < E~L~ < E~M~...). Do đó phát biểu D là KHÔNG ĐÚNG.\n➔ Chọn đáp án D."
        },
        {
            "num": 5,
            "prompt": "Câu 5. Số lượng electron tham gia hình thành liên kết đơn, đôi và ba lần lượt là:",
            "opts": [
                "A. 2, 3 và 4",
                "B. 1, 2 và 3",
                "*C. 2, 4 và 6*",
                "D. 1, 3 và 4"
            ],
            "sol": "- Liên kết đơn được hình thành bởi 1 cặp electron dùng chung = 2 electron.\n- Liên kết đôi được hình thành bởi 2 cặp electron dùng chung = 4 electron.\n- Liên kết ba được hình thành bởi 3 cặp electron dùng chung = 6 electron.\n➔ Chọn đáp án C."
        },
        {
            "num": 6,
            "prompt": "Câu 6. Cho các nguyên tố sau: K (Z = 19), N (Z = 7), Si (Z = 14), Mg (Z = 12). Dãy gồm các nguyên tố được sắp xếp theo chiều giảm dần bán kính nguyên tử từ trái sang phải là",
            "opts": [
                "A. Mg, K, Si, N.",
                "B. N, Si, Mg, K.",
                "*C. K, Mg, Si, N.*",
                "D. K, Mg, N, Si."
            ],
            "sol": "- K (Z=19): Chu kì 4 (có 4 lớp e) ➔ Bán kính nguyên tử lớn nhất.\n- Mg (Z=12) và Si (Z=14): Cùng chu kì 3 (có 3 lớp e). Trong cùng chu kì, theo chiều tăng điện tích hạt nhân bán kính giảm dần: r(Mg) > r(Si).\n- N (Z=7): Chu kì 2 (có 2 lớp e) ➔ Bán kính nhỏ nhất.\n➔ Thứ tự bán kính nguyên tử giảm dần: K > Mg > Si > N.\n➔ Chọn đáp án C."
        },
        {
            "num": 7,
            "prompt": "Câu 7. Khi nói về hợp chất tạo thành giữa Na^+^ và O^2-^, một học sinh đưa ra các nhận định sau:\n(1). Là hợp chất ion.\n(2). Có công thức hoá học là NaO.\n(3). Trong điều kiện thường, tồn tại ở thể khí.\n(4). Trong điều kiện thường, tồn tại ở thể rắn.\n(5). Có nhiệt độ nóng chảy và nhiệt độ sôi cao.\n(6). Lực tương tác giữa Na^+^ và O^2-^ là lực tĩnh điện.\nCó bao nhiêu nhận định đúng?",
            "opts": [
                "A. 3",
                "*B. 4*",
                "C. 2",
                "D. 5"
            ],
            "sol": "Hợp chất tạo thành giữa Na^+^ và O^2-^ là Na~2~O (sodium oxide):\n- (1) ĐÚNG: Na~2~O là hợp chất ion.\n- (2) SAI: Công thức hóa học là Na~2~O.\n- (3) SAI, (4) ĐÚNG: Hợp chất ion ở điều kiện thường tồn tại ở trạng thái rắn (tinh thể ion).\n- (5) ĐÚNG: Mạng tinh thể ion bền vững nên có nhiệt độ nóng chảy và nhiệt độ sôi cao.\n- (6) ĐÚNG: Bản chất liên kết ion là lực hút tĩnh điện giữa các ion mang điện tích trái dấu.\n➔ Có 4 nhận định đúng: (1), (4), (5), (6).\n➔ Chọn đáp án B."
        },
        {
            "num": 8,
            "prompt": "Câu 8. Hợp chất nào sau đây có liên kết ion?",
            "opts": [
                "*A. KCl.*",
                "B. H~2~O.",
                "C. HCl.",
                "D. NH~3~."
            ],
            "sol": "KCl được tạo thành từ kim loại kiềm điển hình (K) và phi kim halogen điển hình (Cl), liên kết hình thành bởi lực hút tĩnh điện giữa cation K^+^ và anion Cl^-^ ➔ Liên kết ion.\n(Các chất H~2~O, HCl, NH~3~ đều có liên kết cộng hóa trị phân cực).\n➔ Chọn đáp án A."
        },
        {
            "num": 9,
            "prompt": "Câu 9. Biểu diễn sự tạo thành ion nào sau đây đúng?",
            "opts": [
                "*A. Ca ➔ Ca^2+^ + 2e.*",
                "B. O~2~ + 2e ➔ 2O^2-^.",
                "C. Cl~2~ ➔ 2Cl^-^ + 2e.",
                "D. Na + 1e ➔ Na^+^."
            ],
            "sol": "- Ca (Z = 20) là kim loại có 2e lớp ngoài cùng, nhường 2e để tạo cation Ca^2+^: Ca ➔ Ca^2+^ + 2e (Đúng).\n- Phương án B sai vì O~2~ + 4e ➔ 2O^2-^.\n- Phương án C sai vì Cl~2~ + 2e ➔ 2Cl^-^.\n- Phương án D sai vì Na ➔ Na^+^ + 1e.\n➔ Chọn đáp án A."
        },
        {
            "num": 10,
            "prompt": "Câu 10. Có hai nguyên tố X (Z = 19) và Y (Z = 17). Hợp chất tạo bởi X và Y có công thức và kiểu liên kết là:",
            "opts": [
                "A. XY~2~, liên kết cộng hóa trị có cực.",
                "*B. XY, liên kết ion.*",
                "C. X~2~Y, liên kết ion.",
                "D. XY, liên kết cộng hóa trị có cực."
            ],
            "sol": "- X (Z = 19): [Ar]4s^1^ (kim loại kiềm K, nhóm IA), nhường 1e tạo ion X^+^.\n- Y (Z = 17): [Ne]3s^2^3p^5^ (phi kim halogen Cl, nhóm VIIA), nhận 1e tạo ion Y^-^.\n➔ Hợp chất tạo thành là XY (KCl), liên kết ion.\n➔ Chọn đáp án B."
        },
        {
            "num": 11,
            "prompt": "Câu 11. Các orbital s và p có dạng hình cầu và dạng hình số 8 nổi và có sự định hướng trong không gian. Orbital p chứa tối đa bao nhiêu electron?",
            "opts": [
                "*A. 2 electron.*",
                "B. 10 electron.",
                "C. 8 electron.",
                "D. 6 electron."
            ],
            "sol": "Theo nguyên lí Pauli, mỗi orbital nguyên tử (AO) dù là s, p, d hay f chỉ chứa tối đa 2 electron có spin ngược chiều nhau. Do đó, một orbital p chứa tối đa 2 electron.\n(Lưu ý: Phân lớp p gồm 3 orbital p~x~, p~y~, p~z~ chứa tối đa 6 electron).\n➔ Chọn đáp án A."
        },
        {
            "num": 12,
            "prompt": "Câu 12. Khi nhường hoặc nhận electron, nguyên tử nào sau đây có cấu trúc electron không thỏa mãn quy tắc octet?",
            "opts": [
                "*A. Hydrogen.*",
                "B. Oxygen.",
                "C. Nitrogen.",
                "D. Sodium."
            ],
            "sol": "Nguyên tử Hydrogen (Z = 1) khi tham gia liên kết hoặc tạo ion chỉ có xu hướng đạt cấu hình bền vững của khí hiếm Helium với 2 electron ở lớp ngoài cùng (quy tắc duplet / octet 2 electron) chứ không phải 8 electron (octet).\n➔ Chọn đáp án A."
        },
        {
            "num": 13,
            "prompt": "Câu 13. Cấu hình electron của nguyên tử oxygen là 1s^2^2s^2^2p^4^. Vị trí của oxygen trong bảng tuần hoàn là:",
            "opts": [
                "*A. ô số 8, chu kì 2, nhóm VIA.*",
                "B. ô số 6, chu kì 3, nhóm VIB.",
                "C. ô số 6, chu kì 2, nhóm VIA.",
                "D. ô số 8, chu kì 2, nhóm VIB."
            ],
            "sol": "- Số hiệu nguyên tử Z = 8 ➔ Ô số 8.\n- Có 2 lớp electron ➔ Chu kì 2.\n- Có 6 electron lớp ngoài cùng (2s^2^2p^4^), nguyên tố p ➔ Nhóm VIA.\n➔ Chọn đáp án A."
        },
        {
            "num": 14,
            "prompt": "Câu 14. Phát biểu nào sau đây không đúng?",
            "opts": [
                "A. Có hai kiểu xen phủ hình thành liên kết là xen phủ trục và xen phủ bên.",
                "*B. Chỉ có các AO (orbital nguyên tử) có hình dạng giống nhau mới xen phủ với nhau để tạo liên kết.*",
                "C. Khi hình thành liên kết cộng hoá trị giữa hai nguyên tử, luôn có một liên kết σ.",
                "D. Liên kết σ bền vững hơn liên kết π."
            ],
            "sol": "Phát biểu B sai vì các AO có hình dạng khác nhau vẫn có thể xen phủ với nhau tạo liên kết, ví dụ: AO s hình cầu xen phủ với AO p hình số 8 nổi tạo liên kết cộng hóa trị s-p (như trong phân tử HF, HCl...).\n➔ Chọn đáp án B."
        },
        {
            "num": 15,
            "prompt": "Câu 15. Cho các nguyên tố: _13_Al, _12_Mg, _19_K, _11_Na. Dãy các nguyên tố nào sau đây có tính kim loại giảm dần?",
            "opts": [
                "A. Na > K > Mg > Al.",
                "*B. K > Na > Mg > Al.*",
                "C. Al > Mg > Na > K.",
                "D. Mg > Al > K > Na."
            ],
            "sol": "- Trong chu kì 3: theo chiều tăng dần của điện tích hạt nhân, tính kim loại giảm dần: Na > Mg > Al.\n- Trong nhóm IA: theo chiều tăng dần của điện tích hạt nhân, tính kim loại tăng dần: K > Na.\n➔ Thứ tự tính kim loại giảm dần: K > Na > Mg > Al.\n➔ Chọn đáp án B."
        },
        {
            "num": 16,
            "prompt": "Câu 16. Nguyên tử nguyên tố X có 32 hạt mang điện. Điện tích hạt nhân của nguyên tử X là",
            "opts": [
                "A. -32.",
                "B. +32.",
                "*C. +16.*",
                "D. -16."
            ],
            "sol": "Hạt mang điện trong nguyên tử gồm proton và electron (p = e = Z).\nTổng số hạt mang điện = 2Z = 32 ➔ Z = 16.\nĐiện tích hạt nhân là +Ze = +16.\n➔ Chọn đáp án C."
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
        title="Phần II: Trắc nghiệm dạng Đúng/Sai (2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai."
    )

    # Câu 1 phần II
    p_q1_p2 = doc.add_paragraph()
    p_q1_p2.paragraph_format.space_before = Pt(3)
    p_q1_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q1_p2, "Câu 1. \"Nước\" là tên trạng thái lỏng của H~2~O ở điều kiện tiêu chuẩn về nhiệt độ và áp suất. Nước bao phủ 71% bề mặt Trái đất, chủ yếu ở các biển và đại dương. Nước đóng một vai trò quan trọng trong nền kinh tế thế giới. Khoảng 70% lượng nước ngọt mà con người sử dụng dùng cho nông nghiệp. Thủy sản, hải sản ở các vùng nước mặn và nước ngọt là nguồn cung cấp thực phẩm chính cho nhiều nơi trên thế giới. Cho số hiệu nguyên tử của H là 1 và O là 8; giá trị độ âm điện của H là 2,20 và O là 3,44.")

    stmts_c1 = [
        ("a. Liên kết H – O là liên kết cộng hoá trị phân cực.", "Đúng"),
        ("b. Cặp electron dùng chung trong liên kết H – O phân bố đều giữa hai nguyên tử.", "Sai"),
        ("c. Nguyên tử O còn bốn cặp electron hoá trị riêng ( electron chưa tham gia vào liên kết).", "Sai"),
        ("d. Nước là dung môi phân cực có thể hòa tan các chất phân cực.", "Đúng")
    ]
    for stmt, ans in stmts_c1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        sol_c1_p2 = (
            "a. ĐÚNG: Hiệu độ âm điện Δχ(O-H) = 3,44 - 2,20 = 1,24 (0,4 ≤ 1,24 < 1,7) ➔ Liên kết cộng hóa trị phân cực.\n"
            "b. SAI: Do O có độ âm điện lớn hơn H nên cặp electron dùng chung bị lệch về phía nguyên tử O (phân bố không đều).\n"
            "c. SAI: Nguyên tử O có 6e hóa trị, dùng 2e tạo 2 liên kết với 2 nguyên tử H ➔ còn lại 4 electron hóa trị chưa liên kết, tương ứng với 2 cặp electron tự do (chứ không phải 4 cặp).\n"
            "d. ĐÚNG: Phân tử H~2~O có cấu trúc góc nên là phân tử phân cực, có thể hòa tan tốt các chất phân cực và hợp chất ion."
        )
        add_solution_no_table(doc, sol_c1_p2)

    # Câu 2 phần II
    p_q2_p2 = doc.add_paragraph()
    p_q2_p2.paragraph_format.space_before = Pt(4)
    p_q2_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q2_p2, "Câu 2. F (Z = 9), Cl (Z = 17), Br (Z = 35) đều thuộc nhóm VIIA (nhóm halogen) trong bảng tuần hoàn. Đơn chất halogen tồn tại dạng phân tử X~2~. Hợp chất hydrogen halide có công thức chung HX.")

    stmts_c2 = [
        ("a. Cấu hình electron lớp ngoài cùng của nguyên tử X có dạng ns^2^np^5^.", "Đúng"),
        ("b. Tính phi kim được sắp xếp theo chiều giảm dần : Br> Cl> F .", "Sai"),
        ("c. Liên kết giữa các nguyên tử trong HX là liên kết cộng hoá trị phân cực.", "Đúng"),
        ("d. Trong hợp chất khí với Hydrogen, X chiếm 97,26% về khối lượng thì X là Br.", "Sai")
    ]
    for stmt, ans in stmts_c2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        sol_c2_p2 = (
            "a. ĐÚNG: Các nguyên tố halogen thuộc nhóm VIIA đều có 7 electron lớp ngoài cùng dạng ns^2^np^5^.\n"
            "b. SAI: Trong một nhóm A, từ trên xuống dưới tính phi kim giảm dần: F > Cl > Br (phát biểu ghi ngược lại).\n"
            "c. ĐÚNG: Các nguyên tử halogen có độ âm điện lớn hơn H nên liên kết H-X là liên kết cộng hóa trị phân cực.\n"
            "d. SAI: %m_X = [X / (X + 1)] × 100% = 97,26% ➔ X ≈ 35,5 (Chlorine - Cl). (Nếu là Br thì %m_Br = 80/81 = 98,77%)."
        )
        add_solution_no_table(doc, sol_c2_p2)

    # -------------------------------------------------------------
    # PHẦN III
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="Phần III: Trắc nghiệm trả lời ngắn (1,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4. Mỗi câu đúng được 0,25 điểm."
    )

    questions_p3 = [
        {
            "num": 1,
            "prompt": "Câu 1. Nguyên tố sulfur (lưu huỳnh) nằm ở ô thứ 16 trong bảng hệ thống tuần hoàn. Số electron thuộc lớp L trong nguyên tử sulfur là bao nhiêu?",
            "ans": "8",
            "sol": "Cấu hình electron của S (Z = 16): 1s^2^2s^2^2p^6^3s^2^3p^4^.\n- Lớp K (n = 1): 2 electron (1s^2^).\n- Lớp L (n = 2): 2 + 6 = 8 electron (2s^2^2p^6^).\n- Lớp M (n = 3): 2 + 4 = 6 electron (3s^2^3p^4^).\n➔ Số electron thuộc lớp L là 8."
        },
        {
            "num": 2,
            "prompt": "Câu 2. Cho các ion sau. K^+^, N^3-^, Cl^-^, O^2-^, Mg^2+^ và Ca^2+^. Có bao nhiêu ion có cấu hình electron của khí hiếm argon? (cho điện tích hạt nhân của K, N, Cl, O, Mg, Ca lần lượt là 19, 7, 17, 8, 12, 20).",
            "ans": "3",
            "sol": "Khí hiếm argon (Ar) có 18 electron.\nXét số electron của từng ion:\n- K^+^ (Z=19): 19 - 1 = 18e (cấu hình Ar).\n- N^3-^ (Z=7): 7 + 3 = 10e (cấu hình Ne).\n- Cl^-^ (Z=17): 17 + 1 = 18e (cấu hình Ar).\n- O^2-^ (Z=8): 8 + 2 = 10e (cấu hình Ne).\n- Mg^2+^ (Z=12): 12 - 2 = 10e (cấu hình Ne).\n- Ca^2+^ (Z=20): 20 - 2 = 18e (cấu hình Ar).\n➔ Có 3 ion có cấu hình electron của khí hiếm argon là: K^+^, Cl^-^, Ca^2+^."
        },
        {
            "num": 3,
            "prompt": "Câu 3. Cho các chất sau: KCl, H~2~O, Na~2~O, BaCl~2~, CaF~2~, HCl, (NH~4~)~2~SO~4~. Trong số các hợp chất trên có bao nhiêu hợp chất phân tử có liên kết ion?",
            "ans": "5",
            "sol": "Các hợp chất có liên kết ion gồm:\n1. KCl (tạo bởi K^+^ và Cl^-^)\n2. Na~2~O (tạo bởi Na^+^ và O^2-^)\n3. BaCl~2~ (tạo bởi Ba^2+^ và Cl^-^)\n4. CaF~2~ (tạo bởi Ca^2+^ và F^-^)\n5. (NH~4~)~2~SO~4~ (tạo bởi NH~4~^+^ và SO~4~^2-^)\n(H~2~O và HCl chỉ chứa liên kết cộng hóa trị phân cực).\n➔ Có 5 hợp chất có liên kết ion."
        },
        {
            "num": 4,
            "prompt": "Câu 4. Cho các phát biểu sau về phân tử CO~2~:\n(a) Liên kết giữa hai nguyên tử C và O là liên kết cộng hoá trị không phân cực\n(b) Liên kết giữa hai nguyên tử C và O là liên kết cộng hoá trị phân cực\n(c) Phân tử CO~2~, mỗi nguyên tử oxygen tạo 2 cặp electron chung với nguyên tử carbon.\n(d) Trong phân tử CO~2~ có 2 liên kết σ và 2 liên kết π.\n(e) Trong phân tử CO~2~ có cấu tạo thẳng.\nSố phát biểu đúng là bao nhiêu?",
            "ans": "4",
            "sol": "Công thức cấu tạo của CO~2~ là O = C = O:\n- (a) SAI: Liên kết C=O là liên kết CHT phân cực do O có độ âm điện lớn hơn C.\n- (b) ĐÚNG: Δχ(O-C) = 3,44 - 2,55 = 0,89 ➔ Liên kết CHT phân cực.\n- (c) ĐÚNG: Mỗi nguyên tử O góp 2 electron tạo 2 cặp electron dùng chung (liên kết đôi) với nguyên tử C.\n- (d) ĐÚNG: Phân tử gồm 2 liên kết đôi, mỗi liên kết đôi gồm 1 liên kết σ và 1 liên kết π ➔ Tổng có 2 liên kết σ và 2 liên kết π.\n- (e) ĐÚNG: Phân tử CO~2~ có cấu tạo đường thẳng (góc liên kết 180°).\n➔ Có 4 phát biểu đúng là (b), (c), (d), (e)."
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
        title="B. TỰ LUẬN (3,0 điểm)",
        subtitle="Học sinh trình bày lời giải chi tiết cho các câu hỏi sau:"
    )

    # Câu 1 Tự luận
    p_tl1 = doc.add_paragraph()
    p_tl1.paragraph_format.space_before = Pt(3)
    p_tl1.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_tl1, "Câu 1. (1,25 điểm) X, Y là hai nguyên tố liên tiếp trong một nhóm A (Z~X~ < Z~Y~) và có tổng số đơn vị điện tích hạt nhân là 32, trong X và Y đều có số neutron bằng số proton.")

    p_tl1_sub = doc.add_paragraph()
    p_tl1_sub.paragraph_format.left_indent = Cm(0.5)
    p_tl1_sub.paragraph_format.space_before = Pt(1)
    p_tl1_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_tl1_sub, "a) So sánh bán kính của X và Y. Giải thích.\nb) Cho 22 gam hỗn hợp X, Y tác dụng với dung dịch HCl dư thu được 18,5925 lít khí (đkc). Xác định khối lượng của X và Y trong hỗn hợp?")

    if is_teacher:
        sol_tl1 = (
            "*Xác định nguyên tố X và Y:*\n"
            "- Vì X, Y thuộc cùng một nhóm A ở hai chu kì liên tiếp (Z~X~ < Z~Y~) nên hiệu số điện tích hạt nhân giữa chúng thường là 8:\n"
            "  { Z~X~ + Z~Y~ = 32\n"
            "  { Z~Y~ - Z~X~ = 8\n"
            "  ➔ Z~X~ = 12 (Magnesium - Mg) và Z~Y~ = 20 (Calcium - Ca).\n"
            "  (Cả Mg và Ca đều thuộc nhóm IIA, có số neutron bằng proton: A(Mg) = 24, A(Ca) = 40, hoàn toàn thỏa mãn giả thiết).\n"
            "a) So sánh bán kính của X và Y:\n"
            "- Mg (Z = 12) có 3 lớp electron: 1s^2^2s^2^2p^6^3s^2^.\n"
            "- Ca (Z = 20) có 4 lớp electron: 1s^2^2s^2^2p^6^3s^2^3p^6^4s^2^.\n"
            "- Trong cùng nhóm IIA, theo chiều tăng điện tích hạt nhân (Mg ➔ Ca), số lớp electron tăng dần nên bán kính nguyên tử tăng dần:\n"
            "  ➔ Bán kính nguyên tử của Ca lớn hơn Mg (r~Ca~ > r~Mg~).\n"
            "b) Xác định khối lượng của X và Y trong hỗn hợp:\n"
            "- Ở điều kiện chuẩn: n~H2~ = 18,5925 / 24,79 = 0,75 mol.\n"
            "- Gọi x, y lần lượt là số mol của Mg và Ca trong hỗn hợp (x, y > 0).\n"
            "- Phương trình hóa học:\n"
            "  Mg + 2HCl ➔ MgCl~2~ + H~2~ ↑\n"
            "  Ca + 2HCl ➔ CaCl~2~ + H~2~ ↑\n"
            "- Ta có hệ phương trình:\n"
            "  { 24x + 40y = 22 (khối lượng hỗn hợp)\n"
            "  { x + y = 0,75 (số mol H~2~)\n"
            "  ➔ Giải hệ ta được: x = 0,5 mol và y = 0,25 mol.\n"
            "- Khối lượng của từng kim loại:\n"
            "  + m~Mg~ = 0,5 × 24 = 12 gam.\n"
            "  + m~Ca~ = 0,25 × 40 = 10 gam."
        )
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Tự luận
    p_tl2 = doc.add_paragraph()
    p_tl2.paragraph_format.space_before = Pt(4)
    p_tl2.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_tl2, "Câu 2. (1,0 điểm) Nguyên tử của một nguyên tố X có cấu hình electron: [Ne]3s^1^. Nguyên tố này là chất điện giải, giúp cân bằng nồng độ chất lỏng, giữ nước cho cơ thể. Ngoài ra X còn cần thiết cho việc duy trì huyết áp ổn định.")

    p_tl2_sub = doc.add_paragraph()
    p_tl2_sub.paragraph_format.left_indent = Cm(0.5)
    p_tl2_sub.paragraph_format.space_before = Pt(1)
    p_tl2_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_tl2_sub, "(a) Hãy cho biết tính chất hóa học cơ bản của X là gì? (là kim loại hay phi kim, mạnh hay yếu). Viết công thức oxide và hydroxide cao nhất của X, cho biết chúng có tính acid hay base?\n(b) Viết phương trình phản ứng hóa học khi cho oxide và hydroxide của X lần lượt tác dụng với H~2~O, HCl (nếu có).")

    if is_teacher:
        sol_tl2 = (
            "Cấu hình electron của X: [Ne]3s^1^ ➔ Z = 11 (nguyên tố Sodium - Na).\n"
            "(a) Tính chất hóa học cơ bản và công thức oxide, hydroxide:\n"
            "- Na có 1 electron ở lớp ngoài cùng, thuộc nhóm IA nên tính chất hóa học cơ bản là tính kim loại mạnh (kim loại điển hình).\n"
            "- Công thức oxide cao nhất: Na~2~O, mang tính base (basic oxide).\n"
            "- Công thức hydroxide cao nhất: NaOH, mang tính base mạnh (base kiềm mạnh).\n"
            "(b) Phương trình phản ứng hóa học:\n"
            "- Với oxide (Na~2~O):\n"
            "  + Tác dụng với H~2~O: Na~2~O + H~2~O ➔ 2NaOH\n"
            "  + Tác dụng với HCl: Na~2~O + 2HCl ➔ 2NaCl + H~2~O\n"
            "- Với hydroxide (NaOH):\n"
            "  + Hòa tan trong H~2~O: NaOH tan vô hạn trong nước, phân li thành ion (NaOH ➔ Na^+^ + OH^-^).\n"
            "  + Tác dụng với HCl: NaOH + HCl ➔ NaCl + H~2~O"
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 3 Tự luận
    p_tl3 = doc.add_paragraph()
    p_tl3.paragraph_format.space_before = Pt(4)
    p_tl3.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_tl3, "Câu 3. (0,75 điểm) Hợp chất khí A tạo bởi nguyên tố R và hydrogen có dạng RH~3~, được sử dụng để trung hoà các thành phần acid của dầu thô, bảo vệ thiết bị không bị ăn mòn trong ngành công nghiệp dầu khí, sản xuất phân bón,... Nguyên tố này chiếm 25,93% về khối lượng trong oxide cao nhất.")

    p_tl3_sub = doc.add_paragraph()
    p_tl3_sub.paragraph_format.left_indent = Cm(0.5)
    p_tl3_sub.paragraph_format.space_before = Pt(1)
    p_tl3_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_tl3_sub, "a) Nguyên tử khối của R là bao nhiêu?\nb) Liên kết trong A thuộc loại liên kết gì ? Viết công thức Lewis và công thức cấu tạo của A.")

    if is_teacher:
        sol_tl3 = (
            "a) Xác định nguyên tử khối của R:\n"
            "- Hợp chất khí với hydrogen là RH~3~ ➔ R thuộc nhóm VA ➔ Oxide cao nhất là R~2~O~5~.\n"
            "- Phần trăm khối lượng của R trong oxide cao nhất:\n"
            "  %m_R = [2R / (2R + 5 × 16)] × 100% = 25,93%\n"
            "  ➔ 2R / (2R + 80) = 0,2593 ➔ 2R = 0,5186 R + 20,744 ➔ 1,4814 R = 20,744 ➔ R ≈ 14.\n"
            "  Vậy nguyên tử khối của R là 14 (R là Nitrogen - N). Hợp chất khí A là NH~3~.\n"
            "b) Loại liên kết, công thức Lewis và công thức cấu tạo của A (NH~3~):\n"
            "- Hiệu độ âm điện giữa N (3,04) và H (2,20) là Δχ = 3,04 - 2,20 = 0,84 (0,4 ≤ 0,84 < 1,7) ➔ Liên kết giữa N và H là liên kết cộng hóa trị phân cực.\n"
            "- Công thức Lewis của NH~3~: H - N̈(H) - H (nguyên tử N ở trung tâm liên kết đơn với 3 nguyên tử H và còn 1 cặp electron tự do).\n"
            "- Công thức cấu tạo của NH~3~: H - N(H) - H"
        )
        add_solution_no_table(doc, sol_tl3)
    else:
        add_student_writing_lines(doc, num_lines=7)

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
    f_hs = os.path.join(out_dir, "De_Kiem_Tra_CK1_Hoa_10_DangHuyTru_Ma104_HocSinh.docx")
    f_gv = os.path.join(out_dir, "De_Kiem_Tra_CK1_Hoa_10_DangHuyTru_Ma104_GiaoVien.docx")
    create_exam_file(is_teacher=False, output_path=f_hs)
    create_exam_file(is_teacher=True, output_path=f_gv)
