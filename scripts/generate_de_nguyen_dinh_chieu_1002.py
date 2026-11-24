# -*- coding: utf-8 -*-
"""
Script tạo 2 file Word chuẩn Form Đề Kiểm Tra Hóa học (Bản Học Sinh & Bản Giáo Viên)
Đề: TRƯỜNG THPT NGUYỄN ĐÌNH CHIỂU - MÃ ĐỀ 1002 - CUỐI HỌC KÌ I - LỚP 10 (2025-2026)
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
        school_line="SỞ GD&ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT NGUYỄN ĐÌNH CHIỂU",
        exam_title="ĐỀ KIỂM TRA CUỐI HỌC KÌ I - NĂM HỌC 2025 - 2026",
        subject_line="MÔN: HÓA HỌC - KHỐI 10",
        num_pages="02",
        duration="45 phút",
        exam_code="1002",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; Si = 28; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; I = 127",
        is_teacher=is_teacher
    )

    # -------------------------------------------------------------
    # BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO BẢN GIÁO VIÊN ĐẶT NGAY ĐẦU HOẶC CUỐI)
    # -------------------------------------------------------------
    if is_teacher:
        p_ans_title = doc.add_paragraph()
        p_ans_title.paragraph_format.space_before = Pt(4)
        p_ans_title.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans_title, "*BẢNG ĐÁP ÁN RÚT GỌN (TRẮC NGHIỆM & TỰ LUẬN)*", base_size=11, base_bold=True, base_color=COLOR_PRIMARY)

        # Bảng đáp án trắc nghiệm Phần I
        p_p1 = doc.add_paragraph()
        p_p1.paragraph_format.space_before = Pt(2)
        p_p1.paragraph_format.space_after = Pt(1)
        render_rich_text(p_p1, "*Phần I (Mỗi câu 0,25 điểm):*", base_size=10, base_bold=True)

        tbl1 = doc.add_table(rows=2, cols=12)
        tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
        ans_p1 = ["1. C", "2. A", "3. A", "4. D", "5. C", "6. D", "7. A", "8. B", "9. B", "10. B", "11. C", "12. D"]
        for col_idx in range(12):
            cell_top = tbl1.cell(0, col_idx)
            cell_top.text = f"C{col_idx+1}"
            cell_top.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run_style(cell_top.paragraphs[0].runs[0], font_size=9.5, bold=True)

            cell_bot = tbl1.cell(1, col_idx)
            cell_bot.text = ans_p1[col_idx].split(". ")[1]
            cell_bot.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run_style(cell_bot.paragraphs[0].runs[0], font_size=9.5, bold=True, color=COLOR_CORRECT)

        # Đáp án Phần II & Phần III & Phần Tự luận
        p_p2 = doc.add_paragraph()
        p_p2.paragraph_format.space_before = Pt(3)
        p_p2.paragraph_format.space_after = Pt(1)
        render_rich_text(p_p2, "*Phần II (Đúng/Sai):* Câu 1: *a-Đ, b-Đ, c-S, d-Đ*  |  Câu 2: *a-Đ, b-S, c-Đ, d-S*", base_size=9.5)

        p_p3 = doc.add_paragraph()
        p_p3.paragraph_format.space_before = Pt(1)
        p_p3.paragraph_format.space_after = Pt(1)
        render_rich_text(p_p3, "*Phần III (Trả lời ngắn):* Câu 1: *6* (hoặc 15)  |  Câu 2: *4*  |  Câu 3: *14*  |  Câu 4: *17*", base_size=9.5)

        p_p4 = doc.add_paragraph()
        p_p4.paragraph_format.space_before = Pt(1)
        p_p4.paragraph_format.space_after = Pt(4)
        render_rich_text(p_p4, "*Phần Tự luận:* Câu 1: *p = e = 13, n = 14; A = 27*  |  Câu 2: *X: 1s^2^2s^2^2p^6^3s^1^; Y: 1s^2^2s^2^2p^6^3s^2^3p^5^*  |  Câu 3: *Mg, Al, Si; tính base: Mg(OH)~2~ > Al(OH)~3~ > H~2~SiO~3~*", base_size=9.5)

    # -------------------------------------------------------------
    # PHẦN I
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="A. TRẮC NGHIỆM KHÁCH QUAN (7,0 điểm)",
        subtitle="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,0 điểm). Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    questions_p1 = [
        {
            "num": 1,
            "prompt": "Câu 1: Trong nguyên tử, hạt mang điện là",
            "opts": [
                "A. neutron và electron.",
                "B. proton và neutron.",
                "*C. electron và proton.*",
                "D. electron, proton và neutron."
            ],
            "sol": "Trong nguyên tử, các hạt mang điện là proton (mang điện tích dương +1) và electron (mang điện tích âm -1). Neutron là hạt không mang điện.\n➔ Chọn đáp án C."
        },
        {
            "num": 2,
            "prompt": "Câu 2: Nguyên tử X có cấu hình electron [Ne]3s^2^3p^5^. Vị trí của X trong bảng hệ thống tuần hoàn là",
            "opts": [
                "*A. Ô 17, chu kì 3, nhóm VIIA.*",
                "B. Ô 15, chu kì 3, nhóm VIIA.",
                "C. Ô 17, chu kì 3, nhóm VA.",
                "D. Ô 15, chu kì 4, nhóm VA."
            ],
            "sol": "Số hiệu nguyên tử Z = 10 + 2 + 5 = 17 ➔ Ô 17.\nCó 3 lớp electron ➔ Chu kì 3.\nCó 7 electron hóa trị (3s^2^3p^5^), electron cuối cùng điền vào phân lớp p ➔ Nhóm VIIA (nguyên tố phi kim Clo - Cl).\n➔ Chọn đáp án A."
        },
        {
            "num": 3,
            "prompt": "Câu 3: Khi hình thành liên kết hóa học, nguyên tử có số hiệu nào sau đây có xu hướng nhường 2 electron để đạt tới cấu hình electron bền vững theo quy tắc octet?",
            "opts": [
                "*A. Magnesium (Z = 12).*",
                "B. Sodium (Z = 11).",
                "C. Neon (Z = 10).",
                "D. Fluorine (Z = 9)."
            ],
            "sol": "- Magnesium (Z = 12): 1s^2^2s^2^2p^6^3s^2^ (có 2 electron ở lớp ngoài cùng) ➔ Có xu hướng nhường 2 electron để đạt cấu hình octet bền vững của khí hiếm neon ([Ne]).\n- Sodium (Z = 11) nhường 1 electron; Neon (Z = 10) là khí hiếm bền vững; Fluorine (Z = 9) nhận 1 electron.\n➔ Chọn đáp án A."
        },
        {
            "num": 4,
            "prompt": "Câu 4: Dãy các nguyên tố nào sau đây có tính kim loại tăng dần?",
            "opts": [
                "A. Ca < K < Al < Mg.",
                "B. Al < Mg < K < Ca.",
                "C. K < Mg < Al < Ca.",
                "*D. Al < Mg < Ca < K.*"
            ],
            "sol": "- Trong chu kì 3, theo chiều tăng điện tích hạt nhân, tính kim loại giảm: Al (Z=13) < Mg (Z=12).\n- Trong nhóm IIA, theo chiều tăng điện tích hạt nhân, tính kim loại tăng: Mg (Z=12) < Ca (Z=20).\n- Trong chu kì 4, theo chiều tăng điện tích hạt nhân, tính kim loại giảm: Ca (Z=20) < K (Z=19).\n➔ Thứ tự tính kim loại tăng dần: Al < Mg < Ca < K.\n➔ Chọn đáp án D."
        },
        {
            "num": 5,
            "prompt": "Câu 5: Công thức nào sau đây là công thức Lewis?",
            "opts": [
                "A. H-N-H (kèm liên kết đơn với H phía dưới).",
                "B. O=C=O.",
                "*C. :O=O: (có các cặp chấm electron hóa trị tự do quanh O).*",
                "D. H:Cl: (toàn bộ là electron, công thức electron)."
            ],
            "sol": "Công thức Lewis biểu diễn liên kết bằng các gạch nối giữa các nguyên tử và biểu diễn đầy đủ các electron hóa trị chưa tham gia liên kết bằng các dấu chấm bao quanh kí hiệu nguyên tố. Ở phương án C, phân tử O~2~ có liên kết đôi O=O và mỗi nguyên tử O còn lại 2 cặp electron tự do (4 chấm) ➔ Đây chính là công thức Lewis.\n➔ Chọn đáp án C."
        },
        {
            "num": 6,
            "prompt": "Câu 6: X, Y là 2 nguyên tố kế tiếp nhau, thuộc cùng một chu kì trong bảng tuần hoàn các nguyên tố hoá học. Tổng số hiệu nguyên tử của X, Y là 27. X, Y lần lượt thuộc nhóm",
            "opts": [
                "A. IIA và IIIA.",
                "B. IVA và VA.",
                "C. VA và VIA.",
                "*D. IIIA và IVA.*"
            ],
            "sol": "Vì X và Y kế tiếp nhau trong cùng một chu kì nên Z~Y~ - Z~X~ = 1.\nTheo đề bài: Z~X~ + Z~Y~ = 27.\nGiải hệ phương trình ta được: Z~X~ = 13 (Al) và Z~Y~ = 14 (Si).\n- Al (Z = 13): [Ne]3s^2^3p^1^ ➔ Thuộc nhóm IIIA.\n- Si (Z = 14): [Ne]3s^2^3p^2^ ➔ Thuộc nhóm IVA.\n➔ Chọn đáp án D."
        },
        {
            "num": 7,
            "prompt": "Câu 7: Nguyên tử X có 3 lớp electron, lớp ngoài cùng có 2 electron. Phát biểu nào sau đây không đúng?",
            "opts": [
                "*A. X là nguyên tố p.*",
                "B. X là nguyên tố kim loại.",
                "C. Ở trạng thái cơ bản, các phân lớp electron của X đã bão hòa.",
                "D. Điện tích hạt nhân của nguyên tử X là +12."
            ],
            "sol": "Nguyên tử X có 3 lớp electron và lớp ngoài cùng có 2 electron ➔ Cấu hình electron: 1s^2^2s^2^2p^6^3s^2^.\n- Electron cuối cùng điền vào phân lớp s ➔ X là nguyên tố s (không phải nguyên tố p). Phát biểu A sai.\n- X có 2 electron lớp ngoài cùng nên là kim loại (Mg). Điện tích hạt nhân là +12. Các phân lớp 1s, 2s, 2p, 3s đều đã bão hòa.\n➔ Chọn đáp án A."
        },
        {
            "num": 8,
            "prompt": "Câu 8: X là nguyên tố rất cần thiết cho sự chuyển hóa của calcium, phosphorus, sodium, potassium, vitamin C và các vitamin nhóm B. Ở trạng thái cơ bản, cấu hình electron lớp ngoài cùng của nguyên tử X là 3s^2^. Số hiệu nguyên tử của nguyên tố X là",
            "opts": [
                "A. 13.",
                "*B. 12.*",
                "C. 11.",
                "D. 14."
            ],
            "sol": "Ở trạng thái cơ bản, cấu hình electron đầy đủ của X là 1s^2^2s^2^2p^6^3s^2^.\nTổng số electron = 2 + 2 + 6 + 2 = 12 ➔ Số hiệu nguyên tử Z = 12 (Magnesium).\n➔ Chọn đáp án B."
        },
        {
            "num": 9,
            "prompt": "Câu 9: Nguyên tử potassium có 19 proton, 19 electron và 20 neutron. Số khối của nguyên tử potassium là",
            "opts": [
                "A. 19.",
                "*B. 39.*",
                "C. 58.",
                "D. 38."
            ],
            "sol": "Số khối A = Z + N = số proton + số neutron = 19 + 20 = 39.\n➔ Chọn đáp án B."
        },
        {
            "num": 10,
            "prompt": "Câu 10: Nguyên tử của nguyên tố M có cấu hình electron là 1s^2^2s^2^2p^4^. Số electron độc thân của M là",
            "opts": [
                "A. 3.",
                "*B. 2.*",
                "C. 1.",
                "D. 0."
            ],
            "sol": "Phân lớp 2p có 3 orbital (p~x~, p~y~, p~z~), chứa 4 electron. Theo quy tắc Hund:\n- 1 orbital chứa 1 cặp electron đã ghép đôi (2 electron).\n- 2 orbital còn lại, mỗi orbital chứa 1 electron độc thân.\n➔ Nguyên tử M có 2 electron độc thân.\n➔ Chọn đáp án B."
        },
        {
            "num": 11,
            "prompt": "Câu 11: Nguyên tử của nguyên tố nào có bán kính lớn nhất trong các nguyên tử sau đây?",
            "opts": [
                "A. S (Z = 16).",
                "B. Al (Z = 13).",
                "*C. K (Z = 19).*",
                "D. P (Z = 15)."
            ],
            "sol": "- Các nguyên tố Al (Z=13), P (Z=15), S (Z=16) đều thuộc chu kì 3 (có 3 lớp electron).\n- Potassium K (Z=19): [Ar]4s^1^ thuộc chu kì 4 (có 4 lớp electron).\nNguyên tử có càng nhiều lớp electron thì bán kính nguyên tử càng lớn ➔ K có bán kính nguyên tử lớn nhất.\n➔ Chọn đáp án C."
        },
        {
            "num": 12,
            "prompt": "Câu 12: Dãy các chất nào dưới đây mà tất cả các phân tử đều có liên kết cộng hoá trị không phân cực?",
            "opts": [
                "A. N~2~, CO~2~, Cl~2~, H~2~.",
                "B. N~2~, Cl~2~, H~2~, HCl.",
                "C. N~2~, HI, Cl~2~, CH~4~.",
                "*D. Cl~2~, O~2~, N~2~, F~2~.*"
            ],
            "sol": "Liên kết cộng hóa trị không phân cực là liên kết hình thành giữa các nguyên tử có hiệu độ âm điện Δχ < 0,4 (thường là giữa 2 nguyên tử của cùng một phi kim, Δχ = 0).\n- Cl~2~, O~2~, N~2~, F~2~ đều là các phân tử đơn chất phi kim gồm 2 nguyên tử cùng loại ➔ Tất cả các liên kết đều là liên kết cộng hóa trị không phân cực.\n(Các chất CO~2~, HCl, HI, CH~4~ ở các phương án A, B, C đều chứa liên kết cộng hóa trị phân cực).\n➔ Chọn đáp án D."
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
        title="PHẦN II. Câu trắc nghiệm đúng sai (2,0 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1 phần II
    p_q1_p2 = doc.add_paragraph()
    p_q1_p2.paragraph_format.space_before = Pt(3)
    p_q1_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q1_p2, "Câu 1: Potassium iodide (KI) được sử dụng như một loại thuốc long đờm, giúp làm lỏng và phá vỡ chất nhầy trong đường thở, thường dùng cho các bệnh nhân hen suyễn, viêm phế quản mãn tính. Trong trường hợp bị nhiễm phóng xạ, KI còn giúp ngăn tuyến giáp hấp thụ iodide phóng xạ, bảo vệ và làm giảm nguy cơ ung thư tuyến giáp.")

    stmts_c1 = [
        ("a) Hợp chất KI được tạo bởi lực hút tĩnh điện giữa hai ion trái dấu là K^+^ và I^-^.", "Đúng"),
        ("b) Trong phân tử KI, các nguyên tử K^+^ và I^-^ đều đã đạt được cấu hình bền của khí hiếm gần nhất.", "Đúng"),
        ("c) Liên kết hóa học trong phân tử KI là liên kết cộng hóa trị phân cực.", "Sai"),
        ("d) Ở điều kiện thường, KI là chất rắn, khó nóng chảy, khó bay hơi.", "Đúng")
    ]
    for stmt, ans in stmts_c1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        sol_c1_p2 = (
            "a) ĐÚNG: KI là hợp chất ion được tạo nên bởi lực hút tĩnh điện giữa cation K^+^ và anion I^-^.\n"
            "b) ĐÚNG: K (Z=19) nhường 1e thành K^+^ đạt cấu hình octet bền vững của Ar ([Ar]); I (Z=53) nhận 1e thành I^-^ đạt cấu hình octet của Xe ([Xe]).\n"
            "c) SAI: Liên kết giữa kim loại điển hình K và phi kim điển hình I là liên kết ion (hiệu độ âm điện Δχ = 2,66 - 0,82 = 1,84 ≥ 1,7).\n"
            "d) ĐÚNG: KI là tinh thể ion có lực hút tĩnh điện mạnh nên ở điều kiện thường là chất rắn, khó nóng chảy, khó bay hơi."
        )
        add_solution_no_table(doc, sol_c1_p2)

    # Câu 2 phần II
    p_q2_p2 = doc.add_paragraph()
    p_q2_p2.paragraph_format.space_before = Pt(4)
    p_q2_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q2_p2, "Câu 2: Nguyên tử X và Y có cấu hình electron lớp ngoài cùng là 3s^x^ và 3p^y^. Biết Y là nguyên tố khí hiếm và x + y = 8.")

    stmts_c2 = [
        ("a) Nguyên tử X là kim loại.", "Đúng"),
        ("b) Ở trạng thái cơ bản, nguyên tử X có 1 electron độc thân.", "Sai"),
        ("c) Ở trạng thái cơ bản, các phân lớp electron của nguyên tử Y đã bão hòa.", "Đúng"),
        ("d) Số electron của X, Y lần lượt là 11 và 18.", "Sai")
    ]
    for stmt, ans in stmts_c2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        sol_c2_p2 = (
            "- Y là khí hiếm có phân lớp ngoài cùng là 3p^y^ ➔ Phân lớp 3p chứa tối đa 6 electron ➔ y = 6 (cấu hình ngoài cùng 3s^2^3p^6^, Y là Argon Ar, Z = 18).\n"
            "- Theo giả thiết: x + y = 8 ➔ x = 8 - 6 = 2. Cấu hình ngoài cùng của X là 3s^2^ (X là Magnesium Mg, Z = 12).\n"
            "Xét các phát biểu:\n"
            "a) ĐÚNG: X là kim loại Mg (nhóm IIA).\n"
            "b) SAI: Phân lớp ngoài cùng 3s^2^ của X đã bão hòa, orbital s chứa cặp electron ghép đôi ➔ X có 0 electron độc thân.\n"
            "c) ĐÚNG: Cấu hình của Y là 1s^2^2s^2^2p^6^3s^2^3p^6^, tất cả các phân lớp 1s, 2s, 2p, 3s, 3p đều đã bão hòa electron.\n"
            "d) SAI: Số electron của X là 12 (Mg) và của Y là 18 (Ar)."
        )
        add_solution_no_table(doc, sol_c2_p2)

    # -------------------------------------------------------------
    # PHẦN III
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm trả lời ngắn (2,0 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Mỗi câu đúng được 0,5 điểm."
    )

    questions_p3 = [
        {
            "num": 1,
            "prompt": "Câu 1: Nguyên tố đồng có 2 đồng vị: ^63^_29_Cu, ^65^_29_Cu; nguyên tố oxygen có 3 đồng vị: ^16^_8_O, ^17^_8_O và ^18^_8_O. Có bao nhiêu phân tử oxide được tạo thành từ các đồng vị trên?",
            "ans": "6 (hoặc 15)",
            "sol": "- Nếu xét oxide dạng CuO (1 nguyên tử Cu và 1 nguyên tử O):\n  Số loại phân tử CuO = 2 (đồng vị Cu) × 3 (đồng vị O) = 6 phân tử.\n- Nếu xét oxide dạng Cu~2~O (2 nguyên tử Cu và 1 nguyên tử O):\n  Số cặp đồng vị Cu có thể tạo thành: (^63^Cu, ^63^Cu), (^65^Cu, ^65^Cu), (^63^Cu, ^65^Cu) ➔ 3 kiểu cặp.\n  Kết hợp với 3 đồng vị của O ➔ 3 × 3 = 9 loại phân tử Cu~2~O.\n- Tổng số phân tử oxide của đồng tạo thành từ các đồng vị là: 6 + 9 = 15 phân tử (hoặc 6 phân tử nếu chỉ tính CuO phổ biến).\n➔ Đáp án: 6 hoặc 15."
        },
        {
            "num": 2,
            "prompt": "Câu 2: Trong các chất sau, có bao nhiêu chất có liên kết ion trong phân tử: AlCl~3~; Al~2~O~3~; AlBr~3~; NaCl; NaBr; Na~2~O? Biết giá trị độ âm điện: Na (0,93), O (3,44), Al (1,61), Cl (3,16), Br (2,96).",
            "ans": "4",
            "sol": "Hiệu độ âm điện Δχ của liên kết trong từng phân tử:\n- AlCl~3~: Δχ = 3,16 - 1,61 = 1,55 < 1,7 ➔ Liên kết cộng hóa trị phân cực.\n- Al~2~O~3~: Δχ = 3,44 - 1,61 = 1,83 ≥ 1,7 ➔ Liên kết ion.\n- AlBr~3~: Δχ = 2,96 - 1,61 = 1,35 < 1,7 ➔ Liên kết cộng hóa trị phân cực.\n- NaCl: Δχ = 3,16 - 0,93 = 2,23 ≥ 1,7 ➔ Liên kết ion.\n- NaBr: Δχ = 2,96 - 0,93 = 2,03 ≥ 1,7 ➔ Liên kết ion.\n- Na~2~O: Δχ = 3,44 - 0,93 = 2,51 ≥ 1,7 ➔ Liên kết ion.\n➔ Có 4 chất có liên kết ion: Al~2~O~3~, NaCl, NaBr, Na~2~O.\n➔ Đáp án: 4."
        },
        {
            "num": 3,
            "prompt": "Câu 3: Nguyên tố X thuộc chu kì 2, nhóm VA. Tổng số hạt mang điện có trong nguyên tử X là bao nhiêu?",
            "ans": "14",
            "sol": "- X ở chu kì 2 (2 lớp electron), nhóm VA (5 electron hóa trị lớp ngoài cùng) ➔ Cấu hình electron: 1s^2^2s^2^2p^3^.\n- Số hiệu nguyên tử Z = 2 + 2 + 3 = 7 (Nitrogen, N).\n- Hạt mang điện trong nguyên tử gồm proton (p) và electron (e), với p = e = Z = 7.\n- Tổng số hạt mang điện = p + e = 2Z = 2 × 7 = 14 hạt.\n➔ Đáp án: 14."
        },
        {
            "num": 4,
            "prompt": "Câu 4: Các electron của nguyên tử nguyên tố X được phân bố trên 3 lớp. Lớp thứ 3 có 7 electron. X nằm ở ô thứ mấy trong bảng tuần hoàn?",
            "ans": "17",
            "sol": "- Nguyên tử có 3 lớp electron: Lớp 1 có 2 electron (1s^2^), Lớp 2 có 8 electron (2s^2^2p^6^), Lớp 3 có 7 electron (3s^2^3p^5^).\n- Cấu hình electron đầy đủ: 1s^2^2s^2^2p^6^3s^2^3p^5^.\n- Tổng số electron = 2 + 8 + 7 = 17 ➔ Số hiệu nguyên tử Z = 17 ➔ X nằm ở ô số 17 (nguyên tố Clo - Cl).\n➔ Đáp án: 17."
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
        subtitle="Thí sinh trình bày lời giải chi tiết cho các câu hỏi sau:"
    )

    # Câu 1 Tự luận
    p_tl1 = doc.add_paragraph()
    p_tl1.paragraph_format.space_before = Pt(3)
    p_tl1.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_tl1, "Câu 1 (1,0 điểm): Hợp kim chứa nguyên tố X nhẹ và bền, dùng chế tạo vỏ máy bay, tên lửa. Nguyên tố X còn được sử dụng trong xây dựng, ngành điện và đồ gia dụng. Nguyên tử của nguyên tố X có tổng số hạt (proton, electron, neutron) là 40. Tổng số hạt mang điện nhiều hơn tổng số hạt không mang điện là 12.")

    p_tl1_sub = doc.add_paragraph()
    p_tl1_sub.paragraph_format.left_indent = Cm(0.5)
    p_tl1_sub.paragraph_format.space_before = Pt(1)
    p_tl1_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_tl1_sub, "a) Tính số hạt mỗi loại (proton, electron, neutron) trong nguyên tử X?\nb) Tính số khối của nguyên tử X?")

    if is_teacher:
        sol_tl1 = (
            "Gọi số proton, electron, neutron trong nguyên tử X lần lượt là p, e, n (p = e = Z).\n"
            "Theo bài ra, ta có hệ phương trình:\n"
            "  { 2Z + n = 40 (1)\n"
            "  { 2Z - n = 12 (2)\n"
            "Cộng (1) và (2) vế theo vế: 4Z = 52 ➔ Z = 13.\n"
            "Thay Z = 13 vào (1) ➔ n = 40 - 2 × 13 = 14.\n"
            "a) Số hạt mỗi loại trong nguyên tử X:\n"
            "  - Số proton = số electron = p = e = 13.\n"
            "  - Số neutron = n = 14.\n"
            "b) Số khối của nguyên tử X:\n"
            "  A = Z + n = 13 + 14 = 27 (X là nguyên tố nhôm - Al)."
        )
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Câu 2 Tự luận
    p_tl2 = doc.add_paragraph()
    p_tl2.paragraph_format.space_before = Pt(4)
    p_tl2.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_tl2, "Câu 2 (1,0 điểm): Nguyên tử X và Y có cấu hình electron phân lớp ngoài cùng lần lượt là 3s^x^ và 3p^y^. Biết phân lớp 3s của hai nguyên tử hơn kém nhau một electron và x + y = 6.")

    p_tl2_sub = doc.add_paragraph()
    p_tl2_sub.paragraph_format.left_indent = Cm(0.5)
    p_tl2_sub.paragraph_format.space_before = Pt(1)
    p_tl2_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_tl2_sub, "a) Viết cấu hình electron của X và Y?\nb) Cho biết để được cấu hình bền của khí hiếm gần nó nhất thì X, Y phải nhường hoặc nhận bao nhiêu electron. Viết cấu hình electron của ion tương ứng?")

    if is_teacher:
        sol_tl2 = (
            "a) Xác định cấu hình electron của X và Y:\n"
            "- Y có phân lớp ngoài cùng là 3p^y^ (1 ≤ y ≤ 6) ➔ Trước khi điền vào 3p, phân lớp 3s của Y bắt buộc đã bão hòa 2 electron (3s^2^).\n"
            "- Phân lớp 3s của hai nguyên tử hơn kém nhau 1 electron ➔ Phân lớp 3s của X có số electron là: x = 2 - 1 = 1.\n"
            "- Lại có: x + y = 6 ➔ y = 6 - 1 = 5.\n"
            "Vậy cấu hình electron đầy đủ:\n"
            "  + Nguyên tử X (x = 1, Z = 11): 1s^2^2s^2^2p^6^3s^1^ (hoặc [Ne]3s^1^).\n"
            "  + Nguyên tử Y (y = 5, Z = 17): 1s^2^2s^2^2p^6^3s^2^3p^5^ (hoặc [Ne]3s^2^3p^5^).\n"
            "b) Sự hình thành ion và cấu hình electron của ion:\n"
            "- Nguyên tử X (Na) có 1 electron ở lớp ngoài cùng ➔ Có xu hướng nhường đi 1 electron để đạt cấu hình electron bền vững của khí hiếm Ne:\n"
            "  Na ➔ Na^+^ + 1e\n"
            "  Cấu hình electron của ion X^+^ (Na^+^): 1s^2^2s^2^2p^6^ (hoặc [Ne]).\n"
            "- Nguyên tử Y (Cl) có 7 electron ở lớp ngoài cùng ➔ Có xu hướng nhận thêm 1 electron để đạt cấu hình electron bền vững của khí hiếm Ar:\n"
            "  Cl + 1e ➔ Cl^-^\n"
            "  Cấu hình electron của ion Y^-^ (Cl^-^): 1s^2^2s^2^2p^6^3s^2^3p^6^ (hoặc [Ar])."
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 3 Tự luận
    p_tl3 = doc.add_paragraph()
    p_tl3.paragraph_format.space_before = Pt(4)
    p_tl3.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_tl3, "Câu 3 (1,0 điểm): Ba nguyên tố X, Y, Z thuộc cùng một chu kì và có tổng số hiệu nguyên tử là 39. Số hiệu nguyên tử Y bằng trung bình cộng số hiệu của nguyên tử X và Z. Nguyên tử của ba nguyên tố trên hầu như không phản ứng với H~2~O ở điều kiện thường.")

    p_tl3_sub = doc.add_paragraph()
    p_tl3_sub.paragraph_format.left_indent = Cm(0.5)
    p_tl3_sub.paragraph_format.space_before = Pt(1)
    p_tl3_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_tl3_sub, "a/ X, Y, Z là những nguyên tố nào? Hãy xác định vị trí của X, Y, Z trong bảng tuần hoàn?\nb/ Viết công thức hydroxide của X, Y, Z và so sánh tính base của chúng?")

    if is_teacher:
        sol_tl3 = (
            "a/ Xác định X, Y, Z và vị trí trong bảng tuần hoàn:\n"
            "- Tổng số hiệu nguyên tử: Z~X~ + Z~Y~ + Z~Z~ = 39 (1)\n"
            "- Số hiệu của Y bằng trung bình cộng số hiệu của X và Z:\n"
            "  Z~Y~ = (Z~X~ + Z~Z~) / 2 ➔ Z~X~ + Z~Z~ = 2 Z~Y~ (2)\n"
            "Thay (2) vào (1): 3 Z~Y~ = 39 ➔ Z~Y~ = 13 (Aluminium - Al, chu kì 3).\n"
            "- Vì X, Y, Z thuộc cùng chu kì 3 ➔ 11 ≤ Z~X~, Z~Z~ ≤ 18.\n"
            "  Lại có Z~X~ + Z~Z~ = 26. Giả sử Z~X~ < 13 < Z~Z~:\n"
            "  Nếu Z~X~ = 11 (Na) ➔ Z~Z~ = 15 (P). Nhưng Na phản ứng mãnh liệt với H~2~O ở điều kiện thường (trái với giả thiết: cả 3 nguyên tố hầu như không phản ứng với H~2~O) ➔ Loại.\n"
            "  Vậy Z~X~ = 12 (Mg) ➔ Z~Z~ = 26 - 12 = 14 (Si). Mg, Al và Si đều hầu như không tác dụng với nước ở nhiệt độ thường (thỏa mãn giả thiết).\n"
            "*Vị trí trong bảng tuần hoàn:*\n"
            "  + X là Magnesium (Mg, Z = 12): 1s^2^2s^2^2p^6^3s^2^ ➔ Ô số 12, chu kì 3, nhóm IIA.\n"
            "  + Y là Aluminium (Al, Z = 13): 1s^2^2s^2^2p^6^3s^2^3p^1^ ➔ Ô số 13, chu kì 3, nhóm IIIA.\n"
            "  + Z là Silicon (Si, Z = 14): 1s^2^2s^2^2p^6^3s^2^3p^2^ ➔ Ô số 14, chu kì 3, nhóm IVA.\n"
            "b/ Công thức hydroxide và so sánh tính base:\n"
            "- Công thức hydroxide tương ứng:\n"
            "  + Của Mg (Z=12): Mg(OH)~2~ (base yếu).\n"
            "  + Của Al (Z=13): Al(OH)~3~ (hydroxide lưỡng tính).\n"
            "  + Của Si (Z=14): H~2~SiO~3~ (silicic acid - acid rất yếu).\n"
            "- So sánh tính base: Trong cùng chu kì 3, theo chiều tăng điện tích hạt nhân (Mg ➔ Al ➔ Si), tính base của hydroxide giảm dần, tính acid tăng dần:\n"
            "  Mg(OH)~2~ > Al(OH)~3~ > H~2~SiO~3~\n"
            "  (Mg(OH)~2~ là base yếu > Al(OH)~3~ là hydroxide lưỡng tính > H~2~SiO~3~ là acid yếu)."
        )
        add_solution_no_table(doc, sol_tl3)
    else:
        add_student_writing_lines(doc, num_lines=8)

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
    f_hs = os.path.join(out_dir, "De_Kiem_Tra_Hoa_10_THPT_Nguyen_Dinh_Chieu_Ma1002_HocSinh.docx")
    f_gv = os.path.join(out_dir, "De_Kiem_Tra_Hoa_10_THPT_Nguyen_Dinh_Chieu_Ma1002_GiaoVien.docx")
    create_exam_file(is_teacher=False, output_path=f_hs)
    create_exam_file(is_teacher=True, output_path=f_gv)
