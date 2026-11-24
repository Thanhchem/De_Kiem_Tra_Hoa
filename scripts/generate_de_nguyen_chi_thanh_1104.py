# -*- coding: utf-8 -*-
"""
Script tạo 2 file Word chuẩn Form Đề Kiểm Tra Hóa học (Bản Học Sinh & Bản Giáo Viên)
Đề: TRƯỜNG THPT NGUYỄN CHÍ THANH - MÃ ĐỀ 1104 - CUỐI HỌC KÌ I - LỚP 10 (2025-2026)
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
        school_line="SỞ GD & ĐT TP. HUẾ\nTRƯỜNG THPT NGUYỄN CHÍ THANH",
        exam_title="ĐỀ KIỂM TRA CUỐI HỌC KÌ 1 – NĂM HỌC 2025 - 2026",
        subject_line="MÔN: HÓA HỌC – LỚP 10",
        num_pages="03",
        duration="45 Phút",
        exam_code="1104",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 23; Si = 28; P = 31; S = 32; Cl = 35,5; Fe = 56; Cu = 64",
        is_teacher=is_teacher
    )

    # -------------------------------------------------------------
    # BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO BẢN GIÁO VIÊN)
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

        tbl1 = doc.add_table(rows=2, cols=16)
        tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
        ans_p1 = ["1. B", "2. C", "3. A", "4. D", "5. B", "6. C", "7. A", "8. C", 
                  "9. D", "10. A", "11. C", "12. D", "13. B", "14. A", "15. B", "16. C"]
        for col_idx in range(16):
            cell_top = tbl1.cell(0, col_idx)
            cell_top.text = f"C{col_idx+1}"
            cell_top.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run_style(cell_top.paragraphs[0].runs[0], font_size=8.5, bold=True)

            cell_bot = tbl1.cell(1, col_idx)
            cell_bot.text = ans_p1[col_idx].split(". ")[1]
            cell_bot.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run_style(cell_bot.paragraphs[0].runs[0], font_size=8.5, bold=True, color=COLOR_CORRECT)

        # Đáp án Phần II & Phần III & Phần Tự luận
        p_p2 = doc.add_paragraph()
        p_p2.paragraph_format.space_before = Pt(3)
        p_p2.paragraph_format.space_after = Pt(1)
        render_rich_text(p_p2, "*Phần II (Đúng/Sai):* Câu 1: *a-S, b-S, c-Đ, d-Đ*  |  Câu 2: *a-Đ, b-S, c-Đ, d-S*", base_size=9.5)

        p_p3 = doc.add_paragraph()
        p_p3.paragraph_format.space_before = Pt(1)
        p_p3.paragraph_format.space_after = Pt(1)
        render_rich_text(p_p3, "*Phần III (Trả lời ngắn):* Câu 1: *3*  |  Câu 2: *3*  |  Câu 3: *98* (hoặc 50)  |  Câu 4: *9*", base_size=9.5)

        p_p4 = doc.add_paragraph()
        p_p4.paragraph_format.space_before = Pt(1)
        p_p4.paragraph_format.space_after = Pt(4)
        render_rich_text(p_p4, "*Phần Tự luận:* C1: *S (Ô 16, CK 3, nhóm VIA); H-S-H*  |  C2: *P (M=31); %m(P trong P~2~O~5~) = 43,7%*  |  C3: *%^63^Cu = 73%, %^65^Cu = 27%; %m(^65^Cu trong Cu~2~S) = 22,06%*", base_size=9.5)

    # -------------------------------------------------------------
    # PHẦN I
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="A. PHẦN TRẮC NGHIỆM ( 7 điểm)",
        subtitle="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn ( 16 câu – 4 điểm). Thí sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    questions_p1 = [
        {
            "num": 1,
            "prompt": "Câu 1. Có các phát biểu sau\n(a). Tổng số electron tối đa ở lớp M (n = 3) là 18.\n(b). Các nguyên tố trong cùng một nhóm có tính chất hóa học tương tự nhau.\n(c). Trong bảng tuần hoàn các nguyên tố hóa học, oxygen là nguyên tố có độ âm điện lớn nhất.\n(d). Liên kết trong phân tử N~2~ được tạo thành bằng hai cặp electron dùng chung.\n(e). Nhiệt độ sôi của CH~3~-CH~2~-OH ( ethyl alcohol) và CH~3~-O-CH~3~ ( dimethyl ether) bằng nhau.\nSố phát biểu đúng là",
            "opts": [
                "A. 4.",
                "*B. 2.*",
                "C. 5",
                "D. 3."
            ],
            "sol": "- (a) ĐÚNG: Lớp M (n = 3) chứa tối đa 2n^2^ = 2 × 3^2^ = 18 electron.\n- (b) ĐÚNG: Các nguyên tố trong cùng một nhóm A có cấu hình electron lớp ngoài cùng tương tự nhau nên có tính chất hóa học tương tự nhau.\n- (c) SAI: Fluorine là nguyên tố có độ âm điện lớn nhất (χ ≈ 3,98), oxygen đứng thứ hai (χ ≈ 3,44).\n- (d) SAI: Phân tử N~2~ có liên kết ba (N≡N) được tạo bởi 3 cặp electron dùng chung.\n- (e) SAI: CH~3~CH~2~OH tạo được liên kết hydrogen liên phân tử nên có nhiệt độ sôi cao hơn nhiều so với CH~3~OCH~3~.\n➔ Có 2 phát biểu đúng: (a) và (b).\n➔ Chọn đáp án B."
        },
        {
            "num": 2,
            "prompt": "Câu 2. Tổng số electron tối đa trong phân lớp p là",
            "opts": [
                "A. 2.",
                "B. 10.",
                "*C. 6.*",
                "D. 14."
            ],
            "sol": "Phân lớp p gồm 3 orbital (p~x~, p~y~, p~z~), mỗi orbital chứa tối đa 2 electron ➔ Phân lớp p chứa tối đa 6 electron.\n➔ Chọn đáp án C."
        },
        {
            "num": 3,
            "prompt": "Câu 3. Quá trình tạo thành ion từ nguyên tử Al nào sau đây là đúng ?",
            "opts": [
                "*A. Al ➔ Al^3+^ + 3e.*",
                "B. Al+ 3e ➔ Al^3+^.",
                "C. Al ➔ Al^2+^ + 2e.",
                "D. Al^3+^ + 3e ➔ Al."
            ],
            "sol": "Al (Z = 13) có 3 electron ở lớp ngoài cùng (3s^2^3p^1^), có xu hướng nhường 3 electron để đạt cấu hình octet bền vững của khí hiếm neon: Al ➔ Al^3+^ + 3e.\n➔ Chọn đáp án A."
        },
        {
            "num": 4,
            "prompt": "Câu 4. Chromium được sử dụng nhiều trong luyện kim để chế tạo hợp kim chống ăn mòn và đánh bóng bề mặt. Nguyên tử chromium có cấu hình electron viết gọn là [Ar]3d^5^4s^1^. Vị trí chromium trong bảng tuần hoàn là",
            "opts": [
                "A. ô số 24, chu kì 3, nhóm VB.",
                "B. ô số 17, chu kì 4, nhóm IA.",
                "C. ô số 27, chu kì 4, nhóm IB.",
                "*D. ô số 24, chu kì 4, nhóm VIB.*"
            ],
            "sol": "Chromium có cấu hình: [Ar]3d^5^4s^1^ ➔ Z = 18 + 5 + 1 = 24 ➔ Ô số 24.\nCó 4 lớp electron ➔ Chu kì 4.\nTổng electron phân lớp 3d và 4s là 5 + 1 = 6, nguyên tố d ➔ Nhóm VIB.\n➔ Chọn đáp án D."
        },
        {
            "num": 5,
            "prompt": "Câu 5. Trong bảng tuần hoàn các nguyên tố hóa học, các nguyên tố được sắp xếp theo ba nguyên tắc, nguyên tắc nào sau đây là đúng?",
            "opts": [
                "A. Cùng số electron hóa trị xếp cùng hàng.",
                "*B. Điện tích hạt nhân tăng dần.*",
                "C. Nguyên tử khối tăng dần.",
                "D. Cùng số lớp electron xếp cùng một cột."
            ],
            "sol": "Ba nguyên tắc sắp xếp các nguyên tố trong bảng tuần hoàn:\n1. Các nguyên tố được sắp xếp theo chiều tăng dần của điện tích hạt nhân nguyên tử.\n2. Các nguyên tố có cùng số lớp electron trong nguyên tử được xếp thành một hàng (chu kì).\n3. Các nguyên tố có cùng số electron hóa trị được xếp thành một cột (nhóm).\n➔ Phát biểu B là đúng.\n➔ Chọn đáp án B."
        },
        {
            "num": 6,
            "prompt": "Câu 6. Số đồng vị bền của nguyên tố neon (Ne) được suy ra từ phổ khối lượng ở hình bên là:",
            "opts": [
                "A. 1.",
                "B. 4.",
                "*C. 3.*",
                "D. 2."
            ],
            "sol": "Trên phổ khối lượng của Neon xuất hiện 3 pic tín hiệu tương ứng với các giá trị m/z bằng 20 (90,9%), 21 (0,3%) và 22 (8,8%). Điều này chứng tỏ nguyên tố Neon có 3 đồng vị bền trong tự nhiên là ^20^Ne, ^21^Ne và ^22^Ne.\n➔ Chọn đáp án C."
        },
        {
            "num": 7,
            "prompt": "Câu 7. Liên kết trong phân tử chất X được hình thành nhờ sự xen phủ orbital trục giữa orbital s (hình cầu) và orbital p (hình số 8 nổi). Chất X nào sau đây không thỏa mãn với sự xen phủ trên ?",
            "opts": [
                "*A. Cl~2~.*",
                "B. HBr.",
                "C. HF.",
                "D. HCl."
            ],
            "sol": "Hình vẽ mô tả sự xen phủ trục s - p (giữa orbital 1s của H và orbital p của halogen).\n- Các phân tử HF, HCl, HBr đều có liên kết hình thành do sự xen phủ giữa AO 1s của H và AO p của halogen.\n- Phân tử Cl~2~ có liên kết hình thành do sự xen phủ trục giữa 2 orbital p của 2 nguyên tử Cl (xen phủ p - p).\n➔ Cl~2~ không thỏa mãn xen phủ s - p.\n➔ Chọn đáp án A."
        },
        {
            "num": 8,
            "prompt": "Câu 8. Fluorine và hợp chất của nó được sử dụng làm chất chống sâu răng, chất cách điện, chất làm lạnh, vật liệu chống dính,... Nguyên tử fluorine chứa 9 hạt electron và 10 hạt neutron. Tổng số hạt proton, electron và neutron trong nguyên tử fluorine là",
            "opts": [
                "A. 32.",
                "B. 19.",
                "*C. 28.*",
                "D. 30."
            ],
            "sol": "Trong nguyên tử trung hòa về điện: số proton = số electron = 9.\nSố neutron = 10.\nTổng số hạt = p + e + n = 9 + 9 + 10 = 28 hạt.\n➔ Chọn đáp án C."
        },
        {
            "num": 9,
            "prompt": "Câu 9. Nguyên tố hoá học X thuộc chu kỳ 3 nhóm VA. Cấu hình electron của nguyên tử X là",
            "opts": [
                "A. 1s^2^2s^2^2p^6^3s^2^3p^4^.",
                "B. 1s^2^2s^2^2p^6^3s^2^3p^2^.",
                "C. 1s^2^2s^2^2p^6^3s^2^3p^5^.",
                "*D. 1s^2^2s^2^2p^6^3s^2^3p^3^.*"
            ],
            "sol": "X thuộc chu kì 3 (có 3 lớp electron) và nhóm VA (có 5 electron lớp ngoài cùng dạng 3s^2^3p^3^).\nCấu hình electron đầy đủ của X là: 1s^2^2s^2^2p^6^3s^2^3p^3^ (Phosphorus - P, Z = 15).\n➔ Chọn đáp án D."
        },
        {
            "num": 10,
            "prompt": "Câu 10. Trong một chu kì, theo chiều tăng dần của điện tích hạt nhân",
            "opts": [
                "*A. bán kính nguyên tử giảm dần, tính phi kim tăng dần.*",
                "B. bán kính nguyên tử tăng dần, tính phi kim giảm dần.",
                "C. bán kính nguyên tử tăng dần, tính phi kim tăng dần.",
                "D. bán kính nguyên tử giảm dần, tính kim loại tăng dần."
            ],
            "sol": "Trong một chu kì, theo chiều tăng của điện tích hạt nhân (từ trái qua phải): lực hút giữa hạt nhân và electron lớp ngoài cùng tăng làm bán kính nguyên tử giảm dần, khả năng nhận electron tăng nên tính phi kim tăng dần (tính kim loại giảm dần).\n➔ Chọn đáp án A."
        },
        {
            "num": 11,
            "prompt": "Câu 11. Dãy các nguyên tố được sắp xếp theo chiều tăng dần của bán kính nguyên tử là",
            "opts": [
                "A. N < O < P.",
                "B. P < O < N.",
                "*C. O < N < P.*",
                "D. P < N < O."
            ],
            "sol": "- O (Z = 8) và N (Z = 7) thuộc chu kì 2 (có 2 lớp e): trong chu kì 2, bán kính giảm theo chiều tăng Z ➔ r(O) < r(N).\n- P (Z = 15) thuộc chu kì 3 (có 3 lớp e) ➔ r(P) lớn hơn r(N) và r(O).\n➔ Thứ tự bán kính tăng dần: O < N < P.\n➔ Chọn đáp án C."
        },
        {
            "num": 12,
            "prompt": "Câu 12. Hợp chất nào sau đây tạo được liên kết hydrogen giữa các phân tử?",
            "opts": [
                "A. HCl.",
                "B. CH~4~.",
                "C. H~2~S.",
                "*D. H~2~O.*"
            ],
            "sol": "Điều kiện để tạo liên kết hydrogen liên phân tử là phân tử có nguyên tử H liên kết với các nguyên tử có độ âm điện lớn và bán kính nhỏ như F, O, N và còn cặp electron tự do. Phân tử H~2~O có liên kết O-H phân cực mạnh nên tạo được liên kết hydrogen liên phân tử.\n➔ Chọn đáp án D."
        },
        {
            "num": 13,
            "prompt": "Câu 13. Trong phân tử SO~2~, tổng số các cặp electron hóa trị chưa tham gia liên kết là",
            "opts": [
                "A. 7.",
                "*B. 6.*",
                "C. 5.",
                "D. 4."
            ],
            "sol": "Theo quy tắc bát phân (octet) chuẩn trong SGK:\n- Nguyên tử S ở trung tâm tạo 1 liên kết đôi với 1 nguyên tử O và 1 liên kết cho nhận với nguyên tử O còn lại.\n- Nguyên tử S còn 1 cặp electron tự do (lone pair).\n- Nguyên tử O liên kết đôi còn 2 cặp electron tự do.\n- Nguyên tử O liên kết cho nhận còn 3 cặp electron tự do.\n➔ Tổng số cặp electron chưa tham gia liên kết là: 1 + 2 + 3 = 6 cặp.\n(Ghi chú: Nếu viết theo công thức mở rộng bát phân O=S=O thì có 5 cặp electron tự do).\n➔ Chọn đáp án B."
        },
        {
            "num": 14,
            "prompt": "Câu 14. Một nguyên tử X có tổng số electron ở các phân lớp s là 6 và tổng số electron lớp ngoài cùng là 6. Cho biết X là nguyên tố hoá học nào sau đây?",
            "opts": [
                "*A. S.*",
                "B. Cl",
                "C. F",
                "D. N"
            ],
            "sol": "- Tổng electron ở các phân lớp s là 6 ➔ Gồm 1s^2^, 2s^2^, 3s^2^ ➔ X có 3 lớp electron.\n- Tổng electron ở lớp ngoài cùng là 6 ➔ Phân lớp ngoài cùng là 3s^2^3p^4^.\n➔ Cấu hình đầy đủ: 1s^2^2s^2^2p^6^3s^2^3p^4^ (Z = 16, nguyên tố Sulfur - S).\n➔ Chọn đáp án A."
        },
        {
            "num": 15,
            "prompt": "Câu 15. Phân tử nào sau đây có chứa liên kết cộng hóa trị phân cực ?",
            "opts": [
                "A. NaF",
                "*B. HCl*",
                "C. MgO",
                "D. KCl"
            ],
            "sol": "NaF, MgO, KCl là các hợp chất ion được tạo bởi kim loại điển hình và phi kim điển hình (Δχ ≥ 1,7).\nHCl có hiệu độ âm điện Δχ = 3,16 - 2,20 = 0,96 (0,4 ≤ Δχ < 1,7) ➔ Liên kết cộng hóa trị phân cực.\n➔ Chọn đáp án B."
        },
        {
            "num": 16,
            "prompt": "Câu 16. Hợp chất ion thường có đặc điểm nào sau đây?",
            "opts": [
                "A. Khó tan trong nước.",
                "B. Dẫn điện ở trạng thái rắn.",
                "*C. Nhiệt độ nóng chảy và nhiệt độ sôi cao.*",
                "D. Dễ tan trong dung môi hữu cơ."
            ],
            "sol": "Hợp chất ion có cấu trúc mạng tinh thể ion bền vững nhờ lực hút tĩnh điện mạnh giữa các ion trái dấu nên có nhiệt độ nóng chảy và nhiệt độ sôi cao, thường dễ tan trong nước và không dẫn điện ở trạng thái rắn.\n➔ Chọn đáp án C."
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
        title="PHẦN II. Câu trắc nghiệm đúng sai( 2 câu – 2 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1 phần II
    p_q1_p2 = doc.add_paragraph()
    p_q1_p2.paragraph_format.space_before = Pt(3)
    p_q1_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q1_p2, "Câu 1. Ammonia (NH~3~) có rất nhiều ứng dụng, chủ yếu làm phân đạm (urê, ammonium nitrate) trong nông nghiệp, nguyên liệu sản xuất nitric acid, thuốc nhuộm, dược phẩm....Mô hình phân tử NH~3~ được biểu diễn như sau (mô hình chóp tam giác):")

    stmts_c1 = [
        ("a) Tổng số hạt mang điện trong phân tử NH~3~ là 17.", "Sai"),
        ("b) Phân tử NH~3~ còn 2 cặp electron hóa trị chưa tham gia liên kết.", "Sai"),
        ("c) Liên kết N-H là liên kết cộng hóa trị phân cực.", "Đúng"),
        ("d) Giữa các phân tử NH~3~ có thể tạo liên kết hydrogen liên phân tử.", "Đúng")
    ]
    for stmt, ans in stmts_c1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        sol_c1_p2 = (
            "a) SAI: Phân tử NH~3~ có tổng số hạt mang điện = 2Z~N~ + 3 × 2Z~H~ = 2 × 7 + 3 × 2 = 20 hạt (gồm 10 proton và 10 electron).\n"
            "b) SAI: Nguyên tử N có 5e hóa trị, dùng 3e tạo liên kết với 3 nguyên tử H ➔ còn lại 2e hóa trị, tương ứng với 1 cặp electron tự do (chưa liên kết).\n"
            "c) ĐÚNG: Δχ(N-H) = 3,04 - 2,20 = 0,84 (0,4 ≤ Δχ < 1,7) ➔ Liên kết cộng hóa trị phân cực (cặp e dùng chung lệch về phía N).\n"
            "d) ĐÚNG: Phân tử NH~3~ có nguyên tử H liên kết với N có độ âm điện lớn và còn cặp e tự do ➔ Tạo được liên kết hydrogen liên phân tử."
        )
        add_solution_no_table(doc, sol_c1_p2)

    # Câu 2 phần II
    p_q2_p2 = doc.add_paragraph()
    p_q2_p2.paragraph_format.space_before = Pt(4)
    p_q2_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q2_p2, "Câu 2. Tổng số proton trong nguyên tử của ba nguyên tố X, Y, Z kế tiếp nhau trong một chu kì bằng 36 ( Z~X~ < Z~Y~< Z~Z~).")

    stmts_c2 = [
        ("a) Hydroxide của X có tính base mạnh.", "Đúng"),
        ("b) Z ở chu kỳ 3 nhóm IIA", "Sai"),
        ("c) X có tính kim loại mạnh nhất", "Đúng"),
        ("d) Công thức oxide cao nhất của Y có dạng Y~2~O~3~", "Sai")
    ]
    for stmt, ans in stmts_c2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        sol_c2_p2 = (
            "- Vì X, Y, Z kế tiếp nhau trong một chu kì nên: Z~Y~ = Z~X~ + 1 và Z~Z~ = Z~X~ + 2.\n"
            "- Tổng số proton: Z~X~ + (Z~X~ + 1) + (Z~X~ + 2) = 36 ➔ 3 Z~X~ = 33 ➔ Z~X~ = 11 (Na), Z~Y~ = 12 (Mg), Z~Z~ = 13 (Al).\n"
            "Xét các phát biểu:\n"
            "a) ĐÚNG: Hydroxide của X (Na) là NaOH, là một base mạnh (kiềm mạnh).\n"
            "b) SAI: Z là Al (Z = 13): [Ne]3s^2^3p^1^ thuộc chu kì 3, nhóm IIIA (không phải nhóm IIA).\n"
            "c) ĐÚNG: Trong chu kì 3, theo chiều tăng Z (Na ➔ Mg ➔ Al), tính kim loại giảm dần ➔ Na (X) có tính kim loại mạnh nhất.\n"
            "d) SAI: Y là Mg (nhóm IIA) có hóa trị II nên oxide cao nhất là YO (MgO), không phải Y~2~O~3~."
        )
        add_solution_no_table(doc, sol_c2_p2)

    # -------------------------------------------------------------
    # PHẦN III
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm trả lời ngắn( 4 câu – 1 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Mỗi câu đúng được 0,25 điểm."
    )

    questions_p3 = [
        {
            "num": 1,
            "prompt": "Câu 1. Cho bảng giá trị hiệu độ âm điện của một số nguyên tố như sau: H (2,20); Na (0,93); Al (1,61); O (3,44); Cl (3,16); C (2,55); S (2,58). Trong các phân tử : HCl, Na~2~O, MgO, H~2~O, H~2~S, MgCl~2~, AlCl~3~ có bao nhiêu phân tử có chứa liên kết ion ?",
            "ans": "3",
            "sol": "Hiệu độ âm điện Δχ của liên kết trong từng phân tử:\n- HCl: Δχ = 3,16 - 2,20 = 0,96 < 1,7 ➔ CHT phân cực.\n- Na~2~O: Δχ = 3,44 - 0,93 = 2,51 ≥ 1,7 ➔ Liên kết ion.\n- MgO: Δχ = 3,44 - 1,31 = 2,13 ≥ 1,7 ➔ Liên kết ion.\n- H~2~O: Δχ = 3,44 - 2,20 = 1,24 < 1,7 ➔ CHT phân cực.\n- H~2~S: Δχ = 2,58 - 2,20 = 0,38 < 0,4 ➔ CHT không phân cực.\n- MgCl~2~: Δχ = 3,16 - 1,31 = 1,85 ≥ 1,7 ➔ Liên kết ion.\n- AlCl~3~: Δχ = 3,16 - 1,61 = 1,55 < 1,7 ➔ CHT phân cực.\n➔ Có 3 phân tử có chứa liên kết ion: Na~2~O, MgO, MgCl~2~."
        },
        {
            "num": 2,
            "prompt": "Câu 2. Có cấu hình electron nguyên tử của các nguyên tố như sau :\n(1) 1s^2^2s^2^2p^4^.       (2) 1s^2^2s^2^2p^6^3s^2^3p^1^      (3) 1s^2^2s^2^2p^1^\n(4) 1s^2^2s^2^2p^6^3s^1^     (5) 1s^2^2s^2^2p^6^             (6) 1s^2^2s^2^2p^6^3s^2^3p^6^4s^2^.\nSố cấu hình electron nguyên tử của các nguyên tố kim loại là bao nhiêu ?",
            "ans": "3",
            "sol": "- (1) O: 6e lớp ngoài cùng ➔ Phi kim.\n- (2) Al: 3e lớp ngoài cùng (3s^2^3p^1^) ➔ Kim loại.\n- (3) B: 3e lớp ngoài cùng (2s^2^2p^1^) ➔ Phi kim (á kim).\n- (4) Na: 1e lớp ngoài cùng (3s^1^) ➔ Kim loại.\n- (5) Ne: 8e lớp ngoài cùng ➔ Khí hiếm.\n- (6) Ca: 2e lớp ngoài cùng (4s^2^) ➔ Kim loại.\n➔ Có 3 cấu hình electron của nguyên tố kim loại: (2), (4), (6)."
        },
        {
            "num": 3,
            "prompt": "Câu 3. Tổng số hạt mang điện trong ion SO~4~^2-^ là bao nhiêu ?",
            "ans": "98 (hoặc 50)",
            "sol": "- Số proton trong ion SO~4~^2-^ = Z~S~ + 4 × Z~O~ = 16 + 4 × 8 = 48 hạt.\n- Số electron trong ion SO~4~^2-^ = 48 + 2 = 50 hạt.\n- Hạt mang điện trong ion gồm proton và electron ➔ Tổng số hạt mang điện = 48 + 50 = 98 hạt.\n(Ghi chú: Nếu đề hỏi số electron thì đáp án là 50)."
        },
        {
            "num": 4,
            "prompt": "Câu 4. Copper (đồng) và oxygen có các đồng vị sau: ^65^_29_Cu, ^63^_29_Cu ; ^16^_8_O, ^17^_8_O, ^18^_8_O . Có thể có bao nhiêu loại phân tử copper (I) oxide ( Cu~2~O) khác nhau tạo nên từ các đồng vị của hai nguyên tố đó ?",
            "ans": "9",
            "sol": "Phân tử Cu~2~O gồm 2 nguyên tử Cu và 1 nguyên tử O.\n- Số cặp đồng vị 2 nguyên tử Cu tạo thành: (^63^Cu, ^63^Cu), (^65^Cu, ^65^Cu), (^63^Cu, ^65^Cu) ➔ 3 kiểu cặp.\n- Kết hợp với 3 đồng vị của Oxygen (^16^O, ^17^O, ^18^O) ➔ Số phân tử Cu~2~O = 3 × 3 = 9 loại phân tử."
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
        title="B. PHẦN TỰ LUẬN ( 3 câu – 3 điểm)",
        subtitle="Thí sinh trình bày lời giải chi tiết cho các câu hỏi sau:"
    )

    # Câu 1 Tự luận
    p_tl1 = doc.add_paragraph()
    p_tl1.paragraph_format.space_before = Pt(3)
    p_tl1.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_tl1, "Câu 1. (1,0 điểm) Sulfur (S) có ứng dụng rộng rãi, chủ yếu dùng sản xuất sulfuric acid, phân bón và bột giặt, làm chất lưu hóa cao su, nguyên liệu diêm, thuốc súng, pháo hoa, thuốc diệt nấm, trừ sâu. Trong y tế và mỹ phẩm, nó trị mụn, vảy nến; còn trong nông nghiệp, nó cải thiện cây trồng và bảo quản thực phẩm (sulfite).")

    p_tl1_sub = doc.add_paragraph()
    p_tl1_sub.paragraph_format.left_indent = Cm(0.5)
    p_tl1_sub.paragraph_format.space_before = Pt(1)
    p_tl1_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_tl1_sub, "a. Viết cấu hình electron nguyên tử, xác định vị trí của nguyên tố S trong bảng tuần hoàn ?\nb. Viết công thức Lewis, công thức cấu tạo của phân tử H~2~S ?")

    if is_teacher:
        sol_tl1 = (
            "a. Cấu hình electron và vị trí của S (Z = 16):\n"
            "- Cấu hình electron: 1s^2^2s^2^2p^6^3s^2^3p^4^ (hoặc [Ne]3s^2^3p^4^).\n"
            "- Vị trí trong bảng tuần hoàn:\n"
            "  + Ô số 16 (vì có số hiệu nguyên tử Z = 16).\n"
            "  + Chu kì 3 (vì có 3 lớp electron).\n"
            "  + Nhóm VIA (vì là nguyên tố p, có 6 electron hóa trị ở lớp ngoài cùng 3s^2^3p^4^).\n"
            "b. Công thức Lewis và công thức cấu tạo của phân tử H~2~S:\n"
            "- Nguyên tử S (Z = 16) có 6 electron hóa trị, liên kết đơn với 2 nguyên tử H bằng 2 cặp electron dùng chung; trên S còn lại 2 cặp electron tự do.\n"
            "- Công thức Lewis: H - S̈̈ - H (hoặc H - S̈(dưới)̈(trên) - H).\n"
            "- Công thức cấu tạo: H - S - H."
        )
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 2 Tự luận
    p_tl2 = doc.add_paragraph()
    p_tl2.paragraph_format.space_before = Pt(4)
    p_tl2.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_tl2, "Câu 2. (1,0 điểm) Nguyên tử của nguyên tố X có cấu hình electron lớp ngoài cùng là ns^2^np^3^. Trong hợp chất khí của nguyên tố X với hydgogen, X chiếm 91,18 % về khối lượng. Phần trăm khối lượng của nguyên tố X trong oxide cao nhất là bao nhiêu ? ( làm tròn đến hàng phần mười)")

    if is_teacher:
        sol_tl2 = (
            "- X có cấu hình electron lớp ngoài cùng là ns^2^np^3^ ➔ X có 5 electron lớp ngoài cùng ➔ X thuộc nhóm VA.\n"
            "- Công thức hợp chất khí với hydrogen là XH~3~.\n"
            "- Phần trăm khối lượng của X trong hợp chất khí:\n"
            "  %m_X = [M_X / (M_X + 3 × 1)] × 100% = 91,18%\n"
            "  ➔ M_X / (M_X + 3) = 0,9118 ➔ M_X = 0,9118 M_X + 2,7354 ➔ 0,0882 M_X = 2,7354 ➔ M_X ≈ 31 (g/mol).\n"
            "  Vậy nguyên tố X là Phosphorus (P = 31).\n"
            "- Hóa trị cao nhất của X với oxygen là V (do thuộc nhóm VA) ➔ Công thức oxide cao nhất là X~2~O~5~ (P~2~O~5~).\n"
            "- Phần trăm khối lượng của X trong oxide cao nhất:\n"
            "  %m_X = [2 × 31 / (2 × 31 + 5 × 16)] × 100% = (62 / 142) × 100% ≈ 43,66% ≈ 43,7%."
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 3 Tự luận
    p_tl3 = doc.add_paragraph()
    p_tl3.paragraph_format.space_before = Pt(4)
    p_tl3.paragraph_format.space_after = Pt(1)
    render_question_prompt(p_tl3, "Câu 3. (1,0 điểm) Trong tự nhiên, nguyên tố copper có hai đồng vị là ^63^Cu và ^65^Cu. Nguyên tử khối trung bình của copper là 63,54.")

    p_tl3_sub = doc.add_paragraph()
    p_tl3_sub.paragraph_format.left_indent = Cm(0.5)
    p_tl3_sub.paragraph_format.space_before = Pt(1)
    p_tl3_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_tl3_sub, "a. Tính phần trăm số nguyên tử của mỗi đồng vị ?\nb. Tính thành phần phần trăm về khối của ^65^Cu trong Cu~2~S ?")

    if is_teacher:
        sol_tl3 = (
            "a. Tính phần trăm số nguyên tử của mỗi đồng vị:\n"
            "- Gọi x, y lần lượt là phần trăm số nguyên tử của ^63^Cu và ^65^Cu (x, y > 0).\n"
            "- Ta có hệ phương trình:\n"
            "  { x + y = 100\n"
            "  { (63x + 65y) / 100 = 63,54 ➔ 63x + 65y = 6354\n"
            "  ➔ Giải hệ ta được: x = 73% và y = 27%.\n"
            "  Vậy: %^63^Cu = 73% và %^65^Cu = 27%.\n"
            "b. Tính thành phần phần trăm về khối lượng của ^65^Cu trong Cu~2~S:\n"
            "- Khối lượng mol của Cu~2~S: M(Cu~2~S) = 2 × 63,54 + 32 = 159,08 g/mol.\n"
            "- Trong 1 mol phân tử Cu~2~S có chứa 2 mol nguyên tử Cu.\n"
            "  Số mol nguyên tử đồng vị ^65^Cu trong 1 mol Cu~2~S là: n(^65^Cu) = 2 × 27% = 0,54 mol.\n"
            "- Khối lượng của ^65^Cu trong 1 mol Cu~2~S: m(^65^Cu) = 0,54 × 65 = 35,1 gam.\n"
            "- Phần trăm khối lượng của ^65^Cu trong Cu~2~S:\n"
            "  %m(^65^Cu) = (35,1 / 159,08) × 100% ≈ 22,06%."
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
    f_hs = os.path.join(out_dir, "De_Kiem_Tra_CK1_Hoa_10_NguyenChiThanh_Ma1104_HocSinh.docx")
    f_gv = os.path.join(out_dir, "De_Kiem_Tra_CK1_Hoa_10_NguyenChiThanh_Ma1104_GiaoVien.docx")
    create_exam_file(is_teacher=False, output_path=f_hs)
    create_exam_file(is_teacher=True, output_path=f_gv)
