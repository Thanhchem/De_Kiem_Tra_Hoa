# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file Word cho ĐỀ SỐ 2 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De02.docx (Bản Học Sinh)
- De_Kiem_Tra_Hoa_11_De02_GiaoVien.docx (Bản Giáo Viên)
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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de02")

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
        exam_title="ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 1 (CÂN BẰNG HÓA HỌC)",
        subject_line="MÔN: HOÁ HỌC - LỚP 11",
        num_pages="04",
        duration="45 phút",
        exam_code="102",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Ba = 137",
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
    render_question_prompt(p, "Câu 1. Cân bằng hóa học liên quan đến loại phản ứng")
    opts_c1 = [
        "A. Không thuận nghịch.",
        "*B. Thuận nghịch.*",
        "C. Một chiều.",
        "D. Oxi hóa – khử."
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Cân bằng hóa học là trạng thái của phản ứng thuận nghịch khi tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Yếu tố nào sau đây luôn luôn không làm dịch chuyển cân bằng của hệ phản ứng?")
    opts_c2 = [
        "A. Nhiệt độ.",
        "B. Áp suất.",
        "C. Nồng độ.",
        "*D. Chất xúc tác.*"
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Chất xúc tác làm tăng tốc độ phản ứng thuận và tốc độ phản ứng nghịch với số lần bằng nhau, giúp hệ nhanh đạt trạng thái cân bằng chứ không làm dịch chuyển cân bằng hóa học.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Khi phản ứng thuận nghịch ở trạng thái cân bằng thì nó")
    opts_c3 = [
        "A. Không xảy ra nữa.",
        "*B. Vẫn tiếp tục xảy ra.*",
        "C. Chỉ xảy ra theo chiều thuận.",
        "D. Chỉ xảy ra theo chiều nghịch."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Cân bằng hóa học là cân bằng động, ở trạng thái cân bằng phản ứng thuận và phản ứng nghịch vẫn tiếp tục xảy ra với tốc độ bằng nhau (v~thuận~ = v~nghịch~).")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4 [CTST - SGK]. Yếu tố nào sau đây luôn luôn không làm dịch chuyển cân bằng của hệ phản ứng?")
    opts_c4 = [
        "A. Nhiệt độ",
        "B. Áp suất",
        "C. Nồng độ",
        "*D. Chất xúc tác*"
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Chất xúc tác không làm biến đổi nồng độ các chất ở trạng thái cân bằng, do đó không làm chuyển dịch cân bằng.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Chất nào sau là acid theo thuyết Brønsted – Lowry?")
    opts_c5 = [
        "A. NaNO~3~",
        "*B. HBrO~3~*",
        "C. K~2~SO~4~",
        "D. CsOH"
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo thuyết Brønsted – Lowry, acid là chất có khả năng nhường proton H^+. HBrO~3~ có khả năng phân li nhường H^+ trong dung dịch nước: HBrO~3~ ➔ H^+ + BrO~3~^-.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Phản ứng nào dưới đây không phải là phản ứng acid-base?")
    opts_c6 = [
        "A. HCl + KOH ➔ KCl + H~2~O",
        "B. H~2~SO~4~ + CaO ➔ CaSO~4~ + H~2~O",
        "*C. H~2~SO~4~ + BaCl~2~ ➔ BaSO~4~↓ + 2HCl*",
        "D. 2HNO~3~ + Cu(OH)~2~ ➔ Cu(NO~3~)~2~ + 2H~2~O"
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phản ứng C (H~2~SO~4~ + BaCl~2~ ➔ BaSO~4~↓ + 2HCl) là phản ứng trao đổi ion tạo chất kết tủa, không có sự cho nhận proton H^+ giữa acid và base. Các phản ứng A, B, D đều là phản ứng acid - base.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 7. Các dung dịch sau đây có cùng nồng độ 0,10 mol/l, dung dịch nào dẫn điện kém nhất?")
    opts_c7 = [
        "A. HCl",
        "*B. HF*",
        "C. HI",
        "D. HBr"
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "HCl, HBr, HI là các hydrohalic acid mạnh, điện li hoàn toàn trong nước. HF là acid yếu, điện li một phần nhỏ thành ion nên có mật độ ion tự do nhỏ nhất ➔ dẫn điện kém nhất.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 8. Hòa tan một acid vào nước ở 25°C, kết quả là:")
    opts_c8 = [
        "A. [H^+] < [OH^-]",
        "B. [H^+] = [OH^-]",
        "*C. [H^+] > [OH^-]*",
        "D. [H^+].[OH^-] > 1,0.10^-14^"
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khi hòa tan acid vào nước, [H^+] > 1,0.10^-7^ M > [OH^-]. Tích số ion của nước ở 25°C luôn giữ hằng số: [H^+].[OH^-] = 1,0.10^-14^.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 9. Chất nào sau đây không dẫn điện được?")
    opts_c9 = [
        "*A. KCl rắn, khan*",
        "B. CaCl~2~ nóng chảy",
        "C. NaOH nóng chảy",
        "D. HBr hòa tan trong nước."
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "KCl rắn, khan ở dạng mạng tinh thể ion liên kết bền chặt, các ion K^+ và Cl^- không thể di chuyển tự do nên không dẫn điện. CaCl~2~ nóng chảy, NaOH nóng chảy và dung dịch HBr đều có các ion tự do chuyển động nên dẫn điện.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 10 (SBT - KNTT). Cho phản ứng hoá học sau: N~2~O~4~(g) ⇌ 2NO~2~(g)    K~C~ = 4,84.10^-3^.\nPhương án nào sau đây là nồng độ của các chất tại thời điểm cân bằng?")
    opts_c10 = [
        "A. [N~2~O~4~(g)] = 4,84.10^-1^M; [NO~2~(g)] = 1,0.10^-4^M.",
        "B. [N~2~O~4~(g)] = 1,0.10^-1^M; [NO~2~(g)] = 4,84.10^-4^M.",
        "*C. [N~2~O~4~(g)] = 1,0.10^-1^M; [NO~2~(g)] = 2,20.10^-2^M.*",
        "D. [N~2~O~4~(g)] = 5,0.10^-2^M; [NO~2~(g)] = 1,10.10^-2^M."
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Biểu thức: K~C~ = [NO~2~]^2^ / [N~2~O~4~].\nThay số với phương án C: (2,20.10^-2^)^2^ / (1,0.10^-1^) = 4,84.10^-4^ / 0,1 = 4,84.10^-3^ = K~C~ (thỏa mãn).")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 11. Cho cân bằng hoá học sau: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g); Δ~r~H°~298~ < 0.\nCho các biện pháp: (1) tăng nhiệt độ, (2) tăng áp suất chung của hệ phản ứng, (3) hạ nhiệt độ, (4) dùng thêm chất xúc tác V~2~O~5~, (5) giảm nồng độ SO~3~, (6) giảm áp suất chung của hệ phản ứng. Có mấy biện pháp làm cân bằng chuyển dịch theo chiều thuận?")
    opts_c11 = [
        "A. 2.",
        "*B. 3.*",
        "C. 4.",
        "D. 5."
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phản ứng thuận tỏa nhiệt (Δ~r~H° < 0) và giảm số mol khí (từ 3 xuống 2 mol).\nCác biện pháp làm chuyển dịch theo chiều thuận gồm:\n- (2) Tăng áp suất chung (chuyển dịch về phía giảm số mol khí).\n- (3) Hạ nhiệt độ (chuyển dịch về phía tỏa nhiệt).\n- (5) Giảm nồng độ SO~3~ (chuyển dịch về phía tạo thêm SO~3~).\n➔ Có 3 biện pháp: (2), (3), (5).")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 12. Cho phản ứng: 2NaHCO~3~(s) ⇌ Na~2~CO~3~(s) + CO~2~(g) + H~2~O(g)    Δ~r~H°~298~ = 129 kJ. Phản ứng xảy ra theo chiều nghịch khi:")
    opts_c12 = [
        "*A. Giảm nhiệt độ*",
        "B. Tăng nhiệt độ",
        "C. Giảm áp suất",
        "D. Tăng nhiệt độ và giảm áp suất"
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Δ~r~H°~298~ = 129 kJ > 0 ➔ phản ứng thuận thu nhiệt, phản ứng nghịch tỏa nhiệt. Khi giảm nhiệt độ, cân bằng chuyển dịch theo chiều phản ứng tỏa nhiệt (chiều nghịch).")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 13. Hằng số cân bằng của phản ứng N~2~O~4~(g) ⇌ 2NO~2~(g) là:")
    opts_c13 = [
        "*A. K = [NO~2~]^2^ / [N~2~O~4~]*",
        "B. K = [NO~2~] / [N~2~O~4~]",
        "C. K = [NO~2~] / [N~2~O~4~]^2^",
        "D. Kết quả khác."
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Biểu thức hằng số cân bằng K~C~ bằng tích nồng độ sản phẩm với số mũ bằng hệ số tỉ lượng chia cho tích nồng độ chất tham gia: K~C~ = [NO~2~]^2^ / [N~2~O~4~].")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 14. Xét cân bằng: Fe~2~O~3~(s) + 3CO(g) ⇌ 2Fe(s) + 3CO~2~(g). Biểu thức hằng số cân bằng của hệ là:")
    opts_c14 = [
        "A. K = ([Fe]^2^.[CO~2~]^3^) / ([Fe~2~O~3~].[CO]^3^)",
        "B. K = ([Fe~2~O~3~].[CO]^3^) / ([Fe]^2^.[CO~2~]^3^)",
        "C. K = [CO]^3^ / [CO~2~]^3^",
        "*D. K = [CO~2~]^3^ / [CO]^3^*"
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nồng độ của các chất rắn (Fe~2~O~3~ và Fe) là hằng số và quy ước bằng 1, do đó không xuất hiện trong biểu thức hằng số cân bằng K~C~. Vậy K~C~ = [CO~2~]^3^ / [CO]^3^.")

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 15. Dịch vị dạ dày thường có pH trong khoảng từ 1,5 - 3,5. Những người nào bị mắc bệnh viêm loét dạ dày, tá tràng thường có pH < 1,5. Để chữa căn bệnh này, người bệnh thường uống trước bữa ăn chất nào sau đây?")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau15_ph_he_tieu_hoa.png"), width_cm=6.5)
    opts_c15 = [
        "*A. Dung dịch sodium hydrogen carbonate*",
        "B. Nước đun sôi để nguội",
        "C. Nước đường saccarose",
        "D. Một ít giấm ăn"
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Thuốc muối NaHCO~3~ có tính kiềm yếu, giúp trung hòa bớt acid HCl dư trong dạ dày: NaHCO~3~ + HCl ➔ NaCl + CO~2~ + H~2~O, nâng pH dạ dày về mức an toàn.")

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 16. Để trung hòa 20 ml dung dịch HCl 0,1M cần 10 ml dung dịch NaOH nồng độ x mol/l. Giá trị của x là")
    opts_c16 = [
        "A. 0,3",
        "B. 0,4",
        "*C. 0,2.*",
        "D. 0,1."
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "n(HCl) = n(NaOH) ➔ 0,020 . 0,1 = 0,010 . x ➔ x = 0,2 M.")

    # Câu 17
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 17. Trong dung dịch muối dichromate luôn có cân bằng:\nCr~2~O~7~^2-^ (da cam) + H~2~O ⇌ 2CrO~4~^2-^ (vàng chanh) + 2H^+\nNếu thêm dung dịch acid HBr đặc và dư vào dung dịch K~2~Cr~2~O~7~ thì dung dịch chuyển thành")
    opts_c17 = [
        "*A. Màu da cam.*",
        "B. Màu vàng.",
        "C. Màu xanh lục.",
        "D. Không màu."
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khi thêm acid HBr đặc dư, nồng độ ion H^+ tăng lên, theo nguyên lí Le Chatelier cân bằng chuyển dịch theo chiều giảm nồng độ H^+ (chiều nghịch) ➔ tạo thêm Cr~2~O~7~^2-^ màu da cam.")

    # Câu 18
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 18. Cho các phát biểu sau:\n(1) Phản ứng thuận nghịch là phản ứng xảy ra theo 2 chiều ngược nhau trong cùng điều kiện.\n(2) Chất xúc tác có tác dụng làm tăng tốc độ phản ứng thuận và nghịch.\n(3) Cân bằng hóa học là trạng thái mà phản ứng đã xảy ra hoàn toàn.\n(4) Khi phản ứng thuận nghịch đạt trạng thái cân bằng hóa học, nồng độ các chất sẽ không đổi.\n(5) Khi phản ứng thuận nghịch đạt trạng thái cân bằng hóa học, phản ứng dừng lại.\n(6) Sự chuyển dịch cân bằng của phản ứng thuận nghịch 2NO~2~ ⇌ N~2~O~4~ không phụ thuộc sự thay đổi áp suất.\nSố phát biểu sai là")
    opts_c18 = [
        "A. 1.",
        "B. 2.",
        "*C. 3.*",
        "D. 4."
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Có 3 phát biểu sai gồm: (3), (5), (6).\n- (3) Sai vì cân bằng hóa học không bao giờ xảy ra hoàn toàn.\n- (5) Sai vì cân bằng hóa học là cân bằng động, phản ứng thuận và nghịch vẫn diễn ra.\n- (6) Sai vì phản ứng có biến thiên số mol khí (2 mol ⇌ 1 mol) nên áp suất có ảnh hưởng đến chuyển dịch cân bằng.")

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
    render_question_prompt(p, "Câu 1. Cho cân bằng hoá học trong bình kín: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g), biết phản ứng thuận là phản ứng toả nhiệt.")
    stmts_p2_c1 = [
        ("a) Khi giảm nồng độ O~2~, cân bằng chuyển dịch theo chiều nghịch.", "ĐÚNG"),
        ("b) Khi giảm nhiệt độ cân bằng chuyển dịch theo chiều thuận.", "ĐÚNG"),
        ("c) Khi giảm áp suất, cân bằng chuyển dịch theo chiều thuận.", "SAI"),
        ("d) Khi giảm nồng độ SO~3~, cân bằng sẽ chuyển dịch theo chiều nghịch.", "SAI")
    ]
    for stmt, ans in stmts_p2_c1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c1 = (
            "a) Đúng vì giảm nồng độ O~2~ làm cân bằng chuyển dịch theo chiều tạo thêm O~2~ (chiều nghịch).\n"
            "b) Đúng vì phản ứng thuận tỏa nhiệt, giảm nhiệt độ cân bằng chuyển dịch theo chiều tỏa nhiệt (chiều thuận).\n"
            "c) Sai vì khi giảm áp suất cân bằng chuyển dịch theo chiều tăng số mol khí (từ 2 lên 3 mol, chiều nghịch).\n"
            "d) Sai vì giảm nồng độ SO~3~ cân bằng chuyển dịch theo chiều tạo thêm SO~3~ (chiều thuận)."
        )
        add_solution_no_table(doc, sol_p2_c1)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Cho các cân bằng sau:\n(1) 2NH~3~(g) ⇌ N~2~(g) + 3H~2~(g)    Δ~r~H°~298~ > 0\n(2) 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g)    Δ~r~H°~298~ < 0\n(3) CaCO~3~(r) ⇌ CaO(r) + CO~2~(g)    Δ~r~H°~298~ > 0\n(4) H~2~(g) + I~2~(g) ⇌ 2HI(g)    Δ~r~H°~298~ < 0\nTrong các cân bằng trên cân bằng nào sẽ chuyển dịch theo chiều thuận khi tăng nhiệt độ và giảm áp suất:")
    stmts_p2_c2 = [
        ("a) phương trình 1,2 sẽ chuyển dịch theo chiều thuận khi tăng nhiệt độ và giảm áp suất.", "SAI"),
        ("b) phương trình 1,3 sẽ chuyển dịch theo chiều thuận khi tăng nhiệt độ và giảm áp suất.", "ĐÚNG"),
        ("c) phương trình 1,2,3 cân bằng dịch theo chiều thuận khi giảm áp suất.", "SAI"),
        ("d) phương trình 4 cân bằng không chuyển dịch khi tăng áp suất.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c2 = (
            "Tăng nhiệt độ làm cân bằng chuyển dịch theo chiều thu nhiệt (Δ~r~H° > 0), gồm (1) và (3).\n"
            "Giảm áp suất làm cân bằng chuyển dịch theo chiều tăng số mol khí: (1) tăng từ 2 lên 4 mol; (3) tăng từ 0 lên 1 mol khí.\n"
            "➔ Do đó cả 2 điều kiện làm (1) và (3) chuyển dịch theo chiều thuận. Ý a sai, ý b đúng.\n"
            "c) Sai vì (2) có số mol khí giảm (từ 3 xuống 2) nên giảm áp suất chuyển dịch theo chiều nghịch.\n"
            "d) Đúng vì (4) có số mol khí 2 vế bằng nhau (1+1 = 2) nên áp suất không làm chuyển dịch cân bằng."
        )
        add_solution_no_table(doc, sol_p2_c2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Thực hiện thí nghiệm với 3 chất tan khác nhau là ethanol, potassium chloride và acetic acid như hình dưới đây.")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau3_p2_thi_nghiem_dan_dien.png"), width_cm=11.0)
    stmts_p2_c3 = [
        ("a) Dung dịch KCl và acetic acid có khả năng dẫn điện.", "ĐÚNG"),
        ("b) Dung dịch hòa tan ethanol không dẫn điện.", "ĐÚNG"),
        ("c) Các phân tử acetic acid có 1 đầu mang điện dương và một đầu mang điện âm, chuyển động trong dung dịch tạo thành dòng điện.", "SAI"),
        ("d) Ethanol là chất điện li.", "SAI")
    ]
    for stmt, ans in stmts_p2_c3:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c3 = (
            "a) Đúng vì KCl là chất điện li mạnh (đèn sáng mạnh), CH~3~COOH là chất điện li yếu (đèn sáng mờ).\n"
            "b) Đúng vì ethanol là chất không điện li, không tạo ion tự do nên đèn không sáng.\n"
            "c) Sai vì dòng điện trong dung dịch là dòng chuyển dịch có hướng của các ion tự do (CH~3~COO^- và H^+), không phải do phân tử lưỡng cực chuyển động.\n"
            "d) Sai vì ethanol là chất không điện li."
        )
        add_solution_no_table(doc, sol_p2_c3)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Trong lịch sử các thuyết về định nghĩa acid – base có thuyết Arrhenius và thuyết Brønsted–Lowry:\n- Thuyết Arrhenius: 'Acid là chất phân ly trong dung dịch nước tạo thành H^+ (proton) còn base là chất phân ly trong dung dịch nước tạo thành OH^-'\n- Thuyết Brønsted–Lowry: 'Acid là chất cho proton (H^+) và base là chất nhận proton'")
    stmts_p2_c4 = [
        ("a) Theo thuyết Brønsted–Lowry phản ứng HClO~4~ + CH~3~COOH ⇌ ClO~4~^- + CH~3~COOH~2~^+, trong phản ứng thuận thì CH~3~COOH là base, trong phản ứng nghịch thì ClO~4~^- là acid.", "SAI"),
        ("b) Thuyết Brønsted–Lowry tổng quát hơn thuyết Arrhenius, vì phân tử không có nhóm –OH như NH~3~, SO~3~^2-^, CO~3~^2-^ cũng là base.", "ĐÚNG"),
        ("c) Theo thuyết Brønsted–Lowry, các ion Fe^3+^, Cu^2+^, Mg^2+^, Na^+ được xem là acid.", "SAI"),
        ("d) Theo thuyết Arrhenius thì acid trong phân tử phải có nguyên tử H và base trong phân tử phải có nhóm –OH. Thuyết lí luận này đúng trong trường hợp dung môi sử dụng là nước.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c4:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c4 = (
            "a) Sai vì trong phản ứng nghịch, ClO~4~^- nhận proton nên là base; CH~3~COOH~2~^+ nhường proton nên là acid.\n"
            "b) Đúng vì mở rộng khái niệm base cho các chất không có nhóm –OH nhưng có khả năng nhận proton.\n"
            "c) Sai vì ion Na^+ không có khả năng nhường proton trong dung dịch nước (không bị thủy phân).\n"
            "d) Đúng theo phạm vi áp dụng của thuyết Arrhenius."
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
    render_question_prompt(p, "Câu 1. Cho dãy các chất: KAl(SO~4~)~2~.12H~2~O, C~2~H~5~OH, C~12~H~22~O~11~ (saccharose), CH~3~COOH, Ca(OH)~2~, CH~3~COONH~4~. Số chất điện li là")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *4*", base_size=10.5)
        add_solution_no_table(doc, "Có 4 chất điện li gồm:\n- Muối: KAl(SO~4~)~2~.12H~2~O, CH~3~COONH~4~.\n- Acid: CH~3~COOH.\n- Base: Ca(OH)~2~.\nHai chất còn lại (C~2~H~5~OH, C~12~H~22~O~11~) là chất không điện li.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Cho các phản ứng:\n(1) NaOH + HCl ➔ NaCl + H~2~O\n(2) H~2~ + I~2~ ⇌ 2HI\n(3) CaCO~3~ ⇌ CaO + CO~2~\n(4) 2KClO~3~ ➔ 2KCl + 3O~2~\nSố phản ứng thuận nghịch là:")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *2*", base_size=10.5)
        add_solution_no_table(doc, "Có 2 phản ứng thuận nghịch là (2) và (3). Phản ứng (1) và (4) là phản ứng một chiều.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3 (SBT-CTST). Trộn lẫn V ml dung dịch NaOH 0,01M với V ml dung dịch HCl 0,03M thu được 2V ml dung dịch Y. Tính pH của dung dịch Y.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *2*", base_size=10.5)
        add_solution_no_table(doc, "n(OH^-) = 0,01V mol; n(H^+) = 0,03V mol.\nH^+ + OH^- ➔ H~2~O  ➔  n(H^+ dư) = 0,03V - 0,01V = 0,02V mol.\n[H^+] = 0,02V / (2V) = 0,01 M = 10^-2^ M ➔ pH = 2.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Tính pH của dung dịch thu được sau khi trộn 40 mL dung dịch HCl 0,5 M với 60 mL dung dịch NaOH 0,5 M.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *13*", base_size=10.5)
        add_solution_no_table(doc, "n(HCl) = 0,040 . 0,5 = 0,02 mol; n(NaOH) = 0,060 . 0,5 = 0,03 mol.\nV~sau~ = 40 + 60 = 100 mL = 0,1 L.\nn(NaOH dư) = 0,03 - 0,02 = 0,01 mol ➔ [OH^-] = 0,01 / 0,1 = 0,1 M.\npOH = 1 ➔ pH = 14 - 1 = 13.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5 (SBT - KNTT). Ascorbic acid (vitamin C) là một acid hữu cơ được kí hiệu đơn giản là HAsc, phân tử khối là 176. Một học sinh hoà tan 5,0 g ascorbic acid vào 250 mL nước. Tính pH của dung dịch thu được, biết trong dung dịch có cân bằng sau: HAsc ⇌ H^+ + Asc^-; K~a~ = 8.10^-5^")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *2,5*", base_size=10.5)
        add_solution_no_table(doc, "C~0~ = 5,0 / (176 . 0,25) ≈ 0,114 M.\nHAsc ⇌ H^+ + Asc^-   K~a~ = 8.10^-5^\nCân bằng: [H^+] = x, [Asc^-] = x, [HAsc] = 0,114 - x.\nK~a~ = x^2^ / (0,114 - x) = 8.10^-5^ ➔ x ≈ 2,98.10^-3^ M ➔ pH = -log(2,98.10^-3^) ≈ 2,5.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6 (SBT - KNTT). Ethanol và propanoic acid phản ứng với nhau tạo thành ethyl propanoate theo phản ứng hoá học sau:\nC~2~H~5~OH(l) + C~2~H~5~COOH(l) ⇌ C~2~H~5~COOC~2~H~5~(l) + H~2~O(l)\nỞ 50°C, giá trị K~C~ của phản ứng trên là 7,5. Nếu cho 23,0 g ethanol phản ứng với 37,0 g propanoic acid ở 50°C thì khối lượng của ethyl propanoate thu được trong hỗn hợp ở trạng thái cân bằng là bao nhiêu? (Coi tổng thể tích của hệ phản ứng không đổi, kết quả tính theo đơn vị gam và làm tròn đến hàng phần trăm).")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *37,32*", base_size=10.5)
        add_solution_no_table(doc, "n(ethanol) = 23,0 / 46 = 0,5 mol; n(propanoic acid) = 37,0 / 74 = 0,5 mol.\nGọi số mol ester tạo thành ở trạng thái cân bằng là x (mol).\nK~C~ = x^2^ / (0,5 - x)^2^ = 7,5  ➔  x / (0,5 - x) = √7,5 ≈ 2,7386\n➔ x ≈ 0,366 mol.\nm(ester) = 0,366 . 102 ≈ 37,322 g ≈ 37,32 g.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # ==========================================
    # PHẦN IV
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN IV. Tự luận",
        subtitle="Học sinh giải chi tiết các bài toán sau vào giấy làm bài."
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. [CD - SBT] Quan sát Hình 1.1 và ghép mỗi đối tượng ở cột A với một mô tả thích hợp ở cột B.\n- Cột A: a) Đường (a); b) t~1~; c) Đường (b); d) t~2~\n- Cột B: 1. không phải là thời điểm bắt đầu trạng thái cân bằng; 2. mô tả biến thiên nồng độ chất sản phẩm theo thời gian; 3. Là thời điểm phản ứng đạt trạng thái cân bằng; 4. Mô tả biến thiên nồng độ chất phản ứng theo thời gian.")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau1_p4_do_thi_nong_do.png"), width_cm=7.5)
    if is_teacher:
        sol_p4_c1 = (
            "Ghép nối chi tiết:\n"
            "- a - 4: Đường (a) nồng độ giảm dần từ cực đại đến khi không đổi ➔ mô tả biến thiên nồng độ chất phản ứng.\n"
            "- b - 1: Tại t~1~ nồng độ chất tham gia bằng chất sản phẩm (cắt nhau), không phải là thời điểm bắt đầu cân bằng.\n"
            "- c - 2: Đường (b) nồng độ tăng dần từ 0 đến khi không đổi ➔ mô tả biến thiên nồng độ chất sản phẩm.\n"
            "- d - 3: Tại t~2~ nồng độ các chất bắt đầu không đổi theo thời gian ➔ là thời điểm đạt trạng thái cân bằng.\n"
            "➔ Kết quả ghép: a – 4, b – 1, c – 2, d – 3."
        )
        add_solution_no_table(doc, sol_p4_c1)
    else:
        add_student_writing_lines(doc, num_lines=5)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. [CD - SGK] Trong môi trường acid, diệp lục có màu vàng đến đỏ; còn trong môi trường kiềm, diệp lục có màu xanh.\na) Giải thích vì sao khi vắt chanh vào nước luộc rau muống thì màu xanh của nước lại bị nhạt đi.\nb*) Vì sao khi luộc bánh chưng, cho thêm một chút thuốc muối (NaHCO~3~) sẽ làm lá dong gói bánh có màu xanh đẹp hơn?")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau2_p4_chanh_rau_muong_banh_chung.png"), width_cm=10.0)
    if is_teacher:
        sol_p4_c2 = (
            "a) Trong nước luộc rau muống có chứa một lượng chất diệp lục tan vào nước. Khi vắt chanh (chứa citric acid là acid hữu cơ yếu), nồng độ H^+ tăng làm thay đổi môi trường sang acid, ion H^+ thế chỗ ion Mg^2+^ trong nhân porphyrin của diệp lục tạo thành pheophytin có màu vàng nhạt, làm màu xanh của nước rau muống bị nhạt đi.\n"
            "b) Trong nước, muối NaHCO~3~ bị thuỷ phân tạo môi trường kiềm yếu:\n"
            "   NaHCO~3~ ➔ Na^+ + HCO~3~^-\n"
            "   HCO~3~^- + H~2~O ⇌ H~2~CO~3~ + OH^-\n"
            "Môi trường kiềm này bảo vệ nhân Mg^2+^ của diệp lục không bị acid hóa và biến tính ở nhiệt độ cao, giúp lá dong gói bánh chưng giữ được màu xanh tươi đẹp mắt."
        )
        add_solution_no_table(doc, sol_p4_c2)
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Lưu file
    filename = "De_Kiem_Tra_Hoa_11_De02_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_11_De02.docx"
    filepath = os.path.join(OUT_DIR, filename)
    doc.save(filepath)
    print(f"Saved: {filepath}")

if __name__ == "__main__":
    print("Generating De_Kiem_Tra_Hoa_11_De02.docx (Học Sinh)...")
    build_exam_document(is_teacher=False)
    print("Generating De_Kiem_Tra_Hoa_11_De02_GiaoVien.docx (Giáo Viên)...")
    build_exam_document(is_teacher=True)
    print("Done!")
