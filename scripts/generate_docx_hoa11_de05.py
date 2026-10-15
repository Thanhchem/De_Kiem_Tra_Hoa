# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file Word cho ĐỀ SỐ 5 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De05.docx (Bản Học Sinh)
- De_Kiem_Tra_Hoa_11_De05_GiaoVien.docx (Bản Giáo Viên)
Chương 2: NITROGEN – SULFUR
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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de05")

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
        exam_title="ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 2 (NITROGEN – SULFUR)",
        subject_line="MÔN: HOÁ HỌC - LỚP 11",
        num_pages="04",
        duration="45 phút",
        exam_code="105",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; Ar = 40; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Ba = 137",
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
    render_question_prompt(p, "Câu 1. Nguyên tố dinh dưỡng nào sau đây có trong cấu tạo của protein?")
    opts_c1 = [
        "A. Sắt.",
        "B. Kẽm.",
        "*C. Nitrogen.*",
        "D. Đồng."
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Protein là hợp chất hữu cơ cao phân tử chứa các nguyên tố chính là C, H, O, N và có thể chứa S, P,... Trong đó nguyên tố nitrogen (N) là thành phần cấu tạo đặc trưng của các đơn vị amino acid tạo nên chuỗi polypeptide của protein.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Để tạo độ xốp cho một số loại bánh, có thể dùng muối nào sau đây?")
    opts_c2 = [
        "A. CaCO~3~.",
        "*B. NH~4~HCO~3~.*",
        "C. NH~4~Cl.",
        "D. NaCl."
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Muối ammonium hydrogencarbonate (NH~4~HCO~3~, thường gọi là bột nở) dễ bị nhiệt phân hủy khi nướng bánh theo phản ứng:\nNH~4~HCO~3~ -(t°)-> NH~3~ + CO~2~ + H~2~O.\nKhí NH~3~ và CO~2~ sinh ra thoát ra ngoài tạo thành các lỗ xốp li ti làm bánh nở xốp mềm.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Mưa acid ảnh hưởng tới hệ thực vật, phá hủy các vật liệu bằng kim loại, các bức tượng bằng đá, gây bệnh cho con người và động vật. Hiện tượng trên gây ra chủ yếu do khí thải của nhà máy nhiệt điện, phương tiện giao thông và sản xuất công nghiệp. Tác nhân chủ yếu trong khí thải gây ra mưa acid là")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau3_so_do_mua_acid.png"), width_cm=8.5)
    opts_c3 = [
        "*A. SO~2~ và NO~2~.*",
        "B. CH~4~ và NH~3~.",
        "C. CO và CH~4~.",
        "D. CO và CO~2~."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hai khí thải chủ yếu gây ra hiện tượng mưa acid là sulfur dioxide (SO~2~) và các oxide của nitrogen (NO~x~, chủ yếu là NO~2~). Các khí này bị oxi hóa trong khí quyển kết hợp với nước mưa tạo ra sulfuric acid (H~2~SO~4~) và nitric acid (HNO~3~).")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4 (SBT - KNTT). Oxide của nitrogen được tạo thành ở nhiệt độ cao, khi nitrogen có trong không khí bị oxi hóa được gọi là")
    opts_c4 = [
        "A. NO~x~ tức thời.",
        "*B. NO~x~ nhiệt.*",
        "C. NO~x~ nhiên liệu.",
        "D. NO~x~ tự nhiên."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khi đốt cháy nhiên liệu ở nhiệt độ cao (trên 1200 °C - 1500 °C) trong không khí hoặc khi có hiện tượng phóng điện sấm sét, N~2~ trong không khí phản ứng với O~2~ tạo thành các oxide của nitrogen. Nguồn NO~x~ này được gọi là NO~x~ nhiệt (thermal NO~x~).")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5 (SBT - KNTT). Công thức hoá học của diêm tiêu Chile là")
    opts_c5 = [
        "A. Ca(NO~3~)~2~.",
        "B. NH~4~NO~3~.",
        "C. NH~4~Cl.",
        "*D. NaNO~3~.*"
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Diêm tiêu Chile (Chile saltpeter) là tên gọi khoáng vật tự nhiên của sodium nitrate có công thức hóa học là NaNO~3~ (phân biệt với diêm tiêu thông thường là KNO~3~).")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6 (SBT - KNTT). Sulfur đóng vai trò chất khử khi tác dụng với đơn chất nào sau đây?")
    opts_c6 = [
        "A. Fe.",
        "*B. O~2~.*",
        "C. H~2~.",
        "D. Hg."
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khi tác dụng với oxygen: S^0^ + O~2~ -(t°)-> S^+4^O~2~. Số oxi hóa của S tăng từ 0 lên +4 nên sulfur đóng vai trò là chất khử. Trong các phản ứng với kim loại (Fe, Hg) hoặc hydrogen (H~2~), số oxi hóa của S giảm từ 0 xuống -2 nên S đóng vai trò là chất oxi hóa.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 7 (SBT - CTST). Cách pha loãng dung dịch H~2~SO~4~ đặc nào sau đây đúng?")
    opts_c7 = [
        "A. Rót nhanh acid vào nước và khuấy đều.",
        "B. Rót nhanh nước vào acid và khuấy đều.",
        "C. Rót từ từ nước vào acid và khuấy đều.",
        "*D. Rót từ từ acid vào nước và khuấy đều.*"
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Quy tắc an toàn khi pha loãng H~2~SO~4~ đặc: Luôn luôn rót từ từ dung dịch sulfuric acid đặc dọc theo đũa thủy tinh vào chậu/cốc nước và khuấy nhẹ đều. Tuyệt đối không được làm ngược lại (rót nước vào acid) vì nhiệt hydrate hóa tỏa ra cực lớn sẽ làm nước sôi đột ngột và bắn các giọt acid nóng ra ngoài gây bỏng nguy hiểm.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 8 (SBT - KNTT). Trong nghiên cứu, khí nitrogen thường được dùng để tạo bầu khí quyển trơ dựa trên cơ sở nào?")
    opts_c8 = [
        "A. Nitrogen có tính oxi hoá mạnh.",
        "*B. Nitrogen rất bền với nhiệt.*",
        "C. Nitrogen khó hoá lỏng.",
        "D. Nitrogen không có cực."
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tử nitrogen (N~2~) có liên kết ba N≡N rất bền vững với năng lượng liên kết cực lớn (945 kJ/mol). Ở nhiệt độ thường và nhiệt độ vừa phải, nitrogen rất bền với nhiệt và rất trơ về mặt hóa học, do đó được dùng làm môi trường trơ để bảo quản hóa chất và vật liệu dễ bị oxi hóa.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 9 (SBT - KNTT). Xét cân bằng hóa học trong dung dịch nước:\nNH~3~ + H~2~O ⇌ NH~4~^+ + OH^-\nHằng số cân bằng (K~C~) của phản ứng được biểu diễn bằng biểu thức nào sau đây?")
    opts_c9 = [
        "*A. K~C~ = ([NH~4~^+][OH^-]) / [NH~3~]*",
        "B. K~C~ = ([NH~4~^+][OH^-]) / ([NH~3~][H~2~O])",
        "C. K~C~ = ([NH~4~^+][OH^-]) / [H~2~O]",
        "D. K~C~ = [NH~4~^+] / [NH~3~]"
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong dung dịch nước loãng, nồng độ của dung môi H~2~O rất lớn và coi như không đổi nên được gộp vào hằng số cân bằng (hằng số base K~b~ hoặc K~C~ của dung dịch). Biểu thức hằng số cân bằng: K~C~ = ([NH~4~^+][OH^-]) / [NH~3~].")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 10. Trong các phản ứng dưới đây, phản ứng nào NH~3~ không thể hiện tính khử?")
    opts_c10 = [
        "A. 4NH~3~ + 5O~2~ -(xt, t°)-> 4NO + 6H~2~O",
        "*B. NH~3~ + HCl ➔ NH~4~Cl*",
        "C. 8NH~3~ + 3Cl~2~ ➔ 6NH~4~Cl + N~2~",
        "D. 2NH~3~ + 3CuO -(t°)-> 3Cu + 3H~2~O + N~2~"
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong phản ứng NH~3~ + HCl ➔ NH~4~Cl, số oxi hóa của nitrogen giữ nguyên là -3; phản ứng này thể hiện tính base của NH~3~ (nhận proton H^+). Ở 3 phản ứng còn lại, số oxi hóa của N tăng từ -3 lên +2 (trong NO) hoặc 0 (trong N~2~), chứng tỏ NH~3~ đóng vai trò chất khử.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 11. Thí nghiệm với dung dịch HNO~3~ thường sinh ra khí độc NO~2~. Để hạn chế khí NO~2~ thoát ra từ ống nghiệm, người ta nút ống nghiệm bằng:\n(a) bông khô.                (b) bông có tẩm nước.\n(c) bông có tẩm nước vôi.        (d) bông có tẩm giấm ăn.\nTrong 4 biện pháp trên, biện pháp có hiệu quả nhất là")
    opts_c11 = [
        "A. (d)",
        "*B. (c)*",
        "C. (a)",
        "D. (b)"
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khí NO~2~ là một oxide acid độc hại. Khi tẩm nước vôi trong Ca(OH)~2~ vào nút bông, NO~2~ sẽ phản ứng nhanh và bị giữ lại hoàn toàn dưới dạng muối tan theo phản ứng:\n2NO~2~ + Ca(OH)~2~ ➔ Ca(NO~2~)~2~ + Ca(NO~3~)~2~ + H~2~O.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 12. Ứng dụng nào sau đây không phải của SO~2~?")
    opts_c12 = [
        "A. Sản xuất sulfuric acid.",
        "B. Tẩy màu dung dịch đường.",
        "*C. Khử mùi không khí.*",
        "D. Tẩy trắng giấy, bột giấy."
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Sulfur dioxide (SO~2~) là một chất khí có mùi hắc rất khó chịu, gây cay mắt, ho, kích ứng hệ hô hấp mạnh và có độc tính, vì vậy tuyệt đối không dùng để khử mùi không khí. Ứng dụng chính của SO~2~ là sản xuất H~2~SO~4~, tẩy trắng bột giấy, tẩy màu nước đường và chống mốc nông sản.")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 13. Trong điều kiện thích hợp, xảy ra các phản ứng sau:\n(a) 2H~2~SO~4~ + C ➔ 2SO~2~ + CO~2~ + 2H~2~O\n(b) H~2~SO~4~ + Fe(OH)~2~ ➔ FeSO~4~ + 2H~2~O\n(c) 4H~2~SO~4~ + 2FeO ➔ Fe~2~(SO~4~)~3~ + SO~2~ + 4H~2~O\n(d) 6H~2~SO~4~ + 2Fe ➔ Fe~2~(SO~4~)~3~ + 3SO~2~ + 6H~2~O\nTrong các phản ứng trên, phản ứng xảy ra với dung dịch H~2~SO~4~ loãng là")
    opts_c13 = [
        "A. (a)",
        "B. (c)",
        "*C. (b)*",
        "D. (d)"
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phản ứng (b) là phản ứng trung hòa trao đổi giữa acid loãng và base không tan Fe(OH)~2~ tạo muối Fe(II) và nước. Các phản ứng (a), (c), (d) là phản ứng oxi hóa - khử giải phóng khí SO~2~ chỉ xảy ra khi dùng dung dịch H~2~SO~4~ đặc nóng.")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 14. Dẫn khí SO~2~ vào dung dịch KMnO~4~ hiện tượng quan sát được là")
    opts_c14 = [
        "A. Dung dịch không màu chuyển sang màu tím.",
        "B. Dung dịch màu tím bị vẩn đục màu vàng.",
        "C. Màu tím của dung dịch KMnO~4~ chuyển sang màu vàng.",
        "*D. Màu tím của dung dịch KMnO~4~ chuyển sang không màu.*"
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Do SO~2~ có tính khử mạnh tác dụng với chất oxi hóa mạnh KMnO~4~, làm dung dịch thuốc tím chuyển sang không màu:\n5SO~2~ + 2KMnO~4~ + 2H~2~O ➔ K~2~SO~4~ + 2MnSO~4~ + 2H~2~SO~4~ (dung dịch MnSO~4~ không màu).")

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 15. Nguyên nhân chủ yếu gây ra hiện tượng phú dưỡng trong các thủy vực là gì?")
    opts_c15 = [
        "A. Sự tích tụ các chất hữu cơ do phân hủy xác sinh vật.",
        "*B. Sự gia tăng quá mức các chất dinh dưỡng như nitrate (NO~3~^-) và phosphate (PO~4~^3^-) từ nước thải sinh hoạt, phân bón nông nghiệp và công nghiệp.*",
        "C. Sự thay đổi nhiệt độ nước trong thủy vực.",
        "D. Quá trình xói mòn đất làm thay đổi hệ sinh thái nước ngọt."
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hiện tượng phú dưỡng (eutrophication) xảy ra khi nguồn nước tích tụ nồng độ quá cao các nguyên tố dinh dưỡng như nitrogen (NO~3~^-, NH~4~^+) và phosphorus (PO~4~^3^-) từ phân bón nông nghiệp dư thừa và nước thải sinh hoạt chưa xử lý, làm bùng phát các loài tảo và vi sinh vật.")

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 16. Mưa acid có thể gây ra tác động tiêu cực nào đối với các công trình xây dựng?")
    opts_c16 = [
        "A. Làm tăng độ bền của các công trình xây dựng bằng đá vôi.",
        "*B. Gây ăn mòn các công trình làm từ đá vôi, đá cẩm thạch và kim loại.*",
        "C. Tăng khả năng chống chịu của các vật liệu xây dựng trước sự tác động của môi trường.",
        "D. Không có tác động đến các công trình xây dựng."
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nước mưa có chứa các acid mạnh như H~2~SO~4~ và HNO~3~ tác dụng ăn mòn đá vôi và đá cẩm thạch (chứa CaCO~3~), đồng thời oxi hóa ăn mòn các kết cấu thép, cầu sắt và tượng đài kim loại.")

    # Câu 17
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 17. Biện pháp nào dưới đây có thể giúp giảm thiểu hiện tượng mưa acid?")
    opts_c17 = [
        "A. Tăng cường việc sử dụng các loại nhiên liệu hóa thạch.",
        "*B. Giảm lượng khí thải SO~2~ và NO~x~ từ các nhà máy công nghiệp và phương tiện giao thông.*",
        "C. Sử dụng nhiều hơn các loại phân bón chứa phosphate.",
        "D. Khuyến khích việc phát triển giao thông đường bộ."
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Cắt giảm lượng phát thải các khí tiền thân gây acid là SO~2~ và NO~x~ bằng cách xử lý khí thải bằng nước vôi/kiềm, chuyển dịch sang năng lượng tái tạo và sử dụng động cơ đạt chuẩn khí thải cao là giải pháp hiệu quả nhất ngăn chặn mưa acid.")

    # Câu 18
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 18. Cho các phát biểu sau:\n(1). Cho thanh Fe (iron) vào dung dịch H~2~SO~4~ loãng, iron tan ra, sủi bọt khí.\n(2). Cho dung dịch BaCl~2~ vào dung dịch H~2~SO~4~, xuất hiện kết tủa đen.\n(3). Ở nhiệt độ cao, sulfur tác dụng với hydrogen tạo thành hydrogen sulfide (H~2~S).\n(4). Cho từng giọt sulfuric acid đặc vào đường saccharose (C~12~H~22~O~11~) màu trắng, đường chuyển màu nâu ➔ đen sau đó trào lên khỏi miệng cốc.\n(5). Nitrogen lỏng được dùng để làm lạnh nhanh, bảo quản mẫu vật phẩm trong y học.\n(6). Nitrogen bền nhiệt, có tính trơ nên không bảo quản được hương vị thực phẩm.\n(7). Không khí có chứa oxygen là tác nhân oxi hóa thực phẩm, làm thực phẩm ôi thiu chậm hơn trong khí nitrogen.\nSố phát biểu đúng là:")
    opts_c18 = [
        "A. 3.",
        "*B. 4.*",
        "C. 5.",
        "D. 6."
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    if is_teacher:
        sol_c18 = (
            "Có 4 phát biểu đúng gồm: (1), (3), (4), (5).\n"
            "- (1) Đúng: Fe + H~2~SO~4~ ➔ FeSO~4~ + H~2~↑ (sủi bọt khí).\n"
            "- (2) Sai: Sinh ra kết tủa trắng barium sulfate (BaSO~4~), không phải kết tủa đen.\n"
            "- (3) Đúng: H~2~ + S -(t°)-> H~2~S.\n"
            "- (4) Đúng: H~2~SO~4~ đặc háo nước than hóa đường thành C (đen), sau đó oxi hóa C thành CO~2~ và SO~2~ đẩy trào cột than lên.\n"
            "- (5) Đúng: Nitrogen lỏng có nhiệt độ -196 °C dùng bảo quản mô, mẫu sinh học.\n"
            "- (6) Sai: Nhờ tính trơ và ức chế vi sinh vật nên nitrogen bảo quản tốt độ giòn và hương vị thực phẩm.\n"
            "- (7) Sai: Oxygen là tác nhân gây oxi hóa thực phẩm làm nhanh ôi thiu hơn so với khi bảo quản trong khí trơ nitrogen."
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
    render_question_prompt(p, "Câu 1. Để giảm sốt hoặc giảm đau, người ta có thể dùng túi chườm lạnh chứa hóa chất. Hãy tìm hiểu về loại túi chườm lạnh này:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau1_p2_tui_chuom_lanh.png"), width_cm=7.0)
    stmts_p2_c1 = [
        ("a) Hóa chất trong túi chườm lạnh gồm ammonium chloride rắn và nước cất.", "ĐÚNG"),
        ("b) Túi chườm sẽ hoạt động khi phá vỡ lớp ngăn cách giữa muối ammonium chloride và nước cất.", "ĐÚNG"),
        ("c) Do sự hòa tan NH~4~Cl(s) vào nước là một quá trình tỏa nhiệt: NH~4~Cl(s) ➔ NH~4~Cl(aq)  Δ~r~H°~298~ < 0.", "SAI"),
        ("d) Phản ứng tỏa nhiệt, muối hấp thụ nhiệt từ nước, làm nước giảm nhiệt độ.", "SAI")
    ]
    for stmt, ans in stmts_p2_c1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c1 = (
            "a) Đúng: Túi chườm lạnh hóa học thường chứa muối ammonium chloride rắn và túi nước riêng biệt.\n"
            "b) Đúng: Khi bóp mạnh làm vỡ túi nước ngăn cách, muối hòa tan vào nước tạo hiệu ứng làm lạnh nhanh chóng.\n"
            "c) Sai: Quá trình hòa tan NH~4~Cl vào nước là một quá trình thu nhiệt rất mạnh (Δ~r~H°~298~ > 0), làm nhiệt độ giảm sâu.\n"
            "d) Sai: Đây là quá trình thu nhiệt; túi chườm lạnh hấp thụ nhiệt từ vùng cơ thể tiếp xúc giúp hạ nhiệt độ cơ thể và giảm đau sưng."
        )
        add_solution_no_table(doc, sol_p2_c1)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Biện pháp nào sau đây dùng để hạn chế hiện tượng phú dưỡng?")
    stmts_p2_c2 = [
        ("a) Tạo điều kiện để nước trong kênh rạch, ao, hồ được lưu thông.", "ĐÚNG"),
        ("b) Xả nước thải (nông nghiệp, công nghiệp và sinh hoạt) trực tiếp đến ao, hồ thông qua các cống dẫn nước cố định.", "SAI"),
        ("c) Sử dụng phân bón đúng liều lượng, đúng cách, đúng thời điểm trong năm để hạn chế sự rửa trôi của ion nitrate và phosphate từ nguồn phân bón dư thừa vào kênh rạch, ao, hồ.", "ĐÚNG"),
        ("d) Xử lí nước thải trước khi cho vào kênh rạch, ao, hồ.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c2 = (
            "a) Đúng: Nước được lưu thông làm tăng lượng oxygen hòa tan, ngăn chặn sự tù đọng và tích tụ chất dinh dưỡng.\n"
            "b) Sai: Việc xả trực tiếp nước thải chưa qua xử lý vào ao hồ là nguyên nhân chính gây phú dưỡng và ô nhiễm nặng nề.\n"
            "c) Đúng: Bón phân hợp lý và đúng kỹ thuật hạn chế tối đa lượng nitrate và phosphate dư thừa bị mưa rửa trôi xuống thủy vực.\n"
            "d) Đúng: Xử lý nước thải sinh hoạt và công nghiệp giúp khử bỏ triệt để các hợp chất chứa N và P trước khi xả thải."
        )
        add_solution_no_table(doc, sol_p2_c2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Với lứa tuổi học sinh có thể giúp giảm thiểu lượng sulfur dioxide thải vào không khí bằng các việc làm phù hợp:")
    stmts_p2_c3 = [
        ("a) Khuyến khích sử dụng xe điện, sử dụng nguyên liệu thân thiện môi trường như xăng E5,...", "ĐÚNG"),
        ("b) Khuyến khích đốt rác thải.", "SAI"),
        ("c) Tuyên truyền tác hại của SO~2~ và lợi ích của nhiên liệu tái tạo, bảo vệ môi trường,...", "ĐÚNG"),
        ("d) Giảm rác thải, phân loại rác, tái chế, tái sử dụng.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c3:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c3 = (
            "a) Đúng: Giảm sử dụng nhiên liệu hóa thạch chứa lưu huỳnh bằng việc đi xe điện, xe đạp hoặc nhiên liệu sinh học.\n"
            "b) Sai: Việc đốt rác thải tự phát làm phát sinh khói bụi chứa nhiều khí độc hại và khí SO~2~ gây ô nhiễm trầm trọng.\n"
            "c) Đúng: Nâng cao nhận thức xã hội qua các hoạt động tuyên truyền học đường.\n"
            "d) Đúng: Giảm thiểu, tái sử dụng và tái chế rác giúp tiết kiệm năng lượng và tài nguyên sản xuất, gián tiếp cắt giảm lượng khí thải độc hại."
        )
        add_solution_no_table(doc, sol_p2_c3)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Sulfuric acid có tính acid, tính oxi hóa mạnh, tính háo nước:")
    stmts_p2_c4 = [
        ("a) Sulfuric acid đặc phản ứng với carbon trong than.", "ĐÚNG"),
        ("b) Sulfuric acid đặc oxi hoá hợp chất Fe(II) thành hợp chất Fe(III).", "ĐÚNG"),
        ("c) Cho bột sodium hydrogencarbonate vào sulfuric acid loãng xuất hiện kết tủa trắng.", "SAI"),
        ("d) Chất rắn copper (II) sulfate pentahydrate có màu xanh tác dụng với sulfuric acid đặc sẽ biến thành copper (II) sulfate khan màu trắng.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c4:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c4 = (
            "a) Đúng: 2H~2~SO~4~(đặc) + C -(t°)-> 2SO~2~ + CO~2~ + 2H~2~O.\n"
            "b) Đúng: H~2~SO~4~ đặc nóng oxi hoá Fe(II) thành muối Fe(III): 2FeO + 4H~2~SO~4~ ➔ Fe~2~(SO~4~)~3~ + SO~2~ + 4H~2~O.\n"
            "c) Sai: 2NaHCO~3~ + H~2~SO~4~ ➔ Na~2~SO~4~ + 2CO~2~↑ + 2H~2~O. Phản ứng sinh ra bọt khí không màu và dung dịch trong suốt, không xuất hiện kết tủa.\n"
            "d) Đúng: Do tính háo nước mãnh liệt của H~2~SO~4~ đặc, nó hút nước kết tinh từ tinh thể ngậm nước: CuSO~4~.5H~2~O (màu xanh) ➔ CuSO~4~ khan (màu trắng) + 5H~2~O."
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
    render_question_prompt(p, "Câu 1. Phú dưỡng là hệ quả sau khi ao ngòi, sông hồ nhận quá nhiều các nguồn thải chứa các chất dinh dưỡng chứa nguyên tố nitrogen, phosphorus. Khi hàm lượng phosphorus vượt quá 20 μg/L và hàm lượng nitrogen vượt quá ngưỡng bao nhiêu μg/L thì xảy ra hiện tượng phú dưỡng?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *300*", base_size=10.5)
        add_solution_no_table(doc, "Theo tiêu chuẩn sinh thái học môi trường thủy vực, khi hàm lượng nitrogen trong nước vượt quá 300 μg/L và phosphorus vượt quá 20 μg/L thì nước hồ chuyển sang tình trạng phú dưỡng.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Cho các chất khí sau: H~2~S, NO, NO~2~, SO~2~. Số khí gây ô nhiễm môi trường khi phát thải vào không khí là bao nhiêu?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *4*", base_size=10.5)
        add_solution_no_table(doc, "Cả 4 khí trên (H~2~S, NO, NO~2~, SO~2~) đều là những chất khí gây ô nhiễm môi trường không khí và có độc tính đối với sức khỏe sinh vật.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3 (SBT - KNTT). Tính khối lượng riêng (g/L) của không khí ở điều kiện chuẩn, giả thiết thành phần không khí gồm: 78% nitrogen, 21% oxygen và 1% argon (làm tròn kết quả đến hàng phần trăm).")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *1,17*", base_size=10.5)
        add_solution_no_table(doc, "Khối lượng mol trung bình của không khí:\nM_kk = 0,78 . 28,014 + 0,21 . 31,998 + 0,01 . 39,948 = 21,851 + 6,720 + 0,399 = 28,97 g/mol.\nỞ điều kiện chuẩn (25 °C, 1 bar), thể tích mol phân tử của khí là V = 24,79 L.\nKhối lượng riêng: D = m / V = 28,97 / 24,79 ≈ 1,1686 g/L ≈ 1,17 g/L.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Trong công nghiệp, ammonia được sản xuất theo phản ứng pha khí:\nN~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)   Δ~r~H°~298~\nCho biết các giá trị năng lượng liên kết E_b (kJ·mol^-1^): E_b(N≡N) = 945; E_b(H–H) = 436; E_b(N–H) = 386.\nNhiệt tạo thành Δ~f~H°~298~ (kJ·mol^-1^) của NH~3~(g) là -x kJ·mol^-1^. Giá trị của x là bao nhiêu?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *31,5*", base_size=10.5)
        add_solution_no_table(doc, "Biến thiên enthalpy phản ứng:\nΔ~r~H°~298~ = E_b(N≡N) + 3.E_b(H–H) - 6.E_b(N–H) = 945 + 3.436 - 6.386 = -63 kJ.\nVì phản ứng tạo thành 2 mol NH~3~(g) từ các đơn chất ở trạng thái chuẩn, nên:\nΔ~f~H°~298~(NH~3~) = Δ~r~H°~298~ / 2 = -63 / 2 = -31,5 kJ/mol ➔ x = 31,5.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Sự phụ thuộc của độ tan khí ammonia trong nước vào nhiệt độ được mô tả ở hình bên.\nDựa vào đồ thị ở hình bên, hãy xác định: Nồng độ phần trăm của dung dịch ammonia bão hòa ở 30 °C (làm tròn kết quả đến hàng phần mười).")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau5_p3_do_thi_do_tan_nh3.png"), width_cm=7.5)
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *28,6%* (hoặc *28,6*)", base_size=10.5)
        add_solution_no_table(doc, "Từ đồ thị, tại 30 °C, độ tan của khí NH~3~ trong nước là S = 40 g NH~3~ trong 100 g H~2~O.\nKhối lượng dung dịch bão hòa: m_dd = m_ct + m_dm = 40 + 100 = 140 g.\nNồng độ phần trăm của dung dịch: C% = (40 / 140) . 100% ≈ 28,57% ≈ 28,6%.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Sau hiện tượng phú dưỡng, một hồ nước có diện tích bề mặt 1 km^2^ và độ sâu trung bình 5 m chứa một lượng lớn xác tảo. Lượng oxygen hòa tan (DO) ban đầu trong hồ là 8 mg/L, nhưng sau quá trình phân hủy xác tảo, nồng độ oxygen giảm xuống còn 2 mg/L. Lượng oxygen (tính bằng kg) đã bị tiêu thụ trong toàn bộ hồ là x nghìn kg. Tính giá trị x?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *30*", base_size=10.5)
        add_solution_no_table(doc, "Thể tích nước trong hồ: V = S . h = 1 km^2^ . 5 m = 10^6^ m^2^ . 5 m = 5.10^6^ m^3^ = 5.10^9^ L.\nLượng oxygen tiêu thụ trên mỗi lít nước: ΔDO = 8 - 2 = 6 mg/L.\nTổng lượng oxygen tiêu thụ: m(O~2~) = 6 mg/L . 5.10^9^ L = 30.10^9^ mg = 30 000 kg = 30 nghìn kg ➔ x = 30.")
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
    render_question_prompt(p, "Câu 1 [CĐ - SGK]. Viết phương trình hóa học minh họa tác động của mưa acid đối với calcium carbonate trong núi đá vôi và với kim loại iron có trong thép.")
    if is_teacher:
        sol_p4_c1 = (
            "Phương trình hóa học minh họa tác động của mưa acid:\n"
            "- Tác động đối với calcium carbonate (CaCO~3~) trong núi đá vôi:\n"
            "  CaCO~3~(s) + 2HNO~3~(aq) ➔ Ca(NO~3~)~2~(aq) + CO~2~(g) + H~2~O(l)\n"
            "  CaCO~3~(s) + H~2~SO~4~(aq) ➔ CaSO~4~(aq, s) + CO~2~(g) + H~2~O(l)\n"
            "- Tác động đối với kim loại iron (Fe) có trong thép:\n"
            "  Fe(s) + H~2~SO~4~(aq) ➔ FeSO~4~(aq) + H~2~(g)\n"
            "  Fe(s) + 4HNO~3~(loãng) ➔ Fe(NO~3~)~3~(aq) + NO(g) + 2H~2~O(l)"
        )
        add_solution_no_table(doc, sol_p4_c1)
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 (SBT - KNTT). Trong công nghiệp, ammonia được sản xuất theo phản ứng pha khí:\nN~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)   Δ~r~H°~298~\nCho biết các giá trị năng lượng liên kết E_b (kJ·mol^-1^): E_b(N≡N) = 945; E_b(H–H) = 436; E_b(N–H) = 386.\na) Tính nhiệt phản ứng Δ~r~H°~298~ của phản ứng ở điều kiện chuẩn, nhận xét về dấu và độ lớn của giá trị tìm được.\nb) Tính nhiệt tạo thành Δ~f~H°~298~ (kJ·mol^-1^) của NH~3~(k).")
    if is_teacher:
        sol_p4_c2 = (
            "a) Tính nhiệt phản ứng Δ~r~H°~298~ theo năng lượng liên kết:\n"
            "   Δ~r~H°~298~ = E_b(N≡N) + 3.E_b(H–H) - 6.E_b(N–H)\n"
            "   = 945 + 3 . 436 - 6 . 386 = 945 + 1308 - 2316 = -63 kJ.\n"
            "   (Ghi chú: Nếu lấy số liệu E_b(N–H) = 389 kJ/mol theo một số tài liệu thì Δ~r~H°~298~ = -78 kJ).\n"
            "   Nhận xét: Giá trị Δ~r~H°~298~ mang dấu âm (< 0), chứng tỏ đây là một phản ứng tỏa nhiệt; nhiệt lượng tỏa ra khi tổng hợp 2 mol NH~3~(g) ở điều kiện chuẩn là 63 kJ.\n"
            "b) Tính nhiệt tạo thành Δ~f~H°~298~ của NH~3~(g):\n"
            "   Vì phản ứng trên tạo ra 2 mol NH~3~(g) từ các đơn chất N~2~(g) và H~2~(g) ở trạng thái chuẩn bền nhất, nên nhiệt tạo thành của 1 mol NH~3~(g) là:\n"
            "   Δ~f~H°~298~(NH~3~, g) = Δ~r~H°~298~ / 2 = -63 / 2 = -31,5 kJ/mol (hoặc -39 kJ/mol)."
        )
        add_solution_no_table(doc, sol_p4_c2)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Lưu file
    filename = "De_Kiem_Tra_Hoa_11_De05_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_11_De05.docx"
    filepath = os.path.join(OUT_DIR, filename)
    doc.save(filepath)
    print(f"Saved: {filepath}")

if __name__ == "__main__":
    print("Generating De_Kiem_Tra_Hoa_11_De05.docx (Học Sinh)...")
    build_exam_document(is_teacher=False)
    print("Generating De_Kiem_Tra_Hoa_11_De05_GiaoVien.docx (Giáo Viên)...")
    build_exam_document(is_teacher=True)
    print("Done!")
