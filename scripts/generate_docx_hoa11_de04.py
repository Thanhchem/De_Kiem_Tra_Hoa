# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file Word cho ĐỀ SỐ 4 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De04.docx (Bản Học Sinh)
- De_Kiem_Tra_Hoa_11_De04_GiaoVien.docx (Bản Giáo Viên)
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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de04")

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
        exam_code="104",
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
    render_question_prompt(p, "Câu 1. Không khí sạch là không khí ít tạp chất chứa các thành phần cơ bản của không khí. Trường hợp nào sau đây được coi là không khí sạch?")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cau1_khong_khi_sach.png"), width_cm=6.5)
    opts_c1 = [
        "A. Không khí chứa 78% N~2~, 17% O~2~, 3% CO~2~, 1% CO, 1% SO~2~.",
        "B. Không khí chứa 78% N~2~, 18% O~2~, 4% hỗn hợp CO~2~, SO~2~, HCl.",
        "C. Không khí chứa 78% N~2~, 20% O~2~, 2% hỗn hợp CO~2~, CH~4~ và bụi.",
        "*D. Không khí chứa 78% N~2~, 21% O~2~, 1% hỗn hợp CO~2~, H~2~O.*"
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Thành phần tự nhiên của không khí sạch gồm khoảng 78% thể tích N~2~, 21% O~2~ và khoảng 1% các khí khác (CO~2~, hơi nước, khí hiếm). Các trường hợp có chứa khí độc như CO, SO~2~, HCl hay bụi mịn đều bị ô nhiễm.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 [CTST - SGK]. Có các loại phân bón như NH~4~Cl, NH~4~NO~3~, (NH~4~)~2~SO~4~, các loại phân bón này không thích hợp bón cho đất nào sau đây?")
    opts_c2 = [
        "*A. Đất chua*",
        "B. Đất phù sa",
        "C. Đất bạc màu",
        "D. Đất nghèo dinh dưỡng."
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Các phân bón này chứa ion NH~4~^+ là acid cation theo thuyết Brønsted – Lowry, khi thuỷ phân tạo ra ion H^+ (hoặc H~3~O^+): NH~4~^+ + H~2~O ⇌ NH~3~ + H~3~O^+ làm tăng độ chua của đất, nên không bón cho đất chua.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Biện pháp nào sau đây không hạn chế tình trạng mưa acid?")
    opts_c3 = [
        "A. Cho lắp đặt các thiết bị khử sunphua tại những nhà máy nhiệt điện.",
        "B. Loại bỏ nitrogen và sulfur trong dầu mỏ, than.",
        "C. Sử dụng các loại năng lượng có khả năng tái tạo thân thiện với môi trường.",
        "*D. Phát tán chất NO~2~ và SO~2~ vào bầu khí quyển.*"
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "NO~2~ và SO~2~ là hai nguyên nhân chính gây ra mưa acid, việc phát tán các khí này vào khí quyển làm gia tăng mưa acid.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4 [KNTT - SBT]. Mưa acid là một thảm họa thiên nhiên toàn cầu, ảnh hưởng đến sự sống của các sinh vật. Mưa acid là hiện tượng mưa có pH")
    opts_c4 = [
        "*A. < 5,6.*",
        "B. = 7",
        "C. 6 – 7.",
        "D. > 8."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nước mưa thông thường hòa tan CO~2~ có pH khoảng 5,6. Hiện tượng mưa acid xảy ra khi nước mưa có pH < 5,6 do chứa acid sulfuric và acid nitric.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5 (SBT - KNTT). Hiện tượng phú dưỡng là một biểu hiện của môi trường ao, hồ bị ô nhiễm do dư thừa các chất dinh dưỡng. Sự dư thừa dinh dưỡng chủ yếu do hàm lượng các ion nào sau đây vượt quá mức cho phép?")
    opts_c5 = [
        "A. Sodium, potassium.",
        "B. Calcium, magnesium.",
        "*C. Nitrate, phosphate.*",
        "D. Chloride, sulfate."
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Hiện tượng phú dưỡng xảy ra do sự tích tụ quá mức của nguyên tố dinh dưỡng nitrogen (ở dạng ion nitrate NO~3~^-) và phosphorus (ở dạng ion phosphate PO~4~^3-^).")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6 (SBT - KNTT). Trong số các chất khí: SO~2~, CO~2~, O~2~, N~2~, khí tan tốt trong nước ở điều kiện thường là")
    opts_c6 = [
        "A. O~2~.",
        "B. CO~2~.",
        "*C. SO~2~.*",
        "D. N~2~."
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "SO~2~ là phân tử phân cực mạnh và phản ứng thuận nghịch với nước tạo H~2~SO~3~, ở 20°C 1 lít nước hòa tan được khoảng 40 lít khí SO~2~. CO~2~ tan ít hơn nhiều, còn O~2~ và N~2~ rất ít tan.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 7 (SBT - CTST). Dung dịch sulfuric acid đặc khác dung dịch sulfuric acid loãng ở tính chất hoá học nào?")
    opts_c7 = [
        "A. Tính base mạnh.",
        "*B. Tính oxi hóa mạnh.*",
        "C. Tính acid mạnh.",
        "D. Tính khử mạnh."
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "H~2~SO~4~ đặc có tính oxi hoá rất mạnh (do S ở mức oxi hoá cực đại +6), có khả năng oxi hoá hầu hết kim loại (trừ Au, Pt), nhiều phi kim (C, S, P) và nhiều hợp chất khử. H~2~SO~4~ loãng chỉ có tính oxi hóa của H^+.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 8 (SBT - KNTT). Bậc liên kết và năng lượng liên kết trong phân tử nitrogen tương ứng là")
    opts_c8 = [
        "A. 2 và 418 kJ/mol",
        "B. 1 và 167 kJ/mol.",
        "C. 1 và 386 kJ/mol.",
        "*D. 3 và 945 kJ/mol*"
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Phân tử nitrogen có liên kết ba N≡N rất bền vững với bậc liên kết là 3 và năng lượng liên kết rất lớn là 945 kJ/mol, do đó N~2~ trơ về mặt hoá học ở nhiệt độ thường.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 9 (SBT - KNTT). Xét cân bằng hóa học: NH~3~ + H~2~O ⇌ NH~4~^+ + OH^-. Cân bằng sẽ chuyển dịch theo chiều thuận khi cho thêm vài giọt dung dịch nào sau đây?")
    opts_c9 = [
        "A. NH~4~Cl.",
        "B. NaOH.",
        "*C. HCl.*",
        "D. NaCl."
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Dung dịch HCl cung cấp ion H^+, H^+ tác dụng với OH^- tạo H~2~O làm giảm nồng độ OH^-. Theo nguyên lí Le Chatelier, cân bằng chuyển dịch theo chiều tạo thêm OH^- (chiều thuận).")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 10. Tính base của NH~3~ do nguyên nhân nào sau đây?")
    opts_c10 = [
        "*A. Trên N còn cặp electron tự do.*",
        "B. Phân tử có 3 liên kết cộng hóa trị phân cực.",
        "C. NH~3~ tan được nhiều trong nước.",
        "D. NH~3~ tác dụng với nước tạo NH~4~OH."
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trên nguyên tử nitrogen của phân tử NH~3~ còn một cặp electron chưa liên kết (cặp electron tự do), có khả năng nhận proton H^+ tạo liên kết cho - nhận, thể hiện tính base theo thuyết Brønsted – Lowry.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 11. HNO~3~ đặc thể hiện tính oxi hóa khi tác dụng với chất nào dưới đây?")
    opts_c11 = [
        "A. Fe~2~O~3~",
        "*B. Fe(OH)~2~*",
        "C. CuCl~2~",
        "D. Al~2~O~3~"
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Fe(OH)~2~ có nguyên tử Fe ở số oxi hoá +2 (chưa cực đại) nên bị HNO~3~ đặc oxi hoá lên mức +3: Fe(OH)~2~ + 4HNO~3~ ➔ Fe(NO~3~)~3~ + NO~2~ + 3H~2~O. Các chất Fe~2~O~3~ và Al~2~O~3~ đã đạt số oxi hoá cao nhất (+3) nên chỉ xảy ra phản ứng acid - base trao đổi thông thường.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 12. Phát biểu nào sau đây sai về SO~2~?")
    opts_c12 = [
        "A. Chất chống nấm mốc lương thực, thực phẩm.",
        "B. Sản xuất H~2~SO~4~ trong công nghiệp.",
        "*C. Không gây ô nhiễm môi trường.*",
        "D. Làm chất tẩy trắng giấy và bột giấy."
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "SO~2~ là chất khí độc gây ô nhiễm không khí nghiêm trọng và là tác nhân chủ yếu gây mưa acid.")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 13. Nhóm các chất nào sau đây đều tác dụng được với dung dịch H~2~SO~4~ loãng?")
    opts_c13 = [
        "A. Fe, CuO, Cu(OH)~2~, BaCl~2~, NaCl.",
        "B. FeO, Cu, Cu(OH)~2~, BaCl~2~, Na~2~CO~3~.",
        "*C. Fe~2~O~3~, Cu(OH)~2~, Zn, Na~2~SO~3~, Ba(NO~3~)~2~.*",
        "D. Fe(OH)~3~, Ag, CuO, KHCO~3~, MgS."
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nhóm C gồm: Fe~2~O~3~ (basic oxide), Cu(OH)~2~ (base), Zn (kim loại đứng trước H), Na~2~SO~3~ (muối acid yếu giải phóng SO~2~), Ba(NO~3~)~2~ (tạo kết tủa BaSO~4~↓). Nhóm A có NaCl không phản ứng; nhóm B có Cu không phản ứng; nhóm D có Ag không phản ứng.")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 14. Phát biểu nào sau đây sai?")
    opts_c14 = [
        "A. SO~2~ dùng làm chất tẩy trắng giấy và bột giấy.",
        "B. SO~2~ dùng làm chất chống mốc lương thực, thực phẩm.",
        "*C. Trong phòng thí nghiệm, SO~2~ được điều chế bằng cách đốt quặng pyrit (pyrite).*",
        "D. SO~3~ là chất trung gian để sản xuất sulfuric acid."
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Đốt quặng pyrit (FeS~2~) là phương pháp sản xuất SO~2~ trong công nghiệp. Trong phòng thí nghiệm, SO~2~ được điều chế bằng cách cho muối sulfite tác dụng với acid mạnh: Na~2~SO~3~ + H~2~SO~4~ ➔ Na~2~SO~4~ + SO~2~↑ + H~2~O.")

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 15. Hiện tượng phú dưỡng là gì?")
    opts_c15 = [
        "*A. Quá trình tích lũy quá mức các chất dinh dưỡng trong thủy vực, dẫn đến sự phát triển mạnh của thực vật thủy sinh, đặc biệt là tảo.*",
        "B. Quá trình thiếu hụt các chất dinh dưỡng trong nước, làm giảm năng suất sinh học của hệ sinh thái.",
        "C. Sự suy giảm chất lượng nước do sự gia tăng vi sinh vật gây bệnh.",
        "D. Hiện tượng mưa acid làm giảm độ pH của nước hồ."
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo định nghĩa chuẩn SGK Hóa học 11, hiện tượng phú dưỡng là hiện tượng ao, hồ bị tích tụ quá nhiều chất dinh dưỡng (nitrogen, phosphorus) dẫn đến sự bùng phát của tảo và các thực vật thủy sinh.")

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 16. Mưa acid được hình thành khi nào?")
    opts_c16 = [
        "A. Khi có sự bốc hơi của nước từ sông hồ.",
        "*B. Khi các khí như SO~2~ và NO~x~ phản ứng với nước trong khí quyển tạo thành sulfuric acid và nitric acid.*",
        "C. Khi các hạt bụi trong không khí kết hợp với nước mưa.",
        "D. Khi carbon dioxide (CO~2~) trong không khí phản ứng với nước tạo thành carbonic acid."
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Mưa acid sinh ra khi khí quyển bị ô nhiễm bởi các oxide acid SO~2~ và NO~x~ từ khí thải nhà máy, phương tiện giao thông. Các khí này bị oxi hoá và hoà tan vào nước mưa tạo thành H~2~SO~4~ và HNO~3~.")

    # Câu 17
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 17. Hậu quả chính của mưa acid đối với môi trường là gì?")
    opts_c17 = [
        "A. Làm tăng độ pH của nước sông, hồ, gây hại cho hệ sinh thái nước ngọt.",
        "B. Làm tăng lượng chất dinh dưỡng trong đất, giúp cây cối phát triển.",
        "*C. Làm giảm độ pH của đất và nước, gây thiệt hại cho cây cối, thủy sinh vật và các tòa nhà.*",
        "D. Làm tăng sự phát triển của vi khuẩn trong môi trường đất."
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Mưa acid làm giảm độ pH của đất (chua đất) và nước, phá hủy thảm thực vật rừng, giết chết sinh vật thủy sinh và hòa tan ăn mòn kim loại, đá vôi của các công trình xây dựng.")

    # Câu 18
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 18. Cho các phát biểu sau:\n(1) Đối với môi trường, khí sulfur dioxide là nguyên nhân quan trọng gây mưa acid.\n(2) Chuyển hoá sulfur dioxide thành các chất ít gây ô nhiễm hơn bằng các hóa chất như vôi sống (thành phần chính là CaO), vôi tôi (thành phần chính là Ca(OH)~2~) hoặc đá vôi nghiền (thành phần chính là CaCO~3~).\n(3) Khí sulfur dioxide có màu vàng, ít tan trong nước, thể hiện cả tính oxi hoá và tính khử.\n(4) Đối với con người, khí sulfur dioxide gây ra cảm giác khó thở, bệnh viêm đường hô hấp, đau mắt.\n(5) Đơn chất N~2~ khó hóa lỏng và tan nhiều trong nước.\n(6) Tương tác van der Waals giữa các phân tử nitrogen yếu nên khó hóa lỏng.\nSố phát biểu đúng là")
    opts_c18 = [
        "A. 3.",
        "*B. 4.*",
        "C. 5.",
        "D. 6."
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Có 4 phát biểu đúng gồm: (1), (2), (4), (6).\n- (3) Sai vì SO~2~ là khí không màu, tan nhiều trong nước.\n- (5) Sai vì N~2~ rất ít tan trong nước (ở điều kiện thường 1 lít nước chỉ tan được khoảng 15 mL N~2~).")

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
    render_question_prompt(p, "Câu 1. Trong phương trình hóa học của phản ứng giữa nitrogen và oxygen:")
    stmts_p2_c1 = [
        ("a) Phương trình hóa học: N~2~ + O~2~ ⇌ 2NO (ở 3000°C hoặc tia lửa điện).", "ĐÚNG"),
        ("b) Oxygen giảm từ 0 xuống -1 ➔ oxygen đóng vai trò chất oxi hóa trong phản ứng này.", "SAI"),
        ("c) Nitrogen tăng từ 0 lên +4 ➔ nitrogen đóng vai trò là chất khử trong phản ứng này.", "SAI"),
        ("d) Phản ứng này xảy ra ở nhiệt độ cao trên 3000°C hoặc có tia lửa điện nhưng hiệu suất tạo NO rất thấp nên thực tế không dùng để tạo NO.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c1 = (
            "a) Đúng, ở nhiệt độ cao khoảng 3000°C hoặc khi có tia lửa điện xảy ra phản ứng thuận nghịch giữa N~2~ và O~2~.\n"
            "b) Sai vì số oxi hoá của oxygen giảm từ 0 xuống -2 trong hợp chất NO.\n"
            "c) Sai vì số oxi hoá của nitrogen tăng từ 0 lên +2 trong hợp chất NO.\n"
            "d) Đúng vì phản ứng thu nhiệt mạnh (Δ~r~H°~298~ = +180 kJ/mol) và có hằng số cân bằng rất nhỏ nên hiệu suất rất thấp."
        )
        add_solution_no_table(doc, sol_p2_c1)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Mưa acid gây tác động xấu đến môi trường, con người và sinh vật:")
    stmts_p2_c2 = [
        ("a) Khi nước mưa có pH nhỏ hơn 5,6 thì gọi là hiện tượng mưa acid.", "ĐÚNG"),
        ("b) Hai tác nhân chính gây mưa acid là Cl~2~, HCl.", "SAI"),
        ("c) Mưa acid ăn mòn công trình xây dựng, kiến trúc bằng đá, cẩm thạch và kim loại.", "ĐÚNG"),
        ("d) Mưa acid làm giảm độ pH của đất và nguồn nước sông hồ, ảnh hưởng nghiêm trọng đến đời sống thủy sinh.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c2 = (
            "a) Đúng theo quy chuẩn phân loại hiện tượng mưa acid (pH < 5,6).\n"
            "b) Sai vì hai tác nhân chính gây mưa acid là SO~2~ và các oxide của nitrogen (NO~x~).\n"
            "c) Đúng vì H~2~SO~4~ và HNO~3~ trong mưa acid tác dụng với CaCO~3~ trong đá và các kết cấu kim loại.\n"
            "d) Đúng vì acid hóa nguồn nước và đất canh tác là tác hại môi trường lớn nhất của mưa acid."
        )
        add_solution_no_table(doc, sol_p2_c2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Thành phần chính của khí thiên nhiên là các hydrocarbon như methane (khoảng 80 - 85%), ethane, propane, butane cùng lượng nhỏ các khí carbon dioxide, hydrogen sulfide, nitrogen. Thành phần chính của than là carbon, ngoài ra còn có một số hợp chất của các nguyên tố H, S, O, N,...")
    stmts_p2_c3 = [
        ("a) Khi sử dụng khí thiên nhiên hoặc than làm nhiên liệu đều thải vào không khí các chất khí gây ô nhiễm môi trường không khí.", "ĐÚNG"),
        ("b) Khí thiên nhiên và than khi cháy sẽ toả rất nhiều nhiệt, tạo điều kiện cho các hợp chất như hydrogen sulfide và các hợp chất chứa S, N bị oxi hoá sinh ra các khí độc gây ô nhiễm như SO~2~, NO, NO~2~,...", "ĐÚNG"),
        ("c) Các khí SO~2~, NO~2~ này là nguyên nhân chính gây ra hiệu ứng nhà kính.", "SAI"),
        ("d) Khí thiên nhiên và than đá khi cháy sẽ tạo ra sản phẩm chính là khí CO~2~ gây mưa acid.", "SAI")
    ]
    for stmt, ans in stmts_p2_c3:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c3 = (
            "a) Đúng vì quá trình cháy không bao giờ hoàn toàn và luôn tạo kèm các khí thải gây ô nhiễm.\n"
            "b) Đúng, nhiệt lượng cao thúc đẩy sự tạo thành các oxide độc hại của lưu huỳnh và nitơ.\n"
            "c) Sai vì SO~2~ và NO~2~ là tác nhân gây mưa acid, còn chất gây hiệu ứng nhà kính chủ yếu là CO~2~ và CH~4~.\n"
            "d) Sai vì CO~2~ tạo acid H~2~CO~3~ rất yếu không được coi là nguyên nhân gây mưa acid mà là khí nhà kính chính."
        )
        add_solution_no_table(doc, sol_p2_c3)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Khi pha loãng dung dịch H~2~SO~4~ đặc:")
    stmts_p2_c4 = [
        ("a) Sulfuric acid tan tốt trong nước, quá trình hoà tan toả ra một lượng nhiệt lớn.", "ĐÚNG"),
        ("b) Sulfuric acid tan tốt trong nước, quá trình hoà tan thu một lượng nhiệt lớn.", "SAI"),
        ("c) Để đảm bảo an toàn phải cho từ từ nước vào acid và khuấy đều.", "SAI"),
        ("d) Để đảm bảo an toàn phải cho từ từ acid vào nước và khuấy đều.", "ĐÚNG")
    ]
    for stmt, ans in stmts_p2_c4:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        sol_p2_c4 = (
            "a) Đúng vì sự hydrate hóa phân tử H~2~SO~4~ giải phóng lượng nhiệt khổng lồ.\n"
            "b) Sai vì quá trình hòa tan tỏa nhiệt cực mạnh chứ không thu nhiệt.\n"
            "c) Sai vì rót nước vào acid đặc sẽ làm nước sôi đột ngột và bắn các giọt acid nóng ra ngoài gây bỏng nghiêm trọng.\n"
            "d) Đúng, quy tắc vàng trong phòng thí nghiệm: rót từ từ acid đặc dọc theo đũa thủy tinh vào lượng nước lớn và khuấy nhẹ."
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
    render_question_prompt(p, "Câu 1. Cho sơ đồ chuyển hóa nitrogen trong khí quyển thành phân đạm:\nN~2~ ➔(+O~2~) NO ➔(+O~2~) NO~2~ ➔(+O~2~ + H~2~O) HNO~3~ ➔ NO~3~^-\nSố phản ứng thuộc loại oxi hóa – khử trong sơ đồ là:")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *3*", base_size=10.5)
        add_solution_no_table(doc, "Có 3 phản ứng oxi hoá – khử (có sự thay đổi số oxi hoá của N):\n(1) N~2~^0 + O~2~ ⇌ 2N^+2^O\n(2) 2N^+2^O + O~2~ ➔ 2N^+4^O~2~\n(3) 4N^+4^O~2~ + O~2~ + 2H~2~O ➔ 4HN^+5^O~3~\nGiai đoạn tạo muối nitrate: HNO~3~ ➔ H^+ + NO~3~^- không thay đổi số oxi hoá (N vẫn là +5).")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Có bao nhiêu ý kiến sau đây về sulfur dioxide (SO~2~) là đúng?\n(1) Có độc tính đối với con người.\n(2) Phản ứng được với đá vôi.\n(3) Khí này được tạo thành từ hoạt động của núi lửa trong tự nhiên, từ quá trình đốt cháy nhiên liệu hóa thạch của con người,...\n(4) Là oxide lưỡng tính.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *3*", base_size=10.5)
        add_solution_no_table(doc, "Có 3 ý kiến đúng gồm: (1), (2), (3).\n- (1) Đúng, SO~2~ kích ứng mắt và đường hô hấp.\n- (2) Đúng, SO~2~ phản ứng với CaCO~3~ (đá vôi) trong công nghệ xử lý khí thải: SO~2~ + CaCO~3~ ➔ CaSO~3~ + CO~2~.\n- (3) Đúng về nguồn gốc tự nhiên và nhân tạo.\n- (4) Sai vì SO~2~ là acidic oxide (oxide acid), không phải oxide lưỡng tính.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Tính thể tích (mL) dung dịch H~2~SO~4~ 98% (khối lượng riêng D = 1,84 g/mL) cần dùng để pha chế thành 500 mL dung dịch H~2~SO~4~ 0,05 M (kết quả làm tròn đến hai chữ số phần thập phân).")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *1,36*", base_size=10.5)
        add_solution_no_table(doc, "n(H~2~SO~4~) = 0,5 . 0,05 = 0,025 mol ➔ m(H~2~SO~4~) = 0,025 . 98 = 2,45 g.\nm(dd H~2~SO~4~ 98%) = 2,45 / 0,98 = 2,5 g.\nV(dd) = m / D = 2,5 / 1,84 ≈ 1,3587 mL ≈ 1,36 mL.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4 (SBT - KNTT). Tính phân tử khối trung bình của không khí, giả thiết thành phần không khí gồm: 78% nitrogen, 21% oxygen và 1% argon (làm tròn kết quả đến phần nguyên).")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *29*", base_size=10.5)
        add_solution_no_table(doc, "M_kk = 28,014 . 0,78 + 31,998 . 0,21 + 39,948 . 0,01 = 21,85 + 6,72 + 0,40 = 28,97 ≈ 29 g/mol.")
    else:
        add_student_writing_lines(doc, num_lines=2)

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Hỗn hợp X gồm N~2~ và H~2~ có khối lượng mol trung bình M_X = 12,4 g/mol. Dẫn X đi qua bình đựng bột Fe rồi nung nóng (hiệu suất tổng hợp NH~3~ đạt 40%), thu được hỗn hợp Y. Giá trị khối lượng mol trung bình M_Y là bao nhiêu g/mol? (làm tròn kết quả đến hàng phần mười).")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *14,8*", base_size=10.5)
        add_solution_no_table(doc, "Quy tắc đường chéo: n(N~2~) / n(H~2~) = (12,4 - 2) / (28 - 12,4) = 10,4 / 15,6 = 2 / 3.\nChọn n(N~2~) = 2 mol, n(H~2~) = 3 mol ➔ n_X = 5 mol, m_X = 2 . 28 + 3 . 2 = 62 g.\nVì n(H~2~)/3 = 1 < n(N~2~)/1 = 2 nên hiệu suất tính theo H~2~.\nH = 40% ➔ n(H~2~ pứ) = 3 . 0,40 = 1,2 mol ➔ n(N~2~ pứ) = 0,4 mol; n(NH~3~) = 0,8 mol.\nSố mol hỗn hợp Y: n_Y = (2 - 0,4) + (3 - 1,2) + 0,8 = 4,2 mol.\nBảo toàn khối lượng: m_Y = m_X = 62 g ➔ M_Y = 62 / 4,2 ≈ 14,76 ≈ 14,8 g/mol.")
    else:
        add_student_writing_lines(doc, num_lines=3)

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Một khu vực nông nghiệp gần hồ có diện tích 10 km^2^, hằng năm lượng nước chảy tràn từ mưa cuốn theo phân bón vào hồ chứa khoảng 5 mg/L nitrate (NO~3~^-). Nếu lượng nước chảy tràn chiếm 2% lượng mưa, và lượng mưa trung bình hàng năm là 800 mm, tính khối lượng nitrate (theo đơn vị kg) bị thải vào hồ mỗi năm.")
    if is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_ans, "➔ *Đáp số:* *800*", base_size=10.5)
        add_solution_no_table(doc, "Diện tích: S = 10 km^2^ = 10^7^ m^2^.\nLượng nước mưa hàng năm: V_mưa = 800 . 10^-3^ m . 10^7^ m^2^ = 8.10^6^ m^3^ = 8.10^9^ L.\nLượng nước chảy tràn vào hồ: V_tràn = 0,02 . 8.10^9^ = 1,6.10^8^ L.\nKhối lượng nitrate thải vào hồ: m(NO~3~^-) = 5 mg/L . 1,6.10^8^ L = 8.10^8^ mg = 800 kg.")
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
    render_question_prompt(p, "Câu 1. (SBT - CD) Trong quy trình sản xuất tơ, mỗi năm có hàng triệu tấn cyclohexanone (C~6~H~10~O) được cho phản ứng với HNO~3~ để tạo adipic acid (C~6~H~10~O~4~) theo phản ứng:\nC~6~H~10~O + HNO~3~ ➔ C~6~H~10~O~4~ + N~2~O + H~2~O\na) Cân bằng phương trình hoá học của phản ứng trên theo phương pháp thăng bằng electron.\nb) Cho biết vai trò của HNO~3~ trong phản ứng trên. Giải thích.")
    if is_teacher:
        sol_p4_c1 = (
            "a) Cân bằng theo phương pháp thăng bằng electron:\n"
            "   C~6~H~10~O (C có số oxi hóa trung bình -4/3) ➔ C~6~H~10~O~4~ (C có số oxi hóa trung bình -1/3):\n"
            "   6C^-4/3^ ➔ 6C^-1/3^ + 6e  (hoặc C~6~H~10~O + 3H~2~O ➔ C~6~H~10~O~4~ + 6H^+ + 6e)   | x 4\n"
            "   2N^+5^ + 8e ➔ 2N^+1^ (trong N~2~O)                                                 | x 3\n"
            "   Phương trình hoàn chỉnh:\n"
            "   4C~6~H~10~O + 6HNO~3~ ➔ 4C~6~H~10~O~4~ + 3N~2~O + 3H~2~O\n"
            "b) Trong phản ứng trên, HNO~3~ đóng vai trò là chất oxi hóa vì số oxi hóa của nguyên tử nitrogen giảm từ +5 (trong HNO~3~) xuống +1 (trong N~2~O)."
        )
        add_solution_no_table(doc, sol_p4_c1)
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2 (SBT - CTST). Muối NH~4~NO~3~ sẽ nhiệt phân theo phản ứng nào trong 2 phản ứng sau? Giải thích.\n(1) NH~4~NO~3~(s) ➔[t°] NH~3~(g) + HNO~3~(g)\n(2) NH~4~NO~3~(s) ➔[t°] N~2~O(g) + 2H~2~O(g)\nBiết enthalpy tạo thành chuẩn Δ~f~H°~298~ (kJ/mol) của các chất như sau:\n- NH~4~NO~3~(s): -365,61 kJ/mol\n- NH~3~(g): -45,90 kJ/mol\n- N~2~O(g): 82,05 kJ/mol\n- HNO~3~(g): -134,31 kJ/mol\n- H~2~O(g): -241,82 kJ/mol")
    if is_teacher:
        sol_p4_c2 = (
            "Tính biến thiên enthalpy chuẩn của hai phản ứng:\n"
            "- Phản ứng (1):\n"
            "  Δ~r~H°~298~(1) = Δ~f~H°~298~(NH~3~) + Δ~f~H°~298~(HNO~3~) - Δ~f~H°~298~(NH~4~NO~3~)\n"
            "  = (-45,90) + (-134,31) - (-365,61) = +185,40 kJ > 0 (phản ứng thu nhiệt).\n"
            "- Phản ứng (2):\n"
            "  Δ~r~H°~298~(2) = Δ~f~H°~298~(N~2~O) + 2.Δ~f~H°~298~(H~2~O) - Δ~f~H°~298~(NH~4~NO~3~)\n"
            "  = 82,05 + 2.(-241,82) - (-365,61) = -35,97 kJ < 0 (phản ứng toả nhiệt).\n"
            "Kết luận: Vì Δ~r~H°~298~(2) < 0 nên phản ứng (2) là phản ứng toả nhiệt, thuận lợi hơn nhiều về mặt năng lượng, do đó muối NH~4~NO~3~ sẽ nhiệt phân theo phản ứng (2) tạo khí cười N~2~O và hơi nước."
        )
        add_solution_no_table(doc, sol_p4_c2)
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Lưu file
    filename = "De_Kiem_Tra_Hoa_11_De04_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_11_De04.docx"
    filepath = os.path.join(OUT_DIR, filename)
    doc.save(filepath)
    print(f"Saved: {filepath}")

if __name__ == "__main__":
    print("Generating De_Kiem_Tra_Hoa_11_De04.docx (Học Sinh)...")
    build_exam_document(is_teacher=False)
    print("Generating De_Kiem_Tra_Hoa_11_De04_GiaoVien.docx (Giáo Viên)...")
    build_exam_document(is_teacher=True)
    print("Done!")
