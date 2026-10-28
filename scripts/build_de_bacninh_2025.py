# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Hóa Học 12 - Cụm trường THPT Bắc Ninh
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_Hoa_12_BacNinh_2025_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_Hoa_12_BacNinh_2025_GiaoVien.docx
"""

import os
import sys

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Import thư viện chuẩn
sys.path.append(os.path.abspath(r"C:\Antigravity_Thanh\scripts"))
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
    COLOR_PRIMARY,
    COLOR_CORRECT,
    COLOR_SOLUTION
)

IMG_DIR = r"C:\Antigravity_Thanh\San_Pham\images_bacninh_2025"

def add_centered_image(doc, img_path, width_cm=8.0):
    """Chèn hình ảnh căn giữa văn bản"""
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run()
        run.add_picture(img_path, width=Cm(width_cm))

def build_exam_bacninh(is_teacher=False, output_path=None):
    doc = docx.Document()
    setup_header_footer(doc)

    # 1. Khối tiêu đề
    build_title_block(
        doc,
        school_line="SỞ GD&ĐT BẮC NINH - CỤM TRƯỜNG THPT",
        exam_title="HƯỚNG TỚI KÌ THI TỐT NGHIỆP THPT",
        subject_line="MÔN: HOÁ HỌC - LỚP 12",
        num_pages="04",
        duration="50 phút",
        exam_code="101",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; P = 31; K = 39; Cl = 35,5; Fe = 56; Ag = 108",
        is_teacher=is_teacher
    )

    # =========================================================================
    # PHẦN I
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN I. Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.",
        subtitle="Mỗi câu trả lời đúng thí sinh được 0,25 điểm"
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Các đồng vị của cùng một nguyên tố hoá học là những nguyên tử có cùng số proton nhưng khác nhau ...(1).... Nội dung phù hợp trong chỗ trống (1) là:")
    opts1 = [
        "A. số electron.",
        "B. điện tích hạt nhân.",
        "*C. số neutron.*" if is_teacher else "C. số neutron.",
        "D. số hiệu nguyên tử."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Theo định nghĩa, đồng vị là những nguyên tử của cùng một nguyên tố hoá học, có cùng số proton (do đó cùng điện tích hạt nhân và số hiệu nguyên tử) nhưng khác nhau về số neutron, dẫn đến số khối khác nhau.\n➔ Đáp án đúng: C.")

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Dụng cụ trong hình bên dưới có tên gọi là gì?")
    add_centered_image(doc, os.path.join(IMG_DIR, "cau2_ong_sinh_han.png"), width_cm=7.5)
    opts2 = [
        "A. Pipette.",
        "*B. Ống sinh hàn.*" if is_teacher else "B. Ống sinh hàn.",
        "C. Ống đong.",
        "D. Burette."
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Dụng cụ trong hình là ống sinh hàn (thường dùng để làm lạnh và ngưng tụ hơi chất lỏng trong phương pháp chưng cất).\n➔ Đáp án đúng: B.")

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Cho cấu hình electron nguyên tử của các nguyên tố: X: 1s^2^2s^2^2p^6^; Y: 1s^2^2s^2^2p^6^3s^2^; Z: 1s^2^2s^2^2p^3^; T: 1s^2^2s^2^2p^6^3s^2^3p^5^. Nhận xét nào sau đây đúng?")
    opts3 = [
        "A. X là khí hiếm, Z là kim loại.",
        "B. Tất cả các nguyên tố trên đều thuộc cùng một chu kì trong bảng tuần hoàn.",
        "C. Trong 4 nguyên tố trên, chỉ có Z là phi kim.",
        "*D. Liên kết trong hợp chất tạo bởi Y và T là liên kết ion.*" if is_teacher else "D. Liên kết trong hợp chất tạo bởi Y và T là liên kết ion."
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "• X có Z = 10 (Ne, khí hiếm, chu kì 2).\n• Y có Z = 12 (Mg, kim loại điển hình, chu kì 3).\n• Z có Z = 7 (N, phi kim, chu kì 2).\n• T có Z = 17 (Cl, phi kim điển hình, chu kì 3).\n- A sai vì Z là phi kim (Nitrogen).\n- B sai vì X, Z thuộc chu kì 2 còn Y, T thuộc chu kì 3.\n- C sai vì cả Z (N) và T (Cl) đều là phi kim.\n- D đúng vì Y (Mg, kim loại mạnh) kết hợp với T (Cl, phi kim mạnh) tạo hợp chất MgCl~2~ chứa liên kết ion.\n➔ Đáp án đúng: D.")

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Trong máu người trưởng thành, khỏe mạnh vào lúc đói có một lượng nhỏ glucose với nồng độ khoảng 4,4 – 7,2 mmol/L (hay 80 – 130 mg/dL). Công thức cấu tạo dạng mạch hở của glucose là")
    opts4 = [
        "*A. CH~2~OH[CHOH]~4~CHO.*" if is_teacher else "A. CH~2~OH[CHOH]~4~CHO.",
        "B. CH~2~OH[CHOH]~5~CHO.",
        "C. CH~2~OH[CHOH]~4~CH~2~OH.",
        "D. CH~2~OH[CHOH]~4~COOH."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Glucose là hợp chất tạp chức, dạng mạch hở là một polyhydroxy aldehyde có 6 carbon: CH~2~OH[CHOH]~4~CHO (gồm 5 nhóm -OH kề nhau và 1 nhóm -CHO).\n➔ Đáp án đúng: A.")

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 5. Liên kết sigma (σ) là liên kết được hình thành do")
    opts5 = [
        "A. lực hút tĩnh điện giữa hai ion.",
        "B. sự xen phủ bên của 2 orbital.",
        "*C. sự xen phủ trục của hai orbital.*" if is_teacher else "C. sự xen phủ trục của hai orbital.",
        "D. sự nhường và nhận electron giữa 2 nguyên tử."
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Liên kết sigma (σ) là liên kết cộng hóa trị được hình thành do sự xen phủ trục (xen phủ dọc theo trục liên kết nối hai hạt nhân) giữa hai orbital nguyên tử.\n➔ Đáp án đúng: C.")

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 6. Cho phản ứng: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g), ΔH < 0. Yếu tố nào sau đây không làm chuyển dịch cân bằng của phản ứng trên?")
    opts6 = [
        "A. Áp suất chung của hệ.",
        "*B. Chất xúc tác.*" if is_teacher else "B. Chất xúc tác.",
        "C. Nồng độ các chất phản ứng.",
        "D. Nhiệt độ của hệ."
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Chất xúc tác làm tăng đồng thời tốc độ phản ứng thuận và tốc độ phản ứng nghịch với cùng một tỉ lệ, giúp hệ nhanh chóng đạt đến trạng thái cân bằng chứ không làm chuyển dịch cân bằng hoá học.\n➔ Đáp án đúng: B.")

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 7. Amine nào sau đây là amine bậc hai?")
    opts7 = [
        "A. Methylamine.",
        "*B. Dimethylamine.*" if is_teacher else "B. Dimethylamine.",
        "C. Ethylamine.",
        "D. Trimethylamine."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "• Methylamine (CH~3~NH~2~) và ethylamine (C~2~H~5~NH~2~) là amine bậc một.\n• Dimethylamine ((CH~3~)~2~NH) là amine bậc hai.\n• Trimethylamine ((CH~3~)~3~N) là amine bậc ba.\n➔ Đáp án đúng: B.")

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 8. Thuỷ phân hoàn toàn saccharose (C~12~H~22~O~11~) trong môi trường acid, thu được sản phẩm nào?")
    opts8 = [
        "*A. Glucose và fructose.*" if is_teacher else "A. Glucose và fructose.",
        "B. Chỉ có fructose.",
        "C. Chỉ có glucose.",
        "D. Glucose và maltose."
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Saccharose cấu tạo từ 1 gốc α-glucose và 1 gốc β-fructose qua liên kết α-1,2-glycosidic. Khi thuỷ phân hoàn toàn trong môi trường acid đun nóng tạo ra 1 phân tử glucose và 1 phân tử fructose.\n➔ Đáp án đúng: A.")

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 9. Nhỏ vài giọt nước bromine vào ống nghiệm chứa aniline, hiện tượng quan sát được là")
    opts9 = [
        "*A. có kết tủa màu trắng.*" if is_teacher else "A. có kết tủa màu trắng.",
        "B. có bọt khí thoát ra.",
        "C. xuất hiện màu xanh.",
        "D. xuất hiện màu tím."
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Nhóm -NH~2~ trên vòng benzen định hướng thế vào các vị trí ortho và para, tạo kết tủa trắng 2,4,6-tribromoaniline:\nC~6~H~5~NH~2~ + 3Br~2~ ➔ C~6~H~2~Br~3~NH~2~↓ (trắng) + 3HBr.\n➔ Đáp án đúng: A.")

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 10. Để phân biệt ba bình khí không màu, mất nhãn: propane, ethylene và acetylene, người ta có thể dùng lần lượt các thuốc thử nào sau đây?")
    opts10 = [
        "*A. Dung dịch AgNO~3~/NH~3~ và dung dịch Br~2~.*" if is_teacher else "A. Dung dịch AgNO~3~/NH~3~ và dung dịch Br~2~.",
        "B. Dung dịch KMnO~4~ và dung dịch NaOH.",
        "C. Dung dịch Br~2~ và que đóm cháy dở.",
        "D. Quỳ tím ẩm và dung dịch Br~2~."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "• Dẫn lần lượt từng khí qua dung dịch AgNO~3~/NH~3~: Khí tạo kết tủa vàng nhạt (CAg≡CAg↓) là acetylene.\n• Hai khí còn lại dẫn qua dung dịch Br~2~: Khí làm nhạt màu dung dịch Br~2~ là ethylene (C~2~H~4~); khí không làm mất màu Br~2~ là propane (C~3~H~8~).\n➔ Đáp án đúng: A.")

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 11. Acid béo omega-3 và omega-6 là các acid béo không no với liên kết đôi đầu tiên ở vị trí số 3 và 6 khi đánh số từ nhóm methyl, chúng đều có vai trò quan trọng đối với cơ thể, giúp phòng ngừa nhiều loại bệnh. Chất nào sau đây thuộc loại acid béo omega-3?")
    add_centered_image(doc, os.path.join(IMG_DIR, "cau11_acid_beo.png"), width_cm=14.0)
    opts11 = [
        "*A. Chất A.*" if is_teacher else "A. Chất A.",
        "B. Chất B.",
        "C. Chất C.",
        "D. Chất D."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Đánh số carbon bắt đầu từ nguyên tử carbon của nhóm methyl (đầu đối diện với nhóm -COOH):\n• C~1~(CH~3~) – C~2~ – C~3~ = C~4~: liên kết đôi đầu tiên xuất hiện ở vị trí C số 3, do đó chất ở phương án A thuộc loại acid béo omega-3 (cụ thể là ALA - linolenic acid).\n• Chất B là acid béo omega-9 (oleic acid, liên kết đôi tại C9).\n• Chất D là acid béo omega-6 (linoleic acid, liên kết đôi tại C6).\n➔ Đáp án đúng: A.")

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 12. Isoamyl acetate là một ester có mùi thơm dễ chịu của chuối chín, được dùng làm hương liệu thực phẩm. Công thức cấu tạo thu gọn của isoamyl acetate là")
    opts12 = [
        "A. HCOOCH~2~CH~2~CH(CH~3~)~2~.",
        "*B. CH~3~COOCH~2~CH~2~CH(CH~3~)~2~.*" if is_teacher else "B. CH~3~COOCH~2~CH~2~CH(CH~3~)~2~.",
        "C. CH~3~COOCH(CH~3~)CH~2~CH~3~.",
        "D. CH~3~COOCH~2~CH(CH~3~)~2~."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Isoamyl acetate được tạo thành từ acetic acid (CH~3~COOH) và isoamyl alcohol ((CH~3~)~2~CHCH~2~CH~2~OH):\nCông thức: CH~3~COOCH~2~CH~2~CH(CH~3~)~2~.\n➔ Đáp án đúng: B.")

    # Câu 13
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 13. Sodium lauryl sulfate là một trong những chất giặt rửa tổng hợp phổ biến có công thức cấu tạo như sau:")
    add_centered_image(doc, os.path.join(IMG_DIR, "cau13_sls.png"), width_cm=7.5)
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(1)
    p_sub.paragraph_format.space_after = Pt(1)
    render_rich_text(p_sub, "Đặc điểm của nhóm được khoanh tròn trong công thức trên là gì?")
    opts13 = [
        "A. Tạo kết tủa với nước cứng.",
        "B. Không phân cực.",
        "C. Kị nước.",
        "*D. Ưa nước.*" if is_teacher else "D. Ưa nước."
    ]
    render_options_no_table(doc, opts13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Chất giặt rửa tổng hợp gồm 2 phần:\n• Phần đuôi hydrocarbon dài không phân cực có tính kị nước.\n• Nhóm ion phân cực –OSO~3~^- Na^+ (được khoanh tròn) mang điện tích, có khả năng tạo liên kết ion - lưỡng cực mạnh với các phân tử nước nên có tính ưa nước (đầu ưa nước).\n➔ Đáp án đúng: D.")

    # Câu 14
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 14. Để khảo sát ảnh hưởng của cấu tạo gốc hydrocarbon đến lực base của các amine no, đơn chức, mạch không nhánh, một nhóm học sinh tiến hành đo pH của các dung dịch amine gồm: methylamine, ethylamine và propylamine. Trong kế hoạch thí nghiệm, nhóm đã liệt kê các yếu tố sau:\n(1) Nồng độ mol của các dung dịch amine.\n(2) Nhiệt độ của các dung dịch khi đo pH.\n(3) Loại amine được sử dụng trong mỗi mẫu đo.\n(4) Thể tích dung dịch amine lấy ra để đo pH.\nNhững yếu tố nào cần được thiết lập giống nhau ở tất cả các mẫu thử?")
    opts14 = [
        "A. (1), (2), (3).",
        "B. (1), (3), (4).",
        "C. (2), (3), (4).",
        "*D. (1), (2), (4).*" if is_teacher else "D. (1), (2), (4)."
    ]
    render_options_no_table(doc, opts14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Trong thí nghiệm này:\n• Yếu tố cần khảo sát (biến độc lập) là loại amine (3) (gốc methyl, ethyl, propyl khác nhau).\n• Các biến kiểm soát (cần giữ cố định, giống nhau giữa các mẫu thử để đảm bảo tính khách quan của phép đo) là: nồng độ mol của dung dịch amine (1), nhiệt độ đo (2) và thể tích mẫu dung dịch lấy ra đo pH (4).\n➔ Đáp án đúng: D.")

    # Câu 15
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 15. Chất nào sau đây thuộc loại hợp chất alcohol?")
    opts15 = [
        "*A. C~3~H~5~(OH)~3~.*" if is_teacher else "A. C~3~H~5~(OH)~3~.",
        "B. CH~3~COOH.",
        "C. H-CHO.",
        "D. CH~3~–CO–CH~3~."
    ]
    render_options_no_table(doc, opts15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "• C~3~H~5~(OH)~3~ (glycerol) là một polyalcohol (chứa 3 nhóm -OH liên kết trực tiếp với 3 carbon no).\n• CH~3~COOH là carboxylic acid.\n• H-CHO là aldehyde.\n• CH~3~–CO–CH~3~ là ketone.\n➔ Đáp án đúng: A.")

    # Câu 16
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 16. Chất X có màu trắng, dạng sợi, không mùi vị, không tan trong nước, là thành phần chính tạo nên màng tế bào thực vật. Chất X là")
    opts16 = [
        "A. saccharose.",
        "*B. cellulose.*" if is_teacher else "B. cellulose.",
        "C. tinh bột.",
        "D. glucose."
    ]
    render_options_no_table(doc, opts16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Cellulose là chất rắn màu trắng, dạng sợi, không mùi vị, không tan trong nước và là thành phần cấu tạo nên thành tế bào thực vật, tạo nên bộ khung của cây cối.\n➔ Đáp án đúng: B.")

    # Câu 17
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 17. Ở điều kiện thường, thực hiện thí nghiệm với khí NH~3~ như sau:\nBước 1. Nạp đầy khí NH~3~ vào bình thủy tinh rồi đậy bình bằng nút cao su.\nBước 2. Dùng ống thủy tinh vuốt nhọn đầu nhúng vào nước, xuyên ống thủy tinh qua nút cao su rồi lắp bình thủy tinh lên giá như hình vẽ bên.")
    add_centered_image(doc, os.path.join(IMG_DIR, "cau17_dai_phun_nuoc.png"), width_cm=5.0)
    p_sub17 = doc.add_paragraph()
    p_sub17.paragraph_format.space_before = Pt(1)
    p_sub17.paragraph_format.space_after = Pt(1)
    render_rich_text(p_sub17, "Phát biểu nào sau đây là không đúng?")
    opts17 = [
        "A. Trong phản ứng NH~3~ + H~2~O ⇌ NH~4~^+ + OH^-. Theo thuyết Brønsted – Lowry, NH~3~ đóng vai trò là base.",
        "B. Thí nghiệm trên chứng minh khí NH~3~ tan tốt trong nước.",
        "*C. Tia nước phun mạnh vào bình do áp suất trong bình cao hơn áp suất không khí.*" if is_teacher else "C. Tia nước phun mạnh vào bình do áp suất trong bình cao hơn áp suất không khí.",
        "D. Dung dịch NH~3~ có môi trường base."
    ]
    render_options_no_table(doc, opts17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Khi một lượng nhỏ nước xâm nhập vào bình, khí NH~3~ tan rất nhiều trong nước làm áp suất khí trong bình GIẢM MẠNH (áp suất trong bình nhỏ hơn áp suất khí quyển bên ngoài). Do chênh lệch áp suất, áp suất khí quyển đẩy nước từ cốc phun mạnh qua ống vuốt nhọn vào bình tạo thành tia nước. Vì vậy C là phát biểu không đúng.\n➔ Đáp án đúng: C.")

    # Câu 18
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 18. Cho phản ứng tổng hợp ammonia: N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g). Cho năng lượng liên kết E~b~ ở điều kiện chuẩn: E~b~(N≡N) = 945 kJ/mol; E~b~(H–H) = 436 kJ/mol; E~b~(N–H) = 386 kJ/mol. Biến thiên enthalpy chuẩn (Δ~r~H°~298~) của phản ứng là:")
    opts18 = [
        "A. -935 kJ.",
        "B. -92 kJ.",
        "C. 1095 kJ.",
        "*D. -63 kJ.*" if is_teacher else "D. -63 kJ."
    ]
    render_options_no_table(doc, opts18, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(doc, "Biến thiên enthalpy chuẩn tính theo năng lượng liên kết:\nΔ~r~H°~298~ = ΣE~b~(chất đầu) - ΣE~b~(sản phẩm)\n= [1·E~b~(N≡N) + 3·E~b~(H–H)] - [2·3·E~b~(N–H)]\n= (945 + 3·436) - (6·386) = (945 + 1308) - 2316 = 2253 - 2316 = -63 kJ.\n➔ Đáp án đúng: D.")

    # =========================================================================
    # PHẦN II
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN II. Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.",
        subtitle="Điểm tối đa cho 01 câu hỏi là 1,0 điểm (Đúng 1 ý: 0,1đ; 2 ý: 0,25đ; 3 ý: 0,5đ; 4 ý: 1,0đ)"
    )

    # Câu 1 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. Cellulose là thành phần chính của thành tế bào thực vật, làm nên bộ khung của cây cối. Cellulose có nhiều trong bông, đay, tre, nứa, gỗ,... Từ cellulose người ta điều chế được cellulose acetate (có thành phần chủ yếu là cellulose triacetate và cellulose diacetate) theo các phương trình sau:\n(1) [C~6~H~7~O~2~(OH)~3~]~n~ + 2n(CH~3~CO)~2~O ➔ [C~6~H~7~O~2~(OCOCH~3~)~2~(OH)]~n~ + 2nCH~3~COOH.\n(2) [C~6~H~7~O~2~(OH)~3~]~n~ + 3n(CH~3~CO)~2~O ➔ [C~6~H~7~O~2~(OCOCH~3~)~3~]~n~ + 3nCH~3~COOH.")
    
    stmts1 = [
        ("a) Trong mỗi phân tử cellulose có 3 nhóm hydroxy nên cellulose có khả năng phản ứng với Cu(OH)~2~ trong môi trường kiềm ở nhiệt độ thường tạo dung dịch màu xanh lam.", "Sai"),
        ("b) Phân tử khối của cellulose là 162.", "Sai"),
        ("c) Cellulose là nguyên liệu trong sản xuất ethanol và thuốc súng không khói.", "Đúng"),
        ("d) Khi phân tích thành phần nguyên tố của một mẫu cellulose acetate người ta thu được phần trăm khối lượng carbon là 49,875%. Vậy thành phần phần trăm khối lượng oxygen trong cellulose acetate trên là 34,64%. Giả sử trong cellulose acetate chỉ có cellulose triacetate và cellulose diacetate. (kết quả làm tròn đến hàng phần trăm).", "Sai")
    ]
    for stmt, ans in stmts1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(doc, "a) Sai: Mặc dù mỗi mắt xích có 3 nhóm -OH nhưng các nhóm này không nằm ở vị trí kề nhau có thể tạo phức chelate bền với Cu(OH)~2~ ở nhiệt độ thường, hơn nữa cấu trúc mạch thẳng rất dài kết tụ thành bó sợi không tan nên không có phản ứng tạo phức xanh lam.\nb) Sai: 162 là khối lượng của một mắt xích C~6~H~10~O~5~. Phân tử khối của cellulose là 162n (n rất lớn, từ 1.000 đến 14.000).\nc) Đúng: Cellulose thuỷ phân tạo glucose rồi lên men thành ethanol; tác dụng với HNO~3~/H~2~SO~4~ đặc tạo cellulose trinitrate là thuốc súng không khói.\nd) Sai: Xét 1 mol mắt xích gồm x mol diacetate [C~6~H~7~O~2~(OCOCH~3~)~2~(OH)] (M = 246 g/mol, chứa 10 C, 14 H, 7 O) và y mol triacetate [C~6~H~7~O~2~(OCOCH~3~)~3~] (M = 288 g/mol, chứa 12 C, 16 H, 8 O) với x + y = 1.\nTa có: %m~C~ = 12·(10x + 12y) / (246x + 288y) = 0,49875 ➔ x = 0,1179; y = 0,8821.\nKhối lượng trung bình của mắt xích M~tb~ = 246·0,1179 + 288·0,8821 = 283,05 g/mol.\nSố mol O trong 1 mol mắt xích n~O~ = 7x + 8y = 7·0,1179 + 8·0,8821 = 7,8821 mol.\n➔ %m~O~ = (16 · 7,8821 / 283,05) · 100% = 44,56% (khác 34,64%).")

    # Câu 2 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Cho các phát biểu sau:")
    stmts2 = [
        ("a) Trong không khí, N~2~ chiếm khoảng 78% về thể tích.", "Đúng"),
        ("b) Trong phản ứng giữa N~2~ và H~2~ thì N~2~ vừa là chất oxi hóa, vừa là chất khử.", "Sai"),
        ("c) Khi điều chế khí nitrogen, cách thu khí ở hình 1 là phù hợp nhất.", "Sai"),
        ("d) Phần lớn N~2~ được sử dụng để tổng hợp NH~3~ từ đó sản xuất nitric acid, phân bón, ...", "Đúng")
    ]
    p_st = doc.add_paragraph()
    p_st.paragraph_format.left_indent = Cm(0.5)
    render_statement(p_st, stmts2[0][0], ans=stmts2[0][1], is_teacher=is_teacher)

    p_st = doc.add_paragraph()
    p_st.paragraph_format.left_indent = Cm(0.5)
    render_statement(p_st, stmts2[1][0], ans=stmts2[1][1], is_teacher=is_teacher)

    p_st = doc.add_paragraph()
    p_st.paragraph_format.left_indent = Cm(0.5)
    render_statement(p_st, stmts2[2][0], ans=stmts2[2][1], is_teacher=is_teacher)

    add_centered_image(doc, os.path.join(IMG_DIR, "p2_cau2_thu_khi.png"), width_cm=9.0)

    p_st = doc.add_paragraph()
    p_st.paragraph_format.left_indent = Cm(0.5)
    render_statement(p_st, stmts2[3][0], ans=stmts2[3][1], is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(doc, "a) Đúng: Trong khí quyển, khí nitrogen chiếm khoảng 78% thể tích (khoảng 4/5 thể tích không khí).\nb) Sai: Trong phản ứng N~2~ + 3H~2~ ⇌ 2NH~3~, số oxi hoá của N giảm từ 0 xuống -3 nên N~2~ chỉ đóng vai trò chất oxi hoá.\nc) Sai: Khí nitrogen có M = 28 g/mol, rất xấp xỉ khối lượng mol trung bình của không khí (29 g/mol) nên không thể thu bằng phương pháp dời không khí (úp hay ngửa bình). Do N~2~ rất ít tan trong nước, phương pháp thích hợp nhất để thu N~2~ tinh khiết là đẩy nước (hình 3).\nd) Đúng: Phần lớn lượng khí N~2~ công nghiệp được dùng để tổng hợp NH~3~, nguyên liệu then chốt sản xuất nitric acid, phân đạm và các hợp chất khác.")

    # Câu 3 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Aspirin (hay acetylsalicylic acid), là một dẫn xuất của salicylic acid được sử dụng để hạ sốt và giảm đau nhẹ đến trung bình do tình trạng đau nhức cơ, răng, cảm lạnh đau đầu và sưng tấy do viêm khớp. Một nhóm học sinh thực hiện tổng hợp Aspirin trong phòng thí nghiệm như sau:\nBước 1: Cho 75,0 gam salicylic acid khan và 105,0 mL acetic anhydride (d = 1,08 g/mL) vào bình cầu 250 mL; thêm 4,0 mL sulfuric acid đặc vào và lắc kỹ. Sau đó, khuấy khối phản ứng ở 50–60°C trong khoảng 45 phút cho đến khi tan hết phần chất rắn.\nBước 2: Dùng nước đá để làm lạnh hỗn hợp phản ứng đến nhiệt độ nhỏ hơn 10°C. Cho hỗn hợp vào cốc lớn, thêm từ từ 750 mL nước cất và khuấy kỹ, aspirin sẽ kết tủa. Lọc lấy sản phẩm bằng phễu lọc Buchner.\nBước 3: Hòa tan aspirin thô trong 100 mL ethanol 90°, đun nóng để tan hoàn toàn, sau đó đổ dung dịch này từ từ vào 450 mL nước nóng khoảng 50°C, đun nóng nhẹ cho chất rắn tan hết. Để nguội dung dịch thu được đến nhiệt độ phòng. Aspirin kết tinh dưới dạng tinh thể, lọc và hút hết nước, sấy khô ở 50°C.\nBước 4: Cân sản phẩm thu được 83,1 gam aspirin. Hiệu suất của quá trình là x%.\nBiết phương trình hóa học điều chế aspirin từ salicylic acid được biểu diễn như sau:")
    add_centered_image(doc, os.path.join(IMG_DIR, "p2_cau3_aspirin.png"), width_cm=11.0)
    
    stmts3 = [
        ("a) Giá trị của x là 85% (kết quả làm tròn đến hàng đơn vị).", "Đúng"),
        ("b) Aspirin tan trong nước nhiều hơn salicylic acid.", "Sai"),
        ("c) Aspirin là hợp chất hữu cơ tạp chức, trong đó có chức ester và chức phenol.", "Sai"),
        ("d) Phương pháp tách và tinh chế aspirin ở trên là phương pháp kết tinh.", "Đúng")
    ]
    for stmt, ans in stmts3:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(doc, "a) Đúng: Ta có:\n• Salicylic acid: M = 138 g/mol; n~salicylic~ = 75,0 / 138 ≈ 0,5435 mol.\n• Acetic anhydride: m = 105,0 · 1,08 = 113,4 gam; M = 102 g/mol ➔ n = 113,4 / 102 ≈ 1,1118 mol (dư).\n➔ Lượng aspirin lý thuyết tạo thành: n~lt~ = 0,5435 mol ➔ m~lt~ = 0,5435 · 180 = 97,83 gam.\n➔ Hiệu suất x% = (83,1 / 97,83) · 100% = 84,94% ≈ 85%.\nb) Sai: Aspirin có nhóm -OH phenol bị acyl hoá thành nhóm ester ít phân cực hơn, độ tan trong nước nhỏ hơn nhiều so với salicylic acid.\nc) Sai: Aspirin có cấu tạo CH~3~COOC~6~H~4~COOH, gồm chức ester (-OCOCH~3~) và chức carboxylic acid (-COOH), không còn nhóm -OH liên kết trực tiếp với vòng benzen nên không còn chức phenol.\nd) Đúng: Quy trình hoà tan trong hỗn hợp cồn - nước nóng rồi làm nguội chậm để tinh thể lắng xuống và lọc lấy tinh thể là phương pháp kết tinh lại.")

    # Câu 4 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Trong quản lý chất lượng dầu mỡ công nghiệp và thực phẩm, chỉ số acid được định nghĩa là số miligam KOH cần dùng để trung hòa hết lượng acid béo tự do có trong 1,00 gam chất béo; chỉ số xà phòng hóa là số miligam KOH cần dùng để xà phòng hóa hoàn toàn các triglyceride và trung hòa hết lượng acid béo tự do có trong 1,00 gam chất béo. Để kiểm định một lô dầu cọ tinh luyện (mẫu X) gồm các triglyceride và một lượng nhỏ acid béo tự do (giả thiết các gốc hydrocarbon mạch dài trong X chỉ gồm palmitate – C~15~H~31~, stearate – C~17~H~35~ và oleate – C~17~H~33~), phòng thí nghiệm tiến hành ba thí nghiệm độc lập sau:\n- Thí nghiệm 1: Hòa tan 10,00 gam mẫu X vào hỗn hợp dung môi trung tính (ethanol – ether) rồi tiến hành chuẩn độ bằng dung dịch chuẩn KOH 0,10 M. Đến điểm kết thúc chuẩn độ, thể tích dung dịch KOH đã tiêu tốn là 4,50 mL.\n- Thí nghiệm 2: Đun nóng 50,00 gam mẫu X với 150,00 mL dung dịch KOH 1,50 M trong cồn đến khi phản ứng xà phòng hóa xảy ra hoàn toàn. Để trung hòa lượng KOH còn dư sau phản ứng cần vừa đủ 51,00 mL dung dịch chuẩn HCl 1,00 M.\n- Thí nghiệm 3: Tiến hành phản ứng hydrogen hóa hoàn toàn 50,00 gam mẫu X (xúc tác Ni, t°, p), thấy thể tích khí H~2~ phản ứng tối đa là 1,2395 lít (đkc).\nCho các phát biểu sau:")
    
    stmts4 = [
        ("a) Ở Thí nghiệm 3, phản ứng hydrogen hóa làm tăng số liên kết pi trong phân tử các triglyceride, từ đó giúp dầu cọ chuyển từ dạng lỏng sang dạng rắn và hạn chế bị ôi thiu.", "Sai"),
        ("b) Trong Thí nghiệm 1, dung môi được sử dụng là hỗn hợp ethanol – ether với mục đích hòa tan hoàn toàn mẫu dầu mà không làm thay đổi lượng acid béo tự do cần chuẩn độ.", "Đúng"),
        ("c) Chỉ số xà phòng hóa của mẫu dầu cọ X luôn có giá trị lớn hơn chỉ số acid của nó.", "Đúng"),
        ("d) Từ các số liệu thực nghiệm, xác định được chỉ số acid và chỉ số xà phòng hóa của mẫu dầu cọ X lần lượt là 2,52 và 194,88.", "Đúng")
    ]
    for stmt, ans in stmts4:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(doc, "a) Sai: Phản ứng hydrogen hóa cộng H~2~ vào liên kết C=C (liên kết pi), làm GIẢM số liên kết pi trong phân tử triglyceride, chuyển chất béo lỏng (không no) thành chất béo rắn (no).\nb) Đúng: Hỗn hợp ethanol - ether là dung môi hữu cơ trung tính hoà tan tốt chất béo mà không phản ứng với acid béo tự do, không làm thay đổi lượng acid béo tự do cần chuẩn độ.\nc) Đúng: Chỉ số xà phòng hoá = lượng KOH tác dụng với cả triglyceride và acid béo tự do; chỉ số acid = lượng KOH chỉ tác dụng với acid béo tự do. Do mẫu dầu cọ chứa triglyceride nên chỉ số xà phòng hoá luôn lớn hơn chỉ số acid.\nd) Đúng: \n• Thí nghiệm 1: n~KOH~ = 0,10 · 0,0045 = 0,45 · 10^-3 mol. Khối lượng KOH = 0,45 · 10^-3 · 56 · 10^3 = 25,2 mg trong 10,00 gam chất béo. ➔ Chỉ số acid = 25,2 / 10,00 = 2,52.\n• Thí nghiệm 2: n~KOH ban đầu~ = 0,150 · 1,50 = 0,225 mol; n~KOH dư~ = n~HCl~ = 0,051 · 1,00 = 0,051 mol.\n➔ n~KOH phản ứng~ = 0,225 - 0,051 = 0,174 mol trong 50,00 gam chất béo.\n➔ m~KOH phản ứng~ = 0,174 · 56 · 10^3 = 9744 mg.\n➔ Chỉ số xà phòng hoá = 9744 / 50,00 = 194,88.")

    # =========================================================================
    # PHẦN III
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN III. Thí sinh trả lời từ câu 1 đến câu 6.",
        subtitle="Mỗi câu trả lời đúng thí sinh được 0,25 điểm"
    )

    # Câu 1 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. Cho các phát biểu sau:\n(a) Phân tử methyl methacrylate có một liên kết pi trong phân tử.\n(b) Cho dung dịch methylamine vào ống nghiệm đựng dung dịch iron(III) chloride thấy xuất hiện kết tủa nâu đỏ.\n(c) Methylamine làm quỳ tím ẩm hóa xanh.\n(d) Glucose và maltose đều có nhóm -OH hemiacetal.\n(e) Dung dịch Fructose và dung dịch saccharose đều hòa tan được Cu(OH)~2~, tạo thành dung dịch ở điều kiện thường.\nCó bao nhiêu phát biểu đúng?")
    
    if not is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* ...................................................................................................................................................")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* *4*", base_color=COLOR_CORRECT)
        add_solution_no_table(doc, "(a) Sai: Methyl methacrylate có cấu tạo CH~2~=C(CH~3~)COOCH~3~, gồm 1 liên kết đôi C=C (1 liên kết π) và 1 liên kết đôi C=O trong nhóm ester (1 liên kết π) ➔ có tổng cộng 2 liên kết π.\n(b) Đúng: 3CH~3~NH~2~ + FeCl~3~ + 3H~2~O ➔ Fe(OH)~3~↓ (nâu đỏ) + 3CH~3~NH~3~Cl.\n(c) Đúng: Methylamine tan tốt trong nước tạo môi trường kiềm yếu, làm quỳ tím ẩm hoá xanh.\n(d) Đúng: Dạng mạch vòng của glucose có -OH hemiacetal tự do; maltose tạo bởi 2 gốc glucose nhưng liên kết qua nhóm -OH hemiacetal của gốc thứ nhất với C4 của gốc thứ hai nên gốc glucose thứ hai vẫn giữ nguyên nhóm -OH hemiacetal tự do.\n(e) Đúng: Cả fructose và saccharose đều là polyalcohol có nhiều nhóm -OH liền kề nên đều hoà tan Cu(OH)~2~ tạo dung dịch phức màu xanh lam ở điều kiện thường.\n➔ Có 4 phát biểu đúng: (b), (c), (d), (e).\n➔ Đáp số: 4.")

    # Câu 2 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Xăng sinh học E5 (chứa 5% ethanol về thể tích, còn lại là xăng, giả thiết chỉ là octane). Khi được đốt cháy hoàn toàn, 1 mol ethanol tỏa ra lượng nhiệt là 1365,0 kJ và 1 mol isooctane (C~8~H~18~) tỏa ra lượng nhiệt là 5928,7 kJ. Trung bình, một chiếc xe máy di chuyển được 1 km thì cần một nhiệt lượng chuyển thành công cơ học có độ lớn là 212 kJ. Để đi được quãng đường 100 km cần dùng bao nhiêu lít xăng E5? (Không làm tròn kết quả các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng phần trăm). Biết rằng, hiệu suất sử dụng nhiên liệu của động cơ là 25%, khối lượng riêng của ethanol là 0,8 g/mL, của isooctane là 0,7 g/mL.")
    
    if not is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* ...................................................................................................................................................")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* *2,37*", base_color=COLOR_CORRECT)
        add_solution_no_table(doc, "• Công cơ học cần thiết để xe đi 100 km: A = 100 · 212 = 21 200 kJ.\n• Hiệu suất sử dụng nhiệt lượng của động cơ là H = 25% = 0,25.\n➔ Tổng nhiệt lượng cần toả ra từ việc đốt cháy xăng E5:\n  Q~cần~ = A / H = 21 200 / 0,25 = 84 800 kJ.\n• Xét 1 lít (1000 mL) xăng E5:\n  - Thể tích ethanol: V~et~ = 5% · 1000 = 50 mL.\n    Khối lượng ethanol: m~et~ = 50 · 0,8 = 40 g ➔ n~et~ = 40 / 46 mol.\n    Nhiệt toả ra từ ethanol: Q~et~ = (40 / 46) · 1365,0 ≈ 1186,9565 kJ.\n  - Thể tích isooctane: V~oct~ = 95% · 1000 = 950 mL.\n    Khối lượng isooctane: m~oct~ = 950 · 0,7 = 665 g ➔ n~oct~ = 665 / 114 mol.\n    Nhiệt toả ra từ isooctane: Q~oct~ = (665 / 114) · 5928,7 ≈ 34584,0833 kJ.\n  - Tổng nhiệt toả ra khi đốt 1 lít xăng E5:\n    Q~1L~ = Q~et~ + Q~oct~ ≈ 1186,9565 + 34584,0833 = 35771,0398 kJ/L.\n• Thể tích xăng E5 cần dùng:\n  V = Q~cần~ / Q~1L~ = 84 800 / 35771,0398 ≈ 2,3706 L.\nLàm tròn đến hàng phần trăm: 2,37 L.\n➔ Đáp số: 2,37.")

    # Câu 3 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Cho phương trình hóa học của các phản ứng được đánh số thứ tự từ 1 tới 6 dưới đây:\n(1) 6nCO~2~ + 5nH~2~O ➔ [chlorophyll, ánh sáng] ➔ (C~6~H~10~O~5~)~n~ + 6nO~2~.\n(2) C~6~H~12~O~6~ (glucose) ➔ [enzyme] ➔ 2CH~3~-CHOH-COOH (lactic acid).\n(3) (C~6~H~10~O~5~)~n~ (tinh bột) + nH~2~O ➔ [H^+, t°] ➔ nC~6~H~12~O~6~ (glucose).\n(4) C~6~H~12~O~6~ (glucose) ➔ [enzyme] ➔ 2C~2~H~5~OH + 2CO~2~.\n(5) (C~17~H~35~COO)~3~C~3~H~5~ (tristearin) + 3NaOH ➔ [t°] ➔ 3C~17~H~35~COONa + C~3~H~5~(OH)~3~.\n(6) C~6~H~12~O~6~ (glucose) + 2[Ag(NH~3~)~2~]OH ➔ [t°] ➔ CH~2~OH(CHOH)~4~COONH~4~ + 2Ag + 3NH~3~ + H~2~O.\nGán số thứ tự phương trình hoá học của các phản ứng theo tên gọi: xà phòng hóa, quang hợp, lên men rượu, tráng gương và sắp xếp theo trình tự thành dãy bốn số (Ví dụ: 1234, 4321,...)")
    
    if not is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* ...................................................................................................................................................")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* *5146*", base_color=COLOR_CORRECT)
        add_solution_no_table(doc, "Gán số thứ tự các phản ứng theo tên gọi:\n• Phản ứng xà phòng hoá: (5) (thuỷ phân tristearin trong dung dịch kiềm NaOH).\n• Phản ứng quang hợp: (1) (tổng hợp tinh bột từ CO~2~ và H~2~O nhờ diệp lục).\n• Phản ứng lên men rượu: (4) (lên men glucose tạo C~2~H~5~OH và CO~2~).\n• Phản ứng tráng gương: (6) (phản ứng khử thuốc thử Tollens tạo kết tủa kim loại Ag).\n➔ Sắp xếp theo trình tự: 5 - 1 - 4 - 6.\n➔ Dãy bốn số: 5146.")

    # Câu 4 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Phân bón NPK hỗn hợp được sản xuất từ các muối gồm: diammonium phosphate ((NH~4~)~2~HPO~4~), ammonium nitrate (NH~4~NO~3~), potassium chloride (KCl), cùng một lượng nhỏ tạp chất và chất độn không chứa các nguyên tố N, P, K. Thành phần dinh dưỡng quy đổi về % khối lượng các chất/nguyên tố trong một loại phân bón NPK được cho trong bảng sau:")
    
    # Bảng dinh dưỡng biểu diễn bằng Paragraph chuẩn (Tab stops: 4.5, 9.0, 13.5 cm)
    p_tb1 = doc.add_paragraph()
    p_tb1.paragraph_format.left_indent = Cm(1.0)
    p_tb1.paragraph_format.tab_stops.add_tab_stop(Cm(6.0), WD_TAB_ALIGNMENT.LEFT)
    p_tb1.paragraph_format.tab_stops.add_tab_stop(Cm(10.0), WD_TAB_ALIGNMENT.LEFT)
    p_tb1.paragraph_format.tab_stops.add_tab_stop(Cm(14.0), WD_TAB_ALIGNMENT.LEFT)
    render_rich_text(p_tb1, "*Thành phần quy đổi:* \t*N* \t*P~2~O~5~* \t*K~2~O* \t*Thành phần khác*", base_bold=True)

    p_tb2 = doc.add_paragraph()
    p_tb2.paragraph_format.left_indent = Cm(1.0)
    p_tb2.paragraph_format.tab_stops.add_tab_stop(Cm(6.0), WD_TAB_ALIGNMENT.LEFT)
    p_tb2.paragraph_format.tab_stops.add_tab_stop(Cm(10.0), WD_TAB_ALIGNMENT.LEFT)
    p_tb2.paragraph_format.tab_stops.add_tab_stop(Cm(14.0), WD_TAB_ALIGNMENT.LEFT)
    render_rich_text(p_tb2, "*% khối lượng:* \t17,00% \t14,20% \t18,80% \t50,00%")

    p_sub4 = doc.add_paragraph()
    p_sub4.paragraph_format.space_before = Pt(2)
    p_sub4.paragraph_format.space_after = Pt(2)
    render_rich_text(p_sub4, "Phần trăm theo khối lượng của hợp chất ammonium nitrate (NH~4~NO~3~) trong loại phân bón NPK trên là x %.\nXác định giá trị của x (không làm tròn kết quả các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng phần mười).")

    if not is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* ...................................................................................................................................................")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* *32,6*", base_color=COLOR_CORRECT)
        add_solution_no_table(doc, "Xét trong 100,0 gam mẫu phân bón NPK:\n• m~P2O5~ = 14,20 gam ➔ n~P2O5~ = 14,20 / 142 = 0,10 mol.\n  Bảo toàn nguyên tố P: n~P~ = 2·n~P2O5~ = 0,20 mol.\n  Toàn bộ P nằm trong (NH~4~)~2~HPO~4~ (DAP) ➔ n~DAP~ = 0,20 mol.\n  Số mol N có trong DAP: n~N(DAP)~ = 2 · 0,20 = 0,40 mol.\n• Khối lượng N tổng cộng trong 100 g phân bón là m~N~ = 17,00 gam:\n  ➔ n~N tổng~ = 17,00 / 14 = 1,214286 mol.\n• Nguồn cung cấp N gồm (NH~4~)~2~HPO~4~ và NH~4~NO~3~ (mỗi phân tử chứa 2 N):\n  ➔ n~N(NH4NO3)~ = n~N tổng~ - n~N(DAP)~ = 1,214286 - 0,40 = 0,814286 mol.\n  ➔ n~NH4NO3~ = 0,814286 / 2 = 0,407143 mol.\n• Khối lượng NH~4~NO~3~ trong 100 gam phân bón (M = 80 g/mol):\n  m~NH4NO3~ = 0,407143 · 80 = 32,5714 gam.\n• Phần trăm khối lượng của NH~4~NO~3~ trong phân bón:\n  x% = (32,5714 / 100) · 100% = 32,5714% ≈ 32,6%.\n➔ Đáp số: 32,6.")

    # Câu 5 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 5. Cho các chất sau: triolein, saccharose, amylopectin, methyl acetate, fructose, cellulose. Có bao nhiêu chất bị thủy phân trong môi trường acid?")
    
    if not is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* ...................................................................................................................................................")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* *5*", base_color=COLOR_CORRECT)
        add_solution_no_table(doc, "Các chất bị thuỷ phân trong môi trường acid gồm:\n(1) Triolein (chất béo, ester ba chức).\n(2) Saccharose (disaccharide).\n(3) Amylopectin (polysaccharide của tinh bột).\n(4) Methyl acetate (ester đơn chức).\n(5) Cellulose (polysaccharide).\nFructose là monosaccharide đơn giản nhất nên không bị thuỷ phân.\n➔ Có tất cả 5 chất bị thuỷ phân.\n➔ Đáp số: 5.")

    # Câu 6 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 6. Theo Tiêu chuẩn Việt Nam TCVN 7624 : 2007, khi chế tạo gương, chiều dày lớp bạc phủ trên bề mặt tấm kính (quy ra tổng lượng bạc trên một đơn vị m^2^ kính) phải đạt tối thiểu 0,7 g·m^-2^. Một công ty cần sản xuất 18000 m^2^ gương có độ dày lớp bạc phủ ở mức 0,72 g·m^-2^. Để tạo ra bạc, người ta tiến hành theo sơ đồ phản ứng như sau:\nSaccharose ➔ [+H~2~O, xt H^+, t°] ➔ Dung dịch X ➔ [trung hòa acid] ➔ Dung dịch Y ➔ [+dd [Ag(NH~3~)~2~]OH, t°] ➔ Ag.\nBiết hiệu suất cả quá trình là 80%. Để sản xuất lượng gương với độ phủ bạc như trên, công ty đó cần sử dụng lượng saccharose ít nhất là bao nhiêu kg? (kết quả làm tròn đến chữ số phần chục, chỉ làm tròn ở phép tính cuối).")
    
    if not is_teacher:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* ...................................................................................................................................................")
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ans, "*Trả lời:* *12,8*", base_color=COLOR_CORRECT)
        add_solution_no_table(doc, "• Tổng lượng bạc cần phủ lên bề mặt gương:\n  m~Ag~ = 18 000 m^2 · 0,72 g/m^2 = 12 960 gam = 12,96 kg.\n  Số mol Ag cần tạo ra: n~Ag~ = 12 960 / 108 = 120 mol.\n• Sơ đồ phản ứng thuỷ phân và tráng bạc:\n  C~12~H~22~O~11~ (saccharose) + H~2~O ➔ C~6~H~12~O~6~ (glucose) + C~6~H~12~O~6~ (fructose)\n  Trong môi trường kiềm của thuốc thử Tollens, fructose chuyển hoá thành glucose; cả glucose và fructose đều tham gia phản ứng tráng bạc với tỉ lệ 1 mol monosaccharide ➔ 2 mol Ag.\n  ➔ 1 mol C~12~H~22~O~11~ ➔ 1 mol Glu + 1 mol Fruc ➔ 4 mol Ag.\n• Số mol saccharose cần theo lí thuyết:\n  n~sac(lt)~ = n~Ag~ / 4 = 120 / 4 = 30 mol.\n• Do hiệu suất của cả quá trình là H = 80% = 0,80:\n  Lượng saccharose thực tế cần dùng:\n  n~sac(tt)~ = n~sac(lt)~ / 0,80 = 30 / 0,80 = 37,5 mol.\n• Khối lượng saccharose thực tế cần dùng (M = 342 g/mol):\n  m~sac~ = 37,5 · 342 = 12 825 gam = 12,825 kg.\nLàm tròn đến chữ số phần chục (hàng phần mười): 12,8 kg.\n➔ Đáp số: 12,8.")

    # Thêm dòng kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(8)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("-------------------------- HẾT --------------------------")
    set_run_style(r_end, font_size=10.5, bold=True)

    # Lưu file
    doc.save(output_path)
    print(f"Đã tạo thành công: {output_path}")

if __name__ == "__main__":
    out_dir = r"C:\Antigravity_Thanh\San_Pham"
    student_file = os.path.join(out_dir, "De_Kiem_Tra_Hoa_12_BacNinh_2025_HocSinh.docx")
    teacher_file = os.path.join(out_dir, "De_Kiem_Tra_Hoa_12_BacNinh_2025_GiaoVien.docx")

    build_exam_bacninh(is_teacher=False, output_path=student_file)
    build_exam_bacninh(is_teacher=True, output_path=teacher_file)
