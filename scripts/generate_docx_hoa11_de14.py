# -*- coding: utf-8 -*-
"""
Script sinh tự động 2 file Word (.docx) cho ĐỀ SỐ 14 (Chương 5: DẪN XUẤT HALOGEN - ALCOHOL - PHENOL):
1. Bản Học sinh (De_Kiem_Tra_Hoa_11_De14.docx)
2. Bản Giáo viên (De_Kiem_Tra_Hoa_11_De14_GiaoVien.docx)

Tuân thủ nghiêm ngặt Quy chuẩn Form Đề Kiểm Tra (GEMINI.md):
- Không dùng bảng (Table) trong nội dung câu hỏi/đáp án/lời giải.
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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de14")

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
        exam_title="ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 5 (DẪN XUẤT HALOGEN - ALCOHOL - PHENOL)",
        subject_line="MÔN: HOÁ HỌC - LỚP 11",
        num_pages="04",
        duration="50 phút",
        exam_code="114",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; F = 19; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Cr = 52; Mn = 55; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108; I = 127.",
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
    render_question_prompt(p, "Câu 1. Chất nào không phải là dẫn xuất halogen của hydrocarbon?")
    opts_c1 = [
        "A. CH~2~=CH-CH~2~Br.",
        "B. ClBrCH-CF~3~.",
        "*C. Cl~2~CH-CF~2~-O-CH~3~.*",
        "D. C~6~H~6~Cl~6~."
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Dẫn xuất halogen của hydrocarbon chỉ được cấu tạo từ các nguyên tố C, H và halogen (F, Cl, Br, I). "
            "Chất C (Cl~2~CH-CF~2~-O-CH~3~) chứa thêm nguyên tố O nên là một ether chứa halogen, không phải dẫn xuất halogen của hydrocarbon."
        )

    # Câu 2
    p = add_p(doc)
    render_question_prompt(p, "Câu 2. Dẫn xuất halogen nào sau đây không có đồng phân cis-trans?")
    opts_c2 = [
        "A. CHCl=CHCl.",
        "*B. CH~2~=CH-CH~2~F.*",
        "C. CH~3~CH=CBrCH~3~.",
        "D. CH~3~CH~2~CH=CHCHClCH~3~."
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Điều kiện để có đồng phân hình học cis-trans là mỗi nguyên tử carbon mang liên kết đôi phải liên kết với hai nhóm nguyên tử khác nhau. "
            "Chất B có nhóm đầu mạch CH~2~= chứa 2 nguyên tử H giống nhau nên không có đồng phân cis-trans."
        )

    # Câu 3
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 3. Chloromethane, còn được gọi với những cái tên khác là methyl chloride, Refrigerant-40, R-40 hoặc HCC 40, "
        "là một hợp chất hóa học thuộc nhóm haloalkane. Nó từng được sử dụng rộng rãi như một chất làm lạnh. "
        "Công thức phân tử của chloromethane là"
    )
    opts_c3 = [
        "A. CH~2~Cl~2~.",
        "*B. CH~3~Cl.*",
        "C. CHCl~3~.",
        "D. CCl~4~."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Chloromethane (methyl chloride) có công thức phân tử là CH~3~Cl."
        )

    # Câu 4
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 4. Nhiệt độ nóng chảy, nhiệt độ sôi của alcohol đều cao hơn so với hydrocarbon, dẫn xuất halogen, "
        "ether có phân tử lượng tương đương là do"
    )
    opts_c4 = [
        "A. trong phân tử alcohol có liên kết cộng hoá trị.",
        "*B. giữa các phân tử alcohol có liên kết hydrogen liên phân tử.*",
        "C. alcohol có nguyên tử oxygen trong phân tử.",
        "D. alcohol có phản ứng với Na."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Giữa các phân tử alcohol có khả năng tạo liên kết hydrogen liên phân tử bền vững, làm cho các phân tử liên kết chặt chẽ hơn với nhau, "
            "do đó cần nhiều năng lượng hơn để phá vỡ nên nhiệt độ sôi và nhiệt độ nóng chảy của alcohol cao hơn nhiều so với hydrocarbon, dẫn xuất halogen, ether có phân tử lượng tương đương."
        )

    # Câu 5
    p = add_p(doc)
    render_question_prompt(p, "Câu 5. Alcohol nào sau đây khi bị oxi hóa bởi CuO (t°) tạo acetone?")
    opts_c5 = [
        "*A. propan-2-ol.*",
        "B. butan-1-ol.",
        "C. 2-methylpropan-1-ol.",
        "D. propan-1-ol."
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Khi bị oxi hóa không hoàn toàn bởi CuO nung nóng, alcohol bậc II bị oxi hóa tạo ketone. "
            "Propan-2-ol (CH~3~CH(OH)CH~3~) là alcohol bậc II tạo ra acetone (CH~3~COCH~3~):\n"
            "CH~3~CH(OH)CH~3~ + CuO --t°--> CH~3~COCH~3~ + Cu + H~2~O."
        )

    # Câu 6
    p = add_p(doc)
    render_question_prompt(p, "Câu 6. Khi đun nóng dẫn xuất halogen X với dung dịch NaOH tạo thành hợp chất acetic aldehyde. Tên của hợp chất X là")
    opts_c6 = [
        "A. 1,2-dibromoethane.",
        "*B. 1,1-dibromoethane.*",
        "C. ethyl chloride.",
        "D. A và B đúng."
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Phương trình: CH~3~CHBr~2~ + 2NaOH --t°--> CH~3~CHO + 2NaBr + H~2~O.\n"
            "Hợp chất gem-dihalide có hai nguyên tử halogen gắn vào cùng một nguyên tử carbon đầu mạch khi thủy phân sẽ tạo aldehyde. Hợp chất X là 1,1-dibromoethane (CH~3~CHBr~2~)."
        )

    # Câu 7
    p = add_p(doc)
    render_question_prompt(p, "Câu 7. Tên gọi theo danh pháp thay thế của alcohol (CH~3~)~2~CHCH~2~CH~2~OH là")
    opts_c7 = [
        "A. 2-methylbutan-1-ol.",
        "*B. 3-methylbutan-1-ol.*",
        "C. 3-methylbutan-2-ol.",
        "D. 1,1-dimethylpropan-2-ol."
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Đánh số mạch chính từ phía gần nhóm -OH nhất:\n"
            "C~4~H~3~ - C~3~H(CH~3~) - C~2~H~2~ - C~1~H~2~OH.\n"
            "Nhóm methyl ở vị trí C số 3, nhóm -OH ở vị trí C số 1 => Tên thay thế: 3-methylbutan-1-ol."
        )

    # Câu 8
    p = add_p(doc)
    render_question_prompt(p, "Câu 8. Tên thay thế của alcohol có công thức cấu tạo thu gọn CH~3~CH~2~CH~2~OH là")
    opts_c8 = [
        "*A. propan-1-ol.*",
        "B. propan-2-ol.",
        "C. pentan-1-ol.",
        "D. pentan-2-ol."
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Mạch chính có 3 carbon no, nhóm -OH ở vị trí số 1 nên tên thay thế là propan-1-ol."
        )

    # Câu 9
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 9. Catechin là chất chống oxi hoá mạnh có trong lá chè tươi và nhiều loại quả, có khả năng phòng chống ung thư và tim mạch. "
        "Công thức cấu tạo của catechin được cho như hình bên:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "catechin_structure.png"), width_cm=7.0)
    p_c9 = add_p(doc, space_before=2, space_after=3)
    render_rich_text(p_c9, "Phát biểu nào sau đây *không đúng*?")
    opts_c9 = [
        "A. Công thức phân tử của catechin là C~15~H~14~O~6~.",
        "*B. Phân tử catechin có 5 nhóm OH phenol.*",
        "C. Catechin phản ứng được với dung dịch NaOH.",
        "D. Catechin thuộc loại hợp chất thơm."
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Quan sát công thức cấu tạo của catechin:\n"
            "- Có 4 nhóm -OH gắn trực tiếp vào vòng benzene (thuộc loại nhóm OH phenol): 2 nhóm ở vòng A và 2 nhóm ở vòng B.\n"
            "- Có 1 nhóm -OH gắn vào carbon no của dị vòng chứa O (thuộc loại nhóm OH alcohol).\n"
            "Do đó, phân tử catechin chỉ có 4 nhóm OH phenol và 1 nhóm OH alcohol (không phải 5 nhóm OH phenol).\n"
            "Công thức phân tử là C~15~H~14~O~6~. Do có nhóm OH phenol nên phản ứng được với dung dịch NaOH; phân tử chứa vòng benzene nên là hợp chất thơm."
        )

    # Câu 10
    p = add_p(doc)
    render_question_prompt(p, "Câu 10. Trong thực tế, phenol được dùng làm nguyên liệu để sản xuất")
    opts_c10 = [
        "A. nhựa poly(vinyl chloride), nhựa novolac và chất diệt cỏ 2,4-D.",
        "B. nhựa resol, nhựa resit và thuốc trừ sâu 666.",
        "*C. poly(phenol-formaldehyde), chất diệt cỏ 2,4-D và picric acid.*",
        "D. nhựa resit, chất diệt cỏ 2,4-D và thuốc nổ TNT."
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Phenol là nguyên liệu quan trọng trong công nghiệp hóa chất để sản xuất nhựa poly(phenol - formaldehyde) (novolac, resol), "
            "chất diệt cỏ 2,4-D (2,4-dichlorophenoxyacetic acid) và chất nổ picric acid (2,4,6-trinitrophenol)."
        )

    # Câu 11
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 11. Để chứng minh sự ảnh hưởng qua lại của nhóm -OH và vòng benzene trong phenol (C~6~H~5~OH) "
        "thì cần cho phenol tác dụng với các chất nào sau đây?"
    )
    opts_c11 = [
        "A. Na và nước Br~2~.",
        "B. Dung dịch NaOH và khí CO~2~.",
        "*C. Dung dịch NaOH và nước Br~2~.*",
        "D. Quỳ tím và nước Br~2~."
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- Ảnh hưởng của vòng benzene lên nhóm -OH: Làm tăng độ phân cực của liên kết O-H, phenol có tính acid yếu, phản ứng được với dung dịch NaOH (trong khi alcohol không phản ứng).\n"
            "- Ảnh hưởng của nhóm -OH lên vòng benzene: Hoạt hóa các vị trí ortho và para của vòng thơm, làm phenol phản ứng dễ dàng với dung dịch nước bromine (Br~2~) tạo kết tủa trắng 2,4,6-tribromophenol (trong khi benzene không phản ứng ở điều kiện này).\n"
            "Do đó cần dùng dung dịch NaOH và nước Br~2~."
        )

    # Câu 12
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 12. Hiện tượng xảy ra khi nhỏ vài giọt dung dịch HCl đặc vào ống nghiệm chứa một ít dung dịch C~6~H~5~ONa rồi lắc mạnh là"
    )
    opts_c12 = [
        "A. có sự phân lớp; dung dịch trong suốt hóa đục.",
        "*B. dung dịch trong suốt hóa đục.*",
        "C. có phân lớp; dung dịch trong suốt.",
        "D. xuất hiện sự phân lớp ở ống nghiệm."
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Phương trình: C~6~H~5~ONa + HCl -> C~6~H~5~OH + NaCl.\n"
            "Phenol sinh ra rất ít tan trong nước lạnh ở điều kiện thường nên làm cho dung dịch từ trong suốt chuyển sang hóa đục."
        )

    # Câu 13
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 13. Sản phẩm chính tạo thành khi cho 2-bromobutane tác dụng với dung dịch KOH trong ethanol, đun nóng là"
    )
    opts_c13 = [
        "A. methylcyclopropane.",
        "B. but-2-ol.",
        "C. but-1-ene.",
        "*D. but-2-ene.*"
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Theo quy tắc tách Zaitsev, khi tách hydrogen halide khỏi dẫn xuất halogen, nguyên tử halogen bị tách ưu tiên cùng với nguyên tử H ở carbon bậc cao hơn bên cạnh để tạo alkene có nhiều nhóm thế hơn (bền hơn):\n"
            "CH~3~-CHBr-CH~2~-CH~3~ + KOH --(C~2~H~5~OH, t°)--> CH~3~-CH=CH-CH~3~ (sp chính: but-2-ene) + KBr + H~2~O."
        )

    # Câu 14
    p = add_p(doc)
    render_question_prompt(p, "Câu 14. (SBT-KNTT) Sản phẩm chính thu được khi tách nước từ 3-methylbutan-2-ol là")
    opts_c14 = [
        "A. 3-methylbut-1-ene.",
        "*B. 2-methylbut-2-ene.*",
        "C. 3-methylbut-2-ene.",
        "D. 2-methylbut-3-ene."
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Tách nước từ 3-methylbutan-2-ol: CH~3~-CH(CH~3~)-CH(OH)-CH~3~ --(H~2~SO~4~ đặc, t°)-->\n"
            "- Tách H cùng carbon bậc cao hơn (C số 3 là carbon bậc III có 1 H) theo quy tắc Zaitsev tạo sản phẩm chính: CH~3~-C(CH~3~)=CH-CH~3~ (2-methylbut-2-ene).\n"
            "- Tách H ở C số 1 tạo sản phẩm phụ: CH~3~-CH(CH~3~)-CH=CH~2~ (3-methylbut-1-ene)."
        )

    # Câu 15
    p = add_p(doc)
    render_question_prompt(p, "Câu 15. (SBT-KNTT) Oxi hoá alcohol nào sau đây thu được sản phẩm là ketone?")
    opts_c15 = [
        "A. C~2~H~5~OH.",
        "B. CH~3~CH~2~CH~2~OH.",
        "*C. CH~3~CH(OH)CH~3~.*",
        "D. (CH~3~)~2~C(OH)CH~3~."
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Oxi hóa alcohol bậc II bằng CuO nung nóng thu được ketone:\n"
            "CH~3~-CH(OH)-CH~3~ + CuO --t°--> CH~3~-CO-CH~3~ + Cu + H~2~O.\n"
            "Chất C là propan-2-ol (isopropyl alcohol) tạo acetone.\n"
            "(A và B là alcohol bậc I cho aldehyde; D là alcohol bậc III không bị oxi hóa trong điều kiện này)."
        )

    # Câu 16
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 16. Cho các phát biểu sau:\n"
        "(1) Phenol tan vô hạn trong nước ở 66°C.\n"
        "(2) Phenol có lực acid mạnh hơn alcohol ethylic.\n"
        "(3) Phản ứng thế vào benzene dễ hơn phản ứng thế vào nhân thơm của phenol.\n"
        "(4) Phenol tan tốt trong ethanol.\n"
        "(5) Phenol làm quỳ tím hóa đỏ.\n"
        "(6) Nhóm OH phenol không bị thế bởi gốc acid như nhóm OH alcohol.\n"
        "Có bao nhiêu phát biểu đúng?"
    )
    opts_c16 = [
        "A. 5.",
        "*B. 4.*",
        "C. 3.",
        "D. 6."
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "(1) Đúng. Phenol ít tan trong nước lạnh nhưng tan vô hạn trong nước ở nhiệt độ trên 66°C.\n"
            "(2) Đúng. Do vòng benzene hút electron làm tăng độ phân cực của liên kết O-H nên phenol có tính acid mạnh hơn alcohol ethylic.\n"
            "(3) Sai. Nhóm -OH đẩy electron vào vòng benzene làm tăng mật độ electron của vòng thơm nên phản ứng thế vào nhân thơm của phenol dễ hơn nhiều so với benzene.\n"
            "(4) Đúng. Phenol tan tốt trong các dung môi hữu cơ như ethanol, ether,...\n"
            "(5) Sai. Phenol có tính acid rất yếu (yếu hơn carbonic acid H~2~CO~3~) nên không làm đổi màu quỳ tím.\n"
            "(6) Đúng. Nhóm -OH phenol liên kết trực tiếp với vòng thơm có hiệu ứng liên hợp p-pi, liên kết C-O rất bền vững nên nhóm -OH không bị thế bởi gốc acid như trong phản ứng ester hóa thông thường của alcohol (phenol không tác dụng trực tiếp với carboxylic acid).\n"
            "=> Có 4 phát biểu đúng là (1), (2), (4), (6)."
        )

    # Câu 17
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 17. X là hỗn hợp gồm hai alkene (ở thể khí trong điều kiện thường). Hydrate hóa X được hỗn hợp Y gồm 4 alcohol (không có alcohol bậc III). X gồm"
    )
    opts_c17 = [
        "*A. propene và but-1-ene.*",
        "B. ethylene và propene.",
        "C. propene và but-2-ene.",
        "D. propene và 2-methylpropene."
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- Hai alkene ở thể khí trong điều kiện thường: Có số nguyên tử C ≤ 4.\n"
            "- Hỗn hợp Y gồm 4 alcohol và không có alcohol bậc III.\n"
            "Xét các phương án:\n"
            "+ A. Propene (CH~2~=CH-CH~3~) hydrate hóa tạo 2 alcohol: propan-1-ol (bậc I) và propan-2-ol (bậc II).\n"
            "  But-1-ene (CH~2~=CH-CH~2~-CH~3~) hydrate hóa tạo 2 alcohol: butan-1-ol (bậc I) và butan-2-ol (bậc II).\n"
            "  Cả 4 alcohol này đều khác nhau và đều là bậc I hoặc bậc II (không có bậc III). Thỏa mãn.\n"
            "+ B. Ethylene chỉ tạo 1 alcohol (C~2~H~5~OH), propene tạo 2 alcohol => tổng cộng 3 alcohol (loại).\n"
            "+ C. But-2-ene đối xứng chỉ tạo 1 alcohol (butan-2-ol), propene tạo 2 alcohol => tổng cộng 3 alcohol (loại).\n"
            "+ D. 2-methylpropene hydrate hóa tạo 2-methylpropan-2-ol là alcohol bậc III (loại)."
        )

    # Câu 18
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 18. Nghiên cứu số liệu về năng lượng liên kết, độ dài liên kết và độ phân cực carbon - halogen:\n"
        "• Bảng năng lượng và độ dài liên kết carbon - halogen:\n"
        "   - C-F : Năng lượng liên kết 453 kJ/mol; Độ dài liên kết 133 pm\n"
        "   - C-Cl: Năng lượng liên kết 339 kJ/mol; Độ dài liên kết 177 pm\n"
        "   - C-Br: Năng lượng liên kết 276 kJ/mol; Độ dài liên kết 194 pm\n"
        "   - C-I : Năng lượng liên kết 216 kJ/mol; Độ dài liên kết 213 pm"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "cx_polarity.png"), width_cm=11.0)
    p_c18 = add_p(doc, space_before=2, space_after=3)
    render_rich_text(p_c18, "Khả năng phản ứng thế nguyên tử halogen bằng nhóm -OH tăng theo chiều nào sau đây:")
    opts_c18 = [
        "*A. C-F < C-Cl < C-Br < C-I.*",
        "B. C-I < C-Cl < C-Br < C-F.",
        "C. C-F < C-Br < C-Cl < C-I.",
        "D. C-I < C-Br < C-Cl < C-F."
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Từ C-F đến C-I: Năng lượng liên kết C-X giảm dần (từ 453 xuống 216 kJ/mol), độ dài liên kết tăng dần (từ 133 lên 213 pm), "
            "liên kết C-X trở nên kém bền hơn và dễ bị phân cắt hơn. Đồng thời độ phân cực giảm dần làm giảm lực hút tĩnh điện giữa C và X. "
            "Do đó, khả năng phản ứng thế nguyên tử halogen bằng nhóm -OH tăng dần theo chiều: C-F < C-Cl < C-Br < C-I."
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
        "Câu 1. Danh pháp và cấu tạo của dẫn xuất halogen.\n"
        "Trong thực tế sản xuất, methyl chloride (CH~3~Cl) được sử dụng rộng rãi trong công nghiệp. "
        "Hãy xác định đúng/sai cho các mệnh đề sau:"
    )
    stmts_c1_p2 = [
        ("a) Methyl chloride thuộc nhóm dẫn xuất halogen.", "ĐÚNG"),
        ("b) Công thức cấu tạo của methyl chloride là CH~3~Cl, trong đó nguyên tử Cl liên kết với nguyên tử C.", "ĐÚNG"),
        ("c) Theo danh pháp thay thế, CH~3~Cl được gọi là chloroethane.", "SAI"),
        ("d) Methyl chloride là một hợp chất khí dễ bay hơi, có thể gây nguy hiểm khi hít phải trong thời gian dài.", "ĐÚNG")
    ]
    for stmt, ans in stmts_c1_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) ĐÚNG: CH~3~Cl được tạo thành khi thay thế 1 nguyên tử H của methane bằng 1 nguyên tử Cl, thuộc nhóm dẫn xuất halogen.\n"
            "- b) ĐÚNG: Nguyên tử carbon liên kết trực tiếp với 3 nguyên tử H và 1 nguyên tử Cl.\n"
            "- c) SAI: Theo danh pháp thay thế, CH~3~Cl có tên là chloromethane (không phải chloroethane).\n"
            "- d) ĐÚNG: Methyl chloride là chất khí không màu, dễ bay hơi, độc hại khi hít phải trong thời gian dài, gây ức chế thần kinh trung ương."
        )

    # Câu 2 Phần II
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 2. Hợp chất X có tác dụng kháng khuẩn, chống vi sinh vật kí sinh trên da (chấy, rận,...). "
        "X có công thức phân tử C~7~H~8~O và có chứa vòng benzene, phổ IR của X có peak hấp thụ rộng ở vùng 3 300 cm^-1. "
        "Oxi hóa X bằng CuO nung nóng, thu được hợp chất Y có peak hấp thụ đặc trưng ở khoảng 1 700 cm^-1."
    )
    stmts_c2_p2 = [
        ("a) Phổ IR của X có peak hấp thụ rộng ở vùng 3 300 cm^-1 nên X có nhóm C=O.", "SAI"),
        ("b) Phổ IR của Y có peak hấp thụ rộng ở vùng 1 700 cm^-1 nên Y có nhóm -OH.", "SAI"),
        ("c) X là benzyl alcohol, Y là aldehyde benzoic.", "ĐÚNG"),
        ("d) Y tan nhiều trong nước hơn X.", "SAI")
    ]
    for stmt, ans in stmts_c2_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) SAI: Peak hấp thụ rộng ở vùng 3 200 - 3 600 cm^-1 đặc trưng cho liên kết O-H (alcohol hoặc phenol), không phải C=O.\n"
            "- b) SAI: Peak hấp thụ mạnh ở khoảng 1 700 cm^-1 đặc trưng cho nhóm carbonyl C=O (aldehyde hoặc ketone), không phải -OH.\n"
            "- c) ĐÚNG: X chứa vòng benzene, công thức C~7~H~8~O, có nhóm -OH và bị oxi hóa bởi CuO tạo aldehyde Y nên X là alcohol bậc I: benzyl alcohol (C~6~H~5~CH~2~OH), và Y là benzaldehyde (aldehyde benzoic, C~6~H~5~CHO).\n"
            "- d) SAI: Cả X và Y đều chứa gốc phenyl kỵ nước lớn nên ít tan trong nước, nhưng benzyl alcohol (X) tạo được liên kết hydrogen với phân tử nước nên độ tan trong nước của X lớn hơn so với benzaldehyde (Y)."
        )

    # Câu 3 Phần II
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 3. Thực hiện thí nghiệm copper(II) hydroxide tác dụng với alcohol đa chức:\n"
        "Bước 1: Cho vào ống nghiệm 1 mL dung dịch CuSO~4~ 2% và 1 mL dung dịch NaOH 10%.\n"
        "Bước 2: Lắc nhẹ ống nghiệm rồi nhỏ vào 5 giọt glycerol.\n"
        "- Sau bước 1, thấy xuất hiện kết tủa xanh lơ trong ống nghiệm.\n"
        "- Sau bước 2, kết tủa tan thành dung dịch xanh lam đậm."
    )
    stmts_c3_p2 = [
        ("a) Thí nghiệm chứng tỏ glycerol có các nhóm OH liền kề.", "ĐÚNG"),
        ("b) Nếu thay glycerol bằng ethanol, hiện tượng vẫn tương tự.", "SAI"),
        ("c) Có thể dùng thí nghiệm này để nhận biết glycerol và ethanol.", "ĐÚNG"),
        ("d) Không thể dùng Cu(OH)~2~ để nhận biết methanol và ethylene glycol.", "SAI")
    ]
    for stmt, ans in stmts_c3_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) ĐÚNG: Phản ứng hòa tan kết tủa Cu(OH)~2~ tạo dung dịch phức màu xanh lam đậm đặc trưng cho polyalcohol có ít nhất hai nhóm -OH liền kề (như glycerol, ethylene glycol).\n"
            "- b) SAI: Ethanol là alcohol đơn chức, không có khả năng tạo phức hòa tan Cu(OH)~2~, kết tủa xanh không tan.\n"
            "- c) ĐÚNG: Glycerol hòa tan Cu(OH)~2~ tạo dung dịch màu xanh lam đậm, còn ethanol không hòa tan Cu(OH)~2~.\n"
            "- d) SAI: Hoàn toàn có thể dùng Cu(OH)~2~ để nhận biết methanol và ethylene glycol vì ethylene glycol hòa tan Cu(OH)~2~ tạo dung dịch màu xanh lam đậm, còn methanol không có hiện tượng."
        )

    # Câu 4 Phần II
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 4. Eugenol là thành phần chính của tinh dầu hương nhu có công thức phân tử C~10~H~12~O~2~. Eugenol có công thức cấu tạo:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "eugenol_structure.png"), width_cm=5.0)
    p_c4_stmts = add_p(doc, space_before=2, space_after=3)
    render_rich_text(p_c4_stmts, "Cho các phát biểu sau:")
    stmts_c4_p2 = [
        ("a) Eugenol tác dụng được với Na và NaOH theo tỉ lệ mol 1:1.", "ĐÚNG"),
        ("b) Eugenol không có đồng phân hình học.", "ĐÚNG"),
        ("c) Hydrogen hoá hoàn toàn eugenol, thu được sản phẩm hữu cơ có tên gọi là 2-methoxy-4-propylcyclohexanol (nhóm methoxy có công thức là CH~3~O-).", "ĐÚNG"),
        ("d) Eugenol có tính acid mạnh.", "SAI")
    ]
    for stmt, ans in stmts_c4_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) ĐÚNG: Phân tử eugenol chứa 1 nhóm -OH phenol nên phản ứng được với cả Na và dung dịch NaOH theo tỉ lệ mol 1:1. Nhóm ether -OCH~3~ và liên kết đôi C=C không tác dụng với NaOH.\n"
            "- b) ĐÚNG: Nhóm allyl -CH~2~-CH=CH~2~ có carbon nối đôi ở đầu mạch chứa 2 nguyên tử H giống nhau (=CH~2~) nên không có đồng phân hình học cis-trans.\n"
            "- c) ĐÚNG: Khi hydrogen hóa hoàn toàn, liên kết đôi C=C ở nhánh allyl và vòng benzene đều bị hydrogen hóa hoàn toàn thành vòng cyclohexane no và nhóm propyl, sản phẩm thu được là 2-methoxy-4-propylcyclohexanol.\n"
            "- d) SAI: Eugenol chứa nhóm -OH phenol nên có tính acid rất yếu (không làm đổi màu quỳ tím, yếu hơn cả carbonic acid), không phải acid mạnh."
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
    render_question_prompt(p, "Câu 1. Cho công thức khung phân tử của hợp chất X như hình dưới đây:")
    add_centered_picture(doc, os.path.join(IMG_DIR, "propane_br_structure.png"), width_cm=4.5)
    p_q1 = add_p(doc, space_before=2, space_after=3)
    render_rich_text(p_q1, "Nguyên tử Br gắn vào nguyên tử carbon ở vị trí số mấy?")
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Mạch chính có 3 carbon: C~3~H~3~ - C~2~H~2~ - C~1~H~2~ - Br.\n"
            "Đánh số mạch carbon ưu tiên từ đầu gần nhóm thế Br nhất, do đó nguyên tử Br gắn vào carbon ở vị trí số 1 (1-bromopropane).\n"
            "=> Đáp số: 1."
        )

    # Câu 2 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 2. Chloromethane được hình thành từ methane và chlorine theo phản ứng sau:\n"
        "CH~4~ + 3Cl~2~ --asmt--> CHCl~3~ + 3HCl   (1)\n"
        "Cho nhiệt tạo thành chuẩn của một chất tương ứng trong phương trình như sau:\n"
        "• CH~4~(g): -75 kJ/mol\n"
        "• CHCl~3~(g): -103 kJ/mol\n"
        "• HCl(g): -92 kJ/mol\n"
        "Biến thiên enthalpy chuẩn của phản ứng (1) bằng bao nhiêu kJ?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Áp dụng công thức tính biến thiên enthalpy chuẩn từ nhiệt tạo thành:\n"
            "Δ~r~H°~298~ = Δ~f~H°~298~(CHCl~3~(g)) + 3 × Δ~f~H°~298~(HCl(g)) - Δ~f~H°~298~(CH~4~(g))\n"
            "Δ~r~H°~298~ = -103 + 3 × (-92) - (-75) = -103 - 276 + 75 = -304 kJ.\n"
            "=> Đáp số: -304."
        )

    # Câu 3 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 3. Cho các chất sau:\n"
        "(1) Dung dịch HCl đặc;\n"
        "(2) Dung dịch Cu(OH)~2~;\n"
        "(3) Nước bromine;\n"
        "(4) Dung dịch H~2~SO~4~ đặc, nóng;\n"
        "(5) Na;\n"
        "(6) CuO (t°, xt);\n"
        "(7) CH~3~COOH;\n"
        "(8) O~2~ (t°).\n"
        "Có bao nhiêu chất tác dụng được với ethyl alcohol?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Ethyl alcohol (C~2~H~5~OH) tác dụng được với 6 chất:\n"
            "- (1) Dung dịch HCl đặc: C~2~H~5~OH + HCl --t°--> C~2~H~5~Cl + H~2~O.\n"
            "- (4) Dung dịch H~2~SO~4~ đặc, nóng (tách nước tạo C~2~H~4~ hoặc (C~2~H~5~)~2~O).\n"
            "- (5) Na: 2C~2~H~5~OH + 2Na -> 2C~2~H~5~ONa + H~2~.\n"
            "- (6) CuO (t°): C~2~H~5~OH + CuO --t°--> CH~3~CHO + Cu + H~2~O.\n"
            "- (7) CH~3~COOH (phản ứng ester hóa với xúc tác H~2~SO~4~ đặc).\n"
            "- (8) O~2~ (t°): phản ứng cháy tạo CO~2~ và H~2~O (hoặc lên men giấm).\n"
            "Không tác dụng với: (2) Dung dịch Cu(OH)~2~; (3) Nước bromine.\n"
            "=> Đáp số: 6."
        )

    # Câu 4 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 4. Cho dãy các chất/dung dịch sau: NaOH, NaHCO~3~, HCl, NaNO~3~, nước Br~2~ phản ứng với phenol. "
        "Có bao nhiêu chất phản ứng với phenol?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Phenol (C~6~H~5~OH) có tính acid yếu và phản ứng thế nhân thơm:\n"
            "- Tác dụng với dung dịch NaOH: C~6~H~5~OH + NaOH -> C~6~H~5~ONa + H~2~O.\n"
            "- Tác dụng với dung dịch nước bromine (Br~2~): C~6~H~5~OH + 3Br~2~ -> C~6~H~2~Br~3~OH (trắng) + 3HBr.\n"
            "Phenol không phản ứng với NaHCO~3~ (do tính acid yếu hơn H~2~CO~3~), HCl và NaNO~3~.\n"
            "=> Đáp số: 2."
        )

    # Câu 5 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 5. Hợp chất X có tên 3,5-dibromo-8,8-dichloro-2,6-dimethyl-1,6-octadiene là một dẫn xuất halogen tách ra từ biển đỏ chi Laurencia "
        "có tác dụng chống ung thư vòm họng. Biết rằng X có chứa 4 nguyên tố với thành phần khối lượng như sau: "
        "32,88% C; 3,84% H; 43,84% Br; còn lại là chlorine. Biết rằng trong X có chứa hai nguyên tử chlorine. "
        "Số nguyên tử C trong phân tử X là bao nhiêu?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Thành phần phần trăm khối lượng của chlorine trong X:\n"
            "%Cl = 100% - (32,88% + 3,84% + 43,84%) = 19,44%.\n"
            "Vì phân tử X chứa 2 nguyên tử chlorine:\n"
            "M_X = (2 × 35,5) / 19,44% ≈ 365,23 g/mol.\n"
            "Số nguyên tử C trong phân tử X là:\n"
            "n_C = (365,23 × 32,88%) / 12 ≈ 10.\n"
            "=> Đáp số: 10."
        )

    # Câu 6 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 6. Củ sắn khô chứa 38% khối lượng là tinh bột, còn lại là các chất không có khả năng lên men thành ethyl alcohol. "
        "Xăng E5 có 5% thể tích là ethyl alcohol. Dùng toàn bộ lượng ethyl alcohol thu được ở trên để pha chế xăng E5. "
        "Tính thể tích xăng E5 thu được sau khi pha trộn (theo lít), biết khối lượng riêng của ethyl alcohol là 0,8 kg/L, "
        "hiệu suất của cả quá trình lên men từ tinh bột là 81%, lấy 1 tấn củ sắn khô làm nguyên liệu."
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Khối lượng củ sắn khô: m = 1 tấn = 1000 kg.\n"
            "Khối lượng tinh bột: m_tinhbot = 1000 × 38% = 380 kg.\n"
            "Sơ đồ phản ứng lên men:\n"
            "(C~6~H~10~O~5~)~n~ --(+nH~2~O)--> n C~6~H~12~O~6~ --(men)--> 2n C~2~H~5~OH + 2n CO~2~\n"
            "Theo lí thuyết: 162 kg tinh bột tạo 92 kg C~2~H~5~OH.\n"
            "Khối lượng C~2~H~5~OH thực tế thu được với H = 81%:\n"
            "m_ethanol = (380 × 92 / 162) × 81% = 174,8 kg.\n"
            "Thể tích ethyl alcohol thu được:\n"
            "V_ethanol = m / D = 174,8 / 0,8 = 218,5 lít.\n"
            "Vì ethyl alcohol chiếm 5% thể tích xăng E5, nên thể tích xăng E5 thu được là:\n"
            "V_xang = 218,5 / 5% = (218,5 × 100) / 5 = 4370 lít.\n"
            "=> Đáp số: 4370."
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
        "Câu 1. Tìm thông tin thích hợp điền vào chỗ trống trong mỗi phát biểu sau:\n"
        "a) Propane-1,2,3-triol có tên thông thường là ............\n"
        "b) Cho ethane-1,2-diol vào ống nghiệm có Cu(OH)~2~ và dung dịch NaOH, lắc nhẹ, hiện tượng quan sát được là ............\n"
        "c) Đun nóng hỗn hợp gồm ethanol, methanol và H~2~SO~4~ đặc thu được tối đa ....(1)...... ether có công thức cấu tạo là .....(2).....\n"
        "d) Cho a mol alcohol R(OH)~n~ phản ứng với Na (dư) thu được tối đa a mol khí H~2~. Giá trị của n là ......."
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "a) glycerol (hoặc glixerin).\n"
            "b) kết tủa tan, tạo dung dịch màu xanh lam đậm.\n"
            "c) (1) : 3\n"
            "   (2) : CH~3~-O-CH~3~ ; C~2~H~5~-O-C~2~H~5~ ; CH~3~-O-C~2~H~5~.\n"
            "d) Phản ứng: R(OH)~n~ + n Na -> R(ONa)~n~ + (n/2) H~2~.\n"
            "Theo phương trình: n_H~2~ = (n/2) × n_alcohol <=> a = (n/2) × a => n/2 = 1 => n = 2.\n"
            "Giá trị của n là 2."
        )
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Phần IV
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 2. (SBT-CD) Hợp chất A là dẫn xuất monochloro của alkylbenzene (B). Phân tử khối của A bằng 126,5.\n"
        "a) Tìm công thức phân tử và viết công thức cấu tạo có thể có của A.\n"
        "b) Chất A có phản ứng thuỷ phân khi đun nóng với dung dịch NaOH, tạo ra chất E có mùi thơm, có khả năng hoà tan nhiều chất hữu cơ, ức chế sự sinh sản của vi khuẩn nên được dùng nhiều trong công nghiệp sản xuất mỹ phẩm. Tìm công thức cấu tạo đúng của A. Viết phương trình hoá học của phản ứng.\n"
        "c) Viết phương trình hoá học của phản ứng điều chế trực tiếp A từ B, ghi rõ điều kiện của phản ứng."
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "a) A là dẫn xuất monochloro của alkylbenzene nên công thức có dạng C~n~H~2n-7~Cl (n ≥ 7).\n"
            "Phân tử khối: M_A = 14n + 28,5 = 126,5 => 14n = 98 => n = 7.\n"
            "Công thức phân tử của A là C~7~H~7~Cl.\n"
            "Các công thức cấu tạo có thể có của A gồm 4 đồng phân:\n"
            "- (1) 2-chlorotoluene (o-chlorotoluene): o-CH~3~-C~6~H~4~-Cl\n"
            "- (2) 3-chlorotoluene (m-chlorotoluene): m-CH~3~-C~6~H~4~-Cl\n"
            "- (3) 4-chlorotoluene (p-chlorotoluene): p-CH~3~-C~6~H~4~-Cl\n"
            "- (4) benzyl chloride: C~6~H~5~CH~2~Cl.\n"
            "(Xem hình vẽ 4 đồng phân cấu tạo bên dưới)\n"
            "b) Các đồng phân (1), (2), (3) có nguyên tử Cl liên kết trực tiếp với vòng benzene rất bền, không bị thủy phân bởi dung dịch NaOH đun nóng ở điều kiện thường.\n"
            "Chỉ có đồng phân (4) benzyl chloride (C~6~H~5~CH~2~Cl) có nguyên tử Cl ở mạch nhánh dễ dàng bị thủy phân tạo benzyl alcohol (E):\n"
            "C~6~H~5~CH~2~Cl + NaOH --t°--> C~6~H~5~CH~2~OH + NaCl.\n"
            "Benzyl alcohol (E) là chất lỏng có mùi thơm dịu, kháng khuẩn, dùng trong sản xuất mỹ phẩm.\n"
            "Vậy công thức cấu tạo đúng của A là (4): C~6~H~5~CH~2~Cl (benzyl chloride).\n"
            "c) Alkylbenzene B là toluene (C~6~H~5~CH~3~). Điều chế trực tiếp benzyl chloride (A) từ toluene bằng phản ứng clo hóa nhánh với ánh sáng:\n"
            "C~6~H~5~CH~3~ + Cl~2~ --as--> C~6~H~5~CH~2~Cl + HCl."
        )
        add_centered_picture(doc, os.path.join(IMG_DIR, "monochloro_isomers.png"), width_cm=11.0)
    else:
        add_student_writing_lines(doc, num_lines=12)

    return doc

def check_and_save():
    os.makedirs(OUT_DIR, exist_ok=True)

    # 1. Bản Học sinh
    doc_hs = build_exam_document(is_teacher=False)
    tables_hs = len(doc_hs.tables)
    print(f"Bản Học sinh có {tables_hs} bảng (Chỉ được phép có 1 bảng Header block).")
    assert tables_hs <= 1, "LỖI: Bản Học sinh chứa bảng trong nội dung!"
    path_hs = os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_11_De14.docx")
    doc_hs.save(path_hs)
    print("Đã lưu Bản Học sinh:", path_hs)

    # 2. Bản Giáo viên
    doc_gv = build_exam_document(is_teacher=True)
    tables_gv = len(doc_gv.tables)
    print(f"Bản Giáo viên có {tables_gv} bảng (Chỉ được phép có 1 bảng Header block).")
    assert tables_gv <= 1, "LỖI: Bản Giáo viên chứa bảng trong nội dung!"
    path_gv = os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_11_De14_GiaoVien.docx")
    doc_gv.save(path_gv)
    print("Đã lưu Bản Giáo viên:", path_gv)

if __name__ == "__main__":
    check_and_save()
