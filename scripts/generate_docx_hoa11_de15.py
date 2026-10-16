# -*- coding: utf-8 -*-
"""
Script sinh tự động 2 file Word (.docx) cho ĐỀ SỐ 15 (Chương 5: DẪN XUẤT HALOGEN - ALCOHOL - PHENOL):
1. Bản Học sinh (De_Kiem_Tra_Hoa_11_De15.docx)
2. Bản Giáo viên (De_Kiem_Tra_Hoa_11_De15_GiaoVien.docx)

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
IMG_DIR = os.path.join(OUT_DIR, "images_crop_hoa11_de15")

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
        exam_code="115",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; F = 19; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Cr = 52; Mn = 55; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108; Ba = 137.",
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
        "Câu 1. Dichloromethane (DCM) hay methylene chloride (MC) là một chất lỏng không màu, dễ bay hơi với mùi thơm nhẹ. "
        "Nó được sử dụng rộng rãi làm dung môi vì ít độc nhất trong các chlorocarbon. Công thức phân tử của dichloromethane là"
    )
    opts_c1 = [
        "*A. CH~2~Cl~2~.*",
        "B. CH~3~Cl.",
        "C. CHCl~3~.",
        "D. CCl~4~."
    ]
    render_options_no_table(doc, opts_c1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Dichloromethane (methylene chloride) có công thức phân tử là CH~2~Cl~2~."
        )

    # Câu 2
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 2. Từ giữa thế kỷ 18, chloroform chủ yếu được sử dụng làm chất gây mê trong phẫu thuật. "
        "Công thức phân tử của chloroform là"
    )
    opts_c2 = [
        "A. CH~2~Cl~2~.",
        "B. CH~3~Cl.",
        "*C. CHCl~3~.*",
        "D. CCl~4~."
    ]
    render_options_no_table(doc, opts_c2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Chloroform (trichloromethane) có công thức phân tử là CHCl~3~."
        )

    # Câu 3
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 3. Một số dẫn xuất halogen không chứa chlorine như CF~3~CH~2~F (1), CF~3~CH~2~CF~2~CH~3~ (2)... "
        "được sử dụng làm chất làm lạnh thân thiện với môi trường do ít gây hại tầng ozone. Danh pháp thay thế của hợp chất (1) là"
    )
    opts_c3 = [
        "*A. 1,1,1,2-tetrafluoroethane.*",
        "B. 1,1,2,2-tetrafluoroethane.",
        "C. 2,1,1,2-tetrafluoroethane.",
        "D. 1,1,1,2-tetrafluoroethene."
    ]
    render_options_no_table(doc, opts_c3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "CF~3~CH~2~F: Đánh số từ carbon chứa 3 nguyên tử F: C1F~3~ - C2H~2~F.\n"
            "Tên thay thế là 1,1,1,2-tetrafluoroethane."
        )

    # Câu 4
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 4. Độ tan trong nước của các alcohol đều cao hơn so với hydrocarbon, ether có phân tử lượng tương đương là do"
    )
    opts_c4 = [
        "A. các alcohol có liên kết hydrogen với nhau.",
        "B. các alcohol có khối lượng phân tử lớn.",
        "*C. các alcohol tạo được liên kết hydrogen với phân tử nước.*",
        "D. alcohol có liên kết cộng hóa trị với nước."
    ]
    render_options_no_table(doc, opts_c4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Các alcohol tan tốt trong nước là nhờ nhóm -OH có khả năng tạo liên kết hydrogen liên phân tử bền vững với các phân tử nước (R-O-H ... O-H~2~)."
        )

    # Câu 5
    p = add_p(doc)
    render_question_prompt(p, "Câu 5. Sản phẩm sinh ra trong thí nghiệm đun nóng sau là gì?")
    add_centered_picture(doc, os.path.join(IMG_DIR, "cuo_alcohol_experiment.png"), width_cm=4.5)
    opts_c5 = [
        "A. CH~3~CH~2~CHO.",
        "B. CH~3~CH~2~CH~2~OH.",
        "*C. CH~3~-CO-CH~3~.*",
        "D. (CH~3~CH~2~)~2~O."
    ]
    render_options_no_table(doc, opts_c5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Propan-2-ol (CH~3~CH(OH)CH~3~) là alcohol bậc II, khi bị oxi hóa bởi CuO nung nóng sẽ tạo ketone tương ứng là acetone:\n"
            "CH~3~CH(OH)CH~3~ + CuO --t°--> CH~3~COCH~3~ + Cu + H~2~O."
        )

    # Câu 6
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 6. X là dẫn xuất chlorine của ethane. Đun nóng X trong dung dịch NaOH dư thu được chất hữu cơ Y "
        "vừa tác dụng với Na vừa hòa tan Cu(OH)~2~ ở nhiệt độ thường tạo dung dịch xanh lam đậm. Chất X là"
    )
    opts_c6 = [
        "A. 1,1,2,2-tetrachloroethane.",
        "*B. 1,2-dichloroethane.*",
        "C. 1,1-dichloroethane.",
        "D. 1,1,1-trichloroethane."
    ]
    render_options_no_table(doc, opts_c6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Chất Y hòa tan Cu(OH)~2~ ở nhiệt độ thường tạo dung dịch xanh lam đậm nên Y là polyalcohol có ít nhất hai nhóm -OH liền kề.\n"
            "Y dẫn xuất từ ethane nên Y là ethylene glycol (HO-CH~2~-CH~2~-OH).\n"
            "Do đó chất ban đầu X là 1,2-dichloroethane (Cl-CH~2~-CH~2~-Cl):\n"
            "ClCH~2~CH~2~Cl + 2NaOH --t°--> HOCH~2~CH~2~OH + 2NaCl."
        )

    # Câu 7
    p = add_p(doc)
    render_question_prompt(p, "Câu 7. Công thức cấu tạo đúng của tert-butyl alcohol là")
    opts_c7 = [
        "*A. (CH~3~)~3~COH.*",
        "B. (CH~3~)~3~CCH~2~OH.",
        "C. (CH~3~)~2~CHCH~2~OH.",
        "D. CH~3~CH(OH)CH~2~CH~3~."
    ]
    render_options_no_table(doc, opts_c7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Gốc tert-butyl có cấu tạo -(CH~3~)~3~C, do đó tert-butyl alcohol có công thức là (CH~3~)~3~C-OH (2-methylpropan-2-ol)."
        )

    # Câu 8
    p = add_p(doc)
    render_question_prompt(p, "Câu 8. Sơ đồ liên kết hydrogen nào sau đây biểu diễn sai?")
    opts_c8 = [
        "A. ... O(C~2~H~5~) - H ... O(H) - H.",
        "B. ... O(C~2~H~5~) - H ... O(C~2~H~5~) - H.",
        "C. HO(CH~2~-CH~2~) ... H - O.",
        "*D. H - C(=O) - OH ... H - C(=O) - OH.*"
    ]
    render_options_no_table(doc, opts_c8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Liên kết hydrogen trong carboxylic acid được hình thành giữa nguyên tử H mang điện tích dương của nhóm -OH phân tử này với nguyên tử oxygen mang điện tích âm của nhóm C=O phân tử kia: ... C=O ... H-O-. Sơ đồ D biểu diễn O-H ... H-C là sai bản chất."
        )

    # Câu 9
    p = add_p(doc)
    render_question_prompt(p, "Câu 9. Phenol không được sử dụng trong ngành công nghiệp sản xuất nào sau đây?")
    opts_c9 = [
        "A. Chất dẻo.",
        "B. Dược phẩm.",
        "*C. Cao su.*",
        "D. Thuốc nổ."
    ]
    render_options_no_table(doc, opts_c9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Phenol được dùng để sản xuất chất dẻo (nhựa phenol-formaldehyde), dược phẩm (aspirin, salol), thuốc nổ (picric acid), phẩm nhuộm, chất diệt cỏ (2,4-D), nhưng không dùng làm nguyên liệu sản xuất cao su."
        )

    # Câu 10
    p = add_p(doc)
    render_question_prompt(p, "Câu 10. Ảnh hưởng của nhóm -OH đến gốc C~6~H~5~- trong phân tử phenol thể hiện qua phản ứng giữa phenol với")
    opts_c10 = [
        "A. H~2~ (Ni, t°).",
        "B. Na kim loại.",
        "*C. dung dịch nước Br~2~.*",
        "D. dung dịch NaOH."
    ]
    render_options_no_table(doc, opts_c10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Nhóm -OH đẩy electron vào vòng thơm làm tăng mật độ electron ở các vị trí ortho và para, giúp phenol dễ dàng phản ứng thế với dung dịch nước bromine tạo kết tủa trắng 2,4,6-tribromophenol ngay ở nhiệt độ thường."
        )

    # Câu 11
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 11. Tiến hành thí nghiệm (A, B, C) ở điều kiện thường về phenol (C~6~H~5~OH) và muối C~6~H~5~ONa như hình vẽ sau đây:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "phenol_3tubes_experiment.png"), width_cm=8.5)
    p_c11 = add_p(doc, space_before=2, space_after=3)
    render_rich_text(p_c11, "Thông qua các thí nghiệm trên, khẳng định nào sau đây là chính xác?")
    opts_c11 = [
        "A. Phenol ít tan trong nước lạnh, nhưng tan nhiều trong dung dịch base, có lực kiềm yếu hơn cả nấc thứ nhất của carbonic acid.",
        "B. Phenol ít tan trong nước nóng, tan nhiều trong dung dịch kiềm, có lực acid mạnh hơn nấc thứ nhất của carbonic acid.",
        "C. Phenol tan nhiều trong nước nóng, tan nhiều trong dung dịch kiềm, có lực acid mạnh hơn nấc thứ nhất của carbonic acid.",
        "*D. Phenol ít tan trong nước lạnh, tan nhiều trong dung dịch kiềm, có lực acid yếu hơn nấc thứ nhất của carbonic acid.*"
    ]
    render_options_no_table(doc, opts_c11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- Ống A: Phenol ít tan trong nước lạnh tạo vẩn đục/phân lớp đáy.\n"
            "- Ống B: Phenol tan tốt trong kiềm tạo dung dịch trong suốt: C~6~H~5~OH + NaOH -> C~6~H~5~ONa + H~2~O.\n"
            "- Ống C: Khi sục CO~2~ vào dung dịch C~6~H~5~ONa, phenol bị đẩy ra làm dung dịch vẩn đục trở lại: C~6~H~5~ONa + CO~2~ + H~2~O -> C~6~H~5~OH + NaHCO~3~.\n"
            "=> Phenol có lực acid yếu hơn nấc thứ nhất của carbonic acid (H~2~CO~3~)."
        )

    # Câu 12
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 12. Cho hợp chất thơm p-HO-C~6~H~4~-CH~2~OH lần lượt tác dụng với: Na; dung dịch NaOH; dung dịch nước Br~2~; CuO nung nóng. "
        "Số trường hợp xảy ra phản ứng là"
    )
    opts_c12 = [
        "A. 3.",
        "B. 2.",
        "C. 1.",
        "*D. 4.*"
    ]
    render_options_no_table(doc, opts_c12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Chất p-HO-C~6~H~4~-CH~2~OH có cả nhóm -OH phenol và -OH alcohol:\n"
            "- Tác dụng với Na: cả hai nhóm -OH đều phản ứng giải phóng H~2~.\n"
            "- Tác dụng với NaOH: nhóm -OH phenol phản ứng tạo muối phenolate.\n"
            "- Tác dụng với nước Br~2~: phản ứng thế vào vị trí ortho của vòng thơm.\n"
            "- Tác dụng với CuO (t°): nhóm -CH~2~OH (alcohol bậc I) bị oxi hóa tạo aldehyde.\n"
            "=> Cả 4 trường hợp đều xảy ra phản ứng."
        )

    # Câu 13
    p = add_p(doc)
    render_question_prompt(p, "Câu 13. (SBT-CTST) Sản phẩm chính của phản ứng nào sau đây không đúng?")
    opts_c13 = [
        "A. CH~3~CH(Cl)CH~3~ + NaOH -> CH~3~CH(OH)CH~3~ + NaCl.",
        "*B. CH~3~CH~2~Cl + KOH -> CH~2~=CH~2~ + KCl + H~2~O.*",
        "C. CH~3~Br + KOH -> CH~3~OH + KBr.",
        "D. CH~3~CH~2~CH(Br)CH~3~ + KOH (ethanol) -> CH~3~CH=CHCH~3~ + KBr + H~2~O."
    ]
    render_options_no_table(doc, opts_c13, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Phản ứng ở B không ghi rõ điều kiện dung môi alcohol và nhiệt độ thì phản ứng ưu tiên là thế tạo alcohol (CH~3~CH~2~OH). Để tạo alkene cần dùng KOH trong ethanol và đun nóng."
        )

    # Câu 14
    p = add_p(doc)
    render_question_prompt(p, "Câu 14. (SBT-KNTT) Phương pháp nào sau đây được dùng để sản xuất ethanol sinh học?")
    opts_c14 = [
        "A. Cho hỗn hợp khí ethylene và hơi nước đi qua tháp chứa H~3~PO~4~.",
        "B. Cộng nước vào ethylene với xúc tác acid H~2~SO~4~.",
        "*C. Lên men tinh bột.*",
        "D. Thuỷ phân dẫn xuất C~2~H~5~Br trong môi trường kiềm."
    ]
    render_options_no_table(doc, opts_c14, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Ethanol sinh học được sản xuất bằng phương pháp lên men các nguồn nguyên liệu tái tạo giàu tinh bột (ngô, sắn, gạo,...) hoặc đường (mía, củ cải đường)."
        )

    # Câu 15
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 15. Cồn 70° là dung dịch ethyl alcohol thường được dùng để sát trùng y tế. Mô tả nào sau đây về cồn 70° là đúng?"
    )
    opts_c15 = [
        "A. 100 gam dung dịch có 70 mL ethyl alcohol nguyên chất.",
        "*B. 100 mL dung dịch có 70 mL ethyl alcohol nguyên chất.*",
        "C. 1 000 gam dung dịch có 70 mol ethyl alcohol nguyên chất.",
        "D. 100 mL dung dịch có 70 gam ethyl alcohol nguyên chất."
    ]
    render_options_no_table(doc, opts_c15, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Độ cồn là số mL ethyl alcohol nguyên chất có trong 100 mL dung dịch rượu. Cồn 70° chứa 70 mL ethanol nguyên chất trong 100 mL dung dịch."
        )

    # Câu 16
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 16. Cho các phát biểu sau:\n"
        "(1) Phenol C~6~H~5~OH là một alcohol thơm.\n"
        "(2) Phenol tác dụng được với NaOH tạo thành muối và nước.\n"
        "(3) Phenol tham gia phản ứng thế bromine và thế nitro dễ hơn benzene.\n"
        "(4) Dung dịch phenol làm quỳ tím hóa đỏ do có tính acid.\n"
        "(5) Giữa nhóm OH và vòng benzene trong phân tử phenol có sự ảnh hưởng qua lại lẫn nhau.\n"
        "Số phát biểu không đúng là"
    )
    opts_c16 = [
        "A. 1.",
        "B. 4.",
        "*C. 2.*",
        "D. 3."
    ]
    render_options_no_table(doc, opts_c16, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- (1) Sai: Phenol không phải là alcohol. Nhóm -OH liên kết trực tiếp với vòng thơm là phenol.\n"
            "- (2) Đúng: C~6~H~5~OH + NaOH -> C~6~H~5~ONa + H~2~O.\n"
            "- (3) Đúng: Do hiệu ứng đẩy electron của nhóm -OH vào vòng thơm.\n"
            "- (4) Sai: Phenol có tính acid rất yếu (yếu hơn H~2~CO~3~) nên không làm đổi màu quỳ tím.\n"
            "- (5) Đúng: Bản chất tương tác p-pi giữa nhóm -OH và vòng thơm.\n"
            "=> Có 2 phát biểu không đúng là (1) và (4)."
        )

    # Câu 17
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 17. Một hợp chất X có công thức phân tử C~4~H~8~O. X làm mất màu nước bromine, tác dụng với Na. "
        "Sản phẩm oxi hóa X bởi CuO (t°) không phải là aldehyde. Tên gọi của X là"
    )
    opts_c17 = [
        "A. but-3-en-1-ol.",
        "*B. but-3-en-2-ol.*",
        "C. 2-methylpropenol.",
        "D. butan-2-ol."
    ]
    render_options_no_table(doc, opts_c17, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- X tác dụng với Na => X có nhóm -OH.\n"
            "- X làm mất màu nước bromine => X có liên kết đôi C=C.\n"
            "- Oxi hóa bởi CuO không tạo aldehyde => X là alcohol bậc II (tạo ketone).\n"
            "Chất thỏa mãn là but-3-en-2-ol (CH~2~=CH-CH(OH)-CH~3~)."
        )

    # Câu 18
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 18. Cho sơ đồ tổng hợp glycerol từ propylene:\n"
        "C~3~H~6~ --(Cl~2~, as, 500°C)--> X --(+Cl~2~)--> Y --(KOH, H~2~O, t°)--> glycerol\n"
        "Các chất X, Y tương ứng là"
    )
    opts_c18 = [
        "*A. X: CH~2~=CH-CH~2~Cl; Y: CH~2~Cl-CHCl-CH~2~Cl.*",
        "B. X: CH~2~Cl-CHCl-CH~3~; Y: CH~2~Cl-CHCl-CH~2~Cl.",
        "C. X: CH~2~Cl-CHCl-CH~3~; Y: CH~2~=CH-CH~2~Cl.",
        "D. X: CHCl~2~-CH=CH~2~; Y: CH~2~Cl-CHCl-CHCl~2~."
    ]
    render_options_no_table(doc, opts_c18, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- Giai đoạn 1: Thế allyl ở 500°C tạo allyl chloride (X): CH~2~=CH-CH~3~ + Cl~2~ --500°C--> CH~2~=CH-CH~2~Cl + HCl.\n"
            "- Giai đoạn 2: Cộng Cl~2~ vào liên kết đôi tạo 1,2,3-trichloropropane (Y): CH~2~=CH-CH~2~Cl + Cl~2~ -> CH~2~Cl-CHCl-CH~2~Cl.\n"
            "- Giai đoạn 3: Thủy phân Y trong kiềm tạo glycerol: CH~2~Cl-CHCl-CH~2~Cl + 3KOH --t°--> CH~2~OH-CHOH-CH~2~OH + 3KCl."
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
        "Câu 1. Ứng dụng và tác hại của dẫn xuất halogen.\n"
        "Trong đời sống, các hợp chất chlorofluorocarbon (CFC) từng được sử dụng rộng rãi trong công nghệ làm lạnh nhưng cũng gây tác hại cho tầng ozone."
    )
    stmts_c1_p2 = [
        ("a) CFC là một loại dẫn xuất halogen được sử dụng trong các thiết bị làm lạnh.", "ĐÚNG"),
        ("b) Việc lạm dụng CFC gây ra hiện tượng thủng tầng ozone, làm tăng nguy cơ ung thư da.", "ĐÚNG"),
        ("c) Các hợp chất CFC hoàn toàn không gây hại cho môi trường nếu sử dụng đúng cách.", "SAI"),
        ("d) Việc thay thế CFC bằng các chất làm lạnh thân thiện môi trường là một giải pháp bền vững.", "ĐÚNG")
    ]
    for stmt, ans in stmts_c1_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) ĐÚNG: CFC (chlorofluorocarbon) là các hợp chất chứa C, Cl, F dùng làm môi chất lạnh.\n"
            "- b) ĐÚNG: Dưới tác dụng của tia tử ngoại UV, CFC phân rã tạo gốc Cl• phá hủy O~3~ làm thủng tầng ozone, tăng bức xạ UV gây ung thư da.\n"
            "- c) SAI: CFC rất bền vững trong khí quyển, khi phát tán đều phá hủy tầng ozone và gây hiệu ứng nhà kính.\n"
            "- d) ĐÚNG: Thay thế CFC bằng HFC, HFO thân thiện môi trường là giải pháp bền vững."
        )

    # Câu 2 Phần II
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 2. Thực hiện thí nghiệm chưng cất ethanol từ dung dịch ethanol – nước:\n"
        "- Cho 60 mL rượu nấu thủ công vào bình cầu có nhánh (chất lỏng không quá 2/3 thể tích), thêm vài viên đá bọt.\n"
        "- Lắp dụng cụ như hình bên dưới:\n"
        "- Đun nóng từ từ đến khi hỗn hợp sôi, quan sát nhiệt kế thấy nhiệt độ tăng dần rồi ổn định ở nhiệt độ sôi của ethanol."
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "distillation_setup.png"), width_cm=8.5)
    stmts_c2_p2 = [
        ("a) Nhiệt kế dùng để đo nhiệt độ của hơi chất đang chưng cất.", "ĐÚNG"),
        ("b) Bình hứng thu được nước nguyên chất.", "SAI"),
        ("c) Đá bọt có vai trò điều hòa quá trình sôi, tránh hiện tượng quá sôi gây bùng nổ.", "ĐÚNG"),
        ("d) Có thể tách ethanol ra khỏi nước bằng phương pháp chiết lỏng - lỏng.", "SAI")
    ]
    for stmt, ans in stmts_c2_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) ĐÚNG: Bầu nhiệt kế đặt ngang nhánh bình cầu để đo nhiệt độ của hơi chưng cất ngưng tụ.\n"
            "- b) SAI: Ethanol có nhiệt độ sôi (78,3°C) thấp hơn nước (100°C) nên bay hơi trước, bình hứng thu được dung dịch giàu ethanol.\n"
            "- c) ĐÚNG: Đá bọt tạo tâm sôi giúp chất lỏng sôi êm dịu, chống quá sôi.\n"
            "- d) SAI: Ethanol tan vô hạn trong nước tạo dung dịch đồng nhất không phân lớp nên không thể tách bằng phương pháp chiết, phải dùng chưng cất."
        )

    # Câu 3 Phần II
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 3. Tiến hành thí nghiệm thử tính chất của glycerol và ethanol với copper(II) hydroxide:\n"
        "Bước 1: Cho vào 2 ống nghiệm, mỗi ống 3 - 4 giọt dung dịch CuSO~4~ 2% và 2 - 3 giọt dung dịch NaOH 10%, lắc nhẹ.\n"
        "Bước 2: Nhỏ 2 - 3 giọt glycerol vào ống thứ nhất, 2 - 3 giọt ethanol vào ống thứ hai. Lắc nhẹ cả hai ống nghiệm."
    )
    stmts_c3_p2 = [
        ("a) Sau bước 1, trong cả hai ống nghiệm đều có kết tủa màu xanh lam của copper(II) hydroxide.", "ĐÚNG"),
        ("b) Sau bước 2, trong cả hai ống nghiệm kết tủa tan, tạo thành dung dịch màu xanh lam thẫm.", "SAI"),
        ("c) Thí nghiệm trên được dùng để phân biệt ethanol và glycerol.", "ĐÚNG"),
        ("d) Ở bước 1, nếu thay dung dịch NaOH bằng dung dịch Ba(OH)~2~ thì hiện tượng thu được hoàn toàn không đổi.", "SAI")
    ]
    for stmt, ans in stmts_c3_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) ĐÚNG: Xuất hiện kết tủa xanh Cu(OH)~2~: CuSO~4~ + 2NaOH -> Cu(OH)~2~ v + Na~2~SO~4~.\n"
            "- b) SAI: Chỉ có glycerol hòa tan kết tủa tạo phức xanh lam đậm, ethanol không hòa tan kết tủa.\n"
            "- c) ĐÚNG: Dùng để phân biệt polyalcohol có nhóm -OH liền kề (glycerol) với alcohol đơn chức (ethanol).\n"
            "- d) SAI: Nếu dùng Ba(OH)~2~ thì tạo thêm kết tủa trắng BaSO~4~ làm vẫn đục dung dịch, hiện tượng bị thay đổi."
        )

    # Câu 4 Phần II
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 4. Vitamin E là chất chống oxi hoá tự nhiên ngăn chặn tổn thương tế bào. "
        "BHT (butylated hydroxytoluene) là chất chống oxi hoá tổng hợp thường dùng trong thực phẩm. Cấu trúc của 2 chất cho như hình bên:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "vitamin_e_bht.png"), width_cm=11.0)
    stmts_c4_p2 = [
        ("a) Cả vitamin E và BHT đều có chứa nhóm OH phenol.", "ĐÚNG"),
        ("b) Công thức phân tử của BHT là C~15~H~36~O.", "SAI"),
        ("c) Vitamin E là một hợp chất hữu cơ tạp chức.", "ĐÚNG"),
        ("d) BHT có nhiệt độ nóng chảy và nhiệt độ sôi cao hơn ethanol.", "SAI")
    ]
    for stmt, ans in stmts_c4_p2:
        p_st = add_p(doc, space_before=1, space_after=1)
        p_st.paragraph_format.left_indent = Cm(0.5)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "- a) ĐÚNG: Cả hai phân tử đều có nhóm -OH gắn trực tiếp vào vòng benzene.\n"
            "- b) SAI: Công thức phân tử của BHT là C~15~H~24~O (không phải C~15~H~36~O).\n"
            "- c) ĐÚNG: Vitamin E vừa chứa nhóm chức phenol (-OH) vừa chứa nhóm chức ether (-O-).\n"
            "- d) SAI (theo đáp án quy chuẩn của đề)."
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
        "Câu 1. Cho các hợp chất sau: CCl~4~, NaCl, CH~3~COOH, C~2~H~5~OH, CH~3~CH~2~CH(Br)CH~3~, CH~2~BrCl, HIO~4~, "
        "CH~2~=CH-CH~3~, KBr, CH~2~=CHCl. Có bao nhiêu hợp chất hữu cơ thuộc loại dẫn xuất halogen của hydrocarbon?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Các dẫn xuất halogen của hydrocarbon gồm 4 chất:\n"
            "1. CCl~4~ (tetrachloromethane)\n"
            "2. CH~3~CH~2~CH(Br)CH~3~ (2-bromobutane)\n"
            "3. CH~2~BrCl (bromochloromethane)\n"
            "4. CH~2~=CHCl (vinyl chloride)\n"
            "=> Đáp số: 4."
        )

    # Câu 2 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 2. Dẫn xuất halogen A có công thức C~3~H~8-x~Br~x~, biết phần trăm khối lượng của H trong A bằng 5,7%. "
        "Giá trị của x là bao nhiêu?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Phân tử khối: M_A = 12 × 3 + (8 - x) + 80x = 44 + 79x.\n"
            "%H = [(8 - x) / (44 + 79x)] × 100% = 5,7%\n"
            "<=> 100(8 - x) = 5,7(44 + 79x) <=> 800 - 100x = 250,8 + 450,3x <=> 550,3x = 549,2 => x ≈ 1.\n"
            "=> Đáp số: 1."
        )

    # Câu 3 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 3. Cho các chất sau:\n"
        "(1) CH~3~-CH~2~-OH;\n"
        "(2) CH~2~=CH-CH~2~-OH;\n"
        "(3) CH~2~=CH-OH;\n"
        "(4) CH~3~-COOH;\n"
        "(5) CH~2~OH-CHOH-CH~2~OH;\n"
        "(6) C~6~H~5~OH;\n"
        "(7) C~6~H~5~CH~2~OH.\n"
        "Có bao nhiêu chất thuộc loại alcohol?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Alcohol là hợp chất có nhóm -OH gắn trực tiếp vào nguyên tử carbon no:\n"
            "- (1) CH~3~CH~2~OH (ethanol) : alcohol.\n"
            "- (2) CH~2~=CH-CH~2~OH (allyl alcohol) : alcohol.\n"
            "- (3) CH~2~=CH-OH không bền tự chuyển vị thành CH~3~CHO.\n"
            "- (4) CH~3~COOH : carboxylic acid.\n"
            "- (5) CH~2~OH-CHOH-CH~2~OH (glycerol) : polyalcohol.\n"
            "- (6) C~6~H~5~OH : phenol.\n"
            "- (7) C~6~H~5~CH~2~OH (benzyl alcohol) : alcohol thơm.\n"
            "=> Có 4 chất là alcohol: (1), (2), (5), (7).\n"
            "=> Đáp số: 4."
        )

    # Câu 4 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 4. Hợp chất hữu cơ có công thức phân tử C~7~H~8~O chứa vòng benzene. "
        "Có bao nhiêu đồng phân cấu tạo phản ứng được với dung dịch NaOH?"
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Để phản ứng được với NaOH, hợp chất thơm C~7~H~8~O phải là phenol (nhóm -OH gắn trực tiếp vào vòng benzene):\n"
            "Gồm 3 đồng phân cresol (CH~3~-C~6~H~4~-OH):\n"
            "- 1. o-cresol (2-methylphenol)\n"
            "- 2. m-cresol (3-methylphenol)\n"
            "- 3. p-cresol (4-methylphenol)\n"
            "=> Đáp số: 3."
        )
        add_centered_picture(doc, os.path.join(IMG_DIR, "cresol_isomers.png"), width_cm=7.0)

    # Câu 5 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 5. Củ sắn khô chứa 38% khối lượng là tinh bột, còn lại là các chất không có khả năng lên men thành ethyl alcohol. "
        "Tính khối lượng ethyl alcohol thu được (theo kg) khi lên men 1 tấn củ sắn khô với hiệu suất của cả quá trình là 81%."
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Khối lượng sắn khô: m = 1 tấn = 1000 kg.\n"
            "Khối lượng tinh bột: m_tinhbot = 1000 × 38% = 380 kg.\n"
            "Sơ đồ: (C~6~H~10~O~5~)~n~ -> 2n C~2~H~5~OH.\n"
            "Theo lí thuyết: 162 kg tinh bột tạo 92 kg C~2~H~5~OH.\n"
            "Khối lượng ethanol thực tế thu được với H = 81%:\n"
            "m_ethanol = (380 × 92 / 162) × 81% = 174,8 kg.\n"
            "=> Đáp số: 174,8."
        )

    # Câu 6 Phần III
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 6. Cho sơ đồ chuyển hóa: CH~4~ -> C~2~H~2~ -> C~2~H~3~Cl -> PVC. "
        "Theo sơ đồ trên, từ 495,8 m^3 khí thiên nhiên (ở đkc) thì tổng hợp được bao nhiêu kg PVC? "
        "Biết CH~4~ chiếm 80% thể tích khí thiên nhiên và hiệu suất của cả quá trình là 50%."
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Thể tích CH~4~: V = 495,8 × 80% = 396,64 m^3 = 396 640 L.\n"
            "Số mol CH~4~ (đkc: 24,79 L/mol): n = 396 640 / 24,79 = 16 000 mol = 16 kmol.\n"
            "Bảo toàn C: 2 CH~4~ -> 1 mắt xích PVC (M = 62,5 g/mol).\n"
            "Khối lượng PVC theo lí thuyết: m_lt = (16 / 2) × 62,5 = 500 kg.\n"
            "Với hiệu suất 50%: m_tt = 500 × 50% = 250 kg.\n"
            "=> Đáp số: 250."
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
        "Câu 1. Điền các thông tin thích hợp vào ô trống để hoàn thành bảng mô tả về các đồng phân có công thức phân tử C~3~H~8~O sau:"
    )
    add_centered_picture(doc, os.path.join(IMG_DIR, "c3h8o_isomers_table.png"), width_cm=12.0)
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Bảng hoàn thành thông tin mô tả các đồng phân C~3~H~8~O:\n"
            "(1) Công thức cấu tạo của Ethyl methyl ether: CH~3~-CH~2~-O-CH~3~\n"
            "(2) Công thức cấu tạo alcohol bậc một: CH~3~-CH~2~-CH~2~-OH\n"
            "(3) Công thức cấu tạo alcohol bậc hai: CH~3~-CH(OH)-CH~3~\n"
            "(4) Tên gọi alcohol bậc một: Propan-1-ol\n"
            "(5) Tên gọi alcohol bậc hai: Propan-2-ol\n"
            "(6) Phản ứng của ether với Na: Không xảy ra\n"
            "(7) Phản ứng của alcohol bậc một với Na: tạo ra H~2~\n"
            "(8) Phản ứng của alcohol bậc hai với Na: tạo ra H~2~\n"
            "(9) Phản ứng của ether với CuO, t°: Không xảy ra\n"
            "(10) Phản ứng của alcohol bậc một với CuO, t°: tạo ra aldehyde CH~3~-CH~2~-CHO\n"
            "(11) Phản ứng của alcohol bậc hai với CuO, t°: tạo ra ketone CH~3~-CO-CH~3~"
        )
        add_centered_picture(doc, os.path.join(IMG_DIR, "c3h8o_isomers_solution_table.png"), width_cm=12.0)
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 2 Phần IV
    p = add_p(doc)
    render_question_prompt(
        p,
        "Câu 2. Khối lượng ethyl alcohol nguyên chất có trong 0,5 lít cồn 96° là bao nhiêu gam? "
        "Biết ethyl alcohol nguyên chất có khối lượng riêng D = 0,8 g/mL."
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Lời giải chi tiết:\n"
            "Thể tích dung dịch cồn: V_dd = 0,5 L = 500 mL.\n"
            "Độ cồn bằng 96° nên thể tích ethyl alcohol nguyên chất là:\n"
            "V_alcohol = (V_dd × 96) / 100 = (500 × 96) / 100 = 480 mL.\n"
            "Khối lượng ethyl alcohol nguyên chất là:\n"
            "m_alcohol = V_alcohol × D = 480 mL × 0,8 g/mL = 384 gam.\n"
            "=> Đáp số: 384 gam."
        )
    else:
        add_student_writing_lines(doc, num_lines=6)

    return doc

def check_and_save():
    os.makedirs(OUT_DIR, exist_ok=True)

    # 1. Bản Học sinh
    doc_hs = build_exam_document(is_teacher=False)
    tables_hs = len(doc_hs.tables)
    print(f"Bản Học sinh có {tables_hs} bảng (Chỉ được phép có 1 bảng Header block).")
    assert tables_hs <= 1, "LỖI: Bản Học sinh chứa bảng trong nội dung!"
    path_hs = os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_11_De15.docx")
    doc_hs.save(path_hs)
    print("Đã lưu Bản Học sinh:", path_hs)

    # 2. Bản Giáo viên
    doc_gv = build_exam_document(is_teacher=True)
    tables_gv = len(doc_gv.tables)
    print(f"Bản Giáo viên có {tables_gv} bảng (Chỉ được phép có 1 bảng Header block).")
    assert tables_gv <= 1, "LỖI: Bản Giáo viên chứa bảng trong nội dung!"
    path_gv = os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_11_De15_GiaoVien.docx")
    doc_gv.save(path_gv)
    print("Đã lưu Bản Giáo viên:", path_gv)

if __name__ == "__main__":
    check_and_save()
