# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Kỳ I Hóa Học 12 - THPT Số 2 Quang Trung (Gia Lai) - Mã đề 2102
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)

Xuất bản 2 file Word:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_12_THPT_So2_QuangTrung_Ma2102.docx
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_12_THPT_So2_QuangTrung_Ma2102_GiaoVien.docx (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

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
    add_p_border_bottom,
    COLOR_PRIMARY,
    COLOR_CORRECT,
    COLOR_SOLUTION,
    COLOR_GRAY
)

OUTPUT_DIR = r"C:\Antigravity_Thanh\San_Pham"
IMG_DIR = os.path.join(OUTPUT_DIR, "images_quangtrung2_hoa12_2102")

def set_cell_border(cell, **kwargs):
    """Cài đặt viền cho ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            val, sz, space, color = edge_data
            tag = f'<w:{edge} {nsdecls("w")} w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'
            tcBorders.append(parse_xml(tag))
    tcPr.append(tcBorders)

def set_cell_shading(cell, color_hex):
    """Tô màu nền cho ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def add_quick_answer_table_2102(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu MÃ ĐỀ 2102.
    """
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 2102)*", base_size=12, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(4)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (16 CÂU - 4,0 ĐIỂM)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=4, cols=9)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_p1_1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8"]
    answers_p1_1 = ["Đ/A", "A", "B", "A", "C", "A", "C", "B", "C"]
    headers_p1_2 = ["Câu", "9", "10", "11", "12", "13", "14", "15", "16"]
    answers_p1_2 = ["Đ/A", "A", "A", "D", "A", "A", "D", "C", "B"]

    col_w_0 = Cm(2.0)
    col_w = Cm(2.0)

    for r_idx, r_data in enumerate([headers_p1_1, answers_p1_1, headers_p1_2, answers_p1_2]):
        row = tbl1.rows[r_idx]
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = col_w_0 if c_idx == 0 else col_w
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            is_header_row = (r_idx in [0, 2])
            is_first_col = (c_idx == 0)

            if is_header_row:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                set_cell_shading(cell, "FFFFFF" if not is_first_col else "F5F5F5")
                c_color = COLOR_CORRECT if not is_first_col else COLOR_PRIMARY
                render_rich_text(p, f"*{val}*", base_size=10.5, base_bold=True, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # 2. PHẦN II
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(6)
    p_p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2, "*PHẦN II: TRẮC NGHIỆM ĐÚNG - SAI (2 CÂU - 2,0 ĐIỂM)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=5)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl2.autofit = False

    p2_data = [
        ["Câu", "Lệnh hỏi a", "Lệnh hỏi b", "Lệnh hỏi c", "Lệnh hỏi d"],
        ["Câu 1", "S", "Đ", "S", "S"],
        ["Câu 2", "S", "Đ", "Đ", "Đ"]
    ]
    col_w2 = [Cm(3.0), Cm(3.75), Cm(3.75), Cm(3.75), Cm(3.75)]
    for r_idx, row_vals in enumerate(p2_data):
        row = tbl2.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = col_w2[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                set_cell_shading(cell, "FFFFFF" if c_idx > 0 else "F5F5F5")
                c_color = COLOR_CORRECT if c_idx > 0 else COLOR_PRIMARY
                render_rich_text(p, f"*{val}*", base_size=10.5, base_bold=True, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # 3. PHẦN III
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(6)
    p_p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3, "*PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (4 CÂU - 2,0 ĐIỂM)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=5)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False

    p3_data = [
        ["Câu", "Câu 1", "Câu 2", "Câu 3", "Câu 4"],
        ["Đáp án", "2", "4", "3", "3"]
    ]
    col_w3 = [Cm(3.0), Cm(3.75), Cm(3.75), Cm(3.75), Cm(3.75)]
    for r_idx, row_vals in enumerate(p3_data):
        row = tbl3.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = col_w3[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                set_cell_shading(cell, "FFFFFF" if c_idx > 0 else "F5F5F5")
                c_color = COLOR_CORRECT if c_idx > 0 else COLOR_PRIMARY
                render_rich_text(p, f"*{val}*", base_size=10.5, base_bold=True, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # 4. PHẦN IV TỰ LUẬN TÓM TẮT ĐÁP ÁN
    p_p4 = doc.add_paragraph()
    p_p4.paragraph_format.space_before = Pt(6)
    p_p4.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p4, "*PHẦN IV: TỰ LUẬN (2 CÂU - 2,0 ĐIỂM)*", base_size=10.5, base_bold=True, base_color=COLOR_PRIMARY)
    p_p4_ans = doc.add_paragraph()
    p_p4_ans.paragraph_format.space_before = Pt(2)
    p_p4_ans.paragraph_format.space_after = Pt(6)
    render_rich_text(p_p4_ans, "*Câu 1 (1,0 điểm):* a) Ống 1: phức xanh lam thẫm (nhiều nhóm -OH kề nhau). Ống 2: kết tủa đỏ gạch Cu~2~O (nhóm aldehyde -CHO). b) PTHH: CH~2~OH[CHOH]~4~CHO + 2Cu(OH)~2~ + NaOH -> CH~2~OH[CHOH]~4~COONa + Cu~2~O + 3H~2~O.\n*Câu 2 (1,0 điểm):* a) (C~17~H~35~COO)~3~C~3~H~5~ + 3NaOH -> 3C~17~H~35~COONa + C~3~H~5~(OH)~3~. b) NaCl bão hòa làm tăng tỉ trọng và giảm độ tan của xà phòng giúp xà phòng tách lớp nổi lên. c) Khối lượng tristearin tối thiểu: *3,56 kg*.", base_size=10, base_color=COLOR_GRAY)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(10)
    add_p_border_bottom(p_div, color="8BC390", sz="12")


def build_exam_document(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="SỞ GD & ĐT GIA LAI\nTRƯỜNG THPT SỐ 2 QUANG TRUNG",
        exam_title="ĐỀ KIỂM TRA GIỮA KỲ I, NĂM HỌC 2026 - 2027",
        subject_line="MÔN: HÓA HỌC - KHỐI 12",
        num_pages="04",
        duration="45 phút (không kể thời gian phát đề)",
        exam_code="2102",
        atomic_masses="C = 12; H = 1; O = 16; Na = 23; Br = 80; N = 14",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_2102(doc)

    # ==========================================
    # PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (4,0 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    # Câu 1
    p1 = doc.add_paragraph()
    render_question_prompt(
        p1,
        "Câu 1. Cho các phát biểu sau:\n"
        "(1) Alkylamine bậc một tác dụng với nitrous acid ở nhiệt độ thường tạo alcohol và giải phóng khí nitrogen.\n"
        "(2) Dung dịch methylamine tác dụng với dung dịch FeCl~3~ tạo kết tủa nâu đỏ Fe(OH)~3~.\n"
        "(3) Tripeptide Ala-Gly-Ala tham gia phản ứng màu biuret.\n"
        "(4) Ở điều kiện thường, các amino acid là chất lỏng, ít tan trong nước.\n"
        "(5) Protein trong dung dịch lòng trắng trứng bị đông tụ khi đun nóng.\n"
        "(6) Aniline dễ tham gia phản ứng thế nguyên tử hydrogen ở nhân thơm hơn benzene.\n"
        "Số phát biểu đúng là"
    )
    render_options_no_table(
        doc,
        ["*A. 5.*", "B. 4.", "C. 2.", "D. 3."] if is_teacher else ["A. 5.", "B. 4.", "C. 2.", "D. 3."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "• (1) Đúng: Alkylamine bậc một tác dụng với HNO~2~ ở nhiệt độ thường sinh ra alcohol và giải phóng khí N~2~ (RNH~2~ + HNO~2~ -> ROH + N~2~↑ + H~2~O).\n"
            "• (2) Đúng: Dung dịch methylamine có tính base, tác dụng với muối Fe^3+^ tạo kết tủa nâu đỏ Fe(OH)~3~:\n"
            "   3CH~3~NH~2~ + FeCl~3~ + 3H~2~O -> Fe(OH)~3~↓ (nâu đỏ) + 3CH~3~NH~3~Cl.\n"
            "• (3) Đúng: Tripeptide có 2 liên kết peptide nên có phản ứng màu biuret với Cu(OH)~2~ trong môi trường kiềm.\n"
            "• (4) Sai: Ở điều kiện thường, các amino acid là chất rắn kết tinh, không màu, vị hơi ngọt, dễ tan trong nước do tồn tại ở dạng ion lưỡng cực.\n"
            "• (5) Đúng: Protein bị đông tụ dưới tác dụng của nhiệt độ cao hoặc acid/base/ion kim loại nặng.\n"
            "• (6) Đúng: Do nhóm -NH~2~ đẩy electron vào vòng benzene làm tăng mật độ electron ở các vị trí ortho và para, giúp aniline phản ứng thế halogen/HNO~3~ dễ dàng hơn nhiều so với benzene.\n"
            "=> Có 5 phát biểu đúng: (1), (2), (3), (5), (6)."
        )

    # Câu 2
    p2 = doc.add_paragraph()
    render_question_prompt(
        p2,
        "Câu 2. Mùi tanh của cá chủ yếu do trimethylamine gây ra. Phát biểu nào sau đây đúng về tính chất vật lí của amine?"
    )
    render_options_no_table(
        doc,
        [
            "A. Các amine có số nguyên tử carbon nhỏ đều không tan trong nước.",
            "*B. Trimethylamine là chất khí ở điều kiện thường.*",
            "C. Tất cả các amine đều là chất rắn ở điều kiện thường.",
            "D. Aniline là chất khí, tan tốt trong nước."
        ] if is_teacher else [
            "A. Các amine có số nguyên tử carbon nhỏ đều không tan trong nước.",
            "B. Trimethylamine là chất khí ở điều kiện thường.",
            "C. Tất cả các amine đều là chất rắn ở điều kiện thường.",
            "D. Aniline là chất khí, tan tốt trong nước."
        ],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "• Methylamine, dimethylamine, trimethylamine và ethylamine là những chất khí ở điều kiện thường, mùi khai khó chịu, độc, tan nhiều trong nước.\n"
            "• Aniline là chất lỏng, không màu, rất độc, ít tan trong nước, nặng hơn nước.\n"
            "=> Phát biểu đúng là B: Trimethylamine là chất khí ở điều kiện thường."
        )

    # Câu 3
    p3 = doc.add_paragraph()
    render_question_prompt(p3, "Câu 3. Dãy nào sau đây gồm polysaccharide?")
    render_options_no_table(
        doc,
        ["*A. Tinh bột, cellulose.*", "B. Saccharose, maltose.", "C. Glucose, fructose.", "D. Glucose, saccharose."] if is_teacher else ["A. Tinh bột, cellulose.", "B. Saccharose, maltose.", "C. Glucose, fructose.", "D. Glucose, saccharose."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "• Tinh bột và cellulose là các polysaccharide.\n"
            "• Saccharose và maltose là các disaccharide.\n"
            "• Glucose và fructose là các monosaccharide."
        )

    # Câu 4
    p4 = doc.add_paragraph()
    render_question_prompt(
        p4,
        "Câu 4. Saccharose được dùng nhiều trong sản xuất bánh kẹo và nước giải khát. Saccharose có nhiều trong"
    )
    render_options_no_table(
        doc,
        ["A. hạt gạo và hạt ngô.", "B. gỗ và sợi bông.", "*C. cây mía, củ cải đường và hoa thốt nốt.*", "D. hạt đang nảy mầm."] if is_teacher else ["A. hạt gạo và hạt ngô.", "B. gỗ và sợi bông.", "C. cây mía, củ cải đường và hoa thốt nốt.", "D. hạt đang nảy mầm."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Saccharose (đường mía) có nhiều nhất trong cây mía, củ cải đường và cụm hoa thốt nốt."
        )

    # Câu 5
    p5 = doc.add_paragraph()
    render_question_prompt(p5, "Câu 5. Amine CH~3~NHCH~3~ có tên gốc - chức là")
    render_options_no_table(
        doc,
        ["*A. dimethylamine.*", "B. methylamine.", "C. trimethylamine.", "D. ethylamine."] if is_teacher else ["A. dimethylamine.", "B. methylamine.", "C. trimethylamine.", "D. ethylamine."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Phân tử CH~3~-NH-CH~3~ gồm 2 gốc methyl gắn vào nguyên tử nitrogen mang chức amine nên có tên gốc - chức là dimethylamine."
        )

    # Câu 6
    p6 = doc.add_paragraph()
    render_question_prompt(
        p6,
        "Câu 6. Ở pH = 2, khi đặt dung dịch glycine trong điện trường thì glycine di chuyển về phía"
    )
    render_options_no_table(
        doc,
        ["A. cực dương (+).", "B. không di chuyển.", "*C. cực âm (-).* ", "D. cả hai cực."] if is_teacher else ["A. cực dương (+).", "B. không di chuyển.", "C. cực âm (-).", "D. cả hai cực."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Ở pH = 2 (môi trường acid mạnh, pH < pH~I~ của glycine ≈ 6,0), nhóm carboxyl nhận proton tạo thành cation: ^+^H~3~N-CH~2~-COOH mang điện tích dương. Khi đặt trong điện trường, cation này sẽ di chuyển về phía cực âm (-)."
        )

    # Câu 7
    p7 = doc.add_paragraph()
    render_question_prompt(p7, "Câu 7. Ester CH~3~CH~2~COOCH~3~ có tên gọi là")
    render_options_no_table(
        doc,
        ["A. methyl ethanoate.", "*B. methyl propanoate.*", "C. ethyl ethanoate.", "D. propyl methanoate."] if is_teacher else ["A. methyl ethanoate.", "B. methyl propanoate.", "C. ethyl ethanoate.", "D. propyl methanoate."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Ester CH~3~CH~2~COOCH~3~ gồm gốc alcohol là methyl (-CH~3~) và gốc carboxylate là propanoate (CH~3~CH~2~COO-) -> Tên gọi là methyl propanoate."
        )

    # Câu 8
    p8 = doc.add_paragraph()
    render_question_prompt(
        p8,
        "Câu 8. Mỡ lợn, dầu lạc, dầu hướng dương là những nguồn cung cấp chất béo quen thuộc trong bữa ăn hằng ngày. Về cấu tạo, chất béo là"
    )
    render_options_no_table(
        doc,
        ["A. ester của glycerol với acid vô cơ.", "B. monoester của glycerol với acid béo.", "*C. triester của glycerol với acid béo.*", "D. diester của glycerol với acid béo."] if is_teacher else ["A. ester của glycerol với acid vô cơ.", "B. monoester của glycerol với acid béo.", "C. triester của glycerol với acid béo.", "D. diester của glycerol với acid béo."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Theo định nghĩa: Chất béo là triester của glycerol với acid béo, gọi chung là triglyceride hay triacylglycerol."
        )

    # Câu 9
    p9 = doc.add_paragraph()
    render_question_prompt(
        p9,
        "Câu 9. Dung dịch saccharose không phản ứng với thuốc thử Tollens. Đun nóng dung dịch saccharose với dung dịch H~2~SO~4~ loãng, để nguội rồi trung hòa acid dư bằng dung dịch NaOH. Thêm thuốc thử Tollens vào dung dịch sau xử lí và đun cách thủy, thấy xuất hiện bạc. Nguyên nhân là do:"
    )
    render_options_no_table(
        doc,
        [
            "*A. saccharose bị thủy phân tạo thành glucose và fructose, các sản phẩm này phản ứng được với thuốc thử Tollens.*",
            "B. trong dung dịch, saccharose tồn tại chủ yếu ở dạng mạch hở có nhóm -CHO.",
            "C. acid loãng oxi hóa saccharose thành các chất có nhóm -CHO.",
            "D. acid loãng khử ion Ag^+^ trong thuốc thử Tollens thành Ag."
        ] if is_teacher else [
            "A. saccharose bị thủy phân tạo thành glucose và fructose, các sản phẩm này phản ứng được với thuốc thử Tollens.",
            "B. trong dung dịch, saccharose tồn tại chủ yếu ở dạng mạch hở có nhóm -CHO.",
            "C. acid loãng oxi hóa saccharose thành các chất có nhóm -CHO.",
            "D. acid loãng khử ion Ag^+^ trong thuốc thử Tollens thành Ag."
        ],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Saccharose không có nhóm -CHO hay -OH hemiacetal nên không tráng bạc. Khi đun với acid, saccharose bị thủy phân thành glucose và fructose: C~12~H~22~O~11~ + H~2~O -(H^+^, t°)-> C~6~H~12~O~6~ (glucose) + C~6~H~12~O~6~ (fructose). Cả hai chất này đều tham gia phản ứng tráng bạc với thuốc thử Tollens."
        )

    # Câu 10
    p10 = doc.add_paragraph()
    render_question_prompt(p10, "Câu 10. Chất nào sau đây là amine bậc hai?")
    render_options_no_table(
        doc,
        ["*A. CH~3~NHC~2~H~5~.*", "B. (CH~3~)~3~N.", "C. CH~3~NH~2~.", "D. C~6~H~5~NH~2~."] if is_teacher else ["A. CH~3~NHC~2~H~5~.", "B. (CH~3~)~3~N.", "C. CH~3~NH~2~.", "D. C~6~H~5~NH~2~."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "• CH~3~NHC~2~H~5~: nguyên tử N liên kết trực tiếp với 2 gốc hydrocarbon -> amine bậc hai.\n"
            "• (CH~3~)~3~N: amine bậc ba; CH~3~NH~2~ và C~6~H~5~NH~2~: amine bậc một."
        )

    # Câu 11
    p11 = doc.add_paragraph()
    render_question_prompt(
        p11,
        "Câu 11. Ethyl butanoate và ethyl ethanoate là những ester có mùi thơm, được dùng làm chất tạo hương trong công nghiệp thực phẩm và mĩ phẩm. Khi thay thế nhóm -OH trong nhóm carboxyl (-COOH) của carboxylic acid bằng nhóm -OR' (R' là gốc hydrocarbon) thì thu được hợp chất thuộc loại nào sau đây?"
    )
    render_options_no_table(
        doc,
        ["A. Amino acid.", "B. Alcohol.", "C. Amine.", "*D. Ester.*"] if is_teacher else ["A. Amino acid.", "B. Alcohol.", "C. Amine.", "D. Ester."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Khi thay thế nhóm -OH trong nhóm carboxyl (-COOH) của carboxylic acid bằng nhóm -OR' (R' là gốc hydrocarbon) thì thu được ester (RCOOR')."
        )

    # Câu 12
    p12 = doc.add_paragraph()
    render_question_prompt(
        p12,
        "Câu 12. Một cơ sở tận dụng dầu ăn đã qua sử dụng để sản xuất biodiesel. Trong quy trình tái chế, chất béo trong dầu ăn phản ứng với methanol có xúc tác thích hợp, tạo các methyl ester của acid béo và glycerol theo phương trình tổng quát sau:\n"
        "   (RCOO)~3~C~3~H~5~ + 3CH~3~OH -(xúc tác, t°)-> 3RCOOCH~3~ + C~3~H~5~(OH)~3~\n"
        "Cơ sở sử dụng 100 kg dầu ăn đã qua sử dụng chứa 88,4% triolein theo khối lượng, phần còn lại là tạp chất không tham gia phản ứng và không tạo biodiesel. Cho lượng dầu này phản ứng với methanol dư, hiệu suất chuyển hóa triolein thành methyl oleate đạt 90%. Khối lượng biodiesel thu được là bao nhiêu, biết biodiesel trong bài toán chỉ gồm methyl oleate?"
    )
    render_options_no_table(
        doc,
        ["*A. 79,92 kg.*", "B. 79,56 kg.", "C. 88,80 kg.", "D. 26,64 kg."] if is_teacher else ["A. 79,92 kg.", "B. 79,56 kg.", "C. 88,80 kg.", "D. 26,64 kg."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "• Khối lượng triolein ban đầu: m~triolein~ = 100 × 88,4% = 88,4 kg.\n"
            "• Phân tử khối của triolein (C~17~H~33~COO)~3~C~3~H~5~: M = 884 g/mol.\n"
            "  => Số mol triolein: n = 88,4 / 884 = 0,1 kmol = 100 mol.\n"
            "• Phương trình: (C~17~H~33~COO)~3~C~3~H~5~ + 3CH~3~OH -> 3C~17~H~33~COOCH~3~ + C~3~H~5~(OH)~3~\n"
            "• Phân tử khối methyl oleate C~17~H~33~COOCH~3~: M = 296 g/mol.\n"
            "  Theo phương trình: n~este (LT)~ = 3 × n~triolein~ = 0,3 kmol.\n"
            "  => m~este (LT)~ = 0,3 × 296 = 88,8 kg.\n"
            "• Do hiệu suất phản ứng H = 90%, khối lượng biodiesel thực tế thu được là:\n"
            "  m = 88,8 × 90% = *79,92 kg*."
        )

    # Câu 13
    p13 = doc.add_paragraph()
    render_question_prompt(
        p13,
        "Câu 13. Trong sản xuất thực phẩm, dầu thực vật có thể được chế biến thành chất béo có độ rắn cao hơn, dùng làm nguyên liệu sản xuất bơ thực vật. Quá trình hóa học nào sau đây giúp thực hiện sự chuyển đổi này?"
    )
    render_options_no_table(
        doc,
        [
            "*A. Chất béo lỏng chứa gốc acid béo không no cộng H~2~ (có xúc tác, nhiệt độ thích hợp) tạo thành chất béo chứa gốc acid béo no ở dạng rắn.*",
            "B. Thủy phân chất béo trong dầu thực vật thành glycerol và acid béo.",
            "C. Xà phòng hóa chất béo trong dầu thực vật bằng dung dịch NaOH.",
            "D. Oxi hóa chất béo trong dầu thực vật bằng oxygen không khí."
        ] if is_teacher else [
            "A. Chất béo lỏng chứa gốc acid béo không no cộng H~2~ (có xúc tác, nhiệt độ thích hợp) tạo thành chất béo chứa gốc acid béo no ở dạng rắn.",
            "B. Thủy phân chất béo trong dầu thực vật thành glycerol và acid béo.",
            "C. Xà phòng hóa chất béo trong dầu thực vật bằng dung dịch NaOH.",
            "D. Oxi hóa chất béo trong dầu thực vật bằng oxygen không khí."
        ],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Phản ứng hydrogen hóa chất béo lỏng (chứa gốc acid béo không no như oleic, linoleic) xúc tác Ni, đun nóng chuyển thành chất béo rắn (chứa gốc no stearic), được dùng trong công nghiệp sản xuất bơ thực vật (margarine)."
        )

    # Câu 14
    p14 = doc.add_paragraph()
    render_question_prompt(
        p14,
        "Câu 14. Dung dịch glucose 5% được dùng để truyền tĩnh mạch nhằm cung cấp năng lượng nhanh cho bệnh nhân. Glucose thuộc loại carbohydrate nào sau đây?"
    )
    render_options_no_table(
        doc,
        ["A. Disaccharide.", "B. Lipid.", "C. Polysaccharide.", "*D. Monosaccharide.*"] if is_teacher else ["A. Disaccharide.", "B. Lipid.", "C. Polysaccharide.", "D. Monosaccharide."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "Glucose là monosaccharide (carbohydrate đơn giản nhất không bị thủy phân thành phân tử nhỏ hơn)."
        )

    # Câu 15
    p15 = doc.add_paragraph()
    render_question_prompt(
        p15,
        "Câu 15. Ethanol sinh học (cồn sinh học) được dùng làm nhiên liệu sinh học, được sản xuất bằng cách lên men các nguyên liệu chứa tinh bột. Từ 1 tấn sắn khô chứa 81% tinh bột về khối lượng, với hiệu suất của toàn bộ quá trình là 70%, khối lượng ethanol thu được là"
    )
    render_options_no_table(
        doc,
        ["A. 230 kg.", "B. 414 kg.", "*C. 322 kg.*", "D. 460 kg."] if is_teacher else ["A. 230 kg.", "B. 414 kg.", "C. 322 kg.", "D. 460 kg."],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "• Sơ đồ sản xuất: (C~6~H~10~O~5~)~n~ -> n C~6~H~12~O~6~ -> 2n C~2~H~5~OH + 2n CO~2~\n"
            "  Tỉ lệ: 162 kg tinh bột -> 2 × 46 = 92 kg C~2~H~5~OH.\n"
            "• Khối lượng tinh bột trong 1 tấn sắn khô: m = 1000 × 81% = 810 kg.\n"
            "• Khối lượng ethanol lý thuyết: m~LT~ = (810 × 92) / 162 = 460 kg.\n"
            "• Vì hiệu suất quá trình là 70%, khối lượng ethanol thực tế thu được là:\n"
            "  m~TT~ = 460 × 70% = *322 kg*."
        )

    # Câu 16
    p16 = doc.add_paragraph()
    render_question_prompt(
        p16,
        "Câu 16. Tiến hành thí nghiệm như hình vẽ sau: nhỏ một giọt dung dịch methylamine 0,1 M lên mẩu giấy quỳ tím đặt trên mặt kính đồng hồ (1), nhỏ một giọt dung dịch aniline 0,1 M lên mẩu giấy quỳ tím đặt trên mặt kính đồng hồ (2)."
    )
    p_img16 = doc.add_paragraph()
    p_img16.alignment = WD_ALIGN_PARAGRAPH.CENTER
    img16_path = os.path.join(IMG_DIR, "hinh_cau16_quy_tim.png")
    if os.path.exists(img16_path):
        p_img16.add_run().add_picture(img16_path, width=Inches(4.5))

    p_prompt16_sub = doc.add_paragraph()
    p_prompt16_sub.paragraph_format.space_before = Pt(2)
    p_prompt16_sub.paragraph_format.space_after = Pt(3)
    render_rich_text(p_prompt16_sub, "Hiện tượng quan sát được là", base_size=10.5)

    render_options_no_table(
        doc,
        [
            "A. cả hai mẩu giấy quỳ tím đều chuyển sang màu xanh.",
            "*B. giấy quỳ tím (1) chuyển sang màu xanh, giấy quỳ tím (2) không đổi màu.*",
            "C. giấy quỳ tím (1) không đổi màu, giấy quỳ tím (2) chuyển sang màu xanh.",
            "D. cả hai mẩu giấy quỳ tím đều không đổi màu."
        ] if is_teacher else [
            "A. cả hai mẩu giấy quỳ tím đều chuyển sang màu xanh.",
            "B. giấy quỳ tím (1) chuyển sang màu xanh, giấy quỳ tím (2) không đổi màu.",
            "C. giấy quỳ tím (1) không đổi màu, giấy quỳ tím (2) chuyển sang màu xanh.",
            "D. cả hai mẩu giấy quỳ tím đều không đổi màu."
        ],
        is_teacher=is_teacher
    )
    if is_teacher:
        add_solution_no_table(
            doc,
            "• Methylamine có tính base mạnh hơn ammonia do nhóm alkyl (+I) đẩy electron làm tăng mật độ electron trên nguyên tử nitrogen, nên làm quỳ tím chuyển sang màu xanh.\n"
            "• Aniline có gốc phenyl (-I, -C) hút electron mạnh làm giảm mật độ electron trên nitrogen, tính base rất yếu (yếu hơn ammonia), không làm đổi màu giấy quỳ tím.\n"
            "=> Kết luận: Giấy quỳ tím (1) hóa xanh, giấy quỳ tím (2) không đổi màu."
        )

    # ==========================================
    # PHẦN II: TRẮC NGHIỆM ĐÚNG - SAI
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI (2,0 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1 (Phần II)
    p2_c1 = doc.add_paragraph()
    render_question_prompt(
        p2_c1,
        "Câu 1. Tinh bột có nhiều trong gạo, ngô, khoai; cellulose là thành phần chính của sợi bông và có trong gỗ. Hai chất này có vai trò quan trọng trong đời sống nhưng khác nhau về cấu tạo và tính chất. Hãy xác định các phát biểu sau là đúng hay sai:"
    )
    
    p2_c1_a = doc.add_paragraph()
    render_statement(p2_c1_a, "a) Tinh bột và cellulose đều là polysaccharide có công thức tổng quát (C~6~H~10~O~5~)~n~ nên có cấu tạo hoàn toàn giống nhau.", ans="Sai" if is_teacher else None, is_teacher=is_teacher)
    
    p2_c1_b = doc.add_paragraph()
    render_statement(p2_c1_b, "b) Trong cơ thể người, tinh bột được thủy phân qua các sản phẩm trung gian, cuối cùng tạo glucose. Glucose được hấp thu vào máu, cung cấp năng lượng hoặc chuyển thành glycogen dự trữ.", ans="Đúng" if is_teacher else None, is_teacher=is_teacher)
    
    p2_c1_c = doc.add_paragraph()
    render_statement(p2_c1_c, "c) Tinh bột là chất rắn màu trắng, tan tốt trong nước lạnh.", ans="Sai" if is_teacher else None, is_teacher=is_teacher)
    
    p2_c1_d = doc.add_paragraph()
    render_statement(p2_c1_d, "d) Hồ tinh bột đã trộn với dung dịch iodine mất màu khi đun nóng nhẹ và xuất hiện lại màu xanh tím khi để nguội vì tinh bột đã thủy phân hoàn toàn thành glucose.", ans="Sai" if is_teacher else None, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(
            doc,
            "• a) Sai: Tinh bột cấu tạo từ các gốc α-glucose (gồm amylose mạch thẳng và amylopectin mạch phân nhánh), còn cellulose cấu tạo từ các gốc β-glucose liên kết với nhau bằng liên kết β-1,4-glycoside tạo thành chuỗi mạch kéo dài, không phân nhánh. Cấu tạo của chúng hoàn toàn khác nhau.\n"
            "• b) Đúng: Trong cơ thể người có enzyme amylase thủy phân tinh bột qua dextrin, maltose thành glucose. Glucose hấp thu vào máu đi nuôi cơ thể, phần dư thừa được tổng hợp thành glycogen dự trữ tại gan và cơ bắp.\n"
            "• c) Sai: Tinh bột không tan trong nước lạnh. Trong nước nóng từ 65 °C trở lên, các hạt tinh bột ngậm nước, trương phồng tạo thành dung dịch keo gọi là hồ tinh bột.\n"
            "• d) Sai: Màu xanh tím biến mất khi đun nóng là do các phân tử iodine bị giải phóng khỏi cấu trúc xoắn của amylose. Khi để nguội, phân tử amylose cuộn xoắn trở lại giữ chặt iodine làm màu xanh tím xuất hiện trở lại. Tinh bột không hề bị thủy phân chỉ bằng việc đun nóng nhẹ."
        )

    # Câu 2 (Phần II)
    p2_c2 = doc.add_paragraph()
    render_question_prompt(
        p2_c2,
        "Câu 2. Một nhóm học sinh tìm hiểu tính chất của protein trong lòng trắng trứng bằng thí nghiệm được mô tả trong hình dưới đây, kết quả thu được dung dịch X."
    )
    p_img_p2_c2 = doc.add_paragraph()
    p_img_p2_c2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    img_p2_c2_path = os.path.join(IMG_DIR, "hinh_p2_cau2_biuret.png")
    if os.path.exists(img_p2_c2_path):
        p_img_p2_c2.add_run().add_picture(img_p2_c2_path, width=Inches(4.8))

    p_p2_c2_sub = doc.add_paragraph()
    p_p2_c2_sub.paragraph_format.space_before = Pt(2)
    p_p2_c2_sub.paragraph_format.space_after = Pt(3)
    render_rich_text(p_p2_c2_sub, "Hãy xác định các phát biểu sau là đúng hay sai:", base_size=10.5)

    p2_c2_a = doc.add_paragraph()
    render_statement(p2_c2_a, "a) Nếu thay dung dịch lòng trắng trứng bằng dung dịch glycine thì vẫn xuất hiện màu tím đặc trưng của phản ứng biuret.", ans="Sai" if is_teacher else None, is_teacher=is_teacher)

    p2_c2_b = doc.add_paragraph()
    render_statement(p2_c2_b, "b) Các protein dạng hình sợi như keratin (trong tóc, móng), collagen (trong da, sụn) không tan trong nước.", ans="Đúng" if is_teacher else None, is_teacher=is_teacher)

    p2_c2_c = doc.add_paragraph()
    render_statement(p2_c2_c, "c) Dung dịch X có màu tím.", ans="Đúng" if is_teacher else None, is_teacher=is_teacher)

    p2_c2_d = doc.add_paragraph()
    render_statement(p2_c2_d, "d) Khi nấu ăn người ta dùng nước ép dứa tươi để làm mềm thịt vì enzyme bromelain trong nước ép dứa tươi có khả năng xúc tác phản ứng thủy phân protein.", ans="Đúng" if is_teacher else None, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(
            doc,
            "• a) Sai: Glycine là một amino acid đơn lẻ, không chứa liên kết peptide (-CO-NH-) nên không tham gia phản ứng màu biuret.\n"
            "• b) Đúng: Protein dạng hình sợi (như keratin của tóc, móng, sừng; myosin của cơ bắp; collagen của da, sụn) hoàn toàn không tan trong nước và rất bền vững.\n"
            "• c) Đúng: Albumin trong lòng trắng trứng là protein dạng hình cầu, tan trong nước và có nhiều liên kết peptide, tạo phức chất màu tím đặc trưng với ion Cu^2+^ trong môi trường kiềm (phản ứng biuret).\n"
            "• d) Đúng: Trong quả dứa tươi có chứa enzyme bromelain (thuộc nhóm protease), có khả năng xúc tác thủy phân các liên kết peptide trong chuỗi protein của thịt động vật, giúp thịt nhanh mềm và dễ tiêu hóa."
        )

    # ==========================================
    # PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN III. CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Điền kết quả trả lời vào phiếu làm bài."
    )

    # Câu 1 (Phần III)
    p3_c1 = doc.add_paragraph()
    render_question_prompt(
        p3_c1,
        "Câu 1. Saccharose được dùng phổ biến làm đường ăn, còn maltose có trong sản phẩm thủy phân tinh bột. Xét các đặc điểm cấu tạo sau:\n"
        "(1) Trong cấu tạo phân tử có hai đơn vị monosaccharide liên kết với nhau.\n"
        "(2) Có 8 nhóm -OH trong phân tử.\n"
        "(3) Chỉ được tạo thành từ các đơn vị glucose.\n"
        "(4) Có 1 nhóm -OH hemiacetal tự do.\n"
        "Có bao nhiêu đặc điểm đúng với cả saccharose và maltose ở dạng vòng?"
    )
    if not is_teacher:
        p_ans1 = doc.add_paragraph()
        render_rich_text(p_ans1, "*Trả lời:* ..................................................................................................................", base_size=10.5)
    else:
        add_solution_no_table(
            doc,
            "• (1) Đúng: Cả hai đều là disaccharide, cấu tạo từ hai gốc monosaccharide liên kết qua liên kết glycoside.\n"
            "• (2) Đúng: Công thức phân tử của cả hai là C~12~H~22~O~11~, ở dạng mạch vòng đều chứa đúng 8 nhóm -OH.\n"
            "• (3) Sai: Maltose tạo từ hai gốc α-glucose, còn saccharose tạo từ một gốc α-glucose và một gốc β-fructose.\n"
            "• (4) Sai: Ở dạng vòng, saccharose liên kết bằng C1 của glucose và C2 của fructose nên không còn nhóm -OH hemiacetal hay hemiketal tự do; chỉ có maltose còn 1 nhóm -OH hemiacetal tự do.\n"
            "=> Có 2 đặc điểm đúng ((1) và (2)).\n"
            "**Đáp số: 2**"
        )

    # Câu 2 (Phần III)
    p3_c2 = doc.add_paragraph()
    render_question_prompt(
        p3_c2,
        "Câu 2. Xà phòng và chất giặt rửa tổng hợp được sử dụng phổ biến trong đời sống. Xét các phát biểu sau về phương pháp sản xuất và cách sử dụng hai loại sản phẩm này:\n"
        "a) Xà phòng có thể được sản xuất bằng cách đun nóng chất béo với dung dịch NaOH hoặc KOH.\n"
        "b) Phương pháp chủ yếu để sản xuất chất giặt rửa tổng hợp là đun nóng chất béo với dung dịch kiềm.\n"
        "c) Trong sản xuất chất giặt rửa tổng hợp, có thể sulfo hoá alkylbenzene rồi trung hoà sản phẩm bằng NaOH để thu được sodium alkylbenzenesulfonate.\n"
        "d) Khi giặt bằng nước cứng, xà phòng tạo kết tủa với ion Ca^2+^, Mg^2+^ nên bị giảm tác dụng giặt rửa.\n"
        "e) Cần sử dụng xà phòng và chất giặt rửa tổng hợp đúng mục đích, đúng liều lượng theo hướng dẫn, tránh để sản phẩm tiếp xúc với mắt.\n"
        "Có bao nhiêu phát biểu đúng?"
    )
    if not is_teacher:
        p_ans2 = doc.add_paragraph()
        render_rich_text(p_ans2, "*Trả lời:* ..................................................................................................................", base_size=10.5)
    else:
        add_solution_no_table(
            doc,
            "• a) Đúng: Xà phòng được sản xuất bằng phản ứng xà phòng hóa chất béo với kiềm NaOH hoặc KOH.\n"
            "• b) Sai: Đun nóng chất béo với kiềm là sản xuất xà phòng. Chất giặt rửa tổng hợp chủ yếu được sản xuất từ hydrocarbon của dầu mỏ.\n"
            "• c) Đúng: Alkylbenzene được sulfo hóa thành alkylbenzenesulfonic acid rồi trung hòa bằng NaOH để tạo muối sodium alkylbenzenesulfonate (chất hoạt động bề mặt tổng hợp).\n"
            "• d) Đúng: Trong nước cứng, các anion gốc acid béo kết hợp với Ca^2+^, Mg^2+^ tạo thành muối không tan (kết tủa) làm giảm tác dụng giặt rửa và bám bẩn vải.\n"
            "• e) Đúng: Cần sử dụng đúng hướng dẫn để bảo vệ da tay, quần áo và sức khỏe.\n"
            "=> Có 4 phát biểu đúng (a, c, d, e).\n"
            "**Đáp số: 4**"
        )

    # Câu 3 (Phần III)
    p3_c3 = doc.add_paragraph()
    render_question_prompt(
        p3_c3,
        "Câu 3. Cho Cu(OH)~2~ mới tạo vào dung dịch Gly-Ala-Gly trong môi trường kiềm, lắc đều thấy xuất hiện màu tím. Xét các nhận định:\n"
        "(1) Gly-Ala-Gly có hai liên kết peptide trong phân tử.\n"
        "(2) Màu tím xuất hiện do peptide tạo phức với ion Cu^2+^ trong môi trường kiềm.\n"
        "(3) Thay Gly-Ala-Gly bằng Gly-Ala vẫn thu được màu tím đặc trưng tương tự.\n"
        "(4) Sau khi thuỷ phân hoàn toàn Gly-Ala-Gly, hỗn hợp sản phẩm không còn cho phản ứng màu biuret.\n"
        "Có bao nhiêu nhận định đúng?"
    )
    if not is_teacher:
        p_ans3 = doc.add_paragraph()
        render_rich_text(p_ans3, "*Trả lời:* ..................................................................................................................", base_size=10.5)
    else:
        add_solution_no_table(
            doc,
            "• (1) Đúng: Tripeptide cấu tạo từ 3 gốc amino acid nên có số liên kết peptide = 3 - 1 = 2 liên kết -CO-NH-.\n"
            "• (2) Đúng: Các liên kết peptide tham gia phối trí tạo phức chất chelate màu tím đặc trưng với ion Cu^2+^.\n"
            "• (3) Sai: Gly-Ala là dipeptide (chỉ có 1 liên kết peptide), không có khả năng phản ứng màu biuret.\n"
            "• (4) Đúng: Khi thủy phân hoàn toàn, thu được hỗn hợp các amino acid tự do (glycine và alanine). Các amino acid này không còn liên kết peptide nên không cho phản ứng màu biuret.\n"
            "=> Có 3 nhận định đúng: (1), (2), (4).\n"
            "**Đáp số: 3**"
        )

    # Câu 4 (Phần III)
    p3_c4 = doc.add_paragraph()
    render_question_prompt(
        p3_c4,
        "Câu 4. Tiến hành thí nghiệm nhỏ aniline loãng vào nước bromine dư như hình sau:"
    )
    p_img_p3_c4 = doc.add_paragraph()
    p_img_p3_c4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    img_p3_c4_path = os.path.join(IMG_DIR, "hinh_p3_cau4_aniline_bromine.png")
    if os.path.exists(img_p3_c4_path):
        p_img_p3_c4.add_run().add_picture(img_p3_c4_path, width=Inches(3.8))

    p_p3_c4_sub = doc.add_paragraph()
    p_p3_c4_sub.paragraph_format.space_before = Pt(2)
    p_p3_c4_sub.paragraph_format.space_after = Pt(3)
    render_rich_text(
        p_p3_c4_sub,
        "Xét các nhận định:\n"
        "(1) Xuất hiện kết tủa trắng là 2,4,6-tribromoaniline.\n"
        "(2) Do ảnh hưởng của nhóm -NH~2~ lên vòng benzene nên aniline tham gia phản ứng thế nguyên tử hydrogen dễ hơn benzene.\n"
        "(3) Nếu thay dung dịch aniline trong thí nghiệm trên bằng dung dịch alanine thì vẫn thu được kết tủa trắng.\n"
        "(4) Cho 0,01 mol aniline phản ứng hoàn toàn với nước bromine dư thì thấy 4,8 gam bromine phản ứng.\n"
        "Có bao nhiêu nhận định đúng?",
        base_size=10.5
    )
    if not is_teacher:
        p_ans4 = doc.add_paragraph()
        render_rich_text(p_ans4, "*Trả lời:* ..................................................................................................................", base_size=10.5)
    else:
        add_solution_no_table(
            doc,
            "• Phương trình phản ứng: C~6~H~5~NH~2~ + 3Br~2~ -> C~6~H~2~Br~3~NH~2~↓ (trắng) + 3HBr\n"
            "• (1) Đúng: Kết tủa trắng tạo thành là 2,4,6-tribromoaniline.\n"
            "• (2) Đúng: Cặp electron không chia của nitrogen liên hợp với hệ electron pi của nhân thơm làm tăng mật độ electron ở các vị trí ortho và para, phản ứng thế xảy ra dễ dàng hơn benzene mà không cần xúc tác.\n"
            "• (3) Sai: Alanine (CH~3~CH(NH~2~)COOH) là amino acid no, không tác dụng với nước bromine tạo kết tủa trắng.\n"
            "• (4) Đúng: Theo phương trình, n~Br2~ = 3 × n~aniline~ = 3 × 0,01 = 0,03 mol.\n"
            "  Khối lượng bromine tham gia phản ứng là: m~Br2~ = 0,03 × 160 = *4,8 gam*.\n"
            "=> Có 3 nhận định đúng: (1), (2), (4).\n"
            "**Đáp số: 3**"
        )

    # ==========================================
    # PHẦN IV: TỰ LUẬN
    # ==========================================
    add_section_header(
        doc,
        title="PHẦN IV. TỰ LUẬN (2,0 điểm)",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trình bày chi tiết lời giải vào giấy làm bài thi."
    )

    # Câu 1 (Phần IV)
    p4_c1 = doc.add_paragraph()
    render_question_prompt(
        p4_c1,
        "Câu 1. (1,0 điểm) Để tìm hiểu tính chất hóa học của glucose tiến hành thí nghiệm như sau:\n"
        "- Bước 1: Cho vào mỗi ống nghiệm (1) và (2) lần lượt 3 mL dung dịch NaOH 10%, 0,6 mL dung dịch CuSO~4~ 5% rồi lắc đều. Tiếp tục thêm vào mỗi ống 3 mL dung dịch glucose 5% và lắc đều.\n"
        "- Bước 2: Giữ ống nghiệm (1) ở nhiệt độ thường, đun nóng nhẹ ống nghiệm (2) bằng đèn cồn trong vài phút. Cách bố trí thí nghiệm được mô tả trong hình."
    )
    p_img_p4_c1 = doc.add_paragraph()
    p_img_p4_c1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    img_p4_c1_path = os.path.join(IMG_DIR, "hinh_p4_cau1_glucose_cuoh2.png")
    if os.path.exists(img_p4_c1_path):
        p_img_p4_c1.add_run().add_picture(img_p4_c1_path, width=Inches(4.8))

    p_p4_c1_sub = doc.add_paragraph()
    p_p4_c1_sub.paragraph_format.space_before = Pt(2)
    p_p4_c1_sub.paragraph_format.space_after = Pt(3)
    render_rich_text(
        p_p4_c1_sub,
        "a) *(0,5 điểm)* Nêu hiện tượng quan sát được ở mỗi ống nghiệm. Mỗi hiện tượng đó chứng tỏ trong phân tử glucose có nhóm chức nào?\n"
        "b) *(0,5 điểm)* Viết phương trình hóa học của phản ứng xảy ra ở ống nghiệm (2).",
        base_size=10.5
    )
    if not is_teacher:
        add_student_writing_lines(doc, num_lines=10)
    else:
        add_solution_no_table(
            doc,
            "a) *Hiện tượng quan sát được và ý nghĩa chứng minh cấu tạo:*\n"
            "   • **Ống nghiệm (1) (ở nhiệt độ thường):** Kết tủa xanh lam Cu(OH)~2~ bị hòa tan, tạo thành dung dịch đồng nhất có màu xanh lam thẫm (phức chất đồng(II) - glucose).\n"
            "     -> Hiện tượng này chứng minh phân tử glucose có **nhiều nhóm hydroxy (-OH) kề nhau** (tính chất của polyalcohol).\n"
            "   • **Ống nghiệm (2) (khi đun nóng nhẹ):** Xuất hiện kết tủa màu đỏ gạch (Cu~2~O).\n"
            "     -> Hiện tượng này chứng minh phân tử glucose có **nhóm aldehyde (-CHO)** (thể hiện tính khử trong môi trường kiềm).\n\n"
            "b) *Phương trình hóa học của các phản ứng xảy ra ở ống nghiệm (2):*\n"
            "   • Phản ứng tạo Cu(OH)~2~ ban đầu:\n"
            "     CuSO~4~ + 2NaOH -> Cu(OH)~2~↓ + Na~2~SO~4~\n"
            "   • Phản ứng hòa tan Cu(OH)~2~ ở nhiệt độ thường:\n"
            "     2C~6~H~12~O~6~ + Cu(OH)~2~ -> (C~6~H~11~O~6~)~2~Cu + 2H~2~O\n"
            "   • Phản ứng oxi hóa nhóm -CHO bởi Cu(OH)~2~ trong kiềm khi đun nóng:\n"
            "     CH~2~OH[CHOH]~4~CHO + 2Cu(OH)~2~ + NaOH -(t°)-> CH~2~OH[CHOH]~4~COONa (sodium gluconate) + Cu~2~O↓ (đỏ gạch) + 3H~2~O\n"
            "     *(hoặc viết dạng gộp: C~6~H~12~O~6~ + 2Cu(OH)~2~ + NaOH -(t°)-> C~6~H~11~O~7~Na + Cu~2~O↓ + 3H~2~O)*"
        )

    # Câu 2 (Phần IV)
    p4_c2 = doc.add_paragraph()
    render_question_prompt(
        p4_c2,
        "Câu 2. (1,0 điểm) Trong phòng thí nghiệm, một nhóm học sinh tiến hành thí nghiệm cho tristearin tác dụng với dung dịch NaOH và xử lí sản phẩm theo quy trình minh họa dưới đây để tạo xà phòng (sản phẩm mang tính thực nghiệm và chưa được đánh giá chất lượng sử dụng):"
    )
    p_img_p4_c2 = doc.add_paragraph()
    p_img_p4_c2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    img_p4_c2_path = os.path.join(IMG_DIR, "hinh_p4_cau2_quy_trinh_xa_phong.png")
    if os.path.exists(img_p4_c2_path):
        p_img_p4_c2.add_run().add_picture(img_p4_c2_path, width=Inches(5.5))

    p_p4_c2_sub = doc.add_paragraph()
    p_p4_c2_sub.paragraph_format.space_before = Pt(2)
    p_p4_c2_sub.paragraph_format.space_after = Pt(3)
    render_rich_text(
        p_p4_c2_sub,
        "a) *(0,25 điểm)* Viết phương trình hóa học của phản ứng xà phòng hóa tristearin bằng NaOH.\n"
        "b) *(0,25 điểm)* Nêu vai trò của dung dịch NaCl bão hòa ở giai đoạn (2) trong thí nghiệm.\n"
        "c) *(0,5 điểm)* Một cơ sở sản xuất nhận đơn hàng gồm 48 bánh xà phòng, mỗi bánh có khối lượng 80 gam, trong đó sodium stearate chiếm 76,5% khối lượng, phần còn lại là các chất phụ gia. Cơ sở sử dụng tristearin nguyên chất phản ứng hoàn toàn với NaOH dư. Do hao hụt trong quá trình tách và xử lí sản phẩm, khối lượng sodium stearate thực tế thu được bằng 80% khối lượng tính theo lí thuyết. Tính khối lượng tristearin tối thiểu cần dùng để hoàn thành đơn hàng, theo đơn vị kg.",
        base_size=10.5
    )
    if not is_teacher:
        add_student_writing_lines(doc, num_lines=12)
    else:
        add_solution_no_table(
            doc,
            "a) *Phương trình hóa học xà phòng hóa tristearin:*\n"
            "   (C~17~H~35~COO)~3~C~3~H~5~ + 3NaOH -(t°)-> 3C~17~H~35~COONa + C~3~H~5~(OH)~3~\n\n"
            "b) *Vai trò của dung dịch NaCl bão hòa ở giai đoạn (2):*\n"
            "   Dung dịch NaCl bão hòa làm tăng khối lượng riêng (tỉ trọng) của dung dịch và làm giảm độ tan của muối sodium stearate, từ đó giúp xà phòng kết tinh lại và nổi lên trên bề mặt dung dịch, giúp cho việc vớt và lọc tách xà phòng diễn ra dễ dàng hơn.\n\n"
            "c) *Tính khối lượng tristearin tối thiểu:*\n"
            "   • Tổng khối lượng 48 bánh xà phòng cần hoàn thành:\n"
            "     m~xà phòng~ = 48 × 80 = 3840 gam = 3,84 kg.\n"
            "   • Khối lượng sodium stearate thực tế cần có trong đơn hàng:\n"
            "     m~thực tế (C17H35COONa)~ = 3840 × 76,5% = 2937,6 gam.\n"
            "   • Vì hiệu suất thu hồi sản phẩm qua tách/lọc là H = 80%, nên khối lượng sodium stearate cần tạo ra theo lí thuyết là:\n"
            "     m~lí thuyết (C17H35COONa)~ = 2937,6 / 80% = 3672 gam.\n"
            "   • Phân tử khối: M(C~17~H~35~COONa) = 17×12 + 35 + 44 + 23 = 306 g/mol.\n"
            "     M(tristearin) = (17×12 + 35 + 44)×3 + 41 = 283×3 + 41 = 890 g/mol.\n"
            "   • Số mol C~17~H~35~COONa theo lí thuyết: n = 3672 / 306 = 12 mol.\n"
            "   • Theo phương trình hóa học: n~tristearin~ = 1/3 × n~C17H35COONa~ = 12 / 3 = 4 mol.\n"
            "   • Khối lượng tristearin tối thiểu cần dùng là:\n"
            "     m~tristearin~ = 4 × 890 = 3560 gam = *3,56 kg*.\n"
            "   **Đáp số: 3,56 kg.**"
        )

    # Dòng kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(12)
    p_end.paragraph_format.space_after = Pt(12)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "---------- HẾT ----------", base_size=11, base_bold=True, base_color=COLOR_GRAY)

    # Lưu file
    suffix = "_GiaoVien.docx" if is_teacher else ".docx"
    file_name = f"De_Kiem_Tra_GK1_Hoa_12_THPT_So2_QuangTrung_Ma2102{suffix}"
    out_path = os.path.join(OUTPUT_DIR, file_name)
    doc.save(out_path)
    print(f"Exported successfully: {out_path}")

if __name__ == "__main__":
    build_exam_document(is_teacher=False)
    build_exam_document(is_teacher=True)
