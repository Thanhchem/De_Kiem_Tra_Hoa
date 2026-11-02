# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì 1 Hóa Học 10 - Mã Đề 5001
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_Hoa_10_Ma5001_HocSinh.docx
2. Bản Giáo viên: De_Kiem_Tra_Hoa_10_Ma5001_GiaoVien.docx
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
    COLOR_PRIMARY,
    COLOR_CORRECT,
    COLOR_SOLUTION,
    COLOR_GRAY
)

IMG_DIR = r"C:\Antigravity_Thanh\San_Pham\images_gk1_10_5001"

def add_centered_image(doc, img_name, width_cm=8.0, caption=None):
    img_path = os.path.join(IMG_DIR, img_name)
    if os.path.exists(img_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run()
        run.add_picture(img_path, width=Cm(width_cm))
        if caption:
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(1)
            p_cap.paragraph_format.space_after = Pt(4)
            render_rich_text(p_cap, f"_{caption}_", base_size=9.5, base_italic=True)

def style_data_table(table, col_widths, col_alignments):
    """Định dạng bảng số liệu khoa học chuẩn."""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(table.rows):
        # Tránh gãy dòng qua trang
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        if i == 0:
            trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

        for j, cell in enumerate(row.cells):
            cell.width = Cm(col_widths[j])
            p = cell.paragraphs[0]
            p.alignment = col_alignments[j]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            # Viền cell mỏng và padding
            tcPr = cell._tc.get_or_add_tcPr()
            tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>')
            tcPr.append(tcMar)
            # Nền header
            if i == 0:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F4F1"/>')
                tcPr.append(shd)

def build_exam_5001(is_teacher=False, output_path=None):
    doc = docx.Document()
    setup_header_footer(doc)

    # 1. Khối tiêu đề
    build_title_block(
        doc,
        school_line="SỞ GIÁO DỤC VÀ ĐÀO TẠO",
        exam_title="KIỂM TRA GIỮA KỲ 1 - NĂM HỌC 2026 – 2027",
        subject_line="MÔN: HÓA HỌC – KHỐI 10",
        num_pages="04",
        duration="45 phút",
        exam_code="5001",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Ga = 69,7; In = 114,8",
        is_teacher=is_teacher
    )

    # =========================================================================
    # PHẦN I
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        subtitle="Học sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi học sinh chỉ chọn một phương án. (3,0 điểm)"
    )

    # Câu 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 1. Ứng dụng nào sau đây thể hiện vai trò của hóa học trong lĩnh vực bảo vệ môi trường?")
    opts1 = [
        "A. Tổng hợp thuốc kháng sinh để điều trị bệnh nhiễm khuẩn.",
        "*B. Chế tạo vật liệu hấp phụ để loại bỏ ion kim loại nặng trong nước thải nhà máy dệt.*" if is_teacher else "B. Chế tạo vật liệu hấp phụ để loại bỏ ion kim loại nặng trong nước thải nhà máy dệt.",
        "C. Sản xuất xi măng phục vụ xây dựng nhà ở.",
        "D. Sản xuất phân đạm để tăng năng suất cây trồng."
    ]
    render_options_no_table(doc, opts1, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- A thuộc lĩnh vực y dược, chăm sóc sức khỏe.\n"
            "- B: Chế tạo vật liệu hấp phụ để xử lí các ion kim loại nặng trong nước thải công nghiệp là giải pháp trực tiếp trong bảo vệ môi trường.\n"
            "- C thuộc ngành sản xuất vật liệu xây dựng.\n"
            "- D thuộc ngành sản xuất phân bón nông nghiệp.\n"
            "➔ Đáp án đúng: B."
        )

    # Câu 2
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 2. Năm 2026, một nhóm nghiên cứu khoa học đã tận dụng phụ phẩm của cây cải dầu (bã ép, thân, lá, hoa) để chế tạo màng bao gói thực phẩm phân hủy sinh học. Màng được tạo từ chitosan, chiết xuất phenolic và hạt nano bạc, giúp kéo dài thời gian bảo quản rau quả tươi. Hoạt động nào sau đây thuộc phạm vi nghiên cứu của hóa học?")
    opts2 = [
        "A. Điều tra thói quen và mức độ sẵn sàng sử dụng bao gói xanh của người tiêu dùng.",
        "B. Khảo sát chi phí sản xuất và khả năng thương mại hóa màng bao gói trên thị trường.",
        "C. Nghiên cứu đặc điểm sinh trưởng của cây cải dầu nhằm nâng cao năng suất thu hoạch.",
        "*D. Xác định thành phần, cấu trúc của màng và sự biến đổi của màng khi bị phân hủy trong đất.*" if is_teacher else "D. Xác định thành phần, cấu trúc của màng và sự biến đổi của màng khi bị phân hủy trong đất."
    ]
    render_options_no_table(doc, opts2, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- Đối tượng nghiên cứu của hóa học là chất và sự biến đổi của chất. Hoạt động xác định thành phần hóa học, cấu trúc của vật liệu và cơ chế biến đổi hóa học/sinh hóa khi màng phân hủy thuộc phạm vi nghiên cứu của hóa học.\n"
            "- A và B thuộc lĩnh vực điều tra xã hội học và kinh tế học; C thuộc lĩnh vực nông nghiệp - sinh học.\n"
            "➔ Đáp án đúng: D."
        )

    # Câu 3
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 3. Hình vẽ sau mô tả thí nghiệm tìm ra hạt X là một trong những thành phần cấu tạo nên nguyên tử.")
    add_centered_image(doc, "cau3_ong_phong_tia.png", width_cm=7.5, caption="Thí nghiệm tìm ra hạt X")
    p_q3 = doc.add_paragraph()
    p_q3.paragraph_format.space_before = Pt(1)
    p_q3.paragraph_format.space_after = Pt(1)
    render_rich_text(p_q3, "Hạt X là", base_size=10.5)
    opts3 = [
        "A. proton.",
        "B. neutron.",
        "*C. electron.*" if is_teacher else "C. electron.",
        "D. neutron và proton."
    ]
    render_options_no_table(doc, opts3, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- Sơ đồ mô tả thí nghiệm ống phóng tia âm cực của J.J. Thomson (1897). Chùm tia âm cực bị lệch về phía bản cực dương khi đi qua điện trường, chứng tỏ chùm tia này mang điện tích âm. Hạt cấu tạo nên tia âm cực được gọi là electron (kí hiệu là e).\n"
            "➔ Đáp án đúng: C."
        )

    # Câu 4
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 4. Các nguyên tố trong bảng tuần hoàn hiện nay được xếp theo chiều tăng dần của")
    opts4 = [
        "*A. điện tích hạt nhân nguyên tử.*" if is_teacher else "A. điện tích hạt nhân nguyên tử.",
        "B. nguyên tử khối.",
        "C. số neutron.",
        "D. số lớp electron."
    ]
    render_options_no_table(doc, opts4, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- Nguyên tắc sắp xếp các nguyên tố trong bảng tuần hoàn hiện đại là: các nguyên tố được sắp xếp theo chiều tăng dần của điện tích hạt nhân nguyên tử (số hiệu nguyên tử Z).\n"
            "➔ Đáp án đúng: A."
        )

    # Câu 5
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 5. Năm 1911, Rutherford cho chùm hạt α bắn vào một lá vàng rất mỏng và quan sát thấy hầu hết các hạt α đi thẳng qua lá vàng. Kết quả này chứng tỏ")
    opts5 = [
        "A. hạt nhân nguyên tử mang điện tích âm.",
        "B. nguyên tử không chứa hạt mang điện.",
        "C. electron có khối lượng rất lớn.",
        "*D. nguyên tử có cấu tạo rỗng.*" if is_teacher else "D. nguyên tử có cấu tạo rỗng."
    ]
    render_options_no_table(doc, opts5, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- Thí nghiệm tán xạ hạt alpha qua lá vàng của Rutherford cho thấy hầu hết các hạt alpha xuyên thẳng không đổi hướng, chứng tỏ nguyên tử có cấu tạo rỗng (phần lớn thể tích nguyên tử là khoảng chân không, kích thước hạt nhân rất nhỏ so với kích thước toàn bộ nguyên tử).\n"
            "➔ Đáp án đúng: D."
        )

    # Câu 6
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 6. Hạt nhân nguyên tử chứa 28 hạt neutron là")
    opts6 = [
        "A. ^32_16S.",
        "*B. ^54_26Fe.*" if is_teacher else "B. ^54_26Fe.",
        "C. ^58_28Ni.",
        "D. ^32_15P."
    ]
    render_options_no_table(doc, opts6, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- Số neutron N = A - Z:\n"
            "  + Với ^32_16S: N = 32 - 16 = 16.\n"
            "  + Với ^54_26Fe: N = 54 - 26 = 28.\n"
            "  + Với ^58_28Ni: N = 58 - 28 = 30.\n"
            "  + Với ^32_15P: N = 32 - 15 = 17.\n"
            "➔ Đáp án đúng: B."
        )

    # Câu 7
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 7. Nguyên tố hóa học là những nguyên tử có cùng")
    opts7 = [
        "A. số khối.",
        "B. số neutron.",
        "C. số neutron và số proton.",
        "*D. số proton.*" if is_teacher else "D. số proton."
    ]
    render_options_no_table(doc, opts7, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- Định nghĩa: Nguyên tố hóa học là tập hợp các nguyên tử có cùng điện tích hạt nhân (có cùng số proton trong hạt nhân).\n"
            "➔ Đáp án đúng: D."
        )

    # Câu 8
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 8. Phát biểu nào sau đây không đúng?")
    opts8 = [
        "A. Phân lớp p có nhiều orbital hơn phân lớp s.",
        "B. Số electron tối đa trên phân lớp d gấp năm lần số electron tối đa trên phân lớp s.",
        "C. Số lượng electron tối đa trong một phân lớp luôn là một số chẵn.",
        "*D. Electron càng ở xa hạt nhân thì có năng lượng càng thấp.*" if is_teacher else "D. Electron càng ở xa hạt nhân thì có năng lượng càng thấp."
    ]
    render_options_no_table(doc, opts8, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- A đúng: Phân lớp s có 1 orbital, phân lớp p có 3 orbital.\n"
            "- B đúng: Phân lớp s tối đa 2 e, phân lớp d tối đa 10 e; 10 = 5 x 2.\n"
            "- C đúng: Số electron tối đa của các phân lớp s, p, d, f lần lượt là 2, 6, 10, 14, đều là số chẵn.\n"
            "- D sai: Electron càng ở xa hạt nhân thì liên kết với hạt nhân càng yếu và có mức năng lượng càng cao.\n"
            "➔ Đáp án đúng: D."
        )

    # Câu 9
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 9. Hình dưới biểu diễn hình dạng orbital nguyên tử (AO) nào sau đây?")
    add_centered_image(doc, "cau9_orbital_s.png", width_cm=4.2)
    opts9 = [
        "A. p~x~.",
        "B. p~z~.",
        "*C. s.*" if is_teacher else "C. s.",
        "D. p~y~."
    ]
    render_options_no_table(doc, opts9, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- Orbital s có dạng hình cầu đối xứng trong không gian tâm hạt nhân. Các orbital p~x~, p~y~, p~z~ có dạng hình số 8 nổi định hướng theo các trục tọa độ x, y, z tương ứng.\n"
            "➔ Đáp án đúng: C."
        )

    # Câu 10
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 10. Phát biểu nào sau đây đúng về chu kì 3 trong bảng tuần hoàn các nguyên tố hóa học?")
    opts10 = [
        "A. bắt đầu từ nguyên tố có Z = 19 và kết thúc ở nguyên tố có Z = 36.",
        "B. bắt đầu từ nguyên tố có Z = 11 và kết thúc ở nguyên tố có Z = 19.",
        "*C. bắt đầu từ nguyên tố có Z = 11 và kết thúc ở nguyên tố có Z = 18.*" if is_teacher else "C. bắt đầu từ nguyên tố có Z = 11 và kết thúc ở nguyên tố có Z = 18.",
        "D. bắt đầu từ nguyên tố có Z = 3 và kết thúc ở nguyên tố có Z = 10."
    ]
    render_options_no_table(doc, opts10, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- Chu kì 3 gồm 8 nguyên tố, bắt đầu từ Sodium (^23_11Na) và kết thúc ở Argon (^40_18Ar).\n"
            "➔ Đáp án đúng: C."
        )

    # Câu 11
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 11. Cho mô hình cấu tạo nguyên tử như sau:")
    add_centered_image(doc, "cau11_mo_hinh_nguyen_tu.png", width_cm=6.0)
    p_c11 = doc.add_paragraph()
    p_c11.paragraph_format.space_before = Pt(1)
    p_c11.paragraph_format.space_after = Pt(2)
    render_rich_text(p_c11, "Kí hiệu đúng của nguyên tử này là:", base_size=10.5)
    add_centered_image(doc, "cau11_o_nguyen_to.png", width_cm=14.0)
    opts11 = [
        "*A. Ô số 6 (Carbon - C).*" if is_teacher else "A. Ô số 6 (Carbon - C).",
        "B. Ô số 9 (Fluorine - F).",
        "C. Ô số 12 (Magnesium - Mg).",
        "D. Ô số 11 (Sodium - Na)."
    ]
    render_options_no_table(doc, opts11, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- Quan sát mô hình: Nguyên tử có 6 proton (hạt tròn đen) và 6 neutron (hạt tròn trắng) trong hạt nhân; lớp vỏ có 6 electron (2 electron ở lớp thứ nhất và 4 electron ở lớp thứ hai).\n"
            "- Số hiệu nguyên tử Z = số proton = 6, nguyên tố tương ứng là Carbon (kí hiệu C, ô số 6, nguyên tử khối 12,0107).\n"
            "➔ Đáp án đúng: A."
        )

    # Câu 12
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    render_question_prompt(p, "Câu 12. Hợp kim cobalt được dùng rộng rãi trong động cơ máy bay vì có độ bền cao ở nhiệt độ cao. Nguyên tử cobalt có cấu hình electron lớp ngoài cùng là 3d^7^4s^2^. Ở trạng thái cơ bản, số electron độc thân trong nguyên tử cobalt là")
    opts12 = [
        "A. 5.",
        "B. 2.",
        "C. 7.",
        "*D. 3.*" if is_teacher else "D. 3."
    ]
    render_options_no_table(doc, opts12, is_teacher=is_teacher)
    if is_teacher:
        add_solution_no_table(
            doc,
            "- Phân lớp 4s^2^ chứa 2 electron đã ghép đôi (không có electron độc thân).\n"
            "- Phân lớp 3d^7^ có 5 orbital, khi phân bố 7 electron theo quy tắc Hund: 5 electron đầu tiên điền vào 5 orbital với spin song song, 2 electron tiếp theo ghép đôi vào 2 orbital đầu tiên. Do đó còn lại 3 orbital mỗi orbital chứa 1 electron độc thân.\n"
            "- Vậy nguyên tử cobalt có 3 electron độc thân.\n"
            "➔ Đáp án đúng: D."
        )

    # =========================================================================
    # PHẦN II
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN II. Câu hỏi trắc nghiệm đúng sai.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu thí sinh chọn đúng hoặc sai. (2,0 điểm)"
    )

    # Câu 1 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. (1,0 điểm) Đồng thau là hợp kim của hai kim loại, được dùng để làm van nước, ổ khóa, tay nắm cửa và các nhạc cụ kèn đồng nhờ bền, dễ gia công và ít bị ăn mòn. Kết quả phân tích phổ khối lượng một mẫu đồng thau cho thấy mẫu chứa bốn loại nguyên tử chính X, Y, Z, T với thông tin như sau:")

    # Bảng dữ liệu X, Y, Z, T kết hợp ảnh
    tbl_p2_1 = doc.add_table(rows=5, cols=4)
    tbl_p2_1_widths = [2.8, 2.8, 2.8, 6.6]
    tbl_p2_1_aligns = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER]
    
    headers_1 = ["Nguyên tử", "Số proton", "Số neutron", "Tỉ lệ số nguyên tử trong tự nhiên của nguyên tố tương ứng (%)"]
    data_1 = [
        ["X", "29", "34", "69,15"],
        ["Y", "30", "36", "27,7"],
        ["Z", "29", "36", "30,85"],
        ["T", "30", "34", "49,2"]
    ]
    for j, h in enumerate(headers_1):
        p_c = tbl_p2_1.cell(0, j).paragraphs[0]
        render_rich_text(p_c, f"*{h}*", base_size=10, base_bold=True)
    for i, row in enumerate(data_1):
        for j, val in enumerate(val for val in row):
            p_c = tbl_p2_1.cell(i + 1, j).paragraphs[0]
            render_rich_text(p_c, val, base_size=10)
    style_data_table(tbl_p2_1, tbl_p2_1_widths, tbl_p2_1_aligns)

    add_centered_image(doc, "p2_cau1_dong_thau.png", width_cm=7.0, caption="Đồng thau (hợp kim của Cu và Zn)")

    stmts_1 = [
        ("a) X và T có cùng số neutron nên là đồng vị của nhau.", "Sai"),
        ("b) Bốn loại nguyên tử trong bảng thông tin trên thuộc hai nguyên tố hóa học.", "Đúng"),
        ("c) Nguyên tử khối trung bình của nguyên tố tạo nên X là 63,6 (làm tròn kết quả đến hàng phần mười).", "Đúng"),
        ("d) Nguyên tử khối trung bình của nguyên tố tạo nên T được tính bằng biểu thức (64 × 49,2 + 66 × 27,7) / 100.", "Sai")
    ]
    for stmt, ans in stmts_1:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(
            doc,
            "- a) Sai: Đồng vị là những nguyên tử có cùng số proton nhưng khác số neutron. X (Z = 29) và T (Z = 30) khác số proton nên thuộc hai nguyên tố khác nhau, không phải là đồng vị.\n"
            "- b) Đúng: Trong 4 loại nguyên tử chỉ có 2 giá trị số proton là Z = 29 (kim loại Đồng - Cu, gồm đồng vị X và Z) và Z = 30 (kim loại Kẽm - Zn, gồm đồng vị Y và T). Do đó chúng thuộc hai nguyên tố hóa học.\n"
            "- c) Đúng: Nguyên tố tạo nên X là Cu (Z = 29), gồm 2 đồng vị là X (A_X = 29 + 34 = 63, chiếm 69,15%) và Z (A_Z = 29 + 36 = 65, chiếm 30,85%).\n"
            "  A_tb = (63 × 69,15 + 65 × 30,85) / 100 = 6361,7 / 100 = 63,617 ≈ 63,6.\n"
            "- d) Sai: Nguyên tố tạo nên T là Zn (Z = 30), có 2 đồng vị nêu trong bảng là T (A = 64, chiếm 49,2%) và Y (A = 66, chiếm 27,7%). Tổng tỉ lệ của 2 đồng vị này là 49,2% + 27,7% = 76,9% (Zn trong tự nhiên còn chứa các đồng vị khác như ^67Zn, ^68Zn, ^70Zn). Biểu thức chia cho 100 là sai.\n"
            "➔ Kết luận: a) Sai | b) Đúng | c) Đúng | d) Sai."
        )

    # Câu 2 Phần II
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. (1,0 điểm) Một học sinh hỏi chatbot AI: “Mô hình nguyên tử hiện đại khác mô hình Rutherford – Bohr như thế nào?” Chatbot đưa ra các ý sau về mô hình nguyên tử hiện đại:")
    add_centered_image(doc, "p2_cau2_chatbot_ai.png", width_cm=7.2)

    stmts_2 = [
        ("a) Trong mô hình hiện đại, electron chuyển động trên các quỹ đạo tròn có bán kính xác định quanh hạt nhân.", "Sai"),
        ("b) Orbital nguyên tử là vùng không gian xung quanh hạt nhân mà xác suất tìm thấy electron khoảng 90%.", "Đúng"),
        ("c) Mọi orbital đều có dạng hình cầu, chỉ khác nhau về kích thước.", "Sai"),
        ("d) Mỗi orbital chứa tối đa 2 electron và hai electron này có chiều tự quay ngược nhau.", "Đúng")
    ]
    for stmt, ans in stmts_2:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        add_solution_no_table(
            doc,
            "- a) Sai: Electron chuyển động trên quỹ đạo tròn/bầu dục có bán kính xác định là nội dung của mô hình Rutherford – Bohr. Theo mô hình hiện đại, electron chuyển động rất nhanh trong không gian xung quanh hạt nhân không theo những quỹ đạo xác định.\n"
            "- b) Đúng: Orbital nguyên tử (AO) là khu vực không gian xung quanh hạt nhân mà tại đó xác suất tìm thấy electron là lớn nhất (khoảng 90%).\n"
            "- c) Sai: Chỉ có orbital s mới có dạng hình cầu; các orbital p có dạng hình số 8 nổi; các orbital d và f có hình dạng phức tạp hơn.\n"
            "- d) Đúng: Theo nguyên lí Pauli, mỗi orbital nguyên tử chỉ chứa tối đa 2 electron và 2 electron này có chiều tự quay quanh trục riêng ngược nhau (spin đối song).\n"
            "➔ Kết luận: a) Sai | b) Đúng | c) Sai | d) Đúng."
        )

    # =========================================================================
    # PHẦN III
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN III. Trắc nghiệm trả lời ngắn.",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4. (2,0 điểm)"
    )

    # Câu 1 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. Khối lượng của một nguyên tử magnesium là 39,8271.10^-27^ kg. Biết 1 amu = 1,6605.10^-27^ kg. Khối lượng nguyên tử magnesium tính theo đơn vị amu là bao nhiêu? _(Kết quả làm tròn đến hàng đơn vị.)_")
    if is_teacher:
        p_ans = doc.add_paragraph()
        render_rich_text(p_ans, "*Trả lời:* *24*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(
            doc,
            "Khối lượng nguyên tử magnesium theo đơn vị amu:\n"
            "m = (39,8271 × 10^-27 kg) / (1,6605 × 10^-27 kg/amu) ≈ 23,985 amu.\n"
            "Làm tròn đến hàng đơn vị ta được 24.\n"
            "➔ Đáp án: 24."
        )
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # Câu 2 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. Cho các phát biểu sau:")
    lines_c2 = [
        "(1) Điện tích của proton và electron có cùng độ lớn nhưng ngược dấu.",
        "(2) Có những nguyên tử không chứa neutron nào.",
        "(3) Khối lượng nguyên tử hầu hết tập trung ở hạt nhân nguyên tử.",
        "(4) Điện tích của proton và neutron có cùng độ lớn nhưng ngược dấu.",
        "(5) Khối lượng của proton và neutron xấp xỉ bằng nhau và lớn hơn nhiều khối lượng của electron."
    ]
    for line in lines_c2:
        p_l = doc.add_paragraph()
        p_l.paragraph_format.space_before = Pt(0)
        p_l.paragraph_format.space_after = Pt(1)
        p_l.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_l, line, base_size=10)

    p_c2_sub = doc.add_paragraph()
    p_c2_sub.paragraph_format.space_before = Pt(2)
    p_c2_sub.paragraph_format.space_after = Pt(2)
    render_rich_text(p_c2_sub, "Hãy sắp xếp các phát biểu *không đúng* trong các phát biểu trên theo số thứ tự tăng dần (ví dụ 12, 134, ...)", base_size=10.5)

    if is_teacher:
        p_ans = doc.add_paragraph()
        render_rich_text(p_ans, "*Trả lời:* *4*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(
            doc,
            "- (1) Đúng: Proton mang điện +1, electron mang điện -1, độ lớn điện tích đều bằng 1,602 × 10^-19 C.\n"
            "- (2) Đúng: Đồng vị hydrogen nhẹ nhất ^1_1H (protium) chỉ có 1 proton và 1 electron, không có neutron.\n"
            "- (3) Đúng: Do khối lượng của electron rất bé so với proton và neutron (~ 1/1836).\n"
            "- (4) Sai: Hạt neutron không mang điện (điện tích bằng 0), không phải có điện tích ngược dấu với proton.\n"
            "- (5) Đúng: m_p ≈ m_n ≈ 1 amu >> m_e ≈ 0,00055 amu.\n"
            "Như vậy phát biểu không đúng duy nhất là phát biểu (4).\n"
            "➔ Đáp án: 4."
        )
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # Câu 3 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. Nguyên tử của nguyên tố X có tổng số electron trên các phân lớp p là 11. Số hiệu nguyên tử của X là bao nhiêu?")
    if is_teacher:
        p_ans = doc.add_paragraph()
        render_rich_text(p_ans, "*Trả lời:* *17*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(
            doc,
            "- Phân lớp p được điền theo thứ tự mức năng lượng: 2p tối đa 6 electron, tiếp đến 3p.\n"
            "- Tổng số electron trên các phân lớp p là 11 = 6 (trên 2p) + 5 (trên 3p).\n"
            "- Cấu hình electron hoàn chỉnh của X: 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^5^.\n"
            "- Tổng số electron = 2 + 2 + 6 + 2 + 5 = 17 => Số hiệu nguyên tử Z = 17 (Chlorine - Cl).\n"
            "➔ Đáp án: 17."
        )
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # Câu 4 Phần III
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 4. Dựa vào bảng tuần hoàn các nguyên tố hóa học, cho thông tin về nguyên tố X như sau:")

    tbl_p3_4 = doc.add_table(rows=2, cols=5)
    tbl_p3_4_widths = [2.5, 3.5, 4.5, 2.5, 2.5]
    tbl_p3_4_aligns = [WD_ALIGN_PARAGRAPH.CENTER] * 5
    headers_4 = ["Kí hiệu", "Số hiệu nguyên tử", "Cấu hình e lớp ngoài cùng", "Chu kì", "Nhóm"]
    data_4 = ["X", "12", "3s^2^", "3", "IIA"]
    for j, h in enumerate(headers_4):
        p_c = tbl_p3_4.cell(0, j).paragraphs[0]
        render_rich_text(p_c, f"*{h}*", base_size=10, base_bold=True)
    for j, val in enumerate(data_4):
        p_c = tbl_p3_4.cell(1, j).paragraphs[0]
        render_rich_text(p_c, val, base_size=10)
    style_data_table(tbl_p3_4, tbl_p3_4_widths, tbl_p3_4_aligns)

    p_c4_sub = doc.add_paragraph()
    p_c4_sub.paragraph_format.space_before = Pt(3)
    p_c4_sub.paragraph_format.space_after = Pt(2)
    render_rich_text(p_c4_sub, "Nguyên tố Y thuộc cùng chu kì với X và thuộc nhóm VIIA. X và Y kết hợp với nhau tạo thành hợp chất có công thức XY~2~. Tổng số proton có trong một phân tử XY~2~ là bao nhiêu?", base_size=10.5)

    if is_teacher:
        p_ans = doc.add_paragraph()
        render_rich_text(p_ans, "*Trả lời:* *46*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        add_solution_no_table(
            doc,
            "- Nguyên tố X có Z_X = 12 (Magnesium - Mg).\n"
            "- Nguyên tố Y thuộc cùng chu kì 3 (có 3 lớp electron) và nhóm VIIA (có 7 electron lớp ngoài cùng dạng 3s^2^3p^5^). Cấu hình e đầy đủ của Y là 1s^2^ 2s^2^ 2p^6^ 3s^2^ 3p^5^ => Z_Y = 17 (Chlorine - Cl).\n"
            "- Phân tử hợp chất XY~2~ (chính là MgCl~2~) có tổng số proton:\n"
            "  ∑p = Z_X + 2 × Z_Y = 12 + 2 × 17 = 12 + 34 = 46.\n"
            "➔ Đáp án: 46."
        )
    else:
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(4)
        render_rich_text(p_ans, "*Trả lời:* ............................................................", base_size=10.5)

    # =========================================================================
    # PHẦN IV
    # =========================================================================
    add_section_header(
        doc,
        "PHẦN IV. TỰ LUẬN.",
        subtitle="HỌC SINH LÀM RA GIẤY THI TRẢ LỜI TỪ CÂU 1 ĐẾN CÂU 3 (3,0 điểm)"
    )

    # Câu 1 Phần IV
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 1. (0,5 điểm) Nguyên tố chlorine (Cl) có nguyên tử khối trung bình là 35,5. Biết chlorine gồm hai đồng vị bền là ^35^Cl và ^37^Cl. Để xác định phần trăm các đồng vị bền của chlorine trong tự nhiên, người ta dùng phương pháp phổ khối lượng.")
    
    p_c1_a = doc.add_paragraph()
    p_c1_a.paragraph_format.space_before = Pt(1)
    p_c1_a.paragraph_format.space_after = Pt(1)
    p_c1_a.paragraph_format.left_indent = Cm(0.5)
    render_rich_text(p_c1_a, "*a)* Hãy xác định phần trăm mỗi đồng vị bền của chlorine. _(0,25 điểm)_", base_size=10.5)

    p_c1_b = doc.add_paragraph()
    p_c1_b.paragraph_format.space_before = Pt(1)
    p_c1_b.paragraph_format.space_after = Pt(2)
    p_c1_b.paragraph_format.left_indent = Cm(0.5)
    render_rich_text(p_c1_b, "*b)* Hãy cho biết hình ảnh phổ khối lượng nào sau đây của chlorine là phù hợp nhất? Vì sao? _(0,25 điểm)_", base_size=10.5)

    add_centered_image(doc, "p4_cau1_pho_khoi_luong.png", width_cm=14.0)

    if is_teacher:
        add_solution_no_table(
            doc,
            "a) Gọi phần trăm số nguyên tử của đồng vị ^35^Cl là x (%), thì phần trăm của đồng vị ^37^Cl là (100 - x) (%).\n"
            "   Áp dụng công thức tính nguyên tử khối trung bình:\n"
            "   A_tb = [35 × x + 37 × (100 - x)] / 100 = 35,5\n"
            "   <=> 3700 - 2x = 3550  <=>  2x = 150  <=>  x = 75 (%)\n"
            "   Vậy đồng vị ^35^Cl chiếm 75% và đồng vị ^37^Cl chiếm 25%.\n"
            "b) Hình ảnh phổ khối lượng phù hợp nhất là Hình 1 (A).\n"
            "   Vì trên phổ khối lượng, tỉ lệ chiều cao pic ứng với tỉ lệ số nguyên tử của các đồng vị: ^35^Cl : ^37^Cl = 75% : 25% = 3 : 1. Trong Hình 1 (A), pic tại m/z = 35 có cường độ 100% (pic cơ bản) và pic tại m/z = 37 có cường độ khoảng 33,3%, hoàn toàn trùng khớp với tỉ lệ thực tế 3 : 1."
        )
    else:
        add_student_writing_lines(doc, num_lines=7)

    # Câu 2 Phần IV
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 2. (0,5 điểm) Các đám mây gây hiện tượng sấm sét được tạo nên bởi những hạt nước nhỏ li ti mang điện tích. Một phép đo thực nghiệm cho thấy một giọt nước hình cầu có đường kính 5.10^-3^ cm mang điện tích âm là -3,220.10^-17^ C. Điện tích âm này là do giọt nước nhận thêm electron.")
    
    p_eq = doc.add_paragraph()
    p_eq.paragraph_format.space_before = Pt(1)
    p_eq.paragraph_format.space_after = Pt(1)
    p_eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_eq, "Số phân tử nước trong giọt nước được tính theo biểu thức: *N~H2O~ = (4/3) . 3,14 . r^3^ . (D / 18) . N~A~*", base_size=10.5)

    p_note2 = doc.add_paragraph()
    p_note2.paragraph_format.space_before = Pt(1)
    p_note2.paragraph_format.space_after = Pt(2)
    render_rich_text(p_note2, "Trong đó D (g/cm^3^) là khối lượng riêng của nước, r (cm) là bán kính của giọt nước. Biết khối lượng riêng của nước là 1 g/cm^3^; điện tích của một electron là -1,602.10^-19^ C; số Avogadro N~A~ = 6,022.10^23^ mol^-1^.", base_size=10)

    p_c2_a = doc.add_paragraph()
    p_c2_a.paragraph_format.space_before = Pt(1)
    p_c2_a.paragraph_format.space_after = Pt(1)
    p_c2_a.paragraph_format.left_indent = Cm(0.5)
    render_rich_text(p_c2_a, "*a)* Tính số phân tử nước trong giọt nước. _(0,25 điểm)_", base_size=10.5)

    p_c2_b = doc.add_paragraph()
    p_c2_b.paragraph_format.space_before = Pt(1)
    p_c2_b.paragraph_format.space_after = Pt(2)
    p_c2_b.paragraph_format.left_indent = Cm(0.5)
    render_rich_text(p_c2_b, "*b)* Trung bình cứ x.10^12^ phân tử nước có trong giọt nước thì có 1 hạt electron dư. Tính giá trị của x _(Kết quả làm tròn đến hàng đơn vị)._ _(0,25 điểm)_", base_size=10.5)

    if is_teacher:
        add_solution_no_table(
            doc,
            "a) Bán kính giọt nước: r = d / 2 = (5 × 10^-3 cm) / 2 = 2,5 × 10^-3 cm.\n"
            "   Số phân tử nước trong giọt nước:\n"
            "   N~H2O~ = (4/3) × 3,14 × (2,5 × 10^-3)^3 × (1 / 18) × 6,022 × 10^23\n"
            "          = (4/3) × 3,14 × 15,625 × 10^-9 × (1 / 18) × 6,022 × 10^23\n"
            "          ≈ 2,1886 × 10^15 (phân tử nước).\n"
            "b) Số hạt electron dư giọt nước đã nhận thêm:\n"
            "   n_e = q / q_e = (-3,220 × 10^-17 C) / (-1,602 × 10^-19 C) ≈ 201 (hạt electron).\n"
            "   Trung bình số phân tử nước chứa 1 hạt electron dư:\n"
            "   Tỉ số = N~H2O~ / n_e = (2,1886 × 10^15) / 201 ≈ 1,0888 × 10^13 = 10,888 × 10^12 (phân tử).\n"
            "   Theo đề bài: Tỉ số = x × 10^12 => x ≈ 10,888.\n"
            "   Làm tròn đến hàng đơn vị: x = 11."
        )
    else:
        add_student_writing_lines(doc, num_lines=8)

    # Câu 3 Phần IV
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(2)
    render_question_prompt(p, "Câu 3. (2,0 điểm) Năm 2019 được chọn là Năm Quốc tế Bảng tuần hoàn các nguyên tố hóa học, kỉ niệm 150 năm Dmitri Ivanovich Mendeleev công bố bảng tuần hoàn (1869). Khi xây dựng bảng, Mendeleev đã để trống một số ô và dự đoán tính chất của các nguyên tố chưa được phát hiện. Một trong số đó là “eka-aluminium”, nguyên tố nằm ngay dưới nhôm trong cùng một nhóm.")

    p_hist = doc.add_paragraph()
    p_hist.paragraph_format.space_before = Pt(1)
    p_hist.paragraph_format.space_after = Pt(2)
    render_rich_text(p_hist, "Năm 1875, nhà hóa học người Pháp Paul-Émile Lecoq de Boisbaudran công bố tìm ra một nguyên tố mới. Nguyên tố này có nhiều tính chất trùng khớp với eka-aluminium, nhưng khối lượng riêng ông đo được lại khác xa giá trị Mendeleev tiên đoán. Mendeleev đề nghị ông đo lại. Kết quả đo lần sau, với mẫu đã được tinh chế, trùng khớp với dự đoán của Mendeleev. Một cách đơn giản để xác định khối lượng riêng của một mẫu kim loại là dùng bình tràn, như minh họa ở Hình 1.", base_size=10.5)

    add_centered_image(doc, "p4_cau3_binh_tran.png", width_cm=8.5, caption="Hình 1. Minh họa phương pháp đo khối lượng riêng bằng bình tràn (số liệu trong hình chỉ mang tính minh họa).")

    p_exp = doc.add_paragraph()
    p_exp.paragraph_format.space_before = Pt(1)
    p_exp.paragraph_format.space_after = Pt(2)
    render_rich_text(p_exp, "Áp dụng phương pháp này để mô phỏng tái hiện lại phép đo trong lịch sử, một nhóm học sinh cân một mẫu kim loại M (ở 20 °C) được khối lượng *35,46 g*. Sau đó, các em thả nhẹ mẫu vào một bình tràn chứa đầy nước và thu được *6,00 mL* nước tràn ra ống đong. Cho bảng số liệu sau:", base_size=10.5)

    tbl_p4_3 = doc.add_table(rows=3, cols=6)
    tbl_p4_3_widths = [4.5, 2.3, 2.3, 2.3, 2.3, 2.3]
    tbl_p4_3_aligns = [WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 5
    headers_3 = ["Kim loại", "Mg", "Al", "Cu", "Zn", "Ga"]
    row_z = ["Số hiệu nguyên tử (Z)", "12", "13", "29", "30", "31"]
    row_d = ["Khối lượng riêng (g/cm^3^)", "1,74", "2,70", "8,96", "7,13", "5,91"]
    for j, h in enumerate(headers_3):
        p_c = tbl_p4_3.cell(0, j).paragraphs[0]
        render_rich_text(p_c, f"*{h}*", base_size=10, base_bold=True)
    for j, val in enumerate(row_z):
        p_c = tbl_p4_3.cell(1, j).paragraphs[0]
        render_rich_text(p_c, val, base_size=10)
    for j, val in enumerate(row_d):
        p_c = tbl_p4_3.cell(2, j).paragraphs[0]
        render_rich_text(p_c, val, base_size=10)
    style_data_table(tbl_p4_3, tbl_p4_3_widths, tbl_p4_3_aligns)

    items_c3 = [
        ("a)", "Tính khối lượng riêng của mẫu kim loại M. Dựa vào bảng số liệu, cho biết M là kim loại nào? _(0,25 điểm)_"),
        ("b)", "Viết cấu hình electron của nguyên tử M. _(0,25 điểm)_"),
        ("c)", "Xác định vị trí của M trong bảng tuần hoàn (ô, chu kì, nhóm). _(0,5 điểm)_"),
        ("d)", "M là nguyên tố s, p, d hay f? _(0,25 điểm)_"),
        ("e)", "Nguyên tử M có bao nhiêu electron hóa trị? _(0,25 điểm)_"),
        ("f)", "Dựa vào vị trí của M và Al trong bảng tuần hoàn, hãy cho biết tính chất hoá học của M và Al có tương tự nhau hay không? Vì sao? _(0,25 điểm)_"),
        ("g)", "Biết nguyên tố In (Z = 49) nằm cùng nhóm với nguyên tố nhôm, ngay phía dưới nguyên tố “eka-aluminium”. Hãy xác định số electron hoá trị của In? _(0,25 điểm)_")
    ]
    for lbl, txt in items_c3:
        p_it = doc.add_paragraph()
        p_it.paragraph_format.space_before = Pt(1)
        p_it.paragraph_format.space_after = Pt(1)
        p_it.paragraph_format.left_indent = Cm(0.5)
        render_rich_text(p_it, f"*{lbl}* {txt}", base_size=10.5)

    if is_teacher:
        add_solution_no_table(
            doc,
            "a) Thể tích của mẫu kim loại M đúng bằng thể tích nước tràn ra:\n"
            "   V = 6,00 mL = 6,00 cm^3.\n"
            "   Khối lượng riêng của mẫu kim loại M:\n"
            "   D = m / V = 35,46 / 6,00 = 5,91 g/cm^3.\n"
            "   So sánh với bảng số liệu: D = 5,91 g/cm^3 tương ứng với kim loại Gallium (Ga, Z = 31).\n"
            "   Vậy M là kim loại Gallium (Ga).\n"
            "b) Nguyên tử Ga có Z = 31:\n"
            "   Cấu hình electron đầy đủ: 1s^2 2s^2 2p^6 3s^2 3p^6 3d^10 4s^2 4p^1  (hoặc viết gọn: [Ar] 3d^10 4s^2 4p^1).\n"
            "c) Vị trí của M (Ga) trong bảng tuần hoàn:\n"
            "   - Ô số: 31 (vì Z = 31).\n"
            "   - Chu kì: 4 (vì có 4 lớp electron).\n"
            "   - Nhóm: IIIA (vì là nguyên tố p, có 3 electron lớp ngoài cùng: 4s^2 4p^1).\n"
            "d) M là nguyên tố p (vì electron cuối cùng được điền vào phân lớp p, cụ thể là phân lớp 4p).\n"
            "e) Nguyên tử Ga có 3 electron hóa trị (gồm 2 electron 4s và 1 electron 4p; phân lớp 3d^10 đã bão hòa và thuộc lớp bên trong nên không tính vào electron hóa trị của nguyên tố nhóm A).\n"
            "f) Tính chất hóa học của M (Ga) và Al tương tự nhau.\n"
            "   Vì sao: Vì Ga và Al cùng thuộc nhóm IIIA trong bảng tuần hoàn, nguyên tử của chúng đều có cấu hình electron lớp ngoài cùng dạng ns^2 np^1 (có 3 electron hóa trị). Do đó, chúng đều có xu hướng nhường đi 3 electron khi tham gia phản ứng hóa học để đạt cấu hình electron bền vững của khí hiếm, thể hiện số oxi hóa đặc trưng là +3 trong các hợp chất.\n"
            "g) Nguyên tố In (Z = 49) nằm cùng nhóm IIIA với nhôm và gallium (eka-aluminium). Các nguyên tố trong cùng một nhóm A có số electron hóa trị bằng nhau và bằng số thứ tự của nhóm.\n"
            "   Do đó, nguyên tử In có 3 electron hóa trị (cấu hình electron của In: [Kr] 4d^10 5s^2 5p^1, có 3 e hóa trị ở phân lớp 5s^2 5p^1)."
        )
    else:
        add_student_writing_lines(doc, num_lines=16)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(4)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "*--- HẾT ---*", base_size=11, base_bold=True)

    if output_path:
        doc.save(output_path)
        print(f"Saved: {output_path}")

if __name__ == "__main__":
    out_dir = r"C:\Antigravity_Thanh\San_Pham"
    os.makedirs(out_dir, exist_ok=True)
    
    hs_path = os.path.join(out_dir, "De_Kiem_Tra_Hoa_10_Ma5001_HocSinh.docx")
    gv_path = os.path.join(out_dir, "De_Kiem_Tra_Hoa_10_Ma5001_GiaoVien.docx")
    
    build_exam_5001(is_teacher=False, output_path=hs_path)
    build_exam_5001(is_teacher=True, output_path=gv_path)
