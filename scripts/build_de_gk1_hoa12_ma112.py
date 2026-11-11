# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kỳ I - Môn: Hóa Học 12 - Mã Đề 112
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)

Xuất bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_12_Ma112_HocSinh.docx / .tex / .pdf
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_12_Ma112_GiaoVien.docx / .tex / .pdf
"""

import os
import sys
import subprocess
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
    add_p_border_bottom,
    COLOR_PRIMARY,
    COLOR_CORRECT,
    COLOR_SOLUTION,
    COLOR_GRAY
)

OUTPUT_DIR = r"C:\Antigravity_Thanh\San_Pham"
IMG_PATH = os.path.join(OUTPUT_DIR, "images_de112", "sodium_lauryl_sulfate.png")

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

def add_quick_answer_table_112(doc):
    """Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu Mã đề 112."""
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 112)*", base_size=11.5, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(3)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*PHẦN I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - 0,25 đ/câu)*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

    tbl_p1 = doc.add_table(rows=2, cols=12)
    tbl_p1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_p1.autofit = False

    p1_answers = [
        ("1", "A"), ("2", "D"), ("3", "A"), ("4", "B"),
        ("5", "B"), ("6", "D"), ("7", "A"), ("8", "A"),
        ("9", "A"), ("10", "A"), ("11", "C"), ("12", "D")
    ]

    for col_idx, (q_num, ans) in enumerate(p1_answers):
        cell_q = tbl_p1.cell(0, col_idx)
        cell_a = tbl_p1.cell(1, col_idx)
        cell_q.width = Cm(1.5)
        cell_a.width = Cm(1.5)

        set_cell_shading(cell_q, "E8F5E9")
        set_cell_border(cell_q, top=("single", "4", "0", "8BC390"), bottom=("single", "4", "0", "8BC390"),
                                left=("single", "4", "0", "8BC390"), right=("single", "4", "0", "8BC390"))
        set_cell_border(cell_a, top=("single", "4", "0", "8BC390"), bottom=("single", "4", "0", "8BC390"),
                                left=("single", "4", "0", "8BC390"), right=("single", "4", "0", "8BC390"))

        pq = cell_q.paragraphs[0]
        pq.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pq.paragraph_format.space_before = Pt(1)
        pq.paragraph_format.space_after = Pt(1)
        rq = pq.add_run(f"C{q_num}")
        set_run_style(rq, font_size=9.5, bold=True, color=COLOR_PRIMARY)

        pa = cell_a.paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pa.paragraph_format.space_before = Pt(1)
        pa.paragraph_format.space_after = Pt(1)
        ra = pa.add_run(ans)
        set_run_style(ra, font_size=10, bold=True, color=COLOR_CORRECT)

    # 2. PHẦN II & PHẦN III
    p_p23 = doc.add_paragraph()
    p_p23.paragraph_format.space_before = Pt(4)
    p_p23.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p23, "*PHẦN II & PHẦN III: Trắc nghiệm Đúng/Sai (2,0 điểm) & Trả lời ngắn (2,0 điểm)*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

    tbl_p23 = doc.add_table(rows=2, cols=6)
    tbl_p23.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_p23.autofit = False

    headers_p23 = [
        "Câu 1 (Đ/S)", "Câu 2 (Đ/S)",
        "Câu 1 (TLN)", "Câu 2 (TLN)", "Câu 3 (TLN)", "Câu 4 (TLN)"
    ]
    p23_answers = [
        "a: Đ | b: S\nc: S | d: S",
        "a: Đ | b: S\nc: Đ | d: Đ",
        "3", "7,25", "4370", "245"
    ]

    for col_idx in range(6):
        cell_h = tbl_p23.cell(0, col_idx)
        cell_a = tbl_p23.cell(1, col_idx)
        cell_h.width = Cm(3.0)
        cell_a.width = Cm(3.0)

        set_cell_shading(cell_h, "E8F5E9")
        set_cell_border(cell_h, top=("single", "4", "0", "8BC390"), bottom=("single", "4", "0", "8BC390"),
                                left=("single", "4", "0", "8BC390"), right=("single", "4", "0", "8BC390"))
        set_cell_border(cell_a, top=("single", "4", "0", "8BC390"), bottom=("single", "4", "0", "8BC390"),
                                left=("single", "4", "0", "8BC390"), right=("single", "4", "0", "8BC390"))

        ph = cell_h.paragraphs[0]
        ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ph.paragraph_format.space_before = Pt(1)
        ph.paragraph_format.space_after = Pt(1)
        rh = ph.add_run(headers_p23[col_idx])
        set_run_style(rh, font_size=9.5, bold=True, color=COLOR_PRIMARY)

        pa = cell_a.paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pa.paragraph_format.space_before = Pt(1)
        pa.paragraph_format.space_after = Pt(1)
        ra = pa.add_run(p23_answers[col_idx])
        set_run_style(ra, font_size=9.5, bold=True, color=COLOR_CORRECT)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(6)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


def build_docx(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="SỞ GIÁO DỤC VÀ ĐÀO TẠO",
        exam_title="KIỂM TRA GIỮA HỌC KỲ I",
        subject_line="MÔN: HÓA HỌC - LỚP 12",
        num_pages="03" if not is_teacher else "05",
        duration="45 phút",
        exam_code="112",
        atomic_masses="H=1, C=12, O=16, N=14, S=32, Na=23, Ag=108, Br=80",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table_112(doc)

    # ==========================
    # PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN
    # ==========================
    add_section_header(
        doc,
        "PHẦN I. (3 điểm) Câu trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    questions_p1 = [
        {
            "q": "Câu 1. Chất nào sau đây là ester :",
            "opts": [
                "*A. CH~3~COOCH~3~.*" if is_teacher else "A. CH~3~COOCH~3~.",
                "B. HCOOH.",
                "C. CH~3~CHO.",
                "D. CH~3~OH."
            ],
            "sol": "Ester là hợp chất hữu cơ được tạo thành khi thay thế nhóm -OH ở nhóm carboxyl (-COOH) của carboxylic acid bằng nhóm -OR'.\n- CH~3~COOCH~3~: methyl acetate (ester).\n- HCOOH: formic acid (carboxylic acid).\n- CH~3~CHO: acetaldehyde (aldehyde).\n- CH~3~OH: methanol (alcohol).\n➔ *Đáp án A.*"
        },
        {
            "q": "Câu 2. Chất nào sau đây là thành phần chính của chất giặt rửa tổng hợp?",
            "opts": [
                "A. C~3~H~5~(OH)~3~.",
                "B. CH~3~[CH~2~]~14~COONa.",
                "C. CH~3~[CH~2~]~16~COOK.",
                "*D. CH~3~[CH~2~]~11~C~6~H~4~SO~3~Na.*" if is_teacher else "D. CH~3~[CH~2~]~11~C~6~H~4~SO~3~Na."
            ],
            "sol": "Chất giặt rửa tổng hợp thường là muối sodium alkylbenzene sulfonate (như sodium dodecylbenzenesulfonate CH~3~[CH~2~]~11~C~6~H~4~SO~3~Na) hoặc muối sodium alkylsulfate.\n- C~3~H~5~(OH)~3~: glycerol.\n- CH~3~[CH~2~]~14~COONa và CH~3~[CH~2~]~16~COOK là muối của acid béo (thành phần của xà phòng).\n➔ *Đáp án D.*"
        },
        {
            "q": "Câu 3. Sodium lauryl sulfate là một trong những chất giặt rửa tổng hợp phổ biến có công thức cấu tạo như sau:",
            "has_image": True,
            "sub_prompt": "Đặc điểm của nhóm được khoanh tròn trong công thức trên là gì?",
            "opts": [
                "*A. Ưa nước.*" if is_teacher else "A. Ưa nước.",
                "B. Tạo kết tủa với nước cứng.",
                "C. Không phân cực.",
                "D. Kị nước."
            ],
            "sol": "Phân tử chất giặt rửa tổng hợp gồm 2 phần:\n- Đuôi hydrocarbon dài (chuỗi alkyl CH~3~-[CH~2~]~11~-): không phân cực, kị nước (ưa dầu mỡ).\n- Nhóm đầu phân cực (-OSO~3~^- Na^+ được khoanh tròn): mang điện tích, phân cực mạnh, có khả năng tạo liên kết hydro với nước nên là nhóm *ưa nước*.\n➔ *Đáp án A.*"
        },
        {
            "q": "Câu 4. Cho 6 carbohydrate sau: glucose, fructose, maltose, saccharose, tinh bột và cellulose. Có bao nhiêu carbohydrate đã cho thuộc nhóm monosaccharide ?",
            "opts": ["A. 1", "*B. 2*" if is_teacher else "B. 2", "C. 3", "D. 4"],
            "sol": "Trong 6 chất đã cho:\n- Monosaccharide gồm 2 chất: glucose và fructose.\n- Disaccharide gồm 2 chất: maltose và saccharose.\n- Polysaccharide gồm 2 chất: tinh bột và cellulose.\nVậy có 2 carbohydrate thuộc nhóm monosaccharide.\n➔ *Đáp án B.*"
        },
        {
            "q": "Câu 5. Saccharose được cấu tạo từ",
            "opts": [
                "A. hai đơn vị α-glucose.",
                "*B. một đơn vị α-glucose và một đơn vị β-fructose.*" if is_teacher else "B. một đơn vị α-glucose và một đơn vị β-fructose.",
                "C. hai đơn vị β-fructose.",
                "D. nhiều đơn vị α-glucose."
            ],
            "sol": "Phân tử saccharose được cấu tạo từ một đơn vị α-glucose và một đơn vị β-fructose, liên kết với nhau qua nguyên tử oxygen giữa C1 của gốc α-glucose và C2 của gốc β-fructose (liên kết α-1,2-glycoside).\n➔ *Đáp án B.*"
        },
        {
            "q": "Câu 6. Phát biểu nào sau đây đúng khi nói về tinh bột và cellulose?",
            "opts": [
                "A. Đều tham gia phản ứng tráng bạc.",
                "B. Đều có phản ứng màu với dung dịch iodine.",
                "C. Đều tan tốt trong nước lạnh.",
                "*D. Thủy phân hoàn toàn trong môi trường acid đều thu được glucose.*" if is_teacher else "D. Thủy phân hoàn toàn trong môi trường acid đều thu được glucose."
            ],
            "sol": "- Cả tinh bột (-C~6~H~10~O~5-)~n và cellulose (-C~6~H~10~O~5-)~n khi bị thủy phân hoàn toàn trong môi trường acid (đun nóng) đều thu được monosaccharide duy nhất là glucose: (-C~6~H~10~O~5-)~n + nH~2~O ➔ nC~6~H~12~O~6~.\n- Tinh bột và cellulose đều không có nhóm -OH hemiacetal tự do nên không tham gia phản ứng tráng bạc.\n- Chỉ có hồ tinh bột cho phản ứng màu xanh tím với iodine, cellulose thì không.\n- Cả hai đều không tan trong nước lạnh.\n➔ *Đáp án D.*"
        },
        {
            "q": "Câu 7. Cellulose thuộc loại polysaccharide, là thành phần chính tạo nên màng tế bào thực vật, có nhiều trong gỗ, bông nõn. Công thức của cellulose là",
            "opts": [
                "*A. (C~6~H~10~O~5~)~n~.*" if is_teacher else "A. (C~6~H~10~O~5~)~n~.",
                "B. C~12~H~22~O~11~.",
                "C. C~6~H~12~O~6~.",
                "D. C~2~H~4~O~2~."
            ],
            "sol": "Cellulose là một polymer thiên nhiên có công thức phân tử là (C~6~H~10~O~5~)~n~ (hay [C~6~H~7~O~2~(OH)~3~]~n~), gồm nhiều gốc β-glucose liên kết với nhau bằng liên kết β-1,4-glycoside tạo thành mạch kéo dài không phân nhánh.\n➔ *Đáp án A.*"
        },
        {
            "q": "Câu 8. Các amino acid tồn tại chủ yếu ở trạng thái ion lưỡng cực, do đó chúng",
            "opts": [
                "*A. có nhiệt độ nóng chảy cao và tan tốt trong nước.*" if is_teacher else "A. có nhiệt độ nóng chảy cao và tan tốt trong nước.",
                "B. có nhiệt độ nóng chảy cao và ít tan trong nước.",
                "C. dễ nóng chảy và tan tốt trong nước.",
                "D. dễ nóng chảy và ít tan trong nước."
            ],
            "sol": "Ở điều kiện thường, phân tử amino acid tồn tại chủ yếu ở dạng ion lưỡng cực H~3~N^+-CH(R)-COO^-. Do có cấu trúc ion tương tự như các hợp chất ion nên lực hút tĩnh điện giữa các phân tử rất mạnh, làm cho các amino acid là chất rắn có nhiệt độ nóng chảy cao và dễ tan trong nước.\n➔ *Đáp án A.*"
        },
        {
            "q": "Câu 9. Glutamic acid có vai trò quan trọng trong quá trình xây dựng cấu trúc tế bào của con người. Glutamic acid có điểm đẳng điện pI = 3,2 (pI là giá trị pH tại đó tổng điện tích trong amino acid bằng không). Khi đặt dung dịch glutamic acid ở pH = 6 trong điện trường, glutamic acid chủ yếu tồn tại ở dạng",
            "opts": [
                "*A. anion và di chuyển về cực dương.*" if is_teacher else "A. anion và di chuyển về cực dương.",
                "B. ion lưỡng cực và đứng yên.",
                "C. anion và di chuyển về cực âm.",
                "D. cation và di chuyển về cực âm."
            ],
            "sol": "- Điểm đẳng điện của glutamic acid là pI = 3,2.\n- Khi đặt ở môi trường pH = 6 > pI: Môi trường có tính kiềm hơn điểm đẳng điện, phân tử glutamic acid có xu hướng nhường proton H^+ từ nhóm carboxyl để tạo thành dạng tích điện âm (anion).\n- Trong điện trường, các hạt mang điện tích âm (anion) sẽ bị hút và di chuyển về phía cực dương (anode).\n➔ *Đáp án A.*"
        },
        {
            "q": "Câu 10. Cho 11,25 gam glycine tác dụng với dung dịch NaOH vừa đủ, cô cạn dung dịch sau phản ứng, thu được m gam muối khan. Giá trị của m là",
            "opts": [
                "*A. 14,55*" if is_teacher else "A. 14,55",
                "B. 13,32.",
                "C. 11,10.",
                "D. 16,65."
            ],
            "sol": "- Glycine: H~2~N-CH~2~-COOH (M = 75 g/mol).\n- Số mol glycine: n = 11,25 / 75 = 0,15 mol.\n- Phản ứng: H~2~N-CH~2~-COOH + NaOH ➔ H~2~N-CH~2~-COONa + H~2~O.\n- Muối thu được là sodium glycinate H~2~N-CH~2~-COONa (M = 97 g/mol).\n- Khối lượng muối khan: m = 0,15 × 97 = *14,55 gam*.\n➔ *Đáp án A.*"
        },
        {
            "q": "Câu 11. Trong môi trường base, protein có phản ứng màu biuret với",
            "opts": [
                "A. HNO~3~.",
                "B. NaCl.",
                "*C. Cu(OH)~2~.*" if is_teacher else "C. Cu(OH)~2~.",
                "D. Mg(OH)~2~."
            ],
            "sol": "Các phân tử protein chứa nhiều liên kết peptide (-CO-NH-) nên trong môi trường kiềm (base) sẽ phản ứng với Cu(OH)~2~ tạo phức chất màu tím đặc trưng. Đây là phản ứng màu biuret dùng để nhận biết peptide (từ tripeptide trở lên) và protein.\n➔ *Đáp án C.*"
        },
        {
            "q": "Câu 12. Câu nào sau đây không đúng?",
            "opts": [
                "A. Khi thuỷ phân hoàn toàn protein thu được hỗn hợp các α-amino acid.",
                "B. Có thể sử dụng phản ứng màu biuret để nhận biết sự có mặt của protein.",
                "C. Protein có thể tạo hợp chất màu vàng khi tác dụng với nitric acid.",
                "*D. Tất cả các loại protein đều không tan trong nước.*" if is_teacher else "D. Tất cả các loại protein đều không tan trong nước."
            ],
            "sol": "- Phương án D sai vì chỉ có các protein dạng hình sợi (keratin trong tóc/móng, fibroin trong tơ tằm, collagen trong da/gân) là không tan trong nước; còn các protein dạng hình cầu (như albumin trong lòng trắng trứng, hemoglobin trong máu) tan tốt trong nước tạo dung dịch keo.\n- Các phương án A, B, C đều là những phát biểu chính xác về tính chất của protein.\n➔ *Đáp án D.*"
        }
    ]

    for item in questions_p1:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, item["q"])

        if item.get("has_image"):
            if os.path.exists(IMG_PATH):
                p_img = doc.add_paragraph()
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(IMG_PATH, width=Inches(4.2))
            if item.get("sub_prompt"):
                p_sub = doc.add_paragraph()
                p_sub.paragraph_format.space_before = Pt(1)
                p_sub.paragraph_format.space_after = Pt(2)
                render_rich_text(p_sub, item["sub_prompt"], base_size=10.5)

        render_options_no_table(doc, item["opts"], is_teacher=is_teacher)
        if is_teacher:
            add_solution_no_table(doc, item["sol"])

    # ==========================
    # PHẦN II: TRẮC NGHIỆM ĐÚNG/SAI
    # ==========================
    add_section_header(
        doc,
        "PHẦN II. (2 điểm) Câu trắc nghiệm đúng sai.",
        "Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1
    p_q1_p2 = doc.add_paragraph()
    p_q1_p2.paragraph_format.space_before = Pt(4)
    p_q1_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q1_p2, "Câu 1. Cho các phát biểu sau:")

    c1_stmts = [
        ("a) Glucose và fructose là hai monosaccharide, có công thức phân tử C~6~H~12~O~6~.", "Đúng"),
        ("b) Trong môi trường acid, glucose và fructose có thể chuyển hóa lẫn nhau.", "Sai"),
        ("c) Dùng dung dịch AgNO~3~ trong NH~3~ để phân biệt glucose và fructose.", "Sai"),
        ("d) Cho 27 gam glucose lên men thành ethyl alcohol, toàn bộ khí CO~2~ sinh ra trong quá trình này được hấp thụ hết vào dung dịch Ca(OH)~2~ dư, tạo ra 40 gam kết tủa. Biết hiệu suất quá trình lên men đạt 75%.", "Sai")
    ]

    for stmt, ans in c1_stmts:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        c1_sol = (
            "- a) *Đúng.* Glucose và fructose đều là monosaccharide đồng phân của nhau, có cùng CTPT C~6~H~12~O~6~.\n"
            "- b) *Sai.* Glucose và fructose chỉ chuyển hóa qua lại lẫn nhau trong *môi trường kiềm* (base), không chuyển hóa trong môi trường acid.\n"
            "- c) *Sai.* Trong môi trường kiềm của dung dịch AgNO~3~/NH~3~ (thuốc thử Tollens), fructose chuyển hóa một phần thành glucose nên *cả hai đều tham gia phản ứng tráng bạc*. Để phân biệt glucose và fructose người ta dùng nước bromine (glucose làm mất màu nước bromine, fructose thì không).\n"
            "- d) *Sai.* Ta có:\n"
            "  + n(glucose) = 27 / 180 = 0,15 mol.\n"
            "  + Phản ứng: C~6~H~12~O~6~ ➔ 2C~2~H~5~OH + 2CO~2~.\n"
            "  + n(CO~2~ lý thuyết) = 2 × 0,15 = 0,3 mol.\n"
            "  + Với hiệu suất H = 75%: n(CO~2~ thực tế) = 0,3 × 75% = 0,225 mol.\n"
            "  + Dẫn CO~2~ vào Ca(OH)~2~ dư: n(CaCO~3~) = n(CO~2~) = 0,225 mol.\n"
            "  + Khối lượng kết tủa tạo thành: m = 0,225 × 100 = *22,5 gam* (không phải 40 gam)."
        )
        add_solution_no_table(doc, c1_sol)

    # Câu 2
    p_q2_p2 = doc.add_paragraph()
    p_q2_p2.paragraph_format.space_before = Pt(4)
    p_q2_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(
        p_q2_p2,
        "Câu 2. Thủy phân hoàn toàn 1 mol peptide X, thu được 1 mol glycine (Gly), 2 mol alanine (Ala), 1 mol valine (Val). Thủy phân không hoàn toàn X thu được tripeptide Gly-Ala-Val và dipeptide Ala-Gly."
    )

    c2_stmts = [
        ("a) X là tetrapeptide, có 3 liên kết peptide.", "Đúng"),
        ("b) X hòa tan Cu(OH)~2~ trong môi trường kiềm tạo dung dịch màu xanh lam.", "Sai"),
        ("c) Peptide X luôn có amino acid đầu N là Ala, amino acid đầu C của X là Val.", "Đúng"),
        ("d) Thủy phân hoàn toàn 31,6 gam X bằng dung dịch NaOH dư, thu được 45,8 gam muối. Mặt khác thủy phân hoàn toàn 31,6 gam X bằng dung dịch HCl dư, thu được 51,6 gam muối.", "Đúng")
    ]

    for stmt, ans in c2_stmts:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        c2_sol = (
            "- a) *Đúng.* Thủy phân hoàn toàn 1 mol X thu được 1 Gly + 2 Ala + 1 Val = 4 gốc amino acid ➔ X là tetrapeptide, có 4 - 1 = 3 liên kết peptide.\n"
            "- b) *Sai.* Tetrapeptide X có phản ứng màu biuret với Cu(OH)~2~ trong môi trường kiềm tạo dung dịch phức chất có màu *tím* đặc trưng (không phải màu xanh lam).\n"
            "- c) *Đúng.* Thủy phân không hoàn toàn tạo ra dipeptide Ala-Gly và tripeptide Gly-Ala-Val. Ghép nối các đoạn peptide: Ala-Gly + Gly-Ala-Val ➔ Ala-Gly-Ala-Val. Vậy X có cấu tạo duy nhất là Ala-Gly-Ala-Val: amino acid đầu N là Ala, amino acid đầu C là Val.\n"
            "- d) *Đúng.* Ta có:\n"
            "  + M(X) = M(Ala) + M(Gly) + M(Ala) + M(Val) - 3 × 18 = 89 + 75 + 89 + 117 - 54 = 316 g/mol.\n"
            "  + n(X) = 31,6 / 316 = 0,1 mol.\n"
            "  + Thủy phân trong NaOH dư: X + 4NaOH ➔ Muối + H~2~O.\n"
            "    Bảo toàn khối lượng: m(muối) = m(X) + m(NaOH) - m(H~2~O) = 31,6 + 0,4 × 40 - 0,1 × 18 = *45,8 gam* (khớp đề).\n"
            "  + Thủy phân trong HCl dư: X + 3H~2~O + 4HCl ➔ Muối.\n"
            "    Bảo toàn khối lượng: m(muối) = m(X) + m(H~2~O) + m(HCl) = 31,6 + 0,3 × 18 + 0,4 × 36,5 = *51,6 gam* (khớp đề)."
        )
        add_solution_no_table(doc, c2_sol)

    # ==========================
    # PHẦN III: TRẢ LỜI NGẮN
    # ==========================
    add_section_header(
        doc,
        "PHẦN III. (2 điểm) Câu trắc nghiệm yêu cầu trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 4."
    )

    questions_p3 = [
        {
            "q": "Câu 1. Cho các chất: CH~3~[CH~2~]~14~COONa, CH~3~[CH~2~]~10~CH~2~OSO~3~Na, CH~3~[CH~2~]~16~COOK, CH~3~[CH~2~]~7~CH=CH[CH~2~]~7~COONa, CH~3~COONa, (C~15~H~31~COO)~3~C~3~H~5~. Có bao nhiêu chất có thể là thành phần chính của xà phòng?",
            "ans": "3",
            "sol": "Xà phòng là hỗn hợp muối sodium hoặc potassium của các acid béo (mạch cacbon dài từ 12 carbon trở lên):\n"
                   "- CH~3~[CH~2~]~14~COONa: sodium palmitate (C~15~H~31~COONa) ➔ là thành phần của xà phòng.\n"
                   "- CH~3~[CH~2~]~10~CH~2~OSO~3~Na: sodium lauryl sulfate ➔ là chất giặt rửa tổng hợp.\n"
                   "- CH~3~[CH~2~]~16~COOK: potassium stearate (C~17~H~35~COOK) ➔ là thành phần của xà phòng.\n"
                   "- CH~3~[CH~2~]~7~CH=CH[CH~2~]~7~COONa: sodium oleate (C~17~H~33~COONa) ➔ là thành phần của xà phòng.\n"
                   "- CH~3~COONa: sodium acetate (muối acid mạch ngắn C2) ➔ không phải xà phòng.\n"
                   "- (C~15~H~31~COO)~3~C~3~H~5~: tripalmitin ➔ là chất béo.\n"
                   "➔ Có *3* chất có thể là thành phần chính của xà phòng."
        },
        {
            "q": "Câu 2. Thủy phân hoàn toàn 4,34 gam tripeptide mạch hở X (được tạo nên từ hai α-amino acid có công thức dạng H~2~NC~x~H~y~COOH) bằng dung dịch NaOH dư, thu được 6,38 gam muối. Mặt khác thủy phân hoàn toàn 4,34 gam X bằng dung dịch HCl dư, thu được bao nhiêu gam muối? (làm tròn đến hàng phần trăm)",
            "ans": "7,25",
            "sol": "1. Tìm số mol tripeptide X:\n"
                   "- Các amino acid tạo nên X có 1 nhóm -NH~2~ và 1 nhóm -COOH.\n"
                   "- Thủy phân trong NaOH: X + 3NaOH ➔ Muối + H~2~O.\n"
                   "- Gọi n(X) = a mol ➔ n(NaOH) = 3a mol; n(H~2~O) = a mol.\n"
                   "- Bảo toàn khối lượng: m(X) + m(NaOH) = m(muối) + m(H~2~O)\n"
                   "  ➔ 4,34 + 3a × 40 = 6,38 + 18a ➔ 102a = 2,04 ➔ a = 0,02 mol.\n"
                   "2. Thủy phân trong dung dịch HCl dư:\n"
                   "- Phương trình: X + 2H~2~O + 3HCl ➔ Muối.\n"
                   "- n(H~2~O) = 2a = 0,04 mol ➔ m(H~2~O) = 0,04 × 18 = 0,72 gam.\n"
                   "- n(HCl) = 3a = 0,06 mol ➔ m(HCl) = 0,06 × 36,5 = 2,19 gam.\n"
                   "- Bảo toàn khối lượng:\n"
                   "  m(muối) = m(X) + m(H~2~O) + m(HCl) = 4,34 + 0,72 + 2,19 = *7,25 gam*.\n"
                   "➔ Kết quả: *7,25*."
        },
        {
            "q": "Câu 3. Củ sắn khô chứa 38% khối lượng là tinh bột, còn lại là các chất không có khả năng lên men thành ancol etylic. Lên men 1 tấn sắn khô với hiệu suất cả quá trình là 81%. Toàn bộ lượng ancol etylic sinh ra để điều chế xăng E5 (có chứa 5% thể tích ancol etylic). Biết khối lượng riêng của ancol etylic là 0,8 g/mL, thể tích xăng E5 thu được là bao nhiêu lít? ( làm tròn đến hàng đơn vị)",
            "ans": "4370",
            "sol": "1. Tính khối lượng ethanol sinh ra:\n"
                   "- Khối lượng tinh bột trong 1 tấn sắn khô: m(tinh bột) = 1000 kg × 38% = 380 kg.\n"
                   "- Sơ đồ: (-C~6~H~10~O~5-)~n ➔ n C~6~H~12~O~6~ ➔ 2n C~2~H~5~OH + 2n CO~2~.\n"
                   "- Khối lượng ethanol lý thuyết: m_lt = 380 × (2 × 46 / 162) = 215,8025 kg.\n"
                   "- Với hiệu suất cả quá trình H = 81%: m_tt = 215,8025 × 81% = 174,8 kg.\n"
                   "2. Tính thể tích ethanol và thể tích xăng E5:\n"
                   "- D = 0,8 g/mL = 0,8 kg/L.\n"
                   "- Thể tích ethanol nguyên chất: V(ethanol) = 174,8 / 0,8 = 218,5 lít.\n"
                   "- Thể tích xăng E5 (chứa 5% thể tích ethanol):\n"
                   "  V(xăng E5) = 218,5 / 5% = 218,5 / 0,05 = *4370 lít*.\n"
                   "➔ Kết quả: *4370*."
        },
        {
            "q": "Câu 4. Cho các nhận định sau:\n(1) Protein dạng hình cầu và dạng hình sợi tan tốt trong nước.\n(2) Một trong những tính chất hoá học đặc trưng của protein là phản ứng thuỷ phân.\n(3) Phản ứng của protein với nitric acid cho sản phẩm có màu tím.\n(4) Khi đun nóng lòng trắng trứng sẽ xảy ra hiện tượng đông tụ.\n(5) Trong cơ thể, enzyme đóng vai trò là chất xúc tác sinh học.\nLiệt kê các nhận định đúng theo dãy số thứ tự tăng dần ( Ví dụ: 1234, 134, 24...)",
            "ans": "245",
            "sol": "Phân tích từng nhận định:\n"
                   "- (1) *Sai*: Protein hình sợi không tan trong nước, chỉ protein hình cầu tan trong nước.\n"
                   "- (2) *Đúng*: Protein bị thủy phân bởi acid, kiềm hoặc men enzyme.\n"
                   "- (3) *Sai*: Phản ứng với HNO~3~ đặc tạo kết tủa màu *vàng* (phản ứng xanthoproteic), không phải màu tím.\n"
                   "- (4) *Đúng*: Dưới tác dụng của nhiệt, protein bị đông tụ lại.\n"
                   "- (5) *Đúng*: Hầu hết enzyme là protein và đóng vai trò là chất xúc tác sinh học tốc độ cực cao trong các phản ứng sinh hóa.\n"
                   "➔ Các nhận định đúng là (2), (4), (5). Dãy số tăng dần: *245*."
        }
    ]

    for item in questions_p3:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, item["q"])

        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(2)
        if is_teacher:
            render_rich_text(p_ans, f"*Trả lời:* *{item['ans']}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
            add_solution_no_table(doc, item["sol"])
        else:
            render_rich_text(p_ans, "*Trả lời:* .................................................................................................", base_size=10.5)

    # ==========================
    # PHẦN IV / III: TỰ LUẬN
    # ==========================
    add_section_header(
        doc,
        "PHẦN IV. Tự luận (3,0 điểm)",
        "Thí sinh trình bày chi tiết lời giải các câu hỏi sau:"
    )

    # Câu 1
    p_tl1 = doc.add_paragraph()
    p_tl1.paragraph_format.space_before = Pt(4)
    p_tl1.paragraph_format.space_after = Pt(2)
    render_question_prompt(
        p_tl1,
        "Câu 1 (1,0 điểm). Viết tất cả các công thức cấu tạo của amine có công thức phân tử C~3~H~9~N? Gọi tên thay thế?"
    )

    if is_teacher:
        sol_tl1 = (
            "Amine có công thức phân tử C~3~H~9~N có 4 đồng phân cấu tạo, gồm:\n"
            "1. Amine bậc 1 (2 đồng phân):\n"
            "   - CH~3~-CH~2~-CH~2~-NH~2~ : *propan-1-amine*\n"
            "   - CH~3~-CH(NH~2~)-CH~3~ : *propan-2-amine*\n"
            "2. Amine bậc 2 (1 đồng phân):\n"
            "   - CH~3~-CH~2~-NH-CH~3~ : *N-methylethanamine*\n"
            "3. Amine bậc 3 (1 đồng phân):\n"
            "   - (CH~3~)~3~N (hoặc CH~3~-N(CH~3~)-CH~3~) : *N,N-dimethylmethanamine*"
        )
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=5)

    # Câu 2
    p_tl2 = doc.add_paragraph()
    p_tl2.paragraph_format.space_before = Pt(4)
    p_tl2.paragraph_format.space_after = Pt(2)
    render_question_prompt(
        p_tl2,
        "Câu 2 (1,0 điểm). Nêu hiện tượng, viết phương trình phản ứng khi cho:\n"
        "a) Saccharose phản ứng với Cu(OH)~2~ ở nhiệt độ thường.\n"
        "b) Dẫn methylamine vào dung dịch muối FeCl~3~."
    )

    if is_teacher:
        sol_tl2 = (
            "a) Cho saccharose tác dụng với Cu(OH)~2~ ở nhiệt độ thường:\n"
            "   - Hiện tượng: Kết tủa màu xanh lam nhạt Cu(OH)~2~ tan dần, thu được dung dịch phức chất đồng-saccharose có màu xanh lam đậm.\n"
            "   - Phương trình hóa học:\n"
            "     2C~12~H~22~O~11~ + Cu(OH)~2~ ➔ (C~12~H~21~O~11~)~2~Cu + 2H~2~O\n"
            "b) Dẫn methylamine vào dung dịch FeCl~3~:\n"
            "   - Hiện tượng: Xuất hiện kết tủa màu nâu đỏ của Fe(OH)~3~ và dung dịch mất dần màu vàng nâu của muối sắt(III).\n"
            "   - Phương trình hóa học:\n"
            "     3CH~3~NH~2~ + 3H~2~O + FeCl~3~ ➔ Fe(OH)~3~↓ (nâu đỏ) + 3CH~3~NH~3~Cl\n"
            "     (Giải thích: Methylamine có tính base, nhận H^+ của nước tạo ion OH^-, OH^- kết hợp với Fe^3+ tạo kết tủa Fe(OH)~3~)."
        )
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=5)

    # Câu 3
    p_tl3 = doc.add_paragraph()
    p_tl3.paragraph_format.space_before = Pt(4)
    p_tl3.paragraph_format.space_after = Pt(2)
    render_question_prompt(
        p_tl3,
        "Câu 3 (1,0 điểm).\n"
        "a) Cho dãy các chất: phenyl acetate, allyl acetate, methyl acetate, ethyl formate, tripalmitin. Có bao nhiêu chất trong dãy khi thủy phân trong dung dịch NaOH (dư), đun nóng sinh ra alcohol ? Viết phương trình phản ứng?\n"
        "b) Aspirin là một hợp chất được sử dụng làm giảm đau, hạ sốt được điều chế theo phản ứng sau:\n"
        "   (CH~3~CO)~2~O + HOC~6~H~4~COOH ➔ CH~3~COOC~6~H~4~COOH + CH~3~COOH\n"
        "   acetic anhydride   salicylic acid     aspirin\n"
        "Để sản xuất 500 nghìn viên thuốc aspirin cần tối thiểu m kg salicylic acid. Biết rằng mỗi viên thuốc có chứa 81 mg aspirin và hiệu suất phản ứng đạt 75%. Giá trị của m là bao nhiêu?"
    )

    if is_teacher:
        sol_tl3 = (
            "a) Số chất thủy phân trong NaOH sinh ra alcohol là *4* chất (allyl acetate, methyl acetate, ethyl formate, tripalmitin).\n"
            "   (Phenyl acetate CH~3~COOC~6~H~5~ là ester của phenol, khi thủy phân trong kiềm sinh ra hai muối CH~3~COONa, C~6~H~5~ONa và H~2~O, không tạo alcohol).\n"
            "   - Phương trình hóa học của 4 phản ứng:\n"
            "     (1) CH~3~COOCH~2~CH=CH~2~ + NaOH ➔ CH~3~COONa + CH~2~=CH-CH~2~OH (allyl alcohol)\n"
            "     (2) CH~3~COOCH~3~ + NaOH ➔ CH~3~COONa + CH~3~OH (methyl alcohol)\n"
            "     (3) HCOOC~2~H~5~ + NaOH ➔ HCOONa + C~2~H~5~OH (ethyl alcohol)\n"
            "     (4) (C~15~H~31~COO)~3~C~3~H~5~ + 3NaOH ➔ 3C~15~H~31~COONa + C~3~H~5~(OH)~3~ (glycerol)\n"
            "b) Tính giá trị của m:\n"
            "   - Khối lượng aspirin cần sản xuất:\n"
            "     m(aspirin) = 500 000 × 81 mg = 40 500 000 mg = 40,5 kg = 40 500 gam.\n"
            "   - Khối lượng mol phân tử:\n"
            "     + Salicylic acid (HOC~6~H~4~COOH, C~7~H~6~O~3~): M = 138 g/mol.\n"
            "     + Aspirin (CH~3~COOC~6~H~4~COOH, C~9~H~8~O~4~): M = 180 g/mol.\n"
            "   - Số mol aspirin cần sản xuất: n(aspirin) = 40 500 / 180 = 225 mol = 0,225 kmol.\n"
            "   - Theo phương trình hóa học, tỉ lệ phản ứng là 1 : 1.\n"
            "   - Do hiệu suất phản ứng là 75%, số mol salicylic acid tối thiểu cần dùng là:\n"
            "     n(salicylic acid) = 0,225 / 75% = 0,3 kmol.\n"
            "   - Khối lượng salicylic acid cần dùng là:\n"
            "     m = 0,3 kmol × 138 kg/kmol = *41,4 kg*.\n"
            "   ➔ Giá trị của m là *41,4*."
        )
        add_solution_no_table(doc, sol_tl3)
    else:
        add_student_writing_lines(doc, num_lines=6)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(r_end, font_size=10.5, bold=True)

    filename = (
        "De_Kiem_Tra_GK1_Hoa_12_Ma112_GiaoVien.docx"
        if is_teacher
        else "De_Kiem_Tra_GK1_Hoa_12_Ma112_HocSinh.docx"
    )
    filepath = os.path.join(OUTPUT_DIR, filename)
    doc.save(filepath)
    print(f"Đã tạo file Word: {filepath}")
    return filepath


def build_latex(is_teacher=False):
    """Tạo mã nguồn LaTeX chuẩn và biên dịch sang PDF."""
    tex = []
    tex.append(r"""\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{geometry}
\geometry{top=1.5cm,bottom=1.2cm,left=1.4cm,right=1.4cm,headheight=28pt,headsep=12pt,footskip=18pt}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{tabularx}
\usepackage{array}
\usepackage{colortbl}
\usepackage{fancyhdr}
\usepackage{enumitem}

\definecolor{hfRed}{HTML}{D32F2F}
\definecolor{hfGreen}{HTML}{8BC390}
\definecolor{mainGreen}{HTML}{2E7D32}
\definecolor{solBlue}{HTML}{1565C0}
\definecolor{tableGreen}{HTML}{E8F5E9}
\definecolor{ansRed}{HTML}{C62828}

\pagestyle{fancy}
\fancyhf{}
\renewcommand{\headrulewidth}{0.75pt}
\renewcommand{\headrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \headrulewidth\hfill}}
\renewcommand{\footrulewidth}{0.75pt}
\renewcommand{\footrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \footrulewidth\hfill}}

\fancyhead[L]{\small\itshape\color{hfRed}Tài liệu lưu hành nội bộ}
\fancyhead[R]{\small\color{hfRed}Thầy \textbf{TRẦN VĂN THẠNH} - 0777.470.803}
\fancyfoot[L]{\footnotesize\itshape\color{hfRed}CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế \quad | \quad CS2: 24 Đặng Thái Thân, TP Huế}
\fancyfoot[R]{\small\bfseries\color{hfRed}Trang \thepage}

\setlength{\parindent}{0pt}
\setlength{\parskip}{3pt}

\newcommand{\loigiai}[1]{
    \par\vspace{2pt}
    \begingroup\leftskip=12pt
    \textcolor{solBlue}{\textbf{$\blacktriangleright$ Lời giải chi tiết:}}\par
    \small #1
    \par\endgroup\vspace{3pt}
}

\begin{document}
""")

    # Header block
    tex.append(r"""
\begin{table}[h!]
\centering
\begin{tabularx}{\textwidth}{@{}p{8.5cm}X@{}}
\textbf{SỞ GIÁO DỤC VÀ ĐÀO TẠO} & \textbf{KIỂM TRA GIỮA HỌC KỲ I} \\
\textbf{ĐỀ CHÍNH THỨC} & \textbf{MÔN: HÓA HỌC - LỚP 12} \\
\textit{(Đề thi có """ + ("05" if is_teacher else "03") + r""" trang)} & Thời gian làm bài: 45 phút \\
\textbf{ } & \textbf{MÃ ĐỀ: 112} \\
\end{tabularx}
\end{table}
""")

    if is_teacher:
        tex.append(r"""\begin{center}
\textbf{\large\color{ansRed}(BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}
\end{center}
""")

    tex.append(r"""
\noindent\textit{(Nguyên tử khối trung bình của: H=1, C=12, O=16, N=14, S=32, Na=23, Ag=108, Br=80)}
\par\vspace{-2pt}\noindent\textcolor{gray!40}{\rule{\textwidth}{0.5pt}}\par\vspace{2pt}
""")

    if is_teacher:
        # Bảng đáp án rút gọn
        tex.append(r"""
\begin{center}
\textbf{\color{ansRed}BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 112)}
\end{center}
\textbf{\color{mainGreen}PHẦN I: Trắc nghiệm nhiều lựa chọn}
\begin{center}
\small
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} \\
\hline
\textbf{\color{ansRed}A} & \textbf{\color{ansRed}D} & \textbf{\color{ansRed}A} & \textbf{\color{ansRed}B} & \textbf{\color{ansRed}B} & \textbf{\color{ansRed}D} & \textbf{\color{ansRed}A} & \textbf{\color{ansRed}A} & \textbf{\color{ansRed}A} & \textbf{\color{ansRed}A} & \textbf{\color{ansRed}C} & \textbf{\color{ansRed}D} \\
\hline
\end{tabular}
\end{center}

\textbf{\color{mainGreen}PHẦN II \& PHẦN III: Đúng/Sai \& Trả lời ngắn}
\begin{center}
\small
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu 1 (Đ/S)} & \textbf{Câu 2 (Đ/S)} & \textbf{Câu 1 (TLN)} & \textbf{Câu 2 (TLN)} & \textbf{Câu 3 (TLN)} & \textbf{Câu 4 (TLN)} \\
\hline
\textbf{\color{ansRed}a-Đ, b-S, c-S, d-S} & \textbf{\color{ansRed}a-Đ, b-S, c-Đ, d-Đ} & \textbf{\color{ansRed}3} & \textbf{\color{ansRed}7,25} & \textbf{\color{ansRed}4370} & \textbf{\color{ansRed}245} \\
\hline
\end{tabular}
\end{center}
\par\vspace{-2pt}\noindent\textcolor{hfGreen}{\rule{\textwidth}{0.75pt}}\par\vspace{4pt}
""")

    # PHẦN I
    tex.append(r"""
\noindent\textbf{\large\color{mainGreen}PHẦN I. (3 điểm) Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{4pt}
""")

    questions_p1_tex = [
        (
            "Chất nào sau đây là ester :",
            ["\\ce{CH3COOCH3}.", "\\ce{HCOOH}.", "\\ce{CH3CHO}.", "\\ce{CH3OH}."],
            0,
            "Ester là hợp chất hữu cơ sinh ra khi thay thế nhóm -OH ở nhóm carboxyl (-COOH) của acid bằng nhóm -OR'.\\\\- \\ce{CH3COOCH3}: methyl acetate (ester).\\\\- \\ce{HCOOH}: formic acid (carboxylic acid).\\\\- \\ce{CH3CHO}: acetaldehyde (aldehyde).\\\\- \\ce{CH3OH}: methanol (alcohol).\\\\$\\to$ \\textbf{Chọn A.}"
        ),
        (
            "Chất nào sau đây là thành phần chính của chất giặt rửa tổng hợp?",
            ["\\ce{C3H5(OH)3}.", "\\ce{CH3[CH2]14COONa}.", "\\ce{CH3[CH2]16COOK}.", "\\ce{CH3[CH2]11C6H4SO3Na}."],
            3,
            "Chất giặt rửa tổng hợp là muối sodium alkylbenzene sulfonate (như \\ce{CH3[CH2]11C6H4SO3Na}) hoặc muối sodium alkylsulfate.\\\\- \\ce{CH3[CH2]14COONa} và \\ce{CH3[CH2]16COOK} là xà phòng.\\\\$\\to$ \\textbf{Chọn D.}"
        ),
        (
            "Sodium lauryl sulfate là một trong những chất giặt rửa tổng hợp phổ biến có công thức cấu tạo như sau:\\\\\n"
            r"\begin{center}\includegraphics[width=9cm]{images_de112/sodium_lauryl_sulfate.png}\end{center}" + "\n"
            "Đặc điểm của nhóm được khoanh tròn trong công thức trên là gì?",
            ["Ưa nước.", "Tạo kết tủa với nước cứng.", "Không phân cực.", "Kị nước."],
            0,
            "Phân tử chất giặt rửa tổng hợp gồm đầu phân cực (nhóm \\ce{-OSO3^- Na+} được khoanh tròn) ưa nước và đuôi alkyl dài kị nước.\\\\$\\to$ \\textbf{Chọn A.}"
        ),
        (
            "Cho 6 carbohydrate sau: glucose, fructose, maltose, saccharose, tinh bột và cellulose. Có bao nhiêu carbohydrate đã cho thuộc nhóm monosaccharide ?",
            ["1", "2", "3", "4"],
            1,
            "Monosaccharide gồm 2 chất: glucose và fructose. Maltose và saccharose là disaccharide. Tinh bột và cellulose là polysaccharide.\\\\$\\to$ \\textbf{Chọn B.}"
        ),
        (
            "Saccharose được cấu tạo từ",
            ["hai đơn vị $\\alpha$-glucose.", "một đơn vị $\\alpha$-glucose và một đơn vị $\\beta$-fructose.", "hai đơn vị $\\beta$-fructose.", "nhiều đơn vị $\\alpha$-glucose."],
            1,
            "Phân tử saccharose được cấu tạo từ một đơn vị $\\alpha$-glucose và một đơn vị $\\beta$-fructose liên kết $\\alpha$-1,2-glycoside.\\\\$\\to$ \\textbf{Chọn B.}"
        ),
        (
            "Phát biểu nào sau đây đúng khi nói về tinh bột và cellulose?",
            ["Đều tham gia phản ứng tráng bạc.", "Đều có phản ứng màu với dung dịch iodine.", "Đều tan tốt trong nước lạnh.", "Thủy phân hoàn toàn trong môi trường acid đều thu được glucose."],
            3,
            "Cả tinh bột và cellulose khi thủy phân hoàn toàn trong acid đều thu được glucose.\\\\$\\to$ \\textbf{Chọn D.}"
        ),
        (
            "Cellulose thuộc loại polysaccharide, là thành phần chính tạo nên màng tế bào thực vật, có nhiều trong gỗ, bông nõn. Công thức của cellulose là",
            ["\\ce{(C6H10O5)_n}.", "\\ce{C12H22O11}.", "\\ce{C6H12O6}.", "\\ce{C2H4O2}."],
            0,
            "Công thức phân tử của cellulose là \\ce{(C6H10O5)_n}.\\\\$\\to$ \\textbf{Chọn A.}"
        ),
        (
            "Các amino acid tồn tại chủ yếu ở trạng thái ion lưỡng cực, do đó chúng",
            ["có nhiệt độ nóng chảy cao và tan tốt trong nước.", "có nhiệt độ nóng chảy cao và ít tan trong nước.", "dễ nóng chảy và tan tốt trong nước.", "dễ nóng chảy và ít tan trong nước."],
            0,
            "Ở điều kiện thường, amino acid tồn tại dạng ion lưỡng cực nên có lực hút tĩnh điện liên phân tử lớn, dẫn tới nhiệt độ nóng chảy cao và dễ tan trong nước.\\\\$\\to$ \\textbf{Chọn A.}"
        ),
        (
            "Glutamic acid có vai trò quan trọng trong quá trình xây dựng cấu trúc tế bào của con người. Glutamic acid có điểm đẳng điện pI = 3,2 (pI là giá trị pH tại đó tổng điện tích trong amino acid bằng không). Khi đặt dung dịch glutamic acid ở pH = 6 trong điện trường, glutamic acid chủ yếu tồn tại ở dạng",
            ["anion và di chuyển về cực dương.", "ion lưỡng cực và đứng yên.", "anion và di chuyển về cực âm.", "cation và di chuyển về cực âm."],
            0,
            "Ở pH = 6 > pI = 3,2, môi trường kiềm hơn điểm đẳng điện nên glutamic acid nhường proton tạo dạng anion tích điện âm, trong điện trường sẽ di chuyển về cực dương (anode).\\\\$\\to$ \\textbf{Chọn A.}"
        ),
        (
            "Cho 11,25 gam glycine tác dụng với dung dịch NaOH vừa đủ, cô cạn dung dịch sau phản ứng, thu được m gam muối khan. Giá trị của m là",
            ["14,55", "13,32.", "11,10.", "16,65."],
            0,
            "$n_{\\text{glycine}} = 11,25 / 75 = 0,15\\text{ mol} \\Rightarrow m_{\\text{muối}} = 0,15 \\times 97 = \\textbf{14,55 gam}$.\\\\$\\to$ \\textbf{Chọn A.}"
        ),
        (
            "Trong môi trường base, protein có phản ứng màu biuret với",
            ["\\ce{HNO3}.", "\\ce{NaCl}.", "\\ce{Cu(OH)2}.", "\\ce{Mg(OH)2}."],
            2,
            "Protein có nhiều liên kết peptide phản ứng với \\ce{Cu(OH)2} trong môi trường kiềm tạo phức chất màu tím đặc trưng (phản ứng biuret).\\\\$\\to$ \\textbf{Chọn C.}"
        ),
        (
            "Câu nào sau đây không đúng?",
            ["Khi thuỷ phân hoàn toàn protein thu được hỗn hợp các $\\alpha$-amino acid.", "Có thể sử dụng phản ứng màu biuret để nhận biết sự có mặt của protein.", "Protein có thể tạo hợp chất màu vàng khi tác dụng với nitric acid.", "Tất cả các loại protein đều không tan trong nước."],
            3,
            "D sai vì protein hình cầu (albumin trong lòng trắng trứng, hemoglobin trong máu) tan tốt trong nước tạo dung dịch keo.\\\\$\\to$ \\textbf{Chọn D.}"
        )
    ]

    for idx, (prompt, opts, c_idx, sol) in enumerate(questions_p1_tex, 1):
        tex.append(f"\\textbf{{Câu {idx}.}} {prompt}\\\\[2pt]")
        formatted_opts = []
        labels = ["A", "B", "C", "D"]
        for o_i, opt in enumerate(opts):
            if is_teacher and o_i == c_idx:
                formatted_opts.append(f"\\textbf{{\\color{{ansRed}}{labels[o_i]}.}} \\textbf{{\\color{{ansRed}}{opt}}}")
            else:
                formatted_opts.append(f"\\textbf{{{labels[o_i]}.}} {opt}")

        max_l = max(len(o) for o in opts)
        if max_l <= 20:
            tex.append(r"\begin{tabularx}{\textwidth}{@{}XXXX@{}}" + " & ".join(formatted_opts) + r"\end{tabularx}\\[2pt]")
        elif max_l <= 45:
            tex.append(r"\begin{tabularx}{\textwidth}{@{}XX@{}}" + f"{formatted_opts[0]} & {formatted_opts[1]} \\\\ {formatted_opts[2]} & {formatted_opts[3]}" + r"\end{tabularx}\\[2pt]")
        else:
            for fo in formatted_opts:
                tex.append(f"\\quad {fo}\\\\[2pt]")

        if is_teacher:
            tex.append(f"\\loigiai{{{sol}}}")
        tex.append(r"\vspace{2pt}")

    # PHẦN II
    tex.append(r"""
\vspace{4pt}
\noindent\textbf{\large\color{mainGreen}PHẦN II. (2 điểm) Câu trắc nghiệm đúng sai.}\\
\textit{Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{4pt}
""")

    # Câu 1 phần II
    tex.append(r"\textbf{Câu 1.} Cho các phát biểu sau:\\[2pt]")
    c1_ii = [
        ("Glucose và fructose là hai monosaccharide, có công thức phân tử \\ce{C6H12O6}.", "Đúng"),
        ("Trong môi trường acid, glucose và fructose có thể chuyển hóa lẫn nhau.", "Sai"),
        ("Dùng dung dịch \\ce{AgNO3} trong \\ce{NH3} để phân biệt glucose và fructose.", "Sai"),
        ("Cho 27 gam glucose lên men thành ethyl alcohol, toàn bộ khí \\ce{CO2} sinh ra trong quá trình này được hấp thụ hết vào dung dịch \\ce{Ca(OH)2} dư, tạo ra 40 gam kết tủa. Biết hiệu suất quá trình lên men đạt 75\\%.", "Sai")
    ]
    sub_labels = ["a)", "b)", "c)", "d)"]
    for s_i, (stmt, ans) in enumerate(c1_ii):
        if is_teacher:
            tex.append(f"\\quad \\textbf{{{sub_labels[s_i]}}} {stmt} \\quad $\\to$ \\textbf{{\\color{{ansRed}}({ans})}}\\\\[2pt]")
        else:
            tex.append(f"\\quad \\textbf{{{sub_labels[s_i]}}} {stmt}\\\\[2pt]")

    if is_teacher:
        c1_ii_sol = r"""- a) \textbf{Đúng.} Glucose và fructose là hai monosaccharide đồng phân của nhau, có cùng CTPT \ce{C6H12O6}.\\
- b) \textbf{Sai.} Chúng chỉ chuyển hóa lẫn nhau trong môi trường \textbf{kiềm} (base).\\
- c) \textbf{Sai.} Thuốc thử Tollens có môi trường kiềm nên cả hai đều tráng bạc; để phân biệt dùng nước bromine.\\
- d) \textbf{Sai.} $n_{\text{glucose}} = 27 / 180 = 0,15\text{ mol} \Rightarrow n_{\ce{CO2}\text{ lt}} = 0,3\text{ mol}$. Với $H = 75\%$, $n_{\ce{CO2}\text{ tt}} = 0,225\text{ mol} \Rightarrow m_{\ce{CaCO3}} = 22,5\text{ gam} \neq 40\text{ gam}$."""
        tex.append(f"\\loigiai{{{c1_ii_sol}}}")

    # Câu 2 phần II
    tex.append(r"""\vspace{4pt}
\textbf{Câu 2.} Thủy phân hoàn toàn 1 mol peptide X, thu được 1 mol glycine (Gly), 2 mol alanine (Ala), 1 mol valine (Val). Thủy phân không hoàn toàn X thu được tripeptide Gly-Ala-Val và dipeptide Ala-Gly.\\[2pt]
""")
    c2_ii = [
        ("X là tetrapeptide, có 3 liên kết peptide.", "Đúng"),
        ("X hòa tan \\ce{Cu(OH)2} trong môi trường kiềm tạo dung dịch màu xanh lam.", "Sai"),
        ("Peptide X luôn có amino acid đầu N là Ala, amino acid đầu C của X là Val.", "Đúng"),
        ("Thủy phân hoàn toàn 31,6 gam X bằng dung dịch NaOH dư, thu được 45,8 gam muối. Mặt khác thủy phân hoàn toàn 31,6 gam X bằng dung dịch HCl dư, thu được 51,6 gam muối.", "Đúng")
    ]
    for s_i, (stmt, ans) in enumerate(c2_ii):
        if is_teacher:
            tex.append(f"\\quad \\textbf{{{sub_labels[s_i]}}} {stmt} \\quad $\\to$ \\textbf{{\\color{{ansRed}}({ans})}}\\\\[2pt]")
        else:
            tex.append(f"\\quad \\textbf{{{sub_labels[s_i]}}} {stmt}\\\\[2pt]")

    if is_teacher:
        c2_ii_sol = r"""- a) \textbf{Đúng.} X gồm 1 Gly + 2 Ala + 1 Val = 4 gốc amino acid $\to$ tetrapeptide, có 3 liên kết peptide.\\
- b) \textbf{Sai.} Tetrapeptide tạo dung dịch phức màu \textbf{tím} đặc trưng (phản ứng biuret).\\
- c) \textbf{Đúng.} Ghép đoạn: Ala-Gly + Gly-Ala-Val $\to$ Ala-Gly-Ala-Val. Đầu N là Ala, đầu C là Val.\\
- d) \textbf{Đúng.} $M_{\text{X}} = 89 + 75 + 89 + 117 - 54 = 316\text{ g/mol} \Rightarrow n_{\text{X}} = 0,1\text{ mol}$.\\
+ Thủy phân trong NaOH: $m_{\text{muối}} = 31,6 + 0,4 \times 40 - 0,1 \times 18 = \textbf{45,8 gam}$.\\
+ Thủy phân trong HCl: $m_{\text{muối}} = 31,6 + 0,3 \times 18 + 0,4 \times 36,5 = \textbf{51,6 gam}$."""
        tex.append(f"\\loigiai{{{c2_ii_sol}}}")

    # PHẦN III
    tex.append(r"""
\vspace{4pt}
\noindent\textbf{\large\color{mainGreen}PHẦN III. (2 điểm) Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4.}
\vspace{4pt}
""")

    p3_items = [
        (
            "Cho các chất: \\ce{CH3[CH2]14COONa}, \\ce{CH3[CH2]10CH2OSO3Na}, \\ce{CH3[CH2]16COOK}, \\ce{CH3[CH2]7CH=CH[CH2]7COONa}, \\ce{CH3COONa}, \\ce{(C15H31COO)3C3H5}. Có bao nhiêu chất có thể là thành phần chính của xà phòng?",
            "3",
            r"""Các chất là thành phần chính của xà phòng gồm muối sodium/potassium của acid béo:\\
- \ce{CH3[CH2]14COONa} (sodium palmitate);\\
- \ce{CH3[CH2]16COOK} (potassium stearate);\\
- \ce{CH3[CH2]7CH=CH[CH2]7COONa} (sodium oleate).\\
$\Rightarrow$ Có \textbf{3} chất."""
        ),
        (
            "Thủy phân hoàn toàn 4,34 gam tripeptide mạch hở X (được tạo nên từ hai $\\alpha$-amino acid có công thức dạng \\ce{H2NC_x H_y COOH}) bằng dung dịch NaOH dư, thu được 6,38 gam muối. Mặt khác thủy phân hoàn toàn 4,34 gam X bằng dung dịch HCl dư, thu được bao nhiêu gam muối? (làm tròn đến hàng phần trăm)",
            "7,25",
            r"""Đặt $n_{\text{X}} = a\text{ mol}$.\\
Phản ứng với NaOH: $\text{X} + 3\text{NaOH} \to \text{Muối} + \ce{H2O}$.\\
BTKL: $4,34 + 120a = 6,38 + 18a \Rightarrow 102a = 2,04 \Rightarrow a = 0,02\text{ mol}$.\\
Phản ứng với HCl: $\text{X} + 2\ce{H2O} + 3\text{HCl} \to \text{Muối}$.\\
$m_{\text{muối}} = 4,34 + 0,04 \times 18 + 0,06 \times 36,5 = \textbf{7,25 gam}$."""
        ),
        (
            "Củ sắn khô chứa 38\\% khối lượng là tinh bột, còn lại là các chất không có khả năng lên men thành ancol etylic. Lên men 1 tấn sắn khô với hiệu suất cả quá trình là 81\\%. Toàn bộ lượng ancol etylic sinh ra để điều chế xăng E5 (có chứa 5\\% thể tích ancol etylic). Biết khối lượng riêng của ancol etylic là 0,8 g/mL, thể tích xăng E5 thu được là bao nhiêu lít? ( làm tròn đến hàng đơn vị)",
            "4370",
            r"""$m_{\text{tinh bột}} = 1000 \times 38\% = 380\text{ kg}$.\\
$m_{\text{etanol}} = 380 \times \frac{92}{162} \times 81\% = 174,8\text{ kg}$.\\
$V_{\text{etanol}} = 174,8 / 0,8 = 218,5\text{ lít}$.\\
$V_{\text{xăng E5}} = 218,5 / 5\% = \textbf{4370 lít}$."""
        ),
        (
            "Cho các nhận định sau:\\\\\n"
            "(1) Protein dạng hình cầu và dạng hình sợi tan tốt trong nước.\\\\\n"
            "(2) Một trong những tính chất hoá học đặc trưng của protein là phản ứng thuỷ phân.\\\\\n"
            "(3) Phản ứng của protein với nitric acid cho sản phẩm có màu tím.\\\\\n"
            "(4) Khi đun nóng lòng trắng trứng sẽ xảy ra hiện tượng đông tụ.\\\\\n"
            "(5) Trong cơ thể, enzyme đóng vai trò là chất xúc tác sinh học.\\\\\n"
            "Liệt kê các nhận định đúng theo dãy số thứ tự tăng dần ( Ví dụ: 1234, 134, 24...)",
            "245",
            r"""Các nhận định đúng là (2), (4), (5).\\
$\Rightarrow$ Thứ tự tăng dần: \textbf{245}."""
        )
    ]

    for p_idx, (prompt, ans, sol) in enumerate(p3_items, 1):
        tex.append(f"\\textbf{{Câu {p_idx}.}} {prompt}\\\\[2pt]")
        if is_teacher:
            tex.append(f"\\textbf{{Trả lời:}} \\textbf{{\\color{{ansRed}}{ans}}}\\\\[2pt]")
            tex.append(f"\\loigiai{{{sol}}}")
        else:
            tex.append(r"\textbf{Trả lời:} \dotfill\\[4pt]")
        tex.append(r"\vspace{2pt}")

    # PHẦN IV
    tex.append(r"""
\vspace{4pt}
\noindent\textbf{\large\color{mainGreen}PHẦN IV. Tự luận (3,0 điểm)}\\
\textit{Thí sinh trình bày chi tiết lời giải các câu hỏi sau:}
\vspace{4pt}
""")

    p4_items = [
        (
            "Câu 1 (1,0 điểm). Viết tất cả các công thức cấu tạo của amine có công thức phân tử \\ce{C3H9N}? Gọi tên thay thế?",
            r"""Các đồng phân amine \ce{C3H9N} (4 đồng phân):\\
1. Amine bậc 1:\\
- \ce{CH3-CH2-CH2-NH2}: \textbf{propan-1-amine}\\
- \ce{CH3-CH(NH2)-CH3}: \textbf{propan-2-amine}\\
2. Amine bậc 2:\\
- \ce{CH3-CH2-NH-CH3}: \textbf{N-methylethanamine}\\
3. Amine bậc 3:\\
- \ce{(CH3)3N}: \textbf{N,N-dimethylmethanamine}"""
        ),
        (
            "Câu 2 (1,0 điểm). Nêu hiện tượng, viết phương trình phản ứng khi cho:\\\\\n"
            "a) Saccharose phản ứng với \\ce{Cu(OH)2} ở nhiệt độ thường.\\\\\n"
            "b) Dẫn methylamine vào dung dịch muối \\ce{FeCl3}.",
            r"""a) Cho saccharose tác dụng với \ce{Cu(OH)2}:\\
- Hiện tượng: Kết tủa xanh lam \ce{Cu(OH)2} tan, tạo dung dịch phức đồng màu xanh lam đậm.\\
- PTHH: \ce{2C12H22O11 + Cu(OH)2 -> (C12H21O11)2Cu + 2H2O}.\\
b) Dẫn methylamine vào dung dịch \ce{FeCl3}:\\
- Hiện tượng: Xuất hiện kết tủa nâu đỏ của \ce{Fe(OH)3}, dung dịch mất dần màu vàng nâu.\\
- PTHH: \ce{3CH3NH2 + 3H2O + FeCl3 -> Fe(OH)3 v} \text{(nâu đỏ)} + \ce{3CH3NH3Cl}."""
        ),
        (
            "Câu 3 (1,0 điểm).\\\\\n"
            "a) Cho dãy các chất: phenyl acetate, allyl acetate, methyl acetate, ethyl formate, tripalmitin. Có bao nhiêu chất trong dãy khi thủy phân trong dung dịch NaOH (dư), đun nóng sinh ra alcohol ? Viết phương trình phản ứng?\\\\\n"
            "b) Aspirin là một hợp chất được sử dụng làm giảm đau, hạ sốt được điều chế theo phản ứng sau:\\\\\n"
            r"\ce{(CH3CO)2O + HOC6H4COOH -> CH3COOC6H4COOH + CH3COOH}" + "\\\\\n"
            "Để sản xuất 500 nghìn viên thuốc aspirin cần tối thiểu m kg salicylic acid. Biết rằng mỗi viên thuốc có chứa 81 mg aspirin và hiệu suất phản ứng đạt 75\\%. Giá trị của m là bao nhiêu?",
            r"""a) Có \textbf{4} chất khi thủy phân trong NaOH sinh ra alcohol (allyl acetate, methyl acetate, ethyl formate, tripalmitin):\\
(1) \ce{CH3COOCH2CH=CH2 + NaOH ->[t^o] CH3COONa + CH2=CH-CH2OH}\\
(2) \ce{CH3COOCH3 + NaOH ->[t^o] CH3COONa + CH3OH}\\
(3) \ce{HCOOC2H5 + NaOH ->[t^o] HCOONa + C2H5OH}\\
(4) \ce{(C15H31COO)3C3H5 + 3NaOH ->[t^o] 3C15H31COONa + C3H5(OH)3}\\
b) Tính giá trị m:\\
- $m_{\text{aspirin}} = 500\,000 \times 81\text{ mg} = 40,5\text{ kg} \Rightarrow n_{\text{aspirin}} = 40,5 / 180 = 0,225\text{ kmol}$.\\
- Với hiệu suất $H = 75\%$, số mol salicylic acid cần tối thiểu là:\\
$n = 0,225 / 75\% = 0,3\text{ kmol}$.\\
- Khối lượng salicylic acid cần dùng:\\
$m = 0,3 \times 138 = \textbf{41,4 kg}$."""
        )
    ]

    for p_q, p_sol in p4_items:
        tex.append(f"{p_q}\\\\[2pt]")
        if is_teacher:
            tex.append(f"\\loigiai{{{p_sol}}}")
        else:
            tex.append(r"""\par\noindent\dotfill\\
\noindent\dotfill\\
\noindent\dotfill\\
\noindent\dotfill\\
\noindent\dotfill\\[4pt]""")
        tex.append(r"\vspace{2pt}")

    tex.append(r"""
\begin{center}
\textbf{---------- HẾT ----------}
\end{center}
\end{document}
""")

    tex_filename = (
        "De_Kiem_Tra_GK1_Hoa_12_Ma112_GiaoVien.tex"
        if is_teacher
        else "De_Kiem_Tra_GK1_Hoa_12_Ma112_HocSinh.tex"
    )
    tex_filepath = os.path.join(OUTPUT_DIR, tex_filename)
    with open(tex_filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(tex))
    print(f"Đã tạo file LaTeX: {tex_filepath}")

    # Biên dịch pdflatex
    cmd = ["pdflatex", "-interaction=nonstopmode", tex_filename]
    res = subprocess.run(cmd, cwd=OUTPUT_DIR, capture_output=True, text=True)
    if res.returncode == 0:
        print(f"Biên dịch PDF thành công: {tex_filename.replace('.tex', '.pdf')}")
    else:
        print(f"Biên dịch có cảnh báo hoặc lỗi. Exit code: {res.returncode}")

    # Dọn dẹp file phụ
    base_name = os.path.splitext(tex_filename)[0]
    for ext in [".aux", ".log", ".out", ".thm", ".synctex.gz"]:
        f_junk = os.path.join(OUTPUT_DIR, base_name + ext)
        if os.path.exists(f_junk):
            os.remove(f_junk)

    return tex_filepath


if __name__ == "__main__":
    print("=== BẮT ĐẦU TẠO ĐỀ KIỂM TRA GIỮA KÌ I HÓA 12 - MÃ ĐỀ 112 ===")
    docx_hs = build_docx(is_teacher=False)
    docx_gv = build_docx(is_teacher=True)
    tex_hs = build_latex(is_teacher=False)
    tex_gv = build_latex(is_teacher=True)
    print("=== HOÀN TẤT TẤT CẢ FILE THÀNH PHẨM ===")
