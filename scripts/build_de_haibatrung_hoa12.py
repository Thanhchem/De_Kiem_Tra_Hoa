# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Kì I Năm Học 2026 - 2027
Trường THPT Hai Bà Trưng - Huế - Môn: Hóa Học 12
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)

Xuất bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_12_HaiBaTrung_Hue_HocSinh.docx / .tex / .pdf
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_12_HaiBaTrung_Hue_GiaoVien.docx / .tex / .pdf
"""

import os
import sys
import subprocess
import glob
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

def add_quick_answer_table(doc):
    """Thêm BẢNG ĐÁP ÁN RÚT GỌN cho Giáo viên ở đầu tài liệu."""
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_title, "*BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN)*", base_size=11.5, base_bold=True, base_color=COLOR_CORRECT)

    # 1. PHẦN I
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(3)
    p_p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1, "*PHẦN I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - 0,25 đ/câu)*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

    tbl_p1 = doc.add_table(rows=2, cols=12)
    tbl_p1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_p1.autofit = False

    p1_answers = [
        ("1", "B"), ("2", "A"), ("3", "C"), ("4", "C"),
        ("5", "B"), ("6", "C"), ("7", "C"), ("8", "B"),
        ("9", "A"), ("10", "B"), ("11", "B"), ("12", "B")
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

    tbl_p23 = doc.add_table(rows=3, cols=6)
    tbl_p23.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_p23.autofit = False

    headers_p23 = [
        "Câu 1 (Đ/S)", "Câu 2 (Đ/S)",
        "Câu 1 (TLN)", "Câu 2 (TLN)", "Câu 3 (TLN)", "Câu 4 (TLN)"
    ]
    p23_answers = [
        "a: Đ | b: S\nc: Đ | d: S",
        "a: Đ | b: Đ\nc: Đ | d: Đ",
        "70", "7", "16,7", "5 (hoặc 4)"
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
        school_line="SỞ GD&ĐT TP HUẾ\nTRƯỜNG THPT HAI BÀ TRƯNG",
        exam_title="KIỂM TRA GIỮA KÌ I NĂM HỌC 2026 - 2027",
        subject_line="MÔN: HÓA HỌC 12",
        num_pages="03" if not is_teacher else "05",
        duration="45 phút",
        exam_code="CHÍNH THỨC",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; K = 39; Cl = 35,5; Ag = 108",
        is_teacher=is_teacher
    )

    if is_teacher:
        add_quick_answer_table(doc)

    # ==========================
    # PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN
    # ==========================
    add_section_header(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,0 điểm)",
        "Thí sinh trả lời từ Câu 1 đến Câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    questions_p1 = [
        {
            "q": "Câu 1. Chất nào sau đây thuộc loại ester?",
            "opts": ["A. CH~3~COOH.", "*B. CH~3~COOCH~3~.*" if is_teacher else "B. CH~3~COOCH~3~.", "C. C~2~H~5~OH.", "D. CH~3~CHO."],
            "sol": "Ester là hợp chất hữu cơ sinh ra khi thay thế nhóm -OH ở nhóm carboxyl (-COOH) của acid bằng nhóm -OR'.\n- CH~3~COOH: acetic acid (carboxylic acid).\n- CH~3~COOCH~3~: methyl acetate (ester).\n- C~2~H~5~OH: ethanol (alcohol).\n- CH~3~CHO: acetaldehyde (aldehyde).\n➔ *Đáp án B.*"
        },
        {
            "q": "Câu 2. Thành phần chính của xà phòng thường là",
            "opts": [
                "*A. muối sodium hoặc potassium của acid béo.*" if is_teacher else "A. muối sodium hoặc potassium của acid béo.",
                "B. muối sodium alkylsulfate.",
                "C. muối sodium alkylbenzene sulfonate.",
                "D. triester của glycerol với acid béo."
            ],
            "sol": "Xà phòng thường dùng là hỗn hợp muối sodium hoặc potassium của các acid béo (thường là stearic, palmitic, oleic acid...).\n- Phương án B và C là thành phần chính của chất giặt rửa tổng hợp.\n- Phương án D là chất béo (triglyceride).\n➔ *Đáp án A.*"
        },
        {
            "q": "Câu 3. Carbohydrate nào sau đây thuộc loại monosaccharide?",
            "opts": ["A. Saccharose.", "B. Tinh bột.", "*C. Glucose.*" if is_teacher else "C. Glucose.", "D. Cellulose."],
            "sol": "Monosaccharide là nhóm carbohydrate đơn giản nhất không bị thủy phân tiếp thành các carbohydrate nhỏ hơn. Hai monosaccharide quan trọng là glucose và fructose.\n- Saccharose là disaccharide.\n- Tinh bột và cellulose là polysaccharide.\n➔ *Đáp án C.*"
        },
        {
            "q": "Câu 4. Ở dạng mạch hở, trong phân tử glucose có chứa nhóm chức nào sau đây?",
            "opts": [
                "A. Chỉ có nhóm hydroxy (-OH).",
                "B. Một nhóm ketone (-CO-) và năm nhóm -OH.",
                "*C. Một nhóm aldehyde (-CHO) và năm nhóm -OH.*" if is_teacher else "C. Một nhóm aldehyde (-CHO) và năm nhóm -OH.",
                "D. Một nhóm carboxyl (-COOH) và bốn nhóm -OH."
            ],
            "sol": "Ở dạng mạch hở, phân tử glucose có cấu tạo gồm 1 nhóm aldehyde (-CHO) ở đầu mạch và 5 nhóm hydroxy (-OH) ở các vị trí liền kề: CH~2~OH-[CHOH]~4~-CHO.\n(Fructose mạch hở chứa 1 nhóm ketone và 5 nhóm -OH).\n➔ *Đáp án C.*"
        },
        {
            "q": "Câu 5. Cho glucose tác dụng với lượng dư thuốc thử Tollens ([Ag(NH~3~)~2~]OH), đun nóng nhẹ. Hiện tượng quan sát được là",
            "opts": [
                "A. xuất hiện kết tủa màu đỏ gạch.",
                "*B. có lớp kim loại bạc sáng bóng bám vào thành ống nghiệm.*" if is_teacher else "B. có lớp kim loại bạc sáng bóng bám vào thành ống nghiệm.",
                "C. dung dịch chuyển sang màu xanh lam đậm.",
                "D. có sủi bọt khí không màu thoát ra."
            ],
            "sol": "Do phân tử glucose có nhóm aldehyde (-CHO) nên tác dụng với thuốc thử Tollens ([Ag(NH~3~)~2~]OH) khi đun nóng giải phóng kim loại bạc sáng bóng bám vào thành ống nghiệm (phản ứng tráng bạc):\nCH~2~OH[CHOH]~4~CHO + 2[Ag(NH~3~)~2~]OH ➔ CH~2~OH[CHOH]~4~COONH~4~ + 2Ag↓ + 3NH~3~ + H~2~O.\n➔ *Đáp án B.*"
        },
        {
            "q": "Câu 6. Chất nào sau đây là thành phần chính của đường mía, đường phèn thường dùng trong pha chế đồ uống hàng ngày?",
            "opts": ["A. Glucose.", "B. Fructose.", "*C. Saccharose.*" if is_teacher else "C. Saccharose.", "D. Maltose."],
            "sol": "Saccharose có nhiều trong cây mía, củ cải đường, thốt nốt. Trong đời sống, saccharose là thành phần chủ yếu của đường mía, đường phèn, đường cát, đường phên dùng làm gia vị và giải khát hàng ngày.\n➔ *Đáp án C.*"
        },
        {
            "q": "Câu 7. Chất nào sau đây thuộc loại amine bậc một?",
            "opts": ["A. CH~3~-NH-CH~3~.", "B. (CH~3~)~3~N.", "*C. C~2~H~5~NH~2~.*" if is_teacher else "C. C~2~H~5~NH~2~.", "D. C~6~H~5~-NH-CH~3~."],
            "sol": "Bậc của amine bằng số gốc hydrocarbon thay thế nguyên tử H trong NH~3~:\n- C~2~H~5~NH~2~: Amine bậc một (có nhóm chức -NH~2~).\n- CH~3~-NH-CH~3~ và C~6~H~5~-NH-CH~3~: Amine bậc hai (có nhóm -NH-).\n- (CH~3~)~3~N: Amine bậc ba (có nhóm -N<).\n➔ *Đáp án C.*"
        },
        {
            "q": "Câu 8. Ở điều kiện thường, amine nào sau đây tồn tại ở thể khí?",
            "opts": ["A. Aniline.", "*B. Methylamine.*" if is_teacher else "B. Methylamine.", "C. Benzylamine.", "D. Hexan-1-amine."],
            "sol": "Ở điều kiện thường, bốn amine đồng đẳng đầu dãy là methylamine, dimethylamine, trimethylamine và ethylamine là chất khí có mùi khai khó chịu, tan nhiều trong nước. Các amine còn lại thường ở thể lỏng hoặc rắn.\n➔ *Đáp án B.*"
        },
        {
            "q": "Câu 9. Nhỏ vài giọt nước bromine vào ống nghiệm chứa aniline, hiện tượng quan sát được là",
            "opts": [
                "*A. xuất hiện kết tủa trắng.*" if is_teacher else "A. xuất hiện kết tủa trắng.",
                "B. dung dịch chuyển sang màu xanh lam.",
                "C. có bọt khí bay lên.",
                "D. xuất hiện kết tủa màu đỏ gạch."
            ],
            "sol": "Do nhóm -NH~2~ làm tăng hoạt tính của nhân thơm ở các vị trí ortho và para, aniline phản ứng thế dễ dàng với nước bromine tạo kết tủa trắng 2,4,6-tribromoaniline:\nC~6~H~5~NH~2~ + 3Br~2~ ➔ C~6~H~2~Br~3~NH~2~↓ (trắng) + 3HBr.\n➔ *Đáp án A.*"
        },
        {
            "q": "Câu 10. Công thức cấu tạo thu gọn của Alanine là",
            "opts": [
                "A. H~2~N-CH~2~-COOH.",
                "*B. CH~3~-CH(NH~2~)-COOH.*" if is_teacher else "B. CH~3~-CH(NH~2~)-COOH.",
                "C. H~2~N-(CH~2~)~2~-COOH.",
                "D. HOOC-CH~2~-CH(NH~2~)-COOH."
            ],
            "sol": "- Alanine là 2-aminopropanoic acid: CH~3~-CH(NH~2~)-COOH.\n- H~2~N-CH~2~-COOH là Glycine.\n- H~2~N-(CH~2~)~2~-COOH là β-alanine.\n- HOOC-CH~2~-CH(NH~2~)-COOH là Aspartic acid.\n➔ *Đáp án B.*"
        },
        {
            "q": "Câu 11. Số liên kết peptide có trong phân tử tripeptide Gly-Ala-Val là",
            "opts": ["A. 3.", "*B. 2.*" if is_teacher else "B. 2.", "C. 1.", "D. 4."],
            "sol": "Tripeptide được tạo nên từ 3 gốc α-amino acid liên kết bởi 2 liên kết peptide (-CO-NH-).\nTổng quát: Phân tử peptide có n gốc α-amino acid thì có (n - 1) liên kết peptide. Vậy phân tử tripeptide có 3 - 1 = 2 liên kết peptide.\n➔ *Đáp án B.*"
        },
        {
            "q": "Câu 12. Chất nào sau đây thuộc loại protein hình cầu, tan được trong nước tạo dung dịch keo?",
            "opts": [
                "A. Keratin (trong tóc, móng).",
                "*B. Albumin (trong lòng trắng trứng).*" if is_teacher else "B. Albumin (trong lòng trắng trứng).",
                "C. Fibroin (trong tơ tằm).",
                "D. Collagen (trong da, gân)."
            ],
            "sol": "- Albumin (trong lòng trắng trứng) và hemoglobin (trong máu) là protein dạng hình cầu, tan được trong nước tạo dung dịch keo.\n- Keratin, fibroin, collagen là các protein dạng hình sợi, không tan trong nước.\n➔ *Đáp án B.*"
        }
    ]

    for item in questions_p1:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, item["q"])
        render_options_no_table(doc, item["opts"], is_teacher=is_teacher)
        if is_teacher:
            add_solution_no_table(doc, item["sol"])

    # ==========================
    # PHẦN II: TRẮC NGHIỆM ĐÚNG/SAI
    # ==========================
    add_section_header(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai (2,0 điểm)",
        "Thí sinh trả lời từ Câu 1 đến Câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    # Câu 1
    p_q1_p2 = doc.add_paragraph()
    p_q1_p2.paragraph_format.space_before = Pt(4)
    p_q1_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q1_p2, "Câu 1. Thực hiện thí nghiệm xà phòng hóa chất béo trong phòng thí nghiệm:")
    p_exp1 = doc.add_paragraph()
    p_exp1.paragraph_format.left_indent = Cm(0.5)
    p_exp1.paragraph_format.space_before = Pt(1)
    p_exp1.paragraph_format.space_after = Pt(2)
    render_rich_text(p_exp1, "_Bước 1:_ Cho vào bát sứ nhỏ khoảng 1 gam mỡ lợn (chứa tripalmitin) và 2,5 mL dung dịch NaOH 40%.\n_Bước 2:_ Đun sôi nhẹ hỗn hợp, liên tục khuấy đều bằng đũa thủy tinh trong khoảng 8 - 10 phút và thỉnh thoảng thêm vài giọt nước cất.\n_Bước 3:_ Rót thêm vào hỗn hợp 5 mL dung dịch NaCl bão hòa nóng, khuấy nhẹ, sau đó để nguội.")

    c1_stmts = [
        ("a) Tripalmitin là triester của glycerol với palmitic acid, có công thức là (C~15~H~31~COO)~3~C~3~H~5~.", "Đúng"),
        ("b) Ở điều kiện thường, tripalmitin là chất lỏng, nhẹ hơn nước, không tan trong nước nhưng tan nhiều trong các dung môi hữu cơ không phân cực.", "Sai"),
        ("c) Mục đích chính của việc thêm dung dịch NaCl bão hòa ở bước 3 là làm cho xà phòng (muối sodium của acid béo) tách ra khỏi dung dịch, nổi lên trên mặt dung dịch.", "Đúng"),
        ("d) Xà phòng hóa hoàn toàn 8,06 gam tripalmitin bằng dung dịch NaOH vừa đủ, cô cạn thu được tối đa 8,44 gam muối sodium palmitate khan.", "Sai")
    ]

    for stmt, ans in c1_stmts:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        c1_sol = (
            "- a) *Đúng.* Palmitic acid là C~15~H~31~COOH, do đó tripalmitin là (C~15~H~31~COO)~3~C~3~H~5~.\n"
            "- b) *Sai.* Tripalmitin chứa các gốc hydrocarbon no C~15~H~31~- nên ở điều kiện thường là chất rắn (chất béo rắn gồm mỡ động vật như mỡ lợn, mỡ bò...). Chỉ có chất béo không no mới ở trạng thái lỏng (dầu thực vật).\n"
            "- c) *Đúng.* Dung dịch NaCl bão hòa làm tăng khối lượng riêng của dung dịch và làm giảm độ tan của xà phòng (muối sodium của acid béo), giúp xà phòng kết tinh nổi lên bề mặt để dễ dàng thu gom.\n"
            "- d) *Sai.* Ta có:\n"
            "  + M của tripalmitin (C~15~H~31~COO)~3~C~3~H~5~: M = 806 g/mol.\n"
            "  + n(tripalmitin) = 8,06 / 806 = 0,01 mol.\n"
            "  + Phản ứng: (C~15~H~31~COO)~3~C~3~H~5~ + 3NaOH ➔ 3C~15~H~31~COONa + C~3~H~5~(OH)~3~.\n"
            "  + n(muối) = 3 × 0,01 = 0,03 mol; M(C~15~H~31~COONa) = 278 g/mol.\n"
            "  + Khối lượng muối thu được tối đa: m = 0,03 × 278 = *8,34 gam* (không phải 8,44 gam)."
        )
        add_solution_no_table(doc, c1_sol)

    # Câu 2
    p_q2_p2 = doc.add_paragraph()
    p_q2_p2.paragraph_format.space_before = Pt(4)
    p_q2_p2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_q2_p2, "Câu 2. Cho các thông tin về cấu tạo và tính chất của tinh bột và cellulose:")

    c2_stmts = [
        ("a) Tinh bột là polymer thiên nhiên gồm amylose (mạch không phân nhánh, xoắn lò xo) và amylopectin (mạch phân nhánh).", "Đúng"),
        ("b) Hồ tinh bột tác dụng với dung dịch iodine ở nhiệt độ thường cho hợp chất màu xanh tím; khi đun nóng màu xanh tím biến mất, để nguội màu xanh tím xuất hiện trở lại.", "Đúng"),
        ("c) Cellulose có 3 nhóm -OH tự do, khi phản ứng với nitric acid đặc có mặt sulfuric acid đặc xúc tác, đun nóng tạo thành cellulose trinitrate dùng làm thuốc súng không khói.", "Đúng"),
        ("d) Cellulose là nguồn nguyên liệu chính trong công nghiệp sản xuất giấy, tơ sợi nhân tạo (tơ visco, tơ acetate) và vật liệu xây dựng.", "Đúng")
    ]

    for stmt, ans in c2_stmts:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.left_indent = Cm(0.5)
        p_st.paragraph_format.space_before = Pt(1)
        p_st.paragraph_format.space_after = Pt(1)
        render_statement(p_st, stmt, ans=ans, is_teacher=is_teacher)

    if is_teacher:
        c2_sol = (
            "- a) *Đúng.* Tinh bột là hỗn hợp hai thành phần: amylose (20 - 30%, mạch không phân nhánh xoắn lò xo gồm các gốc α-glucose) và amylopectin (70 - 80%, mạch phân nhánh).\n"
            "- b) *Đúng.* Các phân tử iodine len lỏi vào ống xoắn amylose tạo hợp chất bao màu xanh tím. Đun nóng làm duỗi mạch giải phóng iodine nên mất màu; làm nguội xoắn trở lại và tái hiện màu xanh tím.\n"
            "- c) *Đúng.* Trong mỗi mắt xích C~6~H~10~O~5~ của cellulose có 3 nhóm -OH tự do: [C~6~H~7~O~2~(OH)~3~]~n~. Khi tác dụng với HNO~3~ đặc (xúc tác H~2~SO~4~ đặc, đun nóng), cellulose tạo thành cellulose trinitrate [C~6~H~7~O~2~(ONO~2~)~3~]~n~ dùng làm thuốc súng không khói.\n"
            "- d) *Đúng.* Cellulose (bông gai, gỗ) được dùng làm vật liệu xây dựng, đồ gỗ, bột giấy làm giấy viết và chế tạo tơ sợi nhân tạo như tơ visco, tơ cellulose acetate."
        )
        add_solution_no_table(doc, c2_sol)

    # ==========================
    # PHẦN III: TRẢ LỜI NGẮN
    # ==========================
    add_section_header(
        doc,
        "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn (2,0 điểm)",
        "Thí sinh trả lời từ Câu 1 đến Câu 4. Thí sinh ghi đáp số hoặc kết quả vào phiếu trả lời."
    )

    questions_p3 = [
        {
            "q": "Câu 1. Để tráng một số lượng gương soi có diện tích bề mặt 0,35 m^2 với độ dày lớp bạc là 0,1 µm, người ta đun nóng dung dịch chứa 26,82 gam glucose với lượng dư thuốc thử Tollens. Biết khối lượng riêng của bạc là 10,49 g/cm^3 và hiệu suất phản ứng tráng bạc đạt 80% (tính theo glucose). Tính số lượng gương soi tối đa sản xuất được (làm tròn đến phần nguyên).",
            "ans": "70",
            "sol": "1. Tính số mol Ag thu được:\n"
                   "- n(glucose) = 26,82 / 180 = 0,149 mol.\n"
                   "- Phản ứng tráng bạc: C~6~H~12~O~6~ ➔ 2Ag.\n"
                   "- Với H = 80%, n(Ag) = 0,149 × 2 × 80% = 0,2384 mol.\n"
                   "- Khối lượng bạc tạo thành: m(Ag) = 0,2384 × 108 = 25,7472 gam.\n"
                   "2. Tính lượng bạc cần cho một chiếc gương:\n"
                   "- Diện tích bề mặt: S = 0,35 m^2 = 3500 cm^2.\n"
                   "- Độ dày lớp bạc: h = 0,1 µm = 10^-5 cm.\n"
                   "- Thể tích bạc trên 1 gương: V = S × h = 3500 × 10^-5 = 0,035 cm^3.\n"
                   "- Khối lượng bạc trên 1 gương: m_1 = V × D = 0,035 × 10,49 = 0,36715 gam.\n"
                   "3. Số lượng gương tối đa sản xuất được:\n"
                   "- N = m(Ag) / m_1 = 25,7472 / 0,36715 ≈ 70,127 chiếc.\n"
                   "➔ Làm tròn đến phần nguyên: *70*."
        },
        {
            "q": "Câu 2. Hợp chất hữu cơ X là amine bậc một, no, đơn chức, mạch hở. Cho 9,0 gam X tác dụng với dung dịch HCl dư, cô cạn cẩn thận dung dịch sau phản ứng thu được 16,3 gam muối khan. Xác định số nguyên tử hydrogen trong phân tử amine X.",
            "ans": "7",
            "sol": "- Đặt công thức phân tử amine no, đơn chức, mạch hở X là C~n~H~2n+3~N.\n"
                   "- Phản ứng: C~n~H~2n+3~N + HCl ➔ C~n~H~2n+4~NCl.\n"
                   "- Bảo toàn khối lượng: m(HCl) = m(muối) - m(X) = 16,3 - 9,0 = 7,3 gam.\n"
                   "- n(HCl) = 7,3 / 36,5 = 0,2 mol ➔ n(X) = n(HCl) = 0,2 mol.\n"
                   "- M(X) = 9,0 / 0,2 = 45 g/mol ➔ 14n + 17 = 45 ➔ n = 2.\n"
                   "- Công thức phân tử của X là C~2~H~7~N (công thức cấu tạo: C~2~H~5~NH~2~).\n"
                   "- Số nguyên tử hydrogen trong phân tử amine X là *7*."
        },
        {
            "q": "Câu 3. Cho m gam glycine tác dụng vừa đủ với 150 mL dung dịch HCl 1,0 M thu được dung dịch X. Cô cạn cẩn thận dung dịch X thu được lượng muối khan là bao nhiêu gam? (Kết quả làm tròn đến một chữ số thập phân).",
            "ans": "16,7",
            "sol": "- n(HCl) = 0,150 × 1,0 = 0,15 mol.\n"
                   "- Glycine là H~2~N-CH~2~-COOH (M = 75 g/mol).\n"
                   "- Phản ứng: H~2~N-CH~2~-COOH + HCl ➔ ClH~3~N-CH~2~-COOH.\n"
                   "- Muối thu được có M = 75 + 36,5 = 111,5 g/mol.\n"
                   "- Do phản ứng vừa đủ: n(muối) = n(HCl) = 0,15 mol.\n"
                   "- Khối lượng muối khan: m = 0,15 × 111,5 = 16,725 gam.\n"
                   "➔ Làm tròn đến một chữ số thập phân: *16,7*."
        },
        {
            "q": "Câu 4. Cho các dung dịch sau: glucose, saccharose, lòng trắng trứng (albumin), glycerol, ethylamine, Gly-Ala. Có bao nhiêu dung dịch phản ứng được với Cu(OH)2 ở điều kiện thường tạo thành dung dịch có màu xanh lam hoặc màu tím đặc trưng?",
            "ans": "5 (hoặc 4)",
            "sol": "- Glucose: Polyalcohol có nhiều nhóm -OH liền kề ➔ tạo dung dịch phức màu xanh lam đặc trưng.\n"
                   "- Saccharose: Polyalcohol có nhiều nhóm -OH kề nhau ➔ tạo dung dịch phức màu xanh lam.\n"
                   "- Lòng trắng trứng (albumin): Protein ➔ phản ứng màu biuret với Cu(OH)~2~ trong môi trường kiềm tạo phức chất màu tím đặc trưng.\n"
                   "- Glycerol: Triol kề nhau ➔ hòa tan Cu(OH)~2~ tạo dung dịch xanh lam đặc trưng.\n"
                   "- Ethylamine: Amine có cặp electron tự do trên N ➔ hòa tan Cu(OH)~2~ tạo phức [Cu(C~2~H~5~NH~2~)~4~](OH)~2~ dung dịch màu xanh lam đậm đặc trưng.\n"
                   "- Gly-Ala: Là dipeptide, chỉ có 1 liên kết peptide nên *không* có phản ứng màu biuret (phản ứng biuret chỉ có từ tripeptide trở lên).\n"
                   "➔ Kết luận: Có *5* chất thỏa mãn (nếu tính theo tính chất tạo phức của amine); hoặc *4* chất (nếu đề chỉ xét kiến thức polyalcohol và protein).\n"
                   "➔ Đáp số ghi: *5* (hoặc *4*)."
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
    # PHẦN IV: TỰ LUẬN
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
        "Câu 1 (1,0 điểm). Một cơ sở sản xuất thủ công sử dụng 120 kg muối sodium của acid béo (thành phần chính của xà phòng) để sản xuất các bánh xà phòng thơm. Mỗi bánh xà phòng thành phẩm có khối lượng tịnh 92 gam và chứa 80% khối lượng muối của acid béo (phần còn lại là các chất phụ gia, chất diệt khuẩn và hương liệu).\n"
        "a) Viết phương trình hóa học của phản ứng điều chế xà phòng từ tripalmitin và dung dịch NaOH.\n"
        "b) Tính số bánh xà phòng tối đa mà cơ sở trên sản xuất được (lấy kết quả nguyên)."
    )

    if is_teacher:
        sol_tl1 = (
            "a) Phương trình hóa học của phản ứng thủy phân tripalmitin trong NaOH (xà phòng hóa):\n"
            "   (C~15~H~31~COO)~3~C~3~H~5~ + 3NaOH ➔ 3C~15~H~31~COONa + C~3~H~5~(OH)~3~\n"
            "b) Tính số bánh xà phòng tối đa sản xuất được:\n"
            "   - Khối lượng muối acid béo có trong một bánh xà phòng thành phẩm là:\n"
            "     m_muối = 92 × 80% = 73,6 gam = 0,0736 kg.\n"
            "   - Tổng lượng muối dùng để sản xuất là 120 kg.\n"
            "   - Số bánh xà phòng tối đa sản xuất được là:\n"
            "     N = 120 / 0,0736 ≈ 1630,43 bánh.\n"
            "   ➔ Vì kết quả lấy số nguyên nên cơ sở trên sản xuất được tối đa *1630* bánh xà phòng."
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
        "Câu 2 (1,0 điểm). Một nhà máy chế biến đường mía ép 10 tấn mía tươi (chứa 13% saccharose về khối lượng) thu được nước mía. Sau quá trình tinh chế với hiệu suất toàn bộ đạt 85%, nhà máy thu được m kg đường saccharose tinh khiết.\n"
        "a) Nêu đặc điểm cấu tạo của phân tử saccharose (gốc monosaccharide tạo thành, liên kết glycoside, dạng mạch tồn tại).\n"
        "b) Xác định giá trị của m."
    )

    if is_teacher:
        sol_tl2 = (
            "a) Đặc điểm cấu tạo của phân tử saccharose:\n"
            "   - Gốc monosaccharide: Phân tử saccharose được tạo nên từ 1 gốc α-glucose và 1 gốc β-fructose.\n"
            "   - Liên kết glycoside: Hai gốc liên kết qua nguyên tử oxygen giữa C1 của α-glucose và C2 của β-fructose (liên kết α-1,2-glycoside).\n"
            "   - Dạng mạch: Cả hai nhóm -OH hemiacetal và hemiketal đều tham gia tạo liên kết glycoside nên phân tử saccharose không thể mở vòng, chỉ tồn tại ở dạng mạch vòng khép kín (không còn nhóm aldehyde tự do nên không có tính khử).\n"
            "b) Xác định giá trị m:\n"
            "   - Khối lượng saccharose có trong 10 tấn mía tươi theo lý thuyết:\n"
            "     m_lt = 10 tấn × 13% = 1,3 tấn = 1300 kg.\n"
            "   - Do hiệu suất quá trình tinh chế đạt 85%, khối lượng saccharose tinh chế thu được là:\n"
            "     m = 1300 × 85% = *1105 kg*."
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
        "Câu 3 (1,0 điểm). Thủy phân hoàn toàn 14,6 gam dipeptide Gly-Ala bằng một lượng vừa đủ dung dịch NaOH, đun nóng nhẹ.\n"
        "a) Viết phương trình hóa học của phản ứng thủy phân Gly-Ala trong dung dịch NaOH.\n"
        "b) Cô cạn dung dịch sau phản ứng thu được m gam chất rắn khan. Tính giá trị của m."
    )

    if is_teacher:
        sol_tl3 = (
            "a) Công thức dipeptide Gly-Ala: H~2~N-CH~2~-CONH-CH(CH~3~)-COOH.\n"
            "   Phương trình hóa học thủy phân trong dung dịch NaOH:\n"
            "   H~2~N-CH~2~-CONH-CH(CH~3~)-COOH + 2NaOH ➔ H~2~N-CH~2~-COONa + CH~3~-CH(NH~2~)-COONa + H~2~O\n"
            "   (Hoặc: Gly-Ala + 2NaOH ➔ Gly-Na + Ala-Na + H~2~O)\n"
            "b) Tính giá trị của m:\n"
            "   - M(Gly-Ala) = 75 + 89 - 18 = 146 g/mol.\n"
            "   - n(Gly-Ala) = 14,6 / 146 = 0,1 mol.\n"
            "   - Theo phương trình:\n"
            "     + n(NaOH) = 2 × n(Gly-Ala) = 0,2 mol ➔ m(NaOH) = 0,2 × 40 = 8,0 gam.\n"
            "     + n(H~2~O) = n(Gly-Ala) = 0,1 mol ➔ m(H~2~O) = 0,1 × 18 = 1,8 gam.\n"
            "   - Áp dụng định luật bảo toàn khối lượng:\n"
            "     m_rắn = m(Gly-Ala) + m(NaOH) - m(H~2~O) = 14,6 + 8,0 - 1,8 = *20,8 gam*.\n"
            "   (Cách 2: m = m(Gly-Na) + m(Ala-Na) = 0,1 × 97 + 0,1 × 111 = 9,7 + 11,1 = *20,8 gam*)."
        )
        add_solution_no_table(doc, sol_tl3)
    else:
        add_student_writing_lines(doc, num_lines=5)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(2)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(r_end, font_size=10.5, bold=True)

    filename = (
        "De_Kiem_Tra_GK1_Hoa_12_HaiBaTrung_Hue_GiaoVien.docx"
        if is_teacher
        else "De_Kiem_Tra_GK1_Hoa_12_HaiBaTrung_Hue_HocSinh.docx"
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

\newcommand{\dapan}[1]{\textcolor{ansRed}{\textbf{#1}}}
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
\textbf{SỞ GD\&ĐT TP HUẾ} & \textbf{KIỂM TRA GIỮA KÌ I NĂM HỌC 2026 - 2027} \\
\textbf{TRƯỜNG THPT HAI BÀ TRƯNG} & \textbf{MÔN: HÓA HỌC 12} \\
\textbf{ĐỀ CHÍNH THỨC} & Thời gian làm bài: 45 phút \\
\textit{(Đề thi có """ + ("05" if is_teacher else "03") + r""" trang)} & \textbf{MÃ ĐỀ: 101} \\
\end{tabularx}
\end{table}
""")

    if is_teacher:
        tex.append(r"""\begin{center}
\textbf{\large\color{ansRed}(BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}
\end{center}
""")

    tex.append(r"""
\noindent\textbf{Họ, tên thí sinh:} \dotfill \textbf{Lớp:} \dots\dots\dots\dots \textbf{SBD:} \dots\dots\dots\dots\\
\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; K = 39; Cl = 35,5; Ag = 108).}
\par\vspace{-2pt}\noindent\textcolor{gray!40}{\rule{\textwidth}{0.5pt}}\par\vspace{2pt}
""")

    if is_teacher:
        # Bảng đáp án rút gọn
        tex.append(r"""
\begin{center}
\textbf{\color{ansRed}BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN)}
\end{center}
\textbf{\color{mainGreen}PHẦN I: Trắc nghiệm nhiều lựa chọn}
\begin{center}
\small
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} \\
\hline
\textbf{\color{ansRed}B} & \textbf{\color{ansRed}A} & \textbf{\color{ansRed}C} & \textbf{\color{ansRed}C} & \textbf{\color{ansRed}B} & \textbf{\color{ansRed}C} & \textbf{\color{ansRed}C} & \textbf{\color{ansRed}B} & \textbf{\color{ansRed}A} & \textbf{\color{ansRed}B} & \textbf{\color{ansRed}B} & \textbf{\color{ansRed}B} \\
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
\textbf{\color{ansRed}a-Đ, b-S, c-Đ, d-S} & \textbf{\color{ansRed}a-Đ, b-Đ, c-Đ, d-Đ} & \textbf{\color{ansRed}70} & \textbf{\color{ansRed}7} & \textbf{\color{ansRed}16,7} & \textbf{\color{ansRed}5 (hoặc 4)} \\
\hline
\end{tabular}
\end{center}
\par\vspace{-2pt}\noindent\textcolor{hfGreen}{\rule{\textwidth}{0.75pt}}\par\vspace{4pt}
""")

    # PHẦN I
    tex.append(r"""
\noindent\textbf{\large\color{mainGreen}PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,0 điểm)}\\
\textit{Thí sinh trả lời từ Câu 1 đến Câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{4pt}
""")

    questions_p1_tex = [
        (
            "Chất nào sau đây thuộc loại ester?",
            ["\\ce{CH3COOH}.", "\\ce{CH3COOCH3}.", "\\ce{C2H5OH}.", "\\ce{CH3CHO}."],
            1,
            "Ester là hợp chất hữu cơ sinh ra khi thay thế nhóm -OH ở nhóm carboxyl (-COOH) bằng nhóm -OR'.\\\\- \\ce{CH3COOH}: acetic acid (carboxylic acid).\\\\- \\ce{CH3COOCH3}: methyl acetate (ester).\\\\- \\ce{C2H5OH}: ethanol (alcohol).\\\\- \\ce{CH3CHO}: acetaldehyde (aldehyde).\\\\$\\to$ \\textbf{Chọn B.}"
        ),
        (
            "Thành phần chính của xà phòng thường là",
            ["muối sodium hoặc potassium của acid béo.", "muối sodium alkylsulfate.", "muối sodium alkylbenzene sulfonate.", "triester của glycerol với acid béo."],
            0,
            "Xà phòng là hỗn hợp muối sodium hoặc potassium của các acid béo (như sodium stearate, sodium palmitate...).\\\\Các chất ở phương án B và C là chất giặt rửa tổng hợp. Phương án D là chất béo.\\\\$\\to$ \\textbf{Chọn A.}"
        ),
        (
            "Carbohydrate nào sau đây thuộc loại monosaccharide?",
            ["Saccharose.", "Tinh bột.", "Glucose.", "Cellulose."],
            2,
            "Monosaccharide là nhóm carbohydrate đơn giản nhất không bị thủy phân tiếp. Hai monosaccharide tiêu biểu là glucose và fructose.\\\\- Saccharose là disaccharide.\\\\- Tinh bột và cellulose là polysaccharide.\\\\$\\to$ \\textbf{Chọn C.}"
        ),
        (
            "Ở dạng mạch hở, trong phân tử glucose có chứa nhóm chức nào sau đây?",
            ["Chỉ có nhóm hydroxy (-OH).", "Một nhóm ketone (-CO-) và năm nhóm -OH.", "Một nhóm aldehyde (-CHO) và năm nhóm -OH.", "Một nhóm carboxyl (-COOH) và bốn nhóm -OH."],
            2,
            "Glucose mạch hở có công thức cấu tạo: \\ce{CH2OH-[CHOH]4-CHO}, gồm 1 nhóm aldehyde (-CHO) và 5 nhóm hydroxy (-OH) kề nhau.\\\\$\\to$ \\textbf{Chọn C.}"
        ),
        (
            "Cho glucose tác dụng với lượng dư thuốc thử Tollens ([\\ce{Ag(NH3)2}]OH), đun nóng nhẹ. Hiện tượng quan sát được là",
            ["xuất hiện kết tủa màu đỏ gạch.", "có lớp kim loại bạc sáng bóng bám vào thành ống nghiệm.", "dung dịch chuyển sang màu xanh lam đậm.", "có sủi bọt khí không màu thoát ra."],
            1,
            "Glucose có nhóm -CHO nên tác dụng với thuốc thử Tollens giải phóng Ag sáng bóng bám lên thành ống nghiệm (phản ứng tráng bạc):\\\\\\ce{CH2OH[CHOH]4CHO + 2[Ag(NH3)2]OH ->[t^o] CH2OH[CHOH]4COONH4 + 2Ag v + 3NH3 + H2O}.\\\\$\\to$ \\textbf{Chọn B.}"
        ),
        (
            "Chất nào sau đây là thành phần chính của đường mía, đường phèn thường dùng trong pha chế đồ uống hàng ngày?",
            ["Glucose.", "Fructose.", "Saccharose.", "Maltose."],
            2,
            "Saccharose là thành phần chính của đường mía, đường phèn, đường cát, đường thốt nốt dùng làm chất tạo ngọt trong gia đình.\\\\$\\to$ \\textbf{Chọn C.}"
        ),
        (
            "Chất nào sau đây thuộc loại amine bậc một?",
            ["\\ce{CH3-NH-CH3}.", "\\ce{(CH3)3N}.", "\\ce{C2H5NH2}.", "\\ce{C6H5-NH-CH3}."],
            2,
            "Bậc của amine bằng số gốc hydrocarbon thay thế nguyên tử H trong \\ce{NH3}:\\\\- \\ce{C2H5NH2} (ethylamine) là amine bậc một (chứa nhóm \\ce{-NH2}).\\\\- \\ce{CH3-NH-CH3} và \\ce{C6H5-NH-CH3} là amine bậc hai.\\\\- \\ce{(CH3)3N} là amine bậc ba.\\\\$\\to$ \\textbf{Chọn C.}"
        ),
        (
            "Ở điều kiện thường, amine nào sau đây tồn tại ở thể khí?",
            ["Aniline.", "Methylamine.", "Benzylamine.", "Hexan-1-amine."],
            1,
            "Bốn amine đầu dãy (methylamine, dimethylamine, trimethylamine, ethylamine) là các chất khí ở nhiệt độ thường, có mùi khai khó chịu, tan nhiều trong nước.\\\\$\\to$ \\textbf{Chọn B.}"
        ),
        (
            "Nhỏ vài giọt nước bromine vào ống nghiệm chứa aniline, hiện tượng quan sát được là",
            ["xuất hiện kết tủa trắng.", "dung dịch chuyển sang màu xanh lam.", "có bọt khí bay lên.", "xuất hiện kết tủa màu đỏ gạch."],
            0,
            "Aniline phản ứng thế dễ dàng với nước bromine tạo kết tủa trắng của 2,4,6-tribromoaniline:\\\\\\ce{C6H5NH2 + 3Br2 -> C6H2Br3NH2 v} \\text{(trắng)} + \\ce{3HBr}.\\\\$\\to$ \\textbf{Chọn A.}"
        ),
        (
            "Công thức cấu tạo thu gọn của Alanine là",
            ["\\ce{H2N-CH2-COOH}.", "\\ce{CH3-CH(NH2)-COOH}.", "\\ce{H2N-(CH2)2-COOH}.", "\\ce{HOOC-CH2-CH(NH2)-COOH}."],
            1,
            "Alanine là 2-aminopropanoic acid: \\ce{CH3-CH(NH2)-COOH}.\\\\- \\ce{H2N-CH2-COOH} là Glycine.\\\\- \\ce{H2N-(CH2)2-COOH} là $\\beta$-alanine.\\\\- \\ce{HOOC-CH2-CH(NH2)-COOH} là Aspartic acid.\\\\$\\to$ \\textbf{Chọn B.}"
        ),
        (
            "Số liên kết peptide có trong phân tử tripeptide Gly-Ala-Val là",
            ["3.", "2.", "1.", "4."],
            1,
            "Tripeptide gồm 3 gốc $\\alpha$-amino acid liên kết bởi 2 liên kết peptide (-CO-NH-).\\\\$\\to$ \\textbf{Chọn B.}"
        ),
        (
            "Chất nào sau đây thuộc loại protein hình cầu, tan được trong nước tạo dung dịch keo?",
            ["Keratin (trong tóc, móng).", "Albumin (trong lòng trắng trứng).", "Fibroin (trong tơ tằm).", "Collagen (trong da, gân)."],
            1,
            "Albumin (lòng trắng trứng) và hemoglobin là protein hình cầu tan được trong nước tạo dung dịch keo. Keratin, fibroin, collagen là protein hình sợi không tan trong nước.\\\\$\\to$ \\textbf{Chọn B.}"
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
\noindent\textbf{\large\color{mainGreen}PHẦN II. Câu trắc nghiệm đúng sai (2,0 điểm)}\\
\textit{Thí sinh trả lời từ Câu 1 đến Câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{4pt}
""")

    # Câu 1 phần II
    tex.append(r"""
\textbf{Câu 1.} Thực hiện thí nghiệm xà phòng hóa chất béo trong phòng thí nghiệm:\\
\textit{Bước 1:} Cho vào bát sứ nhỏ khoảng 1 gam mỡ lợn (chứa tripalmitin) và 2,5 mL dung dịch NaOH 40\%.\\
\textit{Bước 2:} Đun sôi nhẹ hỗn hợp, liên tục khuấy đều bằng đũa thủy tinh trong khoảng 8 - 10 phút và thỉnh thoảng thêm vài giọt nước cất.\\
\textit{Bước 3:} Rót thêm vào hỗn hợp 5 mL dung dịch NaCl bão hòa nóng, khuấy nhẹ, sau đó để nguội.\\[2pt]
""")
    c1_ii = [
        ("Tripalmitin là triester của glycerol với palmitic acid, có công thức là \\ce{(C15H31COO)3C3H5}.", "Đúng"),
        ("Ở điều kiện thường, tripalmitin là chất lỏng, nhẹ hơn nước, không tan trong nước nhưng tan nhiều trong các dung môi hữu cơ không phân cực.", "Sai"),
        ("Mục đích chính của việc thêm dung dịch NaCl bão hòa ở bước 3 là làm cho xà phòng (muối sodium của acid béo) tách ra khỏi dung dịch, nổi lên trên mặt dung dịch.", "Đúng"),
        ("Xà phòng hóa hoàn toàn 8,06 gam tripalmitin bằng dung dịch NaOH vừa đủ, cô cạn thu được tối đa 8,44 gam muối sodium palmitate khan.", "Sai")
    ]
    sub_labels = ["a)", "b)", "c)", "d)"]
    for s_i, (stmt, ans) in enumerate(c1_ii):
        if is_teacher:
            tex.append(f"\\quad \\textbf{{{sub_labels[s_i]}}} {stmt} \\quad $\\to$ \\textbf{{\\color{{ansRed}}({ans})}}\\\\[2pt]")
        else:
            tex.append(f"\\quad \\textbf{{{sub_labels[s_i]}}} {stmt}\\\\[2pt]")

    if is_teacher:
        c1_ii_sol = r"""- a) \textbf{Đúng.} Tripalmitin là triester của glycerol với palmitic acid \ce{C15H31COOH}, có công thức là \ce{(C15H31COO)3C3H5}.\\
- b) \textbf{Sai.} Gốc palmitate là gốc no nên tripalmitin ở điều kiện thường là chất rắn (chất béo rắn).\\
- c) \textbf{Đúng.} NaCl làm tăng khối lượng riêng dung dịch và giảm độ tan của muối acid béo, giúp xà phòng kết tinh nổi lên trên.\\
- d) \textbf{Sai.} $M = 806\text{ g/mol} \Rightarrow n = 8,06 / 806 = 0,01\text{ mol} \Rightarrow n_{\text{muối}} = 3 \times 0,01 = 0,03\text{ mol}$.\\
Khối lượng muối thu được tối đa là $0,03 \times 278 = \textbf{8,34 gam}$ (không phải 8,44 gam)."""
        tex.append(f"\\loigiai{{{c1_ii_sol}}}")

    # Câu 2 phần II
    tex.append(r"""\vspace{4pt}
\textbf{Câu 2.} Cho các thông tin về cấu tạo và tính chất của tinh bột và cellulose:\\[2pt]
""")
    c2_ii = [
        ("Tinh bột là polymer thiên nhiên gồm amylose (mạch không phân nhánh, xoắn lò xo) và amylopectin (mạch phân nhánh).", "Đúng"),
        ("Hồ tinh bột tác dụng với dung dịch iodine ở nhiệt độ thường cho hợp chất màu xanh tím; khi đun nóng màu xanh tím biến mất, để nguội màu xanh tím xuất hiện trở lại.", "Đúng"),
        ("Cellulose có 3 nhóm -OH tự do, khi phản ứng với nitric acid đặc có mặt sulfuric acid đặc xúc tác, đun nóng tạo thành cellulose trinitrate dùng làm thuốc súng không khói.", "Đúng"),
        ("Cellulose là nguồn nguyên liệu chính trong công nghiệp sản xuất giấy, tơ sợi nhân tạo (tơ visco, tơ acetate) và vật liệu xây dựng.", "Đúng")
    ]
    for s_i, (stmt, ans) in enumerate(c2_ii):
        if is_teacher:
            tex.append(f"\\quad \\textbf{{{sub_labels[s_i]}}} {stmt} \\quad $\\to$ \\textbf{{\\color{{ansRed}}({ans})}}\\\\[2pt]")
        else:
            tex.append(f"\\quad \\textbf{{{sub_labels[s_i]}}} {stmt}\\\\[2pt]")

    if is_teacher:
        c2_ii_sol = r"""- a) \textbf{Đúng.} Tinh bột gồm amylose (mạch không phân nhánh xoắn lò xo) và amylopectin (mạch phân nhánh).\\
- b) \textbf{Đúng.} Ở nhiệt độ thường iodine chui vào lòng xoắn amylose tạo hợp chất màu xanh tím; đun nóng duỗi xoắn giải phóng iodine làm mất màu; để nguội xoắn trở lại tạo màu xanh tím.\\
- c) \textbf{Đúng.} Trong mỗi mắt xích \ce{C6H10O5} của cellulose có 3 nhóm -OH tự do: \ce{[C6H7O2(OH)3]_n}. Đun nóng với \ce{HNO3} đặc/\ce{H2SO4} đặc tạo cellulose trinitrate dùng làm thuốc súng không khói.\\
- d) \textbf{Đúng.} Cellulose dùng làm vật liệu xây dựng, sản xuất giấy và chế tạo tơ visco, tơ acetate."""
        tex.append(f"\\loigiai{{{c2_ii_sol}}}")

    # PHẦN III
    tex.append(r"""
\vspace{4pt}
\noindent\textbf{\large\color{mainGreen}PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn (2,0 điểm)}\\
\textit{Thí sinh trả lời từ Câu 1 đến Câu 4. Thí sinh ghi đáp số hoặc kết quả vào phiếu trả lời.}
\vspace{4pt}
""")

    p3_items = [
        (
            "Để tráng một số lượng gương soi có diện tích bề mặt $0,35\\text{ m}^2$ với độ dày lớp bạc là $0,1\\ \\mu\\text{m}$, người ta đun nóng dung dịch chứa $26,82\\text{ gam}$ glucose với lượng dư thuốc thử Tollens. Biết khối lượng riêng của bạc là $10,49\\text{ g/cm}^3$ và hiệu suất phản ứng tráng bạc đạt $80\\%$ (tính theo glucose). Tính số lượng gương soi tối đa sản xuất được (làm tròn đến phần nguyên).",
            "70",
            r"""$n_{\text{glucose}} = 26,82 / 180 = 0,149\text{ mol}$.\\
Phản ứng: \ce{C6H12O6 -> 2Ag}.\\
$n_{\text{Ag}} = 0,149 \times 2 \times 80\% = 0,2384\text{ mol} \Rightarrow m_{\text{Ag}} = 0,2384 \times 108 = 25,7472\text{ gam}$.\\
Lượng Ag trên 1 chiếc gương:\\
$S = 0,35\text{ m}^2 = 3500\text{ cm}^2; \quad h = 0,1\ \mu\text{m} = 10^{-5}\text{ cm} \Rightarrow V = S \cdot h = 0,035\text{ cm}^3$.\\
$m_1 = V \cdot D = 0,035 \times 10,49 = 0,36715\text{ gam}$.\\
Số gương tối đa: $N = \frac{25,7472}{0,36715} \approx 70,127 \Rightarrow \textbf{70}$ chiếc gương."""
        ),
        (
            "Hợp chất hữu cơ X là amine bậc một, no, đơn chức, mạch hở. Cho $9,0\\text{ gam}$ X tác dụng với dung dịch HCl dư, cô cạn cẩn thận dung dịch sau phản ứng thu được $16,3\\text{ gam}$ muối khan. Xác định số nguyên tử hydrogen trong phân tử amine X.",
            "7",
            r"""Công thức phân tử amine: \ce{C_n H_{2n+3} N}.\\
Bảo toàn khối lượng: $m_{\text{HCl}} = 16,3 - 9,0 = 7,3\text{ gam} \Rightarrow n_{\text{HCl}} = 7,3 / 36,5 = 0,2\text{ mol}$.\\
$n_{\text{X}} = n_{\text{HCl}} = 0,2\text{ mol} \Rightarrow M_{\text{X}} = 9,0 / 0,2 = 45\text{ g/mol} \Rightarrow 14n + 17 = 45 \Rightarrow n = 2$.\\
Công thức phân tử là \ce{C2H7N} (amine bậc 1 là \ce{C2H5NH2}).\\
Số nguyên tử hydrogen trong phân tử là \textbf{7}."""
        ),
        (
            "Cho m gam glycine tác dụng vừa đủ với $150\\text{ mL}$ dung dịch HCl $1,0\\text{ M}$ thu được dung dịch X. Cô cạn cẩn thận dung dịch X thu được lượng muối khan là bao nhiêu gam? (Kết quả làm tròn đến một chữ số thập phân).",
            "16,7",
            r"""$n_{\text{HCl}} = 0,150 \times 1,0 = 0,15\text{ mol}$.\\
Glycine là \ce{H2N-CH2-COOH} ($M = 75\text{ g/mol}$).\\
Phản ứng: \ce{H2N-CH2-COOH + HCl -> ClH3N-CH2-COOH}.\\
$M_{\text{muối}} = 75 + 36,5 = 111,5\text{ g/mol}$.\\
$m_{\text{muối}} = 0,15 \times 111,5 = 16,725\text{ gam} \approx \textbf{16,7}\text{ gam}$."""
        ),
        (
            "Cho các dung dịch sau: glucose, saccharose, lòng trắng trứng (albumin), glycerol, ethylamine, Gly-Ala. Có bao nhiêu dung dịch phản ứng được với \\ce{Cu(OH)2} ở điều kiện thường tạo thành dung dịch có màu xanh lam hoặc màu tím đặc trưng?",
            "5 (hoặc 4)",
            r"""- Glucose, saccharose, glycerol: polyalcohol có các nhóm -OH liền kề $\to$ tạo dung dịch xanh lam.\\
- Lòng trắng trứng (albumin): protein $\to$ phản ứng biuret tạo màu tím đặc trưng.\\
- Ethylamine: amine hòa tan \ce{Cu(OH)2} tạo phức chất dung dịch màu xanh lam đậm đặc trưng.\\
- Gly-Ala: dipeptide nên không có phản ứng biuret.\\
$\Rightarrow$ Có \textbf{5} chất thỏa mãn (hoặc \textbf{4} chất nếu không tính tính tạo phức của amine)."""
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
            "Câu 1 (1,0 điểm). Một cơ sở sản xuất thủ công sử dụng 120 kg muối sodium của acid béo (thành phần chính của xà phòng) để sản xuất các bánh xà phòng thơm. Mỗi bánh xà phòng thành phẩm có khối lượng tịnh 92 gam và chứa 80\\% khối lượng muối của acid béo (phần còn lại là các chất phụ gia, chất diệt khuẩn và hương liệu).\\\\\n"
            "a) Viết phương trình hóa học của phản ứng điều chế xà phòng từ tripalmitin và dung dịch NaOH.\\\\\n"
            "b) Tính số bánh xà phòng tối đa mà cơ sở trên sản xuất được (lấy kết quả nguyên).",
            r"""a) Phương trình hóa học của phản ứng xà phòng hóa tripalmitin:\\
\ce{(C15H31COO)3C3H5 + 3NaOH ->[t^o] 3C15H31COONa + C3H5(OH)3}.\\
b) Khối lượng muối acid béo trong mỗi bánh xà phòng thành phẩm là:\\
$m_1 = 92 \times 80\% = 73,6\text{ gam} = 0,0736\text{ kg}$.\\
Số bánh xà phòng tối đa sản xuất được là:\\
$N = \frac{120}{0,0736} \approx 1630,43$ bánh $\Rightarrow$ Tối đa \textbf{1630} bánh."""
        ),
        (
            "Câu 2 (1,0 điểm). Một nhà máy chế biến đường mía ép 10 tấn mía tươi (chứa 13\\% saccharose về khối lượng) thu được nước mía. Sau quá trình tinh chế với hiệu suất toàn bộ đạt 85\\%, nhà máy thu được m kg đường saccharose tinh khiết.\\\\\n"
            "a) Nêu đặc điểm cấu tạo của phân tử saccharose (gốc monosaccharide tạo thành, liên kết glycoside, dạng mạch tồn tại).\\\\\n"
            "b) Xác định giá trị của m.",
            r"""a) Đặc điểm cấu tạo phân tử saccharose:\\
- Gồm một gốc $\alpha$-glucose và một gốc $\beta$-fructose.\\
- Liên kết với nhau qua nguyên tử oxygen giữa C1 của $\alpha$-glucose và C2 của $\beta$-fructose (liên kết $\alpha$-1,2-glycoside).\\
- Cả hai nhóm -OH hemiacetal và hemiketal đều tham gia liên kết glycoside nên phân tử saccharose chỉ tồn tại ở dạng mạch vòng khép kín, không thể mở vòng.\\
b) Khối lượng saccharose lý thuyết có trong 10 tấn mía tươi:\\
$m_{\text{lt}} = 10\text{ tấn} \times 13\% = 1,3\text{ tấn} = 1300\text{ kg}$.\\
Khối lượng saccharose tinh chế thực tế thu được:\\
$m = 1300 \times 85\% = \textbf{1105 kg}$."""
        ),
        (
            "Câu 3 (1,0 điểm). Thủy phân hoàn toàn 14,6 gam dipeptide Gly-Ala bằng một lượng vừa đủ dung dịch NaOH, đun nóng nhẹ.\\\\\n"
            "a) Viết phương trình hóa học của phản ứng thủy phân Gly-Ala trong dung dịch NaOH.\\\\\n"
            "b) Cô cạn dung dịch sau phản ứng thu được m gam chất rắn khan. Tính giá trị của m.",
            r"""a) Phương trình hóa học:\\
\ce{H2N-CH2-CONH-CH(CH3)-COOH + 2NaOH ->[t^o] H2N-CH2-COONa + CH3-CH(NH2)-COONa + H2O}\\
(hoặc viết gọn: \ce{Gly-Ala + 2NaOH ->[t^o] Gly-Na + Ala-Na + H2O}).\\
b) Tính giá trị m:\\
- $M_{\text{Gly-Ala}} = 75 + 89 - 18 = 146\text{ g/mol} \Rightarrow n_{\text{Gly-Ala}} = 14,6 / 146 = 0,1\text{ mol}$.\\
- Theo phương trình: $n_{\text{NaOH}} = 0,2\text{ mol} \Rightarrow m_{\text{NaOH}} = 8,0\text{ gam}$; $n_{\ce{H2O}} = 0,1\text{ mol} \Rightarrow m_{\ce{H2O}} = 1,8\text{ gam}$.\\
- Bảo toàn khối lượng:\\
$m = m_{\text{Gly-Ala}} + m_{\text{NaOH}} - m_{\ce{H2O}} = 14,6 + 8,0 - 1,8 = \textbf{20,8 gam}$.\\
(Cách khác: $m = 0,1 \times 97 + 0,1 \times 111 = 9,7 + 11,1 = \textbf{20,8 gam}$)."""
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
        "De_Kiem_Tra_GK1_Hoa_12_HaiBaTrung_Hue_GiaoVien.tex"
        if is_teacher
        else "De_Kiem_Tra_GK1_Hoa_12_HaiBaTrung_Hue_HocSinh.tex"
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
    print("=== BẮT ĐẦU TẠO ĐỀ KIỂM TRA GIỮA KÌ I HÓA 12 - THPT HAI BÀ TRƯNG ===")
    docx_hs = build_docx(is_teacher=False)
    docx_gv = build_docx(is_teacher=True)
    tex_hs = build_latex(is_teacher=False)
    tex_gv = build_latex(is_teacher=True)
    print("=== HOÀN TẤT TẤT CẢ FILE THÀNH PHẨM ===")
