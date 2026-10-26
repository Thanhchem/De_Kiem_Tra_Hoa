# -*- coding: utf-8 -*-
"""
Script biên soạn trọn bộ 3 tài liệu chuẩn theo FORM ĐỀ KIỂM TRA cho KHTN 6 (2026-2027):
1. De_Cuong_GK1_KHTN6_HocSinh.docx (Bản Học Sinh - Chuẩn A4, có dòng kẻ làm bài)
2. De_Cuong_GK1_KHTN6_GiaoVien.docx (Bản Giáo Viên - Có đáp án đỏ và lời giải chi tiết 100%)
3. De_Cuong_GK1_KHTN6_DayTrenManHinh.pdf (Slide PDF tỉ lệ 16:9 Full HD dạy trên Tivi 65 inch không viền đen)
Và phiên bản docx: De_Cuong_GK1_KHTN6_DayTrenManHinh.docx
Thầy TRẦN VĂN THẠNH - 0777.470.803
"""

import os
import sys
import subprocess
import shutil

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Import các hàm chuẩn từ form_de_kiem_tra
sys.path.insert(0, r"c:\Antigravity_Thanh\scripts")
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
    COLOR_HF_RED,
    COLOR_HF_GREEN_HEX
)

from khtn6_ontap_gk1_data import (
    INFO,
    VATLI_THEORY, VATLI_MCQ, VATLI_SHORT, VATLI_TF, VATLI_ESSAY,
    HOA_THEORY,
    SINH_CHAPTER1_THEORY, SINH_CHAPTER1_EXERCISES,
    SINH_CHAPTER5_THEORY, SINH_CHAPTER5_EXERCISES,
    SINH_CHAPTER6_THEORY, SINH_CHAPTER6_EXERCISES
)

OUTPUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

ONEDRIVE_DIRS = [
    r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn6\KHTN6-DEKIEMTRADINHKI\New folder",
    r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn6\KHTN6-DEKIEMTRADINHKI",
    r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn6"
]

COLOR_PART = RGBColor(198, 40, 40)   # Đỏ đậm cho Phân môn
COLOR_SUBSECTION = RGBColor(0, 77, 64) # Xanh lục trầm cho bài tập

def add_part_heading(doc, title):
    """Tiêu đề phân môn lớn."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(title)
    set_run_style(r, font_size=12.5, bold=True, color=COLOR_PART)
    # Đường kẻ dưới tiêu đề phân môn
    from form_de_kiem_tra import add_p_border_bottom
    add_p_border_bottom(p, color="C62828", sz="8")

# ==============================================================================
# 1. HÀM SINH FILE WORD (HỌC SINH / GIÁO VIÊN / MÀN HÌNH)
# ==============================================================================

def generate_word_document(mode="student"):
    """
    mode: 'student', 'teacher', 'screen'
    """
    is_teacher = (mode == "teacher")
    is_screen = (mode == "screen")

    if mode == "teacher":
        filename = "De_Cuong_GK1_KHTN6_GiaoVien.docx"
    elif mode == "screen":
        filename = "De_Cuong_GK1_KHTN6_DayTrenManHinh.docx"
    else:
        filename = "De_Cuong_GK1_KHTN6_HocSinh.docx"

    filepath = os.path.join(OUTPUT_DIR, filename)

    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="TRƯỜNG TRUNG HỌC CƠ SỞ",
        exam_title="ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA KÌ I - NĂM HỌC 2026 - 2027",
        subject_line="MÔN: KHOA HỌC TỰ NHIÊN 6",
        num_pages="04",
        duration="60 phút",
        exam_code="KHTN6-GK1",
        atomic_masses=None,
        is_teacher=is_teacher
    )

    base_size = 11.5 if is_screen else 10.5

    # --------------------------------------------------------------------------
    # PHẦN 1: PHÂN MÔN VẬT LÝ
    # --------------------------------------------------------------------------
    add_part_heading(doc, "PHẦN I. PHÂN MÔN VẬT LÝ")

    # I. Lý thuyết
    add_section_header(doc, "I. LÝ THUYẾT TRỌNG TÂM")
    for q in VATLI_THEORY:
        p_qt = doc.add_paragraph()
        p_qt.paragraph_format.space_before = Pt(4)
        p_qt.paragraph_format.space_after = Pt(2)
        render_rich_text(p_qt, f"*{q['title']}*", base_size=base_size, base_bold=True, base_color=COLOR_SOLUTION)
        for item in q["items"]:
            p_item = doc.add_paragraph()
            p_item.paragraph_format.left_indent = Cm(0.4)
            p_item.paragraph_format.space_before = Pt(1)
            p_item.paragraph_format.space_after = Pt(1)
            render_rich_text(p_item, item, base_size=base_size)

    # II. Trắc nghiệm nhiều lựa chọn
    add_section_header(doc, "II. TRẮC NGHIỆM NHIỀU LỰA CHỌN", "Chọn một phương án trả lời đúng nhất trong mỗi câu hỏi sau:")
    for q in VATLI_MCQ:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, f"Câu {q['id']}: {q['content']}", base_size=base_size)

        render_options_no_table(doc, q["options"], is_teacher=is_teacher)

        if is_teacher and q.get("solution"):
            add_solution_no_table(doc, f"Đáp án đúng: *{q['answer']}*.\n{q['solution']}")
        elif is_screen:
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.left_indent = Cm(0.6)
            p_sp.paragraph_format.space_before = Pt(2)
            p_sp.paragraph_format.space_after = Pt(2)
            render_rich_text(p_sp, "✎ Ghi chú / Giải thích: ........................................................................................................", base_size=base_size-1)

    # III. Trắc nghiệm trả lời ngắn
    add_section_header(doc, "III. TRẮC NGHIỆM TRẢ LỜI NGẮN", "Học sinh tính toán và viết câu trả lời ngắn gọn vào chỗ trống:")
    for q in VATLI_SHORT:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, f"Câu {q['id']}. {q['content']}", base_size=base_size)

        for sub in q["sub_questions"]:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.left_indent = Cm(0.4)
            p_sub.paragraph_format.space_before = Pt(2)
            p_sub.paragraph_format.space_after = Pt(1)
            render_rich_text(p_sub, f"*{sub['sub_id']})* {sub['text']}", base_size=base_size)

            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.8)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(2)

            if is_teacher:
                render_rich_text(p_ans, f"➔ *Đáp số:* {sub['answer']}", base_size=base_size, base_color=COLOR_CORRECT)
                if sub.get("solution"):
                    add_solution_no_table(doc, sub["solution"])
            elif is_screen:
                render_rich_text(p_ans, "➔ Trả lời: ...................................................................................................................", base_size=base_size)
            else:
                render_rich_text(p_ans, "Đáp số: ......................................................................................................................", base_size=base_size)

    # IV. Trắc nghiệm đúng sai
    add_section_header(doc, "IV. TRẮC NGHIỆM ĐÚNG SAI", "Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn Đúng hoặc Sai:")
    for q in VATLI_TF:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, f"Câu {q['id']}: {q['context']}", base_size=base_size)

        for item in q["items"]:
            p_it = doc.add_paragraph()
            p_it.paragraph_format.left_indent = Cm(0.5)
            p_it.paragraph_format.space_before = Pt(1)
            p_it.paragraph_format.space_after = Pt(1)
            ans_str = "Đúng" if item["answer"] else "Sai"

            if is_teacher:
                render_statement(p_it, f"{item['id']}) {item['content']}", ans=ans_str, is_teacher=True, base_size=base_size)
                # Kèm giải thích
                p_rea = doc.add_paragraph()
                p_rea.paragraph_format.left_indent = Cm(1.0)
                p_rea.paragraph_format.space_before = Pt(0)
                p_rea.paragraph_format.space_after = Pt(2)
                render_rich_text(p_rea, f"↳ _Giải thích:_ {item['reason']}", base_size=base_size-1, base_italic=True, base_color=COLOR_SOLUTION)
            else:
                render_rich_text(p_it, f"{item['id']}) {item['content']}   [   ] Đúng     [   ] Sai", base_size=base_size)

    # V. Tự luận
    add_section_header(doc, "V. PHẦN TỰ LUẬN", "Học sinh trình bày chi tiết lời giải vào bài làm:")
    for q in VATLI_ESSAY:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, f"Câu {q['id']}: {q['content']}", base_size=base_size)

        if is_teacher:
            add_solution_no_table(doc, q["solution"])
        elif is_screen:
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.left_indent = Cm(0.6)
            render_rich_text(p_sp, "✎ Bài giải chi tiết: ........................................................................................................", base_size=base_size)
        else:
            add_student_writing_lines(doc, num_lines=5)

    # --------------------------------------------------------------------------
    # PHẦN 2: PHÂN MÔN HOÁ HỌC
    # --------------------------------------------------------------------------
    add_part_heading(doc, "PHẦN II. PHÂN MÔN HOÁ HỌC")
    add_section_header(doc, "LÝ THUYẾT TRỌNG TÂM HÓA HỌC")
    for q in HOA_THEORY:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_rich_text(p_q, f"*{q['title']}*", base_size=base_size, base_bold=True, base_color=COLOR_SOLUTION)
        for item in q["items"]:
            p_it = doc.add_paragraph()
            p_it.paragraph_format.left_indent = Cm(0.4)
            p_it.paragraph_format.space_before = Pt(1)
            p_it.paragraph_format.space_after = Pt(1)
            render_rich_text(p_it, item, base_size=base_size)

    # --------------------------------------------------------------------------
    # PHẦN 3: PHÂN MÔN SINH HỌC
    # --------------------------------------------------------------------------
    add_part_heading(doc, "PHẦN III. PHÂN MÔN SINH HỌC")

    # Chương I
    add_section_header(doc, "CHƯƠNG I: MỞ ĐẦU - DỤNG CỤ QUAN SÁT")
    for item in SINH_CHAPTER1_THEORY:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.left_indent = Cm(0.4)
        p_t.paragraph_format.space_before = Pt(1)
        p_t.paragraph_format.space_after = Pt(2)
        render_rich_text(p_t, f"• {item}", base_size=base_size)

    p_bt_head = doc.add_paragraph()
    p_bt_head.paragraph_format.space_before = Pt(3)
    p_bt_head.paragraph_format.space_after = Pt(2)
    render_rich_text(p_bt_head, "*BÀI TẬP MINH HOẠ CHƯƠNG I:*", base_size=base_size, base_bold=True, base_color=COLOR_SUBSECTION)

    for ex in SINH_CHAPTER1_EXERCISES:
        p_ex = doc.add_paragraph()
        p_ex.paragraph_format.left_indent = Cm(0.4)
        p_ex.paragraph_format.space_before = Pt(2)
        p_ex.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ex, f"*{ex['id']}.* {ex['question']}", base_size=base_size)

        if is_teacher:
            add_solution_no_table(doc, ex["solution"])
        elif is_screen:
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.left_indent = Cm(0.8)
            render_rich_text(p_sp, "✎ Trả lời: ...................................................................................................................", base_size=base_size)
        else:
            add_student_writing_lines(doc, num_lines=3)

    # Chương V
    add_section_header(doc, "CHƯƠNG V: TẾ BÀO – ĐƠN VỊ CƠ SỞ CỦA SỰ SỐNG")
    for item in SINH_CHAPTER5_THEORY:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.left_indent = Cm(0.4)
        p_t.paragraph_format.space_before = Pt(1)
        p_t.paragraph_format.space_after = Pt(1)
        render_rich_text(p_t, item, base_size=base_size)

    p_bt5_head = doc.add_paragraph()
    p_bt5_head.paragraph_format.space_before = Pt(3)
    p_bt5_head.paragraph_format.space_after = Pt(2)
    render_rich_text(p_bt5_head, "*BÀI TẬP MINH HOẠ CHƯƠNG V:*", base_size=base_size, base_bold=True, base_color=COLOR_SUBSECTION)

    for ex in SINH_CHAPTER5_EXERCISES:
        p_ex = doc.add_paragraph()
        p_ex.paragraph_format.left_indent = Cm(0.4)
        p_ex.paragraph_format.space_before = Pt(2)
        p_ex.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ex, f"*{ex['id']}.* {ex['question']}", base_size=base_size)

        if is_teacher:
            add_solution_no_table(doc, ex["solution"])
        elif is_screen:
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.left_indent = Cm(0.8)
            render_rich_text(p_sp, "✎ Trả lời: ...................................................................................................................", base_size=base_size)
        else:
            add_student_writing_lines(doc, num_lines=4)

    # Chương VI
    add_section_header(doc, "CHƯƠNG VI: TỪ TẾ BÀO ĐẾN CƠ THỂ")
    for item in SINH_CHAPTER6_THEORY:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.left_indent = Cm(0.4)
        p_t.paragraph_format.space_before = Pt(1)
        p_t.paragraph_format.space_after = Pt(1)
        render_rich_text(p_t, item, base_size=base_size)

    p_bt6_head = doc.add_paragraph()
    p_bt6_head.paragraph_format.space_before = Pt(3)
    p_bt6_head.paragraph_format.space_after = Pt(2)
    render_rich_text(p_bt6_head, "*BÀI TẬP MINH HOẠ CHƯƠNG VI:*", base_size=base_size, base_bold=True, base_color=COLOR_SUBSECTION)

    for ex in SINH_CHAPTER6_EXERCISES:
        p_ex = doc.add_paragraph()
        p_ex.paragraph_format.left_indent = Cm(0.4)
        p_ex.paragraph_format.space_before = Pt(2)
        p_ex.paragraph_format.space_after = Pt(2)
        render_rich_text(p_ex, f"*{ex['id']}.* {ex['question']}", base_size=base_size)

        if is_teacher:
            add_solution_no_table(doc, ex["solution"])
        elif is_screen:
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.left_indent = Cm(0.8)
            render_rich_text(p_sp, "✎ Trả lời: ...................................................................................................................", base_size=base_size)
        else:
            add_student_writing_lines(doc, num_lines=4)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(12)
    p_end.paragraph_format.space_after = Pt(6)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("--- HẾT ---")
    set_run_style(r_end, font_size=11, bold=True, italic=True, color=COLOR_PRIMARY)

    doc.save(filepath)
    print(f"[OK] Đã tạo thành công file Word: {filepath} ({os.path.getsize(filepath):,} bytes)")

    # Đồng bộ thêm tên file phụ
    if mode == "student":
        alt_name = "De_Cuong_KHTN_6_On_Tap_GK1_HocSinh.docx"
    elif mode == "teacher":
        alt_name = "De_Cuong_KHTN_6_On_Tap_GK1_GiaoVien.docx"
    else:
        alt_name = "De_Cuong_KHTN_6_On_Tap_GK1_DayTrenManHinh.docx"
    alt_path = os.path.join(OUTPUT_DIR, alt_name)
    shutil.copy2(filepath, alt_path)

    return filepath


# ==============================================================================
# 2. HÀM SINH FILE LATEX & BIÊN DỊCH PDF TRÌNH CHIẾU TIVI 16:9
# ==============================================================================

def generate_tivi_presentation_pdf():
    tex_filename = "De_Cuong_GK1_KHTN6_DayTrenManHinh.tex"
    pdf_filename = "De_Cuong_GK1_KHTN6_DayTrenManHinh.pdf"
    tex_path = os.path.join(SCRATCH_DIR, tex_filename)
    pdf_dest = os.path.join(OUTPUT_DIR, pdf_filename)

    tex_lines = []
    # Preamble chuẩn 16:9 widescreen không viền đen
    tex_lines.append(r"""\documentclass{article}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{tabularx}
\usepackage{array}
\usepackage{booktabs}
\usepackage{xcolor}
\usepackage[version=4]{mhchem}
\usepackage[
  paperwidth=29.7cm,
  paperheight=16.706cm,
  top=0.9cm,
  bottom=0.9cm,
  left=1.2cm,
  right=1.2cm,
  footskip=0pt,
  headheight=0pt,
  headsep=0pt
]{geometry}

\pagestyle{empty}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}

\definecolor{brandRed}{RGB}{211, 47, 47}
\definecolor{brandGreen}{RGB}{46, 125, 50}
\definecolor{brandBlue}{RGB}{21, 101, 192}
\definecolor{textColor}{RGB}{20, 20, 20}

\color{textColor}

\begin{document}
""")

    # --------------------------------------------------------------------------
    # SLIDE 0: TIÊU ĐỀ ĐỀ CƯƠNG
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 0: TIÊU ĐỀ ĐỀ CƯƠNG
% ==============================================================================
\begin{center}
\vspace*{0.2cm}
{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}TRƯỜNG TRUNG HỌC CƠ SỞ -- NĂM HỌC 2026 -- 2027}}\\[10pt]
{\fontsize{34}{42}\selectfont \textbf{\color{brandBlue}ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA KÌ I}}\\[8pt]
{\fontsize{28}{36}\selectfont \textbf{\color{brandGreen}MÔN: KHOA HỌC TỰ NHIÊN 6}}\\[8pt]
{\fontsize{22}{28}\selectfont \textit{(Gồm 3 Phân môn: Vật lý -- Hóa học -- Sinh học)}}\\[16pt]
\rule{16cm}{1.5pt}\\[20pt]
{\fontsize{26}{34}\selectfont \textbf{Giáo viên hướng dẫn: Thầy TRẦN VĂN THẠNH -- 0777.470.803}}\\[8pt]
{\fontsize{22}{28}\selectfont \textit{Cơ sở 1: 6/15 Nguyễn Hoàng, Kim Long | Cơ sở 2: 24 Đặng Thái Thân, TP Huế}}\\[16pt]
{\fontsize{24}{30}\selectfont \textbf{CẤU TRÚC: TRẮC NGHIỆM ABCD -- TRẢ LỜI NGẮN -- ĐÚNG/SAI -- TỰ LUẬN}}
\end{center}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 1: LÝ THUYẾT ĐO ĐỘ DÀI
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 1: LÝ THUYẾT ĐO ĐỘ DÀI
% ==============================================================================
{\fontsize{25}{33}\selectfont \textbf{\color{brandRed}I. PHÂN MÔN VẬT LÝ -- BÀI 1: ĐO CHIỀU DÀI}}
\vspace{6pt}

\fontsize{21}{28}\selectfont
\begin{itemize}
  \item \textbf{Đơn vị đo:} Đơn vị đo độ dài hợp pháp của nước ta theo hệ SI là \textbf{mét} (kí hiệu: $\text{m}$).\\
  \textit{Đơn vị khác:} $1\text{ km} = 1000\text{ m}$; $1\text{ m} = 10\text{ dm} = 100\text{ cm} = 1000\text{ mm}$.
  \item \textbf{Dụng cụ đo:} Thước thẳng, thước dây, thước cuộn, thước kẹp\dots
  \item \textbf{Giới hạn đo (GHĐ):} Là độ dài lớn nhất ghi trên thước.
  \item \textbf{Độ chia nhỏ nhất (ĐCNN):} Là độ dài giữa hai vạch chia liên tiếp trên thước.
  \item \textbf{Quy trình 5 bước đo chiều dài:}
  \begin{enumerate}
    \item \textbf{Bước 1 (Ước lượng):} Ước lượng chiều dài cần đo để chọn thước có GHĐ và ĐCNN phù hợp.
    \item \textbf{Bước 2 (Chọn thước):} Chọn thước có $\text{GHĐ} \ge$ vật cần đo và ĐCNN thích hợp.
    \item \textbf{Bước 3 (Đặt thước):} Đặt dọc theo chiều dài, 1 đầu ngang bằng vạch số 0.
    \item \textbf{Bước 4 (Đặt mắt):} Mắt nhìn vuông góc với cạnh thước ở đầu kia của vật.
    \item \textbf{Bước 5 (Đọc và ghi):} Đọc theo vạch chia gần nhất và ghi kèm đơn vị đo theo ĐCNN.
  \end{enumerate}
\end{itemize}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 2: LÝ THUYẾT ĐO KHỐI LƯỢNG & THỜI GIAN
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 2: LÝ THUYẾT ĐO KHỐI LƯỢNG & ĐO THỜI GIAN
% ==============================================================================
{\fontsize{25}{33}\selectfont \textbf{\color{brandRed}I. PHÂN MÔN VẬT LÝ -- ĐO KHỐI LƯỢNG \& THỜI GIAN}}
\vspace{6pt}

\fontsize{21}{28}\selectfont
\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{\color{brandBlue}1. ĐO KHỐI LƯỢNG:} & \textbf{\color{brandGreen}2. ĐO THỜI GIAN:} \\
• \textbf{Khái niệm:} Khối lượng là số đo lượng chất của một vật. & • \textbf{Đơn vị cơ bản (SI):} Là \textbf{giây} (kí hiệu: $\text{s}$). \\
• \textbf{Đơn vị hợp pháp (SI):} Là \textbf{kilôgam} (kí hiệu: $\text{kg}$). & • \textit{Đơn vị khác:} Giờ ($\text{h}$), phút ($\text{min}$), ngày, tháng, năm, mili giây ($\text{ms}$). \\
• \textit{Đơn vị khác:} $1\text{ tấn} = 1000\text{ kg}$; $1\text{ tạ} = 100\text{ kg}$; $1\text{ yến} = 10\text{ kg}$; $1\text{ lạng} = 1\text{ hg} = 100\text{ g}$; $1\text{ kg} = 1000\text{ g}$; $1\text{ g} = 1000\text{ mg}$. & • \textit{Quy đổi:} $1\text{ ngày} = 24\text{ h}$; $1\text{ h} = 60\text{ min} = 3600\text{ s}$; $1\text{ s} = 1000\text{ ms}$. \\
• \textbf{Dụng cụ đo:} Cân (cân đồng hồ, cân điện tử, cân y tế, cân đòn, cân Rô-béc-van\dots). & • \textbf{Dụng cụ đo:} Đồng hồ (bấm giây, đeo tay, treo tường, quả lắc, điện tử\dots). \\
• \textbf{Bước đầu tiên:} Ước lượng khối lượng để chọn cân phù hợp. & • \textbf{Thi đấu thể thao:} Dùng \textbf{đồng hồ bấm giây} để đo chính xác thời gian. \\
\end{tabularx}
\newpage
""")

    # --------------------------------------------------------------------------
    # TRẮC NGHIỆM ABCD: 2 câu / slide
    # --------------------------------------------------------------------------
    # Slide 3: Câu 1 & 2
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 3: TRẮC NGHIỆM CÂU 1 & CÂU 2
% ==============================================================================
\fontsize{27}{36}\selectfont
\textbf{Câu 1.} Đơn vị đo độ dài hợp pháp của nước Việt Nam theo hệ đo lường quốc tế (SI) là gì?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Centimét ($\text{cm}$). & \textbf{B.} Kilômét ($\text{km}$).\\[6pt]
\textbf{C.} Mét ($\text{m}$). & \textbf{D.} Milimét ($\text{mm}$).
\end{tabularx}

\vspace{18pt}

\textbf{Câu 2.} Giới hạn đo (GHĐ) của thước là gì?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Là độ dài nhỏ nhất ghi trên thước. & \textbf{B.} Là độ dài lớn nhất ghi trên thước.\\[6pt]
\textbf{C.} Khoảng cách giữa hai vạch liên tiếp. & \textbf{D.} Là chiều dài thực tế của thước.
\end{tabularx}
\newpage
""")

    # Slide 4: Câu 3 & 4
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 4: TRẮC NGHIỆM CÂU 3 & CÂU 4
% ==============================================================================
\fontsize{27}{36}\selectfont
\textbf{Câu 3.} Dụng cụ nào sau đây \textbf{không phải} là dụng cụ dùng để đo chiều dài?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Thước thẳng. & \textbf{B.} Thước cuộn.\\[6pt]
\textbf{C.} Cân đồng hồ. & \textbf{D.} Thước kẹp.
\end{tabularx}

\vspace{18pt}

\textbf{Câu 4.} Khi đo chiều dài của một vật bằng thước, bước đầu tiên cần thực hiện là gì?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Đặt mắt vuông góc với vạch chia. & \textbf{B.} Ước lượng chiều dài của vật cần đo.\\[6pt]
\textbf{C.} Đặt thước dọc theo chiều dài vật. & \textbf{D.} Đọc kết quả đo theo vạch gần nhất.
\end{tabularx}
\newpage
""")

    # Slide 5: Câu 5 & 6
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 5: TRẮC NGHIỆM CÂU 5 & CÂU 6
% ==============================================================================
\fontsize{27}{36}\selectfont
\textbf{Câu 5.} ĐCNN (độ chia nhỏ nhất) của thước là gì?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Độ dài lớn nhất ghi trên thước. & \textbf{B.} Độ dài giữa hai vạch chia liên tiếp.\\[6pt]
\textbf{C.} Vạch số 0 của thước. & \textbf{D.} Chiều dài của vật cần đo.
\end{tabularx}

\vspace{18pt}

\textbf{Câu 6.} Đơn vị đo khối lượng hợp pháp của Việt Nam theo hệ SI là gì?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Gam ($\text{g}$). & \textbf{B.} Kilôgam ($\text{kg}$).\\[6pt]
\textbf{C.} Miligam ($\text{mg}$). & \textbf{D.} Tấn.
\end{tabularx}
\newpage
""")

    # Slide 6: Câu 7 & 8
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 6: TRẮC NGHIỆM CÂU 7 & CÂU 8
% ==============================================================================
\fontsize{27}{36}\selectfont
\textbf{Câu 7.} Trong các dụng cụ sau đây, dụng cụ nào dùng để đo khối lượng?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Thước cuộn. & \textbf{B.} Bình chia độ.\\[6pt]
\textbf{C.} Cân đồng hồ. & \textbf{D.} Nhiệt kế.
\end{tabularx}

\vspace{18pt}

\textbf{Câu 8.} Khi dùng cân đồng hồ để đo khối lượng của một vật, bước đầu tiên cần làm là gì?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Đặt vật lên đĩa cân. & \textbf{B.} Ước lượng khối lượng của vật cần cân.\\[6pt]
\textbf{C.} Đọc và ghi kết quả đo. & \textbf{D.} Chỉnh kim cân đúng vạch số 0.
\end{tabularx}
\newpage
""")

    # Slide 7: Câu 9 & 10
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 7: TRẮC NGHIỆM CÂU 9 & CÂU 10
% ==============================================================================
\fontsize{27}{36}\selectfont
\textbf{Câu 9.} Đơn vị đo thời gian trong hệ thống đo lường chính thức (SI) của nước ta là gì?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Giờ. & \textbf{B.} Phút.\\[6pt]
\textbf{C.} Giây ($\text{s}$). & \textbf{D.} Ngày.
\end{tabularx}

\vspace{18pt}

\textbf{Câu 10.} Dụng cụ thông dụng nhất dùng để đo thời gian là:
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Thước thẳng. & \textbf{B.} Cân.\\[6pt]
\textbf{C.} Đồng hồ. & \textbf{D.} Nhiệt kế.
\end{tabularx}
\newpage
""")

    # Slide 8: Câu 11
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 8: TRẮC NGHIỆM CÂU 11
% ==============================================================================
\fontsize{28}{38}\selectfont
\textbf{Câu 11.} Trong các đồng hồ sau đây, loại đồng hồ nào thường dùng để đo thời gian trong các cuộc thi đấu thể thao (như chạy $100\text{m}$, $200\text{m}$)?
\vspace{8pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Đồng hồ cát. & \textbf{B.} Đồng hồ bấm giây.\\[10pt]
\textbf{C.} Đồng hồ treo tường. & \textbf{D.} Đồng hồ để bàn.
\end{tabularx}

\vspace{22pt}
{\fontsize{24}{32}\selectfont \textbf{\color{brandBlue}Giải thích:}} \textit{Đồng hồ bấm giây có độ chia rất nhỏ (thường đến $0,01\text{s}$) và khả năng bấm bắt đầu / dừng tức thì, rất phù hợp đo các sự kiện diễn ra nhanh.}
\newpage
""")

    # --------------------------------------------------------------------------
    # TRẢ LỜI NGẮN: Font 20pt/26pt gọn gàng, vừa vặn 1 slide
    # --------------------------------------------------------------------------
    # Slide 9: Trả lời ngắn Câu 1
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 9: TRẢ LỜI NGẮN CÂU 1
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}PHẦN TRẢ LỜI NGẮN -- CÂU 1}}
\vspace{4pt}

\fontsize{20}{26}\selectfont
\textbf{Câu 1.} Thước thẳng có vạch chia từ $0\text{ cm}$ đến $30\text{ cm}$. Giữa vạch 0 và 1 có 5 khoảng bằng nhau.

\vspace{5pt}
\textbf{a)} Độ chia nhỏ nhất (ĐCNN) của thước là bao nhiêu mi-li-mét?
\newline \textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$ \hfill \textit{(Đáp số: $10\text{mm} : 5 = 2\text{mm}$)}

\vspace{4pt}
\textbf{b)} Giới hạn đo (GHĐ) của cây thước này là bao nhiêu mét?
\newline \textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$ \hfill \textit{(Đáp số: $30\text{cm} = 0,3\text{m}$)}

\vspace{4pt}
\textbf{c)} Đo chiều dài bút bi ghi: $22,50\text{cm}$; $22,2\text{cm}$; $22,25\text{cm}$. Kết quả nào ghi đúng quy tắc?
\newline \textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$ \hfill \textit{(Đáp số: $22,2\text{cm}$)}

\vspace{4pt}
\textbf{d)} Đầu kia quyển sách nằm tại vạch số 3 sau vạch $22\text{cm}$. Chiều dài quyển sách là bao nhiêu?
\newline \textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$ \hfill \textit{(Đáp số: $22 + 3 \times 0,2 = 22,6\text{cm}$)}
\newpage
""")

    # Slide 10: Trả lời ngắn Câu 2
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 10: TRẢ LỜI NGẮN CÂU 2
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}PHẦN TRẢ LỜI NGẮN -- CÂU 2}}
\vspace{4pt}

\fontsize{20}{26}\selectfont
\textbf{Câu 2.} Cân đồng hồ có vạch chia từ 0 và ghi $2\text{kg}$. Giữa vạch 0 và $100\text{g}$ có 10 khoảng bằng nhau.

\vspace{5pt}
\textbf{a)} Độ chia nhỏ nhất (ĐCNN) của cân này là bao nhiêu kg?
\newline \textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$ \hfill \textit{(Đáp số: $100\text{g} : 10 = 10\text{g} = 0,01\text{kg}$)}

\vspace{4pt}
\textbf{b)} Giới hạn đo (GHĐ) của cân này là bao nhiêu kg?
\newline \textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$ \hfill \textit{(Đáp số: $2\text{kg}$)}

\vspace{4pt}
\textbf{c)} Cân túi cam và ghi: $820\text{g}$; $855\text{g}$; $870\text{g}$. Kết quả nào ghi \textbf{không đúng} quy tắc đo?
\newline \textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$ \hfill \textit{(Đáp số: $855\text{g}$ vì không chia hết cho $10\text{g}$)}

\vspace{4pt}
\textbf{d)} Kim cân nằm tại vạch số 7 sau vạch $1,3\text{kg}$. Khối lượng của túi bánh là bao nhiêu?
\newline \textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$ \hfill \textit{(Đáp số: $1300 + 7 \times 10 = 1370\text{g} = 1,37\text{kg}$)}
\newpage
""")

    # --------------------------------------------------------------------------
    # TRẮC NGHIỆM ĐÚNG SAI: 1 câu / slide
    # --------------------------------------------------------------------------
    # Slide 11: Đúng / Sai Câu 1
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 11: TRẮC NGHIỆM ĐÚNG SAI CÂU 1
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}PHẦN TRẮC NGHIỆM ĐÚNG SAI -- CÂU 1}}
\vspace{4pt}

\fontsize{22}{29}\selectfont
\textbf{Câu 1.} Học sinh dùng thước thẳng có $\text{GHĐ} = 1\text{m}$ và $\text{ĐCNN} = 1\text{mm}$ đo chiều dài bàn học. Vạch 0 trùng 1 đầu bàn, đầu kia nằm ở vạch $115,5\text{cm}$.

\vspace{8pt}
\textbf{a)} Cây thước học sinh này chọn là dụng cụ phù hợp nhất để đo chiếc bàn.\hfill $[\quad]$\\[4pt]
\textbf{b)} Kết quả đo chiều dài chiếc bàn ghi $115,5\text{cm}$ là đúng quy tắc và chính xác.\hfill $[\quad]$\\[4pt]
\textbf{c)} Để đo chính xác và thuận tiện hơn, nên thay bằng thước cuộn/dây có $\text{GHĐ} > 115,5\text{cm}$.\hfill $[\quad]$\\[4pt]
\textbf{d)} Độ chia nhỏ nhất của thước là $1\text{mm}$ cho phép đọc kết quả chính xác đến mi-li-mét.\hfill $[\quad]$

\vspace{10pt}
{\fontsize{20}{26}\selectfont \textbf{\color{brandGreen}Chốt đáp án:}} \textbf{a) \color{brandRed}Sai} \quad | \quad \textbf{b) \color{brandRed}Sai} \quad | \quad \textbf{c) \color{brandGreen}Đúng} \quad | \quad \textbf{d) \color{brandGreen}Đúng}
\newpage
""")

    # Slide 12: Đúng / Sai Câu 2
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 12: TRẮC NGHIỆM ĐÚNG SAI CÂU 2
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}PHẦN TRẮC NGHIỆM ĐÚNG SAI -- CÂU 2}}
\vspace{4pt}

\fontsize{22}{29}\selectfont
\textbf{Câu 2.} An dùng cân đồng hồ có $\text{GHĐ} = 2\text{kg}$ và $\text{ĐCNN} = 10\text{g}$ để cân hộp bánh có khối lượng tịnh ghi $750\text{g}$. Khi đặt hộp bánh lên cân, kim cân chỉ vạch $760\text{g}$.

\vspace{8pt}
\textbf{a)} Chiếc cân An đang dùng hoàn toàn phù hợp để đo khối lượng của hộp bánh này.\hfill $[\quad]$\\[4pt]
\textbf{b)} Khối lượng thực tế của hộp bánh đúng $750\text{g}$ như số ghi trên vỏ hộp.\hfill $[\quad]$\\[4pt]
\textbf{c)} Sai số của phép đo trong trường hợp này lớn hơn ĐCNN của cân.\hfill $[\quad]$\\[4pt]
\textbf{d)} Nếu thay hộp bánh bằng bao gạo nặng $2,5\text{kg}$ thì cân trên vẫn đo tốt và bền.\hfill $[\quad]$

\vspace{10pt}
{\fontsize{20}{26}\selectfont \textbf{\color{brandGreen}Chốt đáp án:}} \textbf{a) \color{brandGreen}Đúng} \quad | \quad \textbf{b) \color{brandRed}Sai} \quad | \quad \textbf{c) \color{brandRed}Sai} \quad | \quad \textbf{d) \color{brandRed}Sai}
\newpage
""")

    # --------------------------------------------------------------------------
    # TỰ LUẬN: 1 câu / slide
    # --------------------------------------------------------------------------
    # Slide 13: Tự luận Câu 1
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 13: TỰ LUẬN CÂU 1
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}BÀI TẬP TỰ LUẬN -- CÂU 1}}
\vspace{6pt}

\fontsize{21}{28}\selectfont
\textbf{Câu 1.} Trên bao bì gói bánh COSY có ghi: \textbf{Khối lượng tịnh $500\text{g}$}. Bảng giá trị dinh dưỡng ghi: cứ $10\text{g}$ bánh sẽ cung cấp khoảng $100\text{cal}$.

\vspace{6pt}
\textbf{a)} Con số ``$500\text{g}$'' đó cho em biết điều gì?
\newline \textbf{b)} Bạn Lan đã ăn hết toàn bộ gói bánh này. Tính lượng năng lượng bạn Lan đã nạp vào cơ thể?

\vspace{12pt}
{\fontsize{21}{28}\selectfont \textbf{\color{brandRed}Hướng dẫn giải chi tiết:}}
\begin{itemize}
  \item a) Con số ``$500\text{g}$'' là khối lượng tịnh của riêng phần bánh bên trong (không tính vỏ gói).
  \item b) Số năng lượng bạn Lan đã nạp vào cơ thể:
  \[
  E = \frac{500\text{ g}}{10\text{ g}} \times 100\text{ cal} = 50 \times 100\text{ cal} = 5000\text{ cal} \quad (\text{hay } 5\text{ kcal}).
  \]
\end{itemize}
\newpage
""")

    # Slide 14: Tự luận Câu 2
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 14: TỰ LUẬN CÂU 2
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}BÀI TẬP TỰ LUẬN -- CÂU 2}}
\vspace{6pt}

\fontsize{21}{28}\selectfont
\textbf{Câu 2.} Trên một thùng bánh COSY chứa 20 gói bánh cùng loại, ghi: \textbf{``Khối lượng tịnh của mỗi gói: $100\text{g}$''}.

\vspace{6pt}
\textbf{a)} Con số $100\text{g}$ ghi trên vỏ gói bánh cho biết điều gì?
\newline \textbf{b)} Biết khối lượng của thùng giấy là $200\text{g}$. Tổng khối lượng cả bánh và thùng giấy là bao nhiêu kg?

\vspace{12pt}
{\fontsize{21}{28}\selectfont \textbf{\color{brandRed}Hướng dẫn giải chi tiết:}}
\begin{itemize}
  \item a) Con số $100\text{g}$ là khối lượng tịnh của riêng phần bánh bên trong một gói.
  \item b) Tổng khối lượng cả bánh và thùng giấy:
  \begin{itemize}
    \item Khối lượng của 20 gói bánh: $m_{\text{bánh}} = 20 \times 100\text{ g} = 2000\text{ g} = 2\text{ kg}$.
    \item Tổng khối lượng: $m_{\text{tổng}} = 2000\text{ g} + 200\text{ g} = 2200\text{ g} = 2,2\text{ kg}$.
  \end{itemize}
\end{itemize}
\newpage
""")

    # --------------------------------------------------------------------------
    # PHÂN MÔN HÓA HỌC: 3 SLIDES
    # --------------------------------------------------------------------------
    # Slide 15: Hóa học - Chất & Vật thể, Tính chất vật lí/hóa học
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 15: HÓA HỌC - CHẤT VÀ VẬT THỂ
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}II. PHÂN MÔN HÓA HỌC -- CHẤT VÀ TÍNH CHẤT CỦA CHẤT}}
\vspace{6pt}

\fontsize{20}{26}\selectfont
\textbf{1. Chất tạo nên vật thể (Ở đâu có vật thể là ở đó có chất):}
\begin{itemize}
  \item \textbf{Vật sống (hữu sinh):} Có trao đổi chất, lớn lên, sinh sản (\textit{VD: con sư tử, em bé, cây mai}).
  \item \textbf{Vật không sống (vô sinh):} Không trao đổi chất, không lớn lên (\textit{VD: bánh mì, cầu Long Biên}).
  \item \textbf{Vật thể tự nhiên:} Có sẵn trong tự nhiên (\textit{VD: núi đá vôi, con sư tử, mủ cao su}).
  \item \textbf{Vật thể nhân tạo:} Do con người tạo ra (\textit{VD: cầu Long Biên, cái bàn, bánh mì}).
\end{itemize}

\vspace{5pt}
\textbf{2. Tính chất của chất:}
\begin{itemize}
  \item \textbf{Tính chất vật lí:} Thể (rắn, lỏng, khí), màu sắc, mùi vị, nhiệt độ sôi, nóng chảy, tính tan... \textit{(Không biến đổi thành chất mới, VD: nước đá tan thành nước lỏng)}.
  \item \textbf{Tính chất hóa học:} Sự biến đổi từ chất này thành chất khác \textit{(Có chất mới tạo thành, VD: đường đun cháy thành than và hơi nước)}.
\end{itemize}
\newpage
""")

    # Slide 16: Hóa học - 3 Thể của chất & Chuyển thể
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 16: HÓA HỌC - CÁC THỂ CỦA CHẤT VÀ SỰ CHUYỂN THỂ
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}CÁC THỂ CỦA CHẤT VÀ SỰ CHUYỂN THỂ}}
\vspace{2pt}

\fontsize{18}{23}\selectfont
\begin{tabularx}{\linewidth}{|m{3.5cm}|m{7.2cm}|X|}
\hline
\rowcolor{gray!15}
\textbf{Thể của chất} & \textbf{Đặc điểm hình dạng \& Thể tích} & \textbf{Khả năng nén \& Chảy} \\
\hline
\textbf{Thể rắn} & Hình dạng và thể tích cố định & Rất khó bị nén; không tự chảy. \\
\hline
\textbf{Thể lỏng} & Thể tích cố định; hình dạng phần vật chứa & Khó bị nén; chảy tràn, rót được. \\
\hline
\textbf{Thể khí (hơi)} & Chiếm toàn bộ thể tích bình chứa, không có hình dạng xác định & Rất dễ bị nén; lan tỏa mọi hướng. \\
\hline
\end{tabularx}

\vspace{5pt}
{\fontsize{19}{25}\selectfont \textbf{\color{brandBlue}Các quá trình chuyển thể:}}
\begin{itemize}
  \setlength{\itemsep}{1pt}
  \item \textbf{Nóng chảy:} Rắn $\rightarrow$ Lỏng (ở nhiệt độ nóng chảy xác định, VD: băng tan ở $0^\circ\text{C}$).
  \item \textbf{Đông đặc:} Lỏng $\rightarrow$ Rắn (cùng nhiệt độ với nóng chảy, VD: nước đóng băng ở $0^\circ\text{C}$).
  \item \textbf{Bay hơi:} Hóa hơi trên bề mặt chất lỏng (xảy ra ở \textbf{mọi nhiệt độ}).
  \item \textbf{Sự sôi:} Hóa hơi cả trên bề mặt và trong lòng chất lỏng (ở \textbf{nhiệt độ sôi xác định}, VD: $100^\circ\text{C}$).
  \item \textbf{Ngưng tụ:} Khí (hơi) $\rightarrow$ Lỏng (xảy ra ở \textbf{mọi nhiệt độ}).
\end{itemize}
\newpage
""")

    # Slide 17: Hóa học - Oxygen & Vai trò
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 17: HÓA HỌC - TÍNH CHẤT VÀ VAI TRÒ CỦA OXYGEN
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}TÍNH CHẤT VÀ VAI TRÒ CỦA OXYGEN ($\ce{O2}$)}}
\vspace{6pt}

\fontsize{21}{29}\selectfont
\textbf{1. Tính chất vật lí của Oxygen:}
\begin{itemize}
  \item Điều kiện thường: Là chất khí không màu, không mùi, không vị.
  \item Ít tan trong nước, nặng hơn không khí ($d = 32/29 \approx 1,1$).
  \item Hóa lỏng ở $-183^\circ\text{C}$, hóa rắn ở $-218^\circ\text{C}$. Ở thể lỏng và rắn có màu \textbf{xanh nhạt}.
\end{itemize}

\vspace{8pt}
\textbf{2. Tầm quan trọng của Oxygen trong đời sống và sản xuất:}
\begin{itemize}
  \item \textbf{Duy trì sự sống:} Oxygen cần thiết cho quá trình hô hấp của con người, động vật, thực vật.
  \item \textbf{Duy trì sự cháy:} Cần thiết cho quá trình đốt cháy các nhiên liệu (củi, than đá, khí gas, xăng dầu) trong đun nấu, nhiệt điện, giao thông vận tải, thắp sáng\dots
\end{itemize}
\newpage
""")

    # --------------------------------------------------------------------------
    # PHÂN MÔN SINH HỌC: 4 SLIDES
    # --------------------------------------------------------------------------
    # Slide 18: Sinh học - Chương I (Kính lúp & Kính hiển vi)
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 18: SINH HỌC - CHƯƠNG I: DỤNG CỤ QUAN SÁT
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}III. PHÂN MÔN SINH HỌC -- CHƯƠNG I: DỤNG CỤ QUAN SÁT}}
\vspace{4pt}

\fontsize{19}{25}\selectfont
\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{1. Kính lúp:} & \textbf{2. Kính hiển vi quang học:} \\
• Cấu tạo: Khung kính, tay cầm, tấm kính hội tụ. & • Hệ thống phóng đại: \textbf{Thị kính} (đặt mắt vào) và \textbf{Vật kính} (sát mẫu vật). \\
• Sử dụng: Cầm tay cầm, đưa kính từ gần ra xa mẫu vật đến khi nhìn rõ nét. & • Điều chỉnh rõ nét: Dùng \textbf{ốc to (sơ cấp)} nâng hạ ống kính, sau đó dùng \textbf{ốc nhỏ (vi cấp)} lấy nét thật nét. \\
• Bảo quản: Lau bằng khăn mềm, tránh trầy xước. & • Nơi đặt mắt nhìn mẫu vật: \textbf{Thị kính}. \\
\end{tabularx}

\vspace{8pt}
{\fontsize{21}{27}\selectfont \textbf{\color{brandBlue}Giải đáp bài tập minh họa Chương I:}}
\begin{itemize}
  \item \textbf{BT1:} Bộ phận chỉnh hình ảnh rõ nét: \textbf{Ốc vi cấp (núm vi chỉnh)}.
  \item \textbf{BT2:} Hệ thống phóng đại gồm: \textbf{Thị kính} và \textbf{Vật kính}.
  \item \textbf{BT3:} Lau kính bằng khăn mềm, tránh để mặt kính tiếp xúc bề mặt nhám.
  \item \textbf{BT4:} Bộ phận để mắt nhìn vào mẫu vật: \textbf{Thị kính}.
\end{itemize}
\newpage
""")

    # Slide 19: Sinh học - Chương V: Tế bào (Cấu tạo & Phân chia)
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 19: CHƯƠNG V: TẾ BÀO – ĐƠN VỊ CƠ SỞ CỦA SỰ SỐNG
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}CHƯƠNG V: TẾ BÀO -- ĐƠN VỊ CƠ SỞ CỦA SỰ SỐNG}}
\vspace{6pt}

\fontsize{20}{27}\selectfont
\textbf{1. Khái niệm và chức năng:}
\begin{itemize}
  \item Tế bào là đơn vị cơ sở cấu tạo nên mọi cơ thể sống và thực hiện các chức năng sống.
\end{itemize}

\vspace{3pt}
\textbf{2. Ba thành phần chính cấu tạo nên tế bào:}
\begin{itemize}
  \item \textbf{Màng tế bào (màng sinh chất):} Bao bọc, bảo vệ và kiểm soát trao đổi chất.
  \item \textbf{Chất tế bào:} Chứa các bào quan, nơi diễn ra hầu hết hoạt động sống.
  \item \textbf{Nhân tế bào (hoặc Vùng nhân):} Chứa vật chất di truyền, điều khiển mọi hoạt động.
\end{itemize}

\vspace{3pt}
\textbf{3. Bào quan quang hợp \& Sự sinh sản của tế bào:}
\begin{itemize}
  \item \textbf{Lục lạp:} Bào quan thực hiện chức năng quang hợp ở tế bào thực vật (cây xanh).
  \item \textbf{Phân chia tế bào:} $1 \text{ tế bào} \rightarrow 2 \rightarrow 4 \rightarrow 8 \rightarrow \dots \rightarrow 2^n \text{ tế bào con}$. Giúp cơ thể lớn lên và thay thế tế bào già, tổn thương.
\end{itemize}
\newpage
""")

    # Slide 20: Sinh học - Bài tập minh họa Chương V
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 20: BÀI TẬP MINH HỌA CHƯƠNG V (TẾ BÀO)
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}BÀI TẬP MINH HỌA VỀ TẾ BÀO (CHƯƠNG V)}}
\vspace{4pt}

\fontsize{19}{25}\selectfont
\textbf{BT1. So sánh tế bào vảy hành (đa giác, cố định) và tế bào trứng cá (tròn, tách rời):}
\begin{itemize}
  \item \textit{Tế bào vảy hành (thực vật):} Có \textbf{thành tế bào} bằng cellulose cứng cáp $\rightarrow$ Giữ hình đa giác và gắn chặt nhau.
  \item \textit{Tế bào trứng cá (động vật):} Không có thành tế bào cứng, màng mềm dẻo, sức căng bề mặt tạo hình cầu; nằm rời nhau.
\end{itemize}

\vspace{3pt}
\textbf{BT2. Kích thước tế bào \& Dụng cụ quan sát:}
\begin{itemize}
  \item Đa số kích thước hiển vi ($1 - 100\,\mu\text{m}$) $\rightarrow$ Dùng \textbf{kính hiển vi quang học}.
  \item Tế bào kích thước lớn (trứng cá, tép bưởi, trứng gà) $\rightarrow$ Dùng \textbf{kính lúp} hoặc \textbf{mắt thường}.
\end{itemize}

\vspace{3pt}
\textbf{BT3 \& BT4. Hình dạng tế bào tương ứng với chức năng:}
\begin{itemize}
  \item Hình dạng phù hợp với \textbf{chức năng}: Hồng cầu hình đĩa lõm hai mặt (tăng diện tích gắn $\ce{O2}$); Tế bào thần kinh có sợi nhánh, sợi trục dài (dẫn truyền xung thần kinh).
  \item Các dạng tế bào: Hình cầu (trứng cá), hình đĩa (hồng cầu), hình đa giác (biểu bì hành), hình thoi (cơ trơn).
\end{itemize}
\newpage
""")

    # Slide 21: Sinh học - Chương VI: Từ tế bào đến cơ thể
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 21: CHƯƠNG VI: TỪ TẾ BÀO ĐẾN CƠ THỂ
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}CHƯƠNG VI: TỪ TẾ BÀO ĐẾN CƠ THỂ}}
\vspace{4pt}

\fontsize{19}{25}\selectfont
\textbf{1. Phân biệt cơ thể đơn bào và cơ thể đa bào:}
\begin{itemize}
  \item \textbf{Cơ thể đơn bào:} Chỉ gồm một tế bào duy nhất nhưng thực hiện trọn vẹn chức năng sống (VD: Vi khuẩn $E.coli$, trùng roi xanh, trùng giày, tảo lam, nấm men).
  \item \textbf{Cơ thể đa bào:} Gồm nhiều tế bào phân hóa thành mô, cơ quan, hệ cơ quan (VD: Cây lúa, hoa mai, con người, con sư tử).
\end{itemize}

\vspace{3pt}
\textbf{2. Giải đáp BT1: Phân biệt nhân sơ hay nhân thực:}
\begin{itemize}
  \item \textbf{Vi khuẩn $E.coli$:} Tế bào \textbf{nhân sơ} (chưa có màng nhân bao bọc vật chất di truyền).
  \item \textbf{Tế bào biểu bì vảy hành \& Tế bào trứng cá:} Tế bào \textbf{nhân thực} (đã có màng nhân hoàn chỉnh).
\end{itemize}

\vspace{3pt}
\textbf{3. Giải đáp BT2: Các thành phần của cơ thể trùng roi xanh:}
\begin{itemize}
  \item Gồm 7 thành phần chính: 1. \textbf{Roi bơi}; 2. \textbf{Điểm mắt}; 3. \textbf{Hạt diệp lục (lục lạp)}; 4. \textbf{Màng tế bào}; 5. \textbf{Chất tế bào}; 6. \textbf{Nhân tế bào}; 7. \textbf{Không bào co bóp}.
\end{itemize}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 22: BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ CƯƠNG
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 22: BẢNG TỔNG HỢP ĐÁP ÁN TOÀN BỘ ĐỀ CƯƠNG
% ==============================================================================
\begin{center}
{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}BẢNG TỔNG HỢP ĐÁP ÁN ÔN TẬP KIỂM TRA GIỮA KÌ I}}
\end{center}
\vspace{2pt}

\fontsize{19}{24}\selectfont
\textbf{\color{brandBlue}1. ĐÁP ÁN TRẮC NGHIỆM ABCD (VẬT LÝ):}
\vspace{2pt}

\begin{center}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{C1} & \textbf{C2} & \textbf{C3} & \textbf{C4} & \textbf{C5} & \textbf{C6} & \textbf{C7} & \textbf{C8} & \textbf{C9} & \textbf{C10} & \textbf{C11} \\
\hline
\textbf{\color{brandRed}C} & \textbf{\color{brandRed}B} & \textbf{\color{brandRed}C} & \textbf{\color{brandRed}B} & \textbf{\color{brandRed}B} & \textbf{\color{brandRed}B} & \textbf{\color{brandRed}C} & \textbf{\color{brandRed}B} & \textbf{\color{brandRed}C} & \textbf{\color{brandRed}C} & \textbf{\color{brandRed}B} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\textbf{\color{brandBlue}2. ĐÁP ÁN ĐÚNG -- SAI (VẬT LÝ):}
\vspace{2pt}

\begin{center}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu hỏi} & \textbf{Ý a} & \textbf{Ý b} & \textbf{Ý c} & \textbf{Ý d} \\
\hline
\textbf{Câu 1 (Thước đo bàn)} & \textbf{\color{brandRed}Sai} & \textbf{\color{brandRed}Sai} & \textbf{\color{brandGreen}Đúng} & \textbf{\color{brandGreen}Đúng} \\
\hline
\textbf{Câu 2 (Cân hộp bánh)} & \textbf{\color{brandGreen}Đúng} & \textbf{\color{brandRed}Sai} & \textbf{\color{brandRed}Sai} & \textbf{\color{brandRed}Sai} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\textbf{\color{brandBlue}3. ĐÁP SỐ TRẢ LỜI NGẮN \& TỰ LUẬN:}
\vspace{1pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
• \textbf{TLN C1:} a) $2\text{mm}$; b) $0,3\text{m}$; c) $22,2\text{cm}$; d) $22,6\text{cm}$. & • \textbf{TLN C2:} a) $0,01\text{kg}$ ($10\text{g}$); b) $2\text{kg}$; c) $855\text{g}$; d) $1,37\text{kg}$.\\
• \textbf{Tự luận C1:} a) Khối lượng tịnh bánh; b) $5000\text{cal}$ ($5\text{kcal}$). & • \textbf{Tự luận C2:} a) Khối lượng tịnh 1 gói; b) Tổng $2,2\text{kg}$.
\end{tabularx}

\end{document}
""")

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write("".join(tex_lines))
    print(f"[OK] Đã ghi file LaTeX: {tex_path}")

    # Biên dịch pdflatex 2 lần để ổn định layout
    cmd = ["pdflatex", "-interaction=nonstopmode", "-output-directory", SCRATCH_DIR, tex_path]
    res1 = subprocess.run(cmd, capture_output=True, text=True)
    res2 = subprocess.run(cmd, capture_output=True, text=True)

    pdf_scratch = os.path.join(SCRATCH_DIR, pdf_filename)
    if os.path.exists(pdf_scratch):
        shutil.copy2(pdf_scratch, pdf_dest)
        print(f"[OK] Đã biên dịch và copy PDF thành công: {pdf_dest} ({os.path.getsize(pdf_dest):,} bytes)")

        # Copy thêm tên phụ
        alt_dest = os.path.join(OUTPUT_DIR, "De_Cuong_KHTN_6_On_Tap_GK1_DayTrenManHinh.pdf")
        shutil.copy2(pdf_dest, alt_dest)
        return pdf_dest
    else:
        print("[ERROR] Không tìm thấy file PDF được tạo. Log:")
        print(res2.stdout[-1500:])
        return None

# ==============================================================================
# HÀM MAIN THỰC THI TOÀN BỘ
# ==============================================================================

def main():
    print("=" * 75)
    print("BẮT ĐẦU TẠO 3 FILE CHO ĐỀ CƯƠNG ÔN TẬP GIỮA KÌ I KHTN 6 (2026-2027)")
    print("1. FILE HỌC SINH (.docx)")
    print("2. FILE GIÁO VIÊN (.docx có lời giải chi tiết)")
    print("3. FILE DẠY TRÊN MÀN HÌNH (Slide PDF 16:9 widescreen cho Tivi 65 inch)")
    print("=" * 75)

    f_hs = generate_word_document(mode="student")
    f_gv = generate_word_document(mode="teacher")
    f_screen_docx = generate_word_document(mode="screen")
    f_screen_pdf = generate_tivi_presentation_pdf()

    # Đồng bộ OneDrive nếu có
    all_files = [f_hs, f_gv, f_screen_docx, f_screen_pdf]
    print("\n[+] Đồng bộ các file sang OneDrive (nếu có):")
    for od in ONEDRIVE_DIRS:
        if os.path.exists(od):
            for src_path in all_files:
                if src_path and os.path.exists(src_path):
                    shutil.copy2(src_path, os.path.join(od, os.path.basename(src_path)))
                    print(f"    -> Đã copy {os.path.basename(src_path)} vào {od}")

    print("\n" + "=" * 75)
    print("HOÀN THÀNH TẤT CẢ 3 FILE THEO ĐÚNG CHUẨN FORM ĐỀ KIỂM TRA!")
    print("=" * 75)

if __name__ == "__main__":
    main()
