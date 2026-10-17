# -*- coding: utf-8 -*-
"""
Script sinh tài liệu LaTeX (.tex) và biên dịch ra PDF cho:
Đề kiểm tra Giữa học kì I Hóa học 10 (Mã đề 101 / Đề thêm 01)
Năm học 2026 - 2027 - Thầy TRẦN VĂN THẠNH

Sinh 3 file:
1. De_Kiem_Tra_Hoa_10_Giua_HK1_DeThem01_HocSinh.tex & .pdf
2. De_Kiem_Tra_Hoa_10_Giua_HK1_DeThem01_GiaoVien.tex & .pdf
3. De_Kiem_Tra_Hoa_10_Giua_HK1_DeThem01_DayTrenManHinh.tex & .pdf
"""

import os
import sys
import re
import subprocess
import shutil

sys.stdout.reconfigure(encoding='utf-8')
ROOT_DIR = r"c:\Antigravity_Thanh"
OUT_DIR = os.path.join(ROOT_DIR, "San_Pham")
IMG_DIR = os.path.join(OUT_DIR, "images_hoa10")
ONEDRIVE_DIR = r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\GK1\10"
os.makedirs(OUT_DIR, exist_ok=True)

sys.path.append(os.path.join(ROOT_DIR, "scripts"))
from hoa10_de01_data import (
    EXAM_INFO, PART_1_QUESTIONS, PART_2_QUESTIONS,
    PART_3_QUESTIONS, PART_4_QUESTIONS
)

def clean_for_latex(text):
    """Xử lý triệt để ký tự đặc biệt, công thức hóa học, hạt nhân cho LaTeX."""
    if not text:
        return ""
    t = text
    
    # 1. Ký tự Unicode đặc biệt
    t = t.replace('α', r'$\alpha$')
    t = t.replace('►', r'$\triangleright$')
    
    # 2. Xử lý đồng vị hạt nhân dạng ^A_ZX
    t = re.sub(r'\^([0-9]+)_([0-9]+)([A-Za-z]+)', r'$^{\1}_{\2}\\text{\3}$', t)
    t = re.sub(r'\^A_ZX', r'$^{A}_{Z}\\text{X}$', t)
    
    # 3. Xử lý Subscript ~...~
    t = re.sub(r'~([^~]+)~', r'$_{\1}$', t)
    
    # 4. Xử lý Superscript ^...^
    t = re.sub(r'\^([^\^]+)\^', r'$^{\1}$', t)
    
    # 5. Xử lý ký hiệu % nếu chưa có \ trước
    t = re.sub(r'(?<!\\)%', r'\%', t)
    
    # 6. Sửa lỗi dính toán học kép $$...$$ trong math mode
    t = t.replace('$$', '$')
    
    return t

def generate_tex(mode="hs"):
    is_teacher = (mode == "gv")
    is_screen = (mode == "screen")

    font_size = "13pt" if is_screen else "11pt"
    line_stretch = "1.22" if is_screen else "1.06"

    title_sub = ""
    if is_teacher:
        title_sub = r"{\color{red!80!black}\textbf{(BẢN GIÁO VIÊN -- CÓ ĐÁP ÁN \& LỜI GIẢI CHI TIẾT)}}\par\vspace{3pt}"
    elif is_screen:
        title_sub = r"{\color{purple!80!black}\textbf{(BẢN TRÌNH CHIẾU / GIẢNG DẠY MÀN HÌNH -- CÓ CHỪA CHỖ TRỐNG)}}\par\vspace{3pt}"

    header_template = r"""\documentclass[__FONT_SIZE__,a4paper]{article}
\usepackage[utf8]{vietnam}
\usepackage{amsmath,amssymb}
\usepackage{geometry}
\geometry{top=1.5cm,bottom=1.2cm,left=1.4cm,right=1.4cm,headheight=28pt,headsep=12pt,footskip=18pt}
\usepackage{tikz}
\usepackage{xcolor}
\usepackage{graphicx}
\usepackage{fancyhdr}
\usepackage{tabularx}
\usepackage{enumitem}

\renewcommand{\baselinestretch}{__LINE_STRETCH__}
\setlength{\parskip}{2pt}
\setlength{\parindent}{0pt}

\definecolor{hfRed}{RGB}{211,47,47}
\definecolor{hfGreen}{RGB}{139,195,144}
\definecolor{myGreen}{RGB}{46,125,50}
\definecolor{myBlue}{RGB}{21,101,192}
\definecolor{mySub}{RGB}{0,105,92}
\definecolor{myPurple}{RGB}{106,27,154}

\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small\itshape\color{hfRed} Tài liệu lưu hành nội bộ}
\fancyhead[R]{\small\itshape\color{hfRed} Thầy \textbf{TRẦN VĂN THẠNH} - 0777.470.803}
\renewcommand{\headrulewidth}{0.8pt}
\renewcommand{\headrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \headrulewidth\hfill}}

\fancyfoot[L]{\small\itshape\color{hfRed} CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế \quad|\quad CS2: 24 Đặng Thái Thân, TP Huế}
\fancyfoot[R]{\small\bfseries\color{hfRed} Trang \thepage}
\renewcommand{\footrulewidth}{0.8pt}
\renewcommand{\footrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \footrulewidth\hfill}}

\begin{document}

% --- HEADER BLOCK ---
\begin{tabularx}{\textwidth}{@{}l X r@{}}
\textbf{\small SỞ GIÁO DỤC VÀ ĐÀO TẠO} & & \textbf{\small ĐỀ KIỂM TRA GIỮA HỌC KÌ I} \\
\textbf{\normalsize TRƯỜNG THPT CHUYÊN / THPT} & & \textbf{\normalsize MÔN: HÓA HỌC - LỚP 10} \\
\textit{\small (Đề thi gồm có 03 trang)} & & \textit{\small Thời gian làm bài: 45 phút (không kể phát đề)}
\end{tabularx}

\begin{center}
__TITLE_SUB__
\textbf{Họ và tên thí sinh:} \dotfill \quad \textbf{Lớp:} \dots\dots\dots \quad \textbf{Số báo danh:} \dots\dots\dots \hfill \textbf{\large Mã đề 101}
\end{center}
\vspace{-4pt}
{\color{gray!50}\hrule height 0.8pt}
\vspace{6pt}
"""
    tex = header_template.replace("__FONT_SIZE__", font_size).replace("__LINE_STRETCH__", line_stretch).replace("__TITLE_SUB__", title_sub)

    # -------------------------------------------------------------------------
    # PHẦN I
    # -------------------------------------------------------------------------
    tex += r"""
{\color{myGreen}\bfseries\large PHẦN I (3,0 ĐIỂM). Câu trắc nghiệm nhiều phương án lựa chọn.}

\textit{\small Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{3pt}
"""

    for q in PART_1_QUESTIONS:
        qid = q["id"]
        content_tex = clean_for_latex(q['content'])
        tex += f"\n\\textbf{{Câu {qid}.}} {content_tex}\n\n"

        if q.get("options_type") == "images":
            tex += r"""\begin{center}
\begin{tabularx}{\textwidth}{*{4}{>{\centering\arraybackslash}X}}
"""
            opt_cells = []
            for opt in q["options"]:
                lbl = opt["label"]
                img_path = f"images_hoa10/{opt['image']}"
                is_c = (is_teacher and lbl == q["correct"])
                if is_c:
                    opt_cells.append(f"\\textbf{{\\color{{red}} {lbl}.}} \\raisebox{{-0.3\\height}}{{\\includegraphics[height=0.9cm]{{{img_path}}}}}")
                else:
                    opt_cells.append(f"\\textbf{{{lbl}.}} \\raisebox{{-0.3\\height}}{{\\includegraphics[height=0.9cm]{{{img_path}}}}}")
            tex += " & ".join(opt_cells) + r""" \\
\end{tabularx}
\end{center}
"""
        else:
            opts = [clean_for_latex(opt) for opt in q["options"]]
            lens = [len(opt) for opt in opts]
            max_len = max(lens)

            if max_len <= 25:
                # 4 trên 1 dòng
                row = []
                for opt in opts:
                    lbl = opt[:1]
                    is_c = (is_teacher and lbl == q["correct"])
                    if is_c:
                        row.append(f"\\textbf{{\\color{{red}} {opt}}}")
                    else:
                        row.append(f"\\textbf{{{lbl}.}} {opt[3:]}")
                tex += r"\begin{tabularx}{\textwidth}{@{}XXXX@{}}" + "\n"
                tex += " & ".join(row) + r" \\" + "\n"
                tex += r"\end{tabularx}" + "\n"
            elif max_len <= 50:
                # 2x2
                tex += r"\begin{tabularx}{\textwidth}{@{}XX@{}}" + "\n"
                for r_idx in range(2):
                    opt1 = opts[r_idx * 2]
                    opt2 = opts[r_idx * 2 + 1]
                    lbl1, lbl2 = opt1[:1], opt2[:1]
                    c1 = f"\\textbf{{\\color{{red}} {opt1}}}" if (is_teacher and lbl1 == q["correct"]) else f"\\textbf{{{lbl1}.}} {opt1[3:]}"
                    c2 = f"\\textbf{{\\color{{red}} {opt2}}}" if (is_teacher and lbl2 == q["correct"]) else f"\\textbf{{{lbl2}.}} {opt2[3:]}"
                    tex += f"{c1} & {c2} \\\\\n"
                tex += r"\end{tabularx}" + "\n"
            else:
                # 4 dòng riêng
                for opt in opts:
                    lbl = opt[:1]
                    is_c = (is_teacher and lbl == q["correct"])
                    if is_c:
                        tex += f"\\quad \\textbf{{\\color{{red}} {opt}}}\\\\\n"
                    else:
                        tex += f"\\quad \\textbf{{{lbl}.}} {opt[3:]}\\\\\n"

        if is_teacher:
            exp_text = clean_for_latex(q['explanation']).replace('\n', '\n\n')
            tex += r"""
\begin{quote}\small
{\color{myBlue}\textbf{► Lời giải chi tiết:}} Đáp án đúng: \textbf{\color{red} """ + q['correct'] + r"""}\\
""" + exp_text + r"""
\end{quote}
\vspace{2pt}
"""
        elif is_screen:
            tex += r"""
\vspace{1.8cm}
"""

    # -------------------------------------------------------------------------
    # PHẦN II
    # -------------------------------------------------------------------------
    tex += r"""
\vspace{6pt}
{\color{myGreen}\bfseries\large PHẦN II (2,0 ĐIỂM). Câu trắc nghiệm đúng sai.}

\textit{\small Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{3pt}
"""

    for q in PART_2_QUESTIONS:
        qid = q["id"]
        intro_tex = clean_for_latex(q['intro'])
        tex += f"\n\\textbf{{Câu {qid}.}} {intro_tex}\n\n"

        if q.get("image"):
            img_path = f"images_hoa10/{q['image']}"
            tex += r"""\begin{center}
\includegraphics[width=0.98\textwidth]{""" + img_path + r"""}
\end{center}
"""

        tex += f"\\textit{{{clean_for_latex(q['prompt'])}}}\n\n"
        for st in q["statements"]:
            sub_id = st["sub_id"]
            st_text = clean_for_latex(st["text"])
            correct_val = st["correct"]

            if is_teacher:
                color_tag = "red" if correct_val == "Đúng" else "mySub"
                tex += f"\\textbf{{{sub_id})}} [\\textbf{{\\color{{{color_tag}}} {correct_val.upper()}}}] {st_text}\n\n"
                exp_st = clean_for_latex(st['explanation'])
                tex += r"""\begin{quote}\small
{\color{myBlue}\textbf{► Giải thích ý """ + sub_id + r"""):}} """ + exp_st + r"""
\end{quote}
"""
            else:
                tex += f"\\textbf{{{sub_id})}} {st_text}\\\\\n"

        if is_screen:
            tex += r"\vspace{2.2cm}" + "\n"

    # -------------------------------------------------------------------------
    # PHẦN III
    # -------------------------------------------------------------------------
    tex += r"""
\vspace{6pt}
{\color{myGreen}\bfseries\large PHẦN III (2,0 ĐIỂM). Câu trắc nghiệm trả lời ngắn.}

\textit{\small Thí sinh trả lời từ câu 1 đến câu 8.}
\vspace{3pt}
"""

    for q in PART_3_QUESTIONS:
        qid = q["id"]
        content_tex = clean_for_latex(q['content'])
        tex += f"\n\\textbf{{Câu {qid}.}} {content_tex}\n\n"

        if is_teacher:
            exp_text = clean_for_latex(q['explanation']).replace('\n', '\n\n')
            tex += r"""
\begin{quote}\small
{\color{myBlue}\textbf{► Đáp án:}} \textbf{\color{red} """ + q['answer'] + r"""}\\
""" + exp_text + r"""
\end{quote}
\vspace{2pt}
"""
        elif is_screen:
            tex += r"\vspace{2.0cm}" + "\n"
        else:
            tex += r"\textit{Trả lời:} \dotfill\par\vspace{2pt}" + "\n"

    # -------------------------------------------------------------------------
    # PHẦN IV
    # -------------------------------------------------------------------------
    tex += r"""
\vspace{6pt}
{\color{myGreen}\bfseries\large PHẦN IV (3,0 ĐIỂM). Tự luận.}

\textit{\small Thí sinh làm từ câu 1 đến câu 3 ra giấy làm bài thi tự luận.}
\vspace{3pt}
"""

    for q in PART_4_QUESTIONS:
        qid = q["id"]
        pts = q["points"]
        content_tex = clean_for_latex(q['content'])
        tex += f"\n\\textbf{{Câu {qid} ({pts}).}} {content_tex}\n\n"

        for sub_q in q["sub_questions"]:
            sub_id = sub_q["sub_id"]
            sub_pts = sub_q["points"]
            sub_text = clean_for_latex(sub_q['text'])
            tex += f"\\textbf{{{sub_id})}} \\textit{{({sub_pts})}} {sub_text}\n\n"

            if is_teacher:
                sol_text = clean_for_latex(sub_q['solution']).replace('\n', '\n\n')
                tex += r"""
\begin{quote}\small
{\color{myBlue}\textbf{► Lời giải \& Biểu điểm (""" + sub_pts + r"""):}}\\
""" + sol_text + r"""
\end{quote}
\vspace{2pt}
"""

        if is_screen:
            tex += r"\vspace{3.5cm}" + "\n"
        elif not is_teacher:
            tex += r"""
\noindent\dotfill\\
\noindent\dotfill\\
\noindent\dotfill\\
\noindent\dotfill\par\vspace{4pt}
"""

    tex += r"""
\begin{center}
\textbf{--- HẾT ---}
\end{center}
"""

    # Bảng tổng hợp đáp án cho Giáo viên
    if is_teacher:
        tex += r"""
\vspace{8pt}
\begin{center}
{\color{red!80!black}\large\bfseries BẢNG TỔNG HỢP ĐÁP ÁN NHANH (MÃ ĐỀ 101)}
\end{center}

\textbf{\color{myGreen} 1. Phần I (Trắc nghiệm nhiều lựa chọn):}
\begin{center}
\begin{tabular}{|*{12}{c|}}
\hline
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} \\
\hline
\textbf{\color{red} D} & \textbf{\color{red} A} & \textbf{\color{red} B} & \textbf{\color{red} C} & \textbf{\color{red} C} & \textbf{\color{red} A} & \textbf{\color{red} B} & \textbf{\color{red} A} & \textbf{\color{red} B} & \textbf{\color{red} C} & \textbf{\color{red} B} & \textbf{\color{red} B} \\
\hline
\end{tabular}
\end{center}

\textbf{\color{myGreen} 2. Phần II (Trắc nghiệm đúng/sai):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Ý a} & \textbf{Ý b} & \textbf{Ý c} & \textbf{Ý d} \\
\hline
\textbf{Câu 1} & \textbf{\color{red} Đúng} & \textbf{\color{mySub} Sai} & \textbf{\color{red} Đúng} & \textbf{\color{red} Đúng} \\
\hline
\textbf{Câu 2} & \textbf{\color{mySub} Sai} & \textbf{\color{red} Đúng} & \textbf{\color{red} Đúng} & \textbf{\color{red} Đúng} \\
\hline
\end{tabular}
\end{center}

\textbf{\color{myGreen} 3. Phần III (Trả lời ngắn):}
\begin{center}
\begin{tabular}{|*{8}{c|}}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} & \textbf{Câu 7} & \textbf{Câu 8} \\
\hline
\textbf{\color{red} 95} & \textbf{\color{red} 8} & \textbf{\color{red} 2} & \textbf{\color{red} 40} & \textbf{\color{red} 1} & \textbf{\color{red} 26} & \textbf{\color{red} 1} & \textbf{\color{red} 15} \\
\hline
\end{tabular}
\end{center}
"""

    tex += r"\end{document}"
    return tex

def compile_pdf(tex_path):
    """Biên dịch file .tex thành .pdf bằng pdflatex."""
    cwd = os.path.dirname(tex_path)
    base_name = os.path.splitext(os.path.basename(tex_path))[0]
    cmd = ["pdflatex", "-interaction=nonstopmode", os.path.basename(tex_path)]
    print(f"Compiling {base_name}.tex...")
    res = subprocess.run(cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout_txt = res.stdout.decode('utf-8', errors='ignore') if res.stdout else ""
    if res.returncode == 0:
        print(f"Compilation SUCCESS: {base_name}.pdf")
    else:
        print(f"Compilation WARNING/ERROR: {base_name}.pdf (code {res.returncode})")
        lines = stdout_txt.split('\n')
        for l in lines:
            if l.startswith('!') or 'Error' in l:
                print("  ", l)

    # Chạy lần 2 để cập nhật số trang \thepage
    subprocess.run(cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # Dọn dẹp file rác
    for ext in ['.aux', '.log', '.out', '.thm']:
        junk = os.path.join(cwd, base_name + ext)
        if os.path.exists(junk):
            try:
                os.remove(junk)
            except Exception:
                pass

def main():
    configs = [
        ("De_Kiem_Tra_Hoa_10_Giua_HK1_DeThem01_HocSinh.tex", "hs"),
        ("De_Kiem_Tra_Hoa_10_Giua_HK1_DeThem01_GiaoVien.tex", "gv"),
        ("De_Kiem_Tra_Hoa_10_Giua_HK1_DeThem01_DayTrenManHinh.tex", "screen"),
    ]

    for filename, mode in configs:
        tex_path = os.path.join(OUT_DIR, filename)
        content = generate_tex(mode=mode)
        with open(tex_path, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"Generated TeX: {tex_path}")
        compile_pdf(tex_path)

        # Copy PDF sang thư mục OneDrive
        pdf_name = os.path.splitext(filename)[0] + ".pdf"
        pdf_path = os.path.join(OUT_DIR, pdf_name)
        if os.path.exists(pdf_path) and os.path.exists(ONEDRIVE_DIR):
            try:
                shutil.copy2(pdf_path, os.path.join(ONEDRIVE_DIR, pdf_name))
                print(f"Copied PDF to OneDrive: {pdf_name}")
            except Exception as e:
                print(f"Error copying PDF to OneDrive: {e}")

if __name__ == "__main__":
    main()
