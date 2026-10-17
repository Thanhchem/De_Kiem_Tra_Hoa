# -*- coding: utf-8 -*-
"""
Script tạo tài liệu LaTeX (.tex) và biên dịch ra PDF cho:
Đề thi thử Giữa học kì I Hóa học 11 (Mã đề 104 / Đề số 04)
Năm học 2026 - 2027 - Thầy TRẦN VĂN THẠNH

Sinh 3 file:
1. De_Thi_Thu_Hoa_11_Giua_HK1_De04_HocSinh.tex & .pdf
2. De_Thi_Thu_Hoa_11_Giua_HK1_De04_GiaoVien.tex & .pdf
3. De_Thi_Thu_Hoa_11_Giua_HK1_De04_DayTrenManHinh.tex & .pdf
"""

import os
import sys
import re
import subprocess
import shutil

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

ROOT_DIR = r"c:\Antigravity_Thanh"
OUT_DIR = os.path.join(ROOT_DIR, "San_Pham")
IMG_DIR = os.path.join(OUT_DIR, "images_hoa11_de04")
ONEDRIVE_DIR = r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\11"
os.makedirs(OUT_DIR, exist_ok=True)

sys.path.append(os.path.join(ROOT_DIR, "scripts"))
from hoa11_de04_data import (
    EXAM_INFO, PART_1_QUESTIONS, PART_2_QUESTIONS, PART_3_QUESTIONS
)

def clean_for_latex(text):
    """Xử lý triệt để ký tự đặc biệt, công thức hóa học cho LaTeX."""
    if not text:
        return ""
    t = text
    
    # 1. Ký tự Unicode đặc biệt
    t = t.replace('►', r'$\blacktriangleright$')
    t = t.replace('⇒', r'$\Rightarrow$')
    t = t.replace('⇌', r'$\rightleftharpoons$')
    t = t.replace('→', r'$\longrightarrow$')
    t = t.replace('≡', r'$\equiv$')
    t = t.replace('Δ', r'$\Delta$')
    t = t.replace('°', r'$^\circ$')
    t = t.replace('✎', r'$\triangleright$')
    t = t.replace('σ', r'$\sigma$')
    t = t.replace('π', r'$\pi$')
    
    # 2. Placeholders cho nhiệt độ để tránh va chạm regex
    t = t.replace('^o^C', '___DEGC___')
    t = t.replace('^o^', '___DEG___')
    
    # 3. Xử lý Subscript & Superscript liền nhau (ví dụ: NO~3~^-^ -> NO$_{3}^{-}$)
    t = re.sub(r'~([a-zA-Z0-9+\-_]+)~\^([a-zA-Z0-9+\-_]+)\^', r'$_{\1}^{\2}$', t)
    t = re.sub(r'\^([a-zA-Z0-9+\-_]+)\^~([a-zA-Z0-9+\-_]+)~', r'$^{\1}_{\2}$', t)
    
    # 4. Xử lý Subscript ~...~ và Superscript ^...^ độc lập
    t = re.sub(r'~([a-zA-Z0-9+\-_]+)~', r'$_{\1}$', t)
    t = re.sub(r'\^([a-zA-Z0-9+\-_]+)\^', r'$^{\1}$', t)
    
    # 5. Phục hồi ký hiệu độ
    t = t.replace('___DEGC___', r'$^\circ\text{C}$')
    t = t.replace('___DEG___', r'$^\circ$')
    
    # 6. Xử lý in đậm markdown **...** và in nghiêng *...*
    t = re.sub(r'\*\*([^\*\n]+?)\*\*', r'\\textbf{\1}', t)
    t = re.sub(r'\*([^\*\n]+?)\*', r'\\textit{\1}', t)
    
    # 7. Xử lý ký hiệu % nếu chưa có \ trước
    t = re.sub(r'(?<!\\)%', r'\%', t)
    
    return t

def generate_tex(mode="hs"):
    is_teacher = (mode == "gv")
    is_screen = (mode == "screen")

    font_size = "13pt" if is_screen else "11pt"
    line_stretch = "1.25" if is_screen else "1.06"

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
\textbf{\small SỞ GIÁO DỤC VÀ ĐÀO TẠO} & & \textbf{\small ĐỀ THI THỬ GIỮA HỌC KÌ I} \\
\textbf{\normalsize TRƯỜNG THPT CHUYÊN / THPT} & & \textbf{\normalsize MÔN: HÓA HỌC - LỚP 11} \\
\textit{\small (Đề thi gồm có 04 trang)} & & \textit{\small Thời gian làm bài: 45 phút (không kể phát đề)}
\end{tabularx}

\begin{center}
__TITLE_SUB__
\textbf{Họ và tên thí sinh:} \dotfill \quad \textbf{Lớp:} \dots\dots\dots \quad \textbf{Số báo danh:} \dots\dots\dots \hfill \textbf{\large Mã đề 104}
\end{center}
\vspace{-4pt}
{\color{gray!50}\hrule height 0.8pt}
\vspace{6pt}
"""
    tex = header_template.replace("__FONT_SIZE__", font_size).replace("__LINE_STRETCH__", line_stretch).replace("__TITLE_SUB__", title_sub)

    # -------------------------------------------------------------------------
    # PHẦN I: 18 CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN
    # -------------------------------------------------------------------------
    tex += r"""
{\color{myGreen}\bfseries\large PHẦN I (4,5 ĐIỂM). Câu trắc nghiệm nhiều phương án lựa chọn.}

\textit{\small Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{3pt}
"""

    for q in PART_1_QUESTIONS:
        qid = q["id"]
        content_tex = clean_for_latex(q['content'])
        tex += f"\n\\textbf{{Câu {qid}.}} {content_tex}\n\n"

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
{\color{myBlue}\textbf{$\blacktriangleright$ Lời giải chi tiết:}} Đáp án đúng: \textbf{\color{red} """ + q['correct'] + r"""}\\
""" + exp_text + r"""
\end{quote}
\vspace{2pt}
"""
        elif is_screen:
            tex += r"""
\vspace{1.6cm}
"""

    # -------------------------------------------------------------------------
    # PHẦN II: 4 CÂU ĐÚNG / SAI
    # -------------------------------------------------------------------------
    tex += r"""
\vspace{6pt}
{\color{myGreen}\bfseries\large PHẦN II (4,0 ĐIỂM). Câu trắc nghiệm đúng sai.}

\textit{\small Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{3pt}
"""

    for q in PART_2_QUESTIONS:
        qid = q["id"]
        content_tex = clean_for_latex(q['content'])
        tex += f"\n\\textbf{{Câu {qid}.}} {content_tex}\n\n"

        if q.get("image"):
            img_path = f"images_hoa11_de04/{q['image']}"
            tex += r"""\begin{center}
\includegraphics[width=0.42\textwidth]{""" + img_path + r"""}
\end{center}
"""

        for st in q["statements"]:
            sub_id = st["id"]
            st_text = clean_for_latex(st["text"])
            correct_val = "Đúng" if st["correct"] else "Sai"

            if is_teacher:
                color_tag = "red" if st["correct"] else "mySub"
                tex += f"\\textbf{{{sub_id})}} [\\textbf{{\\color{{{color_tag}}} {correct_val.upper()}}}] {st_text}\\\\\n"
            else:
                tex += f"\\textbf{{{sub_id})}} {st_text}\\\\\n"

        if is_teacher:
            exp_text = clean_for_latex(q['explanation']).replace('\n', '\n\n')
            tex += r"""
\begin{quote}\small
{\color{myBlue}\textbf{$\blacktriangleright$ Giải thích chi tiết:}}\\
""" + exp_text + r"""
\end{quote}
\vspace{2pt}
"""
        elif is_screen:
            tex += r"\vspace{2.0cm}" + "\n"

    # -------------------------------------------------------------------------
    # PHẦN III: 6 CÂU TRẢ LỜI NGẮN
    # -------------------------------------------------------------------------
    tex += r"""
\vspace{6pt}
{\color{myGreen}\bfseries\large PHẦN III (1,5 ĐIỂM). Câu trắc nghiệm trả lời ngắn.}

\textit{\small Thí sinh trả lời từ câu 1 đến câu 6.}
\vspace{3pt}
"""

    for q in PART_3_QUESTIONS:
        qid = q["id"]
        content_tex = clean_for_latex(q['content']).replace('\n', '\n\n')
        tex += f"\n\\textbf{{Câu {qid}.}} {content_tex}\n\n"

        if is_teacher:
            exp_text = clean_for_latex(q['explanation']).replace('\n', '\n\n')
            tex += r"""
\begin{quote}\small
{\color{myBlue}\textbf{$\blacktriangleright$ Đáp án:}} \textbf{\color{red} """ + q['correct'] + r"""}\\
""" + exp_text + r"""
\end{quote}
\vspace{2pt}
"""
        elif is_screen:
            tex += r"\vspace{2.0cm}" + "\n"
        else:
            tex += r"\textit{Trả lời:} \dotfill\par\vspace{2pt}" + "\n"

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
{\color{red!80!black}\large\bfseries BẢNG TỔNG HỢP ĐÁP ÁN NHANH (MÃ ĐỀ 104)}
\end{center}

\textbf{\color{myGreen} 1. Phần I (Trắc nghiệm nhiều lựa chọn - 18 câu):}
\begin{center}
\begin{tabular}{|*{9}{c|}}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} & \textbf{Câu 7} & \textbf{Câu 8} & \textbf{Câu 9} \\
\hline
\textbf{\color{red} A} & \textbf{\color{red} C} & \textbf{\color{red} B} & \textbf{\color{red} A} & \textbf{\color{red} B} & \textbf{\color{red} D} & \textbf{\color{red} B} & \textbf{\color{red} B} & \textbf{\color{red} B} \\
\hline
\textbf{Câu 10} & \textbf{Câu 11} & \textbf{Câu 12} & \textbf{Câu 13} & \textbf{Câu 14} & \textbf{Câu 15} & \textbf{Câu 16} & \textbf{Câu 17} & \textbf{Câu 18} \\
\hline
\textbf{\color{red} C} & \textbf{\color{red} B} & \textbf{\color{red} C} & \textbf{\color{red} B} & \textbf{\color{red} B} & \textbf{\color{red} C} & \textbf{\color{red} A} & \textbf{\color{red} C} & \textbf{\color{red} B} \\
\hline
\end{tabular}
\end{center}

\textbf{\color{myGreen} 2. Phần II (Trắc nghiệm đúng/sai - 4 câu):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu hỏi} & \textbf{Lệnh hỏi a} & \textbf{Lệnh hỏi b} & \textbf{Lệnh hỏi c} & \textbf{Lệnh hỏi d} \\
\hline
\textbf{Câu 1} & \textbf{\color{red} Đúng} & \textbf{\color{red} Đúng} & \textbf{\color{red} Đúng} & \textbf{\color{red} Đúng} \\
\hline
\textbf{Câu 2} & \textbf{\color{mySub} Sai} & \textbf{\color{red} Đúng} & \textbf{\color{red} Đúng} & \textbf{\color{mySub} Sai} \\
\hline
\textbf{Câu 3} & \textbf{\color{red} Đúng} & \textbf{\color{mySub} Sai} & \textbf{\color{red} Đúng} & \textbf{\color{red} Đúng} \\
\hline
\textbf{Câu 4} & \textbf{\color{red} Đúng} & \textbf{\color{mySub} Sai} & \textbf{\color{red} Đúng} & \textbf{\color{red} Đúng} \\
\hline
\end{tabular}
\end{center}

\textbf{\color{myGreen} 3. Phần III (Trả lời ngắn - 6 câu):}
\begin{center}
\begin{tabular}{|*{6}{c|}}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\
\hline
\textbf{\color{red} 3} & \textbf{\color{red} 2} & \textbf{\color{red} 2} & \textbf{\color{red} 7} & \textbf{\color{red} 5} & \textbf{\color{red} 4} \\
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
        for l in stdout_txt.split('\n'):
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
        ("De_Thi_Thu_Hoa_11_Giua_HK1_De04_HocSinh.tex", "hs"),
        ("De_Thi_Thu_Hoa_11_Giua_HK1_De04_GiaoVien.tex", "gv"),
        ("De_Thi_Thu_Hoa_11_Giua_HK1_De04_DayTrenManHinh.tex", "screen"),
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
