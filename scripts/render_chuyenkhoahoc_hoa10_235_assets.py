# -*- coding: utf-8 -*-
"""
Script sinh hình ảnh cho Đề kiểm tra Giữa kì 1 Hóa 10 - THPT Chuyên Khoa Học Huế - Mã đề 235
Sử dụng TikZ + pdflatex + pymupdf để tạo ảnh PNG sắc nét (300 DPI)
"""

import os
import sys

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import subprocess
import pymupdf

IMG_DIR = r"C:\Antigravity_Thanh\San_Pham\images_chuyenkhoahoc_hoa10_235"
SCRATCH_DIR = r"C:\Antigravity_Thanh\scratch"

os.makedirs(IMG_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

def render_tikz_to_png(tikz_code, output_png_path, dpi=300):
    tex_content = r"""\documentclass[tikz,border=6pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{amsmath,amssymb}
\usepackage{tgheros}
\usepackage{sansmath}
\sansmath
\renewcommand{\familydefault}{\sfdefault}
\usetikzlibrary{arrows.meta,calc}

\begin{document}
""" + tikz_code + r"""
\end{document}
"""
    base_name = os.path.splitext(os.path.basename(output_png_path))[0]
    tex_path = os.path.join(SCRATCH_DIR, f"{base_name}.tex")
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_content)
    
    cmd = ["pdflatex", "-interaction=nonstopmode", f"-output-directory={SCRATCH_DIR}", tex_path]
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=SCRATCH_DIR)
    
    pdf_path = os.path.join(SCRATCH_DIR, f"{base_name}.pdf")
    if os.path.exists(pdf_path):
        doc = pymupdf.open(pdf_path)
        page = doc[0]
        pix = page.get_pixmap(dpi=dpi)
        pix.save(output_png_path)
        doc.close()
        print(f"[OK] Đã xuất ảnh: {output_png_path}")
    else:
        print(f"[LỖI] Không tạo được PDF cho {base_name}")

# 1. Câu 5: Bộ 4 đáp án ô orbital A, B, C, D
tikz_cau5 = r"""
\begin{tikzpicture}[line cap=round, line join=round, scale=0.75]
  \def\s{0.75}
  \tikzset{
    box/.style={line width=1.0pt, draw=black},
    spin_up/.style={line width=1.0pt, -{Stealth[length=3.5pt, width=2.5pt]}, draw=black},
    spin_down/.style={line width=1.0pt, -{Stealth[length=3.5pt, width=2.5pt]}, draw=black}
  }

  % Hàm vẽ 5 ô: 1s, 2s, 3 ô 2p
  % A: [↑↓] [↑↓] [↑↓][↑↓][  ]
  \begin{scope}[shift={(0, 2.8)}]
    \node[left, font=\bfseries] at (-0.2, 0.5*\s) {A.};
    % 1s
    \draw[box] (0, 0) rectangle (\s, \s);
    \draw[spin_up] (0.3*\s, 0.12*\s) -- (0.3*\s, 0.88*\s);
    \draw[spin_down] (0.7*\s, 0.88*\s) -- (0.7*\s, 0.12*\s);
    % 2s
    \draw[box] (1.3*\s, 0) rectangle (2.3*\s, \s);
    \draw[spin_up] (1.6*\s, 0.12*\s) -- (1.6*\s, 0.88*\s);
    \draw[spin_down] (2.0*\s, 0.88*\s) -- (2.0*\s, 0.12*\s);
    % 2p
    \draw[box] (2.6*\s, 0) rectangle (3.6*\s, \s);
    \draw[box] (3.6*\s, 0) rectangle (4.6*\s, \s);
    \draw[box] (4.6*\s, 0) rectangle (5.6*\s, \s);
    \draw[spin_up] (2.9*\s, 0.12*\s) -- (2.9*\s, 0.88*\s);
    \draw[spin_down] (3.3*\s, 0.88*\s) -- (3.3*\s, 0.12*\s);
    \draw[spin_up] (3.9*\s, 0.12*\s) -- (3.9*\s, 0.88*\s);
    \draw[spin_down] (4.3*\s, 0.88*\s) -- (4.3*\s, 0.12*\s);
  \end{scope}

  % B: [↑↓] [↑↓] [↑↓][↑ ][↓ ]
  \begin{scope}[shift={(7.5, 2.8)}]
    \node[left, font=\bfseries] at (-0.2, 0.5*\s) {B.};
    % 1s
    \draw[box] (0, 0) rectangle (\s, \s);
    \draw[spin_up] (0.3*\s, 0.12*\s) -- (0.3*\s, 0.88*\s);
    \draw[spin_down] (0.7*\s, 0.88*\s) -- (0.7*\s, 0.12*\s);
    % 2s
    \draw[box] (1.3*\s, 0) rectangle (2.3*\s, \s);
    \draw[spin_up] (1.6*\s, 0.12*\s) -- (1.6*\s, 0.88*\s);
    \draw[spin_down] (2.0*\s, 0.88*\s) -- (2.0*\s, 0.12*\s);
    % 2p
    \draw[box] (2.6*\s, 0) rectangle (3.6*\s, \s);
    \draw[box] (3.6*\s, 0) rectangle (4.6*\s, \s);
    \draw[box] (4.6*\s, 0) rectangle (5.6*\s, \s);
    \draw[spin_up] (2.9*\s, 0.12*\s) -- (2.9*\s, 0.88*\s);
    \draw[spin_down] (3.3*\s, 0.88*\s) -- (3.3*\s, 0.12*\s);
    \draw[spin_up] (4.1*\s, 0.12*\s) -- (4.1*\s, 0.88*\s);
    \draw[spin_down] (5.1*\s, 0.88*\s) -- (5.1*\s, 0.12*\s);
  \end{scope}

  % C: [↑↓] [↑↑] [↑↓][↑ ][↑ ]
  \begin{scope}[shift={(0, 0)}]
    \node[left, font=\bfseries] at (-0.2, 0.5*\s) {C.};
    % 1s
    \draw[box] (0, 0) rectangle (\s, \s);
    \draw[spin_up] (0.3*\s, 0.12*\s) -- (0.3*\s, 0.88*\s);
    \draw[spin_down] (0.7*\s, 0.88*\s) -- (0.7*\s, 0.12*\s);
    % 2s (2 mũi tên cùng lên)
    \draw[box] (1.3*\s, 0) rectangle (2.3*\s, \s);
    \draw[spin_up] (1.6*\s, 0.12*\s) -- (1.6*\s, 0.88*\s);
    \draw[spin_up] (2.0*\s, 0.12*\s) -- (2.0*\s, 0.88*\s);
    % 2p
    \draw[box] (2.6*\s, 0) rectangle (3.6*\s, \s);
    \draw[box] (3.6*\s, 0) rectangle (4.6*\s, \s);
    \draw[box] (4.6*\s, 0) rectangle (5.6*\s, \s);
    \draw[spin_up] (2.9*\s, 0.12*\s) -- (2.9*\s, 0.88*\s);
    \draw[spin_down] (3.3*\s, 0.88*\s) -- (3.3*\s, 0.12*\s);
    \draw[spin_up] (4.1*\s, 0.12*\s) -- (4.1*\s, 0.88*\s);
    \draw[spin_up] (5.1*\s, 0.12*\s) -- (5.1*\s, 0.88*\s);
  \end{scope}

  % D: [↑↓] [↑↓] [↑↓][↑ ][↑ ]
  \begin{scope}[shift={(7.5, 0)}]
    \node[left, font=\bfseries] at (-0.2, 0.5*\s) {D.};
    % 1s
    \draw[box] (0, 0) rectangle (\s, \s);
    \draw[spin_up] (0.3*\s, 0.12*\s) -- (0.3*\s, 0.88*\s);
    \draw[spin_down] (0.7*\s, 0.88*\s) -- (0.7*\s, 0.12*\s);
    % 2s
    \draw[box] (1.3*\s, 0) rectangle (2.3*\s, \s);
    \draw[spin_up] (1.6*\s, 0.12*\s) -- (1.6*\s, 0.88*\s);
    \draw[spin_down] (2.0*\s, 0.88*\s) -- (2.0*\s, 0.12*\s);
    % 2p
    \draw[box] (2.6*\s, 0) rectangle (3.6*\s, \s);
    \draw[box] (3.6*\s, 0) rectangle (4.6*\s, \s);
    \draw[box] (4.6*\s, 0) rectangle (5.6*\s, \s);
    \draw[spin_up] (2.9*\s, 0.12*\s) -- (2.9*\s, 0.88*\s);
    \draw[spin_down] (3.3*\s, 0.88*\s) -- (3.3*\s, 0.12*\s);
    \draw[spin_up] (4.1*\s, 0.12*\s) -- (4.1*\s, 0.88*\s);
    \draw[spin_up] (5.1*\s, 0.12*\s) -- (5.1*\s, 0.88*\s);
  \end{scope}
\end{tikzpicture}
"""

# 2. Câu 8: Mô phỏng 3 nguyên tử
tikz_cau8 = r"""
\begin{tikzpicture}[line cap=round, line join=round, scale=1.0]
  % 1. Nguyên tử 1 (8p, 8n)
  \begin{scope}[shift={(0,0)}]
    % Hạt nhân
    \draw[line width=1.0pt, fill=black!75] (0,0) circle (0.55cm);
    \node[white, font=\bfseries\footnotesize, align=center] at (0,0) {8p\\8n};
    % 2 lớp electron
    \draw[line width=0.8pt] (0,0) circle (1.05cm);
    \draw[line width=0.8pt] (0,0) circle (1.55cm);
    % Lớp 1 (2e)
    \fill[black] (0, 1.05) circle (2.2pt);
    \fill[black] (0, -1.05) circle (2.2pt);
    % Lớp 2 (6e)
    \foreach \ang in {0, 60, 120, 180, 240, 300} {
      \fill[black] (\ang:1.55) circle (2.2pt);
    }
    \node[below=1.75cm, font=\bfseries] at (0,0) {1};
  \end{scope}

  % 2. Nguyên tử 2 (8p, 9n)
  \begin{scope}[shift={(4.2,0)}]
    % Hạt nhân
    \draw[line width=1.0pt, fill=black!75] (0,0) circle (0.55cm);
    \node[white, font=\bfseries\footnotesize, align=center] at (0,0) {8p\\9n};
    % 2 lớp electron
    \draw[line width=0.8pt] (0,0) circle (1.05cm);
    \draw[line width=0.8pt] (0,0) circle (1.55cm);
    % Lớp 1 (2e)
    \fill[black] (0, 1.05) circle (2.2pt);
    \fill[black] (0, -1.05) circle (2.2pt);
    % Lớp 2 (6e)
    \foreach \ang in {0, 60, 120, 180, 240, 300} {
      \fill[black] (\ang:1.55) circle (2.2pt);
    }
    \node[below=1.75cm, font=\bfseries] at (0,0) {2};
  \end{scope}

  % 3. Nguyên tử 3 (9p, 10n)
  \begin{scope}[shift={(8.4,0)}]
    % Hạt nhân
    \draw[line width=1.0pt, fill=black!75] (0,0) circle (0.55cm);
    \node[white, font=\bfseries\footnotesize, align=center] at (0,0) {9p\\10n};
    % 2 lớp electron
    \draw[line width=0.8pt] (0,0) circle (1.05cm);
    \draw[line width=0.8pt] (0,0) circle (1.55cm);
    % Lớp 1 (2e)
    \fill[black] (0, 1.05) circle (2.2pt);
    \fill[black] (0, -1.05) circle (2.2pt);
    % Lớp 2 (7e)
    \foreach \ang in {0, 51.4, 102.8, 154.2, 205.7, 257.1, 308.5} {
      \fill[black] (\ang:1.55) circle (2.2pt);
    }
    \node[below=1.75cm, font=\bfseries] at (0,0) {3};
  \end{scope}
\end{tikzpicture}
"""

# 3. Câu 12: Bộ 4 đáp án hình dạng orbital A (s), B (px), C (py), D (pz)
tikz_cau12 = r"""
\begin{tikzpicture}[line cap=round, line join=round, scale=0.8]
  % A: Orbital s (hình cầu)
  \begin{scope}[shift={(0, 3.2)}]
    \node[left, font=\bfseries] at (-1.5, 0) {A.};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (-1.2,0) -- (1.2,0) node[right, font=\tiny] {$x$};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (0,-1.2) -- (0,1.2) node[above, font=\tiny] {$z$};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (-0.7,-0.5) -- (0.7,0.5) node[above right, font=\tiny] {$y$};
    \fill[black!70] (0,0) circle (0.45cm);
    \node[below=1.3cm, font=\footnotesize] at (0,0) {s};
  \end{scope}

  % B: Orbital px
  \begin{scope}[shift={(6.0, 3.2)}]
    \node[left, font=\bfseries] at (-1.5, 0) {B.};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (-1.2,0) -- (1.2,0) node[right, font=\tiny] {$x$};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (0,-1.2) -- (0,1.2) node[above, font=\tiny] {$z$};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (-0.7,-0.5) -- (0.7,0.5) node[above right, font=\tiny] {$y$};
    \fill[black!70] (0.45, 0) ellipse (0.4cm and 0.18cm);
    \fill[black!70] (-0.45, 0) ellipse (0.4cm and 0.18cm);
    \node[below=1.3cm, font=\footnotesize] at (0,0) {$p_x$};
  \end{scope}

  % C: Orbital py
  \begin{scope}[shift={(0, 0)}]
    \node[left, font=\bfseries] at (-1.5, 0) {C.};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (-1.2,0) -- (1.2,0) node[right, font=\tiny] {$x$};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (0,-1.2) -- (0,1.2) node[above, font=\tiny] {$z$};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (-0.8,-0.6) -- (0.8,0.6) node[above right, font=\tiny] {$y$};
    \fill[black!70, rotate=37] (0.45, 0) ellipse (0.4cm and 0.18cm);
    \fill[black!70, rotate=37] (-0.45, 0) ellipse (0.4cm and 0.18cm);
    \node[below=1.3cm, font=\footnotesize] at (0,0) {$p_y$};
  \end{scope}

  % D: Orbital pz
  \begin{scope}[shift={(6.0, 0)}]
    \node[left, font=\bfseries] at (-1.5, 0) {D.};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (-1.2,0) -- (1.2,0) node[right, font=\tiny] {$x$};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (0,-1.2) -- (0,1.2) node[above, font=\tiny] {$z$};
    \draw[line width=0.8pt, -{Stealth[scale=0.8]}] (-0.7,-0.5) -- (0.7,0.5) node[above right, font=\tiny] {$y$};
    \fill[black!70, rotate=90] (0.45, 0) ellipse (0.4cm and 0.18cm);
    \fill[black!70, rotate=90] (-0.45, 0) ellipse (0.4cm and 0.18cm);
    \node[below=1.3cm, font=\footnotesize] at (0,0) {$p_z$};
  \end{scope}
\end{tikzpicture}
"""

def main():
    print("Đang tạo các hình vẽ TikZ độ nét cao cho THPT Chuyên Khoa Học Huế (Mã đề 235)...")
    render_tikz_to_png(tikz_cau5, os.path.join(IMG_DIR, "hinh_cau5_orbital_abcd.png"))
    render_tikz_to_png(tikz_cau8, os.path.join(IMG_DIR, "hinh_cau8_mo_phong_nguyen_tu.png"))
    render_tikz_to_png(tikz_cau12, os.path.join(IMG_DIR, "hinh_cau12_orbital_abcd.png"))
    print("Hoàn tất xuất tất cả hình vẽ sang PNG 300 DPI!")

if __name__ == "__main__":
    main()
