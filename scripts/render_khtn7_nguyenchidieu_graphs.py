# -*- coding: utf-8 -*-
"""
Vẽ đồ thị Câu 3 và Câu 4 trong Đề cương KHTN 7 THCS Nguyễn Chí Diểu bằng TikZ độ nét cao.
"""

import sys
if sys.stdout:
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr:
    sys.stderr.reconfigure(encoding='utf-8')

import os
import subprocess
import shutil

SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
IMG_DIR = r"c:\Antigravity_Thanh\San_Pham\images_khtn7_nguyenchidieu"
os.makedirs(SCRATCH_DIR, exist_ok=True)
os.makedirs(IMG_DIR, exist_ok=True)

# ------------------------------------------------------------------------------
# 1. Đồ thị Câu 4 (4 giai đoạn O-A-B-C-D)
# ------------------------------------------------------------------------------
tikz_cau4 = r"""\documentclass[tikz,border=5pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{tikz}
\usetikzlibrary{arrows.meta}

\begin{document}
\begin{tikzpicture}[scale=0.9, >=Stealth]
  % Grid
  \draw[step=1, gray!30, very thin] (0,0) grid (8.5,10.5);

  % Axes
  \draw[->, thick] (0,0) -- (8.8,0) node[below right] {$t\text{ (s)}$};
  \draw[->, thick] (0,0) -- (0,10.8) node[above left] {$s\text{ (m)}$};

  % Ticks
  \node[below left] at (0,0) {$O$};
  \foreach \x in {2,4,6,8} {
    \draw (\x,0.1) -- (\x,-0.1) node[below] {\small $\x$};
  }
  \foreach \y in {2,4,6,8,10} {
    \draw (0.1,\y) -- (-0.1,\y) node[left] {\small $\y$};
  }

  % Dashed projection lines
  \draw[dashed, gray] (2,0) -- (2,4) -- (0,4);
  \draw[dashed, gray] (4,0) -- (4,4);
  \draw[dashed, gray] (6,0) -- (6,8) -- (0,8);
  \draw[dashed, gray] (8,0) -- (8,10) -- (0,10);

  % Graph line
  \draw[line width=1.5pt, blue!80!black] (0,0) -- (2,4) -- (4,4) -- (6,8) -- (8,10);

  % Dots
  \filldraw[blue!80!black] (0,0) circle (2pt);
  \filldraw[blue!80!black] (2,4) circle (2pt);
  \filldraw[blue!80!black] (4,4) circle (2pt);
  \filldraw[blue!80!black] (6,8) circle (2pt);
  \filldraw[blue!80!black] (8,10) circle (2pt);

  % Labels A, B, C, D
  \node[above left, red!80!black] at (1,2.2) {\textbf{A}};
  \node[above, red!80!black] at (3,4.1) {\textbf{B}};
  \node[above left, red!80!black] at (5,6.2) {\textbf{C}};
  \node[above left, red!80!black] at (7,9.2) {\textbf{D}};
\end{tikzpicture}
\end{document}
"""

# ------------------------------------------------------------------------------
# 2. Đồ thị Câu 3 (Vẽ đồ thị từ bảng số liệu s - t)
# ------------------------------------------------------------------------------
tikz_cau3 = r"""\documentclass[tikz,border=5pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{tikz}
\usetikzlibrary{arrows.meta}

\begin{document}
\begin{tikzpicture}[scale=1.1, >=Stealth]
  % Grid
  \draw[step=1, gray!30, very thin] (0,0) grid (4.5,4.5);

  % Axes
  \draw[->, thick] (0,0) -- (4.8,0) node[below right] {$t\text{ (h)}$};
  \draw[->, thick] (0,0) -- (0,4.8) node[above left] {$s\text{ (km)}$};

  % Ticks
  \node[below left] at (0,0) {$0$};
  \foreach \x in {1,2,3,4} {
    \draw (\x,0.1) -- (\x,-0.1) node[below] {\small $\x$};
  }
  \foreach \y/\val in {1/60, 2/120, 3/180, 4/240} {
    \draw (0.1,\y) -- (-0.1,\y) node[left] {\small $\val$};
  }

  % Dashed projections
  \foreach \i in {1,2,3,4} {
    \draw[dashed, gray] (\i,0) -- (\i,\i) -- (0,\i);
    \filldraw[red!80!black] (\i,\i) circle (2pt);
  }

  % Line
  \draw[line width=1.5pt, red!80!black] (0,0) -- (4.2,4.2);
  \node[above left, red!80!black] at (4.2,4.2) {\small $s = 60t$};
\end{tikzpicture}
\end{document}
"""

def compile_tikz(tex_content, name):
    tex_file = os.path.join(SCRATCH_DIR, f"{name}.tex")
    with open(tex_file, "w", encoding="utf-8") as f:
        f.write(tex_content)
    subprocess.run(["pdflatex", "-interaction=nonstopmode", "-output-directory", SCRATCH_DIR, tex_file],
                   stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    pdf_file = os.path.join(SCRATCH_DIR, f"{name}.pdf")
    png_file = os.path.join(IMG_DIR, f"{name}.png")
    
    # Dùng pdftoppm để xuất PNG 600 DPI
    subprocess.run(["pdftoppm", "-png", "-r", "600", "-singlefile", pdf_file, os.path.join(IMG_DIR, name)],
                   stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    print(f"[OK] Đã tạo ảnh: {png_file}")

if __name__ == "__main__":
    compile_tikz(tikz_cau4, "dothi_cau4_khtn7_nguyenchidieu")
    compile_tikz(tikz_cau3, "dothi_cau3_khtn7_nguyenchidieu")
