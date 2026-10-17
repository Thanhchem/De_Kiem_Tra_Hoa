# -*- coding: utf-8 -*-
import os
import subprocess
import pymupdf

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham\images_khtn7"
os.makedirs(OUT_DIR, exist_ok=True)

# 1. do_thi_cau5.tex
tex_cau5 = r"""\documentclass[border=8pt]{standalone}
\usepackage[utf8]{inputenc}
\usepackage{times}
\usepackage{tikz}
\usetikzlibrary{arrows.meta}

\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, >=Latex, font=\large]
  % Grid
  \draw[dashed, line width=0.75pt, black!80] (0,0) grid[xstep=1.2, ystep=1.0] (4.8,4.0);
  % Axes
  \draw[->, line width=1.3pt] (0,0) -- (5.6,0) node[below=2pt] {$t\mbox{ (s)}$};
  \draw[->, line width=1.3pt] (0,0) -- (0,4.8) node[left=2pt] {$s\mbox{ (m)}$};
  % Ticks & Labels
  \node[below left=2pt] at (0,0) {\textbf{O}};
  \node[below=3pt] at (1.2,0) {2};
  \node[below=3pt] at (2.4,0) {4};
  \node[below=3pt] at (3.6,0) {6};
  \node[below=3pt] at (4.8,0) {8};
  \node[left=3pt] at (0,1.0) {5};
  \node[left=3pt] at (0,2.0) {10};
  \node[left=3pt] at (0,3.0) {15};
  \node[left=3pt] at (0,4.0) {20};
  % Line
  \draw[line width=2.0pt, black] (0,0) -- (4.8,4.0);
  \fill[black] (0,0) circle (2.2pt);
  \fill[black] (1.2,1.0) circle (2.2pt);
  \fill[black] (2.4,2.0) circle (2.2pt);
  \fill[black] (3.6,3.0) circle (2.2pt);
  \fill[black] (4.8,4.0) circle (2.2pt);
\end{tikzpicture}
\end{document}
"""

# 2. do_thi_cau6.tex
tex_cau6 = r"""\documentclass[border=8pt]{standalone}
\usepackage[utf8]{inputenc}
\usepackage{times}
\usepackage{tikz}
\usetikzlibrary{arrows.meta}

\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, >=Latex, font=\large]
  % Grid
  \draw[dashed, line width=0.65pt, gray!70] (0,0) grid[xstep=1.1, ystep=0.7] (5.5,4.9);
  % Axes
  \draw[->, line width=1.3pt] (0,0) -- (6.3,0) node[below=2pt] {$t\mbox{ (s)}$};
  \draw[->, line width=1.3pt] (0,0) -- (0,5.6) node[left=2pt] {$s\mbox{ (m)}$};
  % Labels
  \node[below left=2pt] at (0,0) {\textbf{O}};
  \node[below=3pt] at (1.1,0) {10};
  \node[below=3pt] at (2.2,0) {20};
  \node[below=3pt] at (3.3,0) {30};
  \node[below=3pt] at (4.4,0) {40};
  \node[below=3pt] at (5.5,0) {50};
  
  \node[left=3pt] at (0,0.7) {14};
  \node[left=3pt] at (0,1.4) {28};
  \node[left=3pt] at (0,2.1) {42};
  \node[left=3pt] at (0,2.8) {56};
  \node[left=3pt] at (0,3.5) {70};
  
  % Points & Line
  \draw[line width=2.0pt, blue!80!black] (0,0) -- (5.5,3.5);
  \fill[red!80!black] (0,0) circle (2.2pt) node[below right=2pt, font=\small, black] {$O(0;0)$};
  \fill[red!80!black] (1.1,0.7) circle (2.2pt) node[below right=1pt, font=\small, black] {$A(10;14)$};
  \fill[red!80!black] (2.2,1.4) circle (2.2pt) node[below right=1pt, font=\small, black] {$B(20;28)$};
  \fill[red!80!black] (3.3,2.1) circle (2.2pt) node[below right=1pt, font=\small, black] {$C(30;42)$};
  \fill[red!80!black] (4.4,2.8) circle (2.2pt) node[below right=1pt, font=\small, black] {$D(40;56)$};
  \fill[red!80!black] (5.5,3.5) circle (2.2pt) node[below right=1pt, font=\small, black] {$E(50;70)$};
\end{tikzpicture}
\end{document}
"""

# 3. mo_hinh_carbon.tex
tex_carbon = r"""\documentclass[border=8pt]{standalone}
\usepackage[utf8]{inputenc}
\usepackage{times}
\usepackage{tikz}

\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, font=\large]
  % Concentric shells
  \draw[dash pattern=on 3pt off 2pt, line width=1.0pt, blue!70!black] (0,0) circle (1.3);
  \draw[dash pattern=on 3pt off 2pt, line width=1.0pt, blue!70!black] (0,0) circle (2.3);
  
  % Nucleus
  \fill[red!85!black] (0,0) circle (0.65);
  \draw[line width=1.2pt, red!40!black] (0,0) circle (0.65);
  \node[white, font=\bfseries] at (0,0.12) {$+6$};
  \node[white, font=\small] at (0,-0.22) {6p, 6n};
  
  % Shell 1: 2 electrons
  \fill[blue!90!black] (0, 1.3) circle (0.13);
  \draw[line width=0.6pt, black] (0, 1.3) circle (0.13);
  \node[white, font=\footnotesize\bfseries] at (0, 1.3) {$-$};
  
  \fill[blue!90!black] (0, -1.3) circle (0.13);
  \draw[line width=0.6pt, black] (0, -1.3) circle (0.13);
  \node[white, font=\footnotesize\bfseries] at (0, -1.3) {$-$};
  
  % Shell 2: 4 electrons
  \fill[blue!90!black] (2.3, 0) circle (0.13);
  \draw[line width=0.6pt, black] (2.3, 0) circle (0.13);
  \node[white, font=\footnotesize\bfseries] at (2.3, 0) {$-$};
  
  \fill[blue!90!black] (0, 2.3) circle (0.13);
  \draw[line width=0.6pt, black] (0, 2.3) circle (0.13);
  \node[white, font=\footnotesize\bfseries] at (0, 2.3) {$-$};
  
  \fill[blue!90!black] (-2.3, 0) circle (0.13);
  \draw[line width=0.6pt, black] (-2.3, 0) circle (0.13);
  \node[white, font=\footnotesize\bfseries] at (-2.3, 0) {$-$};
  
  \fill[blue!90!black] (0, -2.3) circle (0.13);
  \draw[line width=0.6pt, black] (0, -2.3) circle (0.13);
  \node[white, font=\footnotesize\bfseries] at (0, -2.3) {$-$};
  
  % Caption
  \node[below=8pt, font=\bfseries] at (0,-2.5) {Mo hinh nguyen tu Carbon (C)};
\end{tikzpicture}
\end{document}
"""

# 4. mo_hinh_aluminium.tex
tex_al = r"""\documentclass[border=8pt]{standalone}
\usepackage[utf8]{inputenc}
\usepackage{times}
\usepackage{tikz}

\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, font=\large]
  % Concentric shells
  \draw[dash pattern=on 3pt off 2pt, line width=1.0pt, blue!70!black] (0,0) circle (1.2);
  \draw[dash pattern=on 3pt off 2pt, line width=1.0pt, blue!70!black] (0,0) circle (2.0);
  \draw[dash pattern=on 3pt off 2pt, line width=1.0pt, blue!70!black] (0,0) circle (2.8);
  
  % Nucleus
  \fill[red!85!black] (0,0) circle (0.7);
  \draw[line width=1.2pt, red!40!black] (0,0) circle (0.7);
  \node[white, font=\bfseries] at (0,0.13) {$+13$};
  \node[white, font=\small] at (0,-0.23) {13p, 14n};
  
  % Shell 1: 2 electrons
  \fill[blue!90!black] (0, 1.2) circle (0.12);
  \draw[line width=0.6pt, black] (0, 1.2) circle (0.12);
  \node[white, font=\footnotesize\bfseries] at (0, 1.2) {$-$};
  
  \fill[blue!90!black] (0, -1.2) circle (0.12);
  \draw[line width=0.6pt, black] (0, -1.2) circle (0.12);
  \node[white, font=\footnotesize\bfseries] at (0, -1.2) {$-$};
  
  % Shell 2: 8 electrons
  \foreach \ang in {0, 45, 90, 135, 180, 225, 270, 315} {
    \fill[blue!90!black] (\ang:2.0) circle (0.12);
    \draw[line width=0.6pt, black] (\ang:2.0) circle (0.12);
    \node[white, font=\footnotesize\bfseries] at (\ang:2.0) {$-$};
  }
  
  % Shell 3: 3 electrons
  \foreach \ang in {30, 150, 270} {
    \fill[blue!90!black] (\ang:2.8) circle (0.12);
    \draw[line width=0.6pt, black] (\ang:2.8) circle (0.12);
    \node[white, font=\footnotesize\bfseries] at (\ang:2.8) {$-$};
  }
  
  % Caption
  \node[below=8pt, font=\bfseries] at (0,-3.0) {Mo hinh nguyen tu Nhom (Al)};
\end{tikzpicture}
\end{document}
"""

items = [
    ('do_thi_cau5', tex_cau5),
    ('do_thi_cau6', tex_cau6),
    ('mo_hinh_carbon', tex_carbon),
    ('mo_hinh_aluminium', tex_al)
]

for name, code in items:
    tex_path = os.path.join(OUT_DIR, f"{name}.tex")
    pdf_path = os.path.join(OUT_DIR, f"{name}.pdf")
    png_path = os.path.join(OUT_DIR, f"{name}.png")
    
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(code)
    
    subprocess.run(["pdflatex", "-interaction=nonstopmode", f"{name}.tex"], cwd=OUT_DIR, capture_output=True)
    if os.path.exists(pdf_path):
        doc = pymupdf.open(pdf_path)
        pix = doc[0].get_pixmap(dpi=300)
        pix.save(png_path)
        print(f"Successfully rendered: {png_path} ({pix.width}x{pix.height})")
    else:
        print(f"Error rendering {name}")

print("All asset rendering finished.")
