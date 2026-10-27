# -*- coding: utf-8 -*-
import subprocess
import pymupdf
import os

os.makedirs("San_Pham/images_khtn7_nguyenhoang", exist_ok=True)
os.makedirs("scratch", exist_ok=True)

def render_tikz(name, tikz_code):
    tex = r"""\documentclass[border=6pt]{standalone}
\usepackage{tikz}
\usepackage{amsmath}
\begin{document}
""" + tikz_code + r"""
\end{document}
"""
    tex_path = f"scratch/{name}.tex"
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex)
    
    subprocess.run(["pdflatex", "-interaction=nonstopmode", "-output-directory", "scratch", tex_path], capture_output=True)
    doc = pymupdf.open(f"scratch/{name}.pdf")
    pix = doc[0].get_pixmap(dpi=300)
    out_path = f"San_Pham/images_khtn7_nguyenhoang/{name}.png"
    pix.save(out_path)
    print(f"[OK] Generated: {out_path} ({pix.width}x{pix.height})")

c29_code = r"""
\begin{tikzpicture}[x=0.085cm,y=0.0055cm,>=stealth]
  % Axes
  \draw[->,thick] (0,0) -- (72,0) node[right,font=\small] {$t\text{ (s)}$};
  \draw[->,thick] (0,0) -- (0,1050) node[above,font=\small] {$s\text{ (m)}$};
  
  % Ticks X
  \foreach \x in {10,20,30,40,50,60} {
    \draw (\x,15) -- (\x,-15) node[below,font=\small] {\x};
  }
  \node[below left,font=\small] at (0,0) {O};
  
  % Ticks Y
  \foreach \y in {150,300,450,600,750,900} {
    \draw (1.5,\y) -- (-1.5,\y) node[left,font=\small] {\y};
  }
  
  % Dashed lines
  \draw[dashed,thick,gray!80] (0,450) -- (40,450) -- (40,0);
  \draw[dashed,thick,gray!80] (0,900) -- (60,900) -- (60,0);
  
  % Plot lines
  \draw[ultra thick,blue!85!black] (0,0) -- (40,450) -- (60,900);
  
  % Points
  \filldraw[black] (0,0) circle (2pt);
  \filldraw[black] (40,450) circle (2.5pt);
  \filldraw[black] (60,900) circle (2.5pt);
  
  % Labels
  \node[above left,font=\bfseries] at (22,250) {(1)};
  \node[above left,font=\bfseries] at (52,700) {(2)};
\end{tikzpicture}
"""

c30_code = r"""
\begin{tikzpicture}[x=1.6cm,y=0.015cm,>=stealth]
  % Grid
  \draw[xstep=0.8cm,ystep=0.75cm,gray!50,dashed] (0,0) grid (4,300);
  
  % Axes
  \draw[->,thick] (0,0) -- (4.5,0) node[right,font=\small] {$t\text{ (min)}$};
  \draw[->,thick] (0,0) -- (0,350) node[above,font=\small] {$s\text{ (m)}$};
  
  % Ticks X
  \foreach \x in {1,2,3,4} {
    \draw (\x,6) -- (\x,-6) node[below,font=\small] {\x};
  }
  \node[below left,font=\small] at (0,0) {0};
  
  % Ticks Y
  \foreach \y in {100,200,300} {
    \draw (0.06,\y) -- (-0.06,\y) node[left,font=\small] {\y};
  }
  
  % Plot lines
  \draw[ultra thick,red!80!black] (0,0) -- (1,100) -- (3,100) -- (4,300);
  
  % Points
  \filldraw[black] (0,0) circle (2pt);
  \filldraw[black] (0.5,50) circle (2pt) node[above left,font=\bfseries] {A};
  \filldraw[black] (1,100) circle (2pt);
  \node[above,font=\bfseries] at (2,100) {B};
  \filldraw[black] (3,100) circle (2pt);
  \filldraw[black] (3.5,200) circle (2pt) node[above left,font=\bfseries] {C};
  \filldraw[black] (4,300) circle (2pt);
\end{tikzpicture}
"""

if __name__ == "__main__":
    render_tikz("dothi_cau29_hd", c29_code)
    render_tikz("dothi_cau30_hd", c30_code)
