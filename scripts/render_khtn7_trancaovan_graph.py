# -*- coding: utf-8 -*-
import subprocess
import pymupdf
import os

os.makedirs("San_Pham/images_khtn7_trancaovan", exist_ok=True)
os.makedirs("scratch", exist_ok=True)

def render_tikz(name, tikz_code):
    tex = r"""\documentclass[border=6pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{times}
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
    out_path = f"San_Pham/images_khtn7_trancaovan/{name}.png"
    pix.save(out_path)
    print(f"[OK] Generated: {out_path} ({pix.width}x{pix.height})")

# Đồ thị Bài 5: t = 0..6 h, s = 0..260 km
# (0,0) -> (1, 60) -> (2, 120) -> (3, 180) -> (4, 180) -> (5, 220) -> (6, 260)
dothi_bai5_code = r"""
\begin{tikzpicture}[x=1.3cm,y=0.02cm,>=stealth]
  % Grid
  \draw[xstep=1.3cm,ystep=0.8cm,gray!40,dashed] (0,0) grid (6.5,280);
  
  % Axes
  \draw[->,thick] (0,0) -- (7.0,0) node[right,font=\small] {$t\text{ (h)}$};
  \draw[->,thick] (0,0) -- (0,300) node[above,font=\small] {$s\text{ (km)}$};
  
  % Ticks X
  \foreach \x in {1,2,3,4,5,6} {
    \draw (\x,4) -- (\x,-4) node[below,font=\small] {\x};
  }
  \node[below left,font=\small] at (0,0) {0};
  
  % Ticks Y
  \foreach \y in {60,120,180,220,260} {
    \draw (0.08,\y) -- (-0.08,\y) node[left,font=\small] {\y};
  }
  
  % Plot line
  \draw[ultra thick,blue!85!black] (0,0) -- (1,60) -- (2,120) -- (3,180) -- (4,180) -- (5,220) -- (6,260);
  
  % Points
  \filldraw[black] (0,0) circle (2pt);
  \filldraw[black] (1,60) circle (2pt);
  \filldraw[black] (2,120) circle (2pt);
  \filldraw[black] (3,180) circle (2.5pt);
  \filldraw[black] (4,180) circle (2.5pt);
  \filldraw[black] (5,220) circle (2pt);
  \filldraw[black] (6,260) circle (2.5pt);
  
  % Labels
  \node[above,font=\footnotesize\bfseries,blue!80!black] at (3.5,185) {Nghỉ (đứng yên)};
\end{tikzpicture}
"""

if __name__ == "__main__":
    render_tikz("dothi_oto_bai5", dothi_bai5_code)
