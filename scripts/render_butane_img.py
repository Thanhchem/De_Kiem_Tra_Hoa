# -*- coding: utf-8 -*-
"""
Render ảnh công thức cấu tạo butane độ nét cao (600 DPI)
Dùng TikZ / LaTeX và PyMuPDF
"""
import os
import subprocess
import fitz

SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
OUT_IMG = r"c:\Antigravity_Thanh\San_Pham\images_khtn9\khtn9_butane_structure.png"
os.makedirs(SCRATCH_DIR, exist_ok=True)
os.makedirs(os.path.dirname(OUT_IMG), exist_ok=True)

tex_code = r"""\documentclass[border=4pt]{standalone}
\usepackage{tikz}
\usepackage{mathptmx}
\begin{document}
\begin{tikzpicture}[scale=0.95, every node/.style={font=\Large\bfseries}]
  \foreach \i in {1,2,3,4} {
    \node (C\i) at (1.4*\i, 0) {C};
    \node (Ht\i) at (1.4*\i, 1.1) {H};
    \node (Hb\i) at (1.4*\i, -1.1) {H};
    \draw[line width=1.1pt] (C\i) -- (Ht\i);
    \draw[line width=1.1pt] (C\i) -- (Hb\i);
  }
  \node (Hl) at (0.3, 0) {H};
  \node (Hr) at (6.7, 0) {H};
  \draw[line width=1.1pt] (Hl) -- (C1);
  \draw[line width=1.1pt] (C1) -- (C2);
  \draw[line width=1.1pt] (C2) -- (C3);
  \draw[line width=1.1pt] (C3) -- (C4);
  \draw[line width=1.1pt] (C4) -- (Hr);
\end{tikzpicture}
\end{document}
"""

tex_path = os.path.join(SCRATCH_DIR, "butane.tex")
with open(tex_path, "w", encoding="utf-8") as f:
    f.write(tex_code)

res = subprocess.run(["pdflatex", "-interaction=nonstopmode", "butane.tex"], cwd=SCRATCH_DIR, capture_output=True, text=True)
if res.returncode != 0:
    print("pdflatex error:", res.stderr)
else:
    pdf_path = os.path.join(SCRATCH_DIR, "butane.pdf")
    doc = fitz.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(dpi=400)
    pix.save(OUT_IMG)
    print("Saved butane image to:", OUT_IMG)
