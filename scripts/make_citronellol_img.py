# -*- coding: utf-8 -*-
import subprocess
import os
import fitz

TEX_CONTENT = r"""\documentclass[border=10pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{tikz}
\usepackage{helvet}
\renewcommand{\familydefault}{\sfdefault}
\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, line width=1.2pt, line cap=round, line join=round]

% --- Cong thuc (A) Menthol ---
\begin{scope}[shift={(0, 0)}]
    \draw (1.5, 2.2) -- (2.3, 1.7) -- (2.3, 0.7) -- (1.5, 0.2) -- (0.7, 0.7) -- (0.7, 1.7) -- cycle;
    \draw (1.5, 2.2) -- (1.5, 2.9) node[above] {CH$_3$};
    \draw (2.3, 1.7) -- (2.9, 2.05) node[right] {OH};
    \draw (1.5, 0.2) -- (1.5, -0.4);
    \draw (1.5, -0.4) -- (1.0, -0.8) node[below] {H$_3$C};
    \draw (1.5, -0.4) -- (2.0, -0.8) node[below] {CH$_3$};
    \node at (1.5, -1.8) {\large\textbf{Công thức (A)}};
\end{scope}

% --- Cong thuc (B) Citronellol ---
\begin{scope}[shift={(4.5, 0)}]
    \node[left] at (0, 0.6) {H$_3$C};
    \draw (0, 0.6) -- (0.8, 0.1);
    \node[below] at (0.8, -0.7) {CH$_3$};
    \draw (0.8, 0.1) -- (0.8, -0.7);
    
    \draw (0.8, 0.1) -- (1.7, 0.6);
    \draw (0.85, 0.22) -- (1.7, 0.7);
    
    \draw (1.7, 0.6) -- (2.5, 0.1) -- (3.3, 0.6) -- (4.1, 0.1);
    \draw (4.1, 0.1) -- (4.1, -0.7) node[below] {CH$_3$};
    \draw (4.1, 0.1) -- (4.9, 0.6) -- (5.7, 0.1) node[right] {OH};
    
    \node at (2.8, -1.8) {\large\textbf{Công thức (B)}};
\end{scope}

% --- Cong thuc (C) Geraniol ---
\begin{scope}[shift={(12.0, 0)}]
    \node[left] at (0, 0.6) {H$_3$C};
    \draw (0, 0.6) -- (0.8, 0.1);
    \node[below] at (0.8, -0.7) {CH$_3$};
    \draw (0.8, 0.1) -- (0.8, -0.7);
    
    \draw (0.8, 0.1) -- (1.7, 0.6);
    \draw (0.85, 0.22) -- (1.7, 0.7);
    
    \draw (1.7, 0.6) -- (2.5, 0.1) -- (3.3, 0.6);
    \draw (3.3, 0.6) -- (3.3, -0.2) node[below] {CH$_3$};
    
    \draw (3.3, 0.6) -- (4.2, 1.1);
    \draw (3.35, 0.72) -- (4.2, 1.2);
    
    \draw (4.2, 1.1) -- (5.0, 0.6) node[right] {OH};
    
    \node at (2.6, -1.8) {\large\textbf{Công thức (C)}};
\end{scope}

\end{tikzpicture}
\end{document}
"""

out_dir = r"c:\Antigravity_Thanh\San_Pham\images_crop_hoa11_de09"
tex_path = os.path.join(out_dir, "citronellol_abc.tex")
with open(tex_path, "w", encoding="utf-8") as f:
    f.write(TEX_CONTENT)

res = subprocess.run(["pdflatex", "-interaction=nonstopmode", "citronellol_abc.tex"], cwd=out_dir, capture_output=True, text=True)
print("pdflatex returncode:", res.returncode)

pdf_path = os.path.join(out_dir, "citronellol_abc.pdf")
doc = fitz.open(pdf_path)
print("Pages:", len(doc))
pix = doc[0].get_pixmap(dpi=300)
img_path = os.path.join(out_dir, "cau4_p2_citronellol_ABC.png")
pix.save(img_path)
print(f"Saved {img_path} ({pix.width}x{pix.height}) successfully!")
