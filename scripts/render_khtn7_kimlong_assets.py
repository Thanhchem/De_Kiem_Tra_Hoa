# -*- coding: utf-8 -*-
import os
import subprocess
import pymupdf

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham\images_khtn7_kimlong"
os.makedirs(OUT_DIR, exist_ok=True)

# 1. Sulfur atomic model
tex_s = r"""\documentclass[border=5pt]{standalone}
\usepackage[utf8]{inputenc}
\usepackage{times}
\usepackage{tikz}

\begin{document}
\begin{tikzpicture}[font=\bfseries]
  \fill[red!85!black] (0,0) circle (0.45cm);
  \node[white, scale=1.3] at (0,0) {16+};
  \draw[dashed, line width=1pt, blue!75!black] (0,0) circle (1.1cm);
  \fill[blue!85!black] (90:1.1cm) circle (3.2pt);
  \fill[blue!85!black] (270:1.1cm) circle (3.2pt);
  \draw[dashed, line width=1pt, blue!75!black] (0,0) circle (1.85cm);
  \foreach \a in {0, 45, 90, 135, 180, 225, 270, 315} {
    \fill[blue!85!black] (\a:1.85cm) circle (3.2pt);
  }
  \draw[dashed, line width=1pt, blue!75!black] (0,0) circle (2.6cm);
  \foreach \a in {30, 90, 150, 210, 270, 330} {
    \fill[blue!85!black] (\a:2.6cm) circle (3.2pt);
  }
\end{tikzpicture}
\end{document}
"""

# 2. Aluminium atomic model
tex_al = r"""\documentclass[border=5pt]{standalone}
\usepackage[utf8]{inputenc}
\usepackage{times}
\usepackage{tikz}

\begin{document}
\begin{tikzpicture}[font=\bfseries]
  \fill[red!85!black] (0,0) circle (0.45cm);
  \node[white, scale=1.3] at (0,0) {13+};
  \draw[dashed, line width=1pt, blue!75!black] (0,0) circle (1.1cm);
  \fill[blue!85!black] (90:1.1cm) circle (3.2pt);
  \fill[blue!85!black] (270:1.1cm) circle (3.2pt);
  \draw[dashed, line width=1pt, blue!75!black] (0,0) circle (1.85cm);
  \foreach \a in {0, 45, 90, 135, 180, 225, 270, 315} {
    \fill[blue!85!black] (\a:1.85cm) circle (3.2pt);
  }
  \draw[dashed, line width=1pt, blue!75!black] (0,0) circle (2.6cm);
  \foreach \a in {90, 210, 330} {
    \fill[blue!85!black] (\a:2.6cm) circle (3.2pt);
  }
\end{tikzpicture}
\end{document}
"""

# 3. s-t graph
tex_st = r"""\documentclass[border=6pt]{standalone}
\usepackage[utf8]{inputenc}
\usepackage{times}
\usepackage{tikz}
\usetikzlibrary{arrows.meta}

\begin{document}
\begin{tikzpicture}[x=2.2cm, y=0.45cm, >=Latex, font=\bfseries]
  \draw[dashed, line width=0.65pt, gray!60] (0,0) grid[xstep=0.5, ystep=2.5] (2.0,10.0);
  \draw[->, line width=1.3pt] (0,0) -- (2.35,0) node[below=3pt] {$t\text{ (h)}$};
  \draw[->, line width=1.3pt] (0,0) -- (0,11.2) node[left=3pt] {$s\text{ (km)}$};
  \node[below left=2pt] at (0,0) {O};
  \node[below=3pt] at (0.5,0) {0,5};
  \node[below=3pt] at (1.0,0) {1,0};
  \node[below=3pt] at (1.5,0) {1,5};
  \node[below=3pt] at (2.0,0) {2,0};
  \node[left=3pt] at (0,2.5) {2,5};
  \node[left=3pt] at (0,5.0) {5,0};
  \node[left=3pt] at (0,7.5) {7,5};
  \node[left=3pt] at (0,10.0) {10,0};
  \draw[line width=2.0pt, red!85!black] (0,0) -- (2.0,10.0);
  \foreach \t/\s in {0/0, 0.5/2.5, 1.0/5.0, 1.5/7.5, 2.0/10.0} {
    \fill[red!85!black] (\t,\s) circle (2.8pt);
  }
\end{tikzpicture}
\end{document}
"""

def render():
    for name, code in [('so_do_sulfur', tex_s), ('so_do_aluminium', tex_al), ('do_thi_quang_duong_thoi_gian', tex_st)]:
        tex_file = os.path.join(OUT_DIR, f"{name}.tex")
        with open(tex_file, "w", encoding="utf-8") as f:
            f.write(code)
        cmd = ["pdflatex", "-interaction=nonstopmode", f"-output-directory={OUT_DIR}", tex_file]
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pdf_file = os.path.join(OUT_DIR, f"{name}.pdf")
        png_file = os.path.join(OUT_DIR, f"{name}.png")
        if os.path.exists(pdf_file):
            doc = pymupdf.open(pdf_file)
            pix = doc[0].get_pixmap(dpi=300)
            pix.save(png_file)
            doc.close()
            print(f"Rendered: {png_file}")
            for ext in [".aux", ".log", ".pdf", ".tex"]:
                p = os.path.join(OUT_DIR, f"{name}{ext}")
                if os.path.exists(p):
                    os.remove(p)

if __name__ == "__main__":
    render()
