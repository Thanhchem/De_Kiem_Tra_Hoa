import os
import subprocess
import pymupdf

TEX_CONTENT = r"""\documentclass[tikz,border=10pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{tgheros}
\usepackage{sansmath}
\usepackage{amsmath}
\sansmath
\renewcommand{\familydefault}{\sfdefault}

\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, line cap=round, line join=round]
  \tikzset{
    axis/.style={line width=1.1pt, draw=black},
    tick/.style={line width=0.8pt, draw=black},
    peak/.style={line width=1.8pt, draw=black!90},
    axis_lbl/.style={font=\sffamily\fontsize{10pt}{12pt}\selectfont},
    tick_num/.style={font=\sffamily\fontsize{9.5pt}{11pt}\selectfont},
    label_peak/.style={font=\sffamily\fontsize{9.5pt}{11.5pt}\selectfont}
  }

  % Coordinates:
  % X: 0 to 45 (1 unit x = 0.16 cm) -> 0 to 7.2 cm
  % Y: 0 to 110 (1 unit y = 0.045 cm) -> 0 to 5.0 cm
  \def\xs{0.16}
  \def\ys{0.045}

  % Axes
  \draw[axis, ->, >=stealth] (0, 0) -- (46*\xs, 0) node[below=4pt, right=0pt, axis_lbl] {$m/z^{(+)}$};
  \draw[axis, ->, >=stealth] (0, 0) -- (0, 115*\ys) node[above=4pt, left=0pt, axis_lbl, align=center] {Tỉ lệ \% số nguyên tử};

  % Y-ticks: 0, 25, 50, 75, 100
  \foreach \y in {25, 50, 75, 100} {
    \draw[tick] (-0.12, \y*\ys) -- (0, \y*\ys);
    \node[left=3pt, tick_num] at (0, \y*\ys) {\y};
    \draw[dotted, gray!50] (0, \y*\ys) -- (44*\xs, \y*\ys);
  }
  \node[left=3pt, tick_num] at (0, 0) {0};

  % X-ticks: 10, 20, 30, 40
  \foreach \x in {10, 20, 30, 40} {
    \draw[tick] (\x*\xs, -0.12) -- (\x*\xs, 0);
    \node[below=3pt, tick_num] at (\x*\xs, 0) {\x};
  }

  % Peaks:
  % 20Ne: 90.0% at x=20
  \draw[peak] (20*\xs, 0) -- (20*\xs, 90.0*\ys);
  \node[above=3pt, label_peak, align=center] at (20*\xs, 90.0*\ys) {$^{20}_{10}\text{Ne}\ (90,0\%)$};

  % 21Ne: 1.0% at x=21
  \draw[peak] (21*\xs, 0) -- (21*\xs, 1.0*\ys);
  \node[above=4pt, left=-2pt, label_peak, align=center] at (21*\xs, 8.0*\ys) {$^{21}_{10}\text{Ne}\ (1,0\%)$};

  % 22Ne: 9.0% at x=22
  \draw[peak] (22*\xs, 0) -- (22*\xs, 9.0*\ys);
  \node[above=3pt, right=-4pt, label_peak, align=center] at (22*\xs, 14.0*\ys) {$^{22}_{10}\text{Ne}\ (9,0\%)$};

\end{tikzpicture}
\end{document}
"""

SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
SAN_PHAM_IMG = r"c:\Antigravity_Thanh\San_Pham\images_hoa10\pho_khoi_luong_neon_1201.png"
os.makedirs(SCRATCH_DIR, exist_ok=True)

tex_path = os.path.join(SCRATCH_DIR, "pho_neon.tex")
with open(tex_path, "w", encoding="utf-8") as f:
    f.write(TEX_CONTENT)

# Compile with pdflatex
cmd = ["pdflatex", "-interaction=nonstopmode", "-output-directory", SCRATCH_DIR, tex_path]
res = subprocess.run(cmd, capture_output=True, text=True)
pdf_path = os.path.join(SCRATCH_DIR, "pho_neon.pdf")

if os.path.exists(pdf_path):
    doc = pymupdf.open(pdf_path)
    page = doc[0]
    # 300 DPI (zoom 300/72 = 4.166)
    zoom = 4.166
    mat = pymupdf.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    pix.save(SAN_PHAM_IMG)
    print(f"Saved image to: {SAN_PHAM_IMG}")
else:
    print("Compilation failed!")
    print(res.stdout)
    print(res.stderr)
