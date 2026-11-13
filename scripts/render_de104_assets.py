import os
import subprocess
import pymupdf

SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
SAN_PHAM_IMG_DIR = r"c:\Antigravity_Thanh\San_Pham\images_hoa10"
os.makedirs(SCRATCH_DIR, exist_ok=True)
os.makedirs(SAN_PHAM_IMG_DIR, exist_ok=True)

# 1. PHỔ KHỐI LƯỢNG SILVER (Ag)
TEX_AG = r"""\documentclass[tikz,border=8pt]{standalone}
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
    peak/.style={line width=2.0pt, draw=black!90},
    axis_lbl/.style={font=\sffamily\fontsize{9pt}{11pt}\selectfont},
    tick_num/.style={font=\sffamily\fontsize{8.5pt}{10pt}\selectfont},
    label_peak/.style={font=\sffamily\bfseries\fontsize{8.5pt}{10pt}\selectfont}
  }

  % Coordinates:
  % X-axis: 105 to 111 (width ~ 7cm). Origin x=105 -> (x - 105)*1.2
  % Y-axis: 47.0 to 53.0 (height ~ 5.5cm). Origin y=47.0 -> (y - 47.0)*0.9
  \def\ox{105.0}
  \def\oy{47.0}
  \def\xs{1.2}
  \def\ys{0.9}

  % Axes
  \draw[axis, ->, >=stealth] (0, 0) -- (6.5, 0) node[below=8pt, left=15pt, axis_lbl] {$m/z$};
  \draw[axis, ->, >=stealth] (0, 0) -- (0, 5.8) node[above=4pt, right=-5pt, axis_lbl, align=center] {Phần trăm số nguyên tử đồng vị (\%)};

  % Y-ticks: 47.5, 48, 48.5, 49, 49.5, 50, 50.5, 51, 51.5, 52, 52.5
  \foreach \y in {47.5, 48.0, 48.5, 49.0, 49.5, 50.0, 50.5, 51.0, 51.5, 52.0, 52.5} {
    \draw[tick] (-0.1, {(\y - \oy)*\ys}) -- (0, {(\y - \oy)*\ys});
    \node[left=2pt, tick_num] at (0, {(\y - \oy)*\ys}) {\pgfmathprintnumber[fixed,precision=1]{\y}};
  }

  % X-ticks: 106, 107, 108, 109, 110
  \foreach \x in {106, 107, 108, 109, 110} {
    \draw[tick] ({(\x - \ox)*\xs}, -0.1) -- ({(\x - \ox)*\xs}, 0);
    \node[below=3pt, tick_num] at ({(\x - \ox)*\xs}, 0) {\x};
  }

  % Peak 1: 107Ag = 51.8%
  \draw[peak] ({(107 - \ox)*\xs}, 0) -- ({(107 - \ox)*\xs}, {(51.8 - \oy)*\ys});
  \node[above=2pt, label_peak] at ({(107 - \ox)*\xs}, {(51.8 - \oy)*\ys}) {51,8};

  % Peak 2: 109Ag = 48.2%
  \draw[peak] ({(109 - \ox)*\xs}, 0) -- ({(109 - \ox)*\xs}, {(48.2 - \oy)*\ys});
  \node[above=2pt, label_peak] at ({(109 - \ox)*\xs}, {(48.2 - \oy)*\ys}) {48,2};

\end{tikzpicture}
\end{document}
"""

# 2. HÌNH ORBITAL (a) VÀ (b)
TEX_ORBITALS = r"""\documentclass[tikz,border=8pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{tgheros}
\usepackage{sansmath}
\sansmath
\renewcommand{\familydefault}{\sfdefault}

\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, line cap=round, line join=round]
  \tikzset{
    axis/.style={line width=0.8pt, ->, >=stealth, draw=black},
    lbl/.style={font=\sffamily\itshape\fontsize{10pt}{12pt}\selectfont},
    cap_lbl/.style={font=\sffamily\bfseries\fontsize{10.5pt}{13pt}\selectfont}
  }

  % --- Orbital (a): AO s ---
  \begin{scope}[shift={(0,0)}]
    % Axes
    \draw[axis] (-1.8, 0) -- (1.8, 0) node[right=1pt, lbl] {$y$};
    \draw[axis] (0, -1.8) -- (0, 1.8) node[above=1pt, lbl] {$z$};
    \draw[axis] (1.2, 1.2) -- (-1.2, -1.2) node[below left=0pt, lbl] {$x$};

    % Sphere shading
    \shade[ball color=gray!70, opacity=0.85] (0,0) circle (0.85cm);
    \draw[line width=0.6pt, black!80] (0,0) circle (0.85cm);

    \node[below=2.2cm, cap_lbl] at (0,0) {(a)};
  \end{scope}

  % --- Orbital (b): AO p_y (hình số tám nổi) ---
  \begin{scope}[shift={(5.0,0)}]
    % Axes
    \draw[axis] (-2.2, 0) -- (2.2, 0) node[right=1pt, lbl] {$y$};
    \draw[axis] (0, -1.8) -- (0, 1.8) node[above=1pt, lbl] {$z$};
    \draw[axis] (1.2, 1.2) -- (-1.2, -1.2) node[below left=0pt, lbl] {$x$};

    % Lobe 1 (right, along y)
    \shade[ball color=gray!70, opacity=0.85] (0.8,0) ellipse (0.8cm and 0.45cm);
    \draw[line width=0.6pt, black!80] (0.8,0) ellipse (0.8cm and 0.45cm);

    % Lobe 2 (left, along y)
    \shade[ball color=gray!70, opacity=0.85] (-0.8,0) ellipse (0.8cm and 0.45cm);
    \draw[line width=0.6pt, black!80] (-0.8,0) ellipse (0.8cm and 0.45cm);

    \node[below=2.2cm, cap_lbl] at (0,0) {(b)};
  \end{scope}

\end{tikzpicture}
\end{document}
"""

def render_tikz(tex_code, out_name):
    tex_path = os.path.join(SCRATCH_DIR, f"{out_name}.tex")
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_code)
    
    cmd = ["pdflatex", "-interaction=nonstopmode", f"-output-directory={SCRATCH_DIR}", tex_path]
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    
    pdf_path = os.path.join(SCRATCH_DIR, f"{out_name}.pdf")
    png_path = os.path.join(SAN_PHAM_IMG_DIR, f"{out_name}.png")
    
    if os.path.exists(pdf_path):
        doc = pymupdf.open(pdf_path)
        page = doc[0]
        mat = pymupdf.Matrix(4.166, 4.166) # 300 DPI
        pix = page.get_pixmap(matrix=mat, alpha=False)
        pix.save(png_path)
        print(f"[OK] Đã lưu ảnh: {png_path}")
    else:
        print(f"[LỖI] Không thể biên dịch {out_name}")

render_tikz(TEX_AG, "pho_khoi_luong_silver_ag")
render_tikz(TEX_ORBITALS, "hinh_orbital_s_va_p")
