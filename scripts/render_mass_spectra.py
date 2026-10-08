import os
import shutil
import subprocess
import pymupdf

SAN_PHAM_DIR = r"c:\Antigravity_Thanh\San_Pham"
SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
ARTIFACT_DIR = r"C:\Users\Admin\.gemini\antigravity\brain\a6ca2d96-ae73-43ed-8b7f-ad044c4b259a"

os.makedirs(SAN_PHAM_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

# -------------------------------------------------------------------------
# 1. TEX CODE FOR CHLORINE MASS SPECTRUM
# -------------------------------------------------------------------------
tex_cl = r"""\documentclass[tikz,border=15pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{tgheros}
\usepackage{sansmath}
\usepackage{amsmath}
\sansmath
\renewcommand{\familydefault}{\sfdefault}

\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, line cap=round, line join=round]
  \tikzset{
    axis/.style={line width=1.15pt, draw=black},
    tick/.style={line width=0.95pt, draw=black},
    peak/.style={line width=3.0pt, draw=black!90},
    axis_lbl/.style={font=\sffamily\fontsize{10.5pt}{12.5pt}\selectfont},
    tick_num/.style={font=\sffamily\fontsize{10.5pt}{12.5pt}\selectfont},
    label_peak/.style={font=\sffamily\fontsize{10pt}{12pt}\selectfont},
    caption_txt/.style={font=\sffamily\bfseries\itshape\fontsize{11pt}{13.5pt}\selectfont}
  }

  % Coordinates scale:
  % X-axis: 0 to 48 (1 unit x = 0.23 cm -> width ~ 11.0 cm)
  % Y-axis: 0 to 110 (1 unit y = 0.046 cm -> height ~ 5.1 cm)
  \def\xs{0.23}
  \def\ys{0.046}

  % Axes
  \draw[axis] (0, 0) -- (48*\xs, 0);
  \draw[axis] (0, 0) -- (0, 112*\ys);

  % Y-axis ticks and labels
  \foreach \y in {25, 50, 75, 100} {
    \draw[tick] (0, \y*\ys) -- (-0.18, \y*\ys);
    \node[tick_num, anchor=east] at (-0.22, \y*\ys) {\y};
  }

  % Y-axis title (rotated 90)
  \node[axis_lbl, rotate=90, anchor=south] at (-1.15, 55*\ys) {Tỉ lệ \% số nguyên tử};

  % X-axis ticks and labels
  \foreach \x in {10, 20, 30, 40} {
    \draw[tick] (\x*\xs, 0) -- (\x*\xs, -0.18);
    \node[tick_num, anchor=north] at (\x*\xs, -0.22) {\x};
  }

  % X-axis label m/z = A
  \node[axis_lbl, anchor=west] at (42.5*\xs, -0.42) {$m/z = A$};

  % Peaks:
  % Peak 1: 35Cl (a%) at x=35, height = 75.77%
  \draw[peak] (35*\xs, 0) -- (35*\xs, 75.77*\ys);
  \node[label_peak, anchor=south] at (35*\xs - 0.2, 75.77*\ys + 0.12) {$^{35}_{\ 17}\text{Cl}\ (a\%)$};

  % Peak 2: 37Cl (b%) at x=37, height = 24.23%
  \draw[peak] (37*\xs, 0) -- (37*\xs, 24.23*\ys);
  \node[label_peak, anchor=south west, inner sep=1pt] at (36.8*\xs, 24.23*\ys + 0.12) {$^{37}_{\ 17}\text{Cl}\ (b\%)$};

  % Caption
  \node[caption_txt, anchor=north] at (22*\xs, -1.05) {Hình 2.2. Phổ khối lượng của chlorine};

\end{tikzpicture}
\end{document}
"""

# -------------------------------------------------------------------------
# 2. TEX CODE FOR MAGNESIUM MASS SPECTRUM
# -------------------------------------------------------------------------
tex_mg = r"""\documentclass[tikz,border=15pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{tgheros}
\usepackage{sansmath}
\usepackage{amsmath}
\sansmath
\renewcommand{\familydefault}{\sfdefault}

\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, line cap=round, line join=round]
  \tikzset{
    axis/.style={line width=1.15pt, draw=black},
    tick/.style={line width=0.95pt, draw=black},
    peak/.style={line width=2.8pt, draw=black!90},
    axis_lbl/.style={font=\sffamily\fontsize{10.5pt}{12.5pt}\selectfont},
    tick_num/.style={font=\sffamily\fontsize{10.5pt}{12.5pt}\selectfont},
    label_peak/.style={font=\sffamily\fontsize{10pt}{12pt}\selectfont}
  }

  % Coordinates scale:
  % Y-axis: 0 to 90 (1 unit y = 0.055 cm -> 80 is 4.4 cm, 90 is 4.95 cm)
  \def\ys{0.055}

  % Positions of 24, 25, 26 on X-axis:
  \def\xposA{4.2}
  \def\xposB{6.2}
  \def\xposC{8.2}
  \def\xmax{11.5}

  % Axes
  \draw[axis] (0, 0) -- (\xmax, 0);
  \draw[axis] (0, 0) -- (0, 92*\ys);

  % Y-axis ticks and labels
  \foreach \y in {20, 40, 60, 80} {
    \draw[tick] (0, \y*\ys) -- (-0.18, \y*\ys);
    \node[tick_num, anchor=east] at (-0.22, \y*\ys) {\y};
  }

  % Y-axis title (rotated 90)
  \node[axis_lbl, rotate=90, anchor=south] at (-1.1, 46*\ys) {Tỉ lệ \% số nguyên tử};

  % X-axis ticks and labels at 24, 25, 26
  \draw[tick] (\xposA, 0) -- (\xposA, -0.18);
  \node[tick_num, anchor=north] at (\xposA, -0.22) {24};

  \draw[tick] (\xposB, 0) -- (\xposB, -0.18);
  \node[tick_num, anchor=north] at (\xposB, -0.22) {25};

  \draw[tick] (\xposC, 0) -- (\xposC, -0.18);
  \node[tick_num, anchor=north] at (\xposC, -0.22) {26};

  % X-axis label m/z
  \node[axis_lbl, anchor=west] at (\xmax - 0.75, -0.42) {$m/z$};

  % Peaks:
  % Peak 1: 24Mg (78.99%) at xposA
  \draw[peak] (\xposA, 0) -- (\xposA, 78.99*\ys);
  \node[label_peak, anchor=south] at (\xposA + 0.3, 78.99*\ys + 0.12) {$^{24}\text{Mg}\ (78{,}99\%)$};

  % Peak 2: 25Mg (10%) at xposB
  \draw[peak] (\xposB, 0) -- (\xposB, 10*\ys);
  \node[label_peak, anchor=south] at (\xposB, 10*\ys + 0.15) {$^{25}\text{Mg}\ (10\%)$};

  % Peak 3: 26Mg (11.01%) at xposC
  \draw[peak] (\xposC, 0) -- (\xposC, 11.01*\ys);
  \node[label_peak, anchor=south] at (\xposC + 0.5, 11.01*\ys + 0.15) {$^{26}\text{Mg}\ (11{,}01\%)$};

\end{tikzpicture}
\end{document}
"""

def render(name, tex_str):
    tex_file = os.path.join(SCRATCH_DIR, f"{name}.tex")
    pdf_file = os.path.join(SCRATCH_DIR, f"{name}.pdf")
    svg_file = os.path.join(SCRATCH_DIR, f"{name}.svg")

    with open(tex_file, "w", encoding="utf-8") as f:
        f.write(tex_str)

    print(f"Compiling {name}.tex with pdflatex...")
    res = subprocess.run(["pdflatex", "-interaction=nonstopmode", f"{name}.tex"], cwd=SCRATCH_DIR, capture_output=True, text=True)
    if res.returncode != 0:
        print(f"Error compiling {name}:", res.stdout[-500:])
        return

    # Render PNG via PyMuPDF at 300 DPI
    doc = pymupdf.open(pdf_file)
    page = doc[0]
    pix = page.get_pixmap(dpi=300)
    png_file = os.path.join(SCRATCH_DIR, f"{name}.png")
    pix.save(png_file)
    print(f"Rendered PNG: {png_file} ({pix.width}x{pix.height})")

    # Render SVG via dvisvgm
    try:
        subprocess.run(["dvisvgm", "--pdf", f"--output={svg_file}", f"{name}.pdf"], cwd=SCRATCH_DIR, capture_output=True)
        print(f"Rendered SVG: {svg_file}")
    except Exception as e:
        print("dvisvgm error:", e)

    # Copy files to San_Pham and Artifact directory
    for f_ext in [".tex", ".pdf", ".png", ".svg"]:
        src = os.path.join(SCRATCH_DIR, f"{name}{f_ext}")
        if os.path.exists(src):
            shutil.copyfile(src, os.path.join(SAN_PHAM_DIR, f"{name}{f_ext}"))
            shutil.copyfile(src, os.path.join(ARTIFACT_DIR, f"{name}{f_ext}"))
            print(f"Copied {name}{f_ext} to San_Pham and Artifact directory")

if __name__ == "__main__":
    render("pho_khoi_luong_chlorine", tex_cl)
    render("pho_khoi_luong_magnesium", tex_mg)
