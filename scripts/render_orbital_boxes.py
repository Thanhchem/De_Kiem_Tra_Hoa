import os
import shutil
import subprocess
import pymupdf

SAN_PHAM_DIR = r"c:\Antigravity_Thanh\San_Pham"
CROP_DIR = r"c:\Antigravity_Thanh\San_Pham\images_crop_101"
SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
ARTIFACT_DIR = r"C:\Users\Admin\.gemini\antigravity\brain\a6ca2d96-ae73-43ed-8b7f-ad044c4b259a"

os.makedirs(SAN_PHAM_DIR, exist_ok=True)
os.makedirs(CROP_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

# Common TikZ template
TIKZ_TEMPLATE = r"""\documentclass[tikz,border=2pt]{standalone}
\usepackage{tgheros}
\usetikzlibrary{arrows.meta}

\begin{document}
\begin{tikzpicture}[line cap=round, line join=round]
  \def\s{0.7} % kích thước mỗi ô vuông 0.7cm x 0.7cm
  \tikzset{
    box/.style={line width=1.2pt, draw=black},
    spin_up/.style={line width=1.2pt, -{Stealth[length=4.5pt, width=3.4pt]}, draw=black},
    spin_down/.style={line width=1.2pt, -{Stealth[length=4.5pt, width=3.4pt]}, draw=black}
  }

__CONTENT__

\end{tikzpicture}
\end{document}
"""

# 1. HÌNH A: 3 ô orbital, mỗi ô 1 mũi tên lên [↑] [↑] [↑]
draw_a = r"""
  % 3 ô liền kề
  \draw[box] (0, 0) rectangle (\s, \s);
  \draw[box] (\s, 0) rectangle (2*\s, \s);
  \draw[box] (2*\s, 0) rectangle (3*\s, \s);

  % Ô 1: 1 electron quay lên ở giữa
  \draw[spin_up] (0.5*\s, 0.12*\s) -- (0.5*\s, 0.88*\s);

  % Ô 2: 1 electron quay lên ở giữa
  \draw[spin_up] (1.5*\s, 0.12*\s) -- (1.5*\s, 0.88*\s);

  % Ô 3: 1 electron quay lên ở giữa
  \draw[spin_up] (2.5*\s, 0.12*\s) -- (2.5*\s, 0.88*\s);
"""

# 2. HÌNH B: 1 ô orbital, 2 mũi tên cùng quay lên [↑↑]
draw_b = r"""
  % 1 ô đơn
  \draw[box] (0, 0) rectangle (\s, \s);

  % 2 electron cùng quay lên song song
  \draw[spin_up] (0.32*\s, 0.12*\s) -- (0.32*\s, 0.88*\s);
  \draw[spin_up] (0.68*\s, 0.12*\s) -- (0.68*\s, 0.88*\s);
"""

# 3. HÌNH C: 3 ô orbital: [↑↓] [↑ ] [  ]
draw_c = r"""
  % 3 ô liền kề
  \draw[box] (0, 0) rectangle (\s, \s);
  \draw[box] (\s, 0) rectangle (2*\s, \s);
  \draw[box] (2*\s, 0) rectangle (3*\s, \s);

  % Ô 1: cặp e ghép đôi (lên - xuống)
  \draw[spin_up] (0.32*\s, 0.12*\s) -- (0.32*\s, 0.88*\s);
  \draw[spin_down] (0.68*\s, 0.88*\s) -- (0.68*\s, 0.12*\s);

  % Ô 2: 1 e quay lên ở giữa
  \draw[spin_up] (1.5*\s, 0.12*\s) -- (1.5*\s, 0.88*\s);

  % Ô 3: ô trống
"""

# 4. HÌNH D: 3 ô orbital: [↑↑] [↑ ] [↑ ]
draw_d = r"""
  % 3 ô liền kề
  \draw[box] (0, 0) rectangle (\s, \s);
  \draw[box] (\s, 0) rectangle (2*\s, \s);
  \draw[box] (2*\s, 0) rectangle (3*\s, \s);

  % Ô 1: 2 e cùng quay lên
  \draw[spin_up] (0.32*\s, 0.12*\s) -- (0.32*\s, 0.88*\s);
  \draw[spin_up] (0.68*\s, 0.12*\s) -- (0.68*\s, 0.88*\s);

  % Ô 2: 1 e quay lên ở giữa
  \draw[spin_up] (1.5*\s, 0.12*\s) -- (1.5*\s, 0.88*\s);

  % Ô 3: 1 e quay lên ở giữa
  \draw[spin_up] (2.5*\s, 0.12*\s) -- (2.5*\s, 0.88*\s);
"""

# 5. TỔNG HỢP CẢ 4 PHƯƠNG ÁN (A, B, C, D)
draw_all = r"""
  \node[font=\bfseries\fontsize{11pt}{13pt}\selectfont] at (-0.3, 0.35) {A.};
  \draw[box] (0, 0) rectangle (\s, \s);
  \draw[box] (\s, 0) rectangle (2*\s, \s);
  \draw[box] (2*\s, 0) rectangle (3*\s, \s);
  \draw[spin_up] (0.5*\s, 0.12*\s) -- (0.5*\s, 0.88*\s);
  \draw[spin_up] (1.5*\s, 0.12*\s) -- (1.5*\s, 0.88*\s);
  \draw[spin_up] (2.5*\s, 0.12*\s) -- (2.5*\s, 0.88*\s);

  \begin{scope}[xshift=3.2cm]
    \node[font=\bfseries\fontsize{11pt}{13pt}\selectfont] at (-0.3, 0.35) {B.};
    \draw[box] (0, 0) rectangle (\s, \s);
    \draw[spin_up] (0.32*\s, 0.12*\s) -- (0.32*\s, 0.88*\s);
    \draw[spin_up] (0.68*\s, 0.12*\s) -- (0.68*\s, 0.88*\s);
  \end{scope}

  \begin{scope}[xshift=5.3cm]
    \node[font=\bfseries\fontsize{11pt}{13pt}\selectfont] at (-0.3, 0.35) {C.};
    \draw[box] (0, 0) rectangle (\s, \s);
    \draw[box] (\s, 0) rectangle (2*\s, \s);
    \draw[box] (2*\s, 0) rectangle (3*\s, \s);
    \draw[spin_up] (0.32*\s, 0.12*\s) -- (0.32*\s, 0.88*\s);
    \draw[spin_down] (0.68*\s, 0.88*\s) -- (0.68*\s, 0.12*\s);
    \draw[spin_up] (1.5*\s, 0.12*\s) -- (1.5*\s, 0.88*\s);
  \end{scope}

  \begin{scope}[xshift=8.7cm]
    \node[font=\bfseries\fontsize{11pt}{13pt}\selectfont] at (-0.3, 0.35) {D.};
    \draw[box] (0, 0) rectangle (\s, \s);
    \draw[box] (\s, 0) rectangle (2*\s, \s);
    \draw[box] (2*\s, 0) rectangle (3*\s, \s);
    \draw[spin_up] (0.32*\s, 0.12*\s) -- (0.32*\s, 0.88*\s);
    \draw[spin_up] (0.68*\s, 0.12*\s) -- (0.68*\s, 0.88*\s);
    \draw[spin_up] (1.5*\s, 0.12*\s) -- (1.5*\s, 0.88*\s);
    \draw[spin_up] (2.5*\s, 0.12*\s) -- (2.5*\s, 0.88*\s);
  \end{scope}
"""

FIGURES = {
    "orbital_cau5_A": draw_a,
    "orbital_cau5_B": draw_b,
    "orbital_cau5_C": draw_c,
    "orbital_cau5_D": draw_d,
    "orbital_cau5_all": draw_all,
}

def render_figure(name, content):
    tex_str = TIKZ_TEMPLATE.replace("__CONTENT__", content)
    tex_file = os.path.join(SCRATCH_DIR, f"{name}.tex")
    pdf_file = os.path.join(SCRATCH_DIR, f"{name}.pdf")
    svg_file = os.path.join(SCRATCH_DIR, f"{name}.svg")

    with open(tex_file, "w", encoding="utf-8") as f:
        f.write(tex_str)

    res = subprocess.run(["pdflatex", "-interaction=nonstopmode", f"{name}.tex"], cwd=SCRATCH_DIR, capture_output=True, text=True)
    if res.returncode != 0:
        print(f"Error compiling {name}:", res.stdout[-400:])
        return

    # Render PNG via PyMuPDF at 400 DPI
    doc = pymupdf.open(pdf_file)
    page = doc[0]
    pix = page.get_pixmap(dpi=400)
    png_file = os.path.join(SCRATCH_DIR, f"{name}.png")
    pix.save(png_file)
    print(f"Rendered {name}.png ({pix.width}x{pix.height})")

    # Render SVG via dvisvgm
    try:
        subprocess.run(["dvisvgm", "--pdf", f"--output={svg_file}", f"{name}.pdf"], cwd=SCRATCH_DIR, capture_output=True)
    except Exception as e:
        pass

    # Copy files
    for f_ext in [".tex", ".pdf", ".png", ".svg"]:
        src = os.path.join(SCRATCH_DIR, f"{name}{f_ext}")
        if os.path.exists(src):
            shutil.copyfile(src, os.path.join(SAN_PHAM_DIR, f"{name}{f_ext}"))
            shutil.copyfile(src, os.path.join(ARTIFACT_DIR, f"{name}{f_ext}"))

    # Also update the images in images_crop_101 for the exam!
    if name == "orbital_cau5_A":
        shutil.copyfile(png_file, os.path.join(CROP_DIR, "cau5_a.png"))
    elif name == "orbital_cau5_B":
        shutil.copyfile(png_file, os.path.join(CROP_DIR, "cau5_b.png"))
    elif name == "orbital_cau5_C":
        shutil.copyfile(png_file, os.path.join(CROP_DIR, "cau5_c.png"))
    elif name == "orbital_cau5_D":
        shutil.copyfile(png_file, os.path.join(CROP_DIR, "cau5_d.png"))

if __name__ == "__main__":
    for name, content in FIGURES.items():
        render_figure(name, content)
    print("All orbital figures rendered successfully!")
