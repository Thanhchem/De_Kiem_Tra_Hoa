import os
import shutil
import subprocess

SAN_PHAM_DIR = r"C:\Antigravity_Thanh\San_Pham"
SCRATCH_DIR = r"C:\Antigravity_Thanh\scratch"
ARTIFACT_DIR = r"C:\Users\Admin\.gemini\antigravity\brain\646435be-5264-4e8f-a76a-55d03927ffcc"

os.makedirs(SAN_PHAM_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

# -------------------------------------------------------------
# 1. TEX CODE FOR SDS (Sodium Dodecyl Sulfate)
# -------------------------------------------------------------
tex_sds = r"""\documentclass[tikz,border=12pt]{standalone}
\usepackage{sansmath}
\usepackage{amsmath}
\usetikzlibrary{calc}
\sansmath
\renewcommand{\familydefault}{\sfdefault}

\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, line cap=round, line join=round]
  \tikzset{
    bond/.style={line width=1.15pt, draw=black},
    dbond/.style={line width=1.05pt, draw=black},
    atom/.style={inner sep=1.5pt, font=\sffamily\fontsize{14pt}{16pt}\selectfont}
  }

  % Central Sulfur
  \coordinate (S_pos) at (0, 0);
  \node[atom] (S) at (S_pos) {S};

  % Left Oxygen O1
  \coordinate (O1_pos) at (-1.10, 0);
  \node[atom] (O1) at (O1_pos) {O};
  \draw[bond] (O1) -- (S);

  % Top Oxygen
  \coordinate (Otop_pos) at (0, 1.05);
  \node[atom] (Otop) at (Otop_pos) {O};
  \draw[dbond] ($(S.north) + (-0.07, 0.04)$) -- ($(Otop.south) + (-0.07, -0.04)$);
  \draw[dbond] ($(S.north) + (0.07, 0.04)$) -- ($(Otop.south) + (0.07, -0.04)$);

  % Bottom Oxygen
  \coordinate (Obot_pos) at (0, -1.05);
  \node[atom] (Obot) at (Obot_pos) {O};
  \draw[dbond] ($(S.south) + (-0.07, -0.04)$) -- ($(Obot.north) + (-0.07, 0.04)$);
  \draw[dbond] ($(S.south) + (0.07, -0.04)$) -- ($(Obot.north) + (0.07, 0.04)$);

  % Right Oxygen with negative charge O^-
  \coordinate (Oright_pos) at (1.10, 0);
  \node[atom] (Oright) at (Oright_pos) {O$^{\boldsymbol{-}}$};
  \draw[bond] (S) -- (Oright);

  % Sodium cation Na+
  \coordinate (Na_pos) at (2.05, 0);
  \node[atom] (Na) at (Na_pos) {Na$^{\boldsymbol{+}}$};

  % Ellipse enclosing the head group (hydrophilic head)
  \coordinate (head_center) at (0.55, 0);
  \draw[line width=0.85pt, draw=black!90] (head_center) ellipse (2.25cm and 1.55cm);

  % Alkyl chain: 12 carbons (C1 to C12) - Hydrophobic tail
  \def\dx{0.55}
  \def\dy{0.34}

  \coordinate (C12) at ($(O1_pos) + (-\dx, \dy)$);
  \coordinate (C11) at ($(C12) + (-\dx, -\dy)$);
  \coordinate (C10) at ($(C11) + (-\dx, \dy)$);
  \coordinate (C9)  at ($(C10) + (-\dx, -\dy)$);
  \coordinate (C8)  at ($(C9)  + (-\dx, \dy)$);
  \coordinate (C7)  at ($(C8)  + (-\dx, -\dy)$);
  \coordinate (C6)  at ($(C7)  + (-\dx, \dy)$);
  \coordinate (C5)  at ($(C6)  + (-\dx, -\dy)$);
  \coordinate (C4)  at ($(C5)  + (-\dx, \dy)$);
  \coordinate (C3)  at ($(C4)  + (-\dx, -\dy)$);
  \coordinate (C2)  at ($(C3)  + (-\dx, \dy)$);
  \coordinate (C1)  at ($(C2)  + (-\dx, -\dy)$);

  % Draw alkyl chain
  \draw[bond] (C1) -- (C2) -- (C3) -- (C4) -- (C5) -- (C6) -- (C7) -- (C8) -- (C9) -- (C10) -- (C11) -- (C12);
  
  % Bond C12 straight to O1
  \draw[bond] (C12) -- (O1);

\end{tikzpicture}
\end{document}
"""

# -------------------------------------------------------------
# 2. TEX CODE FOR ENZYME ACTIVITY VS pH (Pepsin & Trypsin)
# -------------------------------------------------------------
tex_enzyme = r"""\documentclass[tikz,border=12pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{sansmath}
\usepackage{amsmath}
\usetikzlibrary{calc,arrows.meta}
\sansmath
\renewcommand{\familydefault}{\sfdefault}

\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, line cap=round, line join=round, >=Stealth]
  \tikzset{
    axis/.style={line width=1.1pt, draw=black},
    tick/.style={line width=0.9pt, draw=black},
    curve/.style={line width=1.35pt, draw=black},
    dashed_line/.style={line width=0.85pt, draw=black, dashed, dash pattern=on 3pt off 2.5pt},
    box_node/.style={draw=black, line width=0.85pt, inner sep=4.5pt, font=\sffamily\fontsize{11.5pt}{13.5pt}\selectfont, fill=white},
    ybox_node/.style={draw=black, line width=0.85pt, inner sep=4.5pt, font=\sffamily\fontsize{11pt}{13pt}\selectfont, rotate=90, fill=white},
    label_num/.style={font=\sffamily\fontsize{11pt}{13pt}\selectfont, anchor=north},
    axis_lbl/.style={font=\sffamily\fontsize{12pt}{14pt}\selectfont}
  }

  % Scaling parameters:
  \def\xscale{1.05}
  \def\peakheight{3.6} % peak height
  \def\ymax{5.2} % y-axis height
  \def\xmax{12.4} % x-axis end

  % Axes
  % X-axis from 0 to xmax
  \draw[axis, ->] (0, 0) -- (\xmax, 0);
  % Y-axis from 0 to ymax
  \draw[axis, ->] (0, 0) -- (0, \ymax);

  % Ticks on X-axis (0 to 10)
  \foreach \i in {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10} {
    \ifnum\i>0
      \draw[tick] (\i*\xscale, 0) -- (\i*\xscale, 0.18);
    \fi
    \node[label_num] at (\i*\xscale, -0.15) {\textbf{\i}};
  }

  % Label "pH" at the end of X-axis
  \node[axis_lbl, anchor=north] at (\xmax - 0.45, -0.15) {\textbf{pH}};

  % Label for Y-axis inside rectangular box, rotated 90 degrees
  \node[ybox_node] at (-0.75, 3.2) {\textbf{Độ hoạt động của enzyme}};

  % Curve 1: Pepsin (pH 0 to 4, peak at 2)
  \draw[curve] plot[domain=0:4, samples=100] (\x* \xscale, {\peakheight * (sin(\x * 180 / 4))^2});

  % Dashed vertical line for Pepsin (pH = 2)
  \draw[dashed_line] (2*\xscale, \peakheight) -- (2*\xscale, 0);

  % Pepsin box label above the peak
  \node[box_node] at (2*\xscale, \peakheight + 0.62) {\textbf{Pepsin}};

  % Curve 2: Trypsin (pH 6 to 10, peak at 8)
  \draw[curve] plot[domain=6:10, samples=100] (\x* \xscale, {\peakheight * (sin((\x - 6) * 180 / 4))^2});

  % Dashed vertical line for Trypsin (pH = 8)
  \draw[dashed_line] (8*\xscale, \peakheight) -- (8*\xscale, 0);

  % Trypsin box label above the peak
  \node[box_node] at (8*\xscale, \peakheight + 0.62) {\textbf{Trypsin}};

\end{tikzpicture}
\end{document}
"""

def build_asset(name, tex_content):
    tex_path = os.path.join(SCRATCH_DIR, f"{name}.tex")
    pdf_path = os.path.join(SCRATCH_DIR, f"{name}.pdf")
    png_prefix = os.path.join(SCRATCH_DIR, f"{name}_render")
    svg_path = os.path.join(SCRATCH_DIR, f"{name}.svg")

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_content)

    print(f"Compiling {name} via pdflatex...")
    p1 = subprocess.run(["pdflatex", "-interaction=nonstopmode", f"{name}.tex"], cwd=SCRATCH_DIR, capture_output=True)
    if p1.returncode != 0:
        print(f"Error compiling {name}:", p1.stdout.decode("utf-8", errors="replace")[-500:])
        return False

    print(f"Rendering {name} to 300 DPI PNG via pdftoppm...")
    subprocess.run(["pdftoppm", "-png", "-r", "300", f"{name}.pdf", f"{name}_render"], cwd=SCRATCH_DIR, capture_output=True)

    print(f"Rendering {name} to SVG via dvisvgm...")
    subprocess.run(["dvisvgm", "--pdf", f"--output={svg_path}", f"{name}.pdf"], cwd=SCRATCH_DIR, capture_output=True)

    rendered_png = f"{png_prefix}-1.png"

    # Copy to San_Pham
    dest_tex = os.path.join(SAN_PHAM_DIR, f"{name}.tex")
    dest_pdf = os.path.join(SAN_PHAM_DIR, f"{name}.pdf")
    dest_svg = os.path.join(SAN_PHAM_DIR, f"{name}.svg")
    dest_png = os.path.join(SAN_PHAM_DIR, f"{name}.png")

    shutil.copyfile(tex_path, dest_tex)
    shutil.copyfile(pdf_path, dest_pdf)
    if os.path.exists(svg_path):
        shutil.copyfile(svg_path, dest_svg)
    if os.path.exists(rendered_png):
        shutil.copyfile(rendered_png, dest_png)

    # Also copy to Artifact directory so it can be viewed / embedded
    art_png = os.path.join(ARTIFACT_DIR, f"{name}.png")
    art_svg = os.path.join(ARTIFACT_DIR, f"{name}.svg")
    if os.path.exists(rendered_png):
        shutil.copyfile(rendered_png, art_png)
    if os.path.exists(svg_path):
        shutil.copyfile(svg_path, art_svg)

    print(f"Successfully generated all assets for {name} in {SAN_PHAM_DIR} and {ARTIFACT_DIR}!")
    return True

build_asset("cau_truc_sds", tex_sds)
build_asset("do_thi_enzyme_ph", tex_enzyme)
