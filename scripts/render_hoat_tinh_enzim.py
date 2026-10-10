import os
import shutil
import subprocess
import pymupdf

SAN_PHAM_DIR = r"c:\Antigravity_Thanh\San_Pham"
SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
ARTIFACT_DIR = r"C:\Users\Admin\.gemini\antigravity\brain\0945c976-cd16-48a2-916a-a6e216ca46b3"

os.makedirs(SAN_PHAM_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)
os.makedirs(ARTIFACT_DIR, exist_ok=True)

def generate_tex(label_y="Hoạt tính của enzim"):
    return rf"""\documentclass[tikz,border=10pt]{{standalone}}
\usepackage[utf8]{{vietnam}}
\usepackage{{tgheros}}
\usepackage{{amsmath}}
\renewcommand{{\familydefault}}{{\sfdefault}}
\usetikzlibrary{{arrows.meta,calc}}

\begin{{document}}
\begin{{tikzpicture}}[x=1cm, y=1cm, line cap=round, line join=round]
  \tikzset{{
    axis/.style={{-{{Stealth[scale=1.15, length=6.5pt, width=4.5pt]}}, line width=1.15pt, draw=black}},
    tick/.style={{line width=0.95pt, draw=black}},
    dashed_line/.style={{dash pattern=on 3pt off 2.5pt, line width=0.85pt, draw=black}},
    curve/.style={{line width=1.45pt, draw=black}},
    lbl/.style={{font=\fontsize{{11pt}}{{13pt}}\selectfont}},
    num/.style={{font=\fontsize{{10.5pt}}{{12.5pt}}\selectfont}}
  }}

  % Coordinates scale
  \def\xO{{0}}
  \def\xten{{1.63}}
  \def\xtwenty{{3.35}}
  \def\xthirty{{5.08}}
  \def\xfive{{5.98}}      % 35 deg C
  \def\xpeak{{6.35}}      % ~37 deg C (Optimum)
  \def\xforty{{6.80}}     % 40 deg C
  \def\xmax{{8.45}}
  \def\ymax{{5.30}}
  \def\ypeak{{4.25}}

  % Axes
  \draw[axis] (\xO, 0) -- (\xmax, 0);
  \draw[axis] (\xO, 0) -- (\xO, \ymax);

  % Origin label
  \node[lbl, anchor=north] at (-0.08, -0.06) {{O}};

  % Ticks on X-axis (crossing axis slightly)
  \draw[tick] (\xten, 0.08) -- (\xten, -0.08);
  \draw[tick] (\xtwenty, 0.08) -- (\xtwenty, -0.08);
  \draw[tick] (\xthirty, 0.08) -- (\xthirty, -0.08);

  % Tick 35
  \draw[tick] (\xfive, 0.08) -- (\xfive, -0.08);
  \node[num, anchor=north] at (\xfive, -0.08) {{35}};

  % Tick 40
  \draw[tick] (\xforty, 0.08) -- (\xforty, -0.08);
  \node[num, anchor=north] at (\xforty, -0.08) {{40}};

  % Axis labels
  \node[lbl, anchor=north] at (\xmax - 0.20, -0.08) {{\textsf{{\textit{{t}}\textsuperscript{{o}}}}}};
  \node[lbl, rotate=90, anchor=south] at (-0.45, \ymax*0.5) {{{label_y}}};

  % Dashed projection line from peak to X-axis
  \draw[dashed_line] (\xpeak, \ypeak) -- (\xpeak, 0);

  % Enzyme activity curve
  % Very smooth Catmull-Rom spline faithfully matching original textbook curve
  \draw[curve] plot[smooth, tension=0.65] coordinates {{
    (0.41, 0.25)
    (0.85, 0.42)
    (1.45, 0.68)
    (2.15, 1.08)
    (2.95, 1.70)
    (3.75, 2.38)
    (4.55, 3.05)
    (5.35, 3.70)
    (5.98, 4.10)
    (6.35, 4.25)
    (6.65, 4.16)
    (6.85, 3.96)
    (7.15, 3.42)
    (7.42, 2.45)
    (7.62, 1.40)
    (7.75, 0.46)
  }};

\end{{tikzpicture}}
\end{{document}}
"""

def render_figure(base_name, label_y):
    tex_str = generate_tex(label_y=label_y)
    tex_path = os.path.join(SCRATCH_DIR, f"{base_name}.tex")
    pdf_path = os.path.join(SCRATCH_DIR, f"{base_name}.pdf")
    png_path = os.path.join(SCRATCH_DIR, f"{base_name}.png")
    svg_path = os.path.join(SCRATCH_DIR, f"{base_name}.svg")

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_str)

    print(f"Compiling {base_name}.tex...")
    res = subprocess.run(["pdflatex", "-interaction=nonstopmode", f"{base_name}.tex"], cwd=SCRATCH_DIR, capture_output=True, text=True)
    if res.returncode != 0 and not os.path.exists(pdf_path):
        print(f"Compilation error for {base_name}:", res.stdout[-400:])
        return

    # Render PNG at 400 DPI (High resolution)
    doc = pymupdf.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(dpi=400)
    pix.save(png_path)
    print(f"Rendered PNG: {png_path} ({pix.width}x{pix.height})")

    # Render SVG via dvisvgm
    try:
        subprocess.run(["dvisvgm", "--pdf", f"--output={svg_path}", f"{base_name}.pdf"], cwd=SCRATCH_DIR, capture_output=True)
        print(f"Rendered SVG: {svg_path}")
    except Exception as e:
        print(f"Error generating SVG: {e}")

    # Copy to San_Pham and Artifact directory
    for ext in [".tex", ".pdf", ".png", ".svg"]:
        f_src = os.path.join(SCRATCH_DIR, f"{base_name}{ext}")
        if os.path.exists(f_src):
            shutil.copyfile(f_src, os.path.join(SAN_PHAM_DIR, f"{base_name}{ext}"))
            shutil.copyfile(f_src, os.path.join(ARTIFACT_DIR, f"{base_name}{ext}"))
            print(f"Saved {base_name}{ext} to San_Pham & Artifact")

if __name__ == "__main__":
    # Version 1: Original spelling ("enzim")
    render_figure("do_thi_hoat_tinh_enzim", "Hoạt tính của enzim")
    # Version 2: Modern 2018 curriculum spelling ("enzyme")
    render_figure("do_thi_hoat_tinh_enzyme", "Hoạt tính của enzyme")
