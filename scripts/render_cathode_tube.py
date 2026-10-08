import os
import shutil
import subprocess
import pymupdf

SAN_PHAM_DIR = r"c:\Antigravity_Thanh\San_Pham"
CROP_DIR = r"c:\Antigravity_Thanh\San_Pham\images_crop_101"
SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
ARTIFACT_DIR = r"C:\Users\Admin\.gemini\antigravity\brain\a6ca2d96-ae73-43ed-8b7f-ad044c4b259a"

tex_cathode = r"""\documentclass[tikz,border=18pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{tgheros}
\usepackage{sansmath}
\usepackage{amsmath}
\usetikzlibrary{calc,arrows.meta}
\sansmath
\renewcommand{\familydefault}{\sfdefault}

\begin{document}
\begin{tikzpicture}[x=1cm, y=1cm, line cap=round, line join=round]

  % Styles
  \tikzset{
    wire/.style={line width=1.1pt, draw=black!90},
    glass_edge/.style={line width=0.95pt, draw=blue!45!black!65},
    glass_fill/.style={top color=blue!5!cyan!4, bottom color=blue!12!cyan!8, fill opacity=0.35},
    pointer/.style={line width=0.65pt, draw=black!85},
    label_txt/.style={font=\sffamily\fontsize{11pt}{13.5pt}\selectfont, inner sep=2pt}
  }

  % -------------------------------------------------------------
  % 1. NGUỒN ĐIỆN (Battery / Power Supply) at bottom left
  % -------------------------------------------------------------
  \coordinate (B_bl) at (0.3, -3.2);
  \coordinate (B_br) at (2.4, -3.2);
  \coordinate (B_tr) at (2.4, -2.1);
  \coordinate (B_tl) at (0.3, -2.1);
  
  \coordinate (B_tbl) at ($(B_tl) + (0.45, 0.32)$);
  \coordinate (B_tbr) at ($(B_tr) + (0.45, 0.32)$);
  \coordinate (B_rbr) at ($(B_br) + (0.45, 0.32)$);

  % Top face (Red/Pinkish)
  \filldraw[draw=black!80, line width=0.8pt, top color=red!35!brown!25, bottom color=red!45!brown!35]
    (B_tl) -- (B_tr) -- (B_tbr) -- (B_tbl) -- cycle;

  % Right face (Dark gray)
  \filldraw[draw=black!80, line width=0.8pt, fill=black!25]
    (B_tr) -- (B_tbr) -- (B_rbr) -- (B_br) -- cycle;

  % Front face (Light gray)
  \filldraw[draw=black!80, line width=0.8pt, top color=black!10, bottom color=black!18]
    (B_bl) -- (B_br) -- (B_tr) -- (B_tl) -- cycle;

  % Battery Terminals (Cọc bình)
  % Left terminal (-)
  \filldraw[draw=black!80, line width=0.6pt, fill=black!40] (0.9, -1.95) ellipse (0.12cm and 0.05cm);
  \filldraw[draw=black!80, line width=0.6pt, fill=black!30] (0.78, -1.95) rectangle (1.02, -1.75);
  \filldraw[draw=black!80, line width=0.6pt, fill=black!15] (0.9, -1.75) ellipse (0.12cm and 0.05cm);

  % Right terminal (+)
  \filldraw[draw=black!80, line width=0.6pt, fill=black!40] (1.9, -1.95) ellipse (0.12cm and 0.05cm);
  \filldraw[draw=black!80, line width=0.6pt, fill=black!30] (1.78, -1.95) rectangle (2.02, -1.75);
  \filldraw[draw=black!80, line width=0.6pt, fill=black!15] (1.9, -1.75) ellipse (0.12cm and 0.05cm);

  % -------------------------------------------------------------
  % 2. DÂY NỐI VÀ KÍ HIỆU (+) (-)
  % -------------------------------------------------------------
  % Left wire (-) from terminal (0.9, -1.75) -> left -> up -> Cathode (1.0, 0)
  \draw[wire] (0.9, -1.75) -- (0.9, -1.35) -- (-0.4, -1.35) -- (-0.4, 0) -- (1.0, 0);
  
  % Circle (-) on left wire
  \filldraw[fill=white, draw=black!85, line width=0.8pt] (-0.4, -0.68) circle (0.24cm);
  \node[font=\fontsize{11pt}{12pt}\selectfont\bfseries] at (-0.4, -0.68) {$-$};

  % Right wire (+) from terminal (1.9, -1.75) -> right -> up -> Anode (3.2, -0.52)
  \draw[wire] (1.9, -1.75) -- (1.9, -1.35) -- (3.2, -1.35) -- (3.2, -0.52);

  % Circle (+) on right wire
  \filldraw[fill=white, draw=black!85, line width=0.8pt] (3.2, -0.92) circle (0.24cm);
  \node[font=\fontsize{10.5pt}{12pt}\selectfont\bfseries] at (3.2, -0.92) {$+$};

  % -------------------------------------------------------------
  % 3. ỐNG THỦY TINH CHÂN KHÔNG (Glass Vacuum Tube)
  % -------------------------------------------------------------
  % Background glass body
  \fill[glass_fill] 
    (0.6, 0.55) -- (3.8, 0.55) 
    to[out=0, in=185] (7.5, 0.95)
    to[out=5, in=195] (9.6, 1.55)
    to[out=15, in=90] (10.05, 0)
    to[out=-90, in=-15] (9.6, -1.55)
    to[out=165, in=-5] (7.5, -0.95)
    to[out=175, in=0] (3.8, -0.55)
    -- (0.6, -0.55)
    arc (-90:90:0.25cm and 0.55cm) -- cycle;

  % Left dome cap
  \draw[glass_edge] (0.6, -0.55) arc (-90:90:0.25cm and 0.55cm);

  % Green glow between Cathode and Anode
  \fill[green!55!teal!35, fill opacity=0.45] 
    (1.15, -0.48) rectangle (3.2, 0.48);

  % Cathode Disc (Cực âm) at x=1.1
  \filldraw[draw=black!90, line width=0.8pt, fill=black!75]
    (1.0, 0) ellipse (0.08cm and 0.48cm);
  \filldraw[draw=black!90, line width=0.8pt, fill=black!55]
    (1.15, 0) ellipse (0.08cm and 0.48cm);
  \draw[line width=0.7pt, draw=black!90] (1.0, 0.48) -- (1.15, 0.48);
  \draw[line width=0.7pt, draw=black!90] (1.0, -0.48) -- (1.15, -0.48);

  % Anode Disc with Hole (Cực dương có lỗ) at x=3.2
  \filldraw[draw=orange!60!brown!90, line width=0.9pt, fill=orange!70!brown!60]
    (3.2, 0) ellipse (0.16cm and 0.52cm);
  \draw[orange!50!brown!90, line width=0.8pt] (3.2, 0) ellipse (0.16cm and 0.52cm);
  % Central hole (Lỗ)
  \filldraw[draw=black!85, fill=white, line width=0.7pt] 
    (3.2, 0) ellipse (0.05cm and 0.12cm);

  % Cathode Ray (Tia âm cực) - Green beam
  \draw[line width=3.5pt, green!50!teal!60, opacity=0.35] 
    (1.15, 0) -- (10.05, 0);
  \draw[line width=1.6pt, green!75!black!90] 
    (3.2, 0) -- (10.05, 0);

  % Upper glass tube profile
  \draw[glass_edge] (0.6, 0.55) -- (3.8, 0.55) 
    to[out=0, in=185] (7.5, 0.95)
    to[out=5, in=195] (9.6, 1.55);

  % Lower glass tube profile
  \draw[glass_edge] (0.6, -0.55) -- (3.8, -0.55) 
    to[out=0, in=175] (7.5, -0.95)
    to[out=-5, in=165] (9.6, -1.55);

  % Phosphor Screen (Màn phosphor) at right end
  \filldraw[glass_edge, fill=cyan!15!green!15, fill opacity=0.55]
    (9.6, 1.55) to[out=15, in=90] (10.05, 0) to[out=-90, in=-15] (9.6, -1.55)
    to[out=140, in=-140] (9.6, 1.55) -- cycle;

  % Flared edge rim
  \draw[glass_edge, line width=1.1pt] 
    (9.6, 0) ellipse (0.75cm and 1.55cm);

  % Green spot where beam strikes phosphor
  \fill[green!80!black!90] (10.05, 0) circle (0.08cm);
  \fill[green!50!white, opacity=0.85] (10.05, 0) circle (0.04cm);

  % Glass shine highlights (phản quang thủy tinh)
  \draw[white, line width=1.6pt, opacity=0.6] 
    (4.2, 0.45) to[out=0, in=185] (7.5, 0.85) to[out=5, in=195] (9.4, 1.4);
  \draw[white, line width=1.4pt, opacity=0.5] 
    (1.4, 0.42) -- (3.5, 0.42);

  % -------------------------------------------------------------
  % 4. CHÚ THÍCH (Callout Labels with Pointers)
  % -------------------------------------------------------------

  % 1. "Ống chân không"
  \node[label_txt] (lbl_ong) at (3.2, 2.3) {Ống chân không};
  \draw[pointer] (lbl_ong.south) -- (3.8, 0.85);

  % 2. "Cathode (cực âm)"
  \node[label_txt] (lbl_cat) at (0.8, 1.45) {Cathode (cực âm)};
  \draw[pointer] (lbl_cat.south) -- (1.1, 0.55);

  % 3. "Tia âm cực"
  \node[label_txt] (lbl_tia) at (6.8, 1.55) {Tia âm cực};
  \draw[pointer] (lbl_tia.south) -- (6.2, 0.05);

  % 4. "Lỗ"
  \node[label_txt] (lbl_lo) at (4.8, -0.65) {Lỗ};
  \draw[pointer] (lbl_lo.west) -- (3.35, -0.05);

  % 5. "Anode (cực dương)"
  \node[label_txt] (lbl_ano) at (4.9, -1.45) {Anode (cực dương)};
  \draw[pointer] (lbl_ano.west) -- (3.3, -0.35);

  % 6. "Nguồn"
  \node[label_txt] (lbl_nguon) at (3.1, -2.65) {Nguồn};
  \draw[pointer] (lbl_nguon.west) -- (2.4, -2.65);

  % 7. "Màn phosphor"
  \node[label_txt] (lbl_man) at (8.7, -1.85) {Màn phosphor};
  \draw[pointer] (lbl_man.north) -- (9.4, -0.4);

\end{tikzpicture}
\end{document}
"""

tex_file = os.path.join(SCRATCH_DIR, "ong_phong_dien_tia_am_cuc.tex")
pdf_file = os.path.join(SCRATCH_DIR, "ong_phong_dien_tia_am_cuc.pdf")
png_file = os.path.join(SCRATCH_DIR, "ong_phong_dien_tia_am_cuc.png")
svg_file = os.path.join(SCRATCH_DIR, "ong_phong_dien_tia_am_cuc.svg")

with open(tex_file, "w", encoding="utf-8") as f:
    f.write(tex_cathode)

res = subprocess.run(["pdflatex", "-interaction=nonstopmode", "ong_phong_dien_tia_am_cuc.tex"], cwd=SCRATCH_DIR, capture_output=True, text=True)
if res.returncode == 0:
    doc = pymupdf.open(pdf_file)
    pix = doc[0].get_pixmap(dpi=400)
    pix.save(png_file)
    print("Rendered PNG at 400 DPI!")

    try:
        subprocess.run(["dvisvgm", "--pdf", f"--output={svg_file}", pdf_file], cwd=SCRATCH_DIR, capture_output=True)
    except Exception:
        pass

    for ext in [".tex", ".pdf", ".png", ".svg"]:
        s = os.path.join(SCRATCH_DIR, f"ong_phong_dien_tia_am_cuc{ext}")
        if os.path.exists(s):
            shutil.copyfile(s, os.path.join(SAN_PHAM_DIR, f"ong_phong_dien_tia_am_cuc{ext}"))
            shutil.copyfile(s, os.path.join(ARTIFACT_DIR, f"ong_phong_dien_tia_am_cuc{ext}"))

    shutil.copyfile(png_file, os.path.join(CROP_DIR, "cau13_clean.png"))
    print("Done and updated!")
