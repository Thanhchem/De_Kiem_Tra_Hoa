import os
import subprocess
import fitz # PyMuPDF

tex_template = r"""\documentclass{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese,english]{babel}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{tikz}
\usetikzlibrary{shapes.geometric, arrows.meta, positioning, calc, backgrounds}

\usepackage[
  paperwidth=21.0cm,
  paperheight=16.0cm,
  margin=0.4cm,
  footskip=0pt,
  headheight=0pt,
  headsep=0pt
]{geometry}

\pagestyle{empty}

\definecolor{arrowRed}{RGB}{185, 28, 28}
\definecolor{badgeGreen}{RGB}{124, 179, 66}

\begin{document}
\selectlanguage{vietnamese}

\begin{center}
\begin{tikzpicture}
  % Canvas boundaries
  \useasboundingbox (0, 0) rectangle (20.2, 15.2);
  
  % ---------------- 1. TẾ BÀO (CELL) ----------------
  % Circular Image
  \node at (4.5, 11.0) {
    \includegraphics[width=3.2cm,height=3.2cm]{San_Pham/hd_cell.png}
  };
  \draw[arrowRed, line width=2.5pt] (4.5, 11.0) circle (1.6cm);
  % Label
  \node[above] at (4.5, 12.8) {{\fontsize{14}{17}\selectfont \textbf{Tế bào}}};

  % ---------------- 2. MÔ (TISSUE) ----------------
  % Circular Image
  \node at (15.5, 11.2) {
    \includegraphics[width=3.2cm,height=3.2cm]{San_Pham/hd_tissue.png}
  };
  \draw[arrowRed, line width=2.5pt] (15.5, 11.2) circle (1.6cm);
  % Label
  \node[above] at (17.2, 12.8) {{\fontsize{14}{17}\selectfont \textbf{Mô}}};

  % ---------------- 3. CƠ QUAN (STOMACH) ----------------
  % Circular Image
  \node at (15.5, 4.3) {
    \includegraphics[width=3.2cm,height=3.2cm]{San_Pham/hd_stomach.png}
  };
  \draw[arrowRed, line width=2.5pt] (15.5, 4.3) circle (1.6cm);
  % Label
  \node[align=center, below] at (15.5, 2.5) {
    {\fontsize{13}{16}\selectfont \textbf{Cơ quan}}\\[2pt]
    {\fontsize{11}{14}\selectfont \textit{(Ví dụ: dạ dày)}}
  };

  % ---------------- 4. HỆ CƠ QUAN (DIGESTIVE TORSO) ----------------
  \node at (9.6, 5.6) {
    \includegraphics[width=4.5cm]{San_Pham/hd_torso.png}
  };
  % Label
  \node[align=center, below] at (6.6, 2.5) {
    {\fontsize{13}{16}\selectfont \textbf{Hệ cơ quan}}\\[2pt]
    {\fontsize{11}{14}\selectfont \textit{(Ví dụ: hệ tiêu hoá)}}
  };

  % ---------------- 5. CƠ THỂ (HUMAN BODY) ----------------
  \node at (2.2, 7.3) {
    \includegraphics[width=2.85cm]{San_Pham/hd_body.png}
  };
  % Label
  \node[below] at (2.2, 4.0) {{\fontsize{13}{16}\selectfont \textbf{Cơ thể}}};

  % ---------------- CONNECTING ARROWS ----------------
  % Arrow 1: Tế bào -> Mô (top curve)
  \draw[-{Stealth[length=4.8mm, width=3.2mm]}, arrowRed, line width=2.3pt]
    (6.3, 12.1) to[out=22, in=165] (13.6, 12.4);

  % Arrow 2: Mô -> Cơ quan (right outward curve)
  \draw[-{Stealth[length=4.8mm, width=3.2mm]}, arrowRed, line width=2.3pt]
    (17.3, 10.4) to[out=-35, in=45] (17.4, 5.5);

  % Arrow 3: Cơ quan -> Hệ cơ quan (bottom-left curve)
  \draw[-{Stealth[length=4.8mm, width=3.2mm]}, arrowRed, line width=2.3pt]
    (13.6, 3.7) to[out=195, in=-15] (11.6, 4.0);

  % Arrow 4: Hệ cơ quan -> Cơ thể (up-left curve)
  \draw[-{Stealth[length=4.8mm, width=3.2mm]}, arrowRed, line width=2.3pt]
    (7.8, 4.7) to[out=155, in=-35] (3.4, 6.2);

  % ---------------- BOTTOM CAPTION ----------------
  \node[anchor=center] at (10.1, 0.7) {
    \tikz[baseline=-0.6ex]{
      \node[fill=badgeGreen, text=white, rounded corners=4pt, inner sep=5pt, font=\fontsize{12}{15}\selectfont\bfseries] {Hình 23.1};
    }\hspace{0.35cm}
    {\fontsize{13}{16}\selectfont \textit{\textbf{Sơ đồ các cấp tổ chức cấu tạo cơ thể người}}}
  };

\end{tikzpicture}
\end{center}

\end{document}
"""

def main():
    tex_path = "San_Pham/Hinh_23_1_So_Do_Cac_Cap_To_Chuc_Vector.tex"
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_template)
    
    cmd = ["pdflatex", "-interaction=nonstopmode", "-output-directory=San_Pham", tex_path]
    res = subprocess.run(cmd, capture_output=True)
    if res.returncode != 0:
        print("pdflatex failed:")
        print(res.stdout.decode('utf-8', errors='ignore')[-1500:])
        return
    print("PDF compiled successfully!")
    
    pdf_path = "San_Pham/Hinh_23_1_So_Do_Cac_Cap_To_Chuc_Vector.pdf"
    doc = fitz.open(pdf_path)
    page = doc[0]
    mat = fitz.Matrix(300/72, 300/72)
    pix = page.get_pixmap(matrix=mat)
    png_path = "San_Pham/Hinh_23_1_So_Do_Cac_Cap_To_Chuc_Vector.png"
    pix.save(png_path)
    print(f"High-res PNG saved: {png_path} ({pix.width}x{pix.height})")

    for ext in [".aux", ".log"]:
        f_rm = "San_Pham/Hinh_23_1_So_Do_Cac_Cap_To_Chuc_Vector" + ext
        if os.path.exists(f_rm):
            os.remove(f_rm)

if __name__ == '__main__':
    main()
