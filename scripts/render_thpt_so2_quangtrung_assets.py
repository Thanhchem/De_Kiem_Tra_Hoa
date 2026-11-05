# -*- coding: utf-8 -*-
"""
Script sinh hình ảnh cho Đề kiểm tra Giữa kì 1 Hóa 10 - THPT Số 2 Quang Trung (Gia Lai) - Mã đề 1343
Sử dụng TikZ + pdflatex + pymupdf để tạo ảnh PNG sắc nét (300 DPI)
"""

import os
import sys

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import subprocess
import pymupdf

IMG_DIR = r"C:\Antigravity_Thanh\San_Pham\images_thpt_so2_quangtrung"
SCRATCH_DIR = r"C:\Antigravity_Thanh\scratch"

os.makedirs(IMG_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

def render_tikz_to_png(tikz_code, output_png_path, dpi=300):
    tex_content = r"""\documentclass[tikz,border=8pt]{standalone}
\usepackage[utf8]{vietnam}
\usepackage{amsmath,amssymb}
\usepackage{tgheros}
\usepackage{sansmath}
\sansmath
\renewcommand{\familydefault}{\sfdefault}
\usetikzlibrary{arrows.meta,decorations.pathmorphing,shapes.geometric,calc}

\begin{document}
""" + tikz_code + r"""
\end{document}
"""
    base_name = os.path.splitext(os.path.basename(output_png_path))[0]
    tex_path = os.path.join(SCRATCH_DIR, f"{base_name}.tex")
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_content)
    
    cmd = ["pdflatex", "-interaction=nonstopmode", f"-output-directory={SCRATCH_DIR}", tex_path]
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=SCRATCH_DIR)
    
    pdf_path = os.path.join(SCRATCH_DIR, f"{base_name}.pdf")
    if os.path.exists(pdf_path):
        doc = pymupdf.open(pdf_path)
        page = doc[0]
        pix = page.get_pixmap(dpi=dpi)
        pix.save(output_png_path)
        doc.close()
        print(f"[OK] Đã xuất ảnh: {output_png_path}")
    else:
        print(f"[LỖI] Không tạo được PDF cho {base_name}")

# 1. Câu 1: Sơ đồ 4 lớp electron
tikz_cau1 = r"""
\begin{tikzpicture}[line cap=round, line join=round, scale=1.0]
  % Hạt nhân
  \fill[black] (0,0) circle (3.5pt);
  
  % 4 lớp electron đồng tâm
  \draw[line width=1.0pt] (0,0) circle (0.65cm);
  \draw[line width=1.0pt] (0,0) circle (1.15cm);
  \draw[line width=1.0pt] (0,0) circle (1.65cm);
  \draw[line width=1.0pt] (0,0) circle (2.15cm);
  
  % Đường chỉ dẫn tới các lớp
  \draw[line width=0.8pt] (35:2.15) -- (3.5, 1.25) node[right, font=\bfseries\small] {N ($n = 4$)};
  \draw[line width=0.8pt] (30:1.65) -- (3.5, 0.8) node[right, font=\bfseries\small] {M ($n = 3$)};
  \draw[line width=0.8pt] (25:1.15) -- (3.5, 0.35) node[right, font=\bfseries\small] {L ($n = 2$)};
  \draw[line width=0.8pt] (20:0.65) -- (3.5, -0.1) node[right, font=\bfseries\small] {K ($n = 1$)};
  
  % Đường chỉ dẫn tới Hạt nhân
  \draw[line width=0.8pt] (0,0) -- (3.5, -0.9) node[right, font=\bfseries\small] {Hạt nhân};
\end{tikzpicture}
"""

# 2. Câu 6: Biểu đồ độ âm điện
tikz_cau6 = r"""
\begin{tikzpicture}[line cap=round, line join=round, scale=1.0]
  % Hệ trục tọa độ
  \draw[line width=1.1pt, -{Stealth[scale=1.1]}] (0,0) -- (6.0, 0) node[below=8pt, midway, font=\bfseries] {Nguyên tố};
  \draw[line width=1.1pt, -{Stealth[scale=1.1]}] (0,0) -- (0, 4.3) node[above, font=\bfseries] {Độ âm điện};
  
  % Trục tung chia vạch
  \foreach \y/\lbl in {0.6/0,5, 1.2/1,0, 1.8/1,5, 2.4/2,0, 3.0/2,5, 3.6/3,0, 4.2/3,5} {
    \draw[line width=0.7pt] (-0.12, \y) -- (0, \y);
    \node[left, font=\footnotesize] at (-0.15, \y) {\lbl};
    \draw[dashed, gray!40, line width=0.5pt] (0, \y) -- (5.6, \y);
  }
  \node[left, font=\footnotesize] at (-0.15, 0) {0};
  
  % Các cột độ âm điện: 1 unit độ âm điện = 1.2 cm
  % X (0.93 -> 1.116 cm)
  \draw[line width=1.0pt, fill=black!75] (0.8, 0) rectangle (1.4, 0.93*1.2);
  \node[above, font=\bfseries\footnotesize] at (1.1, 0.93*1.2) {0,93};
  \node[below=3pt, font=\bfseries\small] at (1.1, 0) {X};
  
  % Y (1.61 -> 1.932 cm)
  \draw[line width=1.0pt, fill=black!75] (2.0, 0) rectangle (2.6, 1.61*1.2);
  \node[above, font=\bfseries\footnotesize] at (2.3, 1.61*1.2) {1,61};
  \node[below=3pt, font=\bfseries\small] at (2.3, 0) {Y};
  
  % Z (2.19 -> 2.628 cm)
  \draw[line width=1.0pt, fill=black!75] (3.2, 0) rectangle (3.8, 2.19*1.2);
  \node[above, font=\bfseries\footnotesize] at (3.5, 2.19*1.2) {2,19};
  \node[below=3pt, font=\bfseries\small] at (3.5, 0) {Z};
  
  % T (3.16 -> 3.792 cm)
  \draw[line width=1.0pt, fill=black!75] (4.4, 0) rectangle (5.0, 3.16*1.2);
  \node[above, font=\bfseries\footnotesize] at (4.7, 3.16*1.2) {3,16};
  \node[below=3pt, font=\bfseries\small] at (4.7, 0) {T};
\end{tikzpicture}
"""

# 3. Câu 9: Hình dạng orbital p (px, py, pz)
tikz_cau9 = r"""
\begin{tikzpicture}[line cap=round, line join=round, scale=1.0]
  % 1. Orbital px (dọc theo trục x)
  \begin{scope}[shift={(0,0)}]
    \draw[line width=0.8pt, -{Stealth[scale=0.9]}] (-1.6,0) -- (1.6,0) node[right, font=\footnotesize] {$x$};
    \draw[line width=0.8pt, -{Stealth[scale=0.9]}] (0,-1.4) -- (0,1.4) node[above, font=\footnotesize] {$z$};
    \draw[line width=0.8pt, -{Stealth[scale=0.9]}] (-0.8,-0.6) -- (0.8,0.6) node[above right, font=\footnotesize] {$y$};
    
    % Hai thùy dọc theo trục x
    \fill[black!70, rotate=0] (0.6, 0) ellipse (0.55cm and 0.22cm);
    \fill[black!70, rotate=0] (-0.6, 0) ellipse (0.55cm and 0.22cm);
    \draw[line width=0.8pt] (0.6, 0) ellipse (0.55cm and 0.22cm);
    \draw[line width=0.8pt] (-0.6, 0) ellipse (0.55cm and 0.22cm);
    \node[below=1.4cm, font=\bfseries\small] at (0,0) {$p_x$};
  \end{scope}

  % 2. Orbital py (dọc theo trục y)
  \begin{scope}[shift={(3.8,0)}]
    \draw[line width=0.8pt, -{Stealth[scale=0.9]}] (-1.6,0) -- (1.6,0) node[right, font=\footnotesize] {$x$};
    \draw[line width=0.8pt, -{Stealth[scale=0.9]}] (0,-1.4) -- (0,1.4) node[above, font=\footnotesize] {$z$};
    \draw[line width=0.8pt, -{Stealth[scale=0.9]}] (-1.0,-0.75) -- (1.0,0.75) node[above right, font=\footnotesize] {$y$};
    
    % Hai thùy dọc theo trục y (góc khoảng 37 độ)
    \fill[black!70, rotate=37] (0.6, 0) ellipse (0.55cm and 0.22cm);
    \fill[black!70, rotate=37] (-0.6, 0) ellipse (0.55cm and 0.22cm);
    \draw[line width=0.8pt, rotate=37] (0.6, 0) ellipse (0.55cm and 0.22cm);
    \draw[line width=0.8pt, rotate=37] (-0.6, 0) ellipse (0.55cm and 0.22cm);
    \node[below=1.4cm, font=\bfseries\small] at (0,0) {$p_y$};
  \end{scope}

  % 3. Orbital pz (dọc theo trục z)
  \begin{scope}[shift={(7.6,0)}]
    \draw[line width=0.8pt, -{Stealth[scale=0.9]}] (-1.6,0) -- (1.6,0) node[right, font=\footnotesize] {$x$};
    \draw[line width=0.8pt, -{Stealth[scale=0.9]}] (0,-1.4) -- (0,1.4) node[above, font=\footnotesize] {$z$};
    \draw[line width=0.8pt, -{Stealth[scale=0.9]}] (-0.8,-0.6) -- (0.8,0.6) node[above right, font=\footnotesize] {$y$};
    
    % Hai thùy dọc theo trục z (thẳng đứng)
    \fill[black!70, rotate=90] (0.6, 0) ellipse (0.55cm and 0.22cm);
    \fill[black!70, rotate=90] (-0.6, 0) ellipse (0.55cm and 0.22cm);
    \draw[line width=0.8pt, rotate=90] (0.6, 0) ellipse (0.55cm and 0.22cm);
    \draw[line width=0.8pt, rotate=90] (-0.6, 0) ellipse (0.55cm and 0.22cm);
    \node[below=1.4cm, font=\bfseries\small] at (0,0) {$p_z$};
  \end{scope}
\end{tikzpicture}
"""

# 4. Câu 12: Phổ khối lượng Magnesium
tikz_cau12 = r"""
\begin{tikzpicture}[line cap=round, line join=round, scale=1.0]
  % Hệ trục
  \draw[line width=1.1pt, -{Stealth[scale=1.1]}] (0,0) -- (6.5, 0) node[below=8pt, midway, font=\bfseries] {$m/z$};
  \draw[line width=1.1pt, -{Stealth[scale=1.1]}] (0,0) -- (0, 4.5) node[above, font=\bfseries, align=center] {Tỉ lệ \%\\số nguyên tử};
  
  % Trục tung chia vạch (1 unit = 0.045 cm => 80% = 3.6 cm)
  \foreach \y in {20, 40, 60, 80} {
    \draw[line width=0.7pt] (-0.12, \y*0.045) -- (0, \y*0.045);
    \node[left, font=\footnotesize] at (-0.15, \y*0.045) {\y};
    \draw[dashed, gray!30, line width=0.5pt] (0, \y*0.045) -- (6.0, \y*0.045);
  }
  
  % Các vạch phổ (peak)
  % m/z = 12: cao 78.99% => 3.555 cm
  \draw[line width=2.2pt, color=black] (2.0, 0) -- (2.0, 78.99*0.045);
  \node[above, font=\bfseries\footnotesize] at (2.0, 78.99*0.045) {$^{24}\text{Mg}\ (78,99\%)$};
  \node[below=3pt, font=\bfseries\footnotesize] at (2.0, 0) {12};
  \draw[line width=0.7pt] (2.0, 0) -- (2.0, -0.1);
  
  % m/z = 12.5: cao 10% => 0.45 cm
  \draw[line width=2.2pt, color=black] (3.6, 0) -- (3.6, 10.0*0.045);
  \node[above, font=\bfseries\footnotesize] at (3.5, 10.0*0.045) {$^{25}\text{Mg}\ (10\%)$};
  \node[below=3pt, font=\bfseries\footnotesize] at (3.6, 0) {12,5};
  \draw[line width=0.7pt] (3.6, 0) -- (3.6, -0.1);
  
  % m/z = 13: cao 11.01% => 0.495 cm
  \draw[line width=2.2pt, color=black] (5.2, 0) -- (5.2, 11.01*0.045);
  \node[above, font=\bfseries\footnotesize] at (5.3, 11.01*0.045) {$^{26}\text{Mg}\ (11,01\%)$};
  \node[below=3pt, font=\bfseries\footnotesize] at (5.2, 0) {13};
  \draw[line width=0.7pt] (5.2, 0) -- (5.2, -0.1);
\end{tikzpicture}
"""

# 5. Câu 13: Ô orbital của Ion M3+
tikz_cau13 = r"""
\begin{tikzpicture}[line cap=round, line join=round, scale=0.75]
  \def\s{0.8}
  \tikzset{
    box/.style={line width=1.1pt, draw=black},
    spin_up/.style={line width=1.1pt, -{Stealth[length=4pt, width=3pt]}, draw=black},
    spin_down/.style={line width=1.1pt, -{Stealth[length=4pt, width=3pt]}, draw=black}
  }

  % 14 ô orbital: 9 ô cặp e + 5 ô 1e
  \foreach \i in {0,...,13} {
    \draw[box] (\i*\s, 0) rectangle (\i*\s + \s, \s);
  }

  % 9 ô đầu tiên: cặp e ngược chiều [↑↓]
  \foreach \i in {0,...,8} {
    \draw[spin_up] (\i*\s + 0.3*\s, 0.12*\s) -- (\i*\s + 0.3*\s, 0.88*\s);
    \draw[spin_down] (\i*\s + 0.7*\s, 0.88*\s) -- (\i*\s + 0.7*\s, 0.12*\s);
  }

  % 5 ô tiếp theo: 1 electron quay lên [↑]
  \foreach \i in {9,...,13} {
    \draw[spin_up] (\i*\s + 0.5*\s, 0.12*\s) -- (\i*\s + 0.5*\s, 0.88*\s);
  }
\end{tikzpicture}
"""

# 6. Phần II - Câu 1: Mô phỏng nguyên tử Lithium
tikz_lithium = r"""
\begin{tikzpicture}[line cap=round, line join=round, scale=1.0]
  % Hạt nhân
  \draw[dashed, line width=0.6pt, fill=gray!15] (0,0) circle (0.75cm);
  
  % Các hạt trong hạt nhân (3 proton đỏ +, 4 neutron xám)
  % Proton (đỏ)
  \filldraw[fill=red!80!black, draw=black, line width=0.6pt] (-0.22, 0.22) circle (0.16cm) node[white, font=\bfseries\tiny] {+};
  \filldraw[fill=red!80!black, draw=black, line width=0.6pt] (0.26, 0.18) circle (0.16cm) node[white, font=\bfseries\tiny] {+};
  \filldraw[fill=red!80!black, draw=black, line width=0.6pt] (0.0, -0.32) circle (0.16cm) node[white, font=\bfseries\tiny] {+};
  
  % Neutron (xám)
  \filldraw[fill=gray!60, draw=black, line width=0.6pt] (0.05, 0.05) circle (0.16cm);
  \filldraw[fill=gray!60, draw=black, line width=0.6pt] (-0.32, -0.15) circle (0.16cm);
  \filldraw[fill=gray!60, draw=black, line width=0.6pt] (0.35, -0.15) circle (0.16cm);
  \filldraw[fill=gray!60, draw=black, line width=0.6pt] (-0.05, 0.42) circle (0.16cm);

  % 2 lớp electron
  \draw[dashed, line width=0.8pt] (0,0) circle (1.55cm);
  \draw[dashed, line width=0.8pt] (0,0) circle (2.45cm);

  % Electron lớp trong (lớp K): 2 electron (xanh dương -)
  \filldraw[fill=blue!80!black, draw=black, line width=0.6pt] (0, 1.55) circle (0.14cm) node[white, font=\bfseries\tiny] {-};
  \filldraw[fill=blue!80!black, draw=black, line width=0.6pt] (-1.55, 0) circle (0.14cm) node[white, font=\bfseries\tiny] {-};

  % Electron lớp ngoài (lớp L): 1 electron (xanh dương -)
  \filldraw[fill=blue!80!black, draw=black, line width=0.6pt] (45:2.45) circle (0.14cm) node[white, font=\bfseries\tiny] {-};

  % Chú thích trên hình
  \draw[line width=0.7pt, -{Stealth[scale=0.8]}] (-1.6, 2.2) node[above, font=\bfseries\footnotesize] {Neutron} -- (-0.2, 0.4);
  \draw[line width=0.7pt, -{Stealth[scale=0.8]}] (-2.2, -1.8) node[below, font=\bfseries\footnotesize] {Proton (+)} -- (-0.25, -0.3);
  \draw[line width=0.7pt, -{Stealth[scale=0.8]}] (2.2, 2.3) node[above, font=\bfseries\footnotesize] {Electron (-)} -- (45:2.55);

  % Chú giải bên dưới
  \begin{scope}[shift={(0,-3.2)}]
    \filldraw[fill=red!80!black, draw=black, line width=0.6pt] (-3.2, 0) circle (0.14cm) node[white, font=\bfseries\tiny] {+};
    \node[right, font=\bfseries\footnotesize] at (-3.0, 0) {Proton (+)};
    
    \filldraw[fill=gray!60, draw=black, line width=0.6pt] (-0.5, 0) circle (0.14cm);
    \node[right, font=\bfseries\footnotesize] at (-0.3, 0) {Neutron};
    
    \filldraw[fill=blue!80!black, draw=black, line width=0.6pt] (1.8, 0) circle (0.14cm) node[white, font=\bfseries\tiny] {-};
    \node[right, font=\bfseries\footnotesize] at (2.0, 0) {Electron (-)};
  \end{scope}
\end{tikzpicture}
"""

# 7. Tự luận Câu 1b: Ô orbital của Fluorine
tikz_fluorine = r"""
\begin{tikzpicture}[line cap=round, line join=round, scale=0.8]
  \def\s{0.8}
  \tikzset{
    box/.style={line width=1.1pt, draw=black},
    spin_up/.style={line width=1.1pt, -{Stealth[length=4pt, width=3pt]}, draw=black},
    spin_down/.style={line width=1.1pt, -{Stealth[length=4pt, width=3pt]}, draw=black}
  }

  % Phân lớp 1s
  \draw[box] (0, 0) rectangle (\s, \s);
  \draw[spin_up] (0.3*\s, 0.12*\s) -- (0.3*\s, 0.88*\s);
  \draw[spin_down] (0.7*\s, 0.88*\s) -- (0.7*\s, 0.12*\s);
  \node[below=3pt, font=\bfseries\small] at (0.5*\s, 0) {1s};

  % Phân lớp 2s
  \begin{scope}[shift={(1.8, 0)}]
    \draw[box] (0, 0) rectangle (\s, \s);
    \draw[spin_up] (0.3*\s, 0.12*\s) -- (0.3*\s, 0.88*\s);
    \draw[spin_down] (0.7*\s, 0.88*\s) -- (0.7*\s, 0.12*\s);
    \node[below=3pt, font=\bfseries\small] at (0.5*\s, 0) {2s};
  \end{scope}

  % Phân lớp 2p (3 ô)
  \begin{scope}[shift={(3.6, 0)}]
    \draw[box] (0, 0) rectangle (\s, \s);
    \draw[box] (\s, 0) rectangle (2*\s, \s);
    \draw[box] (2*\s, 0) rectangle (3*\s, \s);
    
    % Ô 1: [↑↓]
    \draw[spin_up] (0.3*\s, 0.12*\s) -- (0.3*\s, 0.88*\s);
    \draw[spin_down] (0.7*\s, 0.88*\s) -- (0.7*\s, 0.12*\s);
    
    % Ô 2: [↑↓]
    \draw[spin_up] (1.3*\s, 0.12*\s) -- (1.3*\s, 0.88*\s);
    \draw[spin_down] (1.7*\s, 0.88*\s) -- (1.7*\s, 0.12*\s);
    
    % Ô 3: [↑ ]
    \draw[spin_up] (2.5*\s, 0.12*\s) -- (2.5*\s, 0.88*\s);
    
    \node[below=3pt, font=\bfseries\small] at (1.5*\s, 0) {2p};
  \end{scope}
\end{tikzpicture}
"""

def main():
    print("Đang tạo các hình vẽ TikZ độ nét cao...")
    render_tikz_to_png(tikz_cau1, os.path.join(IMG_DIR, "hinh_cau1_lop_e.png"))
    render_tikz_to_png(tikz_cau6, os.path.join(IMG_DIR, "hinh_cau6_do_am_dien.png"))
    render_tikz_to_png(tikz_cau9, os.path.join(IMG_DIR, "hinh_cau9_orbital_p.png"))
    render_tikz_to_png(tikz_cau12, os.path.join(IMG_DIR, "hinh_cau12_pho_khoi_mg.png"))
    render_tikz_to_png(tikz_cau13, os.path.join(IMG_DIR, "hinh_cau13_orbital_m3.png"))
    render_tikz_to_png(tikz_lithium, os.path.join(IMG_DIR, "hinh_p2_cau1_lithium.png"))
    render_tikz_to_png(tikz_fluorine, os.path.join(IMG_DIR, "hinh_tl_cau1b_fluorine.png"))
    print("Hoàn tất xuất tất cả hình vẽ sang PNG 300 DPI!")

if __name__ == "__main__":
    main()
