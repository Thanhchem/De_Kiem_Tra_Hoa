import os
import subprocess

def main():
    out_dir = r"c:\Antigravity_Thanh\San_Pham\images_de324"
    os.makedirs(out_dir, exist_ok=True)

    # 1. Vẽ Menthol bằng TikZ
    menthol_tikz = r'''\documentclass[border=4mm]{standalone}
\usepackage{tikz}
\usepackage{amsmath}
\usepackage{mathptmx}
\usetikzlibrary{calc}

\begin{document}
\begin{tikzpicture}[line cap=round, line join=round, line width=1.1pt, font=\fontsize{12}{14}\selectfont]
    \def\r{0.9}
    \coordinate (c1) at (90:\r);
    \coordinate (c2) at (30:\r);
    \coordinate (c3) at (330:\r);
    \coordinate (c4) at (270:\r);
    \coordinate (c5) at (210:\r);
    \coordinate (c6) at (150:\r);
    
    % Vòng 6 cạnh
    \draw (c1) -- (c2) -- (c3) -- (c4) -- (c5) -- (c6) -- cycle;
    
    % Nhánh methyl ở đỉnh trên
    \draw (c1) -- ($(c1)+(0,0.65)$);
    
    % Nhánh OH ở cạnh phải
    \draw (c2) -- ($(c2)+(0.55,-0.15)$) node[right=-2pt] {\textbf{OH}};
    
    % Nhánh isopropyl ở đỉnh dưới
    \draw (c4) -- ($(c4)+(0,-0.55)$) coordinate (iso);
    \draw (iso) -- ($(iso)+(-0.45,-0.45)$);
    \draw (iso) -- ($(iso)+(0.45,-0.45)$);
\end{tikzpicture}
\end{document}
'''
    with open(os.path.join(out_dir, "menthol.tex"), "w", encoding="utf-8") as f:
        f.write(menthol_tikz)

    # 2. Vẽ 4 chất ở Câu 5 Phần III bằng TikZ
    cau5_tikz = r'''\documentclass[border=4mm]{standalone}
\usepackage{tikz}
\usepackage{amsmath}
\usepackage{mathptmx}
\usetikzlibrary{calc}

\begin{document}
\begin{tikzpicture}[line cap=round, line join=round, line width=1.0pt, font=\fontsize{11}{13}\selectfont]
    % (1) H2N-CH2-CH=CH-CH3
    \begin{scope}[shift={(0,0)}]
        \node (N1) at (0,0) {$\mathbf{H_2N}$};
        \draw (N1.east) -- ++(30:0.6) -- ++(-30:0.6) coordinate (m1);
        % Nối đôi
        \draw (m1) -- ++(30:0.6) coordinate (m2);
        \draw ($(m1)+(0,0.08)$) -- ($(m2)+(0,0.08)$);
        \draw (m2) -- ++(-30:0.6);
        \node at (1.5,-0.9) {(1)};
    \end{scope}

    % (2) CH3-CH=CH-CH(OH)-CH3
    \begin{scope}[shift={(3.8,0)}]
        \coordinate (p0) at (0,-0.1);
        \draw (p0) -- ++(30:0.6) coordinate (p1);
        % Nối đôi
        \draw (p1) -- ++(-30:0.6) coordinate (p2);
        \draw ($(p1)+(0,0.08)$) -- ($(p2)+(0,0.08)$);
        % Nối tiếp
        \draw (p2) -- ++(30:0.6) coordinate (p3);
        \draw (p3) -- ++(-30:0.6);
        % Nhóm OH ở p3
        \draw (p3) -- ++(0,-0.6) node[below=-2pt] {$\mathbf{OH}$};
        \node at (1.2,-0.9) {(2)};
    \end{scope}

    % (3) CH3-C(CH3)=CH-CH2-COOH
    \begin{scope}[shift={(7.8,0)}]
        \coordinate (q0) at (0,-0.1);
        \draw (q0) -- ++(30:0.6) coordinate (q1);
        % Nhánh methyl hướng lên
        \draw (q1) -- ++(0,0.6);
        % Nối đôi
        \draw (q1) -- ++(-30:0.6) coordinate (q2);
        \draw ($(q1)+(0,-0.08)$) -- ($(q2)+(0,-0.08)$);
        % Nối tiếp CH2-COOH
        \draw (q2) -- ++(30:0.6) coordinate (q3);
        \draw (q3) -- ++(-30:0.6) node[right=-2pt] {$\mathbf{COOH}$};
        \node at (1.2,-0.9) {(3)};
    \end{scope}

    % (4) CH3-C(CH3)=CH-CH=O
    \begin{scope}[shift={(12.2,0)}]
        \coordinate (w0) at (0,-0.1);
        \draw (w0) -- ++(30:0.6) coordinate (w1);
        % Nhánh methyl hướng lên
        \draw (w1) -- ++(0,0.6);
        % Nối đôi
        \draw (w1) -- ++(-30:0.6) coordinate (w2);
        \draw ($(w1)+(0,-0.08)$) -- ($(w2)+(0,-0.08)$);
        % Nối tiếp CH=O
        \draw (w2) -- ++(30:0.6) node[right=-2pt] {$\mathbf{CH=O}$};
        \node at (1.0,-0.9) {(4)};
    \end{scope}
\end{tikzpicture}
\end{document}
'''
    with open(os.path.join(out_dir, "cau5_4chat.tex"), "w", encoding="utf-8") as f:
        f.write(cau5_tikz)

    # Biên dịch sang PDF rồi render PNG 300 DPI
    for fname in ["menthol", "cau5_4chat"]:
        subprocess.run(["pdflatex", "-interaction=nonstopmode", f"{fname}.tex"], cwd=out_dir, check=True)
        subprocess.run(["pdftoppm", "-png", "-r", "300", f"{fname}.pdf", f"{fname}_raw"], cwd=out_dir, check=True)
        raw_png = os.path.join(out_dir, f"{fname}_raw-1.png")
        target_png = os.path.join(out_dir, f"{fname}.png")
        if os.path.exists(raw_png):
            if os.path.exists(target_png):
                os.remove(target_png)
            os.rename(raw_png, target_png)
            print(f"[OK] Rendered {target_png}")

if __name__ == "__main__":
    main()
