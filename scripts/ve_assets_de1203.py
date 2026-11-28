import os
import subprocess

def main():
    out_dir = r"c:\Antigravity_Thanh\San_Pham\images_de1203"
    os.makedirs(out_dir, exist_ok=True)

    # 1. Vẽ Logo tái chế PVC (03)
    pvc_tex = r'''\documentclass[border=3mm]{standalone}
\usepackage{tikz}
\usepackage{mathptmx}
\usetikzlibrary{arrows.meta,calc}

\begin{document}
\begin{tikzpicture}[line cap=round, line join=round]
    % Tam giác tái chế 3 mũi tên
    % 3 góc ở 90 độ, 210 độ, 330 độ
    \def\R{1.2}
    \coordinate (A) at (90:\R);
    \coordinate (B) at (210:\R);
    \coordinate (C) at (330:\R);
    
    % Vẽ 3 nhánh mũi tên bo góc
    \draw[line width=2.2pt, -{Latex[length=4.5pt, width=4.5pt]}] 
        ($(A)!0.2!(B)$) -- ($(B)!0.8!(A)$);
    \draw[line width=2.2pt, -{Latex[length=4.5pt, width=4.5pt]}] 
        ($(B)!0.2!(C)$) -- ($(C)!0.8!(B)$);
    \draw[line width=2.2pt, -{Latex[length=4.5pt, width=4.5pt]}] 
        ($(C)!0.2!(A)$) -- ($(A)!0.8!(C)$);
        
    % Số 03 ở giữa
    \node at (0, 0.05) {\fontsize{16}{16}\selectfont\textbf{03}};
    % Chữ PVC ở dưới
    \node at (0, -1.25) {\fontsize{12}{14}\selectfont\textbf{PVC}};
\end{tikzpicture}
\end{document}
'''
    with open(os.path.join(out_dir, "pvc_logo.tex"), "w", encoding="utf-8") as f:
        f.write(pvc_tex)

    # 2. Vẽ Sơ đồ điều chế Caprolactam
    capro_tex = r'''\documentclass[border=4mm]{standalone}
\usepackage[utf8]{inputenc}
\usepackage[T5]{fontenc}
\usepackage{mathptmx}
\usepackage{chemfig}
\usepackage{tikz}
\usetikzlibrary{arrows.meta}

\begin{document}
\begin{tikzpicture}[>=stealth, font=\fontsize{11}{13}\selectfont]
    % Node 1: Cyclohexanone
    \node (N1) at (0,0) {
        \chemfig{[:30]*6(-=-=-=)} % dummy or real
    };
    % Thay bằng chemfig thực tế
\end{tikzpicture}
\end{document}
'''
    # Để vẽ sơ đồ caprolactam đẹp nhất, dùng tikz thuần túy rất sắc nét và dễ căn chỉnh
    capro_tikz = r'''\documentclass[border=4mm]{standalone}
\usepackage[utf8]{inputenc}
\usepackage[T5]{fontenc}
\usepackage{mathptmx}
\usepackage{tikz}
\usetikzlibrary{arrows.meta,calc}

\begin{document}
\begin{tikzpicture}[line cap=round, line join=round, font=\fontsize{11}{13}\selectfont]
    \tikzset{
        chemline/.style={line width=1.0pt},
        rxarrow/.style={line width=1.0pt, -{Latex[length=6pt, width=4pt]}}
    }

    % --- 1. Cyclohexanone ---
    \begin{scope}[shift={(0,0)}]
        % 6 cạnh: top là đỉnh (90), các đỉnh khác: 30, 330, 270, 210, 150
        \def\r{0.8}
        \coordinate (c1) at (90:\r);
        \coordinate (c2) at (30:\r);
        \coordinate (c3) at (330:\r);
        \coordinate (c4) at (270:\r);
        \coordinate (c5) at (210:\r);
        \coordinate (c6) at (150:\r);
        
        \draw[chemline] (c1) -- (c2) -- (c3) -- (c4) -- (c5) -- (c6) -- cycle;
        % Liên kết đôi =O ở đỉnh c1
        \draw[chemline] ($(c1)+(-0.06,0)$) -- ($(c1)+(-0.06,0.65)$);
        \draw[chemline] ($(c1)+(0.06,0)$) -- ($(c1)+(0.06,0.65)$);
        \node at ($(c1)+(0,0.85)$) {\textbf{O}};
        \node at (0,-1.25) {\textit{cyclohexanone}};
    \end{scope}

    % --- Mũi tên 1 ---
    \draw[rxarrow] (1.3,0) -- (3.0,0) node[midway, above=2pt] {$\mathrm{NH_2OH}$};

    % --- 2. Cyclohexanoxime ---
    \begin{scope}[shift={(4.3,0)}]
        \def\r{0.8}
        \coordinate (o1) at (90:\r);
        \coordinate (o2) at (30:\r);
        \coordinate (o3) at (330:\r);
        \coordinate (o4) at (270:\r);
        \coordinate (o5) at (210:\r);
        \coordinate (o6) at (150:\r);
        
        \draw[chemline] (o1) -- (o2) -- (o3) -- (o4) -- (o5) -- (o6) -- cycle;
        % Liên kết đôi =N-OH
        \draw[chemline] ($(o1)+(-0.06,0)$) -- ($(o1)+(-0.06,0.5)$);
        \draw[chemline] ($(o1)+(0.06,0)$) -- ($(o1)+(0.06,0.5)$);
        \node (Nnode) at ($(o1)+(0,0.7)$) {\textbf{N}};
        \draw[chemline] (Nnode.north) -- ($(Nnode.north)+(0,0.45)$) node[above=-2pt] {\textbf{OH}};
        \node at (0,-1.25) {\textit{cyclohexanoxime}};
    \end{scope}

    % --- Mũi tên 2 ---
    \draw[rxarrow] (5.6,0) -- (7.3,0) node[midway, above=2pt] {$\mathrm{H_2SO_4}$};

    % --- 3. Caprolactam (vòng 7 cạnh) ---
    \begin{scope}[shift={(8.9,0)}]
        % 7 đỉnh vòng: 
        % 1: C=O (đỉnh trên trái), 2: NH (đỉnh trên phải), 3..7
        \coordinate (v1) at (-0.35, 0.7);  % C=O
        \coordinate (v2) at (0.45, 0.6);   % N
        \coordinate (v3) at (0.85, 0.05);
        \coordinate (v4) at (0.6, -0.65);
        \coordinate (v5) at (-0.1, -0.85);
        \coordinate (v6) at (-0.75, -0.5);
        \coordinate (v7) at (-0.8, 0.2);

        \draw[chemline] (v1) -- ($(v2)+(-0.25, -0.05)$);
        \node (vN) at (v2) {\textbf{N}};
        \node at ($(v2)+(0.12,0.3)$) {\textbf{H}};
        \draw[chemline] ($(vN.south)+(0.05,-0.05)$) -- (v3) -- (v4) -- (v5) -- (v6) -- (v7) -- (v1);

        % Nhóm =O ở v1
        \draw[chemline] ($(v1)+(-0.06,0)$) -- ($(v1)+(-0.06,0.65)$);
        \draw[chemline] ($(v1)+(0.06,0)$) -- ($(v1)+(0.06,0.65)$);
        \node at ($(v1)+(0,0.85)$) {\textbf{O}};
        \node at (0,-1.25) {\textit{caprolactam}};
    \end{scope}

\end{tikzpicture}
\end{document}
'''
    with open(os.path.join(out_dir, "so_do_caprolactam.tex"), "w", encoding="utf-8") as f:
        f.write(capro_tikz)

    # 3. Vẽ phản ứng điều chế Benzyl salicylate cho phần tự luận
    salicylate_tikz = r'''\documentclass[border=4mm]{standalone}
\usepackage[utf8]{inputenc}
\usepackage[T5]{fontenc}
\usepackage{amsmath}
\usepackage{mathptmx}
\usepackage{tikz}
\usetikzlibrary{arrows.meta,calc}

\begin{document}
\begin{tikzpicture}[line cap=round, line join=round, >=stealth, font=\fontsize{11}{13}\selectfont]
    \tikzset{
        chemline/.style={line width=0.9pt},
        rxarrow/.style={line width=0.9pt}
    }

    % --- 1. Salicylic acid ---
    \begin{scope}[shift={(0,0)}]
        % Vòng benzen 6 cạnh
        \def\r{0.75}
        \coordinate (b1) at (90:\r);
        \coordinate (b2) at (30:\r);
        \coordinate (b3) at (330:\r);
        \coordinate (b4) at (270:\r);
        \coordinate (b5) at (210:\r);
        \coordinate (b6) at (150:\r);
        \draw[chemline] (b1) -- (b2) -- (b3) -- (b4) -- (b5) -- (b6) -- cycle;
        \draw[chemline] (0,0) circle (0.45);
        % Nhánh COOH ở b1
        \draw[chemline] (b1) -- ($(b1)+(0,0.45)$) node[above=-2pt] {$\mathrm{COOH}$};
        % Nhánh OH ở b2
        \draw[chemline] (b2) -- ($(b2)+(0.4,0.15)$) node[right=-2pt] {$\mathrm{OH}$};
        \node at (0,-1.15) {\textit{salicylic acid}};
    \end{scope}

    \node at (1.5,0) {\textbf{+}};

    % --- 2. Benzyl alcohol ---
    \begin{scope}[shift={(3.2,0)}]
        \def\r{0.75}
        \coordinate (a1) at (90:\r);
        \coordinate (a2) at (30:\r);
        \coordinate (a3) at (330:\r);
        \coordinate (a4) at (270:\r);
        \coordinate (a5) at (210:\r);
        \coordinate (a6) at (150:\r);
        \draw[chemline] (a1) -- (a2) -- (a3) -- (a4) -- (a5) -- (a6) -- cycle;
        \draw[chemline] (0,0) circle (0.45);
        % Nhánh CH2OH ở a1
        \draw[chemline] (a1) -- ($(a1)+(0,0.45)$) node[above=-2pt] {$\mathrm{CH_2OH}$};
        \node at (0,-1.15) {\textit{benzyl alcohol}};
    \end{scope}

    % --- Mũi tên thuận nghịch ---
    \begin{scope}[shift={(4.9,0)}]
        \draw[rxarrow, -{Latex[length=4pt, width=3pt]}] (-0.4, 0.08) -- (0.5, 0.08) node[midway, above=-1pt] {\scriptsize $\mathrm{H_2SO_4\ \text{đặc},\ t^\circ}$};
        \draw[rxarrow, {Latex[length=4pt, width=3pt]}-] (-0.4, -0.08) -- (0.5, -0.08);
    \end{scope}

    % --- 3. Benzyl salicylate ---
    \begin{scope}[shift={(7.3,0)}]
        \def\r{0.75}
        \coordinate (s1) at (90:\r);
        \coordinate (s2) at (30:\r);
        \coordinate (s3) at (330:\r);
        \coordinate (s4) at (270:\r);
        \coordinate (s5) at (210:\r);
        \coordinate (s6) at (150:\r);
        \draw[chemline] (s1) -- (s2) -- (s3) -- (s4) -- (s5) -- (s6) -- cycle;
        \draw[chemline] (0,0) circle (0.45);
        \draw[chemline] (s2) -- ($(s2)+(0.4,0.15)$) node[right=-2pt] {$\mathrm{OH}$};

        % Cầu nối COOCH2
        \coordinate (Ctop) at ($(s1)+(0,0.45)$);
        \draw[chemline] (s1) -- (Ctop) node[above=-2pt] (Cnode) {$\mathrm{COOCH_2}$};

        % Vòng benzyl thứ hai
        \begin{scope}[shift={(8.9, 1.1)}]
            \def\rB{0.65}
            \coordinate (k1) at (210:\rB);
            \coordinate (k2) at (150:\rB);
            \coordinate (k3) at (90:\rB);
            \coordinate (k4) at (30:\rB);
            \coordinate (k5) at (330:\rB);
            \coordinate (k6) at (270:\rB);
            \draw[chemline] (k1) -- (k2) -- (k3) -- (k4) -- (k5) -- (k6) -- cycle;
            \draw[chemline] (0,0) circle (0.38);
            \draw[chemline] (Cnode.east) -- (k1);
        \end{scope}
        
        \node at (0.8,-1.15) {\textit{benzyl salicylate}};
    \end{scope}

    \node at (10.6,0) {\textbf{+} $\mathbf{H_2O}$};

\end{tikzpicture}
\end{document}
'''
    with open(os.path.join(out_dir, "benzyl_salicylate.tex"), "w", encoding="utf-8") as f:
        f.write(salicylate_tikz)

    # Biên dịch cả 3 file tex sang pdf rồi render png bằng pdftoppm
    files = ["pvc_logo", "so_do_caprolactam", "benzyl_salicylate"]
    for fname in files:
        tex_file = os.path.join(out_dir, f"{fname}.tex")
        subprocess.run(["pdflatex", "-interaction=nonstopmode", f"{fname}.tex"], cwd=out_dir, check=True)
        # pdftoppm -png -r 300
        pdf_file = os.path.join(out_dir, f"{fname}.pdf")
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
