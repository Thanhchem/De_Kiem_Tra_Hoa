import os
import subprocess
import pymupdf

def main():
    output_dir = r"c:\Antigravity_Thanh\San_Pham\So_Do_Thuy_Phan_Salicin"
    os.makedirs(output_dir, exist_ok=True)

    tex_path = os.path.join(output_dir, "so_do_thuy_phan_salicin.tex")
    pdf_path = os.path.join(output_dir, "so_do_thuy_phan_salicin.pdf")
    png_path = os.path.join(output_dir, "so_do_thuy_phan_salicin.png")
    png_hd_path = os.path.join(output_dir, "so_do_thuy_phan_salicin_hd.png")
    png_trans_path = os.path.join(output_dir, "so_do_thuy_phan_salicin_transparent.png")

    tikz_code = r'''\documentclass[border=6mm]{standalone}
\usepackage[utf8]{inputenc}
\usepackage[T5]{fontenc}
\usepackage{mathptmx}
\usepackage{tikz}
\usetikzlibrary{arrows.meta,calc}

\begin{document}
\begin{tikzpicture}[
    line cap=round, line join=round,
    >=Stealth,
    font=\fontsize{12}{14}\selectfont
]

% Định nghĩa nét vẽ chuẩn ChemDraw
\tikzset{
    chemline/.style={line width=0.85pt},
    thickbond/.style={line width=3.0pt},
    rxnarrow/.style={line width=0.85pt, -{Latex[length=5.2pt, width=3.4pt]}}
}

% Macro vẽ vòng benzen chuẩn Kekule với 3 liên kết đôi xen kẽ
% Tâm = #1, Bán kính R = 0.68cm
\newcommand{\drawBenzeneRing}[1]{
    \def\R{0.68}
    \def\xR{0.5889} % R * sqrt(3)/2
    \def\yR{0.34}   % R / 2
    
    % 6 đỉnh: 90, 150, 210, 270, 330, 30
    \coordinate (#1-90)  at ($(#1) + (0, \R)$);
    \coordinate (#1-150) at ($(#1) + (-\xR, \yR)$);
    \coordinate (#1-210) at ($(#1) + (-\xR, -\yR)$);
    \coordinate (#1-270) at ($(#1) + (0, -\R)$);
    \coordinate (#1-330) at ($(#1) + (\xR, -\yR)$);
    \coordinate (#1-30)  at ($(#1) + (\xR, \yR)$);
    
    % Vòng 6 cạnh ngoài
    \draw[chemline] (#1-90) -- (#1-150) -- (#1-210) -- (#1-270) -- (#1-330) -- (#1-30) -- cycle;
    
    % 3 liên kết đôi bên trong song song tuyệt đối:
    % 1. Cạnh đứng phải: từ 330 đến 30
    \draw[chemline] ($($(#1-330)!0.16!(#1-30)$) + (-0.11, 0)$) -- ($($(#1-30)!0.16!(#1-330)$) + (-0.11, 0)$);
    % 2. Cạnh chéo trên trái: từ 90 đến 150
    \draw[chemline] ($($(#1-90)!0.16!(#1-150)$) + (0.055, -0.095)$) -- ($($(#1-150)!0.16!(#1-90)$) + (0.055, -0.095)$);
    % 3. Cạnh chéo dưới trái: từ 210 đến 270
    \draw[chemline] ($($(#1-210)!0.16!(#1-270)$) + (0.055, 0.095)$) -- ($($(#1-270)!0.16!(#1-210)$) + (0.055, 0.095)$);
}

% =============================================================
% DÒNG 1: PHẢN ỨNG THỦY PHÂN SALICIN
% =============================================================

% --- 1. Phân tử Salicin ---
\coordinate (C4) at (0, 0);
\coordinate (C5) at (1.05, 0.82);
\coordinate (RingOPos) at (2.2, 0.82);
\coordinate (C1) at (3.0, 0);
\coordinate (C2) at (2.2, -0.76);
\coordinate (C3) at (1.05, -0.76);

% Vòng pyranose phía sau: C4 -> C5 -> RingO -> C1
\node[inner sep=2pt] (RingO) at (RingOPos) {O};
\draw[chemline] (C4) -- (C5) -- (RingO.west);
\draw[chemline] (RingO.south east) -- (C1);

% Mặt trước: nêm đậm C4-C3, vạch đậm C3-C2, nêm đậm C2-C1
\fill (C4) -- ($(C3) + (0, 0.06)$) -- ($(C3) + (0, -0.06)$) -- cycle;
\draw[thickbond] (C3) -- (C2);
\fill (C1) -- ($(C2) + (0, 0.06)$) -- ($(C2) + (0, -0.06)$) -- cycle;

% Nhóm thế trên vòng pyranose:
% C4: H hướng lên, OH hướng xuống
\draw[chemline] ($(C4) + (0, 0.45)$) node[above, inner sep=1pt] {H} -- ($(C4) + (0, -0.45)$) node[below, inner sep=1pt] {OH};

% C5: CH2OH hướng lên (gắn vào C), vạch ngắn hướng xuống
\draw[chemline] (C5) -- ++(0, 0.45) node[above, anchor=south, xshift=13pt, inner sep=1pt] {CH$_2$OH};
\draw[chemline] (C5) -- ++(0, -0.26);

% C3: OH hướng lên, H hướng xuống
\draw[chemline] (C3) -- ++(0, 0.36) node[above, inner sep=1pt] {OH};
\draw[chemline] (C3) -- ++(0, -0.45) node[below, inner sep=1pt] {H};

% C2: H hướng lên, OH hướng xuống
\draw[chemline] (C2) -- ++(0, 0.36) node[above, inner sep=1pt] {H};
\draw[chemline] (C2) -- ++(0, -0.45) node[below, inner sep=1pt] {OH};

% C1: H hướng xuống, O liên kết glycoside hướng lên
\draw[chemline] (C1) -- ++(0, -0.45) node[below, inner sep=1pt] {H};
\draw[chemline] (C1) -- ++(0, 0.45) node[above, inner sep=1.5pt] (GlycO) {O};

% Vòng benzen của Salicin
\coordinate (BZ1) at (4.15, 1.35);
\drawBenzeneRing{BZ1};

% Liên kết từ GlycO đến vị trí 210 của vòng benzen
\draw[chemline] (GlycO.north east) -- (BZ1-210);

% Nhánh -CH2-OH ở vị trí 150
\coordinate (BZ1-CH2) at ($(BZ1-150) + (-0.38, 0.30)$);
\draw[chemline] (BZ1-150) -- (BZ1-CH2) -- ++(-0.38, -0.22) node[left, inner sep=1pt] {HO};

% --- Mũi tên 1: thủy phân ---
\draw[rxnarrow] (5.4, 0) -- (7.6, 0) node[midway, above=2.5pt] {thủy phân};

% --- Sản phẩm: glucose + salicylic alcohol ---
\node at (8.7, 0) {glucose};
\node at (9.8, 0) {+};

% Phân tử salicylic alcohol (trên)
\coordinate (BZ2) at (11.8, 0);
\drawBenzeneRing{BZ2};
\coordinate (BZ2-CH2) at ($(BZ2-150) + (-0.38, 0.30)$);
\draw[chemline] (BZ2-150) -- (BZ2-CH2) -- ++(-0.38, -0.22) node[left, inner sep=1pt] {HO};
\draw[chemline] (BZ2-210) -- ++(-0.42, -0.25) node[left, inner sep=1pt] {HO};
\node at (11.8, -1.45) {salicylic alcohol};

% =============================================================
% DÒNG 2: PHẢN ỨNG OXI HÓA SALICYLIC ALCOHOL
% =============================================================

% --- Phân tử salicylic alcohol (dưới) ---
\coordinate (BZ3) at (1.8, -3.2);
\drawBenzeneRing{BZ3};
\coordinate (BZ3-CH2) at ($(BZ3-150) + (-0.38, 0.30)$);
\draw[chemline] (BZ3-150) -- (BZ3-CH2) -- ++(-0.38, -0.22) node[left, inner sep=1pt] {HO};
\draw[chemline] (BZ3-210) -- ++(-0.42, -0.25) node[left, inner sep=1pt] {HO};

% --- Mũi tên 2: oxi hóa ---
\draw[rxnarrow] (3.0, -3.2) -- (5.2, -3.2) node[midway, above=2.5pt] {oxi hóa};

% --- Sản phẩm: salicylic acid ---
\coordinate (BZ4) at (6.8, -3.2);
\drawBenzeneRing{BZ4};
\draw[chemline] (BZ4-150) -- ++(-0.46, 0.28) node[left, inner sep=1pt] {HOOC};
\draw[chemline] (BZ4-210) -- ++(-0.42, -0.25) node[left, inner sep=1pt] {HO};

\end{tikzpicture}
\end{document}
'''

    with open(tex_path, 'w', encoding='utf-8') as f:
        f.write(tikz_code)

    res = subprocess.run(['pdflatex', '-interaction=nonstopmode', f'-output-directory={output_dir}', tex_path], capture_output=True, text=True)
    if res.returncode != 0:
        print("Lỗi biên dịch LaTeX:")
        print(res.stdout[-600:])
        return

    doc = pymupdf.open(pdf_path)
    page = doc[0]

    # Standard 300 DPI (nền trắng)
    pix = page.get_pixmap(dpi=300, alpha=False)
    pix.save(png_path)

    # Ultra HD 600 DPI (nền trắng)
    pix_hd = page.get_pixmap(dpi=600, alpha=False)
    pix_hd.save(png_hd_path)

    # Transparent 300 DPI (nền trong suốt)
    pix_trans = page.get_pixmap(dpi=300, alpha=True)
    pix_trans.save(png_trans_path)

    # Dọn dẹp file tạm
    for ext in ['.aux', '.log']:
        f_clean = os.path.join(output_dir, "so_do_thuy_phan_salicin" + ext)
        if os.path.exists(f_clean):
            os.remove(f_clean)
    print("Hoàn tất tạo bộ file sơ đồ thủy phân Salicin chất lượng cao!")

if __name__ == '__main__':
    main()
