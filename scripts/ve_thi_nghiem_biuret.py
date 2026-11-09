# -*- coding: utf-8 -*-
"""
Script vẽ lại chính xác sơ đồ thí nghiệm:
Phản ứng màu Biuret của protein (lòng trắng trứng) với Cu(OH)2
Xuất ra 4 định dạng:
1. thi_nghiem_biuret_long_trang_trung.pdf (Vector)
2. thi_nghiem_biuret_long_trang_trung.png (300 DPI, nền trắng)
3. thi_nghiem_biuret_long_trang_trung_hd.png (600 DPI siêu nét)
4. thi_nghiem_biuret_long_trang_trung_transparent.png (300 DPI, nền trong suốt)
Đồng thời cập nhật vào thư mục đề thi THPT Số 2 Quang Trung.
"""

import os
import subprocess
import pymupdf

def main():
    output_dir = r"c:\Antigravity_Thanh\San_Pham\Thi_Nghiem_Phan_Ung_Mau_Biuret"
    os.makedirs(output_dir, exist_ok=True)

    tex_path = os.path.join(output_dir, "thi_nghiem_biuret_long_trang_trung.tex")
    pdf_path = os.path.join(output_dir, "thi_nghiem_biuret_long_trang_trung.pdf")
    png_path = os.path.join(output_dir, "thi_nghiem_biuret_long_trang_trung.png")
    png_hd_path = os.path.join(output_dir, "thi_nghiem_biuret_long_trang_trung_hd.png")
    png_trans_path = os.path.join(output_dir, "thi_nghiem_biuret_long_trang_trung_transparent.png")

    tikz_code = r'''\documentclass[border=5mm]{standalone}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{mathptmx}
\usepackage{tikz}
\usetikzlibrary{arrows.meta,calc,decorations.pathmorphing,patterns}

\begin{document}
\begin{tikzpicture}[
    line cap=round, line join=round,
    >=Stealth,
    font=\fontsize{11}{13}\selectfont
]

% Màu sắc khoa học thanh lịch
\definecolor{glassLine}{HTML}{2B2B2B}      % Nét viền thủy tinh
\definecolor{liquidLight}{HTML}{E8F0FE}    % Dung dịch trong suốt hơi xanh nhạt
\definecolor{cuPpt}{HTML}{42A5F5}          % Kết tủa Cu(OH)2 màu xanh lam
\definecolor{biuretPurple}{HTML}{E1BEE7}   % Phức biuret tím nhạt
\definecolor{biuretPurpleDeep}{HTML}{AB47BC}
\definecolor{rubberBulb}{HTML}{5D4037}     % Quả bóp cao su
\definecolor{arrowColor}{HTML}{1A237E}     % Mũi tên tiến trình

% Style chung
\tikzset{
    glass/.style={line width=0.85pt, draw=glassLine},
    liquidSurface/.style={line width=0.6pt, draw=glassLine!70},
    leadLine/.style={line width=0.6pt, draw=glassLine!80},
    procArrow/.style={line width=0.9pt, ->, >=Latex, draw=arrowColor}
}

% Macro vẽ ống nghiệm
% #1: Tọa độ tâm đáy (x, y)
% #2: Bán kính r (thường 0.42cm)
% #3: Chiều cao thành ống h (thường 3.5cm)
\newcommand{\drawTestTube}[3]{
    % Thành ống bên trái, đáy bán cầu, thành bên phải
    \draw[glass] 
        ($(#1) + (-#2, #3)$) -- ($(#1) + (-#2, 0)$)
        arc[start angle=180, end angle=360, radius=#2]
        -- ($(#1) + (#2, #3)$);
    
    % Miệng ống nghiệm loe nhẹ (rim/lip)
    \draw[glass] ($(#1) + (-#2-0.05, #3)$) -- ($(#1) + (-#2, #3)$);
    \draw[glass] ($(#1) + (#2+0.05, #3)$) -- ($(#1) + (#2, #3)$);
    \draw[glass] ($(#1) + (0, #3)$) ellipse [x radius=#2+0.05, y radius=0.07];
}

% Macro vẽ ống nhỏ giọt (dropper)
% #1: Điểm đầu nhọn
% #2: Góc nghiêng (độ)
% #3: Chiều dài thân
\newcommand{\drawDropper}[3]{
    \begin{scope}[shift={(#1)}, rotate=#2]
        % Đầu nhọn thủy tinh
        \draw[glass] (-0.035, 0) -- (-0.04, 0.3) -- (-0.12, 0.7) -- (-0.12, #3)
                     (0.035, 0) -- (0.04, 0.3) -- (0.12, 0.7) -- (0.12, #3);
        % Ngấn nối quả bóp
        \draw[glass] (-0.14, #3) rectangle (0.14, #3+0.08);
        % Quả bóp cao su
        \draw[glass, fill=rubberBulb!70] 
            (-0.12, #3+0.08) .. controls (-0.25, #3+0.4) and (-0.22, #3+0.9) .. (0, #3+1.05)
            .. controls (0.22, #3+0.9) and (0.25, #3+0.4) .. (0.12, #3+0.08) -- cycle;
    \end{scope}
}

% ==============================================================================
% ỐNG NGHIỆM (1): Cho NaOH và nhỏ giọt CuSO4
% ==============================================================================
\def\xOne{0}
\def\rTube{0.44}
\def\hTube{3.4}

% Dung dịch trong ống (1): 1 mL NaOH 30%
\fill[fill=liquidLight!80] 
    ($(\xOne, 0) + (-\rTube, 0)$) arc[start angle=180, end angle=360, radius=\rTube]
    -- ($(\xOne, 0.8) + (\rTube, 0)$) 
    arc[start angle=0, end angle=-180, x radius=\rTube, y radius=0.06] -- cycle;
\draw[liquidSurface] (\xOne, 0.8) ellipse [x radius=\rTube, y radius=0.06];

% Vẽ vỏ ống nghiệm (1)
\drawTestTube{\xOne, 0}{\rTube}{\hTube}

% Ống nhỏ giọt CuSO4 nghiêng khoảng 40 độ phía trên
\coordinate (tipOne) at ($(\xOne, \hTube) + (-0.35, 0.95)$);
\drawDropper{tipOne}{-42}{1.3}

% Các giọt dung dịch rơi xuống
\draw[fill=cuPpt, draw=glassLine, line width=0.4pt] 
    ($(\xOne, \hTube) + (-0.15, 0.65)$) circle [radius=0.045];
\draw[fill=cuPpt, draw=glassLine, line width=0.4pt] 
    ($(\xOne, \hTube) + (-0.05, 0.3)$) circle [radius=0.05];

% Chú thích trên ống nhỏ giọt (1)
\draw[leadLine] ($(\xOne, \hTube) + (0.35, 1.3)$) -- ($(\xOne, \hTube) + (0.7, 1.45)$);
\node[right, align=left, inner sep=1pt] at ($(\xOne, \hTube) + (0.7, 1.45)$) {
    \textit{2--3 giọt dung dịch}\\
    \textit{\ce{CuSO4} 2\%}
};

% Chú thích dung dịch NaOH (1)
\draw[leadLine] ($(\xOne, 0.35) + (0.1, 0)$) -- ($(\xOne, 0.35) + (0.9, -0.3)$);
\node[right, align=left, inner sep=1pt] at ($(\xOne, 0.35) + (0.9, -0.3)$) {
    \textit{1 mL dung dịch}\\
    \textit{NaOH 30\%}
};

% Nhãn ống (1)
\node[below=4pt] at (\xOne, -\rTube) {\textit{\textbf{(1)}}};

% ==============================================================================
% MŨI TÊN (1) -> (2): Lắc đều
% ==============================================================================
\def\xTwo{5.0}
\draw[procArrow] ($(\xOne+\rTube+0.45, 1.6)$) -- ($(\xTwo-\rTube-0.45, 1.6)$)
    node[midway, above=2pt] {\textit{\textbf{Lắc đều}}};

% ==============================================================================
% ỐNG NGHIỆM (2): Kết tủa Cu(OH)2 màu xanh
% ==============================================================================
% Dung dịch và kết tủa
\fill[fill=liquidLight!80] 
    ($(\xTwo, 0) + (-\rTube, 0)$) arc[start angle=180, end angle=360, radius=\rTube]
    -- ($(\xTwo, 0.85) + (\rTube, 0)$) 
    arc[start angle=0, end angle=-180, x radius=\rTube, y radius=0.06] -- cycle;
\draw[liquidSurface] (\xTwo, 0.85) ellipse [x radius=\rTube, y radius=0.06];

% Kết tủa dạng vẩn hạt màu xanh lam đặc trưng
\begin{scope}
    \clip ($(\xTwo, 0) + (-\rTube, 0)$) arc[start angle=180, end angle=360, radius=\rTube]
          -- ($(\xTwo, 0.85) + (\rTube, 0)$) -- ($(\xTwo, 0.85) + (-\rTube, 0)$) -- cycle;
    % Đám kết tủa Cu(OH)2 lơ lửng & tụ dưới đáy
    \foreach \i in {1,...,38} {
        \fill[cuPpt!80!black] 
            ($(\xTwo-0.34, -0.3) + (rnd*0.68, rnd*1.0)$) circle [radius=0.035+rnd*0.035];
    }
    \fill[cuPpt, opacity=0.35] 
        ($(\xTwo, 0) + (-\rTube, 0)$) arc[start angle=180, end angle=360, radius=\rTube]
        -- ($(\xTwo, 0.6) + (\rTube, 0)$) -- ($(\xTwo, 0.6) + (-\rTube, 0)$) -- cycle;
\end{scope}

% Vẽ vỏ ống nghiệm (2)
\drawTestTube{\xTwo, 0}{\rTube}{\hTube}

% Chú thích kết tủa Cu(OH)2 (2)
\draw[leadLine] ($(\xTwo, 0.35) + (0.1, 0)$) -- ($(\xTwo, 0.35) + (0.9, 0.3)$);
\node[right, align=left, inner sep=1pt] at ($(\xTwo, 0.35) + (0.9, 0.3)$) {
    \textit{\textbf{Kết tủa \ce{Cu(OH)2}}}\\
    \textit{\textbf{màu xanh}}
};

% Nhãn ống (2)
\node[below=4pt] at (\xTwo, -\rTube) {\textit{\textbf{(2)}}};

% ==============================================================================
% MŨI TÊN (2) -> (3): Thêm lòng trắng trứng, lắc đều
% ==============================================================================
\def\xThree{10.2}
\draw[procArrow] ($(\xTwo+\rTube+0.45, 1.8)$) -- ($(\xThree-\rTube-0.45, 1.8)$)
    node[midway, above=2pt, align=center] {
        \textit{\textbf{Thêm lòng trắng trứng,}}\\
        \textit{\textbf{lắc đều}}
    };

% ==============================================================================
% ỐNG NGHIỆM (3): Dung dịch X (phức biuret màu tím)
% ==============================================================================
% Dung dịch dâng cao khoảng 2.0 cm
\fill[fill=biuretPurple!90] 
    ($(\xThree, 0) + (-\rTube, 0)$) arc[start angle=180, end angle=360, radius=\rTube]
    -- ($(\xThree, 2.0) + (\rTube, 0)$) 
    arc[start angle=0, end angle=-180, x radius=\rTube, y radius=0.06] -- cycle;
\draw[liquidSurface, draw=biuretPurpleDeep!80] 
    (\xThree, 2.0) ellipse [x radius=\rTube, y radius=0.06];

% Vẽ vỏ ống nghiệm (3)
\drawTestTube{\xThree, 0}{\rTube}{\hTube}

% Pipette/ống hút nhỏ lòng trắng trứng dài hơn
\coordinate (tipThree) at ($(\xThree, \hTube) + (0.05, 0.4)$);
\begin{scope}[shift={(tipThree)}, rotate=48]
    % Thân pipette thủy tinh dài
    \draw[glass] (-0.035, 0) -- (-0.045, 0.4) -- (-0.11, 0.8) -- (-0.11, 2.4)
                 (0.035, 0) -- (0.045, 0.4) -- (0.11, 0.8) -- (0.11, 2.4);
    % Vạch chia độ nhỏ
    \foreach \y in {1.0, 1.3, 1.6, 1.9, 2.2} {
        \draw[glassLine!60, line width=0.4pt] (-0.08, \y) -- (-0.02, \y);
    }
    % Dung dịch bên trong pipette
    \fill[biuretPurple!40, opacity=0.6] 
        (-0.09, 0.8) rectangle (0.09, 2.1);
    % Quả bóp cao su
    \draw[glass] (-0.13, 2.4) rectangle (0.13, 2.48);
    \draw[glass, fill=rubberBulb!70] 
        (-0.11, 2.48) .. controls (-0.22, 2.7) and (-0.2, 3.1) .. (0, 3.25)
        .. controls (0.2, 3.1) and (0.22, 2.7) .. (0.11, 2.48) -- cycle;
\end{scope}

% Giọt chất lỏng chảy vào miệng ống
\draw[fill=biuretPurpleDeep!40, draw=glassLine, line width=0.4pt] 
    ($(\xThree, \hTube) + (0.02, 0.15)$) circle [radius=0.045];

% Chú thích trên pipette (3)
\node[right, align=left, inner sep=1pt] at ($(\xThree, \hTube) + (1.2, 1.8)$) {
    \textit{\textbf{Khoảng 4 mL dung dịch}}\\
    \textit{\textbf{lòng trắng trứng}}
};
\draw[leadLine] ($(\xThree, \hTube) + (0.8, 1.6)$) -- ($(\xThree, \hTube) + (1.2, 1.7)$);

% Chú thích Dung dịch X (3)
\draw[leadLine] ($(\xThree, 1.1) + (0.2, 0)$) -- ($(\xThree, 1.1) + (1.1, -0.1)$);
\node[right, align=left, inner sep=1pt] at ($(\xThree, 1.1) + (1.1, -0.1)$) {
    \textbf{\textit{\large Dung dịch X}}
};

% Nhãn ống (3)
\node[below=4pt] at (\xThree, -\rTube) {\textit{\textbf{(3)}}};

\end{tikzpicture}
\end{document}
'''

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tikz_code)

    print("Compiling TikZ to PDF...")
    cmd = ["pdflatex", "-interaction=nonstopmode", "thi_nghiem_biuret_long_trang_trung.tex"]
    res = subprocess.run(cmd, cwd=output_dir, capture_output=True, text=True)
    if res.returncode != 0:
        print("Error compiling TikZ:")
        print(res.stdout)
        return

    print("Converting PDF to high-resolution PNGs...")
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

    # Đồng thời sao chép đè vào thư mục ảnh đề thi Quang Trung 2 để các file đề tự động cập nhật nét đẹp
    qt2_img_dir = r"c:\Antigravity_Thanh\San_Pham\images_quangtrung2_hoa12_2102"
    if os.path.exists(qt2_img_dir):
        dest_img = os.path.join(qt2_img_dir, "hinh_p2_cau2_biuret.png")
        pix.save(dest_img)
        print(f"Updated exam asset: {dest_img}")

    # Dọn dẹp file phụ
    for ext in ['.aux', '.log']:
        f_clean = os.path.join(output_dir, "thi_nghiem_biuret_long_trang_trung" + ext)
        if os.path.exists(f_clean):
            os.remove(f_clean)

    print("Hoan tat tao bo anh vector va raster chat luong cao!")

if __name__ == "__main__":
    main()
