# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Giữa Kì 1 Hóa Học 10 - THPT Nguyễn Trường Tộ - Mã đề 104
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_10_NguyenTruongTo_Ma104_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_10_NguyenTruongTo_Ma104_GiaoVien.pdf (Kèm bảng đáp án rút gọn và lời giải chi tiết)
"""

import os
import sys
import subprocess
import glob

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

SAN_PHAM_DIR = r"C:\Antigravity_Thanh\San_Pham"

def build_latex_content(is_teacher=False):
    tex = []
    tex.append(r"""\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{geometry}
\geometry{top=1.5cm,bottom=1.2cm,left=1.4cm,right=1.4cm,headheight=28pt,headsep=12pt,footskip=18pt}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{tabularx}
\usepackage{array}
\usepackage{tikz}
\usepackage{fancyhdr}
\usepackage{enumitem}
\usepackage{multido}
\usepackage{tcolorbox}
\usepackage{colortbl}

\definecolor{hfRed}{HTML}{D32F2F}
\definecolor{hfGreen}{HTML}{8BC390}
\definecolor{mainGreen}{HTML}{2E7D32}
\definecolor{solBlue}{HTML}{1565C0}
\definecolor{tableGreen}{HTML}{E8F5E9}

\usepackage{ex_test}

\pagestyle{fancy}
\fancyhf{}
\renewcommand{\headrulewidth}{0.75pt}
\renewcommand{\headrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \headrulewidth\hfill}}
\renewcommand{\footrulewidth}{0.75pt}
\renewcommand{\footrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \footrulewidth\hfill}}

\fancyhead[L]{\small\itshape\color{hfRed}Tài liệu lưu hành nội bộ}
\fancyhead[R]{\small\color{hfRed}Thầy \textbf{TRẦN VĂN THẠNH} - 0777.470.803}

\fancyfoot[L]{\footnotesize\itshape\color{hfRed}CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế \quad|\quad CS2: 24 Đặng Thái Thân, TP Huế}
\fancyfoot[R]{\small\textbf{\color{hfRed}Trang \thepage}}

\setlength{\parindent}{0pt}
\setlength{\parskip}{2pt}

\begin{document}

\begin{minipage}[t]{0.48\textwidth}
\textbf{TRƯỜNG THPT NGUYỄN TRƯỜNG TỘ}\\
\textbf{TỔ HÓA - TIN}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 03 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}KIỂM TRA GIỮA KÌ I - NĂM HỌC 2026 - 2027}\\
\textbf{MÔN: HÓA HỌC, LỚP 10}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{MÃ ĐỀ 104}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ và tên:} \dotfill \textit{Lớp:} \dotfill

\vspace{2pt}
\noindent
\textit{\footnotesize (Cho biết nguyên tử khối: $\ce{H=1; C=12; N=14; O=16; Na=23; Mg=24; Al=27; Si=28; S=32; Cl=35,5; K=39; Ca=40; Ag=107,964}$).}

\vspace{2pt}
\hrule color{hfGreen} height 0.5pt
""")

    if is_teacher:
        tex.append(r"""
\vspace{4pt}
\begin{center}
\textbf{\color{hfRed}\large (BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}
\end{center}

\vspace{2pt}
\begin{center}
{\color{hfRed}\textbf{\large BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN)}}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}1. Phần I: Trắc nghiệm nhiều lựa chọn (16 câu)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} \\
\hline
\textbf{Đ/A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}2. Phần II: Trắc nghiệm Đúng / Sai (2 câu)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Lệnh a)} & \textbf{Lệnh b)} & \textbf{Lệnh c)} & \textbf{Lệnh d)} \\
\hline
\textbf{Câu 1} & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{solBlue}SAI} & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{solBlue}SAI} \\
\hline
\textbf{Câu 2} & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{solBlue}SAI} & \textbf{\color{solBlue}SAI} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}3. Phần III: Trả lời ngắn (4 câu)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}12} & \textbf{\color{hfRed}46} & \textbf{\color{hfRed}24} & \textbf{\color{hfRed}3} \\
\hline
\end{tabular}
\end{center}

\vspace{4pt}
\hrule color{hfGreen} height 0.5pt
""")

    # NỘI DUNG ĐỀ THI
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\Large A. PHẦN TRẮC NGHIỆM (7 ĐIỂM)}}

\vspace{3pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn}}\\
\textit{\small Học sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi học sinh chỉ chọn một phương án.}
\vspace{2pt}
""")

    # Câu 1
    tex.append(r"""
\begin{ex}
Tổng số electron tối đa có trong mỗi lớp L và M lần lượt là
\choice
{2 và 8.}
{8 và 32.}
{18 và 32.}
{\True 8 và 18.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Số electron tối đa trong lớp thứ $n$ là $2n^2$:\\
- Lớp L ($n = 2$): tối đa $2 \times 2^2 = 8$ electron.\\
- Lớp M ($n = 3$): tối đa $2 \times 3^2 = 18$ electron.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2
    tex.append(r"""
\begin{ex}
Cho phổ khối lượng MS của silver (Ag) ở hình dưới đây:
\begin{center}
\includegraphics[width=6.5cm]{images_hoa10/pho_khoi_luong_silver_ag.png}
\end{center}
Nguyên tử khối trung bình của nguyên tố silver là
\choice
{108,012.}
{\True 107,964.}
{108,000.}
{107,987.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Dựa vào phổ khối lượng:\\
- $^{107}\ce{Ag}$ chiếm 51,8\%.\\
- $^{109}\ce{Ag}$ chiếm 48,2\%.\\
Nguyên tử khối trung bình của silver:\\
$$\overline{A} = \frac{107 \times 51,8 + 109 \times 48,2}{100} = \frac{5542,6 + 5253,8}{100} = \frac{10796,4}{100} = 107,964.$$
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 3
    tex.append(r"""
\begin{ex}
Cho các phát biểu sau:\\
(a) Phân lớp 4s có mức năng lượng cao hơn phân lớp 3d.\\
(b) Trong mỗi orbital chỉ chứa tối đa 2 electron có chiều tự quay ngược nhau.\\
(c) Lớp M có 4 phân lớp.\\
(d) Số electron tối đa trong phân lớp 3d là 18.\\
(e) Số orbital của phân lớp s, p, d lần lượt là 1, 3, 5.\\
(f) Những electron ở gần hạt nhân có mức năng lượng cao nhất.\\
Số phát biểu đúng là
\choice
{5.}
{3.}
{\True 2.}
{4.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- (a) Sai: Mức năng lượng phân lớp $4s < 3d$.\\
- (b) Đúng: Nguyên lí Pauli quy định mỗi orbital chỉ chứa tối đa 2 electron có spin đối song.\\
- (c) Sai: Lớp M ($n=3$) chỉ có 3 phân lớp ($3s, 3p, 3d$).\\
- (d) Sai: Phân lớp d chứa tối đa 10 electron.\\
- (e) Đúng: Số AO của s, p, d lần lượt là 1, 3, 5.\\
- (f) Sai: Electron càng gần hạt nhân liên kết càng bền, có mức năng lượng thấp nhất.\\
Vậy có 2 phát biểu đúng là (b) và (e).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 4
    tex.append(r"""
\begin{ex}
Trong hạt nhân nguyên tử, hạt mang điện là
\choice
{proton và neutron.}
{\True proton.}
{electron.}
{proton và electron.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Trong hạt nhân chỉ chứa hai loại hạt: proton mang điện dương và neutron không mang điện. Do đó hạt mang điện trong hạt nhân chỉ có proton.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 5
    tex.append(r"""
\begin{ex}
Dãy nào dưới đây gồm các đồng vị của cùng một nguyên tố hóa học?
\choice
{$^{40}_{18}\ce{X}, ^{40}_{19}\ce{Y}.$}
{$^{19}_{9}\ce{X}, ^{20}_{10}\ce{Y}.$}
{$^{14}_{6}\ce{X}, ^{14}_{7}\ce{Y}.$}
{\True $^{28}_{14}\ce{X}, ^{29}_{14}\ce{Y}.$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Các đồng vị của cùng một nguyên tố là những nguyên tử có cùng số proton $Z$. Dãy D đều có $Z = 14$ (Silicon).\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 6
    tex.append(r"""
\begin{ex}
Nguyên tố hóa học là tập hợp các nguyên tử có cùng
\choice
{số hạt neutron.}
{tổng số hạt proton và neutron.}
{\True điện tích hạt nhân.}
{số khối.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Nguyên tố hóa học là tập hợp những nguyên tử có cùng điện tích hạt nhân (cùng số proton $Z$).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 7
    tex.append(r"""
\begin{ex}
Sự phân bố electron theo ô orbital nào dưới đây là đúng?
\choice
{$[\uparrow\uparrow]$}
{$[\uparrow\uparrow][\uparrow][\phantom{\uparrow}]$}
{\True $[\uparrow][\uparrow][\uparrow]$}
{$[\uparrow\downarrow][\uparrow][\phantom{\uparrow}]$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Theo quy tắc Hund, các electron phân bố vào các AO sao cho số electron độc thân là cực đại và có chiều spin song song: $[\uparrow][\uparrow][\uparrow]$.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 8
    tex.append(r"""
\begin{ex}
Cho cấu hình electron nguyên tử của một số nguyên tố như sau: $[\ce{Ne}]3s^2$, $[\ce{Ar}]3d^{10}4s^1$, $[\ce{Ar}]3d^{10}4s^2 4p^5$, $[\ce{He}]2s^2 2p^6$, $[\ce{Ar}]3d^7 4s^2$. Số nguyên tố thuộc nhóm A và nhóm B lần lượt là
\choice
{4; 1.}
{2; 2.}
{\True 3; 2.}
{2; 3.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
1. $[\ce{Ne}]3s^2$: nguyên tố s $\implies$ nhóm IIA (nhóm A).\\
2. $[\ce{Ar}]3d^{10}4s^1$: nguyên tố d $\implies$ nhóm IB (nhóm B).\\
3. $[\ce{Ar}]3d^{10}4s^2 4p^5$: nguyên tố p $\implies$ nhóm VIIA (nhóm A).\\
4. $[\ce{He}]2s^2 2p^6$: nguyên tố p $\implies$ nhóm VIIIA (nhóm A).\\
5. $[\ce{Ar}]3d^7 4s^2$: nguyên tố d $\implies$ nhóm VIIIB (nhóm B).\\
Vậy có 3 nguyên tố nhóm A và 2 nguyên tố nhóm B.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 9
    tex.append(r"""
\begin{ex}
Nguyên tử X có 56 electron và 81 neutron. Kí hiệu nguyên tử của X là
\choice
{\True $^{137}_{56}\ce{X}$.}
{$^{56}_{81}\ce{X}$.}
{$^{81}_{56}\ce{X}$.}
{$^{112}_{56}\ce{X}$.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$Z = e = 56$, $N = 81 \implies A = Z + N = 56 + 81 = 137$. Kí hiệu: $^{137}_{56}\ce{X}$.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 10
    tex.append(r"""
\begin{ex}
Số nguyên tố trong chu kì 3 và chu kì 5 lần lượt là
\choice
{18 và 18.}
{18 và 8.}
{\True 8 và 18.}
{8 và 8.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Chu kì 3 có 8 nguyên tố; chu kì 5 có 18 nguyên tố.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 11
    tex.append(r"""
\begin{ex}
Biết rằng một loại nguyên tử Sodium (Na) có 11 proton và 12 neutron. Phát biểu nào sau đây không đúng?
\choice
{Số khối của hạt nhân nguyên tử Sodium là 23.}
{Số hạt mang điện trong nguyên tử Sodium là 22.}
{Vỏ nguyên tử Sodium có 11 electron.}
{\True Điện tích hạt nhân nguyên tử Sodium là 11.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Điện tích hạt nhân phải là $+11$ hoặc $+11e$ (hay $+1,762 \cdot 10^{-18}\text{ C}$), còn 11 chỉ là số đơn vị điện tích hạt nhân ($Z=11$).\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 12
    tex.append(r"""
\begin{ex}
Nguyên tử của nguyên tố X có cấu hình electron là $\ce{1s^2 2s^1}$. Nguyên tố hóa học có cấu hình electron nào sau đây thuộc cùng một nhóm với nguyên tố X?
\choice
{$\ce{1s^2 2s^2 2p^1}$.}
{$\ce{1s^2 2s^2 2p^6 3s^2 3p^6 3d^5 4s^1}$.}
{\True $\ce{1s^2 2s^2 2p^6 3s^1}$.}
{$\ce{1s^2 2s^2 2p^6 3s^2 3p^1}$.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
X ($\ce{1s^2 2s^1}$) có 1 electron lớp ngoài cùng dạng $ns^1$, thuộc nhóm IA (kim loại kiềm). Cấu hình $\ce{1s^2 2s^2 2p^6 3s^1}$ (Na) cũng thuộc nhóm IA.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 13
    tex.append(r"""
\begin{ex}
Nguyên tử của nguyên tố X có electron ở mức năng lượng cao nhất là 3p. Nguyên tử của nguyên tố Y có một electron ở lớp N. Nguyên tử X và Y có số electron hơn kém nhau là 3. Nguyên tử Y, X lần lượt là
\choice
{\True kim loại, phi kim.}
{kim loại, kim loại.}
{khí hiếm, kim loại.}
{phi kim, kim loại.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- Y có 1e ở lớp N (lớp 4), cấu hình tận cùng là $4s^1 \implies$ Y là Potassium (K, $Z = 19$), là kim loại.\\
- Số electron hơn kém nhau 3: $Z_X = 19 - 3 = 16$ (Sulfur - S, $\ce{[Ne]3s^2 3p^4}$, electron mức năng lượng cao nhất là 3p). S là phi kim.\\
- Thứ tự hỏi: ``Nguyên tử Y, X lần lượt là'' $\implies$ kim loại, phi kim.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 14
    tex.append(r"""
\begin{ex}
Nguyên tố X được sử dụng như một nguyên liệu cơ bản trong sản xuất phân bón. X là chất vô cùng cần thiết cho sự tăng trưởng và phát triển của các cây xanh. Ngoài những công dụng trên, X còn là một khoáng chất có vai trò rất quan trọng đối với sức khỏe con người. Nguyên tử của nguyên tố X có tổng số hạt (proton, electron, neutron) là 58, trong đó tổng số hạt mang điện nhiều hơn số hạt không mang điện là 18. Nguyên tử X có số hạt proton là
\choice
{\True 19.}
{20.}
{38.}
{39.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Hệ phương trình: $2Z + N = 58$ và $2Z - N = 18 \implies 4Z = 76 \implies Z = 19$ (K).\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 15
    tex.append(r"""
\begin{ex}
Năm 1911, Rutherford cho chùm hạt $\alpha$ bắn vào một lá vàng rất mỏng và quan sát thấy hầu hết các hạt $\alpha$ đi thẳng qua lá vàng. Kết quả này chứng tỏ
\choice
{\True nguyên tử có cấu tạo rỗng.}
{hạt nhân nguyên tử mang điện tích âm.}
{electron có khối lượng rất lớn.}
{nguyên tử không chứa hạt mang điện.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Hầu hết các hạt alpha đi thẳng xuyên qua lá vàng chứng tỏ nguyên tử có cấu tạo rỗng.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 16
    tex.append(r"""
\begin{ex}
Trong bảng tuần hoàn, chu kì gồm các nguyên tố mà nguyên tử của chúng có cùng
\choice
{số phân lớp electron.}
{cấu hình electron.}
{số electron hóa trị.}
{\True số lớp electron.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Chu kì gồm các nguyên tố mà nguyên tử của chúng có cùng số lớp electron.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN II
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN II. Câu trắc nghiệm đúng sai}}\\
\textit{\small Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai.}
\vspace{2pt}
""")

    # Câu 1 Phần II
    tex.append(r"""
\begin{ex}
Sulfur dạng kem bôi được sử dụng để điều trị mụn trứng cá. Nguyên tử sulfur có phân lớp electron ngoài cùng là $\ce{3p^4}$.
\choiceTF
{\True Sulfur thuộc chu kì 3, nhóm VIA và là nguyên tố phi kim}
{Sulfur có số hiệu nguyên tử là 14}
{\True Nguyên tử $^{34}_{16}\ce{X}$ là đồng vị của sulfur}
{Ở trạng thái cơ bản, nguyên tử sulfur có 4 electron hóa trị}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Cấu hình electron đầy đủ của Sulfur: $\ce{1s^2 2s^2 2p^6 3s^2 3p^4}$ ($Z = 16$).\\
a) \textbf{ĐÚNG}: Có 3 lớp electron $\implies$ chu kì 3; có 6 electron lớp ngoài cùng $\implies$ nhóm VIA; là phi kim.\\
b) \textbf{SAI}: $Z = 16$ (không phải 14).\\
c) \textbf{ĐÚNG}: Có cùng số proton $Z = 16$ nên là đồng vị của sulfur.\\
d) \textbf{SAI}: Sulfur có 6 electron hóa trị ($\ce{3s^2 3p^4}$), không phải 4.
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần II
    tex.append(r"""
\begin{ex}
Hình dưới mô tả orbital (a) và orbital (b) chứa electron trong nguyên tử Sodium (Na) ở trạng thái cơ bản. Mức năng lượng của orbital (a) cao hơn orbital (b).
\begin{center}
\includegraphics[width=8cm]{images_hoa10/hinh_orbital_s_va_p.png}
\end{center}
\choiceTF
{\True Số electron trong orbital (b) gấp đôi số electron trong orbital (a)}
{\True Số tám nổi (b) là hình dạng của AO p}
{Electron trên orbital (a) nằm gần hạt nhân hơn electron trên orbital (b)}
{Electron trong orbital (a) và (b) cùng thuộc 1 lớp electron}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Nguyên tử Sodium (Na, $Z = 11$): $\ce{1s^2 2s^2 2p^6 3s^1}$.\\
Hình (a) là AO s, hình (b) là AO p. Vì năng lượng AO (a) cao hơn AO (b) nên (a) là orbital 3s và (b) là orbital 2p.\\
a) \textbf{ĐÚNG}: Orbital 3s chỉ có 1 electron ($3s^1$); orbital 2p chứa 2 electron ghép đôi. Do đó số e trong (b) gấp đôi số e trong (a).\\
b) \textbf{ĐÚNG}: AO p có dạng số tám nổi.\\
c) \textbf{SAI}: AO 3s ở lớp $n=3$ xa hạt nhân hơn AO 2p ở lớp $n=2$.\\
d) \textbf{SAI}: AO 3s thuộc lớp thứ 3, AO 2p thuộc lớp thứ 2, thuộc 2 lớp khác nhau.
}
""")
    tex.append(r"\end{ex}")

    # PHẦN III
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN III. Câu trắc nghiệm trả lời ngắn}}\\
\textit{\small Học sinh trả lời từ câu 1 đến câu 4.}
\vspace{2pt}
""")

    # Câu 1 Phần III
    tex.append(r"""
\begin{ex}
Khí carbon dioxide ($\ce{CO2}$) là nguyên nhân chính gây nên hiệu ứng nhà kính. Carbon có 2 đồng vị $^{12}_6\ce{C}, ^{13}_6\ce{C}$; oxygen có 3 đồng vị $^{16}_8\ce{O}, ^{17}_8\ce{O}, ^{18}_8\ce{O}$. Số loại phân tử $\ce{CO2}$ tạo thành là bao nhiêu?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phân tử $\ce{CO2}$ có dạng $\ce{O-C-O}$:\\
- Chọn 1 nguyên tử C từ 2 đồng vị: 2 cách.\\
- Chọn 2 nguyên tử O từ 3 đồng vị: $3 + \binom{3}{2} = 3 + 3 = 6$ cách.\\
Số loại phân tử $\ce{CO2}$ tạo thành: $2 \times 6 = 12$.\\
\textbf{Đáp án: 12.}
}
""")
    else:
        tex.append(r"\par\textbf{Trả lời:} \dotfill")
    tex.append(r"\end{ex}")

    # Câu 2 Phần III
    tex.append(r"""
\begin{ex}
Nguyên tố X thuộc chu kì 3, nhóm IIA. Nguyên tố Y thuộc cùng chu kì với X và thuộc nhóm VIIA. X và Y kết hợp với nhau tạo thành hợp chất có công thức $\ce{XY2}$. Tổng số proton có trong một phân tử $\ce{XY2}$ là bao nhiêu?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- X thuộc chu kì 3, nhóm IIA $\implies$ Mg ($Z_X = 12$).\\
- Y thuộc chu kì 3, nhóm VIIA $\implies$ Cl ($Z_Y = 17$).\\
Hợp chất $\ce{MgCl2}$ có tổng số proton là: $Z_X + 2Z_Y = 12 + 2 \times 17 = 46$.\\
\textbf{Đáp án: 46.}
}
""")
    else:
        tex.append(r"\par\textbf{Trả lời:} \dotfill")
    tex.append(r"\end{ex}")

    # Câu 3 Phần III
    tex.append(r"""
\begin{ex}
Cho các phát biểu sau:\\
(1) Nguyên tử có kích thước vô cùng nhỏ so với kích thước hạt nhân nguyên tử.\\
(2) Số khối của nguyên tử bằng tổng số proton và số neutron.\\
(3) Tất cả các nguyên tử đều chứa proton, electron và neutron.\\
(4) Khối lượng của nguyên tử tập trung hầu hết ở hạt nhân nguyên tử.\\
(5) Đồng vị là các nguyên tố có cùng số proton nhưng khác nhau số neutron.\\
Viết số thứ tự của các phát biểu đúng theo thứ tự tăng dần từ trái sang phải (ví dụ: 12 hoặc 123 hoặc 23...).
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- (1) Sai: Kích thước nguyên tử vô cùng lớn so với hạt nhân (lớn hơn $\sim 10^4$ lần).\\
- (2) Đúng: $A = Z + N$.\\
- (3) Sai: Đồng vị $^1_1\ce{H}$ không có neutron.\\
- (4) Đúng: Khối lượng electron rất bé nên khối lượng tập trung ở hạt nhân.\\
- (5) Sai theo SGK: Đồng vị là các nguyên tử (không phải các nguyên tố).\\
Các phát biểu đúng là (2) và (4) $\implies$ Viết: 24 (chấp nhận 245 nếu xét thông tục).\\
\textbf{Đáp án: 24.}
}
""")
    else:
        tex.append(r"\par\textbf{Trả lời:} \dotfill")
    tex.append(r"\end{ex}")

    # Câu 4 Phần III
    tex.append(r"""
\begin{ex}
Cho 7 nguyên tố có số hiệu nguyên tử lần lượt là: 5, 18, 11, 13, 16, 19, 2. Trong các nguyên tố trên, có bao nhiêu nguyên tố là kim loại?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- $Z = 5$ (B: phi kim); $Z = 18$ (Ar: khí hiếm); $Z = 11$ (Na: kim loại); $Z = 13$ (Al: kim loại); $Z = 16$ (S: phi kim); $Z = 19$ (K: kim loại); $Z = 2$ (He: khí hiếm).\\
Có 3 nguyên tố kim loại là Na ($Z=11$), Al ($Z=13$), K ($Z=19$).\\
\textbf{Đáp án: 3.}
}
""")
    else:
        tex.append(r"\par\textbf{Trả lời:} \dotfill")
    tex.append(r"\end{ex}")

    # PHẦN B: TỰ LUẬN
    tex.append(r"""
\vspace{6pt}
\noindent
{\color{mainGreen}\textbf{\Large B. PHẦN TỰ LUẬN (3 ĐIỂM)}}
\vspace{3pt}
""")

    # Câu 1 Tự luận
    tex.append(r"""
\begin{ex}
Silicon (Si) được sử dụng trong nhiều ngành công nghiệp: gốm, men sứ, thủy tinh, luyện thép, vật liệu bán dẫn \dots\\
a) Viết cấu hình electron nguyên tử của nguyên tố silicon ($Z = 14$).\\
b) Xác định vị trí của silicon (ô nguyên tố, chu kì, nhóm) trong bảng tuần hoàn và giải thích.
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) Cấu hình electron của Si ($Z = 14$):\\
$\ce{1s^2 2s^2 2p^6 3s^2 3p^2}$ (hoặc $[\ce{Ne}]\ce{3s^2 3p^2}$). \hfill \textit{(0,5 điểm)}\\
b) Vị trí của silicon trong bảng tuần hoàn:\\
- Ô nguyên tố: số 14 (vì $Z = 14$, có 14 proton).\\
- Chu kì: Chu kì 3 (vì có 3 lớp electron).\\
- Nhóm: Nhóm IVA (vì có 4 electron hóa trị ở lớp ngoài cùng $\ce{3s^2 3p^2}$, electron cuối cùng điền vào phân lớp p). \hfill \textit{(0,5 điểm)}
}
""")
    else:
        tex.append(r"""
\vspace{0.3cm}
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 2 Tự luận
    tex.append(r"""
\begin{ex}
Cho X, Y, T là ba nguyên tố liên tiếp nhau trong một chu kì của bảng tuần hoàn các nguyên tố hóa học. Tổng số các hạt mang điện trong thành phần cấu tạo nguyên tử của X, Y, T bằng 72. Xác định số hiệu nguyên tử của các nguyên tố X, Y, T?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Vì X, Y, T là ba nguyên tố liên tiếp nhau trong một chu kì nên số hiệu nguyên tử hơn kém nhau 1 đơn vị. Giả sử $Z_X < Z_Y < Z_T$:\\
$Z_X = Z_Y - 1$ và $Z_T = Z_Y + 1$.\\
Tổng số hạt mang điện là 72:\\
$$2Z_X + 2Z_Y + 2Z_T = 72 \implies Z_X + Z_Y + Z_T = 36$$
$$\implies (Z_Y - 1) + Z_Y + (Z_Y + 1) = 36 \implies 3Z_Y = 36 \implies Z_Y = 12\text{ (Mg)}.$$
Suy ra: $Z_X = 12 - 1 = 11$ (Na) và $Z_T = 12 + 1 = 13$ (Al).\\
Kiểm tra: $\ce{Na, Mg, Al}$ đều thuộc chu kì 3 (thỏa mãn).\\
Vậy: $Z_X = 11$, $Z_Y = 12$, $Z_T = 13$. \hfill \textit{(1,0 điểm)}
}
""")
    else:
        tex.append(r"""
\vspace{0.3cm}
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 3 Tự luận
    tex.append(r"""
\begin{ex}
Một nguyên tố R có 3 đồng vị X, Y, Z, tổng số hạt electron, proton và neutron của 3 đồng vị bằng 111. Số neutron của đồng vị X bằng số proton, số neutron của đồng vị Z hơn đồng vị Y 1 hạt.\\
a) Xác định số khối của 3 đồng vị.\\
b) Cho biết tỷ lệ số nguyên tử của các đồng vị như sau: $\ce{X : Y} = 393 : 50$ và $\ce{Y : Z} = 50 : 57$, xác định nguyên tử khối trung bình của R và khối lượng của $9,033 \cdot 10^{22}$ nguyên tử R.\\
\textit{(Biết số Avogadro $= 6,022 \cdot 10^{23}$)}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) Xác định số khối của 3 đồng vị:\\
- Gọi $Z$ là số proton của nguyên tố R; $N_X, N_Y, N_Z$ là số neutron của 3 đồng vị.\\
Tổng số hạt của 3 đồng vị:\\
$$(2Z + N_X) + (2Z + N_Y) + (2Z + N_Z) = 111 \implies 6Z + N_X + N_Y + N_Z = 111.\quad (1)$$
Theo đề: $N_X = Z$ (2) và $N_Z = N_Y + 1$ (3).\\
Thay (2) và (3) vào (1):
$$6Z + Z + N_Y + (N_Y + 1) = 111 \implies 7Z + 2N_Y = 110 \implies 2N_Y = 110 - 7Z.$$
Vì $2N_Y$ chẵn nên $Z$ chẵn ($Z < 110/7 \approx 15,7$). Với đồng vị bền ($1 \le N_Y/Z \le 1,5$):\\
- Với $Z = 12$ (Mg): $2N_Y = 110 - 84 = 26 \implies N_Y = 13 \implies N_Y/Z = 13/12 \approx 1,083$ (thỏa mãn bền).\\
Suy ra:\\
- $N_X = Z = 12 \implies A_X = 12 + 12 = 24$.\\
- $N_Y = 13 \implies A_Y = 12 + 13 = 25$.\\
- $N_Z = 13 + 1 = 14 \implies A_Z = 12 + 14 = 26$.\\
Vậy số khối 3 đồng vị là: $A_X = 24$, $A_Y = 25$, $A_Z = 26$. \hfill \textit{(0,5 điểm)}\\
b) Tính $\overline{A}$ và khối lượng:\\
- Tỉ lệ số nguyên tử: $\ce{X : Y : Z} = 393 : 50 : 57$ (tổng số phần $= 500$).\\
$$\overline{A}_R = \frac{24 \times 393 + 25 \times 50 + 26 \times 57}{500} = \frac{12164}{500} = 24,328.$$
- Số mol của $9,033 \cdot 10^{22}$ nguyên tử R:
$$n = \frac{9,033 \cdot 10^{22}}{6,022 \cdot 10^{23}} = 0,15\text{ mol.}$$
Khối lượng của $9,033 \cdot 10^{22}$ nguyên tử R là:
$$m = n \times \overline{A}_R = 0,15 \times 24,328 = 3,6492\text{ gam (xấp xỉ } 3,65\text{ gam)}. \hfill \textit{(0,5 điểm)}$$
}
""")
    else:
        tex.append(r"""
\vspace{0.3cm}
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill
""")
    tex.append(r"\end{ex}")

    # Kết thúc
    tex.append(r"""
\vspace{6pt}
\begin{center}
\textbf{---------- HẾT ----------}\\
\textit{\small - Học sinh không được sử dụng tài liệu.\\
- Cán bộ coi thi không giải thích gì thêm.}
\end{center}
\end{document}
""")

    content = "".join(tex)
    if not is_teacher:
        content = content.replace(r"\True ", "")
    return content

def compile_pdf(tex_path):
    out_dir = os.path.dirname(tex_path)
    base_name = os.path.splitext(os.path.basename(tex_path))[0]
    
    for i in range(2):
        cmd = ["pdflatex", "-interaction=nonstopmode", f"-output-directory={out_dir}", tex_path]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, cwd=out_dir)
        if res.returncode != 0 and i == 1:
            print(f"[CẢNH BÁO] pdflatex kết thúc với mã {res.returncode}")
            log_file = os.path.join(out_dir, f"{base_name}.log")
            if os.path.exists(log_file):
                with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
                    lines = f.readlines()
                    print("".join(lines[-30:]))
    
    for ext in [".aux", ".log", ".thm", ".out", ".synctex.gz"]:
        f_junk = os.path.join(out_dir, base_name + ext)
        if os.path.exists(f_junk):
            try:
                os.remove(f_junk)
            except Exception:
                pass
    for f in glob.glob(os.path.join(out_dir, "ans-*.tex")):
        try:
            os.remove(f)
        except Exception:
            pass
    print(f"[OK] Đã biên dịch xong PDF: {os.path.join(out_dir, base_name + '.pdf')}")

def main():
    tex_hs = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_10_NguyenTruongTo_Ma104_HocSinh.tex")
    tex_gv = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_10_NguyenTruongTo_Ma104_GiaoVien.tex")

    print("Đang tạo mã nguồn LaTeX Bản Học sinh...")
    with open(tex_hs, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))

    print("Đang tạo mã nguồn LaTeX Bản Giáo viên...")
    with open(tex_gv, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))

    print("\nĐang biên dịch PDF Bản Học sinh...")
    compile_pdf(tex_hs)

    print("\nĐang biên dịch PDF Bản Giáo viên...")
    compile_pdf(tex_gv)

    print("\n[HOÀN TẤT] Toàn bộ file .tex và .pdf đã được tạo và biên dịch thành công!")

if __name__ == "__main__":
    main()
