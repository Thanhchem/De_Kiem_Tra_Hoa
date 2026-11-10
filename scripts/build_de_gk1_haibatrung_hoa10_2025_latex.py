# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Giữa Kì 1 Hóa Học 10 - THPT Hai Bà Trưng (2024 - 2025)
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_10_HaiBaTrung_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_10_HaiBaTrung_GiaoVien.pdf (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
\textbf{SỞ GD\&ĐT THỪA THIÊN HUẾ}\\
\textbf{TRƯỜNG THPT HAI BÀ TRƯNG}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 03 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}ĐỀ KIỂM TRA GIỮA KÌ I - NĂM HỌC 2024 - 2025}\\
\textbf{MÔN: HOÁ HỌC -- LỚP 10}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{MÃ ĐỀ: GỐC}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ, tên thí sinh:} \dotfill \textit{Lớp:} \dotfill \textit{Số báo danh:} \dotfill

\vspace{3pt}
\noindent
\textit{\footnotesize Cho NTK: $\ce{Na=23; K=39; Mg=24; Ca=40; Ba=137; Al=27; Cl=35,5; H=1; C=12; O=16; N=14}$.\\
(Học sinh không sử dụng bảng hệ thống tuần hoàn các nguyên tố hóa học)}

\vspace{3pt}
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
\textbf{\color{mainGreen}1. Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - mỗi câu đúng 0,25 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} \\
\hline
\textbf{Đ/A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm - 1 ý: 0,1đ; 2 ý: 0,25đ; 3 ý: 0,5đ; 4 ý: 1,0đ)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Nội dung} & \textbf{Lệnh a} & \textbf{Lệnh b} & \textbf{Lệnh c} & \textbf{Lệnh d} \\
\hline
\textbf{Câu 1} & Cấu tạo nguyên tử & \textbf{\color{solBlue}SAI} & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{solBlue}SAI} & \textbf{\color{solBlue}SAI} \\
\hline
\textbf{Câu 2} & Nguyên tử P ($Z=15$) & \textbf{\color{solBlue}SAI} & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{solBlue}SAI} & \textbf{\color{hfRed}ĐÚNG} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}3. Phần III: Trả lời ngắn (2,0 điểm - mỗi câu đúng 0,5 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}3, 2, 1, 4} & \textbf{\color{hfRed}2} & \textbf{\color{hfRed}0,1\%} & \textbf{\color{hfRed}17} \\
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
{\color{mainGreen}\textbf{\Large A. PHẦN TRẮC NGHIỆM: (7,0 điểm)}}

\vspace{3pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN I (3,0 điểm). TRẮC NGHIỆM NHIỀU LỰA CHỌN}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{2pt}
""")

    # Câu 1
    tex.append(r"""
\begin{ex}
Nội dung nào dưới đây không phải là đối tượng nghiên cứu của hóa học ?
\choice
{Thành phần, cấu trúc của chất.}
{Tính chất và sự biến đổi của chất.}
{Ứng dụng của chất.}
{\True Sự lớn lên và sinh sản của tế bào.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Đối tượng nghiên cứu của hóa học là chất và sự biến đổi của chất (bao gồm thành phần, cấu trúc, tính chất và ứng dụng của chất).\\
``Sự lớn lên và sinh sản của tế bào'' là đối tượng nghiên cứu của khoa học Sinh học.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2
    tex.append(r"""
\begin{ex}
Nguyên tử chứa những hạt mang điện là
\choice
{proton và neutron.}
{\True proton và electron.}
{electron và neutron.}
{electron và hạt alpha.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Hạt proton mang điện tích dương (+1), hạt electron mang điện tích âm (-1), còn neutron không mang điện.\\
Do đó, các hạt mang điện trong nguyên tử gồm proton và electron.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 3
    tex.append(r"""
\begin{ex}
Chọn đáp án sai
\choice
{Proton, $m \approx 1\text{ amu}, q = +1$.}
{Electron, $m \approx 0,00055\text{ amu}, q = -1$.}
{Neutron, $m \approx 1\text{ amu}, q = 0$.}
{\True Electron, $m \approx 1\text{ amu}, q = -1$.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Khối lượng của hạt electron rất nhỏ, xấp xỉ $0,00055\text{ amu}$ (khoảng $1/1836\text{ amu}$), có thể bỏ qua khi tính khối lượng nguyên tử.\\
Nhận định D ghi electron có $m \approx 1\text{ amu}$ là sai.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 4
    tex.append(r"""
\begin{ex}
Thông tin nào sau đây không đúng về $^{28}_{14}\ce{N}$ (hoặc $^{28}_{14}\ce{Si}$)?
\choice
{Số khối là 28.}
{Tổng số proton và neutron là 28.}
{\True Số neuton là 28.}
{Số đơn vị điện tích hạt nhân là 14.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Kí hiệu nguyên tử $^{A}_{Z}\ce{X}$ cho biết:\\
- $Z = 14$: Số đơn vị điện tích hạt nhân (D đúng).\\
- $A = 28$: Số khối (A đúng).\\
- Tổng số proton và neutron: $Z + N = A = 28$ (B đúng).\\
- Số neutron: $N = A - Z = 28 - 14 = 14$ (C sai vì ghi là 28).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 5
    tex.append(r"""
\begin{ex}
Nguyên tử của nguyên tố X có 56 electron trong lớp vỏ, 82 neutron trong hạt nhân. Kí hiệu của nguyên tử nguyên tố X là
\choice
{$^{56}_{82}\ce{A}$.}
{$^{82}_{56}\ce{A}$.}
{$^{56}_{138}\ce{A}$.}
{\True $^{138}_{56}\ce{A}$.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- Trong nguyên tử trung hòa điện: $Z = e = 56$.\\
- Số khối: $A = Z + N = 56 + 82 = 138$.\\
- Kí hiệu nguyên tử: $^{A}_{Z}\ce{X}$ hay $^{138}_{56}\ce{A}$ (nguyên tố Barium - Ba).\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 6
    tex.append(r"""
\begin{ex}
Phát biểu nào sau đây không đúng
\choice
{Orbital nguyên tử được kí hiệu là AO (Atomic Orbital).}
{Theo mô hình hiện đại, trong nguyên tử, electron chuyển động rất nhanh, không theo quỹ đạo xác định.}
{\True Có 3 loại orbital nguyên tử: orbital s, orbital p và orbital d.}
{Các orbital s có dạng hình cầu và orbital p có dạng hình số 8 nổi.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Các loại orbital nguyên tử gồm có s, p, d, f... chứ không chỉ giới hạn ở 3 loại s, p, d.\\
Do đó nhận định C là không đúng.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 7
    tex.append(r"""
\begin{ex}
Lớp N có số phân lớp electron bằng
\choice
{1.}
{2.}
{3.}
{\True 4.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Lớp N là lớp thứ $n = 4$ (K: 1, L: 2, M: 3, N: 4).\\
Với $n \le 4$, lớp thứ $n$ có $n$ phân lớp. Vậy lớp N ($n = 4$) có 4 phân lớp gồm: $4s, 4p, 4d, 4f$.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 8
    tex.append(r"""
\begin{ex}
Nguyên tử nguyên tố M có số hiệu nguyên tử bằng 20. Cấu hình electron của ion $\ce{M^{2+}}$ là
\choice
{$\ce{1s^2 2s^2 2p^6 3s^2 3p^6 4s^2}$.}
{\True $\ce{1s^2 2s^2 2p^6 3s^2 3p^6}$.}
{$\ce{1s^2 2s^2 2p^6 3s^2 3p^6 3d^2 4s^2}$.}
{$\ce{1s^2 2s^2 2p^8 3s^2 3p^6}$.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Cấu hình electron của M ($Z = 20$): $\ce{1s^2 2s^2 2p^6 3s^2 3p^6 4s^2}$.\\
Khi tạo ion $\ce{M^{2+}}$, nguyên tử M nhường 2e ở phân lớp ngoài cùng: $\ce{M -> M^{2+} + 2e}$.\\
Cấu hình electron của $\ce{M^{2+}}$ là: $\ce{1s^2 2s^2 2p^6 3s^2 3p^6}$.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 9
    tex.append(r"""
\begin{ex}
Chu kì là dãy các nguyên tố được xếp theo chiều điện tích hạt nhân tăng dần, nguyên tử của chúng có cùng
\choice
{số electron ở lớp ngoài cùng.}
{\True số lớp electron.}
{số electron hóa trị.}
{số electron.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Theo định nghĩa: Chu kì là dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được sắp xếp theo chiều điện tích hạt nhân tăng dần.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 10
    tex.append(r"""
\begin{ex}
Nguyên tố Y thuộc chu kì 4, nhóm IA của bảng tuần hoàn. Phát biểu nào sau đây về Y là đúng ?
\choice
{Y có độ âm điện lớn nhất và bán kính nguyên tử nhỏ nhất trong chu kì 4.}
{Y có độ âm điện lớn nhất và bán kính nguyên tử lớn nhất trong chu kì 4.}
{\True Y có độ âm điện nhỏ nhất và bán kính nguyên tử lớn nhất trong chu kì 4.}
{Y có độ âm điện nhỏ nhất và bán kính nguyên tử nhỏ nhất trong chu kì 4.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Trong một chu kì, từ trái sang phải theo chiều tăng của điện tích hạt nhân: bán kính nguyên tử giảm dần, độ âm điện tăng dần.\\
Nguyên tố Y ở nhóm IA đứng đầu chu kì 4 nên có độ âm điện nhỏ nhất và bán kính nguyên tử lớn nhất trong chu kì 4.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 11
    tex.append(r"""
\begin{ex}
Trong một chu kì, bán kính nguyên tử
\choice
{xấp xỉ bằng nhau.}
{biến đổi ngẫu nhiên.}
{tăng theo chiều tăng dần của điện tích hạt nhân.}
{\True giảm theo chiều tăng dần của điện tích hạt nhân.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Trong một chu kì, khi đi từ trái sang phải theo chiều tăng của điện tích hạt nhân, số lớp electron không đổi nhưng điện tích hạt nhân tăng làm tăng lực hút tĩnh điện giữa hạt nhân với electron, dẫn đến bán kính nguyên tử giảm dần.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 12
    tex.append(r"""
\begin{ex}
Cho các nguyên tố sau: $\ce{Li}\ (Z=3),\ \ce{O}\ (Z=8),\ \ce{F}\ (Z=9),\ \ce{Na}\ (Z=11)$. Dãy gồm các nguyên tố được sắp xếp theo chiều tăng dần bán kính nguyên tử từ trái sang phải là
\choice
{F, Li, O, Na.}
{Li, Na, O, F.}
{F, Na, O, Li.}
{\True F, O, Li, Na.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- Li ($Z=3$), O ($Z=8$), F ($Z=9$) cùng thuộc chu kì 2. Trong một chu kì, bán kính giảm dần: $r_{\ce{F}} < r_{\ce{O}} < r_{\ce{Li}}$.\\
- Li ($Z=3$) và Na ($Z=11$) cùng thuộc nhóm IA. Trong một nhóm A, bán kính tăng dần: $r_{\ce{Li}} < r_{\ce{Na}}$.\\
$\implies$ Chiều tăng dần bán kính: $r_{\ce{F}} < r_{\ce{O}} < r_{\ce{Li}} < r_{\ce{Na}}$ (F, O, Li, Na).\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN II
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN II (2,0 điểm). CÂU TRẮC NGHIỆM ĐÚNG / SAI}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng (Đ) hoặc sai (S).}
\vspace{2pt}
""")

    # Câu 1 Phần II
    tex.append(r"""
\begin{ex}
Các phát biểu sau đây là đúng hay sai?
\choiceTF
{Nguyên tử của nguyên tố hydrogen luôn có 2 proton}
{\True Nguyên tử của nguyên tố hydrogen có thể không có neutron}
{Nếu đường kính của nguyên tử khoảng $10^2\text{ pm}$ thì đường kính của hạt nhân khoảng $10^{-4}\text{ pm}$}
{Nguyên tố hóa học là tập hợp các nguyên tử có cùng số neutron trong hạt nhân}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) \textbf{SAI}: Nguyên tử hydrogen có $Z = 1$ nên luôn chỉ có 1 proton trong hạt nhân.\\
b) \textbf{ĐÚNG}: Đồng vị phổ biến nhất của hydrogen là protium ($^1_1\ce{H}$) không có neutron ($N = 1 - 1 = 0$).\\
c) \textbf{SAI}: Kích thước hạt nhân nhỏ hơn kích thước nguyên tử khoảng $10^4$ lần. Nếu đường kính nguyên tử khoảng $10^2\text{ pm}$ thì đường kính hạt nhân khoảng $10^2 \times 10^{-4} = 10^{-2}\text{ pm}$ (chứ không phải $10^{-4}\text{ pm}$).\\
d) \textbf{SAI}: Nguyên tố hóa học là tập hợp các nguyên tử có cùng số proton (điện tích hạt nhân), không phải cùng số neutron.
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần II
    tex.append(r"""
\begin{ex}
Cho sơ đồ của một nguyên tử X được biểu diễn như hình bên:
\begin{center}
\includegraphics[width=4.5cm]{images_hoa10/so_do_nguyen_tu_p_15.png}
\end{center}
\choiceTF
{Các electron trong nguyên tử X chỉ được phân bố trên 2 lớp K và lớp L}
{\True Ở trạng thái cơ bản, nguyên tử X có 3 electron độc thân}
{Nguyên tử nguyên tố X có tổng số hạt cơ bản là 50 hạt. Số hạt mang điện nhiều hơn số hạt không mang điện là 10 hạt. Số hạt neutron của X là 35}
{\True X thuộc ô 15, chu kỳ 3, nhóm VA}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) \textbf{SAI}: Electron phân bố trên 3 lớp: lớp 1 (K có 2e), lớp 2 (L có 8e), lớp 3 (M có 5e).\\
b) \textbf{ĐÚNG}: Cấu hình electron của X ($Z = 15$): $\ce{1s^2 2s^2 2p^6 3s^2 3p^3}$. Phân lớp $3p^3$ có 3 electron độc thân chiếm 3 orbital $p$ theo quy tắc Hund.\\
c) \textbf{SAI}: Giải hệ: $2Z + N = 50$ và $2Z - N = 10 \implies 4Z = 60 \implies Z = 15$; $N = 50 - 2 \times 15 = 20$ (chứ không phải 35).\\
d) \textbf{ĐÚNG}: $Z = 15$ (ô 15); có 3 lớp electron (chu kì 3); có 5e lớp ngoài cùng và là nguyên tố p $\implies$ nhóm VA (Phosphorus - P).
}
""")
    tex.append(r"\end{ex}")

    # PHẦN III
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN III (2,0 điểm). CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 4.}
\vspace{2pt}
""")

    # Câu 1 Phần III
    tex.append(r"""
\begin{ex}
Cho dãy nguyên tố sau: Na (1), K (2), Rb (3), Li (4). Hãy sắp xếp theo chiều giảm dần bán kính nguyên tử? (ghi thứ tự các số)
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Các nguyên tố Li, Na, K, Rb đều thuộc nhóm IA.\\
Trong cùng nhóm A, theo chiều tăng của $Z$ (từ trên xuống: Li $\to$ Na $\to$ K $\to$ Rb), bán kính nguyên tử tăng dần: $r_{\ce{Li}} < r_{\ce{Na}} < r_{\ce{K}} < r_{\ce{Rb}}$.\\
Sắp xếp theo chiều GIẢM DẦN bán kính nguyên tử: $\ce{Rb} > \ce{K} > \ce{Na} > \ce{Li}$, tương ứng dãy số: 3, 2, 1, 4.\\
\textbf{Đáp số: 3, 2, 1, 4}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần III
    tex.append(r"""
\begin{ex}
Cho số hiệu nguyên tử của $\ce{Li}\ (Z=3),\ \ce{Na}\ (Z=11),\ \ce{P}\ (Z=15),\ \ce{Ar}\ (Z=18)$. Có bao nhiêu nguyên tố p trong các nguyên tố trên?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Cấu hình electron:\\
- $\ce{Li}\ (Z=3): \ce{1s^2 2s^1} \implies$ nguyên tố s.\\
- $\ce{Na}\ (Z=11): \ce{[Ne] 3s^1} \implies$ nguyên tố s.\\
- $\ce{P}\ (Z=15): \ce{[Ne] 3s^2 3p^3} \implies$ nguyên tố p.\\
- $\ce{Ar}\ (Z=18): \ce{[Ne] 3s^2 3p^6} \implies$ nguyên tố p.\\
Vậy có 2 nguyên tố p là P và Ar.\\
\textbf{Đáp số: 2}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 3 Phần III
    tex.append(r"""
\begin{ex}
Đồng vị phóng xạ cobalt ($^{60}_{27}\ce{Co}$) phát ra tia $\gamma$ có khả năng đâm xuyên mạnh, dùng điều trị các khối u ở sâu trong cơ thể. Cobalt có ba đồng vị: $^{59}_{27}\ce{Co}$ (chiếm 98\%), $^{58}_{27}\ce{Co}$ và $^{60}_{27}\ce{Co}$; nguyên tử khối trung bình là 58,982. Xác định hàm lượng \% của đồng vị phóng xạ $^{60}_{27}\ce{Co}$.
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Gọi \% số nguyên tử của $^{58}_{27}\ce{Co}$ là $x\%$ và của $^{60}_{27}\ce{Co}$ là $y\%$.\\
Tổng \% các đồng vị: $98\% + x\% + y\% = 100\% \implies x + y = 2 \implies x = 2 - y$.\\
Biểu thức nguyên tử khối trung bình:
$$\overline{A} = \dfrac{59 \times 98 + 58 \times x + 60 \times y}{100} = 58,982$$
$$\implies 5782 + 58(2 - y) + 60y = 5898,2 \implies 5898 + 2y = 5898,2 \implies 2y = 0,2 \implies y = 0,1\%.$$
\textbf{Đáp số: 0,1\% (hoặc 0,1)}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 4 Phần III
    tex.append(r"""
\begin{ex}
Nguyên tử X có tổng số hạt cơ bản là 49, số hạt không mang điện bằng 53,125\% số hạt mang điện. Số neutron của nguyên tử X là bao nhiêu?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Tổng số hạt: $2Z + N = 49$ (1).\\
Số hạt không mang điện bằng 53,125\% số hạt mang điện:
$$N = 0,53125 \times 2Z = 1,0625 Z \quad (2)$$
Thay (2) vào (1):
$$2Z + 1,0625 Z = 49 \implies 3,0625 Z = 49 \implies Z = 16.$$
Số neutron: $N = 49 - 2 \times 16 = 17$ (nguyên tố Sulfur - S).\\
\textbf{Đáp số: 17}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # PHẦN TỰ LUẬN
    tex.append(r"""
\vspace{6pt}
\noindent
{\color{mainGreen}\textbf{\Large B. PHẦN TỰ LUẬN: (3,0 điểm)}}
\vspace{2pt}
""")

    # Câu 1 Tự luận
    tex.append(r"""
\noindent
\textbf{Câu 1 (1,0 điểm):} Cho nguyên tố A có số hiệu nguyên tử $Z = 13$.\\
a) Viết cấu hình electron nguyên tử của nguyên tố A. Cho biết A là kim loại, phi kim hay khí hiếm?\\
b) Biểu diễn cấu hình electron của A theo ô orbital. Cho biết số electron độc thân của nguyên tử A?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 1:}]
\textbf{a) (0,5 điểm)}
\begin{itemize}[leftmargin=1.0cm]
\item Cấu hình electron của nguyên tử A ($Z = 13$): $\ce{1s^2 2s^2 2p^6 3s^2 3p^1}$ (hoặc $\ce{[Ne] 3s^2 3p^1}$). \hfill \textit{(0,25 đ)}
\item A có 3 electron ở lớp ngoài cùng ($n = 3$), do đó A là nguyên tố kim loại (Aluminium - Al). \hfill \textit{(0,25 đ)}
\end{itemize}
\textbf{b) (0,5 điểm)}
\begin{itemize}[leftmargin=1.0cm]
\item Biểu diễn theo ô orbital:
  \begin{itemize}
  \item Lớp 1 ($1s$): 1 ô orbital chứa 2 electron ghép đôi ($\uparrow\downarrow$).
  \item Lớp 2 ($2s, 2p$): 1 ô $2s$ chứa 2 electron ghép đôi ($\uparrow\downarrow$); 3 ô $2p$ chứa 6 electron ghép đôi ($\uparrow\downarrow\ \uparrow\downarrow\ \uparrow\downarrow$).
  \item Lớp 3 ($3s, 3p$): 1 ô $3s$ chứa 2 electron ghép đôi ($\uparrow\downarrow$); 3 ô $3p$ có 1 electron độc thân ($\uparrow\ \square\ \square$). \hfill \textit{(0,25 đ)}
  \end{itemize}
\item Kết luận: Số electron độc thân của nguyên tử A là 1 electron. \hfill \textit{(0,25 đ)}
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{6}{\dotfill\\[4pt]}
""")

    # Câu 2 Tự luận
    tex.append(r"""
\vspace{4pt}
\noindent
\textbf{Câu 2 (1,0 điểm):} Kim loại M thuộc nhóm IIA của bảng tuần hoàn, là một thành phần dinh dưỡng quan trọng. Sự thiếu hụt rất nhỏ của nó đã ảnh hưởng tới sự hình thành và phát triển của xương, răng. Thừa M có thể dẫn đến sỏi thận. Cho 1,2 gam M tác dụng hết với dung dịch HCl, thu được 0,7437 L khí (đkc).\\
a) Xác định M và viết cấu hình electron của ion $\ce{M^{2+}}$.\\
b) Sắp xếp theo thứ tự tăng dần về bán kính nguyên tử của nguyên tố M với $_{19}\ce{K}$ và $_{12}\ce{Mg}$.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 2:}]
\textbf{a) (0,5 điểm)}
\begin{itemize}[leftmargin=1.0cm]
\item Số mol khí $\ce{H2}$ ở điều kiện chuẩn ($25^\circ\text{C}, 1\text{ bar}$):
$$n_{\ce{H2}} = \dfrac{0,7437}{24,79} = 0,03\text{ mol}. \hfill \textit{(0,1 đ)}$$
\item Phương trình phản ứng: $\ce{M + 2HCl -> MCl2 + H2\uparrow}$
$$\implies n_{\ce{M}} = n_{\ce{H2}} = 0,03\text{ mol}. \hfill \textit{(0,1 đ)}$$
\item Khối lượng mol nguyên tử của M:
$$M_{\ce{M}} = \dfrac{1,2}{0,03} = 40\text{ g/mol} \implies \text{M là kim loại Calcium (Ca, } Z = 20\text{)}. \hfill \textit{(0,15 đ)}$$
\item Cấu hình electron của $\ce{Ca}\ (Z=20)$: $\ce{1s^2 2s^2 2p^6 3s^2 3p^6 4s^2}$ hay $\ce{[Ar] 4s^2}$.\\
$\implies$ Cấu hình electron của ion $\ce{Ca^{2+}}$: $\ce{1s^2 2s^2 2p^6 3s^2 3p^6}$ hay $\ce{[Ar]}$. \hfill \textit{(0,15 đ)}
\end{itemize}
\textbf{b) (0,5 điểm)}
\begin{itemize}[leftmargin=1.0cm]
\item So sánh bán kính nguyên tử giữa $\ce{Mg}\ (Z=12)$, $\ce{Ca}\ (Z=20)$, $\ce{K}\ (Z=19)$:
  \begin{itemize}
  \item $\ce{Mg}$ và $\ce{Ca}$ cùng nhóm IIA: Từ trên xuống dưới số lớp e tăng $\implies r_{\ce{Mg}} < r_{\ce{Ca}}$. \hfill \textit{(0,2 đ)}
  \item $\ce{K}$ và $\ce{Ca}$ cùng chu kì 4: Từ trái sang phải $Z$ tăng $\implies r_{\ce{Ca}} < r_{\ce{K}}$. \hfill \textit{(0,2 đ)}
  \end{itemize}
\item Sắp xếp theo chiều tăng dần bán kính nguyên tử: $r_{\ce{Mg}} < r_{\ce{Ca}} < r_{\ce{K}}$ hay $\ce{Mg < Ca < K}$. \hfill \textit{(0,1 đ)}
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{7}{\dotfill\\[4pt]}
""")

    # Câu 3 Tự luận
    tex.append(r"""
\vspace{4pt}
\noindent
\textbf{Câu 3 (1,0 điểm):} Trong tự nhiên nguyên tố X có hai đồng vị là $^{35}_{17}\ce{X}$ và đồng vị thứ hai $^{37}_{17}\ce{X}$. Biết nguyên tử khối trung bình của X là 35,48.\\
a) Xác định vị trí của X trong bảng tuần hoàn (ô, chu kì, nhóm)?\\
b) Tính \% số nguyên tử của đồng vị $^{35}_{17}\ce{X}$.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 3:}]
\textbf{a) (0,5 điểm)}
\begin{itemize}[leftmargin=1.0cm]
\item Số hiệu nguyên tử $Z = 17 \implies$ Cấu hình electron của X: $\ce{1s^2 2s^2 2p^6 3s^2 3p^5}$ (hoặc $\ce{[Ne] 3s^2 3p^5}$). \hfill \textit{(0,2 đ)}
\item Vị trí của X trong bảng tuần hoàn:
  \begin{itemize}
  \item Ô số: 17 (do $Z = 17$). \hfill \textit{(0,1 đ)}
  \item Chu kì: 3 (do có 3 lớp electron). \hfill \textit{(0,1 đ)}
  \item Nhóm: VIIA (do có 7e lớp ngoài cùng và là nguyên tố p). \hfill \textit{(0,1 đ)}
  \end{itemize}
\item (X là nguyên tố Chlorine - Cl).
\end{itemize}
\textbf{b) (0,5 điểm)}
\begin{itemize}[leftmargin=1.0cm]
\item Gọi phần trăm số nguyên tử của đồng vị $^{35}_{17}\ce{X}$ là $a\%$, suy ra của $^{37}_{17}\ce{X}$ là $(100 - a)\%$. \hfill \textit{(0,1 đ)}
\item Ta có phương trình nguyên tử khối trung bình:
$$\overline{A}_{\ce{X}} = \dfrac{35 \cdot a + 37 \cdot (100 - a)}{100} = 35,48 \hfill \textit{(0,2 đ)}$$
$$\iff 35a + 3700 - 37a = 3548 \iff -2a = -152 \iff a = 76\%. \hfill \textit{(0,1 đ)}$$
\item Vậy \% số nguyên tử của đồng vị $^{35}_{17}\ce{X}$ là $76\%$ (và của $^{37}_{17}\ce{X}$ là $24\%$). \hfill \textit{(0,1 đ)}
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{6}{\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\begin{center}
\textbf{---------- HẾT ----------}\\
\textit{\small (Thí sinh không được sử dụng tài liệu. Cán bộ coi thi không giải thích gì thêm)}
\end{center}

\end{document}
""")
    return "".join(tex)


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
    tex_hs = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_10_HaiBaTrung_HocSinh.tex")
    tex_gv = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_10_HaiBaTrung_GiaoVien.tex")

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
