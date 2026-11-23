# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Giữa Học Kì I Hóa Học 10 - THPT Gia Hội - Huế
Mã đề: 2101 - Năm học 2026 - 2027
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_10_GiaHoi_Ma2101_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_10_GiaHoi_Ma2101_GiaoVien.pdf (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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

SAN_PHAM_DIR = os.path.abspath("San_Pham")

def build_latex_content(is_teacher=False):
    tex = []
    tex.append(r"""\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{geometry}
\geometry{top=1.25cm,bottom=1.15cm,left=1.3cm,right=1.3cm,headheight=26pt,headsep=10pt,footskip=16pt}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{tabularx}
\usepackage{array}
\usepackage{colortbl}
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
\setlength{\parskip}{1.2pt}

\begin{document}

\begin{minipage}[t]{0.48\textwidth}
\textbf{SỞ GD\&ĐT THÀNH PHỐ HUẾ}\\
\textbf{TRƯỜNG THPT GIA HỘI}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 03 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}KIỂM TRA GIỮA KỲ I}\\
\textbf{NĂM HỌC 2026 -- 2027}\\
\textbf{MÔN: HOÁ HỌC 10}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{MÃ ĐỀ: 2101}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ và tên thí sinh:} \dotfill \textit{Lớp:} \dotfill \textit{Số báo danh:} \dotfill

\vspace{3pt}
\noindent
\textit{\footnotesize Cho nguyên tử khối của các nguyên tố: H: 1; C: 12; O: 16; N: 14; S: 32.}

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
\textbf{Đ/A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}A} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm - mỗi câu 1,0 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|l|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Nội dung câu hỏi} & \textbf{Ý a} & \textbf{Ý b} & \textbf{Ý c} & \textbf{Ý d} \\
\hline
\textbf{Câu 1} & Phổ khối lượng của nguyên tố R & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{SAI} & \textbf{\color{hfRed}ĐÚNG} \\
\hline
\textbf{Câu 2} & Sơ đồ cấu tạo nguyên tử X (Lithium-7) & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{SAI} & \textbf{\color{hfRed}ĐÚNG} \\
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
\textbf{Đáp án} & \textbf{\color{hfRed}81} & \textbf{\color{hfRed}35} & \textbf{\color{hfRed}2} & \textbf{\color{hfRed}16} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}4. Tóm tắt đáp số Phần Tự luận (3,0 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|p{11.5cm}|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Điểm} & \textbf{Đáp số cốt lõi} \\
\hline
\textbf{Câu 1} & 1,0 đ & a) $\overline{A}_{\text{Ne}} = 20,19\text{ amu}$. b) $m_{^{20}\text{Ne}} = 5,4\text{ gam}$. \\
\hline
\textbf{Câu 2} & 1,0 đ & a) Hạt electron ($e$), điện tích quy ước $-1$ (tuyệt đối: $-1,602\cdot 10^{-19}\text{ C}$). \\
& & b) Cột 1 ($^{28}_{13}\text{T}$): $Z=13, +13, N=15, A=28$; Cột 2 ($^{31}_{15}\text{P}$): $Z=15, +15, N=16, A=31$; \\
& & Cột 3: $^{13}_{6}\text{C}, +6, A=13$; Cột 4: $^{24}_{12}\text{Mg}, Z=12, N=12$. \\
\hline
\textbf{Câu 3} & 1,0 đ & a) S ($Z=16$): $1s^2 2s^2 2p^6 3s^2 3p^4$, có 2 electron độc thân ở phân lớp $3p$. \\
& & b) Mn ($Z=25$): $1s^2 2s^2 2p^6 3s^2 3p^6 3d^5 4s^2$; Ô 25, chu kì 4, nhóm VIIB. \\
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
{\color{mainGreen}\textbf{\Large I. TRẮC NGHIỆM (7,0 ĐIỂM)}}

\vspace{3pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. (12 câu -- 3,0 điểm)}}\\
\textit{\small Mỗi câu hỏi học sinh chỉ chọn một phương án. Mỗi phương án đúng được 0,25 điểm.}
\vspace{2pt}
""")

    # 12 CÂU PHẦN I
    # Câu 1
    tex.append(r"""
\begin{ex}
Phát biểu nào sau đây đúng?
\choice
{Proton và neutron là các hạt mang điện, electron là hạt không mang điện}
{Proton tạo nên lớp vỏ nguyên tử}
{\True Số lượng proton và electron trong nguyên tử là bằng nhau}
{Khối lượng của proton lớn hơn rất nhiều so với khối lượng của neutron}
\loigiai{
Trong nguyên tử trung hòa về điện:
\begin{itemize}[leftmargin=1.5cm,itemsep=0pt,topsep=1pt]
\item Hạt nhân chứa proton mang điện dương (+1) và neutron không mang điện. Lớp vỏ chứa electron mang điện âm (-1).
\item Số đơn vị điện tích hạt nhân bằng số proton và bằng số electron trong vỏ: $p = e$.
\item Khối lượng của proton xấp xỉ khối lượng neutron ($m_p \approx m_n \approx 1\text{ amu}$), và đều lớn hơn khối lượng electron rất nhiều ($\approx 1836$ lần).
\end{itemize}
\textbf{Đáp án đúng: C.}
}
\end{ex}
""")

    # Câu 2
    tex.append(r"""
\begin{ex}
Nguyên tử của nguyên tố X có tổng số electron trong các phân lớp p là 9. Nguyên tố X là nguyên tố
\choice
{khí hiếm}
{kim loại}
{kim loại hoặc phi kim}
{\True phi kim}
\loigiai{
Cấu hình electron nguyên tử của X được điền lần lượt theo mức năng lượng:
$$1s^2 2s^2 2p^6 3s^2 3p^3$$
Tổng số electron trên các phân lớp p là $6 + 3 = 9$ (thỏa mãn đề bài).\\
Số electron lớp ngoài cùng (lớp $n = 3$) là $2 + 3 = 5$ electron. Với 5 electron ở lớp ngoài cùng, X là nguyên tố phi kim (Phosphorus -- P, $Z = 15$).\\
\textbf{Đáp án đúng: D.}
}
\end{ex}
""")

    # Câu 3
    tex.append(r"""
\begin{ex}
Nguyên tố X có số hiệu nguyên tử bằng 9. Vị trí của X trong bảng tuần hoàn các nguyên tố hoá học là
\choice
{\True Ô 9, chu kì 2, nhóm VIIA}
{Ô 9, chu kì 5, nhóm VA}
{Ô 9, chu kì 2, nhóm VIIB}
{Ô 9, chu kì 3, nhóm VA}
\loigiai{
Nguyên tử X có $Z = 9$. Cấu hình electron: $1s^2 2s^2 2p^5$.\\
- Số hiệu nguyên tử $Z = 9 \Rightarrow$ Ô số 9 trong bảng tuần hoàn.\\
- Có 2 lớp electron $\Rightarrow$ Chu kì 2.\\
- Là nguyên tố p, có 7 electron ở lớp ngoài cùng ($2s^2 2p^5$) $\Rightarrow$ Nhóm VIIA (nguyên tố Fluorine, F).\\
\textbf{Đáp án đúng: A.}
}
\end{ex}
""")

    # Câu 4
    tex.append(r"""
\begin{ex}
Orbital p có dạng hình học là
\choice
{\True hình số tám nổi}
{hình số tám}
{hình tròn}
{hình cầu}
\loigiai{
Theo SGK Hóa học 10 (Chương trình GDPT 2018):
\begin{itemize}[leftmargin=1.5cm,itemsep=0pt,topsep=1pt]
\item Orbital s có dạng hình cầu.
\item Orbital p có dạng hình số tám nổi (gồm 3 orbital $p_x, p_y, p_z$ định hướng theo 3 trục không gian tương ứng $x, y, z$).
\end{itemize}
\textbf{Đáp án đúng: A.}
}
\end{ex}
""")

    # Câu 5
    tex.append(r"""
\begin{ex}
Hình vẽ sau mô tả thí nghiệm tìm ra một hạt cấu tạo nên nguyên tử.
\begin{center}
\includegraphics[width=6.0cm]{images_gk1_giahoi_2101/cau5_ong_tia_am_cuc_clean.png}
\end{center}
Đó là
\choice
{thí nghiệm tìm ra hạt nhân}
{\True thí nghiệm tìm ra electron}
{thí nghiệm tìm ra neutron}
{thí nghiệm tìm ra proton}
\loigiai{
Hình vẽ mô tả thí nghiệm phóng điện qua chất khí trong ống tia catot (ống Crookes) của nhà bác học J.J. Thomson (năm 1897). Chùm tia phát ra từ catot bị lệch về phía cực dương của điện trường chứng tỏ chùm tia gồm các hạt mang điện tích âm, từ đó phát hiện ra hạt electron.\\
\textbf{Đáp án đúng: B.}
}
\end{ex}
""")

    # Câu 6
    tex.append(r"""
\begin{ex}
Khí chlorine ($\ce{Cl2}$) được dùng phổ biến để diệt trùng nước sinh hoạt. Kí hiệu của nguyên tử chlorine có 17 proton và 20 neutron là
\choice
{$^{17}_{37}\ce{Cl}$}
{$^{35}_{17}\ce{Cl}$}
{$^{17}_{35}\ce{Cl}$}
{\True $^{37}_{17}\ce{Cl}$}
\loigiai{
Kí hiệu nguyên tử có dạng $^A_Z\ce{X}$:\\
- Số hiệu nguyên tử $Z = p = 17$.\\
- Số khối $A = Z + N = 17 + 20 = 37$.\\
Do đó, kí hiệu nguyên tử chlorine là $^{37}_{17}\ce{Cl}$.\\
\textbf{Đáp án đúng: D.}
}
\end{ex}
""")

    # Câu 7
    tex.append(r"""
\begin{ex}
Biết điện tích của 1 proton là $+1,602 \cdot 10^{-19}\text{ C}$. Nguyên tử A có điện tích hạt nhân là $+12,816 \cdot 10^{-19}\text{ C}$, trong hạt nhân của A chứa số hạt mang điện là
\choice
{6}
{15}
{13}
{\True 8}
\loigiai{
Trong hạt nhân nguyên tử, chỉ có proton là hạt mang điện dương ($+1,602 \cdot 10^{-19}\text{ C}$).\\
Số hạt mang điện trong hạt nhân của nguyên tử A là:
$$p = \frac{+12,816 \cdot 10^{-19}\text{ C}}{+1,602 \cdot 10^{-19}\text{ C}} = 8\text{ (hạt proton)}.$$
\textbf{Đáp án đúng: D.}
}
\end{ex}
""")

    # Câu 8
    tex.append(r"""
\begin{ex}
Cho nguyên tử có ký hiệu $^{39}_{19}\ce{Z}$, điều khẳng định nào sau đây \textbf{không} đúng?
\choice
{Nguyên tử khối của $^{39}_{19}\ce{Z}$ là 39 amu}
{Nguyên tử $^{39}_{19}\ce{Z}$ có 20 neutron}
{\True Nguyên tử $^{39}_{19}\ce{Z}$ có 20 electron}
{Nguyên tử $^{39}_{19}\ce{Z}$ có 19 proton}
\loigiai{
Từ kí hiệu $^{39}_{19}\ce{Z}$ ta có:
\begin{itemize}[leftmargin=1.5cm,itemsep=0pt,topsep=1pt]
\item Số proton $p = Z = 19$.
\item Trong nguyên tử trung hòa, số electron $e = p = 19$ (khẳng định C ghi 20 electron là \textbf{SAI}).
\item Số khối $A = 39 \Rightarrow$ Nguyên tử khối xấp xỉ 39 amu.
\item Số neutron $N = A - Z = 39 - 19 = 20$.
\end{itemize}
\textbf{Đáp án đúng: C.}
}
\end{ex}
""")

    # Câu 9
    tex.append(r"""
\begin{ex}
Trong bảng tuần hoàn hiện nay, các nguyên tố hoá học thuộc cùng một chu kì có cùng
\choice
{số electron hóa trị}
{\True số lớp electron}
{số proton trong nguyên tử}
{số electron lớp ngoài cùng}
\loigiai{
Chu kì là dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được sắp xếp theo chiều tăng dần của điện tích hạt nhân. Số thứ tự của chu kì bằng số lớp electron trong nguyên tử.\\
\textbf{Đáp án đúng: B.}
}
\end{ex}
""")

    # Câu 10
    tex.append(r"""
\begin{ex}
Để học tốt môn hoá học, theo em cần làm những gì sau đây?
\choice
{Vận dụng kiến thức để giải quyết một số tình huống trong thực tế}
{Đặt giả thuyết khoa học, xây dựng thí nghiệm để chứng minh, phân tích}
{Chịu khó quan sát và đặt câu hỏi}
{\True Tất cả các phương án trên}
\loigiai{
Phương pháp học tập và nghiên cứu Hóa học bao gồm: quan sát hiện tượng, đặt câu hỏi, nêu giả thuyết, tiến hành thí nghiệm kiểm chứng, phân tích kết quả và vận dụng kiến thức vào thực tế. Do đó cần thực hiện tất cả các phương án trên.\\
\textbf{Đáp án đúng: D.}
}
\end{ex}
""")

    # Câu 11
    tex.append(r"""
\begin{ex}
Cấu hình electron của nguyên tử Al ($Z = 13$) ở trạng thái cơ bản là
\choice
{\True $1s^2 2s^2 2p^6 3s^2 3p^1$}
{$1s^2 2s^2 2p^6 3s^1 3p^2$}
{$1s^2 2s^2 2p^6 3s^2$}
{$1s^2 2s^2 2p^6 3s^3$}
\loigiai{
Nguyên tử Aluminium (Al) có $Z = 13$. Cấu hình electron ở trạng thái cơ bản là:
$$1s^2 2s^2 2p^6 3s^2 3p^1 \quad ([\ce{Ne}]3s^2 3p^1).$$
\textbf{Đáp án đúng: A.}
}
\end{ex}
""")

    # Câu 12
    tex.append(r"""
\begin{ex}
Cho 4 nguyên tử: $^{16}_8\ce{X}$, $^{16}_7\ce{Y}$, $^{18}_9\ce{Z}$, $^{19}_9\ce{T}$. Cặp nguyên tử nào sau đây là đồng vị của nhau?
\choice
{\True Z, T}
{Y, Z}
{X, Z}
{X, Y}
\loigiai{
Đồng vị là những nguyên tử có cùng số proton (cùng số hiệu nguyên tử $Z$) nhưng khác nhau về số neutron (dẫn đến số khối $A$ khác nhau).\\
Trong 4 nguyên tử trên, Z và T đều có $Z = 9$ (nguyên tố Fluorine) nhưng có số khối khác nhau ($A_Z = 18$ và $A_T = 19$). Vì vậy, Z và T là hai đồng vị của nhau.\\
\textbf{Đáp án đúng: A.}
}
\end{ex}
""")

    # PHẦN II. Câu trắc nghiệm đúng sai
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN II. Câu trắc nghiệm đúng sai. (2 câu -- 2,0 điểm)}}\\
\textit{\small Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{2pt}
""")

    # Câu 1 Phần II
    tex.append(r"""
\begin{ex}
Cho phổ khối lượng của nguyên tố R ở hình dưới đây, biết nguyên tử khối trung bình của R là 6,92.
\begin{center}
\includegraphics[width=6.2cm]{images_gk1_giahoi_2101/pho_khoi_r.png}
\end{center}
\choiceTF
{\True Đồng vị $^x\ce{R}$ phổ biến hơn so với đồng vị còn lại}
{Trong tự nhiên nguyên tố R có ba đồng vị bền}
{Hai đồng vị bền trên có cùng số neutron}
{\True Số khối x của đồng vị còn lại là 7}
\loigiai{
Quan sát phổ khối lượng của nguyên tố R:
\begin{itemize}[leftmargin=1.5cm,itemsep=1pt,topsep=1pt]
\item[a)] \textbf{Đúng.} Pic ứng với đồng vị $^x\ce{R}$ có cường độ $92,41\%$, phổ biến hơn rất nhiều so với đồng vị còn lại ($7,59\%$).
\item[b)] \textbf{Sai.} Phổ khối lượng chỉ có 2 pic tín hiệu, chứng tỏ trong tự nhiên nguyên tố R chỉ có 2 đồng vị bền.
\item[c)] \textbf{Sai.} Hai đồng vị của cùng một nguyên tố có cùng số proton nhưng khác nhau về số neutron ($N_1 = 6 - Z \ne N_2 = x - Z$).
\item[d)] \textbf{Đúng.} Nguyên tử khối trung bình của R:
$$\overline{A}_{\text{R}} = \frac{6 \times 7,59 + x \times 92,41}{100} = 6,92 \Rightarrow 45,54 + 92,41x = 692 \Rightarrow x \approx 7.$$
Số khối của đồng vị còn lại là 7 ($^7\ce{Li}$).
\end{itemize}
}
\end{ex}
""")

    # Câu 2 Phần II
    tex.append(r"""
\begin{ex}
Cho sơ đồ của một nguyên tử X được biểu diễn như sau:
\begin{center}
\includegraphics[width=5.2cm]{images_gk1_giahoi_2101/so_do_nguyen_tu_x.png}
\end{center}
\choiceTF
{\True X có cấu hình electron là $1s^2 2s^1$}
{Số proton của X bằng 4}
{Trong X, điện tích hạt nhân là +7}
{\True X có số khối là 7}
\loigiai{
Quan sát sơ đồ cấu tạo nguyên tử X:
\begin{itemize}[leftmargin=1.5cm,itemsep=1pt,topsep=1pt]
\item Vỏ nguyên tử có 2 lớp electron: lớp thứ nhất có 2 electron, lớp thứ hai có 1 electron $\Rightarrow e = 3$. Cấu hình electron: $1s^2 2s^1$.
\item Hạt nhân nguyên tử có 3 proton và 4 neutron $\Rightarrow p = 3$; số neutron $n = 4$.
\item[a)] \textbf{Đúng.} Cấu hình electron của X là $1s^2 2s^1$.
\item[b)] \textbf{Sai.} Số proton của X bằng 3 (4 là số neutron).
\item[c)] \textbf{Sai.} Điện tích hạt nhân là $+3$ (7 là số khối).
\item[d)] \textbf{Đúng.} Số khối $A = p + n = 3 + 4 = 7$.
\end{itemize}
}
\end{ex}
""")

    # PHẦN III. Trả lời ngắn
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN III. Câu trắc nghiệm trả lời ngắn. (4 câu -- 2,0 điểm)}}\\
\textit{\small Học sinh trả lời từ câu 1 đến câu 4. Thí sinh điền kết quả vào ô trả lời tương ứng.}
\vspace{2pt}
""")

    # Câu 1 Phần III
    ans_p3_1 = r"\textbf{\color{hfRed}81}" if is_teacher else r"$\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$"
    c1_txt = r"""
\begin{ex}
Boron là nguyên tố có nhiều tác dụng đối với cơ thể người như: làm lành vết thương, điều hòa nội tiết sinh dục, chống viêm khớp,... Do ngọn lửa cháy có màu đặc biệt nên boron vô định hình được dùng làm pháo hoa. Boron có hai đồng vị là $^{10}\ce{B}$ và $^{11}\ce{B}$, nguyên tử khối trung bình là 10,81. Phần trăm số nguyên tử của đồng vị $^{11}\ce{B}$ là bao nhiêu? (Làm tròn đến phần nguyên)

\vspace{2pt}
\textbf{Trả lời:} %%ANS%%
\loigiai{
Gọi phần trăm số nguyên tử của đồng vị $^{11}\ce{B}$ là $x\%$ ($0 < x < 100$).\\
Phần trăm số nguyên tử của đồng vị $^{10}\ce{B}$ là $(100 - x)\%$.\\
Nguyên tử khối trung bình của Boron là:
$$\overline{A} = \frac{11 \cdot x + 10 \cdot (100 - x)}{100} = 10,81 \Rightarrow 11x + 1000 - 10x = 1081 \Rightarrow x = 81\%.$$
\textbf{Đáp án: 81}
}
\end{ex}
""".replace("%%ANS%%", ans_p3_1)
    tex.append(c1_txt)

    # Câu 2 Phần III
    ans_p3_2 = r"\textbf{\color{hfRed}35}" if is_teacher else r"$\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$"
    c2_txt = r"""
\begin{ex}
Tổng số hạt cơ bản trong nguyên tử của nguyên tố X là 114, trong đó số hạt mang điện nhiều hơn số hạt không mang điện là 26. Xác định số đơn vị điện tích hạt nhân của X.

\vspace{2pt}
\textbf{Trả lời:} %%ANS%%
\loigiai{
Gọi số proton, neutron, electron của nguyên tử X lần lượt là $p, n, e$. Trong nguyên tử trung hòa: $p = e$.\\
- Tổng số hạt cơ bản: $2p + n = 114$ \quad (1)\\
- Số hạt mang điện ($2p$) nhiều hơn số hạt không mang điện ($n$) là 26: $2p - n = 26$ \quad (2)\\
Cộng (1) và (2):
$$4p = 114 + 26 = 140 \Rightarrow p = 35.$$
Số đơn vị điện tích hạt nhân của X là $Z = p = 35$ (nguyên tố Bromine, Br).\\
\textbf{Đáp án: 35}
}
\end{ex}
""".replace("%%ANS%%", ans_p3_2)
    tex.append(c2_txt)

    # Câu 3 Phần III
    ans_p3_3 = r"\textbf{\color{hfRed}2}" if is_teacher else r"$\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$"
    c3_txt = r"""
\begin{ex}
Cho 5 nguyên tố có số hiệu nguyên tử lần lượt là 5, 8, 11, 15, 19. Trong số các nguyên tố trên có bao nhiêu nguyên tố ở chu kì 3 trong Bảng tuần hoàn các nguyên tố hoá học?

\vspace{2pt}
\textbf{Trả lời:} %%ANS%%
\loigiai{
Viết cấu hình electron của 5 nguyên tố:
\begin{itemize}[leftmargin=1.5cm,itemsep=0pt,topsep=1pt]
\item $Z = 5\text{ (B)}: 1s^2 2s^2 2p^1 \Rightarrow 2\text{ lớp e} \Rightarrow$ Chu kì 2.
\item $Z = 8\text{ (O)}: 1s^2 2s^2 2p^4 \Rightarrow 2\text{ lớp e} \Rightarrow$ Chu kì 2.
\item $Z = 11\text{ (Na)}: 1s^2 2s^2 2p^6 3s^1 \Rightarrow 3\text{ lớp e} \Rightarrow$ Chu kì 3.
\item $Z = 15\text{ (P)}: 1s^2 2s^2 2p^6 3s^2 3p^3 \Rightarrow 3\text{ lớp e} \Rightarrow$ Chu kì 3.
\item $Z = 19\text{ (K)}: 1s^2 2s^2 2p^6 3s^2 3p^6 4s^1 \Rightarrow 4\text{ lớp e} \Rightarrow$ Chu kì 4.
\end{itemize}
Có 2 nguyên tố thuộc chu kì 3 ($Z = 11$ và $Z = 15$).\\
\textbf{Đáp án: 2}
}
\end{ex}
""".replace("%%ANS%%", ans_p3_3)
    tex.append(c3_txt)

    # Câu 4 Phần III
    ans_p3_4 = r"\textbf{\color{hfRed}16}" if is_teacher else r"$\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$"
    c4_txt = r"""
\begin{ex}
Sulfur (S) là chất rắn, xốp, màu vàng hơi nhạt ở điều kiện thường. Sulfur và hợp chất của nó được sử dụng trong acquy, bột giặt, thuốc diệt nấm; do dễ cháy nên S còn được dùng để sản xuất các loại diêm, thuốc súng, pháo hoa,... Trong bảng tuần hoàn, nguyên tố S nằm ở chu kì 3, nhóm VIA. Nguyên tử của nguyên tố S nằm ở ô thứ bao nhiêu của bảng tuần hoàn các nguyên tố hóa học?

\vspace{2pt}
\textbf{Trả lời:} %%ANS%%
\loigiai{
- Nguyên tố S nằm ở chu kì 3 nên có 3 lớp electron ($n = 3$).\\
- Nằm ở nhóm VIA nên có 6 electron hóa trị ở lớp ngoài cùng ($3s^2 3p^4$).\\
Cấu hình electron đầy đủ của S: $1s^2 2s^2 2p^6 3s^2 3p^4$.\\
Tổng số electron là $2 + 2 + 6 + 2 + 4 = 16 \Rightarrow$ Số hiệu nguyên tử $Z = 16$.\\
Do số thứ tự ô nguyên tố bằng số hiệu nguyên tử $Z$, nguyên tử S nằm ở ô thứ 16.\\
\textbf{Đáp án: 16}
}
\end{ex}
""".replace("%%ANS%%", ans_p3_4)
    tex.append(c4_txt)

    # PHẦN IV. TỰ LUẬN
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\Large II. TỰ LUẬN (3,0 ĐIỂM)}}

\vspace{3pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN IV. Tự luận. (3 câu -- 3,0 điểm)}}\\
\textit{\small Học sinh làm bài theo hình thức tự luận.}
\vspace{2pt}
""")

    # Câu 1 Phần IV
    tex.append(r"""
\begin{ex}
Neon là nguyên tố khí hiếm được dùng nhiều trong các thiết bị điện, điện tử, bảng quảng cáo,... Nguyên tố Neon có 3 đồng vị cho ở bảng sau:
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|l|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Nguyên tử} & $^{20}\ce{Ne}$ & $^{21}\ce{Ne}$ & $^{22}\ce{Ne}$ \\
\hline
\textbf{\% số nguyên tử} & 90,0\% & 1,0\% & 9,0\% \\
\hline
\end{tabular}
\end{center}
a) Tính nguyên tử khối trung bình của Neon.\\
b) Tính khối lượng của đồng vị $^{20}\ce{Ne}$ có trong 6,057 gam neon nguyên chất.
""")
    if not is_teacher:
        tex.append(r"""
\vspace{1pt}
\noindent\textit{Bài làm:}\\
\multido{\n=1+1}{3}{\noindent\dotfill\\[3pt]}
""")
    else:
        tex.append(r"""
\loigiai{
a) Nguyên tử khối trung bình của Neon:
$$\overline{A}_{\text{Ne}} = \frac{20 \times 90,0 + 21 \times 1,0 + 22 \times 9,0}{100} = \frac{1800 + 21 + 198}{100} = \frac{2019}{100} = 20,19\text{ (amu)}.$$
b)
\begin{itemize}[leftmargin=1.5cm,itemsep=1pt,topsep=1pt]
\item Số mol của mẫu neon nguyên chất:
$$n_{\text{Ne}} = \frac{m_{\text{Ne}}}{\overline{A}_{\text{Ne}}} = \frac{6,057}{20,19} = 0,3\text{ (mol)}.$$
\item Vì đồng vị $^{20}\ce{Ne}$ chiếm $90,0\%$ tổng số nguyên tử neon:
$$n_{^{20}\text{Ne}} = 0,3 \times 90,0\% = 0,27\text{ (mol)}.$$
\item Khối lượng của đồng vị $^{20}\ce{Ne}$ là:
$$m_{^{20}\text{Ne}} = n_{^{20}\text{Ne}} \times M_{^{20}\text{Ne}} = 0,27 \times 20 = 5,4\text{ gam}.$$
\end{itemize}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần IV
    t_c3 = r"\textbf{\color{hfRed}$^{13}_{6}\ce{C}$}" if is_teacher else ""
    t_c4 = r"\textbf{\color{hfRed}$^{24}_{12}\ce{Mg}$}" if is_teacher else ""
    t_z1 = r"\textbf{\color{hfRed}13}" if is_teacher else ""
    t_z2 = r"\textbf{\color{hfRed}15}" if is_teacher else ""
    t_z4 = r"\textbf{\color{hfRed}12}" if is_teacher else ""
    t_q1 = r"\textbf{\color{hfRed}+13}" if is_teacher else ""
    t_q2 = r"\textbf{\color{hfRed}+15}" if is_teacher else ""
    t_q3 = r"\textbf{\color{hfRed}+6}" if is_teacher else ""
    t_n1 = r"\textbf{\color{hfRed}15}" if is_teacher else ""
    t_n2 = r"\textbf{\color{hfRed}16}" if is_teacher else ""
    t_n4 = r"\textbf{\color{hfRed}12}" if is_teacher else ""
    t_a1 = r"\textbf{\color{hfRed}28}" if is_teacher else ""
    t_a2 = r"\textbf{\color{hfRed}31}" if is_teacher else ""
    t_a3 = r"\textbf{\color{hfRed}13}" if is_teacher else ""

    c2_p4_tpl = r"""
\begin{ex}
a) Trong nguyên tử, cho biết loại hạt nào mang điện âm và điện tích của nó?\\
b) Hoàn thành bảng sau:
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|l|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Ký hiệu nguyên tử} & $^{28}_{13}\ce{T}$ & $^{31}_{15}\ce{P}$ & __C3__ & __C4__ \\
\hline
\textbf{Số hiệu nguyên tử} & __Z1__ & __Z2__ & 6 & __Z4__ \\
\hline
\textbf{Điện tích hạt nhân} & __Q1__ & __Q2__ & __Q3__ & +12 \\
\hline
\textbf{Số neutron} & __N1__ & __N2__ & 7 & __N4__ \\
\hline
\textbf{Số khối} & __A1__ & __A2__ & __A3__ & 24 \\
\hline
\end{tabular}
\end{center}
"""
    c2_p4_txt = c2_p4_tpl.replace("__C3__", t_c3).replace("__C4__", t_c4)\
                         .replace("__Z1__", t_z1).replace("__Z2__", t_z2).replace("__Z4__", t_z4)\
                         .replace("__Q1__", t_q1).replace("__Q2__", t_q2).replace("__Q3__", t_q3)\
                         .replace("__N1__", t_n1).replace("__N2__", t_n2).replace("__N4__", t_n4)\
                         .replace("__A1__", t_a1).replace("__A2__", t_a2).replace("__A3__", t_a3)
    tex.append(c2_p4_txt)
    if not is_teacher:
        tex.append(r"""
\vspace{1pt}
\noindent\textit{Bài làm:}\\
\multido{\n=1+1}{2}{\noindent\dotfill\\[3pt]}
""")
    else:
        tex.append(r"""
\loigiai{
a) Loại hạt mang điện âm là hạt \textbf{electron} ($e$).\\
- Điện tích quy ước (tương đối): $-1$.\\
- Điện tích tuyệt đối: $q_e = -1,602 \cdot 10^{-19}\text{ C}$.\\
b) Hoàn thành bảng:\\
Áp dụng mối liên hệ: $Z = p = e = \text{Số hiệu nguyên tử}$; Điện tích hạt nhân $= +Z$; $A = Z + N$; $N = A - Z$.\\
- Cột 1 ($^{28}_{13}\ce{T}$): $Z = 13$, điện tích hạt nhân là $+13$, $N = 28 - 13 = 15$, $A = 28$.\\
- Cột 2 ($^{31}_{15}\ce{P}$): $Z = 15$, điện tích hạt nhân là $+15$, $N = 31 - 15 = 16$, $A = 31$.\\
- Cột 3 (Cho $Z = 6, N = 7$): Ký hiệu là $^{13}_{6}\ce{C}$, điện tích hạt nhân là $+6$, $A = 6 + 7 = 13$.\\
- Cột 4 (Cho điện tích hạt nhân $+12$, $A = 24$): Ký hiệu là $^{24}_{12}\ce{Mg}$, $Z = 12$, $N = 24 - 12 = 12$.
}
""")
    tex.append(r"\end{ex}")

    # Câu 3 Phần IV
    tex.append(r"""
\begin{ex}
Sulfur ($Z = 16$) là nguyên tố phổ biến thứ 5 trên trái đất, thường xuất hiện ở các vùng suối nước nóng, mỏ núi lửa... Manganese ($Z = 25$) là nguyên tố được ứng dụng rộng rãi trong công nghiệp và đời sống.\\
a) Viết cấu hình electron của nguyên tử Sulfur, phân bố các electron trên các orbital.\\
b) Viết cấu hình electron của nguyên tử Manganese. Xác định vị trí (ô, chu kì, nhóm) của nguyên tử Manganese trong bảng tuần hoàn.
""")
    if not is_teacher:
        tex.append(r"""
\vspace{1pt}
\noindent\textit{Bài làm:}\\
\multido{\n=1+1}{3}{\noindent\dotfill\\[3pt]}
""")
    else:
        tex.append(r"""
\loigiai{
a) Nguyên tử Sulfur ($Z = 16$):
\begin{itemize}[leftmargin=1.5cm,itemsep=1pt,topsep=1pt]
\item Cấu hình electron: $1s^2 2s^2 2p^6 3s^2 3p^4$ (hoặc $[\ce{Ne}] 3s^2 3p^4$).
\item Sự phân bố electron vào các ô orbital:
  \begin{itemize}[itemsep=1pt]
  \item Phân lớp $1s$: 1 ô orbital chứa 2 electron ghép đôi ($\uparrow\downarrow$).
  \item Phân lớp $2s$: 1 ô orbital chứa 2 electron ghép đôi ($\uparrow\downarrow$).
  \item Phân lớp $2p$: 3 ô orbital bão hòa chứa 6 electron ghép đôi ($\uparrow\downarrow$ $\uparrow\downarrow$ $\uparrow\downarrow$).
  \item Phân lớp $3s$: 1 ô orbital chứa 2 electron ghép đôi ($\uparrow\downarrow$).
  \item Phân lớp $3p$: 3 ô orbital chứa 4 electron gồm 1 ô ghép đôi và 2 ô chứa electron độc thân ($\uparrow\downarrow$ $\uparrow$ $\uparrow$).
  \end{itemize}
  $\Rightarrow$ Nguyên tử Sulfur có 2 electron độc thân ở phân lớp $3p$.
\end{itemize}

b) Nguyên tử Manganese ($Z = 25$):
\begin{itemize}[leftmargin=1.5cm,itemsep=1pt,topsep=1pt]
\item Mức năng lượng: $1s^2 2s^2 2p^6 3s^2 3p^6 4s^2 3d^5$.
\item Cấu hình electron theo thứ tự lớp: $1s^2 2s^2 2p^6 3s^2 3p^6 3d^5 4s^2$ (hoặc $[\ce{Ar}] 3d^5 4s^2$).
\item Vị trí trong bảng tuần hoàn:
  \begin{itemize}[itemsep=1pt]
  \item \textbf{Ô nguyên tố:} Ô thứ 25 (vì có $Z = 25$).
  \item \textbf{Chu kì:} Chu kì 4 (vì có 4 lớp electron).
  \item \textbf{Nhóm:} Nhóm VIIB (vì là nguyên tố khối d, electron hóa trị điền vào $3d$ và $4s$ với tổng số electron là $5 + 2 = 7$).
  \end{itemize}
\end{itemize}
}
""")
    tex.append(r"\end{ex}")

    tex.append(r"""
\vspace{4pt}
\begin{center}
\textbf{---------- HẾT ----------}
\end{center}

\end{document}
""")

    return "".join(tex)

def compile_latex(filename_base):
    tex_file = os.path.join(SAN_PHAM_DIR, f"{filename_base}.tex")
    pdf_file = os.path.join(SAN_PHAM_DIR, f"{filename_base}.pdf")

    print(f"--- Đang biên dịch {filename_base}.tex ---")
    cmd = ["pdflatex", "-interaction=nonstopmode", f"{filename_base}.tex"]
    res = subprocess.run(cmd, cwd=SAN_PHAM_DIR, capture_output=True, text=True)

    if os.path.exists(pdf_file):
        print(f"Biên dịch thành công: {pdf_file}")
    else:
        print(f"Biên dịch thất bại: {filename_base}.tex")
        print("Chi tiết lỗi:\n", res.stdout[-1000:])

def cleanup_aux_files():
    patterns = ["*.aux", "*.log", "*.thm", "ans-*.tex", "*.out"]
    for p in patterns:
        for f in glob.glob(os.path.join(SAN_PHAM_DIR, p)):
            try:
                os.remove(f)
            except Exception:
                pass
    print("Đã dọn dẹp các tệp tạm.")

if __name__ == "__main__":
    base_hs = "De_Kiem_Tra_GK1_Hoa_10_GiaHoi_Ma2101_HocSinh"
    base_gv = "De_Kiem_Tra_GK1_Hoa_10_GiaHoi_Ma2101_GiaoVien"

    # Ghi file tex
    with open(os.path.join(SAN_PHAM_DIR, f"{base_hs}.tex"), "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))

    with open(os.path.join(SAN_PHAM_DIR, f"{base_gv}.tex"), "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))

    # Biên dịch
    compile_latex(base_hs)
    compile_latex(base_gv)

    # Dọn dẹp
    cleanup_aux_files()
