# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Giữa Kì 1 Hóa Học 11 - THPT Nguyễn Chí Thanh - Mã đề 1002
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_11_NguyenChiThanh_Ma1002_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_11_NguyenChiThanh_Ma1002_GiaoVien.pdf (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
\textbf{SỞ GIÁO DỤC VÀ ĐÀO TẠO TP. HUẾ}\\
\textbf{TRƯỜNG THPT NGUYỄN CHÍ THANH}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 04 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}ĐỀ KIỂM TRA GIỮA HỌC KỲ I – NĂM HỌC 2026 – 2027}\\
\textbf{MÔN: HOÁ HỌC. LỚP 11}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{Mã đề: 1002}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ, tên thí sinh:} \dotfill \textit{Lớp:} \dotfill \textit{Số báo danh:} \dotfill

\vspace{2pt}
\noindent
\textit{\footnotesize (Cho nguyên tử khối của: $\ce{H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; S = 32; Cl = 35{,}5; Zn = 65; Ba = 137}$).}

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
\textbf{\color{mainGreen}1. Phần I: Trắc nghiệm nhiều lựa chọn (4,0 điểm - mỗi câu đúng 0,25 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\small
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} \\
\hline
\textbf{Đ/A} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Lệnh a)} & \textbf{Lệnh b)} & \textbf{Lệnh c)} & \textbf{Lệnh d)} \\
\hline
\textbf{Câu 1} & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{solBlue}SAI} & \textbf{\color{hfRed}ĐÚNG} \\
\hline
\textbf{Câu 2} & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{solBlue}SAI} & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{solBlue}SAI} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}3. Phần III: Trắc nghiệm Trả lời ngắn (1,0 điểm - mỗi câu đúng 0,25 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu hỏi} & \textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}1,7} & \textbf{\color{hfRed}12} & \textbf{\color{hfRed}4} & \textbf{\color{hfRed}3} \\
\hline
\end{tabular}
\end{center}

\vspace{4pt}
\hrule color{hfGreen} height 0.5pt
""")

    # ==========================================
    # PHẦN I
    # ==========================================
    tex.append(r"""
\vspace{4pt}
\noindent
\textbf{\large\color{mainGreen}PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn: (16 câu - 4 điểm)}

\textit{Thí sinh trả lời từ câu 1 đến câu 16. Mỗi câu thí sinh chỉ chọn một phương án.}
\vspace{4pt}
""")

    # Câu 1
    c1_sol = r"""\begin{loigiai}
Hiện tượng nóng lên toàn cầu là do sự gia tăng nồng độ của các khí nhà kính ($\ce{CO2}, \ce{CH4}, \ce{N2O},\dots$) giữ nhiệt trong khí quyển, không phải là tác hại chính của mưa acid. Các tác hại chính của mưa acid gồm hủy hoại rừng cây, làm chết sinh vật thủy sinh, và ăn mòn các công trình kiến trúc.\\
$\rightarrow$ \textbf{Chọn đáp án A.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C1]
Hiện tượng nào sau đây không phải là tác hại chính của mưa acid?
\choice
{\True Làm gia tăng hiện tượng nóng lên toàn cầu}
{Làm hủy hoại rừng và thảm thực vật cây cối}
{Gây ăn mòn và hủy hoại các công trình kiến trúc}
{Gây hại và làm suy giảm sinh vật dưới nước}
""" + c1_sol + r"""
\end{ex}
""")

    # Câu 2
    c2_sol = r"""\begin{loigiai}
Xét pH của các dung dịch:
\begin{itemize}[leftmargin=0.8cm]
\item Dung dịch $\ce{HCl}$ $0{,}1\,\text{M}$: acid mạnh phân li hoàn toàn $\Rightarrow [\ce{H+}] = 0{,}1\,\text{M} \Rightarrow \text{pH} = 1$.
\item Dung dịch $\ce{CH3COOH}$ $0{,}1\,\text{M}$: acid yếu phân li thuận nghịch không hoàn toàn $\Rightarrow [\ce{H+}] < 0{,}1\,\text{M} \Rightarrow \text{pH} > 1$ (thực tế $\text{pH} \approx 2{,}88$).
\item Dung dịch $\ce{NaCl}$ $0{,}1\,\text{M}$: muối trung tính $\Rightarrow \text{pH} = 7$.
\item Dung dịch $\ce{NaOH}$ $0{,}01\,\text{M}$: base mạnh $\Rightarrow [\ce{OH-}] = 0{,}01\,\text{M} \Rightarrow \text{pOH} = 2 \Rightarrow \text{pH} = 12$.
\end{itemize}
Vậy dung dịch có pH nhỏ nhất là dung dịch $\ce{HCl}$ $0{,}1\,\text{M}$ ($\text{pH} = 1$).\\
$\rightarrow$ \textbf{Chọn đáp án A.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C2]
pH của dung dịch nào sau đây có giá trị nhỏ nhất?
\choice
{\True Dung dịch $\ce{HCl}$ $0{,}1\text{M}$.}
{Dung dịch $\ce{CH3COOH}$ $0{,}1\text{M}$.}
{Dung dịch $\ce{NaCl}$ $0{,}1\text{M}$.}
{Dung dịch $\ce{NaOH}$ $0{,}01\text{M}$.}
""" + c2_sol + r"""
\end{ex}
""")

    # Câu 3
    c3_sol = r"""\begin{loigiai}
Bản chất của phản ứng chuẩn độ acid mạnh ($\ce{HCl}$) bằng base mạnh ($\ce{NaOH}$) là phản ứng trung hòa ion giữa $\ce{H+}$ và $\ce{OH-}$ tạo thành nước: $\ce{H+ + OH- -> H2O}$.\\
$\rightarrow$ \textbf{Chọn đáp án D.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C3]
Chuẩn độ dung dịch $\ce{NaOH}$ chưa biết chính xác nồng độ (biết nồng độ trong khoảng gần với $0{,}1\,\text{M}$) bằng dung dịch chuẩn $\ce{HCl}$ $0{,}1\,\text{M}$ với chỉ thị phenolphthalein. Phương trình thể hiện bản chất của phản ứng của quá trình chuẩn độ là ?
\choice
{$\ce{NaCl -> Na+ + Cl-}$}
{$\ce{Na+ + Cl- -> NaCl}$}
{$\ce{H2O <=> H+ + OH-}$}
{\True $\ce{H+ + OH- -> H2O}$}
""" + c3_sol + r"""
\end{ex}
""")

    # Câu 4
    c4_sol = r"""\begin{loigiai}
Số mol khí $\ce{NO}$ ở điều kiện chuẩn: $n_{\ce{NO}} = \dfrac{4{,}958}{24{,}79} = 0{,}2\,\text{mol}$.\\
Quá trình nhường -- nhận electron:
$$\ce{Al -> Al^{3+} + 3e}\quad\quad\quad \ce{N^{+5} + 3e -> N^{+2} (NO)}$$
Bảo toàn electron: $3n_{\ce{Al}} = 3n_{\ce{NO}} \Rightarrow n_{\ce{Al}} = n_{\ce{NO}} = 0{,}2\,\text{mol}$.\\
Khối lượng $\ce{Al}$: $m = 0{,}2 \times 27 = 5{,}4\,\text{gam}$.\\
$\rightarrow$ \textbf{Chọn đáp án C.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C4]
Cho $m$ gam $\ce{Al}$ phản ứng hoàn toàn với dung dịch $\ce{HNO3}$ loãng (dư), thu được $4{,}958$ lít khí $\ce{NO}$ (đkc, sản phẩm khử duy nhất). Giá trị của $m$ là
\choice
{4,5.}
{8,1}
{\True 5,4.}
{2,7}
""" + c4_sol + r"""
\end{ex}
""")

    # Câu 5
    c5_sol = r"""\begin{loigiai}
Trong phản ứng với $\ce{O2}$: $\ce{N2 + O2 <=> 2NO}$ (ở khoảng $3000^\circ\text{C}$ hoặc có tia lửa điện).\\
Số oxi hóa của $\ce{N}$ tăng từ $0$ lên $+2$, do đó $\ce{N2}$ thể hiện tính khử.\\
Trong phản ứng với $\ce{H2}, \ce{Li}, \ce{Mg}$, số oxi hóa của $\ce{N}$ giảm từ $0$ xuống $-3$ nên $\ce{N2}$ thể hiện tính oxi hóa.\\
$\rightarrow$ \textbf{Chọn đáp án C.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C5]
Nitrogen ($\ce{N2}$) thể hiện tính khử trong phản ứng với chất nào trong các chất sau?
\choice
{$\ce{H2}$.}
{$\ce{Li}$.}
{\True $\ce{O2}$.}
{$\ce{Mg}$}
""" + c5_sol + r"""
\end{ex}
""")

    # Câu 6
    c6_sol = r"""\begin{loigiai}
Phân tử ammonia ($\ce{NH3}$) có cặp electron tự do trên nguyên tử nitrogen có khả năng nhận proton $\ce{H+}$ nên có tính base yếu. Trên $\ce{N}$ có mức oxi hóa $-3$ thấp nhất nên dễ nhường electron thể hiện tính khử mạnh.\\
$\rightarrow$ \textbf{Chọn đáp án D.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C6]
Tính chất hóa học cơ bản của ammonia ($\ce{NH3}$) là
\choice
{tính base mạnh, tính khử.}
{tính base yếu, tính oxi hóa.}
{tính base mạnh, tính oxi hóa.}
{\True tính khử mạnh, tính base yếu.}
""" + c6_sol + r"""
\end{ex}
""")

    # Câu 7
    c7_sol = r"""\begin{loigiai}
Theo nguyên lí Le Chatelier, các yếu tố làm chuyển dịch cân bằng hóa học là: nồng độ, nhiệt độ và áp suất (đối với hệ có chất khí). Chất xúc tác và diện tích bề mặt chỉ làm tăng tốc độ đạt đến cân bằng chứ không làm chuyển dịch cân bằng.\\
$\rightarrow$ \textbf{Chọn đáp án B.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C7]
Các yếu tố có thể ảnh hưởng đến sự chuyển dịch cân bằng hoá học là :
\choice
{nồng độ, áp suất và diện tích bề mặt.}
{\True nồng độ, nhiệt độ và áp suất.}
{áp suất, nhiệt độ và chất xúc tác.}
{nồng độ, nhiệt độ và chất xúc tác.}
""" + c7_sol + r"""
\end{ex}
""")

    # Câu 8
    c8_sol = r"""\begin{loigiai}
Trong $\ce{HNO3}$: $\ce{H}$ mang số oxi hóa $+1$, $\ce{O}$ mang $-2$. Gọi $x$ là số oxi hóa của $\ce{N}$:\\
$(+1) + x + 3 \times (-2) = 0 \Rightarrow x = +5$.\\
$\rightarrow$ \textbf{Chọn đáp án B.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C8]
Số oxi hoá của nitrogen ($\ce{N}$) trong phân tử nitric acid ($\ce{HNO3}$) là
\choice
{$+4$.}
{\True $+5$.}
{$+2$.}
{$+3$.}
""" + c8_sol + r"""
\end{ex}
""")

    # Câu 9
    c9_sol = r"""\begin{loigiai}
Hiện tượng phú dưỡng xảy ra khi nguồn nước bị ô nhiễm dư thừa các nguyên tố dinh dưỡng như nitrogen ($\ce{N}$) và phosphorus ($\ce{P}$). Hành vi đánh bắt thủy sản bằng xung điện và chất nổ là hành vi tận diệt môi sinh, không phải là nguyên nhân chính gây tích tụ dinh dưỡng dẫn đến hiện tượng phú dưỡng.\\
$\rightarrow$ \textbf{Chọn đáp án B.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C9]
Hành động nào sau đây của con người không phải là nguyên nhân chính gây ra hiện tượng phú dưỡng?
\choice
{Xả thải nước thải chăn nuôi, công nghiệp chứa hàm lượng cao các nguyên tố dinh dưỡng ($\ce{N, P}$).}
{\True Đánh bắt thuỷ hải sản bằng chích điện và chất nổ.}
{Sử dụng dư thừa lượng phân bón chứa ion $\ce{PO4^{3-}}$ ngấm vào nguồn nước mặt.}
{Sử dụng dư thừa lượng phân bón chứa các ion $\ce{NO2^-}, \ce{NO3^-}$ trong sản xuất nông nghiệp.}
""" + c9_sol + r"""
\end{ex}
""")

    # Câu 10
    c10_sol = r"""\begin{loigiai}
Theo thuyết Brønsted -- Lowry, acid là chất cho proton ($\ce{H+}$), base là chất nhận proton ($\ce{H+}$).\\
$\rightarrow$ \textbf{Chọn đáp án A.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C10]
Phát biểu nào sau đây về acid – base theo thuyết Bronsted – Lowry là đúng?
\choice
{\True Acid là chất cho proton ($\ce{H+}$).}
{Base là chất nhận electron.}
{Base là chất cho proton ($\ce{H+}$).}
{Acid là chất nhận proton ($\ce{H+}$).}
""" + c10_sol + r"""
\end{ex}
""")

    # Câu 11
    c11_sol = r"""\begin{loigiai}
$\ce{HNO3}$ đóng vai trò là chất oxi hóa khi số oxi hóa của nitrogen giảm từ $+5$ xuống các mức thấp hơn. Xét từng chất:
\begin{enumerate}
\item $\ce{Cu}$: $\ce{Cu + 4HNO3 -> Cu(NO3)2 + 2NO2 + 2H2O}$ ($\ce{N^{+5}}$ nhận electron $\Rightarrow \ce{HNO3}$ là chất oxi hóa).
\item $\ce{NaOH}$: $\ce{NaOH + HNO3 -> NaNO3 + H2O}$ (phản ứng trung hòa trao đổi, $\ce{HNO3}$ chỉ đóng vai trò acid).
\item $\ce{C}$: $\ce{C + 4HNO3 -> CO2 + 4NO2 + 2H2O}$ ($\ce{C^0}$ bị oxi hóa lên $\ce{C^{+4}} \Rightarrow \ce{HNO3}$ là chất oxi hóa).
\item $\ce{FeO}$: $\ce{3FeO + 10HNO3 -> 3Fe(NO3)3 + NO + 5H2O}$ ($\ce{Fe^{+2}}$ bị oxi hóa lên $\ce{Fe^{+3}} \Rightarrow \ce{HNO3}$ là chất oxi hóa).
\item $\ce{Na2CO3}$: $\ce{Na2CO3 + 2HNO3 -> 2NaNO3 + CO2 + H2O}$ (phản ứng trao đổi ion, $\ce{HNO3}$ chỉ đóng vai trò acid).
\end{enumerate}
Vậy có 3 phản ứng $\ce{HNO3}$ đóng vai trò chất oxi hóa (với $\ce{Cu, C, FeO}$).\\
$\rightarrow$ \textbf{Chọn đáp án D.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C11]
Cho dung dịch nitric acid ($\ce{HNO3}$) tác dụng với các chất sau: $\ce{Cu, NaOH, C, FeO, Na2CO3}$. Khi phản ứng xảy ra, có bao nhiêu phản ứng trong đó nitric acid ($\ce{HNO3}$) đóng vai trò oxi hóa ?
\choice
{4.}
{5.}
{2.}
{\True 3.}
""" + c11_sol + r"""
\end{ex}
""")

    # Câu 12
    c12_sol = r"""\begin{loigiai}
Sự dẫn điện của dung dịch muối, acid, base là do trong dung dịch nước chúng có khả năng phân li ra các ion dương và ion âm mang điện tích tự do chuyển động (chất điện li).\\
$\rightarrow$ \textbf{Chọn đáp án D.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C12]
Các dung dịch của muối, acid, base dẫn điện là do
\choice
{phân tử của chúng dẫn được điện.}
{có sự di chuyển của electron tạo thành dòng electron.}
{các ion hợp phần có khả năng dẫn điện.}
{\True muối, acid, base có khả năng phân li ra ion khi tan trong nước.}
""" + c12_sol + r"""
\end{ex}
""")

    # Câu 13
    c13_sol = r"""\begin{loigiai}
Theo định luật tác dụng khối lượng, đối với cân bằng: $\ce{H2(g) + I2(g) <=> 2HI(g)}$, hằng số cân bằng $K_\text{c}$ có biểu thức:\\
$$K_\text{c} = \dfrac{[\ce{HI}]^2}{[\ce{H2}]\cdot[\ce{I2}]}$$
$\rightarrow$ \textbf{Chọn đáp án D.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C13]
Hệ phản ứng sau ở trạng thái cân bằng : $\ce{H2(g) + I2(g) <=> 2HI(g)}$\\
Biểu thức hằng số cân bằng của phản ứng trên là
\choice
{$K_\text{c} = \dfrac{[2\ce{HI}]}{[\ce{H2}]\cdot[\ce{I2}]}$}
{$K_\text{c} = \dfrac{[\ce{H2}]\cdot[\ce{I2}]}{2[\ce{HI}]}$}
{$K_\text{c} = \dfrac{[\ce{H2}]\cdot[\ce{I2}]}{[\ce{HI}]^2}$}
{\True $K_\text{c} = \dfrac{[\ce{HI}]^2}{[\ce{H2}]\cdot[\ce{I2}]}$}
""" + c13_sol + r"""
\end{ex}
""")

    # Câu 14
    c14_sol = r"""\begin{loigiai}
Muối ammonium ($\ce{NH4NO3, (NH4)2SO4, NH4Cl},\dots$) được ứng dụng rộng rãi và quan trọng nhất là dùng làm phân bón hóa học (phân đạm), cung cấp dinh dưỡng nitrogen cho cây trồng.\\
$\rightarrow$ \textbf{Chọn đáp án C.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C14]
Ứng dụng phổ biến và quan trọng nhất của muối ammonium trong đời sống và sản xuất là
\choice
{sản xuất nhiên liệu xăng sinh học và dầu diesel.}
{sản xuất xi măng và vật liệu chịu lửa trong xây dựng.}
{\True sản xuất phân bón hoá học, cung cấp nguyên tố nitrogen dinh dưỡng cho cây trồng.}
{điều chế kim loại nguyên chất bằng phương pháp nhiệt luyện.}
""" + c14_sol + r"""
\end{ex}
""")

    # Câu 15
    c15_sol = r"""\begin{loigiai}
Khí gây cười (funky ball) là dinitrogen monoxide có công thức hóa học là $\ce{N2O}$. Lạm dụng khí này gây tổn thương thần kinh và nguy cơ tử vong.\\
$\rightarrow$ \textbf{Chọn đáp án D.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C15]
Bóng cười (hay "funky ball") là một quả bóng được bơm đầy khí X. Khí X có vị ngọt nhẹ, không màu, gây tác động hưng phấn, ảo giác cười cho người hít phải. Lạm dụng bóng cười gây tác hại nghiêm trọng đến hệ thần kinh, tim mạch, tụt huyết áp, co giật và nguy cơ tử vong. Công thức hoá học của khí X là
\choice
{$\ce{O2}$.}
{$\ce{N2}$.}
{$\ce{NO}$.}
{\True $\ce{N2O}$.}
""" + c15_sol + r"""
\end{ex}
""")

    # Câu 16
    c16_sol = r"""\begin{loigiai}
Phản ứng thuận nghịch là phản ứng trong cùng điều kiện xác định diễn ra theo cả hai chiều ngược nhau, biểu diễn bằng mũi tên thuận nghịch $\rightleftharpoons$:\\
$\ce{N2 + 3H2 <=> 2NH3}$.\\
$\rightarrow$ \textbf{Chọn đáp án D.}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-C16]
Phản ứng nào sau đây là phản ứng thuận nghịch ?
\choice
{$\ce{Fe + 2HCl -> FeCl2 + H2}$.}
{$\ce{H2 + Cl2 -> 2HCl}$.}
{$\ce{2H2 + O2 -> 2H2O}$.}
{\True $\ce{N2 + 3H2 <=> 2NH3}$.}
""" + c16_sol + r"""
\end{ex}
""")

    # ==========================================
    # PHẦN II. TRẮC NGHIỆM ĐÚNG SAI
    # ==========================================
    tex.append(r"""
\vspace{6pt}
\noindent
\textbf{\large\color{mainGreen}PHẦN II. Câu trắc nghiệm đúng sai: (2 câu - 2 điểm)}

\textit{Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu thí sinh chọn đúng hoặc sai.}
\vspace{4pt}
""")

    # Câu 1 Phần II
    p2_c1_sol = r"""\begin{loigiai}
\begin{itemize}[leftmargin=1.5cm]
\item[\textbf{a) ĐÚNG:}] Khi thêm acid $\ce{HCl}$, nồng độ $\ce{H+}$ tăng lên. Theo nguyên lí Le Chatelier, cân bằng chuyển dịch theo chiều làm giảm nồng độ $\ce{H+}$, tức là chiều nghịch, ngăn ngừa sự thủy phân tạo kết tủa keo $\ce{Fe(OH)3}$ giúp bảo quản dung dịch trong suốt.
\item[\textbf{b) ĐÚNG:}] Trong dung dịch nước, phức aqua $\ce{[Fe(H2O)6]^{3+}}$ nhường proton $\ce{H+}$ cho $\ce{H2O}$ để tạo $\ce{H3O+}$, do đó $\ce{Fe^{3+}}$ đóng vai trò là một acid theo Brønsted -- Lowry.
\item[\textbf{c) SAI:}] Thêm $\ce{NaOH}$, ion $\ce{OH-}$ trung hòa $\ce{H+}$ ($\ce{H+ + OH- -> H2O}$) làm giảm nồng độ $\ce{H+}$. Cân bằng chuyển dịch theo chiều tạo thêm $\ce{H+}$, tức là \textbf{CHIỀU THUẬN} (tạo kết tủa $\ce{Fe(OH)3}$), không phải chiều nghịch.
\item[\textbf{d) ĐÚNG:}] Cân bằng thủy phân $\ce{Fe^{3+}}$ giải phóng nhiều ion $\ce{H+}$, làm tăng nồng độ $\ce{H+}$ và làm giảm $\text{pH}$ của đất ($\text{pH} < 7$), nên các loại đất giàu sắt là đất chua.
\end{itemize}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-P2-C1]
Trong dung dịch muối $\ce{Fe^{3+}}$ tồn tại cân bằng hóa học sau:
$$\ce{Fe^{3+} + 3H2O <=> Fe(OH)3 v + 3H+}$$
\choiceTF
{\True Trong phòng thí nghiệm, để bảo quản dung dịch $\ce{Fe^{3+}}$, người ta thường thêm vào bình đựng vài giọt dung dịch acid $\ce{HCl}$ để cân bằng chuyển dịch theo chiều nghịch}
{\True Dựa vào thuyết acid – base của Brønsted – Lowry, $\ce{Fe^{3+}}$ đóng vai trò acid}
{Thêm vào dung dịch muối $\ce{Fe^{3+}}$ vài giọt dung dịch $\ce{NaOH}$ thì cân bằng chuyển dịch theo chiều nghịch}
{\True Các loại đất có chứa nhiều ion $\ce{Fe^{3+}}$ có giá trị pH thấp hay còn gọi là đất chua}
""" + p2_c1_sol + r"""
\end{ex}
""")

    # Câu 2 Phần II
    p2_c2_sol = r"""\begin{loigiai}
\begin{itemize}[leftmargin=1.5cm]
\item[\textbf{a) ĐÚNG:}] Ở phản ứng (1): $\ce{N2 + O2 <=> 2NO}$. Số oxi hóa của $\ce{N}$ tăng từ $0$ lên $+2$, do đó $\ce{N2}$ đóng vai trò là chất khử.
\item[\textbf{b) SAI:}] Tính toán lượng khí $\ce{N2}$ cần dùng:
\begin{itemize}
\item Số mol $\ce{NO3-}$ cần tạo ra: $n_{\ce{NO3-}} = \dfrac{124\,000}{62} = 2000\,\text{mol} = 2\,\text{kmol}$.
\item Sơ đồ: $\ce{N2 -> 2NO3-}$. Theo lí thuyết ($H = 100\%$): $n_{\ce{N2}} = \dfrac{2000}{2} = 1000\,\text{mol}$.
\item Với hiệu suất $H = 5\% = 0{,}05$: $n_{\ce{N2}\,\text{thực tế}} = \dfrac{1000}{0{,}05} = 20\,000\,\text{mol}$.
\item Thể tích khí $\ce{N2}$ ở đkc: $V_{\ce{N2}} = 20\,000 \times 24{,}79 = 495\,800\,\text{lít} = \mathbf{495{,}8\,\text{m}^3}$ (không phải $991{,}6\,\text{m}^3$).
\end{itemize}
\item[\textbf{c) ĐÚNG:}] Trong động cơ buồng đốt của phương tiện giao thông, nhiệt độ cao làm $\ce{N2}$ và $\ce{O2}$ trong không khí phản ứng tạo thành $\ce{NO}$, sau đó tiếp tục tạo thành $\ce{HNO3}$ theo đúng sơ đồ trên.
\item[\textbf{d) SAI:}] Quá trình (4) là sự điện li của acid mạnh $\ce{HNO3}$ trong nước: $\ce{HNO3 -> H+ + NO3-}$, tự xảy ra hoàn toàn trong dung dịch nước mưa mà \textbf{không cần} năng lượng sấm sét. Sấm sét chỉ cần thiết cho phản ứng (1).
\end{itemize}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-P2-C2]
Trong những cơn mưa dông kèm sấm sét, sự chuyển hoá của nitrogen đơn chất xảy ra theo sơ đồ sau:
$$\ce{N2 ->[+O_2][(1)] NO ->[+O_2][(2)] NO2 ->[+O_2 + H_2O][(3)] HNO3 ->[(4)] H+ + NO3^-}$$
Ion nitrate ($\ce{NO3-}$) là một dạng đạm cây trồng hấp thụ được để sinh trưởng, phát triển.
\choiceTF
{\True Trong phản ứng (1) $\ce{N2}$ đóng vai trò là chất khử}
{Để tạo ra được $124\,\text{kg}$ ion nitrate cần dùng ở điều kiện chuẩn là $991{,}6\,\text{m}^3$ khí $\ce{N2}$. Biết hiệu suất của toàn bộ quá trình là $5\%$}
{\True Sơ đồ cũng áp dụng để giải thích sự tạo thành $\ce{HNO3}$ từ hoạt động đốt cháy nhiên liệu bằng không khí của phương tiện giao thông vận tải}
{Năng lượng trong sấm sét gây ra sự phân li của $\ce{HNO3}$ trong quá trình (4)}
""" + p2_c2_sol + r"""
\end{ex}
""")

    # ==========================================
    # PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN
    # ==========================================
    tex.append(r"""
\vspace{6pt}
\noindent
\textbf{\large\color{mainGreen}PHẦN III. Câu trắc nghiệm trả lời ngắn: (4 câu - 1 điểm)}

\textit{Thí sinh trả lời từ câu 1 đến câu 4.}
\vspace{4pt}
""")

    # Câu 1 Phần III
    p3_c1_ans_line = r"\textbf{Trả lời: } \textbf{\color{hfRed}1,7}" if is_teacher else r"\textbf{Trả lời: } \dotfill"
    p3_c1_sol = r"""\begin{loigiai}
\begin{itemize}[leftmargin=0.8cm]
\item Khối lượng $\ce{HNO3}$ nguyên chất cần điều chế:
$$m_{\ce{HNO3}} = 8{,}4 \times 60\% = 5{,}04\,\text{tấn}.$$
\item Số mol $\ce{HNO3}$: $n_{\ce{HNO3}} = \dfrac{5{,}04}{63} = 0{,}08\,\text{kmol}$ (triệu mol).
\item Sơ đồ bảo toàn nguyên tố $\ce{N}$: $\ce{NH3 -> HNO3}$.
$$n_{\ce{NH3}\,\text{lí thuyết}} = n_{\ce{HNO3}} = 0{,}08\,\text{kmol}.$$
\item Với hiệu suất toàn bộ quá trình $H = 80\% = 0{,}8$:
$$n_{\ce{NH3}\,\text{thực tế}} = \dfrac{0{,}08}{0{,}8} = 0{,}1\,\text{kmol}.$$
\item Khối lượng ammonia cần dùng:
$$m_{\ce{NH3}} = 0{,}1 \times 17 = \mathbf{1{,}7}\,\text{tấn}.$$
\end{itemize}
$\rightarrow$ \textbf{Đáp số: 1,7}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-P3-C1]
Trong công nghiệp, người ta sản xuất nitric acid ($\ce{HNO3}$) từ ammonia theo sơ đồ chuyển hoá sau:
$$\ce{NH3 ->[+O_2,\, t^\circ,\, \text{xt}] NO ->[+O_2] NO2 ->[+O_2 + H_2O] HNO3}$$
Để điều chế $8{,}4$ tấn dung dịch nitric acid có nồng độ $60\%$ cần dùng bao nhiêu tấn ammonia? Biết rằng hiệu suất của toàn bộ quá trình sản xuất nitric acid theo sơ đồ trên là $80\%$.

""" + p3_c1_ans_line + r"""
""" + p3_c1_sol + r"""
\end{ex}
""")

    # Câu 2 Phần III
    p3_c2_ans_line = r"\textbf{Trả lời: } \textbf{\color{hfRed}12}" if is_teacher else r"\textbf{Trả lời: } \dotfill"
    p3_c2_sol = r"""\begin{loigiai}
\begin{itemize}[leftmargin=0.8cm]
\item Số mol $\ce{NaOH}$: $n_{\ce{NaOH}} = \dfrac{0{,}8}{40} = 0{,}02\,\text{mol}$.
\item Nồng độ mol dung dịch $\ce{NaOH}$ ($V = 2000\,\text{mL} = 2\,\text{L}$):
$$C_{\ce{NaOH}} = \dfrac{0{,}02}{2} = 0{,}01\,\text{M} = 10^{-2}\,\text{M}.$$
\item $\ce{NaOH}$ phân li hoàn toàn: $[\ce{OH-}] = 0{,}01\,\text{M} = 10^{-2}\,\text{M} \Rightarrow \text{pOH} = -\lg(10^{-2}) = 2$.
\item Giá trị pH ở $25^\circ\text{C}$: $\text{pH} = 14 - \text{pOH} = 14 - 2 = \mathbf{12}$.
\end{itemize}
$\rightarrow$ \textbf{Đáp số: 12}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-P3-C2]
Hòa tan hết $0{,}8$ gam $\ce{NaOH}$ vào nước để được $2000\,\text{mL}$ dung dịch ở $25^\circ\text{C}$. pH của dung dịch thu được bằng mấy?

""" + p3_c2_ans_line + r"""
""" + p3_c2_sol + r"""
\end{ex}
""")

    # Câu 3 Phần III
    p3_c3_ans_line = r"\textbf{Trả lời: } \textbf{\color{hfRed}4}" if is_teacher else r"\textbf{Trả lời: } \dotfill"
    p3_c3_sol = r"""\begin{loigiai}
Theo thuyết Brønsted -- Lowry, base là chất hoặc ion có khả năng nhận proton ($\ce{H+}$).
\begin{enumerate}
\item $\ce{NaOH}$: phân li ra ion $\ce{OH-}$, $\ce{OH-}$ nhận proton ($\ce{OH- + H+ -> H2O}$) $\Rightarrow$ là base. (1)
\item $\ce{H2SO4}$: là acid cho proton.
\item $\ce{NH3}$: nhận proton ($\ce{NH3 + H+ <=> NH4+}$) $\Rightarrow$ là base. (2)
\item $\ce{CO3^{2-}}$: nhận proton ($\ce{CO3^{2-} + H+ <=> HCO3-}$) $\Rightarrow$ là base. (3)
\item $\ce{S^{2-}}$: nhận proton ($\ce{S^{2-} + H+ <=> HS-}$ $\Rightarrow$ là base. (4)
\item $\ce{NH4+}$: là acid nhường proton ($\ce{NH4+ <=> NH3 + H+}$).
\end{enumerate}
Vậy có tất cả 4 chất và ion đóng vai trò là base ($\ce{NaOH, NH3, CO3^{2-}, S^{2-}}$).\\
$\rightarrow$ \textbf{Đáp số: 4}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-P3-C3]
Cho các chất (phân tử và ion): $\ce{NaOH, H2SO4, NH3, CO3^{2-}, S^{2-}, NH4+}$. Theo quan điểm của Bronsted – Lowry có bao nhiêu chất (phân tử và ion) có thể là base?

""" + p3_c3_ans_line + r"""
""" + p3_c3_sol + r"""
\end{ex}
""")

    # Câu 4 Phần III
    p3_c4_ans_line = r"\textbf{Trả lời: } \textbf{\color{hfRed}3}" if is_teacher else r"\textbf{Trả lời: } \dotfill"
    p3_c4_sol = r"""\begin{loigiai}
Phản ứng thuận tỏa nhiệt ($\Delta_\text{r}H^\circ_{298} < 0$) và làm giảm số mol khí (từ $3\,\text{mol}$ khí về $2\,\text{mol}$ khí).
\begin{itemize}[leftmargin=0.8cm]
\item (1) Giảm nhiệt độ: Cân bằng chuyển dịch theo chiều tỏa nhiệt $\Rightarrow$ CHIỀU THUẬN. $\checkmark$
\item (2) Tăng áp suất: Cân bằng chuyển dịch theo chiều làm giảm số mol khí $\Rightarrow$ CHIỀU THUẬN. $\checkmark$
\item (3) Dùng xúc tác $\ce{V2O5}$: Không làm chuyển dịch cân bằng.
\item (4) Tăng nồng độ $\ce{SO3}$: Cân bằng chuyển dịch theo chiều làm giảm $\ce{SO3}$ $\Rightarrow$ CHIỀU NGHỊCH.
\item (5) Tăng nồng độ $\ce{SO2}$: Cân bằng chuyển dịch theo chiều làm giảm $\ce{SO2}$ $\Rightarrow$ CHIỀU THUẬN. $\checkmark$
\end{itemize}
Có 3 biện pháp làm cân bằng chuyển dịch theo chiều thuận: (1), (2), (5).\\
$\rightarrow$ \textbf{Đáp số: 3}
\end{loigiai}""" if is_teacher else ""
    tex.append(r"""
\begin{ex}%[1002-P3-C4]
Cho cân bằng hoá học sau: $\ce{2SO2(g) + O2(g) <=> 2SO3(g)};\ \Delta_\text{r}H^\circ_{298} < 0$\\
Người ta thực hiện các biện pháp sau:
\begin{enumerate}[label=(\arabic*)]
\item Giảm nhiệt độ,
\item Tăng áp suất chung của hệ phản ứng,
\item Dùng thêm chất xúc tác $\ce{V2O5}$,
\item Tăng nồng độ $\ce{SO3}$,
\item Tăng nồng độ $\ce{SO2}$.
\end{enumerate}
Có bao nhiêu biện pháp làm cân bằng trên chuyển dịch theo chiều thuận?

""" + p3_c4_ans_line + r"""
""" + p3_c4_sol + r"""
\end{ex}
""")

    # ==========================================
    # PHẦN IV. TỰ LUẬN
    # ==========================================
    tex.append(r"""
\vspace{6pt}
\noindent
\textbf{\large\color{mainGreen}PHẦN IV. TỰ LUẬN: (3 câu - 3 điểm)}

\textit{Thí sinh trả lời từ câu 1 đến câu 3.}
\vspace{4pt}
""")

    # Tự luận Câu 1
    tl1_sol = r"""\begin{loigiai}
Các phương trình phản ứng hoàn chỉnh (mỗi phương trình 0,25 điểm):
\begin{align*}
\text{a)}\quad &\ce{N2(g) + O2(g) <=>[3000^\circ\text{C}\text{ hoặc tia lửa điện}] 2NO(g)}\\
\text{b)}\quad &\ce{NH3(g) + HCl(g) -> NH4Cl(s)}\quad \text{(khói trắng)}\\
\text{c)}\quad &\ce{4HNO3 (đặc) + FeO ->[t^\circ] Fe(NO3)3 + NO2 ^ + 2H2O}\\
\text{d)}\quad &\ce{NH4HCO3(s) ->[t^\circ] NH3 ^ + CO2 ^ + H2O}
\end{align*}
\end{loigiai}""" if is_teacher else r"""
\noindent\textit{Bài làm:}
\multido{\i=1+1}{7}{\noindent\makebox[\linewidth]{\dotfill}\\[1.5pt]}
"""
    tex.append(r"""
\begin{ex}%[1002-TL1]
\textbf{(1,0 điểm)} Hoàn thành các phương trình phản ứng sau (ghi rõ điều kiện phản ứng)
\begin{enumerate}[label=\textbf{\alph*)}]
\item $\ce{N2 + O2 ->}$
\item $\ce{NH3 + HCl ->}$
\item $\ce{HNO3 \text{đặc} + FeO ->}$
\item Nhiệt phân $\ce{NH4HCO3 ->}$
\end{enumerate}
""" + tl1_sol + r"""
\end{ex}
""")

    # Tự luận Câu 2
    tl2_sol = r"""\begin{loigiai}
\begin{itemize}[leftmargin=0.8cm]
\item Biểu thức hằng số cân bằng nồng độ $K_\text{c}$ của phản ứng:
$$K_\text{c} = \dfrac{[\ce{NH3}]^2}{[\ce{N2}]\cdot[\ce{H2}]^3}\quad \text{(0,5 điểm)}$$
\item Thay nồng độ của các chất ở trạng thái cân bằng vào biểu thức:
$$K_\text{c} = \dfrac{(0{,}30)^2}{0{,}05 \times (0{,}10)^3} = \dfrac{0{,}09}{0{,}05 \times 0{,}001} = \dfrac{0{,}09}{0{,}00005} = \mathbf{1800}.\quad \text{(0,5 điểm)}$$
\end{itemize}
$\rightarrow$ \textbf{Đáp số: } $K_\text{c} = 1800$.
\end{loigiai}""" if is_teacher else r"""
\noindent\textit{Bài làm:}
\multido{\i=1+1}{6}{\noindent\makebox[\linewidth]{\dotfill}\\[1.5pt]}
"""
    tex.append(r"""
\begin{ex}%[1002-TL2]
\textbf{(1,0 điểm)} Cho phản ứng tổng hợp ammonia ở trạng thái cân bằng tại $430^\circ\text{C}$.
$$\ce{N2(g) + 3H2(g) <=> 2NH3(g)} .$$
Ở trạng thái cân bằng, nồng độ của $\ce{NH3}, \ce{N2}$ và $\ce{H2}$ lần lượt là $0{,}30\,\text{M}$; $0{,}05\,\text{M}$ và $0{,}10\,\text{M}$.\\
Tính hằng số cân bằng của phản ứng ở nhiệt độ trên.
""" + tl2_sol + r"""
\end{ex}
""")

    # Tự luận Câu 3
    tl3_sol = r"""\begin{loigiai}
\begin{itemize}[leftmargin=0.8cm]
\item Lượng $\ce{HCl}$ ban đầu có trong $1{,}5\,\text{L}$ dịch vị dạ dày:
$$n_{\ce{HCl}\,\text{ban đầu}} = \dfrac{0{,}09125}{36{,}5} = 0{,}0025\,\text{mol}.\quad \text{(0,25 điểm)}$$
\item Để duy trì độ pH trong dạ dày ở mức 3:
$$[\ce{H+}] = 10^{-3}\,\text{M} = 0{,}001\,\text{M}.$$
Lượng ion $\ce{H+}$ (tương ứng $\ce{HCl}$) còn lại trong $1{,}5\,\text{L}$ dịch vị sau khi trung hòa:
$$n_{\ce{HCl}\,\text{còn lại}} = [\ce{H+}] \times V = 0{,}001 \times 1{,}5 = 0{,}0015\,\text{mol}.\quad \text{(0,25 điểm)}$$
\item Lượng $\ce{HCl}$ đã được trung hòa bởi $\ce{NaHCO3}$:
$$n_{\ce{HCl}\,\text{phản ứng}} = n_{\ce{HCl}\,\text{ban đầu}} - n_{\ce{HCl}\,\text{còn lại}} = 0{,}0025 - 0{,}0015 = 0{,}001\,\text{mol}.$$
\item Phương trình hóa học của phản ứng trung hòa:
$$\ce{NaHCO3 + HCl -> NaCl + CO2 ^ + H2O}$$
Theo phương trình: $n_{\ce{NaHCO3}\,\text{phản ứng}} = n_{\ce{HCl}\,\text{phản ứng}} = 0{,}001\,\text{mol}$.\\
Khối lượng $\ce{NaHCO3}$ thực tế phản ứng (tiêu thụ):
$$m_{\ce{NaHCO3}\,\text{phản ứng}} = 0{,}001 \times 84 = 0{,}084\,\text{gam}.\quad \text{(0,25 điểm)}$$
\item Do khả năng tiêu thụ $\ce{NaHCO3}$ của cơ thể người bệnh là $70\%$ (hiệu suất $H = 70\% = 0{,}7$), khối lượng $\ce{NaHCO3}$ người bệnh cần đưa vào cơ thể là:
$$m_{\ce{NaHCO3}\,\text{cần đưa vào}} = \dfrac{m_{\ce{NaHCO3}\,\text{phản ứng}}}{70\%} = \dfrac{0{,}084}{0{,}7} = \mathbf{0{,}12\,\text{gam}}.\quad \text{(0,25 điểm)}$$
\end{itemize}
$\rightarrow$ \textbf{Đáp số: } Cần đưa vào cơ thể $0{,}12\,\text{gam}$ $\ce{NaHCO3}$.
\end{loigiai}""" if is_teacher else r"""
\noindent\textit{Bài làm:}
\multido{\i=1+1}{10}{\noindent\makebox[\linewidth]{\dotfill}\\[1.5pt]}
"""
    tex.append(r"""
\begin{ex}%[1002-TL3]
\textbf{(1,0 điểm)} Hydrochloric acid ($\ce{HCl}$) là một chất được phát hiện trong dịch vị dạ dày có nồng độ $10^{-4} - 10^{-3}\,\text{mol/L}$ và độ pH duy trì ở mức $3 - 4$ đối với người bình thường. Nếu thiếu acid trong dạ dày thì thức ăn không chuyển hóa được lâu dần gây suy nhược cơ thể, nếu dư acid lâu ngày thì nó sẽ phá hủy đường ruột gây viêm loét dạ dày. Khi cơ thể dư acid $\ce{HCl}$, người bệnh có thể sử dụng thuốc giảm đau dạ dày có tên gọi là Nabica (thành phần chính là $\ce{NaHCO3}$), được dùng để trung hòa bớt lượng acid dư trong dạ dày.\\
Giả sử dịch vị dạ dày người bệnh chứa $1{,}5$ lít dung dịch hỗn hợp thức ăn lỏng, trong đó chứa $0{,}09125$ gam $\ce{HCl}$. Nếu khả năng tiêu thụ $\ce{NaHCO3}$ của cơ thể người bệnh là $70\%$ thì khối lượng $\ce{NaHCO3}$ người đó cần đưa vào cơ thể để duy trì độ pH trong dạ dày ở mức 3 là bao nhiêu?
""" + tl3_sol + r"""
\end{ex}
""")

    # Kết thúc
    tex.append(r"""
\vspace{8pt}
\begin{center}
\textbf{---------------------- HẾT ----------------------}
\end{center}
\begin{center}
\textit{\small - Thí sinh không được sử dụng tài liệu.\\ - Cán bộ coi thi không giải thích gì thêm.}
\end{center}

\end{document}
""")
    return "".join(tex)


def compile_latex(tex_path):
    work_dir = os.path.dirname(tex_path)
    base_name = os.path.splitext(os.path.basename(tex_path))[0]
    cmd = ["pdflatex", "-interaction=nonstopmode", f"{base_name}.tex"]
    print(f"Biên dịch LaTeX: {base_name}.tex...")
    
    # Run 1
    res1 = subprocess.run(cmd, cwd=work_dir, capture_output=True, text=True, encoding='utf-8', errors='ignore')
    # Run 2 for cross references
    res2 = subprocess.run(cmd, cwd=work_dir, capture_output=True, text=True, encoding='utf-8', errors='ignore')
    
    pdf_path = os.path.join(work_dir, f"{base_name}.pdf")
    if os.path.exists(pdf_path):
        print(f"-> Tạo thành công PDF: {pdf_path}")
    else:
        print(f"Lỗi biên dịch {base_name}.tex:")
        print(res2.stdout[-1500:])
        return False

    # Cleanup rác
    for ext in [".aux", ".log", ".out", ".toc", ".thm", ".ans"]:
        junk = os.path.join(work_dir, f"{base_name}{ext}")
        if os.path.exists(junk):
            try:
                os.remove(junk)
            except Exception:
                pass
    ans_files = glob.glob(os.path.join(work_dir, "ans-*.tex")) + glob.glob(os.path.join(work_dir, "ans*.tex"))
    for af in ans_files:
        try:
            os.remove(af)
        except Exception:
            pass
    return True

if __name__ == "__main__":
    os.makedirs(SAN_PHAM_DIR, exist_ok=True)

    # 1. Bản Học sinh
    print("\n--- TẠO BẢN HỌC SINH ---")
    hs_base = "De_Kiem_Tra_GK1_Hoa_11_NguyenChiThanh_Ma1002_HocSinh"
    hs_tex_path = os.path.join(SAN_PHAM_DIR, hs_base + ".tex")
    with open(hs_tex_path, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))
    print(f"Đã ghi file: {hs_tex_path}")
    compile_latex(hs_tex_path)

    # 2. Bản Giáo viên
    print("\n--- TẠO BẢN GIÁO VIÊN ---")
    gv_base = "De_Kiem_Tra_GK1_Hoa_11_NguyenChiThanh_Ma1002_GiaoVien"
    gv_tex_path = os.path.join(SAN_PHAM_DIR, gv_base + ".tex")
    with open(gv_tex_path, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))
    print(f"Đã ghi file: {gv_tex_path}")
    compile_latex(gv_tex_path)

    print("\nHOÀN THÀNH BIÊN DỊCH CẢ 2 BẢN LATEX & PDF!")
