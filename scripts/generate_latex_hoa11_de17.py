# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file LaTeX và biên dịch ra PDF cho ĐỀ SỐ 17 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De17_HocSinh.tex / .pdf
- De_Kiem_Tra_Hoa_11_De17_GiaoVien.tex / .pdf
Chương 6: HỢP CHẤT CARBONYL - CARBOXYLIC ACID
Tuân thủ Form chuẩn GEMINI.md, sử dụng gói ex_test.sty và mhchem.
"""

import os
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import subprocess
import shutil

TEX_DIR = r"c:\Antigravity_Thanh\San_Pham"

if not os.path.exists(os.path.join(TEX_DIR, "ex_test.sty")):
    shutil.copyfile(r"c:\Antigravity_Thanh\ex_test.sty", os.path.join(TEX_DIR, "ex_test.sty"))

def make_tex(is_teacher=False):
    tex = r"""\documentclass[12pt,a4paper]{article}
\usepackage[utf8]{vietnam}
\usepackage{amsmath,amssymb}
\usepackage{geometry}
\geometry{top=1.5cm,bottom=1.2cm,left=1.4cm,right=1.4cm,headheight=28pt,headsep=12pt,footskip=18pt}
\usepackage{tikz}
\usepackage{xcolor}
\usepackage{enumerate}
\usepackage{graphicx}
\usepackage[version=4]{mhchem}
\usepackage{fancyhdr}
\usepackage{ex_test}

\renewcommand{\baselinestretch}{0.96}
\setlength{\parskip}{0pt}

\definecolor{hfRed}{RGB}{211,47,47}
\definecolor{hfGreen}{RGB}{139,195,144}

\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small\itshape\color{hfRed} Tài liệu lưu hành nội bộ}
\fancyhead[R]{\small\itshape\color{hfRed} Thầy \textbf{TRẦN VĂN THẠNH} - 0777.470.803}
\renewcommand{\headrulewidth}{0.8pt}
\renewcommand{\headrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \headrulewidth\hfill}}

\fancyfoot[L]{\small\itshape\color{hfRed} CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế \quad|\quad CS2: 24 Đặng Thái Thân, TP Huế}
\fancyfoot[R]{\small\bfseries\color{hfRed} Trang \thepage}
\renewcommand{\footrulewidth}{0.8pt}
\renewcommand{\footrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \footrulewidth\hfill}}

\fancypagestyle{firstpage}{
  \fancyhf{}
  \renewcommand{\headrulewidth}{0pt}
  \fancyfoot[L]{\small\itshape\color{hfRed} CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế \quad|\quad CS2: 24 Đặng Thái Thân, TP Huế}
  \fancyfoot[R]{\small\bfseries\color{hfRed} Trang \thepage}
  \renewcommand{\footrulewidth}{0.8pt}
  \renewcommand{\footrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \footrulewidth\hfill}}
}

\begin{document}
\thispagestyle{firstpage}

\noindent
\begin{minipage}[t]{0.44\textwidth}
	\centering\fontsize{10pt}{12pt}\selectfont
	\textbf{SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ}\\
	\textbf{ĐỀ CHÍNH THỨC}\\
	\textit{(Đề thi có 04 trang)}
\end{minipage}\hfill
\begin{minipage}[t]{0.54\textwidth}
	\centering\fontsize{10.5pt}{13pt}\selectfont
	\textbf{ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 6}\\
	\textbf{(HỢP CHẤT CARBONYL - CARBOXYLIC ACID)}\\
	\textbf{MÔN: HOÁ HỌC - LỚP 11}\\
	\textit{Thời gian làm bài: 50 phút, không kể thời gian phát đề}
	
	\smallskip
	\fbox{\textbf{MÃ ĐỀ: 117}}
\end{minipage}
"""
    if is_teacher:
        tex += r"""
\smallskip
\begin{center}
	{\color{hfRed}\textbf{\large (BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}}
\end{center}
"""

    tex += r"""
\smallskip
\noindent\hrulefill
\smallskip

\noindent\textbf{Họ, tên thí sinh:} \dotfill\ \textbf{Lớp:} \dotfill\ \textbf{Số báo danh:} \dotfill

\smallskip
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Cr = 52; Mn = 55; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108; Ba = 137.)}

\smallskip
\noindent\textbf{\large PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.}
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{0.2cm}
"""

    # C1
    tex += r"""
\begin{ex}
Chất nào sau đây là acetic aldehyde?
\choice
{\ce{HCHO}}
{""" + (r"\True " if is_teacher else "") + r"""\ce{CH3CHO}}
{\ce{C2H5OH}}
{\ce{C2H6}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Acetic aldehyde (ethanal, acetaldehyde) có công thức là \ce{CH3CHO}.\\
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C2
    tex += r"""
\begin{ex}
Tên theo danh pháp thay thế của chất \ce{CH3-CH2-CH2-CHO} là gì?
\choice
{Propan-1-al}
{propanal}
{""" + (r"\True " if is_teacher else "") + r"""butanal}
{butyraldehyde}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Mạch chính có 4 carbon, no, chứa nhóm \ce{-CHO} ở vị trí số 1 $\Rightarrow$ tên thay thế là butanal.\\
$\Rightarrow$ Chọn phương án \textbf{C}.
}
"""
    tex += r"""\end{ex}
"""

    # C3
    tex += r"""
\begin{ex}
Benzaldehyde tác dụng với \ce{Cu(OH)2} trong môi trường base ($t^\circ\text{C}$) thu được kết tủa màu đỏ gạch là
\choice
{Cu}
{""" + (r"\True " if is_teacher else "") + r"""\ce{Cu2O}}
{CuO}
{\ce{Cu2O}, Cu}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Aldehyde tác dụng với \ce{Cu(OH)2} trong môi trường kiềm đun nóng tạo kết tủa đỏ gạch \ce{Cu2O}.\\
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C4
    tex += r"""
\begin{ex}
Phản ứng giữa \ce{CH3CHO} với HCN thuộc loại phản ứng
\choice
{phản ứng thuận nghịch}
{""" + (r"\True " if is_teacher else "") + r"""phản ứng cộng}
{phản ứng thế}
{phản ứng phân hủy}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Phản ứng cộng nucleophile của HCN vào liên kết đôi \ce{C=O} tạo cyanohydrin (\ce{CH3-CH(OH)-CN}).\\
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C5
    tex += r"""
\begin{ex}
\textbf{[KNTT - SBT]} Có bốn chất lỏng có thể tích bằng nhau là ethanol, acetone, acetaldehyde, acetic acid. Tiến hành chưng cất hỗn hợp này, sau một thời gian, hàm lượng chất nào trong bình chưng cất còn lại lớn nhất?
\choice
{Ethanol}
{Acetone}
{Acetaldehyde}
{""" + (r"\True " if is_teacher else "") + r"""Acetic acid}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Nhiệt độ sôi: acetaldehyde ($21^\circ\text{C}$) $<$ acetone ($56^\circ\text{C}$) $<$ ethanol ($78,3^\circ\text{C}$) $<$ acetic acid ($118^\circ\text{C}$). Chất có nhiệt độ sôi cao nhất bay hơi chậm nhất nên còn lại nhiều nhất trong bình chưng cất.\\
$\Rightarrow$ Chọn phương án \textbf{D}.
}
"""
    tex += r"""\end{ex}
"""

    # C6
    tex += r"""
\begin{ex}
\textbf{[KNTT - SBT]} Cho ba chất lỏng riêng biệt sau: \ce{C2H5OH}, \ce{CH3CHO}, \ce{CH3COOH}. Cách nào sau đây phù hợp để phân biệt ba chất lỏng trên?
\choice
{Dùng quỳ tím, sau đó dùng dung dịch \ce{NaOH}}
{""" + (r"\True " if is_teacher else "") + r"""Dùng quỳ tím, sau đó dùng dung dịch \ce{AgNO3} trong \ce{NH3}}
{Dùng Na sau đó dùng dung dịch \ce{AgNO3} trong \ce{NH3}}
{Dùng Na sau đó dùng quỳ tím}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- Dùng quỳ tím: \ce{CH3COOH} làm quỳ hóa đỏ, hai chất còn lại không đổi màu.\\
- Dùng \ce{AgNO3/NH3} đun nóng: \ce{CH3CHO} tạo kết tủa bạc sáng bóng, \ce{C2H5OH} không phản ứng.\\
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C7
    tex += r"""
\begin{ex}
\textbf{(SBT - CD)} Dãy nào sau đây gồm các chất đều tác dụng với dung dịch \ce{AgNO3/NH3}?
\choice
{Acetaldehyde, but-1-yne, ethylene}
{Acetaldehyde, acetylene, but-2-yne}
{""" + (r"\True " if is_teacher else "") + r"""Formaldehyde, vinylacetylene, propyne}
{Formaldehyde, acetylene, ethylene}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- Formaldehyde (\ce{HCHO}) cho phản ứng tráng bạc.\\
- Vinylacetylene và propyne có liên kết ba đầu mạch tác dụng tạo kết tủa vàng nhạt.\\
$\Rightarrow$ Chọn phương án \textbf{C}.
}
"""
    tex += r"""\end{ex}
"""

    # C8
    tex += r"""
\begin{ex}
\textbf{(SBT - CD)} Trong các chất sau: (1) \ce{CH3CH2CHO}, (2) \ce{CH3CH(OH)CH3}, (3) \ce{(CH3)2CHCHO}, (4) \ce{CH2=CHCH2OH}. Những chất nào phản ứng với \ce{H2} (Ni, $t^\circ$) hoặc \ce{NaBH4} sinh ra cùng một sản phẩm?
\choice
{""" + (r"\True " if is_teacher else "") + r"""(1) và (4)}
{(2) và (4)}
{(1) và (2)}
{(3) và (4)}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Cả (1) \ce{CH3CH2CHO} và (4) \ce{CH2=CH-CH2OH} khi phản ứng với lượng dư \ce{H2} (Ni, $t^\circ$) đều tạo propan-1-ol (\ce{CH3CH2CH2OH}).\\
$\Rightarrow$ Chọn phương án \textbf{A}.
}
"""
    tex += r"""\end{ex}
"""

    # C9
    tex += r"""
\begin{ex}
\textbf{(SBT - CTST)} Cho dãy các chất sau: \ce{HCHO}, \ce{CH3COOH}, \ce{CH3CHO}, \ce{HCOOH}, \ce{C2H5OH}. Số chất trong dãy có thể tham gia phản ứng tráng bạc là
\choice
{""" + (r"\True " if is_teacher else "") + r"""3}
{6}
{4}
{5}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Gồm 3 chất có nhóm \ce{-CHO}: \ce{HCHO}, \ce{CH3CHO}, \ce{HCOOH}.\\
$\Rightarrow$ Chọn phương án \textbf{A}.
}
"""
    tex += r"""\end{ex}
"""

    # C10
    tex += r"""
\begin{ex}
Hai hợp chất hữu cơ X, Y có cùng công thức phân tử \ce{C3H6O2}. Cả X và Y đều tác dụng với Na; X tác dụng được với \ce{NaHCO3} còn Y có khả năng tham gia phản ứng tráng bạc. Công thức cấu tạo của X và Y lần lượt là
\choice
{\ce{C2H5COOH} và \ce{HCOOC2H5}}
{\ce{HCOOC2H5} và \ce{HOCH2COCH3}}
{\ce{HCOOC2H5} và \ce{HOCH2CH2CHO}}
{""" + (r"\True " if is_teacher else "") + r"""\ce{C2H5COOH} và \ce{CH3CH(OH)CHO}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- X tác dụng \ce{NaHCO3} $\Rightarrow$ X là carboxylic acid: \ce{C2H5COOH}.\\
- Y tác dụng Na (có \ce{-OH}) và tráng bạc (có \ce{-CHO}) $\Rightarrow$ Y là \ce{CH3CH(OH)CHO}.\\
$\Rightarrow$ Chọn phương án \textbf{D}.
}
"""
    tex += r"""\end{ex}
"""

    # C11
    tex += r"""
\begin{ex}
Dãy gồm các chất có thể điều chế trực tiếp (bằng một phản ứng) tạo ra acetic acid là:
\choice
{\ce{CH3CHO}, \ce{C2H5OH}, \ce{C2H5COOCH3}}
{\ce{CH3CHO}, \ce{C6H12O6} (glucose), \ce{CH3OH}}
{""" + (r"\True " if is_teacher else "") + r"""\ce{CH3OH}, \ce{C2H5OH}, \ce{CH3CHO}}
{\ce{C2H4(OH)2}, \ce{CH3OH}, \ce{CH3CHO}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- \ce{CH3OH + CO ->[{xt, t$^\circ$}] CH3COOH}\\
- \ce{C2H5OH + O2 ->[{men giấm}] CH3COOH + H2O}\\
- \ce{2CH3CHO + O2 ->[{xt, t$^\circ$}] 2CH3COOH}\\
$\Rightarrow$ Chọn phương án \textbf{C}.
}
"""
    tex += r"""\end{ex}
"""

    # C12
    tex += r"""
\begin{ex}
Để điều chế \ce{CH3CHO} từ \ce{Al4C3} cần ít nhất bao nhiêu phương trình phản ứng?
\choice
{2}
{""" + (r"\True " if is_teacher else "") + r"""3}
{4}
{5}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Cần ít nhất 3 phản ứng: \ce{Al4C3 -> CH4 -> C2H2 -> CH3CHO}.\\
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C13
    tex += r"""
\begin{ex}
Cho dãy các chất: \ce{HCHO}, \ce{CH3COOH}, \ce{HCOONa}, \ce{HCOOH}, \ce{C2H5OH}, \ce{HCOOCH3}. Số chất trong dãy tham gia phản ứng tráng gương là
\choice
{3}
{6}
{""" + (r"\True " if is_teacher else "") + r"""4}
{5}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Các chất có nhóm \ce{H-C=O} tráng gương được gồm 4 chất: \ce{HCHO}, \ce{HCOONa}, \ce{HCOOH}, \ce{HCOOCH3}.\\
$\Rightarrow$ Chọn phương án \textbf{C}.
}
"""
    tex += r"""\end{ex}
"""

    # C14
    tex += r"""
\begin{ex}
Aldehyde bị khử bằng \ce{NaBH4} hoặc \ce{LiAlH4} sinh ra
\choice
{alcohol bậc hai}
{""" + (r"\True " if is_teacher else "") + r"""alcohol bậc một}
{ketone}
{alcohol bậc ba}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Aldehyde \ce{R-CHO} bị khử bởi \ce{NaBH4} hoặc \ce{LiAlH4} tạo alcohol bậc một \ce{R-CH2OH}.\\
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C15
    tex += r"""
\begin{ex}
Cho sơ đồ phản ứng: \ce{C2H5OH -> X -> CH3COOH} (mỗi mũi tên ứng với một phản ứng). Chất X là
\choice
{\ce{HCHO}}
{\ce{C2H5CHO}}
{\ce{CH4}}
{""" + (r"\True " if is_teacher else "") + r"""\ce{CH3CHO}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
\ce{C2H5OH} oxi hóa thành \ce{CH3CHO}, sau đó \ce{CH3CHO} oxi hóa tiếp thành \ce{CH3COOH}.\\
$\Rightarrow$ Chọn phương án \textbf{D}.
}
"""
    tex += r"""\end{ex}
"""

    # C16
    tex += r"""
\begin{ex}
\textbf{(SBT - CTST)} Trung hoà 100 mL dung dịch carboxylic acid (X) nồng độ 0,1 M cần vừa đủ 16 gam dung dịch \ce{NaOH} 5\%, thu được 1,48 gam muối. Công thức cấu tạo của (X) là
\choice
{""" + (r"\True " if is_teacher else "") + r"""\ce{HOOC-CH2-COOH}}
{\ce{HOOC-COOH}}
{\ce{CH3-COOH}}
{\ce{HOOC-CH2-CH2-COOH}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
$n_{\ce{X}} = 0,1 \times 0,1 = 0,01\text{ mol}$.\\
$n_{\ce{NaOH}} = \dfrac{16 \times 5\%}{40} = 0,02\text{ mol} = 2 n_{\ce{X}} \Rightarrow$ X có 2 nhóm \ce{-COOH}.\\
$M_{\text{muối}} = \dfrac{1,48}{0,01} = 148\text{ g/mol} \Rightarrow R + 134 = 148 \Rightarrow R = 14$ (\ce{-CH2-}).\\
$\Rightarrow$ X là \ce{HOOC-CH2-COOH}. Chọn phương án \textbf{A}.
}
"""
    tex += r"""\end{ex}
"""

    # C17
    tex += r"""
\begin{ex}
\textbf{(SBT - CTST)} Cho dung dịch acetic acid có nồng độ $x\%$ tác dụng vừa đủ với dung dịch \ce{NaOH} 20\% thu được dung dịch muối sodium acetate có nồng độ 10,25\%. Giá trị của $x$ là
\choice
{""" + (r"\True " if is_teacher else "") + r"""10\%}
{15\%}
{18,67\%}
{20\%}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Giả sử có 1 mol \ce{CH3COOH} phản ứng:\\
$m_{\text{dd NaOH}} = \dfrac{40 \times 100}{20} = 200\text{ g}$.\\
$m_{\text{dd muối}} = \dfrac{82 \times 100}{10,25} = 800\text{ g}$.\\
$m_{\text{dd acid}} = 800 - 200 = 600\text{ g} \Rightarrow x\% = \dfrac{60}{600} \times 100\% = 10\%$.\\
$\Rightarrow$ Chọn phương án \textbf{A}.
}
"""
    tex += r"""\end{ex}
"""

    # C18
    tex += r"""
\begin{ex}
\textbf{(SBT - CD)} Acetic acid được tổng hợp từ nguồn khí than đá theo các phản ứng sau:
\begin{align*}
\text{(1) } & \ce{CO + 2H2 ->[{t$^\circ$, xt}] CH3OH} \\
\text{(2) } & \ce{CH3OH + CO ->[{t$^\circ$, xt}] CH3COOH}
\end{align*}
Giả sử hiệu suất của các phản ứng (1) và (2) đều đạt 90\%. Để sản xuất 1000 lít acetic acid ($D = 1,05\text{ g mL}^{-1}$), cần thể tích khí CO và \ce{H2} (ở điều kiện chuẩn) lần lượt là
\choice
{$964,06\text{ m}^3$ và $1\,928,12\text{ m}^3$}
{$535,6\text{ m}^3$ và $1\,071,17\text{ m}^3$}
{$964,06\text{ m}^3$ và $964,06\text{ m}^3$}
{""" + (r"\True " if is_teacher else "") + r"""$1\,017,6\text{ m}^3$ và $1\,071,2\text{ m}^3$}
"""
    if is_teacher:
        tex += r"""
\loigiai{
$m = 1\,000 \times 1,05 = 1\,050\text{ kg} \Rightarrow n_{\text{acid}} = \dfrac{1\,050}{60} = 17,5\text{ kmol}$.\\
- Từ (2): $n_{\ce{CH3OH}} = \dfrac{17,5}{0,9} = \dfrac{175}{9}\text{ kmol}$; $n_{\ce{CO}(2)} = \dfrac{175}{9}\text{ kmol}$.\\
- Từ (1): $n_{\ce{CO}(1)} = \dfrac{175/9}{0,9} = \dfrac{1750}{81}\text{ kmol}$; $n_{\ce{H2}} = 2 \times n_{\ce{CO}(1)} = \dfrac{3500}{81}\text{ kmol}$.\\
$V_{\ce{CO}} = \left(\dfrac{1750}{81} + \dfrac{175}{9}\right) \times 24,79 \approx 1\,017,6\text{ m}^3$.\\
$V_{\ce{H2}} = \dfrac{3500}{81} \times 24,79 \approx 1\,071,2\text{ m}^3$.\\
$\Rightarrow$ Chọn phương án \textbf{D}.
}
"""
    tex += r"""\end{ex}
"""

    # PHẦN II
    tex += r"""
\bigskip
\noindent\textbf{\large PHẦN II. Câu trắc nghiệm đúng sai.}
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{0.2cm}
"""

    # C1 P2
    tex += r"""
\begin{ex}
Cinnamaldehyde (\ce{C6H5CH=CHCHO}) là hợp chất tạo mùi đặc trưng của tinh dầu quế. Cho các nhận định sau đây về Cinnamaldehyde:
\choiceTF
{""" + (r"\True " if is_teacher else "") + r"""Cinnamaldehyde có đồng phân hình học}
{""" + (r"\True " if is_teacher else "") + r"""Cinnamaldehyde có khả năng tham gia phản ứng tráng bạc}
{Ở điều kiện thường 1 mol Cinnamaldehyde có thể làm mất màu tối đa 1 mol bromine}
{Cinnamaldehyde phản ứng với Na kim loại giải phóng khí hydrogen}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- a) Đúng: Có liên kết đôi \ce{-CH=CH-} với 2 nhóm thế khác nhau nên có đồng phân cis-trans.\\
- b) Đúng: Có nhóm aldehyde \ce{-CHO} tráng bạc tạo Ag.\\
- c) Sai: 1 mol cinnamaldehyde phản ứng với 2 mol \ce{Br2} (1 mol cộng vào \ce{C=C} và 1 mol oxi hóa \ce{-CHO}).\\
- d) Sai: Không có nguyên tử H linh động nên không tác dụng với Na.\\
$\Rightarrow$ Đáp án: a-Đ, b-Đ, c-S, d-S.
}
"""
    tex += r"""\end{ex}
"""

    # C2 P2
    tex += r"""
\begin{ex}
Thực hiện thí nghiệm theo các bước như sau:\\
\textit{Bước 1:} Cho 2 mL dung dịch \ce{NaOH} 10\% và 2 mL dung dịch \ce{I2/KI} vào ống nghiệm có chứa 0,5 mL \ce{CH3CHO} 10\%.\\
\textit{Bước 2:} Lắc nhẹ ống nghiệm, hơ nhẹ ống nghiệm trên ngọn lửa đèn cồn hoặc đặt vào cốc nước nóng, sau đó làm nguội.
\choiceTF
{""" + (r"\True " if is_teacher else "") + r"""Đây là thí nghiệm nghiên cứu phản ứng tạo iodoform}
{""" + (r"\True " if is_teacher else "") + r"""Có thể thay \ce{CH3CHO} bằng \ce{CH3COCH3} vẫn thu được hiện tượng tương tự}
{Sau bước 2, thu được kết tủa màu trắng}
{Trong phản ứng trên, \ce{I2} đóng vai trò là chất khử}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- a) Đúng: Đây là phản ứng tạo iodoform (\ce{CHI3}).\\
- b) Đúng: Cả \ce{CH3CHO} và \ce{CH3COCH3} đều có nhóm \ce{CH3-C=O} nên đều tạo kết tủa vàng \ce{CHI3}.\\
- c) Sai: Iodoform (\ce{CHI3}) là kết tủa màu vàng nhạt.\\
- d) Sai: \ce{I2} đóng vai trò là chất oxi hóa.\\
$\Rightarrow$ Đáp án: a-Đ, b-Đ, c-S, d-S.
}
"""
    tex += r"""\end{ex}
"""

    # C3 P2
    tex += r"""
\begin{ex}
Malic acid là hợp chất hữu cơ có công thức cấu tạo \ce{HOOC-CH(OH)-CH2-COOH}, là nguyên nhân chính gây nên vị chua của quả táo.
\choiceTF
{""" + (r"\True " if is_teacher else "") + r"""Cứ 1 mol malic acid phản ứng được với tối đa 2 mol \ce{NaHCO3}}
{Malic là hợp chất hữu cơ đa chức}
{Công thức phân tử của Malic là \ce{C5H8O4}}
{""" + (r"\True " if is_teacher else "") + r"""Làm quỳ tím hóa đỏ}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- a) Đúng: Có 2 nhóm \ce{-COOH} nên phản ứng với 2 mol \ce{NaHCO3}.\\
- b) Sai: Malic acid vừa có nhóm \ce{-OH} vừa có nhóm \ce{-COOH} nên là hợp chất tạp chức.\\
- c) Sai: Công thức phân tử là \ce{C4H6O5}.\\
- d) Đúng: Có 2 nhóm \ce{-COOH} làm quỳ tím chuyển màu đỏ.\\
$\Rightarrow$ Đáp án: a-Đ, b-S, c-S, d-Đ.
}
"""
    tex += r"""\end{ex}
"""

    # C4 P2
    tex += r"""
\begin{ex}
\ce{CH3(CH2)3COOH} (chất A) là chất lỏng không màu, có mùi khó chịu. Đun nóng A với methanol (\ce{CH3OH}) có mặt của chất xúc tác acid sau một vài giờ thu được hỗn hợp chứa A, methanol và một sản phẩm hữu cơ \ce{CH3(CH2)3COOCH3} (chất V) có mùi trái cây dễ chịu. Một hợp chất hữu cơ được xác định bằng máy quang phổ IR như sau:
\begin{center}
\includegraphics[width=0.6\linewidth]{images_crop_hoa11_de17/ir_spectrum_pentanoic.png}
\end{center}
\choiceTF
{""" + (r"\True " if is_teacher else "") + r"""Dựa vào quang phổ IR xác định có nhóm chức \ce{-COOH} của (A)}
{Phổ MS sẽ xác định được khối lượng của chất V thu được từ thí nghiệm trên}
{""" + (r"\True " if is_teacher else "") + r"""Phản ứng của chất A với methanol tạo thành chất V là phản ứng thuận nghịch}
{Dựa vào quang phổ IR có thể xác định được công thức phân tử của một hợp chất hữu cơ}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- a) Đúng: Vùng hấp thụ rộng mạnh của \ce{O-H} acid ($2500 - 3300\text{ cm}^{-1}$) và peak \ce{C=O} ($1710\text{ cm}^{-1}$) xác nhận nhóm \ce{-COOH}.\\
- b) Sai: Phổ MS xác định phân tử khối, không đo khối lượng gam chất thu được trong thí nghiệm.\\
- c) Đúng: Phản ứng ester hóa là phản ứng thuận nghịch.\\
- d) Sai: Phổ IR dùng để xác định các nhóm chức, không xác định công thức phân tử.\\
$\Rightarrow$ Đáp án: a-Đ, b-S, c-Đ, d-S.
}
"""
    tex += r"""\end{ex}
"""

    # PHẦN III
    tex += r"""
\bigskip
\noindent\textbf{\large PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.}
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}
\vspace{0.2cm}
"""

    # C1 P3
    tex += r"""
\begin{ex}
Cho các chất sau: \ce{CH3-CH2-CHO} (1), \ce{CH2=CH-CHO} (2), \ce{(CH3)2CH-CHO} (3), \ce{CH2=CH-CH2-OH} (4). Chất phản ứng hoàn toàn với lượng dư [H] cùng tạo ra \ce{CH3-CH2-CH2-OH} là bao nhiêu chất?
"""
    if is_teacher:
        tex += r"""
\loigiai{
Gồm 3 chất: (1), (2), (4) khi bị khử đều tạo propan-1-ol.\\
$\Rightarrow$ Đáp số: \textbf{3}.
}
"""
    else:
        tex += r"""\vspace{0.8cm}
"""
    tex += r"""\end{ex}
"""

    # C2 P3
    tex += r"""
\begin{ex}
Cho các tác nhân phản ứng: (a) \ce{H2/Pd}, $t^\circ$; (b) \ce{LiAlH4}; (c) \ce{NaBH4}; (d) dung dịch bromine. Có bao nhiêu tác nhân ở trên có thể chuyển hóa RCHO thành \ce{RCH2OH}?
"""
    if is_teacher:
        tex += r"""
\loigiai{
Gồm 3 tác nhân khử: (a), (b), (c). Tác nhân (d) là chất oxi hóa.\\
$\Rightarrow$ Đáp số: \textbf{3}.
}
"""
    else:
        tex += r"""\vspace{0.8cm}
"""
    tex += r"""\end{ex}
"""

    # C3 P3
    tex += r"""
\begin{ex}
Cho các phát biểu sau:
\begin{enumerate}[(1)]
  \item Aldehyde phản ứng được với nước bromine.
  \item Ketone không phản ứng được với \ce{Cu(OH)2/OH^-}.
  \item Aldehyde tác dụng với dung dịch \ce{AgNO3/NH3} tạo ra bạc.
  \item Trong các hợp chất carbonyl, chỉ aldehyde bị khử bởi \ce{NaBH4}.
\end{enumerate}
Số phát biểu đúng là bao nhiêu?
"""
    if is_teacher:
        tex += r"""
\loigiai{
Có 3 phát biểu đúng: (1), (2), (3). Phát biểu (4) sai vì cả aldehyde và ketone đều bị khử bởi \ce{NaBH4}.\\
$\Rightarrow$ Đáp số: \textbf{3}.
}
"""
    else:
        tex += r"""\vspace{0.8cm}
"""
    tex += r"""\end{ex}
"""

    # C4 P3
    tex += r"""
\begin{ex}
Để điều chế \ce{CH3CHO} từ \ce{Al4C3} cần ít nhất bao nhiêu phương trình phản ứng?
"""
    if is_teacher:
        tex += r"""
\loigiai{
Cần ít nhất 3 phương trình: \ce{Al4C3 -> CH4 -> C2H2 -> CH3CHO}.\\
$\Rightarrow$ Đáp số: \textbf{3}.
}
"""
    else:
        tex += r"""\vspace{0.8cm}
"""
    tex += r"""\end{ex}
"""

    # C5 P3
    tex += r"""
\begin{ex}
Quế có vị cay, mùi thơm nồng, được sử dụng phổ biến làm gia vị, vị thuốc trong Đông y. Hợp chất hữu cơ cinnamaldehyde tạo mùi đặc trưng của quế. Trong phân tử cinnamaldehyde chứa vòng benzene có một nhóm thế. Cinnamaldehyde bị oxi hóa bởi thuốc thử Tollens và có đồng phân hình học dạng trans. Công thức cấu tạo của cinnamaldehyde là:
\begin{center}
\includegraphics[width=0.28\linewidth]{images_crop_hoa11_de17/cinnamaldehyde_trans.png}
\end{center}
Phân tử khối của cinnamaldehyde là bao nhiêu?
"""
    if is_teacher:
        tex += r"""
\loigiai{
Công thức phân tử: \ce{C9H8O}.\\
Phân tử khối: $M = 9 \times 12 + 8 \times 1 + 16 = 132$.\\
$\Rightarrow$ Đáp số: \textbf{132}.
}
"""
    else:
        tex += r"""\vspace{0.8cm}
"""
    tex += r"""\end{ex}
"""

    # C6 P3
    tex += r"""
\begin{ex}
Hợp chất hữu cơ X được dùng nhiều trong tổng hợp hữu cơ. Bằng phương pháp phân tích nguyên tố, người ta xác định được X chứa 62,07\% C; 10,34\% H; còn lại là O. Trên phổ MS của X, người ta thấy có peak ion phân tử $[\text{M}^+]$ có giá trị $m/z$ bằng 58. Trên phổ IR của X có một peak trong vùng 2380 -- 2695 cm$^{-1}$ và một peak trong vùng 1740 -- 1685 cm$^{-1}$. Chất X phản ứng với \ce{Cu(OH)2} trong môi trường kiềm để tạo ra kết tủa màu đỏ gạch. Xác định tổng số nhóm \ce{CH2} trong công thức cấu tạo của X.
"""
    if is_teacher:
        tex += r"""
\loigiai{
- $\%O = 100\% - 62,07\% - 10,34\% = 27,59\%$.\\
- $M_X = 58 \Rightarrow$ công thức phân tử là \ce{C3H6O}.\\
- Tín hiệu IR và phản ứng với \ce{Cu(OH)2} tạo \ce{Cu2O} đỏ gạch $\Rightarrow$ X là propanal: \ce{CH3-CH2-CHO}.\\
- Trong công thức \ce{CH3-CH2-CHO} có đúng 1 nhóm \ce{-CH2-}.\\
$\Rightarrow$ Đáp số: \textbf{1}.
}
"""
    else:
        tex += r"""\vspace{0.8cm}
"""
    tex += r"""\end{ex}
"""

    # PHẦN IV
    tex += r"""
\bigskip
\noindent\textbf{\large PHẦN IV. Câu hỏi tự luận.}
\textit{Thí sinh trình bày lời giải chi tiết vào giấy làm bài.}
\vspace{0.2cm}
"""

    # C1 P4
    tex += r"""
\begin{ex}
\textbf{(SBT - CTST)} Trong các phản ứng sau đây, phản ứng nào không thể xảy ra? Hoàn thành phương trình hoá học các phản ứng xảy ra.
\begin{enumerate}[a)]
  \item \ce{CH3CH2-CO-CH3 + [O] ->}
  \item \ce{CH3CH2-OH + [O] ->}
  \item \ce{C6H6 + Cl2 ->[{FeBr$_3$, $t^\circ$}]}
  \item \ce{(CH3)2C(OH)-CH2CH3 + [O] ->}
\end{enumerate}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- Phản ứng không thể xảy ra:\\
  + a) \ce{CH3CH2-CO-CH3} là ketone không bị oxi hóa thông thường.\\
  + d) \ce{(CH3)2C(OH)-CH2CH3} là alcohol bậc ba không bị oxi hóa.\\
- Hoàn thành các phản ứng xảy ra:\\
  b) \ce{CH3CH2-OH + O2 ->[{men giấm}] CH3COOH + H2O} (hoặc tác dụng CuO tạo \ce{CH3CHO}).\\
  c) \ce{C6H6 + Cl2 ->[{FeBr$_3$, $t^\circ$}] C6H5Cl + HCl}
}
"""
    else:
        tex += r"""
\vspace{0.2cm}
\noindent\dotfill\\[4pt]
\noindent\dotfill\\[4pt]
\noindent\dotfill\\[4pt]
\noindent\dotfill\\[4pt]
\noindent\dotfill
"""
    tex += r"""\end{ex}
"""

    # C2 P4
    tex += r"""
\begin{ex}
Cho thí nghiệm: Chuẩn độ acetic acid bằng dung dịch sodium hydroxide 0,15 M. Sau khi thực hiện thí nghiệm chuẩn độ này 3 lần, thu được bảng dưới đây:
\begin{center}
\begin{tabular}{|c|c|c|}
\hline
\textbf{Số lần} & \textbf{Thể tích acetic acid (mL)} & \textbf{Thể tích dung dịch \ce{NaOH} cần dùng (mL)} \\ \hline
1 & 6,0 & 20,0 \\ \hline
2 & 6,0 & 19,9 \\ \hline
3 & 6,0 & 20,0 \\ \hline
\end{tabular}
\end{center}
Xác định nồng độ mol của dung dịch acetic acid.
"""
    if is_teacher:
        tex += r"""
\loigiai{
- Thể tích trung bình của dung dịch \ce{NaOH} sau 3 lần đo:\\
  $V_{\ce{NaOH}} = \dfrac{20,0 + 19,9 + 20,0}{3} \approx 20,0\text{ mL}$.\\
- Phương trình phản ứng: \ce{CH3COOH + NaOH -> CH3COONa + H2O}.\\
- Áp dụng công thức: $V_{\text{acid}} \times C_{\text{acid}} = V_{\ce{NaOH}} \times C_{\ce{NaOH}}$.\\
$\Rightarrow C_{\text{acid}} = \dfrac{20,0 \times 0,15}{6,0} = 0,5\text{ M}$.\\
$\Rightarrow$ Đáp số: \textbf{0,5 M}.
}
"""
    else:
        tex += r"""
\vspace{0.2cm}
\noindent\dotfill\\[4pt]
\noindent\dotfill\\[4pt]
\noindent\dotfill\\[4pt]
\noindent\dotfill\\[4pt]
\noindent\dotfill
"""
    tex += r"""\end{ex}
"""

    # Bảng đáp án ở cuối bản giáo viên
    if is_teacher:
        tex += r"""
\newpage
\noindent\textbf{\large BẢNG ĐÁP ÁN THAM KHẢO PHẦN I}
\vspace{0.2cm}

\begin{center}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1. B} & \textbf{2. C} & \textbf{3. B} & \textbf{4. B} & \textbf{5. D} & \textbf{6. B} & \textbf{7. C} & \textbf{8. A} & \textbf{9. A} \\
\hline
\textbf{10. D} & \textbf{11. C} & \textbf{12. B} & \textbf{13. C} & \textbf{14. B} & \textbf{15. D} & \textbf{16. A} & \textbf{17. A} & \textbf{18. D} \\
\hline
\end{tabular}
\end{center}

\bigskip
\noindent\textbf{\large BẢNG ĐÁP ÁN THAM KHẢO PHẦN II}
\vspace{0.2cm}

\begin{center}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Ý a)} & \textbf{Ý b)} & \textbf{Ý c)} & \textbf{Ý d)} \\
\hline
\textbf{Câu 1} & Đ & Đ & S & S \\
\hline
\textbf{Câu 2} & Đ & Đ & S & S \\
\hline
\textbf{Câu 3} & Đ & S & S & Đ \\
\hline
\textbf{Câu 4} & Đ & S & Đ & S \\
\hline
\end{tabular}
\end{center}

\bigskip
\noindent\textbf{\large BẢNG ĐÁP ÁN THAM KHẢO PHẦN III}
\vspace{0.2cm}

\begin{center}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\
\hline
3 & 3 & 3 & 3 & 132 & 1 \\
\hline
\end{tabular}
\end{center}
"""

    tex += r"""
\bigskip
\begin{center}
\textbf{---------- HẾT ----------}
\end{center}
\end{document}
"""
    return tex

def compile_pdf(tex_path):
    cmd = f'pdflatex -interaction=nonstopmode "{os.path.basename(tex_path)}"'
    for _ in range(2):
        res = subprocess.run(
            cmd,
            cwd=os.path.dirname(tex_path),
            shell=True,
            capture_output=True,
            text=True
        )
    return res.returncode == 0

def cleanup(base_name):
    for ext in [".aux", ".log", ".out", ".thm"]:
        f = os.path.join(TEX_DIR, base_name + ext)
        if os.path.exists(f):
            try:
                os.remove(f)
            except:
                pass

def main():
    # 1. Học sinh
    hs_base = "De_Kiem_Tra_Hoa_11_De17_HocSinh"
    hs_tex = os.path.join(TEX_DIR, hs_base + ".tex")
    with open(hs_tex, "w", encoding="utf-8") as f:
        f.write(make_tex(is_teacher=False))
    print(f"[OK] Đã tạo file: {hs_tex}")

    # 2. Giáo viên
    gv_base = "De_Kiem_Tra_Hoa_11_De17_GiaoVien"
    gv_tex = os.path.join(TEX_DIR, gv_base + ".tex")
    with open(gv_tex, "w", encoding="utf-8") as f:
        f.write(make_tex(is_teacher=True))
    print(f"[OK] Đã tạo file: {gv_tex}")

    # Biên dịch PDF
    print("Đang biên dịch PDF Học sinh...")
    ok_hs = compile_pdf(hs_tex)
    cleanup(hs_base)
    print(f"Học sinh PDF: {'Thành công' if ok_hs else 'Thất bại'}")

    print("Đang biên dịch PDF Giáo viên...")
    ok_gv = compile_pdf(gv_tex)
    cleanup(gv_base)
    print(f"Giáo viên PDF: {'Thành công' if ok_gv else 'Thất bại'}")

if __name__ == "__main__":
    main()
