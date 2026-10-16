# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file LaTeX và biên dịch ra PDF cho ĐỀ SỐ 18 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De18_HocSinh.tex / .pdf
- De_Kiem_Tra_Hoa_11_De18_GiaoVien.tex / .pdf
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
	\textit{(Đề thi có """ + ("06" if is_teacher else "05") + r""" trang)}
\end{minipage}\hfill
\begin{minipage}[t]{0.54\textwidth}
	\centering\fontsize{10.5pt}{13pt}\selectfont
	\textbf{ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 6}\\
	\textbf{(HỢP CHẤT CARBONYL - CARBOXYLIC ACID)}\\
	\textbf{MÔN: HOÁ HỌC - LỚP 11}\\
	\textit{Thời gian làm bài: 50 phút, không kể thời gian phát đề}
	
	\smallskip
	\fbox{\textbf{MÃ ĐỀ: 118}}
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
Gọi tên một hợp chất có công thức cấu tạo như sau: \ce{(CH3)3C-CHO}
\choice
{2-ethyl-2-methylethanal}
{2,2-dimethylpentanal}
{Tert-butyletanal}
{""" + (r"\True " if is_teacher else "") + r"""2,2-dimethylpropanal}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Mạch chính dài nhất chứa nhóm \ce{-CHO} gồm 3 nguyên tử carbon (đánh số 1 từ nhóm \ce{-CHO}), tại C số 2 có 2 nhóm \ce{-CH3}.\\
Tên gọi thay thế là 2,2-dimethylpropanal.\\
$\Rightarrow$ Chọn phương án \textbf{D}.
}
"""
    tex += r"""\end{ex}
"""

    # C2
    tex += r"""
\begin{ex}
Công thức cấu tạo thu gọn của acetone là
\choice
{\ce{OHC-CHO}}
{""" + (r"\True " if is_teacher else "") + r"""\ce{CH3COCH3}}
{\ce{CH3CHO}}
{\ce{CH2=CHCHO}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Acetone (propan-2-one hay dimethyl ketone) có công thức cấu tạo thu gọn là \ce{CH3COCH3}.\\
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C3
    tex += r"""
\begin{ex}
Tên gọi của aldehyde có khối lượng phân tử nhỏ nhất là
\choice
{aldehyde acetic}
{""" + (r"\True " if is_teacher else "") + r"""methanal}
{acetone}
{methanol}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Aldehyde đơn giản nhất và có phân tử khối nhỏ nhất là \ce{HCHO} ($M = 30$), có tên là methanal (hay formaldehyde).\\
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C4
    tex += r"""
\begin{ex}
Phản ứng giữa \ce{CH3COCH3} với \ce{HCN} thuộc loại phản ứng
\choice
{phản ứng thuận nghịch}
{""" + (r"\True " if is_teacher else "") + r"""phản ứng cộng}
{phản ứng thế}
{phản ứng phân hủy}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Nhóm carbonyl $>$\ce{C=O} trong acetone tham gia phản ứng cộng nucleophile với \ce{HCN} tạo ra cyanohydrin:
$$\ce{CH3COCH3 + HCN -> (CH3)2C(OH)CN}$$
Đây là phản ứng cộng.\\
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C5
    tex += r"""
\begin{ex}
Điều kiện của phản ứng giữa aldehyde hoặc ketone với iodine trong dung dịch \ce{NaOH} tạo iodoform là
\choice
{""" + (r"\True " if is_teacher else "") + r"""aldehyde hoặc ketone có nhóm methyl cạnh nhóm carbonyl (\ce{CH3CO-})}
{aldehyde hoặc ketone có nhóm methyl cạnh nhóm carbonyl (\ce{C2H3CO-})}
{aldehyde hoặc ketone có nhóm methyl cạnh nhóm carbonyl (\ce{C2H5CO-})}
{aldehyde hoặc ketone có nhóm methyl cạnh nhóm carbonyl (\ce{C3H7CO-})}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Phản ứng iodoform (tạo kết tủa vàng \ce{CHI3}) xảy ra với các hợp chất carbonyl có chứa nhóm methyl ketone (\ce{CH3-C=O}) hoặc acetaldehyde (\ce{CH3-CHO}).\\
$\Rightarrow$ Chọn phương án \textbf{A}.
}
"""
    tex += r"""\end{ex}
"""

    # C6
    tex += r"""
\begin{ex}
\textit{(SBT - CTST)} Cho dãy các chất sau: \ce{HCHO}, \ce{CH3COOH}, \ce{CH3CHO}, \ce{HCOOH}, \ce{C2H5OH}. Số chất trong dãy có thể tham gia phản ứng tráng bạc là
\choice
{""" + (r"\True " if is_teacher else "") + r"""3}
{6}
{4}
{5}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Các chất có nhóm formyl (\ce{-CHO}) tham gia phản ứng tráng bạc. Trong dãy gồm 3 chất: \ce{HCHO}, \ce{CH3CHO} và \ce{HCOOH}.\\
$\Rightarrow$ Chọn phương án \textbf{A}.
}
"""
    tex += r"""\end{ex}
"""

    # C7
    tex += r"""
\begin{ex}
\textit{(SBT - CTST)} Có 3 dung dịch: \ce{CH3CHO}, \ce{CH3COOH}, \ce{HCOOH} đựng trong 3 lọ mất nhãn. Hóa chất có thể dùng để phân biệt ba dung dịch trên là
\choice
{quỳ tím, \ce{CuO}}
{quỳ tím, \ce{Na}}
{""" + (r"\True " if is_teacher else "") + r"""quỳ tím, dung dịch \ce{AgNO3} trong \ce{NH3} dư}
{dung dịch \ce{AgNO3} trong \ce{NH3} dư, \ce{CuO}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- Dùng quỳ tím: \ce{CH3CHO} không đổi màu quỳ; \ce{CH3COOH} và \ce{HCOOH} làm quỳ tím hóa đỏ.\\
- Cho dung dịch \ce{AgNO3/NH3} đun nóng vào 2 lọ acid: \ce{HCOOH} có nhóm \ce{-CHO} nên tráng bạc tạo kết tủa \ce{Ag} sáng bóng; \ce{CH3COOH} không hiện tượng.\\
$\Rightarrow$ Chọn phương án \textbf{C}.
}
"""
    tex += r"""\end{ex}
"""

    # C8
    tex += r"""
\begin{ex}
\textit{(SBT - CD)} Phản ứng nào sau đây thể hiện tính oxi hóa của propanal?
\choice
{\ce{C2H5CHO + 2[Ag(NH3)2]OH -> C2H5COONH4 + 3NH3 + 2Ag v + H2O}}
{\ce{C2H5CHO + Br2 + H2O -> C2H5COOH + 2HBr}}
{\ce{C2H5CHO + 2Cu(OH)2 + NaOH -> C2H5COONa + Cu2O v + 3H2O}}
{""" + (r"\True " if is_teacher else "") + r"""\ce{C2H5CHO + 2[H] ->[{LiAlH$_4$}] CH3CH2CH2OH}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Trong phản ứng khử với \ce{LiAlH4}, propanal nhận electron/H (số oxi hóa carbon nhóm \ce{-CHO} giảm từ $+1$ xuống $-1$ trong alcohol bậc 1), do đó propanal thể hiện tính oxi hóa.\\
$\Rightarrow$ Chọn phương án \textbf{D}.
}
"""
    tex += r"""\end{ex}
"""

    # C9
    tex += r"""
\begin{ex}
\textit{(SBT - CD)} Khi cho ethanal phản ứng với \ce{Cu(OH)2} trong môi trường kiềm ở nhiệt độ thích hợp, hiện tượng nào sau đây sẽ xảy ra?
\choice
{\ce{Cu(OH)2} bị tan ra, tạo dung dịch màu xanh}
{Có mùi chua của giấm, do phản ứng sinh ra acetic acid}
{""" + (r"\True " if is_teacher else "") + r"""Tạo kết tủa đỏ gạch do phản ứng sinh ra \ce{Cu2O}}
{Sinh ra \ce{CuO} màu đen}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Ethanal khử \ce{Cu(OH)2} trong môi trường kiềm đun nóng tạo kết tủa đỏ gạch \ce{Cu2O}:
$$\ce{CH3CHO + 2Cu(OH)2 + NaOH ->[{t$^\circ$}] CH3COONa + Cu2O v (đỏ gạch) + 3H2O}$$
$\Rightarrow$ Chọn phương án \textbf{C}.
}
"""
    tex += r"""\end{ex}
"""

    # C10
    tex += r"""
\begin{ex}
Có thể dùng thuốc thử nào trong các chất dưới đây để nhận biết các chất: ethyl alcohol, glycerol, acetic aldehyde đựng trong ba lọ mất nhãn?
\choice
{""" + (r"\True " if is_teacher else "") + r"""\ce{Cu(OH)2} trong môi trường kiềm}
{Quỳ tím}
{Kim loại sodium}
{Dung dịch \ce{AgNO3/NH3}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- Cho \ce{Cu(OH)2} ở nhiệt độ thường: glycerol hòa tan tạo dung dịch màu xanh lam thẫm (phức đồng); ethyl alcohol và ethanal không phản ứng.\\
- Đun nóng 2 mẫu còn lại với \ce{Cu(OH)2/OH^-}: ethanal tạo kết tủa đỏ gạch \ce{Cu2O}; ethyl alcohol không hiện tượng.\\
$\Rightarrow$ Chọn phương án \textbf{A}.
}
"""
    tex += r"""\end{ex}
"""

    # C11
    tex += r"""
\begin{ex}
Có bao nhiêu aldehyde hai chức có công thức đơn giản nhất là \ce{C2H3O}?
\choice
{""" + (r"\True " if is_teacher else "") + r"""2}
{4}
{1}
{3}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Aldehyde hai chức có 2 nhóm \ce{-CHO} nên chứa 2 nguyên tử O trong phân tử.\\
Vì công thức đơn giản nhất là \ce{C2H3O} nên CTPT là \ce{(C2H3O)2} = \ce{C4H6O2}.\\
Các đồng phân aldehyde hai chức:\\
(1) \ce{OHC-CH2-CH2-CHO} (butanedial)\\
(2) \ce{OHC-CH(CH3)-CHO} (2-methylpropanedial)\\
$\Rightarrow$ Có 2 đồng phân thỏa mãn. Chọn phương án \textbf{A}.
}
"""
    tex += r"""\end{ex}
"""

    # C12
    tex += r"""
\begin{ex}
Có bao nhiêu phản ứng xảy ra khi cho các chất \ce{C6H5OH}, \ce{NaHCO3}, \ce{NaOH}, \ce{HCl} tác dụng với nhau từng đôi một?
\choice
{3}
{""" + (r"\True " if is_teacher else "") + r"""4}
{5}
{6}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Có 4 phản ứng xảy ra khi trộn từng đôi một:\\
(1) \ce{C6H5OH + NaOH -> C6H5ONa + H2O}\\
(2) \ce{NaHCO3 + NaOH -> Na2CO3 + H2O}\\
(3) \ce{NaHCO3 + HCl -> NaCl + CO2 ^ + H2O}\\
(4) \ce{NaOH + HCl -> NaCl + H2O}\\
(Phenol tính acid yếu hơn \ce{H2CO3} nên không phản ứng với \ce{NaHCO3}).\\
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C13
    tex += r"""
\begin{ex}
Cho các hợp chất có công thức sau: \ce{CH3OH}; \ce{CH3CHO}; \ce{HCOOH}; \ce{C2H5NH2}; \ce{C6H6}; \ce{CH3COCH3}; \ce{C2H6}; \ce{C2H5Cl}. Trong các hợp chất trên, có bao nhiêu chất chứa nhóm carbonyl trong phân tử?
\choice
{""" + (r"\True " if is_teacher else "") + r"""3}
{4}
{2}
{5}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Các chất chứa nhóm carbonyl ($>$\ce{C=O}) trong dãy gồm:\\
- \ce{CH3CHO} (aldehyde)\\
- \ce{CH3COCH3} (ketone)\\
- \ce{HCOOH} (carboxylic acid chứa nhóm C=O)\\
$\Rightarrow$ Có 3 chất. Chọn phương án \textbf{A}.
}
"""
    tex += r"""\end{ex}
"""

    # C14
    tex += r"""
\begin{ex}
Quế có vị cay, mùi thơm nồng, được sử dụng phổ biến làm gia vị, vị thuốc trong Đông y. Hợp chất hữu cơ X tạo mùi đặc trưng của quế, có công thức phân tử là \ce{C9H8O}. Trong phân tử X chứa vòng benzene. X tham gia phản ứng tráng bạc và có đồng phân hình học dạng trans. Xác định công thức cấu tạo của X.
\begin{center}
\includegraphics[width=0.88\linewidth]{images_crop_hoa11_de18/""" + ("c14_grid_teacher.png" if is_teacher else "c14_grid_student.png") + r"""}
\end{center}
\choice
{""" + (r"\True " if is_teacher else "") + r"""Phương án A}
{Phương án B}
{Phương án C}
{Phương án D}
"""
    if is_teacher:
        tex += r"""
\loigiai{
X tráng bạc nên X là aldehyde (chứa nhóm \ce{-CHO}), có đồng phân hình học dạng trans, công thức \ce{C9H8O} chứa vòng benzene.\\
$\Rightarrow$ X là trans-cinnamaldehyde (\ce{C6H5-CH=CH-CHO}), tương ứng với cấu trúc A.\\
$\Rightarrow$ Chọn phương án \textbf{A}.
}
"""
    tex += r"""\end{ex}
"""

    # C15
    tex += r"""
\begin{ex}
Một thí nghiệm được mô tả như hình sau đây:
\begin{center}
\includegraphics[width=0.48\linewidth]{images_crop_hoa11_de18/c15_experiment.png}
\end{center}
Chất lỏng thu được ở ống nghiệm B có mùi táo, có tên gọi là
\choice
{ethyl formate}
{""" + (r"\True " if is_teacher else "") + r"""methyl propionate}
{ethyl propionate}
{propyl formate}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Ống nghiệm A chứa \ce{CH3OH} và \ce{CH3CH2COOH} xúc tác \ce{H2SO4} đặc đun nóng.\\
Phản ứng ester hóa tạo ester methyl propionate có mùi táo:\\
$$\ce{CH3CH2COOH + CH3OH <=>[{H$_2$SO$_4$, $t^\circ$}] CH3CH2COOCH3 + H2O}$$
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C16
    tex += r"""
\begin{ex}
Cho giá trị $K_{\text{a}}$ của một số carboxylic acid ở 25 $^\circ$C dưới đây:
\begin{center}
\begin{tabular}{|l|c|c|c|c|}
\hline
\textbf{Carboxylic acid} & \ce{HCOOH} & \ce{CH3COOH} & \ce{CH3CH2COOH} & \ce{C6H5COOH} \\ \hline
$K_{\text{a}}$ & $17,72 \times 10^{-5}$ & $1,75 \times 10^{-5}$ & $1,35 \times 10^{-5}$ & $6,25 \times 10^{-5}$ \\ \hline
\end{tabular}
\end{center}
Tính acid của carboxylic acid nào là mạnh nhất?
\choice
{""" + (r"\True " if is_teacher else "") + r"""\ce{HCOOH}}
{\ce{CH3COOH}}
{\ce{CH3CH2COOH}}
{\ce{C6H5COOH}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Hằng số phân li acid $K_{\text{a}}$ càng lớn thì tính acid càng mạnh.\\
Trong bảng, \ce{HCOOH} có $K_{\text{a}} = 17,72 \times 10^{-5}$ là lớn nhất nên có tính acid mạnh nhất.\\
$\Rightarrow$ Chọn phương án \textbf{A}.
}
"""
    tex += r"""\end{ex}
"""

    # C17
    tex += r"""
\begin{ex}
\textit{(SBT - CD)} Để trung hòa 40 mL giấm ăn cần 25 mL dung dịch \ce{NaOH} 1 M. Biết khối lượng riêng của giấm ăn xấp xỉ là 1 g/mL. Mẫu giấm ăn này có nồng độ phần trăm là
\choice
{3,5\%}
{""" + (r"\True " if is_teacher else "") + r"""3,75\%}
{4\%}
{5\%}
"""
    if is_teacher:
        tex += r"""
\loigiai{
$n_{\ce{NaOH}} = 0,025 \times 1 = 0,025\text{ mol}$.\\
Phản ứng: \ce{CH3COOH + NaOH -> CH3COONa + H2O}\\
$\Rightarrow n_{\ce{CH3COOH}} = 0,025\text{ mol} \Rightarrow m_{\ce{CH3COOH}} = 0,025 \times 60 = 1,5\text{ gam}$.\\
$m_{\text{dd giấm}} = V \times D = 40 \times 1 = 40\text{ gam}$.\\
$C\% = \dfrac{1,5}{40} \times 100\% = 3,75\%$.\\
$\Rightarrow$ Chọn phương án \textbf{B}.
}
"""
    tex += r"""\end{ex}
"""

    # C18
    tex += r"""
\begin{ex}
Đun 12 gam acetic acid với một lượng dư ethanol (\ce{H2SO4} đặc làm xúc tác). Đến khi dừng thí nghiệm thu được 8,36 gam ester. Hiệu suất của phản ứng ester hóa là
\choice
{65,15\%}
{54,54\%}
{34,84\%}
{""" + (r"\True " if is_teacher else "") + r"""47,50\%}
"""
    if is_teacher:
        tex += r"""
\loigiai{
$n_{\ce{CH3COOH}} = \dfrac{12}{60} = 0,2\text{ mol}$.\\
Ester tạo thành là ethyl acetate \ce{CH3COOC2H5} ($M = 88$).\\
$n_{\text{ester thu}} = \dfrac{8,36}{88} = 0,095\text{ mol}$.\\
Theo lí thuyết: $n_{\text{ester lt}} = 0,2\text{ mol}$.\\
$H = \dfrac{0,095}{0,2} \times 100\% = 47,50\%$.\\
$\Rightarrow$ Chọn phương án \textbf{D}.
}
"""
    tex += r"""\end{ex}
"""

    # =========================================================================
    # PHẦN II
    # =========================================================================
    tex += r"""
\vspace{0.3cm}
\noindent\textbf{\large PHẦN II. Câu trắc nghiệm đúng sai.}
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{0.2cm}
"""

    # C1 P2
    tex += r"""
\begin{ex}
Muscone là hợp chất tạo nên mùi thơm đặc trưng của xạ hương, có công thức cấu tạo như sau:
\begin{center}
\includegraphics[width=0.35\linewidth]{images_crop_hoa11_de18/muscone.png}
\end{center}
\choiceTF
{""" + (r"\True " if is_teacher else "") + r"""Muscone thuộc hợp chất ketone}
{""" + (r"\True " if is_teacher else "") + r"""Muscone có công thức phân tử là \ce{C16H30O}}
{Muscone có thể tham gia phản ứng tráng bạc với thuốc thử Tollens}
{""" + (r"\True " if is_teacher else "") + r"""Muscone có nhiệt độ sôi cao hơn so với ethanal}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- a) \textbf{Đúng}: Phân tử có nhóm $>$\ce{C=O} gắn với 2 nguyên tử carbon trong vòng kín nên là ketone.\\
- b) \textbf{Đúng}: Vòng 15 cạnh + 1 nhóm \ce{-CH3} = 16 nguyên tử C; 1 vòng và 1 liên kết $\pi$ ($k = 2$) $\Rightarrow$ CTPT là \ce{C16H30O} ($M = 238$).\\
- c) \textbf{Sai}: Ketone không tham gia phản ứng tráng bạc với thuốc thử Tollens.\\
- d) \textbf{Đúng}: Muscone ($M = 238$, sôi ở khoảng 328 $^\circ$C) có phân tử khối lớn hơn rất nhiều so với ethanal ($M = 44$, sôi ở 20,8 $^\circ$C) nên nhiệt độ sôi cao hơn rất nhiều.
}
"""
    tex += r"""\end{ex}
"""

    # C2 P2
    tex += r"""
\begin{ex}
Trong nọc độc của một số côn trùng như kiến, ong, muỗi... có chứa hợp chất acid X gây bỏng rát và ngứa.
\choiceTF
{X là acetic acid, có công thức cấu tạo là \ce{CH3COOH}}
{""" + (r"\True " if is_teacher else "") + r"""X là formic acid, có công thức phân tử là \ce{HCOOH}}
{""" + (r"\True " if is_teacher else "") + r"""Khi bị côn trùng đốt thì nên bôi vôi vào vết côn trùng đốt nó sẽ trung hòa lượng acid X làm cho vết đốt đỡ ngứa rát: \ce{2HCOOH + Ca(OH)2 -> (HCOO)2Ca + 2H2O}}
{""" + (r"\True " if is_teacher else "") + r"""Cho dung dịch acid X vào ống nghiệm chứa dung dịch \ce{AgNO3/NH3}, sau đó đun nóng nhẹ thì thấy thành ống nghiệm sáng bóng như gương}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- a) \textbf{Sai}: Acid trong nọc kiến là formic acid (\ce{HCOOH}), không phải acetic acid.\\
- b) \textbf{Đúng}: Acid X là formic acid có CTPT là \ce{HCOOH}.\\
- c) \textbf{Đúng}: Vôi tôi \ce{Ca(OH)2} có tính base giúp trung hòa acid \ce{HCOOH} giảm sưng tấy.\\
- d) \textbf{Đúng}: \ce{HCOOH} có nhóm \ce{-CHO} nên có khả năng tráng bạc tạo gương sáng.
}
"""
    tex += r"""\end{ex}
"""

    # C3 P2
    tex += r"""
\begin{ex}
Tiến hành phản ứng tráng bạc aldehyde acetic với thuốc thử Tollens theo các bước sau đây:\\
\textit{Bước 1:} Rửa sạch ống nghiệm bằng cách nhỏ vài giọt dung dịch kiềm \ce{NaOH}, đun nóng nhẹ, tráng đều, sau đó đổ đi và tráng lại bằng nước cất.\\
\textit{Bước 2:} Nhỏ vào ống nghiệm 2 giọt dung dịch \ce{AgNO3} và 1 giọt dung dịch \ce{NH3}, xuất hiện kết tủa nâu xám, nhỏ tiếp dung dịch \ce{NH3} đến khi kết tủa tan hết.\\
\textit{Bước 3:} Nhỏ tiếp vào ống nghiệm 2 giọt dung dịch \ce{CH3CHO}.\\
\textit{Bước 4:} Lắc đều ống nghiệm, đun cách thủy (trong cốc nước nóng) vài phút ở 60 -- 70 $^\circ$C.
\choiceTF
{""" + (r"\True " if is_teacher else "") + r"""Sau bước 4 quan sát thấy thành ống nghiệm sáng bóng như gương}
{Ở bước 3 nếu thay aldehyde acetic bằng acetone thì thu được hiện tượng tương tự}
{Ở bước 4 xảy ra phản ứng oxi hóa -- khử trong đó aldehyde acetic là chất bị khử}
{""" + (r"\True " if is_teacher else "") + r"""Sản phẩm hữu cơ thu được trong dung dịch sau bước 4 có công thức phân tử là \ce{C2H7NO2}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- a) \textbf{Đúng}: Bạc kim loại tạo thành bám vào thành ống nghiệm tạo lớp gương sáng.\\
- b) \textbf{Sai}: Acetone là ketone không tham gia phản ứng tráng bạc.\\
- c) \textbf{Sai}: \ce{CH3CHO} đóng vai trò là chất khử (bị oxi hóa thành ion ammonium acetate).\\
- d) \textbf{Đúng}: Muối \ce{CH3COONH4} có công thức phân tử là \ce{C2H7NO2}.
}
"""
    tex += r"""\end{ex}
"""

    # C4 P2
    tex += r"""
\begin{ex}
Hãy chọn các nhận định về hiện tượng và phương trình hóa học xảy ra đúng trong các trường hợp sau:
\choiceTF
{""" + (r"\True " if is_teacher else "") + r"""Cho acetic aldehyde tác dụng với nước bromine thấy nước bromine bị mất màu do xảy ra phản ứng: \ce{CH3CHO + Br2 + H2O -> CH3COOH + 2HBr}}
{""" + (r"\True " if is_teacher else "") + r"""Cho formic aldehyde tác dụng với dung dịch \ce{AgNO3} trong \ce{NH3} dư thấy xuất hiện kết tủa trắng sáng do xảy ra phản ứng: \ce{HCHO + 4[Ag(NH3)2]OH ->[{t$^\circ$}] (NH4)2CO3 + 4Ag v + 6NH3 + 2H2O}}
{Cho acetone tác dụng với \ce{I2} trong dung dịch \ce{NaOH} dư không có hiện tượng gì xảy ra}
{Cho acetic aldehyde tác dụng với \ce{Cu(OH)2} trong \ce{NaOH} dư thấy xuất hiện kết tủa đỏ gạch do xảy ra phản ứng: \ce{C2H5CHO + 2Cu(OH)2 + NaOH ->[{t$^\circ$}] C2H5COONa + Cu2O v + 3H2O}}
"""
    if is_teacher:
        tex += r"""
\loigiai{
- a) \textbf{Đúng}: Acetic aldehyde bị nước bromine oxi hóa thành acetic acid làm mất màu nước bromine.\\
- b) \textbf{Đúng}: \ce{HCHO} phản ứng tráng bạc tỉ lệ 1 : 4 \ce{Ag} tạo kết tủa bạc kim loại sáng bóng.\\
- c) \textbf{Sai}: Acetone phản ứng iodoform với \ce{I2/NaOH} sinh ra kết tủa vàng \ce{CHI3}.\\
- d) \textbf{Sai}: Acetic aldehyde là \ce{CH3CHO}, phương trình trên viết sai chất phản ứng (viết thành \ce{C2H5CHO}).
}
"""
    tex += r"""\end{ex}
"""

    # =========================================================================
    # PHẦN III
    # =========================================================================
    tex += r"""
\vspace{0.3cm}
\noindent\textbf{\large PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.}
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}
\vspace{0.2cm}
"""

    # C1 P3
    tex += r"""
\begin{ex}
Cho các hóa chất sau: dung dịch \ce{HCN}, dung dịch \ce{Br2}, dung dịch \ce{NaOH}, dung dịch \ce{[Ag(NH3)2]OH}, \ce{LiAlH4}, \ce{Cu(OH)2/NaOH}, \ce{I2/NaOH}. Có bao nhiêu chất phản ứng được với acetone ở điều kiện nhiệt độ thích hợp?
"""
    if is_teacher:
        tex += r"""
\loigiai{
3 chất phản ứng được với acetone (\ce{CH3COCH3}) là:
\begin{enumerate}[(1)]
  \item Dung dịch \ce{HCN}: \ce{CH3COCH3 + HCN -> (CH3)2C(OH)CN} (phản ứng cộng nucleophile)
  \item \ce{LiAlH4}: \ce{CH3COCH3 + 2[H] ->[{LiAlH$_4$}] CH3CH(OH)CH3} (khử thành propan-2-ol)
  \item Dung dịch \ce{I2/NaOH}: \ce{CH3COCH3 + 3I2 + 4NaOH -> CH3COONa + CHI3 v (vàng) + 3NaI + 3H2O}
\end{enumerate}
$\Rightarrow$ Đáp số: \textbf{3}.
}
"""
    else:
        tex += r"""
\vspace{0.2cm}
\noindent\textbf{Đáp số:} \dotfill
"""
    tex += r"""\end{ex}
"""

    # C2 P3
    tex += r"""
\begin{ex}
Aldehyde (X) no, đơn chức, mạch hở không nhánh. Phân tử khối của (X) được xác định thông qua kết quả phổ khối lượng với peak ion phân tử có giá trị $m/z$ lớn nhất trong hình dưới đây. Khối lượng mol của X là bao nhiêu g/mol?
\begin{center}
\includegraphics[width=0.45\linewidth]{images_crop_hoa11_de18/ms_spectrum.png}
\end{center}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Trên phổ MS, peak ion phân tử $[\text{M}^+]$ có giá trị $m/z$ lớn nhất là 72.\\
$\Rightarrow$ Phân tử khối của X là 72 amu, hay khối lượng mol $M_{\text{X}} = 72\text{ g/mol}$ (ứng với \ce{CH3CH2CH2CHO}).\\
$\Rightarrow$ Đáp số: \textbf{72}.
}
"""
    else:
        tex += r"""
\vspace{0.2cm}
\noindent\textbf{Đáp số:} \dotfill
"""
    tex += r"""\end{ex}
"""

    # C3 P3
    tex += r"""
\begin{ex}
Cho các chất sau: \ce{Cu}, \ce{Na}, \ce{Mg}, \ce{KOH}, \ce{Na2CO3}, \ce{NaCl}, \ce{HCOOH}, \ce{C2H5OH}. Có bao nhiêu chất phản ứng được với \ce{CH3COOH}?
"""
    if is_teacher:
        tex += r"""
\loigiai{
Có 5 chất phản ứng được với \ce{CH3COOH} là: \ce{Na}, \ce{Mg}, \ce{KOH}, \ce{Na2CO3}, \ce{C2H5OH}.
\begin{enumerate}[(1)]
  \item \ce{2CH3COOH + 2Na -> 2CH3COONa + H2 ^}
  \item \ce{2CH3COOH + Mg -> (CH3COO)2Mg + H2 ^}
  \item \ce{CH3COOH + KOH -> CH3COOK + H2O}
  \item \ce{2CH3COOH + Na2CO3 -> 2CH3COONa + CO2 ^ + H2O}
  \item \ce{CH3COOH + C2H5OH <=>[{H$_2$SO$_4$, $t^\circ$}] CH3COOC2H5 + H2O}
\end{enumerate}
$\Rightarrow$ Đáp số: \textbf{5}.
}
"""
    else:
        tex += r"""
\vspace{0.2cm}
\noindent\textbf{Đáp số:} \dotfill
"""
    tex += r"""\end{ex}
"""

    # C4 P3
    tex += r"""
\begin{ex}
Hợp chất carboxylic acid X có công thức đơn giản nhất là \ce{CH2O}. Phổ MS cho thấy X có các tín hiệu sau:
\begin{center}
\begin{tabular}{|c|c|}
\hline
\textbf{$m/z$} & \textbf{Cường độ tương đối (\%)} \\ \hline
29 & 19 \\ \hline
31 & 100 \\ \hline
60 & 39 \\ \hline
\end{tabular}
\end{center}
Biết mảnh $[\text{M}^+]$ có giá trị $m/z$ lớn nhất. Tổng số nguyên tử trong một phân tử của X là bao nhiêu?
"""
    if is_teacher:
        tex += r"""
\loigiai{
- Giá trị $m/z$ lớn nhất là 60 $\Rightarrow$ phân tử khối của X là 60 amu.\\
- Công thức đơn giản nhất \ce{CH2O} ($M = 30$) $\Rightarrow$ CTPT có dạng \ce{(CH2O)}_n.\\
  Ta có: $30n = 60 \Rightarrow n = 2 \Rightarrow$ CTPT là \ce{C2H4O2} (acetic acid \ce{CH3COOH}).\\
- Tổng số nguyên tử trong phân tử \ce{C2H4O2} là: $2 + 4 + 2 = 8$ nguyên tử.\\
$\Rightarrow$ Đáp số: \textbf{8}.
}
"""
    else:
        tex += r"""
\vspace{0.2cm}
\noindent\textbf{Đáp số:} \dotfill
"""
    tex += r"""\end{ex}
"""

    # C5 P3
    tex += r"""
\begin{ex}
Nước ép chanh chứa khoảng 5\% (khoảng 0,3 mol/lít) citric acid, điều này giúp chanh có vị chua và độ pH từ 2 -- 3. Citric acid có công thức cấu tạo là:
\begin{center}
\includegraphics[width=0.35\linewidth]{images_crop_hoa11_de18/citric_acid.png}
\end{center}
Phân tử khối của citric acid là bao nhiêu amu?
"""
    if is_teacher:
        tex += r"""
\loigiai{
Từ CTCT, citric acid có CTPT là \ce{C6H8O7}.\\
Phân tử khối: $M = 12 \times 6 + 1 \times 8 + 16 \times 7 = 72 + 8 + 112 = 192\text{ amu}$.\\
$\Rightarrow$ Đáp số: \textbf{192}.
}
"""
    else:
        tex += r"""
\vspace{0.2cm}
\noindent\textbf{Đáp số:} \dotfill
"""
    tex += r"""\end{ex}
"""

    # C6 P3
    tex += r"""
\begin{ex}
Thực hiện phản ứng oxi hóa 4,958 L \ce{C2H4} (đkc) bằng \ce{O2} (xúc tác \ce{PdCl2}, \ce{CuCl2}) thu được chất X đơn chức. Cho toàn bộ lượng chất X tác dụng với hydrogen cyanide (\ce{HCN}) dư, thu được 7,1 gam \ce{CH3CH(CN)OH} (cyanohydrin). Hiệu suất của quá trình tạo \ce{CH3CH(CN)OH} từ \ce{C2H4} là $x$\%. Giá trị của $x$ là bao nhiêu?
"""
    if is_teacher:
        tex += r"""
\loigiai{
- $n_{\ce{C2H4}} = \dfrac{4,958}{24,79} = 0,2\text{ mol}$.\\
- Khối lượng mol \ce{CH3CH(CN)OH} là $M = 71\text{ g/mol} \Rightarrow n_{\text{cyanohydrin}} = \dfrac{7,1}{71} = 0,1\text{ mol}$.\\
- Sơ đồ phản ứng:\\
  $$\ce{2C2H4 + O2 ->[{PdCl$_2$, CuCl$_2$}] 2CH3CHO}$$
  $$\ce{CH3CHO + HCN -> CH3CH(CN)OH}$$
  Theo lí thuyết: $n_{\text{cyanohydrin lt}} = n_{\ce{C2H4}} = 0,2\text{ mol}$.\\
- Hiệu suất: $x\% = \dfrac{0,1}{0,2} \times 100\% = 50\% \Rightarrow x = 50$.\\
$\Rightarrow$ Đáp số: \textbf{50}.
}
"""
    else:
        tex += r"""
\vspace{0.2cm}
\noindent\textbf{Đáp số:} \dotfill
"""
    tex += r"""\end{ex}
"""

    # =========================================================================
    # PHẦN IV
    # =========================================================================
    tex += r"""
\vspace{0.3cm}
\noindent\textbf{\large PHẦN IV. Tự luận.}
\textit{Thí sinh trình bày bài giải chi tiết vào giấy làm bài.}
\vspace{0.2cm}
"""

    # C1 P4
    tex += r"""
\begin{ex}
\textit{(1,0 điểm)} Hoàn thành dãy chuyển hóa sau bằng các phương trình hóa học:
$$\ce{Ethane ->[(1)] Ethyl chloride ->[(2)] Ethanol ->[(3)] Ethanal ->[(4)] Acetic acid}$$
"""
    if is_teacher:
        tex += r"""
\loigiai{
Các phương trình hóa học thực hiện dãy chuyển hóa:
\begin{enumerate}[(1)]
  \item \ce{CH3-CH3 + Cl2 ->[{as (1:1)}] CH3-CH2Cl + HCl}
  \item \ce{CH3-CH2Cl + NaOH ->[{t$^\circ$}] CH3-CH2OH + NaCl}
  \item \ce{CH3-CH2OH + CuO ->[{t$^\circ$}] CH3-CHO + Cu v + H2O}
  \item \ce{CH3-CHO + Br2 + H2O -> CH3-COOH + 2HBr}\\
        \textit{(hoặc: \ce{2CH3CHO + O2 ->[{xt, t$^\circ$}] 2CH3COOH})}
\end{enumerate}
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
\textit{(1,0 điểm)} Để xác định được hợp chất hữu cơ X, người ta thực hiện các phân tích sau:\\
-- Bằng phương pháp phân tích nguyên tố, xác định được X chứa \%C = 40,0\%; \%H = 6,67\%, còn lại là oxygen.\\
-- Bằng phương pháp đo phổ MS, kết quả cho thấy trên phổ xuất hiện peak ion phân tử $[\text{M}^+]$ có giá trị $m/z$ bằng 60.\\
-- Bằng phương pháp đo phổ IR, kết quả cho thấy trên phổ xuất hiện peak rộng trong vùng từ 2 500 đến 3 200 cm$^{-1}$, và một peak hấp thụ mạnh ở 1 745 cm$^{-1}$.\\
Xác định công thức phân tử và công thức cấu tạo thu gọn của hợp chất X.
"""
    if is_teacher:
        tex += r"""
\loigiai{
\textbf{1. Xác định công thức phân tử của X:}\\
- Phần trăm khối lượng của oxygen: $\%O = 100\% - 40,0\% - 6,67\% = 53,33\%$.\\
- Peak ion phân tử $[\text{M}^+]$ có $m/z = 60 \Rightarrow M_{\text{X}} = 60\text{ g/mol}$.\\
- Gọi công thức phân tử của X là \ce{C_xH_yO_z}:\\
  $$x = \dfrac{40,0 \times 60}{12 \times 100} = 2;\quad y = \dfrac{6,67 \times 60}{1 \times 100} = 4;\quad z = \dfrac{53,33 \times 60}{16 \times 100} = 2$$
  $\Rightarrow$ CTPT của X là \ce{C2H4O2}.\\[6pt]
\textbf{2. Xác định công thức cấu tạo thu gọn của X:}\\
- Trên phổ IR:\\
  + Peak rộng vùng 2 500 -- 3 200 cm$^{-1}$ đặc trưng cho liên kết \ce{O-H} của carboxylic acid.\\
  + Peak hấp thụ mạnh ở 1 745 cm$^{-1}$ đặc trưng cho liên kết \ce{C=O} trong nhóm carboxyl (\ce{-COOH}).\\
- Kết hợp CTPT \ce{C2H4O2} chứa nhóm \ce{-COOH} $\Rightarrow$ X là carboxylic acid no, đơn chức 2 carbon.\\
$\Rightarrow$ Công thức cấu tạo thu gọn của X là \ce{CH3COOH} (acetic acid).
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
\textbf{1. D} & \textbf{2. B} & \textbf{3. B} & \textbf{4. B} & \textbf{5. A} & \textbf{6. A} & \textbf{7. C} & \textbf{8. D} & \textbf{9. C} \\
\hline
\textbf{10. A} & \textbf{11. A} & \textbf{12. B} & \textbf{13. A} & \textbf{14. A} & \textbf{15. B} & \textbf{16. A} & \textbf{17. B} & \textbf{18. D} \\
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
\textbf{Câu 1} & Đ & Đ & S & Đ \\
\hline
\textbf{Câu 2} & S & Đ & Đ & Đ \\
\hline
\textbf{Câu 3} & Đ & S & S & Đ \\
\hline
\textbf{Câu 4} & Đ & Đ & S & S \\
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
3 & 72 & 5 & 8 & 192 & 50 \\
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
    hs_base = "De_Kiem_Tra_Hoa_11_De18_HocSinh"
    hs_tex = os.path.join(TEX_DIR, hs_base + ".tex")
    with open(hs_tex, "w", encoding="utf-8") as f:
        f.write(make_tex(is_teacher=False))
    print(f"[OK] Đã tạo file: {hs_tex}")

    # 2. Giáo viên
    gv_base = "De_Kiem_Tra_Hoa_11_De18_GiaoVien"
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
