# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file LaTeX và biên dịch ra PDF cho ĐỀ SỐ 11 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De11_HocSinh.tex / .pdf
- De_Kiem_Tra_Hoa_11_De11_GiaoVien.tex / .pdf
Chương 4: HYDROCARBON
Tuân thủ Form chuẩn GEMINI.md, sử dụng gói ex_test.sty và mhchem.
"""

import os
import sys
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
import subprocess
import shutil

TEX_DIR = r"c:\Antigravity_Thanh\San_Pham"

if not os.path.exists(os.path.join(TEX_DIR, "ex_test.sty")):
    shutil.copyfile(r"c:\Antigravity_Thanh\ex_test.sty", os.path.join(TEX_DIR, "ex_test.sty"))

def make_tex(is_teacher=False):
    ver_tag = "gv" if is_teacher else "hs"
    ans_file = f"ans-hoa11-de11{ver_tag}"
    
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
	\textbf{ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 4 (HYDROCARBON)}\\
	\textbf{MÔN: HOÁ HỌC - LỚP 11}\\
	\textit{Thời gian làm bài: 50 phút, không kể thời gian phát đề}
	
	\smallskip
	\fbox{\textbf{MÃ ĐỀ: 111}}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Mn = 55; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108. Thể tích khí đo ở đkc 25 $^\circ\text{C}$, 1 bar là 24,79 L/mol).}
\smallskip

\noindent\textbf{\large PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.}
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{0.2cm}

\Opensolutionfile{ans}[""" + ans_file + r"""]

\begin{ex}
[KNTT - SBT] $\ce{(CH3)2CH-CH3}$ có tên theo danh pháp thay thế là
\choice
{\True 2-methylpropane}
{isobutane}
{butane}
{2-methylbutane}
\loigiai{
Mạch chính dài nhất chứa 3 nguyên tử carbon (propane), nhóm methyl thế ở vị trí carbon số 2. Danh pháp thay thế là 2-methylpropane. (Isobutane là tên thông thường).
}
\end{ex}

\begin{ex}
(SBT - CTST). Đồ thị dưới đây thể hiện mối tương quan giữa nhiệt độ sôi và số nguyên tử carbon trong phân tử alkane không phân nhánh được biểu diễn như sau:
\begin{center}
    \includegraphics[width=8.5cm]{images_crop_hoa11_de11/cau2_dothi_alkane.png}
\end{center}
Dựa vào đồ thị đã cho, số phân tử alkane không phân nhánh ở thể khí trong điều kiện thường ($25\ ^\circ\text{C}$) là
\choice
{\True 4}
{3}
{2}
{1}
\loigiai{
Ở điều kiện thường ($25\ ^\circ\text{C}$), các chất có nhiệt độ sôi thấp hơn nhiệt độ môi trường sẽ ở thể khí. Quan sát đồ thị ta thấy các alkane không phân nhánh có từ 1 đến 4 nguyên tử carbon ($\ce{CH4, C2H6, C3H8, C4H10}$) đều có nhiệt độ sôi $< 25\ ^\circ\text{C}$ (nhỏ hơn $0\ ^\circ\text{C}$), do đó có 4 alkane ở thể khí.
}
\end{ex}

\begin{ex}
[KNTT - SBT] Phát biểu nào sau đây \textbf{không đúng} về phản ứng reforming alkane?
\choice
{Chuyển alkane mạch không phân nhánh thành các alkane mạch phân nhánh}
{Chuyển alkane mạch không phân nhánh thành các hydrocarbon mạch vòng}
{Số nguyên tử carbon của chất tham gia và của sản phẩm bằng nhau}
{\True Nhiệt độ sôi của sản phẩm lớn hơn nhiều so với alkane tham gia phản ứng}
\loigiai{
Quá trình reforming làm chuyển alkane mạch không phân nhánh thành alkane mạch phân nhánh hoặc arene có cùng số C. Cấu trúc mạch phân nhánh có diện tích tiếp xúc phân tử nhỏ hơn làm giảm tương tác van der Waals, dẫn tới nhiệt độ sôi thường thấp hơn alkane mạch thẳng tương ứng, không phải lớn hơn nhiều.
}
\end{ex}

\begin{ex}
(SBT - CD). Phát biểu nào sau đây về alkane là \textbf{không đúng}?
\choice
{Trong phân tử alkane chỉ có liên kết đơn}
{\True Chỉ các alkane là chất khí ở điều kiện thường được dùng làm nhiên liệu}
{Các alkane lỏng được dùng sản xuất xăng, dầu và làm dung môi}
{Các alkane rắn được dùng làm nhựa đường, nguyên liệu cho quá trình cracking}
\loigiai{
Cả alkane thể khí (khí gas LPG), alkane thể lỏng (xăng, dầu hoả, dầu diesel) và thể rắn (parafin) đều được sử dụng làm nhiên liệu hoặc chất đốt. Do đó phát biểu ``Chỉ các alkane là chất khí... được dùng làm nhiên liệu'' là không đúng.
}
\end{ex}

\begin{ex}
[KNTT – SBT] Hợp chất nào sau đây là một alkene?
\choice
{$\ce{CH3CH2CH3}$}
{\True $\ce{CH3CH=CH2}$}
{$\ce{CH3C#CH}$}
{$\ce{CH2=C=CH2}$}
\loigiai{
Alkene là hydrocarbon không no mạch hở chứa duy nhất 1 liên kết đôi $\ce{C=C}$ trong phân tử ($\ce{C_nH_{2n}}$, $n \ge 2$). Chất $\ce{CH3CH=CH2}$ (propene) là một alkene.
}
\end{ex}

\begin{ex}
(SBT - CTST). Biểu đồ dưới đây thể hiện mối tương quan giữa nhiệt độ sôi và một số alkene:
\begin{center}
    \includegraphics[width=8.5cm]{images_crop_hoa11_de11/cau6_bieudo_alkene.png}
\end{center}
Có bao nhiêu alkene trong biểu đồ ở thể khí trong điều kiện thường ($25\ ^\circ\text{C}$)?
\choice
{4}
{2}
{3}
{\True 5}
\loigiai{
Đường màu xanh trên biểu đồ biểu thị mốc $25\ ^\circ\text{C}$. Cả 5 alkene trong biểu đồ (ethene, propene, 2-methylpropene, but-1-ene, cis-but-2-ene) đều có nhiệt độ sôi nằm dưới mức $25\ ^\circ\text{C}$, do đó ở điều kiện thường cả 5 alkene này đều tồn tại ở thể khí.
}
\end{ex}

\begin{ex}
Trùng hợp ethylene, sản phẩm thu được có công thức cấu tạo là
\choice
{$\ce{[-CH2=CH2-]_n}$}
{\True $\ce{[-CH2-CH2-]_n}$}
{$\ce{[-CH-CH-]_n}$}
{$\ce{[-CH3-CH3-]_n}$}
\loigiai{
Phản ứng trùng hợp ethylene làm đứt liên kết $\pi$ tạo poly(ethylene) (PE):
\[n\ce{CH2=CH2 ->[t^\circ, p, xt] [-CH2-CH2-]_n}\]
}
\end{ex}

\begin{ex}
(SBT - CD). Phát biểu nào sau đây là \textbf{không đúng}?
\choice
{Trong phòng thí nghiệm, người ta điều chế ethene bằng cách tách nước ethanol và thu bằng cách dời chỗ của nước}
{Một ứng dụng quan trọng của acetylene là làm nhiên liệu trong đèn xì oxygen - acetylene}
{\True Trong công nghiệp, người ta điều chế acetylene bằng cách nhiệt phân nhanh methane có xúc tác hoặc cho calcium carbide (thành phần chính của đất đèn) tác dụng với nước}
{Một ứng dụng quan trọng của acetylene là làm nguyên liệu tổng hợp ethylene}
\loigiai{
Trong công nghiệp hiện đại, acetylene được điều chế bằng phản ứng nhiệt phân nhanh methane ở khoảng $1500\ ^\circ\text{C}$ rồi làm lạnh nhanh, phản ứng không dùng xúc tác:
\[2\ce{CH4 ->[1500^\circ\text{C}, \text{làm lạnh nhanh}] C2H2 + 3H2}\]
}
\end{ex}

\begin{ex}
(SBT - CTST). Công thức phân tử nào dưới đây \textbf{không thể} là của một arene?
\choice
{$\ce{C7H8}$}
{$\ce{C10H8}$}
{\True $\ce{C11H18}$}
{$\ce{C8H8}$}
\loigiai{
Arene phải chứa ít nhất 1 vòng benzene (gồm 3 liên kết $\pi$ và 1 vòng) nên có độ bất bão hòa $k \ge 4$. Đối với $\ce{C11H18}$:
\[k = \frac{2 \times 11 + 2 - 18}{2} = 3 < 4\]
Do đó $\ce{C11H18}$ không thể là một arene.
}
\end{ex}

\begin{ex}
(SBT - CTST). Biết độ dài liên kết C=C là 134 pm, liên kết C-C là 154 pm. Thực tế, 3 liên kết $\pi$ trong vòng benzene không cố định mà trải đều trên toàn bộ vòng benzene. Giá trị nào dưới đây phù hợp với độ dài liên kết giữa carbon và carbon trong phân tử benzene?
\choice
{125 pm}
{132 pm}
{160 pm}
{\True 139 pm}
\loigiai{
Hệ electron $\pi$ liên hợp giải tỏa đều toàn bộ vòng benzene khiến cho 6 liên kết C-C hoàn toàn tương đương nhau và có tính chất trung gian giữa liên kết đơn (154 pm) và liên kết đôi (134 pm). Giá trị thực tế đo được là 139 pm.
}
\end{ex}

\begin{ex}
(SBT – KNTT). Nhận xét nào sau đây \textbf{không đúng} đối với phản ứng cộng chlorine vào benzene?
\choice
{Khó hơn phản ứng cộng chlorine vào ethylene}
{Xảy ra với điều kiện ánh sáng tử ngoại và đun nóng}
{Sản phẩm thu được là 1,2,3,4,5,6-hexachlorocyclohexane}
{\True Tỉ lệ mol của các chất tham gia phản ứng là 1 : 1}
\loigiai{
Phản ứng cộng chlorine vào benzene diễn ra theo phương trình:
\[\ce{C6H6 + 3Cl2 ->[as] C6H6Cl6}\]
Tỉ lệ mol giữa benzene và chlorine tham gia phản ứng là $1 : 3$, không phải $1 : 1$.
}
\end{ex}

\begin{ex}
(SBT - CD). Một trong những ứng dụng của toluene là
\choice
{\True làm phụ gia để tăng chỉ số octane của nhiên liệu}
{làm chất đầu để sản xuất methylcyclohexane}
{làm chất đầu để điều chế phenol}
{làm chất đầu để sản xuất polystyrene}
\loigiai{
Toluene có chỉ số octane cao (~120) nên được dùng làm chất phụ gia pha trộn vào nhiên liệu (xăng) để nâng cao khả năng chống kích nổ.
}
\end{ex}

\begin{ex}
(SBT - CD). Cho butane phản ứng với chlorine (chiếu sáng, tỉ lệ 1 : 1) thu được sản phẩm chính là
\choice
{\True 2-chlorobutane}
{1-chlorobutane}
{3-chlorobutane}
{4-chlorobutane}
\loigiai{
Butane ($\ce{CH3-CH2-CH2-CH3}$) có hai nguyên tử carbon bậc II ở vị trí 2 và 3. Theo quy tắc thế halogen, phản ứng ưu tiên thế vào nguyên tử H gắn với C bậc cao hơn, do đó 2-chlorobutane là sản phẩm chính.
}
\end{ex}

\begin{ex}
(SBT - CTST). Alkane (A) có công thức phân tử $\ce{C8H18}$. (A) tác dụng với chlorine đun nóng chỉ tạo một dẫn xuất monochloro duy nhất. Tên gọi của (A) là
\choice
{octane}
{2-methylheptane}
{2,2-dimethylhexane}
{\True 2,2,3,3-tetramethylbutane}
\loigiai{
Chất 2,2,3,3-tetramethylbutane có cấu tạo $\ce{(CH3)3C-C(CH3)3}$ đối xứng hoàn hảo, tất cả 18 nguyên tử H thuộc 6 nhóm $\ce{-CH3}$ đều tương đương nhau, do đó khi monoclo hóa chỉ tạo duy nhất 1 dẫn xuất monochloro.
}
\end{ex}

\begin{ex}
(SBT - CD). Dẫn dòng khí gồm acetylene và ethylene lần lượt đi vào ống nghiệm (1) đựng dung dịch $\ce{AgNO3/NH3}$ ở điều kiện thường, sau đó dẫn tiếp qua ống nghiệm (2) đựng nước bromine. Hiện tượng thí nghiệm nào sau đây là \textbf{không đúng}?
\choice
{Ở ống nghiệm (1) có kết tủa màu vàng nhạt}
{Ở ống nghiệm (2) màu của nước bromine nhạt dần}
{Ở ống nghiệm (2) chất lỏng chia thành hai lớp}
{\True Ở ống nghiệm (2) thu được chất lỏng đồng nhất}
\loigiai{
Khí acetylene bị giữ lại hoàn toàn ở ống (1) do phản ứng tạo kết tủa vàng nhạt $\ce{AgC#CAg}$. Khí ethylene thoát ra đi vào ống (2) phản ứng với nước bromine tạo 1,2-dibromoethane là chất lỏng không tan trong nước, nặng hơn nước nên chìm xuống dưới tách thành 2 lớp chất lỏng phân biệt. Phát biểu ``thu được chất lỏng đồng nhất'' là không đúng.
}
\end{ex}

\begin{ex}
(SBT – KNTT). Tên gọi của chất nào sau đây \textbf{không đúng}?
\choice
{\True \raisebox{-0.3\height}{\includegraphics[height=1.2cm]{images_crop_hoa11_de11/cau16_ct_A_clean.png}}\ but-2-ene}
{\raisebox{-0.3\height}{\includegraphics[height=1.2cm]{images_crop_hoa11_de11/cau16_ct_B_clean.png}}\ 3-methylbut-1-yne}
{\raisebox{-0.3\height}{\includegraphics[height=1.2cm]{images_crop_hoa11_de11/cau16_ct_C_clean.png}}\ 2,2,4-trimethylpentane}
{\raisebox{-0.3\height}{\includegraphics[height=1.2cm]{images_crop_hoa11_de11/cau16_ct_D_clean.png}}\ 1-ethyl-2-methylbenzene}
\loigiai{
Công thức ở đáp án A có mạch chính 5 carbon, chứa liên kết đôi ở vị trí C số 2 ($\ce{CH3-CH=CH-CH2-CH3}$), tên gọi thay thế đúng phải là pent-2-ene. Tên gọi ``but-2-ene'' là sai.
}
\end{ex}

\begin{ex}
(SBT – KNTT). Chất nào sau đây khi tác dụng với hỗn hợp $\ce{HNO3}$ và $\ce{H2SO4}$ đặc nóng tạo một sản phẩm mononitro hoá duy nhất?
\choice
{\True Benzene}
{Toluene}
{o-xylene}
{Naphthalene}
\loigiai{
Benzene có cấu trúc phân tử đối xứng phẳng cao, cả 6 vị trí trên vòng đều tương đương nhau nên khi nitro hóa ở tỉ lệ $1 : 1$ chỉ tạo duy nhất 1 sản phẩm mononitro hóa là nitrobenzene ($\ce{C6H5NO2}$).
}
\end{ex}

\begin{ex}
Toluene có giá trị thương mại thấp hơn so với benzene. Người ta loại bỏ nhóm methyl khỏi toluene bằng quá trình dealkyl hóa: toluene được trộn với hydrogen ở $550 - 650\ ^\circ\text{C}$, $30 - 50\ \text{atm}$ cùng xúc tác $\ce{SiO2}$ và $\ce{Al2O3}$. Phương trình hóa học của phản ứng là:
\begin{center}
    \includegraphics[width=7.5cm]{images_crop_hoa11_de11/cau18_pt_dealkyl.png}
\end{center}
\choice
{$\ce{C6H5-CH3 + 3H2 ->[xt, t^\circ] C6H11-CH3}$}
{$\ce{C6H5-CH3 + H2 ->[xt, t^\circ] C6H6 + C2H6}$}
{\True $\ce{C6H5-CH3 + H2 ->[xt, t^\circ] C6H6 + CH4}$}
{Sau phản ứng vẫn còn toluene và hydrogen dư}
\loigiai{
Phản ứng dealkyl hóa (hydrodealkylation) cắt bỏ nhóm methyl ($\ce{-CH3}$) khỏi nhân thơm tạo ra benzene và methane theo phương trình:
\[\ce{C6H5-CH3 + H2 ->[xt, t^\circ] C6H6 + CH4}\]
}
\end{ex}

\Closesolutionfile{ans}

\vspace{0.3cm}
\noindent\textbf{\large PHẦN II. Câu trắc nghiệm đúng sai.}
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{0.2cm}

\begin{ex}
Cho các alkane: Butane, hexane và decane; 2-methylheptane, octane và 2,2,4-trimethylpentane.
\choiceTF
{\True Nhiệt độ sôi của các alkane mạch không phân nhánh theo thứ tự: Butane < hexane < decane}
{Nhiệt độ sôi của các alkane mạch không phân nhánh theo thứ tự: Butane > hexane > decane}
{\True Nhiệt độ sôi của các đồng phân $\ce{C8H18}$ theo thứ tự: 2,2,4-trimethylpentane < 2-methylheptane < octane}
{Nhiệt độ sôi của các đồng phân $\ce{C8H18}$ theo thứ tự: octane < 2-methylheptane < 2,2,4-trimethylpentane}
\loigiai{
a) Đúng, b) Sai: Với các alkane mạch thẳng, khi phân tử khối tăng, diện tích tiếp xúc phân tử tăng làm tăng tương tác van der Waals giữa các phân tử, dẫn đến nhiệt độ sôi tăng dần: butane ($\ce{C4}$, $-0{,}5\ ^\circ\text{C}$) < hexane ($\ce{C6}$, $69\ ^\circ\text{C}$) < decane ($\ce{C10}$, $174\ ^\circ\text{C}$).\\
c) Đúng, d) Sai: Các chất 2,2,4-trimethylpentane, 2-methylheptane và octane đều là đồng phân có cùng CTPT $\ce{C8H18}$. Phân tử càng nhiều nhánh thì cấu trúc càng gọn cầu, diện tích tiếp xúc phân tử giảm làm tương tác van der Waals yếu hơn, nhiệt độ sôi thấp hơn: 2,2,4-trimethylpentane (3 nhánh, $99\ ^\circ\text{C}$) < 2-methylheptane (1 nhánh, $117{,}6\ ^\circ\text{C}$) < octane (không nhánh, $125{,}7\ ^\circ\text{C}$).
}
\end{ex}

\begin{ex}
Trong thí nghiệm điều chế và thử tính chất của acetylene trong phòng thí nghiệm:
\begin{center}
    \includegraphics[width=8.5cm]{images_crop_hoa11_de11/p2_cau2_tn_acetylene.png}
\end{center}
\choiceTF
{\True Phương trình hóa học của phản ứng điều chế $\ce{C2H2}$ trong thí nghiệm là: $\ce{CaC2 + 2H2O -> C2H2 ^ + Ca(OH)2}$}
{\True Dung dịch bromine ban đầu có màu vàng cam, sau phản ứng sục khí acetylene thì dung dịch bromine mất màu}
{Dung dịch bromine ban đầu không màu, sau phản ứng sục khí acetylene thì dung dịch chuyển sang màu nâu đỏ}
{\True Trong phân tử $\ce{C2H2}$, liên kết $\pi$ ở liên kết ba kém bền vững, dễ bị đứt tạo thành liên kết $\sigma$ với các nguyên tử bromine khi xảy ra phản ứng cộng}
\loigiai{
a) Đúng: Khí acetylene được điều chế bằng phản ứng thủy phân calcium carbide: $\ce{CaC2 + 2H2O -> C2H2 ^ + Ca(OH)2}$. Khí sinh ra được dẫn qua bình chứa dung dịch $\ce{NaOH}$ đặc để giữ lại các tạp chất có tính acid ($\ce{H2S, SO2}$) trước khi thử tính chất.\\
b) Đúng, c) Sai: Nước bromine có màu vàng da cam hoặc nâu đỏ, khi sục $\ce{C2H2}$ vào phản ứng cộng xảy ra: $\ce{CH#CH + 2Br2 -> CHBr2-CHBr2}$ làm dung dịch mất màu. Nước bromine không phải là chất không màu.\\
d) Đúng: Liên kết ba $\ce{-C#C-}$ gồm 1 liên kết $\sigma$ bền và 2 liên kết $\pi$ kém bền, dễ bị bẻ gãy khi tham gia phản ứng cộng để tạo thành liên kết $\sigma$ bền vững hơn.
}
\end{ex}

\begin{ex}
Terephthalic acid là nguyên liệu để tổng hợp nhựa poly(ethylene terephthalate) (PET) dùng để sản xuất tơ sợi, chai nhựa. Terephthalic acid có thể được tổng hợp từ arene X có công thức phân tử $\ce{C8H10}$ bằng cách oxi hóa X bởi dung dịch thuốc tím:
\begin{center}
    \includegraphics[width=7.5cm]{images_crop_hoa11_de11/p2_cau3_sodo_pet.png}
\end{center}
\choiceTF
{\True Chất X có độ bất bão hòa $k = 4$}
{\True Chất X có tên là 1,4-dimethylbenzene hay p-xylene}
{Chất X có tên là 1,3-dimethylbenzene hay m-xylene}
{Chất X có tên là 1,2-dimethylbenzene hay o-xylene}
\loigiai{
a) Đúng: Với arene $\ce{C8H10}$, độ bất bão hòa $k = \frac{2 \times 8 + 2 - 10}{2} = 4$ (tương ứng 1 vòng benzene gồm 3 liên kết $\pi$ và 1 vòng).\\
b) Đúng, c) Sai, d) Sai: Terephthalic acid là benzene-1,4-dicarboxylic acid (hai nhóm $\ce{-COOH}$ ở vị trí para đối diện nhau). Quá trình oxi hóa nhóm alkyl gắn trên nhân thơm bởi $\ce{KMnO4}$ giữ nguyên vị trí khung vòng, nên X phải là 1,4-dimethylbenzene (p-xylene). Oxi hóa o-xylene tạo phthalic acid (vị trí 1,2) và m-xylene tạo isophthalic acid (vị trí 1,3).
\begin{center}
    \includegraphics[width=2.5cm]{images_crop_hoa11_de11/p2_cau3_pxylene.png}
\end{center}
}
\end{ex}

\begin{ex}
Có một số loại trái cây chưa chín mà chúng ta lại muốn được sớm thưởng thức chúng, chẳng hạn một quả bơ, xoài, ...
\choiceTF
{\True Có một cách giải quyết đơn giản là cho quả bơ vào túi giấy cùng với vài quả chuối, bơ sẽ chín nhanh hơn nhiều}
{\True Chuối khi chín sẽ giải phóng một loại hormone thực vật ở thể khí là ethylene}
{Chuối khi chín sẽ giải phóng một loại hormone thực vật ở thể khí là acetylene}
{Khí ethylene chỉ giúp đẩy nhanh quá trình chín của chuối, còn nhiều loại quả khác thì không có tác dụng}
\loigiai{
a) Đúng, b) Đúng: Chuối chín sản sinh ra lượng lớn hormone thực vật tự nhiên ở thể khí là ethylene ($\ce{C2H4}$). Khi giữ trong túi giấy kín, khí ethylene tích tụ thúc đẩy các enzyme xúc tác quá trình hô hấp và phân giải tinh bột, giúp bơ và xoài nhanh chín mềm.\\
c) Sai: Hormone thực vật tự nhiên kích thích chín là ethylene, không phải acetylene.\\
d) Sai: Khí ethylene có tác động thúc đẩy chín đối với hầu hết các loại trái cây hô hấp đột biến (climacteric fruits) như bơ, chuối, xoài, cà chua, đu đủ, táo, lê,...
}
\end{ex}

\vspace{0.3cm}
\noindent\textbf{\large PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.}
\textit{Thí sinh trả lời từ câu 1 đến câu 6. Điền kết quả vào ô trả lời.}
\vspace{0.2cm}

\begin{ex}
Khí đốt hoá lỏng (Liquified Petroleum Gas, viết tắt là LPG) hay còn gọi là gas, là hỗn hợp khí chủ yếu gồm propane ($\ce{C3H8}$) và alkane Y đã được hoá lỏng. Alkane Y có chứa 4 nguyên tử carbon trong phân tử. Tỉ lệ số nguyên tử hydrogen : số nguyên tử carbon trong Y là $y$. Giá trị của $y$ bằng bao nhiêu?
\shortans{2{,}5}
\loigiai{
Alkane Y chứa 4 nguyên tử carbon nên công thức phân tử của Y là $\ce{C4H10}$. Tỉ lệ số nguyên tử H : C là $y = \frac{10}{4} = 2{,}5$.
}
\end{ex}

\begin{ex}
Alkene $\ce{C4H8}$ có bao nhiêu đồng phân cấu tạo mà khi tác dụng với dung dịch HCl chỉ cho một sản phẩm hữu cơ duy nhất?
\shortans{1}
\loigiai{
Để khi phản ứng cộng HCl chỉ cho duy nhất 1 sản phẩm hữu cơ thì phân tử alkene phải có cấu trúc đối xứng qua liên kết đôi $\ce{C=C}$. Các đồng phân cấu tạo alkene $\ce{C4H8}$ gồm:
\begin{itemize}
    \setlength\itemsep{0pt}
    \item but-1-ene: $\ce{CH2=CH-CH2-CH3}$ (bất đối xứng $\implies$ 2 sản phẩm: 2-chlorobutane và 1-chlorobutane).
    \item 2-methylpropene: $\ce{CH2=C(CH3)2}$ (bất đối xứng $\implies$ 2 sản phẩm).
    \item but-2-ene: $\ce{CH3-CH=CH-CH3}$ (đối xứng hoàn toàn $\implies$ chỉ cho 1 sản phẩm duy nhất là 2-chlorobutane).
\end{itemize}
Do đề bài chỉ hỏi số đồng phân cấu tạo, đáp án là 1.
}
\end{ex}

\begin{ex}
Cho các phát biểu sau về hydrocarbon thơm:
\begin{enumerate}[(1)]
    \setlength\itemsep{0pt}
    \item Arene còn gọi là hydrocarbon thơm là những hydrocarbon trong phân tử có chứa một hay nhiều vòng benzene.
    \item Arene còn gọi là hydrocarbon thơm là những hydrocarbon trong phân tử có chứa một hay hai vòng benzene.
    \item Hydrocarbon thơm gồm alkylbenzene và hydrocarbon thơm khác.
    \item Alkylbenzene có công thức chung là $\ce{C_nH_{2n-6}}$ ($n \ge 6$).
    \item Hydrocarbon thơm khác gồm styrene và naphthalene có công thức chung là $\ce{C_nH_{2n-6}}$ ($n \ge 6$).
\end{enumerate}
Số phát biểu đúng là bao nhiêu?
\shortans{3}
\loigiai{
Phân tích các phát biểu theo tài liệu gốc:
\begin{itemize}
    \setlength\itemsep{0pt}
    \item (2) đúng (theo phạm vi SGK xét 1 hoặc 2 vòng benzene; định nghĩa tổng quát (1) cũng chuẩn).
    \item (3) đúng: Arene được phân thành alkylbenzene và các hydrocarbon thơm khác (như styrene, naphthalene).
    \item (4) đúng: Dãy đồng đẳng alkylbenzene có CTTQ là $\ce{C_nH_{2n-6}}$ ($n \ge 6$).
    \item (5) sai vì styrene là $\ce{C8H8}$ ($\ce{C_nH_{2n-8}}$), naphthalene là $\ce{C10H8}$ ($\ce{C_nH_{2n-12}}$), không tuân theo $\ce{C_nH_{2n-6}}$.
\end{itemize}
Số phát biểu đúng là 3.
}
\end{ex}

\begin{ex}
Ethylene là một hormone thực vật có vai trò kích thích quá trình tạo mủ và làm cho mủ cao su chậm đông, chảy dai hơn. Trong nông nghiệp hiện đại, người ta áp khí ethylene trực tiếp vào vỏ thân cây để kích thích tiết mủ:
\begin{center}
    \includegraphics[width=7.5cm]{images_crop_hoa11_de11/p3_cau4_cao_su.png}
\end{center}
Trong công nghiệp, ethylene được điều chế từ phản ứng dehydro hóa ethane ở $500\ ^\circ\text{C}$ có xúc tác: $\ce{C2H6 ->[xt, 500^\circ\text{C}] C2H4 + H2}$. Để bơm vào túi khí 60 cc (đkc: $25\ ^\circ\text{C}$, 1 bar) khí ethylene thì cần dùng bao nhiêu miligam ethane? (Làm tròn đến hàng đơn vị).
\shortans{73}
\loigiai{
- Thể tích ethylene cần dùng: $V = 60\text{ cc} = 60\text{ mL} = 0{,}06\text{ L}$.\\
- Số mol $\ce{C2H4}$ (ở đkc: 24,79 L/mol): $n = \frac{0{,}06}{24{,}79}\text{ mol}$.\\
- Phương trình phản ứng: $\ce{C2H6 ->[xt, t^\circ] C2H4 + H2}$.\\
- Số mol ethane cần dùng: $n(\ce{C2H6}) = n(\ce{C2H4}) = \frac{0{,}06}{24{,}79}\text{ mol}$.\\
- Khối lượng ethane: $m = \frac{0{,}06}{24{,}79} \times 30 \approx 0{,}07261\text{ g} = 72{,}61\text{ mg} \approx 73\text{ mg}$.
}
\end{ex}

\begin{ex}
Đất đèn có thành phần chính là calcium carbide ($\ce{CaC2}$). Cho 20 gam một mẫu đất đèn tác dụng với nước dư thu được 7,437 lít khí acetylene (ở đkc). Cho rằng các tạp chất trong mẫu trơ với nước và phản ứng xảy ra hoàn toàn. Độ tinh khiết của mẫu $\ce{CaC2}$ trên là bao nhiêu phần trăm?
\shortans{96\%}
\loigiai{
- Số mol khí acetylene: $n(\ce{C2H2}) = \frac{7{,}437}{24{,}79} = 0{,}3\text{ mol}$.\\
- Phương trình phản ứng: $\ce{CaC2 + 2H2O -> Ca(OH)2 + C2H2 ^}$\\
- Theo phương trình: $n(\ce{CaC2}) = n(\ce{C2H2}) = 0{,}3\text{ mol}$.\\
- Khối lượng $\ce{CaC2}$ nguyên chất: $m = 0{,}3 \times 64 = 19{,}2\text{ gam}$.\\
- Độ tinh khiết của mẫu: $\% \ce{CaC2} = \frac{19{,}2}{20} \times 100\% = 96\%$.
}
\end{ex}

\begin{ex}
Butane là một trong hai thành phần chính của khí đốt hóa lỏng (LPG). Khi đốt cháy hoàn toàn 1 mol butane tỏa ra lượng nhiệt là 2497 kJ. Để đun nóng 1 gam nước tăng thêm $1\ ^\circ\text{C}$ cần cung cấp nhiệt lượng là 4,18 J. Tính khối lượng butane (theo gam, làm tròn đến một chữ số thập phân) cần đốt để đun sôi 2 lít nước từ $25\ ^\circ\text{C}$ lên $100\ ^\circ\text{C}$. Biết khối lượng riêng của nước là $1\text{ g/mL}$ và hiệu suất hấp thụ nhiệt của nước là 60\%.
\shortans{24{,}3}
\loigiai{
- Khối lượng 2 lít nước: $m(\ce{H2O}) = 2000\text{ mL} \times 1\text{ g/mL} = 2000\text{ gam}$.\\
- Nhiệt lượng cần thiết để đun nóng nước từ $25\ ^\circ\text{C}$ lên $100\ ^\circ\text{C}$:\\
  \[Q(\text{thu}) = m \cdot c \cdot \Delta t = 2000 \times 4{,}18 \times (100 - 25) = 627\,000\text{ J} = 627\text{ kJ}.\]
- Vì hiệu suất hấp thụ nhiệt là 60\%, nhiệt lượng thực tế cần đốt cháy tỏa ra là:\\
  \[Q(\text{tỏa}) = \frac{Q(\text{thu})}{60\%} = \frac{627}{0{,}60} = 1045\text{ kJ}.\]
- Số mol butane cần đốt: $n(\ce{C4H10}) = \frac{1045}{2497} \approx 0{,}4185\text{ mol}$.\\
- Khối lượng butane cần dùng: $m = 0{,}4185 \times 58 \approx 24{,}27\text{ g} \approx 24{,}3\text{ gam}$.
}
\end{ex}

\vspace{0.3cm}
\noindent\textbf{\large PHẦN IV. Tự luận.}
\textit{Thí sinh trình bày chi tiết lời giải vào bài làm.}
\vspace{0.2cm}

\begin{ex}
(SBT - CTST). Viết công thức khung phân tử của các chất sau:
\begin{enumerate}[a)]
    \setlength\itemsep{0pt}
    \item propene.
    \item pent-1-ene.
    \item 3-methylpent-1-yne.
    \item cis-pent-2-ene.
    \item trans-pent-2-ene.
\end{enumerate}
"""
    if is_teacher:
        tex += r"""
\loigiai{
Công thức khung phân tử tương ứng của các hợp chất:
\begin{center}
    \includegraphics[width=13.0cm]{images_crop_hoa11_de11/p4_cau1_khung_pt.png}
\end{center}
\begin{itemize}
    \setlength\itemsep{0pt}
    \item \textbf{a) propene:} Khung 3 nguyên tử carbon với liên kết đôi ở đầu mạch.
    \item \textbf{b) pent-1-ene:} Mạch hở 5 nguyên tử carbon dạng gấp khúc, liên kết đôi tại C số 1.
    \item \textbf{c) 3-methylpent-1-yne:} Mạch chính 5 carbon chứa liên kết ba ở C số 1, nhánh methyl ở vị trí C số 3.
    \item \textbf{d) cis-pent-2-ene:} Đồng phân hình học có 2 mạch nhánh alkyl ($\ce{-CH3}$ và $\ce{-C2H5}$) cùng nằm về một phía của mặt phẳng liên kết đôi $\ce{C=C}$.
    \item \textbf{e) trans-pent-2-ene:} Đồng phân hình học có 2 mạch nhánh alkyl nằm về hai phía đối diện của mặt phẳng liên kết đôi $\ce{C=C}$.
\end{itemize}
}
"""
    else:
        tex += r"""
\vspace{5.0cm}
"""
    tex += r"""
\end{ex}

\begin{ex}
Trinitrotoluene (TNT) là một hợp chất hữu cơ có công thức $\ce{C6H2(NO2)3CH3}$, được tổng hợp bằng phản ứng trinitro hóa toluene.
\begin{enumerate}[a)]
    \setlength\itemsep{0pt}
    \item Viết phương trình hóa học điều chế 2,4,6-trinitrotoluene từ toluene và nitric acid đặc (xúc tác sulfuric acid đặc, đun nóng).
    \item Tính khối lượng toluene (theo tấn, làm tròn đến một chữ số thập phân) cần dùng để điều chế 36 000 tấn trinitrotoluene, biết hiệu suất phản ứng của toàn bộ quá trình đạt 70\%.
\end{enumerate}
"""
    if is_teacher:
        tex += r"""
\loigiai{
\textbf{a)} Phương trình hóa học điều chế TNT:
\[\ce{C6H5-CH3 + 3HNO3 ->[H2SO4\text{ đặc}, t^\circ] C6H2(NO2)3-CH3 + 3H2O}\]
\textbf{b)} Phân tử khối:
\begin{itemize}
    \setlength\itemsep{0pt}
    \item $M(\text{toluene}, \ce{C7H8}) = 92\text{ g/mol}$.
    \item $M(\text{TNT}, \ce{C7H5N3O6}) = 12 \times 7 + 5 + 14 \times 3 + 16 \times 6 = 227\text{ g/mol}$.
\end{itemize}
- Theo phương trình hóa học: Cứ 92 tấn toluene phản ứng 100\% sẽ tạo thành 227 tấn TNT.\\
- Khối lượng toluene lý thuyết cần dùng để tạo 36 000 tấn TNT:
\[m(\text{lt}) = \frac{36\,000 \times 92}{227} \approx 14\,590{,}31\text{ tấn}.\]
- Do hiệu suất của cả quá trình chỉ đạt 70\%, lượng toluene thực tế cần lấy là:
\[m(\text{tt}) = \frac{m(\text{lt})}{H} = \frac{14\,590{,}31}{70\%} = \frac{14\,590{,}31}{0{,}70} \approx 20\,843{,}3\text{ tấn}.\]
}
"""
    else:
        tex += r"""
\vspace{5.0cm}
"""
    tex += r"""
\end{ex}
"""

    if is_teacher:
        tex += r"""
\bigskip
\newpage
\noindent\textbf{\large BẢNG ĐÁP ÁN THAM KHẢO PHẦN I}
\vspace{0.2cm}

\noindent
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1. A} & \textbf{2. A} & \textbf{3. D} & \textbf{4. B} & \textbf{5. B} & \textbf{6. D} & \textbf{7. B} & \textbf{8. C} & \textbf{9. C} & \textbf{10. D} \\
\hline
\textbf{11. D} & \textbf{12. A} & \textbf{13. A} & \textbf{14. D} & \textbf{15. D} & \textbf{16. A} & \textbf{17. A} & \textbf{18. C} & & \\
\hline
\end{tabular}
\end{center}

\bigskip
\noindent\textbf{\large BẢNG ĐÁP ÁN THAM KHẢO PHẦN II}
\vspace{0.2cm}

\noindent
\begin{center}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Lệnh hỏi a} & \textbf{Lệnh hỏi b} & \textbf{Lệnh hỏi c} & \textbf{Lệnh hỏi d} \\
\hline
\textbf{Câu 1} & Đ & S & Đ & S \\
\hline
\textbf{Câu 2} & Đ & Đ & S & Đ \\
\hline
\textbf{Câu 3} & Đ & Đ & S & S \\
\hline
\textbf{Câu 4} & Đ & Đ & S & S \\
\hline
\end{tabular}
\end{center}

\bigskip
\noindent\textbf{\large BẢNG ĐÁP ÁN THAM KHẢO PHẦN III}
\vspace{0.2cm}

\noindent
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\
\hline
2{,}5 & 1 & 3 & 73 & 96\% & 24{,}3 \\
\hline
\end{tabular}
\end{center}
"""

    tex += r"""
\begin{center}
\textbf{--------- HẾT ---------}
\end{center}

\end{document}
"""
    return tex

def main():
    hs_tex = make_tex(is_teacher=False)
    hs_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De11_HocSinh.tex")
    with open(hs_path, "w", encoding="utf-8") as f:
        f.write(hs_tex)
    print("Saved:", hs_path)

    gv_tex = make_tex(is_teacher=True)
    gv_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De11_GiaoVien.tex")
    with open(gv_path, "w", encoding="utf-8") as f:
        f.write(gv_tex)
    print("Saved:", gv_path)

    # Compile with pdflatex
    for name in ["De_Kiem_Tra_Hoa_11_De11_HocSinh", "De_Kiem_Tra_Hoa_11_De11_GiaoVien"]:
        print(f"Compiling {name}.tex...")
        cmd = f'pdflatex -interaction=nonstopmode "{name}.tex"'
        proc = subprocess.run(cmd, cwd=TEX_DIR, shell=True, capture_output=True, text=True)
        pdf_file = os.path.join(TEX_DIR, f"{name}.pdf")
        if os.path.exists(pdf_file):
            print(f"Successfully created: {pdf_file} ({os.path.getsize(pdf_file) / 1024:.1f} KB)")
        else:
            print(f"Compilation failed for {name}!")
            print(proc.stdout[-500:])

    # Clean up auxiliary files
    print("Cleaning auxiliary files...")
    for ext in [".aux", ".log", ".thm", ".out"]:
        for root, dirs, files in os.walk(TEX_DIR):
            for file in files:
                if file.endswith(ext) or file.startswith("ans-hoa11-de11"):
                    try:
                        os.remove(os.path.join(root, file))
                    except:
                        pass
    print("LaTeX generation completed successfully!")

if __name__ == "__main__":
    main()
