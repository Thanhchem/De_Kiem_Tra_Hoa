# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file LaTeX và biên dịch ra PDF cho ĐỀ SỐ 5 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De05_HocSinh.tex / .pdf
- De_Kiem_Tra_Hoa_11_De05_GiaoVien.tex / .pdf
Chương 2: NITROGEN – SULFUR
Tuân thủ Form chuẩn GEMINI.md, sử dụng gói ex_test.sty và mhchem.
"""

import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import subprocess
import shutil

TEX_DIR = r"c:\Antigravity_Thanh\San_Pham"

if not os.path.exists(os.path.join(TEX_DIR, "ex_test.sty")):
    shutil.copyfile(r"c:\Antigravity_Thanh\ex_test.sty", os.path.join(TEX_DIR, "ex_test.sty"))

def make_tex(is_teacher=False):
    ver_tag = "gv" if is_teacher else "hs"
    ans_file = f"ans-hoa11-de05{ver_tag}"
    
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
	\centering\small
	\textbf{ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 2 (NITROGEN -- SULFUR)}\\
	\textbf{MÔN: HOÁ HỌC -- LỚP 11}\\
	\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
	\fbox{\textbf{MÃ ĐỀ: 105}}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; Ar = 40; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Ba = 137).}
\smallskip

\noindent\textbf{\large PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.}
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{0.2cm}

\Opensolutionfile{ans}[""" + ans_file + r"""]
"""

    # Phần I - 18 câu
    # Câu 1
    tex += r"""
\begin{ex}
Nguyên tố dinh dưỡng nào sau đây có trong cấu tạo của protein?
\choice
{Sắt}
{Kẽm}
{\True Nitrogen}
{Đồng}
""" + (r"""
\loigiai{
Protein là hợp chất polypeptide cao phân tử được cấu tạo từ các amino acid chứa các nguyên tố C, H, O, N và một số nguyên tố khác. Do đó, nitrogen là nguyên tố đặc trưng luôn có trong cấu tạo của protein.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 2
    tex += r"""
\begin{ex}
Để tạo độ xốp cho một số loại bánh, có thể dùng muối nào sau đây?
\choice
{$\ce{CaCO3}$}
{\True $\ce{NH4HCO3}$}
{$\ce{NH4Cl}$}
{$\ce{NaCl}$}
""" + (r"""
\loigiai{
$\ce{NH4HCO3}$ (bột nở) dễ bị nhiệt phân hủy khi nướng bánh theo phản ứng:\\
$\ce{NH4HCO3 ->[t^\circ] NH3 ^ + CO2 ^ + H2O}$.\\
Các khí sinh ra thoát ra ngoài tạo thành các lỗ xốp li ti làm bánh nở đều và xốp mềm.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 3
    tex += r"""
\begin{ex}
Mưa acid ảnh hưởng tới hệ thực vật, phá hủy các vật liệu bằng kim loại, các bức tượng bằng đá, gây bệnh cho con người và động vật. Hiện tượng trên gây ra chủ yếu do khí thải của nhà máy nhiệt điện, phương tiện giao thông và sản xuất công nghiệp. Tác nhân chủ yếu trong khí thải gây ra mưa acid là
\begin{center}
	\includegraphics[width=0.62\textwidth]{images_crop_hoa11_de05/cau3_so_do_mua_acid.png}
\end{center}
\choice
{\True $\ce{SO2}$ và $\ce{NO2}$}
{$\ce{CH4}$ và $\ce{NH3}$}
{$\ce{CO}$ và $\ce{CH4}$}
{$\ce{CO}$ và $\ce{CO2}$}
""" + (r"""
\loigiai{
Hai khí thải chủ yếu trong công nghiệp và giao thông gây mưa acid là $\ce{SO2}$ và các oxide của nitrogen ($\ce{NO_x}$, chủ yếu là $\ce{NO2}$). Chúng bị oxi hóa trong khí quyển kết hợp với nước mưa tạo ra sulfuric acid ($\ce{H2SO4}$) và nitric acid ($\ce{HNO3}$).
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 4
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Oxide của nitrogen được tạo thành ở nhiệt độ cao, khi nitrogen có trong không khí bị oxi hóa được gọi là
\choice
{$\ce{NO_x}$ tức thời}
{\True $\ce{NO_x}$ nhiệt}
{$\ce{NO_x}$ nhiên liệu}
{$\ce{NO_x}$ tự nhiên}
""" + (r"""
\loigiai{
Ở nhiệt độ rất cao (trên $1200^\circ\text{C}$), khí $\ce{N2}$ trong không khí phản ứng trực tiếp với $\ce{O2}$ tạo ra các oxide của nitrogen. Nguồn $\ce{NO_x}$ sinh ra theo cơ chế này được gọi là $\ce{NO_x}$ nhiệt (Thermal $\ce{NO_x}$).
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 5
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Công thức hoá học của diêm tiêu Chile là
\choice
{$\ce{Ca(NO3)2}$}
{$\ce{NH4NO3}$}
{$\ce{NH4Cl}$}
{\True $\ce{NaNO3}$}
""" + (r"""
\loigiai{
Diêm tiêu Chile (Chile saltpeter) là tên khoáng vật tự nhiên của sodium nitrate có công thức hóa học là $\ce{NaNO3}$.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 6
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Sulfur đóng vai trò chất khử khi tác dụng với đơn chất nào sau đây?
\choice
{$\ce{Fe}$}
{\True $\ce{O2}$}
{$\ce{H2}$}
{$\ce{Hg}$}
""" + (r"""
\loigiai{
Khi tác dụng với $\ce{O2}$: $\ce{S^0 + O2 ->[t^\circ] S^{+4}O2}$. Số oxi hóa của S tăng từ 0 lên +4 nên sulfur đóng vai trò là chất khử. Với các kim loại ($\ce{Fe, Hg}$) hoặc $\ce{H2}$, số oxi hóa của S giảm từ 0 xuống -2 nên S đóng vai trò là chất oxi hóa.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 7
    tex += r"""
\begin{ex}
\textbf{(SBT - CTST)} Cách pha loãng dung dịch $\ce{H2SO4}$ đặc nào sau đây đúng?
\choice
{Rót nhanh acid vào nước và khuấy đều}
{Rót nhanh nước vào acid và khuấy đều}
{Rót từ từ nước vào acid và khuấy đều}
{\True Rót từ từ acid vào nước và khuấy đều}
""" + (r"""
\loigiai{
Quá trình hòa tan $\ce{H2SO4}$ đặc vào nước tỏa ra lượng nhiệt rất lớn. Để đảm bảo an toàn, cần rót từ từ acid vào nước dọc theo đũa thủy tinh và khuấy nhẹ đều. Tuyệt đối không được rót nước vào acid đặc vì nước sôi đột ngột sẽ bắn các giọt acid nóng ra ngoài gây bỏng nguy hiểm.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 8
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Trong nghiên cứu, khí nitrogen thường được dùng để tạo bầu khí quyển trơ dựa trên cơ sở nào?
\choice
{Nitrogen có tính oxi hoá mạnh}
{\True Nitrogen rất bền với nhiệt}
{Nitrogen khó hoá lỏng}
{Nitrogen không có cực}
""" + (r"""
\loigiai{
Phân tử $\ce{N2}$ có liên kết ba $\ce{N\bond{3}N}$ rất bền vững với năng lượng liên kết khổng lồ ($945\text{ kJ/mol}$). Ở nhiệt độ thường và nhiệt độ trung bình, nitrogen rất bền nhiệt và rất trơ về mặt hóa học, do đó được dùng làm khí quyển trơ bảo quản hóa chất và vật liệu.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 9
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Xét cân bằng hóa học trong dung dịch nước:\\
$\ce{NH3 + H2O <=> NH4+ + OH-}$\\
Hằng số cân bằng ($K_\text{C}$) của phản ứng được biểu diễn bằng biểu thức nào sau đây?
\choice
{\True $K_\text{C} = \dfrac{[\ce{NH4+}][\ce{OH-}]}{[\ce{NH3}]}$}
{$K_\text{C} = \dfrac{[\ce{NH4+}][\ce{OH-}]}{[\ce{NH3}][\ce{H2O}]}$}
{$K_\text{C} = \dfrac{[\ce{NH4+}][\ce{OH-}]}{[\ce{H2O}]}$}
{$K_\text{C} = \dfrac{[\ce{NH4+}]}{[\ce{NH3}]}$}
""" + (r"""
\loigiai{
Trong dung dịch nước loãng, nồng độ của dung môi $\ce{H2O}$ rất lớn và coi như không đổi nên được gộp vào hằng số cân bằng. Biểu thức chuẩn của hằng số cân bằng là $K_\text{C} = \dfrac{[\ce{NH4+}][\ce{OH-}]}{[\ce{NH3}]}$.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 10
    tex += r"""
\begin{ex}
Trong các phản ứng dưới đây, phản ứng nào $\ce{NH3}$ \textbf{không} thể hiện tính khử?
\choice
{$\ce{4NH3 + 5O2 ->[xt, t^\circ] 4NO + 6H2O}$}
{\True $\ce{NH3 + HCl -> NH4Cl}$}
{$\ce{8NH3 + 3Cl2 -> 6NH4Cl + N2}$}
{$\ce{2NH3 + 3CuO ->[t^\circ] 3Cu + 3H2O + N2}$}
""" + (r"""
\loigiai{
Trong phản ứng $\ce{NH3 + HCl -> NH4Cl}$, số oxi hóa của nitrogen giữ nguyên là -3 (phản ứng thể hiện tính base do nhận proton $\ce{H+}$). Ở 3 phản ứng còn lại, số oxi hóa của N tăng từ -3 lên +2 hoặc 0 nên $\ce{NH3}$ thể hiện tính khử.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 11
    tex += r"""
\begin{ex}
Thí nghiệm với dung dịch $\ce{HNO3}$ thường sinh ra khí độc $\ce{NO2}$. Để hạn chế khí $\ce{NO2}$ thoát ra từ ống nghiệm, người ta nút ống nghiệm bằng:
\begin{enumerate}[(a)]
	\item bông khô.
	\item bông có tẩm nước.
	\item bông có tẩm nước vôi.
	\item bông có tẩm giấm ăn.
\end{enumerate}
Trong 4 biện pháp trên, biện pháp có hiệu quả nhất là
\choice
{(d)}
{\True (c)}
{(a)}
{(b)}
""" + (r"""
\loigiai{
Khí $\ce{NO2}$ là acidic oxide độc hại. Khi nút ống nghiệm bằng bông tẩm dung dịch nước vôi $\ce{Ca(OH)2}$, $\ce{NO2}$ phản ứng tạo muối tan bị giữ lại hoàn toàn: $\ce{2NO2 + Ca(OH)2 -> Ca(NO2)2 + Ca(NO3)2 + H2O}$.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 12
    tex += r"""
\begin{ex}
Ứng dụng nào sau đây \textbf{không} phải của $\ce{SO2}$?
\choice
{Sản xuất sulfuric acid}
{Tẩy màu dung dịch đường}
{\True Khử mùi không khí}
{Tẩy trắng giấy, bột giấy}
""" + (r"""
\loigiai{
Khí $\ce{SO2}$ có mùi hắc cay, nồng, gây ho và kích ứng hô hấp, có độc tính đối với sức khỏe nên tuyệt đối không được dùng để khử mùi không khí.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 13
    tex += r"""
\begin{ex}
Trong điều kiện thích hợp, xảy ra các phản ứng sau:
\begin{enumerate}[(a)]
	\item $\ce{2H2SO4 + C -> 2SO2 + CO2 + 2H2O}$
	\item $\ce{H2SO4 + Fe(OH)2 -> FeSO4 + 2H2O}$
	\item $\ce{4H2SO4 + 2FeO -> Fe2(SO4)3 + SO2 + 4H2O}$
	\item $\ce{6H2SO4 + 2Fe -> Fe2(SO4)3 + 3SO2 + 6H2O}$
\end{enumerate}
Trong các phản ứng trên, phản ứng xảy ra với dung dịch $\ce{H2SO4}$ loãng là
\choice
{(a)}
{(c)}
{\True (b)}
{(d)}
""" + (r"""
\loigiai{
Phản ứng (b) là phản ứng trao đổi acid-base xảy ra giữa dung dịch $\ce{H2SO4}$ loãng và base $\ce{Fe(OH)2}$. Các phản ứng (a), (c), (d) là phản ứng oxi hóa - khử giải phóng khí $\ce{SO2}$, chỉ xảy ra khi dùng $\ce{H2SO4}$ đặc nóng.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 14
    tex += r"""
\begin{ex}
Dẫn khí $\ce{SO2}$ vào dung dịch $\ce{KMnO4}$ hiện tượng quan sát được là
\choice
{Dung dịch không màu chuyển sang màu tím}
{Dung dịch màu tím bị vẩn đục màu vàng}
{Màu tím của dung dịch $\ce{KMnO4}$ chuyển sang màu vàng}
{\True Màu tím của dung dịch $\ce{KMnO4}$ chuyển sang không màu}
""" + (r"""
\loigiai{
Do $\ce{SO2}$ có tính khử mạnh, khử ion $\ce{MnO4-}$ (màu tím) xuống $\ce{Mn^{2+}}$ (không màu):\\
$\ce{5SO2 + 2KMnO4 + 2H2O -> K2SO4 + 2MnSO4 + 2H2SO4}$.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 15
    tex += r"""
\begin{ex}
Nguyên nhân chủ yếu gây ra hiện tượng phú dưỡng trong các thủy vực là gì?
\choice
{Sự tích tụ các chất hữu cơ do phân hủy xác sinh vật}
{\True Sự gia tăng quá mức các chất dinh dưỡng như nitrate ($\ce{NO3-}$) và phosphate ($\ce{PO4^3-}$) từ nước thải sinh hoạt, phân bón nông nghiệp và công nghiệp}
{Sự thay đổi nhiệt độ nước trong thủy vực}
{Quá trình xói mòn đất làm thay đổi hệ sinh thái nước ngọt}
""" + (r"""
\loigiai{
Hiện tượng phú dưỡng (eutrophication) xảy ra khi nguồn nước tích tụ nồng độ quá cao các nguyên tố dinh dưỡng nitrogen ($\ce{NO3-}$, $\ce{NH4+}$) và phosphorus ($\ce{PO4^3-}$) từ phân bón nông nghiệp dư thừa và nước thải, làm bùng phát các loài tảo và vi sinh vật.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 16
    tex += r"""
\begin{ex}
Mưa acid có thể gây ra tác động tiêu cực nào đối với các công trình xây dựng?
\choice
{Làm tăng độ bền của các công trình xây dựng bằng đá vôi}
{\True Gây ăn mòn các công trình làm từ đá vôi, đá cẩm thạch và kim loại}
{Tăng khả năng chống chịu của các vật liệu xây dựng trước sự tác động của môi trường}
{Không có tác động đến các công trình xây dựng}
""" + (r"""
\loigiai{
Các acid $\ce{H2SO4}$ và $\ce{HNO3}$ trong nước mưa hòa tan và ăn mòn đá vôi ($\ce{CaCO3}$), đá cẩm thạch và ăn mòn kim loại trong kết cấu công trình, tượng đài.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 17
    tex += r"""
\begin{ex}
Biện pháp nào dưới đây có thể giúp giảm thiểu hiện tượng mưa acid?
\choice
{Tăng cường việc sử dụng các loại nhiên liệu hóa thạch}
{\True Giảm lượng khí thải $\ce{SO2}$ và $\ce{NO_x}$ từ các nhà máy công nghiệp và phương tiện giao thông}
{Sử dụng nhiều hơn các loại phân bón chứa phosphate}
{Khuyến khích việc phát triển giao thông đường bộ}
""" + (r"""
\loigiai{
Cắt giảm triệt để lượng phát thải các khí tiền thân tạo acid là $\ce{SO2}$ và $\ce{NO_x}$ bằng các biện pháp xử lý khí thải công nghiệp và sử dụng năng lượng sạch là giải pháp then chốt để hạn chế mưa acid.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 18
    tex += r"""
\begin{ex}
Cho các phát biểu sau:
\begin{enumerate}[(1)]
	\item Cho thanh Fe (iron) vào dung dịch $\ce{H2SO4}$ loãng, iron tan ra, sủi bọt khí.
	\item Cho dung dịch $\ce{BaCl2}$ vào dung dịch $\ce{H2SO4}$, xuất hiện kết tủa đen.
	\item Ở nhiệt độ cao, sulfur tác dụng với hydrogen tạo thành hydrogen sulfide ($\ce{H2S}$).
	\item Cho từng giọt sulfuric acid đặc vào đường saccharose ($\ce{C12H22O11}$) màu trắng, đường chuyển màu nâu $\rightarrow$ đen sau đó trào lên khỏi miệng cốc.
	\item Nitrogen lỏng được dùng để làm lạnh nhanh, bảo quản mẫu vật phẩm trong y học.
	\item Nitrogen bền nhiệt, có tính trơ nên không bảo quản được hương vị thực phẩm.
	\item Không khí có chứa oxygen là tác nhân oxi hóa thực phẩm, làm thực phẩm ôi thiu chậm hơn trong khí nitrogen.
\end{enumerate}
Số phát biểu đúng là
\choice
{3}
{\True 4}
{5}
{6}
""" + (r"""
\loigiai{
Có 4 phát biểu đúng gồm: (1), (3), (4), (5).\\
- (1) Đúng: $\ce{Fe + H2SO4 -> FeSO4 + H2 ^}$.\\
- (2) Sai: Xuất hiện kết tủa trắng $\ce{BaSO4}$, không phải kết tủa đen.\\
- (3) Đúng: $\ce{H2 + S ->[t^\circ] H2S}$.\\
- (4) Đúng: $\ce{H2SO4}$ đặc háo nước than hóa đường thành than (đen), sau đó oxi hóa C thành $\ce{CO2, SO2}$ đẩy trào cột than lên.\\
- (5) Đúng: Nitrogen lỏng sôi ở $-196^\circ\text{C}$ dùng bảo quản mẫu sinh học.\\
- (6) Sai: Nhờ tính trơ hóa học nên nitrogen bảo quản rất tốt hương vị và độ giòn của thực phẩm.\\
- (7) Sai: Oxygen là tác nhân oxi hóa làm thực phẩm nhanh ôi thiu hơn so với trong khí nitrogen.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    tex += r"""
\Closesolutionfile{ans}

\vspace{0.3cm}
\noindent\textbf{\large PHẦN II. Câu trắc nghiệm đúng sai.}
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{0.2cm}
"""

    # Phần II - Câu 1
    tex += r"""
\begin{ex}
Để giảm sốt hoặc giảm đau, người ta có thể dùng túi chườm lạnh chứa hóa chất. Hãy tìm hiểu về loại túi chườm lạnh này:
\begin{center}
	\includegraphics[width=0.45\textwidth]{images_crop_hoa11_de05/cau1_p2_tui_chuom_lanh.png}
\end{center}
\choiceTF
{\True Hóa chất trong túi chườm lạnh gồm ammonium chloride rắn và nước cất}
{\True Túi chườm sẽ hoạt động khi phá vỡ lớp ngăn cách giữa muối ammonium chloride và nước cất}
{Do sự hòa tan $\ce{NH4Cl(s)}$ vào nước là một quá trình tỏa nhiệt: $\ce{NH4Cl(s) -> NH4Cl(aq)}\ \Delta_r H^\circ_{298} < 0$}
{Phản ứng tỏa nhiệt, muối hấp thụ nhiệt từ nước, làm nước giảm nhiệt độ}
""" + (r"""
\loigiai{
\begin{itemize}
	\item a) Đúng: Thành phần túi chườm lạnh hóa học thường chứa muối ammonium chloride rắn và một túi chứa nước riêng biệt.
	\item b) Đúng: Khi bóp mạnh làm vỡ túi nước ngăn cách, muối hòa tan vào nước tạo hiệu ứng làm lạnh nhanh chóng.
	\item c) Sai: Quá trình hòa tan $\ce{NH4Cl}$ vào nước là một quá trình thu nhiệt rất mạnh ($\Delta_r H^\circ_{298} > 0$), làm nhiệt độ dung dịch giảm sâu.
	\item d) Sai: Đây là quá trình thu nhiệt; túi chườm lạnh hấp thụ nhiệt từ vùng cơ thể tiếp xúc giúp hạ nhiệt độ cơ thể và giảm đau sưng.
\end{itemize}
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Phần II - Câu 2
    tex += r"""
\begin{ex}
Biện pháp nào sau đây dùng để hạn chế hiện tượng phú dưỡng?
\choiceTF
{\True Tạo điều kiện để nước trong kênh rạch, ao, hồ được lưu thông}
{Xả nước thải (nông nghiệp, công nghiệp và sinh hoạt) trực tiếp đến ao, hồ thông qua các cống dẫn nước cố định}
{\True Sử dụng phân bón đúng liều lượng, đúng cách, đúng thời điểm trong năm để hạn chế sự rửa trôi của ion nitrate và phosphate từ nguồn phân bón dư thừa vào kênh rạch, ao, hồ}
{\True Xử lí nước thải trước khi cho vào kênh rạch, ao, hồ}
""" + (r"""
\loigiai{
\begin{itemize}
	\item a) Đúng: Khơi thông dòng chảy giúp tăng cường nồng độ oxygen hòa tan và hạn chế tích tụ chất dinh dưỡng tù đọng.
	\item b) Sai: Xả trực tiếp nước thải chưa qua xử lý vào ao hồ là nguyên nhân chính gây phú dưỡng nghiêm trọng.
	\item c) Đúng: Bón phân hợp lý và đúng kỹ thuật hạn chế tối đa lượng nitrate và phosphate dư thừa bị mưa rửa trôi xuống nguồn nước.
	\item d) Đúng: Trạm xử lý nước thải sinh hoạt và công nghiệp giúp loại bỏ triệt để các hợp chất chứa N và P trước khi xả ra môi trường.
\end{itemize}
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Phần II - Câu 3
    tex += r"""
\begin{ex}
Với lứa tuổi học sinh có thể giúp giảm thiểu lượng sulfur dioxide thải vào không khí bằng các việc làm phù hợp:
\choiceTF
{\True Khuyến khích sử dụng xe điện, sử dụng nguyên liệu thân thiện môi trường như xăng E5,...}
{Khuyến khích đốt rác thải}
{\True Tuyên truyền tác hại của $\ce{SO2}$ và lợi ích của nhiên liệu tái tạo, bảo vệ môi trường,...}
{\True Giảm rác thải, phân loại rác, tái chế, tái sử dụng}
""" + (r"""
\loigiai{
\begin{itemize}
	\item a) Đúng: Sử dụng xe đạp, xe điện hoặc nhiên liệu sinh học giúp giảm phát thải các khí độc hại từ nhiên liệu hóa thạch chứa lưu huỳnh.
	\item b) Sai: Đốt rác thải tự phát làm phát sinh lượng lớn khói bụi độc hại và khí $\ce{SO2}$ gây ô nhiễm không khí.
	\item c) Đúng: Tuyên truyền học đường nâng cao nhận thức bảo vệ bầu không khí trong lành.
	\item d) Đúng: Thực hiện mô hình 3R (Reduce - Reuse - Recycle) giúp tiết kiệm năng lượng và giảm thiểu ô nhiễm môi trường.
\end{itemize}
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Phần II - Câu 4
    tex += r"""
\begin{ex}
Sulfuric acid có tính acid, tính oxi hóa mạnh, tính háo nước:
\choiceTF
{\True Sulfuric acid đặc phản ứng với carbon trong than}
{\True Sulfuric acid đặc oxi hoá hợp chất Fe(II) thành hợp chất Fe(III)}
{Cho bột sodium hydrogencarbonate vào sulfuric acid loãng xuất hiện kết tủa trắng}
{\True Chất rắn copper (II) sulfate pentahydrate có màu xanh tác dụng với sulfuric acid đặc sẽ biến thành copper (II) sulfate khan màu trắng}
""" + (r"""
\loigiai{
\begin{itemize}
	\item a) Đúng: $\ce{2H2SO4_{(đặc)} + C ->[t^\circ] 2SO2 + CO2 + 2H2O}$.
	\item b) Đúng: $\ce{H2SO4}$ đặc nóng oxi hoá Fe(II) thành muối Fe(III):\\
	$\ce{2FeO + 4H2SO4_{(đặc)} ->[t^\circ] Fe2(SO4)3 + SO2 + 4H2O}$.
	\item c) Sai: $\ce{2NaHCO3 + H2SO4 -> Na2SO4 + 2CO2 ^ + 2H2O}$. Phản ứng sinh ra bọt khí không màu và dung dịch trong suốt, không xuất hiện kết tủa.
	\item d) Đúng: Do tính háo nước mãnh liệt của $\ce{H2SO4}$ đặc, nó hút nước kết tinh từ tinh thể ngậm nước:\\
	$\ce{CuSO4.5H2O (xanh) -> CuSO4 (trắng) + 5H2O}$.
\end{itemize}
}""" if is_teacher else "") + r"""
\end{ex}
"""

    tex += r"""
\vspace{0.3cm}
\noindent\textbf{\large PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.}
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}
\vspace{0.2cm}
"""

    # Phần III - Câu 1
    tex += r"""
\begin{ex}
Phú dưỡng là hệ quả sau khi ao ngòi, sông hồ nhận quá nhiều các nguồn thải chứa các chất dinh dưỡng chứa nguyên tố nitrogen, phosphorus. Khi hàm lượng phosphorus vượt quá $20\ \mu\text{g/L}$ và hàm lượng nitrogen vượt quá ngưỡng bao nhiêu $\mu\text{g/L}$ thì xảy ra hiện tượng phú dưỡng?
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{300}\\
Theo tiêu chuẩn sinh thái học môi trường thủy vực, khi hàm lượng nitrogen trong nước vượt quá $300\ \mu\text{g/L}$ và phosphorus vượt quá $20\ \mu\text{g/L}$ thì nguồn nước xảy ra hiện tượng phú dưỡng.
}""" if is_teacher else r"""
\vspace{1.0cm}
""") + r"""
\end{ex}
"""

    # Phần III - Câu 2
    tex += r"""
\begin{ex}
Cho các chất khí sau: $\ce{H2S, NO, NO2, SO2}$. Số khí gây ô nhiễm môi trường khi phát thải vào không khí là bao nhiêu?
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{4}\\
Cả 4 khí trên ($\ce{H2S, NO, NO2, SO2}$) đều là những chất khí gây ô nhiễm môi trường không khí và có độc tính đối với sức khỏe sinh vật.
}""" if is_teacher else r"""
\vspace{1.0cm}
""") + r"""
\end{ex}
"""

    # Phần III - Câu 3
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Tính khối lượng riêng (g/L) của không khí ở điều kiện chuẩn, giả thiết thành phần không khí gồm: 78\% nitrogen, 21\% oxygen và 1\% argon (làm tròn kết quả đến hàng phần trăm).
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{1,17}\\
Khối lượng mol trung bình của không khí:\\
$\overline{M}_\text{kk} = 0{,}78 \times 28{,}014 + 0{,}21 \times 31{,}998 + 0{,}01 \times 39{,}948 = 21{,}851 + 6{,}720 + 0{,}399 = 28{,}97\text{ g/mol}$.\\
Ở điều kiện chuẩn ($25^\circ\text{C}, 1\text{ bar}$), thể tích mol phân tử của chất khí là $V = 24{,}79\text{ L}$.\\
Khối lượng riêng: $D = \dfrac{m}{V} = \dfrac{28{,}97}{24{,}79} \approx 1{,}1686\text{ g/L} \approx \mathbf{1{,}17\text{ g/L}}$.
}""" if is_teacher else r"""
\vspace{1.5cm}
""") + r"""
\end{ex}
"""

    # Phần III - Câu 4
    tex += r"""
\begin{ex}
Trong công nghiệp, ammonia được sản xuất theo phản ứng pha khí:
\[\ce{N2(g) + 3H2(g) <=> 2NH3(g)} \quad \Delta_r H^\circ_{298}\]
Cho biết các giá trị năng lượng liên kết $E_\text{b}\ (\text{kJ}\cdot\text{mol}^{-1})$:
\begin{center}
\begin{tabular}{|c|c|c|c|}
\hline
\textbf{Liên kết} & $\ce{N\bond{3}N}$ & $\ce{H-H}$ & $\ce{N-H}$ \\ \hline
$E_\text{b}$ & 945 & 436 & 386 \\ \hline
\end{tabular}
\end{center}
Nhiệt tạo thành $\Delta_f H^\circ_{298}\ (\text{kJ}\cdot\text{mol}^{-1})$ của $\ce{NH3(g)}$ là $-x\text{ kJ}\cdot\text{mol}^{-1}$. Giá trị của $x$ là bao nhiêu?
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{31,5}\\
Biến thiên enthalpy phản ứng:\\
$\Delta_r H^\circ_{298} = E_\text{b}(\ce{N\bond{3}N}) + 3 \cdot E_\text{b}(\ce{H-H}) - 6 \cdot E_\text{b}(\ce{N-H}) = 945 + 3 \times 436 - 6 \times 386 = -63\text{ kJ}$.\\
Vì phản ứng tạo thành 2 mol $\ce{NH3(g)}$ từ các đơn chất ở trạng thái chuẩn, nên:\\
$\Delta_f H^\circ_{298}(\ce{NH3}) = \dfrac{\Delta_r H^\circ_{298}}{2} = \dfrac{-63}{2} = -31{,}5\text{ kJ/mol} \Rightarrow x = \mathbf{31{,}5}$.
}""" if is_teacher else r"""
\vspace{1.5cm}
""") + r"""
\end{ex}
"""

    # Phần III - Câu 5
    tex += r"""
\begin{ex}
Sự phụ thuộc của độ tan khí ammonia trong nước vào nhiệt độ được mô tả ở hình bên.
\begin{center}
	\includegraphics[width=0.48\textwidth]{images_crop_hoa11_de05/cau5_p3_do_thi_do_tan_nh3.png}
\end{center}
Dựa vào đồ thị ở hình bên, hãy xác định: Nồng độ phần trăm của dung dịch ammonia bão hòa ở $30^\circ\text{C}$ (làm tròn kết quả đến hàng phần mười).
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{28,6\%} (hoặc \textbf{28,6})\\
Từ đồ thị, tại $30^\circ\text{C}$, độ tan của khí $\ce{NH3}$ trong nước là $S = 40\text{ g }\ce{NH3}$ trong $100\text{ g }\ce{H2O}$.\\
Khối lượng dung dịch bão hòa: $m_\text{dd} = m_\text{ct} + m_\text{dm} = 40 + 100 = 140\text{ g}$.\\
Nồng độ phần trăm của dung dịch:\\
$C\% = \dfrac{40}{140} \times 100\% \approx 28{,}57\% \approx \mathbf{28{,}6\%}$.
}""" if is_teacher else r"""
\vspace{1.5cm}
""") + r"""
\end{ex}
"""

    # Phần III - Câu 6
    tex += r"""
\begin{ex}
Sau hiện tượng phú dưỡng, một hồ nước có diện tích bề mặt $1\text{ km}^2$ và độ sâu trung bình $5\text{ m}$ chứa một lượng lớn xác tảo. Lượng oxygen hòa tan (DO) ban đầu trong hồ là $8\text{ mg/L}$, nhưng sau quá trình phân hủy xác tảo, nồng độ oxygen giảm xuống còn $2\text{ mg/L}$. Lượng oxygen (tính bằng kg) đã bị tiêu thụ trong toàn bộ hồ là $x$ nghìn kg. Tính giá trị $x$?
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{30}\\
Thể tích nước trong hồ:\\
$V = S \cdot h = 1\text{ km}^2 \times 5\text{ m} = 10^6\text{ m}^2 \times 5\text{ m} = 5 \times 10^6\text{ m}^3 = 5 \times 10^9\text{ L}$.\\
Lượng oxygen tiêu thụ trên mỗi lít nước: $\Delta\text{DO} = 8 - 2 = 6\text{ mg/L}$.\\
Tổng lượng oxygen tiêu thụ:\\
$m(\ce{O2}) = 6\text{ mg/L} \times 5 \times 10^9\text{ L} = 30 \times 10^9\text{ mg} = 30\,000\text{ kg} = 30\text{ nghìn kg} \Rightarrow x = \mathbf{30}$.
}""" if is_teacher else r"""
\vspace{1.5cm}
""") + r"""
\end{ex}
"""

    tex += r"""
\vspace{0.3cm}
\noindent\textbf{\large PHẦN IV. Tự luận}
\textit{Học sinh giải chi tiết các câu hỏi sau vào giấy làm bài.}
\vspace{0.2cm}
"""

    # Phần IV - Câu 1
    tex += r"""
\begin{ex}
\textbf{[CĐ - SGK]} Viết phương trình hóa học minh họa tác động của mưa acid đối với calcium carbonate trong núi đá vôi và với kim loại iron có trong thép.
""" + (r"""
\loigiai{
Phương trình hóa học minh họa tác động của mưa acid:
\begin{itemize}
	\item Tác động đối với calcium carbonate ($\ce{CaCO3}$) trong núi đá vôi:
	\[\ce{CaCO3(s) + 2HNO3(aq) -> Ca(NO3)2(aq) + CO2(g) ^ + H2O(l)}\]
	\[\ce{CaCO3(s) + H2SO4(aq) -> CaSO4(aq, s) + CO2(g) ^ + H2O(l)}\]
	\item Tác động đối với kim loại iron ($\ce{Fe}$) có trong thép:
	\[\ce{Fe(s) + H2SO4(aq) -> FeSO4(aq) + H2(g) ^}\]
	\[\ce{Fe(s) + 4HNO3_{(loãng)} -> Fe(NO3)3(aq) + NO(g) ^ + 2H2O(l)}\]
\end{itemize}
}""" if is_teacher else r"""
\vspace{2.5cm}
""") + r"""
\end{ex}
"""

    # Phần IV - Câu 2
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Trong công nghiệp, ammonia được sản xuất theo phản ứng pha khí:
\[\ce{N2(g) + 3H2(g) <=> 2NH3(g)} \quad \Delta_r H^\circ_{298}\]
Cho biết các giá trị năng lượng liên kết $E_\text{b}\ (\text{kJ}\cdot\text{mol}^{-1})$:
\begin{center}
\begin{tabular}{|c|c|c|c|}
\hline
\textbf{Liên kết} & $\ce{N\bond{3}N}$ & $\ce{H-H}$ & $\ce{N-H}$ \\ \hline
$E_\text{b}$ & 945 & 436 & 386 \\ \hline
\end{tabular}
\end{center}
a) Tính nhiệt phản ứng $\Delta_r H^\circ_{298}$ của phản ứng ở điều kiện chuẩn, nhận xét về dấu và độ lớn của giá trị tìm được.\\
b) Tính nhiệt tạo thành $\Delta_f H^\circ_{298}\ (\text{kJ}\cdot\text{mol}^{-1})$ của $\ce{NH3(k)}$.
""" + (r"""
\loigiai{
a) Tính nhiệt phản ứng $\Delta_r H^\circ_{298}$ theo năng lượng liên kết:\\
$\Delta_r H^\circ_{298} = E_\text{b}(\ce{N\bond{3}N}) + 3 \cdot E_\text{b}(\ce{H-H}) - 6 \cdot E_\text{b}(\ce{N-H})$\\
$= 945 + 3 \times 436 - 6 \times 386 = 945 + 1308 - 2316 = -63\text{ kJ}$.\\
\textit{(Ghi chú: Nếu lấy số liệu $E_\text{b}(\ce{N-H}) = 389\text{ kJ/mol}$ theo một số tài liệu thì $\Delta_r H^\circ_{298} = -78\text{ kJ}$).}\\
\textbf{Nhận xét:} Giá trị $\Delta_r H^\circ_{298} < 0\ (-63\text{ kJ})$ chứng tỏ đây là một phản ứng tỏa nhiệt; nhiệt lượng tỏa ra khi tổng hợp 2 mol $\ce{NH3(g)}$ ở điều kiện chuẩn là $63\text{ kJ}$.\\
b) Tính nhiệt tạo thành $\Delta_f H^\circ_{298}$ của $\ce{NH3(g)}$:\\
Vì phản ứng trên tạo ra 2 mol $\ce{NH3(g)}$ từ các đơn chất $\ce{N2(g)}$ và $\ce{H2(g)}$ ở trạng thái chuẩn bền nhất, nên nhiệt tạo thành của 1 mol $\ce{NH3(g)}$ là:\\
$\Delta_f H^\circ_{298}(\ce{NH3, g}) = \dfrac{\Delta_r H^\circ_{298}}{2} = \dfrac{-63}{2} = \mathbf{-31{,}5\text{ kJ/mol}}$ \textit{(hoặc $-39\text{ kJ/mol}$)}.
}""" if is_teacher else r"""
\vspace{3.5cm}
""") + r"""
\end{ex}
"""

    if is_teacher:
        tex += r"""
\newpage
\noindent
\textbf{\large BẢNG ĐÁP ÁN VÀ THANG ĐIỂM THAM KHẢO}
\smallskip

\noindent
\textbf{1. PHẦN I (Trắc nghiệm nhiều phương án):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1. C} & \textbf{2. B} & \textbf{3. A} & \textbf{4. B} & \textbf{5. D} & \textbf{6. B} & \textbf{7. D} & \textbf{8. B} & \textbf{9. A} \\ \hline
\textbf{10. B} & \textbf{11. B} & \textbf{12. C} & \textbf{13. C} & \textbf{14. D} & \textbf{15. B} & \textbf{16. B} & \textbf{17. B} & \textbf{18. B} \\ \hline
\end{tabular}
\end{center}

\noindent
\textbf{2. PHẦN II (Trắc nghiệm đúng sai):}
\begin{itemize}
	\item \textbf{Câu 1:} a) Đúng \quad|\quad b) Đúng \quad|\quad c) Sai \quad|\quad d) Sai
	\item \textbf{Câu 2:} a) Đúng \quad|\quad b) Sai \quad|\quad c) Đúng \quad|\quad d) Đúng
	\item \textbf{Câu 3:} a) Đúng \quad|\quad b) Sai \quad|\quad c) Đúng \quad|\quad d) Đúng
	\item \textbf{Câu 4:} a) Đúng \quad|\quad b) Đúng \quad|\quad c) Sai \quad|\quad d) Đúng
\end{itemize}

\noindent
\textbf{3. PHẦN III (Trả lời ngắn):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\ \hline
300 & 4 & 1,17 & 31,5 & 28,6 & 30 \\ \hline
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
    hs_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De05_HocSinh.tex")
    with open(hs_path, "w", encoding="utf-8") as f:
        f.write(hs_tex)
    print("Saved:", hs_path)

    gv_tex = make_tex(is_teacher=True)
    gv_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De05_GiaoVien.tex")
    with open(gv_path, "w", encoding="utf-8") as f:
        f.write(gv_tex)
    print("Saved:", gv_path)

    # Compile with pdflatex
    for name in ["De_Kiem_Tra_Hoa_11_De05_HocSinh", "De_Kiem_Tra_Hoa_11_De05_GiaoVien"]:
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
                if file.endswith(ext) or file.startswith("ans-hoa11-de05"):
                    try:
                        os.remove(os.path.join(root, file))
                    except:
                        pass
    print("LaTeX generation completed successfully!")

if __name__ == "__main__":
    main()
