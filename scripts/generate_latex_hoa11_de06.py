# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file LaTeX và biên dịch ra PDF cho ĐỀ SỐ 6 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De06_HocSinh.tex / .pdf
- De_Kiem_Tra_Hoa_11_De06_GiaoVien.tex / .pdf
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
    ans_file = f"ans-hoa11-de06{ver_tag}"
    
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
	\fbox{\textbf{MÃ ĐỀ: 106}}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Ag = 108; Ba = 137; Au = 197).}
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
\textbf{(SBT - CTST)} Ở trạng thái tự nhiên, nitrogen
\choice
{\True tồn tại ở dạng đơn chất và hợp chất}
{chỉ tồn tại ở dạng đơn chất}
{chỉ tồn tại ở dạng hợp chất}
{tự do chiếm khoảng 20\% thể tích không khí}
""" + (r"""
\loigiai{
Trong tự nhiên, nitrogen tồn tại ở cả dạng đơn chất (khí $\ce{N2}$ chiếm khoảng 78\% thể tích khí quyển) và dạng hợp chất (trong khoáng vật diêm tiêu $\ce{NaNO3, KNO3}$, protein, nucleic acid và sinh khối động - thực vật).
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 2
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Trong phương pháp Ostwald, ammonia bị oxi hóa bởi oxygen không khí tạo thành sản phẩm chính là
\choice
{\True $\ce{NO}$}
{$\ce{N2}$}
{$\ce{N2O}$}
{$\ce{NO2}$}
""" + (r"""
\loigiai{
Trong phương pháp Ostwald sản xuất $\ce{HNO3}$, ammonia được oxi hóa bởi oxygen không khí ở $850 - 900^\circ\text{C}$ với xúc tác $\ce{Pt}$ tạo sản phẩm chính là nitrogen monoxide ($\ce{NO}$):\\
$\ce{4NH3 + 5O2 ->[Pt, t^\circ] 4NO + 6H2O}$.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 3
    tex += r"""
\begin{ex}
\textbf{[KNTT - SBT]} Cho các chất khí sau: $\ce{H2S, NO, NO2, SO2}$. Số khí gây ô nhiễm môi trường khi phát thải vào không khí là
\choice
{1}
{\True 4}
{3}
{2}
""" + (r"""
\loigiai{
Cả 4 khí trên ($\ce{H2S, NO, NO2, SO2}$) đều là các khí độc hại gây ô nhiễm môi trường không khí nghiêm trọng và là tác nhân chính gây ra mưa acid và sương mù quang hóa.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 4
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Oxide của nitrogen được tạo thành khi nguyên tố nitrogen trong nhiên liệu hoặc sinh khối kết hợp với oxygen dư thừa trong không khí được gọi là
\choice
{$\ce{NO_x}$ tức thời}
{$\ce{NO_x}$ nhiệt}
{\True $\ce{NO_x}$ nhiên liệu}
{$\ce{NO_x}$ tự nhiên}
""" + (r"""
\loigiai{
Khi đốt cháy than đá, dầu mỏ hoặc sinh khối chứa hợp chất nitrogen hữu cơ, các nguyên tử N liên kết bị oxi hóa tạo thành oxide của nitrogen; nguồn này được gọi là $\ce{NO_x}$ nhiên liệu (Fuel $\ce{NO_x}$).
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 5
    tex += r"""
\begin{ex}
\textbf{[CD - SBT]} Phát biểu nào sau đây về đơn chất nitrogen ($\ce{N2}$) là \textbf{không} đúng?
\choice
{\True Dù phân tử $\ce{N2}$ có tính kém hoạt động hóa học, nhưng vẫn hoạt động hóa học mạnh hơn chlorine, $\ce{Cl2}$}
{Đơn chất nitrogen không phản ứng với hydrogen, oxygen ở điều kiện thường}
{Do có nhiệt độ rất thấp nên nitrogen lỏng được sử dụng bảo quản một số loại mẫu vật}
{Trong bầu khí quyển, khi có sấm chớp, khí nitrogen tạo các nitrogen oxide, là một nguyên nhân làm cho nước mưa có tính acid}
""" + (r"""
\loigiai{
Phát biểu A sai vì phân tử $\ce{N2}$ có liên kết ba $\ce{N\bond{3}N}$ vô cùng bền vững (năng lượng liên kết $945\text{ kJ/mol}$) nên ở nhiệt độ thường trơ và kém hoạt động hơn rất nhiều so với khí chlorine ($\ce{Cl2}$ hoạt động hóa học mạnh).
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 6
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Ở điều kiện thích hợp, sulfur dioxide đóng vai trò là chất oxi hoá khi tham gia phản ứng với chất nào sau đây?
\choice
{$\ce{NO2}$}
{\True $\ce{H2S}$}
{$\ce{NaOH}$}
{$\ce{Ca(OH)2}$}
""" + (r"""
\loigiai{
Trong phản ứng: $\ce{SO2 + 2H2S -> 3S + 2H2O}$, số oxi hóa của sulfur trong $\ce{SO2}$ giảm từ +4 xuống 0, do đó $\ce{SO2}$ đóng vai trò là chất oxi hóa ($\ce{H2S}$ là chất khử).
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 7
    tex += r"""
\begin{ex}
\textbf{(SBT - CTST)} Điều nào sau đây đúng về tính chất hóa học của dung dịch $\ce{H2SO4}$ đặc?
\choice
{Dung dịch $\ce{H2SO4}$ đặc có tính khử mạnh}
{\True Dung dịch $\ce{H2SO4}$ đặc có tính oxi hóa mạnh}
{Dung dịch $\ce{H2SO4}$ đặc vừa có tính khử, vừa có tính oxi hóa}
{Dung dịch $\ce{H2SO4}$ đặc không có tính khử, không có tính oxi hóa}
""" + (r"""
\loigiai{
Trong $\ce{H2SO4}$ đặc, nguyên tử sulfur đạt số oxi hóa cực đại là +6 nên dung dịch $\ce{H2SO4}$ đặc có tính oxi hóa rất mạnh và tính háo nước mạnh.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 8
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Xét phản ứng trong quá trình tạo ra $\ce{NO_x}$ nhiệt:
\[\ce{N2(g) + O2(g) -> 2NO(g)} \quad \Delta_r H^\circ_{298} = 180{,}6\text{ kJ}\]
Nhiệt tạo thành chuẩn của $\ce{NO(g)}$ là
\choice
{$180{,}6\text{ kJ/mol}$}
{$-180{,}6\text{ kJ/mol}$}
{$-90{,}3\text{ kJ/mol}$}
{\True $90{,}3\text{ kJ/mol}$}
""" + (r"""
\loigiai{
Phản ứng trên tạo ra 2 mol $\ce{NO(g)}$ từ các đơn chất ở trạng thái chuẩn:\\
$\Delta_r H^\circ_{298} = 2 \cdot \Delta_f H^\circ_{298}(\ce{NO, g}) = 180{,}6\text{ kJ} \Rightarrow \Delta_f H^\circ_{298}(\ce{NO, g}) = \dfrac{180{,}6}{2} = \mathbf{90{,}3\text{ kJ/mol}}$.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 9
    tex += r"""
\begin{ex}
\textbf{(SBT - KNTT)} Xét cân bằng hóa học: $\ce{N2(g) + 3H2(g) <=> 2NH3(g)} \quad \Delta_r H^\circ_{298} < 0$.\\
Hiệu suất phản ứng khi hệ đạt cân bằng ở nhiệt độ $400^\circ\text{C}$ và $500^\circ\text{C}$ lần lượt bằng $x\%$ và $y\%$. Mối quan hệ giữa $x$ và $y$ là
\choice
{$x < y$}
{$x = y$}
{\True $x > y$}
{$5x = 4y$}
""" + (r"""
\loigiai{
Phản ứng tổng hợp $\ce{NH3}$ theo chiều thuận là phản ứng tỏa nhiệt ($\Delta_r H^\circ_{298} < 0$). Khi tăng nhiệt độ từ $400^\circ\text{C}$ lên $500^\circ\text{C}$, theo nguyên lý Le Chatelier cân bằng chuyển dịch theo chiều thu nhiệt (chiều nghịch), làm giảm hiệu suất tổng hợp ammonia $\Rightarrow x > y$.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 10
    tex += r"""
\begin{ex}
Muốn cho cân bằng của phản ứng tổng hợp ammonia chuyển dịch sang phải cần đồng thời:
\[\ce{N2 + 3H2 <=> 2NH3} \quad \Delta_r H^\circ_{298} < 0\]
\choice
{Tăng áp suất và tăng nhiệt độ}
{\True Tăng áp suất và giảm nhiệt độ}
{Giảm áp suất và giảm nhiệt độ}
{Giảm áp suất và tăng nhiệt độ}
""" + (r"""
\loigiai{
- Về áp suất: Phía chất đầu có 4 mol khí, phía sản phẩm có 2 mol khí. Tăng áp suất làm cân bằng chuyển dịch theo chiều giảm số mol khí (chiều thuận sang phải).\\
- Về nhiệt độ: Phản ứng tỏa nhiệt ($\Delta_r H^\circ_{298} < 0$), giảm nhiệt độ làm cân bằng chuyển dịch theo chiều tỏa nhiệt (chiều thuận sang phải).
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 11
    tex += r"""
\begin{ex}
Hợp chất vô cơ X có các tính chất: X tinh khiết là chất lỏng không màu, bốc khói mạnh trong không khí ẩm; dung dịch X đặc để lâu có màu vàng; tan tốt trong nước; có tính oxi hóa mạnh. Chất X là
\choice
{$\ce{H3PO4}$}
{\True $\ce{HNO3}$}
{$\ce{H2SO4}$}
{$\ce{HCl}$}
""" + (r"""
\loigiai{
Nitric acid ($\ce{HNO3}$) tinh khiết là chất lỏng không màu, bốc khói mạnh trong không khí ẩm. Dưới tác dụng của ánh sáng, $\ce{HNO3}$ bị phân hủy một phần sinh ra khí $\ce{NO2}$ có màu nâu đỏ tan lại trong acid làm dung dịch có màu vàng.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 12
    tex += r"""
\begin{ex}
\textbf{[KNTT - SBT]} Cho cân bằng hóa học sau: $\ce{2SO2(g) + O2(g) <=> 2SO3(g)} \quad \Delta_r H^\circ_{298} < 0$.\\
Khi tăng nhiệt độ thì
\choice
{Tổng số mol khí trong hệ giảm}
{Hiệu suất phản ứng tăng}
{\True Cân bằng chuyển dịch theo chiều nghịch}
{Nồng độ khí sản phẩm tăng}
""" + (r"""
\loigiai{
Vì phản ứng theo chiều thuận tỏa nhiệt ($\Delta_r H^\circ_{298} < 0$), khi tăng nhiệt độ theo nguyên lý Le Chatelier cân bằng chuyển dịch theo chiều thu nhiệt (chiều nghịch).
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 13
    tex += r"""
\begin{ex}
\textbf{[KNTT - SBT]} Trong công nghiệp sản xuất sulfuric acid, hai nguồn nguyên liệu được khai thác từ mỏ để cung cấp nguyên tố lưu huỳnh là
\choice
{$\ce{ZnS, PbS}$}
{$\ce{H2S, SO2}$}
{$\ce{CaSO4, BaSO4}$}
{\True $\ce{S, FeS2}$}
""" + (r"""
\loigiai{
Hai nguồn nguyên liệu mỏ chính để sản xuất sulfuric acid trong công nghiệp là lưu huỳnh đơn chất ($\ce{S}$) và quặng pyrite sắt ($\ce{FeS2}$).
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 14
    tex += r"""
\begin{ex}
Nếu ta dẫn khí $\ce{SO2}$ từ từ đến dư vào dung dịch bromine sẽ có phản ứng hóa học xảy ra kèm hiện tượng gì sau đây?
\choice
{Dung dịch không đổi màu}
{Dung dịch bị vẩn đục do tạo kết tủa và kết tủa không tan}
{Dung dịch sẽ có màu đỏ nâu đậm dần lên hơn lúc ban đầu}
{\True Dung dịch sẽ nhạt màu đỏ nâu rồi mất hẳn màu}
""" + (r"""
\loigiai{
Khí $\ce{SO2}$ có tính khử mạnh khử dung dịch bromine màu đỏ nâu tạo thành dung dịch trong suốt không màu:\\
$\ce{SO2 + Br2 + 2H2O -> 2HBr + H2SO4}$.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 15
    tex += r"""
\begin{ex}
Hậu quả chính của hiện tượng phú dưỡng đối với môi trường nước là gì?
\choice
{Làm tăng độ trong suốt của nước và giúp cải thiện hệ sinh thái thủy sinh}
{\True Làm giảm lượng oxygen hòa tan trong nước, gây ra hiện tượng cá chết hàng loạt và phá hủy hệ sinh thái}
{Tăng nồng độ oxygen hòa tan trong nước, giúp thủy sinh vật phát triển mạnh mẽ hơn}
{Giảm sự phát triển của vi sinh vật trong nước}
""" + (r"""
\loigiai{
Khi phú dưỡng xảy ra, tảo bùng phát làm giảm ánh sáng; khi tảo chết đi quá trình phân hủy xác tảo tiêu thụ lượng lớn oxygen hòa tan (DO), dẫn đến cạn kiệt dưỡng khí gây ngạt và chết hàng loạt các sinh vật thủy sinh.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 16
    tex += r"""
\begin{ex}
Loại khí nào dưới đây là nguyên nhân chính gây ra mưa acid?
\choice
{Khí $\ce{CO2}$}
{Khí $\ce{O2}$}
{\True Khí $\ce{SO2}$ và $\ce{NO_x}$}
{Khí $\ce{CH4}$}
""" + (r"""
\loigiai{
Sulfur dioxide ($\ce{SO2}$) và các oxide của nitrogen ($\ce{NO_x}$) là hai tác nhân chính gây ra hiện tượng mưa acid khi phản ứng với nước và oxygen trong khí quyển tạo thành $\ce{H2SO4}$ và $\ce{HNO3}$.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 17
    tex += r"""
\begin{ex}
Khu vực nào có nguy cơ bị ảnh hưởng nghiêm trọng bởi mưa acid?
\choice
{Những khu vực có nhiều cây cối và hệ sinh thái phong phú}
{\True Những khu vực gần các nhà máy điện, nhà máy công nghiệp và giao thông dày đặc}
{Những khu vực có độ cao lớn}
{Những khu vực gần biển}
""" + (r"""
\loigiai{
Những khu vực gần nhà máy nhiệt điện, cơ sở luyện kim và đô thị mật độ giao thông dày đặc là nơi phát thải tập trung lượng lớn khí $\ce{SO2}$ và $\ce{NO_x}$, nên có nguy cơ cao nhất hứng chịu mưa acid.
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Câu 18
    tex += r"""
\begin{ex}
Cho các phát biểu sau:
\begin{enumerate}[(1)]
	\item Phương pháp tiếp xúc gồm ba công đoạn chính sản xuất: $\ce{SO2, SO3, H2SO4}$.
	\item Ở giai đoạn tạo ra $\ce{SO3}$, người ta chọn điều kiện phản ứng ở nhiệt độ thấp.
	\item Nhờ có chất xúc tác nên phản ứng giữa $\ce{SO2}$ và $\ce{O2}$ ưu tiên diễn ra theo chiều thuận.
	\item Khí sulfur trioxide đi từ dưới lên đỉnh tháp, sulfuric acid đặc chảy từ đỉnh tháp xuống dưới.
	\item Tương tác van der Waals giữa các phân tử nitrogen với nước rất yếu nên ít tan trong nước.
	\item Phân tử nitrogen nặng hơn không khí và không phân cực.
	\item Nitrogen lỏng ở nhiệt độ thấp, $-196^\circ\text{C}$ và có tính trơ.
\end{enumerate}
Số phát biểu đúng là
\choice
{3}
{\True 4}
{5}
{6}
""" + (r"""
\loigiai{
Có 4 phát biểu đúng gồm: (1), (4), (5), (7).\\
- (1) Đúng: Ba công đoạn là sản xuất $\ce{SO2} \rightarrow$ oxi hóa thành $\ce{SO3} \rightarrow$ hấp thụ tạo $\ce{H2SO4}$.\\
- (2) Sai: Thực tế tiến hành ở $450^\circ\text{C}$ để đảm bảo tốc độ phản ứng.\\
- (3) Sai: Chất xúc tác chỉ làm tăng tốc độ đạt cân bằng chứ không làm dịch chuyển cân bằng.\\
- (4) Đúng: Chuyển động ngược chiều tối đa hóa tiếp xúc giữa khí và acid.\\
- (5) Đúng: Phân tử không phân cực, tương tác van der Waals với nước rất yếu nên ít tan.\\
- (6) Sai: $M(\ce{N2}) = 28$ nhẹ hơn không khí ($29$).\\
- (7) Đúng: Nitrogen lỏng có nhiệt độ sôi $-196^\circ\text{C}$ và rất trơ hóa học.
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
Hợp chất có công thức hóa học $\ce{NH4NO3}$ được giới chức quốc gia Lebanon xác định là nguyên nhân gây ra vụ nổ thảm khốc ở thủ đô Beirut vào ngày 04/08/2020. Tia lửa hàn trong quá trình sửa chữa nhà kho có thể đã làm 2 750 tấn $\ce{NH4NO3}$ cất trữ phát nổ, phá hủy nhiều nhà cửa, dẫn đến nhiều người thiệt mạng.
\[\ce{NH4NO3 ->[t^\circ] N2O + 2H2O} \quad \Delta_r H^\circ_{298} = -35{,}9\text{ kJ}\]
\choiceTF
{Ammonium nitrate khi ở nhiệt độ cao bị phân hủy thành khí $\ce{N2O}$ và hơi nước, là một phản ứng thu nhiệt và năng lượng lớn}
{\True Khi phản ứng nổ xảy ra, năng lượng được giải phóng một cách đột ngột dưới áp lực rất cao, tăng nhanh, còn được gọi là sóng nổ hoặc sóng xung kích}
{\True Sóng xung kích gây ra thiệt hại lớn cho môi trường xung quanh. Ammonium nitrate có thể tự phân hủy qua thời gian}
{\True Tia lửa hàn trong quá trình sửa chữa nhà kho đã khơi mào phản ứng phân hủy ammonium nitrate gây nổ}
""" + (r"""
\loigiai{
\begin{itemize}
	\item a) Sai: Vì $\Delta_r H^\circ_{298} = -35{,}9\text{ kJ} < 0$ nên phản ứng phân hủy ammonium nitrate là phản ứng tỏa nhiệt mạnh.
	\item b) Đúng: Sự giải phóng thể tích khí khổng lồ và nhiệt năng cực nhanh tạo nên sóng nổ/sóng xung kích áp lực cao.
	\item c) Đúng: Sóng xung kích tàn phá các công trình; $\ce{NH4NO3}$ bảo quản lâu ngày có thể tự phân hủy sinh nhiệt tích tụ.
	\item d) Đúng: Tia lửa hàn cung cấp nhiệt lượng ban đầu khơi mào phản ứng dây chuyền gây nổ.
\end{itemize}
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Phần II - Câu 2
    tex += r"""
\begin{ex}
Nguồn nào sau đây phát sinh oxide của nitrogen trong không khí?
\choiceTF
{\True Đốt cháy các nhiên liệu hóa thạch}
{\True Quá trình sản xuất và sử dụng nitric acid}
{\True Trong khí quyển khi có sấm sét}
{Quá trình quang hợp của cây xanh}
""" + (r"""
\loigiai{
\begin{itemize}
	\item a) Đúng: Đốt cháy than đá, xăng dầu ở nhiệt độ cao tạo $\ce{NO_x}$ nhiệt và $\ce{NO_x}$ nhiên liệu.
	\item b) Đúng: Khí thải công nghiệp sản xuất và sử dụng $\ce{HNO3}$ chứa oxide của nitrogen.
	\item c) Đúng: Sấm sét tạo nhiệt độ cực cao làm kết hợp $\ce{N2}$ và $\ce{O2}$ thành $\ce{NO}$.
	\item d) Sai: Quá trình quang hợp của cây xanh hấp thụ $\ce{CO2}$ và giải phóng $\ce{O2}$, không tạo $\ce{NO_x}$.
\end{itemize}
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Phần II - Câu 3
    tex += r"""
\begin{ex}
Sulfur dioxide là một chất có nhiều ứng dụng trong công nghiệp (dùng để sản xuất sulfuric acid, tẩy trắng bột giấy công nghiệp giấy, tẩy trắng dung dịch đường trong sản xuất đường tinh luyện,...) và giúp ngăn cản sự phát triển của một số loại vi khuẩn và nấm gây hại cho thực phẩm. Ở áp suất 1 bar và nhiệt độ $25^\circ\text{C}$, phản ứng giữa 1 mol sulfur với oxygen xảy ra theo phương trình:
\[\ce{S(s) + O2(g) -> SO2(g)}\]
và tỏa ra một lượng nhiệt là $296{,}9\text{ kJ}$.
\begin{center}
	\includegraphics[width=0.75\textwidth]{images_crop_hoa11_de06/cau3_p2_so2_ungdung.png}
\end{center}
\choiceTF
{Biến thiên enthalpy chuẩn của phản ứng là $296{,}9\text{ kJ}$}
{\True Enthalpy tạo thành chuẩn của sulfur dioxide bằng $-296{,}9\text{ kJ}\cdot\text{mol}^{-1}$}
{\True Sulfur dioxide vừa có thể là chất khử vừa có thể là chất oxi hóa, tùy thuộc vào phản ứng mà nó tham gia}
{\True 0,5 mol sulfur tác dụng hết oxygen giải phóng $148{,}45\text{ kJ}$ năng lượng dưới dạng nhiệt}
""" + (r"""
\loigiai{
\begin{itemize}
	\item a) Sai: Vì phản ứng tỏa nhiệt nên biến thiên enthalpy chuẩn mang dấu âm: $\Delta_r H^\circ_{298} = -296{,}9\text{ kJ}$.
	\item b) Đúng: Vì $\ce{SO2(g)}$ được tạo thành từ các đơn chất bền ở điều kiện chuẩn nên $\Delta_f H^\circ_{298}(\ce{SO2, g}) = -296{,}9\text{ kJ/mol}$.
	\item c) Đúng: Số oxi hóa của S trong $\ce{SO2}$ là +4 (trung gian) nên vừa thể hiện tính khử vừa thể hiện tính oxi hóa.
	\item d) Đúng: Nhiệt lượng giải phóng khi đốt cháy 0,5 mol sulfur là $0{,}5 \times 296{,}9 = 148{,}45\text{ kJ}$.
\end{itemize}
}""" if is_teacher else "") + r"""
\end{ex}
"""

    # Phần II - Câu 4
    tex += r"""
\begin{ex}
Sulfuric acid có hoạt tính hóa học rất mạnh:
\choiceTF
{Nhôm (aluminum) bị thụ động trong dung dịch $\ce{H2SO4}$ loãng nguội}
{\True Hỗn hợp gồm $\ce{Cu}$ và $\ce{Fe2O3}$ có tỉ lệ mol 1:1 tan hết trong dung dịch $\ce{H2SO4}$ loãng}
{\True Khi cho $\ce{H2SO4}$ đặc vào đường kính trắng ($\ce{C12H22O11}$), khí thoát ra làm mất màu dung dịch $\ce{KMnO4}$}
{$\ce{H2SO4}$ đặc có thể dùng để làm khô các khí $\ce{CO2, Cl2, N2, H2S}$}
""" + (r"""
\loigiai{
\begin{itemize}
	\item a) Sai: Kim loại Al chỉ bị thụ động trong $\ce{H2SO4}$ đặc nguội; trong dung dịch $\ce{H2SO4}$ loãng nguội Al tan nhanh giải phóng $\ce{H2}$.
	\item b) Đúng: $\ce{Fe2O3 + 3H2SO4 -> Fe2(SO4)3 + 3H2O}$; sau đó $\ce{Cu + Fe2(SO4)3 -> CuSO4 + 2FeSO4}$. Hỗn hợp tỉ lệ mol 1:1 tan hoàn toàn.
	\item c) Đúng: $\ce{H2SO4}$ đặc than hóa đường rồi oxi hóa carbon sinh khí $\ce{SO2}$ làm mất màu dung dịch thuốc tím $\ce{KMnO4}$.
	\item d) Sai: Khí $\ce{H2S}$ có tính khử mạnh phản ứng với $\ce{H2SO4}$ đặc nên không thể dùng $\ce{H2SO4}$ đặc làm khô khí $\ce{H2S}$.
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
Cho các chất sau: $\ce{H2SO4, CO2, CO, N2, NH3}$. Số chất tan tốt trong nước ở nhiệt độ thường là bao nhiêu?
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{2}\\
Có 2 chất tan tốt trong nước ở điều kiện thường là $\ce{H2SO4}$ (tan vô hạn và tỏa nhiều nhiệt) và $\ce{NH3}$ (1 lít nước hòa tan khoảng 700 lít $\ce{NH3}$ ở $20^\circ\text{C}$).
}""" if is_teacher else r"""
\vspace{1.0cm}
""") + r"""
\end{ex}
"""

    # Phần III - Câu 2
    tex += r"""
\begin{ex}
Cho các phản ứng:
\begin{enumerate}[(a)]
	\item $\ce{S + O2 ->[t^\circ] SO2}$
	\item $\ce{S + 3F2 -> SF6}$
	\item $\ce{Hg + S -> HgS}$
	\item $\ce{H2 + S ->[t^\circ] H2S}$
\end{enumerate}
Số phản ứng trong đó sulfur đơn chất đóng vai trò chất oxi hóa là bao nhiêu?
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{2}\\
Sulfur đóng vai trò là chất oxi hóa khi số oxi hóa giảm từ 0 xuống -2. Điều này xảy ra ở phản ứng (c) và (d). Ở phản ứng (a) và (b), số oxi hóa của S tăng lên +4 và +6 nên S là chất khử.
}""" if is_teacher else r"""
\vspace{1.0cm}
""") + r"""
\end{ex}
"""

    # Phần III - Câu 3
    tex += r"""
\begin{ex}
Cho các chất: $\ce{HCl}$ (aq hoặc khí), $\ce{O2}\ (t^\circ)$, $\ce{AlCl3}$ (aq), $\ce{H2SO4}$ (aq), $\ce{NaOH}$ (aq), $\ce{Na2CO3}$ (aq), $\ce{CuO}$. Có bao nhiêu chất phản ứng được với $\ce{NH3}$?
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{5}\\
Có 5 chất phản ứng được với $\ce{NH3}$ gồm: $\ce{HCl, O2, AlCl3, H2SO4, CuO}$.
}""" if is_teacher else r"""
\vspace{1.5cm}
""") + r"""
\end{ex}
"""

    # Phần III - Câu 4
    tex += r"""
\begin{ex}
Cho hỗn hợp khí (X) gồm $\ce{N2, H2, NH3}$ có tỉ khối so với khí hydrogen là 8. Dẫn hỗn hợp khí (X) đi qua dung dịch $\ce{H2SO4}$ đặc, dư thì thể tích khí còn lại một nửa. Thành phần \% theo thể tích của $\ce{NH3}$ trong hỗn hợp (X) là $x\%$. Xác định giá trị $x$.
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{50}\\
Dung dịch $\ce{H2SO4}$ đặc hấp thụ toàn bộ lượng khí $\ce{NH3}$:\\
$\ce{2NH3 + H2SO4 -> (NH4)2SO4}$.\\
Thể tích khí còn lại là của $\ce{N2}$ và $\ce{H2}$, bằng một nửa thể tích ban đầu $\Rightarrow$ thể tích khí $\ce{NH3}$ chiếm 50\% thể tích hỗn hợp (X) $\Rightarrow x = \mathbf{50}$.
}""" if is_teacher else r"""
\vspace{1.5cm}
""") + r"""
\end{ex}
"""

    # Phần III - Câu 5
    tex += r"""
\begin{ex}
Tại nhiều làng nghề thủ công mĩ nghệ, sulfur dioxide được dùng là chất chống mốc cho các sản phẩm mây tre đan. Trong một ngày, một làng nghề đốt cháy $20\text{ kg}$ sulfur để tạo thành sulfur dioxide. Giả thiết có 20\% lượng khí $\ce{SO2}$ trên bay vào khí quyển và chuyển hoá hết thành $\ce{H2SO4}$ trong nước mưa theo sơ đồ:
\[\ce{SO2 ->[+O2, xt] SO3 ->[+H2O] H2SO4}\]
Thể tích nước mưa là $x$ nghìn $\text{m}^3$ bị nhiễm acid nếu nồng độ $\ce{H2SO4}$ trong nước mưa là $1{,}25 \cdot 10^{-5}\text{ M}$. Xác định giá trị của $x$?
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{10}\\
Số mol sulfur đốt cháy: $n_\text{S} = \dfrac{20\,000}{32} = 625\text{ mol} \Rightarrow n(\ce{SO2}) = 625\text{ mol}$.\\
Lượng $\ce{SO2}$ bay vào khí quyển: $n(\ce{SO2\ thoát}) = 0{,}20 \times 625 = 125\text{ mol}$.\\
Theo sơ đồ: $n(\ce{H2SO4}) = n(\ce{SO2}) = 125\text{ mol}$.\\
Thể tích nước mưa bị nhiễm acid:\\
$V = \dfrac{n}{C_\text{M}} = \dfrac{125}{1{,}25 \cdot 10^{-5}} = 10\,000\,000\text{ L} = 10\,000\text{ m}^3 = 10\text{ nghìn m}^3 \Rightarrow x = \mathbf{10}$.
}""" if is_teacher else r"""
\vspace{1.5cm}
""") + r"""
\end{ex}
"""

    # Phần III - Câu 6
    tex += r"""
\begin{ex}
Tại một nhà máy phân bón, ammophos được sản xuất từ ammonia và phosphoric acid, thu được $\ce{NH4H2PO4}$ và $\ce{(NH4)2HPO4}$ với tỉ lệ mol là 1:1. Cho khí ammonia tác dụng vừa đủ với dung dịch chứa $5{,}88\text{ tấn}$ phosphoric acid. Tính khối lượng ammophos thu được (theo đơn vị tấn)?
""" + (r"""
\loigiai{
\textbf{Đáp số:} \textbf{7,41}\\
Số mol $\ce{H3PO4} = \dfrac{5{,}88 \cdot 10^6}{98} = 60\,000\text{ mol}$.\\
Gọi số mol của $\ce{NH4H2PO4}$ và $\ce{(NH4)2HPO4}$ lần lượt là $a$ và $a$ mol (tỉ lệ 1:1):\\
Bảo toàn nguyên tố P: $a + a = 60\,000 \Rightarrow a = 30\,000\text{ mol}$.\\
Số mol $\ce{NH3}$ phản ứng: $n(\ce{NH3}) = a + 2a = 3a = 90\,000\text{ mol}$.\\
Khối lượng $\ce{NH3}$: $m(\ce{NH3}) = 90\,000 \times 17 = 1\,530\,000\text{ g} = 1{,}53\text{ tấn}$.\\
Theo định luật bảo toàn khối lượng:\\
$m(\text{ammophos}) = m(\ce{H3PO4}) + m(\ce{NH3}) = 5{,}88 + 1{,}53 = \mathbf{7{,}41\text{ tấn}}$.
}""" if is_teacher else r"""
\vspace{1.5cm}
""") + r"""
\end{ex}
"""

    # ==========================================
    # PHẦN IV
    # ==========================================
    tex += r"""
\vspace{0.3cm}
\noindent\textbf{\large PHẦN IV. Tự luận}
\textit{Học sinh giải chi tiết các câu hỏi sau vào giấy làm bài.}
\vspace{0.2cm}
"""

    # Phần IV - Câu 1
    tex += r"""
\begin{ex}
\textbf{(SBT - CD)} Vàng tan trong hỗn hợp gồm dung dịch nitric acid đặc và dung dịch hydrochloric acid đặc (tỉ lệ 1 : 3 về thể tích) tạo ra hợp chất tan của $\ce{Au^{3+}}$ theo phản ứng sau:
\[\ce{Au + HNO3 + HCl -> HAuCl4 + H2O + NO}\]
a) Cân bằng phương trình hoá học của phản ứng trên theo phương pháp thăng bằng electron.\\
b) Cho biết acid nào đóng vai trò chất oxi hoá trong phản ứng trên. Giải thích.
""" + (r"""
\loigiai{
a) Cân bằng theo phương pháp thăng bằng electron:
\begin{align*}
	\ce{Au^0 + 4Cl- &-> Au^{+3}Cl4- + 3e} &|\times 1 \\
	\ce{N^{+5} + 3e &-> N^{+2} (trong\ NO)} &|\times 1
\end{align*}
Phương trình hóa học hoàn chỉnh:
\[\ce{Au + HNO3 + 4HCl -> HAuCl4 + 2H2O + NO ^}\]
b) Trong phản ứng trên, $\ce{HNO3}$ đóng vai trò là chất oxi hóa vì số oxi hóa của nguyên tử nitrogen giảm từ +5 (trong $\ce{HNO3}$) xuống +2 (trong khí $\ce{NO}$). Acid $\ce{HCl}$ đóng vai trò môi trường và chất tạo phức.
}""" if is_teacher else r"""
\vspace{2.5cm}
""") + r"""
\end{ex}
"""

    # Phần IV - Câu 2
    tex += r"""
\begin{ex}
\textbf{(SBT - CTST)} Đặt hai cốc (A), (B) có khối lượng bằng nhau lên 2 đĩa cân, cân ở vị trí thăng bằng. Cho $120\text{ gam}$ hỗn hợp potassium hydrogencarbonate và sodium hydrogencarbonate vào cốc (A); $85\text{ gam}$ silver nitrate vào cốc (B). Thêm từ từ $100\text{ gam}$ dung dịch sulfuric acid 19,6\% vào cốc (A); $100\text{ gam}$ dung dịch hydrochloric acid 36,5\% vào cốc (B). Sau thí nghiệm cân có ở vị trí thăng bằng không? Nếu cân không ở vị trí thăng bằng thì cần thêm bao nhiêu gam dung dịch hydrochloric acid 36,5\% vào cốc nào để cân trở lại vị trí thăng bằng? Giả thiết khí $\ce{CO2}$ không tan trong nước, bỏ qua quá trình bay hơi của nước và hydrogen chloride.
""" + (r"""
\loigiai{
Hướng dẫn giải chi tiết:
\begin{itemize}
	\item \textbf{Xét cốc (A):} Đặt công thức chung hai muối là $\ce{RHCO3}$ ($M = 84 \div 100$).\\
	Số mol tối đa của muối: $n(\ce{RHCO3}) \le \dfrac{120}{84} \approx 1{,}43\text{ mol}$.\\
	Số mol acid: $n(\ce{H2SO4}) = \dfrac{100 \times 19{,}6\%}{98} = 0{,}2\text{ mol}$.\\
	Phương trình: $\ce{2RHCO3 + H2SO4 -> R2SO4 + 2CO2 ^ + 2H2O}$.\\
	Vì $\dfrac{n(\ce{RHCO3})}{2} > 0{,}7 > n(\ce{H2SO4}) = 0{,}2$ nên $\ce{RHCO3}$ dư, $\ce{H2SO4}$ hết.\\
	Số mol khí $\ce{CO2}$ thoát ra: $n(\ce{CO2}) = 2 \times 0{,}2 = 0{,}4\text{ mol}$.\\
	Khối lượng cốc (A) sau phản ứng:\\
	$m_A = 120 + 100 - 0{,}4 \times 44 = 202{,}4\text{ gam}$.
	\item \textbf{Xét cốc (B):}\\
	$n(\ce{AgNO3}) = \dfrac{85}{170} = 0{,}5\text{ mol}$; $n(\ce{HCl}) = \dfrac{100 \times 36{,}5\%}{36{,}5} = 1{,}0\text{ mol}$.\\
	Phương trình: $\ce{AgNO3 + HCl -> AgCl v + HNO3}$.\\
	Vì $n(\ce{AgNO3}) = 0{,}5 < n(\ce{HCl}) = 1{,}0$ nên $\ce{HCl}$ dư, $\ce{AgNO3}$ hết.\\
	Vì kết tủa $\ce{AgCl}$ nằm trong cốc và không có khí thoát ra:\\
	$m_B = 85 + 100 = 185\text{ gam}$.
	\item \textbf{So sánh hai cốc:}\\
	Vì $m_A = 202{,}4\text{ g} > m_B = 185\text{ g}$ nên cân không ở vị trí thăng bằng (cốc A nặng hơn cốc B).\\
	Để cân trở lại thăng bằng, cần thêm vào cốc (B) một lượng dung dịch $\ce{HCl}$ 36,5\% có khối lượng là:\\
	$\Delta m = 202{,}4 - 185 = \mathbf{17{,}4\text{ gam}}$.
\end{itemize}
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
\textbf{1. A} & \textbf{2. A} & \textbf{3. B} & \textbf{4. C} & \textbf{5. A} & \textbf{6. B} & \textbf{7. B} & \textbf{8. D} & \textbf{9. C} \\ \hline
\textbf{10. B} & \textbf{11. B} & \textbf{12. C} & \textbf{13. D} & \textbf{14. D} & \textbf{15. B} & \textbf{16. C} & \textbf{17. B} & \textbf{18. B} \\ \hline
\end{tabular}
\end{center}

\noindent
\textbf{2. PHẦN II (Trắc nghiệm đúng sai):}
\begin{itemize}
	\item \textbf{Câu 1:} a) Sai \quad|\quad b) Đúng \quad|\quad c) Đúng \quad|\quad d) Đúng
	\item \textbf{Câu 2:} a) Đúng \quad|\quad b) Đúng \quad|\quad c) Đúng \quad|\quad d) Sai
	\item \textbf{Câu 3:} a) Sai \quad|\quad b) Đúng \quad|\quad c) Đúng \quad|\quad d) Đúng
	\item \textbf{Câu 4:} a) Sai \quad|\quad b) Đúng \quad|\quad c) Đúng \quad|\quad d) Sai
\end{itemize}

\noindent
\textbf{3. PHẦN III (Trả lời ngắn):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\ \hline
2 & 2 & 5 & 50 & 10 & 7,41 \\ \hline
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
    hs_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De06_HocSinh.tex")
    with open(hs_path, "w", encoding="utf-8") as f:
        f.write(hs_tex)
    print("Saved:", hs_path)

    gv_tex = make_tex(is_teacher=True)
    gv_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De06_GiaoVien.tex")
    with open(gv_path, "w", encoding="utf-8") as f:
        f.write(gv_tex)
    print("Saved:", gv_path)

    # Compile with pdflatex
    for name in ["De_Kiem_Tra_Hoa_11_De06_HocSinh", "De_Kiem_Tra_Hoa_11_De06_GiaoVien"]:
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
                if file.endswith(ext) or file.startswith("ans-hoa11-de06"):
                    try:
                        os.remove(os.path.join(root, file))
                    except:
                        pass
    print("LaTeX generation completed successfully!")

if __name__ == "__main__":
    main()
