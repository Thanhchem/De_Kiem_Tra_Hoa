# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file LaTeX và biên dịch ra PDF cho ĐỀ SỐ 3 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De03_HocSinh.tex / .pdf
- De_Kiem_Tra_Hoa_11_De03_GiaoVien.tex / .pdf
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
    ans_file = f"ans-hoa11-de03{ver_tag}"
    
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
	\textbf{ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 1 (CÂN BẰNG HÓA HỌC)}\\
	\textbf{MÔN: HOÁ HỌC -- LỚP 11}\\
	\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
	\fbox{\textbf{MÃ ĐỀ: 103}}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; I = 127; Ba = 137).}
\smallskip

\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\Opensolutionfile{ans}[""" + ans_file + r"""]

\begin{ex}
Cho cân bằng hoá học: $\ce{H2(g) + I2(g) <=> 2HI(g)} \quad \Delta_r H^\circ_{298} > 0$. Cân bằng không bị chuyển dịch khi
\choice
{Tăng nhiệt độ của hệ.}
{Giảm nồng độ $\ce{HI}$.}
{Tăng nồng độ $\ce{H2}$.}
""" + (r"{\True Giảm áp suất chung của hệ.}" if is_teacher else r"{Giảm áp suất chung của hệ.}") + r"""
\loigiai{""" + (r"""
Phản ứng có tổng số mol khí ở 2 vế bằng nhau ($1 + 1 = 2$) nên sự biến đổi áp suất không làm chuyển dịch cân bằng hóa học.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho cân bằng hóa học: $\ce{N2(g) + 3H2(g) <=> 2NH3(g)}$. Phản ứng thuận là phản ứng tỏa nhiệt. Cân bằng hóa học không bị dịch chuyển khi
\choice
{Thay đổi áp suất của hệ.}
{Thay đổi nồng độ $\ce{N2}$.}
{Thay đổi nhiệt độ.}
""" + (r"{\True Thêm chất xúc tác.}" if is_teacher else r"{Thêm chất xúc tác.}") + r"""
\loigiai{""" + (r"""
Chất xúc tác làm tăng tốc độ phản ứng thuận và tốc độ phản ứng nghịch với số lần bằng nhau, giúp hệ nhanh đạt trạng thái cân bằng chứ không làm dịch chuyển cân bằng.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho cân bằng hoá học: $\ce{2SO2(g) + O2(g) <=> 2SO3(g)}$; phản ứng thuận là phản ứng tỏa nhiệt. Phát biểu đúng là:
\choice
{Cân bằng chuyển dịch theo chiều thuận khi tăng nhiệt độ.}
""" + (r"{\True Cân bằng chuyển dịch theo chiều nghịch khi giảm nồng độ $\ce{O2}$.}" if is_teacher else r"{Cân bằng chuyển dịch theo chiều nghịch khi giảm nồng độ $\ce{O2}$.}") + r"""
{Cân bằng chuyển dịch theo chiều thuận khi giảm áp suất hệ phản ứng.}
{Cân bằng chuyển dịch theo chiều nghịch khi giảm nồng độ $\ce{SO3}$.}
\loigiai{""" + (r"""
Khi giảm nồng độ $\ce{O2}$, theo nguyên lí Le Chatelier cân bằng chuyển dịch theo chiều tạo thêm $\ce{O2}$ (chiều nghịch).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Các yếu tố ảnh hưởng đến cân bằng hoá học là
\choice
{Nồng độ, nhiệt độ và chất xúc tác.}
{Nồng độ, áp suất và diện tích bề mặt.}
""" + (r"{\True Nồng độ, nhiệt độ và áp suất.}" if is_teacher else r"{Nồng độ, nhiệt độ và áp suất.}") + r"""
{Áp suất, nhiệt độ và chất xúc tác.}
\loigiai{""" + (r"""
Ba yếu tố làm chuyển dịch cân bằng hóa học là nồng độ, nhiệt độ và áp suất (đối với phản ứng có chất khí và có biến thiên số mol khí). Chất xúc tác và diện tích bề mặt không làm dịch chuyển cân bằng.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Dãy các chất gồm những chất điện li mạnh là:
\choice
""" + (r"{\True $\ce{NaF, NaOH, KCl, BaCl2}$}" if is_teacher else r"{$\ce{NaF, NaOH, KCl, BaCl2}$}") + r"""
{$\ce{HCl, NaCl, Na2CO3, Hg(CN)2}$}
{$\ce{KNO3, MgCl2, HNO3, HF}$}
{$\ce{NaOH, KCl, H2SO4, KOH, HClO}$}
\loigiai{""" + (r"""
$\ce{NaF, KCl, BaCl2}$ là các muối tan; $\ce{NaOH}$ là base mạnh $\rightarrow$ đều là chất điện li mạnh. ($\ce{HF, HClO}$ là acid yếu; $\ce{Hg(CN)2}$ là chất điện li rất yếu).\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Nhúng giấy quỳ vào dung dịch có $\text{pH} = 4$ giấy quỳ chuyển thành màu:
\choice
""" + (r"{\True Đỏ}" if is_teacher else r"{Đỏ}") + r"""
{Xanh}
{Không đổi màu}
{Chưa xác định được}
\loigiai{""" + (r"""
Dung dịch có $\text{pH} = 4 < 7$ có môi trường acid nên làm quỳ tím chuyển sang màu đỏ.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Dung dịch chất nào sau đây làm quỳ tím chuyển thành màu đỏ?
\choice
""" + (r"{\True $\ce{H2SO4}$.}" if is_teacher else r"{$\ce{H2SO4}$.}") + r"""
{$\ce{KOH}$.}
{$\ce{NaCl}$.}
{$\ce{C2H5OH}$.}
\loigiai{""" + (r"""
$\ce{H2SO4}$ là acid mạnh làm quỳ tím hóa đỏ. $\ce{KOH}$ làm quỳ hóa xanh, $\ce{NaCl}$ và $\ce{C2H5OH}$ không làm đổi màu quỳ.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Dung dịch của một base ở $25^\circ\text{C}$ có:
\choice
{$[\ce{H+}] = 1{,}0 \cdot 10^{-7}\text{ M}$}
""" + (r"{\True $[\ce{H+}] < 1{,}0 \cdot 10^{-7}\text{ M}$}" if is_teacher else r"{$[\ce{H+}] < 1{,}0 \cdot 10^{-7}\text{ M}$}") + r"""
{$[\ce{H+}] > 1{,}0 \cdot 10^{-7}\text{ M}$}
{$[\ce{H+}][\ce{OH-}] > 1{,}0 \cdot 10^{-14}$}
\loigiai{""" + (r"""
Dung dịch base có $[\ce{OH-}] > 1{,}0 \cdot 10^{-7}\text{ M} \rightarrow [\ce{H+}] = \dfrac{10^{-14}}{[\ce{OH-}]} < 1{,}0 \cdot 10^{-7}\text{ M}$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Dung dịch acid mạnh $\ce{H2SO4}\ 0{,}10\text{M}$ có:
\choice
{$\text{pH} = 1{,}00$}
""" + (r"{\True $\text{pH} < 1{,}00$}" if is_teacher else r"{$\text{pH} < 1{,}00$}") + r"""
{$\text{pH} > 1{,}00$}
{$[\ce{H+}] > 0{,}20\text{M}$}
\loigiai{""" + (r"""
$\ce{H2SO4}$ nấc 1 phân li hoàn toàn: $\ce{H2SO4 -> H+ + HSO4-}$ (tạo $0{,}10\text{ M } \ce{H+}$). Nấc 2 tiếp tục phân li một phần: $\ce{HSO4- <=> H+ + SO4^2-}$. Do đó tổng $[\ce{H+}] > 0{,}10\text{ M} \rightarrow \text{pH} = -\log[\ce{H+}] < -\log(0{,}10) = 1{,}00$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Xét các hệ cân bằng sau đây trong một bình kín:
\begin{enumerate}[(1)]
	\item $\ce{2NaHCO3(s) <=> Na2CO3(s) + H2O(g) + CO2(g)}$
	\item $\ce{CO2(g) + CaO(s) <=> CaCO3(s)}$
	\item $\ce{C(s) + CO2(g) <=> 2CO(g)}$
	\item $\ce{CO(g) + H2O(g) <=> CO2(g) + H2(g)}$
\end{enumerate}
Khi thêm $\ce{CO2}$ vào hệ thì số cân bằng chuyển dịch theo chiều thuận là
\choice
""" + (r"{\True 2.}" if is_teacher else r"{2.}") + r"""
{1.}
{4.}
{3.}
\loigiai{""" + (r"""
Khi thêm $\ce{CO2}$, cân bằng chuyển dịch theo chiều làm giảm nồng độ $\ce{CO2}$ (chiều tiêu thụ $\ce{CO2}$).\\
- Phản ứng (2) và (3) có $\ce{CO2}$ ở vế trái (chất tham gia) nên chuyển dịch theo chiều thuận.\\
- Phản ứng (1) và (4) có $\ce{CO2}$ ở vế phải (chất sản phẩm) nên chuyển dịch theo chiều nghịch.\\
$\rightarrow$ Có 2 cân bằng chuyển dịch theo chiều thuận.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho cân bằng hoá học sau: $\ce{2NH3(g) <=> N2(g) + 3H2(g)}$. Khi tăng nhiệt độ của hệ thì tỉ khối của hỗn hợp so với hydrogen giảm. Nhận xét nào sau đây là đúng?
\choice
""" + (r"{\True Khi tăng nhiệt độ của hệ, cân bằng chuyển dịch theo chiều thuận.}" if is_teacher else r"{Khi tăng nhiệt độ của hệ, cân bằng chuyển dịch theo chiều thuận.}") + r"""
{Phản ứng thuận là phản ứng toả nhiệt.}
{Khi tăng áp suất của hệ, cân bằng chuyển dịch theo chiều thuận.}
{Khi tăng nồng độ của $\ce{NH3}$, cân bằng chuyển dịch theo chiều nghịch.}
\loigiai{""" + (r"""
Tỉ khối so với $\ce{H2}$ giảm $\rightarrow \overline{M}_{\text{hh}}$ giảm. Khối lượng hỗn hợp không đổi ($m_{\text{hh}} = \text{const}$) nên số mol khí $n_{\text{hh}}$ tăng. Chiều tăng số mol khí là chiều thuận (từ 2 mol lên 4 mol khí). Vậy tăng nhiệt độ làm cân bằng chuyển dịch theo chiều thuận $\rightarrow$ Phản ứng thuận là phản ứng thu nhiệt ($\Delta_r H > 0$).\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Sự dịch chuyển cân bằng hoá học là sự di chuyển từ trạng thái cân bằng hoá học này
\choice
{Sang trạng thái cân bằng hóa học khác không cần có tác động của các yếu tố từ bên ngoài tác động lên cân bằng.}
{Sang trạng thái không cân bằng do tác động của các yếu tố từ bên ngoài tác động lên cân bằng.}
""" + (r"{\True Sang trạng thái cân bằng hoá học khác do tác động của các yếu tố từ bên ngoài tác động lên cân bằng.}" if is_teacher else r"{Sang trạng thái cân bằng hoá học khác do tác động của các yếu tố từ bên ngoài tác động lên cân bằng.}") + r"""
{Sang trạng thái cân bằng hoá học khác do cân bằng hóa học tác động lên các yếu tố bên ngoài.}
\loigiai{""" + (r"""
Sự chuyển dịch cân bằng hoá học là sự dịch chuyển từ trạng thái cân bằng này sang trạng thái cân bằng khác do tác động của các yếu tố từ bên ngoài tác động lên cân bằng.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
\textbf{(SBT - KNTT)} Cho phản ứng hoá học sau: $\ce{PCl3(g) + Cl2(g) <=> PCl5(g)}$ ở $T^\circ\text{C}$, nồng độ các chất ở trạng thái cân bằng như sau: $[\ce{PCl5}] = 0{,}059\text{ mol/L}$; $[\ce{PCl3}] = [\ce{Cl2}] = 0{,}035\text{ mol/L}$. Hằng số cân bằng ($K_C$) của phản ứng tại $T^\circ\text{C}$ là
\choice
{$1{,}68.$}
""" + (r"{\True $48{,}16.$}" if is_teacher else r"{$48{,}16.$}") + r"""
{$0{,}02.$}
{$16{,}95.$}
\loigiai{""" + (r"""
Biểu thức: $K_C = \dfrac{[\ce{PCl5}]}{[\ce{PCl3}][\ce{Cl2}]} = \dfrac{0{,}059}{0{,}035 \cdot 0{,}035} \approx 48{,}16$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong một bình kín có cân bằng hóa học sau: $\ce{2NO2(g) <=> N2O4(g)}$. Tỉ khối hơi của hỗn hợp khí trong bình so với $\ce{H2}$ ở nhiệt độ $T_1$ bằng $27{,}6$ và ở nhiệt độ $T_2$ bằng $34{,}5$. Biết $T_1 > T_2$. Phát biểu nào sau đây về cân bằng trên là đúng?
\choice
""" + (r"{\True Phản ứng thuận là phản ứng tỏa nhiệt.}" if is_teacher else r"{Phản ứng thuận là phản ứng tỏa nhiệt.}") + r"""
{Khi tăng nhiệt độ, áp suất chung của hệ cân bằng giảm.}
{Khi giảm nhiệt độ, áp suất chung của hệ cân bằng tăng.}
{Phản ứng nghịch là phản ứng tỏa nhiệt.}
\loigiai{""" + (r"""
$T_1 > T_2$ (nhiệt độ giảm). Khi giảm nhiệt độ, tỉ khối hỗn hợp tăng $\rightarrow \overline{M}_{\text{hh}}$ tăng $\rightarrow$ số mol khí giảm (chiều thuận từ 2 mol $\ce{NO2}$ thành 1 mol $\ce{N2O4}$). Vậy hạ nhiệt độ làm cân bằng dịch chuyển theo chiều thuận $\rightarrow$ Phản ứng thuận là phản ứng tỏa nhiệt ($\Delta_r H < 0$).\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Khi hoà tan $\ce{SO2}$ vào nước có cân bằng sau: $\ce{SO2 + H2O <=> HSO3- + H+}$. Khi cho thêm $\ce{NaOH}$ và khi cho thêm $\ce{H2SO4}$ loãng vào dung dịch trên thì cân bằng sẽ chuyển dịch tương ứng là
\choice
{Thuận và thuận.}
""" + (r"{\True Thuận và nghịch.}" if is_teacher else r"{Thuận và nghịch.}") + r"""
{Nghịch và thuận.}
{Nghịch và nghịch.}
\loigiai{""" + (r"""
- Thêm $\ce{NaOH}$ (chứa $\ce{OH-}$), $\ce{OH-}$ trung hòa $\ce{H+}$ làm nồng độ $\ce{H+}$ giảm $\rightarrow$ cân bằng chuyển dịch theo chiều thuận.\\
- Thêm $\ce{H2SO4}$ (chứa $\ce{H+}$), nồng độ $\ce{H+}$ tăng $\rightarrow$ cân bằng chuyển dịch theo chiều nghịch làm giảm $\ce{H+}$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các cân bằng sau:
\begin{enumerate}[(1)]
	\item $\ce{N2(g) + 3H2(g) <=> 2NH3(g)}$
	\item $\ce{H2(g) + I2(g) <=> 2HI(g)}$
	\item $\ce{CaCO3(s) <=> CaO(s) + CO2(g)}$
\end{enumerate}
Khi tăng áp suất chung của hệ thì cân bằng:
\choice
{(1) và (3) dịch chuyển theo chiều thuận; (2) không dịch chuyển.}
""" + (r"{\True (1) dịch chuyển theo chiều thuận; (2) không dịch chuyển; (3) dịch chuyển theo chiều nghịch.}" if is_teacher else r"{(1) dịch chuyển theo chiều thuận; (2) không dịch chuyển; (3) dịch chuyển theo chiều nghịch.}") + r"""
{(1) và (3) dịch chuyển theo chiều nghịch; (2) không dịch chuyển.}
{(1) dịch chuyển theo chiều nghịch; (2) không dịch chuyển; (3) dịch chuyển theo chiều thuận.}
\loigiai{""" + (r"""
Khi tăng áp suất chung, cân bằng chuyển dịch theo chiều làm giảm số mol khí:\\
- (1): 4 mol khí $\rightleftharpoons$ 2 mol khí $\rightarrow$ chuyển dịch theo chiều thuận.\\
- (2): 2 mol khí $\rightleftharpoons$ 2 mol khí $\rightarrow$ không chuyển dịch.\\
- (3): 0 mol khí $\rightleftharpoons$ 1 mol khí $\rightarrow$ chuyển dịch theo chiều nghịch.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trộn hai thể tích bằng nhau của dung dịch $\ce{H2SO4}$ và dung dịch $\ce{NaOH}$ có cùng nồng độ mol, pH của dung dịch sau phản ứng sẽ như thế nào?
\choice
{$\text{pH} = 7$}
{$\text{pH} > 7$}
""" + (r"{\True $\text{pH} < 7$}" if is_teacher else r"{$\text{pH} < 7$}") + r"""
{Không xác định được}
\loigiai{""" + (r"""
Phương trình: $\ce{H2SO4 + 2NaOH -> Na2SO4 + 2H2O}$. Vì trộn 2 thể tích bằng nhau cùng nồng độ mol nên $n_{\ce{H2SO4}} = n_{\ce{NaOH}}$. Theo tỉ lệ phản ứng $1:2$ thì $\ce{H2SO4}$ dư $\rightarrow$ dung dịch có môi trường acid $\rightarrow \text{pH} < 7$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho phản ứng: $\ce{H2(g) + I2(g) <=> 2HI(g)}$. Ở nhiệt độ $430^\circ\text{C}$, hằng số cân bằng $K_C$ của phản ứng trên bằng $53{,}96$. Đun nóng một bình kín dung tích không đổi 10 lít chứa $4{,}0\text{ gam } \ce{H2}$ và $406{,}4\text{ gam } \ce{I2}$. Khi hệ phản ứng đạt trạng thái cân bằng ở $430^\circ\text{C}$, nồng độ của $\ce{HI}$ là
\choice
""" + (r"{\True $0{,}275\text{M}.$}" if is_teacher else r"{$0{,}275\text{M}.$}") + r"""
{$0{,}225\text{M}.$}
{$0{,}151\text{M}.$}
{$0{,}320\text{M}.$}
\loigiai{""" + (r"""
$n_{\ce{H2}} = \dfrac{4}{2} = 2\text{ mol}$; $n_{\ce{I2}} = \dfrac{406{,}4}{254} = 1{,}6\text{ mol}$.\\
Gọi $x$ là số mol $\ce{H2}$ phản ứng. Tại cân bằng: $n_{\ce{H2}} = 2 - x$; $n_{\ce{I2}} = 1{,}6 - x$; $n_{\ce{HI}} = 2x$.\\
$K_C = \dfrac{(2x/10)^2}{\left(\dfrac{2 - x}{10}\right)\left(\dfrac{1{,}6 - x}{10}\right)} = \dfrac{4x^2}{(2 - x)(1{,}6 - x)} = 53{,}96$\\
$\rightarrow 49{,}96x^2 - 194{,}256x + 172{,}672 = 0 \rightarrow x \approx 1{,}375\text{ mol}$ (nhận vì $x < 1{,}6$).\\
$\rightarrow [\ce{HI}] = \dfrac{2x}{10} = \dfrac{2 \cdot 1{,}375}{10} = 0{,}275\text{ M}$.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\Closesolutionfile{ans}

\bigskip
\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\begin{ex}
Cho các phát biểu sau về trạng thái cân bằng hóa học:
\choiceTF
""" + (r"{\True Ở trạng thái cân bằng, tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch.}" if is_teacher else r"{Ở trạng thái cân bằng, tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch.}") + r"""
""" + (r"{Ở trạng thái cân bằng, các chất không phản ứng với nhau.}" if is_teacher else r"{Ở trạng thái cân bằng, các chất không phản ứng với nhau.}") + r"""
""" + (r"{Ở trạng thái cân bằng, nồng độ các chất sản phẩm luôn lớn hơn nồng độ các chất đầu.}" if is_teacher else r"{Ở trạng thái cân bằng, nồng độ các chất sản phẩm luôn lớn hơn nồng độ các chất đầu.}") + r"""
""" + (r"{\True Ở trạng thái cân bằng, nồng độ các chất không thay đổi.}" if is_teacher else r"{Ở trạng thái cân bằng, nồng độ các chất không thay đổi.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Đúng} theo bản chất cân bằng động.\\
b) \textbf{Sai} vì ở trạng thái cân bằng phản ứng thuận và phản ứng nghịch vẫn tiếp diễn với tốc độ bằng nhau.\\
c) \textbf{Sai} vì nồng độ các chất ở trạng thái cân bằng giữ nguyên không đổi chứ không nhất thiết chất sản phẩm phải lớn hơn chất đầu.\\
d) \textbf{Đúng} vì tốc độ sinh ra bằng tốc độ tiêu hao của mỗi chất.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Khi hoà tan khí chlorine vào nước tạo thành dung dịch có màu vàng nhạt gọi là nước chlorine. Trong nước chlorine xảy ra cân bằng hoá học sau:
\[\ce{Cl2 + H2O <=> HClO + HCl \quad (1)}\]
Acid $\ce{HClO}$ sinh ra không bền, dễ bị phân huỷ theo phản ứng:
\[\ce{HClO -> HCl + O \quad (2)}\]
\choiceTF
""" + (r"{\True Nước chlorine sẽ nhạt màu dần theo thời gian, không bảo quản được lâu.}" if is_teacher else r"{Nước chlorine sẽ nhạt màu dần theo thời gian, không bảo quản được lâu.}") + r"""
""" + (r"{\True Phản ứng phân huỷ acid $\ce{HClO}$ làm giảm nồng độ của chất này, cân bằng (1) chuyển dịch theo chiều thuận.}" if is_teacher else r"{Phản ứng phân huỷ acid $\ce{HClO}$ làm giảm nồng độ của chất này, cân bằng (1) chuyển dịch theo chiều thuận.}") + r"""
""" + (r"{Nước chlorine sẽ đậm màu dần theo thời gian, nên bảo quản được lâu.}" if is_teacher else r"{Nước chlorine sẽ đậm màu dần theo thời gian, nên bảo quản được lâu.}") + r"""
""" + (r"{Phản ứng phân huỷ acid $\ce{HClO}$ làm tăng nồng độ của chất này, cân bằng (1) chuyển dịch theo chiều nghịch.}" if is_teacher else r"{Phản ứng phân huỷ acid $\ce{HClO}$ làm tăng nồng độ của chất này, cân bằng (1) chuyển dịch theo chiều nghịch.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Đúng} vì $\ce{HClO}$ phân hủy liên tục làm tiêu hao dần $\ce{Cl2}$ tan trong nước, khiến dung dịch mất dần màu vàng lục nhạt.\\
b) \textbf{Đúng} vì $\ce{HClO}$ bị phân hủy làm giảm nồng độ sản phẩm, theo Le Chatelier cân bằng (1) chuyển dịch theo chiều thuận.\\
c) \textbf{Sai} vì nước chlorine nhạt màu dần và mất tính tẩy màu theo thời gian.\\
d) \textbf{Sai} vì phản ứng (2) làm giảm nồng độ $\ce{HClO}$.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hoa đậu biếc tên tiếng Anh là butterfly pea, là một loại hoa được trồng phổ biến tại các quốc gia Nam Á và Đông Nam Á như Ấn Độ, Thái Lan, Singapore, Việt Nam... Hoa có tên gọi như vậy vì cây đậu biếc thuộc họ đậu, sống dây leo và cánh hoa có màu xanh tím giống hình con bướm. Khi sử dụng đậu biếc làm chất tạo màu tự nhiên, người ta thường đun sôi hoa đậu biếc với nước, hoặc ngâm cánh hoa trong nước sôi từ 10 đến 15 phút, màu hoa sẽ được trích ly tạo thành màu xanh biếc. Sắc tố tạo màu đặc trưng cho đậu biếc là các hợp chất thuộc nhóm anthocyanin, một trong những chất chống oxi hóa tự nhiên. Điểm đặc biệt của nhóm anthocyanin là màu của chúng thay đổi dưới tác dụng pH của môi trường. Ở môi trường $\text{pH} < 7$ (môi trường Acid), anthocyanin chuyển sang đỏ tím, ngược lại chúng chuyển sang màu xanh sẫm khi môi trường $\text{pH} \ge 7$.
\choiceTF
""" + (r"{\True Vắt nước cốt chanh vào cốc nước hoa đậu biếc đã được ngâm nóng trong nước sôi thì ta thấy màu của cốc nước chuyển từ màu xanh biếc sang màu đỏ tím.}" if is_teacher else r"{Vắt nước cốt chanh vào cốc nước hoa đậu biếc đã được ngâm nóng trong nước sôi thì ta thấy màu của cốc nước chuyển từ màu xanh biếc sang màu đỏ tím.}") + r"""
""" + (r"{\True Cho nước rửa bát vào cốc nước hoa đậu biếc đã được ngâm nóng trong nước sôi thì ta thấy màu của cốc nước chuyển từ màu xanh biếc sang màu xanh sẫm ngay lập tức.}" if is_teacher else r"{Cho nước rửa bát vào cốc nước hoa đậu biếc đã được ngâm nóng trong nước sôi thì ta thấy màu của cốc nước chuyển từ màu xanh biếc sang màu xanh sẫm ngay lập tức.}") + r"""
""" + (r"{\True Cho nước muối sinh lí vào cốc nước hoa đậu biếc đã được ngâm nóng trong nước sôi thì ta thấy màu của cốc nước không thay đổi.}" if is_teacher else r"{Cho nước muối sinh lí vào cốc nước hoa đậu biếc đã được ngâm nóng trong nước sôi thì ta thấy màu của cốc nước không thay đổi.}") + r"""
""" + (r"{\True Hoa đậu biếc là một chất chỉ thị màu trong tự nhiên.}" if is_teacher else r"{Hoa đậu biếc là một chất chỉ thị màu trong tự nhiên.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Đúng} vì nước chanh chứa citric acid có $\text{pH} < 7$, làm anthocyanin chuyển sang màu đỏ tím.\\
b) \textbf{Đúng} vì nước rửa bát có tính base nhẹ ($\text{pH} \ge 7$), làm anthocyanin chuyển sang màu xanh sẫm.\\
c) \textbf{Đúng} vì nước muối sinh lí $\ce{NaCl}\ 0{,}9\%$ có môi trường trung tính ($\text{pH} \approx 7$) nên không làm đổi màu hoa đậu biếc.\\
d) \textbf{Đúng} vì màu sắc của nước hoa đậu biếc biến đổi nhạy bén theo pH của dung dịch.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
$\ce{NaCl}$ được gọi là chất điện li, tan vào nước tạo ra ion và thu được dung dịch dẫn điện, gọi là dung dịch chất điện li:
\begin{center}
	\includegraphics[width=12cm]{images_crop_hoa11_de03/cau4_p2_phan_li_nacl.png}
\end{center}
\choiceTF
""" + (r"{\True Các phân tử nước kéo các ion sodium và ion chloride ra xa nhau, phá vỡ liên kết ion giữa hai ion sodium và ion chloride.}" if is_teacher else r"{Các phân tử nước kéo các ion sodium và ion chloride ra xa nhau, phá vỡ liên kết ion giữa hai ion sodium và ion chloride.}") + r"""
""" + (r"{\True Sau khi các hợp chất muối được tách ra, các ion sodium và ion chloride được bao quanh bởi các phân tử nước. Lúc này, muối sẽ bị hòa tan, tạo thành dung dịch đồng nhất (dung dịch chất dẫn điện).}" if is_teacher else r"{Sau khi các hợp chất muối được tách ra, các ion sodium và ion chloride được bao quanh bởi các phân tử nước. Lúc này, muối sẽ bị hòa tan, tạo thành dung dịch đồng nhất (dung dịch chất dẫn điện).}") + r"""
""" + (r"{Đầu âm của các phân tử nước bị thu hút bởi các ion chloride tích điện dương, còn đầu dương của các phân tử nước bị thu hút bởi các ion sodium tích điện âm.}" if is_teacher else r"{Đầu âm của các phân tử nước bị thu hút bởi các ion chloride tích điện dương, còn đầu dương của các phân tử nước bị thu hút bởi các ion sodium tích điện âm.}") + r"""
""" + (r"{\True Sự điện li của $\ce{NaCl}$ trong nước có thể được biểu diễn bằng phương trình điện li như sau: $\ce{NaCl + (a+b)H2O -> Na+.aH2O + Cl-.bH2O}$.}" if is_teacher else r"{Sự điện li của $\ce{NaCl}$ trong nước có thể được biểu diễn bằng phương trình điện li như sau: $\ce{NaCl + (a+b)H2O -> Na+.aH2O + Cl-.bH2O}$.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Đúng} theo cơ chế solvate hoá của phân tử nước phân cực.\\
b) \textbf{Đúng} vì các ion hydrate hoá phân tán đồng đều tạo thành dung dịch điện li dẫn điện.\\
c) \textbf{Sai} vì phân tử nước có cực âm ở nguyên tử oxygen (hút ion dương $\ce{Na+}$) và cực dương ở nguyên tử hydrogen (hút ion âm $\ce{Cl-}$).\\
d) \textbf{Đúng}, đây là phương trình biểu diễn chính xác sự điện li kèm theo sự hydrate hoá ion.""" if is_teacher else "") + r"""}
\end{ex}

\bigskip
\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}

\begin{ex}
Độ pH của da thường từ $4{,}5 - 5{,}5$, có tính Acid nhẹ. Điều này giúp da của chúng ta chống lại vi khuẩn trên bề mặt. Đối với da đầu, độ pH của da đầu thường từ $4 - 7$. Sản phẩm dành cho tóc dầu sẽ có độ pH cao hơn $5{,}5$, chứa chất tẩy rửa mạnh giúp loại bỏ chất dầu nhờn trên da đầu. Bên cạnh đó, nếu bạn thường xuyên uốn nhuộm thì những dòng sản phẩm dành cho tóc hư tổn với khả năng bổ sung tính Acid giúp cân bằng độ pH trên da đầu và tóc sẽ rất hữu ích.
\begin{center}
	\includegraphics[width=6.5cm]{images_crop_hoa11_de03/cau1_p3_dau_goi.png}
\end{center}
Một loại dầu gội đầu có nồng độ ion $\ce{OH-}$ là $10^{-5{,}17}\text{ mol/L}$. Tính pH của loại dầu gội nói trên (kết quả làm tròn đến hai chữ số phần thập phân).
""" + (r"""
\loigiai{
$\text{pOH} = -\log[\ce{OH-}] = 5{,}17 \rightarrow \text{pH} = 14 - \text{pOH} = 14 - 5{,}17 = 8{,}83$.\\
\textbf{Đáp số: 8,83.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
\textbf{(SBT - KNTT)} Aspirin là một loại thuốc có thành phần chính là acetylsalicylic acid. Nếu hoà tan thuốc này vào nước, người ta xác định được pH của dung dịch tạo thành là $2{,}8$. Nồng độ $\ce{OH-}$ của dung dịch tạo thành là $10^x$. Xác định $x$?
""" + (r"""
\loigiai{
$[\ce{H+}] = 10^{-2{,}8}\text{ M} \rightarrow [\ce{OH-}] = \dfrac{10^{-14}}{10^{-2{,}8}} = 10^{-11{,}2}\text{ M} \rightarrow x = -11{,}2$.\\
\textbf{Đáp số: -11,2.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
\textbf{[CD - SGK]} Chuẩn độ dung dịch $\ce{NaOH}$ chưa biết chính xác nồng độ (biết nồng độ trong khoảng gần với $0{,}1\text{ M}$) bằng dung dịch chuẩn $\ce{HCl}\ 0{,}1\text{ M}$ với chất chỉ thị phenolphthalein hết $10\text{ mL}$. Giả sử khi kết thúc chuẩn độ, thể tích dung dịch $\ce{NaOH}$ (tính trung bình sau 3 lần chuẩn độ) đã sử dụng ở burette là $10{,}27\text{ mL}$. Nồng độ của dung dịch $\ce{NaOH}$ là $x\text{ (M)}$. Xác định giá trị của $x$ (kết quả làm tròn đến hàng phần mười).
""" + (r"""
\loigiai{
$n_{\ce{HCl}} = 0{,}010 \cdot 0{,}1 = 0{,}001\text{ mol}$.\\
$\ce{HCl + NaOH -> NaCl + H2O} \rightarrow n_{\ce{NaOH}} = 0{,}001\text{ mol}$.\\
$C_M(\ce{NaOH}) = \dfrac{0{,}001}{10{,}27 \cdot 10^{-3}} \approx 0{,}0974\text{ M} \approx 0{,}1\text{ M} \rightarrow x = 0{,}1$.\\
\textbf{Đáp số: 0,1.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
Cần bao nhiêu gam $\ce{NaOH}$ để pha chế $250{,}0\text{ mL}$ dung dịch có $\text{pH} = 12$?
""" + (r"""
\loigiai{
$\text{pH} = 12 \rightarrow \text{pOH} = 2 \rightarrow [\ce{OH-}] = 10^{-2}\text{ M} = 0{,}01\text{ M}$.\\
$n_{\ce{NaOH}} = 0{,}01 \cdot 0{,}250 = 0{,}0025\text{ mol} \rightarrow m_{\ce{NaOH}} = 0{,}0025 \cdot 40 = 0{,}1\text{ gam}$.\\
\textbf{Đáp số: 0,1.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
\textbf{[CD - SBT]} Xét phản ứng: $\ce{H2(g) + I2(g) <=> 2HI(g)}$. Một hỗn hợp phản ứng chứa trong bình dung tích $3{,}67\text{ lít}$ ở một nhiệt độ xác định, ban đầu chứa $0{,}763\text{ gam } \ce{H2}$ và $96{,}6\text{ gam } \ce{I2}$. Ở trạng thái cân bằng, bình chứa $90{,}4\text{ gam } \ce{HI}$. Tính hằng số cân bằng ($K_C$) cho phản ứng ở nhiệt độ này (làm tròn kết quả đến phần nguyên).
""" + (r"""
\loigiai{
$[\ce{H2}]_{\text{ban đầu}} = \dfrac{0{,}763}{2 \cdot 3{,}67} = 0{,}1040\text{ M}$; $[\ce{I2}]_{\text{ban đầu}} = \dfrac{96{,}6}{254 \cdot 3{,}67} = 0{,}1040\text{ M}$.\\
$[\ce{HI}]_{\text{cân bằng}} = \dfrac{90{,}4}{128 \cdot 3{,}67} = 0{,}1924\text{ M} \rightarrow 2x = 0{,}1924 \rightarrow x = 0{,}0962\text{ M}$.\\
$[\ce{H2}] = [\ce{I2}] = 0{,}1040 - 0{,}0962 = 0{,}0078\text{ M}$.\\
$K_C = \dfrac{(0{,}1924)^2}{(0{,}0078)^2} \approx 608{,}4 \approx 608$.\\
\textbf{Đáp số: 608.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
Trong công nghiệp, hydrogen được sản xuất từ phản ứng:
\[\ce{CH4(g) + H2O(g) <=> 3H2(g) + CO(g)}\]
Tính hằng số cân bằng $K_C$ của phản ứng trên ở $760^\circ\text{C}$. Biết ở nhiệt độ này, tất cả các chất đều ở thể khí và nồng độ mol của $\ce{CH4, H2O, H2}$ và $\ce{CO}$ ở trạng thái cân bằng lần lượt là $0{,}126\text{ M}$; $0{,}242\text{ M}$; $1{,}150\text{ M}$ và $0{,}126\text{ M}$ (làm tròn kết quả đến hàng phần trăm).
""" + (r"""
\loigiai{
$K_C = \dfrac{[\ce{H2}]^3 [\ce{CO}]}{[\ce{CH4}] [\ce{H2O}]} = \dfrac{1{,}150^3 \cdot 0{,}126}{0{,}126 \cdot 0{,}242} = \dfrac{1{,}150^3}{0{,}242} \approx 6{,}28$.\\
\textbf{Đáp số: 6,28.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\bigskip
\noindent
\textbf{PHẦN IV: Tự luận}\\
\textit{Học sinh giải chi tiết các câu hỏi sau vào giấy làm bài.}

\begin{ex}
\textbf{[CD - SBT]} Điền từ ngữ thích hợp vào chỗ trống trong mỗi phát biểu sau:\\
a) Quá trình phân li của các chất khi tan trong nước thành các ion được gọi là .....(1)..... Chất điện li là chất khi tan trong nước phân li thành các .....(2)..... .....(3)..... là chất khi tan trong nước không phân li thành các ion.\\
b) Theo thuyết Brønsted -- Lowry, .....(1)..... là những chất có khả năng cho $\ce{H+}$, .....(2)..... là những chất có khả năng nhận $\ce{H+}$. Acid mạnh và base mạnh phân li .....(3)..... trong nước; acid yếu và base yếu phân li .....(4)..... trong nước.
""" + (r"""
\loigiai{
Điền từ thích hợp:\\
a) (1) sự điện li; \quad (2) ion; \quad (3) Chất không điện li.\\
b) (1) acid; \quad (2) base; \quad (3) hoàn toàn; \quad (4) một phần (hoặc không hoàn toàn).
}""" if is_teacher else r"""
\vspace{2.5cm}
""") + r"""
\end{ex}

\begin{ex}
\textbf{(SBT - KNTT)}\\
a) $\ce{CH3COOH}$ (có trong giấm ăn) là một acid yếu. Tính pH của dung dịch $\ce{CH3COOH}\ 0{,}1\text{ M}$ (biết hằng số cân bằng của sự phân li $\ce{CH3COOH}$ là $1{,}8 \cdot 10^{-5}$, bỏ qua sự phân li của nước).\\
b) Trong dung dịch nước ion $\ce{CH3COO-}$ nhận proton của nước. Viết phương trình thuỷ phân và cho biết môi trường của dung dịch $\ce{CH3COONa}$.\\
c) Cho $10\text{ mL}$ dung dịch $\ce{NaOH}\ 0{,}1\text{ M}$ vào $10\text{ mL}$ dung dịch $\ce{CH3COOH}\ 0{,}2\text{ M}$ thu được $20\text{ mL}$ dung dịch A. Tính pH của dung dịch A.
""" + (r"""
\loigiai{
a) Cân bằng phân li: $\ce{CH3COOH <=> CH3COO- + H+} \quad K_a = 1{,}8 \cdot 10^{-5}$\\
$K_a = \dfrac{x^2}{0{,}1 - x} = 1{,}8 \cdot 10^{-5} \rightarrow x \approx 1{,}33 \cdot 10^{-3}\text{ M} \rightarrow \text{pH} = -\log(1{,}33 \cdot 10^{-3}) \approx 2{,}88$.\\
b) Muối $\ce{CH3COONa}$ phân li hoàn toàn: $\ce{CH3COONa -> Na+ + CH3COO-}$\\
Ion $\ce{CH3COO-}$ bị thuỷ phân:\\
\[\ce{CH3COO- + H2O <=> CH3COOH + OH-}\]
Do tạo ra ion $\ce{OH-}$ nên dung dịch $\ce{CH3COONa}$ có môi trường base (kiềm, $\text{pH} > 7$).\\
c) $n_{\ce{NaOH}} = 0{,}010 \cdot 0{,}1 = 1{,}0 \cdot 10^{-3}\text{ mol}$; $n_{\ce{CH3COOH}} = 0{,}010 \cdot 0{,}2 = 2{,}0 \cdot 10^{-3}\text{ mol}$.\\
$\ce{CH3COOH + NaOH -> CH3COONa + H2O}$\\
Sau phản ứng: $n_{\ce{CH3COOH}\text{ dư}} = 1{,}0 \cdot 10^{-3}\text{ mol}$; $n_{\ce{CH3COONa}} = 1{,}0 \cdot 10^{-3}\text{ mol}$.\\
$V_{\text{dd}} = 20\text{ mL} = 0{,}020\text{ L} \rightarrow [\ce{CH3COOH}] = [\ce{CH3COO-}] = 0{,}05\text{ M}$.\\
Cân bằng: $\ce{CH3COOH <=> CH3COO- + H+}$\\
$[\ce{H+}] = K_a \cdot \dfrac{[\ce{CH3COOH}]}{[\ce{CH3COO-}]} = 1{,}8 \cdot 10^{-5} \cdot \dfrac{0{,}05}{0{,}05} = 1{,}8 \cdot 10^{-5}\text{ M}$.\\
$\rightarrow \text{pH} = -\log(1{,}8 \cdot 10^{-5}) \approx 4{,}74$.
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
\textbf{1. D} & \textbf{2. D} & \textbf{3. B} & \textbf{4. C} & \textbf{5. A} & \textbf{6. A} & \textbf{7. A} & \textbf{8. B} & \textbf{9. B} \\ \hline
\textbf{10. A} & \textbf{11. A} & \textbf{12. C} & \textbf{13. B} & \textbf{14. A} & \textbf{15. B} & \textbf{16. B} & \textbf{17. C} & \textbf{18. A} \\ \hline
\end{tabular}
\end{center}

\noindent
\textbf{2. PHẦN II (Trắc nghiệm đúng sai):}
\begin{itemize}
	\item \textbf{Câu 1:} a) Đúng \quad|\quad b) Sai \quad|\quad c) Sai \quad|\quad d) Đúng
	\item \textbf{Câu 2:} a) Đúng \quad|\quad b) Đúng \quad|\quad c) Sai \quad|\quad d) Sai
	\item \textbf{Câu 3:} a) Đúng \quad|\quad b) Đúng \quad|\quad c) Đúng \quad|\quad d) Đúng
	\item \textbf{Câu 4:} a) Đúng \quad|\quad b) Đúng \quad|\quad c) Sai \quad|\quad d) Đúng
\end{itemize}

\noindent
\textbf{3. PHẦN III (Trả lời ngắn):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\ \hline
8,83 & -11,2 & 0,1 & 0,1 & 608 & 6,28 \\ \hline
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
    hs_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De03_HocSinh.tex")
    with open(hs_path, "w", encoding="utf-8") as f:
        f.write(hs_tex)
    print("Saved:", hs_path)

    gv_tex = make_tex(is_teacher=True)
    gv_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De03_GiaoVien.tex")
    with open(gv_path, "w", encoding="utf-8") as f:
        f.write(gv_tex)
    print("Saved:", gv_path)

    # Compile with pdflatex
    for name in ["De_Kiem_Tra_Hoa_11_De03_HocSinh", "De_Kiem_Tra_Hoa_11_De03_GiaoVien"]:
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
                if file.endswith(ext) or file.startswith("ans-hoa11-de03"):
                    try:
                        os.remove(os.path.join(root, file))
                    except:
                        pass
    print("LaTeX generation completed successfully!")

if __name__ == "__main__":
    main()
