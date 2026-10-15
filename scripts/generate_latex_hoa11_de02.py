# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file LaTeX và biên dịch ra PDF cho ĐỀ SỐ 2 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De02_HocSinh.tex / .pdf
- De_Kiem_Tra_Hoa_11_De02_GiaoVien.tex / .pdf
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
    ans_file = f"ans-hoa11-de02{ver_tag}"
    
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
	\fbox{\textbf{MÃ ĐỀ: 102}}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Ba = 137).}
\smallskip

\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\Opensolutionfile{ans}[""" + ans_file + r"""]

\begin{ex}
Cân bằng hóa học liên quan đến loại phản ứng
\choice
{Không thuận nghịch.}
""" + (r"{\True Thuận nghịch.}" if is_teacher else r"{Thuận nghịch.}") + r"""
{Một chiều.}
{Oxi hóa – khử.}
\loigiai{""" + (r"""
Cân bằng hóa học là trạng thái của phản ứng thuận nghịch khi tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Yếu tố nào sau đây luôn luôn không làm dịch chuyển cân bằng của hệ phản ứng?
\choice
{Nhiệt độ.}
{Áp suất.}
{Nồng độ.}
""" + (r"{\True Chất xúc tác.}" if is_teacher else r"{Chất xúc tác.}") + r"""
\loigiai{""" + (r"""
Chất xúc tác làm tăng tốc độ phản ứng thuận và tốc độ phản ứng nghịch với số lần bằng nhau, giúp hệ nhanh đạt trạng thái cân bằng chứ không làm dịch chuyển cân bằng hóa học.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Khi phản ứng thuận nghịch ở trạng thái cân bằng thì nó
\choice
{Không xảy ra nữa.}
""" + (r"{\True Vẫn tiếp tục xảy ra.}" if is_teacher else r"{Vẫn tiếp tục xảy ra.}") + r"""
{Chỉ xảy ra theo chiều thuận.}
{Chỉ xảy ra theo chiều nghịch.}
\loigiai{""" + (r"""
Cân bằng hóa học là cân bằng động, ở trạng thái cân bằng phản ứng thuận và phản ứng nghịch vẫn tiếp tục xảy ra với tốc độ bằng nhau ($v_{\text{thuận}} = v_{\text{nghịch}}$).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
\textbf{[CTST - SGK]} Yếu tố nào sau đây luôn luôn không làm dịch chuyển cân bằng của hệ phản ứng?
\choice
{Nhiệt độ}
{Áp suất}
{Nồng độ}
""" + (r"{\True Chất xúc tác}" if is_teacher else r"{Chất xúc tác}") + r"""
\loigiai{""" + (r"""
Chất xúc tác không làm biến đổi nồng độ các chất ở trạng thái cân bằng, do đó không làm chuyển dịch cân bằng.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Chất nào sau là acid theo thuyết Brønsted -- Lowry?
\choice
{$\ce{NaNO3}$}
""" + (r"{\True $\ce{HBrO3}$}" if is_teacher else r"{$\ce{HBrO3}$}") + r"""
{$\ce{K2SO4}$}
{$\ce{CsOH}$}
\loigiai{""" + (r"""
Theo thuyết Brønsted -- Lowry, acid là chất có khả năng nhường proton $\ce{H+}$. $\ce{HBrO3}$ phân li nhường $\ce{H+}$ trong nước: $\ce{HBrO3 -> H+ + BrO3-}$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phản ứng nào dưới đây không phải là phản ứng acid-base?
\choice
{$\ce{HCl + KOH -> KCl + H2O}$}
{$\ce{H2SO4 + CaO -> CaSO4 + H2O}$}
""" + (r"{\True $\ce{H2SO4 + BaCl2 -> BaSO4 v + 2HCl}$}" if is_teacher else r"{$\ce{H2SO4 + BaCl2 -> BaSO4 v + 2HCl}$}") + r"""
{$\ce{2HNO3 + Cu(OH)2 -> Cu(NO3)2 + 2H2O}$}
\loigiai{""" + (r"""
Phản ứng $\ce{H2SO4 + BaCl2 -> BaSO4 v + 2HCl}$ là phản ứng trao đổi ion tạo kết tủa, không có sự cho nhận proton $\ce{H+}$ giữa acid và base. Các phản ứng còn lại đều là phản ứng acid - base.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Các dung dịch sau đây có cùng nồng độ $0{,}10\text{ mol/L}$, dung dịch nào dẫn điện kém nhất?
\choice
{$\ce{HCl}$}
""" + (r"{\True $\ce{HF}$}" if is_teacher else r"{$\ce{HF}$}") + r"""
{$\ce{HI}$}
{$\ce{HBr}$}
\loigiai{""" + (r"""
$\ce{HCl, HBr, HI}$ là các hydrohalic acid mạnh, điện li hoàn toàn trong nước. $\ce{HF}$ là acid yếu, điện li một phần nhỏ thành ion nên có mật độ ion tự do nhỏ nhất $\rightarrow$ dẫn điện kém nhất.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hòa tan một acid vào nước ở $25^\circ\text{C}$, kết quả là:
\choice
{$[\ce{H+}] < [\ce{OH-}]$}
{$[\ce{H+}] = [\ce{OH-}]$}
""" + (r"{\True $[\ce{H+}] > [\ce{OH-}]$}" if is_teacher else r"{$[\ce{H+}] > [\ce{OH-}]$}") + r"""
{$[\ce{H+}][\ce{OH-}] > 1{,}0 \cdot 10^{-14}$}
\loigiai{""" + (r"""
Khi hòa tan acid vào nước, $[\ce{H+}] > 1{,}0 \cdot 10^{-7}\text{ M} > [\ce{OH-}]$. Tích số ion của nước ở $25^\circ\text{C}$ luôn là hằng số: $[\ce{H+}][\ce{OH-}] = 1{,}0 \cdot 10^{-14}$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Chất nào sau đây không dẫn điện được?
\choice
""" + (r"{\True $\ce{KCl}$ rắn, khan}" if is_teacher else r"{$\ce{KCl}$ rắn, khan}") + r"""
{$\ce{CaCl2}$ nóng chảy}
{$\ce{NaOH}$ nóng chảy}
{$\ce{HBr}$ hòa tan trong nước.}
\loigiai{""" + (r"""
$\ce{KCl}$ rắn, khan ở dạng mạng tinh thể ion liên kết chặt chẽ, các ion $\ce{K+}$ và $\ce{Cl-}$ không di chuyển tự do nên không dẫn điện. $\ce{CaCl2}$ nóng chảy, $\ce{NaOH}$ nóng chảy và dung dịch $\ce{HBr}$ đều có các ion chuyển động tự do nên dẫn điện.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
\textbf{(SBT - KNTT)} Cho phản ứng hoá học sau: $\ce{N2O4(g) <=> 2NO2(g)} \quad K_C = 4{,}84 \cdot 10^{-3}$.\\
Phương án nào sau đây là nồng độ của các chất tại thời điểm cân bằng?
\choice
{$[\ce{N2O4(g)}] = 4{,}84 \cdot 10^{-1}\text{ M}; [\ce{NO2(g)}] = 1{,}0 \cdot 10^{-4}\text{ M}$.}
{$[\ce{N2O4(g)}] = 1{,}0 \cdot 10^{-1}\text{ M}; [\ce{NO2(g)}] = 4{,}84 \cdot 10^{-4}\text{ M}$.}
""" + (r"{\True $[\ce{N2O4(g)}] = 1{,}0 \cdot 10^{-1}\text{ M}; [\ce{NO2(g)}] = 2{,}20 \cdot 10^{-2}\text{ M}$.}" if is_teacher else r"{$[\ce{N2O4(g)}] = 1{,}0 \cdot 10^{-1}\text{ M}; [\ce{NO2(g)}] = 2{,}20 \cdot 10^{-2}\text{ M}$.}") + r"""
{$[\ce{N2O4(g)}] = 5{,}0 \cdot 10^{-2}\text{ M}; [\ce{NO2(g)}] = 1{,}10 \cdot 10^{-2}\text{ M}$.}
\loigiai{""" + (r"""
Biểu thức: $K_C = \dfrac{[\ce{NO2}]^2}{[\ce{N2O4}]}$.\\
Thay số với phương án C: $\dfrac{(2{,}20 \cdot 10^{-2})^2}{1{,}0 \cdot 10^{-1}} = \dfrac{4{,}84 \cdot 10^{-4}}{0{,}1} = 4{,}84 \cdot 10^{-3} = K_C$ (thỏa mãn).\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho cân bằng hoá học sau: $\ce{2SO2(g) + O2(g) <=> 2SO3(g)} \quad \Delta_r H^\circ_{298} < 0$.\\
Cho các biện pháp: (1) tăng nhiệt độ, (2) tăng áp suất chung của hệ phản ứng, (3) hạ nhiệt độ, (4) dùng thêm chất xúc tác $\ce{V2O5}$, (5) giảm nồng độ $\ce{SO3}$, (6) giảm áp suất chung của hệ phản ứng. Có mấy biện pháp làm cân bằng chuyển dịch theo chiều thuận?
\choice
{2.}
""" + (r"{\True 3.}" if is_teacher else r"{3.}") + r"""
{4.}
{5.}
\loigiai{""" + (r"""
Phản ứng thuận tỏa nhiệt ($\Delta_r H^\circ < 0$) và giảm số mol khí (từ 3 mol khí xuống 2 mol khí).\\
Các biện pháp làm chuyển dịch theo chiều thuận gồm:\\
- (2) Tăng áp suất chung (chuyển dịch về phía giảm số mol khí).\\
- (3) Hạ nhiệt độ (chuyển dịch về phía phản ứng tỏa nhiệt).\\
- (5) Giảm nồng độ $\ce{SO3}$ (chuyển dịch về phía bù lại lượng $\ce{SO3}$).\\
$\rightarrow$ Có 3 biện pháp: (2), (3), (5).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho phản ứng: $\ce{2NaHCO3(s) <=> Na2CO3(s) + CO2(g) + H2O(g)} \quad \Delta_r H^\circ_{298} = 129\text{ kJ}$.\\
Phản ứng xảy ra theo chiều nghịch khi:
\choice
""" + (r"{\True Giảm nhiệt độ}" if is_teacher else r"{Giảm nhiệt độ}") + r"""
{Tăng nhiệt độ}
{Giảm áp suất}
{Tăng nhiệt độ và giảm áp suất}
\loigiai{""" + (r"""
$\Delta_r H^\circ_{298} = 129\text{ kJ} > 0 \rightarrow$ phản ứng thuận thu nhiệt, phản ứng nghịch tỏa nhiệt. Khi giảm nhiệt độ, cân bằng chuyển dịch theo chiều phản ứng tỏa nhiệt (chiều nghịch).\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hằng số cân bằng của phản ứng $\ce{N2O4(g) <=> 2NO2(g)}$ là:
\choice
""" + (r"{\True $K = \dfrac{[\ce{NO2}]^2}{[\ce{N2O4}]}$}" if is_teacher else r"{$K = \dfrac{[\ce{NO2}]^2}{[\ce{N2O4}]}$}") + r"""
{$K = \dfrac{[\ce{NO2}]}{[\ce{N2O4}]}$}
{$K = \dfrac{[\ce{NO2}]}{[\ce{N2O4}]^2}$}
{Kết quả khác.}
\loigiai{""" + (r"""
Hằng số cân bằng nồng độ $K_C = \dfrac{[\ce{NO2}]^2}{[\ce{N2O4}]}$.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Xét cân bằng: $\ce{Fe2O3(s) + 3CO(g) <=> 2Fe(s) + 3CO2(g)}$. Biểu thức hằng số cân bằng của hệ là:
\choice
{$K = \dfrac{[\ce{Fe}]^2 [\ce{CO2}]^3}{[\ce{Fe2O3}] [\ce{CO}]^3}$}
{$K = \dfrac{[\ce{Fe2O3}] [\ce{CO}]^3}{[\ce{Fe}]^2 [\ce{CO2}]^3}$}
{$K = \dfrac{[\ce{CO}]^3}{[\ce{CO2}]^3}$}
""" + (r"{\True $K = \dfrac{[\ce{CO2}]^3}{[\ce{CO}]^3}$}" if is_teacher else r"{$K = \dfrac{[\ce{CO2}]^3}{[\ce{CO}]^3}$}") + r"""
\loigiai{""" + (r"""
Nồng độ chất rắn ($\ce{Fe2O3}$ và $\ce{Fe}$) là hằng số và quy ước bằng 1 trong biểu thức hằng số cân bằng. Do đó $K_C = \dfrac{[\ce{CO2}]^3}{[\ce{CO}]^3}$.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Dịch vị dạ dày thường có pH trong khoảng từ $1{,}5 - 3{,}5$. Những người nào bị mắc bệnh viêm loét dạ dày, tá tràng thường có $\text{pH} < 1{,}5$. Để chữa căn bệnh này, người bệnh thường uống trước bữa ăn chất nào sau đây?
\begin{center}
	\includegraphics[width=6.5cm]{images_crop_hoa11_de02/cau15_ph_he_tieu_hoa.png}
\end{center}
\choice
""" + (r"{\True Dung dịch sodium hydrogen carbonate}" if is_teacher else r"{Dung dịch sodium hydrogen carbonate}") + r"""
{Nước đun sôi để nguội}
{Nước đường saccarose}
{Một ít giấm ăn}
\loigiai{""" + (r"""
Thuốc muối $\ce{NaHCO3}$ có tính kiềm yếu, giúp trung hòa bớt acid $\ce{HCl}$ dư trong dạ dày: $\ce{NaHCO3 + HCl -> NaCl + CO2 + H2O}$, nâng pH về khoảng an toàn.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Để trung hòa $20\text{ mL}$ dung dịch $\ce{HCl}\ 0{,}1\text{ M}$ cần $10\text{ mL}$ dung dịch $\ce{NaOH}$ nồng độ $x\text{ mol/L}$. Giá trị của $x$ là
\choice
{$0{,}3$}
{$0{,}4$}
""" + (r"{\True $0{,}2.$}" if is_teacher else r"{$0{,}2.$}") + r"""
{$0{,}1.$}
\loigiai{""" + (r"""
Phương trình: $\ce{HCl + NaOH -> NaCl + H2O}$.\\
Số mol: $n_{\ce{HCl}} = n_{\ce{NaOH}} \rightarrow 0{,}020 \cdot 0{,}1 = 0{,}010 \cdot x \rightarrow x = 0{,}2\text{ M}$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong dung dịch muối dichromate luôn có cân bằng:
\[\ce{Cr2O7^2- (da\ cam) + H2O <=> 2CrO4^2- (vang\ chanh) + 2H+}\]
Nếu thêm dung dịch acid $\ce{HBr}$ đặc và dư vào dung dịch $\ce{K2Cr2O7}$ thì dung dịch chuyển thành
\choice
""" + (r"{\True Màu da cam.}" if is_teacher else r"{Màu da cam.}") + r"""
{Màu vàng.}
{Màu xanh lục.}
{Không màu.}
\loigiai{""" + (r"""
Khi thêm acid $\ce{HBr}$ đặc dư, nồng độ ion $\ce{H+}$ tăng cao làm cân bằng chuyển dịch theo chiều giảm nồng độ $\ce{H+}$ (chiều nghịch), tạo thêm ion $\ce{Cr2O7^2-}$ có màu da cam.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các phát biểu sau:\\
(1) Phản ứng thuận nghịch là phản ứng xảy ra theo 2 chiều ngược nhau trong cùng điều kiện.\\
(2) Chất xúc tác có tác dụng làm tăng tốc độ phản ứng thuận và nghịch.\\
(3) Cân bằng hóa học là trạng thái mà phản ứng đã xảy ra hoàn toàn.\\
(4) Khi phản ứng thuận nghịch đạt trạng thái cân bằng hóa học, nồng độ các chất sẽ không đổi.\\
(5) Khi phản ứng thuận nghịch đạt trạng thái cân bằng hóa học, phản ứng dừng lại.\\
(6) Sự chuyển dịch cân bằng của phản ứng thuận nghịch $\ce{2NO2 <=> N2O4}$ không phụ thuộc sự thay đổi áp suất.\\
Số phát biểu sai là
\choice
{1.}
{2.}
""" + (r"{\True 3.}" if is_teacher else r"{3.}") + r"""
{4.}
\loigiai{""" + (r"""
Có 3 phát biểu sai gồm: (3), (5), (6).\\
- (3) Sai vì phản ứng thuận nghịch không bao giờ xảy ra hoàn toàn.\\
- (5) Sai vì cân bằng hóa học là cân bằng động, phản ứng vẫn diễn ra.\\
- (6) Sai vì phản ứng có biến thiên số mol khí ($2 \rightleftharpoons 1$) nên áp suất ảnh hưởng đến chuyển dịch cân bằng.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\Closesolutionfile{ans}

\bigskip
\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\begin{ex}
Cho cân bằng hoá học trong bình kín: $\ce{2SO2(g) + O2(g) <=> 2SO3(g)}$, biết phản ứng thuận là phản ứng toả nhiệt.
\choiceTF
""" + (r"{\True Khi giảm nồng độ $\ce{O2}$, cân bằng chuyển dịch theo chiều nghịch.}" if is_teacher else r"{Khi giảm nồng độ $\ce{O2}$, cân bằng chuyển dịch theo chiều nghịch.}") + r"""
""" + (r"{\True Khi giảm nhiệt độ cân bằng chuyển dịch theo chiều thuận.}" if is_teacher else r"{Khi giảm nhiệt độ cân bằng chuyển dịch theo chiều thuận.}") + r"""
""" + (r"{Khi giảm áp suất, cân bằng chuyển dịch theo chiều thuận.}" if is_teacher else r"{Khi giảm áp suất, cân bằng chuyển dịch theo chiều thuận.}") + r"""
""" + (r"{Khi giảm nồng độ $\ce{SO3}$, cân bằng sẽ chuyển dịch theo chiều nghịch.}" if is_teacher else r"{Khi giảm nồng độ $\ce{SO3}$, cân bằng sẽ chuyển dịch theo chiều nghịch.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Đúng} vì giảm nồng độ $\ce{O2}$ làm cân bằng chuyển dịch theo chiều tạo thêm $\ce{O2}$ (chiều nghịch).\\
b) \textbf{Đúng} vì phản ứng thuận tỏa nhiệt, hạ nhiệt độ cân bằng chuyển dịch theo chiều tỏa nhiệt (chiều thuận).\\
c) \textbf{Sai} vì khi giảm áp suất cân bằng chuyển dịch theo chiều tăng số mol khí (từ 2 lên 3 mol, tức chiều nghịch).\\
d) \textbf{Sai} vì giảm nồng độ $\ce{SO3}$ cân bằng chuyển dịch theo chiều tạo thêm $\ce{SO3}$ (chiều thuận).""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các cân bằng sau:\\
(1) $\ce{2NH3(g) <=> N2(g) + 3H2(g)} \quad \Delta_r H^\circ_{298} > 0$\\
(2) $\ce{2SO2(g) + O2(g) <=> 2SO3(g)} \quad \Delta_r H^\circ_{298} < 0$\\
(3) $\ce{CaCO3(r) <=> CaO(r) + CO2(g)} \quad \Delta_r H^\circ_{298} > 0$\\
(4) $\ce{H2(g) + I2(g) <=> 2HI(g)} \quad \Delta_r H^\circ_{298} < 0$\\
Trong các cân bằng trên cân bằng nào sẽ chuyển dịch theo chiều thuận khi tăng nhiệt độ và giảm áp suất:
\choiceTF
""" + (r"{phương trình 1,2 sẽ chuyển dịch theo chiều thuận khi tăng nhiệt độ và giảm áp suất.}" if is_teacher else r"{phương trình 1,2 sẽ chuyển dịch theo chiều thuận khi tăng nhiệt độ và giảm áp suất.}") + r"""
""" + (r"{\True phương trình 1,3 sẽ chuyển dịch theo chiều thuận khi tăng nhiệt độ và giảm áp suất.}" if is_teacher else r"{phương trình 1,3 sẽ chuyển dịch theo chiều thuận khi tăng nhiệt độ và giảm áp suất.}") + r"""
""" + (r"{phương trình 1,2,3 cân bằng dịch theo chiều thuận khi giảm áp suất.}" if is_teacher else r"{phương trình 1,2,3 cân bằng dịch theo chiều thuận khi giảm áp suất.}") + r"""
""" + (r"{\True phương trình 4 cân bằng không chuyển dịch khi tăng áp suất.}" if is_teacher else r"{phương trình 4 cân bằng không chuyển dịch khi tăng áp suất.}") + r"""
\loigiai{""" + (r"""
- Tăng nhiệt độ làm cân bằng chuyển dịch theo chiều thu nhiệt ($\Delta_r H^\circ > 0$), gồm (1) và (3).\\
- Giảm áp suất làm cân bằng chuyển dịch theo chiều tăng số mol khí: (1) tăng từ 2 lên 4 mol; (3) tăng từ 0 lên 1 mol khí.\\
$\rightarrow$ Cả hai điều kiện làm (1) và (3) chuyển dịch theo chiều thuận. Ý a sai, ý b đúng.\\
c) \textbf{Sai} vì (2) giảm số mol khí nên giảm áp suất chuyển dịch theo chiều nghịch.\\
d) \textbf{Đúng} vì (4) có tổng số mol khí 2 vế bằng nhau ($1+1=2$) nên áp suất không làm chuyển dịch cân bằng.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Thực hiện thí nghiệm với 3 chất tan khác nhau là ethanol, potassium chloride và acetic acid như hình dưới đây:
\begin{center}
	\includegraphics[width=11cm]{images_crop_hoa11_de02/cau3_p2_thi_nghiem_dan_dien.png}
\end{center}
\choiceTF
""" + (r"{\True Dung dịch $\ce{KCl}$ và acetic acid có khả năng dẫn điện.}" if is_teacher else r"{Dung dịch $\ce{KCl}$ và acetic acid có khả năng dẫn điện.}") + r"""
""" + (r"{\True Dung dịch hòa tan ethanol không dẫn điện.}" if is_teacher else r"{Dung dịch hòa tan ethanol không dẫn điện.}") + r"""
""" + (r"{Các phân tử acetic acid có 1 đầu mang điện dương và một đầu mang điện âm, chuyển động trong dung dịch tạo thành dòng điện.}" if is_teacher else r"{Các phân tử acetic acid có 1 đầu mang điện dương và một đầu mang điện âm, chuyển động trong dung dịch tạo thành dòng điện.}") + r"""
""" + (r"{Ethanol là chất điện li.}" if is_teacher else r"{Ethanol là chất điện li.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Đúng} vì $\ce{KCl}$ là chất điện li mạnh (đèn sáng mạnh), $\ce{CH3COOH}$ là chất điện li yếu (đèn sáng mờ).\\
b) \textbf{Đúng} vì ethanol là chất không điện li, không có ion tự do nên đèn không sáng.\\
c) \textbf{Sai} vì dòng điện trong dung dịch là dòng chuyển dịch có hướng của các ion tự do ($\ce{CH3COO-}$ và $\ce{H+}$), không phải phân tử lưỡng cực.\\
d) \textbf{Sai} vì ethanol là chất không điện li.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong lịch sử các thuyết về định nghĩa acid -- base có thuyết Arrhenius và thuyết Brønsted--Lowry:\\
-- Thuyết Arrhenius: ``Acid là chất phân ly trong dung dịch nước tạo thành $\ce{H+}$ (proton) còn base là chất phân ly trong dung dịch nước tạo thành $\ce{OH-}$''\\
-- Thuyết Brønsted--Lowry: ``Acid là chất cho proton ($\ce{H+}$) và base là chất nhận proton''
\choiceTF
""" + (r"{Theo thuyết Brønsted--Lowry phản ứng $\ce{HClO4 + CH3COOH <=> ClO4- + CH3COOH2+}$, trong phản ứng thuận thì $\ce{CH3COOH}$ là base, trong phản ứng nghịch thì $\ce{ClO4-}$ là acid.}" if is_teacher else r"{Theo thuyết Brønsted--Lowry phản ứng $\ce{HClO4 + CH3COOH <=> ClO4- + CH3COOH2+}$, trong phản ứng thuận thì $\ce{CH3COOH}$ là base, trong phản ứng nghịch thì $\ce{ClO4-}$ là acid.}") + r"""
""" + (r"{\True Thuyết Brønsted--Lowry tổng quát hơn thuyết Arrhenius, vì phân tử không có nhóm --OH như $\ce{NH3}$, $\ce{SO3^2-}$, $\ce{CO3^2-}$ cũng là base.}" if is_teacher else r"{Thuyết Brønsted--Lowry tổng quát hơn thuyết Arrhenius, vì phân tử không có nhóm --OH như $\ce{NH3}$, $\ce{SO3^2-}$, $\ce{CO3^2-}$ cũng là base.}") + r"""
""" + (r"{Theo thuyết Brønsted--Lowry, các ion $\ce{Fe^3+}$, $\ce{Cu^2+}$, $\ce{Mg^2+}$, $\ce{Na+}$ được xem là acid.}" if is_teacher else r"{Theo thuyết Brønsted--Lowry, các ion $\ce{Fe^3+}$, $\ce{Cu^2+}$, $\ce{Mg^2+}$, $\ce{Na+}$ được xem là acid.}") + r"""
""" + (r"{\True Theo thuyết Arrhenius thì acid trong phân tử phải có nguyên tử H và base trong phân tử phải có nhóm --OH. Thuyết lí luận này đúng trong trường hợp dung môi sử dụng là nước.}" if is_teacher else r"{Theo thuyết Arrhenius thì acid trong phân tử phải có nguyên tử H và base trong phân tử phải có nhóm --OH. Thuyết lí luận này đúng trong trường hợp dung môi sử dụng là nước.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Sai} vì trong phản ứng nghịch, $\ce{ClO4-}$ nhận proton nên là base; $\ce{CH3COOH2+}$ nhường proton nên là acid.\\
b) \textbf{Đúng} vì mở rộng khái niệm base cho các chất không có nhóm --OH nhưng có khả năng nhận proton.\\
c) \textbf{Sai} vì ion $\ce{Na+}$ không phân li nhường proton trong nước.\\
d) \textbf{Đúng} theo phạm vi của thuyết Arrhenius.""" if is_teacher else "") + r"""}
\end{ex}

\bigskip
\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}

\begin{ex}
Cho dãy các chất: $\ce{KAl(SO4)2.12H2O}$, $\ce{C2H5OH}$, $\ce{C12H22O11}$ (saccharose), $\ce{CH3COOH}$, $\ce{Ca(OH)2}$, $\ce{CH3COONH4}$. Số chất điện li là
""" + (r"""
\loigiai{
Có \textbf{4} chất điện li gồm:\\
- Muối: $\ce{KAl(SO4)2.12H2O}$, $\ce{CH3COONH4}$.\\
- Acid: $\ce{CH3COOH}$.\\
- Base: $\ce{Ca(OH)2}$.\\
$\ce{C2H5OH}$ và $\ce{C12H22O11}$ là chất không điện li.\\
\textbf{Đáp số: 4.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
Cho các phản ứng:\\
(1) $\ce{NaOH + HCl -> NaCl + H2O}$\\
(2) $\ce{H2 + I2 <=> 2HI}$\\
(3) $\ce{CaCO3 <=> CaO + CO2}$\\
(4) $\ce{2KClO3 -> 2KCl + 3O2}$\\
Số phản ứng thuận nghịch là:
""" + (r"""
\loigiai{
Có \textbf{2} phản ứng thuận nghịch là (2) và (3). Phản ứng (1) và (4) là phản ứng một chiều.\\
\textbf{Đáp số: 2.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
\textbf{(SBT - CTST)} Trộn lẫn $V\text{ mL}$ dung dịch $\ce{NaOH}\ 0{,}01\text{ M}$ với $V\text{ mL}$ dung dịch $\ce{HCl}\ 0{,}03\text{ M}$ thu được $2V\text{ mL}$ dung dịch Y. Tính pH của dung dịch Y.
""" + (r"""
\loigiai{
$n_{\ce{OH-}} = 0{,}01V\text{ mol}$; $n_{\ce{H+}} = 0{,}03V\text{ mol}$.\\
$\ce{H+ + OH- -> H2O} \rightarrow n_{\ce{H+}\text{ dư}} = 0{,}03V - 0{,}01V = 0{,}02V\text{ mol}$.\\
$[\ce{H+}] = \dfrac{0{,}02V}{2V} = 0{,}01\text{ M} = 10^{-2}\text{ M} \rightarrow \text{pH} = 2$.\\
\textbf{Đáp số: 2.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
Tính pH của dung dịch thu được sau khi trộn $40\text{ mL}$ dung dịch $\ce{HCl}\ 0{,}5\text{ M}$ với $60\text{ mL}$ dung dịch $\ce{NaOH}\ 0{,}5\text{ M}$.
""" + (r"""
\loigiai{
$n_{\ce{HCl}} = 0{,}040 \cdot 0{,}5 = 0{,}02\text{ mol}$; $n_{\ce{NaOH}} = 0{,}060 \cdot 0{,}5 = 0{,}03\text{ mol}$.\\
$V_{\text{sau}} = 40 + 60 = 100\text{ mL} = 0{,}1\text{ L}$.\\
$n_{\ce{NaOH}\text{ dư}} = 0{,}03 - 0{,}02 = 0{,}01\text{ mol} \rightarrow [\ce{OH-}] = \dfrac{0{,}01}{0{,}1} = 0{,}1\text{ M}$.\\
$\text{pOH} = 1 \rightarrow \text{pH} = 14 - 1 = 13$.\\
\textbf{Đáp số: 13.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
\textbf{(SBT - KNTT)} Ascorbic acid (vitamin C) là một acid hữu cơ được kí hiệu đơn giản là $\ce{HAsc}$, phân tử khối là 176. Một học sinh hoà tan $5{,}0\text{ g}$ ascorbic acid vào $250\text{ mL}$ nước. Tính pH của dung dịch thu được, biết trong dung dịch có cân bằng sau: $\ce{HAsc <=> H+ + Asc-}; K_a = 8 \cdot 10^{-5}$.
""" + (r"""
\loigiai{
$C_0 = \dfrac{5{,}0}{176 \cdot 0{,}25} \approx 0{,}114\text{ M}$.\\
Cân bằng: $[\ce{H+}] = x, [\ce{Asc-}] = x, [\ce{HAsc}] = 0{,}114 - x$.\\
$K_a = \dfrac{x^2}{0{,}114 - x} = 8 \cdot 10^{-5} \rightarrow x \approx 2{,}98 \cdot 10^{-3}\text{ M} \rightarrow \text{pH} = -\log(2{,}98 \cdot 10^{-3}) \approx 2{,}5$.\\
\textbf{Đáp số: 2,5.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
\textbf{(SBT - KNTT)} Ethanol và propanoic acid phản ứng với nhau tạo thành ethyl propanoate theo phản ứng hoá học sau:
\[\ce{C2H5OH(l) + C2H5COOH(l) <=> C2H5COOC2H5(l) + H2O(l)}\]
Ở $50^\circ\text{C}$, giá trị $K_C$ của phản ứng trên là $7{,}5$. Nếu cho $23{,}0\text{ g}$ ethanol phản ứng với $37{,}0\text{ g}$ propanoic acid ở $50^\circ\text{C}$ thì khối lượng của ethyl propanoate thu được trong hỗn hợp ở trạng thái cân bằng là bao nhiêu gam? (Làm tròn đến hàng phần trăm).
""" + (r"""
\loigiai{
$n_{\text{ethanol}} = \dfrac{23{,}0}{46} = 0{,}5\text{ mol}$; $n_{\text{propanoic acid}} = \dfrac{37{,}0}{74} = 0{,}5\text{ mol}$.\\
Gọi số mol ester tạo thành ở cân bằng là $x\text{ (mol)}$.\\
$K_C = \dfrac{x^2}{(0{,}5 - x)^2} = 7{,}5 \rightarrow \dfrac{x}{0{,}5 - x} = \sqrt{7{,}5} \approx 2{,}7386 \rightarrow x \approx 0{,}366\text{ mol}$.\\
$m_{\text{ester}} = 0{,}366 \cdot 102 \approx 37{,}32\text{ g}$.\\
\textbf{Đáp số: 37,32.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\bigskip
\noindent
\textbf{PHẦN IV: Tự luận}\\
\textit{Học sinh giải chi tiết các câu hỏi sau vào giấy làm bài.}

\begin{ex}
\textbf{[CD - SBT]} Quan sát Hình 1.1 và ghép mỗi đối tượng ở cột A với một mô tả thích hợp ở cột B:
\begin{itemize}
	\item \textbf{Cột A:} a) Đường (a); \quad b) $t_1$; \quad c) Đường (b); \quad d) $t_2$.
	\item \textbf{Cột B:} 1. không phải là thời điểm bắt đầu trạng thái cân bằng; \quad 2. mô tả biến thiên nồng độ chất sản phẩm theo thời gian; \quad 3. Là thời điểm phản ứng đạt trạng thái cân bằng; \quad 4. Mô tả biến thiên nồng độ chất phản ứng theo thời gian.
\end{itemize}
\begin{center}
	\includegraphics[width=7.5cm]{images_crop_hoa11_de02/cau1_p4_do_thi_nong_do.png}
\end{center}
""" + (r"""
\loigiai{
Ghép nối chi tiết:\\
- \textbf{a -- 4:} Đường (a) nồng độ giảm dần từ cực đại đến khi không đổi $\rightarrow$ mô tả biến thiên nồng độ chất phản ứng.\\
- \textbf{b -- 1:} Tại $t_1$ nồng độ chất tham gia bằng chất sản phẩm, không phải thời điểm bắt đầu cân bằng.\\
- \textbf{c -- 2:} Đường (b) nồng độ tăng dần từ 0 đến khi không đổi $\rightarrow$ mô tả biến thiên nồng độ chất sản phẩm.\\
- \textbf{d -- 3:} Tại $t_2$ nồng độ các chất bắt đầu không đổi $\rightarrow$ là thời điểm đạt trạng thái cân bằng.\\
$\rightarrow$ \textbf{Kết quả ghép: a -- 4, b -- 1, c -- 2, d -- 3.}
}""" if is_teacher else r"""
\vspace{2.5cm}
""") + r"""
\end{ex}

\begin{ex}
\textbf{[CD - SGK]} Trong môi trường acid, diệp lục có màu vàng đến đỏ; còn trong môi trường kiềm, diệp lục có màu xanh.\\
a) Giải thích vì sao khi vắt chanh vào nước luộc rau muống thì màu xanh của nước lại bị nhạt đi.\\
b*) Vì sao khi luộc bánh chưng, cho thêm một chút thuốc muối ($\ce{NaHCO3}$) sẽ làm lá dong gói bánh có màu xanh đẹp hơn?
\begin{center}
	\includegraphics[width=10.5cm]{images_crop_hoa11_de02/cau2_p4_chanh_rau_muong_banh_chung.png}
\end{center}
""" + (r"""
\loigiai{
a) Trong nước luộc rau muống có chứa một lượng chất diệp lục tan vào nước. Khi vắt chanh (chứa citric acid là acid hữu cơ yếu), nồng độ $\ce{H+}$ tăng làm môi trường chuyển sang acid, ion $\ce{H+}$ thế chỗ ion $\ce{Mg^2+}$ trong nhân porphyrin của diệp lục tạo thành pheophytin có màu vàng nhạt, làm màu xanh của nước rau muống bị nhạt đi.\\
b) Trong nước, muối $\ce{NaHCO3}$ bị thuỷ phân tạo môi trường kiềm yếu:
\[\ce{NaHCO3 -> Na+ + HCO3-}\]
\[\ce{HCO3- + H2O <=> H2CO3 + OH-}\]
Môi trường kiềm này bảo vệ nhân $\ce{Mg^2+}$ của diệp lục không bị acid hóa và biến tính ở nhiệt độ cao, giúp lá dong gói bánh chưng giữ được màu xanh tươi đẹp mắt.
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
\textbf{1. B} & \textbf{2. D} & \textbf{3. B} & \textbf{4. D} & \textbf{5. B} & \textbf{6. C} & \textbf{7. B} & \textbf{8. C} & \textbf{9. A} \\ \hline
\textbf{10. C} & \textbf{11. B} & \textbf{12. A} & \textbf{13. A} & \textbf{14. D} & \textbf{15. A} & \textbf{16. C} & \textbf{17. A} & \textbf{18. C} \\ \hline
\end{tabular}
\end{center}

\noindent
\textbf{2. PHẦN II (Trắc nghiệm đúng sai):}
\begin{itemize}
	\item \textbf{Câu 1:} a) Đúng \quad|\quad b) Đúng \quad|\quad c) Sai \quad|\quad d) Sai
	\item \textbf{Câu 2:} a) Sai \quad|\quad b) Đúng \quad|\quad c) Sai \quad|\quad d) Đúng
	\item \textbf{Câu 3:} a) Đúng \quad|\quad b) Đúng \quad|\quad c) Sai \quad|\quad d) Sai
	\item \textbf{Câu 4:} a) Sai \quad|\quad b) Đúng \quad|\quad c) Sai \quad|\quad d) Đúng
\end{itemize}

\noindent
\textbf{3. PHẦN III (Trả lời ngắn):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\ \hline
4 & 2 & 2 & 13 & 2,5 & 37,32 \\ \hline
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
    hs_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De02_HocSinh.tex")
    with open(hs_path, "w", encoding="utf-8") as f:
        f.write(hs_tex)
    print("Saved:", hs_path)

    gv_tex = make_tex(is_teacher=True)
    gv_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De02_GiaoVien.tex")
    with open(gv_path, "w", encoding="utf-8") as f:
        f.write(gv_tex)
    print("Saved:", gv_path)

    # Compile with pdflatex
    for name in ["De_Kiem_Tra_Hoa_11_De02_HocSinh", "De_Kiem_Tra_Hoa_11_De02_GiaoVien"]:
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
                if file.endswith(ext) or file.startswith("ans-hoa11-de02"):
                    try:
                        os.remove(os.path.join(root, file))
                    except:
                        pass
    print("LaTeX generation completed successfully!")

if __name__ == "__main__":
    main()
