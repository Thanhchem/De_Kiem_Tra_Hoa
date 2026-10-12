# -*- coding: utf-8 -*-
"""
Script sinh file LaTeX (.tex) và biên dịch ra PDF cho:
ĐỀ KIỂM TRA SỐ 23 - CHƯƠNG 8: KIM LOẠI CHUYỂN TIẾP THỨ NHẤT & PHỨC CHẤT (HÓA HỌC 12)
Mã đề thi: 123
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De23_HocSinh.pdf (Chuẩn chính xác 4 trang A4)
2. De_Kiem_Tra_Hoa_12_De23_GiaoVien.pdf (Bản Giáo Viên có lời giải chi tiết và đáp án)
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
import subprocess
import glob
import pymupdf

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"

def create_latex_de23(is_teacher=False):
    ans_file = "ans-de23-gv" if is_teacher else "ans-de23-hs"
    
    tex = r"""\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{geometry}
\geometry{top=0.60cm,bottom=0.60cm,left=1.20cm,right=1.20cm,headheight=26pt,headsep=3pt,footskip=13pt}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{tabularx}
\usepackage{array}
\usepackage{tikz}
\usepackage{multicol}
\usepackage{fancyhdr}
\usepackage{enumitem}

\definecolor{hfRed}{HTML}{D32F2F}
\definecolor{hfGreen}{HTML}{8BC390}
\definecolor{mainGreen}{HTML}{2E7D32}

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
\setlength{\parskip}{""" + (r"0.8pt" if is_teacher else r"0.0pt") + r"""}

\def\loigiai#1{""" + (r"""
\par\begingroup\color{blue!80!black}\small
\noindent\textbf{\textit{Lời giải chi tiết:}}\par
#1\par\endgroup""" if is_teacher else "") + r"""}

\begin{document}
\Opensolutionfile{ans}[""" + ans_file + r"""]

\noindent
\begin{minipage}[t]{0.48\textwidth}
	\textbf{SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ}\\
	\textbf{TRƯỜNG THPT CHUYÊN QUỐC HỌC}\\
	\textbf{ĐỀ CHÍNH THỨC}\\
	\textit{(Đề thi có 04 trang)}
\end{minipage}\hfill
\begin{minipage}[t]{0.48\textwidth}
	\raggedleft
	\textbf{\color{mainGreen}ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 8 -- HOÁ HỌC 12}\\
	\textbf{MÔN: HOÁ HỌC -- LỚP 12}\\
	\textit{Thời gian làm bài: 50 phút (không kể thời gian phát đề)}\\
	\textbf{Mã đề thi: 123}
\end{minipage}

"""

    if is_teacher:
        tex += r"""
\vspace{0.1cm}
\begin{center}
	\textbf{\color{hfRed}\Large (BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}
\end{center}
"""

    tex += r"""
\vspace{1pt}
\noindent\textit{Họ, tên thí sinh: \dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots Lớp: \dots\dots\dots\dots SBD: \dots\dots\dots\dots}\\
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Cr = 52; Mn = 55; Fe = 56; Co = 59; Ni = 58,7; Cu = 64; Zn = 65; Ag = 108; Ba = 137; Pt = 195).}

\vspace{1pt}
\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\begin{ex}
[KNTT - SBT] Sự hình thành các nguyên tố chuyển tiếp dãy thứ nhất là do có sự sắp xếp lần lượt các electron vào phân lớp nào sau đây?
\choice
""" + (r"{\True 3d.}" if is_teacher else r"{3d.}") + r"""
{4s.}
{4p.}
{3p.}
\loigiai{""" + (r"""
Các nguyên tố kim loại chuyển tiếp dãy thứ nhất (từ Sc đến Cu) có cấu hình electron với các electron được điền dần vào phân lớp 3d.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
[KNTT - SBT] Ion nào sau đây không có electron trên phân lớp 3d và không có màu trong dung dịch nước?
\choice
{\ce{Fe^3+}.}
{\ce{Cr^3+}.}
{\ce{Ti^3+}.}
""" + (r"{\True \ce{Sc^3+}.}" if is_teacher else r"{\ce{Sc^3+}.}") + r"""
\loigiai{""" + (r"""
Nguyên tử Sc ($Z=21$) có cấu hình $\text{[Ar]}3d^1 4s^2$. Cation $\ce{Sc^3+}$ có cấu hình $\text{[Ar]}$ (phân lớp 3d trống, không có electron) nên ion không có màu trong nước.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
[KNTT - SBT] Oxide nào sau đây có màu trắng?
\choice
{\ce{Fe2O3}.}
{\ce{Cr2O3}.}
""" + (r"{\True \ce{Al2O3}.}" if is_teacher else r"{\ce{Al2O3}.}") + r"""
{\ce{CuO}.}
\loigiai{""" + (r"""
$\ce{Al2O3}$ là oxide màu trắng. $\ce{Fe2O3}$ màu nâu đỏ, $\ce{Cr2O3}$ màu lục thẫm, $\ce{CuO}$ màu đen.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
[CTST - SBT] Trong thí nghiệm xác định hàm lượng muối Fe(II) bằng dung dịch thuốc tím (\ce{KMnO4}) trong môi trường acid, nhận xét nào sau đây là không đúng?
\choice
{Ion \ce{Fe^2+} là chất bị oxi hóa.}
{\ce{H2SO4} là chất tạo môi trường phản ứng.}
{Ion \ce{MnO4-} là chất bị khử.}
""" + (r"{\True Dung dịch muối Fe(II) có màu vàng nhạt.}" if is_teacher else r"{Dung dịch muối Fe(II) có màu vàng nhạt.}") + r"""
\loigiai{""" + (r"""
Dung dịch muối Fe(II) tinh khiết không màu hoặc có màu xanh rất nhạt. Dung dịch muối Fe(III) mới có màu vàng nhạt đến nâu đỏ.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Để trở thành phối tử trong phức chất thì phân tử hoặc anion cần phải có
\choice
{các orbital trống.}
""" + (r"{\True cặp electron hoá trị riêng.}" if is_teacher else r"{cặp electron hoá trị riêng.}") + r"""
{ít nhất 4 orbital trống.}
{ít nhất hai cặp electron hoá trị riêng.}
\loigiai{""" + (r"""
Phối tử đóng vai trò là base Lewis, cần có ít nhất một cặp electron hoá trị tự do (chưa liên kết) để cho vào orbital trống của nguyên tử/ion trung tâm tạo liên kết phối trí.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong phức chất $\ce{[Co(H2O)6]Cl3}$, số phối tử liên kết với nguyên tử trung tâm là
\choice
{2.}
{3.}
{9.}
""" + (r"{\True 6.}" if is_teacher else r"{6.}") + r"""
\loigiai{""" + (r"""
Cầu nội phức chất là ion phức $\ce{[Co(H2O)6]^3+}$, có 6 phân tử \ce{H2O} đóng vai trò là phối tử liên kết với cation trung tâm $\ce{Co^3+}$.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong ion phức $\ce{[Zn(NH3)4]^2+}$, ion phức có dạng hình học nào sau đây?
\choice
{Hình vuông phẳng.}
{Hình lập phương.}
{Hình bát diện.}
""" + (r"{\True Hình tứ diện.}" if is_teacher else r"{Hình tứ diện.}") + r"""
\loigiai{""" + (r"""
Cation $\ce{Zn^2+}$ có cấu hình $3d^{10}$ bền vững, tham gia lai hoá $sp^3$ liên kết với 4 phối tử \ce{NH3} tạo cấu trúc hình học tứ diện.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phức chất nào sau đây có cấu trúc hình học tứ diện?
\choice
{$\ce{[Pt(NH3)4]^2+}$.}
""" + (r"{\True $\ce{[Zn(NH3)4]^2+}$.}" if is_teacher else r"{$\ce{[Zn(NH3)4]^2+}$.}") + r"""
{$\ce{[Fe(H2O)6]^2+}$.}
{$\ce{[Co(H2O)6]^3+}$.}
\loigiai{""" + (r"""
$\ce{[Zn(NH3)4]^2+}$ có số phối trí 4 và có cấu trúc tứ diện đều. $\ce{[Pt(NH3)4]^2+}$ vuông phẳng. $\ce{[Fe(H2O)6]^2+}$ và $\ce{[Co(H2O)6]^3+}$ bát diện.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Chọn phát biểu sai trong các phát biểu sau:
\choice
{Phức chất có số phối trí 4, nguyên tử trung tâm lai hoá $sp^3$ thì có dạng hình học tứ diện.}
{Phức chất có số phối trí 4, nguyên tử trung tâm lai hoá $dsp^2$ thì có dạng vuông phẳng.}
{Phức chất có số phối tử là 6, nguyên tử trung tâm lai hoá $d^2sp^3$ hoặc $sp^3d^2$ thì có dạng bát diện.}
""" + (r"{\True Phức chất có số phối trí là 4 thì thường có dạng hình học vuông phẳng.}" if is_teacher else r"{Phức chất có số phối trí là 4 thì thường có dạng hình học vuông phẳng.}") + r"""
\loigiai{""" + (r"""
Phát biểu D sai vì phức chất có số phối trí 4 có dạng tứ diện rất phổ biến (đặc biệt với cation $d^{10}$), không thể khẳng định là thường có dạng vuông phẳng.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\begin{ex}
Nhỏ vài giọt dung dịch NaCl vào ống nghiệm chứa dung dịch \ce{AgNO3}, thu được kết tủa trắng AgCl. Tiếp tục nhỏ dung dịch \ce{NH3} đến dư thấy kết tủa tan, chứng tỏ phức chất $\ce{[Ag(NH3)2]+}$ đã được tạo thành. Dấu hiệu nhận biết của phản ứng tạo phức trên dựa vào
\choice
{sự ngưng tụ.}
{sự đổi màu sắc.}
""" + (r"{\True sự hoà tan kết tủa.}" if is_teacher else r"{sự hoà tan kết tủa.}") + r"""
{sự thay đổi nhiệt độ sôi.}
\loigiai{""" + (r"""
Dấu hiệu của phản ứng tạo phức là kết tủa trắng AgCl bị hoà tan trong dung dịch \ce{NH3} dư do tạo phức tan $\ce{[Ag(NH3)2]+}$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Nhỏ vài giọt dung dịch NaOH vào ống nghiệm chứa dung dịch \ce{AlCl3}, thấy xuất hiện kết tủa trắng keo, chứng tỏ phức chất $\ce{[Al(OH)3(H2O)3]}$ đã tạo thành. Dấu hiệu nhận biết của phản ứng tạo phức trên dựa vào
\choice
{sự ngưng tụ.}
""" + (r"{\True sự kết tủa.}" if is_teacher else r"{sự kết tủa.}") + r"""
{sự bay hơi.}
{sự thay đổi nhiệt độ sôi.}
\loigiai{""" + (r"""
Dấu hiệu nhận biết là sự xuất hiện kết tủa trắng keo $\ce{[Al(OH)3(H2O)3]}$ không tan.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hoàn thành phương trình phản ứng sau: $\ce{NiCl2}(s) + \text{?} \rightarrow \ce{[Ni(OH2)6]^2+}(aq) + \text{?}$
\choice
""" + (r"{\True $\ce{NiCl2(s) + 6H2O(l) -> [Ni(OH2)6]^2+(aq) + 2Cl-(aq)}$.}" if is_teacher else r"{$\ce{NiCl2(s) + 6H2O(l) -> [Ni(OH2)6]^2+(aq) + 2Cl-(aq)}$.}") + r"""
{$\ce{NiCl2(s) + 2H2O(l) -> [Ni(OH2)6]^2+(aq) + 2Cl-(aq)}$.}
{$\ce{NiCl2(s) + 6H2O(l) -> [Ni(OH2)6]^2+(s) + 2Cl-(aq)}$.}
{$\ce{NiCl2(s) + 3H2O(l) -> [Ni(OH2)6]^2+(aq) + Cl-(aq)}$.}
\loigiai{""" + (r"""
Quá trình hoà tan và hydrate hoá tạo phức hexaaquanickel(II):\\
$\ce{NiCl2(s) + 6H2O(l) -> [Ni(OH2)6]^2+(aq) + 2Cl-(aq)}$.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phức chất của $\ce{Fe^2+}$ cấu tạo nên nhân heme, thành phần của hemoglobin trong hồng cầu của máu. Trong hệ thống tuần hoàn, heme có vai trò chính trong việc vận chuyển và tích luỹ oxygen trong cơ thể. Nguyên tử/ion trung tâm trong phức chất này là
\choice
{\ce{Cl-}.}
""" + (r"{\True \ce{Fe^2+}.}" if is_teacher else r"{\ce{Fe^2+}.}") + r"""
{\ce{OH-}.}
{\ce{NH3}.}
\loigiai{""" + (r"""
Cation trung tâm trong nhân heme là ion $\ce{Fe^2+}$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"""\begin{ex}
\begin{minipage}[t]{0.68\textwidth}
Phân tử phức chất \textit{cis}-dichloro(ethylenediamine)platinum(II) có công thức cấu tạo như hình bên:\\
Có bao nhiêu loại phối tử có trong phân tử phức chất đó?
\choice
{1.}
{\True 2.}
{3.}
{4.}
\end{minipage}\hfill
\begin{minipage}[t]{0.30\textwidth}
\centering\vspace{0pt}
\includegraphics[width=\linewidth]{images_crop_de23/p1_cau14_cis_pt_en.png}
\end{minipage}
\loigiai{
Phức chất có 2 loại phối tử: phối tử một càng chloride (\ce{Cl-}) và phối tử hai càng ethylenediamine (\ce{en}: \ce{H2N-CH2-CH2-NH2}).\\
\textbf{Chọn B.}
}
\end{ex}""" if is_teacher else r"""\begin{ex}
\begin{minipage}[t]{0.68\textwidth}
Phân tử phức chất \textit{cis}-dichloro(ethylenediamine)platinum(II) có công thức cấu tạo như hình bên:\\
Có bao nhiêu loại phối tử có trong phân tử phức chất đó?
\choice
{1.}
{2.}
{3.}
{4.}
\end{minipage}\hfill
\begin{minipage}[t]{0.30\textwidth}
\centering\vspace{0pt}
\includegraphics[width=\linewidth]{images_crop_de23/p1_cau14_cis_pt_en.png}
\end{minipage}
\end{ex}""") + r"""

\begin{ex}
EDTA là một phối tử sáu càng có thể tạo phức bát diện bền với hầu hết các cation kim loại theo tỉ lệ mol phối tử : cation là $1 : 1$. Số nguyên tử nitrogen có trong phân tử phối tử EDTA tham gia tạo liên kết phối trí với ion kim loại là
\choice
""" + (r"{\True 2.}" if is_teacher else r"{2.}") + r"""
{4.}
{6.}
{8.}
\loigiai{""" + (r"""
Phân tử EDTA có 6 nguyên tử cho electron tạo liên kết phối trí gồm 2 nguyên tử N (của 2 nhóm amino) và 4 nguyên tử O (của 4 nhóm carboxylate).\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho sơ đồ chuyển hoá sau:
\par\vspace{1pt}\centerline{\includegraphics[width=0.55\textwidth]{images_crop_de23/p1_cau16_so_do.png}}\vspace{1pt}
Những phản ứng có sự tạo thành phức chất tan là
\choice
{(1) và (2).}
""" + (r"{\True (1) và (3).}" if is_teacher else r"{(1) và (3).}") + r"""
{(2) và (3).}
{Cả (1), (2) và (3).}
\loigiai{""" + (r"""
Phản ứng (1) tạo phức aqua tan $\ce{[Cu(OH2)6]^2+}$. Phản ứng (2) tạo kết tủa $\ce{[Cu(OH2)4(OH)2](s)}$. Phản ứng (3) kết tủa tan trong \ce{NH3} tạo phức tan $\ce{[Cu(NH3)4(OH2)2]^2+}$. Do đó các phản ứng tạo phức tan là (1) và (3).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho sơ đồ chuyển hoá của bạc (Ag) sau:
\par\vspace{1pt}\centerline{\includegraphics[width=0.55\textwidth]{images_crop_de23/p1_cau17_so_do.png}}\vspace{1pt}
Những phản ứng có sự tạo thành phức chất tan là
\choice
{(1), (2), (3).}
""" + (r"{\True (1), (3), (4).}" if is_teacher else r"{(1), (3), (4).}") + r"""
{(2), (3), (4).}
{(1), (2), (4).}
\loigiai{""" + (r"""
Phản ứng (1) tạo phức aqua tan $\ce{[Ag(OH2)4]+}$. Phản ứng (2) tạo kết tủa đen $\ce{Ag2O}$. Phản ứng (3) tạo phức tan $\ce{[Ag(NH3)2]+}$. Phản ứng (4) tạo phức cyano tan $\ce{[Ag(CN)2]-}$. Do đó các phản ứng tạo phức tan là (1), (3), (4).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho cấu hình electron nguyên tử của các nguyên tố: (X): $\text{[Ar]}3d^2 4s^2$; (Y): $\text{[Ar]}4s^1$; (Z): $\text{[Ar]}3d^{10} 4s^2$; (T): $\text{[Ar]}3d^{10} 4s^2 4p^3$. Trong các nhận định sau:\\
(a) Nguyên tố (X) không phải là kim loại chuyển tiếp. \quad (b) Nguyên tố (Y) là kim loại chuyển tiếp.\\
(c) Nguyên tố (T) là kim loại chuyển tiếp. \quad (d) Nguyên tố (Z) không phải là kim loại chuyển tiếp (theo định nghĩa IUPAC).\\
Số nhận định đúng là
\choice
{2.}
""" + (r"{\True 1.}" if is_teacher else r"{1.}") + r"""
{3.}
{4.}
\loigiai{""" + (r"""
• (X) có phân lớp 3d đang điền electron (Ti) $\rightarrow$ là kim loại chuyển tiếp $\rightarrow$ (a) sai.\\
• (Y) là K (kim loại kiềm nhóm IA) $\rightarrow$ (b) sai.\\
• (T) là As (phi kim nhóm VA) $\rightarrow$ (c) sai.\\
• (Z) là Zn ($3d^{10} 4s^2$), cation $\ce{Zn^2+}$ ($3d^{10}$), theo IUPAC kim loại chuyển tiếp phải có phân lớp d chưa bão hoà ở nguyên tử hoặc ion bền $\rightarrow$ Zn không phải là kim loại chuyển tiếp $\rightarrow$ (d) đúng.\\
Vậy có 1 nhận định đúng.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\setcounter{ex}{0}

\begin{ex}
[CĐ – SBT] Cho các thông tin về thế điện cực chuẩn của các cặp oxi hoá – khử sau:
\begin{center}
\renewcommand{\arraystretch}{1.1}
\begin{tabular}{|l|c|}
\hline
\textbf{Cặp oxi hoá -- khử} & \textbf{Thế điện cực chuẩn (V)} \\
\hline
$\ce{Fe^3+/Fe^2+}$ & $+0,77$ \\
\hline
$\ce{Cr2O7^2- + 14H+ / 2Cr^3+ + 7H2O}$ & $+1,33$ \\
\hline
$\ce{MnO4- + 8H+ / Mn^2+ + 4H2O}$ & $+1,51$ \\
\hline
\end{tabular}
\end{center}
Biết $\ce{Cr2O7^2-(aq)}$ có màu da cam và $\ce{Cr^3+(aq)}$ có màu xanh lá cây.\\
Xét tính đúng/sai của các phát biểu sau:
\choiceTF
{Trong môi trường acid, anion $\ce{Cr2O7^2-}$ (từ $\ce{K2Cr2O7}$) có tính oxi hoá mạnh hơn anion $\ce{MnO4-}$ (từ $\ce{KMnO4}$).}
""" + (r"{\True Chuẩn độ được $\ce{Fe^2+}$ trong dung dịch gồm $\ce{Fe^2+}$, $\ce{SO4^2-}$ và $\ce{H+}$ bằng dung dịch chất chuẩn $\ce{KMnO4}$.}" if is_teacher else r"{Chuẩn độ được $\ce{Fe^2+}$ trong dung dịch gồm $\ce{Fe^2+}$, $\ce{SO4^2-}$ và $\ce{H+}$ bằng dung dịch chất chuẩn $\ce{KMnO4}$.}") + r"""
{Không chuẩn độ được $\ce{Fe^2+}$ trong dung dịch gồm $\ce{Fe^2+}$, $\ce{SO4^2-}$ và $\ce{H+}$ bằng dung dịch chất chuẩn $\ce{K2Cr2O7}$.}
{Có diễn ra phản ứng oxi hoá – khử tự diễn biến theo phương trình:\\ $6\ce{Fe^3+}(aq) + 2\ce{Cr^3+}(aq) + 7\ce{H2O}(l) \rightarrow 6\ce{Fe^2+}(aq) + \ce{Cr2O7^2-}(aq) + 14\ce{H+}(aq)$.}
\loigiai{""" + (r"""
a) \textbf{Sai.} Do $E^\circ(\ce{MnO4-/Mn^2+}) = 1,51\text{ V} > E^\circ(\ce{Cr2O7^2-/2Cr^3+}) = 1,33\text{ V}$ nên $\ce{MnO4-}$ có tính oxi hoá mạnh hơn $\ce{Cr2O7^2-}$.\\
b) \textbf{Đúng.} $\ce{MnO4-}$ oxi hoá hoàn toàn $\ce{Fe^2+}$ lên $\ce{Fe^3+}$, kết thúc chuẩn độ dung dịch chuyển sang màu tím hồng nhạt của $\ce{MnO4-}$ dư.\\
c) \textbf{Sai.} Do $E^\circ(\ce{Cr2O7^2-/2Cr^3+}) = 1,33\text{ V} > E^\circ(\ce{Fe^3+/Fe^2+}) = 0,77\text{ V}$ nên hoàn toàn chuẩn độ được $\ce{Fe^2+}$ bằng $\ce{K2Cr2O7}$ (dùng chỉ thị diphenylamine).\\
d) \textbf{Sai.} Phản ứng tự diễn biến theo chiều ngược lại: $\ce{Cr2O7^2-}$ oxi hoá $\ce{Fe^2+}$ thành $\ce{Fe^3+}$ và sinh ra $\ce{Cr^3+}$.
""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
[CTST - SBT] Khi tìm hiểu về hợp chất phức chlorophyll (chất diệp lục), xét tính đúng/sai của các phát biểu sau:
\choiceTF
{Chlorophyll chỉ được tìm thấy trong tế bào thực vật.}
""" + (r"{\True Chlorophyll đóng vai trò hấp thụ năng lượng ánh sáng và xúc tác cho quá trình quang hợp.}" if is_teacher else r"{Chlorophyll đóng vai trò hấp thụ năng lượng ánh sáng và xúc tác cho quá trình quang hợp.}") + r"""
{Trong chất diệp lục chỉ chứa duy nhất một loại chlorophyll.}
""" + (r"{\True Chlorophyll là thành phần sắc tố quan trọng nhất trong lục lạp của lá cây.}" if is_teacher else r"{Chlorophyll là thành phần sắc tố quan trọng nhất trong lục lạp của lá cây.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Sai.} Chlorophyll còn được tìm thấy trong vi khuẩn lam, tảo và vi khuẩn quang hợp lưu huỳnh màu lục.\\
b) \textbf{Đúng.} Chlorophyll hấp thụ photon ánh sáng xúc tác phản ứng tổng hợp carbohydrate từ $\ce{CO2}$ và $\ce{H2O}$.\\
c) \textbf{Sai.} Có nhiều loại chlorophyll như chlorophyll a, b, c, d...\\
d) \textbf{Đúng.} Chlorophyll là sắc tố chính tạo màu xanh và thực hiện chức năng quang hợp.
""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
[KNTT - SBT] Nhỏ muối thiocyanate ($\ce{SCN-}$) vào dung dịch muối $\ce{Fe^3+}$ loãng, dung dịch từ màu vàng nhạt chuyển sang màu đỏ máu do một phối tử nước trong phức chất aqua bát diện của $\ce{Fe^3+}$ bị thay thế bởi một phối tử $\ce{SCN-}$:
\choiceTF
""" + (r"{\True Phức chất aqua của sắt(III) có công thức hoá học là $\ce{[Fe(H2O)6]^3+}$.}" if is_teacher else r"{Phức chất aqua của sắt(III) có công thức hoá học là $\ce{[Fe(H2O)6]^3+}$.}") + r"""
{Phức chất có màu đỏ máu là phức chất của $\ce{Fe^3+}$ chứa 1 phối tử $\ce{SCN-}$ và 6 phối tử nước.}
""" + (r"{\True Phức chất màu đỏ máu có công thức hoá học là $\ce{[Fe(H2O)5(SCN)]^2+}$.}" if is_teacher else r"{Phức chất màu đỏ máu có công thức hoá học là $\ce{[Fe(H2O)5(SCN)]^2+}$.}") + r"""
{Phức chất màu đỏ máu có điện tích là +3.}
\loigiai{""" + (r"""
a) \textbf{Đúng.} Ion $\ce{Fe^3+}$ trong dung dịch nước tồn tại ở dạng phức aqua bát diện $\ce{[Fe(H2O)6]^3+}$.\\
b) \textbf{Sai.} Khi 1 phối tử $\ce{SCN-}$ thay thế 1 phân tử $\ce{H2O}$ thì phức chỉ còn 5 phối tử nước (tổng số phối trí là 6).\\
c) \textbf{Đúng.} Công thức phức màu đỏ máu là $\ce{[Fe(H2O)5(SCN)]^2+}$.\\
d) \textbf{Sai.} Điện tích ion phức là $(+3) + (-1) = +2$.
""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hemoglobin (Hb) là một metalloprotein chiếm khoảng 35\% trọng lượng của hồng cầu, có nhiệm vụ vận chuyển oxygen từ phổi đến các mô. Cấu trúc gồm 4 tiểu đơn vị, mỗi tiểu đơn vị chứa một nhân heme (chiếm 5\% khối lượng Hb) và một chuỗi polypeptide globin. Mỗi nhân heme chứa một ion $\ce{Fe^2+}$ ở trung tâm vòng porphyrin và chỉ liên kết tối đa với 1 phân tử $\ce{O2}$. Đánh giá tính đúng/sai của các phát biểu sau:
\choiceTF
""" + (r"{\True Dạng hình học của phức chất Fe(II) trong nhân heme khi liên kết với oxygen là bát diện.}" if is_teacher else r"{Dạng hình học của phức chất Fe(II) trong nhân heme khi liên kết với oxygen là bát diện.}") + r"""
""" + (r"{\True Mỗi phân tử hemoglobin có thể kết hợp tối đa với bốn phân tử oxygen.}" if is_teacher else r"{Mỗi phân tử hemoglobin có thể kết hợp tối đa với bốn phân tử oxygen.}") + r"""
""" + (r"{\True Trong các nhân heme, nguyên tử trung tâm $\ce{Fe^2+}$ liên kết với phối tử qua các nguyên tử nitrogen.}" if is_teacher else r"{Trong các nhân heme, nguyên tử trung tâm $\ce{Fe^2+}$ liên kết với phối tử qua các nguyên tử nitrogen.}") + r"""
{Trong phân tử hemoglobin, nguyên tố iron (Fe) chiếm 5\% theo khối lượng.}
\loigiai{""" + (r"""
a) \textbf{Đúng.} $\ce{Fe^2+}$ liên kết với 4 nguyên tử N của porphyrin, 1 nguyên tử N của histidine và 1 phân tử $\ce{O2}$, có số phối trí 6 (dạng bát diện).\\
b) \textbf{Đúng.} Phân tử Hb có 4 nhân heme nên liên kết tối đa 4 phân tử $\ce{O2}$.\\
c) \textbf{Đúng.} $\ce{Fe^2+}$ phối trí với các nguyên tử N của vòng porphyrin và N của chuỗi globin.\\
d) \textbf{Sai.} 4 nhân heme chiếm 5\% khối lượng Hb, nhưng trong heme Fe chỉ là một nguyên tử trung tâm nên \% khối lượng Fe trong Hb chỉ khoảng 0,34\%.
""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}

\setcounter{ex}{0}

\begin{ex}
Số phối tử không mang điện tích trong phức chất $\ce{[Co(NH3)5Cl]Cl2}$ là bao nhiêu?
\loigiai{""" + (r"""
Trong cầu nội $\ce{[Co(NH3)5Cl]^2+}$, có 5 phân tử \ce{NH3} là phối tử trung hoà không mang điện tích.\\
\textbf{Đáp số: 5.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Trả lời:} \dotfill
\vspace{0.2cm}""" if not is_teacher else "") + r"""

\begin{ex}
Số nguyên tử kim loại ở cầu ngoại trong một phân tử phức chất $\ce{Na2[Pt(CN)4]}$ là bao nhiêu?
\loigiai{""" + (r"""
Cầu ngoại gồm 2 cation $\ce{Na+}$, tức có 2 nguyên tử natri ở cầu ngoại.\\
\textbf{Đáp số: 2.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Trả lời:} \dotfill
\vspace{0.2cm}""" if not is_teacher else "") + r"""

\begin{ex}
Trong dung dịch nước, ion $\ce{Co^2+}$ tạo thành phức chất aqua có dạng hình học bát diện $\ce{[Co(H2O)6]^2+}$. Số phối tử liên kết với nguyên tử trung tâm trong ion phức này là bao nhiêu?
\loigiai{""" + (r"""
Ion phức aqua bát diện $\ce{[Co(H2O)6]^2+}$ có 6 phối tử phân tử nước \ce{H2O}.\\
\textbf{Đáp số: 6.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Trả lời:} \dotfill
\vspace{0.2cm}""" if not is_teacher else "") + r"""

\begin{ex}
Cho các phức chất và ion phức sau: $\ce{[Co(NH3)6]^3+}$, $\ce{[Cr(H2O)6]^3+}$, $\ce{[Ni(CO)4]}$, $\ce{[BF4]-}$, $\ce{[Fe(CN)6]^4-}$, $\ce{[Ag(NH3)2]+}$. Có bao nhiêu phức anion (phức chất mang điện tích âm) trong số các chất trên?
\loigiai{""" + (r"""
Các phức anion mang điện tích âm là $\ce{[BF4]-}$ và $\ce{[Fe(CN)6]^4-}$. Tổng cộng có 2 phức anion.\\
\textbf{Đáp số: 2.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Trả lời:} \dotfill
\vspace{0.2cm}""" if not is_teacher else "") + r"""

\begin{ex}
Cho 16,0 gam bột CuO tan hết trong dung dịch $\ce{H2SO4}$ 20\% vừa đủ, sau đó làm nguội dung dịch thu được đến 10$^\circ$C thấy tách ra $m$ gam tinh thể $\ce{CuSO4.5H2O}$. Biết độ tan của $\ce{CuSO4}$ ở 10$^\circ$C là 17,4 gam trong 100 gam nước. Giá trị của $m$ là bao nhiêu? \textit{(Làm tròn kết quả đến hàng phần mười)}.
\loigiai{""" + (r"""
• $n_{\ce{CuO}} = \frac{16,0}{80} = 0,2\text{ mol}$.\\
• Phản ứng: $\ce{CuO + H2SO4 -> CuSO4 + H2O}$.\\
• $n_{\ce{H2SO4}} = 0,2\text{ mol} \rightarrow m_{\text{dd }\ce{H2SO4}} = \frac{0,2 \times 98}{0,20} = 98\text{ g}$.\\
• Khối lượng dung dịch sau phản ứng: $m_{\text{dd}} = 16,0 + 98 = 114\text{ g}$.\\
• $n_{\ce{CuSO4}} = 0,2\text{ mol} \rightarrow m_{\ce{CuSO4}} = 0,2 \times 160 = 32\text{ g}$.\\
• Gọi $n_{\ce{CuSO4.5H2O}} = x\text{ mol} \rightarrow m = 250x\text{ g}$.\\
  Khối lượng $\ce{CuSO4}$ còn lại: $32 - 160x\text{ (g)}$.\\
  Khối lượng $\ce{H2O}$ còn lại: $(114 - 32) - 90x = 82 - 90x\text{ (g)}$.\\
• Ở 10$^\circ$C, độ tan $S = 17,4\text{ g / 100 g }\ce{H2O}$:\\
  $\frac{32 - 160x}{82 - 90x} = \frac{17,4}{100} \Rightarrow 3200 - 16000x = 1426,8 - 1566x \Rightarrow 14434x = 1773,2 \Rightarrow x \approx 0,12285\text{ mol}$.\\
• Khối lượng tinh thể: $m = 250 \times 0,12285 \approx 30,71\text{ g} \rightarrow \mathbf{30,7\text{ g}}$.\\
\textbf{Đáp số: 30,7.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Trả lời:} \dotfill
\vspace{0.2cm}""" if not is_teacher else "") + r"""

\begin{ex}
Đem 9,80 gam muối Mohr $\ce{FeSO4.(NH4)2SO4.6H2O}$ hoà tan trong nước cất rồi định mức thành 250 mL dung dịch trong bình định mức với sự có mặt của $\ce{H2SO4}$ loãng. Tiến hành chuẩn độ 25,0 mL dung dịch này bằng dung dịch $\ce{KMnO4}$ chưa rõ nồng độ, thể tích dung dịch $\ce{KMnO4}$ tiêu tốn trung bình qua 3 lần chuẩn độ là 24,65 mL. Nồng độ mol/L của dung dịch $\ce{KMnO4}$ là bao nhiêu? \textit{(Làm tròn đến hai chữ số thập phân)}.
\loigiai{""" + (r"""
• Khối lượng mol muối Mohr: $M = 56 + 96 + 2 \times 18 + 96 + 6 \times 18 = 392\text{ g/mol}$.\\
• $n_{\text{muối}} = \frac{9,80}{392} = 0,025\text{ mol} = 25\text{ mmol}$ (trong 250 mL).\\
• Trong 25,0 mL dung dịch chuẩn độ có: $n_{\ce{Fe^2+}} = 2,5\text{ mmol}$.\\
• Phương trình chuẩn độ: $5\ce{Fe^2+} + \ce{MnO4-} + 8\ce{H+} \rightarrow 5\ce{Fe^3+} + \ce{Mn^2+} + 4\ce{H2O}$.\\
• $n_{\ce{KMnO4}} = \frac{2,5}{5} = 0,50\text{ mmol}$.\\
• $C_M(\ce{KMnO4}) = \frac{0,50\text{ mmol}}{24,65\text{ mL}} \approx 0,02028\text{ mol/L} \rightarrow \mathbf{0,02\text{ M}}$.\\
\textbf{Đáp số: 0,02.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Trả lời:} \dotfill
\vspace{0.2cm}""" if not is_teacher else "") + r"""

\Closesolutionfile{ans}

"""

    if is_teacher:
        tex += r"""
\newpage
\begin{center}
	\textbf{\color{hfRed}\Large BẢNG ĐÁP ÁN ĐỀ SỐ 23 (MÃ ĐỀ 123)}
\end{center}

\vspace{0.2cm}
\noindent\textbf{\large 1. PHẦN I: CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN}\\
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} \\
\hline
\textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} \\
\hline
\hline
\textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} & \textbf{17} & \textbf{18} & \multicolumn{2}{c|}{} \\
\cline{1-8}
\textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \multicolumn{2}{c|}{} \\
\cline{1-8}
\end{tabular}
\end{center}

\vspace{0.3cm}
\noindent\textbf{\large 2. PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG SAI}\\
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Lệnh hỏi a} & \textbf{Lệnh hỏi b} & \textbf{Lệnh hỏi c} & \textbf{Lệnh hỏi d} \\
\hline
\textbf{Câu 1} & \textbf{\color{hfRed}S} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}S} & \textbf{\color{hfRed}S} \\
\hline
\textbf{Câu 2} & \textbf{\color{hfRed}S} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}S} & \textbf{\color{hfRed}Đ} \\
\hline
\textbf{Câu 3} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}S} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}S} \\
\hline
\textbf{Câu 4} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}S} \\
\hline
\end{tabular}
\end{center}

\vspace{0.3cm}
\noindent\textbf{\large 3. PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN}\\
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\
\hline
\textbf{\color{hfRed}5} & \textbf{\color{hfRed}2} & \textbf{\color{hfRed}6} & \textbf{\color{hfRed}2} & \textbf{\color{hfRed}30,7} & \textbf{\color{hfRed}0,02} \\
\hline
\end{tabular}
\end{center}
"""

    tex += r"""
\end{document}
"""
    return tex

def compile_tex(tex_path):
    base_name = os.path.splitext(os.path.basename(tex_path))[0]
    cmd = ["pdflatex", "-interaction=nonstopmode", base_name + ".tex"]
    res = subprocess.run(cmd, cwd=OUT_DIR, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if res.returncode != 0:
        print(f"Lỗi biên dịch {base_name}:")
        print(res.stderr[-500:] if res.stderr else res.stdout[-500:])
        return False
    # Chạy lần 2 để cập nhật label/page
    subprocess.run(cmd, cwd=OUT_DIR, capture_output=True, text=True, encoding="utf-8", errors="replace")
    
    # Xoá file rác
    for ext in [".aux", ".log", ".out"]:
        f_junk = os.path.join(OUT_DIR, base_name + ext)
        if os.path.exists(f_junk):
            os.remove(f_junk)
    for f in glob.glob(os.path.join(OUT_DIR, "ans-*.tex")):
        os.remove(f)
    return True

def main():
    print("--- Đang tạo file LaTeX và PDF cho ĐỀ SỐ 23 ---")
    
    # 1. Học sinh
    hs_tex = create_latex_de23(is_teacher=False)
    hs_tex_path = os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_12_De23_HocSinh.tex")
    with open(hs_tex_path, "w", encoding="utf-8") as f:
        f.write(hs_tex)
    
    p_hs = 0
    ok_hs = compile_tex(hs_tex_path)
    if ok_hs:
        doc_hs = pymupdf.open(os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_12_De23_HocSinh.pdf"))
        p_hs = len(doc_hs)
        print(f"Biên dịch thành công: De_Kiem_Tra_Hoa_12_De23_HocSinh.pdf ({p_hs} trang)")
    
    # 2. Giáo viên
    p_gv = 0
    gv_tex = create_latex_de23(is_teacher=True)
    gv_tex_path = os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_12_De23_GiaoVien.tex")
    with open(gv_tex_path, "w", encoding="utf-8") as f:
        f.write(gv_tex)
        
    ok_gv = compile_tex(gv_tex_path)
    if ok_gv:
        doc_gv = pymupdf.open(os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_12_De23_GiaoVien.pdf"))
        p_gv = len(doc_gv)
        print(f"Biên dịch thành công: De_Kiem_Tra_Hoa_12_De23_GiaoVien.pdf ({p_gv} trang)")
        
    print("\n--- KẾT QUẢ BIÊN DỊCH ---")
    print(f"Học sinh: {p_hs} trang (Yêu cầu chuẩn: 4 trang)")
    print(f"Giáo viên: {p_gv} trang")
    if p_hs == 4:
        print("XÁC NHẬN: Bản Học Sinh ĐẠT CHUẨN CHÍNH XÁC 4 TRANG A4!")
    else:
        print("CẢNH BÁO: Số trang bản Học Sinh chưa đúng 4 trang!")

if __name__ == "__main__":
    main()
