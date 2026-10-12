# -*- coding: utf-8 -*-
"""
Script sinh file LaTeX (.tex) và biên dịch ra PDF cho:
ĐỀ KIỂM TRA SỐ 15 - CHƯƠNG 5: PIN ĐIỆN – ĐIỆN PHÂN (HÓA HỌC 12)
Mã đề thi: 115
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De15_HocSinh.pdf (Chuẩn chính xác 4 trang A4)
2. De_Kiem_Tra_Hoa_12_De15_GiaoVien.pdf (Bản Giáo Viên có lời giải chi tiết và đáp án)
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
import subprocess
import glob
import fitz  # PyMuPDF

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"

def create_latex_de15(is_teacher=False):
    ans_file = "ans-de15-gv" if is_teacher else "ans-de15-hs"
    
    tex = r"""\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{geometry}
\geometry{top=0.9cm,bottom=0.9cm,left=1.35cm,right=1.35cm,headheight=24pt,headsep=5pt,footskip=11pt}
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
\setlength{\parskip}{""" + (r"1.2pt" if is_teacher else r"0.0pt") + r"""}

\def\loigiai#1{""" + (r"""
\par\begingroup\color{blue!80!black}\small
\noindent\textbf{\textit{Lời giải chi tiết:}}\par
#1\par\endgroup""" if is_teacher else "") + r"""}

\begin{document}
\Opensolutionfile{ans}[""" + ans_file + r"""]

\noindent
\begin{minipage}[t]{0.48\textwidth}
	\textbf{SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ}\\
	\textbf{TRƯỜNG THPT HAI BÀ TRƯNG}\\
	\textbf{ĐỀ CHÍNH THỨC}\\
	\textit{(Đề thi có 04 trang)}
\end{minipage}\hfill
\begin{minipage}[t]{0.48\textwidth}
	\raggedleft
	\textbf{\color{mainGreen}ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 5 (PIN ĐIỆN – ĐIỆN PHÂN)}\\
	\textbf{MÔN: HOÁ HỌC -- LỚP 12}\\
	\textit{Thời gian làm bài: 50 phút (không kể thời gian phát đề)}\\
	\textbf{Mã đề thi: 115}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Mn = 55; Fe = 56; Cu = 64; Zn = 65; Ag = 108; Cd = 112; Au = 197; Pb = 207; hằng số Faraday F = 96500 C/mol).}

\vspace{2pt}
\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\setcounter{ex}{0}

\begin{ex}
(SBT – KNTT). Trong pin điện hoá Zn -- Cu, ở anode (cực âm) xảy ra quá trình
\choice
""" + (r"{\True oxi hoá Zn thành ion \ce{Zn^2+}}" if is_teacher else r"{oxi hoá Zn thành ion \ce{Zn^2+}}") + r"""
{khử ion \ce{Cu^2+} thành Cu}
{khử Cu thành ion \ce{Cu^2+}}
{oxi hoá ion \ce{Zn^2+} thành Zn}
\loigiai{""" + (r"""
Trong pin điện hoá Zn -- Cu: Cực kẽm là cực âm (anode), tại đây xảy ra quá trình oxi hoá Zn: $\ce{Zn -> Zn^2+ + 2e}$.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Pin điện là gì?
\choice
""" + (r"{\True Một thiết bị lưu trữ và cung cấp điện năng}" if is_teacher else r"{Một thiết bị lưu trữ và cung cấp điện năng}") + r"""
{Một thiết bị dùng để đo nhiệt độ}
{Một thiết bị dùng để đo áp suất}
{Một thiết bị dùng để lưu trữ năng lượng cơ học}
\loigiai{""" + (r"""
Pin điện là thiết bị chuyển hoá trực tiếp hoá năng của phản ứng hoá học thành điện năng, dùng để lưu trữ và cung cấp nguồn điện năng cho các thiết bị.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Pin Galvani còn được gọi là gì?
\choice
{Pin nhiên liệu}
""" + (r"{\True Pin điện hóa}" if is_teacher else r"{Pin điện hóa}") + r"""
{Pin nhiệt điện}
{Pin quang điện}
\loigiai{""" + (r"""
Pin Galvani (Galvanic cell) là thuật ngữ khoa học chỉ pin điện hoá cơ bản dựa trên phản ứng oxi hoá -- khử tự diễn biến.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Điện cực dương trong pin gọi là gì?
\choice
{Anode}
""" + (r"{\True Cathode}" if is_teacher else r"{Cathode}") + r"""
{Điện cực trơ}
{Cực âm}
\loigiai{""" + (r"""
Trong pin điện hoá, cực dương là cathode (nơi xảy ra sự khử), cực âm là anode (nơi xảy ra sự oxi hoá).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Trong quá trình mạ bạc cho một chiếc vòng bằng thép thì ở anode xảy ra quá trình
\choice
""" + (r"{\True \ce{Ag -> Ag+ + 1e}}" if is_teacher else r"{\ce{Ag -> Ag+ + 1e}}") + r"""
{\ce{Fe -> Fe^2+ + 2e}}
{\ce{2H2O -> 4H+ + O2 + 4e}}
{\ce{C -> C^4+ + 4e}}
\loigiai{""" + (r"""
Khi mạ bạc (phương pháp dương cực tan): Anode làm bằng thanh kim loại bạc (Ag), cathode là chiếc vòng bằng thép. Ở anode xảy ra quá trình oxi hoá thanh bạc tan ra: $\ce{Ag -> Ag+ + 1e}$.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Khi điện phân dung dịch \ce{NaCl} bằng dòng điện một chiều (với điện cực trơ, có màng ngăn xốp) thì ở cathode xảy ra quá trình
\choice
{oxi hoá \ce{H2O} thành \ce{H+} và \ce{O2}}
{khử \ce{Cl-} thành \ce{Cl2}}
{oxi hoá \ce{Cl-} thành \ce{Cl2}}
""" + (r"{\True khử \ce{H2O} thành \ce{H2} và \ce{OH-}}" if is_teacher else r"{khử \ce{H2O} thành \ce{H2} và \ce{OH-}}") + r"""
\loigiai{""" + (r"""
Ở cathode (cực âm), do ion \ce{Na+} không bị khử trong dung dịch nước nên phân tử \ce{H2O} nhận electron (bị khử):
$$\ce{2H2O + 2e -> H2 ^ + 2OH-.}$$
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Khi điện phân dung dịch \ce{CuSO4} bằng dòng điện một chiều (với điện cực anode bằng Cu) thì ở anode xảy ra quá trình
\choice
{oxi hoá \ce{H2O} thành \ce{H+} và \ce{O2}}
{khử \ce{Cu^2+} thành Cu}
""" + (r"{\True oxi hoá Cu thành \ce{Cu^2+}}" if is_teacher else r"{oxi hoá Cu thành \ce{Cu^2+}}") + r"""
{khử \ce{H2O} thành \ce{H2} và \ce{OH-}}
\loigiai{""" + (r"""
Hiện tượng dương cực tan: Điện cực anode bằng đồng (Cu) bị oxi hoá tan vào dung dịch theo bán phản ứng: $\ce{Cu -> Cu^2+ + 2e}$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
So sánh hai kim loại A và B có thế điện cực chuẩn lần lượt là $-0{,}25\text{ V}$ và $+0{,}34\text{ V}$, kim loại nào có khả năng khử mạnh hơn?
\choice
""" + (r"{\True Kim loại A}" if is_teacher else r"{Kim loại A}") + r"""
{Kim loại B}
{Cả hai đều có khả năng khử như nhau}
{Không thể xác định}
\loigiai{""" + (r"""
Kim loại có thế điện cực chuẩn nhỏ hơn (âm hơn) thì có tính khử mạnh hơn. Do $E^\circ_{\text{A}} (-0{,}25\text{ V}) < E^\circ_{\text{B}} (+0{,}34\text{ V})$ nên kim loại A có khả năng khử mạnh hơn kim loại B.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Cho pin điện hoá Zn--Cu. Quá trình xảy ra ở cực dương của pin là
\choice
{\ce{Zn^2+ + 2e -> Zn}}
""" + (r"{\True \ce{Cu^2+ + 2e -> Cu}}" if is_teacher else r"{\ce{Cu^2+ + 2e -> Cu}}") + r"""
{\ce{Zn -> Zn^2+ + 2e}}
{\ce{Cu -> Cu^2+ + 2e}}
\loigiai{""" + (r"""
Trong pin điện hoá Zn -- Cu, cực dương là cathode (thanh Cu). Ở cathode xảy ra quá trình khử ion \ce{Cu^2+}: $\ce{Cu^2+ + 2e -> Cu}$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Cho pin điện hoá Al--Pb. Quá trình xảy ra ở cực âm của pin là
\choice
{\ce{Al^3+ + 3e -> Al}}
{\ce{Pb^2+ + 2e -> Pb}}
""" + (r"{\True \ce{Al -> Al^3+ + 3e}}" if is_teacher else r"{\ce{Al -> Al^3+ + 3e}}") + r"""
{\ce{Pb -> Pb^2+ + 2e}}
\loigiai{""" + (r"""
Vì $E^\circ_{\ce{Al^3+/Al}} (-1{,}66\text{ V}) < E^\circ_{\ce{Pb^2+/Pb}} (-0{,}13\text{ V})$ nên kim loại Al có tính khử mạnh hơn, đóng vai trò là cực âm (anode). Ở anode xảy ra quá trình oxi hoá: $\ce{Al -> Al^3+ + 3e}$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\begin{ex}
(SBT – CTST). Phát biểu nào sau đây \textbf{không} đúng về pin Galvani?
\choice
""" + (r"{\True Ở điện cực dương xảy ra quá trình oxi hoá}" if is_teacher else r"{Ở điện cực dương xảy ra quá trình oxi hoá}") + r"""
{Cathode là kim loại yếu hơn, đóng vai trò điện cực dương của pin}
{Anode là kim loại mạnh hơn, đóng vai trò điện cực âm của pin}
{Phản ứng hóa học diễn ra trong pin kèm theo sự giải phóng điện năng}
\loigiai{""" + (r"""
Ở điện cực dương (cathode) của pin điện hoá luôn xảy ra quá trình khử (nhận electron), không phải quá trình oxi hoá. Phát biểu A là sai.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Khi điện phân dung dịch nào sau đây, tại anode xảy ra quá trình oxi hóa nước?
\choice
{Dung dịch \ce{ZnCl2}}
{Dung dịch \ce{CuCl2}}
""" + (r"{\True Dung dịch \ce{AgNO3}}" if is_teacher else r"{Dung dịch \ce{AgNO3}}") + r"""
{Dung dịch \ce{MgCl2}}
\loigiai{""" + (r"""
Anion \ce{Cl-} trong các dung dịch \ce{ZnCl2}, \ce{CuCl2}, \ce{MgCl2} dễ bị oxi hoá hơn nước. Còn anion \ce{NO3-} trong dung dịch \ce{AgNO3} không bị oxi hoá trong nước, do đó phân tử \ce{H2O} bị oxi hoá tại anode giải phóng \ce{O2} và \ce{H+}.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Điện phân dung dịch \ce{CuCl2} với điện cực trơ. Nồng độ của \ce{CuCl2} trong dung dịch
\choice
""" + (r"{\True giảm dần}" if is_teacher else r"{giảm dần}") + r"""
{tăng dần}
{không thay đổi}
{tăng dần rồi giảm dần}
\loigiai{""" + (r"""
Phương trình điện phân dung dịch: $\ce{CuCl2 ->[\text{đpdd}] Cu + Cl2 ^}$. Do chất tan \ce{CuCl2} liên tục bị tiêu hao chuyển thành Cu và \ce{Cl2} nên nồng độ \ce{CuCl2} trong dung dịch giảm dần.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Điện phân dung dịch \ce{CuSO4} với hai điện cực trơ. Sau một thời gian, màu xanh của dung dịch nhạt dần do:
\choice
{khí \ce{H2} sinh ra đã khử màu của dung dịch}
{dung dịch bị pha loãng dần nên màu xanh nhạt dần}
""" + (r"{\True ion \ce{Cu^2+} bị khử dần thành đồng kim loại}" if is_teacher else r"{ion \ce{Cu^2+} bị khử dần thành đồng kim loại}") + r"""
{ion \ce{Cu^2+} được tạo thêm}
\loigiai{""" + (r"""
Màu xanh đặc trưng của dung dịch là do các ion \ce{Cu^2+} tạo nên. Trong quá trình điện phân, cation \ce{Cu^2+} liên tục bị khử ở cathode thành kim loại Cu: $\ce{Cu^2+ + 2e -> Cu}$, làm giảm nồng độ \ce{Cu^2+}, do đó màu xanh của dung dịch nhạt dần.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Điện phân điện cực trơ, có màng ngăn một dung dịch chứa các ion: \ce{Fe^2+}, \ce{Fe^3+}, \ce{Cu^2+}, \ce{Cl-}. Thứ tự điện phân xảy ra ở cathode là
\choice
{\ce{Fe^2+}, \ce{Fe^3+}, \ce{Cu^2+}}
{\ce{Fe^2+}, \ce{Cu^2+}, \ce{Fe^3+}}
""" + (r"{\True \ce{Fe^3+}, \ce{Cu^2+}, \ce{Fe^2+}}" if is_teacher else r"{\ce{Fe^3+}, \ce{Cu^2+}, \ce{Fe^2+}}") + r"""
{\ce{Fe^3+}, \ce{Fe^2+}, \ce{Cu^2+}}
\loigiai{""" + (r"""
Tại cathode (cực âm), ion có tính oxi hoá mạnh hơn sẽ nhận electron trước. Thứ tự tính oxi hoá của các cation:\\
$$\ce{Fe^3+} (E^\circ = 0{,}77\text{ V}) > \ce{Cu^2+} (E^\circ = 0{,}34\text{ V}) > \ce{Fe^2+} (E^\circ = -0{,}44\text{ V}).$$
Do đó thứ tự điện phân xảy ra ở cathode là: \ce{Fe^3+}, \ce{Cu^2+}, \ce{Fe^2+}.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Cho pin điện hoá Mn--Cd có $E^\circ_{\text{pin(Mn-Cd)}} = 0{,}79\text{ V}$ và $E^\circ_{\ce{Cd^2+/Cd}} = -0{,}40\text{ V}$. Thế điện cực chuẩn của cặp oxi hoá khử \ce{Mn^2+/Mn} có giá trị là
\choice
{0,27 V}
""" + (r"{\True -1,19 V}" if is_teacher else r"{-1,19 V}") + r"""
{0,26 V}
{-0,25 V}
\loigiai{""" + (r"""
Ta có công thức:
$$E^\circ_{\text{pin}} = E^\circ_{\ce{Cd^2+/Cd}} - E^\circ_{\ce{Mn^2+/Mn}}$$
$$\Rightarrow E^\circ_{\ce{Mn^2+/Mn}} = E^\circ_{\ce{Cd^2+/Cd}} - E^\circ_{\text{pin}} = -0{,}40 - 0{,}79 = -1{,}19\text{ V.}$$
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Quá trình điện phân dung dịch \ce{NaCl} (điện cực trơ, có màng ngăn) và điện phân \ce{NaCl} nóng chảy (điện cực trơ) có điểm giống nhau là
\choice
{tại cathode xảy ra sự khử ion \ce{Na+}}
{tại cathode xảy ra sự khử phân tử \ce{H2O}}
""" + (r"{\True tại anode xảy ra sự oxi hóa ion \ce{Cl-}}" if is_teacher else r"{tại anode xảy ra sự oxi hóa ion \ce{Cl-}}") + r"""
{tại anode xảy ra sự oxi hóa phân tử \ce{H2O}}
\loigiai{""" + (r"""
Cả hai quá trình đều có anion \ce{Cl-} di chuyển về anode (cực dương) và bị oxi hoá thành khí chlorine:
$$\ce{2Cl- -> Cl2 + 2e.}$$
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho sơ đồ phản ứng:\\
(1) $\ce{NaCl + H2O ->[\text{đpdd, cmn}] X + Y ^}\text{ (ở cathode)} + \ce{Z ^}\text{ (ở anode)}$\\
(2) $\ce{X + CO2 \text{ (dư)} -> T}$\\
Phát biểu nào sau đây đúng?
\choice
{Chất khí Z có thể khử được CaO thành Ca ở nhiệt độ cao}
{Dung dịch X có tính tẩy màu, sát trùng, thường gọi là Javel}
""" + (r"{\True Chất T được dùng làm thuốc giảm đau dạ dày}" if is_teacher else r"{Chất T được dùng làm thuốc giảm đau dạ dày}") + r"""
{Chất khí Y không có màu, mùi, vị và Y có thể duy trì sự cháy}
\loigiai{""" + (r"""
Phương trình phản ứng:\\
(1) $\ce{2NaCl + 2H2O ->[\text{đpdd, cmn}] 2NaOH (X) + H2 ^ (Y) + Cl2 ^ (Z).}$\\
(2) $\ce{NaOH + CO2 \text{ (dư)} -> NaHCO3 (T).}$\\
Chất T là \ce{NaHCO3} (natri hydrocarbonat hay baking soda) được sử dụng làm thuốc đau dạ dày do có tác dụng trung hoà acid HCl dư trong dạ dày. Phát biểu C đúng.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\setcounter{ex}{0}

\begin{ex}
(SBT – CĐ). Chọn tính đúng sai cho mỗi phát biểu sau về acquy:
\choiceTF
""" + (r"{\True Một ưu điểm của acquy là tái sử dụng được nhiều lần}" if is_teacher else r"{Một ưu điểm của acquy là tái sử dụng được nhiều lần}") + r"""
""" + (r"{\True Phản ứng xảy ra trong acquy cũng giống như phản ứng xảy ra trong pin Galvani nhưng có thể đảo ngược}" if is_teacher else r"{Phản ứng xảy ra trong acquy cũng giống như phản ứng xảy ra trong pin Galvani nhưng có thể đảo ngược}") + r"""
{Acquy không gây ô nhiễm môi trường}
{Acquy là nguồn điện hoá học có thể hoạt động liên tục}
\loigiai{""" + (r"""
a) \textbf{Đúng}. Acquy là nguồn điện thứ cấp (pin sạc) có thể sạc và tái sử dụng nhiều lần.\\
b) \textbf{Đúng}. Khi phóng điện phản ứng diễn ra như pin Galvani; khi sạc điện phản ứng diễn ra ngược lại như bình điện phân.\\
c) \textbf{Sai}. Acquy chứa kim loại nặng độc hại (chì, axit \ce{H2SO4}, lithi...) gây nguy cơ ô nhiễm nghiêm trọng nếu không xử lý đúng quy chuẩn.\\
d) \textbf{Sai}. Acquy chỉ hoạt động trong một khoảng thời gian nhất định rồi cạn năng lượng, phải nạp lại mới hoạt động tiếp, không thể chạy liên tục vô hạn.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Em hãy cho biết những thí nghiệm sau đây có thể tạo thành một pin điện hóa là đúng hay sai:
\choiceTF
""" + (r"{\True Một điện cực nhôm và một điện cực đồng cùng nhúng vào dung dịch nước muối}" if is_teacher else r"{Một điện cực nhôm và một điện cực đồng cùng nhúng vào dung dịch nước muối}") + r"""
{Một điện cực nhôm và một điện cực đồng cùng nhúng vào nước cất}
{Hai điện cực làm bằng đồng cùng nhúng vào dung dịch nước vôi trong}
{Hai điện cực làm bằng nhựa nhúng vào dầu hoả}
\loigiai{""" + (r"""
Điều kiện để tạo thành pin điện hoá: Hai điện cực khác bản chất cùng tiếp xúc với dung dịch chất điện li và nối kín mạch ngoài.\\
a) \textbf{Đúng}. Al và Cu khác bản chất, nước muối \ce{NaCl} là dung dịch chất điện li.\\
b) \textbf{Sai}. Nước cất không phải là dung dịch dẫn điện tốt (không phải dung dịch điện li).\\
c) \textbf{Sai}. Cả hai cực cùng bằng Cu (cùng bản chất) nên sức điện động bằng 0.\\
d) \textbf{Sai}. Nhựa không dẫn điện, dầu hoả không phải chất điện li.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Trong quá trình điện phân dung dịch \ce{CuSO4} với anode bằng đồng. Chọn tính đúng sai cho mỗi phát biểu sau:
\choiceTF
{Ở anode xảy ra quá trình oxi hóa nước}
{Khối lượng anode không thay đổi}
{Nồng độ \ce{CuSO4} trong dung dịch giảm dần}
""" + (r"{\True Khối lượng cathode tăng}" if is_teacher else r"{Khối lượng cathode tăng}") + r"""
\loigiai{""" + (r"""
Đây là hiện tượng dương cực tan:\\
- Ở anode: Đồng bị oxi hoá tan ra: $\ce{Cu -> Cu^2+ + 2e} \Rightarrow$ khối lượng anode giảm, nước không bị oxi hoá $\Rightarrow$ ý a, b \textbf{Sai}.\\
- Ở cathode: Ion \ce{Cu^2+} bị khử tạo Cu bám vào cathode: $\ce{Cu^2+ + 2e -> Cu} \Rightarrow$ khối lượng cathode tăng $\Rightarrow$ ý d \textbf{Đúng}.\\
- Lượng \ce{Cu^2+} sinh ra ở anode bằng lượng \ce{Cu^2+} mất đi ở cathode $\Rightarrow$ nồng độ \ce{CuSO4} trong dung dịch không đổi $\Rightarrow$ ý c \textbf{Sai}.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Điện phân dung dịch \ce{NaCl} bão hoà (với điện cực trơ, màng ngăn xốp) đến khi nồng độ \ce{NaCl} giảm đi một nửa thì dừng điện phân. Chọn tính đúng sai cho mỗi phát biểu sau:
\choiceTF
""" + (r"{\True Dung dịch sau điện phân làm phenolphthalein chuyển màu hồng}" if is_teacher else r"{Dung dịch sau điện phân làm phenolphthalein chuyển màu hồng}") + r"""
{Ở cathode chỉ xảy ra quá trình khử ion \ce{Na+}}
""" + (r"{\True Số mol khí \ce{Cl2} thoát ra ở anode bằng số mol \ce{H2} thoát ra ở cathode}" if is_teacher else r"{Số mol khí \ce{Cl2} thoát ra ở anode bằng số mol \ce{H2} thoát ra ở cathode}") + r"""
{Thứ tự điện phân ở anode là \ce{H2O}, \ce{Cl-}}
\loigiai{""" + (r"""
Phương trình điện phân: $\ce{2NaCl + 2H2O -> 2NaOH + H2 ^ + Cl2 ^}$.\\
a) \textbf{Đúng}. Dung dịch chứa base mạnh \ce{NaOH} làm phenolphthalein hoá hồng.\\
b) \textbf{Sai}. Ở cathode xảy ra sự khử \ce{H2O}, ion \ce{Na+} không bị khử.\\
c) \textbf{Đúng}. Theo tỉ lệ phản ứng, $n_{\ce{Cl2}} = n_{\ce{H2}}$.\\
d) \textbf{Sai}. Ở anode ion \ce{Cl-} có tính khử mạnh hơn nước nên bị oxi hoá trước.""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}

\setcounter{ex}{0}

\begin{ex}
Cho bảng số liệu thế điện cực chuẩn: \ce{Al^3+/Al} ($-1{,}676\text{ V}$); \ce{Zn^2+/Zn} ($-0{,}763\text{ V}$); \ce{Fe^2+/Fe} ($-0{,}440\text{ V}$); \ce{Cu^2+/Cu} ($+0{,}340\text{ V}$); \ce{Ag+/Ag} ($+0{,}799\text{ V}$); \ce{Au^3+/Au} ($+1{,}520\text{ V}$). Hãy cho biết có bao nhiêu kim loại trong bảng có tính khử mạnh hơn Cu?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}3}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Kim loại có tính khử mạnh hơn Cu là kim loại có thế điện cực chuẩn nhỏ hơn $E^\circ_{\ce{Cu^2+/Cu}} (+0{,}340\text{ V})$.\\
Các kim loại thoả mãn gồm: Al ($-1{,}676\text{ V}$), Zn ($-0{,}763\text{ V}$), Fe ($-0{,}440\text{ V}$). Có tất cả 3 kim loại.\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho pin điện hóa tạo bởi hai cặp oxi hoá -- khử \ce{Zn^2+/Zn} và \ce{Cu^2+/Cu}. Cho $E^\circ_{\ce{Cu^2+/Cu}} = +0{,}34\text{ V}$ và $E^\circ_{\ce{Zn^2+/Zn}} = -0{,}76\text{ V}$. Tính sức điện động chuẩn của pin (theo đơn vị V).
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}1,1}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Sức điện động chuẩn của pin:
$$E^\circ_{\text{pin}} = E^\circ_{\ce{Cu^2+/Cu}} - E^\circ_{\ce{Zn^2+/Zn}} = +0{,}34 - (-0{,}76) = 1{,}10\text{ V.}$$
\textbf{Đáp số: 1,1.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các kim loại sau: Na, Ca, Fe, Al, Zn, Cu, Ag. Số kim loại được điều chế trong công nghiệp bằng phương pháp điện phân hợp chất nóng chảy của chúng là bao nhiêu?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}3}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Các kim loại có tính khử mạnh (từ Al trở về trước) được điều chế bằng phương pháp điện phân nóng chảy:\\
- Na (điện phân nóng chảy \ce{NaCl}).\\
- Ca (điện phân nóng chảy \ce{CaCl2}).\\
- Al (điện phân nóng chảy \ce{Al2O3} trong cryolite).\\
Vậy có 3 kim loại là Na, Ca và Al.\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các chất rắn sau đây: \ce{NaCl}, \ce{CaCO3}, \ce{CuSO4}, \ce{NaOH}, \ce{CaCl2}, \ce{FeCl2}, \ce{CuCl2}, \ce{MgCl2}, \ce{Al2O3}. Số chất có thể dùng để điều chế kim loại bằng phương pháp điện phân nóng chảy là bao nhiêu?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}5}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Các chất có thể điện phân nóng chảy tạo kim loại tương ứng:\\
1. $\ce{2NaCl ->[\text{đpnc}] 2Na + Cl2}$.\\
2. $\ce{4NaOH ->[\text{đpnc}] 4Na + O2 + 2H2O}$.\\
3. $\ce{CaCl2 ->[\text{đpnc}] Ca + Cl2}$.\\
4. $\ce{MgCl2 ->[\text{đpnc}] Mg + Cl2}$.\\
5. $\ce{2Al2O3 ->[\text{đpnc}] 4Al + 3O2}$.\\
Tổng cộng có 5 chất.\\
\textbf{Đáp số: 5.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Một viên pin lithium-ion của điện thoại có điện áp danh định 3,7 V và dung lượng 4000 mAh. Điện thoại tiêu thụ trung bình dòng điện 500 mA khi sử dụng liên tục. Hỏi điện thoại có thể hoạt động liên tục trong bao lâu trước khi cần sạc lại (theo đơn vị giờ)?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}8}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Thời gian hoạt động liên tục của điện thoại:
$$t = \frac{\text{Dung lượng pin (mAh)}}{\text{Cường độ dòng tiêu thụ (mA)}} = \frac{4000}{500} = 8\text{ giờ.}$$
\textbf{Đáp số: 8.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Một nhà máy muốn mạ một lớp đồng lên một tấm kim loại có diện tích $1\text{ m}^2$ để chống ăn mòn. Họ sử dụng dung dịch \ce{CuSO4} và dòng điện có cường độ 500 A trong thời gian 2 giờ. Biết khối lượng mol của Cu là 64 g/mol. Hãy tính khối lượng đồng (theo đơn vị kg) bám vào tấm kim loại. (Làm tròn kết quả đến hàng phần trăm).
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}1,19}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
- Điện lượng cung cấp:
$$Q = I \times t = 500 \times (2 \times 3600) = 3\,600\,000\text{ C.}$$
- Số mol electron trao đổi:
$$n_e = \frac{Q}{F} = \frac{3\,600\,000}{96500} \approx 37{,}3057\text{ mol.}$$
- Quá trình ở cathode: $\ce{Cu^2+ + 2e -> Cu} \Rightarrow n_{\ce{Cu}} = \frac{n_e}{2} \approx 18{,}6528\text{ mol}$.\\
- Khối lượng đồng bám vào:
$$m = n_{\ce{Cu}} \times M_{\ce{Cu}} = 18{,}6528 \times 64 \approx 1193{,}78\text{ g} \approx 1{,}19\text{ kg.}$$
\textbf{Đáp số: 1,19.}""" if is_teacher else "") + r"""}
\end{ex}

"""

    if is_teacher:
        tex += r"""
\Closesolutionfile{ans}

\vspace{0.5cm}
\begin{center}
	\textbf{\Large\color{mainGreen}BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ SỐ 15}
\end{center}

\vspace{0.2cm}
\noindent\textbf{\color{mainGreen}PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} & \textbf{17} & \textbf{18} \\ \hline
\textbf{\color{red}A} & \textbf{\color{red}A} & \textbf{\color{red}B} & \textbf{\color{red}B} & \textbf{\color{red}A} & \textbf{\color{red}D} & \textbf{\color{red}C} & \textbf{\color{red}A} & \textbf{\color{red}B} & \textbf{\color{red}C} & \textbf{\color{red}A} & \textbf{\color{red}C} & \textbf{\color{red}A} & \textbf{\color{red}C} & \textbf{\color{red}C} & \textbf{\color{red}B} & \textbf{\color{red}C} & \textbf{\color{red}C} \\ \hline
\end{tabular}
\end{center}

\vspace{0.3cm}
\noindent\textbf{\color{mainGreen}PHẦN II: TRẮC NGHIỆM ĐÚNG SAI}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Lệnh hỏi a} & \textbf{Lệnh hỏi b} & \textbf{Lệnh hỏi c} & \textbf{Lệnh hỏi d} \\ \hline
\textbf{1} & \textbf{\color{red}Đ} & \textbf{\color{red}Đ} & \textbf{S} & \textbf{S} \\ \hline
\textbf{2} & \textbf{\color{red}Đ} & \textbf{S} & \textbf{S} & \textbf{S} \\ \hline
\textbf{3} & \textbf{S} & \textbf{S} & \textbf{S} & \textbf{\color{red}Đ} \\ \hline
\textbf{4} & \textbf{\color{red}Đ} & \textbf{S} & \textbf{\color{red}Đ} & \textbf{S} \\ \hline
\end{tabular}
\end{center}

\vspace{0.3cm}
\noindent\textbf{\color{mainGreen}PHẦN III: TRẢ LỜI NGẮN}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\ \hline
\textbf{\color{red}3} & \textbf{\color{red}1,1} & \textbf{\color{red}3} & \textbf{\color{red}5} & \textbf{\color{red}8} & \textbf{\color{red}1,19} \\ \hline
\end{tabular}
\end{center}
"""
    else:
        tex += r"""
\Closesolutionfile{ans}
"""

    tex += r"""
\end{document}
"""
    return tex

def compile_and_check(tex_filename, is_teacher=False):
    tex_path = os.path.join(OUT_DIR, tex_filename)
    pdf_filename = tex_filename.replace(".tex", ".pdf")
    pdf_path = os.path.join(OUT_DIR, pdf_filename)
    
    # Run pdflatex twice in San_Pham
    for i in range(2):
        proc = subprocess.run(
            ["pdflatex", "-interaction=nonstopmode", tex_filename],
            cwd=OUT_DIR,
            capture_output=True
        )
        stdout_str = proc.stdout.decode("utf-8", errors="replace") if proc.stdout else ""
        if proc.returncode != 0:
            print(f"Error compiling {tex_filename} (pass {i+1}, code {proc.returncode}):")
            lines = stdout_str.splitlines()
            err_lines = [l for l in lines if l.startswith("!") or "error" in l.lower() or "fatal" in l.lower()]
            print("Errors found:")
            print("\n".join(err_lines[:15]))
            print("Last 25 lines of log:")
            print("\n".join(lines[-25:]))
            return False
            
    # Check page count
    doc = fitz.open(pdf_path)
    page_count = len(doc)
    doc.close()
    print(f"Compiled {pdf_filename}: {page_count} pages.")
    
    # Clean temporary files
    for ext in ["aux", "log", "thm", "out"]:
        for f in glob.glob(os.path.join(OUT_DIR, f"*.{ext}")):
            try:
                os.remove(f)
            except:
                pass
    for f in glob.glob(os.path.join(OUT_DIR, "ans-*.tex")):
        try:
            os.remove(f)
        except:
            pass
            
    return page_count

if __name__ == "__main__":
    # Generate and compile Student version
    hs_tex_file = "De_Kiem_Tra_Hoa_12_De15_HocSinh.tex"
    hs_tex_content = create_latex_de15(is_teacher=False)
    with open(os.path.join(OUT_DIR, hs_tex_file), "w", encoding="utf-8") as f:
        f.write(hs_tex_content)
    print(f"Saved {hs_tex_file}")
    
    hs_pages = compile_and_check(hs_tex_file, is_teacher=False)
    print(f"Student version page count: {hs_pages} (Target: EXACTLY 4)")
    
    # Generate and compile Teacher version
    gv_tex_file = "De_Kiem_Tra_Hoa_12_De15_GiaoVien.tex"
    gv_tex_content = create_latex_de15(is_teacher=True)
    with open(os.path.join(OUT_DIR, gv_tex_file), "w", encoding="utf-8") as f:
        f.write(gv_tex_content)
    print(f"Saved {gv_tex_file}")
    
    gv_pages = compile_and_check(gv_tex_file, is_teacher=True)
    print(f"Teacher version page count: {gv_pages}")
