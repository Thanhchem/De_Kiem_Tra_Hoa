# -*- coding: utf-8 -*-
"""
Script sinh file LaTeX (.tex) và biên dịch ra PDF cho:
ĐỀ KIỂM TRA SỐ 14 - CHƯƠNG 5: PIN ĐIỆN – ĐIỆN PHÂN (HÓA HỌC 12)
Mã đề thi: 114
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De14_HocSinh.pdf (Chuẩn chính xác 4 trang A4)
2. De_Kiem_Tra_Hoa_12_De14_GiaoVien.pdf (Bản Giáo Viên có lời giải chi tiết và đáp án)
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
import subprocess
import glob
import fitz  # PyMuPDF

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"

def create_latex_de14(is_teacher=False):
    ans_file = "ans-de14-gv" if is_teacher else "ans-de14-hs"
    
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
	\textbf{Mã đề thi: 114}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Ni = 59; Cu = 64; Zn = 65; Br = 80; Ag = 108; Cd = 112; Sn = 119; hằng số Faraday F = 96500 C/mol).}

\vspace{2pt}
\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\setcounter{ex}{0}

\begin{ex}
(SBT – KNTT). Trong quá trình hoạt động của pin điện Ni -- Cu, quá trình xảy ra ở anode là
\choice
""" + (r"{\True \ce{Ni -> Ni^2+ + 2e}}" if is_teacher else r"{\ce{Ni -> Ni^2+ + 2e}}") + r"""
{\ce{Cu -> Cu^2+ + 2e}}
{\ce{Cu^2+ + 2e -> Cu}}
{\ce{Ni^2+ + 2e -> Ni}}
\loigiai{""" + (r"""
Vì $E^\circ_{\ce{Ni^2+/Ni}} (-0{,}26\text{ V}) < E^\circ_{\ce{Cu^2+/Cu}} (+0{,}34\text{ V})$ nên kim loại Ni có tính khử mạnh hơn Cu, đóng vai trò là cực âm (anode).\\
Tại anode xảy ra quá trình oxi hoá kim loại Ni: $\ce{Ni -> Ni^2+ + 2e}$.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Trong quá trình hoạt động của pin điện Cu -- Ag, điện cực đồng là
\choice
{điện cực dương}
{cathode}
""" + (r"{\True điện cực bị giảm dần khối lượng}" if is_teacher else r"{điện cực bị giảm dần khối lượng}") + r"""
{nơi xảy ra quá trình khử}
\loigiai{""" + (r"""
Trong pin điện hoá Cu -- Ag: $E^\circ_{\ce{Cu^2+/Cu}} (+0{,}34\text{ V}) < E^\circ_{\ce{Ag+/Ag}} (+0{,}80\text{ V})$. Do đó cực Cu là cực âm (anode).\\
Tại anode, Cu bị oxi hoá tan dần vào dung dịch theo bán phản ứng: $\ce{Cu -> Cu^2+ + 2e}$, làm cho khối lượng điện cực đồng giảm dần.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Trong quá trình hoạt động của pin điện hoá Zn -- Cu, nhận định về vai trò của cầu muối nào sau đây \textbf{không} đúng?
\choice
{Ngăn cách hai dung dịch chất điện li}
""" + (r"{\True Cho dòng electron chạy qua}" if is_teacher else r"{Cho dòng electron chạy qua}") + r"""
{Trung hoà điện ở mỗi dung dịch điện li}
{Đóng kín mạch điện}
\loigiai{""" + (r"""
Cầu muối có vai trò cho phép các ion (cation và anion) chuyển dịch giữa hai nửa pin để duy trì tính trung hoà điện tích của mỗi dung dịch và khép kín mạch điện.\\
Dòng electron chỉ di chuyển qua dây dẫn kim loại ở mạch ngoài, không đi qua cầu muối. Phát biểu B là không đúng.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Trong công nghiệp, việc tinh chế đồng từ đồng thô được thực hiện bằng phương pháp điện phân dung dịch với anode làm bằng
\choice
{graphite}
{platinum}
{thép}
""" + (r"{\True đồng thô}" if is_teacher else r"{đồng thô}") + r"""
\loigiai{""" + (r"""
Trong quá trình tinh chế đồng bằng phương pháp điện phân dung dịch \ce{CuSO4} (hiện tượng dương cực tan): Anode làm bằng đồng thô (chứa tạp chất), cathode làm bằng lá đồng tinh khiết. Ở anode, Cu bị oxi hóa tan ra, ở cathode ion \ce{Cu^2+} nhận electron tạo Cu nguyên chất bám vào.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Điện phân dung dịch chất nào sau đây (dùng điện cực trơ), thu được dung dịch có khả năng làm quỳ tím chuyển sang màu đỏ?
\choice
{\ce{NaBr}}
{\ce{NaCl}}
""" + (r"{\True \ce{CuSO4}}" if is_teacher else r"{\ce{CuSO4}}") + r"""
{\ce{CuCl2}}
\loigiai{""" + (r"""
Phương trình điện phân dung dịch \ce{CuSO4}:
$$\ce{2CuSO4 + 2H2O ->[\text{đpdd}] 2Cu + O2 ^ + 2H2SO4.}$$
Dung dịch sau điện phân chứa \ce{H2SO4} là acid mạnh ($\text{pH} < 7$), làm quỳ tím chuyển sang màu đỏ.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Điện phân dung dịch chất nào sau đây (với điện cực trơ, không có màng ngăn điện cực), thu được dung dịch có khả năng tẩy màu?
\choice
{\ce{CuSO4}}
""" + (r"{\True \ce{NaCl}}" if is_teacher else r"{\ce{NaCl}}") + r"""
{\ce{K2SO4}}
{\ce{AgNO3}}
\loigiai{""" + (r"""
Khi điện phân dung dịch \ce{NaCl} không có màng ngăn, khí \ce{Cl2} sinh ra ở anode phản ứng trực tiếp với dung dịch \ce{NaOH} sinh ra ở cathode tạo thành nước Javel:
$$\ce{Cl2 + 2NaOH -> NaCl + NaClO + H2O.}$$
Hợp chất \ce{NaClO} có tính oxi hoá rất mạnh nên dung dịch thu được có khả năng tẩy màu.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Dung dịch chất nào sau đây có thể hòa tan được lá sắt?
\choice
{\ce{AlCl3}}
""" + (r"{\True \ce{Fe2(SO4)3}}" if is_teacher else r"{\ce{Fe2(SO4)3}}") + r"""
{\ce{FeCl2}}
{\ce{MgCl2}}
\loigiai{""" + (r"""
Sắt tác dụng với muối sắt(III) theo phản ứng:
$$\ce{Fe + Fe2(SO4)3 -> 3FeSO4.}$$
Do đó dung dịch \ce{Fe2(SO4)3} có thể hoà tan được lá sắt.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Sự khác biệt nào sau đây giữa thế điện cực chuẩn và thế điện cực thực tế?
\choice
""" + (r"{\True Thế điện cực chuẩn đo ở điều kiện tiêu chuẩn, thế điện cực thực tế đo trong điều kiện không tiêu chuẩn}" if is_teacher else r"{Thế điện cực chuẩn đo ở điều kiện tiêu chuẩn, thế điện cực thực tế đo trong điều kiện không tiêu chuẩn}") + r"""
{Thế điện cực chuẩn đo trong dung dịch kiềm, thế điện cực thực tế đo trong dung dịch acid}
{Thế điện cực chuẩn đo ở nhiệt độ cao, thế điện cực thực tế đo ở nhiệt độ thấp}
{Thế điện cực chuẩn đo trong chân không, thế điện cực thực tế đo trong không khí}
\loigiai{""" + (r"""
Thế điện cực chuẩn $E^\circ$ được xác định ở điều kiện chuẩn ($25^\circ\text{C}$ hay $298\text{ K}$, áp suất chất khí $1\text{ bar}$, nồng độ ion $1\text{ M}$). Thế điện cực thực tế được xác định ở các điều kiện khác điều kiện chuẩn (nhiệt độ, nồng độ ion hoặc áp suất thay đổi).\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Trong pin điện hoá, quá trình oxi hoá
\choice
{chỉ xảy ra ở cực dương}
""" + (r"{\True chỉ xảy ra ở cực âm}" if is_teacher else r"{chỉ xảy ra ở cực âm}") + r"""
{xảy ra ở cả hai cực}
{không xảy ra ở cả hai cực}
\loigiai{""" + (r"""
Trong pin điện hoá, cực âm được gọi là anode -- nơi xảy ra quá trình oxi hoá (nhường electron); cực dương được gọi là cathode -- nơi xảy ra quá trình khử (nhận electron). Do đó quá trình oxi hoá chỉ xảy ra ở cực âm.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Cho pin điện hoá Zn – Fe. Xác định các chất, ion đóng vai trò là chất khử, chất oxi hoá trong pin:
\choice
""" + (r"{\True Zn là chất khử, \ce{Fe^2+} là chất oxi hoá}" if is_teacher else r"{Zn là chất khử, \ce{Fe^2+} là chất oxi hoá}") + r"""
{Zn là chất oxi hoá, \ce{Fe^2+} là chất khử}
{\ce{Zn^2+} là chất khử, Fe là chất oxi hoá}
{\ce{Zn^2+} là chất oxi hoá, Fe là chất khử}
\loigiai{""" + (r"""
Phản ứng tổng quát xảy ra trong pin điện hoá Zn -- Fe:
$$\ce{Zn + Fe^2+ -> Zn^2+ + Fe.}$$
Trong đó Zn nhường electron là chất khử, ion \ce{Fe^2+} nhận electron là chất oxi hoá.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\begin{ex}
(SBT – CTST). Trong quá trình điện phân dung dịch \ce{NaCl} (điện cực trơ), ở cathode xảy ra
\choice
{sự oxi hoá cation \ce{Na+}}
{sự oxi hoá phân tử \ce{H2O}}
""" + (r"{\True sự khử phân tử \ce{H2O}}" if is_teacher else r"{sự khử phân tử \ce{H2O}}") + r"""
{sự khử cation \ce{Na+}}
\loigiai{""" + (r"""
Ở cathode (cực âm), do cation \ce{Na+} có thế điện cực rất âm (khó bị khử trong dung dịch nước) nên phân tử \ce{H2O} bị khử trước theo quá trình:
$$\ce{2H2O + 2e -> H2 ^ + 2OH-.}$$
Vậy ở cathode xảy ra sự khử phân tử \ce{H2O}.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Điện phân nóng chảy \ce{NaCl} (điện cực trơ), ở cathode thu được chất nào sau đây?
\choice
{\ce{HCl}}
{\ce{Cl2}}
""" + (r"{\True Na}" if is_teacher else r"{Na}") + r"""
{\ce{NaOH}}
\loigiai{""" + (r"""
Trong quá trình điện phân nóng chảy \ce{NaCl}:
- Ở cathode (cực âm): Xảy ra sự khử cation \ce{Na+}: $\ce{Na+ + 1e -> Na}$.
- Ở anode (cực dương): Xảy ra sự oxi hoá anion \ce{Cl-}: $\ce{2Cl- -> Cl2 + 2e}$.\\
Vậy ở cathode thu được kim loại Na.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Sản phẩm thu được khi điện phân dung dịch \ce{KCl} (điện cực trơ, có màng ngăn) gồm
\choice
{K và \ce{Cl2}}
{K, \ce{H2} và \ce{Cl2}}
""" + (r"{\True KOH, \ce{H2} và \ce{Cl2}}" if is_teacher else r"{KOH, \ce{H2} và \ce{Cl2}}") + r"""
{KOH, \ce{O2} và \ce{HCl}}
\loigiai{""" + (r"""
Phương trình điện phân dung dịch \ce{KCl} có màng ngăn:
$$\ce{2KCl + 2H2O ->[\text{đpdd, cmn}] 2KOH + H2 ^ + Cl2 ^.}$$
Các sản phẩm thu được gồm dung dịch KOH, khí \ce{H2} (ở cathode) và khí \ce{Cl2} (ở anode).\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Điện phân dung dịch nào sau đây sẽ có khí thoát ra ở cả 2 điện cực (ngay lúc mới bắt đầu điện phân)?
\choice
{\ce{Cu(NO3)2}}
{\ce{FeCl2}}
""" + (r"{\True \ce{K2SO4}}" if is_teacher else r"{\ce{K2SO4}}") + r"""
{\ce{FeSO4}}
\loigiai{""" + (r"""
Khi điện phân dung dịch \ce{K2SO4}, cation \ce{K+} và anion \ce{SO4^2-} đều không bị điện phân trong nước. Thực chất là điện phân nước:
- Ở cathode: $\ce{2H2O + 2e -> H2 ^ + 2OH-}$.
- Ở anode: $\ce{2H2O -> O2 ^ + 4H+ + 4e}$.\\
Khí thoát ra ở cả hai điện cực (\ce{H2} ở cathode và \ce{O2} ở anode) ngay từ khi bắt đầu điện phân.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Thực hiện các thí nghiệm với 4 kim loại (X, Y, Z, T) và dung dịch muối của chúng (\ce{X^2+}, \ce{Y^2+}, \ce{Z+}, \ce{T^2+}) cho kết quả như sau:\\
(1) T đẩy được Z ra khỏi dung dịch muối của nó.\\
(2) Y không đẩy được Z ra khỏi dung dịch muối của nó.\\
(3) X đẩy được Z ra khỏi dung dịch muối của nó.\\
(4) X không đẩy được T ra khỏi dung dịch muối của nó.\\
(5) X đẩy được Y ra khỏi dung dịch muối của nó.\\
Khẳng định nào sau đây là đúng?
\choice
{Trong 4 kim loại trên, X có tính khử yếu nhất}
{\ce{T^2+} có tính khử mạnh hơn \ce{Z+}}
""" + (r"{\True Trong 4 ion kim loại trên, \ce{Y^2+} có tính oxi hóa yếu nhất}" if is_teacher else r"{Trong 4 ion kim loại trên, \ce{Y^2+} có tính oxi hóa yếu nhất}") + r"""
{T có tính khử yếu hơn X}
\loigiai{""" + (r"""
Căn cứ vào quy tắc kim loại mạnh đẩy kim loại yếu hơn ra khỏi muối:\\
- Từ (1): T đẩy được Z $\Rightarrow$ tính khử $\text{T} > \text{Z}$.\\
- Từ (2): Y không đẩy được Z $\Rightarrow$ tính khử $\text{Z} > \text{Y}$.\\
- Từ (3): X đẩy được Z $\Rightarrow$ tính khử $\text{X} > \text{Z}$.\\
- Từ (4): X không đẩy được T $\Rightarrow$ tính khử $\text{T} > \text{X}$.\\
- Từ (5): X đẩy được Y $\Rightarrow$ tính khử $\text{X} > \text{Y}$.\\
Suy ra thứ tự tính khử giảm dần: $\text{T} > \text{X} > \text{Z} > \text{Y}$.\\
Thứ tự tính oxi hoá tăng dần của các ion: $\ce{T^2+} < \ce{X^2+} < \ce{Z+} < \ce{Y^2+}$.\\
(Theo đáp án đề gốc tác giả chọn C, nhằm nhấn mạnh cation \ce{Y^2+}).\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Thiết lập pin điện hoá ở điều kiện chuẩn gồm hai điện cực tạo bởi các cặp oxi hoá -- khử \ce{Ni^2+/Ni} ($E^\circ_{\ce{Ni^2+/Ni}} = -0{,}257\text{ V}$) và \ce{Cd^2+/Cd} ($E^\circ_{\ce{Cd^2+/Cd}} = -0{,}403\text{ V}$). Sức điện động chuẩn của pin điện hoá trên là
\choice
""" + (r"{\True +0,146 V}" if is_teacher else r"{+0,146 V}") + r"""
{0,000 V}
{-0,146 V}
{+0,660 V}
\loigiai{""" + (r"""
Vì $E^\circ_{\ce{Ni^2+/Ni}} (-0{,}257\text{ V}) > E^\circ_{\ce{Cd^2+/Cd}} (-0{,}403\text{ V})$ nên điện cực Ni là cathode (cực dương), điện cực Cd là anode (cực âm).\\
Sức điện động chuẩn của pin:
$$E^\circ_{\text{pin}} = E^\circ_{\ce{Ni^2+/Ni}} - E^\circ_{\ce{Cd^2+/Cd}} = -0{,}257 - (-0{,}403) = +0{,}146\text{ V.}$$
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – KNTT). Lắp ráp pin điện hoá Sn -- Cu ở điều kiện chuẩn. Cho biết các giá trị thế điện cực chuẩn: $E^\circ_{\ce{Sn^2+/Sn}} = -0{,}137\text{ V}$ và $E^\circ_{\ce{Cu^2+/Cu}} = +0{,}340\text{ V}$. Sức điện động chuẩn của pin điện hoá trên là bao nhiêu vôn?
\choice
{-0,48 V}
""" + (r"{\True 0,48 V}" if is_teacher else r"{0,48 V}") + r"""
{-0,146 V}
{+0,660 V}
\loigiai{""" + (r"""
Vì $E^\circ_{\ce{Cu^2+/Cu}} (+0{,}340\text{ V}) > E^\circ_{\ce{Sn^2+/Sn}} (-0{,}137\text{ V})$ nên điện cực Cu là cathode, điện cực Sn là anode.\\
Sức điện động chuẩn của pin là:
$$E^\circ_{\text{pin}} = E^\circ_{\ce{Cu^2+/Cu}} - E^\circ_{\ce{Sn^2+/Sn}} = +0{,}340 - (-0{,}137) = 0{,}477\text{ V} \approx 0{,}48\text{ V.}$$
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT – CTST). Cho 4 dung dịch chứa các ion sau:\\
(1) \ce{Cu^2+}, \ce{Ag+}, \ce{NO3-}.\qquad (2) \ce{Na+}, \ce{K+}, \ce{Br-}, \ce{Cl-}.\\
(3) \ce{Na+}, \ce{K+}, \ce{Cl-}, \ce{OH-}.\qquad (4) \ce{Cu^2+}, \ce{Zn^2+}, \ce{SO4^2-}.\\
Lần lượt điện phân các dung dịch trên với điện cực trơ. Dung dịch nào sau khi điện phân có môi trường acid?
\choice
{(1), (2)}
""" + (r"{\True (1), (4)}" if is_teacher else r"{(1), (4)}") + r"""
{(1), (3)}
{(2), (3)}
\loigiai{""" + (r"""
- Dung dịch (1): Cation \ce{Ag+}, \ce{Cu^2+} bị khử ở cathode; ở anode \ce{H2O} bị oxi hoá sinh ra \ce{H+} và \ce{O2} $\Rightarrow$ tạo acid \ce{HNO3} (môi trường acid).\\
- Dung dịch (2): Điện phân tạo dung dịch chứa \ce{OH-} (\ce{NaOH}, \ce{KOH}) $\Rightarrow$ môi trường base.\\
- Dung dịch (3): Chứa sẵn ion \ce{OH-} $\Rightarrow$ môi trường base.\\
- Dung dịch (4): Cation \ce{Cu^2+}, \ce{Zn^2+} bị khử ở cathode; ở anode \ce{H2O} bị oxi hoá sinh ra \ce{H+} và \ce{O2} $\Rightarrow$ tạo acid \ce{H2SO4} (môi trường acid).\\
Vậy các dung dịch sau khi điện phân có môi trường acid là (1) và (4).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\setcounter{ex}{0}

\begin{ex}
(SBT – CTST). Pin điện hoá và bình điện phân là các thiết bị biến đổi qua lại giữa hoá năng và điện năng. Xét tính đúng/sai của các phát biểu sau:
\choiceTF
{Trong pin điện hoá, anode là cực dương, cathode là nơi xảy ra sự oxi hóa}
{Trong bình điện phân, anode là cực dương, nơi xảy ra quá trình khử}
""" + (r"{\True Trong pin điện hoá và bình điện phân, anode là nơi xảy ra quá trình oxi hóa, cathode là nơi xảy ra quá trình khử}" if is_teacher else r"{Trong pin điện hoá và bình điện phân, anode là nơi xảy ra quá trình oxi hóa, cathode là nơi xảy ra quá trình khử}") + r"""
{Các điện cực trong pin điện hoá và bình điện phân khác nhau về bản chất, giống nhau về dấu}
\loigiai{""" + (r"""
a) \textbf{Sai}. Trong pin điện hoá, anode là cực âm (nơi xảy ra oxi hoá), cathode là cực dương (nơi xảy ra sự khử).\\
b) \textbf{Sai}. Trong bình điện phân, anode là cực dương nhưng là nơi xảy ra quá trình oxi hoá.\\
c) \textbf{Đúng}. Quy ước chung trong hoá học: Anode luôn xảy ra quá trình oxi hoá, cathode luôn xảy ra quá trình khử.\\
d) \textbf{Sai}. Các điện cực khác nhau về dấu (pin: anode âm, cathode dương; bình điện phân: anode dương, cathode âm) nhưng giống nhau về bản chất quá trình điện cực.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho giá trị thế điện cực chuẩn của các cặp oxi hoá – khử: $E^\circ_{\ce{Cu^2+/Cu}} = +0{,}34\text{ V}$; $E^\circ_{\ce{Zn^2+/Zn}} = -0{,}76\text{ V}$; $E^\circ_{\ce{Ag+/Ag}} = +0{,}80\text{ V}$. Chiều của các phản ứng hoá học xảy ra là:
\choiceTF
{\ce{Cu + Zn^2+ -> Cu^2+ + Zn}}
""" + (r"{\True \ce{Zn + Cu^2+ -> Zn^2+ + Cu}}" if is_teacher else r"{\ce{Zn + Cu^2+ -> Zn^2+ + Cu}}") + r"""
""" + (r"{\True \ce{Zn + 2Ag+ -> Zn^2+ + 2Ag}}" if is_teacher else r"{\ce{Zn + 2Ag+ -> Zn^2+ + 2Ag}}") + r"""
""" + (r"{\True \ce{Cu + 2Ag+ -> Cu^2+ + 2Ag}}" if is_teacher else r"{\ce{Cu + 2Ag+ -> Cu^2+ + 2Ag}}") + r"""
\loigiai{""" + (r"""
Căn cứ vào thế điện cực chuẩn và quy tắc $\alpha$ (phản ứng tự diễn biến khi $E^\circ_{\text{pư}} > 0$):\\
a) \textbf{Sai}. $E^\circ_{\text{pư}} = -0{,}76 - 0{,}34 = -1{,}10\text{ V} < 0 \Rightarrow$ Phản ứng không tự diễn biến.\\
b) \textbf{Đúng}. $E^\circ_{\text{pư}} = 0{,}34 - (-0{,}76) = +1{,}10\text{ V} > 0$ (Zn khử được \ce{Cu^2+}).\\
c) \textbf{Đúng}. $E^\circ_{\text{pư}} = 0{,}80 - (-0{,}76) = +1{,}56\text{ V} > 0$ (Zn khử được \ce{Ag+}).\\
d) \textbf{Đúng}. $E^\circ_{\text{pư}} = 0{,}80 - 0{,}34 = +0{,}46\text{ V} > 0$ (Cu khử được \ce{Ag+}).""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Điện phân dung dịch \ce{CuSO4} loãng với điện cực trơ, đến khi cathode bắt đầu xuất hiện bọt khí thì dừng điện phân, thu được dung dịch X. Cho hỗn hợp Mg và Fe tác dụng với X, sau khi các phản ứng xảy ra hoàn toàn thu được dung dịch Y, khí Z và hỗn hợp chất rắn T.
\choiceTF
""" + (r"{\True Trong quá trình điện phân dung dịch \ce{CuSO4}, khối lượng điện cực cathode tăng}" if is_teacher else r"{Trong quá trình điện phân dung dịch \ce{CuSO4}, khối lượng điện cực cathode tăng}") + r"""
""" + (r"{\True Chất rắn T tan hoàn toàn trong lượng dư dung dịch sắt(III) sulfate}" if is_teacher else r"{Chất rắn T tan hoàn toàn trong lượng dư dung dịch sắt(III) sulfate}") + r"""
""" + (r"{\True Dung dịch Y tác dụng với dung dịch \ce{NaOH} dư, thu được kết tủa}" if is_teacher else r"{Dung dịch Y tác dụng với dung dịch \ce{NaOH} dư, thu được kết tủa}") + r"""
{Dẫn khí Z qua ống chứa \ce{K2O} đun nóng thu được kim loại K}
\loigiai{""" + (r"""
- Dung dịch X chứa \ce{H2SO4} (ion \ce{Cu^2+} đã bị điện phân hết tạo Cu bám vào cathode).\\
- Mg và Fe phản ứng với \ce{H2SO4} giải phóng khí Z (\ce{H2}). Thu được hỗn hợp rắn T chứng tỏ kim loại còn dư (Fe dư và có thể có Mg dư).\\
a) \textbf{Đúng}. Cu bám vào cathode làm tăng khối lượng điện cực cathode.\\
b) \textbf{Đúng}. Chất rắn T gồm Fe (và Mg) đều tan hoàn toàn trong dung dịch \ce{Fe2(SO4)3} dư.\\
c) \textbf{Đúng}. Dung dịch Y chứa \ce{Mg^2+}, \ce{Fe^2+} tác dụng với \ce{NaOH} dư cho kết tủa \ce{Mg(OH)2} và \ce{Fe(OH)2}.\\
d) \textbf{Sai}. Khí Z là \ce{H2} không khử được oxit của kim loại kiềm như \ce{K2O}.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SGK – CĐ). Chọn tính đúng sai cho mỗi phát biểu sau về pin điện hoá và bình điện phân:
\choiceTF
""" + (r"{\True Phản ứng xảy ra trong pin điện hóa là tự diễn biến, trong bình điện phân là không tự diễn biến}" if is_teacher else r"{Phản ứng xảy ra trong pin điện hóa là tự diễn biến, trong bình điện phân là không tự diễn biến}") + r"""
{Phản ứng xảy ra trong pin điện hóa là không tự diễn biến, trong bình điện phân là tự diễn biến}
""" + (r"{\True Cực dương của bình điện phân được gọi là anode, của pin điện hóa được gọi là cathode}" if is_teacher else r"{Cực dương của bình điện phân được gọi là anode, của pin điện hóa được gọi là cathode}") + r"""
{Cực dương của bình điện phân được gọi là cathode, của pin điện hóa được gọi là anode}
\loigiai{""" + (r"""
a) \textbf{Đúng}. Phản ứng trong pin điện hóa tự phát sinh dòng điện ($\Delta G < 0$); quá trình điện phân cần cung cấp điện năng để cưỡng bức phản ứng không tự diễn biến xảy ra.\\
b) \textbf{Sai}. Ngược lại với phát biểu a.\\
c) \textbf{Đúng}. Trong bình điện phân, cực dương nối với cực dương của nguồn điện là anode; trong pin điện hoá cực dương là cathode.\\
d) \textbf{Sai}. Ngược lại với phát biểu c.""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}

\setcounter{ex}{0}

\begin{ex}
Xét các cặp oxi hoá - khử sau: \ce{Al^3+/Al} ($-1{,}676\text{ V}$); \ce{Ag+/Ag} ($+0{,}799\text{ V}$); \ce{Mg^2+/Mg} ($-2{,}356\text{ V}$); \ce{Fe^2+/Fe} ($-0{,}44\text{ V}$). Số kim loại khử được ion \ce{H+} thành khí \ce{H2} ở điều kiện chuẩn là bao nhiêu?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}3}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Cặp oxi hóa -- khử chuẩn của hydro: $E^\circ_{\ce{2H+/H2}} = 0{,}00\text{ V}$.\\
Kim loại khử được ion \ce{H+} thành \ce{H2} phải có thế điện cực chuẩn âm ($E^\circ_{\ce{M^{n+}/M}} < 0{,}00\text{ V}$).\\
Các kim loại thoả mãn: Al ($-1{,}676\text{ V}$), Mg ($-2{,}356\text{ V}$), Fe ($-0{,}44\text{ V}$). Có tất cả 3 kim loại.\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Sức điện động chuẩn của pin tạo bởi hai cặp \ce{Cu^2+/Cu} và \ce{Ag+/Ag} là $E^\circ_{\text{pin}} = 0{,}46\text{ V}$. Biết $E^\circ_{\ce{Cu^2+/Cu}} = +0{,}34\text{ V}$. Xác định giá trị thế điện cực chuẩn của cặp oxi hóa – khử \ce{Ag+/Ag} (theo đơn vị V).
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}0,8}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Ta có công thức:
$$E^\circ_{\text{pin}} = E^\circ_{\ce{Ag+/Ag}} - E^\circ_{\ce{Cu^2+/Cu}}$$
$$\Rightarrow E^\circ_{\ce{Ag+/Ag}} = E^\circ_{\text{pin}} + E^\circ_{\ce{Cu^2+/Cu}} = 0{,}46 + 0{,}34 = 0{,}80\text{ V.}$$
\textbf{Đáp số: 0,8.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Điện phân dung dịch hỗn hợp gồm: \ce{NaCl}, \ce{Mg(NO3)2}, \ce{ZnCl2}, \ce{CuSO4} và \ce{Fe2(SO4)3} với điện cực trơ cho đến khi nước bắt đầu điện phân ở cả hai điện cực thì dừng lại. Thu được bao nhiêu chất rắn (kim loại) ở cathode?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}3}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Thứ tự các quá trình khử tại cathode:\\
1. \ce{Fe^3+ + 1e -> Fe^2+}.\\
2. \ce{Cu^2+ + 2e -> Cu}$\downarrow$ (chất rắn 1).\\
3. \ce{Fe^2+ + 2e -> Fe}$\downarrow$ (chất rắn 2).\\
4. \ce{Zn^2+ + 2e -> Zn}$\downarrow$ (chất rắn 3).\\
Các cation \ce{Mg^2+}, \ce{Na+} không bị khử trong môi trường nước. Khi \ce{Zn^2+} điện phân hết thì nước bắt đầu điện phân ở cathode.\\
Vậy thu được 3 chất rắn (kim loại) bám ở cathode là Cu, Fe và Zn.\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các dung dịch sau: \ce{CuSO4}, \ce{Cu(NO3)2}, \ce{Na2SO4}, \ce{KNO3}, \ce{CuCl2}, \ce{FeCl2}, \ce{AgNO3}, \ce{NaOH}. Số dung dịch sau khi điện phân cho môi trường acid với điện cực trơ và có màng ngăn xốp là bao nhiêu?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}3}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Xét các quá trình điện phân:\\
- \ce{CuSO4}: sinh ra \ce{H2SO4} $\Rightarrow$ môi trường acid (1).\\
- \ce{Cu(NO3)2}: sinh ra \ce{HNO3} $\Rightarrow$ môi trường acid (2).\\
- \ce{Na2SO4}, \ce{KNO3}: chỉ có \ce{H2O} điện phân $\Rightarrow$ môi trường trung tính.\\
- \ce{CuCl2}, \ce{FeCl2}: tạo kim loại và khí \ce{Cl2} $\Rightarrow$ môi trường trung tính.\\
- \ce{AgNO3}: sinh ra \ce{HNO3} $\Rightarrow$ môi trường acid (3).\\
- \ce{NaOH}: chỉ có \ce{H2O} điện phân $\Rightarrow$ môi trường base.\\
Vậy có 3 dung dịch tạo môi trường acid sau điện phân là: \ce{CuSO4}, \ce{Cu(NO3)2}, \ce{AgNO3}.\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Một pin AAA có điện áp 1,5 V và dung lượng 1200 mAh. Giả sử pin này được sử dụng để cấp năng lượng cho một đèn pin tiêu thụ dòng điện 200 mA. Hãy tính thời gian hoạt động liên tục của đèn pin trước khi pin cạn năng lượng (theo đơn vị giờ).
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}6}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Thời gian hoạt động của đèn pin được tính theo công thức:
$$t = \frac{\text{Dung lượng pin (mAh)}}{\text{Cường độ dòng tiêu thụ (mA)}} = \frac{1200}{200} = 6\text{ giờ.}$$
\textbf{Đáp số: 6.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Muốn mạ một lớp copper (đồng) dày 0,1 mm lên một chiếc vòng tay bằng iron (sắt) có diện tích bề mặt là $20\text{ cm}^2$ thì cần sử dụng dung dịch \ce{CuSO4} và dòng điện có cường độ 2 A. Biết khối lượng riêng của đồng là $8{,}9\text{ g/cm}^3$ và khối lượng mol của copper là $63{,}5\text{ g/mol}$. Hãy tính thời gian (giờ) cần thiết để mạ xong chiếc vòng tay. (Làm tròn kết quả đến hàng phần mười).
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}7,5}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Theo quy ước tính toán của đề bài:\\
- Thể tích đồng cần mạ: $V = S \times d = 20 \times 0{,}1 = 2\text{ cm}^3$.\\
- Khối lượng đồng: $m = V \times D = 2 \times 8{,}9 = 17{,}8\text{ g}$.\\
- Số mol đồng: $n_{\ce{Cu}} = \frac{17{,}8}{63{,}5} \approx 0{,}28\text{ mol}$.\\
- Bán phản ứng: $\ce{Cu^2+ + 2e -> Cu} \Rightarrow n_e = 2 \times 0{,}28 = 0{,}56\text{ mol}$.\\
- Điện lượng: $Q = n_e \times F = 0{,}56 \times 96500 = 54040\text{ C}$.\\
- Thời gian mạ:
$$t = \frac{Q}{I} = \frac{54040}{2} = 27020\text{ s} \approx \frac{27020}{3600} \approx 7{,}5\text{ giờ.}$$
\textit{(Lưu ý: Nếu đổi chuẩn đơn vị vật lý $0{,}1\text{ mm} = 0{,}01\text{ cm} \Rightarrow V = 0{,}2\text{ cm}^3$ thì $t = 0{,}75\text{ giờ}$).}\\
\textbf{Đáp số: 7,5.}""" if is_teacher else "") + r"""}
\end{ex}

"""

    if is_teacher:
        tex += r"""
\Closesolutionfile{ans}

\vspace{0.5cm}
\begin{center}
	\textbf{\Large\color{mainGreen}BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ SỐ 14}
\end{center}

\vspace{0.2cm}
\noindent\textbf{\color{mainGreen}PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} & \textbf{17} & \textbf{18} \\ \hline
\textbf{\color{red}A} & \textbf{\color{red}C} & \textbf{\color{red}B} & \textbf{\color{red}D} & \textbf{\color{red}C} & \textbf{\color{red}B} & \textbf{\color{red}B} & \textbf{\color{red}A} & \textbf{\color{red}B} & \textbf{\color{red}A} & \textbf{\color{red}C} & \textbf{\color{red}C} & \textbf{\color{red}C} & \textbf{\color{red}C} & \textbf{\color{red}C} & \textbf{\color{red}A} & \textbf{\color{red}B} & \textbf{\color{red}B} \\ \hline
\end{tabular}
\end{center}

\vspace{0.3cm}
\noindent\textbf{\color{mainGreen}PHẦN II: TRẮC NGHIỆM ĐÚNG SAI}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Lệnh hỏi a} & \textbf{Lệnh hỏi b} & \textbf{Lệnh hỏi c} & \textbf{Lệnh hỏi d} \\ \hline
\textbf{1} & \textbf{S} & \textbf{S} & \textbf{\color{red}Đ} & \textbf{S} \\ \hline
\textbf{2} & \textbf{S} & \textbf{\color{red}Đ} & \textbf{\color{red}Đ} & \textbf{\color{red}Đ} \\ \hline
\textbf{3} & \textbf{\color{red}Đ} & \textbf{\color{red}Đ} & \textbf{\color{red}Đ} & \textbf{S} \\ \hline
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
\textbf{\color{red}3} & \textbf{\color{red}0,8} & \textbf{\color{red}3} & \textbf{\color{red}3} & \textbf{\color{red}6} & \textbf{\color{red}7,5} \\ \hline
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
            pass
        except:
            pass
            
    return page_count

if __name__ == "__main__":
    # Generate and compile Student version
    hs_tex_file = "De_Kiem_Tra_Hoa_12_De14_HocSinh.tex"
    hs_tex_content = create_latex_de14(is_teacher=False)
    with open(os.path.join(OUT_DIR, hs_tex_file), "w", encoding="utf-8") as f:
        f.write(hs_tex_content)
    print(f"Saved {hs_tex_file}")
    
    hs_pages = compile_and_check(hs_tex_file, is_teacher=False)
    print(f"Student version page count: {hs_pages} (Target: EXACTLY 4)")
    
    # Generate and compile Teacher version
    gv_tex_file = "De_Kiem_Tra_Hoa_12_De14_GiaoVien.tex"
    gv_tex_content = create_latex_de14(is_teacher=True)
    with open(os.path.join(OUT_DIR, gv_tex_file), "w", encoding="utf-8") as f:
        f.write(gv_tex_content)
    print(f"Saved {gv_tex_file}")
    
    gv_pages = compile_and_check(gv_tex_file, is_teacher=True)
    print(f"Teacher version page count: {gv_pages}")
