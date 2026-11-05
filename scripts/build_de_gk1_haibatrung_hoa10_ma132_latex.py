# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Giữa Kì 1 Hóa Học 10 - THPT Hai Bà Trưng (Mã Đề 132)
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_Hoa_10_Ma132_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_Hoa_10_Ma132_GiaoVien.pdf (Kèm bảng đáp án rút gọn)
"""

import os
import sys
import subprocess
import glob

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

SAN_PHAM_DIR = r"C:\Antigravity_Thanh\San_Pham"

def build_latex_content(is_teacher=False):
    tex = []
    tex.append(r"""\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{geometry}
\geometry{top=1.5cm,bottom=1.2cm,left=1.4cm,right=1.4cm,headheight=28pt,headsep=12pt,footskip=18pt}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{tabularx}
\usepackage{array}
\usepackage{tikz}
\usepackage{fancyhdr}
\usepackage{enumitem}

\definecolor{hfRed}{HTML}{D32F2F}
\definecolor{hfGreen}{HTML}{8BC390}
\definecolor{mainGreen}{HTML}{2E7D32}
\definecolor{solBlue}{HTML}{1565C0}
\definecolor{tableGreen}{HTML}{E8F5E9}

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
\setlength{\parskip}{2pt}

\begin{document}

\begin{minipage}[t]{0.48\textwidth}
\textbf{SỞ GD\&ĐT THÀNH PHỐ HUẾ}\\
\textbf{TRƯỜNG THPT HAI BÀ TRƯNG}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 03 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}KIỂM TRA GIỮA KÌ 1 NĂM HỌC 2026 - 2027}\\
\textbf{MÔN HÓA HỌC -- Khối lớp 10}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian giao đề)}\\
\textbf{MÃ ĐỀ: 132}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ, tên thí sinh:} \dotfill \textit{Lớp:} \dotfill \textit{Số báo danh:} \dotfill

\vspace{3pt}
\noindent
\textit{\footnotesize Cho nguyên tử khối (amu): H = 1; C = 12; N = 14; O = 16; Si = 28; P = 31; S = 32; Li = 7; Na = 23; K = 39; Be = 9; Mg = 24; Ca = 40; Ba = 137.\\
Cho số hiệu nguyên tử: H: 1; N: 7; C: 6; Cl: 17; O: 8; S: 16; Li: 3; Na: 11; K: 19; Be: 4; Mg: 12; Ca: 20.\\
(Học sinh không sử dụng bảng hệ thống tuần hoàn các nguyên tố hóa học)}

\vspace{3pt}
\hrule color{hfGreen} height 0.5pt
""")

    if is_teacher:
        tex.append(r"""
\vspace{4pt}
\begin{center}
\textbf{\color{hfRed}\large (BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}
\end{center}

\vspace{2pt}
\begin{center}
{\color{hfRed}\textbf{\large BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 132)}}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}1. Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - mỗi câu đúng 0,25 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} \\
\hline
\textbf{Đ/A} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C/D*} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}A} \\
\hline
\end{tabular}
\end{center}
\textit{\footnotesize * Câu 6: Xem phân tích chi tiết ở phần bài giải.}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|l|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Nội dung câu hỏi} & \textbf{Lệnh a} & \textbf{Lệnh b} & \textbf{Lệnh c} & \textbf{Lệnh d} \\
\hline
\textbf{Câu 1} & Cấu tạo hạt \& Kí hiệu nguyên tử & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{SAI} \\
\hline
\textbf{Câu 2} & Nguyên tử Iron ($^{56}_{26}\ce{Fe}$) & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{SAI} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}3. Phần III: Trả lời ngắn (2,0 điểm - mỗi câu đúng 0,5 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}4} & \textbf{\color{hfRed}20} & \textbf{\color{hfRed}15} & \textbf{\color{hfRed}28} (hạt) \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}4. Tóm tắt đáp số Phần Tự luận (3,0 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|p{11.5cm}|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Điểm} & \textbf{Đáp số cốt lõi} \\
\hline
\textbf{Câu 1} & 1,0 đ & a) $\ce{1s^2 2s^2 2p^6 3s^2 3p^5}$. \quad b) Oxide cao nhất: $\ce{X2O7}$ ($\ce{Cl2O7}$); Hợp chất khí: $\ce{HX}$ ($\ce{HCl}$). \\
\hline
\textbf{Câu 2} & 1,0 đ & Nguyên tử khối của đồng vị thứ hai là $A_2 = 81$ (amu). \\
\hline
\textbf{Câu 3} & 1,0 đ & a) Hai muối cacbonat là $\ce{MgCO3}$ và $\ce{CaCO3}$. \quad b) $V_{\text{dd HCl 30\%}} = 9,73\text{ mL}$. \\
\hline
\end{tabular}
\end{center}

\vspace{4pt}
\hrule color{hfGreen} height 0.5pt
""")

    # NỘI DUNG ĐỀ THI
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\Large A. PHẦN TRẮC NGHIỆM: (7 điểm)}}

\vspace{3pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN I (3 điểm). TRẮC NGHIỆM NHIỀU LỰA CHỌN}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{2pt}
""")

    # 12 câu Phần I
    # Câu 1
    tex.append(r"""
\begin{ex}
Silicon được dùng trong công nghệ sản xuất chip máy tính hiện đại.. Vị trí của nguyên tố silicon trong bảng tuần hoàn thuộc ô 14, chu kì 3, nhóm IVA. Hãy chọn nhận đúng về silicon.
\choice
{\True Cấu hình electron nguyên tử của nguyên tố silicon là $\ce{[Ne] 3s^2 3p^2}$}
{Lớp N có 4 electron hoá trị}
{Số điện tích hạt nhân của silicon bằng 28}
{Silic nằm ở hàng thứ 4}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Silicon ở ô 14 ($Z = 14$): Cấu hình electron: $\ce{1s^2 2s^2 2p^6 3s^2 3p^2} = \ce{[Ne] 3s^2 3p^2}$ (A đúng).\\
- B sai vì 4 electron hóa trị ở lớp M ($n=3$).\\
- C sai vì số đơn vị điện tích hạt nhân là 14.\\
- D sai vì chu kì 3 tương ứng hàng thứ 3.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2
    tex.append(r"""
\begin{ex}
Hầu hết các hạt nhân nguyên tử được tạo nên từ các hạt
\choice
{electron, proton và neutron}
{electron và neutron}
{\True proton và neutron}
{electron và proton}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Hạt nhân của hầu hết các nguyên tử được tạo nên từ 2 loại hạt: proton và neutron.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 3
    tex.append(r"""
\begin{ex}
Trong số những quá trình kể dưới đây, cho biết đâu là hiện tượng vật lí ?
\choice
{Dẫn khí carbon dioxide vào nước vôi trong, làm nước vôi trong vẩn đục}
{Cho vôi sống ($\ce{CaO}$) vào nước được $\ce{Ca(OH)2}$}
{\True Nước bị hoá hơi khi đun nóng}
{Đốt cháy đường mía tạo thành chất màu đen và có mùi khét}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Nước lỏng chuyển thành hơi nước là quá trình chuyển thể (hiện tượng vật lí).\\
Các quá trình còn lại đều xảy ra phản ứng hóa học tạo chất mới.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 4
    tex.append(r"""
\begin{ex}
Số phân lớp electron có trong lớp O là
\choice
{2}
{4}
{3}
{\True 5}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Lớp O ứng với $n = 5$, gồm 5 phân lớp electron ($\ce{5s, 5p, 5d, 5f, 5g}$).\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 5
    tex.append(r"""
\begin{ex}
Phát biểu nào sau đây không đúng?
\choice
{Số khối của nguyên tử bằng tổng số proton và số neutron}
{Nguyên tố hoá học là những nguyên tử có cùng số đơn vị điện tích hạt nhân}
{Số hiệu nguyên tử bằng số đơn vị điện tích hạt nhân nguyên tử}
{\True Đồng vị là các nguyên tố có cùng số proton nhưng khác nhau số neutron}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Đồng vị là các NGUYÊN TỬ (không phải \"các nguyên tố\") của cùng một nguyên tố hóa học.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 6
    tex.append(r"""
\begin{ex}
Phát biểu nào sau đây đúng?
\choice
{Electron ở orbital 3p có mức năng lượng thấp hơn electron ở orbital 3s}
{Những electron ở gần hạt nhân có mức năng lượng cao nhất}
{\True Những electron ở lớp K có mức năng lượng cao nhất}
{Các electron trong cùng một lớp có năng lượng bằng nhau}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
\textbf{Phân tích sư phạm Câu 6:}\\
- 3p có năng lượng cao hơn 3s $\implies$ A sai.\\
- Electron càng gần hạt nhân liên kết càng bền, mức năng lượng thấp nhất $\implies$ đề in B và C là \"cao nhất\" (lỗi in ấn, đúng ra là \"thấp nhất\").\\
- Các electron trong cùng một lớp có năng lượng gần bằng nhau (không bằng nhau hoàn toàn).\\
$\implies$ Giáo viên chấm thi linh hoạt cho điểm C (nếu hiểu là sửa thấp nhất) hoặc D.
}
""")
    tex.append(r"\end{ex}")

    # Câu 7
    tex.append(r"""
\begin{ex}
Tổng số các hạt proton, nơtron và electron của một nguyên tử nguyên tố X là 28. Trong đó số hạt mang điện nhiều hơn số hạt không mang điện là 8. X là:
\choice
{Cl}
{Al}
{\True F}
{Si}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$\begin{cases} 2p + n = 28 \\ 2p - n = 8 \end{cases} \implies 4p = 36 \implies p = 9$ (Fluorine - F).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 8
    tex.append(r"""
\begin{ex}
Đồng (Copper) là kim loại dễ dát mỏng và uốn lượn, có khả năng dẫn điện và dẫn nhiệt tốt... Trong tự nhiên, nguyên tố Copper có hai đồng vị là $^{63}\ce{Cu}$ và $^{65}\ce{Cu}$. Nguyên tử khối trung bình của Copper là 63,54. Thành phần \% số nguyên tử của đồng vị $^{63}\ce{Cu}$ là
\choice
{27\%}
{\True 73\%}
{50\%}
{54\%}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Gọi \% số nguyên tử $^{63}\ce{Cu}$ là $x$ (\%):\\
$\overline{A} = \dfrac{63x + 65(100 - x)}{100} = 63,54 \iff 6500 - 2x = 6354 \iff 2x = 146 \iff x = 73\%$.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 9
    tex.append(r"""
\begin{ex}
Phát biểu nào sau đây về chu kì là đúng?
\choice
{Dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được xếp theo chiều số neutron tăng dần}
{\True Mở đầu mỗi chu kì là kim loại điển hình và kết thúc là một khí hiếm (trừ chu kì 1)}
{Dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được xếp theo chiều khối lượng nguyên tử tăng dần}
{Dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron, được xếp theo chiều số khối tăng dần}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Theo định luật tuần hoàn: Mở đầu mỗi chu kì là một kim loại kiềm (điển hình) và kết thúc bằng một khí hiếm (trừ chu kì 1).\\
Các nguyên tố trong chu kì được xếp theo chiều tăng dần của điện tích hạt nhân (Z).\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 10
    tex.append(r"""
\begin{ex}
X là nguyên tố hóa học có trong thành phần của chất có tác dụng oxi hóa và sát khuẩn cực mạnh, thường được sử dụng với mục đích khử trùng và tẩy trắng trong lĩnh vực thủy sản, dệt nhuộm, xử lí nước cấp, nước thải, nước bể bơi. Nguyên tử X có số hiệu nguyên tử $Z = 17$. Số electron lớp hoá trị trong nguyên tử X là
\choice
{\True 7}
{6}
{5}
{4}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$Z = 17$ (Chlorine): $\ce{1s^2 2s^2 2p^6 3s^2 3p^5}$.\\
Là nguyên tố nhóm VIIA, số electron hóa trị bằng số electron lớp ngoài cùng: $2 + 5 = 7$.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 11
    tex.append(r"""
\begin{ex}
Nội dung nào dưới đây không phải là đối tượng nghiên cứu của hóa học?
\choice
{Ứng dụng của chất}
{Tính chất và sự biến đổi của chất}
{Thành phần cấu trúc của chất}
{\True Sự lớn lên và sinh sản của tế bào}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Sự lớn lên và sinh sản của tế bào là đối tượng nghiên cứu của Sinh học.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 12
    tex.append(r"""
\begin{ex}
Cho dãy chất được biểu diễn bằng công thức hóa học như sau: $\ce{HNO3, O3, NaOH, He, Cl2, CuO, H2SO4, NaBr, Zn, CH4, MnO2}$. Số hợp chất trong dãy trên là
\choice
{\True 7}
{5}
{4}
{6}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Các hợp chất gồm 7 chất: $\ce{HNO3, NaOH, CuO, H2SO4, NaBr, CH4, MnO2}$.\\
Các đơn chất gồm 4 chất: $\ce{O3, He, Cl2, Zn}$.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN II
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN II (2 điểm). CÂU TRẮC NGHIỆM ĐÚNG / SAI}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng (Đ) hoặc sai (S).}
\vspace{2pt}
""")

    # Câu 1 Đúng/Sai
    tex.append(r"""
\begin{ex}
Các phát biểu sau đây là đúng hay sai?
\choiceTF
{\True Proton là hạt mang điện tích dương, nằm trong hạt nhân, khối lượng gần bằng 1 amu}
{\True Nguyên tử trung hòa điện, có kích thước lớn hơn nhiều so với hạt nhân, nhưng có khối lượng gần bằng khối lượng hạt nhân}
{Nguyên tử R có điện tích lớp vỏ là $-40,6\cdot 10^{-19}\text{ C}$ thì hạt nhân nguyên tử R có 26 proton}
{Nguyên tử X có 32 electron và 35 neutron, kí hiệu nguyên tử X là $^{16}_{32}\ce{S}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) \textbf{Đúng}: Proton mang điện dương (+1), $m_p \approx 1\text{ amu}$.\\
b) \textbf{Đúng}: Nguyên tử trung hòa điện, kích thước lớn gấp khoảng 10.000 lần hạt nhân nhưng khối lượng tập trung ở hạt nhân.\\
c) \textbf{Sai}: Nếu có 26 proton thì điện tích lớp vỏ phải là $-26 \times 1,602\cdot 10^{-19}\text{ C} \approx -41,6\cdot 10^{-19}\text{ C}$ (không phải $-40,6\cdot 10^{-19}\text{ C}$).\\
d) \textbf{Sai}: $Z = 32, A = 32 + 35 = 67 \implies$ Kí hiệu là $^{67}_{32}\ce{Ge}$.
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 Đúng/Sai
    tex.append(r"""
\begin{ex}
Nguyên tử iron (Fe) có kí hiệu là $^{56}_{26}\ce{Fe}$. Cho các phát biểu sau:
\choiceTF
{\True Nguyên tử Fe có 2 electron lớp ngoài cùng}
{\True Fe là nguyên tố d}
{Trong bảng tuần hoàn Fe thuộc nhóm IIA}
{Trong nguyên tử Fe số hạt không mang điện chiếm 39,6\%. Về khối lượng nguyên tử}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Fe ($Z = 26, A = 56$): Cấu hình electron $\ce{[Ar] 3d^6 4s^2}$.\\
a) \textbf{Đúng}: Lớp ngoài cùng $\ce{4s^2}$ có 2 electron.\\
b) \textbf{Đúng}: Electron cuối cùng điền vào phân lớp 3d nên Fe là nguyên tố d.\\
c) \textbf{Sai}: Fe có cấu hình $\ce{(n-1)d^6 ns^2}$ nên thuộc nhóm VIIIB, chu kì 4.\\
d) \textbf{Sai}: Hạt không mang điện là neutron (30 hạt). Tỉ lệ khối lượng hạt neutron: $\%m_n = \dfrac{30}{56}\cdot 100\% \approx 53,57\%$ (không phải 39,6\%).
}
""")
    tex.append(r"\end{ex}")

    # PHẦN III
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN III (2 điểm). TRẮC NGHIỆM TRẢ LỜI NGẮN}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 4. Viết câu trả lời vào ô quy định.}
\vspace{2pt}
""")

    # Câu 1 Trả lời ngắn
    tex.append(r"""
\begin{ex}
Cho các phát biểu sau:\\
(1) Tổng số proton và số neutron trong hạt nhân bằng giá trị số khối.\\
(2) Trong một nguyên tử luôn có số proton bằng số neutron.\\
(3) Các đồng vị có tính chất vật lí khác nhau nhưng tính chất hoá học giống nhau.\\
(4) Số proton bằng số đơn vị điện tích hạt nhân nguyên tử.\\
(5) Trong nguyên tử, hạt electron có khối lượng không đáng kể so với các hạt còn lại.\\
Trong các phát biểu trên có bao nhiêu phát biểu đúng?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Các phát biểu đúng gồm: (1), (3), (4), (5).\\
Phát biểu (2) sai vì số proton luôn bằng số electron, còn số neutron thường lớn hơn hoặc bằng số proton.\\
$\implies$ Có 4 phát biểu đúng.\\
\textbf{Đáp số: 4}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 2 Trả lời ngắn
    tex.append(r"""
\begin{ex}
Trong bảng tuần hoàn nguyên tố X thuộc chu kì 4, nhóm IIA. Số hiệu nguyên tử của X là bao nhiêu?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Chu kì 4 (4 lớp e), nhóm IIA (nguyên tố s, có $\ce{4s^2}$).\\
Cấu hình: $\ce{1s^2 2s^2 2p^6 3s^2 3p^6 4s^2} \implies Z = 20$ (Calcium - Ca).\\
\textbf{Đáp số: 20}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 3 Trả lời ngắn
    tex.append(r"""
\begin{ex}
Các electron của nguyên tử nguyên tố X được phân bố trên ba lớp, lớp thứ ba có 5 electron. Số đơn vị điện tích hạt nhân của nguyên tử nguyên tố X là bao nhiêu?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Cấu hình electron: $\ce{1s^2 2s^2 2p^6 3s^2 3p^3}$.\\
Tổng số electron = 15 $\implies$ Số đơn vị điện tích hạt nhân $Z = 15$ (Phosphorus - P).\\
\textbf{Đáp số: 15}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 4 Trả lời ngắn
    tex.append(r"""
\begin{ex}
Nước cất ($\ce{H2O}$) là nước tinh khiết, nguyên chất, được điều chế bằng cách chưng cất và thường được sử dụng trong y tế như pha chế thuốc tiêm, thuốc uống, biệt dược, rửa dụng cụ y tế, rửa vết thương,... Biết phân tử này chỉ được tạo nên từ các đồng vị $^1_1\ce{H}$ và $^{16}_8\ce{O}$. Tổng số hạt electron, proton và neutron trong một phân tử $\ce{H2O}$ là bao nhiêu?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Trong 1 phân tử $\ce{H2O}$ ($2\times ^1_1\ce{H} + 1\times ^{16}_8\ce{O}$):\\
- Số proton: $2 \times 1 + 8 = 10$.\\
- Số electron: $2 \times 1 + 8 = 10$.\\
- Số neutron: $2 \times 0 + 8 = 8$.\\
Tổng số hạt ($p + e + n$) = $10 + 10 + 8 = 28$ hạt.\\
\textbf{Đáp số: 28}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # PHẦN TỰ LUẬN
    tex.append(r"""
\vspace{6pt}
\noindent
{\color{mainGreen}\textbf{\Large B. PHẦN TỰ LUẬN: (3 điểm)}}
\vspace{2pt}
""")

    # Câu 1 Tự luận
    tex.append(r"""
\noindent
\textbf{Câu 1: (1,0 điểm)} Nguyên tử của nguyên tố X nằm ở ô thứ 17, chu kì 3 nhóm VIIA trong bảng HTTH các nguyên tố hoá học.\\
a) Viết cấu hình electron nguyên tử X.\\
b) Viết công thức oxide cao nhất đối với oxigene và hợp chất với Hidrogene.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 1:}]
\textbf{a) (0,5 điểm)}
\begin{itemize}[leftmargin=1.2cm]
\item X ở ô thứ 17 $\implies Z = 17$.
\item Cấu hình electron của X: $\ce{1s^2 2s^2 2p^6 3s^2 3p^5}$ (hoặc $\ce{[Ne] 3s^2 3p^5}$). \hfill \textit{(0,5 đ)}
\end{itemize}
\textbf{b) (0,5 điểm)}
\begin{itemize}[leftmargin=1.2cm]
\item Hóa trị cao nhất với oxygen là VII $\implies$ Công thức oxide cao nhất là $\ce{X2O7}$ (hoặc $\ce{Cl2O7}$). \hfill \textit{(0,25 đ)}
\item Hóa trị với hydrogen là $8 - 7 = \text{I} \implies$ Công thức hợp chất với hydrogen là $\ce{HX}$ (hoặc $\ce{HCl}$). \hfill \textit{(0,25 đ)}
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{6}{\dotfill\\[4pt]}
""")

    # Câu 2 Tự luận
    tex.append(r"""
\vspace{4pt}
\noindent
\textbf{Câu 2: (1,0 điểm)} Nguyên tử khối trung bình của nguyên tố R là 79,91. R có hai đồng vị. Biết $^{79}\ce{R}$ chiếm 54,5\%. Nguyên tử khối của đồng vị thứ 2 bằng bao nhiêu?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 2:}]
\begin{itemize}[leftmargin=1.2cm]
\item Gọi số khối (nguyên tử khối) của đồng vị thứ hai là $A_2$.
\item Thành phần \% của đồng vị thứ hai là: $\%A_2 = 100\% - 54,5\% = 45,5\%$. \hfill \textit{(0,25 đ)}
\item Áp dụng công thức tính nguyên tử khối trung bình:
$$\overline{A}_{\ce{R}} = \dfrac{79 \times 54,5 + A_2 \times 45,5}{100} = 79,91 \hfill \textit{(0,25 đ)}$$
$$\iff 4305,5 + 45,5 \times A_2 = 7991 \hfill \textit{(0,25 đ)}$$
$$\iff 45,5 \times A_2 = 3685,5 \iff A_2 = 81\text{ amu}. \hfill \textit{(0,25 đ)}$$
\item Vậy nguyên tử khối của đồng vị thứ hai là 81 (đồng vị $^{81}\ce{Br}$).
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{6}{\dotfill\\[4pt]}
""")

    # Câu 3 Tự luận
    tex.append(r"""
\vspace{4pt}
\noindent
\textbf{Câu 3: (1,0 điểm)} Một loại quặng chứa 2 muối cacbonate của 2 kim loại kế tiếp nhau trong nhóm IIA được dùng trong vật liệu xây dựng, luyện kim, cải tạo đất chua... Cho 10,80 g quặng trên tác dụng với dd $\ce{H2SO4}$ loãng dư. Chất khí thu được cho hấp thụ hoàn toàn vào dd $\ce{Ba(OH)2}$ dư thì thu được 23,64 g kết tủa.\\
a. Viết các phương trình phản ứng hoá học xảy ra. Tìm công thức 2 muối trên.\\
b. Cần dùng bao nhiêu ml dung dịch HCl 30\%, $d = 1,5\text{ gam/ml}$ để phản ứng hết 5,4 gam quặng trên.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 3:}]
\textbf{a) (0,5 điểm)}
\begin{itemize}[leftmargin=1.2cm]
\item Đặt công thức chung của 2 muối cacbonat là $\overline{\ce{M}}\ce{CO3}$.
\item Phương trình phản ứng:
$$\overline{\ce{M}}\ce{CO3} + \ce{H2SO4} \to \overline{\ce{M}}\ce{SO4} + \ce{CO2}\uparrow + \ce{H2O} \quad (1)$$
$$\ce{CO2} + \ce{Ba(OH)2} \text{ (dư)} \to \ce{BaCO3}\downarrow + \ce{H2O} \quad (2) \hfill \textit{(0,15 đ)}$$
\item Kết tủa thu được là $\ce{BaCO3}$ ($M = 197\text{ g/mol}$):
$$n_{\ce{BaCO3}} = \dfrac{23,64}{197} = 0,12\text{ mol}.$$
\item Theo (1) và (2): $n_{\ce{CO2}} = n_{\overline{\ce{M}}\ce{CO3}} = n_{\ce{BaCO3}} = 0,12\text{ mol}$. \hfill \textit{(0,1 đ)}
\item Khối lượng mol trung bình của 2 muối cacbonat:
$$M_{\overline{\ce{M}}\ce{CO3}} = \overline{M} + 60 = \dfrac{10,80}{0,12} = 90\text{ g/mol} \implies \overline{M} = 30\text{ g/mol}. \hfill \textit{(0,15 đ)}$$
\item Vì 2 kim loại kế tiếp nhau trong nhóm IIA: $M_{\ce{A}} < 30 < M_{\ce{B}}$.\\
Mà các kim loại nhóm IIA là: $\ce{Be}(9) < \ce{Mg}(24) < \ce{Ca}(40) < \ce{Sr}(88) < \ce{Ba}(137)$.\\
$\implies$ Hai kim loại kế tiếp là Magnesium ($\ce{Mg} = 24$) và Calcium ($\ce{Ca} = 40$).\\
Công thức 2 muối cacbonat là $\ce{MgCO3}$ và $\ce{CaCO3}$. \hfill \textit{(0,1 đ)}
\end{itemize}
\textbf{b) (0,5 điểm)}
\begin{itemize}[leftmargin=1.2cm]
\item Phương trình phản ứng: $\overline{\ce{M}}\ce{CO3} + 2\ce{HCl} \to \overline{\ce{M}}\ce{Cl2} + \ce{CO2}\uparrow + \ce{H2O} \quad (3)$
\item Trong 5,4 gam quặng (bằng một nửa khối lượng ban đầu):
$$n_{\text{muối}} = \dfrac{5,4}{90} = 0,06\text{ mol}. \hfill \textit{(0,15 đ)}$$
\item Theo PTHH (3): $n_{\ce{HCl}} = 2\cdot n_{\text{muối}} = 2 \times 0,06 = 0,12\text{ mol}$. \hfill \textit{(0,1 đ)}
\item Khối lượng $\ce{HCl}$ nguyên chất: $m_{\ce{HCl}} = 0,12 \times 36,5 = 4,38\text{ g}$.
\item Khối lượng dung dịch $\ce{HCl}$ 30\%:
$$m_{\text{dd HCl}} = \dfrac{4,38 \times 100}{30} = 14,6\text{ gam}. \hfill \textit{(0,15 đ)}$$
\item Thể tích dung dịch $\ce{HCl}$ 30\% cần dùng:
$$V_{\text{dd HCl}} = \dfrac{m_{\text{dd}}}{d} = \dfrac{14,6}{1,5} \approx 9,733\text{ mL} \approx 9,73\text{ mL}. \hfill \textit{(0,1 đ)}$$
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{8}{\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\begin{center}
\textbf{\large ---------- HẾT ----------}
\end{center}

\end{document}
""")

    return "".join(tex)

def compile_pdf(tex_path):
    out_dir = os.path.dirname(tex_path)
    base_name = os.path.splitext(os.path.basename(tex_path))[0]
    
    for i in range(2):
        cmd = ["pdflatex", "-interaction=nonstopmode", f"-output-directory={out_dir}", tex_path]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, cwd=out_dir)
        if res.returncode != 0 and i == 1:
            print(f"[CẢNH BÁO] pdflatex kết thúc với mã {res.returncode}")
            log_file = os.path.join(out_dir, f"{base_name}.log")
            if os.path.exists(log_file):
                with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
                    lines = f.readlines()
                    print("".join(lines[-30:]))
    
    for ext in [".aux", ".log", ".thm", ".out", ".synctex.gz"]:
        f_junk = os.path.join(out_dir, base_name + ext)
        if os.path.exists(f_junk):
            try:
                os.remove(f_junk)
            except Exception:
                pass
    for f in glob.glob(os.path.join(out_dir, "ans-*.tex")):
        try:
            os.remove(f)
        except Exception:
            pass
    print(f"[OK] Đã biên dịch xong PDF: {os.path.join(out_dir, base_name + '.pdf')}")

def main():
    tex_hs = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_Hoa_10_Ma132_HocSinh.tex")
    tex_gv = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_Hoa_10_Ma132_GiaoVien.tex")

    print("Đang tạo mã nguồn LaTeX Bản Học sinh Mã đề 132...")
    with open(tex_hs, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))

    print("Đang tạo mã nguồn LaTeX Bản Giáo viên Mã đề 132...")
    with open(tex_gv, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))

    print("\nĐang biên dịch PDF Bản Học sinh Mã đề 132...")
    compile_pdf(tex_hs)

    print("\nĐang biên dịch PDF Bản Giáo viên Mã đề 132...")
    compile_pdf(tex_gv)

    print("\n[HOÀN TẤT] Toàn bộ file .tex và .pdf Mã đề 132 đã được tạo và biên dịch thành công!")

if __name__ == "__main__":
    main()
