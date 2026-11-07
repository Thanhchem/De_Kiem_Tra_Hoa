# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Giữa Kỳ 1 Hóa Học 11 - THPT Chuyên Khoa Học Huế (Mã Đề 135)
Năm học 2026 - 2027
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_11_ChuyenKhoaHocHue_Ma135_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_11_ChuyenKhoaHocHue_Ma135_GiaoVien.pdf (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
\usepackage{multido}
\usepackage{tcolorbox}
\usepackage{colortbl}

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
\textbf{TRƯỜNG ĐẠI HỌC KHOA HỌC}\\
\textbf{TRƯỜNG THPT CHUYÊN KHOA HỌC HUẾ}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 03 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}ĐỀ KIỂM TRA GIỮA HỌC KỲ I - NĂM HỌC 2026 - 2027}\\
\textbf{Môn: Hóa học -- Lớp 11 (CHUNG)}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian giao đề)}\\
\textbf{Mã đề: 135}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ và tên học sinh:} \dotfill \textit{Lớp:} \dotfill \textit{Số báo danh:} \dotfill

\vspace{3pt}
\noindent
\textit{\footnotesize Cho nguyên tử khối: $\ce{H}=1; \ce{Li}=7; \ce{Be}=9; \ce{N}=14; \ce{O}=16; \ce{Na}=23; \ce{Mg}=24; \ce{Al}=27; \ce{P}=31; \ce{S}=32; \ce{Cl}=35,5; \ce{K}=39; \ce{Ca}=40; \ce{Ba}=137$.}

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
{\color{hfRed}\textbf{\large BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 135)}}
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
\textbf{Đ/A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} \\
\hline
\end{tabular}
\end{center}

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
\textbf{Câu 1} & Khí thải $\ce{NH3}$ \& muối ammonium & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{SAI} & \textbf{\color{hfRed}ĐÚNG} \\
\hline
\textbf{Câu 2} & Hệ cân bằng $\ce{CO2 + H2 <=> CO + H2O}$ & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{SAI} & \textbf{SAI} \\
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
\textbf{Đáp án} & \textbf{\color{hfRed}6,12} & \textbf{\color{hfRed}5} & \textbf{\color{hfRed}0,75} & \textbf{\color{hfRed}2} \\
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
\textbf{Câu 1} & 1,0 đ & PTHH điện li: $\ce{KOH -> K+ + OH-}$; $\ce{Al2(SO4)3 -> 2Al^3+ + 3SO4^2-}$; $\ce{CH3COOH <=> CH3COO- + H+}$; $\ce{(NH4)3PO4 -> 3NH4+ + PO4^3-}$. \\
\hline
\textbf{Câu 2} & 1,0 đ & a) 3 PTHH chu trình Ostwald điều chế $\ce{HNO3}$. \quad b) $m_{\text{tt}}(\ce{NH3}) = 33\,660,03\text{ tấn}$ (hoặc $\approx 33\,660\text{ tấn}$). \\
\hline
\textbf{Câu 3} & 1,0 đ & Dung dịch sau pư có $\ce{OH-}$ dư $0,032\text{ mol} \implies [\ce{OH-}] = 0,16\text{ M} \implies \text{pOH} \approx 0,796 \implies \text{pH} = 13,2$. \\
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
{\color{mainGreen}\textbf{\Large A. TRẮC NGHIỆM (7 điểm)}}

\vspace{3pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN I. (3 điểm) Câu trắc nghiệm nhiều phương án lựa chọn.}}\\
\textit{\small Học sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi học sinh chỉ chọn một phương án.}
\vspace{2pt}
""")

    # Câu 1
    tex.append(r"""
\begin{ex}
Có thể nhận biết muối ammonium bằng cách cho muối tác dụng với dung dịch kiềm thấy thoát ra một chất khí làm xanh quỳ tím ẩm. Chất khí đó là
\choice
{$\ce{NO2}$}
{$\ce{NO}$}
{\True $\ce{NH3}$}
{$\ce{H2}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Muối ammonium tác dụng với dung dịch kiềm sinh ra khí ammonia ($\ce{NH3}$) có mùi khai đặc trưng, làm xanh quỳ tím ẩm:
$$\ce{NH4+ + OH- -> NH3 ^ + H2O}$$
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2
    tex.append(r"""
\begin{ex}
Trong các dung dịch có cùng nồng độ $0,10\text{ mol/L}$ dưới đây, dung dịch nào dẫn điện kém nhất?
\choice
{\True $\ce{HF}$}
{$\ce{HI}$}
{$\ce{HCl}$}
{$\ce{HBr}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$\ce{HI}, \ce{HCl}, \ce{HBr}$ là các acid mạnh, phân li hoàn toàn thành ion (chất điện li mạnh).\\
$\ce{HF}$ là acid yếu, phân li thuận nghịch (chất điện li yếu): $\ce{HF <=> H+ + F-}$.\\
Ở cùng nồng độ $0,10\text{ mol/L}$, dung dịch $\ce{HF}$ có nồng độ ion tự do nhỏ nhất nên dẫn điện kém nhất.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 3
    tex.append(r"""
\begin{ex}
Acid trong dạ dày giúp kích hoạt enzyme pepsin, hỗ trợ tiêu hóa protein và bảo vệ cơ thể khỏi vi khuẩn có hại. Nếu trong dạ dày có $\text{pH} = 2$ thì nồng độ $\ce{H+}$ bằng bao nhiêu?
\choice
{$10\text{ M}$}
{$10^2\text{ M}$}
{\True $10^{-2}\text{ M}$}
{$2\text{ M}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Ta có: $\text{pH} = -\lg[\ce{H+}] \implies [\ce{H+}] = 10^{-\text{pH}} = 10^{-2}\text{ M}$.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 4
    tex.append(r"""
\begin{ex}
Nhận xét nào dưới đây về nitrogen \textbf{không đúng}?
\choice
{\True Khí nitrogen duy trì sự cháy}
{$\ce{N2}$ phản ứng với $\ce{O2}$ ở nhiệt độ rất cao}
{Khí nitrogen rất ít tan trong nước}
{$\ce{N2}$ vừa có tính oxi hóa vừa có tính khử}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Khí nitrogen ($\ce{N2}$) không duy trì sự cháy và không duy trì sự hô hấp. Do đó nhận xét A là không đúng.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 5
    tex.append(r"""
\begin{ex}
Phản ứng thuận nghịch là phản ứng
\choice
{có phương trình hoá học được biểu diễn bằng mũi tên một chiều}
{chỉ xảy ra theo một chiều nhất định}
{\True trong cùng điều kiện, phản ứng xảy ra theo hai chiều trái ngược nhau}
{xảy ra giữa hai chất khí}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Theo định nghĩa: Phản ứng thuận nghịch là phản ứng trong cùng điều kiện xác định, xảy ra đồng thời theo hai chiều trái ngược nhau.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 6
    tex.append(r"""
\begin{ex}
Các yếu tố ảnh hưởng đến cân bằng hoá học là :
\choice
{nồng độ, nhiệt độ và chất xúc tác}
{nồng độ, áp suất và diện tích bề mặt}
{\True nồng độ, nhiệt độ và áp suất}
{áp suất, nhiệt độ và chất xúc tác}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Theo nguyên lý Le Chatelier, các yếu tố làm chuyển dịch cân bằng hoá học gồm có: nồng độ, nhiệt độ và áp suất (đối với hệ phản ứng có chất khí). Chất xúc tác và diện tích bề mặt chỉ làm tăng tốc độ phản ứng mà không làm chuyển dịch cân bằng.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 7
    tex.append(r"""
\begin{ex}
Hành động nào của con người \textbf{không phải} nguyên nhân chính gây nên hiện tượng phú dưỡng?
\choice
{Sử dụng dư thừa lượng phân bón chứa ion $\ce{PO4^3-}$}
{\True Quá trình đánh bắt cá bằng chích điện và thuốc nổ của con người}
{Sử dụng dư thừa lượng phân bón chứa ion $\ce{NO2^-}, \ce{NO3^-}$}
{Nước thải chăn nuôi, nước thải công nghiệp chứa nhiều các nguyên tố dinh dưỡng như N, P}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Hiện tượng phú dưỡng là sự tích tụ quá mức các chất dinh dưỡng chứa nguyên tố N và P trong các thủy vực. Đánh bắt cá bằng kích điện hay thuốc nổ phá hoại hệ sinh vật nhưng không phải là nguồn xả thải dinh dưỡng N, P gây phú dưỡng.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 8
    tex.append(r"""
\begin{ex}
Trong phản ứng sau đây, những chất, ion nào đóng vai trò là acid theo thuyết Bronsted-Lowry?
$$\ce{H2S(aq) + H2O <=> HS^-(aq) + H3O^+(aq)}$$
\choice
{$\ce{H2S}$ và $\ce{HS^-}$}
{$\ce{H2O}$ và $\ce{H3O^+}$}
{$\ce{H2S}$ và $\ce{H2O}$}
{\True $\ce{H2S}$ và $\ce{H3O^+}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Theo thuyết Brønsted – Lowry, acid là chất cho proton ($\ce{H+}$):\\
- Chiều thuận: $\ce{H2S}$ nhường $\ce{H+}$ cho $\ce{H2O} \implies \ce{H2S}$ là acid.\\
- Chiều nghịch: $\ce{H3O+}$ nhường $\ce{H+}$ cho $\ce{HS-} \implies \ce{H3O+}$ là acid.\\
Vậy các acid là $\ce{H2S}$ và $\ce{H3O+}$.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 9
    tex.append(r"""
\begin{ex}
Nitrogen tương đối trơ về mặt hoá học ở nhiệt độ thường là do
\choice
{\True phân tử $\ce{N2}$ có liên kết ba với năng lượng liên kết lớn}
{nitrogen có độ âm điện lớn}
{phân tử $\ce{N2}$ có liên kết cộng hoá trị không phân cực}
{phân tử $\ce{N2}$ có liên kết ion}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phân tử $\ce{N2}$ có liên kết ba $\ce{N#N}$ với năng lượng liên kết rất lớn ($E_b = 945\text{ kJ/mol}$), rất khó bị bẻ gãy ở nhiệt độ thường nên nitrogen tương đối trơ về mặt hoá học.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 10
    tex.append(r"""
\begin{ex}
Khi tăng áp suất, những cân bằng hóa học nào dưới đây chuyển dịch theo chiều thuận?\\
(1) $\ce{H2(g) + Cl2(g) <=> 2HCl(g)}$\\
(2) $\ce{NH4Cl(s) <=> NH3(g) + HCl(g)}$\\
(3) $\ce{2H2O2(aq) <=> 2H2O(l) + O2(g)}$\\
(4) $\ce{3O2(g) <=> 2O3(g)}$
\choice
{(1) và (2)}
{(1) và (4)}
{(2) và (3)}
{\True Chỉ có (4)}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Khi tăng áp suất, cân bằng chuyển dịch theo chiều làm giảm số mol khí:\\
- (1) Vế trái $2\text{ mol khí}$, vế phải $2\text{ mol khí} \implies$ áp suất không làm chuyển dịch cân bằng.\\
- (2) Vế trái $0\text{ mol khí}$, vế phải $2\text{ mol khí} \implies$ tăng áp suất chuyển dịch theo chiều nghịch.\\
- (3) Vế trái $0\text{ mol khí}$, vế phải $1\text{ mol khí} \implies$ tăng áp suất chuyển dịch theo chiều nghịch.\\
- (4) Vế trái $3\text{ mol khí}$, vế phải $2\text{ mol khí} \implies$ số mol khí giảm, cân bằng chuyển dịch theo chiều thuận.\\
Vậy chỉ có (4) chuyển dịch theo chiều thuận.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 11
    tex.append(r"""
\begin{ex}
Dãy chất nào sau đây đều là chất điện li yếu?
\choice
{$\ce{H2S}, \ce{H2SO3}, \ce{H2SO4}$}
{$\ce{H2CO3}, \ce{H3PO4}, \ce{CH3COOH}, \ce{Ba(OH)2}$}
{\True $\ce{H2S}, \ce{CH3COOH}, \ce{HClO}$}
{$\ce{H2CO3}, \ce{H2SO3}, \ce{HClO}, \ce{Al2(SO4)3}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- A loại $\ce{H2SO4}$ (acid mạnh).\\
- B loại $\ce{Ba(OH)2}$ (base mạnh).\\
- D loại $\ce{Al2(SO4)3}$ (muối tan, điện li mạnh).\\
- C gồm $\ce{H2S}, \ce{CH3COOH}, \ce{HClO}$ đều là các acid yếu nên là chất điện li yếu.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 12
    tex.append(r"""
\begin{ex}
Dung dịch nào sau đây có $\text{pH} > 7$?
\choice
{\True $\ce{NaOH}$}
{$\ce{H2SO4}$}
{$\ce{NaCl}$}
{$\ce{HNO3}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$\ce{NaOH}$ là base mạnh tạo môi trường kiềm có $\text{pH} > 7$.\\
$\ce{H2SO4}, \ce{HNO3}$ có môi trường acid ($\text{pH} < 7$).\\
$\ce{NaCl}$ có môi trường trung tính ($\text{pH} = 7$).\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN II
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\Large PHẦN II (2 điểm). Câu trắc nghiệm đúng sai.}}\\
\textit{\small Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai.}
\vspace{2pt}
""")

    # Câu 1 Phần II
    tex.append(r"""
\begin{ex}
Trong khí thải của quy trình sản xuất thuốc trừ sâu, phân bón hóa học có lẫn khí $\ce{NH3}$. Khí này rất độc đối với sức khoẻ của con người và gây ô nhiễm môi trường. Con người hít phải khí này với lượng lớn sẽ gây ngộ độc: ho, đau ngực (nặng: đau thắt ngực, khó thở, thở nhanh, thở khò khè; chảy nước mắt và bỏng mắt, mù mắt, đau họng nặng, đau miệng; mạch nhanh, yếu, sốc; lẫn lộn, đi lại khó khăn, chóng mặt, thiếu sự phối hợp, bồn chồn, ngẩn ngơ).
\choiceTF
{\True Để xử lí $\ce{NH3}$ lẫn trong khí thải, người ta có thể dẫn khí thải qua một bể lọc chứa dung dịch $\ce{HCl}$}
{Tất cả các muối của $\ce{NH3}$ đều bền với nhiệt}
{$\ce{NH3}$ có tính base mạnh và tính khử mạnh}
{\True Đưa hai đũa thuỷ tinh có nhúng $\ce{NH3}$ đặc và $\ce{HCl}$ đặc lại gần thấy khói màu trắng xuất hiện}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) \textbf{Đúng}: $\ce{NH3}$ có tính base, tan tốt và tác dụng với dung dịch acid như $\ce{HCl}$ tạo muối $\ce{NH4Cl}$ giữ lại trong dung dịch, giúp khử sạch khí độc.\\
b) \textbf{Sai}: Tất cả các muối ammonium đều kém bền với nhiệt và dễ bị nhiệt phân.\\
c) \textbf{Sai}: $\ce{NH3}$ chỉ có tính base yếu (không phải base mạnh).\\
d) \textbf{Đúng}: Khói trắng là các hạt tinh thể $\ce{NH4Cl}$ lơ lửng: $\ce{NH3(g) + HCl(g) -> NH4Cl(s)}$.\\
$\implies$ Kết luận: a -- Đ, b -- S, c -- S, d -- Đ.
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần II
    tex.append(r"""
\begin{ex}
Trong bình kín có hệ cân bằng hóa học sau:
$$\ce{CO2(g) + H2(g) <=> CO(g) + H2O(g)};\quad \Delta_{\mathrm{r}}H^{\circ}_{298} > 0$$
Cho các phát biểu về sự tác động đến hệ cân bằng dưới đây, hãy xét tính đúng sai của các phát biểu đó.
\choiceTF
{\True Tăng nhiệt độ của hệ thì cân bằng chuyển dịch theo chiều thuận}
{Thêm chất xúc tác thì cân bằng chuyển dịch theo chiều thuận}
{Giảm áp suất chung của hệ thì cân bằng chuyển dịch theo chiều nghịch}
{Nếu nồng độ ban đầu $\ce{CO2}$ và $\ce{H2}$ đều là $0,03\text{ M}$; tại cân bằng $[\ce{CO}] = 0,018\text{ M}$ thì giá trị $K_c$ là $1,5$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) \textbf{Đúng}: $\Delta_{\mathrm{r}}H^{\circ}_{298} > 0$ là phản ứng thu nhiệt. Khi tăng nhiệt độ, cân bằng chuyển dịch theo chiều thu nhiệt (chiều thuận).\\
b) \textbf{Sai}: Chất xúc tác chỉ làm tăng tốc độ phản ứng, không làm chuyển dịch cân bằng.\\
c) \textbf{Sai}: Số mol khí ở 2 vế bằng nhau ($1+1=2$) nên sự thay đổi áp suất không làm chuyển dịch cân bằng.\\
d) \textbf{Sai}:\\
Tại cân bằng: $[\ce{CO}] = [\ce{H2O}] = 0,018\text{ M}$, $[\ce{CO2}] = [\ce{H2}] = 0,03 - 0,018 = 0,012\text{ M}$.\\
$K_c = \dfrac{[\ce{CO}][\ce{H2O}]}{[\ce{CO2}][\ce{H2}]} = \dfrac{0,018^2}{0,012^2} = 1,5^2 = 2,25 \ne 1,5$.\\
$\implies$ Kết luận: a -- Đ, b -- S, c -- S, d -- S.
}
""")
    tex.append(r"\end{ex}")

    # PHẦN III
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\Large PHẦN III (2 điểm). Câu trắc nghiệm trả lời ngắn.}}\\
\textit{\small Học sinh trả lời từ câu 1 đến câu 4.}
\vspace{2pt}
""")

    # Câu 1 Phần III
    tex.append(r"""
\begin{ex}
Trong công nghiệp ammonia được tổng hợp từ nitrogen và hydrogen. Cho $14,874\text{ lít }\ce{N2}$ (đkc) tác dụng với lượng dư khí $\ce{H2}$. Biết hiệu suất của phản ứng là $30\%$, khối lượng $\ce{NH3}$ tạo thành là bao nhiêu gam? \textit{(Kết quả làm tròn đến hàng phần trăm)}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$n_{\ce{N2}} = \dfrac{14,874}{24,79} = 0,6\text{ mol}$.\\
PTHH: $\ce{N2 + 3H2 <=> 2NH3}$\\
Theo lí thuyết: $n_{\ce{NH3(lt)}} = 2 \times 0,6 = 1,2\text{ mol}$.\\
Do $H = 30\% \implies n_{\ce{NH3(tt)}} = 1,2 \times 0,3 = 0,36\text{ mol}$.\\
Khối lượng $\ce{NH3}$: $m = 0,36 \times 17 = 6,12\text{ gam}$.\\
\textbf{Đáp số: 6,12}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần III
    tex.append(r"""
\begin{ex}
Cho các chất dưới đây: $\ce{HClO4}, \ce{HClO}, \ce{HF}, \ce{HNO3}, \ce{H2S}, \ce{H2SO3}, \ce{NaOH}, \ce{NaCl}, \ce{CuSO4}, \ce{CH3COOH}$. Số chất thuộc loại chất điện li mạnh là bao nhiêu?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Các chất điện li mạnh gồm 5 chất:\\
(1) $\ce{HClO4}$ (acid rất mạnh);\\
(2) $\ce{HNO3}$ (acid mạnh);\\
(3) $\ce{NaOH}$ (base mạnh);\\
(4) $\ce{NaCl}$ (muối tan);\\
(5) $\ce{CuSO4}$ (muối tan).\\
Các chất còn lại ($\ce{HClO}, \ce{HF}, \ce{H2S}, \ce{H2SO3}, \ce{CH3COOH}$) đều là acid yếu $\implies$ điện li yếu.\\
\textbf{Đáp số: 5}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 3 Phần III
    tex.append(r"""
\begin{ex}
Thực hiện thí nghiệm chuẩn độ dung dịch $\ce{HCl}$ chưa biết nồng độ bằng dung dịch chuẩn $\ce{NaOH}\ 0,5\text{M}$ và thuốc thử phenolphthalein. Cho $10,00\text{ mL}$ dung dịch $\ce{HCl}$ vào bình tam giác cùng với 1 -- 2 giọt thuốc thử. Thể tích dung dịch $\ce{NaOH}$ ở các lần chuẩn độ được ghi lại như sau:
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Lần 1} & \textbf{Lần 2} & \textbf{Lần 3} \\
\hline
$15,1\text{ mL}$ & $14,7\text{ mL}$ & $15,2\text{ mL}$ \\
\hline
\end{tabular}
\end{center}
Dung dịch $\ce{HCl}$ có nồng độ bằng bao nhiêu mol/L? \textit{(Kết quả làm tròn đến hàng phần trăm)}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Thể tích $\ce{NaOH}$ trung bình:\\
$\overline{V}_{\ce{NaOH}} = \dfrac{15,1 + 14,7 + 15,2}{3} = 15,00\text{ mL}$.\\
Phương trình chuẩn độ: $\ce{HCl + NaOH -> NaCl + H2O}$.\\
Tại điểm tương đương: $C_{\ce{HCl}} \times V_{\ce{HCl}} = C_{\ce{NaOH}} \times \overline{V}_{\ce{NaOH}}$\\
$\implies C_{\ce{HCl}} = \dfrac{0,5 \times 15,00}{10,00} = 0,75\text{ M (hay mol/L)}$.\\
\textbf{Đáp số: 0,75}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 4 Phần III
    tex.append(r"""
\begin{ex}
Cho phản ứng: $\ce{CO(g) + H2O(g) <=> CO2(g) + H2(g)}\quad \Delta_{\mathrm{r}}H^{\circ}_{298} = -41\text{ kJ}$.\\
Các phát biểu sau đây, có bao nhiêu phát biểu \textbf{đúng} về phản ứng trên?\\
(1) Phản ứng theo chiều thuận là phản ứng thu nhiệt.\\
(2) Tăng nhiệt độ cân bằng chuyển dịch theo chiều nghịch.\\
(3) Thêm khí $\ce{H2}$ vào hệ cân bằng chuyển dịch theo chiều nghịch.\\
(4) Tăng áp suất chung bằng cách nén cho thể tích của hệ giảm xuống cân bằng chuyển dịch theo chiều thuận.
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- (1) Sai vì $\Delta_{\mathrm{r}}H^{\circ}_{298} = -41\text{ kJ} < 0 \implies$ chiều thuận tỏa nhiệt.\\
- (2) Đúng vì tăng nhiệt độ thì chuyển dịch theo chiều thu nhiệt (chiều nghịch).\\
- (3) Đúng vì thêm $\ce{H2}$ (sản phẩm vế phải) cân bằng chuyển dịch sang trái (chiều nghịch).\\
- (4) Sai vì 2 vế có cùng $2\text{ mol khí}$ nên áp suất không làm chuyển dịch cân bằng.\\
$\implies$ Có 2 phát biểu đúng là (2) và (3).\\
\textbf{Đáp số: 2}
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
{\color{mainGreen}\textbf{\Large B. TỰ LUẬN (3 điểm)}}
\vspace{2pt}
""")

    # Câu 1 Tự luận
    tex.append(r"""
\noindent
\textbf{Câu 1. (1,0 điểm)} Viết phương trình điện li của các chất sau trong nước: $\ce{KOH}$, $\ce{Al2(SO4)3}$, $\ce{CH3COOH}$, $\ce{(NH4)3PO4}$.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 1:}]
\begin{itemize}[leftmargin=1.2cm]
\item $\ce{KOH}$ là base mạnh, chất điện li mạnh:
$$\ce{KOH -> K+ + OH-} \hfill \textit{(0,25 đ)}$$
\item $\ce{Al2(SO4)3}$ là muối tan, chất điện li mạnh:
$$\ce{Al2(SO4)3 -> 2Al^3+ + 3SO4^2-} \hfill \textit{(0,25 đ)}$$
\item $\ce{CH3COOH}$ là acid yếu, chất điện li yếu, phân li thuận nghịch:
$$\ce{CH3COOH <=> CH3COO- + H+} \quad (\text{hoặc } \ce{CH3COOH + H2O <=> CH3COO- + H3O+}) \hfill \textit{(0,25 đ)}$$
\item $\ce{(NH4)3PO4}$ là muối tan, chất điện li mạnh:
$$\ce{(NH4)3PO4 -> 3NH4+ + PO4^3-} \hfill \textit{(0,25 đ)}$$
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
\textbf{Câu 2. (1,0 điểm)} Trong công nghiệp, người ta sản xuất nitric acid ($\ce{HNO3}$) từ ammonia theo sơ đồ chuyển hoá sau:
$$\ce{NH3 ->[+O2, t^\circ, xt] NO ->[+O2] NO2 ->[+O2 + H2O] HNO3}$$
a) Viết các phương trình hoá học xảy ra.\\
b) Để điều chế $200\,000\text{ tấn}$ nitric acid có nồng độ $60\%$ cần dùng bao nhiêu tấn ammonia? Biết rằng hiệu suất của quá trình sản xuất nitric acid theo sơ đồ trên là $96,2\%$.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 2:}]
\textbf{a) (0,5 điểm)} Các phương trình hoá học xảy ra:
\begin{align*}
&\text{(1) } \ce{4NH3 + 5O2 ->[\text{Pt},\, 850 - 900^{\circ}\text{C}] 4NO + 6H2O} \hfill \textit{(0,2 đ)}\\
&\text{(2) } \ce{2NO + O2 -> 2NO2} \hfill \textit{(0,15 đ)}\\
&\text{(3) } \ce{4NO2 + O2 + 2H2O -> 4HNO3} \hfill \textit{(0,15 đ)}
\end{align*}
\textbf{b) (0,5 điểm)} Tính lượng ammonia cần dùng:
\begin{itemize}[leftmargin=1.2cm]
\item Khối lượng $\ce{HNO3}$ nguyên chất cần sản xuất:
$$m_{\ce{HNO3}} = 200\,000 \times 60\% = 120\,000\text{ tấn}. \hfill \textit{(0,15 đ)}$$
\item Sơ đồ bảo toàn nguyên tố $\ce{N}$:
$$\underset{17\text{ tấn}}{\ce{NH3}} \longrightarrow \underset{63\text{ tấn}}{\ce{HNO3}}$$
Theo lý thuyết (hiệu suất 100\%):
$$m_{\text{lt}}(\ce{NH3}) = \dfrac{120\,000 \times 17}{63} \approx 32\,380,95\text{ tấn}. \hfill \textit{(0,15 đ)}$$
\item Do hiệu suất toàn bộ quá trình là $H = 96,2\%$, khối lượng $\ce{NH3}$ thực tế cần dùng là:
$$m_{\text{tt}}(\ce{NH3}) = \dfrac{m_{\text{lt}}}{96,2\%} = \dfrac{32\,380,9524}{0,962} \approx 33\,660,03\text{ tấn} \quad (\text{hoặc } \approx 33\,660\text{ tấn}). \hfill \textit{(0,2 đ)}$$
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{8}{\dotfill\\[4pt]}
""")

    # Câu 3 Tự luận
    tex.append(r"""
\vspace{4pt}
\noindent
\textbf{Câu 3. (1,0 điểm)} Trộn $80\text{ mL}$ dung dịch $\ce{HCl}\ 0,5\text{M}$ với $120\text{ mL}$ dung dịch $\ce{NaOH}\ 0,6\text{M}$. Tính pH của dung dịch thu được. \textit{(Kết quả làm tròn đến hàng phần mười)}.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 3:}]
\begin{itemize}[leftmargin=1.2cm]
\item Thể tích dung dịch sau khi trộn: $V_{\text{dd}} = 80 + 120 = 200\text{ mL} = 0,2\text{ lít}$. \hfill \textit{(0,1 đ)}
\item Số mol các chất ban đầu:
\begin{align*}
n_{\ce{H+}} &= n_{\ce{HCl}} = 0,08 \times 0,5 = 0,04\text{ mol}. \hfill \textit{(0,2 đ)}\\
n_{\ce{OH-}} &= n_{\ce{NaOH}} = 0,12 \times 0,6 = 0,072\text{ mol}. \hfill \textit{(0,2 đ)}
\end{align*}
\item Phương trình ion thu gọn:
$$\ce{H+ + OH- -> H2O}$$
Do $n_{\ce{OH-}} > n_{\ce{H+}}$ nên sau phản ứng $\ce{OH-}$ còn dư:
$$n_{\ce{OH^-}\text{ (dư)}} = 0,072 - 0,04 = 0,032\text{ mol}. \hfill \textit{(0,2 đ)}$$
\item Nồng độ $\ce{OH-}$ dư trong dung dịch thu được:
$$[\ce{OH-}] = \dfrac{0,032}{0,2} = 0,16\text{ M}. \hfill \textit{(0,1 đ)}$$
\item Tính pH dung dịch:
$$\text{pOH} = -\lg[\ce{OH-}] = -\lg(0,16) \approx 0,796 \implies \text{pH} = 14 - \text{pOH} = 14 - 0,796 = 13,204 \approx 13,2. \hfill \textit{(0,2 đ)}$$
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{6}{\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\begin{center}
\textbf{\large ---------- HẾT ----------}\\
\textit{\small - Học sinh không được sử dụng tài liệu; Giám thị không giải thích gì thêm. -}
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
    tex_hs = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_11_ChuyenKhoaHocHue_Ma135_HocSinh.tex")
    tex_gv = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_11_ChuyenKhoaHocHue_Ma135_GiaoVien.tex")

    with open(tex_hs, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))
    print(f"Đã tạo file TeX Học sinh: {tex_hs}")

    with open(tex_gv, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))
    print(f"Đã tạo file TeX Giáo viên: {tex_gv}")

    print("\n--- Bắt đầu biên dịch PDF Bản Học sinh ---")
    compile_pdf(tex_hs)

    print("\n--- Bắt đầu biên dịch PDF Bản Giáo viên ---")
    compile_pdf(tex_gv)


if __name__ == "__main__":
    main()
