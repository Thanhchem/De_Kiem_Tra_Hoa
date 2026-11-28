# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Cuối Học Kỳ I Hóa Học 12 - THPT Nguyễn Chí Thanh (TP Huế)
Mã đề: 1001 - Năm học 2025 - 2026
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_CK1_Hoa_12_NguyenChiThanh_Ma1001_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_CK1_Hoa_12_NguyenChiThanh_Ma1001_GiaoVien.pdf (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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

SAN_PHAM_DIR = os.path.abspath("San_Pham")
DIR_IMAGES = os.path.join(SAN_PHAM_DIR, "images_de1001")

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
\usepackage{colortbl}
\usepackage{tikz}
\usepackage{fancyhdr}
\usepackage{enumitem}
\usepackage{multido}
\usepackage{tcolorbox}

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
\textbf{SỞ GD VÀ ĐT THÀNH PHỐ HUẾ}\\
\textbf{TRƯỜNG THPT NGUYỄN CHÍ THANH}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(gồm có: 04 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}ĐỀ KIỂM TRA CUỐI HỌC KÌ I}\\
\textbf{NĂM HỌC 2025 -- 2026}\\
\textbf{MÔN: HÓA HỌC -- LỚP 12}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{MÃ ĐỀ: 1001}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ và tên:} \dotfill \textit{Số báo danh:} \dotfill

\vspace{3pt}
\noindent
\textit{\footnotesize Cho biết nguyên tử khối của các nguyên tố (amu): $\ce{C}=12,\ \ce{H}=1,\ \ce{O}=16,\ \ce{Cu}=64,\ \ce{Ag}=108,\ \ce{Cl}=35,5,\ \ce{Na}=23,\ \ce{N}=14$.}

\vspace{3pt}
\hrule color{hfGreen} height 0.5pt
""")

    if is_teacher:
        # BẢNG ĐÁP ÁN RÚT GỌN CHO GIÁO VIÊN
        tex.append(r"""
\vspace{6pt}
\begin{center}
{\Large\textbf{\color{hfRed}BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN -- MÃ ĐỀ 1001)}}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}1. Phần I: Trắc nghiệm nhiều lựa chọn (4,0 điểm -- 0,25 đ/câu)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\setlength{\tabcolsep}{3.5pt}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} \\
\hline
\textbf{Đ/A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|p{7.5cm}|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Nội dung câu hỏi} & \textbf{Lệnh a} & \textbf{Lệnh b} & \textbf{Lệnh c} & \textbf{Lệnh d} \\
\hline
\textbf{Câu 1} & Thế điện cực chuẩn và pin điện hóa Fe -- Cu & \textbf{\color{gray}S} & \textbf{\color{gray}S} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}Đ} \\
\hline
\textbf{Câu 2} & Polymer tơ capron từ caprolactam & \textbf{\color{gray}S} & \textbf{\color{hfRed}Đ} & \textbf{\color{gray}S} & \textbf{\color{gray}S} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}3. Phần III: Trắc nghiệm Trả lời ngắn (1,0 điểm -- 0,25 đ/câu)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}74,1\%} & \textbf{\color{hfRed}4} & \textbf{\color{hfRed}3} & \textbf{\color{hfRed}165} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}4. Tóm tắt đáp số Phần IV: Tự luận (3,0 điểm -- 1,0 đ/câu)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|p{11.5cm}|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Điểm} & \textbf{Đáp số cốt lõi} \\
\hline
\textbf{Câu 1} & 1,0 đ & a) CTCT: $\ce{H2N-CH2-COOH}$, tên IUPAC: aminoethanoic acid; b) PTHH chứng minh tính lưỡng tính (+ $\ce{HCl}$, + $\ce{NaOH}$). \\
\hline
\textbf{Câu 2} & 1,0 đ & $n(\ce{CH4}\text{ lt}) = 32000\text{ mol} \Rightarrow V_\text{tt}(\ce{CH4}) = 991,6\text{ m}^3 \Rightarrow V(\text{khí thiên nhiên}) \approx \mathbf{1043,8\text{ m}^3}$. \\
\hline
\textbf{Câu 3} & 1,0 đ & a) Muối sodium/potassium của acid béo cao; b) $x = 0,5;\ y = 0,02 \Rightarrow m(\text{muối}) = 893,24\text{ g} \Rightarrow b \approx \mathbf{14\text{ bánh}}$. \\
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
{\color{mainGreen}\textbf{\Large A. PHẦN TRẮC NGHIỆM}}

\vspace{3pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN 1: TRẮC NGHIỆM NHIỀU LỰA CHỌN (4 điểm – 16 câu).}}\\
\textit{\small Thí sinh trả lời câu hỏi từ câu 1 đến câu 16, mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{2pt}
""")

    # CÂU 1 -> 16 PHẦN I
    # Câu 1
    tex.append(r"""
\begin{ex}
Methyl acetate có công thức cấu tạo nào sau đây?
\choice
{$\ce{CH3COOC3H7}$}
{$\ce{HCOOC2H5}$}
{\True $\ce{CH3COOCH3}$}
{$\ce{CH3COOC2H5}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Methyl acetate được tạo từ acetic acid ($\ce{CH3COOH}$) và methyl alcohol ($\ce{CH3OH}$), có công thức cấu tạo là $\ce{CH3COOCH3}$.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2
    tex.append(r"""
\begin{ex}
Amine nào sau đây là amine bậc hai?
\choice
{$\ce{(CH3)3N}$}
{$\ce{CH5N}$}
{\True $\ce{(CH3)2NH}$}
{$\ce{C2H7NH2}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Amine bậc hai có nguyên tử nitrogen liên kết trực tiếp với 2 gốc hydrocarbon: $\ce{(CH3)2NH}$ (dimethylamine).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 3
    tex.append(r"""
\begin{ex}
Khi nấu món canh từ thịt cua, tôm, tép có nhiều mảng thịt đóng rắn lại. Hiện tượng trên gây ra bởi tính chất nào sau đây?
\choice
{Sự thủy phân protein bởi nhiệt độ}
{Sự đông tụ protein bởi sự thay đổi pH}
{\True Sự đông tụ protein bởi nhiệt độ}
{Kết tủa carbonate của các chất khoáng có trong vỏ}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Dưới tác dụng của nhiệt độ khi đun sôi, cấu trúc không gian của các phân tử protein hòa tan bị biến tính và vón kết lại thành từng mảng đóng rắn nổi lên (hiện tượng đông tụ protein bởi nhiệt độ).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 4
    tex.append(r"""
\begin{ex}
Cho sức điện động chuẩn của các pin điện hoá sau:
$E^\circ_{\text{pin}(\text{T-X})} = 2,46\text{ V}$; $E^\circ_{\text{pin}(\text{T-Y})} = 2,00\text{ V}$ (với X, Y, T là ba kim loại). Dãy các kim loại được sắp xếp theo chiều tăng dần tính khử là:
\choice
{\True $\text{X} < \text{Y} < \text{T}$}
{$\text{Y} < \text{X} < \text{T}$}
{$\text{T} < \text{Y} < \text{X}$}
{$\text{T} < \text{X} < \text{Y}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Trong pin điện hóa T-X và T-Y, kim loại T đóng vai trò là anode (cực âm) có tính khử mạnh nhất. Ta có:\\
$E^\circ_{\text{pin}(\text{T-X})} = E^\circ(\text{X}^{n+}/\text{X}) - E^\circ(\text{T}^{m+}/\text{T}) = 2,46\text{ V} \Rightarrow E^\circ(\text{X}) = E^\circ(\text{T}) + 2,46\text{ V}$\\
$E^\circ_{\text{pin}(\text{T-Y})} = E^\circ(\text{Y}^{k+}/\text{Y}) - E^\circ(\text{T}^{m+}/\text{T}) = 2,00\text{ V} \Rightarrow E^\circ(\text{Y}) = E^\circ(\text{T}) + 2,00\text{ V}$\\
$\Rightarrow E^\circ(\text{X}) > E^\circ(\text{Y}) > E^\circ(\text{T})$.\\
Thế điện cực chuẩn $E^\circ$ càng lớn thì tính khử của kim loại càng yếu. Do đó chiều tăng dần tính khử là: $\text{X} < \text{Y} < \text{T}$.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 5
    tex.append(r"""
\begin{ex}
Chất nào sau đây thuộc loại monosaccharide?
\choice
{Cellulose}
{Saccharose}
{Tinh bột}
{\True Glucose}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Glucose (và fructose) là các monosaccharide (đường đơn). Saccharose là disaccharide; tinh bột và cellulose là polysaccharide.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 6
    tex.append(r"""
\begin{ex}
Tinh bột thuộc loại polysaccharide, có nhiều trong các loại ngũ cốc như hạt lúa, hạt ngô, củ khoai, củ sắn, .... Công thức phân tử của tinh bột là:
\choice
{$\ce{C12H22O11}$}
{$\ce{C6H12O6}$}
{\True $\ce{(C6H10O5)_n}$}
{$\ce{C2H4O2}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Tinh bột có công thức phân tử là $\ce{(C6H10O5)_n}$.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 7
    path_c7 = os.path.join(DIR_IMAGES, "cau7_amino_acid_forms.png").replace("\\", "/")
    tex_c7 = r"""
\begin{ex}
Dạng ion chủ yếu nào của amino acid tồn tại trong môi trường acid mạnh (pH khoảng 1-2)?
\begin{center}
\includegraphics[width=11.5cm]{__PATH_C7__}
\end{center}
\choice
{Dạng anion $\ce{H2N-CH2-COO-}$}
{Dạng ion lưỡng cực $\ce{^{+}H3N-CH2-COO-}$}
{Dạng anion $\ce{CH3-CH(NH2)-COO-}$}
{\True Dạng cation $\ce{CH3-CH(^{+}NH3)-COOH}$}
""".replace("__PATH_C7__", path_c7)
    tex.append(tex_c7)
    if is_teacher:
        tex.append(r"""
\loigiai{
Trong môi trường acid mạnh ($\text{pH} \approx 1-2$), nồng độ $\ce{H+}$ rất cao, nhóm $-\ce{COO-}$ nhận $\ce{H+}$ trở thành $-\ce{COOH}$, nhóm $-\ce{NH2}$ nhận $\ce{H+}$ trở thành $-\ce{NH3+}$. Do đó amino acid tồn tại chủ yếu ở dạng cation $\ce{^{+}H3N-CH(R)-COOH}$ (hình D).\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 8
    tex.append(r"""
\begin{ex}
Sản phẩm thu được khi điện phân dung dịch KCl (điện cực trơ, có màng ngăn) gồm:
\choice
{K và $\ce{Cl2}$}
{K, $\ce{H2}$ và $\ce{Cl2}$}
{\True KOH, $\ce{H2}$ và $\ce{Cl2}$}
{KOH, $\ce{O2}$ và HCl}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phương trình điện phân: $2\ce{KCl} + 2\ce{H2O} \xrightarrow{\text{đpdd, cmn}} 2\ce{KOH} + \ce{H2} + \ce{Cl2}$.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 9
    tex.append(r"""
\begin{ex}
Saccharose được cấu tạo từ:
\choice
{hai đơn vị $\alpha$-glucose qua liên kết $\alpha$-1,4-glycoside}
{một đơn vị $\alpha$-glucose và một đơn vị $\beta$-fructose qua liên kết $\alpha$-1,4-glycoside}
{\True một đơn vị $\alpha$-glucose và một đơn vị $\beta$-fructose qua liên kết $\alpha$-1,2-glycoside}
{hai đơn vị $\beta$-fructose qua liên kết $\beta$-1,4-glycoside}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phân tử saccharose được cấu tạo từ một gốc $\alpha$-glucose và một gốc $\beta$-fructose liên kết với nhau qua nguyên tử oxygen giữa C1 của $\alpha$-glucose và C2 của $\beta$-fructose (liên kết $\alpha$-1,2-glycosidic).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 10
    path_c10 = os.path.join(DIR_IMAGES, "pin_galvani_de1001.png").replace("\\", "/")
    tex_c10 = r"""
\begin{ex}
Một pin điện hoá có điện cực Zn nhúng trong dung dịch $\ce{ZnSO4}$ và điện cực Cu nhúng trong dung dịch $\ce{CuSO4}$. Sau một thời gian pin đó phóng điện thì:
\begin{center}
\includegraphics[width=7.5cm]{__PATH_C10__}
\end{center}
\choice
{khối lượng cả hai điện cực Zn và Cu đều tăng}
{khối lượng cả hai điện cực Zn và Cu đều giảm}
{khối lượng điện cực Zn tăng còn khối lượng điện cực Cu giảm}
{\True khối lượng điện cực Zn giảm còn khối lượng điện cực Cu tăng}
""".replace("__PATH_C10__", path_c10)
    tex.append(tex_c10)
    if is_teacher:
        tex.append(r"""
\loigiai{
Tại cực âm anode Zn: Zn bị tan ra tạo $\ce{Zn^2+}$ nên khối lượng thanh Zn giảm; tại cực dương cathode Cu: $\ce{Cu^2+}$ bị khử tạo Cu bám vào điện cực nên khối lượng thanh Cu tăng.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 11
    tex.append(r"""
\begin{ex}
Kết quả phân tích nguyên tố của một $\alpha$-amino acid X như sau: $\%C=46,60\%$; $\%H=8,74\%$; $\%N=13,59\%$ (về khối lượng); còn lại là oxygen. Bằng phổ khối lượng (MS), xác định được phân tử khối của X bằng 103. Số công thức cấu tạo thỏa mãn của X là:
\choice
{4}
{3}
{\True 2}
{1}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$\%O = 31,07\% \Rightarrow n_\ce{C} = 4;\ n_\ce{H} = 9;\ n_\ce{N} = 1;\ n_\ce{O} = 2 \Rightarrow \ce{C4H9NO2}$ ($M=103$).\\
Các đồng phân cấu tạo là $\alpha$-amino acid (nhóm $-\ce{NH2}$ ở C số 2 kề nhóm $-\ce{COOH}$):\\
(1) $\ce{CH3-CH2-CH(NH2)-COOH}$ (2-aminobutanoic acid)\\
(2) $\ce{(CH3)2C(NH2)-COOH}$ (2-amino-2-methylpropanoic acid)\\
$\Rightarrow$ Có \textbf{2} công thức cấu tạo thỏa mãn.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 12
    path_c12 = os.path.join(DIR_IMAGES, "linoleic_acid.png").replace("\\", "/")
    tex_c12 = r"""
\begin{ex}
Linoleic acid có công thức khung phân tử như bên dưới thuộc loại acid béo nào sau đây?
\begin{center}
\includegraphics[width=9.5cm]{__PATH_C12__}
\end{center}
\choice
{Omega-3}
{\True Omega-6}
{Omega-9}
{Omega-12}
""".replace("__PATH_C12__", path_c12)
    tex.append(tex_c12)
    if is_teacher:
        tex.append(r"""
\loigiai{
Đếm từ nguyên tử carbon ở nhóm methyl cuối mạch ($-\ce{CH3}$, đầu $\omega$): liên kết đôi $\ce{C=C}$ đầu tiên xuất hiện ở vị trí carbon số 6, do đó linoleic acid thuộc loại acid béo Omega-6.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 13
    tex.append(r"""
\begin{ex}
Trong quá trình điện phân $\ce{MgCl2}$ nóng chảy với các điện cực trơ, ở cathode xảy ra quá trình:
\choice
{khử ion $\ce{Cl-}$}
{oxi hóa ion $\ce{Cl-}$}
{\True khử ion $\ce{Mg^2+}$}
{oxi hóa ion $\ce{Mg^2+}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Ở cathode (cực âm) xảy ra quá trình khử cation: $\ce{Mg^2+} + 2\ce{e} \to \ce{Mg}$.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 14
    tex.append(r"""
\begin{ex}
Polymer nào sau đây được tổng hợp bằng phản ứng trùng hợp?
\choice
{tinh bột}
{tơ visco}
{tơ nylon-6}
{\True cao su buna}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Cao su buna được tổng hợp bằng phản ứng trùng hợp buta-1,3-diene: $n\,\ce{CH2=CH-CH=CH2} \xrightarrow{t^\circ, xt, p} \ce{[-CH2-CH=CH-CH2-]_n}$.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 15
    tex.append(r"""
\begin{ex}
Số liên kết peptide trong phân tử peptide Gly-Ala-Val-Tyr-Phe là:
\choice
{2}
{1}
{3}
{\True 4}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Pentapeptide gồm 5 gốc amino acid thì có $5 - 1 = \mathbf{4}$ liên kết peptide ($-\ce{CO-NH-}$).\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 16
    tex.append(r"""
\begin{ex}
Trong các loại tơ: nylon-6,6; tơ visco; nylon-6; tơ capron; sợi bông; tơ cellulose acetate. Số loại tơ có nguồn gốc từ cellulose là:
\choice
{5}
{\True 3}
{2}
{4}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Các loại tơ có nguồn gốc từ cellulose gồm: tơ visco, sợi bông, tơ cellulose acetate (3 loại).\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN II
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN 2: TRẮC NGHIỆM ĐÚNG -- SAI (2 điểm – 2 câu).}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{2pt}
""")

    # Câu 1 Phần II
    tex.append(r"""
\begin{ex}
Cho thế điện cực chuẩn của các cặp oxi hóa -- khử như sau:
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Cặp oxi hóa -- khử} & $\ce{Fe^2+}/\ce{Fe}$ & $\ce{Cu^2+}/\ce{Cu}$ & $\ce{Fe^3+}/\ce{Fe^2+}$ & $\ce{Ag+}/\ce{Ag}$ \\
\hline
$E^\circ\ (\text{V})$ & -0,04 & 0,34 & 0,77 & 0,80 \\
\hline
\end{tabular}
\end{center}
\choiceTF
{Các kim loại được xếp theo chiều tính khử tăng dần: $\text{Ag} < \text{Fe} < \text{Cu}$}
{Cho Fe tác dụng với dung dịch $\ce{AgNO3}$ dư thu được dung dịch Y chứa 2 muối là $\ce{Fe(NO3)2}$ và $\ce{AgNO3}$}
{\True Sức điện động chuẩn của pin điện hóa Fe-Cu là: 0,38V}
{\True Khi nhúng thanh đồng có khối lượng $6,4\text{ gam}$ vào $400\text{ mL}$ dung dịch $\ce{AgNO3}\ 0,2\text{M}$. Giả thiết toàn bộ lượng Ag giải phóng đều bám vào thanh đồng thì khối lượng thanh đồng sau khi phản ứng kết thúc là $12,48\text{ gam}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
\begin{itemize}[leftmargin=15pt,itemsep=1pt]
\item a) \textbf{SAI}: Căn cứ thế điện cực chuẩn, tính khử của Fe mạnh hơn Cu: Chiều tăng dần đúng là $\text{Ag} < \text{Cu} < \text{Fe}$.
\item b) \textbf{SAI}: Do $\ce{AgNO3}$ dư nên Fe bị oxi hóa lên sắt(III): $\ce{Fe + 3AgNO3 -> Fe(NO3)3 + 3Ag}$. Dung dịch thu được gồm $\ce{Fe(NO3)3}$ và $\ce{AgNO3}$ dư, không chứa $\ce{Fe(NO3)2}$.
\item c) \textbf{ĐÚNG}: $E^\circ_\text{pin}(\text{Fe-Cu}) = E^\circ(\ce{Cu^2+}/\ce{Cu}) - E^\circ(\ce{Fe^2+}/\ce{Fe}) = 0,34 - (-0,04) = \mathbf{0,38\text{ V}}$.
\item d) \textbf{ĐÚNG}: $n(\ce{AgNO3}) = 0,4 \times 0,2 = 0,08\text{ mol}$.\\
$\ce{Cu + 2Ag+ -> Cu^2+ + 2Ag}$\\
$n(\ce{Cu}\text{ tan}) = 0,04\text{ mol};\ n(\ce{Ag}\text{ bám}) = 0,08\text{ mol}$.\\
$\Delta m = 0,08 \times 108 - 0,04 \times 64 = 8,64 - 2,56 = 6,08\text{ gam}$.\\
Khối lượng thanh đồng sau phản ứng: $m = 6,4 + 6,08 = \mathbf{12,48\text{ gam}}$.
\end{itemize}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần II
    path_capro = os.path.join(DIR_IMAGES, "caprolactam_scheme.png").replace("\\", "/")
    tex_c2 = r"""
\begin{ex}
Tơ capron là loại tơ có tính dai, độ đàn hồi và độ bóng cao, ít thấm nước và được sử dụng để dệt vải may mặc. Tơ capron được tổng hợp từ caprolactam theo phương trình hóa học sau:
\begin{center}
\includegraphics[width=9.5cm]{__PATH_CAPRO__}
\end{center}
\choiceTF
{Phản ứng (1) thuộc loại phản ứng trùng ngưng}
{\True Tơ capron là tơ tổng hợp}
{Do tơ capron có các nhóm $-\ce{CO-NH}-$ nên tơ này bền với dung dịch kiềm mạnh}
{Do tơ capron có các nhóm $-\ce{CO-NH}-$ nên tơ capron là loại polypeptide}
""".replace("__PATH_CAPRO__", path_capro)
    tex.append(tex_c2)
    if is_teacher:
        tex.append(r"""
\loigiai{
\begin{itemize}[leftmargin=15pt,itemsep=1pt]
\item a) \textbf{SAI}: Phản ứng mở vòng caprolactam tạo tơ capron là phản ứng TRÙNG HỢP (trùng hợp mở vòng), không giải phóng phân tử nhỏ nên không phải trùng ngưng.
\item b) \textbf{ĐÚNG}: Tơ capron được tổng hợp hóa học nên thuộc loại tơ tổng hợp.
\item c) \textbf{SAI}: Liên kết amide $-\ce{CO-NH}-$ rất dễ bị thủy phân trong dung dịch kiềm và acid nên tơ capron kém bền trong môi trường kiềm mạnh.
\item d) \textbf{SAI}: Polypeptide là polymer tạo từ các $\alpha$-amino acid. Tơ capron tạo từ $\varepsilon$-aminocaproic acid nên là polyamide, không phải polypeptide.
\end{itemize}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN III
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN 3: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (1 điểm - 4 câu).}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 4.}
\vspace{2pt}
""")

    # Câu 1 Phần III
    tex.append(r"""
\begin{ex}
Isoamyl acetate có mùi thơm đặc trưng của chuối chín nên còn được gọi là dầu chuối. Khi đun nóng hỗn hợp gồm $16,2\text{ gam}$ acetic acid và $26,4\text{ gam}$ isoamyl alcohol ($\ce{((CH3)2CHCH2CH2OH)}$) với xúc tác là $\ce{H2SO4}$ đặc, đun nóng thu được $26\text{ gam}$ dầu chuối. Tính hiệu suất của phản ứng điều chế dầu chuối trên. (Kết quả làm tròn đến hàng phần mười).
\hfill \textbf{Trả lời:} \dotfill
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$n(\ce{CH3COOH}) = \frac{16,2}{60} = 0,27\text{ mol}$; $n(\text{alcohol}) = \frac{26,4}{88} = 0,30\text{ mol} \Rightarrow$ Hiệu suất tính theo $\ce{CH3COOH}$.\\
$m_\text{LT}(\text{ester}) = 0,27 \times 130 = 35,1\text{ gam}$.\\
Hiệu suất: $H = \frac{26}{35,1} \times 100\% \approx \mathbf{74,1\%}$.\\
\textbf{Đáp án: 74,1\%}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần III
    tex.append(r"""
\begin{ex}
Cho các chất: glucose, amylose, cellulose, maltose, fructose, saccharose. Có bao nhiêu chất tham gia phản ứng thuỷ phân trong môi trường acid?
\hfill \textbf{Trả lời:} \dotfill
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Các chất bị thủy phân trong acid gồm disaccharide và polysaccharide: amylose, cellulose, maltose, saccharose (4 chất). Glucose và fructose là monosaccharide nên không bị thủy phân.\\
\textbf{Đáp án: 4}
}
""")
    tex.append(r"\end{ex}")

    # Câu 3 Phần III
    tex.append(r"""
\begin{ex}
Thủy phân một tripeptide X (mạch hở) thu được hỗn hợp gồm glycine, alanine (có tỉ lệ mol 2:1). Có bao nhiêu đồng phân cấu tạo của X thỏa mãn tính chất trên?
\hfill \textbf{Trả lời:} \dotfill
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Tripeptide X gồm 2 Gly và 1 Ala. Các đồng phân cấu tạo của X là: Gly-Gly-Ala, Gly-Ala-Gly, Ala-Gly-Gly $\Rightarrow$ Có \textbf{3} đồng phân cấu tạo.\\
\textbf{Đáp án: 3}
}
""")
    tex.append(r"\end{ex}")

    # Câu 4 Phần III
    tex.append(r"""
\begin{ex}
Poly(hexamethylene adipamide) còn được gọi là tơ nylon-6,6 được điều chế từ phản ứng trùng ngưng của adipic acid và hexamethylene diamine. Tơ nylon-6,6 có tính dai, mềm, ít thấm nước nên được dùng để dệt vải may mặc, bện dây cáp, dây dù, đan lưới. Một loại tơ nylon-6,6 có khối lượng phân tử là $37290\text{ đvC}$. Có bao nhiêu đơn vị mắt xích có trong phân tử loại tơ sợi này?
\hfill \textbf{Trả lời:} \dotfill
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Công thức mắt xích: $\ce{[-NH-(CH2)6-NH-CO-(CH2)4-CO-]}$.\\
Khối lượng mắt xích: $M = 14 \times 12 + 26 \times 1 + 2 \times 14 + 2 \times 16 = 226\text{ đvC}$.\\
Số lượng mắt xích: $n = \frac{37290}{226} = \mathbf{165}$.\\
\textbf{Đáp án: 165}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN IV: TỰ LUẬN
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\Large B. PHẦN TỰ LUẬN (3 câu – 3 điểm)}}\\
\textit{\small Thí sinh trình bày chi tiết lời giải từ câu 1 đến câu 3 vào giấy làm bài thi.}
\vspace{3pt}
""")

    # Câu 1 Tự luận
    tex.append(r"""
\noindent
\textbf{Câu 1. (1,0 điểm)} Glycine là một amino acid có khối lượng phân tử nhỏ nhất trong số 20 amino acid thường được tìm thấy trong protein. Glycine tham gia vào nhiều chức năng của cơ thể, thúc đẩy quá trình phát triển cơ, cải thiện hệ tiêu hóa, làm chậm quá trình lão hóa,...
\begin{enumerate}[label=\alph*.,leftmargin=20pt,itemsep=1pt]
\item Hãy viết công thức cấu tạo và gọi tên hệ thống của Glycine.
\item Viết phương trình hóa học chứng minh tính lưỡng tính của Glycine.
\end{enumerate}
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=tableGreen!30,colframe=mainGreen!70,title=\textbf{Hướng dẫn giải chi tiết Câu 1},fonttitle=\bfseries]
\textbf{a. Công thức cấu tạo và tên hệ thống của Glycine:}
\begin{itemize}[leftmargin=15pt,itemsep=1pt]
\item Công thức cấu tạo: $\ce{H2N-CH2-COOH}$
\item Tên hệ thống (IUPAC): \textbf{aminoethanoic acid} (hoặc 2-aminoethanoic acid).
\end{itemize}
\textbf{b. Phương trình hóa học chứng minh tính lưỡng tính của Glycine:}
\begin{itemize}[leftmargin=15pt,itemsep=1pt]
\item Tác dụng với acid mạnh (thể hiện tính base do nhóm $-\ce{NH2}$):
\[
\ce{H2N-CH2-COOH + HCl -> ClH3N-CH2-COOH}
\]
\item Tác dụng với base mạnh (thể hiện tính acid do nhóm $-\ce{COOH}$):
\[
\ce{H2N-CH2-COOH + NaOH -> H2N-CH2-COONa + H2O}
\]
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{}{7}{\noindent\makebox[\linewidth]{\dotfill}\\[14pt]}
""")

    # Câu 2 Tự luận
    tex.append(r"""
\vspace{2pt}
\noindent
\textbf{Câu 2. (1,0 điểm)} Poly(vinyl chloride) hay còn được gọi là PVC được điều chế từ khí thiên nhiên theo sơ đồ:
\[
\ce{CH4 -> C2H2 -> CH2=CHCl -> PVC}
\]
Để sản xuất 1 tấn PVC theo sơ đồ trên cần bao nhiêu $\text{m}^3$ khí thiên nhiên (ở đkc), giả sử hiệu suất toàn bộ quá trình là $80\%$? Biết khí thiên nhiên chứa $95\%$ methane về thể tích. (Kết quả làm tròn đến hàng phần mười).
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=tableGreen!30,colframe=mainGreen!70,title=\textbf{Hướng dẫn giải chi tiết Câu 2},fonttitle=\bfseries]
1. Sơ đồ bảo toàn nguyên tố C:
\[
2\,\ce{CH4} \longrightarrow \ce{[-CH2-CH(Cl)-]_n}
\]
\[
2\text{ mol } (2 \times 24,79\text{ L}) \longrightarrow 62,5\text{ gam}
\]
2. Lượng PVC cần sản xuất: $m(\text{PVC}) = 1\text{ tấn} = 10^6\text{ gam}$.\\
Số mol mắt xích PVC: $n(\text{PVC}) = \frac{10^6}{62,5} = 16000\text{ mol}$.\\
3. Số mol $\ce{CH4}$ cần dùng theo lý thuyết:
\[
n(\ce{CH4}) = 2 \times n(\text{PVC}) = 2 \times 16000 = 32000\text{ mol}
\]
4. Thể tích $\ce{CH4}$ ở điều kiện chuẩn ($24,79\text{ L/mol}$):
\[
V_\text{LT}(\ce{CH4}) = 32000 \times 24,79\text{ L} = 793280\text{ L} = 793,28\text{ m}^3
\]
5. Vì hiệu suất toàn bộ quá trình là $80\%$, thể tích $\ce{CH4}$ thực tế cần dùng là:
\[
V_\text{TT}(\ce{CH4}) = \frac{793,28}{80\%} = 991,6\text{ m}^3
\]
6. Thể tích khí thiên nhiên cần dùng (chứa $95\%$ methane):
\[
V(\text{khí thiên nhiên}) = \frac{991,6}{95\%} \approx \mathbf{1043,8\text{ m}^3}
\]
\textbf{Kết luận:} Cần khoảng $\mathbf{1043,8\text{ m}^3}$ khí thiên nhiên (ở đkc).
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{}{7}{\noindent\makebox[\linewidth]{\dotfill}\\[14pt]}
""")

    # Câu 3 Tự luận
    path_soap = os.path.join(DIR_IMAGES, "banh_xa_bong.png").replace("\\", "/")
    tex_c3 = r"""
\vspace{2pt}
\noindent
\textbf{Câu 3. (1,0 điểm)}
\begin{enumerate}[label=\alph*.,leftmargin=20pt,itemsep=1pt]
\item Xà phòng (xà bông) là một chất giặt rửa các vết bẩn, vết dầu mỡ,... Hãy nêu thành phần chính của xà phòng?
\begin{center}
\includegraphics[width=4.0cm]{__PATH_SOAP__}
\end{center}
\item Hỗn hợp X gồm tristearin, tripalmitin, stearic acid và palmitic acid (tỉ lệ mol $n_\text{tristearin} : n_\text{tripalmitin} = 1 : 1$ và $n_\text{palmitic acid} : n_\text{stearic acid} = 2 : 1$). Để sản xuất $b$ bánh xà phòng (mỗi bánh đều nặng $75\text{ gam}$ và chứa $75\%$ muối của các acid béo theo khối lượng), người ta cho $863,92\text{ gam}$ hỗn hợp X tác dụng vừa đủ với $306\text{ gam}$ dung dịch NaOH $40\%$, đun nóng. Biết hiệu suất toàn bộ quá trình sản xuất xà phòng là $88\%$. Tính giá trị của $b$ (Kết quả làm tròn đến hàng đơn vị).
\end{enumerate}
""".replace("__PATH_SOAP__", path_soap)
    tex.append(tex_c3)
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=tableGreen!30,colframe=mainGreen!70,title=\textbf{Hướng dẫn giải chi tiết Câu 3},fonttitle=\bfseries]
\textbf{a. Thành phần chính của xà phòng:}
Thành phần chính của xà phòng là muối sodium hoặc potassium của các acid béo cao (như sodium stearate, sodium palmitate, sodium oleate,...), ngoài ra có thêm các chất độn, chất tạo màu và hương liệu.

\textbf{b. Tính giá trị của $b$:}
\begin{enumerate}[leftmargin=15pt,itemsep=1pt]
\item Số mol NaOH:
\[
n(\ce{NaOH}) = \frac{306 \times 40\%}{40} = 3,06\text{ mol}
\]
\item Đặt ẩn:
Gọi $n(\text{tristearin}) = n(\text{tripalmitin}) = x\text{ mol}$.\\
Gọi $n(\text{stearic acid}) = y\text{ mol} \Rightarrow n(\text{palmitic acid}) = 2y\text{ mol}$.
\item Lập hệ phương trình:
\begin{itemize}
\item Theo số mol NaOH phản ứng:
\[
3x + 3x + y + 2y = 6x + 3y = 3,06 \Rightarrow 2x + y = 1,02 \quad (1)
\]
\item Theo khối lượng hỗn hợp X:
\[
890x + 806x + 284y + 256(2y) = 1696x + 796y = 863,92 \quad (2)
\]
\end{itemize}
Giải hệ (1) và (2) ta được: $x = 0,5\text{ mol};\ y = 0,02\text{ mol}$.
\item Tính khối lượng muối acid béo theo bảo toàn khối lượng:
\begin{itemize}
\item $n(\text{glycerol}) = 2x = 1,0\text{ mol} \Rightarrow m(\text{glycerol}) = 1,0 \times 92 = 92\text{ gam}$.
\item $n(\ce{H2O}) = 3y = 0,06\text{ mol} \Rightarrow m(\ce{H2O}) = 0,06 \times 18 = 1,08\text{ gam}$.
\item $m(\text{muối}) = m_X + m(\ce{NaOH}) - m(\text{glycerol}) - m(\ce{H2O})$
\[
= 863,92 + 122,4 - 92 - 1,08 = 893,24\text{ gam}
\]
\end{itemize}
\item Tính số bánh xà phòng $b$:
\begin{itemize}
\item Khối lượng muối thực tế thu được ($H = 88\%$):
\[
m(\text{muối TT}) = 893,24 \times 88\% = 786,0512\text{ gam}
\]
\item Lượng muối acid béo trong 1 bánh xà phòng:
\[
m(\text{muối / 1 bánh}) = 75 \times 75\% = 56,25\text{ gam}
\]
\item Số bánh xà phòng sản xuất được:
\[
b = \frac{786,0512}{56,25} \approx 13,974 \approx \mathbf{14\text{ (bánh)}}
\]
\end{itemize}
\end{enumerate}
\textbf{Kết luận:} $b = 14$ bánh xà phòng.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{}{9}{\noindent\makebox[\linewidth]{\dotfill}\\[14pt]}
""")

    tex.append(r"""
\vfill
\begin{center}
\textbf{---------- HẾT ----------}
\end{center}
\end{document}
""")

    return "".join(tex)


def compile_pdf(tex_path):
    out_dir = os.path.dirname(tex_path)
    base_name = os.path.splitext(os.path.basename(tex_path))[0]
    
    cmd = ["pdflatex", "-interaction=nonstopmode", f"-output-directory={out_dir}", tex_path]
    
    for run_idx in range(2):
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="ignore")
        if res.returncode != 0:
            print(f"[CẢNH BÁO] Lần chạy {run_idx+1} có mã trả về {res.returncode}")
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
    tex_hs = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_CK1_Hoa_12_NguyenChiThanh_Ma1001_HocSinh.tex")
    tex_gv = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_CK1_Hoa_12_NguyenChiThanh_Ma1001_GiaoVien.tex")

    print("Đang tạo mã nguồn LaTeX Bản Học sinh THPT Nguyễn Chí Thanh Hóa 12...")
    with open(tex_hs, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))

    print("Đang tạo mã nguồn LaTeX Bản Giáo viên THPT Nguyễn Chí Thanh Hóa 12...")
    with open(tex_gv, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))

    print("\nĐang biên dịch PDF Bản Học sinh THPT Nguyễn Chí Thanh Hóa 12...")
    compile_pdf(tex_hs)

    print("\nĐang biên dịch PDF Bản Giáo viên THPT Nguyễn Chí Thanh Hóa 12...")
    compile_pdf(tex_gv)

if __name__ == "__main__":
    main()
