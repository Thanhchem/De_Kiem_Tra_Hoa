# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Cuối Kỳ I Hóa Học 12 - THPT Đặng Huy Trứ (TP Huế)
Mã đề: 104 - Năm học 2025 - 2026
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_CK1_Hoa_12_DangHuyTru_Ma104_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_CK1_Hoa_12_DangHuyTru_Ma104_GiaoVien.pdf (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
DIR_IMAGES = os.path.join(SAN_PHAM_DIR, "images_de104")

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
\textbf{SỞ GD - ĐT THÀNH PHỐ HUẾ}\\
\textbf{TRƯỜNG THPT ĐẶNG HUY TRỨ}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề gồm có 03 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}KIỂM TRA CUỐI KỲ I}\\
\textbf{NĂM HỌC: 2025 -- 2026}\\
\textbf{Môn: HÓA HỌC -- Lớp 12}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{MÃ ĐỀ: 104}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ và tên học sinh:} \dotfill \textit{Lớp:} \dotfill \textit{SBD:} \dotfill

\vspace{3pt}
\noindent
\textit{\footnotesize Cho NTK các nguyên tố: $\ce{Cu}=64;\ \ce{Ag}=108;\ \ce{C}=12;\ \ce{H}=1;\ \ce{O}=16;\ \ce{N}=14;\ \ce{K}=39;\ \ce{Na}=23;\ \ce{Zn}=65;\ \ce{Ni}=59;\ \ce{Sn}=119;\ \ce{Mg}=24$.}

\vspace{3pt}
\hrule color{hfGreen} height 0.5pt
""")

    if is_teacher:
        # BẢNG ĐÁP ÁN RÚT GỌN CHO GIÁO VIÊN
        tex.append(r"""
\vspace{6pt}
\begin{center}
{\Large\textbf{\color{hfRed}BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN -- MÃ ĐỀ 104)}}
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
\textbf{Đ/A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} \\
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
\textbf{Câu 1} & Polymer Poly(ethylene terephthalate) (PET) & \textbf{\color{hfRed}Đ} & \textbf{\color{gray}S} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}Đ} \\
\hline
\textbf{Câu 2} & Phản ứng thanh đồng Cu trong dung dịch $\ce{AgNO3}$ & \textbf{\color{gray}S} & \textbf{\color{gray}S} & \textbf{\color{hfRed}Đ} & \textbf{\color{gray}S} \\
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
\textbf{Đáp án} & \textbf{\color{hfRed}5} & \textbf{\color{hfRed}3} & \textbf{\color{hfRed}4} & \textbf{\color{hfRed}3} \\
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
\textbf{Câu 1} & 1,0 đ & PTHH xà phòng hóa tristearin; $n(\text{tristearin}) = 0,2\text{ mol} \Rightarrow n(\ce{C17H35COOK}) = 0,6\text{ mol} \Rightarrow m = \mathbf{193,2\text{ gam}}$. \\
\hline
\textbf{Câu 2} & 1,0 đ & 4 PTHH: oxi hóa glucose bằng $\ce{Br2}$; glycine + $\ce{NaOH}$; propan-1-amine + $\ce{HCl}$; methyl acetate + $\ce{KOH}$. \\
\hline
\textbf{Câu 3} & 1,0 đ & a) $\ce{Zn(s) + Sn^2+(aq) -> Zn^2+(aq) + Sn(s)}$; b) $E^\circ_\text{pin}(\text{Zn-Sn}) = 0,51 + 0,11 = \mathbf{0,62\text{ V}}$. \\
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
{\color{mainGreen}\textbf{\Large A. TRẮC NGHIỆM (7,0 điểm).}}

\vspace{3pt}
\noindent
{\color{mainGreen}\textbf{\large Phần I: Trắc nghiệm nhiều lựa chọn (4,0 điểm).}}\\
\textit{\small Học sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi học sinh chỉ chọn một phương án.}
\vspace{2pt}
""")

    # CÂU 1 -> 16 PHẦN I
    # Câu 1
    tex.append(r"""
\begin{ex}
Cho các thế điện cực chuẩn của các cặp oxi hóa -- khử của kim loại như sau:
(1) $\ce{Ag+}/\ce{Ag}\ (= 0,80\text{V})$; (2) $\ce{Cu^2+}/\ce{Cu}\ (= 0,34\text{V})$; (3) $\ce{Na+}/\ce{Na}\ (-2,71\text{V})$; (4) $\ce{Mg^2+}/\ce{Mg}\ (-2,37\text{V})$.
Ion kim loại có tính oxi hoá mạnh nhất là ion kim loại trong cặp oxi hóa khử
\choice
{(2)}
{(4)}
{(3)}
{\True (1)}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Căn cứ vào dãy thế điện cực chuẩn: Cặp có $E^\circ$ càng lớn thì dạng oxi hóa càng có tính oxi hóa mạnh. Cặp (1) $\ce{Ag+}/\ce{Ag}$ có $E^\circ = +0,80\text{ V}$ lớn nhất nên ion $\ce{Ag+}$ có tính oxi hóa mạnh nhất.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2
    tex.append(r"""
\begin{ex}
Chất nào dưới đây không phải là amino acid?
\choice
{Glycine}
{Lysine}
{Glutamic acid}
{\True Aniline}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Aniline ($\ce{C6H5NH2}$) là amine thơm đơn chức bậc I, không chứa nhóm $-\ce{COOH}$ nên không phải là amino acid. Glycine, lysine và glutamic acid đều là các amino acid.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 3
    tex.append(r"""
\begin{ex}
Chất nào sau đây có thể tham gia phản ứng trùng ngưng?
\choice
{$\ce{CH2=CHC6H5}$}
{\True $\ce{H2N[CH2]5COOH}$}
{$\ce{CH2=CHCOOH}$}
{$\ce{CH2=CH(CH3)}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$\ce{H2N[CH2]5COOH}$ ($\varepsilon$-aminocaproic acid) chứa đồng thời 2 nhóm chức $-\ce{NH2}$ và $-\ce{COOH}$ có khả năng phản ứng với nhau tách nước nên tham gia phản ứng trùng ngưng tạo tơ capron. Các chất còn lại chỉ tham gia trùng hợp.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 4
    tex.append(r"""
\begin{ex}
Một pin điện hoá có điện cực là Zn nhúng trong dung dịch $\ce{ZnSO4}$ và điện cực là Cu nhúng trong dung dịch $\ce{CuSO4}$. Sau một thời gian pin đó phóng điện thì
\choice
{khối lượng cả hai điện cực Zn và Cu đều giảm}
{khối lượng cả hai điện cực Zn và Cu đều tăng}
{khối lượng điện cực Zn tăng còn khối lượng điện cực Cu giảm}
{\True khối lượng điện cực Zn giảm còn khối lượng điện cực Cu tăng}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Khi pin hoạt động:\\
- Tại anode (cực âm Zn): Zn bị oxi hóa tan ra: $\ce{Zn} \to \ce{Zn^2+} + 2\ce{e} \Rightarrow$ khối lượng Zn giảm.\\
- Tại cathode (cực dương Cu): $\ce{Cu^2+}$ bị khử tạo Cu bám vào: $\ce{Cu^2+} + 2\ce{e} \to \ce{Cu} \Rightarrow$ khối lượng Cu tăng.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 5
    tex.append(r"""
\begin{ex}
Chất nào dưới đây không tan trong nước nhưng tan được trong dung dịch Schweizer?
\choice
{Saccharose}
{\True Cellulose}
{Maltose}
{Fructose}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Cellulose có mạch phân tử kéo dài, nhiều liên kết hydrogen liên phân tử chặt chẽ nên không tan trong nước, nhưng tan được trong dung dịch Schweizer (phức chất $\ce{[Cu(NH3)4](OH)2}$).\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 6
    tex.append(r"""
\begin{ex}
Cho tetrapeptide X: Gly-Ala-Glu-Val. Amino acid đầu N là
\choice
{Val}
{\True Gly}
{Glu}
{Ala}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Theo quy ước, amino acid đầu N (còn nhóm $-\ce{NH2}$ tự do) nằm ở đầu bên trái của chuỗi peptide, đó là Gly. Amino acid đầu C là Val.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 7
    path_c7 = os.path.join(DIR_IMAGES, "ky_hieu_tai_che_6.png").replace("\\", "/")
    tex_c7 = r"""
\begin{ex}
Hình dưới đây là ký hiệu của 6 polymer nhiệt dẻo phổ biến có thể tái chế:
\begin{center}
\includegraphics[width=9.5cm]{__PATH_C7__}
\end{center}
Các ký hiệu này thường được in trên bao bì, vỏ hộp, đồ dùng,... để giúp nhận biết vật liệu polymer cũng như thuận lợi cho việc thu gom, tái chế. Polymer có ký hiệu số 5 được điều chế bằng phản ứng trùng hợp monomer nào dưới đây?
\choice
{$\ce{CH2=CH2}$}
{$\ce{CH2=CH-Cl}$}
{\True $\ce{CH2=CH-CH3}$}
{$\ce{CH2=CH-C6H5}$}
""".replace("__PATH_C7__", path_c7)
    tex.append(tex_c7)
    if is_teacher:
        tex.append(r"""
\loigiai{
Ký hiệu số 5 là nhựa PP (polypropylene), được điều chế bằng phản ứng trùng hợp propylene (propene, $\ce{CH2=CH-CH3}$).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 8
    path_c8 = os.path.join(DIR_IMAGES, "ctct_amino_acid_cau8.png").replace("\\", "/")
    tex_c8 = r"""
\begin{ex}
Cho các hợp chất có công thức cấu tạo dưới đây:
\begin{center}
\includegraphics[width=8.5cm]{__PATH_C8__}
\end{center}
Những hợp chất nào trong số các chất trên thuộc loại $\alpha$ -- amino acid?
\choice
{\True Chất (1), chất (3)}
{Chất (1) và chất (2)}
{Chất (1), chất (2) và chất (3)}
{Chất (2) và chất (3)}
""".replace("__PATH_C8__", path_c8)
    tex.append(tex_c8)
    if is_teacher:
        tex.append(r"""
\loigiai{
$\alpha$-amino acid là amino acid có nhóm $-\ce{NH2}$ đính trực tiếp vào carbon $\alpha$ (C số 2, kề nhóm $-\ce{COOH}$):\\
- Chất (1): $\ce{CH3-CH2-CH(NH2)-COOH}$ ($\alpha$-aminobutyric acid, nhóm $-\ce{NH2}$ ở vị trí $\alpha$).\\
- Chất (2): $\ce{H2N-CH2-CH2-COOH}$ ($\beta$-alanine, nhóm $-\ce{NH2}$ ở vị trí $\beta$).\\
- Chất (3): $\ce{HOOC-CH(NH2)-CH3}$ (alanine, nhóm $-\ce{NH2}$ ở vị trí $\alpha$).\\
Vậy chất (1) và chất (3) thuộc loại $\alpha$-amino acid.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 9
    tex.append(r"""
\begin{ex}
Khẳng định nào sau đây đúng?
\choice
{Chất béo là triester của glycerol với acid vô cơ}
{Chất béo là chất rắn tan tốt trong nước}
{\True Chất béo không tan trong nước nhưng tan nhiều trong dung môi hữu cơ}
{Dầu ăn và mỡ bôi trơn có cùng thành phần nguyên tố}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Chất béo nhẹ hơn nước, không tan trong nước nhưng tan nhiều trong các dung môi hữu cơ không phân cực (benzene, ether, chloroform, xăng,...). A sai vì là acid béo hữu cơ; B sai vì không tan trong nước; D sai vì dầu ăn chứa C, H, O còn mỡ bôi trơn là hydrocarbon (chỉ chứa C, H).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 10
    tex.append(r"""
\begin{ex}
Xà phòng có thành phần chính là
\choice
{muối sodium hoặc potassium của acid bất kì}
{glycerol và sodium chloride hoặc potassium chloride}
{\True muối sodium hoặc potassium của acid béo}
{muối sodium hoặc potassium của carboxylic acid}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Xà phòng thông thường có thành phần chính là muối sodium hoặc potassium của các acid béo cao (như sodium stearate, sodium palmitate, sodium oleate,...).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 11
    tex.append(r"""
\begin{ex}
Nhận định nào sau đây đúng khi nói về amino acid?
\choice
{dễ nóng chảy và ít tan trong nước}
{\True có nhiệt độ nóng chảy cao và tan tốt trong nước}
{có nhiệt độ nóng chảy cao và ít tan trong nước}
{dễ nóng chảy và tan tốt trong nước}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Ở trạng thái tinh thể, amino acid tồn tại dưới dạng ion lưỡng cực ($\ce{H3N+-CHR-COO-}$). Lực hút tĩnh điện giữa các ion rất mạnh nên amino acid có nhiệt độ nóng chảy cao và tan tốt trong nước.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 12
    tex.append(r"""
\begin{ex}
Phát biểu nào sau đây đúng?
\choice
{\True PVC được điều chế bằng phản ứng trùng hợp}
{Tơ visco thuộc loại tơ tổng hợp}
{Amylose có cấu trúc mạch phân nhánh}
{Tơ tằm thuộc loại tơ nhân tạo}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
A đúng: Poly(vinyl chloride) (PVC) được điều chế bằng phản ứng trùng hợp monomer vinyl chloride ($\ce{CH2=CHCl}$). B sai vì tơ visco là tơ bán tổng hợp (nhân tạo); C sai vì amylose có mạch không phân nhánh; D sai vì tơ tằm là tơ thiên nhiên.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 13
    tex.append(r"""
\begin{ex}
Một tripeptide X được cấu thành từ 2 phân tử Ala và 1 phân tử Gly. Công thức cấu tạo của X không thể là
\choice
{\True Gly-Ala-Gly}
{Gly-Ala-Ala}
{Ala-Gly-Ala}
{Ala-Ala-Gly}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Tripeptide X được tạo từ 2 phân tử Ala và 1 phân tử Gly nên thành phần cấu tạo phải có đúng 2 gốc Ala và 1 gốc Gly. Gly-Ala-Gly có 2 gốc Gly và 1 gốc Ala nên không thể là công thức của X.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 14
    tex.append(r"""
\begin{ex}
Có thể phân biệt glucose và fructose bằng cách cho từng mẫu chất tác dụng với
\choice
{thuốc thử Tollens}
{\True nước bromine}
{dung dịch $\ce{NH3}$}
{$\ce{Cu(OH)2}$ trong môi trường kiềm, đun nóng}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Dùng nước bromine: Glucose chứa nhóm aldehyde làm mất màu nước bromine; fructose chứa nhóm ketone không phản ứng với nước bromine (do môi trường nước bromine có tính acid yếu, không chuyển hóa fructose thành glucose). Thuốc thử Tollens và $\ce{Cu(OH)2}$/kiềm nóng đều phản ứng với cả hai do có môi trường kiềm chuyển hóa fructose thành glucose.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 15
    tex.append(r"""
\begin{ex}
Cao su buna -- N (hay còn gọi là cao su nitrile, có kí hiệu là NBR) là loại cao su tổng hợp có khả năng chịu dầu mỡ tốt nên dùng làm ống dẫn nhiên liệu, gioăng làm kín trong các máy móc. Thực hiện phản ứng trùng hợp các chất nào dưới đây thu được sản phẩm cao su buna -- N?
\choice
{$\ce{CH2=CHCH=CH2}$ và $\ce{C6H5CH=CH2}$}
{$\ce{CH2=C(CH3)CH=CH2}$ và $\ce{CH2=CHCN}$}
{\True $\ce{CH2=CHCH=CH2}$ và $\ce{CH2=CHCN}$}
{$\ce{CH2=CHCH=CH2}$ và $\ce{N2}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Cao su buna-N được sản xuất bằng phản ứng đồng trùng hợp giữa buta-1,3-diene ($\ce{CH2=CH-CH=CH2}$) và acrylonitrile (vinyl cyanide, $\ce{CH2=CH-CN}$).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 16
    tex.append(r"""
\begin{ex}
Tên gọi của $\ce{HCOOC2H5}$ là
\choice
{methyl acetate}
{ethyl acetate}
{\True ethyl formate}
{methyl formate}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Gốc alcohol $-\ce{C2H5}$ là ethyl; gốc acid $\ce{HCOO-}$ là formate. Do đó tên gọi của ester là ethyl formate.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN II
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large Phần II: Trắc nghiệm dạng Đúng/Sai (2,0 điểm).}}\\
\textit{\small Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai.}
\vspace{2pt}
""")

    # Câu 1 Phần II
    path_pet = os.path.join(DIR_IMAGES, "phan_ung_pet.png").replace("\\", "/")
    tex_c1 = r"""
\begin{ex}
Poly(ethylene terephthalate) (viết tắt là PET) là một polymer được điều chế từ terephthalic acid và ethylene glycol theo phương trình:
\begin{center}
\includegraphics[width=12.5cm]{__PATH_PET__}
\end{center}
PET được sử dụng để sản xuất tơ, chai đựng nước uống, hộp đựng thực phẩm. Để thuận lợi cho việc nhận biết, sử dụng và tái chế thì các đồ nhựa làm từ vật liệu chứa PET thường được in kí hiệu như hình trên.
\choiceTF
{\True Trong một mắt xích PET, phần trăm khối lượng carbon là $62,5\%$}
{Phản ứng tổng hợp PET từ terephthalic acid và ethylene glycol thuộc loại phản ứng trùng hợp}
{\True Sản phẩm làm bằng nhựa số 01(PET) có thể được tái chế nhưng không nên tái sử dụng để đựng thực phẩm, nước uống}
{\True Tơ được chế tạo từ PET thuộc loại tơ tổng hợp}
""".replace("__PATH_PET__", path_pet)
    tex.append(tex_c1)
    if is_teacher:
        tex.append(r"""
\loigiai{
\begin{itemize}[leftmargin=15pt,itemsep=1pt]
\item a) \textbf{ĐÚNG}: Một mắt xích PET có cấu trúc $\ce{[-O-CH2-CH2-OOC-C6H4-CO-]}$ có CTPT $\ce{C10H8O4}$ ($M = 192\text{ g/mol}$). Phần trăm khối lượng carbon là $\%C = \frac{10 \times 12}{192} \times 100\% = 62,5\%$.
\item b) \textbf{SAI}: Phản ứng kết hợp các monomer đồng thời giải phóng phân tử nhỏ $\ce{H2O}$ nên thuộc loại phản ứng TRÙNG NGƯNG (đồng trùng ngưng), không phải trùng hợp.
\item c) \textbf{ĐÚNG}: Nhựa PET số 01 chỉ an toàn cho sử dụng một lần. Tái sử dụng nhiều lần bề mặt dễ xốp, tích tụ vi khuẩn và thôi nhiễm chất độc hại vào thực phẩm.
\item d) \textbf{ĐÚNG}: Tơ kéo từ PET (tơ lapsan/dacron) là polymer tổng hợp nên thuộc loại tơ tổng hợp.
\end{itemize}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần II
    tex.append(r"""
\begin{ex}
Trong phòng thí nghiệm, một học sinh nhúng thanh đồng có khối lượng $12,340\text{ g}$ vào $255\text{ mL}$ dung dịch $\ce{AgNO3}\ 0,125\text{M}$. Bằng quan sát, học sinh đó đã khẳng định có phản ứng xảy ra.
\choiceTF
{Tính oxi hóa của $\ce{Cu^2+}$ mạnh hơn $\ce{Ag+}$}
{Dung dịch từ không màu chuyển sang màu nâu đỏ}
{\True Phương trình hóa học của phản ứng: $\ce{Cu + 2AgNO3 -> Cu(NO3)2 + 2Ag}$}
{Giả thuyết toàn bộ lượng Ag giải phóng đều bám vào thanh đồng thì khối lượng thanh đồng sau khi phản ứng kết thúc là $13,35\text{ gam}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
\begin{itemize}[leftmargin=15pt,itemsep=1pt]
\item a) \textbf{SAI}: $E^\circ(\ce{Ag+}/\ce{Ag}) = +0,80\text{ V} > E^\circ(\ce{Cu^2+}/\ce{Cu}) = +0,34\text{ V}$ nên tính oxi hóa của $\ce{Ag+}$ mạnh hơn $\ce{Cu^2+}$.
\item b) \textbf{SAI}: Dung dịch muối $\ce{Cu(NO3)2}$ sinh ra có màu xanh lam đặc trưng, không phải màu nâu đỏ.
\item c) \textbf{ĐÚNG}: Phương trình phản ứng đúng: $\ce{Cu + 2AgNO3 -> Cu(NO3)2 + 2Ag}$.
\item d) \textbf{SAI}: $n(\ce{AgNO3}) = 0,255 \times 0,125 = 0,031875\text{ mol}$.\\
$\Rightarrow n(\ce{Cu}\text{ tan}) = 0,0159375\text{ mol};\ n(\ce{Ag}\text{ bám}) = 0,031875\text{ mol}$.\\
Độ tăng khối lượng: $\Delta m = 0,031875 \times 108 - 0,0159375 \times 64 = 3,4425 - 1,02 = 2,4225\text{ gam}$.\\
Khối lượng thanh Cu sau phản ứng: $m = 12,340 + 2,4225 = 14,7625\text{ gam} \neq 13,35\text{ gam}$.
\end{itemize}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN III
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large Phần III: Trắc nghiệm trả lời ngắn (1,0 điểm).}}\\
\textit{\small Học sinh trả lời từ câu 1 đến câu 4.}
\vspace{2pt}
""")

    # Câu 1 Phần III
    tex.append(r"""
\begin{ex}
Cho các phát biểu:
\begin{enumerate}[label=(\alph*),leftmargin=20pt,itemsep=1pt]
\item Oxi hóa glucose bằng thuốc thử Tollens tạo ra gluconic acid.
\item Dung dịch của glucose hay saccharose đều có thể hòa tan $\ce{Cu(OH)2}$ ở điều kiện thường.
\item Cellulose trinitrate là nguyên liệu để sản xuất thuốc súng không khói.
\item Trong cellulose, các gốc $\beta$-glucose chỉ liên kết với nhau bằng liên kết $\beta$-1,4-glycoside.
\item Glucose và fructose có thể chuyển hóa lẫn nhau trong môi trường kiềm.
\item Glucose là hợp chất thuộc loại monosaccharide.
\end{enumerate}
Số phát biểu đúng là bao nhiêu?
\hfill \textbf{Trả lời:} \dotfill
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- (a) SAI: Thuốc thử Tollens oxi hóa glucose tạo ammonium gluconate ($\ce{CH2OH[CHOH]4COONH4}$). Muốn tạo gluconic acid cần dùng nước bromine.\\
- (b) ĐÚNG: Đều là polyalcohol có nhiều nhóm $-\ce{OH}$ kề nhau hòa tan $\ce{Cu(OH)2}$ tạo phức xanh lam.\\
- (c) ĐÚNG: Cellulose trinitrate dùng chế tạo thuốc súng không khói.\\
- (d) ĐÚNG: Cellulose chỉ có duy nhất liên kết $\beta$-1,4-glycosidic.\\
- (e) ĐÚNG: Glucose và fructose chuyển hóa qua lại trong môi trường base kiềm.\\
- (f) ĐÚNG: Glucose là monosaccharide.\\
$\Rightarrow$ Có \textbf{5} phát biểu đúng: (b), (c), (d), (e), (f).\\
\textbf{Đáp án: 5}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần III
    tex.append(r"""
\begin{ex}
Cho các loại tơ: bông, tơ capron, tơ cellulose acetate, tơ tằm, len, tơ nitron, nylon-6,6. Có bao nhiêu tơ thuộc loại tơ tổng hợp?
\hfill \textbf{Trả lời:} \dotfill
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- Tơ thiên nhiên: bông, tơ tằm, len.\\
- Tơ bán tổng hợp (nhân tạo): tơ cellulose acetate.\\
- Tơ tổng hợp: tơ capron, tơ nitron, nylon-6,6 (3 loại).\\
\textbf{Đáp án: 3}
}
""")
    tex.append(r"\end{ex}")

    # Câu 3 Phần III
    tex.append(r"""
\begin{ex}
Xét pin Galvani hoạt động với phương trình tương ứng: $\ce{Zn(s) + Cu^2+(aq) -> Cu(s) + Zn^2+(aq)}$\\
Cho các phát biểu:
\begin{enumerate}[label=(\alph*),leftmargin=20pt,itemsep=1pt]
\item Điện cực Copper giảm khối lượng và điện cực Copper là cực dương.
\item Sức điện động của pin $E^\circ_\text{pin} = E^\circ(+) - E^\circ(-)$.
\item Điện cực Zinc giảm khối lượng và điện cực Zinc là cực âm.
\item Khi pin hoạt động dòng điện có chiều từ Copper sang Zinc.
\item Thế điện cực chuẩn của cặp $\ce{Zn^2+}/\ce{Zn}$ nhỏ hơn của cặp $\ce{Cu^2+}/\ce{Cu}$.
\end{enumerate}
Có bao nhiêu phát biểu đúng trong các phát biểu trên?
\hfill \textbf{Trả lời:} \dotfill
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- (a) SAI: Khối lượng điện cực Cu tăng do $\ce{Cu^2+}$ bị khử tạo Cu bám vào.\\
- (b) ĐÚNG: $E^\circ_\text{pin} = E^\circ(\text{cathode}) - E^\circ(\text{anode}) = E^\circ(+) - E^\circ(-)$.\\
- (c) ĐÚNG: Zn bị oxi hóa tan ra, là cực âm anode.\\
- (d) ĐÚNG: Chiều dòng điện quy ước chạy từ cực dương (Cu) sang cực âm (Zn).\\
- (e) ĐÚNG: $E^\circ(\ce{Zn^2+}/\ce{Zn}) = -0,76\text{ V} < E^\circ(\ce{Cu^2+}/\ce{Cu}) = +0,34\text{ V}$.\\
$\Rightarrow$ Có \textbf{4} phát biểu đúng: (b), (c), (d), (e).\\
\textbf{Đáp án: 4}
}
""")
    tex.append(r"\end{ex}")

    # Câu 4 Phần III
    tex.append(r"""
\begin{ex}
Có bao nhiêu hợp chất đơn chức ứng với công thức phân tử $\ce{C3H6O2}$ đều tác dụng được với dung dịch NaOH?
\hfill \textbf{Trả lời:} \dotfill
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Hợp chất đơn chức $\ce{C3H6O2}$ ($k = 1$, 2 nguyên tử O) tác dụng được với dung dịch NaOH gồm:\\
1. Carboxylic acid: $\ce{CH3-CH2-COOH}$ (propanoic acid) (1 chất).\\
2. Ester:\\
   - $\ce{HCOO-CH2-CH3}$ (ethyl formate)\\
   - $\ce{CH3-COO-CH3}$ (methyl acetate)\\
   (2 chất).\\
$\Rightarrow$ Tổng cộng có: $1 + 2 = \mathbf{3}$ chất.\\
\textbf{Đáp án: 3}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN IV: TỰ LUẬN
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\Large B. TỰ LUẬN (3,0 điểm).}}\\
\textit{\small Thí sinh trình bày chi tiết lời giải từ câu 1 đến câu 3 vào giấy làm bài thi.}
\vspace{3pt}
""")

    # Câu 1 Tự luận
    tex.append(r"""
\noindent
\textbf{Câu 1. (1,0 điểm)} Mỡ bò có hàm lượng tristearin $80\%$ về khối lượng, các loại chất béo khác có hàm lượng không đáng kể và một số tạp chất khác. Xà phòng hóa hoàn toàn $222,5\text{ gam}$ mỡ bò trong dung dịch KOH, thu được $m\text{ gam}$ potassium stearate.
\begin{enumerate}[label=\alph*),leftmargin=20pt,itemsep=1pt]
\item Viết phương trình hóa học?
\item Tính giá trị của $m$? Cho rằng các tạp chất không tham gia phản ứng với KOH.
\end{enumerate}
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=tableGreen!30,colframe=mainGreen!70,title=\textbf{Hướng dẫn giải chi tiết Câu 1},fonttitle=\bfseries]
a) Phương trình phản ứng xà phòng hóa tristearin bằng dung dịch KOH:
\[
\ce{(C17H35COO)3C3H5 + 3KOH ->[t^\circ] 3C17H35COOK + C3H5(OH)3}
\]
b) Tính giá trị của $m$:
\begin{itemize}[leftmargin=15pt,itemsep=1pt]
\item Khối lượng tristearin có trong mỡ bò:
\[
m(\text{tristearin}) = 222,5 \times 80\% = 178\text{ gam}
\]
\item Phân tử khối tristearin $\ce{(C17H35COO)3C3H5}$: $M = 3 \times 283 + 41 = 890\text{ g/mol}$.
\item Số mol tristearin:
\[
n(\text{tristearin}) = \frac{178}{890} = 0,2\text{ mol}
\]
\item Theo phương trình hóa học:
\[
n(\ce{C17H35COOK}) = 3 \times n(\text{tristearin}) = 3 \times 0,2 = 0,6\text{ mol}
\]
\item Phân tử khối $\ce{C17H35COOK}$: $M = 283 + 39 = 322\text{ g/mol}$.
\item Khối lượng muối potassium stearate thu được:
\[
m = 0,6 \times 322 = \mathbf{193,2\text{ gam}}
\]
\end{itemize}
\textbf{Kết luận:} $m = 193,2\text{ gam}$.
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
\textbf{Câu 2. (1,0 điểm)} Hoàn thành các phương trình phản ứng sau:
\begin{enumerate}[label=\alph*),leftmargin=20pt,itemsep=1pt]
\item Glucose + dung dịch $\ce{Br2}$
\item Glycine + NaOH
\item Propan-1-amine + HCl
\item Methyl acetate + KOH ($t^\circ$)
\end{enumerate}
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=tableGreen!30,colframe=mainGreen!70,title=\textbf{Hướng dẫn giải chi tiết Câu 2},fonttitle=\bfseries]
Các phương trình phản ứng hóa học hoàn chỉnh:
\begin{enumerate}[label=\alph*),leftmargin=15pt,itemsep=3pt]
\item $\ce{CH2OH[CHOH]4CHO + Br2 + H2O -> CH2OH[CHOH]4COOH + 2HBr}$
\item $\ce{H2N-CH2-COOH + NaOH -> H2N-CH2-COONa + H2O}$
\item $\ce{CH3CH2CH2NH2 + HCl -> CH3CH2CH2NH3Cl}$
\item $\ce{CH3COOCH3 + KOH ->[t^\circ] CH3COOK + CH3OH}$
\end{enumerate}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{}{7}{\noindent\makebox[\linewidth]{\dotfill}\\[14pt]}
""")

    # Câu 3 Tự luận
    tex.append(r"""
\vspace{2pt}
\noindent
\textbf{Câu 3. (1,0 điểm)} Cho các pin điện hoá và sức điện động chuẩn tương ứng:
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Pin điện hóa} & \textbf{Ni -- Cu} & \textbf{Zn -- Cu} & \textbf{Ni -- Sn} \\
\hline
\textbf{Sức điện động chuẩn (V)} & 0,59 & 1,10 & 0,11 \\
\hline
\end{tabular}
\end{center}
\begin{enumerate}[label=\alph*),leftmargin=20pt,itemsep=1pt]
\item Viết phương trình ion diễn ra khi pin Zn -- Sn hoạt động.
\item Sức điện động chuẩn của pin Zn -- Sn bằng bao nhiêu V(volt)?
\end{enumerate}
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=tableGreen!30,colframe=mainGreen!70,title=\textbf{Hướng dẫn giải chi tiết Câu 3},fonttitle=\bfseries]
\textbf{a) Phương trình ion diễn ra khi pin Zn -- Sn hoạt động:}
\begin{itemize}[leftmargin=15pt,itemsep=1pt]
\item So sánh thế điện cực chuẩn: Cặp $\ce{Zn^2+}/\ce{Zn}$ có tính khử mạnh hơn cặp $\ce{Sn^2+}/\ce{Sn}$ ($E^\circ(\ce{Zn^2+}/\ce{Zn}) < E^\circ(\ce{Sn^2+}/\ce{Sn})$).
\item Do đó, tại anode (cực âm Zn) xảy ra quá trình oxi hóa: $\ce{Zn(s) -> Zn^2+(aq) + 2e}$.
\item Tại cathode (cực dương Sn) xảy ra quá trình khử: $\ce{Sn^2+(aq) + 2e -> Sn(s)}$.
\item Phương trình ion diễn ra trong pin:
\[
\ce{Zn(s) + Sn^2+(aq) -> Zn^2+(aq) + Sn(s)}
\]
\end{itemize}

\textbf{b) Tính sức điện động chuẩn của pin Zn -- Sn:}
\begin{itemize}[leftmargin=15pt,itemsep=1pt]
\item Theo công thức sức điện động chuẩn của pin Galvani:
\begin{align}
E^\circ_\text{pin}(\text{Zn--Cu}) &= E^\circ(\ce{Cu^2+}/\ce{Cu}) - E^\circ(\ce{Zn^2+}/\ce{Zn}) = 1,10\text{ V} \\
E^\circ_\text{pin}(\text{Ni--Cu}) &= E^\circ(\ce{Cu^2+}/\ce{Cu}) - E^\circ(\ce{Ni^2+}/\ce{Ni}) = 0,59\text{ V} \\
E^\circ_\text{pin}(\text{Ni--Sn}) &= E^\circ(\ce{Sn^2+}/\ce{Sn}) - E^\circ(\ce{Ni^2+}/\ce{Ni}) = 0,11\text{ V}
\end{align}
\item Lấy biểu thức (1) trừ biểu thức (2):
\[
E^\circ(\ce{Ni^2+}/\ce{Ni}) - E^\circ(\ce{Zn^2+}/\ce{Zn}) = 1,10 - 0,59 = 0,51\text{ V}
\]
\item Cộng với biểu thức (3):
\begin{align*}
E^\circ_\text{pin}(\text{Zn--Sn}) &= E^\circ(\ce{Sn^2+}/\ce{Sn}) - E^\circ(\ce{Zn^2+}/\ce{Zn}) \\
&= [E^\circ(\ce{Sn^2+}/\ce{Sn}) - E^\circ(\ce{Ni^2+}/\ce{Ni})] + [E^\circ(\ce{Ni^2+}/\ce{Ni}) - E^\circ(\ce{Zn^2+}/\ce{Zn})] \\
&= 0,11 + 0,51 = \mathbf{0,62\text{ (V)}}
\end{align*}
\end{itemize}
\textbf{Kết luận:} Sức điện động chuẩn của pin Zn -- Sn là $\mathbf{0,62\text{ V}}$.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{}{8}{\noindent\makebox[\linewidth]{\dotfill}\\[14pt]}
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
    tex_hs = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_CK1_Hoa_12_DangHuyTru_Ma104_HocSinh.tex")
    tex_gv = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_CK1_Hoa_12_DangHuyTru_Ma104_GiaoVien.tex")

    print("Đang tạo mã nguồn LaTeX Bản Học sinh THPT Đặng Huy Trứ Hóa 12...")
    with open(tex_hs, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))

    print("Đang tạo mã nguồn LaTeX Bản Giáo viên THPT Đặng Huy Trứ Hóa 12...")
    with open(tex_gv, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))

    print("\nĐang biên dịch PDF Bản Học sinh THPT Đặng Huy Trứ Hóa 12...")
    compile_pdf(tex_hs)

    print("\nĐang biên dịch PDF Bản Giáo viên THPT Đặng Huy Trứ Hóa 12...")
    compile_pdf(tex_gv)

if __name__ == "__main__":
    main()
