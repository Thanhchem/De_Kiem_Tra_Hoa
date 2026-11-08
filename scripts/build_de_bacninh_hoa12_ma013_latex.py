# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho:
Đề Thi Thử Tốt Nghiệp THPT Quốc Gia Lần 1 (Năm Học 2026 - 2027)
Sở Giáo Dục và Đào Tạo Bắc Ninh - Môn: Hoá Học - Mã Đề 013
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)

Xuất bản:
1. Bản Học sinh: De_Kiem_Tra_Hoa_12_BacNinh_Ma013_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_Hoa_12_BacNinh_Ma013_GiaoVien.pdf
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
IMG_REL_DIR = "images_bacninh_12_ma013"

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
\textbf{SỞ GIÁO DỤC VÀ ĐÀO TẠO BẮC NINH}\\
\textbf{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 04 trang - 28 câu hỏi)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}THI THỬ TN THPT QUỐC GIA LẦN 1 -- NĂM HỌC 2026-2027}\\
\textbf{Môn: HOÁ HỌC -- LỚP 12}\\
\textit{Thời gian: 50 phút (không tính thời gian phát đề)}\\
\textbf{MÃ ĐỀ: 013}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ, tên thí sinh:} \dotfill \textit{Lớp:} \dotfill \textit{Số báo danh:} \dotfill

\vspace{3pt}
\noindent
\textit{\footnotesize (Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; K = 39; S = 32; Cl = 35,5; Ba = 137)}

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
{\color{hfRed}\textbf{\large BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 013)}}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (18 CÂU)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} \\
\hline
\textbf{Đ/A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} \\
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} & \textbf{17} & \textbf{18} & \multicolumn{2}{c|}{} \\
\cline{1-9}
\textbf{Đ/A} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \multicolumn{2}{c|}{} \\
\cline{1-9}
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (4 CÂU)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu hỏi} & \textbf{Lệnh a)} & \textbf{Lệnh b)} & \textbf{Lệnh c)} & \textbf{Lệnh d)} \\
\hline
\textbf{Câu 19} & \textbf{\color{hfRed}Đúng} & \textbf{Sai} & \textbf{Sai} & \textbf{\color{hfRed}Đúng} \\
\hline
\textbf{Câu 20} & \textbf{\color{hfRed}Đúng} & \textbf{Sai} & \textbf{\color{hfRed}Đúng} & \textbf{\color{hfRed}Đúng} \\
\hline
\textbf{Câu 21} & \textbf{\color{hfRed}Đúng} & \textbf{Sai} & \textbf{\color{hfRed}Đúng} & \textbf{Sai} \\
\hline
\textbf{Câu 22} & \textbf{Sai} & \textbf{\color{hfRed}Đúng} & \textbf{Sai} & \textbf{\color{hfRed}Đúng} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (6 CÂU)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Câu 23} & \textbf{Câu 24} & \textbf{Câu 25} & \textbf{Câu 26} & \textbf{Câu 27} & \textbf{Câu 28} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}1256} & \textbf{\color{hfRed}3241} & \textbf{\color{hfRed}42} & \textbf{\color{hfRed}1686} & \textbf{\color{hfRed}29,9} & \textbf{\color{hfRed}3} \\
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
{\color{mainGreen}\textbf{\large PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{2pt}
""")

    # Câu 1
    tex.append(r"""
\begin{ex}
Chất nào sau đây được sử dụng làm chất giặt rửa tổng hợp?
\choice
{$\ce{CH3COOK}$}
{\True $\ce{CH3[CH2]11OSO3Na}$}
{$\ce{C15H31COONa}$}
{$\ce{C15H31COOCH3}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$\ce{CH3[CH2]11OSO3Na}$ (sodium lauryl sulfate) là chất giặt rửa tổng hợp gốc alkyl sulfate.\\
$\ce{CH3COOK}$ và $\ce{C15H31COONa}$ là các muối carboxylate của xà phòng.\\
$\ce{C15H31COOCH3}$ là ester (methyl palmitate).\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2
    tex.append(r"""
\begin{ex}
Cho dung dịch $\ce{H2SO4}$ loãng vào dung dịch X thấy sủi bọt khí và sinh ra kết tủa. Dung dịch X là
\choice
{$\ce{Na2SO4}$}
{$\ce{KNO3}$}
{\True $\ce{Ba(HCO3)2}$}
{$\ce{BaCl2}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phản ứng: $\ce{Ba(HCO3)2 + H2SO4 -> BaSO4 v + 2CO2 ^ + 2H2O}$.\\
Vừa tạo kết tủa trắng $\ce{BaSO4}$, vừa sủi bọt khí $\ce{CO2}$.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 3
    tex.append(r"""
\begin{ex}
Một dung dịch baking soda có $\text{pH} = 8{,}3$. Nồng độ ion $\ce{OH-}$ trong dung dịch trên là
\choice
{$5{,}7\text{ M}$}
{$8{,}3\text{ M}$}
{$10^{-8{,}3}\text{ M}$}
{\True $10^{-5{,}7}\text{ M}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$\text{pH} = 8{,}3 \implies \text{pOH} = 14 - 8{,}3 = 5{,}7 \implies [\ce{OH-}] = 10^{-\text{pOH}} = 10^{-5{,}7}\text{ M}$.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 4
    tex.append(r"""
\begin{ex}
Phát biểu nào sau đây đúng?
\choice
{Amylose và amylopectin đều có cấu trúc mạch phân nhánh}
{Trong phân tử glucose chỉ có 4 nhóm alcohol (OH)}
{\True Saccharose không tham gia phản ứng tráng bạc}
{Glucose và fructose đều làm mất màu nước bromine}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
A sai vì amylose có cấu trúc mạch không phân nhánh.\\
B sai vì glucose có 5 nhóm -OH kề nhau.\\
C đúng vì saccharose không có nhóm -OH hemiacetal/hemiketal tự do để mở vòng nên không tráng bạc.\\
D sai vì fructose không làm mất màu nước bromine.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 5
    tex.append(r"""
\begin{ex}
Cho 4 dung dịch đựng trong 4 lọ mất nhãn sau: $\ce{HCl}$, $\ce{Na2SO4}$, $\ce{NaCl}$, $\ce{Ba(OH)2}$. Chỉ dùng thêm 1 thuốc thử nào dưới đây để nhận biết các dung dịch trên?
\choice
{\True quỳ tím}
{$\ce{H2SO4}$}
{$\ce{AgNO3}$}
{$\ce{BaCl2}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Dùng quỳ tím:\\
- Quỳ hóa đỏ: $\ce{HCl}$.\\
- Quỳ hóa xanh: $\ce{Ba(OH)2}$.\\
- Không đổi màu: $\ce{Na2SO4}$ và $\ce{NaCl}$.\\
Lấy $\ce{Ba(OH)2}$ vừa nhận biết nhỏ vào 2 mẫu còn lại: mẫu xuất hiện kết tủa trắng là $\ce{Na2SO4}$ ($\ce{BaSO4}\downarrow$), mẫu không hiện tượng là $\ce{NaCl}$.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 6
    tex.append(r"""
\begin{ex}
Chất nào sau đây thuộc loại acid béo omega-3?
\begin{center}
\includegraphics[width=13.5cm]{images_bacninh_12_ma013/cau6_acid_beo.png}
\end{center}
\choice
{3}
{\True 4}
{2}
{1}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Acid béo omega-3 ($\omega$-3) là acid béo không no đa liên kết đôi có liên kết đôi đầu tiên ở vị trí C số 3 tính từ đầu methyl ($\ce{-CH3}$, tức vị trí $\omega$).\\
- Chất (1): oleic acid ($\omega$-9).\\
- Chất (2): arachidonic acid ($\omega$-6).\\
- Chất (3): linoleic acid ($\omega$-6).\\
- Chất (4): $\alpha$-linolenic acid (ALA), liên kết đôi đầu tiên bắt đầu từ C số 3 tính từ đầu methyl ($\omega$-3).\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 7
    tex.append(r"""
\begin{ex}
Thủy phân hoàn toàn hỗn hợp phenyl acetate và ethyl acetate trong dung dịch $\ce{NaOH}$, thu được sản phẩm gồm
\choice
{\True 2 muối và 1 alcohol}
{2 muối và 2 alcohol}
{1 muối và 1 alcohol}
{1 muối và 2 alcohol}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$\ce{CH3COOC6H5 + 2NaOH -> CH3COONa + C6H5ONa + H2O}$ (2 muối, 0 alcohol).\\
$\ce{CH3COOC2H5 + NaOH -> CH3COONa + C2H5OH}$ (1 muối, 1 alcohol).\\
Tổng sản phẩm gồm 2 muối ($\ce{CH3COONa}$, $\ce{C6H5ONa}$) và 1 alcohol ($\ce{C2H5OH}$).\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 8
    tex.append(r"""
\begin{ex}
Ester X có công thức phân tử là $\ce{C4H8O2}$, X có khả năng tham gia phản ứng với dung dịch $\ce{AgNO3}$ trong $\ce{NH3}$. Số công thức cấu tạo của ester X thỏa mãn điều kiện trên là
\choice
{3}
{4}
{\True 2}
{1}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
X là ester của formic acid: $\ce{HCOOC3H7}$.\\
Gốc $\ce{-C3H7}$ có 2 đồng phân: $\ce{HCOOCH2CH2CH3}$ (propyl formate) và $\ce{HCOOCH(CH3)2}$ (isopropyl formate).\\
Vậy có 2 công thức cấu tạo thỏa mãn.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 9
    tex.append(r"""
\begin{ex}
Các khí sinh ra trong thí nghiệm phản ứng của saccharose với dung dịch $\ce{H2SO4}$ đặc bao gồm
\choice
{$\ce{SO2}$, $\ce{H2S}$}
{\True $\ce{CO2}$, $\ce{SO2}$}
{$\ce{CO2}$, $\ce{SO3}$}
{$\ce{CO2}$, $\ce{H2S}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Saccharose bị than hóa: $\ce{C12H22O11 ->[H2SO4 đặc] 12C + 11H2O}$.\\
Carbon phản ứng với $\ce{H2SO4}$ đặc nóng: $\ce{C + 2H2SO4 đặc -> CO2 ^ + 2SO2 ^ + 2H2O}$.\\
Khí sinh ra là $\ce{CO2}$ và $\ce{SO2}$.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 10
    tex.append(r"""
\begin{ex}
Để tráng một lớp bạc lên ruột phích, người ta cho chất X phản ứng với lượng dư dung dịch $\ce{AgNO3}$ trong $\ce{NH3}$, đun nóng nhẹ 60--70$^\circ\text{C}$. Chất X là
\choice
{tinh bột}
{ethyl acetate}
{saccharose}
{\True glucose}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Trong công nghiệp tráng gương ruột phích, người ta dùng glucose vì không độc hại và giá thành hợp lý.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 11
    tex.append(r"""
\begin{ex}
Trong quá trình trồng trọt, người nông dân được khuyến cáo không bón vôi sống (thành phần chính là $\ce{CaO}$) cùng với phân đạm ammonium. Nguyên nhân của khuyến cáo này là
\choice
{\True thất thoát đạm vì giải phóng ammonia}
{tạo thành hỗn hợp gây cháy nổ}
{tạo acid làm ảnh hưởng tới cây trồng}
{làm tăng độ chua của đất}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$\ce{CaO + H2O -> Ca(OH)2}$.\\
$\ce{Ca(OH)2 + 2NH4+ -> Ca^2+ + 2NH3 ^ + 2H2O}$.\\
Khí $\ce{NH3}$ bay hơi làm thất thoát lượng đạm cung cấp cho cây.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 12
    tex.append(r"""
\begin{ex}
Công thức phân tử nào sau đây không thể của ester?
\choice
{$\ce{C3H6O2}$}
{\True $\ce{C3H8O2}$}
{$\ce{C3H4O2}$}
{$\ce{C4H6O2}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phân tử $\ce{C3H8O2}$ có độ bất bão hòa $k = \dfrac{2\times 3 + 2 - 8}{2} = 0$, không có liên kết $\pi$ hay vòng nên không thể chứa nhóm chức ester $\ce{-COO-}$ (cần $k \ge 1$).\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 13
    tex.append(r"""
\begin{ex}
Polysaccharide X là chất rắn, màu trắng, dạng sợi, trong bông nõn có gần 98\% chất X. Chất X là
\choice
{glucose}
{\True cellulose}
{maltose}
{saccharose}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Cellulose là polysaccharide dạng sợi, chiếm gần 98\% khối lượng trong bông nõn.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 14
    tex.append(r"""
\begin{ex}
Có nhiều vụ tai nạn giao thông xảy ra do người lái xe uống rượu. Hàm lượng alcohol ethanol trong máu người lái xe không được vượt quá 0,02\% theo khối lượng. Để xác định hàm lượng đó ta chuẩn độ alcohol bằng $\ce{K2Cr2O7}$ trong môi trường acid theo phản ứng:\\
$3\ce{CH3CH2OH} + \ce{K2Cr2O7} + 4\ce{H2SO4} \rightarrow 3\ce{CH3CHO} + \ce{Cr2(SO4)3} + \ce{K2SO4} + 7\ce{H2O}$\\
Khi chuẩn độ 25,0 gam huyết tương máu của một người lái xe cần dùng 20,0 mL dung dịch $\ce{K2Cr2O7}$ 0,010M. Nồng độ phần trăm ethanol trong 25 gam huyết tương máu của một người lái xe là
\choice
{\True 0,11\%}
{0,03\%}
{0,05\%}
{0,027\%}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$n_{\ce{K2Cr2O7}} = 0{,}020 \times 0{,}010 = 2{,}0 \cdot 10^{-4}\text{ mol}$.\\
Theo phương trình: $n_{\ce{C2H5OH}} = 3 \times n_{\ce{K2Cr2O7}} = 6{,}0 \cdot 10^{-4}\text{ mol}$.\\
$m_{\ce{C2H5OH}} = 6{,}0 \cdot 10^{-4} \times 46 = 0{,}0276\text{ gam}$.\\
$C\% = \dfrac{0{,}0276}{25{,}0} \times 100\% = 0{,}1104\% \approx \mathbf{0{,}11\%}$.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 15
    tex.append(r"""
\begin{ex}
Cho phản ứng hóa học sau: $\ce{CH3COOH(l) + CH3OH(l) <=> CH3COOCH3(l) + H2O(l)}$. Biểu thức tính hằng số cân bằng nồng độ $K_{\text{C}}$ của phản ứng trên là
\choice
{\True $K_{\text{C}} = \dfrac{[\ce{CH3COOCH3}][\ce{H2O}]}{[\ce{CH3COOH}][\ce{CH3OH}]}$}
{$K_{\text{C}} = \dfrac{[\ce{CH3COOCH3}]}{[\ce{CH3COOH}][\ce{CH3OH}]}$}
{$K_{\text{C}} = \dfrac{[\ce{CH3COOH}][\ce{CH3OH}]}{[\ce{CH3COOCH3}][\ce{H2O}]}$}
{$K_{\text{C}} = \dfrac{[\ce{CH3COOH}][\ce{CH3OH}]}{[\ce{CH3COOCH3}]}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Trong hệ phản ứng ester hóa lỏng đồng thể (không dùng dung môi nước), $\ce{H2O(l)}$ là một cấu tử trong hệ cân bằng nồng độ biến thiên nên tham gia vào biểu thức hằng số cân bằng:\\
$K_{\text{C}} = \dfrac{[\ce{CH3COOCH3}][\ce{H2O}]}{[\ce{CH3COOH}][\ce{CH3OH}]}$.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 16
    tex.append(r"""
\begin{ex}
Methyl acrylate là một chất kích thích mạnh, có thể gây chóng mặt, đau đầu, hoa mắt và khó thở khi tiếp xúc với da hoặc hít phải, ester này có công thức cấu tạo thu gọn là
\choice
{$\ce{CH3COOC2H5}$}
{$\ce{CH3COOCH3}$}
{$\ce{C2H5COOCH3}$}
{\True $\ce{CH2=CHCOOCH3}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Gốc acrylate là $\ce{CH2=CH-COO-}$, gốc methyl là $\ce{-CH3}$.\\
Công thức thu gọn của methyl acrylate là $\ce{CH2=CHCOOCH3}$.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 17
    tex.append(r"""
\begin{ex}
Phương trình hoá học của phản ứng hydrate hóa ethylene để điều chế ethanol là: $\ce{CH2=CH2 + H2O -> CH3CH2OH}$. Giai đoạn (1) trong cơ chế của phản ứng trên xảy ra như sau:
\begin{center}
\includegraphics[width=7.5cm]{images_bacninh_12_ma013/cau17_co_che.png}
\end{center}
Nhận định nào sau đây \textbf{không đúng}?
\choice
{Phản ứng hydrate hóa ethylene là phản ứng cộng}
{Trong giai đoạn (1) có sự phân cắt liên kết $\pi$}
{Trong giai đoạn (1) có sự hình thành liên kết $\sigma$}
{\True Trong phân tử ethylene có 6 liên kết $\sigma$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phân tử ethylene $\ce{H2C=CH2}$ có 4 liên kết $\sigma_{\ce{C-H}}$ và 1 liên kết đôi $\ce{C=C}$ (gồm 1 liên kết $\sigma$ và 1 liên kết $\pi$). Tổng cộng chỉ có 5 liên kết $\sigma$, không phải 6.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 18
    tex.append(r"""
\begin{ex}
Tiến hành thí nghiệm theo các bước sau:\\
• Bước 1: Cho khoảng 2mL dung dịch $\ce{AgNO3}$ 1\% vào ống nghiệm sạch.\\
• Bước 2: Thêm từ từ từng giọt dung dịch ammonia 5\%, lắc đều cho đến khi kết tủa tan hết. Dung dịch thu được là thuốc thử Tollens.\\
• Bước 3: Thêm tiếp khoảng 2mL dung dịch glucose 2\% lắc đều. Sau đó, ngâm ống nghiệm vào cốc thuỷ tinh chứa nước nóng trong vài phút.\\
Phát biểu nào sau đây sai?
\choice
{Sản phẩm hữu cơ thu được sau bước 3 là ammonium gluconate}
{\True Thí nghiệm trên chứng minh glucose có tính chất của polyalcohol}
{Sau bước 3, có lớp bạc (silver) kim loại bám trên thành ống nghiệm}
{Trong phản ứng ở bước 3, glucose đóng vai trò là chất khử}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phản ứng tráng bạc chứng minh glucose có tính chất của nhóm aldehyde ($\ce{-CHO}$), không chứng minh tính chất của polyalcohol.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN II
    tex.append(r"""
\vspace{6pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN II. Câu trắc nghiệm đúng sai}}\\
\textit{\small Thí sinh trả lời từ câu 19 đến câu 22. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{2pt}
""")

    # Câu 19
    tex.append(r"""
\begin{ex}
Methyl cinnamate là một ester có công thức phân tử $\ce{C10H10O2}$ và có mùi thơm của dâu tây (strawberry) được sử dụng trong ngành công nghiệp hương liệu và nước hoa.
\begin{center}
\includegraphics[width=5.5cm]{images_bacninh_12_ma013/cau19_methyl_cinnamate.png}
\end{center}
\choiceTF
{\True Methyl cinnamate có đồng phân hình học.}
{Methyl cinnamate có công thức cấu tạo là $\ce{CH3COO-CH=CH-C6H5}$.}
{Methyl cinnamate phản ứng với $\ce{NaOH}$ với tỉ lệ $1 : 2$.}
{\True Để điều chế 24,3 gam ester methyl cinnamate, người ta cho 29,6 gam cinnamic acid phản ứng với lượng dư methyl alcohol. Hiệu suất phản ứng ester hóa trong trường hợp này là 75\%.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) \textbf{Đúng.} Liên kết đôi $\ce{-CH=CH-}$ có cấu hình cis/trans do 2 nhóm thế ở mỗi C khác nhau.\\
b) \textbf{Sai.} Công thức đúng là $\ce{C6H5-CH=CH-COOCH3}$.\\
c) \textbf{Sai.} Đây là ester của alcohol đơn chức nên tỉ lệ phản ứng với $\ce{NaOH}$ là $1 : 1$.\\
d) \textbf{Đúng.} $n_{\text{acid}} = \dfrac{29{,}6}{148} = 0{,}2\text{ mol} \implies m_{\text{lt}} = 0{,}2 \times 162 = 32{,}4\text{ gam} \implies H = \dfrac{24{,}3}{32{,}4} \times 100\% = 75\%$.
}
""")
    tex.append(r"\end{ex}")

    # Câu 20
    tex.append(r"""
\begin{ex}
Tiến hành thí nghiệm điều chế ethyl acetate theo các bước sau đây:\\
• Bước 1: Cho 1mL $\ce{C2H5OH}$, 1mL $\ce{CH3COOH}$ và vài giọt dung dịch $\ce{H2SO4}$ đặc vào ống nghiệm.\\
• Bước 2: Lắc đều ống nghiệm, đun cách thủy (trong nồi nước nóng) khoảng 5 -- 6 phút ở 60 -- 70$^\circ\text{C}$.\\
• Bước 3: Làm lạnh, sau đó rót 2mL dung dịch $\ce{NaCl}$ bão hòa vào ống nghiệm.
\choiceTF
{\True Ở bước 1, có thể thay bằng dung dịch $\ce{H3PO4}$ đậm đặc.}
{Mục đích chính của việc thêm dung dịch $\ce{NaCl}$ bão hòa là để tránh phân hủy sản phẩm.}
{\True Sau bước 3, chất lỏng trong ống nghiệm tách thành hai lớp.}
{\True $\ce{H2SO4}$ đặc có vai trò vừa làm chất xúc tác vừa làm tăng hiệu suất tạo sản phẩm.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) \textbf{Đúng.} $\ce{H3PO4}$ đậm đặc là acid vô cơ có tính hút nước, có thể dùng làm xúc tác ester hóa thay cho $\ce{H2SO4}$ đặc.\\
b) \textbf{Sai.} Thêm $\ce{NaCl}$ bão hòa làm tăng khối lượng riêng của dung dịch và giảm độ tan của ester, giúp ester tách lớp.\\
c) \textbf{Đúng.} Ethyl acetate nhẹ hơn và không tan nên nổi lên trên tách thành 2 lớp.\\
d) \textbf{Đúng.} $\ce{H2SO4}$ đặc xúc tác và hút nước làm cân bằng dịch chuyển sang chiều thuận.
}
""")
    tex.append(r"\end{ex}")

    # Câu 21
    tex.append(r"""
\begin{ex}
Hợp chất hữu cơ mạch hở X ($\ce{C7H10O5}$) tác dụng với lượng dư dung dịch $\ce{NaOH}$ đun nóng thu được glycerol và hỗn hợp 2 muối carboxylate Y và Z ($M_{\text{Y}} < M_{\text{Z}}$).
\choiceTF
{\True Y tham gia phản ứng tráng bạc.}
{Có 2 công thức cấu tạo thỏa mãn tính chất của X.}
{\True X là hợp chất hữu cơ tạp chức.}
{Carboxylic acid của muối Z có đồng phân hình học.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- Độ bất bão hòa của X ($\ce{C7H10O5}$): $k = 3$. X tác dụng $\ce{NaOH}$ cho glycerol (3C) và 2 muối Y, Z.\\
$\implies$ X là ester 2 chức có 1 nhóm $\ce{-OH}$ tự do. Tổng số C của 2 muối là $7 - 3 = 4\text{ C}$.\\
- Vì $M_{\text{Y}} < M_{\text{Z}}$ và có 1 liên kết $\pi$ ở gốc acid nên: Y là $\ce{HCOONa}$, Z là $\ce{CH2=CH-COONa}$.\\
a) \textbf{Đúng.} $\ce{HCOONa}$ tham gia phản ứng tráng bạc.\\
b) \textbf{Sai.} Có 3 công thức cấu tạo: nhóm $\ce{-OH}$ ở vị trí C1 (2 đồng phân do hoán đổi 2 nhóm ester) hoặc ở vị trí C2 (1 đồng phân).\\
c) \textbf{Đúng.} X vừa có nhóm ester vừa có nhóm alcohol $\ce{-OH}$.\\
d) \textbf{Sai.} Acid của Z là acrylic acid $\ce{CH2=CH-COOH}$, không có đồng phân hình học.
}
""")
    tex.append(r"\end{ex}")

    # Câu 22
    tex.append(r"""
\begin{ex}
Thực vật sử dụng quá trình quang hợp để chuyển đổi năng lượng ánh sáng từ mặt trời thành năng lượng hóa học. Trong quá trình quang hợp xảy ra phản ứng giữa khí carbonic và nước theo phương trình hóa học:\\
$6\ce{CO2(g) + 6H2O(l) -> C6H12O6(aq) + 6O2(g)}$\\
Cho enthalpy tạo thành chuẩn của $\ce{CO2(g)} = -393{,}5\text{ kJ/mol}$; $\ce{H2O(l)} = -285{,}8\text{ kJ/mol}$; $\ce{C6H12O6(aq)} = -1271\text{ kJ/mol}$.
\choiceTF
{Phản ứng quang hợp của cây xanh là phản ứng tỏa nhiệt.}
{\True Nhiệt tạo thành chuẩn của $\ce{O2(g)}$ bằng 0.}
{Biến thiên enthalpy chuẩn của phản ứng trên có giá trị là $-2804{,}8\text{ kJ}$.}
{\True Trồng nhiều cây xanh sẽ làm giảm được nhiệt độ của môi trường xung quanh.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) \textbf{Sai.} Quang hợp hấp thụ năng lượng ánh sáng mặt trời nên là phản ứng thu nhiệt ($\Delta_r H^\circ_{298} > 0$).\\
b) \textbf{Đúng.} $\ce{O2(g)}$ là đơn chất bền ở điều kiện chuẩn nên $\Delta_f H^\circ_{298} = 0$.\\
c) \textbf{Sai.} $\Delta_r H^\circ_{298} = -1271 - [6\times(-393{,}5) + 6\times(-285{,}8)] = +2804{,}8\text{ kJ}$ (dấu dương, không phải dấu âm).\\
d) \textbf{Đúng.} Quá trình quang hợp thu nhiệt và thoát hơi nước làm mát không khí xung quanh.
}
""")
    tex.append(r"\end{ex}")

    # PHẦN III
    tex.append(r"""
\vspace{6pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN III. Câu trắc nghiệm trả lời ngắn}}\\
\textit{\small Thí sinh trả lời từ câu 23 đến câu 28.}
\vspace{2pt}
""")

    # Câu 23
    tex.append(r"""
\begin{ex}
Có các phát biểu sau:\\
(1) Thủy phân chất béo trong môi trường kiềm luôn thu được glycerol.\\
(2) Triolein làm mất màu nước bromine.\\
(3) Các ester chỉ được điều chế từ carboxylic acid và alcohol.\\
(4) Benzyl acetate là ester có mùi chuối chín.\\
(5) Đốt cháy ethyl acetate thu được số mol nước bằng số mol khí carbonic.\\
(6) Hydrogen hoá hoàn toàn chất béo lỏng (xúc tác Ni, $t^\circ$) thu được chất béo rắn.\\
Liệt kê theo trình tự từ nhỏ đến lớn các số ứng với các phát biểu đúng trên (ví dụ: 56, 124, 3456,...).
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Các phát biểu đúng: (1), (2), (5), (6).\\
(3) Sai vì có thể điều chế từ acid và alkyne (vinyl acetate) hoặc acid anhydride.\\
(4) Sai vì benzyl acetate có mùi hoa nhài; isoamyl acetate mới có mùi chuối chín.\\
\textbf{Đáp án: 1256}
}
""")
    else:
        tex.append(r"""
\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 24
    tex.append(r"""
\begin{ex}
Cho các phản ứng được đánh số thứ tự như sau:\\
(1) $6n\ce{CO2} + 5n\ce{H2O} \xrightarrow{\text{ánh sáng}} (\ce{C6H10O5})_n + 6n\ce{O2}$\\
(2) $(\ce{C6H10O5})_n + n\ce{H2O} \xrightarrow{\text{enzyme}} n\ce{C6H12O6}$\\
(3) $\ce{C6H12O6} \xrightarrow{\text{enzyme}} 2\ce{C2H5OH} + 2\ce{CO2}$\\
(4) $\ce{C2H5OH} + 3\ce{O2} \rightarrow 2\ce{CO2} + 3\ce{H2O}$\\
Gán số thứ tự phương trình hoá học của các phản ứng theo tên gọi: lên men rượu, thủy phân, cháy, quang hợp và sắp xếp theo trình tự thành dãy bốn số.
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- Lên men rượu: (3)\\
- Thủy phân: (2)\\
- Cháy: (4)\\
- Quang hợp: (1)\\
\textbf{Đáp án: 3241}
}
""")
    else:
        tex.append(r"""
\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 25
    tex.append(r"""
\begin{ex}
Một nhà máy đường trung bình mỗi ngày sử dụng 3000 tấn mía để sản xuất. Sau khi nước mía được cô đặc và kết tinh để lấy đường cát, phần chất lỏng màu nâu sẫm, đặc sệt còn lại chính là rỉ đường. Tỷ lệ rỉ đường bằng 4\% khối lượng mía ban đầu. Thành phần chính của rỉ đường gồm: Saccharose (khoảng 40\%), monosaccharide (tổng glucose và fructose khoảng 15\%), nước và các khoáng chất. Nhằm tận dụng nguồn nguyên liệu, nhà máy dùng rỉ đường để sản xuất ethanol, người ta pha loãng rỉ đường và tiến hành lên men bằng nấm men (Saccharomyces cerevisiae). Quá trình gồm thuỷ phân saccharose và lên men monosaccharide (cả glucose và fructose đều lên men tạo ethanol). Hiệu suất thuỷ phân Saccharose và lên men đều đạt 90\%, các thành phần khác không tham gia quá trình tạo ethanol. Thể tích dung dịch ethanol 88$^\circ$ trung bình nhà máy sản xuất được trong 1 ngày là $V$ ($\text{m}^3$), biết khối lượng riêng của ethanol nguyên chất là 0,79 g/mL. Giá trị của $V$ bằng bao nhiêu? (Không làm tròn các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng đơn vị).
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- Khối lượng rỉ đường $= 3000 \times 4\% = 120\text{ tấn}$.\\
- $m_{\text{saccharose}} = 120 \times 40\% = 48\text{ tấn} \implies n_{\text{saccharose}} = \dfrac{48}{342}\text{ tấn mol}$.\\
  $n_{\text{mono sinh ra}} = 2 \times \dfrac{48}{342} \times 90\% = \dfrac{86{,}4}{342}\text{ tấn mol}$.\\
- $m_{\text{mono ban đầu}} = 120 \times 15\% = 18\text{ tấn} \implies n_{\text{mono ban đầu}} = \dfrac{18}{180} = 0{,}1\text{ tấn mol}$.\\
- Tổng $n_{\text{mono}} = 0{,}1 + \dfrac{86{,}4}{342} \approx 0{,}35263158\text{ tấn mol}$.\\
- $n_{\ce{C2H5OH}} = 2 \times 0{,}35263158 \times 90\% \approx 0{,}63473684\text{ tấn mol}$.\\
- $m_{\ce{C2H5OH}} = 0{,}63473684 \times 46 \approx 29{,}197895\text{ tấn} = 29\,197\,895\text{ g}$.\\
- $V_{\text{nguyên chất}} = \dfrac{29\,197\,895}{0{,}79} \approx 36\,959\,360\text{ mL} \approx 36{,}95936\text{ m}^3$.\\
- $V_{\text{dd 88}^\circ} = \dfrac{36{,}95936 \times 100}{88} \approx 41{,}9993\text{ m}^3 \approx \mathbf{42\text{ m}^3}$.\\
\textbf{Đáp án: 42}
}
""")
    else:
        tex.append(r"""
\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 26
    tex.append(r"""
\begin{ex}
Xác định giá trị số sóng $\nu$ ($\text{cm}^{-1}$) hấp thụ đặc trưng của nhóm $\ce{C=O}$ (ketone) trên phổ hồng ngoại IR của methyl phenyl ketone cho dưới đây:
\begin{center}
\includegraphics[width=10.5cm]{images_bacninh_12_ma013/cau26_pho_ir.png}
\end{center}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Trên phổ hồng ngoại IR, tín hiệu hấp thụ đặc trưng của nhóm carbonyl $\ce{C=O}$ liên hợp trong methyl phenyl ketone là pic có cường độ mạnh và nhọn nhất ở số sóng 1686 $\text{cm}^{-1}$.\\
\textbf{Đáp án: 1686}
}
""")
    else:
        tex.append(r"""
\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 27
    tex.append(r"""
\begin{ex}
Chỉ số xà phòng hoá của chất béo là số mg KOH cần dùng để xà phòng hoá hoàn toàn 1 gam chất béo (bao gồm cả trung hoà acid béo tự do có trong chất béo). Cho bảng số liệu:
\begin{center}
\renewcommand{\arraystretch}{1.15}
\begin{tabular}{|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Chất béo} & \textbf{Chỉ số xà phòng hoá} & \textbf{Chất béo} & \textbf{Chỉ số xà phòng hoá} \\
\hline
Dầu dừa & 257 & Dầu olive & 188 \\
\hline
\textbf{Dầu cọ} & \textbf{199} & Mỡ vịt & 194 \\
\hline
Dầu lạc & 192 & Mỡ gà & 195 \\
\hline
Dầu mè & 188 & Mỡ lợn & 198 \\
\hline
\end{tabular}
\end{center}
Giả sử quá trình thực hành điều chế xà phòng, học sinh được yêu cầu điều chế xà phòng từ 200 gam dầu cọ và $\ce{NaOH}$ (rắn). Biết xà phòng được điều chế bằng cách cho chất béo tác dụng vừa đủ với dung dịch $\ce{NaOH}$. Cần cân $m$ (gam) hoá chất từ lọ đựng $\ce{NaOH}$ rắn trong phòng thí nghiệm, biết rằng loại $\ce{NaOH}$ được dùng trong phòng thí nghiệm đã hút ẩm nên lượng $\ce{NaOH}$ chiếm 95\%. Giá trị của $m$ bằng bao nhiêu? (Các quá trình tính toán không làm tròn, kết quả cuối cùng làm tròn đến hàng phần mười).
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- Dầu cọ có chỉ số xà phòng hóa là 199 $\implies 200\text{ g}$ dầu cọ cần lượng KOH là:\\
  $m_{\ce{KOH}} = 200 \times 199\text{ mg} = 39\,800\text{ mg} = 39{,}8\text{ g}$.\\
- $n_{\ce{KOH}} = \dfrac{39{,}8}{56}\text{ mol} \implies n_{\ce{NaOH}} = \dfrac{39{,}8}{56}\text{ mol}$.\\
- $m_{\ce{NaOH}\text{ nguyên chất}} = \dfrac{39{,}8}{56} \times 40 \approx 28{,}42857\text{ gam}$.\\
- Lọ hóa chất chứa 95\% $\ce{NaOH} \implies m = \dfrac{28{,}42857}{0{,}95} \approx 29{,}9248\text{ gam} \approx \mathbf{29{,}9\text{ gam}}$.\\
\textbf{Đáp án: 29,9}
}
""")
    else:
        tex.append(r"""
\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 28
    tex.append(r"""
\begin{ex}
Cho các chất: $\ce{CH3[CH2]14COONa}$, $\ce{CH3[CH2]10CH2OSO3Na}$, $\ce{CH3[CH2]16COOK}$, $\ce{CH3[CH2]7CH=CH[CH2]7COONa}$, $\ce{CH3COONa}$, $(\ce{C15H31COO})_3\ce{C3H5}$. Có bao nhiêu chất có thể là thành phần chính của xà phòng?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Xà phòng là muối sodium hoặc potassium của các acid béo:\\
(1) $\ce{CH3[CH2]14COONa}$ (sodium palmitate): thành phần xà phòng.\\
(2) $\ce{CH3[CH2]10CH2OSO3Na}$: chất giặt rửa tổng hợp.\\
(3) $\ce{CH3[CH2]16COOK}$ (potassium stearate): thành phần xà phòng mềm.\\
(4) $\ce{CH3[CH2]7CH=CH[CH2]7COONa}$ (sodium oleate): thành phần xà phòng.\\
(5) $\ce{CH3COONa}$: muối của acid mạch ngắn, không phải xà phòng.\\
(6) $(\ce{C15H31COO})_3\ce{C3H5}$: chất béo, là nguyên liệu.\\
Vậy có đúng 3 chất có thể là thành phần chính của xà phòng.\\
\textbf{Đáp án: 3}
}
""")
    else:
        tex.append(r"""
\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Kết thúc
    tex.append(r"""
\vspace{10pt}
\begin{center}
\textbf{------------------- HẾT -------------------}\\
\textit{\small (Học sinh không được sử dụng tài liệu. Cán bộ coi thi không giải thích gì thêm)}
\end{center}

\end{document}
""")

    return "".join(tex)


def compile_latex(tex_name, is_teacher=False):
    tex_content = build_latex_content(is_teacher=is_teacher)
    tex_path = os.path.join(SAN_PHAM_DIR, f"{tex_name}.tex")
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_content)
    print(f"Đã ghi file TeX: {tex_path}")

    # Copy ex_test.sty to SAN_PHAM_DIR if needed
    src_sty = r"C:\Antigravity_Thanh\ex_test.sty"
    dst_sty = os.path.join(SAN_PHAM_DIR, "ex_test.sty")
    if not os.path.exists(dst_sty) and os.path.exists(src_sty):
        import shutil
        shutil.copy2(src_sty, dst_sty)

    # Chạy pdflatex 2 lần để giải quyết số trang
    for run in range(1, 3):
        cmd = ["pdflatex", "-interaction=nonstopmode", f"{tex_name}.tex"]
        proc = subprocess.run(cmd, cwd=SAN_PHAM_DIR, capture_output=True, text=True)
        if proc.returncode != 0 and run == 2:
            print(f"Lỗi biên dịch lần {run}:")
            print(proc.stdout[-1500:])
            return False

    pdf_path = os.path.join(SAN_PHAM_DIR, f"{tex_name}.pdf")
    if os.path.exists(pdf_path):
        print(f"Biên dịch PDF thành công: {pdf_path}")

    # Dọn dẹp file tạm
    for ext in ["aux", "log", "thm", "out"]:
        for f in glob.glob(os.path.join(SAN_PHAM_DIR, f"{tex_name}.{ext}")):
            try:
                os.remove(f)
            except Exception:
                pass
    for f in glob.glob(os.path.join(SAN_PHAM_DIR, "ans-*.tex")):
        try:
            os.remove(f)
        except Exception:
            pass

    return True

if __name__ == "__main__":
    compile_latex("De_Kiem_Tra_Hoa_12_BacNinh_Ma013_HocSinh", is_teacher=False)
    compile_latex("De_Kiem_Tra_Hoa_12_BacNinh_Ma013_GiaoVien", is_teacher=True)
