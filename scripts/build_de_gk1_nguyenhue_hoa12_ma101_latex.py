# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho:
Đề Kiểm Tra Giữa Kì I Lớp 12 (Năm Học 2026 - 2027)
Trường THPT Nguyễn Huệ - Tổ Lí – Hóa – Công nghệ - MÃ ĐỀ: 101
Chuẩn 100% FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)

Xuất bản:
1. Bản Học sinh: San_Pham/De_Kiem_Tra_GK1_Hoa_12_THPT_NguyenHue_Ma101_HocSinh.pdf
2. Bản Giáo viên: San_Pham/De_Kiem_Tra_GK1_Hoa_12_THPT_NguyenHue_Ma101_GiaoVien.pdf
"""

import os
import sys
import subprocess
import glob

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

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
""")

    # Tiêu đề đề thi
    tex.append(r"""
\noindent
\begin{minipage}[t]{0.48\textwidth}
\textbf{TRƯỜNG THPT NGUYỄN HUỆ}\\
\textbf{TỔ LÍ -- HÓA -- CÔNG NGHỆ}\\
\textbf{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 03 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}ĐỀ KIỂM TRA GIỮA KÌ I -- NĂM HỌC 2026 -- 2027}\\
\textbf{MÔN: HOÁ HỌC -- LỚP 12}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{MÃ ĐỀ: 101}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ, tên thí sinh:} \dotfill \textit{Lớp:} \dots\dots\dots \textit{Số báo danh:} \dotfill

\vspace{3pt}
\noindent
\textit{\footnotesize (Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; S = 32; Cl = 35,5; K = 39; Ca = 40; Ag = 108).}

\vspace{2pt}
\hrule height 0.5pt
""")

    if is_teacher:
        tex.append(r"""
\vspace{4pt}
\begin{center}
\textbf{\color{hfRed}\large (BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}
\end{center}

\vspace{2pt}
\begin{center}
{\color{hfRed}\textbf{\large BẢNG ĐÁP ÁN THAM KHẢO NHANH (MÃ ĐỀ 101)}}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (12 CÂU -- 3,0 ĐIỂM)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} \\
\hline
\textbf{Đ/A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2 CÂU -- 2,0 ĐIỂM)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Lệnh a)} & \textbf{Lệnh b)} & \textbf{Lệnh c)} & \textbf{Lệnh d)} \\
\hline
\textbf{Câu 1} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}S} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}Đ} \\
\hline
\textbf{Câu 2} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}S} & \textbf{\color{hfRed}Đ} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (4 CÂU -- 2,0 ĐIỂM)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu hỏi} & \textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}3} & \textbf{\color{hfRed}4} & \textbf{\color{hfRed}2} & \textbf{\color{hfRed}16,3} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}4. PHẦN IV: TỰ LUẬN (3 CÂU -- 3,0 ĐIỂM)}
\begin{itemize}[leftmargin=*,itemsep=1pt]
\item \textbf{Câu 1 (1,0 điểm):} a) Ống 1 không tráng bạc; Ống 2 xuất hiện lớp bạc sáng bóng bám vào thành ống. b) $\ce{C12H22O11 + H2O ->[H^+,\, t^\circ] C6H12O6 (glucose) + C6H12O6 (fructose)}$.
\item \textbf{Câu 2 (1,0 điểm):} a) $\ce{(C17H35COO)3C3H5 + 3NaOH -> 3C17H35COONa + C3H5(OH)3}$. b) $m_{\text{glycerol}} = \mathbf{1{,}656\text{ kg}}$. c) Số bánh xà phòng tối đa: \textbf{\color{hfRed}209 bánh}.
\item \textbf{Câu 3 (1,0 điểm):} a) $m_{\text{ethanol (LT)}} = \mathbf{368\text{ kg}}$. b) Thể tích cồn y tế $70^\circ$ tối đa: \textbf{\color{hfRed}525,71 lít}.
\end{itemize}

\vspace{4pt}
\hrule color{hfGreen} height 0.5pt
""")

    # ==========================================
    # PHẦN I
    # ==========================================
    tex.append(r"""
\vspace{6pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 điểm)}}

\textit{Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\vspace{2pt}
""")

    # Câu 1
    tex.append(r"""
\begin{ex}
Chất nào sau đây là ester?
\choice
{$\ce{CH3COOH}$}
{\True $\ce{CH3COOC2H5}$}
{$\ce{C2H5OH}$}
{$\ce{CH3CHO}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$\ce{CH3COOC2H5}$ (ethyl acetate) là hợp chất thuộc loại ester. $\ce{CH3COOH}$ là carboxylic acid, $\ce{C2H5OH}$ là alcohol, $\ce{CH3CHO}$ là aldehyde. Chọn \textbf{B}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 2
    tex.append(r"""
\begin{ex}
Thủy phân hoàn toàn ethyl acetate trong dung dịch $\ce{NaOH}$ đun nóng, thu được sản phẩm gồm
\choice
{$\ce{CH3COONa}$ và $\ce{CH3OH}$}
{\True $\ce{CH3COONa}$ và $\ce{C2H5OH}$}
{$\ce{C2H5COONa}$ và $\ce{CH3OH}$}
{$\ce{HCOONa}$ và $\ce{C2H5OH}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phương trình xà phòng hóa: $\ce{CH3COOC2H5 + NaOH ->[t^\circ] CH3COONa + C2H5OH}$. Sản phẩm gồm sodium acetate và ethanol. Chọn \textbf{B}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 3
    tex.append(r"""
\begin{ex}
Phát biểu nào sau đây \textbf{không đúng} về xà phòng và chất giặt rửa tổng hợp?
\choice
{Xà phòng thường là muối sodium hoặc potassium của các acid béo}
{Khi giặt rửa trong nước cứng, xà phòng bị giảm tác dụng tẩy rửa do tạo muối kết tủa}
{\True Phân tử chất giặt rửa tổng hợp không có đầu ưa nước và đuôi kị nước như xà phòng}
{Chất giặt rửa tổng hợp có ưu điểm là có thể sử dụng được trong nước cứng và môi trường acid nhẹ}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phát biểu C sai vì phân tử chất giặt rửa tổng hợp cũng gồm 2 phần: đầu ưa nước (thường là nhóm $\ce{-SO3Na}$ hoặc $\ce{-OSO3Na}$) và đuôi kị nước (gốc hydrocarbon dài) tương tự xà phòng. Chọn \textbf{C}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 4
    tex.append(r"""
\begin{ex}
Carbohydrate nào sau đây thuộc loại monosaccharide?
\choice
{\True Glucose}
{Saccharose}
{Maltose}
{Tinh bột}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Glucose ($\ce{C6H12O6}$) là monosaccharide đơn giản nhất. Saccharose và maltose là disaccharide; tinh bột là polysaccharide. Chọn \textbf{A}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 5
    tex.append(r"""
\begin{ex}
Trong điều kiện thường, chất nào sau đây là chất rắn màu trắng, dạng sợi, không mùi, không tan trong nước và là thành phần chính tạo nên màng tế bào thực vật?
\choice
{Glucose}
{Saccharose}
{\True Cellulose}
{Amylose}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Cellulose là polymer thiên nhiên mạch không phân nhánh, dạng sợi màu trắng, không tan trong nước, tạo nên khung màng tế bào thực vật. Chọn \textbf{C}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 6
    tex.append(r"""
\begin{ex}
Thuốc thử dùng để nhận biết hồ tinh bột bằng phản ứng tạo dung dịch màu xanh tím đặc trưng là
\choice
{\True dung dịch iodine ($\ce{I2}$)}
{thuốc thử Tollens}
{dung dịch $\ce{Cu(OH)2}$}
{nước bromine}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Cấu trúc xoắn của tinh bột hấp phụ các phân tử iodine ($\ce{I2}$) tạo hợp chất bọc màu xanh tím đặc trưng. Chọn \textbf{A}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 7
    tex.append(r"""
\begin{ex}
Trong các chất sau, chất nào có lực base \textbf{yếu nhất}?
\choice
{Methylamine}
{Dimethylamine}
{Ammonia}
{\True Aniline}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Lực base giảm dần: $\ce{(CH3)2NH} > \ce{CH3NH2} > \ce{NH3} > \ce{C6H5NH2}$ (aniline). Gốc phenyl hút electron làm giảm mật độ electron trên nitrogen nên lực base của aniline rất yếu, không làm đổi màu quỳ tím. Chọn \textbf{D}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 8
    tex.append(r"""
\begin{ex}
Amino acid đơn giản nhất có tên gọi là glycine. Công thức phân tử của glycine là
\choice
{\True $\ce{C2H5NO2}$}
{$\ce{C3H7NO2}$}
{$\ce{C4H9NO2}$}
{$\ce{C5H9NO4}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Glycine có CTCT $\ce{H2N-CH2-COOH}$ nên CTPT là $\ce{C2H5NO2}$ ($M = 75\text{ g/mol}$). Chọn \textbf{A}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 9
    tex.append(r"""
\begin{ex}
Số liên kết peptide trong một phân tử tripeptide Gly-Ala-Val là
\choice
{1}
{\True 2}
{3}
{4}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Tripeptide tạo bởi 3 gốc $\alpha$-amino acid liên kết nhau bằng 2 liên kết peptide ($\ce{-CO-NH-}$). Chọn \textbf{B}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 10
    tex.append(r"""
\begin{ex}
Hiện tượng protein bị đông tụ và tách ra khỏi dung dịch xảy ra khi
\choice
{\True đun nóng dung dịch lòng trắng trứng}
{hòa tan lòng trắng trứng vào nước lạnh}
{thêm nước cất vào dung dịch lòng trắng trứng}
{khuấy đều dung dịch lòng trắng trứng ở nhiệt độ thường}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Khi đun nóng dung dịch protein (lòng trắng trứng), liên kết hydro bị đứt gãy làm mất cấu trúc không gian ba chiều tự nhiên dẫn đến hiện tượng đông tụ. Chọn \textbf{A}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 11
    tex.append(r"""
\begin{ex}
Thuốc thử dùng để phân biệt tripeptide với dipeptide bằng phản ứng màu biuret là
\choice
{dung dịch $\ce{AgNO3/NH3}$}
{\True $\ce{Cu(OH)2}$ trong môi trường kiềm}
{nước bromine}
{dung dịch $\ce{I2}$ trong $\ce{KI}$}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Trong môi trường kiềm, các hợp chất có từ 2 liên kết peptide trở lên (tripeptide trở lên) phản ứng với $\ce{Cu(OH)2}$ tạo phức màu tím biuret. Dipeptide không có phản ứng này. Chọn \textbf{B}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 12
    tex.append(r"""
\begin{ex}
Phát biểu nào sau đây về enzyme là \textbf{không đúng}?
\choice
{Hầu hết các enzyme có bản chất là protein}
{Enzyme là chất xúc tác sinh học có tính đặc hiệu rất cao}
{\True Enzyme xúc tác cho các phản ứng sinh hóa ở nhiệt độ rất cao (trên $150^\circ\text{C}$) để tăng tốc độ phản ứng}
{Tốc độ của phản ứng do enzyme xúc tác thường lớn hơn rất nhiều so với chất xúc tác hóa học thông thường}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phát biểu C không đúng vì enzyme hoạt động tối ưu ở nhiệt độ cơ thể (khoảng $37^\circ\text{C}$). Ở nhiệt độ cao (trên $60 - 70^\circ\text{C}$), protein enzyme bị biến tính và mất hoàn toàn hoạt tính xúc tác. Chọn \textbf{C}.
}
""")
    tex.append(r"\end{ex}")

    # ==========================================
    # PHẦN II
    # ==========================================
    tex.append(r"""
\vspace{6pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI (2,0 điểm)}}

\textit{Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\vspace{2pt}
""")

    # Câu 1 (Phần II)
    tex.append(r"""
\begin{ex}
Glucose và cellulose là hai carbohydrate phổ biến trong tự nhiên, đóng vai trò quan trọng trong sự sống và sản xuất công nghiệp. Cho các phát biểu sau:
\choiceTF
{\True Glucose có công thức phân tử là $\ce{C6H12O6}$, ở dạng mạch hở phân tử có chứa 5 nhóm hydroxy ($\ce{-OH}$) và 1 nhóm aldehyde ($\ce{-CHO}$).}
{Phân tử cellulose được cấu tạo từ các đơn vị $\alpha$-glucose liên kết với nhau qua liên kết $\alpha$-1,4-glycoside tạo thành mạch phân nhánh.}
{\True Thủy phân hoàn toàn cellulose trong môi trường acid vô cơ đun nóng thu được sản phẩm duy nhất là glucose, dung dịch sản phẩm có phản ứng tráng bạc với thuốc thử Tollens.}
{\True Đun nóng $32{,}4\text{ kg}$ cellulose với lượng dư hỗn hợp $\ce{HNO3}$ đặc trong $\ce{H2SO4}$ đặc, thu được cellulose trinitrate $[\ce{C6H7O2(ONO2)3}]_n$ dùng làm thuốc súng không khói. Biết hiệu suất phản ứng đạt $80\%$, khối lượng cellulose trinitrate thu được là $47{,}52\text{ kg}$.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) \textbf{Đúng.} Glucose mạch hở là $\ce{HOCH2-[CHOH]4-CHO}$ gồm 5 nhóm $\ce{-OH}$ kề nhau và 1 nhóm $\ce{-CHO}$.\\
b) \textbf{Sai.} Cellulose gồm các mắt xích $\beta$-glucose liên kết qua liên kết $\beta$-1,4-glycoside tạo mạch kéo dài không phân nhánh.\\
c) \textbf{Đúng.} $\ce{(C6H10O5)_n + nH2O ->[H^+,\, t^\circ] nC6H12O6}$ (glucose). Glucose tham gia tráng bạc sinh ra $\ce{Ag}$.\\
d) \textbf{Đúng.} $n_{\text{mắt xích}} = \dfrac{32{,}4}{162} = 0{,}2\text{ kmol}$. Khối lượng cellulose trinitrate lý thuyết: $0{,}2 \times 297 = 59{,}4\text{ kg}$. Với $H = 80\% \Rightarrow m_{\text{TT}} = 59{,}4 \times 80\% = \mathbf{47{,}52\text{ kg}}$.
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 (Phần II)
    tex.append(r"""
\begin{ex}
Cho các hợp chất hữu cơ chứa nitrogen gồm: methylamine ($\ce{CH3NH2}$), aniline ($\ce{C6H5NH2}$), glycine ($\ce{H2N-CH2-COOH}$), glutamic acid và tripeptide Gly-Ala-Val. Cho các phát biểu sau:
\choiceTF
{\True Ở điều kiện thường, methylamine là chất khí có mùi khai khó chịu, độc, tan nhiều trong nước và làm quỳ tím ẩm hóa xanh.}
{\True Glycine là amino acid đơn giản nhất, ở trạng thái tinh thể tồn tại chủ yếu dưới dạng ion lưỡng cực ($\ce{^+H3N-CH2-COO^-}$).}
{Dung dịch glutamic acid làm đổi màu quỳ tím sang màu xanh tương tự như dung dịch lysine.}
{\True Thủy phân hoàn toàn $0{,}1\text{ mol}$ peptide Gly-Ala-Val trong dung dịch chứa $0{,}4\text{ mol}$ $\ce{NaOH}$ đun nóng. Sau khi phản ứng xảy ra hoàn toàn, cô cạn dung dịch thu được $38{,}7\text{ gam}$ chất rắn khan.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) \textbf{Đúng.} Methylamine là chất khí có tính base mạnh hơn $\ce{NH3}$, làm quỳ tím ẩm hóa xanh.\\
b) \textbf{Đúng.} Amino acid tồn tại chủ yếu dưới dạng ion lưỡng cực nên có nhiệt độ nóng chảy cao và tan tốt trong nước.\\
c) \textbf{Sai.} Glutamic acid có 2 nhóm $\ce{-COOH}$ và 1 nhóm $\ce{-NH2}$ nên làm quỳ tím hóa đỏ (hồng); còn lysine có 2 nhóm $\ce{-NH2}$ và 1 nhóm $\ce{-COOH}$ làm quỳ tím hóa xanh.\\
d) \textbf{Đúng.} $M_{\text{Gly-Ala-Val}} = 75 + 89 + 117 - 2 \times 18 = 245\text{ g/mol}$.\\
Phản ứng: $\text{Gly-Ala-Val} + 3\ce{NaOH} \to \text{Muối} + \ce{H2O}$.\\
$n_{\ce{NaOH}\text{ phản ứng}} = 3 \times 0{,}1 = 0{,}3\text{ mol} \Rightarrow n_{\ce{NaOH}\text{ dư}} = 0{,}4 - 0{,}3 = 0{,}1\text{ mol}$.\\
$n_{\ce{H2O}} = 0{,}1\text{ mol} \Rightarrow m_{\text{muối}} = 0{,}1 \times 245 + 0{,}3 \times 40 - 0{,}1 \times 18 = 34{,}7\text{ gam}$.\\
Chất rắn khan gồm muối ($34{,}7\text{ g}$) và $\ce{NaOH}$ dư ($0{,}1 \times 40 = 4{,}0\text{ g}$) $\Rightarrow m_{\text{rắn}} = 34{,}7 + 4{,}0 = \mathbf{38{,}7\text{ gam}}$.
}
""")
    tex.append(r"\end{ex}")

    # ==========================================
    # PHẦN III
    # ==========================================
    tex.append(r"""
\vspace{6pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN III. CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 điểm)}}

\textit{Thí sinh trả lời từ câu 1 đến câu 4. Mỗi câu trả lời đúng được 0,5 điểm.}

\vspace{2pt}
""")

    # Câu 1 (Phần III)
    tex.append(r"""
\begin{ex}
Cho các chất sau: sodium stearate, sodium dodecylbenzenesulfonate, saccharose, glycerol, potassium palmitate, ethyl acetate. Có bao nhiêu chất là thành phần chính của xà phòng hoặc chất giặt rửa tổng hợp?
""")
    if not is_teacher:
        tex.append(r"\par\noindent\textbf{Trả lời:} \dotfill" + "\n")
    else:
        tex.append(r"""
\loigiai{
Có \textbf{3} chất thỏa mãn:\\
- Xà phòng: sodium stearate ($\ce{C17H35COONa}$), potassium palmitate ($\ce{C15H31COOK}$).\\
- Chất giặt rửa tổng hợp: sodium dodecylbenzenesulfonate ($\ce{CH3[CH2]11C6H4SO3Na}$).\\
\textbf{Đáp số: 3.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 (Phần III)
    tex.append(r"""
\begin{ex}
Trong các amine sau: methylamine, dimethylamine, trimethylamine, ethylamine, aniline; có bao nhiêu amine ở thể khí ở điều kiện thường?
""")
    if not is_teacher:
        tex.append(r"\par\noindent\textbf{Trả lời:} \dotfill" + "\n")
    else:
        tex.append(r"""
\loigiai{
Có \textbf{4} amine ở thể khí ở điều kiện thường là: methylamine, dimethylamine, trimethylamine và ethylamine. (Aniline là chất lỏng không màu).\\
\textbf{Đáp số: 4.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 3 (Phần III)
    tex.append(r"""
\begin{ex}
Nhỏ từng giọt dung dịch của mỗi chất sau vào các mẩu giấy quỳ tím riêng biệt: glycine, ethylamine, aniline, lysine, glutamic acid. Có bao nhiêu chất làm quỳ tím chuyển sang màu xanh?
""")
    if not is_teacher:
        tex.append(r"\par\noindent\textbf{Trả lời:} \dotfill" + "\n")
    else:
        tex.append(r"""
\loigiai{
Có \textbf{2} chất làm quỳ tím chuyển sang màu xanh là: ethylamine (amine béo có tính base) và lysine (có 2 nhóm $\ce{-NH2}$ và 1 nhóm $\ce{-COOH}$).\\
(Glycine không làm đổi màu quỳ tím; aniline có tính base rất yếu không đổi màu quỳ tím; glutamic acid làm quỳ tím hóa đỏ).\\
\textbf{Đáp số: 2.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 4 (Phần III)
    tex.append(r"""
\begin{ex}
Cho $9{,}0\text{ gam}$ ethylamine ($\ce{C2H5NH2}$) tác dụng vừa đủ với dung dịch hydrochloric acid ($\ce{HCl}$). Sau khi phản ứng hoàn toàn, khối lượng muối ethylammonium chloride ($\ce{C2H5NH3Cl}$) thu được là bao nhiêu gam?
""")
    if not is_teacher:
        tex.append(r"\par\noindent\textbf{Trả lời:} \dotfill" + "\n")
    else:
        tex.append(r"""
\loigiai{
Phương trình phản ứng: $\ce{C2H5NH2 + HCl -> C2H5NH3Cl}$.\\
$n_{\ce{C2H5NH2}} = \dfrac{9{,}0}{45} = 0{,}2\text{ mol} \Rightarrow n_{\text{muối}} = 0{,}2\text{ mol}$.\\
Khối lượng muối thu được: $m = 0{,}2 \times (45 + 36{,}5) = 0{,}2 \times 81{,}5 = \mathbf{16{,}3\text{ gam}}$.\\
\textbf{Đáp số: 16,3.}
}
""")
    tex.append(r"\end{ex}")

    # ==========================================
    # PHẦN IV
    # ==========================================
    tex.append(r"""
\vspace{6pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN IV. TỰ LUẬN (3,0 điểm)}}

\textit{Thí sinh làm bài trực tiếp vào đề thi hoặc giấy làm bài thi. Điểm tối đa: 3,0 điểm.}

\vspace{2pt}
""")

    # Câu 1 (Phần IV)
    tex.append(r"""
\begin{ex}
\textbf{(1,0 điểm)} Có hai ống nghiệm chứa dung dịch saccharose không màu. Tiến hành hai thí nghiệm sau:
\begin{itemize}[leftmargin=1.5cm,itemsep=1pt]
\item \textit{Ống nghiệm 1:} Thêm lượng dư thuốc thử Tollens $[\ce{Ag(NH3)2}]\ce{OH}$, sau đó ngâm ống nghiệm trong cốc nước nóng (khoảng $70 - 80^\circ\text{C}$) trong 5 phút.
\item \textit{Ống nghiệm 2:} Thêm $1\text{ mL}$ dung dịch $\ce{H2SO4}$ $10\%$ rồi đun nóng nhẹ trong 5 phút. Để nguội, trung hòa acid dư bằng dung dịch $\ce{NaOH}$ loãng, sau đó thêm lượng dư thuốc thử Tollens và ngâm trong cốc nước nóng.
\end{itemize}
a) Hãy nêu hiện tượng quan sát được ở mỗi ống nghiệm.\\
b) Viết phương trình hóa học của phản ứng thủy phân xảy ra ở ống nghiệm 2 và phản ứng tráng bạc giải thích hiện tượng quan sát được.
""")
    if not is_teacher:
        tex.append(r"""
\vspace{0.2cm}
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill
""")
    else:
        tex.append(r"""
\loigiai{
a) \textbf{Hiện tượng quan sát được:}\\
- \textit{Ống nghiệm 1:} Không có hiện tượng tráng bạc, không xuất hiện lớp kim loại sáng bám vào thành ống (do phân tử saccharose gồm 1 gốc $\alpha$-glucose và 1 gốc $\beta$-fructose liên kết qua C1 và C2, không còn nhóm hemiacetal/hemiketal tự do nên không mở vòng tạo nhóm aldehyde $\ce{-CHO}$, không có tính khử).\\
- \textit{Ống nghiệm 2:} Thành ống nghiệm xuất hiện lớp kim loại bạc sáng bóng bám vào như gương (hoặc kết tủa bạc xám đen).\\
b) \textbf{Phương trình hóa học:}\\
- Phản ứng thủy phân saccharose trong môi trường acid:\\
\[ \ce{C12H22O11 + H2O ->[H^+,\, t^\circ] C6H12O6 (glucose) + C6H12O6 (fructose)} \]
- Phản ứng tráng bạc của glucose sinh ra:\\
\[ \ce{CH2OH[CHOH]4CHO + 2[Ag(NH3)2]OH ->[t^\circ] CH2OH[CHOH]4COONH4 + 2Ag v + 3NH3 + H2O} \]
\textit{(Lưu ý: Trong môi trường kiềm của thuốc thử Tollens, fructose cũng chuyển hóa thành glucose nên cả 2 monosaccharide đều tham gia phản ứng tráng bạc).}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 (Phần IV)
    tex.append(r"""
\begin{ex}
\textbf{(1,0 điểm)} Một cơ sở sản xuất xà phòng thủ công tiến hành xà phòng hóa hoàn toàn $17{,}8\text{ kg}$ một loại chất béo chứa $90\%$ khối lượng là tristearin $\ce{(C17H35COO)3C3H5}$ (còn lại là tạp chất trơ không phản ứng) bằng một lượng vừa đủ dung dịch $\ce{NaOH}$, đun nóng.\\
a) Viết phương trình hóa học của phản ứng xà phòng hóa tristearin bằng dung dịch $\ce{NaOH}$.\\
b) Tính khối lượng glycerol thu được sau phản ứng (giả sử hiệu suất phản ứng xà phòng hóa đạt $100\%$).\\
c) Toàn bộ lượng muối sodium stearate ($\ce{C17H35COONa}$) thu được sau khi cô đặc và phối trộn phụ gia (chiếm $25\%$ khối lượng bánh xà phòng thương phẩm) được ép thành các bánh xà phòng. Mỗi bánh xà phòng có khối lượng $100\text{ gam}$ và hiệu suất thu hồi xà phòng đạt $95\%$. Tính số bánh xà phòng tối đa mà cơ sở trên sản xuất được.
""")
    if not is_teacher:
        tex.append(r"""
\vspace{0.2cm}
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill
""")
    else:
        tex.append(r"""
\loigiai{
a) \textbf{Phương trình hóa học xà phòng hóa:}\\
\[ \ce{(C17H35COO)3C3H5 + 3NaOH ->[t^\circ] 3C17H35COONa + C3H5(OH)3} \]
b) \textbf{Tính khối lượng glycerol:}\\
- Khối lượng tristearin nguyên chất: $m = 17{,}8 \times 90\% = 16{,}02\text{ kg} = 16020\text{ gam}$.\\
- Số mol tristearin: $n = \dfrac{16020}{890} = 18\text{ mol}$.\\
- Theo PTHH: $n_{\text{glycerol}} = n_{\text{tristearin}} = 18\text{ mol}$.\\
- Khối lượng glycerol: $m_{\text{glycerol}} = 18 \times 92 = 1656\text{ gam} = \mathbf{1{,}656\text{ kg}}$.\\
c) \textbf{Tính số bánh xà phòng tối đa:}\\
- Số mol sodium stearate lý thuyết: $n_{\text{muối (LT)}} = 3 \times 18 = 54\text{ mol}$.\\
- Khối lượng sodium stearate lý thuyết: $m = 54 \times 306 = 16524\text{ gam}$.\\
- Do hiệu suất thu hồi đạt $95\% \Rightarrow m_{\text{muối (TT)}} = 16524 \times 95\% = 15697{,}8\text{ gam}$.\\
- Trong bánh xà phòng, sodium stearate chiếm: $100\% - 25\% = 75\%$ khối lượng.\\
- Tổng khối lượng xà phòng bánh thu được: $m_{\text{xà phòng}} = \dfrac{15697{,}8}{75\%} = 20930{,}4\text{ gam}$.\\
- Số bánh xà phòng ($100\text{ g/bánh}$): $N = \dfrac{20930{,}4}{100} = 209{,}304 \Rightarrow$ Tối đa sản xuất được \textbf{\color{hfRed}209 bánh}.
}
""")
    tex.append(r"\end{ex}")

    # Câu 3 (Phần IV)
    tex.append(r"""
\begin{ex}
\textbf{(1,0 điểm)} Để sản xuất cồn y tế $70^\circ$ (dung dịch ethanol $70\%$ theo thể tích) phục vụ sát trùng trong y tế, người ta sử dụng nguyên liệu ban đầu là 1 tấn hạt ngô chứa $64{,}8\%$ tinh bột (còn lại là chất xơ và tạp chất trơ không sinh cồn). Quá trình sản xuất trải qua hai giai đoạn theo sơ đồ sau:
\[ \ce{(C6H10O5)_n ->[+\ce{H2O},\, \text{enzyme/acid}] nC6H12O6 ->[\text{lên men}] 2nC2H5OH + 2nCO2} \]
a) Tính khối lượng ethanol lý thuyết thu được từ 1 tấn hạt ngô trên.\\
b) Biết hiệu suất của toàn bộ quá trình sản xuất trên đạt $80\%$ và khối lượng riêng của ethanol nguyên chất là $D = 0{,}8\text{ g/mL}$. Tính thể tích (lít) dung dịch cồn y tế $70^\circ$ tối đa pha chế được từ lượng ethanol thực tế thu được.
""")
    if not is_teacher:
        tex.append(r"""
\vspace{0.2cm}
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill\\
\dotfill
""")
    else:
        tex.append(r"""
\loigiai{
a) \textbf{Tính khối lượng ethanol lý thuyết:}\\
- Khối lượng tinh bột trong 1 tấn ($1000\text{ kg}$) ngô: $m_{\text{tinh bột}} = 1000 \times 64{,}8\% = 648\text{ kg}$.\\
- Sơ đồ chuyển hóa:
\[ \begin{matrix}
\ce{(C6H10O5)_n} & \longrightarrow & \ce{2nC2H5OH} \\
162n\text{ kg} & & 2n \times 46 = 92n\text{ kg} \\
648\text{ kg} & & m_{\text{ethanol (LT)}}
\end{matrix} \]
$\Rightarrow m_{\text{ethanol (LT)}} = \dfrac{648 \times 92n}{162n} = \mathbf{368\text{ kg}}$.\\
b) \textbf{Tính thể tích dung dịch cồn $70^\circ$:}\\
- Khối lượng ethanol thực tế thu được ($H = 80\%$):
\[ m_{\text{ethanol (TT)}} = 368 \times 80\% = 294{,}4\text{ kg} = 294400\text{ gam}. \]
- Thể tích ethanol nguyên chất:
\[ V_{\text{ethanol}} = \dfrac{m}{D} = \dfrac{294400}{0{,}8} = 368000\text{ mL} = 368\text{ lít}. \]
- Thể tích dung dịch cồn y tế $70^\circ$ pha chế được:
\[ V_{\text{cồn } 70^\circ} = \dfrac{V_{\text{ethanol}}}{70\%} = \dfrac{368}{0{,}7} \approx \mathbf{525{,}71\text{ lít}}. \]
}
""")
    tex.append(r"\end{ex}")

    tex.append(r"""
\vspace{8pt}
\begin{center}
\textbf{---------- HẾT ----------}
\end{center}

\end{document}
""")

    return "".join(tex)


def compile_tex(tex_path):
    out_dir = os.path.dirname(tex_path)
    base_name = os.path.splitext(os.path.basename(tex_path))[0]
    print(f"Biên dịch LaTeX: {base_name}...")
    for run_i in range(2):
        cmd = ["pdflatex", "-interaction=nonstopmode", f"-output-directory={out_dir}", tex_path]
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, cwd=out_dir)
        if proc.returncode != 0 and run_i == 1:
            print(f"[Cảnh báo] pdflatex returncode {proc.returncode} cho {base_name}")

    # Dọn dẹp file rác
    for ext in [".aux", ".log", ".thm", ".out"]:
        junk = os.path.join(out_dir, base_name + ext)
        if os.path.exists(junk):
            try:
                os.remove(junk)
            except Exception:
                pass
    pdf_path = os.path.join(out_dir, base_name + ".pdf")
    if os.path.exists(pdf_path):
        print(f"-> Xuất thành công PDF: {pdf_path}")
    else:
        print(f"-> [Lỗi] Không tìm thấy file PDF: {pdf_path}")


if __name__ == "__main__":
    hs_tex_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_THPT_NguyenHue_Ma101_HocSinh.tex")
    gv_tex_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_THPT_NguyenHue_Ma101_GiaoVien.tex")

    print("Ghi file LaTeX Bản Học sinh...")
    with open(hs_tex_path, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))

    print("Ghi file LaTeX Bản Giáo viên...")
    with open(gv_tex_path, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))

    print("\nTiến hành biên dịch PDF bằng pdflatex...")
    compile_tex(hs_tex_path)
    compile_tex(gv_tex_path)
    print("\nHoàn tất biên dịch LaTeX và PDF!")
