# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Giữa Kỳ I Hóa Học 12 - THPT Nguyễn Trường Tộ - Mã Đề 103
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma103_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma103_GiaoVien.pdf
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
""")

    # Tiêu đề
    tex.append(r"""
\noindent
\begin{tabularx}{\textwidth}{@{}X r@{}}
\textbf{TRƯỜNG THPT NGUYỄN TRƯỜNG TỘ} & \textbf{ĐỀ KIỂM TRA GIỮA KỲ I} \\
\textbf{TỔ HÓA - TIN} & \textbf{NĂM HỌC 2026 - 2027} \\
\textbf{ĐỀ CHÍNH THỨC} & \textbf{Môn: Hóa học lớp 12} \\
\textit{(Đề có 04 trang)} & \textbf{Thời gian làm bài: 45 phút} \\
& \textbf{Mã đề: 103}
\end{tabularx}
""")

    if is_teacher:
        tex.append(r"""
\begin{center}
{\color{hfRed}\textbf{\large (BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}}
\end{center}
""")

    tex.append(r"""
\vspace{2pt}
\noindent\textbf{Họ và tên:} \dotfill \textbf{Số báo danh:} \dots\dots\dots\dots

\vspace{2pt}
\noindent\textit{Cho biết nguyên tử khối của các nguyên tố: H = 1; C = 12; N = 14; O = 16; Na = 23; Br = 80; Cu = 64; Ag = 108.}
\vspace{2pt}
\hrule height 0.5pt
\vspace{6pt}
""")

    if is_teacher:
        tex.append(r"""
\begin{center}
{\color{hfRed}\textbf{\large BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 103)}}
\end{center}

\noindent{\color{mainGreen}\textbf{1. Phần I: Trắc nghiệm nhiều lựa chọn (16 câu)}}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|*{8}{>{\centering\arraybackslash}m{0.95cm}|}}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}C} \\
\hline
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} \\
\hline
\end{tabular}
\end{center}

\vspace{3pt}
\noindent{\color{mainGreen}\textbf{2. Phần II: Trắc nghiệm Đúng / Sai}}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Lệnh hỏi} & \textbf{Ý a)} & \textbf{Ý b)} & \textbf{Ý c)} & \textbf{Ý d)} \\
\hline
\textbf{Câu 1} & Đúng / Sai & \textbf{\color{hfRed}Đ} & \textbf{S} & \textbf{\color{hfRed}Đ} & \textbf{S} \\
\hline
\textbf{Câu 2} & Đúng / Sai & \textbf{\color{hfRed}Đ} & \textbf{S} & \textbf{\color{hfRed}Đ} & \textbf{S} \\
\hline
\end{tabular}
\end{center}

\vspace{3pt}
\noindent{\color{mainGreen}\textbf{3. Phần III: Trắc nghiệm Trả lời ngắn}}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}5} & \textbf{\color{hfRed}4} & \textbf{\color{hfRed}4402} & \textbf{\color{hfRed}8} \\
\hline
\end{tabular}
\end{center}
\vspace{6pt}
\hrule height 0.5pt
\vspace{6pt}
""")

    # PHẦN A
    tex.append(r"""
\noindent{\color{mainGreen}\textbf{\Large TRẮC NGHIỆM KHÁCH QUAN (7,0 ĐIỂM)}}
\vspace{4pt}

\noindent{\color{mainGreen}\textbf{\large PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.}} \\
\textit{Thí sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{4pt}
""")

    p1_items = [
        (r"Dipeptide có công thức: \ce{H2NCH2CONHCH(CH3)COOH}. Tên viết tắt của X là",
         ["Ala-Val.", "Gly-Ala.", "Ala-Gly.", "Gly-Val."],
         1,
         r"Dipeptide gồm gốc glycine (\ce{H2NCH2CO-}) và alanine (\ce{-NHCH(CH3)COOH}) nên có tên viết tắt là Gly-Ala."),

        ("Chất nào sau đây là thành phần chủ yếu của xà phòng?",
         [r"\ce{CH3COONa}.", r"\ce{C15H31COOK}.", r"\ce{CH2=CHCOONa}.", r"\ce{CH3(CH2)3COOK}."],
         1,
         r"Xà phòng là muối của acid béo mạch dài: \ce{C15H31COOK} (potassium palmitate)."),

        ("Công thức cấu tạo của ethylamine là",
         [r"\ce{CH3NH2}.", r"\ce{CH3NHCH3}.", r"\ce{C2H5NH2}.", r"\ce{(CH3)3N}."],
         2,
         r"Ethylamine có công thức \ce{C2H5NH2}."),

        ("Cơ thể người mã hoá được loại amino acid nào sau đây trong tổng hợp protein cho cơ thể?",
         [r"\ce{H2N-CH2-CH2-OH}.", r"\ce{^+NH3-CH2-CH2-COO^-}.", r"\ce{^+NH3-CH2-COO^-}.", r"\ce{^+NH3-(CH2)3-COO^-}."],
         2,
         r"Protein được tạo nên từ các $\alpha$-amino acid. \ce{^+NH3-CH2-COO^-} là dạng ion lưỡng cực của glycine ($\alpha$-amino acid)."),

        (r"Carbohydrate nào dưới đây không có nhóm -OH hemiacetal hoặc nhóm -OH hemiketal?",
         ["Maltose.", "Fructose.", "Glucose.", "Saccharose."],
         3,
         "Trong phân tử saccharose, cả hai nhóm -OH hemiacetal và hemiketal đều đã tham gia liên kết glycoside."),

        (r"Giá trị pH mà khi đó amino acid có nồng độ ion lưỡng cực là cực đại được gọi là điểm đẳng điện (pI). Cho các giá trị pI: glycine ($6{,}0$); glutamic acid ($3{,}2$); lysine ($9{,}7$). Giá trị pH tối ưu nhất để tách ba chất trên ra khỏi dung dịch hỗn hợp là",
         [r"$\text{pH} = 3{,}2$.", r"$\text{pH} = 9{,}7$.", r"$\text{pH} = 6{,}0$.", r"$\text{pH} = 14{,}0$."],
         2,
         r"Tại $\text{pH} = 6{,}0$: glycine đứng yên; glutamic acid ($\text{pI} = 3{,}2$) về cực $(+)$; lysine ($\text{pI} = 9{,}7$) về cực $(-)$."),

        ("Polymer thiên nhiên X được sinh ra trong quá trình quang hợp của cây xanh. Ở nhiệt độ thường, X tạo với dung dịch iodine hợp chất có màu xanh tím. Polymer X là",
         ["tinh bột.", "cellulose.", "saccharose.", "glycogen."],
         0,
         "Tinh bột tạo màu xanh tím đặc trưng với dung dịch iodine."),

        ("Cho các chất: glycerol, glucose, maltose, acetic acid, fructose và cellulose. Có bao nhiêu chất thuộc loại carbohydrate?",
         ["2.", "3.", "4.", "5."],
         2,
         "Có 4 chất là carbohydrate: glucose, maltose, fructose, cellulose."),

        (r"Một ester có công thức phân tử là \ce{C3H6O2}, có tham gia phản ứng tráng bạc. Công thức cấu tạo thu gọn của ester đó là",
         [r"\ce{HCOOC3H7}.", r"\ce{CH3COOCH3}.", r"\ce{C2H5COOCH3}.", r"\ce{HCOOC2H5}."],
         3,
         r"Ester tráng bạc có dạng \ce{HCOOR'}. Với \ce{C3H6O2} là ethyl formate: \ce{HCOOC2H5}."),

        (r"Để chứng minh trong phân tử của glucose có nhiều nhóm hydroxyl kề nhau, người ta cho dung dịch glucose phản ứng với",
         [r"\ce{Cu(OH)2} ở nhiệt độ thường.", "kim loại Na.", r"\ce{AgNO3} trong dung dịch \ce{NH3}, đun nóng.", r"\ce{Cu(OH)2} trong \ce{NaOH}, đun nóng."],
         0,
         r"\ce{Cu(OH)2} hòa tan ở nhiệt độ thường tạo phức màu xanh lam đặc trưng cho các chất có nhiều nhóm -OH kề nhau."),

        (r"Polysaccharide X là chất rắn, màu trắng, dạng sợi. Trong bông nõn có gần 98\% chất X. Thủy phân X, thu được monosaccharide Y. Phát biểu nào sau đây là đúng?",
         ["X có tính chất của aldehyde.", "X dễ tan trong nước và có vị ngọt.", "Y là nguồn lương thực chính của con người và một số động vật.", "X được sử dụng làm vật liệu xây dựng, sản xuất giấy và sợi nhân tạo."],
         3,
         "X là cellulose, được dùng làm vật liệu xây dựng, sản xuất giấy và tơ nhân tạo."),

        ("Glucose là một loại monosaccharide có nhiều trong quả nho chín. Công thức phân tử của glucose là",
         [r"\ce{C12H22O11}.", r"\ce{C2H4O2}.", r"\ce{C6H12O6}.", r"\ce{(C6H10O5)_n}."],
         2,
         r"Glucose có công thức phân tử là \ce{C6H12O6}."),

        ("Methyl acetate có công thức cấu tạo là",
         [r"\ce{HCOOCH3}.", r"\ce{HCOOC2H5}.", r"\ce{CH3COOCH3}.", r"\ce{CH3COOC2H5}."],
         2,
         r"Methyl acetate có công thức cấu tạo là \ce{CH3COOCH3}."),

        ("Mùi tanh của cá là hỗn hợp các amine và một số tạp chất khác. Để khử mùi tanh của cá trước khi chế biến thực phẩm, nên áp dụng cách nào sau đây?",
         ["Ngâm cá trong nước để amine tan vào nước.", "Rửa cá bằng giấm ăn.", "Rửa cá bằng dung dịch nước muối.", "Rửa cá bằng dung dịch soda."],
         1,
         r"Giấm ăn chứa acid acetic trung hòa amine tạo muối tan không mùi."),

        ("Cho các chất sau: \ce{CH3COOCH3}, \ce{HCOOCH3}, \ce{HCOOC6H5}, \ce{CH3COOC2H5}. Chất có nhiệt độ sôi thấp nhất là",
         [r"\ce{HCOOC6H5}.", r"\ce{CH3COOC2H5}.", r"\ce{HCOOCH3}.", r"\ce{CH3COOCH3}."],
         2,
         r"\ce{HCOOCH3} có phân tử khối nhỏ nhất ($M = 60$) nên có nhiệt độ sôi thấp nhất ($\approx 31{,}5^\circ\text{C}$)."),

        ("Khi nấu canh cua thấy các mảng ``gạch cua'' nổi lên là do",
         ["sự đông tụ lipid.", "phản ứng thủy phân protein.", "sự đông tụ protein.", "phản ứng màu biuret."],
         2,
         "Protein hòa tan bị nhiệt độ làm biến tính và đông tụ nổi lên mặt nước.")
    ]

    for idx, (q_text, opts, correct_idx, sol) in enumerate(p1_items, 1):
        tex.append(f"\\begin{{ex}}\n{q_text}\n\\choice\n")
        opts_formatted = []
        for o_idx, opt in enumerate(opts):
            if is_teacher and o_idx == correct_idx:
                opts_formatted.append(f"\\textbf{{\\color{{hfRed}}{opt}}}")
            else:
                opts_formatted.append(opt)
        tex.append(f"{{{opts_formatted[0]}}}\n{{{opts_formatted[1]}}}\n{{{opts_formatted[2]}}}\n{{{opts_formatted[3]}}}\n")
        if is_teacher:
            labels = ["A", "B", "C", "D"]
            tex.append(f"\\loigiai{{\\textbf{{Chọn {labels[correct_idx]}.}} {sol}}}\n")
        else:
            tex.append("\\loigiai{}\n")
        tex.append("\\end{ex}\n")

    # PHẦN II
    tex.append(r"""
\vspace{6pt}
\noindent{\color{mainGreen}\textbf{\large PHẦN II. Câu trắc nghiệm đúng sai.}} \\
\textit{Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{4pt}
""")

    # CÂU 1 PHẦN II
    tex.append(r"""
\begin{ex}
Ethyl acetate là một chất lỏng không màu, dễ bay hơi, ít độc hại, không hút ẩm, có mùi ngọt và dễ chịu, thường được sử dụng trong việc sản xuất chất tẩy sơn móng tay, mực in, keo dán, thuốc lá, nước hoa.
\choiceTF
""")
    tf1 = [
        (r"Đun nóng hỗn hợp gồm với 9,0 gam acetic acid và 9,2 gam ethyl alcohol có \ce{H2SO4} đặc làm chất xúc tác, thu được 10,56 gam este. Biết hiệu suất phản ứng đạt 80\%.", True),
        ("Ethyl acetate có khả năng tham gia phản ứng tráng bạc.", False),
        ("Ethyl acetate thuộc loại ester no, mạch hở, đơn chức.", True),
        (r"Công thức cấu tạo của ethyl acetate là \ce{CH3CH2COOCH3}.", False)
    ]
    for stmt, val in tf1:
        if val:
            tex.append(f"\\True {stmt}\n")
        else:
            tex.append(f"{stmt}\n")
    if is_teacher:
        tex.append(r"""\loigiai{
\textbf{Ý a) ĐÚNG.} $n_{\ce{CH3COOH}} = 0{,}15\text{ mol}; n_{\ce{C2H5OH}} = 0{,}2\text{ mol} \Rightarrow m_{\text{lt}} = 0{,}15 \times 88 = 13{,}2\text{ g} \Rightarrow H = \frac{10{,}56}{13{,}2} \times 100\% = 80\%$.\\
\textbf{Ý b) SAI.} Ethyl acetate không chứa nhóm aldehyde nên không tráng bạc.\\
\textbf{Ý c) ĐÚNG.} Công thức \ce{C4H8O2} là ester no, đơn chức, mạch hở.\\
\textbf{Ý d) SAI.} Ethyl acetate là \ce{CH3COOCH2CH3}. \ce{CH3CH2COOCH3} là methyl propionate.
}
""")
    else:
        tex.append("\\loigiai{}\n")
    tex.append("\\end{ex}\n")

    # CÂU 2 PHẦN II
    tex.append(r"""
\begin{ex}
Trong công nghiệp sản xuất bia có các bước chính sau:\\
\centerline{Hạt đại mạch $\xrightarrow[\text{enzyme}]{\ce{H2O}, t^o}$ Dịch nấu (maltose, glucose) $\xrightarrow{\text{enzyme}}$ Ethanol}
\choiceTF
""")
    tf2 = [
        (r"Để biết được thời điểm kết thúc quá trình thủy phân tinh bột ta có thể kiểm tra bằng thuốc thử là dung dịch \ce{I2} trong \ce{KI}.", True),
        ("Maltose và glucose thuộc nhóm monosaccharide, khi lên men thu được ethanol.", False),
        ("Thành phần trong hạt đại mạch bị thủy phân tạo ra maltose, glucose là tinh bột.", True),
        ("Sản phẩm thủy phân của tinh bột (hạt đại mạch), ngoài maltose và glucose còn thu được sản phẩm phụ là fructose.", False)
    ]
    for stmt, val in tf2:
        if val:
            tex.append(f"\\True {stmt}\n")
        else:
            tex.append(f"{stmt}\n")
    if is_teacher:
        tex.append(r"""\loigiai{
\textbf{Ý a) ĐÚNG.} Khi tinh bột bị thủy phân hết, dung dịch không còn tạo màu xanh tím với dung dịch \ce{I2}.\\
\textbf{Ý b) SAI.} Maltose là disaccharide.\\
\textbf{Ý c) ĐÚNG.} Tinh bột trong mầm đại mạch bị enzyme thủy phân.\\
\textbf{Ý d) SAI.} Tinh bột chỉ gồm các mắt xích glucose, không tạo ra fructose.
}
""")
    else:
        tex.append("\\loigiai{}\n")
    tex.append("\\end{ex}\n")

    # PHẦN III
    tex.append(r"""
\vspace{6pt}
\noindent{\color{mainGreen}\textbf{\large PHẦN III: Câu trắc nghiệm yêu cầu trả lời ngắn.}} \\
\textit{Thí sinh trả lời từ câu 1 đến câu 4.}
\vspace{4pt}
""")

    # CÂU 1 PHẦN III
    tex.append(r"""
\begin{ex}
Cho các chất sau: ethyl acetate, tripalmitin, saccharose, ethylamine, Gly-Ala, fibroin, Ala-Ala-Val. Số chất tham gia phản ứng thủy phân trong môi trường kiềm là bao nhiêu?
\shortans{5}
\loigiai{
\textbf{Trả lời: 5}\\
Gồm 5 chất: ethyl acetate, tripalmitin, Gly-Ala, fibroin, Ala-Ala-Val.
}
\end{ex}
""")

    # CÂU 2 PHẦN III
    tex.append(r"""
\begin{ex}
Có bao nhiêu đồng phân cấu tạo amine bậc 1 ứng với công thức phân tử \ce{C4H11N}?
\shortans{4}
\loigiai{
\textbf{Trả lời: 4}\\
Gốc butyl \ce{C4H9-} có 4 đồng phân cấu tạo gắn nhóm \ce{-NH2}:
\begin{enumerate}[noitemsep,topsep=1pt]
\item \ce{CH3-CH2-CH2-CH2-NH2} (butan-1-amine)
\item \ce{CH3-CH2-CH(NH2)-CH3} (butan-2-amine)
\item \ce{(CH3)2CH-CH2-NH2} (2-methylpropan-1-amine)
\item \ce{(CH3)3C-NH2} (2-methylpropan-2-amine)
\end{enumerate}
}
\end{ex}
""")

    # CÂU 3 PHẦN III
    tex.append(r"""
\begin{ex}
Để sản xuất ethanol, người ta dùng nguyên liệu là mùn cưa và vỏ bao từ gỗ chứa 50\% cellulose. Nếu muốn điều chế 1 tấn ethanol, hiệu suất cả quá trình là 80\% thì khối lượng nguyên liệu cần dùng là bao nhiêu kg? \textit{(Kết quả làm tròn đến hàng đơn vị)}.
\shortans{4402}
\loigiai{
\textbf{Trả lời: 4402}\\
$m_{\text{cellulose, lt}} = 1000 \times \frac{162}{92} = 1760{,}87\text{ kg}$.\\
Với $H = 80\% \Rightarrow m_{\text{cellulose, tt}} = \frac{1760{,}87}{80\%} = 2201{,}087\text{ kg}$.\\
Khối lượng nguyên liệu: $m = \frac{2201{,}087}{50\%} \approx \mathbf{4402\text{ kg}}$.
}
\end{ex}
""")

    # CÂU 4 PHẦN III
    tex.append(r"""
\begin{ex}
Saccharose có công thức phân tử \ce{C12H22O11}, cấu tạo từ một đơn vị $\alpha$-glucose và một đơn vị $\beta$-fructose qua liên kết $\alpha$-1,2-glycoside. Tổng số nhóm -OH trong phân tử saccharose là bao nhiêu?
\shortans{8}
\loigiai{
\textbf{Trả lời: 8}\\
Gốc glucose còn 4 nhóm -OH, gốc fructose còn 4 nhóm -OH. Tổng cộng có: $4 + 4 = \mathbf{8}$ nhóm -OH.
}
\end{ex}
""")

    # PHẦN B
    tex.append(r"""
\vspace{6pt}
\noindent{\color{mainGreen}\textbf{\Large B. PHẦN TỰ LUẬN (3,0 ĐIỂM)}}
\vspace{4pt}

\noindent\textbf{Câu 1.} Một loại chất béo có chứa 80\% tristearin về khối lượng. Để sản xuất 9,2 triệu chai nước rửa tay (có chứa chất dưỡng ẩm glycerol) cần dùng tối thiểu x tấn loại chất béo trên cho phản ứng với dung dịch \ce{NaOH}, đun nóng. Biết rằng trong mỗi chai nước rửa tay có chứa 6 gam glycerol. Tính x?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
- Khối lượng glycerol: $m = 9\,200\,000 \times 6\text{ g} = 55\,200\,000\text{ g} = 55{,}2\text{ tấn} \Rightarrow n = 600\text{ kmol}$.\\
- Phương trình: $\ce{(C17H35COO)3C3H5 + 3NaOH -> 3C17H35COONa + C3H5(OH)3}$.\\
- $n_{\text{tristearin}} = n_{\text{glycerol}} = 600\text{ kmol} \Rightarrow m = 600 \times 890 = 534\,000\text{ kg} = 534\text{ tấn}$.\\
- Khối lượng chất béo chứa 80\% tristearin: $x = \frac{534}{80\%} = \mathbf{667{,}5\text{ tấn}}$.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{6pt}
\multido{\i=1+1}{6}{\noindent\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\noindent\textbf{Câu 2.} Glucose tồn tại đồng thời dạng mạch hở và mạch vòng ($\alpha$ và $\beta$). Glucose có tính chất của aldehyde và của polyalcohol. Cho các chất sau: \ce{Cu(OH)2}, nước bromine, dung dịch \ce{I2} trong \ce{KI}, thuốc thử Tollens và \ce{CH3OH}/\ce{HCl} khan. Ở điều kiện thích hợp, liệt kê các chất tác dụng được với glucose?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
Các chất tác dụng được với glucose gồm:
\begin{itemize}[noitemsep,topsep=1pt]
\item \ce{Cu(OH)2}: tạo phức xanh lam (nhiệt độ thường) và kết tủa đỏ gạch \ce{Cu2O} (đun nóng).
\item Nước bromine: oxi hóa nhóm -CHO làm mất màu nước bromine.
\item Thuốc thử Tollens ($[\ce{Ag(NH3)2}]\ce{OH}$): phản ứng tráng bạc tạo kết tủa Ag.
\item \ce{CH3OH}/\ce{HCl} khan: phản ứng nhóm -OH hemiacetal tạo methyl glucoside.
\end{itemize}
$\Rightarrow$ Liệt kê: \textbf{\ce{Cu(OH)2}, nước bromine, thuốc thử Tollens, \ce{CH3OH}/\ce{HCl} khan}.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{6pt}
\multido{\i=1+1}{5}{\noindent\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\noindent\textbf{Câu 3.} Thực hiện phản ứng theo sơ đồ sau: $\text{Ala} + \ce{NaOH -> X + H2O}$.\\
Viết công thức cấu tạo và tính phân tử khối của X?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
- Phương trình: $\ce{CH3-CH(NH2)-COOH + NaOH -> CH3-CH(NH2)-COONa + H2O}$.\\
- CTCT của X: \textbf{\ce{CH3-CH(NH2)-COONa}} (sodium alaninate).\\
- Phân tử khối của X: $M_X = 89 + 23 - 1 = \mathbf{111\text{ g/mol}}$.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{6pt}
\multido{\i=1+1}{4}{\noindent\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\noindent\textbf{Câu 4.} Cây xanh nhả khí oxygen sau khi hoàn tất quá trình quang hợp, con người và nhiều loài động vật khác cần khí oxygen để duy trì sự sống. Phản ứng quang hợp được mô tả theo phương trình sau:\\
\centerline{$n\ce{CO2} + m\ce{H2O} \xrightarrow[\text{diệp lục}]{\text{ánh sáng}} \ce{C_n(H2O)_m} + n\ce{O2}$}\\
Giả sử một cây xanh trưởng thành tiêu thụ lượng \ce{CO2} trung bình khoảng 22 kg/năm để thực hiện quá trình quang hợp và trung bình một năm mỗi người cần tối thiểu 1,2 tấn oxygen để duy trì sự sống. Để cung cấp đủ lượng oxygen cho 8 tỉ người trên Trái đất trong một năm thì cần bao nhiêu tỉ cây xanh như trên?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
- Tổng lượng \ce{O2} cần cho 8 tỉ người trong 1 năm: $8 \times 10^9 \times 1{,}2\text{ tấn} = 9{,}6 \times 10^{12}\text{ kg}$.\\
- Theo phương trình: Cứ 44 kg \ce{CO2} giải phóng 32 kg \ce{O2}.\\
- Lượng \ce{O2} do 1 cây xanh tạo ra trong 1 năm: $m_{\ce{O2}} = 22 \times \frac{32}{44} = 16\text{ kg}$.\\
- Số cây xanh cần dùng: $\text{Số cây} = \frac{9{,}6 \times 10^{12}}{16} = 0{,}6 \times 10^{12}\text{ cây} = \mathbf{600\text{ tỉ cây xanh}}$.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{6pt}
\multido{\i=1+1}{6}{\noindent\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\noindent\textbf{Câu 5.} Tiến hành thí nghiệm theo các bước sau:\\
- Bước 1: Cho khoảng 1 mL nước bromine vào ống nghiệm.\\
- Bước 2: Thêm từ từ vài giọt dung dịch aniline loãng vào ống nghiệm.\\
Nêu hiện tượng và viết phương trình phản ứng?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
- \textbf{Hiện tượng:} Dung dịch nước bromine bị mất màu, đồng thời xuất hiện kết tủa trắng.\\
- \textbf{Phương trình phản ứng:}\\
$\ce{C6H5NH2 + 3Br2 -> C6H2Br3NH2\downarrow\text{ (trắng)} + 3HBr}$ (2,4,6-tribromoaniline).
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{6pt}
\multido{\i=1+1}{5}{\noindent\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\noindent\textbf{Câu 6.} Tiến hành thí nghiệm với các chất X, Y, Z là một trong các chất sau: hồ tinh bột, albumin, triolein, vinyl acetate. Kết quả được ghi ở bảng sau:
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|p{8.5cm}|c|}
\hline
\rowcolor{tableGreen}
\textbf{Mẫu thử} & \textbf{Thí nghiệm} & \textbf{Hiện tượng} \\
\hline
\textbf{X} & Đun nóng với dung dịch \ce{NaOH} (loãng, dư), để nguội. Thêm tiếp vài giọt dung dịch \ce{CuSO4} & Tạo dung dịch màu xanh lam \\
\hline
\textbf{Y} & Đun nóng với dung dịch \ce{NaOH} loãng (vừa đủ). Thêm tiếp dung dịch \ce{AgNO3} trong \ce{NH3}, đun nóng & Tạo kết tủa Ag \\
\hline
\textbf{Z} & Tác dụng với dung dịch \ce{I2} loãng & Có màu xanh tím \\
\hline
\end{tabular}
\end{center}
Tìm các chất X, Y, Z.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
- Mẫu Z tác dụng với \ce{I2} tạo màu xanh tím $\Rightarrow$ \textbf{Z là hồ tinh bột}.\\
- Mẫu Y thủy phân tạo \ce{CH3CHO} tham gia tráng bạc sinh kết tủa Ag $\Rightarrow$ \textbf{Y là vinyl acetate} (\ce{CH3COOCH=CH2}).\\
- Mẫu X thủy phân trong kiềm tạo glycerol, glycerol tác dụng với \ce{Cu(OH)2} tạo phức màu xanh lam $\Rightarrow$ \textbf{X là triolein} (\ce{(C17H33COO)3C3H5}).\\
\textbf{Kết luận:} X là triolein; Y là vinyl acetate; Z là hồ tinh bột.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{6pt}
\multido{\i=1+1}{5}{\noindent\dotfill\\[4pt]}
""")

    tex.append(r"""
\begin{center}
\textbf{--- HẾT ---}
\end{center}
\end{document}
""")

    return "".join(tex)


def compile_latex(tex_path, output_pdf):
    work_dir = os.path.dirname(tex_path)
    base_name = os.path.splitext(os.path.basename(tex_path))[0]
    cmd = ["pdflatex", "-interaction=nonstopmode", f"{base_name}.tex"]
    subprocess.run(cmd, cwd=work_dir, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    subprocess.run(cmd, cwd=work_dir, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    pdf_temp = os.path.join(work_dir, f"{base_name}.pdf")
    if os.path.exists(pdf_temp) and pdf_temp != output_pdf:
        if os.path.exists(output_pdf):
            os.remove(output_pdf)
        os.rename(pdf_temp, output_pdf)

    # Dọn dẹp rác
    for ext in [".aux", ".log", ".thm", ".out", ".toc"]:
        junk = os.path.join(work_dir, f"{base_name}{ext}")
        if os.path.exists(junk):
            os.remove(junk)
    for ans_f in glob.glob(os.path.join(work_dir, "ans-*.tex")):
        os.remove(ans_f)


def main():
    os.makedirs(SAN_PHAM_DIR, exist_ok=True)

    # 1. Bản Học sinh
    tex_hs = build_latex_content(is_teacher=False)
    tex_hs_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma103_HocSinh.tex")
    with open(tex_hs_path, "w", encoding="utf-8") as f:
        f.write(tex_hs)
    pdf_hs_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma103_HocSinh.pdf")
    compile_latex(tex_hs_path, pdf_hs_path)
    print(f"[OK] Đã biên dịch xong PDF Học sinh: {pdf_hs_path}")

    # 2. Bản Giáo viên
    tex_gv = build_latex_content(is_teacher=True)
    tex_gv_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma103_GiaoVien.tex")
    with open(tex_gv_path, "w", encoding="utf-8") as f:
        f.write(tex_gv)
    pdf_gv_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma103_GiaoVien.pdf")
    compile_latex(tex_gv_path, pdf_gv_path)
    print(f"[OK] Đã biên dịch xong PDF Giáo viên: {pdf_gv_path}")


if __name__ == "__main__":
    main()
