# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Giữa Kỳ I Hóa Học 12 - THPT Nguyễn Trường Tộ - Mã Đề 106
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma106_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma106_GiaoVien.pdf
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
& \textbf{Mã đề: 106}
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
\noindent\textit{Cho biết nguyên tử khối của các nguyên tố: H = 1; C = 12; N = 14; O = 16; Na = 23; S = 32; Cu = 64; Ag = 108.}
\vspace{2pt}
\hrule height 0.5pt
\vspace{6pt}
""")

    if is_teacher:
        tex.append(r"""
\begin{center}
{\color{hfRed}\textbf{\large BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 106)}}
\end{center}

\noindent{\color{mainGreen}\textbf{1. Phần I: Trắc nghiệm nhiều lựa chọn (16 câu)}}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|*{8}{>{\centering\arraybackslash}m{0.95cm}|}}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} \\
\hline
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}A} \\
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
\textbf{Câu 1} & Đúng / Sai & \textbf{S} & \textbf{S} & \textbf{\color{hfRed}Đ} & \textbf{\color{hfRed}Đ} \\
\hline
\textbf{Câu 2} & Đúng / Sai & \textbf{S} & \textbf{S} & \textbf{\color{hfRed}Đ} & \textbf{S} \\
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
\textbf{Đáp án} & \textbf{\color{hfRed}2159} & \textbf{\color{hfRed}5} & \textbf{\color{hfRed}3} & \textbf{\color{hfRed}22} \\
\hline
\end{tabular}
\end{center}
\vspace{6pt}
\hrule height 0.5pt
\vspace{6pt}
""")

    # PHẦN A. TRẮC NGHIỆM
    tex.append(r"""
\noindent{\color{mainGreen}\textbf{\Large A. PHẦN TRẮC NGHIỆM}}
\vspace{4pt}

\noindent{\color{mainGreen}\textbf{\large Phần I: Câu trắc nghiệm nhiều phương án lựa chọn.}} \\
\textit{Thí sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{4pt}
""")

    # 16 câu Phần I
    p1_items = [
        ("Khi nấu canh cua thấy các mảng ``gạch cua'' nổi lên là do",
         ["sự đông tụ lipid.", "phản ứng thủy phân protein.", "phản ứng màu biuret.", "sự đông tụ protein."],
         3,
         "Thịt cua chứa nhiều protein hòa tan, khi đun nóng protein bị biến tính và đông tụ tạo thành các mảng nổi lên trên bề mặt nước canh."),

        ("Công thức cấu tạo của methylamine là",
         [r"\ce{C2H5NH2}.", r"\ce{CH3NH2}.", r"\ce{CH3NHCH3}.", r"\ce{(CH3)3N}."],
         1,
         "Methylamine có công thức cấu tạo là \ce{CH3NH2}."),

        (r"Dung dịch chất X hòa tan \ce{Cu(OH)2}, thu được dung dịch màu xanh lam. Mặt khác, X bị thủy phân khi đun nóng trong môi trường acid. Chất X là",
         ["glucose.", "saccharose.", "amylose.", "fructose."],
         1,
         r"Saccharose (\ce{C12H22O11}) là disaccharide chứa nhiều nhóm -OH kề nhau nên hòa tan \ce{Cu(OH)2} tạo dung dịch xanh lam, đồng thời bị thủy phân trong acid tạo glucose và fructose."),

        (r"Ester X có công thức phân tử \ce{C4H8O2}. Thủy phân X trong dung dịch \ce{H2SO4} loãng, đun nóng, thu được sản phẩm gồm acetic acid và chất hữu cơ Y. Công thức của Y là",
         [r"\ce{HCOOH}.", r"\ce{C2H5OH}.", r"\ce{CH3COOH}.", r"\ce{CH3OH}."],
         1,
         r"Ester X là ethyl acetate: $\ce{CH3COOC2H5 + H2O <=> CH3COOH + C2H5OH}$. Vậy Y là alcohol \ce{C2H5OH}."),

        ("Mùi tanh của cá là hỗn hợp các amine và một số tạp chất khác. Để khử mùi tanh của cá trước khi chế biến thực phẩm, nên áp dụng cách nào sau đây?",
         ["Rửa cá bằng dung dịch nước muối.", "Ngâm cá trong nước để amine tan vào nước.", "Rửa cá bằng dung dịch soda.", "Rửa cá bằng giấm ăn."],
         3,
         "Các amine gây mùi tanh có tính base yếu, khi rửa bằng giấm ăn (chứa acetic acid \ce{CH3COOH}) xảy ra phản ứng trung hòa tạo muối tan dễ bị nước rửa trôi."),

        ("Khi thuỷ phân lipid trong môi trường base thì thu được muối của acid béo và",
         ["ester đơn chức.", "phenol.", "alcohol đơn chức.", "glycerol."],
         3,
         "Lipid (triglyceride) thủy phân trong dung dịch kiềm (xà phòng hóa) luôn sinh ra muối của acid béo và glycerol: $\ce{(RCOO)3C3H5 + 3NaOH -> 3RCOONa + C3H5(OH)3}$."),

        (r"Giá trị pH mà khi đó amino acid có nồng độ ion lưỡng cực là cực đại được gọi là điểm đẳng điện (kí hiệu là pI). Khi $\text{pH} < \text{pI}$ thì amino acid đó tồn tại chủ yếu ở dạng cation, còn khi $\text{pH} > \text{pI}$ thì amino acid đó tồn tại chủ yếu ở dạng anion. Khi đặt trong một điện trường dạng anion sẽ di chuyển về cực $(+)$ còn dạng cation sẽ di chuyển về cực $(-)$. Cho các giá trị pI: glycine ($6{,}0$); glutamic acid ($3{,}2$); lysine ($9{,}7$). Giá trị pH tối ưu nhất để tách ba chất trên ra khỏi dung dịch hỗn hợp là",
         [r"$\text{pH} = 14{,}0$.", r"$\text{pH} = 9{,}7$.", r"$\text{pH} = 3{,}2$.", r"$\text{pH} = 6{,}0$."],
         3,
         r"Tại $\text{pH} = 6{,}0$: glycine có $\text{pH} = \text{pI} = 6{,}0$ nên ở dạng ion lưỡng cực không tích điện (đứng yên); glutamic acid ($\text{pI} = 3{,}2 < 6{,}0$) tồn tại dạng anion di chuyển về cực $(+)$; lysine ($\text{pI} = 9{,}7 > 6{,}0$) tồn tại dạng cation di chuyển về cực $(-)$. Ba chất di chuyển theo 3 hướng khác nhau giúp tách triệt để."),

        (r"Tripeptide là hợp chất mà phân tử có",
         [r"ba liên kết peptide, hai gốc $\alpha$-amino acid.",
          r"hai liên kết peptide, ba gốc $\alpha$-amino acid.",
          r"ba liên kết peptide, ba gốc $\alpha$-amino acid.",
          r"hai liên kết peptide, ba gốc $\beta$-amino acid."],
         1,
         r"Tripeptide gồm 3 gốc $\alpha$-amino acid liên kết với nhau bằng 2 liên kết peptide (-CO-NH-)."),

        (r"Chất nào sau đây vừa tác dụng được với \ce{H2NCH2COOH}, vừa tác dụng được với \ce{CH3NH2}?",
         [r"\ce{NaOH}.", r"\ce{HCl}.", r"\ce{NaCl}.", r"\ce{CH3OH}."],
         1,
         r"\ce{CH3NH2} là amine có tính base, tác dụng được với acid \ce{HCl}. \ce{H2NCH2COOH} lưỡng tính cũng tác dụng được với \ce{HCl}."),

        ("Phần không phân cực (``đuôi'' kị nước) trong xà phòng và chất giặt rửa là",
         ["nhóm sulfonate.", "gốc hydrocarbon dài.", "nhóm sulfate.", "Nhóm carboxylate."],
         1,
         "Phân tử xà phòng và chất giặt rửa có 'đầu' phân cực ưa nước (-COO-, -SO3-) và 'đuôi' kị nước là gốc hydrocarbon mạch dài."),

        ("Glucose là một loại monosaccharide có nhiều trong quả nho chín. Công thức phân tử của glucose là",
         [r"\ce{C6H12O6}.", r"\ce{(C6H10O5)_n}.", r"\ce{C2H4O2}.", r"\ce{C12H22O11}."],
         0,
         "Glucose có công thức phân tử là \ce{C6H12O6}."),

        ("Nhóm chức nào sau đây không có trong cấu tạo của glucose?",
         ["Aldehyde.", "Ketone.", "Hemiacetal.", "Hydroxy."],
         1,
         "Glucose mạch hở có nhóm -CHO và các nhóm -OH; dạng mạch vòng có nhóm hemiacetal. Glucose không có nhóm chức ketone (nhóm C=O ketone có trong fructose)."),

        (r"Ở động vật, tinh bột được dự trữ dưới dạng glycogen ở đâu?",
         ["Ruột", "Dạ dày", "Gan", "Máu"],
         2,
         "Glycogen được tổng hợp và dự trữ chủ yếu ở gan và các mô cơ."),

        (r"Số ester có cùng công thức phân tử \ce{C3H6O2} là",
         ["4.", "3.", "2.", "5."],
         2,
         r"Có 2 đồng phân ester: \ce{HCOOC2H5} và \ce{CH3COOCH3}."),

        (r"Monosaccharide X được dùng trong công nghiệp để tráng bạc lên bề mặt thuỷ tinh trong sản xuất ruột phích. Cùng với Ag, sản phẩm hữu cơ được tạo thành khi cho X tác dụng với lượng dư dung dịch \ce{AgNO3} trong \ce{NH3} là",
         ["ammonium gluconate.", "ammonium carbonate.", "gluconic acid.", "khí carbon dioxide."],
         0,
         r"Glucose phản ứng với thuốc thử Tollens tạo muối ammonium gluconate: $\ce{CH2OH[CHOH]4CHO + 2[Ag(NH3)2]OH -> CH2OH[CHOH]4COONH4 + 2Ag\downarrow + 3NH3 + H2O}$."),

        (r"Chất có công thức phân tử \ce{C12H22O11}, được tạo thành trong quá trình thuỷ phân không hoàn toàn amylose có trong tinh bột là",
         ["maltose.", "glucose.", "saccharose.", "fructose."],
         0,
         "Thủy phân không hoàn toàn amylose trong tinh bột tạo disaccharide maltose (\ce{C12H22O11}).")
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

    # PHẦN II. ĐÚNG SAI
    tex.append(r"""
\vspace{6pt}
\noindent{\color{mainGreen}\textbf{\large Phần II: Câu hỏi trả lời đúng sai.}} \\
\textit{Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{4pt}
""")

    # CÂU 1 PHẦN II
    tex.append(r"""
\begin{ex}
Tiến hành thí nghiệm thủy phân ester theo các bước sau:\\
- Bước 1: Cho vào hai ống nghiệm mỗi ống 2 mL ethyl acetate.\\
- Bước 2: Thêm 2 mL dung dịch \ce{H2SO4} 20\% vào ống thứ nhất; 4 mL dung dịch \ce{NaOH} 30\% vào ống thứ hai.\\
- Bước 3: Lắc đều cả hai ống nghiệm, lắp ống sinh hàn, đun sôi nhẹ trong khoảng 5 phút, để nguội.
\choiceTF
""")
    tf1 = [
        ("Phản ứng xảy ra trong ống nghiệm 1 nhanh hơn phản ứng trong ống 2.", False),
        (r"Vai trò của dung dịch \ce{H2SO4} 20\% và dung dịch \ce{NaOH} 30\% trong thí nghiệm trên là chất xúc tác và không bị tiêu hao sau phản ứng.", False),
        ("Sau bước 2, chất lỏng trong cả hai ống nghiệm đều phân thành hai lớp.", True),
        ("Sau bước 3: ống nghiệm (1) thể tích lớp chất lỏng phía trên giảm. Ống nghiệm (2) tạo thành hỗn hợp đồng nhất.", True)
    ]
    for stmt, val in tf1:
        if val:
            tex.append(f"\\True {stmt}\n")
        else:
            tex.append(f"{stmt}\n")
    if is_teacher:
        tex.append(r"""\loigiai{
\textbf{Ý a) SAI.} Phản ứng xà phòng hóa trong môi trường kiềm (ống 2) là phản ứng một chiều xảy ra nhanh hơn thủy phân thuận nghịch trong môi trường acid (ống 1).\\
\textbf{Ý b) SAI.} \ce{H2SO4} là chất xúc tác, còn \ce{NaOH} là chất tham gia phản ứng và bị tiêu hao.\\
\textbf{Ý c) ĐÚNG.} Ethyl acetate nhẹ hơn nước và không tan nên nổi lên tạo 2 lớp ở cả 2 ống nghiệm trước khi đun nóng.\\
\textbf{Ý d) ĐÚNG.} Ống 1 phản ứng thuận nghịch tiêu hao một phần ester nên lớp nổi giảm; ống 2 phản ứng hoàn toàn tạo \ce{CH3COONa} và \ce{C2H5OH} tan vô hạn tạo dung dịch đồng nhất.
}
""")
    else:
        tex.append("\\loigiai{}\n")
    tex.append("\\end{ex}\n")

    # CÂU 2 PHẦN II
    tex.append(r"""
\begin{ex}
Hình dưới đây mô tả cấu tạo phân tử saccharose gồm một gốc $\alpha$-glucose và một gốc $\beta$-fructose liên kết với nhau qua liên kết $\alpha$-1,2-glycoside:
\begin{center}
\includegraphics[width=0.65\textwidth]{C:/Users/Admin/.gemini/antigravity/brain/c29808c0-7d22-4aec-89d3-7a04648f0518/.user_uploaded/media_1794322695260_e258ffd0.jpg}
\end{center}
\choiceTF
""")
    tf2 = [
        ("Ngoài cấu tạo dạng mạch vòng, saccharose còn có cả cấu tạo ở dạng mở vòng của đơn vị cấu tạo glucose và fructose.", False),
        (r"Mỗi phân tử saccharose gồm một đơn vị $\alpha$-glucose và một đơn vị $\alpha$-fructose liên kết với nhau qua liên kết $\alpha$-1,2-glycoside.", False),
        ("Saccharose và maltose đều có 42,1\\% C về khối lượng.", True),
        ("Phân tử saccharose không có nhóm -OH hemiacetal, nhưng có nhóm -OH hemiketal.", False)
    ]
    for stmt, val in tf2:
        if val:
            tex.append(f"\\True {stmt}\n")
        else:
            tex.append(f"{stmt}\n")
    if is_teacher:
        tex.append(r"""\loigiai{
\textbf{Ý a) SAI.} Các nhóm -OH hemiacetal và hemiketal đều đã bị khóa trong liên kết glycosidic nên saccharose không thể mở vòng.\\
\textbf{Ý b) SAI.} Saccharose gồm một gốc $\alpha$-glucose và một gốc $\beta$-fructose (đề ghi $\alpha$-fructose là sai).\\
\textbf{Ý c) ĐÚNG.} Cùng công thức \ce{C12H22O11}: $\%C = \frac{12 \times 12}{342} \times 100\% = 42{,}1\%$.\\
\textbf{Ý d) SAI.} Cả hai nhóm -OH hemiacetal và hemiketal đều không còn trong phân tử saccharose.
}
""")
    else:
        tex.append("\\loigiai{}\n")
    tex.append("\\end{ex}\n")

    # PHẦN III. TRẢ LỜI NGẮN
    tex.append(r"""
\vspace{6pt}
\noindent{\color{mainGreen}\textbf{\large Phần III: Câu trắc nghiệm trả lời ngắn.}} \\
\textit{Thí sinh trả lời từ câu 1 đến câu 4.}
\vspace{4pt}
""")

    # CÂU 1 PHẦN III
    tex.append(r"""
\begin{ex}
Ethanol có thể được sản xuất từ cellulose hoặc tinh bột. Loại ethanol này được dùng để sản xuất xăng E10 (xăng chứa 10\% ethanol về thể tích). Lượng ethanol thu được từ 1 tấn mùn cưa (chứa 50\% cellulose, phần còn lại là chất trơ) có thể dùng để pha chế bao nhiêu L xăng E10? Biết hiệu suất quá trình sản xuất ethanol từ cellulose là 60\% và ethanol có khối lượng riêng là 0,789 g/mL. \textit{(Kết quả làm tròn đến hàng đơn vị)}.
\shortans{2159}
\loigiai{
\textbf{Trả lời: 2159}\\
$m_{\text{cellulose}} = 1000 \times 50\% = 500\text{ kg} = 500\,000\text{ gam}$.\\
Sơ đồ: $\ce{(C6H10O5)_n -> n C6H12O6 -> 2n C2H5OH}$\\
$m_{\text{ethanol}} = 500\,000 \times \frac{92}{162} \times 60\% = 170\,370{,}37\text{ gam}$.\\
$V_{\text{ethanol}} = \frac{170\,370{,}37}{0{,}789} = 215\,932\text{ mL} = 215{,}932\text{ L}$.\\
$V_{\text{xăng E10}} = \frac{215{,}932}{0{,}10} = 2159{,}32\text{ L} \approx \mathbf{2159}$.
}
\end{ex}
""")

    # CÂU 2 PHẦN III
    tex.append(r"""
\begin{ex}
Cho các chất sau: ethyl acetate, tripalmitin, saccharose, ethylamine, Gly-Ala, albumin, Ala-Ala-Val. Số chất tham gia phản ứng thủy phân trong môi trường kiềm là bao nhiêu?
\shortans{5}
\loigiai{
\textbf{Trả lời: 5}\\
Các chất thủy phân trong môi trường kiềm gồm 5 chất:
\begin{enumerate}[noitemsep,topsep=1pt]
\item Ethyl acetate (ester).
\item Tripalmitin (triglyceride / chất béo).
\item Gly-Ala (dipeptide).
\item Albumin (protein).
\item Ala-Ala-Val (tripeptide).
\end{enumerate}
(Saccharose không bị thủy phân trong kiềm; ethylamine là amine không bị thủy phân).
}
\end{ex}
""")

    # CÂU 3 PHẦN III
    tex.append(r"""
\begin{ex}
Cho các dung dịch sau: \ce{HCl}, \ce{NaOH}, \ce{CH3COOH}, \ce{CuSO4}. Có bao nhiêu dung dịch tác dụng được với ethylamine?
\shortans{3}
\loigiai{
\textbf{Trả lời: 3}\\
Ethylamine (\ce{C2H5NH2}) tác dụng với 3 dung dịch:
\begin{itemize}[noitemsep,topsep=1pt]
\item \ce{HCl}: $\ce{C2H5NH2 + HCl -> C2H5NH3Cl}$.
\item \ce{CH3COOH}: $\ce{C2H5NH2 + CH3COOH -> CH3COONH3C2H5}$.
\item \ce{CuSO4}: tạo kết tủa \ce{Cu(OH)2} rồi hòa tan tạo phức xanh lam $[\ce{Cu(C2H5NH2)4}]^{2+}$.
\end{itemize}
(\ce{NaOH} không phản ứng vì cùng có tính base).
}
\end{ex}
""")

    # CÂU 4 PHẦN III
    tex.append(r"""
\begin{ex}
Một phân tử saccharose có bao nhiêu nguyên tử hydrogen?
\shortans{22}
\loigiai{
\textbf{Trả lời: 22}\\
Công thức phân tử saccharose là \ce{C12H22O11}, do đó có 22 nguyên tử hydrogen.
}
\end{ex}
""")

    # PHẦN B. TỰ LUẬN
    tex.append(r"""
\vspace{6pt}
\noindent{\color{mainGreen}\textbf{\Large B. PHẦN TỰ LUẬN}}
\vspace{4pt}

\noindent\textbf{Câu 1.} Hoàn thành phương trình phản ứng sau:\\
\ce{CH3COOC6H5} (phenyl acetate) + \ce{NaOH} dư (dung dịch, đun nóng).
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
Phương trình phản ứng:\\
$\ce{CH3COOC6H5 + 2NaOH ->[t^o] CH3COONa + C6H5ONa + H2O}$
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{6pt}
\multido{\i=1+1}{4}{\noindent\dotfill\\[4pt]}
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
\item Thuốc thử Tollens ($[\ce{Ag(NH3)2}]\ce{OH}$): phản ứng tráng bạc tạo kết tủa \ce{Ag}.
\item \ce{CH3OH}/\ce{HCl} khan: phản ứng nhóm -OH hemiacetal tạo methyl glucoside.
\end{itemize}
(Dung dịch \ce{I2} trong \ce{KI} không phản ứng với glucose).\\
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
\noindent\textbf{Câu 3.} Viết công thức cấu tạo của Alanine.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
Công thức cấu tạo của Alanine:\\
\textbf{\ce{CH3-CH(NH2)-COOH}} (hoặc \ce{H2N-CH(CH3)-COOH}).
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{6pt}
\multido{\i=1+1}{3}{\noindent\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\noindent\textbf{Câu 4.} Trong cơ thể, glucose là sản phẩm của quá trình thủy phân tinh bột dưới tác dụng của enzyme. Trong quá trình hô hấp tế bào, 1 mol glucose bị oxi hóa hoàn toàn qua các phản ứng sinh hóa tạo thành carbon dioxide, nước và giải phóng năng lượng là 2880 kJ. Trong khẩu phần ăn hằng ngày của một người trưởng thành cần 300 gam gạo. Hàm lượng tinh bột trong gạo là 80\%. Giả sử 60\% lượng tinh bột trong gạo chuyển hóa thành glucose và toàn bộ lượng glucose này tham gia vào quá trình hô hấp tế bào. Năng lượng giải phóng từ sự oxi hóa hoàn toàn glucose chiếm bao nhiêu \% so với nhu cầu năng lượng trung bình (2000 Calori/ngày) của một người trưởng thành? \textit{(Kết quả làm tròn đến hàng phần mười)}.\\
\textit{Cho biết: 1 Calori = 1 kcal = 1000 cal = 4184 J.}
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
- Nhu cầu năng lượng hàng ngày: $E_{\text{nhu cầu}} = 2000 \times 4184\text{ J} = 8\,368\,000\text{ J} = 8368\text{ kJ}$.\\
- Khối lượng tinh bột trong 300 gam gạo: $m = 300 \times 80\% = 240\text{ gam}$.\\
- Khối lượng tinh bột thực tế chuyển hóa: $m_{\text{pư}} = 240 \times 60\% = 144\text{ gam}$.\\
- Phương trình: $\ce{(C6H10O5)_n -> n C6H12O6} \Rightarrow n_{\text{glucose}} = \frac{144}{162} = \frac{8}{9}\text{ mol}$.\\
- Năng lượng giải phóng: $E_{\text{giải phóng}} = \frac{8}{9} \times 2880 = 2560\text{ kJ}$.\\
- Phần trăm năng lượng: $\% = \frac{2560}{8368} \times 100\% = \mathbf{30{,}6\%}$.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{6pt}
\multido{\i=1+1}{6}{\noindent\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\noindent\textbf{Câu 5.} Amine X có phân tử khối nhỏ hơn 80. Trong phân tử X, nitrogen chiếm 19,18\% về khối lượng. Cho X tác dụng với nitrous acid (\ce{HNO2}) ở nhiệt độ thường thu được alcohol Y. Oxi hóa không hoàn toàn Y thu được ketone Z. Viết công thức cấu tạo của Y.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
- Phân tử khối amine $M_X < 80$ và $\%N = 19{,}18\% \Rightarrow$ số nguyên tử $\ce{N} = 1 \Rightarrow M_X = \frac{14}{19{,}18\%} \approx 73\text{ g/mol}$.\\
- X tác dụng với \ce{HNO2} ở nhiệt độ thường cho alcohol Y $\Rightarrow$ X là amine no, đơn chức, bậc 1 dạng $\ce{C_nH_{2n+1}NH2}$.\\
$\Rightarrow 14n + 17 = 73 \Rightarrow n = 4 \Rightarrow$ CTPT là \ce{C4H11N}, alcohol Y là \ce{C4H9OH}.\\
- Oxi hóa không hoàn toàn Y thu được ketone Z $\Rightarrow$ Y phải là alcohol bậc 2.\\
$\Rightarrow$ CTCT của alcohol Y là: \textbf{\ce{CH3-CH(OH)-CH2-CH3}} (butan-2-ol).
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{6pt}
\multido{\i=1+1}{6}{\noindent\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\noindent\textbf{Câu 6.} Tiến hành thí nghiệm với các chất X, Y, Z là một trong các chất sau: hồ tinh bột, albumin, ethyl acetate, vinyl acetate. Kết quả được ghi ở bảng sau:
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|p{8.5cm}|c|}
\hline
\rowcolor{tableGreen}
\textbf{Mẫu thử} & \textbf{Thí nghiệm} & \textbf{Hiện tượng} \\
\hline
\textbf{X} & Tác dụng với \ce{Cu(OH)2} trong môi trường kiềm & Có màu tím \\
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
- Mẫu X tạo màu tím với \ce{Cu(OH)2} trong môi trường kiềm $\Rightarrow$ phản ứng màu biuret đặc trưng của protein $\Rightarrow$ \textbf{X là albumin}.\\
- Mẫu Z tác dụng với \ce{I2} tạo màu xanh tím đặc trưng $\Rightarrow$ \textbf{Z là hồ tinh bột}.\\
- Mẫu Y thủy phân trong kiềm tạo chất tráng bạc:\\
$\ce{CH3COOCH=CH2 + NaOH ->[t^o] CH3COONa + CH3CHO}$\\
$\ce{CH3CHO + 2[Ag(NH3)2]OH ->[t^o] CH3COONH4 + 2Ag\downarrow + 3NH3 + H2O}$\\
$\Rightarrow$ \textbf{Y là vinyl acetate} (\ce{CH3COOCH=CH2}).\\
\textbf{Kết luận:} X là albumin; Y là vinyl acetate; Z là hồ tinh bột.
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
    tex_hs_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma106_HocSinh.tex")
    with open(tex_hs_path, "w", encoding="utf-8") as f:
        f.write(tex_hs)
    pdf_hs_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma106_HocSinh.pdf")
    compile_latex(tex_hs_path, pdf_hs_path)
    print(f"[OK] Đã biên dịch xong PDF Học sinh: {pdf_hs_path}")

    # 2. Bản Giáo viên
    tex_gv = build_latex_content(is_teacher=True)
    tex_gv_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma106_GiaoVien.tex")
    with open(tex_gv_path, "w", encoding="utf-8") as f:
        f.write(tex_gv)
    pdf_gv_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenTruongTo_Ma106_GiaoVien.pdf")
    compile_latex(tex_gv_path, pdf_gv_path)
    print(f"[OK] Đã biên dịch xong PDF Giáo viên: {pdf_gv_path}")


if __name__ == "__main__":
    main()
