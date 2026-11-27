# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Cuối Học Kì I Hóa Học 11 - THPT Phong Điền - Huế
Mã đề: 1020 - Năm học 2025 - 2026
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_CK1_Hoa_11_PhongDien_Ma1020_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_CK1_Hoa_11_PhongDien_Ma1020_GiaoVien.pdf (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
\usepackage{chemfig}
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
\textbf{SỞ GD\&ĐT THÀNH PHỐ HUẾ}\\
\textbf{TRƯỜNG THPT PHONG ĐIỀN}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 04 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}ĐỀ KIỂM TRA CUỐI KÌ I}\\
\textbf{NĂM HỌC 2025 -- 2026}\\
\textbf{MÔN: HÓA HỌC LỚP 11}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{MÃ ĐỀ: 1020}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ và tên:} \dotfill \textit{Số báo danh:} \dotfill

\vspace{3pt}
\noindent
\textit{\footnotesize Cho biết nguyên tử khối của các nguyên tố: $\ce{Na} = 23; \ce{Cu} = 64; \ce{Ag} = 108; \ce{Fe} = 56; \ce{S} = 32; \ce{N} = 14; \ce{O} = 16; \ce{H} = 1; \ce{Cl} = 35{,}5; \ce{Br} = 80$.}

\vspace{3pt}
\hrule color{hfGreen} height 0.5pt
""")

    if is_teacher:
        tex.append(r"""
\vspace{6pt}
\begin{center}
{\large\bfseries\color{hfRed} BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 1020)}
\end{center}

\vspace{4pt}
{\bfseries\color{mainGreen} 1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (3,0 ĐIỂM)}

\vspace{2pt}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} \\ \hline
\textbf{Đ/A} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} \\ \hline
\end{tabular}
\end{center}

\vspace{4pt}
{\bfseries\color{mainGreen} 2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM)}

\vspace{2pt}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|p{7.5cm}|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Nội dung câu hỏi} & \textbf{Lệnh a} & \textbf{Lệnh b} & \textbf{Lệnh c} & \textbf{Lệnh d} \\ \hline
\textbf{Câu 1} & Thí nghiệm đốt cháy hexane trong bát sứ & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{SAI} \\ \hline
\textbf{Câu 2} & Thí nghiệm chưng cất tách ethanol và nước & \textbf{SAI} & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{\color{hfRed}ĐÚNG} \\ \hline
\end{tabular}
\end{center}

\vspace{4pt}
{\bfseries\color{mainGreen} 3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM)}

\vspace{2pt}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu hỏi} & \textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} \\ \hline
\textbf{Đáp án} & \textbf{\color{hfRed}2 - 3 - 1} & \textbf{\color{hfRed}5} & \textbf{\color{hfRed}6, 8, 6} & \textbf{\color{hfRed}4} \\ \hline
\end{tabular}
\end{center}

\vspace{4pt}
{\bfseries\color{mainGreen} 4. TÓM TẮT ĐÁP SỐ PHẦN TỰ LUẬN (3,0 ĐIỂM)}

\vspace{2pt}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|p{13.2cm}|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Điểm} & \textbf{Đáp số cốt lõi \& Hướng dẫn chấm} \\ \hline
\textbf{Câu 1} & 0,5 đ & 3 luận điểm: Trật tự liên kết (cấu tạo hóa học); Carbon hóa trị 4 tạo các loại mạch C; Tính chất phụ thuộc vào thành phần và cấu tạo hóa học. \\ \hline
\textbf{Câu 2} & 0,5 đ & 3 bước: 1) Chọn mạch chính dài nhất, nhiều nhánh nhất; 2) Đánh số C từ đầu gần nhánh hơn; 3) Tên = Vị trí nhánh - Tên nhánh + Tên alkane mạch chính. \\ \hline
\textbf{Câu 3} & 0,5 đ & a) $\ce{Cu(OH)2 + H2SO4(đặc, nóng) -> CuSO4 + 2H2O}$ (0,25 đ)\\
& & b) $\ce{2Fe + 6H2SO4(đặc, nóng) -> Fe2(SO4)3 + 3SO2 ^ + 6H2O}$ (0,25 đ) \\ \hline
\textbf{Câu 4} & 0,5 đ & Dựa vào sự đổi màu đột ngột của chất chỉ thị màu acid - base (phenolphthalein/methyl da cam), do tại điểm tương đương có bước nhảy pH làm chất chỉ thị biến đổi cấu trúc và đổi màu. \\ \hline
\textbf{Câu 5} & 1,0 đ & 1. a) Phản ứng tỏa nhiệt ($\Delta H < 0$). b) Không dùng nhiệt độ quá cao vì làm cân bằng chuyển dịch chiều nghịch (giảm hiệu suất $\ce{SO3}$) và giảm tuổi thọ $\ce{V2O5}$. (0,5 đ)\\
& & 2. Điều chỉnh: Hạ nhiệt độ vừa phải ($\approx 450^\circ\text{C}$); Tăng áp suất; Tăng nồng độ $\ce{SO2}$, $\ce{O2}$ (dùng dư không khí) và liên tục tách $\ce{SO3}$ ra khỏi hệ. (0,5 đ) \\ \hline
\end{tabular}
\end{center}

\vspace{4pt}
\hrule color{hfGreen} height 0.5pt
""")

    # IR table
    tex.append(r"""
\vspace{4pt}
\textit{\small Bảng tín hiệu phổ hồng ngoại của một số nhóm chức cơ bản:}
\begin{center}
\small
\renewcommand{\arraystretch}{1.15}
\begin{tabular}{|l|c|c||l|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Hợp chất} & \textbf{Liên kết} & \textbf{Số sóng ($\text{cm}^{-1}$)} & \textbf{Hợp chất} & \textbf{Liên kết} & \textbf{Số sóng ($\text{cm}^{-1}$)} \\ \hline
Alcohol & $\ce{O-H}$ & $3500 - 3200$ & Carboxylic & $\ce{C=O}$ & $1760 - 1690$ \\ \cline{1-3}
Ketone & $\ce{C=O}$ & $1715 - 1666$ & acid & $\ce{O-H}$ & $3300 - 2500$ \\ \hline
Aldehyde & $\ce{C-H}$ & $2830 - 2695$ & Ester & $\ce{C=O}$ & $1750 - 1715$ \\ \cline{2-3} \cline{5-6}
& $\ce{C=O}$ & $1740 - 1685$ & & $\ce{C-O}$ & $1300 - 1000$ \\ \hline
\end{tabular}
\end{center}

\vspace{4pt}
{\large\bfseries\color{mainGreen} PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,0 điểm)}

\textit{Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{4pt}
""")

    # Questions 1 to 12
    questions_p1 = [
        # Cau 1
        (r"""Để pha loãng $\ce{H2SO4}$ đặc cách làm nào sau đây đúng?
\begin{center}
\includegraphics[width=0.55\textwidth]{images_phongdien_ma1020/pha_loang_h2so4_clean.png}
\end{center}""",
         r"""cách 1.""",
         r"""cách 2.""",
         r"""cách 3.""",
         r"""cách 1 và 2.""",
         1,
         r"""Pha loãng sulfuric acid đặc phải rót từ từ acid vào nước và khuấy nhẹ bằng đũa thủy tinh (cách 1), tuyệt đối không làm ngược lại vì acid háo nước tỏa nhiệt dữ dội làm nước sôi bắn acid ra ngoài."""),

        # Cau 2
        (r"""Hợp chất hữu cơ nào sau đây có mạch carbon phân nhánh?""",
         r"""$\ce{CH3-CH(OH)-CH3}$""",
         r"""$\ce{cyclo-C4H8}$""",
         r"""$\ce{CH3-CH(CH3)-CH3}$""",
         r"""$\ce{CH3-CH2-CH2-CH2-CH3}$""",
         3,
         r"""$\ce{CH3-CH(CH3)-CH3}$ (isobutane) có 1 nhánh methyl ($-\ce{CH3}$) tại vị trí C2 nên là mạch carbon phân nhánh."""),

        # Cau 3
        (r"""Kết tinh là phương pháp tách biệt và tinh chế hỗn hợp các chất ....(1) dựa vào độ tan khác nhau và sự thay đổi độ tan của chúng theo .....(2). (1) và (2) lần lượt là""",
         r"""rắn -- thời gian.""",
         r"""lỏng -- nhiệt độ.""",
         r"""rắn -- nhiệt độ.""",
         r"""lỏng -- thời gian.""",
         3,
         r"""Kết tinh là phương pháp tách biệt và tinh chế hỗn hợp các chất rắn (1) dựa vào độ tan khác nhau và sự thay đổi độ tan của chúng theo nhiệt độ (2)."""),

        # Cau 4
        (r"""Trước kia, “phẩm đỏ” dùng để nhuộm áo choàng cho các Hồng y giáo chủ được tách chiết từ một loài ốc biển. Đó là một hợp chất có thành phần nguyên tố như sau: 45,70 \%C; 1,90 \%H; 7,60 \%O; 6,70 \%N; 38,10 \%Br. Công thức đơn giản của phẩm đỏ là""",
         r"""\ce{C8H4ONBr}.""",
         r"""\ce{C4H2ONBr}.""",
         r"""\ce{C4H8O2NBr2}.""",
         r"""\ce{C2H4ONBr}.""",
         1,
         r"""Tỉ lệ mol: $x : y : z : t : u = \frac{45{,}70}{12} : \frac{1{,}90}{1} : \frac{7{,}60}{16} : \frac{6{,}70}{14} : \frac{38{,}10}{80} = 3{,}808 : 1{,}900 : 0{,}475 : 0{,}479 : 0{,}476 \approx 8 : 4 : 1 : 1 : 1 \Rightarrow \ce{C8H4ONBr}$."""),

        # Cau 5
        (r"""Điền vào khoảng trống bằng cụm từ thích hợp: “Cân bằng hóa học là cân bằng ...(1)... vì tại cân bằng phản ứng ...(2)...”""",
         r"""(1) tĩnh; (2) dừng lại.""",
         r"""(1) tính; (2) tiếp tục xảy ra.""",
         r"""(1) động; (2) tiếp tục xảy ra.""",
         r"""(1) động; (2) dừng lại.""",
         3,
         r"""Cân bằng hóa học là cân bằng động vì tại trạng thái cân bằng, phản ứng thuận và nghịch vẫn tiếp tục xảy ra với tốc độ bằng nhau ($v_t = v_n > 0$)."""),

        # Cau 6
        (r"""Liên kết hóa học trong phân tử $\ce{HNO3}$ là liên kết""",
         r"""cộng hóa trị có cực.""",
         r"""hydrogen.""",
         r"""ion.""",
         r"""cộng hóa trị không cực.""",
         1,
         r"""Trong phân tử $\ce{HNO3}$, các liên kết $\ce{N-O}$ và $\ce{O-H}$ đều là liên kết cộng hóa trị có cực do chênh lệch độ âm điện."""),

        # Cau 7
        (r"""Phát biểu nào sau đây đúng khi nói về sự điện li?""",
         r"""Sự điện li là quá trình phân li một chất trong nước thành ion.""",
         r"""Sự điện li là quá trình phân li một chất dưới tác dụng của dòng điện.""",
         r"""Sự điện li quá trình hoà tan một chất vào nước tạo thành dung dịch.""",
         r"""Sự điện li thực chất là quá trình oxi hoá - khử.""",
         1,
         r"""Định nghĩa: Sự điện li là quá trình phân li các chất khi tan trong nước thành các ion."""),

        # Cau 8
        (r"""Cho hình vẽ mô tả thí nghiệm như sau:
\begin{center}
\includegraphics[width=0.38\textwidth]{images_phongdien_ma1020/dai_phun_nuoc_nh3_clean.png}
\end{center}
Hình vẽ mô tả thí nghiệm để chứng minh""",
         r"""tính base của $\ce{NH3}$.""",
         r"""tính tan nhiều trong nước và tính base của $\ce{NH3}$.""",
         r"""tính tan nhiều trong nước của $\ce{NH3}$.""",
         r"""tính khử của $\ce{NH3}$.""",
         2,
         r"""Khí $\ce{NH3}$ tan rất nhiều trong nước tạo áp suất giảm hút nước phun thành tia (tính tan nhiều) và làm phenolphthalein chuyển sang màu hồng/đỏ (tính base)."""),

        # Cau 9
        (r"""Dãy nào sau đây gồm các chất đều thuộc dãy đồng đẳng của methane?""",
         r"""$\ce{C2H2}, \ce{C3H4}, \ce{C4H6}$.""",
         r"""$\ce{CH4}, \ce{C2H2}, \ce{C3H4}$.""",
         r"""$\ce{C2H6}, \ce{C3H8}, \ce{C5H12}$.""",
         r"""$\ce{CH4}, \ce{C2H6}, \ce{C4H14}$.""",
         3,
         r"""Dãy đồng đẳng của methane là các alkane có công thức chung $\ce{C_nH_{2n+2}}$: $\ce{C2H6}, \ce{C3H8}, \ce{C5H12}$."""),

        # Cau 10
        (r"""Tác nhân chính gây ra hiện tượng mưa acid là""",
         r"""$\ce{NH3}, \ce{NO2}$.""",
         r"""$\ce{NO_x}, \ce{SO2}$.""",
         r"""$\ce{CO}, \ce{SO2}$.""",
         r"""$\ce{CO}, \ce{NH3}$.""",
         2,
         r"""Khí $\ce{SO2}$ và $\ce{NO_x}$ trong không khí bị oxi hóa kết hợp với nước mưa tạo acid $\ce{H2SO4}$ và $\ce{HNO3}$ gây mưa acid."""),

        # Cau 11
        (r"""Ngoài đơn chất nitrogen thì nguyên tố nitrogen còn tồn tại dưới dạng nào?""",
         r"""Khoáng vật sodium nitrate ($\ce{NaNO3}$) gọi là diêm tiêu sodium""",
         r"""Có trong protein""",
         r"""Có trong nucleic acid""",
         r"""Có trong diêm tiêu sodium, protein, nucleic acid...""",
         4,
         r"""Trong tự nhiên nitrogen tồn tại ở dạng khoáng vật diêm tiêu ($\ce{NaNO3}$) và là thành phần của protein, nucleic acid..."""),

        # Cau 12
        (r"""Hợp chất hữu cơ là""",
         r"""Hợp chất khó tan trong nước.""",
         r"""Hợp chất của cacbon trừ $\ce{CO}, \ce{CO2}, \ce{H2CO3}$, muối cacbonat kim loại...""",
         r"""Hợp chất của cacbon và một số nguyên tố khác trừ N, Cl, O.""",
         r"""Hợp chất có nhiệt độ sôi cao.""",
         2,
         r"""Định nghĩa: Hợp chất hữu cơ là hợp chất của carbon (trừ $\ce{CO}, \ce{CO2}, \ce{H2CO3}$, muối carbonate, carbide, cyanide...).""")
    ]

    for q_text, op1, op2, op3, op4, ans_idx, sol in questions_p1:
        tex.append(r"""\begin{ex}
""" + q_text + r"""
\choice
""")
        ops = [op1, op2, op3, op4]
        ops_tex = []
        for i, op in enumerate(ops):
            if is_teacher and (i + 1 == ans_idx):
                ops_tex.append(r"{\textbf{\color{hfRed}" + op + r"}}")
            else:
                ops_tex.append(op)
        tex.append(r"{" + ops_tex[0] + r"}" + "\n")
        tex.append(r"{" + ops_tex[1] + r"}" + "\n")
        tex.append(r"{" + ops_tex[2] + r"}" + "\n")
        tex.append(r"{" + ops_tex[3] + r"}" + "\n")
        if is_teacher:
            tex.append(r"""\loigiai{
""" + sol + r"""
}
""")
        else:
            tex.append(r"""\loigiai{}
""")
        tex.append(r"""\end{ex}
""")

    # Part II: Trắc nghiệm đúng sai
    tex.append(r"""
\vspace{6pt}
{\large\bfseries\color{mainGreen} PHẦN II. Câu trắc nghiệm đúng sai (2,0 điểm)}

\textit{Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai. Điểm tối đa của 01 câu hỏi là 1 điểm: 01 ý được 0,1đ; 02 ý được 0,25đ; 03 ý được 0,5đ; 04 ý được 1đ.}
\vspace{4pt}
""")

    # Cau 1 Part II
    tex.append(r"""\begin{ex}
Cho khoảng 1 mL hexane vào bát sứ nhỏ, cẩn thận đưa que đóm đang cháy vào bề mặt chất lỏng, hexane bốc cháy cho ngọn lửa màu vàng.
\choiceTF
""")
    tf1 = [
        ("Đốt hexane trong điều kiện thiếu oxygen sinh ra carbon monooxide.", True),
        ("Không nên đốt hexane trong cốc thuỷ tinh do phản ứng này toả ra lượng nhiệt lớn có thể làm vỡ cốc thuỷ tinh.", True),
        ("Phản ứng trên là phản ứng thu nhiệt.", False),
        ("Hexane thành phần chủ yếu của khí dầu mỏ hóa lỏng (LPG).", False)
    ]
    for stmt, val in tf1:
        t_str = r"\True" if (is_teacher and val) else ""
        f_str = r"\True" if (is_teacher and not val) else ""
        if val:
            tex.append(r"{" + t_str + " " + stmt + "}\n")
        else:
            tex.append(r"{" + f_str + " " + stmt + "}\n")
    if is_teacher:
        tex.append(r"""\loigiai{
$\bullet$ Ý a: \textbf{ĐÚNG}. Cháy không hoàn toàn trong điều kiện thiếu oxygen sinh khí độc $\ce{CO}$ và muội than.\\
$\bullet$ Ý b: \textbf{ĐÚNG}. Phản ứng tỏa lượng nhiệt lớn làm cốc thủy tinh giãn nở không đều có thể bị vỡ nứt.\\
$\bullet$ Ý c: \textbf{SAI}. Phản ứng đốt cháy hydrocarbon luôn là phản ứng tỏa nhiệt mạnh ($\Delta H < 0$).\\
$\bullet$ Ý d: \textbf{SAI}. Khí dầu mỏ hóa lỏng (LPG) có thành phần chính là propane ($\ce{C3H8}$) và butane ($\ce{C4H10}$).
}
""")
    else:
        tex.append(r"""\loigiai{}
""")
    tex.append(r"""\end{ex}
""")

    # Cau 2 Part II
    tex.append(r"""\begin{ex}
Thực hiện thí nghiệm tách cồn (ethanol) ra khỏi hỗn hợp cồn -- nước, được mô phỏng như hình sau:
\begin{center}
\includegraphics[width=0.6\textwidth]{images_phongdien_ma1020/thi_nghiem_chung_cat_clean.png}
\end{center}
\choiceTF
""")
    tf2 = [
        ("Đầu nước vào và nước ra của ống sinh hàn có thể được đặt ngược lại.", False),
        ("Thí nghiệm đã sử dụng phương pháp chưng cất.", True),
        ("Phương pháp tách ethanol và nước dựa trên sự khác nhau về khối lượng riêng.", False),
        ("Độ cồn của sản phẩm sẽ lớn hơn so với độ cồn ban đầu do sản phẩm thu được tinh khiết hơn lẫn ít nước hơn cồn ban đầu.", True)
    ]
    for stmt, val in tf2:
        t_str = r"\True" if (is_teacher and val) else ""
        f_str = r"\True" if (is_teacher and not val) else ""
        if val:
            tex.append(r"{" + t_str + " " + stmt + "}\n")
        else:
            tex.append(r"{" + f_str + " " + stmt + "}\n")
    if is_teacher:
        tex.append(r"""\loigiai{
$\bullet$ Ý a: \textbf{SAI}. Nước vào phải ở đầu thấp và ra ở đầu cao để ống sinh hàn luôn ngập đầy nước làm mát và tránh sốc nhiệt.\\
$\bullet$ Ý b: \textbf{ĐÚNG}. Thí nghiệm dùng nhiệt độ sôi để làm bay hơi và ngưng tụ chất lỏng là phương pháp chưng cất.\\
$\bullet$ Ý c: \textbf{SAI}. Chưng cất dựa trên sự khác nhau về nhiệt độ sôi của các chất (ethanol sôi ở $78{,}3^\circ\text{C}$, nước sôi ở $100^\circ\text{C}$).\\
$\bullet$ Ý d: \textbf{ĐÚNG}. Ethanol sôi ở nhiệt độ thấp hơn nên bay hơi trước, dịch ngưng thu được có hàm lượng ethanol cao hơn hỗn hợp ban đầu.
}
""")
    else:
        tex.append(r"""\loigiai{}
""")
    tex.append(r"""\end{ex}
""")

    # Part III: Trắc nghiệm trả lời ngắn
    tex.append(r"""
\vspace{6pt}
{\large\bfseries\color{mainGreen} PHẦN III: Câu trắc nghiệm yêu cầu trả lời ngắn (2,0 điểm)}

\textit{Thí sinh làm tròn kết quả đến hàng phần mười.}
\vspace{4pt}
""")

    # Cau 1 Part III
    tex.append(r"""\begin{ex}
Một cơ sở nhỏ sản xuất tinh dầu sả từ thân và lá sả tươi. Sau khi đun sôi hỗn hợp sả với nước, người ta thu được một hỗn hợp gồm nước và tinh dầu sả, trong đó tinh dầu sả không tan trong nước, nhẹ hơn nước và còn lẫn một lượng nhỏ tạp chất rắn từ nguyên liệu thực vật. Để tách riêng và tinh chế tinh dầu sả đạt độ tinh khiết cao, cách lựa chọn và ưu tiên thứ tự sử dụng các phương pháp hóa học nào sau đây là hợp lí nhất?
\begin{enumerate}[label=\arabic*.,leftmargin=20pt,itemsep=0pt]
\item Làm khô và tinh chế bằng kết tinh nhằm loại bỏ tạp chất còn sót lại.
\item Chưng cất thường để thu hỗn hợp
\item chiết để tách tinh dầu khỏi nước
\end{enumerate}
""")
    if is_teacher:
        tex.append(r"""\par\textbf{Trả lời:} \textbf{\color{hfRed}2 -- 3 -- 1}
\loigiai{
Trình tự tách và tinh chế: Chưng cất (2) $\to$ Chiết tách lớp tinh dầu khỏi nước (3) $\to$ Làm khô và tinh chế (1). Thứ tự là 2 - 3 - 1.
}
""")
    else:
        tex.append(r"""\par\textit{Trả lời:} \dotfill
\loigiai{}
""")
    tex.append(r"""\end{ex}
""")

    # Cau 2 Part III
    tex.append(r"""\begin{ex}
Cho lần lượt các chất sau $\ce{FeS}, \ce{Fe3O4}, \ce{Fe2O3}, \ce{FeO}, \ce{Fe(OH)2}, \ce{Fe(OH)3}, \ce{FeSO4}, \ce{Fe2(SO4)3}$ tác dụng với dung dịch $\ce{H2SO4}$ đặc nóng. Số phản ứng oxi hóa -- khử là?
""")
    if is_teacher:
        tex.append(r"""\par\textbf{Trả lời:} \textbf{\color{hfRed}5}
\loigiai{
Dung dịch $\ce{H2SO4}$ đặc nóng oxi hóa các chất có nguyên tố Fe hoặc S chưa đạt số oxi hóa cao nhất (+3):
1) $\ce{FeS}$; 2) $\ce{Fe3O4}$; 3) $\ce{FeO}$; 4) $\ce{Fe(OH)2}$; 5) $\ce{FeSO4}$.
Vậy có 5 chất tham gia phản ứng oxi hóa - khử.
}
""")
    else:
        tex.append(r"""\par\textit{Trả lời:} \dotfill
\loigiai{}
""")
    tex.append(r"""\end{ex}
""")

    # Cau 3 Part III
    tex.append(r"""\begin{ex}
Người ta xác định được \% khối lượng các nguyên tố trong vitamin C: $\%C = 40{,}91\%$; $\%H = 4{,}545\%$; $\%O = 54{,}545\%$. Số nguyên tử của C, H, O trong CTPT vitamin C. biết $M_{\text{vitamin C}} = 176$.
""")
    if is_teacher:
        tex.append(r"""\par\textbf{Trả lời:} \textbf{\color{hfRed}6, 8, 6}
\loigiai{
$x = \frac{176 \times 40{,}91\%}{12} = 6$; $y = \frac{176 \times 4{,}545\%}{1} = 8$; $z = \frac{176 \times 54{,}545\%}{16} = 6 \Rightarrow \ce{C6H8O6}$.
Số nguyên tử C, H, O lần lượt là 6, 8, 6.
}
""")
    else:
        tex.append(r"""\par\textit{Trả lời:} \dotfill
\loigiai{}
""")
    tex.append(r"""\end{ex}
""")

    # Cau 4 Part III
    tex.append(r"""\begin{ex}
Hợp chất 3-methylpentane có công thức cấu tạo dưới đây:
\[
\ce{CH3 - CH2 - CH(CH3) - CH2 - CH3}
\]
Khi cho 3-methylpentane tác dụng với khí chlorine thu được số đồng phân cấu tạo là dẫn xuất monochloro tối đa là
""")
    if is_teacher:
        tex.append(r"""\par\textbf{Trả lời:} \textbf{\color{hfRed}4}
\loigiai{
Phân tử có 4 vị trí C không tương đương:
- Thế tại C1 (hoặc C5): 1-chloro-3-methylpentane (1).
- Thế tại C2 (hoặc C4): 2-chloro-3-methylpentane (2).
- Thế tại C3: 3-chloro-3-methylpentane (3).
- Thế tại nhóm $-\ce{CH3}$ nhánh ở C3: 1-chloro-2-ethylbutane (4).
$\Rightarrow$ Có 4 đồng phân cấu tạo monochloro.
}
""")
    else:
        tex.append(r"""\par\textit{Trả lời:} \dotfill
\loigiai{}
""")
    tex.append(r"""\end{ex}
""")

    # Part IV: Tự luận (3,0 điểm)
    tex.append(r"""
\vspace{6pt}
{\large\bfseries\color{mainGreen} PHẦN IV: Tự luận (3,0 điểm)}

\textit{Thí sinh trình bày lời giải chi tiết vào phần giấy làm bài thi.}
\vspace{4pt}
""")

    # Cau 1 Tu luan
    tex.append(r"""\textbf{Câu 1: (0,5 điểm)} Trình bày các luận điểm chính của Thuyết cấu tạo hóa học?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=2mm,boxrule=0.5pt,title=\textbf{Lời giải Câu 1 (Tự luận)}]
1. Trong phân tử hợp chất hữu cơ, các nguyên tử liên kết với nhau theo đúng hóa trị và theo một trật tự nhất định gọi là cấu tạo hóa học. Thay đổi trật tự liên kết sẽ tạo ra chất mới.\\
2. Trong hợp chất hữu cơ, carbon luôn có hóa trị 4. Các nguyên tử carbon liên kết với nguyên tố khác và liên kết trực tiếp với nhau tạo mạch carbon (mạch thẳng, mạch nhánh, mạch vòng).\\
3. Tính chất của chất phụ thuộc vào thành phần phân tử và cấu tạo hóa học, đồng thời phụ thuộc vào sự ảnh hưởng qua lại giữa các nguyên tử.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{2pt}
\multido{}{6}{\noindent\dotfill\par\vspace{2pt}}
""")

    # Cau 2 Tu luan
    tex.append(r"""
\vspace{4pt}
\textbf{Câu 2: (0,5 điểm)} Trình bày các bước gọi tên thay thế của Alkane phân nhánh?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=2mm,boxrule=0.5pt,title=\textbf{Lời giải Câu 2 (Tự luận)}]
$\bullet$ \textbf{Bước 1 (Chọn mạch chính):} Chọn mạch carbon dài nhất và có nhiều nhánh nhất làm mạch chính.\\
$\bullet$ \textbf{Bước 2 (Đánh số mạch chính):} Đánh số C trên mạch chính từ đầu gần nhánh hơn (tổng chỉ số vị trí nhánh nhỏ nhất).\\
$\bullet$ \textbf{Bước 3 (Gọi tên):} Tên alkane = [Vị trí nhánh] - [Tên nhánh] + [Tên alkane mạch chính].
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{2pt}
\multido{}{6}{\noindent\dotfill\par\vspace{2pt}}
""")

    # Cau 3 Tu luan
    tex.append(r"""
\vspace{4pt}
\textbf{Câu 3: (0,5 điểm)} Hoàn thành các phương trình phản ứng sau:
\begin{enumerate}[label=\alph*),leftmargin=25pt,itemsep=2pt]
\item $\ce{Cu(OH)2 + H2SO4(đặc, nóng) ->}$
\item $\ce{Fe + H2SO4(đặc, nóng) ->}$
\end{enumerate}
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=2mm,boxrule=0.5pt,title=\textbf{Lời giải Câu 3 (Tự luận)}]
a) $\ce{Cu(OH)2 + H2SO4(đặc, nóng) -> CuSO4 + 2H2O}$ (phản ứng acid - base).\\
b) $\ce{2Fe + 6H2SO4(đặc, nóng) ->[t^\circ] Fe2(SO4)3 + 3SO2 ^ + 6H2O}$.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{2pt}
\multido{}{5}{\noindent\dotfill\par\vspace{2pt}}
""")

    # Cau 4 Tu luan
    tex.append(r"""
\vspace{4pt}
\textbf{Câu 4: (0,5 điểm)} Chuẩn độ acid- base là phương pháp được sử dụng để xác định nồng độ acid hoặc base chưa biết nồng độ bằng dung dịch acid hoặc dung dịch base đã biết chính xác nồng độ. Khi đó dung dịch đã biết chính xác nồng độ gọi là dung dịch chuẩn. Khi chuẩn độ, người ta thêm từ từ dung dịch đựng trong burette vào dung dịch đựng trong bình tam giác. Tại điểm mà hai chất tác dụng vừa đủ với nhau gọi là điểm tương đương. Người ta thường dựa vào đặc điểm nào để nhận biết điểm tương đương trong phép chuẩn độ trên? Vì sao?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=2mm,boxrule=0.5pt,title=\textbf{Lời giải Câu 4 (Tự luận)}]
$\bullet$ \textbf{Đặc điểm nhận biết:} Dựa vào sự đổi màu đột ngột của chất chỉ thị màu acid - base (như phenolphthalein hoặc methyl da cam).\\
$\bullet$ \textbf{Vì sao:} Tại điểm tương đương, phản ứng vừa đủ xảy ra bước nhảy pH rất nhanh và đột ngột, làm cho chất chỉ thị biến đổi cấu trúc phân tử và đổi màu rõ rệt, giúp nhận biết thời điểm dừng chuẩn độ.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{2pt}
\multido{}{7}{\noindent\dotfill\par\vspace{2pt}}
""")

    # Cau 5 Tu luan
    tex.append(r"""
\vspace{4pt}
\textbf{Câu 5: (1,0 điểm)} Một nhà máy sản xuất sulfuric acid áp dụng phương pháp tiếp xúc với các giai đoạn chính sau:
\begin{itemize}[leftmargin=20pt,itemsep=0pt]
\item Đốt Sulfur hoặc quặng pyrite tạo $\ce{SO2}$
\item Oxi hóa $\ce{SO2}$ thành $\ce{SO3}$ trên xúc tác $\ce{V2O5}$
\item Hấp thụ $\ce{SO3}$ để thu sulfuric acid
\end{itemize}
Trong quá trình vận hành, nhà máy phải đồng thời đảm bảo hiệu suất cao, tiết kiệm năng lượng và giảm thiểu ô nhiễm môi trường.

1. Xét phản ứng oxi hóa: $\ce{2SO2(g) + O2(g) <=> 2SO3(g)},\quad (\Delta H < 0)$
\begin{enumerate}[label=\alph*),leftmargin=20pt,itemsep=0pt]
\item Cho biết phản ứng trên là thu nhiệt hay tỏa nhiệt.
\item Giải thích vì sao trong thực tế không tiến hành phản ứng ở nhiệt độ quá cao, mặc dù tốc độ phản ứng sẽ tăng.
\end{enumerate}
2. Để tăng hiệu suất tổng hợp $\ce{SO3}$ người ta phải điều chỉnh các yếu tố: Nhiệt độ, Áp suất, Nồng độ $\ce{SO2}$ và $\ce{O2}$ như thế nào?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=2mm,boxrule=0.5pt,title=\textbf{Lời giải Câu 5 (Tự luận)}]
\textbf{1.} \\
a) Do $\Delta H < 0$ nên phản ứng trên là \textbf{phản ứng tỏa nhiệt}.\\
b) Theo nguyên lí Le Chatelier, vì phản ứng thuận tỏa nhiệt nên nếu tăng nhiệt độ lên quá cao, cân bằng sẽ dịch chuyển theo chiều nghịch (chiều thu nhiệt) làm giảm hiệu suất tạo $\ce{SO3}$. Ngoài ra nhiệt độ quá cao làm giảm tuổi thọ của chất xúc tác $\ce{V2O5}$ và tốn kém năng lượng. (Thực tế chọn $\approx 450^\circ\text{C}$).\\[4pt]
\textbf{2. Điều chỉnh các yếu tố để tăng hiệu suất tổng hợp $\ce{SO3}$:}\\
$\bullet$ \textbf{Nhiệt độ:} Giảm nhiệt độ (chọn nhiệt độ tối ưu khoảng $450^\circ\text{C}$) để cân bằng dịch chuyển sang chiều thuận (tỏa nhiệt).\\
$\bullet$ \textbf{Áp suất:} Tăng áp suất, vì phản ứng thuận làm giảm số mol khí (từ 3 mol khí còn 2 mol khí) nên tăng áp suất cân bằng dịch chuyển theo chiều thuận.\\
$\bullet$ \textbf{Nồng độ $\ce{SO2}$ và $\ce{O2}$:} Tăng nồng độ các chất phản ứng (dùng dư không khí / dư $\ce{O2}$) để cân bằng dịch chuyển sang chiều thuận; đồng thời liên tục tách $\ce{SO3}$ ra khỏi hệ (hấp thụ vào $\ce{H2SO4}$ đặc).
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{2pt}
\multido{}{10}{\noindent\dotfill\par\vspace{2pt}}
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
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8', errors='replace', cwd=out_dir)
    
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
    pdf_file = os.path.join(out_dir, base_name + '.pdf')
    if os.path.exists(pdf_file):
        print(f"[OK] Đã biên dịch xong PDF: {pdf_file}")
    else:
        print(f"[CẢNH BÁO] Không tìm thấy file PDF: {pdf_file}")

def main():
    tex_hs = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_CK1_Hoa_11_PhongDien_Ma1020_HocSinh.tex")
    tex_gv = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_CK1_Hoa_11_PhongDien_Ma1020_GiaoVien.tex")

    print("Đang tạo mã nguồn LaTeX Bản Học sinh THPT Phong Điền...")
    with open(tex_hs, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))

    print("Đang tạo mã nguồn LaTeX Bản Giáo viên THPT Phong Điền...")
    with open(tex_gv, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))

    print("\nĐang biên dịch PDF Bản Học sinh THPT Phong Điền...")
    compile_pdf(tex_hs)

    print("\nĐang biên dịch PDF Bản Giáo viên THPT Phong Điền...")
    compile_pdf(tex_gv)

    print("\n[HOÀN TẤT] Toàn bộ file .tex và .pdf đã được tạo và biên dịch thành công!")

if __name__ == "__main__":
    main()
