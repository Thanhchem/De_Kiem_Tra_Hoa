# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Cuối Học Kì I Hóa Học 11 - THPT Nguyễn Chí Thanh - Huế
Mã đề: 1101 - Năm học 2025 - 2026
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_CK1_Hoa_11_NguyenChiThanh_Ma1101_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_CK1_Hoa_11_NguyenChiThanh_Ma1101_GiaoVien.pdf (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
\textbf{TRƯỜNG THPT NGUYỄN CHÍ THANH}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 04 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}ĐỀ KIỂM TRA CUỐI HỌC KÌ 1}\\
\textbf{NĂM HỌC 2025 -- 2026}\\
\textbf{MÔN: HÓA HỌC - LỚP 11}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{MÃ ĐỀ: 1101}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ và tên:} \dotfill \textit{Số báo danh:} \dotfill

\vspace{3pt}
\noindent
\textit{\footnotesize Cho biết nguyên tử khối của các nguyên tố: $\ce{Cu} = 64; \ce{H} = 1; \ce{O} = 16; \ce{N} = 14; \ce{Fe} = 56; \ce{Al} = 27; \ce{S} = 32; \ce{C} = 12; \ce{Ba} = 137$.}

\vspace{3pt}
\hrule color{hfGreen} height 0.5pt
""")

    if is_teacher:
        tex.append(r"""
\vspace{6pt}
\begin{center}
{\large\bfseries\color{hfRed} BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 1101)}
\end{center}

\vspace{4pt}
{\bfseries\color{mainGreen} 1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (4,0 ĐIỂM)}

\vspace{2pt}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} \\ \hline
\textbf{Đ/A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} \\ \hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} \\ \hline
\textbf{Đ/A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} \\ \hline
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
\textbf{Câu 1} & Chuẩn độ dung dịch $\ce{H2SO4}$ bằng dung dịch $\ce{NaOH}$ & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{\color{hfRed}ĐÚNG} \\ \hline
\textbf{Câu 2} & Đốt cháy hoàn toàn hợp chất hữu cơ X ($\ce{C4H10}$) & \textbf{SAI} & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{SAI} \\ \hline
\end{tabular}
\end{center}

\vspace{4pt}
{\bfseries\color{mainGreen} 3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (1,0 ĐIỂM)}

\vspace{2pt}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu hỏi} & \textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} \\ \hline
\textbf{Đáp án} & \textbf{\color{hfRed}6,72} & \textbf{\color{hfRed}3} & \textbf{\color{hfRed}3} & \textbf{\color{hfRed}2} \\ \hline
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
\textbf{Câu 1} & 0,5 đ & • Butane: $\ce{CH3-CH2-CH2-CH3}$, khung mạch dích dắc 4 đỉnh.\\
& & • Isobutane: $\ce{CH3-CH(CH3)-CH3}$, khung chữ Y 3 nhánh. \\ \hline
\textbf{Câu 2} & 1,0 đ & a) Peak ion phân tử $[M^+]$ tại $m/z = 194 \Rightarrow M = 194\text{ g/mol}$. (0,25 đ)\\
& & b) $\%O = 16{,}50\% \Rightarrow \text{Số C} = 8, \text{Số H} = 10, \text{Số N} = 4, \text{Số O} = 2 \Rightarrow \ce{C8H10N4O2}$. (0,75 đ) \\ \hline
\textbf{Câu 3} & 1,5 đ & a) 4 phương trình: (1) Đốt quặng $\ce{FeS2}$; (2) Oxi hóa $\ce{SO2}$; (3) Hấp thụ tạo oleum; (4) Pha loãng oleum. (0,5 đ)\\
& & b) Trong 50,7 g X: $n_{\ce{S}} = 0{,}6\text{ mol} \Rightarrow n = 3$ (oleum $\ce{H2SO4.3SO3}$).\\
& & Trong 1000 g dd 98\%: $n_{\ce{H2SO4}} = 10\text{ mol}; n_{\ce{H2O}} = \frac{10}{9}\text{ mol} \Rightarrow n_{\ce{SO3(tổng)}} = \frac{310}{9}\text{ mol}$.\\
& & $n_{\ce{FeS2}} = \frac{155}{9}\text{ mol} \Rightarrow m = \frac{155}{9} \times 120 = 2066{,}67\text{ g} \approx 2{,}067\text{ kg}$ (hoặc 2,07 kg). (1,0 đ) \\ \hline
\end{tabular}
\end{center}

\vspace{4pt}
\hrule color{hfGreen} height 0.5pt
""")

    # Part I: Trắc nghiệm
    tex.append(r"""
\vspace{6pt}
{\large\bfseries\color{mainGreen} A. PHẦN TRẮC NGHIỆM}

\vspace{3pt}
{\bfseries\color{mainGreen} PHẦN 1: TRẮC NGHIỆM NHIỀU LỰA CHỌN (16 câu -- 4 điểm)}
\textit{Thí sinh trả lời câu hỏi từ câu 1 đến câu 16, mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{4pt}
""")

    questions_p1 = [
        # Cau 1
        (r"""Cho cân bằng hóa học: $\ce{N2(g) + 3H2(g) <=> 2NH3(g)}\quad \Delta_r H = -92\text{ kJ}$
Cân bằng trên chuyển dịch theo chiều thuận khi""",
         r"""tăng nhiệt độ của hệ phản ứng.""",
         r"""thêm chất xúc tác vào hệ phản ứng.""",
         r"""tăng nồng độ $\ce{NH3}$.""",
         r"""tăng áp suất của hệ phản ứng.""",
         4,
         r"""Phản ứng thuận làm giảm số mol khí (từ 4 mol thành 2 mol), nên khi tăng áp suất, cân bằng chuyển dịch theo chiều thuận."""),

        # Cau 2
        (r"""Hỗn hợp X gồm các alkane: pentane ($t_s = 36{,}1^\circ\text{C}$), heptane ($t_s = 98{,}4^\circ\text{C}$), octane ($t_s = 125{,}7^\circ\text{C}$) và nonane ($t_s = 150{,}8^\circ\text{C}$). Có thể tách riêng các chất đó một cách thuận lợi bằng phương pháp nào sau đây?""",
         r"""Chưng cất.""",
         r"""Chiết.""",
         r"""Kết tinh.""",
         r"""Sắc kí.""",
         1,
         r"""Các alkane là chất lỏng tan vào nhau có nhiệt độ sôi khác nhau rõ rệt nên tách bằng phương pháp chưng cất phân đoạn."""),

        # Cau 3
        (r"""Để tạo độ xốp cho một số loại bánh, có thể dùng chất nào sau đây?""",
         r"""\ce{NaCl}.""",
         r"""\ce{(NH4)3PO4}.""",
         r"""\ce{CaCO3}.""",
         r"""\ce{NH4HCO3}.""",
         4,
         r"""$\ce{NH4HCO3}$ (bột nở) bị nhiệt phân hủy khi nướng: $\ce{NH4HCO3 ->[t^\circ] NH3 ^ + CO2 ^ + H2O ^}$ tạo các bọt khí làm xốp bánh."""),

        # Cau 4
        (r"""Sulfur là chất oxi hóa trong phản ứng nào sau đây?""",
         r"""$\ce{S + 3F2 -> SF6}$""",
         r"""$\ce{S + O2 ->[t^\circ] SO2}$""",
         r"""$\ce{S + Fe ->[t^\circ] FeS}$""",
         r"""$\ce{S + 2H2SO4 ->[t^\circ] 3SO2 + 2H2O}$""",
         3,
         r"""Trong phản ứng $\ce{S + Fe -> FeS}$, số oxi hóa của S giảm từ 0 xuống -2 nên S là chất oxi hóa."""),

        # Cau 5
        (r"""Nhóm chức là""",
         r"""Là một nguyên tử bất kì trong phân tử chất hữu cơ""",
         r"""Là một nguyên tử (hoặc nhóm nguyên tử) gây ra những phản ứng hoá học đặc trưng của phân tử hợp chất hữu cơ""",
         r"""Là một nhóm nguyên tử khác biệt trong chất hữu cơ""",
         r"""Là một nhóm nguyên tử có cấu trúc không gian đặc biệt mà trong đó các nguyên tử liên kết với nhau không theo quy tắc hoá trị nào""",
         2,
         r"""Định nghĩa chuẩn: Nhóm chức là nguyên tử hoặc nhóm nguyên tử gây ra những phản ứng hoá học đặc trưng của hợp chất hữu cơ."""),

        # Cau 6
        (r"""Chất nào sau đây là chất điện li?""",
         r"""saccharose""",
         r"""sulfuric acid""",
         r"""glucose""",
         r"""ethyl alcohol""",
         2,
         r"""Sulfuric acid ($\ce{H2SO4}$) phân li trong nước tạo thành các ion nên là chất điện li mạnh. Các chất còn lại là chất không điện li."""),

        # Cau 7
        (r"""Hợp chất X có công thức khung phân tử như sau:
\begin{center}
\includegraphics[width=0.25\textwidth]{images_nguyenchithanh_ma1101/khung_c7_clean.png}
\end{center}
Tổng số nguyên tử hydrogen trong X là""",
         r"""10""",
         r"""12""",
         r"""6""",
         r"""8""",
         1,
         r"""Công thức của X là 2-methylbut-1-ene: $\ce{CH2=C(CH3)-CH2-CH3}$ ($\ce{C5H10}$), có 10 nguyên tử H."""),

        # Cau 8
        (r"""Cho iron (III) oxide tác dụng với nitric acid loãng thì sản phẩm thu được là""",
         r"""$\ce{Fe(NO3)3}, \ce{NO2}$ và $\ce{H2O}$.""",
         r"""$\ce{Fe(NO3)3}$ và $\ce{H2O}$.""",
         r"""$\ce{Fe(NO3)3}, \ce{NO}$ và $\ce{H2O}$.""",
         r"""$\ce{Fe(NO3)3}, \ce{N2}$ và $\ce{H2O}$.""",
         2,
         r"""Trong $\ce{Fe2O3}$, Fe đã có số oxi hóa cao nhất (+3) nên chỉ xảy ra phản ứng trao đổi: $\ce{Fe2O3 + 6HNO3 -> 2Fe(NO3)3 + 3H2O}$."""),

        # Cau 9
        (r"""Hợp chất X tan trong nước tạo dung dịch không màu. Dung dịch này không tạo kết tủa với dung dịch $\ce{BaCl2}$, khi phản ứng với $\ce{NaOH}$ tạo ra khí mùi khai, khi phản ứng với dung dịch $\ce{HCl}$ tạo ra khí làm đục nước vôi trong và làm mất màu dung dịch nước bromine. Chất X là""",
         r"""\ce{(NH4)2SO3}""",
         r"""\ce{(NH4)2CO3}""",
         r"""\ce{NH4HCO3}""",
         r"""\ce{NH4HSO3}""",
         4,
         r"""Tác dụng $\ce{NaOH}$ tạo khí mùi khai $\Rightarrow$ có cation $\ce{NH4+}$. Không kết tủa với $\ce{BaCl2}$ loại sulfite và carbonate trung hòa. Tác dụng $\ce{HCl}$ tạo khí $\ce{SO2}$ làm đục nước vôi trong và làm mất màu nước bromine $\Rightarrow$ gốc $\ce{HSO3-}$, chất X là $\ce{NH4HSO3}$."""),

        # Cau 10
        (r"""Chất nào dưới đây là đồng đẳng của ethanol $\ce{CH3CH2OH}$?""",
         r"""$\ce{CH3OCH2CH3}$.""",
         r"""$\ce{CH2=CHCH2OH}$.""",
         r"""$\ce{CH3OCH3}$""",
         r"""$\ce{CH3CH2CH2CH2OH}$.""",
         4,
         r"""$\ce{CH3CH2CH2CH2OH}$ (butan-1-ol) là alcohol no, đơn chức, mạch hở thuộc cùng dãy đồng đẳng với ethanol."""),

        # Cau 11
        (r"""Một số cơ sở sản xuất thuốc Bắc thường đốt một chất bột rắn X màu vàng (là một đơn chất) để tạo ra khí Y nhằm mục đích tẩy trắng, chống mốc. Tuy nhiên, theo các nhà khoa học thì khí Y có ảnh hưởng không tốt đến cơ quan nội tạng và khí Y cũng là một trong những nguyên nhân gây ra “mưa acid”. Chất rắn X là""",
         r"""iodine.""",
         r"""carbon.""",
         r"""sulfur.""",
         r"""phosphorus.""",
         3,
         r"""Chất rắn màu vàng là sulfur (S), khi đốt tạo khí $\ce{SO2}$ diệt khuẩn chống mốc nhưng rất độc và gây mưa acid."""),

        # Cau 12
        (r"""Dụng cụ này có thể được sử dụng để tách riêng các chất trong hỗn hợp nào dưới đây?
\begin{center}
\includegraphics[width=0.15\textwidth]{images_nguyenchithanh_ma1101/pheu_chiet_c12_clean.png}
\end{center}""",
         r"""Xăng và dầu ăn.""",
         r"""Nước và rượu.""",
         r"""Nước và muối ăn.""",
         r"""Dầu ăn và nước.""",
         4,
         r"""Phễu chiết dùng để chiết tách hai chất lỏng không hòa tan vào nhau tạo 2 lớp riêng biệt: Dầu ăn và nước."""),

        # Cau 13
        (r"""Chất nào sau đây không thuộc loại chất hữu cơ?""",
         r"""\ce{CH3COONa}""",
         r"""\ce{Al4C3}""",
         r"""\ce{CH3Cl}""",
         r"""\ce{CCl4}""",
         2,
         r"""$\ce{Al4C3}$ là muối carbide kim loại, thuộc nhóm các hợp chất vô cơ của carbon."""),

        # Cau 14
        (r"""Khí thải công nghiệp, khí thải động cơ (ô tô, xe máy ...) là nguyên nhân chủ yếu gây mưa acid. Dãy nào dưới đây gồm các chất có trong các loại khí thải này và là nguyên nhân trực tiếp gây mưa acid?""",
         r"""$\ce{SO2}, \ce{CO}, \ce{NO}$.""",
         r"""$\ce{SO2}, \ce{CO}, \ce{NO2}$.""",
         r"""$\ce{NO}, \ce{NO2}, \ce{SO2}$.""",
         r"""$\ce{NO2}, \ce{CO2}, \ce{CO}$.""",
         3,
         r"""Các oxide của nitrogen ($\ce{NO}, \ce{NO2}$) và sulfur dioxide ($\ce{SO2}$) là tác nhân trực tiếp gây mưa acid."""),

        # Cau 15
        (r"""Khi cho saccharose ($\ce{C12H22O11}$) tác dụng với dung dịch $\ce{H2SO4}$ đặc trong cốc thủy tinh thấy có bọt khí đẩy carbon trào lên khỏi cốc. Thí nghiệm trên chứng minh được tính chất nào sau đây của $\ce{H2SO4}$ đặc?""",
         r"""chỉ có tính oxi hóa mạnh.""",
         r"""tính háo nước và tính oxi hóa mạnh.""",
         r"""chỉ có tính háo nước.""",
         r"""tính háo nước và tính khử mạnh.""",
         2,
         r"""$\ce{H2SO4}$ đặc hút nước hóa than đường (tính háo nước), sau đó oxi hóa C thành $\ce{CO2}$ và $\ce{SO2}$ đẩy bọt khí trào lên (tính oxi hóa mạnh)."""),

        # Cau 16
        (r"""Kim loại nào sau đây không tác dụng với dung dịch acid $\ce{H2SO4}$ loãng nguội?""",
         r"""Na""",
         r"""Zn""",
         r"""Cu""",
         r"""Al""",
         3,
         r"""Đồng (Cu) đứng sau hydrogen trong dãy hoạt động hóa học nên không tác dụng với $\ce{H2SO4}$ loãng.""")
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
{\large\bfseries\color{mainGreen} PHẦN 2: TRẮC NGHIỆM ĐÚNG -- SAI (2 câu -- 2 điểm)}

\textit{Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{4pt}
""")

    # Cau 1 Part II
    tex.append(r"""\begin{ex}
Pha 1,0 L dung dịch sulfuric acid. Sau đó lấy 10,0 mL dung dịch acid cho vào bình tam giác, thêm vài giọt dung dịch phenolphathalein. Nhỏ từ từ dung dịch $\ce{NaOH}\ 0{,}10\text{ M}$ chứa trên burette vào bình tam giác đến khi dung dịch xuất hiện màu hồng nhạt, đọc thể tích $\ce{NaOH}$ đã dùng trên burette. Lặp lại thí nghiệm 3 lần, các giá trị thể tích $\ce{NaOH}$ đọc được lần lượt là 17,9 mL; 18,0 mL; 18,1 mL.
\choiceTF
""")
    tf1 = [
        ("Khi chuẩn độ xảy ra phản ứng hoá học $\\ce{H2SO4 + 2NaOH -> Na2SO4 + 2H2O}$.", True),
        ("pH của dung dịch acid gần bằng 1,05.", True),
        ("Để pha 1 lít dung dịch sulfuric acid trên cần hoà 19,6 ml (làm tròn đến phần chục) dung dịch $\\ce{H2SO4}$ 98\\% (d=1,84g/ml) vào 980,4 ml nước.", False),
        ("Phương pháp trong thí nghiệm trên là phương pháp chuẩn độ acid-base.", True)
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
$\bullet$ Thể tích $\ce{NaOH}$ trung bình: $\bar{V} = \frac{17{,}9 + 18{,}0 + 18{,}1}{3} = 18{,}0\text{ mL} \Rightarrow n_{\ce{NaOH}} = 0{,}0018\text{ mol}$.\\
$\bullet$ Ý a: \textbf{ĐÚNG}. $\ce{H2SO4 + 2NaOH -> Na2SO4 + 2H2O}$.\\
$\bullet$ Ý b: \textbf{ĐÚNG}. $n_{\ce{H2SO4(10 mL)}} = 0{,}0009\text{ mol} \Rightarrow C_M = 0{,}09\text{ M} \Rightarrow \text{pH} = -\log(0{,}09) \approx 1{,}046 \approx 1{,}05$.\\
$\bullet$ Ý c: \textbf{SAI}. Trong 1 L dung dịch có $0{,}09\text{ mol } \ce{H2SO4} \Rightarrow m_{\text{dd}} = \frac{0{,}09 \times 98}{0{,}98} = 9{,}0\text{ g} \Rightarrow V = \frac{9{,}0}{1{,}84} \approx 4{,}89\text{ mL}$ (không phải 19,6 mL).\\
$\bullet$ Ý d: \textbf{ĐÚNG}. Đây là phương pháp chuẩn độ acid - base.
}
""")
    else:
        tex.append(r"""\loigiai{}
""")
    tex.append(r"""\end{ex}
""")

    # Cau 2 Part II
    tex.append(r"""\begin{ex}
Đốt cháy hoàn toàn 7,25 gam hợp chất hữu cơ X thu được 11,2 lít $\ce{CO2}$ (đktc) và 11,25 gam $\ce{H2O}$. Biết tỉ khối của X so với khí nitrogen bằng 29/14.
\choiceTF
""")
    tf2 = [
        ("Phần trăm khối lượng của nguyên tố hydrogen trong X xấp xỉ 8,62\\%", False),
        ("X có chứa 4 nguyên tử carbon.", True),
        ("Phân tử khối của X bằng 29.", False),
        ("Phân tử X chứa 3 loại nguyên tố là C, H, O.", False)
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
$n_{\ce{CO2}} = 0{,}5\text{ mol} \Rightarrow m_{\ce{C}} = 6{,}0\text{ g}$; $n_{\ce{H2O}} = 0{,}625\text{ mol} \Rightarrow m_{\ce{H}} = 1{,}25\text{ g}$.\\
$m_{\ce{C}} + m_{\ce{H}} = 7{,}25\text{ g} = m_{\ce{X}} \Rightarrow$ X chỉ chứa C và H (không có O).\\
$M_X = 28 \times \frac{29}{14} = 58\text{ g/mol} \Rightarrow \ce{C4H10}$.\\
$\bullet$ Ý a: \textbf{SAI}. $\%H = \frac{1{,}25}{7{,}25} \times 100\% \approx 17{,}24\%$ (không phải 8,62\%).\\
$\bullet$ Ý b: \textbf{ĐÚNG}. X có 4 nguyên tử C ($\ce{C4H10}$).\\
$\bullet$ Ý c: \textbf{SAI}. Phân tử khối của X là 58.\\
$\bullet$ Ý d: \textbf{SAI}. X chỉ chứa 2 nguyên tố C và H.
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
{\large\bfseries\color{mainGreen} PHẦN 3: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (4 câu - 1 điểm)}

\textit{Thí sinh trả lời từ câu 1 đến câu 4.}
\vspace{4pt}
""")

    # Cau 1 Part III
    tex.append(r"""\begin{ex}
Hoà tan 5,4 gam aluminium (Al) bằng acid $\ce{H2SO4}$ loãng dư sinh ra V lít khí (ở đktc). Tính V.
""")
    if is_teacher:
        tex.append(r"""\par\textbf{Trả lời:} \textbf{\color{hfRed}6,72}
\loigiai{
$n_{\ce{Al}} = \frac{5{,}4}{27} = 0{,}2\text{ mol} \Rightarrow n_{\ce{H2}} = \frac{3}{2} n_{\ce{Al}} = 0{,}3\text{ mol} \Rightarrow V = 0{,}3 \times 22{,}4 = 6{,}72\text{ lít}$.
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
Khi cho dung dịch $\ce{HNO3}$ đặc, nóng lần lượt tác dụng với các chất sau: $\ce{Cu}, \ce{Fe2O3}, \ce{C}$, dung dịch $\ce{Ba(OH)2}$, dung dịch $\ce{Na2SO3}$. Số phản ứng mà $\ce{HNO3}$ đóng vai trò là chất oxi hóa là bao nhiêu?
""")
    if is_teacher:
        tex.append(r"""\par\textbf{Trả lời:} \textbf{\color{hfRed}3}
\loigiai{
$\ce{HNO3}$ đóng vai trò là chất oxi hóa khi tác dụng với các chất khử: $\ce{Cu}, \ce{C}, \ce{Na2SO3}$. Vậy có 3 phản ứng.
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
Cho các cân bằng hoá học:
\begin{align*}
(1) &\quad \ce{N2(g) + 3H2(g) <=> 2NH3(g)} \\
(2) &\quad \ce{H2(g) + I2(g) <=> 2HI(g)} \\
(3) &\quad \ce{2SO2(g) + O2(g) <=> 2SO3(g)} \\
(4) &\quad \ce{2NO2(g) <=> N2O4(g)}
\end{align*}
Khi thay đổi áp suất có bao nhiêu cân bằng hóa học bị chuyển dịch?
""")
    if is_teacher:
        tex.append(r"""\par\textbf{Trả lời:} \textbf{\color{hfRed}3}
\loigiai{
Cân bằng bị chuyển dịch khi thay đổi áp suất có $\Delta n_{\text{khí}} \neq 0$: (1), (3), (4). Cân bằng (2) có $\Delta n = 0$ nên không bị chuyển dịch. Vậy có 3 cân bằng.
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
Ứng với CTPT $\ce{C3H8O}$ có bao nhiêu đồng phân cấu tạo thuộc loại alcohol?
""")
    if is_teacher:
        tex.append(r"""\par\textbf{Trả lời:} \textbf{\color{hfRed}2}
\loigiai{
Có 2 đồng phân alcohol: $\ce{CH3-CH2-CH2-OH}$ (propan-1-ol) và $\ce{CH3-CH(OH)-CH3}$ (propan-2-ol).
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
{\large\bfseries\color{mainGreen} B. PHẦN TỰ LUẬN (3 câu -- 3 điểm)}

\textit{Thí sinh trình bày lời giải chi tiết vào phần giấy làm bài thi.}
\vspace{4pt}
""")

    # Cau 1 Tu luan
    tex.append(r"""\textbf{Câu 1: (0,5 điểm)} Viết công thức cấu tạo thu gọn và công thức khung phân tử các đồng phân cấu tạo ứng với công thức phân tử $\ce{C4H10}$.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=2mm,boxrule=0.5pt,title=\textbf{Lời giải Câu 1 (Tự luận)}]
1. \textbf{Butane (mạch thẳng):}\\
$\bullet$ CTCT thu gọn: $\ce{CH3-CH2-CH2-CH3}$\\
$\bullet$ Công thức khung phân tử: Mạch dích dắc 4 đỉnh.\\
2. \textbf{2-Methylpropane / Isobutane (mạch phân nhánh):}\\
$\bullet$ CTCT thu gọn: $\ce{CH3-CH(CH3)-CH3}$\\
$\bullet$ Công thức khung phân tử: Hình chữ Y 3 nhánh.
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
\textbf{Câu 2: (1,0 điểm)} Caffeine là chất kích thích tự nhiên được tìm thấy trong cây trà, cà phê và ca cao. Chúng hoạt động bằng cách kích thích não và hệ thần kinh trung ương, giúp con người tỉnh táo và ngăn ngừa sự mệt mỏi. Thành phần phần trăm về khối lượng của các nguyên tố trong phân tử caffeine như sau: 49,48\% carbon; 5,15\% hydrogen; 28,87\% nitrogen còn lại là oxygen. Phổ MS của caffeine được cho như hình dưới đây:

\begin{center}
\includegraphics[width=0.7\textwidth]{images_nguyenchithanh_ma1101/ms_caffeine_clean.png}
\end{center}

a. Dựa vào phổ khối lượng phân tử (MS) hãy xác định phân tử khối của caffeine?\\
b. Xác định công thức phân tử của caffeine.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=2mm,boxrule=0.5pt,title=\textbf{Lời giải Câu 2 (Tự luận)}]
\textbf{a. Xác định phân tử khối:}\\
Trên phổ MS, peak ion phân tử $[M^+]$ có $m/z = 194 \Rightarrow$ Phân tử khối của caffeine là $M = 194\text{ g/mol}$.\\[4pt]
\textbf{b. Xác định công thức phân tử:}\\
$\bullet\ \%O = 100\% - 49{,}48\% - 5{,}15\% - 28{,}87\% = 16{,}50\%$.\\
$\bullet$ Đặt CTPT caffeine là $\ce{C}_x\ce{H}_y\ce{N}_z\ce{O}_t$:\\
$x = \frac{194 \times 49{,}48\%}{12} = 8{,}00 \Rightarrow x = 8$\\
$y = \frac{194 \times 5{,}15\%}{1} = 9{,}99 \Rightarrow y = 10$\\
$z = \frac{194 \times 28{,}87\%}{14} = 4{,}00 \Rightarrow z = 4$\\
$t = \frac{194 \times 16{,}50\%}{16} = 2{,}00 \Rightarrow t = 2$\\
$\Rightarrow$ Công thức phân tử của caffeine là $\ce{C8H10N4O2}$.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{2pt}
\multido{}{8}{\noindent\dotfill\par\vspace{2pt}}
""")

    # Cau 3 Tu luan
    tex.append(r"""
\vspace{4pt}
\textbf{Câu 3: (1,5 điểm)} Sulfuric acid là một trong những hoá chất quan trọng nhất được sử dụng trong công nghiệp; được sản xuất hàng trăm triệu tấn mỗi năm, chiếm nhiều nhất trong ngành công nghiệp hoá chất. Sulfuric acid được sản xuất từ quặng pyrite theo sơ đồ sau:
\[
\ce{FeS2 ->[(1)] SO2 ->[(2)] SO3 ->[(3)] oleum ->[(4)] H2SO4}
\]
a. Hoàn thành sơ đồ phản ứng trên, ghi rõ điều kiện phản ứng nếu có.\\
b. Hấp thụ toàn bộ lượng $\ce{SO3}$ điều chế được từ $m\text{ kg } \ce{FeS2}$ vào 1000 gam dung dịch $\ce{H2SO4}\ 98\%$ thu được oleum X. Hoà tan 50,7 gam X vào nước được dung dịch Y. Cho dung dịch $\ce{BaCl2}$ dư vào Y thu được 139,8 gam kết tủa. Giả sử các phản ứng xảy ra hoàn toàn. Tìm $m$.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=2mm,boxrule=0.5pt,title=\textbf{Lời giải Câu 3 (Tự luận)}]
\textbf{a. Hoàn thành sơ đồ phản ứng:}\\
(1) $\ce{4FeS2 + 11O2 ->[t^\circ] 2Fe2O3 + 8SO2 ^}$\\
(2) $\ce{2SO2 + O2 <=>[\ce{V2O5}, 450^\circ\text{C}] 2SO3}$\\
(3) $\ce{H2SO4 + nSO3 -> H2SO4.nSO3}\text{ (oleum)}$\\
(4) $\ce{H2SO4.nSO3 + nH2O -> (n+1) H2SO4}$\\[4pt]
\textbf{b. Tìm giá trị $m$:}\\
$\bullet$ Trong 50,7 g oleum X ($\ce{H2SO4.nSO3}$):\\
$n_{\ce{BaSO4}} = \frac{139{,}8}{233} = 0{,}6\text{ mol} \Rightarrow n_{\ce{S}} = 0{,}6\text{ mol}$.\\
Khối lượng cho 1 mol S: $\frac{98 + 80n}{n + 1} = \frac{50{,}7}{0{,}6} = 84{,}5 \Rightarrow 4{,}5n = 13{,}5 \Rightarrow n = 3$ (oleum là $\ce{H2SO4.3SO3}$).\\
$\bullet$ Trong 1000 g dung dịch $\ce{H2SO4}\ 98\%$:\\
$n_{\ce{H2SO4}} = \frac{980}{98} = 10\text{ mol}$; $n_{\ce{H2O}} = \frac{20}{18} = \frac{10}{9}\text{ mol}$.\\
$\ce{SO3 + H2O -> H2SO4} \Rightarrow n_{\ce{SO3(pư H2O)}} = \frac{10}{9}\text{ mol}$.\\
Tổng $\ce{H2SO4} = 10 + \frac{10}{9} = \frac{100}{9}\text{ mol}$.\\
$\ce{H2SO4 + 3SO3 -> H2SO4.3SO3} \Rightarrow n_{\ce{SO3(tạo oleum)}} = 3 \times \frac{100}{9} = \frac{300}{9}\text{ mol}$.\\
Tổng $n_{\ce{SO3}} = \frac{10}{9} + \frac{300}{9} = \frac{310}{9}\text{ mol}$.\\
$\bullet$ Sơ đồ bảo toàn: $\ce{FeS2 -> 2SO3} \Rightarrow n_{\ce{FeS2}} = \frac{1}{2} n_{\ce{SO3}} = \frac{155}{9}\text{ mol}$.\\
$m_{\ce{FeS2}} = \frac{155}{9} \times 120 = 2066{,}67\text{ g} \approx 2{,}067\text{ kg}$.\\
Vậy $m \approx 2{,}067\text{ kg}$ (hoặc 2,07 kg).
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{2pt}
\multido{}{12}{\noindent\dotfill\par\vspace{2pt}}
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
    tex_hs = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_CK1_Hoa_11_NguyenChiThanh_Ma1101_HocSinh.tex")
    tex_gv = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_CK1_Hoa_11_NguyenChiThanh_Ma1101_GiaoVien.tex")

    print("Đang tạo mã nguồn LaTeX Bản Học sinh THPT Nguyễn Chí Thanh...")
    with open(tex_hs, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))

    print("Đang tạo mã nguồn LaTeX Bản Giáo viên THPT Nguyễn Chí Thanh...")
    with open(tex_gv, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))

    print("\nĐang biên dịch PDF Bản Học sinh THPT Nguyễn Chí Thanh...")
    compile_pdf(tex_hs)

    print("\nĐang biên dịch PDF Bản Giáo viên THPT Nguyễn Chí Thanh...")
    compile_pdf(tex_gv)

    print("\n[HOÀN TẤT] Toàn bộ file .tex và .pdf đã được tạo và biên dịch thành công!")

if __name__ == "__main__":
    main()
