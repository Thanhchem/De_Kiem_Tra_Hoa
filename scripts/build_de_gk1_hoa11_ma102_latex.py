# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Giữa Học Kì I Hóa Học 11 - Mã Đề 102
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_GK1_Hoa_11_Ma102_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_GK1_Hoa_11_Ma102_GiaoVien.pdf
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

    # Tiêu đề đề thi
    tex.append(r"""
\noindent
\begin{tabularx}{\textwidth}{@{}X r@{}}
\textbf{TRƯỜNG THPT CHUYÊN / THPT TỈNH} & \textbf{ĐỀ KIỂM TRA GIỮA HỌC KÌ I} \\
\textbf{ĐỀ CHÍNH THỨC} & \textbf{MÔN: HOÁ HỌC - LỚP 11} \\
\textit{(Đề thi có 04 trang)} & \textbf{Thời gian làm bài: 50 phút} \textit{(không kể thời gian phát đề)} \\
& \textbf{MÃ ĐỀ THI: 102}
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
\noindent\textbf{Họ, tên thí sinh:} \dotfill \textbf{Lớp:} \dots\dots\dots \textbf{Số báo danh:} \dots\dots\dots

\vspace{2pt}
\noindent\textit{(Cho biết nguyên tử khối: Cu = 64; Fe = 56; N = 14; O = 16; S = 32; H = 1; K = 39; Cl = 35,5; C = 12; Ca = 40).}
\vspace{2pt}
\hrule height 0.5pt
\vspace{6pt}
""")

    # Bảng đáp án nhanh cho giáo viên
    if is_teacher:
        tex.append(r"""
\begin{center}
{\color{hfRed}\textbf{\large BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - MÃ ĐỀ 102)}}
\end{center}

\noindent{\color{mainGreen}\textbf{1. Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - mỗi câu đúng 0,25 điểm)}}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|*{12}{>{\centering\arraybackslash}m{0.95cm}|}}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} \\
\hline
\end{tabular}
\end{center}

\vspace{3pt}
\noindent{\color{mainGreen}\textbf{2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm)}}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Lệnh hỏi} & \textbf{Ý a)} & \textbf{Ý b)} & \textbf{Ý c)} & \textbf{Ý d)} \\
\hline
\textbf{Câu 13} & Đúng / Sai & \textbf{\color{hfRed}Đ} & \textbf{S} & \textbf{S} & \textbf{\color{hfRed}Đ} \\
\hline
\textbf{Câu 14} & Đúng / Sai & \textbf{S} & \textbf{\color{hfRed}Đ} & \textbf{S} & \textbf{\color{hfRed}Đ} \\
\hline
\end{tabular}
\end{center}

\vspace{3pt}
\noindent{\color{mainGreen}\textbf{3. Phần III: Trắc nghiệm Trả lời ngắn (2,0 điểm - mỗi câu đúng 0,5 điểm)}}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Câu 15} & \textbf{Câu 16} & \textbf{Câu 17} & \textbf{Câu 18} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}1} & \textbf{\color{hfRed}64,0} & \textbf{\color{hfRed}3} & \textbf{\color{hfRed}123} \\
\hline
\end{tabular}
\end{center}
\vspace{6pt}
\hrule height 0.5pt
\vspace{6pt}
""")

    # PHẦN A. TRẮC NGHIỆM
    tex.append(r"""
\noindent{\color{mainGreen}\textbf{\Large A. PHẦN TRẮC NGHIỆM (7,0 ĐIỂM)}}
\vspace{4pt}

\noindent{\color{mainGreen}\textbf{\large PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.}} \\
\textit{Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án. (3,0 điểm)}
\vspace{4pt}
""")

    # 12 CÂU PHẦN I
    p1_items = [
        ("Sự chuyển dịch cân bằng hóa học là sự dịch chuyển từ trạng thái cân bằng hóa học này",
         ["sang trạng thái không cân bằng do tác động của các yếu tố từ bên ngoài tác động lên cân bằng.",
          "sang trạng thái cân bằng hóa học khác do tác động của các yếu tố từ bên ngoài tác động lên cân bằng.",
          "sang trạng thái cân bằng hóa học khác không cần có tác động của các yếu tố từ bên ngoài tác động lên cân bằng.",
          "sang trạng thái cân bằng hóa học khác do cân bằng hóa học tác động lên các yếu tố bên ngoài."],
         1,
         r"Sự chuyển dịch cân bằng hóa học là sự dịch chuyển từ trạng thái cân bằng này sang trạng thái cân bằng khác do tác động của các yếu tố bên ngoài (nhiệt độ, nồng độ, áp suất) lên cân bằng."),

        (r"Cho \ce{Cu} phản ứng với dung dịch \ce{HNO3} loãng, nóng thu được một chất khí không màu hóa nâu trong không khí, khí đó là",
         [r"\ce{NO}.", r"\ce{NH3}.", r"\ce{N2O}.", r"\ce{N2}."],
         0,
         r"Phản ứng: $3\ce{Cu} + 8\ce{HNO3}\text{ (loãng)} \xrightarrow{t^o} 3\ce{Cu(NO3)2} + 2\ce{NO}\uparrow + 4\ce{H2O}$. Khí \ce{NO} không màu hóa nâu ngoài không khí tạo \ce{NO2}: $2\ce{NO} + \ce{O2} \to 2\ce{NO2}$."),

        (r"Trong thực tế, khi bảo quản dung dịch muối sắt(III) chloride (\ce{FeCl3}) trong phòng thí nghiệm, người ta thường phải cho thêm vài giọt dung dịch \ce{HCl} đặc vào lọ đựng. Việc làm này nhằm mục đích chủ yếu là:",
         [r"Ngăn cản quá trình bay hơi của dung dịch \ce{FeCl3}, làm cho dung dịch trong suốt, không màu.",
          r"Ức chế sự thủy phân của ion \ce{Fe^3+} tạo kết tủa \ce{Fe(OH)3} và ngăn sự oxi hóa \ce{Fe^2+} (nếu có).",
          r"Làm tăng màu vàng nâu đặc trưng của dung dịch \ce{FeCl3}.",
          r"Tăng tốc độ điện li của muối \ce{FeCl3}."],
         1,
         r"Trong nước: $\ce{Fe^3+} + 3\ce{H2O} \rightleftharpoons \ce{Fe(OH)3}\downarrow + 3\ce{H+}$. Thêm vài giọt \ce{HCl} đặc (tăng nồng độ $\ce{H+}$) làm cân bằng dịch chuyển sang chiều nghịch, ức chế sự thủy phân tạo kết tủa."),

        (r"Ở điều kiện thường, sulfur tồn tại ở dạng tinh thể, được tạo nên từ các phân tử sulfur. Số nguyên tử trong mỗi phân tử sulfur là",
         ["4.", "8.", "6.", "2."],
         1,
         r"Ở điều kiện thường, lưu huỳnh tồn tại ở dạng tinh thể gồm các phân tử $\ce{S8}$ liên kết vòng 8 nguyên tử."),

        (r"Với xúc tác của các ion kim loại trong khói bụi, các oxide của sulfur và nitrogen bị oxi hóa bởi oxygen, ozone, hydrogen peroxide, gốc tự do,... rồi hòa tan vào nước tạo thành các acid tương ứng. Hai acid tạo thành từ quá trình trên là",
         [r"\ce{H2S} và \ce{H2SO4}.", r"\ce{H2CO3} và \ce{HNO3}.", r"\ce{H2S} và \ce{HNO3}.", r"\ce{H2SO4} và \ce{HNO3}."],
         3,
         r"Đây là cơ chế tạo mưa acid: các oxide $\ce{SO2}$ tạo $\ce{H2SO4}$ và $\ce{NO}_x$ tạo $\ce{HNO3}$."),

        (r"Trong khí quyển, nguyên tố nitrogen tồn tại chủ yếu dưới dạng",
         [r"\ce{N2}.", r"\ce{NH4+}.", r"\ce{NO3-}.", r"\ce{NO2-}."],
         0,
         r"Khí quyển chứa khoảng $78\%$ thể tích là đơn chất khí nitrogen (\ce{N2})."),

        (r"Hiện tượng phú dưỡng có thể quan sát được thông qua sự xuất hiện dày đặc của tảo xanh trong nước. Khi đó, lượng oxygen trong nước sẽ nhanh chóng giảm đi. Điều này là do:\\"
         r"(a) Quá trình phân hủy tảo chết bởi vi khuẩn đã tiêu tốn lượng lớn oxygen trong nước.\\"
         r"(b) Sự phát triển của tảo xanh đã ngăn cản ánh sáng và không khí chứa oxygen khuếch tán vào nước.\\"
         r"(c) Sự hoạt động của lượng lớn vi khuẩn đã hấp thụ đáng kể oxygen hòa tan trong nước.\\"
         r"Số nhận định \textbf{đúng} là",
         ["1.", "0.", "2.", "3."],
         3,
         r"Cả 3 nhận định (a), (b), (c) đều đúng mô tả quá trình suy giảm nồng độ oxygen hòa tan (DO) khi ao hồ bị phú dưỡng."),

        (r"Chất nào sau đây thuộc loại chất điện li yếu?",
         [r"\ce{NH4Cl}.", r"\ce{KCl}.", r"\ce{HNO3}.", r"\ce{HF}."],
         3,
         r"\ce{HF} là acid yếu nên chỉ phân li một phần trong nước, thuộc loại chất điện li yếu. Các chất còn lại là muối tan và acid mạnh (chất điện li mạnh)."),

        (r"Khí nitrogen ít tan trong nước là do",
         ["nitrogen có độ âm điện lớn.",
          "phân tử nitrogen không phân cực.",
          "nitrogen có bán kính nguyên tử nhỏ.",
          "phân tử nitrogen có liên kết ba bền vững."],
         1,
         r"Phân tử nitrogen (\ce{N\equiv N}) có cấu tạo đối xứng hoàn toàn, không phân cực nên rất ít tan trong dung môi phân cực như nước."),

        (r"Phân tử ammonia có dạng hình học nào sau đây?",
         ["Chóp tam giác.", "Chóp tứ giác.", "Chữ T.", "Tam giác đều."],
         0,
         r"Nguyên tử $\ce{N}$ trong phân tử $\ce{NH3}$ ở trạng thái lai hóa $sp^3$, có 1 cặp electron chưa liên kết, tạo nên dạng hình học chóp tam giác."),

        (r"Chất chỉ thị acid - base là những chất",
         ["Chỉ có khả năng đổi màu trong môi trường base mạnh.",
          "Có màu sắc thay đổi tùy thuộc vào giá trị pH của dung dịch.",
          "Chỉ có khả năng đổi màu trong môi trường acid mạnh.",
          "Có màu sắc không thay đổi khi tiếp xúc với các acid hoặc base."],
         1,
         r"Chất chỉ thị acid - base là chất có màu biến đổi theo giá trị pH của dung dịch (quỳ tím, phenolphthalein,...)."),

        (r"Khi tiến hành các thí nghiệm khảo sát tính háo nước và tính oxi hóa mạnh của \ce{H2SO4} đặc trên các mẫu vật liệu: tờ giấy trắng, hạt gạo, và đường mía; hiện tượng chung quan sát được trên các mẫu vật này là gì và giải thích \textbf{đúng nhất} là?",
         ["Cả ba mẫu vật chỉ bị mất nước tạm thời, khi cho rửa lại bằng nước sạch thì các mẫu vật sẽ trở lại trạng thái ban đầu.",
          "Giấy và đường mía bị hóa đen do bị cháy do nhiệt độ cao, còn gạo không có hiện tượng gì.",
          r"Tất cả các mẫu vật đều bị hóa đen; nguyên nhân là do \ce{H2SO4} đặc chiếm giữ các nguyên tố H và O (thành phần cấu tạo nên carbohydrate) để lại khung carbon.",
          r"H2SO4 đặc thủy phân hoàn toàn các chất này thành dung dịch acid hữu cơ làm tờ giấy trắng, hạt gạo, và đường mía sủi bọt khí hydrogen H2."],
         2,
         r"Các chất trên đều là carbohydrate $\ce{C_n(H2O)_m}$. \ce{H2SO4} đặc có tính háo nước cực mạnh đã hút các nguyên tố H và O, biến chúng thành carbon (than) có màu đen: $\ce{C_n(H2O)_m ->[H2SO4\text{ đ}] n C\text{ (đen)} + m H2O}$.")
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

    # PHẦN II. TRẮC NGHIỆM ĐÚNG SAI
    tex.append(r"""
\vspace{6pt}
\noindent{\color{mainGreen}\textbf{\large PHẦN II. Câu trắc nghiệm đúng sai.}} \\
\textit{Thí sinh trả lời từ câu 13 đến câu 14. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai. (2,0 điểm)}
\vspace{4pt}
""")

    # CÂU 13
    tex.append(r"""
\begin{ex}
Cho phản ứng: $\ce{N2(g) + 3H2(g) <=> 2NH3(g)} \quad \Delta_r H^0_{298} = -92\text{ kJ}$.
\choiceTF
""")
    tf13 = [
        (r"Ở một nhiệt độ xác định, biểu thức xác định hằng số cân bằng $K_c = \frac{[\ce{NH3}]^2}{[\ce{N2}][\ce{H2}]^3}$.", True),
        (r"Vì phản ứng thuận là phản ứng tỏa nhiệt, nên để nâng cao hiệu suất tổng hợp \ce{NH3} trong công nghiệp, người ta thường thực hiện phản ứng ở nhiệt độ vừa phải $500^\circ\text{C}$.", False),
        (r"\ce{NH3} một chất khí không màu, không mùi, làm xanh giấy quỳ tím ẩm.", False),
        (r"Biến thiên enthalpy của phản ứng có giá trị âm $\Delta_r H^0_{298} < 0$ chứng tỏ hệ tỏa nhiệt ra môi trường khi phản ứng thuận xảy ra.", True)
    ]
    for stmt, val in tf13:
        if val:
            tex.append(f"\\True {stmt}\n")
        else:
            tex.append(f"{stmt}\n")
    if is_teacher:
        tex.append(r"""\loigiai{
\textbf{Ý a) ĐÚNG.} Theo định luật tác dụng khối lượng: $K_c = \frac{[\ce{NH3}]^2}{[\ce{N2}][\ce{H2}]^3}$.\\
\textbf{Ý b) SAI.} Để nâng cao hiệu suất (chuyển dịch theo chiều thuận tỏa nhiệt) về mặt nhiệt động học phải \textit{giảm nhiệt độ}. Ở công nghiệp chọn $\approx 500^\circ\text{C}$ là nhiệt độ tối ưu để đạt vận tốc phản ứng đủ nhanh với xúc tác, chứ nhiệt độ này làm giảm hiệu suất so với nhiệt độ thấp.\\
\textbf{Ý c) SAI.} Khí \ce{NH3} có mùi khai xốc đặc trưng, không phải ``không mùi''.\\
\textbf{Ý d) ĐÚNG.} $\Delta_r H^0_{298} = -92\text{ kJ} < 0$ biểu thị phản ứng thuận tỏa nhiệt.
}
""")
    else:
        tex.append("\\loigiai{}\n")
    tex.append("\\end{ex}\n")

    # CÂU 14
    tex.append(r"""
\begin{ex}
Sulfur dioxide (\ce{SO2}) là một trong các khí gây ô nhiễm môi trường, là nguyên nhân chính gây ra mưa acid và nhiều tác hại đối với sức khỏe con người. Trong công nghiệp, khí này được tạo ra từ phản ứng đốt cháy sulfur theo cân bằng sau: $\ce{S(s) + O2(g) ->[t^o] SO2(g)}$.
\choiceTF
""")
    tf14 = [
        (r"Nước mưa bình thường không ô nhiễm có pH khoảng 5,6 do có \ce{CO2} hòa tan. Nếu nước mưa tại vùng công nghiệp nhiễm \ce{SO2}, tạo ra acid \ce{H2SO4} làm pH giảm xuống dưới 5,0 thì quỳ tím sẽ chuyển sang màu xanh.", False),
        (r"Trong tự nhiên, nguyên tố sulfur (lưu huỳnh) tồn tại phổ biến ở dạng đơn chất trong lòng đất (mỏ lưu huỳnh) và dạng hợp chất trong các khoáng vật như quặng pyrite và thạch cao.", True),
        (r"Để giảm thiểu tác hại của \ce{SO2}, các nhà máy được khuyến khích xây dựng các ống khói thật cao để khí \ce{SO2} khuếch tán nhanh lên tầng cao khí quyển, chuyển hóa triệt để thành chất an toàn mà không gây mưa acid.", False),
        (r"Để xác định hàm lượng \ce{SO2} trong không khí, người ta dẫn lượng khí này qua dung dịch \ce{KMnO4}. Một học sinh tiến hành chuẩn độ 50 mL dung dịch acid \ce{H2SO4} (được tạo ra sau khi thu khí \ce{SO2} bằng dung dịch \ce{NaOH} 0,0001 M). Thể tích \ce{NaOH} tiêu tốn trung bình là 15 mL khi dung dịch chuyển sang màu hồng nhạt (chỉ thị phenolphthalein). Nồng độ mol của dung dịch \ce{H2SO4} ban đầu thu được là $1{,}5\cdot 10^{-5}\text{ M}$.", True)
    ]
    for stmt, val in tf14:
        if val:
            tex.append(f"\\True {stmt}\n")
        else:
            tex.append(f"{stmt}\n")
    if is_teacher:
        tex.append(r"""\loigiai{
\textbf{Ý a) SAI.} Khi pH giảm dưới 5,0 (môi trường acid), quỳ tím chuyển sang màu đỏ (hoặc hồng), không phải màu xanh.\\
\textbf{Ý b) ĐÚNG.} Sulfur tồn tại ở dạng đơn chất trong mỏ ngầm và dạng hợp chất trong pyrite (\ce{FeS2}), thạch cao (\ce{CaSO4.2H2O})...\\
\textbf{Ý c) SAI.} Xây ống khói cao chỉ phân tán khí ra xa hơn chứ không triệt tiêu ô nhiễm, \ce{SO2} vẫn gây mưa acid diện rộng ở nơi khác.\\
\textbf{Ý d) ĐÚNG.} $n_{\ce{NaOH}} = 0{,}0001 \times 0{,}015 = 1{,}5\cdot 10^{-6}\text{ mol} \Rightarrow n_{\ce{H2SO4}} = 0{,}75\cdot 10^{-6}\text{ mol} \Rightarrow C_M = \frac{0{,}75\cdot 10^{-6}}{0{,}050} = 1{,}5\cdot 10^{-5}\text{ M}$.
}
""")
    else:
        tex.append("\\loigiai{}\n")
    tex.append("\\end{ex}\n")

    # PHẦN III. TRẢ LỜI NGẮN
    tex.append(r"""
\vspace{6pt}
\noindent{\color{mainGreen}\textbf{\large PHẦN III. Câu hỏi trắc nghiệm yêu cầu trả lời ngắn.}} \\
\textit{Thí sinh trả lời từ câu 15 đến câu 18. (2,0 điểm)}
\vspace{4pt}
""")

    # CÂU 15
    tex.append(r"""
\begin{ex}
Giá trị pH của dung dịch \ce{H2SO4} 0,05 M là bao nhiêu?
\shortans{1}
\loigiai{
\textbf{Trả lời: 1}\\
Dung dịch \ce{H2SO4} phân li hoàn toàn: $\ce{H2SO4 -> 2H+ + SO4^2-}$.\\
$[\ce{H+}] = 2 \times 0{,}05 = 0{,}1\text{ M} = 10^{-1}\text{ M} \Rightarrow \text{pH} = -\log[\ce{H+}] = 1$.
}
\end{ex}
""")

    # CÂU 16
    tex.append(r"""
\begin{ex}
Cho a gam \ce{Cu} phản ứng hoàn toàn với dung dịch \ce{H2SO4} đặc, nóng, dư thu được 24,79 lít khí \ce{SO2} (ở đkc, sản phẩm khử duy nhất). Giá trị của a? \textit{(Làm tròn kết quả đến hàng phần mười)}
\shortans{64,0}
\loigiai{
\textbf{Trả lời: 64,0}\\
Ở đkc: $n_{\ce{SO2}} = \frac{24{,}79}{24{,}79} = 1{,}0\text{ mol}$.\\
Phương trình: $\ce{Cu + 2H2SO4\text{ (đặc)} ->[t^o] CuSO4 + SO2\uparrow + 2H2O}$.\\
Theo phương trình: $n_{\ce{Cu}} = n_{\ce{SO2}} = 1{,}0\text{ mol} \Rightarrow a = 1{,}0 \times 64 = 64{,}0\text{ gam}$.
}
\end{ex}
""")

    # CÂU 17
    tex.append(r"""
\begin{ex}
Cho các chất: \ce{Au}, \ce{Al}, \ce{FeO}, \ce{Fe(OH)2}, \ce{CaCO3}, \ce{K2SO4} lần lượt tác dụng với dung dịch \ce{HNO3} loãng. Có bao nhiêu trường hợp xảy ra phản ứng oxi hóa -- khử?
\shortans{3}
\loigiai{
\textbf{Trả lời: 3}\\
Các trường hợp xảy ra phản ứng oxi hóa -- khử gồm:
\begin{itemize}[noitemsep,topsep=1pt]
\item \ce{Al} là kim loại có tính khử: $\ce{Al + 4HNO3 -> Al(NO3)3 + NO\uparrow + 2H2O}$.
\item \ce{FeO} chứa $\ce{Fe^2+}$: $3\ce{FeO} + 10\ce{HNO3} \to 3\ce{Fe(NO3)3} + \ce{NO}\uparrow + 5\ce{H2O}$.
\item \ce{Fe(OH)2} chứa $\ce{Fe^2+}$: $3\ce{Fe(OH)2} + 10\ce{HNO3} \to 3\ce{Fe(NO3)3} + \ce{NO}\uparrow + 8\ce{H2O}$.
\end{itemize}
(\ce{Au} không phản ứng; \ce{CaCO3} là phản ứng trao đổi; \ce{K2SO4} không phản ứng).
}
\end{ex}
""")

    # CÂU 18
    tex.append(r"""
\begin{ex}
Xét trạng thái tự nhiên của sulfur:
\begin{enumerate}[label=\arabic*),noitemsep,topsep=2pt]
\item Trong tự nhiên sulfur có cả ở dạng đơn chất và hợp chất.
\item Trong cơ thể người, sulfur có trong thành phần của nhiều protein và enzyme.
\item Dạng hợp chất sulfur thường có ở các khoáng vật như pyrite (\ce{FeS2}), chu sa (\ce{HgS}), ...
\item Dạng đơn chất sulfur thường có ở gần các hồ nước mặn.
\end{enumerate}
Hãy sắp xếp các mệnh đề \textbf{đúng} theo số thứ tự tăng dần (ví dụ 12, 123, ...)
\shortans{123}
\loigiai{
\textbf{Trả lời: 123}\\
Mệnh đề (1), (2), (3) đúng. Mệnh đề (4) sai vì đơn chất sulfur thường có ở vùng mỏ ngầm hoặc gần miệng núi lửa, hồ nước mặn chứa các muối sulfate hòa tan. Thứ tự các mệnh đề đúng: 123.
}
\end{ex}
""")

    # PHẦN B. TỰ LUẬN
    tex.append(r"""
\vspace{6pt}
\noindent{\color{mainGreen}\textbf{\Large B. PHẦN TỰ LUẬN (3,0 ĐIỂM)}}
\vspace{4pt}

\noindent\textbf{Câu 19. (1,0 điểm)} Trình bày cách pha loãng dung dịch sulfuric acid đặc an toàn và nguyên tắc xử lí sơ bộ khi bị bỏng sulfuric acid.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
1. \textbf{Cách pha loãng acid sulfuric đặc an toàn:}\\
- Rót từ từ acid sulfuric đặc vào nước dọc theo thân đũa thủy tinh và khuấy nhẹ.\\
- \textit{Tuyệt đối không làm ngược lại} (không rót nước vào acid đặc), vì \ce{H2SO4} đặc tan trong nước tỏa nhiệt rất mạnh làm nước sôi đột ngột và bắn các giọt acid nóng ra ngoài gây bỏng nguy hiểm.\\
2. \textbf{Nguyên tắc xử lí sơ bộ khi bị bỏng sulfuric acid:}\\
- Cởi bỏ ngay quần áo dính acid.\\
- Rửa vết bỏng dưới vòi nước sạch chảy liên tục trong ít nhất 15 -- 20 phút để trôi acid và hạ nhiệt độ.\\
- Rửa lại nhẹ bằng dung dịch \ce{NaHCO3} loãng ($2\% - 5\%$) để trung hòa lượng acid còn sót lại.\\
- Băng nhẹ bằng gạc vô trùng và nhanh chóng đưa nạn nhân đến cơ sở y tế gần nhất.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{10pt}
\multido{\i=1+1}{6}{\noindent\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\noindent\textbf{Câu 20. (1,0 điểm)} Cho 100 mL dung dịch \ce{KOH} x M vào dung dịch \ce{(NH4)2SO4} dư, đun nóng nhẹ thu được 2,479 lít khí (đo ở điều kiện chuẩn đkc).\\
a) Khí thu được có thể nhận biết bằng cách nào?\\
b) Tính x?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
Phương trình phản ứng: $2\ce{KOH} + \ce{(NH4)2SO4} \xrightarrow{t^o} \ce{K2SO4} + 2\ce{NH3}\uparrow + 2\ce{H2O}$.\\
a) Khí sinh ra là ammonia (\ce{NH3}). Cách nhận biết:\\
- Khí có mùi khai xốc đặc trưng.\\
- Quỳ tím ẩm hóa xanh khi tiếp xúc với khí \ce{NH3}.\\
- Hoặc dùng đũa thủy tinh nhúng dung dịch \ce{HCl} đặc đưa lại gần thấy xuất hiện làn khói trắng (\ce{NH4Cl}).\\
b) Tính x:\\
- Ở đkc: $n_{\ce{NH3}} = \frac{2{,}479}{24{,}79} = 0{,}1\text{ mol}$.\\
- Theo phương trình: $n_{\ce{KOH}} = n_{\ce{NH3}} = 0{,}1\text{ mol}$.\\
- $V_{\text{dd}} = 100\text{ mL} = 0{,}1\text{ L} \Rightarrow x = C_M = \frac{0{,}1}{0{,}1} = \mathbf{1{,}0\text{ M}}$.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{10pt}
\multido{\i=1+1}{6}{\noindent\dotfill\\[4pt]}
""")

    tex.append(r"""
\vspace{6pt}
\noindent\textbf{Câu 21. (1,0 điểm)} Cho cân bằng hóa học sau trong pha khí:\\
\centerline{$2\ce{SO2(g) + O2(g) <=> 2SO3(g)} \quad \Delta_r H^0_{298} < 0$}\\
Cân bằng trên dịch chuyển theo chiều nào khi thay đổi các điều kiện sau (giải thích ngắn gọn)?\\
a) Tăng nhiệt độ.\\
b) Tăng áp suất chung của hệ phản ứng.\\
c) Dùng thêm chất xúc tác \ce{V2O5}.\\
d) Tăng nồng độ \ce{SO2}.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,arc=1mm,boxrule=0.5pt,top=2mm,bottom=2mm,left=3mm,right=3mm]
{\color{solBlue}\textbf{► Lời giải chi tiết:}}\\
Áp dụng nguyên lí chuyển dịch cân bằng Le Chatelier:\\
a) \textbf{Tăng nhiệt độ:} Chiều thuận tỏa nhiệt ($\Delta H < 0$). Khi tăng nhiệt độ, cân bằng chuyển dịch theo chiều thu nhiệt ($\Delta H > 0$) $\Rightarrow$ \textbf{chiều nghịch}.\\
b) \textbf{Tăng áp suất chung:} Số mol khí vế trái ($3\text{ mol}$) > vế phải ($2\text{ mol}$). Khi tăng áp suất, cân bằng dịch chuyển theo chiều làm giảm số mol khí $\Rightarrow$ \textbf{chiều thuận}.\\
c) \textbf{Dùng thêm xúc tác \ce{V2O5}:} Xúc tác làm tăng đồng thời tốc độ phản ứng thuận và nghịch với mức độ như nhau $\Rightarrow$ \textbf{cân bằng không chuyển dịch}.\\
d) \textbf{Tăng nồng độ \ce{SO2}:} Cân bằng dịch chuyển theo chiều làm giảm nồng độ chất thêm vào $\Rightarrow$ \textbf{chiều thuận}.
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{10pt}
\multido{\i=1+1}{7}{\noindent\dotfill\\[4pt]}
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
    tex_hs_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_11_Ma102_HocSinh.tex")
    with open(tex_hs_path, "w", encoding="utf-8") as f:
        f.write(tex_hs)
    pdf_hs_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_11_Ma102_HocSinh.pdf")
    compile_latex(tex_hs_path, pdf_hs_path)
    print(f"[OK] Đã biên dịch xong PDF Học sinh: {pdf_hs_path}")

    # 2. Bản Giáo viên
    tex_gv = build_latex_content(is_teacher=True)
    tex_gv_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_11_Ma102_GiaoVien.tex")
    with open(tex_gv_path, "w", encoding="utf-8") as f:
        f.write(tex_gv)
    pdf_gv_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_11_Ma102_GiaoVien.pdf")
    compile_latex(tex_gv_path, pdf_gv_path)
    print(f"[OK] Đã biên dịch xong PDF Giáo viên: {pdf_gv_path}")


if __name__ == "__main__":
    main()
