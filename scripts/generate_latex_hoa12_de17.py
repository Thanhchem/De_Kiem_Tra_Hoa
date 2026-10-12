# -*- coding: utf-8 -*-
"""
Script sinh file LaTeX (.tex) và biên dịch ra PDF cho:
ĐỀ KIỂM TRA SỐ 17 - CHƯƠNG 6: ĐẠI CƯƠNG KIM LOẠI (HÓA HỌC 12)
Mã đề thi: 117
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De17_HocSinh.pdf (Chuẩn chính xác 4 trang A4)
2. De_Kiem_Tra_Hoa_12_De17_GiaoVien.pdf (Bản Giáo Viên có lời giải chi tiết và đáp án)
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
import subprocess
import glob
import fitz  # PyMuPDF

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"

def create_latex_de17(is_teacher=False):
    ans_file = "ans-de17-gv" if is_teacher else "ans-de17-hs"
    
    tex = r"""\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{geometry}
\geometry{top=0.9cm,bottom=0.9cm,left=1.35cm,right=1.35cm,headheight=24pt,headsep=5pt,footskip=11pt}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{tabularx}
\usepackage{array}
\usepackage{tikz}
\usepackage{multicol}
\usepackage{fancyhdr}
\usepackage{enumitem}

\definecolor{hfRed}{HTML}{D32F2F}
\definecolor{hfGreen}{HTML}{8BC390}
\definecolor{mainGreen}{HTML}{2E7D32}

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
\setlength{\parskip}{""" + (r"1.2pt" if is_teacher else r"0.0pt") + r"""}

\def\loigiai#1{""" + (r"""
\par\begingroup\color{blue!80!black}\small
\noindent\textbf{\textit{Lời giải chi tiết:}}\par
#1\par\endgroup""" if is_teacher else "") + r"""}

\begin{document}
\Opensolutionfile{ans}[""" + ans_file + r"""]

\noindent
\begin{minipage}[t]{0.48\textwidth}
	\textbf{SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ}\\
	\textbf{TRƯỜNG THPT HAI BÀ TRƯNG}\\
	\textbf{ĐỀ CHÍNH THỨC}\\
	\textit{(Đề thi có 04 trang)}
\end{minipage}\hfill
\begin{minipage}[t]{0.48\textwidth}
	\raggedleft
	\textbf{\color{mainGreen}ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 6 (ĐẠI CƯƠNG KIM LOẠI)}\\
	\textbf{MÔN: HOÁ HỌC -- LỚP 12}\\
	\textit{Thời gian làm bài: 50 phút (không kể thời gian phát đề)}\\
	\textbf{Mã đề thi: 117}
\end{minipage}

"""

    if is_teacher:
        tex += r"""
\vspace{0.1cm}
\begin{center}
	\textbf{\color{hfRed}\Large (BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}
\end{center}
"""

    tex += r"""
\vspace{1pt}
\noindent\textit{Họ, tên thí sinh: \dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots Lớp: \dots\dots\dots\dots SBD: \dots\dots\dots\dots}\\
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108; Ba = 137; Au = 197; Pb = 207; hằng số Faraday F = 96500 C/mol).}

\vspace{2pt}
\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\setcounter{ex}{0}

\begin{ex}
Ở trạng thái cơ bản, cấu hình electron lớp ngoài cùng của nguyên tử Al ($Z = 13$) là
\choice
{$\ce{3s^1}$.}
{$\ce{3s^2}$.}
""" + (r"{\True $\ce{3s^2 3p^1}$.}" if is_teacher else r"{$\ce{3s^2 3p^1}$.}") + r"""
{$\ce{3p^1}$.}
\loigiai{""" + (r"""
Nguyên tử Al có $Z = 13$: $\ce{1s^2 2s^2 2p^6 3s^2 3p^1}$. Lớp ngoài cùng (lớp $n = 3$) có cấu hình là $\ce{3s^2 3p^1}$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Ở trạng thái cơ bản, nguyên tử của nguyên tố nào sau đây có số electron lớp ngoài cùng ít nhất?
\choice
{C ($Z = 6$).}
""" + (r"{\True Li ($Z = 3$).}" if is_teacher else r"{Li ($Z = 3$).}") + r"""
{O ($Z = 8$).}
{F ($Z = 9$).}
\loigiai{""" + (r"""
Cấu hình electron lớp ngoài cùng:\\
-- C ($Z = 6$): $\ce{2s^2 2p^2}$ (4 electron).\\
-- Li ($Z = 3$): $\ce{2s^1}$ (1 electron).\\
-- O ($Z = 8$): $\ce{2s^2 2p^4}$ (6 electron).\\
-- F ($Z = 9$): $\ce{2s^2 2p^5}$ (7 electron).\\
Vậy Li có số electron lớp ngoài cùng ít nhất (1 electron).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Ở trạng thái cơ bản, nguyên tử của nguyên tố nào sau đây có 3 electron ở lớp ngoài cùng?
\choice
{$_{11}\ce{Na}$.}
""" + (r"{\True $_{13}\ce{Al}$.}" if is_teacher else r"{$_{13}\ce{Al}$.}") + r"""
{$_{20}\ce{Ca}$.}
{$_{26}\ce{Fe}$.}
\loigiai{""" + (r"""
-- $_{11}\ce{Na}$: $[\ce{Ne}] \ce{3s^1}$ (1 electron lớp ngoài cùng).\\
-- $_{13}\ce{Al}$: $[\ce{Ne}] \ce{3s^2 3p^1}$ (3 electron lớp ngoài cùng).\\
-- $_{20}\ce{Ca}$: $[\ce{Ar}] \ce{4s^2}$ (2 electron lớp ngoài cùng).\\
-- $_{26}\ce{Fe}$: $[\ce{Ar}] \ce{3d^6 4s^2}$ (2 electron lớp ngoài cùng).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Thành phần nào sau đây \textbf{không} có trong mạng tinh thể kim loại?
\choice
{Ion kim loại.}
{Electron.}
{Nguyên tử kim loại.}
""" + (r"{\True Anion gốc acid.}" if is_teacher else r"{Anion gốc acid.}") + r"""
\loigiai{""" + (r"""
Mạng tinh thể kim loại bao gồm cation kim loại ở các nút mạng, các electron hoá trị tự do chuyển động và một phần nguyên tử kim loại trung hoà. Mạng tinh thể kim loại không chứa anion gốc acid.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Đốt một sợi dây kim loại X trong bình khí chlorine (\ce{Cl2}) thấy tạo ra khói màu nâu đỏ. X là kim loại nào sau đây?
\choice
{Mg.}
{Al.}
""" + (r"{\True Fe.}" if is_teacher else r"{Fe.}") + r"""
{Cu.}
\loigiai{""" + (r"""
Sắt cháy trong khí chlorine theo phản ứng: $\ce{2Fe + 3Cl2 -> 2FeCl3}$ (tạo khói màu nâu đỏ đặc trưng của các hạt tinh thể \ce{FeCl3} rắn).\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Thuỷ ngân (Hg) là chất lỏng ở điều kiện thường, dễ bay hơi và hơi thuỷ ngân rất độc. Khi nhiệt kế thuỷ ngân bị vỡ, để tránh thuỷ ngân phân tán ra môi trường xung quanh, người ta gom thuỷ ngân lại rồi trộn với chất nào sau đây (ở dạng bột)?
\choice
{Carbon (C).}
{Magnesium (Mg).}
{Iron (Fe).}
""" + (r"{\True Sulfur (S).}" if is_teacher else r"{Sulfur (S).}") + r"""
\loigiai{""" + (r"""
Dùng bột lưu huỳnh (S) rắc lên thuỷ ngân vì phản ứng xảy ra ngay ở nhiệt độ thường tạo muối HgS kết tủa bền, không bay hơi và không độc hại: $\ce{Hg + S -> HgS}$.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Tính chất hoá học đặc trưng của kim loại là
\choice
{tính oxi hoá.}
{tính acid.}
{tính base.}
""" + (r"{\True tính khử.}" if is_teacher else r"{tính khử.}") + r"""
\loigiai{""" + (r"""
Các nguyên tử kim loại có ít electron lớp ngoài cùng, năng lượng ion hoá thấp và bán kính nguyên tử tương đối lớn nên dễ nhường electron: $\ce{M -> M^{n+} + ne}$. Do đó, tính chất hoá học đặc trưng của kim loại là tính khử.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong vỏ Trái Đất, kim loại nào sau đây có thể tồn tại ở dạng đơn chất?
\choice
{Na, Mg.}
{Al, Fe.}
{Cu, Zn.}
""" + (r"{\True Ag, Au.}" if is_teacher else r"{Ag, Au.}") + r"""
\loigiai{""" + (r"""
Các kim loại rất kém hoạt động hoá học như bạc (Ag), vàng (Au), platinum (Pt) có thế điện cực chuẩn dương lớn, trơ về mặt hoá học nên có thể tồn tại trong tự nhiên ở dạng đơn chất.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Ion kim loại nào sau đây có nhiều nhất trong nước biển?
\choice
{\ce{Mg^2+}.}
{\ce{K+}.}
""" + (r"{\True \ce{Na+}.}" if is_teacher else r"{\ce{Na+}.}") + r"""
{\ce{Ca^2+}.}
\loigiai{""" + (r"""
Nước biển chứa hàm lượng muối \ce{NaCl} rất cao (chiếm khoảng $77{,}8\%$ lượng muối hoà tan), do đó cation kim loại dồi dào nhất trong nước biển là \ce{Na+}.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hai nguyên tố kim loại phổ biến nhất trong vỏ Trái Đất là
\choice
{Mg, Ca.}
{Na, Ca.}
{Au, Ag.}
""" + (r"{\True Al, Fe.}" if is_teacher else r"{Al, Fe.}") + r"""
\loigiai{""" + (r"""
Trong vỏ Trái Đất, nhôm (Al, chiếm khoảng 8\%) là kim loại phổ biến nhất, tiếp theo là sắt (Fe, chiếm khoảng 5\%).\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\begin{ex}
Nhôm (Al) là nguyên tố phổ biến thứ ba và là kim loại phổ biến nhất trong vỏ Trái Đất. Trong tự nhiên, quặng chính chứa nhôm là bauxite và quặng này là nguyên liệu chính để sản xuất nhôm trong công nghiệp. Thành phần chính của quặng bauxite là
\choice
{\ce{Na3AlF6}.}
""" + (r"{\True \ce{Al2O3.2H2O}.}" if is_teacher else r"{\ce{Al2O3.2H2O}.}") + r"""
{\ce{KAl(SO4)2.12H2O}.}
{\ce{K2O.Al2O3.6SiO2}.}
\loigiai{""" + (r"""
Quặng bauxite có thành phần chính là nhôm oxide ngậm nước \ce{Al2O3.2H2O} (hoặc \ce{Al2O3.xH2O} chứa tạp chất \ce{Fe2O3}, \ce{SiO2},\dots).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hợp kim là vật liệu kim loại chứa một kim loại cơ bản với
\choice
""" + (r"{\True một số kim loại khác hoặc phi kim.}" if is_teacher else r"{một số kim loại khác hoặc phi kim.}") + r"""
{một số oxide của kim loại đó.}
{một số oxide kim loại khác hoặc phi kim.}
{một số phi kim và oxide của phi kim đó.}
\loigiai{""" + (r"""
Hợp kim là vật liệu kim loại chứa một kim loại cơ bản (kim loại nền) kết hợp với một số kim loại khác hoặc phi kim nhằm cải thiện các tính chất cơ học, vật lý và hoá học.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phát biểu nào sau đây sai?
\choice
{Tính chất vật lí của hợp kim khác nhiều so với của đơn chất thành phần.}
{Tính chất hoá học của hợp kim tương tự tính chất của đơn chất thành phần.}
{Tính chất cơ học của hợp kim phụ thuộc vào thành phần đơn chất của hợp kim.}
""" + (r"{\True Tính chất cơ học của hợp kim không phụ thuộc vào hàm lượng các đơn chất.}" if is_teacher else r"{Tính chất cơ học của hợp kim không phụ thuộc vào hàm lượng các đơn chất.}") + r"""
\loigiai{""" + (r"""
Tính chất cơ học (độ cứng, độ bền, độ đàn hồi...) của hợp kim phụ thuộc chặt chẽ vào thành phần và tỉ lệ phần trăm hàm lượng của các nguyên tố trong hợp kim. Do đó phát biểu D là sai.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trường hợp nào sau đây đã sử dụng phương pháp điện hoá để chống ăn mòn kim loại?
\choice
{Mạ vàng lên quai đồng hồ.}
{Bôi dầu mỡ lên xích xe đạp.}
""" + (r"{\True Gắn tấm kẽm lên mặt ngoài vỏ tàu biển.}" if is_teacher else r"{Gắn tấm kẽm lên mặt ngoài vỏ tàu biển.}") + r"""
{Sơn kín bề mặt khung cửa thép.}
\loigiai{""" + (r"""
Gắn tấm kẽm (Zn) lên vỏ tàu biển bằng thép là phương pháp bảo vệ điện hoá (dùng kim loại hy sinh). Kẽm có tính khử mạnh hơn sắt nên đóng vai trò là anode bị ăn mòn, bảo vệ vỏ tàu sắt không bị gỉ. Các phương án A, B, D là phương pháp phủ bảo vệ bề mặt.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trường hợp nào sau đây \textbf{không} xảy ra sự ăn mòn kim loại?
\choice
{Tàu đánh cá làm bằng thép bị hoen gỉ sau thời gian đi biển về.}
{Trống đồng bị chuyển màu xanh khi để lâu ngày trong không khí ẩm.}
{Vòng tay làm bằng bạc kim loại bị hoá đen khi sử dụng lâu ngày.}
""" + (r"{\True Nấu chảy vàng để đúc khuôn khi chế tác vàng trang sức.}" if is_teacher else r"{Nấu chảy vàng để đúc khuôn khi chế tác vàng trang sức.}") + r"""
\loigiai{""" + (r"""
Nấu chảy vàng chỉ là quá trình chuyển thể vật lý từ thể rắn sang thể lỏng, không có sự biến đổi chất hay quá trình oxi hoá kim loại nên không phải là sự ăn mòn kim loại.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Khi kim loại bị ăn mòn, luôn xảy ra quá trình nào sau đây?
\choice
""" + (r"{\True Oxi hoá kim loại thành ion kim loại.}" if is_teacher else r"{Oxi hoá kim loại thành ion kim loại.}") + r"""
{Khử ion kim loại thành kim loại.}
{Khử oxygen trong không khí.}
{Khử kim loại thành ion kim loại.}
\loigiai{""" + (r"""
Bản chất của mọi quá trình ăn mòn kim loại là sự oxi hoá kim loại thành ion kim loại: $\ce{M -> M^{n+} + ne}$.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trường hợp nào sau đây kim loại bị phá huỷ chủ yếu do ăn mòn hoá học?
\choice
""" + (r"{\True Thiết bị làm bằng thép trong lò đốt lâu ngày bị phá huỷ.}" if is_teacher else r"{Thiết bị làm bằng thép trong lò đốt lâu ngày bị phá huỷ.}") + r"""
{Thép xây dựng bị gỉ khi để lâu ngày trong không khí ẩm.}
{Ống nước làm bằng gang bị gỉ khi chôn dưới đất lâu ngày.}
{Vỏ tàu biển làm bằng thép bị gỉ sau một thời gian sử dụng.}
\loigiai{""" + (r"""
Thiết bị làm bằng thép trong lò đốt tiếp xúc trực tiếp với các khí oxi hoá (\ce{O2}, \ce{CO2}, \ce{SO2}\dots) ở nhiệt độ rất cao, các electron chuyển trực tiếp sang môi trường mà không phát sinh dòng điện nên đây là ăn mòn hoá học. Các trường hợp còn lại đều xảy ra trong dung dịch chất điện li nên là ăn mòn điện hoá học.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trường hợp nào sau đây kim loại bị phá huỷ chủ yếu do ăn mòn điện hoá học?
\choice
{Thiết bị làm bằng sắt trong lò đốt lâu ngày bị phá huỷ.}
{Thước nhôm để trong không khí bị phủ bởi lớp \ce{Al2O3} ở bề mặt.}
""" + (r"{\True Cánh cửa làm bằng thép nhanh bị gỉ khi để ngoài trời mưa.}" if is_teacher else r"{Cánh cửa làm bằng thép nhanh bị gỉ khi để ngoài trời mưa.}") + r"""
{Đốt dây sắt trong bình khí chlorine.}
\loigiai{""" + (r"""
Thép là hợp kim của sắt và carbon. Khi gặp nước mưa (chứa chất điện li), tạo thành vô số pin vi mô Fe -- C dẫn đến ăn mòn điện hoá học của sắt.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\setcounter{ex}{0}

\begin{ex}
Tùy thuộc vào tính chất vật lí riêng của mỗi kim loại mà chúng được sử dụng vào những mục đích khác nhau. Mỗi phát biểu sau đây là đúng hay sai?
\choiceTF
""" + (r"{\True Kim loại chì (Pb) và cadmium (Cd) có nhiệt nóng chảy khá thấp nên được sử dụng làm dây chảy trong cầu chì.}" if is_teacher else r"{Kim loại chì (Pb) và cadmium (Cd) có nhiệt nóng chảy khá thấp nên được sử dụng làm dây chảy trong cầu chì.}") + r"""
""" + (r"{\True Kim loại tungsten (W) có độ bền nhiệt và nhiệt độ nóng chảy rất cao nên được sử dụng làm dây tóc bóng đèn, thiết bị sưởi.}" if is_teacher else r"{Kim loại tungsten (W) có độ bền nhiệt và nhiệt độ nóng chảy rất cao nên được sử dụng làm dây tóc bóng đèn, thiết bị sưởi.}") + r"""
""" + (r"{\True Do có tính dẻo và độ cứng phù hợp nên nhôm (Al) thường được gia công làm vật liệu như khung cửa, khung thiết bị.}" if is_teacher else r"{Do có tính dẻo và độ cứng phù hợp nên nhôm (Al) thường được gia công làm vật liệu như khung cửa, khung thiết bị.}") + r"""
{Do kim loại magnesium (Mg) có khối lượng riêng là $1{,}735\text{ g/cm}^3$ nên được dùng để chế tạo các hợp kim nặng.}
\loigiai{""" + (r"""
a) \textbf{Đúng}. Chì và cadmium có nhiệt độ nóng chảy thấp nên hợp kim của chúng được ứng dụng làm dây chảy cầu chì tự ngắt khi quá nhiệt.\\
b) \textbf{Đúng}. Tungsten có nhiệt độ nóng chảy cao nhất trong các kim loại ($3422^\circ\text{C}$) nên được dùng làm dây tóc bóng đèn sợi đốt.\\
c) \textbf{Đúng}. Nhôm có khối lượng nhẹ, tính dẻo cao và dễ uốn tạo hình khung cửa, khung kết cấu.\\
d) \textbf{Sai}. Magnesium có khối lượng riêng nhỏ ($1{,}735\text{ g/cm}^3 < 5\text{ g/cm}^3$) nên là kim loại nhẹ, hợp kim của nó là hợp kim nhẹ.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Đồng (Cu) là kim loại có tính khử yếu, không tan trong dung dịch \ce{H2SO4} loãng, nhưng tan được trong dung dịch \ce{H2SO4} đặc, nóng hoặc trong dung dịch \ce{H2SO4} loãng khi có mặt của \ce{O2} ngay ở nhiệt độ thường theo các phương trình hoá học sau:\\
(1) $\ce{Cu + 2H2SO4} \text{ (đặc, nóng)} \rightarrow \ce{CuSO4 + SO2 + 2H2O}$\\
(2) $\ce{2Cu + 2H2SO4} \text{ (loãng)} + \ce{O2} \rightarrow \ce{2CuSO4 + 2H2O}$\\
Mỗi phát biểu sau đây là đúng hay sai?
\choiceTF
{Trong hai phản ứng trên, \ce{H2SO4} đều đóng vai trò là chất oxi hoá.}
""" + (r"{\True Cùng một lượng \ce{CuSO4} tạo ra thì phản ứng (1) tiêu tốn \ce{H2SO4} nhiều hơn phản ứng (2).}" if is_teacher else r"{Cùng một lượng \ce{CuSO4} tạo ra thì phản ứng (1) tiêu tốn \ce{H2SO4} nhiều hơn phản ứng (2).}") + r"""
{Do có giá trị thế điện cực chuẩn dương nên Cu tác dụng được với dung dịch \ce{H2SO4} loãng ở (2).}
""" + (r"{\True Từ phản ứng (2) chứng tỏ $E^\circ_{\ce{Cu^2+/Cu}} < E^\circ_{\ce{O2, 4H+/2H2O}}$.}" if is_teacher else r"{Từ phản ứng (2) chứng tỏ $E^\circ_{\ce{Cu^2+/Cu}} < E^\circ_{\ce{O2, 4H+/2H2O}}$.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Sai}. Trong phản ứng (2), chất oxi hoá là \ce{O2} (số oxi hoá giảm từ 0 xuống $-2$), \ce{H2SO4} chỉ đóng vai trò môi trường.\\
b) \textbf{Đúng}. Ở phản ứng (1), để tạo 1 mol \ce{CuSO4} cần 2 mol \ce{H2SO4}; ở phản ứng (2), để tạo 1 mol \ce{CuSO4} chỉ cần 1 mol \ce{H2SO4}.\\
c) \textbf{Sai}. $E^\circ_{\ce{Cu^2+/Cu}} = +0{,}34\text{ V} > 0\text{ V}$ nên Cu không đẩy được \ce{H+} của acid loãng; phản ứng (2) xảy ra là do có mặt chất oxi hoá mạnh hơn \ce{O2} trong môi trường acid.\\
d) \textbf{Đúng}. Phản ứng xảy ra tự phát chứng tỏ chất oxi hoá mạnh hơn (\ce{O2} trong acid, $E^\circ = +1{,}23\text{ V}$) oxi hoá được Cu ($E^\circ = +0{,}34\text{ V}$).""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các phát biểu sau về ứng dụng và tính chất của kim loại. Mỗi phát biểu sau đây là đúng hay sai?
\choiceTF
{Trong công nghiệp, nhôm được sản xuất bằng cách điện phân dung dịch \ce{AlCl3}.}
{Cho mẩu nhỏ Na vào dung dịch \ce{CuSO4}, thu được chất rắn là kim loại Cu.}
""" + (r"{\True Tái chế kim loại là giải pháp giúp con người sử dụng hiệu quả hơn nguồn tài nguyên.}" if is_teacher else r"{Tái chế kim loại là giải pháp giúp con người sử dụng hiệu quả hơn nguồn tài nguyên.}") + r"""
{Tất cả hợp kim của sắt đều dễ bị ăn mòn trong không khí ẩm.}
\loigiai{""" + (r"""
a) \textbf{Sai}. Điện phân dung dịch \ce{AlCl3} tại cathode nước bị khử sinh ra \ce{H2}, không thu được Al. Nhôm được điều chế bằng điện phân nóng chảy \ce{Al2O3}.\\
b) \textbf{Sai}. Na tác dụng với nước tạo NaOH và \ce{H2}, sau đó NaOH tạo kết tủa \ce{Cu(OH)2} màu xanh, không thu được Cu kim loại.\\
c) \textbf{Đúng}. Tái chế kim loại giúp tiết kiệm tài nguyên khoáng sản, giảm tiêu tốn năng lượng và giảm ô nhiễm môi trường.\\
d) \textbf{Sai}. Thép không gỉ (inox) chứa Cr tạo lớp màng oxide thụ động bảo vệ rất bền vững, không bị gỉ trong không khí ẩm.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Tiến hành thí nghiệm (như hình vẽ bên): Rót dung dịch \ce{NaCl} bão hoà vào cốc 1, cốc 2, cốc 3; cho dầu nhờn vào cốc 4. Cho vào cốc 1 và cốc 4 một đinh sắt sạch; cho vào cốc 2 đinh sắt sạch được quấn bởi dây kẽm; cho vào cốc 3 đinh sắt sạch được quấn bởi dây đồng. Để 4 cốc trong không khí khoảng 5 ngày.
\begin{center}
	\includegraphics[width=0.68\textwidth]{images_crop_de17/p2_cau4_an_mon_sat.png}
\end{center}
Mỗi phát biểu sau đây là đúng hay sai?
\choiceTF
""" + (r"{\True Ở cốc 4, đinh sắt không bị gỉ (không bị ăn mòn).}" if is_teacher else r"{Ở cốc 4, đinh sắt không bị gỉ (không bị ăn mòn).}") + r"""
{Ở cốc 2, đinh sắt không bị gỉ, dây Zn bị ăn mòn và có khí thoát ra.}
""" + (r"{\True Ở cốc 3, đinh sắt bị gỉ nhiều nhất và dây đồng không bị ăn mòn.}" if is_teacher else r"{Ở cốc 3, đinh sắt bị gỉ nhiều nhất và dây đồng không bị ăn mòn.}") + r"""
{Ở cốc 1, đinh sắt bị gỉ và dung dịch có màu vàng của \ce{FeCl2}.}
\loigiai{""" + (r"""
a) \textbf{Đúng}. Dầu nhờn cách li đinh sắt hoàn toàn khỏi không khí và nước nên không xảy ra ăn mòn.\\
b) \textbf{Sai}. Zn bị ăn mòn bảo vệ sắt, nhưng trong môi trường \ce{NaCl} trung tính, ở cathode xảy ra sự khử oxi hoà tan ($\ce{O2 + 2H2O + 4e -> 4OH-}$), không có khí thoát ra.\\
c) \textbf{Đúng}. Cặp Fe -- Cu tạo pin điện hoá trong đó Fe là anode bị ăn mòn nhanh hơn nhiều so với cốc 1; Cu là cathode được bảo vệ không bị ăn mòn.\\
d) \textbf{Sai}. Gỉ sắt là \ce{Fe2O3.nH2O} kết tủa màu nâu đỏ; dung dịch \ce{FeCl2} loãng có màu lục rất nhạt hoặc không màu, không phải màu vàng.""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}

\setcounter{ex}{0}

\begin{ex}
Cho các nhận định sau:\\
(1) Các kim loại thuộc nhóm IA và IIA bao gồm các nguyên tố s.\\
(2) Nguyên tử của hầu hết các kim loại có từ 1 đến 3 electron ở lớp ngoài cùng.\\
(3) Nguyên tử kim loại dễ nhường electron hơn so với các nguyên tử phi kim.\\
(4) Trong mạng tinh thể kim loại, các electron chuyển động tự do theo nhiều hướng.\\
Có bao nhiêu nhận định là đúng?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}4}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Cả 4 nhận định đều đúng:\\
-- (1) Nhóm IA có cấu hình $\ce{ns^1}$, nhóm IIA có $\ce{ns^2}$ nên đều là nguyên tố s.\\
-- (2) Hầu hết nguyên tử kim loại có 1, 2 hoặc 3 electron lớp ngoài cùng.\\
-- (3) Kim loại có năng lượng ion hoá thấp nên dễ nhường electron hơn phi kim.\\
-- (4) Electron tự do chuyển động hỗn loạn không định hướng trong toàn bộ tinh thể.\\
\textbf{Đáp số: 4.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Sắt (Fe) có số hiệu nguyên tử là 26. Trong vỏ Trái Đất, sắt là nguyên tố kim loại phổ biến thứ 2 (sau nhôm). Ứng dụng chủ yếu của sắt là để tạo ra các hợp kim thép dùng trong xây dựng và chế tạo. Cho các phát biểu sau:\\
(a) Ở trạng thái cơ bản, nguyên tử Fe có 6 electron ở lớp ngoài cùng.\\
(b) Hợp kim thép carbon khi để trong không khí ẩm sẽ bị ăn mòn điện hoá học.\\
(c) Trong vỏ Trái Đất, sắt tồn tại ở dạng hợp chất chủ yếu trong các quặng.\\
(d) Trong hợp chất, số oxi hoá của sắt chủ yếu là +2 và +3.\\
(e) Kim loại sắt (dư) tác dụng với chlorine tạo ra sản phẩm là \ce{FeCl2}.\\
Hãy cho biết có bao nhiêu phát biểu đúng trong các phát biểu trên.
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}3}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Các phát biểu đúng là (b), (c), (d):\\
-- (a) Sai vì Fe ($[\ce{Ar}] \ce{3d^6 4s^2}$) có 2 electron ở lớp ngoài cùng (lớp $n = 4$).\\
-- (b) Đúng vì cặp Fe -- C tiếp xúc với màng nước điện li bị ăn mòn điện hoá.\\
-- (c) Đúng vì sắt tồn tại trong quặng hematite, magnetite, siderite, pyrite\dots\\
-- (d) Đúng vì sắt có số oxi hoá phổ biến là +2 và +3.\\
-- (e) Sai vì chlorine là chất oxi hoá mạnh, luôn oxi hoá sắt lên mức cao nhất là \ce{FeCl3} ngay cả khi sắt dư.\\
Tổng cộng có 3 phát biểu đúng.\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Tiến hành các thí nghiệm:\\
(1) Đốt dây nhôm trong bình khí oxygen.\\
(2) Cho đinh thép sạch vào dung dịch \ce{H2SO4} loãng.\\
(3) Cho lá Al vào dung dịch \ce{CuSO4}.\\
(4) Nối dây nhôm với dây đồng rồi để trong không khí ẩm.\\
Liệt kê những thí nghiệm có xảy ra ăn mòn điện hoá học theo số thứ tự tăng dần (ví dụ: 1, 2, 3).
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}2, 3, 4}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
-- (1) Đốt nhôm trong khí \ce{O2} ở nhiệt độ cao là ăn mòn hoá học.\\
-- (2) Cặp Fe -- C trong thép tiếp xúc dung dịch acid \ce{H2SO4} xảy ra ăn mòn điện hoá học.\\
-- (3) Al đẩy \ce{Cu^2+} tạo kim loại Cu bám vào lá Al, tạo cặp pin Al -- Cu trong dung dịch chất điện li xảy ra ăn mòn điện hoá học.\\
-- (4) Dây Al nối dây Cu trong không khí ẩm tạo pin Al -- Cu xảy ra ăn mòn điện hoá học.\\
Các thí nghiệm xảy ra ăn mòn điện hoá học là: 2, 3, 4.\\
\textbf{Đáp số: 2, 3, 4.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Để xảy ra ăn mòn điện hoá học thì cần đủ các yếu tố nào sau đây?\\
(1) Phải có hai điện cực khác chất nhau (hai kim loại khác nhau hoặc kim loại với phi kim).\\
(2) Hai điện cực phải tiếp xúc với nhau (trực tiếp hoặc gián tiếp qua dây dẫn).\\
(3) Hai điện cực phải được nối qua vôn kế hoặc ampe kế.\\
(4) Hai điện cực phải cùng tiếp xúc với một dung dịch chất điện li.\\
Liệt kê các yếu tố cần thiết theo số thứ tự tăng dần (ví dụ: 1, 2, 4).
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}1, 2, 4}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Điều kiện để xảy ra ăn mòn điện hoá học gồm 3 yếu tố đồng thời:\\
(1) Các điện cực khác nhau về bản chất.\\
(2) Các điện cực phải tiếp xúc trực tiếp hoặc gián tiếp qua dây dẫn.\\
(4) Các điện cực cùng tiếp xúc với một dung dịch chất điện li.\\
Yếu tố (3) không cần thiết vì dòng điện ăn mòn tự khép kín mạch ngoài mà không cần thiết bị đo.\\
\textbf{Đáp số: 1, 2, 4.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong công nghiệp sản xuất gang ở nước ta hiện nay, muốn sản xuất ra 1 tấn gang cần phải sử dụng 1,7 đến 1,8 tấn quặng sắt, 0,6 đến 0,7 tấn đá vôi làm chất trợ dung, 0,6 đến 0,8 tấn than cốc. Giả thiết trong đá vôi \ce{CaCO3} chiếm 97\% khối lượng và 90\% than cốc chuyển hoá thành \ce{CO2}. Trong điều kiện sản xuất như trên, khi sản xuất được 1 tấn gang, nhà máy đã thải ra môi trường tối thiểu bao nhiêu $\text{m}^3$ khí \ce{CO2} ở điều kiện chuẩn? (Làm tròn kết quả đến hàng đơn vị).
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}1260}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Để lượng khí \ce{CO2} thải ra tối thiểu, ta lấy lượng chất tiêu tốn ở mức thấp nhất:\\
-- Đá vôi tối thiểu: $m_{\text{đá vôi}} = 0{,}6\text{ tấn} = 600\text{ kg}$.\\
$\Rightarrow m_{\ce{CaCO3}} = 600 \times 97\% = 582\text{ kg} \Rightarrow n_{\ce{CaCO3}} = \frac{582}{100} = 5{,}82\text{ kmol}$.\\
Nhiệt phân đá vôi tạo: $n_1(\ce{CO2}) = 5{,}82\text{ kmol}$.\\
-- Than cốc tối thiểu: $m_{\text{than cốc}} = 0{,}6\text{ tấn} = 600\text{ kg}$.\\
Lượng C chuyển thành \ce{CO2}: $n_2(\ce{CO2}) = \frac{600 \times 90\%}{12} = 45\text{ kmol}$.\\
-- Tổng lượng \ce{CO2} tối thiểu sinh ra:
$$n(\ce{CO2}) = 5{,}82 + 45 = 50{,}82\text{ kmol.}$$
-- Thể tích khí \ce{CO2} ở điều kiện chuẩn ($25^\circ\text{C}, 1\text{ bar}$):
$$V(\ce{CO2}) = 50{,}82 \times 24{,}79 \approx 1259{,}83 \approx 1260\text{ m}^3.$$
\textbf{Đáp số: 1260.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Bằng phân tích quang phổ, người ta phát hiện trong “khí quyển” của sao Thuỷ có kim loại X. Biết X có cấu trúc mạng tinh thể lập phương tâm khối (độ đặc khít 68\%) và bán kính nguyên tử là $0{,}230\text{ nm}$. Khối lượng riêng của nguyên tố này là $0{,}862\text{ g/cm}^3$. (Biết Al = 27; K = 39; Zn = 65; Ba = 137). Xác định nguyên tử khối kim loại X.
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}39}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
-- Bán kính nguyên tử: $R = 0{,}230\text{ nm} = 2{,}30 \times 10^{-8}\text{ cm}$.\\
-- Thể tích 1 nguyên tử hình cầu:
$$V_1 = \frac{4}{3} \pi R^3 = \frac{4}{3} \pi (2{,}30 \times 10^{-8})^3 \approx 5{,}0965 \times 10^{-23}\text{ cm}^3.$$
-- Xét 1 mol nguyên tử X (chứa $N_A = 6{,}022 \times 10^{23}$ nguyên tử):
$$V_{\text{thực}} = V_1 \times N_A = 5{,}0965 \times 10^{-23} \times 6{,}022 \times 10^{23} \approx 30{,}691\text{ cm}^3.$$
-- Do độ đặc khít là 68\%, thể tích của 1 mol tinh thể là:
$$V_{\text{tinh thể}} = \frac{V_{\text{thực}}}{0{,}68} = \frac{30{,}691}{0{,}68} \approx 45{,}134\text{ cm}^3.$$
-- Khối lượng 1 mol tinh thể X:
$$m = V_{\text{tinh thể}} \times D = 45{,}134 \times 0{,}862 \approx 38{,}91\text{ g} \approx 39\text{ g.}$$
-- Nguyên tử khối kim loại X là 39 (Potassium / Kali, K).\\
\textbf{Đáp số: 39.}""" if is_teacher else "") + r"""}
\end{ex}

"""

    if is_teacher:
        tex += r"""
\Closesolutionfile{ans}

\vspace{0.5cm}
\begin{center}
	\textbf{\Large\color{mainGreen}BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ SỐ 17}
\end{center}

\vspace{0.2cm}
\noindent\textbf{\color{mainGreen}PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} & \textbf{17} & \textbf{18} \\ \hline
\textbf{\color{red}C} & \textbf{\color{red}B} & \textbf{\color{red}B} & \textbf{\color{red}D} & \textbf{\color{red}C} & \textbf{\color{red}D} & \textbf{\color{red}D} & \textbf{\color{red}D} & \textbf{\color{red}C} & \textbf{\color{red}D} & \textbf{\color{red}B} & \textbf{\color{red}A} & \textbf{\color{red}D} & \textbf{\color{red}C} & \textbf{\color{red}D} & \textbf{\color{red}A} & \textbf{\color{red}A} & \textbf{\color{red}C} \\ \hline
\end{tabular}
\end{center}

\vspace{0.3cm}
\noindent\textbf{\color{mainGreen}PHẦN II: TRẮC NGHIỆM ĐÚNG SAI}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Lệnh hỏi a} & \textbf{Lệnh hỏi b} & \textbf{Lệnh hỏi c} & \textbf{Lệnh hỏi d} \\ \hline
\textbf{1} & \textbf{\color{red}Đ} & \textbf{\color{red}Đ} & \textbf{\color{red}Đ} & \textbf{S} \\ \hline
\textbf{2} & \textbf{S} & \textbf{\color{red}Đ} & \textbf{S} & \textbf{\color{red}Đ} \\ \hline
\textbf{3} & \textbf{S} & \textbf{S} & \textbf{\color{red}Đ} & \textbf{S} \\ \hline
\textbf{4} & \textbf{\color{red}Đ} & \textbf{S} & \textbf{\color{red}Đ} & \textbf{S} \\ \hline
\end{tabular}
\end{center}

\vspace{0.3cm}
\noindent\textbf{\color{mainGreen}PHẦN III: TRẢ LỜI NGẮN}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\ \hline
\textbf{\color{red}4} & \textbf{\color{red}3} & \textbf{\color{red}2, 3, 4} & \textbf{\color{red}1, 2, 4} & \textbf{\color{red}1260} & \textbf{\color{red}39} \\ \hline
\end{tabular}
\end{center}
"""
    else:
        tex += r"""
\Closesolutionfile{ans}
"""

    tex += r"""
\end{document}
"""
    return tex

def compile_and_check(tex_filename, is_teacher=False):
    tex_path = os.path.join(OUT_DIR, tex_filename)
    pdf_filename = tex_filename.replace(".tex", ".pdf")
    pdf_path = os.path.join(OUT_DIR, pdf_filename)
    
    # Run pdflatex twice in San_Pham
    for i in range(2):
        proc = subprocess.run(
            ["pdflatex", "-interaction=nonstopmode", tex_filename],
            cwd=OUT_DIR,
            capture_output=True
        )
        stdout_str = proc.stdout.decode("utf-8", errors="replace") if proc.stdout else ""
        if proc.returncode != 0:
            print(f"Error compiling {tex_filename} (pass {i+1}, code {proc.returncode}):")
            lines = stdout_str.splitlines()
            err_lines = [l for l in lines if l.startswith("!") or "error" in l.lower() or "fatal" in l.lower()]
            print("Errors found:")
            print("\n".join(err_lines[:15]))
            print("Last 25 lines of log:")
            print("\n".join(lines[-25:]))
            return False
            
    # Check page count
    doc = fitz.open(pdf_path)
    page_count = len(doc)
    doc.close()
    print(f"Compiled {pdf_filename}: {page_count} pages.")
    
    # Clean temporary files
    for ext in ["aux", "log", "thm", "out"]:
        for f in glob.glob(os.path.join(OUT_DIR, f"*.{ext}")):
            try:
                os.remove(f)
            except:
                pass
    for f in glob.glob(os.path.join(OUT_DIR, "ans-*.tex")):
        try:
            os.remove(f)
        except:
            pass
            
    return page_count

if __name__ == "__main__":
    # Generate and compile Student version
    hs_tex_file = "De_Kiem_Tra_Hoa_12_De17_HocSinh.tex"
    hs_tex_content = create_latex_de17(is_teacher=False)
    with open(os.path.join(OUT_DIR, hs_tex_file), "w", encoding="utf-8") as f:
        f.write(hs_tex_content)
    print(f"Saved {hs_tex_file}")
    
    hs_pages = compile_and_check(hs_tex_file, is_teacher=False)
    print(f"Student version page count: {hs_pages} (Target: EXACTLY 4)")
    
    # Generate and compile Teacher version
    gv_tex_file = "De_Kiem_Tra_Hoa_12_De17_GiaoVien.tex"
    gv_tex_content = create_latex_de17(is_teacher=True)
    with open(os.path.join(OUT_DIR, gv_tex_file), "w", encoding="utf-8") as f:
        f.write(gv_tex_content)
    print(f"Saved {gv_tex_file}")
    
    gv_pages = compile_and_check(gv_tex_file, is_teacher=True)
    print(f"Teacher version page count: {gv_pages}")
