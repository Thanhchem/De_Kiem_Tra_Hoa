# -*- coding: utf-8 -*-
"""
Script sinh file LaTeX (.tex) và biên dịch ra PDF cho:
ĐỀ KIỂM TRA SỐ 19 - CHƯƠNG 7: KIM LOẠI NHÓM IA & NHÓM IIA (HÓA HỌC 12)
Mã đề thi: 119
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De19_HocSinh.pdf (Chuẩn chính xác 4 trang A4)
2. De_Kiem_Tra_Hoa_12_De19_GiaoVien.pdf (Bản Giáo Viên có lời giải chi tiết và đáp án)
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
import subprocess
import glob
import pymupdf

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"

def create_latex_de19(is_teacher=False):
    ans_file = "ans-de19-gv" if is_teacher else "ans-de19-hs"
    
    tex = r"""\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{geometry}
\geometry{top=0.65cm,bottom=0.65cm,left=1.25cm,right=1.25cm,headheight=20pt,headsep=3pt,footskip=9pt}
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
\setlength{\parskip}{""" + (r"1.0pt" if is_teacher else r"0.0pt") + r"""}

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
	\textbf{\color{mainGreen}ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 7 (KIM LOẠI NHÓM IA \& IIA)}\\
	\textbf{MÔN: HOÁ HỌC -- LỚP 12}\\
	\textit{Thời gian làm bài: 50 phút (không kể thời gian phát đề)}\\
	\textbf{Mã đề thi: 119}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Br = 80; Rb = 85,5; Ag = 108; Cs = 133; Ba = 137).}

\vspace{1pt}
\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\begin{ex}
Các đại dương là những kho muối vô tận với nhiều khoáng chất có giá trị dinh dưỡng cao. Trong nước biển, hai nguyên tố kim loại có nhiều nhất là
\choice
""" + (r"{\True sodium và magnesium.}" if is_teacher else r"{sodium và magnesium.}") + r"""
{đồng (copper) và kẽm (zinc).}
{nhôm (aluminium) và sắt (iron).}
{vàng (gold) và bạc (silver).}
\loigiai{""" + (r"""
Trong nước biển, hai cation kim loại dồi dào nhất là \ce{Na+} và \ce{Mg^2+}. Do đó, sodium và magnesium là hai nguyên tố kim loại phong phú nhất.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cấu hình electron lớp ngoài cùng của nguyên tử kim loại kiềm là
\choice
{$ns^2np^2$.}
{$ns^2np^1$.}
""" + (r"{\True $ns^1$.}" if is_teacher else r"{$ns^1$.}") + r"""
{$ns^2$.}
\loigiai{""" + (r"""
Các kim loại kiềm thuộc nhóm IA trong bảng tuần hoàn, nguyên tử có 1 electron lớp ngoài cùng dạng $ns^1$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Kim loại nào sau đây là kim loại kiềm?
\choice
""" + (r"{\True Na.}" if is_teacher else r"{Na.}") + r"""
{Ca.}
{Al.}
{Fe.}
\loigiai{""" + (r"""
Sodium (Na) thuộc nhóm IA nên là kim loại kiềm. Ca thuộc nhóm IIA (kiềm thổ), Al thuộc nhóm IIIA, Fe thuộc nhóm VIIIB.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho một mẩu Na vào dung dịch \ce{MgSO4} dư, thu được kết tủa X và chất khí Y. Hai chất X, Y lần lượt là
\choice
{Mg và \ce{H2}.}
""" + (r"{\True \ce{Mg(OH)2} và \ce{H2}.}" if is_teacher else r"{\ce{Mg(OH)2} và \ce{H2}.}") + r"""
{Mg và \ce{O2}.}
{\ce{Na2SO4} và \ce{H2}.}
\loigiai{""" + (r"""
Na phản ứng với nước trong dung dịch: $\ce{2Na + 2H2O -> 2NaOH + H2 ^}$ (khí Y là \ce{H2}).\\
Sau đó xảy ra phản ứng trao đổi: $\ce{2NaOH + MgSO4 -> Mg(OH)2 v + Na2SO4}$ (kết tủa X là \ce{Mg(OH)2}).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong cốc nước chứa nhiều các ion sau: \ce{Ca^2+}, \ce{Mg^2+}, \ce{Cl-}, \ce{SO4^2-}. Nước trong cốc trên thuộc loại
\choice
""" + (r"{\True có tính cứng vĩnh cửu.}" if is_teacher else r"{có tính cứng vĩnh cửu.}") + r"""
{không có tính cứng.}
{có tính cứng tạm thời.}
{có tính cứng toàn phần.}
\loigiai{""" + (r"""
Nước cứng chứa ion \ce{Ca^2+}, \ce{Mg^2+} cùng với \ce{Cl-}, \ce{SO4^2-} (không có \ce{HCO3-}) là nước có tính cứng vĩnh cửu.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Nhận định nào \textbf{không} đúng về cấu tạo và tính chất của các kim loại nhóm IIA?
\choice
{Khối lượng riêng tương đối nhỏ, là những kim loại nhẹ hơn nhôm (trừ Ba).}
{Độ cứng cao hơn kim loại nhóm IA, nhưng mềm hơn kim loại nhôm.}
""" + (r"{\True Mạng tinh thể của chúng đều có kiểu lập phương tâm khối.}" if is_teacher else r"{Mạng tinh thể của chúng đều có kiểu lập phương tâm khối.}") + r"""
{Nhiệt độ nóng chảy và nhiệt độ sôi tương đối thấp (trừ Be).}
\loigiai{""" + (r"""
Kim loại nhóm IIA có các kiểu mạng tinh thể khác nhau: Be và Mg có mạng lục phương; Ca và Sr có mạng lập phương tâm diện; Ba có mạng lập phương tâm khối.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phương trình hóa học của phản ứng nào sau đây \textbf{không} đúng?
\choice
{$\ce{Ca + 2H2O -> Ca(OH)2 + H2}$.}
""" + (r"{\True $\ce{BeO + H2O -> Be(OH)2}$.}" if is_teacher else r"{$\ce{BeO + H2O -> Be(OH)2}$.}") + r"""
{$\ce{Mg + H2SO4 -> MgSO4 + H2}$.}
{$\ce{Be + 2NaOH -> Na2BeO2 + H2}$.}
\loigiai{""" + (r"""
Be và BeO hoàn toàn không tác dụng với nước ở bất kì điều kiện nhiệt độ nào (kể cả ở nhiệt độ cao đỏ).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cách nào sau đây \textbf{không} thu được NaOH sau phản ứng?
\choice
""" + (r"{\True Điện phân dung dịch \ce{NaCl} không có màng ngăn xốp.}" if is_teacher else r"{Điện phân dung dịch \ce{NaCl} không có màng ngăn xốp.}") + r"""
{Cho kim loại Na tác dụng với nước.}
{Điện phân dung dịch \ce{NaCl} có màng ngăn xốp, điện cực trơ.}
{Cho dung dịch \ce{Ca(OH)2} tác dụng với dung dịch \ce{Na2CO3}.}
\loigiai{""" + (r"""
Khi điện phân dd \ce{NaCl} không có màng ngăn, \ce{Cl2} sinh ra ở anode phản ứng ngay với \ce{NaOH} sinh ra ở cathode tạo thành nước Javel ($\ce{NaCl + NaClO + H2O}$), không thu được NaOH.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hai chất nào sau đây được dùng để làm mềm nước có tính cứng vĩnh cửu?
\choice
{\ce{NaCl} và \ce{Ca(OH)2}.}
""" + (r"{\True \ce{Na2CO3} và \ce{Na3PO4}.}" if is_teacher else r"{\ce{Na2CO3} và \ce{Na3PO4}.}") + r"""
{\ce{Na2CO3} và \ce{Ca(OH)2}.}
{\ce{Na2CO3} và \ce{HCl}.}
\loigiai{""" + (r"""
Để làm mềm nước cứng vĩnh cửu (hoặc toàn phần), người ta dùng \ce{Na2CO3} hoặc \ce{Na3PO4} để kết tủa \ce{Ca^2+} và \ce{Mg^2+} dưới dạng carbonate hoặc phosphate không tan.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\begin{ex}
Hợp chất Y của calcium là thành phần chính của vỏ các loại ốc, sò,... Ngoài ra Y được sử dụng rộng rãi trong sản xuất vôi, xi măng, thuỷ tinh, chất phụ gia trong công nghiệp thực phẩm. Hợp chất Y là
\choice
{\ce{CaO}.}
""" + (r"{\True \ce{CaCO3}.}" if is_teacher else r"{\ce{CaCO3}.}") + r"""
{\ce{Ca(OH)2}.}
{\ce{Ca3(PO4)2}.}
\loigiai{""" + (r"""
Thành phần chính của vỏ các loại sò, ốc, hến, đá vôi, đá phấn, cẩm thạch là calcium carbonate (\ce{CaCO3}).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Kim loại \textbf{không} phản ứng với nước ở nhiệt độ thường là
\choice
""" + (r"{\True Be.}" if is_teacher else r"{Be.}") + r"""
{Ca.}
{Li.}
{K.}
\loigiai{""" + (r"""
Beryllium (Be) không phản ứng với nước ở nhiệt độ thường cũng như khi đun nóng.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Xét phản ứng phân huỷ muối carbonate của kim loại nhóm IIA: $\ce{MCO3(s) -> MO(s) + CO2(g)} \quad (\Delta_r H^\circ_{298})$.\\
Từ \ce{MgCO3} đến \ce{BaCO3}, biến thiên enthalpy chuẩn của phản ứng biến đổi như thế nào?
\choice
{Không đổi.}
{Giảm dần.}
""" + (r"{\True Tăng dần.}" if is_teacher else r"{Tăng dần.}") + r"""
{Không có quy luật.}
\loigiai{""" + (r"""
Độ bền nhiệt của các muối carbonate tăng dần từ \ce{MgCO3} đến \ce{BaCO3}, năng lượng cần cung cấp để phá vỡ liên kết lớn hơn nên phản ứng thu nhiệt nhiều hơn $\rightarrow \Delta_r H^\circ_{298}$ tăng dần.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Ở nhiệt độ thường, độ tan của các hydroxide tăng dần trong dãy từ \ce{Mg(OH)2} đến \ce{Ba(OH)2}. Từ thông tin này có thể dự đoán được khả năng phản ứng với nước của các kim loại từ Mg đến Ba biến đổi như thế nào?
\choice
""" + (r"{\True Tăng dần.}" if is_teacher else r"{Tăng dần.}") + r"""
{Không đổi.}
{Không có quy luật.}
{Giảm dần.}
\loigiai{""" + (r"""
Sản phẩm hydroxide tạo thành càng dễ tan thì càng nhanh chóng khuếch tán vào dung dịch, giải phóng bề mặt kim loại để tiếp tục tiếp xúc với nước $\rightarrow$ Khả năng phản ứng với nước tăng dần từ Mg đến Ba.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Nước chứa nhiều các ion nào sau đây có tính cứng toàn phần?
\choice
""" + (r"{\True \ce{Mg^2+}, \ce{Ca^2+}, \ce{HCO3-}, \ce{SO4^2-}, \ce{Cl-}.}" if is_teacher else r"{\ce{Mg^2+}, \ce{Ca^2+}, \ce{HCO3-}, \ce{SO4^2-}, \ce{Cl-}.}") + r"""
{\ce{Na+}, \ce{K+}, \ce{SO4^2-}, \ce{Cl-}.}
{\ce{Mg^2+}, \ce{Ca^2+}, \ce{HCO3-}.}
{\ce{Mg^2+}, \ce{Ca^2+}, \ce{SO4^2-}, \ce{Cl-}.}
\loigiai{""" + (r"""
Nước có tính cứng toàn phần chứa đồng thời tính cứng tạm thời (chứa \ce{HCO3-}) và tính cứng vĩnh cửu (chứa \ce{Cl-}, \ce{SO4^2-}) cùng với các cation \ce{Ca^2+}, \ce{Mg^2+}.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phân tích một mẫu nước tự nhiên thấy chứa nhiều các ion: \ce{Na+}, \ce{Ca^2+}, \ce{HCO3-}, \ce{Cl-} và \ce{SO4^2-}. Chất nào sau đây có thể làm mềm mẫu nước trên?
\choice
""" + (r"{\True \ce{Na2CO3}.}" if is_teacher else r"{\ce{Na2CO3}.}") + r"""
{\ce{Ca(OH)2}.}
{\ce{NaOH}.}
{\ce{HCl}.}
\loigiai{""" + (r"""
Mẫu nước là nước cứng toàn phần. Dùng soda \ce{Na2CO3} làm kết tủa hoàn toàn \ce{Ca^2+}: $\ce{Ca^2+ + CO3^2- -> CaCO3 v}$.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Một loại nước cứng khi đun sôi thì trở thành nước mềm. Trong loại nước này có hoà tan những hợp chất nào sau đây?
\choice
""" + (r"{\True \ce{Ca(HCO3)2} và \ce{Mg(HCO3)2}.}" if is_teacher else r"{\ce{Ca(HCO3)2} và \ce{Mg(HCO3)2}.}") + r"""
{\ce{Ca(HCO3)2} và \ce{MgSO4}.}
{\ce{CaSO4} và \ce{MgCl2}.}
{\ce{MgCl2} và \ce{CaCl2}.}
\loigiai{""" + (r"""
Nước mất hoàn toàn tính cứng khi đun sôi là nước cứng tạm thời, chứa các muối hydrogencarbonate \ce{Ca(HCO3)2} và \ce{Mg(HCO3)2}. Khi đun sôi các muối này bị nhiệt phân tạo kết tủa carbonate.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Tiến hành thí nghiệm theo các bước sau:\\
\textbf{Bước 1:} Chuẩn bị hai ống nghiệm, ống (1) chứa 2 mL dung dịch \ce{CaCl2} 1 M, ống (2) chứa 2 mL dung dịch \ce{BaCl2} 1 M.\\
\textbf{Bước 2:} Nhỏ đồng thời vào mỗi ống nghiệm 3 giọt dung dịch \ce{CuSO4} 1 M, thấy ống (1) xuất hiện kết tủa chậm hơn và ít hơn so với ống (2).\\
Nhận định nào sau đây đúng khi so sánh \ce{CaSO4} với \ce{BaSO4}?
\choice
{Khó nhiệt phân hơn.}
{Khó thuỷ phân hơn.}
{Dễ kết tủa hơn.}
""" + (r"{\True Dễ tan hơn.}" if is_teacher else r"{Dễ tan hơn.}") + r"""
\loigiai{""" + (r"""
Ống (1) tạo \ce{CaSO4}, ống (2) tạo \ce{BaSO4}. Vì ống (1) xuất hiện kết tủa chậm hơn và ít hơn nhiều so với ống (2) nên \ce{CaSO4} có độ tan lớn hơn (dễ tan hơn) so với \ce{BaSO4}.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho sơ đồ chuyển hoá sau:\\
$\ce{CaO} \xrightarrow{+\text{X}} \text{Y} \xrightarrow{+\text{Z}} \ce{CaCO3} \xrightarrow{+\text{Z} + \text{X}} \text{T} \xrightarrow{+\text{E}} \ce{CaSO4}$\\
Biết: X, Y, Z, T, E là các hợp chất khác nhau; mỗi mũi tên ứng với một phương trình hoá học. Các chất Z, E thoả mãn sơ đồ trên lần lượt là
\choice
{\ce{Na2CO3}, \ce{H2SO4}.}
""" + (r"{\True \ce{CO2}, \ce{KHSO4}.}" if is_teacher else r"{\ce{CO2}, \ce{KHSO4}.}") + r"""
{\ce{NaHCO3}, \ce{Na2SO4}.}
{\ce{CO2}, \ce{BaSO4}.}
\loigiai{""" + (r"""
X: \ce{H2O}; Y: \ce{Ca(OH)2}; Z: \ce{CO2}; T: \ce{Ca(HCO3)2}; E: \ce{KHSO4}.\\
$\ce{Ca(HCO3)2 + 2KHSO4 -> CaSO4 v + K2SO4 + 2CO2 ^ + 2H2O}$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\setcounter{ex}{0}

\begin{ex}
Quặng sylvinite là một khoáng chất phổ biến có thành phần chính là \ce{NaCl.KCl}. Sự phụ thuộc của độ tan các muối vào nhiệt độ được biểu diễn ở đồ thị sau:
\par\vspace{1pt}\centerline{\includegraphics[height=4.6cm]{images_crop_de19/p2_cau1_do_thi_do_tan.png}}\vspace{1pt}
\choiceTF
{Độ tan của \ce{KCl} giảm chậm khi giảm nhiệt độ từ $100^\circ\text{C}$ về $0^\circ\text{C}$.}
""" + (r"{\True Tách được \ce{KCl} khỏi dung dịch \ce{NaCl} bằng phương pháp kết tinh phân đoạn.}" if is_teacher else r"{Tách được \ce{KCl} khỏi dung dịch \ce{NaCl} bằng phương pháp kết tinh phân đoạn.}") + r"""
{Độ tan của \ce{NaCl} tăng nhanh khi nhiệt độ từ $0^\circ\text{C}$ đến $100^\circ\text{C}$.}
""" + (r"{\True Độ tan của \ce{KCl} giảm nhanh hơn của \ce{NaCl} khi giảm nhiệt độ từ $100^\circ\text{C}$ về $0^\circ\text{C}$.}" if is_teacher else r"{Độ tan của \ce{KCl} giảm nhanh hơn của \ce{NaCl} khi giảm nhiệt độ từ $100^\circ\text{C}$ về $0^\circ\text{C}$.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Sai}. Đồ thị độ tan của \ce{KCl} rất dốc, khi nhiệt độ giảm từ $100^\circ\text{C}$ về $0^\circ\text{C}$ thì độ tan giảm rất nhanh.\\
b) \textbf{Đúng}. Độ dốc của hai đường tan chênh lệch rất lớn nên dễ dàng tách \ce{KCl} bằng phương pháp kết tinh.\\
c) \textbf{Sai}. Đường độ tan của \ce{NaCl} gần như nằm ngang, biến thiên rất ít theo nhiệt độ.\\
d) \textbf{Đúng}. Độ dốc của đường biểu diễn \ce{KCl} lớn hơn nhiều so với \ce{NaCl}.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Mỗi phát biểu sau đây về kim loại kiềm là đúng hay sai?
\choiceTF
""" + (r"{\True Thứ tự giảm dần tính kim loại của các kim loại kiềm là: Cs, Rb, K, Na, Li.}" if is_teacher else r"{Thứ tự giảm dần tính kim loại của các kim loại kiềm là: Cs, Rb, K, Na, Li.}") + r"""
{Phương pháp chung để điều chế kim loại kiềm là điện phân dung dịch.}
{Để bảo quản kim loại Na cần ngâm Na trong cồn tinh khiết.}
{\ce{Na2O} tan trong nước tạo dung dịch trong suốt và thoát ra khí \ce{H2}.}
\loigiai{""" + (r"""
a) \textbf{Đúng}. Trong nhóm IA theo chiều tăng Z, tính kim loại tăng dần: $\ce{Li < Na < K < Rb < Cs}$, do đó giảm dần là $\ce{Cs, Rb, K, Na, Li}$.\\
b) \textbf{Sai}. Kim loại kiềm có tính khử rất mạnh, ion kim loại kiềm khó bị khử nên chỉ điều chế bằng phương pháp điện phân nóng chảy.\\
c) \textbf{Sai}. Na phản ứng với cồn sinh khí \ce{H2}, Na phải được bảo quản ngâm trong dầu hoả.\\
d) \textbf{Sai}. $\ce{Na2O + H2O -> 2NaOH}$, phản ứng không sinh ra khí \ce{H2}.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các mẫu nước được đánh dấu X, Y, Z, T có chứa các ion theo bảng sau:
\par\vspace{1pt}{\centering
\small\renewcommand{\arraystretch}{0.92}
\begin{tabular}{|c|l|}
\hline
\textbf{Mẫu nước} & \textbf{Các ion chủ yếu có trong mẫu nước} \\
\hline
X & \ce{Ca^2+}, \ce{Mg^2+} (nhiều), \ce{Cl-}, \ce{SO4^2-} \\
Y & \ce{Ca^2+}, \ce{Mg^2+} (nhiều), \ce{HCO3-} \\
Z & \ce{Na+}, \ce{K+}, \ce{HCO3-} \\
T & \ce{Ca^2+}, \ce{Mg^2+} (nhiều), \ce{Cl-}, \ce{SO4^2-}, \ce{HCO3-} \\
\hline
\end{tabular}\par}\vspace{1pt}
\choiceTF
{Nước X và Y đều là nước có tính cứng vĩnh cửu.}
{Có thể loại bỏ tính cứng của nước T bằng cách đun nóng.}
""" + (r"{\True Có thể loại bỏ tính cứng của nước Y bằng dung dịch \ce{Na2CO3} dư.}" if is_teacher else r"{Có thể loại bỏ tính cứng của nước Y bằng dung dịch \ce{Na2CO3} dư.}") + r"""
""" + (r"{\True Có thể phân biệt X, Y, Z và T bằng cách đun nóng nhẹ và dùng dung dịch \ce{NaOH}.}" if is_teacher else r"{Có thể phân biệt X, Y, Z và T bằng cách đun nóng nhẹ và dùng dung dịch \ce{NaOH}.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Sai}. X là nước cứng vĩnh cửu, Y là nước cứng tạm thời.\\
b) \textbf{Sai}. T là nước cứng toàn phần; đun nóng chỉ loại bỏ tính cứng tạm thời, phần vĩnh cửu (\ce{Cl-}, \ce{SO4^2-}) vẫn còn.\\
c) \textbf{Đúng}. \ce{Na2CO3} cung cấp ion \ce{CO3^2-} kết tủa \ce{Ca^2+}, \ce{Mg^2+} làm mềm nước Y.\\
d) \textbf{Đúng}. Đun nhẹ: Y và T xuất hiện kết tủa trắng, X và Z không. Dùng \ce{NaOH} thử tiếp: phân biệt được từng mẫu riêng biệt.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hợp chất của kim loại calcium có nhiều ứng dụng quan trọng trong thực tiễn đời sống và sản xuất.
\choiceTF
""" + (r"{\True Hợp chất của calcium được dùng để đúc tượng, bó bột khi gãy xương là thạch cao nung (\ce{CaSO4.H2O}).}" if is_teacher else r"{Hợp chất của calcium được dùng để đúc tượng, bó bột khi gãy xương là thạch cao nung (\ce{CaSO4.H2O}).}") + r"""
""" + (r"{\True Vôi sống (\ce{CaO}) được dùng phổ biến để khử chua cho đất trồng trọt.}" if is_teacher else r"{Vôi sống (\ce{CaO}) được dùng phổ biến để khử chua cho đất trồng trọt.}") + r"""
{Phương trình hoá học của phản ứng tạo thạch nhũ trong các hang động đá vôi là: $\ce{CaCO3 + CO2 + H2O -> Ca(HCO3)2}$.}
""" + (r"{\True Thành phần hoá học chính của vỏ các loài ốc, sò, hến là calcium carbonate (\ce{CaCO3}).}" if is_teacher else r"{Thành phần hoá học chính của vỏ các loài ốc, sò, hến là calcium carbonate (\ce{CaCO3}).}") + r"""
\loigiai{""" + (r"""
a) \textbf{Đúng}. Thạch cao nung nở thể tích khi đông cứng nên dùng đúc tượng, bó bột.\\
b) \textbf{Đúng}. Vôi sống trung hoà acid trong đất phèn khử chua.\\
c) \textbf{Sai}. Phản ứng tạo thạch nhũ là: $\ce{Ca(HCO3)2 -> CaCO3 v + CO2 ^ + H2O}$. Chiều ngược lại là phản ứng xâm thực đá vôi.\\
d) \textbf{Đúng}. Vỏ ốc, sò, hến chứa trên 95\% là \ce{CaCO3}.""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}

\setcounter{ex}{0}

\begin{ex}
Cho các kim loại sau: Li, Na, Ca, Mg, Fe, Cu, K, Zn, Ag, Rb. Có bao nhiêu kim loại kiềm trong dãy trên?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}4}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Có 4 kim loại kiềm (nhóm IA) là: Li, Na, K, Rb.\\
\textbf{Đáp số: 4.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho sơ đồ các phản ứng sau theo đúng tỉ lệ mol:\\
(1) $\ce{2X1 + 2X2 -> 2X3 + H2}$\\
(2) $\ce{X3 + CO2 -> X4}$\\
(3) $\ce{X3 + X4 -> X5 + X2}$\\
(4) $\ce{2X6 + 3X5 + 3X2 -> 2Fe(OH)3 + 3CO2 + 6KCl}$\\
Khối lượng mol của chất $\text{X}_5$ là bao nhiêu g/mol?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}138}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Từ phản ứng (4): cation kim loại là K. Các chất lần lượt là:\\
$\text{X}_1$: K; $\text{X}_2$: \ce{H2O}; $\text{X}_3$: \ce{KOH}; $\text{X}_4$: \ce{KHCO3}; $\text{X}_5$: \ce{K2CO3}; $\text{X}_6$: \ce{FeCl3}.\\
Khối lượng mol của $\text{X}_5$ (\ce{K2CO3}): $M = 39 \times 2 + 12 + 16 \times 3 = 138\text{ g/mol}$.\\
\textbf{Đáp số: 138.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các phát biểu sau:\\
(1) Soda là chất bột màu trắng, tan trong nước tạo môi trường trung tính.\\
(2) Soda dùng làm mềm nước cứng.\\
(3) Soda bền với nhiệt hơn so với baking soda.\\
(4) Chất béo bị thuỷ phân trong dung dịch soda tạo thành xà phòng.\\
(5) Có thể dùng baking soda thay cho soda trong việc tẩy rửa lớp dầu, mỡ bám vào bồn cầu.\\
Có bao nhiêu phát biểu đúng trong các phát biểu trên?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}3}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Có 3 phát biểu đúng là (2), (3), (4):\\
-- (1) Sai vì dung dịch soda \ce{Na2CO3} có tính kiềm (pH > 7).\\
-- (2) Đúng vì ion \ce{CO3^2-} kết tủa \ce{Ca^2+}, \ce{Mg^2+}.\\
-- (3) Đúng vì \ce{Na2CO3} không bị nhiệt phân, \ce{NaHCO3} dễ bị nhiệt phân.\\
-- (4) Đúng vì môi trường kiềm xà phòng hoá chất béo.\\
-- (5) Sai vì baking soda có tính kiềm yếu hơn nhiều so với soda, tẩy nhờn kém.\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong số các chất sau: \ce{Ca(OH)2}, \ce{Na2CO3}, \ce{Na3PO4}, \ce{KCl}, có bao nhiêu chất có thể dùng để làm mềm nước có tính cứng tạm thời?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}3}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Có 3 chất là \ce{Ca(OH)2} (lượng vừa đủ), \ce{Na2CO3} và \ce{Na3PO4}.\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong công nghiệp, người ta sản xuất xút (\ce{NaOH}) từ muối ăn bằng phương pháp điện phân dung dịch có màng ngăn xốp. Tính khối lượng (tấn) \ce{NaCl} cần dùng để sản xuất 15 tấn \ce{NaOH} với hiệu suất của toàn bộ quá trình là 80\% (kết quả làm tròn đến hàng phần mười).
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}27,4}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
$\ce{2NaCl + 2H2O -> 2NaOH + Cl2 + H2}$\\
$m_{\ce{NaCl}\text{ (lý thuyết)}} = \dfrac{15 \times 58{,}5}{40} = 21{,}9375\text{ tấn}$.\\
Với hiệu suất $H = 80\%$, khối lượng thực tế cần dùng là:\\
$m_{\ce{NaCl}\text{ (thực tế)}} = \dfrac{21{,}9375}{0{,}80} = 27{,}421875\text{ tấn} \approx 27{,}4\text{ tấn}$.\\
\textbf{Đáp số: 27,4.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Carnallite là muối khoáng kép ngậm nước có công thức dạng $\ce{KCl.MgCl2.nH2O}$. Khi nung nóng 5,55 gam carnallite, thu được 3,39 gam muối khan. Mặt khác, nếu cho 5,55 gam carnallite tác dụng hoàn toàn với dung dịch \ce{NaOH} dư, thu được kết tủa. Lọc và nung kết tủa đến khối lượng không đổi thì khối lượng chất rắn giảm 0,36 gam so với khối lượng kết tủa ban đầu. Xác định giá trị của $n$.
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}6}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
-- Khối lượng nước kết tinh bay ra khi nung carnallite:\\
$m_{\ce{H2O}} = 5{,}55 - 3{,}39 = 2{,}16\text{ g} \rightarrow n_{\ce{H2O}} = \dfrac{2{,}16}{18} = 0{,}12\text{ mol}$.\\
-- Nung kết tủa $\ce{Mg(OH)2 -> MgO + H2O}$. Độ giảm khối lượng 0,36 g là nước từ $\ce{Mg(OH)2}$:\\
$n_{\ce{Mg(OH)2}} = \dfrac{0{,}36}{18} = 0{,}02\text{ mol} \rightarrow n_{\ce{MgCl2}} = 0{,}02\text{ mol}$.\\
-- Giá trị của $n$: $n = \dfrac{n_{\ce{H2O}}}{n_{\ce{MgCl2}}} = \dfrac{0{,}12}{0{,}02} = 6$.\\
\textbf{Đáp số: 6.}""" if is_teacher else "") + r"""}
\end{ex}
"""

    if is_teacher:
        tex += r"""
\vspace{0.3cm}
\noindent
\textbf{\color{hfRed}\large BẢNG ĐÁP ÁN THAM KHẢO TOÀN ĐỀ}

\vspace{0.1cm}
\noindent
\textbf{\color{mainGreen}PHẦN I:} 1.A \quad 2.C \quad 3.A \quad 4.B \quad 5.A \quad 6.C \quad 7.B \quad 8.A \quad 9.B \quad 10.B \quad 11.A \quad 12.C \quad 13.A \quad 14.A \quad 15.A \quad 16.A \quad 17.D \quad 18.B

\vspace{0.1cm}
\noindent
\textbf{\color{mainGreen}PHẦN II:} Câu 1: a-S, b-Đ, c-S, d-Đ \quad|\quad Câu 2: a-Đ, b-S, c-S, d-S\\
Câu 3: a-S, b-S, c-Đ, d-Đ \quad|\quad Câu 4: a-Đ, b-Đ, c-S, d-Đ

\vspace{0.1cm}
\noindent
\textbf{\color{mainGreen}PHẦN III:} Câu 1: \textbf{4} \quad|\quad Câu 2: \textbf{138} \quad|\quad Câu 3: \textbf{3} \quad|\quad Câu 4: \textbf{3} \quad|\quad Câu 5: \textbf{27,4} \quad|\quad Câu 6: \textbf{6}
"""

    tex += r"""
\Closesolutionfile{ans}
\end{document}
"""
    return tex

def compile_latex(tex_content, tex_filename, pdf_filename):
    tex_path = os.path.join(OUT_DIR, tex_filename)
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_content)
    
    # Biên dịch pdflatex 2 lần để giải quyết tham chiếu & ex_test
    for i in range(2):
        res = subprocess.run(
            ["pdflatex", "-interaction=nonstopmode", tex_filename],
            cwd=OUT_DIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
    
    # Dọn dẹp file rác
    for ext in ["aux", "log", "thm", "out"]:
        for garbage in glob.glob(os.path.join(OUT_DIR, f"*.{ext}")):
            try:
                os.remove(garbage)
            except:
                pass
    for ans_f in glob.glob(os.path.join(OUT_DIR, "ans-*.tex")):
        try:
            os.remove(ans_f)
        except:
            pass

    pdf_path = os.path.join(OUT_DIR, pdf_filename)
    if os.path.exists(pdf_path):
        doc = pymupdf.open(pdf_path)
        page_count = len(doc)
        print(f"Biên dịch thành công: {pdf_filename} ({page_count} trang)")
        return page_count
    else:
        print(f"LỖI: Không tìm thấy {pdf_filename}")
        return 0

def main():
    print("--- Đang tạo file LaTeX và PDF cho ĐỀ SỐ 19 ---")
    
    # 1. Bản Học Sinh
    tex_hs = create_latex_de19(is_teacher=False)
    pages_hs = compile_latex(tex_hs, "De_Kiem_Tra_Hoa_12_De19_HocSinh.tex", "De_Kiem_Tra_Hoa_12_De19_HocSinh.pdf")
    
    # 2. Bản Giáo Viên
    tex_gv = create_latex_de19(is_teacher=True)
    pages_gv = compile_latex(tex_gv, "De_Kiem_Tra_Hoa_12_De19_GiaoVien.tex", "De_Kiem_Tra_Hoa_12_De19_GiaoVien.pdf")
    
    print("\n--- KẾT QUẢ BIÊN DỊCH ---")
    print(f"Học sinh: {pages_hs} trang (Yêu cầu chuẩn: 4 trang)")
    print(f"Giáo viên: {pages_gv} trang")
    
    if pages_hs == 4:
        print("XÁC NHẬN: Bản Học Sinh ĐẠT CHUẨN CHÍNH XÁC 4 TRANG A4!")
    else:
        print(f"CẢNH BÁO: Bản Học Sinh có {pages_hs} trang, cần điều chỉnh!")

if __name__ == "__main__":
    main()
