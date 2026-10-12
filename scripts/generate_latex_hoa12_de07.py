# -*- coding: utf-8 -*-
"""
Script sinh file LaTeX (.tex) và biên dịch ra PDF cho:
ĐỀ KIỂM TRA SỐ 7 - CHƯƠNG 3: HỢP CHẤT CHỨA NITROGEN (HÓA HỌC 12)
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De07_HocSinh.pdf (Đúng chuẩn 4 trang A4)
2. De_Kiem_Tra_Hoa_12_De07_GiaoVien.pdf (Bản Giáo Viên có lời giải chi tiết và đáp án)
"""

import os
import subprocess
import glob

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"

def create_latex_de07(is_teacher=False):
    ans_file = "ans-de07-gv" if is_teacher else "ans-de07-hs"
    
    tex = r"""\documentclass[11pt,a4paper]{article}
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
\usepackage{multicol}
\usepackage{fancyhdr}
\usepackage{enumitem}

\definecolor{hfRed}{HTML}{D32F2F}
\definecolor{hfGreen}{HTML}{8BC390}
\definecolor{mainGreen}{HTML}{2E7D32}

\usepackage{ex_test}
\bracketExpl{true}

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
\setlength{\parskip}{""" + (r"1.5pt" if is_teacher else r"2pt") + r"""}

\def\loigiai#1{""" + (r"""
\par\begingroup\color{blue!80!black}\small
\noindent\textbf{\textit{Lời giải chi tiết:}}\par
#1\par\endgroup""" if is_teacher else "") + r"""}

\begin{document}
"""

    if is_teacher:
        tex += r"""
\Opensolutionfile{ans}[""" + ans_file + r"""]
"""
    else:
        tex += r"""
\Opensolutionfile{ans}[""" + ans_file + r"""]
"""

    # First page header
    tex += r"""
\noindent
\begin{minipage}[t]{0.48\textwidth}
	\centering
	\textbf{SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ}\\
	\textbf{ĐỀ CHÍNH THỨC}\\
	\textit{(Đề thi có 04 trang)}
\end{minipage}\hfill
\begin{minipage}[t]{0.48\textwidth}
	\centering
	\textbf{\color{mainGreen}ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 3 (HỢP CHẤT NITROGEN)}\\
	\textbf{MÔN: HOÁ HỌC -- LỚP 12}\\
	\textit{Thời gian làm bài: 50 phút (không kể thời gian phát đề)}\\
	\textbf{Mã đề thi: 107}
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
\vspace{0.1cm}
\noindent\textit{Họ, tên thí sinh: \dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots Lớp: \dots\dots\dots\dots SBD: \dots\dots\dots\dots}\\
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; S = 32; Cl = 35,5; Fe = 56; Cu = 64).}

\vspace{0.15cm}
\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\setcounter{ex}{0}

\begin{ex}
Chất nào sau đây là amine thơm?
\choice
""" + (r"{\True Aniline}" if is_teacher else r"{Aniline}") + r"""
{Cyclohexylamine}
{Alanine}
{Trimethylamine}
\loigiai{""" + (r"""
Aniline (\ce{C6H5NH2}) có nhóm amino -\ce{NH2} liên kết trực tiếp với vòng benzene nên là một amine thơm.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Aniline có công thức phân tử là:
\choice
{\ce{C3H7O2N}}
{\ce{C2H5O2N}}
{\ce{C7H9N}}
""" + (r"{\True \ce{C6H7N}}" if is_teacher else r"{\ce{C6H7N}}") + r"""
\loigiai{""" + (r"""
Aniline có công thức cấu tạo là \ce{C6H5NH2} => Công thức phân tử là \ce{C6H7N}.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phân tử amine có cấu tạo tương tự phân tử nào sau đây nhất?
\choice
{\ce{H2O}}
""" + (r"{\True \ce{NH3}}" if is_teacher else r"{\ce{NH3}}") + r"""
{\ce{HF}}
{\ce{N2}}
\loigiai{""" + (r"""
Amine là dẫn xuất của ammonia (\ce{NH3}) khi thay thế nguyên tử H bằng gốc hydrocarbon. Trong phân tử amine, nguyên tử N vẫn còn một cặp electron chưa liên kết tương tự trong \ce{NH3}.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Số nhóm amino (-\ce{NH2}) có trong một phân tử Lysine là
\choice
{1}
{3}
{4}
""" + (r"{\True 2}" if is_teacher else r"{2}") + r"""
\loigiai{""" + (r"""
Lysine (acid 2,6-diaminohexanoic) có công thức là \ce{H2N-[CH2]4-CH(NH2)-COOH}, phân tử chứa 2 nhóm -\ce{NH2} và 1 nhóm -\ce{COOH}.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- CTST). Ở điều kiện thường, trạng thái tồn tại của amino acid là
\choice
{thể lỏng}
{thể khí}
""" + (r"{\True thể rắn}" if is_teacher else r"{thể rắn}") + r"""
{thể rắn và lỏng}
\loigiai{""" + (r"""
Do tồn tại chủ yếu ở dạng ion lưỡng cực có lực hút tĩnh điện liên phân tử rất mạnh, các amino acid ở điều kiện thường đều là chất rắn kết tinh, nhiệt độ nóng chảy cao.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Khi đun nóng, các $\varepsilon$-amino acid hoặc $\omega$-amino acid có thể phản ứng với nhau để tạo thành polymer, đồng thời tách ra các phân tử nước gọi là phản ứng
\choice
{ester hóa}
""" + (r"{\True trùng ngưng}" if is_teacher else r"{trùng ngưng}") + r"""
{trùng hợp}
{hydrogen hóa}
\loigiai{""" + (r"""
Phản ứng kết hợp nhiều phân tử nhỏ tạo polymer đồng thời giải phóng phân tử nhỏ khác (như \ce{H2O}) là phản ứng trùng ngưng.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Peptide: \ce{H2NCH2CONHCH(CH3)CONHCH2COOH} được biểu diễn dạng kí hiệu như sau
\choice
{Gly -- Gly -- Ala}
""" + (r"{\True Gly -- Ala -- Gly}" if is_teacher else r"{Gly -- Ala -- Gly}") + r"""
{Ala -- Gly -- Gly}
{Ala -- Gly -- Ala}
\loigiai{""" + (r"""
Các gốc amino acid lần lượt từ đầu N đến đầu C là: \ce{H2N-CH2-CO-} (Gly), \ce{-NH-CH(CH3)-CO-} (Ala), \ce{-NH-CH2-COOH} (Gly). Kí hiệu: Gly-Ala-Gly.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong phân tử tetrapeptide Gly-Val-Glu-Ala, amino acid đầu C là
\choice
{Gly}
{Val}
{Glu}
""" + (r"{\True Ala}" if is_teacher else r"{Ala}") + r"""
\loigiai{""" + (r"""
Quy ước viết mạch peptide: amino acid đầu N ở bên trái ngoài cùng (Gly), amino acid đầu C ở bên phải ngoài cùng (Ala).\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Insulin, một loại hormone thuộc loại protein, được sản sinh bởi tuyến tuỵ, có chức năng điều hoà quá trình chuyển hoá glucose trong cơ thể. Insulin thúc đẩy sự hấp thu glucose của các tế bào và dự trữ glucose dư thừa trong gan và cơ. Insulin là
\choice
""" + (r"{\True protein đơn giản}" if is_teacher else r"{protein đơn giản}") + r"""
{protein phức tạp}
{dipeptide}
{monomer}
\loigiai{""" + (r"""
Insulin là một protein đơn giản (gồm 51 amino acid tạo bởi 2 chuỗi polypeptide liên kết qua cầu disulfide), chỉ cấu tạo hoàn toàn từ các gốc $\alpha$-amino acid.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- KNTT). Trong cơ thể, enzyme có chức năng nào sau đây?
\choice
{Cấu trúc tế bào}
{Chất điện giải}
{Chất dự trữ năng lượng}
""" + (r"{\True Xúc tác sinh học}" if is_teacher else r"{Xúc tác sinh học}") + r"""
\loigiai{""" + (r"""
Enzyme là các chất xúc tác sinh học có bản chất protein, thúc đẩy tốc độ phản ứng sinh hoá trong điều kiện nhiệt độ cơ thể với độ chọn lọc cực cao.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\begin{ex}
(SBT -- CTST). Amine nào sau đây phản ứng được với nitrous acid tạo thành muối diazonium bền ở nhiệt độ thấp?
\choice
{Methanamine}
{Methanediamine}
""" + (r"{\True Benzenamine}" if is_teacher else r"{Benzenamine}") + r"""
{Phenylmethanamine}
\loigiai{""" + (r"""
Benzenamine (aniline, \ce{C6H5NH2}) là amine bậc 1 thơm phản ứng với \ce{HNO2} ở nhiệt độ thấp ($0 - 5^\circ\text{C}$) tạo muối diazonium bền: $\ce{C6H5NH2 + HNO2 + HCl -> [C6H5N2]+Cl- + 2H2O}$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- CTST). Dạng ion chủ yếu nào của amino acid có trong môi trường acid mạnh (pH thấp)?
{\centering
	\includegraphics[width=10.0cm]{images_crop_de07/cau12_options.png}\par}
\loigiai{""" + (r"""
Trong môi trường acid mạnh (nồng độ \ce{H+} cao), nhóm carboxyl -\ce{COO-} nhận \ce{H+} trở thành -\ce{COOH}, nhóm -\ce{NH2} nhận \ce{H+} thành -\ce{NH3+}. Amino acid tồn tại chủ yếu ở dạng cation: \ce{H3N+-CH(CH3)-COOH}.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- CĐ). Một tripeptide X được cấu thành từ 2 phân tử Ala và 1 phân tử Gly. Công thức cấu tạo của X \textbf{không} thể là
\choice
{Ala-Ala-Gly}
{Ala-Gly-Ala}
{Gly-Ala-Ala}
""" + (r"{\True Gly-Ala-Gly}" if is_teacher else r"{Gly-Ala-Gly}") + r"""
\loigiai{""" + (r"""
Tripeptide gồm 2 Ala và 1 Gly chỉ có 3 đồng phân cấu tạo là Ala-Ala-Gly, Ala-Gly-Ala, Gly-Ala-Ala. Công thức Gly-Ala-Gly gồm 2 Gly và 1 Ala nên không thỏa mãn.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- CĐ). Cho peptide X có công thức cấu tạo như hình bên dưới:
{\centering
	\includegraphics[width=4.5cm]{images_crop_de07/cau14_dipeptide.png}\par}
Khi thuỷ phân hoàn toàn X trong môi trường NaOH thu được sản phẩm hữu cơ có công thức là
\choice
{\ce{H2NCH(CH3)COOH} và \ce{H2NCH2COOH}}
""" + (r"{\True \ce{H2NCH(CH3)COONa} và \ce{H2NCH2COONa}}" if is_teacher else r"{\ce{H2NCH(CH3)COONa} và \ce{H2NCH2COONa}}") + r"""
{\ce{H2NCH(CH3)COONa} và \ce{H2NCH2COOH}}
{\ce{H2NCH(CH3)COOH} và \ce{H2NCH2COONa}}
\loigiai{""" + (r"""
Peptide X là Ala-Gly. Trong môi trường NaOH, liên kết peptide bị thủy phân và các nhóm acid tạo thành muối sodium: \ce{H2NCH(CH3)COONa} và \ce{H2NCH2COONa}.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- CĐ). Leucine có công thức cấu tạo \ce{HOOCCH(NH2)CH2CH(CH3)2}, là amino acid có khả năng điều hòa sự tổng hợp protein của cơ. Tên theo danh pháp thay thế của leucine là
\choice
{2-aminoisohexanoic acid}
""" + (r"{\True 2-amino-4-methylpentanoic acid}" if is_teacher else r"{2-amino-4-methylpentanoic acid}") + r"""
{4-amino-2-methylpentanoic acid}
{2-amino-isohexanoic acid}
\loigiai{""" + (r"""
Mạch chính có 5 carbon, đánh số từ nhóm -\ce{COOH} (C1): C2 gắn nhóm amino, C4 gắn nhóm methyl => Tên thay thế: 2-amino-4-methylpentanoic acid.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- CTST). Dung dịch của chất nào sau đây có môi trường base?
{\centering
	\includegraphics[width=10.5cm]{images_crop_de07/cau16_options.png}\par}
\loigiai{""" + (r"""
Chất A là lysine ở dạng ion lưỡng cực: $\ce{{}^+H3N-[CH2]4-CH(NH2)-COO-}$, có 2 nhóm amino (tính base) và 1 nhóm carboxyl (tính acid) nên dung dịch có tính base ($pH > 7$).\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- CTST). Nguyên nhân chủ yếu dẫn đến tính tan của amino acid trong nước là do
\choice
""" + (r"{\True phân tử phân cực mạnh nên dễ tan trong nước}" if is_teacher else r"{phân tử phân cực mạnh nên dễ tan trong nước}") + r"""
{cấu tạo lưỡng cực của phân tử, có tính kị nước nên ít tan trong nước}
{năng lượng liên kết của phân tử lớn, khó phá vỡ nên ít tan trong nước}
{hình thành liên kết hydrogen với nước nên dễ tan trong nước}
\loigiai{""" + (r"""
Amino acid tồn tại ở dạng ion lưỡng cực phân cực rất mạnh, lực tương tác tĩnh điện giữa các ion này với các phân tử nước phân cực rất lớn (solvat hoá), làm cho amino acid dễ tan trong nước.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các phát biểu sau:\\
(a) Phenylamine có lực base yếu hơn ammonia.\\
(b) Phenylamine tan ít trong nước nhưng tan tốt trong dung dịch HCl.\\
(c) Khi nấu canh cua, hiện tượng riêu cua nổi lên trên là do sự đông tụ protein.\\
(d) Trong tơ tằm có các gốc $\alpha$-amino acid.\\
(e) Dùng giấm ăn hoặc chanh khử được mùi tanh trong cá do amine gây ra.\\
Số phát biểu đúng là
\choice
""" + (r"{\True 5}" if is_teacher else r"{5}") + r"""
{2}
{3}
{4}
\loigiai{""" + (r"""
- (a) Đúng, vòng benzene hút e làm giảm mật độ e trên nguyên tử N.\\
- (b) Đúng, phản ứng tạo muối tan \ce{C6H5NH3Cl}.\\
- (c) Đúng, nhiệt độ làm đông tụ protein trong thịt cua nổi lên.\\
- (d) Đúng, tơ tằm là sợi protein thiên nhiên tạo bởi các $\alpha$-amino acid.\\
- (e) Đúng, acid trong chanh/giấm trung hòa amine tạo muối không bay hơi.\\
=> Cả 5 phát biểu đều đúng.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\Closesolutionfile{ans}

""" + (r"\newpage" if not is_teacher else "") + r"""

\vspace{0.15cm}
\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\setcounter{ex}{0}

\begin{ex}
(SBT -- KNTT). Các phát biểu về điều chế và ứng dụng của amine:
\choiceTF
""" + (r"{\True Một số amine có thể được điều chế bằng cách alkyl hoá ammonia}" if is_teacher else r"{Một số amine có thể được điều chế bằng cách alkyl hoá ammonia}") + r"""
""" + (r"{\True Một số amine có thể được điều chế bằng cách khử hợp chất nitro}" if is_teacher else r"{Một số amine có thể được điều chế bằng cách khử hợp chất nitro}") + r"""
""" + (r"{\True Amine được sử dụng để tổng hợp một số loại dược phẩm}" if is_teacher else r"{Amine được sử dụng để tổng hợp một số loại dược phẩm}") + r"""
""" + (r"{\True Amine được sử dụng để tổng hợp một số loại polymer}" if is_teacher else r"{Amine được sử dụng để tổng hợp một số loại polymer}") + r"""
\loigiai{""" + (r"""
a) Đúng, alkyl hóa ammonia bằng dẫn xuất halogen: $\ce{RX + NH3 -> RNH2 + HX}$.\\
b) Đúng, khử hợp chất nitro (ví dụ khử nitrobenzene bằng Fe/HCl hoặc \ce{H2/Ni}) tạo aniline.\\
c) Đúng, amine dùng sản xuất nhiều loại dược phẩm quan trọng.\\
d) Đúng, ví dụ hexamethylenediamine dùng tổng hợp tơ nylon-6,6.\\
\textbf{Đáp án: a - Đ, b - Đ, c - Đ, d - Đ.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- CTST). Cho một số hợp chất chứa nguyên tố nitrogen như hình bên dưới:\\
Về tính chất hoá học của các chất, hãy cho biết các phát biểu sau đúng hay sai:
{\centering
	\includegraphics[width=9.5cm]{images_crop_de07/p2_cau2_compounds.png}\par}
\choiceTF
""" + (r"{\True Đều có phản ứng với dung dịch HCl, sinh ra muối}" if is_teacher else r"{Đều có phản ứng với dung dịch HCl, sinh ra muối}") + r"""
""" + (r"{\True Có 1 chất tham gia phản ứng thuỷ phân}" if is_teacher else r"{Có 1 chất tham gia phản ứng thuỷ phân}") + r"""
""" + (r"{\False Có thể trùng ngưng tạo thành peptide}" if is_teacher else r"{Có thể trùng ngưng tạo thành peptide}") + r"""
""" + (r"{\False Chỉ tạo phức với \ce{Cu(OH)2} tạo thành dung dịch có màu xanh}" if is_teacher else r"{Chỉ tạo phức với \ce{Cu(OH)2} tạo thành dung dịch có màu xanh}") + r"""
\loigiai{""" + (r"""
a) Đúng, cả 3 chất đều có nhóm amino -\ce{NH2} có tính base nên tác dụng với HCl tạo muối.\\
b) Đúng, chỉ có dipeptide Gly-Gly có liên kết peptide -\ce{CO-NH}- bị thuỷ phân.\\
c) Sai, ethylamine không có nhóm -\ce{COOH} nên không trùng ngưng; $\beta$-alanine không tạo peptide thông thường.\\
d) Sai, ethylamine và amino acid tạo phức chất xanh lam đặc trưng với \ce{Cu^2+}, dipeptide tạo phức chelate khác.\\
\textbf{Đáp án: a - Đ, b - Đ, c - S, d - S.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Collagen được coi là một loại protein quan trọng đối với làn da, trong đó peptide lại có vai trò lớn thúc đẩy sản xuất collagen. Điều này giải thích cho việc peptide có khả năng chống lại sự lão hóa, giúp làm mờ các vết nhăn làm làn da luôn căng mịn.
\choiceTF
""" + (r"{\False Trong một phân tử tetrapeptide mạch hở có 4 liên kết peptide}" if is_teacher else r"{Trong một phân tử tetrapeptide mạch hở có 4 liên kết peptide}") + r"""
""" + (r"{\True Peptide là loại hợp chất chứa gốc $\alpha$-amino acid}" if is_teacher else r"{Peptide là loại hợp chất chứa gốc $\alpha$-amino acid}") + r"""
""" + (r"{\False Trong môi trường base, dipeptide mạch hở tác dụng được với \ce{Cu(OH)2} cho hợp chất màu tím}" if is_teacher else r"{Trong môi trường base, dipeptide mạch hở tác dụng được với \ce{Cu(OH)2} cho hợp chất màu tím}") + r"""
""" + (r"{\False Các hợp chất peptide kém bền trong môi trường base nhưng bền trong môi trường acid}" if is_teacher else r"{Các hợp chất peptide kém bền trong môi trường base nhưng bền trong môi trường acid}") + r"""
\loigiai{""" + (r"""
a) Sai, phân tử tetrapeptide có 4 gốc amino acid nên chỉ có $4 - 1 = 3$ liên kết peptide.\\
b) Đúng, peptide cấu tạo từ 2 đến 50 gốc $\alpha$-amino acid liên kết qua liên kết peptide.\\
c) Sai, phản ứng màu biuret cần từ 2 liên kết peptide trở lên (từ tripeptide trở lên). Dipeptide không cho màu tím.\\
d) Sai, peptide bị thuỷ phân trong cả môi trường acid và base nên kém bền trong cả hai môi trường.\\
\textbf{Đáp án: a - S, b - Đ, c - S, d - S.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Thành phần tạo nên vị ngọt đặc trưng của nước mắm (được sản xuất từ cá) và nước tương (được sản xuất từ đậu nành) là các amino acid tạo thành từ sự thuỷ phân hoàn toàn của protein có trong cá hoặc đậu nành.
\choiceTF
""" + (r"{\False Độ đạm tương ứng với hàm lượng oxygen có trong nước mắm, nước tương}" if is_teacher else r"{Độ đạm tương ứng với hàm lượng oxygen có trong nước mắm, nước tương}") + r"""
""" + (r"{\True Độ đạm có tỉ lệ thuận với hàm lượng amino acid có trong nước mắm, nước tương}" if is_teacher else r"{Độ đạm có tỉ lệ thuận với hàm lượng amino acid có trong nước mắm, nước tương}") + r"""
""" + (r"{\False Vì hàm lượng nitrogen quyết định đến độ đạm trong khi nitrogen có trong thành phần của amino acid nên độ đạm có tỉ lệ nghịch với hàm lượng amino acid}" if is_teacher else r"{Vì hàm lượng nitrogen quyết định đến độ đạm trong khi nitrogen có trong thành phần của amino acid nên độ đạm có tỉ lệ nghịch với hàm lượng amino acid}") + r"""
""" + (r"{\True Độ đạm trong nước mắm là quá trình thủy phân cá, các chất protein được thủy phân tạo thành các amino acid nước mắm sẽ có mùi thơm đặc trưng riêng}" if is_teacher else r"{Độ đạm trong nước mắm là quá trình thủy phân cá, các chất protein được thủy phân tạo thành các amino acid nước mắm sẽ có mùi thơm đặc trưng riêng}") + r"""
\loigiai{""" + (r"""
a) Sai, độ đạm là hàm lượng nitrogen tổng số tính theo gam trong 1 lít nước mắm.\\
b) Đúng, amino acid càng nhiều thì lượng nitrogen càng lớn, độ đạm càng cao.\\
c) Sai, mối quan hệ là tỉ lệ thuận chứ không phải tỉ lệ nghịch.\\
d) Đúng, quá trình ủ chượp cá thủy phân protein thành amino acid tạo hương vị đặc trưng.\\
\textbf{Đáp án: a - S, b - Đ, c - S, d - Đ.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\vspace{0.15cm}
\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6. Mỗi câu hỏi thí sinh điền câu trả lời vào ô tương ứng.}

\setcounter{ex}{0}

\begin{ex}
Cho dãy các chất: \ce{H2NCH2COOH}, \ce{HCOONH4}, \ce{(CH3NH3)2CO3}, \ce{C6H5NH2}, \ce{C2H5NH2}, \ce{CH3COOH}, \ce{H2NCH2CONHCH(CH3)COOH}. Số lượng chất trong dãy phản ứng được với cả 2 dung dịch NaOH và dung dịch HCl là
\loigiai{""" + (r"""
Các chất vừa tác dụng với NaOH vừa tác dụng với HCl gồm:\\
1. \ce{H2NCH2COOH} (glycine - amino acid lưỡng tính).\\
2. \ce{HCOONH4} (muối ammonium của acid yếu).\\
3. \ce{(CH3NH3)2CO3} (muối alkylammonium của acid yếu carbonic).\\
4. \ce{H2NCH2CONHCH(CH3)COOH} (dipeptide Gly-Ala, lưỡng tính).\\
- \ce{C6H5NH2} và \ce{C2H5NH2} chỉ tác dụng HCl.\\
- \ce{CH3COOH} chỉ tác dụng NaOH.\\
=> Có 4 chất thỏa mãn.\\
\textbf{Đáp số: 4.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Đáp số:} \dotfill\vspace{0.25cm}""" if not is_teacher else "") + r"""

\begin{ex}
Cho các chất sau: \ce{CH3COONH4}, \ce{CH3COONH3CH3}, \ce{C2H5NH2}, \ce{H2NCH2COOC2H5}. Số chất trong dãy vừa tác dụng được với NaOH và vừa tác dụng được với HCl trong dung dịch là:
\loigiai{""" + (r"""
Các chất vừa tác dụng với NaOH vừa tác dụng với HCl gồm:\\
1. \ce{CH3COONH4} (muối ammonium).\\
2. \ce{CH3COONH3CH3} (muối methylammonium).\\
3. \ce{H2NCH2COOC2H5} (ester của glycine: nhóm -\ce{NH2} tác dụng HCl, nhóm ester thuỷ phân trong NaOH).\\
- \ce{C2H5NH2} chỉ tác dụng HCl.\\
=> Có 3 chất thỏa mãn.\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Đáp số:} \dotfill\vspace{0.25cm}""" if not is_teacher else "") + r"""

\begin{ex}
Nếu thuỷ phân không hoàn toàn pentapeptide Gly-Ala-Gly-Ala-Gly thì thu được tối đa bao nhiêu dipeptide khác nhau?
\loigiai{""" + (r"""
Mạch pentapeptide: Gly-Ala-Gly-Ala-Gly.\\
Các đoạn 2 mắt xích liên tiếp: Gly-Ala, Ala-Gly, Gly-Ala, Ala-Gly.\\
=> Thu được tối đa 2 dipeptide khác nhau là Gly-Ala và Ala-Gly.\\
\textbf{Đáp số: 2.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Đáp số:} \dotfill\vspace{0.25cm}""" if not is_teacher else "") + r"""

\begin{ex}
Tiến hành thí nghiệm phản ứng màu biuret của lòng trắng trứng (protein) theo các bước sau đây:\\
\textit{Bước 1:} Cho vào ống nghiệm 1 giọt dung dịch \ce{CuSO4} 2\% + 1 mL dung dịch NaOH 30\%.\\
\textit{Bước 2:} Lắc nhẹ, gạn lớp dung dịch để giữ kết tủa.\\
\textit{Bước 3:} Thêm 4 mL dung dịch lòng trắng trứng vào ống nghiệm, lắc đều.\\
Cho các nhận định sau:\\
(1) Cần lấy dư dung dịch NaOH để đảm bảo môi trường cho phản ứng tạo phức.\\
(2) Có thể thay thế dung dịch lòng trắng trứng bằng dung dịch Gly-Ala.\\
(3) Sau bước 3, kết tủa bị hoà tan và dung dịch có màu tím đặc trưng.\\
(4) Sau bước 1, trong ống nghiệm xuất hiện kết tủa màu xanh lam.\\
Có bao nhiêu nhận định đúng trong các nhận định trên?
\loigiai{""" + (r"""
- (1) Đúng, phản ứng tạo phức biuret cần môi trường kiềm.\\
- (2) Sai, Gly-Ala là dipeptide chỉ có 1 liên kết peptide, không cho phản ứng biuret.\\
- (3) Đúng, protein hoà tan \ce{Cu(OH)2} tạo dung dịch phức màu tím.\\
- (4) Đúng, phản ứng tạo kết tủa xanh lam \ce{Cu(OH)2}.\\
=> Có 3 nhận định đúng là (1), (3), (4).\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Đáp số:} \dotfill\vspace{0.25cm}""" if not is_teacher else "") + r"""

\begin{ex}
Xác định phân tử khối gần đúng của một hemoglobin (huyết cầu tố) chứa 0,4\% Fe về khối lượng (mỗi phân tử hemoglobin chỉ chứa 1 nguyên tử iron).
\loigiai{""" + (r"""
Mỗi phân tử hemoglobin chứa 1 nguyên tử Fe ($M_{\text{Fe}} = 56\text{ amu}$).\\
Phân tử khối gần đúng của hemoglobin là:
$$M = \frac{56 \times 100}{0,4} = 14\,000\text{ amu}.$$
\textbf{Đáp số: 14000.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Đáp số:} \dotfill\vspace{0.25cm}""" if not is_teacher else "") + r"""

\begin{ex}
Một khẩu phần thịt gà nướng 100g chứa khoảng 25g protein. Một người trong ngày đã ăn 150g thịt gà, 200g cá hồi (chứa 22g protein/100g) và 50g phô mai (chứa 7g protein/100g). Hãy tính tổng lượng protein (theo đơn vị gram) người đó đã hấp thụ từ các thực phẩm này.
\loigiai{""" + (r"""
Lượng protein từ từng loại thực phẩm:\\
- Thịt gà: $150 \times \frac{25}{100} = 37,5\text{ g}$.\\
- Cá hồi: $200 \times \frac{22}{100} = 44\text{ g}$.\\
- Phô mai: $50 \times \frac{7}{100} = 3,5\text{ g}$.\\
Tổng lượng protein hấp thụ là: $37,5 + 44 + 3,5 = 85\text{ g}$.\\
\textbf{Đáp số: 85.}""" if is_teacher else "") + r"""}
\end{ex}
""" + (r"""\noindent\textit{Đáp số:} \dotfill\vspace{0.25cm}""" if not is_teacher else "") + r"""

\vspace{0.1cm}
\begin{center}
	\textbf{---------- HẾT ----------}
\end{center}
"""

    if is_teacher:
        tex += r"""
\vspace{0.3cm}
\noindent
\textbf{\large\color{mainGreen}BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ SỐ 07}

\vspace{0.2cm}
\noindent\textbf{PHẦN I:}\\
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} & \textbf{17} & \textbf{18} \\
\hline
A & D & B & D & C & B & B & D & A & D & C & D & D & B & B & A & A & A \\
\hline
\end{tabular}

\vspace{0.3cm}
\noindent\textbf{PHẦN II:}\\
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Lệnh hỏi a} & \textbf{Lệnh hỏi b} & \textbf{Lệnh hỏi c} & \textbf{Lệnh hỏi d} \\
\hline
\textbf{1} & Đ & Đ & Đ & Đ \\
\hline
\textbf{2} & Đ & Đ & S & S \\
\hline
\textbf{3} & S & Đ & S & S \\
\hline
\textbf{4} & S & Đ & S & Đ \\
\hline
\end{tabular}

\vspace{0.3cm}
\noindent\textbf{PHẦN III:}\\
\begin{tabular}{|c|c|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} \\
\hline
\textbf{Đáp số} & 4 & 3 & 2 & 3 & 14000 & 85 \\
\hline
\end{tabular}
"""

    tex += r"""
\end{document}
"""
    return tex

def compile_pdf(tex_path):
    cmd = ["pdflatex", "-interaction=nonstopmode", os.path.basename(tex_path)]
    work_dir = os.path.dirname(tex_path)
    res = subprocess.run(cmd, cwd=work_dir, capture_output=True, text=True, errors="ignore")
    # run second time for references / page numbers
    subprocess.run(cmd, cwd=work_dir, capture_output=True, text=True, errors="ignore")
    
    # cleanup temp files
    base = os.path.splitext(os.path.basename(tex_path))[0]
    for ext in [".aux", ".log", ".thm", ".out"]:
        f = os.path.join(work_dir, base + ext)
        if os.path.exists(f):
            try: os.remove(f)
            except: pass
    for ans_f in glob.glob(os.path.join(work_dir, "ans-*.tex")):
        try: os.remove(ans_f)
        except: pass
        
    pdf_name = base + ".pdf"
    pdf_path = os.path.join(work_dir, pdf_name)
    if os.path.exists(pdf_path):
        print(f"Compiled successfully: {pdf_name}")
    else:
        print(f"Compilation failed for: {tex_path}")
        print("STDOUT:", res.stdout[-800:])
        print("STDERR:", res.stderr[-800:])

if __name__ == "__main__":
    # 1. Hoc Sinh
    hs_tex_path = os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_12_De07_HocSinh.tex")
    with open(hs_tex_path, "w", encoding="utf-8") as f:
        f.write(create_latex_de07(is_teacher=False))
    print(f"Wrote {hs_tex_path}")
    compile_pdf(hs_tex_path)
    
    # 2. Giao Vien
    gv_tex_path = os.path.join(OUT_DIR, "De_Kiem_Tra_Hoa_12_De07_GiaoVien.tex")
    with open(gv_tex_path, "w", encoding="utf-8") as f:
        f.write(create_latex_de07(is_teacher=True))
    print(f"Wrote {gv_tex_path}")
    compile_pdf(gv_tex_path)
