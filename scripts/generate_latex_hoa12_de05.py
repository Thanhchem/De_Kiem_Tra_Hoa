# -*- coding: utf-8 -*-
"""
Script biên dịch LaTeX và PDF chuẩn "FORM ĐỀ KIỂM TRA" cho:
ĐỀ KIỂM TRA SỐ 5 - CHƯƠNG 2: CARBOHYDRATE (HÓA HỌC 12)
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De05_HocSinh.pdf (Chính xác 4 trang)
2. De_Kiem_Tra_Hoa_12_De05_GiaoVien.pdf (Có lời giải chi tiết và bảng đáp án)
"""

import os
import subprocess
import shutil

TEX_DIR = r"c:\Antigravity_Thanh\San_Pham"

# Ensure ex_test.sty is present in San_Pham
if not os.path.exists(os.path.join(TEX_DIR, "ex_test.sty")):
    shutil.copyfile(r"c:\Antigravity_Thanh\ex_test.sty", os.path.join(TEX_DIR, "ex_test.sty"))

def make_tex(is_teacher=False):
    ver_tag = "gv" if is_teacher else "hs"
    ans_file = f"ans-de05{ver_tag}"
    
    tex = r"""\documentclass[12pt,a4paper]{article}
\usepackage[utf8]{vietnam}
\usepackage{amsmath,amssymb}
\usepackage{geometry}
\geometry{top=1.2cm,bottom=1.0cm,left=1.3cm,right=1.3cm,headheight=26pt,headsep=10pt,footskip=14pt}
\usepackage{tikz}
\usepackage{xcolor}
\usepackage{enumerate}
\usepackage{graphicx}
\usepackage[version=4]{mhchem}
\usepackage{fancyhdr}
\usepackage{ex_test}

\renewcommand{\baselinestretch}{0.92}
\setlength{\parskip}{0pt}

\definecolor{hfRed}{RGB}{211,47,47}
\definecolor{hfGreen}{RGB}{139,195,144}

\providecommand{\False}{}

\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small\itshape\color{hfRed} Tài liệu lưu hành nội bộ}
\fancyhead[R]{\small\itshape\color{hfRed} Thầy \textbf{TRẦN VĂN THẠNH} - 0777.470.803}
\renewcommand{\headrulewidth}{0.8pt}
\renewcommand{\headrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \headrulewidth\hfill}}

\fancyfoot[L]{\small\itshape\color{hfRed} CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế \quad|\quad CS2: 24 Đặng Thái Thân, TP Huế}
\fancyfoot[R]{\small\bfseries\color{hfRed} Trang \thepage}
\renewcommand{\footrulewidth}{0.8pt}
\renewcommand{\footrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \footrulewidth\hfill}}

\fancypagestyle{firstpage}{
  \fancyhf{}
  \renewcommand{\headrulewidth}{0pt}
  \fancyfoot[L]{\small\itshape\color{hfRed} CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế \quad|\quad CS2: 24 Đặng Thái Thân, TP Huế}
  \fancyfoot[R]{\small\bfseries\color{hfRed} Trang \thepage}
  \renewcommand{\footrulewidth}{0.8pt}
  \renewcommand{\footrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \footrulewidth\hfill}}
}

\begin{document}
\thispagestyle{firstpage}

\noindent
\begin{minipage}[t]{0.44\textwidth}
	\centering\fontsize{10pt}{12pt}\selectfont
	\textbf{SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ}\\
	\textbf{ĐỀ CHÍNH THỨC}\\
	\textit{(Đề thi có 04 trang)}
\end{minipage}\hfill
\begin{minipage}[t]{0.54\textwidth}
	\centering\small
	\textbf{ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 2 (CARBOHYDRATE)}\\
	\textbf{MÔN: HOÁ HỌC -- LỚP 12}\\
	\textit{Thời gian làm bài: 50 phút (không kể thời gian phát đề)}\\
	\fbox{\textbf{MÃ ĐỀ: 105}}
\end{minipage}
"""
    if is_teacher:
        tex += r"""
\smallskip
\begin{center}
	{\color{hfRed}\textbf{\large (BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}}
\end{center}
"""

    tex += r"""
\smallskip
\noindent\hrulefill
\smallskip

\noindent\textbf{Họ, tên thí sinh:} \dotfill\ \textbf{Lớp:} \dotfill\ \textbf{Số báo danh:} \dotfill

\smallskip
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; K = 39; Cu = 64; Ag = 108).}

\vspace{0.2cm}
\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\Opensolutionfile{ans}[""" + ans_file + r"""]

\begin{ex}
Saccharose và fructose đều thuộc loại
\choice
{monosaccharide.}
{disaccharide.}
{polysaccharide.}
""" + (r"{\True carbohydrate.}" if is_teacher else r"{carbohydrate.}") + r"""
\loigiai{""" + (r"""
Saccharose là disaccharide, fructose là monosaccharide; cả hai đều thuộc nhóm carbohydrate.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Một chất khi thủy phân trong môi trường acid, đun nóng \textbf{không} tạo ra glucose. Chất đó là
\choice
{tinh bột.}
{saccharose.}
{cellulose.}
""" + (r"{\True tripalmitin.}" if is_teacher else r"{tripalmitin.}") + r"""
\loigiai{""" + (r"""
Tripalmitin là chất béo (triglyceride), khi thủy phân trong môi trường acid tạo ra palmitic acid và glycerol, không sinh ra glucose.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- CTST). Carbohydrate nào sau đây là thành phần chính của mật ong?
\choice
{Glucose.}
{Maltose.}
{Saccharose.}
""" + (r"{\True Fructose.}" if is_teacher else r"{Fructose.}") + r"""
\loigiai{""" + (r"""
Fructose chiếm khoảng $40\%$ khối lượng mật ong, là thành phần đường chiếm tỉ lệ cao nhất (glucose chiếm khoảng $30\%$).\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- KNTT). Saccharose thường được tìm thấy trong loại thực vật nào sau đây?
\choice
{Cây đậu nành.}
{Cây lúa mì.}
""" + (r"{\True Cây mía.}" if is_teacher else r"{Cây mía.}") + r"""
{Cây cà phê.}
\loigiai{""" + (r"""
Saccharose có nhiều nhất trong cây mía (đường mía), hoa thốt nốt và củ cải đường.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- KNTT). Công thức phân tử của saccharose là
\choice
{$\ce{C5H10O5}$.}
{$\ce{C6H12O6}$.}
""" + (r"{\True $\ce{C12H22O11}$.}" if is_teacher else r"{$\ce{C12H22O11}$.}") + r"""
{$(\ce{C6H10O5})_n$.}
\loigiai{""" + (r"""
Saccharose là disaccharide có CTPT là $\ce{C12H22O11}$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- KNTT). Tinh bột và cellulose đều là
\choice
{disaccharide.}
{monosaccharide.}
""" + (r"{\True polysaccharide.}" if is_teacher else r"{polysaccharide.}") + r"""
{oligosaccharide.}
\loigiai{""" + (r"""
Tinh bột và cellulose đều là polymer tự nhiên thuộc nhóm polysaccharide có công thức chung $(\ce{C6H10O5})_n$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- KNTT). Phản ứng màu với dung dịch iodine là tính chất của chất nào sau đây?
\choice
{Glucose.}
{Fructose.}
{Saccharose.}
""" + (r"{\True Tinh bột.}" if is_teacher else r"{Tinh bột.}") + r"""
\loigiai{""" + (r"""
Amylose trong tinh bột hấp phụ các phân tử iodine tạo hợp chất bọc có màu xanh tím đặc trưng.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Glucose có thể lên men tạo lactic acid có trong sữa chua qua phản ứng nào sau đây?
\choice
{$2\ce{C6H12O6 + Cu(OH)2 -> (C6H11O6)2Cu + 2H2O}$}
{$\ce{C6H12O6} \xrightarrow{\text{enzyme}} 2\ce{C2H5OH + 2CO2}$}
""" + (r"{\True $\ce{C6H12O6} \xrightarrow{\text{enzyme}} 2\ce{CH3CH(OH)COOH}$}" if is_teacher else r"{$\ce{C6H12O6} \xrightarrow{\text{enzyme}} 2\ce{CH3CH(OH)COOH}$}") + r"""
{$\ce{CH2OH[CHOH]4CHO + Br2 + H2O} \xrightarrow{t^\circ} \ce{CH2OH[CHOH]4COOH + 2HBr}$}
\loigiai{""" + (r"""
Lên men lactic chuyển hóa glucose thành lactic acid ($\ce{CH3CH(OH)COOH}$).\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SGK 12 -- CTST). Cho các carbohydrate sau: glucose, fructose, saccharose và maltose. Số carbohydrate có khả năng mở vòng trong dung dịch nước là
\choice
{1.}
{2.}
""" + (r"{\True 3.}" if is_teacher else r"{3.}") + r"""
{4.}
\loigiai{""" + (r"""
Glucose, fructose, maltose đều có nhóm hemiacetal/hemiketal tự do nên có thể mở vòng trong dung dịch nước; saccharose bị khóa liên kết glycosidic ở cả 2 gốc nên không thể mở vòng.\\
$\Rightarrow$ Có 3 chất thỏa mãn (glucose, fructose, maltose).\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Công thức cấu tạo dạng mạch vòng $\beta$-fructose là:
\begin{center}
\begin{tabular}{cccc}
\textbf{A} & \textbf{B} & \textbf{C} & \textbf{D} \\
\includegraphics[width=3.2cm]{images_crop_de05/cau10_opt_A.png} &
\includegraphics[width=3.2cm]{images_crop_de05/cau10_opt_B.png} &
\includegraphics[width=3.2cm]{images_crop_de05/cau10_opt_C.png} &
\includegraphics[width=3.2cm]{images_crop_de05/cau10_opt_D.png}
\end{tabular}
\end{center}
\choice
{Hình A.}
{Hình B.}
{Hình C.}
""" + (r"{\True Hình D.}" if is_teacher else r"{Hình D.}") + r"""
\loigiai{""" + (r"""
Fructose tạo vòng furanose 5 cạnh; ở dạng $\beta$-fructose, nhóm hemiketal $-\ce{OH}$ ở C2 hướng lên trên cùng phía với nhóm $-\ce{CH2OH}$ ở C5 (như hình D).\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Mô tả nào dưới đây \textbf{không} đúng với glucose?
\choice
""" + (r"{\True Chất rắn, khó tan trong nước và có vị ngọt.}" if is_teacher else r"{Chất rắn, khó tan trong nước và có vị ngọt.}") + r"""
{Có mặt trong hầu hết các bộ phận của cây, nhất là trong quả chín.}
{Còn có tên gọi là đường nho.}
{Có 80-130 mg/dL trong máu người trưởng thành, khỏe mạnh và lúc đói.}
\loigiai{""" + (r"""
Glucose là chất rắn kết tinh, không màu, \textbf{dễ tan trong nước} (không phải khó tan).\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các tính chất sau: (1) có vị ngọt, (2) dễ tan trong nước, (3) có phản ứng tráng bạc, (4) bị thủy phân trong môi trường acid, (5) hòa tan $\ce{Cu(OH)2}$ thành dung dịch màu xanh lam.\\
Số tính chất đúng với saccharose là
\choice
""" + (r"{\True 4.}" if is_teacher else r"{4.}") + r"""
{3.}
{5.}
{2.}
\loigiai{""" + (r"""
Saccharose có 4 tính chất đúng: (1), (2), (4), (5). Tính chất (3) sai vì phân tử saccharose không có nhóm $-\ce{CHO}$ hay hemiacetal tự do để tráng bạc.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Nhận định sai là
\choice
{Phân biệt glucose và saccharose bằng phản ứng tráng gương.}
{Phân biệt tinh bột và cellulose bằng $\ce{I2}$.}
""" + (r"{\True Phân biệt saccharose và glycerol bằng $\ce{Cu(OH)2}$.}" if is_teacher else r"{Phân biệt saccharose và glycerol bằng $\ce{Cu(OH)2}$.}") + r"""
{Phân biệt tinh bột và saccharose bằng $\ce{I2}$.}
\loigiai{""" + (r"""
Cả saccharose và glycerol đều hòa tan $\ce{Cu(OH)2}$ ở điều kiện thường tạo dung dịch màu xanh lam, do đó không thể dùng $\ce{Cu(OH)2}$ ở điều kiện thường để phân biệt hai chất này.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- CTST). Hầu hết những phản ứng sinh hoá xảy ra nhờ sự xúc tác của các enzyme. Biểu đồ dưới đây thể hiện tốc độ của một phản ứng sinh hoá do enzyme X làm chất tác theo nhiệt độ:
\begin{center}
	\includegraphics[width=6.5cm]{images_crop_de05/cau14_do_thi_toc_do.png}
\end{center}
Nhiệt độ tối ưu cho enzyme X hoạt động là
\choice
{0$^\circ$C.}
""" + (r"{\True 50$^\circ$C.}" if is_teacher else r"{50$^\circ$C.}") + r"""
{60$^\circ$C.}
{70$^\circ$C.}
\loigiai{""" + (r"""
Đỉnh cao nhất của đồ thị tốc độ phản ứng ứng với nhiệt độ 50$^\circ$C, đây là nhiệt độ tối ưu cho enzyme X hoạt động.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
(SBT -- CTST). Carbohydrate nào có cấu trúc phân tử được biểu diễn dưới đây?
\begin{center}
	\includegraphics[width=8.5cm]{images_crop_de05/cau15_cau_truc_amylose.png}
\end{center}
\choice
{Saccharose.}
{Cellulose.}
{Maltose.}
""" + (r"{\True Amylose.}" if is_teacher else r"{Amylose.}") + r"""
\loigiai{""" + (r"""
Cấu trúc gồm các mắt xích $\alpha$-glucose liên kết qua liên kết $\alpha$-1,4-glycosidic tạo mạch không phân nhánh dài khoảng 1000 mắt xích, đây là cấu tạo của amylose (thành phần của tinh bột).\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các phát biểu sau về carbohydrate:\\
(a) Tất cả các carbohydrate đều có phản ứng thủy phân.\\
(b) Thủy phân hoàn toàn tinh bột thu được glucose.\\
(c) Glucose, fructose đều có phản ứng tráng bạc.\\
(d) Glucose làm mất màu nước bromine.\\
Số phát biểu đúng là:
\choice
""" + (r"{\True 3.}" if is_teacher else r"{3.}") + r"""
{4.}
{1.}
{2.}
\loigiai{""" + (r"""
\begin{itemize}
	\item (a) \textbf{Sai.} Monosaccharide (glucose, fructose) không bị thủy phân.
	\item (b) \textbf{Đúng.} Tinh bột thủy phân hoàn toàn chỉ tạo glucose.
	\item (c) \textbf{Đúng.} Glucose tráng bạc trực tiếp; fructose tráng bạc trong môi trường base do chuyển hóa thành glucose.
	\item (d) \textbf{Đúng.} Nhóm aldehyde $-\ce{CHO}$ của glucose bị oxi hóa bởi nước bromine.
\end{itemize}
$\Rightarrow$ Có 3 phát biểu đúng: (b), (c), (d).\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các phát biểu sau:\\
(a) Đốt cháy hoàn toàn ester no, đơn chức, mạch hở luôn thu được số mol $\ce{CO2}$ bằng số mol $\ce{H2O}$.\\
(b) Trong hợp chất hữu cơ nhất thiết phải có carbon và hydrogen.\\
(c) Những hợp chất hữu cơ có thành phần nguyên tố giống nhau, thành phần phân tử hơn kém nhau một hay nhiều nhóm $\ce{CH2}$ là đồng đẳng của nhau.\\
(d) Dung dịch glucose bị khử bởi $\ce{AgNO3}$ trong $\ce{NH3}$ tạo ra Ag.\\
(e) Saccharose chỉ có cấu tạo mạch vòng.\\
Số phát biểu đúng là
\choice
{4.}
{5.}
{3.}
""" + (r"{\True 2.}" if is_teacher else r"{2.}") + r"""
\loigiai{""" + (r"""
\begin{itemize}
	\item (a) \textbf{Đúng.} $\ce{C_nH_{2n}O2} + \dots \to n\ce{CO2} + n\ce{H2O}$.
	\item (b) \textbf{Sai.} Hợp chất hữu cơ nhất thiết phải có carbon, có thể không có hydrogen (ví dụ: $\ce{CCl4}, \ce{CF4}$).
	\item (c) \textbf{Sai.} Các chất đồng đẳng bắt buộc phải có cấu tạo hóa học tương tự nhau (tính chất tương tự nhau).
	\item (d) \textbf{Sai.} Glucose đóng vai trò là chất khử (bị oxi hóa); $\ce{Ag^+}$ bị khử thành Ag.
	\item (e) \textbf{Đúng.} Saccharose không có nhóm hemiacetal/hemiketal tự do nên chỉ tồn tại ở dạng mạch vòng.
\end{itemize}
$\Rightarrow$ Có 2 phát biểu đúng: (a), (e).\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các phát biểu sau:\\
(a) Hydrogen hóa hoàn toàn glucose tạo ra acid gluconic.\\
(b) Ở điều kiện thường, glucose và saccharose đều là những chất rắn, dễ tan trong nước.\\
(c) Cellulose trinitrate là nguyên liệu để sản xuất tơ nhân tạo và chế tạo thuốc súng không khói.\\
(d) Amylopectin trong tinh bột chỉ có các liên kết $\alpha$-1,4-glycoside.\\
(e) Saccharose bị hóa đen trong $\ce{H2SO4}$ đặc.\\
(f) Trong công nghiệp dược phẩm, saccharose được dùng để pha chế thuốc.\\
Trong các phát biểu trên, số phát biểu đúng là
\choice
{4.}
""" + (r"{\True 3.}" if is_teacher else r"{3.}") + r"""
{2.}
{5.}
\loigiai{""" + (r"""
\begin{itemize}
	\item (a) \textbf{Sai.} Khử glucose bằng $\ce{H2}$ (xúc tác Ni, $t^\circ$) tạo sorbitol, không phải gluconic acid.
	\item (b) \textbf{Đúng.} Cả hai đều là chất rắn kết tinh, không màu, dễ tan trong nước.
	\item (c) \textbf{Sai.} Cellulose trinitrate dùng làm thuốc súng không khói, không dùng sản xuất tơ nhân tạo (cellulose acetate hoặc viscose mới dùng làm tơ nhân tạo).
	\item (d) \textbf{Sai.} Amylopectin có cấu tạo phân nhánh chứa cả liên kết $\alpha$-1,4 và $\alpha$-1,6-glycosidic.
	\item (e) \textbf{Đúng.} $\ce{H2SO4}$ đặc háo nước làm than hóa saccharose thành C màu đen.
	\item (f) \textbf{Đúng.} Saccharose dùng làm siro, tá dược trong sản xuất thuốc.
\end{itemize}
$\Rightarrow$ Có 3 phát biểu đúng: (b), (e), (f).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\Closesolutionfile{ans}

\vspace{0.15cm}
\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\setcounter{ex}{0}

\begin{ex}
(SBT -- KNTT).
\choiceTF[t]
""" + (r"{\False Tất cả các loại carbohydrate đều tan hoàn toàn trong nước}" if is_teacher else r"{Tất cả các loại carbohydrate đều tan hoàn toàn trong nước}") + r"""
""" + (r"{\True Một số đường đơn (monosaccharide) có vị ngọt}" if is_teacher else r"{Một số đường đơn (monosaccharide) có vị ngọt}") + r"""
""" + (r"{\False Carbohydrate chỉ được tìm thấy trong thực phẩm có nguồn gốc thực vật}" if is_teacher else r"{Carbohydrate chỉ được tìm thấy trong thực phẩm có nguồn gốc thực vật}") + r"""
""" + (r"{\True Glucose và fructose đều là chất rắn, vị ngọt, tan tốt trong nước}" if is_teacher else r"{Glucose và fructose đều là chất rắn, vị ngọt, tan tốt trong nước}") + r"""
\loigiai{""" + (r"""
\begin{itemize}
	\item a) \textbf{Sai.} Tinh bột không tan trong nước lạnh (tan một phần trong nước nóng tạo hồ tinh bột), cellulose không tan trong nước.
	\item b) \textbf{Đúng.} Glucose, fructose đều có vị ngọt.
	\item c) \textbf{Sai.} Carbohydrate còn có trong cơ thể động vật (glycogen trong gan, lactose trong sữa).
	\item d) \textbf{Đúng.} Đều là chất rắn kết tinh, vị ngọt, dễ tan trong nước.
\end{itemize}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Saccharose được tiêu thụ với lượng lớn trên toàn cầu.
\begin{center}
	\includegraphics[width=9.5cm]{images_crop_de05/p2_cau2_cay_mia_thot_not_cu_cai.png}
\end{center}
\choiceTF[t]
""" + (r"{\True Saccharose còn được gọi là đường ăn, là chất rắn, vị ngọt, dễ tan trong nước}" if is_teacher else r"{Saccharose còn được gọi là đường ăn, là chất rắn, vị ngọt, dễ tan trong nước}") + r"""
""" + (r"{\True Saccharose có nhiều trong cây mía, hoa thốt nốt, củ cải đường}" if is_teacher else r"{Saccharose có nhiều trong cây mía, hoa thốt nốt, củ cải đường}") + r"""
""" + (r"{\False Thủy phân hoàn toàn saccharose trong môi trường acid chỉ thu được glucose}" if is_teacher else r"{Thủy phân hoàn toàn saccharose trong môi trường acid chỉ thu được glucose}") + r"""
""" + (r"{\False Cấu tạo phân tử saccharose gồm 1 gốc $\alpha$-glucose và 1 gốc $\beta$-fructose liên kết với nhau qua nguyên tử carbon}" if is_teacher else r"{Cấu tạo phân tử saccharose gồm 1 gốc $\alpha$-glucose và 1 gốc $\beta$-fructose liên kết với nhau qua nguyên tử carbon}") + r"""
\loigiai{""" + (r"""
\begin{itemize}
	\item a) \textbf{Đúng.} Saccharose là đường ăn thông dụng, vị ngọt thanh, dễ tan trong nước.
	\item b) \textbf{Đúng.} Cây mía, củ cải đường và thốt nốt là nguồn cung cấp saccharose chủ yếu.
	\item c) \textbf{Sai.} Thủy phân hoàn toàn thu được hỗn hợp gồm 1 phân tử glucose và 1 phân tử fructose.
	\item d) \textbf{Sai.} Hai gốc monosaccharide liên kết với nhau qua nguyên tử oxygen (liên kết $\alpha$-1,2-glycosidic), không phải qua nguyên tử carbon.
\end{itemize}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cellulose là một polysaccharide gồm nhiều gốc $\beta$-glucose liên kết với nhau tạo ra dạng mạch không nhánh, kéo dài thành dạng sợi có cấu tạo như sau:
\begin{center}
	\includegraphics[width=8.5cm]{images_crop_de05/p2_cau3_cau_truc_cellulose.png}
\end{center}
\choiceTF[t]
""" + (r"{\False Đoạn mạch trên có cả hai loại liên kết là: liên kết $\beta$-1,4-glycoside và liên kết $\beta$-1,6-glycoside}" if is_teacher else r"{Đoạn mạch trên có cả hai loại liên kết là: liên kết $\beta$-1,4-glycoside và liên kết $\beta$-1,6-glycoside}") + r"""
""" + (r"{\True Các gốc $\beta$-glucose trong đoạn mạch trên liên kết với nhau bởi liên kết $\beta$-1,4-glycoside}" if is_teacher else r"{Các gốc $\beta$-glucose trong đoạn mạch trên liên kết với nhau bởi liên kết $\beta$-1,4-glycoside}") + r"""
""" + (r"{\True Công thức mỗi gốc $\beta$-glucose là $\ce{C6H10O5}$}" if is_teacher else r"{Công thức mỗi gốc $\beta$-glucose là $\ce{C6H10O5}$}") + r"""
""" + (r"{\False Đoạn mạch trên có cấu tạo phân nhánh ở liên kết $\alpha$-1,6-glycoside}" if is_teacher else r"{Đoạn mạch trên có cấu tạo phân nhánh ở liên kết $\alpha$-1,6-glycoside}") + r"""
\loigiai{""" + (r"""
\begin{itemize}
	\item a) \textbf{Sai.} Phân tử cellulose mạch thẳng chỉ chứa duy nhất liên kết $\beta$-1,4-glycosidic.
	\item b) \textbf{Đúng.} Các gốc $\beta$-glucose liên kết với nhau qua nguyên tử oxygen giữa C1 của gốc này và C4 của gốc kế tiếp.
	\item c) \textbf{Đúng.} Mỗi mắt xích có công thức $-\ce{C6H10O5}-$.
	\item d) \textbf{Sai.} Cellulose có cấu tạo mạch không phân nhánh.
\end{itemize}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phản ứng thuỷ phân cellulose trong môi trường acid\\
\textit{Chuẩn bị:}\\
Hoá chất: cellulose (bông), dung dịch $\ce{H2SO4}$ 70\%, dung dịch $\ce{NaOH}$ 10\%, dung dịch $\ce{CuSO4}$ 5\%, $\ce{NaHCO3}$ rắn, nước nóng.\\
Dụng cụ: ống nghiệm, cốc thuỷ tinh 250 mL, đũa thuỷ tinh, đèn cồn, giá đựng ống nghiệm.\\
\textit{Tiến hành:}\\
\textit{Bước 1:} Cho 10 mL dung dịch $\ce{H2SO4}$ 70\% vào cốc thuỷ tinh, thêm một lượng nhỏ cellulose (bông) vào cốc và dùng đũa thuỷ tinh khuấy đều. Sau đó, đặt cốc thuỷ tinh vào cốc nước nóng và khuấy trong khoảng 3 phút để cellulose tan hết tạo dung dịch đồng nhất.\\
\textit{Bước 2:} Trung hoà dung dịch bằng cách thêm từ từ $\ce{NaHCO3}$ đến khi dừng sủi bọt khí, sau đó thêm tiếp 5 mL dung dịch $\ce{NaOH}$ 10\%.\\
\textit{Bước 3:} Cho 5 mL dung dịch thu được ở trên vào ống nghiệm chứa $\ce{Cu(OH)2}$ (được điều chế bằng cách cho 0,5 mL dung dịch $\ce{CuSO4}$ 5\% vào 2 mL dung dịch $\ce{NaOH}$ 10\%, lắc nhẹ). Đun nóng đều ống nghiệm khoảng 2 phút, sau đó để ống nghiệm trên giá khoảng 3 phút.
\choiceTF[t]
""" + (r"{\False Sau bước 1, sản phẩm thu được có glucose}" if is_teacher else r"{Sau bước 1, sản phẩm thu được có glucose}") + r"""
""" + (r"{\False Ở bước 2 thêm từ từ $\ce{NaHCO3}$ để phản ứng diễn ra nhanh hơn}" if is_teacher else r"{Ở bước 2 thêm từ từ $\ce{NaHCO3}$ để phản ứng diễn ra nhanh hơn}") + r"""
""" + (r"{\True Sau bước 3, lấy sản phẩm thu được có kết tủa màu đỏ gạch}" if is_teacher else r"{Sau bước 3, lấy sản phẩm thu được có kết tủa màu đỏ gạch}") + r"""
""" + (r"{\True Thí nghiệm trên chứng minh trong cellulose có thể bị thủy phân trong môi trường acid tạo glucose}" if is_teacher else r"{Thí nghiệm trên chứng minh trong cellulose có thể bị thủy phân trong môi trường acid tạo glucose}") + r"""
\loigiai{""" + (r"""
\begin{itemize}
	\item a) \textbf{Sai.} Sau bước 1 trong nước nóng 3 phút cellulose mới bị thủy phân dở dang thành dextrin ngắn, chưa tạo glucose tự do.
	\item b) \textbf{Sai.} Thêm $\ce{NaHCO3}$ để trung hòa lượng acid $\ce{H2SO4}$ dư (để tránh hòa tan kết tủa $\ce{Cu(OH)2}$ ở bước 3).
	\item c) \textbf{Đúng.} Glucose sinh ra khử $\ce{Cu(OH)2}$ khi đun nóng tạo kết tủa đỏ gạch $\ce{Cu2O}$.
	\item d) \textbf{Đúng.} Phản ứng chứng minh cellulose bị thủy phân trong môi trường acid sinh ra monosaccharide có tính khử (glucose).
\end{itemize}""" if is_teacher else "") + r"""}
\end{ex}

\vspace{0.15cm}
\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6. Mỗi câu hỏi thí sinh điền câu trả lời ngắn gọn theo yêu cầu.}

\setcounter{ex}{0}

\begin{ex}
Cho dãy các chất: tinh bột, saccharose, glucose, fructose, cellulose. Số chất trong dãy thuộc loại monosaccharide là bao nhiêu?
\loigiai{""" + (r"""
Có 2 chất là glucose và fructose (saccharose là disaccharide; tinh bột và cellulose là polysaccharide).\\
$\Rightarrow$ Có \textbf{2} chất.\\
\textbf{Đáp số: 2.}""" if is_teacher else "") + r"""}
""" + (r"\noindent\textbf{Đáp số: \color{hfRed}{2}}" if is_teacher else r"\noindent\textbf{Đáp số:} \dotfill") + r"""
\end{ex}

\begin{ex}
Cho các tính chất sau: (1) có vị ngọt, (2) dễ tan trong nước, (3) có phản ứng tráng bạc, (4) bị thủy phân trong môi trường acid, (5) hòa tan $\ce{Cu(OH)2}$ thành dung dịch màu xanh lam. Số tính chất đúng với saccharose là bao nhiêu?
\loigiai{""" + (r"""
Saccharose có 4 tính chất đúng: (1), (2), (4), (5). Tính chất (3) không có vì không chứa nhóm $-\ce{CHO}$ tự do.\\
$\Rightarrow$ Có \textbf{4} tính chất đúng.\\
\textbf{Đáp số: 4.}""" if is_teacher else "") + r"""}
""" + (r"\noindent\textbf{Đáp số: \color{hfRed}{4}}" if is_teacher else r"\noindent\textbf{Đáp số:} \dotfill") + r"""
\end{ex}

\begin{ex}
(SBT -- KNTT). Trong số các chất saccharose, maltose, tinh bột và cellulose, có bao nhiêu chất khi thuỷ phân hoàn toàn sản phẩm thu được chỉ là glucose?
\loigiai{""" + (r"""
Có 3 chất khi thủy phân hoàn toàn chỉ thu được glucose là:\\
- Maltose: $\ce{C12H22O11 + H2O -> 2C6H12O6}$ (glucose)\\
- Tinh bột: $(\ce{C6H10O5})_n + n\ce{H2O} \to n\ce{C6H12O6}$ (glucose)\\
- Cellulose: $(\ce{C6H10O5})_n + n\ce{H2O} \to n\ce{C6H12O6}$ (glucose)\\
(Saccharose thủy phân tạo 1 glucose và 1 fructose nên không thỏa mãn).\\
$\Rightarrow$ Có \textbf{3} chất.\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
""" + (r"\noindent\textbf{Đáp số: \color{hfRed}{3}}" if is_teacher else r"\noindent\textbf{Đáp số:} \dotfill") + r"""
\end{ex}

\begin{ex}
Giả thiết phản ứng tổng hợp glucose ở cây lương thực như sau:
\[
\ce{6CO2 + 6H2O} \xrightarrow[\text{chlorophyll}]{\text{asmt}} \ce{C6H12O6 + 6O2} \quad \Delta_r H^\circ_{298} = 2800\text{ kJ}
\]
Trung bình mỗi phút, $1\text{ m}^2$ mặt đất nhận được từ Mặt Trời $60\text{ kJ}$ nhưng chỉ $1\%$ được hấp thụ ở quá trình quang hợp. Khối lượng glucose tổng hợp được trong 7 giờ trên 1 sào ($360\text{ m}^2$) trồng cây lương thực là bao nhiêu gam?
\loigiai{""" + (r"""
- Tổng năng lượng ánh sáng hấp thụ trong 7 giờ trên 1 sào ($360\text{ m}^2$):\\
  $Q = 360 \times (7 \times 60) \times 60 \times 1\% = 90720\text{ kJ}$.\\
- Số mol glucose tổng hợp được:\\
  $n_{\text{glucose}} = \dfrac{90720}{2800} = 32{,}4\text{ mol}$.\\
- Khối lượng glucose:\\
  $m_{\text{glucose}} = 32{,}4 \times 180 = \mathbf{5832\text{ gam}}$.\\
\textbf{Đáp số: 5832.}""" if is_teacher else "") + r"""}
""" + (r"\noindent\textbf{Đáp số: \color{hfRed}{5832}}" if is_teacher else r"\noindent\textbf{Đáp số:} \dotfill") + r"""
\end{ex}

\begin{ex}
Để tráng bạc (silver) một số ruột phích, người ta phải thuỷ phân $34{,}2\text{ gam}$ saccharose rồi tiến hành phản ứng tráng gương. Có $m\text{ gam}$ Ag tạo thành sau phản ứng, biết hiệu suất mỗi quá trình là $70\%$. Xác định giá trị $m$? (Làm tròn kết quả đến chữ số thập phân thứ nhất).
\loigiai{""" + (r"""
- $n_{\text{saccharose}} = \dfrac{34{,}2}{342} = 0{,}1\text{ mol}$.\\
- Sơ đồ phản ứng:\\
  $\ce{C12H22O11} \xrightarrow{\text{thủy phân}} 1\text{ glucose} + 1\text{ fructose} \xrightarrow{\text{tráng gương}} 4\ce{Ag}$.\\
- Do hiệu suất mỗi quá trình là $70\%$ ($0{,}7$):\\
  $n_{\ce{Ag}} = 0{,}1 \times 4 \times 0{,}7 \times 0{,}7 = 0{,}196\text{ mol}$.\\
- Khối lượng Ag thu được:\\
  $m_{\ce{Ag}} = 0{,}196 \times 108 = 21{,}168\text{ gam} \approx \mathbf{21{,}2\text{ gam}}$.\\
\textbf{Đáp số: 21,2.}""" if is_teacher else "") + r"""}
""" + (r"\noindent\textbf{Đáp số: \color{hfRed}{21,2}}" if is_teacher else r"\noindent\textbf{Đáp số:} \dotfill") + r"""
\end{ex}

\begin{ex}
Xăng E5 được sản xuất bằng cách phối trộn xăng khoáng RON 92 với ethanol ($D = 0{,}8\text{ gam/mL}$) theo tỉ lệ thể tích tương ứng là $95 : 5$. Ethanol được sản xuất từ tinh bột bằng phương pháp lên men. Tinh bột có nhiều trong sắn, ngô. Cho bảng thông tin sau:
\vspace{0.1cm}

\noindent
\begin{minipage}{0.54\textwidth}
\small
\begin{tabular}{|l|c|c|}
\hline
\textbf{Nguyên liệu} & \textbf{Sắn tươi} & \textbf{Ngô khô} \\ \hline
Hàm lượng tinh bột & 30\% & 75\% \\ \hline
Giá 1 kg nguyên liệu & 1.850 đ & 6.500 đ \\ \hline
\end{tabular}
\end{minipage}\hfill
\begin{minipage}{0.44\textwidth}
\centering
\includegraphics[width=5.5cm]{images_crop_de05/p3_cau6_san_va_ngo.png}
\end{minipage}

\vspace{0.1cm}
Biết hiệu suất của toàn bộ quá trình sản xuất từ sắn tươi đạt $79{,}83\%$, từ ngô khô đạt $84{,}783\%$. Để điều chế $50.000\text{ lít}$ xăng E5 từ sắn tươi cần chi $x$ triệu đồng, từ ngô khô cần chi $y$ triệu đồng. Tính chênh lệch tiền nguyên liệu là $z$ (triệu đồng) khi dùng sắn tươi so với dùng ngô khô để sản xuất. Tính giá trị của $z$ (làm tròn kết quả đến phần triệu).
\loigiai{""" + (r"""
- Thể tích ethanol trong $50.000\text{ lít}$ xăng E5: $V_{\text{ethanol}} = 50000 \times 5\% = 2500\text{ lít}$.\\
- Khối lượng ethanol: $m_{\text{ethanol}} = 2500 \times 0{,}8 = 2000\text{ kg}$.\\
- Quá trình lên men: $(\ce{C6H10O5})_n \to n\ce{C6H12O6} \to 2n\ce{C2H5OH}$.\\
  Khối lượng tinh bột lý thuyết cần: $m_{\text{tinh bột (LT)}} = 2000 \times \dfrac{162}{92} \approx 3521{,}739\text{ kg}$.\\
- Trường hợp dùng sắn tươi:\\
  $x = 3521{,}739 \times \dfrac{100}{30} \times \dfrac{100}{79{,}83} \times 1850 \approx 27{,}203 \times 10^6\text{ đồng} \approx 27\text{ triệu đồng}$.\\
- Trường hợp dùng ngô khô:\\
  $y = 3521{,}739 \times \dfrac{100}{75} \times \dfrac{100}{84{,}783} \times 6500 \approx 36{,}000 \times 10^6\text{ đồng} \approx 36\text{ triệu đồng}$.\\
- Chênh lệch chi phí nguyên liệu:\\
  $z = |y - x| = 36 - 27 = \mathbf{9\text{ triệu đồng}}$.\\
\textbf{Đáp số: 9.}""" if is_teacher else "") + r"""}
""" + (r"\noindent\textbf{Đáp số: \color{hfRed}{9}}" if is_teacher else r"\noindent\textbf{Đáp số:} \dotfill") + r"""
\end{ex}

"""
    if is_teacher:
        tex += r"""
\newpage
\noindent
{\color{hfRed}\textbf{\Large BẢNG ĐÁP ÁN ĐỀ KIỂM TRA SỐ 5 (CHƯƠNG 2: CARBOHYDRATE)}}

\vspace{0.3cm}
\noindent\textbf{PHẦN I:}
\begin{center}
\input{""" + ans_file + r"""}
\end{center}

\vspace{0.2cm}
\noindent\textbf{PHẦN II:}
\begin{center}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Ý a} & \textbf{Ý b} & \textbf{Ý c} & \textbf{Ý d} \\ \hline
\textbf{1} & S & Đ & S & Đ \\ \hline
\textbf{2} & Đ & Đ & S & S \\ \hline
\textbf{3} & S & Đ & Đ & S \\ \hline
\textbf{4} & S & S & Đ & Đ \\ \hline
\end{tabular}
\end{center}

\vspace{0.2cm}
\noindent\textbf{PHẦN III:}
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\ \hline
2 & 4 & 3 & 5832 & 21,2 & 9 \\ \hline
\end{tabular}
\end{center}
"""

    tex += r"""
\par\vspace{0.1cm}
\centerline{\textbf{---------- HẾT ----------}}

\end{document}
"""
    return tex

def compile_pdf(tex_content, filename_base):
    tex_path = os.path.join(TEX_DIR, f"{filename_base}.tex")
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_content)
    
    pdflatex_cmd = r"C:\Users\Admin\AppData\Local\Programs\MiKTeX\miktex\bin\x64\pdflatex.exe"
    res = None
    for _ in range(2):
        res = subprocess.run([pdflatex_cmd, "-interaction=nonstopmode", f"{filename_base}.tex"],
                             cwd=TEX_DIR, capture_output=True, encoding="utf-8", errors="ignore")
    
    # Cleanup auxiliary files
    for ext in [".aux", ".log", ".thm"]:
        junk = os.path.join(TEX_DIR, f"{filename_base}{ext}")
        if os.path.exists(junk):
            try:
                os.remove(junk)
            except:
                pass
    for f in os.listdir(TEX_DIR):
        if f.startswith("ans-") and f.endswith(".tex"):
            try:
                os.remove(os.path.join(TEX_DIR, f))
            except:
                pass
    print(f"Compiled: {filename_base}.pdf (code {res.returncode if res else 'None'})")

if __name__ == "__main__":
    os.makedirs(TEX_DIR, exist_ok=True)
    
    # 1. Học sinh
    tex_hs = make_tex(is_teacher=False)
    compile_pdf(tex_hs, "De_Kiem_Tra_Hoa_12_De05_HocSinh")
    
    # 2. Giáo viên
    tex_gv = make_tex(is_teacher=True)
    compile_pdf(tex_gv, "De_Kiem_Tra_Hoa_12_De05_GiaoVien")
