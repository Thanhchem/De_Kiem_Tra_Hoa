import os
import subprocess
import shutil

TEX_DIR = r"c:\Antigravity_Thanh\San_Pham"

# Ensure ex_test.sty is present in San_Pham
if not os.path.exists(os.path.join(TEX_DIR, "ex_test.sty")):
    shutil.copyfile(r"c:\Antigravity_Thanh\ex_test.sty", os.path.join(TEX_DIR, "ex_test.sty"))

def make_tex(is_teacher=False):
    ver_tag = "gv" if is_teacher else "hs"
    ans_file = f"ans-de01{ver_tag}"
    
    tex = r"""\documentclass[12pt,a4paper]{article}
\usepackage[utf8]{vietnam}
\usepackage{amsmath,amssymb}
\usepackage{geometry}
\geometry{top=1.5cm,bottom=1.2cm,left=1.4cm,right=1.4cm,headheight=28pt,headsep=12pt,footskip=18pt}
\usepackage{tikz}
\usepackage{xcolor}
\usepackage{enumerate}
\usepackage{graphicx}
\usepackage[version=4]{mhchem}
\usepackage{fancyhdr}
\usepackage{ex_test}

\renewcommand{\baselinestretch}{0.96}
\setlength{\parskip}{0pt}

\definecolor{hfRed}{RGB}{211,47,47}
\definecolor{hfGreen}{RGB}{139,195,144}

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
	\textbf{TRƯỜNG THPT CHUYÊN QUỐC HỌC}\\
	\textbf{ĐỀ CHÍNH THỨC}\\
	\textit{(Đề thi có 04 trang)}
\end{minipage}\hfill
\begin{minipage}[t]{0.54\textwidth}
	\centering\small
	\textbf{ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 1 (ESTER -- LIPID)}\\
	\textbf{MÔN: HOÁ HỌC -- LỚP 12}\\
	\textit{Thời gian làm bài: 50 phút (không kể thời gian phát đề)}\\
	\fbox{\textbf{MÃ ĐỀ: 101}}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; K = 39; Br = 80).}
\smallskip

\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\Opensolutionfile{ans}[""" + ans_file + r"""]

\begin{ex}
Chất nào sau đây thuộc loại ester?
\choice
""" + (r"{\True $\ce{CH3COOC2H5}$}" if is_teacher else r"{$\ce{CH3COOC2H5}$}") + r"""
{$\ce{HOOCCH3}$}
{$\ce{H2N-CH2-COOH}$}
{$\ce{CH3CHO}$}
\loigiai{""" + (r"""
$\ce{CH3COOC2H5}$ (ethyl acetate) thuộc loại ester vì chứa nhóm chức $-\ce{COO}-$ liên kết với gốc hydrocarbon.\\
$\ce{HOOCCH3}$ là carboxylic acid, $\ce{H2N-CH2-COOH}$ là amino acid, $\ce{CH3CHO}$ là aldehyde.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phản ứng điều chế xà phòng từ chất béo được gọi là phản ứng
\choice
{ester hóa}
""" + (r"{\True xà phòng hóa}" if is_teacher else r"{xà phòng hóa}") + r"""
{trung hòa}
{hydrate hóa}
\loigiai{""" + (r"""
Phản ứng thủy phân chất béo trong môi trường kiềm (dung dịch $\ce{NaOH}$ hoặc $\ce{KOH}$) đun nóng để điều chế xà phòng và glycerol được gọi là \textbf{phản ứng xà phòng hóa}.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Dầu chuối là ester có tên isoamyl acetate, được điều chế từ
\choice
{$\ce{CH3OH}$, $\ce{CH3COOH}$}
{$\ce{(CH3)2CH-CH2OH}$, $\ce{CH3COOH}$}
{$\ce{C2H5COOH}$, $\ce{C2H5OH}$}
""" + (r"{\True $\ce{CH3COOH}$, $\ce{(CH3)2CH-CH2-CH2OH}$}" if is_teacher else r"{$\ce{CH3COOH}$, $\ce{(CH3)2CH-CH2-CH2OH}$}") + r"""
\loigiai{""" + (r"""
Isoamyl acetate ($\ce{CH3COOCH2CH2CH(CH3)2}$) được tổng hợp từ phản ứng ester hóa giữa acetic acid ($\ce{CH3COOH}$) và isoamyl alcohol ($\ce{(CH3)2CH-CH2-CH2OH}$).\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Một số ester được dùng trong hương liệu, mĩ phẩm, bột giặt là nhờ các ester
\choice
{là chất lỏng dễ bay hơi}
""" + (r"{\True có mùi thơm, an toàn với người}" if is_teacher else r"{có mùi thơm, an toàn với người}") + r"""
{có thể bay hơi nhanh sau khi sử dụng}
{đều có nguồn gốc từ thiên nhiên}
\loigiai{""" + (r"""
Nhờ có mùi thơm hoa quả đặc trưng dễ chịu và an toàn cho sức khỏe con người nên nhiều ester được dùng làm chất tạo hương.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Thủy phân ester nào sau đây trong dung dịch $\ce{NaOH}$ thu được sodium formate?
\choice
{$\ce{CH3COOCH3}$}
{$\ce{CH3COOC2H5}$}
""" + (r"{\True $\ce{HCOOC2H5}$}" if is_teacher else r"{$\ce{HCOOC2H5}$}") + r"""
{$\ce{CH3COOC3H7}$}
\loigiai{""" + (r"""
Phương trình: $\ce{HCOOC2H5 + NaOH ->[t^\circ] HCOONa + C2H5OH}$. $\ce{HCOONa}$ là sodium formate.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Số đồng phân ester ứng với công thức phân tử $\ce{C4H8O2}$ là
\choice
{2}
{3}
""" + (r"{\True 4}" if is_teacher else r"{4}") + r"""
{5}
\loigiai{""" + (r"""
$\ce{C4H8O2}$ có 4 đồng phân ester:\\
(1) $\ce{HCOOCH2CH2CH3}$ (propyl formate); (2) $\ce{HCOOCH(CH3)2}$ (isopropyl formate);\\
(3) $\ce{CH3COOCH2CH3}$ (ethyl acetate); (4) $\ce{C2H5COOCH3}$ (methyl propionate).\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Công thức của tristearin là
\choice
{$\ce{(C2H5COO)3C3H5}$}
""" + (r"{\True $\ce{(C17H35COO)3C3H5}$}" if is_teacher else r"{$\ce{(C17H35COO)3C3H5}$}") + r"""
{$\ce{(CH3COO)3C3H5}$}
{$\ce{(HCOO)3C3H5}$}
\loigiai{""" + (r"""
Tristearin là triester của stearic acid ($\ce{C17H35COOH}$) và glycerol: $\ce{(C17H35COO)3C3H5}$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Thực hiện phản ứng ester hoá giữa $\ce{HOOC-COOH}$ với hỗn hợp $\ce{CH3OH}$ và $\ce{C2H5OH}$ thu được tối đa bao nhiêu ester hai chức?
\choice
{2}
""" + (r"{\True 3}" if is_teacher else r"{3}") + r"""
{1}
{4}
\loigiai{""" + (r"""
Acid hai chức $\ce{HOOC-COOH}$ tạo tối đa 3 ester hai chức:\\
(1) $\ce{CH3OOC-COOCH3}$, (2) $\ce{C2H5OOC-COOC2H5}$, (3) $\ce{CH3OOC-COOC2H5}$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Công thức phân tử của oleic acid là
\choice
{$\ce{C2H5COOH}$}
{$\ce{HCOOOH}$}
{$\ce{CH3COOH}$}
""" + (r"{\True $\ce{C17H33COOH}$}" if is_teacher else r"{$\ce{C17H33COOH}$}") + r"""
\loigiai{""" + (r"""
Oleic acid là acid béo không no đơn chức mang 1 liên kết đôi $\ce{C=C}$, công thức là $\ce{C17H33COOH}$.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Công thức của triolein là
\choice
""" + (r"{\True $\ce{(C17H33COO)3C3H5}$}" if is_teacher else r"{$\ce{(C17H33COO)3C3H5}$}") + r"""
{$\ce{(HCOO)3C3H5}$}
{$\ce{(C2H5COO)3C3H5}$}
{$\ce{(CH3COO)3C3H5}$}
\loigiai{""" + (r"""
Triolein là triglyceride của oleic acid ($\ce{C17H33COOH}$), công thức là $\ce{(C17H33COO)3C3H5}$.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các chất sau: (1) alcohol ethylic, (2) acetic acid, (3) nước, (4) methyl formate. Thứ tự nhiệt độ sôi giảm dần là
\choice
{(1) $>$ (4) $>$ (3) $>$ (2)}
{(1) $>$ (2) $>$ (3) $>$ (4)}
{(1) $>$ (3) $>$ (2) $>$ (4)}
""" + (r"{\True (2) $>$ (3) $>$ (1) $>$ (4)}" if is_teacher else r"{(2) $>$ (3) $>$ (1) $>$ (4)}") + r"""
\loigiai{""" + (r"""
Nhiệt độ sôi: Acetic acid ($117{,}9^\circ\text{C}$) $>$ Nước ($100^\circ\text{C}$) $>$ Alcohol ethylic ($78{,}4^\circ\text{C}$) $>$ Methyl formate ($31{,}5^\circ\text{C}$).\\
Thứ tự giảm dần: (2) $>$ (3) $>$ (1) $>$ (4).\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Xà phòng và chất giặt rửa có đặc điểm chung nào sau đây?
\choice
{Không tan trong nước}
{Là muối sodium hoặc potassium của acid béo}
{Là muối sulfonate hoặc sulfate của acid béo}
""" + (r"{\True Thường có cấu tạo gồm hai phần là phần không phân cực (kị nước) và phần phân cực (ưa nước)}" if is_teacher else r"{Thường có cấu tạo gồm hai phần là phần không phân cực (kị nước) và phần phân cực (ưa nước)}") + r"""
\loigiai{""" + (r"""
Phân tử xà phòng và chất giặt rửa đều có cấu tạo lưỡng cực gồm: phần đầu phân cực (ưa nước) và phần đuôi hydrocarbon dài không phân cực (kị nước, ưa dầu mỡ).\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các chất sau: $\ce{CH3[CH2]7CH=CH[CH2]7COONa}$, $\ce{CH3[CH2]14COOK}$, $\ce{CH3[CH2]10COOK}$ và $\ce{CH3COONa}$. Trong các chất nêu trên, có bao nhiêu chất có thể là thành phần chính của xà phòng?
\choice
{1}
{2}
""" + (r"{\True 3}" if is_teacher else r"{3}") + r"""
{4}
\loigiai{""" + (r"""
Xà phòng là muối sodium hoặc potassium của acid béo (mạch C dài từ 12--24 nguyên tử C):\\
• $\ce{CH3[CH2]7CH=CH[CH2]7COONa}$ (18C) $\rightarrow$ Đúng;\\
• $\ce{CH3[CH2]14COOK}$ (16C) $\rightarrow$ Đúng;\\
• $\ce{CH3[CH2]10COOK}$ (12C) $\rightarrow$ Đúng;\\
• $\ce{CH3COONa}$ (2C) $\rightarrow$ Mạch quá ngắn, không có tính giặt rửa.\\
Vậy có 3 chất thỏa mãn.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phát biểu nào sau đây về xà phòng là đúng?
\choice
{Xà phòng có thành phần chính là muối sodium hoặc potassium của carboxylic acid}
{Các phân tử xà phòng đều có đầu kị nước gắn với đuôi dài ưa nước}
""" + (r"{\True Xà phòng mất tính giặt rửa khi sử dụng với nước cứng}" if is_teacher else r"{Xà phòng mất tính giặt rửa khi sử dụng với nước cứng}") + r"""
{Nhược điểm của xà phòng là khó bị phân huỷ hoặc phân huỷ chậm, do đó gây hại cho hệ sinh thái}
\loigiai{""" + (r"""
Trong nước cứng chứa $\ce{Ca^{2+}}, \ce{Mg^{2+}}$ tạo kết tủa $(\ce{RCOO})_2\ce{Ca}$ và $(\ce{RCOO})_2\ce{Mg}$, làm mất bọt và tính giặt rửa của xà phòng.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong số các vật phẩm tiêu dùng sau: xà phòng bánh, dầu gội đầu, nước bồ kết và baking soda ($\ce{NaHCO3}$), số vật phẩm có thành phần chất giặt rửa tự nhiên và tổng hợp là
\choice
{1}
""" + (r"{\True 2}" if is_teacher else r"{2}") + r"""
{3}
{4}
\loigiai{""" + (r"""
Nước bồ kết chứa chất giặt rửa tự nhiên (saponin); dầu gội đầu chứa chất giặt rửa tổng hợp (sodium lauryl sulfate,...). Có 2 vật phẩm.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Loại dầu nào sau đây không phải là chất béo?
\choice
{dầu vừng}
{dầu oliu}
{dầu gan cá}
""" + (r"{\True dầu luyn}" if is_teacher else r"{dầu luyn}") + r"""
\loigiai{""" + (r"""
Dầu luyn là hỗn hợp hydrocarbon từ dầu mỏ dùng bôi trơn máy móc. Dầu vừng, dầu oliu, dầu gan cá là chất béo.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các phản ứng sau:\\
(1) Thuỷ phân ester trong môi trường acid.\\
(2) Thuỷ phân ester trong dung dịch $\ce{NaOH}$, đun nóng.\\
(3) Cho ester tác dụng với dung dịch $\ce{KOH}$, đun nóng.\\
(4) Thuỷ phân dẫn xuất halogen trong dung dịch $\ce{NaOH}$, đun nóng.\\
(5) Cho carboxylic acid tác dụng với dung dịch $\ce{NaOH}$.\\
Những phản ứng nào không được gọi là phản ứng xà phòng hoá?
\choice
{(1), (2), (3), (4)}
""" + (r"{\True (1), (4), (5)}" if is_teacher else r"{(1), (4), (5)}") + r"""
{(1), (3), (4), (5)}
{(3), (4), (5)}
\loigiai{""" + (r"""
Phản ứng xà phòng hóa là phản ứng thủy phân ester trong dung dịch kiềm ($\ce{NaOH, KOH}$) đun nóng $\rightarrow$ (2) và (3) là phản ứng xà phòng hóa. Các phản ứng (1), (4), (5) không phải phản ứng xà phòng hóa.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Đun nóng acid acetic với isoamyl alcohol $\ce{(CH3)2CH-CH2CH2OH}$ có $\ce{H2SO4}$ đặc xúc tác thu được isoamyl acetate (dầu chuối). Tính lượng dầu chuối thu được từ $132{,}35\text{ gam}$ acid acetic đun nóng với $200\text{ gam}$ isoamyl alcohol (biết hiệu suất phản ứng đạt 68\%).
\choice
{$97{,}5\text{ gam}$}
""" + (r"{\True $195{,}0\text{ gam}$}" if is_teacher else r"{$195{,}0\text{ gam}$}") + r"""
{$292{,}5\text{ gam}$}
{$159{,}0\text{ gam}$}
\loigiai{""" + (r"""
Phương trình: $\ce{CH3COOH + (CH3)2CHCH2CH2OH <=> CH3COOCH2CH2CH(CH3)2 + H2O}$.\\
$n_{\text{acid}} = \dfrac{132{,}35}{60} \approx 2{,}206\text{ mol}$; $n_{\text{alcohol}} = \dfrac{200}{88} \approx 2{,}273\text{ mol} > n_{\text{acid}}$.\\
$\Rightarrow$ Hiệu suất tính theo $\ce{CH3COOH}$: $n_{\text{ester}} = 2{,}206 \times 68\% = 1{,}50\text{ mol}$.\\
$m_{\text{ester}} = 1{,}50 \times 130 = \mathbf{195{,}0\text{ gam}}$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\Closesolutionfile{ans}

\vspace{0.3cm}
\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\setcounter{ex}{0}

\begin{ex}
Cho các triglyceride X, Y với công thức cấu tạo sau:
\begin{center}
	\includegraphics[width=14cm]{images_crop_de01/triglyceride_XY.png}
\end{center}
Em hãy cho biết các phát biểu sau đây là đúng hay sai:
\choiceTF[t]
""" + (r"{\False Triglyceride X có tên gọi là tripalmitin.}" if is_teacher else r"{Triglyceride X có tên gọi là tripalmitin.}") + r"""
""" + (r"{\True X là chất béo no, Y là chất béo không no.}" if is_teacher else r"{X là chất béo no, Y là chất béo không no.}") + r"""
""" + (r"{\False X, Y đều tan tốt trong nước.}" if is_teacher else r"{X, Y đều tan tốt trong nước.}") + r"""
""" + (r"{\True Hydrogen hoá Y thu được X.}" if is_teacher else r"{Hydrogen hoá Y thu được X.}") + r"""
\loigiai{""" + (r"""
\begin{itemize}
	\item a) \textbf{Sai.} Gốc acid béo trong X có 17 carbon ở đuôi hydrocarbon ($\ce{C17H35COO-}$), do đó X là tristearin ($\ce{(C17H35COO)3C3H5}$), không phải tripalmitin.
	\item b) \textbf{Đúng.} Gốc stearate không có liên kết $\pi$ mạch C nên X là chất béo no; gốc oleate có 1 nối đôi $\ce{C=C}$ nên Y là chất béo không no.
	\item c) \textbf{Sai.} Chất béo hầu như không tan trong nước do phân tử không phân cực và nhẹ hơn nước.
	\item d) \textbf{Đúng.} Hydrogen hóa hoàn toàn triolein (Y) với xúc tác Ni, $t^\circ$ sẽ cộng $\ce{H2}$ tạo thành tristearin (X).
\end{itemize}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Nhiệt độ sôi và độ tan của một số ester, carboxylic acid và alcohol có cùng số nguyên tử carbon được cho trong bảng sau:
\begin{center}
\begin{tabular}{|l|c|c|}
\hline
\textbf{Công thức} & \textbf{Nhiệt độ sôi ($^\circ\text{C}$)} & \textbf{Độ tan ở 25 $^\circ\text{C}$ (g/100 g nước)} \\ \hline
$\ce{HCOOCH3}$ & $31{,}5$ & $23{,}0$ \\ \hline
$\ce{HCOOC2H5}$ & $54{,}2$ & $12{,}0$ \\ \hline
$\ce{CH3COOH}$ & $117{,}9$ & Tan vô hạn \\ \hline
$\ce{C2H5COOH}$ & $141{,}0$ & Tan vô hạn \\ \hline
$\ce{C2H5OH}$ & $78{,}4$ & Tan vô hạn \\ \hline
$\ce{CH3CH2CH2OH}$ & $97{,}2$ & Tan vô hạn \\ \hline
\end{tabular}
\end{center}
Em hãy cho biết các phát biểu sau đây là đúng hay sai:
\choiceTF[t]
""" + (r"{\True Do không tạo được liên kết hydrogen giữa các phân tử nên ester có nhiệt độ sôi thấp hơn nhiệt độ sôi của carboxylic acid và alcohol có cùng số nguyên tử carbon.}" if is_teacher else r"{Do không tạo được liên kết hydrogen giữa các phân tử nên ester có nhiệt độ sôi thấp hơn nhiệt độ sôi của carboxylic acid và alcohol có cùng số nguyên tử carbon.}") + r"""
""" + (r"{\True Do có khả năng tạo liên kết hydrogen yếu với nước nên ester thường ít tan trong nước hơn so với carboxylic acid và alcohol có cùng số carbon.}" if is_teacher else r"{Do có khả năng tạo liên kết hydrogen yếu với nước nên ester thường ít tan trong nước hơn so với carboxylic acid và alcohol có cùng số carbon.}") + r"""
""" + (r"{\True Carboxylic acid có nhiệt độ sôi cao hơn alcohol có cùng số nguyên tử carbon.}" if is_teacher else r"{Carboxylic acid có nhiệt độ sôi cao hơn alcohol có cùng số nguyên tử carbon.}") + r"""
""" + (r"{\True Methanol có khả năng tan vô hạn trong nước.}" if is_teacher else r"{Methanol có khả năng tan vô hạn trong nước.}") + r"""
\loigiai{""" + (r"""
\begin{itemize}
	\item a) \textbf{Đúng.} Phân tử ester không tạo được liên kết hydrogen liên phân tử nên nhiệt độ sôi thấp hơn nhiều so với acid và alcohol.
	\item b) \textbf{Đúng.} Ester chỉ nhận liên kết hydrogen từ nước, không tự cho liên kết hydrogen, do đó độ tan kém hơn hẳn acid và alcohol.
	\item c) \textbf{Đúng.} Phân tử acid tạo liên kết hydrogen dạng nhị hợp bền hơn dạng chuỗi của alcohol.
	\item d) \textbf{Đúng.} Methanol ($\ce{CH3OH}$) tạo liên kết hydrogen rất mạnh với nước nên tan vô hạn ở mọi tỉ lệ.
\end{itemize}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Các phát biểu sau đây về xà phòng và chất giặt rửa là đúng hay sai?
\choiceTF[t]
""" + (r"{\True Xà phòng và chất giặt rửa thường có cấu tạo gồm hai phần: ưa nước và kị nước.}" if is_teacher else r"{Xà phòng và chất giặt rửa thường có cấu tạo gồm hai phần: ưa nước và kị nước.}") + r"""
""" + (r"{\False Xà phòng hoá tripalmitin với dung dịch $\ce{NaOH}$ thu được sản phẩm là $\ce{C15H29COONa}$ và glycerol.}" if is_teacher else r"{Xà phòng hoá tripalmitin với dung dịch $\ce{NaOH}$ thu được sản phẩm là $\ce{C15H29COONa}$ và glycerol.}") + r"""
""" + (r"{\False Chất giặt rửa tổng hợp thường được điều chế từ chất béo.}" if is_teacher else r"{Chất giặt rửa tổng hợp thường được điều chế từ chất béo.}") + r"""
""" + (r"{\True Mỡ động vật, dầu thực vật là nguyên liệu để sản xuất xà phòng.}" if is_teacher else r"{Mỡ động vật, dầu thực vật là nguyên liệu để sản xuất xà phòng.}") + r"""
\loigiai{""" + (r"""
\begin{itemize}
	\item a) \textbf{Đúng.} Đều có cấu trúc phân tử lưỡng cực gồm đầu ưa nước và đuôi kị nước.
	\item b) \textbf{Sai.} Thủy phân tripalmitin thu được sodium palmitate là $\ce{C15H31COONa}$, không phải $\ce{C15H29COONa}$.
	\item c) \textbf{Sai.} Chất giặt rửa tổng hợp được điều chế từ dầu mỏ, không phải từ chất béo.
	\item d) \textbf{Đúng.} Dầu thực vật và mỡ động vật là nguồn nguyên liệu tự nhiên để nấu xà phòng.
\end{itemize}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Các phát biểu sau đây là đúng hay sai?
\choiceTF[t]
""" + (r"{\True Chất giặt rửa thường là muối sodium alkylsulfate hoặc alkylbenzene sulfonate.}" if is_teacher else r"{Chất giặt rửa thường là muối sodium alkylsulfate hoặc alkylbenzene sulfonate.}") + r"""
""" + (r"{\True Phân tử chất giặt rửa gồm một đầu kị nước gắn với một đầu ưa nước.}" if is_teacher else r"{Phân tử chất giặt rửa gồm một đầu kị nước gắn với một đầu ưa nước.}") + r"""
""" + (r"{\False Khi giặt rửa bằng nước cứng nên sử dụng xà phòng.}" if is_teacher else r"{Khi giặt rửa bằng nước cứng nên sử dụng xà phòng.}") + r"""
""" + (r"{\True Phản ứng thuỷ phân chất béo trong môi trường kiềm ($\ce{NaOH, KOH}$) thuộc loại phản ứng xà phòng hoá.}" if is_teacher else r"{Phản ứng thuỷ phân chất béo trong môi trường kiềm ($\ce{NaOH, KOH}$) thuộc loại phản ứng xà phòng hoá.}") + r"""
\loigiai{""" + (r"""
\begin{itemize}
	\item a) \textbf{Đúng.} Đây là 2 loại chất giặt rửa tổng hợp thông dụng nhất.
	\item b) \textbf{Đúng.} Cấu tạo lưỡng cực đặc trưng của chất hoạt động bề mặt.
	\item c) \textbf{Sai.} Nước cứng làm kết tủa xà phòng; trường hợp này nên sử dụng chất giặt rửa tổng hợp.
	\item d) \textbf{Đúng.} Phản ứng thủy phân chất béo trong dung dịch kiềm được gọi là phản ứng xà phòng hóa.
\end{itemize}""" if is_teacher else "") + r"""}
\end{ex}

\vspace{0.3cm}
\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6. Mỗi câu hỏi thí sinh điền câu trả lời ngắn gọn theo yêu cầu.}

\setcounter{ex}{0}

\begin{ex}
Khi xà phòng hóa triglyceride X bằng dung dịch $\ce{NaOH}$ dư, đun nóng, thu được sản phẩm gồm glycerol, sodium oleate, sodium stearate và sodium palmitate. Số đồng phân cấu tạo thỏa mãn tính chất trên của X là bao nhiêu?
\loigiai{""" + (r"""
Triglyceride chứa 3 gốc acid béo khác nhau gắn vào khung glycerol. Số đồng phân cấu tạo là: $\dfrac{3!}{2} = \mathbf{3}$ đồng phân.\\
\textbf{Đáp số: 3.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho triolein lần lượt vào mỗi ống nghiệm chứa riêng biệt: $\ce{Na}$, $\ce{Cu(OH)2}$, $\ce{CH3OH}$, dung dịch $\ce{Br2}$, dung dịch $\ce{NaOH}$. Trong điều kiện thích hợp, số phản ứng xảy ra là bao nhiêu?
\loigiai{""" + (r"""
Triolein ($\ce{(C17H33COO)3C3H5}$) có 3 liên kết đôi $\ce{C=C}$ và 3 nhóm ester:\\
(1) Phản ứng cộng dung dịch $\ce{Br2}$ vào liên kết đôi $\ce{C=C}$.\\
(2) Phản ứng thủy phân trong dung dịch $\ce{NaOH}$.\\
Triolein không phản ứng với $\ce{Na}$, $\ce{Cu(OH)2}$, $\ce{CH3OH}$. $\Rightarrow$ Có \textbf{2} phản ứng.\\
\textbf{Đáp số: 2.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Một loại chất béo có chứa 80\% tristearin về khối lượng. Để sản xuất ba nghìn (3000) bánh xà phòng cần dùng tối thiểu $x\text{ kg}$ loại chất béo trên cho phản ứng với dung dịch $\ce{NaOH}$, đun nóng. Biết hiệu suất phản ứng đạt 90\%. Biết rằng trong mỗi bánh xà phòng có chứa $60\text{ gam}$ sodium stearate. Giá trị của $x$ là bao nhiêu? \textit{(Làm tròn kết quả đến một chữ số thập phân).}
\begin{flushright}
	\includegraphics[width=3.5cm]{images_crop_de01/soap_bar.png}
\end{flushright}
\loigiai{""" + (r"""
$m_{\text{sodium stearate}} = 3000 \times 60\text{ g} = 180\text{ kg}$.\\
Phương trình: $\ce{(C17H35COO)3C3H5 + 3NaOH -> 3C17H35COONa + C3H5(OH)3}$\\
Theo lí thuyết: $m_{\text{tristearin (LT)}} = 180 \times \dfrac{890}{3 \times 306} = 174{,}51\text{ kg}$.\\
Với $H = 90\%$ và hàm lượng tristearin $80\%$, khối lượng chất béo thực tế cần dùng là:
\[x = \dfrac{174{,}51}{0{,}90 \times 0{,}80} = \dfrac{180 \times 890 \times 100 \times 100}{918 \times 90 \times 80} \approx \mathbf{242{,}4\text{ kg}}.\]
\textbf{Đáp số: 242,4.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho $0{,}1\text{ mol}$ butanoic acid tác dụng với $0{,}1\text{ mol}$ methyl alcohol có mặt $\ce{H2SO4}$ đặc làm xúc tác. Tính khối lượng ester tạo thành theo gam, biết 67\% alcohol chuyển hoá thành ester. \textit{(Làm tròn kết quả đến hai chữ số thập phân).}
\loigiai{""" + (r"""
Phương trình: $\ce{CH3CH2CH2COOH + CH3OH <=> CH3CH2CH2COOCH3 + H2O}$.\\
Khối lượng mol của methyl butanoate: $M = 102\text{ g/mol}$.\\
Số mol ester: $n_{\text{ester}} = 0{,}1 \times 67\% = 0{,}067\text{ mol}$.\\
Khối lượng ester: $m_{\text{ester}} = 0{,}067 \times 102 = \mathbf{6{,}83\text{ gam}}$.\\
\textbf{Đáp số: 6,83.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Một loại dầu thực vật trong đó thành phần chất béo chứa hai gốc linoleate, một gốc oleate và thành phần phần trăm khối lượng chất béo trong dầu thực vật là 88\%. Tính chỉ số ester của dầu thực vật đó. (Biết chỉ số ester là số miligam $\ce{KOH}$ dùng để xà phòng hoá hết lượng triglyceride có trong 1 g chất béo).
\loigiai{""" + (r"""
Công thức cấu tạo: $\ce{(C17H31COO)2(C17H33COO)C3H5}$ ($M = 880\text{ g/mol}$).\\
Trong 1 gam dầu thực vật có $0{,}88\text{ gam}$ chất béo $\Rightarrow n_{\text{béo}} = \dfrac{0{,}88}{880} = 0{,}001\text{ mol}$.\\
Số mol $\ce{KOH}$ phản ứng: $n_{\ce{KOH}} = 3 \times n_{\text{béo}} = 0{,}003\text{ mol}$.\\
Khối lượng $\ce{KOH}$: $m_{\ce{KOH}} = 0{,}003 \times 56 = 0{,}168\text{ gam} = \mathbf{168\text{ mg}}$.\\
Vậy chỉ số ester là 168.\\
\textbf{Đáp số: 168.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Số miligam $\ce{KOH}$ dùng để xà phòng hoá hết lượng triglyceride có trong 1 g chất béo được gọi là chỉ số ester hoá của loại chất béo đó. Tính chỉ số ester của một loại chất béo chứa 65\% tristearin và 23\% triolein (còn lại là tạp chất không phản ứng). \textit{(Kết quả làm tròn đến phần nguyên).}
\loigiai{""" + (r"""
Trong 1 gam chất béo có $0{,}65\text{ gam}$ tristearin ($M = 890$) và $0{,}23\text{ gam}$ triolein ($M = 884$).\\
Số mol $\ce{KOH}$: $n_{\ce{KOH}} = 3 \times \left(\dfrac{0{,}65}{890} + \dfrac{0{,}23}{884}\right) \approx 0{,}0029715\text{ mol}$.\\
Khối lượng $\ce{KOH}$: $m_{\ce{KOH}} = 0{,}0029715 \times 56 \approx 0{,}1664\text{ gam} = 166{,}4\text{ mg} \approx \mathbf{166\text{ mg}}$.\\
Vậy chỉ số ester là 166.\\
\textbf{Đáp số: 166.}""" if is_teacher else "") + r"""}
\end{ex}
"""

    if is_teacher:
        tex += r"""
\vspace{0.4cm}
\noindent\hrulefill
\smallskip

\noindent
{\color{hfRed}\textbf{\large BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ SỐ 1}}

\smallskip
\noindent
\textbf{1. PHẦN I (Trắc nghiệm nhiều phương án):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1. A} & \textbf{2. B} & \textbf{3. D} & \textbf{4. B} & \textbf{5. C} & \textbf{6. C} & \textbf{7. B} & \textbf{8. B} & \textbf{9. D} \\ \hline
\textbf{10. A} & \textbf{11. D} & \textbf{12. D} & \textbf{13. C} & \textbf{14. C} & \textbf{15. B} & \textbf{16. D} & \textbf{17. B} & \textbf{18. B} \\ \hline
\end{tabular}
\end{center}

\noindent
\textbf{2. PHẦN II (Trắc nghiệm đúng sai):}
\begin{itemize}
	\item \textbf{Câu 1:} a) Sai \quad|\quad b) Đúng \quad|\quad c) Sai \quad|\quad d) Đúng
	\item \textbf{Câu 2:} a) Đúng \quad|\quad b) Đúng \quad|\quad c) Đúng \quad|\quad d) Đúng
	\item \textbf{Câu 3:} a) Đúng \quad|\quad b) Sai \quad|\quad c) Sai \quad|\quad d) Đúng
	\item \textbf{Câu 4:} a) Đúng \quad|\quad b) Đúng \quad|\quad c) Sai \quad|\quad d) Đúng
\end{itemize}

\noindent
\textbf{3. PHẦN III (Trả lời ngắn):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\ \hline
3 & 2 & 242{,}4 & 6{,}83 & 168 & 166 \\ \hline
\end{tabular}
\end{center}
"""

    tex += r"""
\begin{center}
\textbf{--------- HẾT ---------}
\end{center}

\end{document}
"""
    return tex

def main():
    hs_tex = make_tex(is_teacher=False)
    hs_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_12_De01_HocSinh.tex")
    with open(hs_path, "w", encoding="utf-8") as f:
        f.write(hs_tex)
    print("Saved:", hs_path)

    gv_tex = make_tex(is_teacher=True)
    gv_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_12_De01_GiaoVien.tex")
    with open(gv_path, "w", encoding="utf-8") as f:
        f.write(gv_tex)
    print("Saved:", gv_path)

    # Compile with pdflatex
    for name in ["De_Kiem_Tra_Hoa_12_De01_HocSinh", "De_Kiem_Tra_Hoa_12_De01_GiaoVien"]:
        print(f"Compiling {name}.tex...")
        cmd = f'pdflatex -interaction=nonstopmode "{name}.tex"'
        proc = subprocess.run(cmd, cwd=TEX_DIR, shell=True, capture_output=True, text=True)
        pdf_file = os.path.join(TEX_DIR, f"{name}.pdf")
        if os.path.exists(pdf_file):
            print(f"Successfully created: {pdf_file} ({os.path.getsize(pdf_file) / 1024:.1f} KB)")
        else:
            print(f"Compilation failed for {name}!")
            print(proc.stdout[-500:])

    # Clean up auxiliary files
    print("Cleaning auxiliary files...")
    for ext in [".aux", ".log", ".thm", ".out"]:
        for root, dirs, files in os.walk(TEX_DIR):
            for file in files:
                if file.endswith(ext) or file.startswith("ans-de01"):
                    try:
                        os.remove(os.path.join(root, file))
                    except:
                        pass
    print("LaTeX generation completed successfully!")

if __name__ == "__main__":
    main()
