import os
import subprocess
import pymupdf
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

ROOT_DIR = r"c:\Antigravity_Thanh"
SAN_PHAM_DIR = os.path.join(ROOT_DIR, "San_Pham")
IMAGES_DIR = os.path.join(ROOT_DIR, "images_crop")
os.makedirs(SAN_PHAM_DIR, exist_ok=True)
pdflatex = r"C:\Users\Admin\AppData\Local\Programs\MiKTeX\miktex\bin\x64\pdflatex.exe"

# ==============================================================================
# 1. FILE LATEX ĐỀ GIÁO VIÊN (KÈM LỜI GIẢI CHI TIẾT VÀ BẢNG ĐÁP ÁN)
# ==============================================================================
tex_giao_vien = r'''\documentclass[12pt,a4paper]{article}
\usepackage[utf8]{vietnam}
\usepackage{amsmath,amssymb}
\usepackage{geometry}
\geometry{top=1.5cm,bottom=1.8cm,left=1.4cm,right=1.4cm,headheight=15pt,headsep=10pt,footskip=22pt}
\usepackage{tikz}
\usepackage{xcolor}
\usepackage{enumerate}
\usepackage{chemfig}
\usepackage[version=4]{mhchem}
\usepackage{fancyhdr}
\usepackage{ex_test}

\makeatletter
\@ifundefined{c@bt}{\newcounter{bt}}{}
\makeatother

\renewcommand{\baselinestretch}{1.0}
\setlength{\parskip}{2pt}

% Hiển thị lời giải chi tiết
\printanswers

% Header & Footer theo file mẫu 3C2-26-27-DE của Thầy Trần Văn Thạnh
\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{\small\itshape\color{blue!50!green} Tài liệu lưu hành nội bộ \quad\vrule\quad THPT Hai Bà Trưng -- Huế}
\fancyhead[R]{\small\itshape\bfseries\color{blue!50!green} Thầy TRẦN VĂN THẠNH -- 0777.470.803}
\renewcommand{\headrulewidth}{0.4pt}
\renewcommand{\headrule}{\hbox to\headwidth{\color{blue!50!green}\leaders\hrule height \headrulewidth\hfill}}

\fancyfoot[L]{%
  \scriptsize
  \textbf{CS1:} 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế \quad\vrule\quad \textbf{CS2:} 24 Đặng Thái Thân, TP Huế\\
  \textit{Lớp Hóa Thầy Thạnh -- SĐT: 0777.470.803}%
}
\fancyfoot[R]{\small\textbf{Trang \thepage}}
\renewcommand{\footrulewidth}{0.4pt}
\renewcommand{\footrule}{\hbox to\headwidth{\color{blue!50!green}\leaders\hrule height \footrulewidth\hfill}}

\begin{document}

\noindent
\begin{minipage}[t]{0.45\textwidth}
	\centering
	\textbf{SỞ GIÁO DỤC \& ĐÀO TẠO TP HUẾ}\\
	\textbf{TRƯỜNG THPT HAI BÀ TRƯNG}\\
	\textbf{Mã đề thi: 209}
\end{minipage}
\hfill
\begin{minipage}[t]{0.52\textwidth}
	\centering
	\textbf{ĐỀ KIỂM TRA CUỐI KÌ II - NĂM HỌC 2024-2025}\\
	\textbf{MÔN HÓA HỌC LỚP 11}\\
	\textbf{\color{red!80!black}(HƯỚNG DẪN CHẤM VÀ LỜI GIẢI CHI TIẾT)}
\end{minipage}

\smallskip
\noindent
\textit{Cho NTK: \ce{H=1}; \ce{C=12}; \ce{O=16}; \ce{N=14}; \ce{Ag=108}; \ce{Na=23}.}

\medskip
\noindent
\textbf{A. PHẦN TRẮC NGHIỆM: [7,0 điểm]}

\begin{flushleft}
	\color{blue!50!green}\fbox{\fontfamily{qag}\bfseries\selectfont PHẦN 1. Câu trắc nghiệm 4 phương án}
\end{flushleft}
\setcounter{ex}{0}
\setcounter{bt}{0}

\textit{PHẦN I. [3,0 điểm]. Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án đúng (0,25 điểm/câu).}

\begin{ex}
Cho $x\text{ mol}$ phenol (\ce{C6H5OH}) tác dụng với \ce{Na} dư, thấy thoát ra $0{,}1\text{ mol}$ khí \ce{H2}. Giá trị của $x$ là
\choice
{$0{,}1$}
{$0{,}05$}
{$0{,}15$}
{\True $0{,}2$}
\loigiai{
Phương trình phản ứng:
$$\ce{C6H5OH + Na -> C6H5ONa + 1/2 H2 ^}$$
Theo phương trình: $n_{\ce{C6H5OH}} = 2 \cdot n_{\ce{H2}} = 2 \cdot 0{,}1 = 0{,}2\text{ mol} \Rightarrow x = 0{,}2$.\\
\textbf{Chọn D.}
}
\end{ex}

\begin{ex}
Cho các chất sau: methane, ethylene, acetylene, benzene, toluene và naphthalene. Số chất ở thể lỏng trong điều kiện thường là
\choice
{$1$}
{\True $2$}
{$3$}
{$4$}
\loigiai{
Ở điều kiện thường:
\begin{itemize}
  \item Thể khí: methane (\ce{CH4}), ethylene (\ce{C2H4}), acetylene (\ce{C2H2}).
  \item Thể lỏng: benzene (\ce{C6H6}), toluene (\ce{C7H8}) $\Rightarrow$ có $2$ chất ở thể lỏng.
  \item Thể rắn: naphthalene (\ce{C10H8}).
\end{itemize}
\textbf{Chọn B.}
}
\end{ex}

\begin{ex}
Aldehyde nào sau đây là đồng đẳng của \ce{CH3CHO}?
\choice
{\ce{CH2=CH-CHO}}
{\ce{C6H5-OH}}
{\ce{CH#C-CHO}}
{\True \ce{C2H5CHO}}
\loigiai{
Dãy đồng đẳng của acetaldehyde (\ce{CH3CHO}) là các aldehyde no, đơn chức, mạch hở có công thức chung \ce{C_nH_{2n}O} ($n \ge 1$). Chất đồng đẳng kế tiếp là propanal (\ce{C2H5CHO}).\\
\textbf{Chọn D.}
}
\end{ex}

\begin{ex}
Số hợp chất hữu cơ có công thức phân tử \ce{C3H8O} phản ứng được với \ce{Na} là
\choice
{$1$}
{\True $2$}
{$3$}
{$4$}
\loigiai{
Hợp chất có công thức \ce{C3H8O} phản ứng được với \ce{Na} phải là alcohol:
\begin{enumerate}
  \item Propan-1-ol: \ce{CH3-CH2-CH2-OH}
  \item Propan-2-ol: \ce{CH3-CH(OH)-CH3}
\end{enumerate}
Vậy có $2$ hợp chất alcohol phản ứng được với \ce{Na}.\\
\textbf{Chọn B.}
}
\end{ex}

\begin{ex}
Khi nhỏ từ từ dung dịch bromine vào ống nghiệm chứa dung dịch phenol, hiện tượng quan sát được trong ống nghiệm là
\choice
{không xảy ra hiện tượng gì}
{\True dung dịch brom mất màu và xuất hiện kết tủa trắng}
{xuất hiện kết tủa vàng}
{dung dịch trong suốt}
\loigiai{
Do ảnh hưởng của nhóm \ce{-OH}, mật độ electron ở vị trí ortho và para trên vòng benzen của phenol tăng lên, phenol phản ứng dễ dàng với nước bromine tạo kết tủa trắng $2,4,6$-tribromophenol và làm mất màu dung dịch brom:
$$\ce{C6H5OH + 3Br2 -> C6H2Br3OH v + 3HBr}$$
\textbf{Chọn B.}
}
\end{ex}

\begin{ex}
Tên gọi theo danh pháp thay thế của dẫn xuất halogen có công thức khung phân tử sau là:
\begin{center}
\begin{tikzpicture}[scale=0.45, baseline=-2pt]
  \draw[thick] (0,0) -- (0.8,0.6) -- (1.6,0) -- (2.4,0.6) -- (3.1,0.1) node[right=-2pt] {\ce{Br}};
\end{tikzpicture}
\end{center}
\choice
{\True $1$-bromobutane}
{$1$-bromopentane}
{$4$-bromobutane}
{$2$-bromobutane}
\loigiai{
Mạch chính có $4$ nguyên tử carbon, đánh số từ đầu gần nhóm thế bromine nhất:
$$\ce{\overset{4}{C}H3-\overset{3}{C}H2-\overset{2}{C}H2-\overset{1}{C}H2-Br}$$
Tên theo danh pháp thay thế là $1$-bromobutane.\\
\textbf{Chọn A.}
}
\end{ex}

\begin{ex}
\immini{
Catechin là một chất chống oxi hoá mạnh, ức chế hoạt động của các gốc tự do nên có khả năng phòng chống bệnh ung thư, nhồi máu cơ tim. Trong lá chè tươi, catechin chiếm khoảng $25 - 35\%$ tổng trọng lượng khô. Ngoài ra, catechin còn có trong táo, lê, nho,\dots\ Công thức cấu tạo của catechin cho như hình bên:

Phát biểu nào sau đây là \textbf{không} đúng?
}{
\setchemfig{atom sep=1.4em, bond offset=1pt}
\chemfig{*6(=(-[6]OH)-(*6(--(<[:-30]OH)-(<:[:30]*6(=-(-[:-30]OH)=(-[:30]OH)-=-))-O-))=-=(-[:150]HO)-)}
}
\choice
{\True Phân tử catechin có $5$ nhóm \ce{-OH} phenol}
{Catechin thuộc loại hợp chất thơm}
{Catechin phản ứng được với dung dịch \ce{NaOH}}
{Công thức phân tử của catechin là \ce{C15H14O6}}
\loigiai{
\begin{itemize}
  \item Phân tử catechin có $4$ nhóm \ce{-OH} liên kết trực tiếp vào vòng benzen (nhóm \ce{-OH} phenol) và $1$ nhóm \ce{-OH} liên kết với nguyên tử carbon no ở vị trí $\text{C}_3$ của vòng dị vòng dihydropyran (nhóm \ce{-OH} alcohol bậc hai). Do đó phát biểu ``có $5$ nhóm \ce{-OH} phenol'' là \textbf{sai}.
  \item Các phát biểu còn lại đều đúng: chứa vòng benzen nên là hợp chất thơm; các nhóm \ce{-OH} phenol phản ứng được với dung dịch \ce{NaOH}; công thức phân tử là \ce{C15H14O6}.
\end{itemize}
\textbf{Chọn A.}
}
\end{ex}

\begin{ex}
Trong các chất sau đây: \ce{CH3CH2OH}, \ce{CH3CHO}, \ce{CH3COOH}, \ce{CH3CH2CH2CH3}. Chất nào có nhiệt độ sôi cao nhất?
\choice
{\ce{CH3CHO}}
{\ce{CH3CH2OH}}
{\True \ce{CH3COOH}}
{\ce{CH3CH2CH2CH3}}
\loigiai{
Acetic acid (\ce{CH3COOH}) có liên kết hydrogen liên phân tử bền vững hơn alcohol (\ce{C2H5OH}) nhờ liên kết $\ce{O-H}$ phân cực mạnh hơn và tạo thành dạng dimer. Aldehyde và alkane không tạo được liên kết hydrogen liên phân tử nên có nhiệt độ sôi thấp hơn nhiều. Thứ tự nhiệt độ sôi:
$$\ce{CH3COOH} (118^\circ\text{C}) > \ce{CH3CH2OH} (78{,}3^\circ\text{C}) > \ce{CH3CHO} (20^\circ\text{C}) > \ce{CH3CH2CH2CH3} (-0{,}5^\circ\text{C})$$
\textbf{Chọn C.}
}
\end{ex}

\begin{ex}
Để phân biệt styrene và phenylacetylene có thể dùng chất nào sau đây?
\choice
{Khí oxygen dư}
{Nước bromine}
{Dung dịch \ce{KMnO4}}
{\True Dung dịch \ce{AgNO3} trong \ce{NH3}}
\loigiai{
Phenylacetylene (\ce{C6H5-C#CH}) là alkyne có liên kết ba đầu mạch (terminal alkyne), có nguyên tử $\ce{H}$ linh động nên phản ứng với dung dịch \ce{AgNO3} trong \ce{NH3} tạo kết tủa màu vàng nhạt bạc phenylacetylide (\ce{C6H5-C#CAg v}).\\
Styrene (\ce{C6H5-CH=CH2}) không có liên kết ba đầu mạch nên không phản ứng. Cả hai chất đều làm mất màu nước bromine và dung dịch \ce{KMnO4} nên không thể dùng để phân biệt.\\
\textbf{Chọn D.}
}
\end{ex}

\begin{ex}
Alkyne là những hydrocarbon mạch hở, chỉ chứa các liên kết đơn và một liên kết ba \ce{C#C} trong phân tử, có công thức chung là
\choice
{\True \ce{C_nH_{2n-2}} ($n \ge 2$)}
{\ce{C_nH_{2n+2}} ($n \ge 1$)}
{\ce{C_nH_{2n}} ($n \ge 2$)}
{\ce{C_nH_{2n-6}} ($n \ge 6$)}
\loigiai{
Alkyne mạch hở chứa $1$ liên kết ba $\ce{C#C}$ (tương ứng với độ bất bão hòa $k = 2$) có công thức phân tử chung là \ce{C_nH_{2n-2}} ($n \ge 2$).\\
\textbf{Chọn A.}
}
\end{ex}

\begin{ex}
Geraniol có trong tinh dầu hoa hồng (công thức cấu tạo thu gọn như hình bên dưới) được sử dụng phổ biến trong công nghiệp hương liệu, thực phẩm,\dots\ vì có mùi thơm đặc trưng.
\begin{center}
\begin{tikzpicture}[scale=0.5, font=\small, baseline=0]
  \node[left] at (-0.2,0.6) {\ce{H3C}};
  \draw[thick] (-0.2,0.6) -- (0.4,0.1);
  \draw[thick] (0.4,0.1) -- (0.4,-0.6) node[below] {\ce{CH3}};
  \draw[thick] (0.4,0.1) -- (1.2,0.6);
  \draw[thick] (0.45,0.25) -- (1.15,0.7);
  \draw[thick] (1.2,0.6) -- (1.9,0.1) -- (2.6,0.6) -- (3.3,0.1);
  \draw[thick] (3.3,0.1) -- (3.3,-0.6) node[below] {\ce{CH3}};
  \draw[thick] (3.3,0.1) -- (4.1,0.6);
  \draw[thick] (3.35,0.25) -- (4.05,0.7);
  \draw[thick] (4.1,0.6) -- (4.8,0.1) -- (5.4,0.45) node[right] {\ce{OH}};
\end{tikzpicture}
\end{center}
Geraniol thuộc loại hợp chất hữu cơ nào sau đây?
\choice
{Hydrocarbon}
{Carboxylic acid}
{\True Alcohol}
{Aldehyde}
\loigiai{
Trong công thức phân tử của geraniol có nhóm chức hydroxyl (\ce{-OH}) liên kết trực tiếp với nguyên tử carbon no (\ce{-CH2-OH}) $\Rightarrow$ Geraniol là một alcohol không no (chứa $2$ liên kết đôi $\ce{C=C}$).\\
\textbf{Chọn C.}
}
\end{ex}

\begin{ex}
Dẫn xuất halogen nào sau đây khi tác dụng với dung dịch \ce{NaOH} đun nóng không tạo thành alcohol?
\choice
{\True \ce{C6H5Cl}}
{\ce{CH3CH(Br)CH3}}
{\ce{C6H5CH2Br}}
{\ce{C2H5Cl}}
\loigiai{
Chlorobenzene (\ce{C6H5Cl}) có nguyên tử chlorine liên kết trực tiếp với vòng benzen. Do hiệu ứng liên hợp $p-\pi$, liên kết $\ce{C-Cl}$ rất bền vững, không bị thủy phân bởi dung dịch \ce{NaOH} loãng đun nóng thường (chỉ phản ứng ở điều kiện khắc nghiệt $300^\circ\text{C}$, $200\text{ atm}$ tạo sodium phenolate \ce{C6H5ONa}, không tạo alcohol).\\
Các dẫn xuất alkyl halogen và benzyl halogen khác đều dễ dàng thủy phân tạo alcohol tương ứng.\\
\textbf{Chọn A.}
}
\end{ex}

\bigskip
\begin{flushleft}
	{\color{blue!50!green}\fbox{\fontfamily{qag}\bfseries\selectfont PHẦN 2. Câu trắc nghiệm đúng sai}}
\end{flushleft}
\setcounter{ex}{0}
\setcounter{bt}{0}

\textit{PHẦN II. [2,0 điểm]. Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng (Đ) hoặc sai (S).}

\begin{ex}
Nhắc đến ethanol nhiều người thường nghĩ ngay đến đồ uống có cồn, tuy nhiên đây cũng là thành phần quan trọng được sử dụng trong y tế, là dung môi phổ biến cho nhiều ngành công nghiệp\dots
\choiceTF
{\True Ethanol còn gọi là ethyl alcohol}
{Ethanol là chất lỏng, dễ bay hơi, không mùi}
{\True Ethanol được tạo thành từ phản ứng thuỷ phân bromoethane bằng dung dịch \ce{NaOH} có đun nóng}
{\True Cho $4{,}6\text{ gam}$ ethanol tác dụng với \ce{Na} dư, thể tích khí \ce{H2} thu được (đkc) là $1{,}2395\text{ L}$}
\loigiai{
\begin{itemize}
  \item a) \textbf{Đúng.} Theo danh pháp gốc - chức, ethanol có tên là ethyl alcohol.
  \item b) \textbf{Sai.} Ethanol là chất lỏng không màu, dễ bay hơi, có mùi thơm nhẹ đặc trưng, vị cay, tan vô hạn trong nước.
  \item c) \textbf{Đúng.} Phản ứng thủy phân bromoethane:
  $$\ce{C2H5Br + NaOH ->[t^o] C2H5OH + NaBr}$$
  \item d) \textbf{Đúng.} Ta có: $n_{\ce{C2H5OH}} = \dfrac{4{,}6}{46} = 0{,}1\text{ mol}$.\\
  Phản ứng: $\ce{C2H5OH + Na -> C2H5ONa + 1/2 H2} \Rightarrow n_{\ce{H2}} = 0{,}05\text{ mol}$.\\
  Ở điều kiện chuẩn ($25^\circ\text{C}, 1\text{ bar}$): $V_{\ce{H2}} = 0{,}05 \times 24{,}79 = 1{,}2395\text{ L}$.
\end{itemize}
\textbf{Đáp án: a - Đ, b - S, c - Đ, d - Đ.}
}
\end{ex}

\begin{ex}
Hợp chất carbonyl đơn giản nhất là aldehyde và ketone đơn chức. Chúng có nhiều ứng dụng trong ngành công nghiệp hoá chất cũng như trong thiên nhiên.
\choiceTF
{Cho $0{,}44\text{ gam}$ ethanal vào dung dịch \ce{AgNO3} trong \ce{NH3} dư, đến khi phản ứng hoàn toàn thì thu được $10{,}8\text{ gam}$ \ce{Ag}}
{\True Trong tinh dầu thảo mộc có chứa những aldehyde, dùng dung dịch \ce{AgNO3} trong \ce{NH3} để nhận biết thành phần aldehyde trong tinh dầu}
{Tên thay thế của \ce{(CH3)2CHCH2CHO} là $2$-methylbutanal}
{Ethanal và propanone không thuộc loại hợp chất carbonyl}
\loigiai{
\begin{itemize}
  \item a) \textbf{Sai.} $n_{\ce{CH3CHO}} = \dfrac{0{,}44}{44} = 0{,}01\text{ mol}$.\\
  Phản ứng tráng bạc: $\ce{CH3CHO + 2[Ag(NH3)2]OH -> CH3COONH4 + 2Ag v + 3NH3 + H2O}$.\\
  $n_{\ce{Ag}} = 2 \cdot n_{\ce{CH3CHO}} = 0{,}02\text{ mol} \Rightarrow m_{\ce{Ag}} = 0{,}02 \times 108 = 2{,}16\text{ gam} \ne 10{,}8\text{ gam}$.
  \item b) \textbf{Đúng.} Phản ứng với thuốc thử Tollens (\ce{AgNO3} trong \ce{NH3}) là phản ứng đặc trưng để nhận biết nhóm chức aldehyde nhờ sự xuất hiện của lớp bạc sáng bóng bám trên thành ống nghiệm.
  \item c) \textbf{Sai.} Công thức cấu tạo: $\ce{\overset{4}{C}H3-\overset{3}{C}H(CH3)-\overset{2}{C}H2-\overset{1}{C}HO}$. Đánh số từ nhóm \ce{-CHO} là $\text{C}_1$, nhóm thế methyl ở vị trí số $3$, do đó tên thay thế đúng là $3$-methylbutanal (không phải $2$-methylbutanal).
  \item d) \textbf{Sai.} Hợp chất carbonyl là hợp chất hữu cơ trong phân tử có chứa nhóm carbonyl ($\ce{>C=O}$). Ethanal là aldehyde (\ce{CH3CHO}) và propanone là ketone (\ce{CH3COCH3}), cả hai đều thuộc loại hợp chất carbonyl.
\end{itemize}
\textbf{Đáp án: a - S, b - Đ, c - S, d - S.}
}
\end{ex}

\bigskip
\begin{flushleft}
	{\color{blue!50!green}\fbox{\fontfamily{qag}\bfseries\selectfont PHẦN 3. Câu trắc nghiệm trả lời ngắn}}
\end{flushleft}
\setcounter{ex}{0}
\setcounter{bt}{0}

\textit{PHẦN III. [2,0 điểm]. Câu trắc nghiệm yêu cầu trả lời ngắn. Thí sinh trả lời từ câu 1 đến câu 4 (0,5 điểm/câu).}

\begin{ex}
Ethene và acetylene là những hydrocarbon không no đơn giản nhất và có nhiều ứng dụng quan trọng. Một ứng dụng quan trọng của acetylene là làm nhiên liệu trong đèn xì oxygen - acetylene. Khi đèn hoạt động, hai khí này được trộn vào nhau để thực hiện phản ứng đốt cháy theo sơ đồ \ce{C2H2 + O2 ->[t^o] 2CO2 + H2O}. Đốt cháy hoàn toàn $V\text{ lít}$ (đkc) khí acetylene thu được $7{,}2\text{ gam}$ \ce{H2O}. Nếu cho tất cả sản phẩm cháy hấp thụ hết vào bình đựng nước vôi trong dư thì khối lượng bình tăng $m\text{ gam}$. Tính giá trị của $m$?
\par\shortans[oly]{42{,}4}
\loigiai{
Số mol nước thu được: $n_{\ce{H2O}} = \dfrac{7{,}2}{18} = 0{,}4\text{ mol}$.\\
Phương trình đốt cháy:
$$\ce{C2H2 + 5/2 O2 ->[t^o] 2 CO2 + H2O}$$
Theo phương trình: $n_{\ce{CO2}} = 2 \cdot n_{\ce{H2O}} = 2 \times 0{,}4 = 0{,}8\text{ mol}$.\\
Khi dẫn toàn bộ sản phẩm cháy (\ce{CO2} và \ce{H2O}) vào bình đựng nước vôi trong dư, cả hai chất đều bị giữ lại trong bình:\\
$\Delta m_{\text{bình tăng}} = m_{\ce{CO2}} + m_{\ce{H2O}} = 0{,}8 \times 44 + 7{,}2 = 35{,}2 + 7{,}2 = 42{,}4\text{ gam}$.\\
\textbf{Đáp số: 42,4}
}
\end{ex}

\begin{ex}
Phenol được dùng để sản xuất phẩm nhuộm, nhựa phenol-formaldehyde, thuốc nổ ($2{,}4{,}6$-trinitrophenol), chất diệt cỏ 2,4-D, chất diệt nấm mốc (các đồng phân của nitrophenol),\dots\ Do có tính diệt khuẩn nên phenol được dùng làm chất khử trùng, tẩy uế. Thuốc xịt chloraseptic chứa $1{,}4\%$ phenol được dùng làm thuốc chữa đau họng. Cho các phát biểu sau:
\begin{enumerate}[a)]
	\item Phenol tan vô hạn trong nước lạnh ở điều kiện thường.
	\item Nhiệt độ nóng chảy của phenol cao hơn ethanol.
	\item Phenol có khả năng tác dụng với dung dịch bromine tạo kết tủa trắng.
	\item Phenol dùng để sản xuất phẩm nhuộm, chất diệt nấm mốc, thuốc nổ TNT.
\end{enumerate}
Có bao nhiêu phát biểu đúng?
\par\shortans[oly]{2}
\loigiai{
Phân tích từng phát biểu:
\begin{itemize}
  \item a) \textbf{Sai.} Ở nhiệt độ thường, phenol ít tan trong nước lạnh (khoảng $8{,}3\text{ g}/100\text{ g nước}$), tan vô hạn khi đun nóng trên $66^\circ\text{C}$.
  \item b) \textbf{Đúng.} Ở điều kiện thường phenol là chất rắn có nhiệt độ nóng chảy $43^\circ\text{C}$, cao hơn ethanol là chất lỏng có nhiệt độ nóng chảy rất thấp ($-114{,}1^\circ\text{C}$).
  \item c) \textbf{Đúng.} Phenol tác dụng dễ dàng với dung dịch brom tạo kết tủa trắng $2,4,6$-tribromophenol.
  \item d) \textbf{Sai.} Thuốc nổ sản xuất từ phenol là acid picric ($2,4,6$-trinitrophenol), còn thuốc nổ TNT ($2,4,6$-trinitrotoluene) được sản xuất từ toluene (\ce{C7H8}).
\end{itemize}
Có $2$ phát biểu đúng (b và c).\\
\textbf{Đáp số: 2}
}
\end{ex}

\begin{ex}
Ngày nay, nhu cầu về đồ gỗ nội thất ngày càng nhiều song nguồn gỗ tự nhiên không còn dồi dào nên việc chuyển sang sử dụng gỗ công nghiệp đang là xu hướng của nhiều nước trên thế giới. Việc sử dụng gỗ công nghiệp góp phần bảo vệ rừng, bảo vệ môi trường. Quy trình sản xuất gỗ công nghiệp là nghiền các cây gỗ trồng ngắn ngày như keo, bạch đàn, cao su,\dots, sau đó sử dụng keo để kết dính và ép để tạo độ dày ván gỗ. Keo được sử dụng trong gỗ công nghiệp thường chứa dư lượng formaldehyde, là một hoá chất độc hại đối với sức khoẻ con người. Tại các nước phát triển như ở châu Âu và Mỹ, dư lượng formaldehyde được kiểm soát rất nghiêm ngặt. Châu Âu quy định tiêu chuẩn dư lượng formaldehyde trong gỗ công nghiệp là $120\ \mu\mathrm{g}\cdot\mathrm{m}^{-3}$. Cơ quan kiểm định lấy $300\text{ gam}$ gỗ trong một lô gỗ của một doanh nghiệp Việt Nam xuất khẩu sang châu Âu và kiểm tra bằng phương pháp sắc kí thấy chứa $0{,}03\ \mu\mathrm{g}$ formaldehyde. Biết khối lượng riêng của loại gỗ này là $800\ \mathrm{kg}\cdot\mathrm{m}^{-3}$. Hàm lượng formaldehyde có trong $800\text{ kg}$ (hay $1\ \mathrm{m}^3$) gỗ là bao nhiêu $\mu\mathrm{g}$?
\par\shortans[oly]{80}
\loigiai{
Đổi đơn vị: $800\text{ kg} = 800\,000\text{ gam}$.\\
Theo kết quả kiểm nghiệm: trong $300\text{ gam}$ gỗ có chứa $0{,}03\ \mu\text{g}$ formaldehyde.\\
Hàm lượng formaldehyde có trong $800\,000\text{ gam}$ ($1\text{ m}^3$) gỗ là:
$$m_{\text{formaldehyde}} = \frac{0{,}03\ \mu\text{g}}{300\text{ g}} \times 800\,000\text{ g} = 0{,}0001 \times 800\,000 = 80\ \mu\text{g}$$
(Giá trị này $80\ \mu\mathrm{g}\cdot\mathrm{m}^{-3} < 120\ \mu\mathrm{g}\cdot\mathrm{m}^{-3}$, đạt chuẩn chất lượng xuất khẩu sang châu Âu).\\
\textbf{Đáp số: 80}
}
\end{ex}

\begin{ex}
Cho các chất sau: \ce{H-CHO}; \ce{H-COOH}; \ce{CH3-COOH}; \ce{C2H5-OH}; \ce{CH2=CH-COOH}; \ce{C6H5-COOH}; \ce{HOOC-COOH}; \ce{C6H6}; \ce{HOOC-CH2-COOH}; \ce{C2H5COOH}.\\
Có bao nhiêu hợp chất carboxylic acid trong các chất trên?
\par\shortans[oly]{7}
\loigiai{
Hợp chất carboxylic acid là hợp chất hữu cơ có nhóm carboxyl (\ce{-COOH}) liên kết với nguyên tử carbon hoặc nguyên tử hydrogen. Liệt kê các chất:\\
\begin{enumerate}
  \item \ce{H-COOH} (methanoic acid / formic acid)
  \item \ce{CH3-COOH} (ethanoic acid / acetic acid)
  \item \ce{CH2=CH-COOH} (acrylic acid / prop-2-enoic acid)
  \item \ce{C6H5-COOH} (benzoic acid)
  \item \ce{HOOC-COOH} (oxalic acid / ethanedioic acid)
  \item \ce{HOOC-CH2-COOH} (malonic acid / propanedioic acid)
  \item \ce{C2H5COOH} (propanoic acid)
\end{enumerate}
Các chất còn lại: \ce{H-CHO} (aldehyde), \ce{C2H5-OH} (alcohol), \ce{C6H6} (hydrocarbon).\\
Vậy có tất cả $7$ hợp chất carboxylic acid.\\
\textbf{Đáp số: 7}
}
\end{ex}

\bigskip
\noindent
\textbf{B. PHẦN TỰ LUẬN: [3,0 điểm]}

\begin{flushleft}
	{\color{blue!50!green}\fbox{ \fontfamily{qag}\bfseries\selectfont Tự Luận:}}
\end{flushleft}
\setcounter{ex}{0}
\setcounter{bt}{0}

\begin{ex}
\textbf{[1,0 điểm]} Các nhà hoá học đã tìm ra một số dẫn xuất halogen không chứa chlorine như: \ce{CBr2F2}, \ce{CF3-CHF2}, \ce{CH2=CF-CF3}, \ce{CF3CH2CF2CH3},\dots\ đang được sử dụng trong công nghiệp nhiệt lạnh, vì sự phân huỷ các hợp chất này nhanh chóng sau khi phát tán vào không khí nên ảnh hưởng rất ít đến tầng ozone hay sự ấm lên toàn cầu thấp. Gọi tên theo danh pháp thay thế các hợp chất đó.
\loigiai{
Tên gọi theo danh pháp thay thế của các hợp chất:
\begin{itemize}
  \item \ce{CBr2F2}: \textbf{dibromodifluoromethane} \hfill \textit{[0,25 điểm]}
  \item \ce{CF3-CHF2}: \textbf{1,1,1,2,2-pentafluoroethane} \hfill \textit{[0,25 điểm]}
  \item \ce{CH2=CF-CF3}: \textbf{2,3,3,3-tetrafluoroprop-1-ene} (hoặc $2,3,3,3$-tetrafluoropropene) \hfill \textit{[0,25 điểm]}
  \item \ce{CF3CH2CF2CH3}: \textbf{1,1,1,3,3-pentafluorobutane} \hfill \textit{[0,25 điểm]}
\end{itemize}
}
\end{ex}

\begin{ex}
\textbf{[1,0 điểm]} Trong công nghiệp chế biến đường từ mía sẽ tạo ra sản phẩm phụ, gọi là rỉ đường hay rỉ mật, sử dụng rỉ đường để lên men tạo ra ethanol trong điều kiện thích hợp, hiệu suất cả quá trình là $80\%$. Tính khối lượng ethanol thu được từ $1{,}5\text{ tấn}$ rỉ đường mía theo 2 phương trình:
$$\ce{C12H22O11 + H2O -> C6H12O6 + C6H12O6}$$
$$\ce{C6H12O6 -> 2C2H5OH + 2CO2}$$
\loigiai{
\textbf{Lời giải:}
Sơ đồ phản ứng tổng hợp từ saccharose ra ethanol:
$$\ce{C12H22O11 -> 2 C6H12O6 -> 4 C2H5OH}$$
Khối lượng mol: $M_{\ce{C12H22O11}} = 342\text{ g/mol}$; $4 \cdot M_{\ce{C2H5OH}} = 4 \times 46 = 184\text{ g/mol}$. \hfill \textit{[0,25 điểm]}\\
Khối lượng ethanol theo lý thuyết thu được từ $1{,}5\text{ tấn}$ đường mía:
$$m_{\text{ethanol (lý thuyết)}} = 1{,}5 \times \frac{184}{342} \approx 0{,}8070\text{ tấn} = 807{,}0\text{ kg}$$ \hfill \textit{[0,50 điểm]}\\
Do hiệu suất của toàn bộ quá trình là $80\%$, khối lượng ethanol thực tế thu được là:
$$m_{\text{ethanol (thực tế)}} = 0{,}8070 \times 80\% = 0{,}6456\text{ tấn} = 645{,}6\text{ kg}$$ \hfill \textit{[0,25 điểm]}
}
\end{ex}

\begin{ex}
\textbf{[1,0 điểm]} Thị trường tiêu thụ phenol trên toàn thế giới khoảng $11{,}67\text{ triệu tấn}$ trong năm 2022. Phenol được sử dụng để sản xuất nhiều loại hoá chất như bisphenol A, nhựa phenol-formaldehyde, picric acid và các chất khác. Khoảng $90\%$ lượng phenol được sản xuất từ cumene. Khối lượng cumene cần dùng để sản xuất phenol cho năm 2022 là bao nhiêu?
\loigiai{
\textbf{Lời giải:}
Lượng phenol được sản xuất từ cumene trong năm 2022 là:
$$m_{\text{phenol}} = 11{,}67 \times 90\% = 10{,}503\text{ triệu tấn}$$ \hfill \textit{[0,25 điểm]}\\
Phương trình phản ứng điều chế phenol từ cumene (isopropylbenzene):
$$\ce{C6H5-CH(CH3)2 + O2 ->[H2SO4] C6H5OH + CH3COCH3}$$
Tỉ lệ mol giữa cumene và phenol là $1 : 1$.\\
Khối lượng mol của cumene (\ce{C9H12}) là $120\text{ g/mol}$; của phenol (\ce{C6H5OH}) là $94\text{ g/mol}$. \hfill \textit{[0,25 điểm]}\\
Khối lượng cumene cần dùng để sản xuất lượng phenol trên là:
$$m_{\text{cumene}} = 10{,}503 \times \frac{120}{94} \approx 13{,}408\text{ triệu tấn}$$ \hfill \textit{[0,50 điểm]}\\
\textbf{Đáp số:} Khoảng $13{,}41\text{ triệu tấn}$ cumene.
}
\end{ex}

\bigskip
\begin{center}
\textbf{\large BẢNG TỔNG HỢP ĐÁP ÁN TRẮC NGHIỆM ĐỀ 209}
\end{center}

\smallskip
\noindent
\textbf{PHẦN 1. Câu trắc nghiệm 4 phương án (0,25 điểm/câu):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} \\
\hline
\textbf{D} & \textbf{B} & \textbf{D} & \textbf{B} & \textbf{B} & \textbf{A} & \textbf{A} & \textbf{C} & \textbf{D} & \textbf{A} & \textbf{C} & \textbf{A} \\
\hline
\end{tabular}
\end{center}

\smallskip
\noindent
\textbf{PHẦN 2. Câu trắc nghiệm đúng/sai (1,0 điểm/câu):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Ý a)} & \textbf{Ý b)} & \textbf{Ý c)} & \textbf{Ý d)} \\
\hline
\textbf{Câu 1} & \textbf{Đ} & \textbf{S} & \textbf{Đ} & \textbf{Đ} \\
\hline
\textbf{Câu 2} & \textbf{S} & \textbf{Đ} & \textbf{S} & \textbf{S} \\
\hline
\end{tabular}
\end{center}

\smallskip
\noindent
\textbf{PHẦN 3. Câu trắc nghiệm trả lời ngắn (0,5 điểm/câu):}
\begin{center}
\begin{tabular}{|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} \\
\hline
\textbf{42,4} & \textbf{2} & \textbf{80} & \textbf{7} \\
\hline
\end{tabular}
\end{center}

\end{document}
'''

tex_gv_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_Hoa_11_Ma209_GiaoVien.tex")
with open(tex_gv_path, "w", encoding="utf-8") as f:
    f.write(tex_giao_vien)

res = subprocess.run([pdflatex, "-interaction=nonstopmode", "De_Kiem_Tra_Hoa_11_Ma209_GiaoVien.tex"], cwd=SAN_PHAM_DIR, capture_output=True, text=True)
print("Giao Vien PDF status:", res.returncode)
pdf_gv_path = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_Hoa_11_Ma209_GiaoVien.pdf")
if os.path.exists(pdf_gv_path):
    doc = pymupdf.open(pdf_gv_path)
    print("Giao Vien PDF Pages:", len(doc))
    doc.close()
