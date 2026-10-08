import os
import subprocess

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"

def build_tex(is_teacher=False):
    tex_filename = "De_Kiem_Tra_Hoa_10_Ma101_GiaoVien.tex" if is_teacher else "De_Kiem_Tra_Hoa_10_Ma101_HocSinh.tex"
    tex_path = os.path.join(OUT_DIR, tex_filename)

    header_status = r"\textbf{\color{red!80!black}(HƯỚNG DẪN CHẤM VÀ LỜI GIẢI CHI TIẾT)}" if is_teacher else r"\textit{Thời gian làm bài: 45 phút (không kể thời gian giao đề)}"

    teacher_macros = r"""
\renewcommand{\loigiai}[1]{%
  \par\smallskip\noindent{\color{blue!80!black}\textbf{Lời giải:}}\ #1\par\smallskip
}
\renewcommand{\True}{\bfseries\color{red!80!black}}
""" if is_teacher else r"""
\renewcommand{\loigiai}[1]{}
\renewcommand{\True}{}
"""

    ans_prefix = "ans-gv-101" if is_teacher else "ans-hs-101"

    student_space_23 = "" if is_teacher else r"""
\vspace{4.0cm}
\noindent\dotfill
"""

    student_space_24 = "" if is_teacher else r"""
\vspace{5.0cm}
\noindent\dotfill
"""

    student_space_25 = "" if is_teacher else r"""
\vspace{5.5cm}
\noindent\dotfill
"""

    template = r"""\documentclass[12pt,a4paper]{article}
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

\providecommand{\False}{}
__TEACHER_MACROS__

\begin{document}
\thispagestyle{firstpage}

\noindent
\begin{minipage}[t]{0.44\textwidth}
	\centering\fontsize{10pt}{12pt}\selectfont
	\textbf{TRƯỜNG ĐẠI HỌC KHOA HỌC}\\
	\textbf{TRƯỜNG THPT CHUYÊN KHOA HỌC HUẾ}\\
	\fbox{\textbf{MÃ ĐỀ: 101}}\\
	\textit{(Đề gồm 04 trang)}
\end{minipage}\hfill
\begin{minipage}[t]{0.54\textwidth}
	\centering\small
	\textbf{ĐỀ KIỂM TRA GIỮA HỌC KỲ I -- NĂM HỌC 2024--2025}\\
	\textbf{MÔN: HÓA HỌC -- LỚP 10}\\
	__HEADER_STATUS__
\end{minipage}

\smallskip
\noindent
Họ, tên thí sinh: \dotfill Lớp: \makebox[2.5cm]{\dotfill} Số báo danh: \makebox[2.5cm]{\dotfill}

\smallskip
\noindent
\textit{(Cho biết nguyên tử khối: $\ce{H}=1$; $\ce{C}=12$; $\ce{N}=14$; $\ce{O}=16$; $\ce{Na}=23$; $\ce{Mg}=24$; $\ce{Al}=27$; $\ce{P}=31$; $\ce{S}=32$; $\ce{Cl}=35{,}5$; $\ce{K}=39$; $\ce{Ca}=40$; $\ce{Fe}=56$; $\ce{Cu}=63{,}54$; thể tích mol khí ở đkc $25^\circ\text{C}, 1\text{ bar}$ là $24{,}79\text{ L}$).}

\smallskip
\noindent
\textbf{PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (5,0 điểm -- 20 câu)}

\Opensolutionfile{ans}[__ANS_PREFIX__]

\begin{ex}
Cho nguyên tử X có 11 electron ở lớp vỏ. Điện tích hạt nhân nguyên tử X là
\choice
{$-1{,}76 \cdot 10^{-18}\text{ C}$}
{$-1{,}826 \cdot 10^{-18}\text{ C}$}
{$+1{,}826 \cdot 10^{-18}\text{ C}$}
{\True $+1{,}76 \cdot 10^{-18}\text{ C}$}
\loigiai{
Trong nguyên tử trung hòa về điện: $Z = p = e = 11$. Hạt nhân mang điện tích dương:\\
$q_{\text{hạt nhân}} = +Z \cdot e_0 = +11 \times 1{,}602 \cdot 10^{-19}\text{ C} \approx +1{,}76 \cdot 10^{-18}\text{ C}$.\\
\textbf{Chọn D.}
}
\end{ex}

\begin{ex}
Cấu hình electron nguyên tử của oxygen là $1s^2 2s^2 2p^4$. Vị trí của oxygen trong bảng tuần hoàn là
\choice
{\True chu kì 2, nhóm VIA}
{chu kì 3, nhóm VIA}
{chu kì 2, nhóm IVA}
{chu kì 2, nhóm VIB}
\loigiai{
Oxygen có 2 lớp electron $\rightarrow$ chu kì 2. Lớp ngoài cùng (lớp 2) có $2 + 4 = 6$ electron, electron cuối cùng điền vào phân lớp 2p (nguyên tố p) $\rightarrow$ nhóm VIA.\\
\textbf{Chọn A.}
}
\end{ex}

\begin{ex}
Đối tượng nghiên cứu của hóa học là gì?
\choice
{Vật chất, năng lượng và sự vận động của chúng}
{Thế giới sinh vật gần gũi với đời sống hằng ngày của học sinh}
{\True Chất và sự biến đổi của chất}
{Nghệ thuật ngôn từ}
\loigiai{
Hóa học là ngành khoa học tự nhiên nghiên cứu về chất và sự biến đổi của chất cũng như ứng dụng của chúng.\\
\textbf{Chọn C.}
}
\end{ex}

\begin{ex}
Số hiệu nguyên tử cho biết thông tin nào sau đây?
\choice
{\True Số proton}
{Số neutron}
{Số khối}
{Nguyên tử khối}
\loigiai{
Số hiệu nguyên tử $Z$ cho biết số đơn vị điện tích hạt nhân, số proton (và số electron trong nguyên tử trung hòa).\\
\textbf{Chọn A.}
}
\end{ex}

\begin{ex}
Sự phân bố electron theo ô orbital nào dưới đây là đúng?
\choice
{\True \includegraphics[height=0.45cm]{images_crop_101/cau5_a.png}}
{\includegraphics[height=0.45cm]{images_crop_101/cau5_b.png}}
{\includegraphics[height=0.45cm]{images_crop_101/cau5_c.png}}
{\includegraphics[height=0.45cm]{images_crop_101/cau5_d.png}}
\loigiai{
Theo nguyên lý Pauli, trong 1 orbital chỉ chứa tối đa 2 electron và có spin ngược nhau ($\uparrow\downarrow$). Phương án B và D vi phạm nguyên lý Pauli. Theo quy tắc Hund, các electron phân bố sao cho số electron độc thân cực đại và có spin cùng chiều; phương án C vi phạm quy tắc Hund. Phương án A phân bố đúng.\\
\textbf{Chọn A.}
}
\end{ex}

\begin{ex}
Nguyên tử Chlorine ($Z = 17$) có số electron hóa trị là
\choice
{1}
{3}
{5}
{\True 7}
\loigiai{
Cấu hình electron của Chlorine ($Z = 17$): $1s^2 2s^2 2p^6 3s^2 3p^5$ (nhóm VIIA, nguyên tố p). Với nguyên tố nhóm A, số electron hóa trị bằng số electron lớp ngoài cùng $= 2 + 5 = 7$.\\
\textbf{Chọn D.}
}
\end{ex}

\begin{ex}
Mỗi orbital nguyên tử chứa tối đa
\choice
{1 electron}
{\True 2 electron}
{3 electron}
{4 electron}
\loigiai{
Theo nguyên lý Pauli, mỗi orbital nguyên tử chỉ có thể chứa tối đa 2 electron có chiều tự quay ngược nhau.\\
\textbf{Chọn B.}
}
\end{ex}

\begin{ex}
Trong tự nhiên, copper có 2 đồng vị là $^{63}\ce{Cu}$ và $^{65}\ce{Cu}$, trong đó đồng vị $^{65}\ce{Cu}$ chiếm 27\% nguyên tử. Phần trăm khối lượng của $^{63}\ce{Cu}$ trong $\ce{Cu2O}$ là: (cho biết đồng vị oxygen $^{16}_8\ce{O}$)
\choice
{73\%}
{63\%}
{32{,}14\%}
{\True 64{,}29\%}
\loigiai{
Phần trăm số nguyên tử của đồng vị $^{63}\ce{Cu}$: $100\% - 27\% = 73\%$.\\
Nguyên tử khối trung bình của copper:
\[\overline{A}_{\ce{Cu}} = \dfrac{63 \times 73 + 65 \times 27}{100} = 63{,}54.\]
Phân tử khối của $\ce{Cu2O}$: $M = 2 \times 63{,}54 + 16 = 143{,}08\text{ g/mol}$.\\
Trong 1 mol $\ce{Cu2O}$ có 2 mol Cu, khối lượng đồng vị $^{63}\ce{Cu}$ là $2 \times 73\% \times 63 = 91{,}98\text{ gam}$.\\
Phần trăm khối lượng của $^{63}\ce{Cu}$ trong $\ce{Cu2O}$ là:
\[\%m_{^{63}\ce{Cu}} = \dfrac{91{,}98}{143{,}08} \times 100\% \approx 64{,}29\%.\]
\textbf{Chọn D.}
}
\end{ex}

\begin{ex}
Cho các phát biểu sau:
\begin{enumerate}[(1)]
	\item Nguyên tử K có điện tích hạt nhân là $+3{,}0438 \cdot 10^{-18}\text{ C}$.
	\item Khối lượng hạt nhân được xem như là khối lượng nguyên tử.
	\item 1 amu bằng 1/12 khối lượng của nguyên tử carbon - 12.
	\item Đường kính hạt nhân gần bằng đường kính nguyên tử.
\end{enumerate}
Số phát biểu đúng là
\choice
{1}
{2}
{\True 3}
{4}
\loigiai{
(1) Đúng: Nguyên tử K có $Z = 19$, điện tích hạt nhân $q = +19 \times 1{,}602 \cdot 10^{-19}\text{ C} = +3{,}0438 \cdot 10^{-18}\text{ C}$.\\
(2) Đúng: Khối lượng electron rất nhỏ không đáng kể so với proton và neutron nên khối lượng hạt nhân được coi là khối lượng nguyên tử.\\
(3) Đúng: Định nghĩa 1 amu bằng 1/12 khối lượng nguyên tử carbon-12.\\
(4) Sai: Đường kính nguyên tử lớn hơn đường kính hạt nhân khoảng 10 000 lần.\\
Số phát biểu đúng là 3 gồm (1), (2), (3).\\
\textbf{Chọn C.}
}
\end{ex}

\begin{ex}
Sự phân bố electron vào các lớp và phân lớp căn cứ vào
\choice
{nguyên tử khối tăng dần}
{điện tích hạt nhân tăng dần}
{số khối tăng dần}
{\True mức năng lượng electron}
\loigiai{
Theo nguyên lý vững bền, các electron trong nguyên tử lần lượt chiếm các mức năng lượng từ thấp đến cao.\\
\textbf{Chọn D.}
}
\end{ex}

\begin{ex}
Số nguyên tố thuộc chu kì 3 của bảng tuần hoàn là
\choice
{2}
{\True 8}
{18}
{32}
\loigiai{
Chu kì 3 bắt đầu từ nguyên tố Sodium ($_{11}\ce{Na}$) đến Argon ($_{18}\ce{Ar}$), gồm tổng cộng 8 nguyên tố.\\
\textbf{Chọn B.}
}
\end{ex}

\begin{ex}
Các hạt cấu tạo nên hạt nhân nguyên tử là
\choice
{neutron và electron}
{electron, proton và neutron}
{electron và proton}
{\True proton và neutron}
\loigiai{
Hạt nhân nguyên tử được cấu tạo từ các hạt proton (mang điện tích dương) và neutron (không mang điện tích).\\
\textbf{Chọn D.}
}
\end{ex}

\begin{ex}
\noindent
\begin{minipage}{0.58\textwidth}
Năm 1897, nhà vật lý người Anh Joseph John Thomson thực hiện thí nghiệm phóng điện trong ống thủy tinh gần như chân không với hiệu điện thế lớn ($15\text{ kV}$). Mô hình thí nghiệm như hình vẽ bên. Nếu đặt một chong chóng nhẹ trên đường đi của tia âm cực thì chong chóng sẽ quay. Hiện tượng này chứng tỏ điều gì về tia âm cực?
\end{minipage}\hfill
\begin{minipage}{0.40\textwidth}
\centering
\includegraphics[width=\linewidth]{images_crop_101/cau13_clean.png}
\end{minipage}
\choice
{Tia âm cực mang điện tích âm}
{Tia âm cực là một loại ánh sáng trắng như ánh sáng mặt trời}
{Tia âm cực có phương truyền thẳng}
{\True Tia âm cực là chùm hạt vật chất chuyển động với vận tốc rất lớn}
\loigiai{
Chong chóng quay chứng tỏ tia âm cực có động lượng, tức là chùm hạt vật chất có khối lượng và chuyển động với vận tốc rất lớn va chạm cơ học làm quay chong chóng.\\
\textbf{Chọn D.}
}
\end{ex}

\begin{ex}
Trong tự nhiên oxygen có 3 đồng vị bền: $^{16}_8\ce{O}, ^{17}_8\ce{O}, ^{18}_8\ce{O}$, còn carbon có 2 đồng vị bền: $^{12}_6\ce{C}, ^{13}_6\ce{C}$. Số lượng phân tử $\ce{CO2}$ tạo ra từ các đồng vị trên là:
\choice
{8}
{10}
{\True 12}
{6}
\loigiai{
Phân tử $\ce{CO2}$ có dạng cấu trúc $\ce{O-C-O}$ đối xứng. Số cặp gồm 2 nguyên tử oxygen là $\dfrac{3 \times 4}{2} = 6$ cặp. Với 2 đồng vị của carbon, mỗi cặp oxygen kết hợp với 1 nguyên tử C tạo ra 1 phân tử $\ce{CO2}$. Tổng số phân tử $\ce{CO2}$ tạo thành là $6 \times 2 = 12$ phân tử.\\
\textbf{Chọn C.}
}
\end{ex}

\begin{ex}
Nguyên tử X có tổng số hạt cơ bản là 40. Trong đó tổng số hạt mang điện nhiều hơn số hạt không mang điện là 12 hạt. Nguyên tử X và số hiệu nguyên tử là
\choice
{$\ce{Na}$ ($Z = 11$)}
{$\ce{Mg}$ ($Z = 12$)}
{\True $\ce{Al}$ ($Z = 13$)}
{$\ce{Cl}$ ($Z = 17$)}
\loigiai{
Gọi $p, n, e$ là số proton, neutron, electron của X ($p = e$). Ta có:\\
$2p + n = 40$ và $2p - n = 12 \implies 4p = 52 \implies p = 13$ ($\ce{Al}$) và $n = 14$.\\
Số hiệu nguyên tử $Z = p = 13$, X là Aluminium ($\ce{Al}$).\\
\textbf{Chọn C.}
}
\end{ex}

\begin{ex}
Trong các hạt sau đây, hạt nào không mang điện tích?
\choice
{Electron}
{\True Neutron}
{Electron và proton}
{Proton}
\loigiai{
Proton mang điện tích $+1$, electron mang điện tích $-1$, neutron là hạt trung hòa về điện (không mang điện tích).\\
\textbf{Chọn B.}
}
\end{ex}

\begin{ex}
Nguyên tố nào sau đây thuộc nhóm A?
\choice
{\True $[\ce{Ne}]3s^2 3p^3$}
{$[\ce{Ar}]3d^1 4s^2$}
{$[\ce{Ar}]3d^7 4s^2$}
{$[\ce{Ar}]3d^5 4s^2$}
\loigiai{
Nguyên tố nhóm A là các nguyên tố s và p. Cấu hình $[\ce{Ne}]3s^2 3p^3$ có electron cuối cùng điền vào phân lớp 3p nên là nguyên tố p (thuộc nhóm VA). Các cấu hình còn lại là nguyên tố d thuộc nhóm B.\\
\textbf{Chọn A.}
}
\end{ex}

\begin{ex}
Khối lượng nguyên tử gần bằng khối lượng hạt nhân vì
\choice
{khối lượng electron gần bằng khối lượng hạt nhân}
{số lượng electron quá ít}
{\True tổng khối lượng electron không đáng kể}
{khối lượng nhân quá lớn}
\loigiai{
Vì khối lượng của mỗi electron rất nhỏ ($m_e \approx \dfrac{1}{1836} m_p$) nên tổng khối lượng các electron trong lớp vỏ là không đáng kể so với hạt nhân. Do đó khối lượng nguyên tử gần bằng khối lượng hạt nhân.\\
\textbf{Chọn C.}
}
\end{ex}

\begin{ex}
So sánh nào dưới đây về mức năng lượng của các phân lớp là không phù hợp?
\choice
{\True $3d < 4s$}
{$3p < 3d$}
{$1s < 2s$}
{$4s > 3s$}
\loigiai{
Theo thứ tự mức năng lượng tăng dần: $1s < 2s < 2p < 3s < 3p < 4s < 3d...$ Phân lớp 4s có mức năng lượng thấp hơn phân lớp 3d ($4s < 3d$). Do đó so sánh $3d < 4s$ là không phù hợp (sai).\\
\textbf{Chọn A.}
}
\end{ex}

\begin{ex}
Sulfur dạng kem bôi được sử dụng để điều trị mụn trứng cá. Nguyên tử sulfur có phân lớp electron ngoài cùng là $3p^4$. Phát biểu nào sau đây không đúng về nguyên tử sulfur?
\choice
{\True Lớp ngoài cùng có 4 electron}
{Nguyên tử có 16 electron}
{Thuộc chu kỳ 3 trong bảng tuần hoàn}
{Thuộc nhóm VIA trong bảng tuần hoàn}
\loigiai{
Cấu hình electron của sulfur: $1s^2 2s^2 2p^6 3s^2 3p^4$. Lớp ngoài cùng (lớp thứ 3) có $2 + 4 = 6$ electron (chứ không phải 4 electron). Phát biểu A sai.\\
\textbf{Chọn A.}
}
\end{ex}

\Closesolutionfile{ans}

\vspace{0.3cm}
\noindent
\textbf{PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2,0 điểm -- 2 câu)}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn ĐÚNG hoặc SAI.}

\begin{ex}
X là nguyên tố phổ biến thứ 4 trong vỏ trái đất, X có trong hemoglobin của máu làm nhiệm vụ vận chuyển oxygen, duy trì sự sống. Nguyên tử X có 26 proton trong hạt nhân.
\choiceTF[t]
{\False X có 26 neutron trong hạt nhân.}
{\True X có 26 electron ở vỏ nguyên tử.}
{\True X có điện tích hạt nhân là + 26.}
{\False Khối lượng nguyên tử X là 26 amu.}
\loigiai{
Nguyên tử X có 26 proton ($Z = 26$) $\rightarrow$ X là nguyên tố Iron (sắt, $\ce{Fe}$).
\begin{itemize}
	\item a) \textbf{Sai.} Đồng vị bền phổ biến nhất của sắt là $^{56}_{26}\ce{Fe}$ có số neutron $N = 56 - 26 = 30$ neutron. Hạt nhân Fe không có 26 neutron.
	\item b) \textbf{Đúng.} Trong nguyên tử trung hòa về điện, số electron ở vỏ bằng số proton: $e = p = 26$.
	\item c) \textbf{Đúng.} Hạt nhân có 26 proton nên điện tích hạt nhân là $+26$ (theo đơn vị điện tích nguyên tố $e_0$).
	\item d) \textbf{Sai.} Khối lượng nguyên tử Fe xấp xỉ số khối ($A \approx 56\text{ amu}$), 26 chỉ là số proton hoặc số đơn vị điện tích hạt nhân.
\end{itemize}
}
\end{ex}

\begin{ex}
Hình dưới mô tả orbital (a) và orbital (b) chứa electron trong nguyên tử sodium ($\ce{Na}$) ở trạng thái cơ bản. Mức năng lượng của orbital (a) cao hơn orbital (b).
\begin{center}
	\includegraphics[width=8.5cm]{images_crop_101/cau22_clean.png}
\end{center}
Cho các phát biểu sau:
\choiceTF[t]
{\False Electron trong các orbital (a) và (b) thuộc cùng lớp electron.}
{\False Số electron trong 1 orbital (b) gấp ba số electron trong orbital (a).}
{\False Electron trên orbital (a) nằm gần hạt nhân hơn electron trên orbital (b).}
{\True Orbital (a) và (b) khác nhau về định hướng trong không gian.}
\loigiai{
Sodium ($\ce{Na}$, $Z = 11$): $1s^2 2s^2 2p^6 3s^1$. Orbital (a) hình cầu có năng lượng cao hơn orbital (b) hình số 8 nổi ($2p$), do đó orbital (a) là orbital $3s$ (lớp $n = 3$), orbital (b) là orbital $2p$ (lớp $n = 2$).
\begin{itemize}
	\item a) \textbf{Sai.} Orbital (a) thuộc lớp 3, orbital (b) thuộc lớp 2; chúng không thuộc cùng lớp electron.
	\item b) \textbf{Sai.} Trong nguyên tử Na, 1 orbital $2p$ chứa 2 electron, orbital $3s$ chứa 1 electron ($3s^1$). Số electron trong 1 orbital (b) gấp 2 lần (chứ không phải gấp ba) orbital (a).
	\item c) \textbf{Sai.} Orbital $3s$ (a) ở lớp 3 nằm xa hạt nhân hơn orbital $2p$ (b) ở lớp 2.
	\item d) \textbf{Đúng.} Orbital s đối xứng cầu không có hướng ưu tiên, orbital p có hướng xác định trong không gian theo các trục Ox, Oy, Oz.
\end{itemize}
}
\end{ex}

\vspace{0.3cm}
\noindent
\textbf{PHẦN III: TỰ LUẬN (3,0 điểm -- 3 câu)}\\
\textit{Thí sinh trình bày chi tiết các bước giải và đáp số vào bài làm.}

\begin{ex}
\textbf{(1,0 điểm):} Trong tự nhiên, magnesium có 3 đồng vị bền là $^{24}\ce{Mg}$, $^{25}\ce{Mg}$ và $^{26}\ce{Mg}$. Phương pháp phổ khối lượng xác nhận đồng vị $^{26}\ce{Mg}$ chiếm tỉ lệ phần trăm số nguyên tử là 11\%. Biết rằng nguyên tử khối trung bình của $\ce{Mg}$ là 24,32. Tính \% số nguyên tử của đồng vị $^{24}\ce{Mg}$, đồng vị $^{25}\ce{Mg}$?
__STUDENT_SPACE_23__
\loigiai{
\begin{itemize}
	\item Gọi $x$ (\%) và $y$ (\%) lần lượt là phần trăm số nguyên tử của đồng vị $^{24}\ce{Mg}$ và $^{25}\ce{Mg}$ ($x, y > 0$).
	\item Tổng phần trăm số nguyên tử của 3 đồng vị bằng $100\%$:
	\[x + y + 11 = 100 \implies x + y = 89 \quad (1)\]
	\item Theo công thức tính nguyên tử khối trung bình của Mg:
	\[\overline{A} = \dfrac{24x + 25y + 26 \times 11}{100} = 24{,}32\]
	\[\implies 24x + 25y + 286 = 2432 \implies 24x + 25y = 2146 \quad (2)\]
	\item Giải hệ phương trình (1) và (2):
	\[\begin{cases} x + y = 89 \\ 24x + 25y = 2146 \end{cases} \implies \begin{cases} x = 79 \\ y = 10 \end{cases}\]
	\item \textbf{Kết luận:} Phần trăm số nguyên tử của đồng vị $^{24}\ce{Mg}$ là \textbf{79\%}, đồng vị $^{25}\ce{Mg}$ là \textbf{10\%}.
\end{itemize}
}
\end{ex}

\begin{ex}
\textbf{(1,0 điểm):} Cho hai nguyên tử X ($Z = 15$) và Y ($Z = 26$).
\begin{enumerate}[a)]
	\item Viết cấu hình electron và xác định vị trí của mỗi nguyên tử trong bảng tuần hoàn hóa học.
	\item Cho biết X và Y là nguyên tố kim loại, phi kim hay khí hiếm? Giải thích.
\end{enumerate}
__STUDENT_SPACE_24__
\loigiai{
\begin{enumerate}[a)]
	\item Cấu hình electron và vị trí trong bảng tuần hoàn:
	\begin{itemize}
		\item Nguyên tử X ($Z = 15$):
		\begin{itemize}
			\item Cấu hình electron: $1s^2 2s^2 2p^6 3s^2 3p^3$ (hoặc $[\ce{Ne}]3s^2 3p^3$).
			\item Ô số: 15 (vì $Z = 15$).
			\item Chu kì: 3 (vì có 3 lớp electron).
			\item Nhóm: VA (vì là nguyên tố p và có 5 electron lớp ngoài cùng).
		\end{itemize}
		\item Nguyên tử Y ($Z = 26$):
		\begin{itemize}
			\item Cấu hình electron: $1s^2 2s^2 2p^6 3s^2 3p^6 3d^6 4s^2$ (hoặc $[\ce{Ar}]3d^6 4s^2$).
			\item Ô số: 26 (vì $Z = 26$).
			\item Chu kì: 4 (vì có 4 lớp electron).
			\item Nhóm: VIIIB (vì là nguyên tố d, tổng số electron phân lớp $3d$ và $4s$ là $6 + 2 = 8$).
		\end{itemize}
	\end{itemize}
	\item Xác định tính chất nguyên tố:
	\begin{itemize}
		\item X là \textbf{phi kim} vì có 5 electron ở lớp ngoài cùng ($3s^2 3p^3$).
		\item Y là \textbf{kim loại} (kim loại chuyển tiếp) vì có 2 electron ở lớp ngoài cùng ($4s^2$) và phân lớp 3d chưa bão hòa.
	\end{itemize}
\end{enumerate}
}
\end{ex}

\begin{ex}
\textbf{(1,0 điểm):} Hòa tan hoàn toàn 20 gam hỗn hợp 2 nguyên tố A và B thuộc nhóm IIA, ở 2 chu kì liên tiếp nhau vào dung dịch $\ce{HCl}$ dư thu được $17{,}353\text{ lít}$ khí (đktc). Xác định tên 2 nguyên tố A, B và thành phần \% về khối lượng của mỗi nguyên tố trong hỗn hợp.
__STUDENT_SPACE_25__
\loigiai{
\begin{itemize}
	\item Đổi số mol khí $\ce{H2}$ thoát ra theo điều kiện chuẩn (đkc: $25^\circ\text{C}, 1\text{ bar}$):
	\[n_{\ce{H2}} = \dfrac{17{,}353}{24{,}79} = 0{,}7\text{ mol}.\]
	\textit{(Ghi chú: Đề bài ghi kí hiệu đktc theo thói quen cũ nhưng dùng số liệu chuẩn $V_{\text{mol}} = 24{,}79\text{ L/mol}$ theo chương trình GDPT 2018).}
	\item Gọi $\overline{M}$ là nguyên tử khối trung bình của hai kim loại A và B.
	\[\overline{M} + 2\ce{HCl ->} \overline{M}\ce{Cl2} + \ce{H2 ^}\]
	Theo phương trình: $n_{\text{hh kim loại}} = n_{\ce{H2}} = 0{,}7\text{ mol}$.
	\item Khối lượng mol trung bình của hai kim loại:
	\[\overline{M} = \dfrac{m_{\text{hh}}}{n_{\text{hh}}} = \dfrac{20}{0{,}7} \approx 28{,}57\text{ g/mol}.\]
	\item Vì A và B là 2 kim loại thuộc nhóm IIA ở 2 chu kì liên tiếp:
	\[M_{\text{Be}} (9) < M_{\ce{Mg}} (24) < \overline{M} = 28{,}57 < M_{\ce{Ca}} (40) < M_{\ce{Sr}} (88) < M_{\ce{Ba}} (137)\]
	Suy ra: $M_{\text{A}} = 24$ ($\ce{Mg}$, chu kì 3) và $M_{\text{B}} = 40$ ($\ce{Ca}$, chu kì 4).\\
	Vậy hai kim loại là \textbf{Magnesium ($\ce{Mg}$)} và \textbf{Calcium ($\ce{Ca}$)}.
	\item Tính thành phần \% khối lượng của mỗi kim loại:
	Gọi $a, b$ lần lượt là số mol của $\ce{Mg}$ và $\ce{Ca}$ trong 20 gam hỗn hợp ($a, b > 0$):
	\[\begin{cases} a + b = 0{,}7 \\ 24a + 40b = 20 \end{cases} \implies \begin{cases} a = 0{,}5\text{ mol} \\ b = 0{,}2\text{ mol} \end{cases}\]
	\item Khối lượng của từng kim loại:
	\[m_{\ce{Mg}} = 0{,}5 \times 24 = 12\text{ gam} \implies \%m_{\ce{Mg}} = \dfrac{12}{20} \times 100\% = \mathbf{60\%}.\]
	\[m_{\ce{Ca}} = 0{,}2 \times 40 = 8\text{ gam} \implies \%m_{\ce{Ca}} = \dfrac{8}{20} \times 100\% = \mathbf{40\%}.\]
\end{itemize}
}
\end{ex}

\vspace{0.5cm}
\noindent
\textit{\small Ghi chú: Cán bộ coi thi không được giải thích gì thêm.}

\begin{center}
\textbf{--------- HẾT ---------}
\end{center}

\vspace{0.3cm}
\noindent
\begin{minipage}[t]{0.48\textwidth}
	\centering
	\textbf{DUYỆT}\\
	\textit{(Ký và ghi rõ họ tên)}\\[1.5cm]
	\textbf{Nguyễn Hồ Ngọc Thư}
\end{minipage}\hfill
\begin{minipage}[t]{0.48\textwidth}
	\centering
	\textbf{CÁN BỘ RA ĐỀ}\\
	\textit{(Ký và ghi rõ họ tên)}\\[1.5cm]
	\textbf{Tôn Nữ Mỹ Phương}
\end{minipage}

\end{document}
"""
    final_text = template.replace("__TEACHER_MACROS__", teacher_macros)
    final_text = final_text.replace("__HEADER_STATUS__", header_status)
    final_text = final_text.replace("__ANS_PREFIX__", ans_prefix)
    final_text = final_text.replace("__STUDENT_SPACE_23__", student_space_23)
    final_text = final_text.replace("__STUDENT_SPACE_24__", student_space_24)
    final_text = final_text.replace("__STUDENT_SPACE_25__", student_space_25)

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(final_text)
    print(f"Saved: {tex_path}")

def compile_pdf(is_teacher=False):
    tex_filename = "De_Kiem_Tra_Hoa_10_Ma101_GiaoVien.tex" if is_teacher else "De_Kiem_Tra_Hoa_10_Ma101_HocSinh.tex"
    base_name = os.path.splitext(tex_filename)[0]

    cmd = ["pdflatex", "-interaction=nonstopmode", tex_filename]
    print(f"Compiling {tex_filename} (run 1)...")
    res1 = subprocess.run(cmd, cwd=OUT_DIR, capture_output=True, text=True)
    print(f"Compiling {tex_filename} (run 2)...")
    res2 = subprocess.run(cmd, cwd=OUT_DIR, capture_output=True, text=True)

    if res2.returncode == 0:
        print(f"Successfully compiled {base_name}.pdf")
    else:
        print(f"Error compiling {tex_filename}: returncode = {res2.returncode}")
        # print error lines from log
        log_file = os.path.join(OUT_DIR, f"{base_name}.log")
        if os.path.exists(log_file):
            with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    if line.startswith("!"):
                        print("  ", line.strip())

    # Cleanup temp files
    extensions = [".aux", ".log", ".thm", ".out"]
    for ext in extensions:
        temp_f = os.path.join(OUT_DIR, base_name + ext)
        if os.path.exists(temp_f):
            os.remove(temp_f)
    for ans_f in ["ans-gv-101.tex", "ans-hs-101.tex", "ans-101.tex"]:
        temp_ans = os.path.join(OUT_DIR, ans_f)
        if os.path.exists(temp_ans):
            os.remove(temp_ans)

if __name__ == "__main__":
    build_tex(is_teacher=False)
    build_tex(is_teacher=True)
    compile_pdf(is_teacher=False)
    compile_pdf(is_teacher=True)
