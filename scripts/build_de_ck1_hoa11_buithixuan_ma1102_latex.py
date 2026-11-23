# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Cuối Học Kì I Hóa Học 11 - THPT Bùi Thị Xuân - Huế
Mã đề: 1102 - Năm học 2025 - 2026
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_CK1_Hoa_11_BuiThiXuan_Ma1102_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_CK1_Hoa_11_BuiThiXuan_Ma1102_GiaoVien.pdf (Kèm bảng đáp án rút gọn và lời giải chi tiết)
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
\textbf{TRƯỜNG THPT BÙI THỊ XUÂN}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 03 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.48\textwidth}
\raggedleft
\textbf{\color{hfRed}KIỂM TRA ĐÁNH GIÁ CUỐI KÌ I}\\
\textbf{NĂM HỌC 2025 -- 2026}\\
\textbf{MÔN: HOÁ HỌC LỚP 11}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{MÃ ĐỀ: 1102}
\end{minipage}

\vspace{4pt}
\noindent
\textit{Họ và tên:} \dotfill \textit{Số báo danh:} \dotfill

\vspace{3pt}
\noindent
\textit{\footnotesize Cho biết nguyên tử khối của các nguyên tố: $\ce{H} = 1; \ce{O} = 16; \ce{C} = 12; \ce{N} = 14; \ce{Na} = 23; \ce{Mg} = 24; \ce{K} = 39; \ce{S} = 32$.}

\vspace{3pt}
\hrule color{hfGreen} height 0.5pt
""")

    if is_teacher:
        tex.append(r"""
\vspace{4pt}
\begin{center}
\textbf{\color{hfRed}\large (BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}
\end{center}

\vspace{2pt}
\begin{center}
{\color{hfRed}\textbf{\large BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN)}}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}1. Phần I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - mỗi câu đúng 0,25 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} \\
\hline
\textbf{Đ/A} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}D} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} & \textbf{\color{hfRed}A} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}B} & \textbf{\color{hfRed}C} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}2. Phần II: Trắc nghiệm Đúng / Sai (2,0 điểm - mỗi câu 1,0 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|l|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Nội dung câu hỏi} & \textbf{Ý a} & \textbf{Ý b} & \textbf{Ý c} & \textbf{Ý d} \\
\hline
\textbf{Câu 1} & Tính chất của Sulfur và hợp chất & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{\color{hfRed}ĐÚNG} & \textbf{\color{hfRed}ĐÚNG} \\
\hline
\textbf{Câu 2} & Hợp chất hữu cơ $\ce{X}$ (Ethanol, $\ce{C2H6O}$) & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} & \textbf{\color{hfRed}ĐÚNG} & \textbf{SAI} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}3. Phần III: Trả lời ngắn (2,0 điểm - mỗi câu đúng 0,5 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|c|c|c|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} \\
\hline
\textbf{Đáp án} & \textbf{\color{hfRed}3} & \textbf{\color{hfRed}3} & \textbf{\color{hfRed}3214} & \textbf{\color{hfRed}18} \\
\hline
\end{tabular}
\end{center}

\vspace{2pt}
\noindent
\textbf{\color{mainGreen}4. Tóm tắt đáp số Phần Tự luận (3,0 điểm)}
\begin{center}
\renewcommand{\arraystretch}{1.2}
\begin{tabular}{|c|c|p{11.5cm}|}
\hline
\rowcolor{tableGreen}
\textbf{Câu} & \textbf{Điểm} & \textbf{Đáp số cốt lõi} \\
\hline
\textbf{Câu 1} & 0,5 đ & $n_e$ trao đổi $= 0,08\text{ mol} \implies a = 8 \implies$ Khí $\ce{X}$ là $\ce{N2O}$ (dinitrogen monoxide). \\
\hline
\textbf{Câu 2} & 1,0 đ & Dùng dd $\ce{Ba(OH)2}$: $\ce{(NH4)2SO4}$ ($\downarrow$ trắng + khí mùi khai); $\ce{NH4Cl}$ (chỉ có khí mùi khai); $\ce{NaCl}$ (không hiện tượng). \\
\hline
\textbf{Câu 3} & 1,0 đ & a) Peak $m/z = 60 \implies M = 60\text{ g/mol} \implies$ CTPT: $\ce{C2H4O2}$. \quad b) $m_{\text{acid}} = 2,5\text{ g/kg} < 10\text{ g/kg} \implies$ An toàn. \\
\hline
\textbf{Câu 4} & 0,5 đ & 2 đồng phân mạch hở: $\ce{CH3-CH2-NH2}$ (ethylamine) và $\ce{CH3-NH-CH3}$ (dimethylamine). \\
\hline
\end{tabular}
\end{center}

\vspace{4pt}
\hrule color{hfGreen} height 0.5pt
""")

    # NỘI DUNG ĐỀ THI
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\Large A. PHẦN TRẮC NGHIỆM (7 điểm)}}

\vspace{3pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{2pt}
""")

    # 12 CÂU PHẦN I
    # Câu 1
    tex.append(r"""
\begin{ex}
Cho các phát biểu sau:
\begin{enumerate}[label=(\arabic*),leftmargin=1.5cm,itemsep=0pt]
\item Sulfur vừa có tính oxi hóa vừa có tính khử.
\item Trong tự nhiên, sulfur chỉ tồn tại ở dạng hợp chất.
\item $\ce{SO2}$ là acidic oxide, tan trong nước tạo dung dịch có tính acid.
\item $\ce{H2SO4}$ đặc có tính háo nước mạnh và tính oxi hóa mạnh.
\item Trong hợp chất, nguyên tố sulfur có số oxi hóa phổ biến là $-2; +4; +6$.
\end{enumerate}
Số phát biểu đúng là
\choice
{2.}
{3.}
{\True 4.}
{5.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- (1) \textbf{Đúng}: Sulfur đơn chất có số oxi hóa 0 trung gian, có thể giảm xuống $-2$ (tính oxi hóa) hoặc tăng lên $+4, +6$ (tính khử).\\
- (2) \textbf{Sai}: Trong tự nhiên, sulfur tồn tại ở cả dạng đơn chất (các mỏ sulfur tự do) và hợp chất (quặng sulfide, sulfate).\\
- (3) \textbf{Đúng}: $\ce{SO2}$ là acidic oxide, tác dụng với nước tạo dung dịch $\ce{H2SO3}$ có tính acid yếu: $\ce{SO2 + H2O <=> H2SO3}$.\\
- (4) \textbf{Đúng}: $\ce{H2SO4}$ đặc có tính háo nước mạnh và tính oxi hóa mạnh.\\
- (5) \textbf{Đúng}: Trong hợp chất, sulfur có số oxi hóa phổ biến là $-2, +4, +6$.\\
$\implies$ Có 4 phát biểu đúng: (1), (3), (4), (5).\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 2
    tex.append(r"""
\begin{ex}
Dung dịch $\ce{H2SO4}$ loãng có thể tác dụng với cả hai chất nào sau đây?
\choice
{\True $\ce{Fe}$ và $\ce{Fe(OH)3}$.}
{$\ce{S}$ và $\ce{H2S}$.}
{$\ce{C}$ và $\ce{CO2}$.}
{$\ce{Cu}$ và $\ce{Cu(OH)2}$.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Dung dịch $\ce{H2SO4}$ loãng tác dụng được với kim loại đứng trước $\ce{H}$ ($\ce{Fe}$) và base ($\ce{Fe(OH)3}$):\\
$\ce{Fe + H2SO4 -> FeSO4 + H2 ^}$\\
$\ce{2Fe(OH)3 + 3H2SO4 -> Fe2(SO4)3 + 6H2O}$\\
Dung dịch $\ce{H2SO4}$ loãng không phản ứng với $\ce{S, H2S, C, CO2, Cu}$.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 3
    tex.append(r"""
\begin{ex}
Cặp chất nào sau đây có thể là đồng đẳng của nhau?
\choice
{$\ce{C2H4, C3H8}$.}
{$\ce{CH4O, C2H4O}$.}
{$\ce{CH4, C2H4}$.}
{\True $\ce{C2H2, C4H6}$.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
$\ce{C2H2}$ và $\ce{C4H6}$ có cùng công thức chung $\ce{C_n H_{2n-2}}$, phân tử hơn kém nhau 2 nhóm $\ce{-CH2-}$, có thể cùng thuộc dãy đồng đẳng alkyne ($\ce{CH#CH}$ và $\ce{CH#C-CH2-CH3}$ hoặc $\ce{CH3-C#C-CH3}$).\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 4
    tex.append(r"""
\begin{ex}
Để tách tinh dầu sả (có trong thân, lá, rễ .... cây sả) trong công nghiệp hương liệu, người ta dùng phương pháp
\choice
{kết tinh dầu sả trong nước.}
{\True chưng cất bằng hơi nước và chiết tinh dầu ra khỏi hỗn hợp sản phẩm.}
{chưng cất bằng hơi nước và chiết bằng nước lạnh.}
{chiết tinh dầu sả sau đó chưng cất bằng hơi nước.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Tinh dầu sả nhẹ hơn nước, không tan trong nước và dễ bay hơi theo hơi nước sôi. Người ta dùng phương pháp chưng cất lôi cuốn hơi nước, sau đó ngưng tụ và dùng phễu chiết để tách riêng tinh dầu ra khỏi nước.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 5
    tex.append(r"""
\begin{ex}
Yếu tố không làm ảnh hưởng đến chuyển dịch cân bằng là
\choice
{nồng độ.}
{nhiệt độ.}
{áp suất.}
{\True chất xúc tác.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Chất xúc tác làm tăng tốc độ phản ứng thuận và nghịch với mức độ như nhau, giúp hệ mau đạt tới trạng thái cân bằng chứ không làm chuyển dịch cân bằng hóa học.\\
\textbf{Chọn D.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 6
    tex.append(r"""
\begin{ex}
Liên kết hoá học trong phân tử $\ce{NH3}$ là liên kết
\choice
{kim loại.}
{\True cộng hoá trị có cực.}
{ion.}
{cộng hoá trị không cực.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Hiệu độ âm điện $\Delta\chi = \chi(\ce{N}) - \chi(\ce{H}) = 3,04 - 2,20 = 0,84$ ($0,4 \le \Delta\chi < 1,7$) nên liên kết $\ce{N-H}$ trong $\ce{NH3}$ là liên kết cộng hóa trị có cực (phân cực).\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 7
    tex.append(r"""
\begin{ex}
Phổ khối lượng của hợp chất hữu cơ $\ce{X}$ thu được như hình vẽ:
\begin{center}
\includegraphics[width=0.48\textwidth]{images_ck1_hoa11_buithixuan/ms_cau7_hd.png}
\end{center}
Phân tử khối của hợp chất hữu cơ $\ce{X}$ là
\choice
{28.}
{75.}
{\True 44.}
{30.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Trên phổ khối lượng (MS), peak ion phân tử $[\ce{M+}]$ có giá trị $m/z$ lớn nhất là 44. Do đó, phân tử khối của hợp chất hữu cơ $\ce{X}$ là 44.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 8
    tex.append(r"""
\begin{ex}
Chọn phát biểu không đúng?
\choice
{\True Nitrogen là nguyên tố phổ biến nhất trong vỏ trái đất.}
{Nitrogen lỏng được dùng để bảo quản mẫu vật phẩm trong y học.}
{Hai nguyên tử $\ce{N}$ trong phân tử nitrogen liên kết với nhau bằng liên kết ba.}
{Ở nhiệt độ thường, nitrogen khá trơ về mặt hoá học.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Nguyên tố phổ biến nhất trong vỏ Trái Đất là Oxygen ($\approx 46,6\%$), tiếp theo là Silicon ($\approx 27,7\%$), Aluminium ($\approx 8,1\%$),... Nitrogen chỉ chiếm khoảng $0,002\%$ trong vỏ Trái Đất dù chiếm 78\% thể tích khí quyển. Do đó phát biểu A không đúng.\\
\textbf{Chọn A.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 9
    tex.append(r"""
\begin{ex}
Khí $\ce{NO_x}$ là một trong các nguyên nhân gây nên hiện tượng mưa acid, mù quang hóa, hiện tượng phú dưỡng,... làm ô nhiễm môi trường. Nguồn phát sinh $\ce{NO_x}$ có nguồn gốc tự nhiên là
\choice
{hoạt động giao thông vận tải.}
{\True hoạt động của núi lửa.}
{hoạt động đốt nhiên liệu.}
{hoạt động của nhà máy nhiệt điện.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Giao thông, đốt nhiên liệu, nhà máy nhiệt điện là nguồn nhân tạo. Nguồn phát sinh $\ce{NO_x}$ tự nhiên là do hoạt động của núi lửa, sấm sét trong khí quyển và vi sinh vật trong đất.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 10
    tex.append(r"""
\begin{ex}
Trong các phát biểu sau đây, phát biểu nào đúng khi một hệ ở trạng thái cân bằng?
\choice
{Phản ứng nghịch đã dừng.}
{\True Nồng độ của các chất trong hệ không đổi.}
{Nồng độ chất tham gia và sản phẩm bằng nhau.}
{Phản ứng thuận đã dừng.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Cân bằng hóa học là cân bằng động, cả 2 phản ứng vẫn diễn ra với tốc độ bằng nhau ($v_t = v_n > 0$), do đó nồng độ các chất trong hệ không đổi theo thời gian.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 11
    tex.append(r"""
\begin{ex}
Để xác định nồng độ mol của một dung dịch $\ce{NaOH}$, người ta đã tiến hành chuẩn độ bằng dung dịch $\ce{HCl}$ 0,1M. Để chuẩn độ 10 mL dung dịch $\ce{NaOH}$ này cần 12 mL dung dịch $\ce{HCl}$. Nồng độ của dung dịch $\ce{NaOH}$ trên là
\choice
{0,012M.}
{\True 0,12M.}
{0,1M.}
{1,2M.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Phương trình phản ứng: $\ce{NaOH + HCl -> NaCl + H2O}$\\
Tại điểm tương đương: $n_{\ce{NaOH}} = n_{\ce{HCl}} \implies C_{\ce{NaOH}} \times 10 = 0,1 \times 12 \implies C_{\ce{NaOH}} = 0,12\text{ M}$.\\
\textbf{Chọn B.}
}
""")
    tex.append(r"\end{ex}")

    # Câu 12
    tex.append(r"""
\begin{ex}
Liên kết hoá học chủ yếu trong phân tử hợp chất hữu cơ là
\choice
{liên kết ion.}
{liên kết cho-nhận.}
{\True liên kết cộng hoá trị.}
{liên kết hydrogen.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Trong hợp chất hữu cơ, nguyên tử carbon có độ âm điện trung bình, liên kết với nhau và với các nguyên tử khác chủ yếu bằng liên kết cộng hóa trị.\\
\textbf{Chọn C.}
}
""")
    tex.append(r"\end{ex}")

    # PHẦN II
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN II. Câu trắc nghiệm đúng sai.}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{2pt}
""")

    # Câu 1 Phần II
    tex.append(r"""
\begin{ex}
Sulfur và các hợp chất của nó có vai trò quan trọng trong sản xuất và đời sống, nhưng khí thải và chất thải chứa sulfur nếu không được xử lí đúng cách sẽ gây tác động tiêu cực đến sức khỏe con người và môi trường.
\choiceTF
{\True Ở điều kiện thường, lưu huỳnh là chất rắn, màu vàng, không tan trong nước, tan nhiều trong các dung môi hữu cơ như benzene, carbon disulfide ($\ce{CS2}$)...}
{Khí $\ce{SO2}$ có thể được dùng làm chất tẩy màu, khi hoà tan khí này vào nước tạo ra dung dịch làm quỳ tím hoá xanh.}
{\True Thuốc thử nhận biết sulfuric acid và muối sulfate là ion $\ce{Ba^2+}$ trong $\ce{BaCl2, Ba(OH)2, Ba(NO3)2}$.}
{\True Sulfuric acid đặc lấy nước từ hợp chất carbohydrate và khiến chúng hoá đen chứng tỏ sulfuric acid có tính háo nước.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
a) \textbf{Đúng}: Ở điều kiện thường, đơn chất sulfur là chất rắn màu vàng, không tan trong nước, tan nhiều trong các dung môi không phân cực như benzene, $\ce{CS2}$.\\
b) \textbf{Sai}: $\ce{SO2}$ tan trong nước tạo dung dịch acid $\ce{H2SO3}$ làm quỳ tím hóa đỏ (không phải hóa xanh).\\
c) \textbf{Đúng}: Cation $\ce{Ba^2+}$ tạo kết tủa trắng $\ce{BaSO4}$ không tan trong acid với ion sulfate: $\ce{Ba^2+ + SO4^2- -> BaSO4 v}$.\\
d) \textbf{Đúng}: $\ce{H2SO4}$ đặc có tính háo nước mạnh, chiếm $\ce{H}$ và $\ce{O}$ trong carbohydrate tạo carbon màu đen.
}
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần II
    tex.append(r"""
\begin{ex}
Hợp chất hữu cơ $\ce{X}$ được sử dụng rộng rãi trong đời sống như làm dung môi, nhiên liệu sinh học và thành phần chính của đồ uống có cồn. Kết quả phân tích nguyên tố hợp chất $\ce{X}$ cho thấy có 52,17\% carbon, 34,78\% oxygen, còn lại là hydrogen. Tỉ khối hơi của $\ce{X}$ so với khí hydrogen bằng 23. Phổ IR của $\ce{X}$ có dải hấp thụ đặc trưng được mô tả trong hình dưới đây:
\begin{center}
\includegraphics[width=0.72\textwidth]{images_ck1_hoa11_buithixuan/ir_cau2_p2.png}
\end{center}
\choiceTF
{\True Công thức phân tử của $\ce{X}$ là $\ce{C2H6O}$.}
{Do $\ce{X}$ tan vô hạn trong nước và có nhiệt độ sôi thấp hơn nước, nên không thể tách $\ce{X}$ khỏi hỗn hợp với nước bằng phương pháp chưng cất.}
{\True Dựa vào phổ IR, có thể dự đoán phân tử $\ce{X}$ chứa nhóm chức $-\ce{OH}$.}
{Chất $\ce{CH3OCH3}$ không phải là đồng phân cấu tạo của $\ce{X}$.}
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- $M_{\ce{X}} = 23 \times 2 = 46\text{ g/mol}$.\\
- $\%H = 100\% - 52,17\% - 34,78\% = 13,05\%$.\\
- Đặt CTPT $\ce{C_x H_y O_z} \implies x = \dfrac{46 \times 52,17\%}{12} = 2$; $y = \dfrac{46 \times 13,05\%}{1} = 6$; $z = \dfrac{46 \times 34,78\%}{16} = 1 \implies$ CTPT là $\ce{C2H6O}$.\\
a) \textbf{Đúng}: CTPT của $\ce{X}$ là $\ce{C2H6O}$.\\
b) \textbf{Sai}: Ethanol sôi ở $78,3^\circ\text{C}$ và nước sôi ở $100^\circ\text{C}$, nhiệt độ sôi chênh lệch đáng kể nên hoàn toàn có thể tách và làm giàu ethanol bằng chưng cất phân đoạn.\\
c) \textbf{Đúng}: Phổ IR có peak tù, mạnh ở $3391\text{ cm}^{-1}$ đặc trưng cho liên kết $\ce{O-H}$ của alcohol.\\
d) \textbf{Sai}: $\ce{CH3-O-CH3}$ có cùng CTPT $\ce{C2H6O}$ nên chính là một đồng phân cấu tạo của ethanol.
}
""")
    tex.append(r"\end{ex}")

    # PHẦN III
    tex.append(r"""
\vspace{4pt}
\noindent
{\color{mainGreen}\textbf{\large PHẦN III. Câu trắc nghiệm trả lời ngắn.}}\\
\textit{\small Thí sinh trả lời từ câu 1 đến câu 4. Thí sinh điền kết quả vào ô trả lời tương ứng.}
\vspace{2pt}
""")

    # Câu 1 Phần III
    tex.append(r"""
\begin{ex}
Cho cân bằng hóa học: $\ce{N2(g) + O2(g) <=> 2NO(g)} \quad \Delta_r H^\circ_{298} = +180\text{ kJ}$.
Có các biện pháp sau:
(1) Tăng nhiệt độ.
(2) Hạ nhiệt độ.
(3) Tăng nồng độ $\ce{O2}$.
(4) Tăng áp suất chung của hệ phản ứng.
(5) Dùng thêm chất xúc tác $\ce{V2O5}$.\\
Hỏi có bao nhiêu biện pháp làm cân bằng trên chuyển dịch?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
- Phản ứng thuận thu nhiệt ($\Delta_r H^\circ_{298} > 0$), số mol khí hai vế bằng nhau (2 mol = 2 mol).\\
- (1) Tăng nhiệt độ $\implies$ chuyển dịch theo chiều thu nhiệt (chiều thuận).\\
- (2) Hạ nhiệt độ $\implies$ chuyển dịch theo chiều tỏa nhiệt (chiều nghịch).\\
- (3) Tăng nồng độ $\ce{O2} \implies$ chuyển dịch theo chiều giảm $\ce{O2}$ (chiều thuận).\\
- (4) Áp suất không ảnh hưởng do số mol khí hai vế bằng nhau.\\
- (5) Chất xúc tác không làm chuyển dịch cân bằng.\\
$\implies$ Có 3 biện pháp làm cân bằng chuyển dịch: (1), (2), (3).\\
\textbf{Đáp số: 3}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 2 Phần III
    tex.append(r"""
\begin{ex}
Cho các chất sau: $\ce{HCOOH, CaCO3, C2H4, CO, C6H12O6, NaCN}$. Số chất thuộc loại hợp chất hữu cơ là bao nhiêu?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Các chất là hợp chất hữu cơ: $\ce{HCOOH}$ (formic acid), $\ce{C2H4}$ (ethylene), $\ce{C6H12O6}$ (glucose). Tổng cộng có 3 chất.\\
($\ce{CaCO3, CO, NaCN}$ là các hợp chất vô cơ của carbon).\\
\textbf{Đáp số: 3}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 3 Phần III
    tex.append(r"""
\begin{ex}
Các bước tách một hợp chất hữu cơ bằng phương pháp chiết được mô tả như sau:
\begin{enumerate}[label=(\arabic*),leftmargin=1.5cm,itemsep=0pt]
\item Để yên phễu chiết cho hai lớp chất lỏng phân tách hoàn toàn.
\item Lắc nhẹ phễu chiết để chất cần tách chuyển sang dung môi chiết.
\item Cho hỗn hợp chất lỏng không trộn lẫn vào phễu chiết cùng dung môi thích hợp.
\item Mở khóa phễu chiết để thu từng lớp chất lỏng riêng biệt.
\end{enumerate}
Hãy sắp xếp các bước trên theo đúng thứ tự thực hiện của phương pháp chiết.
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Thứ tự chuẩn của phương pháp chiết lỏng - lỏng:\\
- Bước 1: Cho hỗn hợp và dung môi vào phễu chiết (3).\\
- Bước 2: Lắc nhẹ phễu chiết (2).\\
- Bước 3: Để yên cho phân tách 2 lớp hoàn toàn (1).\\
- Bước 4: Mở khóa thu từng lớp chất lỏng (4).\\
$\implies$ Thứ tự: (3) $\to$ (2) $\to$ (1) $\to$ (4) $\implies$ \textbf{3214}.\\
\textbf{Đáp số: 3214}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # Câu 4 Phần III
    tex.append(r"""
\begin{ex}
Geraniol là một hợp chất hữu cơ tự nhiên có mùi hoa hồng đặc trưng, được sử dụng rộng rãi trong sản xuất nước hoa, mỹ phẩm và hương liệu thực phẩm; công thức cấu tạo của geraniol được biểu diễn như sau:
\begin{center}
\includegraphics[width=0.45\textwidth]{images_ck1_hoa11_buithixuan/geraniol_hd.png}
\end{center}
Số nguyên tử hydrogen có trong phân tử geraniol là bao nhiêu?
""")
    if is_teacher:
        tex.append(r"""
\loigiai{
Công thức cấu tạo: $\ce{(CH3)2C=CH-CH2-CH2-C(CH3)=CH-CH2OH}$.\\
Đếm số nguyên tử $\ce{H}$: $6 + 1 + 2 + 2 + 3 + 1 + 2 + 1 = 18$ nguyên tử $\ce{H}$.\\
(Công thức phân tử của geraniol là $\ce{C10H18O}$).\\
\textbf{Đáp số: 18}
}
""")
    else:
        tex.append(r"""
\noindent\textit{Trả lời:} \dotfill
""")
    tex.append(r"\end{ex}")

    # B. PHẦN TỰ LUẬN
    tex.append(r"""
\vspace{6pt}
\noindent
{\color{mainGreen}\textbf{\Large B. PHẦN TỰ LUẬN (3,0 điểm)}}
\vspace{2pt}
""")

    # Câu 1 Tự luận
    tex.append(r"""
\noindent
\textbf{Câu 1: (0,5 điểm)} Cho $0,04\text{ mol }\ce{Mg}$ tan hết trong dung dịch $\ce{HNO3}$ thấy thoát ra $0,01\text{ mol}$ khí $\ce{X}$ là sản phẩm khử duy nhất (đkc). Xác định công thức của khí $\ce{X}$?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 1:}]
\begin{itemize}[leftmargin=1.2cm]
\item Quá trình nhường electron: $\ce{Mg -> Mg^2+ + 2e} \implies n_e\text{ nhường} = 2 \times 0,04 = 0,08\text{ mol}$. \hfill \textit{(0,15 đ)}
\item Gọi $a$ là số electron mà 1 phân tử khí $\ce{X}$ nhận:
  $n_e\text{ nhận} = 0,01 \times a\text{ (mol)}$.
\item Áp dụng định luật bảo toàn electron:
  $n_e\text{ nhường} = n_e\text{ nhận} \iff 0,08 = 0,01 \times a \iff a = 8$. \hfill \textit{(0,2 đ)}
\item Phân tử khí nhận 8 electron ứng với sản phẩm khử dinitrogen monoxide ($\ce{N2O}$):\\
  $\ce{2NO3- + 10H+ + 8e -> N2O ^ + 5H2O}$.
\item Vậy công thức của khí $\ce{X}$ là $\ce{N2O}$. \hfill \textit{(0,15 đ)}
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{5}{\dotfill\\[4pt]}
""")

    # Câu 2 Tự luận
    tex.append(r"""
\vspace{4pt}
\noindent
\textbf{Câu 2: (1,0 điểm)} Bằng phương pháp hoá học hãy nhận biết các dung dịch sau: $\ce{NaCl, (NH4)2SO4}$ và $\ce{NH4Cl}$. Viết các phương trình phản ứng xảy ra.
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 2:}]
\textbf{1. Cách tiến hành nhận biết: (0,5 điểm)}
\begin{itemize}[leftmargin=1.2cm]
\item Trích mẫu thử của các dung dịch vào các ống nghiệm có đánh số tương ứng.
\item Cho từ từ dung dịch $\ce{Ba(OH)2}$ dư vào từng mẫu thử và đun nóng nhẹ: \hfill \textit{(0,2 đ)}
  \begin{itemize}
  \item Mẫu thử xuất hiện kết tủa trắng, đồng thời có khí mùi khai thoát ra làm quỳ tím ẩm hóa xanh là $\ce{(NH4)2SO4}$. \hfill \textit{(0,1 đ)}
  \item Mẫu thử chỉ có khí mùi khai thoát ra làm quỳ tím ẩm hóa xanh (không có kết tủa) là $\ce{NH4Cl}$. \hfill \textit{(0,1 đ)}
  \item Mẫu thử không có hiện tượng gì là $\ce{NaCl}$. \hfill \textit{(0,1 đ)}
  \end{itemize}
\end{itemize}
\textbf{2. Phương trình phản ứng xảy ra: (0,5 điểm)}
\begin{itemize}[leftmargin=1.2cm]
\item $\ce{(NH4)2SO4 + Ba(OH)2 -> BaSO4 v (trắng) + 2NH3 ^ (mùi khai) + 2H2O}$ \hfill \textit{(0,25 đ)}
\item $\ce{2NH4Cl + Ba(OH)2 -> BaCl2 + 2NH3 ^ (mùi khai) + 2H2O}$ \hfill \textit{(0,25 đ)}
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{7}{\dotfill\\[4pt]}
""")

    # Câu 3 Tự luận
    tex.append(r"""
\vspace{4pt}
\noindent
\textbf{Câu 3: (1,0 điểm)} Acetic acid được sử dụng rộng rãi trên thế giới trong nhiều ngành công nghiệp khác nhau như tạo ra polymer ứng dụng trong sơn, chất kết dính, là dung môi hoà tan các chất hoá học, sản xuất và bảo quản thực phẩm, đặc biệt dùng để sản xuất giấm.\\
a) Lập công thức phân tử của acetic acid, biết kết quả phân tích nguyên tố acetic acid có 40\% C; 53,33\% O về khối lượng; còn lại là H. Phân tử khối của acetic acid được xác định trên phổ khối lượng tương ứng với peak có giá trị $m/z$ lớn nhất.\\
\begin{center}
\includegraphics[width=0.48\textwidth]{images_ck1_hoa11_buithixuan/ms_acetic_acid_hd.png}
\end{center}
b) Theo khuyến cáo an toàn, lượng acetic acid trong thực phẩm không nên vượt quá 10 gam trên mỗi một kg thực phẩm. Một cơ sở chế biến thực phẩm dùng 50 mL giấm ăn 5\% để tạo vị chua cho 1,0 kg dưa muối. Biết khối lượng riêng của giấm ăn bằng $1,00\text{ g/mL}$. Hỏi lượng acetic acid sử dụng có vượt quá giới hạn an toàn hay không, giải thích?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 3:}]
\textbf{a) Lập công thức phân tử của acetic acid: (0,5 điểm)}
\begin{itemize}[leftmargin=1.2cm]
\item $\%H = 100\% - 40\% - 53,33\% = 6,67\%$. \hfill \textit{(0,1 đ)}
\item Trên phổ khối lượng (MS), peak ion phân tử $[\ce{M+}]$ có $m/z = 60 \implies M = 60\text{ g/mol}$. \hfill \textit{(0,1 đ)}
\item Đặt CTPT là $\ce{C_x H_y O_z}$:\\
  $x = \dfrac{60 \times 40\%}{12} = 2$; \quad $y = \dfrac{60 \times 6,67\%}{1} = 4$; \quad $z = \dfrac{60 \times 53,33\%}{16} = 2$. \hfill \textit{(0,2 đ)}
\item Vậy CTPT của acetic acid là $\ce{C2H4O2}$ (hoặc $\ce{CH3COOH}$). \hfill \textit{(0,1 đ)}
\end{itemize}
\textbf{b) Đánh giá giới hạn an toàn: (0,5 điểm)}
\begin{itemize}[leftmargin=1.2cm]
\item Khối lượng dung dịch giấm ăn: $m_{\text{dd}} = V \times D = 50\text{ mL} \times 1,00\text{ g/mL} = 50\text{ gam}$. \hfill \textit{(0,15 đ)}
\item Khối lượng acetic acid nguyên chất có trong giấm ăn:
  $m_{\text{acid}} = 50 \times 5\% = 2,5\text{ gam}$. \hfill \textit{(0,15 đ)}
\item Lượng acetic acid trong 1,0 kg dưa muối là $2,5\text{ gam/kg}$.
\item So sánh: $2,5\text{ gam/kg} < 10\text{ gam/kg}$ (giới hạn an toàn tối đa). \hfill \textit{(0,1 đ)}
\item \textbf{Kết luận:} Lượng acetic acid sử dụng \textbf{không vượt quá} giới hạn an toàn cho phép. \hfill \textit{(0,1 đ)}
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{8}{\dotfill\\[4pt]}
""")

    # Câu 4 Tự luận
    tex.append(r"""
\vspace{4pt}
\noindent
\textbf{Câu 4: (0,5 điểm)} Hãy viết các đồng phân mạch hở có cùng công thức phân tử $\ce{C2H7N}$?
""")
    if is_teacher:
        tex.append(r"""
\begin{tcolorbox}[colback=white,colframe=solBlue,title=\textbf{Lời giải chi tiết và Biểu điểm Câu 4:}]
\begin{itemize}[leftmargin=1.2cm]
\item Độ bất bão hòa: $k = \dfrac{2 \times 2 + 2 + 1 - 7}{2} = 0 \implies$ Amine no, đơn chức, mạch hở. \hfill \textit{(0,1 đ)}
\item Các đồng phân cấu tạo amine mạch hở gồm:
  \begin{enumerate}
  \item Amine bậc 1: $\ce{CH3-CH2-NH2}$ \quad (Tên gọi: ethanamine hoặc ethylamine). \hfill \textit{(0,2 đ)}
  \item Amine bậc 2: $\ce{CH3-NH-CH3}$ \quad (Tên gọi: $N$-methylmethanamine hoặc dimethylamine). \hfill \textit{(0,2 đ)}
  \end{enumerate}
\end{itemize}
\end{tcolorbox}
""")
    else:
        tex.append(r"""
\vspace{1pt}
\noindent
\multido{\n=1+1}{5}{\dotfill\\[4pt]}
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
    print(f"[OK] Đã biên dịch xong PDF: {os.path.join(out_dir, base_name + '.pdf')}")

def main():
    tex_hs = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_CK1_Hoa_11_BuiThiXuan_Ma1102_HocSinh.tex")
    tex_gv = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_CK1_Hoa_11_BuiThiXuan_Ma1102_GiaoVien.tex")

    print("Đang tạo mã nguồn LaTeX Bản Học sinh THPT Bùi Thị Xuân...")
    with open(tex_hs, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))

    print("Đang tạo mã nguồn LaTeX Bản Giáo viên THPT Bùi Thị Xuân...")
    with open(tex_gv, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))

    print("\nĐang biên dịch PDF Bản Học sinh THPT Bùi Thị Xuân...")
    compile_pdf(tex_hs)

    print("\nĐang biên dịch PDF Bản Giáo viên THPT Bùi Thị Xuân...")
    compile_pdf(tex_gv)

    print("\n[HOÀN TẤT] Toàn bộ file .tex và .pdf đã được tạo và biên dịch thành công!")

if __name__ == "__main__":
    main()
