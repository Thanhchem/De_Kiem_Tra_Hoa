# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file LaTeX và biên dịch ra PDF cho ĐỀ SỐ 12 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De12_HocSinh.tex / .pdf
- De_Kiem_Tra_Hoa_11_De12_GiaoVien.tex / .pdf
Chương 4: HYDROCARBON
Tuân thủ Form chuẩn GEMINI.md, sử dụng gói ex_test.sty và mhchem.
"""

import os
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import subprocess
import shutil

TEX_DIR = r"c:\Antigravity_Thanh\San_Pham"

if not os.path.exists(os.path.join(TEX_DIR, "ex_test.sty")):
    shutil.copyfile(r"c:\Antigravity_Thanh\ex_test.sty", os.path.join(TEX_DIR, "ex_test.sty"))

def make_tex(is_teacher=False):
    ver_tag = "gv" if is_teacher else "hs"
    ans_file = f"ans-hoa11-de12{ver_tag}"
    
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
	\textbf{ĐỀ CHÍNH THỨC}\\
	\textit{(Đề thi có 04 trang)}
\end{minipage}\hfill
\begin{minipage}[t]{0.54\textwidth}
	\centering\fontsize{10.5pt}{13pt}\selectfont
	\textbf{ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 4 (HYDROCARBON)}\\
	\textbf{MÔN: HOÁ HỌC - LỚP 11}\\
	\textit{Thời gian làm bài: 50 phút, không kể thời gian phát đề}
	
	\smallskip
	\fbox{\textbf{MÃ ĐỀ: 112}}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Mn = 55; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108. Thể tích khí đo ở đkc 25 $^\circ\text{C}$, 1 bar là 24,79 L/mol).}
\smallskip

\noindent\textbf{\large PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.}
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}
\vspace{0.2cm}

\Opensolutionfile{ans}[""" + ans_file + r"""]

% Câu 1
\begin{ex}
(SBT -- KNTT) Chất nào sau đây không phải là hydrocarbon?
\choice
{\ce{H3C-CH3}}
{\ce{H2C=CH2}}
{\ce{HC#CH}}
{\True \ce{CH3-CH2-OH}}
\loigiai{
\ce{CH3-CH2-OH} chứa nhóm chức alcohol (-\ce{OH}) gồm 3 nguyên tố C, H, O nên là dẫn xuất chứa oxygen của hydrocarbon, không phải là hydrocarbon (chỉ chứa C và H).
}
\end{ex}

% Câu 2
\begin{ex}
(SBT-CD) Những yếu tố nào sau đây không quyết định đến độ lớn của nhiệt độ sôi của các alkane?\\
(a) Phân tử khối.\\
(b) Tương tác van der Waals giữa các phân tử.\\
(c) Độ tan trong nước.\\
(d) Liên kết hydrogen giữa các phân tử.
\choice
{(a) và (b)}
{(b) và (c)}
{\True (c) và (d)}
{(a) và (d)}
\loigiai{
Nhiệt độ sôi của alkane phụ thuộc vào phân tử khối và độ lớn tương tác van der Waals giữa các phân tử. Alkane không tan trong nước và không có liên kết hydrogen liên phân tử, do đó (c) và (d) không quyết định độ lớn nhiệt độ sôi.
}
\end{ex}

% Câu 3
\begin{ex}
[KNTT - SBT] Phát biểu nào sau đây về ứng dụng của alkane \textbf{không} đúng?
\choice
{Propane \ce{C3H8} và butane \ce{C4H10} được sử dụng làm khí đốt}
{Các alkane \ce{C6}, \ce{C7}, \ce{C8} là nguyên liệu để sản xuất một số hydrocarbon thơm}
{Các alkane lỏng được sử dụng làm nhiên liệu như xăng hay dầu diesel}
{\True Các alkane từ \ce{C11} đến \ce{C20} được dùng làm nến và sáp}
\loigiai{
Các alkane rắn (thường có từ 20 nguyên tử carbon trở lên, sáp paraffin) mới được dùng làm nến và sáp. Các alkane từ \ce{C11} đến \ce{C20} ở trạng thái lỏng được dùng làm dầu hoả, nhiên liệu phản lực và dầu diesel.
}
\end{ex}

% Câu 4
\begin{ex}
(SBT-CD) Để tăng chất lượng của xăng, dầu, người ta thực hiện cách nào sau đây?
\choice
{\True Thực hiện phản ứng reforming để thay đổi cấu trúc của các alkane mạch không nhánh thành hydrocarbon mạch nhánh hoặc mạch vòng có chỉ số octane cao}
{Thực hiện phản ứng cracking để thay đổi cấu trúc các alkane mạch dài chuyển thành các alkene và alkane mạch ngắn hơn}
{Thực hiện phản ứng hydrogen hoá để chuyển các alkene thành alkane}
{Bổ sung thêm heptane vào xăng, dầu}
\loigiai{
Phản ứng reforming biến đổi alkane mạch không phân nhánh thành alkane phân nhánh hoặc hydrocarbon thơm (chỉ số octane cao), giúp nhiên liệu cháy êm, chống kích nổ hiệu quả.
}
\end{ex}

% Câu 5
\begin{ex}
[KNTT -- SBT] Hợp chất nào sau đây là một alkyne?
\choice
{\ce{CH3CH2CH2CH3}}
{\ce{CH3CH=CH2}}
{\True \ce{CH3CH2C#CH}}
{\ce{CH2=CH-CH=CH2}}
\loigiai{
Alkyne là hydrocarbon không no, mạch hở, trong phân tử có 1 liên kết ba $\ce{C#C}$. Do đó \ce{CH3CH2C#CH} (but-1-yne) là alkyne.
}
\end{ex}

% Câu 6
\begin{ex}
(SBT-CD) Các alkene \textbf{không} có tính chất vật lí đặc trưng nào sau đây?
\choice
{\True Tan tốt trong nước và các dung môi hữu cơ}
{Có khối lượng riêng nhỏ hơn khối lượng riêng của nước}
{Có nhiệt độ sôi thấp hơn alkane phân tử có cùng số nguyên tử carbon}
{Không dẫn điện}
\loigiai{
Phân tử các alkene không hoặc rất ít phân cực, nên không tan trong nước (dung môi phân cực mạnh), chỉ tan tốt trong các dung môi hữu cơ không phân cực.
}
\end{ex}

% Câu 7
\begin{ex}
[KNTT -- SBT] Phản ứng nào sau đây \textbf{không} phải là phản ứng đặc trưng của hydrocarbon không no?
\choice
{Phản ứng cộng}
{Phản ứng trùng hợp}
{Phản ứng oxi hóa -- khử}
{\True Phản ứng thế}
\loigiai{
Phản ứng thế là phản ứng đặc trưng của hydrocarbon no (alkane). Hydrocarbon không no có phản ứng đặc trưng là phản ứng cộng vào liên kết bội và phản ứng trùng hợp.
}
\end{ex}

% Câu 8
\begin{ex}
(SBT-CD) Cho các chất có công thức cấu tạo sau:\\
(1) \ce{ClCH2CH=CHCH3};\\
(2) \ce{CH3CH=CHCH3};\\
(3) \ce{BrCH2C(CH3)=C(CH2CH3)2};\\
(4) \ce{ClCH2CH=CH2};\\
(5) \ce{ClCH2CH=CHCH2CH3};\\
(6) \ce{(CH3)2C=CH2}.\\
Trong số các chất trên, bao nhiêu chất có đồng phân hình học?
\choice
{\True 3}
{4}
{5}
{6}
\loigiai{
Điều kiện có đồng phân hình học quanh liên kết đôi $\ce{C=C}$: mỗi nguyên tử carbon mang liên kết đôi phải liên kết với 2 nhóm nguyên tử khác nhau.\\
- (1) \ce{ClCH2CH=CHCH3}: có đphh (\ce{ClCH2} $\neq$ \ce{H} và \ce{CH3} $\neq$ \ce{H}).\\
- (2) \ce{CH3CH=CHCH3}: có đphh (\ce{CH3} $\neq$ \ce{H} và \ce{CH3} $\neq$ \ce{H}).\\
- (3) \ce{BrCH2C(CH3)=C(CH2CH3)2}: nguyên tử C thứ hai liên kết với hai nhóm \ce{-CH2CH3} giống nhau $\to$ không có đphh.\\
- (4) \ce{ClCH2CH=CH2}: C cuối liên kết với hai nguyên tử H giống nhau $\to$ không có đphh.\\
- (5) \ce{ClCH2CH=CHCH2CH3}: có đphh (\ce{ClCH2} $\neq$ \ce{H} và \ce{C2H5} $\neq$ \ce{H}).\\
- (6) \ce{(CH3)2C=CH2}: không có đphh.\\
Vậy có 3 chất: (1), (2), (5).
}
\end{ex}

% Câu 9
\begin{ex}
(SBT-CD) Phát biểu nào sau đây là \textbf{không} đúng?
\choice
{Hydrocarbon thơm là những hydrocarbon trong phân tử có chứa vòng benzene}
{\True Các chất trong phân tử có vòng benzene được gọi là hydrocarbon thơm}
{Những hydrocarbon trong phân tử có chứa vòng benzene được gọi là arene}
{Dãy đồng đẳng của benzene có công thức chung là \ce{C_nH_{2n-6}} ($n \ge 6$)}
\loigiai{
Phát biểu B sai vì các chất chứa vòng benzene nhưng có thêm các nguyên tố O, N, halogen (như phenol \ce{C6H5OH}, aniline \ce{C6H5NH2}) là dẫn xuất thơm, không phải là hydrocarbon thơm (hydrocarbon chỉ chứa C và H).
}
\end{ex}

% Câu 10
\begin{ex}
(SBT -- KNTT) Cho các chất sau: (X) \textit{o}-bromotoluene; (Y) \textit{m}-bromotoluene; (Z) \textit{p}-bromotoluene. Sản phẩm chính của phản ứng giữa toluene với bromine ở nhiệt độ cao có mặt iron(III) bromide là
\choice
{(X) và (Y)}
{(Y) và (Z)}
{\True (X) và (Z)}
{(Y)}
\loigiai{
Nhóm methyl (\ce{-CH3}) là nhóm đẩy electron làm tăng mật độ điện tích tại vị trí \textit{ortho} và \textit{para}, định hướng phản ứng thế electrophile vào hai vị trí này, do đó sản phẩm chính là (X) \textit{o}-bromotoluene và (Z) \textit{p}-bromotoluene.
}
\end{ex}

% Câu 11
\begin{ex}
(SBT -- KNTT) Nhận xét nào sau đây về tính chất hoá học của benzene là \textbf{không} đúng?
\choice
{Benzene khó tham gia phản ứng cộng hơn ethylene}
{Benzene dễ tham gia phản ứng thế hơn so với phản ứng cộng}
{Benzene không bị oxi hoá bởi tác nhân oxi hoá thông thường}
{\True Benzene làm mất màu dung dịch nước bromine ở điều kiện thường}
\loigiai{
Hệ liên hợp $\pi$ khép kín của benzene rất bền vững, benzene không phản ứng và không làm mất màu nước bromine ở điều kiện thường.
}
\end{ex}

% Câu 12
\begin{ex}
(SBT-CD) Một số chất gây ô nhiễm môi trường như benzene, toluene có trong khí thải đốt cháy nhiên liệu xăng, dầu. Để giảm thiểu nguyên nhân gây ô nhiễm này cần
\choice
{cấm sử dụng nhiên liệu xăng}
{\True hạn chế sử dụng nhiên liệu hoá thạch}
{thay xăng bằng khí gas}
{cấm sử dụng xe cá nhân}
\loigiai{
Hạn chế sử dụng nhiên liệu hoá thạch và chuyển dần sang các nguồn năng lượng tái tạo, năng lượng sạch là giải pháp căn cơ, hiệu quả và thực tiễn nhất.
}
\end{ex}

% Câu 13
\begin{ex}
(SBT-CD) Hydrocarbon Y có công thức cấu tạo: \ce{(CH3)2CHCH2CH3}. Khi cho Y phản ứng với bromine theo tỉ lệ mol 1:1 có chiếu sáng có thể thu được bao nhiêu dẫn xuất monobromo là đồng phân cấu tạo của nhau?
\choice
{3}
{\True 4}
{5}
{6}
\loigiai{
Phân tử \ce{(CH3)2CH-CH2-CH3} có 4 vị trí carbon không tương đương (2 nhóm \ce{-CH3} ở C1, C2, C3, C4) nên tạo ra 4 dẫn xuất monobromo đồng phân cấu tạo.
}
\end{ex}

% Câu 14
\begin{ex}
(SBT-CD) Tên thay thế của hydrocarbon có công thức cấu tạo \ce{(CH3)3CCH2CH2CH3} là
\choice
{\True 2,2-dimethylpentane}
{2,3-dimethylpentane}
{2,2,3-trimethylbutane}
{2,2-dimethylbutane}
\loigiai{
Mạch chính dài nhất gồm 5 carbon (pentane), đánh số từ đầu gần nhánh hơn: vị trí C số 2 mang 2 nhóm methyl $\to$ 2,2-dimethylpentane.
}
\end{ex}

% Câu 15
\begin{ex}
(SBT-CD) Để phân biệt but-2-yne (\ce{CH3C#CCH3}) với but-1-yne (\ce{CH#CCH2CH3}) có thể dùng thuốc thử nào sau đây?
\choice
{Dung dịch \ce{HCl}}
{\True Dung dịch \ce{AgNO3/NH3}}
{Nước bromine}
{Dung dịch \ce{KMnO4}}
\loigiai{
But-1-yne có liên kết ba đầu mạch tạo kết tủa màu vàng nhạt \ce{AgC#CCH2CH3} với dung dịch \ce{AgNO3/NH3}; but-2-yne có liên kết ba giữa mạch nên không phản ứng.
}
\end{ex}

% Câu 16
\begin{ex}
(SBT -- KNTT) Đun nóng hydrocarbon thơm X có công thức phân tử \ce{C8H10} với dung dịch \ce{KMnO4} nóng thu được dung dịch có chứa \ce{C6H5COOK} và \ce{K2CO3}. Chất X là
\choice
{\textit{o}-xylene}
{\textit{p}-xylene}
{\True ethylbenzene}
{styrene}
\loigiai{
Phản ứng oxi hoá ethylbenzene bằng \ce{KMnO4} nóng bẻ gãy mạch nhánh tại carbon benzylic:\\
\ce{C6H5CH2CH3 + 4KMnO4 -> C6H5COOK + K2CO3 + 4MnO2 + KOH + 2H2O}.\\
Do đó chất X là ethylbenzene.
}
\end{ex}

% Câu 17
\begin{ex}
Có 4 tên gọi: \textit{o}-xylene; \textit{o}-dimethylbenzene; 1,2-dimethylbenzene; ethylbenzene. Đó là tên của mấy chất?
\choice
{1 chất}
{\True 2 chất}
{3 chất}
{4 chất}
\loigiai{
- \textit{o}-xylene, \textit{o}-dimethylbenzene và 1,2-dimethylbenzene đều là tên của cùng một chất: 1,2-\ce{(CH3)2C6H4}.\\
- ethylbenzene là chất thứ hai: \ce{C6H5CH2CH3}.\\
Vậy có tất cả 2 chất khác nhau.
}
\end{ex}

% Câu 18
\begin{ex}
Khu vực trạm xăng dầu, khu vực có nhiều xe cơ giới qua lại, nơi có khói thuốc lá... lại được xem là nơi gây tổn hại đến sức khoẻ con người. Phát biểu nào sau đây đúng?
\choice
{\True Khu vực có trạm xăng dầu, khu vực có nhiều xe cơ giới qua lại, nơi có khói thuốc lá,... là nơi có nguồn hydrocarbon thơm, đây là các chất độc gây tổn hại đến sức khoẻ con người}
{Trong xăng tồn tại ethanol}
{Trong thuốc lá tồn tại nicotine (là hợp chất chứa vòng benzene) không độc}
{Trong xăng không tồn tại benzene (\ce{C6H6})}
\loigiai{
Khói thuốc lá và khí thải xăng dầu chứa nhiều hydrocarbon thơm độc hại (benzene, toluene, PAHs) gây ức chế tủy xương và nguy cơ ung thư cao.
}
\end{ex}

\Closesolutionfile{ans}

\bigskip
\noindent\textbf{\large PHẦN II. Câu trắc nghiệm đúng sai.}
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}
\vspace{0.2cm}

% Câu 1 Phần II
\begin{ex}
So sánh và giải thích nhiệt độ sôi của alkane mạch không phân nhánh với alkane mạch phân nhánh khi chúng có cùng số nguyên tử carbon trong phân tử.
\choiceTF
{\True Các phân tử alkane phân nhánh có nhiệt độ sôi thấp hơn các phân tử alkane không phân nhánh dù cho chúng có cùng số nguyên tử carbon}
{Các phân tử alkane phân nhánh có nhiệt độ sôi cao hơn các phân tử alkane không phân nhánh dù cho chúng có cùng số nguyên tử carbon}
{\True Các phân tử alkane không phân nhánh có diện tích bề mặt lớn hơn so với phân tử alkane phân nhánh nên tồn tại tương tác van der Waals giữa các phân tử lớn hơn, dẫn đến nhiệt độ sôi cao hơn}
{Các phân tử alkane không phân nhánh có diện tích bề mặt nhỏ hơn so với phân tử alkane phân nhánh nên tồn tại tương tác van der Waals giữa các phân tử nhỏ hơn, dẫn đến nhiệt độ sôi thấp hơn}
\loigiai{
- Alkane mạch không phân nhánh có diện tích tiếp xúc phân tử lớn hơn alkane mạch phân nhánh (cấu trúc gọn hình cầu).\\
- Tương tác van der Waals giữa các phân tử alkane mạch thẳng lớn hơn, cần nhiều nhiệt lượng hơn để bứt rời phân tử khi sôi $\to$ nhiệt độ sôi cao hơn.\\
$\Rightarrow$ a: ĐÚNG, b: SAI, c: ĐÚNG, d: SAI.
}
\end{ex}

% Câu 2 Phần II
\begin{ex}
Không được dùng nước dập tắt đám cháy có mặt đất đèn (có thành phần chính là \ce{CaC2}).
\choiceTF
{\True \ce{CaC2} tác dụng được với nước có thể tạo ra lượng lớn acetylene (\ce{C2H2}), là chất khí dễ cháy nổ}
{\True Khi acetylene được đốt cháy trong không khí ở nhiệt độ cao, nó phát ra nhiều nhiệt và có thể làm đám cháy lớn hơn và toả ra nhiều nhiệt hơn nên gây nguy hiểm cho người dập tắt đám cháy và môi trường xung quanh}
{\True Phản ứng của \ce{CaC2} với nước: \ce{CaC2 + 2H2O -> Ca(OH)2 + C2H2}}
{\True Phản ứng đốt cháy \ce{C2H2} với oxygen: \ce{2C2H2 + 5O2 ->[{t$^\circ$}] 4CO2 + 2H2O}, $\Delta_r H^0_{298} = -2602\text{ kJ}$}
\loigiai{
- Đất đèn phản ứng mãnh liệt với nước tạo khí acetylene dễ cháy.\\
- Acetylene cháy toả nhiệt lượng khổng lồ (-2602 kJ cho 2 mol) làm đám cháy bùng phát dữ dội hơn, cực kỳ nguy hiểm.\\
$\Rightarrow$ Cả 4 ý a, b, c, d đều ĐÚNG.
}
\end{ex}

% Câu 3 Phần II
\begin{ex}
Nitrobenzene là sản phẩm hữu cơ được tạo thành khi nitro hoá benzene bằng hỗn hợp \ce{HNO3} đặc và \ce{H2SO4} đặc ở nhiệt độ $\le 50^\circ\text{C}$.
\choiceTF
{Nitrobenzene được ứng dụng trong nông nghiệp để tổng hợp chất kích thích tăng trưởng thực vật}
{\True Nitrobenzene là nguyên liệu sản xuất aniline, tiền chất để tổng hợp polyaniline, cao su nhân tạo}
{Nitrobenzene là nguyên liệu để sản xuất thuốc nổ TNB sử dụng trong quân sự, công nghiệp và khai thác mỏ}
{Nitrobenzene là chất lỏng, màu trắng, không tan trong nước, sánh như dầu}
\loigiai{
- a) Sai: Nitrobenzene rất độc hại, không dùng làm chất kích thích sinh trưởng nông nghiệp.\\
- b) Đúng: Khử nitrobenzene bằng H$_2$ hoặc kim loại thu được aniline, nguyên liệu then chốt tổng hợp polyaniline và cao su nhân tạo.\\
- c) Sai: TNB là trinitrobenzene, nitrobenzene chủ yếu dùng điều chế aniline.\\
- d) Sai: Nitrobenzene là chất lỏng màu vàng nhạt (không phải màu trắng), mùi hạnh nhân đắng.\\
$\Rightarrow$ a: SAI, b: ĐÚNG, c: SAI, d: SAI.
}
\end{ex}

% Câu 4 Phần II
\begin{ex}
Trước những năm 50 của thế kỷ XX, công nghiệp tổng hợp hữu cơ dựa trên nguyên liệu chính là acetylene. Ngày nay, người ta thường dùng ethylene.
\choiceTF
{\True Ethylene là nguyên liệu rẻ hơn, tiện lợi hơn nhiều so với acetylene (ethylene thu được từ quá trình khai thác và chế biến dầu mỏ)}
{Acetylene là nguyên liệu rẻ hơn, tiện lợi hơn nhiều so với ethylene (ethylene thu được từ quá trình khai thác và chế biến dầu mỏ)}
{\True Phương pháp điều chế các monomer để tổng hợp các polymer đi từ ethylene kinh tế hơn và ít ảnh hưởng đến môi trường}
{\True Sơ đồ điều chế vinyl chloride từ ethylene dưới đây có chất thải ra môi trường chỉ là nước:}
\begin{center}
	\includegraphics[width=0.88\textwidth]{images_crop_hoa11_de12/so_do_vinyl_chloride.png}
\end{center}
\loigiai{
- Quá trình oxy-clo hoá ethylene khép kín: \ce{2CH2=CH2 + Cl2 + 1/2 O2 -> 2CH2=CHCl + H2O}.\\
- Toàn bộ \ce{HCl} sinh ra được tái sử dụng, chất thải duy nhất ra môi trường là nước (\ce{H2O}).\\
$\Rightarrow$ a: ĐÚNG, b: SAI, c: ĐÚNG, d: ĐÚNG.
}
\end{ex}

\bigskip
\noindent\textbf{\large PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.}
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}
\vspace{0.2cm}

% Câu 1 Phần III
\begin{ex}
Số dẫn xuất monobromine tối đa thu được khi cho isopentane tác dụng với bromine trong điều kiện có chiếu sáng là bao nhiêu?
\shortans{4}
\loigiai{
Isopentane \ce{(CH3)2CH-CH2-CH3} có 4 vị trí C không tương đương:\\
1) \ce{CH2Br-CH(CH3)-CH2-CH3} (1-bromo-2-methylbutane)\\
2) \ce{CH3-CBr(CH3)-CH2-CH3} (2-bromo-2-methylbutane, sản phẩm chính)\\
3) \ce{CH3-CH(CH3)-CHBr-CH3} (2-bromo-3-methylbutane)\\
4) \ce{CH3-CH(CH3)-CH2-CH2Br} (1-bromo-3-methylbutane)\\
$\Rightarrow$ Đáp số: \textbf{4}.
}
\end{ex}

% Câu 2 Phần III
\begin{ex}
Có bao nhiêu đồng phân alkyne ứng với công thức phân tử \ce{C5H8}?
\shortans{3}
\loigiai{
Các đồng phân alkyne mạch hở \ce{C5H8}:\\
(1) \ce{CH#C-CH2-CH2-CH3} (pent-1-yne)\\
(2) \ce{CH3-C#C-CH2-CH3} (pent-2-yne)\\
(3) \ce{CH#C-CH(CH3)2} (3-methylbut-1-yne)\\
$\Rightarrow$ Đáp số: \textbf{3}.
}
\end{ex}

% Câu 3 Phần III
\begin{ex}
Terephthalic acid (\textit{p}-\ce{HOOCC6H4COOH}) là nguyên liệu để sản xuất poly(ethylene terephthalate) (PET, một loại polymer quan trọng được sử dụng làm sợi dệt và chai nhựa,...). Cho các phát biểu sau:\\
a) Phương trình hoá học của phản ứng điều chế terephthalic acid từ \textit{p}-xylene:
\begin{center}
	\includegraphics[width=0.72\textwidth]{images_crop_hoa11_de12/so_do_terephthalic_acid.png}
\end{center}
b) \textit{p}-xylene còn có tên gọi khác là 1,3-dimethylbenzene.\\
c) Terephthalic acid (\textit{p}-\ce{HOOCC6H4COOH}) là một chất hữu cơ tạp chức.\\
d) Trong phân tử \textit{p}-xylene có chứa 8 nguyên tử carbon.\\
Số phát biểu đúng là bao nhiêu?
\shortans{2}
\loigiai{
- a) Đúng: Oxi hoá \textit{p}-xylene bằng không khí có xúc tác tạo terephthalic acid.\\
- b) Sai: \textit{p}-xylene là 1,4-dimethylbenzene.\\
- c) Sai: Terephthalic acid chỉ chứa 2 nhóm \ce{-COOH} giống nhau nên là chất đa chức, không phải tạp chức.\\
- d) Đúng: \textit{p}-xylene có CTPT là \ce{C8H10}, chứa 8 nguyên tử carbon.\\
$\Rightarrow$ Có 2 phát biểu đúng (a và d). Đáp số: \textbf{2}.
}
\end{ex}

% Câu 4 Phần III
\begin{ex}
Acetylene dùng trong đèn xì oxygen-acetylene để hàn, cắt kim loại. Trong công nghiệp acetylene là nguyên liệu để sản xuất poly(vinyl chloride) dùng sản xuất nhựa PVC và nhiều hóa chất khác. Acetylene cháy theo phương trình hóa học sau:
\[
\ce{C2H2}(g) + 2{,}5\ce{O2}(g) \xrightarrow{t^\circ} 2\ce{CO2}(g) + \ce{H2O}(g)
\]
Dựa vào nhiệt tạo thành chuẩn ở 298 K của các chất: $\Delta_f H^\circ_{298}$ của \ce{C2H2}(g) = $+227{,}0\text{ kJ/mol}$; \ce{CO2}(g) = $-393{,}5\text{ kJ/mol}$; \ce{H2O}(g) = $-241{,}82\text{ kJ/mol}$. Hãy cho biết biến thiên enthalpy của phản ứng cháy trên là $a\text{ kJ}$. Xác định $a$ (làm tròn đến hàng đơn vị).
\shortans{-1256}
\loigiai{
Biến thiên enthalpy của phản ứng:
\[
\Delta_r H^\circ_{298} = 2 \Delta_f H^\circ_{298}(\ce{CO2}, g) + \Delta_f H^\circ_{298}(\ce{H2O}, g) - \Delta_f H^\circ_{298}(\ce{C2H2}, g) - 2{,}5 \Delta_f H^\circ_{298}(\ce{O2}, g)
\]
Thay số:
\[
\Delta_r H^\circ_{298} = 2(-393{,}5) + (-241{,}82) - (+227{,}0) - 0 = -787{,}0 - 241{,}82 - 227{,}0 = -1255{,}82\text{ kJ}.
\]
Làm tròn đến hàng đơn vị: $a = -1256$.\\
$\Rightarrow$ Đáp số: \textbf{-1256} (hoặc $-1255{,}8$).
}
\end{ex}

% Câu 5 Phần III
\begin{ex}
Khí gas chứa chủ yếu các thành phần chính: Propane (\ce{C3H8}), butane (\ce{C4H10}) và một số thành phần khác. Để tạo mùi cho gas nhà sản xuất đã pha trộn thêm chất tạo mùi đặc trưng như methanethiol (\ce{CH3SH}), có mùi giống tỏi, hành tây. Thành phần khí gas, tỉ lệ hòa trộn phổ biến của propane : butane theo thứ tự 30 : 70 đến 50 : 50.\\
Cho phương trình nhiệt hóa học sau:
\[
\ce{C3H8}(g) + 5\ce{O2}(g) \xrightarrow{t^\circ} 3\ce{CO2}(g) + 4\ce{H2O}(g) \quad \Delta_r H^\circ_{298} = -2220\text{ kJ}
\]
\[
\ce{C4H10}(g) + \frac{13}{2}\ce{O2}(g) \xrightarrow{t^\circ} 4\ce{CO2}(g) + 5\ce{H2O}(g) \quad \Delta_r H^\circ_{298} = -2874\text{ kJ}
\]
Giả sử một hộ gia đình cần 6 000 kJ nhiệt mỗi ngày, sau bao nhiêu ngày sẽ sử dụng hết 1 bình gas 12 kg với tỉ lệ khối lượng propane : butane là 30 : 70 (thành phần khác không đáng kể)? (Biết hiệu suất hấp thụ nhiệt khoảng 60\%, làm tròn đến hàng đơn vị).
\shortans{60}
\loigiai{
- Khối lượng từng khí trong bình 12 kg:\\
$m(\ce{C3H8}) = 12 \times 0{,}30 = 3{,}6\text{ kg} = 3600\text{ g} \implies n(\ce{C3H8}) = \dfrac{3600}{44} \approx 81{,}818\text{ mol}$.\\
$m(\ce{C4H10}) = 12 \times 0{,}70 = 8{,}4\text{ kg} = 8400\text{ g} \implies n(\ce{C4H10}) = \dfrac{8400}{58} \approx 144{,}828\text{ mol}$.\\
- Tổng nhiệt lượng toả ra khi đốt cháy hết bình gas:
\[
Q_{\text{toả}} = 81{,}818 \times 2220 + 144{,}828 \times 2874 \approx 181\,636 + 416\,236 = 597\,872\text{ kJ}.
\]
- Nhiệt lượng có ích thực tế gia đình nhận được ($H = 60\%$):
\[
Q_{\text{có ích}} = 597\,872 \times 60\% \approx 358\,723\text{ kJ}.
\]
- Số ngày sử dụng:
\[
N = \dfrac{358\,723}{6000} \approx 59{,}79 \approx 60\text{ ngày}.
\]
$\Rightarrow$ Đáp số: \textbf{60}.
}
\end{ex}

% Câu 6 Phần III
\begin{ex}
Một loại xăng chứa hỗn hợp hexane, heptane và 2,2,4-trimethylpentane (isooctane). Hóa hơi lượng xăng này được hơi xăng có tỉ khối so với \ce{H2} là 54,9. Tính tỉ lệ thể tích hơi xăng và không khí (20\% thể tích \ce{O2}, 80\% thể tích \ce{N2}) vừa đủ đốt cháy hết lượng xăng này là $\dfrac{n_{\text{xăng}}}{n_{\text{không khí}}} = \dfrac{1}{x}$. Xác định $x$ (làm tròn đến hai chữ số thập phân).
\shortans{60,25}
\loigiai{
- Hỗn hợp xăng gồm các alkane có công thức chung \ce{C_{\bar{n}}H_{2\bar{n}+2}}.\\
$\bar{M} = 54{,}9 \times 2 = 109{,}8\text{ g/mol} \implies 14\bar{n} + 2 = 109{,}8 \implies \bar{n} = 7{,}7$.\\
- Phương trình đốt cháy: \ce{C_{\bar{n}}H_{2\bar{n}+2} + \dfrac{3\bar{n}+1}{2} O2 -> \bar{n} CO2 + (\bar{n}+1) H2O}.\\
Tỉ lệ mol giữa \ce{O2} và xăng:
\[
\dfrac{n_{\ce{O2}}}{n_{\text{xăng}}} = \dfrac{3 \times 7{,}7 + 1}{2} = 12{,}05.
\]
- Do oxygen chiếm 20\% thể tích không khí:
\[
n_{\text{không khí}} = 5 \times n_{\ce{O2}} = 5 \times 12{,}05 \times n_{\text{xăng}} = 60{,}25 \times n_{\text{xăng}}.
\]
\[
\implies \dfrac{n_{\text{xăng}}}{n_{\text{không khí}}} = \dfrac{1}{60{,}25} \implies x = 60{,}25.
\]
$\Rightarrow$ Đáp số: \textbf{60,25}.
}
\end{ex}

\bigskip
\noindent\textbf{\large PHẦN IV. Câu hỏi tự luận.}
\textit{Thí sinh trình bày lời giải chi tiết vào giấy làm bài.}
\vspace{0.2cm}

% Câu 1 Tự luận
\begin{ex}
(SBT - CTST) Nhu cầu sử dụng PVC trên toàn thế giới liên tục tăng trong các năm qua. Để thu được PVC, cần đi từ monomer là vinyl chloride. Có thể điều chế vinyl chloride từ acetylene hoặc ethylene. Một trong những cách điều chế vinyl chloride từ ethylene hiện nay là theo sơ đồ:
\[
\ce{C2H4} \xrightarrow{+\ce{Cl2}} (X) \xrightarrow{-\ce{HCl},\, 500^\circ\text{C}} \ce{CH2=CH-Cl}
\]
Viết các phương trình phản ứng minh hoạ cho sơ đồ trên.
"""
    if is_teacher:
        tex += r"""
\loigiai{
Các phương trình phản ứng minh hoạ:\\
(1) Phản ứng cộng chlorine vào ethylene tạo 1,2-dichloroethane (chất X):
\[
\ce{CH2=CH2 + Cl2 ->[\ce{FeCl3},\, 50-60^\circ\text{C}] ClCH2-CH2Cl}
\]
(2) Nhiệt phân tách \ce{HCl} ở $500^\circ\text{C}$ tạo vinyl chloride:
\[
\ce{ClCH2-CH2Cl ->[500^\circ\text{C}] CH2=CH-Cl + HCl}
\]
\textit{(Lưu ý: trong phòng thí nghiệm có thể đun nóng 1,2-dichloroethane với KOH trong ethanol: \ce{ClCH2-CH2Cl + KOH ->[\ce{C2H5OH},\, t^\circ] CH2=CH-Cl + KCl + H2O})}.
}
"""
    else:
        tex += r"""
\vspace{4.0cm}
"""
    tex += r"""
\end{ex}

% Câu 2 Tự luận
\begin{ex}
(SBT-CD) Khí đốt hoá lỏng (Liquified Petroleum Gas, viết tắt là LPG) hay còn được gọi là gas, là hỗn hợp khí chủ yếu gồm propane (\ce{C3H8}) và butane (\ce{C4H10}) đã được hoá lỏng. Một loại gas dân dụng chứa khí hoá lỏng có tỉ lệ mol propane : butane là 40 : 60. Đốt cháy 1 lít khí gas này (ở 25 $^\circ\text{C}$, 1 bar) thì toả ra một lượng nhiệt bằng bao nhiêu? Biết khi đốt cháy hoàn toàn 1 mol mỗi chất propane và butane toả ra lượng nhiệt tương ứng là 2 220 kJ và 2 875 kJ.
"""
    if is_teacher:
        tex += r"""
\loigiai{
- Ở điều kiện chuẩn ($25^\circ\text{C}$, 1 bar), 1 mol khí bất kỳ chiếm thể tích $24{,}79\text{ lít}$.\\
Tổng số mol khí trong 1 lít gas:
\[
n_{\text{gas}} = \dfrac{1}{24{,}79} \approx 0{,}040339\text{ mol}.
\]
- Số mol từng chất theo tỉ lệ mol 40 : 60:
\[
n_{\text{propane}} = 0{,}040339 \times 0{,}40 \approx 0{,}016136\text{ mol}.
\]
\[
n_{\text{butane}} = 0{,}040339 \times 0{,}60 \approx 0{,}024203\text{ mol}.
\]
- Lượng nhiệt toả ra khi đốt cháy 1 lít khí gas:
\[
Q = 0{,}016136 \times 2220 + 0{,}024203 \times 2875 \approx 35{,}822 + 69{,}584 = 105{,}41\text{ kJ}.
\]
\textbf{Kết luận:} Khi đốt cháy 1 lít khí gas trên toả ra khoảng \textbf{105,4 kJ} nhiệt lượng.
}
"""
    else:
        tex += r"""
\vspace{4.0cm}
"""
    tex += r"""
\end{ex}
"""

    if is_teacher:
        tex += r"""
\newpage
\noindent\textbf{\large BẢNG ĐÁP ÁN THAM KHẢO PHẦN I}
\vspace{0.2cm}

\begin{center}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1. D} & \textbf{2. C} & \textbf{3. D} & \textbf{4. A} & \textbf{5. C} & \textbf{6. A} & \textbf{7. D} & \textbf{8. A} & \textbf{9. B} \\
\hline
\textbf{10. C} & \textbf{11. D} & \textbf{12. B} & \textbf{13. B} & \textbf{14. A} & \textbf{15. B} & \textbf{16. C} & \textbf{17. B} & \textbf{18. A} \\
\hline
\end{tabular}
\end{center}

\bigskip
\noindent\textbf{\large BẢNG ĐÁP ÁN THAM KHẢO PHẦN II}
\vspace{0.2cm}

\begin{center}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Ý a)} & \textbf{Ý b)} & \textbf{Ý c)} & \textbf{Ý d)} \\
\hline
\textbf{Câu 1} & Đ & S & Đ & S \\
\hline
\textbf{Câu 2} & Đ & Đ & Đ & Đ \\
\hline
\textbf{Câu 3} & S & Đ & S & S \\
\hline
\textbf{Câu 4} & Đ & S & Đ & Đ \\
\hline
\end{tabular}
\end{center}

\bigskip
\noindent\textbf{\large BẢNG ĐÁP ÁN THAM KHẢO PHẦN III}
\vspace{0.2cm}

\begin{center}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\
\hline
4 & 3 & 2 & -1256 & 60 & 60{,}25 \\
\hline
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
    hs_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De12_HocSinh.tex")
    with open(hs_path, "w", encoding="utf-8") as f:
        f.write(hs_tex)
    print("Saved:", hs_path)

    gv_tex = make_tex(is_teacher=True)
    gv_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De12_GiaoVien.tex")
    with open(gv_path, "w", encoding="utf-8") as f:
        f.write(gv_tex)
    print("Saved:", gv_path)

    # Compile with pdflatex
    for name in ["De_Kiem_Tra_Hoa_11_De12_HocSinh", "De_Kiem_Tra_Hoa_11_De12_GiaoVien"]:
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
                if file.endswith(ext) or file.startswith("ans-hoa11-de12"):
                    try:
                        os.remove(os.path.join(root, file))
                    except:
                        pass
    print("LaTeX generation completed successfully!")

if __name__ == "__main__":
    main()
