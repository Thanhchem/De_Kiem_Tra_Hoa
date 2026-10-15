# -*- coding: utf-8 -*-
"""
Tự động tạo 2 file LaTeX và biên dịch ra PDF cho ĐỀ SỐ 4 - HÓA HỌC 11:
- De_Kiem_Tra_Hoa_11_De04_HocSinh.tex / .pdf
- De_Kiem_Tra_Hoa_11_De04_GiaoVien.tex / .pdf
Chương 2: NITROGEN – SULFUR
Tuân thủ Form chuẩn GEMINI.md, sử dụng gói ex_test.sty và mhchem.
"""

import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import subprocess
import shutil

TEX_DIR = r"c:\Antigravity_Thanh\San_Pham"

if not os.path.exists(os.path.join(TEX_DIR, "ex_test.sty")):
    shutil.copyfile(r"c:\Antigravity_Thanh\ex_test.sty", os.path.join(TEX_DIR, "ex_test.sty"))

def make_tex(is_teacher=False):
    ver_tag = "gv" if is_teacher else "hs"
    ans_file = f"ans-hoa11-de04{ver_tag}"
    
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
	\centering\small
	\textbf{ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 2 (NITROGEN -- SULFUR)}\\
	\textbf{MÔN: HOÁ HỌC -- LỚP 11}\\
	\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
	\fbox{\textbf{MÃ ĐỀ: 104}}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; Ar = 40; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Ba = 137).}
\smallskip

\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\Opensolutionfile{ans}[""" + ans_file + r"""]

\begin{ex}
Không khí sạch là không khí ít tạp chất chứa các thành phần cơ bản của không khí. Trường hợp nào sau đây được coi là không khí sạch?
\begin{center}
	\includegraphics[width=6.5cm]{images_crop_hoa11_de04/cau1_khong_khi_sach.png}
\end{center}
\choice
{Không khí chứa 78\% $\ce{N2}$, 17\% $\ce{O2}$, 3\% $\ce{CO2}$, 1\% $\ce{CO}$, 1\% $\ce{SO2}$.}
{Không khí chứa 78\% $\ce{N2}$, 18\% $\ce{O2}$, 4\% hỗn hợp $\ce{CO2}$, $\ce{SO2}$, $\ce{HCl}$.}
{Không khí chứa 78\% $\ce{N2}$, 20\% $\ce{O2}$, 2\% hỗn hợp $\ce{CO2}$, $\ce{CH4}$ và bụi.}
""" + (r"{\True Không khí chứa 78\% $\ce{N2}$, 21\% $\ce{O2}$, 1\% hỗn hợp $\ce{CO2}$, $\ce{H2O}$.}" if is_teacher else r"{Không khí chứa 78\% $\ce{N2}$, 21\% $\ce{O2}$, 1\% hỗn hợp $\ce{CO2}$, $\ce{H2O}$.}") + r"""
\loigiai{""" + (r"""
Thành phần tự nhiên của không khí sạch gồm khoảng 78\% thể tích $\ce{N2}$, 21\% $\ce{O2}$ và khoảng 1\% các khí khác ($\ce{CO2}$, hơi nước, khí hiếm). Các trường hợp có chứa khí độc như $\ce{CO}$, $\ce{SO2}$, $\ce{HCl}$ hay bụi mịn đều bị ô nhiễm.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
\textbf{[CTST - SGK]} Có các loại phân bón như $\ce{NH4Cl}$, $\ce{NH4NO3}$, $\ce{(NH4)2SO4}$, các loại phân bón này không thích hợp bón cho đất nào sau đây?
\choice
""" + (r"{\True Đất chua}" if is_teacher else r"{Đất chua}") + r"""
{Đất phù sa}
{Đất bạc màu}
{Đất nghèo dinh dưỡng.}
\loigiai{""" + (r"""
Các phân bón này chứa cation $\ce{NH4+}$, khi thuỷ phân trong nước tạo môi trường acid: $\ce{NH4+ + H2O <=> NH3 + H3O+}$ làm tăng độ chua của đất, do đó không thích hợp bón cho đất chua.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Biện pháp nào sau đây không hạn chế tình trạng mưa acid?
\choice
{Cho lắp đặt các thiết bị khử sunphua tại những nhà máy nhiệt điện.}
{Loại bỏ nitrogen và sulfur trong dầu mỏ, than.}
{Sử dụng các loại năng lượng có khả năng tái tạo thân thiện với môi trường.}
""" + (r"{\True Phát tán chất $\ce{NO2}$ và $\ce{SO2}$ vào bầu khí quyển.}" if is_teacher else r"{Phát tán chất $\ce{NO2}$ và $\ce{SO2}$ vào bầu khí quyển.}") + r"""
\loigiai{""" + (r"""
$\ce{NO2}$ và $\ce{SO2}$ là hai tác nhân chính gây ra mưa acid, việc phát tán các khí này vào khí quyển làm gia tăng mưa acid.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
\textbf{[KNTT - SBT]} Mưa acid là một thảm họa thiên nhiên toàn cầu, ảnh hưởng đến sự sống của các sinh vật. Mưa acid là hiện tượng mưa có pH
\choice
""" + (r"{\True $< 5{,}6.$}" if is_teacher else r"{$< 5{,}6.$}") + r"""
{$= 7$}
{$6 - 7.$}
{$> 8.$}
\loigiai{""" + (r"""
Nước mưa thông thường hòa tan $\ce{CO2}$ có $\text{pH} \approx 5{,}6$. Mưa acid xảy ra khi nước mưa có $\text{pH} < 5{,}6$ do hòa tan các acid mạnh $\ce{H2SO4}$ và $\ce{HNO3}$.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
\textbf{(SBT - KNTT)} Hiện tượng phú dưỡng là một biểu hiện của môi trường ao, hồ bị ô nhiễm do dư thừa các chất dinh dưỡng. Sự dư thừa dinh dưỡng chủ yếu do hàm lượng các ion nào sau đây vượt quá mức cho phép?
\choice
{Sodium, potassium.}
{Calcium, magnesium.}
""" + (r"{\True Nitrate, phosphate.}" if is_teacher else r"{Nitrate, phosphate.}") + r"""
{Chloride, sulfate.}
\loigiai{""" + (r"""
Hiện tượng phú dưỡng xảy ra do sự tích tụ quá mức của nguyên tố dinh dưỡng nitrogen (dưới dạng ion nitrate $\ce{NO3-}$) và phosphorus (dưới dạng ion phosphate $\ce{PO4^3-}$).\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
\textbf{(SBT - KNTT)} Trong số các chất khí: $\ce{SO2}$, $\ce{CO2}$, $\ce{O2}$, $\ce{N2}$, khí tan tốt trong nước ở điều kiện thường là
\choice
{$\ce{O2}$.}
{$\ce{CO2}$.}
""" + (r"{\True $\ce{SO2}$.}" if is_teacher else r"{$\ce{SO2}$.}") + r"""
{$\ce{N2}$.}
\loigiai{""" + (r"""
$\ce{SO2}$ là phân tử phân cực mạnh và phản ứng thuận nghịch với nước tạo $\ce{H2SO3}$, ở $20^\circ\text{C}$ 1 lít nước hòa tan được tới khoảng 40 lít khí $\ce{SO2}$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
\textbf{(SBT - CTST)} Dung dịch sulfuric acid đặc khác dung dịch sulfuric acid loãng ở tính chất hoá học nào?
\choice
{Tính base mạnh.}
""" + (r"{\True Tính oxi hóa mạnh.}" if is_teacher else r"{Tính oxi hóa mạnh.}") + r"""
{Tính acid mạnh.}
{Tính khử mạnh.}
\loigiai{""" + (r"""
$\ce{H2SO4}$ đặc có tính oxi hoá rất mạnh (do nguyên tử S ở mức oxi hoá cao nhất +6), có khả năng oxi hoá hầu hết kim loại (trừ Au, Pt), nhiều phi kim và hợp chất khử. $\ce{H2SO4}$ loãng chỉ thể hiện tính oxi hóa của $\ce{H+}$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
\textbf{(SBT - KNTT)} Bậc liên kết và năng lượng liên kết trong phân tử nitrogen tương ứng là
\choice
{2 và $418\text{ kJ/mol}$}
{1 và $167\text{ kJ/mol}.$}
{1 và $386\text{ kJ/mol}.$}
""" + (r"{\True 3 và $945\text{ kJ/mol}$}" if is_teacher else r"{3 và $945\text{ kJ/mol}$}") + r"""
\loigiai{""" + (r"""
Phân tử nitrogen có liên kết ba $\ce{N#N}$ với bậc liên kết bằng 3 và năng lượng liên kết rất lớn là $945\text{ kJ/mol}$.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
\textbf{(SBT - KNTT)} Xét cân bằng hóa học: $\ce{NH3 + H2O <=> NH4+ + OH-}$. Cân bằng sẽ chuyển dịch theo chiều thuận khi cho thêm vài giọt dung dịch nào sau đây?
\choice
{$\ce{NH4Cl}$.}
{$\ce{NaOH}$.}
""" + (r"{\True $\ce{HCl}$.}" if is_teacher else r"{$\ce{HCl}$.}") + r"""
{$\ce{NaCl}$.}
\loigiai{""" + (r"""
Dung dịch $\ce{HCl}$ cung cấp $\ce{H+}$, trung hòa bớt $\ce{OH-}$, làm nồng độ $\ce{OH-}$ giảm, cân bằng dịch chuyển theo chiều thuận tạo thêm $\ce{OH-}$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Tính base của $\ce{NH3}$ do nguyên nhân nào sau đây?
\choice
""" + (r"{\True Trên N còn cặp electron tự do.}" if is_teacher else r"{Trên N còn cặp electron tự do.}") + r"""
{Phân tử có 3 liên kết cộng hóa trị phân cực.}
{$\ce{NH3}$ tan được nhiều trong nước.}
{$\ce{NH3}$ tác dụng với nước tạo $\ce{NH4OH}$.}
\loigiai{""" + (r"""
Trên nguyên tử nitrogen còn một cặp electron chưa liên kết, có khả năng nhận proton $\ce{H+}$ tạo ion $\ce{NH4+}$, thể hiện tính base.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
$\ce{HNO3}$ đặc thể hiện tính oxi hóa khi tác dụng với chất nào dưới đây?
\choice
{$\ce{Fe2O3}$}
""" + (r"{\True $\ce{Fe(OH)2}$}" if is_teacher else r"{$\ce{Fe(OH)2}$}") + r"""
{$\ce{CuCl2}$}
{$\ce{Al2O3}$}
\loigiai{""" + (r"""
$\ce{Fe(OH)2}$ có số oxi hóa của Fe là +2 (chưa cực đại) nên bị $\ce{HNO3}$ đặc oxi hoá lên mức +3: $\ce{Fe(OH)2 + 4HNO3 -> Fe(NO3)3 + NO2 ^ + 3H2O}$. Các oxide $\ce{Fe2O3}$ và $\ce{Al2O3}$ đã ở mức +3 nên chỉ xảy ra phản ứng trao đổi acid - base thông thường.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phát biểu nào sau đây sai về $\ce{SO2}$?
\choice
{Chất chống nấm mốc lương thực, thực phẩm.}
{Sản xuất $\ce{H2SO4}$ trong công nghiệp.}
""" + (r"{\True Không gây ô nhiễm môi trường.}" if is_teacher else r"{Không gây ô nhiễm môi trường.}") + r"""
{Làm chất tẩy trắng giấy và bột giấy.}
\loigiai{""" + (r"""
$\ce{SO2}$ là chất khí độc hại, gây ô nhiễm không khí và là tác nhân chủ yếu gây mưa acid.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Nhóm các chất nào sau đây đều tác dụng được với dung dịch $\ce{H2SO4}$ loãng?
\choice
{$\ce{Fe, CuO, Cu(OH)2, BaCl2, NaCl}$.}
{$\ce{FeO, Cu, Cu(OH)2, BaCl2, Na2CO3}$.}
""" + (r"{\True $\ce{Fe2O3, Cu(OH)2, Zn, Na2SO3, Ba(NO3)2}$.}" if is_teacher else r"{$\ce{Fe2O3, Cu(OH)2, Zn, Na2SO3, Ba(NO3)2}$.}") + r"""
{$\ce{Fe(OH)3, Ag, CuO, KHCO3, MgS}$.}
\loigiai{""" + (r"""
Nhóm C gồm: $\ce{Fe2O3}$ (basic oxide), $\ce{Cu(OH)2}$ (base), $\ce{Zn}$ (kim loại đứng trước H), $\ce{Na2SO3}$ (muối acid yếu giải phóng $\ce{SO2}$), $\ce{Ba(NO3)2}$ (tạo kết tủa $\ce{BaSO4 v}$).\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Phát biểu nào sau đây sai?
\choice
{$\ce{SO2}$ dùng làm chất tẩy trắng giấy và bột giấy.}
{$\ce{SO2}$ dùng làm chất chống mốc lương thực, thực phẩm.}
""" + (r"{\True Trong phòng thí nghiệm, $\ce{SO2}$ được điều chế bằng cách đốt quặng pyrit (pyrite).}" if is_teacher else r"{Trong phòng thí nghiệm, $\ce{SO2}$ được điều chế bằng cách đốt quặng pyrit (pyrite).}") + r"""
{$\ce{SO3}$ là chất trung gian để sản xuất sulfuric acid.}
\loigiai{""" + (r"""
Đốt quặng pyrit ($\ce{FeS2}$) là phương pháp sản xuất $\ce{SO2}$ trong công nghiệp. Trong phòng thí nghiệm, $\ce{SO2}$ được điều chế bằng phản ứng: $\ce{Na2SO3 + H2SO4 -> Na2SO4 + SO2 ^ + H2O}$.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hiện tượng phú dưỡng là gì?
\choice
""" + (r"{\True Quá trình tích lũy quá mức các chất dinh dưỡng trong thủy vực, dẫn đến sự phát triển mạnh của thực vật thủy sinh, đặc biệt là tảo.}" if is_teacher else r"{Quá trình tích lũy quá mức các chất dinh dưỡng trong thủy vực, dẫn đến sự phát triển mạnh của thực vật thủy sinh, đặc biệt là tảo.}") + r"""
{Quá trình thiếu hụt các chất dinh dưỡng trong nước, làm giảm năng suất sinh học của hệ sinh thái.}
{Sự suy giảm chất lượng nước do sự gia tăng vi sinh vật gây bệnh.}
{Hiện tượng mưa acid làm giảm độ pH của nước hồ.}
\loigiai{""" + (r"""
Hiện tượng phú dưỡng là sự tích tụ quá mức các chất dinh dưỡng (nitrogen, phosphorus) trong các nguồn nước, kích thích tảo và thực vật thủy sinh phát triển bùng nổ.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Mưa acid được hình thành khi nào?
\choice
{Khi có sự bốc hơi của nước từ sông hồ.}
""" + (r"{\True Khi các khí như $\ce{SO2}$ và $\ce{NO_x}$ phản ứng với nước trong khí quyển tạo thành sulfuric acid và nitric acid.}" if is_teacher else r"{Khi các khí như $\ce{SO2}$ và $\ce{NO_x}$ phản ứng với nước trong khí quyển tạo thành sulfuric acid và nitric acid.}") + r"""
{Khi các hạt bụi trong không khí kết hợp với nước mưa.}
{Khi carbon dioxide ($\ce{CO2}$) trong không khí phản ứng với nước tạo thành carbonic acid.}
\loigiai{""" + (r"""
Mưa acid được hình thành khi các khí độc hại $\ce{SO2}$ và $\ce{NO_x}$ trong khí quyển bị oxi hóa và hòa tan vào các giọt nước mưa tạo thành $\ce{H2SO4}$ và $\ce{HNO3}$.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hậu quả chính của mưa acid đối với môi trường là gì?
\choice
{Làm tăng độ pH của nước sông, hồ, gây hại cho hệ sinh thái nước ngọt.}
{Làm tăng lượng chất dinh dưỡng trong đất, giúp cây cối phát triển.}
""" + (r"{\True Làm giảm độ pH của đất và nước, gây thiệt hại cho cây cối, thủy sinh vật và các tòa nhà.}" if is_teacher else r"{Làm giảm độ pH của đất và nước, gây thiệt hại cho cây cối, thủy sinh vật và các tòa nhà.}") + r"""
{Làm tăng sự phát triển của vi khuẩn trong môi trường đất.}
\loigiai{""" + (r"""
Mưa acid làm giảm độ pH của đất và nước, gây thiệt hại nghiêm trọng cho hệ sinh thái rừng, giết chết sinh vật thủy sinh và phá hủy kết cấu kim loại, đá của các công trình xây dựng.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các phát biểu sau:\\
(1) Đối với môi trường, khí sulfur dioxide là nguyên nhân quan trọng gây mưa acid.\\
(2) Chuyển hoá sulfur dioxide thành các chất ít gây ô nhiễm hơn bằng các hóa chất như vôi sống (thành phần chính là $\ce{CaO}$), vôi tôi (thành phần chính là $\ce{Ca(OH)2}$) hoặc đá vôi nghiền (thành phần chính là $\ce{CaCO3}$).\\
(3) Khí sulfur dioxide có màu vàng, ít tan trong nước, thể hiện cả tính oxi hoá và tính khử.\\
(4) Đối với con người, khí sulfur dioxide gây ra cảm giác khó thở, bệnh viêm đường hô hấp, đau mắt.\\
(5) Đơn chất $\ce{N2}$ khó hóa lỏng và tan nhiều trong nước.\\
(6) Tương tác van der Waals giữa các phân tử nitrogen yếu nên khó hóa lỏng.\\
Số phát biểu đúng là
\choice
{3.}
""" + (r"{\True 4.}" if is_teacher else r"{4.}") + r"""
{5.}
{6.}
\loigiai{""" + (r"""
Có 4 phát biểu đúng gồm: (1), (2), (4), (6).\\
- (3) Sai vì $\ce{SO2}$ là khí không màu, tan nhiều trong nước.\\
- (5) Sai vì $\ce{N2}$ rất ít tan trong nước.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\Closesolutionfile{ans}

\bigskip
\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\begin{ex}
Trong phương trình hóa học của phản ứng giữa nitrogen và oxygen:
\choiceTF
""" + (r"{\True Phương trình hóa học: $\ce{N2 + O2 <=> 2NO}$ (ở $3000^\circ\text{C}$ hoặc tia lửa điện).}" if is_teacher else r"{Phương trình hóa học: $\ce{N2 + O2 <=> 2NO}$ (ở $3000^\circ\text{C}$ hoặc tia lửa điện).}") + r"""
""" + (r"{Oxygen giảm từ 0 xuống -1 $\rightarrow$ oxygen đóng vai trò chất oxi hóa trong phản ứng này.}" if is_teacher else r"{Oxygen giảm từ 0 xuống -1 $\rightarrow$ oxygen đóng vai trò chất oxi hóa trong phản ứng này.}") + r"""
""" + (r"{Nitrogen tăng từ 0 lên +4 $\rightarrow$ nitrogen đóng vai trò là chất khử trong phản ứng này.}" if is_teacher else r"{Nitrogen tăng từ 0 lên +4 $\rightarrow$ nitrogen đóng vai trò là chất khử trong phản ứng này.}") + r"""
""" + (r"{\True Phản ứng này xảy ra ở nhiệt độ cao trên $3000^\circ\text{C}$ hoặc có tia lửa điện nhưng hiệu suất tạo NO rất thấp nên thực tế không dùng để tạo NO.}" if is_teacher else r"{Phản ứng này xảy ra ở nhiệt độ cao trên $3000^\circ\text{C}$ hoặc có tia lửa điện nhưng hiệu suất tạo NO rất thấp nên thực tế không dùng để tạo NO.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Đúng}, ở $3000^\circ\text{C}$ hoặc có tia lửa điện $\ce{N2}$ phản ứng thuận nghịch với $\ce{O2}$.\\
b) \textbf{Sai} vì số oxi hoá của oxygen giảm từ 0 xuống -2 trong $\ce{NO}$.\\
c) \textbf{Sai} vì số oxi hoá của nitrogen tăng từ 0 lên +2 trong $\ce{NO}$.\\
d) \textbf{Đúng} vì phản ứng thu nhiệt mạnh ($\Delta_r H^\circ_{298} = +180\text{ kJ/mol}$) nên hiệu suất rất thấp.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Mưa acid gây tác động xấu đến môi trường, con người và sinh vật:
\choiceTF
""" + (r"{\True Khi nước mưa có pH nhỏ hơn 5,6 thì gọi là hiện tượng mưa acid.}" if is_teacher else r"{Khi nước mưa có pH nhỏ hơn 5,6 thì gọi là hiện tượng mưa acid.}") + r"""
""" + (r"{Hai tác nhân chính gây mưa acid là $\ce{Cl2}$, $\ce{HCl}$.}" if is_teacher else r"{Hai tác nhân chính gây mưa acid là $\ce{Cl2}$, $\ce{HCl}$.}") + r"""
""" + (r"{\True Mưa acid ăn mòn công trình xây dựng, kiến trúc bằng đá, cẩm thạch và kim loại.}" if is_teacher else r"{Mưa acid ăn mòn công trình xây dựng, kiến trúc bằng đá, cẩm thạch và kim loại.}") + r"""
""" + (r"{\True Mưa acid làm giảm độ pH của đất và nguồn nước sông hồ, ảnh hưởng nghiêm trọng đến đời sống thủy sinh.}" if is_teacher else r"{Mưa acid làm giảm độ pH của đất và nguồn nước sông hồ, ảnh hưởng nghiêm trọng đến đời sống thủy sinh.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Đúng} theo định nghĩa mưa acid ($\text{pH} < 5{,}6$).\\
b) \textbf{Sai} vì hai tác nhân chính gây mưa acid là $\ce{SO2}$ và các oxide của nitrogen ($\ce{NO_x}$).\\
c) \textbf{Đúng} vì acid trong nước mưa tác dụng với $\ce{CaCO3}$ trong đá và ăn mòn kim loại.\\
d) \textbf{Đúng} vì acid hóa làm thay đổi môi trường sống của các loài thủy sinh.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Thành phần chính của khí thiên nhiên là các hydrocarbon như methane (khoảng 80 -- 85\%), ethane, propane, butane cùng lượng nhỏ các khí carbon dioxide, hydrogen sulfide, nitrogen. Thành phần chính của than là carbon, ngoài ra còn có một số hợp chất của các nguyên tố H, S, O, N,...
\choiceTF
""" + (r"{\True Khi sử dụng khí thiên nhiên hoặc than làm nhiên liệu đều thải vào không khí các chất khí gây ô nhiễm môi trường không khí.}" if is_teacher else r"{Khi sử dụng khí thiên nhiên hoặc than làm nhiên liệu đều thải vào không khí các chất khí gây ô nhiễm môi trường không khí.}") + r"""
""" + (r"{\True Khí thiên nhiên và than khi cháy sẽ toả rất nhiều nhiệt, tạo điều kiện cho các hợp chất như hydrogen sulfide và các hợp chất chứa S, N bị oxi hoá sinh ra các khí độc gây ô nhiễm như $\ce{SO2}$, $\ce{NO}$, $\ce{NO2}$,...}" if is_teacher else r"{Khí thiên nhiên và than khi cháy sẽ toả rất nhiều nhiệt, tạo điều kiện cho các hợp chất như hydrogen sulfide và các hợp chất chứa S, N bị oxi hoá sinh ra các khí độc gây ô nhiễm như $\ce{SO2}$, $\ce{NO}$, $\ce{NO2}$,...}") + r"""
""" + (r"{Các khí $\ce{SO2}$, $\ce{NO2}$ này là nguyên nhân chính gây ra hiệu ứng nhà kính.}" if is_teacher else r"{Các khí $\ce{SO2}$, $\ce{NO2}$ này là nguyên nhân chính gây ra hiệu ứng nhà kính.}") + r"""
""" + (r"{Khí thiên nhiên và than đá khi cháy sẽ tạo ra sản phẩm chính là khí $\ce{CO2}$ gây mưa acid.}" if is_teacher else r"{Khí thiên nhiên và than đá khi cháy sẽ tạo ra sản phẩm chính là khí $\ce{CO2}$ gây mưa acid.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Đúng} vì quá trình đốt cháy luôn sinh ra các khí thải độc hại.\\
b) \textbf{Đúng}, nhiệt độ cao oxi hóa các hợp chất chứa lưu huỳnh và nitơ.\\
c) \textbf{Sai} vì $\ce{SO2}$ và $\ce{NO2}$ là tác nhân gây mưa acid, không phải khí nhà kính chính.\\
d) \textbf{Sai} vì $\ce{CO2}$ là tác nhân gây hiệu ứng nhà kính, không phải nguyên nhân chính của mưa acid.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Khi pha loãng dung dịch $\ce{H2SO4}$ đặc:
\choiceTF
""" + (r"{\True Sulfuric acid tan tốt trong nước, quá trình hoà tan toả ra một lượng nhiệt lớn.}" if is_teacher else r"{Sulfuric acid tan tốt trong nước, quá trình hoà tan toả ra một lượng nhiệt lớn.}") + r"""
""" + (r"{Sulfuric acid tan tốt trong nước, quá trình hoà tan thu một lượng nhiệt lớn.}" if is_teacher else r"{Sulfuric acid tan tốt trong nước, quá trình hoà tan thu một lượng nhiệt lớn.}") + r"""
""" + (r"{Để đảm bảo an toàn phải cho từ từ nước vào acid và khuấy đều.}" if is_teacher else r"{Để đảm bảo an toàn phải cho từ từ nước vào acid và khuấy đều.}") + r"""
""" + (r"{\True Để đảm bảo an toàn phải cho từ từ acid vào nước và khuấy đều.}" if is_teacher else r"{Để đảm bảo an toàn phải cho từ từ acid vào nước và khuấy đều.}") + r"""
\loigiai{""" + (r"""
a) \textbf{Đúng} vì sự hydrate hoá $\ce{H2SO4}$ giải phóng nhiệt lượng rất lớn.\\
b) \textbf{Sai} vì quá trình này toả nhiệt mạnh chứ không thu nhiệt.\\
c) \textbf{Sai} vì rót nước vào acid đặc sẽ làm sôi cục bộ và bắn acid ra ngoài gây nguy hiểm.\\
d) \textbf{Đúng}, quy tắc an toàn: rót từ từ acid vào nước và khuấy đều.""" if is_teacher else "") + r"""}
\end{ex}

\bigskip
\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}

\begin{ex}
Cho sơ đồ chuyển hóa nitrogen trong khí quyển thành phân đạm:
\[\ce{N2 ->[+O2] NO ->[+O2] NO2 ->[+O2 + H2O] HNO3 -> NO3-}\]
Số phản ứng thuộc loại oxi hóa -- khử trong sơ đồ là:
""" + (r"""
\loigiai{
Có \textbf{3} phản ứng oxi hoá -- khử:\\
(1) $\ce{N2^0 + O2 -> 2N^{+2}O}$\\
(2) $\ce{2N^{+2}O + O2 -> 2N^{+4}O2}$\\
(3) $\ce{4N^{+4}O2 + O2 + 2H2O -> 4HN^{+5}O3}$\\
Giai đoạn $\ce{HNO3 -> NO3-}$ là phản ứng trao đổi / điện li, số oxi hoá của N không đổi (+5).\\
\textbf{Đáp số: 3.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
Có bao nhiêu ý kiến sau đây về sulfur dioxide ($\ce{SO2}$) là đúng?\\
(1) Có độc tính đối với con người.\\
(2) Phản ứng được với đá vôi.\\
(3) Khí này được tạo thành từ hoạt động của núi lửa trong tự nhiên, từ quá trình đốt cháy nhiên liệu hóa thạch của con người,...\\
(4) Là oxide lưỡng tính.
""" + (r"""
\loigiai{
Có \textbf{3} ý kiến đúng gồm: (1), (2), (3).\\
- (1) Đúng, $\ce{SO2}$ gây hại cho đường hô hấp.\\
- (2) Đúng, $\ce{SO2 + CaCO3 -> CaSO3 + CO2 ^}$.\\
- (3) Đúng về nguồn gốc tự nhiên và nhân tạo.\\
- (4) Sai vì $\ce{SO2}$ là acidic oxide (oxide acid), không phải oxide lưỡng tính.\\
\textbf{Đáp số: 3.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
Tính thể tích ($\text{mL}$) dung dịch $\ce{H2SO4}\ 98\%$ (khối lượng riêng $D = 1{,}84\text{ g/mL}$) cần dùng để pha chế thành $500\text{ mL}$ dung dịch $\ce{H2SO4}\ 0{,}05\text{ M}$ (làm tròn đến hai chữ số phần thập phân).
""" + (r"""
\loigiai{
$n_{\ce{H2SO4}} = 0{,}5 \cdot 0{,}05 = 0{,}025\text{ mol} \rightarrow m_{\ce{H2SO4}} = 0{,}025 \cdot 98 = 2{,}45\text{ g}$.\\
$m_{\text{dd } \ce{H2SO4}} = \dfrac{2{,}45}{0{,}98} = 2{,}5\text{ g} \rightarrow V_{\text{dd}} = \dfrac{2{,}5}{1{,}84} \approx 1{,}36\text{ mL}$.\\
\textbf{Đáp số: 1,36.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
\textbf{(SBT - KNTT)} Tính phân tử khối trung bình của không khí, giả thiết thành phần không khí gồm: $78\%$ nitrogen, $21\%$ oxygen và $1\%$ argon (làm tròn kết quả đến phần nguyên).
""" + (r"""
\loigiai{
$\overline{M}_{\text{kk}} = 28{,}014 \cdot 0{,}78 + 31{,}998 \cdot 0{,}21 + 39{,}948 \cdot 0{,}01 \approx 28{,}97 \approx 29\text{ g/mol}$.\\
\textbf{Đáp số: 29.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
Hỗn hợp X gồm $\ce{N2}$ và $\ce{H2}$ có khối lượng mol trung bình $\overline{M}_X = 12{,}4\text{ g/mol}$. Dẫn X đi qua bình đựng bột Fe rồi nung nóng (hiệu suất tổng hợp $\ce{NH3}$ đạt $40\%$), thu được hỗn hợp Y. Giá trị khối lượng mol trung bình $\overline{M}_Y$ là bao nhiêu $\text{g/mol}$? (làm tròn kết quả đến hàng phần mười).
""" + (r"""
\loigiai{
Đường chéo: $\dfrac{n_{\ce{N2}}}{n_{\ce{H2}}} = \dfrac{12{,}4 - 2}{28 - 12{,}4} = \dfrac{10{,}4}{15{,}6} = \dfrac{2}{3}$.\\
Chọn $n_{\ce{N2}} = 2\text{ mol}$, $n_{\ce{H2}} = 3\text{ mol} \rightarrow n_X = 5\text{ mol}$, $m_X = 2 \cdot 28 + 3 \cdot 2 = 62\text{ g}$.\\
Hiệu suất tính theo $\ce{H2}$: $n_{\ce{H2}\text{ pứ}} = 3 \cdot 0{,}40 = 1{,}2\text{ mol}$.\\
$n_Y = n_X - 2 \cdot \dfrac{n_{\ce{H2}\text{ pứ}}}{3} = 5 - 0{,}8 = 4{,}2\text{ mol}$.\\
$\overline{M}_Y = \dfrac{m_Y}{n_Y} = \dfrac{62}{4{,}2} \approx 14{,}8\text{ g/mol}$.\\
\textbf{Đáp số: 14,8.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\begin{ex}
Một khu vực nông nghiệp gần hồ có diện tích $10\text{ km}^2$, hằng năm lượng nước chảy tràn từ mưa cuốn theo phân bón vào hồ chứa khoảng $5\text{ mg/L}$ nitrate ($\ce{NO3-}$). Nếu lượng nước chảy tràn chiếm $2\%$ lượng mưa, và lượng mưa trung bình hàng năm là $800\text{ mm}$, tính khối lượng nitrate (theo đơn vị $\text{kg}$) bị thải vào hồ mỗi năm.
""" + (r"""
\loigiai{
Diện tích: $S = 10\text{ km}^2 = 10^7\text{ m}^2$.\\
Thể tích nước mưa: $V_{\text{mưa}} = 800 \cdot 10^{-3}\text{ m} \cdot 10^7\text{ m}^2 = 8 \cdot 10^6\text{ m}^3 = 8 \cdot 10^9\text{ L}$.\\
Thể tích nước chảy tràn: $V_{\text{tràn}} = 0{,}02 \cdot 8 \cdot 10^9 = 1{,}6 \cdot 10^8\text{ L}$.\\
Khối lượng nitrate thải vào hồ: $m_{\ce{NO3-}} = 5\text{ mg/L} \cdot 1,6 \cdot 10^8\text{ L} = 8 \cdot 10^8\text{ mg} = 800\text{ kg}$.\\
\textbf{Đáp số: 800.}
}""" if is_teacher else r"""
\vspace{1.2cm}
""") + r"""
\end{ex}

\bigskip
\noindent
\textbf{PHẦN IV: Tự luận}\\
\textit{Học sinh giải chi tiết các câu hỏi sau vào giấy làm bài.}

\begin{ex}
\textbf{(SBT - CD)} Trong quy trình sản xuất tơ, mỗi năm có hàng triệu tấn cyclohexanone ($\ce{C6H10O}$) được cho phản ứng với $\ce{HNO3}$ để tạo adipic acid ($\ce{C6H10O4}$) theo phản ứng:
\[\ce{C6H10O + HNO3 -> C6H10O4 + N2O + H2O}\]
a) Cân bằng phương trình hoá học của phản ứng trên theo phương pháp thăng bằng electron.\\
b) Cho biết vai trò của $\ce{HNO3}$ trong phản ứng trên. Giải thích.
""" + (r"""
\loigiai{
a) Quá trình thăng bằng electron:\\
$\ce{C6H10O + 3H2O -> C6H10O4 + 6H+ + 6e} \quad |\times 4$\\
$\ce{2N^{+5} + 8e -> 2N^{+1} (trong\ N2O)} \quad |\times 3$\\
Phương trình hoàn chỉnh:\\
\[\ce{4C6H10O + 6HNO3 -> 4C6H10O4 + 3N2O + 3H2O}\]
b) Trong phản ứng trên, $\ce{HNO3}$ đóng vai trò là chất oxi hóa vì số oxi hóa của nguyên tử nitrogen giảm từ +5 (trong $\ce{HNO3}$) xuống +1 (trong $\ce{N2O}$).
}""" if is_teacher else r"""
\vspace{2.5cm}
""") + r"""
\end{ex}

\begin{ex}
\textbf{(SBT - CTST)} Muối $\ce{NH4NO3}$ sẽ nhiệt phân theo phản ứng nào trong 2 phản ứng sau? Giải thích.
\begin{align*}
	\ce{NH4NO3(s) &->[t^\circ] NH3(g) + HNO3(g)} \tag{1} \\
	\ce{NH4NO3(s) &->[t^\circ] N2O(g) + 2H2O(g)} \tag{2}
\end{align*}
Biết enthalpy tạo thành chuẩn $\Delta_f H^\circ_{298}\ (\text{kJ/mol})$ của các chất như sau:
$\ce{NH4NO3(s)}: -365{,}61$; $\ce{NH3(g)}: -45{,}90$; $\ce{N2O(g)}: 82{,}05$; $\ce{HNO3(g)}: -134{,}31$; $\ce{H2O(g)}: -241{,}82$.
""" + (r"""
\loigiai{
Tính biến thiên enthalpy chuẩn của hai phản ứng:\\
- Phản ứng (1):\\
$\Delta_r H^\circ_{298}(1) = \Delta_f H^\circ_{298}(\ce{NH3}) + \Delta_f H^\circ_{298}(\ce{HNO3}) - \Delta_f H^\circ_{298}(\ce{NH4NO3})$\\
$= (-45{,}90) + (-134{,}31) - (-365{,}61) = +185{,}40\text{ kJ} > 0$ (phản ứng thu nhiệt).\\
- Phản ứng (2):\\
$\Delta_r H^\circ_{298}(2) = \Delta_f H^\circ_{298}(\ce{N2O}) + 2 \cdot \Delta_f H^\circ_{298}(\ce{H2O}) - \Delta_f H^\circ_{298}(\ce{NH4NO3})$\\
$= 82{,}05 + 2 \cdot (-241{,}82) - (-365{,}61) = -35{,}97\text{ kJ} < 0$ (phản ứng toả nhiệt).\\
\textbf{Kết luận:} Vì $\Delta_r H^\circ_{298}(2) < 0$ nên phản ứng (2) toả nhiệt, thuận lợi hơn về mặt nhiệt động học, do đó muối $\ce{NH4NO3}$ sẽ nhiệt phân theo phản ứng (2) tạo khí $\ce{N2O}$ và hơi nước.
}""" if is_teacher else r"""
\vspace{3.5cm}
""") + r"""
\end{ex}
"""

    if is_teacher:
        tex += r"""
\newpage
\noindent
\textbf{\large BẢNG ĐÁP ÁN VÀ THANG ĐIỂM THAM KHẢO}
\smallskip

\noindent
\textbf{1. PHẦN I (Trắc nghiệm nhiều phương án):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1. D} & \textbf{2. A} & \textbf{3. D} & \textbf{4. A} & \textbf{5. C} & \textbf{6. C} & \textbf{7. B} & \textbf{8. D} & \textbf{9. C} \\ \hline
\textbf{10. A} & \textbf{11. B} & \textbf{12. C} & \textbf{13. C} & \textbf{14. C} & \textbf{15. A} & \textbf{16. B} & \textbf{17. C} & \textbf{18. B} \\ \hline
\end{tabular}
\end{center}

\noindent
\textbf{2. PHẦN II (Trắc nghiệm đúng sai):}
\begin{itemize}
	\item \textbf{Câu 1:} a) Đúng \quad|\quad b) Sai \quad|\quad c) Sai \quad|\quad d) Đúng
	\item \textbf{Câu 2:} a) Đúng \quad|\quad b) Sai \quad|\quad c) Đúng \quad|\quad d) Đúng
	\item \textbf{Câu 3:} a) Đúng \quad|\quad b) Đúng \quad|\quad c) Sai \quad|\quad d) Sai
	\item \textbf{Câu 4:} a) Đúng \quad|\quad b) Sai \quad|\quad c) Sai \quad|\quad d) Đúng
\end{itemize}

\noindent
\textbf{3. PHẦN III (Trả lời ngắn):}
\begin{center}
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} \\ \hline
3 & 3 & 1,36 & 29 & 14,8 & 800 \\ \hline
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
    hs_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De04_HocSinh.tex")
    with open(hs_path, "w", encoding="utf-8") as f:
        f.write(hs_tex)
    print("Saved:", hs_path)

    gv_tex = make_tex(is_teacher=True)
    gv_path = os.path.join(TEX_DIR, "De_Kiem_Tra_Hoa_11_De04_GiaoVien.tex")
    with open(gv_path, "w", encoding="utf-8") as f:
        f.write(gv_tex)
    print("Saved:", gv_path)

    # Compile with pdflatex
    for name in ["De_Kiem_Tra_Hoa_11_De04_HocSinh", "De_Kiem_Tra_Hoa_11_De04_GiaoVien"]:
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
                if file.endswith(ext) or file.startswith("ans-hoa11-de04"):
                    try:
                        os.remove(os.path.join(root, file))
                    except:
                        pass
    print("LaTeX generation completed successfully!")

if __name__ == "__main__":
    main()
