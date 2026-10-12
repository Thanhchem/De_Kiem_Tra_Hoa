# -*- coding: utf-8 -*-
"""
Script sinh file LaTeX (.tex) và biên dịch ra PDF cho:
ĐỀ KIỂM TRA SỐ 18 - CHƯƠNG 6: ĐẠI CƯƠNG KIM LOẠI (HÓA HỌC 12)
Mã đề thi: 118
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De18_HocSinh.pdf (Chuẩn chính xác 4 trang A4)
2. De_Kiem_Tra_Hoa_12_De18_GiaoVien.pdf (Bản Giáo Viên có lời giải chi tiết và đáp án)
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
import subprocess
import glob
import pymupdf

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"

def create_latex_de18(is_teacher=False):
    ans_file = "ans-de18-gv" if is_teacher else "ans-de18-hs"
    
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
	\textbf{\color{mainGreen}ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 6 (ĐẠI CƯƠNG KIM LOẠI)}\\
	\textbf{MÔN: HOÁ HỌC -- LỚP 12}\\
	\textit{Thời gian làm bài: 50 phút (không kể thời gian phát đề)}\\
	\textbf{Mã đề thi: 118}
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
\noindent\textit{(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108; Ba = 137; Au = 197; Pb = 207; hằng số Faraday F = 96485 C/mol).}

\vspace{1pt}
\noindent
\textbf{PHẦN I (4,5 điểm -- 18 câu): Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}

\begin{ex}
Ở điều kiện thường, kim loại nào sau đây ở trạng thái lỏng?
\choice
{Au.}
{Cu.}
{Na.}
""" + (r"{\True Hg.}" if is_teacher else r"{Hg.}") + r"""
\loigiai{""" + (r"""
Ở điều kiện thường ($25^\circ\text{C}$), thuỷ ngân (Hg) là kim loại duy nhất ở trạng thái lỏng (nhiệt độ nóng chảy $-38{,}83^\circ\text{C}$).\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Kim loại có độ cứng lớn nhất là
\choice
""" + (r"{\True Cr.}" if is_teacher else r"{Cr.}") + r"""
{Al.}
{Mg.}
{Na.}
\loigiai{""" + (r"""
Chromium (Cr) là kim loại có độ cứng lớn nhất trong tất cả các kim loại (độ cứng bằng 9 trên thang Mohs, có thể rạch được thuỷ tinh).\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Kim loại có nhiệt độ nóng chảy cao nhất là
\choice
{Na.}
{Pb.}
{Hg.}
""" + (r"{\True W.}" if is_teacher else r"{W.}") + r"""
\loigiai{""" + (r"""
Tungsten (W - vonfam) có nhiệt độ nóng chảy cao nhất trong các kim loại ($3422^\circ\text{C}$), được dùng làm dây tóc bóng đèn sợi đốt.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Kim loại có khối lượng riêng nhỏ nhất là
\choice
{Na.}
""" + (r"{\True Li.}" if is_teacher else r"{Li.}") + r"""
{Fe.}
{Al.}
\loigiai{""" + (r"""
Lithium (Li) có khối lượng riêng nhỏ nhất ($D = 0{,}534\text{ g/cm}^3$, nhẹ hơn nước và nhẹ nhất trong các kim loại).\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong mạng tinh thể kim loại chứa các cation kim loại được sắp xếp theo một trật tự nhất định cùng với các electron chuyển động
\choice
{theo một quỹ đạo xác định.}
{xung quanh một vị trí xác định.}
""" + (r"{\True tự do trong toàn bộ mạng tinh thể.}" if is_teacher else r"{tự do trong toàn bộ mạng tinh thể.}") + r"""
{trong một khu vực không gian nhất định.}
\loigiai{""" + (r"""
Trong mạng tinh thể kim loại, các electron hoá trị tách khỏi nguyên tử trở thành electron tự do, chuyển động hỗn loạn không ngừng trong toàn bộ không gian mạng tinh thể.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Hình dưới đây mô tả tính chất vật lí nào của kim loại? (hình tròn to mô tả ion kim loại, hình tròn nhỏ mô tả electron tự do)
\par\vspace{1pt}\centerline{\includegraphics[height=1.4cm]{images_crop_de18/p1_cau6_tinh_dan_dien.png}}\vspace{1pt}
\choice
{Tính dẻo.}
""" + (r"{\True Tính dẫn điện.}" if is_teacher else r"{Tính dẫn điện.}") + r"""
{Tính dẫn nhiệt.}
{Tính cứng.}
\loigiai{""" + (r"""
Khi đặt một điện trường ngoài (cực âm bên trái, cực dương bên phải), các electron tự do mang điện tích âm chuyển động có hướng về phía cực dương (tạo nên dòng điện) $\rightarrow$ Hình vẽ mô tả tính dẫn điện.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cầu chì là một phần tử hay thiết bị bảo vệ mạch điện bằng cách làm đứt mạch điện. Cầu chì được sử dụng nhằm phòng tránh các hiện tượng quá tải trên đường dây. Dây cầu chì thường được làm bằng kim loại chì (Pb), thiếc trắng (Sn) hoặc cadmium (Cd).
\par\vspace{1pt}\centerline{\includegraphics[height=1.4cm]{images_crop_de18/p1_cau7_cau_chi.png}}\vspace{1pt}
Ứng dụng này dựa trên tính chất nào của các kim loại trên?
\choice
{Có tính dẻo cao.}
""" + (r"{\True Có nhiệt độ nóng chảy tương đối thấp.}" if is_teacher else r"{Có nhiệt độ nóng chảy tương đối thấp.}") + r"""
{Có độ dẫn điện cao.}
{Có độ cứng tương đối thấp.}
\loigiai{""" + (r"""
Chì ($327^\circ\text{C}$), thiếc ($232^\circ\text{C}$), cadmium ($321^\circ\text{C}$) có nhiệt độ nóng chảy tương đối thấp. Khi dòng điện quá tải sinh nhiệt lớn, dây chì nóng chảy đứt trước để ngắt mạch bảo vệ an toàn cho hệ thống điện.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong nước nguyên chất khi phản ứng xảy ra sẽ có quá trình sau:\\
$\ce{2H2O + 2e -> H2 + 2OH-} \quad E^\circ_{\ce{2H2O/H2 + 2OH-}} = -0{,}42\text{ V}$.\\
Cho $E^\circ_{\ce{Na+/Na}} = -2{,}71\text{ V}$; $E^\circ_{\ce{Cu^2+/Cu}} = +0{,}34\text{ V}$; $E^\circ_{\ce{Mg^2+/Mg}} = -2{,}36\text{ V}$; $E^\circ_{\ce{Al^3+/Al}} = -1{,}68\text{ V}$.\\
Hãy cho biết kim loại nào sau đây \textbf{không} phản ứng được với nước ở điều kiện chuẩn.
\choice
""" + (r"{\True Cu.}" if is_teacher else r"{Cu.}") + r"""
{Na.}
{Mg.}
{Al.}
\loigiai{""" + (r"""
Để kim loại có khả năng khử nước ở điều kiện chuẩn tạo khí \ce{H2}, thế điện cực chuẩn của nó phải âm hơn $-0{,}42\text{ V}$. Vì $E^\circ_{\ce{Cu^2+/Cu}} = +0{,}34\text{ V} > -0{,}42\text{ V}$ nên Cu không phản ứng với nước ở điều kiện chuẩn.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các kim loại: Na, Ca, Cu, Ag, K. Số kim loại có khả năng khử được nước ở điều kiện thường giải phóng \ce{H2} là
\choice
{2.}
{1.}
{4.}
""" + (r"{\True 3.}" if is_teacher else r"{3.}") + r"""
\loigiai{""" + (r"""
Có 3 kim loại phản ứng trực tiếp với nước ở điều kiện thường là Na, Ca, K sinh ra dung dịch kiềm và khí \ce{H2}. Cu và Ag trơ với nước.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\begin{ex}
Lấy một đinh sắt đã được làm sạch bề mặt rồi ngâm vào dung dịch \ce{CuSO4} một thời gian. Hiện tượng nào sau đây \textbf{không} xuất hiện trong thí nghiệm trên?
\choice
{Có chất rắn màu đỏ bám lên đinh sắt.}
{Màu xanh của dung dịch nhạt dần.}
""" + (r"{\True Có bọt khí thoát ra mạnh trên bề mặt đinh sắt.}" if is_teacher else r"{Có bọt khí thoát ra mạnh trên bề mặt đinh sắt.}") + r"""
{Đinh sắt bị ăn mòn một phần.}
\loigiai{""" + (r"""
Phản ứng: $\ce{Fe + CuSO4 -> FeSO4 + Cu}$. Không có phản ứng sinh khí, nên hiện tượng bọt khí thoát ra mạnh là không xuất hiện.\\
\textbf{Chọn C.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Nguyên tắc chung để điều chế kim loại từ hợp chất của chúng là
\choice
""" + (r"{\True khử cation kim loại thành kim loại.}" if is_teacher else r"{khử cation kim loại thành kim loại.}") + r"""
{oxi hoá ion kim loại thành kim loại.}
{oxi hoá kim loại thành ion kim loại.}
{khử kim loại thành ion kim loại.}
\loigiai{""" + (r"""
Trong hợp chất, kim loại ở dạng cation $\ce{M^{n+}}$. Nguyên tắc điều chế kim loại là khử cation kim loại thành kim loại tự do: $\ce{M^{n+} + ne -> M}$.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Gang và thép là hai hợp kim quan trọng nhất của sắt. Chúng có rất nhiều ứng dụng trong đời sống và sản xuất. Để sản xuất gang, người ta khử oxide của sắt (\ce{Fe2O3} hoặc \ce{Fe3O4}) trong quặng sắt bằng khí CO (được tạo ra do đốt cháy than cốc). Sắt được tạo ra trong quá trình sản xuất gang bằng phương pháp nào?
\choice
{Kết tinh.}
""" + (r"{\True Nhiệt luyện.}" if is_teacher else r"{Nhiệt luyện.}") + r"""
{Thuỷ luyện.}
{Chiết.}
\loigiai{""" + (r"""
Dùng chất khử CO khử oxide sắt ở nhiệt độ cao trong lò cao thuộc phương pháp nhiệt luyện.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Kim loại kẽm (zinc, Zn) được sản xuất trong công nghiệp từ quặng sphalerite (có thành phần chính là ZnS) theo sơ đồ: $\ce{ZnS} \xrightarrow{+\ce{O2},\, t^\circ} \ce{ZnO} \xrightarrow{+\ce{C},\, t^\circ} \ce{Zn}$. Phương pháp điều chế kim loại nào đã được sử dụng trong quá trình sản xuất zinc theo sơ đồ trên?
\choice
{Thuỷ luyện.}
""" + (r"{\True Nhiệt luyện.}" if is_teacher else r"{Nhiệt luyện.}") + r"""
{Điện phân.}
{Kết tinh.}
\loigiai{""" + (r"""
Giai đoạn khử ZnO bằng than cốc C ở nhiệt độ cao tạo Zn kim loại: $\ce{ZnO + C -> Zn + CO}$ là phương pháp nhiệt luyện.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Vàng (Au) tồn tại trong tự nhiên ở dạng đơn chất. Tuy nhiên, hàm lượng Au trong quặng thường rất thấp nên rất khó tách Au bằng cơ học. Trong công nghiệp, người ta tách vàng từ quặng theo sơ đồ sau:\\
Quặng chứa vàng $(\ce{Au}) \xrightarrow{+\ce{O2} + \ce{KCN} + \ce{H2O}} \ce{K[Au(CN)2]} \text{ (aq)} \xrightarrow{+\ce{Zn dư}} \ce{Au} \text{ (s)}$\\
Phương pháp điều chế kim loại nào đã được sử dụng trong sơ đồ trên?
\choice
""" + (r"{\True Thuỷ luyện.}" if is_teacher else r"{Thuỷ luyện.}") + r"""
{Nhiệt luyện.}
{Điện phân.}
{Chiết.}
\loigiai{""" + (r"""
Dùng dung dịch hoá chất (KCN) hoà tan kim loại tạo phức tan, sau đó dùng kim loại có tính khử mạnh hơn (Zn) đẩy Au ra khỏi dung dịch phức là phương pháp thuỷ luyện.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
\begin{minipage}[t]{0.72\textwidth}
Thực hiện thí nghiệm sau:\\
\textbf{Bước 1:} Cho dung dịch \ce{NaCl} 5\% vào ống chữ U như hình bên.\\
\textbf{Bước 2:} Nhúng thanh Cu và thanh Zn đã làm sạch vào 2 đầu ống.\\
\textbf{Bước 3:} Nối hai thanh kim loại bằng dây dẫn.\\
Sau bước 2, kim loại nào bị ăn mòn?
\end{minipage}\hfill
\begin{minipage}[t]{0.26\textwidth}
\centering
\vspace{-4pt}
\includegraphics[height=1.55cm]{images_crop_de18/p1_cau15_u_tube.png}
\end{minipage}
\choice
{Đồng.}
{Cả hai đều bị ăn mòn.}
{Kẽm.}
""" + (r"{\True Không kim loại nào bị ăn mòn.}" if is_teacher else r"{Không kim loại nào bị ăn mòn.}") + r"""
\loigiai{""" + (r"""
Sau bước 2, hai thanh kim loại chưa được nối bằng dây dẫn nên không hình thành pin điện hoá. Mặt khác, trong dung dịch \ce{NaCl} trung tính, cả Zn và Cu đều không phản ứng hoá học với nước hay \ce{NaCl} ở điều kiện thường $\rightarrow$ Không có kim loại nào bị ăn mòn.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Để khảo sát quá trình ăn mòn kim loại, tiến hành ba thí nghiệm như hình dưới đây. Mỗi ống nghiệm chứa một chiếc đinh sắt trong các môi trường khác nhau. Phát biểu nào dưới đây là đúng?
\par\vspace{1pt}\centerline{\includegraphics[height=1.5cm]{images_crop_de18/p1_cau16_an_mon_ong_nghiem.png}}\vspace{1pt}
\choice
{Sau một thời gian, đinh sắt trong cả ba ống nghiệm đều bị gỉ.}
{Sau một thời gian, đinh sắt trong cả ba ống nghiệm đều không có sự thay đổi.}
{Sau một thời gian, đinh sắt trong ống nghiệm số 1 và 3 bị gỉ.}
""" + (r"{\True Sau một thời gian, chỉ có đinh sắt trong ống nghiệm số 1 bị gỉ.}" if is_teacher else r"{Sau một thời gian, chỉ có đinh sắt trong ống nghiệm số 1 bị gỉ.}") + r"""
\loigiai{""" + (r"""
Ăn mòn sắt trong không khí ẩm đòi hỏi đồng thời nước và oxygen:\\
-- Ống 1: Có cả nước và oxygen hoà tan $\rightarrow$ Đinh sắt bị gỉ.\\
-- Ống 2: Nước cất đun sôi đã đuổi hết khí, phủ lớp dầu ngăn cách không cho $\ce{O2}$ khuếch tán vào $\rightarrow$ Đinh sắt không bị gỉ.\\
-- Ống 3: Không khí khô do có $\ce{CaCl2}$ hút ẩm $\rightarrow$ Không có nước $\rightarrow$ Đinh sắt không bị gỉ.\\
Vậy chỉ có đinh sắt trong ống nghiệm số 1 bị gỉ.\\
\textbf{Chọn D.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Trong số các thí nghiệm sau đây, thí nghiệm nào có tốc độ ăn mòn đinh sắt chậm nhất?
\par\vspace{1pt}\centerline{\includegraphics[height=1.1cm]{images_crop_de18/p1_cau17_bon_ong_nghiem.png}}\vspace{1pt}
\choice
{(a).}
""" + (r"{\True (b).}" if is_teacher else r"{(b).}") + r"""
{(c).}
{(d).}
\loigiai{""" + (r"""
Ống (b) ngâm chìm hoàn toàn trong dầu thực vật, cách li hoàn toàn đinh sắt với nước và không khí ẩm $\rightarrow$ Tốc độ ăn mòn chậm nhất.\\
\textbf{Chọn B.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Tiến hành thí nghiệm của ba kim loại khác nhau với dung dịch hydrochloric acid loãng theo hình vẽ sau:
\par\vspace{1pt}\centerline{\includegraphics[height=1.2cm]{images_crop_de18/p1_cau18_ba_kim_loai_hcl.png}}\vspace{1pt}
Xác định các kim loại phù hợp với hình vẽ (theo thứ tự: Kim loại W, Kim loại X, Kim loại Y)?
\choice
""" + (r"{\True Cu, Mg, Zn.}" if is_teacher else r"{Cu, Mg, Zn.}") + r"""
{Zn, Mg, Cu.}
{Mg, Zn, Cu.}
{Cu, Zn, Mg.}
\loigiai{""" + (r"""
Quan sát lượng bọt khí thoát ra:\\
-- Ống W: Hoàn toàn không có bọt khí thoát ra $\rightarrow$ W là kim loại không tác dụng với HCl loãng (Cu đứng sau H).\\
-- Ống X: Bọt khí thoát ra rất mạnh và mãnh liệt $\rightarrow$ X là Mg (kim loại hoạt động hoá học mạnh nhất trong 3 kim loại).\\
-- Ống Y: Bọt khí thoát ra với lượng vừa phải $\rightarrow$ Y là Zn.\\
Vậy các kim loại tương ứng là Cu, Mg, Zn.\\
\textbf{Chọn A.}""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN II (4,0 điểm -- 4 câu): Câu trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}

\setcounter{ex}{0}

\begin{ex}
Kim loại Cu có độ dẫn điện kém hơn Ag, kim loại Al có độ dẫn điện kém hơn Au nhưng dây dẫn điện chủ yếu làm bằng kim loại Cu và Al mà không phải là Ag và Au. Ứng dụng làm dây dẫn điện của Cu và Al là do chúng
\choiceTF
""" + (r"{\True đều dẫn điện tốt.}" if is_teacher else r"{đều dẫn điện tốt.}") + r"""
""" + (r"{\True có giá thành phù hợp với thị trường tiêu dùng.}" if is_teacher else r"{có giá thành phù hợp với thị trường tiêu dùng.}") + r"""
""" + (r"{\True sản xuất được với quy mô công nghiệp.}" if is_teacher else r"{sản xuất được với quy mô công nghiệp.}") + r"""
{đều có khối lượng riêng nhỏ.}
\loigiai{""" + (r"""
a) \textbf{Đúng}. Cu dẫn điện tốt thứ 2 (sau Ag) và Al dẫn điện tốt thứ 4 (sau Ag, Cu, Au).\\
b) \textbf{Đúng}. Giá thành Cu và Al rẻ hơn rất nhiều so với Ag và Au, phù hợp cho tiêu dùng đại trà.\\
c) \textbf{Đúng}. Quặng Cu và Al phong phú, công nghệ luyện kim phát triển cho phép sản xuất quy mô công nghiệp hàng triệu tấn mỗi năm.\\
d) \textbf{Sai}. Chỉ có Al có khối lượng riêng nhỏ ($D = 2{,}7\text{ g/cm}^3 < 5\text{ g/cm}^3$), còn Cu là kim loại nặng ($D = 8{,}96\text{ g/cm}^3 > 5\text{ g/cm}^3$).""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Tiến hành thí nghiệm theo các bước sau:\\
\textbf{Bước 1.} Cho vào 3 ống nghiệm, mỗi ống 2 mL dung dịch \ce{H2SO4} 1 M.\\
\textbf{Bước 2.} Cho 3 lá kim loại có kích thước như nhau gồm lá nhôm (Al) đã làm sạch lớp bề mặt vào ống nghiệm (1), lá sắt (iron, Fe) vào ống nghiệm (2) và lá đồng (Cu) vào ống nghiệm (3).\\
Biết: $E^\circ_{\ce{Al^3+/Al}} = -1{,}676\text{ V}$; $E^\circ_{\ce{Fe^2+/Fe}} = -0{,}440\text{ V}$; $E^\circ_{\ce{Cu^2+/Cu}} = +0{,}340\text{ V}$.
\choiceTF
{Ở bước 2, ở cả ba ống nghiệm đều có khí thoát ra.}
""" + (r"{\True Tốc độ thoát khí ở ống (1) nhanh hơn ống (2).}" if is_teacher else r"{Tốc độ thoát khí ở ống (1) nhanh hơn ống (2).}") + r"""
{Nếu thay \ce{H2SO4} loãng bằng \ce{H2SO4} đặc thì hiện tượng ở bước 2 sẽ không đổi.}
{Ở bước 2, nếu thêm tiếp 2 mL \ce{H2SO4} 1 M vào cả 3 ống thì tốc độ thoát khí ở cả ba ống sẽ tăng.}
\loigiai{""" + (r"""
a) \textbf{Sai}. $E^\circ_{\ce{Cu^2+/Cu}} = +0{,}340\text{ V} > 0\text{ V}$ nên Cu không tác dụng với \ce{H2SO4} loãng, ống 3 không có khí thoát ra.\\
b) \textbf{Đúng}. Al có tính khử mạnh hơn Fe ($E^\circ_{\ce{Al^3+/Al}} = -1{,}676\text{ V} < E^\circ_{\ce{Fe^2+/Fe}} = -0{,}440\text{ V}$) nên tốc độ thoát khí ở ống (1) nhanh hơn ống (2).\\
c) \textbf{Sai}. Nếu thay bằng \ce{H2SO4} đặc nguội thì Al và Fe bị thụ động hoá; đặc nóng thì cả ba kim loại đều tác dụng giải phóng khí \ce{SO2} mùi hắc thay vì khí \ce{H2}.\\
d) \textbf{Sai}. Thêm dung dịch cùng nồng độ 1 M thì $[\ce{H+}]$ không đổi nên tốc độ phản ứng không đổi; hơn nữa Cu ở ống 3 vẫn không phản ứng.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Nhôm phế liệu thường lẫn các tạp chất là các chất hữu cơ và vô cơ (có trong nhãn, mác do in hoặc sơn). Khi tái chế, phế liệu được cắt, băm nhỏ rồi cho vào lò nung đến khi chảy lỏng. Phần lớn các tạp chất này biến thành xỉ lỏng, nổi lên và được vớt ra khỏi lò. Phần còn lại là nhôm tái chế ở trạng thái lỏng và được đưa vào đúc khuôn.
\choiceTF
""" + (r"{\True Việc tái chế nhôm giúp giảm giá thành sản phẩm.}" if is_teacher else r"{Việc tái chế nhôm giúp giảm giá thành sản phẩm.}") + r"""
""" + (r"{\True Việc tái chế nhôm giúp giảm chất thải ra môi trường.}" if is_teacher else r"{Việc tái chế nhôm giúp giảm chất thải ra môi trường.}") + r"""
""" + (r"{\True Không nên dùng nhôm tái chế để chế tạo dụng cụ nhà bếp và y tế.}" if is_teacher else r"{Không nên dùng nhôm tái chế để chế tạo dụng cụ nhà bếp và y tế.}") + r"""
{Đem cắt, băm nhỏ nhôm phế liệu để quá trình khử \ce{Al2O3} xảy ra dễ hơn.}
\loigiai{""" + (r"""
a) \textbf{Đúng}. Tái chế nhôm tiết kiệm khoảng 95\% năng lượng so với sản xuất từ quặng bauxite, giảm giá thành.\\
b) \textbf{Đúng}. Giảm thiểu rác thải rắn phế liệu và giảm lượng khí thải ra môi trường.\\
c) \textbf{Đúng}. Nhôm tái chế khó loại bỏ hoàn toàn các tạp chất kim loại nặng độc hại (chì, cadmium...) nên không an toàn cho đồ dùng nhà bếp và thiết bị y tế.\\
d) \textbf{Sai}. Quá trình tái chế nhôm phế liệu là quá trình nấu chảy vật lí để loại bỏ xỉ và đúc lại nhôm, không phải phản ứng hoá học khử \ce{Al2O3}.""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Tiến hành thí nghiệm theo các bước sau:\\
\textbf{Bước 1.} Cho vào cốc thuỷ tinh 30 mL dung dịch \ce{H2SO4} 0,5 M.\\
\textbf{Bước 2.} Cho một lá nhôm và một lá đồng vào cốc sao cho chúng không tiếp xúc với nhau. Để yên khoảng 1 phút.\\
\textbf{Bước 3.} Dùng dây dẫn điện nối lá nhôm và lá đồng với một vôn kế (như hình bên).\\
Biết: $E^\circ_{\ce{Al^3+/Al}} = -1{,}676\text{ V}$; $E^\circ_{\ce{2H+/H2}} = 0\text{ V}$; $E^\circ_{\ce{Cu^2+/Cu}} = +0{,}340\text{ V}$.
\par\vspace{1pt}\centerline{\includegraphics[height=2.2cm]{images_crop_de18/p2_cau4_pin_al_cu.png}}\vspace{1pt}
\choiceTF
""" + (r"{\True Ở bước 2, lá nhôm bị ăn mòn hoá học còn lá đồng không bị ăn mòn.}" if is_teacher else r"{Ở bước 2, lá nhôm bị ăn mòn hoá học còn lá đồng không bị ăn mòn.}") + r"""
{Ở bước 3, có sự tạo thành cặp pin điện hoá, trong đó lá nhôm là cathode và lá đồng là anode.}
{Ở bước 3, khí chỉ thoát ra ở bề mặt lá đồng, còn ở lá nhôm không có khí thoát ra.}
{Ở bước 3, kim vôn kế chỉ 1,336 V.}
\loigiai{""" + (r"""
a) \textbf{Đúng}. Ở bước 2 chưa nối dây, Al phản ứng trực tiếp với \ce{H2SO4} (ăn mòn hoá học), Cu không phản ứng nên không bị ăn mòn.\\
b) \textbf{Sai}. Al có tính khử mạnh hơn nên đóng vai trò là cực âm (anode), Cu là cực dương (cathode).\\
c) \textbf{Sai}. Khi nối dây, ở cathode (Cu) xảy ra sự khử \ce{H+} sinh khí \ce{H2}; đồng thời trên bề mặt Al phản ứng hoá học trực tiếp giữa Al và acid vẫn tiếp tục xảy ra song song, nên khí \ce{H2} vẫn thoát ra ở cả lá nhôm.\\
d) \textbf{Sai}. Sức điện động chuẩn của pin: $E^\circ_{\text{pin}} = E^\circ_{\text{cathode}} - E^\circ_{\text{anode}} = 0 - (-1{,}676) = 1{,}676\text{ V}$ (hoặc nếu đo cặp Al -- Cu là $0{,}340 - (-1{,}676) = 2{,}016\text{ V}$), không phải $1{,}336\text{ V}$.""" if is_teacher else "") + r"""}
\end{ex}

""" + (r"\newpage" if not is_teacher else "") + r"""

\noindent
\textbf{PHẦN III (1,5 điểm -- 6 câu): Câu trắc nghiệm yêu cầu trả lời ngắn.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 6.}

\setcounter{ex}{0}

\begin{ex}
Quá trình sản xuất gang từ nguyên liệu quặng hematite, than cốc, chất chảy trong lò cao xảy ra các phản ứng chính:
\par\vspace{1pt}{\centering
\footnotesize\renewcommand{\arraystretch}{0.88}
\begin{tabular}{|c|l||c|l|}
\hline
\textbf{Nhiệt độ} & \textbf{Phương trình phản ứng} & \textbf{Nhiệt độ} & \textbf{Phương trình phản ứng} \\
\hline
$400^\circ\text{C}$ & $\ce{3Fe2O3 + CO ->[t^\circ] 2Fe3O4 + CO2}$ & $1300^\circ\text{C}$ & $\ce{CaO + SiO2 ->[t^\circ] CaSiO3}$ \\
$500 - 600^\circ\text{C}$ & $\ce{Fe3O4 + CO ->[t^\circ] 3FeO + CO2}$ & $1500^\circ\text{C}$ & $\ce{C + CO2 ->[t^\circ] 2CO}$ \\
$700 - 800^\circ\text{C}$ & $\ce{FeO + CO ->[t^\circ] Fe + CO2}$ & $1800^\circ\text{C}$ & $\ce{C + O2 ->[t^\circ] CO2}$ \\
$1000^\circ\text{C}$ & $\ce{CaCO3 ->[t^\circ] CaO + CO2}$ & & \\
\hline
\end{tabular}\par}\vspace{1pt}
Trong các phản ứng trên có bao nhiêu phản ứng oxi hoá -- khử?
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}5}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Các phản ứng oxi hoá -- khử (có sự thay đổi số oxi hoá của các nguyên tố):\\
(1) $\ce{3Fe2O3 + CO -> 2Fe3O4 + CO2}$ (thay đổi số oxi hoá của Fe và C).\\
(2) $\ce{Fe3O4 + CO -> 3FeO + CO2}$ (thay đổi số oxi hoá của Fe và C).\\
(3) $\ce{FeO + CO -> Fe + CO2}$ (thay đổi số oxi hoá của Fe và C).\\
(4) $\ce{C + CO2 -> 2CO}$ (thay đổi số oxi hoá của C).\\
(5) $\ce{C + O2 -> CO2}$ (thay đổi số oxi hoá của C và O).\\
Hai phản ứng phân huỷ $\ce{CaCO3}$ và tạo xỉ $\ce{CaSiO3}$ không thay đổi số oxi hoá.\\
Tổng cộng có 5 phản ứng oxi hoá -- khử.\\
\textbf{Đáp số: 5.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Để bảo vệ vật bằng thép khỏi bị ăn mòn điện hoá, trong thực tế người ta đã thực hiện một số cách sau:\\
(1) Sơn kín bề mặt cánh cửa làm bằng thép.\\
(2) Tráng kẽm lên tấm thép mỏng khi sản xuất tôn.\\
(3) Gắn một số tấm kẽm lên vỏ tàu làm bằng thép.\\
(4) Cho dầu mỡ lên các ốc vít trên đường ray.\\
Liệt kê những cách sử dụng phương pháp bảo vệ bề mặt theo số thứ tự tăng dần.
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}1 2 4}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
-- Phương pháp bảo vệ bề mặt: Phủ lớp màng cách li kim loại với môi trường (sơn, tráng kẽm, bôi dầu mỡ) $\rightarrow$ gồm (1), (2), (4).\\
-- Cách (3): Gắn tấm kẽm là phương pháp bảo vệ điện hoá (dùng kim loại hi sinh Zn).\\
Vậy các cách bảo vệ bề mặt là: 1 2 4.\\
\textbf{Đáp số: 1 2 4.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Cho các tính chất vật lí ở cột (A) và giải thích ở cột (B) như sau:
\par\vspace{1pt}{\centering
\footnotesize\renewcommand{\arraystretch}{0.88}
\begin{tabular}{|l|p{12.2cm}|}
\hline
\textbf{Cột A} & \textbf{Cột B (Giải thích)} \\
\hline
(a) Tính dẻo & (1) Dưới tác dụng của điện trường, các electron tự do chuyển động có hướng. \\
(b) Tính dẫn điện & (2) Các electron tự do phản xạ ánh sáng nhìn thấy. \\
(c) Tính dẫn nhiệt & (3) Do các electron tự do liên kết các ion kim loại khi bị biến dạng cơ học. \\
(d) Tính ánh kim & (4) Các electron ở nhiệt độ cao truyền năng lượng cho ion dương và electron khác. \\
\hline
\end{tabular}\par}\vspace{1pt}
Hãy sắp xếp thứ tự giải thích của cột B với thứ tự tính chất vật lí ở cột A cho phù hợp.
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}3142}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Ghép tương ứng:\\
-- (a) Tính dẻo $\leftrightarrow$ (3) Electron tự do liên kết các ion kim loại khi bị biến dạng cơ học.\\
-- (b) Tính dẫn điện $\leftrightarrow$ (1) Electron tự do chuyển động có hướng trong điện trường.\\
-- (c) Tính dẫn nhiệt $\leftrightarrow$ (4) Electron vùng nhiệt cao truyền năng lượng nhiệt cho phần còn lại.\\
-- (d) Tính ánh kim $\leftrightarrow$ (2) Electron tự do phản xạ ánh sáng nhìn thấy.\\
Thứ tự ghép đôi với (a)-(b)-(c)-(d) là: 3142.\\
\textbf{Đáp số: 3142.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Nhôm (Al) là kim loại phổ biến nhất trong vỏ Trái Đất. Nhôm được sử dụng để chế tạo hợp kim nhẹ, bền, được dùng trong kĩ thuật hàng không, trong xây dựng và làm nội thất. Nhôm có số hiệu nguyên tử là 13 và $E^\circ_{\ce{Al^3+/Al}} = -1{,}676\text{ V}$. Cho các phát biểu sau:\\
(1) Nhôm và đồ vật làm bằng nhôm dễ bị ăn mòn trong không khí ẩm.\\
(2) Ở trạng thái cơ bản, nguyên tử Al có 1 electron ở lớp ngoài cùng.\\
(3) Nhôm phản ứng được với nước ngay ở điều kiện chuẩn.\\
(4) Nhôm có tính khử mạnh hơn Mg (biết $E^\circ_{\ce{Mg^2+/Mg}} = -2{,}356\text{ V}$).\\
(5) Trong vỏ Trái Đất, nhôm chỉ tồn tại ở dạng hợp chất.\\
Liệt kê các phát biểu đúng theo số thứ tự tăng dần.
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}3 5}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
(1) Sai vì nhôm có màng oxide $\ce{Al2O3}$ mỏng, mịn và rất bền bảo vệ nên không bị ăn mòn trong không khí ẩm.\\
(2) Sai vì Al ($[\ce{Ne}]\ce{3s^2 3p^1}$) có 3 electron ở lớp ngoài cùng.\\
(3) Đúng vì $E^\circ_{\ce{Al^3+/Al}} = -1{,}676\text{ V} < -0{,}42\text{ V}$ nên về mặt nhiệt động học Al phản ứng được với nước ở điều kiện chuẩn (nhưng nhanh chóng dừng lại do tạo màng $\ce{Al(OH)3}$ không tan bảo vệ).\\
(4) Sai vì $E^\circ_{\ce{Mg^2+/Mg}} = -2{,}356\text{ V} < E^\circ_{\ce{Al^3+/Al}} = -1{,}676\text{ V}$ nên Mg có tính khử mạnh hơn Al.\\
(5) Đúng vì nhôm hoạt động hoá học mạnh nên chỉ tồn tại ở dạng hợp chất.\\
Vậy các phát biểu đúng là: 3 5.\\
\textbf{Đáp số: 3 5.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Để mạ đồng một vật dụng kim loại có diện tích $10\text{ cm}^2$, người ta điện phân dung dịch \ce{CuSO4} với cực âm là vật mạ và cực dương là lá đồng thô ($I = 2\text{ A}$, hiệu suất điện phân $90\%$, $D_{\ce{Cu}} = 8{,}94\text{ g/cm}^3$). Khối lượng đồng tính theo công thức Faraday: $m = \frac{A \cdot I \cdot t}{n \cdot F}$ ($A_{\ce{Cu}} = 64$, $F = 96485\text{ C/mol}$, $n = 2$). Thời gian điện phân để lớp mạ có độ dày $0{,}1\text{ mm}$ là bao nhiêu phút? (Làm tròn đến hàng đơn vị)
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}25}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Đổi độ dày: $d = 0{,}1\text{ mm} = 0{,}01\text{ cm}$.\\
Thể tích lớp mạ: $V = S \cdot d = 10 \cdot 0{,}01 = 0{,}1\text{ cm}^3$.\\
Khối lượng đồng cần mạ: $m = V \cdot D = 0{,}1 \cdot 8{,}94 = 0{,}894\text{ g}$.\\
Theo định luật Faraday với hiệu suất $H = 90\%$:
$$m = \frac{A \cdot I \cdot t}{n \cdot F} \cdot \frac{H}{100\%} \Rightarrow 0{,}894 = \frac{64 \cdot 2 \cdot t}{2 \cdot 96485} \cdot 0{,}9 \Rightarrow t \approx 1497{,}53\text{ giây.}$$
Đổi sang phút: $t = \frac{1497{,}53}{60} \approx 25\text{ phút}$.\\
\textbf{Đáp số: 25.}""" if is_teacher else "") + r"""}
\end{ex}

\begin{ex}
Đồng thau là hợp kim của đồng và kẽm. Để xác định hàm lượng Cu trong đồng thau, người ta cho 10 gam mảnh đồng thau vào dung dịch \ce{H2SO4} loãng dư; sau phản ứng hoàn toàn, lọc lấy chất rắn làm khô cân được 6,5 gam chất rắn. Hàm lượng đồng trong loại đồng thau trên bằng bao nhiêu \%? (Làm tròn đến hàng phần mười)
""" + (r"""\par\noindent\textbf{Đáp số:} \textbf{\color{red}65,0}""" if is_teacher else r"""\par\noindent\textit{Trả lời:} \dotfill""") + r"""
\loigiai{""" + (r"""
Khi cho đồng thau (Cu, Zn) vào dung dịch \ce{H2SO4} loãng dư, chỉ có Zn phản ứng:
$$\ce{Zn + H2SO4 -> ZnSO4 + H2 ^}$$
Kim loại Cu đứng sau H trong dãy điện hoá nên không tác dụng với \ce{H2SO4} loãng.\\
Chất rắn không tan thu được chính là Cu: $m_{\ce{Cu}} = 6{,}5\text{ gam}$.\\
Hàm lượng đồng trong loại đồng thau là:
$$\%\ce{Cu} = \frac{6{,}5}{10} \cdot 100\% = 65{,}0\%.$$
\textbf{Đáp số: 65,0.}""" if is_teacher else "") + r"""}
\end{ex}

"""

    if is_teacher:
        tex += r"""
\vspace{0.3cm}
\noindent
\textbf{\color{hfRed}\large BẢNG ĐÁP ÁN THAM KHẢO TOÀN ĐỀ}

\vspace{0.1cm}
\noindent
\textbf{\color{mainGreen}PHẦN I:} 1.D \quad 2.A \quad 3.D \quad 4.B \quad 5.C \quad 6.B \quad 7.B \quad 8.A \quad 9.D \quad 10.C \quad 11.A \quad 12.B \quad 13.B \quad 14.A \quad 15.D \quad 16.D \quad 17.B \quad 18.A

\vspace{0.1cm}
\noindent
\textbf{\color{mainGreen}PHẦN II:} Câu 1: a-Đ, b-Đ, c-Đ, d-S \quad|\quad Câu 2: a-S, b-Đ, c-S, d-S\\
Câu 3: a-Đ, b-Đ, c-Đ, d-S \quad|\quad Câu 4: a-Đ, b-S, c-S, d-S

\vspace{0.1cm}
\noindent
\textbf{\color{mainGreen}PHẦN III:} Câu 1: \textbf{5} \quad|\quad Câu 2: \textbf{1 2 4} \quad|\quad Câu 3: \textbf{3142} \quad|\quad Câu 4: \textbf{3 5} \quad|\quad Câu 5: \textbf{25} \quad|\quad Câu 6: \textbf{65,0}
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
    print("--- Đang tạo file LaTeX và PDF cho ĐỀ SỐ 18 ---")
    
    # 1. Bản Học Sinh
    tex_hs = create_latex_de18(is_teacher=False)
    pages_hs = compile_latex(tex_hs, "De_Kiem_Tra_Hoa_12_De18_HocSinh.pdf".replace(".pdf", ".tex"), "De_Kiem_Tra_Hoa_12_De18_HocSinh.pdf")
    
    # 2. Bản Giáo Viên
    tex_gv = create_latex_de18(is_teacher=True)
    pages_gv = compile_latex(tex_gv, "De_Kiem_Tra_Hoa_12_De18_GiaoVien.pdf".replace(".pdf", ".tex"), "De_Kiem_Tra_Hoa_12_De18_GiaoVien.pdf")
    
    print("\n--- KẾT QUẢ BIÊN DỊCH ---")
    print(f"Học sinh: {pages_hs} trang (Yêu cầu chuẩn: 4 trang)")
    print(f"Giáo viên: {pages_gv} trang")
    
    if pages_hs == 4:
        print("XÁC NHẬN: Bản Học Sinh ĐẠT CHUẨN CHÍNH XÁC 4 TRANG A4!")
    else:
        print(f"CẢNH BÁO: Bản Học Sinh có {pages_hs} trang, cần điều chỉnh!")

if __name__ == "__main__":
    main()
