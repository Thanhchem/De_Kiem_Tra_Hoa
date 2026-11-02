# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Giữa Kì 1 Hóa Học 10 - Mã Đề 5001
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
Sinh cả 2 phiên bản:
1. Bản Học sinh: De_Kiem_Tra_Hoa_10_Ma5001_HocSinh.pdf
2. Bản Giáo viên: De_Kiem_Tra_Hoa_10_Ma5001_GiaoVien.pdf
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

SAN_PHAM_DIR = r"C:\Antigravity_Thanh\San_Pham"

def build_latex_content(is_teacher=False):
    doc_type = "gv" if is_teacher else "hs"
    
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
\usepackage{tikz}
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
\setlength{\parskip}{1.5pt}

\begin{document}

\begin{minipage}[t]{0.45\textwidth}
\textbf{SỞ GIÁO DỤC VÀ ĐÀO TẠO}\\
\textbf{TRƯỜNG THPT}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề gồm 4 phần, 4 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.52\textwidth}
\raggedleft
\textbf{\color{hfRed}KIỂM TRA GIỮA KỲ 1 -- NĂM HỌC 2026 -- 2027}\\
\textbf{MÔN: HÓA HỌC -- KHỐI 10}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{MÃ ĐỀ: 5001}
\end{minipage}

\vspace{3pt}
\noindent
\textit{Họ, tên thí sinh:} \dotfill \textit{Lớp:} \dotfill \textit{SBD:} \dotfill

\vspace{2pt}
\noindent
\textit{\footnotesize (Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; P = 31; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Ga = 69,7; In = 114,8).}
""")

    if is_teacher:
        tex.append(r"""
\vspace{2pt}
\begin{center}
\textbf{\color{hfRed}\large (BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)}
\end{center}
""")

    tex.append(r"""
\vspace{2pt}
\noindent
\textbf{\color{mainGreen}PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Học sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi học sinh chỉ chọn một phương án. (3,0 điểm)}

\Opensolutionfile{ans}[ans-5001-""" + doc_type + r"""]

\begin{ex}
Ứng dụng nào sau đây thể hiện vai trò của hóa học trong lĩnh vực bảo vệ môi trường?
\choice
{Tổng hợp thuốc kháng sinh để điều trị bệnh nhiễm khuẩn.}
{\True Chế tạo vật liệu hấp phụ để loại bỏ ion kim loại nặng trong nước thải nhà máy dệt.}
{Sản xuất xi măng phục vụ xây dựng nhà ở.}
{Sản xuất phân đạm để tăng năng suất cây trồng.}
""" + (r"""\loigiai{
- A thuộc lĩnh vực y dược, chăm sóc sức khỏe.\\
- B: Chế tạo vật liệu hấp phụ để loại bỏ ion kim loại nặng trong nước thải công nghiệp là giải pháp trực tiếp trong bảo vệ môi trường.\\
- C thuộc ngành vật liệu xây dựng; D thuộc ngành nông nghiệp.\\
\textbf{Chọn B.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Năm 2026, một nhóm nghiên cứu khoa học đã tận dụng phụ phẩm của cây cải dầu (bã ép, thân, lá, hoa) để chế tạo màng bao gói thực phẩm phân hủy sinh học. Màng được tạo từ chitosan, chiết xuất phenolic và hạt nano bạc, giúp kéo dài thời gian bảo quản rau quả tươi. Hoạt động nào sau đây thuộc phạm vi nghiên cứu của hóa học?
\choice
{Điều tra thói quen và mức độ sẵn sàng sử dụng bao gói xanh của người tiêu dùng.}
{Khảo sát chi phí sản xuất và khả năng thương mại hóa màng bao gói trên thị trường.}
{Nghiên cứu đặc điểm sinh trưởng của cây cải dầu nhằm nâng cao năng suất thu hoạch.}
{\True Xác định thành phần, cấu trúc của màng và sự biến đổi của màng khi bị phân hủy trong đất.}
""" + (r"""\loigiai{
- Đối tượng nghiên cứu của hóa học là chất và sự biến đổi của chất. Hoạt động xác định thành phần, cấu trúc của màng và sự biến đổi của màng khi bị phân hủy trong đất thuộc phạm vi nghiên cứu của hóa học.\\
- Các phương án A, B thuộc lĩnh vực kinh tế - xã hội học; C thuộc lĩnh vực nông nghiệp - sinh học.\\
\textbf{Chọn D.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Hình vẽ sau mô tả thí nghiệm tìm ra hạt X là một trong những thành phần cấu tạo nên nguyên tử.
\begin{center}
\includegraphics[width=7cm]{images_gk1_10_5001/cau3_ong_phong_tia.png}
\end{center}
Hạt X là
\choice
{proton.}
{neutron.}
{\True electron.}
{neutron và proton.}
""" + (r"""\loigiai{
Sơ đồ mô tả thí nghiệm ống phóng tia âm cực của J.J. Thomson. Chùm tia lệch về phía bản cực dương của điện trường, chứng tỏ mang điện tích âm, cấu tạo từ các hạt electron ($e$).\\
\textbf{Chọn C.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Các nguyên tố trong bảng tuần hoàn hiện nay được xếp theo chiều tăng dần của
\choice
{\True điện tích hạt nhân nguyên tử.}
{nguyên tử khối.}
{số neutron.}
{số lớp electron.}
""" + (r"""\loigiai{
Các nguyên tố trong bảng tuần hoàn hiện đại được sắp xếp theo chiều tăng dần của điện tích hạt nhân nguyên tử ($Z$).\\
\textbf{Chọn A.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Năm 1911, Rutherford cho chùm hạt $\alpha$ bắn vào một lá vàng rất mỏng và quan sát thấy hầu hết các hạt $\alpha$ đi thẳng qua lá vàng. Kết quả này chứng tỏ
\choice
{hạt nhân nguyên tử mang điện tích âm.}
{nguyên tử không chứa hạt mang điện.}
{electron có khối lượng rất lớn.}
{\True nguyên tử có cấu tạo rỗng.}
""" + (r"""\loigiai{
Hầu hết các hạt $\alpha$ đi thẳng xuyên qua lá vàng chứng tỏ nguyên tử có cấu tạo rỗng (phần lớn thể tích nguyên tử là khoảng trống giữa hạt nhân và lớp vỏ electron).\\
\textbf{Chọn D.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Hạt nhân nguyên tử chứa 28 hạt neutron là
\choice
{\ce{^{32}_{16}S}.}
{\True \ce{^{54}_{26}Fe}.}
{\ce{^{58}_{28}Ni}.}
{\ce{^{32}_{15}P}.}
""" + (r"""\loigiai{
Số neutron $N = A - Z$:\\
+ Với \ce{^{54}_{26}Fe}: $N = 54 - 26 = 28$.\\
\textbf{Chọn B.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Nguyên tố hóa học là những nguyên tử có cùng
\choice
{số khối.}
{số neutron.}
{số neutron và số proton.}
{\True số proton.}
""" + (r"""\loigiai{
Nguyên tố hóa học là tập hợp các nguyên tử có cùng điện tích hạt nhân (có cùng số proton).\\
\textbf{Chọn D.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Phát biểu nào sau đây không đúng?
\choice
{Phân lớp p có nhiều orbital hơn phân lớp s.}
{Số electron tối đa trên phân lớp d gấp năm lần số electron tối đa trên phân lớp s.}
{Số lượng electron tối đa trong một phân lớp luôn là một số chẵn.}
{\True Electron càng ở xa hạt nhân thì có năng lượng càng thấp.}
""" + (r"""\loigiai{
- Electron càng ở xa hạt nhân thì lực hút giữa hạt nhân và electron càng yếu, năng lượng của electron càng cao. Do đó phát biểu D sai.\\
\textbf{Chọn D.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Hình dưới biểu diễn hình dạng orbital nguyên tử (AO) nào sau đây?
\begin{center}
\includegraphics[width=3.8cm]{images_gk1_10_5001/cau9_orbital_s.png}
\end{center}
\choice
{$p_x$.}
{$p_z$.}
{\True s.}
{$p_y$.}
""" + (r"""\loigiai{
Orbital s có dạng hình cầu đối xứng tâm trong không gian quanh hạt nhân.\\
\textbf{Chọn C.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Phát biểu nào sau đây đúng về chu kì 3 trong bảng tuần hoàn các nguyên tố hóa học?
\choice
{bắt đầu từ nguyên tố có $Z = 19$ và kết thúc ở nguyên tố có $Z = 36$.}
{bắt đầu từ nguyên tố có $Z = 11$ và kết thúc ở nguyên tố có $Z = 19$.}
{\True bắt đầu từ nguyên tố có $Z = 11$ và kết thúc ở nguyên tố có $Z = 18$.}
{bắt đầu từ nguyên tố có $Z = 3$ và kết thúc ở nguyên tố có $Z = 10$.}
""" + (r"""\loigiai{
Chu kì 3 gồm 8 nguyên tố, bắt đầu từ Sodium (\ce{Na}, $Z = 11$) và kết thúc ở Argon (\ce{Ar}, $Z = 18$).\\
\textbf{Chọn C.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Cho mô hình cấu tạo nguyên tử như sau:
\begin{center}
\includegraphics[width=5.5cm]{images_gk1_10_5001/cau11_mo_hinh_nguyen_tu.png}
\end{center}
Kí hiệu đúng của nguyên tử này là
\begin{center}
\includegraphics[width=13.5cm]{images_gk1_10_5001/cau11_o_nguyen_to.png}
\end{center}
\choice
{\True Ô số 6 (Carbon -- C).}
{Ô số 9 (Fluorine -- F).}
{Ô số 12 (Magnesium -- Mg).}
{Ô số 11 (Sodium -- Na).}
""" + (r"""\loigiai{
Mô hình có 6 proton (hạt đen) và 6 electron (lớp 1 có 2e, lớp 2 có 4e). Số hiệu $Z = 6$ là nguyên tố Carbon (\ce{C}).\\
\textbf{Chọn A.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Hợp kim cobalt được dùng rộng rãi trong động cơ máy bay vì có độ bền cao ở nhiệt độ cao. Nguyên tử cobalt có cấu hình electron lớp ngoài cùng là $3d^74s^2$. Ở trạng thái cơ bản, số electron độc thân trong nguyên tử cobalt là
\choice
{5.}
{2.}
{7.}
{\True 3.}
""" + (r"""\loigiai{
Phân lớp $4s^2$ đã bão hòa (0 e độc thân). Phân lớp $3d^7$ có 5 orbital, khi điền 7 electron sẽ có 2 orbital chứa cặp e ghép đôi và 3 orbital chứa 1 e độc thân. Vậy số e độc thân là 3.\\
\textbf{Chọn D.}}""" if is_teacher else "") + r"""
\end{ex}

\Closesolutionfile{ans}

\vspace{5pt}
\noindent
\textbf{\color{mainGreen}PHẦN II. Câu hỏi trắc nghiệm đúng sai.}\\
\textit{Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu thí sinh chọn đúng hoặc sai. (2,0 điểm)}

\begin{ex}
Đồng thau là hợp kim của hai kim loại, được dùng để làm van nước, ổ khóa, tay nắm cửa và các nhạc cụ kèn đồng nhờ bền, dễ gia công và ít bị ăn mòn. Kết quả phân tích phổ khối lượng một mẫu đồng thau cho thấy mẫu chứa bốn loại nguyên tử chính X, Y, Z, T với thông tin như sau:
\begin{center}
\begin{tabular}{|c|c|c|c|}
\hline
\textbf{Nguyên tử} & \textbf{Số proton} & \textbf{Số neutron} & \textbf{Tỉ lệ số nguyên tử trong tự nhiên (\%)} \\
\hline
X & 29 & 34 & 69,15 \\
\hline
Y & 30 & 36 & 27,7 \\
\hline
Z & 29 & 36 & 30,85 \\
\hline
T & 30 & 34 & 49,2 \\
\hline
\end{tabular}
\end{center}
\begin{center}
\includegraphics[width=6.5cm]{images_gk1_10_5001/p2_cau1_dong_thau.png}
\end{center}
\choiceTF
{X và T có cùng số neutron nên là đồng vị của nhau}
{\True Bốn loại nguyên tử trong bảng thông tin trên thuộc hai nguyên tố hóa học}
{\True Nguyên tử khối trung bình của nguyên tố tạo nên X là 63,6 (làm tròn kết quả đến hàng phần mười)}
{Nguyên tử khối trung bình của nguyên tố tạo nên T được tính bằng biểu thức $\dfrac{64 \times 49,2 + 66 \times 27,7}{100}$}
""" + (r"""\loigiai{
\textbf{a) Sai.} Đồng vị là các nguyên tử có cùng số proton. X ($Z = 29$) và T ($Z = 30$) khác $Z$ nên không phải là đồng vị.\\
\textbf{b) Đúng.} Có 2 giá trị $Z$: $Z = 29$ (\ce{Cu}) và $Z = 30$ (\ce{Zn}), thuộc 2 nguyên tố hóa học.\\
\textbf{c) Đúng.} Nguyên tố X là \ce{Cu}, gồm đồng vị X ($A = 63$, 69,15\%) và Z ($A = 65$, 30,85\%):\\
$\overline{A} = \dfrac{63 \times 69,15 + 65 \times 30,85}{100} = 63,617 \approx 63,6$.\\
\textbf{d) Sai.} Nguyên tố T là \ce{Zn}, 2 đồng vị T ($A = 64$, 49,2\%) và Y ($A = 66$, 27,7\%) có tổng tỉ lệ là $76,9\% \neq 100\%$ (kẽm còn có các đồng vị khác). Biểu thức chia cho 100 là sai.}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Một học sinh hỏi chatbot AI: ``Mô hình nguyên tử hiện đại khác mô hình Rutherford -- Bohr như thế nào?'' Chatbot đưa ra các ý sau về mô hình nguyên tử hiện đại:
\begin{center}
\includegraphics[width=6.5cm]{images_gk1_10_5001/p2_cau2_chatbot_ai.png}
\end{center}
\choiceTF
{Trong mô hình hiện đại, electron chuyển động trên các quỹ đạo tròn có bán kính xác định quanh hạt nhân}
{\True Orbital nguyên tử là vùng không gian xung quanh hạt nhân mà xác suất tìm thấy electron khoảng 90\%}
{Mọi orbital đều có dạng hình cầu, chỉ khác nhau về kích thước}
{\True Mỗi orbital chứa tối đa 2 electron và hai electron này có chiều tự quay ngược nhau}
""" + (r"""\loigiai{
\textbf{a) Sai.} Quỹ đạo tròn có bán kính xác định là mô hình Rutherford -- Bohr. Mô hình hiện đại xem electron chuyển động không theo quỹ đạo xác định.\\
\textbf{b) Đúng.} Định nghĩa orbital nguyên tử (AO) theo SGK.\\
\textbf{c) Sai.} Chỉ có AO s có dạng hình cầu; AO p có dạng số 8 nổi.\\
\textbf{d) Đúng.} Theo nguyên lí Pauli, mỗi AO chứa tối đa 2 electron có spin đối song.}""" if is_teacher else "") + r"""
\end{ex}

\vspace{5pt}
\noindent
\textbf{\color{mainGreen}PHẦN III. Trắc nghiệm trả lời ngắn.}\\
\textit{Học sinh trả lời từ câu 1 đến câu 4. (2,0 điểm)}

\begin{ex}
Khối lượng của một nguyên tử magnesium là $39,8271 \cdot 10^{-27}\text{ kg}$. Biết $1\text{ amu} = 1,6605 \cdot 10^{-27}\text{ kg}$. Khối lượng nguyên tử magnesium tính theo đơn vị amu là bao nhiêu? \textit{(Kết quả làm tròn đến hàng đơn vị.)}
""" + (r"""\loigiai{
$m = \dfrac{39,8271 \cdot 10^{-27}}{1,6605 \cdot 10^{-27}} \approx 23,985\text{ amu} \approx 24\text{ amu}$.\\
\textbf{Đáp số: 24.}}""" if is_teacher else r"""\newline\textbf{Trả lời:} \dotfill""") + r"""
\end{ex}

\begin{ex}
Cho các phát biểu sau:\\
(1) Điện tích của proton và electron có cùng độ lớn nhưng ngược dấu.\\
(2) Có những nguyên tử không chứa neutron nào.\\
(3) Khối lượng nguyên tử hầu hết tập trung ở hạt nhân nguyên tử.\\
(4) Điện tích của proton và neutron có cùng độ lớn nhưng ngược dấu.\\
(5) Khối lượng của proton và neutron xấp xỉ bằng nhau và lớn hơn nhiều khối lượng của electron.\\
Hãy sắp xếp các phát biểu \textbf{không đúng} trong các phát biểu trên theo số thứ tự tăng dần (ví dụ 12, 134, \dots)
""" + (r"""\loigiai{
- (1), (2), (3), (5) đều đúng.\\
- (4) Sai vì neutron không mang điện (điện tích bằng 0).\\
Phát biểu không đúng duy nhất là (4).\\
\textbf{Đáp số: 4.}}""" if is_teacher else r"""\newline\textbf{Trả lời:} \dotfill""") + r"""
\end{ex}

\begin{ex}
Nguyên tử của nguyên tố X có tổng số electron trên các phân lớp p là 11. Số hiệu nguyên tử của X là bao nhiêu?
""" + (r"""\loigiai{
Cấu hình e: $1s^2 2s^2 2p^6 3s^2 3p^5$. Tổng số e trên phân lớp p: $6 + 5 = 11$.\\
Tổng số electron = $2 + 2 + 6 + 2 + 5 = 17 \implies Z = 17$ (\ce{Cl}).\\
\textbf{Đáp số: 17.}}""" if is_teacher else r"""\newline\textbf{Trả lời:} \dotfill""") + r"""
\end{ex}

\begin{ex}
Dựa vào bảng tuần hoàn các nguyên tố hóa học, cho thông tin về nguyên tố X như sau:
\begin{center}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Kí hiệu} & \textbf{Số hiệu nguyên tử} & \textbf{Cấu hình e lớp ngoài cùng} & \textbf{Chu kì} & \textbf{Nhóm} \\
\hline
X & 12 & $3s^2$ & 3 & IIA \\
\hline
\end{tabular}
\end{center}
Nguyên tố Y thuộc cùng chu kì với X và thuộc nhóm VIIA. X và Y kết hợp với nhau tạo thành hợp chất có công thức \ce{XY2}. Tổng số proton có trong một phân tử \ce{XY2} là bao nhiêu?
""" + (r"""\loigiai{
- X có $Z_X = 12$ (\ce{Mg}).\\
- Y cùng chu kì 3, nhóm VIIA $\implies$ cấu hình e lớp ngoài là $3s^23p^5 \implies Z_Y = 17$ (\ce{Cl}).\\
- Phân tử \ce{XY2} (\ce{MgCl2}) có: $\sum p = Z_X + 2 Z_Y = 12 + 2 \times 17 = 46$.\\
\textbf{Đáp số: 46.}}""" if is_teacher else r"""\newline\textbf{Trả lời:} \dotfill""") + r"""
\end{ex}

\vspace{5pt}
\noindent
\textbf{\color{mainGreen}PHẦN IV. TỰ LUẬN.}\\
\textit{HỌC SINH LÀM RA GIẤY THI TRẢ LỜI TỪ CÂU 1 ĐẾN CÂU 3 (3,0 điểm)}

\begin{ex}
\textbf{(0,5 điểm)} Nguyên tố chlorine (\ce{Cl}) có nguyên tử khối trung bình là 35,5. Biết chlorine gồm hai đồng vị bền là \ce{^{35}Cl} và \ce{^{37}Cl}. Để xác định phần trăm các đồng vị bền của chlorine trong tự nhiên, người ta dùng phương pháp phổ khối lượng.\\
\textbf{a)} Hãy xác định phần trăm mỗi đồng vị bền của chlorine. \textit{(0,25 điểm)}\\
\textbf{b)} Hãy cho biết hình ảnh phổ khối lượng nào sau đây của chlorine là phù hợp nhất? Vì sao? \textit{(0,25 điểm)}
\begin{center}
\includegraphics[width=13.5cm]{images_gk1_10_5001/p4_cau1_pho_khoi_luong.png}
\end{center}
""" + (r"""\loigiai{
\textbf{a)} Gọi $x$ là \% số nguyên tử của \ce{^{35}Cl}, \% của \ce{^{37}Cl} là $(100 - x)$.\\
$\overline{A} = \dfrac{35x + 37(100 - x)}{100} = 35,5 \iff 3700 - 2x = 3550 \iff x = 75\%$.\\
Vậy \ce{^{35}Cl} chiếm 75\% và \ce{^{37}Cl} chiếm 25\%.\\
\textbf{b)} Hình ảnh phổ khối lượng phù hợp nhất là \textbf{Hình 1 (A)}.\\
Vì tỉ lệ cường độ pic ứng với tỉ lệ số nguyên tử các đồng vị: \ce{^{35}Cl} : \ce{^{37}Cl} = 75\% : 25\% = 3 : 1. Trong Hình 1 (A), pic tại $m/z = 35$ đạt 100\% và pic tại $m/z = 37$ đạt khoảng 33,3\%, hoàn toàn phù hợp với tỉ lệ 3 : 1.}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
\textbf{(0,5 điểm)} Các đám mây gây hiện tượng sấm sét được tạo nên bởi những hạt nước nhỏ li ti mang điện tích. Một phép đo thực nghiệm cho thấy một giọt nước hình cầu có đường kính $5 \cdot 10^{-3}\text{ cm}$ mang điện tích âm là $-3{,}220 \cdot 10^{-17}\text{ C}$. Điện tích âm này là do giọt nước nhận thêm electron.\\
Số phân tử nước trong giọt nước được tính theo biểu thức:
\[
N_{\ce{H2O}} = \dfrac{4}{3} \cdot 3{,}14 \cdot r^3 \cdot \dfrac{D}{18} \cdot N_A
\]
Trong đó $D\text{ (g/cm}^3\text{)}$ là khối lượng riêng của nước, $r\text{ (cm)}$ là bán kính của giọt nước. Biết khối lượng riêng của nước là $1\text{ g/cm}^3$; điện tích của một electron là $-1{,}602 \cdot 10^{-19}\text{ C}$; số Avogadro $N_A = 6{,}022 \cdot 10^{23}\text{ mol}^{-1}$.\\
\textbf{a)} Tính số phân tử nước trong giọt nước. \textit{(0,25 điểm)}\\
\textbf{b)} Trung bình cứ $x \cdot 10^{12}$ phân tử nước có trong giọt nước thì có 1 hạt electron dư. Tính giá trị của $x$ \textit{(Kết quả làm tròn đến hàng đơn vị).} \textit{(0,25 điểm)}
""" + (r"""\loigiai{
\textbf{a)} Bán kính giọt nước: $r = \dfrac{5 \cdot 10^{-3}}{2} = 2{,}5 \cdot 10^{-3}\text{ cm}$.\\
Số phân tử nước:\\
$N_{\ce{H2O}} = \dfrac{4}{3} \cdot 3{,}14 \cdot (2{,}5 \cdot 10^{-3})^3 \cdot \dfrac{1}{18} \cdot 6{,}022 \cdot 10^{23} \approx 2{,}1886 \cdot 10^{15}\text{ (phân tử)}.\\
\textbf{b)} Số electron dư:\\
$n_e = \dfrac{-3{,}220 \cdot 10^{-17}}{-1{,}602 \cdot 10^{-19}} \approx 201\text{ (hạt)}.\\
Tỉ số phân tử nước / electron dư:\\
$\dfrac{N_{\ce{H2O}}}{n_e} = \dfrac{2{,}1886 \cdot 10^{15}}{201} \approx 1{,}0888 \cdot 10^{13} = 10{,}888 \cdot 10^{12}\text{ (phân tử)}.\\
$\implies x \approx 10{,}888$. Làm tròn đến hàng đơn vị được $x = 11$.}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
\textbf{(2,0 điểm)} Năm 2019 được chọn là Năm Quốc tế Bảng tuần hoàn các nguyên tố hóa học, kỉ niệm 150 năm Dmitri Ivanovich Mendeleev công bố bảng tuần hoàn (1869). Khi xây dựng bảng, Mendeleev đã để trống một số ô và dự đoán tính chất của các nguyên tố chưa được phát hiện. Một trong số đó là ``eka-aluminium'', nguyên tố nằm ngay dưới nhôm trong cùng một nhóm.\\
Năm 1875, nhà hóa học người Pháp Paul-Émile Lecoq de Boisbaudran công bố tìm ra một nguyên tố mới. Nguyên tố này có nhiều tính chất trùng khớp với eka-aluminium, nhưng khối lượng riêng ông đo được lại khác xa giá trị Mendeleev tiên đoán. Mendeleev đề nghị ông đo lại. Kết quả đo lần sau, với mẫu đã được tinh chế, trùng khớp với dự đoán của Mendeleev. Một cách đơn giản để xác định khối lượng riêng của một mẫu kim loại là dùng bình tràn, như minh họa ở Hình 1.
\begin{center}
\includegraphics[width=8cm]{images_gk1_10_5001/p4_cau3_binh_tran.png}\\
\textit{\footnotesize Hình 1. Minh họa phương pháp đo khối lượng riêng bằng bình tràn (số liệu trong hình chỉ mang tính minh họa).}
\end{center}
Áp dụng phương pháp này để mô phỏng tái hiện lại phép đo trong lịch sử, một nhóm học sinh cân một mẫu kim loại M (ở $20^\circ\text{C}$) được khối lượng \textbf{35,46 g}. Sau đó, các em thả nhẹ mẫu vào một bình tràn chứa đầy nước và thu được \textbf{6,00 mL} nước tràn ra ống đong. Cho bảng số liệu sau:
\begin{center}
\begin{tabular}{|l|c|c|c|c|c|}
\hline
\textbf{Kim loại} & \textbf{Mg} & \textbf{Al} & \textbf{Cu} & \textbf{Zn} & \textbf{Ga} \\
\hline
Số hiệu nguyên tử (Z) & 12 & 13 & 29 & 30 & 31 \\
\hline
Khối lượng riêng ($\text{g/cm}^3$) & 1,74 & 2,70 & 8,96 & 7,13 & 5,91 \\
\hline
\end{tabular}
\end{center}
\textbf{a)} Tính khối lượng riêng của mẫu kim loại M. Dựa vào bảng số liệu, cho biết M là kim loại nào? \textit{(0,25 điểm)}\\
\textbf{b)} Viết cấu hình electron của nguyên tử M. \textit{(0,25 điểm)}\\
\textbf{c)} Xác định vị trí của M trong bảng tuần hoàn (ô, chu kì, nhóm). \textit{(0,5 điểm)}\\
\textbf{d)} M là nguyên tố s, p, d hay f? \textit{(0,25 điểm)}\\
\textbf{e)} Nguyên tử M có bao nhiêu electron hóa trị? \textit{(0,25 điểm)}\\
\textbf{f)} Dựa vào vị trí của M và Al trong bảng tuần hoàn, hãy cho biết tính chất hoá học của M và Al có tương tự nhau hay không? Vì sao? \textit{(0,25 điểm)}\\
\textbf{g)} Biết nguyên tố In ($Z = 49$) nằm cùng nhóm với nguyên tố nhôm, ngay phía dưới nguyên tố ``eka-aluminium''. Hãy xác định số electron hoá trị của In? \textit{(0,25 điểm)}
""" + (r"""\loigiai{
\textbf{a)} Thể tích mẫu: $V = 6{,}00\text{ mL} = 6{,}00\text{ cm}^3$.\\
Khối lượng riêng: $D = \dfrac{m}{V} = \dfrac{35{,}46}{6{,}00} = 5{,}91\text{ g/cm}^3$.\\
Tra bảng số liệu: $D = 5{,}91\text{ g/cm}^3$ ứng với kim loại Gallium (\ce{Ga}, $Z = 31$). Vậy M là kim loại Gallium (\ce{Ga}).\\
\textbf{b)} Cấu hình electron của \ce{Ga} ($Z = 31$): $1s^2 2s^2 2p^6 3s^2 3p^6 3d^{10} 4s^2 4p^1$ (hoặc $[\ce{Ar}] 3d^{10} 4s^2 4p^1$).\\
\textbf{c)} Vị trí của \ce{Ga} trong bảng tuần hoàn:\\
- Ô số: 31 (vì $Z = 31$).\\
- Chu kì: 4 (vì có 4 lớp electron).\\
- Nhóm: IIIA (vì là nguyên tố p, có 3 e lớp ngoài cùng: $4s^2 4p^1$).\\
\textbf{d)} \ce{Ga} là nguyên tố p (vì electron cuối cùng điền vào phân lớp 4p).\\
\textbf{e)} Nguyên tử \ce{Ga} có 3 electron hóa trị (ở phân lớp $4s^2 4p^1$).\\
\textbf{f)} Tính chất hóa học của M (\ce{Ga}) và \ce{Al} tương tự nhau.\\
Vì sao: Vì \ce{Ga} và \ce{Al} cùng thuộc nhóm IIIA, đều có 3 electron hóa trị lớp ngoài cùng ($ns^2 np^1$), dễ nhường 3e khi tham gia phản ứng hóa học để đạt cấu hình electron bền khí hiếm, thể hiện số oxi hóa đặc trưng là $+3$.\\
\textbf{g)} Nguyên tố In ($Z = 49$) thuộc cùng nhóm IIIA với \ce{Al} và \ce{Ga}. Các nguyên tố trong cùng nhóm A có số e hóa trị bằng số thứ tự của nhóm. Do đó In có 3 electron hóa trị (cấu hình: $[\ce{Kr}] 4d^{10} 5s^2 5p^1$).}""" if is_teacher else "") + r"""
\end{ex}

\begin{center}
\textbf{--- HẾT ---}
\end{center}

\end{document}
""")

    return "".join(tex)

def compile_latex(tex_path, output_dir):
    try:
        cmd = ["pdflatex", "-interaction=nonstopmode", os.path.basename(tex_path)]
        res = subprocess.run(cmd, cwd=output_dir, capture_output=True, text=True, encoding="utf-8", errors="replace")
        # Chạy lần 2 để cập nhật số trang và references
        subprocess.run(cmd, cwd=output_dir, capture_output=True, text=True, encoding="utf-8", errors="replace")
        if res.returncode == 0:
            print(f"Biên dịch thành công: {tex_path}")
            return True
        else:
            print(f"Lỗi biên dịch {tex_path}: {res.stderr[:300]}")
            return False
    except Exception as e:
        print(f"Exception compiling {tex_path}: {e}")
        return False

def cleanup_temp_files(output_dir):
    extensions = ["*.aux", "*.log", "*.thm", "ans-*.tex", "*.out"]
    for ext in extensions:
        for f in glob.glob(os.path.join(output_dir, ext)):
            try:
                os.remove(f)
            except Exception:
                pass

if __name__ == "__main__":
    os.makedirs(SAN_PHAM_DIR, exist_ok=True)
    
    # 1. Bản Học Sinh
    hs_tex_file = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_Hoa_10_Ma5001_HocSinh.tex")
    with open(hs_tex_file, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))
    print(f"Saved: {hs_tex_file}")
    compile_latex(hs_tex_file, SAN_PHAM_DIR)
    
    # 2. Bản Giáo Viên
    gv_tex_file = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_Hoa_10_Ma5001_GiaoVien.tex")
    with open(gv_tex_file, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))
    print(f"Saved: {gv_tex_file}")
    compile_latex(gv_tex_file, SAN_PHAM_DIR)
    
    # Dọn dẹp file tạm
    cleanup_temp_files(SAN_PHAM_DIR)
    print("Hoàn tất biên dịch PDF và dọn dẹp file rác.")
