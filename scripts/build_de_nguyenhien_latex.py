# -*- coding: utf-8 -*-
"""
Script tạo file LaTeX và biên dịch PDF cho Đề Kiểm Tra Giữa Kì 1 - Hóa 12
Trường THPT Nguyễn Hiền (TP. Hồ Chí Minh) - Năm học 2026 - 2027
Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)
"""

import os
import sys
import subprocess

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
\textbf{SỞ GD\&ĐT TP. HỒ CHÍ MINH}\\
\textbf{TRƯỜNG T.H.P.T NGUYỄN HIỀN}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề gồm 3 phần, 4 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.52\textwidth}
\raggedleft
\textbf{\color{hfRed}ĐỀ KIỂM TRA GIỮA KÌ 1. NĂM HỌC 2026 -- 2027}\\
\textbf{MÔN: HOÁ HỌC -- KHỐI 12}\\
\textit{Thời gian làm bài: 45 phút (không kể thời gian phát đề)}\\
\textbf{MÃ ĐỀ: Gốc}
\end{minipage}

\vspace{3pt}
\noindent
\textit{Họ, tên thí sinh:} \dotfill \textit{Lớp:} \dotfill \textit{SBD:} \dotfill

\vspace{2pt}
\noindent
\textit{\footnotesize (Cho biết nguyên tử khối: H = 1, C = 12, N = 14, O = 16, S = 32, Cl = 35,5, Na = 23, K = 39, Ca = 40, Br = 80).}
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
\textbf{\color{mainGreen}PHẦN I (4,5 điểm). Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.}\\
\textit{Mỗi câu trả lời đúng thí sinh được 0,25 điểm.}

\Opensolutionfile{ans}[ans-nh-""" + doc_type + r"""]

\begin{ex}
Methyl acetate có công thức là
\choice
{\ce{HCOOCH2CH3}.}
{\True \ce{CH3COOCH3}.}
{\ce{CH3COOCH2CH3}.}
{\ce{CH2=CHCOOCH3}.}
""" + (r"""\loigiai{
Methyl acetate có công thức cấu tạo là \ce{CH3COOCH3}.\\
\textbf{Chọn B.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Thủy phân tristearin \ce{(C17H35COO)3C3H5} trong dung dịch \ce{NaOH}, thu được muối ứng dụng làm xà phòng có công thức là
\choice
{\ce{C3H5COONa}.}
{\True \ce{C17H35COONa}.}
{\ce{C3H5(OH)3}.}
{\ce{C17H33COONa}.}
""" + (r"""\loigiai{
Phương trình xà phòng hóa tristearin:\\
\ce{(C17H35COO)3C3H5 + 3NaOH -> 3C17H35COONa + C3H5(OH)3}.\\
Muối sodium stearate (\ce{C17H35COONa}) là thành phần làm xà phòng.\\
\textbf{Chọn B.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Nhỏ dung dịch methylamine dư vào dung dịch nào sau đây thấy xuất hiện kết tủa, sau đó kết tủa tan dần và tạo dung dịch phức chất có màu xanh?
\choice
{Dung dịch \ce{FeCl3}.}
{Dung dịch \ce{H2SO4}.}
{\True Dung dịch \ce{CuSO4}.}
{Dung dịch \ce{HCl}.}
""" + (r"""\loigiai{
Methylamine có tính base tương tự ammonia:\\
\ce{2CH3NH2 + CuSO4 + 2H2O -> Cu(OH)2 v + (CH3NH3)2SO4}.\\
Khi methylamine dư, kết tủa \ce{Cu(OH)2} tan tạo phức xanh lam:\\
\ce{Cu(OH)2 + 4CH3NH2 -> [Cu(CH3NH2)4](OH)2}.\\
\textbf{Chọn C.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Ester nào sau đây khi đun nóng với dung dịch \ce{NaOH} không thu được sản phẩm alcohol?
\choice
{\True \ce{CH3COOCH=CH-CH3}.}
{\ce{CH2=CHCOOCH2-CH3}.}
{\ce{CH3COOCH2-CH3}.}
{\ce{CH3COOCH2-CH=CH2}.}
""" + (r"""\loigiai{
\ce{CH3COOCH=CH-CH3} có nhóm vinyl dạng enol không bền, thủy phân tạo aldehyde propanal:\\
\ce{CH3COOCH=CH-CH3 + NaOH -> CH3COONa + CH3CH2CHO}.\\
\textbf{Chọn A.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Chất nào sau đây tác dụng với bromine thu được kết tủa màu trắng?
\choice
{glucose.}
{ethyl amine.}
{triolein.}
{\True aniline.}
""" + (r"""\loigiai{
Aniline phản ứng với nước bromine tạo kết tủa trắng 2,4,6-tribromoaniline:\\
\ce{C6H5NH2 + 3Br2 -> C6H2Br3NH2 v + 3HBr}.\\
\textbf{Chọn D.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Nhiệt độ sôi của amine phụ thuộc vào khối lượng phân tử và khả năng tạo liên kết hydrogen liên phân tử. Amine nào sau đây có nhiệt độ sôi cao nhất?
\choice
{\ce{(CH3)3N}.}
{\ce{CH3NHCH2CH3}.}
{\True \ce{CH3CH2CH2NH2}.}
{\ce{CH3CH2NH2}.}
""" + (r"""\loigiai{
\ce{CH3CH2CH2NH2} là amine bậc một có 2 liên kết N--H nên tạo liên kết hydrogen liên phân tử mạnh nhất, nhiệt độ sôi cao nhất ($\sim 48\,^\circ\text{C}$).\\
\textbf{Chọn C.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Một mắt xích của cellulose có cấu trúc như hình bên:
\begin{center}
\includegraphics[width=5.5cm]{images_gk1nguyenhien/cau7_cellulose.png}
\end{center}
Nitrate hoá cellulose bằng tác nhân \ce{HNO3} (xúc tác \ce{H2SO4}) thấy tạo thành cellulose trinitrate. Các nguyên tử oxygen ở vị trí nào (trong mỗi mắt xích) đã tham gia phản ứng?
\choice
{\True Vị trí (1), (3), (5).}
{Vị trí (2), (3), (5).}
{Vị trí (2), (3), (4).}
{Vị trí (3), (4), (5).}
""" + (r"""\loigiai{
Ba nhóm \ce{-OH} tự do ở các vị trí (1) tại C2, (5) tại C3, và (3) tại C6 phản ứng este hóa với \ce{HNO3}. Các vị trí (2) và (4) là cầu ether/glycoside không tham gia phản ứng.\\
\textbf{Chọn A.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Phát biểu nào sau đây về xà phòng và chất giặt rửa không đúng?
\choice
{Chất giặt rửa tổng hợp có thể được sản xuất từ các sản phẩm của dầu mỏ.}
{Xà phòng là hỗn hợp muối sodium hoặc potassium của acid béo và một số chất phụ gia.}
{\True Xà phòng giặt rửa tốt trong nước cứng vì muối calcium, magnesium của acid béo tan tốt trong nước.}
{Phân tử chất giặt rửa gồm phần ưa nước và phần kị nước (ưa dầu mỡ).}
""" + (r"""\loigiai{
C sai vì muối calcium và magnesium của acid béo kết tủa không tan trong nước cứng, làm mất tác dụng giặt rửa của xà phòng.\\
\textbf{Chọn C.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Khi nghiên cứu về phản ứng điều chế ethyl acetate từ ethanol $90^\circ$ và acetic acid nguyên chất, một nhóm học sinh muốn khảo sát ảnh hưởng của lượng \ce{H2SO4} đặc đến hiệu suất điều chế ester. Để thực hiện điều này, nhóm học sinh đã xây dựng kế hoạch cho ba thí nghiệm điều chế, trong đó có liệt kê một số yếu tố sau:\\
(1) Nhiệt độ khi thực hiện phản ứng.\\
(2) Thể tích dung dịch \ce{H2SO4} 98\%.\\
(3) Thể tích ethanol và acetic acid ban đầu.\\
(4) Thời gian thực hiện phản ứng.\\
Những yếu tố nào cần được thiết lập giống nhau ở cả ba thí nghiệm?
\choice
{\True (1), (3), (4).}
{(1), (2), (4).}
{(1), (2), (3).}
{(2), (3), (4).}
""" + (r"""\loigiai{
Yếu tố khảo sát là lượng \ce{H2SO4} đặc (2) nên phải thay đổi. Do đó các yếu tố kiểm soát (1), (3), (4) cần giữ cố định giống nhau.\\
\textbf{Chọn A.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Trong quá trình sản xuất bia bằng phương pháp lên men sinh học, dưới tác dụng của enzyme sẽ xảy ra quá trình chuyển hoá sau: $\text{X} \longrightarrow \text{maltose} \longrightarrow \text{Y}$. Chất X, Y tương ứng là
\choice
{tinh bột và fructose.}
{cellulose và glucose.}
{cellulose và fructose.}
{\True tinh bột và glucose.}
""" + (r"""\loigiai{
Tinh bột (X) bị enzyme amylase thủy phân thành maltose, sau đó enzyme maltase thủy phân maltose thành glucose (Y).\\
\textbf{Chọn D.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Chất nào sau đây không thuộc nhóm hợp chất carbohydrate?
\choice
{Fructose.}
{Tinh bột.}
{\True Glucosamine.}
{Glucose.}
""" + (r"""\loigiai{
Glucosamine (\ce{C6H13NO5}) là một dẫn xuất amino-sugar chứa nitrogen, không phải carbohydrate thuần túy.\\
\textbf{Chọn C.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Sodium stearate (\ce{C17H35COONa}) có khả năng làm sạch vết dầu mỡ bám trên vải là do
\choice
{sodium stearate phản ứng hóa học với dầu mỡ tạo thành chất tan trong nước.}
{gốc \ce{C17H35-} ưa nước, còn nhóm \ce{-COONa} kị nước nên kéo vết dầu mỡ vào nước.}
{sodium stearate làm tăng sức căng bề mặt của nước nên nước dễ thấm vào vải.}
{\True gốc \ce{C17H35-} thâm nhập vào vết dầu mỡ, nhóm \ce{-COO^-} ưa nước hướng ra ngoài, giúp vết bẩn phân tán thành các hạt nhỏ trong nước.}
""" + (r"""\loigiai{
Phân tử có đầu ưa nước (\ce{-COO^-}) và đuôi kị nước (\ce{C17H35-}) thâm nhập vào dầu mỡ, nhũ tương hóa vết bẩn vào nước.\\
\textbf{Chọn D.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Tính chất của nhóm -OH hemiacetal trong glucose được thể hiện khi phản ứng với chất nào sau đây?
\choice
{\ce{[Ag(NH3)2]OH} ($t^\circ$).}
{\True \ce{CH3OH} (xúc tác \ce{HCl} khan).}
{\ce{Cu(OH)2/OH^-} ($t^\circ$).}
{\ce{Cu(OH)2}.}
""" + (r"""\loigiai{
Nhóm -OH hemiacetal ở C1 phản ứng chọn lọc với \ce{CH3OH} (xúc tác \ce{HCl} khan) tạo methyl glucoside.\\
\textbf{Chọn B.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Cấu tạo dạng mạch vòng của hai disaccharide: lactose và saccharose được biểu diễn như hình bên:
\begin{center}
\includegraphics[width=9.0cm]{images_gk1nguyenhien/cau14_disaccharide.png}
\end{center}
Phát biểu nào sau đây không đúng?
\choice
{Lactose và saccharose là đồng phân cấu tạo của nhau.}
{Lactose và saccharose đều có khả năng tham gia phản ứng thủy phân trong môi trường acid.}
{\True Trong phân tử lactose, các gốc glucose liên kết với nhau qua liên kết $\alpha$-1,4-glycoside.}
{Ở điều kiện thường, hai dung dịch lactose và saccharose đều hòa tan được \ce{Cu(OH)2} trong môi trường kiềm.}
""" + (r"""\loigiai{
Lactose gồm 1 gốc $\beta$-galactose và 1 gốc $\alpha$-glucose liên kết qua liên kết $\beta$-1,4-glycoside. Vì vậy C sai.\\
\textbf{Chọn C.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Đặc điểm nào sau đây không đúng đối với methylamine?
\choice
{Methylamine có công thức phân tử \ce{CH5N}.}
{\True Methylamine là amine bậc hai.}
{Methylamine làm quỳ tím ẩm đổi thành màu xanh.}
{Methylamine phản ứng được với dung dịch \ce{HCl}.}
""" + (r"""\loigiai{
Methylamine (\ce{CH3NH2}) là amine bậc một (chỉ có 1 gốc hydrocarbon liên kết với nguyên tử N).\\
\textbf{Chọn B.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Propyl tiglate là một loại chất tạo nên mùi hương của quả táo, có công thức cấu tạo như sau:
\begin{center}
\includegraphics[width=5.0cm]{images_gk1nguyenhien/cau16_propyl_tiglate.png}
\end{center}
Phát biểu nào sau đây không đúng?
\choice
{\True Thủy phân propyl tiglate trong dung dịch \ce{NaOH} thu được ethanol.}
{Propyl tiglate có 2 liên kết pi ($\pi$) trong phân tử.}
{Propyl tiglate là ester không no, đơn chức, mạch hở.}
{Công thức phân tử của propyl tiglate là \ce{C8H14O2}.}
""" + (r"""\loigiai{
Thủy phân trong \ce{NaOH} thu được propan-1-ol (\ce{CH3CH2CH2OH}), không phải ethanol. Vì vậy A sai.\\
\textbf{Chọn A.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Phát biểu nào sau đây về fructose không đúng?
\choice
{\True Fructose làm mất màu nước bromine ở điều kiện thường.}
{Fructose là đồng phân của glucose.}
{Fructose thuộc loại monosaccharide.}
{Fructose phản ứng được với \ce{Cu(OH)2} trong môi trường kiềm ở điều kiện thường.}
""" + (r"""\loigiai{
Fructose là polyhydroxy ketone nên không phản ứng với nước bromine ở điều kiện thường. Vì vậy A sai.\\
\textbf{Chọn A.}}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Trong quá trình sản xuất xà phòng, sau khi đun chất béo với dung dịch \ce{NaOH}, người ta thêm dung dịch \ce{NaCl} bão hòa vào hỗn hợp nhằm mục đích
\choice
{\True làm giảm độ tan của xà phòng, giúp xà phòng tách ra rõ rệt và nổi lên trên.}
{làm xúc tác cho phản ứng xà phòng hóa.}
{làm tăng tốc độ phản ứng xà phòng hóa.}
{trung hòa lượng \ce{NaOH} còn dư.}
""" + (r"""\loigiai{
Dung dịch \ce{NaCl} bão hòa làm giảm độ tan của xà phòng (hiệu ứng ion chung) giúp xà phòng kết tinh và nổi lên trên.\\
\textbf{Chọn A.}}""" if is_teacher else "") + r"""
\end{ex}

\Closesolutionfile{ans}

\vspace{5pt}
\noindent
\textbf{\color{mainGreen}PHẦN II (4,0 điểm). Cấu trúc trắc nghiệm đúng sai. Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}\\
\textit{Điểm tối đa cho 01 câu hỏi là 1,0 điểm (Đúng 1 ý: 0,1đ; 2 ý: 0,25đ; 3 ý: 0,5đ; 4 ý: 1,0đ).}

\Opensolutionfile{ans}[ans-nh-p2-""" + doc_type + r"""]

\begin{ex}
Cellulose và tinh bột đều có công thức chung là \ce{(C6H10O5)n}. Tinh bột gồm hai thành phần là amylose và amylopectin. Thủy phân hoàn toàn cellulose và tinh bột trong môi trường acid đun nóng, theo phương trình hóa học sau:
$$\ce{(C6H10O5)n + nH2O ->[H^+, t^o] nC6H12O6}\quad (1)$$
\choiceTF
{Cellulose có khả năng tạo phức chất màu xanh tím với iodine.}
{\True Phản ứng thủy phân cellulose và tinh bột theo phương trình (1) tạo thành glucose.}
{Vì cellulose và tinh bột đều có công thức chung là \ce{(C6H10O5)n} nên phân tử khối của cellulose và tinh bột đều bằng nhau.}
{Cellulose và amylopectin đều cấu tạo từ các đơn vị glucose liên kết với nhau bằng các liên kết glycoside nên chúng đều có cấu tạo mạch không phân nhánh.}
""" + (r"""\loigiai{
a) Sai: Chỉ tinh bột mới tạo phức xanh tím với iodine.\\
b) Đúng: Đều tạo sản phẩm duy nhất là glucose.\\
c) Sai: Hệ số polime hóa $n$ của cellulose lớn hơn rất nhiều so với tinh bột.\\
d) Sai: Amylopectin có mạch phân nhánh do liên kết $\alpha$-1,6-glycoside.}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Trong kiểm định chất lượng dầu, chỉ số acid (AV) được định nghĩa là số miligam \ce{KOH} cần thiết để trung hòa toàn bộ acid béo tự do có trong 1,00 gam chất béo; chỉ số xà phòng hóa (SV) là số miligam \ce{KOH} cần thiết để xà phòng hóa hoàn toàn triglyceride và đồng thời trung hòa acid béo tự do có trong 1,00 gam chất béo. Một mẫu chất béo (X) gồm triglyceride và một lượng nhỏ acid béo tự do. Triglyceride trong mẫu chất béo (X) có công thức cấu tạo như sau:
\begin{center}
\includegraphics[width=7.5cm]{images_gk1nguyenhien/p2_cau2_triglyceride.png}
\end{center}
Phòng thí nghiệm tiến hành hai thí nghiệm độc lập sau:\\
- Thí nghiệm 1: Hòa tan 20,00 gam chất béo (X) vào hỗn hợp dung môi thích hợp, thêm vài giọt phenolphthalein rồi chuẩn độ bằng dung dịch chuẩn \ce{KOH} 0,10 M. Đến điểm kết thúc chuẩn độ, thể tích dung dịch \ce{KOH} đã tiêu tốn là 6,00 mL. Giả thiết trong thí nghiệm này, liên kết ester trong chất béo không tham gia phản ứng.\\
- Thí nghiệm 2: Đun nóng 40,00 gam chất béo (X) với 200,00 mL dung dịch \ce{KOH} 1,00 M trong cồn đến khi phản ứng xà phòng hóa xảy ra hoàn toàn. Để trung hòa lượng \ce{KOH} còn dư sau phản ứng cần vừa đủ 120,00 mL dung dịch chuẩn \ce{HCl} 0,50 M. Giả thiết trong thí nghiệm này, toàn bộ liên kết ester trong chất béo và nhóm carboxylic acid trong acid béo đều tham gia phản ứng.
\choiceTF
{Trong thí nghiệm 1, dung môi thích hợp để hòa tan mẫu (X) hoàn toàn phải là nước cất.}
{\True Gốc hydrocarbon mạch dài trong triglyceride chỉ gồm hai loại gốc: stearate -- \ce{C17H35} và oleate -- \ce{C17H33}.}
{Từ thí nghiệm 1, chỉ số acid (AV) của mẫu chất béo (X) bằng 3,36.}
{\True Từ thí nghiệm 2, chỉ số xà phòng hoá (SV) của mẫu chất béo (X) bằng 196,00.}
""" + (r"""\loigiai{
a) Sai: Chất béo không tan trong nước cất; cần dung môi cồn - ether trung tính.\\
b) Đúng: Triglyceride gồm 1 gốc stearate và 2 gốc oleate.\\
c) Sai: $n_{\ce{KOH}} = 0{,}10 \times 0{,}006 = 0{,}0006\,\text{mol} \Rightarrow m_{\ce{KOH}} = 33{,}6\,\text{mg}$. Trong 1,00 gam: $\text{AV} = 33{,}6 / 20 = 1{,}68\,\text{mg KOH/g}$.\\
d) Đúng: $n_{\ce{KOH\,pư}} = 0{,}200 \times 1 - 0{,}120 \times 0{,}5 = 0{,}140\,\text{mol}$. Trong 1,00 gam: $\text{SV} = (0{,}140 \times 56 \times 1000) / 40 = 196{,}00\,\text{mg KOH/g}$.}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Trong dung dịch, glucose và fructose đều tồn tại cả dạng mạch hở và mạch vòng. Quá trình chuyển từ mạch vòng sang mạch hở là do trong phân tử hai chất này có nhóm -OH hemiacetal hoặc -OH hemiketal. Hình vẽ dưới đây mô tả cấu tạo của $\alpha$-glucose, $\beta$-fructose và saccharose, trong đó nhóm -OH được khoanh tròn là nhóm -OH hemiacetal của $\alpha$-glucose và nhóm -OH hemiketal của $\beta$-fructose:
\begin{center}
\includegraphics[width=10.0cm]{images_gk1nguyenhien/p2_cau3_hemiacetal.png}
\end{center}
\choiceTF
{\True Đơn vị $\alpha$-glucose và đơn vị $\beta$-fructose trong saccharose liên kết với nhau qua liên kết $\alpha$-1,2-glycoside.}
{\True Do cả glucose và fructose đều không bị thuỷ phân nên chúng thuộc loại monosaccharide.}
{\True Khác với glucose và fructose, saccharose không có dạng mở vòng vì trong phân tử saccharose không chứa nhóm --OH hemiacetal và --OH hemiketal.}
{\True Dung dịch saccharose có tính chất của polyalcohol vì phân tử saccharose chứa nhiều nhóm -OH liền kề nhau.}
""" + (r"""\loigiai{
a) Đúng: Liên kết giữa C1 ($\alpha$-glucose) và C2 ($\beta$-fructose) là $\alpha$-1,2-glycoside.\\
b) Đúng: Cả hai đều là monosaccharide.\\
c) Đúng: Cả hai nhóm hemiacetal/hemiketal đều đã tham gia liên kết glycoside nên không thể mở vòng.\\
d) Đúng: Saccharose có 8 nhóm -OH liền kề nhau nên hòa tan \ce{Cu(OH)2} tạo dung dịch xanh lam.}""" if is_teacher else "") + r"""
\end{ex}

\begin{ex}
Một nhóm học sinh tìm hiểu ảnh hưởng của nồng độ đường đến quá trình lên men lactic, với giả thuyết: ``\textit{Nồng độ glucose ban đầu càng cao thì lượng lactic acid thu được càng nhiều}''. Nhóm đã thực hiện thí nghiệm như sau:\\
Chuẩn bị: các dung dịch glucose 2\%, 5\%, 10\%, 15\%, 20\% (khối lượng/thể tích, tức g/100 mL), chế phẩm vi sinh vật lên men lactic (\textit{Lactobacillus}), 6 bình thủy tinh có nắp, tủ ấm điều nhiệt.\\
Cách làm: Lấy 200 mL mỗi dung dịch glucose cho vào các bình tương ứng (mẫu 1 $\longrightarrow$ mẫu 5). Thêm cùng một lượng chế phẩm vi sinh vật vào mỗi bình, đậy nắp và ủ ở $37\,^\circ\text{C}$. Bình thứ 6 chứa 200 mL glucose 10\% nhưng không thêm vi sinh vật (mẫu trắng). Sau 3 ngày, lọc sạch các mẫu, để ổn định ở nhiệt độ phòng thì thấy thể tích dung dịch trong các mẫu hầu như không thay đổi là 200 mL. Lần lượt lấy 5 mL các mẫu: mẫu trắng; các mẫu 1; 2; 3; 4; 5 rồi pha loãng thành 100 mL và tiến hành chuẩn độ bằng dung dịch \ce{NaOH} 0,1 M với chỉ thị phenolphthalein. Kết quả thu được như bảng sau:
\begin{center}
\begin{tabular}{|l|c|c|c|c|c|c|}
\hline
\textbf{Mẫu} & \textbf{Mẫu trắng} & \textbf{Mẫu 1} & \textbf{Mẫu 2} & \textbf{Mẫu 3} & \textbf{Mẫu 4} & \textbf{Mẫu 5} \\
\hline
$V_{\text{dd NaOH}}$ (mL) & 0,0 & 4,8 & 8,4 & 12,0 & 9,6 & 6,0 \\
\hline
\end{tabular}
\end{center}
Biết rằng trong các mẫu chỉ chứa \ce{C6H12O6}, \ce{CH3CH(OH)COOH} (lactic acid) và \ce{H2O}; glucose chỉ chuyển hóa thành lactic acid.
\choiceTF
{\True Trong thí nghiệm trên, lượng lactic acid được tạo thành nhiều nhất trong một mẫu (200 mL) là 4,32 gam.}
{Nếu thay dung dịch \ce{NaOH} 0,1 M bằng dung dịch \ce{NaOH} 0,2 M thì thể tích \ce{NaOH} dùng để chuẩn độ mẫu 3 sẽ nhỏ hơn 5,0 mL.}
{Từ kết quả thí nghiệm, kết luận giả thuyết của nhóm học sinh trên là đúng.}
{\True Hiệu suất chuyển hóa glucose thành lactic acid ở mẫu 1 cao hơn ở mẫu 3.}
""" + (r"""\loigiai{
a) Đúng: Ở mẫu 3: $n_{\text{acid}} = 0{,}012 \times 0{,}1 \times (200 / 5) = 0{,}048\,\text{mol} \Rightarrow m = 0{,}048 \times 90 = 4{,}32\,\text{gam}$.\\
b) Sai: $V_{\ce{NaOH}} = 1{,}2 \times 10^{-3} / 0{,}2 = 6{,}0\,\text{mL} > 5{,}0\,\text{mL}$.\\
c) Sai: Khi nồng độ đường tăng lên 15\% và 20\% thì lượng acid giảm do áp suất thẩm thấu cao.\\
d) Đúng: Hiệu suất mẫu 1 là 43,2\%, mẫu 3 là 21,6\%.}""" if is_teacher else "") + r"""
\end{ex}

\Closesolutionfile{ans}

\vspace{5pt}
\noindent
\textbf{\color{mainGreen}PHẦN III (1,5 điểm). Thí sinh trả lời từ câu 1 đến câu 6.}\\
\textit{Mỗi câu trả lời đúng thí sinh được 0,25 điểm.}

\begin{ex}
Có bao nhiêu đồng phân cấu tạo của amine bậc một có công thức phân tử \ce{C4H11N}?
""" + (r"""\loigiai{
Các đồng phân amine bậc một (\ce{C4H9NH2}) gồm:\\
(1) butan-1-amine; (2) butan-2-amine; (3) 2-methylpropan-1-amine; (4) 2-methylpropan-2-amine.\\
\textbf{Đáp số: 4.}}""" if is_teacher else r"""\par\textbf{Trả lời:} \dotfill""") + r"""
\end{ex}

\begin{ex}
Cho sơ đồ điều chế amine X từ benzene:
\begin{center}
\includegraphics[width=9.0cm]{images_gk1nguyenhien/p3_cau2_aniline.png}
\end{center}
Phân tử khối của X là bao nhiêu?
""" + (r"""\loigiai{
X là aniline (\ce{C6H5NH2}). Phân tử khối: $M = 12 \times 6 + 5 + 14 + 2 = 93\,\text{g/mol}$.\\
\textbf{Đáp số: 93.}}""" if is_teacher else r"""\par\textbf{Trả lời:} \dotfill""") + r"""
\end{ex}

\begin{ex}
Cho các phản ứng sau:\\
(1) \ce{6nCO2 + 5nH2O ->[chlorophyll] (C6H10O5)n + 6nO2}\\
(2) \ce{C6H12O6 ->[enzyme] 2CH3-CHOH-COOH}\\
(3) \ce{(C6H10O5)n + nH2O ->[H^+, t^o] nC6H12O6}\\
(4) \ce{C6H12O6 ->[enzyme] 2C2H5OH + 2CO2}\\
(5) \ce{(C17H35COO)3C3H5 + 3NaOH ->[t^o] 3C17H35COONa + C3H5(OH)3}\\
(6) \ce{C6H12O6 + 2[Ag(NH3)2]OH ->[t^o] CH2OH(CHOH)4COONH4 + 2Ag + 3NH3 + H2O}\\
Gán số thứ tự phương trình hoá học của các phản ứng theo tên gọi: xà phòng hóa, quang hợp, lên men lactic, tráng gương và sắp xếp theo trình tự thành dãy bốn số (Ví dụ: 1234, 4321, ...)
""" + (r"""\loigiai{
Xà phòng hóa (5) -- Quang hợp (1) -- Lên men lactic (2) -- Tráng gương (6).\\
\textbf{Đáp số: 5126.}}""" if is_teacher else r"""\par\textbf{Trả lời:} \dotfill""") + r"""
\end{ex}

\begin{ex}
Xà phòng giặt kém hiệu quả trong nước cứng vì ion \ce{Ca^2+} kết tủa với anion của acid béo:
$$\ce{Ca^2+ + 2C17H35COO^- -> (C17H35COO)2Ca v}$$
Một gia đình dùng 50 lít nước giếng khoan có nồng độ ion \ce{Ca^2+} là 59 mg/L để giặt quần áo bằng xà phòng bánh chứa 84\% sodium stearate (\ce{C17H35COONa}) về khối lượng. Tính khối lượng (gam) xà phòng bánh bị hao phí do tạo kết tủa với toàn bộ lượng ion \ce{Ca^2+} trong 50 lít nước trên. (Làm tròn kết quả đến hàng phần mười.)
""" + (r"""\loigiai{
$m_{\ce{Ca^2+}} = 59 \times 50 = 2950\,\text{mg} = 2{,}95\,\text{gam} \Rightarrow n_{\ce{Ca^2+}} = 0{,}07375\,\text{mol}$.\\
$n_{\ce{C17H35COONa}} = 2 \times 0{,}07375 = 0{,}1475\,\text{mol} \Rightarrow m_{\text{muối}} = 0{,}1475 \times 306 = 45{,}135\,\text{gam}$.\\
$m_{\text{xà phòng bánh}} = 45{,}135 / 0{,}84 \approx 53{,}732\,\text{gam} \approx 53{,}7\,\text{gam}$.\\
\textbf{Đáp số: 53,7.}}""" if is_teacher else r"""\par\textbf{Trả lời:} \dotfill""") + r"""
\end{ex}

\begin{ex}
Biodiesel được sản xuất từ dầu ăn đã qua sử dụng (dầu thải) bằng phương trình phản ứng tổng quát như sau: \ce{(RCOO)3C3H5 + 3CH3OH -> 3RCOOCH3 + C3H5(OH)3}.\\
Phân tích thành phần của 100 gam biodiesel từ dầu chiên thải có thành phần (coi chỉ gồm 3 ester) như sau:
\begin{center}
\begin{tabular}{|l|c|c|c|}
\hline
\textbf{Chất} & \ce{C15H31COOCH3} & \ce{C17H33COOCH3} & \ce{C17H35COOCH3} \\
\hline
\textbf{Khối lượng (gam)} & 40 & 50 & 10 \\
\hline
\end{tabular}
\end{center}
Hệ số phát thải (EF) là khối lượng \ce{CO2} (gam) tạo thành trên 1 MJ nhiệt lượng tỏa ra. Cho EF của diesel truyền thống là 74,1 g \ce{CO2}/MJ.\\
Biết rằng biodiesel có khối lượng riêng 0,880 kg/L. Đốt cháy hoàn toàn 1 lít biodiesel tỏa ra 32,0 MJ nhiệt lượng. Phần carbon trong gốc acid béo (\ce{R-COO-}) có nguồn gốc từ thực vật; \ce{CO2} sinh ra từ phần này được coi là carbon trung hòa. Riêng carbon trong nhóm \ce{-CH3} của ester có nguồn gốc từ methanol (\ce{CH3OH}) sản xuất từ khí thiên nhiên; \ce{CO2} sinh ra từ quá trình cháy phần này được tính là \ce{CO2} có nguồn gốc hóa thạch.\\
Phần trăm giảm phát thải \ce{CO2} nguồn gốc hóa thạch khi thay diesel bằng biodiesel (cho cùng một nhiệt lượng) được tính theo công thức:
$$\%\Delta = \frac{\text{EF}_{\text{diesel}} - \text{EF}_{\text{hóa thạch}}}{\text{EF}_{\text{diesel}}} \times 100\%$$
Tính giá trị $\%\Delta$ của mẫu biodiesel trên. (Chỉ xét khí thải từ quá trình đốt cháy. Không làm tròn kết quả các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng phần mười.)
""" + (r"""\loigiai{
Số mol C hóa thạch trong 100 gam biodiesel:\\
$n_{\ce{C}} = \frac{40}{270} + \frac{50}{296} + \frac{10}{298} = \frac{4}{27} + \frac{25}{148} + \frac{5}{149} \approx 0{,}350624\,\text{mol}$.\\
Khối lượng biodiesel sinh 1 MJ: $m = 880 / 32{,}0 = 27{,}5\,\text{gam}$.\\
$\text{EF}_{\text{hóa thạch}} = 44 \times 0{,}275 \times \left(\frac{4}{27} + \frac{25}{148} + \frac{5}{149}\right) \approx 4{,}24255\,\text{g \ce{CO2}/MJ}$.\\
$\%\Delta = \frac{74{,}1 - 4{,}24255}{74{,}1} \times 100\% \approx 94{,}3\%$.\\
\textbf{Đáp số: 94,3.}}""" if is_teacher else r"""\par\textbf{Trả lời:} \dotfill""") + r"""
\end{ex}

\begin{ex}
Cellulose là thành phần chính của thành tế bào thực vật, làm nên bộ khung của cây cối. Cellulose có nhiều trong bông, đay, tre, nứa, gỗ,... Từ cellulose người ta điều chế được cellulose acetate (có thành phần chủ yếu là cellulose triacetate và cellulose diacetate) bằng phản ứng acetyl hóa theo các phương trình sau:\\
(1) \ce{[C6H7O2(OH)3]n + 2n(CH3CO)2O -> [C6H7O2(OCOCH3)2(OH)]n + 2nCH3COOH}\\
(2) \ce{[C6H7O2(OH)3]n + 3n(CH3CO)2O -> [C6H7O2(OCOCH3)3]n + 3nCH3COOH}\\
Acetyl hóa hoàn toàn 16,2 gam cellulose bằng acetic anhydride \ce{(CH3CO)2O} dư trong điều kiện thích hợp, thu được 26,7 gam hỗn hợp X chỉ gồm cellulose diacetate và cellulose triacetate. Giả sử toàn bộ cellulose chuyển thành X, mạch polymer không bị cắt và bỏ qua ảnh hưởng của các nhóm chức ở hai đầu mạch. Cho biết mỗi nhóm \ce{-OH} được acetyl hóa chuyển thành nhóm \ce{-OCOCH3}.\\
Xác định phần trăm khối lượng cellulose triacetate trong X. (làm tròn đến hàng phần mười).
""" + (r"""\loigiai{
Số mol mắt xích cellulose: $n = 16{,}2 / 162 = 0{,}10\,\text{mol}$.\\
Hệ phương trình:\\
$x + y = 0{,}10$\\
$246x + 288y = 26{,}7$\\
$\Rightarrow x = 0{,}05\,\text{mol};\; y = 0{,}05\,\text{mol}$.\\
$\%m_{\text{triacetate}} = \frac{0{,}05 \times 288}{26{,}7} \times 100\% \approx 53{,}9\%$.\\
\textbf{Đáp số: 53,9.}}""" if is_teacher else r"""\par\textbf{Trả lời:} \dotfill""") + r"""
\end{ex}

\begin{center}
\textbf{--- HẾT ---}
\end{center}

\end{document}
""")

    return "".join(tex)

def compile_pdf(tex_path):
    out_dir = os.path.dirname(tex_path)
    basename = os.path.splitext(os.path.basename(tex_path))[0]
    cmd = ["pdflatex", "-interaction=nonstopmode", f"-output-directory={out_dir}", tex_path]
    
    print(f"Compiling {tex_path}...")
    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, cwd=out_dir)
    # Run second time for page numbers and references
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, cwd=out_dir)
    
    # Clean aux files
    for ext in [".aux", ".log", ".out", ".thm"]:
        f_del = os.path.join(out_dir, basename + ext)
        if os.path.exists(f_del):
            try:
                os.remove(f_del)
            except Exception:
                pass
    
    for f in os.listdir(out_dir):
        if f.startswith("ans-nh-") and f.endswith(".tex"):
            try:
                os.remove(os.path.join(out_dir, f))
            except Exception:
                pass

if __name__ == "__main__":
    hs_tex = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenHien_2026_HocSinh.tex")
    gv_tex = os.path.join(SAN_PHAM_DIR, "De_Kiem_Tra_GK1_Hoa_12_NguyenHien_2026_GiaoVien.tex")
    
    with open(hs_tex, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=False))
        
    with open(gv_tex, "w", encoding="utf-8") as f:
        f.write(build_latex_content(is_teacher=True))
        
    compile_pdf(hs_tex)
    compile_pdf(gv_tex)
    print("Done compiling LaTeX to PDF!")
