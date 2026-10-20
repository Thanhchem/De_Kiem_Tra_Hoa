# -*- coding: utf-8 -*-
"""
Script biên soạn duy nhất 1 file PDF mẫu ĐỀ SỐ 1 chuẩn TIVI 16:9 (KHÔNG VIỀN ĐEN):
- Tỉ lệ chuẩn 16:9 widescreen (paperwidth=29.7cm, paperheight=16.706cm)
- Hoàn toàn KHÔNG DÙNG KHUNG (bỏ tất cả ra khỏi khung)
- Cỡ chữ to 20pt (rõ nét từ xa 6 - 8 mét trên Tivi 65 inch)
- Phân bổ 2 câu hỏi vào 1 trang hợp lý, dàn đều toàn bộ chiều cao màn hình (không để trống nửa dưới)
- Không header/footer trang sách
- Slide Bảng đáp án ở cuối đề
"""

import os
import sys
import subprocess
import shutil
import pymupdf

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

TEX_FILE = os.path.join(SCRATCH_DIR, "De_On_Tap_Hoa_12_De01_16_9_Tivi.tex")
PDF_OUT = os.path.join(OUT_DIR, "De_On_Tap_Hoa_12_De01_A5_NAM_NGANG_TrinhChieu_Tivi.pdf")

latex_content = r"""\documentclass[17pt]{extarticle}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{xcolor}
\usepackage{tabularx}
\usepackage{array}
\usepackage{booktabs}
\usepackage[
  paperwidth=29.7cm,
  paperheight=16.706cm,
  top=0.8cm,
  bottom=0.8cm,
  left=1.3cm,
  right=1.3cm,
  footskip=0pt,
  headheight=0pt,
  headsep=0pt
]{geometry}

\pagestyle{empty}

% Bảng màu chuyên dụng trình chiếu Tivi 65 inch
\definecolor{hdrBlue}{RGB}{13, 71, 161}      % Xanh dương đậm tiêu đề (#0D47A1)
\definecolor{subGreen}{RGB}{0, 105, 92}      % Xanh ngọc đậm (#00695C)
\definecolor{qBlue}{RGB}{21, 101, 192}       % Xanh dương cho Câu X. (#1565C0)
\definecolor{ansRed}{RGB}{198, 40, 40}       % Đỏ cho đáp án
\definecolor{textDark}{RGB}{15, 15, 15}      % Chữ đen nhánh rõ nét

\linespread{1.22}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}

% Lệnh in tiêu đề Phần I
\newcommand{\slideHeaderI}{%
  {\color{hdrBlue}\fontsize{22}{26}\selectfont\textbf{ĐỀ ÔN TẬP CHƯƠNG 1, 2, 3 -- ĐỀ SỐ 1}}\\[3pt]
  {\color{subGreen}\fontsize{17.5}{21}\selectfont\textbf{Phần I. Câu trắc nghiệm nhiều phương án lựa chọn}}\\[2pt]
  {\fontsize{13.5}{16}\selectfont\textit{Mỗi câu hỏi, thí sinh chỉ chọn một phương án.}}\\[6pt]
}

% Lệnh in tiêu đề Phần II
\newcommand{\slideHeaderII}{%
  {\color{hdrBlue}\fontsize{22}{26}\selectfont\textbf{ĐỀ ÔN TẬP CHƯƠNG 1, 2, 3 -- ĐỀ SỐ 1}}\\[3pt]
  {\color{subGreen}\fontsize{17.5}{21}\selectfont\textbf{Phần II. Câu trắc nghiệm đúng sai}}\\[2pt]
  {\fontsize{13.5}{16}\selectfont\textit{Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.}}\\[6pt]
}

% Lệnh in tiêu đề Phần III
\newcommand{\slideHeaderIII}{%
  {\color{hdrBlue}\fontsize{22}{26}\selectfont\textbf{ĐỀ ÔN TẬP CHƯƠNG 1, 2, 3 -- ĐỀ SỐ 1}}\\[3pt]
  {\color{subGreen}\fontsize{17.5}{21}\selectfont\textbf{Phần III. Câu trắc nghiệm trả lời ngắn}}\\[2pt]
  {\fontsize{13.5}{16}\selectfont\textit{Thí sinh trả lời từ câu 1 đến câu 6.}}\\[6pt]
}

\begin{document}
\fontsize{19.5}{25}\selectfont
\color{textDark}

% ==============================================================================
% SLIDE 1: CÂU 1 & CÂU 2
% ==============================================================================
\slideHeaderI
\vfill

{\color{qBlue}\textbf{Câu 1.}} Thành phần chính của chất giặt rửa tổng hợp là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{C}_{17}\text{H}_{33}\text{COOK}$. & \textbf{B.} $(\text{C}_{17}\text{H}_{33}\text{COO})_2\text{Ca}$.\\[8pt]
\textbf{C.} $\text{CH}_3\text{COONa}$. & \textbf{D.} $\text{CH}_3[\text{CH}_2]_{11}\text{C}_6\text{H}_4\text{SO}_3\text{Na}$.
\end{tabularx}

\vfill

{\color{qBlue}\textbf{Câu 2.}} Ester nào sau đây có công thức $\text{CH}_3\text{COOCH}=\text{CH}_2$?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Methyl acetate. & \textbf{B.} Vinyl acetate.\\[8pt]
\textbf{C.} Ethyl formate. & \textbf{D.} Methyl acrylate.
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 2: CÂU 3 & CÂU 4
% ==============================================================================
\slideHeaderI
\vfill

{\color{qBlue}\textbf{Câu 3.}} Hợp chất $\text{CH}_3\text{NHCH}_3$ thuộc loại amine nào?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Bậc 1. & \textbf{B.} Bậc 2.\\[8pt]
\textbf{C.} Bậc 3. & \textbf{D.} Amine thơm.
\end{tabularx}

\vfill

{\color{qBlue}\textbf{Câu 4.}} Trong phân tử cellulose, các đơn vị glucose liên kết với nhau bằng liên kết
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\alpha$-1,4-glycoside. & \textbf{B.} $\beta$-1,4-glycoside.\\[8pt]
\textbf{C.} $\alpha$-1,6-glycoside. & \textbf{D.} $\beta$-1,6-glycoside.
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 3: CÂU 5 & CÂU 6
% ==============================================================================
\slideHeaderI
\vfill

{\color{qBlue}\textbf{Câu 5.}} Amino acid là hợp chất hữu cơ trong phân tử chứa các nhóm chức
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $-\text{COOH}$ và $-\text{OH}$. & \textbf{B.} $-\text{NH}_2$ và $-\text{COOH}$.\\[8pt]
\textbf{C.} $-\text{CHO}$ và $-\text{OH}$. & \textbf{D.} $-\text{NH}_2$ và $-\text{OH}$.
\end{tabularx}

\vfill

{\color{qBlue}\textbf{Câu 6.}} Để phân biệt glucose và glycerol ở nhiệt độ thường, thuốc thử phù hợp là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Dung dịch $\text{NaOH}$. & \textbf{B.} $\text{Cu(OH)}_2$ đun nóng.\\[8pt]
\textbf{C.} Thuốc thử Tollens, $t^\circ$. & \textbf{D.} Dung dịch $\text{HCl}$.
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 4: CÂU 7 & CÂU 8
% ==============================================================================
\slideHeaderI
\vfill

{\color{qBlue}\textbf{Câu 7.}} Chất chủ yếu trong tinh bột tạo nên quá trình hồ hóa khi đun nóng trong nước là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Amylose. & \textbf{B.} Amylopectin.\\[8pt]
\textbf{C.} Glucose. & \textbf{D.} Saccharose.
\end{tabularx}

\vfill

{\color{qBlue}\textbf{Câu 8.}} Thủy phân hoàn toàn peptide bởi dung dịch kiềm nóng thu được sản phẩm là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Ester. & \textbf{B.} Các muối của $\alpha$-amino acid.\\[8pt]
\textbf{C.} Amine. & \textbf{D.} Alcohol.
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 5: CÂU 9 & CÂU 10
% ==============================================================================
\slideHeaderI
\vfill

{\color{qBlue}\textbf{Câu 9.}} Chất nào sau đây vừa phản ứng với dung dịch $\text{NaOH}$ vừa phản ứng với dung dịch $\text{HCl}$?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{H}_2\text{N-CH}_2\text{-COOH}$. & \textbf{B.} $\text{CH}_3\text{COOH}$.\\[8pt]
\textbf{C.} $\text{C}_2\text{H}_5\text{OH}$. & \textbf{D.} $\text{CH}_3\text{NH}_2$.
\end{tabularx}

\vfill

{\color{qBlue}\textbf{Câu 10.}} Phản ứng xà phòng hóa chất béo tristearin bằng dung dịch $\text{NaOH}$ thu được xà phòng và
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Ethylene glycol. & \textbf{B.} Glycerol.\\[8pt]
\textbf{C.} Ethanol. & \textbf{D.} Methanol.
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 6: CÂU 11 & CÂU 12
% ==============================================================================
\slideHeaderI
\vfill

{\color{qBlue}\textbf{Câu 11.}} Công thức phân tử của saccharose là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{C}_6\text{H}_{12}\text{O}_6$. & \textbf{B.} $\text{C}_{12}\text{H}_{22}\text{O}_{11}$.\\[8pt]
\textbf{C.} $(\text{C}_6\text{H}_{10}\text{O}_5)_n$. & \textbf{D.} $\text{C}_3\text{H}_6\text{O}_2$.
\end{tabularx}

\vfill

{\color{qBlue}\textbf{Câu 12.}} Tên gọi của dipeptide $\text{H}_2\text{N-CH}_2\text{-CO-NH-CH(CH}_3\text{)COOH}$ là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Ala-Gly. & \textbf{B.} Gly-Ala.\\[8pt]
\textbf{C.} Gly-Val. & \textbf{D.} Ala-Val.
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 7: CÂU 13 & CÂU 14
% ==============================================================================
\slideHeaderI
\vfill

{\color{qBlue}\textbf{Câu 13.}} Carbohydrate nào sau đây thuộc loại disaccharide?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Glucose. & \textbf{B.} Fructose.\\[8pt]
\textbf{C.} Saccharose. & \textbf{D.} Tinh bột.
\end{tabularx}

\vfill

{\color{qBlue}\textbf{Câu 14.}} Hydrogen hóa hoàn toàn triolein (xúc tác $\text{Ni}, t^\circ$) thu được
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Tristearin. & \textbf{B.} Tripalmitin.\\[8pt]
\textbf{C.} Glycerol. & \textbf{D.} Oleic acid.
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 8: CÂU 15 & CÂU 16
% ==============================================================================
\slideHeaderI
\vfill

{\color{qBlue}\textbf{Câu 15.}} Amine nào sau đây là arylamine (amine thơm)?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{CH}_3\text{NH}_2$. & \textbf{B.} $\text{C}_6\text{H}_5\text{NH}_2$.\\[8pt]
\textbf{C.} $\text{C}_2\text{H}_5\text{NH}_2$. & \textbf{D.} $(\text{CH}_3)_2\text{NH}$.
\end{tabularx}

\vfill

{\color{qBlue}\textbf{Câu 16.}} Ở điều kiện thường, chất nào sau đây là chất khí?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Aniline. & \textbf{B.} Methylamine.\\[8pt]
\textbf{C.} Glycine. & \textbf{D.} Glucose.
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 9: CÂU 17 & CÂU 18
% ==============================================================================
\slideHeaderI
\vfill

{\color{qBlue}\textbf{Câu 17.}} Cho dung dịch glucose tác dụng với $\text{Cu(OH)}_2/\text{NaOH}$ đun nóng tạo kết tủa màu
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Xanh lam. & \textbf{B.} Đỏ gạch.\\[8pt]
\textbf{C.} Trắng. & \textbf{D.} Vàng.
\end{tabularx}

\vfill

{\color{qBlue}\textbf{Câu 18.}} Peptide Gly-Ala có số liên kết peptide là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} 1. & \textbf{B.} 2.\\[8pt]
\textbf{C.} 3. & \textbf{D.} 0.
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 10: PHẦN II - CÂU 1
% ==============================================================================
\slideHeaderII
\vfill

{\color{qBlue}\textbf{Câu 1.}} Về cấu tạo và tính chất của chất béo Trilinolein $(\text{C}_{17}\text{H}_{31}\text{COO})_3\text{C}_3\text{H}_5$:
\vspace{12pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} Công thức phân tử của trilinolein là $\text{C}_{57}\text{H}_{98}\text{O}_6$. & $[\quad]$\\[10pt]
\textbf{b)} Thủy phân hoàn toàn trong acid thu được acid béo không no. & $[\quad]$\\[10pt]
\textbf{c)} Hydrogen hóa hoàn toàn trilinolein thu được chất béo lỏng. & $[\quad]$\\[10pt]
\textbf{d)} Trilinolein làm mất màu dung dịch nước bromine. & $[\quad]$
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 11: PHẦN II - CÂU 2
% ==============================================================================
\slideHeaderII
\vfill

{\color{qBlue}\textbf{Câu 2.}} Cho các điểm đẳng điện pI: Glycine (6,0), Glutamic acid (3,2), Lysine (9,7):
\vspace{12pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} Tên bán hệ thống của glycine là $\alpha$-aminoacetic acid. & $[\quad]$\\[10pt]
\textbf{b)} Trong dung dịch $\text{pH} = 6,0$, glycine tồn tại dạng cation về cực âm. & $[\quad]$\\[10pt]
\textbf{c)} Ở $\text{pH} = 11,0$, lysine tồn tại ở dạng anion về cực dương. & $[\quad]$\\[10pt]
\textbf{d)} Lysine có 2 nhóm $-\text{NH}_2$ và 1 nhóm $-\text{COOH}$ trong phân tử. & $[\quad]$
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 12: PHẦN II - CÂU 3
% ==============================================================================
\slideHeaderII
\vfill

{\color{qBlue}\textbf{Câu 3.}} Xét peptide X có công thức Ala-Gly-Val:
\vspace{12pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} X là một tripeptide. & $[\quad]$\\[10pt]
\textbf{b)} Amino acid đầu N là Valine, đầu C là Alanine. & $[\quad]$\\[10pt]
\textbf{c)} X có phản ứng màu biuret tạo dung dịch màu tím đặc trưng. & $[\quad]$\\[10pt]
\textbf{d)} Phân tử X chứa 3 liên kết peptide. & $[\quad]$
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 13: PHẦN II - CÂU 4
% ==============================================================================
\slideHeaderII
\vfill

{\color{qBlue}\textbf{Câu 4.}} Về thí nghiệm ester hóa giữa $\text{CH}_3\text{COOH}$ và $\text{C}_2\text{H}_5\text{OH}$ (xúc tác $\text{H}_2\text{SO}_4$ đặc):
\vspace{12pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} Phản ứng xảy ra là phản ứng thuận nghịch. & $[\quad]$\\[10pt]
\textbf{b)} $\text{H}_2\text{SO}_4$ đặc vừa làm xúc tác vừa tăng hiệu suất bằng cách hút nước. & $[\quad]$\\[10pt]
\textbf{c)} Tách ethyl acetate bằng phương pháp chiết vì không tan, nhẹ hơn nước. & $[\quad]$\\[10pt]
\textbf{d)} Sau phản ứng thu được dung dịch hoàn toàn đồng nhất. & $[\quad]$
\end{tabularx}

\vfill
\newpage

% ==============================================================================
% SLIDE 14: PHẦN III - CÂU 1 & CÂU 2
% ==============================================================================
\slideHeaderIII
\vfill

{\color{qBlue}\textbf{Câu 1.}} Thứ tự đúng các bước thí nghiệm điều chế xà phòng: (1) Thêm $\text{NaCl}$ bão hòa nóng; (2) Cho dầu ăn + $\text{NaOH}$ vào bát sứ; (3) Đun sôi nhẹ và khuấy đều; (4) Thêm nước cất giữ thể tích không đổi.\\[6pt]
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\vfill

{\color{qBlue}\textbf{Câu 2.}} Tính thu nhập từ nấu rượu gạo: 320 kg gạo (chứa 80\% tinh bột) lên men thu rượu 40$^\circ$ với hiệu suất 70\%, $D(\text{ethanol}) = 0,8\text{ g/mL}$, giá 50.000 đ/lít. Doanh thu thu được là bao nhiêu triệu đồng? \textit{(làm tròn đến hàng phần mười)}\\[6pt]
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\vfill
\newpage

% ==============================================================================
% SLIDE 15: PHẦN III - CÂU 3 & CÂU 4
% ==============================================================================
\slideHeaderIII
\vfill

{\color{qBlue}\textbf{Câu 3.}} Có bao nhiêu chất trong dãy tác dụng được với dung dịch $\text{NaOH}$ nóng: lysine, aniline, methyl acrylate, ethyl formate, maltose, tripalmitin?\\[6pt]
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\vfill

{\color{qBlue}\textbf{Câu 4.}} Số nguyên tử oxygen trong phân tử tripeptide Gly-Ala-Ala là bao nhiêu?\\[6pt]
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\vfill
\newpage

% ==============================================================================
% SLIDE 16: PHẦN III - CÂU 5 & CÂU 6
% ==============================================================================
\slideHeaderIII
\vfill

{\color{qBlue}\textbf{Câu 5.}} Thủy phân hoàn toàn 0,1 mol phenyl acetate $\text{CH}_3\text{COOC}_6\text{H}_5$ trong dung dịch $\text{NaOH}$ dư. Khối lượng muối thu được là bao nhiêu gam?\\[6pt]
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\vfill

{\color{qBlue}\textbf{Câu 6.}} Để sản xuất 1 tấn glucose với hiệu suất 80\%, khối lượng tinh bột cần dùng là bao nhiêu tấn? \textit{(làm tròn đến hàng phần trăm)}\\[6pt]
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\vfill
\newpage

% ==============================================================================
% SLIDE 17: BẢNG ĐÁP ÁN ĐỀ SỐ 1
% ==============================================================================
\begin{center}
{\color{hdrBlue}\fontsize{24}{28}\selectfont\textbf{BẢNG ĐÁP ÁN CHÍNH THỨC -- ĐỀ SỐ 1}}\\[12pt]
\end{center}

{\color{qBlue}\fontsize{18}{22}\selectfont\textbf{1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (4,5 ĐIỂM)}}\\[6pt]
\renewcommand{\arraystretch}{1.3}
\begin{tabularx}{\linewidth}{|*{9}{>{\centering\arraybackslash}X|}}
\hline
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} \\
\hline
\color{ansRed}\textbf{D} & \color{ansRed}\textbf{B} & \color{ansRed}\textbf{B} & \color{ansRed}\textbf{B} & \color{ansRed}\textbf{B} & \color{ansRed}\textbf{C} & \color{ansRed}\textbf{B} & \color{ansRed}\textbf{B} & \color{ansRed}\textbf{A} \\
\hline
\textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} & \textbf{17} & \textbf{18} \\
\hline
\color{ansRed}\textbf{B} & \color{ansRed}\textbf{B} & \color{ansRed}\textbf{B} & \color{ansRed}\textbf{C} & \color{ansRed}\textbf{A} & \color{ansRed}\textbf{B} & \color{ansRed}\textbf{B} & \color{ansRed}\textbf{B} & \color{ansRed}\textbf{A} \\
\hline
\end{tabularx}

\vspace{18pt}

{\color{subGreen}\fontsize{18}{22}\selectfont\textbf{2. PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG SAI (4,0 ĐIỂM)}}\\[6pt]
\begin{tabularx}{\linewidth}{@{}*{4}{>{\raggedright\arraybackslash}X}@{}}
\textbf{Câu 1:} [\color{ansRed}\textbf{Đ--Đ--S--Đ}\color{textDark}] &
\textbf{Câu 2:} [\color{ansRed}\textbf{Đ--S--Đ--Đ}\color{textDark}] &
\textbf{Câu 3:} [\color{ansRed}\textbf{Đ--S--Đ--S}\color{textDark}] &
\textbf{Câu 4:} [\color{ansRed}\textbf{Đ--Đ--Đ--S}\color{textDark}]
\end{tabularx}

\vspace{18pt}

{\color{ansRed}\fontsize{18}{22}\selectfont\textbf{3. PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (1,5 ĐIỂM)}}\\[6pt]
\begin{tabularx}{\linewidth}{@{}XXX@{}}
\textbf{Câu 1:} {\color{ansRed}\textbf{2341}} & \textbf{Câu 2:} {\color{ansRed}\textbf{12,8}} & \textbf{Câu 3:} {\color{ansRed}\textbf{4}} \\[8pt]
\textbf{Câu 4:} {\color{ansRed}\textbf{4}} & \textbf{Câu 5:} {\color{ansRed}\textbf{17,6}} & \textbf{Câu 6:} {\color{ansRed}\textbf{1,13}}
\end{tabularx}

\end{document}
"""

with open(TEX_FILE, "w", encoding="utf-8") as f:
    f.write(latex_content)

print(f"Đã ghi file TeX: {TEX_FILE}")

cmd = ["pdflatex", "-interaction=nonstopmode", "-output-directory", SCRATCH_DIR, TEX_FILE]
res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

pdf_temp = os.path.join(SCRATCH_DIR, "De_On_Tap_Hoa_12_De01_16_9_Tivi.pdf")
if os.path.exists(pdf_temp):
    shutil.copy2(pdf_temp, PDF_OUT)
    print(f"BIÊN DỊCH THÀNH CÔNG! File PDF: {PDF_OUT} ({os.path.getsize(PDF_OUT):,} bytes)")

    # Sao lưu sang OneDrive
    od1 = r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12\De_On_Tap_Hoa_12_De01_A5_NAM_NGANG_TrinhChieu_Tivi.pdf"
    try:
        shutil.copy2(PDF_OUT, od1)
        print(f"Đã sao lưu sang OneDrive: {od1}")
    except Exception as e:
        print(f"Lỗi OneDrive: {e}")
else:
    print("LỖI BIÊN DỊCH:")
    print(res.stdout[-1500:])
