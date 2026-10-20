# -*- coding: utf-8 -*-
"""
Tạo duy nhất 1 file PDF mẫu ĐỀ SỐ 1 khổ A5 NẰM NGANG chuyên dụng TRÌNH CHIẾU TIVI 65 INCH
Bộ Ôn tập Chương 1, 2, 3 - Hóa học 12 (Năm học 2026 - 2027)
- Khổ A5 nằm ngang (Landscape: 210mm x 148mm)
- Cỡ chữ lớn 14pt (extarticle), phông Times New Roman chuẩn
- Tối đa 1-2 câu / slide, câu hỏi nguyên vẹn, không ngắt đôi
- KHÔNG header, KHÔNG footer
- Chỉ chiếu đề bài, kèm trang BẢNG ĐÁP ÁN ở cuối đề
"""

import os
import sys
import subprocess
import shutil

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

TEX_FILE = os.path.join(SCRATCH_DIR, "De_On_Tap_Hoa_12_De01_A5_Tivi.tex")
PDF_OUT = os.path.join(OUT_DIR, "De_On_Tap_Hoa_12_De01_A5_NAM_NGANG_TrinhChieu_Tivi.pdf")

latex_content = r"""\documentclass[14pt]{extarticle}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{graphicx}
\usepackage[table]{xcolor}
\usepackage{tcolorbox}
\usepackage{tabularx}
\usepackage{array}
\usepackage{booktabs}
\usepackage[a5paper, landscape, top=0.7cm, bottom=0.7cm, left=0.8cm, right=0.8cm, footskip=0pt, headheight=0pt, headsep=0pt]{geometry}

\pagestyle{empty}

% Định nghĩa màu sắc chuyên dụng Tivi 65 inch
\definecolor{tvBlue}{RGB}{13, 71, 161}      % Xanh dương đậm sang trọng
\definecolor{tvCyan}{RGB}{0, 131, 143}     % Xanh mòng két
\definecolor{tvRed}{RGB}{198, 40, 40}      % Đỏ nổi bật
\definecolor{tvDark}{RGB}{33, 33, 33}      % Chữ đen đậm
\definecolor{tvBoxBg}{RGB}{245, 247, 250}  % Nền hộp câu hỏi
\definecolor{tvBorder}{RGB}{176, 190, 197} % Viền hộp
\definecolor{ansBg}{RGB}{232, 245, 233}    % Nền bảng đáp án
\definecolor{ansBorder}{RGB}{76, 175, 80}  % Viền bảng đáp án

\tcbuselibrary{skins,breakable}

% Box kiểu dáng hiện đại cho từng câu hỏi
\newtcolorbox{qbox}[2][]{
  enhanced,
  colback=white,
  colframe=tvBlue,
  coltitle=white,
  fonttitle=\bfseries\fontsize{13}{15}\selectfont,
  title={#2},
  arc=3mm,
  boxrule=1.2pt,
  left=6pt, right=6pt, top=4pt, bottom=4pt,
  boxsep=2pt,
  before skip=3pt, after skip=3pt,
  #1
}

\newtcolorbox{qboxcyan}[2][]{
  enhanced,
  colback=white,
  colframe=tvCyan,
  coltitle=white,
  fonttitle=\bfseries\fontsize{13}{15}\selectfont,
  title={#2},
  arc=3mm,
  boxrule=1.2pt,
  left=6pt, right=6pt, top=4pt, bottom=4pt,
  boxsep=2pt,
  before skip=3pt, after skip=3pt,
  #1
}

\linespread{1.15}
\setlength{\parindent}{0pt}
\setlength{\parskip}{2pt}

\begin{document}

% ==============================================================================
% SLIDE 1: TRANG TIÊU ĐỀ ĐỀ SỐ 1
% ==============================================================================
\begin{center}
\vspace*{0.3cm}
{\color{tvBlue}\fontsize{15}{18}\selectfont\textbf{BỘ ĐỀ ÔN TẬP HÓA HỌC 12 -- CHƯƠNG 1, 2, 3}}\\[0.2cm]
{\color{tvDark}\fontsize{13}{16}\selectfont\textit{NĂM HỌC 2026 -- 2027}}\\[0.35cm]

\begin{tcolorbox}[enhanced, colback=tvBlue!8, colframe=tvBlue, arc=4mm, width=0.92\linewidth, halign=center, boxrule=1.8pt]
  \vspace{0.2cm}
  {\color{tvRed}\fontsize{24}{28}\selectfont\textbf{ĐỀ ÔN TẬP SỐ 01}}\\[0.3cm]
  {\color{tvCyan}\fontsize{15}{18}\selectfont\textbf{ESTERS -- LIPIDS $\bullet$ CARBOHYDRATES $\bullet$ AMINES -- AMINO ACIDS -- PROTEINS}}\\[0.25cm]
  {\color{tvDark}\fontsize{13}{16}\selectfont\textbf{Thời gian làm bài: 30 phút -- Thầy TRẦN VĂN THẠNH (0777.470.803)}}
  \vspace{0.2cm}
\end{tcolorbox}

\vspace{0.4cm}
{\color{tvDark}\fontsize{12}{15}\selectfont\textbf{Cấu trúc đề:}\ \textbf{Phần I:} 18 câu trắc nghiệm (4,5đ)\ $\bullet$\ \textbf{Phần II:} 4 câu Đúng/Sai (4,0đ)\ $\bullet$\ \textbf{Phần III:} 6 câu trả lời ngắn (1,5đ)}
\end{center}
\newpage

% ==============================================================================
% SLIDE 2: PHẦN I - CÂU 1 & CÂU 2
% ==============================================================================
{\color{tvBlue}\fontsize{12}{14}\selectfont\textbf{PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (4,5 ĐIỂM) -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 1: Thành phần chất giặt rửa tổng hợp}
Thành phần chính của chất giặt rửa tổng hợp là
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} $\text{C}_{17}\text{H}_{33}\text{COOK}$. & \textbf{B.} $(\text{C}_{17}\text{H}_{33}\text{COO})_2\text{Ca}$.\\[3pt]
\textbf{C.} $\text{CH}_3\text{COONa}$. & \textbf{D.} $\text{CH}_3[\text{CH}_2]_{11}\text{C}_6\text{H}_4\text{SO}_3\text{Na}$.
\end{tabularx}
\end{qbox}

\begin{qbox}{Câu 2: Công thức cấu tạo của ester}
Ester nào sau đây có công thức cấu tạo $\text{CH}_3\text{COOCH}=\text{CH}_2$?
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} Methyl acetate. & \textbf{B.} Vinyl acetate.\\[3pt]
\textbf{C.} Ethyl formate. & \textbf{D.} Methyl acrylate.
\end{tabularx}
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 3: PHẦN I - CÂU 3 & CÂU 4
% ==============================================================================
{\color{tvBlue}\fontsize{12}{14}\selectfont\textbf{PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 3: Bậc của amine}
Hợp chất $\text{CH}_3\text{NHCH}_3$ thuộc loại amine nào?
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} Bậc 1. & \textbf{B.} Bậc 2.\\[3pt]
\textbf{C.} Bậc 3. & \textbf{D.} Amine thơm.
\end{tabularx}
\end{qbox}

\begin{qbox}{Câu 4: Liên kết trong phân tử cellulose}
Trong phân tử cellulose, các đơn vị glucose liên kết với nhau bằng liên kết
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} $\alpha$-1,4-glycoside. & \textbf{B.} $\beta$-1,4-glycoside.\\[3pt]
\textbf{C.} $\alpha$-1,6-glycoside. & \textbf{D.} $\beta$-1,6-glycoside.
\end{tabularx}
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 4: PHẦN I - CÂU 5 & CÂU 6
% ==============================================================================
{\color{tvBlue}\fontsize{12}{14}\selectfont\textbf{PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 5: Nhóm chức trong amino acid}
Amino acid là hợp chất hữu cơ trong phân tử chứa các nhóm chức
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} $-\text{COOH}$ và $-\text{OH}$. & \textbf{B.} $-\text{NH}_2$ và $-\text{COOH}$.\\[3pt]
\textbf{C.} $-\text{CHO}$ và $-\text{OH}$. & \textbf{D.} $-\text{NH}_2$ và $-\text{OH}$.
\end{tabularx}
\end{qbox}

\begin{qbox}{Câu 6: Phân biệt glucose và glycerol}
Để phân biệt glucose và glycerol ở nhiệt độ thường, thuốc thử phù hợp là
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} Dung dịch $\text{NaOH}$. & \textbf{B.} $\text{Cu(OH)}_2$ đun nóng.\\[3pt]
\textbf{C.} Thuốc thử Tollens, $t^\circ$. & \textbf{D.} Dung dịch $\text{HCl}$.
\end{tabularx}
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 5: PHẦN I - CÂU 7 & CÂU 8
% ==============================================================================
{\color{tvBlue}\fontsize{12}{14}\selectfont\textbf{PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 7: Quá trình hồ hóa tinh bột}
Chất chủ yếu trong tinh bột tạo nên quá trình hồ hóa khi đun nóng trong nước là
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} Amylose. & \textbf{B.} Amylopectin.\\[3pt]
\textbf{C.} Glucose. & \textbf{D.} Saccharose.
\end{tabularx}
\end{qbox}

\begin{qbox}{Câu 8: Thủy phân hoàn toàn peptide}
Thủy phân hoàn toàn peptide bởi dung dịch kiềm nóng thu được sản phẩm là
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} Ester. & \textbf{B.} Các muối của $\alpha$-amino acid.\\[3pt]
\textbf{C.} Amine. & \textbf{D.} Alcohol.
\end{tabularx}
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 6: PHẦN I - CÂU 9 & CÂU 10
% ==============================================================================
{\color{tvBlue}\fontsize{12}{14}\selectfont\textbf{PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 9: Tính chất lưỡng tính}
Chất nào sau đây vừa phản ứng với dung dịch $\text{NaOH}$ vừa phản ứng với dung dịch $\text{HCl}$?
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} $\text{H}_2\text{N-CH}_2\text{-COOH}$. & \textbf{B.} $\text{CH}_3\text{COOH}$.\\[3pt]
\textbf{C.} $\text{C}_2\text{H}_5\text{OH}$. & \textbf{D.} $\text{CH}_3\text{NH}_2$.
\end{tabularx}
\end{qbox}

\begin{qbox}{Câu 10: Xà phòng hóa tristearin}
Phản ứng xà phòng hóa chất béo tristearin bằng dung dịch $\text{NaOH}$ thu được xà phòng và
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} Ethylene glycol. & \textbf{B.} Glycerol.\\[3pt]
\textbf{C.} Ethanol. & \textbf{D.} Methanol.
\end{tabularx}
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 7: PHẦN I - CÂU 11 & CÂU 12
% ==============================================================================
{\color{tvBlue}\fontsize{12}{14}\selectfont\textbf{PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 11: Công thức phân tử saccharose}
Công thức phân tử của saccharose là
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} $\text{C}_6\text{H}_{12}\text{O}_6$. & \textbf{B.} $\text{C}_{12}\text{H}_{22}\text{O}_{11}$.\\[3pt]
\textbf{C.} $(\text{C}_6\text{H}_{10}\text{O}_5)_n$. & \textbf{D.} $\text{C}_3\text{H}_6\text{O}_2$.
\end{tabularx}
\end{qbox}

\begin{qbox}{Câu 12: Tên gọi dipeptide}
Tên gọi của dipeptide $\text{H}_2\text{N-CH}_2\text{-CO-NH-CH(CH}_3\text{)COOH}$ là
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} Ala-Gly. & \textbf{B.} Gly-Ala.\\[3pt]
\textbf{C.} Gly-Val. & \textbf{D.} Ala-Val.
\end{tabularx}
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 8: PHẦN I - CÂU 13 & CÂU 14
% ==============================================================================
{\color{tvBlue}\fontsize{12}{14}\selectfont\textbf{PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 13: Phân loại disaccharide}
Carbohydrate nào sau đây thuộc loại disaccharide?
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} Glucose. & \textbf{B.} Fructose.\\[3pt]
\textbf{C.} Saccharose. & \textbf{D.} Tinh bột.
\end{tabularx}
\end{qbox}

\begin{qbox}{Câu 14: Hydrogen hóa triolein}
Hydrogen hóa hoàn toàn triolein (xúc tác $\text{Ni}, t^\circ$) thu được
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} Tristearin. & \textbf{B.} Tripalmitin.\\[3pt]
\textbf{C.} Glycerol. & \textbf{D.} Oleic acid.
\end{tabularx}
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 9: PHẦN I - CÂU 15 & CÂU 16
% ==============================================================================
{\color{tvBlue}\fontsize{12}{14}\selectfont\textbf{PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 15: Amine thơm (Arylamine)}
Amine nào sau đây là arylamine (amine thơm)?
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} $\text{CH}_3\text{NH}_2$. & \textbf{B.} $\text{C}_6\text{H}_5\text{NH}_2$.\\[3pt]
\textbf{C.} $\text{C}_2\text{H}_5\text{NH}_2$. & \textbf{D.} $(\text{CH}_3)_2\text{NH}$.
\end{tabularx}
\end{qbox}

\begin{qbox}{Câu 16: Trạng thái của amine}
Ở điều kiện thường, chất nào sau đây là chất khí?
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} Aniline. & \textbf{B.} Methylamine.\\[3pt]
\textbf{C.} Glycine. & \textbf{D.} Glucose.
\end{tabularx}
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 10: PHẦN I - CÂU 17 & CÂU 18
% ==============================================================================
{\color{tvBlue}\fontsize{12}{14}\selectfont\textbf{PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 17: Phản ứng tạo kết tủa Cu2O}
Cho dung dịch glucose tác dụng với $\text{Cu(OH)}_2/\text{NaOH}$ đun nóng tạo kết tủa màu
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} Xanh lam. & \textbf{B.} Đỏ gạch.\\[3pt]
\textbf{C.} Trắng. & \textbf{D.} Vàng.
\end{tabularx}
\end{qbox}

\begin{qbox}{Câu 18: Số liên kết peptide}
Peptide Gly-Ala có số liên kết peptide là
\begin{tabularx}{\linewidth}{XX}
\textbf{A.} 1. & \textbf{B.} 2.\\[3pt]
\textbf{C.} 3. & \textbf{D.} 0.
\end{tabularx}
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 11: PHẦN II - CÂU 1 (ĐÚNG / SAI)
% ==============================================================================
{\color{tvCyan}\fontsize{12}{14}\selectfont\textbf{PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG SAI (4,0 ĐIỂM) -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qboxcyan}{Câu 1: Cấu tạo và tính chất của chất béo Trilinolein}
Về cấu tạo và tính chất của chất béo Trilinolein $(\text{C}_{17}\text{H}_{31}\text{COO})_3\text{C}_3\text{H}_5$:
\vspace{4pt}
\begin{itemize}
  \item[\textbf{a)}] Công thức phân tử của trilinolein là $\text{C}_{57}\text{H}_{98}\text{O}_6$. \hfill $\square$
  \item[\textbf{b)}] Thủy phân hoàn toàn trilinolein trong môi trường acid thu được acid béo không no. \hfill $\square$
  \item[\textbf{c)}] Hydrogen hóa hoàn toàn trilinolein thu được chất béo lỏng. \hfill $\square$
  \item[\textbf{d)}] Trilinolein làm mất màu dung dịch nước bromine. \hfill $\square$
\end{itemize}
\end{qboxcyan}
\newpage

% ==============================================================================
% SLIDE 12: PHẦN II - CÂU 2 (ĐÚNG / SAI)
% ==============================================================================
{\color{tvCyan}\fontsize{12}{14}\selectfont\textbf{PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG SAI (4,0 ĐIỂM) -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qboxcyan}{Câu 2: Điểm đẳng điện (pI) của amino acid}
Cho các điểm đẳng điện pI: Glycine (6,0), Glutamic acid (3,2), Lysine (9,7):
\vspace{4pt}
\begin{itemize}
  \item[\textbf{a)}] Tên bán hệ thống của glycine là $\alpha$-aminoacetic acid. \hfill $\square$
  \item[\textbf{b)}] Trong dung dịch $\text{pH} = 6,0$, glycine tồn tại chủ yếu ở dạng cation di chuyển về cực âm. \hfill $\square$
  \item[\textbf{c)}] Ở $\text{pH} = 11,0$, lysine tồn tại ở dạng anion di chuyển về cực dương trong điện trường. \hfill $\square$
  \item[\textbf{d)}] Lysine có 2 nhóm $-\text{NH}_2$ và 1 nhóm $-\text{COOH}$ trong phân tử. \hfill $\square$
\end{itemize}
\end{qboxcyan}
\newpage

% ==============================================================================
% SLIDE 13: PHẦN II - CÂU 3 (ĐÚNG / SAI)
% ==============================================================================
{\color{tvCyan}\fontsize{12}{14}\selectfont\textbf{PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG SAI (4,0 ĐIỂM) -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qboxcyan}{Câu 3: Cấu tạo và tính chất của peptide X}
Xét peptide X có công thức Ala-Gly-Val:
\vspace{4pt}
\begin{itemize}
  \item[\textbf{a)}] X là một tripeptide. \hfill $\square$
  \item[\textbf{b)}] Amino acid đầu N là Valine, đầu C là Alanine. \hfill $\square$
  \item[\textbf{c)}] X có phản ứng màu biuret tạo dung dịch màu tím đặc trưng. \hfill $\square$
  \item[\textbf{d)}] Phân tử X chứa 3 liên kết peptide. \hfill $\square$
\end{itemize}
\end{qboxcyan}
\newpage

% ==============================================================================
% SLIDE 14: PHẦN II - CÂU 4 (ĐÚNG / SAI)
% ==============================================================================
{\color{tvCyan}\fontsize{12}{14}\selectfont\textbf{PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG SAI (4,0 ĐIỂM) -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qboxcyan}{Câu 4: Thí nghiệm ester hóa điều chế ethyl acetate}
Về thí nghiệm ester hóa giữa $\text{CH}_3\text{COOH}$ và $\text{C}_2\text{H}_5\text{OH}$ (xúc tác $\text{H}_2\text{SO}_4$ đặc):
\vspace{4pt}
\begin{itemize}
  \item[\textbf{a)}] Phản ứng xảy ra là phản ứng thuận nghịch. \hfill $\square$
  \item[\textbf{b)}] $\text{H}_2\text{SO}_4$ đặc vừa làm xúc tác vừa làm tăng hiệu suất bằng cách hút nước. \hfill $\square$
  \item[\textbf{c)}] Tách ethyl acetate bằng phương pháp chiết vì nó không tan và nhẹ hơn nước. \hfill $\square$
  \item[\textbf{d)}] Sau phản ứng thu được dung dịch hoàn toàn đồng nhất. \hfill $\square$
\end{itemize}
\end{qboxcyan}
\newpage

% ==============================================================================
% SLIDE 15: PHẦN III - CÂU 1 & CÂU 2 (TRẢ LỜI NGẮN)
% ==============================================================================
{\color{tvRed}\fontsize{12}{14}\selectfont\textbf{PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (1,5 ĐIỂM) -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 1: Thứ tự các bước điều chế xà phòng}
Thứ tự đúng các bước thí nghiệm điều chế xà phòng:\\
(1) Thêm $\text{NaCl}$ bão hòa nóng;\\
(2) Cho dầu ăn + $\text{NaOH}$ vào bát sứ;\\
(3) Đun sôi nhẹ và khuấy đều;\\
(4) Thêm nước cất giữ thể tích không đổi.\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\end{qbox}

\begin{qbox}{Câu 2: Doanh thu nấu rượu gạo}
Tính thu nhập từ nấu rượu gạo: 320 kg gạo (chứa 80\% tinh bột) lên men thu rượu 40$^\circ$ với hiệu suất 70\%, $D(\text{ethanol}) = 0,8\text{ g/mL}$, giá 50.000 đ/lít. Doanh thu thu được là bao nhiêu triệu đồng? \textit{(làm tròn đến hàng phần mười)}\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 16: PHẦN III - CÂU 3 & CÂU 4 (TRẢ LỜI NGẮN)
% ==============================================================================
{\color{tvRed}\fontsize{12}{14}\selectfont\textbf{PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 3: Đếm số chất tác dụng với dung dịch NaOH nóng}
Có bao nhiêu chất trong dãy tác dụng được với dung dịch $\text{NaOH}$ nóng: lysine, aniline, methyl acrylate, ethyl formate, maltose, tripalmitin?\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\end{qbox}

\begin{qbox}{Câu 4: Số nguyên tử oxygen trong peptide}
Số nguyên tử oxygen trong phân tử tripeptide Gly-Ala-Ala là bao nhiêu?\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 17: PHẦN III - CÂU 5 & CÂU 6 (TRẢ LỜI NGẮN)
% ==============================================================================
{\color{tvRed}\fontsize{12}{14}\selectfont\textbf{PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN -- ĐỀ SỐ 1}}
\vspace{1pt}

\begin{qbox}{Câu 5: Khối lượng muối từ thủy phân phenyl acetate}
Thủy phân hoàn toàn 0,1 mol phenyl acetate $\text{CH}_3\text{COOC}_6\text{H}_5$ trong dung dịch $\text{NaOH}$ dư. Khối lượng muối thu được là bao nhiêu gam?\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\end{qbox}

\begin{qbox}{Câu 6: Khối lượng tinh bột để sản xuất glucose}
Để sản xuất 1 tấn glucose với hiệu suất 80\%, khối lượng tinh bột cần dùng là bao nhiêu tấn? \textit{(làm tròn đến hàng phần trăm)}\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\end{qbox}
\newpage

% ==============================================================================
% SLIDE 18: BẢNG ĐÁP ÁN ĐỀ SỐ 1
% ==============================================================================
\begin{center}
{\color{tvRed}\fontsize{18}{22}\selectfont\textbf{BẢNG ĐÁP ÁN CHÍNH THỨC -- ĐỀ SỐ 01}}\\[0.2cm]
\end{center}

\begin{tcolorbox}[enhanced, colback=ansBg, colframe=ansBorder, arc=3mm, boxrule=1.5pt, left=8pt, right=8pt, top=6pt, bottom=6pt]
{\color{tvBlue}\textbf{1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (4,5 ĐIỂM)}}\\[3pt]
\begin{tabularx}{\linewidth}{|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} & \textbf{17} & \textbf{18} \\
\hline
\color{tvRed}\textbf{D} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{C} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{A} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{C} & \color{tvRed}\textbf{A} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{B} & \color{tvRed}\textbf{A} \\
\hline
\end{tabularx}

\vspace{8pt}
{\color{tvCyan}\textbf{2. PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG SAI (4,0 ĐIỂM)}}\\[3pt]
\begin{tabularx}{\linewidth}{XXXX}
\textbf{Câu 1:} [ \textbf{Đ -- Đ -- S -- Đ} ] & \textbf{Câu 2:} [ \textbf{Đ -- S -- Đ -- Đ} ] & \textbf{Câu 3:} [ \textbf{Đ -- S -- Đ -- S} ] & \textbf{Câu 4:} [ \textbf{Đ -- Đ -- Đ -- S} ]
\end{tabularx}

\vspace{8pt}
{\color{tvRed}\textbf{3. PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (1,5 ĐIỂM)}}\\[3pt]
\begin{tabularx}{\linewidth}{XXX}
\textbf{Câu 1:} {\color{tvRed}\textbf{2341}} & \textbf{Câu 2:} {\color{tvRed}\textbf{12,8}} triệu đồng & \textbf{Câu 3:} {\color{tvRed}\textbf{4}} chất \\[4pt]
\textbf{Câu 4:} {\color{tvRed}\textbf{4}} nguyên tử O & \textbf{Câu 5:} {\color{tvRed}\textbf{17,6}} gam & \textbf{Câu 6:} {\color{tvRed}\textbf{1,13}} tấn
\end{tabularx}
\end{tcolorbox}

\vspace{0.3cm}
\begin{center}
{\color{tvBlue}\fontsize{13}{16}\selectfont\textbf{Thầy TRẦN VĂN THẠNH -- Chúc các em học tập và ôn thi đạt kết quả cao nhất!}}
\end{center}

\end{document}
"""

with open(TEX_FILE, "w", encoding="utf-8") as f:
    f.write(latex_content)

print(f"Đã ghi file TeX: {TEX_FILE}")

# Biên dịch pdflatex
cmd = ["pdflatex", "-interaction=nonstopmode", "-output-directory", SCRATCH_DIR, TEX_FILE]
res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

pdf_temp = os.path.join(SCRATCH_DIR, "De_On_Tap_Hoa_12_De01_A5_Tivi.pdf")
if os.path.exists(pdf_temp):
    shutil.copy2(pdf_temp, PDF_OUT)
    print(f"BIÊN DỊCH THÀNH CÔNG! File PDF: {PDF_OUT} ({os.path.getsize(PDF_OUT):,} bytes)")
    
    # Sao lưu sang OneDrive
    od1 = r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12\De_On_Tap_Hoa_12_De01_A5_NAM_NGANG_TrinhChieu_Tivi.pdf"
    try:
        shutil.copy2(PDF_OUT, od1)
        print(f"Đã sao lưu sang OneDrive: {od1}")
    except Exception as e:
        print(f"Lỗi sao lưu OneDrive: {e}")
else:
    print("LỖI BIÊN DỊCH:")
    print(res.stdout[-1500:])
