# -*- coding: utf-8 -*-
"""
Script tạo file PDF ĐỀ SỐ 1 chuẩn TIVI 16:9 THEO ĐÚNG HÌNH ẢNH MẪU CỦA THẦY:
- Tỉ lệ chuẩn 16:9 widescreen (paperwidth=29.7cm, paperheight=16.706cm)
- Hoàn toàn KHÔNG DÙNG KHUNG, KHÔNG DÙNG HEADER RƯỜM RÀ
- Cỡ chữ lớn (26pt/34pt), màu đen tuyền (#000000), nét đậm rõ ràng
- Phân bố 2 câu hỏi / trang cho Phần I, khoảng cách giữa 2 câu rộng rãi (\vspace*{\fill})
- Từng câu Phần II trọn vẹn 1 trang
- Mỗi trang 2 câu Phần III
- Trang cuối: Bảng đáp án chính thức
"""

import os
import sys
import subprocess
import pymupdf
import shutil

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
os.makedirs(SCRATCH_DIR, exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)

TEX_FILE = os.path.join(SCRATCH_DIR, "De01_Tivi_16_9_ChuanMau.tex")
PDF_OUT = os.path.join(OUT_DIR, "De_On_Tap_Hoa_12_De01_TrinhChieu_Tivi_16_9.pdf")
OD_OUT = r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12\De_On_Tap_Hoa_12_De01_TrinhChieu_Tivi_16_9.pdf"

latex_content = r"""\documentclass{article}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{tabularx}
\usepackage{array}
\usepackage{booktabs}
\usepackage{xcolor}
\usepackage[
  paperwidth=29.7cm,
  paperheight=16.706cm,
  top=1.0cm,
  bottom=1.0cm,
  left=1.2cm,
  right=1.2cm,
  footskip=0pt,
  headheight=0pt,
  headsep=0pt
]{geometry}

\pagestyle{empty}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}

% Màu sắc
\definecolor{hdrBlue}{RGB}{13, 71, 161}
\definecolor{subGreen}{RGB}{0, 105, 92}
\definecolor{ansRed}{RGB}{198, 40, 40}

\begin{document}
\fontsize{30}{39}\selectfont

% ==============================================================================
% SLIDE 1: CÂU 1 & CÂU 2
% ==============================================================================
\textbf{Câu 1.} Thành phần chính của chất giặt rửa tổng hợp là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{C}_{17}\text{H}_{33}\text{COOK}$. & \textbf{B.} $(\text{C}_{17}\text{H}_{33}\text{COO})_2\text{Ca}$.\\[8pt]
\textbf{C.} $\text{CH}_3\text{COONa}$. & \textbf{D.} $\text{CH}_3[\text{CH}_2]_{11}\text{C}_6\text{H}_4\text{SO}_3\text{Na}$.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 2.} Ester nào sau đây có công thức $\text{CH}_3\text{COOCH}=\text{CH}_2$?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Methyl acetate. & \textbf{B.} Vinyl acetate.\\[8pt]
\textbf{C.} Ethyl formate. & \textbf{D.} Methyl acrylate.
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 2: CÂU 3 & CÂU 4
% ==============================================================================
\textbf{Câu 3.} Hợp chất $\text{CH}_3\text{NHCH}_3$ thuộc loại amine nào?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Bậc 1. & \textbf{B.} Bậc 2.\\[8pt]
\textbf{C.} Bậc 3. & \textbf{D.} Amine thơm.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 4.} Trong phân tử cellulose, các đơn vị glucose liên kết với nhau bằng liên kết
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\alpha$-1,4-glycoside. & \textbf{B.} $\beta$-1,4-glycoside.\\[8pt]
\textbf{C.} $\alpha$-1,6-glycoside. & \textbf{D.} $\beta$-1,6-glycoside.
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 3: CÂU 5 & CÂU 6
% ==============================================================================
\textbf{Câu 5.} Amino acid là hợp chất hữu cơ trong phân tử chứa các nhóm chức
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $-\text{COOH}$ và $-\text{OH}$. & \textbf{B.} $-\text{NH}_2$ và $-\text{COOH}$.\\[8pt]
\textbf{C.} $-\text{CHO}$ và $-\text{OH}$. & \textbf{D.} $-\text{NH}_2$ và $-\text{OH}$.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 6.} Để phân biệt glucose và glycerol ở nhiệt độ thường, thuốc thử phù hợp là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Dung dịch $\text{NaOH}$. & \textbf{B.} $\text{Cu(OH)}_2$ đun nóng.\\[8pt]
\textbf{C.} Thuốc thử Tollens, $t^\circ$. & \textbf{D.} Dung dịch $\text{HCl}$.
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 4: CÂU 7 & CÂU 8
% ==============================================================================
\textbf{Câu 7.} Chất chủ yếu trong tinh bột tạo nên quá trình hồ hóa khi đun nóng trong nước là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Amylose. & \textbf{B.} Amylopectin.\\[8pt]
\textbf{C.} Glucose. & \textbf{D.} Saccharose.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 8.} Thủy phân hoàn toàn peptide bởi dung dịch kiềm nóng thu được sản phẩm là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Ester. & \textbf{B.} Các muối của $\alpha$-amino acid.\\[8pt]
\textbf{C.} Amine. & \textbf{D.} Alcohol.
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 5: CÂU 9 & CÂU 10
% ==============================================================================
\textbf{Câu 9.} Chất nào sau đây vừa phản ứng với dung dịch $\text{NaOH}$ vừa phản ứng với dung dịch $\text{HCl}$?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{H}_2\text{N-CH}_2\text{-COOH}$. & \textbf{B.} $\text{CH}_3\text{COOH}$.\\[8pt]
\textbf{C.} $\text{C}_2\text{H}_5\text{OH}$. & \textbf{D.} $\text{CH}_3\text{NH}_2$.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 10.} Phản ứng xà phòng hóa chất béo tristearin bằng dung dịch $\text{NaOH}$ thu được xà phòng và
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Ethylene glycol. & \textbf{B.} Glycerol.\\[8pt]
\textbf{C.} Ethanol. & \textbf{D.} Methanol.
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 6: CÂU 11 & CÂU 12
% ==============================================================================
\textbf{Câu 11.} Công thức phân tử của saccharose là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{C}_6\text{H}_{12}\text{O}_6$. & \textbf{B.} $\text{C}_{12}\text{H}_{22}\text{O}_{11}$.\\[8pt]
\textbf{C.} $(\text{C}_6\text{H}_{10}\text{O}_5)_n$. & \textbf{D.} $\text{C}_3\text{H}_6\text{O}_2$.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 12.} Tên gọi của dipeptide $\text{H}_2\text{N-CH}_2\text{-CO-NH-CH(CH}_3\text{)COOH}$ là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Ala-Gly. & \textbf{B.} Gly-Ala.\\[8pt]
\textbf{C.} Gly-Val. & \textbf{D.} Ala-Val.
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 7: CÂU 13 & CÂU 14
% ==============================================================================
\textbf{Câu 13.} Carbohydrate nào sau đây thuộc loại disaccharide?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Glucose. & \textbf{B.} Fructose.\\[8pt]
\textbf{C.} Saccharose. & \textbf{D.} Tinh bột.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 14.} Hydrogen hóa hoàn toàn triolein (xúc tác $\text{Ni}, t^\circ$) thu được
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Tristearin. & \textbf{B.} Tripalmitin.\\[8pt]
\textbf{C.} Glycerol. & \textbf{D.} Oleic acid.
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 8: CÂU 15 & CÂU 16
% ==============================================================================
\textbf{Câu 15.} Amine nào sau đây là arylamine (amine thơm)?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{CH}_3\text{NH}_2$. & \textbf{B.} $\text{C}_6\text{H}_5\text{NH}_2$.\\[8pt]
\textbf{C.} $\text{C}_2\text{H}_5\text{NH}_2$. & \textbf{D.} $(\text{CH}_3)_2\text{NH}$.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 16.} Ở điều kiện thường, chất nào sau đây là chất khí?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Aniline. & \textbf{B.} Methylamine.\\[8pt]
\textbf{C.} Glycine. & \textbf{D.} Glucose.
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 9: CÂU 17 & CÂU 18
% ==============================================================================
\textbf{Câu 17.} Cho dung dịch glucose tác dụng với $\text{Cu(OH)}_2/\text{NaOH}$ đun nóng tạo kết tủa màu
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Xanh lam. & \textbf{B.} Đỏ gạch.\\[8pt]
\textbf{C.} Trắng. & \textbf{D.} Vàng.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 18.} Peptide Gly-Ala có số liên kết peptide là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} 1. & \textbf{B.} 2.\\[8pt]
\textbf{C.} 3. & \textbf{D.} 0.
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 10: PHẦN II - CÂU 1
% ==============================================================================
\fontsize{26}{34}\selectfont
\textbf{Câu 1 (Đúng/Sai).} Về cấu tạo và tính chất của chất béo Trilinolein $(\text{C}_{17}\text{H}_{31}\text{COO})_3\text{C}_3\text{H}_5$:
\vspace{14pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} Công thức phân tử của trilinolein là $\text{C}_{57}\text{H}_{98}\text{O}_6$. & $[\quad]$\\[12pt]
\textbf{b)} Thủy phân hoàn toàn trong acid thu được acid béo không no. & $[\quad]$\\[12pt]
\textbf{c)} Hydrogen hóa hoàn toàn trilinolein thu được chất béo lỏng. & $[\quad]$\\[12pt]
\textbf{d)} Trilinolein làm mất màu dung dịch nước bromine. & $[\quad]$
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 11: PHẦN II - CÂU 2
% ==============================================================================
\textbf{Câu 2 (Đúng/Sai).} Cho các điểm đẳng điện pI: Glycine (6,0), Glutamic acid (3,2), Lysine (9,7):
\vspace{14pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} Tên bán hệ thống của glycine là $\alpha$-aminoacetic acid. & $[\quad]$\\[12pt]
\textbf{b)} Trong dung dịch $\text{pH} = 6,0$, glycine tồn tại dạng cation về cực âm. & $[\quad]$\\[12pt]
\textbf{c)} Ở $\text{pH} = 11,0$, lysine tồn tại ở dạng anion về cực dương. & $[\quad]$\\[12pt]
\textbf{d)} Lysine có 2 nhóm $-\text{NH}_2$ và 1 nhóm $-\text{COOH}$ trong phân tử. & $[\quad]$
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 12: PHẦN II - CÂU 3
% ==============================================================================
\textbf{Câu 3 (Đúng/Sai).} Xét peptide X có công thức Ala-Gly-Val:
\vspace{14pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} X là một tripeptide. & $[\quad]$\\[12pt]
\textbf{b)} Amino acid đầu N là Valine, đầu C là Alanine. & $[\quad]$\\[12pt]
\textbf{c)} X có phản ứng màu biuret tạo dung dịch màu tím đặc trưng. & $[\quad]$\\[12pt]
\textbf{d)} Phân tử X chứa 3 liên kết peptide. & $[\quad]$
\end{tabularx}

\newpage

% ==============================================================================
% SLIDE 13: PHẦN II - CÂU 4
% ==============================================================================
\textbf{Câu 4 (Đúng/Sai).} Về thí nghiệm ester hóa giữa $\text{CH}_3\text{COOH}$ và $\text{C}_2\text{H}_5\text{OH}$ (xúc tác $\text{H}_2\text{SO}_4$ đặc):
\vspace{14pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} Phản ứng xảy ra là phản ứng thuận nghịch. & $[\quad]$\\[12pt]
\textbf{b)} $\text{H}_2\text{SO}_4$ đặc vừa làm xúc tác vừa tăng hiệu suất bằng cách hút nước. & $[\quad]$\\[12pt]
\textbf{c)} Tách ethyl acetate bằng phương pháp chiết vì không tan, nhẹ hơn nước. & $[\quad]$\\[12pt]
\textbf{d)} Sau phản ứng thu được dung dịch hoàn toàn đồng nhất. & $[\quad]$
\end{tabularx}

\newpage

% ==============================================================================
% PHẦN III - CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (CỠ CHỮ 32, MỖI CÂU 1 SLIDE)
% ==============================================================================
\fontsize{32}{42}\selectfont

% SLIDE 14: CÂU 1
\textbf{Câu 1 (Trả lời ngắn).} Thứ tự đúng các bước thí nghiệm điều chế xà phòng: (1) Thêm $\text{NaCl}$ bão hòa nóng; (2) Cho dầu ăn + $\text{NaOH}$ vào bát sứ; (3) Đun sôi nhẹ và khuấy đều; (4) Thêm nước cất giữ thể tích không đổi.
\vspace{22pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\newpage

% SLIDE 15: CÂU 2
\textbf{Câu 2 (Trả lời ngắn).} Tính thu nhập từ nấu rượu gạo: 320 kg gạo (chứa 80\% tinh bột) lên men thu rượu 40$^\circ$ với hiệu suất 70\%, $D = 0,8\text{ g/mL}$, giá 50.000 đ/lít. Doanh thu thu được là bao nhiêu triệu đồng? \textit{(làm tròn đến hàng phần mười)}
\vspace{22pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\newpage

% SLIDE 16: CÂU 3
\textbf{Câu 3 (Trả lời ngắn).} Có bao nhiêu chất trong dãy tác dụng được với dung dịch $\text{NaOH}$ nóng: lysine, aniline, methyl acrylate, ethyl formate, maltose, tripalmitin?
\vspace{22pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\newpage

% SLIDE 17: CÂU 4
\textbf{Câu 4 (Trả lời ngắn).} Số nguyên tử oxygen trong phân tử tripeptide Gly-Ala-Ala là bao nhiêu?
\vspace{22pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\newpage

% SLIDE 18: CÂU 5
\textbf{Câu 5 (Trả lời ngắn).} Thủy phân hoàn toàn 0,1 mol phenyl acetate $\text{CH}_3\text{COOC}_6\text{H}_5$ trong dung dịch $\text{NaOH}$ dư. Khối lượng muối thu được là bao nhiêu gam?
\vspace{22pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\newpage

% SLIDE 19: CÂU 6
\textbf{Câu 6 (Trả lời ngắn).} Để sản xuất 1 tấn glucose với hiệu suất 80\%, khối lượng tinh bột cần dùng là bao nhiêu tấn? \textit{(làm tròn đến hàng phần trăm)}
\vspace{22pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\newpage

% ==============================================================================
% SLIDE 20: BẢNG ĐÁP ÁN ĐỀ SỐ 1
% ==============================================================================
\fontsize{17}{21}\selectfont
\begin{center}
{\fontsize{23}{27}\selectfont\textbf{BẢNG ĐÁP ÁN CHÍNH THỨC -- ĐỀ SỐ 1}}\\[8pt]
\end{center}

\textbf{1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (4,5 ĐIỂM)}\\[4pt]
\renewcommand{\arraystretch}{1.2}
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

\vspace{10pt}

\textbf{2. PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG SAI (4,0 ĐIỂM)}\\[4pt]
\begin{tabularx}{\linewidth}{@{}*{4}{>{\raggedright\arraybackslash}X}@{}}
\textbf{Câu 1:} [\color{ansRed}\textbf{Đ--Đ--S--Đ}\color{black}] &
\textbf{Câu 2:} [\color{ansRed}\textbf{Đ--S--Đ--Đ}\color{black}] &
\textbf{Câu 3:} [\color{ansRed}\textbf{Đ--S--Đ--S}\color{black}] &
\textbf{Câu 4:} [\color{ansRed}\textbf{Đ--Đ--Đ--S}\color{black}]
\end{tabularx}

\vspace{10pt}

\textbf{3. PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (1,5 ĐIỂM)}\\[4pt]
\begin{tabularx}{\linewidth}{@{}XXX@{}}
\textbf{Câu 1:} {\color{ansRed}\textbf{2341}} & \textbf{Câu 2:} {\color{ansRed}\textbf{12,8}} & \textbf{Câu 3:} {\color{ansRed}\textbf{4}} \\[6pt]
\textbf{Câu 4:} {\color{ansRed}\textbf{4}} & \textbf{Câu 5:} {\color{ansRed}\textbf{17,6}} & \textbf{Câu 6:} {\color{ansRed}\textbf{1,13}}
\end{tabularx}

\end{document}
"""

with open(TEX_FILE, "w", encoding="utf-8") as f:
    f.write(latex_content)

print(f"Đã ghi file TeX: {TEX_FILE}")
res = subprocess.run(["pdflatex", "-interaction=nonstopmode", "-output-directory", SCRATCH_DIR, TEX_FILE], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

pdf_temp = os.path.join(SCRATCH_DIR, "De01_Tivi_16_9_ChuanMau.pdf")
if os.path.exists(pdf_temp):
    shutil.copy2(pdf_temp, PDF_OUT)
    print(f"BIÊN DỊCH THÀNH CÔNG: {PDF_OUT}")
    
    # Check page count
    doc = pymupdf.open(PDF_OUT)
    print(f"Tổng số trang: {len(doc)}")
    
    # Save slide 1 as image to compare
    page1 = doc[0]
    pix1 = page1.get_pixmap(dpi=150)
    pix1.save(os.path.join(SCRATCH_DIR, "slide_01_check.png"))
    print("Đã render slide 1 thành ảnh:", pix1.width, pix1.height)

    # Sync to OneDrive
    try:
        shutil.copy2(PDF_OUT, OD_OUT)
        print(f"Đã đồng bộ sang OneDrive: {OD_OUT}")
    except Exception as e:
        print(f"Lỗi OneDrive: {e}")
else:
    print("LỖI BIÊN DỊCH:")
    print(res.stdout[-1500:])
