# -*- coding: utf-8 -*-
r"""
Script biên dịch file PDF TRÌNH CHIẾU TIVI 65 INCH (16:9 Widescreen)
HƯỚNG DẪN ÔN TẬP GIỮA HỌC KÌ I - MÔN KHTN 9 (2026 - 2027)
TRƯỜNG THCS TRẦN CAO VÂN - HUẾ
Thầy TRẦN VĂN THẠNH - 0777.470.803

Đúng 100% QUY CHUẨN MỤC 5 TRONG GEMINI.MD:
- paperwidth=29.7cm, paperheight=16.706cm (16:9 widescreen không viền đen)
- Căn lề: top=1.0cm, bottom=1.0cm, left=1.2cm, right=1.2cm
- Không Header/Footer: headheight=0pt, headsep=0pt, footskip=0pt, \pagestyle{empty}
- Nền trắng tinh (#FFFFFF), chữ đen tuyền (#000000), font Times New Roman, KHÔNG DÙNG KHUNG (no tcolorbox)
- Nhãn câu in đậm \textbf{Câu X.}
- Cỡ chữ cực lớn 26pt - 32pt
- Slide bảng đáp án ở cuối
"""

import os
import sys
import subprocess
import shutil

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

ROOT_DIR = r"c:\Antigravity_Thanh"
OUT_DIR = os.path.join(ROOT_DIR, "San_Pham")
IMG_CON_LAC = os.path.join(OUT_DIR, "images_khtn9_trancao_van", "con_lac_hd.png").replace("\\", "/")

TEX_PATH = os.path.join(OUT_DIR, "De_Cuong_KHTN_9_Giua_HK1_TranCaoVan_TrinhChieu_Tivi_16x9.tex")
PDF_PATH = os.path.join(OUT_DIR, "De_Cuong_KHTN_9_Giua_HK1_TranCaoVan_TrinhChieu_Tivi_16x9.pdf")

latex_code = r"""\documentclass{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese,english]{babel}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{tabularx}
\usepackage{array}
\usepackage{booktabs}
\usepackage{graphicx}
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

% Định nghĩa màu sắc thương hiệu
\definecolor{brandGreen}{RGB}{46, 125, 50}
\definecolor{brandRed}{RGB}{211, 47, 47}
\definecolor{brandBlue}{RGB}{21, 101, 192}

\begin{document}
\selectlanguage{vietnamese}

% ==============================================================================
% SLIDE 1: TRANG TIÊU ĐỀ
% ==============================================================================
\begin{center}
\vspace*{0.6cm}
{\fontsize{24}{30}\selectfont \textbf{TRƯỜNG THCS TRẦN CAO VÂN -- TP HUẾ}}\\[0.6cm]
{\fontsize{36}{44}\selectfont \textbf{\color{brandGreen}HƯỚNG DẪN ÔN TẬP GIỮA HỌC KÌ I}}\\[0.4cm]
{\fontsize{30}{38}\selectfont \textbf{MÔN: KHOA HỌC TỰ NHIÊN 9 -- NĂM HỌC 2026 - 2027}}\\[1.0cm]
{\fontsize{26}{34}\selectfont \textit{Phân môn: Năng lượng (Vật lí) -- Chất và sự biến đổi chất (Hóa học) -- Vật sống (Sinh học)}}\\[1.2cm]
{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}Thầy TRẦN VĂN THẠNH -- 0777.470.803}}\\[0.3cm]
{\fontsize{20}{26}\selectfont CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế \quad | \quad CS2: 24 Đặng Thái Thân, TP Huế}
\end{center}
\newpage

% ==============================================================================
% PHẦN I: PHÂN MÔN NĂNG LƯỢNG (VẬT LÍ)
% ==============================================================================
\fontsize{28}{36}\selectfont
\textbf{\color{brandGreen}A. PHÂN MÔN: NĂNG LƯỢNG (VẬT LÍ)}
\vspace{14pt}

\textbf{Câu 1.} Một tòa nhà cao 10 tầng, mỗi tầng cao $3{,}4\text{ m}$, có một thang chở tối đa được 20 người, mỗi người có khối lượng trung bình $50\text{ kg}$. Mỗi chuyến lên tầng 10, nếu không dừng ở các tầng khác, mất một phút. Công suất tối thiểu của động cơ thang máy phải là bao nhiêu?
\vspace{14pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace*{\fill}
\newpage

% SLIDE 3: VẬT LÍ CÂU 2 & CÂU 4
\fontsize{28}{36}\selectfont
\textbf{Câu 2.} Một máy nâng cần đưa 100 viên gạch, mỗi viên nặng $1{,}5\text{ kg}$ lên cao $5\text{ m}$. Tổng công mà máy cần thực hiện là bao nhiêu Jun?
\vspace{8pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\vspace*{\fill}

\textbf{Câu 4.} Một người kéo đều một vật từ giếng sâu $12\text{ m}$ trong nửa phút. Người ấy phải dùng một lực $F = 150\text{ N}$. Công và công suất của người kéo là bao nhiêu?
\vspace{8pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace*{\fill}
\newpage

% SLIDE 4: VẬT LÍ CÂU 3 & CÂU 5
\fontsize{26}{34}\selectfont
\textbf{Câu 3.} Một học sinh thả một vật có khối lượng $0{,}1\text{ kg}$ ở độ cao $4{,}5\text{ m}$ rơi tự do xuống đất. Bỏ qua mọi lực cản ($g = 10\text{ m/s}^2$). Các nhận định sau đúng hay sai?
\vspace{4pt}

a) Tốc độ của vật khi vật vừa gần chạm đất là $4\text{ m/s}$. \hfill $[\quad]$\\[4pt]
b) Khi vật ở độ cao $1{,}5\text{ m}$ thì động năng gấp 2 lần thế năng. \hfill $[\quad]$

\vspace*{\fill}

\textbf{Câu 5.} Con ngựa kéo xe chuyển động đều với tốc độ $9\text{ km/h}$. Lực kéo của ngựa là $200\text{ N}$. Công suất của ngựa là bao nhiêu?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X X X@{}}
\textbf{A.} $6480\text{ J}$. & \textbf{B.} $6550\text{ J}$. & \textbf{C.} $6480\text{ kJ}$. & \textbf{D.} $6550\text{ kJ}$.
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 5: VẬT LÍ CÂU 6
\fontsize{26}{34}\selectfont
\textbf{Câu 6.} Hai vật ở cùng một độ cao so với gốc thế năng, biết khối lượng của vật thứ nhất gấp 3 lần khối lượng của vật thứ hai ($m_1 = 3m_2$). Các nhận định sau là đúng hay sai?
\vspace{12pt}

a) Thế năng trọng trường của vật thứ nhất gấp 3 lần của vật thứ hai. \hfill $[\quad]$\\[8pt]
b) Thế năng trọng trường của vật thứ hai gấp 3 lần của vật thứ nhất. \hfill $[\quad]$\\[8pt]
c) Cơ năng của vật thứ nhất gấp 3 lần cơ năng của vật thứ hai. \hfill $[\quad]$\\[8pt]
d) Cơ năng của vật thứ hai gấp 3 lần cơ năng của vật thứ nhất. \hfill $[\quad]$

\vspace*{\fill}
\newpage

% SLIDE 6: VẬT LÍ CÂU 7
\fontsize{28}{36}\selectfont
\textbf{Câu 7.} Một công nhân vác một bao xi măng có trọng lượng $500\text{ N}$ trên vai, đứng trên sân thượng toà nhà cao $20\text{ m}$ so với mặt đất. Độ cao của bao xi măng so với mặt sân thượng là $1{,}4\text{ m}$. Tính thế năng trọng trường của bao xi măng:
\vspace{10pt}

a) Chọn gốc thế năng tại mặt sân thượng toà nhà.\\[6pt]
b) Chọn gốc thế năng tại mặt đất.
\vspace{14pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace*{\fill}
\newpage

% SLIDE 7: VẬT LÍ CÂU 8 (HÌNH CON LẮC)
\fontsize{23}{30}\selectfont
\textbf{Câu 8.} Một con lắc gồm vật nặng có khối lượng $2\text{ kg}$ được treo vào đầu sợi dây nhẹ, không dãn. Từ VTCB O ban đầu, vật được nâng lên $0{,}5\text{ m}$ đến điểm A rồi thả nhẹ (hình vẽ). Chọn gốc thế năng tại O. Coi cơ năng của vật không đổi ($g = 10\text{ m/s}^2$).
\vspace{6pt}

\begin{minipage}[c]{0.68\linewidth}
a) Tính cơ năng của vật tại A.\\[6pt]
b) Khi vật chuyển động từ A đến O, động năng và thế năng của nó thay đổi như thế nào? Vì sao?\\[6pt]
c) Tính tốc độ của vật khi đi qua điểm O.\\[10pt]
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\end{minipage}
\hfill
\begin{minipage}[c]{0.29\linewidth}
\centering
\includegraphics[width=0.95\linewidth]{""" + IMG_CON_LAC + r"""}
\end{minipage}
\vspace*{\fill}
\newpage

% ==============================================================================
% PHẦN II: PHÂN MÔN CHẤT VÀ SỰ BIẾN ĐỔI CHẤT (HÓA HỌC)
% ==============================================================================
\fontsize{28}{36}\selectfont
\textbf{\color{brandGreen}B. PHÂN MÔN: CHẤT VÀ SỰ BIẾN ĐỔI CHẤT (HÓA HỌC)}
\vspace{10pt}

\textbf{Trọng tâm kiến thức cần nhớ (Bài 18, 19, 20, 21):}
\vspace{8pt}

\fontsize{24}{32}\selectfont
1. \textbf{Tính chất hóa học của kim loại:} Tác dụng với phi kim ($\text{O}_2, \text{Cl}_2$), acid ($\text{HCl}, \text{H}_2\text{SO}_4$), dd muối, nước.\\[6pt]
2. \textbf{Phương pháp tách kim loại:} Điện phân nóng chảy (K, Na, Ca, Mg, Al); Nhiệt luyện (kim loại sau Al: Fe, Zn, Cu).\\[6pt]
3. \textbf{Hợp kim quan trọng:} Gang ($2\%-5\%\text{ C}$), Thép ($<2\%\text{ C}$), Duralumin (nhẹ, bền), Inox (chống gỉ).\\[6pt]
4. \textbf{Dãy HĐHH:} $\text{K, Na, Ca, Mg, Al, Zn, Fe, Pb, (H), Cu, Ag, Au}$.
\vspace*{\fill}
\newpage

% SLIDE 9: HÓA CÂU 1 & CÂU 2
\fontsize{28}{36}\selectfont
\textbf{Câu 1.} Dãy kim loại sau được sắp xếp theo chiều tăng dần mức độ hoạt động hóa học:
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{Cu, Fe, Mg, Al, K}$. & \textbf{B.} $\text{K, Cu, Al, Mg, Fe}$.\\[8pt]
\textbf{C.} $\text{K, Mg, Al, Fe, Cu}$. & \textbf{D.} $\text{Cu, Fe, Al, Mg, K}$.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 2.} Cho các kim loại $\text{Fe, Cu, Ag, Na, Mg}$. Kết luận nào sau đây là sai?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Kim loại tác dụng với $\text{H}_2\text{SO}_4$ loãng là $\text{Na, Fe, Mg}$. & \textbf{B.} Kim loại không tác dụng $\text{HCl}$ là $\text{Cu, Ag}$.\\[8pt]
\textbf{C.} Kim loại không tác dụng với $\text{O}_2$ là $\text{Ag}$. & \textbf{D.} Cả 5 kim loại không tan trong nước ở đk thường.
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 10: HÓA CÂU 3 & CÂU 4
\fontsize{28}{36}\selectfont
\textbf{Câu 3.} Ứng dụng của lưu huỳnh (sulfur) trong đời sống là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Nhựa PVC. & \textbf{B.} Lưu hóa cao su làm săm, lốp xe.\\[8pt]
\textbf{C.} Nilon. & \textbf{D.} Chất tẩy rửa.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 4.} Chlorine thường được ứng dụng để sản xuất
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Sulfuric acid. & \textbf{B.} Lõi lọc nước.\\[8pt]
\textbf{C.} Chất tẩy rửa và nhựa PVC. & \textbf{D.} Điện cực và ruột bút chì.
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 11: HÓA CÂU 5 & CÂU 6
\fontsize{28}{36}\selectfont
\textbf{Câu 5.} Phương pháp điện phân nóng chảy được dùng để điều chế dãy kim loại nào?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{Al, Fe, Ag, K}$. & \textbf{B.} $\text{Al, Na, Ca, K}$.\\[8pt]
\textbf{C.} $\text{Cu, Fe, Ag, K}$. & \textbf{D.} $\text{Al, Cu, Ag, Na}$.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 6.} Phương pháp nhiệt luyện được sử dụng để điều chế các kim loại nào sau đây?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{Al, Fe, Ag}$. & \textbf{B.} $\text{Na, Fe, K}$.\\[8pt]
\textbf{C.} $\text{Fe, Zn, Pb}$. & \textbf{D.} $\text{Zn, Fe, Au}$.
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 12: HÓA CÂU 7 & CÂU 8
\fontsize{28}{36}\selectfont
\textbf{Câu 7.} Đặc điểm tính chất nổi trội của hợp kim inox là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Dẻo và cứng. & \textbf{B.} Khó bị gỉ.\\[8pt]
\textbf{C.} Độ bền cao, dẫn nhiệt tốt. & \textbf{D.} Nhẹ và bền.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 8.} Đặc điểm tính chất nổi trội của hợp kim duralumin là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Dẻo và cứng. & \textbf{B.} Khó bị gỉ.\\[8pt]
\textbf{C.} Độ bền cao, dẫn nhiệt tốt. & \textbf{D.} Nhẹ và bền.
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 13: HÓA CÂU 9 & CÂU 10
\fontsize{28}{36}\selectfont
\textbf{Câu 9.} Gang thường được sử dụng để
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} làm chi tiết máy móc. & \textbf{B.} làm vật liệu xây dựng.\\[8pt]
\textbf{C.} chế tạo vỏ máy bay. & \textbf{D.} dao, kéo,\dots
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 10.} Cho mẩu Sodium vào nước. Hiện tượng xảy ra là:
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Chìm trong nước và sủi bọt khí. & \textbf{B.} Chạy vo tròn trên mặt nước, phát lửa.\\[8pt]
\textbf{C.} Có khí thoát ra và kết tủa trắng. & \textbf{D.} Có khí thoát ra và kết tủa xanh lam.
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 14: HÓA CÂU 11 & CÂU 12
\fontsize{28}{36}\selectfont
\textbf{Câu 11.} Cho thanh Iron vào dung dịch Copper(II) chloride. Hiện tượng xảy ra là:
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Rắn đỏ gạch bám ngoài, dd nhạt màu. & \textbf{B.} Không thấy hiện tượng gì.\\[8pt]
\textbf{C.} Có khí thoát ra và kết tủa trắng. & \textbf{D.} Có khí thoát ra và kết tủa xanh.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 12.} Hoà tan hỗn hợp $\text{Fe}$ và $\text{Cu}$ vào dung dịch $\text{HCl}$ (vừa đủ). Các sản phẩm thu được là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{FeCl}_2$ và khí $\text{H}_2$. & \textbf{B.} $\text{FeCl}_2$, $\text{Cu}$ và khí $\text{H}_2$.\\[8pt]
\textbf{C.} $\text{Cu}$ và khí $\text{H}_2$. & \textbf{D.} $\text{FeCl}_2$ và $\text{Cu}$.
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 15: HÓA CÂU 13 & CÂU 14
\fontsize{28}{36}\selectfont
\textbf{Câu 13.} Khi vỡ nhiệt kế thủy ngân thì dùng chất nào sau đây rắc lên để khử độc?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X X X@{}}
\textbf{A.} Bột Iron. & \textbf{B.} Bột sulfur. & \textbf{C.} Bột than. & \textbf{D.} Muối ăn.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 14.} Có dung dịch muối $\text{AlCl}_3$ lẫn tạp chất $\text{CuCl}_2$. Dùng kim loại nào để làm sạch?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X X X@{}}
\textbf{A.} Silver. & \textbf{B.} Magnesium. & \textbf{C.} Aluminium. & \textbf{D.} Copper.
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 16: HÓA CÂU 15 & CÂU 16
\fontsize{28}{36}\selectfont
\textbf{Câu 15.} Kim loại có nhiệt độ nóng chảy cao nhất thường dùng làm dây tóc bóng đèn là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X X X@{}}
\textbf{A.} Tungsten (W). & \textbf{B.} Copper (Cu). & \textbf{C.} Iron (Fe). & \textbf{D.} Zinc (Zn).
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 16.} Trong sản xuất gang người ta sử dụng than cốc chủ yếu để làm gì?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Tạo khí carbon monoxide (CO). & \textbf{B.} Tạo hợp chất với iron.\\[8pt]
\textbf{C.} Tạo xỉ. & \textbf{D.} Phản ứng với đá vôi.
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 17: HÓA CÂU 17 & CÂU 18
\fontsize{28}{36}\selectfont
\textbf{Câu 17.} Đặc điểm tính chất của gang là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Dẻo và cứng. & \textbf{B.} Khó bị gỉ.\\[8pt]
\textbf{C.} Độ cứng cao, chịu lực tốt. & \textbf{D.} Nhẹ và bền.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 18.} Hợp kim của sắt với $\text{Cr, Ni, C}\dots$ có tính chống gỉ cao được gọi là gì?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X X X@{}}
\textbf{A.} Thép carbon. & \textbf{B.} Gang. & \textbf{C.} Inox. & \textbf{D.} Duralumin.
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 18: HÓA CÂU 19 & CÂU 20
\fontsize{28}{36}\selectfont
\textbf{Câu 19.} Trong nhiệt luyện, người ta không sử dụng chất nào để phản ứng với oxide kim loại?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X X X@{}}
\textbf{A.} $\text{Al}$. & \textbf{B.} $\text{C}$. & \textbf{C.} $\text{CO}$. & \textbf{D.} $\text{CO}_2$.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 20.} Gang thường được sử dụng nhiều nhất để
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} làm chi tiết máy móc. & \textbf{B.} làm vật liệu xây dựng.\\[8pt]
\textbf{C.} chế tạo vỏ máy bay. & \textbf{D.} dao, kéo,\dots
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 19: HÓA ĐÚNG/SAI CÂU 1
\fontsize{26}{34}\selectfont
\textbf{Câu 21.} Cho các phát biểu sau về tính chất của kim loại, phát biểu nào đúng, phát biểu nào sai:
\vspace{12pt}

a) Vẻ sáng lấp lánh của các kim loại ở điều kiện thường được gọi là ánh kim. \hfill $[\quad]$\\[8pt]
b) Tungsten (W) làm sợi đốt bóng đèn là do kim loại này dẫn điện tốt nhất. \hfill $[\quad]$\\[8pt]
c) Kim loại thủy ngân thường được dùng trong nhiệt kế, áp kế. \hfill $[\quad]$\\[8pt]
d) Sắt là kim loại có độ cứng cao nhất trong tất cả các kim loại. \hfill $[\quad]$

\vspace*{\fill}
\newpage

% SLIDE 20: HÓA ĐÚNG/SAI CÂU 2
\fontsize{25}{33}\selectfont
\textbf{Câu 22.} Thí nghiệm với kim loại X: X đẩy được sắt ra khỏi muối $\text{FeSO}_4$ trong dung dịch; X không phản ứng với nước ở nhiệt độ thường. Các phát biểu sau đúng hay sai:
\vspace{12pt}

a) X có thể phản ứng với hơi nước ở nhiệt độ cao. \hfill $[\quad]$\\[8pt]
b) X hoạt động hoá học yếu hơn sắt. \hfill $[\quad]$\\[8pt]
c) X phản ứng được với dung dịch $\text{HCl}$ sinh ra khí $\text{H}_2$. \hfill $[\quad]$\\[8pt]
d) X là kim loại kiềm. \hfill $[\quad]$

\vspace*{\fill}
\newpage

% SLIDE 21: HÓA TRẢ LỜI NGẮN 1 & 2
\fontsize{26}{34}\selectfont
\textbf{Câu 23.} Trong gang có chứa bao nhiêu \% carbon? Thép chứa bao nhiêu \% carbon?
\vspace{8pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\vspace*{\fill}

\textbf{Câu 24.} Cho: $\text{H}_2\text{O}$ (thường), $\text{H}_2\text{O}$ (nhiệt cao), $\text{H}_2\text{SO}_4$, $\text{O}_2$, $\text{Cl}_2$, $\text{NaOH}$, $\text{NaCl}$, $\text{AlCl}_3$, $\text{AgNO}_3$. Những chất nào tác dụng được với $\text{Fe}$?
\vspace{8pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace*{\fill}
\newpage

% SLIDE 22: HÓA TRẢ LỜI NGẮN 3 & 4
\fontsize{26}{34}\selectfont
\textbf{Câu 25.} Muốn sản xuất 5 tấn thép chứa $98\%\text{ Fe}$, cần bao nhiêu tấn gang chứa $94{,}5\%\text{ Fe}$? Biết hiệu suất quá trình là $85\%$.
\vspace{8pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\vspace*{\fill}

\textbf{Câu 26.} Có các nguyên liệu: (1) Quặng sắt; (2) Quặng Cromit; (3) Quặng Bauxite; (4) Than cốc; (5) Than đá; (6) $\text{CaCO}_3$; (7) $\text{O}_2$. Những nguyên liệu dùng để luyện gang là gì?
\vspace{8pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace*{\fill}
\newpage

% SLIDE 23: HÓA TRẢ LỜI NGẮN 5 & 6
\fontsize{26}{34}\selectfont
\textbf{Câu 27.} Cho 2 ví dụ kim loại tác dụng với khí $\text{Cl}_2$ và dung dịch $\text{HCl}$ tạo cùng một muối?
\vspace{8pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\vspace*{\fill}

\textbf{Câu 28.} Cho các chất: (1) Sulfur; (2) Iron; (3) Copper; (4) Aluminium; (5) Phosphorus. Có bao nhiêu chất dùng phổ biến làm lõi dây dẫn điện?
\vspace{8pt}

\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace*{\fill}
\newpage

% ==============================================================================
% PHẦN III: PHÂN MÔN VẬT SỐNG (SINH HỌC)
% ==============================================================================
\fontsize{28}{36}\selectfont
\textbf{\color{brandGreen}C. PHÂN MÔN: VẬT SỐNG (SINH HỌC)}
\vspace{10pt}

\textbf{Câu 1.} Tên gọi đầy đủ của phân tử DNA là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} deoxyribonucleic acid. & \textbf{B.} ribonucleic acid.\\[8pt]
\textbf{C.} nucleotide. & \textbf{D.} nucleic acid.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 2.} Mục đích của phép lai phân tích nhằm xác định
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} kiểu gene, kiểu hình của cá thể trội. & \textbf{B.} kiểu hình của cá thể trội.\\[8pt]
\textbf{C.} kiểu gene của tất cả tính trạng. & \textbf{D.} kiểu gene của cá thể mang tính trạng trội.
\end{tabularx}
\vspace*{\fill}
\newpage

% SLIDE 25: SINH CÂU 3 (ĐẾM SỐ NHẬN ĐỊNH)
\fontsize{24}{32}\selectfont
\textbf{Câu 3.} Tiến hành giao phấn giữa hai dòng đậu Hà Lan hạt trơn thuần chủng và hạt nhăn thuần chủng. $\text{F}_1$ thu được $100\%$ hạt trơn. Cho $\text{F}_1$ tự thụ phấn, ở $\text{F}_2$ thu được 300 hạt trơn và 100 hạt nhăn. Có bao nhiêu nhận định đúng trong các nhận định sau?
\vspace{8pt}

(1) Hạt trơn và hạt nhăn là một cặp tính trạng tương phản.\\[4pt]
(2) Cây hạt trơn ở thế hệ P là cơ thể mang tính trạng lặn.\\[4pt]
(3) Cơ thể thuần chủng là cây hạt trơn tc hoặc cây hạt nhăn tc ở thế hệ P.\\[4pt]
(4) Hạt nhăn là tính trạng trội và chỉ biểu hiện ở cơ thể đồng hợp lặn.
\vspace{12pt}

\textbf{Đáp số:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace*{\fill}
\newpage

% SLIDE 26: SINH CÂU 4 (SƠ ĐỒ LAI ĐẬU HÀ LAN)
\fontsize{26}{34}\selectfont
\textbf{Câu 4.} Ở đậu Hà Lan, lai giữa các cá thể thuần chủng thân cao với thân thấp. $\text{F}_1$ thu được $100\%$ thân cao. $\text{F}_2$ thu được cả thân cao và thân thấp với tỉ lệ 3 thân cao : 1 thân thấp.
\vspace{8pt}

Sử dụng kí hiệu và thuật ngữ để mô tả thí nghiệm trên bằng sơ đồ lai.
\vspace{14pt}

\textbf{Bài làm:}
\vspace*{\fill}
\newpage

% SLIDE 27: SINH CÂU 5 (SƠ ĐỒ LAI CÀ CHUA)
\fontsize{26}{34}\selectfont
\textbf{Câu 5.} Ở cà chua, gen A quy định hoa đỏ trội hoàn toàn gen a hoa trắng. Cho hai cây cà chua thuần chủng khác nhau về màu sắc hoa lai với nhau.
\vspace{8pt}

a) Viết sơ đồ lai và xác định tỉ lệ kiểu gen, kiểu hình ở $\text{F}_1$.\\[6pt]
b) Cho $\text{F}_1$ tự thụ phấn, viết sơ đồ lai và xác định tỉ lệ kiểu gen, kiểu hình ở $\text{F}_2$.
\vspace{14pt}

\textbf{Bài làm:}
\vspace*{\fill}
\newpage

% SLIDE 28: SINH CÂU 6 (LAI 2 CẶP TÍNH TRẠNG)
\fontsize{24}{32}\selectfont
\textbf{Câu 6.} Ở đậu Hà Lan, allele A hạt vàng trội hoàn toàn allele a hạt xanh; allele B vỏ trơn trội hoàn toàn allele b vỏ nhăn. Hai cặp gene nằm trên 2 cặp NST khác nhau và phân li độc lập.
Lai cây thuần chủng hạt vàng, vỏ trơn với cây thuần chủng hạt xanh, vỏ nhăn thu được $\text{F}_1$. Cho $\text{F}_1$ tự thụ phấn thu được $\text{F}_2$.
\vspace{8pt}

a) Xác định kiểu gene P và kiểu gene, kiểu hình $\text{F}_1$.\\[4pt]
b) Lập sơ đồ lai từ P đến $\text{F}_2$. Xác định tỉ lệ kiểu gene và kiểu hình ở $\text{F}_2$.
\vspace{12pt}

\textbf{Bài làm:}
\vspace*{\fill}
\newpage

% SLIDE 29: SINH CÂU 7 (ĐÚNG/SAI MENĐEN)
\fontsize{25}{33}\selectfont
\textbf{Câu 7.} Các phát biểu sau đây về quy luật di truyền của Mendel là đúng hay sai:
\vspace{10pt}

a) Theo quy luật phân li, mỗi tính trạng do một cặp allele quy định; khi giảm phân tạo giao tử, các allele phân li đồng đều về các giao tử. \hfill $[\quad]$\\[8pt]
b) Phép lai phân tích là lai cơ thể lặn với cơ thể lặn khác để xác định độ thuần chủng của cơ thể lặn. \hfill $[\quad]$\\[8pt]
c) Lai 2 giống đậu thuần chủng 2 cặp tính trạng tương phản ($\text{P}_{\text{tc}}$: Vàng, trơn $\times$ Xanh, nhăn), ở $\text{F}_2$ thu được tỉ lệ kiểu hình xấp xỉ 9 : 3 : 3 : 1. \hfill $[\quad]$\\[8pt]
d) Trong phép lai một cặp tính trạng của Mendel, ở $\text{F}_2$ thu được $100\%$ cá thể có kiểu hình mang tính trạng trội. \hfill $[\quad]$

\vspace*{\fill}
\newpage

% ==============================================================================
% SLIDE 30: BẢNG TỔNG HỢP ĐÁP ÁN CHÍNH THỨC
% ==============================================================================
\begin{center}
{\fontsize{26}{34}\selectfont \textbf{\color{brandGreen}BẢNG TỔNG HỢP ĐÁP ÁN TRẮC NGHIỆM KHÁCH QUAN}}\\[10pt]

\fontsize{20}{26}\selectfont
\textbf{1. Trắc nghiệm Phân môn Hóa học (20 câu):}\\[4pt]
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{C1} & \textbf{C2} & \textbf{C3} & \textbf{C4} & \textbf{C5} & \textbf{C6} & \textbf{C7} & \textbf{C8} & \textbf{C9} & \textbf{C10} \\
\hline
\textbf{\color{brandRed}D} & \textbf{\color{brandRed}D} & \textbf{\color{brandRed}B} & \textbf{\color{brandRed}C} & \textbf{\color{brandRed}B} & \textbf{\color{brandRed}C} & \textbf{\color{brandRed}B} & \textbf{\color{brandRed}D} & \textbf{\color{brandRed}A} & \textbf{\color{brandRed}B} \\
\hline
\hline
\textbf{C11} & \textbf{C12} & \textbf{C13} & \textbf{C14} & \textbf{C15} & \textbf{C16} & \textbf{C17} & \textbf{C18} & \textbf{C19} & \textbf{C20} \\
\hline
\textbf{\color{brandRed}A} & \textbf{\color{brandRed}B} & \textbf{\color{brandRed}B} & \textbf{\color{brandRed}C} & \textbf{\color{brandRed}A} & \textbf{\color{brandRed}A} & \textbf{\color{brandRed}C} & \textbf{\color{brandRed}C} & \textbf{\color{brandRed}D} & \textbf{\color{brandRed}A} \\
\hline
\end{tabular}

\vspace{16pt}
\textbf{2. Trắc nghiệm Vật lí \& Sinh học:}\\[4pt]
\begin{tabular}{|c|c|c|}
\hline
\textbf{Vật lí -- Câu 5} & \textbf{Sinh học -- Câu 1} & \textbf{Sinh học -- Câu 2} \\
\hline
\textbf{\color{brandRed}A (500 W)} & \textbf{\color{brandRed}A} & \textbf{\color{brandRed}D} \\
\hline
\end{tabular}

\vspace{18pt}
{\fontsize{22}{28}\selectfont \textbf{\color{brandBlue}Chúc các em học sinh ôn tập thật tốt và đạt kết quả cao nhất!}}
\end{center}

\end{document}
"""

with open(TEX_PATH, "w", encoding="utf-8") as f:
    f.write(latex_code)

print("Đã ghi file LaTeX:", TEX_PATH)
print("Đang biên dịch pdflatex...")
cmd = ["pdflatex", "-interaction=nonstopmode", os.path.basename(TEX_PATH)]
res = subprocess.run(cmd, cwd=OUT_DIR, capture_output=True, text=True)

if res.returncode == 0:
    print("Biên dịch thành công:", PDF_PATH)
    print("Kích thước PDF:", os.path.getsize(PDF_PATH), "bytes")
    # Dọn dẹp file phụ
    for ext in [".aux", ".log", ".out"]:
        f_aux = os.path.splitext(TEX_PATH)[0] + ext
        if os.path.exists(f_aux):
            try:
                os.remove(f_aux)
            except Exception:
                pass
    # Copy sang OneDrive
    for od in [
        r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9\KHTN9-DEKIEMTRADINHKI",
        r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9"
    ]:
        if os.path.exists(od):
            dest = os.path.join(od, os.path.basename(PDF_PATH))
            try:
                shutil.copy2(PDF_PATH, dest)
                print("Đã sao lưu sang OneDrive:", dest)
            except Exception as e:
                print("Lỗi copy OneDrive:", e)
else:
    print("Lỗi biên dịch LaTeX:")
    print(res.stdout[-1500:])
