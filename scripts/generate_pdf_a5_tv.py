# -*- coding: utf-8 -*-
"""
Script tạo file PDF khổ A5 NẰM NGANG (LANDSCAPE) CHUYÊN DỤNG TRÌNH CHIẾU TIVI 65 INCH
- Khổ giấy: A5 NẰM NGANG (210mm x 148mm)
- Cỡ chữ to hơn: Sử dụng \\documentclass[14pt]{extarticle}
- Giới hạn: MỖI TRANG A5 KHÔNG QUÁ 2 CÂU (Tối đa 2 câu/slide, câu dài 1 câu/slide)
- KHÔNG header, KHÔNG footer (đúng yêu cầu người dùng)
- Tối ưu kích thước chữ, khoảng cách, hình ảnh để học sinh ngồi xa đọc rõ trên màn hình lớn.
- Bố trí minipage chống tràn trang, câu hỏi nguyên vẹn trên từng slide trình chiếu.
- Trang bìa mở đầu slide đẹp mắt, trang kết thúc chúc thi tốt trang trọng.
- Tạo 2 file: Bản Học Sinh (đề bài) và Bản Giáo Viên (có đáp án nổi bật).
"""

import os
import re
import math
import subprocess
import shutil
import pymupdf
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Import dữ liệu
sys.path.append(os.path.dirname(__file__))
from khtn8_data_part1 import ALL_PART1
from khtn8_data_part2_3_4 import PART2_INFO, PART3_INFO, PART4_ESSAY

SAN_PHAM_DIR = r"c:\Antigravity_Thanh\San_Pham"
SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
IMG_BALLOON = r"c:/Antigravity_Thanh/San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay.png"
IMG_WRENCHES = r"c:/Antigravity_Thanh/San_Pham/Thi_Nghiem_Co_Le_Bu_Long/thi_nghiem_co_le_bu_long.png"

def tex_escape(text):
    """Chuyển đổi văn bản thông thường sang cú pháp LaTeX chuẩn."""
    if not text:
        return ""
    t = text
    t = t.replace('\\', '\\textbackslash{}')
    t = t.replace('&', r'\&')
    t = t.replace('%', r'\%')
    t = t.replace('#', r'\#')
    t = t.replace('$', r'\$')
    t = t.replace('_', r'\_')
    
    # Chỉ số trên ^...^
    t = re.sub(r'\^([^^]+)\^', r'\\textsuperscript{\1}', t)
    # Chỉ số dưới ~...~
    t = re.sub(r'~([^~]+)~', r'\\textsubscript{\1}', t)
    
    # Chữ in đậm *...*
    t = re.sub(r'\*([^*]+)\*', r'\\textbf{\1}', t)
    
    # Mũi tên và kí hiệu toán học / hóa học
    t = t.replace('→', r'$\rightarrow$')
    t = t.replace('->', r'$\rightarrow$')
    t = t.replace('↑', r'$\uparrow$')
    t = t.replace('≈', r'$\approx$')
    t = t.replace('≤', r'$\le$')
    t = t.replace('≥', r'$\ge$')
    t = t.replace('·', r'$\cdot$')
    t = t.replace('•', r'\textbullet{}')
    
    # Xuống dòng
    t = t.replace('\n', '\n\n')
    return t

def clean_opt_text(opt_str):
    """Lấy nhãn A/B/C/D, nội dung và trạng thái đáp án đúng."""
    s = opt_str.strip()
    is_correct = False
    if s.startswith('*') and s.endswith('*'):
        is_correct = True
        s = s[1:-1].strip()
    m = re.match(r'^([A-D]\.\s*)(.*)$', s)
    if m:
        lbl = m.group(1).strip()
        body = m.group(2).strip()
    else:
        lbl = ""
        body = s
    return lbl, body, is_correct

def est_part1_lines(q, is_teacher):
    """Ước lượng chiều cao số dòng của câu trắc nghiệm phần I (14pt extarticle)."""
    q_clean = re.sub(r'^Câu\s*\d+\.\s*', '', q["question"]).strip()
    q_lines = max(1, math.ceil(len(q_clean) / 72))
    
    parsed_opts = [clean_opt_text(opt) for opt in q["options"]]
    max_len = max(len(body) for lbl, body, is_corr in parsed_opts) if parsed_opts else 0
    
    if max_len <= 14 and len(parsed_opts) == 4:
        opt_lines = 1.0  # 4 cột
    elif max_len <= 36 and len(parsed_opts) == 4:
        opt_lines = 2.0  # 2 cột
    else:
        # 1 cột: mỗi phương án 1 dòng trở lên
        opt_lines = sum(max(1, math.ceil(len(body) / 65)) for lbl, body, is_corr in parsed_opts)
        
    sol_lines = 0.0
    if is_teacher and q.get("solution"):
        sol_lines = max(1, math.ceil(len(q["solution"]) / 72)) + 0.5
        
    return q_lines + opt_lines + sol_lines + 1.6

def est_part3_lines(q, is_teacher):
    """Ước lượng chiều cao số dòng của câu trả lời ngắn phần III."""
    if q.get("has_table") and q["id"] == 8:
        return 99.0
    q_clean = re.sub(r'^Câu\s*\d+\.\s*', '', q["question"]).strip()
    q_lines = max(1, math.ceil(len(q_clean) / 72))
    ans_lines = 1.5
    sol_lines = 0.0
    if is_teacher and q.get("solution"):
        sol_lines = max(1, math.ceil(len(q["solution"]) / 72)) + 0.5
    return q_lines + ans_lines + sol_lines + 1.4

def est_part4_student_lines(q):
    """Ước lượng chiều cao số dòng câu tự luận học sinh."""
    if q.get("has_image") and q["id"] == 11:
        return 99.0
    q_clean = re.sub(r'^Câu\s*\d+\.\s*', '', q["question"]).strip()
    q_lines = max(1, math.ceil(len(q_clean) / 72))
    dot_lines = 3.5
    return q_lines + dot_lines + 1.8

def build_latex_content(is_teacher=False):
    """Xây dựng mã LaTeX A5 NẰM NGANG chuẩn 14pt trình chiếu Tivi 65 inch, tối đa 2 câu/trang."""
    lines = []
    
    # Header Preamble: BẮT BUỘC A5 LANDSCAPE (NẰM NGANG), 14PT EXTARTICLE
    preamble = r"""\documentclass[14pt]{extarticle}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{graphicx}
\usepackage[table]{xcolor}
\usepackage{tcolorbox}
\usepackage{multicol}
\usepackage{tabularx}
\usepackage{array}
\usepackage{booktabs}
\usepackage[a5paper, landscape, top=0.8cm, bottom=0.8cm, left=0.9cm, right=0.9cm, footskip=0pt, headheight=0pt, headsep=0pt]{geometry}

% TẮT HOÀN TOÀN HEADER VÀ FOOTER
\pagestyle{empty}

% BẢNG MÀU CHUYÊN DỤNG TRÌNH CHIẾU TIVI
\definecolor{primaryBlue}{RGB}{13, 71, 161}
\definecolor{accentTeal}{RGB}{0, 105, 92}
\definecolor{correctRed}{RGB}{211, 47, 47}
\definecolor{boxGray}{RGB}{90, 95, 105}
\definecolor{tableHdr}{RGB}{225, 235, 250}
\definecolor{solBg}{RGB}{245, 248, 255}

\linespread{1.12}
\setlength{\parindent}{0pt}
\setlength{\parskip}{2pt}
\raggedbottom

\begin{document}
"""
    lines.append(preamble)

    # --- SLIDE 1: TRANG TIÊU ĐỀ BÌA TRÌNH CHIẾU SANG TRỌNG ---
    teacher_sub = r"\\[2.5mm]\textcolor{correctRed}{\textbf{\large [BẢN GIÁO VIÊN -- CÓ ĐÁP ÁN VÀ LỜI GIẢI CHI TIẾT]}}" if is_teacher else ""
    title_slide = r"""
\begin{center}
\vspace*{0.5cm}
{\normalsize\textbf{UBND PHƯỜNG THUỶ XUÂN --- TRƯỜNG THCS NGUYỄN THỊ MINH KHAI}}\\[3.5mm]
{\LARGE\textbf{\textcolor{primaryBlue}{ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA HỌC KÌ I}}}\\[3mm]
{\Large\textbf{MÔN: KHOA HỌC TỰ NHIÊN 8}}\\[1.5mm]
{\normalsize\textbf{NĂM HỌC 2026 -- 2027}}\\[2.5mm]
\textit{\normalsize Thời gian ôn tập: Học kì I}__TEACHER_SUB__\\[4mm]
\rule{0.65\linewidth}{0.8pt}\\[3.5mm]
\textit{\normalsize Họ và tên học sinh: \dotfill\ Lớp: \dotfill\ SBD: \dotfill}
\end{center}
\clearpage
""".replace("__TEACHER_SUB__", teacher_sub)
    lines.append(title_slide)

    # --- PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (TỐI ĐA 2 CÂU/SLIDE) ---
    lines.append(r"""
\noindent{\Large\textbf{\textcolor{primaryBlue}{A. PHẦN TRẮC NGHIỆM KHÁCH QUAN}}}
\vspace{1.5mm}

\noindent{\large\textbf{\textcolor{accentTeal}{Phần I. Câu trắc nghiệm nhiều lựa chọn}}} \\
\textit{\small Mỗi câu hỏi, học sinh chọn một phương án đúng nhất.}
\vspace{2.5mm}
""")

    p1_on_page = 0
    p1_lines_on_page = 3.5  # Khởi đầu có tiêu đề phần I

    for q in ALL_PART1:
        qid = q["id"]
        q_raw = q["question"]
        q_clean = re.sub(r'^Câu\s*\d+\.\s*', '', q_raw).strip()
        q_tex = tex_escape(q_clean)
        
        q_l = est_part1_lines(q, is_teacher)
        
        # Kiểm tra điều kiện ngắt trang: TỐI ĐA 2 CÂU/TRANG hoặc VƯỢT DÒNG CHO PHÉP
        if p1_on_page >= 2 or (p1_on_page == 1 and p1_lines_on_page + q_l > 17.5):
            lines.append(r"\clearpage")
            p1_on_page = 0
            p1_lines_on_page = 0.0

        if p1_on_page > 0:
            lines.append(r"\vspace{4mm}")

        # Bọc câu hỏi vào minipage để không bị đứt trang giữa chừng khi chiếu TV
        lines.append(r"\noindent\begin{minipage}{\linewidth}")
        lines.append(r"\textbf{\textcolor{primaryBlue}{Câu " + str(qid) + r".}} " + q_tex + r"\par")
        lines.append(r"\vspace{1.2mm}")

        # Xử lý 4 phương án A, B, C, D
        opts = q["options"]
        parsed_opts = [clean_opt_text(opt) for opt in opts]
        max_len = max(len(body) for lbl, body, is_corr in parsed_opts) if parsed_opts else 0

        def format_opt(lbl, body, is_corr):
            b_tex = tex_escape(body)
            if is_teacher and is_corr:
                return r"\textcolor{correctRed}{\textbf{" + lbl + " " + b_tex + r"}}"
            else:
                return r"\textbf{" + lbl + "} " + b_tex

        # Phân chia cột trong A5 Landscape (chiều rộng 19.2 cm)
        if max_len <= 14 and len(parsed_opts) == 4:
            w_col = r"0.25\linewidth"
            row_str = "".join([r"\makebox[" + w_col + r"][l]{" + format_opt(*opt) + r"}" for opt in parsed_opts])
            lines.append(r"\noindent " + row_str + r"\par")
        elif max_len <= 36 and len(parsed_opts) == 4:
            w_col = r"0.5\linewidth"
            r1 = r"\makebox[" + w_col + r"][l]{" + format_opt(*parsed_opts[0]) + r"}\makebox[" + w_col + r"][l]{" + format_opt(*parsed_opts[1]) + r"}"
            r2 = r"\makebox[" + w_col + r"][l]{" + format_opt(*parsed_opts[2]) + r"}\makebox[" + w_col + r"][l]{" + format_opt(*parsed_opts[3]) + r"}"
            lines.append(r"\noindent " + r1 + r"\\[1.2mm]")
            lines.append(r"\noindent " + r2 + r"\par")
        else:
            for idx, opt in enumerate(parsed_opts):
                if idx < len(parsed_opts) - 1:
                    lines.append(r"\noindent " + format_opt(*opt) + r"\\[1.2mm]")
                else:
                    lines.append(r"\noindent " + format_opt(*opt) + r"\par")

        if is_teacher and q.get("solution"):
            sol_tex = tex_escape(q["solution"])
            lines.append(r"\vspace{1mm}")
            lines.append(r"\noindent\hspace*{3mm}\textit{\small\textcolor{blue!70!black}{☞ Lời giải: " + sol_tex + r"}}\par")

        lines.append(r"\end{minipage}")
        
        p1_on_page += 1
        p1_lines_on_page += q_l

    # --- PHẦN II: ĐÚNG - SAI (MỖI CÂU 1 SLIDE RIÊNG BIỆT) ---
    lines.append(r"""
\clearpage
\noindent{\Large\textbf{\textcolor{primaryBlue}{Phần II. Câu trắc nghiệm đúng -- sai}}} \\
\textit{\normalsize Mỗi câu hỏi có 4 ý a, b, c, d --- học sinh chọn đúng hoặc sai cho mỗi ý.}
\vspace{3mm}
""")
    # Hộp YÊU CẦU CẦN ĐẠT nằm trên slide mở đầu Phần II
    req_items = PART2_INFO.get("req", [])
    if req_items:
        req_box = [r"\begin{tcolorbox}[colback=accentTeal!5!white, colframe=accentTeal!75!black, arc=1.5mm, boxrule=0.6pt, left=3.5mm, right=3.5mm, top=2.5mm, bottom=2.5mm]",
                   r"\textbf{\normalsize YÊU CẦU CẦN ĐẠT:}\\[1.5mm]"]
        for r_line in req_items[1:]:
            req_box.append(r"\textit{\small " + tex_escape(r_line) + r"}\\[1mm]")
        req_box.append(r"\end{tcolorbox}")
        lines.append("\n".join(req_box))

    for q in PART2_INFO["questions"]:
        lines.append(r"\clearpage")  # Mỗi câu Đúng/Sai 1 slide riêng biệt để chữ to và thoáng
        qid = q["id"]
        ctx_clean = re.sub(r'^Câu\s*\d+\.\s*', '', q["context"]).strip()
        lines.append(r"\noindent\textbf{\textcolor{primaryBlue}{\large Câu " + str(qid) + r".}} " + tex_escape(ctx_clean) + r"\par")
        lines.append(r"\vspace{1.5mm}")

        for item_str, is_true, exp in q["items"]:
            m_it = re.match(r'^([a-d][\.\)])\s*(.*)$', item_str)
            if m_it:
                it_lbl = m_it.group(1)
                it_body = m_it.group(2)
            else:
                it_lbl = ""
                it_body = item_str

            it_body_tex = tex_escape(it_body)
            if is_teacher:
                status_tex = r"\mbox{\textcolor{correctRed}{\textbf{[ ĐÚNG ]}}}" if is_true else r"\mbox{\textcolor{correctRed}{\textbf{[ SAI ]}}}"
                lines.append(r"\noindent\begin{minipage}[t]{0.82\linewidth}\textbf{" + it_lbl + r"} " + it_body_tex + r"\end{minipage}\hfill\begin{minipage}[t]{0.17\linewidth}\raggedleft " + status_tex + r"\end{minipage}\par")
                if exp:
                    lines.append(r"\noindent\hspace*{4mm}\textit{\footnotesize\textcolor{blue!60!black}{Giải thích: " + tex_escape(exp) + r"}}\\[1mm]")
                else:
                    lines.append(r"\vspace{1mm}")
            else:
                status_tex = r"\mbox{\textbf{(Đúng / Sai)}}"
                lines.append(r"\noindent\begin{minipage}[t]{0.78\linewidth}\textbf{" + it_lbl + r"} " + it_body_tex + r"\end{minipage}\hfill\begin{minipage}[t]{0.21\linewidth}\raggedleft " + status_tex + r"\end{minipage}\\[2mm]")

    # --- PHẦN III: TRẢ LỜI NGẮN (TỐI ĐA 2 CÂU/SLIDE) ---
    lines.append(r"""
\clearpage
\noindent{\Large\textbf{\textcolor{primaryBlue}{Phần III. Câu trắc nghiệm trả lời ngắn}}} \\
\textit{\normalsize Mỗi câu hỏi, học sinh ghi câu trả lời không quá 4 kí tự.}
\vspace{2.5mm}
""")

    p3_on_page = 0
    p3_lines_on_page = 3.5

    for q in PART3_INFO["questions"]:
        qid = q["id"]
        q_raw = q["question"]
        q_clean = re.sub(r'^Câu\s*\d+\.\s*', '', q_raw).strip()
        q_tex = tex_escape(q_clean)

        # Nếu là câu 8 có hình vẽ bình tam giác nối bóng bay và bảng ghép nối
        if q.get("has_table") and qid == 8:
            lines.append(r"\clearpage") # Cho Câu 8 riêng một trang slide nằm ngang
            p3_on_page = 0
            p3_lines_on_page = 0.0
            
            lines.append(r"\noindent\textbf{\textcolor{primaryBlue}{\large Câu " + str(qid) + r".}} " + q_tex + r"\par")
            lines.append(r"\vspace{1.5mm}")
            
            t_data = q["table_data"]
            # Bố cục A5 Ngang: Ảnh bình tam giác bóng bay bên trái, bảng ghép nối bên phải
            lines.append(r"""
\noindent
\begin{minipage}[c]{0.35\linewidth}
\centering
\includegraphics[width=\linewidth]{__IMG_BALLOON__}
\end{minipage}%
\hfill
\begin{minipage}[c]{0.63\linewidth}
\centering
\small
\begin{tabular}{|p{6.8cm}|p{3.5cm}|}
\hline
\rowcolor{tableHdr} \textbf{__H1__} & \textbf{__H2__} \\
\hline
__R11__ & __R12__ \\
\hline
__R21__ & __R22__ \\
\hline
__R31__ & __R32__ \\
\hline
__R41__ & __R42__ \\
\hline
\end{tabular}
\end{minipage}
""".replace("__IMG_BALLOON__", IMG_BALLOON)\
   .replace("__H1__", tex_escape(t_data[0][0])).replace("__H2__", tex_escape(t_data[0][1]))\
   .replace("__R11__", tex_escape(t_data[1][0])).replace("__R12__", tex_escape(t_data[1][1]))\
   .replace("__R21__", tex_escape(t_data[2][0])).replace("__R22__", tex_escape(t_data[2][1]))\
   .replace("__R31__", tex_escape(t_data[3][0])).replace("__R32__", tex_escape(t_data[3][1]))\
   .replace("__R41__", tex_escape(t_data[4][0])).replace("__R42__", tex_escape(t_data[4][1]))
            )

            lines.append(r"\vspace{2mm}")
            if is_teacher:
                ans_val = q["ans"]
                lines.append(r"\noindent\textbf{☞ Đáp án:} \textcolor{correctRed}{\fbox{\rule{0pt}{1.4ex}\textbf{ " + tex_escape(ans_val) + r" }}}")
                if q.get("solution"):
                    lines.append(r" \quad \textit{\small\textcolor{blue!60!black}{(" + tex_escape(q['solution']) + r")}}")
                lines.append(r"\par")
            else:
                lines.append(r"\noindent\textbf{☞ Đáp án:} \fbox{\rule{0pt}{1.4ex}\rule{2.2cm}{0pt} }\par")

            p3_on_page = 2  # Câu 8 chiếm trọn slide, câu tiếp theo phải sang trang mới
            p3_lines_on_page = 15.0
            continue

        q_l = est_part3_lines(q, is_teacher)
        if p3_on_page >= 2 or (p3_on_page == 1 and p3_lines_on_page + q_l > 17.5):
            lines.append(r"\clearpage")
            p3_on_page = 0
            p3_lines_on_page = 0.0

        if p3_on_page > 0:
            lines.append(r"\vspace{4mm}")

        # Các câu trả lời ngắn thông thường
        lines.append(r"\noindent\begin{minipage}{\linewidth}")
        lines.append(r"\textbf{\textcolor{primaryBlue}{Câu " + str(qid) + r".}} " + q_tex + r"\par")
        lines.append(r"\vspace{1.2mm}")
        if is_teacher:
            ans_val = q["ans"]
            lines.append(r"\noindent\textbf{☞ Đáp án:} \textcolor{correctRed}{\fbox{\rule{0pt}{1.4ex}\textbf{ " + tex_escape(ans_val) + r" }}}")
            if q.get("solution"):
                lines.append(r" \quad \textit{\small\textcolor{blue!60!black}{(" + tex_escape(q['solution']) + r")}}")
            lines.append(r"\par")
        else:
            lines.append(r"\noindent\textbf{☞ Đáp án:} \fbox{\rule{0pt}{1.4ex}\rule{2.2cm}{0pt} }\par")
        lines.append(r"\end{minipage}")
        
        p3_on_page += 1
        p3_lines_on_page += q_l

    # --- B. TỰ LUẬN (BẢN GV: 1 CÂU/SLIDE; BẢN HS: TỐI ĐA 2 CÂU/SLIDE) ---
    lines.append(r"""
\clearpage
\noindent{\Large\textbf{\textcolor{primaryBlue}{B. TỰ LUẬN}}}
\vspace{2.5mm}
""")

    if is_teacher:
        # BẢN GIÁO VIÊN: MỖI CÂU TỰ LUẬN 1 SLIDE RIÊNG BIỆT (Kèm hộp lời giải chi tiết đẹp mắt)
        for q in PART4_ESSAY:
            lines.append(r"\clearpage")
            qid = q["id"]
            q_raw = q["question"]
            q_clean = re.sub(r'^Câu\s*\d+\.\s*', '', q_raw).strip()
            q_tex = tex_escape(q_clean)

            if q.get("has_image") and qid == 11:
                # Câu 11 có hình 3 cờ lê bu lông
                sol_tex = tex_escape(q["solution"])
                lines.append(r"\noindent\textbf{\textcolor{primaryBlue}{Câu " + str(qid) + r".}} " + q_tex + r"\par")
                lines.append(r"""
\vspace{1.5mm}
\noindent
\begin{minipage}[c]{0.44\linewidth}
\centering
\includegraphics[width=\linewidth]{__IMG_WRENCHES__}
\end{minipage}%
\hfill
\begin{minipage}[c]{0.54\linewidth}
\begin{tcolorbox}[colback=solBg, colframe=primaryBlue!75, arc=1.5mm, boxrule=0.6pt, left=2mm, right=2mm, top=1.5mm, bottom=1.5mm]
\textbf{\textcolor{primaryBlue}{\small Lời giải chi tiết:}}\\[0.5mm]
{\scriptsize __SOL__}
\end{tcolorbox}
\end{minipage}
""".replace("__IMG_WRENCHES__", IMG_WRENCHES).replace("__SOL__", sol_tex))
            else:
                lines.append(r"\noindent\textbf{\textcolor{primaryBlue}{Câu " + str(qid) + r".}} " + q_tex + r"\par")
                sol_tex = tex_escape(q["solution"])
                lines.append(r"""
\vspace{1mm}
\begin{tcolorbox}[colback=solBg, colframe=primaryBlue!75, arc=1.5mm, boxrule=0.5pt, left=2.5mm, right=2.5mm, top=1mm, bottom=1mm]
\textbf{\textcolor{primaryBlue}{\small Lời giải chi tiết:}}\\[0.5mm]
{\small __SOL__}
\end{tcolorbox}
""".replace("__SOL__", sol_tex))

    else:
        # BẢN HỌC SINH: TỐI ĐA 2 CÂU/SLIDE (Câu dài 1 câu/slide)
        p4_on_page = 0
        p4_lines_on_page = 3.5

        for q in PART4_ESSAY:
            qid = q["id"]
            q_raw = q["question"]
            q_clean = re.sub(r'^Câu\s*\d+\.\s*', '', q_raw).strip()
            q_tex = tex_escape(q_clean)

            if q.get("has_image") and qid == 11:
                lines.append(r"\clearpage")
                lines.append(r"\noindent\textbf{\textcolor{primaryBlue}{Câu " + str(qid) + r".}} " + q_tex + r"\par")
                lines.append(r"""
\vspace{2mm}
\begin{center}
\includegraphics[width=13.5cm]{__IMG_WRENCHES__}
\end{center}
\vspace{2mm}
\noindent\dotfill\\
\noindent\dotfill\\
\noindent\dotfill\par
""".replace("__IMG_WRENCHES__", IMG_WRENCHES))
                p4_on_page = 2
                p4_lines_on_page = 15.0
                continue

            q_l = est_part4_student_lines(q)
            if p4_on_page >= 2 or (p4_on_page == 1 and p4_lines_on_page + q_l > 17.5):
                lines.append(r"\clearpage")
                p4_on_page = 0
                p4_lines_on_page = 0.0

            if p4_on_page > 0:
                lines.append(r"\vspace{4mm}")

            lines.append(r"\noindent\begin{minipage}{\linewidth}")
            lines.append(r"\textbf{\textcolor{primaryBlue}{Câu " + str(qid) + r".}} " + q_tex + r"\par")
            lines.append(r"""
\vspace{1.5mm}
\noindent\dotfill\\
\noindent\dotfill\\
\noindent\dotfill\par
""")
            lines.append(r"\end{minipage}")
            p4_on_page += 1
            p4_lines_on_page += q_l

    # Slide kết thúc
    lines.append(r"""
\clearpage
\vspace*{2.0cm}
\begin{center}
{\Huge\textbf{\textcolor{primaryBlue}{CHÚC CÁC EM ÔN TẬP VÀ THI TỐT!}}}\\[6mm]
\rule{0.6\linewidth}{1.2pt}
\end{center}
\end{document}
""")

    return "\n".join(lines)

def compile_latex(tex_path, output_dir):
    """Biên dịch mã LaTeX bằng pdflatex."""
    cmd = ["pdflatex", "-interaction=nonstopmode", f"-output-directory={output_dir}", tex_path]
    res = subprocess.run(cmd, capture_output=True, text=True, errors="replace")
    base = os.path.splitext(os.path.basename(tex_path))[0]
    pdf_path = os.path.join(output_dir, base + ".pdf")
    
    # Dọn dẹp file tạm
    for ext in [".aux", ".log", ".out"]:
        f_tmp = os.path.join(output_dir, base + ext)
        if os.path.exists(f_tmp):
            try:
                os.remove(f_tmp)
            except:
                pass
                
    if os.path.exists(pdf_path):
        return pdf_path, True
    else:
        print(f"Lỗi biên dịch {tex_path}:")
        print(res.stdout[-1500:])
        return None, False

def main():
    print("=== TẠO FILE PDF A5 NẰM NGANG (LANDSCAPE) TRÌNH CHIẾU TIVI 65 INCH ===")
    print("=== CỠ CHỮ 14PT EXTARTICLE, TỐI ĐA 2 CÂU/SLIDE ===")
    os.makedirs(SAN_PHAM_DIR, exist_ok=True)
    os.makedirs(SCRATCH_DIR, exist_ok=True)

    configs = [
        # (is_teacher, file_base, description)
        (False, "De_Cuong_KHTN_8_A5_NAM_NGANG_TrinhChieu_Tivi", "Bản Học Sinh (A5 NẰM NGANG - Chiếu Tivi 65 inch)"),
        (True,  "De_Cuong_KHTN_8_A5_NAM_NGANG_TrinhChieu_Tivi_GiaoVien", "Bản Giáo Viên (A5 NẰM NGANG - Có đáp án & Lời giải)")
    ]

    for is_teacher, file_base, desc in configs:
        print(f"\n--> Đang xử lý: {desc}...")
        tex_content = build_latex_content(is_teacher=is_teacher)
        tex_path = os.path.join(SCRATCH_DIR, file_base + ".tex")
        with open(tex_path, "w", encoding="utf-8") as f:
            f.write(tex_content)
            
        pdf_path, ok = compile_latex(tex_path, SCRATCH_DIR)
        if ok:
            dest_pdf = os.path.join(SAN_PHAM_DIR, file_base + ".pdf")
            shutil.copy2(pdf_path, dest_pdf)
            doc = pymupdf.open(dest_pdf)
            print(f"    ✔ Thành công! File: {dest_pdf} ({len(doc)} trang)")
            # Kiểm tra kích thước chính xác
            rect = doc[0].rect
            print(f"    Kích thước trang: {rect.width * 25.4 / 72:.1f} mm x {rect.height * 25.4 / 72:.1f} mm (NẰM NGANG CHUẨN: Rộng > Cao)")
            doc.close()
        else:
            print(f"    ✘ Thất bại khi biên dịch {file_base}")

    print("\n=== HOÀN TẤT 100% FILE PDF A5 NẰM NGANG TRÌNH CHIẾU TIVI ===")

if __name__ == "__main__":
    main()
