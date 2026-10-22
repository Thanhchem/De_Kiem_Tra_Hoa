# -*- coding: utf-8 -*-
"""
Script biên dịch file PDF TRÌNH CHIẾU TIVI 65 INCH (16:9 Widescreen) cho:
CHUYÊN ĐỀ: PIN ĐIỆN VÀ ĐIỆN PHÂN (TỜ SỐ 36) - HÓA HỌC 12
Thầy TRẦN VĂN THẠNH - 0777.470.803

Đúng 100% QUY CHUẨN MỤC 5 TRONG GEMINI.MD:
- paperwidth=29.7cm, paperheight=16.706cm (16:9 widescreen không viền đen)
- Căn lề: top=1.0cm, bottom=1.0cm, left=1.2cm, right=1.2cm
- Không Header/Footer: headheight=0pt, headsep=0pt, footskip=0pt, \pagestyle{empty}
- Nền trắng tinh (#FFFFFF), chữ đen tuyền (#000000), font Times New Roman, KHÔNG DÙNG KHUNG
- Cỡ chữ cực lớn 24pt - 30pt
- Slide bảng đáp án ở cuối
"""

import os
import sys
import re
import subprocess
import shutil

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.append(os.path.dirname(__file__))
import hoa12_to36_data as d

ROOT_DIR = r"c:\Antigravity_Thanh"
OUT_DIR = os.path.join(ROOT_DIR, "San_Pham")
IMG_DIR = os.path.join(OUT_DIR, "images_to36")

TEX_PATH = os.path.join(OUT_DIR, "De_To36_Hoa12_PinDien_DienPhan_TrinhChieu_Tivi_16x9.tex")
PDF_PATH = os.path.join(OUT_DIR, "De_To36_Hoa12_PinDien_DienPhan_TrinhChieu_Tivi_16x9.pdf")

def to_latex_text(text):
    """Chuyển đổi chuỗi text giàu (rich text) sang mã LaTeX an toàn."""
    if not text:
        return ""
    
    # Escape các ký tự đặc biệt của LaTeX
    s = text.replace("&", r"\&").replace("%", r"\%")
    
    # Subscript: ~abc~ -> $_{\text{abc}}$
    s = re.sub(r'~([^~]+)~', r'$_{\\text{\1}}$', s)
    # Superscript: ^abc^ -> $^{\text{abc}}$
    s = re.sub(r'\^([^\^]+)\^', r'$^{\\text{\1}}$', s)
    # Bold: *abc* -> \textbf{abc}
    s = re.sub(r'\*([^\*]+)\*', r'\\textbf{\1}', s)
    # Italic: _abc_ -> \textit{abc}
    s = re.sub(r'_([^_]+)_', r'\\textit{\1}', s)
    
    # Chuyển đổi dấu xuống dòng
    s = s.replace("\n", r"\\ " + "\n")
    return s

def build_latex_document():
    lines = []
    
    # Preamble
    lines.append(r"\documentclass{article}")
    lines.append(r"\usepackage[utf8]{inputenc}")
    lines.append(r"\usepackage[vietnamese,english]{babel}")
    lines.append(r"\usepackage{times}")
    lines.append(r"\usepackage{amsmath,amssymb}")
    lines.append(r"\usepackage{tabularx}")
    lines.append(r"\usepackage{array}")
    lines.append(r"\usepackage{booktabs}")
    lines.append(r"\usepackage{graphicx}")
    lines.append(r"\usepackage{xcolor}")
    lines.append(r"\usepackage[")
    lines.append(r"  paperwidth=29.7cm,")
    lines.append(r"  paperheight=16.706cm,")
    lines.append(r"  top=1.0cm,")
    lines.append(r"  bottom=1.0cm,")
    lines.append(r"  left=1.2cm,")
    lines.append(r"  right=1.2cm,")
    lines.append(r"  footskip=0pt,")
    lines.append(r"  headheight=0pt,")
    lines.append(r"  headsep=0pt")
    lines.append(r"]{geometry}")
    lines.append(r"\pagestyle{empty}")
    lines.append(r"\setlength{\parindent}{0pt}")
    lines.append(r"\setlength{\parskip}{0pt}")
    lines.append("")
    lines.append(r"% Màu sắc")
    lines.append(r"\definecolor{brandGreen}{RGB}{46, 125, 50}")
    lines.append(r"\definecolor{brandRed}{RGB}{211, 47, 47}")
    lines.append(r"\definecolor{brandBlue}{RGB}{21, 101, 192}")
    lines.append("")
    lines.append(r"\begin{document}")
    lines.append(r"\selectlanguage{vietnamese}")
    lines.append("")

    # =========================================================================
    # SLIDE 1: TRANG TIÊU ĐỀ
    # =========================================================================
    lines.append(r"% SLIDE 1: TIÊU ĐỀ")
    lines.append(r"\begin{center}")
    lines.append(r"\vspace*{0.8cm}")
    lines.append(r"{\fontsize{24}{30}\selectfont \textbf{TRUNG TÂM LUYỆN THI HÓA HỌC THẦY THẠNH}}\\[0.6cm]")
    lines.append(r"{\fontsize{34}{42}\selectfont \textbf{\color{brandGreen}CHUYÊN ĐỀ: PIN ĐIỆN VÀ ĐIỆN PHÂN (TỜ SỐ 36)}}\\[0.4cm]")
    lines.append(r"{\fontsize{28}{36}\selectfont \textbf{MÔN: HÓA HỌC 12 -- CHƯƠNG TRÌNH MỚI 2026 - 2027}}\\[0.6cm]")
    lines.append(r"{\fontsize{22}{28}\selectfont \textit{Hệ thống lý thuyết chuyên sâu: Mạ điện -- Anodizing -- Điện phân $Al_2O_3$ -- Tinh chế đồng}}\\[1.0cm]")
    lines.append(r"{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}Thầy TRẦN VĂN THẠNH -- 0777.470.803}}\\[0.3cm]")
    lines.append(r"{\fontsize{20}{26}\selectfont CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế \quad | \quad CS2: 24 Đặng Thái Thân, TP Huế}")
    lines.append(r"\end{center}")
    lines.append(r"\newpage")
    lines.append("")

    # =========================================================================
    # CÁC SLIDE BÀI TẬP VÀ LÝ THUYẾT
    # =========================================================================
    # 1. Dạng 1 (Câu 1 đến Câu 9)
    for q in d.EXERCISES_DANG1:
        lines.append(f"% DẠNG 1: CÂU {q['num']}")
        lines.append(r"\fontsize{24}{32}\selectfont")
        content_tex = to_latex_text(q['content'])
        lines.append(f"\\textbf{{Câu {q['num']} (Mạ kim loại).}} {content_tex}")
        lines.append(r"\vspace{12pt}")
        lines.append(r"\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$")
        lines.append(r"\vspace*{\fill}")
        lines.append(r"\newpage")
        lines.append("")

    # 2. Dạng 2 (Câu 1 đến Câu 3)
    for q in d.EXERCISES_DANG2:
        lines.append(f"% DẠNG 2: CÂU {q['num']}")
        lines.append(r"\fontsize{24}{32}\selectfont")
        content_tex = to_latex_text(q['content'])
        lines.append(f"\\textbf{{Câu {q['num']} (Anode hóa nhôm).}} {content_tex}")
        lines.append(r"\vspace{12pt}")
        lines.append(r"\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$")
        lines.append(r"\vspace*{\fill}")
        lines.append(r"\newpage")
        lines.append("")

    # 3. Dạng 3 (Câu 1 đến Câu 11)
    for q in d.EXERCISES_DANG3:
        lines.append(f"% DẠNG 3: CÂU {q['num']}")
        q_type = q.get('type')
        if q_type == 'mcq':
            lines.append(r"\fontsize{24}{32}\selectfont")
            content_tex = to_latex_text(q['content'])
            lines.append(f"\\textbf{{Câu {q['num']}.}} {content_tex}")
            lines.append(r"\vspace{10pt}")
            opts = q.get('options', [])
            lines.append(r"\begin{tabularx}{\linewidth}{@{}X X@{}}")
            lines.append(f"{to_latex_text(opts[0])} & {to_latex_text(opts[1])}\\\\[8pt]")
            lines.append(f"{to_latex_text(opts[2])} & {to_latex_text(opts[3])}")
            lines.append(r"\end{tabularx}")
        elif q_type == 'tf':
            lines.append(r"\fontsize{22}{29}\selectfont")
            content_tex = to_latex_text(q['content'])
            lines.append(f"\\textbf{{Câu {q['num']}.}} {content_tex}")
            lines.append(r"\vspace{6pt}")
            items = q.get('items', [])
            for it in items:
                lines.append(f"\\textbf{{{it['id']})}} [\\quad] {to_latex_text(it['text'])}\\\\[4pt]")
        else: # essay
            lines.append(r"\fontsize{24}{32}\selectfont")
            content_tex = to_latex_text(q['content'])
            lines.append(f"\\textbf{{Câu {q['num']}.}} {content_tex}")
            lines.append(r"\vspace{12pt}")
            lines.append(r"\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$")
        lines.append(r"\vspace*{\fill}")
        lines.append(r"\newpage")
        lines.append("")

    # 4. Dạng 4 (Câu 1 đến Câu 8)
    for q in d.EXERCISES_DANG4:
        lines.append(f"% DẠNG 4: CÂU {q['num']}")
        q_type = q.get('type')
        if q_type == 'tf':
            lines.append(r"\fontsize{22}{29}\selectfont")
            content_tex = to_latex_text(q['content'])
            lines.append(f"\\textbf{{Câu {q['num']}.}} {content_tex}")
            lines.append(r"\vspace{6pt}")
            items = q.get('items', [])
            for it in items:
                lines.append(f"\\textbf{{{it['id']})}} [\\quad] {to_latex_text(it['text'])}\\\\[4pt]")
        else:
            lines.append(r"\fontsize{24}{32}\selectfont")
            content_tex = to_latex_text(q['content'])
            lines.append(f"\\textbf{{Câu {q['num']}.}} {content_tex}")
            lines.append(r"\vspace{12pt}")
            lines.append(r"\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$")
        lines.append(r"\vspace*{\fill}")
        lines.append(r"\newpage")
        lines.append("")

    # 5. Bài tập tổng hợp (Câu 1 đến Câu 36)
    # Lặp qua các câu hỏi tổng hợp
    for q in d.TONG_HOP_MCQ_TF:
        lines.append(f"% TỔNG HỢP: CÂU {q['num']}")
        q_type = q.get('type')
        if q_type == 'mcq':
            lines.append(r"\fontsize{24}{32}\selectfont")
            content_tex = to_latex_text(q['content'])
            lines.append(f"\\textbf{{Câu {q['num']}.}} {content_tex}")
            lines.append(r"\vspace{10pt}")
            opts = q.get('options', [])
            if len(opts) == 4:
                lines.append(r"\begin{tabularx}{\linewidth}{@{}X X@{}}")
                lines.append(f"{to_latex_text(opts[0])} & {to_latex_text(opts[1])}\\\\[8pt]")
                lines.append(f"{to_latex_text(opts[2])} & {to_latex_text(opts[3])}")
                lines.append(r"\end{tabularx}")
        elif q_type == 'tf':
            lines.append(r"\fontsize{22}{29}\selectfont")
            content_tex = to_latex_text(q['content'])
            lines.append(f"\\textbf{{Câu {q['num']}.}} {content_tex}")
            lines.append(r"\vspace{6pt}")
            items = q.get('items', [])
            for it in items:
                lines.append(f"\\textbf{{{it['id']})}} [\\quad] {to_latex_text(it['text'])}\\\\[4pt]")
        lines.append(r"\vspace*{\fill}")
        lines.append(r"\newpage")
        lines.append("")

    # 6. Phần III. Trả lời ngắn (Câu 37 đến Câu 74)
    for q in d.PHAN3_SHORT_ANSWER:
        lines.append(f"% PHẦN 3: CÂU {q['num']}")
        q_type = q.get('type')
        if q_type == 'tf':
            lines.append(r"\fontsize{22}{29}\selectfont")
            content_tex = to_latex_text(q['content'])
            lines.append(f"\\textbf{{Câu {q['num']}.}} {content_tex}")
            lines.append(r"\vspace{6pt}")
            items = q.get('items', [])
            for it in items:
                lines.append(f"\\textbf{{{it['id']})}} [\\quad] {to_latex_text(it['text'])}\\\\[4pt]")
        else:
            lines.append(r"\fontsize{24}{32}\selectfont")
            content_tex = to_latex_text(q['content'])
            lines.append(f"\\textbf{{Câu {q['num']}.}} {content_tex}")
            lines.append(r"\vspace{12pt}")
            lines.append(r"\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$")
        lines.append(r"\vspace*{\fill}")
        lines.append(r"\newpage")
        lines.append("")

    # =========================================================================
    # SLIDE BẢNG ĐÁP ÁN NHANH Ở CUỐI
    # =========================================================================
    lines.append(r"% SLIDE: BẢNG ĐÁP ÁN")
    lines.append(r"\begin{center}")
    lines.append(r"{\fontsize{26}{34}\selectfont \textbf{\color{brandGreen}BẢNG ĐÁP ÁN NHANH BÀI TẬP TỔNG HỢP}}\\[10pt]")
    lines.append(r"\fontsize{18}{24}\selectfont")
    lines.append(r"\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|}")
    lines.append(r"\hline")
    # In hàng 1-12
    hdr1 = " & ".join([f"\\textbf{{{i}}}" for i in range(1, 13)])
    lines.append(hdr1 + r"\\")
    lines.append(r"\hline")
    ans1 = []
    for i in range(1, 13):
        item = d.TONG_HOP_MCQ_TF[i-1]
        if item['type'] == 'mcq':
            ans1.append(f"\\textbf{{{item.get('ans', '')}}}")
        else:
            sub = "".join([it['ans'] for it in item.get('items', [])])
            ans1.append(f"\\textbf{{{sub}}}")
    lines.append(" & ".join(ans1) + r"\\")
    lines.append(r"\hline")
    # In hàng 13-24
    hdr2 = " & ".join([f"\\textbf{{{i}}}" for i in range(13, 25)])
    lines.append(hdr2 + r"\\")
    lines.append(r"\hline")
    ans2 = []
    for i in range(13, 25):
        item = d.TONG_HOP_MCQ_TF[i-1]
        if item['type'] == 'mcq':
            ans2.append(f"\\textbf{{{item.get('ans', '')}}}")
        else:
            sub = "".join([it['ans'] for it in item.get('items', [])])
            ans2.append(f"\\textbf{{{sub}}}")
    lines.append(" & ".join(ans2) + r"\\")
    lines.append(r"\hline")
    # In hàng 25-36
    hdr3 = " & ".join([f"\\textbf{{{i}}}" for i in range(25, 37)])
    lines.append(hdr3 + r"\\")
    lines.append(r"\hline")
    ans3 = []
    for i in range(25, 37):
        item = d.TONG_HOP_MCQ_TF[i-1]
        if item['type'] == 'mcq':
            ans3.append(f"\\textbf{{{item.get('ans', '')}}}")
        else:
            sub = "".join([it['ans'] for it in item.get('items', [])])
            ans3.append(f"\\textbf{{{sub}}}")
    lines.append(" & ".join(ans3) + r"\\")
    lines.append(r"\hline")
    lines.append(r"\end{tabular}")
    lines.append(r"\end{center}")
    lines.append(r"\vspace*{\fill}")
    lines.append(r"\newpage")
    lines.append("")

    lines.append(r"\end{document}")
    
    with open(TEX_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Generated LaTeX: {TEX_PATH}")

def compile_pdf():
    print("Compiling LaTeX to PDF via pdflatex...")
    cmd = [
        "pdflatex",
        "-interaction=nonstopmode",
        f"-output-directory={OUT_DIR}",
        TEX_PATH
    ]
    # Run twice for references
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=OUT_DIR)
    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=OUT_DIR)
    
    # Dọn dẹp rác
    for ext in [".aux", ".log", ".out", ".thm", ".synctex.gz"]:
        waste = os.path.join(OUT_DIR, "De_To36_Hoa12_PinDien_DienPhan_TrinhChieu_Tivi_16x9" + ext)
        if os.path.exists(waste):
            try:
                os.remove(waste)
            except Exception:
                pass
                
    if os.path.exists(PDF_PATH):
        print(f"SUCCESS: Generated 16:9 TV PDF: {PDF_PATH} ({os.path.getsize(PDF_PATH):,} bytes)")
        # Copy to OneDrive
        od_dirs = [
            r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12",
            r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12\kiem-tra-dinh-ky"
        ]
        for od in od_dirs:
            if os.path.exists(od):
                shutil.copy2(PDF_PATH, os.path.join(od, os.path.basename(PDF_PATH)))
                print(f"Copied TV PDF to OneDrive: {od}")
    else:
        print("FAILED to compile PDF. Check log.")

if __name__ == "__main__":
    build_latex_document()
    compile_pdf()
