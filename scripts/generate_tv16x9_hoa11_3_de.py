# -*- coding: utf-8 -*-
"""
CHƯƠNG TRÌNH BIÊN SOẠN FILE PDF TRÌNH CHIẾU TIVI 16:9 CHO 3 ĐỀ KIỂM TRA HÓA HỌC 11
Mã đề: 101 (Đề 01), 102 (Đề 02), 103 (Đề 03)
Áp dụng tuyệt đối chuẩn FORM CHIẾU TIVI 16:9 (65 INCH WIDESCREEN):
- Kích thước: paperwidth=29.7cm, paperheight=16.706cm (100% tràn viền không viền đen).
- Margins: top=1.0cm, bottom=1.0cm, left=1.2cm, right=1.2cm, pagestyle=empty.
- Không khung viền, nền trắng tinh, chữ đen tuyền font Times New Roman.
- Phân bổ slide CHÍNH XÁC 16 SLIDE / ĐỀ (KHÔNG TRÀN TRANG):
  + Phần I (12 câu): 2 câu / slide (6 slide, cỡ chữ 24pt/32pt).
  + Phần II (2 câu Đúng/Sai): 1 câu / slide (2 slide, cỡ chữ 25pt/33pt).
  + Phần III (4 câu Trả lời ngắn): 1 câu / slide (4 slide, cỡ chữ 30pt/40pt).
  + Phần IV (3 câu Tự luận): 1 câu / slide (3 slide, cỡ chữ 24pt/32pt).
  + Bảng đáp án: 1 slide cuối cùng (1 slide, cỡ chữ 13.5pt/16.5pt).
- Tổng cộng đúng 16 slide / đề.
"""

import os
import sys
import re
import subprocess
import pymupdf

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.insert(0, r'c:\Antigravity_Thanh\scripts')
from generate_3_de_gk1_hoa11_chuan_matran import DE01_DATA, DE02_DATA, DE03_DATA

SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
os.makedirs(SCRATCH_DIR, exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)

def tex_clean(text):
    """Chuyển đổi cú pháp hóa học và ký tự đặc biệt sang mã LaTeX chuẩn."""
    t = text
    # Các ký tự đặc thù
    t = t.replace('Δ~r~H^o^~298~', r'$\Delta_{\text{r}}H^{\circ}_{298}$')
    t = t.replace('ΔrH°298', r'$\Delta_{\text{r}}H^{\circ}_{298}$')
    t = t.replace('ΔrH', r'$\Delta_{\text{r}}H$')
    t = t.replace('Δ', r'$\Delta$')
    t = t.replace('⇌', r'$\rightleftharpoons$')
    t = t.replace('→', r'$\rightarrow$')
    t = t.replace('↑', r'$\uparrow$')
    t = t.replace('↓', r'$\downarrow$')
    t = t.replace('≡', r'$\equiv$')
    t = t.replace('–', '--')
    t = t.replace('—', '---')
    t = t.replace('°C', r'$^\circ\text{C}$')
    t = t.replace('°', r'$^\circ$')
    t = t.replace('%', r'\%')
    t = t.replace('⁺', r'$^+$')
    t = t.replace('⁻', r'$^-$')
    t = t.replace('₀', r'$_0$')
    t = t.replace('₁', r'$_1$')
    t = t.replace('₂', r'$_2$')
    t = t.replace('₃', r'$_3$')
    t = t.replace('₄', r'$_4$')
    t = t.replace('₅', r'$_5$')

    # Chỉ số trên o
    t = t.replace('^o^', r'$^{\circ}$')

    # Chỉ số trên và dưới chuẩn
    t = re.sub(r'~([0-9a-zA-Z\+\-]+)~', r'$_{\1}$', t)
    t = re.sub(r'\^([0-9a-zA-Z\+\-]+)\^', r'$^{\1}$', t)

    # In đậm markdown
    t = re.sub(r'\*([^\*]+)\*', r'\\textbf{\1}', t)

    # Dọn dẹp $$ lặp
    t = t.replace('$$', '')
    return t

def clean_opt_text(opt_str):
    """Bỏ nhãn A. B. C. D. và dấu sao để lấy nội dung phương án."""
    s = opt_str.strip()
    if s.startswith('*') and s.endswith('*'):
        s = s[1:-1].strip()
    m = re.match(r'^[A-D]\.\s*(.*)$', s)
    if m:
        return tex_clean(m.group(1).strip())
    return tex_clean(s)

def build_latex_exam(exam_data):
    code = exam_data["exam_code"]
    exam_num = exam_data["exam_num"]

    lines = []
    # Preamble
    lines.append(r"""\documentclass{article}
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

\definecolor{ansRed}{RGB}{198, 40, 40}
\definecolor{darkGreen}{RGB}{46, 125, 50}

\begin{document}
""")

    # ==========================================================================
    # PHẦN I: 12 CÂU (6 SLIDE - 2 CÂU / SLIDE - CỠ CHỮ 24PT/32PT)
    # ==========================================================================
    p1_qs = exam_data["p1"]
    for slide_idx in range(0, 12, 2):
        q1 = p1_qs[slide_idx]
        q2 = p1_qs[slide_idx + 1]

        lines.append(r"\fontsize{24}{32}\selectfont")
        lines.append(f"% --- SLIDE PHẦN I: CÂU {q1['num']} & CÂU {q2['num']} ---")

        # Câu 1 của slide
        stem1 = tex_clean(q1["stem"])
        lines.append(f"\\textbf{{Câu {q1['num']}.}} {stem1}")
        lines.append(r"\vspace{4pt}")
        lines.append(r"\begin{tabularx}{\linewidth}{@{}X X@{}}")
        optA1 = clean_opt_text(q1["opts"][0])
        optB1 = clean_opt_text(q1["opts"][1])
        optC1 = clean_opt_text(q1["opts"][2])
        optD1 = clean_opt_text(q1["opts"][3])
        lines.append(f"\\textbf{{A.}} {optA1} & \\textbf{{B.}} {optB1}\\\\[6pt]")
        lines.append(f"\\textbf{{C.}} {optC1} & \\textbf{{D.}} {optD1}")
        lines.append(r"\end{tabularx}")

        lines.append(r"\vspace*{\fill}")

        # Câu 2 của slide
        stem2 = tex_clean(q2["stem"])
        lines.append(f"\\textbf{{Câu {q2['num']}.}} {stem2}")
        lines.append(r"\vspace{4pt}")
        lines.append(r"\begin{tabularx}{\linewidth}{@{}X X@{}}")
        optA2 = clean_opt_text(q2["opts"][0])
        optB2 = clean_opt_text(q2["opts"][1])
        optC2 = clean_opt_text(q2["opts"][2])
        optD2 = clean_opt_text(q2["opts"][3])
        lines.append(f"\\textbf{{A.}} {optA2} & \\textbf{{B.}} {optB2}\\\\[6pt]")
        lines.append(f"\\textbf{{C.}} {optC2} & \\textbf{{D.}} {optD2}")
        lines.append(r"\end{tabularx}")

        lines.append(r"\newpage")
        lines.append("")

    # ==========================================================================
    # PHẦN II: 2 CÂU ĐÚNG - SAI (2 SLIDE - 1 CÂU / SLIDE - CỠ CHỮ 25PT/33PT)
    # ==========================================================================
    p2_qs = exam_data["p2"]
    for idx, q in enumerate(p2_qs):
        lines.append(r"\fontsize{25}{33}\selectfont")
        lines.append(f"% --- SLIDE PHẦN II: CÂU {13 + idx} ---")
        stem_p2 = tex_clean(q["stem"])
        lines.append(f"\\textbf{{Câu {13 + idx} (Đúng/Sai).}} {stem_p2}")
        lines.append(r"\vspace{12pt}")

        lines.append(r"\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}")
        for it_idx, (lbl, text_it, ans_val) in enumerate(q["items"]):
            clean_it = tex_clean(text_it)
            end_row = r"\\[10pt]" if it_idx < 3 else ""
            lines.append(f"\\textbf{{{lbl})}} {clean_it} & $[\\quad]${end_row}")
        lines.append(r"\end{tabularx}")

        lines.append(r"\newpage")
        lines.append("")

    # ==========================================================================
    # PHẦN III: 4 CÂU TRẢ LỜI NGẮN (4 SLIDE - 1 CÂU / SLIDE - CỠ CHỮ 30PT/40PT)
    # ==========================================================================
    p3_qs = exam_data["p3"]
    for idx, q in enumerate(p3_qs):
        lines.append(r"\fontsize{30}{40}\selectfont")
        lines.append(f"% --- SLIDE PHẦN III: CÂU {15 + idx} ---")
        stem_p3 = tex_clean(q["stem"]).replace("\n", r"\par\vspace{4pt}")
        lines.append(f"\\textbf{{Câu {15 + idx} (Trả lời ngắn).}} {stem_p3}")
        lines.append(r"\vspace{24pt}")
        lines.append(r"\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$")

        lines.append(r"\newpage")
        lines.append("")

    # ==========================================================================
    # PHẦN IV: 3 CÂU TỰ LUẬN (3 SLIDE - 1 CÂU / SLIDE - CỠ CHỮ 24PT/32PT)
    # ==========================================================================
    p4_qs = exam_data["p4"]
    for idx, q in enumerate(p4_qs):
        lines.append(r"\fontsize{24}{32}\selectfont")
        lines.append(f"% --- SLIDE PHẦN IV: CÂU {19 + idx} ---")
        stem_p4 = tex_clean(q["stem"]).replace("\n", r"\par\vspace{4pt}")
        lines.append(f"\\textbf{{Câu {19 + idx} (Tự luận -- {q['points']}).}} {stem_p4}")
        lines.append(r"\vspace{16pt}")
        lines.append(r"\textit{(Thí sinh trình bày chi tiết lời giải vào giấy làm bài)}")

        lines.append(r"\newpage")
        lines.append("")

    # ==========================================================================
    # SLIDE 16: BẢNG ĐÁP ÁN CHÍNH THỨC (CỠ CHỮ 13.5PT/16.5PT - NẰM TRỌN 1 SLIDE)
    # ==========================================================================
    lines.append(r"""% --- SLIDE BẢNG ĐÁP ÁN CHÍNH THỨC ---
\fontsize{13.5}{16.5}\selectfont
\begin{center}
{\fontsize{20}{24}\selectfont\textbf{BẢNG ĐÁP ÁN CHÍNH THỨC -- ĐỀ SỐ """ + str(exam_num) + r""" (MÃ ĐỀ """ + str(code) + r""")}}\\[6pt]
\end{center}

\textbf{1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (12 CÂU -- 3,0 ĐIỂM)}\\[2pt]
\renewcommand{\arraystretch}{1.15}
\begin{tabularx}{\linewidth}{|*{12}{>{\centering\arraybackslash}X|}}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} & \textbf{Câu 7} & \textbf{Câu 8} & \textbf{Câu 9} & \textbf{Câu 10} & \textbf{Câu 11} & \textbf{Câu 12}\\
\hline
""")
    row_ans = " & ".join([f"\\textbf{{{q['correct_letter']}}}" for q in p1_qs])
    lines.append(row_ans + r"\\")
    lines.append(r"""\hline
\end{tabularx}

\vspace{8pt}
\textbf{2. PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (2 CÂU -- 2,0 ĐIỂM)}\\[2pt]
\begin{tabularx}{\linewidth}{|c|*{4}{>{\centering\arraybackslash}X|}}
\hline
\textbf{Câu hỏi} & \textbf{Lệnh hỏi a)} & \textbf{Lệnh hỏi b)} & \textbf{Lệnh hỏi c)} & \textbf{Lệnh hỏi d)}\\
\hline
""")
    for r_i, q2 in enumerate(p2_qs):
        c_vals = [f"\\textbf{{{it[2].upper()}}}" for it in q2["items"]]
        lines.append(f"\\textbf{{Câu {13 + r_i}}} & " + " & ".join(c_vals) + r"\\")
        lines.append(r"\hline")
    lines.append(r"""\end{tabularx}

\vspace{8pt}
\textbf{3. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (4 CÂU -- 2,0 ĐIỂM)}\\[2pt]
\begin{tabularx}{\linewidth}{|*{4}{>{\centering\arraybackslash}X|}}
\hline
\textbf{Câu 15} & \textbf{Câu 16} & \textbf{Câu 17} & \textbf{Câu 18}\\
\hline
""")
    row_p3 = " & ".join([f"\\textbf{{{tex_clean(q['answer'])}}}" for q in p3_qs])
    lines.append(row_p3 + r"\\")
    lines.append(r"""\hline
\end{tabularx}

\vspace{8pt}
\textbf{4. PHẦN IV: HƯỚNG DẪN CHẤM TỰ LUẬN (3 CÂU -- 3,0 ĐIỂM)}\\[2pt]
\begin{tabularx}{\linewidth}{|c|X|c|}
\hline
\textbf{Câu} & \textbf{Nội dung yêu cầu / Đáp số chuẩn} & \textbf{Điểm}\\
\hline
""")
    for r_i, q4 in enumerate(p4_qs):
        sum_ans = tex_clean(q4.get("summary_ans", ""))
        lines.append(f"\\textbf{{Câu {19 + r_i}}} & {sum_ans} & \\textbf{{1,0 đ}}\\\\")
        lines.append(r"\hline")
    lines.append(r"""\end{tabularx}

\end{document}
""")

    return "\n".join(lines)

def compile_exam_to_pdf(exam_data):
    code = exam_data["exam_code"]
    exam_num = exam_data["exam_num"]

    tex_name = f"De_Kiem_Tra_GK1_Hoa_11_De{exam_num}_Ma{code}_Tivi.tex"
    tex_path = os.path.join(SCRATCH_DIR, tex_name)

    content = build_latex_exam(exam_data)
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"[+] Đang biên dịch Đề số {exam_num} (Mã đề {code}) bằng pdflatex...")
    for pass_num in range(2):
        cmd = ["pdflatex", "-interaction=nonstopmode", f"-output-directory={SCRATCH_DIR}", tex_path]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="ignore")
        if res.returncode != 0 and pass_num == 1:
            print(f"  [!] Cảnh báo biên dịch ở lần {pass_num+1}: {res.returncode}")
            for line in res.stdout.split("\n")[-30:]:
                if "!" in line or "Error" in line:
                    print("     ", line)

    pdf_temp = os.path.join(SCRATCH_DIR, f"De_Kiem_Tra_GK1_Hoa_11_De{exam_num}_Ma{code}_Tivi.pdf")
    pdf_out = os.path.join(OUT_DIR, f"De_Kiem_Tra_GK1_Hoa_11_De{exam_num}_Ma{code}_TrinhChieu_Tivi_16x9.pdf")

    if os.path.exists(pdf_temp):
        import shutil
        shutil.copyfile(pdf_temp, pdf_out)
        print(f"  -> Thành công: {pdf_out}")

        doc = pymupdf.open(pdf_out)
        page_count = len(doc)
        p0 = doc[0]
        w_cm = round(p0.rect.width * 2.54 / 72, 3)
        h_cm = round(p0.rect.height * 2.54 / 72, 3)
        ratio = round(w_cm / h_cm, 3)
        print(f"     Số trang: {page_count} slide | Kích thước: {w_cm} x {h_cm} cm (Tỉ lệ: {ratio} ~ 16:9)")
        doc.close()
        return pdf_out
    else:
        print(f"  [X] Thất bại khi tạo file PDF cho Đề số {exam_num}!")
        return None

def main():
    print("=" * 70)
    print("BẮT ĐẦU TẠO FILE PDF TRÌNH CHIẾU TIVI 16:9 CHO 3 ĐỀ HÓA HỌC 11")
    print("=" * 70)

    exams = [DE01_DATA, DE02_DATA, DE03_DATA]
    success_files = []

    for exam in exams:
        out_f = compile_exam_to_pdf(exam)
        if out_f:
            success_files.append(out_f)

    print("\n" + "=" * 70)
    print(f"HOÀN THÀNH TẠO {len(success_files)}/3 FILE PDF TRÌNH CHIẾU TIVI 16:9:")
    for f in success_files:
        print(f"  * {os.path.basename(f)}")
    print("=" * 70)

if __name__ == "__main__":
    main()
