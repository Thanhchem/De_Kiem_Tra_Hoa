# -*- coding: utf-8 -*-
"""
Script biên soạn toàn bộ 10 ĐỀ ÔN TẬP vào 1 FILE PDF DUY NHẤT CHUẨN TIVI 16:9
Theo đúng Quy chuẩn FORM CHIẾU TIVI (GEMINI.md Mục 5):
- Tỉ lệ 16:9 widescreen (paperwidth=29.7cm, paperheight=16.706cm), tràn viền 100%, không viền đen.
- Nền trắng tinh (#FFFFFF), chữ đen tuyền (#000000), font Times New Roman, KHÔNG DÙNG KHUNG/BẢNG viền.
- Không có header/footer trang sách trên các slide câu hỏi.
- Phần I (Trắc nghiệm nhiều lựa chọn): Cỡ chữ 30pt (hoặc tự điều chỉnh 26-23pt cho câu dài), 2 câu/slide, dãn cách rộng rãi.
- Phần II (Trắc nghiệm Đúng/Sai): Cỡ chữ 26pt (hoặc 24-22pt cho câu dài), 1 câu/slide (4 ý a, b, c, d có ô [   ]).
- Phần III (Trắc nghiệm Trả lời ngắn): Cỡ chữ 32pt, 1 câu/slide, có dòng chấm trả lời, để trống nửa dưới.
- Slide Bảng đáp án chính thức ở cuối mỗi đề thi.
- Bookmark PDF cho từng đề và từng phần để giáo viên chuyển slide nhanh trên Tivi.
- ĐẢM BẢO CHÍNH XÁC 20 SLIDE CHO MỖI ĐỀ -> TỔNG CỘNG 200 SLIDE KHÔNG TRÀN TRANG!
"""

import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import os
import re
import json
import subprocess
import shutil
import pymupdf

SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
PDF_SOURCE = r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12\10-de-on-tap-C123.pdf"

TEX_FILE = os.path.join(SCRATCH_DIR, "De_On_Tap_Hoa_12_Full_10_De_Tivi_16_9.tex")
PDF_FINAL = os.path.join(OUT_DIR, "De_On_Tap_Hoa_12_Full_10_De_TrinhChieu_Tivi_16_9.pdf")
OD_FINAL = r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12\De_On_Tap_Hoa_12_Full_10_De_TrinhChieu_Tivi_16_9.pdf"

# 1. Đọc và parse dữ liệu từ PDF gốc
doc_src = pymupdf.open(PDF_SOURCE)
all_text = ""
for i in range(25):
    all_text += doc_src[i].get_text() + "\n"

ans_text = doc_src[25].get_text() + "\n" + doc_src[26].get_text()

def parse_answers(text):
    ans_dict = {}
    blocks = re.split(r'ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI ĐỀ SỐ\s+(\d+)', text)
    for k in range(1, len(blocks), 2):
        de_num = int(blocks[k])
        de_ans = blocks[k+1]
        
        # P1
        p1_m = re.search(r'PHẦN I:\s*(.*?)(?=PHẦN II:|$)', de_ans, re.DOTALL)
        p1_ans = {}
        if p1_m:
            items = re.findall(r'(\d+)\.([A-D])', p1_m.group(1))
            for num, opt in items:
                p1_ans[int(num)] = opt
        
        # P2
        p2_m = re.search(r'PHẦN II:\s*(.*?)(?=PHẦN III:|$)', de_ans, re.DOTALL)
        p2_ans = {}
        if p2_m:
            items = re.findall(r'Câu\s+(\d+):\s*\[([ĐS]{4})\]', p2_m.group(1))
            for num, s in items:
                p2_ans[int(num)] = "--".join(list(s))
        
        # P3
        p3_m = re.search(r'PHẦN III:\s*(.*?)(?=ĐÁP ÁN|$)', de_ans, re.DOTALL)
        p3_ans = {}
        if p3_m:
            items = re.findall(r'Câu\s+(\d+):\s*([^|\n]+)', p3_m.group(1))
            for num, val in items:
                p3_ans[int(num)] = val.strip()
        
        ans_dict[de_num] = {
            'P1': p1_ans,
            'P2': p2_ans,
            'P3': p3_ans
        }
    return ans_dict

answers = parse_answers(ans_text)

def parse_options_sequential(q_text):
    mA = re.search(r'(?:^|\s)A\.\s*', q_text)
    if not mA: return None
    posA_start, posA_end = mA.start(), mA.end()
    
    mB = re.search(r'(?:^|\s)B\.\s*', q_text[posA_end:])
    if not mB: return None
    posB_start = posA_end + mB.start()
    posB_end = posA_end + mB.end()
    
    mC = re.search(r'(?:^|\s)C\.\s*', q_text[posB_end:])
    if not mC: return None
    posC_start = posB_end + mC.start()
    posC_end = posB_end + mC.end()
    
    mD = re.search(r'(?:^|\s)D\.\s*', q_text[posC_end:])
    if not mD: return None
    posD_start = posC_end + mD.start()
    posD_end = posC_end + mD.end()
    
    stem = ' '.join(q_text[:posA_start].strip().split())
    optA = ' '.join(q_text[posA_end:posB_start].strip().split())
    optB = ' '.join(q_text[posB_end:posC_start].strip().split())
    optC = ' '.join(q_text[posC_end:posD_start].strip().split())
    optD = ' '.join(q_text[posD_end:].strip().split())
    
    return stem, {'A': optA, 'B': optB, 'C': optC, 'D': optD}

def parse_part2_sequential(q_text):
    ma = re.search(r'(?:^|\s)a[\)\.]\s*', q_text)
    if not ma: return None
    pos_a_start, pos_a_end = ma.start(), ma.end()
    
    mb = re.search(r'(?:^|\s)b[\)\.]\s*', q_text[pos_a_end:])
    if not mb: return None
    pos_b_start = pos_a_end + mb.start()
    pos_b_end = pos_a_end + mb.end()
    
    mc = re.search(r'(?:^|\s)c[\)\.]\s*', q_text[pos_b_end:])
    if not mc: return None
    pos_c_start = pos_b_end + mc.start()
    pos_c_end = pos_b_end + mc.end()
    
    md = re.search(r'(?:^|\s)d[\)\.]\s*', q_text[pos_c_end:])
    if not md: return None
    pos_d_start = pos_c_end + md.start()
    pos_d_end = pos_c_end + md.end()
    
    stem = ' '.join(q_text[:pos_a_start].strip().split())
    opta = ' '.join(q_text[pos_a_end:pos_b_start].strip().split())
    optb = ' '.join(q_text[pos_b_end:pos_c_start].strip().split())
    optc = ' '.join(q_text[pos_c_end:pos_d_start].strip().split())
    optd = ' '.join(q_text[pos_d_end:].strip().split())
    
    return stem, {'a': opta, 'b': optb, 'c': optc, 'd': optd}

tests_raw = re.split(r'ĐỀ\s+ÔN\s+TẬP\s+CHƯƠNG\s+1,\s*2,\s*3', all_text)[1:]
all_tests = []

for idx, t_raw in enumerate(tests_raw, 1):
    t = re.sub(r'^\s*\d+\s*$', '', t_raw, flags=re.MULTILINE)
    t = re.sub(r'[-–—]{3,}.*?[-–—]{3,}', '', t)
    
    p1_m = re.search(r'PHẦN I\..*?(?=PHẦN II\.|$)', t, re.DOTALL)
    p2_m = re.search(r'PHẦN II\..*?(?=PHẦN III\.|$)', t, re.DOTALL)
    p3_m = re.search(r'PHẦN III\..*?$', t, re.DOTALL)
    
    p1_str = p1_m.group(0) if p1_m else ""
    p2_str = p2_m.group(0) if p2_m else ""
    p3_str = p3_m.group(0) if p3_m else ""
    
    # P1
    q1_splits = list(re.finditer(r'Câu\s+(\d+)\.\s*', p1_str))
    p1_list = []
    for i in range(len(q1_splits)):
        num = int(q1_splits[i].group(1))
        start = q1_splits[i].end()
        end = q1_splits[i+1].start() if i+1 < len(q1_splits) else len(p1_str)
        stem, opts = parse_options_sequential(p1_str[start:end])
        p1_list.append({'num': num, 'stem': stem, 'opts': opts})
        
    # P2
    q2_splits = list(re.finditer(r'Câu\s+(\d+)\.\s*', p2_str))
    p2_list = []
    for i in range(len(q2_splits)):
        num = int(q2_splits[i].group(1))
        start = q2_splits[i].end()
        end = q2_splits[i+1].start() if i+1 < len(q2_splits) else len(p2_str)
        stem, items = parse_part2_sequential(p2_str[start:end])
        p2_list.append({'num': num, 'stem': stem, 'items': items})
        
    # P3
    q3_splits = list(re.finditer(r'Câu\s+(\d+)\.\s*', p3_str))
    p3_list = []
    for i in range(len(q3_splits)):
        num = int(q3_splits[i].group(1))
        start = q3_splits[i].end()
        end = q3_splits[i+1].start() if i+1 < len(q3_splits) else len(p3_str)
        stem = ' '.join(p3_str[start:end].strip().split())
        p3_list.append({'num': num, 'stem': stem})
        
    all_tests.append({
        'test_num': idx,
        'p1': p1_list,
        'p2': p2_list,
        'p3': p3_list,
        'ans': answers.get(idx, {})
    })

print(f"Đã nạp toàn bộ {len(all_tests)} đề ôn tập.")

def format_chem(text):
    if not text:
        return ""
    
    text = text.replace("&", r"\&")
    text = re.sub(r'(?<!\\)%', r'\%', text)
    text = text.replace("α", r"$\alpha$").replace("β", r"$\beta$").replace("γ", r"$\gamma$").replace("ω", r"$\omega$")
    text = text.replace("→", r"$\to$")
    text = text.replace("t°", r"$t^\circ$").replace("t °", r"$t^\circ$")
    text = re.sub(r'(\d+)\s*°(?![a-zA-Z])', r'$\1^{\\circ}$', text)
    
    exclude_words = {'DNA', 'RNA', 'ATP', 'ADP', 'pH', 'Ala', 'Gly', 'Val', 'Glu', 'Lys', 'Ser', 'His', 'Leu', 'M', 'X', 'Y', 'Z', 'A', 'B', 'C', 'D'}
    known_inorganics = {'NaOH', 'HCl', 'KOH', 'NaCl', 'AgNO3', 'NH3', 'Cu(OH)2', 'H2SO4', 'HNO3', 'H2O', 'CO2', 'SO2', 'Cu2O', 'Fe2O3', 'KMnO4', 'Ba(OH)2', 'BaSO4', 'CaCO3'}
    
    def repl(m):
        w = m.group(0)
        if w in exclude_words or w.isdigit():
            return w
        if (re.search(r'\d', w) and re.search(r'[A-Z]', w) and not re.search(r'[àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]', w, re.I)) or (w in known_inorganics):
            return f"\\ce{{{w}}}"
        return w

    res = re.sub(r'\b[A-Z][A-Za-z0-9\(\)\[\]\-\=\+]*\b', repl, text)
    return res

# 2. XÂY DỰNG FILE TEX HOÀN CHỈNH CHO 10 ĐỀ
tex_out = []

tex_out.append(r"""\documentclass{article}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{tabularx}
\usepackage{array}
\usepackage{booktabs}
\usepackage{xcolor}
\usepackage[version=4]{mhchem}
\usepackage{bookmark}
\usepackage[
  paperwidth=29.7cm,
  paperheight=16.706cm,
  top=0.9cm,
  bottom=0.9cm,
  left=1.2cm,
  right=1.2cm,
  footskip=0pt,
  headheight=0pt,
  headsep=0pt
]{geometry}

\pagestyle{empty}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}

\definecolor{hdrBlue}{RGB}{13, 71, 161}
\definecolor{subGreen}{RGB}{0, 105, 92}
\definecolor{ansRed}{RGB}{198, 40, 40}

\begin{document}
""")

for t_data in all_tests:
    de_num = t_data['test_num']
    ans = t_data['ans']
    
    tex_out.append(f"% ==============================================================================")
    tex_out.append(f"% ĐỀ SỐ {de_num}")
    tex_out.append(f"% ==============================================================================")
    tex_out.append(f"\\pdfbookmark[0]{{ĐỀ SỐ {de_num}}}{{de_{de_num}}}")
    
    # --------------------------------------------------------------------------
    # PHẦN I: 18 câu (2 câu / slide = 9 slides) - Cỡ chữ 30pt (linh hoạt 26-23pt nếu câu dài)
    # --------------------------------------------------------------------------
    tex_out.append(f"\\pdfbookmark[1]{{Phần I (Câu 1 - 18)}}{{de_{de_num}_p1}}")
    
    p1 = t_data['p1']
    for s_idx in range(9):
        qA = p1[s_idx * 2]
        qB = p1[s_idx * 2 + 1]
        
        # Tính toán độ dài để chọn cỡ chữ tối ưu không tràn trang
        t1_raw = qA['stem'] + ' ' + ' '.join(qA['opts'].values())
        t2_raw = qB['stem'] + ' ' + ' '.join(qB['opts'].values())
        total_chars = len(t1_raw) + len(t2_raw)
        
        if total_chars <= 300:
            font_cmd = r"\fontsize{30}{39}\selectfont"
            row_gap = "8pt"
        elif total_chars <= 400:
            font_cmd = r"\fontsize{26}{34}\selectfont"
            row_gap = "6pt"
        else:
            font_cmd = r"\fontsize{23}{30}\selectfont"
            row_gap = "5pt"
        
        # Câu thứ nhất
        stemA = format_chem(qA['stem'])
        optA_A = format_chem(qA['opts']['A'])
        optA_B = format_chem(qA['opts']['B'])
        optA_C = format_chem(qA['opts']['C'])
        optA_D = format_chem(qA['opts']['D'])
        
        # Câu thứ hai
        stemB = format_chem(qB['stem'])
        optB_A = format_chem(qB['opts']['A'])
        optB_B = format_chem(qB['opts']['B'])
        optB_C = format_chem(qB['opts']['C'])
        optB_D = format_chem(qB['opts']['D'])
        
        slide_tex = f"""{font_cmd}
\\textbf{{Câu {qA['num']}.}} {stemA}
\\vspace{{5pt}}

\\begin{{tabularx}}{{\\linewidth}}{{@{{}}X X@{{}}}}
\\textbf{{A.}} {optA_A} & \\textbf{{B.}} {optA_B}\\\\[{row_gap}]
\\textbf{{C.}} {optA_C} & \\textbf{{D.}} {optA_D}
\\end{{tabularx}}

\\vspace*{{\\fill}}

\\textbf{{Câu {qB['num']}.}} {stemB}
\\vspace{{5pt}}

\\begin{{tabularx}}{{\\linewidth}}{{@{{}}X X@{{}}}}
\\textbf{{A.}} {optB_A} & \\textbf{{B.}} {optB_B}\\\\[{row_gap}]
\\textbf{{C.}} {optB_C} & \\textbf{{D.}} {optB_D}
\\end{{tabularx}}

\\newpage
"""
        tex_out.append(slide_tex)
        
    # --------------------------------------------------------------------------
    # PHẦN II: 4 câu (1 câu / slide = 4 slides) - Cỡ chữ 26pt (linh hoạt nếu dài)
    # --------------------------------------------------------------------------
    tex_out.append(f"\\pdfbookmark[1]{{Phần II (Câu 1 - 4 Đúng/Sai)}}{{de_{de_num}_p2}}")
    
    p2 = t_data['p2']
    for q in p2:
        total_p2_chars = len(q['stem']) + sum(len(v) for v in q['items'].values())
        if total_p2_chars <= 350:
            font_cmd = r"\fontsize{26}{34}\selectfont"
            row_gap = "12pt"
            top_gap = "14pt"
        elif total_p2_chars <= 440:
            font_cmd = r"\fontsize{24}{31}\selectfont"
            row_gap = "10pt"
            top_gap = "12pt"
        else:
            font_cmd = r"\fontsize{22}{28}\selectfont"
            row_gap = "8pt"
            top_gap = "10pt"
            
        stem = format_chem(q['stem'])
        item_a = format_chem(q['items']['a'])
        item_b = format_chem(q['items']['b'])
        item_c = format_chem(q['items']['c'])
        item_d = format_chem(q['items']['d'])
        
        slide_tex = f"""{font_cmd}
\\textbf{{Câu {q['num']} (Đúng/Sai).}} {stem}
\\vspace{{{top_gap}}}

\\begin{{tabularx}}{{\\linewidth}}{{@{{}}p{{0.88\\linewidth}} c@{{}}}}
\\textbf{{a)}} {item_a} & $[\\quad]$\\\\[{row_gap}]
\\textbf{{b)}} {item_b} & $[\\quad]$\\\\[{row_gap}]
\\textbf{{c)}} {item_c} & $[\\quad]$\\\\[{row_gap}]
\\textbf{{d)}} {item_d} & $[\\quad]$
\\end{{tabularx}}

\\newpage
"""
        tex_out.append(slide_tex)

    # --------------------------------------------------------------------------
    # PHẦN III: 6 câu (1 câu / slide = 6 slides) - Cỡ chữ 32pt
    # --------------------------------------------------------------------------
    tex_out.append(f"\\pdfbookmark[1]{{Phần III (Câu 1 - 6 Trả lời ngắn)}}{{de_{de_num}_p3}}")
    tex_out.append(r"\fontsize{32}{42}\selectfont")
    
    p3 = t_data['p3']
    for q in p3:
        stem = format_chem(q['stem'])
        
        slide_tex = f"""\\textbf{{Câu {q['num']} (Trả lời ngắn).}} {stem}
\\vspace{{22pt}}

\\textbf{{Trả lời:}} $\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots$

\\newpage
"""
        tex_out.append(slide_tex)

    # --------------------------------------------------------------------------
    # SLIDE BẢNG ĐÁP ÁN ĐỀ SỐ X (1 slide)
    # --------------------------------------------------------------------------
    tex_out.append(f"\\pdfbookmark[1]{{Bảng đáp án Đề số {de_num}}}{{de_{de_num}_ans}}")
    
    p1_ans = ans.get('P1', {})
    p2_ans = ans.get('P2', {})
    p3_ans = ans.get('P3', {})
    
    row1_nums = " & ".join([f"\\textbf{{{i}}}" for i in range(1, 10)])
    row1_vals = " & ".join([f"\\color{{ansRed}}\\textbf{{{p1_ans.get(i, '')}}}" for i in range(1, 10)])
    row2_nums = " & ".join([f"\\textbf{{{i}}}" for i in range(10, 19)])
    row2_vals = " & ".join([f"\\color{{ansRed}}\\textbf{{{p1_ans.get(i, '')}}}" for i in range(10, 19)])
    
    p2_cols = " &\n".join([f"\\textbf{{Câu {i}:}} [\\color{{ansRed}}\\textbf{{{p2_ans.get(i, '')}}}\\color{{black}}]" for i in range(1, 5)])
    
    p3_row1 = " & ".join([f"\\textbf{{Câu {i}:}} {{\\color{{ansRed}}\\textbf{{{p3_ans.get(i, '')}}}}}" for i in [1, 2, 3]])
    p3_row2 = " & ".join([f"\\textbf{{Câu {i}:}} {{\\color{{ansRed}}\\textbf{{{p3_ans.get(i, '')}}}}}" for i in [4, 5, 6]])
    
    ans_slide_tex = f"""\\fontsize{{17}}{{21}}\\selectfont
\\begin{{center}}
{{\\fontsize{{23}}{{27}}\\selectfont\\textbf{{BẢNG ĐÁP ÁN CHÍNH THỨC -- ĐỀ SỐ {de_num}}}}}\\\\[8pt]
\\end{{center}}

\\textbf{{1. PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (4,5 ĐIỂM)}}\\\\[4pt]
\\renewcommand{{\\arraystretch}}{{1.2}}
\\begin{{tabularx}}{{\\linewidth}}{{|*{{9}}{{>{{\\centering\\arraybackslash}}X|}}}}
\\hline
{row1_nums} \\\\
\\hline
{row1_vals} \\\\
\\hline
{row2_nums} \\\\
\\hline
{row2_vals} \\\\
\\hline
\\end{{tabularx}}

\\vspace{{10pt}}

\\textbf{{2. PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG SAI (4,0 ĐIỂM)}}\\\\[4pt]
\\begin{{tabularx}}{{\\linewidth}}{{@{{}}*{{4}}{{>{{\\raggedright\\arraybackslash}}X}}@{{}}}}
{p2_cols}
\\end{{tabularx}}

\\vspace{{10pt}}

\\textbf{{3. PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (1,5 ĐIỂM)}}\\\\[4pt]
\\begin{{tabularx}}{{\\linewidth}}{{@{{}}XXX@{{}}}}
{p3_row1} \\\\[6pt]
{p3_row2}
\\end{{tabularx}}

\\newpage
"""
    tex_out.append(ans_slide_tex)

tex_out.append(r"\end{document}")

# Ghi file TeX
full_tex = "\n".join(tex_out)
with open(TEX_FILE, "w", encoding="utf-8") as f:
    f.write(full_tex)

print(f"Đã tạo file TeX hoàn chỉnh cho 10 đề: {TEX_FILE}")

# Biên dịch pdflatex 2 lần để bookmark chuẩn xác
print("Đang tiến hành biên dịch pdflatex lần 1...")
res = subprocess.run(["pdflatex", "-interaction=nonstopmode", "-output-directory", SCRATCH_DIR, TEX_FILE], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

print("Đang tiến hành biên dịch pdflatex lần 2 (cập nhật bookmark outline)...")
res2 = subprocess.run(["pdflatex", "-interaction=nonstopmode", "-output-directory", SCRATCH_DIR, TEX_FILE], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

pdf_temp = os.path.join(SCRATCH_DIR, "De_On_Tap_Hoa_12_Full_10_De_Tivi_16_9.pdf")
if os.path.exists(pdf_temp):
    shutil.copy2(pdf_temp, PDF_FINAL)
    print(f"BIÊN DỊCH THÀNH CÔNG RỰC RỠ: {PDF_FINAL}")
    
    doc_res = pymupdf.open(PDF_FINAL)
    print(f"TỔNG SỐ TRANG CỦA FILE PDF 10 ĐỀ: {len(doc_res)} trang")
    
    # Sao lưu sang OneDrive
    try:
        shutil.copy2(PDF_FINAL, OD_FINAL)
        print(f"ĐÃ SAO LƯU SANG ONEDRIVE: {OD_FINAL}")
    except Exception as e:
        print(f"Lỗi sao lưu OneDrive: {e}")
else:
    print("LỖI BIÊN DỊCH PDFLATEX:")
    print(res.stdout[-2000:])
