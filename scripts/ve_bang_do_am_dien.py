# -*- coding: utf-8 -*-
"""
Script vẽ lại Bảng giá trị độ âm điện của các nguyên tố (Pauling)
Na: 0,93 | H: 2,20 | Cl: 3,16 | C: 2,55 | O: 3,44 | S: 2,58 | N: 3,04 | Ca: 1,00 | Br: 2,96 | K: 0,82
Chuẩn đồ họa vector cao cấp (LaTeX Tabular / PyMuPDF / SVG)
Biên soạn: Thầy TRẦN VĂN THẠNH - 0777.470.803

Xuất bản:
1. Bản tiêu chuẩn (300 DPI, nền trắng): San_Pham/Bang_Do_Am_Dien/bang_do_am_dien.png
2. Bản trong suốt (Transparent): San_Pham/Bang_Do_Am_Dien/bang_do_am_dien_transparent.png
3. Bản cực nét (600 DPI in ấn): San_Pham/Bang_Do_Am_Dien/bang_do_am_dien_hd.png
4. Bản vector PDF (dùng chèn LaTeX/in ấn): San_Pham/Bang_Do_Am_Dien/bang_do_am_dien.pdf
5. Bản vector SVG: San_Pham/Bang_Do_Am_Dien/bang_do_am_dien.svg
6. Mã nguồn LaTeX: San_Pham/Bang_Do_Am_Dien/bang_do_am_dien.tex
7. Bản kèm dòng chữ dẫn đề: bang_do_am_dien_kem_chu (.png, .pdf, .svg, _transparent.png)
"""

import os
import sys
import subprocess
import pymupdf

OUT_DIR = os.path.abspath(r"San_Pham\Bang_Do_Am_Dien")
os.makedirs(OUT_DIR, exist_ok=True)

# 1. Mã nguồn LaTeX Bảng chuẩn (thuần bảng)
TEX_TABLE_ONLY = r"""\documentclass[border=4pt]{standalone}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{mathptmx}
\usepackage{array}

\newcolumntype{C}{>{\centering\arraybackslash}p{1.2cm}}

\begin{document}
\renewcommand{\arraystretch}{1.55}
\setlength{\tabcolsep}{0pt}
\begin{tabular}{|*{10}{C|}}
\hline
\textbf{\large Na} & \textbf{\large H} & \textbf{\large Cl} & \textbf{\large C} & \textbf{\large O} & \textbf{\large S} & \textbf{\large N} & \textbf{\large Ca} & \textbf{\large Br} & \textbf{\large K} \\
\hline
\large 0,93 & \large 2,20 & \large 3,16 & \large 2,55 & \large 3,44 & \large 2,58 & \large 3,04 & \large 1,00 & \large 2,96 & \large 0,82 \\
\hline
\end{tabular}
\end{document}
"""

# 2. Mã nguồn LaTeX Bảng kèm chữ dẫn đề thi
TEX_WITH_CAPTION = r"""\documentclass[border=5pt]{standalone}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{mathptmx}
\usepackage{array}

\newcolumntype{C}{>{\centering\arraybackslash}p{1.2cm}}

\begin{document}
\renewcommand{\arraystretch}{1.55}
\setlength{\tabcolsep}{0pt}
\begin{tabular}{l}
\textit{\normalsize Cho giá trị độ âm điện của một số nguyên tố:}\\[4pt]
\begin{tabular}{|*{10}{C|}}
\hline
\textbf{\large Na} & \textbf{\large H} & \textbf{\large Cl} & \textbf{\large C} & \textbf{\large O} & \textbf{\large S} & \textbf{\large N} & \textbf{\large Ca} & \textbf{\large Br} & \textbf{\large K} \\
\hline
\large 0,93 & \large 2,20 & \large 3,16 & \large 2,55 & \large 3,44 & \large 2,58 & \large 3,04 & \large 1,00 & \large 2,96 & \large 0,82 \\
\hline
\end{tabular}
\end{tabular}
\end{document}
"""

def compile_and_export(base_name, tex_content):
    tex_path = os.path.join(OUT_DIR, f"{base_name}.tex")
    pdf_path = os.path.join(OUT_DIR, f"{base_name}.pdf")
    png_path = os.path.join(OUT_DIR, f"{base_name}.png")
    png_hd_path = os.path.join(OUT_DIR, f"{base_name}_hd.png")
    png_trans_path = os.path.join(OUT_DIR, f"{base_name}_transparent.png")

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_content)
    print(f"Đã ghi file TeX: {tex_path}")

    # Biên dịch pdflatex
    cmd = ["pdflatex", "-interaction=nonstopmode", f"{base_name}.tex"]
    res = subprocess.run(cmd, cwd=OUT_DIR, capture_output=True, text=True)
    if res.returncode != 0:
        print(f"Lỗi biên dịch LaTeX cho {base_name}:")
        print(res.stdout[-600:])
        return False
    print(f"Đã biên dịch PDF: {pdf_path}")

    # Xuất PNG qua PyMuPDF
    doc = pymupdf.open(pdf_path)
    page = doc[0]

    # Bản 300 DPI nền trắng
    pix = page.get_pixmap(dpi=300, alpha=False)
    pix.save(png_path)
    print(f"Đã tạo PNG (300 DPI): {png_path} ({pix.width}x{pix.height})")

    # Bản 600 DPI Ultra HD
    pix_hd = page.get_pixmap(dpi=600, alpha=False)
    pix_hd.save(png_hd_path)
    print(f"Đã tạo PNG Ultra HD (600 DPI): {png_hd_path} ({pix_hd.width}x{pix_hd.height})")

    # Bản 300 DPI Transparent (nền trong suốt)
    pix_trans = page.get_pixmap(dpi=300, alpha=True)
    pix_trans.save(png_trans_path)
    print(f"Đã tạo PNG Transparent (300 DPI): {png_trans_path}")

    # Xóa file rác
    for ext in [".aux", ".log"]:
        f_temp = os.path.join(OUT_DIR, f"{base_name}{ext}")
        if os.path.exists(f_temp):
            try:
                os.remove(f_temp)
            except Exception:
                pass
    return True

def generate_svg():
    # 1. SVG thuần bảng
    svg_table_path = os.path.join(OUT_DIR, "bang_do_am_dien.svg")
    
    # 10 cột, mỗi cột rộng 58px, 2 hàng mỗi hàng cao 36px
    col_w = 58
    row_h = 36
    total_w = col_w * 10
    total_h = row_h * 2
    
    elements = ["Na", "H", "Cl", "C", "O", "S", "N", "Ca", "Br", "K"]
    values   = ["0,93", "2,20", "3,16", "2,55", "3,44", "2,58", "3,04", "1,00", "2,96", "0,82"]

    svg_lines = []
    svg_lines.append(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {total_w + 4} {total_h + 4}" width="{total_w + 4}" height="{total_h + 4}">')
    svg_lines.append(f'  <rect x="2" y="2" width="{total_w}" height="{total_h}" fill="#ffffff" stroke="#000000" stroke-width="1.2" />')
    
    # Kẻ đường ngang giữa
    svg_lines.append(f'  <line x1="2" y1="{2 + row_h}" x2="{2 + total_w}" y2="{2 + row_h}" stroke="#000000" stroke-width="1" />')
    
    # Kẻ các đường dọc và điền chữ
    for i in range(10):
        x_left = 2 + i * col_w
        x_mid = x_left + col_w / 2
        # Đường dọc phân cách cột
        if i > 0:
            svg_lines.append(f'  <line x1="{x_left}" y1="2" x2="{x_left}" y2="{2 + total_h}" stroke="#000000" stroke-width="1" />')
        # Chữ ký hiệu nguyên tố (hàng 1 - in đậm)
        svg_lines.append(f'  <text x="{x_mid}" y="{2 + 25}" text-anchor="middle" font-family="\'Times New Roman\', Times, serif" font-weight="bold" font-size="20" fill="#000000">{elements[i]}</text>')
        # Giá trị độ âm điện (hàng 2)
        svg_lines.append(f'  <text x="{x_mid}" y="{2 + row_h + 24}" text-anchor="middle" font-family="\'Times New Roman\', Times, serif" font-size="18" fill="#000000">{values[i]}</text>')
    
    svg_lines.append('</svg>')
    
    with open(svg_table_path, "w", encoding="utf-8") as f:
        f.write("\n".join(svg_lines))
    print(f"Đã tạo SVG: {svg_table_path}")

    # 2. SVG kèm dòng chữ dẫn
    svg_caption_path = os.path.join(OUT_DIR, "bang_do_am_dien_kem_chu.svg")
    pad_top = 30
    svg_cap_lines = []
    svg_cap_lines.append(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {total_w + 4} {total_h + pad_top + 4}" width="{total_w + 4}" height="{total_h + pad_top + 4}">')
    svg_cap_lines.append(f'  <text x="2" y="20" font-family="\'Times New Roman\', Times, serif" font-style="italic" font-size="17" fill="#000000">Cho giá trị độ âm điện của một số nguyên tố:</text>')
    svg_cap_lines.append(f'  <rect x="2" y="{pad_top + 2}" width="{total_w}" height="{total_h}" fill="#ffffff" stroke="#000000" stroke-width="1.2" />')
    svg_cap_lines.append(f'  <line x1="2" y1="{pad_top + 2 + row_h}" x2="{2 + total_w}" y2="{pad_top + 2 + row_h}" stroke="#000000" stroke-width="1" />')
    
    for i in range(10):
        x_left = 2 + i * col_w
        x_mid = x_left + col_w / 2
        if i > 0:
            svg_cap_lines.append(f'  <line x1="{x_left}" y1="{pad_top + 2}" x2="{x_left}" y2="{pad_top + 2 + total_h}" stroke="#000000" stroke-width="1" />')
        svg_cap_lines.append(f'  <text x="{x_mid}" y="{pad_top + 2 + 25}" text-anchor="middle" font-family="\'Times New Roman\', Times, serif" font-weight="bold" font-size="20" fill="#000000">{elements[i]}</text>')
        svg_cap_lines.append(f'  <text x="{x_mid}" y="{pad_top + 2 + row_h + 24}" text-anchor="middle" font-family="\'Times New Roman\', Times, serif" font-size="18" fill="#000000">{values[i]}</text>')
    
    svg_cap_lines.append('</svg>')

    with open(svg_caption_path, "w", encoding="utf-8") as f:
        f.write("\n".join(svg_cap_lines))
    print(f"Đã tạo SVG kèm chữ: {svg_caption_path}")

if __name__ == "__main__":
    compile_and_export("bang_do_am_dien", TEX_TABLE_ONLY)
    compile_and_export("bang_do_am_dien_kem_chu", TEX_WITH_CAPTION)
    generate_svg()
    print("\n[HOÀN TẤT] Đã tạo thành công toàn bộ các định dạng vector và hình ảnh bảng độ âm điện!")
