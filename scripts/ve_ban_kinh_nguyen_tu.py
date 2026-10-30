# -*- coding: utf-8 -*-
"""
Script vẽ lại sơ đồ Bán kính nguyên tử trong cùng một chu kì (Y), (R), (X), (T)
Chuẩn đồ họa vector cao cấp (3D Shaded Spheres / LaTeX TikZ / SVG)
Biên soạn: Thầy TRẦN VĂN THẠNH - 0777.470.803
Xuất bản:
1. Bản tiêu chuẩn (300 DPI, nền trắng): San_Pham/Ban_Kinh_Nguyen_Tu/so_do_ban_kinh_nguyen_tu.png
2. Bản trong suốt (Transparent): San_Pham/Ban_Kinh_Nguyen_Tu/so_do_ban_kinh_nguyen_tu_transparent.png
3. Bản cực nét (600 DPI in ấn): San_Pham/Ban_Kinh_Nguyen_Tu/so_do_ban_kinh_nguyen_tu_hd.png
4. Bản vector PDF (dùng chèn LaTeX): San_Pham/Ban_Kinh_Nguyen_Tu/so_do_ban_kinh_nguyen_tu.pdf
5. Bản vector SVG: San_Pham/Ban_Kinh_Nguyen_Tu/so_do_ban_kinh_nguyen_tu.svg
6. Mã nguồn TikZ: San_Pham/Ban_Kinh_Nguyen_Tu/so_do_ban_kinh_nguyen_tu.tex
"""

import os
import sys

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import subprocess
import fitz
from PIL import Image

OUT_DIR = r"C:\Antigravity_Thanh\San_Pham\Ban_Kinh_Nguyen_Tu"
os.makedirs(OUT_DIR, exist_ok=True)

# 1. Mã nguồn LaTeX TikZ
TIKZ_TEX = r"""\documentclass[tikz,border=8pt]{standalone}
\usepackage{tikz}
\usepackage{mathptmx} % Times New Roman font
\usetikzlibrary{shadings}

\begin{document}
\begin{tikzpicture}
  % Định nghĩa phong cách quả cầu 3D với hiệu ứng bóng đổ và độ chuyển màu tinh xảo
  % Các nguyên tử Y, R, X, T nằm trên cùng một chu kì với bán kính giảm dần:
  % r_Y (1.25cm) > r_R (0.98cm) > r_X (0.72cm) > r_T (0.50cm)
  % Tất cả các quả cầu đều đặt tiếp xúc trên cùng một đường đáy (baseline y = 0)
  
  % --- Quả cầu (Y) ---
  \shadedraw[shading=ball, ball color=red!92!white, draw=red!75!black, line width=0.55pt] 
    (1.4, 1.25) circle (1.25);
  \node[below, font=\fontsize{14}{17}\selectfont] at (1.4, -0.15) {(Y)};

  % --- Quả cầu (R) ---
  \shadedraw[shading=ball, ball color=red!92!white, draw=red!75!black, line width=0.55pt] 
    (4.6, 0.98) circle (0.98);
  \node[below, font=\fontsize{14}{17}\selectfont] at (4.6, -0.15) {(R)};

  % --- Quả cầu (X) ---
  \shadedraw[shading=ball, ball color=red!92!white, draw=red!75!black, line width=0.55pt] 
    (7.4, 0.72) circle (0.72);
  \node[below, font=\fontsize{14}{17}\selectfont] at (7.4, -0.15) {(X)};

  % --- Quả cầu (T) ---
  \shadedraw[shading=ball, ball color=red!92!white, draw=red!75!black, line width=0.55pt] 
    (9.6, 0.50) circle (0.50);
  \node[below, font=\fontsize{14}{17}\selectfont] at (9.6, -0.15) {(T)};

\end{tikzpicture}
\end{document}
"""

def generate_tikz_files():
    tex_path = os.path.join(OUT_DIR, "so_do_ban_kinh_nguyen_tu.tex")
    pdf_path = os.path.join(OUT_DIR, "so_do_ban_kinh_nguyen_tu.pdf")
    png_path = os.path.join(OUT_DIR, "so_do_ban_kinh_nguyen_tu.png")
    png_hd_path = os.path.join(OUT_DIR, "so_do_ban_kinh_nguyen_tu_hd.png")
    png_trans_path = os.path.join(OUT_DIR, "so_do_ban_kinh_nguyen_tu_transparent.png")

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(TIKZ_TEX)
    print(f"Đã ghi file TeX: {tex_path}")

    # Biên dịch pdflatex
    cmd = ["pdflatex", "-interaction=nonstopmode", "so_do_ban_kinh_nguyen_tu.tex"]
    res = subprocess.run(cmd, cwd=OUT_DIR, capture_output=True, text=True)
    if res.returncode != 0:
        print("Lỗi biên dịch LaTeX:")
        print(res.stdout[-600:])
        return False
    print(f"Đã biên dịch PDF: {pdf_path}")

    # Render PNG tiêu chuẩn (300 DPI)
    doc = fitz.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(dpi=300, alpha=False)
    pix.save(png_path)
    print(f"Đã tạo PNG (300 DPI): {png_path} ({pix.width}x{pix.height})")

    # Render PNG HD (600 DPI)
    pix_hd = page.get_pixmap(dpi=600, alpha=False)
    pix_hd.save(png_hd_path)
    print(f"Đã tạo PNG Ultra HD (600 DPI): {png_hd_path} ({pix_hd.width}x{pix_hd.height})")

    # Render PNG Transparent (300 DPI)
    pix_trans = page.get_pixmap(dpi=300, alpha=True)
    pix_trans.save(png_trans_path)
    print(f"Đã tạo PNG Transparent (300 DPI): {png_trans_path}")

    # Xóa file phụ trợ (.aux, .log)
    for ext in [".aux", ".log"]:
        f_temp = os.path.join(OUT_DIR, "so_do_ban_kinh_nguyen_tu" + ext)
        if os.path.exists(f_temp):
            try:
                os.remove(f_temp)
            except Exception:
                pass
    return True

def generate_svg_file():
    svg_path = os.path.join(OUT_DIR, "so_do_ban_kinh_nguyen_tu.svg")
    svg_content = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 620 180" width="620" height="180">
  <defs>
    <!-- Hiệu ứng chuyển màu 3D Sphere bóng đỏ cao cấp -->
    <radialGradient id="sphereGrad" cx="35%" cy="35%" r="65%" fx="28%" fy="26%">
      <stop offset="0%" stop-color="#ffffff" />
      <stop offset="15%" stop-color="#ffb8be" />
      <stop offset="38%" stop-color="#ff1e32" />
      <stop offset="78%" stop-color="#c80012" />
      <stop offset="100%" stop-color="#7a000a" />
    </radialGradient>
  </defs>

  <!-- Đáy tiếp xúc của các quả cầu tại y = 110 -->
  <!-- Quả cầu (Y): bán kính r = 48, tâm (80, 62) -->
  <circle cx="80" cy="62" r="48" fill="url(#sphereGrad)" stroke="#a0000e" stroke-width="0.8" />
  <text x="80" y="152" text-anchor="middle" font-family="'Times New Roman', Times, serif" font-size="28" fill="#000000">(Y)</text>

  <!-- Quả cầu (R): bán kính r = 37, tâm (260, 73) -->
  <circle cx="260" cy="73" r="37" fill="url(#sphereGrad)" stroke="#a0000e" stroke-width="0.8" />
  <text x="260" y="152" text-anchor="middle" font-family="'Times New Roman', Times, serif" font-size="28" fill="#000000">(R)</text>

  <!-- Quả cầu (X): bán kính r = 27, tâm (425, 83) -->
  <circle cx="425" cy="83" r="27" fill="url(#sphereGrad)" stroke="#a0000e" stroke-width="0.8" />
  <text x="425" y="152" text-anchor="middle" font-family="'Times New Roman', Times, serif" font-size="28" fill="#000000">(X)</text>

  <!-- Quả cầu (T): bán kính r = 19, tâm (550, 91) -->
  <circle cx="550" cy="91" r="19" fill="url(#sphereGrad)" stroke="#a0000e" stroke-width="0.8" />
  <text x="550" y="152" text-anchor="middle" font-family="'Times New Roman', Times, serif" font-size="28" fill="#000000">(T)</text>
</svg>
"""
    with open(svg_path, "w", encoding="utf-8") as f:
        f.write(svg_content)
    print(f"Đã tạo file SVG: {svg_path}")

if __name__ == "__main__":
    generate_tikz_files()
    generate_svg_file()
    print("Hoàn thành toàn bộ quy trình vẽ lại sơ đồ bán kính nguyên tử!")
