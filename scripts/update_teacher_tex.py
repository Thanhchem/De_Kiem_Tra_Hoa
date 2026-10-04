import subprocess
import pymupdf

with open(r'c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_Hoa_11_Ma209_GiaoVien.tex', 'r', encoding='utf-8') as f:
    c = f.read()

# Custom loigiai definition
custom_defs = r'''
\renewcommand{\loigiai}[1]{%
  \par\smallskip\noindent{\color{blue!80!black}\textbf{Lời giải:}}\ #1\par\smallskip
}
\renewcommand{\True}{\bfseries\color{red!80!black}}
'''

if r'\renewcommand{\loigiai}' not in c:
    c = c.replace(r'\begin{document}', custom_defs + "\n\\begin{document}")

with open(r'c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_Hoa_11_Ma209_GiaoVien.tex', 'w', encoding='utf-8') as f:
    f.write(c)

pdflatex = r'C:\Users\Admin\AppData\Local\Programs\MiKTeX\miktex\bin\x64\pdflatex.exe'
res = subprocess.run([pdflatex, '-interaction=nonstopmode', 'De_Kiem_Tra_Hoa_11_Ma209_GiaoVien.tex'], cwd=r'c:\Antigravity_Thanh\San_Pham', capture_output=True, text=True)
print('Compile status:', res.returncode)
doc = pymupdf.open(r'c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_Hoa_11_Ma209_GiaoVien.pdf')
print('Teacher pages:', len(doc))
for i in range(len(doc)):
    doc[i].get_pixmap(dpi=150).save(f'c:/Antigravity_Thanh/San_Pham/page_gv_{i+1}.png')
print('Saved pages')
doc.close()
