import subprocess
import pymupdf
import os

pdflatex = r'C:\Users\Admin\AppData\Local\Programs\MiKTeX\miktex\bin\x64\pdflatex.exe'
os.makedirs('images_crop', exist_ok=True)

figs = {
    'fig_bromobutane': r'''\documentclass[border=2pt]{standalone}
\usepackage{tikz}
\usepackage[version=4]{mhchem}
\begin{document}
\begin{tikzpicture}[scale=0.5, baseline=-2pt]
  \draw[thick] (0,0) -- (0.8,0.6) -- (1.6,0) -- (2.4,0.6) -- (3.1,0.1) node[right=-2pt] {\ce{Br}};
\end{tikzpicture}
\end{document}''',

    'fig_catechin': r'''\documentclass[border=3pt]{standalone}
\usepackage{chemfig}
\begin{document}
\setchemfig{atom sep=1.8em, bond offset=1.5pt}
\chemfig{*6(=(-[6]OH)-(*6(--(<[:-30]OH)-(<:[:30]*6(=-(-[:-30]OH)=(-[:30]OH)-=-))-O-))=-=(-[:150]HO)-)}
\end{document}''',

    'fig_geraniol': r'''\documentclass[border=2pt]{standalone}
\usepackage{tikz}
\usepackage[version=4]{mhchem}
\begin{document}
\begin{tikzpicture}[scale=0.55, font=\small, baseline=0]
  \node[left] at (-0.2,0.6) {\ce{H3C}};
  \draw[thick] (-0.2,0.6) -- (0.4,0.1);
  \draw[thick] (0.4,0.1) -- (0.4,-0.6) node[below] {\ce{CH3}};
  \draw[thick] (0.4,0.1) -- (1.2,0.6);
  \draw[thick] (0.45,0.25) -- (1.15,0.7);
  \draw[thick] (1.2,0.6) -- (1.9,0.1) -- (2.6,0.6) -- (3.3,0.1);
  \draw[thick] (3.3,0.1) -- (3.3,-0.6) node[below] {\ce{CH3}};
  \draw[thick] (3.3,0.1) -- (4.1,0.6);
  \draw[thick] (3.35,0.25) -- (4.05,0.7);
  \draw[thick] (4.1,0.6) -- (4.8,0.1) -- (5.4,0.45) node[right] {\ce{OH}};
\end{tikzpicture}
\end{document}''',

    'fig_shortans': r'''\documentclass[border=2pt]{standalone}
\usepackage{tikz}
\begin{document}
\begin{tikzpicture}[baseline=0]
  \node[left] at (0,0.25) {\textbf{KQ:}};
  \draw[step=0.5cm,gray!80,thick] (0,0) grid (2,0.5);
\end{tikzpicture}
\end{document}'''
}

for name, code in figs.items():
    tex_name = f'temp_{name}.tex'
    pdf_name = f'temp_{name}.pdf'
    with open(tex_name, 'w', encoding='utf-8') as f:
        f.write(code)
    subprocess.run([pdflatex, '-interaction=nonstopmode', tex_name], capture_output=True)
    doc = pymupdf.open(pdf_name)
    pix = doc[0].get_pixmap(dpi=300)
    out_png = os.path.join('images_crop', f'{name}.png')
    pix.save(out_png)
    doc.close()
    print(f'Generated {out_png}')
    for ext in ['.tex', '.pdf', '.aux', '.log']:
        f_to_del = f'temp_{name}{ext}'
        if os.path.exists(f_to_del):
            try:
                os.remove(f_to_del)
            except Exception:
                pass
print('All crops generated successfully!')
