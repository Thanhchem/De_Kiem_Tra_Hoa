import os
import subprocess
import pymupdf

OUT_DIR = r"San_Pham\So_Do_Dieu_Che_Aniline"
os.makedirs(OUT_DIR, exist_ok=True)

# Common LaTeX header for standalone chemfig
HEADER = r"""\documentclass[border=3mm]{standalone}
\usepackage[utf8]{inputenc}
\usepackage{amsmath}
\usepackage{chemfig}
\usepackage[scaled=0.95]{helvet}
\renewcommand{\familydefault}{\sfdefault}

\setchemfig{
    atom sep=2.0em,
    bond offset=1.8pt,
    bond style={line width=0.9pt, line cap=round, line join=round},
    arrow style={line width=0.9pt, -{Latex[length=4.5pt, width=3.2pt]}}
}
\renewcommand*\printatom[1]{\ensuremath{\mathsf{#1}}}

\begin{document}
"""

FOOTER = r"""
\end{document}
"""

def compile_and_export(name, body, dpi=300):
    tex_code = HEADER + body + FOOTER
    tex_path = os.path.join(OUT_DIR, f"{name}.tex")
    pdf_path = os.path.join(OUT_DIR, f"{name}.pdf")
    
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_code)
    
    res = subprocess.run(["pdflatex", "-interaction=nonstopmode", f"{name}.tex"], cwd=OUT_DIR, capture_output=True, text=True)
    if res.returncode != 0:
        print(f"Error compiling {name}:", res.stderr or res.stdout)
        return False
    
    doc = pymupdf.open(pdf_path)
    page = doc[0]
    
    # White background (300 DPI)
    pix_white = page.get_pixmap(dpi=dpi, alpha=False)
    pix_white.save(os.path.join(OUT_DIR, f"{name}.png"))
    
    # Transparent background (300 DPI)
    pix_trans = page.get_pixmap(dpi=dpi, alpha=True)
    pix_trans.save(os.path.join(OUT_DIR, f"{name}_transparent.png"))

    # Ultra HD (600 DPI)
    pix_hd = page.get_pixmap(dpi=600, alpha=False)
    pix_hd.save(os.path.join(OUT_DIR, f"{name}_hd.png"))
    
    print(f"Successfully generated: {name}")
    return True

# 1. Full reaction scheme
scheme_body = r"""\schemestart
\chemfig{*6(-=-(-[:30,0.75,,,draw=none]\phantom{NO_2})=-=)}
\arrow{->[][(1)]}[,1.6]
\chemfig{*6(-=-(-[:30,0.75]NO_2)=-=)}
\arrow{->[Zn + HCl][(2)]}[,1.8]
\chemfig{*6(-=-(-[:30,0.75]NH_3Cl)=-=)}
\arrow{->[NaOH][(3)]}[,1.8]
\chemfig{*6(-=-(-[:30,0.75]NH_2)=-=)}
\schemestop"""

compile_and_export("so_do_dieu_che_aniline", scheme_body)

# 2. Individual molecules
molecules = {
    "01_benzen": r"\chemfig{*6(-=-=-=)}",
    "02_nitrobenzen": r"\chemfig{*6(-=-(-[:30,0.75]NO_2)=-=)}",
    "03_phenylamoni_clorua": r"\chemfig{*6(-=-(-[:30,0.75]NH_3Cl)=-=)}",
    "04_aniline": r"\chemfig{*6(-=-(-[:30,0.75]NH_2)=-=)}"
}

for name, body in molecules.items():
    compile_and_export(name, body)

# Clean aux and log
for f in os.listdir(OUT_DIR):
    if f.endswith((".aux", ".log")):
        try:
            os.remove(os.path.join(OUT_DIR, f))
        except:
            pass

print("ALL SCHEMES COMPILED SUCCESSFULLY!")
