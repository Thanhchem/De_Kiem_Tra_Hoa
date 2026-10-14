import os
import subprocess
import pymupdf

OUT_DIR = r"San_Pham\Cong_Thuc_Hoa_Hoc"
os.makedirs(OUT_DIR, exist_ok=True)

# Common LaTeX header for standalone chemfig
HEADER = r"""\documentclass[border=3mm]{standalone}
\usepackage[utf8]{inputenc}
\usepackage{amsmath}
\usepackage{chemfig}
\usepackage[scaled=0.95]{helvet}
\renewcommand{\familydefault}{\sfdefault}

\setchemfig{
    atom sep=2.2em,
    bond offset=1.8pt,
    bond style={line width=1.0pt, line cap=round, line join=round}
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
    
    # Render with PyMuPDF
    doc = pymupdf.open(pdf_path)
    page = doc[0]
    
    # White background
    pix_white = page.get_pixmap(dpi=dpi, alpha=False)
    pix_white.save(os.path.join(OUT_DIR, f"{name}.png"))
    
    # Transparent background
    pix_trans = page.get_pixmap(dpi=dpi, alpha=True)
    pix_trans.save(os.path.join(OUT_DIR, f"{name}_transparent.png"))
    
    print(f"Successfully generated: {name}.png, {name}_transparent.png, {name}.pdf")
    return True

# 1. Individual formulas
formulas = {
    "ct_A": r"\chemfig{^{+}H_3N-[:30]-[:-30]-[:30]COO^{-}}",
    "ct_B": r"\chemfig{-[:-30](-[:-90]NH_3^{+})-[:30]COO^{-}}",
    "ct_B_anh_goc": r"\chemfig{-[:-30](-[:-90]NH^{+})-[:30]COO^{-}}",
    "ct_C": r"\chemfig{^{+}H_3N-[:30](-[:90]NH_2)-[:-30]-[:30]COO^{-}}",
    "ct_D": r"\chemfig{H_2N-[:30]-[:-30]OH}"
}

for name, body in formulas.items():
    compile_and_export(name, body)

# 2. Combined 1-row layout (A, B, C, D)
row_body = r"""\textbf{A.} $\vcenter{\hbox{\chemfig{^{+}H_3N-[:30]-[:-30]-[:30]COO^{-}}}}$ \qquad\quad
\textbf{B.} $\vcenter{\hbox{\chemfig{-[:-30](-[:-90]NH_3^{+})-[:30]COO^{-}}}}$ \qquad\quad
\textbf{C.} $\vcenter{\hbox{\chemfig{^{+}H_3N-[:30](-[:90]NH_2)-[:-30]-[:30]COO^{-}}}}$ \qquad\quad
\textbf{D.} $\vcenter{\hbox{\chemfig{H_2N-[:30]-[:-30]OH}}}$"""
compile_and_export("de_ABCD_hang_ngang", row_body)

# 3. Combined 2x2 grid layout (as in original image)
grid_body = r"""\begin{tabular}{ll}
\textbf{A.} $\vcenter{\hbox{\chemfig{^{+}H_3N-[:30]-[:-30]-[:30]COO^{-}}}}$ &
\qquad\qquad \textbf{B.} $\vcenter{\hbox{\chemfig{-[:-30](-[:-90]NH_3^{+})-[:30]COO^{-}}}}$ \\[3.5em]
\textbf{C.} $\vcenter{\hbox{\chemfig{^{+}H_3N-[:30](-[:90]NH_2)-[:-30]-[:30]COO^{-}}}}$ &
\qquad\qquad \textbf{D.} $\vcenter{\hbox{\chemfig{H_2N-[:30]-[:-30]OH}}}$
\end{tabular}"""
compile_and_export("de_ABCD_luoi_2x2", grid_body)

# Clean temporary aux/log files
for f in os.listdir(OUT_DIR):
    if f.endswith((".aux", ".log")):
        try:
            os.remove(os.path.join(OUT_DIR, f))
        except:
            pass

print("ALL DONE!")
