import os
import shutil
import glob

ROOT_DIR = r"c:\Antigravity_Thanh"
SCRIPTS_DIR = os.path.join(ROOT_DIR, "scripts")
SAN_PHAM_DIR = os.path.join(ROOT_DIR, "San_Pham")
PDF_GOC_DIR = os.path.join(ROOT_DIR, "PDF_Goc")

os.makedirs(SCRIPTS_DIR, exist_ok=True)
os.makedirs(SAN_PHAM_DIR, exist_ok=True)
os.makedirs(PDF_GOC_DIR, exist_ok=True)

# 1. Copy cropped images into San_Pham/images_crop
sp_img_dir = os.path.join(SAN_PHAM_DIR, "images_crop")
os.makedirs(sp_img_dir, exist_ok=True)
src_crop_dir = os.path.join(ROOT_DIR, "images_crop")
if os.path.exists(src_crop_dir):
    for f in os.listdir(src_crop_dir):
        shutil.copy2(os.path.join(src_crop_dir, f), os.path.join(sp_img_dir, f))

# 2. Move build & tool scripts to scripts/
scripts_to_move = [
    "build_student_exam.py",
    "build_teacher_exam.py",
    "build_word_exams.py",
    "crop_figures.py",
    "generate_all_docx.py",
    "generate_crops.py",
    "make_student_pdf.py",
    "merge_images.py",
    "update_teacher_tex.py"
]
for s in scripts_to_move:
    p = os.path.join(ROOT_DIR, s)
    if os.path.exists(p):
        shutil.move(p, os.path.join(SCRIPTS_DIR, s))

# 3. Clean temporary files in root
root_temps_patterns = [
    "test_*.*",
    "export_images.*",
    "*.aux",
    "*.log",
    "*.thm",
    "*.txt",
    "page_hf_*.png",
    "De_Kiem_Tra_Hoa_11_Ma209.aux",
    "De_Kiem_Tra_Hoa_11_Ma209.log",
    "De_Kiem_Tra_Hoa_11_Ma209.thm",
    "De_Kiem_Tra_Hoa_11_Ma209.pdf",
    "De_Kiem_Tra_Hoa_11_Ma209.tex",
    "De_Kiem_Tra_Hoa_11_Ma209.docx",
    "De_Kiem_Tra_Hoa_11_Ma209_Goc.pdf",
    "test_header.docx"
]
for pat in root_temps_patterns:
    for f in glob.glob(os.path.join(ROOT_DIR, pat)):
        try:
            if os.path.isfile(f):
                os.remove(f)
        except Exception as e:
            print(f"Error removing {f}: {e}")

# 4. Clean temporary files in San_Pham
sp_temps_patterns = [
    "*.aux",
    "*.log",
    "*.thm",
    "page_*.png"
]
for pat in sp_temps_patterns:
    for f in glob.glob(os.path.join(SAN_PHAM_DIR, pat)):
        try:
            if os.path.isfile(f):
                os.remove(f)
        except Exception as e:
            print(f"Error removing {f}: {e}")

print("Workspace organized successfully!")
