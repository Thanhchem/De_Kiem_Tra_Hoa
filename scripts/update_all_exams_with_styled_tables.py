# -*- coding: utf-8 -*-
"""
Script cập nhật tự động 4 file tạo đề thi và tạo lại toàn bộ file Word thành phẩm.
"""

import re
import subprocess
import sys
import os

configs = [
    ("scripts/generate_de_nguyen_dinh_chieu_1002.py", 0),
    ("scripts/generate_de_hoa_chau_132.py", 1),
    ("scripts/generate_de_dang_huy_tru_104.py", 2),
    ("scripts/generate_de_nguyen_chi_thanh_1104.py", 3),
]

for script_path, data_idx in configs:
    with open(script_path, "r", encoding="utf-8") as f:
        content = f.read()

    # Tìm đoạn if is_teacher: ... đến trước add_section_header
    pattern = r'(# -{30,}\s+# BẢNG ĐÁP ÁN RÚT GỌN.*?\n\s+if is_teacher:)(.*?)(?=\n\s+# -{30,}\s+# PHẦN I)'
    replacement = rf"""# -------------------------------------------------------------
    # BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO BẢN GIÁO VIÊN THEO MẪU ẢNH CHUẨN)
    # -------------------------------------------------------------
    if is_teacher:
        from build_bang_dap_an_rut_gon import add_styled_answer_block, EXAMS_DATA
        add_styled_answer_block(doc, EXAMS_DATA[{data_idx}])"""

    new_content = re.sub(pattern, replacement, content, flags=re.DOTALL)
    if new_content != content:
        with open(script_path, "w", encoding="utf-8") as f:
            f.write(new_content)
        print(f"Đã cập nhật {script_path}")
    else:
        print(f"Không tìm thấy đoạn pattern trong {script_path}")

print("Chạy lại tất cả các script sinh file Word...")
for script_path, _ in configs:
    ret = subprocess.run([sys.executable, script_path], capture_output=True, text=True, encoding="utf-8")
    print(ret.stdout.strip())
    if ret.stderr:
        print("Lỗi:", ret.stderr.strip())

# Tạo file master
ret_master = subprocess.run([sys.executable, "scripts/build_bang_dap_an_rut_gon.py"], capture_output=True, text=True, encoding="utf-8")
print(ret_master.stdout.strip())
