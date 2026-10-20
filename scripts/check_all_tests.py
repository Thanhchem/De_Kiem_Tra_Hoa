# -*- coding: utf-8 -*-
"""
Trích xuất từng đề và kiểm tra số lượng câu hỏi trong mỗi phần
"""
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import pymupdf
import re

PDF_PATH = r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12\10-de-on-tap-C123.pdf"
doc = pymupdf.open(PDF_PATH)

# Ghép toàn bộ nội dung từ trang 1 đến trang 25
all_pages_text = []
for i in range(25):
    all_pages_text.append(doc[i].get_text())

full_exam_text = "\n===PAGE_BREAK===\n".join(all_pages_text)

# Tìm vị trí bắt đầu của các đề
# Các tiêu đề: ĐỀ SỐ 1, ĐỀ SỐ 2, ..., ĐỀ SỐ 10 (chú ý đề 10 ở trang 24 bị ghi nhầm thành ĐỀ SỐ 8)
# Hãy tách theo pattern 'ĐỀ\s+ÔN\s+TẬP\s+CHƯƠNG\s+1,\s*2,\s*3'
tests_raw = re.split(r'ĐỀ\s+ÔN\s+TẬP\s+CHƯƠNG\s+1,\s*2,\s*3', full_exam_text)
print(f"Số phần tách được bằng 'ĐỀ ÔN TẬP CHƯƠNG 1, 2, 3': {len(tests_raw)}")

# Phần 0 là phần trước đề 1 (nếu có)
# Từ phần 1 đến 10 là 10 đề
for idx in range(1, len(tests_raw)):
    t_text = tests_raw[idx]
    # Lấy tiêu đề ĐỀ SỐ ...
    header_m = re.search(r'ĐỀ\s+SỐ\s+(\d+)', t_text)
    header_title = header_m.group(0) if header_m else f"ĐỀ #{idx}"
    if idx == 10:
        header_title = "ĐỀ SỐ 10 (Gốc ghi nhầm Đề số 8)"
    
    # Kiểm tra các phần
    p1_m = re.search(r'PHẦN I\..*?(?=PHẦN II\.|$)', t_text, re.DOTALL)
    p2_m = re.search(r'PHẦN II\..*?(?=PHẦN III\.|$)', t_text, re.DOTALL)
    p3_m = re.search(r'PHẦN III\..*?$', t_text, re.DOTALL)
    
    p1_text = p1_m.group(0) if p1_m else ""
    p2_text = p2_m.group(0) if p2_m else ""
    p3_text = p3_m.group(0) if p3_m else ""
    
    # Đếm số câu hỏi trong từng phần
    c1 = len(re.findall(r'Câu\s+(\d+)\.', p1_text))
    c2 = len(re.findall(r'Câu\s+(\d+)\.', p2_text))
    c3 = len(re.findall(r'Câu\s+(\d+)\.', p3_text))
    print(f"Đề {idx:02d} ({header_title}): P1 = {c1} câu, P2 = {c2} câu, P3 = {c3} câu")
