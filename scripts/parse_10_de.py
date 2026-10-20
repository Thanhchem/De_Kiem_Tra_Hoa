# -*- coding: utf-8 -*-
"""
Trích xuất toàn bộ 10 đề thi từ file PDF 10-de-on-tap-C123.pdf
"""

import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import pymupdf
import re
import json

PDF_PATH = r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12\10-de-on-tap-C123.pdf"

doc = pymupdf.open(PDF_PATH)
full_text = []
for i, page in enumerate(doc):
    t = page.get_text()
    full_text.append((i + 1, t))

print(f"Tổng số trang đã đọc: {len(full_text)}")

# Trang 26-27 chứa đáp án
ans_text = full_text[25][1] + "\n" + full_text[26][1]

# In thử cấu trúc đáp án
print("=== TRÍCH XUẤT ĐÁP ÁN ===")
ans_blocks = re.split(r'ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI ĐỀ SỐ\s+(\d+)', ans_text)
print(f"Số block đáp án tìm thấy: {len(ans_blocks)}")
for k in range(1, len(ans_blocks), 2):
    de_num = ans_blocks[k]
    de_ans = ans_blocks[k+1]
    print(f"\n--- ĐÁP ÁN ĐỀ {de_num} ---")
    p1 = re.search(r'PHẦN I:\s*(.*?)(?=PHẦN II:|$)', de_ans, re.DOTALL)
    p2 = re.search(r'PHẦN II:\s*(.*?)(?=PHẦN III:|$)', de_ans, re.DOTALL)
    p3 = re.search(r'PHẦN III:\s*(.*?)(?=ĐÁP ÁN|$)', de_ans, re.DOTALL)
    if p1: print("P1:", p1.group(1).strip().replace('\n', ' ')[:80])
    if p2: print("P2:", p2.group(1).strip().replace('\n', ' ')[:80])
    if p3: print("P3:", p3.group(1).strip().replace('\n', ' ')[:80])
