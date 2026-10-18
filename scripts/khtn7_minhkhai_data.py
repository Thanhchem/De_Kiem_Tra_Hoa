# -*- coding: utf-8 -*-
"""
Dữ liệu tổng hợp Đề Cương Ôn Tập Giữa Kì I KHTN 7
Trường THCS Nguyễn Thị Minh Khai - Năm học 2026 - 2027
Thầy TRẦN VĂN THẠNH - 0777.470.803
"""

import sys
import os

sys.path.append(os.path.dirname(__file__))
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from khtn7_minhkhai_data_part1 import PART1_P1
from khtn7_minhkhai_data_part2 import PART1_P2
from khtn7_minhkhai_data_part3 import TF_QUESTIONS, SHORT_QUESTIONS
from khtn7_minhkhai_data_part4 import ESSAY_QUESTIONS

INFO = {
    "department": "UBND PHƯỜNG THỦY XUÂN",
    "school": "TRƯỜNG THCS NGUYỄN THỊ MINH KHAI",
    "title": "ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA HỌC KÌ I",
    "subject": "MÔN: KHOA HỌC TỰ NHIÊN 7",
    "year": "Năm học: 2026 - 2027",
    "time_str": "Thời gian ôn tập & làm bài: 45 - 90 phút"
}

# Tổng hợp 90 câu trắc nghiệm nhiều lựa chọn Phần I
MC_QUESTIONS = PART1_P1 + PART1_P2

# Kiểm tra tính toàn vẹn của dữ liệu
assert len(MC_QUESTIONS) == 90, f"Lỗi số lượng câu Phần I: {len(MC_QUESTIONS)} != 90"
assert len(TF_QUESTIONS) == 2, f"Lỗi số lượng câu Phần II: {len(TF_QUESTIONS)} != 2"
assert len(SHORT_QUESTIONS) == 38, f"Lỗi số lượng câu Phần III: {len(SHORT_QUESTIONS)} != 38"
assert len(ESSAY_QUESTIONS) == 18, f"Lỗi số lượng câu Phần Tự luận: {len(ESSAY_QUESTIONS)} != 18"

if __name__ == "__main__":
    print(f"Tổng số câu Phần I (Trắc nghiệm 4 lựa chọn): {len(MC_QUESTIONS)}")
    print(f"Tổng số câu Phần II (Trắc nghiệm Đúng - Sai): {len(TF_QUESTIONS)}")
    print(f"Tổng số câu Phần III (Trắc nghiệm Trả lời ngắn): {len(SHORT_QUESTIONS)}")
    print(f"Tổng số câu Phần Tự luận: {len(ESSAY_QUESTIONS)}")
    print("Dữ liệu hợp lệ 100%!")
