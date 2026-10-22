# -*- coding: utf-8 -*-
"""
Tạo hoa12_to36_data.py hoàn chỉnh
"""
import os
import sys

DATA_FILE = r"c:\Antigravity_Thanh\scripts\hoa12_to36_data.py"

header_content = '''# -*- coding: utf-8 -*-
"""
Module dữ liệu chuẩn cho TỜ SỐ 36 - CHƯƠNG 5: PIN ĐIỆN VÀ ĐIỆN PHÂN (HÓA HỌC 12)
Thầy TRẦN VĂN THẠNH - 0777.470.803
Bản quyền tài liệu: Thầy Trần Văn Thạnh - TP Huế
"""

import os

ROOT_DIR = r"c:\\Antigravity_Thanh"
IMG_DIR = os.path.join(ROOT_DIR, "San_Pham", "images_to36")

INFO = {
    "title": "CHUYÊN ĐỀ 5: PIN ĐIỆN VÀ ĐIỆN PHÂN (TỜ SỐ 36)",
    "sub_title": "CHUYÊN ĐỀ ÔN TẬP VÀ PHÁT TRIỂN NĂNG LỰC HÓA HỌC 12",
    "subject": "HÓA HỌC 12 - CHƯƠNG TRÌNH MỚI 2026 - 2027",
    "grade": "12",
    "teacher": "Thầy TRẦN VĂN THẠNH",
    "phone": "0777.470.803",
    "branch1": "CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế",
    "branch2": "CS2: 24 Đặng Thái Thân, TP Huế",
}
'''

with open(DATA_FILE, "w", encoding="utf-8") as f:
    f.write(header_content)

print("Created header in hoa12_to36_data.py")
