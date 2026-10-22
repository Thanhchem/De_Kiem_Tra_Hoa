# -*- coding: utf-8 -*-
"""
Master Builder cho hoa12_to36_data.py
Đảm bảo 100% cú pháp Python hợp lệ, chuẩn UTF-8, không lỗi quote.
"""

import os
import sys

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

DATA_FILE = r"c:\Antigravity_Thanh\scripts\hoa12_to36_data.py"

# Bước 1: Header + Dạng 1 + Dạng 2
print("1. Đang ghi Header, Dạng 1 và Dạng 2...")
exec(open(r"c:\Antigravity_Thanh\scripts\build_hoa12_to36_data.py", encoding="utf-8").read())
exec(open(r"c:\Antigravity_Thanh\scripts\append_dang1_dang2.py", encoding="utf-8").read())

# Bước 2: Dạng 3 + Dạng 4 (đã sửa quote bên trong)
print("2. Đang ghi Dạng 3 và Dạng 4...")
# Đọc append_dang3_dang4.py và thay các quote bên trong
with open(r"c:\Antigravity_Thanh\scripts\append_dang3_dang4.py", "r", encoding="utf-8") as f:
    c34 = f.read()
c34 = c34.replace('"khi điện phân, chỉ có tạp chất trong đồng xu tan hết vào trong dung dịch, còn lại sẽ là đồng tinh khiết"',
                  "'khi điện phân, chỉ có tạp chất trong đồng xu tan hết vào trong dung dịch, còn lại sẽ là đồng tinh khiết'")
with open(r"c:\Antigravity_Thanh\scripts\append_dang3_dang4.py", "w", encoding="utf-8") as f:
    f.write(c34)

exec(open(r"c:\Antigravity_Thanh\scripts\append_dang3_dang4.py", encoding="utf-8").read())

# Bước 3: Tổng hợp (đã sửa quote bên trong)
print("3. Đang ghi Bài tập tổng hợp...")
with open(r"c:\Antigravity_Thanh\scripts\append_tong_hop.py", "r", encoding="utf-8") as f:
    cth = f.read()
cth = cth.replace('"Nồng độ CuSO~4~ trước và sau quá trình mạ là không đổi"',
                  "'Nồng độ CuSO~4~ trước và sau quá trình mạ là không đổi'")
cth = cth.replace('"Nếu kim loại đồng trong mảnh nhẫn tan hết thì khối lượng mảnh nhẫn không giảm nữa"',
                  "'Nếu kim loại đồng trong mảnh nhẫn tan hết thì khối lượng mảnh nhẫn không giảm nữa'")
with open(r"c:\Antigravity_Thanh\scripts\append_tong_hop.py", "w", encoding="utf-8") as f:
    f.write(cth)

exec(open(r"c:\Antigravity_Thanh\scripts\append_tong_hop.py", encoding="utf-8").read())

# Bước 4: Phần 3 Trả lời ngắn
print("4. Đang ghi Phần 3 Trả lời ngắn...")
exec(open(r"c:\Antigravity_Thanh\scripts\append_phan3.py", encoding="utf-8").read())

print("Hoàn tất tạo hoa12_to36_data.py!")
