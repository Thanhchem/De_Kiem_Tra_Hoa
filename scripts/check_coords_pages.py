import re
import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('scratch/decompressed_press.xml', 'r', encoding='utf-8', errors='ignore') as f:
    text = f.read()

pages = text.split('<page ')

print("=== Checking Page index 105 (page_106.png) around Cau 16 ===")
for b in re.findall(r'<block\s+l="(\d+)"\s+t="(\d+)"\s+w="(\d+)"\s+h="(\d+)"\s*>(.*?)</block>', pages[105], re.DOTALL):
    txt = " ".join(re.findall(r'<text[^>]*>(.*?)</text>', b[4], re.DOTALL))
    if any(k in txt for k in ['Câu 16', 'cấu tạo như hình', 'Số liên kết', 'Câu 17']):
        print(f"[{b[0]},{b[1]},{b[2]},{b[3]}] {txt}")

print("\n=== Checking Page index 106 (page_107.png) around Cau 1 Part II ===")
for b in re.findall(r'<block\s+l="(\d+)"\s+t="(\d+)"\s+w="(\d+)"\s+h="(\d+)"\s*>(.*?)</block>', pages[106], re.DOTALL):
    txt = " ".join(re.findall(r'<text[^>]*>(.*?)</text>', b[4], re.DOTALL))
    if any(k in txt for k in ['Câu 1:', 'quy trình Kroll', 'Phương trình hoá học của phản ứng A', 'Hướng dẫn giải']):
        print(f"[{b[0]},{b[1]},{b[2]},{b[3]}] {txt}")

print("\n=== Checking Page index 107 (page_108.png) around Cau 4 Part II and Cau 1 Part III ===")
for b in re.findall(r'<block\s+l="(\d+)"\s+t="(\d+)"\s+w="(\d+)"\s+h="(\d+)"\s*>(.*?)</block>', pages[107], re.DOTALL):
    txt = " ".join(re.findall(r'<text[^>]*>(.*?)</text>', b[4], re.DOTALL))
    if any(k in txt for k in ['Câu 4:', 'Mỗi phát biểu sau', 'PHẦN III', 'Câu 1:', 'Có bao nhiêu liên kết sigma']):
        print(f"[{b[0]},{b[1]},{b[2]},{b[3]}] {txt}")
