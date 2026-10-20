# -*- coding: utf-8 -*-
"""
Trích xuất chi tiết từng câu hỏi, từng phương án của 10 đề thi
"""
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import pymupdf
import re
import json

PDF_PATH = r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\12\10-de-on-tap-C123.pdf"
doc = pymupdf.open(PDF_PATH)

all_text = ""
for i in range(25):
    all_text += doc[i].get_text() + "\n"

ans_text = doc[25].get_text() + "\n" + doc[26].get_text()

# Hàm parse đáp án
def parse_answers(text):
    ans_dict = {}
    blocks = re.split(r'ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI ĐỀ SỐ\s+(\d+)', text)
    for k in range(1, len(blocks), 2):
        de_num = int(blocks[k])
        de_ans = blocks[k+1]
        
        # P1
        p1_m = re.search(r'PHẦN I:\s*(.*?)(?=PHẦN II:|$)', de_ans, re.DOTALL)
        p1_ans = {}
        if p1_m:
            items = re.findall(r'(\d+)\.([A-D])', p1_m.group(1))
            for num, opt in items:
                p1_ans[int(num)] = opt
        
        # P2
        p2_m = re.search(r'PHẦN II:\s*(.*?)(?=PHẦN III:|$)', de_ans, re.DOTALL)
        p2_ans = {}
        if p2_m:
            items = re.findall(r'Câu\s+(\d+):\s*\[([ĐS]{4})\]', p2_m.group(1))
            for num, s in items:
                p2_ans[int(num)] = "--".join(list(s))
        
        # P3
        p3_m = re.search(r'PHẦN III:\s*(.*?)(?=ĐÁP ÁN|$)', de_ans, re.DOTALL)
        p3_ans = {}
        if p3_m:
            items = re.findall(r'Câu\s+(\d+):\s*([^|\n]+)', p3_m.group(1))
            for num, val in items:
                p3_ans[int(num)] = val.strip()
        
        ans_dict[de_num] = {
            'P1': p1_ans,
            'P2': p2_ans,
            'P3': p3_ans
        }
    return ans_dict

answers = parse_answers(ans_text)
print(f"Đã trích xuất đáp án cho {len(answers)} đề.")

# Hàm parse câu hỏi trắc nghiệm Phần I
def parse_part1_question(q_text):
    # Tìm các phương án A., B., C., D.
    # Pattern: A. ... B. ... C. ... D. ...
    opt_matches = list(re.finditer(r'([A-D])\.\s*', q_text))
    if len(opt_matches) < 4:
        # Thử tìm dạng A. B. C. D. không có khoảng trắng hoặc xuống dòng
        return q_text.strip(), {}
    
    stem = q_text[:opt_matches[0].start()].strip()
    opts = {}
    for i in range(len(opt_matches)):
        opt_char = opt_matches[i].group(1)
        start_pos = opt_matches[i].end()
        end_pos = opt_matches[i+1].start() if i+1 < len(opt_matches) else len(q_text)
        opts[opt_char] = q_text[start_pos:end_pos].strip()
    return stem, opts

# Hàm parse câu hỏi Đúng/Sai Phần II
def parse_part2_question(q_text):
    # Tìm a), b), c), d) hoặc a. b. c. d.
    item_matches = list(re.finditer(r'([a-d])[\)\.]\s*', q_text))
    if len(item_matches) < 4:
        return q_text.strip(), {}
    stem = q_text[:item_matches[0].start()].strip()
    items = {}
    for i in range(len(item_matches)):
        char = item_matches[i].group(1)
        start_pos = item_matches[i].end()
        end_pos = item_matches[i+1].start() if i+1 < len(item_matches) else len(q_text)
        items[char] = q_text[start_pos:end_pos].strip()
    return stem, items

tests_raw = re.split(r'ĐỀ\s+ÔN\s+TẬP\s+CHƯƠNG\s+1,\s*2,\s*3', all_text)[1:]
all_tests_data = []

for idx, t_raw in enumerate(tests_raw, 1):
    # Clean text
    t = re.sub(r'^\s*\d+\s*$', '', t_raw, flags=re.MULTILINE)
    t = re.sub(r'[-–—]{3,}.*?[-–—]{3,}', '', t)
    
    p1_m = re.search(r'PHẦN I\..*?(?=PHẦN II\.|$)', t, re.DOTALL)
    p2_m = re.search(r'PHẦN II\..*?(?=PHẦN III\.|$)', t, re.DOTALL)
    p3_m = re.search(r'PHẦN III\..*?$', t, re.DOTALL)
    
    p1_str = p1_m.group(0) if p1_m else ""
    p2_str = p2_m.group(0) if p2_m else ""
    p3_str = p3_m.group(0) if p3_m else ""
    
    # Parse P1
    q1_splits = list(re.finditer(r'Câu\s+(\d+)\.\s*', p1_str))
    p1_data = []
    for i in range(len(q1_splits)):
        num = int(q1_splits[i].group(1))
        start = q1_splits[i].end()
        end = q1_splits[i+1].start() if i+1 < len(q1_splits) else len(p1_str)
        q_raw = p1_str[start:end]
        stem, opts = parse_part1_question(q_raw)
        p1_data.append({'num': num, 'stem': stem, 'opts': opts})
        
    # Parse P2
    q2_splits = list(re.finditer(r'Câu\s+(\d+)\.\s*', p2_str))
    p2_data = []
    for i in range(len(q2_splits)):
        num = int(q2_splits[i].group(1))
        start = q2_splits[i].end()
        end = q2_splits[i+1].start() if i+1 < len(q2_splits) else len(p2_str)
        q_raw = p2_str[start:end]
        stem, items = parse_part2_question(q_raw)
        p2_data.append({'num': num, 'stem': stem, 'items': items})
        
    # Parse P3
    q3_splits = list(re.finditer(r'Câu\s+(\d+)\.\s*', p3_str))
    p3_data = []
    for i in range(len(q3_splits)):
        num = int(q3_splits[i].group(1))
        start = q3_splits[i].end()
        end = q3_splits[i+1].start() if i+1 < len(q3_splits) else len(p3_str)
        stem = p3_str[start:end].strip()
        p3_data.append({'num': num, 'stem': stem})
        
    all_tests_data.append({
        'test_num': idx,
        'p1': p1_data,
        'p2': p2_data,
        'p3': p3_data,
        'ans': answers.get(idx, {})
    })

print(f"Đã xử lý xong toàn bộ 10 đề thi.")
# Kiểm tra xem có câu nào bị thiếu phương án A,B,C,D không
for d in all_tests_data:
    de_num = d['test_num']
    for q in d['p1']:
        if len(q['opts']) != 4:
            print(f"Cảnh báo Đề {de_num} P1 Câu {q['num']}: chỉ có {len(q['opts'])} phương án!")
    for q in d['p2']:
        if len(q['items']) != 4:
            print(f"Cảnh báo Đề {de_num} P2 Câu {q['num']}: chỉ có {len(q['items'])} ý!")
    if len(d['p3']) != 6:
        print(f"Cảnh báo Đề {de_num} P3: chỉ có {len(d['p3'])} câu!")

# Ghi ra file JSON tạm
with open(r"c:\Antigravity_Thanh\scratch\all_10_tests.json", "w", encoding="utf-8") as f:
    json.dump(all_tests_data, f, ensure_ascii=False, indent=2)
print("Đã lưu all_10_tests.json thành công!")
