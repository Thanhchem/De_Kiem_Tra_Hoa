# -*- coding: utf-8 -*-
"""
Script phân tích dữ liệu 10 Đề ôn tập Chương 1, 2, 3 Hóa học 12,
chuẩn hóa công thức hóa học, sinh lời giải chi tiết và tạo các file Word.
"""

import os
import sys
import re
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Import mô-đun chuẩn
sys.path.append(os.path.abspath(r"C:\Antigravity_Thanh\scripts"))
from form_de_kiem_tra import (
    setup_header_footer,
    build_title_block,
    add_section_header,
    render_question_prompt,
    render_options_no_table,
    render_statement,
    add_solution_no_table,
    add_student_writing_lines,
    render_rich_text,
    set_run_style,
    COLOR_PRIMARY,
    COLOR_CORRECT,
    COLOR_SOLUTION
)

from normalize_chem import normalize_chem_text

RAW_FILE = r"C:\Antigravity_Thanh\scripts\raw_10_de.txt"
OUT_DIR = r"C:\Antigravity_Thanh\San_Pham"
SUB_DIR = os.path.join(OUT_DIR, "10_De_On_Tap_Chuong_123_Hoa12")
os.makedirs(SUB_DIR, exist_ok=True)

def parse_all_exams():
    with open(RAW_FILE, "r", encoding="utf-8") as f:
        content = f.read()

    # Tách 10 đề dựa trên dòng bắt đầu bằng 'ĐỀ SỐ X'
    parts = re.split(r'(?m)^ĐỀ SỐ\s+(\d+)\s*$', content)
    exams = []
    
    # parts[0] là header tổng hợp, sau đó là [ '1', de_1_text, '2', de_2_text, ... ]
    for i in range(1, len(parts), 2):
        de_num = int(parts[i])
        de_text = parts[i+1]
        
        # Parse đề
        exam = parse_single_exam(de_num, de_text)
        exams.append(exam)
        
    return exams

def parse_single_exam(de_num, text):
    # Tách các phần: PHẦN I, PHẦN II, PHẦN III, ĐÁP ÁN
    p1_match = re.search(r'PHẦN I\..*?(?=PHẦN II\.)', text, re.DOTALL)
    p2_match = re.search(r'PHẦN II\..*?(?=PHẦN III\.)', text, re.DOTALL)
    p3_match = re.search(r'PHẦN III\..*?(?=ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI)', text, re.DOTALL)
    ans_match = re.search(r'ĐÁP ÁN VÀ HƯỚNG DẪN GIẢI.*?(?=――――|$)', text, re.DOTALL)

    p1_text = p1_match.group(0) if p1_match else ""
    p2_text = p2_match.group(0) if p2_match else ""
    p3_text = p3_match.group(0) if p3_match else ""
    ans_text = ans_match.group(0) if ans_match else ""

    # Parse đáp án
    ans_p1 = {}
    for m in re.finditer(r'(\d+)\.([A-D])', ans_text):
        ans_p1[int(m.group(1))] = m.group(2)

    ans_p2 = {}
    for m in re.finditer(r'Câu\s+(\d+):\s*\[([ĐS]{4})\]', ans_text):
        c_num = int(m.group(1))
        ans_p2[c_num] = list(m.group(2)) # ['Đ', 'Đ', 'S', 'Đ']

    ans_p3 = {}
    p3_ans_line = ""
    p3_m = re.search(r'PHẦN III:\s*(.*)', ans_text)
    if p3_m:
        p3_ans_line = p3_m.group(1)
        for m in re.finditer(r'Câu\s+(\d+):\s*([^\|]+)', p3_ans_line):
            c_num = int(m.group(1))
            val = m.group(2).strip()
            ans_p3[c_num] = val

    # 1. Parse Phần I (18 câu)
    part1_questions = []
    q_matches = list(re.finditer(r'Câu\s+(\d+)\.\s*(.*?)(?=Câu\s+\d+\.|$)', p1_text, re.DOTALL))
    for qm in q_matches:
        q_idx = int(qm.group(1))
        q_body = qm.group(2).strip()
        
        # Tách options A, B, C, D
        # Pattern: A. ... B. ... C. ... D. ...
        opt_m = re.search(r'A\.\s*(.*?)\s+B\.\s*(.*?)\s+C\.\s*(.*?)\s+D\.\s*(.*?)$', q_body, re.DOTALL)
        if opt_m:
            prompt = q_body[:opt_m.start()].strip()
            opts = [
                "A. " + normalize_chem_text(opt_m.group(1).strip()),
                "B. " + normalize_chem_text(opt_m.group(2).strip()),
                "C. " + normalize_chem_text(opt_m.group(3).strip()),
                "D. " + normalize_chem_text(opt_m.group(4).strip())
            ]
        else:
            prompt = q_body
            opts = []

        correct_opt = ans_p1.get(q_idx, "A")
        part1_questions.append({
            "idx": q_idx,
            "prompt": normalize_chem_text(prompt),
            "opts": opts,
            "correct": correct_opt
        })

    # 2. Parse Phần II (4 câu)
    part2_questions = []
    q2_matches = list(re.finditer(r'Câu\s+(\d+)\.\s*(.*?)(?=Câu\s+\d+\.|$)', p2_text, re.DOTALL))
    for qm in q2_matches:
        q_idx = int(qm.group(1))
        q_body = qm.group(2).strip()
        
        # Tìm các ý a), b), c), d)
        a_m = re.search(r'a\)\s*(.*?)(?=b\)|\Z)', q_body, re.DOTALL)
        b_m = re.search(r'b\)\s*(.*?)(?=c\)|\Z)', q_body, re.DOTALL)
        c_m = re.search(r'c\)\s*(.*?)(?=d\)|\Z)', q_body, re.DOTALL)
        d_m = re.search(r'd\)\s*(.*?)$', q_body, re.DOTALL)

        intro = q_body[:a_m.start()].strip() if a_m else q_body

        c_ans = ans_p2.get(q_idx, ['Đ', 'Đ', 'Đ', 'Đ'])
        if len(c_ans) < 4:
            c_ans = ['Đ', 'Đ', 'Đ', 'Đ']

        items = [
            ("a) " + normalize_chem_text(a_m.group(1).strip()) if a_m else "", "Đúng" if c_ans[0] == 'Đ' else "Sai"),
            ("b) " + normalize_chem_text(b_m.group(1).strip()) if b_m else "", "Đúng" if c_ans[1] == 'Đ' else "Sai"),
            ("c) " + normalize_chem_text(c_m.group(1).strip()) if c_m else "", "Đúng" if c_ans[2] == 'Đ' else "Sai"),
            ("d) " + normalize_chem_text(d_m.group(1).strip()) if d_m else "", "Đúng" if c_ans[3] == 'Đ' else "Sai"),
        ]

        part2_questions.append({
            "idx": q_idx,
            "intro": normalize_chem_text(intro),
            "items": items
        })

    # 3. Parse Phần III (6 câu)
    part3_questions = []
    q3_matches = list(re.finditer(r'Câu\s+(\d+)\.\s*(.*?)(?=Câu\s+\d+\.|$)', p3_text, re.DOTALL))
    for qm in q3_matches:
        q_idx = int(qm.group(1))
        q_body = qm.group(2).strip()
        ans_val = ans_p3.get(q_idx, "")
        
        part3_questions.append({
            "idx": q_idx,
            "prompt": normalize_chem_text(q_body),
            "ans": ans_val
        })

    return {
        "de_num": de_num,
        "part1": part1_questions,
        "part2": part2_questions,
        "part3": part3_questions
    }

print("Loaded parser successfully.")
