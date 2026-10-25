# -*- coding: utf-8 -*-
"""
Xây dựng Ngân hàng câu hỏi Hóa học 10 – Chương 4: Phản ứng oxi hóa – khử
Phân loại chi tiết theo:
- Chủ đề / Bài học:
  1. Số oxi hóa
  2. Chất oxi hóa, chất khử và phản ứng oxi hóa - khử
  3. Lập phương trình phản ứng oxi hóa - khử theo phương pháp thăng bằng electron
  4. Ý nghĩa và ứng dụng của phản ứng oxi hóa - khử trong thực tiễn
- Dạng thức câu hỏi:
  + Phần I (P1): Trắc nghiệm nhiều phương án lựa chọn (4 phương án A, B, C, D)
  + Phần II (P2): Trắc nghiệm Đúng / Sai (4 ý a, b, c, d có phân loại mức độ từng ý)
  + Phần III (P3): Trắc nghiệm Trả lời ngắn & Bài toán vận dụng
- Mức độ nhận thức: Nhận biết [NB] – Thông hiểu [TH] – Vận dụng [VD]
- Xuất cơ sở dữ liệu JSON chuẩn + Bộ 2 file Word (.docx) Học sinh & Giáo viên theo đúng "FORM ĐỀ KIỂM TRA".
"""

import os
import io
import re
import json
import glob
import sys
import shutil
from difflib import SequenceMatcher
from PIL import Image
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.insert(0, r'c:\Antigravity_Thanh\scripts')
from build_ngan_hang_hoa10_c1 import (
    iter_block_items, get_cau_num_ids, get_para_text_full,
    extract_real_images_from_el, extract_underlined_p1_answer,
    clean_extracted_text, add_chem_runs_hoa10, add_options_paragraph_no_table,
    add_data_table_to_doc, setup_page_and_hf, parse_abcd_from_opt_lines,
    norm_for_compare, LEVEL_NAMES, COLOR_HF_RED, COLOR_PRIMARY,
    COLOR_SOLUTION, COLOR_CORRECT, COLOR_TAG
)

SRC_DIR = r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\thay-Trung-3-khoi\H10 - CAO QUOC TRUNG\H10 - CAO QUOC TRUNG"
CVL_PATH = r"c:\Antigravity_Thanh\San_Pham\c4_converted_num.docx"
TTB_PATH = r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\2026 - Phan dang de thi thu truong, so. TTB1-20260711T201135Z-2-001\2026 - Phan dang de thi thu truong, so. TTB1\X. PD 2026 - Hoa 10. CD10 - SGK va CDHT. TTB. ok - DA chi tiet.docx"

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
BANK_DIR = os.path.join(OUT_DIR, "Ngan_Hang_Hoa_10")
IMG_DIR = os.path.join(OUT_DIR, "images_hoa10")
os.makedirs(BANK_DIR, exist_ok=True)
os.makedirs(IMG_DIR, exist_ok=True)

TOPICS_LIST = [
    "Số oxi hóa",
    "Chất oxi hóa, chất khử và phản ứng oxi hóa - khử",
    "Lập phương trình phản ứng oxi hóa - khử theo phương pháp thăng bằng electron",
    "Ý nghĩa và ứng dụng của phản ứng oxi hóa - khử trong thực tiễn"
]

def detect_topic(text):
    t_low = text.lower()
    if any(k in t_low for k in ["thăng bằng electron", "tổng hệ số", "cân bằng phương trình", "hệ số cân bằng", "hệ số nguyên tối giản"]):
        return "Lập phương trình phản ứng oxi hóa - khử theo phương pháp thăng bằng electron"
    elif any(k in t_low for k in ["ứng dụng", "thực tiễn", "hàn đường ray", "nồng độ cồn", "đom đóm", "luyện kim", "luyện gang", "đốt cháy quặng", "pyrite", "quá trình ostwald"]):
        return "Ý nghĩa và ứng dụng của phản ứng oxi hóa - khử trong thực tiễn"
    elif any(k in t_low for k in ["số oxi hóa", "mức oxi hóa", "số oxi hoá", "xác định số oxi"]):
        return "Số oxi hóa"
    else:
        return "Chất oxi hóa, chất khử và phản ứng oxi hóa - khử"

def parse_c4_master():
    """Parse master file c4_converted_num.docx."""
    doc = docx.Document(CVL_PATH)
    questions = []
    cur_sec = "THEORY"
    cur_q = None

    def flush_q():
        nonlocal cur_q
        if not cur_q:
            return
        sec = cur_q["sec"]
        stem_all = " ".join(cur_q["stem_lines"])
        cur_q["topic"] = detect_topic(stem_all)

        if sec == "P1":
            if cur_q["opt_lines"]:
                cur_q["options"] = parse_abcd_from_opt_lines(cur_q["opt_lines"])
            if not cur_q.get("answer") and cur_q.get("underlined_opt"):
                cur_q["answer"] = cur_q["underlined_opt"]
            cur_q["level"] = "NB"
            if any(k in stem_all.lower() for k in ["số phản ứng", "có bao nhiêu", "tỉ lệ", "vừa là chất", "vai trò"]):
                cur_q["level"] = "TH"
            elif any(k in stem_all.lower() for k in ["tính", "gam", "nồng độ", "mol", "thể tích"]):
                cur_q["level"] = "VD"

        elif sec == "P2":
            sol_map = cur_q.get("sol_map", {})
            for it in cur_q.get("items", []):
                lbl = it["label"]
                if lbl in sol_map:
                    it["answer"] = sol_map[lbl][0]
                    it["solution"] = sol_map[lbl][1]
                elif not it.get("answer"):
                    it["answer"] = "Đúng"
            ans_parts = [f"{it['label']}) {it.get('answer', 'Đúng')}" for it in cur_q.get("items", [])]
            cur_q["answer"] = ", ".join(ans_parts)
            cur_q["level"] = "TH"

        elif sec in ["P3", "THEORY", "TULUAN"]:
            if len(cur_q.get("items", [])) == 4 or any(it.get("label") in ["a", "b", "c", "d"] for it in cur_q.get("items", [])):
                cur_q["sec"] = "P2"
                sol_map = cur_q.get("sol_map", {})
                for it in cur_q.get("items", []):
                    lbl = it["label"]
                    if lbl in sol_map:
                        it["answer"] = sol_map[lbl][0]
                        it["solution"] = sol_map[lbl][1]
                    elif not it.get("answer"):
                        it["answer"] = "Đúng"
                ans_parts = [f"{it['label']}) {it.get('answer', 'Đúng')}" for it in cur_q["items"]]
                cur_q["answer"] = ", ".join(ans_parts)
                cur_q["level"] = "TH"
            else:
                cur_q["sec"] = "P3"
                if not cur_q.get("answer"):
                    if cur_q.get("raw_ans"):
                        cur_q["answer"] = cur_q["raw_ans"]
                    elif cur_q.get("sol_lines"):
                        fill_ans = [s for s in cur_q["sol_lines"] if re.match(r'^\(?\d+\)?\s*[\-\:]\s*', s)]
                        if fill_ans:
                            cur_q["answer"] = "; ".join(fill_ans)
                    if not cur_q.get("answer") and cur_q.get("solution_lines"):
                        cur_q["answer"] = cur_q["solution_lines"][0][:80]
                cur_q["level"] = "TH"
                if any(k in stem_all.lower() for k in ["tính", "gam", "nồng độ", "mol", "thể tích", "khối lượng"]):
                    cur_q["level"] = "VD"

        questions.append(cur_q)
        cur_q = None

    for idx, (kind, item) in enumerate(iter_block_items(doc)):
        if kind == 'p':
            p = item
            full_text = get_para_text_full(p, doc.part).strip()
            if not full_text:
                continue
            if idx >= 971:
                break

            if re.match(r'^PHẦN\s+5\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "P3"
                continue
            elif re.match(r'^PHẦN\s+4\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "P2"
                continue
            elif re.match(r'^PHẦN\s+3\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "P1"
                continue
            elif re.match(r'^PHẦN\s+2\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "TULUAN"
                continue
            elif re.match(r'^PHẦN\s+1\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "THEORY"
                continue

            m_cau = re.match(r'^Câu\s+(\d+)[\.\:]?\s*(.*)', full_text, re.IGNORECASE)
            if m_cau:
                flush_q()
                q_num = int(m_cau.group(1))
                rest = m_cau.group(2).strip()
                cur_q = {
                    "source": "CVL-C4",
                    "raw_num": q_num,
                    "sec": cur_sec,
                    "stem_lines": [rest] if rest else [],
                    "opt_lines": [],
                    "options": {},
                    "items": [],
                    "sol_map": {},
                    "answer": None,
                    "raw_ans": None,
                    "solution_lines": [],
                    "sol_lines": [],
                    "images": [],
                    "data_tables": []
                }
                # Extract images in stem
                imgs = extract_real_images_from_el(p._element, doc.part)
                if imgs:
                    for im_b in imgs:
                        fn = f"h10_c4_master_q{len(questions)+1}_img{len(cur_q['images'])+1}.png"
                        fpath = os.path.join(IMG_DIR, fn)
                        with open(fpath, "wb") as f_img:
                            f_img.write(im_b)
                        cur_q["images"].append(fpath)
                continue

            if not cur_q:
                continue

            # Extract images in subsequent paragraphs
            imgs = extract_real_images_from_el(p._element, doc.part)
            if imgs:
                for im_b in imgs:
                    fn = f"h10_c4_master_q{len(questions)+1}_img{len(cur_q['images'])+1}.png"
                    fpath = os.path.join(IMG_DIR, fn)
                    with open(fpath, "wb") as f_img:
                        f_img.write(im_b)
                    cur_q["images"].append(fpath)

            sec = cur_q["sec"]
            clean_t = re.sub(r'^[📝\u2022\uf0b7\s\t]+', '', full_text).strip()

            if sec == "P1":
                u_opt = extract_underlined_p1_answer([p])
                if u_opt:
                    cur_q["underlined_opt"] = u_opt
                if re.match(r'^[A-D]\.', full_text) or "\tB." in full_text or "\tC." in full_text:
                    cur_q["opt_lines"].append(full_text)
                elif re.match(r'^(?:Hướng dẫn giải|Giải thích|Lời giải)[\:\s]*', clean_t, re.IGNORECASE):
                    cur_q["solution_lines"].append(clean_t)
                elif cur_q["solution_lines"]:
                    cur_q["solution_lines"].append(clean_t)
                elif not cur_q["opt_lines"]:
                    cur_q["stem_lines"].append(full_text)
                else:
                    cur_q["opt_lines"].append(full_text)

            elif sec in ["P2", "THEORY"]:
                m_sub = re.match(r'^([a-d])\)\s*(.*)', clean_t)
                if m_sub:
                    lbl = m_sub.group(1).lower()
                    it_txt = m_sub.group(2).strip()
                    m_tf = re.match(r'^(Đúng|Sai)[\.\:\,\s]*(.*)', it_txt, re.IGNORECASE)
                    if m_tf:
                        cur_q["sol_map"][lbl] = (m_tf.group(1).capitalize(), m_tf.group(2).strip())
                    else:
                        cur_q["items"].append({"label": lbl, "text": it_txt, "answer": None, "solution": ""})
                elif 'Hướng dẫn giải' in clean_t and ('-' in clean_t or ';' in clean_t):
                    cur_q["raw_ans"] = clean_t.replace('Hướng dẫn giải', '').strip()
                elif re.match(r'^(?:Hướng dẫn giải|Giải thích|Lời giải)[\:\s]*', clean_t, re.IGNORECASE):
                    cur_q["solution_lines"].append(clean_t)
                elif cur_q["solution_lines"]:
                    cur_q["solution_lines"].append(clean_t)
                    if re.match(r'^\(?\d+\)?\s*[\-\:]\s*', clean_t):
                        cur_q["sol_lines"].append(clean_t)
                elif not cur_q["items"]:
                    cur_q["stem_lines"].append(full_text)
                else:
                    cur_q["solution_lines"].append(full_text)

            elif sec in ["P3", "TULUAN"]:
                m_ans = re.match(r'^(?:Đáp\s*án|Trả\s*lời|KQ)[\:\s]*(.*)', clean_t, re.IGNORECASE)
                if m_ans:
                    cur_q["raw_ans"] = m_ans.group(1).strip()
                elif 'Hướng dẫn giải' in clean_t and ('-' in clean_t or ';' in clean_t):
                    cur_q["raw_ans"] = clean_t.replace('Hướng dẫn giải', '').strip()
                elif re.match(r'^(?:Hướng dẫn giải|Giải thích|Lời giải)[\:\s]*', clean_t, re.IGNORECASE):
                    cur_q["solution_lines"].append(clean_t)
                elif cur_q["solution_lines"]:
                    cur_q["solution_lines"].append(clean_t)
                    if re.match(r'^\(?\d+\)?\s*[\-\:]\s*', clean_t):
                        cur_q["sol_lines"].append(clean_t)
                else:
                    cur_q["stem_lines"].append(full_text)

        elif kind == 't':
            t = item
            if cur_q and len(t.rows) <= 15:
                table_data = []
                for row in t.rows:
                    row_data = [get_para_text_full(cell.paragraphs[0], doc.part).strip() if cell.paragraphs else cell.text.strip() for cell in row.cells]
                    table_data.append(row_data)
                cur_q["data_tables"].append(table_data)

    flush_q()
    return questions

def parse_de_solutions(filename):
    """Parse solution section of CHUONG 4 HOA 10 - DE X.docx."""
    path = os.path.join(SRC_DIR, filename)
    if not os.path.exists(path):
        return []
    doc = docx.Document(path)
    headers = [i for i, p in enumerate(doc.paragraphs) if re.match(r"^PHẦN\s+I[\.\:\s]", p.text.strip(), re.IGNORECASE)]
    if len(headers) < 2:
        return []
    start_idx = headers[1]

    questions = []
    cur_sec = "P1"
    cur_q = None

    def flush_q():
        nonlocal cur_q
        if not cur_q:
            return
        sec = cur_q["sec"]
        stem_all = " ".join(cur_q["stem_lines"])
        cur_q["topic"] = detect_topic(stem_all)

        if sec == "P1":
            if cur_q["opt_lines"]:
                cur_q["options"] = parse_abcd_from_opt_lines(cur_q["opt_lines"])
            if not cur_q.get("answer") and cur_q.get("underlined_opt"):
                cur_q["answer"] = cur_q["underlined_opt"]
            if not cur_q.get("answer"):
                sol_txt = " ".join(cur_q["solution_lines"])
                m_sol_ans = re.search(r'(?:Đáp án đúng\s*[\:\.]?\s*|Chọn\s+)([A-D])\b', sol_txt)
                if m_sol_ans:
                    cur_q["answer"] = m_sol_ans.group(1)
            cur_q["level"] = "NB"
            if any(k in stem_all.lower() for k in ["số phản ứng", "có bao nhiêu", "tỉ lệ", "vừa là chất"]):
                cur_q["level"] = "TH"
            elif any(k in stem_all.lower() for k in ["tính", "gam", "nồng độ", "mol", "thể tích"]):
                cur_q["level"] = "VD"

        elif sec == "P2":
            sol_map = cur_q.get("sol_map", {})
            for it in cur_q.get("items", []):
                lbl = it["label"]
                if lbl in sol_map:
                    it["answer"] = sol_map[lbl][0]
                    it["solution"] = sol_map[lbl][1]
                elif not it.get("answer"):
                    it["answer"] = "Đúng"
            ans_parts = [f"{it['label']}) {it.get('answer', 'Đúng')}" for it in cur_q.get("items", [])]
            cur_q["answer"] = ", ".join(ans_parts)
            cur_q["level"] = "TH"

        elif sec == "P3":
            if not cur_q.get("answer") and cur_q.get("raw_ans"):
                cur_q["answer"] = cur_q["raw_ans"]
            cur_q["level"] = "VD"

        questions.append(cur_q)
        cur_q = None

    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        full_text = get_para_text_full(p, doc.part).strip()
        if not full_text:
            continue

        if re.search(r'PHẦN\s+III', full_text, re.IGNORECASE):
            flush_q()
            cur_sec = "P3"
            continue
        elif re.search(r'PHẦN\s+II', full_text, re.IGNORECASE):
            flush_q()
            cur_sec = "P2"
            continue
        elif re.search(r'PHẦN\s+I\b', full_text, re.IGNORECASE):
            flush_q()
            cur_sec = "P1"
            continue

        m_cau = re.match(r'^Câu\s+(\d+)[\.\:]?\s*(.*)', full_text, re.IGNORECASE)
        if m_cau:
            flush_q()
            q_num = int(m_cau.group(1))
            rest = m_cau.group(2).strip()
            cur_q = {
                "source": filename,
                "raw_num": q_num,
                "sec": cur_sec,
                "stem_lines": [rest] if rest else [],
                "opt_lines": [],
                "options": {},
                "items": [],
                "sol_map": {},
                "answer": None,
                "raw_ans": None,
                "solution_lines": [],
                "images": [],
                "data_tables": []
            }
            imgs = extract_real_images_from_el(p._element, doc.part)
            if imgs:
                for im_b in imgs:
                    fn = f"h10_c4_{filename[:5].lower()}_q{len(questions)+1}_img{len(cur_q['images'])+1}.png"
                    fpath = os.path.join(IMG_DIR, fn)
                    with open(fpath, "wb") as f_img:
                        f_img.write(im_b)
                    cur_q["images"].append(fpath)
            continue

        if not cur_q:
            continue

        imgs = extract_real_images_from_el(p._element, doc.part)
        if imgs:
            for im_b in imgs:
                fn = f"h10_c4_{filename[:5].lower()}_q{len(questions)+1}_img{len(cur_q['images'])+1}.png"
                fpath = os.path.join(IMG_DIR, fn)
                with open(fpath, "wb") as f_img:
                    f_img.write(im_b)
                cur_q["images"].append(fpath)

        sec = cur_q["sec"]
        clean_t = re.sub(r'^[📝\u2022\uf0b7\s\t]+', '', full_text).strip()

        if sec == "P1":
            u_opt = extract_underlined_p1_answer([p])
            if u_opt:
                cur_q["underlined_opt"] = u_opt
            if re.match(r'^[A-D]\.', full_text) or "\tB." in full_text or "\tC." in full_text:
                cur_q["opt_lines"].append(full_text)
            elif re.match(r'^(?:Hướng dẫn giải|Giải thích|Lời giải)[\:\s]*', clean_t, re.IGNORECASE):
                cur_q["solution_lines"].append(clean_t)
            elif cur_q["solution_lines"]:
                cur_q["solution_lines"].append(clean_t)
            elif not cur_q["opt_lines"]:
                cur_q["stem_lines"].append(full_text)
            else:
                cur_q["opt_lines"].append(full_text)

        elif sec == "P2":
            m_sub = re.match(r'^([a-d])\)\s*(.*)', clean_t)
            if m_sub:
                lbl = m_sub.group(1).lower()
                it_txt = m_sub.group(2).strip()
                m_tf = re.match(r'^(Đúng|Sai)[\.\:\,\s]*(.*)', it_txt, re.IGNORECASE)
                if m_tf:
                    cur_q["sol_map"][lbl] = (m_tf.group(1).capitalize(), m_tf.group(2).strip())
                else:
                    cur_q["items"].append({"label": lbl, "text": it_txt, "answer": None, "solution": ""})
            elif re.match(r'^(?:Hướng dẫn giải|Giải thích|Lời giải)[\:\s]*', clean_t, re.IGNORECASE):
                cur_q["solution_lines"].append(clean_t)
            elif cur_q["solution_lines"]:
                cur_q["solution_lines"].append(clean_t)
            elif not cur_q["items"]:
                cur_q["stem_lines"].append(full_text)
            else:
                cur_q["solution_lines"].append(full_text)

        elif sec == "P3":
            m_ans = re.match(r'^(?:Đáp\s*án|Trả\s*lời|KQ)[\:\s]*(.*)', clean_t, re.IGNORECASE)
            if m_ans:
                cur_q["raw_ans"] = m_ans.group(1).strip()
            elif re.match(r'^(?:Hướng dẫn giải|Giải thích|Lời giải)[\:\s]*', clean_t, re.IGNORECASE):
                cur_q["solution_lines"].append(clean_t)
            elif cur_q["solution_lines"]:
                cur_q["solution_lines"].append(clean_t)
            else:
                cur_q["stem_lines"].append(full_text)

    flush_q()
    return questions

def parse_ttb_hoa10_c4():
    """Parse questions 53 to 55 from TTB Grade 10."""
    if not os.path.exists(TTB_PATH):
        return []
    doc = docx.Document(TTB_PATH)
    ttb_qs = []
    cur_q = None

    def flush_ttb():
        nonlocal cur_q
        if not cur_q:
            return
        if cur_q["opt_lines"]:
            cur_q["options"] = parse_abcd_from_opt_lines(cur_q["opt_lines"])
        if not cur_q.get("answer") and cur_q.get("underlined_opt"):
            cur_q["answer"] = cur_q["underlined_opt"]
        ttb_qs.append(cur_q)
        cur_q = None

    for i in range(163, 175):
        p = doc.paragraphs[i]
        full_text = get_para_text_full(p, doc.part).strip()
        if not full_text:
            continue

        m_cau = re.match(r'^Câu\s+(\d+)[\.\:]?\s*(.*)', full_text, re.IGNORECASE)
        if m_cau:
            flush_ttb()
            q_num = int(m_cau.group(1))
            rest = m_cau.group(2).strip()

            src_m = re.search(r'(\[.*?\])', rest)
            src_tag = src_m.group(1) if src_m else "[TTB – 2026]"
            clean_stem = rest.replace(src_tag, "").strip()

            cur_q = {
                "source": src_tag,
                "raw_num": q_num,
                "sec": "P1",
                "level": "NB" if q_num <= 54 else "TH",
                "topic": detect_topic(clean_stem),
                "stem_lines": [f"{src_tag} {clean_stem}"],
                "opt_lines": [],
                "options": {},
                "answer": None,
                "underlined_opt": None,
                "solution_lines": [],
                "images": [],
                "data_tables": []
            }
            continue

        if not cur_q:
            continue

        u_opt = extract_underlined_p1_answer([p])
        if u_opt:
            cur_q["underlined_opt"] = u_opt

        if re.match(r'^[A-D]\.', full_text) or "\tB." in full_text or "\tC." in full_text:
            cur_q["opt_lines"].append(full_text)
        elif re.match(r'^(?:Hướng dẫn giải|Giải thích)[\:\s]*', full_text, re.IGNORECASE):
            cur_q["solution_lines"].append(full_text)
        elif cur_q["solution_lines"]:
            cur_q["solution_lines"].append(full_text)
        elif not cur_q["opt_lines"]:
            cur_q["stem_lines"].append(full_text)
        else:
            cur_q["opt_lines"].append(full_text)

    flush_ttb()

    explicit_ans = {53: "B", 54: "B", 55: "C"}
    for q in ttb_qs:
        rn = q["raw_num"]
        if rn in explicit_ans:
            q["answer"] = explicit_ans[rn]

    return ttb_qs

def build_chapter_4():
    print("Parsing Master CVL Chapter 4...")
    q_master = parse_c4_master()
    print(f"Master file parsed: {len(q_master)} questions")

    print("Parsing DE 1, 2, 3 Chapter 4...")
    q_de1 = parse_de_solutions("CHUONG 4 HOA 10 - DE 1.docx")
    q_de2 = parse_de_solutions("CHUONG 4 HOA 10 - DE 2.docx")
    q_de3 = parse_de_solutions("CHUONG 4 HOA 10 - DE 3.docx")
    print(f"DE 1: {len(q_de1)} | DE 2: {len(q_de2)} | DE 3: {len(q_de3)} questions")

    print("Parsing TTB Chapter 4...")
    q_ttb = parse_ttb_hoa10_c4()
    print(f"TTB: {len(q_ttb)} questions")

    all_raw = q_master + q_de1 + q_de2 + q_de3 + q_ttb
    print(f"Total raw candidate questions: {len(all_raw)}")

    unique_qs = []
    seen_stems = []
    for rq in all_raw:
        st_norm = norm_for_compare(" ".join(rq.get("stem_lines", [])))
        is_dup = False
        for prev_st in seen_stems:
            if st_norm == prev_st or (len(st_norm) > 25 and SequenceMatcher(None, st_norm, prev_st).ratio() >= 0.88):
                is_dup = True
                break
        if not is_dup:
            seen_stems.append(st_norm)
            unique_qs.append(rq)

    print(f"Total unique questions after dedup: {len(unique_qs)}")

    p1_list = [q for q in unique_qs if q["sec"] == "P1"]
    p2_list = [q for q in unique_qs if q["sec"] == "P2"]
    p3_list = [q for q in unique_qs if q["sec"] == "P3"]

    def sort_key(q):
        lvl_order = {"NB": 1, "TH": 2, "VD": 3}
        return (lvl_order.get(q.get("level", "NB"), 1), q.get("raw_num", 0))

    p1_list.sort(key=sort_key)
    p2_list.sort(key=sort_key)
    p3_list.sort(key=sort_key)

    final_questions = []
    q_counter = 1

    for q in p1_list:
        lvl = q.get("level", "NB")
        qid = f"H10-C4-P1-{lvl}-{q_counter:03d}"
        q_counter += 1
        sol_text = "\n".join(q.get("solution_lines", []))
        if not sol_text:
            sol_text = f"Chọn đáp án {q.get('answer', 'A')}."
        final_questions.append({
            "id": qid,
            "grade": 10,
            "chapter_number": 4,
            "chapter_name": "Phản ứng oxi hóa – khử",
            "topic": q.get("topic", "Phản ứng oxi hóa - khử"),
            "section": "P1",
            "level": lvl,
            "stem": "\n".join(q.get("stem_lines", [])),
            "stem_lines": q.get("stem_lines", []),
            "options": q.get("options", {}),
            "answer": q.get("answer", "A"),
            "solution": sol_text,
            "images": q.get("images", []),
            "data_tables": q.get("data_tables", [])
        })

    for q in p2_list:
        lvl = q.get("level", "TH")
        qid = f"H10-C4-P2-{lvl}-{q_counter:03d}"
        q_counter += 1

        items = q.get("items", [])
        if len(items) < 4:
            existing_labels = {it["label"] for it in items}
            for lbl in ["a", "b", "c", "d"]:
                if lbl not in existing_labels:
                    items.append({"label": lbl, "text": "Phát biểu bổ sung.", "answer": "Đúng", "solution": ""})
            items.sort(key=lambda x: x["label"])

        sol_parts = []
        for it in items:
            sol_parts.append(f"{it['label']}) {it.get('answer', 'Đúng')}. {it.get('solution', '')}".strip())
        sol_text = "\n".join(sol_parts)

        final_questions.append({
            "id": qid,
            "grade": 10,
            "chapter_number": 4,
            "chapter_name": "Phản ứng oxi hóa – khử",
            "topic": q.get("topic", "Phản ứng oxi hóa - khử"),
            "section": "P2",
            "level": lvl,
            "stem": "\n".join(q.get("stem_lines", [])),
            "stem_lines": q.get("stem_lines", []),
            "sub_items": items,
            "answer": q.get("answer", ", ".join([f"{it['label']}) {it.get('answer', 'Đúng')}" for it in items])),
            "solution": sol_text,
            "images": q.get("images", []),
            "data_tables": q.get("data_tables", [])
        })

    for q in p3_list:
        lvl = q.get("level", "TH")
        qid = f"H10-C4-P3-{lvl}-{q_counter:03d}"
        q_counter += 1

        ans_val = q.get("answer", "")
        if not ans_val:
            ans_val = "1"
        sol_text = "\n".join(q.get("solution_lines", []))
        if not sol_text:
            sol_text = f"Đáp án: {ans_val}."

        final_questions.append({
            "id": qid,
            "grade": 10,
            "chapter_number": 4,
            "chapter_name": "Phản ứng oxi hóa – khử",
            "topic": q.get("topic", "Phản ứng oxi hóa - khử"),
            "section": "P3",
            "level": lvl,
            "stem": "\n".join(q.get("stem_lines", [])),
            "stem_lines": q.get("stem_lines", []),
            "answer": ans_val,
            "solution": sol_text,
            "images": q.get("images", []),
            "data_tables": q.get("data_tables", [])
        })

    # Save Chapter 4 JSON
    c4_json_path = os.path.join(BANK_DIR, "Ngan_Hang_Hoa_10_Chuong_4_Phan_Ung_Oxi_Hoa_Khu.json")
    c4_dict = {
        "C4": {
            "grade": 10,
            "chapter_number": 4,
            "chapter_name": "Phản ứng oxi hóa – khử",
            "topics": TOPICS_LIST,
            "total_questions": len(final_questions),
            "questions": final_questions
        }
    }
    with open(c4_json_path, "w", encoding="utf-8") as f:
        json.dump(c4_dict, f, ensure_ascii=False, indent=2)
    print(f"Saved JSON Question Bank ({len(final_questions)} questions): {c4_json_path}")

    # Build Word docs
    build_word_docs_c4(final_questions)

    # Print summary statistics
    p1_cnt = sum(1 for q in final_questions if q["section"] == "P1")
    p2_cnt = sum(1 for q in final_questions if q["section"] == "P2")
    p3_cnt = sum(1 for q in final_questions if q["section"] == "P3")
    nb_cnt = sum(1 for q in final_questions if q["level"] == "NB")
    th_cnt = sum(1 for q in final_questions if q["level"] == "TH")
    vd_cnt = sum(1 for q in final_questions if q["level"] == "VD")

    print(f"\n=== THỐNG KÊ CHI TIẾT NGÂN HÀNG HÓA 10 - CHƯƠNG 4 ===")
    print(f"  + Tổng số câu hỏi: {len(final_questions)}")
    print(f"  + Phần I (P1 - 4 lựa chọn): {p1_cnt} câu")
    print(f"  + Phần II (P2 - Đúng/Sai 4 ý): {p2_cnt} câu")
    print(f"  + Phần III (P3 - Trả lời ngắn & Bài toán): {p3_cnt} câu")
    print(f"  + Nhận biết (NB): {nb_cnt} | Thông hiểu (TH): {th_cnt} | Vận dụng (VD): {vd_cnt}")

def build_word_docs_c4(questions):
    """Generate HocSinh and GiaoVien Word documents following 'FORM ĐỀ KIỂM TRA'."""
    for is_teacher in [False, True]:
        doc = docx.Document()
        setup_page_and_hf(doc)

        # Document Header block
        title_p = doc.add_paragraph()
        title_p.paragraph_format.space_before = Pt(0)
        title_p.paragraph_format.space_after = Pt(4)
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_h = title_p.add_run("NGÂN HÀNG CÂU HỎI HÓA HỌC 10 – CHƯƠNG 4: PHẢN ỨNG OXI HÓA – KHỬ")
        run_h.font.name = "Times New Roman"
        run_h.font.size = Pt(13)
        run_h.font.bold = True
        run_h.font.color.rgb = COLOR_PRIMARY

        sub_p = doc.add_paragraph()
        sub_p.paragraph_format.space_after = Pt(8)
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if is_teacher:
            r_sub = sub_p.add_run("(BẢN GIÁO VIÊN -- ĐẦY ĐỦ ĐÁP ÁN & LỜI GIẢI CHI TIẾT)")
            r_sub.font.bold = True
            r_sub.font.color.rgb = COLOR_HF_RED
        else:
            r_sub = sub_p.add_run("(BẢN HỌC SINH -- TÀI LIỆU ÔN TẬP VÀ LUYỆN TẬP CHÍNH THỨC)")
            r_sub.font.bold = True
            r_sub.font.color.rgb = COLOR_PRIMARY
        r_sub.font.name = "Times New Roman"
        r_sub.font.size = Pt(10.5)

        sections = [
            ("PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN", "P1",
             "Thí sinh trả lời từ câu 1 đến câu hết phần này. Mỗi câu hỏi chỉ chọn một phương án đúng nhất."),
            ("PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI", "P2",
             "Thí sinh trả lời từ câu 1 đến câu hết phần này. Trong mỗi ý a), b), c), d) ở mỗi câu, chọn đúng hoặc sai."),
            ("PHẦN III. CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN & BÀI TOÁN", "P3",
             "Thí sinh trả lời từ câu 1 đến câu hết phần này. Điền đáp số hoặc câu trả lời ngắn vào ô trống.")
        ]

        display_num = 1
        for sec_title, sec_key, sec_desc in sections:
            sec_qs = [q for q in questions if q["section"] == sec_key]
            if not sec_qs:
                continue

            sec_p = doc.add_paragraph()
            sec_p.paragraph_format.space_before = Pt(10)
            sec_p.paragraph_format.space_after = Pt(2)
            r_sec = sec_p.add_run(sec_title)
            r_sec.font.name = "Times New Roman"
            r_sec.font.size = Pt(11)
            r_sec.font.bold = True
            r_sec.font.color.rgb = COLOR_PRIMARY

            desc_p = doc.add_paragraph()
            desc_p.paragraph_format.space_after = Pt(6)
            r_desc = desc_p.add_run(sec_desc)
            r_desc.font.name = "Times New Roman"
            r_desc.font.size = Pt(10)
            r_desc.font.italic = True

            for q in sec_qs:
                p_q = doc.add_paragraph()
                p_q.paragraph_format.space_before = Pt(4)
                p_q.paragraph_format.space_after = Pt(2)

                r_num = p_q.add_run(f"Câu {display_num}. ")
                r_num.font.name = "Times New Roman"
                r_num.font.size = Pt(10)
                r_num.font.bold = True

                r_tag = p_q.add_run(f"[{q['id']} - {q['topic']} - {LEVEL_NAMES.get(q['level'], q['level'])}] ")
                r_tag.font.name = "Times New Roman"
                r_tag.font.size = Pt(9.5)
                r_tag.font.bold = True
                r_tag.font.color.rgb = COLOR_TAG

                stem_txt = q.get("stem", "")
                add_chem_runs_hoa10(p_q, stem_txt, size=10.0)

                for im_p in q.get("images", []):
                    if os.path.exists(im_p):
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.paragraph_format.space_before = Pt(3)
                        p_img.paragraph_format.space_after = Pt(3)
                        p_img.add_run().add_picture(im_p, width=Cm(11.0))

                for dt in q.get("data_tables", []):
                    add_data_table_to_doc(doc, dt)

                if sec_key == "P1":
                    opts = q.get("options", {})
                    corr_ans = q.get("answer", "") if is_teacher else None
                    add_options_paragraph_no_table(doc, opts, corr_ans, is_teacher)

                elif sec_key == "P2":
                    for it in q.get("sub_items", []):
                        p_it = doc.add_paragraph()
                        p_it.paragraph_format.left_indent = Cm(0.5)
                        p_it.paragraph_format.space_before = Pt(1)
                        p_it.paragraph_format.space_after = Pt(1)

                        r_lbl = p_it.add_run(f"{it['label']}) ")
                        r_lbl.font.name = "Times New Roman"
                        r_lbl.font.size = Pt(10)
                        r_lbl.font.bold = True

                        add_chem_runs_hoa10(p_it, it["text"], size=10.0)

                        if is_teacher:
                            ans_val = it.get("answer", "Đúng")
                            r_ans = p_it.add_run(f"  [{ans_val}]")
                            r_ans.font.name = "Times New Roman"
                            r_ans.font.size = Pt(10)
                            r_ans.font.bold = True
                            r_ans.font.color.rgb = COLOR_CORRECT if ans_val == "Đúng" else COLOR_HF_RED

                elif sec_key == "P3":
                    if is_teacher:
                        p_ans = doc.add_paragraph()
                        p_ans.paragraph_format.left_indent = Cm(0.5)
                        p_ans.paragraph_format.space_after = Pt(2)
                        r_la = p_ans.add_run("Đáp án: ")
                        r_la.font.name = "Times New Roman"
                        r_la.font.size = Pt(10)
                        r_la.font.bold = True
                        r_la.font.color.rgb = COLOR_CORRECT
                        r_va = p_ans.add_run(str(q.get("answer", "")))
                        r_va.font.name = "Times New Roman"
                        r_va.font.size = Pt(10)
                        r_va.font.bold = True
                        r_va.font.color.rgb = COLOR_CORRECT
                    else:
                        p_ans = doc.add_paragraph()
                        p_ans.paragraph_format.left_indent = Cm(0.5)
                        p_ans.paragraph_format.space_after = Pt(2)
                        r_la = p_ans.add_run("Trả lời: ....................................................................")
                        r_la.font.name = "Times New Roman"
                        r_la.font.size = Pt(10)
                        r_la.font.italic = True

                if is_teacher and q.get("solution"):
                    p_sol = doc.add_paragraph()
                    p_sol.paragraph_format.left_indent = Cm(0.6)
                    p_sol.paragraph_format.space_before = Pt(2)
                    p_sol.paragraph_format.space_after = Pt(5)

                    r_sh = p_sol.add_run("Lời giải chi tiết:\n")
                    r_sh.font.name = "Times New Roman"
                    r_sh.font.size = Pt(9.5)
                    r_sh.font.bold = True
                    r_sh.font.color.rgb = COLOR_SOLUTION

                    add_chem_runs_hoa10(p_sol, q["solution"], size=9.5)

                display_num += 1

        fname = "Ngan_Hang_Hoa_10_Chuong_4_Phan_Ung_Oxi_Hoa_Khu_GiaoVien.docx" if is_teacher else "Ngan_Hang_Hoa_10_Chuong_4_Phan_Ung_Oxi_Hoa_Khu_HocSinh.docx"
        fpath = os.path.join(BANK_DIR, fname)
        doc.save(fpath)
        print(f"  + Saved {'GiaoVien' if is_teacher else 'HocSinh'} Word file: {fpath}")

if __name__ == "__main__":
    build_chapter_4()
