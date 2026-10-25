# -*- coding: utf-8 -*-
"""
Xây dựng Ngân hàng câu hỏi Hóa học 10 – Chương 6: Tốc độ phản ứng hóa học
Phân loại chi tiết theo:
- Chủ đề / Bài học:
  1. Khái niệm tốc độ phản ứng và tốc độ trung bình của phản ứng
  2. Định luật tác dụng khối lượng và hằng số tốc độ phản ứng
  3. Các yếu tố ảnh hưởng đến tốc độ phản ứng hóa học
  4. Ý nghĩa thực tiễn và ứng dụng của tốc độ phản ứng
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
CVL_PATH = r"c:\Antigravity_Thanh\San_Pham\c6_converted_num.docx"
TTB_PATH = r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\2026 - Phan dang de thi thu truong, so. TTB1-20260711T201135Z-2-001\2026 - Phan dang de thi thu truong, so. TTB1\X. PD 2026 - Hoa 10. CD10 - SGK va CDHT. TTB. ok - DA chi tiet.docx"

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
BANK_DIR = os.path.join(OUT_DIR, "Ngan_Hang_Hoa_10")
IMG_DIR = os.path.join(OUT_DIR, "images_hoa10")
os.makedirs(BANK_DIR, exist_ok=True)
os.makedirs(IMG_DIR, exist_ok=True)

TOPICS_LIST = [
    "Khái niệm tốc độ phản ứng và tốc độ trung bình của phản ứng",
    "Định luật tác dụng khối lượng và hằng số tốc độ phản ứng",
    "Các yếu tố ảnh hưởng đến tốc độ phản ứng hóa học",
    "Ý nghĩa thực tiễn và ứng dụng của tốc độ phản ứng"
]

def detect_topic(text):
    t_low = text.lower()
    if any(k in t_low for k in ["bảo quản thực phẩm", "than tổ ong", "nồi áp suất", "tủ lạnh", "sản xuất gang", "thực tiễn", "đời sống", "bếp than", "chữa cháy", "lên men"]):
        return "Ý nghĩa thực tiễn và ứng dụng của tốc độ phản ứng"
    elif any(k in t_low for k in ["hằng số tốc độ", "định luật tác dụng", "v = k", "biểu thức tốc độ", "bậc phản ứng", "v=k"]):
        return "Định luật tác dụng khối lượng và hằng số tốc độ phản ứng"
    elif any(k in t_low for k in ["nhiệt độ", "nồng độ", "áp suất", "diện tích", "xúc tác", "hạt", "bột", "thỏi", "tăng", "giảm", "ảnh hưởng", "hệ số nhiệt độ", "van't hoff", "vant hoff", "gamma"]):
        return "Các yếu tố ảnh hưởng đến tốc độ phản ứng hóa học"
    else:
        return "Khái niệm tốc độ phản ứng và tốc độ trung bình của phản ứng"

def parse_c6_master():
    """Parse master file c6_converted_num.docx."""
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
            for opt_k in ["A", "B", "C", "D"]:
                if opt_k in cur_q.get("options", {}):
                    txt = cur_q["options"][opt_k]
                    m_da = re.search(r'[\t\s]+(?:Đáp án|ĐA)[\:\.]?\s*([A-D])\b', txt, re.IGNORECASE)
                    if m_da:
                        if not cur_q.get("answer"):
                            cur_q["answer"] = m_da.group(1).upper()
                        cur_q["options"][opt_k] = re.sub(r'[\t\s]+(?:Đáp án|ĐA)[\:\.]?\s*[A-D]\.?', '', txt).strip()
            if not cur_q.get("answer") and cur_q.get("underlined_opt"):
                cur_q["answer"] = cur_q["underlined_opt"]
            if not cur_q.get("answer") and cur_q.get("solution_lines"):
                sol_txt = " ".join(cur_q["solution_lines"])
                m_sol_ans = re.search(r'(?:Đáp án đúng\s*[\:\.]?\s*|Đáp án\s*[\:\.]?\s*|Chọn\s+)([A-D])\b', sol_txt, re.IGNORECASE)
                if m_sol_ans:
                    cur_q["answer"] = m_sol_ans.group(1).upper()
            cur_q["level"] = "NB"
            if any(k in stem_all.lower() for k in ["ảnh hưởng", "tăng nhiệt độ", "chất xúc tác", "diện tích", "áp suất", "tốc độ trung bình"]):
                cur_q["level"] = "TH"
            if any(k in stem_all.lower() for k in ["tính", "tính tốc độ", "mol/(l.s)", "hằng số", "hệ số", "gamma", "gấp bao nhiêu", "sau bao lâu"]):
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

        elif sec in ["P3", "THEORY"]:
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
                        sol_all = " ".join(cur_q["solution_lines"])
                        m_ans_num = re.search(r'(?:kết quả|giá trị|đáp số|bằng|là)\s*[:=]?\s*([+-]?\d+(?:[\.,]\d+)?)', sol_all, re.IGNORECASE)
                        if m_ans_num:
                            cur_q["answer"] = m_ans_num.group(1).replace(',', '.')
                if not cur_q.get("answer"):
                    cur_q["answer"] = "Xem lời giải chi tiết"
                cur_q["level"] = "VD"

        questions.append(cur_q)
        cur_q = None

    for item in iter_block_items(doc):
        kind, el = item
        if kind == 'p':
            p = el
            full_text = get_para_text_full(p, doc.part).strip()
            if not full_text:
                continue

            if re.match(r'^PHẦN\s+4\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "P3"
                continue
            elif re.match(r'^PHẦN\s+3\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "P2"
                continue
            elif re.match(r'^PHẦN\s+2\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "P1"
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
                    "source": "CVL-C6",
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
                imgs = extract_real_images_from_el(p._element, doc.part)
                if imgs:
                    for im_b in imgs:
                        fn = f"h10_c6_master_q{len(questions)+1}_img{len(cur_q['images'])+1}.png"
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
                    fn = f"h10_c6_master_q{len(questions)+1}_img{len(cur_q['images'])+1}.png"
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

            elif sec == "P3":
                m_ans = re.match(r'^(?:Đáp\s*án|Trả\s*lời|KQ|Đáp\s*số)[\:\s]*(.*)', clean_t, re.IGNORECASE)
                if m_ans:
                    cur_q["raw_ans"] = m_ans.group(1).strip()
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
    """Parse solution section of CHUONG 6 HOA 10 DE X.docx."""
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
            for opt_k in ["A", "B", "C", "D"]:
                if opt_k in cur_q.get("options", {}):
                    txt = cur_q["options"][opt_k]
                    m_da = re.search(r'[\t\s]+(?:Đáp án|ĐA)[\:\.]?\s*([A-D])\b', txt, re.IGNORECASE)
                    if m_da:
                        if not cur_q.get("answer"):
                            cur_q["answer"] = m_da.group(1).upper()
                        cur_q["options"][opt_k] = re.sub(r'[\t\s]+(?:Đáp án|ĐA)[\:\.]?\s*[A-D]\.?', '', txt).strip()
            if not cur_q.get("answer") and cur_q.get("underlined_opt"):
                cur_q["answer"] = cur_q["underlined_opt"]
            if not cur_q.get("answer"):
                sol_txt = " ".join(cur_q["solution_lines"])
                m_sol_ans = re.search(r'(?:Đáp án đúng\s*[\:\.]?\s*|Đáp án\s*[\:\.]?\s*|Chọn\s+)([A-D])\b', sol_txt, re.IGNORECASE)
                if m_sol_ans:
                    cur_q["answer"] = m_sol_ans.group(1).upper()
            cur_q["level"] = "NB"
            if any(k in stem_all.lower() for k in ["ảnh hưởng", "tăng nhiệt độ", "chất xúc tác", "diện tích", "áp suất", "tốc độ trung bình"]):
                cur_q["level"] = "TH"
            if any(k in stem_all.lower() for k in ["tính", "tính tốc độ", "mol/(l.s)", "hằng số", "hệ số", "gamma", "gấp bao nhiêu", "sau bao lâu"]):
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
            if not cur_q.get("answer"):
                if cur_q.get("raw_ans"):
                    cur_q["answer"] = cur_q["raw_ans"]
                elif cur_q.get("solution_lines"):
                    sol_all = " ".join(cur_q["solution_lines"])
                    m_num = re.search(r'(?:kết quả|giá trị|đáp số|bằng|là)\s*[:=]?\s*([+-]?\d+(?:[\.,]\d+)?)', sol_all, re.IGNORECASE)
                    if m_num:
                        cur_q["answer"] = m_num.group(1).replace(',', '.')
            if not cur_q.get("answer"):
                cur_q["answer"] = "Xem lời giải chi tiết"
            cur_q["level"] = "VD"

        questions.append(cur_q)
        cur_q = None

    for i in range(start_idx, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        full_text = get_para_text_full(p, doc.part).strip()
        if not full_text:
            continue

        if re.match(r"^PHẦN\s+I[\.\:\s]", full_text, re.IGNORECASE):
            flush_q()
            cur_sec = "P1"
            continue
        elif re.match(r"^PHẦN\s+II[\.\:\s]", full_text, re.IGNORECASE):
            flush_q()
            cur_sec = "P2"
            continue
        elif re.match(r"^PHẦN\s+III[\.\:\s]", full_text, re.IGNORECASE):
            flush_q()
            cur_sec = "P3"
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
                "sol_lines": [],
                "images": [],
                "data_tables": []
            }
            imgs = extract_real_images_from_el(p._element, doc.part)
            if imgs:
                for im_b in imgs:
                    fn = f"h10_c6_{filename[:5].lower()}_q{len(questions)+1}_img{len(cur_q['images'])+1}.png"
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
                fn = f"h10_c6_{filename[:5].lower()}_q{len(questions)+1}_img{len(cur_q['images'])+1}.png"
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
            m_ans = re.match(r'^(?:Đáp\s*án|Trả\s*lời|KQ|Đáp\s*số)[\:\s]*(.*)', clean_t, re.IGNORECASE)
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

def parse_ttb_hoa10_c6():
    """Parse questions 134 to 144 from TTB Grade 10."""
    if not os.path.exists(TTB_PATH):
        return []
    doc = docx.Document(TTB_PATH)
    ttb_qs = []
    cur_sec = "P1"
    cur_level = "NB"
    cur_q = None

    def flush_q():
        nonlocal cur_q
        if not cur_q:
            return
        if cur_q["opt_lines"]:
            cur_q["options"] = parse_abcd_from_opt_lines(cur_q["opt_lines"])
        for opt_k in ["A", "B", "C", "D"]:
            if opt_k in cur_q.get("options", {}):
                txt = cur_q["options"][opt_k]
                m_da = re.search(r'[\t\s]+(?:Đáp án|ĐA)[\:\.]?\s*([A-D])\b', txt, re.IGNORECASE)
                if m_da:
                    if not cur_q.get("answer"):
                        cur_q["answer"] = m_da.group(1).upper()
                    cur_q["options"][opt_k] = re.sub(r'[\t\s]+(?:Đáp án|ĐA)[\:\.]?\s*[A-D]\.?', '', txt).strip()
        if not cur_q.get("answer") and cur_q.get("underlined_opt"):
            cur_q["answer"] = cur_q["underlined_opt"]
        if not cur_q.get("answer") and cur_q.get("solution_lines"):
            sol_all = " ".join(cur_q["solution_lines"])
            m_sol = re.search(r'(?:Đáp án đúng\s*[\:\.]?\s*|Đáp án\s*[\:\.]?\s*|Chọn\s+)([A-D])\b', sol_all, re.IGNORECASE)
            if m_sol:
                cur_q["answer"] = m_sol.group(1).upper()
            else:
                for opt_k in ["A", "B", "C", "D"]:
                    opt_val = cur_q.get("options", {}).get(opt_k, "").strip()
                    m_val = re.search(r'([+-]?\d+(?:[\.,]\d+)?)', opt_val)
                    if m_val and m_val.group(1) in sol_all:
                        cur_q["answer"] = opt_k
                        break
        cur_q["topic"] = detect_topic(" ".join(cur_q["stem_lines"]))
        ttb_qs.append(cur_q)
        cur_q = None

    in_c6 = False
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if 'VII. TỐC ĐỘ PHẢN ỨNG' in t:
            in_c6 = True
            continue
        elif in_c6 and 'VIII. CƠ SỞ HÓA HỌC' in t:
            break
        elif not in_c6:
            continue

        if '♦ MỨC ĐỘ BIẾT' in t:
            cur_level = "NB"
            continue
        elif '♦ MỨC ĐỘ HIỂU' in t:
            cur_level = "TH"
            continue
        elif '♦ MỨC ĐỘ VẬN DỤNG' in t:
            cur_level = "VD"
            continue

        m_cau = re.match(r'^Câu\s+(\d+)[\.\:]?\s*(.*)', t, re.IGNORECASE)
        if m_cau:
            flush_q()
            num = int(m_cau.group(1))
            cur_q = {
                "source": "TTB",
                "raw_num": num,
                "sec": "P1",
                "level": cur_level,
                "stem_lines": [m_cau.group(2).strip()] if m_cau.group(2).strip() else [],
                "opt_lines": [],
                "options": {},
                "items": [],
                "sol_map": {},
                "answer": None,
                "solution_lines": [],
                "images": [],
                "data_tables": []
            }
            imgs = extract_real_images_from_el(p._element, doc.part)
            if imgs:
                for im_b in imgs:
                    fn = f"h10_c6_ttb_q{num}_img{len(cur_q['images'])+1}.png"
                    fpath = os.path.join(IMG_DIR, fn)
                    with open(fpath, "wb") as f_img:
                        f_img.write(im_b)
                    cur_q["images"].append(fpath)
            continue

        if not cur_q:
            continue

        clean_t = re.sub(r'^[📝\u2022\uf0b7\s\t]+', '', t).strip()
        u_opt = extract_underlined_p1_answer([p])
        if u_opt:
            cur_q["underlined_opt"] = u_opt

        if re.match(r'^[A-D]\.', t) or "\tB." in t or "\tC." in t:
            cur_q["opt_lines"].append(t)
        elif re.match(r'^(?:Hướng dẫn giải|Giải thích|Lời giải)[\:\s]*', clean_t, re.IGNORECASE):
            cur_q["solution_lines"].append(clean_t)
        elif cur_q["solution_lines"]:
            cur_q["solution_lines"].append(clean_t)
        elif not cur_q["opt_lines"]:
            cur_q["stem_lines"].append(t)
        else:
            cur_q["opt_lines"].append(t)

    flush_q()
    return ttb_qs

def build_chapter_6_bank():
    print("Parsing Master CVL Chapter 6...")
    master_qs = parse_c6_master()
    print(f"Master file parsed: {len(master_qs)} questions")

    print("Parsing DE 1, 2, 3 Chapter 6...")
    de1_qs = parse_de_solutions("CHUONG 6 HOA 10 - DE 1.docx")
    de2_qs = parse_de_solutions("CHUONG 6 HOA 10- DE 2.docx")
    de3_qs = parse_de_solutions("CHUONG 6 HOA 10- DE 3.docx")
    print(f"DE 1: {len(de1_qs)} | DE 2: {len(de2_qs)} | DE 3: {len(de3_qs)} questions")

    print("Parsing TTB Chapter 6...")
    ttb_qs = parse_ttb_hoa10_c6()
    print(f"TTB: {len(ttb_qs)} questions")

    all_raw = master_qs + de1_qs + de2_qs + de3_qs + ttb_qs
    print(f"Total raw candidate questions: {len(all_raw)}")

    unique_questions = []
    seen_stems = []

    for q in all_raw:
        stem_txt = " ".join(q["stem_lines"]).strip()
        if not stem_txt:
            continue
        norm_s = norm_for_compare(stem_txt)
        is_dup = False
        for seen in seen_stems:
            sim = SequenceMatcher(None, norm_s, seen).ratio()
            if sim > 0.82:
                is_dup = True
                break
        if not is_dup:
            seen_stems.append(norm_s)
            unique_questions.append(q)

    print(f"Total unique questions after dedup: {len(unique_questions)}")

    p1_list = [q for q in unique_questions if q["sec"] == "P1"]
    p2_list = [q for q in unique_questions if q["sec"] == "P2"]
    p3_list = [q for q in unique_questions if q["sec"] == "P3"]

    def sort_key(q):
        lvl_order = {"NB": 1, "TH": 2, "VD": 3}
        return (lvl_order.get(q.get("level", "NB"), 1), q.get("topic", ""))

    p1_list.sort(key=sort_key)
    p2_list.sort(key=sort_key)
    p3_list.sort(key=sort_key)

    def finalize_q(q, section, idx):
        lvl = q.get("level", "TH")
        qid = f"H10-C6-{section}-{lvl}-{idx:03d}"
        stem = clean_extracted_text(" ".join(q["stem_lines"]))
        sol = "\n".join(q["solution_lines"]).strip()
        if not sol:
            sol = "Xem nội dung và phân tích chi tiết trong tài liệu tham khảo."

        final_dict = {
            "id": qid,
            "grade": 10,
            "chapter_number": 6,
            "chapter_name": "Tốc độ phản ứng hóa học",
            "topic": q.get("topic", "Khái niệm tốc độ phản ứng và tốc độ trung bình của phản ứng"),
            "section": section,
            "level": lvl,
            "stem": stem,
            "stem_lines": [clean_extracted_text(l) for l in q["stem_lines"] if l.strip()],
            "images": q.get("images", []),
            "data_tables": q.get("data_tables", [])
        }

        if section == "P1":
            opts = q.get("options", {})
            for k in ["A", "B", "C", "D"]:
                if k not in opts:
                    opts[k] = ""
                else:
                    opts[k] = clean_extracted_text(opts[k])
            final_dict["options"] = opts
            ans = q.get("answer")
            if not ans:
                for opt_k in ["A", "B", "C", "D"]:
                    opt_val = opts.get(opt_k, "")
                    m_val = re.search(r'([+-]?\d+(?:[\.,]\d+)?)', opt_val)
                    if m_val and m_val.group(1) in sol:
                        ans = opt_k
                        break
            final_dict["answer"] = ans or "A"
            final_dict["solution"] = clean_extracted_text(sol)

        elif section == "P2":
            items = []
            for it in q.get("items", []):
                items.append({
                    "label": it.get("label", "a"),
                    "text": clean_extracted_text(it.get("text", "")),
                    "answer": it.get("answer", "Đúng"),
                    "solution": clean_extracted_text(it.get("solution", ""))
                })
            final_dict["items"] = items
            final_dict["answer"] = q.get("answer", "a) Đúng, b) Đúng, c) Sai, d) Sai")
            final_dict["solution"] = clean_extracted_text(sol)

        elif section == "P3":
            final_dict["answer"] = clean_extracted_text(str(q.get("answer", "Xem lời giải chi tiết")))
            final_dict["solution"] = clean_extracted_text(sol)

        return final_dict

    final_p1 = [finalize_q(q, "P1", i+1) for i, q in enumerate(p1_list)]
    final_p2 = [finalize_q(q, "P2", i+1) for i, q in enumerate(p2_list)]
    final_p3 = [finalize_q(q, "P3", i+1) for i, q in enumerate(p3_list)]

    all_final_qs = final_p1 + final_p2 + final_p3

    c6_db = {
        "C6": {
            "grade": 10,
            "chapter_number": 6,
            "chapter_name": "Tốc độ phản ứng hóa học",
            "topics": TOPICS_LIST,
            "total_questions": len(all_final_qs),
            "questions": all_final_qs
        }
    }

    json_path = os.path.join(BANK_DIR, "Ngan_Hang_Hoa_10_Chuong_6_Toc_Do_Phan_Ung.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(c6_db, f, ensure_ascii=False, indent=2)
    print(f"Saved JSON Question Bank ({len(all_final_qs)} questions): {json_path}")

    # Generate Word files
    def generate_docx(is_teacher=False):
        doc = docx.Document()
        setup_page_and_hf(doc)

        # Header Title Block
        t_table = doc.add_table(rows=1, cols=2)
        t_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_table.autofit = False
        t_table.columns[0].width = Cm(8.0)
        t_table.columns[1].width = Cm(10.0)

        c0 = t_table.cell(0, 0).paragraphs[0]
        c0.paragraph_format.space_before = Pt(0)
        c0.paragraph_format.space_after = Pt(2)
        r01 = c0.add_run("BỘ GIÁO DỤC VÀ ĐÀO TẠO\n")
        r01.font.name = "Times New Roman"
        r01.font.size = Pt(10)
        r01.font.bold = True
        r02 = c0.add_run("NGÂN HÀNG CÂU HỎI HÓA HỌC 10\n(CHƯƠNG TRÌNH GDPT 2018)")
        r02.font.name = "Times New Roman"
        r02.font.size = Pt(9.5)
        r02.font.italic = True

        c1 = t_table.cell(0, 1).paragraphs[0]
        c1.paragraph_format.space_before = Pt(0)
        c1.paragraph_format.space_after = Pt(2)
        c1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r11 = c1.add_run("CHƯƠNG 6: TỐC ĐỘ PHẢN ỨNG HÓA HỌC\n")
        r11.font.name = "Times New Roman"
        r11.font.size = Pt(11)
        r11.font.bold = True
        r11.font.color.rgb = COLOR_PRIMARY
        r12 = c1.add_run(f"Số lượng: {len(all_final_qs)} câu hỏi (Phần I, II, III)\n")
        r12.font.name = "Times New Roman"
        r12.font.size = Pt(9.5)
        r12.font.italic = True
        if is_teacher:
            r13 = c1.add_run("(BẢN GIÁO VIÊN — ĐÁP ÁN & LỜI GIẢI CHI TIẾT)")
            r13.font.name = "Times New Roman"
            r13.font.size = Pt(9.5)
            r13.font.bold = True
            r13.font.color.rgb = COLOR_HF_RED
        else:
            r13 = c1.add_run("(BẢN HỌC SINH — DÙNG CHO ÔN TẬP VÀ LUYỆN ĐỀ)")
            r13.font.name = "Times New Roman"
            r13.font.size = Pt(9.5)
            r13.font.italic = True

        p_div = doc.add_paragraph()
        p_div.paragraph_format.space_before = Pt(4)
        p_div.paragraph_format.space_after = Pt(8)
        p_div_border = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                                 r'<w:bottom w:val="single" w:sz="6" w:space="1" w:color="8BC390"/>'
                                 r'</w:pBdr>')
        p_div._element.get_or_add_pPr().append(p_div_border)

        sections_def = [
            ("PHẦN I. CÂU HỎI TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN",
             "Thí sinh trả lời từ câu 1 đến câu hết phần này. Mỗi câu hỏi chỉ chọn một phương án chính xác nhất.",
             final_p1, "P1"),
            ("PHẦN II. CÂU HỎI TRẮC NGHIỆM ĐÚNG / SAI",
             "Thí sinh trả lời từ câu 1 đến câu hết phần này. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.",
             final_p2, "P2"),
            ("PHẦN III. CÂU HỎI TRẮC NGHIỆM YÊU CẦU TRẢ LỜI NGẮN",
             "Thí sinh trả lời từ câu 1 đến câu hết phần này. Điền kết quả hoặc đáp số tương ứng vào chỗ trống.",
             final_p3, "P3")
        ]

        for sec_title, sec_desc, q_list, sec_type in sections_def:
            if not q_list:
                continue
            h_p = doc.add_paragraph()
            h_p.paragraph_format.space_before = Pt(12)
            h_p.paragraph_format.space_after = Pt(2)
            r_h = h_p.add_run(sec_title)
            r_h.font.name = "Times New Roman"
            r_h.font.size = Pt(11)
            r_h.font.bold = True
            r_h.font.color.rgb = COLOR_PRIMARY

            d_p = doc.add_paragraph()
            d_p.paragraph_format.space_before = Pt(0)
            d_p.paragraph_format.space_after = Pt(6)
            r_d = d_p.add_run(sec_desc)
            r_d.font.name = "Times New Roman"
            r_d.font.size = Pt(10)
            r_d.font.italic = True

            for idx, q in enumerate(q_list):
                q_p = doc.add_paragraph()
                q_p.paragraph_format.space_before = Pt(4)
                q_p.paragraph_format.space_after = Pt(2)

                r_lbl = q_p.add_run(f"Câu {idx+1}. ")
                r_lbl.font.name = "Times New Roman"
                r_lbl.font.size = Pt(10.5)
                r_lbl.font.bold = True
                r_lbl.font.color.rgb = COLOR_PRIMARY

                lvl_txt = f"[{q.get('level', 'TH')}] "
                r_lvl = q_p.add_run(lvl_txt)
                r_lvl.font.name = "Times New Roman"
                r_lvl.font.size = Pt(10)
                r_lvl.font.bold = True
                r_lvl.font.color.rgb = COLOR_TAG

                add_chem_runs_hoa10(q_p, q.get("stem", ""), size=Pt(10.5))

                for img_p in q.get("images", []):
                    if os.path.exists(img_p):
                        try:
                            doc.add_picture(img_p, width=Cm(10.0))
                            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        except Exception:
                            pass

                for t_data in q.get("data_tables", []):
                    add_data_table_to_doc(doc, t_data)

                if sec_type == "P1":
                    corr = q.get("answer", "") if is_teacher else ""
                    add_options_paragraph_no_table(doc, q.get("options", {}), corr, is_teacher)

                elif sec_type == "P2":
                    for it in q.get("items", []):
                        it_p = doc.add_paragraph()
                        it_p.paragraph_format.left_indent = Cm(0.6)
                        it_p.paragraph_format.space_before = Pt(1)
                        it_p.paragraph_format.space_after = Pt(1)
                        r_it_lbl = it_p.add_run(f"{it.get('label', 'a')}) ")
                        r_it_lbl.font.name = "Times New Roman"
                        r_it_lbl.font.size = Pt(10.5)
                        r_it_lbl.font.bold = True
                        add_chem_runs_hoa10(it_p, it.get("text", ""), size=Pt(10.5))

                        if is_teacher and it.get("answer"):
                            r_ans = it_p.add_run(f"  --> {it['answer']}")
                            r_ans.font.name = "Times New Roman"
                            r_ans.font.size = Pt(10)
                            r_ans.font.bold = True
                            r_ans.font.color.rgb = COLOR_CORRECT if it['answer'] == 'Đúng' else COLOR_HF_RED

                elif sec_type == "P3":
                    ans_box = doc.add_paragraph()
                    ans_box.paragraph_format.left_indent = Cm(0.6)
                    ans_box.paragraph_format.space_before = Pt(2)
                    ans_box.paragraph_format.space_after = Pt(2)
                    r_ans_lbl = ans_box.add_run("Trả lời: ")
                    r_ans_lbl.font.name = "Times New Roman"
                    r_ans_lbl.font.size = Pt(10)
                    r_ans_lbl.font.bold = True
                    if is_teacher:
                        r_ans_val = ans_box.add_run(f" {q.get('answer', '')}")
                        r_ans_val.font.name = "Times New Roman"
                        r_ans_val.font.size = Pt(10)
                        r_ans_val.font.bold = True
                        r_ans_val.font.color.rgb = COLOR_CORRECT
                    else:
                        r_dots = ans_box.add_run("....................................................................................")
                        r_dots.font.name = "Times New Roman"
                        r_dots.font.size = Pt(10)

                if is_teacher and q.get("solution"):
                    sol_p = doc.add_paragraph()
                    sol_p.paragraph_format.left_indent = Cm(0.6)
                    sol_p.paragraph_format.space_before = Pt(2)
                    sol_p.paragraph_format.space_after = Pt(4)
                    r_s_tag = sol_p.add_run("Hướng dẫn giải: ")
                    r_s_tag.font.name = "Times New Roman"
                    r_s_tag.font.size = Pt(10)
                    r_s_tag.font.bold = True
                    r_s_tag.font.color.rgb = COLOR_SOLUTION
                    add_chem_runs_hoa10(sol_p, q.get("solution", ""), size=Pt(9.5))

        out_fname = f"Ngan_Hang_Hoa_10_Chuong_6_Toc_Do_Phan_Ung_{'GiaoVien' if is_teacher else 'HocSinh'}.docx"
        out_p = os.path.join(BANK_DIR, out_fname)
        doc.save(out_p)
        print(f"  + Saved {'GiaoVien' if is_teacher else 'HocSinh'} Word file: {out_p}")

    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)

    # Summary
    nb_cnt = sum(1 for q in all_final_qs if q.get("level") == "NB")
    th_cnt = sum(1 for q in all_final_qs if q.get("level") == "TH")
    vd_cnt = sum(1 for q in all_final_qs if q.get("level") == "VD")

    print("\n=== THỐNG KÊ CHI TIẾT NGÂN HÀNG HÓA 10 - CHƯƠNG 6 ===")
    print(f"  + Tổng số câu hỏi: {len(all_final_qs)}")
    print(f"  + Phần I (P1 - 4 lựa chọn): {len(final_p1)} câu")
    print(f"  + Phần II (P2 - Đúng/Sai 4 ý): {len(final_p2)} câu")
    print(f"  + Phần III (P3 - Trả lời ngắn & Bài toán): {len(final_p3)} câu")
    print(f"  + Nhận biết (NB): {nb_cnt} | Thông hiểu (TH): {th_cnt} | Vận dụng (VD): {vd_cnt}")

if __name__ == "__main__":
    build_chapter_6_bank()
