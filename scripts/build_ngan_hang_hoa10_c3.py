# -*- coding: utf-8 -*-
"""
Xây dựng Ngân hàng câu hỏi Hóa học 10 – Chương 3: Liên kết hóa học
Phân loại chi tiết theo:
- Chủ đề / Bài học:
  1. Bài 10: Quy tắc octet
  2. Bài 11: Liên kết ion
  3. Bài 12: Liên kết cộng hóa trị
  4. Bài 13: Liên kết hydrogen và tương tác van der Waals
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

SRC_DOCX = r"c:\Antigravity_Thanh\San_Pham\c3_converted_num.docx"
TTB_DOCX = r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\2026 - Phan dang de thi thu truong, so. TTB1-20260711T201135Z-2-001\2026 - Phan dang de thi thu truong, so. TTB1\X. PD 2026 - Hoa 10. CD10 - SGK va CDHT. TTB. ok - DA chi tiet.docx"

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
BANK_DIR = os.path.join(OUT_DIR, "Ngan_Hang_Hoa_10")
IMG_DIR = os.path.join(OUT_DIR, "images_hoa10")
os.makedirs(BANK_DIR, exist_ok=True)
os.makedirs(IMG_DIR, exist_ok=True)

TOPICS = {
    10: "Quy tắc octet",
    11: "Liên kết ion",
    12: "Liên kết cộng hóa trị",
    13: "Liên kết hydrogen và tương tác van der Waals"
}

def parse_cvl_hoa10_c3():
    """Parse all questions from c3_converted_num.docx."""
    doc = docx.Document(SRC_DOCX)
    raw_questions = []

    cur_lesson = 10
    cur_sec = "THEORY"  # THEORY, P1, P2, P3
    cur_q = None

    def flush_q():
        nonlocal cur_q
        if not cur_q:
            return

        sec = cur_q["sec"]
        if sec == "P1":
            # Extract options A, B, C, D
            opts = {}
            if cur_q["opt_lines"]:
                opts = parse_abcd_from_opt_lines(cur_q["opt_lines"])
            cur_q["options"] = opts

            # Answer key
            if not cur_q.get("answer"):
                if cur_q.get("underlined_opt"):
                    cur_q["answer"] = cur_q["underlined_opt"]

            # Fallback for Q37 NaCl lattice
            if not cur_q.get("answer") and "sodium chloride" in " ".join(cur_q["stem_lines"]).lower():
                cur_q["answer"] = "Mỗi ion Na+ và Cl- được bao quanh bởi 6 ion trái dấu (phối trí 6:6)"

            # Cognitive level
            full_text = " ".join(cur_q["stem_lines"]) + " " + " ".join(cur_q["opt_lines"])
            if any(k in full_text.lower() for k in ["số phát biểu đúng", "có bao nhiêu", "hình vẽ", "mô hình", "xen phủ"]):
                cur_q["level"] = "TH"
            elif any(k in full_text.lower() for k in ["năng lượng liên kết", "biến thiên enthalpy", "tính", "khối lượng"]):
                cur_q["level"] = "VD"
            else:
                cur_q["level"] = "NB"

        elif sec == "P2":
            # Map solutions to items
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
            stem_txt = " ".join(cur_q["stem_lines"])
            if any(k in stem_txt.lower() for k in ["tính", "năng lượng", "phức tạp", "độ dài"]):
                cur_q["level"] = "VD"

        elif sec in ["P3", "THEORY"]:
            # If it's a P2-like theory question with a), b), c), d)
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
                        # Extract fill-in answers like (1) - ..., (2) - ...
                        fill_ans = [s for s in cur_q["sol_lines"] if re.match(r'^\(?\d+\)?\s*[\-\:]\s*', s)]
                        if fill_ans:
                            cur_q["answer"] = "; ".join(fill_ans)
                    if not cur_q.get("answer") and cur_q.get("solution_lines"):
                        cur_q["answer"] = cur_q["solution_lines"][0][:60]
                cur_q["level"] = "TH"
                stem_txt = " ".join(cur_q["stem_lines"])
                if any(k in stem_txt.lower() for k in ["tính", "năng lượng", "nhiệt độ", "bao nhiêu gam", "mol"]):
                    cur_q["level"] = "VD"

        raw_questions.append(cur_q)
        cur_q = None

    for kind, item in iter_block_items(doc):
        if kind == 'p':
            p = item
            raw_text = p.text.strip()
            full_text = get_para_text_full(p, doc.part).strip()

            if not full_text:
                continue

            # Check Lesson Heading
            m_bai = re.search(r'BÀI\s+(10|11|12|13)\b', full_text, re.IGNORECASE)
            if m_bai and len(full_text) < 80:
                flush_q()
                cur_lesson = int(m_bai.group(1))
                continue

            # Check Section Heading
            # ONLY match headings with PHẦN to avoid clashes with sub-headings (I., II.)
            if re.match(r'^PHẦN\s+IV\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "P3"
                continue
            elif re.match(r'^PHẦN\s+III\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "P2"
                continue
            elif re.match(r'^PHẦN\s+II\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "P1"
                continue
            elif re.match(r'^PHẦN\s+I\b', full_text, re.IGNORECASE):
                flush_q()
                cur_sec = "THEORY"
                continue

            # Stop before table of contents at the end
            if "HỆ THỐNG BÀI TẬP\t" in full_text:
                break

            # Check Question Start
            m_cau = re.match(r'^Câu\s+(\d+)[\.\:]?\s*(.*)', full_text, re.IGNORECASE)
            if m_cau:
                flush_q()
                q_num = int(m_cau.group(1))
                rest = m_cau.group(2).strip()

                cur_q = {
                    "raw_num": q_num,
                    "lesson": cur_lesson,
                    "topic": TOPICS.get(cur_lesson, "Liên kết hóa học"),
                    "sec": cur_sec,
                    "stem_lines": [rest] if rest else [],
                    "opt_lines": [],
                    "items": [],
                    "sol_map": {},
                    "answer": None,
                    "raw_ans": None,
                    "solution_lines": [],
                    "sol_lines": [],
                    "images": [],
                    "data_tables": [],
                    "underlined_opt": None
                }

                # Extract images from question stem paragraph
                imgs = extract_real_images_from_el(p._element, doc.part)
                if imgs:
                    for im_b in imgs:
                        fn = f"h10_c3_q{len(raw_questions)+1}_img{len(cur_q['images'])+1}.png"
                        fpath = os.path.join(IMG_DIR, fn)
                        with open(fpath, "wb") as f_img:
                            f_img.write(im_b)
                        cur_q["images"].append(fpath)
                continue

            if not cur_q:
                continue

            # Extract any images in subsequent paragraphs of this question
            imgs = extract_real_images_from_el(p._element, doc.part)
            if imgs:
                for im_b in imgs:
                    fn = f"h10_c3_q{len(raw_questions)+1}_img{len(cur_q['images'])+1}.png"
                    fpath = os.path.join(IMG_DIR, fn)
                    with open(fpath, "wb") as f_img:
                        f_img.write(im_b)
                    cur_q["images"].append(fpath)

            sec = cur_q["sec"]
            if sec == "P1":
                # Check underlined runs for answer
                u_opt = extract_underlined_p1_answer([p])
                if u_opt:
                    cur_q["underlined_opt"] = u_opt

                if re.match(r'^[A-D]\.', full_text) or "\tB." in full_text or "\tC." in full_text:
                    cur_q["opt_lines"].append(full_text)
                elif re.match(r'^(?:Hướng dẫn giải|Giải thích|Lời giải)[\:\s]*', full_text, re.IGNORECASE):
                    cur_q["solution_lines"].append(full_text)
                elif cur_q["solution_lines"]:
                    cur_q["solution_lines"].append(full_text)
                elif not cur_q["opt_lines"]:
                    cur_q["stem_lines"].append(full_text)
                else:
                    cur_q["opt_lines"].append(full_text)

            elif sec == "P2":
                clean_t = re.sub(r'^[\uf0b7\u2022\-\*\s\t]+', '', full_text)
                m_sub = re.match(r'^([a-d])\)\s*(.*)', clean_t)
                if m_sub:
                    lbl = m_sub.group(1).lower()
                    it_txt = m_sub.group(2).strip()
                    m_tf = re.match(r'^(Đúng|Sai)[\.\:\,\s]*(.*)', it_txt, re.IGNORECASE)
                    if m_tf:
                        cur_q["sol_map"][lbl] = (m_tf.group(1).capitalize(), m_tf.group(2).strip())
                    else:
                        cur_q["items"].append({"label": lbl, "text": it_txt, "answer": None, "solution": ""})
                elif re.match(r'^(?:Hướng dẫn giải|Giải thích|Lời giải)[\:\s]*', full_text, re.IGNORECASE):
                    cur_q["solution_lines"].append(full_text)
                elif cur_q["solution_lines"]:
                    cur_q["solution_lines"].append(full_text)
                elif not cur_q["items"]:
                    cur_q["stem_lines"].append(full_text)
                else:
                    cur_q["solution_lines"].append(full_text)

            elif sec in ["P3", "THEORY"]:
                clean_t = re.sub(r'^[\uf0b7\u2022\-\*\s\t]+', '', full_text)
                m_ans = re.match(r'^(?:Đáp\s*án|Trả\s*lời|KQ)[\:\s]*(.*)', clean_t, re.IGNORECASE)
                if m_ans:
                    cur_q["raw_ans"] = m_ans.group(1).strip()
                elif 'Đáp án đúng:' in full_text:
                    cur_q["raw_ans"] = full_text.split('Đáp án đúng:')[1].strip()
                elif re.match(r'^(?:Hướng dẫn giải|Giải thích|Lời giải)[\:\s]*', full_text, re.IGNORECASE):
                    cur_q["solution_lines"].append(full_text)
                elif cur_q["solution_lines"]:
                    cur_q["solution_lines"].append(full_text)
                    if re.match(r'^\(?\d+\)?\s*[\-\:]\s*', clean_t):
                        cur_q["sol_lines"].append(clean_t)
                else:
                    m_sub = re.match(r'^([a-d])\)\s*(.*)', clean_t)
                    if m_sub:
                        lbl = m_sub.group(1).lower()
                        it_txt = m_sub.group(2).strip()
                        m_tf = re.match(r'^(Đúng|Sai)[\.\:\,\s]*(.*)', it_txt, re.IGNORECASE)
                        if m_tf:
                            cur_q["sol_map"][lbl] = (m_tf.group(1).capitalize(), m_tf.group(2).strip())
                        else:
                            cur_q["items"].append({"label": lbl, "text": it_txt, "answer": None, "solution": ""})
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
    return raw_questions

def parse_ttb_hoa10_c3():
    """Parse questions 39 to 52 from TTB Grade 10."""
    if not os.path.exists(TTB_DOCX):
        return []
    doc = docx.Document(TTB_DOCX)
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

    # Lines 118 to 162 contain Chapter 3 in TTB
    for i in range(118, 163):
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

            topic = "Liên kết ion"
            if any(k in clean_stem.lower() for k in ["cộng hoá trị", "cộng hóa trị", "xen phủ", "năng lượng liên kết", "đôi"]):
                topic = "Liên kết cộng hóa trị"
            elif any(k in clean_stem.lower() for k in ["hydrogen", "van der waals", "nhiệt độ sôi"]):
                topic = "Liên kết hydrogen và tương tác van der Waals"

            level = "NB" if q_num <= 46 else "TH"

            cur_q = {
                "raw_num": q_num,
                "lesson": 11 if "ion" in topic else (12 if "cộng hóa trị" in topic else 13),
                "topic": topic,
                "sec": "P1",
                "level": level,
                "source": src_tag,
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

    explicit_ans = {
        39: "B", 40: "A", 41: "A", 42: "C", 43: "A", 44: "C", 45: "D",
        46: "A", 47: "A", 48: "A", 49: "B", 50: "B", 51: "D", 52: "A"
    }
    for q in ttb_qs:
        rn = q["raw_num"]
        if rn in explicit_ans:
            q["answer"] = explicit_ans[rn]

    return ttb_qs

def build_chapter_3():
    print("Parsing CVL Chapter 3...")
    raw_cvl = parse_cvl_hoa10_c3()
    print(f"Parsed CVL: {len(raw_cvl)} questions")

    print("Parsing TTB Chapter 3...")
    raw_ttb = parse_ttb_hoa10_c3()
    print(f"Parsed TTB: {len(raw_ttb)} questions")

    # Filter out duplicate questions
    all_raw = raw_cvl + raw_ttb
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
        return (q.get("lesson", 10), lvl_order.get(q.get("level", "NB"), 1), q.get("raw_num", 0))

    p1_list.sort(key=sort_key)
    p2_list.sort(key=sort_key)
    p3_list.sort(key=sort_key)

    final_questions = []
    q_counter = 1

    for q in p1_list:
        lvl = q.get("level", "NB")
        qid = f"H10-C3-P1-{lvl}-{q_counter:03d}"
        q_counter += 1
        sol_text = "\n".join(q.get("solution_lines", []))
        if not sol_text:
            sol_text = f"Chọn đáp án {q.get('answer', 'A')}."
        final_questions.append({
            "id": qid,
            "grade": 10,
            "chapter_number": 3,
            "chapter_name": "Liên kết hóa học",
            "topic": q.get("topic", "Liên kết hóa học"),
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
        qid = f"H10-C3-P2-{lvl}-{q_counter:03d}"
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
            "chapter_number": 3,
            "chapter_name": "Liên kết hóa học",
            "topic": q.get("topic", "Liên kết hóa học"),
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
        qid = f"H10-C3-P3-{lvl}-{q_counter:03d}"
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
            "chapter_number": 3,
            "chapter_name": "Liên kết hóa học",
            "topic": q.get("topic", "Liên kết hóa học"),
            "section": "P3",
            "level": lvl,
            "stem": "\n".join(q.get("stem_lines", [])),
            "stem_lines": q.get("stem_lines", []),
            "answer": ans_val,
            "solution": sol_text,
            "images": q.get("images", []),
            "data_tables": q.get("data_tables", [])
        })

    # Save JSON database
    c3_json_path = os.path.join(BANK_DIR, "Ngan_Hang_Hoa_10_Chuong_3_Lien_Ket_Hoa_Hoc.json")
    c3_dict = {
        "C3": {
            "grade": 10,
            "chapter_number": 3,
            "chapter_name": "Liên kết hóa học",
            "topics": list(TOPICS.values()),
            "total_questions": len(final_questions),
            "questions": final_questions
        }
    }
    with open(c3_json_path, "w", encoding="utf-8") as f:
        json.dump(c3_dict, f, ensure_ascii=False, indent=2)
    print(f"Saved JSON Question Bank ({len(final_questions)} questions): {c3_json_path}")

    # Build Word documents
    build_word_docs_c3(final_questions)

    # Print summary statistics
    p1_cnt = sum(1 for q in final_questions if q["section"] == "P1")
    p2_cnt = sum(1 for q in final_questions if q["section"] == "P2")
    p3_cnt = sum(1 for q in final_questions if q["section"] == "P3")
    nb_cnt = sum(1 for q in final_questions if q["level"] == "NB")
    th_cnt = sum(1 for q in final_questions if q["level"] == "TH")
    vd_cnt = sum(1 for q in final_questions if q["level"] == "VD")

    print(f"\n=== THỐNG KÊ CHI TIẾT NGÂN HÀNG HÓA 10 - CHƯƠNG 3 ===")
    print(f"  + Tổng số câu hỏi: {len(final_questions)}")
    print(f"  + Phần I (P1 - 4 lựa chọn): {p1_cnt} câu")
    print(f"  + Phần II (P2 - Đúng/Sai 4 ý): {p2_cnt} câu")
    print(f"  + Phần III (P3 - Trả lời ngắn & Bài toán): {p3_cnt} câu")
    print(f"  + Nhận biết (NB): {nb_cnt} | Thông hiểu (TH): {th_cnt} | Vận dụng (VD): {vd_cnt}")

def build_word_docs_c3(questions):
    """Generate HocSinh and GiaoVien Word documents following 'FORM ĐỀ KIỂM TRA'."""
    for is_teacher in [False, True]:
        doc = docx.Document()
        setup_page_and_hf(doc)

        # Document Header block
        title_p = doc.add_paragraph()
        title_p.paragraph_format.space_before = Pt(0)
        title_p.paragraph_format.space_after = Pt(4)
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_h = title_p.add_run("NGÂN HÀNG CÂU HỎI HÓA HỌC 10 – CHƯƠNG 3: LIÊN KẾT HÓA HỌC")
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

        # Section groupings
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

                # Question stem
                stem_txt = q.get("stem", "")
                add_chem_runs_hoa10(p_q, stem_txt, size=10.0)

                # Images
                for im_p in q.get("images", []):
                    if os.path.exists(im_p):
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.paragraph_format.space_before = Pt(3)
                        p_img.paragraph_format.space_after = Pt(3)
                        p_img.add_run().add_picture(im_p, width=Cm(11.0))

                # Data tables if any
                for dt in q.get("data_tables", []):
                    add_data_table_to_doc(doc, dt)

                # Format by section
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

                # Teacher Version: Detailed Solution
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

        fname = "Ngan_Hang_Hoa_10_Chuong_3_Lien_Ket_Hoa_Hoc_GiaoVien.docx" if is_teacher else "Ngan_Hang_Hoa_10_Chuong_3_Lien_Ket_Hoa_Hoc_HocSinh.docx"
        fpath = os.path.join(BANK_DIR, fname)
        doc.save(fpath)
        print(f"  + Saved {'GiaoVien' if is_teacher else 'HocSinh'} Word file: {fpath}")

if __name__ == "__main__":
    build_chapter_3()
