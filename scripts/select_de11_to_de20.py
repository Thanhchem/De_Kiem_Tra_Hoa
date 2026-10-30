# -*- coding: utf-8 -*-
import os
import sys
import json
import re
from difflib import SequenceMatcher

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_5_de_gk1_hoa12 import (
    TESTS_SPEC, BANK_PATH, STEM_PATCHES, P2_ITEM_PATCHES,
    ANSWER_PATCHES, SOLUTION_PATCHES, parse_abcd_from_opt_lines
)
from merge_cvl_c1_to_bank import norm_for_compare

with open(BANK_PATH, "r", encoding="utf-8") as f:
    bank = json.load(f)

bank_map = {}
for ch in ["C1", "C2", "C3"]:
    for q in bank[ch]["questions"]:
        bank_map[(ch, q["q_num"])] = q

used = set()
used_norms = []
for t in TESTS_SPEC[:10]:
    for pk in ["P1", "P2", "P3"]:
        for ch, qn in t[pk]:
            used.add((ch, qn))
            st_lns = STEM_PATCHES.get((ch, qn), bank_map[(ch, qn)]["stem_lines"])
            used_norms.append((ch, qn, pk, norm_for_compare(" ".join(st_lns))))


def is_valid_candidate(ch, q, strict_no_multi_img=True):
    pair = (ch, q["q_num"])
    if pair in used:
        return False
    stem_lns = STEM_PATCHES.get(pair, q.get("stem_lines", []))
    stem = " ".join(stem_lns).strip()
    if len(stem) < 12:
        return False
    if strict_no_multi_img:
        if len(q.get("image_paths", [])) > 1:
            return False
        if "  " in stem or stem.endswith(("là:", "sau:", "như sau:")):
            if not q.get("image_paths") and not q.get("tables") and len(stem_lns) == 1:
                return False
    ans = ANSWER_PATCHES.get(pair, (q.get("answer") or "").strip())
    if not ans:
        return False
    if q["part"] == "P1":
        opts = parse_abcd_from_opt_lines(q.get("opt_lines", []))
        if len(opts) < 4 or any(not opts.get(k) for k in ["A", "B", "C", "D"]):
            return False
        if ans not in ["A", "B", "C", "D"]:
            return False
    elif q["part"] == "P2":
        items = P2_ITEM_PATCHES.get(pair, q.get("p2_items", {}))
        if len(items) < 4 or any(not items.get(k) for k in ["a", "b", "c", "d"]):
            return False
        if "a:" not in ans or "d:" not in ans:
            return False
    nm = norm_for_compare(stem)
    for uch, uqn, upk, unm in used_norms:
        if upk == q["part"]:
            if SequenceMatcher(None, nm, unm).ratio() >= 0.78:
                return False
    return True


def pick_pool(ch, part, topic_filter=None, count=10):
    res = []
    for strict in (True, False):
        if len(res) >= count:
            break
        for q in bank[ch]["questions"]:
            if q["part"] != part:
                continue
            if topic_filter and not topic_filter(q):
                continue
            if is_valid_candidate(ch, q, strict_no_multi_img=strict):
                res.append((ch, q["q_num"]))
                used.add((ch, q["q_num"]))
                st_lns = STEM_PATCHES.get((ch, q["q_num"]), q.get("stem_lines", []))
                used_norms.append((ch, q["q_num"], part, norm_for_compare(" ".join(st_lns))))
                if len(res) == count:
                    break
    return res


# Build pools for 10 exams (Đề 11 -> Đề 20)
# P1:
# Slots 1, 2, 3: C1 Ester (30 questions)
p1_c1_ester = pick_pool("C1", "P1", lambda q: q["topic"] == "Ester", 30)
# Slots 4, 5: C1 Chất béo (20 questions)
p1_c1_lipid = pick_pool("C1", "P1", lambda q: q["topic"] == "Chất béo", 20)
# Slots 6, 7, 8: C1 Xà phòng & CGR + Chất béo/Ester (30 questions)
p1_c1_soap = pick_pool("C1", "P1", lambda q: q["topic"] == "Xà phòng và chất giặt rửa", 30)
if len(p1_c1_soap) < 30:
    need = 30 - len(p1_c1_soap)
    p1_c1_soap += pick_pool("C1", "P1", lambda q: q["topic"] in ("Chất béo", "Ester"), need)

# Slots 9, 10: C2 Glucose – Fructose (20 questions)
p1_c2_glu = pick_pool("C2", "P1", lambda q: q["topic"] == "Glucose – Fructose", 20)
# Slots 11, 12: C2 Saccharose – Maltose (20 questions)
p1_c2_sac = pick_pool("C2", "P1", lambda q: q["topic"] == "Saccharose – Maltose", 20)
if len(p1_c2_sac) < 20:
    p1_c2_sac += pick_pool("C2", "P1", lambda q: True, 20 - len(p1_c2_sac))
# Slots 13, 14, 15: C2 Tinh bột – Cellulose & Tổng hợp Carbohydrate (30 questions)
p1_c2_poly = pick_pool("C2", "P1", lambda q: q["topic"] == "Tinh bột – Cellulose", 20)
p1_c2_all = pick_pool("C2", "P1", lambda q: True, 10)
p1_c2_3slots = p1_c2_poly + p1_c2_all

# Slots 16, 17, 18: C3 Amine (30 questions)
p1_c3_amine = pick_pool("C3", "P1", lambda q: q["topic"] == "Amine", 30)

# P2 (4 slots x 10 exams):
# Slot 1: C1 (10 questions)
p2_c1 = pick_pool("C1", "P2", lambda q: True, 10)
# Slot 2: C2 Glucose/Fructose/Saccharose (10 questions)
p2_c2_s2 = pick_pool("C2", "P2", lambda q: q["topic"] in ("Glucose – Fructose", "Saccharose – Maltose"), 10)
# Slot 3: C2 Tinh bột/Cellulose/Saccharose (10 questions)
p2_c2_s3 = pick_pool("C2", "P2", lambda q: True, 10)
# Slot 4: C3 Amine (+ q 145) (10 questions)
p2_c3 = pick_pool("C3", "P2", lambda q: q["topic"] == "Amine" or q["q_num"] == 145, 10)
if len(p2_c3) < 10:
    p2_c3 += pick_pool("C3", "P2", lambda q: True, 10 - len(p2_c3))

# P3 (6 slots x 10 exams):
# Slot 1: C1 Ester (10 questions)
p3_c1_ester = pick_pool("C1", "P3", lambda q: q["topic"] == "Ester", 10)
# Slot 2: C1 Chất béo (10 questions)
p3_c1_lipid = pick_pool("C1", "P3", lambda q: q["topic"] == "Chất béo", 10)
# Slot 3: C1 Xà phòng & CGR (10 questions)
p3_c1_soap = pick_pool("C1", "P3", lambda q: q["topic"] == "Xà phòng và chất giặt rửa", 10)
# Slot 4: C2 Glucose/Saccharose (10 questions)
p3_c2_s4 = pick_pool("C2", "P3", lambda q: q["topic"] in ("Glucose – Fructose", "Saccharose – Maltose"), 10)
# Slot 5: C2 Tinh bột – Cellulose (10 questions)
p3_c2_s5 = pick_pool("C2", "P3", lambda q: q["topic"] == "Tinh bột – Cellulose", 10)
# Slot 6: C3 Amine (10 questions)
p3_c3_s6 = pick_pool("C3", "P3", lambda q: q["topic"] == "Amine", 10)

print("Pools sizes:")
print("P1:", len(p1_c1_ester), len(p1_c1_lipid), len(p1_c1_soap), len(p1_c2_glu), len(p1_c2_sac), len(p1_c2_3slots), len(p1_c3_amine))
print("P2:", len(p2_c1), len(p2_c2_s2), len(p2_c2_s3), len(p2_c3))
print("P3:", len(p3_c1_ester), len(p3_c1_lipid), len(p3_c1_soap), len(p3_c2_s4), len(p3_c2_s5), len(p3_c3_s6))

new_tests = []
for i in range(10):
    de_num = 11 + i
    ma_de = str(111 + i)
    p1_list = [
        p1_c1_ester[i], p1_c1_ester[10 + i], p1_c1_ester[20 + i],
        p1_c1_lipid[i], p1_c1_lipid[10 + i],
        p1_c1_soap[i], p1_c1_soap[10 + i], p1_c1_soap[20 + i],
        p1_c2_glu[i], p1_c2_glu[10 + i],
        p1_c2_sac[i], p1_c2_sac[10 + i],
        p1_c2_3slots[i], p1_c2_3slots[10 + i], p1_c2_3slots[20 + i],
        p1_c3_amine[i], p1_c3_amine[10 + i], p1_c3_amine[20 + i],
    ]
    p2_list = [p2_c1[i], p2_c2_s2[i], p2_c2_s3[i], p2_c3[i]]
    p3_list = [p3_c1_ester[i], p3_c1_lipid[i], p3_c1_soap[i], p3_c2_s4[i], p3_c2_s5[i], p3_c3_s6[i]]
    new_tests.append({
        "de_num": de_num,
        "ma_de": ma_de,
        "P1": p1_list,
        "P2": p2_list,
        "P3": p3_list,
    })

print("\nGenerated 10 new tests (De 11 -> De 20):")
print(json.dumps(new_tests, ensure_ascii=False, indent=2))
