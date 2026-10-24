# -*- coding: utf-8 -*-
"""
Công cụ Rút đề kiểm tra tự động theo Ma trận đặc tả từ Ngân hàng câu hỏi Hóa học 12.
Hỗ trợ chọn số lượng câu hỏi theo:
  (Chương, Chủ đề, Phần I / II / III, Mức độ Nhận biết [NB] / Thông hiểu [TH] / Vận dụng [VD])
"""

import os
import json
import random
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

BANK_JSON = r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Tron_Bo_9_Chuong.json"


def load_bank():
    with open(BANK_JSON, "r", encoding="utf-8") as f:
        return json.load(f)


def select_questions_by_matrix(matrix_spec, seed=2026):
    """
    matrix_spec: list of dicts, e.g.:
    [
        {"chapter": "C1", "topic": "Ester", "part": "P1", "level": "NB", "count": 4},
        {"chapter": "C1", "topic": "Chất béo", "part": "P1", "level": "TH", "count": 2},
        ...
    ]
    """
    bank = load_bank()
    rng = random.Random(seed)
    selected = {"P1": [], "P2": [], "P3": []}

    for rule in matrix_spec:
        ch = rule["chapter"]
        tp = rule.get("topic")
        pt = rule["part"]
        lv = rule["level"]
        cnt = rule["count"]

        pool = [
            q for q in bank[ch]["questions"]
            if q["part"] == pt
            and (tp is None or q["topic"] == tp)
            and (lv is None or q["level"] == lv)
        ]
        if len(pool) < cnt:
            print(f"[Cảnh báo] Pool ({ch}, {tp}, {pt}, {lv}) chỉ có {len(pool)} câu (< {cnt} yêu cầu). Lấy tối đa {len(pool)} câu.")
            chosen = pool
        else:
            chosen = rng.sample(pool, cnt)
        selected[pt].extend(chosen)

    return selected


if __name__ == "__main__":
    bank = load_bank()
    print("=== TỔNG QUAN NGÂN HÀNG CÂU HỎI HÓA 12 (CHƯƠNG 1, 2, 3) ===")
    total_all = 0
    for ch_code, ch_data in bank.items():
        qs = ch_data["questions"]
        total_all += len(qs)
        print(f"\n* {ch_data['chapter_name']} ({len(qs)} câu):")
        for tp in ch_data["topics"]:
            tp_qs = [q for q in qs if q["topic"] == tp]
            p1_str = "/".join(str(sum(1 for q in tp_qs if q["part"] == "P1" and q["level"] == lv)) for lv in ["NB", "TH", "VD"])
            p2_str = "/".join(str(sum(1 for q in tp_qs if q["part"] == "P2" and q["level"] == lv)) for lv in ["NB", "TH", "VD"])
            p3_str = "/".join(str(sum(1 for q in tp_qs if q["part"] == "P3" and q["level"] == lv)) for lv in ["NB", "TH", "VD"])
            print(f"  - {tp:28s}: Tổng {len(tp_qs):3d} câu | P1(NB/TH/VD)={p1_str:8s} | P2(NB/TH/VD)={p2_str:6s} | P3(NB/TH/VD)={p3_str:6s}")
    print(f"\n=> TỔNG SỐ CÂU HỎI ĐÃ PHÂN LOẠI: {total_all} câu.")
