# -*- coding: utf-8 -*-
"""
Soạn thêm 10 Đề Kiểm Tra Giữa Học Kì I - Hóa Học 12 (Đề 11 -> Đề 20, Mã đề 111 -> 120):
- Mỗi đề gồm 28 câu theo đúng Ma trận đặc tả (Phần I: 18 câu, Phần II: 4 câu Đúng/Sai, Phần III: 6 câu Trả lời ngắn).
- Toàn bộ 280 câu hỏi của Đề 11 -> Đề 20 KHÔNG TRÙNG NHAU và KHÔNG TRÙNG với 280 câu của Đề 01 -> Đề 10 (tổng 560 câu duy nhất).
- Mỗi đề xuất đủ 3 phiên bản Word (.docx):
  1. Bản Học sinh (De_Kiem_Tra_GK1_Hoa_12_DeXX_HocSinh.docx)
  2. Bản Giáo viên có gắn nhãn mã câu hỏi (De_Kiem_Tra_GK1_Hoa_12_DeXX_GiaoVien.docx)
  3. Bản đăng Azota chuẩn cấu trúc Azota (De_Kiem_Tra_GK1_Hoa_12_DeXX_Azota.docx) - xuất trọn bộ cả Đề 01 -> Đề 20.
"""

import os
import sys
import copy
import json
from PIL import Image

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_5_de_gk1_hoa12 import (
    BANK_PATH, OUT_DIR, BANK_DIR, TESTS_SPEC,
    STEM_PATCHES, P2_ITEM_PATCHES, ANSWER_PATCHES, SOLUTION_PATCHES,
    generate_single_test_docx
)
from generate_azota_gk1_hoa12 import generate_azota_docx
from extract_q_assets import extract_chapter_assets
from select_de11_to_de20 import new_tests


def enrich_assets_map_with_json_questions(assets_map, bank_data):
    """Bổ sung hình ảnh (lọc bỏ ảnh công thức inline quá nhỏ) và bảng biểu của các câu mới trong JSON vào assets_map."""
    for ch_int in [1, 2, 3]:
        ch_code = f"C{ch_int}"
        ch_assets = assets_map.setdefault(ch_int, {})
        for q in bank_data[ch_code]["questions"]:
            qnum = q["q_num"]
            if qnum not in ch_assets:
                valid_imgs = []
                raw_imgs = q.get("image_paths", [])
                # Nếu câu chỉ có 1-2 ảnh sơ đồ/thí nghiệm lớn (chiều rộng > 140px và chiều cao > 70px) thì giữ lại
                if len(raw_imgs) <= 2:
                    for imp in raw_imgs:
                        if os.path.exists(imp):
                            try:
                                with Image.open(imp) as im:
                                    w, h = im.size
                                    if w >= 140 and h >= 65:
                                        valid_imgs.append(imp)
                            except Exception:
                                pass
                ch_assets[qnum] = {
                    "images": valid_imgs,
                    "tables": q.get("tables", []),
                }


if __name__ == "__main__":
    with open(BANK_PATH, "r", encoding="utf-8") as f:
        bank_data = json.load(f)

    bank_map = {}
    for ch_code in ["C1", "C2", "C3"]:
        for q in bank_data[ch_code]["questions"]:
            bank_map[(ch_code, q["q_num"])] = q

    all_20_tests = TESTS_SPEC[:10] + new_tests
    all_used = []
    for t in all_20_tests:
        for pk in ["P1", "P2", "P3"]:
            for pair in t[pk]:
                all_used.append((pair[0], pair[1]))

    print(f"Tổng số câu hỏi trong trọn bộ 20 đề (Đề 01 -> Đề 20): {len(all_used)} câu")
    print(f"Số câu hỏi duy nhất (không trùng lặp): {len(set(all_used))} câu")
    assert len(all_used) == len(set(all_used)) == 560, "Phát hiện câu hỏi trùng lặp!"

    print("Đang chuẩn bị tài nguyên hình ảnh & bảng dữ liệu...")
    assets_map = {c: extract_chapter_assets(c) for c in [1, 2, 3]}
    enrich_assets_map_with_json_questions(assets_map, bank_data)

    # 1. Xuất Bản Học sinh + Bản Giáo viên + Bản Azota cho 10 đề mới (Đề 11 -> Đề 20)
    for t_cfg in new_tests:
        de_num = t_cfg["de_num"]
        ma_de = t_cfg["ma_de"]
        print(f"\n--- Đang xuất Đề số {de_num:02d} (Mã đề {ma_de}) ---")
        generate_single_test_docx(t_cfg, bank_map, assets_map, is_teacher=False)
        generate_single_test_docx(t_cfg, bank_map, assets_map, is_teacher=True)
        generate_azota_docx(t_cfg, bank_map, assets_map)

    # 2. Đảm bảo cả 10 đề cũ (Đề 01 -> Đề 10) cũng có đầy đủ file Azota
    print("\n--- Kiểm tra & xuất đầy đủ file Azota cho Đề 01 -> Đề 10 ---")
    for t_cfg in TESTS_SPEC[:10]:
        generate_azota_docx(t_cfg, bank_map, assets_map)

    print("\nHOÀN TẤT XUẤT 10 ĐỀ MỚI (ĐỀ 11 -> ĐỀ 20: HỌC SINH, GIÁO VIÊN, AZOTA) & TRỌN BỘ 20 FILE AZOTA!")
