# -*- coding: utf-8 -*-
import os
import sys

sys.path.append(r"c:\Antigravity_Thanh\scripts")
import hoa12_to36_data as d

img_dir = d.IMG_DIR
all_items = (
    d.EXERCISES_DANG1 + d.EXERCISES_DANG2 + d.EXERCISES_DANG3 +
    d.EXERCISES_DANG4 + d.TONG_HOP_MCQ_TF + d.PHAN3_SHORT_ANSWER
)

missing = 0
found = 0
for ex in all_items:
    img = ex.get("image")
    if img:
        p = os.path.join(img_dir, img)
        if not os.path.exists(p):
            print("MISSING:", img, "in", ex.get("id"))
            missing += 1
        else:
            found += 1

for sec in [d.THEORY_DANG1, d.THEORY_DANG2, d.THEORY_DANG3, d.THEORY_DANG4]:
    for s in sec.get("sections", []):
        img = s.get("image")
        if img:
            p = os.path.join(img_dir, img)
            if not os.path.exists(p):
                print("MISSING:", img, "in theory", sec.get("title"))
                missing += 1
            else:
                found += 1

print(f"Images verified: {found} found, {missing} missing.")
