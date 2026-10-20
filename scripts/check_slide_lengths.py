# -*- coding: utf-8 -*-
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import json

with open(r'c:\Antigravity_Thanh\scratch\all_10_tests.json', encoding='utf-8') as f:
    data = json.load(f)

for d in data:
    de_num = d['test_num']
    for s in range(9):
        q1 = d['p1'][s*2]
        q2 = d['p1'][s*2 + 1]
        t1 = q1['stem'] + ' ' + ' '.join(q1['opts'].values())
        t2 = q2['stem'] + ' ' + ' '.join(q2['opts'].values())
        total_len = len(t1) + len(t2)
        if total_len > 300:
            print(f"De {de_num} Slide {s+1} (Q{q1['num']} & Q{q2['num']}): {total_len} chars")

print("\n--- Kiểm tra Phần II ---")
for d in data:
    de_num = d['test_num']
    for q in d['p2']:
        t = q['stem'] + ' ' + ' '.join(q['items'].values())
        if len(t) > 350:
            print(f"De {de_num} P2 Q{q['num']}: {len(t)} chars")
