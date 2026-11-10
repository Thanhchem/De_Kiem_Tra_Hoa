# -*- coding: utf-8 -*-
import sys, json

sys.stdout.reconfigure(encoding='utf-8')
with open(r'c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Tron_Bo_9_Chuong.json', 'r', encoding='utf-8') as f:
    bank = json.load(f)

def get_candidates(chapter, topic, part, level):
    return [q for q in bank[chapter]['questions'] if q.get('part') == part and (topic is None or q.get('topic') == topic) and q.get('level') == level]

print("C1 Ester NB:", len(get_candidates('C1', 'Ester', 'P1', 'NB')))
print("C1 Ester TH:", len(get_candidates('C1', 'Ester', 'P1', 'TH')))
print("C1 Xà phòng TH:", len(get_candidates('C1', 'Xà phòng và chất giặt rửa', 'P1', 'TH')))
print("C2 Carbo NB:", len(get_candidates('C2', None, 'P1', 'NB')))
print("C3 Amine TH:", len(get_candidates('C3', 'Amine', 'P1', 'TH')))
print("C3 Amino acid NB:", len(get_candidates('C3', 'Amino acid', 'P1', 'NB')))
print("C3 Protein NB:", len(get_candidates('C3', 'Peptide – Protein – Enzyme', 'P1', 'NB')))
print("C3 Protein TH:", len(get_candidates('C3', 'Peptide – Protein – Enzyme', 'P1', 'TH')))
print("P2 Carbo:", len([q for q in bank['C2']['questions'] if q.get('part') == 'P2']))
print("P2 C3:", len([q for q in bank['C3']['questions'] if q.get('part') == 'P2']))
print("P3 Xà phòng NB:", len(get_candidates('C1', 'Xà phòng và chất giặt rửa', 'P3', 'NB')))
print("P3 Amine NB:", len(get_candidates('C3', 'Amine', 'P3', 'NB')))
print("P3 Amino acid TH:", len(get_candidates('C3', 'Amino acid', 'P3', 'TH')))
print("P3 Amine VD:", len(get_candidates('C3', 'Amine', 'P3', 'VD')))
