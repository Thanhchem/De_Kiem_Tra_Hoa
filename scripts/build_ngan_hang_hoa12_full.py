# -*- coding: utf-8 -*-
"""
Xây dựng TRỌN BỘ Ngân hàng câu hỏi Hóa học 12 (Chương 1 -> Chương 8 + Chuyên đề học tập 9)
Phân loại chi tiết theo:
- Chương & Chủ đề (Bài học)
- Dạng thức câu hỏi (Phần I: Nhiều lựa chọn, Phần II: Đúng/Sai, Phần III: Trả lời ngắn)
- Mức độ nhận thức (Nhận biết [NB] - Thông hiểu [TH] - Vận dụng [VD]) đến từng câu và từng ý a,b,c,d
- Xuất cơ sở dữ liệu JSON Toàn bộ 9 Chương + Bộ 18 file Word (.docx) Học sinh & Giáo viên chuẩn "FORM ĐỀ KIỂM TRA".
"""

import os
import re
import json
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from build_ngan_hang_hoa12_c123 import (
    FOLDER, OUT_DIR, BANK_DIR, LEVEL_NAMES, LEVEL_COLORS_HEX,
    COLOR_HF_RED, COLOR_PRIMARY, COLOR_SOLUTION, COLOR_CORRECT,
    iter_block_items, classify_p2_item_level, classify_p2_overall_level,
    extract_underlined_p1_answer, setup_page_and_hf, insert_matrix_summary_at_top,
    CHAPTERS as CHAPTERS_123
)

ALL_CHAPTERS = dict(CHAPTERS_123)
ALL_CHAPTERS.update({
    4: {
        "code": "C4",
        "name": "CHƯƠNG 4: POLYMER",
        "short_name": "Polymer",
        "file_prefix": "Ngan_Hang_Hoa_12_Chuong_4_Polymer",
        "da_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD4 - Polymer. TTB. ok - DA chi tiet.docx"),
        "in_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD4 - Polymer. TTB. ok - in.docx"),
        "topics": ["Đại cương về Polymer", "Vật liệu Polymer"],
    },
    5: {
        "code": "C5",
        "name": "CHƯƠNG 5: PIN ĐIỆN VÀ ĐIỆN PHÂN",
        "short_name": "Pin điện – Điện phân",
        "file_prefix": "Ngan_Hang_Hoa_12_Chuong_5_Pin_Dien_Dien_Phan",
        "da_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD5 - Pin dien - dien phan. TTB. ok - DA chi tiet.docx"),
        "in_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD5 - Pin dien - dien phan. TTB. ok - in.docx"),
        "topics": ["Pin điện", "Điện phân"],
    },
    6: {
        "code": "C6",
        "name": "CHƯƠNG 6: ĐẠI CƯƠNG VỀ KIM LOẠI",
        "short_name": "Đại cương kim loại",
        "file_prefix": "Ngan_Hang_Hoa_12_Chuong_6_Dai_Cuong_Kim_Loai",
        "da_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD6 - Dai cuong kim loai. TTB. ok - DA chi tiet.docx"),
        "in_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD6 - Dai cuong kim loai. TTB . ok - in.docx"),
        "topics": [
            "Cấu tạo và tính chất vật lí của kim loại",
            "Tính chất hóa học và phương pháp tách kim loại",
            "Hợp kim và ăn mòn kim loại"
        ],
    },
    7: {
        "code": "C7",
        "name": "CHƯƠNG 7: NGUYÊN TỐ NHÓM IA VÀ NHÓM IIA",
        "short_name": "Nguyên tố nhóm IA, IIA",
        "file_prefix": "Ngan_Hang_Hoa_12_Chuong_7_Nhom_IA_IIA",
        "da_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD7 - Nguyen to nhom IA, IIA. TTB. ok - DA chi tiet.docx"),
        "in_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD7 - Nguyen to nhom IA, IIA. TTB. ok - in.docx"),
        "topics": ["Nguyên tố nhóm IA", "Nguyên tố nhóm IIA"],
    },
    8: {
        "code": "C8",
        "name": "CHƯƠNG 8: SƠ LƯỢC VỀ DÃY KIM LOẠI CHUYỂN TIẾP THỨ NHẤT VÀ PHỨC CHẤT",
        "short_name": "Kim loại chuyển tiếp & Phức chất",
        "file_prefix": "Ngan_Hang_Hoa_12_Chuong_8_Kim_Loai_Chuyen_Tiep_Phuc_Chat",
        "da_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD8 - Kim loai chuyen tiep va phuc chat. TTB. ok - DA chi tiet.docx"),
        "in_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD8 - Kim loai chuyen tiep va phuc chat. TTB. ok - in.docx"),
        "topics": ["Kim loại chuyển tiếp dãy thứ nhất", "Phức chất"],
    },
    9: {
        "code": "C9",
        "name": "CHUYÊN ĐỀ HỌC TẬP HÓA HỌC 12",
        "short_name": "Chuyên đề học tập Hóa 12",
        "file_prefix": "Ngan_Hang_Hoa_12_Chuyen_De_Hoc_Tap",
        "da_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD9 - Chuyen de hoc tap. TTB. ok - DA chi tiet.docx"),
        "in_file": os.path.join(FOLDER, "PD 2026 - Hoa 12. CD9 - Chuyen de hoc tap. TTB. ok - in.docx"),
        "topics": [
            "CĐ1: Cơ chế phản ứng trong Hóa hữu cơ",
            "CĐ2: Trải nghiệm thực hành Hóa vô cơ",
            "CĐ3: Một số vấn đề cơ bản về phức chất"
        ],
    }
})


def normalize_topic_all(chap_num, raw_topic, q_num, q_text):
    from build_ngan_hang_hoa12_c123 import normalize_topic as norm_123
    if chap_num in (1, 2, 3):
        return norm_123(chap_num, raw_topic, q_num, q_text)

    rt = (raw_topic or "").upper().strip()
    t = q_text.lower()

    if chap_num == 4:
        if "ĐẠI CƯƠNG" in rt:
            return "Đại cương về Polymer"
        if "VẬT LIỆU" in rt:
            return "Vật liệu Polymer"
        # Content inference when raw_topic is empty (e.g. P2 or P3 VD)
        if any(k in t for k in [
            'tơ ', 'cao su', 'chất dẻo', 'keo dán', 'vật liệu', 'lốp xe', 'ống nhựa',
            'nylon', 'nilon', 'capron', 'buna', 'pe ', 'pvc', 'ps', 'pmma', 'thủy tinh hữu cơ',
            'tơ tằm', 'tơ visco', 'tơ nitron', 'olon', 'kevlar', 'nhựa'
        ]):
            return "Vật liệu Polymer"
        return "Đại cương về Polymer"

    elif chap_num == 5:
        if "ĐIỆN PHÂN" in rt:
            return "Điện phân"
        if "PIN ĐIỆN" in rt:
            return "Pin điện"
        if any(k in t for k in ['điện phân', 'bình điện phân', 'mạ điện', 'tinh luyện', 'faraday']):
            return "Điện phân"
        return "Pin điện"

    elif chap_num == 6:
        if "CẤU TẠO" in rt or "VẬT LÍ" in rt:
            return "Cấu tạo và tính chất vật lí của kim loại"
        if "TÍNH CHẤT HÓA HỌC" in rt or "TÁCH KIM LOẠI" in rt:
            return "Tính chất hóa học và phương pháp tách kim loại"
        if "HỢP KIM" in rt or "ĂN MÒN" in rt:
            return "Hợp kim và ăn mòn kim loại"
        # Content inference when raw_topic is empty (e.g. P2 or P3 VD)
        if any(k in t for k in ['ăn mòn', 'hợp kim', 'gang', 'thép', 'duralumin', 'bảo vệ bề mặt', 'phương pháp điện hóa']):
            return "Hợp kim và ăn mòn kim loại"
        if any(k in t for k in ['tính dẻo', 'dẫn điện', 'dẫn nhiệt', 'ánh kim', 'khối lượng riêng', 'nhiệt độ nóng chảy', 'liên kết kim loại', 'mạng tinh thể']):
            return "Cấu tạo và tính chất vật lí của kim loại"
        return "Tính chất hóa học và phương pháp tách kim loại"

    elif chap_num == 7:
        if "NHÓM IA" in rt:
            return "Nguyên tố nhóm IA"
        if "NHÓM IIA" in rt:
            return "Nguyên tố nhóm IIA"
        if any(k in t for k in ['nước cứng', 'thạch cao', 'đá vôi', 'vôi sống', 'caco3', 'caso4', 'ca(oh)2', 'nhóm iia', 'kiềm thổ', 'magnesium', 'calcium', 'barium', 'mg', 'ca2+', 'ba2+']):
            return "Nguyên tố nhóm IIA"
        return "Nguyên tố nhóm IA"

    elif chap_num == 8:
        if "PHỨC CHẤT" in rt:
            return "Phức chất"
        if "CHUYỂN TIẾP" in rt:
            return "Kim loại chuyển tiếp dãy thứ nhất"
        if any(k in t for k in ['phối tử', 'nguyên tử trung tâm', 'cầu nội', 'cầu ngoại', 'phức chất', 'bát diện', 'tứ diện', 'vuông phẳng']):
            return "Phức chất"
        return "Kim loại chuyển tiếp dãy thứ nhất"

    elif chap_num == 9:
        if "CĐ1" in rt or "CƠ CHẾ" in rt:
            return "CĐ1: Cơ chế phản ứng trong Hóa hữu cơ"
        if "CĐ2" in rt or "VÔ CƠ" in rt:
            return "CĐ2: Trải nghiệm thực hành Hóa vô cơ"
        if "CĐ3" in rt or "PHỨC CHẤT" in rt:
            return "CĐ3: Một số vấn đề cơ bản về phức chất"
        return "CĐ1: Cơ chế phản ứng trong Hóa hữu cơ"

    return raw_topic


def parse_chapter_all(chap_num, info):
    doc = docx.Document(info["da_file"])
    blocks = list(iter_block_items(doc))

    questions = []
    current_part = "P1"
    current_level = "NB"
    current_topic = info["topics"][0]
    current_q = None
    seq_counter = 0

    TOPIC_HEADERS = {
        "ESTER", "CHẤT BÉO", "XÀ PHÒNG VÀ CHẤT GIẶT RỬA", "XÀ PHÒNG VÀ CHẤT GIẶT RỬA TỔNG HỢP",
        "KHÁI QUÁT VỀ CARBOHYDRATE – GLUCOSE, FRUCTOSE", "SACCHAROSE – MALTOSE", "TINH BỘT – CELLULOSE",
        "AMINE", "AMINO ACID", "PEPTIDE – PROTEIN - ENZYME",
        "ĐẠI CƯƠNG VỀ POLYMER", "VẬT LIỆU POLYMER",
        "PIN ĐIỆN", "ĐIỆN PHÂN",
        "CẤU TẠO VÀ TÍNH CHẤT VẬT LÍ CỦA KIM LOẠI", "TÍNH CHẤT HÓA HỌC VÀ PP TÁCH KIM LOẠI", "HỢP KIM VÀ ĂN MÒN KIM LOẠI",
        "NGUYÊN TỐ NHÓM IA", "NGUYÊN TỐ NHÓM IIA",
        "KIM LOẠI CHUYỂN TIẾP DÃY THỨ NHẤT", "PHỨC CHẤT",
        "CĐ1: CƠ CHẾ PHẢN ỨNG TRONG HÓA HỌC HỮU CƠ", "CĐ2: TRẢI NGHIỆM THỰC HÀNH HÓA VÔ CƠ", "CĐ3: MỘT SỐ VẤN ĐỀ CƠ BẢN VỀ PHỨC CHẤT"
    }

    def finalize_q(q):
        if not q:
            return
        full_q_text = "\n".join(q["stem_lines"] + q["opt_lines"])
        sol_text = "\n".join(q["sol_lines"]).strip()
        q["topic"] = normalize_topic_all(chap_num, q["raw_topic"], q["q_num"], full_q_text)

        if chap_num == 9 and q["part"] in ("P1", "P3"):
            q["level"] = classify_p2_item_level(full_q_text)

        if q["part"] == "P1":
            ans = q["underlined_ans"]
            if not ans and sol_text:
                m = re.search(r'Đáp\s+(?:án|số)\s*[:\s]*([A-D])\b', sol_text, re.IGNORECASE)
                if m:
                    ans = m.group(1).upper()
                else:
                    m2 = re.search(r'\b([A-D])\s+(?:đúng|sai)', sol_text)
                    if m2:
                        ans = m2.group(1).upper()
            if not ans and chap_num == 3 and q["q_num"] in (12, 77):
                ans = "A"
            q["answer"] = ans or ""
        elif q["part"] == "P2":
            item_levels = {}
            for key in ["a", "b", "c", "d"]:
                stmt = q["p2_items"].get(key, "")
                item_levels[key] = classify_p2_item_level(stmt)
            q["item_levels"] = item_levels
            q["level"] = classify_p2_overall_level(q["q_num"], chap_num, full_q_text, item_levels)

            p2_ans = {}
            for key in ["a", "b", "c", "d"]:
                m = re.search(rf'(?:^|\n)\s*\(?{key}[\)\.]\s*(Đúng|Sai)', sol_text, re.IGNORECASE)
                if m:
                    p2_ans[key] = "Đ" if m.group(1).lower().startswith("đ") else "S"
                elif key in q["underlined_p2"]:
                    p2_ans[key] = "Đ" if q["underlined_p2"][key] else "S"
                else:
                    p2_ans[key] = "S"
            q["answer"] = " - ".join(f"{k}:{p2_ans[k]}" for k in ["a", "b", "c", "d"])
        elif q["part"] == "P3":
            ans = ""
            m = re.search(r'Đáp\s+(?:số|án)\s*[:\.]?\s*([^\n\.]+)', sol_text, re.IGNORECASE)
            if m:
                ans = m.group(1).strip()
                m_num = re.match(r'^([+-]?\d+(?:[\.,]\d+)?%?)', ans)
                if m_num:
                    ans = m_num.group(1)
            q["answer"] = ans

        q["id"] = f"{info['code']}-{q['part']}-{q['level']}-{q['seq_num']:03d}"
        q["solution"] = sol_text
        questions.append(q)

    for btype, block in blocks:
        if btype == 'p':
            txt = block.text.strip()
            if not txt:
                continue
            tl = txt.lower()
            if 'trắc nghiệm nhiều phương án lựa chọn' in tl:
                finalize_q(current_q)
                current_q = None
                current_part = "P1"
                if chap_num != 9:
                    current_topic = ""
                continue
            if 'trắc nghiệm đúng' in tl and 'sai' in tl:
                finalize_q(current_q)
                current_q = None
                current_part = "P2"
                if chap_num != 9:
                    current_topic = ""
                continue
            if 'trắc nghiệm trả lời ngắn' in tl:
                finalize_q(current_q)
                current_q = None
                current_part = "P3"
                if chap_num != 9:
                    current_topic = ""
                continue
            if 'MỨC ĐỘ BIẾT' in txt:
                finalize_q(current_q)
                current_q = None
                current_level = "NB"
                current_topic = ""
                continue
            if 'MỨC ĐỘ HIỂU' in txt:
                finalize_q(current_q)
                current_q = None
                current_level = "TH"
                current_topic = ""
                continue
            if 'MỨC ĐỘ VẬN DỤNG' in txt:
                finalize_q(current_q)
                current_q = None
                current_level = "VD"
                current_topic = ""
                continue
            if re.match(r'^[-_]{4,}', txt) or 'HẾT' in txt:
                continue
            if txt.upper() in TOPIC_HEADERS:
                finalize_q(current_q)
                current_q = None
                current_topic = txt
                continue

            m_q = re.match(r'^Câu\s+(\d+)[\.\:]\s*(.*)', txt)
            if m_q:
                finalize_q(current_q)
                seq_counter += 1
                q_num = int(m_q.group(1))
                rest = m_q.group(2).strip()
                src_m = re.match(r'^(\[[^\]]+\])\s*(.*)', rest)
                source_tag = src_m.group(1) if src_m else ""
                stem_first = src_m.group(2) if src_m else rest
                current_q = {
                    "seq_num": seq_counter,
                    "q_num": q_num,
                    "part": current_part,
                    "level": current_level,
                    "raw_topic": current_topic,
                    "source_tag": source_tag,
                    "stem_lines": [stem_first] if stem_first else [],
                    "opt_lines": [],
                    "p2_items": {},
                    "underlined_ans": None,
                    "underlined_p2": {},
                    "sol_lines": [],
                    "in_sol": False,
                }
                continue

            if current_q:
                if txt.startswith("Hướng dẫn giải"):
                    current_q["in_sol"] = True
                    continue
                if current_q["in_sol"]:
                    current_q["sol_lines"].append(txt)
                else:
                    if current_q["part"] == "P1" and re.match(r'^[A-D]\.', txt):
                        current_q["opt_lines"].append(txt)
                        u_ans = extract_underlined_p1_answer([block])
                        if u_ans:
                            current_q["underlined_ans"] = u_ans
                    elif current_q["part"] == "P2" and re.match(r'^([a-d])[\)\.]\s*(.*)', txt):
                        m_it = re.match(r'^([a-d])[\)\.]\s*(.*)', txt)
                        k = m_it.group(1)
                        current_q["p2_items"][k] = m_it.group(2).strip()
                        current_q["opt_lines"].append(txt)
                        has_u = any(r.underline for r in block.runs if r.text.strip())
                        current_q["underlined_p2"][k] = has_u
                    else:
                        current_q["stem_lines"].append(txt)

        elif btype == 'tbl':
            for row in block.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        pt = p.text.strip()
                        m_q = re.match(r'^Câu\s+(\d+)[\.\:]\s*(.*)', pt)
                        if m_q and (not current_q or int(m_q.group(1)) != current_q["q_num"]):
                            finalize_q(current_q)
                            seq_counter += 1
                            q_num = int(m_q.group(1))
                            rest = m_q.group(2).strip()
                            src_m = re.match(r'^(\[[^\]]+\])\s*(.*)', rest)
                            source_tag = src_m.group(1) if src_m else ""
                            stem_first = src_m.group(2) if src_m else rest
                            current_q = {
                                "seq_num": seq_counter,
                                "q_num": q_num,
                                "part": current_part,
                                "level": current_level,
                                "raw_topic": current_topic,
                                "source_tag": source_tag,
                                "stem_lines": [stem_first] if stem_first else [],
                                "opt_lines": [],
                                "p2_items": {},
                                "underlined_ans": None,
                                "underlined_p2": {},
                                "sol_lines": [],
                                "in_sol": False,
                            }
                        elif current_q and not current_q["in_sol"] and pt:
                            if current_q["part"] == "P1" and re.match(r'^[A-D]\.', pt):
                                u_ans = extract_underlined_p1_answer([p])
                                if u_ans:
                                    current_q["underlined_ans"] = u_ans

    finalize_q(current_q)
    return questions


def inject_question_tags_seq(doc, questions, is_teacher):
    """Gắn nhãn Mã câu hỏi theo thứ tự xuất hiện (seq_num) để hỗ trợ cả các file đánh lại Câu 1 ở từng mục."""
    seq_idx = [0]
    current_q = [None]
    in_sol = [False]

    def tag_question_paragraph(p, q):
        lv_name = LEVEL_NAMES[q["level"]]
        lv_hex = LEVEL_COLORS_HEX[q["level"]]
        if q["part"] == "P2":
            il = q.get("item_levels", {})
            il_str = ", ".join(f"{k}:{LEVEL_NAMES.get(il.get(k,'TH'), 'Hiểu')}" for k in ["a", "b", "c", "d"])
            ans_str = f" | ĐA: {q['answer']}" if is_teacher else ""
            tag_text = f"[{q['id']} | {lv_name} – {q['topic']} ({il_str}){ans_str}] "
        else:
            ans_str = f" | ĐA: {q['answer']}" if (is_teacher and q["answer"]) else ""
            tag_text = f"[{q['id']} | {lv_name} – {q['topic']}{ans_str}] "

        new_r = parse_xml(
            f'<w:r {nsdecls("w")}>'
            f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'<w:b/><w:color w:val="{lv_hex}"/><w:sz w:val="21"/></w:rPr>'
            f'<w:t xml:space="preserve">{tag_text}</w:t>'
            f'</w:r>'
        )
        runs = p._p.xpath('./w:r')
        inserted = False
        acc_text = ""
        for r_el in runs:
            t_str = "".join(r_el.xpath('.//w:t/text()'))
            acc_text += t_str
            if re.match(r'^\s*Câu\s+\d+[\.\:]', acc_text):
                r_el.addnext(new_r)
                inserted = True
                break
        if not inserted and runs:
            runs[0].addnext(new_r)

    def tag_p2_statement_paragraph(p, q, key):
        il = q.get("item_levels", {})
        lv = il.get(key, "TH")
        lv_name = LEVEL_NAMES[lv]
        lv_hex = LEVEL_COLORS_HEX[lv]
        tag_text = f" [{lv_name}] "
        new_r = parse_xml(
            f'<w:r {nsdecls("w")}>'
            f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'<w:b/><w:color w:val="{lv_hex}"/><w:sz w:val="21"/></w:rPr>'
            f'<w:t xml:space="preserve">{tag_text}</w:t>'
            f'</w:r>'
        )
        runs = p._p.xpath('./w:r')
        for r_el in runs:
            t_str = "".join(r_el.xpath('.//w:t/text()'))
            if re.search(rf'^{key}[\)\.]', t_str.strip()):
                r_el.addnext(new_r)
                return
        if runs:
            runs[0].addnext(new_r)

    for btype, block in iter_block_items(doc):
        if btype == 'p':
            txt = block.text.strip()
            if not txt:
                continue
            m_q = re.match(r'^Câu\s+(\d+)[\.\:]', txt)
            if m_q and seq_idx[0] < len(questions):
                current_q[0] = questions[seq_idx[0]]
                seq_idx[0] += 1
                in_sol[0] = False
                tag_question_paragraph(block, current_q[0])
                continue
            if txt.startswith("Hướng dẫn giải"):
                in_sol[0] = True
                continue
            if current_q[0] and not in_sol[0] and current_q[0]["part"] == "P2":
                m_it = re.match(r'^([a-d])[\)\.]\s*', txt)
                if m_it:
                    tag_p2_statement_paragraph(block, current_q[0], m_it.group(1))
        elif btype == 'tbl':
            last_tbl_q_num = None
            for row in block.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        pt = p.text.strip()
                        m_q = re.match(r'^Câu\s+(\d+)[\.\:]', pt)
                        if m_q and int(m_q.group(1)) != last_tbl_q_num and seq_idx[0] < len(questions):
                            last_tbl_q_num = int(m_q.group(1))
                            current_q[0] = questions[seq_idx[0]]
                            seq_idx[0] += 1
                            in_sol[0] = False
                            tag_question_paragraph(p, current_q[0])


def build_chapter_word_files_seq(chap_num, info, questions):
    doc_hs = docx.Document(info["in_file"])
    setup_page_and_hf(doc_hs)
    inject_question_tags_seq(doc_hs, questions, is_teacher=False)
    insert_matrix_summary_at_top(doc_hs, info, questions, is_teacher=False)

    hs_path_1 = os.path.join(OUT_DIR, f"{info['file_prefix']}_HocSinh.docx")
    hs_path_2 = os.path.join(BANK_DIR, f"{info['file_prefix']}_HocSinh.docx")
    doc_hs.save(hs_path_1)
    doc_hs.save(hs_path_2)
    print(f"  + Saved Student Docx: {hs_path_1}")

    doc_gv = docx.Document(info["da_file"])
    setup_page_and_hf(doc_gv)
    inject_question_tags_seq(doc_gv, questions, is_teacher=True)
    insert_matrix_summary_at_top(doc_gv, info, questions, is_teacher=True)

    gv_path_1 = os.path.join(OUT_DIR, f"{info['file_prefix']}_GiaoVien.docx")
    gv_path_2 = os.path.join(BANK_DIR, f"{info['file_prefix']}_GiaoVien.docx")
    doc_gv.save(gv_path_1)
    doc_gv.save(gv_path_2)
    print(f"  + Saved Teacher Docx: {gv_path_1}")


if __name__ == "__main__":
    all_bank_json = {}
    grand_total = 0
    for cnum in sorted(ALL_CHAPTERS.keys()):
        info = ALL_CHAPTERS[cnum]
        qs = parse_chapter_all(cnum, info)
        grand_total += len(qs)
        all_bank_json[info["code"]] = {
            "chapter_number": cnum,
            "chapter_name": info["name"],
            "topics": info["topics"],
            "total_questions": len(qs),
            "questions": qs
        }
        if cnum >= 4:
            print(f"\nProcessing {info['name']} ({len(qs)} câu)...")
            build_chapter_word_files_seq(cnum, info, qs)

    full_json_path = os.path.join(BANK_DIR, "Ngan_Hang_Hoa_12_Tron_Bo_9_Chuong.json")
    with open(full_json_path, "w", encoding="utf-8") as f:
        json.dump(all_bank_json, f, ensure_ascii=False, indent=2)
    print(f"\n==========================================")
    print(f"ĐÃ HOÀN TẤT TRỌN BỘ 9 CHƯƠNG HÓA 12: {grand_total} CÂU HỎI")
    print(f"File JSON Tổng: {full_json_path}")
