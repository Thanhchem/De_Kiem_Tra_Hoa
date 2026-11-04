# -*- coding: utf-8 -*-
"""
Script gom toan bo 180 cau hoi trac nghiem 4 dap an (Phan I)
tu 10 de on tap Chuong 1, 2, 3 Hoa hoc 12 thanh 1 file Word duy nhat (.docx)
chuan upload Azota:
- Danh so lien tuc tu Cau 1 den Cau 180.
- Khong co loi giai chi tiet (theo yeu cau nguoi dung).
- Gach chan dap an dung (underline = True) o nhan dap an (chuan Azota).
- Kem bang dap an 180 cau o cuoi de de Azota nhan dien kep 100%.
- Luu tai San_Pham/10_De_On_Tap_Chuong_123_Hoa12/ va San_Pham/.
"""

import os
import sys
import re
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Import module form chuan
sys.path.append(os.path.abspath(r"C:\Antigravity_Thanh\scripts"))
from form_de_kiem_tra import (
    setup_header_footer,
    build_title_block,
    set_run_style,
    COLOR_PRIMARY,
    COLOR_CORRECT
)

IN_DIR = r"C:\Antigravity_Thanh\San_Pham\10_De_On_Tap_Chuong_123_Hoa12"
OUT_FILE_1 = os.path.join(IN_DIR, "Tong_Hop_180_Cau_Trac_Nghiem_Chuong_123_Hoa12_Azota.docx")
OUT_FILE_2 = r"C:\Antigravity_Thanh\San_Pham\Tong_Hop_180_Cau_Trac_Nghiem_Chuong_123_Hoa12_Azota.docx"

def extract_p1_questions(doc_path):
    """Trích xuất 18 câu trắc nghiệm 4 đáp án từ mỗi file Giáo Viên"""
    doc = docx.Document(doc_path)
    questions = []
    cur_q = None
    state = None
    
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        if 'PHẦN II' in txt:
            if cur_q:
                questions.append(cur_q)
                cur_q = None
            break
            
        m_q = re.match(r'^Câu\s+(\d+)\.\s*(.*)', txt)
        if m_q and int(m_q.group(1)) <= 18:
            if cur_q:
                questions.append(cur_q)
            cur_q = {
                'orig_num': int(m_q.group(1)),
                'stem_p': p,
                'opt_paragraphs': [],
                'correct': None
            }
            state = 'stem'
            continue
            
        if cur_q:
            if '► Lời giải chi tiết:' in txt or 'Lời giải chi tiết:' in txt:
                state = 'sol'
                continue
            if state == 'sol':
                m_c = re.search(r'➔\s*Chọn\s+([A-D])\.', txt)
                if m_c and not cur_q['correct']:
                    cur_q['correct'] = m_c.group(1)
            elif state in ['stem', 'opts']:
                if re.search(r'^[A-D]\.', txt):
                    state = 'opts'
                    cur_q['opt_paragraphs'].append(p)
                    for r in p.runs:
                        if r.font.color and r.font.color.rgb and str(r.font.color.rgb).lower() == 'c62828':
                            m_opt = re.match(r'^([A-D])\.', r.text.strip())
                            if m_opt:
                                cur_q['correct'] = m_opt.group(1)
                elif state == 'stem':
                    # Đoạn văn bản bổ sung của đề (nếu có)
                    pass

    if cur_q:
        questions.append(cur_q)
    return questions

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=60, bottom=60, left=60, right=60):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="90A4AE", sz="4"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'<w:tblBorders {nsdecls("w")}><w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/></w:tblBorders>')
    tblPr.append(borders)

def build_merged_azota_doc():
    doc = docx.Document()
    setup_header_footer(doc)

    # 1. Khối tiêu đề chuẩn
    build_title_block(
        doc,
        school_line="TỔNG HỢP NGÂN HÀNG CÂU HỎI HÓA HỌC 12\nCHƯƠNG 1, 2, 3 (ESTER - CARBOHYDRATE - HỢP CHẤT NITROGEN)",
        exam_title="BỘ 180 CÂU TRẮC NGHIỆM ÔN TẬP CHƯƠNG 1, 2, 3",
        subject_line="MÔN: HOÁ HỌC - LỚP 12",
        num_pages="--",
        duration="180 phút (Dùng thi trực tuyến Azota)",
        exam_code="AZOTA-180",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Br = 80; Ag = 108",
        is_teacher=False
    )

    # Dòng chỉ dẫn Azota
    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.space_before = Pt(6)
    p_intro.paragraph_format.space_after = Pt(8)
    r_intro = p_intro.add_run("Thí sinh trả lời từ câu 1 đến câu 180. Mỗi câu hỏi chỉ chọn một phương án đúng nhất.")
    set_run_style(r_intro, font_size=11, italic=True)

    # 2. Gom toàn bộ câu hỏi từ 10 đề
    all_questions = []
    for de_num in range(1, 11):
        doc_path = os.path.join(IN_DIR, f"De_On_Tap_Hoa_12_De{de_num:02d}_GiaoVien.docx")
        qs = extract_p1_questions(doc_path)
        print(f"Đã nạp Đề {de_num:02d}: {len(qs)} câu (Đáp án hợp lệ: {len([q for q in qs if q['correct']])})")
        for q in qs:
            q["source_de"] = de_num
            all_questions.append(q)

    print(f"==> Tổng số câu hỏi gom được: {len(all_questions)} câu.")

    # 3. Ghi từng câu vào file Word
    for global_idx, q in enumerate(all_questions, 1):
        # 3.1. Thân câu hỏi (Stem)
        p_stem = doc.add_paragraph()
        p_stem.paragraph_format.space_before = Pt(5)
        p_stem.paragraph_format.space_after = Pt(2)

        # Nhãn câu: Câu X.
        r_num = p_stem.add_run(f"Câu {global_idx}. ")
        set_run_style(r_num, font_size=11.5, bold=True)

        # Copy các runs của thân câu hỏi (bỏ run đầu tiên là "Câu X. ")
        orig_stem_p = q['stem_p']
        for r_orig in orig_stem_p.runs[1:]:
            r_new = p_stem.add_run(r_orig.text)
            r_new.font.name = "Times New Roman"
            r_new.font.size = Pt(11.5)
            r_new.bold = r_orig.bold
            r_new.italic = r_orig.italic
            r_new.font.subscript = r_orig.font.subscript
            r_new.font.superscript = r_orig.font.superscript

        # 3.2. Các phương án lựa chọn A, B, C, D
        corr_ans = q['correct']
        for p_opt_orig in q['opt_paragraphs']:
            p_opt = doc.add_paragraph()
            p_opt.paragraph_format.left_indent = Cm(0.5)
            p_opt.paragraph_format.space_before = Pt(1)
            p_opt.paragraph_format.space_after = Pt(1)

            # Sao chép tab stops nếu có
            if p_opt_orig.paragraph_format.tab_stops:
                for ts in p_opt_orig.paragraph_format.tab_stops:
                    p_opt.paragraph_format.tab_stops.add_tab_stop(ts.position, ts.alignment)
            else:
                p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(8.5), WD_TAB_ALIGNMENT.LEFT)

            current_opt_letter = None
            for r_orig in p_opt_orig.runs:
                txt = r_orig.text
                m_lbl = re.match(r'^([A-D])\.\s*$', txt.strip())
                if m_lbl:
                    current_opt_letter = m_lbl.group(1)
                    is_corr = (current_opt_letter == corr_ans)
                    r_new = p_opt.add_run(txt)
                    r_new.font.name = "Times New Roman"
                    r_new.font.size = Pt(11.5)
                    r_new.bold = True
                    r_new.underline = is_corr  # GẠCH CHÂN ĐÁP ÁN ĐÚNG CHO AZOTA
                else:
                    r_new = p_opt.add_run(txt)
                    r_new.font.name = "Times New Roman"
                    r_new.font.size = Pt(11.5)
                    r_new.bold = False
                    r_new.italic = r_orig.italic
                    r_new.font.subscript = r_orig.font.subscript
                    r_new.font.superscript = r_orig.font.superscript

    # Dòng kết thúc phần câu hỏi
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(12)
    p_end.paragraph_format.space_after = Pt(12)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("-------------------------- HẾT 180 CÂU TRẮC NGHIỆM --------------------------")
    set_run_style(r_end, font_size=11, bold=True)

    # 4. Thêm BẢNG ĐÁP ÁN 180 CÂU ở cuối file (giúp Azota nhận diện kép và giáo viên dễ đối soát)
    p_tbl_title = doc.add_paragraph()
    p_tbl_title.paragraph_format.space_before = Pt(12)
    p_tbl_title.paragraph_format.space_after = Pt(6)
    p_tbl_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_tbl_title = p_tbl_title.add_run("BẢNG ĐÁP ÁN 180 CÂU TRẮC NGHIỆM ÔN TẬP CHƯƠNG 1, 2, 3")
    set_run_style(r_tbl_title, font_size=12, bold=True, color=COLOR_PRIMARY)

    # Bảng đáp án gồm 10 cột, 36 hàng (18 cặp hàng: Hàng số câu và Hàng đáp án)
    # Cụ thể: 18 cặp hàng = 18 x 10 = 180 câu
    COLS = 10
    total_pairs = (len(all_questions) + COLS - 1) // COLS # 18 cặp hàng
    ans_table = doc.add_table(rows=total_pairs * 2, cols=COLS)
    ans_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ans_table, color="90A4AE", sz="4")

    for pair_idx in range(total_pairs):
        row_q = pair_idx * 2
        row_a = pair_idx * 2 + 1
        for col_idx in range(COLS):
            q_num = pair_idx * COLS + col_idx + 1
            if q_num <= len(all_questions):
                q_item = all_questions[q_num - 1]
                
                # Ô số câu
                cell_q = ans_table.cell(row_q, col_idx)
                cell_q.width = Cm(1.8)
                set_cell_background(cell_q, "E3F2FD")
                set_cell_margins(cell_q, top=40, bottom=40, left=40, right=40)
                p_q = cell_q.paragraphs[0]
                p_q.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_q = p_q.add_run(f"C{q_num}")
                set_run_style(r_q, font_size=9, bold=True, color=COLOR_PRIMARY)

                # Ô đáp án
                cell_a = ans_table.cell(row_a, col_idx)
                cell_a.width = Cm(1.8)
                set_cell_margins(cell_a, top=40, bottom=40, left=40, right=40)
                p_a = cell_a.paragraphs[0]
                p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_a = p_a.add_run(q_item['correct'])
                set_run_style(r_a, font_size=10, bold=True, color=COLOR_CORRECT)

    # 5. Lưu tài liệu
    def safe_save(d, path):
        try:
            d.save(path)
            print(f"   + Đã lưu thành công: {path}")
        except PermissionError:
            base, ext = os.path.splitext(path)
            alt = base + "_CapNhat" + ext
            d.save(alt)
            print(f"   [!] File đang mở trong Word, đã lưu sang: {alt}")

    safe_save(doc, OUT_FILE_1)
    safe_save(doc, OUT_FILE_2)

    print("=== HOÀN TẤT TẠO FILE GỘP 180 CÂU CHUẨN AZOTA! ===")

if __name__ == "__main__":
    build_merged_azota_doc()
