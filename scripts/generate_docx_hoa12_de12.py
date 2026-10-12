# -*- coding: utf-8 -*-
"""
Script tạo file Word (.docx) chuẩn "FORM ĐỀ KIỂM TRA" cho:
ĐỀ KIỂM TRA SỐ 12 - CHƯƠNG 4: POLYMER (HÓA HỌC 12)
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De12.docx (Bản Học Sinh)
2. De_Kiem_Tra_Hoa_12_De12_GiaoVien.docx (Bản Giáo Viên có lời giải chi tiết)
Tuân thủ 100% quy chuẩn: Header & Footer, Tab Stops (NO TABLE trong đề/lời giải), Sub/Superscript công thức hóa học.
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')
import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
IMG_DIR = r"c:\Antigravity_Thanh\San_Pham\images_crop_de12"

COLOR_HF_RED = RGBColor(211, 47, 47)      # Coral Red (#D32F2F)
COLOR_HF_GREEN_HEX = "8BC390"             # Sage Green (#8BC390)
COLOR_PRIMARY = RGBColor(46, 125, 50)     # Dark Green (#2E7D32)
COLOR_CORRECT = RGBColor(198, 40, 40)     # Red for correct answers (#C62828)
COLOR_SOLUTION = RGBColor(21, 101, 192)   # Blue for solutions (#1565C0)
COLOR_GRAY = RGBColor(97, 97, 97)

def add_page_number_field(p, font_name="Times New Roman", font_size=10, bold=True, color=COLOR_HF_RED):
    run_txt = p.add_run("Trang ")
    set_run_style(run_txt, font_name=font_name, font_size=font_size, bold=bold, color=color)
    
    r = p.add_run()
    set_run_style(r, font_name=font_name, font_size=font_size, bold=bold, color=color)
    fld1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instr1 = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
    fld2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fld3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    r._r.append(fld1); r._r.append(instr1); r._r.append(fld2); r._r.append(fld3)

def set_run_style(run, font_name="Times New Roman", font_size=12, bold=False, italic=False, subscript=False, superscript=False, color=None):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold
    run.italic = italic
    if subscript:
        run.font.subscript = True
    if superscript:
        run.font.superscript = True
    if color:
        run.font.color.rgb = color

def add_chem_run(p, text, base_font="Times New Roman", base_size=12, bold=False, italic=False, color=None):
    tokens = re.split(r'(\^\{[^}]+\}|\^[0-9\+\-]+|~\{[^}]+\}|~[0-9]+)', text)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith('^'):
            content = tok[2:-1] if tok.startswith('^{') else tok[1:]
            r = p.add_run(content)
            set_run_style(r, font_name=base_font, font_size=base_size, bold=bold, italic=italic, superscript=True, color=color)
        elif tok.startswith('~'):
            content = tok[2:-1] if tok.startswith('~{') else tok[1:]
            r = p.add_run(content)
            set_run_style(r, font_name=base_font, font_size=base_size, bold=bold, italic=italic, subscript=True, color=color)
        else:
            r = p.add_run(tok)
            set_run_style(r, font_name=base_font, font_size=base_size, bold=bold, italic=italic, color=color)

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(r'<w:tcBorders %s/>' % nsdecls('w'))
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = parse_xml(r'<%s %s w:val="%s" w:sz="%s" w:space="0" w:color="%s"/>' %
                                (tag, nsdecls('w'), edge_data.get('val', 'single'),
                                 edge_data.get('sz', '4'), edge_data.get('color', 'auto')))
            tcBorders.append(element)
        else:
            tag = 'w:{}'.format(edge)
            element = parse_xml(r'<%s %s w:val="none"/>' % (tag, nsdecls('w')))
            tcBorders.append(element)
    tcPr.append(tcBorders)

def set_cell_margins(cell, top=50, bottom=50, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(r'<w:tcMar %s><w:top w:w="%d" w:type="dxa"/><w:bottom w:w="%d" w:type="dxa"/><w:left w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/></w:tcMar>' % (nsdecls('w'), top, bottom, left, right))
    tcPr.append(tcMar)

def setup_page_and_header_footer(doc):
    for sec in doc.sections:
        sec.top_margin = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
        sec.left_margin = Cm(1.5)
        sec.right_margin = Cm(1.5)
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.header_distance = Cm(0.6)
        sec.footer_distance = Cm(0.6)
        
        # --- HEADER ---
        hdr = sec.header
        hdr_p = hdr.paragraphs[0]
        hdr_p.text = ""
        hdr_p.paragraph_format.space_before = Pt(0)
        hdr_p.paragraph_format.space_after = Pt(0)
        
        tbl_hdr = hdr.add_table(rows=1, cols=2, width=Cm(18.0))
        tbl_hdr.autofit = False
        tbl_hdr.columns[0].width = Cm(9.0)
        tbl_hdr.columns[1].width = Cm(9.0)
        
        c0 = tbl_hdr.cell(0, 0)
        c1 = tbl_hdr.cell(0, 1)
        c0.width = Cm(9.0)
        c1.width = Cm(9.0)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(2)
        r0 = p0.add_run("Tài liệu lưu hành nội bộ")
        set_run_style(r0, font_name="Times New Roman", font_size=10, italic=True, color=COLOR_HF_RED)
        
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(2)
        r1a = p1.add_run("Thầy ")
        set_run_style(r1a, font_name="Times New Roman", font_size=10, color=COLOR_HF_RED)
        r1b = p1.add_run("TRẦN VĂN THẠNH")
        set_run_style(r1b, font_name="Times New Roman", font_size=10, bold=True, color=COLOR_HF_RED)
        r1c = p1.add_run(" - 0777.470.803")
        set_run_style(r1c, font_name="Times New Roman", font_size=10, color=COLOR_HF_RED)
        
        set_cell_border(c0, bottom={'val': 'single', 'sz': '6', 'color': COLOR_HF_GREEN_HEX})
        set_cell_border(c1, bottom={'val': 'single', 'sz': '6', 'color': COLOR_HF_GREEN_HEX})
        set_cell_margins(c0, top=0, bottom=40, left=0, right=0)
        set_cell_margins(c1, top=0, bottom=40, left=0, right=0)
        
        # --- FOOTER ---
        ftr = sec.footer
        ftr_p = ftr.paragraphs[0]
        ftr_p.text = ""
        ftr_p.paragraph_format.space_before = Pt(0)
        ftr_p.paragraph_format.space_after = Pt(0)
        
        tbl_ftr = ftr.add_table(rows=1, cols=2, width=Cm(18.0))
        tbl_ftr.autofit = False
        tbl_ftr.columns[0].width = Cm(13.5)
        tbl_ftr.columns[1].width = Cm(4.5)
        
        fc0 = tbl_ftr.cell(0, 0)
        fc1 = tbl_ftr.cell(0, 1)
        fc0.width = Cm(13.5)
        fc1.width = Cm(4.5)
        
        fp0 = fc0.paragraphs[0]
        fp0.paragraph_format.space_before = Pt(2)
        fp0.paragraph_format.space_after = Pt(0)
        fr0 = fp0.add_run("CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế   |   CS2: 24 Đặng Thái Thân, TP Huế")
        set_run_style(fr0, font_name="Times New Roman", font_size=9.5, italic=True, color=COLOR_HF_RED)
        
        fp1 = fc1.paragraphs[0]
        fp1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        fp1.paragraph_format.space_before = Pt(2)
        fp1.paragraph_format.space_after = Pt(0)
        add_page_number_field(fp1, font_name="Times New Roman", font_size=10, bold=True, color=COLOR_HF_RED)
        
        set_cell_border(fc0, top={'val': 'single', 'sz': '6', 'color': COLOR_HF_GREEN_HEX})
        set_cell_border(fc1, top={'val': 'single', 'sz': '6', 'color': COLOR_HF_GREEN_HEX})
        set_cell_margins(fc0, top=40, bottom=0, left=0, right=0)
        set_cell_margins(fc1, top=40, bottom=0, left=0, right=0)

def add_exam_title_block(doc, is_teacher=False):
    lines_info = [
        ("SỞ GD&ĐT THỪA THIÊN HUẾ", "ĐỀ KIỂM TRA CHƯƠNG 4: POLYMER", True, True),
        ("TRƯỜNG THPT HAI BÀ TRƯNG", "MÔN: HÓA HỌC 12 - NĂM HỌC 2026 - 2027", False, True),
        ("ĐỀ CHÍNH THỨC", "Thời gian làm bài: 50 phút (không kể thời gian phát đề)", True, False, False, True),
        ("(Đề thi gồm 04 trang)", "MÃ ĐỀ THI: 112", False, True, True, False)
    ]
    
    for item in lines_info:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.05
        p.paragraph_format.tab_stops.add_tab_stop(Cm(18.0), WD_TAB_ALIGNMENT.RIGHT)
        
        r_l = p.add_run(item[0])
        set_run_style(r_l, font_name="Times New Roman", font_size=11, bold=item[2], italic=item[4] if len(item)>4 else False)
        
        p.add_run("\t")
        
        r_r = p.add_run(item[1])
        set_run_style(r_r, font_name="Times New Roman", font_size=11, bold=item[3], italic=item[5] if len(item)>5 else False)

    if is_teacher:
        p_note = doc.add_paragraph()
        p_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_note.paragraph_format.space_before = Pt(2)
        p_note.paragraph_format.space_after = Pt(2)
        r_note = p_note.add_run("(BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)")
        set_run_style(r_note, font_name="Times New Roman", font_size=11.5, bold=True, color=COLOR_HF_RED)

    # Student info line
    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(3)
    p_info.paragraph_format.space_after = Pt(1)
    r_info = p_info.add_run("Họ, tên thí sinh: .......................................................................... Lớp: ............. SBD: .............................")
    set_run_style(r_info, font_name="Times New Roman", font_size=11, italic=True)

    # Atomic mass note
    p_mass = doc.add_paragraph()
    p_mass.paragraph_format.space_before = Pt(0)
    p_mass.paragraph_format.space_after = Pt(4)
    r_mass = p_mass.add_run("(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108; Ba = 137).")
    set_run_style(r_mass, font_name="Times New Roman", font_size=9.5, italic=True, color=COLOR_GRAY)

def add_part_heading(doc, part_title, part_instruction):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    
    r_title = p.add_run(part_title)
    set_run_style(r_title, font_name="Times New Roman", font_size=11.5, bold=True, color=COLOR_PRIMARY)
    
    if part_instruction:
        r_inst = p.add_run(" " + part_instruction)
        set_run_style(r_inst, font_name="Times New Roman", font_size=11, italic=True)

def add_solution_paragraph(doc, solution_text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.left_indent = Cm(0.6)
    
    r_label = p.add_run("Lời giải chi tiết: ")
    set_run_style(r_label, font_name="Times New Roman", font_size=10.5, bold=True, color=COLOR_SOLUTION)
    
    add_chem_run(p, solution_text, base_font="Times New Roman", base_size=10.5, italic=True)

# ----------------- DỮ LIỆU ĐỀ KIỂM TRA SỐ 12 -----------------

MCQ_QUESTIONS = [
    {
        "num": 1,
        "text": "Tên gọi của polymer có công thức cho dưới đây là",
        "image": "cau1_buna_n.png",
        "image_w": Cm(9.0),
        "options": ["A. Cao su buna.", "B. Cao su buna-S.", "C. Cao su buna-N.", "D. Cao su isoprene."],
        "correct": "C",
        "sol": "Chọn C. Phân tử gồm các mắt xích buta-1,3-diene (-CH~2~-CH=CH-CH~2~-) và acrylonitrile (-CH(CN)-CH~2~-) liên kết với nhau, đây là sản phẩm của phản ứng đồng trùng hợp buta-1,3-diene với acrylonitrile tạo thành cao su buna-N (cao su nitrile, NBR)."
    },
    {
        "num": 2,
        "text": "(SBT – KNTT). Poly(vinyl chloride) (PVC) là một loại chất dẻo phổ biến, được sử dụng làm vỏ dây điện, ống dẫn nước thải, áo mưa, vải giả da,... PVC có công thức cấu tạo là",
        "image": "cau2_pvc_options.png",
        "image_w": Cm(12.0),
        "options": [
            "A. [-CH~2~-CH~2~-]~n~.",
            "B. [-CH~2~-CH(CH~3~)-]~n~.",
            "C. [-CH~2~-CH(C~6~H~5~)-]~n~.",
            "D. [-CH~2~-CH(Cl)-]~n~."
        ],
        "correct": "D",
        "sol": "Chọn D. Poly(vinyl chloride) (PVC) được điều chế bằng phản ứng trùng hợp vinyl chloride (CH~2~=CHCl), có công thức cấu tạo là [-CH~2~-CH(Cl)-]~n~."
    },
    {
        "num": 3,
        "text": "Polymer nào sau đây có thể cộng với hydrogen chloride?",
        "options": ["A. Polyacrylonitrile.", "B. Polyethylene.", "C. Polyisoprene.", "D. Tinh bột."],
        "correct": "C",
        "sol": "Chọn C. Phân tử polyisoprene còn chứa các liên kết đôi C=C trong mắt xích ([-CH~2~-C(CH~3~)=CH-CH~2~-]~n~) nên có phản ứng cộng với hydrogen chloride (HCl)."
    },
    {
        "num": 4,
        "text": "Trong các phản ứng giữa các cặp chất dưới đây, phản ứng nào làm giảm mạch polymer?",
        "options": [
            "A. Poly(vinyl chloride) + Cl~2~.",
            "B. Cao su thiên nhiên + HCl.",
            "C. Tinh bột + H~2~O (t^o^, H^+^).",
            "D. Poly(vinyl acetate) + NaOH."
        ],
        "correct": "C",
        "sol": "Chọn C. Phản ứng thủy phân tinh bột trong môi trường acid đun nóng (H~2~O, H^+^, t^o^) cắt đứt các liên kết glycosidic, phân cắt các phân tử polysaccharide mạch dài thành các phân tử dextrin, maltose và cuối cùng là glucose có mạch ngắn hơn, do đó làm giảm mạch polymer."
    },
    {
        "num": 5,
        "text": "(SBT – KNTT). Polymer nào sau đây không thuộc loại chất dẻo?",
        "options": [
            "A. Poly(methyl methacrylate).",
            "B. Poly(vinyl chloride).",
            "C. Polystyrene.",
            "D. Polybuta-1,3-diene."
        ],
        "correct": "D",
        "sol": "Chọn D. Polybuta-1,3-diene là polymer có tính đàn hồi cao, thuộc loại vật liệu cao su (cao su buna), không phải chất dẻo."
    },
    {
        "num": 6,
        "text": "(SBT – KNTT). LDPE là một chất dẻo dễ tạo màng, có tính dai bền nên được sử dụng làm túi nylon, màng bọc, bao gói thực phẩm. Trên các bao bì làm từ LDPE thường được in kí hiệu như hình bên. LDPE được tổng hợp từ monomer nào sau đây?",
        "image": "cau6_ldpe_symbol.png",
        "image_w": Cm(2.0),
        "options": [
            "A. CH~2~=CH~2~.",
            "B. CH~2~=CH-CH~3~.",
            "C. CH~2~=CH-C~6~H~5~.",
            "D. CH~2~=CHCl."
        ],
        "correct": "A",
        "sol": "Chọn A. LDPE (Low-Density Polyethylene) là polyethylene mật độ thấp, được tổng hợp bằng phản ứng trùng hợp monomer ethylene (CH~2~=CH~2~) ở áp suất cao."
    },
    {
        "num": 7,
        "text": "(SBT – CTST). Nylon-6,6 thuộc loại tơ nào sau đây?",
        "options": ["A. Tơ polyester.", "B. Tơ bán tổng hợp.", "C. Tơ thiên nhiên.", "D. Tơ tổng hợp."],
        "correct": "D",
        "sol": "Chọn D. Nylon-6,6 được điều chế bằng phản ứng trùng ngưng giữa hexamethylenediamine và adipic acid, là các chất tổng hợp hoàn toàn từ công nghiệp hóa chất, nên thuộc loại tơ tổng hợp (polyamide)."
    },
    {
        "num": 8,
        "text": "(SBT – CTST). Loại polymer thiên nhiên nào có thành phần chính là cellulose?",
        "options": ["A. Bông.", "B. Tơ tằm.", "C. Len.", "D. Cao su thiên nhiên."],
        "correct": "A",
        "sol": "Chọn A. Sợi bông có thành phần chính là cellulose (chiếm khoảng 95 - 98%). Tơ tằm và len có thành phần chính là protein; mủ cao su thiên nhiên là polyisoprene."
    },
    {
        "num": 9,
        "text": "(SBT – KNTT). Cao su buna-N (hay còn gọi là cao su nitrile, có kí hiệu là NBR) là loại cao su tổng hợp có khả năng chịu dầu mỡ tốt nên được dùng làm ống dẫn nhiên liệu, gioăng phớt làm kín trong các máy móc. Thực hiện phản ứng trùng hợp các chất nào dưới đây thu được sản phẩm là cao su buna-N?",
        "options": [
            "A. CH~2~=CH-CH=CH~2~ và C~6~H~5~CH=CH~2~.",
            "B. CH~2~=C(CH~3~)-CH=CH~2~ và CH~2~=CH-CN.",
            "C. CH~2~=CH-CH=CH~2~ và N~2~.",
            "D. CH~2~=CH-CH=CH~2~ và CH~2~=CH-CN."
        ],
        "correct": "D",
        "sol": "Chọn D. Cao su buna-N được điều chế bằng phản ứng đồng trùng hợp buta-1,3-diene (CH~2~=CH-CH=CH~2~) và acrylonitrile (vinyl cyanide, CH~2~=CH-CN)."
    },
    {
        "num": 10,
        "text": "Nhận xét về tính chất vật lí chung của polymer nào dưới đây không đúng?",
        "options": [
            "A. Hầu hết là những chất rắn, không bay hơi.",
            "B. Đa số nóng chảy ở một khoảng nhiệt độ rộng, hoặc không nóng chảy mà bị phân hủy khi đun nóng.",
            "C. Đa số không tan trong các dung môi thông thường, một số tan trong dung môi thích hợp tạo dung dịch nhớt.",
            "D. Hầu hết polymer đều đồng thời có tính dẻo, tính đàn hồi và có thể kéo thành sợi dai, bền."
        ],
        "correct": "D",
        "sol": "Chọn D. Phát biểu D không đúng vì mỗi polymer có những đặc tính cơ học riêng biệt: chất dẻo có tính dẻo, cao su có tính đàn hồi, tơ có thể kéo sợi dai bền,... Hầu hết polymer không đồng thời có cả 3 đặc tính trên."
    },
    {
        "num": 11,
        "text": "Dãy nào sau đây chỉ chứa các polymer thiên nhiên?",
        "options": [
            "A. Poly(vinyl chloride), tinh bột, cellulose.",
            "B. Protein, tinh bột, polyethylene.",
            "C. Protein, cellulose.",
            "D. Protein, tinh bột, cellulose."
        ],
        "correct": "D",
        "sol": "Chọn D. Protein, tinh bột và cellulose đều là các polymer có sẵn trong tự nhiên (thực vật, động vật) nên đều là polymer thiên nhiên. Các chất PVC, polyethylene là polymer tổng hợp."
    },
    {
        "num": 12,
        "text": "Cho dãy gồm các polymer: (1) polybuta-1,3-diene, (2) polyacrylonitrile, (3) poly(vinyl chloride), (4) poly(hexamethylene adipamide). Polymer bị thuỷ phân trong môi trường acid là",
        "options": ["A. (1).", "B. (3).", "C. (4).", "D. (2)."],
        "correct": "C",
        "sol": "Chọn C. (4) Poly(hexamethylene adipamide) (nylon-6,6) chứa các liên kết amide -CO-NH- trong mạch, dễ bị thủy phân trong cả môi trường acid và kiềm."
    },
    {
        "num": 13,
        "text": "Nhận xét nào sau đây đúng?",
        "options": [
            "A. Đa số các polymer dễ hòa tan trong các dung môi thông thường.",
            "B. Đa số các polymer không có nhiệt độ nóng chảy xác định.",
            "C. Các polymer đều bền vững dưới tác động của acid, base.",
            "D. Các polymer dễ bay hơi."
        ],
        "correct": "B",
        "sol": "Chọn B. Đa số polymer gồm hỗn hợp các đại phân tử có khối lượng phân tử và độ dài mạch khác nhau nên không có nhiệt độ nóng chảy xác định mà nóng chảy trong một khoảng nhiệt độ khá rộng."
    },
    {
        "num": 14,
        "text": "Cho các polymer sau: polyethylene, poly(methyl methacrylate), poly(vinyl chloride), polyacrylonitrile. Số polymer điều chế được bằng phản ứng trùng hợp là",
        "options": ["A. 3.", "B. 4.", "C. 1.", "D. 2."],
        "correct": "B",
        "sol": "Chọn B. Cả 4 polymer trên đều được điều chế bằng phản ứng trùng hợp monomer chứa liên kết đôi C=C:\n• Polyethylene từ CH~2~=CH~2~.\n• Poly(methyl methacrylate) từ CH~2~=C(CH~3~)-COOCH~3~.\n• Poly(vinyl chloride) từ CH~2~=CHCl.\n• Polyacrylonitrile từ CH~2~=CH-CN."
    },
    {
        "num": 15,
        "text": "Phát biểu nào sau đây sai?",
        "options": [
            "A. Tơ nitron thuộc loại tơ tổng hợp.",
            "B. Tơ nylon-6,6 thuộc loại tơ tổng hợp.",
            "C. Tơ visco thuộc loại tơ thiên nhiên.",
            "D. Tơ acetate thuộc loại tơ bán tổng hợp."
        ],
        "correct": "C",
        "sol": "Chọn C. Tơ visco là tơ bán tổng hợp (tơ nhân tạo) được điều chế từ cellulose thiên nhiên qua phản ứng hóa học với CS~2~ và kiềm, không phải tơ thiên nhiên."
    },
    {
        "num": 16,
        "text": "Phát biểu nào sau đây sai?",
        "options": [
            "A. Bản chất cấu tạo hoá học của tơ tằm và len là protein; của sợi bông là cellulose.",
            "B. Bản chất cấu tạo hoá học của tơ nylon là polyamide.",
            "C. Quần áo nylon, len, tơ tằm không nên giặt với xà phòng có độ kiềm cao.",
            "D. Tơ nylon, tơ tằm, len rất bền vững với nhiệt."
        ],
        "correct": "D",
        "sol": "Chọn D. Tơ nylon, tơ tằm, len đều kém bền với nhiệt do chứa các liên kết amide/peptide dễ bị phá vỡ hoặc co xoắn ở nhiệt độ cao, khi là (ủi) quá nóng sẽ làm co rúm và hư hỏng sợi vải."
    },
    {
        "num": 17,
        "text": "(SBT – CTST). Trùng ngưng ε-aminocaproic acid thu được m kg capron và 12,6 kg H~2~O với hiệu suất phản ứng 90%. Giá trị của m là",
        "options": ["A. 71,19.", "B. 79,10.", "C. 91,70.", "D. 90,40."],
        "correct": "B",
        "sol": "Chọn B.\nPhương trình phản ứng trùng ngưng:\nn H~2~N-(CH~2~)~5~-COOH -> [-NH-(CH~2~)~5~-CO-]~n~ + n H~2~O\nTheo tỉ lệ phản ứng, cứ n mol H~2~O sinh ra tương ứng với 1 mol mắt xích capron (113n gam).\nSố mol H~2~O thực tế thu được = 12,6 / 18 = 0,7 kmol.\n=> Khối lượng capron thực tế thu được:\nm = 0,7 * 113 = 79,10 kg.\n(Lượng H~2~O đề bài cho là lượng thực tế thu được, nên lượng polymer thu được tương ứng trực tiếp theo tỉ lệ stoichiometry là 79,10 kg)."
    },
    {
        "num": 18,
        "text": "Keo dán là vật liệu được sử dụng rộng rãi trong đời sống. Cho các phát biểu sau về keo dán:\n(a) Keo dán là vật liệu có khả năng kết dính bề mặt của hai vật liệu rắn với nhau mà không làm biến đổi bản chất các vật liệu được kết dính.\n(b) Nhựa vá săm là dung dịch dạng keo của chất dẻo trong dung môi hữu cơ như toluene, xylene, thường được dùng để vá chỗ thủng của săm xe.\n(c) Keo dán epoxy còn gọi là keo dán một thành phần.\n(d) Thành phần chính của keo epoxy là hợp chất chứa hai nhóm epoxy ở hai đầu. Thành phần thứ hai là chất đóng rắn, thường là các polyamine.\n(e) Ưu điểm của keo epoxy là độ kết dính rất cao, chịu nhiệt, chịu nước, chịu dung môi, chịu lực tốt, rất dễ sử dụng. Keo epoxy thường được dùng để dán các kim loại, nhựa, bê tông, các vật bằng kính, sứ, gỗ, đồ gốm.\nSố phát biểu đúng là",
        "options": ["A. 2.", "B. 3.", "C. 5.", "D. 4."],
        "correct": "B",
        "sol": "Chọn B. Các phát biểu đúng là (a), (d), (e) (tổng cộng 3 phát biểu).\n• (a) ĐÚNG theo định nghĩa keo dán.\n• (b) SAI vì nhựa vá săm là dung dịch keo của cao su trong dung môi hữu cơ, không phải chất dẻo.\n• (c) SAI vì keo dán epoxy là keo dán 2 thành phần (keo AB).\n• (d) ĐÚNG vì gồm hợp chất chứa 2 nhóm epoxy và chất đóng rắn polyamine.\n• (e) ĐÚNG về các đặc tính ưu việt và ứng dụng phong phú của keo epoxy."
    }
]

TF_QUESTIONS = [
    {
        "num": 1,
        "intro": "(SBT – CTST). Tơ capron và polyacrylonitrile có công thức cấu tạo như hình dưới:",
        "image": "p2_cau1_polymers.png",
        "items": [
            ("a", "Cả hai loại đều dùng để chế tạo tơ.", True, "Đúng. Capron dùng chế tạo tơ capron (nylon-6), polyacrylonitrile dùng chế tạo tơ nitron (olon)."),
            ("b", "Polyacrylonitrile được điều chế bằng phản ứng trùng ngưng.", False, "Sai. Polyacrylonitrile được điều chế bằng phản ứng trùng hợp monomer acrylonitrile (vinyl cyanide, CH~2~=CH-CN)."),
            ("c", "Capron thuộc loại polypeptide.", False, "Sai. Capron thuộc loại polyamide (polycaproamide), được tạo từ một phân tử amino acid riêng lẻ ε-aminocaproic acid, không phải là polypeptide (polypeptide gồm chuỗi nối từ các α-amino acid)."),
            ("d", "Hai polymer đều có phản ứng phân huỷ mạch trong môi trường kiềm.", False, "Sai. Capron có liên kết amide -CO-NH- bị thủy phân phân hủy mạch trong môi trường kiềm, nhưng polyacrylonitrile có mạch carbon-carbon bền vững không bị phân hủy mạch chính trong kiềm.")
        ]
    },
    {
        "num": 2,
        "intro": "(SBT - CD). Năm 1839, Charles Goodyear đã được cấp bằng sáng chế cho phát minh về quy trình hoá học để chế tạo ra cao su lưu hoá - một loại cao su có cấu trúc đặc biệt, bền cơ học, chịu được sự ma sát, va chạm, đàn hồi tốt và có thể đúc được. Những phát biểu nào sau đây là đúng?",
        "items": [
            ("a", "Cao su lưu hoá còn có tên gọi là cao su buna-S.", False, "Sai. Cao su buna-S là sản phẩm đồng trùng hợp buta-1,3-diene với styrene, khác hoàn toàn với cao su lưu hóa (được chế tạo bằng cách đun nóng cao su với lưu huỳnh)."),
            ("b", "Bản chất của việc lưu hoá cao su là tạo ra cầu nối disulfide -S-S- giữa các mạch cao su nên cao su lưu hoá có tính chất cơ lí nổi trội hơn.", True, "Đúng. Cầu nối disulfide -S-S- nối các mạch polymer riêng rẽ lại tạo cấu trúc không gian ba chiều bền vững."),
            ("c", "Trong mủ cao su thiên nhiên, polymer có tính đàn hồi là polyisoprene.", True, "Đúng. Thành phần chính tạo tính đàn hồi của mủ cao su thiên nhiên là cis-polyisoprene."),
            ("d", "Cao su lưu hoá có cấu trúc mạng không gian nên bền hơn cao su chưa lưu hoá.", True, "Đúng. Cấu trúc mạng không gian giúp cao su lưu hóa bền cơ học, lâu mòn, đàn hồi tốt hơn và chịu nhiệt cao hơn.")
        ]
    },
    {
        "num": 3,
        "intro": "Cây cao su là loại cây công nghiệp có giá trị kinh tế lớn, được đưa vào trồng ở nước ta từ cuối thế kỉ 19. Chất lỏng thu được từ cây cao su giống như nhựa cây (gọi là mủ cao su) là nguyên liệu để sản xuất cao su tự nhiên.",
        "items": [
            ("a", "Thành phần chính của cao su thiên nhiên là polyisoprene có cấu hình dạng cis.", True, "Đúng. Cao su thiên nhiên là cis-1,4-polyisoprene."),
            ("b", "Cao su thiên nhiên có tính đàn hồi, không có tính dẫn điện, không thấm khí và nước, không tan trong nước, nhưng tan nhiều trong xăng, benzene...", True, "Đúng. Đây là những tính chất vật lí đặc trưng của cao su thiên nhiên."),
            ("c", "Cao su thiên nhiên có tính đàn hồi chỉ tồn tại trong khoảng nhiệt độ hẹp. Để cải thiện tính chất này người ta xử lý bằng cách trộn cao su thiên nhiên với styrene gọi là quá trình lưu hóa cao su.", False, "Sai. Quá trình lưu hóa cao su là đun nóng cao su thiên nhiên với lưu huỳnh (S), không phải styrene."),
            ("d", "Mắt xích của cao su thiên nhiên được viết thu gọn là C~5~H~8~.", True, "Đúng. Mắt xích của polyisoprene là [-CH~2~-C(CH~3~)=CH-CH~2~-], có công thức thực nghiệm thu gọn là C~5~H~8~.")
        ]
    },
    {
        "num": 4,
        "intro": "Bản chất kết dính của keo dán là tạo ra lớp màng mỏng bám chắc vào hai mảnh vật liệu để giúp chúng kết dính lại với nhau.",
        "items": [
            ("a", "Keo dán poly(urea – formaldehyde) được điều chế từ phản ứng trùng ngưng.", True, "Đúng. Keo poly(urea – formaldehyde) được tổng hợp bằng phản ứng trùng ngưng giữa urea (NH~2~-CO-NH~2~) và formaldehyde (HCHO)."),
            ("b", "Keo dán poly(urea – formaldehyde) gồm phần chính là poly(urea – formaldehyde) và thêm phần đóng rắn như oxalic acid, ammonium chloride... để tạo polymer có cấu trúc mạch phân nhánh.", False, "Sai. Chất đóng rắn tạo mạng lưới polymer không gian (mạng lưới 3 chiều), không phải mạch phân nhánh."),
            ("c", "Keo dán epoxy có polymer mạng không gian với độ kết dính cao dùng để gắn kết hai bề mặt vật cần dán lại như kim loại, thủy tinh, gỗ...", True, "Đúng. Nhựa epoxy kết hợp với chất đóng rắn tạo polymer mạng không gian có độ bám dính cực kỳ cao và chịu lực rất tốt."),
            ("d", "Keo dán tổng hợp thông dụng thường gặp gồm keo dán urea – formaldehyde, keo dán epoxy và nhựa vá săm.", True, "Đúng. Đây là ba loại keo dán tổng hợp rất phổ biến trong đời sống và kỹ thuật sản xuất.")
        ]
    }
]

SA_QUESTIONS = [
    {
        "num": 1,
        "text": "Cho dãy gồm các nguyên liệu: (1) ε-aminocaproic acid, (2) acrylonitrile, (3) ω-aminoenanthic acid, (4) ethylene glycol và terephthalic acid, (5) hexamethylenediamine và adipic acid. Số nguyên liệu (hoặc cặp nguyên liệu) được dùng để tiến hành phản ứng trùng ngưng tạo thành tơ là",
        "ans": "4",
        "sol": "Có 4 nguyên liệu/cặp nguyên liệu dùng để trùng ngưng tạo tơ:\n• (1) ε-aminocaproic acid -> tơ capron (nylon-6).\n• (3) ω-aminoenanthic acid -> tơ enang (nylon-7).\n• (4) ethylene glycol và terephthalic acid -> tơ lapsan (PET).\n• (5) hexamethylenediamine và adipic acid -> tơ nylon-6,6.\n(Nguyên liệu (2) acrylonitrile dùng điều chế tơ nitron bằng phản ứng trùng hợp)."
    },
    {
        "num": 2,
        "text": "Cho các polymer: amylose (1); cao su isoprene (2); cellulose acetate (3); thủy tinh hữu cơ (4); tơ tằm (5). Số lượng polymer thiên nhiên là",
        "ans": "2",
        "sol": "Có 2 polymer thiên nhiên là: amylose (1) và tơ tằm (5).\n(Cao su isoprene trong đề là cao su tổng hợp polyisoprene; cellulose acetate là polymer bán tổng hợp; thủy tinh hữu cơ là polymer tổng hợp)."
    },
    {
        "num": 3,
        "text": "Phân tử khối của một đoạn mạch tơ nylon-6,6 là 56 500 amu. Số lượng mắt xích trong đoạn mạch nylon-6,6 nêu trên là",
        "ans": "250",
        "sol": "Mắt xích tơ nylon-6,6 có công thức [-NH-(CH~2~)~6~-NH-CO-(CH~2~)~4~-CO-] với phân tử khối M = 116 + 146 - 2 * 18 = 226 amu.\nSố lượng mắt xích n = 56500 / 226 = 250."
    },
    {
        "num": 4,
        "text": "Polystyrene (PS) là loại nhựa dùng chế tạo hộp xốp đựng thức ăn. Hãy tính hệ số polymer hóa của loại nhựa này khi biết phân tử khối của nó bằng 104 000 amu.",
        "ans": "1000",
        "sol": "Polystyrene có công thức [-CH~2~-CH(C~6~H~5~)-]~n~ với phân tử khối một mắt xích là M = 104 amu.\nHệ số polymer hóa n = 104000 / 104 = 1000."
    },
    {
        "num": 5,
        "text": "Trong công nghiệp, để điều chế cao su buna người ta có thể đi từ nguyên liệu khí ethylene thu được từ dầu mỏ theo sơ đồ sau:\n2 C~2~H~4~ -> 2 C~2~H~5~OH -> C~4~H~6~ -> [-CH~2~-CH=CH-CH~2~-]~n~ (cao su buna)\nTính thể tích khí ethylene (đo ở 25 oC và 1 bar) cần lấy để điều chế được 1 tấn cao su buna theo sơ đồ trên. Giả sử hiệu suất phản ứng của mỗi quá trình lần lượt là 65%, 50% và 70% (kết quả làm tròn đến hàng đơn vị theo đơn vị m^3).",
        "ans": "4036",
        "sol": "• Khối lượng cao su buna: m = 1 tấn = 1000 kg.\n  => n_caosu = 1000 / 54 kmol.\n• Theo sơ đồ bảo toàn carbon: 2 C~2~H~4~ tạo 1 mắt xích C~4~H~6~ cao su buna.\n  => n_ethylene (lý thuyết) = 2 * (1000 / 54) = 2000 / 54 kmol.\n• Hiệu suất toàn bộ quá trình: H = 65% * 50% * 70% = 0,2275 (22,75%).\n• Ở 25 oC, 1 bar (điều kiện chuẩn mới): V_mol = 24,79 m^3/kmol.\n  => V_ethylene (thực tế) = (2000 / (54 * 0,2275)) * 24,79 ≈ 4035,82 ≈ 4036 m^3."
    },
    {
        "num": 6,
        "text": "Caprolactam được tổng hợp từ cuối thế kỉ XIX. Hiện nay, nhu cầu sản xuất caprolactam trên thế giới khoảng 10 triệu tấn/năm; 90% trong đó dùng để tổng hợp tơ capron. Trong công nghiệp, caprolactam được điều chế theo sơ đồ sau:\nĐể sản xuất 10 triệu tấn caprolactam, cần sử dụng bao nhiêu triệu tấn cyclohexanone (giả sử hiệu suất trung bình của cả quá trình trên là 60%)? (Kết quả làm tròn đến chữ số hàng phần mười).",
        "image": "p3_cau6_caprolactam_scheme.png",
        "image_w": Cm(12.0),
        "ans": "14,5",
        "sol": "• Sơ đồ chuyển hóa tổng quát:\n  Cyclohexanone (C~6~H~10~O, M = 98) -> Caprolactam (C~6~H~11~NO, M = 113)\n• Khối lượng cyclohexanone cần lấy:\n  m = (10 / 113) * 98 * (100 / 60) ≈ 14,454 ≈ 14,5 triệu tấn."
    }
]

def generate_exam_docx(is_teacher=False):
    filename = "De_Kiem_Tra_Hoa_12_De12_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_12_De12.docx"
    filepath = os.path.join(OUT_DIR, filename)
    
    doc = docx.Document()
    setup_page_and_header_footer(doc)
    add_exam_title_block(doc, is_teacher=is_teacher)
    
    # --- PHẦN I ---
    add_part_heading(doc, "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.", "Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")
    
    for q in MCQ_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(1)
        p_q.paragraph_format.line_spacing = 1.15
        
        r_qnum = p_q.add_run(f"Câu {q['num']}. ")
        set_run_style(r_qnum, font_name="Times New Roman", font_size=11, bold=True)
        
        add_chem_run(p_q, q["text"], base_font="Times New Roman", base_size=11)
        
        if q.get("image"):
            img_path = os.path.join(IMG_DIR, q["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(1)
                p_img.paragraph_format.space_after = Pt(2)
                run_img = p_img.add_run()
                w = q.get("image_w", Cm(8.0))
                run_img.add_picture(img_path, width=w)

        # Format options using Tabs (NO TABLE)
        opts = q["options"]
        max_opt_len = max(len(opt) for opt in opts)
        
        if max_opt_len < 22:
            # 1 dòng 4 lựa chọn (Tab stops 4.5, 9.0, 13.5 cm)
            p_opt = doc.add_paragraph()
            p_opt.paragraph_format.space_before = Pt(1)
            p_opt.paragraph_format.space_after = Pt(2)
            p_opt.paragraph_format.line_spacing = 1.15
            p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(4.5), WD_TAB_ALIGNMENT.LEFT)
            p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
            p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(13.5), WD_TAB_ALIGNMENT.LEFT)
            
            for i, opt in enumerate(opts):
                is_correct = (opt.startswith(q["correct"] + ".")) and is_teacher
                color = COLOR_CORRECT if is_correct else None
                
                prefix = opt[:3]
                body = opt[3:]
                r_pre = p_opt.add_run(prefix)
                set_run_style(r_pre, font_name="Times New Roman", font_size=11, bold=True, color=color)
                add_chem_run(p_opt, body, base_font="Times New Roman", base_size=11, bold=is_correct, color=color)
                
                if i < 3:
                    p_opt.add_run("\t")
        elif max_opt_len < 46:
            # 2 dòng 2 lựa chọn (Tab stop 9.0 cm)
            for row in [(0, 1), (2, 3)]:
                p_opt = doc.add_paragraph()
                p_opt.paragraph_format.space_before = Pt(1)
                p_opt.paragraph_format.space_after = Pt(1.5)
                p_opt.paragraph_format.line_spacing = 1.15
                p_opt.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
                
                for idx_in_row, opt_idx in enumerate(row):
                    opt = opts[opt_idx]
                    is_correct = (opt.startswith(q["correct"] + ".")) and is_teacher
                    color = COLOR_CORRECT if is_correct else None
                    
                    prefix = opt[:3]
                    body = opt[3:]
                    r_pre = p_opt.add_run(prefix)
                    set_run_style(r_pre, font_name="Times New Roman", font_size=11, bold=True, color=color)
                    add_chem_run(p_opt, body, base_font="Times New Roman", base_size=11, bold=is_correct, color=color)
                    
                    if idx_in_row == 0:
                        p_opt.add_run("\t")
        else:
            # 4 dòng riêng biệt
            for opt in opts:
                p_opt = doc.add_paragraph()
                p_opt.paragraph_format.space_before = Pt(1)
                p_opt.paragraph_format.space_after = Pt(1.5)
                p_opt.paragraph_format.line_spacing = 1.15
                
                is_correct = (opt.startswith(q["correct"] + ".")) and is_teacher
                color = COLOR_CORRECT if is_correct else None
                
                prefix = opt[:3]
                body = opt[3:]
                r_pre = p_opt.add_run(prefix)
                set_run_style(r_pre, font_name="Times New Roman", font_size=11, bold=True, color=color)
                add_chem_run(p_opt, body, base_font="Times New Roman", base_size=11, bold=is_correct, color=color)

        if is_teacher and q.get("sol"):
            add_solution_paragraph(doc, q["sol"])

    # --- PHẦN II ---
    add_part_heading(doc, "PHẦN II. Câu trắc nghiệm đúng sai.", "Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")
    
    for q in TF_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        
        r_qnum = p_q.add_run(f"Câu {q['num']}. ")
        set_run_style(r_qnum, font_name="Times New Roman", font_size=11, bold=True)
        
        add_chem_run(p_q, q["intro"], base_font="Times New Roman", base_size=11)
        
        if q.get("image"):
            img_path = os.path.join(IMG_DIR, q["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                run_img = p_img.add_run()
                run_img.add_picture(img_path, width=Cm(9.5))

        for item_key, item_text, is_true, item_sol in q["items"]:
            p_item = doc.add_paragraph()
            p_item.paragraph_format.space_before = Pt(1)
            p_item.paragraph_format.space_after = Pt(1.5)
            p_item.paragraph_format.line_spacing = 1.15
            p_item.paragraph_format.left_indent = Cm(0.5)
            
            r_k = p_item.add_run(f"{item_key}) ")
            set_run_style(r_k, font_name="Times New Roman", font_size=11, bold=True)
            
            add_chem_run(p_item, item_text, base_font="Times New Roman", base_size=11)
            
            if is_teacher:
                status_str = " [ĐÚNG]" if is_true else " [SAI]"
                color_st = COLOR_CORRECT if is_true else COLOR_PRIMARY
                r_st = p_item.add_run(status_str)
                set_run_style(r_st, font_name="Times New Roman", font_size=11, bold=True, color=color_st)
                
                p_sol = doc.add_paragraph()
                p_sol.paragraph_format.space_before = Pt(0)
                p_sol.paragraph_format.space_after = Pt(2)
                p_sol.paragraph_format.left_indent = Cm(1.0)
                r_sol_lbl = p_sol.add_run("-> ")
                set_run_style(r_sol_lbl, font_name="Times New Roman", font_size=10, bold=True, color=COLOR_SOLUTION)
                add_chem_run(p_sol, item_sol, base_font="Times New Roman", base_size=10, italic=True)

    # --- PHẦN III ---
    add_part_heading(doc, "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.", "Thí sinh trả lời từ câu 1 đến câu 6.")
    
    for q in SA_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        
        r_qnum = p_q.add_run(f"Câu {q['num']}. ")
        set_run_style(r_qnum, font_name="Times New Roman", font_size=11, bold=True)
        
        add_chem_run(p_q, q["text"], base_font="Times New Roman", base_size=11)
        
        if q.get("image"):
            img_path = os.path.join(IMG_DIR, q["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                run_img = p_img.add_run()
                w = q.get("image_w", Cm(11.0))
                run_img.add_picture(img_path, width=w)

        if not is_teacher:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(3)
            p_ans.paragraph_format.left_indent = Cm(0.5)
            r_ans_lbl = p_ans.add_run("Trả lời: .................................................................................................................................")
            set_run_style(r_ans_lbl, font_name="Times New Roman", font_size=11, italic=True)
        else:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(2)
            p_ans.paragraph_format.left_indent = Cm(0.5)
            
            r_ans_lbl = p_ans.add_run("Đáp số: ")
            set_run_style(r_ans_lbl, font_name="Times New Roman", font_size=11, bold=True, color=COLOR_CORRECT)
            r_ans_val = p_ans.add_run(q["ans"])
            set_run_style(r_ans_val, font_name="Times New Roman", font_size=11, bold=True, color=COLOR_CORRECT)
            
            add_solution_paragraph(doc, q["sol"])

    # Bảng tổng hợp đáp án ở cuối đề Giáo Viên
    if is_teacher:
        p_table_title = doc.add_paragraph()
        p_table_title.paragraph_format.space_before = Pt(10)
        p_table_title.paragraph_format.space_after = Pt(3)
        r_tt = p_table_title.add_run("BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ SỐ 12")
        set_run_style(r_tt, font_name="Times New Roman", font_size=11.5, bold=True, color=COLOR_PRIMARY)
        
        # Phần I Table
        p_p1 = doc.add_paragraph()
        p_p1.paragraph_format.space_before = Pt(3)
        p_p1.paragraph_format.space_after = Pt(1)
        r_p1 = p_p1.add_run("Phần I: Trắc nghiệm nhiều phương án lựa chọn")
        set_run_style(r_p1, font_name="Times New Roman", font_size=10.5, bold=True)
        
        tbl_ans1 = doc.add_table(rows=2, cols=18)
        tbl_ans1.autofit = False
        for c_idx in range(18):
            cell_q = tbl_ans1.cell(0, c_idx)
            cell_a = tbl_ans1.cell(1, c_idx)
            cell_q.width = Cm(1.0)
            cell_a.width = Cm(1.0)
            
            pq = cell_q.paragraphs[0]
            pq.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pq.paragraph_format.space_before = Pt(1)
            pq.paragraph_format.space_after = Pt(1)
            rq = pq.add_run(str(c_idx + 1))
            set_run_style(rq, font_name="Times New Roman", font_size=9, bold=True)
            
            pa = cell_a.paragraphs[0]
            pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pa.paragraph_format.space_before = Pt(1)
            pa.paragraph_format.space_after = Pt(1)
            ra = pa.add_run(MCQ_QUESTIONS[c_idx]["correct"])
            set_run_style(ra, font_name="Times New Roman", font_size=9.5, bold=True, color=COLOR_CORRECT)
            
            for cell in [cell_q, cell_a]:
                set_cell_border(cell, top={'val':'single','sz':'4','color':'CCCCCC'},
                                      bottom={'val':'single','sz':'4','color':'CCCCCC'},
                                      left={'val':'single','sz':'4','color':'CCCCCC'},
                                      right={'val':'single','sz':'4','color':'CCCCCC'})
                set_cell_margins(cell, top=20, bottom=20, left=10, right=10)

        # Phần II Table
        p_p2 = doc.add_paragraph()
        p_p2.paragraph_format.space_before = Pt(5)
        p_p2.paragraph_format.space_after = Pt(1)
        r_p2 = p_p2.add_run("Phần II: Đúng/Sai")
        set_run_style(r_p2, font_name="Times New Roman", font_size=10.5, bold=True)
        
        tbl_ans2 = doc.add_table(rows=5, cols=5)
        tbl_ans2.autofit = False
        headers_p2 = ["Câu", "a", "b", "c", "d"]
        for j, h in enumerate(headers_p2):
            cell = tbl_ans2.cell(0, j)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(h)
            set_run_style(r, font_name="Times New Roman", font_size=9.5, bold=True)
            set_cell_border(cell, top={'val':'single','sz':'4','color':'CCCCCC'},
                                  bottom={'val':'single','sz':'4','color':'CCCCCC'},
                                  left={'val':'single','sz':'4','color':'CCCCCC'},
                                  right={'val':'single','sz':'4','color':'CCCCCC'})
            set_cell_margins(cell, top=20, bottom=20, left=10, right=10)

        for i, q in enumerate(TF_QUESTIONS):
            cell_q = tbl_ans2.cell(i + 1, 0)
            cell_q.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell_q.paragraphs[0].paragraph_format.space_before = Pt(1)
            cell_q.paragraphs[0].paragraph_format.space_after = Pt(1)
            rq = cell_q.paragraphs[0].add_run(f"Câu {q['num']}")
            set_run_style(rq, font_name="Times New Roman", font_size=9.5, bold=True)
            set_cell_border(cell_q, top={'val':'single','sz':'4','color':'CCCCCC'},
                                    bottom={'val':'single','sz':'4','color':'CCCCCC'},
                                    left={'val':'single','sz':'4','color':'CCCCCC'},
                                    right={'val':'single','sz':'4','color':'CCCCCC'})
            set_cell_margins(cell_q, top=20, bottom=20, left=10, right=10)
            
            for j, item in enumerate(q["items"]):
                cell_item = tbl_ans2.cell(i + 1, j + 1)
                p_it = cell_item.paragraphs[0]
                p_it.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_it.paragraph_format.space_before = Pt(1)
                p_it.paragraph_format.space_after = Pt(1)
                is_true = item[2]
                st_text = "Đ" if is_true else "S"
                st_col = COLOR_CORRECT if is_true else COLOR_PRIMARY
                r_st = p_it.add_run(st_text)
                set_run_style(r_st, font_name="Times New Roman", font_size=9.5, bold=True, color=st_col)
                set_cell_border(cell_item, top={'val':'single','sz':'4','color':'CCCCCC'},
                                           bottom={'val':'single','sz':'4','color':'CCCCCC'},
                                           left={'val':'single','sz':'4','color':'CCCCCC'},
                                           right={'val':'single','sz':'4','color':'CCCCCC'})
                set_cell_margins(cell_item, top=20, bottom=20, left=10, right=10)

        # Phần III Table
        p_p3 = doc.add_paragraph()
        p_p3.paragraph_format.space_before = Pt(5)
        p_p3.paragraph_format.space_after = Pt(1)
        r_p3 = p_p3.add_run("Phần III: Trả lời ngắn")
        set_run_style(r_p3, font_name="Times New Roman", font_size=10.5, bold=True)
        
        tbl_ans3 = doc.add_table(rows=2, cols=6)
        tbl_ans3.autofit = False
        for c_idx in range(6):
            cell_q = tbl_ans3.cell(0, c_idx)
            cell_a = tbl_ans3.cell(1, c_idx)
            cell_q.width = Cm(3.0)
            cell_a.width = Cm(3.0)
            
            pq = cell_q.paragraphs[0]
            pq.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pq.paragraph_format.space_before = Pt(1)
            pq.paragraph_format.space_after = Pt(1)
            rq = pq.add_run(f"Câu {c_idx + 1}")
            set_run_style(rq, font_name="Times New Roman", font_size=9.5, bold=True)
            
            pa = cell_a.paragraphs[0]
            pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pa.paragraph_format.space_before = Pt(1)
            pa.paragraph_format.space_after = Pt(1)
            ra = pa.add_run(SA_QUESTIONS[c_idx]["ans"])
            set_run_style(ra, font_name="Times New Roman", font_size=9.5, bold=True, color=COLOR_CORRECT)
            
            for cell in [cell_q, cell_a]:
                set_cell_border(cell, top={'val':'single','sz':'4','color':'CCCCCC'},
                                      bottom={'val':'single','sz':'4','color':'CCCCCC'},
                                      left={'val':'single','sz':'4','color':'CCCCCC'},
                                      right={'val':'single','sz':'4','color':'CCCCCC'})
                set_cell_margins(cell, top=20, bottom=20, left=10, right=10)

    doc.save(filepath)
    print(f"Đã tạo thành công: {filepath}")

if __name__ == "__main__":
    generate_exam_docx(is_teacher=False)
    generate_exam_docx(is_teacher=True)
