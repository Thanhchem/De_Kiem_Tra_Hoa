# -*- coding: utf-8 -*-
"""
Soạn trọn bộ 5 Đề Kiểm Tra Giữa Học Kì I - Hóa Học 12 (Mã đề 101, 102, 103, 104, 105)
bám sát 100% Ma trận đặc tả của Thầy:
- Mỗi đề gồm 28 câu: Phần I (18 câu), Phần II (4 câu Đúng/Sai), Phần III (6 câu Trả lời ngắn).
- Toàn bộ 140 câu hỏi (5 đề x 28 câu) KHÔNG TRÙNG NHAU giữa các đề.
- Mỗi đề xuất 2 phiên bản Word (.docx) chuẩn "FORM ĐỀ KIỂM TRA":
  1. Bản Học sinh (De_Kiem_Tra_GK1_Hoa_12_De0X_HocSinh.docx): Không đáp án, không gắn nhãn mã câu hỏi.
  2. Bản Giáo viên (De_Kiem_Tra_GK1_Hoa_12_De0X_GiaoVien.docx): Có gắn nhãn Mã câu hỏi Ngân hàng + Mức độ ma trận,
     tô đỏ/gạch chân đáp án đúng, kèm lời giải chi tiết từng câu và bảng đáp án cuối đề.
"""

import os
import re
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
BANK_DIR = r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12"
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(BANK_DIR, exist_ok=True)

COLOR_HF_RED = RGBColor(211, 47, 47)      # #D32F2F Coral Red
COLOR_PRIMARY = RGBColor(46, 125, 50)     # #2E7D32 Dark Green
COLOR_CORRECT = RGBColor(198, 40, 40)     # #C62828 Red
COLOR_SOLUTION = RGBColor(21, 101, 192)   # #1565C0 Blue
COLOR_TAG = RGBColor(0, 105, 92)          # #00695C Teal for Matrix Code Tag


def add_p_border_bottom(p, color="8BC390", sz="6"):
    pPr = p._element.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/></w:pBdr>'))


def add_chem_runs(p, text, bold=False, italic=False, underline=False, color=None, size=11.5):
    """In văn bản kèm xử lý tự động chỉ số dưới/trên cho công thức hóa học và đơn vị."""
    if not text:
        return
    # Tokenize by ~sub~ or ^sup^ or chemical formulas like C2H5OH, H2SO4, Cu(OH)2, [CH2]12, etc.
    # First convert explicit chemical patterns to ~...~ / ^...^ where not already marked
    s = text
    # Degree Celsius / degree alcohol
    s = re.sub(r'(\d+)\s*oC\b', r'\1^o^C', s)
    s = re.sub(r'cm-1\b', r'cm^-1^', s)
    s = re.sub(r'g\.mL-1\b', r'g.mL^-1^', s)
    s = re.sub(r'\b(Ca|Mg|Cu|Fe|Zn|Ba)\s*2\+', r'\1^2+^', s)
    s = re.sub(r'\b(Fe|Al|Cr)\s*3\+', r'\1^3+^', s)
    s = re.sub(r'\b(Na|K|Ag|H|NH4)\+', r'\1^+^', s)
    s = re.sub(r'\b(OH|Cl|Br|NO3|HCO3)-', r'\1^-^', s)
    s = re.sub(r'\b(SO4|CO3)2-', r'\1^2-^', s)
    # Standard chemical subscripts: letter or ) or ] followed by digits
    s = re.sub(r'([A-Za-z\)\]])(\d+)(?=[A-Za-z\(\[\)\]\s\.,;:\-\+\=⇌→]|$)', lambda m: f"{m.group(1)}~{m.group(2)}~", s)
    # Handle cases like C17H35COONa where multiple Letter+Digits occur in one word
    for _ in range(3):
        s = re.sub(r'([A-Za-z\)\]])(\d+)(?=[A-Za-z\(\[\)\]\~]|$)', lambda m: f"{m.group(1)}~{m.group(2)}~", s)

    tokens = re.split(r'(\~[^\~]+\~|\^[^\^]+\^)', s)
    for tok in tokens:
        if not tok:
            continue
        r = p.add_run()
        r.font.name = "Times New Roman"
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        r.underline = underline
        if color:
            r.font.color.rgb = color
        if tok.startswith('~') and tok.endswith('~') and len(tok) > 2:
            r.text = tok[1:-1]
            r.font.subscript = True
        elif tok.startswith('^') and tok.endswith('^') and len(tok) > 2:
            r.text = tok[1:-1]
            r.font.superscript = True
        else:
            r.text = tok


def setup_doc_style(doc):
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11.5)

    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        section.header_distance = Cm(0.6)
        section.footer_distance = Cm(0.6)
        section.different_first_page_header_footer = False

        # Header
        header = section.header
        header.is_linked_to_previous = False
        for p in list(header.paragraphs):
            p._element.getparent().remove(p._element)
        htbl = header.add_table(rows=1, cols=2, width=Cm(18.0))
        htbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        htbl._tbl.tblPr.append(parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:bottom w:val="single" w:sz="6" w:space="4" w:color="8BC390"/>'
            f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'</w:tblBorders>'
        ))
        c0, c1 = htbl.cell(0, 0), htbl.cell(0, 1)
        c0.width, c1.width = Cm(9.0), Cm(9.0)
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(2)
        r0 = p0.add_run("Tài liệu lưu hành nội bộ")
        r0.italic = True
        r0.font.name = "Times New Roman"
        r0.font.size = Pt(10)
        r0.font.color.rgb = COLOR_HF_RED

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_after = Pt(2)
        r1a = p1.add_run("Thầy ")
        r1b = p1.add_run("TRẦN VĂN THẠNH")
        r1b.bold = True
        r1c = p1.add_run(" - 0777.470.803")
        for r in (r1a, r1b, r1c):
            r.font.name = "Times New Roman"
            r.font.size = Pt(10)
            r.font.color.rgb = COLOR_HF_RED

        # Footer
        footer = section.footer
        footer.is_linked_to_previous = False
        for p in list(footer.paragraphs):
            p._element.getparent().remove(p._element)
        ftbl = footer.add_table(rows=1, cols=2, width=Cm(18.0))
        ftbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        ftbl._tbl.tblPr.append(parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="6" w:space="4" w:color="8BC390"/>'
            f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
            f'</w:tblBorders>'
        ))
        fc0, fc1 = ftbl.cell(0, 0), ftbl.cell(0, 1)
        fc0.width, fc1.width = Cm(15.0), Cm(3.0)
        fp0 = fc0.paragraphs[0]
        fp0.paragraph_format.space_before = Pt(2)
        fr0 = fp0.add_run("CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế   |   CS2: 24 Đặng Thái Thân, TP Huế")
        fr0.italic = True
        fr0.font.name = "Times New Roman"
        fr0.font.size = Pt(9.5)
        fr0.font.color.rgb = COLOR_HF_RED

        fp1 = fc1.paragraphs[0]
        fp1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        fp1.paragraph_format.space_before = Pt(2)
        fr1 = fp1.add_run("Trang ")
        fr1.bold = True
        fr1.font.name = "Times New Roman"
        fr1.font.size = Pt(10)
        fr1.font.color.rgb = COLOR_HF_RED
        fp1._p.append(parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w')))


def add_header_block(doc, de_num, ma_de, is_teacher):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl._tbl.tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>'
        f'<w:insideH w:val="none"/><w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    ))
    c0, c1 = tbl.cell(0, 0), tbl.cell(0, 1)
    c0.width, c1.width = Cm(7.5), Cm(10.5)

    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_after = Pt(1)
    r = p0.add_run("SỞ GD&ĐT TP HUẾ\n")
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    r2 = p0.add_run("ĐỀ CHÍNH THỨC — ĐỀ SỐ {:02d}\n".format(de_num))
    r2.bold = True
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(11)
    r2.font.color.rgb = COLOR_PRIMARY
    r3 = p0.add_run("(Đề kiểm tra gồm 04 trang)")
    r3.italic = True
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(10)

    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_after = Pt(1)
    rt1 = p1.add_run("ĐỀ KIỂM TRA GIỮA HỌC KÌ I — NĂM HỌC 2026–2027\n")
    rt1.bold = True
    rt1.font.name = "Times New Roman"
    rt1.font.size = Pt(11.5)
    rt2 = p1.add_run("MÔN: HÓA HỌC — LỚP 12\n")
    rt2.bold = True
    rt2.font.name = "Times New Roman"
    rt2.font.size = Pt(12)
    rt2.font.color.rgb = COLOR_HF_RED
    rt3 = p1.add_run(f"Thời gian làm bài: 45 phút   |   Mã đề: {ma_de}")
    rt3.bold = True
    rt3.font.name = "Times New Roman"
    rt3.font.size = Pt(10.5)

    if is_teacher:
        pgv = doc.add_paragraph()
        pgv.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pgv.paragraph_format.space_before = Pt(3)
        pgv.paragraph_format.space_after = Pt(2)
        rgv = pgv.add_run("(BẢN GIÁO VIÊN — CÓ LỜI GIẢI CHI TIẾT & NHÃN MÃ CÂU HỎI MA TRẬN)")
        rgv.bold = True
        rgv.font.name = "Times New Roman"
        rgv.font.size = Pt(11)
        rgv.font.color.rgb = COLOR_HF_RED

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(3)
    p_info.paragraph_format.space_after = Pt(2)
    r_inf = p_info.add_run("Họ, tên thí sinh: .....................................................................   Lớp: ..............   Số báo danh: ..................")
    r_inf.font.name = "Times New Roman"
    r_inf.font.size = Pt(11)

    p_kl = doc.add_paragraph()
    p_kl.paragraph_format.space_before = Pt(0)
    p_kl.paragraph_format.space_after = Pt(4)
    add_p_border_bottom(p_kl, color="8BC390", sz="6")
    r_kl = p_kl.add_run("Cho biết nguyên tử khối của các nguyên tố: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; S = 32; Cl = 35,5; K = 39; Ca = 40; Br = 80; Ag = 108.")
    r_kl.italic = True
    r_kl.font.name = "Times New Roman"
    r_kl.font.size = Pt(10)


import json
from extract_q_assets import extract_chapter_assets

BANK_PATH = r"c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Tron_Bo_9_Chuong.json"

# Đặc tả Ma trận cho 28 vị trí câu hỏi trong mỗi đề
MATRIX_SLOTS = {
    # PHẦN I (18 câu)
    ("P1", 1):  "1. Ester (HH1.1 – Biết: CTCT ester đơn giản)",
    ("P1", 2):  "1. Ester (HH1.1 – Biết: Danh pháp ester đơn giản)",
    ("P1", 3):  "1. Ester (Thông hiểu: Tính chất, điều chế, thủy phân ester)",
    ("P1", 4):  "2. Lipid (HH1.1 – Biết: Khái niệm lipid, chất béo, acid béo)",
    ("P1", 5):  "2. Lipid (HH1.1 – Biết: Ứng dụng chất béo & omega-3, omega-6)",
    ("P1", 6):  "3. Xà phòng & CGR (HH1.1 – Biết: Khái niệm, thành phần)",
    ("P1", 7):  "3. Xà phòng & CGR (HH1.1 – Biết: Tính chất, phương pháp sản xuất)",
    ("P1", 8):  "3. Xà phòng & CGR (HH3.1 – Vận dụng: Cơ chế giặt rửa, thực tiễn)",
    ("P1", 9):  "4. Glucose & Fructose (HH1.1 – Biết: Trạng thái tự nhiên, tính chất)",
    ("P1", 10): "4. Glucose & Fructose (HH1.2 – Hiểu: Cấu tạo, tính chất hóa học)",
    ("P1", 11): "5. Saccharose & Maltose (HH1.1 – Biết: CTPT, cấu tạo, ứng dụng)",
    ("P1", 12): "5. Saccharose & Maltose (HH1.1 – Biết: Phản ứng với Cu(OH)2, thủy phân)",
    ("P1", 13): "6. Tinh bột & Cellulose (HH1.1 – Biết: Trạng thái tự nhiên, cấu tạo)",
    ("P1", 14): "6. Tinh bột & Cellulose (HH1.1 – Biết: Tính chất vật lí, hóa học)",
    ("P1", 15): "6. Tổng hợp Carbohydrate (HH3.1 – Vận dụng: Tính chất chung carbohydrate)",
    ("P1", 16): "7. Amine (HH1.1 – Biết: Khái niệm, phân loại bậc, danh pháp amine)",
    ("P1", 17): "7. Amine (HH1.1 – Biết: Đặc điểm tính chất vật lí, nhận biết amine)",
    ("P1", 18): "7. Amine (HH3.1 – Vận dụng: Hiện tượng thí nghiệm & giải thích TCHH amine)",
    # PHẦN II (4 câu Đúng/Sai)
    ("P2", 1):  "1. Ester – Chất béo (Thí nghiệm xà phòng hóa / điều chế & thủy phân ester)",
    ("P2", 2):  "2. Glucose & Fructose (Hiện tượng thí nghiệm & giải thích TCHH)",
    ("P2", 3):  "3. Tinh bột & Cellulose (Tính chất hóa học cơ bản & ứng dụng)",
    ("P2", 4):  "4. Amine (Tính chất hóa học đặc trưng, phản ứng base, Br2, điều chế)",
    # PHẦN III (6 câu Trả lời ngắn)
    ("P3", 1):  "1. Ester (HH3.1 – Vận dụng: Xác định chất/lượng ester theo số liệu)",
    ("P3", 2):  "2. Chất béo (HH1.5 – Hiểu: Cấu tạo, phân tử khối, số chất thủy phân)",
    ("P3", 3):  "3. Xà phòng & CGR (HH1.4 – Hiểu/VD: Xác định lượng xà phòng, quy trình)",
    ("P3", 4):  "4. Glucose & Carbohydrate (HH1.3 – Hiểu: Xác định số chất theo TCHH)",
    ("P3", 5):  "5. Tinh bột & Cellulose (HH3.4 – Vận dụng cao: Bài toán lên men xăng E5/rượu/tín chỉ C)",
    ("P3", 6):  "6. Amine (HH1.3 – Hiểu: Công thức cấu tạo và số đồng phân/tính chất amine)",
}

# Bổ sung đầy đủ các công thức/sơ đồ nằm trong đối tượng Equation (OMML) hoặc bảng 2 cột của file gốc
STEM_PATCHES = {
    ("C1", 112): [
        "Ethyl acetate được điều chế bằng cách đun nóng hỗn hợp gồm acetic acid, ethanol và dung dịch H2SO4 98% (xúc tác) trong bình cầu có nhánh, phản ứng tổng hợp xảy ra theo phương trình hoá học sau:",
        "CH3COOH + C2H5OH ⇌ CH3COOC2H5 + H2O (1) (xt: H2SO4 đặc, to)",
        "Mô hình điều chế và tách ethyl acetate như sau:",
        "Chất lỏng trong bình hứng được cho vào phễu chiết chứa sẵn một thể tích tương đương của dung dịch NaCl bão hoà, lắc đều sau đó để lắng và chiết lấy ester."
    ],
    ("C1", 119): [
        "Một nhóm học sinh tiến hành thí nghiệm điều chế isoamyl acetate theo các bước sau:",
        "• Bước 1: Cho vào bình cầu 11 mL isoamyl alcohol ((CH3)2CHCH2CH2OH, khối lượng riêng bằng 0,810 g/cm3), 24 mL acetic acid (CH3COOH, khối lượng riêng bằng 1,049 g/cm3), 1 mL H2SO4 đặc và vài viên đá bọt. Lắp bình cầu vào hệ thống ống sinh hàn đảm bảo chắc chắn và kín khít. Sau đó đun nóng bình cầu trong khoảng 1,5 giờ.",
        "• Bước 2: Sau khi đun, để nguội rồi rót sản phẩm vào phễu chiết, thêm 50 mL nước cất vào phễu, lắc đều rồi để yên khoảng 5 phút, lúc đó chất lỏng tách thành hai lớp, loại bỏ phần chất lỏng phía dưới, lấy phần chất lỏng phía trên.",
        "• Bước 3: Cho từ từ dung dịch NaHCO3 5% vào phần chất lỏng thu lấy ở bước 2 và lắc đều cho đến khi không còn khí thoát ra, thêm tiếp 20 mL dung dịch NaCl bão hòa, lắc đều rồi để yên. Chiết lấy phần chất lỏng phía trên, sau đó loại bỏ nước khỏi hỗn hợp bằng CaCl2 khan.",
        "• Bước 4: Phần chất lỏng thu được sau bước 3 được cho vào hệ thống chưng cất, thu được 7,5 mL isoamyl acetate (khối lượng riêng bằng 0,876 g/cm3)."
    ],
    ("C1", 168): [
        "Thực hiện các phản ứng ester hóa giữa glycerol với acetic acid (xúc tác H2SO4 đặc, đun nóng) theo sơ đồ sau:",
        "Glycerol --(+ CH3COOH)--> X --(+ CH3COOH)--> Y --(+ CH3COOH)--> Z",
        "Phần trăm khối lượng của oxygen trong Y là x%. Giá trị của x bằng bao nhiêu? (Không làm tròn kết quả các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng phần mười)."
    ],
    ("C1", 185): [
        "Bioethanol (cồn sinh học) là thành phần chính để pha chế xăng sinh học, được sản xuất bằng cách lên men các sản phẩm hữu cơ giàu tinh bột hoặc cellulose. Sơ đồ sản xuất bioethanol từ tinh bột gồm hai giai đoạn với hiệu suất lần lượt là 80% và 70% như sau:",
        "(C6H10O5)n --(1) H=80%--> C6H12O6 --(2) H=70%--> C2H5OH",
        "Từ m kg sắn khô chứa 72% tinh bột, sản xuất 400 lít xăng E10 (chứa 10% ethanol về thể tích). Biết ethanol có khối lượng riêng bằng 0,79 g/mL. Giá trị của m là bao nhiêu? (Kết quả các phép tính trung gian không được làm tròn, chỉ kết quả cuối cùng được làm tròn đến hàng đơn vị)."
    ],
    ("C2", 92): [
        "Fructose là loại monosaccharide phổ biến trong đời sống, fructose được tìm thấy trong mật ong, hoa, quả, và hầu hết các loại rau củ. Fructose tồn tại chủ yếu ở hai dạng mạch vòng là α-fructose và β-fructose. Các đồng phân dạng mạch vòng và mạch hở có thể chuyển hóa lẫn nhau như hình sau:"
    ],
    ("C2", 121): [
        "Cho sơ đồ điều chế ethanol bằng phương pháp lên men tinh bột như sau:",
        "(C6H10O5)n --(1) + H2O (xt, to)--> C6H12O6 --(2) enzyme--> C2H5OH",
        "Lên men 338,35 kg sắn khô (chứa 35% khối lượng là tinh bột, phần còn lại là các chất không có khả năng lên men thành ethanol) với hiệu suất mỗi quá trình (1) và (2) lần lượt là 90% và 80%. Toàn bộ lượng ethanol sinh ra để pha chế xăng E5 (chứa 5% thể tích ethanol). Biết khối lượng riêng của ethanol nguyên chất là 0,8 g/mL. Xác định thể tích (tính bằng Lít) xăng E5 thu được sau pha trộn. (Không làm tròn kết quả các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng đơn vị)."
    ],
    ("C2", 122): [
        "Tiến hành sản xuất 1,8 triệu chai rượu vang 13,8^o^ dung tích 750 mL bằng phương pháp lên men rượu với nguyên liệu là m tấn quả nho tươi (chứa 15% glucose về khối lượng). Biết khối lượng riêng của ethyl alcohol là 0,8 g/mL. Giả thiết trong thành phần quả nho tươi chỉ có glucose bị lên men rượu, hiệu suất toàn bộ quá trình sản xuất là 60%; độ rượu là % thể tích ethanol trong rượu (hỗn hợp ethanol và nước). Tính m. (Không làm tròn kết quả các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng đơn vị)."
    ],
    ("C3", 128): [
        "Trong phòng thí nghiệm, aniline (C6H5NH2) được điều chế từ nitrobenzene (C6H5NO2) theo sơ đồ:",
        "C6H5NO2 --(1) + Fe + HCl, to--> C6H5NH3Cl --(2) + dd NaOH--> C6H5NH2",
        "Trong một thí nghiệm, sử dụng 6,15 gam nitrobenzene để điều chế aniline theo sơ đồ trên, thu được 3,72 gam aniline tinh khiết. Cho biết số sóng hấp thụ đặc trưng của một số nhóm chức trên phổ hồng ngoại như sau:"
    ],
    ("C3", 131): [
        "Trong phòng thí nghiệm, aniline được điều chế từ nitrobenzene qua hai giai đoạn chính được thể hiện trong sơ đồ:",
        "C6H5NO2 --(1) + Fe + HCl, to--> C6H5NH3Cl --(2) + dd NaOH--> C6H5NH2",
        "Aniline sau khi được điều chế từ giai đoạn (2) sẽ tiếp tục được tinh chế như mô tả trong hình dưới. Từ 10 mL nitrobenzene (D = 1,20 g.mL-1) ban đầu và các hóa chất khác lấy dư, cuối cùng thu được 5,93 gam aniline. Hiệu suất của quá trình tổng hợp aniline là h%."
    ],
    ("C3", 132): [
        "Cho sơ đồ chuyển hóa sau (X, Y, Z, T đều là các chất hữu cơ):",
        "Benzene --(1) + HNO3/H2SO4 đặc, to--> X --(2) + Fe + HCl, to--> Y --(3) + dd NaOH--> Z --(4) + HNO2 + HCl (0-5 oC)--> T"
    ],
    ("C3", 163): [
        "Amine X có công thức phân tử là C4H11N. Cho X phản ứng với hỗn hợp dung dịch HCl và NaNO2 ở nhiệt độ thường thấy có khí không màu thoát ra. Có bao nhiêu công thức cấu tạo phù hợp với X?"
    ]
}

P2_ITEM_PATCHES = {
    ("C2", 92): {
        "a": "Nhóm –OH (*) ở α-fructose được gọi là nhóm –OH hemiacetal và –OH (**) ở β-fructose được gọi là nhóm –OH hemiketal.",
        "b": "Ở điều kiện thường hoặc khi đun nóng, fructose đều bị oxi hóa bởi Cu(OH)2 trong môi trường kiềm.",
        "c": "Fructose có phản ứng tráng bạc do cả hai dạng α-fructose và β-fructose đều phản ứng trực tiếp với thuốc thử Tollens.",
        "d": "Dạng α-fructose và β-fructose có sự khác nhau về hướng trong không gian của các nhóm liên kết ở vị trí C2."
    }
}

# Danh sách 5 đề x 28 câu = 140 câu hỏi KHÔNG TRÙNG NHAU từ Ngân hàng (Chương, Số thứ tự câu gốc)
TESTS_SPEC = [
    {
        "de_num": 1,
        "ma_de": "101",
        "P1": [
            ("C1", 2), ("C1", 8), ("C1", 58), ("C1", 37), ("C1", 39), ("C1", 46), ("C1", 51), ("C1", 102),
            ("C2", 1), ("C2", 38), ("C2", 3), ("C2", 20), ("C2", 21), ("C2", 29), ("C2", 87),
            ("C3", 1), ("C3", 22), ("C3", 106)
        ],
        "P2": [("C1", 144), ("C2", 92), ("C2", 97), ("C3", 139)],
        "P3": [("C1", 162), ("C1", 171), ("C1", 177), ("C2", 109), ("C2", 121), ("C3", 163)],
    },
    {
        "de_num": 2,
        "ma_de": "102",
        "P1": [
            ("C1", 6), ("C1", 9), ("C1", 59), ("C1", 38), ("C1", 43), ("C1", 47), ("C1", 57), ("C1", 86),
            ("C2", 2), ("C2", 39), ("C2", 17), ("C2", 56), ("C2", 22), ("C2", 36), ("C2", 88),
            ("C3", 3), ("C3", 23), ("C3", 107)
        ],
        "P2": [("C1", 139), ("C2", 93), ("C2", 98), ("C3", 128)],
        "P3": [("C1", 164), ("C1", 157), ("C1", 178), ("C2", 110), ("C1", 185), ("C3", 153)],
    },
    {
        "de_num": 3,
        "ma_de": "103",
        "P1": [
            ("C1", 13), ("C1", 10), ("C1", 60), ("C1", 41), ("C1", 44), ("C1", 48), ("C1", 87), ("C1", 88),
            ("C2", 11), ("C2", 43), ("C2", 18), ("C2", 58), ("C2", 25), ("C2", 71), ("C2", 89),
            ("C3", 4), ("C3", 24), ("C3", 108)
        ],
        "P2": [("C1", 142), ("C2", 94), ("C2", 99), ("C3", 130)],
        "P3": [("C1", 170), ("C1", 172), ("C1", 179), ("C2", 108), ("C2", 122), ("C3", 165)],
    },
    {
        "de_num": 4,
        "ma_de": "104",
        "P1": [
            ("C1", 14), ("C1", 12), ("C1", 62), ("C1", 42), ("C1", 79), ("C1", 49), ("C1", 90), ("C1", 89),
            ("C2", 13), ("C2", 44), ("C2", 19), ("C2", 62), ("C2", 27), ("C2", 74), ("C2", 90),
            ("C3", 5), ("C3", 25), ("C3", 109)
        ],
        "P2": [("C1", 112), ("C2", 96), ("C2", 101), ("C3", 131)],
        "P3": [("C1", 166), ("C1", 174), ("C1", 180), ("C2", 114), ("C2", 123), ("C3", 168)],
    },
    {
        "de_num": 5,
        "ma_de": "105",
        "P1": [
            ("C1", 26), ("C1", 15), ("C1", 71), ("C1", 85), ("C1", 82), ("C1", 55), ("C1", 91), ("C1", 94),
            ("C2", 14), ("C2", 50), ("C2", 64), ("C2", 63), ("C2", 31), ("C2", 77), ("C2", 91),
            ("C3", 17), ("C3", 16), ("C3", 111)
        ],
        "P2": [("C1", 119), ("C2", 95), ("C2", 102), ("C3", 132)],
        "P3": [("C1", 168), ("C1", 176), ("C1", 175), ("C2", 115), ("C2", 119), ("C3", 170)],
    },
]


def parse_abcd_from_opt_lines(opt_lines):
    """Tách 4 phương án A, B, C, D từ danh sách dòng phương án."""
    joined = "\t".join(opt_lines)
    matches = list(re.finditer(r'([A-D])\.\s*', joined))
    opts = {}
    if len(matches) >= 4:
        for i in range(4):
            ch = matches[i].group(1)
            start = matches[i].end()
            end = matches[i+1].start() if i + 1 < len(matches) else len(joined)
            opts[ch] = joined[start:end].strip()
    return opts


def add_options_paragraph_no_table(doc, opts, correct_ans, is_teacher):
    """Trình bày 4 phương án A, B, C, D bằng Paragraph + Tab stops (TUYỆT ĐỐI KHÔNG DÙNG BẢNG)."""
    a_txt = opts.get("A", "")
    b_txt = opts.get("B", "")
    c_txt = opts.get("C", "")
    d_txt = opts.get("D", "")
    max_len = max(len(a_txt), len(b_txt), len(c_txt), len(d_txt))

    def render_opt(p, letter, txt, prefix_tab=False):
        if prefix_tab:
            p.add_run("\t")
        is_corr = is_teacher and (letter == correct_ans)
        r_lbl = p.add_run(f"{letter}. ")
        r_lbl.bold = True
        r_lbl.underline = is_corr
        r_lbl.font.name = "Times New Roman"
        r_lbl.font.size = Pt(11.5)
        r_lbl.font.color.rgb = COLOR_CORRECT if is_corr else COLOR_SOLUTION
        add_chem_runs(p, txt, bold=is_corr, underline=is_corr, color=(COLOR_CORRECT if is_corr else None), size=11.5)

    if max_len <= 18:
        # 1 hàng 4 cột với Tab stops: 4.5cm, 9.0cm, 13.5cm
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(4.5), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(13.5), WD_TAB_ALIGNMENT.LEFT)
        render_opt(p, "A", a_txt, False)
        render_opt(p, "B", b_txt, True)
        render_opt(p, "C", c_txt, True)
        render_opt(p, "D", d_txt, True)
    elif max_len <= 42:
        # 2 hàng x 2 cột với Tab stop: 9.0cm
        for pair in [("A", "B"), ("C", "D")]:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
            render_opt(p, pair[0], opts.get(pair[0], ""), False)
            render_opt(p, pair[1], opts.get(pair[1], ""), True)
    else:
        # 4 hàng riêng biệt
        for letter in ["A", "B", "C", "D"]:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(2)
            render_opt(p, letter, opts.get(letter, ""), False)


def add_data_table_to_doc(doc, tbl_data):
    """Chèn bảng số liệu đề bài (nếu có trong câu hỏi)."""
    if not tbl_data or not tbl_data[0]:
        return
    rows = len(tbl_data)
    cols = len(tbl_data[0])
    t = doc.add_table(rows=rows, cols=cols)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t._tbl.tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="888888"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="888888"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="888888"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="888888"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="888888"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="888888"/>'
        f'</w:tblBorders>'
    ))
    for r_idx, r_vals in enumerate(tbl_data):
        for c_idx, val in enumerate(r_vals):
            if c_idx < cols:
                cell_p = t.cell(r_idx, c_idx).paragraphs[0]
                cell_p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0 or r_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
                cell_p.paragraph_format.space_before = Pt(1.5)
                cell_p.paragraph_format.space_after = Pt(1.5)
                add_chem_runs(cell_p, val, bold=(r_idx == 0), size=10.5)


def add_solution_block(doc, q, opts=None):
    """Trình bày lời giải chi tiết thụt lề 0.6 cm (không dùng bảng)."""
    sol_text = (q.get("solution") or "").strip()
    if not sol_text:
        # Tạo lời giải tự động chuẩn xác cho các câu Nhận biết/Thông hiểu Phần I
        ans = q.get("answer", "")
        chosen_txt = opts.get(ans, "") if opts else ""
        sol_text = f"Chọn đáp án {ans} ({chosen_txt})."

    p_hd = doc.add_paragraph()
    p_hd.paragraph_format.left_indent = Cm(0.6)
    p_hd.paragraph_format.space_before = Pt(2)
    p_hd.paragraph_format.space_after = Pt(2)
    r_title = p_hd.add_run("Hướng dẫn giải: ")
    r_title.bold = True
    r_title.font.name = "Times New Roman"
    r_title.font.size = Pt(11)
    r_title.font.color.rgb = COLOR_SOLUTION

    lines = [ln.strip() for ln in sol_text.split("\n") if ln.strip()]
    if len(lines) == 1:
        add_chem_runs(p_hd, lines[0], size=11)
    else:
        add_chem_runs(p_hd, f"Đáp án: {q.get('answer', '')}", bold=True, color=COLOR_CORRECT, size=11)
        for ln in lines:
            p_ln = doc.add_paragraph()
            p_ln.paragraph_format.left_indent = Cm(0.6)
            p_ln.paragraph_format.space_before = Pt(0.5)
            p_ln.paragraph_format.space_after = Pt(1.5)
            add_chem_runs(p_ln, "• " + ln if not ln.startswith(("•", "-", "+")) else ln, size=11)


def generate_single_test_docx(test_cfg, bank_map, assets_map, is_teacher):
    doc = docx.Document()
    setup_doc_style(doc)
    de_num = test_cfg["de_num"]
    ma_de = test_cfg["ma_de"]
    add_header_block(doc, de_num, ma_de, is_teacher)

    # ================= PHẦN I =================
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_before = Pt(6)
    p_sec1.paragraph_format.space_after = Pt(3)
    r_s1 = p_sec1.add_run("PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (4,5 điểm). ")
    r_s1.bold = True
    r_s1.font.name = "Times New Roman"
    r_s1.font.size = Pt(11.5)
    r_s1.font.color.rgb = COLOR_PRIMARY
    r_s1_sub = p_sec1.add_run("Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")
    r_s1_sub.italic = True
    r_s1_sub.font.name = "Times New Roman"
    r_s1_sub.font.size = Pt(11.5)

    p1_answers = []
    for idx, (ch_code, orig_qnum) in enumerate(test_cfg["P1"], 1):
        q = copy.deepcopy(bank_map[(ch_code, orig_qnum)])
        if (ch_code, orig_qnum) in STEM_PATCHES:
            q["stem_lines"] = STEM_PATCHES[(ch_code, orig_qnum)]
        if (ch_code, orig_qnum) in P2_ITEM_PATCHES:
            q["p2_items"] = P2_ITEM_PATCHES[(ch_code, orig_qnum)]
        # Fix C1 Q47 answer in bank if needed (CH3[CH2]12COONa is D)
        if ch_code == "C1" and orig_qnum == 47:
            q["answer"] = "D"
        p1_answers.append((idx, q["answer"], q["id"]))

        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        r_qnum = p_q.add_run(f"Câu {idx}. ")
        r_qnum.bold = True
        r_qnum.font.name = "Times New Roman"
        r_qnum.font.size = Pt(11.5)
        r_qnum.font.color.rgb = COLOR_PRIMARY

        if is_teacher:
            slot_desc = MATRIX_SLOTS.get(("P1", idx), "")
            r_tag = p_q.add_run(f"[Mã NH: {q['id']} | {slot_desc} | ĐA: {q['answer']}] ")
            r_tag.bold = True
            r_tag.font.name = "Times New Roman"
            r_tag.font.size = Pt(10.5)
            r_tag.font.color.rgb = COLOR_TAG

        stem_lines = q["stem_lines"]
        if stem_lines:
            add_chem_runs(p_q, stem_lines[0], size=11.5)
            for extra_ln in stem_lines[1:]:
                p_ex = doc.add_paragraph()
                p_ex.paragraph_format.space_before = Pt(1)
                p_ex.paragraph_format.space_after = Pt(1)
                add_chem_runs(p_ex, extra_ln, size=11.5)

        # Insert any extracted images or data tables for this question
        c_int = int(ch_code[1])
        q_assets = assets_map.get(c_int, {}).get(orig_qnum, {"images": [], "tables": []})
        for img_path in q_assets["images"]:
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_path, width=Cm(6.5))
        for tbl_data in q_assets["tables"]:
            add_data_table_to_doc(doc, tbl_data)

        opts = parse_abcd_from_opt_lines(q["opt_lines"])
        add_options_paragraph_no_table(doc, opts, q["answer"], is_teacher)

        if is_teacher:
            add_solution_block(doc, q, opts)

    # ================= PHẦN II =================
    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_before = Pt(8)
    p_sec2.paragraph_format.space_after = Pt(3)
    r_s2 = p_sec2.add_run("PHẦN II. Câu trắc nghiệm đúng sai (4,0 điểm). ")
    r_s2.bold = True
    r_s2.font.name = "Times New Roman"
    r_s2.font.size = Pt(11.5)
    r_s2.font.color.rgb = COLOR_PRIMARY
    r_s2_sub = p_sec2.add_run("Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")
    r_s2_sub.italic = True
    r_s2_sub.font.name = "Times New Roman"
    r_s2_sub.font.size = Pt(11.5)

    p2_matrix_tags = {
        "a": "HH1.1 – Biết",
        "b": "HH2.1 – Hiểu",
        "c": "HH2.2 – Hiểu",
        "d": "HH3.1 – Vận dụng"
    }
    p2_answers = []
    for idx, (ch_code, orig_qnum) in enumerate(test_cfg["P2"], 1):
        q = copy.deepcopy(bank_map[(ch_code, orig_qnum)])
        if (ch_code, orig_qnum) in STEM_PATCHES:
            q["stem_lines"] = STEM_PATCHES[(ch_code, orig_qnum)]
        if (ch_code, orig_qnum) in P2_ITEM_PATCHES:
            q["p2_items"] = P2_ITEM_PATCHES[(ch_code, orig_qnum)]
        p2_answers.append((idx, q["answer"], q["id"]))

        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        r_qnum = p_q.add_run(f"Câu {idx}. ")
        r_qnum.bold = True
        r_qnum.font.name = "Times New Roman"
        r_qnum.font.size = Pt(11.5)
        r_qnum.font.color.rgb = COLOR_PRIMARY

        if is_teacher:
            slot_desc = MATRIX_SLOTS.get(("P2", idx), "")
            r_tag = p_q.add_run(f"[Mã NH: {q['id']} | {slot_desc} | ĐA: {q['answer']}] ")
            r_tag.bold = True
            r_tag.font.name = "Times New Roman"
            r_tag.font.size = Pt(10.5)
            r_tag.font.color.rgb = COLOR_TAG

        stem_lines = q["stem_lines"]
        if stem_lines:
            add_chem_runs(p_q, stem_lines[0], size=11.5)
            for extra_ln in stem_lines[1:]:
                p_ex = doc.add_paragraph()
                p_ex.paragraph_format.space_before = Pt(1)
                p_ex.paragraph_format.space_after = Pt(1)
                add_chem_runs(p_ex, extra_ln, size=11.5)

        c_int = int(ch_code[1])
        q_assets = assets_map.get(c_int, {}).get(orig_qnum, {"images": [], "tables": []})
        for img_path in q_assets["images"]:
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_path, width=Cm(8.5))
        for tbl_data in q_assets["tables"]:
            add_data_table_to_doc(doc, tbl_data)

        # Parse a:Đ - b:S - c:Đ - d:S
        ans_map = {}
        for part_ans in q["answer"].split("-"):
            part_ans = part_ans.strip()
            if ":" in part_ans:
                k_a, v_a = part_ans.split(":", 1)
                ans_map[k_a.strip()] = v_a.strip()

        for key in ["a", "b", "c", "d"]:
            stmt_txt = q.get("p2_items", {}).get(key, "")
            p_it = doc.add_paragraph()
            p_it.paragraph_format.left_indent = Cm(0.4)
            p_it.paragraph_format.space_before = Pt(1)
            p_it.paragraph_format.space_after = Pt(1.5)
            r_k = p_it.add_run(f"{key}) ")
            r_k.bold = True
            r_k.font.name = "Times New Roman"
            r_k.font.size = Pt(11.5)
            if is_teacher:
                m_tag = p2_matrix_tags.get(key, "Hiểu")
                r_mt = p_it.add_run(f"[{m_tag}] ")
                r_mt.bold = True
                r_mt.font.name = "Times New Roman"
                r_mt.font.size = Pt(10.5)
                r_mt.font.color.rgb = COLOR_SOLUTION
            add_chem_runs(p_it, stmt_txt, size=11.5)
            if is_teacher:
                ds_val = "Đúng" if ans_map.get(key) == "Đ" else "Sai"
                r_ds = p_it.add_run(f"   → [{ds_val}]")
                r_ds.bold = True
                r_ds.font.name = "Times New Roman"
                r_ds.font.size = Pt(11)
                r_ds.font.color.rgb = COLOR_CORRECT

        if is_teacher:
            add_solution_block(doc, q)

    # ================= PHẦN III =================
    p_sec3 = doc.add_paragraph()
    p_sec3.paragraph_format.space_before = Pt(8)
    p_sec3.paragraph_format.space_after = Pt(3)
    r_s3 = p_sec3.add_run("PHẦN III. Câu trắc nghiệm trả lời ngắn (1,5 điểm). ")
    r_s3.bold = True
    r_s3.font.name = "Times New Roman"
    r_s3.font.size = Pt(11.5)
    r_s3.font.color.rgb = COLOR_PRIMARY
    r_s3_sub = p_sec3.add_run("Thí sinh trả lời từ câu 1 đến câu 6.")
    r_s3_sub.italic = True
    r_s3_sub.font.name = "Times New Roman"
    r_s3_sub.font.size = Pt(11.5)

    p3_answers = []
    for idx, (ch_code, orig_qnum) in enumerate(test_cfg["P3"], 1):
        q = copy.deepcopy(bank_map[(ch_code, orig_qnum)])
        if (ch_code, orig_qnum) in STEM_PATCHES:
            q["stem_lines"] = STEM_PATCHES[(ch_code, orig_qnum)]
        if (ch_code, orig_qnum) in P2_ITEM_PATCHES:
            q["p2_items"] = P2_ITEM_PATCHES[(ch_code, orig_qnum)]
        p3_answers.append((idx, q["answer"], q["id"]))

        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        r_qnum = p_q.add_run(f"Câu {idx}. ")
        r_qnum.bold = True
        r_qnum.font.name = "Times New Roman"
        r_qnum.font.size = Pt(11.5)
        r_qnum.font.color.rgb = COLOR_PRIMARY

        if is_teacher:
            slot_desc = MATRIX_SLOTS.get(("P3", idx), "")
            r_tag = p_q.add_run(f"[Mã NH: {q['id']} | {slot_desc} | Đáp số: {q['answer']}] ")
            r_tag.bold = True
            r_tag.font.name = "Times New Roman"
            r_tag.font.size = Pt(10.5)
            r_tag.font.color.rgb = COLOR_TAG

        stem_lines = q["stem_lines"]
        if stem_lines:
            add_chem_runs(p_q, stem_lines[0], size=11.5)
            for extra_ln in stem_lines[1:]:
                p_ex = doc.add_paragraph()
                p_ex.paragraph_format.space_before = Pt(1)
                p_ex.paragraph_format.space_after = Pt(1)
                add_chem_runs(p_ex, extra_ln, size=11.5)

        c_int = int(ch_code[1])
        q_assets = assets_map.get(c_int, {}).get(orig_qnum, {"images": [], "tables": []})
        for img_path in q_assets["images"]:
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_path, width=Cm(7.5))
        for tbl_data in q_assets["tables"]:
            add_data_table_to_doc(doc, tbl_data)

        if not is_teacher:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(3)
            r_a = p_ans.add_run("Đáp số: ....................................")
            r_a.bold = True
            r_a.italic = True
            r_a.font.name = "Times New Roman"
            r_a.font.size = Pt(11)
        else:
            add_solution_block(doc, q)

    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(4)
    r_end = p_end.add_run("—————— HẾT ——————")
    r_end.bold = True
    r_end.font.name = "Times New Roman"
    r_end.font.size = Pt(11.5)

    # Bảng đáp án tổng hợp cuối đề cho Bản Giáo viên
    if is_teacher:
        p_ba = doc.add_paragraph()
        p_ba.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ba.paragraph_format.space_before = Pt(6)
        p_ba.paragraph_format.space_after = Pt(3)
        r_ba = p_ba.add_run(f"BẢNG ĐÁP ÁN & ĐỐI CHIẾU MÃ NGÂN HÀNG — ĐỀ SỐ {de_num:02d} (MÃ ĐỀ {ma_de})")
        r_ba.bold = True
        r_ba.font.name = "Times New Roman"
        r_ba.font.size = Pt(12)
        r_ba.font.color.rgb = COLOR_HF_RED

        for title_s, ans_list in [
            ("1. Phần I – Trắc nghiệm nhiều phương án lựa chọn (18 câu):", p1_answers),
            ("2. Phần II – Trắc nghiệm Đúng / Sai (4 câu):", p2_answers),
            ("3. Phần III – Trắc nghiệm Trả lời ngắn (6 câu):", p3_answers),
        ]:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(3)
            p_sub.paragraph_format.space_after = Pt(1)
            r_sub = p_sub.add_run(title_s)
            r_sub.bold = True
            r_sub.font.name = "Times New Roman"
            r_sub.font.size = Pt(11)
            r_sub.font.color.rgb = COLOR_PRIMARY

            p_items = doc.add_paragraph()
            p_items.paragraph_format.left_indent = Cm(0.5)
            p_items.paragraph_format.space_after = Pt(2)
            chunks = [f"Câu {q_i}: {a_v} ({q_id})" for q_i, a_v, q_id in ans_list]
            r_it = p_items.add_run("   |   ".join(chunks))
            r_it.font.name = "Times New Roman"
            r_it.font.size = Pt(10.5)

    suffix = "GiaoVien" if is_teacher else "HocSinh"
    fname = f"De_Kiem_Tra_GK1_Hoa_12_De{de_num:02d}_{suffix}.docx"
    out_p1 = os.path.join(OUT_DIR, fname)
    out_p2 = os.path.join(BANK_DIR, fname)
    doc.save(out_p1)
    doc.save(out_p2)
    print(f"  + Saved: {out_p1}")


if __name__ == "__main__":
    import copy
    with open(BANK_PATH, "r", encoding="utf-8") as f:
        bank_data = json.load(f)

    bank_map = {}
    for ch_code in ["C1", "C2", "C3"]:
        for q in bank_data[ch_code]["questions"]:
            bank_map[(ch_code, q["q_num"])] = q

    # Kiểm tra 140 câu hỏi hoàn toàn không trùng nhau
    all_used = []
    for t in TESTS_SPEC:
        for part_k in ["P1", "P2", "P3"]:
            for pair in t[part_k]:
                all_used.append(pair)
    print(f"Tổng số câu hỏi sử dụng trong 5 đề: {len(all_used)} câu | Số câu duy nhất (không trùng): {len(set(all_used))} câu.")
    assert len(all_used) == len(set(all_used)) == 140, "Có câu hỏi trùng lặp!"

    print("Đang trích xuất hình ảnh và bảng dữ liệu từ Ngân hàng...")
    assets_map = {c: extract_chapter_assets(c) for c in [1, 2, 3]}

    for t_cfg in TESTS_SPEC:
        print(f"\nĐang tạo Đề số {t_cfg['de_num']:02d} (Mã đề {t_cfg['ma_de']})...")
        generate_single_test_docx(t_cfg, bank_map, assets_map, is_teacher=False)
        generate_single_test_docx(t_cfg, bank_map, assets_map, is_teacher=True)

    print("\nHOÀN TẤT 10 FILE WORD CHO 5 ĐỀ KIỂM TRA GIỮA KÌ 1 HÓA 12!")
