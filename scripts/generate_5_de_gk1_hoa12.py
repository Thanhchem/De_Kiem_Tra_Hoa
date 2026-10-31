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
import copy
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

OPT_PATCHES = {
    ("C2", 220): ["A. Rau răm.\tB. Hạt gạo.\tC. Hạt óc chó.\tD. Quả nho chín."]
}

STEM_PATCHES = {
    ("C2", 220): ["Tinh bột là sản phẩm quang hóa của cây xanh và được dự trữ trong các loại hạt, củ, quả,... Loại nông sản nào sau đây có chứa hàm lượng tinh bột cao nhất?"],
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
    ],
    # --- Bổ sung STEM_PATCHES cho Đề 06 đến Đề 10 ---
    ("C1", 160): [
        "Methyl cinnamate (ester của cinnamic acid, có công thức cấu tạo C6H5–CH=CH–COOCH3) là chất rắn có mùi thơm nồng của dâu tây.",
        "Tỉ lệ khối lượng của carbon và oxygen trong phân tử methyl cinnamate là k. Tìm giá trị của k (làm tròn kết quả đến hàng phần trăm)."
    ],
    ("C1", 165): [
        "Aspirin (o-CH3COO–C6H4–COOH, M = 180) có tác dụng giảm đau, hạ sốt, chống viêm. Aspirin được điều chế từ phenol (C6H5OH, M = 94) theo sơ đồ chuyển hóa sau:",
        "C6H5OH (phenol) ──► o-HO–C6H4–COOH (salicylic acid) ──(+ (CH3CO)2O, H2SO4 đặc)──► o-CH3COO–C6H4–COOH (aspirin)",
        "Một viên thuốc aspirin có chứa 80 mg aspirin. Tính khối lượng (gam) phenol cần thiết để sản xuất 100 lọ aspirin (mỗi lọ có 100 viên) với hiệu suất cả quá trình là 69,63%. (Làm tròn kết quả đến hàng đơn vị)."
    ],
    ("C2", 42): [
        "Tiến hành thí nghiệm khảo sát tính chất hóa học của glucose theo các bước sau:",
        "• Bước 1: Cho khoảng 2 mL dung dịch NaOH 10% vào ống nghiệm. Sau đó, thêm khoảng 0,5 mL dung dịch CuSO4 5% vào, lắc nhẹ.",
        "• Bước 2: Cho tiếp khoảng 3 mL dung dịch glucose 2% vào ống nghiệm và lắc đều ở nhiệt độ thường.",
        "• Bước 3: Đun nóng nhẹ ống nghiệm chứa hỗn hợp thu được sau bước 2 trên ngọn lửa đèn cồn."
    ],
    ("C2", 46): [
        "Tiến hành thí nghiệm theo các bước sau:",
        "• Bước 1: Cho 5 giọt dung dịch CuSO4 0,5% vào ống nghiệm sạch.",
        "• Bước 2: Thêm 1 mL dung dịch NaOH 10% vào ống nghiệm, lắc đều; gạn phần dung dịch, giữ lại kết tủa.",
        "• Bước 3: Thêm tiếp 2 mL dung dịch glucose 1% vào ống nghiệm, lắc đều ở nhiệt độ thường."
    ],
    ("C2", 51): [
        "Glucose bị lên men tạo thành các sản phẩm khác nhau tùy thuộc vào loại enzyme có trong men được sử dụng. Các phản ứng lên men xảy ra theo sơ đồ sau:",
        "(1) C6H12O6 ──(enzyme 1)──► 2X + 2CO2↑",
        "(2) C6H12O6 ──(enzyme 2)──► 2Z   (biết Z là hợp chất hữu cơ tạp chức có trong sữa chua)."
    ],
    ("C2", 70): [
        "Tiến hành thí nghiệm thủy phân saccharose theo các bước sau:",
        "• Bước 1: Cho khoảng 3 mL dung dịch saccharose 2% vào ống nghiệm, sau đó thêm khoảng 1 mL dung dịch HCl 1 M vào, lắc đều.",
        "• Bước 2: Đặt ống nghiệm trong cốc thủy tinh chứa nước nóng, đun cách thủy trong 10 phút rồi để nguội.",
        "• Bước 3: Thêm từ từ tinh thể NaHCO3 vào ống nghiệm sau bước 2 đến khi ngừng sủi bọt khí.",
        "• Bước 4: Cho khoảng 2 mL dung dịch thu được sau bước 3 vào ống nghiệm chứa Cu(OH)2 (được điều chế bằng cách cho 0,5 mL dung dịch CuSO4 5% vào 2 mL dung dịch NaOH 10%, lắc nhẹ). Sau đó đặt ống nghiệm trong cốc thủy tinh chứa nước nóng khoảng 5 phút."
    ],
    ("C2", 84): [
        "Cho bốn dung dịch mất nhãn gồm: ethanol, glucose, glycerol và hồ tinh bột được kí hiệu ngẫu nhiên là X, Y, Z và T. Tiến hành thí nghiệm nhận biết các dung dịch trên và ghi lại kết quả ở bảng sau:"
    ],
    ("C2", 85): [
        "Glucose và fructose là hai monosaccharide phổ biến trong tự nhiên, đóng vai trò cung cấp năng lượng quan trọng cho tế bào sống. Thực hiện các thí nghiệm khảo sát cấu tạo và tính chất hóa học của glucose và fructose."
    ],
    ("C2", 86): [
        "Saccharose (đường kính, đường mía) và maltose (đường mạch nha) là hai disaccharide quan trọng trong công nghiệp thực phẩm. Tiến hành khảo sát cấu tạo phân tử và tính chất hóa học của saccharose."
    ],
    ("C2", 100): [
        "Cellulose là thành phần chính của thành tế bào thực vật, làm nên bộ khung của cây cối. Cellulose có nhiều trong bông, đay, tre, nứa, gỗ,... Từ cellulose người ta điều chế được cellulose acetate (có thành phần chủ yếu là cellulose triacetate và cellulose diacetate) theo các phương trình hóa học sau:",
        "[C6H7O2(OH)3]n + 3n(CH3CO)2O ──(H2SO4 đặc, to)──► [C6H7O2(OCOCH3)3]n + 3nCH3COOH   (1)",
        "[C6H7O2(OH)3]n + 2n(CH3CO)2O ──(H2SO4 đặc, to)──► [C6H7O2(OH)(OCOCH3)2]n + 2nCH3COOH   (2)"
    ],
    ("C2", 107): [
        "Trong công nghiệp sản xuất gương soi và ruột phích, thay vì dùng dung dịch aldehyde độc hại, người ta thường sử dụng phản ứng tráng bạc từ nguồn nguyên liệu an toàn hơn là saccharose. Để sản xuất một lô gồm 100 chiếc gương để bàn trang điểm (kích thước 60 cm × 120 cm), người ta cần phủ lên mặt 1 chiếc gương với lớp bạc có khối lượng 0,864 gam. Quy trình được thực hiện bằng cách thủy phân m gam saccharose trong môi trường acid, sau đó trung hòa rồi cho toàn bộ sản phẩm tác dụng với lượng dư dung dịch AgNO3 trong NH3, đun nóng nhẹ. Giả thiết hiệu suất của toàn bộ quá trình đạt 80%. Tính giá trị của m. (Không làm tròn kết quả ở các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng phần mười)."
    ],
    ("C2", 112): [
        "Cho các phương trình hóa học của các phản ứng xảy ra được đánh số thứ tự từ (1) đến (4) dưới đây:",
        "(1) C6H12O6 ──(enzyme)──► 2CH3CH(OH)COOH",
        "(2) CH2OH[CHOH]4CHO + 2[Ag(NH3)2]OH ──(to)──► CH2OH[CHOH]4COONH4 + 2Ag↓ + 3NH3 + H2O",
        "(3) C12H22O11 + H2O ──(H+, to)──► C6H12O6 (glucose) + C6H12O6 (fructose)",
        "(4) [C6H7O2(OH)3]n + 3nHNO3 ──(H2SO4 đặc, to)──► [C6H7O2(ONO2)3]n + 3nH2O",
        "Hãy gán số thứ tự phương trình hóa học ở trên tương ứng với tên gọi: phản ứng thủy phân, phản ứng lên men, phản ứng với nitric acid, phản ứng với thuốc thử Tollens thành một dãy bốn số (ví dụ: 1234, 4321,...)."
    ],
    ("C2", 113): [
        "Cho phương trình hóa học của các phản ứng được đánh số thứ tự từ (1) tới (4) dưới đây:",
        "(1) CH3–CH=O + 2[Ag(NH3)2]OH ──(to)──► CH3COONH4 + 2Ag↓ + 3NH3 + H2O",
        "(2) CH3–CH=O + 3I2 + 4NaOH ──► CHI3↓ + HCOONa + 3NaI + 3H2O",
        "(3) C6H12O6 ──(enzyme)──► 2C2H5OH + 2CO2↑",
        "(4) (C6H10O5)n + nH2O ──(H+, to)──► nC6H12O6",
        "Gán số thứ tự phương trình hóa học của các phản ứng theo tên gọi: Phản ứng với thuốc thử Tollens, phản ứng lên men rượu, phản ứng tạo iodoform, phản ứng thủy phân và sắp xếp theo trình tự thành dãy bốn số (ví dụ: 1234, 4321,...)."
    ],
    ("C2", 116): [
        "Một nhóm học sinh tiến hành thí nghiệm thủy phân hoàn toàn tinh bột để điều chế glucose, sau đó thử tính chất của glucose bằng Cu(OH)2 trong môi trường kiềm. Để kiểm tra tiến trình của phản ứng thủy phân, nhóm học sinh sử dụng thêm dung dịch iodine trong KI. Các thao tác thí nghiệm được tiến hành như sau:",
        "• Bước 1: Cho vào ống nghiệm (a) dung dịch hồ tinh bột 2% và dung dịch H2SO4 10%, lắc đều rồi đặt ống nghiệm trong nồi nước sôi. Chuẩn bị ống nghiệm (b) chứa dung dịch I2 trong KI.",
        "• Bước 2: Sau khoảng 20 phút, hút lấy 1 – 2 giọt dung dịch trong ống nghiệm (a) đã để nguội nhỏ vào ống nghiệm (b) thì thấy không xuất hiện màu xanh tím.",
        "• Bước 3: Lấy khoảng 1 mL dung dịch đã thủy phân hoàn toàn sang ống nghiệm (c), thêm dần từng giọt dung dịch NaOH 10% cho đến môi trường kiềm, sau đó thêm vài giọt dung dịch CuSO4 5%, lắc đều rồi đun nóng."
    ],
    ("C2", 117): [
        "Trong phòng thí nghiệm, để khảo sát phản ứng thủy phân tinh bột, một nhóm học sinh chuẩn bị ba ống nghiệm, mỗi ống chứa khoảng 2 mL dung dịch hồ tinh bột 1% (đánh số 1, 2, 3) và tiến hành thí nghiệm:",
        "• Bước 1: Thêm khoảng 1 mL dung dịch HCl 1 M vào ống nghiệm (1); 1 mL dung dịch HCl 2 M vào ống nghiệm (2); 1 mL nước cất vào ống nghiệm (3).",
        "• Bước 2: Lắc đều, đặt cả 3 ống nghiệm trong nồi nước nóng, đun cách thủy trong 10 phút, sau đó để nguội.",
        "• Bước 3: Chia dung dịch trong mỗi ống nghiệm thành hai phần: Phần 1 nhỏ dung dịch iodine vào; Phần 2 trung hòa kiềm hóa bằng dung dịch NaOH 1 M rồi đun cách thủy với thuốc thử Tollens. Kết quả thu được ở bảng dưới đây:"
    ],
    ("C3", 129): [
        "Trong phòng thí nghiệm, aniline (C6H5NH2) được điều chế từ nitrobenzene (C6H5NO2) theo sơ đồ sau:",
        "C6H5NO2 ──(1) + Fe + HCl, to──► C6H5NH3Cl ──(2) + dd NaOH──► C6H5NH2",
        "Trong một thí nghiệm, sử dụng 6,15 gam nitrobenzene để điều chế aniline theo sơ đồ trên, thu được 3,72 gam aniline tinh khiết. Cho biết số sóng hấp thụ đặc trưng của một số nhóm chức trên phổ hồng ngoại như sau:"
    ],
    ("C3", 137): [
        "Benzocaine (H2N–C6H4–COOC2H5, M = 165) là chất gây tê cục bộ, được dùng làm hoạt chất trong nhiều loại thuốc mỡ gây tê. Benzocaine được tổng hợp từ nguyên liệu ban đầu là p-toluidine (CH3–C6H4–NH2, M = 107) qua hai giai đoạn chính theo sơ đồ chuyển hóa sau:",
        "p-CH3–C6H4–NH2 (toluidine) ──(1) oxi hóa──► p-HOOC–C6H4–NH2 (p-aminobenzoic acid) ──(2) + C2H5OH (H2SO4 đặc, to)──► p-H2N–C6H4–COOC2H5 (benzocaine)"
    ],
}

P2_ITEM_PATCHES = {
    ("C2", 92): {
        "a": "Nhóm –OH (*) ở α-fructose được gọi là nhóm –OH hemiacetal và –OH (**) ở β-fructose được gọi là nhóm –OH hemiketal.",
        "b": "Ở điều kiện thường hoặc khi đun nóng, fructose đều bị oxi hóa bởi Cu(OH)2 trong môi trường kiềm.",
        "c": "Fructose có phản ứng tráng bạc do cả hai dạng α-fructose và β-fructose đều phản ứng trực tiếp với thuốc thử Tollens.",
        "d": "Dạng α-fructose và β-fructose có sự khác nhau về hướng trong không gian của các nhóm liên kết ở vị trí C2."
    },
    ("C2", 42): {
        "a": "Ở bước 1, xảy ra phản ứng tạo thành kết tủa copper(II) hydroxide có màu xanh.",
        "b": "Ở bước 2, kết tủa bị hòa tan hoàn toàn tạo thành dung dịch phức chất có màu xanh lam đặc trưng.",
        "c": "Hiện tượng ở bước 2 chứng minh phân tử glucose có nhiều nhóm hydroxy (–OH) liền kề nhau (tính chất của polyalcohol).",
        "d": "Ở bước 3, khi đun nóng xuất hiện kết tủa màu đỏ gạch Cu2O; nếu thay dung dịch glucose bằng dung dịch fructose thì khi đun nóng ở bước 3 không thu được kết tủa đỏ gạch."
    },
    ("C2", 46): {
        "a": "Ở bước 2, sau khi thêm dung dịch NaOH 10% thì thu được kết tủa Cu(OH)2 màu xanh.",
        "b": "Sau bước 3, kết tủa Cu(OH)2 bị hòa tan tạo thành dung dịch màu xanh lam.",
        "c": "Thí nghiệm ở bước 3 dùng để chứng minh glucose có tính chất của nhóm chức aldehyde (–CHO).",
        "d": "Ở bước 3, nếu thay dung dịch glucose bằng dung dịch saccharose hoặc dung dịch glycerol thì hiện tượng hòa tan kết tủa tạo dung dịch xanh lam vẫn xảy ra tương tự."
    },
    ("C2", 51): {
        "a": "Trong phản ứng (1), chất hữu cơ X thu được là ethyl alcohol (C2H5OH).",
        "b": "Trong phản ứng (2), hợp chất Z là lactic acid (CH3–CH(OH)–COOH) chứa đồng thời nhóm chức hydroxy (–OH) và nhóm carboxyl (–COOH).",
        "c": "Cả hai quá trình lên men (1) và (2) đều giải phóng khí carbon dioxide (CO2) thoát ra khỏi hỗn hợp phản ứng.",
        "d": "Nếu lên men 360 gam glucose theo phản ứng (1) với hiệu suất 80% thì thu được 147,2 gam ethyl alcohol."
    },
    ("C2", 70): {
        "a": "Ở bước 2, saccharose bị thủy phân hoàn toàn chỉ tạo thành một loại monosaccharide duy nhất là glucose.",
        "b": "Ở bước 3, vai trò của NaHCO3 là trung hòa hết lượng acid HCl dư sau phản ứng thủy phân.",
        "c": "Ở bước 4, hỗn hợp phản ứng ban đầu có màu xanh lam, sau khi đun nóng chuyển dần sang kết tủa màu đỏ gạch (Cu2O).",
        "d": "Nếu bỏ qua bước 1 và bước 2 mà cho trực tiếp dung dịch saccharose vào Cu(OH)2/OH- rồi đun nóng thì vẫn thu được kết tủa đỏ gạch Cu2O."
    },
    ("C2", 84): {
        "a": "Dung dịch Y tạo màu xanh tím đặc trưng với dung dịch iodine nên Y là hồ tinh bột.",
        "b": "Dung dịch T tham gia phản ứng tráng bạc với thuốc thử Tollens tạo kết tủa Ag nên T là glucose.",
        "c": "Dung dịch X hòa tan Cu(OH)2 ở điều kiện thường tạo dung dịch xanh lam nhưng không tráng bạc nên X là ethanol.",
        "d": "Có thể dùng một thuốc thử duy nhất là Cu(OH)2 trong môi trường kiềm (kết hợp đun nóng) để phân biệt ba dung dịch X, Z và T."
    },
    ("C2", 85): {
        "a": "Glucose và fructose đều là monosaccharide có cùng công thức phân tử C6H12O6 và đều không tham gia phản ứng thủy phân.",
        "b": "Trong môi trường kiềm, fructose và glucose có thể chuyển hóa qua lại lẫn nhau nên cả hai chất đều tham gia phản ứng tráng bạc với thuốc thử Tollens.",
        "c": "Có thể phân biệt dung dịch glucose và dung dịch fructose bằng nước bromine ở nhiệt độ thường.",
        "d": "Khi cho 18 gam hỗn hợp gồm glucose và fructose tác dụng với lượng dư thuốc thử Tollens đun nóng (hiệu suất 100%), thu được tối đa 10,8 gam kết tủa bạc."
    },
    ("C2", 86): {
        "a": "Ở điều kiện thường, saccharose là chất rắn kết tinh, không màu, có vị ngọt và dễ tan trong nước.",
        "b": "Phân tử saccharose được cấu tạo từ một đơn vị α-glucose và một đơn vị β-fructose liên kết với nhau qua nguyên tử oxygen nên không còn nhóm –OH hemiacetal tự do.",
        "c": "Dung dịch saccharose hòa tan được Cu(OH)2 ở điều kiện thường tạo dung dịch màu xanh lam và làm mất màu nước bromine.",
        "d": "Thủy phân hoàn toàn 34,2 gam saccharose trong môi trường acid, trung hòa acid rồi cho toàn bộ sản phẩm tác dụng với lượng dư thuốc thử Tollens đun nóng thì thu được tối đa 43,2 gam Ag."
    },
    ("C2", 100): {
        "a": "Trong mỗi mắt xích của cellulose có 3 nhóm hydroxy (–OH) tự do nên cellulose dễ dàng hòa tan Cu(OH)2 trong môi trường kiềm ở nhiệt độ thường tạo dung dịch màu xanh lam.",
        "b": "Phân tử khối của mỗi mắt xích cellulose là 162 và công thức phân tử của cellulose được viết gọn là (C6H10O5)n.",
        "c": "Bản chất của các phản ứng (1), (2) là nguyên tử H trong nhóm –OH của mắt xích cellulose được thay thế bằng nhóm acetyl (CH3CO–).",
        "d": "Khi phân tích thành phần nguyên tố của một mẫu cellulose acetate (chỉ gồm cellulose triacetate và cellulose diacetate), phần trăm khối lượng carbon là 49,785%; suy ra phần trăm khối lượng oxygen trong mẫu đó là 34,64%."
    },
    ("C2", 116): {
        "a": "Ở bước 1, dung dịch H2SO4 10% đóng vai trò là chất xúc tác cho phản ứng thủy phân tinh bột.",
        "b": "Ở bước 2, việc không xuất hiện màu xanh tím khi thử với dung dịch iodine chứng tỏ tinh bột trong ống nghiệm (a) đã bị thủy phân hoàn toàn.",
        "c": "Ở bước 3, mục đích thêm dung dịch NaOH 10% trước khi cho CuSO4 là để trung hòa hết acid H2SO4 dư và tạo môi trường kiềm.",
        "d": "Sau khi đun nóng ống nghiệm (c) ở bước 3, chỉ thu được dung dịch trong suốt màu xanh lam mà không có kết tủa xuất hiện."
    },
    ("C2", 117): {
        "a": "Ở ống nghiệm (3) chỉ thêm nước cất nên sau 10 phút đun nóng, tinh bột hầu như chưa bị thủy phân thành glucose.",
        "b": "Ở ống nghiệm (1), dung dịch sau thủy phân vừa tạo màu xanh tím với iodine vừa tráng bạc, chứng tỏ tinh bột đã bị thủy phân một phần nhưng chưa hoàn toàn.",
        "c": "Ở ống nghiệm (2), dung dịch sau thủy phân không đổi màu với iodine và có phản ứng tráng bạc, chứng tỏ toàn bộ tinh bột đã bị thủy phân hết.",
        "d": "Từ kết quả thí nghiệm ở ba ống nghiệm, kết luận được rằng nồng độ acid không ảnh hưởng đến tốc độ phản ứng thủy phân tinh bột."
    },
}

ANSWER_PATCHES = {
    ("C1", 47): "D",
    ("C2", 42): "a:Đ - b:Đ - c:Đ - d:S",
    ("C2", 46): "a:Đ - b:Đ - c:S - d:Đ",
    ("C2", 51): "a:Đ - b:Đ - c:S - d:Đ",
    ("C2", 70): "a:S - b:Đ - c:Đ - d:S",
    ("C2", 84): "a:Đ - b:Đ - c:S - d:Đ",
    ("C2", 85): "a:Đ - b:Đ - c:Đ - d:S",
    ("C2", 86): "a:Đ - b:Đ - c:S - d:Đ",
    ("C2", 116): "a:Đ - b:Đ - c:Đ - d:S",
    ("C2", 117): "a:Đ - b:Đ - c:Đ - d:S",
}

SOLUTION_PATCHES = {
    ("C2", 42): (
        "a) Đúng: 2NaOH + CuSO4 → Cu(OH)2↓ (xanh) + Na2SO4.\n"
        "b) Đúng: Glucose hòa tan Cu(OH)2 ở nhiệt độ thường tạo phức đồng-glucose [Cu(C6H11O6)2] màu xanh lam.\n"
        "c) Đúng: Phản ứng hòa tan Cu(OH)2 ở nhiệt độ thường chứng minh glucose có nhiều nhóm –OH liền kề.\n"
        "d) Sai: Trong môi trường kiềm đun nóng, fructose chuyển hóa thành glucose nên vẫn tạo kết tủa đỏ gạch Cu2O."
    ),
    ("C2", 46): (
        "a) Đúng: CuSO4 + 2NaOH → Cu(OH)2↓ (xanh) + Na2SO4.\n"
        "b) Đúng: Glucose hòa tan Cu(OH)2 ở nhiệt độ thường tạo dung dịch phức màu xanh lam.\n"
        "c) Sai: Ở nhiệt độ thường, thí nghiệm này chứng minh tính chất polyalcohol (nhiều nhóm –OH liền kề) của glucose, chưa chứng minh tính chất aldehyde.\n"
        "d) Đúng: Saccharose và glycerol đều có các nhóm –OH liền kề nên đều hòa tan Cu(OH)2 ở nhiệt độ thường tạo dung dịch xanh lam."
    ),
    ("C2", 51): (
        "a) Đúng: Lên men rượu glucose tạo X là C2H5OH (ethyl alcohol) và CO2.\n"
        "b) Đúng: Lên men lactic glucose tạo Z là CH3–CH(OH)–COOH (lactic acid), là hợp chất tạp chức.\n"
        "c) Sai: Chỉ quá trình lên men rượu (1) mới sinh ra khí CO2; quá trình lên men lactic (2) không sinh ra khí CO2.\n"
        "d) Đúng: n(glucose) = 360/180 = 2 mol → n(C2H5OH) = 2 × 2 × 80% = 3,2 mol → m(C2H5OH) = 3,2 × 46 = 147,2 gam."
    ),
    ("C2", 70): (
        "a) Sai: Thủy phân saccharose thu được hai monosaccharide là glucose và fructose.\n"
        "b) Đúng: NaHCO3 tác dụng với HCl dư (NaHCO3 + HCl → NaCl + CO2↑ + H2O) để trung hòa môi trường acid.\n"
        "c) Đúng: Sản phẩm thủy phân hòa tan Cu(OH)2 tạo dung dịch xanh lam, khi đun nóng nhóm –CHO khử Cu(OH)2 thành Cu2O↓ đỏ gạch.\n"
        "d) Sai: Saccharose không có nhóm –CHO (không có tính khử) nên nếu chưa thủy phân sẽ không tạo kết tủa đỏ gạch Cu2O khi đun nóng."
    ),
    ("C2", 84): (
        "a) Đúng: Hồ tinh bột (Y) hấp phụ iodine tạo hợp chất bọc có màu xanh tím.\n"
        "b) Đúng: Glucose (T) vừa hòa tan Cu(OH)2 tạo dung dịch xanh lam vừa tham gia phản ứng tráng bạc.\n"
        "c) Sai: Chất X hòa tan Cu(OH)2 tạo dung dịch xanh lam và không tráng bạc nên X là glycerol, còn Z mới là ethanol.\n"
        "d) Đúng: Dùng Cu(OH)2/OH- ở nhiệt độ thường nhận ra ethanol (Z, không hòa tan Cu(OH)2); đun nóng hai mẫu xanh lam còn lại thì glucose (T) tạo kết tủa đỏ gạch Cu2O, còn glycerol (X) không tạo kết tủa đỏ gạch."
    ),
    ("C2", 85): (
        "a) Đúng: Glucose và fructose là hai monosaccharide đồng phân (C6H12O6), không bị thủy phân.\n"
        "b) Đúng: Trong môi trường kiềm của thuốc thử Tollens, fructose chuyển hóa thuận nghịch thành glucose nên cả hai đều tráng bạc.\n"
        "c) Đúng: Glucose có nhóm –CHO làm mất màu nước bromine, còn fructose không làm mất màu nước bromine.\n"
        "d) Sai: n(hỗn hợp) = 18/180 = 0,1 mol → n(Ag) = 2 × 0,1 = 0,2 mol → m(Ag) = 0,2 × 108 = 21,6 gam (chứ không phải 10,8 gam)."
    ),
    ("C2", 86): (
        "a) Đúng: Saccharose là chất rắn kết tinh, không màu, vị ngọt, tan tốt trong nước.\n"
        "b) Đúng: Phân tử saccharose tạo bởi 1 gốc α-glucose và 1 gốc β-fructose qua liên kết α-1,2-glycoside nên không còn nhóm –OH hemiacetal.\n"
        "c) Sai: Saccharose hòa tan Cu(OH)2 tạo dung dịch xanh lam nhưng không có tính khử nên không làm mất màu nước bromine.\n"
        "d) Đúng: n(saccharose) = 34,2/342 = 0,1 mol → Thủy phân tạo 0,1 mol glucose và 0,1 mol fructose → n(Ag) = 4 × 0,1 = 0,4 mol → m(Ag) = 0,4 × 108 = 43,2 gam."
    ),
    ("C2", 116): (
        "a) Đúng: H2SO4 loãng là chất xúc tác cho phản ứng thủy phân tinh bột.\n"
        "b) Đúng: Không còn màu xanh tím với iodine chứng tỏ tinh bột đã thủy phân hết.\n"
        "c) Đúng: Phải trung hòa hết acid dư và kiềm hóa môi trường trước khi tạo kết tủa Cu(OH)2.\n"
        "d) Sai: Khi đun nóng, glucose sinh ra sẽ khử Cu(OH)2 trong môi trường kiềm tạo kết tủa Cu2O màu đỏ gạch."
    ),
    ("C2", 117): (
        "a) Đúng: Không có xúc tác acid, tinh bột hầu như không bị thủy phân khi đun nóng trong 10 phút.\n"
        "b) Đúng: Còn phản ứng màu xanh tím với I2 chứng tỏ còn tinh bột dư; có phản ứng Tollens chứng tỏ đã có glucose tạo thành.\n"
        "c) Đúng: Không còn màu xanh tím với I2 chứng tỏ tinh bột đã thủy phân hoàn toàn.\n"
        "d) Sai: So sánh ống (1) dùng HCl 1M và ống (2) dùng HCl 2M cho thấy nồng độ acid lớn hơn làm phản ứng thủy phân xảy ra nhanh và hoàn toàn hơn."
    ),
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
    # ================= ĐỀ SỐ 06 ĐẾN ĐỀ SỐ 10 (MÃ ĐỀ 106 - 110) =================
    {
        "de_num": 6,
        "ma_de": "106",
        "P1": [
            ("C1", 1), ("C1", 11), ("C1", 31), ("C1", 45), ("C1", 40), ("C1", 50), ("C1", 75), ("C1", 96),
            ("C2", 10), ("C2", 49), ("C2", 4), ("C2", 57), ("C2", 6), ("C2", 32), ("C2", 78),
            ("C3", 2), ("C3", 13), ("C3", 64)
        ],
        "P2": [("C1", 107), ("C2", 42), ("C2", 100), ("C3", 129)],
        "P3": [("C1", 152), ("C1", 154), ("C1", 181), ("C2", 103), ("C2", 104), ("C3", 152)],
    },
    {
        "de_num": 7,
        "ma_de": "107",
        "P1": [
            ("C1", 3), ("C1", 16), ("C1", 36), ("C1", 73), ("C1", 77), ("C1", 56), ("C1", 84), ("C1", 95),
            ("C2", 12), ("C2", 40), ("C2", 55), ("C2", 59), ("C2", 23), ("C2", 72), ("C2", 79),
            ("C3", 6), ("C3", 15), ("C3", 65)
        ],
        "P2": [("C1", 111), ("C2", 46), ("C2", 70), ("C3", 133)],
        "P3": [("C1", 160), ("C1", 155), ("C1", 183), ("C2", 106), ("C2", 105), ("C3", 164)],
    },
    {
        "de_num": 8,
        "ma_de": "108",
        "P1": [
            ("C1", 7), ("C1", 18), ("C1", 63), ("C1", 74), ("C1", 78), ("C1", 92), ("C1", 32), ("C1", 99),
            ("C2", 15), ("C2", 41), ("C2", 61), ("C2", 60), ("C2", 24), ("C2", 73), ("C2", 80),
            ("C3", 8), ("C3", 18), ("C3", 66)
        ],
        "P2": [("C1", 134), ("C2", 51), ("C2", 86), ("C3", 134)],
        "P3": [("C1", 161), ("C1", 156), ("C1", 184), ("C2", 111), ("C2", 107), ("C3", 166)],
    },
    {
        "de_num": 9,
        "ma_de": "109",
        "P1": [
            ("C1", 23), ("C1", 25), ("C1", 66), ("C1", 76), ("C1", 80), ("C1", 28), ("C1", 34), ("C1", 101),
            ("C2", 16), ("C2", 45), ("C2", 65), ("C2", 67), ("C2", 26), ("C2", 75), ("C2", 81),
            ("C3", 9), ("C3", 19), ("C3", 105)
        ],
        "P2": [("C1", 140), ("C2", 84), ("C2", 116), ("C3", 135)],
        "P3": [("C1", 165), ("C1", 158), ("C1", 186), ("C2", 112), ("C2", 124), ("C3", 167)],
    },
    {
        "de_num": 10,
        "ma_de": "110",
        "P1": [
            ("C1", 24), ("C1", 27), ("C1", 72), ("C1", 83), ("C1", 81), ("C1", 30), ("C1", 35), ("C1", 103),
            ("C2", 33), ("C2", 47), ("C2", 66), ("C2", 68), ("C2", 30), ("C2", 83), ("C2", 82),
            ("C3", 11), ("C3", 20), ("C3", 110)
        ],
        "P2": [("C1", 148), ("C2", 85), ("C2", 117), ("C3", 137)],
        "P3": [("C1", 169), ("C1", 173), ("C1", 187), ("C2", 113), ("C2", 125), ("C3", 171)],
    },
]


def parse_abcd_from_opt_lines(opt_lines):
    """Tách 4 phương án A, B, C, D từ danh sách dòng phương án."""
    joined = "\t".join(opt_lines)
    matches = list(re.finditer(r'(?:^|(?<=[\t\s]))([A-D])\.\s*', joined))
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
        if (ch_code, orig_qnum) in OPT_PATCHES:
            q["opt_lines"] = OPT_PATCHES[(ch_code, orig_qnum)]
        if (ch_code, orig_qnum) in P2_ITEM_PATCHES:
            q["p2_items"] = P2_ITEM_PATCHES[(ch_code, orig_qnum)]
        if (ch_code, orig_qnum) in ANSWER_PATCHES:
            q["answer"] = ANSWER_PATCHES[(ch_code, orig_qnum)]
        if (ch_code, orig_qnum) in SOLUTION_PATCHES:
            q["solution"] = SOLUTION_PATCHES[(ch_code, orig_qnum)]
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
        c_int = int(ch_code[1])
        q_assets = assets_map.get(c_int, {}).get(orig_qnum, {"images": [], "tables": []})
        img_placed_early = False
        img_triggers = ("như sau:", "dưới đây:", "hình vẽ sau:", "hình sau:", "sơ đồ sau:")

        if stem_lines:
            add_chem_runs(p_q, stem_lines[0], size=11.5)
            if any(trig in stem_lines[0].lower() for trig in img_triggers) and q_assets["images"]:
                for img_path in q_assets["images"]:
                    if os.path.exists(img_path):
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.paragraph_format.space_before = Pt(2)
                        p_img.paragraph_format.space_after = Pt(2)
                        p_img.add_run().add_picture(img_path, width=Cm(6.5))
                img_placed_early = True
            for extra_ln in stem_lines[1:]:
                p_ex = doc.add_paragraph()
                p_ex.paragraph_format.space_before = Pt(1)
                p_ex.paragraph_format.space_after = Pt(1)
                add_chem_runs(p_ex, extra_ln, size=11.5)

        # Insert any extracted images or data tables for this question
        if not img_placed_early:
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
        if (ch_code, orig_qnum) in ANSWER_PATCHES:
            q["answer"] = ANSWER_PATCHES[(ch_code, orig_qnum)]
        if (ch_code, orig_qnum) in SOLUTION_PATCHES:
            q["solution"] = SOLUTION_PATCHES[(ch_code, orig_qnum)]
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
        if (ch_code, orig_qnum) in ANSWER_PATCHES:
            q["answer"] = ANSWER_PATCHES[(ch_code, orig_qnum)]
        if (ch_code, orig_qnum) in SOLUTION_PATCHES:
            q["solution"] = SOLUTION_PATCHES[(ch_code, orig_qnum)]
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

    # Kiểm tra 280 câu hỏi trong trọn bộ 10 đề hoàn toàn không trùng nhau
    all_used = []
    for t in TESTS_SPEC:
        for part_k in ["P1", "P2", "P3"]:
            for pair in t[part_k]:
                all_used.append(pair)
    print(f"Tổng số câu hỏi sử dụng trong 10 đề: {len(all_used)} câu | Số câu duy nhất (không trùng): {len(set(all_used))} câu.")
    assert len(all_used) == len(set(all_used)) == 280, "Có câu hỏi trùng lặp!"

    print("Đang trích xuất hình ảnh và bảng dữ liệu từ Ngân hàng...")
    assets_map = {c: extract_chapter_assets(c) for c in [1, 2, 3]}

    for t_cfg in TESTS_SPEC:
        print(f"\nĐang tạo Đề số {t_cfg['de_num']:02d} (Mã đề {t_cfg['ma_de']})...")
        generate_single_test_docx(t_cfg, bank_map, assets_map, is_teacher=False)
        generate_single_test_docx(t_cfg, bank_map, assets_map, is_teacher=True)

    print("\nHOÀN TẤT 20 FILE WORD CHO TRỌN BỘ 10 ĐỀ KIỂM TRA GIỮA KÌ 1 HÓA 12!")
