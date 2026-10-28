# -*- coding: utf-8 -*-
"""
Script biên soạn trọn bộ 3 file:
1. Đề cương Học sinh (.docx)
2. Đề cương Giáo viên có lời giải chi tiết (.docx)
3. Slide Dạy trên màn hình Tivi 65 inch tỉ lệ 16:9 (.pdf và .docx) - KHÔNG CÓ ĐÁP ÁN

ĐỀ CƯƠNG GIỮA KÌ 1. NĂM HỌC 2026 - 2027
MÔN HỌC: KHTN 9 – PHÂN MÔN VẬT SỐNG (SINH HỌC)
TRƯỜNG THCS NGUYỄN HOÀNG
"""

import sys
if sys.stdout:
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr:
    sys.stderr.reconfigure(encoding='utf-8')

import os
import shutil
import subprocess
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# ==============================================================================
# HẰNG SỐ MÀU SẮC & ĐỊNH DẠNG THEO GEMINI.md (FORM ĐỀ KIỂM TRA)
# ==============================================================================
COLOR_PRIMARY_HEX = "D32F2F"       # Đỏ san hô (Coral Red)
COLOR_SECONDARY_HEX = "8BC390"     # Xanh xô thơm (Sage Green)
COLOR_TEXT_HEX = "1A1A1A"          # Đen than nhẹ (Off-black)
COLOR_SOLUTION_TITLE_HEX = "1565C0"# Xanh dương đậm cho tiêu đề lời giải
COLOR_GRAY_HEX = "757575"

COLOR_PRIMARY = RGBColor(0xD3, 0x2F, 0x2F)
COLOR_SECONDARY = RGBColor(0x8B, 0xC3, 0x90)
COLOR_TEXT = RGBColor(0x1A, 0x1A, 0x1A)
COLOR_SOLUTION_TITLE = RGBColor(0x15, 0x65, 0xC0)
COLOR_GRAY = RGBColor(0x75, 0x75, 0x75)

FONT_NAME = "Times New Roman"

BASE_DIR = r"c:\Antigravity_Thanh"
SAN_PHAM_DIR = os.path.join(BASE_DIR, "San_Pham")
SCRATCH_DIR = os.path.join(BASE_DIR, "scratch")
IMG_DIR = os.path.join(SAN_PHAM_DIR, "images_khtn9_nguyenhoang")
DIAGRAM_PATH = os.path.join(IMG_DIR, "so_do_tai_ban_dna.png")

ONEDRIVE_DIR1 = r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9\KHTN9-DEKIEMTRADINHKI"
ONEDRIVE_DIR2 = r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9"

os.makedirs(SAN_PHAM_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)
os.makedirs(IMG_DIR, exist_ok=True)

# ==============================================================================
# DỮ LIỆU ĐỀ CƯƠNG KHTN 9 (PHÂN MÔN VẬT SỐNG) - THCS NGUYỄN HOÀNG
# ==============================================================================

MCQ_DATA = [
    {
        "id": 1,
        "question": "Hiện tượng di truyền là gì?",
        "options": [
            "A. Sự khác nhau giữa các cá thể cùng loài.",
            "B. Sự truyền đạt các tính trạng của bố mẹ, tổ tiên cho các thế hệ con cháu.",
            "C. Sự giống nhau của 2 anh em cùng 1 nhà.",
            "D. Sự xuất hiện 1 số tính trạng mới cả con so với cha mẹ."
        ],
        "answer": "B",
        "solution": "Di truyền là hiện tượng truyền đạt các tính trạng của bố mẹ, tổ tiên cho các thế hệ con cháu. Do đó chọn phương án B."
    },
    {
        "id": 2,
        "question": "Theo ý tưởng của Mendel, yếu tố nào quy định sự di truyền của tính trạng?",
        "options": [
            "A. Các chất dinh dưỡng trong đất.",
            "B. Ánh sáng mà cây nhận được.",
            "C. Các nhân tố di truyền.",
            "D. Lượng nước cây hấp thụ."
        ],
        "answer": "C",
        "solution": "Theo Mendel, mỗi tính trạng của cơ thể do một cặp nhân tố di truyền (sau này được gọi là gene) quy định. Do đó chọn phương án C."
    },
    {
        "id": 3,
        "question": "Hai loại nucleic acid là",
        "options": [
            "A. DNA và RNA.",
            "B. DNA và protein.",
            "C. RNA và lipid.",
            "D. Protein và lipid."
        ],
        "answer": "A",
        "solution": "Trong tế bào có hai loại nucleic acid chính là DNA (deoxyribonucleic acid) và RNA (ribonucleic acid). Do đó chọn phương án A."
    },
    {
        "id": 4,
        "question": "Chức năng của DNA là",
        "options": [
            "A. Cung cấp trực tiếp năng lượng cho mọi hoạt động của tế bào.",
            "B. Vận chuyển oxygen trong máu.",
            "C. Phân giải thức ăn thành chất đơn giản.",
            "D. Lưu giữ, bảo quản và truyền đạt thông tin di truyền."
        ],
        "answer": "D",
        "solution": "DNA có chức năng cơ bản và quan trọng nhất là lưu giữ, bảo quản và truyền đạt thông tin di truyền qua các thế hệ. Do đó chọn phương án D."
    },
    {
        "id": 5,
        "question": "Hiện tượng biến dị là gì?",
        "options": [
            "A. Sự truyền đạt nguyên vẹn mọi tính trạng cho đời con.",
            "B. Sự sai khác toàn bộ các đặc điểm của con cái so với cha mẹ.",
            "C. Hiện tượng con sinh ra có những đặc điểm khác nhau và khác với bố mẹ.",
            "D. Sự khác nhau giữa 2-4 thành viên trong 1 gia đình (cha, mẹ và 2 con)."
        ],
        "answer": "C",
        "solution": "Biến dị là hiện tượng con sinh ra có những đặc điểm khác nhau và khác so với bố mẹ, tổ tiên. Do đó chọn phương án C."
    },
    {
        "id": 6,
        "question": "Ý tưởng nào của Mendel đặt cơ sở cho các nghiên cứu về gene?",
        "options": [
            "A. Tính trạng được di truyền nhờ các nhân tố di truyền.",
            "B. Tính trạng chỉ phụ thuộc vào lượng thức ăn.",
            "C. Tính trạng của bố mẹ luôn hòa trộn vĩnh viễn ở con.",
            "D. Tính trạng ở con không liên quan đến bố mẹ."
        ],
        "answer": "A",
        "solution": "Ý tưởng về 'nhân tố di truyền' tồn tại riêng rẽ trong tế bào, không hòa trộn vào nhau và được truyền đạt qua các thế hệ là nền móng trực tiếp dẫn tới khái niệm gene trong di truyền học hiện đại. Do đó chọn phương án A."
    },
    {
        "id": 7,
        "question": "Nhóm nào gồm đầy đủ hai loại nucleic acid?",
        "options": [
            "A. RNA và carbohydrate.",
            "B. Protein và DNA.",
            "C. Lipid và carbohydrate.",
            "D. RNA và DNA."
        ],
        "answer": "D",
        "solution": "Nhóm chứa đầy đủ hai loại nucleic acid là RNA và DNA. Do đó chọn phương án D."
    },
    {
        "id": 8,
        "question": "Vai trò nào sau đây thuộc về DNA?",
        "options": [
            "A. Xúc tác mọi phản ứng hóa học trong tế bào.",
            "B. Lưu giữ, bảo quản và truyền đạt thông tin di truyền.",
            "C. Tạo ra năng lượng trực tiếp cho sự co cơ.",
            "D. Vận chuyển các chất dinh dưỡng trong máu."
        ],
        "answer": "B",
        "solution": "Vai trò sinh học của DNA là lưu giữ, bảo quản và truyền đạt thông tin di truyền. Do đó chọn phương án B."
    }
]

SHORT_ANS_DATA = [
    {
        "sub_id": "a",
        "intro": "Khi giải thích vì sao chỉ từ bốn loại nucleotide có thể tạo nên nhiều phân tử DNA khác nhau, có các nhận định sau:",
        "statements": [
            "(1) Hai DNA cùng số nucleotide nhưng khác trình tự sắp xếp có thể khác nhau.",
            "(2) Hai DNA có cùng số lượng từng loại nucleotide thì chắc chắn có trình tự giống nhau.",
            "(3) Sự khác nhau về số lượng, thành phần và trình tự nucleotide góp phần tạo nên tính đa dạng của DNA.",
            "(4) Vì chỉ có bốn loại nucleotide nên chỉ có thể tạo ra bốn loại phân tử DNA.",
            "(5) Chỉ cần biết tỉ lệ các loại nucleotide là xác định được duy nhất trình tự của một DNA."
        ],
        "question": "Có bao nhiêu nhận định đúng? Trả lời: ................",
        "answer": "2",
        "solution": (
            "Phân tích từng nhận định:\n"
            "• (1) *ĐÚNG*: Hai phân tử DNA có cùng chiều dài/số nucleotide nhưng trật tự sắp xếp khác nhau sẽ mang thông tin di truyền khác nhau.\n"
            "• (2) *SAI*: Cùng số lượng từng loại nucleotide nhưng có vô số cách sắp xếp trật tự khác nhau, nên trình tự không nhất thiết phải giống nhau.\n"
            "• (3) *ĐÚNG*: Tính đa dạng và đặc thù của DNA được quy định bởi số lượng, thành phần và đặc biệt là trật tự sắp xếp của các nucleotide.\n"
            "• (4) *SAI*: Tuy chỉ có 4 loại nucleotide (A, T, G, C) nhưng với kích thước hàng triệu đến hàng trăm triệu cặp base và vô số cách sắp xếp, tạo ra số lượng phân tử DNA vô cùng phong phú.\n"
            "• (5) *SAI*: Tỉ lệ nucleotide không quyết định được trình tự sắp xếp cụ thể của từng vị trí base trên mạch DNA.\n"
            "==> Có *2* nhận định đúng là (1) và (3)."
        )
    },
    {
        "sub_id": "b",
        "intro": "Quan sát sơ đồ một lần tái bản DNA dưới đây và xét các nhận định sau (coi quá trình tái bản diễn ra chính xác):",
        "has_image": True,
        "image_caption": "Sơ đồ giản lược một lần tái bản DNA (Nét liền: mạch của DNA mẹ; nét đứt: mạch mới).",
        "statements": [
            "(1) Hai mạch DNA mẹ tách nhau để làm khuôn tổng hợp các mạch mới.",
            "(2) Nucleotide tự do ghép với nucleotide trên mạch khuôn theo nguyên tắc bổ sung.",
            "(3) Hai DNA con đều chỉ gồm các mạch mới được tổng hợp.",
            "(4) Khi tái bản chính xác, thông tin di truyền của DNA mẹ được truyền cho cả hai DNA con.",
            "(5) Mỗi DNA con trong sơ đồ có một mạch từ DNA mẹ và một mạch mới.",
            "(6) Mạch mới được tổng hợp trước khi hai mạch DNA mẹ tách nhau.",
            "(7) Nucleotide trên mạch mới được lắp ghép theo nguyên tắc bổ sung với mạch khuôn.",
            "(8) Tái bản chính xác giúp duy trì thông tin di truyền của DNA mẹ trong hai DNA con."
        ],
        "question": "Có bao nhiêu nhận định đúng? Trả lời: ................",
        "answer": "6",
        "solution": (
            "Phân tích từng nhận định đối chiếu với cơ chế tái bản (nhân đôi) DNA:\n"
            "• (1) *ĐÚNG*: Quá trình tái bản bắt đầu bằng việc tách hai mạch đơn của DNA mẹ làm mạch khuôn.\n"
            "• (2) *ĐÚNG*: Các nucleotide tự do trong môi trường nội bào liên kết với các nucleotide trên mạch khuôn theo nguyên tắc bổ sung (A-T, G-C).\n"
            "• (3) *SAI*: Theo nguyên tắc bán bảo toàn, mỗi DNA con gồm một mạch của DNA mẹ và một mạch mới tổng hợp, chứ không phải 'chỉ gồm các mạch mới'.\n"
            "• (4) *ĐÚNG*: Tái bản chính xác đảm bảo thông tin di truyền từ DNA mẹ được truyền nguyên vẹn cho cả 2 DNA con.\n"
            "• (5) *ĐÚNG*: Mỗi DNA con chứa 1 mạch cũ của DNA mẹ (nét liền) và 1 mạch mới tổng hợp (nét đứt) - nguyên tắc bán bảo toàn.\n"
            "• (6) *SAI*: Hai mạch DNA mẹ phải tách nhau ra trước làm khuôn thì mạch mới mới có thể tổng hợp được.\n"
            "• (7) *ĐÚNG*: Mạch mới được kéo dài và lắp ghép chính xác theo nguyên tắc bổ sung với mạch khuôn.\n"
            "• (8) *ĐÚNG*: Quá trình nhân đôi chính xác giúp duy trì tính ổn định của thông tin di truyền qua các thế hệ tế bào.\n"
            "==> Các nhận định đúng gồm: (1), (2), (4), (5), (7), (8). Có *6* nhận định đúng."
        )
    }
]

ESSAY_DATA = [
    {
        "id": 20,
        "title": "Câu 20 (Phần Tự luận)",
        "content": (
            "Trong thí nghiệm của Mendel, lai đậu Hà Lan hạt vàng thuần chủng với đậu hạt xanh thuần chủng thu được F1 toàn hạt vàng; "
            "cho F1 tự thụ phấn, F2 có tỉ lệ xấp xỉ 3 hạt vàng : 1 hạt xanh. Biện luận viết sơ đồ lai từ P đến F2."
        ),
        "solution": (
            "1. Biện luận xác định tính trạng trội, lặn và quy ước gene:\n"
            "• P thuần chủng tương phản: Hạt vàng × Hạt xanh thu được F1 100% hạt vàng.\n"
            "  Theo quy luật phân li của Mendel, tính trạng biểu hiện ngay ở thế hệ lai F1 là tính trạng trội, tính trạng lặn bị lấn át.\n"
            "  ==> Hạt vàng là tính trạng *trội hoàn toàn* so với hạt xanh.\n"
            "• Quy ước gene:\n"
            "  - Gene *A*: quy định hạt vàng (trội).\n"
            "  - Gene *a*: quy định hạt xanh (lặn).\n"
            "• Xác định kiểu gene của bố mẹ P:\n"
            "  - Đậu hạt vàng thuần chủng có kiểu gene là *AA*.\n"
            "  - Đậu hạt xanh (tính trạng lặn) có kiểu gene là *aa*.\n\n"
            "2. Sơ đồ lai từ P đến F2:\n"
            "• Phép lai P:\n"
            "  P_tc:             AA (Hạt vàng)        ×        aa (Hạt xanh)\n"
            "  G_P:               A                               a\n"
            "  F1:                               Aa (100% Hạt vàng)\n\n"
            "• F1 tự thụ phấn (F1 × F1):\n"
            "  F1 × F1:          Aa (Hạt vàng)        ×        Aa (Hạt vàng)\n"
            "  G_F1:          1/2 A : 1/2 a                   1/2 A : 1/2 a\n"
            "  F2:\n"
            "  - Tỉ lệ kiểu gene: *1 AA : 2 Aa : 1 aa* (hay 25% AA : 50% Aa : 25% aa).\n"
            "  - Tỉ lệ kiểu hình: *3 Hạt vàng : 1 Hạt xanh* (hay 75% Hạt vàng : 25% Hạt xanh).\n"
            "• Kết luận: Kết quả phép lai hoàn toàn phù hợp với thực nghiệm của Mendel."
        )
    }
]

# ==============================================================================
# HÀM BỔ TRỢ ĐỊNH DẠNG WORD (.docx) CHUẨN GEMINI.md
# ==============================================================================

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders_horizontal_only(table, border_color_hex="8BC390", border_sz="6"):
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color_hex}"/>\n'
        f'  <w:bottom w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color_hex}"/>\n'
        f'  <w:left w:val="none"/>\n'
        f'  <w:right w:val="none"/>\n'
        f'  <w:insideH w:val="none"/>\n'
        f'  <w:insideV w:val="none"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)

def add_header_footer(doc):
    for s in doc.sections:
        s.top_margin = Inches(0.59)    # 1.5 cm
        s.bottom_margin = Inches(0.59) # 1.5 cm
        s.left_margin = Inches(0.59)   # 1.5 cm
        s.right_margin = Inches(0.59)  # 1.5 cm
        s.header_distance = Inches(0.24) # 0.6 cm
        s.footer_distance = Inches(0.24) # 0.6 cm

        # Header Table
        header = s.header
        header.is_linked_to_previous = False
        tbl_hdr = header.add_table(rows=1, cols=2, width=Inches(7.08))
        tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders_horizontal_only(tbl_hdr, border_color_hex=COLOR_SECONDARY_HEX, border_sz="6")

        c0 = tbl_hdr.cell(0, 0)
        c0.width = Inches(3.54)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r0 = p0.add_run("Tài liệu lưu hành nội bộ")
        r0.font.name = FONT_NAME
        r0.font.size = Pt(10)
        r0.font.italic = True
        r0.font.color.rgb = COLOR_PRIMARY

        c1 = tbl_hdr.cell(0, 1)
        c1.width = Inches(3.54)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r1 = p1.add_run("Thầy ")
        r1.font.name = FONT_NAME
        r1.font.size = Pt(10)
        r1.font.color.rgb = COLOR_PRIMARY
        r1_b = p1.add_run("TRẦN VĂN THẠNH")
        r1_b.font.name = FONT_NAME
        r1_b.font.size = Pt(10)
        r1_b.font.bold = True
        r1_b.font.color.rgb = COLOR_PRIMARY
        r1_ph = p1.add_run(" - 0777.470.803")
        r1_ph.font.name = FONT_NAME
        r1_ph.font.size = Pt(10)
        r1_ph.font.color.rgb = COLOR_PRIMARY

        # Footer Table
        footer = s.footer
        footer.is_linked_to_previous = False
        tbl_ftr = footer.add_table(rows=1, cols=2, width=Inches(7.08))
        tbl_ftr.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders_horizontal_only(tbl_ftr, border_color_hex=COLOR_SECONDARY_HEX, border_sz="6")

        fc0 = tbl_ftr.cell(0, 0)
        fc0.width = Inches(5.3)
        fp0 = fc0.paragraphs[0]
        fp0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fr0 = fp0.add_run("CS1: 6/15 Nguyễn Hoàng, Kim Long   |   CS2: 24 Đặng Thái Thân, TP Huế")
        fr0.font.name = FONT_NAME
        fr0.font.size = Pt(9.5)
        fr0.font.italic = True
        fr0.font.color.rgb = COLOR_PRIMARY

        fc1 = tbl_ftr.cell(0, 1)
        fc1.width = Inches(1.78)
        fp1 = fc1.paragraphs[0]
        fp1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        fr1 = fp1.add_run("Trang ")
        fr1.font.name = FONT_NAME
        fr1.font.size = Pt(10)
        fr1.font.bold = True
        fr1.font.color.rgb = COLOR_PRIMARY

        fldSimple = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
        fp1._p.append(fldSimple)

def add_header_block(doc, is_teacher=False, is_screen=False):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c0 = tbl.cell(0, 0)
    c0.width = Inches(3.2)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.line_spacing = 1.15
    p0.paragraph_format.space_after = Pt(2)

    r_ubnd = p0.add_run("TRƯỜNG THCS NGUYỄN HOÀNG\n")
    r_ubnd.font.name = FONT_NAME
    r_ubnd.font.size = Pt(10.5)
    r_ubnd.font.bold = True
    r_ubnd.font.color.rgb = COLOR_TEXT

    r_nh = p0.add_run("NĂM HỌC 2026 - 2027\n")
    r_nh.font.name = FONT_NAME
    r_nh.font.size = Pt(10)
    r_nh.font.italic = True

    r_de = p0.add_run("ĐỀ CƯƠNG CHÍNH THỨC")
    r_de.font.name = FONT_NAME
    r_de.font.size = Pt(10)
    r_de.font.bold = True
    r_de.font.color.rgb = COLOR_PRIMARY

    c1 = tbl.cell(0, 1)
    c1.width = Inches(3.88)
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.line_spacing = 1.15
    p1.paragraph_format.space_after = Pt(2)

    r_title = p1.add_run("ĐỀ CƯƠNG ÔN TẬP GIỮA KỲ I\n")
    r_title.font.name = FONT_NAME
    r_title.font.size = Pt(11)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_PRIMARY

    r_sub = p1.add_run("MÔN: KHOA HỌC TỰ NHIÊN 9\n")
    r_sub.font.name = FONT_NAME
    r_sub.font.size = Pt(11)
    r_sub.font.bold = True
    r_sub.font.color.rgb = COLOR_PRIMARY

    r_note = p1.add_run("PHÂN MÔN VẬT SỐNG (SINH HỌC)")
    r_note.font.name = FONT_NAME
    r_note.font.size = Pt(10)
    r_note.font.bold = True
    r_note.font.color.rgb = COLOR_SOLUTION_TITLE

    # Teacher / Screen notice
    if is_teacher:
        p_t = doc.add_paragraph()
        p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_t.paragraph_format.space_before = Pt(4)
        p_t.paragraph_format.space_after = Pt(6)
        r_tn = p_t.add_run("(BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT VÀ ĐÁP ÁN ĐẦY ĐỦ)")
        r_tn.font.name = FONT_NAME
        r_tn.font.size = Pt(10.5)
        r_tn.font.bold = True
        r_tn.font.color.rgb = COLOR_PRIMARY
    elif is_screen:
        p_t = doc.add_paragraph()
        p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_t.paragraph_format.space_before = Pt(4)
        p_t.paragraph_format.space_after = Pt(6)
        r_tn = p_t.add_run("(BẢN DẠY TRÊN MÀN HÌNH TIVI 65 INCH -- KHÔNG CÓ ĐÁP ÁN)")
        r_tn.font.name = FONT_NAME
        r_tn.font.size = Pt(10.5)
        r_tn.font.bold = True
        r_tn.font.color.rgb = COLOR_PRIMARY
    else:
        p_stu = doc.add_paragraph()
        p_stu.paragraph_format.space_before = Pt(4)
        p_stu.paragraph_format.space_after = Pt(6)
        r_s = p_stu.add_run("Họ và tên học sinh: ........................................................................ Lớp: .............. Số báo danh: .......................................")
        r_s.font.name = FONT_NAME
        r_s.font.size = Pt(10.5)
        r_s.font.italic = True

def add_part_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT_NAME
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

def add_subpart_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = FONT_NAME
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_SOLUTION_TITLE

def add_solution_paragraph(doc, text):
    lines = text.strip().split("\n")
    for i, line in enumerate(lines):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.28)
        p.paragraph_format.space_before = Pt(2) if i == 0 else Pt(1)
        p.paragraph_format.space_after = Pt(2) if i == len(lines)-1 else Pt(1)
        p.paragraph_format.line_spacing = 1.15
        
        if i == 0:
            r_tag = p.add_run(">> Hướng dẫn giải chi tiết:\n")
            r_tag.font.name = FONT_NAME
            r_tag.font.size = Pt(10.5)
            r_tag.font.bold = True
            r_tag.font.color.rgb = COLOR_SOLUTION_TITLE
        
        parts = line.split("*")
        for idx, part in enumerate(parts):
            if not part:
                continue
            r = p.add_run(part)
            r.font.name = FONT_NAME
            r.font.size = Pt(10.5)
            if idx % 2 == 1:
                r.font.bold = True
                r.font.color.rgb = COLOR_PRIMARY
            else:
                r.font.color.rgb = COLOR_TEXT

def add_dotted_lines(doc, num_lines=4):
    for _ in range(num_lines):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("......................................................................................................................................................................................")
        r.font.name = FONT_NAME
        r.font.size = Pt(9.5)
        r.font.color.rgb = COLOR_GRAY

def add_options_two_by_two(doc, options, correct_key=None, is_teacher=False):
    # Dòng 1: A và B (Tab stop ở 3.5 inch)
    p1 = doc.add_paragraph()
    p1.paragraph_format.left_indent = Inches(0.2)
    p1.paragraph_format.tab_stops.add_tab_stop(Inches(3.5), WD_TAB_ALIGNMENT.LEFT)
    p1.paragraph_format.space_before = Pt(1)
    p1.paragraph_format.space_after = Pt(1)
    p1.paragraph_format.line_spacing = 1.15

    # Phương án A
    r_a = p1.add_run(options[0])
    r_a.font.name = FONT_NAME
    r_a.font.size = Pt(10.5)
    if is_teacher and correct_key == "A":
        r_a.font.bold = True
        r_a.font.color.rgb = COLOR_PRIMARY

    p1.add_run("\t")

    # Phương án B
    r_b = p1.add_run(options[1])
    r_b.font.name = FONT_NAME
    r_b.font.size = Pt(10.5)
    if is_teacher and correct_key == "B":
        r_b.font.bold = True
        r_b.font.color.rgb = COLOR_PRIMARY

    # Dòng 2: C và D
    p2 = doc.add_paragraph()
    p2.paragraph_format.left_indent = Inches(0.2)
    p2.paragraph_format.tab_stops.add_tab_stop(Inches(3.5), WD_TAB_ALIGNMENT.LEFT)
    p2.paragraph_format.space_before = Pt(1)
    p2.paragraph_format.space_after = Pt(2)
    p2.paragraph_format.line_spacing = 1.15

    # Phương án C
    r_c = p2.add_run(options[2])
    r_c.font.name = FONT_NAME
    r_c.font.size = Pt(10.5)
    if is_teacher and correct_key == "C":
        r_c.font.bold = True
        r_c.font.color.rgb = COLOR_PRIMARY

    p2.add_run("\t")

    # Phương án D
    r_d = p2.add_run(options[3])
    r_d.font.name = FONT_NAME
    r_d.font.size = Pt(10.5)
    if is_teacher and correct_key == "D":
        r_d.font.bold = True
        r_d.font.color.rgb = COLOR_PRIMARY

# ==============================================================================
# HÀM TẠO FILE WORD (HỌC SINH HOẶC GIÁO VIÊN HOẶC MÀN HÌNH)
# ==============================================================================
def create_word_document(mode="student"):
    doc = docx.Document()
    add_header_footer(doc)
    is_teacher = (mode == "teacher")
    is_screen = (mode == "screen")
    add_header_block(doc, is_teacher=is_teacher, is_screen=is_screen)

    # --------------------------------------------------------------------------
    # I. PHẦN TRẮC NGHIỆM
    # --------------------------------------------------------------------------
    add_part_heading(doc, "I. PHẦN TRẮC NGHIỆM")
    add_subpart_heading(doc, "1. TRẮC NGHIỆM NHIỀU LỰA CHỌN (1,0 điểm)")
    
    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(1)
    p_note.paragraph_format.space_after = Pt(4)
    r_n = p_note.add_run("(Chọn 1 phương án trả lời đúng nhất. Mỗi câu trả lời đúng được 0,25 điểm)")
    r_n.font.name = FONT_NAME
    r_n.font.size = Pt(10)
    r_n.font.italic = True

    for item in MCQ_DATA:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(1)
        p_q.paragraph_format.line_spacing = 1.15
        
        r_num = p_q.add_run(f"Câu {item['id']}. ")
        r_num.font.name = FONT_NAME
        r_num.font.size = Pt(10.5)
        r_num.font.bold = True
        
        r_cnt = p_q.add_run(item["question"])
        r_cnt.font.name = FONT_NAME
        r_cnt.font.size = Pt(10.5)

        add_options_two_by_two(doc, item["options"], correct_key=item["answer"], is_teacher=is_teacher)

        if is_teacher:
            add_solution_paragraph(doc, item["solution"])

    # --------------------------------------------------------------------------
    # 2. PHẦN TRẮC NGHIỆM TRẢ LỜI NGẮN
    # --------------------------------------------------------------------------
    add_subpart_heading(doc, "2. PHẦN TRẮC NGHIỆM TRẢ LỜI NGẮN (0,5 điểm)")
    
    p_sn_note = doc.add_paragraph()
    p_sn_note.paragraph_format.space_before = Pt(1)
    p_sn_note.paragraph_format.space_after = Pt(4)
    r_snn = p_sn_note.add_run("Học sinh trả lời bằng số. Mỗi ý a), b) trả lời bằng một con số, tối đa 4 kí tự kể cả dấu phẩy. Mỗi ý đúng được 0,25 điểm.")
    r_snn.font.name = FONT_NAME
    r_snn.font.size = Pt(10)
    r_snn.font.italic = True

    for item in SHORT_ANS_DATA:
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(4)
        p_sub.paragraph_format.space_after = Pt(2)
        p_sub.paragraph_format.line_spacing = 1.15
        
        r_sid = p_sub.add_run(f"{item['sub_id']}) ")
        r_sid.font.name = FONT_NAME
        r_sid.font.size = Pt(10.5)
        r_sid.font.bold = True
        
        r_intro = p_sub.add_run(item["intro"])
        r_intro.font.name = FONT_NAME
        r_intro.font.size = Pt(10.5)

        # Chèn ảnh nếu có (ý b)
        if item.get("has_image") and os.path.exists(DIAGRAM_PATH):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(4)
            p_img.paragraph_format.space_after = Pt(2)
            run_img = p_img.add_run()
            run_img.add_picture(DIAGRAM_PATH, width=Inches(6.2))

        # Liệt kê các nhận định
        for st in item["statements"]:
            p_st = doc.add_paragraph()
            p_st.paragraph_format.left_indent = Inches(0.2)
            p_st.paragraph_format.space_before = Pt(1)
            p_st.paragraph_format.space_after = Pt(1)
            p_st.paragraph_format.line_spacing = 1.15
            r_st = p_st.add_run(st)
            r_st.font.name = FONT_NAME
            r_st.font.size = Pt(10.5)

        # Dòng câu hỏi
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(2)
        p_ans.paragraph_format.space_after = Pt(2)
        r_ans = p_ans.add_run(item["question"])
        r_ans.font.name = FONT_NAME
        r_ans.font.size = Pt(10.5)
        r_ans.font.bold = True

        if is_teacher:
            add_solution_paragraph(doc, item["solution"])
        elif not is_screen:
            p_ans_stu = doc.add_paragraph()
            p_ans_stu.paragraph_format.left_indent = Inches(0.2)
            p_ans_stu.paragraph_format.space_before = Pt(1)
            p_ans_stu.paragraph_format.space_after = Pt(2)
            r_as = p_ans_stu.add_run("Trả lời: ............................................................................................................................................")
            r_as.font.name = FONT_NAME
            r_as.font.size = Pt(9.5)
            r_as.font.color.rgb = COLOR_GRAY

    # --------------------------------------------------------------------------
    # II. PHẦN TỰ LUẬN
    # --------------------------------------------------------------------------
    add_part_heading(doc, "II. PHẦN TỰ LUẬN")
    for item in ESSAY_DATA:
        p_eq = doc.add_paragraph()
        p_eq.paragraph_format.space_before = Pt(4)
        p_eq.paragraph_format.space_after = Pt(2)
        p_eq.paragraph_format.line_spacing = 1.15
        
        r_num = p_eq.add_run(f"{item['title']}. ")
        r_num.font.name = FONT_NAME
        r_num.font.size = Pt(10.5)
        r_num.font.bold = True
        
        r_cnt = p_eq.add_run(item["content"])
        r_cnt.font.name = FONT_NAME
        r_cnt.font.size = Pt(10.5)

        if is_teacher:
            add_solution_paragraph(doc, item["solution"])
        elif not is_screen:
            add_dotted_lines(doc, num_lines=8)

    # Dòng kết thúc đề
    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_end.paragraph_format.space_before = Pt(10)
    r_end = p_end.add_run("“HẾT -- Chúc các em làm bài đạt kết quả tốt nhất!”")
    r_end.font.name = FONT_NAME
    r_end.font.size = Pt(10)
    r_end.font.italic = True
    r_end.font.color.rgb = COLOR_PRIMARY

    # Đặt tên file xuất
    if mode == "student":
        filename = "De_Cuong_GK1_KHTN9_NguyenHoang_SinhHoc_HocSinh.docx"
    elif mode == "teacher":
        filename = "De_Cuong_GK1_KHTN9_NguyenHoang_SinhHoc_GiaoVien.docx"
    else:
        filename = "De_Cuong_GK1_KHTN9_NguyenHoang_SinhHoc_DayTrenManHinh.docx"

    out_path = os.path.join(SAN_PHAM_DIR, filename)
    doc.save(out_path)
    print(f"[OK] Đã tạo thành công file Word: {out_path} ({os.path.getsize(out_path):,} bytes)")
    return out_path

# ==============================================================================
# HÀM TẠO FILE SLIDE PDF DẠY TRÊN MÀN HÌNH 16:9 CHO TIVI 65 INCH (KHÔNG CÓ ĐÁP ÁN)
# ==============================================================================
def create_screen_teaching_pdf():
    tex_path = os.path.join(SCRATCH_DIR, "De_Cuong_GK1_KHTN9_NguyenHoang_SinhHoc_DayTrenManHinh.tex")
    
    tex_lines = [r"""\documentclass[12pt,landscape]{article}
\usepackage[paperwidth=29.7cm,paperheight=16.706cm,top=1.0cm,bottom=1.0cm,left=1.2cm,right=1.2cm,headheight=0pt,headsep=0pt,footskip=0pt]{geometry}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{tabularx}
\usepackage{enumitem}

\definecolor{brandRed}{RGB}{211, 47, 47}
\definecolor{brandBlue}{RGB}{21, 101, 192}
\definecolor{brandGreen}{RGB}{46, 125, 50}
\definecolor{offBlack}{RGB}{26, 26, 26}
\definecolor{grayText}{RGB}{117, 117, 117}

\pagestyle{empty}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}

\begin{document}
"""]

    # --------------------------------------------------------------------------
    # SLIDE 1: TRANG TIÊU ĐỀ
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 1: TIÊU ĐỀ
% ==============================================================================
\begin{center}
\vspace*{0.8cm}
{\fontsize{24}{30}\selectfont \textbf{\color{offBlack}TRƯỜNG THCS NGUYỄN HOÀNG -- NĂM HỌC 2026 -- 2027}}\\[12pt]
{\fontsize{34}{44}\selectfont \textbf{\color{brandRed}ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA HỌC KỲ I}}\\[10pt]
{\fontsize{28}{36}\selectfont \textbf{\color{brandBlue}MÔN: KHOA HỌC TỰ NHIÊN 9 -- PHÂN MÔN VẬT SỐNG}}\\[8pt]
{\fontsize{22}{28}\selectfont \textit{\color{grayText}(Chủ đề: Di truyền và biến dị -- Nucleic acid và Tái bản DNA)}}\\[1.2cm]

{\fontsize{20}{26}\selectfont \textbf{Giáo viên hướng dẫn:} Thầy \textbf{\color{brandRed}TRẦN VĂN THẠNH} -- 0777.470.803}\\[6pt]
{\fontsize{18}{24}\selectfont \textit{Cơ sở 1: 6/15 Nguyễn Hoàng, Kim Long \quad|\quad Cơ sở 2: 24 Đặng Thái Thân, TP Huế}}\\[1.0cm]

{\fontsize{18}{24}\selectfont \textbf{\color{brandGreen}SLIDE TRÌNH CHIẾU TƯƠNG TÁC TIVI 65 INCH (KHÔNG HIỂN THỊ ĐÁP ÁN)}}
\end{center}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 2: CÂU 1 & CÂU 2 (TRẮC NGHIỆM NHIỀU LỰA CHỌN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 2: CÂU 1 & CÂU 2
% ==============================================================================
\fontsize{26}{36}\selectfont
\textbf{Câu 1.} Hiện tượng di truyền là gì?
\vspace{6pt}

\begin{tabularx}{\textwidth}{@{}XX@{}}
\textbf{A.} Sự khác nhau giữa các cá thể cùng loài. & \textbf{B.} Sự truyền đạt các tính trạng của bố mẹ, tổ tiên cho con cháu. \\[8pt]
\textbf{C.} Sự giống nhau của 2 anh em cùng 1 nhà. & \textbf{D.} Sự xuất hiện 1 số tính trạng mới của con so với cha mẹ.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 2.} Theo ý tưởng của Mendel, yếu tố nào quy định sự di truyền của tính trạng?
\vspace{6pt}

\begin{tabularx}{\textwidth}{@{}XX@{}}
\textbf{A.} Các chất dinh dưỡng trong đất. & \textbf{B.} Ánh sáng mà cây nhận được. \\[8pt]
\textbf{C.} Các nhân tố di truyền. & \textbf{D.} Lượng nước cây hấp thụ.
\end{tabularx}
\vspace*{0.5cm}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 3: CÂU 3 & CÂU 4 (TRẮC NGHIỆM NHIỀU LỰA CHỌN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 3: CÂU 3 & CÂU 4
% ==============================================================================
\fontsize{26}{36}\selectfont
\textbf{Câu 3.} Hai loại nucleic acid là
\vspace{6pt}

\begin{tabularx}{\textwidth}{@{}XX@{}}
\textbf{A.} DNA và RNA. & \textbf{B.} DNA và protein. \\[8pt]
\textbf{C.} RNA và lipid. & \textbf{D.} Protein và lipid.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 4.} Chức năng của DNA là
\vspace{6pt}

\begin{tabularx}{\textwidth}{@{}XX@{}}
\textbf{A.} Cung cấp trực tiếp năng lượng cho mọi hoạt động của tế bào. & \textbf{B.} Vận chuyển oxygen trong máu. \\[8pt]
\textbf{C.} Phân giải thức ăn thành chất đơn giản. & \textbf{D.} Lưu giữ, bảo quản và truyền đạt thông tin di truyền.
\end{tabularx}
\vspace*{0.5cm}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 4: CÂU 5 & CÂU 6 (TRẮC NGHIỆM NHIỀU LỰA CHỌN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 4: CÂU 5 & CÂU 6
% ==============================================================================
\fontsize{25}{35}\selectfont
\textbf{Câu 5.} Hiện tượng biến dị là gì?
\vspace{6pt}

\begin{tabularx}{\textwidth}{@{}XX@{}}
\textbf{A.} Sự truyền đạt nguyên vẹn mọi tính trạng cho đời con. & \textbf{B.} Sự sai khác toàn bộ các đặc điểm của con cái so với cha mẹ. \\[8pt]
\textbf{C.} Hiện tượng con sinh ra có những đặc điểm khác nhau và khác với bố mẹ. & \textbf{D.} Sự khác nhau giữa 2--4 thành viên trong 1 gia đình (cha, mẹ và 2 con).
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 6.} Ý tưởng nào của Mendel đặt cơ sở cho các nghiên cứu về gene?
\vspace{6pt}

\begin{tabularx}{\textwidth}{@{}XX@{}}
\textbf{A.} Tính trạng được di truyền nhờ các nhân tố di truyền. & \textbf{B.} Tính trạng chỉ phụ thuộc vào lượng thức ăn. \\[8pt]
\textbf{C.} Tính trạng của bố mẹ luôn hòa trộn vĩnh viễn ở con. & \textbf{D.} Tính trạng ở con không liên quan đến bố mẹ.
\end{tabularx}
\vspace*{0.5cm}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 5: CÂU 7 & CÂU 8 (TRẮC NGHIỆM NHIỀU LỰA CHỌN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 5: CÂU 7 & CÂU 8
% ==============================================================================
\fontsize{26}{36}\selectfont
\textbf{Câu 7.} Nhóm nào gồm đầy đủ hai loại nucleic acid?
\vspace{6pt}

\begin{tabularx}{\textwidth}{@{}XX@{}}
\textbf{A.} RNA và carbohydrate. & \textbf{B.} Protein và DNA. \\[8pt]
\textbf{C.} Lipid và carbohydrate. & \textbf{D.} RNA và DNA.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 8.} Vai trò nào sau đây thuộc về DNA?
\vspace{6pt}

\begin{tabularx}{\textwidth}{@{}XX@{}}
\textbf{A.} Xúc tác mọi phản ứng hóa học trong tế bào. & \textbf{B.} Lưu giữ, bảo quản và truyền đạt thông tin di truyền. \\[8pt]
\textbf{C.} Tạo ra năng lượng trực tiếp cho sự co cơ. & \textbf{D.} Vận chuyển các chất dinh dưỡng trong máu.
\end{tabularx}
\vspace*{0.5cm}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 6: TRẢ LỜI NGẮN - CÂU 1a (ĐA DẠNG NUCLEOTIDE CỦA DNA)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 6: TRẢ LỜI NGẮN - CÂU 1a
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}2. TRẮC NGHIỆM TRẢ LỜI NGẮN -- CÂU 1a}}
\vspace{8pt}

\fontsize{21}{28}\selectfont
\textbf{a) Khi giải thích vì sao chỉ từ bốn loại nucleotide có thể tạo nên nhiều phân tử DNA khác nhau, có các nhận định sau:}
\vspace{6pt}

\begin{enumerate}[label=(\arabic*),itemsep=4pt]
  \item Hai DNA cùng số nucleotide nhưng khác trình tự sắp xếp có thể khác nhau.
  \item Hai DNA có cùng số lượng từng loại nucleotide thì chắc chắn có trình tự giống nhau.
  \item Sự khác nhau về số lượng, thành phần và trình tự nucleotide góp phần tạo nên tính đa dạng của DNA.
  \item Vì chỉ có bốn loại nucleotide nên chỉ có thể tạo ra bốn loại phân tử DNA.
  \item Chỉ cần biết tỉ lệ các loại nucleotide là xác định được duy nhất trình tự của một DNA.
\end{enumerate}

\vspace{10pt}
\textbf{Có bao nhiêu nhận định đúng? \quad Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots$
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 7: TRẢ LỜI NGẮN - CÂU 1b (SƠ ĐỒ TÁI BẢN DNA)
    # --------------------------------------------------------------------------
    diagram_unix = DIAGRAM_PATH.replace("\\", "/")
    slide7_tex = r"""
% ==============================================================================
% SLIDE 7: TRẢ LỜI NGẮN - CÂU 1b
% ==============================================================================
{\fontsize{23}{28}\selectfont \textbf{\color{brandBlue}2. TRẮC NGHIỆM TRẢ LỜI NGẮN -- CÂU 1b (TÁI BẢN DNA)}}
\vspace{4pt}

\fontsize{19}{24}\selectfont
\textbf{b) Quan sát sơ đồ một lần tái bản DNA dưới đây và xét các nhận định sau:}
\vspace{4pt}

\begin{center}
\includegraphics[width=0.88\textwidth]{__DIAGRAM_PATH__}
\end{center}
\vspace{2pt}

\begin{tabularx}{\textwidth}{@{}XX@{}}
(1) Hai mạch DNA mẹ tách nhau để làm khuôn. & (5) Mỗi DNA con có 1 mạch từ mẹ và 1 mạch mới. \\
(2) Nu tự do ghép theo nguyên tắc bổ sung. & (6) Mạch mới tổng hợp trước khi mẹ tách nhau. \\
(3) Hai DNA con chỉ gồm các mạch mới. & (7) Nu trên mạch mới lắp ghép theo NTBS. \\
(4) Thông tin mẹ được truyền cho 2 con. & (8) Tái bản chính xác duy trì thông tin di truyền.
\end{tabularx}

\vspace{6pt}
\textbf{Có bao nhiêu nhận định đúng? \quad Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots$
\newpage
""".replace("__DIAGRAM_PATH__", diagram_unix)
    tex_lines.append(slide7_tex)

    # --------------------------------------------------------------------------
    # SLIDE 8: PHẦN TỰ LUẬN - CÂU 20 (THÍ NGHIỆM MENDEL LAI ĐẬU HÀ LAN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 8: PHẦN TỰ LUẬN - CÂU 20
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}II. PHẦN TỰ LUẬN -- CÂU 20}}
\vspace{10pt}

\fontsize{22}{30}\selectfont
\textbf{Đề bài:} Trong thí nghiệm của Mendel, lai đậu Hà Lan hạt vàng thuần chủng với đậu hạt xanh thuần chủng thu được $\text{F}_1$ toàn hạt vàng; cho $\text{F}_1$ tự thụ phấn, $\text{F}_2$ có tỉ lệ xấp xỉ $3\text{ hạt vàng} : 1\text{ hạt xanh}$.\\
\textbf{Yêu cầu:} Biện luận và viết sơ đồ lai từ $\text{P}$ đến $\text{F}_2$.

\vspace*{2.5cm}
\begin{center}
{\fontsize{19}{25}\selectfont \textit{\color{grayText}(Khoảng trống chữa bài: Xác định trội/lặn $\to$ Quy ước gene $\to$ Sơ đồ lai $\text{P} \to \text{F}_1 \to \text{F}_2$)}}
\end{center}
\newpage
""")

    tex_lines.append(r"\end{document}")

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write("\n".join(tex_lines))
    print(f"[OK] Đã ghi file LaTeX slide: {tex_path}")

    cmd = ["pdflatex", "-interaction=nonstopmode", "-output-directory", SCRATCH_DIR, tex_path]
    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    pdf_temp = os.path.join(SCRATCH_DIR, "De_Cuong_GK1_KHTN9_NguyenHoang_SinhHoc_DayTrenManHinh.pdf")
    pdf_out = os.path.join(SAN_PHAM_DIR, "De_Cuong_GK1_KHTN9_NguyenHoang_SinhHoc_DayTrenManHinh.pdf")

    if os.path.exists(pdf_temp):
        shutil.copy2(pdf_temp, pdf_out)
        print(f"[OK] Đã biên dịch và copy PDF thành công: {pdf_out} ({os.path.getsize(pdf_out):,} bytes)")
    else:
        print(f"[ERR] Không tạo được PDF. Mã lỗi: {res.returncode}")
    return pdf_out

# ==============================================================================
# HÀM ĐỒNG BỘ ONEDRIVE
# ==============================================================================
def sync_to_onedrive():
    files_to_sync = [
        "De_Cuong_GK1_KHTN9_NguyenHoang_SinhHoc_HocSinh.docx",
        "De_Cuong_GK1_KHTN9_NguyenHoang_SinhHoc_GiaoVien.docx",
        "De_Cuong_GK1_KHTN9_NguyenHoang_SinhHoc_DayTrenManHinh.docx",
        "De_Cuong_GK1_KHTN9_NguyenHoang_SinhHoc_DayTrenManHinh.pdf",
    ]
    print("\n[+] Đồng bộ các file sang các thư mục OneDrive:")
    for dest_dir in [ONEDRIVE_DIR1, ONEDRIVE_DIR2]:
        if os.path.exists(dest_dir):
            for fname in files_to_sync:
                src_f = os.path.join(SAN_PHAM_DIR, fname)
                dst_f = os.path.join(dest_dir, fname)
                if os.path.exists(src_f):
                    try:
                        shutil.copy2(src_f, dst_f)
                        print(f"    -> Đã copy {fname} vào: {dest_dir}")
                    except Exception as e:
                        print(f"    [!] Cảnh báo file đang mở/khóa tại {dest_dir}: {e}")
        else:
            print(f"    [-] Không tìm thấy thư mục: {dest_dir}")

# ==============================================================================
# MAIN
# ==============================================================================
if __name__ == "__main__":
    print("=" * 75)
    print("BẮT ĐẦU TẠO 3 FILE CHO ĐỀ CƯƠNG ÔN TẬP GIỮA KỲ I KHTN 9 (THCS NGUYỄN HOÀNG)")
    print("PHÂN MÔN VẬT SỐNG (SINH HỌC)")
    print("1. FILE HỌC SINH (.docx)")
    print("2. FILE GIÁO VIÊN (.docx có lời giải chi tiết)")
    print("3. FILE DẠY TRÊN MÀN HÌNH (Slide PDF 16:9 widescreen cho Tivi 65 inch)")
    print("=" * 75)

    create_word_document(mode="student")
    create_word_document(mode="teacher")
    create_word_document(mode="screen")
    create_screen_teaching_pdf()
    sync_to_onedrive()

    print("\n" + "=" * 75)
    print("HOÀN THÀNH TẤT CẢ 3 FILE THEO ĐÚNG CHUẨN FORM ĐỀ KIỂM TRA!")
    print("=" * 75)
