# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Học Kỳ I - Hóa học 12 - Trường THPT Hương Trà (TP Huế)
Mã đề: 001 - Năm học 2025 - 2026
Chuẩn theo Quy chuẩn Form Đề Kiểm Tra của Thầy TRẦN VĂN THẠNH
Tạo 2 phiên bản: Bản Học Sinh & Bản Giáo Viên (có BẢNG ĐÁP ÁN RÚT GỌN BAO GỒM CẢ TỰ LUẬN & lời giải chi tiết)
"""

import os
import sys
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Import module form_de_kiem_tra
sys.path.append(os.path.join(os.path.dirname(__file__)))
from form_de_kiem_tra import (
    setup_header_footer,
    build_title_block,
    add_section_header,
    render_question_prompt,
    render_options_no_table,
    render_statement,
    add_solution_no_table,
    add_student_writing_lines,
    render_rich_text,
    set_run_style,
    add_p_border_bottom,
    COLOR_PRIMARY,
    COLOR_CORRECT,
    COLOR_SOLUTION,
    COLOR_GRAY
)

DIR_SAN_PHAM = os.path.abspath("San_Pham")
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_de001")
os.makedirs(DIR_SAN_PHAM, exist_ok=True)
os.makedirs(DIR_IMAGES, exist_ok=True)

def set_cell_border(cell, **kwargs):
    """Cài đặt viền cho ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            val, sz, space, color = edge_data
            tag = f'<w:{edge} {nsdecls("w")} w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'
            tcBorders.append(parse_xml(tag))
        else:
            tag = f'<w:{edge} {nsdecls("w")} w:val="none"/>'
            tcBorders.append(parse_xml(tag))
    tcPr.append(tcBorders)

def set_cell_shading(cell, color_hex):
    """Tô màu nền cho ô bảng."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def add_compact_summary_answers(doc):
    """
    Thêm BẢNG ĐÁP ÁN RÚT GỌN đầy đủ: Phần I, Phần II, Phần III và Phần IV Tự luận cho Bản Giáo Viên.
    """
    p_box = doc.add_paragraph()
    p_box.paragraph_format.space_before = Pt(8)
    p_box.paragraph_format.space_after = Pt(2)
    p_box.paragraph_format.keep_with_next = True
    r_box = p_box.add_run("BẢNG ĐÁP ÁN RÚT GỌN (DÀNH CHO GIÁO VIÊN - BAO GỒM CẢ TỰ LUẬN)")
    set_run_style(r_box, font_size=11, bold=True, color=COLOR_CORRECT)

    # 1. BẢNG PHẦN I
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(2)
    p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p1, "*1. PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (4,0 ĐIỂM - MỖI CÂU 0,25 Đ)*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

    tbl1 = doc.add_table(rows=2, cols=17)
    tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl1.autofit = False

    headers_p1 = ["Câu", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"]
    answers_p1 = ["Đ/A", "A", "B", "C", "C", "B", "D", "C", "A", "D", "D", "C", "D", "A", "B", "B", "C"]

    w1_c0 = Cm(1.2)
    w1_c = Cm(1.05)

    for r_idx, r_data in enumerate([headers_p1, answers_p1]):
        row = tbl1.rows[r_idx]
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = w1_c0 if c_idx == 0 else w1_c
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=9.0, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                is_label = (c_idx == 0)
                set_cell_shading(cell, "F5F5F5" if is_label else "FFFFFF")
                c_color = COLOR_PRIMARY if is_label else COLOR_CORRECT
                render_rich_text(p, f"*{val}*", base_size=9.5, base_bold=True, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # 2. BẢNG PHẦN II
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(5)
    p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p2, "*2. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM - 1 ý: 0,1đ; 2 ý: 0,25đ; 3 ý: 0,5đ; 4 ý: 1,0đ)*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

    tbl2 = doc.add_table(rows=3, cols=6)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl2.autofit = False

    tbl2_data = [
        ["Câu hỏi", "Nội dung câu hỏi", "Lệnh a", "Lệnh b", "Lệnh c", "Lệnh d"],
        ["Câu 1", "Polymer gia cường sợi Kevlar", "ĐÚNG", "ĐÚNG", "SAI", "ĐÚNG"],
        ["Câu 2", "Pin điện hóa Galvani Zn – Cu", "SAI", "ĐÚNG", "ĐÚNG", "SAI"]
    ]

    w2_widths = [Cm(1.8), Cm(6.2), Cm(2.5), Cm(2.5), Cm(2.5), Cm(2.5)]

    for r_idx, row_vals in enumerate(tbl2_data):
        row = tbl2.rows[r_idx]
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = w2_widths[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx != 1 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=9.0, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                is_label = (c_idx < 2)
                set_cell_shading(cell, "F5F5F5" if is_label else "FFFFFF")
                if is_label:
                    c_color = COLOR_PRIMARY if c_idx == 0 else RGBColor(0, 0, 0)
                else:
                    c_color = COLOR_CORRECT if val == "ĐÚNG" else COLOR_PRIMARY
                render_rich_text(p, f"*{val}*", base_size=9.5, base_bold=True, base_color=c_color)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # 3. BẢNG PHẦN III & PHẦN IV
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(5)
    p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p3, "*3. PHẦN III: TRẢ LỜI NGẮN (0,5 Đ/CÂU) & PHẦN IV: TỰ LUẬN (1,0 Đ/CÂU)*", base_size=10, base_bold=True, base_color=COLOR_PRIMARY)

    tbl3 = doc.add_table(rows=2, cols=7)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl3.autofit = False

    headers_p34 = ["P.III - C1", "P.III - C2", "P.III - C3", "P.III - C4", "P.IV - C1", "P.IV - C2", "P.IV - C3"]
    answers_p34 = ["2", "4", "100", "6,4", "nCH~2~=CHCl -> [-CH~2~-CHCl-]~n~", "+0,80 V", "2,95 giờ (10605 s)"]

    w3_c = Cm(2.57)

    for r_idx, r_data in enumerate([headers_p34, answers_p34]):
        row = tbl3.rows[r_idx]
        for c_idx, val in enumerate(r_data):
            cell = row.cells[c_idx]
            cell.width = w3_c
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)

            if r_idx == 0:
                set_cell_shading(cell, "E8F5E9")
                render_rich_text(p, f"*{val}*", base_size=9.0, base_bold=True, base_color=COLOR_PRIMARY)
            else:
                set_cell_shading(cell, "FFFFFF")
                render_rich_text(p, f"*{val}*", base_size=9.0, base_bold=True, base_color=COLOR_CORRECT)

            set_cell_border(cell,
                            top=("single", "4", "0", "8BC390"),
                            bottom=("single", "4", "0", "8BC390"),
                            left=("single", "4", "0", "8BC390"),
                            right=("single", "4", "0", "8BC390"))

    # Đường kẻ phân cách Sage Green
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(8)
    add_p_border_bottom(p_div, color="8BC390", sz="6")


DATA_EXAM_001 = {
    "school_line": "SỞ GIÁO DỤC VÀ ĐÀO TẠO\nTRƯỜNG THPT HƯƠNG TRÀ\nTỔ HÓA - SINH - CNNN",
    "exam_title": "ĐỀ KIỂM TRA HỌC KỲ I NĂM HỌC 2025 - 2026",
    "subject_line": "MÔN: HÓA HỌC – KHỐI 12",
    "num_pages": "03",
    "duration": "45 phút (không kể thời gian phát đề)",
    "exam_code": "001",
    "atomic_masses": "C = 12; H = 1; O = 16; N = 14; Na = 23; Mg = 24; S = 32; Cl = 35,5; Cu = 64; Ag = 108. F = 96500 C/mol; π = 3,14",

    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1: Polymer nào sau đây được dùng để chế tạo vật liệu có tính dẻo?",
            "options": [
                "*A. Poly(vinyl chloride).* ",
                "B. Poly(vinyl cyanide).",
                "C. Poli(hexamethylene adipamide).",
                "D. Poly(ethylene terephthalate)."
            ],
            "solution": "Poly(vinyl chloride) (PVC) là loại chất dẻo phổ biến trong đời sống. Poly(vinyl cyanide) (tơ olon/nitron), poli(hexamethylene adipamide) (tơ nylon-6,6) và poly(ethylene terephthalate) (tơ lapsan) là các loại tơ sợi tổng hợp."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Trong số các ion: Ag^+, Al^3+^, Fe^2+^, Cu^2+^, ion nào có tính oxi hóa mạnh nhất ở điều kiện chuẩn?",
            "options": [
                "A. Fe^2+^.",
                "*B. Ag^+^.*",
                "C. Al^3+^.",
                "D. Cu^2+^."
            ],
            "solution": "Căn cứ vào dãy thế điện cực chuẩn: E°(Al^3+^/Al) = -1,66 V < E°(Fe^2+^/Fe) = -0,44 V < E°(Cu^2+^/Cu) = +0,34 V < E°(Ag^+^/Ag) = +0,80 V. Cặp có E° càng lớn thì dạng oxi hóa càng mạnh, do đó Ag^+^ có tính oxi hóa mạnh nhất."
        },
        {
            "id": 3,
            "prompt": "Câu 3: Chất nào sau đây là thành phần chủ yếu của xà phòng?",
            "options": [
                "A. CH~2~CH~2~OH.",
                "B. CH~3~COONa.",
                "*C. CH~3~[CH~2~]~14~COONa.*",
                "D. CH~3~[CH~2~]~16~COOH."
            ],
            "solution": "Xà phòng là hỗn hợp muối sodium hoặc potassium của các acid béo cao (palmitic acid, stearic acid,...). Hợp chất CH~3~[CH~2~]~14~COONa chính là sodium palmitate."
        },
        {
            "id": 4,
            "prompt": "Câu 4: Tác hại lớn nhất của việc lạm dụng chất dẻo đối với môi trường là gì?",
            "options": [
                "A. Gây mùi khó chịu.",
                "B. Làm giảm độ màu mỡ của đất.",
                "*C. Khó phân hủy, gây ô nhiễm môi trường lâu dài.*",
                "D. Gây ra hiệu ứng nhà kính."
            ],
            "solution": "Chất dẻo có tính bền vững sinh học rất cao, rất khó bị phân hủy trong điều kiện tự nhiên (mất hàng trăm đến hàng nghìn năm), dẫn tới sự tích tụ rác thải nhựa gây ô nhiễm môi trường đất, nước và hình thành các hạt vi nhựa gây nguy hại cho hệ sinh thái."
        },
        {
            "id": 5,
            "prompt": "Câu 5: Dãy gồm các chất được dùng để tổng hợp cao su Buna-S là:",
            "options": [
                "A. CH~2~=C(CH~3~)-CH=CH~2~, C~6~H~5~CH=CH~2~.",
                "*B. CH~2~=CH-CH=CH~2~, C~6~H~5~CH=CH~2~.*",
                "C. CH~2~=CH-CH=CH~2~, lưu huỳnh.",
                "D. CH~2~=CH-CH=CH~2~, CH~3~-CH=CH~2~."
            ],
            "solution": "Cao su Buna-S được sản xuất bằng phản ứng đồng trùng hợp giữa buta-1,3-diene (CH~2~=CH-CH=CH~2~) và styrene (C~6~H~5~CH=CH~2~) có mặt xúc tác natri."
        },
        {
            "id": 6,
            "prompt": "Câu 6: Thuốc thử nào sau đây được dùng để nhận biết tinh bột?",
            "options": [
                "A. Thuốc thử Tollens.",
                "B. Dung dịch Br~2~.",
                "C. Cu(OH)~2~.",
                "*D. Dung dịch I~2~.*"
            ],
            "solution": "Dung dịch iodine (I~2~) tác dụng với các phân tử amylose hình xoắn của tinh bột tạo thành hợp chất bọc có màu xanh tím đặc trưng. Khi đun nóng màu xanh tím biến mất, để nguội lại xuất hiện."
        },
        {
            "id": 7,
            "prompt": "Câu 7: Điện phân CaCl~2~ nóng chảy, ở cathode xảy ra quá trình nào?",
            "options": [
                "A. Oxi hóa ion Ca^2+^.",
                "B. Khử ion Cl^-^.",
                "*C. Khử ion Ca^2+^.*",
                "D. Oxi hóa ion Cl^-^."
            ],
            "solution": "Ở cathode (cực âm) xảy ra quá trình khử cation kim loại: Ca^2+^ + 2e -> Ca. Ở anode (cực dương) xảy ra quá trình oxi hóa anion: 2Cl^-^ -> Cl~2~ + 2e."
        },
        {
            "id": 8,
            "prompt": "Câu 8: Amino acid không có loại phản ứng nào sau đây?",
            "options": [
                "*A. Phản ứng thủy phân.*",
                "B. Phản ứng ester hoá.",
                "C. Phản ứng với acid và base.",
                "D. Phản ứng trùng ngưng."
            ],
            "solution": "Amino acid là hợp chất hữu cơ tạp chức đơn phân chứa đồng thời nhóm amino (-NH~2~) và nhóm carboxyl (-COOH). Chúng có tính lưỡng tính (tác dụng với acid, base), có phản ứng ester hóa của nhóm -COOH, có phản ứng trùng ngưng tạo polyamide; nhưng không có liên kết peptide hay ester nên không có phản ứng thủy phân."
        },
        {
            "id": 9,
            "prompt": "Câu 9: Khẳng định nào sau đây là sai?",
            "options": [
                "A. Polymer không bay hơi được.",
                "B. Polymer không có nhiệt độ nóng chảy nhất định.",
                "C. Đa số polymer không tan trong các dung môi thông thường.",
                "*D. Thủy tinh hữu cơ là vật liệu trong suốt, giòn và kém bền.*"
            ],
            "solution": "Thủy tinh hữu cơ plexiglas hay poly(methyl methacrylate) là vật liệu có độ bền cơ học cao, chịu lực tốt, khó bị vỡ vụn và bền bỉ hơn nhiều so với thủy tinh silicat vô cơ thông thường (khẳng định giòn và kém bền là sai)."
        },
        {
            "id": 10,
            "prompt": "Câu 10: Chất nào sau đây thuộc loại monosaccharide?",
            "options": [
                "A. Maltose.",
                "B. Cellulose.",
                "C. Saccharose.",
                "*D. Fructose.*"
            ],
            "solution": "Fructose và glucose là các monosaccharide (đường đơn). Saccharose và maltose là các disaccharide. Tinh bột và cellulose là các polysaccharide."
        },
        {
            "id": 11,
            "prompt": "Câu 11: Tinh bột và cellulose là các polymer tự nhiên tạo bởi các mắt xích tương ứng là",
            "options": [
                "A. α-glucose và β-fructose.",
                "B. β-fructose và β-glucose.",
                "*C. α-glucose và β-glucose.*",
                "D. α-fructose và β-glucose."
            ],
            "solution": "Tinh bột được cấu tạo từ các mắt xích α-glucose liên kết qua liên kết α-1,4-glycosidic (amylose) và α-1,6-glycosidic (amylopectin). Cellulose được tạo thành từ các mắt xích β-glucose liên kết qua liên kết β-1,4-glycosidic."
        },
        {
            "id": 12,
            "prompt": "Câu 12: Phát biểu nào sau đây là không đúng?",
            "options": [
                "A. Saccharose có nhiều trong thân cây mía.",
                "B. Glucose có nhiều trong các quả chín.",
                "C. Fructose có nhiều trong mật ong.",
                "*D. Maltose có nhiều trong đường kính.*"
            ],
            "solution": "Đường kính, đường cát, đường phèn sử dụng hàng ngày có thành phần chính là saccharose (chiếm tới 99,8%), không phải maltose (đường mạch nha)."
        },
        {
            "id": 13,
            "prompt": "Câu 13: Kí hiệu cặp oxi hóa – khử ứng với quá trình khử: Fe^3+^ + 1e ⇌ Fe^2+^ là",
            "options": [
                "*A. Fe^3+^/Fe^2+^.*",
                "B. Fe^2+^/Fe^3+^.",
                "C. Fe^3+^/Fe.",
                "D. Fe^2+^/Fe."
            ],
            "solution": "Theo quy ước IUPAC, kí hiệu cặp oxi hóa - khử được viết dưới dạng: Dạng oxi hóa / Dạng khử. Trong bán phản ứng trên, Fe^3+^ là dạng oxi hóa và Fe^2+^ là dạng khử, do đó kí hiệu là Fe^3+^/Fe^2+^."
        },
        {
            "id": 14,
            "prompt": "Câu 14: Số amine bậc I trong số các chất: CH~3~NH~2~, CH~3~NH~3~Cl, (NH~2~)~2~CO, CH~3~NHCH~3~, CH~3~CH~2~NH~2~, (CH~3~)~3~N, C~6~H~5~NH~2~ (aniline) là:",
            "options": [
                "A. 6.",
                "*B. 3.*",
                "C. 4.",
                "D. 5."
            ],
            "solution": "Amine bậc I có dạng R-NH~2~: CH~3~NH~2~, CH~3~CH~2~NH~2~, C~6~H~5~NH~2~ (3 chất). Các chất còn lại: CH~3~NH~3~Cl (muối ammonium), (NH~2~)~2~CO (diamide - urê), CH~3~NHCH~3~ (amine bậc II), (CH~3~)~3~N (amine bậc III)."
        },
        {
            "id": 15,
            "prompt": "Câu 15: Hợp chất X có công thức cấu tạo CH~3~CH~2~COOCH~3~. X có tên gọi nào sau đây?",
            "options": [
                "A. Methyl acetate.",
                "*B. Methyl propionate.*",
                "C. Ethyl propionate.",
                "D. Propyl acetate."
            ],
            "solution": "Gốc alcohol -CH~3~ là methyl; gốc acid CH~3~CH~2~COO- là propionate. Do đó tên gọi của ester là methyl propionate."
        },
        {
            "id": 16,
            "prompt": "Câu 16: Khi nấu món canh từ thịt của tôm tép có nhiều mảng thịt đóng rắn lại. Hiện tượng trên gây ra bởi tính chất nào sau đây?",
            "options": [
                "A. Sự thủy phân protein bởi nhiệt độ.",
                "B. Kết tủa carbonate của các chất khoáng có trong vỏ.",
                "*C. Sự đông tụ protein bởi nhiệt độ.*",
                "D. Sự đông tụ protein bởi sự thay đổi pH."
            ],
            "solution": "Dưới tác dụng của nhiệt độ khi đun sôi, cấu trúc không gian bậc cao của các phân tử protein hòa tan trong thịt tôm tép/cua bị biến tính, các chuỗi polypeptide duỗi xoắn và liên kết lại với nhau tạo thành các mảng kết tủa đóng rắn nổi lên mặt nước (hiện tượng đông tụ protein bởi nhiệt độ)."
        }
    ],

    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1: Kevlar là một polymer có độ bền kéo rất cao. Loại vật liệu này được dùng để sản xuất áo chống đạn và mũ bảo hiểm cho quân đội. Kevlar được tổng hợp bằng phản ứng sau (các điều kiện có đủ):\n"
                      "n H~2~N-C~6~H~4~-NH~2~  +  n HOOC-C~6~H~4~-COOH  ->  [-NH-C~6~H~4~-NH-CO-C~6~H~4~-CO-]~n~  +  2n H~2~O\n"
                      "(1,4-diaminebenzene)          (Terephthalic acid)                     (Kevlar)",
            "image": "ao_chong_dan.png",
            "image_width": Cm(4.5),
            "statements": [
                ("a) Kevlar thuộc loại polymer tổng hợp.", "ĐÚNG"),
                ("b) Phản ứng tổng hợp Kevlar thuộc loại phản ứng trùng ngưng.", "ĐÚNG"),
                ("c) Cả hai chất terephthalic acid và 1,4-diaminebenzene đều là hợp chất tạp chức.", "SAI"),
                ("d) Trong một mắt xích Kevlar, phần trăm khối lượng nitrogen là 11,76% (Kết quả làm tròn đến hàng phần trăm). (Cho nguyên tử khối: C=12; O=16; H=1; N=14)", "ĐÚNG")
            ],
            "solution": "- a) ĐÚNG: Kevlar được tạo ra hoàn toàn bằng con đường tổng hợp hóa học nên thuộc loại polymer tổng hợp.\n"
                       "- b) ĐÚNG: Phản ứng kết hợp các monomer đồng thời giải phóng phân tử nhỏ (H~2~O) nên là phản ứng trùng ngưng.\n"
                       "- c) SAI: Cả 2 chất đều là hợp chất đa chức (chứa 2 nhóm -NH~2~ hoặc 2 nhóm -COOH giống nhau). Hợp chất tạp chức phải chứa các loại nhóm chức khác nhau trong cùng một phân tử.\n"
                       "- d) ĐÚNG: Công thức phân tử một mắt xích Kevlar là C~14~H~10~N~2~O~2~. Phân tử khối: M = 14*12 + 10*1 + 2*14 + 2*16 = 168 + 10 + 28 + 32 = 238. Phần trăm khối lượng nitrogen là: %N = (28 / 238) * 100% ≈ 11,76%."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Cho Pin Galvani Zn-Cu gồm điện cực kẽm và điện cực đồng được nối với nhau bởi cầu muối (thường chứa dung dịch KCl bão hòa) như hình sau:",
            "image": "pin_galvani.png",
            "image_width": Cm(8.5),
            "statements": [
                ("a) Tại anode xảy ra quá trình khử Zn, tại cathode xảy ra quá trình oxi hóa Cu^2+^.", "SAI"),
                ("b) Khi pin hoạt động thì nồng độ của Cu^2+^ giảm, nồng độ của Zn^2+^ tăng lên.", "ĐÚNG"),
                ("c) Dòng electron di chuyển từ thanh Zn qua dây dẫn đến thanh Cu.", "ĐÚNG"),
                ("d) Biết thế điện cực chuẩn của cặp oxi hoá - khử Zn^2+^/Zn và Cu^2+^/Cu là E°(Zn^2+^/Zn) = -0,76 V, E°(Cu^2+^/Cu) = +0,34 V. Sức điện động của pin Galvani Zn-Cu có giá trị bằng 0,42 (V).", "SAI")
            ],
            "solution": "- a) SAI: Trong pin điện hóa, tại cực âm (anode) xảy ra quá trình oxi hóa Zn: Zn -> Zn^2+^ + 2e; tại cực dương (cathode) xảy ra quá trình khử Cu^2+^: Cu^2+^ + 2e -> Cu.\n"
                       "- b) ĐÚNG: Cu^2+^ trong dung dịch nhận electron bám vào điện cực nên [Cu^2+^] giảm; Zn bị oxi hóa tan vào dung dịch nên [Zn^2+^] tăng.\n"
                       "- c) ĐÚNG: Dòng electron chạy từ cực âm Zn qua dây dẫn sang cực dương Cu.\n"
                       "- d) SAI: Sức điện động chuẩn: E°_pin = E°(cathode) - E°(anode) = E°(Cu^2+^/Cu) - E°(Zn^2+^/Zn) = 0,34 - (-0,76) = 1,10 (V), không phải 0,42 V."
        }
    ],

    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1: Tiến hành các thí nghiệm sau:\n"
                      "(1) Đun nóng hỗn hợp 3 mL isoamyl acetate và 5 mL dung dịch sulfuric acid 0,1 M.\n"
                      "(2) Đun nóng hỗn hợp 3 mL ethyl acetate và 5 mL dung dịch NaOH 0,2 M.\n"
                      "(3) Đun 5 gam mỡ động vật trong nước.\n"
                      "(4) Đun nóng hỗn hợp 3 mL ethyl bromide và 5 mL dung dịch NaOH 0,2 M.\n"
                      "(5) Cho 3 mL dung dịch CH~3~COOH 0,1 M vào ống nghiệm chứa 3 mL dung dịch NaOH 0,2 M.\n"
                      "(6) Đun nóng hỗn hợp 3 gam tristearin và 5 mL dung dịch NaOH 0,2 M.\n"
                      "Trong các thí nghiệm trên, có bao nhiêu thí nghiệm xảy ra phản ứng xà phòng hoá?",
            "answer": "2",
            "solution": "Phản ứng xà phòng hóa theo nghĩa rộng là phản ứng thủy phân ester trong môi trường kiềm (dung dịch base kiềm như NaOH, KOH):\n"
                       "- (1) Thủy phân ester trong môi trường acid (không phải xà phòng hóa).\n"
                       "- (2) Thủy phân ethyl acetate trong NaOH: CH~3~COOC~2~H~5~ + NaOH -> CH~3~COONa + C~2~H~5~OH (thỏa mãn).\n"
                       "- (3) Đun mỡ động vật trong nước nguyên chất không có xúc tác acid hay base kiềm (không xảy ra thủy phân đáng kể).\n"
                       "- (4) Thủy phân dẫn xuất halogen (không phải xà phòng hóa ester).\n"
                       "- (5) Phản ứng trung hòa acid - base (không phải xà phòng hóa).\n"
                       "- (6) Thủy phân chất béo tristearin trong dung dịch NaOH tạo muối sodium stearate (xà phòng) và glycerol (thỏa mãn).\n"
                       "-> Có 2 thí nghiệm xảy ra phản ứng xà phòng hóa ((2) và (6)).\n"
                       "*(Ghi chú: Nếu quy ước chặt chẽ chỉ dùng thuật ngữ 'xà phòng hóa' cho việc điều chế xà phòng từ chất béo thì đáp số là 1 (thí nghiệm (6)).)*"
        },
        {
            "id": 2,
            "prompt": "Câu 2: Trong những carbohydrate sau: glucose, fructose, saccharose, maltose, tinh bột và cellulose, có bao nhiêu carbohydrate tham gia phản ứng thuỷ phân trong môi trường acid?",
            "answer": "4",
            "solution": "Các disaccharide và polysaccharide bị thủy phân hoàn toàn trong môi trường acid nóng để tạo thành các monosaccharide tương ứng:\n"
                       "- Saccharose, maltose (disaccharide) -> thủy phân tạo monosaccharide.\n"
                       "- Tinh bột, cellulose (polysaccharide) -> thủy phân tạo glucose.\n"
                       "Glucose và fructose là monosaccharide nên không tham gia phản ứng thủy phân.\n"
                       "-> Có 4 carbohydrate tham gia phản ứng thủy phân."
        },
        {
            "id": 3,
            "prompt": "Câu 3: Khối lượng của một đoạn mạch capron ([-NH-(CH~2~)~5~-CO-]~n~) là 11 300 amu. Số lượng mắt xích trong đoạn mạch capron nêu trên là bao nhiêu? (Kết quả làm tròn đến hàng đơn vị) (Cho nguyên tử khối: C=12; O=16; H=1; N=14)",
            "answer": "100",
            "solution": "Công thức một mắt xích tơ capron là -NH-(CH~2~)~5~-CO-.\n"
                       "Khối lượng một mắt xích: M_mắt xích = 14 + 1 + 5*(12 + 2) + 12 + 16 = 15 + 70 + 28 = 113 (amu).\n"
                       "Số lượng mắt xích (hệ số trùng hợp): n = 11 300 / 113 = 100."
        },
        {
            "id": 4,
            "prompt": "Câu 4: Điện phân 1 lít dung dịch CuSO~4~ 0,1M (điện cực trơ) cho đến khi ở anode thu được 2,479 lít khí (đkc) thì ở cathode thu được bao nhiêu gam kim loại? (Kết quả làm tròn đến hàng phần mười) (Cho nguyên tử khối: Cu=64; O=16; H=1; S=32)",
            "answer": "6,4",
            "solution": "1. Số mol CuSO~4~ ban đầu: n(Cu^2+^) = 1 * 0,1 = 0,1 mol.\n"
                       "2. Khí thu được ở anode (cực dương) là O~2~ do nước bị oxi hóa:\n"
                       "   2H~2~O -> O~2~ + 4H^+^ + 4e\n"
                       "   n(O~2~) = 2,479 / 24,79 = 0,1 mol -> Số mol electron trao đổi ở anode: n_e = 4 * 0,1 = 0,4 mol.\n"
                       "3. Tại cathode (cực âm) xảy ra quá trình khử Cu^2+^:\n"
                       "   Cu^2+^ + 2e -> Cu\n"
                       "   Số mol e cần thiết để khử hết Cu^2+^ là: 2 * n(Cu^2+^) = 2 * 0,1 = 0,2 mol < 0,4 mol.\n"
                       "   Do đó, toàn bộ Cu^2+^ bị khử hết thành kim loại Cu bám trên bề mặt cathode (sau đó nước tiếp tục bị khử sinh khí H~2~).\n"
                       "4. Khối lượng Cu thu được: m(Cu) = 0,1 * 64 = 6,4 gam."
        }
    ],

    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1: PVC là viết tắt của poly(vinyl chloride) là một loại nhựa nhiệt dẻo. PVC là loại vật liệu nhân tạo được tổng hợp sớm nhất và mở ra một sự tiện lợi lớn lao trong sinh hoạt và sản xuất cho nhân loại với độ bền cao và giá thành rẻ, mới đầu khi được khám phá ra, PVC có rất nhiều nhược điểm như cứng, kém ổn định và khó gia công. Mãi đến năm 1926 khi xuất hiện phương pháp hoá dẻo PVC thì chúng mới dần được nhân rộng.\n"
                      "Viết phương trình hóa học tổng hợp PVC từ vinyl chloride?",
            "num_lines": 5,
            "solution": "Phương trình hóa học trùng hợp vinyl chloride:\n"
                       "   n CH~2~=CH-Cl  --[t°, p, xt]-->  [-CH~2~-CH(Cl)-]~n~\n"
                       "   (vinyl chloride)                           (poly(vinyl chloride) - PVC)\n"
                       "Đặc điểm: Phản ứng diễn ra bằng sự kết hợp liên tiếp các phân tử monome vinyl chloride nhờ việc bẻ gãy liên kết đôi C=C để hình thành các mạch polymer dài."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Cho E°_pin(Mg-Ag) = 3,17 V và E°(Mg^2+^/Mg) = -2,37 V. Tính thế điện cực chuẩn của cặp oxi hóa – khử Ag^+^/Ag (E°(Ag^+^/Ag))?",
            "num_lines": 6,
            "solution": "1. Trong pin Galvani Mg - Ag, kim loại Mg có tính khử mạnh hơn kim loại Ag nên:\n"
                       "   - Anode (cực âm): điện cực magnesium (Mg).\n"
                       "   - Cathode (cực dương): điện cực bạc (Ag).\n"
                       "2. Sức điện động chuẩn của pin được tính theo công thức:\n"
                       "   E°_pin = E°(cathode) - E°(anode) = E°(Ag^+^/Ag) - E°(Mg^2+^/Mg)\n"
                       "3. Thay các số liệu vào công thức:\n"
                       "   3,17 = E°(Ag^+^/Ag) - (-2,37)\n"
                       "   3,17 = E°(Ag^+^/Ag) + 2,37\n"
                       "   => E°(Ag^+^/Ag) = 3,17 - 2,37 = +0,80 (V).\n"
                       "Kết luận: Thế điện cực chuẩn của cặp oxi hóa - khử Ag^+^/Ag là +0,80 V."
        },
        {
            "id": 3,
            "prompt": "Câu 3: Hiện nay, mạ điện được sử dụng rộng rãi trong thực tế, mạ điện là quá trình phủ một lớp kim loại lên bề mặt kim loại khác bằng phương pháp điện phân. Giả sử người ta cần mạ Bạc lên một mặt của một chiếc đĩa kim loại hình tròn có bán kính 30 cm, độ dày lớp mạ là 0,01 mm bằng phương pháp điện phân dung dịch AgNO~3~. Nếu được cung cấp nguồn điện một chiều có cường độ dòng điện I = 2,5 A thì thời gian cần dùng là t giờ. Biết khối lượng riêng của Ag là 10,5 g/cm^3^; π = 3,14; Hằng số Faraday (F = 96500 C/mol); Hiệu suất điện phân là 100%. Tính t? (Cho nguyên tử khối: Ag=108; O=16; N=14)",
            "num_lines": 8,
            "solution": "1. Tính diện tích một mặt chiếc đĩa hình tròn cần mạ:\n"
                       "   S = π * R^2 = 3,14 * 30^2 = 3,14 * 900 = 2826 (cm^2).\n"
                       "2. Đổi chiều dày lớp mạ sang cm:\n"
                       "   d = 0,01 mm = 0,001 cm.\n"
                       "3. Thể tích lớp bạc mạ trên bề mặt:\n"
                       "   V = S * d = 2826 * 0,001 = 2,826 (cm^3).\n"
                       "4. Khối lượng Bạc cần mạ:\n"
                       "   m(Ag) = V * D = 2,826 * 10,5 = 29,673 (gam).\n"
                       "5. Áp dụng định luật Faraday về điện phân m = (A * I * t) / (n * F):\n"
                       "   Quá trình khử: Ag^+^ + 1e -> Ag  =>  n = 1 trao đổi electron.\n"
                       "   Thời gian điện phân tính bằng giây (t_s):\n"
                       "   t_s = (m * n * F) / (A * I) = (29,673 * 1 * 96500) / (108 * 2,5) = 2863444,5 / 270 ≈ 10605,35 (giây).\n"
                       "6. Đổi thời gian ra đơn vị giờ:\n"
                       "   t = 10605,35 / 3600 ≈ 2,946 (giờ) ≈ 2,95 giờ (tức 2 giờ 56 phút 45 giây).\n"
                       "Kết luận: Thời gian cần dùng là khoảng 2,95 giờ (hoặc 2,946 giờ)."
        }
    ]
}


def generate_docx(is_teacher=False):
    """Tạo file docx theo chuẩn Thầy TRẦN VĂN THẠNH."""
    doc = Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line=DATA_EXAM_001["school_line"],
        exam_title=DATA_EXAM_001["exam_title"],
        subject_line=DATA_EXAM_001["subject_line"],
        num_pages=DATA_EXAM_001["num_pages"],
        duration=DATA_EXAM_001["duration"],
        exam_code=DATA_EXAM_001["exam_code"],
        atomic_masses=DATA_EXAM_001["atomic_masses"],
        is_teacher=is_teacher
    )

    if is_teacher:
        add_compact_summary_answers(doc)

    # PHẦN I
    add_section_header(
        doc,
        title="PHẦN I: (4,0 điểm) Câu trắc nghiệm nhiều phương án lựa chọn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    for item in DATA_EXAM_001["part1"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        render_options_no_table(doc, item["options"], is_teacher=is_teacher)
        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN II
    add_section_header(
        doc,
        title="PHẦN II: (2,0 điểm) Câu trắc nghiệm đúng sai.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    for item in DATA_EXAM_001["part2"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        if "image" in item:
            img_path = os.path.join(DIR_IMAGES, item["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(3)
                p_img.add_run().add_picture(img_path, width=item.get("image_width", Cm(8.0)))

        for stmt, ans in item["statements"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt, ans=ans, is_teacher=is_teacher, base_size=10.5)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN III
    add_section_header(
        doc,
        title="PHẦN III: (2,0 điểm) Câu trắc nghiệm yêu cầu trả lời ngắn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Thí sinh điền kết quả vào ô trả lời tương ứng."
    )

    for item in DATA_EXAM_001["part3"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(2)
        if is_teacher:
            r_ans_lbl = p_ans.add_run("Trả lời: ")
            set_run_style(r_ans_lbl, font_size=10.5, bold=True)
            r_ans_val = p_ans.add_run(item["answer"])
            set_run_style(r_ans_val, font_size=10.5, bold=True, color=COLOR_CORRECT)
        else:
            r_ans = p_ans.add_run("Trả lời: ............................................................ ")
            set_run_style(r_ans, font_size=10.5, italic=True)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN IV
    add_section_header(
        doc,
        title="PHẦN IV: (3,0 điểm) Tự luận.",
        subtitle="Thí sinh trình bày chi tiết lời giải từ câu 1 đến câu 3 vào giấy làm bài thi."
    )

    for item in DATA_EXAM_001["part4"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])
        elif not is_teacher:
            num_lines = item.get("num_lines", 8)
            add_student_writing_lines(doc, num_lines=num_lines)

    # Kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(0)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(r_end, font_size=11, bold=True)

    suffix = "_GiaoVien.docx" if is_teacher else "_HocSinh.docx"
    filename = f"De_Kiem_Tra_CK1_Hoa_12_HuongTra_Ma001{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
