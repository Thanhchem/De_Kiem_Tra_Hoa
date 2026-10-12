# -*- coding: utf-8 -*-
"""
Script tạo file Word (.docx) chuẩn "FORM ĐỀ KIỂM TRA" cho:
ĐỀ KIỂM TRA SỐ 8 - CHƯƠNG 3: HỢP CHẤT CHỨA NITROGEN (HÓA HỌC 12)
Gồm 2 phiên bản:
1. De_Kiem_Tra_Hoa_12_De08.docx (Bản Học Sinh)
2. De_Kiem_Tra_Hoa_12_De08_GiaoVien.docx (Bản Giáo Viên có lời giải chi tiết)
Tuân thủ 100% quy chuẩn: Header & Footer, Tab Stops (NO TABLE trong đề/lời giải), Sub/Superscript công thức hóa học.
"""

import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
IMG_DIR = r"c:\Antigravity_Thanh\San_Pham\images_crop_de08"

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

def setup_page(doc):
    for sec in doc.sections:
        sec.top_margin = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
        sec.left_margin = Cm(1.5)
        sec.right_margin = Cm(1.5)
        sec.header_distance = Cm(0.6)
        sec.footer_distance = Cm(0.6)
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)

def set_cell_border_bottom(cell, color="8BC390", sz="6"):
    tcPr = cell._tc.get_or_add_tcPr()
    xml = f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/></w:tcBorders>'
    tcPr.append(parse_xml(xml))

def set_cell_border_top(cell, color="8BC390", sz="6"):
    tcPr = cell._tc.get_or_add_tcPr()
    xml = f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/></w:tcBorders>'
    tcPr.append(parse_xml(xml))

def set_cell_margins_zero(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    xml = f'<w:tcMar {nsdecls("w")}><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>'
    tcPr.append(parse_xml(xml))

def setup_header_footer(doc):
    sec = doc.sections[0]
    
    # HEADER
    hdr = sec.header
    htable = hdr.add_table(rows=1, cols=2, width=Cm(18.0))
    htable.autofit = False
    htable.columns[0].width = Cm(9.0)
    htable.columns[1].width = Cm(9.0)
    
    cell_l, cell_r = htable.rows[0].cells[0], htable.rows[0].cells[1]
    set_cell_margins_zero(cell_l)
    set_cell_margins_zero(cell_r)
    set_cell_border_bottom(cell_l, color=COLOR_HF_GREEN_HEX, sz="6")
    set_cell_border_bottom(cell_r, color=COLOR_HF_GREEN_HEX, sz="6")
    
    p_hl = cell_l.paragraphs[0]
    p_hl.paragraph_format.space_before = Pt(0)
    p_hl.paragraph_format.space_after = Pt(2)
    p_hl.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_hl = p_hl.add_run("Tài liệu lưu hành nội bộ")
    set_run_style(r_hl, font_size=10, italic=True, color=COLOR_HF_RED)
    
    p_hr = cell_r.paragraphs[0]
    p_hr.paragraph_format.space_before = Pt(0)
    p_hr.paragraph_format.space_after = Pt(2)
    p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_hr1 = p_hr.add_run("Thầy ")
    set_run_style(r_hr1, font_size=10, color=COLOR_HF_RED)
    r_hr2 = p_hr.add_run("TRẦN VĂN THẠNH")
    set_run_style(r_hr2, font_size=10, bold=True, color=COLOR_HF_RED)
    r_hr3 = p_hr.add_run(" - 0777.470.803")
    set_run_style(r_hr3, font_size=10, color=COLOR_HF_RED)
    
    # FOOTER
    ftr = sec.footer
    ftable = ftr.add_table(rows=1, cols=2, width=Cm(18.0))
    ftable.autofit = False
    ftable.columns[0].width = Cm(14.0)
    ftable.columns[1].width = Cm(4.0)
    
    cell_fl, cell_fr = ftable.rows[0].cells[0], ftable.rows[0].cells[1]
    set_cell_margins_zero(cell_fl)
    set_cell_margins_zero(cell_fr)
    set_cell_border_top(cell_fl, color=COLOR_HF_GREEN_HEX, sz="6")
    set_cell_border_top(cell_fr, color=COLOR_HF_GREEN_HEX, sz="6")
    
    p_fl = cell_fl.paragraphs[0]
    p_fl.paragraph_format.space_before = Pt(2)
    p_fl.paragraph_format.space_after = Pt(0)
    p_fl.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_fl = p_fl.add_run("CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế   |   CS2: 24 Đặng Thái Thân, TP Huế")
    set_run_style(r_fl, font_size=9.5, italic=True, color=COLOR_HF_RED)
    
    p_fr = cell_fr.paragraphs[0]
    p_fr.paragraph_format.space_before = Pt(2)
    p_fr.paragraph_format.space_after = Pt(0)
    p_fr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_page_number_field(p_fr, font_size=10, bold=True, color=COLOR_HF_RED)

def add_header_block(doc, is_teacher=False):
    lines = [
        ("SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ", "ĐỀ KIỂM TRA ĐỊNH KỲ CHƯƠNG 3 (HỢP CHẤT CHỨA NITROGEN)", True, True, 10.5, 11),
        ("ĐỀ CHÍNH THỨC", "MÔN: HOÁ HỌC – LỚP 12", True, True, 11, 11),
        ("(Đề thi có 04 trang)", "Thời gian làm bài: 50 phút (không kể thời gian phát đề)", False, True, 10, 10.5),
        ("", "Mã đề thi: 108", False, True, 10, 11)
    ]
    
    for left_text, right_text, l_bold, r_bold, l_size, r_size in lines:
        p = doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(18.0), WD_TAB_ALIGNMENT.RIGHT)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.05
        
        if left_text:
            r_l = p.add_run(left_text)
            set_run_style(r_l, font_size=l_size, bold=l_bold, italic=("Đề thi có" in left_text))
        
        p.add_run("\t")
        
        if right_text:
            r_r = p.add_run(right_text)
            set_run_style(r_r, font_size=r_size, bold=r_bold, italic=("Thời gian" in right_text), color=COLOR_PRIMARY if "CHƯƠNG 3" in right_text else None)

    if is_teacher:
        p_t = doc.add_paragraph()
        p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_t.paragraph_format.space_before = Pt(2)
        p_t.paragraph_format.space_after = Pt(2)
        r_t = p_t.add_run("(BẢN GIÁO VIÊN — CÓ LỜI GIẢI CHI TIẾT)")
        set_run_style(r_t, font_size=11, bold=True, color=COLOR_CORRECT)

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(4)
    p_info.paragraph_format.space_after = Pt(2)
    r_info = p_info.add_run("Họ, tên thí sinh: .......................................................................... Lớp: ................ Số báo danh: ...............")
    set_run_style(r_info, font_size=10.5, italic=True)
    
    p_mass = doc.add_paragraph()
    p_mass.paragraph_format.space_before = Pt(0)
    p_mass.paragraph_format.space_after = Pt(6)
    r_mass = p_mass.add_run("(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; S = 32; Cl = 35,5; Fe = 56; Cu = 64).")
    set_run_style(r_mass, font_size=10, italic=True)

def add_part_heading(doc, part_title, part_desc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.15
    r_title = p.add_run(part_title)
    set_run_style(r_title, font_size=11.5, bold=True, color=COLOR_PRIMARY)
    
    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_before = Pt(0)
    p_desc.paragraph_format.space_after = Pt(4)
    r_desc = p_desc.add_run(part_desc)
    set_run_style(r_desc, font_size=10.5, italic=True)

def add_solution(doc, sol_text):
    p_sol = doc.add_paragraph()
    p_sol.paragraph_format.space_before = Pt(2)
    p_sol.paragraph_format.space_after = Pt(4)
    p_sol.paragraph_format.left_indent = Cm(0.6)
    p_sol.paragraph_format.line_spacing = 1.15
    
    r_hdr = p_sol.add_run("Lời giải chi tiết:\n")
    set_run_style(r_hdr, font_size=10.5, bold=True, italic=True, color=COLOR_SOLUTION)
    
    lines = sol_text.strip().split("\n")
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        if i > 0:
            p_line = doc.add_paragraph()
            p_line.paragraph_format.space_before = Pt(0)
            p_line.paragraph_format.space_after = Pt(1.5)
            p_line.paragraph_format.left_indent = Cm(0.6)
            p_line.paragraph_format.line_spacing = 1.15
            add_chem_run(p_line, line, base_size=10.5, italic=False)
        else:
            add_chem_run(p_sol, line, base_size=10.5, italic=False)

def add_mcq_options(doc, opt_a, opt_b, opt_c, opt_d, correct_idx=None, is_teacher=False):
    opts = [('A', opt_a), ('B', opt_b), ('C', opt_c), ('D', opt_d)]
    max_len = max(len(text) for _, text in opts)
    
    if max_len < 22:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.tab_stops.add_tab_stop(Cm(4.5), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(13.5), WD_TAB_ALIGNMENT.LEFT)
        
        for idx, (label, text) in enumerate(opts):
            if idx > 0:
                p.add_run("\t")
            is_corr = (is_teacher and idx == correct_idx)
            r_lbl = p.add_run(f"{label}. ")
            set_run_style(r_lbl, font_size=11, bold=True, color=COLOR_CORRECT if is_corr else None)
            add_chem_run(p, text, base_size=11, bold=is_corr, color=COLOR_CORRECT if is_corr else None)
    elif max_len < 46:
        for row in [(0, 1), (2, 3)]:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
            
            for col_idx, opt_num in enumerate(row):
                if col_idx > 0:
                    p.add_run("\t")
                label, text = opts[opt_num]
                is_corr = (is_teacher and opt_num == correct_idx)
                r_lbl = p.add_run(f"{label}. ")
                set_run_style(r_lbl, font_size=11, bold=True, color=COLOR_CORRECT if is_corr else None)
                add_chem_run(p, text, base_size=11, bold=is_corr, color=COLOR_CORRECT if is_corr else None)
    else:
        for idx, (label, text) in enumerate(opts):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1.5)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.left_indent = Cm(0.6)
            is_corr = (is_teacher and idx == correct_idx)
            r_lbl = p.add_run(f"{label}. ")
            set_run_style(r_lbl, font_size=11, bold=True, color=COLOR_CORRECT if is_corr else None)
            add_chem_run(p, text, base_size=11, bold=is_corr, color=COLOR_CORRECT if is_corr else None)

def build_docx_de08(is_teacher=False):
    doc = docx.Document()
    setup_page(doc)
    setup_header_footer(doc)
    add_header_block(doc, is_teacher=is_teacher)
    
    # ==================== PHẦN I ====================
    add_part_heading(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )
    
    part1_data = [
        (
            1,
            "(SBT – KNTT). Amine nào sau đây ở trạng thái lỏng ở nhiệt độ phòng?",
            None,
            "Methylamine.", "Ethylamine.", "Dimethylamine.", "Aniline.",
            3,
            "Ở điều kiện thường, methylamine, dimethylamine, trimethylamine và ethylamine là chất khí có mùi khai khó chịu, độc, tan nhiều trong nước.\nAniline (C~6~H~5~NH~2~) là chất lỏng, sôi ở 184 °C, không màu, rất độc, ít tan trong nước, nặng hơn nước."
        ),
        (
            2,
            "Khi alkylamine bậc một phản ứng với nitrous acid (HNO~2~) qua phương trình sau:\n    RNH~2~ + HNO~2~ -> ROH + N~2~ + H~2~O  (R: gốc alkyl)\nTrong phản ứng trên thì alkylamine đóng vai trò là chất",
            None,
            "oxi hóa.", "base.", "khử.", "acid.",
            2,
            "Trong phân tử alkylamine RNH~2~, nguyên tử nitrogen có số oxi hóa -3 tăng lên 0 trong N~2~.\nNitrous acid (HNO~2~) chứa N(+3) giảm xuống 0 trong N~2~.\nDo đó, alkylamine đóng vai trò là chất khử, còn HNO~2~ là chất oxi hóa."
        ),
        (
            3,
            "Trong phân tử amine: N còn một cặp e chưa liên kết nên có thể tạo ra liên kết ............. (giống NH~3~) do đó các amine có tính chất hóa học tương tự ammonia.",
            ("cau03_amine_hplus.png", 6.8),
            "Cụm từ điền vào khoảng trống là:\nA. cho nhận.", "B. ion.", "C. cộng hóa trị thường.", "D. hydrogen.",  # Special handling below
            0,
            "Cặp electron tự do chưa liên kết trên nguyên tử nitrogen của amine có thể nhường vào orbital trống của ion H^+^ để tạo liên kết phối trí (liên kết cho - nhận), tương tự như phân tử ammonia NH~3~. Do đó amine thể hiện tính base."
        ),
        (
            4,
            "Chất nào sau đây là amino acid?",
            None,
            "Acid glutaric.", "Acid glutamic.", "Glycerol.", "Aniline.",
            1,
            "Glutamic acid (HOOC-CH~2~-CH~2~-CH(NH~2~)-COOH) là một amino acid (axit 2-aminopentanedioic).\nGlutaric acid là dicarboxylic acid; Glycerol là polyalcohol; Aniline là arylamine."
        ),
        (
            5,
            "(SBT – KNTT). Quá trình di chuyển của các amino acid trong điện trường được gọi là",
            None,
            "sự điện di.", "sự điện li.", "sự điện phân.", "sự điện giải.",
            0,
            "Sự điện di (electrophoresis) là hiện tượng chuyển động của các hạt tích điện (như amino acid, protein) dưới tác dụng của một điện trường ngoài."
        ),
        (
            6,
            "(SBT – CTST). Amino acid không có loại phản ứng nào sau đây?",
            None,
            "Phản ứng với acid và base.", "Phản ứng ester hoá.", "Phản ứng trùng ngưng.", "Phản ứng thủy phân.",
            3,
            "Amino acid là hợp chất đơn phân tử chứa đồng thời nhóm chức carboxyl (-COOH) và amino (-NH~2~), không chứa liên kết peptide hay liên kết ester nên không tham gia phản ứng thủy phân. Phản ứng thủy phân là phản ứng đặc trưng của peptide, protein, ester..."
        ),
        (
            7,
            "Số liên kết peptide trong hợp chất pentapeptide là",
            None,
            "3.", "4.", "5.", "6.",
            1,
            "Pentapeptide là peptide được cấu tạo từ 5 gốc α-amino acid, do đó số liên kết peptide bằng 5 - 1 = 4."
        ),
        (
            8,
            "(SBT – KNTT). Loại liên kết được hình thành giữa các amino acid trong peptide được gọi là",
            None,
            "liên kết ion.", "liên kết hydrogen.", "liên kết peptide.", "liên kết cộng hoá trị.",
            2,
            "Liên kết peptide (-CO-NH-) là liên kết amide hình thành giữa nhóm carboxyl của một α-amino acid với nhóm amino của một α-amino acid khác."
        ),
        (
            9,
            "Protein hình cầu tan được trong nước tạo dung dịch keo như",
            None,
            "α-keratin (ở tóc, móng).", "Collagen (ở da, sụn).", "Albumin (lòng trắng trứng).", "Myosin (ở cơ bắp).",
            2,
            "Albumin trong lòng trắng trứng là protein dạng hình cầu, tan được trong nước tạo dung dịch keo.\nKeratin, collagen, myosin là protein dạng hình sợi không tan trong nước."
        ),
        (
            10,
            "Ứng dụng nào sau đây không phải là của protein?",
            None,
            "Là thành phần tạo nên chất dẻo.", "Là thành phần cấu tạo tế bào.", "Là cơ sở tạo nên sự sống.", "Là dinh dưỡng thức ăn.",
            0,
            "Protein đóng vai trò cấu trúc tế bào, vận chuyển chất và cung cấp dinh dưỡng thiết yếu cho cơ thể sống. Protein không được dùng để chế tạo chất dẻo (vật liệu nhựa công nghiệp)."
        ),
        (
            11,
            "(SBT – CTST). Trong dãy các chất sau đây, tính base của amine thể hiện qua phản ứng với các chất:",
            None,
            "HCl, H~2~SO~4~, CuCl~2~.", "Cl~2~, H~2~SO~4~, FeCl~3~.", "NaOH, HCl, FeCl~3~.", "O~2~, HCl, CuCl~2~.",
            0,
            "Tính base của amine thể hiện qua phản ứng với acid vô cơ (HCl, H~2~SO~4~ tạo muối ammonium) và phản ứng tạo kết tủa hydroxide/phức chất với dung dịch muối (CuCl~2~, FeCl~3~). Tuy nhiên dãy B có Cl~2~ (chất oxi hóa), dãy C có NaOH (base), dãy D có O~2~. Dãy phù hợp nhất là HCl, H~2~SO~4~, CuCl~2~."
        ),
        (
            12,
            "(SBT – CTST). Cơ thể người mã hoá được loại amino acid nào sau đây trong tổng hợp protein cho cơ thể?",
            ("cau12_options.png", 14.2),
            None, None, None, None,  # Options are inside image
            1,
            "Protein của cơ thể người được tạo nên từ 20 loại α-amino acid chuẩn do mã di truyền trực tiếp mã hóa.\nCông thức ở phương án B là CH~3~-CH(NH~3~^+^)-COO^-^ (alanine) là α-amino acid.\nPhương án A (β-alanine) và C (γ-aminobutyric acid / GABA) không tham gia cấu tạo protein theo mã di truyền."
        ),
        (
            13,
            "(SBT – KNTT). Trong phản ứng của amine với Cu(OH)~2~, sản phẩm có màu xanh đặc trưng là do",
            None,
            "sự tạo phức giữa Cu^2+^ và -NH~2~.", "sự oxi hoá Cu^2+^ khi gặp amine.", "phản ứng giữa OH^-^ và -NH~2~.", "giải phóng H~2~ khi phản ứng.",
            0,
            "Các alkylamine đơn giản dùng cặp electron tự do trên nguyên tử N tạo liên kết phối trí với ion Cu^2+^, hình thành ion phức [Cu(RNH~2~)~4~]^2+^ có màu xanh lam đặc trưng."
        ),
        (
            14,
            "(SBT – CTST). Cơ thể người sử dụng phân tử nào sau đây để xây dựng protein?",
            None,
            "Tinh bột.", "Chất béo.", "Amino acid.", "Acid béo.",
            2,
            "Amino acid là các viên gạch xây dựng (monomer) nên các đại phân tử polypeptide và protein trong cơ thể sinh vật."
        ),
        (
            15,
            "(SBT – CTST). Các hợp chất amine, amino acid, peptide và protein có nhóm chức chung nào sau đây?",
            None,
            "Carboxyl.", "Amide.", "Ammonium.", "Amine hoặc amino.",
            3,
            "Nhóm chức chứa nguyên tố nitrogen có mặt xuyên suốt trong các phân tử amine, amino acid, peptide và protein là nhóm amine (hoặc amino: -NH~2~, -NH-, =N-)."
        ),
        (
            16,
            "Amine và amino acid nào sau đây có cùng bậc?",
            None,
            "Dimethylamine và alanine.", "Dimethylamine và glutamic acid.", "Dimethylamine và proline.", "Trimethylamine và lysine.",
            2,
            "Dimethylamine ((CH~3~)~2~NH) là amine bậc 2.\nProline là amino acid vòng có nhóm amino bậc 2 (-NH- gắn với 2 nguyên tử C trong khung dị vòng pyrrolidine).\nDo đó dimethylamine và proline có cùng bậc 2."
        ),
        (
            17,
            "(SBT – CTST). Tính chất hoá học của methylamine và alanine đều có",
            None,
            "phản ứng với ninhydrin.", "phản ứng thuỷ phân.", "phản ứng với dung dịch HCl.", "phản ứng với dung dịch NaOH.",
            2,
            "Cả methylamine (CH~3~NH~2~) và alanine (CH~3~CH(NH~2~)COOH) đều có nhóm amino mang tính base nên đều phản ứng được với dung dịch acid HCl tạo muối ammonium.\nMethylamine không phản ứng với NaOH, không có phản ứng màu ninhydrin."
        ),
        (
            18,
            "Cho các phát biểu sau:\n"
            "(a) Trong dung dịch, glycine tồn tại chủ yếu ở dạng ion lưỡng cực.\n"
            "(b) Amino acid là chất rắn kết tinh, dễ tan trong nước.\n"
            "(c) Phản ứng thế bromine vào vòng thơm của aniline dễ hơn benzene.\n"
            "(d) Hợp chất H~2~N-CH~2~-COO-CH~3~ tác dụng được với dung dịch NaOH và dung dịch HCl.\n"
            "(e) Triolein và protein có cùng thành phần nguyên tố.\n"
            "(g) Trong phân tử, các amino acid đều chỉ có một nhóm NH~2~ và một nhóm COOH.\n"
            "Số phát biểu đúng là:",
            None,
            "4.", "5.", "3.", "6.",
            0,
            "(a) Đúng, trong dung dịch glycine tồn tại chủ yếu ở dạng ion lưỡng cực ^+^H~3~N-CH~2~-COO^-^.\n"
            "(b) Đúng, amino acid tồn tại dạng ion lưỡng cực tương tự hợp chất ion nên là chất rắn kết tinh, nhiệt độ nóng chảy cao và dễ tan trong nước.\n"
            "(c) Đúng, nhóm -NH~2~ đẩy electron làm tăng mật độ electron trong nhân thơm, phản ứng thế Br~2~ xảy ra dễ dàng hơn benzene tạo kết tủa trắng.\n"
            "(d) Đúng, methyl glycinate vừa có nhóm -NH~2~ (tác dụng với HCl) vừa có nhóm ester (bị thủy phân bởi NaOH).\n"
            "(e) Sai, triolein chứa C, H, O còn protein bắt buộc có C, H, O, N.\n"
            "(g) Sai, amino acid có thể có nhiều nhóm -NH~2~ hoặc nhiều nhóm -COOH (ví dụ Lysine có 2 nhóm NH~2~, Glutamic acid có 2 nhóm COOH).\n"
            "=> Có 4 phát biểu đúng: (a), (b), (c), (d)."
        )
    ]
    
    for item in part1_data:
        num, text, img_info, oa, ob, oc, od, corr_idx, sol = item
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        
        r_num = p.add_run(f"Câu {num}. ")
        set_run_style(r_num, font_size=11, bold=True)
        add_chem_run(p, text, base_size=11)
        
        # If question has an image
        if img_info:
            img_file, img_w = img_info
            img_path = os.path.join(IMG_DIR, img_file)
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(3)
                p_img.add_run().add_picture(img_path, width=Cm(img_w))
        
        # Options
        if num == 3:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(1)
            p_sub.paragraph_format.space_after = Pt(2)
            add_chem_run(p_sub, "Cụm từ điền vào khoảng trống là:", base_size=11)
            add_mcq_options(doc, "cho nhận.", "ion.", "cộng hóa trị thường.", "hydrogen.", correct_idx=corr_idx, is_teacher=is_teacher)
        elif num == 12:
            # Options already shown in image cau12_options.png
            p_opts = doc.add_paragraph()
            p_opts.paragraph_format.space_before = Pt(1)
            p_opts.paragraph_format.space_after = Pt(3)
            p_opts.paragraph_format.tab_stops.add_tab_stop(Cm(4.5), WD_TAB_ALIGNMENT.LEFT)
            p_opts.paragraph_format.tab_stops.add_tab_stop(Cm(9.0), WD_TAB_ALIGNMENT.LEFT)
            p_opts.paragraph_format.tab_stops.add_tab_stop(Cm(13.5), WD_TAB_ALIGNMENT.LEFT)
            
            opts_12 = [("A", "Cấu tạo A."), ("B", "Cấu tạo B."), ("C", "Cấu tạo C."), ("D", "Cấu tạo D.")]
            for idx, (lbl, txt) in enumerate(opts_12):
                if idx > 0:
                    p_opts.add_run("\t")
                is_c = (is_teacher and idx == corr_idx)
                r_l = p_opts.add_run(f"{lbl}. ")
                set_run_style(r_l, font_size=11, bold=True, color=COLOR_CORRECT if is_c else None)
                r_t = p_opts.add_run(txt)
                set_run_style(r_t, font_size=11, bold=is_c, color=COLOR_CORRECT if is_c else None)
        else:
            add_mcq_options(doc, oa, ob, oc, od, correct_idx=corr_idx, is_teacher=is_teacher)
            
        if is_teacher and sol:
            add_solution(doc, sol)

    # ==================== PHẦN II ====================
    add_part_heading(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )
    
    part2_data = [
        (
            1,
            "(SBT – KNTT). Các phát biểu về tính chất hoá học của dung dịch aniline:",
            None,
            [
                ("a", "Dung dịch aniline làm quỳ tím chuyển sang màu xanh.", False),
                ("b", "Dung dịch aniline tạo kết tủa trắng khi thêm vào nước bromine.", True),
                ("c", "Aniline phản ứng với HCl tạo phenylammonium chloride.", True),
                ("d", "Aniline phản ứng với HNO~2~ tạo muối phenyldiazonium.", True)
            ],
            (
                "a) Sai. Do ảnh hưởng hút electron của nhân benzene, tính base của aniline rất yếu không làm đổi màu quỳ tím.\n"
                "b) Đúng. C~6~H~5~NH~2~ + 3Br~2~ -> 2,4,6-tribromoaniline↓ (trắng) + 3HBr.\n"
                "c) Đúng. C~6~H~5~NH~2~ + HCl -> C~6~H~5~NH~3~Cl (phenylammonium chloride).\n"
                "d) Đúng. Ở 0 – 5 °C: C~6~H~5~NH~2~ + HNO~2~ + HCl -> [C~6~H~5~N~2~]^+^Cl^-^ + 2H~2~O."
            )
        ),
        (
            2,
            "(SBT – CTST). Cho một số hợp chất chứa nguyên tố nitrogen như sau:",
            ("p2_cau2_compounds.png", 13.5),
            [
                ("a", "Có 1 hợp chất hữu cơ thuộc loại amine.", True),
                ("b", "Có 2 hợp chất hữu cơ thuộc loại amino acid.", False),
                ("c", "Có 1 hợp chất hữu cơ đơn chức, 2 hợp chất tạp chức, 3 hợp chất chứa nhóm α-NH~2~.", False),
                ("d", "Trong 3 hợp chất trên chỉ có 1 liên kết peptide.", True)
            ],
            (
                "Chất (1) là ethylamine CH~3~-CH~2~-NH~2~ (amine đơn chức).\n"
                "Chất (2) là β-alanine H~2~N-CH~2~-CH~2~-COOH (amino acid tạp chức, nhóm -NH~2~ ở vị trí β).\n"
                "Chất (3) là dipeptide Gly-Ala H~2~N-CH~2~-CONH-CH(CH~3~)-COOH (tạp chức, có 1 liên kết peptide và nhóm α-NH~2~ ở đầu N).\n"
                "a) Đúng. Chỉ có chất (1) là amine.\n"
                "b) Sai. Chỉ có chất (2) là amino acid; chất (3) là dipeptide.\n"
                "c) Sai. Chất (2) có nhóm β-NH~2~, chất (1) không có nhóm carboxyl, chỉ có chất (3) có nhóm α-NH~2~. Vậy không có 3 hợp chất chứa nhóm α-NH~2~.\n"
                "d) Đúng. Chỉ có dipeptide (3) chứa đúng 1 liên kết peptide -CONH-."
            )
        ),
        (
            3,
            "Acetyl Hexapeptide-8 là một hợp chất peptide. Chất này được biết đến là một thành phần có khả năng liên kết với nước để phục hồi da và là một chất dẫn truyền thần kinh làm giảm nếp nhăn hình thành do lão hóa hoặc bởi các biểu hiện thay đổi trên cơ mặt gây ra.",
            ("p2_cau3_acetyl_hexapeptide.png", 14.5),
            [
                ("a", "Acetyl Hexapeptide-8 là một hợp chất amino acid.", False),
                ("b", "Acetyl Hexapeptide-8 có chứa các nguyên tố C, H, N, O.", False),
                ("c", "Acetyl Hexapeptide-8 là một peptide.", True),
                ("d", "Acetyl Hexapeptide-8 là một hợp chất tạp chức.", True)
            ],
            (
                "a) Sai. Acetyl Hexapeptide-8 là một hexapeptide biến tính (oligopeptide 6 gốc), không phải amino acid đơn phân.\n"
                "b) Sai. Nhìn vào công thức cấu tạo, phân tử có chứa gốc methionine (-CH~2~-CH~2~-S-CH~3~), do đó có chứa thêm nguyên tố lưu huỳnh (S), không chỉ gồm C, H, N, O.\n"
                "c) Đúng. Đây là một hexapeptide có 5 liên kết peptide trong mạch.\n"
                "d) Đúng. Phân tử chứa đồng thời nhiều loại nhóm chức: amide (-CONH-), carboxyl (-COOH), thioether (-S-), hydroxyl (-OH)... nên là hợp chất tạp chức."
            )
        ),
        (
            4,
            "Casein là loại protein chủ yếu có trong sữa. Casein là một dạng protein (đạm) có trong sữa mẹ, sữa bò và sữa dê. Vì là protein hoàn chỉnh nên casein bổ sung nhiều amino acid thiết yếu và các hợp chất khác giúp cơ thể tăng trưởng, phát triển.",
            ("p2_cau4_casein.png", 11.5),
            [
                ("a", "Casein là loại protein tan tốt trong nước.", False),
                ("b", "Uống sữa giúp giảm bớt nguy hiểm khi bị ngộ độc bởi muối chì, muối thuỷ ngân.", True),
                ("c", "Vì casein là một loại protein sẽ đông tụ và tách khỏi dung dịch khi thêm dung dịch acid, base, muối của các kim loại nặng như chì, thủy ngân,... nên uống sữa khi bị ngộ độc muối chì, muối thủy ngân sẽ tạo thành kết tủa và thải được ra ngoài.", True),
                ("d", "Casein là loại protein có duy nhất trong sữa.", False)
            ],
            (
                "a) Sai. Casein là phosphoprotein hầu như không tan trong nước tinh khiết ở điểm đẳng điện (pH ~ 4,6); trong sữa, casein tồn tại ở dạng hạt keo micelle phân tán.\n"
                "b) Đúng. Sữa chứa hàm lượng protein cao (đặc biệt là casein) có khả năng kết tủa ion kim loại nặng.\n"
                "c) Đúng. Protein casein đông tụ liên kết với các ion kim loại nặng (Pb^2+^, Hg^2+^) tạo kết tủa bền không tan, ngăn ngừa chất độc ngấm qua niêm mạc ruột vào máu.\n"
                "d) Sai. Trong sữa ngoài casein (khoảng 80%) còn có whey protein (chiếm khoảng 20%, gồm lactoglobulin, lactalbumin...), nên casein không phải protein duy nhất."
            )
        )
    ]
    
    for item in part2_data:
        num, text, img_info, statements, sol = item
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        
        r_num = p.add_run(f"Câu {num}. ")
        set_run_style(r_num, font_size=11, bold=True)
        add_chem_run(p, text, base_size=11)
        
        if img_info:
            img_file, img_w = img_info
            img_path = os.path.join(IMG_DIR, img_file)
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(3)
                p_img.add_run().add_picture(img_path, width=Cm(img_w))
                
        if num == 2:
            p_lead = doc.add_paragraph()
            p_lead.paragraph_format.space_before = Pt(1)
            p_lead.paragraph_format.space_after = Pt(2)
            add_chem_run(p_lead, "Về phân loại hợp chất và đặc điểm cấu tạo, xét tính đúng/sai của các phát biểu sau:", base_size=11, italic=True)
            
        for stmt_label, stmt_text, is_correct in statements:
            p_stmt = doc.add_paragraph()
            p_stmt.paragraph_format.space_before = Pt(1)
            p_stmt.paragraph_format.space_after = Pt(1.5)
            p_stmt.paragraph_format.line_spacing = 1.15
            p_stmt.paragraph_format.left_indent = Cm(0.6)
            
            r_slbl = p_stmt.add_run(f"{stmt_label}) ")
            set_run_style(r_slbl, font_size=11, bold=True)
            add_chem_run(p_stmt, stmt_text, base_size=11)
            
            if is_teacher:
                status_txt = " — (Đúng)" if is_correct else " — (Sai)"
                r_status = p_stmt.add_run(status_txt)
                set_run_style(r_status, font_size=11, bold=True, color=COLOR_CORRECT)
                
        if is_teacher and sol:
            add_solution(doc, sol)

    # ==================== PHẦN III ====================
    add_part_heading(
        doc,
        "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 6."
    )
    
    part3_data = [
        (
            1,
            "Cho ethylamine lần lượt vào dung dịch FeCl~3~; dung dịch HCl; Cu(OH)~2~; dung dịch NaCl; dung dịch Br~2~. Ethylamine phản ứng được với bao nhiêu chất?",
            "3",
            "Ethylamine (CH~3~CH~2~NH~2~) tác dụng được với 3 chất:\n"
            "- Dung dịch FeCl~3~: 3CH~3~CH~2~NH~2~ + FeCl~3~ + 3H~2~O -> Fe(OH)~3~↓ (nâu đỏ) + 3CH~3~CH~2~NH~3~Cl.\n"
            "- Dung dịch HCl: CH~3~CH~2~NH~2~ + HCl -> CH~3~CH~2~NH~3~Cl.\n"
            "- Cu(OH)~2~: Hòa tan kết tủa tạo phức xanh lam: Cu(OH)~2~ + 4CH~3~CH~2~NH~2~ -> [Cu(CH~3~CH~2~NH~2~)~4~](OH)~2~.\n"
            "Ethylamine không phản ứng với NaCl và dung dịch Br~2~.\n"
            "**Đáp số: 3.**"
        ),
        (
            2,
            "Cho dãy các dung dịch: acetic acid, phenylammonium chloride, sodium acetate, methylamine, glycine, phenol. Số dung dịch trong dãy tác dụng được với dung dịch NaOH là",
            "4",
            "Các dung dịch tác dụng được với dung dịch NaOH gồm:\n"
            "(1) Acetic acid: CH~3~COOH + NaOH -> CH~3~COONa + H~2~O.\n"
            "(2) Phenylammonium chloride: C~6~H~5~NH~3~Cl + NaOH -> C~6~H~5~NH~2~ + NaCl + H~2~O.\n"
            "(3) Glycine: H~2~N-CH~2~-COOH + NaOH -> H~2~N-CH~2~-COONa + H~2~O.\n"
            "(4) Phenol: C~6~H~5~OH + NaOH -> C~6~H~5~ONa + H~2~O.\n"
            "Sodium acetate và methylamine không phản ứng với NaOH.\n"
            "=> Có 4 dung dịch tác dụng được với NaOH.\n"
            "**Đáp số: 4.**"
        ),
        (
            3,
            "Cho các chất: phenylammonium chloride, alanine, lysine, glycine, ethylamine. Số chất làm quỳ tím đổi màu là",
            "3",
            "Xét sự đổi màu quỳ tím của các dung dịch:\n"
            "- Phenylammonium chloride (C~6~H~5~NH~3~Cl): Muối của base rất yếu và acid mạnh -> dung dịch có tính acid làm quỳ tím hóa đỏ (hồng).\n"
            "- Alanine (1 nhóm NH~2~, 1 nhóm COOH): Quỳ tím không đổi màu.\n"
            "- Lysine (2 nhóm NH~2~, 1 nhóm COOH): Tính base trội hơn -> quỳ tím hóa xanh.\n"
            "- Glycine (1 nhóm NH~2~, 1 nhóm COOH): Quỳ tím không đổi màu.\n"
            "- Ethylamine: Tính base mạnh -> quỳ tím hóa xanh.\n"
            "=> Có 3 chất làm đổi màu quỳ tím.\n"
            "**Đáp số: 3.**"
        ),
        (
            4,
            "Cho các phát biểu sau:\n"
            "(1) Dung dịch ethylamine và dung dịch aniline đều làm xanh giấy quỳ tím.\n"
            "(2) Nhỏ từ từ đến dư dung dịch methylamine vào dung dịch copper(II) sulfate, ban đầu thấy xuất hiện kết tủa màu xanh nhạt, sau đó kết tủa tan tạo thành dung dịch màu xanh lam.\n"
            "(3) Cho dung dịch methylamine vào ống nghiệm đựng dung dịch iron(III) chloride thấy xuất hiện kết tủa nâu đỏ.\n"
            "(4) Nhỏ vài giọt dung dịch aniline vào ống nghiệm đựng nước bromine thấy xuất hiện kết tủa trắng.\n"
            "Số phát biểu đúng là",
            "3",
            "(1) Sai, aniline có tính base rất yếu không làm đổi màu quỳ tím (chỉ ethylamine làm xanh quỳ tím).\n"
            "(2) Đúng, ban đầu tạo kết tủa Cu(OH)~2~, sau đó tạo phức chất hòa tan [Cu(CH~3~NH~2~)~4~]^2+^ xanh lam thẫm.\n"
            "(3) Đúng, methylamine tạo kết tủa nâu đỏ Fe(OH)~3~: 3CH~3~NH~2~ + FeCl~3~ + 3H~2~O -> Fe(OH)~3~↓ + 3CH~3~NH~3~Cl.\n"
            "(4) Đúng, aniline phản ứng thế với bromine tạo kết tủa trắng 2,4,6-tribromoaniline.\n"
            "=> Có 3 phát biểu đúng: (2), (3), (4).\n"
            "**Đáp số: 3.**"
        ),
        (
            5,
            "Thuốc Paracetamol là một loại thuốc được sử dụng rất nhiều trong y học. Thuốc có tác dụng giảm đau, hạ sốt tuy nhiên ít có tác dụng kháng viêm. Thuốc Paracetamol được tổng hợp từ p-nitrophenol theo sơ đồ phản ứng sau (biết hiệu suất của cả quá trình phản ứng đạt 90%):",
            "15,44",
            "Khối lượng mol: M(Paracetamol) = 151 g/mol; M(p-nitrophenol) = 139 g/mol.\n"
            "Khối lượng Paracetamol cần sản xuất:\n"
            "m(Paracetamol) = 30,2 * 10^6 * 500 mg = 30,2 * 10^6 * 0,5 g = 15,1 * 10^6 g = 15,1 tấn.\n"
            "Số mol Paracetamol = 15,1 * 10^6 / 151 = 100 000 mol.\n"
            "Do hiệu suất toàn bộ quá trình đạt 90%:\n"
            "n(p-nitrophenol cần dùng) = 100 000 / 0,9 ≈ 111 111,1 mol.\n"
            "Khối lượng p-nitrophenol cần dùng là:\n"
            "m = (100 000 / 0,9) * 139 ≈ 15 444 444 g ≈ 15,44 tấn (hoặc 15,4 tấn).\n"
            "**Đáp số: 15,44.**"
        ),
        (
            6,
            "Một người trưởng thành cần tiêu thụ khoảng 0,8 gram protein cho mỗi kg trọng lượng cơ thể mỗi ngày. Một người nặng 70 kg cần bổ sung bao nhiêu gram protein trong một ngày? Nếu mỗi quả trứng chứa khoảng 6 gram protein, người đó cần ăn bao nhiêu quả trứng để đáp ứng lượng protein cần thiết trong ngày? Biết số quả trứng là một con số chẵn.",
            "10",
            "Lượng protein cần thiết cho người nặng 70 kg trong 1 ngày:\n"
            "m(protein) = 70 * 0,8 = 56 gam.\n"
            "Số quả trứng cần ăn tương ứng:\n"
            "Số trứng = 56 / 6 ≈ 9,33 quả.\n"
            "Vì cần đáp ứng đủ lượng protein và số quả trứng là số chẵn, nên người đó cần ăn 10 quả trứng.\n"
            "**Đáp số: 10.**"
        )
    ]
    
    for q in part3_data:
        num, text, ans, sol = q
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        
        r_num = p.add_run(f"Câu {num}. ")
        set_run_style(r_num, font_size=11, bold=True)
        add_chem_run(p, text, base_size=11)
        
        if num == 5:
            img_path = os.path.join(IMG_DIR, "p3_cau5_paracetamol.png")
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_path, width=Cm(14.5))
            
            p_sub5 = doc.add_paragraph()
            p_sub5.paragraph_format.space_before = Pt(1)
            p_sub5.paragraph_format.space_after = Pt(2)
            add_chem_run(p_sub5, "Để sản xuất 30,2 triệu viên thuốc Paracetamol thì khối lượng p-nitrophenol cần dùng gần nhất với giá trị nào sau đây (theo đơn vị tấn)? Biết rằng mỗi viên thuốc chứa 500 mg Paracetamol.", base_size=11)
            
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(3)
        p_ans.paragraph_format.left_indent = Cm(0.6)
        
        r_ans_lbl = p_ans.add_run("Đáp số: ")
        set_run_style(r_ans_lbl, font_size=11, italic=True)
        
        if is_teacher:
            r_ans_val = p_ans.add_run(ans)
            set_run_style(r_ans_val, font_size=11, bold=True, color=COLOR_CORRECT)
        else:
            r_dots = p_ans.add_run(". . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .")
            set_run_style(r_dots, font_size=11, color=COLOR_GRAY)
            
        if is_teacher and sol:
            add_solution(doc, sol)

    # Summary table for teacher version
    if is_teacher:
        p_tbl_hdr = doc.add_paragraph()
        p_tbl_hdr.paragraph_format.space_before = Pt(12)
        p_tbl_hdr.paragraph_format.space_after = Pt(4)
        r_thdr = p_tbl_hdr.add_run("BẢNG TỔNG HỢP ĐÁP ÁN ĐỀ SỐ 08")
        set_run_style(r_thdr, font_size=11.5, bold=True, color=COLOR_PRIMARY)
        
        p_p1 = doc.add_paragraph()
        p_p1.paragraph_format.space_before = Pt(2)
        p_p1.paragraph_format.space_after = Pt(2)
        r_p1 = p_p1.add_run("Phần I: ")
        set_run_style(r_p1, font_size=10.5, bold=True)
        p1_ans_str = "1.D  2.C  3.A  4.B  5.A  6.D  7.B  8.C  9.C  10.A  11.A  12.B  13.A  14.C  15.D  16.C  17.C  18.A"
        r_p1_val = p_p1.add_run(p1_ans_str)
        set_run_style(r_p1_val, font_size=10.5, bold=True, color=COLOR_CORRECT)
        
        p_p2 = doc.add_paragraph()
        p_p2.paragraph_format.space_before = Pt(2)
        p_p2.paragraph_format.space_after = Pt(2)
        r_p2 = p_p2.add_run("Phần II: ")
        set_run_style(r_p2, font_size=10.5, bold=True)
        p2_ans_str = "Câu 1: a-S, b-Đ, c-Đ, d-Đ  |  Câu 2: a-Đ, b-S, c-S, d-Đ  |  Câu 3: a-S, b-S, c-Đ, d-Đ  |  Câu 4: a-S, b-Đ, c-Đ, d-S"
        r_p2_val = p_p2.add_run(p2_ans_str)
        set_run_style(r_p2_val, font_size=10.5, bold=True, color=COLOR_CORRECT)
        
        p_p3 = doc.add_paragraph()
        p_p3.paragraph_format.space_before = Pt(2)
        p_p3.paragraph_format.space_after = Pt(4)
        r_p3 = p_p3.add_run("Phần III: ")
        set_run_style(r_p3, font_size=10.5, bold=True)
        p3_ans_str = "Câu 1: 3  |  Câu 2: 4  |  Câu 3: 3  |  Câu 4: 3  |  Câu 5: 15,44  |  Câu 6: 10"
        r_p3_val = p_p3.add_run(p3_ans_str)
        set_run_style(r_p3_val, font_size=10.5, bold=True, color=COLOR_CORRECT)

    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(10)
    p_end.paragraph_format.space_after = Pt(6)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(r_end, font_size=11, bold=True)

    out_name = "De_Kiem_Tra_Hoa_12_De08_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_Hoa_12_De08.docx"
    out_path = os.path.join(OUT_DIR, out_name)
    doc.save(out_path)
    print(f"Saved: {out_path}")

if __name__ == "__main__":
    build_docx_de08(is_teacher=False)
    build_docx_de08(is_teacher=True)
