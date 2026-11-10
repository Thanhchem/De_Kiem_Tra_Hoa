# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Kì I Hóa Học 12 - THPT NGUYỄN HUỆ - Năm học 2026 - 2027
Được xây dựng chuẩn 100% theo MA TRẬN yêu cầu và FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH)

Xuất bản 2 file Word:
1. Bản Học sinh: San_Pham/De_Kiem_Tra_GK1_Hoa_12_THPT_NguyenHue_Ma101_HocSinh.docx
2. Bản Giáo viên: San_Pham/De_Kiem_Tra_GK1_Hoa_12_THPT_NguyenHue_Ma101_GiaoVien.docx (Kèm đáp án & lời giải chi tiết)
"""

import os
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

sys.path.append(os.path.abspath(r"C:\Antigravity_Thanh\scripts"))
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
    COLOR_PRIMARY,
    COLOR_CORRECT,
    COLOR_SOLUTION,
    COLOR_GRAY
)

OUTPUT_DIR = r"C:\Antigravity_Thanh\San_Pham"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------------------------------------
# DỮ LIỆU CÂU HỎI THEO ĐÚNG MA TRẬN THPT NGUYỄN HUỆ
# -------------------------------------------------------------

# PHẦN I: 12 câu TN nhiều lựa chọn (3,0 điểm)
PART1_QUESTIONS = [
    {
        "num": 1,
        "title": "Chất nào sau đây là ester?",
        "options": ["A. CH~3~COOH.", "*B. CH~3~COOC~2~H~5~.*", "C. C~2~H~5~OH.", "D. CH~3~CHO."],
        "level": "NB",
        "topic": "Ester - lipid (2 tiết)",
        "sol": "CH~3~COOC~2~H~5~ (ethyl acetate) là hợp chất thuộc loại ester. CH~3~COOH là carboxylic acid, C~2~H~5~OH là alcohol, CH~3~CHO là aldehyde.\nChọn B."
    },
    {
        "num": 2,
        "title": "Thủy phân hoàn toàn ethyl acetate trong dung dịch NaOH đun nóng, thu được sản phẩm gồm",
        "options": ["A. CH~3~COONa và CH~3~OH.", "*B. CH~3~COONa và C~2~H~5~OH.*", "C. C~2~H~5~COONa và CH~3~OH.", "D. HCOONa và C~2~H~5~OH."],
        "level": "TH",
        "topic": "Ester - lipid (2 tiết)",
        "sol": "Phương trình phản ứng xà phòng hóa:\nCH~3~COOC~2~H~5~ + NaOH -(t°)-> CH~3~COONa + C~2~H~5~OH.\nSản phẩm thu được gồm sodium acetate (CH~3~COONa) và ethanol (C~2~H~5~OH).\nChọn B."
    },
    {
        "num": 3,
        "title": "Phát biểu nào sau đây *không đúng* về xà phòng và chất giặt rửa tổng hợp?",
        "options": [
            "A. Xà phòng thường là muối sodium hoặc potassium của các acid béo.",
            "B. Khi giặt rửa trong nước cứng, xà phòng bị giảm tác dụng tẩy rửa do tạo muối kết tủa.",
            "*C. Phân tử chất giặt rửa tổng hợp không có đầu ưa nước và đuôi kị nước như xà phòng.*",
            "D. Chất giặt rửa tổng hợp có ưu điểm là có thể sử dụng được trong nước cứng và môi trường acid nhẹ."
        ],
        "level": "TH",
        "topic": "Xà phòng và chất giặt rửa (1 tiết)",
        "sol": "Phát biểu C sai vì phân tử chất giặt rửa tổng hợp cũng có cấu tạo gồm 2 phần tương tự xà phòng: phần đầu ưa nước (thường là nhóm -SO~3~Na hoặc -OSO~3~Na) và phần đuôi kị nước (gốc hydrocarbon dài), giúp bám vào vết bẩn và phân tán vào nước tạo nhũ tương.\nChọn C."
    },
    {
        "num": 4,
        "title": "Carbohydrate nào sau đây thuộc loại monosaccharide?",
        "options": ["*A. Glucose.*", "B. Saccharose.", "C. Maltose.", "D. Tinh bột."],
        "level": "NB",
        "topic": "Carbohydrate (5 tiết)",
        "sol": "Glucose (C~6~H~12~O~6~) là monosaccharide đơn giản nhất. Saccharose và maltose là disaccharide. Tinh bột là polysaccharide.\nChọn A."
    },
    {
        "num": 5,
        "title": "Trong điều kiện thường, chất nào sau đây là chất rắn màu trắng, dạng sợi, không mùi, không tan trong nước và là thành phần chính tạo nên màng tế bào thực vật?",
        "options": ["A. Glucose.", "B. Saccharose.", "*C. Cellulose.*", "D. Amylose."],
        "level": "NB",
        "topic": "Carbohydrate (5 tiết)",
        "sol": "Cellulose là polymer thiên nhiên mạch không phân nhánh, dạng sợi màu trắng, không tan trong nước, tạo nên khung màng tế bào thực vật.\nChọn C."
    },
    {
        "num": 6,
        "title": "Thuốc thử dùng để nhận biết hồ tinh bột bằng phản ứng tạo dung dịch màu xanh tím đặc trưng là",
        "options": ["*A. dung dịch iodine (I~2~).*", "B. thuốc thử Tollens.", "C. dung dịch Cu(OH)~2~.", "D. nước bromine."],
        "level": "NB",
        "topic": "Carbohydrate (5 tiết)",
        "sol": "Phân tử tinh bột có cấu trúc xoắn tạo lỗ rỗng, hấp phụ các phân tử iodine (I~2~) tạo thành hợp chất bọc có màu xanh tím đặc trưng.\nChọn A."
    },
    {
        "num": 7,
        "title": "Trong các chất sau, chất nào có lực base *yếu nhất*?",
        "options": ["A. Methylamine.", "B. Dimethylamine.", "C. Ammonia.", "*D. Aniline.*"],
        "level": "TH",
        "topic": "Amine (2 tiết)",
        "sol": "Lực base: (CH~3~)~2~NH (amine bậc 2) > CH~3~NH~2~ (amine bậc 1) > NH~3~ > C~6~H~5~NH~2~ (aniline). Gốc phenyl (C~6~H~5~-) hút electron làm giảm mật độ electron trên nguyên tử nitrogen khiến lực base của aniline rất yếu (yếu hơn cả NH~3~ và không đổi màu quỳ tím).\nChọn D."
    },
    {
        "num": 8,
        "title": "Amino acid đơn giản nhất có tên gọi là glycine. Công thức phân tử của glycine là",
        "options": ["*A. C~2~H~5~NO~2~.*", "B. C~3~H~7~NO~2~.", "C. C~4~H~9~NO~2~.", "D. C~5~H~9~NO~4~."],
        "level": "NB",
        "topic": "Amino acid và peptide (2 tiết)",
        "sol": "Glycine có công thức cấu tạo là H~2~N-CH~2~-COOH, công thức phân tử là C~2~H~5~NO~2~ (M = 75 g/mol).\nChọn A."
    },
    {
        "num": 9,
        "title": "Số liên kết peptide trong một phân tử tripeptide Gly-Ala-Val là",
        "options": ["A. 1.", "*B. 2.*", "C. 3.", "D. 4."],
        "level": "NB",
        "topic": "Amino acid và peptide (2 tiết)",
        "sol": "Phân tử tripeptide gồm 3 gốc alpha-amino acid liên kết với nhau bằng 2 liên kết peptide (-CO-NH-).\nChọn B."
    },
    {
        "num": 10,
        "title": "Hiện tượng protein bị đông tụ và tách ra khỏi dung dịch xảy ra khi",
        "options": [
            "*A. đun nóng dung dịch lòng trắng trứng.*",
            "B. hòa tan lòng trắng trứng vào nước lạnh.",
            "C. thêm nước cất vào dung dịch lòng trắng trứng.",
            "D. khuấy đều dung dịch lòng trắng trứng ở nhiệt độ thường."
        ],
        "level": "NB",
        "topic": "Protein và enzyme (1 tiết)",
        "sol": "Khi đun nóng dung dịch protein (như lòng trắng trứng) hoặc cho thêm acid/kiềm/muối, cấu trúc không gian của protein bị phá hủy dẫn đến hiện tượng đông tụ.\nChọn A."
    },
    {
        "num": 11,
        "title": "Thuốc thử dùng để phân biệt tripeptide với dipeptide bằng phản ứng màu biuret là",
        "options": [
            "A. dung dịch AgNO~3~/NH~3~.",
            "*B. Cu(OH)~2~ trong môi trường kiềm.*",
            "C. nước bromine.",
            "D. dung dịch I~2~ trong KI."
        ],
        "level": "NB",
        "topic": "Protein và enzyme (1 tiết)",
        "sol": "Trong môi trường kiềm, các hợp chất peptide có từ 2 liên kết peptide trở lên (tripeptide trở lên) và protein phản ứng với Cu(OH)~2~ tạo phức chất màu tím đặc trưng (phản ứng biuret). Dipeptide không có phản ứng này.\nChọn B."
    },
    {
        "num": 12,
        "title": "Phát biểu nào sau đây về enzyme là *không đúng*?",
        "options": [
            "A. Hầu hết các enzyme có bản chất là protein.",
            "B. Enzyme là chất xúc tác sinh học có tính đặc hiệu rất cao.",
            "*C. Enzyme xúc tác cho các phản ứng sinh hóa ở nhiệt độ rất cao (trên 150 °C) để tăng tốc độ phản ứng.*",
            "D. Tốc độ của phản ứng do enzyme xúc tác thường lớn hơn rất nhiều so với chất xúc tác hóa học thông thường."
        ],
        "level": "TH",
        "topic": "Protein và enzyme (1 tiết)",
        "sol": "Phát biểu C không đúng vì enzyme hoạt động tối ưu ở điều kiện sinh lý êm dịu (khoảng 37 °C, pH thích hợp). Ở nhiệt độ cao (trên 60 - 70 °C), protein enzyme bị biến tính và mất hoạt tính xúc tác hoàn toàn.\nChọn C."
    }
]

# PHẦN II: 2 câu Đúng - Sai (2,0 điểm)
PART2_QUESTIONS = [
    {
        "num": 1,
        "stem": "Glucose và cellulose là hai carbohydrate phổ biến trong tự nhiên, đóng vai trò quan trọng trong sự sống và sản xuất công nghiệp. Cho các phát biểu sau:",
        "topic": "Carbohydrate (5 tiết) [2 Biết, 1 Hiểu, 1 VD]",
        "items": [
            ("a", "Glucose có công thức phân tử là C~6~H~12~O~6~, ở dạng mạch hở phân tử có chứa 5 nhóm hydroxy (-OH) và 1 nhóm aldehyde (-CHO).", "Đ", "NB"),
            ("b", "Phân tử cellulose được cấu tạo từ các đơn vị alpha-glucose liên kết với nhau qua liên kết alpha-1,4-glycoside tạo thành mạch phân nhánh.", "S", "NB"),
            ("c", "Thủy phân hoàn toàn cellulose trong môi trường acid vô cơ đun nóng thu được sản phẩm duy nhất là glucose, dung dịch sản phẩm có phản ứng tráng bạc với thuốc thử Tollens.", "Đ", "TH"),
            ("d", "Đun nóng 32,4 kg cellulose với lượng dư hỗn hợp HNO~3~ đặc trong H~2~SO~4~ đặc, thu được cellulose trinitrate [C~6~H~7~O~2~(ONO~2~)~3~]~n~ dùng làm thuốc súng không khói. Biết hiệu suất phản ứng đạt 80%, khối lượng cellulose trinitrate thu được là 47,52 kg.", "Đ", "VD")
        ],
        "sol": (
            "a) Đúng. Glucose là monosaccharide có CTPT C~6~H~12~O~6~, dạng mạch hở là HOCH~2~-[CHOH]~4~-CHO chứa 5 nhóm -OH kề nhau và 1 nhóm -CHO.\n"
            "b) Sai. Phân tử cellulose cấu tạo từ các mắt xích beta-glucose qua liên kết beta-1,4-glycoside tạo thành mạch kéo dài không phân nhánh.\n"
            "c) Đúng. (C~6~H~10~O~5~)~n~ + nH~2~O -(H+, t°)-> nC~6~H~12~O~6~ (glucose). Glucose tráng bạc tạo 2Ag.\n"
            "d) Đúng. Sơ đồ: [C~6~H~7~O~2~(OH)~3~]~n~ + 3nHNO~3~ -(H~2~SO~4~ đặc, t°)-> [C~6~H~7~O~2~(ONO~2~)~3~]~n~ + 3nH~2~O.\n"
            "Mỗi mắt xích cellulose (162 g) tạo ra 1 mắt xích cellulose trinitrate (297 g).\n"
            "Khối lượng cellulose trinitrate lý thuyết: m~LT~ = (32,4 * 297) / 162 = 59,4 kg.\n"
            "Vì hiệu suất H = 80% => Khối lượng thực tế thu được: m~TT~ = 59,4 * 80% = 47,52 kg."
        )
    },
    {
        "num": 2,
        "stem": "Cho các hợp chất hữu cơ chứa nitrogen gồm: methylamine (CH~3~NH~2~), aniline (C~6~H~5~NH~2~), glycine (H~2~N-CH~2~-COOH), glutamic acid và tripeptide Gly-Ala-Val. Cho các phát biểu sau:",
        "topic": "Hợp chất chứa nitrogen (5 tiết) [1 Biết Amine; 1 Biết, 1 Hiểu, 1 VD Amino acid & Peptide]",
        "items": [
            ("a", "Ở điều kiện thường, methylamine là chất khí có mùi khai khó chịu, độc, tan nhiều trong nước và làm quỳ tím ẩm hóa xanh.", "Đ", "NB (Amine)"),
            ("b", "Glycine là amino acid đơn giản nhất, ở trạng thái tinh thể tồn tại chủ yếu dưới dạng ion lưỡng cực (^+^H~3~N-CH~2~-COO^-^).", "Đ", "NB (Amino acid)"),
            ("c", "Dung dịch glutamic acid làm đổi màu quỳ tím sang màu xanh tương tự như dung dịch lysine.", "S", "TH (Amino acid)"),
            ("d", "Thủy phân hoàn toàn 0,1 mol peptide Gly-Ala-Val trong dung dịch chứa 0,4 mol NaOH đun nóng. Sau khi phản ứng xảy ra hoàn toàn, cô cạn dung dịch thu được 38,7 gam chất rắn khan.", "Đ", "VD (Peptide)")
        ],
        "sol": (
            "a) Đúng. Methylamine là chất khí ở điều kiện thường, mùi khai tương tự NH~3~, tan rất tốt trong nước và có tính base làm quỳ tím ẩm hóa xanh.\n"
            "b) Đúng. Các amino acid ở điều kiện thường là tinh thể không màu, nhiệt độ nóng chảy cao và tan tốt trong nước do tồn tại chủ yếu dưới dạng ion lưỡng cực.\n"
            "c) Sai. Glutamic acid có 2 nhóm -COOH và 1 nhóm -NH~2~ nên dung dịch có môi trường acid, làm quỳ tím chuyển sang màu đỏ (hoặc hồng). Lysine có 2 nhóm -NH~2~ và 1 nhóm -COOH nên dung dịch có môi trường base làm quỳ tím chuyển sang màu xanh.\n"
            "d) Đúng. Tripeptide Gly-Ala-Val có phân tử khối: M = 75 (Gly) + 89 (Ala) + 117 (Val) - 2 * 18 (H~2~O) = 245 g/mol.\n"
            "Phương trình phản ứng: Gly-Ala-Val + 3NaOH -> Muối + H~2~O.\n"
            "n~NaOH phản ứng~ = 3 * n~peptide~ = 3 * 0,1 = 0,3 mol => n~NaOH dư~ = 0,4 - 0,3 = 0,1 mol.\n"
            "n~H2O sinh ra~ = n~peptide~ = 0,1 mol.\n"
            "Bảo toàn khối lượng: m~muối~ = m~peptide~ + m~NaOH phản ứng~ - m~H2O~ = 0,1 * 245 + 0,3 * 40 - 0,1 * 18 = 24,5 + 12 - 1,8 = 34,7 gam.\n"
            "Chất rắn khan thu được gồm muối (34,7 gam) và NaOH dư (0,1 * 40 = 4,0 gam).\n"
            "=> m~rắn~ = 34,7 + 4,0 = 38,7 gam."
        )
    }
]

# PHẦN III: 4 câu Trả lời ngắn (2,0 điểm)
PART3_QUESTIONS = [
    {
        "num": 1,
        "stem": "Cho các chất sau: sodium stearate, sodium dodecylbenzenesulfonate, saccharose, glycerol, potassium palmitate, ethyl acetate. Có bao nhiêu chất là thành phần chính của xà phòng hoặc chất giặt rửa tổng hợp?",
        "ans": "3",
        "level": "NB",
        "topic": "Xà phòng và chất giặt rửa (1 tiết)",
        "sol": "Có 3 chất là thành phần chính của xà phòng hoặc chất giặt rửa tổng hợp:\n- Xà phòng: sodium stearate (C~17~H~35~COONa), potassium palmitate (C~15~H~31~COOK).\n- Chất giặt rửa tổng hợp: sodium dodecylbenzenesulfonate (CH~3~[CH~2~]~11~C~6~H~4~SO~3~Na).\nĐáp án: 3."
    },
    {
        "num": 2,
        "stem": "Trong các amine sau: methylamine, dimethylamine, trimethylamine, ethylamine, aniline; có bao nhiêu amine ở thể khí ở điều kiện thường?",
        "ans": "4",
        "level": "NB",
        "topic": "Amine (2 tiết)",
        "sol": "Có 4 amine ở thể khí ở điều kiện thường là: methylamine (CH~3~NH~2~), dimethylamine ((CH~3~)~2~NH), trimethylamine ((CH~3~)~3~N) và ethylamine (C~2~H~5~NH~2~). Aniline là chất lỏng không màu, ít tan trong nước.\nĐáp án: 4."
    },
    {
        "num": 3,
        "stem": "Nhỏ từng giọt dung dịch của mỗi chất sau vào các mẩu giấy quỳ tím riêng biệt: glycine, ethylamine, aniline, lysine, glutamic acid. Có bao nhiêu chất làm quỳ tím chuyển sang màu xanh?",
        "ans": "2",
        "level": "TH",
        "topic": "Amino acid và peptide (2 tiết)",
        "sol": "Có 2 chất làm quỳ tím chuyển sang màu xanh là:\n- Ethylamine (C~2~H~5~NH~2~): amine béo có tính base mạnh hơn NH~3~.\n- Lysine: có 2 nhóm -NH~2~ và 1 nhóm -COOH (số nhóm -NH~2~ > số nhóm -COOH) nên có tính base.\n(Glycine có 1 -NH~2~ và 1 -COOH không làm đổi màu quỳ tím; Aniline có tính base rất yếu không đổi màu quỳ tím; Glutamic acid làm quỳ tím hóa đỏ).\nĐáp án: 2."
    },
    {
        "num": 4,
        "stem": "Cho 9,0 gam ethylamine (C~2~H~5~NH~2~) tác dụng vừa đủ với dung dịch hydrochloric acid (HCl). Sau khi phản ứng hoàn toàn, khối lượng muối ethylammonium chloride (C~2~H~5~NH~3~Cl) thu được là bao nhiêu gam?",
        "ans": "16,3",
        "level": "VD",
        "topic": "Amine (2 tiết)",
        "sol": "Phương trình phản ứng: C~2~H~5~NH~2~ + HCl -> C~2~H~5~NH~3~Cl.\nSố mol C~2~H~5~NH~2~: n = 9,0 / 45 = 0,2 mol.\nTheo phương trình: n~muối~ = n~amine~ = 0,2 mol.\nPhân tử khối muối C~2~H~5~NH~3~Cl: M = 45 + 36,5 = 81,5 g/mol.\nKhối lượng muối thu được: m = 0,2 * 81,5 = 16,3 gam.\nĐáp án: 16,3."
    }
]

# PHẦN IV: 3 câu Tự luận (3,0 điểm)
PART4_QUESTIONS = [
    {
        "num": 1,
        "points": "1,0 điểm",
        "level": "TH",
        "topic": "Carbohydrate (5 tiết)",
        "stem": (
            "Có hai ống nghiệm chứa dung dịch saccharose không màu. Tiến hành hai thí nghiệm sau:\n"
            "- *Ống nghiệm 1:* Thêm lượng dư thuốc thử Tollens [Ag(NH~3~)~2~]OH, sau đó ngâm ống nghiệm trong cốc nước nóng (khoảng 70 - 80 °C) trong 5 phút.\n"
            "- *Ống nghiệm 2:* Thêm 1 mL dung dịch H~2~SO~4~ 10% rồi đun nóng nhẹ trong 5 phút. Để nguội, trung hòa acid dư bằng dung dịch NaOH loãng, sau đó thêm lượng dư thuốc thử Tollens và ngâm trong cốc nước nóng.\n"
            "a) Hãy nêu hiện tượng quan sát được ở mỗi ống nghiệm.\n"
            "b) Viết phương trình hóa học của phản ứng thủy phân xảy ra ở ống nghiệm 2 và phản ứng tráng bạc giải thích hiện tượng quan sát được."
        ),
        "sol": (
            "a) *Hiện tượng quan sát được:*\n"
            "- Ống 1: Không có hiện tượng tráng bạc, không xuất hiện lớp kim loại sáng bóng trên thành ống nghiệm (do phân tử saccharose gồm 1 gốc alpha-glucose và 1 gốc beta-fructose liên kết qua -O- giữa C1 và C2, không còn nhóm hemiacetal/hemiketal tự do nên không thể mở vòng tạo nhóm aldehyde -CHO, không có tính khử).\n"
            "- Ống 2: Thành ống nghiệm xuất hiện lớp kim loại bạc sáng bóng bám vào như gương (hoặc kết tủa bạc xám đen tách ra).\n\n"
            "b) *Phương trình hóa học:*\n"
            "- Phản ứng thủy phân saccharose trong môi trường acid:\n"
            "  C~12~H~22~O~11~ + H~2~O -(H^+, t°)-> C~6~H~12~O~6~ (glucose) + C~6~H~12~O~6~ (fructose)\n"
            "- Phản ứng tráng bạc của glucose sinh ra sau thủy phân:\n"
            "  CH~2~OH[CHOH]~4~CHO + 2[Ag(NH~3~)~2~]OH -(t°)-> CH~2~OH[CHOH]~4~COONH~4~ + 2Ag + 3NH~3~ + H~2~O\n"
            "*(Lưu ý: Trong môi trường kiềm của thuốc thử Tollens, fructose cũng chuyển hóa một phần thành glucose nên cả hai monosaccharide đều tham gia phản ứng tráng bạc).*"
        )
    },
    {
        "num": 2,
        "points": "1,0 điểm",
        "level": "VD",
        "topic": "Ester - lipid (2 tiết)",
        "stem": (
            "Một cơ sở sản xuất xà phòng thủ công tiến hành xà phòng hóa hoàn toàn 17,8 kg một loại chất béo chứa 90% khối lượng là tristearin (C~17~H~35~COO)~3~C~3~H~5~ (còn lại là tạp chất trơ không phản ứng) bằng một lượng vừa đủ dung dịch NaOH, đun nóng.\n"
            "a) Viết phương trình hóa học của phản ứng xà phòng hóa tristearin bằng dung dịch NaOH.\n"
            "b) Tính khối lượng glycerol thu được sau phản ứng (giả sử hiệu suất phản ứng xà phòng hóa đạt 100%).\n"
            "c) Toàn bộ lượng muối sodium stearate (C~17~H~35~COONa) thu được sau khi cô đặc và phối trộn phụ gia (chiếm 25% khối lượng bánh xà phòng thương phẩm) được ép thành các bánh xà phòng. Mỗi bánh xà phòng có khối lượng 100 gam và hiệu suất thu hồi xà phòng đạt 95%. Tính số bánh xà phòng tối đa mà cơ sở trên sản xuất được."
        ),
        "sol": (
            "a) *Phương trình hóa học xà phòng hóa:*\n"
            "  (C~17~H~35~COO)~3~C~3~H~5~ + 3NaOH -(t°)-> 3C~17~H~35~COONa + C~3~H~5~(OH)~3~\n\n"
            "b) *Tính khối lượng glycerol:*\n"
            "- Khối lượng tristearin nguyên chất: m = 17,8 * 90% = 16,02 kg = 16020 gam.\n"
            "- Số mol tristearin: n = 16020 / 890 = 18 mol.\n"
            "- Theo PTHH: n~glycerol~ = n~tristearin~ = 18 mol.\n"
            "- Khối lượng glycerol thu được: m~glycerol~ = 18 * 92 = 1656 gam = 1,656 kg.\n\n"
            "c) *Tính số bánh xà phòng tối đa sản xuất được:*\n"
            "- Theo PTHH: n~C17H35COONa (LT)~ = 3 * n~tristearin~ = 3 * 18 = 54 mol.\n"
            "- Khối lượng sodium stearate lý thuyết: m~muối (LT)~ = 54 * 306 = 16524 gam.\n"
            "- Vì hiệu suất thu hồi đạt 95%: m~muối (TT)~ = 16524 * 95% = 15697,8 gam.\n"
            "- Trong bánh xà phòng thành phẩm, sodium stearate chiếm: 100% - 25% = 75% khối lượng bánh.\n"
            "- Tổng khối lượng xà phòng bánh thu được: m~xà phòng~ = 15697,8 / 75% = 20930,4 gam.\n"
            "- Số bánh xà phòng tối đa: N = 20930,4 / 100 = 209,304 => Làm tròn đạt tối đa 209 bánh xà phòng hoàn chỉnh."
        )
    },
    {
        "num": 3,
        "points": "1,0 điểm",
        "level": "VD",
        "topic": "Carbohydrate (5 tiết)",
        "stem": (
            "Để sản xuất cồn y tế 70° (dung dịch ethanol 70% theo thể tích) phục vụ sát trùng trong y tế, người ta sử dụng nguyên liệu ban đầu là 1 tấn hạt ngô chứa 64,8% tinh bột (còn lại là chất xơ và tạp chất trơ không sinh cồn). Quá trình sản xuất trải qua hai giai đoạn theo sơ đồ sau:\n"
            "  (C~6~H~10~O~5~)~n~ -(+ H~2~O, xúc tác acid hoặc enzyme)-> nC~6~H~12~O~6~ -(lên men rượu)-> 2nC~2~H~5~OH + 2nCO~2~\n"
            "a) Tính khối lượng ethanol lý thuyết thu được từ 1 tấn hạt ngô trên.\n"
            "b) Biết hiệu suất của toàn bộ quá trình sản xuất trên đạt 80% và khối lượng riêng của ethanol nguyên chất là D = 0,8 g/mL. Tính thể tích (lít) dung dịch cồn y tế 70° tối đa pha chế được từ lượng ethanol thực tế thu được."
        ),
        "sol": (
            "a) *Tính khối lượng ethanol lý thuyết:*\n"
            "- Khối lượng tinh bột nguyên chất có trong 1 tấn (1000 kg) hạt ngô:\n"
            "  m~tinh bột~ = 1000 * 64,8% = 648 kg.\n"
            "- Sơ đồ phản ứng tổng hợp từ tinh bột ra ethanol:\n"
            "  (C~6~H~10~O~5~)~n~  --->  2n C~2~H~5~OH\n"
            "  162n kg       --->    2n * 46 = 92n kg\n"
            "  648 kg        --->    m~ethanol (LT)~ = (648 * 92n) / (162n) = 368 kg.\n\n"
            "b) *Tính thể tích cồn y tế 70° tối đa:*\n"
            "- Khối lượng ethanol thực tế thu được (với hiệu suất toàn quá trình H = 80%):\n"
            "  m~ethanol (TT)~ = 368 * 80% = 294,4 kg = 294400 gam.\n"
            "- Thể tích ethanol nguyên chất thu được:\n"
            "  V~ethanol~ = m / D = 294400 / 0,8 = 368000 mL = 368 lít.\n"
            "- Thể tích dung dịch cồn y tế 70° pha chế được:\n"
            "  Độ cồn = (V~ethanol~ / V~dung dịch~) * 100% = 70%\n"
            "  => V~dung dịch cồn 70°~ = 368 / 70% ≈ 525,71 lít."
        )
    }
]

# -------------------------------------------------------------
# BẢNG ĐÁP ÁN RÚT GỌN (CHO BẢN GIÁO VIÊN)
# -------------------------------------------------------------
def render_summary_answer_table(doc):
    """Vẽ bảng đáp án nhanh ở cuối bản giáo viên."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    render_rich_text(p, "*BẢNG ĐÁP ÁN THAM KHẢO NHANH*", base_size=11.5, base_bold=True, base_color=COLOR_PRIMARY)

    # 1. Phần I
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(2)
    p1.paragraph_format.space_after = Pt(2)
    render_rich_text(p1, "*PHẦN I: Trắc nghiệm nhiều lựa chọn (3,0 điểm - 0,25 đ/câu)*", base_size=10.5, base_bold=True)
    
    t1 = doc.add_table(rows=2, cols=12)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    ans_p1 = ["B", "B", "C", "A", "C", "A", "D", "A", "B", "A", "B", "C"]
    for col in range(12):
        cell_h = t1.cell(0, col)
        cell_h.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell_h.paragraphs[0].paragraph_format.space_before = Pt(1)
        cell_h.paragraphs[0].paragraph_format.space_after = Pt(1)
        render_rich_text(cell_h.paragraphs[0], f"*{col+1}*", base_size=9.5, base_bold=True)
        
        cell_b = t1.cell(1, col)
        cell_b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell_b.paragraphs[0].paragraph_format.space_before = Pt(1)
        cell_b.paragraphs[0].paragraph_format.space_after = Pt(1)
        render_rich_text(cell_b.paragraphs[0], f"*{ans_p1[col]}*", base_size=10, base_bold=True, base_color=COLOR_CORRECT)

    # 2. Phần II
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(6)
    p2.paragraph_format.space_after = Pt(2)
    render_rich_text(p2, "*PHẦN II: Trắc nghiệm Đúng - Sai (2,0 điểm - 0,25 đ/ý đúng)*", base_size=10.5, base_bold=True)
    t2 = doc.add_table(rows=3, cols=5)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers2 = ["Câu", "Lệnh a)", "Lệnh b)", "Lệnh c)", "Lệnh d)"]
    for col in range(5):
        cell = t2.cell(0, col)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        render_rich_text(cell.paragraphs[0], f"*{headers2[col]}*", base_size=9.5, base_bold=True)
    
    p2_data = [
        ("Câu 1", "Đ", "S", "Đ", "Đ"),
        ("Câu 2", "Đ", "Đ", "S", "Đ")
    ]
    for r_idx, row in enumerate(p2_data):
        for c_idx, val in enumerate(row):
            cell = t2.cell(r_idx + 1, c_idx)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            color = COLOR_CORRECT if c_idx > 0 else None
            render_rich_text(cell.paragraphs[0], f"*{val}*", base_size=9.5, base_bold=True, base_color=color)

    # 3. Phần III
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(6)
    p3.paragraph_format.space_after = Pt(2)
    render_rich_text(p3, "*PHẦN III: Trắc nghiệm Trả lời ngắn (2,0 điểm - 0,5 đ/câu)*", base_size=10.5, base_bold=True)
    t3 = doc.add_table(rows=2, cols=4)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    ans_p3 = ["3", "4", "2", "16,3"]
    for col in range(4):
        cell_h = t3.cell(0, col)
        cell_h.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        render_rich_text(cell_h.paragraphs[0], f"*Câu {col+1}*", base_size=9.5, base_bold=True)
        cell_b = t3.cell(1, col)
        cell_b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        render_rich_text(cell_b.paragraphs[0], f"*{ans_p3[col]}*", base_size=10, base_bold=True, base_color=COLOR_CORRECT)

    # 4. Phần IV Thang điểm tóm tắt
    p4 = doc.add_paragraph()
    p4.paragraph_format.space_before = Pt(6)
    p4.paragraph_format.space_after = Pt(2)
    render_rich_text(p4, "*PHẦN IV: Tự luận (3,0 điểm)*", base_size=10.5, base_bold=True)
    p4_note = doc.add_paragraph()
    p4_note.paragraph_format.left_indent = Cm(0.5)
    render_rich_text(p4_note, "- *Câu 1 (1,0 điểm):* a) Nêu đúng hiện tượng ống 1 (0,25 đ) và ống 2 (0,25 đ); b) Viết đúng PTHH thủy phân (0,25 đ) và PTHH tráng bạc (0,25 đ).\n"
                             "- *Câu 2 (1,0 điểm):* a) Viết đúng PTHH (0,25 đ); b) Khối lượng glycerol = 1,656 kg (0,25 đ); c) Tính đúng khối lượng muối và số bánh xà phòng = 209 bánh (0,5 đ).\n"
                             "- *Câu 3 (1,0 điểm):* a) Khối lượng ethanol lý thuyết = 368 kg (0,5 đ); b) Khối lượng ethanol thực tế = 294,4 kg -> Thể tích cồn 70° = 525,71 lít (0,5 đ).", base_size=9.5)


# -------------------------------------------------------------
# BUILD EXAM DOCUMENT
# -------------------------------------------------------------
def build_exam_docx(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="TRƯỜNG THPT NGUYỄN HUỆ",
        exam_title="ĐỀ KIỂM TRA GIỮA KÌ I - NĂM HỌC 2026 - 2027",
        subject_line="MÔN: HOÁ HỌC - LỚP 12",
        num_pages="03" if not is_teacher else "05",
        duration="45 phút",
        exam_code="101",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; S = 32; Cl = 35,5; K = 39; Ca = 40; Ag = 108",
        is_teacher=is_teacher
    )

    # ==================== PHẦN I ====================
    add_section_header(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        "Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )
    for q in PART1_QUESTIONS:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, f"Câu {q['num']}. {q['title']}")
        render_options_no_table(doc, q['options'], is_teacher=is_teacher)
        if is_teacher:
            add_solution_no_table(doc, q['sol'])

    # ==================== PHẦN II ====================
    add_section_header(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai.",
        "Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )
    for q in PART2_QUESTIONS:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, f"Câu {q['num']}. {q['stem']}")
        for label, text, ans, lv in q['items']:
            p_item = doc.add_paragraph()
            p_item.paragraph_format.left_indent = Cm(0.5)
            p_item.paragraph_format.space_before = Pt(1)
            p_item.paragraph_format.space_after = Pt(1)
            render_statement(p_item, f"{label}) {text}", ans=ans, is_teacher=is_teacher)
        if is_teacher:
            add_solution_no_table(doc, q['sol'])

    # ==================== PHẦN III ====================
    add_section_header(
        doc,
        "PHẦN III. Câu trắc nghiệm trả lời ngắn.",
        "Thí sinh trả lời từ câu 1 đến câu 4. Mỗi câu trả lời đúng được 0,5 điểm."
    )
    for q in PART3_QUESTIONS:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, f"Câu {q['num']}. {q['stem']}")
        
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(2)
        if is_teacher:
            render_rich_text(p_ans, f"*Trả lời:* *{q['ans']}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
            add_solution_no_table(doc, q['sol'])
        else:
            render_rich_text(p_ans, "*Trả lời:* .....................................................................................", base_size=10.5)

    # ==================== PHẦN IV ====================
    add_section_header(
        doc,
        "PHẦN IV. Tự luận.",
        "Thí sinh làm bài trực tiếp vào đề thi hoặc giấy làm bài thi. Điểm tối đa: 3,0 điểm."
    )
    for q in PART4_QUESTIONS:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, f"Câu {q['num']} ({q['points']}). {q['stem']}")
        
        if is_teacher:
            add_solution_no_table(doc, q['sol'])
        else:
            add_student_writing_lines(doc, num_lines=8)

    # HẾT
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(4)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "---------- *HẾT* ----------", base_size=10.5, base_bold=True)

    if is_teacher:
        render_summary_answer_table(doc)

    return doc


if __name__ == "__main__":
    print("Đang tạo 2 file Word theo chuẩn FORM ĐỀ KIỂM TRA...")
    
    # 1. Bản Học sinh
    doc_hs = build_exam_docx(is_teacher=False)
    hs_path = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_12_THPT_NguyenHue_Ma101_HocSinh.docx")
    doc_hs.save(hs_path)
    print(f"-> Đã tạo Bản Học sinh: {hs_path}")

    # 2. Bản Giáo viên
    doc_gv = build_exam_docx(is_teacher=True)
    gv_path = os.path.join(OUTPUT_DIR, "De_Kiem_Tra_GK1_Hoa_12_THPT_NguyenHue_Ma101_GiaoVien.docx")
    doc_gv.save(gv_path)
    print(f"-> Đã tạo Bản Giáo viên: {gv_path}")
    print("Hoàn tất tạo file Word thành công 100%!")
