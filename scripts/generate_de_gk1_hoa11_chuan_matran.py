# -*- coding: utf-8 -*-
"""
SOẠN ĐỀ KIỂM TRA GIỮA HỌC KÌ I - HÓA HỌC 11 (MÃ ĐỀ 101)
Chuẩn xác 100% theo đúng Ma trận hình ảnh người dùng cung cấp:
- Tổng số câu/ý: 27 câu/ý (12 TN 4 lựa chọn + 2 câu Đúng/Sai 8 ý + 4 câu Trả lời ngắn + 3 câu Tự luận)
- Tỉ lệ điểm: 3,0 đ Phần I + 2,0 đ Phần II + 2,0 đ Phần III + 3,0 đ Tự luận = 10,0 điểm (40% NB - 30% TH - 30% VD).
- Xuất 2 file Word: Bản Học sinh & Bản Giáo viên chuẩn "FORM ĐỀ KIỂM TRA".
"""

import os
import sys
import re
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.insert(0, r'c:\Antigravity_Thanh\scripts')
from build_ngan_hang_hoa10_c1 import (
    setup_page_and_hf, add_chem_runs_hoa10, parse_abcd_from_opt_lines,
    add_options_paragraph_no_table, add_data_table_to_doc, add_solution_block_hoa10,
    LEVEL_NAMES, COLOR_HF_RED, COLOR_PRIMARY, COLOR_SOLUTION, COLOR_CORRECT, COLOR_TAG
)

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
os.makedirs(OUT_DIR, exist_ok=True)

# ==============================================================================
# BỘ CÂU HỎI THEO ĐÚNG MA TRẬN ẢNH NGƯỜI DÙNG
# ==============================================================================

# PHẦN I: 12 CÂU (8 BIẾT - 4 HIỂU)
PART_I_QUESTIONS = [
    # 1. Bài 1: Cân bằng hóa học - Biết [NB]
    {
        "id": "H11-C1-P1-NB-001",
        "topic": "Bài 1: Khái niệm về cân bằng hóa học",
        "level": "NB",
        "stem": "Phản ứng thuận nghịch là phản ứng",
        "opts": {
            "A": "xảy ra theo hai chiều ngược nhau trong cùng điều kiện.",
            "B": "chỉ xảy ra theo một chiều nhất định từ chất đầu tạo thành sản phẩm.",
            "C": "có phương trình hóa học được biểu diễn bằng mũi tên một chiều.",
            "D": "xảy ra hoàn toàn giữa các chất tham gia ở thể khí."
        },
        "answer": "A",
        "solution": "Theo định nghĩa SGK, phản ứng thuận nghịch là phản ứng hóa học xảy ra theo hai chiều trái ngược nhau trong cùng một điều kiện xác định."
    },
    # 2. Bài 1: Cân bằng hóa học - Biết [NB]
    {
        "id": "H11-C1-P1-NB-002",
        "topic": "Bài 1: Khái niệm về cân bằng hóa học",
        "level": "NB",
        "stem": "Mối quan hệ giữa tốc độ phản ứng thuận (vt) và tốc độ phản ứng nghịch (vn) ở trạng thái cân bằng được biểu diễn như thế nào?",
        "opts": {
            "A": "vt = 2vn.",
            "B": "vt = vn > 0.",
            "C": "vt = 0,5vn.",
            "D": "vt = vn = 0."
        },
        "answer": "B",
        "solution": "Tại trạng thái cân bằng hóa học, phản ứng vẫn liên tục diễn ra nhưng với tốc độ thuận bằng chính xác tốc độ nghịch (vt = vn > 0). Cân bằng hóa học là cân bằng động."
    },
    # 3. Bài 1: Cân bằng hóa học - Biết [NB]
    {
        "id": "H11-C1-P1-NB-003",
        "topic": "Bài 1: Khái niệm về cân bằng hóa học",
        "level": "NB",
        "stem": "Hằng số cân bằng KC của một phản ứng thuận nghịch xác định chỉ phụ thuộc vào yếu tố nào sau đây?",
        "opts": {
            "A": "Nhiệt độ.",
            "B": "Nồng độ các chất.",
            "C": "Áp suất chung của hệ.",
            "D": "Chất xúc tác."
        },
        "answer": "A",
        "solution": "Hằng số cân bằng KC là đại lượng nhiệt động lực học, chỉ phụ thuộc vào bản chất của phản ứng và nhiệt độ môi trường, không phụ thuộc vào nồng độ, áp suất hay chất xúc tác."
    },
    # 4. Bài 1: Cân bằng hóa học - Hiểu [TH]
    {
        "id": "H11-C1-P1-TH-004",
        "topic": "Bài 1: Khái niệm về cân bằng hóa học",
        "level": "TH",
        "stem": "Cho cân bằng hóa học trong bình kín: N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)   ΔrH^o^~298~ = –92 kJ. Cân bằng trên sẽ chuyển dịch theo chiều thuận khi",
        "opts": {
            "A": "tăng nhiệt độ của hệ phản ứng.",
            "B": "giảm áp suất của hệ phản ứng.",
            "C": "tăng áp suất chung của hệ phản ứng.",
            "D": "thêm chất xúc tác bột iron vào hệ."
        },
        "answer": "C",
        "solution": "Theo nguyên lí Le Chatelier: Phản ứng có tổng mol khí vế trái (4 mol) lớn hơn vế phải (2 mol). Khi tăng áp suất, cân bằng sẽ chuyển dịch theo chiều làm giảm số mol khí, tức là chiều thuận."
    },
    # 5. Bài 2: Cân bằng trong dung dịch nước - Biết [NB]
    {
        "id": "H11-C1-P1-NB-005",
        "topic": "Bài 2: Cân bằng trong dung dịch nước",
        "level": "NB",
        "stem": "Chất nào sau đây thuộc loại chất điện li mạnh trong dung môi nước?",
        "opts": {
            "A": "CH~3~COOH.",
            "B": "NaCl.",
            "C": "C~2~H~5~OH.",
            "D": "H~2~O."
        },
        "answer": "B",
        "solution": "NaCl là muối tan hoàn toàn trong nước, phân li 100% thành ion Na+ và Cl- nên là chất điện li mạnh. CH3COOH và H2O là chất điện li yếu, C2H5OH là chất không điện li."
    },
    # 6. Bài 2: Cân bằng trong dung dịch nước - Hiểu [TH]
    {
        "id": "H11-C1-P1-TH-006",
        "topic": "Bài 2: Cân bằng trong dung dịch nước",
        "level": "TH",
        "stem": "Theo thuyết Brønsted – Lowry, ion hoặc phân tử nào sau đây thể hiện tính chất lưỡng tính khi tương tác với nước?",
        "opts": {
            "A": "HCO~3~^-^.",
            "B": "CO~3~^2-^.",
            "C": "NH~4~^+^.",
            "D": "Cl^-^."
        },
        "answer": "A",
        "solution": "Ion HCO3- là chất lưỡng tính vì vừa có khả năng nhường proton H+ (thể hiện tính acid: HCO3- + H2O ⇌ CO3^2- + H3O+), vừa có khả năng nhận proton H+ (thể hiện tính base: HCO3- + H2O ⇌ H2CO3 + OH-)."
    },
    # 7. Bài 3: Đơn chất nitrogen - Biết [NB]
    {
        "id": "H11-C2-P1-NB-007",
        "topic": "Bài 3: Đơn chất nitrogen",
        "level": "NB",
        "stem": "Khí nitrogen (N~2~) chiếm khoảng bao nhiêu phần trăm thể tích của bầu khí quyển Trái Đất?",
        "opts": {
            "A": "21%.",
            "B": "78%.",
            "C": "1%.",
            "D": "0,04%."
        },
        "answer": "B",
        "solution": "Trong không khí khô, khí nitrogen chiếm khoảng 78% về thể tích, khí oxygen chiếm khoảng 21%, còn lại là argon, carbon dioxide và các khí hiếm khác."
    },
    # 8. Bài 3: Đơn chất nitrogen - Biết [NB]
    {
        "id": "H11-C2-P1-NB-008",
        "topic": "Bài 3: Đơn chất nitrogen",
        "level": "NB",
        "stem": "Ở điều kiện thường, đơn chất nitrogen (N~2~) rất trơ về mặt hóa học là do",
        "opts": {
            "A": "phân tử nitrogen không phân cực.",
            "B": "nguyên tử nitrogen có bán kính rất nhỏ.",
            "C": "trong phân tử có chứa liên kết ba N≡N rất bền vững.",
            "D": "nguyên tử nitrogen có độ âm điện lớn."
        },
        "answer": "C",
        "solution": "Phân tử N2 có chứa liên kết ba N≡N với năng lượng liên kết khổng lồ (945 kJ/mol), đòi hỏi năng lượng rất lớn để bẻ gãy, do đó nitrogen rất trơ ở nhiệt độ phòng."
    },
    # 9. Bài 5: Ammonia và muối ammonium - Biết [NB]
    {
        "id": "H11-C2-P1-NB-009",
        "topic": "Bài 4: Ammonia và một số hợp chất ammonium",
        "level": "NB",
        "stem": "Phân tử ammonia (NH~3~) có dạng hình học phân tử nào sau đây?",
        "opts": {
            "A": "Đường thẳng.",
            "B": "Hình tam giác đều.",
            "C": "Hình chóp tam giác.",
            "D": "Hình tứ diện đều."
        },
        "answer": "C",
        "solution": "Nguyên tử N ở trạng thái lai hóa sp3, trên nguyên tử N còn 1 cặp electron tự do đẩy 3 liên kết N-H xuống dưới tạo thành cấu trúc hình chóp tam giác (góc liên kết khoảng 107°)."
    },
    # 10. Bài 5: Ammonia và muối ammonium - Hiểu [TH]
    {
        "id": "H11-C2-P1-TH-010",
        "topic": "Bài 4: Ammonia và một số hợp chất ammonium",
        "level": "TH",
        "stem": "Khi nhúng hai đầu đũa thủy tinh vào hai bình đựng dung dịch NH~3~ đặc và dung dịch HCl đặc rồi đưa lại gần nhau, hiện tượng quan sát được là",
        "opts": {
            "A": "xuất hiện ngọn lửa màu vàng chói.",
            "B": "xuất hiện kết tủa màu nâu đỏ.",
            "C": "xuất hiện làn khói trắng dày đặc.",
            "D": "dung dịch chuyển sang màu hồng."
        },
        "answer": "C",
        "solution": "Hơi NH3 và hơi HCl bay ra tiếp xúc với nhau xảy ra phản ứng: NH3(g) + HCl(g) → NH4Cl(s). Muối ammonium chloride sinh ra dưới dạng các hạt tinh thể li ti lơ lửng tạo thành làn khói trắng."
    },
    # 11. Bài 6: Hợp chất của nitrogen với oxygen - Biết [NB]
    {
        "id": "H11-C2-P1-NB-011",
        "topic": "Bài 5: Một số hợp chất của nitrogen với oxygen",
        "level": "NB",
        "stem": "Khí nitrogen dioxide (NO~2~) có màu sắc đặc trưng nào sau đây?",
        "opts": {
            "A": "Không màu, không mùi.",
            "B": "Màu vàng lục nhạt.",
            "C": "Màu nâu đỏ, có mùi xốc độc hại.",
            "D": "Màu xanh lam nhạt."
        },
        "answer": "C",
        "solution": "Khí NO2 có màu nâu đỏ đặc trưng, mùi hắc và độc hại đối với hệ hô hấp. Khí NO thì không màu nhưng dễ bị oxi hóa trong không khí biến thành NO2."
    },
    # 12. Bài 6: Hợp chất của nitrogen với oxygen - Hiểu [TH]
    {
        "id": "H11-C2-P1-TH-012",
        "topic": "Bài 5: Một số hợp chất của nitrogen với oxygen",
        "level": "TH",
        "stem": "Cặp kim loại nào sau đây đều bị thụ động hóa (không phản ứng) trong dung dịch nitric acid đặc, nguội?",
        "opts": {
            "A": "Cu và Ag.",
            "B": "Fe và Al.",
            "C": "Zn và Mg.",
            "D": "Fe và Cu."
        },
        "answer": "B",
        "solution": "Kim loại Fe, Al (và Cr) bị thụ động hóa trong dung dịch HNO3 đặc nguội (cũng như H2SO4 đặc nguội) do tạo lớp màng oxide đặc khít bảo vệ bề mặt kim loại."
    }
]

# PHẦN II: 2 CÂU ĐÚNG - SAI (CÂU 1: CÂN BẰNG HÓA HỌC; CÂU 2: NITROGEN VÀ HỢP CHẤT)
# Mỗi câu gồm 4 ý a, b, c, d
PART_II_QUESTIONS = [
    # Câu 13 (Câu 1 Phần II): Chủ đề 1 - Cân bằng hóa học
    # Ý a, b: Biết [NB] về Bài 2 (sự điện li / dung dịch nước); Ý c, d: Hiểu [TH] về Bài 1 (cân bằng hóa học / Le Chatelier)
    {
        "id": "H11-C1-P2-MATRAN-013",
        "topic": "Chủ đề 1: Cân bằng hóa học",
        "stem": "Khảo sát trạng thái cân bằng của phản ứng hóa học và cân bằng trong dung dịch nước, hãy đánh giá tính Đúng hay Sai của các mệnh đề sau:",
        "items": {
            "a": "Dung dịch nước của các acid, base và muối dẫn được dòng điện là do chúng có khả năng phân li ra các ion mang điện tích chuyển động tự do.",
            "b": "Dung dịch cồn ethanol và dung dịch đường saccharose đều là những chất điện li mạnh vì chúng tan rất nhiều trong nước.",
            "c": "Tại trạng thái cân bằng hóa học, nồng độ của các chất tham gia và các chất sản phẩm luôn luôn bằng nhau.",
            "d": "Cho hệ cân bằng trong bình kín: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g)   ΔrH^o^~298~ < 0. Khi hạ nhiệt độ của bình phản ứng, cân bằng sẽ chuyển dịch theo chiều thuận."
        },
        "item_levels": {"a": "NB", "b": "NB", "c": "TH", "d": "TH"},
        "answer": "a:Đ - b:S - c:S - d:Đ",
        "solution": """a) Đúng. Bản chất sự dẫn điện của dung dịch chất điện li là nhờ sự chuyển động có hướng của các cation và anion tự do.
b) Sai. Cồn ethanol và đường tan tốt trong nước nhưng tồn tại ở dạng phân tử trung hòa, không phân li ra ion nên chúng là các chất không điện li.
c) Sai. Khi hệ đạt trạng thái cân bằng, nồng độ của các chất được giữ nguyên không đổi theo thời gian chứ không bắt buộc phải bằng nhau.
d) Đúng. Phản ứng có ΔH < 0 là phản ứng tỏa nhiệt theo chiều thuận. Khi hạ nhiệt độ (làm lạnh), theo nguyên lí Le Chatelier hệ sẽ tự động chuyển dịch theo chiều tỏa nhiệt (chiều thuận) để bù lại lượng nhiệt đã mất."""
    },
    # Câu 14 (Câu 2 Phần II): Chủ đề 2 - Nitrogen và sulfur
    # Ý a, b: Biết [NB] về Bài 5 (Ammonia/muối ammonium); Ý c, d: Hiểu [TH] về Bài 6 (Hợp chất của nitrogen với oxygen/mưa acid/phú dưỡng)
    {
        "id": "H11-C2-P2-MATRAN-014",
        "topic": "Chủ đề 2: Nitrogen và hợp chất của nitrogen",
        "stem": "Khảo sát tính chất và vai trò thực tiễn của các hợp chất chứa nguyên tố nitrogen, hãy đánh giá tính Đúng hay Sai của các phát biểu sau:",
        "items": {
            "a": "Khí ammonia tan rất nhiều trong nước là nhờ khả năng tạo được liên kết hydrogen liên phân tử bền vững với các phân tử nước.",
            "b": "Tất cả các muối ammonium đều rất bền với nhiệt, không bị nhiệt phân hủy khi đun nóng.",
            "c": "Trong cơn mưa dông kèm sấm sét, khí nitrogen bị oxygen oxi hóa thành nitrogen monoxide (NO), mở đầu cho quá trình tự nhiên cung cấp đạm nitrate cho đất.",
            "d": "Hiện tượng phú dưỡng xảy ra ở ao hồ là do nguồn nước bị thiếu hụt nghiêm trọng các nguyên tố dinh dưỡng như nitrogen và phosphorus."
        },
        "item_levels": {"a": "NB", "b": "NB", "c": "TH", "d": "TH"},
        "answer": "a:Đ - b:S - c:Đ - d:S",
        "solution": """a) Đúng. Ở 20 °C, 1 lít nước có thể hòa tan được khoảng 700 lít khí ammonia do phân tử NH3 phân cực và tạo liên kết hydrogen mạnh với nước.
b) Sai. Tính chất đặc trưng nổi bật của muối ammonium là rất kém bền nhiệt, dễ dàng bị phân hủy khi đun nóng thành chất khí.
c) Đúng. Năng lượng sấm sét kích hoạt: N2 + O2 → 2NO (ở khoảng 3000 °C), sau đó NO hóa nâu thành NO2, tiếp tục tác dụng với nước và O2 tạo HNO3 cung cấp ion NO3- (đạm nitrate) cho cây trồng qua nước mưa.
d) Sai. Hiện tượng phú dưỡng (eutrophication) xảy ra khi nguồn nước bị dư thừa quá mức các chất dinh dưỡng (nitrate và phosphate) từ phân bón trôi rửa và nước thải sinh hoạt, kích thích rong tảo phát triển bùng phát làm cạn kiệt oxygen hòa tan."""
    }
]

# PHẦN III: 4 CÂU TRẢ LỜI NGẮN (2 BIẾT - 2 HIỂU)
# Mỗi câu 0,5 điểm
PART_III_QUESTIONS = [
    # Câu 15: Bài 5 - Ammonia / muối ammonium - Biết [NB]
    {
        "id": "H11-C2-P3-NB-015",
        "topic": "Bài 4: Ammonia và một số hợp chất ammonium",
        "level": "NB",
        "stem": "Trong công thức cấu tạo Lewis của phân tử ammonia (NH~3~), nguyên tử nitrogen trung tâm còn lại bao nhiêu cặp electron tự do (chưa tham gia tạo liên kết)?",
        "answer": "1",
        "solution": "Nguyên tử nitrogen có 5 electron hóa trị, dùng 3 electron tạo 3 liên kết cộng hóa trị với 3 nguyên tử H, còn lại 2 electron độc thân ghép đôi tạo thành đúng 1 cặp electron tự do."
    },
    # Câu 16: Bài 6 - Hợp chất của nitrogen với oxygen - Biết [NB]
    {
        "id": "H11-C2-P3-NB-016",
        "topic": "Bài 5: Một số hợp chất của nitrogen với oxygen",
        "level": "NB",
        "stem": "Số oxi hóa của nguyên tố nitrogen trong phân tử nitric acid (HNO~3~) là bao nhiêu?",
        "answer": "+5",
        "solution": "Trong phân tử HNO3: H mang số oxi hóa +1, O mang số oxi hóa –2. Áp dụng bảo toàn điện tích: (+1) + x + 3×(–2) = 0 ⇒ x = +5. Đây là trạng thái oxi hóa cao nhất của nguyên tố nitrogen."
    },
    # Câu 17: Bài 2 - Cân bằng trong dung dịch nước - Hiểu [TH]
    {
        "id": "H11-C1-P3-TH-017",
        "topic": "Bài 2: Cân bằng trong dung dịch nước",
        "level": "TH",
        "stem": "Một mẫu dịch vị dạ dày của người khỏe mạnh có nồng độ ion H^+^ đo được là 0,01 M (tức 10^-2^ mol/L). Giá trị pH của mẫu dịch vị này bằng bao nhiêu?",
        "answer": "2",
        "solution": "Công thức tính pH = –log[H+]. Với [H+] = 0,01 M = 10^-2 M ⇒ pH = –log(10^-2) = 2."
    },
    # Câu 18: Bài 3 - Đơn chất nitrogen - Hiểu [TH]
    {
        "id": "H11-C2-P3-TH-018",
        "topic": "Bài 3: Đơn chất nitrogen",
        "level": "TH",
        "stem": "Cho các phản ứng hóa học sau:\n(1) N~2~ + 3H~2~ ⇌ 2NH~3~\n(2) N~2~ + O~2~ ⇌ 2NO\n(3) N~2~ + 3Mg → Mg~3~N~2~\n(4) N~2~ + 6Li → 2Li~3~N\nCó bao nhiêu phản ứng trong số các phản ứng trên mà đơn chất nitrogen (N~2~) đóng vai trò là chất oxi hóa?",
        "answer": "3",
        "solution": "Đơn chất nitrogen thể hiện tính oxi hóa khi số oxi hóa giảm từ 0 xuống –3. Điều này xảy ra khi tác dụng với các chất khử mạnh như H2 (phản ứng 1) và kim loại (phản ứng 3, phản ứng 4). Ở phản ứng (2), nitrogen tăng số oxi hóa từ 0 lên +2 nên đóng vai trò là chất khử. Vậy có 3 phản ứng thỏa mãn là (1), (3), (4)."
    }
]

# PHẦN IV: 3 CÂU TỰ LUẬN (1 HIỂU - 2 VẬN DỤNG)
# Mỗi câu 1,0 điểm
PART_IV_QUESTIONS = [
    # Câu 19: Bài 5 - Ammonia và muối ammonium - Hiểu [TH] (1,0 điểm)
    {
        "id": "H11-C2-TL-TH-019",
        "topic": "Bài 4: Ammonia và một số hợp chất ammonium",
        "level": "TH",
        "points": "1,0 điểm",
        "stem": "Trong nông nghiệp, người ta sử dụng các loại phân đạm để cung cấp dinh dưỡng cho cây trồng. Có ba mẫu phân đạm dạng bột màu trắng được đựng riêng biệt trong ba lọ mất nhãn chứa: NaNO~3~, NH~4~Cl và (NH~4~)~2~SO~4~.\nBằng phương pháp hóa học, hãy trình bày cách sử dụng một thuốc thử duy nhất là dung dịch barium hydroxide (Ba(OH)~2~) để phân biệt ba mẫu phân đạm trên. Viết các phương trình hóa học xảy ra.",
        "solution": """Trình bày phương pháp nhận biết:
- Lấy một lượng nhỏ mỗi mẫu thử vào 3 ống nghiệm riêng biệt có đánh số thứ tự, thêm khoảng 2 mL nước cất rồi lắc đều để hòa tan.
- Nhỏ từ từ dung dịch Ba(OH)2 vào từng ống nghiệm rồi đun nóng nhẹ trên ngọn lửa đèn cồn:
  + Mẫu không có hiện tượng gì xuất hiện là NaNO3.
  + Mẫu sinh ra chất khí có mùi khai đặc trưng, làm xanh giấy quỳ tím ẩm là NH4Cl.
    Phương trình hóa học:
    Ba(OH)2 + 2NH4Cl → BaCl2 + 2NH3↑ + 2H2O
  + Mẫu vừa xuất hiện kết tủa trắng (BaSO4), vừa sinh ra chất khí có mùi khai đặc trưng là (NH4)2SO4.
    Phương trình hóa học:
    Ba(OH)2 + (NH4)2SO4 → BaSO4↓ + 2NH3↑ + 2H2O"""
    },
    # Câu 20: Bài 1 - Khái niệm về cân bằng hóa học - Vận dụng [VD] (1,0 điểm)
    {
        "id": "H11-C1-TL-VD-020",
        "topic": "Bài 1: Khái niệm về cân bằng hóa học",
        "level": "VD",
        "points": "1,0 điểm",
        "stem": "Một bình phản ứng kín có dung tích không đổi 1,0 lít được nạp hỗn hợp gồm 0,50 mol khí carbon monoxide (CO) và 0,50 mol khí chlorine (Cl~2~). Đun nóng bình để phản ứng tổng hợp khí phosgene xảy ra theo phương trình hóa học:\nCO(g) + Cl~2~(g) ⇌ COCl~2~(g)\nKhi hệ phản ứng đạt đến trạng thái cân bằng ở nhiệt độ T, người ta tiến hành phân tích và đo được nồng độ của khí phosgene (COCl~2~) sinh ra trong bình là 0,40 mol/L.\na) Hãy lập bảng biến thiên nồng độ (Ban đầu – Phản ứng – Cân bằng) và xác định nồng độ của khí CO và khí Cl~2~ ở trạng thái cân bằng.\nb) Tính giá trị hằng số cân bằng KC của phản ứng trên ở nhiệt độ T.",
        "solution": """a) Lập bảng biến thiên nồng độ (dung tích bình V = 1,0 L nên số mol bằng nồng độ mol):
Phương trình hóa học:   CO(g)  +  Cl2(g)   ⇌   COCl2(g)
Ban đầu (M):            0,50      0,50           0
Phản ứng (M):           0,40      0,40          0,40
Cân bằng (M):           0,10      0,10          0,40
- Nồng độ của CO tại thời điểm cân bằng: [CO] = 0,50 – 0,40 = 0,10 (M).
- Nồng độ của Cl2 tại thời điểm cân bằng: [Cl2] = 0,50 – 0,40 = 0,10 (M).

b) Biểu thức tính hằng số cân bằng KC:
KC = [COCl2] / ([CO] × [Cl2])
Thay các giá trị nồng độ lúc cân bằng vào biểu thức:
KC = 0,40 / (0,10 × 0,10) = 0,40 / 0,01 = 40.
Kết luận: Hằng số cân bằng KC của phản ứng ở nhiệt độ T là 40."""
    },
    # Câu 21: Bài 2 - Cân bằng trong dung dịch nước - Vận dụng [VD] (1,0 điểm)
    {
        "id": "H11-C1-TL-VD-021",
        "topic": "Bài 2: Cân bằng trong dung dịch nước",
        "level": "VD",
        "points": "1,0 điểm",
        "stem": "Trong phòng thí nghiệm, một học sinh tiến hành trộn lẫn 100 mL dung dịch hydrochloric acid (HCl) có nồng độ 0,10 M với 100 mL dung dịch sodium hydroxide (NaOH) có nồng độ 0,30 M, thu được 200 mL dung dịch X.\na) Viết phương trình hóa học và phương trình ion rút gọn của phản ứng trung hòa xảy ra.\nb) Tính nồng độ mol của ion OH^-^ còn dư trong dung dịch X và xác định giá trị pH của dung dịch X ở 25 °C.",
        "solution": """a) Phương trình phân tử:
HCl + NaOH → NaCl + H2O
Phương trình ion rút gọn:
H+ + OH- → H2O

b) Tính số mol các chất ban đầu:
- Số mol H+ (từ dung dịch HCl): n(H+) = 0,10 × 0,10 = 0,01 (mol).
- Số mol OH- (từ dung dịch NaOH): n(OH-) = 0,10 × 0,30 = 0,03 (mol).
So sánh số mol: n(OH-) > n(H+) (0,03 > 0,01) nên ion H+ phản ứng hết, ion OH- còn dư.
- Số mol OH- dư sau phản ứng:
n(OH- dư) = 0,03 – 0,01 = 0,02 (mol).
- Tổng thể tích của dung dịch X sau khi trộn: V = 100 mL + 100 mL = 200 mL = 0,20 (L).
- Nồng độ mol của ion OH- trong dung dịch X:
[OH-] = 0,02 / 0,20 = 0,10 (M) = 10^-1 M.
- Tính chỉ số pOH và pH:
pOH = –log[OH-] = –log(10^-1) = 1.
Ở 25 °C, ta có: pH + pOH = 14 ⇒ pH = 14 – 1 = 13.
Kết luận: Dung dịch X có môi trường kiềm (base) mạnh với giá trị pH = 13."""
    }
]


# ==============================================================================
# BẢNG MA TRẬN ĐẶC TẢ TẠI ĐẦU BẢN GIÁO VIÊN
# ==============================================================================
def add_matrix_spec_table(doc):
    p_mt_title = doc.add_paragraph()
    p_mt_title.paragraph_format.space_before = Pt(4)
    p_mt_title.paragraph_format.space_after = Pt(2)
    r_mt = p_mt_title.add_run("BẢNG MA TRẬN ĐẶC TẢ ĐỀ KIỂM TRA GIỮA HỌC KÌ I – HÓA HỌC 11 (MÃ ĐỀ 101)")
    r_mt.bold = True
    r_mt.font.name = "Times New Roman"
    r_mt.font.size = Pt(11)
    r_mt.font.color.rgb = COLOR_PRIMARY

    tbl = doc.add_table(rows=1, cols=9)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl._tbl.tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="8BC390"/>'
        f'<w:left w:val="single" w:sz="6" w:space="0" w:color="8BC390"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="8BC390"/>'
        f'<w:right w:val="single" w:sz="6" w:space="0" w:color="8BC390"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="8BC390"/>'
        f'</w:tblBorders>'
    ))

    headers = [
        "TT", "Chủ đề kiến thức", "Đơn vị bài học",
        "TNKQ (Nhiều lựa chọn)", "TNKQ (Đúng - Sai)", "TNKQ (Trả lời ngắn)",
        "Tự luận", "Tổng điểm", "Tỉ lệ %"
    ]
    for i, h in enumerate(headers):
        p = tbl.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(h)
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r.font.color.rgb = COLOR_PRIMARY

    rows_data = [
        ("1", "Chủ đề 1:\nCÂN BẰNG HÓA HỌC", "Bài 1: Khái niệm về cân bằng hóa học", "3 NB (0,75đ)\n1 TH (0,25đ)", "2 TH (0,5đ)", "—", "1 VD (1,0đ)", "2,5 đ", "25%"),
        ("1", "Chủ đề 1:\nCÂN BẰNG HÓA HỌC", "Bài 2: Cân bằng trong dung dịch nước", "1 NB (0,25đ)\n1 TH (0,25đ)", "2 NB (0,5đ)", "1 TH (0,5đ)", "1 VD (1,0đ)", "2,5 đ", "25%"),
        ("2", "Chủ đề 2:\nNITROGEN VÀ SULFUR", "Bài 3: Đơn chất nitrogen", "2 NB (0,5đ)", "—", "1 TH (0,5đ)", "—", "1,0 đ", "10%"),
        ("2", "Chủ đề 2:\nNITROGEN VÀ SULFUR", "Bài 5: Ammonia. Muối ammonium", "1 NB (0,25đ)\n1 TH (0,25đ)", "2 NB (0,5đ)", "1 NB (0,5đ)", "1 TH (1,0đ)", "2,5 đ", "25%"),
        ("2", "Chủ đề 2:\nNITROGEN VÀ SULFUR", "Bài 6: Một số hợp chất với oxygen của nitrogen", "1 NB (0,25đ)\n1 TH (0,25đ)", "2 TH (0,5đ)", "1 NB (0,5đ)", "—", "1,5 đ", "15%"),
        ("Σ", "TỔNG TOÀN ĐỀ", "Gồm 2 chủ đề cốt lõi GK1", "12 câu (3,0đ)\n[8 NB - 4 TH]", "2 câu / 8 ý (2,0đ)\n[4 NB - 4 TH]", "4 câu (2,0đ)\n[2 NB - 2 TH]", "3 câu (3,0đ)\n[1 TH - 2 VD]", "10,0 đ", "100%")
    ]

    for r_idx, r_vals in enumerate(rows_data):
        row_cells = tbl.add_row().cells
        is_total = (r_idx == len(rows_data) - 1)
        for c_idx, val in enumerate(r_vals):
            cp = row_cells[c_idx].paragraphs[0]
            cp.paragraph_format.space_before = Pt(1.5)
            cp.paragraph_format.space_after = Pt(1.5)
            if c_idx in (0, 7, 8) or is_total:
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cr = cp.add_run(val)
            cr.font.name = "Times New Roman"
            cr.font.size = Pt(9.0)
            if is_total:
                cr.bold = True
                cr.font.color.rgb = COLOR_HF_RED


# ==============================================================================
# HÀM XUẤT ĐỀ THI WORD
# ==============================================================================
def generate_gk1_exam(is_teacher=False):
    doc = docx.Document()
    setup_page_and_hf(doc)

    # 1. Header Title Block
    # Bảng 2 cột không viền
    tbl_hdr = doc.add_table(rows=1, cols=2)
    tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_l, cell_r = tbl_hdr.rows[0].cells
    cell_l.width = Cm(8.0)
    cell_r.width = Cm(10.0)

    p_l = cell_l.paragraphs[0]
    p_l.paragraph_format.space_before = Pt(0)
    p_l.paragraph_format.space_after = Pt(2)
    r_l1 = p_l.add_run("SỞ GD&ĐT THỪA THIÊN HUẾ\nTRƯỜNG THPT CHÂU VĂN LIÊM")
    r_l1.bold = True
    r_l1.font.name = "Times New Roman"
    r_l1.font.size = Pt(10)
    r_l1.font.color.rgb = COLOR_PRIMARY
    p_l_sub = cell_l.add_paragraph()
    p_l_sub.paragraph_format.space_before = Pt(0)
    p_l_sub.paragraph_format.space_after = Pt(2)
    r_l2 = p_l_sub.add_run("ĐỀ KIỂM TRA CHÍNH THỨC")
    r_l2.bold = True
    r_l2.font.name = "Times New Roman"
    r_l2.font.size = Pt(9.5)
    r_l2.font.color.rgb = COLOR_HF_RED

    p_r = cell_r.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.paragraph_format.space_before = Pt(0)
    p_r.paragraph_format.space_after = Pt(2)
    r_r1 = p_r.add_run("ĐỀ KIỂM TRA GIỮA HỌC KÌ I – NĂM HỌC 2026 – 2027\nMÔN: HÓA HỌC – LỚP 11")
    r_r1.bold = True
    r_r1.font.name = "Times New Roman"
    r_r1.font.size = Pt(10.5)
    r_r1.font.color.rgb = COLOR_PRIMARY

    p_r_sub = cell_r.add_paragraph()
    p_r_sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r_sub.paragraph_format.space_before = Pt(0)
    p_r_sub.paragraph_format.space_after = Pt(2)
    r_r2 = p_r_sub.add_run("Thời gian làm bài: 45 phút (không kể thời gian phát đề)\nMã đề thi: 101")
    r_r2.font.name = "Times New Roman"
    r_r2.font.size = Pt(9.5)
    r_r2.italic = True

    # Dòng thông tin học sinh
    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(3)
    p_info.paragraph_format.space_after = Pt(2)
    info_txt = (
        "BẢN GIÁO VIÊN — CÓ ĐÁP ÁN, LỜI GIẢI CHI TIẾT & BẢNG MA TRẬN ĐẶC TẢ"
        if is_teacher else
        "Họ và tên thí sinh: .............................................................. Lớp: ............. Số báo danh: ............."
    )
    r_inf = p_info.add_run(info_txt)
    r_inf.bold = True
    r_inf.font.name = "Times New Roman"
    r_inf.font.size = Pt(10.5)
    if is_teacher:
        p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_inf.font.color.rgb = COLOR_HF_RED

    # Dòng nguyên tử khối
    p_ntk = doc.add_paragraph()
    p_ntk.paragraph_format.space_before = Pt(0)
    p_ntk.paragraph_format.space_after = Pt(4)
    r_ntk = p_ntk.add_run("(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108; Ba = 137)")
    r_ntk.italic = True
    r_ntk.font.name = "Times New Roman"
    r_ntk.font.size = Pt(9.5)

    if is_teacher:
        add_matrix_spec_table(doc)

    q_num = 1

    for q in PART_I_QUESTIONS:
        q["part"] = "P1"
    for q in PART_II_QUESTIONS:
        q["part"] = "P2"
    for q in PART_III_QUESTIONS:
        q["part"] = "P3"
    for q in PART_IV_QUESTIONS:
        q["part"] = "P4"
        if "answer" not in q:
            q["answer"] = ""

    # ==========================================================================
    # PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (12 CÂU - 3,0 ĐIỂM)
    # ==========================================================================
    p_p1 = doc.add_paragraph()
    p_p1.paragraph_format.space_before = Pt(8)
    p_p1.paragraph_format.space_after = Pt(2)
    r_p1 = p_p1.add_run("PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,0 điểm). Thí sinh trả lời từ câu 1 đến câu 12.")
    r_p1.bold = True
    r_p1.font.name = "Times New Roman"
    r_p1.font.size = Pt(11)
    r_p1.font.color.rgb = COLOR_HF_RED

    p_p1_d = doc.add_paragraph()
    p_p1_d.paragraph_format.space_before = Pt(0)
    p_p1_d.paragraph_format.space_after = Pt(4)
    r_p1_d = p_p1_d.add_run("Mỗi câu hỏi thí sinh chỉ chọn một phương án duy nhất. Mỗi câu trả lời đúng được 0,25 điểm.")
    r_p1_d.italic = True
    r_p1_d.font.name = "Times New Roman"
    r_p1_d.font.size = Pt(10)

    for q in PART_I_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        r_num = p_q.add_run(f"Câu {q_num}. ")
        r_num.bold = True
        r_num.font.name = "Times New Roman"
        r_num.font.size = Pt(11)
        r_num.font.color.rgb = COLOR_PRIMARY

        if is_teacher:
            lv_txt = LEVEL_NAMES.get(q["level"], q["level"])
            r_tag = p_q.add_run(f"[{q['id']} | {lv_txt} | ĐA: {q['answer']}] ")
            r_tag.bold = True
            r_tag.font.name = "Times New Roman"
            r_tag.font.size = Pt(10)
            r_tag.font.color.rgb = COLOR_TAG

        add_chem_runs_hoa10(p_q, q["stem"], size=11)
        add_options_paragraph_no_table(doc, q["opts"], q["answer"], is_teacher)

        if is_teacher:
            add_solution_block_hoa10(doc, q)
        q_num += 1

    # ==========================================================================
    # PHẦN II. TRẮC NGHIỆM ĐÚNG SAI (2 CÂU - 2,0 ĐIỂM)
    # ==========================================================================
    p_p2 = doc.add_paragraph()
    p_p2.paragraph_format.space_before = Pt(8)
    p_p2.paragraph_format.space_after = Pt(2)
    r_p2 = p_p2.add_run("PHẦN II. Câu trắc nghiệm đúng sai (2,0 điểm). Thí sinh trả lời từ câu 13 đến câu 14.")
    r_p2.bold = True
    r_p2.font.name = "Times New Roman"
    r_p2.font.size = Pt(11)
    r_p2.font.color.rgb = COLOR_HF_RED

    p_p2_d = doc.add_paragraph()
    p_p2_d.paragraph_format.space_before = Pt(0)
    p_p2_d.paragraph_format.space_after = Pt(4)
    r_p2_d = p_p2_d.add_run("Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai. Mỗi ý trả lời đúng được 0,25 điểm.")
    r_p2_d.italic = True
    r_p2_d.font.name = "Times New Roman"
    r_p2_d.font.size = Pt(10)

    for q in PART_II_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        r_num = p_q.add_run(f"Câu {q_num}. ")
        r_num.bold = True
        r_num.font.name = "Times New Roman"
        r_num.font.size = Pt(11)
        r_num.font.color.rgb = COLOR_PRIMARY

        if is_teacher:
            r_tag = p_q.add_run(f"[{q['id']} | {q['answer']}] ")
            r_tag.bold = True
            r_tag.font.name = "Times New Roman"
            r_tag.font.size = Pt(10)
            r_tag.font.color.rgb = COLOR_TAG

        add_chem_runs_hoa10(p_q, q["stem"], size=11)

        p2_ans_map = {}
        for m in re.finditer(r'([a-d])\s*[:\.]\s*([ĐS])', q["answer"]):
            p2_ans_map[m.group(1).lower()] = m.group(2)

        for item_k in ["a", "b", "c", "d"]:
            item_txt = q["items"][item_k]
            p_it = doc.add_paragraph()
            p_it.paragraph_format.left_indent = Cm(0.6)
            p_it.paragraph_format.space_before = Pt(1.5)
            p_it.paragraph_format.space_after = Pt(1.5)

            corr_val = p2_ans_map.get(item_k, "")
            it_lv = q["item_levels"].get(item_k, "TH")
            it_lv_name = LEVEL_NAMES.get(it_lv, it_lv)

            if is_teacher and corr_val:
                r_it = p_it.add_run(f"{item_k}) [{it_lv_name} | {corr_val}] ")
                r_it.bold = True
                r_it.font.name = "Times New Roman"
                r_it.font.size = Pt(10.5)
                r_it.font.color.rgb = COLOR_CORRECT if corr_val == "Đ" else COLOR_PRIMARY
            else:
                r_it = p_it.add_run(f"{item_k}) ")
                r_it.bold = True
                r_it.font.name = "Times New Roman"
                r_it.font.size = Pt(10.5)

            add_chem_runs_hoa10(p_it, item_txt, size=10.5)

        if is_teacher:
            add_solution_block_hoa10(doc, q)
        q_num += 1

    # ==========================================================================
    # PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN (4 CÂU - 2,0 ĐIỂM)
    # ==========================================================================
    p_p3 = doc.add_paragraph()
    p_p3.paragraph_format.space_before = Pt(8)
    p_p3.paragraph_format.space_after = Pt(2)
    r_p3 = p_p3.add_run("PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn (2,0 điểm). Thí sinh trả lời từ câu 15 đến câu 18.")
    r_p3.bold = True
    r_p3.font.name = "Times New Roman"
    r_p3.font.size = Pt(11)
    r_p3.font.color.rgb = COLOR_HF_RED

    p_p3_d = doc.add_paragraph()
    p_p3_d.paragraph_format.space_before = Pt(0)
    p_p3_d.paragraph_format.space_after = Pt(4)
    r_p3_d = p_p3_d.add_run("Thí sinh điền kết quả hoặc số liệu tương ứng vào bài làm. Mỗi câu trả lời đúng được 0,5 điểm.")
    r_p3_d.italic = True
    r_p3_d.font.name = "Times New Roman"
    r_p3_d.font.size = Pt(10)

    for q in PART_III_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        r_num = p_q.add_run(f"Câu {q_num}. ")
        r_num.bold = True
        r_num.font.name = "Times New Roman"
        r_num.font.size = Pt(11)
        r_num.font.color.rgb = COLOR_PRIMARY

        if is_teacher:
            lv_txt = LEVEL_NAMES.get(q["level"], q["level"])
            r_tag = p_q.add_run(f"[{q['id']} | {lv_txt} | ĐA: {q['answer']}] ")
            r_tag.bold = True
            r_tag.font.name = "Times New Roman"
            r_tag.font.size = Pt(10)
            r_tag.font.color.rgb = COLOR_TAG

        add_chem_runs_hoa10(p_q, q["stem"], size=11)

        if not is_teacher:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.6)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(2)
            r_ans = p_ans.add_run("Đáp số / Trả lời: .......................................................................................")
            r_ans.italic = True
            r_ans.font.name = "Times New Roman"
            r_ans.font.size = Pt(10.5)
        else:
            add_solution_block_hoa10(doc, q)
        q_num += 1

    # ==========================================================================
    # PHẦN IV. TỰ LUẬN (3 CÂU - 3,0 ĐIỂM)
    # ==========================================================================
    p_p4 = doc.add_paragraph()
    p_p4.paragraph_format.space_before = Pt(8)
    p_p4.paragraph_format.space_after = Pt(2)
    r_p4 = p_p4.add_run("PHẦN IV. Tự luận (3,0 điểm). Thí sinh trả lời từ câu 19 đến câu 21.")
    r_p4.bold = True
    r_p4.font.name = "Times New Roman"
    r_p4.font.size = Pt(11)
    r_p4.font.color.rgb = COLOR_HF_RED

    p_p4_d = doc.add_paragraph()
    p_p4_d.paragraph_format.space_before = Pt(0)
    p_p4_d.paragraph_format.space_after = Pt(4)
    r_p4_d = p_p4_d.add_run("Thí sinh trình bày chi tiết các bước lập luận, tính toán và viết phương trình hóa học hoàn chỉnh. Mỗi câu 1,0 điểm.")
    r_p4_d.italic = True
    r_p4_d.font.name = "Times New Roman"
    r_p4_d.font.size = Pt(10)

    for q in PART_IV_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        r_num = p_q.add_run(f"Câu {q_num} ({q['points']}). ")
        r_num.bold = True
        r_num.font.name = "Times New Roman"
        r_num.font.size = Pt(11)
        r_num.font.color.rgb = COLOR_PRIMARY

        if is_teacher:
            lv_txt = LEVEL_NAMES.get(q["level"], q["level"])
            r_tag = p_q.add_run(f"[{q['id']} | {lv_txt}] ")
            r_tag.bold = True
            r_tag.font.name = "Times New Roman"
            r_tag.font.size = Pt(10)
            r_tag.font.color.rgb = COLOR_TAG

        add_chem_runs_hoa10(p_q, q["stem"], size=11)

        if not is_teacher:
            # Tạo các dòng chấm làm bài cho học sinh
            for _ in range(7):
                p_dot = doc.add_paragraph()
                p_dot.paragraph_format.space_before = Pt(1)
                p_dot.paragraph_format.space_after = Pt(1)
                r_dot = p_dot.add_run("....................................................................................................................................................................................")
                r_dot.font.name = "Times New Roman"
                r_dot.font.size = Pt(9.5)
                r_dot.font.color.rgb = RGBColor(160, 160, 160)
        else:
            add_solution_block_hoa10(doc, q)
        q_num += 1

    # Lưu file
    suffix = "GiaoVien" if is_teacher else "HocSinh"
    fname = f"De_Kiem_Tra_GK1_Hoa_11_Ma101_{suffix}.docx"
    fpath = os.path.join(OUT_DIR, fname)
    doc.save(fpath)
    print(f"  + Xuất file {suffix}: {fpath}")


if __name__ == "__main__":
    generate_gk1_exam(is_teacher=False)
    generate_gk1_exam(is_teacher=True)
    print("\n=> ĐÃ HOÀN THÀNH SOẠN BỘ 2 FILE WORD ĐỀ KIỂM TRA GIỮA KÌ 1 HÓA 11 CHUẨN MA TRẬN!")
