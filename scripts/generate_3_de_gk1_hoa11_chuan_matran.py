# -*- coding: utf-8 -*-
"""
CHƯƠNG TRÌNH SOẠN THẢO 3 ĐỀ KIỂM TRA GIỮA HỌC KÌ I - HÓA HỌC 11 THEO MA TRẬN
Mã đề: 101 (Đề 01), 102 (Đề 02), 103 (Đề 03)
Đúng chuẩn 100% Ma trận người dùng gửi:
- Phần I (3,0 điểm): 12 câu trắc nghiệm nhiều lựa chọn (8 Biết - 4 Hiểu)
- Phần II (2,0 điểm): 2 câu trắc nghiệm Đúng - Sai (8 ý: 4 Biết - 4 Hiểu)
- Phần III (2,0 điểm): 4 câu trắc nghiệm Trả lời ngắn (2 Biết - 2 Hiểu)
- Phần IV (3,0 điểm): 3 câu Tự luận (1 Hiểu - 2 Vận dụng)
Tổng điểm: 10,0 điểm (40% NB - 30% TH - 30% VD).
Xuất đủ 6 file Word (3 bản Học Sinh + 3 bản Giáo Viên) chuẩn "FORM ĐỀ KIỂM TRA".
"""

import os
import sys
import re
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.insert(0, r'c:\Antigravity_Thanh\scripts')
from form_de_kiem_tra import (
    setup_header_footer, build_title_block, add_section_header,
    render_question_prompt, render_options_no_table, render_statement,
    add_solution_no_table, add_student_writing_lines, render_rich_text,
    COLOR_HF_RED, COLOR_PRIMARY, COLOR_SOLUTION, COLOR_CORRECT
)

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
os.makedirs(OUT_DIR, exist_ok=True)

# ==============================================================================
# DỮ LIỆU ĐỀ 01 - MÃ ĐỀ 101
# ==============================================================================
DE01_DATA = {
    "exam_code": "101",
    "exam_num": "01",
    "p1": [
        {
            "num": 1, "topic": "Bài 1", "level": "NB",
            "stem": "Phản ứng thuận nghịch là phản ứng",
            "opts": [
                "*A. xảy ra theo hai chiều ngược nhau trong cùng điều kiện.*",
                "B. chỉ xảy ra theo một chiều nhất định từ chất đầu tạo thành sản phẩm.",
                "C. có phương trình hóa học được biểu diễn bằng mũi tên một chiều.",
                "D. xảy ra hoàn toàn giữa các chất tham gia ở thể khí."
            ],
            "solution": "Theo định nghĩa SGK Hóa học 11, phản ứng thuận nghịch là phản ứng hóa học xảy ra theo hai chiều trái ngược nhau trong cùng một điều kiện xác định."
        },
        {
            "num": 2, "topic": "Bài 1", "level": "NB",
            "stem": "Mối quan hệ giữa tốc độ phản ứng thuận (v~t~) và tốc độ phản ứng nghịch (v~n~) ở trạng thái cân bằng được biểu diễn như thế nào?",
            "opts": [
                "A. v~t~ = 2v~n~.",
                "*B. v~t~ = v~n~ > 0.*",
                "C. v~t~ = 0,5v~n~.",
                "D. v~t~ = v~n~ = 0."
            ],
            "solution": "Tại trạng thái cân bằng hóa học, phản ứng vẫn liên tục diễn ra nhưng với tốc độ thuận bằng chính xác tốc độ nghịch (v~t~ = v~n~ > 0). Cân bằng hóa học là cân bằng động."
        },
        {
            "num": 3, "topic": "Bài 1", "level": "NB",
            "stem": "Hằng số cân bằng K~C~ của một phản ứng thuận nghịch xác định chỉ phụ thuộc vào yếu tố nào sau đây?",
            "opts": [
                "*A. Nhiệt độ.*",
                "B. Nồng độ các chất.",
                "C. Áp suất chung của hệ.",
                "D. Chất xúc tác."
            ],
            "solution": "Hằng số cân bằng K~C~ là đại lượng nhiệt động lực học, chỉ phụ thuộc vào bản chất của phản ứng và nhiệt độ môi trường, không phụ thuộc vào nồng độ, áp suất hay chất xúc tác."
        },
        {
            "num": 4, "topic": "Bài 1", "level": "TH",
            "stem": "Cho cân bằng hóa học trong bình kín: N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)   Δ~r~H^o^~298~ = –92 kJ. Cân bằng trên sẽ chuyển dịch theo chiều thuận khi",
            "opts": [
                "A. tăng nhiệt độ của hệ phản ứng.",
                "B. giảm áp suất của hệ phản ứng.",
                "*C. tăng áp suất chung của hệ phản ứng.*",
                "D. thêm chất xúc tác bột iron vào hệ."
            ],
            "solution": "Theo nguyên lí Le Chatelier: Phản ứng có tổng mol khí vế trái (4 mol) lớn hơn vế phải (2 mol). Khi tăng áp suất, cân bằng sẽ chuyển dịch theo chiều làm giảm số mol khí, tức là chiều thuận."
        },
        {
            "num": 5, "topic": "Bài 2", "level": "NB",
            "stem": "Chất nào sau đây thuộc loại chất điện li mạnh trong dung môi nước?",
            "opts": [
                "A. CH~3~COOH.",
                "*B. NaCl.*",
                "C. C~2~H~5~OH.",
                "D. H~2~O."
            ],
            "solution": "NaCl là muối tan hoàn toàn trong nước, phân li 100% thành ion Na^+^ và Cl^-^ nên là chất điện li mạnh. CH~3~COOH và H~2~O là chất điện li yếu, C~2~H~5~OH là chất không điện li."
        },
        {
            "num": 6, "topic": "Bài 2", "level": "TH",
            "stem": "Theo thuyết Brønsted – Lowry, ion hoặc phân tử nào sau đây thể hiện tính chất lưỡng tính khi tương tác với nước?",
            "opts": [
                "*A. HCO~3~^-^.*",
                "B. CO~3~^2-^.",
                "C. NH~4~^+^.",
                "D. Cl^-^."
            ],
            "solution": "Ion HCO~3~^-^ là chất lưỡng tính vì vừa có khả năng nhường proton H^+^ (thể hiện tính acid: HCO~3~^-^ + H~2~O ⇌ CO~3~^2-^ + H~3~O^+^), vừa có khả năng nhận proton H^+^ (thể hiện tính base: HCO~3~^-^ + H~2~O ⇌ H~2~CO~3~ + OH^-^)."
        },
        {
            "num": 7, "topic": "Bài 4", "level": "NB",
            "stem": "Khí nitrogen (N~2~) chiếm khoảng bao nhiêu phần trăm thể tích của bầu khí quyển Trái Đất?",
            "opts": [
                "A. 21%.",
                "*B. 78%.*",
                "C. 1%.",
                "D. 0,04%."
            ],
            "solution": "Trong không khí khô, khí nitrogen chiếm khoảng 78% về thể tích, khí oxygen chiếm khoảng 21%, còn lại là argon, carbon dioxide và các khí hiếm khác."
        },
        {
            "num": 8, "topic": "Bài 4", "level": "NB",
            "stem": "Ở điều kiện thường, đơn chất nitrogen (N~2~) rất trơ về mặt hóa học là do",
            "opts": [
                "A. phân tử nitrogen không phân cực.",
                "B. nguyên tử nitrogen có bán kính rất nhỏ.",
                "*C. trong phân tử có chứa liên kết ba N≡N rất bền vững.*",
                "D. nguyên tử nitrogen có độ âm điện lớn."
            ],
            "solution": "Phân tử N~2~ có chứa liên kết ba N≡N với năng lượng liên kết khổng lồ (945 kJ/mol), đòi hỏi năng lượng rất lớn để bẻ gãy, do đó nitrogen rất trơ ở nhiệt độ phòng."
        },
        {
            "num": 9, "topic": "Bài 5", "level": "NB",
            "stem": "Phân tử ammonia (NH~3~) có dạng hình học phân tử nào sau đây?",
            "opts": [
                "A. Đường thẳng.",
                "B. Hình tam giác đều.",
                "*C. Hình chóp tam giác.*",
                "D. Hình tứ diện đều."
            ],
            "solution": "Nguyên tử N ở trạng thái lai hóa sp^3^, trên nguyên tử N còn 1 cặp electron tự do đẩy 3 liên kết N–H xuống dưới tạo thành cấu trúc hình chóp tam giác (góc liên kết khoảng 107°)."
        },
        {
            "num": 10, "topic": "Bài 5", "level": "TH",
            "stem": "Khi nhúng hai đầu đũa thủy tinh vào hai bình đựng dung dịch NH~3~ đặc và dung dịch HCl đặc rồi đưa lại gần nhau, hiện tượng quan sát được là",
            "opts": [
                "A. xuất hiện ngọn lửa màu vàng chói.",
                "B. xuất hiện kết tủa màu nâu đỏ.",
                "*C. xuất hiện làn khói trắng dày đặc.*",
                "D. dung dịch chuyển sang màu hồng."
            ],
            "solution": "Hơi NH~3~ và hơi HCl bay ra tiếp xúc với nhau xảy ra phản ứng: NH~3~(g) + HCl(g) → NH~4~Cl(s). Muối ammonium chloride sinh ra dưới dạng các hạt tinh thể li ti lơ lửng tạo thành làn khói trắng."
        },
        {
            "num": 11, "topic": "Bài 6", "level": "NB",
            "stem": "Khí nitrogen dioxide (NO~2~) có màu sắc đặc trưng nào sau đây?",
            "opts": [
                "A. Không màu, không mùi.",
                "B. Màu vàng lục nhạt.",
                "*C. Màu nâu đỏ, có mùi xốc độc hại.*",
                "D. Màu xanh lam nhạt."
            ],
            "solution": "Khí NO~2~ có màu nâu đỏ đặc trưng, mùi hắc và độc hại đối với hệ hô hấp. Khí NO thì không màu nhưng dễ bị oxi hóa trong không khí biến thành NO~2~."
        },
        {
            "num": 12, "topic": "Bài 6", "level": "TH",
            "stem": "Cặp kim loại nào sau đây đều bị thụ động hóa (không phản ứng) trong dung dịch nitric acid đặc, nguội?",
            "opts": [
                "A. Cu và Ag.",
                "*B. Fe và Al.*",
                "C. Zn và Mg.",
                "D. Fe và Cu."
            ],
            "solution": "Kim loại Fe, Al (và Cr) bị thụ động hóa trong dung dịch HNO~3~ đặc nguội do tạo lớp màng oxide đặc khít bảo vệ bề mặt kim loại."
        }
    ],
    "p2": [
        {
            "num": 1,
            "stem": "Khảo sát trạng thái cân bằng của phản ứng hóa học và cân bằng trong dung dịch nước, hãy đánh giá tính Đúng hay Sai của các mệnh đề sau:",
            "items": [
                ("a", "Dung dịch nước của các acid, base và muối dẫn được dòng điện là do chúng có khả năng phân li ra các ion mang điện tích chuyển động tự do.", "Đúng"),
                ("b", "Dung dịch cồn ethanol và dung dịch đường saccharose đều là những chất điện li mạnh vì chúng tan rất nhiều trong nước.", "Sai"),
                ("c", "Tại trạng thái cân bằng hóa học, nồng độ của các chất tham gia và các chất sản phẩm luôn luôn bằng nhau.", "Sai"),
                ("d", "Cho hệ cân bằng trong bình kín: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g)   Δ~r~H^o^~298~ < 0. Khi hạ nhiệt độ của bình phản ứng, cân bằng sẽ chuyển dịch theo chiều thuận.", "Đúng")
            ],
            "solution": "a) Đúng. Bản chất sự dẫn điện của dung dịch chất điện li là nhờ sự chuyển động có hướng của các cation và anion tự do.\nb) Sai. Cồn ethanol và đường tan tốt trong nước nhưng tồn tại ở dạng phân tử trung hòa, không phân li ra ion nên chúng là các chất không điện li.\nc) Sai. Khi hệ đạt trạng thái cân bằng, nồng độ của các chất được giữ nguyên không đổi theo thời gian chứ không bắt buộc phải bằng nhau.\nd) Đúng. Phản ứng có ΔH < 0 là phản ứng tỏa nhiệt theo chiều thuận. Khi hạ nhiệt độ, hệ sẽ tự động chuyển dịch theo chiều tỏa nhiệt (chiều thuận) để bù lại lượng nhiệt đã mất."
        },
        {
            "num": 2,
            "stem": "Khảo sát tính chất và vai trò thực tiễn của các hợp chất chứa nguyên tố nitrogen, hãy đánh giá tính Đúng hay Sai của các phát biểu sau:",
            "items": [
                ("a", "Khí ammonia tan rất nhiều trong nước là nhờ khả năng tạo được liên kết hydrogen liên phân tử bền vững với các phân tử nước.", "Đúng"),
                ("b", "Tất cả các muối ammonium đều rất bền với nhiệt, không bị nhiệt phân hủy khi đun nóng.", "Sai"),
                ("c", "Trong cơn mưa dông kèm sấm sét, khí nitrogen bị oxygen oxi hóa thành nitrogen monoxide (NO), mở đầu cho quá trình tự nhiên cung cấp đạm nitrate cho đất.", "Đúng"),
                ("d", "Hiện tượng phú dưỡng xảy ra ở ao hồ là do nguồn nước bị thiếu hụt nghiêm trọng các nguyên tố dinh dưỡng như nitrogen và phosphorus.", "Sai")
            ],
            "solution": "a) Đúng. Ở 20 °C, 1 lít nước có thể hòa tan được khoảng 700 lít khí ammonia do phân tử NH~3~ phân cực và tạo liên kết hydrogen mạnh với nước.\nb) Sai. Tính chất đặc trưng nổi bật của muối ammonium là rất kém bền nhiệt, dễ dàng bị phân hủy khi đun nóng thành chất khí.\nc) Đúng. Năng lượng sấm sét kích hoạt: N~2~ + O~2~ → 2NO (ở khoảng 3000 °C), sau đó NO hóa nâu thành NO~2~, tiếp tục tác dụng với nước và O~2~ tạo HNO~3~ cung cấp đạm nitrate cho cây trồng.\nd) Sai. Hiện tượng phú dưỡng xảy ra khi nguồn nước bị dư thừa quá mức các chất dinh dưỡng (nitrate và phosphate), kích thích rong tảo bùng phát."
        }
    ],
    "p3": [
        {
            "num": 1, "topic": "Bài 4", "level": "NB",
            "stem": "Trong một phân tử đơn chất nitrogen (N~2~), có bao nhiêu liên kết cộng hóa trị được hình thành giữa hai nguyên tử nitrogen?",
            "answer": "3",
            "solution": "Giữa hai nguyên tử nitrogen hình thành một liên kết ba N≡N (gồm 1 liên kết σ và 2 liên kết π), tức là có đúng 3 liên kết cộng hóa trị."
        },
        {
            "num": 2, "topic": "Bài 5", "level": "NB",
            "stem": "Trong công thức cấu tạo Lewis của phân tử ammonia (NH~3~), nguyên tử nitrogen trung tâm còn lại bao nhiêu cặp electron tự do (chưa tham gia tạo liên kết)?",
            "answer": "1",
            "solution": "Nguyên tử N có 5 electron hóa trị, dùng 3 electron tạo liên kết với 3 nguyên tử H, còn lại 2 electron ghép đôi tạo thành đúng 1 cặp electron tự do."
        },
        {
            "num": 3, "topic": "Bài 2", "level": "TH",
            "stem": "Một mẫu dung dịch acid có nồng độ ion H^+^ đo được là 10^-3^ M (tức 0,001 mol/L). Giá trị pH của mẫu dung dịch này bằng bao nhiêu?",
            "answer": "3",
            "solution": "Công thức tính: pH = –log[H^+^] = –log(10^-3^) = 3."
        },
        {
            "num": 4, "topic": "Bài 6", "level": "TH",
            "stem": "Cho các phản ứng hóa học sau:\n(1) N~2~ + 3H~2~ ⇌ 2NH~3~\n(2) N~2~ + O~2~ ⇌ 2NO\n(3) N~2~ + 3Mg → Mg~3~N~2~\n(4) N~2~ + 6Li → 2Li~3~N\nCó bao nhiêu phản ứng trong số các phản ứng trên mà đơn chất nitrogen (N~2~) đóng vai trò là chất oxi hóa?",
            "answer": "3",
            "solution": "Đơn chất nitrogen thể hiện tính oxi hóa khi số oxi hóa giảm từ 0 xuống –3. Điều này xảy ra ở phản ứng (1), (3) và (4). Ở phản ứng (2), số oxi hóa của N tăng từ 0 lên +2 nên N~2~ là chất khử. Vậy có 3 phản ứng thỏa mãn."
        }
    ],
    "p4": [
        {
            "num": 1, "topic": "Bài 1", "level": "VD", "points": "1,0 điểm",
            "stem": "Một bình phản ứng kín có dung tích không đổi 1,0 lít được nạp hỗn hợp gồm 0,50 mol khí carbon monoxide (CO) và 0,50 mol khí chlorine (Cl~2~). Đun nóng bình để phản ứng xảy ra theo phương trình hóa học:\nCO(g) + Cl~2~(g) ⇌ COCl~2~(g)\nKhi hệ phản ứng đạt đến trạng thái cân bằng ở nhiệt độ T, nồng độ của khí phosgene (COCl~2~) đo được là 0,40 mol/L.\na) Hãy lập bảng biến thiên nồng độ và xác định nồng độ của khí CO và Cl~2~ ở trạng thái cân bằng.\nb) Tính giá trị hằng số cân bằng K~C~ của phản ứng ở nhiệt độ T.",
            "solution": "a) Lập bảng biến thiên nồng độ (dung tích bình V = 1,0 L nên nồng độ mol bằng số mol):\nPhương trình hóa học:   CO(g)  +  Cl~2~(g)   ⇌   COCl~2~(g)\nBan đầu (M):            0,50      0,50           0\nPhản ứng (M):           0,40      0,40          0,40\nCân bằng (M):           0,10      0,10          0,40\n- Nồng độ của CO tại thời điểm cân bằng: [CO] = 0,50 – 0,40 = 0,10 (M).\n- Nồng độ của Cl~2~ tại thời điểm cân bằng: [Cl~2~] = 0,50 – 0,40 = 0,10 (M).\nb) Giá trị hằng số cân bằng K~C~:\nK~C~ = [COCl~2~] / ([CO] . [Cl~2~]) = 0,40 / (0,10 . 0,10) = 40."
        },
        {
            "num": 2, "topic": "Bài 2", "level": "VD", "points": "1,0 điểm",
            "stem": "Trộn 100 mL dung dịch hydrochloric acid (HCl) nồng độ 0,15 M với 100 mL dung dịch barium hydroxide (Ba(OH)~2~) nồng độ 0,10 M thu được 200 mL dung dịch X.\na) Viết phương trình ion rút gọn của phản ứng trung hòa xảy ra.\nb) Tính nồng độ mol của ion OH^-^ dư và giá trị pH của dung dịch X ở 25 °C.",
            "solution": "a) Phương trình ion rút gọn của phản ứng trung hòa:\nH^+^ + OH^-^ → H~2~O\nb) Tính toán số mol:\n- Số mol H^+^ ban đầu = n(HCl) = 0,1 . 0,15 = 0,015 (mol).\n- Số mol OH^-^ ban đầu = 2 . n(Ba(OH)~2~) = 2 . (0,1 . 0,10) = 0,020 (mol).\n- Vì n(OH^-^) > n(H^+^) nên sau phản ứng OH^-^ còn dư:\nn(OH^-^ dư) = 0,020 – 0,015 = 0,005 (mol).\n- Thể tích dung dịch sau trộn: V = 100 + 100 = 200 mL = 0,2 L.\n- Nồng độ ion OH^-^ trong dung dịch X: [OH^-^] = 0,005 / 0,2 = 0,025 (M).\n- Ta có: pOH = –log[OH^-^] = –log(0,025) ≈ 1,60.\n⇒ pH = 14 – pOH = 14 – 1,60 = 12,40."
        },
        {
            "num": 3, "topic": "Bài 5", "level": "TH", "points": "1,0 điểm",
            "stem": "Có ba mẫu phân đạm vô cơ dạng bột màu trắng được đựng riêng biệt trong ba lọ mất nhãn gồm: NaNO~3~, NH~4~Cl và (NH~4~)~2~SO~4~.\nBằng phương pháp hóa học, hãy trình bày cách sử dụng một thuốc thử duy nhất là dung dịch barium hydroxide (Ba(OH)~2~) để phân biệt ba mẫu phân đạm trên. Viết các phương trình hóa học xảy ra.",
            "solution": "Phương pháp nhận biết:\n- Lấy mỗi chất một lượng nhỏ vào 3 ống nghiệm riêng biệt, hòa tan vào nước cất để tạo thành dung dịch mẫu thử.\n- Nhỏ dung dịch Ba(OH)~2~ vào từng ống nghiệm rồi đun nóng nhẹ:\n  + Mẫu thử không có hiện tượng gì là NaNO~3~.\n  + Mẫu thử sinh ra chất khí có mùi khai, làm xanh giấy quỳ tím ẩm là NH~4~Cl:\n    Ba(OH)~2~ + 2NH~4~Cl → BaCl~2~ + 2NH~3~↑ + 2H~2~O\n  + Mẫu thử vừa xuất hiện kết tủa trắng (BaSO~4~), vừa sinh ra chất khí có mùi khai làm xanh quỳ tím ẩm là (NH~4~)~2~SO~4~:\n    Ba(OH)~2~ + (NH~4~)~2~SO~4~ → BaSO~4~↓ + 2NH~3~↑ + 2H~2~O"
        }
    ]
}

# ==============================================================================
# DỮ LIỆU ĐỀ 02 - MÃ ĐỀ 102
# ==============================================================================
DE02_DATA = {
    "exam_code": "102",
    "exam_num": "02",
    "p1": [
        {
            "num": 1, "topic": "Bài 1", "level": "NB",
            "stem": "Phản ứng một chiều là phản ứng",
            "opts": [
                "*A. chỉ xảy ra theo một chiều từ chất đầu biến thành chất sản phẩm.*",
                "B. xảy ra đồng thời theo hai chiều ngược nhau ở cùng điều kiện.",
                "C. có phương trình hóa học được biểu diễn bằng dấu hai chiều ⇌.",
                "D. luôn luôn xảy ra với tốc độ rất nhanh và tỏa nhiều nhiệt."
            ],
            "solution": "Phản ứng một chiều là phản ứng hóa học chỉ xảy ra theo một chiều nhất định từ chất tham gia tạo thành chất sản phẩm trong điều kiện xác định, phương trình biểu diễn bằng mũi tên một chiều (→)."
        },
        {
            "num": 2, "topic": "Bài 1", "level": "NB",
            "stem": "Trạng thái cân bằng hóa học là trạng thái của phản ứng thuận nghịch tại đó",
            "opts": [
                "A. nồng độ của các chất tham gia đã bằng 0.",
                "B. phản ứng hóa học đã hoàn toàn dừng lại.",
                "*C. tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch.*",
                "D. số mol của các chất tham gia bằng số mol của các sản phẩm."
            ],
            "solution": "Trạng thái cân bằng là trạng thái của phản ứng thuận nghịch khi tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch (v~t~ = v~n~). Ở trạng thái này, phản ứng vẫn xảy ra nhưng nồng độ các chất trong hệ không thay đổi."
        },
        {
            "num": 3, "topic": "Bài 1", "level": "NB",
            "stem": "Cho phản ứng tổng quát dạng: aA + bB ⇌ cC + dD (các chất đều ở thể khí). Biểu thức hằng số cân bằng K~C~ của phản ứng được xác định theo công thức nào?",
            "opts": [
                "A. K~C~ = ([A]^a . [B]^b) / ([C]^c . [D]^d).",
                "*B. K~C~ = ([C]^c . [D]^d) / ([A]^a . [B]^b).*",
                "C. K~C~ = (c.[C] + d.[D]) / (a.[A] + b.[B]).",
                "D. K~C~ = (a.[A] . b.[B]) / (c.[C] . d.[D])."
            ],
            "solution": "Theo định luật tác dụng khối lượng, hằng số cân bằng nồng độ K~C~ bằng tích số nồng độ cân bằng của các chất sản phẩm với lũy thừa là hệ số tỉ lượng tương ứng, chia cho tích số nồng độ cân bằng của các chất tham gia với lũy thừa tương ứng."
        },
        {
            "num": 4, "topic": "Bài 1", "level": "TH",
            "stem": "Cho cân bằng hóa học: CaCO~3~(s) ⇌ CaO(s) + CO~2~(g)   Δ~r~H^o^~298~ = +178 kJ. Biện pháp nào sau đây giúp cân bằng chuyển dịch theo chiều thuận tạo ra nhiều vôi sống CaO hơn?",
            "opts": [
                "*A. Tăng nhiệt độ nung lò.*",
                "B. Tăng áp suất chung của lò nung.",
                "C. Giảm nhiệt độ của hệ phản ứng.",
                "D. Tăng thêm lượng đá vôi rắn CaCO~3~."
            ],
            "solution": "Phản ứng nung đá vôi có ΔH > 0 là phản ứng thu nhiệt. Khi tăng nhiệt độ, cân bằng sẽ chuyển dịch theo chiều thu nhiệt (chiều thuận) để chống lại sự tăng nhiệt độ."
        },
        {
            "num": 5, "topic": "Bài 2", "level": "NB",
            "stem": "Chất nào sau đây là chất điện li yếu trong dung dịch nước?",
            "opts": [
                "A. HCl.",
                "B. NaOH.",
                "*C. CH~3~COOH.*",
                "D. KNO~3~."
            ],
            "solution": "CH~3~COOH (acetic acid) là một acid hữu cơ yếu, trong dung dịch nước chỉ có một phần nhỏ số phân tử phân li ra ion, phần lớn còn lại vẫn tồn tại dưới dạng phân tử phân tán trong nước."
        },
        {
            "num": 6, "topic": "Bài 2", "level": "TH",
            "stem": "Trong phản ứng: NH~3~ + H~2~O ⇌ NH~4~^+^ + OH^-^. Theo thuyết Brønsted – Lowry, phân tử NH~3~ đóng vai trò là",
            "opts": [
                "A. acid vì có khả năng nhường proton H^+^.",
                "*B. base vì có khả năng nhận proton H^+^.*",
                "C. chất khử vì số oxi hóa tăng sau phản ứng.",
                "D. chất lưỡng tính vì vừa nhận vừa nhường proton."
            ],
            "solution": "Trong phản ứng trên, phân tử NH~3~ đã nhận một proton H^+^ từ phân tử nước để biến thành ion NH~4~^+^. Theo thuyết Brønsted – Lowry, chất nhận proton là một base."
        },
        {
            "num": 7, "topic": "Bài 4", "level": "NB",
            "stem": "Trong công nghiệp thực phẩm, khí nitrogen thường được nạp vào các bao bì đóng gói bim bim, bánh kẹo nhằm mục đích",
            "opts": [
                "A. tạo vị ngọt thanh cho thực phẩm.",
                "*B. làm chậm quá trình oxi hóa và bảo quản thực phẩm.*",
                "C. làm tăng khối lượng tịnh của sản phẩm.",
                "D. giúp thực phẩm lên men nhanh chóng."
            ],
            "solution": "Khí nitrogen rất trơ về mặt hóa học, không màu, không mùi, không độc hại. Khi nạp khí N~2~ vào bao bì giúp đẩy khí oxygen ra ngoài, ngăn ngừa quá trình oxi hóa chất béo và chống ẩm mốc."
        },
        {
            "num": 8, "topic": "Bài 4", "level": "NB",
            "stem": "Số oxi hóa của nguyên tố nitrogen trong đơn chất N~2~ là bao nhiêu?",
            "opts": [
                "A. –3.",
                "*B. 0.*",
                "C. +2.",
                "D. +5."
            ],
            "solution": "Trong mọi đơn chất, số oxi hóa của nguyên tố luôn luôn bằng 0. Do đó trong đơn chất nitrogen (N~2~), số oxi hóa của nitrogen là 0."
        },
        {
            "num": 9, "topic": "Bài 5", "level": "NB",
            "stem": "Dung dịch ammonia (NH~3~) làm quỳ tím chuyển sang màu gì?",
            "opts": [
                "A. Màu đỏ.",
                "*B. Màu xanh.*",
                "C. Màu vàng.",
                "D. Không đổi màu."
            ],
            "solution": "Khi tan trong nước, ammonia phản ứng một phần với nước tạo ra các ion OH^-^ (NH~3~ + H~2~O ⇌ NH~4~^+^ + OH^-^). Dung dịch có tính base yếu nên làm quỳ tím hóa xanh."
        },
        {
            "num": 10, "topic": "Bài 5", "level": "TH",
            "stem": "Ammonia (NH~3~) thể hiện tính chất hóa học nào khi tác dụng với oxygen nung nóng không có chất xúc tác: 4NH~3~ + 3O~2~ → 2N~2~ + 6H~2~O?",
            "opts": [
                "A. Tính base.",
                "B. Tính acid.",
                "*C. Tính khử.*",
                "D. Tính oxi hóa."
            ],
            "solution": "Trong phản ứng này, số oxi hóa của nguyên tố nitrogen tăng từ –3 (trong NH~3~) lên 0 (trong N~2~). Chất nhường electron có số oxi hóa tăng là chất khử."
        },
        {
            "num": 11, "topic": "Bài 6", "level": "NB",
            "stem": "Hai chất khí chủ yếu gây ra hiện tượng mưa acid trong bầu khí quyển là",
            "opts": [
                "A. CO~2~ và CH~4~.",
                "*B. SO~2~ và các oxide của nitrogen (NO~x~).*",
                "C. O~2~ và N~2~.",
                "D. H~2~ và Cl~2~."
            ],
            "solution": "Khí SO~2~ (sinh ra từ đốt than đá, dầu mỏ) và các oxide của nitrogen NO~x~ (từ khí thải động cơ đốt trong) khi hòa tan vào nước mưa bị oxi hóa tạo thành H~2~SO~4~ và HNO~3~ gây hiện tượng mưa acid."
        },
        {
            "num": 12, "topic": "Bài 6", "level": "TH",
            "stem": "Dung dịch nitric acid (HNO~3~) đặc, nóng tác dụng được với kim loại nào sau đây sinh ra chất khí màu nâu đỏ?",
            "opts": [
                "A. Au.",
                "B. Pt.",
                "*C. Cu.*",
                "D. Al (ở trạng thái nguội)."
            ],
            "solution": "Dung dịch HNO~3~ đặc nóng là chất oxi hóa rất mạnh, hòa tan được hầu hết các kim loại (kể cả Cu, Ag; trừ Au, Pt): Cu + 4HNO~3~(đặc) → Cu(NO~3~)~2~ + 2NO~2~↑ + 2H~2~O. Khí NO~2~ có màu nâu đỏ."
        }
    ],
    "p2": [
        {
            "num": 1,
            "stem": "Khảo sát tính dẫn điện của dung dịch và các yếu tố ảnh hưởng đến cân bằng hóa học, hãy đánh giá tính Đúng hay Sai của các mệnh đề sau:",
            "items": [
                ("a", "Hydrochloric acid (HCl) và sodium hydroxide (NaOH) khi tan trong nước đều phân li hoàn toàn thành các ion tự do.", "Đúng"),
                ("b", "Nước cất tinh khiết hoàn toàn không phân li thành ion trong mọi điều kiện nhiệt độ.", "Sai"),
                ("c", "Chất xúc tác có tác dụng làm tăng tốc độ phản ứng thuận và tốc độ phản ứng nghịch với số lần như nhau, không làm chuyển dịch cân bằng.", "Đúng"),
                ("d", "Đối với hệ cân bằng: H~2~(g) + I~2~(g) ⇌ 2HI(g), khi tăng áp suất chung của hệ thì cân bằng sẽ chuyển dịch theo chiều nghịch.", "Sai")
            ],
            "solution": "a) Đúng. HCl là acid mạnh và NaOH là base mạnh, tan trong nước phân li hoàn toàn 100% thành các ion tự do.\nb) Sai. Nước là chất điện li vô cùng yếu, ở nhiệt độ phòng xảy ra quá trình tự ion hóa: 2H~2~O ⇌ H~3~O^+^ + OH^-^ với tích số ion K~w~ = 10^-14^.\nc) Đúng. Chất xúc tác làm giảm năng lượng hoạt hóa của cả phản ứng thuận và phản ứng nghịch như nhau nên giúp hệ nhanh đạt cân bằng chứ không làm thay đổi vị trí cân bằng.\nd) Sai. Phản ứng có tổng số mol khí ở vế trái (2 mol) bằng tổng số mol khí ở vế phải (2 mol), do đó áp suất không gây ra sự chuyển dịch cân bằng."
        },
        {
            "num": 2,
            "stem": "Khảo sát các phản ứng của ammonia, muối ammonium và các hợp chất của nitrogen, hãy đánh giá tính Đúng hay Sai của các phát biểu sau:",
            "items": [
                ("a", "Dung dịch muối ammonium tác dụng với dung dịch kiềm mạnh khi đun nóng giải phóng chất khí có mùi khai đặc trưng.", "Đúng"),
                ("b", "Nguyên tử nitrogen trong phân tử NH~3~ còn 1 cặp electron tự do nên ammonia có tính chất của một base yếu.", "Đúng"),
                ("c", "Nước mưa bình thường ở những vùng nông thôn không bị ô nhiễm có giá trị pH chính xác bằng 7,0.", "Sai"),
                ("d", "Dung dịch nitric acid (HNO~3~) có thể hòa tan được kim loại vàng (Au) và platinum (Pt) ở nhiệt độ phòng.", "Sai")
            ],
            "solution": "a) Đúng. Ion NH~4~^+^ phản ứng với OH^-^ khi đun nóng: NH~4~^+^ + OH^-^ → NH~3~↑ + H~2~O sinh ra khí NH~3~ mùi khai.\nb) Đúng. Khả năng nhận proton H^+^ nhờ cặp electron tự do trên nguyên tử N làm cho NH~3~ có tính base.\nc) Sai. Nước mưa bình thường không ô nhiễm đã hòa tan CO~2~ trong khí quyển tạo thành H~2~CO~3~ yếu, do đó có pH hơi acid nhẹ (khoảng 5,6).\nd) Sai. HNO~3~ dù đặc nóng cũng không hòa tan được Au và Pt. Chỉ có hỗn hợp nước cường toan (aqua regia gồm HCl đặc và HNO~3~ đặc theo tỉ lệ 3:1) mới hòa tan được vàng và bạch kim."
        }
    ],
    "p3": [
        {
            "num": 1, "topic": "Bài 4", "level": "NB",
            "stem": "Trong tự nhiên, nguyên tố nitrogen tồn tại chủ yếu ở dạng bao nhiêu đồng vị bền vững?",
            "answer": "2",
            "solution": "Trong tự nhiên, nitrogen tồn tại chủ yếu dưới dạng 2 đồng vị bền là ^14^N (chiếm khoảng 99,63%) và ^15^N (chiếm khoảng 0,37%)."
        },
        {
            "num": 2, "topic": "Bài 5", "level": "NB",
            "stem": "Tổng số nguyên tử có mặt trong một đơn vị công thức của muối ammonium nitrate (NH~4~NO~3~) là bao nhiêu?",
            "answer": "9",
            "solution": "Trong một đơn vị công thức NH~4~NO~3~ có: 2 nguyên tử N + 4 nguyên tử H + 3 nguyên tử O = tổng cộng 9 nguyên tử."
        },
        {
            "num": 3, "topic": "Bài 2", "level": "TH",
            "stem": "Một mẫu nước ép trái cây có nồng độ ion OH^-^ đo được ở 25 °C là 10^-11^ M. Giá trị pH của mẫu nước ép này bằng bao nhiêu?",
            "answer": "3",
            "solution": "Ở 25 °C: [H^+^].[OH^-^] = 10^-14^ ⇒ [H^+^] = 10^-14^ / 10^-11^ = 10^-3^ M ⇒ pH = –log[H^+^] = 3."
        },
        {
            "num": 4, "topic": "Bài 6", "level": "TH",
            "stem": "Số oxi hóa của nguyên tố nitrogen trong phân tử nitrogen dioxide (NO~2~) là bao nhiêu?",
            "answer": "+4",
            "solution": "Trong hợp chất NO~2~: nguyên tử O mang số oxi hóa –2. Gọi x là số oxi hóa của N: x + 2×(–2) = 0 ⇒ x = +4."
        }
    ],
    "p4": [
        {
            "num": 1, "topic": "Bài 1", "level": "VD", "points": "1,0 điểm",
            "stem": "Ở 430 °C, hằng số cân bằng K~C~ của phản ứng tổng hợp hydrogen iodide:\nH~2~(g) + I~2~(g) ⇌ 2HI(g) có giá trị bằng 54,0.\nNạp vào bình phản ứng kín dung tích 2,0 lít hỗn hợp gồm 0,40 mol khí H~2~ và 0,40 mol hơi I~2~, sau đó nung nóng bình ở 430 °C đến khi phản ứng đạt trạng thái cân bằng.\na) Xác định nồng độ mol ban đầu của H~2~ và I~2~ trong bình.\nb) Tính nồng độ mol của các chất H~2~, I~2~ và HI ở trạng thái cân bằng.",
            "solution": "a) Nồng độ ban đầu của các chất:\n[H~2~]~0~ = 0,40 / 2,0 = 0,20 (M).\n[I~2~]~0~ = 0,40 / 2,0 = 0,20 (M).\nb) Gọi x (mol/L) là nồng độ H~2~ đã tham gia phản ứng (0 < x < 0,20).\nPhương trình hóa học:   H~2~(g)  +  I~2~(g)   ⇌   2HI(g)\nBan đầu (M):            0,20      0,20            0\nPhản ứng (M):             x         x            2x\nCân bằng (M):         0,20 – x   0,20 – x        2x\nBiểu thức hằng số cân bằng:\nK~C~ = [HI]^2^ / ([H~2~] . [I~2~]) = (2x)^2^ / (0,20 – x)^2^ = 54,0\nLấy căn bậc hai hai vế:\n2x / (0,20 – x) = √54 ≈ 7,348\n⇒ 2x = 7,348 . (0,20 – x) = 1,470 – 7,348x\n⇒ 9,348x = 1,470 ⇒ x ≈ 0,157 (M).\nVậy tại thời điểm cân bằng:\n- [H~2~] = [I~2~] = 0,20 – 0,157 = 0,043 (M).\n- [HI] = 2 . 0,157 = 0,314 (M)."
        },
        {
            "num": 2, "topic": "Bài 2", "level": "VD", "points": "1,0 điểm",
            "stem": "Để xác định nồng độ mol của một mẫu dung dịch acetic acid (CH~3~COOH), người ta tiến hành phương pháp chuẩn độ bằng dung dịch chuẩn NaOH 0,10 M với chỉ thị phenolphthalein. Dùng burette nhỏ từ từ từng giọt dung dịch NaOH vào bình tam giác chứa 10,0 mL mẫu dung dịch acid. Khi dung dịch xuất hiện màu hồng nhạt bền trong khoảng 30 giây thì ghi nhận thể tích NaOH đã tiêu tốn trung bình qua 3 lần chuẩn độ là 12,5 mL.\na) Viết phương trình hóa học của phản ứng chuẩn độ.\nb) Tính nồng độ mol của acetic acid trong mẫu đem thử.",
            "solution": "a) Phương trình hóa học của phản ứng chuẩn độ:\nCH~3~COOH + NaOH → CH~3~COONa + H~2~O\nb) Tính toán nồng độ:\n- Thể tích dung dịch NaOH tiêu tốn: V(NaOH) = 12,5 mL = 0,0125 L.\n- Số mol NaOH đã phản ứng: n(NaOH) = 0,0125 . 0,10 = 0,00125 (mol).\n- Theo phương trình hóa học: n(CH~3~COOH) = n(NaOH) = 0,00125 (mol).\n- Nồng độ mol của dung dịch acetic acid mẫu:\nC~M~(CH~3~COOH) = n / V = 0,00125 / 0,010 = 0,125 (M)."
        },
        {
            "num": 3, "topic": "Bài 5", "level": "TH", "points": "1,0 điểm",
            "stem": "Dẫn từ từ dòng khí ammonia (NH~3~) đi qua ống sứ chịu nhiệt đựng bột copper(II) oxide (CuO) màu đen nung nóng đỏ.\na) Nêu hiện tượng quan sát được đối với chất rắn trong ống sứ và chất lỏng đọng lại ở thành ống sứ.\nb) Viết phương trình hóa học của phản ứng xảy ra và cho biết ammonia đóng vai trò là chất gì trong phản ứng này.",
            "solution": "a) Hiện tượng quan sát được:\n- Bột CuO màu đen trong ống sứ dần dần chuyển sang màu đỏ gạch của kim loại đồng (Cu).\n- Ở thành ống sứ nguội phía sau xuất hiện những giọt chất lỏng không màu ngưng tụ (nước H~2~O).\nb) Phương trình hóa học:\n2NH~3~ + 3CuO --(t°)-> N~2~ + 3Cu + 3H~2~O\nTrong phản ứng trên, số oxi hóa của nitrogen tăng từ –3 (trong NH~3~) lên 0 (trong N~2~), do đó ammonia đóng vai trò là chất khử (khử oxide kim loại thành kim loại tự do)."
        }
    ]
}

# ==============================================================================
# DỮ LIỆU ĐỀ 03 - MÃ ĐỀ 103
# ==============================================================================
DE03_DATA = {
    "exam_code": "103",
    "exam_num": "03",
    "p1": [
        {
            "num": 1, "topic": "Bài 1", "level": "NB",
            "stem": "Đặc điểm cơ bản của một phản ứng thuận nghịch là",
            "opts": [
                "*A. ở cùng điều kiện, phản ứng xảy ra đồng thời theo hai chiều ngược nhau.*",
                "B. phản ứng luôn xảy ra hoàn toàn tạo thành 100% chất sản phẩm.",
                "C. các chất phản ứng không thể tái tạo lại được sau khi phản ứng xảy ra.",
                "D. chỉ xảy ra khi có mặt chất xúc tác kim loại chuyển tiếp."
            ],
            "solution": "Phản ứng thuận nghịch có đặc điểm là trong cùng điều kiện xác định, phản ứng diễn ra đồng thời theo hai chiều ngược nhau: chiều thuận tạo sản phẩm và chiều nghịch tái tạo lại chất đầu."
        },
        {
            "num": 2, "topic": "Bài 1", "level": "NB",
            "stem": "Theo nguyên lí chuyển dịch cân bằng Le Chatelier, những yếu tố nào sau đây có thể làm chuyển dịch trạng thái cân bằng hóa học?",
            "opts": [
                "A. Nhiệt độ, nồng độ và chất xúc tác.",
                "*B. Nhiệt độ, nồng độ và áp suất (đối với hệ có chất khí).*",
                "C. Áp suất, diện tích tiếp xúc và chất xúc tác.",
                "D. Khối lượng chất rắn, thể tích bình và chất xúc tác."
            ],
            "solution": "Ba yếu tố ảnh hưởng trực tiếp đến trạng thái cân bằng hóa học là: nhiệt độ, nồng độ và áp suất (đối với hệ có chất khí tham gia). Chất xúc tác và diện tích bề mặt chất rắn không làm chuyển dịch cân bằng."
        },
        {
            "num": 3, "topic": "Bài 1", "level": "NB",
            "stem": "Khi một phản ứng thuận nghịch đạt đến trạng thái cân bằng thì khẳng định nào sau đây là đúng?",
            "opts": [
                "A. Phản ứng hóa học đã chấm dứt hoàn toàn.",
                "*B. Tốc độ của phản ứng thuận bằng tốc độ của phản ứng nghịch.*",
                "C. Toàn bộ lượng chất phản ứng ban đầu đã chuyển hết thành sản phẩm.",
                "D. Nồng độ của các chất sản phẩm luôn lớn hơn nồng độ các chất phản ứng."
            ],
            "solution": "Cân bằng hóa học là cân bằng động, phản ứng vẫn liên tục diễn ra nhưng do tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch nên nồng độ các chất trong hệ không còn biến đổi."
        },
        {
            "num": 4, "topic": "Bài 1", "level": "TH",
            "stem": "Cho cân bằng hóa học trong bình kín: 2NO~2~(g) [nâu đỏ] ⇌ N~2~O~4~(g) [không màu]   Δ~r~H^o^~298~ = –58 kJ. Khi ngâm bình phản ứng vào cốc nước đá lạnh (hạ nhiệt độ) thì",
            "opts": [
                "*A. màu nâu đỏ của hỗn hợp khí trong bình sẽ nhạt dần.*",
                "B. màu nâu đỏ của hỗn hợp khí trong bình sẽ đậm dần lên.",
                "C. màu sắc của hỗn hợp khí trong bình hoàn toàn không thay đổi.",
                "D. trong bình xuất hiện kết tủa trắng bám vào thành bình."
            ],
            "solution": "Phản ứng có ΔH < 0 là phản ứng tỏa nhiệt theo chiều thuận. Khi ngâm vào nước đá (hạ nhiệt độ), cân bằng chuyển dịch theo chiều tỏa nhiệt để chống lại việc làm lạnh, tức là chuyển dịch theo chiều thuận tạo thêm N~2~O~4~ không màu làm màu nâu đỏ của khí nhạt dần."
        },
        {
            "num": 5, "topic": "Bài 2", "level": "NB",
            "stem": "Dung dịch chất nào sau đây không có khả năng dẫn điện?",
            "opts": [
                "A. Dung dịch sulfuric acid (H~2~SO~4~).",
                "B. Dung dịch potassium hydroxide (KOH).",
                "C. Dung dịch sodium chloride (NaCl).",
                "*D. Dung dịch đường saccharose (C~12~H~22~O~11~).*"
            ],
            "solution": "Dung dịch đường saccharose khi tan trong nước chỉ phân tán thành các phân tử trung hòa C~12~H~22~O~11~, không phân li ra các ion mang điện tích tự do nên không dẫn được điện."
        },
        {
            "num": 6, "topic": "Bài 2", "level": "TH",
            "stem": "Thang đo pH thường dùng có giá trị từ 1 đến 14. Dung dịch nào sau đây có môi trường kiềm (base)?",
            "opts": [
                "A. Dung dịch có pH = 2,5.",
                "B. Dung dịch có pH = 7,0.",
                "*C. Dung dịch có pH = 11,2.*",
                "D. Dung dịch có [H^+^] = 10^-5^ M."
            ],
            "solution": "Ở 25 °C: Dung dịch trung tính có pH = 7; dung dịch acid có pH < 7; dung dịch kiềm (base) có pH > 7. Với pH = 11,2 > 7 thì dung dịch có môi trường kiềm."
        },
        {
            "num": 7, "topic": "Bài 4", "level": "NB",
            "stem": "Trong lĩnh vực y tế và công nghệ sinh học, nitrogen lỏng (ở nhiệt độ khoảng –196 °C) thường được ứng dụng để",
            "opts": [
                "A. khử trùng nguồn nước uống sinh hoạt.",
                "*B. bảo quản mẫu tế bào, mô, trứng và tinh trùng sinh sản.*",
                "C. làm nhiên liệu thay thế cho cồn y tế.",
                "D. điều trị các bệnh về tiêu hóa."
            ],
            "solution": "Do có nhiệt độ sôi cực thấp (–196 °C) và không độc hại, nitrogen lỏng được dùng làm chất làm lạnh sâu để bảo quản lâu dài mẫu máu, tinh trùng, phôi ghép và các mô sinh học."
        },
        {
            "num": 8, "topic": "Bài 4", "level": "NB",
            "stem": "Khẳng định nào sau đây là đúng về tính tan của khí nitrogen (N~2~) trong nước?",
            "opts": [
                "A. Khí nitrogen tan vô hạn trong nước ở mọi nhiệt độ.",
                "*B. Khí nitrogen tan rất ít trong nước ở điều kiện thường.*",
                "C. Khí nitrogen tan rất nhiều tạo dung dịch acid mạnh.",
                "D. Khí nitrogen phản ứng mãnh liệt với nước tạo bọt khí."
            ],
            "solution": "Ở điều kiện chuẩn 20 °C, 1 lít nước chỉ hòa tan được khoảng 15 mL khí nitrogen. Phân tử N~2~ không phân cực nên rất khó tan trong dung môi phân cực như nước."
        },
        {
            "num": 9, "topic": "Bài 5", "level": "NB",
            "stem": "Muối ammonium là nhóm hợp chất ion có chứa cation nào sau đây?",
            "opts": [
                "A. NO~3~^-^.",
                "*B. NH~4~^+^.*",
                "C. Na^+^.",
                "D. NO~2~^-^."
            ],
            "solution": "Muối ammonium là muối tạo bởi cation ammonium (NH~4~^+^) và anion gốc acid (như Cl^-^, SO~4~^2-^, NO~3~^-^, CO~3~^2-^...)."
        },
        {
            "num": 10, "topic": "Bài 5", "level": "TH",
            "stem": "Để nhận biết ion ammonium (NH~4~^+^) trong dung dịch mẫu thử, người ta cho dung dịch mẫu tác dụng với hóa chất nào và quan sát hiện tượng gì?",
            "opts": [
                "A. Cho tác dụng với dung dịch AgNO~3~ xuất hiện kết tủa trắng.",
                "*B. Cho tác dụng với dung dịch kiềm nóng sinh ra khí có mùi khai làm xanh quỳ tím ẩm.*",
                "C. Cho tác dụng với dung dịch BaCl~2~ xuất hiện kết tủa trắng không tan trong acid.",
                "D. Cho tác dụng với dung dịch HCl sinh ra bọt khí không màu, không mùi."
            ],
            "solution": "Ion NH~4~^+^ tác dụng với OH^-^ sinh ra khí ammonia (NH~3~) có mùi khai đặc trưng và làm đổi màu giấy quỳ tím ẩm sang màu xanh: NH~4~^+^ + OH^-^ --(t°)-> NH~3~↑ + H~2~O."
        },
        {
            "num": 11, "topic": "Bài 6", "level": "NB",
            "stem": "Nước mưa được gọi là 'mưa acid' khi giá trị pH đo được nhỏ hơn",
            "opts": [
                "A. 7,0.",
                "*B. 5,6.*",
                "C. 4,0.",
                "D. 8,2."
            ],
            "solution": "Nước mưa bình thường hòa tan CO~2~ có pH ≈ 5,6. Khi nước mưa có pH < 5,6 do bị nhiễm các acid mạnh (H~2~SO~4~, HNO~3~) thì được gọi là hiện tượng mưa acid."
        },
        {
            "num": 12, "topic": "Bài 6", "level": "TH",
            "stem": "Kim loại copper (Cu) tan được trong dung dịch nitric acid loãng (HNO~3~) giải phóng chất khí không màu, hóa nâu ngoài không khí. Khí đó là",
            "opts": [
                "A. H~2~.",
                "B. NH~3~.",
                "*C. NO.*",
                "D. N~2~O."
            ],
            "solution": "Phản ứng giữa Cu và HNO~3~ loãng: 3Cu + 8HNO~3~(loãng) → 3Cu(NO~3~)~2~ + 2NO↑ + 4H~2~O. Khí sinh ra là nitrogen monoxide (NO) không màu, khi gặp oxygen trong không khí lập tức kết hợp tạo NO~2~ có màu nâu đỏ."
        }
    ],
    "p2": [
        {
            "num": 1,
            "stem": "Khảo sát tính chất của các chất điện li và cân bằng hóa học trong pha khí, hãy đánh giá tính Đúng hay Sai của các mệnh đề sau:",
            "items": [
                ("a", "Chất điện li yếu khi tan trong nước chỉ có một phần nhỏ số phân tử phân li ra ion, phần còn lại tồn tại ở dạng phân tử trung hòa.", "Đúng"),
                ("b", "Muối ăn (NaCl) ở trạng thái tinh thể rắn khan dẫn điện rất tốt vì trong mạng tinh thể có chứa nhiều ion Na^+^ và Cl^-^.", "Sai"),
                ("c", "Khi tăng nồng độ của một chất tham gia trong phản ứng thuận nghịch, cân bằng sẽ chuyển dịch theo chiều làm giảm nồng độ chất đó (chiều thuận).", "Đúng"),
                ("d", "Đối với cân bằng: N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g), việc tăng thể tích của bình phản ứng lên 2 lần ở nhiệt độ không đổi sẽ làm cân bằng chuyển dịch theo chiều nghịch.", "Đúng")
            ],
            "solution": "a) Đúng. Đó là định nghĩa của chất điện li yếu (như CH~3~COOH, HF, H~2~O...).\nb) Sai. Ở trạng thái tinh thể rắn khan, các ion Na^+^ và Cl^-^ bị liên kết chặt chẽ trong mạng tinh thể ion, không chuyển động tự do được nên NaCl rắn khan không dẫn điện.\nc) Đúng. Theo nguyên lí Le Chatelier, khi tăng nồng độ một chất, hệ sẽ tự cân bằng lại bằng cách chuyển dịch theo chiều tiêu thụ bớt chất đó.\nd) Đúng. Tăng thể tích bình làm giảm áp suất chung của hệ khí. Cân bằng sẽ chuyển dịch theo chiều làm tăng số mol khí, tức là chiều nghịch (từ 2 mol biến thành 4 mol)."
        },
        {
            "num": 2,
            "stem": "Khảo sát tính chất vật lí, hóa học và ứng dụng thực tiễn của ammonia và các hợp chất nitrogen, hãy đánh giá tính Đúng hay Sai của các phát biểu sau:",
            "items": [
                ("a", "Khí ammonia có khối lượng mol phân tử là 17 g/mol nên nhẹ hơn không khí (d ≈ 17/29 ≈ 0,59).", "Đúng"),
                ("b", "Không nên bón chung phân đạm ammonium với vôi sống (CaO) vì sẽ gây phản ứng hóa học làm thất thoát dinh dưỡng đạm dưới dạng khí NH~3~.", "Đúng"),
                ("c", "Acid nitric (HNO~3~) là một nguyên liệu công nghiệp quan trọng, được sử dụng phổ biến để sản xuất phân bón đạm nitrate và chất nổ.", "Đúng"),
                ("d", "Sự phóng điện trong khí quyển do sấm sét tạo ra khí NO là nguyên nhân chính gây ngộ độc nguồn nước ngầm của Trái Đất.", "Sai")
            ],
            "solution": "a) Đúng. M(NH~3~) = 17 g/mol < M(không khí) = 29 g/mol nên khí ammonia nhẹ hơn không khí và bay lên cao.\nb) Đúng. Vôi sống gặp nước tạo Ca(OH)~2~ có môi trường kiềm mạnh: Ca(OH)~2~ + 2NH~4~^+^ → Ca^2+^ + 2NH~3~↑ + 2H~2~O làm đạm bay mất vào không khí.\nc) Đúng. Hơn 80% sản lượng HNO~3~ được dùng sản xuất phân bón (NH~4~NO~3~, KNO~3~...), phần còn lại dùng sản xuất thuốc nổ (TNT, nitroglycerin) và phẩm nhuộm.\nd) Sai. Quá trình phóng điện tạo NO trong sấm sét là một quy trình tự nhiên vô cùng quý giá, cung cấp hàng triệu tấn đạm sinh học tự nhiên cho đất trồng hàng năm."
        }
    ],
    "p3": [
        {
            "num": 1, "topic": "Bài 4", "level": "NB",
            "stem": "Hóa trị của nguyên tố nitrogen trong phân tử đơn chất nitrogen (N~2~) là bao nhiêu?",
            "answer": "3",
            "solution": "Trong phân tử N~2~, hai nguyên tử nitrogen dùng chung 3 cặp electron tạo liên kết ba N≡N, do đó hóa trị của nitrogen trong N~2~ là III (hoặc 3)."
        },
        {
            "num": 2, "topic": "Bài 5", "level": "NB",
            "stem": "Số nguyên tử hydrogen liên kết trực tiếp với nguyên tử nitrogen trong một cation ammonium (NH~4~^+^) là bao nhiêu?",
            "answer": "4",
            "solution": "Trong cation ammonium (NH~4~^+^), nguyên tử nitrogen trung tâm liên kết với 4 nguyên tử hydrogen (gồm 3 liên kết cộng hóa trị thông thường và 1 liên kết cho - nhận)."
        },
        {
            "num": 3, "topic": "Bài 2", "level": "TH",
            "stem": "Trộn 50 mL dung dịch hydrochloric acid (HCl) 0,20 M với 50 mL dung dịch sodium hydroxide (NaOH) 0,20 M ở 25 °C. Giá trị pH của dung dịch thu được sau phản ứng bằng bao nhiêu?",
            "answer": "7",
            "solution": "Số mol H^+^ = 0,05 . 0,20 = 0,01 mol; số mol OH^-^ = 0,05 . 0,20 = 0,01 mol. Hai chất phản ứng vừa đủ với nhau tạo thành muối trung tính NaCl và nước, do đó dung dịch sau phản ứng có môi trường trung tính với pH = 7."
        },
        {
            "num": 4, "topic": "Bài 6", "level": "TH",
            "stem": "Cho sơ đồ các chuyển hóa sau của nguyên tố nitrogen:\n(1) N~2~ → NH~3~\n(2) NH~3~ → NO\n(3) NO → NO~2~\n(4) NO~2~ → HNO~3~\nTrong số các phản ứng trên, có bao nhiêu phản ứng mà số oxi hóa của nguyên tố nitrogen tăng lên?",
            "answer": "3",
            "solution": "Xét số oxi hóa của N qua từng phản ứng:\n(1) N~2~(0) → NH~3~(–3): giảm từ 0 xuống –3.\n(2) NH~3~(–3) → NO(+2): tăng từ –3 lên +2.\n(3) NO(+2) → NO~2~(+4): tăng từ +2 lên +4.\n(4) NO~2~(+4) → HNO~3~(+5): tăng từ +4 lên +5.\nVậy có 3 phản ứng mà số oxi hóa của nitrogen tăng lên là (2), (3), (4)."
        }
    ],
    "p4": [
        {
            "num": 1, "topic": "Bài 1", "level": "VD", "points": "1,0 điểm",
            "stem": "Phản ứng nhiệt phân phosphorus pentachloride trong bình kín dung tích không đổi 1,0 lít ở 250 °C xảy ra theo phương trình hóa học:\nPCl~5~(g) ⇌ PCl~3~(g) + Cl~2~(g)\nBan đầu nạp vào bình 0,20 mol khí PCl~5~ rồi đun nóng đến 250 °C. Khi phản ứng đạt trạng thái cân bằng, người ta đo được nồng độ của khí chlorine (Cl~2~) sinh ra trong bình là 0,08 mol/L.\na) Xác định nồng độ mol của PCl~5~ và PCl~3~ ở trạng thái cân bằng.\nb) Tính giá trị hằng số cân bằng K~C~ của phản ứng ở 250 °C.",
            "solution": "a) Lập bảng biến thiên nồng độ (dung tích bình V = 1,0 L nên nồng độ mol bằng số mol):\nPhương trình hóa học:   PCl~5~(g)  ⇌  PCl~3~(g)  +  Cl~2~(g)\nBan đầu (M):             0,20           0            0\nPhản ứng (M):            0,08          0,08         0,08\nCân bằng (M):            0,12          0,08         0,08\n- Nồng độ của PCl~5~ ở trạng thái cân bằng: [PCl~5~] = 0,20 – 0,08 = 0,12 (M).\n- Nồng độ của PCl~3~ ở trạng thái cân bằng: [PCl~3~] = 0,08 (M).\nb) Giá trị hằng số cân bằng K~C~ ở 250 °C:\nK~C~ = ([PCl~3~] . [Cl~2~]) / [PCl~5~] = (0,08 . 0,08) / 0,12 ≈ 0,0533."
        },
        {
            "num": 2, "topic": "Bài 2", "level": "VD", "points": "1,0 điểm",
            "stem": "Trộn 100 mL dung dịch sulfuric acid (H~2~SO~4~) nồng độ 0,05 M với 100 mL dung dịch barium hydroxide (Ba(OH)~2~) nồng độ 0,10 M thu được kết tủa trắng và dung dịch Y.\na) Tính khối lượng kết tủa trắng thu được sau khi phản ứng xảy ra hoàn toàn (Biết Ba = 137, S = 32, O = 16).\nb) Tính nồng độ mol của ion OH^-^ dư và giá trị pH của dung dịch Y ở 25 °C (coi thể tích dung dịch Y bằng 200 mL).",
            "solution": "a) Tính toán số mol các chất ban đầu:\n- n(H~2~SO~4~) = 0,1 . 0,05 = 0,005 (mol) ⇒ n(H^+^) = 0,010 mol; n(SO~4~^2-^) = 0,005 mol.\n- n(Ba(OH)~2~) = 0,1 . 0,10 = 0,010 (mol) ⇒ n(Ba^2+^) = 0,010 mol; n(OH^-^) = 0,020 mol.\nPhản ứng tạo kết tủa: Ba^2+^ + SO~4~^2-^ → BaSO~4~↓\nVì n(Ba^2+^) = 0,010 > n(SO~4~^2-^) = 0,005 nên SO~4~^2-^ phản ứng hết, kết tủa tính theo SO~4~^2-^:\nn(BaSO~4~) = 0,005 (mol).\nKhối lượng kết tủa: m(BaSO~4~) = 0,005 . 233 = 1,165 (gam).\nb) Phản ứng trung hòa: H^+^ + OH^-^ → H~2~O\n- Số mol OH^-^ đã phản ứng = n(H^+^) = 0,010 (mol).\n- Số mol OH^-^ còn dư trong dung dịch Y: n(OH^-^ dư) = 0,020 – 0,010 = 0,010 (mol).\n- Nồng độ mol của ion OH^-^ trong dung dịch Y (V = 200 mL = 0,2 L):\n[OH^-^] = 0,010 / 0,2 = 0,05 (M).\n- Tính pH:\npOH = –log[OH^-^] = –log(0,05) ≈ 1,30.\n⇒ pH = 14 – pOH = 14 – 1,30 = 12,70."
        },
        {
            "num": 3, "topic": "Bài 5", "level": "TH", "points": "1,0 điểm",
            "stem": "Trong thí nghiệm đài phun nước của ammonia: Người ta nạp đầy khí ammonia (NH~3~) khô vào một bình cầu thủy tinh rồi đậy kín bằng nút cao su có cắm ống thủy tinh vuốt nhọn xuyên qua. Nhúng đầu dưới của ống thủy tinh vào chậu nước có pha sẵn vài giọt chỉ thị phenolphthalein. Sau khi bóp một lượng nhỏ nước vào bình cầu để khởi động thí nghiệm:\na) Nêu hiện tượng quan sát được trong bình cầu.\nb) Giải thích nguyên nhân xuất hiện hiện tượng trên.",
            "solution": "a) Hiện tượng quan sát được:\n- Nước từ trong chậu thủy tinh phun mạnh lên thành các tia nước màu hồng bắn tung tóe trong bình cầu giống như một đài phun nước rực rỡ.\nb) Giải thích nguyên nhân:\n- Khí ammonia tan cực kì nhiều trong nước (1 lít nước hòa tan khoảng 700 lít NH~3~). Khi một lượng nhỏ nước vào bình cầu, lượng lớn khí NH~3~ lập tức bị hòa tan làm áp suất trong bình giảm đột ngột so với áp suất khí quyển bên ngoài.\n- Áp suất khí quyển ép nước từ chậu phun mạnh qua ống vuốt nhọn vào bình cầu.\n- Dung dịch ammonia tạo ra có môi trường kiềm (NH~3~ + H~2~O ⇌ NH~4~^+^ + OH^-^) làm chỉ thị phenolphthalein lập tức chuyển sang màu hồng đặc trưng."
        }
    ]
}

# ==============================================================================
# HÀM XUẤT FILE WORD CHO MỘT ĐỀ THI
# ==============================================================================
def export_single_exam_word(exam_data, is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    code = exam_data["exam_code"]
    exam_num = exam_data["exam_num"]
    title_text = f"ĐỀ KIỂM TRA GIỮA HỌC KÌ I - ĐỀ SỐ {exam_num}"

    # Tiêu đề đề thi
    build_title_block(
        doc,
        school_line="SỞ GIÁO DỤC VÀ ĐÀO TẠO THỪA THIÊN HUẾ",
        exam_title=title_text,
        subject_line="MÔN: HOÁ HỌC - LỚP 11",
        num_pages="04",
        duration="45 phút",
        exam_code=code,
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Ba = 137",
        is_teacher=is_teacher
    )

    q_global_num = 1

    # --------------------------------------------------------------------------
    # PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (12 CÂU - 3,0 ĐIỂM)
    # --------------------------------------------------------------------------
    add_section_header(
        doc,
        "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (Thí sinh trả lời từ câu 1 đến câu 12)",
        "Mỗi câu hỏi thí sinh chỉ chọn một phương án chính xác nhất. Mỗi câu trả lời đúng được 0,25 điểm."
    )

    for q in exam_data["p1"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, f"Câu {q_global_num}. {q['stem']}")

        render_options_no_table(doc, q["opts"], is_teacher=is_teacher)

        if is_teacher and q.get("solution"):
            add_solution_no_table(doc, q["solution"])

        q_global_num += 1

    # --------------------------------------------------------------------------
    # PHẦN II: TRẮC NGHIỆM ĐÚNG - SAI (2 CÂU - 2,0 ĐIỂM)
    # --------------------------------------------------------------------------
    add_section_header(
        doc,
        "PHẦN II. Câu trắc nghiệm đúng sai (Thí sinh trả lời từ câu 13 đến câu 14)",
        "Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn Đúng hoặc Sai. Điểm tối đa cho 1 câu là 1,0 điểm."
    )

    for q in exam_data["p2"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, f"Câu {q_global_num}. {q['stem']}")

        for item_lbl, item_txt, ans_val in q["items"]:
            p_it = doc.add_paragraph()
            p_it.paragraph_format.left_indent = Cm(0.6)
            p_it.paragraph_format.space_before = Pt(1.5)
            p_it.paragraph_format.space_after = Pt(1.5)
            render_statement(p_it, f"{item_lbl}) {item_txt}", ans=(ans_val if is_teacher else None), is_teacher=is_teacher)

        if is_teacher and q.get("solution"):
            add_solution_no_table(doc, q["solution"])

        q_global_num += 1

    # --------------------------------------------------------------------------
    # PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (4 CÂU - 2,0 ĐIỂM)
    # --------------------------------------------------------------------------
    add_section_header(
        doc,
        "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn (Thí sinh trả lời từ câu 15 đến câu 18)",
        "Thí sinh viết câu trả lời hoặc số liệu tương ứng vào bài làm. Mỗi câu trả lời đúng được 0,5 điểm."
    )

    for q in exam_data["p3"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, f"Câu {q_global_num}. {q['stem']}")

        if not is_teacher:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.6)
            p_ans.paragraph_format.space_before = Pt(2)
            p_ans.paragraph_format.space_after = Pt(3)
            r_ans = p_ans.add_run("Đáp số / Trả lời: .......................................................................................")
            r_ans.font.name = "Times New Roman"
            r_ans.font.size = Pt(10.5)
            r_ans.italic = True
        else:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.6)
            p_ans.paragraph_format.space_before = Pt(2)
            p_ans.paragraph_format.space_after = Pt(2)
            render_rich_text(p_ans, f"► *Đáp số chuẩn:* *{q['answer']}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
            if q.get("solution"):
                add_solution_no_table(doc, q["solution"])

        q_global_num += 1

    # --------------------------------------------------------------------------
    # PHẦN IV: TỰ LUẬN (3 CÂU - 3,0 ĐIỂM)
    # --------------------------------------------------------------------------
    add_section_header(
        doc,
        "PHẦN IV. Tự luận (Gồm 3 câu: Câu 19, Câu 20, Câu 21 - Tổng điểm: 3,0 điểm)",
        "Thí sinh trình bày chi tiết các bước lập luận, tính toán và viết phương trình hóa học vào bài thi."
    )

    for q in exam_data["p4"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, f"Câu {q_global_num} ({q['points']}). {q['stem']}")

        if not is_teacher:
            p_lb = doc.add_paragraph()
            p_lb.paragraph_format.left_indent = Cm(0.4)
            p_lb.paragraph_format.space_before = Pt(2)
            p_lb.paragraph_format.space_after = Pt(1)
            r_lb = p_lb.add_run("Bài làm:")
            r_lb.font.name = "Times New Roman"
            r_lb.font.size = Pt(10.5)
            r_lb.italic = True
            add_student_writing_lines(doc, num_lines=7)
        else:
            if q.get("solution"):
                add_solution_no_table(doc, q["solution"])

        q_global_num += 1

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(4)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    r_end.font.name = "Times New Roman"
    r_end.font.size = Pt(11)
    r_end.bold = True

    role_suffix = "GiaoVien" if is_teacher else "HocSinh"
    out_name = f"De_Kiem_Tra_GK1_Hoa_11_De{exam_num}_Ma{code}_{role_suffix}.docx"
    out_path = os.path.join(OUT_DIR, out_name)
    doc.save(out_path)
    print(f"  -> Đã tạo thành công: {out_name}")
    return out_path


def main():
    print("=" * 70)
    print("TIẾN HÀNH XUẤT 3 ĐỀ KIỂM TRA HÓA HỌC 11 THEO MA TRẬN CHUẨN")
    print("=" * 70)

    exams = [DE01_DATA, DE02_DATA, DE03_DATA]
    created_files = []

    for exam in exams:
        code = exam["exam_code"]
        num = exam["exam_num"]
        print(f"\n[+] Đang xử lý Đề số {num} (Mã đề {code})...")
        f_hs = export_single_exam_word(exam, is_teacher=False)
        f_gv = export_single_exam_word(exam, is_teacher=True)
        created_files.extend([f_hs, f_gv])

    print("\n" + "=" * 70)
    print(f"HOÀN THÀNH TOÀN BỘ 6 FILE WORD ({len(created_files)} files) TẠI THƯ MỤC:")
    print(f"{OUT_DIR}")
    for f in created_files:
        print(f"  * {os.path.basename(f)}")
    print("=" * 70)


if __name__ == "__main__":
    main()
