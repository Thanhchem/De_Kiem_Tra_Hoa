# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Đánh Giá Cuối Kì I - Hóa học 11 - Trường THPT Bùi Thị Xuân (TP Huế)
Mã đề: 1102 - Năm học 2025 - 2026
Chuẩn theo Quy chuẩn Form Đề Kiểm Tra của Thầy TRẦN VĂN THẠNH
Tạo 2 phiên bản: Bản Học Sinh & Bản Giáo Viên (có đáp án đậm đỏ & lời giải chi tiết)
"""

import os
import sys
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

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
    COLOR_CORRECT,
    COLOR_SOLUTION
)

DIR_SAN_PHAM = os.path.abspath("San_Pham")
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_ck1_hoa11_buithixuan")
os.makedirs(DIR_SAN_PHAM, exist_ok=True)

# Dữ liệu Đề thi mã 1102
DATA_EXAM = {
    "school_line": "SỞ GD&ĐT THÀNH PHỐ HUẾ - TRƯỜNG THPT BÙI THỊ XUÂN",
    "exam_title": "KIỂM TRA ĐÁNH GIÁ CUỐI KÌ I - NĂM HỌC 2025 - 2026",
    "subject_line": "MÔN: HOÁ HỌC LỚP 11",
    "duration": "45 phút (không kể thời gian phát đề)",
    "num_pages": "03",
    "exam_code": "1102",
    "atomic_masses": "H = 1; O = 16; C = 12; N = 14; Na = 23; Mg = 24; K = 39; S = 32",

    # A. PHẦN TRẮC NGHIỆM (7đ)
    # PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. (12 câu - 3,0 điểm)
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1. Cho các phát biểu sau:\n(1) Sulfur vừa có tính oxi hóa vừa có tính khử.\n(2) Trong tự nhiên, sulfur chỉ tồn tại ở dạng hợp chất.\n(3) SO~2~ là acidic oxide, tan trong nước tạo dung dịch có tính acid.\n(4) H~2~SO~4~ đặc có tính háo nước mạnh và tính oxi hóa mạnh.\n(5) Trong hợp chất, nguyên tố sulfur có số oxi hóa phổ biến là -2; +4; +6.\nSố phát biểu đúng là",
            "options": [
                "A. 2.",
                "B. 3.",
                "*C. 4.*",
                "D. 5."
            ],
            "solution": "• (1) Đúng. Sulfur đơn chất có số oxi hóa 0 trung gian, có thể giảm xuống -2 (thể hiện tính oxi hóa) hoặc tăng lên +4, +6 (thể hiện tính khử).\n• (2) Sai. Trong tự nhiên, sulfur tồn tại ở cả dạng đơn chất (các mỏ lưu huỳnh tự do) và dạng hợp chất (quặng sulfide, quặng sulfate).\n• (3) Đúng. SO~2~ là acidic oxide, tan trong nước tạo dung dịch sulfurous acid (H~2~SO~3~) có tính acid yếu: SO~2~ + H~2~O ⇌ H~2~SO~3~.\n• (4) Đúng. H~2~SO~4~ đặc có tính háo nước rất mạnh và tính oxi hóa mạnh (oxi hóa được hầu hết kim loại trừ Au, Pt và nhiều phi kim, hợp chất khử).\n• (5) Đúng. Trong các hợp chất hóa học, số oxi hóa phổ biến nhất của nguyên tố sulfur là -2 (trong sulfide), +4 (trong sulfite, SO~2~) và +6 (trong sulfate, SO~3~, H~2~SO~4~).\n⇒ Có 4 phát biểu đúng: (1), (3), (4), (5).\n*Đáp án đúng: C.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Dung dịch H~2~SO~4~ loãng có thể tác dụng với cả hai chất nào sau đây?",
            "options": [
                "*A. Fe và Fe(OH)~3~.*",
                "B. S và H~2~S.",
                "C. C và CO~2~.",
                "D. Cu và Cu(OH)~2~."
            ],
            "solution": "- Dung dịch H~2~SO~4~ loãng có đầy đủ tính chất của một acid mạnh:\n  + Tác dụng với kim loại đứng trước hydrogen trong dãy hoạt động hóa học giải phóng H~2~ (tác dụng với Fe; không tác dụng với Cu).\n  + Tác dụng với base, basic oxide tạo muối và nước (tác dụng với Fe(OH)~3~, Cu(OH)~2~).\n- Phản ứng:\n  Fe + H~2~SO~4~ → FeSO~4~ + H~2~↑\n  2Fe(OH)~3~ + 3H~2~SO~4~ → Fe~2~(SO~4~)~3~ + 6H~2~O\n- H~2~SO~4~ loãng không tác dụng với S, H~2~S, C, CO~2~, Cu.\n*Đáp án đúng: A.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Cặp chất nào sau đây có thể là đồng đẳng của nhau?",
            "options": [
                "A. C~2~H~4~, C~3~H~8~.",
                "B. CH~4~O, C~2~H~4~O.",
                "C. CH~4~, C~2~H~4~.",
                "*D. C~2~H~2~, C~4~H~6~.*"
            ],
            "solution": "Đồng đẳng là hiện tượng các chất có thành phần phân tử hơn kém nhau một hay nhiều nhóm -CH~2~- nhưng có cấu tạo hóa học và tính chất hóa học tương tự nhau (cùng dãy đồng đẳng).\n- A: C~2~H~4~ (alkene) và C~3~H~8~ (alkane) thuộc 2 dãy khác nhau.\n- B: CH~4~O (CH~3~OH, alcohol no) và C~2~H~4~O (aldehyde hoặc hợp chất không no) khác loại nhóm chức.\n- C: CH~4~ (alkane) và C~2~H~4~ (alkene) thuộc 2 dãy khác nhau.\n- D: C~2~H~2~ và C~4~H~6~ cùng có công thức chung C~n~H~2n-2~, hơn kém nhau 2 nhóm -CH~2~-, có thể cùng thuộc dãy đồng đẳng alkyne (CH≡CH và CH≡C-CH~2~-CH~3~ hoặc CH~3~-C≡C-CH~3~).\n*Đáp án đúng: D.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Để tách tinh dầu sả (có trong thân, lá, rễ .... cây sả) trong công nghiệp hương liệu, người ta dùng phương pháp",
            "options": [
                "A. kết tinh dầu sả trong nước.",
                "*B. chưng cất bằng hơi nước và chiết tinh dầu ra khỏi hỗn hợp sản phẩm.*",
                "C. chưng cất bằng hơi nước và chiết bằng nước lạnh.",
                "D. chiết tinh dầu sả sau đó chưng cất bằng hơi nước."
            ],
            "solution": "Tinh dầu sả (chứa geraniol, citronellal,...) nhẹ hơn nước, thực tế không tan trong nước và dễ bay hơi cùng với hơi nước sôi. Do đó, người ta sử dụng phương pháp chưng cất lôi cuốn hơi nước để đưa hỗn hợp tinh dầu và hơi nước sang bộ ngưng tụ. Hỗn hợp ngưng tụ phân thành 2 lớp lỏng (lớp tinh dầu nhẹ nổi lên trên), sau đó dùng phễu chiết để tách riêng lớp tinh dầu ra khỏi nước.\n*Đáp án đúng: B.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5. Yếu tố không làm ảnh hưởng đến chuyển dịch cân bằng là",
            "options": [
                "A. nồng độ.",
                "B. nhiệt độ.",
                "C. áp suất.",
                "*D. chất xúc tác.*"
            ],
            "solution": "Chất xúc tác làm tăng tốc độ của cả phản ứng thuận và phản ứng nghịch với cùng một mức độ như nhau, giúp hệ nhanh chóng đạt tới trạng thái cân bằng chứ không làm thay đổi nồng độ các chất ở trạng thái cân bằng, do đó chất xúc tác không làm chuyển dịch cân bằng hóa học.\n*Đáp án đúng: D.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6. Liên kết hoá học trong phân tử NH~3~ là liên kết",
            "options": [
                "A. kim loại.",
                "*B. cộng hoá trị có cực.*",
                "C. ion.",
                "D. cộng hoá trị không cực."
            ],
            "solution": "Hiệu độ âm điện giữa N (3,04) và H (2,20) là: Δχ = 3,04 - 2,20 = 0,84 (thuộc khoảng 0,4 ≤ Δχ < 1,7). Do đó, liên kết N-H trong phân tử ammonia (NH~3~) là liên kết cộng hóa trị có cực (cặp electron dùng chung bị lệch về phía nguyên tử nitrogen có độ âm điện lớn hơn).\n*Đáp án đúng: B.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7. Phổ khối lượng của hợp chất hữu cơ X thu được như hình vẽ:",
            "image": "ms_cau7_hd.png",
            "subprompt": "Phân tử khối của hợp chất hữu cơ X là",
            "options": [
                "A. 28.",
                "B. 75.",
                "*C. 44.*",
                "D. 30."
            ],
            "solution": "Trên phổ khối lượng (MS), peak ion phân tử [M^+^] thường là peak có giá trị m/z lớn nhất (ngoại trừ các peak đồng vị có cường độ rất nhỏ), giá trị m/z của peak này chính bằng phân tử khối của hợp chất. Trong phổ khối lượng của X, peak ion phân tử xuất hiện tại m/z = 44. Do đó, phân tử khối của hợp chất hữu cơ X là 44.\n*Đáp án đúng: C.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8. Chọn phát biểu không đúng?",
            "options": [
                "*A. Nitrogen là nguyên tố phổ biến nhất trong vỏ trái đất.*",
                "B. Nitrogen lỏng được dùng để bảo quản mẫu vật phẩm trong y học.",
                "C. Hai nguyên tử N trong phân tử nitrogen liên kết với nhau bằng liên kết ba.",
                "D. Ở nhiệt độ thường, nitrogen khá trơ về mặt hoá học."
            ],
            "solution": "Phát biểu A không đúng vì nguyên tố phổ biến nhất trong vỏ Trái Đất là Oxygen (chiếm khoảng 46,6% khối lượng), tiếp theo là Silicon (27,7%), Aluminium (8,1%),... Nitrogen chỉ chiếm tỉ lệ rất nhỏ trong vỏ Trái Đất (khoảng 0,002%), dù là thành phần chính chiếm 78% thể tích không khí.\n*Đáp án đúng: A.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9. Khí NO~x~ là một trong các nguyên nhân gây nên hiện tượng mưa acid, mù quang hóa, hiện tượng phú dưỡng,... làm ô nhiễm môi trường. Nguồn phát sinh NO~x~ có nguồn gốc tự nhiên là",
            "options": [
                "A. hoạt động giao thông vận tải.",
                "*B. hoạt động của núi lửa.*",
                "C. hoạt động đốt nhiên liệu.",
                "D. hoạt động của nhà máy nhiệt điện."
            ],
            "solution": "Các hoạt động giao thông, đốt nhiên liệu, nhà máy nhiệt điện là nguồn nhân tạo do con người gây ra. Nguồn phát sinh khí NO~x~ tự nhiên trong khí quyển chủ yếu là do hoạt động của núi lửa, các cơn giông bão có tia lửa điện sấm sét và quá trình chuyển hóa sinh học của vi sinh vật trong đất.\n*Đáp án đúng: B.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10. Trong các phát biểu sau đây, phát biểu nào đúng khi một hệ ở trạng thái cân bằng?",
            "options": [
                "A. Phản ứng nghịch đã dừng.",
                "*B. Nồng độ của các chất trong hệ không đổi.*",
                "C. Nồng độ chất tham gia và sản phẩm bằng nhau.",
                "D. Phản ứng thuận đã dừng."
            ],
            "solution": "Cân bằng hóa học là cân bằng động, tại trạng thái cân bằng cả phản ứng thuận và phản ứng nghịch vẫn tiếp tục diễn ra với tốc độ bằng nhau (v~thuận~ = v~nghịch~ > 0), do đó nồng độ của các chất tham gia và sản phẩm không thay đổi theo thời gian.\n*Đáp án đúng: B.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11. Để xác định nồng độ mol của một dung dịch NaOH, người ta đã tiến hành chuẩn độ bằng dung dịch HCl 0,1M. Để chuẩn độ 10 mL dung dịch NaOH này cần 12 mL dung dịch HCl. Nồng độ của dung dịch NaOH trên là",
            "options": [
                "A. 0,012M.",
                "*B. 0,12M.*",
                "C. 0,1M.",
                "D. 1,2M."
            ],
            "solution": "Phương trình phản ứng chuẩn độ: NaOH + HCl → NaCl + H~2~O\nTại điểm tương đương: n(NaOH) = n(HCl)\n⇔ C(NaOH) × V(NaOH) = C(HCl) × V(HCl)\n⇔ C(NaOH) × 10 mL = 0,1 M × 12 mL\n⇔ C(NaOH) = (0,1 × 12) / 10 = 0,12 M.\n*Đáp án đúng: B.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12. Liên kết hoá học chủ yếu trong phân tử hợp chất hữu cơ là",
            "options": [
                "A. liên kết ion.",
                "B. liên kết cho-nhận.",
                "*C. liên kết cộng hoá trị.*",
                "D. liên kết hydrogen."
            ],
            "solution": "Trong phân tử hợp chất hữu cơ, nguyên tử carbon có độ âm điện trung bình (2,55), dễ dàng góp chung electron với các nguyên tử carbon khác và với các nguyên tử phi kim như H, O, N, halogen,... Do đó, liên kết hóa học chủ yếu trong hợp chất hữu cơ là liên kết cộng hóa trị.\n*Đáp án đúng: C.*"
        }
    ],

    # PHẦN II. Câu trắc nghiệm đúng sai. (2 câu - 2,0 điểm)
    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1. Sulfur và các hợp chất của nó có vai trò quan trọng trong sản xuất và đời sống, nhưng khí thải và chất thải chứa sulfur nếu không được xử lí đúng cách sẽ gây tác động tiêu cực đến sức khỏe con người và môi trường.",
            "statements": [
                ("a) Ở điều kiện thường, lưu huỳnh là chất rắn, màu vàng, không tan trong nước, tan nhiều trong các dung môi hữu cơ như benzene, carbon disulfide (CS~2~)...", "Đ"),
                ("b) Khí SO~2~ có thể được dùng làm chất tẩy màu, khi hoà tan khí này vào nước tạo ra dung dịch làm quỳ tím hoá xanh.", "S"),
                ("c) Thuốc thử nhận biết sulfuric acid và muối sulfate là ion Ba^2+^ trong BaCl~2~, Ba(OH)~2~, Ba(NO~3~)~2~.", "Đ"),
                ("d) Sulfuric acid đặc lấy nước từ hợp chất carbohydrate và khiến chúng hoá đen chứng tỏ sulfuric acid có tính háo nước.", "Đ")
            ],
            "solution": "• a) Đúng. Ở điều kiện thường, sulfur là chất rắn màu vàng, không tan trong nước nhưng tan nhiều trong các dung môi không phân cực như benzene, carbon disulfide (CS~2~).\n• b) Sai. SO~2~ tan trong nước tạo thành dung dịch sulfurous acid (H~2~SO~3~), dung dịch này có tính acid làm quỳ tím hóa đỏ (không phải hóa xanh).\n• c) Đúng. Để nhận biết ion SO~4~^2-^ (trong H~2~SO~4~ và muối sulfate), thuốc thử đặc trưng là dung dịch chứa cation Ba^2+^ (như BaCl~2~, Ba(OH)~2~, Ba(NO~3~)~2~), tạo kết tủa trắng barium sulfate (BaSO~4~) bền không tan trong acid.\n• d) Đúng. H~2~SO~4~ đặc có tính háo nước cực mạnh, nó chiếm lấy các nguyên tố H và O (với tỉ lệ 2:1 như nước) từ các carbohydrate (như đường saccarose, tinh bột, xenlulozơ), biến chúng thành carbon xốp màu đen."
        },
        {
            "id": 2,
            "prompt": "Câu 2. Hợp chất hữu cơ X được sử dụng rộng rãi trong đời sống như làm dung môi, nhiên liệu sinh học và thành phần chính của đồ uống có cồn. Kết quả phân tích nguyên tố hợp chất X cho thấy có 52,17% carbon, 34,78% oxygen, còn lại là hydrogen. Tỉ khối hơi của X so với khí hydrogen bằng 23. Phổ IR của X có dải hấp thụ đặc trưng được mô tả trong hình dưới đây:",
            "image": "ir_cau2_p2.png",
            "statements": [
                ("a) Công thức phân tử của X là C~2~H~6~O.", "Đ"),
                ("b) Do X tan vô hạn trong nước và có nhiệt độ sôi thấp hơn nước, nên không thể tách X khỏi hỗn hợp với nước bằng phương pháp chưng cất.", "S"),
                ("c) Dựa vào phổ IR, có thể dự đoán phân tử X chứa nhóm chức -OH.", "Đ"),
                ("d) Chất CH~3~OCH~3~ không phải là đồng phân cấu tạo của X.", "S")
            ],
            "solution": "Phân tích cấu tạo và công thức của X:\n- Khối lượng mol của X: M~X~ = 23 × 2 = 46 g/mol.\n- Phần trăm hydrogen: %H = 100% - 52,17% - 34,78% = 13,05%.\n- Đặt CTPT của X là C~x~H~y~O~z~:\n  x = (46 × 52,17%) / 12 = 2\n  y = (46 × 13,05%) / 1 = 6\n  z = (46 × 34,78%) / 16 = 1\n  ⇒ Công thức phân tử của X là C~2~H~6~O (M = 46).\n- Phổ IR của X xuất hiện peak hấp thụ mạnh và rộng ở 3391 cm^-1^ (nằm trong vùng 3600 - 3300 cm^-1^ đặc trưng của liên kết O-H alcohol) và dải tín hiệu ở 1102, 1055 cm^-1^ (liên kết C-O) ⇒ X là ethanol: C~2~H~5~OH.\n\nĐánh giá các phát biểu:\n• a) Đúng. Công thức phân tử của X là C~2~H~6~O.\n• b) Sai. Mặc dù tan vô hạn trong nước nhưng ethanol sôi ở 78,3 °C còn nước sôi ở 100 °C, nhiệt độ sôi chênh lệch nhau hơn 20 °C nên hoàn toàn có thể tách và nâng cao nồng độ ethanol bằng phương pháp chưng cất phân đoạn.\n• c) Đúng. Dựa vào dải hấp thụ tù, mạnh ở 3391 cm^-1^ trên phổ IR chứng tỏ phân tử X có chứa nhóm chức alcohol -OH.\n• d) Sai. Dimethyl ether (CH~3~OCH~3~) có cùng công thức phân tử C~2~H~6~O với ethanol nhưng khác cấu tạo liên kết, do đó CH~3~OCH~3~ chính là một đồng phân cấu tạo của X."
        }
    ],

    # PHẦN III. Câu trắc nghiệm trả lời ngắn. (4 câu - 2,0 điểm)
    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1. Cho cân bằng hóa học: N~2~(g) + O~2~(g) ⇌ 2NO(g)  Δ~r~H°~298~ = +180 kJ.\nCó các biện pháp sau:\n(1) Tăng nhiệt độ.\n(2) Hạ nhiệt độ.\n(3) Tăng nồng độ O~2~.\n(4) Tăng áp suất chung của hệ phản ứng.\n(5) Dùng thêm chất xúc tác V~2~O~5~.\nHỏi có bao nhiêu biện pháp làm cân bằng trên chuyển dịch?",
            "answer": "3",
            "solution": "- Phản ứng thuận có Δ~r~H°~298~ = +180 kJ > 0 nên phản ứng thuận là phản ứng thu nhiệt.\n- Số mol khí ở 2 vế của phương trình bằng nhau: 1 mol N~2~ + 1 mol O~2~ (2 mol khí) ⇌ 2 mol NO (2 mol khí).\nXét các biện pháp:\n(1) Tăng nhiệt độ: Cân bằng chuyển dịch theo chiều thu nhiệt (chiều thuận) ⇒ CÓ chuyển dịch.\n(2) Hạ nhiệt độ: Cân bằng chuyển dịch theo chiều tỏa nhiệt (chiều nghịch) ⇒ CÓ chuyển dịch.\n(3) Tăng nồng độ O~2~: Cân bằng chuyển dịch theo chiều làm giảm nồng độ O~2~ (chiều thuận) ⇒ CÓ chuyển dịch.\n(4) Tăng áp suất chung của hệ: Vì số mol khí 2 vế bằng nhau nên áp suất KHÔNG làm chuyển dịch cân bằng.\n(5) Dùng chất xúc tác V~2~O~5~: Chất xúc tác chỉ làm tăng tốc độ phản ứng, KHÔNG làm chuyển dịch cân bằng.\n⇒ Có 3 biện pháp làm cân bằng chuyển dịch: (1), (2), (3).\n*Đáp án: 3*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Cho các chất sau: HCOOH, CaCO~3~, C~2~H~4~, CO, C~6~H~12~O~6~, NaCN. Số chất thuộc loại hợp chất hữu cơ là bao nhiêu?",
            "answer": "3",
            "solution": "Theo định nghĩa: Hợp chất hữu cơ là hợp chất của carbon (trừ carbon monoxide CO, carbon dioxide CO~2~, muối carbonate, muối cyanide, carbide,...).\n- HCOOH: Formic acid (hợp chất hữu cơ).\n- CaCO~3~: Calcium carbonate (hợp chất vô cơ).\n- C~2~H~4~: Ethylene (hydrocarbon, hợp chất hữu cơ).\n- CO: Carbon monoxide (hợp chất vô cơ).\n- C~6~H~12~O~6~: Glucose (hợp chất hữu cơ).\n- NaCN: Sodium cyanide (hợp chất vô cơ).\n⇒ Có 3 chất thuộc loại hợp chất hữu cơ là: HCOOH, C~2~H~4~, C~6~H~12~O~6~.\n*Đáp án: 3*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Các bước tách một hợp chất hữu cơ bằng phương pháp chiết được mô tả như sau:\n(1) Để yên phễu chiết cho hai lớp chất lỏng phân tách hoàn toàn.\n(2) Lắc nhẹ phễu chiết để chất cần tách chuyển sang dung môi chiết.\n(3) Cho hỗn hợp chất lỏng không trộn lẫn vào phễu chiết cùng dung môi thích hợp.\n(4) Mở khóa phễu chiết để thu từng lớp chất lỏng riêng biệt.\nHãy sắp xếp các bước trên theo đúng thứ tự thực hiện của phương pháp chiết.",
            "answer": "3214",
            "solution": "Quy trình chuẩn thực hiện phương pháp chiết lỏng - lỏng trong phòng thí nghiệm:\n- Bước 1 (3): Cho hỗn hợp chất lỏng và dung môi chiết thích hợp vào phễu chiết.\n- Bước 2 (2): Lắc nhẹ phễu chiết nhiều lần (kết hợp mở khóa xả áp suất) để chất cần tách chuyển tối đa sang pha dung môi chiết.\n- Bước 3 (1): Đặt phễu lên giá, mở nút phễu và để yên cho hai lớp chất lỏng phân tách ranh giới hoàn toàn.\n- Bước 4 (4): Mở khóa phễu chiết để từ từ tháo lớp chất lỏng phía dưới vào bình hứng, sau đó thu lớp chất lỏng phía trên.\n⇒ Thứ tự thực hiện: (3) → (2) → (1) → (4) ➔ Kết quả: 3214.\n*Đáp án: 3214*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Geraniol là một hợp chất hữu cơ tự nhiên có mùi hoa hồng đặc trưng, được sử dụng rộng rãi trong sản xuất nước hoa, mỹ phẩm và hương liệu thực phẩm; công thức cấu tạo của geraniol được biểu diễn như sau:",
            "image": "geraniol_hd.png",
            "subprompt": "Số nguyên tử hydrogen có trong phân tử geraniol là bao nhiêu?",
            "answer": "18",
            "solution": "Xác định công thức cấu tạo và phân tử của geraniol:\n- Công thức cấu tạo đầy đủ: (CH~3~)~2~C=CH-CH~2~-CH~2~-C(CH~3~)=CH-CH~2~OH\n- Đếm các nguyên tử trong phân tử:\n  + Đầu mạch: 2 nhóm -CH~3~ liên kết với =C: 6 nguyên tử H.\n  + Nhóm =CH-: 1 nguyên tử H.\n  + 2 nhóm -CH~2~- liên tiếp: 2 × 2 = 4 nguyên tử H.\n  + Nhóm =C(CH~3~)-: 3 nguyên tử H.\n  + Nhóm =CH-: 1 nguyên tử H.\n  + Đuôi mạch -CH~2~OH: 2 + 1 = 3 nguyên tử H.\n⇒ Tổng số nguyên tử hydrogen là: 6 + 1 + 4 + 3 + 1 + 3 = 18 nguyên tử H.\n(Công thức phân tử của geraniol là C~10~H~18~O, phân tử có 2 liên kết đôi C=C và 1 nhóm -OH, độ bất bão hòa k = 2).\n*Đáp án: 18*"
        }
    ],

    # B. PHẦN TỰ LUẬN (3đ)
    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1. (0,5 điểm)\nCho 0,04 mol Mg tan hết trong dung dịch HNO~3~ thấy thoát ra 0,01 mol khí X là sản phẩm khử duy nhất (đkc). Xác định công thức của khí X?",
            "num_lines": 6,
            "solution": "• Số mol electron do Mg nhường:\n  Mg → Mg^2+^ + 2e\n  n(e nhường) = 2 × n(Mg) = 2 × 0,04 = 0,08 mol.\n• Gọi a là số electron mà 1 phân tử khí X nhận:\n  n(e nhận) = a × n(X) = 0,01 × a (mol).\n• Áp dụng định luật bảo toàn electron:\n  n(e nhường) = n(e nhận) ⇔ 0,08 = 0,01 × a ⇔ a = 8.\n• Quá trình khử tạo khí nhận 8 electron:\n  2NO~3~^-^ + 10H^+^ + 8e → N~2~O↑ + 5H~2~O (số oxi hóa N giảm từ +5 xuống +1 trong N~2~O).\n⇒ Khí X là dinitrogen monoxide (công thức phân tử: N~2~O)."
        },
        {
            "id": 2,
            "prompt": "Câu 2. (1,0 điểm)\nBằng phương pháp hoá học hãy nhận biết các dung dịch sau: NaCl, (NH~4~)~2~SO~4~ và NH~4~Cl. Viết các phương trình phản ứng xảy ra.",
            "num_lines": 8,
            "solution": "1. Các bước tiến hành nhận biết:\n- Trích mẫu thử từ các lọ dung dịch và đánh số thứ tự tương ứng.\n- Cho từ từ dung dịch Ba(OH)~2~ dư vào từng mẫu thử và đun nóng nhẹ:\n  + Mẫu thử xuất hiện kết tủa trắng, đồng thời có khí mùi khai bay ra làm xanh giấy quỳ tím ẩm là dung dịch (NH~4~)~2~SO~4~.\n  + Mẫu thử chỉ có khí mùi khai bay ra làm xanh giấy quỳ tím ẩm (không có kết tủa) là dung dịch NH~4~Cl.\n  + Mẫu thử không có hiện tượng gì xuất hiện là dung dịch NaCl.\n\n2. Phương trình phản ứng xảy ra:\n• (NH~4~)~2~SO~4~ + Ba(OH)~2~ → BaSO~4~↓ (trắng) + 2NH~3~↑ (mùi khai) + 2H~2~O\n  (hoặc phương trình ion rút gọn: Ba^2+^ + SO~4~^2-^ → BaSO~4~↓ và NH~4~^+^ + OH^-^ → NH~3~↑ + H~2~O)\n• 2NH~4~Cl + Ba(OH)~2~ → BaCl~2~ + 2NH~3~↑ + 2H~2~O\n  (hoặc phương trình ion rút gọn: NH~4~^+^ + OH^-^ → NH~3~↑ + H~2~O)."
        },
        {
            "id": 3,
            "prompt": "Câu 3. (1,0 điểm)\nAcetic acid được sử dụng rộng rãi trên thế giới trong nhiều ngành công nghiệp khác nhau như tạo ra polymer ứng dụng trong sơn, chất kết dính, là dung môi hoà tan các chất hoá học, sản xuất và bảo quản thực phẩm, đặc biệt dùng để sản xuất giấm.\na) Lập công thức phân tử của acetic acid, biết kết quả phân tích nguyên tố acetic acid có 40% C; 53,33% O về khối lượng; còn lại là H. Phân tử khối của acetic acid được xác định trên phổ khối lượng tương ứng với peak có giá trị m/z lớn nhất.\nb) Theo khuyến cáo an toàn, lượng acetic acid trong thực phẩm không nên vượt quá 10 gam trên mỗi một kg thực phẩm. Một cơ sở chế biến thực phẩm dùng 50 mL giấm ăn 5% để tạo vị chua cho 1,0 kg dưa muối. Biết khối lượng riêng của giấm ăn bằng 1,00 g/mL. Hỏi lượng acetic acid sử dụng có vượt quá giới hạn an toàn hay không, giải thích?",
            "image": "ms_acetic_acid_hd.png",
            "num_lines": 12,
            "solution": "a) Lập công thức phân tử của acetic acid:\n- Phần trăm khối lượng của hydrogen:\n  %H = 100% - 40% - 53,33% = 6,67%.\n- Trên phổ khối lượng (MS), peak ion phân tử [M^+^] có giá trị m/z lớn nhất là 60 ⇒ Phân tử khối của acetic acid là M = 60 g/mol.\n- Đặt công thức phân tử của acetic acid là C~x~H~y~O~z~:\n  x = (60 × 40%) / 12 = 2\n  y = (60 × 6,67%) / 1 = 4\n  z = (60 × 53,33%) / 16 = 2\n⇒ Công thức phân tử của acetic acid là C~2~H~4~O~2~ (hoặc CH~3~COOH).\n\nb) Tính toán và đánh giá độ an toàn:\n- Khối lượng dung dịch giấm ăn 5% đã dùng:\n  m(dd giấm) = V × D = 50 mL × 1,00 g/mL = 50 gam.\n- Khối lượng acetic acid nguyên chất có trong 50 mL giấm ăn:\n  m(acetic acid) = m(dd giấm) × 5% = 50 × 0,05 = 2,5 gam.\n- Lượng acetic acid có trong 1,0 kg dưa muối là 2,5 gam.\n- So sánh với khuyến cáo an toàn: 2,5 gam/kg < 10 gam/kg (giới hạn tối đa cho phép).\n⇒ Kết luận: Lượng acetic acid mà cơ sở chế biến sử dụng hoàn toàn KHÔNG vượt quá giới hạn an toàn thực phẩm."
        },
        {
            "id": 4,
            "prompt": "Câu 4. (0,5 điểm)\nHãy viết các đồng phân mạch hở có cùng công thức phân tử C~2~H~7~N?",
            "num_lines": 6,
            "solution": "Độ bất bão hòa của hợp chất C~2~H~7~N:\nk = (2 × 2 + 2 + 1 - 7) / 2 = 0 ⇒ Hợp chất là amine no, đơn chức, mạch hở.\nCác đồng phân cấu tạo amine mạch hở ứng với C~2~H~7~N gồm:\n1. Amine bậc 1: CH~3~-CH~2~-NH~2~ (Tên gọi: ethanamine hoặc ethylamine)\n2. Amine bậc 2: CH~3~-NH-CH~3~ (Tên gọi: N-methylmethanamine hoặc dimethylamine)\n⇒ Có 2 đồng phân mạch hở."
        }
    ]
}

def generate_docx(is_teacher=False):
    doc = Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line=DATA_EXAM["school_line"],
        exam_title=DATA_EXAM["exam_title"],
        subject_line=DATA_EXAM["subject_line"],
        num_pages=DATA_EXAM["num_pages"],
        duration=DATA_EXAM["duration"],
        exam_code=DATA_EXAM["exam_code"],
        atomic_masses=DATA_EXAM["atomic_masses"],
        is_teacher=is_teacher
    )

    # A. PHẦN TRẮC NGHIỆM (7đ)
    add_section_header(
        doc,
        title="A. PHẦN TRẮC NGHIỆM (7 điểm)",
        subtitle="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    for item in DATA_EXAM["part1"]:
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
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_path, width=Inches(3.8))

        if "subprompt" in item:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(1)
            p_sub.paragraph_format.space_after = Pt(1)
            render_rich_text(p_sub, item["subprompt"], base_size=10.5)

        render_options_no_table(doc, item["options"], is_teacher=is_teacher)
        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN II. Câu trắc nghiệm đúng sai.
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    for item in DATA_EXAM["part2"]:
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
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_path, width=Inches(4.5))

        for stmt, ans in item["statements"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt, ans=ans, is_teacher=is_teacher, base_size=10.5)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN III. Câu trắc nghiệm trả lời ngắn.
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm trả lời ngắn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Thí sinh điền kết quả vào ô trả lời tương ứng."
    )

    for item in DATA_EXAM["part3"]:
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
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_path, width=Inches(3.4))

        if "subprompt" in item:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(1)
            p_sub.paragraph_format.space_after = Pt(1)
            render_rich_text(p_sub, item["subprompt"], base_size=10.5)

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

    # B. PHẦN TỰ LUẬN (3đ)
    add_section_header(
        doc,
        title="B. PHẦN TỰ LUẬN (3,0 điểm)",
        subtitle="Thí sinh trình bày chi tiết lời giải vào giấy làm bài thi."
    )

    for item in DATA_EXAM["part4"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        if "image" in item:
            img_path = os.path.join(DIR_IMAGES, item["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(2)
                p_img.add_run().add_picture(img_path, width=Inches(3.8))

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])
        elif not is_teacher:
            num_lines = item.get("num_lines", 6)
            add_student_writing_lines(doc, num_lines=num_lines)

    # Kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(0)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(r_end, font_size=11, bold=True)

    suffix = "_GiaoVien.docx" if is_teacher else "_HocSinh.docx"
    filename = f"De_Kiem_Tra_CK1_Hoa_11_BuiThiXuan_Ma1102{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
