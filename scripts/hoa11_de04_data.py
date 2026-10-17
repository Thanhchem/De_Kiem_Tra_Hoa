# -*- coding: utf-8 -*-
"""
Dữ liệu Đề kiểm tra / Thi thử Giữa học kì I - Môn Hóa học 11 (Đề số 04 / Mã đề 104)
Năm học 2026 - 2027 - Thầy TRẦN VĂN THẠNH
Bao gồm:
- Phần I: 18 câu trắc nghiệm nhiều phương án lựa chọn (4,5 điểm)
- Phần II: 4 câu trắc nghiệm đúng/sai (4,0 điểm)
- Phần III: 6 câu trắc nghiệm trả lời ngắn (1,5 điểm)
Đầy đủ đáp án, phương án, phân tích và lời giải chi tiết chuẩn mực theo chương trình GDPT 2018.
"""

EXAM_INFO = {
    "title_left_1": "SỞ GIÁO DỤC VÀ ĐÀO TẠO",
    "title_left_2": "TRƯỜNG THPT CHUYÊN / THPT",
    "sub_left": "(Đề thi gồm có 04 trang)",
    "title_right_1": "ĐỀ THI THỬ GIỮA HỌC KÌ I",
    "title_right_2": "MÔN: HÓA HỌC - LỚP 11",
    "time_str": "Thời gian làm bài: 45 phút (Không kể thời gian phát đề)",
    "made": "104",
    "year": "NĂM HỌC 2026 - 2027",
}

# -----------------------------------------------------------------------------
# PHẦN I: 18 CÂU TRẮC NGHIỆM NHIỀU LỰA CHỌN (4,5 ĐIỂM - MỖI CÂU 0,25 ĐIỂM)
# -----------------------------------------------------------------------------
PART_1_QUESTIONS = [
    {
        "id": 1,
        "content": "Ở nhiệt độ cao khoảng 3000 ^o^C, nitrogen trong không khí phản ứng với oxygen tạo thành hợp chất nào sau đây?",
        "options": [
            "A. NO.",
            "B. NO~2~.",
            "C. NH~3~.",
            "D. N~2~O."
        ],
        "correct": "A",
        "explanation": "Ở nhiệt độ rất cao khoảng 3000 ^o^C (hoặc khi có tia lửa điện sấm sét), liên kết ba bền vững giữa hai nguyên tử nitrogen (N≡N) bị bẻ gãy, nitrogen phản ứng với oxygen trong không khí tạo thành nitrogen monoxide:\nN~2~(g) + O~2~(g) ⇌ 2NO(g) (ở khoảng 3000 ^o^C)\nSau đó ở điều kiện thường, khí NO không màu dễ dàng kết hợp tiếp với oxygen tạo NO~2~ màu nâu đỏ (2NO + O~2~ → 2NO~2~)."
    },
    {
        "id": 2,
        "content": "Yếu tố nào sau đây luôn luôn *không* làm dịch chuyển cân bằng của hệ phản ứng?",
        "options": [
            "A. Áp suất.",
            "B. Nhiệt độ.",
            "C. Chất xúc tác.",
            "D. Nồng độ."
        ],
        "correct": "C",
        "explanation": "Chất xúc tác làm tăng tốc độ của phản ứng thuận và tốc độ của phản ứng nghịch với số lần bằng nhau, do đó chỉ giúp hệ phản ứng nhanh chóng đạt đến trạng thái cân bằng chứ *không làm dịch chuyển cân bằng hóa học* và không làm thay đổi hằng số cân bằng K~C~."
    },
    {
        "id": 3,
        "content": "Cấu hình electron của nguyên tử nitrogen (Z = 7) là",
        "options": [
            "A. 1s^2^2s^3^2p^3^.",
            "B. 1s^2^2s^2^2p^3^.",
            "C. 1s^2^2s^2^2p^4^.",
            "D. 1s^2^2s^2^2p^5^."
        ],
        "correct": "B",
        "explanation": "Nguyên tử nitrogen có số hiệu nguyên tử Z = 7. Thứ tự mức năng lượng và sự phân bố electron vào các phân lớp theo đúng quy tắc là: 1s^2^ 2s^2^ 2p^3^. Nguyên tử nitrogen có 2 lớp electron (chu kì 2) và có 5 electron lớp ngoài cùng (nhóm VA)."
    },
    {
        "id": 4,
        "content": "Phản ứng thuận nghịch là phản ứng",
        "options": [
            "A. phản ứng xảy ra theo hai chiều ngược nhau trong cùng điều kiện.",
            "B. chỉ xảy ra theo một chiều nhất định.",
            "C. xảy ra giữa hai chất khí.",
            "D. có phương trình hoá học được biểu diễn bằng mũi tên một chiều."
        ],
        "correct": "A",
        "explanation": "Theo định nghĩa SGK Hóa học 11: Phản ứng thuận nghịch là phản ứng hóa học trong đó ở cùng điều kiện, sự chuyển hóa chất phản ứng thành sản phẩm (chiều thuận) và sự chuyển hóa sản phẩm thành chất phản ứng (chiều nghịch) diễn ra đồng thời. Phương trình hóa học của phản ứng thuận nghịch được biểu diễn bằng mũi tên hai chiều (⇌)."
    },
    {
        "id": 5,
        "content": "Mưa acid tàn phá nhiều rừng cây, ăn mòn nhiều công trình kiến trúc bằng đá và kim loại. Tác nhân chính tạo ra mưa acid là",
        "options": [
            "A. CO~2~.",
            "B. SO~2~.",
            "C. H~2~S.",
            "D. CO."
        ],
        "correct": "B",
        "explanation": "Tác nhân chủ yếu gây ra hiện tượng mưa acid là khí sulfur dioxide (SO~2~) và các oxide của nitrogen (NO~x~). Khí SO~2~ bị oxi hóa trong khí quyển với xúc tác ion kim loại hoặc bụi tạo thành SO~3~, sau đó hòa tan vào nước mưa tạo thành sulfuric acid:\n2SO~2~ + O~2~ + 2H~2~O → 2H~2~SO~4~.\nDo đó, trong các phương án đã cho, SO~2~ là tác nhân chính gây ra mưa acid."
    },
    {
        "id": 6,
        "content": "Phương trình nào sau đây *không phải* là phương trình phản ứng thuận nghịch?",
        "options": [
            "A. C + CO~2~ ⇌ 2CO.",
            "B. H~2~ + I~2~ ⇌ 2HI.",
            "C. CaCO~3~ ⇌ CaO + CO~2~.",
            "D. NaOH + HCl → NaCl + H~2~O."
        ],
        "correct": "D",
        "explanation": "Phản ứng trung hòa giữa dung dịch acid mạnh (HCl) và base mạnh (NaOH):\nNaOH + HCl → NaCl + H~2~O\nlà phản ứng một chiều xảy ra hoàn toàn đến cùng, được biểu diễn bằng mũi tên đơn một chiều (→). Các phản ứng ở A, B, C đều là các phản ứng thuận nghịch xảy ra trong bình kín."
    },
    {
        "id": 7,
        "content": "Khi có sấm chớp, trong khí quyển sinh ra chất nào sau đây?",
        "options": [
            "A. Oxide carbon.",
            "B. Oxide nitrogen.",
            "C. Nước.",
            "D. Không có khí gì sinh ra."
        ],
        "correct": "B",
        "explanation": "Năng lượng cực lớn từ tia chớp tạo nhiệt độ rất cao (hàng nghìn độ C), kích thích phản ứng giữa nitrogen và oxygen trong khí quyển tạo thành nitrogen monoxide: N~2~ + O~2~ → 2NO (oxide nitrogen). Khí NO sau đó tiếp tục phản ứng với oxygen tạo NO~2~, hòa tan vào nước mưa tạo acid nitric HNO~3~ rơi xuống đất cung cấp phân đạm tự nhiên cho cây trồng."
    },
    {
        "id": 8,
        "content": "Trong công nghiệp, người ta thường điều chế N~2~ từ nguồn nguyên liệu nào sau đây?",
        "options": [
            "A. Nitric acid.",
            "B. Không khí.",
            "C. Ammonium nitrate.",
            "D. Ammonia."
        ],
        "correct": "B",
        "explanation": "Trong công nghiệp, nitrogen được sản xuất với quy mô lớn bằng phương pháp *chưng cất phân đoạn không khí lỏng*: Không khí sau khi làm sạch bụi, hơi nước và CO~2~ được hóa lỏng dưới áp suất cao và nhiệt độ thấp. Khi nâng dần nhiệt độ của không khí lỏng, khí nitrogen sôi ở -196 ^o^C bay ra trước và được thu hồi."
    },
    {
        "id": 9,
        "content": "Hiện tượng phú dưỡng là một biểu hiện của môi trường ao, hồ bị ô nhiễm do dư thừa các chất dinh dưỡng. Sự dư thừa dinh dưỡng chủ yếu do hàm lượng các ion nào sau đây vượt quá mức cho phép?",
        "options": [
            "A. Calcium, magnesium.",
            "B. Nitrate, phosphate.",
            "C. Sodium, potassium.",
            "D. Chloride, sulfate."
        ],
        "correct": "B",
        "explanation": "Hiện tượng phú dưỡng (eutrophication) xảy ra do nguồn nước mặt tiếp nhận hàm lượng dư thừa các chất dinh dưỡng chứa nguyên tố nitrogen (dưới dạng ion nitrate NO~3~^-^) và phosphorus (dưới dạng ion phosphate PO~4~^3-^, HPO~4~^2-^) từ phân bón nông nghiệp, nước thải sinh hoạt và công nghiệp, kích thích tảo và thực vật thủy sinh bùng phát làm cạn kiệt oxygen hòa tan."
    },
    {
        "id": 10,
        "content": "Hiện tượng mưa acid được định nghĩa là hiện tượng nước mưa có",
        "options": [
            "A. tính chất sẵn có trong tự nhiên.",
            "B. sự bốc hơi của nước rồi ngưng tụ.",
            "C. giá trị pH < 5,6.",
            "D. giá trị pH < 7."
        ],
        "correct": "C",
        "explanation": "Nước mưa thông thường hòa tan khí CO~2~ trong khí quyển có pH khoảng 5,6. Khi nước mưa hòa tan thêm các oxide có tính acid mạnh như SO~2~, NO~x~ tạo ra H~2~SO~4~, HNO~3~ làm giá trị pH giảm xuống *dưới 5,6* (pH < 5,6), hiện tượng đó được gọi là hiện tượng *mưa acid*."
    },
    {
        "id": 11,
        "content": "Bản chất liên kết giữa nguyên tử nitrogen và hydrogen trong phân tử NH~3~ là",
        "options": [
            "A. liên kết hydrogen.",
            "B. liên kết cộng hoá trị phân cực.",
            "C. liên kết ion.",
            "D. liên kết cộng hoá trị không phân cực."
        ],
        "correct": "B",
        "explanation": "Trong phân tử NH~3~, liên kết giữa nguyên tử N và mỗi nguyên tử H được hình thành do sự dùng chung một cặp electron. Vì độ âm điện của nitrogen (3,04) lớn hơn nhiều so với hydrogen (2,20), cặp electron dùng chung bị hút lệch về phía nguyên tử N, do đó liên kết N-H là *liên kết cộng hóa trị phân cực*."
    },
    {
        "id": 12,
        "content": "Dạng hình học của phân tử ammonia (NH~3~) là",
        "options": [
            "A. hình tứ diện đều.",
            "B. hình tam giác đều.",
            "C. hình chóp tam giác.",
            "D. đường thẳng."
        ],
        "correct": "C",
        "explanation": "Nguyên tử N ở trạng thái lai hóa sp^3^, liên kết với 3 nguyên tử H bằng 3 liên kết σ và còn lại 1 cặp electron tự do (chưa liên kết). Lực đẩy giữa cặp electron chưa liên kết và các cặp electron liên kết làm các liên kết N-H bị cụp xuống, tạo nên cấu trúc dạng *hình chóp tam giác* (nguyên tử N ở đỉnh chóp, 3 nguyên tử H ở đáy tam giác, góc liên kết H-N-H khoảng 107^o^)."
    },
    {
        "id": 13,
        "content": "Hằng số cân bằng K~C~ của một phản ứng xác định phụ thuộc vào yếu tố nào sau đây?",
        "options": [
            "A. Chất xúc tác.",
            "B. Nhiệt độ.",
            "C. Áp suất.",
            "D. Nồng độ."
        ],
        "correct": "B",
        "explanation": "Đối với một phản ứng thuận nghịch xác định, hằng số cân bằng K~C~ *chỉ phụ thuộc vào nhiệt độ*. Hằng số K~C~ hoàn toàn không phụ thuộc vào nồng độ ban đầu của các chất, áp suất hay sự có mặt của chất xúc tác."
    },
    {
        "id": 14,
        "content": "Sulfur được dân gian sử dụng để pha chế vào thuốc trị các bệnh ngoài da. Tên gọi dân gian của sulfur là",
        "options": [
            "A. giấm ăn.",
            "B. diêm sinh.",
            "C. phèn chua.",
            "D. đá vôi."
        ],
        "correct": "B",
        "explanation": "Trong dân gian từ xa xưa, sulfur (lưu huỳnh) thường được gọi là *diêm sinh* (hoặc lưu hoàng), được sử dụng để điều chế thuốc trị bệnh ngoài da (trị ghẻ, vảy nến), sấy khô chống nấm mốc dược liệu đông y hoặc dùng chế tạo thuốc súng."
    },
    {
        "id": 15,
        "content": "Phản ứng nào sau đây là phản ứng thuận nghịch?",
        "options": [
            "A. C~2~H~5~OH + 3O~2~ → 2CO~2~ + 3H~2~O.",
            "B. Mg + 2HCl → MgCl~2~ + H~2~.",
            "C. 3H~2~ + N~2~ ⇌ 2NH~3~.",
            "D. 2KClO~3~ → 2KCl + 3O~2~."
        ],
        "correct": "C",
        "explanation": "Phản ứng tổng hợp ammonia từ nitrogen và hydrogen:\nN~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)\nlà phản ứng thuận nghịch kinh điển (quá trình Haber - Bosch) diễn ra ở điều kiện nhiệt độ cao, áp suất cao và có chất xúc tác sắt (Fe). Các phản ứng ở A, B, D đều là phản ứng một chiều xảy ra hoàn toàn."
    },
    {
        "id": 16,
        "content": "Trong công nghiệp, phần lớn lượng sulfur đơn chất sau khi khai thác ở các mỏ được dùng làm nguyên liệu để",
        "options": [
            "A. sản xuất sulfuric acid.",
            "B. điều chế thuốc bảo vệ thực vật.",
            "C. lưu hoá cao su tự nhiên.",
            "D. bào chế thuốc đông y."
        ],
        "correct": "A",
        "explanation": "Trong nền công nghiệp hóa chất hiện đại, khoảng *90%* lượng sulfur đơn chất khai thác được sử dụng làm nguyên liệu trực tiếp để *sản xuất sulfuric acid (H~2~SO~4~)* theo phương pháp tiếp xúc; 10% còn lại dùng để lưu hóa cao su, sản xuất thuốc trừ sâu, phẩm nhuộm, diêm,..."
    },
    {
        "id": 17,
        "content": "Trong phân tử ammonia (NH~3~), nguyên tử nitrogen có số oxi hóa là",
        "options": [
            "A. +3.",
            "B. +4.",
            "C. -3.",
            "D. +5."
        ],
        "correct": "C",
        "explanation": "Trong hợp chất NH~3~, nguyên tử hydrogen có số oxi hóa +1. Gọi số oxi hóa của nitrogen là x, ta có:\nx + 3.(+1) = 0 ⇒ x = -3.\nSố oxi hóa -3 là số oxi hóa thấp nhất của nitrogen, do đó ammonia chỉ thể hiện tính khử trong các phản ứng oxi hóa - khử."
    },
    {
        "id": 18,
        "content": "Các yếu tố có thể ảnh hưởng đến sự chuyển dịch cân bằng hoá học là",
        "options": [
            "A. nồng độ, áp suất và diện tích bề mặt.",
            "B. nồng độ, nhiệt độ và áp suất.",
            "C. áp suất, nhiệt độ và chất xúc tác.",
            "D. nồng độ, nhiệt độ và chất xúc tác."
        ],
        "correct": "B",
        "explanation": "Theo nguyên lí Le Chatelier, các yếu tố làm chuyển dịch cân bằng hóa học gồm có: *nồng độ, nhiệt độ và áp suất* (đối với phản ứng có chất khí và có sự biến thiên tổng số mol khí). Chất xúc tác và diện tích bề mặt chất rắn chỉ ảnh hưởng đến tốc độ phản ứng chứ không làm chuyển dịch vị trí cân bằng."
    }
]

# -----------------------------------------------------------------------------
# PHẦN II: 4 CÂU TRẮC NGHIỆM ĐÚNG/SAI (4,0 ĐIỂM - MỖI CÂU 1,0 ĐIỂM)
# -----------------------------------------------------------------------------
PART_2_QUESTIONS = [
    {
        "id": 1,
        "content": "Khí NO~x~ là một trong các nguyên nhân gây nên hiện tượng mưa acid, sương mù quang hóa, hiện tượng phú dưỡng,… làm ô nhiễm môi trường nghiêm trọng.",
        "image": "acid_rain.png",
        "statements": [
            {"id": "a", "text": "Hoạt động giao thông vận tải làm phát sinh NO~x~.", "correct": True},
            {"id": "b", "text": "Hoạt động của nhà máy nhiệt điện làm phát sinh NO~x~.", "correct": True},
            {"id": "c", "text": "Hoạt động của núi lửa làm phát sinh NO~x~.", "correct": True},
            {"id": "d", "text": "Hoạt động đốt nhiên liệu hóa thạch làm phát sinh NO~x~.", "correct": True}
        ],
        "explanation": (
            "- Ý a) **ĐÚNG**: Động cơ đốt trong của các phương tiện giao thông (ô tô, xe máy, máy bay) vận hành ở nhiệt độ và áp suất cao làm N~2~ và O~2~ trong không khí phản ứng tạo thành NO, sau đó thoát ra môi trường hóa nâu thành NO~2~.\n"
            "- Ý b) **ĐÚNG**: Nhà máy nhiệt điện đốt cháy nhiên liệu than đá hoặc khí đốt ở nhiệt độ cao sinh ra lượng lớn NO~x~ cùng với SO~2~.\n"
            "- Ý c) **ĐÚNG**: Hoạt động phun trào của núi lửa với nhiệt độ macma cực lớn trên 1000 ^o^C kích hoạt phản ứng giữa N~2~ và O~2~ trong khí quyển, đồng thời giải phóng các chất chứa nitrogen từ sâu trong vỏ Trái Đất, sinh ra NO~x~ tự nhiên.\n"
            "- Ý d) **ĐÚNG**: Việc đốt nhiên liệu hóa thạch (than đá, dầu mỏ, khí thiên nhiên) trong các ngành công nghiệp và sinh hoạt là nguồn nhân tạo chủ yếu phát thải NO~x~ vào bầu khí quyển."
        )
    },
    {
        "id": 2,
        "content": "Cho các phát biểu sau về nguyên tố nitrogen (N, Z = 7):",
        "statements": [
            {"id": "a", "text": "Nguyên tố nitrogen thuộc chu kì 3, nhóm VA trong bảng tuần hoàn.", "correct": False},
            {"id": "b", "text": "Trong một số hợp chất, nguyên tử nitrogen có thể dùng cặp electron hoá trị riêng để tạo một liên kết cho – nhận với nguyên tử khác.", "correct": True},
            {"id": "c", "text": "Nguyên tử nguyên tố nitrogen có cấu hình electron là 1s^2^2s^2^2p^3^.", "correct": True},
            {"id": "d", "text": "Nguyên tử nguyên tố nitrogen có 3 electron hoá trị.", "correct": False}
        ],
        "explanation": (
            "- Ý a) **SAI**: Nitrogen (Z = 7) có cấu hình electron là 1s^2^2s^2^2p^3^, có 2 lớp electron nên thuộc **chu kì 2** (không phải chu kì 3), nhóm VA.\n"
            "- Ý b) **ĐÚNG**: Ở lớp vỏ ngoài cùng, nguyên tử nitrogen có 1 cặp electron tự do (chưa liên kết). Trong các phân tử như NH~3~, nitrogen có thể dùng cặp electron này để tạo liên kết cho - nhận với ion H^+^ tạo cation ammonium (NH~4~^+^).\n"
            "- Ý c) **ĐÚNG**: Với số hiệu Z = 7, cấu hình electron chuẩn là 1s^2^2s^2^2p^3^.\n"
            "- Ý d) **SAI**: Lớp ngoài cùng (n = 2) của nitrogen có 2 + 3 = 5 electron, do đó nitrogen có **5 electron hóa trị** (chứ không phải 3; 3 là số electron độc thân ở phân lớp 2p)."
        )
    },
    {
        "id": 3,
        "content": "Hydrogen chloride phân li trong nước theo phương trình sau: HCl → H^+^ + Cl^-^",
        "statements": [
            {"id": "a", "text": "Theo thuyết Bronsted – Lowry về acid – base thì HCl là acid.", "correct": True},
            {"id": "b", "text": "Khi hòa tan trong nước, chỉ một phần các phân tử HCl bị phân li thành ion.", "correct": False},
            {"id": "c", "text": "HCl là chất điện li mạnh.", "correct": True},
            {"id": "d", "text": "Khi hòa tan trong nước, toàn bộ các phân tử HCl đều bị phân li thành ion.", "correct": True}
        ],
        "explanation": (
            "- Ý a) **ĐÚNG**: Theo thuyết Brønsted – Lowry, acid là chất cho proton (H^+^). Trong dung dịch nước: HCl + H~2~O → H~3~O^+^ + Cl^-^, phân tử HCl nhường H^+^ cho H~2~O nên HCl là acid.\n"
            "- Ý b) **SAI**: HCl là chất điện li mạnh, khi tan trong nước phân li hoàn toàn thành ion, không phải \"chỉ một phần\".\n"
            "- Ý c) **ĐÚNG**: HCl là acid mạnh và là chất điện li mạnh điển hình.\n"
            "- Ý d) **ĐÚNG**: Đây là định nghĩa chất điện li mạnh: khi tan trong nước, toàn bộ các phân tử hòa tan đều phân li hoàn toàn thành ion."
        )
    },
    {
        "id": 4,
        "content": "Cho các phát biểu sau về đặc điểm của phản ứng thuận nghịch:",
        "statements": [
            {"id": "a", "text": "Trong hệ phản ứng thuận nghịch luôn tồn tại đồng thời chất sản phẩm và chất tham gia.", "correct": True},
            {"id": "b", "text": "Khi phản ứng thuận nghịch đạt trạng thái cân bằng hóa học, phản ứng dừng lại.", "correct": False},
            {"id": "c", "text": "Chỉ có phản ứng thuận nghịch mới đạt đến trạng thái cân bằng hóa học.", "correct": True},
            {"id": "d", "text": "Hiệu suất của phản ứng thuận nghịch luôn nhỏ hơn 100%.", "correct": True}
        ],
        "explanation": (
            "- Ý a) **ĐÚNG**: Do phản ứng xảy ra đồng thời theo hai chiều ngược nhau trong cùng điều kiện, chất phản ứng không bao giờ chuyển hóa hết thành sản phẩm, do đó trong hệ luôn tồn tại đồng thời cả chất phản ứng và sản phẩm.\n"
            "- Ý b) **SAI**: Cân bằng hóa học là **cân bằng động**. Tại trạng thái cân bằng, phản ứng thuận và nghịch vẫn tiếp tục diễn ra với tốc độ bằng nhau (v~thuận~ = v~nghịch~), chứ phản ứng không hề dừng lại.\n"
            "- Ý c) **ĐÚNG**: Phản ứng một chiều xảy ra cho đến khi chất phản ứng hết hoàn toàn nên không có trạng thái cân bằng; chỉ phản ứng thuận nghịch mới có trạng thái cân bằng hóa học.\n"
            "- Ý d) **ĐÚNG**: Do chất phản ứng không bao giờ phản ứng hết (luôn còn lại ở trạng thái cân bằng), hiệu suất của phản ứng thuận nghịch luôn nhỏ hơn 100% (H < 100%)."
        )
    }
]

# -----------------------------------------------------------------------------
# PHẦN III: 6 CÂU TRẢ LỜI NGẮN (1,5 ĐIỂM - MỖI CÂU 0,25 ĐIỂM HOẶC 0,5 ĐIỂM)
# -----------------------------------------------------------------------------
PART_3_QUESTIONS = [
    {
        "id": 1,
        "content": "Trong phân tử nitric acid, số nguyên tử oxygen bằng bao nhiêu?",
        "correct": "3",
        "explanation": "Nitric acid có công thức phân tử là HNO~3~. Trong một phân tử HNO~3~ gồm có 1 nguyên tử H, 1 nguyên tử N và **3 nguyên tử O**."
    },
    {
        "id": 2,
        "content": "Cho các nhận xét sau:\n(a) Ở trạng thái cân bằng, tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch.\n(b) Ở trạng thái cân bằng, các chất không phản ứng với nhau.\n(c) Ở trạng thái cân bằng, nồng độ chất sản phẩm luôn lớn hơn nồng độ chất ban đầu.\n(d) Ở trạng thái cân bằng, nồng độ các chất không thay đổi theo thời gian.\nCó bao nhiêu nhận xét đúng?",
        "correct": "2",
        "explanation": (
            "- (a) **ĐÚNG**: Theo định nghĩa trạng thái cân bằng hóa học: v~thuận~ = v~nghịch~.\n"
            "- (b) **SAI**: Cân bằng hóa học là cân bằng động, các chất vẫn tiếp tục phản ứng với nhau với tốc độ thuận bằng tốc độ nghịch.\n"
            "- (c) **SAI**: Nồng độ các chất tại cân bằng là không đổi, nhưng nồng độ sản phẩm có thể nhỏ hơn, bằng hoặc lớn hơn nồng độ chất ban đầu tùy thuộc vào hằng số K~C~ và điều kiện phản ứng.\n"
            "- (d) **ĐÚNG**: Vì tốc độ tạo thành bằng tốc độ tiêu hao của mỗi chất, nên nồng độ các chất trong hệ không thay đổi theo thời gian.\n"
            "⇒ Có **2** nhận xét đúng (gồm nhận xét a và d)."
        )
    },
    {
        "id": 3,
        "content": "Cho cân bằng hóa học (trong bình kín) sau:\nN~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)    Δ~r~H^o^~298~ = -92 kJ\nXét các tác động sau:\n(1) Thêm một lượng N~2~ hoặc H~2~;\n(2) Thêm một lượng NH~3~;\n(3) Tăng nhiệt độ của phản ứng;\n(4) Tăng áp suất của phản ứng;\n(5) Dùng thêm chất xúc tác.\nCó bao nhiêu tác động làm cho tỉ khối của hỗn hợp khí trong bình so với H~2~ tăng lên?",
        "correct": "2",
        "explanation": (
            "Tỉ khối của hỗn hợp khí so với H~2~ là: d = M~tb~ / 2 = m~hh~ / (2.n~hh~).\n"
            "- **Tác động (4) - Tăng áp suất:** Chiều thuận làm giảm số mol khí (từ 4 mol về 2 mol). Khi tăng áp suất, cân bằng dịch chuyển theo chiều thuận làm giảm số mol khí n~hh~. Vì bình kín nên m~hh~ bảo toàn, n~hh~ giảm ⇒ M~tb~ tăng ⇒ d tăng. (Thỏa mãn 1)\n"
            "- **Tác động (2) - Thêm một lượng NH~3~:** Khí NH~3~ có M = 17 g/mol, lớn hơn nhiều so với M~tb~ của hỗn hợp giàu H~2~ ban đầu. Thêm NH~3~ làm tăng nồng độ NH~3~, cân bằng dịch chuyển sang trái nhưng không triệt tiêu hết lượng NH~3~ thêm vào ⇒ M~tb~ sau cùng tăng ⇒ d tăng. (Thỏa mãn 2)\n"
            "- **Tác động (3) - Tăng nhiệt độ:** Phản ứng thuận tỏa nhiệt (Δ~r~H^o^~298~ < 0). Khi tăng nhiệt độ, cân bằng chuyển dịch theo chiều thu nhiệt (chiều nghịch) ⇒ n~hh~ tăng ⇒ M~tb~ giảm ⇒ d giảm. (Không thỏa mãn)\n"
            "- **Tác động (5) - Dùng thêm chất xúc tác:** Chất xúc tác không làm chuyển dịch cân bằng ⇒ tỉ khối không đổi. (Không thỏa mãn)\n"
            "- **Tác động (1) - Thêm N~2~ hoặc H~2~:** Nếu thêm H~2~ (M = 2), do H~2~ rất nhẹ làm M~tb~ giảm mạnh ⇒ không thỏa mãn trong mọi trường hợp.\n"
            "⇒ Có **2** tác động làm cho tỉ khối hỗn hợp so với H~2~ tăng lên (yếu tố 2 và 4)."
        )
    },
    {
        "id": 4,
        "content": "Ở 25 ^o^C, môi trường trung tính có giá trị pH bằng bao nhiêu?",
        "correct": "7",
        "explanation": "Ở 25 ^o^C, trong môi trường trung tính (như nước tinh khiết), nồng độ ion [H^+^] = [OH^-^] = 1,0.10^-7^ M ⇒ pH = -log[H^+^] = -log(10^-7^) = **7**."
    },
    {
        "id": 5,
        "content": "Cho các chất sau: N~2~, CuSO~4~, NaCl, NaOH, HCl, C~12~H~22~O~11~ (saccharose), C~2~H~5~OH (ethanol), CH~3~COOH (acetic acid). Có bao nhiêu chất là chất điện li?",
        "correct": "5",
        "explanation": (
            "Chất điện li là chất khi tan trong nước phân li ra các ion.\n"
            "- Các chất điện li gồm **5 chất**:\n"
            "  1. CuSO~4~ (muối tan, điện li mạnh: CuSO~4~ → Cu^2+^ + SO~4~^2-^)\n"
            "  2. NaCl (muối tan, điện li mạnh: NaCl → Na^+^ + Cl^-^)\n"
            "  3. NaOH (base mạnh, điện li mạnh: NaOH → Na^+^ + OH^-^)\n"
            "  4. HCl (acid mạnh, điện li mạnh: HCl → H^+^ + Cl^-^)\n"
            "  5. CH~3~COOH (acid yếu, điện li yếu: CH~3~COOH ⇌ CH~3~COO^-^ + H^+^)\n"
            "- Các chất không điện li: N~2~, C~12~H~22~O~11~ (đường), C~2~H~5~OH (rượu).\n"
            "⇒ Có tất cả **5** chất là chất điện li."
        )
    },
    {
        "id": 6,
        "content": "Cho các chất khí sau: H~2~S, NO, NO~2~, SO~2~. Có bao nhiêu chất khí gây ô nhiễm môi trường khi phát thải vào không khí?",
        "correct": "4",
        "explanation": (
            "Cả 4 chất khí trên đều là các khí độc hại, gây ô nhiễm không khí nghiêm trọng:\n"
            "1. H~2~S: khí độc mùi trứng thối, độc cho hệ hô hấp và thần kinh.\n"
            "2. NO: khí độc, dễ bị oxi hóa thành NO~2~, phá hủy tầng ozone, gây sương mù quang hóa.\n"
            "3. NO~2~: khí màu nâu đỏ có mùi hắc cay, rất độc với phổi, là tác nhân chính tạo mưa acid.\n"
            "4. SO~2~: khí có mùi hắc, gây viêm phế quản, là tác nhân chủ yếu tạo mưa acid.\n"
            "⇒ Cả **4** chất khí đều gây ô nhiễm môi trường."
        )
    }
]
