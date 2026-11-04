# -*- coding: utf-8 -*-
"""
Dữ liệu Đề kiểm tra Giữa học kì I - Môn KHTN 9 (ĐỀ CHÍNH THỨC)
Trường THCS Trần Cao Vân - UBND Phường Phú Xuân - Năm học 2026 - 2027
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

EXAM_INFO = {
    "ubnd": "UBND PHƯỜNG PHÚ XUÂN",
    "school": "TRƯỜNG THCS TRẦN CAO VÂN",
    "exam_title": "ĐỀ KIỂM TRA GIỮA HỌC KỲ I",
    "year": "NĂM HỌC 2026 – 2027",
    "subject": "MÔN: KHOA HỌC TỰ NHIÊN – LỚP 9",
    "time_str": "Thời gian làm bài: 60 phút (không kể thời gian phát đề)",
    "official_label": "ĐỀ CHÍNH THỨC",
    "pages_label": "(Đề kiểm tra có 03 trang)",
    "atomic_mass": "Cho Fe = 56 amu; O = 16 amu; C = 12 amu; H = 1 amu; Ag = 108 amu; N = 14 amu; Cu = 64 amu; Zn = 65 amu"
}

# -----------------------------------------------------------------------------
# PHẦN I. TRẮC NGHIỆM KHÁCH QUAN (7,0 ĐIỂM)
# 1. Trắc nghiệm nhiều lựa chọn (3,0 điểm; câu 1 - câu 12) - Mỗi câu đúng 0,25 điểm
# -----------------------------------------------------------------------------
PART1_QUESTIONS = [
    {
        "id": 1,
        "content": "Dụng cụ nào dưới đây dùng để phát hiện dòng điện cảm ứng?",
        "options": ["A. Ampe kế", "B. Điện kế", "C. Vôn kế", "D. Nhiệt lượng kế"],
        "correct": "B",
        "solution": "Điện kế (galvanometer) là dụng cụ đo có độ nhạy rất cao, chuyên dùng để phát hiện dòng điện cảm ứng có cường độ rất nhỏ hoặc xác định chiều của dòng điện trong mạch cảm ứng điện từ."
    },
    {
        "id": 2,
        "content": "Dụng cụ thí nghiệm nào được sử dụng chuyên dụng để tách các chất theo phương pháp chiết?",
        "options": ["A. Phễu chiết.", "B. Bát sứ.", "C. Bình cầu đáy tròn.", "D. Lưới tản nhiệt."],
        "correct": "A",
        "solution": "Phễu chiết là dụng cụ thí nghiệm chuyên dụng dùng để tách hai chất lỏng không hòa tan vào nhau (tách lớp theo khối lượng riêng) trong phương pháp chiết lỏng – lỏng."
    },
    {
        "id": 3,
        "content": "Các hóa chất cần được bảo quản như thế nào?",
        "options": [
            "A. Trong chai hoặc lọ nhựa, để ở chỗ tối.",
            "B. Trong chai hoặc lọ, được dán nhãn ghi thông tin về hóa chất.",
            "C. Trong chai hoặc lọ, có nắp đậy.",
            "D. Trong chai hoặc lọ, có nắp đậy và được dán nhãn ghi thông tin về hóa chất."
        ],
        "correct": "D",
        "solution": "Hóa chất trong phòng thí nghiệm bắt buộc phải được bảo quản trong chai hoặc lọ chuyên dụng có nắp đậy kín (tránh bay hơi, hút ẩm hoặc phản ứng với không khí) và được dán nhãn rõ ràng ghi đầy đủ thông tin tên, nồng độ, cảnh báo nguy hiểm."
    },
    {
        "id": 4,
        "content": "Công thức tính động năng (W_đ) của một vật có khối lượng m đang chuyển động với tốc độ v là",
        "options": [
            "A. W~đ~ = m.v",
            "B. W~đ~ = 1/2.m.v^2^",
            "C. W~đ~ = P.h",
            "D. W~đ~ = 1/2.m.v"
        ],
        "correct": "B",
        "solution": "Động năng của một vật khối lượng m chuyển động với tốc độ v được xác định bằng công thức: W~đ~ = 1/2·m·v^2^."
    },
    {
        "id": 5,
        "content": "Trong hệ đơn vị đo lường SI, đơn vị nào sau đây là đơn vị công suất?",
        "options": ["A. J", "B. N", "C. W", "D. kg"],
        "correct": "C",
        "solution": "Trong hệ SI, đơn vị đo công suất là Oát (Watt, kí hiệu W). Jun (J) là đơn vị đo công/năng lượng, Niutơn (N) là đơn vị đo lực, kilôgam (kg) là đơn vị đo khối lượng."
    },
    {
        "id": 6,
        "content": "Kim loại nào sau đây có tính dẻo nhất?",
        "options": ["A. Đồng (Cu).", "B. Sắt (Fe).", "C. Vàng (Au).", "D. Chì (Pb)."],
        "correct": "C",
        "solution": "Vàng (Au) là kim loại có tính dẻo cao nhất trong số các kim loại, có thể dát mỏng thành lá vàng có độ dày chỉ khoảng 0,0001 mm hoặc kéo thành sợi cực mảnh."
    },
    {
        "id": 7,
        "content": "Kim loại nào sau đây không tác dụng với dung dịch sulfuric acid (H~2~SO~4~) loãng?",
        "options": ["A. Magnesium (Mg).", "B. Sodium (Na).", "C. Zinc (Zn).", "D. Silver (Ag)."],
        "correct": "D",
        "solution": "Trong dãy hoạt động hóa học, bạc (Silver - Ag) đứng sau hydrogen (H) nên không có khả năng tác dụng với dung dịch acid loãng (như HCl, H~2~SO~4~ loãng) để giải phóng khí H~2~."
    },
    {
        "id": 8,
        "content": "Cho một mẩu dây Copper vào ống nghiệm đựng dung dịch Silver nitrate, sau một thời gian hiện tượng quan sát được là",
        "options": [
            "A. Dung dịch không màu chuyển sang màu xanh, dây Copper màu đỏ có một lớp màu trắng bạc bám vào.",
            "B. Màu xanh của dung dịch nhạt dần, dây Copper màu xanh có một lớp màu trắng bạc bám vào.",
            "C. Dung dịch không màu chuyển sang màu vàng, dây Copper màu đỏ có một lớp màu trắng bạc bám vào.",
            "D. Màu xanh của dung dịch nhạt dần, dây Copper màu đỏ có một lớp màu trắng bạc bám vào."
        ],
        "correct": "A",
        "solution": "Phương trình hóa học: Cu + 2AgNO~3~ → Cu(NO~3~)~2~ + 2Ag↓.\nHiện tượng: Dung dịch AgNO~3~ ban đầu không màu dần chuyển sang màu xanh lam đặc trưng của ion Cu^2+^, đồng thời kim loại bạc (Ag) màu trắng bạc sinh ra bám ngoài bề mặt mẩu dây đồng (màu đỏ)."
    },
    {
        "id": 9,
        "content": "Thép là hợp kim của sắt và carbon, trong đó hàm lượng carbon chiếm bao nhiêu phần trăm?",
        "options": ["A. Từ 2% đến 5%.", "B. Dưới 2%.", "C. Trên 5%.", "D. Từ 5% đến 10%."],
        "correct": "B",
        "solution": "Thép là hợp kim của sắt với carbon và một số nguyên tố khác, trong đó hàm lượng carbon chiếm dưới 2% (khác với gang có hàm lượng carbon từ 2% đến 5%)."
    },
    {
        "id": 10,
        "content": "Bản chất cốt lõi của quá trình chuyển hóa từ gang thành thép là gì?",
        "options": [
            "A. Tăng hàm lượng carbon có trong gang để làm tăng độ đàn hồi của vật liệu.",
            "B. Nung chảy gang rồi pha thêm các hóa chất để loại bỏ hoàn toàn các khí độc hại.",
            "C. Làm giảm các tạp chất C, Si, Mn,... trong gang bằng cách chuyển chúng thành các oxide và loại bỏ chúng để thu được thép.",
            "D. Khử các tạp chất còn lại bằng cách cho tác dụng với dung dịch acid mạnh."
        ],
        "correct": "C",
        "solution": "Bản chất của quá trình luyện gang thành thép là oxi hóa các tạp chất C, Si, Mn, S, P... có trong gang bằng cách thổi khí oxygen vào gang nóng chảy ở nhiệt độ cao, biến chúng thành các oxide dạng khí (CO~2~, SO~2~) hoặc tạo xỉ nổi lên trên để loại bỏ, làm giảm hàm lượng carbon xuống dưới 2%."
    },
    {
        "id": 11,
        "content": "Than hoạt tính được sử dụng trong sản xuất mặt nạ phòng độc, chất khử màu, khử mùi, là do khả năng",
        "options": ["A. dẫn điện tốt.", "B. không tan trong nước.", "C. mềm.", "D. hấp phụ cao."],
        "correct": "D",
        "solution": "Than hoạt tính có cấu trúc xốp với diện tích bề mặt cực lớn nên có khả năng hấp phụ (giữ lại trên bề mặt) rất mạnh các chất khí độc hại, các hạt màu và mùi khó chịu."
    },
    {
        "id": 12,
        "content": "Hai trạng thái biểu hiện trái ngược nhau của cùng một loại tính trạng (ví dụ: hoa tím và hoa trắng ở đậu Hà Lan) được gọi là",
        "options": [
            "A. cặp tính trạng tương phản.",
            "B. cặp tính trạng đồng nhất.",
            "C. cặp tính trạng tương đồng.",
            "D. biến dị tổ hợp."
        ],
        "correct": "A",
        "solution": "Cặp tính trạng tương phản là hai trạng thái biểu hiện trái ngược nhau của cùng một loại tính trạng (ví dụ: thân cao - thân lùn, hạt trơn - hạt nhăn, hoa tím - hoa trắng)."
    }
]

# -----------------------------------------------------------------------------
# 2. Trắc nghiệm Đúng - Sai (2,0 điểm; câu 13 - câu 14) - Mỗi câu 1,0 điểm
# -----------------------------------------------------------------------------
PART2_QUESTIONS = [
    {
        "id": 13,
        "title": "Tại công trường xây dựng, một cần cẩu điện được sử dụng để nâng một kiện hàng vật liệu có khối lượng m = 400 kg từ mặt đất lên độ cao h = 15 m trong thời gian t = 20 s. Bỏ qua mọi ma sát và lực cản của không khí. Chọn gốc thế năng tại mặt đất. Em hãy cho biết mỗi phát biểu sau là đúng hay sai.",
        "statements": [
            {
                "label": "a",
                "text": "Khi kiện hàng được nâng lên đứng yên ở độ cao 15 m, giá trị thế năng trọng trường được tính bằng biểu thức W~t~ = P.h.",
                "correct": True,
                "explain": "Đúng. Thế năng trọng trường của vật ở độ cao h so với mốc thế năng được tính bằng công thức: W~t~ = P.h = 10.m.h."
            },
            {
                "label": "b",
                "text": "Cơ năng của kiện hàng tại một vị trí bất kì bằng hiệu số giữa thế năng và động năng của kiện hàng tại vị trí đó.",
                "correct": False,
                "explain": "Sai. Cơ năng của vật bằng TỔNG giữa động năng và thế năng tại vị trí đó (W = W~đ~ + W~t~), không phải hiệu số."
            },
            {
                "label": "c",
                "text": "Công tối thiểu mà động cơ cần cẩu thực hiện để nâng kiện hàng lên độ cao 15 m là A = 6 000 J.",
                "correct": False,
                "explain": "Sai. Trọng lượng kiện hàng là P = 10.m = 10 × 400 = 4 000 N. Lực nâng tối thiểu F = P = 4 000 N. Công tối thiểu cần thực hiện là A = F.s = 4 000 × 15 = 60 000 J (đề bài ghi 6 000 J là thiếu một chữ số 0)."
            },
            {
                "label": "d",
                "text": "Công suất tối thiểu của động cơ cần cẩu trong quá trình nâng kiện hàng trên là P = 3 000 W.",
                "correct": True,
                "explain": "Đúng. Công suất tối thiểu của động cơ là: P = A / t = 60 000 / 20 = 3 000 W (tương đương 3 kW)."
            }
        ]
    },
    {
        "id": 14,
        "title": "Tại một xưởng mạ điện, người thợ tiến hành cho bột kẽm dư vào dung dịch CuSO~4~ thải ra từ bể mạ để thu hồi kim loại đồng và xử lý nước thải. Dựa vào kiến thức đã học, em hãy cho biết mỗi phát biểu sau là đúng hay sai.",
        "statements": [
            {
                "label": "a",
                "text": "Dãy hoạt động hóa học của các kim loại liên quan sắp xếp theo thứ tự giảm dần là: Cu, Zn.",
                "correct": False,
                "explain": "Sai. Trong dãy hoạt động hóa học, kẽm (Zn) hoạt động mạnh hơn đồng (Cu) nên thứ tự giảm dần phải là: Zn, Cu."
            },
            {
                "label": "b",
                "text": "Phản ứng hóa học xảy ra trong quá trình trên là: Zn + CuSO~4~ → ZnSO~4~ + Cu.",
                "correct": True,
                "explain": "Đúng. Kẽm đứng trước đồng trong dãy hoạt động hóa học nên đẩy được đồng ra khỏi dung dịch muối: Zn + CuSO~4~ → ZnSO~4~ + Cu."
            },
            {
                "label": "c",
                "text": "Kim loại kẽm đã đẩy được đồng ra khỏi dung dịch muối vì kẽm hoạt động hóa học yếu hơn đồng.",
                "correct": False,
                "explain": "Sai. Kẽm đẩy được đồng ra khỏi dung dịch muối vì kẽm có mức độ hoạt động hóa học MẠNH hơn đồng."
            },
            {
                "label": "d",
                "text": "Trong quá trình phản ứng, kẽm tan dần, màu xanh lam của dung dịch ban đầu nhạt dần rồi mất hẳn do tạo thành dung dịch ZnSO~4~ không màu và có kim loại đồng (Cu) màu đỏ nâu bám trên bề mặt kẽm.",
                "correct": True,
                "explain": "Đúng. Do bột kẽm dư nên toàn bộ ion Cu^2+^ (màu xanh lam) phản ứng hết chuyển thành Zn^2+^ không màu, đồng giải phóng có màu đỏ nâu bám vào bề mặt kẽm."
            }
        ]
    }
]

# -----------------------------------------------------------------------------
# 3. Trắc nghiệm trả lời ngắn (2,0 điểm; câu 15 - câu 16) - Mỗi câu 1,0 điểm
# -----------------------------------------------------------------------------
PART3_QUESTIONS = [
    {
        "id": 15,
        "sub_questions": [
            {
                "sub_id": "a",
                "content": "Cho các kim loại sau: Mg, Cu, Ag, Fe. Có bao nhiêu kim loại đẩy được bạc (Ag) ra khỏi dung dịch muối silver nitrate (AgNO~3~)?",
                "correct": "3",
                "solution": "Các kim loại đứng trước Ag trong dãy hoạt động hóa học (không tan trong nước ở điều kiện thường) sẽ đẩy được Ag ra khỏi dung dịch muối AgNO~3~. Đó là: Mg, Cu, Fe (tổng cộng 3 kim loại). Kim loại Ag không phản ứng với muối của chính nó."
            },
            {
                "sub_id": "b",
                "content": "Trong tự nhiên, các kim loại thường tồn tại dưới dạng hợp chất, quặng,... Để tách lấy kim loại ra khỏi hợp chất, ta thường dùng một trong các phương pháp sau: (1) Phương pháp nhiệt luyện; (2) Phương pháp điện phân; (3) Phương pháp thủy luyện. Quặng bauxite chứa một lượng lớn Aluminium dưới dạng Aluminium oxide, quặng sphalerite chứa Zinc dưới dạng Zinc sulfide. Phương pháp để tách lấy Aluminium và Zinc ra khỏi quặng là gì? (Ghi số thứ tự ứng với phương pháp được chọn để tách, theo thứ tự kim loại đã cho: Aluminium, Zinc.)",
                "correct": "(2); (1)",
                "solution": "- Aluminium (Al) là kim loại hoạt động hóa học mạnh, được điều chế bằng phương pháp điện phân nóng chảy Al~2~O~3~ (phương pháp số 2).\n- Zinc (Zn) là kim loại hoạt động trung bình, được điều chế bằng phương pháp nhiệt luyện (sau khi nung quặng ZnS thành ZnO thì khử bằng than cốc C hoặc khí CO ở nhiệt độ cao - phương pháp số 1).\nVậy thứ tự tương ứng là: (2); (1) (hoặc 2; 1)."
            }
        ]
    },
    {
        "id": 16,
        "sub_questions": [
            {
                "sub_id": "a",
                "content": "Bạn Bình dùng các kí hiệu để mô tả thí nghiệm lai một cặp tính trạng ở đậu Hà Lan của Mendel như sau:\nP~tc~: Cây hạt vàng , cây hạt xanh;\nF~1~: 100% cây hạt vàng;\nF~2~: 3 cây hạt vàng : 1 cây hạt xanh.\nBiết rằng cây hạt vàng thuần chủng được lai với cây hạt xanh thuần chủng, F~1~ thu được 100% cây hạt vàng. Cho các cây F~1~ tự thụ phấn, F~2~ thu được cây hạt vàng và cây hạt xanh với tỉ lệ xấp xỉ 3 : 1. Có bao nhiêu lỗi sai trong cách sử dụng kí hiệu của bạn Bình?",
                "correct": "3",
                "solution": "Có 3 lỗi sai chính trong cách dùng kí hiệu di truyền học của bạn Bình:\n1. Ở thế hệ P: Giữa hai cơ thể bố và mẹ đem lai phải dùng dấu phép lai \"×\" (viết là P: Cây hạt vàng × Cây hạt xanh), bạn Bình dùng dấu phẩy \",\" hoặc không có dấu phép lai.\n2. Bỏ qua bước tự thụ phấn của F~1~: Để tạo ra F~2~ thì các cây F~1~ phải tự thụ phấn (kí hiệu dấu ⊗ hoặc F~1~ × F~1~), bạn Bình không thể hiện phép tự thụ phấn này.\n3. Ở tỉ lệ phân li thế hệ F~2~: F~2~ thể hiện tỉ lệ kiểu hình phân li là 3 hạt vàng : 1 hạt xanh (hoặc 3 vàng : 1 xanh), bạn Bình viết kèm đơn vị cá thể \"3 cây... : 1 cây...\" là sai bản chất của tỉ lệ phân li thống kê.\n=> Đáp số: 3 lỗi sai."
            },
            {
                "sub_id": "b",
                "content": "Ở đậu Hà Lan, allele A quy định hạt vàng trội hoàn toàn so với allele a quy định hạt xanh; allele B quy định vỏ trơn trội hoàn toàn so với allele b quy định vỏ nhăn (hai cặp gene nằm trên hai cặp nhiễm sắc thể tương đồng khác nhau, phân li độc lập). Cho cây bố mẹ thuần chủng hạt vàng, vỏ trơn lai với cây hạt xanh, vỏ nhăn thu được F~1~. Cho các cây F~1~ tự thụ phấn thu được F~2~. Theo lí thuyết, cây hạt xanh, vỏ nhăn ở F~2~ chiếm tỉ lệ bao nhiêu? (Viết kết quả dưới dạng phân số.)",
                "correct": "1/16",
                "solution": "Quy ước: A: hạt vàng >> a: hạt xanh; B: vỏ trơn >> b: vỏ nhăn.\n- P thuần chủng: AABB (hạt vàng, trơn) × aabb (hạt xanh, nhăn) → F~1~: 100% AaBb (hạt vàng, trơn).\n- F~1~ tự thụ phấn: AaBb × AaBb.\n- Vì 2 cặp gene phân li độc lập, xét riêng từng cặp tính trạng:\n  + Aa × Aa → F~2~: 3/4 A- (vàng) : 1/4 aa (xanh).\n  + Bb × Bb → F~2~: 3/4 B- (trơn) : 1/4 bb (nhăn).\n- Tỉ lệ cây mang kiểu hình hạt xanh, vỏ nhăn (kiểu gene aabb) ở F~2~ là:\n  Tỉ lệ = 1/4 (aa) × 1/4 (bb) = 1/16."
            }
        ]
    }
]

# -----------------------------------------------------------------------------
# PHẦN II. TỰ LUẬN (3,0 ĐIỂM)
# -----------------------------------------------------------------------------
ESSAY_QUESTIONS = [
    {
        "id": 17,
        "score": "1,0 điểm",
        "sub_questions": [
            {
                "sub_id": "a",
                "score": "0,5 điểm",
                "content": "Hãy liệt kê hai đơn vị thường dùng để đo công cơ học và hai đơn vị thường dùng để đo công suất.",
                "solution": "- Hai đơn vị thường dùng để đo công cơ học là: Jun (J) và kilôjun (kJ) (hoặc calo, kilôoát giờ - kW.h).\n- Hai đơn vị thường dùng để đo công suất là: Oát (W) và kilôoát (kW) (hoặc mêgaoát - MW, mã lực - HP)."
            },
            {
                "sub_id": "b",
                "score": "0,5 điểm",
                "content": "Một vật có khối lượng m được thả rơi tự do không tốc độ ban đầu từ độ cao h so với mặt đất. Bỏ qua mọi lực cản của không khí và chọn gốc thế năng tại mặt đất. Trong quá trình vật rơi, động năng và thế năng của vật thay đổi như thế nào? Vì sao?",
                "solution": "- Trong quá trình vật rơi:\n  + Độ cao h của vật giảm dần nên thế năng trọng trường (W~t~ = 10.m.h) giảm dần.\n  + Tốc độ v của vật tăng dần nên động năng (W~đ~ = 1/2.m.v^2^) tăng dần.\n- Vì sao (Giải thích): Vì bỏ qua lực cản của không khí nên cơ năng của vật được bảo toàn (W = W~đ~ + W~t~ = không đổi). Do đó, thế năng của vật đã chuyển hóa hoàn toàn thành động năng trong suốt quá trình rơi (thế năng giảm bao nhiêu thì động năng tăng bấy nhiêu)."
            }
        ]
    },
    {
        "id": 18,
        "score": "1,0 điểm",
        "sub_questions": [
            {
                "sub_id": "a",
                "score": "0,5 điểm",
                "content": "Nêu các tính chất vật lí chung của kim loại.",
                "solution": "Các tính chất vật lí chung của kim loại gồm có:\n1. Tính dẻo (dễ rèn, kéo sợi, dát mỏng thành lá mỏng).\n2. Tính dẫn điện tốt.\n3. Tính dẫn nhiệt tốt.\n4. Có ánh kim đặc trưng.\n(Các tính chất vật lí chung này đều do các electron tự do chuyển động trong mạng tinh thể kim loại gây ra)."
            },
            {
                "sub_id": "b",
                "score": "0,5 điểm",
                "content": "Cho khí carbon monoxide dư đi qua 2 kg quặng hematite (chứa 80% iron(III) oxide về khối lượng) thì thu được m gam iron. Biết hiệu suất phản ứng là 60%. Tính giá trị m. (Cho Fe = 56 amu; O = 16 amu.)",
                "solution": "- Khối lượng Fe~2~O~3~ nguyên chất có trong 2 kg quặng hematite:\n  m(Fe~2~O~3~) = 2 kg × 80% = 1,6 kg = 1 600 gam.\n- Khối lượng mol phân tử của Fe~2~O~3~: M(Fe~2~O~3~) = 56 × 2 + 16 × 3 = 160 g/mol.\n- Số mol Fe~2~O~3~ đem phản ứng:\n  n(Fe~2~O~3~) = 1 600 / 160 = 10 mol.\n- Phương trình hóa học:\n  Fe~2~O~3~ + 3CO  → (t°)  2Fe + 3CO~2~\n- Theo phương trình hóa học, số mol Fe thu được theo lý thuyết:\n  n(Fe lí thuyết) = 2 × n(Fe~2~O~3~) = 2 × 10 = 20 mol.\n- Khối lượng Fe thu được theo lý thuyết:\n  m(Fe lí thuyết) = 20 × 56 = 1 120 gam.\n- Vì hiệu suất phản ứng H = 60%, khối lượng sắt thực tế thu được là:\n  m = m(Fe lí thuyết) × (60 / 100) = 1 120 × 0,6 = 672 gam."
            }
        ]
    },
    {
        "id": 19,
        "score": "1,0 điểm",
        "sub_questions": [
            {
                "sub_id": "a",
                "score": "0,5 điểm",
                "content": "Nếu cơ thể mang tính trạng trội không thuần chủng đem lai phân tích thì phép lai phân tích có kết quả như thế nào?",
                "solution": "- Cơ thể mang tính trạng trội không thuần chủng có kiểu gene dị hợp (Aa).\n- Phép lai phân tích là phép lai giữa cá thể mang tính trạng trội với cá thể mang tính trạng lặn (aa).\n- Sơ đồ lai phân tích:\n  P: Aa (trội) × aa (lặn)\n  G: (1/2 A : 1/2 a) × a\n  F~a~: 1/2 Aa (kiểu hình trội) : 1/2 aa (kiểu hình lặn).\n- Kết luận: Kết quả của phép lai phân tích là đời con bị phân tính theo tỉ lệ 1 trội : 1 lặn (50% trội : 50% lặn)."
            },
            {
                "sub_id": "b",
                "score": "0,5 điểm",
                "content": "Trong công tác điều tra pháp y, cơ quan chức năng có thể xác định chính xác danh tính của một cá nhân thông qua mẫu DNA. Mặc dù phân tử DNA của tất cả mọi người đều được cấu tạo từ 4 loại nucleotide (A, T, G, C), nhưng trình tự DNA ở mỗi người lại có sự khác nhau. Em hãy giải thích vì sao chỉ từ 4 loại nucleotide này lại có thể tạo nên sự đặc trưng của các phân tử DNA ở mỗi người.",
                "solution": "- Phân tử DNA được cấu tạo theo nguyên tắc đa phân từ 4 loại đơn phân là các nucleotide (A, T, G, C).\n- Mặc dù chỉ gồm 4 loại nucleotide, nhưng tính đa dạng và tính đặc thù của phân tử DNA ở mỗi người được quy định bởi:\n  1. Số lượng các nucleotide trong phân tử.\n  2. Thành phần tỉ lệ giữa các loại nucleotide.\n  3. Quan trọng nhất là trật tự sắp xếp (trình tự các nucleotide) trong chuỗi xoắn kép DNA.\n- Với hàng triệu đến hàng tỉ nucleotide sắp xếp theo các tổ hợp trình tự vô cùng phong phú, mỗi cá nhân sẽ sở hữu một trình tự DNA mang tính đặc thù riêng biệt (ngoại trừ trường hợp sinh đôi cùng trứng). Do đó, phân tích trình tự DNA cho phép cơ quan pháp y định danh chính xác từng cá nhân."
            }
        ]
    }
]
