# -*- coding: utf-8 -*-
"""
Dữ liệu Đề cương ôn tập Giữa kì I - KHTN 6: PHẦN CHUNG
Trường THCS Kim Long - Năm học 2026 - 2027
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

INFO = {
    "school": "TRƯỜNG THCS KIM LONG",
    "year": "NĂM HỌC 2026 - 2027",
    "title": "ĐỀ CƯƠNG ÔN TẬP GIỮA KÌ I – KHTN 6",
    "subject": "MÔN: KHOA HỌC TỰ NHIÊN 6",
    "time_str": "Chương trình GDPT 2018 - Trọn bộ Đề cương Ôn tập Giữa kì I",
}

PHAN_CHUNG_THEORY = [
    "I. Lí thuyết: Học sinh học từ: Bài 1. Giới thiệu về khoa học tự nhiên đến bài 8. Đo nhiệt độ.",
    "II. Bài tập: Học sinh làm tất cả các bài tập trong sách bài tập từ: Bài 1. Giới thiệu về khoa học tự nhiên đến bài 8. Đo nhiệt độ."
]

# -----------------------------------------------------------------------------
# PHẦN CHUNG - 1. TRẮC NGHIỆM NHIỀU LỰA CHỌN (28 CÂU)
# -----------------------------------------------------------------------------
PHAN_CHUNG_MCQ = [
    {
        "id": 1,
        "content": "Hoạt động nào sau đây của con người là hoạt động nghiên cứu khoa học?",
        "options": [
            "A. Trồng hoa với quy mô lớn trong nhà kính",
            "B. Nghiên cứu vaccine phòng chống virus corona trong phòng thí nghiệm",
            "C. Sản xuất muối ăn từ nước biển bằng phương pháp phơi cát",
            "D. Vận hành nhà máy thủy điện để sản xuất điện"
        ],
        "correct": "B",
        "solution": "Nghiên cứu khoa học là hoạt động của con người chủ động tìm tòi, khám phá ra tri thức mới (nghiên cứu cơ chế và điều chế vaccine phòng bệnh). Các hoạt động còn lại là hoạt động sản xuất, kinh doanh và ứng dụng công nghệ."
    },
    {
        "id": 2,
        "content": "Hoạt động nào sau đây của con người không phải là hoạt động nghiên cứu khoa học?",
        "options": [
            "A. Theo dõi nuôi cấy mô cây trồng trong phòng thí nghiệm",
            "B. Làm thí nghiệm điều chế chất mới",
            "C. Lấy mẫu đất để phân loại cây trồng",
            "D. Sản xuất phân bón hóa học"
        ],
        "correct": "D",
        "solution": "Sản xuất phân bón hóa học là hoạt động sản xuất kinh doanh theo quy trình công nghệ sẵn có, không phải là hoạt động tìm tòi, khám phá quy luật mới (nghiên cứu khoa học)."
    },
    {
        "id": 3,
        "content": "Vật nào sau đây gọi là vật không sống?",
        "options": ["A. Con ong", "B. Vi khuẩn", "C. Than củi", "D. Cây cam"],
        "correct": "C",
        "solution": "Than củi không có các dấu hiệu đặc trưng của sự sống như: trao đổi chất và năng lượng, lớn lên, sinh sản, cảm ứng."
    },
    {
        "id": 4,
        "content": "Việc làm nào sau đây được cho là không an toàn trong phòng thực hành?",
        "options": [
            "A. Đeo găng tay khi lấy hoá chất.",
            "B. Tự ý làm các thí nghiệm.",
            "C. Sử dụng kính bảo vệ mắt khi làm thí nghiệm.",
            "D. Rửa tay trước khi ra khỏi phòng thực hành."
        ],
        "correct": "B",
        "solution": "Tự ý làm các thí nghiệm khi chưa có sự cho phép hay hướng dẫn của giáo viên có thể dẫn đến nguy cơ cháy nổ, ngộ độc hóa chất và mất an toàn nghiêm trọng."
    },
    {
        "id": 5,
        "content": "Khi gặp sự cố mất an toàn trong phòng thực hành, em cần",
        "options": [
            "A. báo cáo ngay với giáo viên trong phòng thực hành.",
            "B. tự xử lí và không thông báo với giáo viên.",
            "C. nhờ bạn xử lí sự cố.",
            "D. tiếp tục làm thí nghiệm."
        ],
        "correct": "A",
        "solution": "Khi gặp bất kỳ sự cố mất an toàn nào (đổ vỡ dụng cụ thủy tinh, tràn hóa chất, chập điện, bỏng, cháy...), cần báo cáo ngay với giáo viên để được hướng dẫn và xử lí đúng quy trình kĩ thuật."
    },
    {
        "id": 6,
        "content": "Khoa học tự nhiên không bao gồm lĩnh vực nào sau đây?",
        "options": ["A. Vật lí học", "B. Khoa học Trái Đất", "C. Thiên văn học", "D. Tâm lí học"],
        "correct": "D",
        "solution": "Khoa học tự nhiên nghiên cứu các sự vật, hiện tượng, quy luật của thế giới tự nhiên (Vật lí, Hóa học, Sinh học, Thiên văn học, Khoa học Trái Đất). Tâm lí học nghiên cứu hành vi và tâm trí con người, thuộc lĩnh vực khoa học xã hội."
    },
    {
        "id": 7,
        "content": "Em đang đun nước, sau một thời gian thấy tiếng nước reo và mặt nước sủi lăn tăn, nước bắt đầu sôi. Vậy hiện tượng nước sôi liên quan tới lĩnh vực khoa học nào?",
        "options": ["A. Hóa học", "B. Vật lí học", "C. Sinh học", "D. Hóa học và sinh học"],
        "correct": "B",
        "solution": "Hiện tượng nước sôi và chuyển từ thể lỏng sang thể hơi (hơi nước) là sự biến đổi trạng thái vật lí, thuộc đối tượng nghiên cứu của Vật lí học."
    },
    {
        "id": 8,
        "content": "Hoạt động nào sau đây không được xem là nghiên cứu khoa học tự nhiên?",
        "options": [
            "A. Nghiên cứu hệ thống quạt nước cho đầm nuôi tôm.",
            "B. Nghiên cứu trang phục của các nước.",
            "C. Nghiên cứu xử lí rác thải bảo vệ môi trường.",
            "D. Nghiên cứu cách khắc chữ lên thủy tinh."
        ],
        "correct": "B",
        "solution": "Nghiên cứu trang phục truyền thống của các nước thuộc lĩnh vực văn hóa, lịch sử và xã hội học, không thuộc khoa học tự nhiên."
    },
    {
        "id": 9,
        "content": "Để đảm bảo an toàn trong phòng thực hành cần thực hiện nguyên tắc nào dưới đây?",
        "options": [
            "A. Làm thí nghiệm theo sự hướng dẫn của bàn bè trong lớp.",
            "B. Có thể nhận biết hóa chất bằng cách ngửi hóa chất.",
            "C. Mang đồ ăn vào phòng thực hành.",
            "D. Đọc kĩ nội quy và thực hiện theo nội quy phòng thực hành."
        ],
        "correct": "D",
        "solution": "Đọc kĩ và nghiêm túc tuân thủ nội quy phòng thực hành là nguyên tắc bắt buộc hàng đầu để đảm bảo an toàn cho bản thân và tập thể."
    },
    {
        "id": 10,
        "content": "Cách sử dụng kính lúp nào sau đây là đúng?",
        "options": [
            "A. Đặt kính ở khoảng sao cho nhìn thấy vật rõ nét, mắt nhìn vào mặt kính.",
            "B. Đặt kính cách xa mắt, mắt nhìn vào mặt kính.",
            "C. Đặt kính ở khoảng 20 cm, mắt nhìn vào mặt kính.",
            "D. Đặt kính trong khoảng mắt không phải điều tiết, mắt nhìn vào mặt kính."
        ],
        "correct": "A",
        "solution": "Cách quan sát đúng: Đặt mặt kính lúp gần sát vật mẫu, mắt nhìn vào mặt kính lúp, sau đó dịch chuyển kính lúp dần dần cho đến khi quan sát được hình ảnh phóng to rõ nét nhất."
    },
    {
        "id": 11,
        "content": "Người ta sử dụng kính hiển vi để quan sát:",
        "options": ["A. Hồng cầu", "B. Mặt Trăng", "C. Máy bay", "D. Con kiến"],
        "correct": "A",
        "solution": "Tế bào hồng cầu có kích thước hiển vi (khoảng 7 - 8 µm) nên bắt buộc phải dùng kính hiển vi quang học để quan sát."
    },
    {
        "id": 12,
        "content": "Giới hạn đo của thước là",
        "options": [
            "A. độ dài lớn nhất ghi trên thước.",
            "B. độ dài giữa hai vạch chia liên tiếp trên thước.",
            "C. độ dài nhỏ nhất ghi trên thước.",
            "D. độ dài giữa hai vạch bất kỳ ghi trên thước."
        ],
        "correct": "A",
        "solution": "Giới hạn đo (GHĐ) của thước là độ dài lớn nhất ghi trên thước. Độ dài giữa hai vạch chia liên tiếp gọi là độ chia nhỏ nhất (ĐCNN)."
    },
    {
        "id": 13,
        "content": "Dụng cụ không được sử dụng để đo chiều dài là",
        "options": ["A. Thước dây.", "B. Thước mét.", "C. Thước kẹp.", "D. Compa."],
        "correct": "D",
        "solution": "Compa là dụng cụ dùng để vẽ đường tròn hoặc so sánh khoảng cách hình học, không có vạch chia độ dài trực tiếp để đo chiều dài."
    },
    {
        "id": 14,
        "content": "Thước thích hợp để đo bề dày quyển sách Khoa học tự nhiên 6 là",
        "options": [
            "A. thước kẻ có giới hạn đo 10 cm và độ chia nhỏ nhất 1 mm.",
            "B. thước dây có giới hạn đo 1 m và độ chia nhỏ nhất 1 cm.",
            "C. thước cuộn có giới hạn đo 3 m và độ chia nhỏ nhất 5 cm.",
            "D. thước thẳng có giới hạn đo 1,5 m và độ chia nhỏ nhất 1 cm."
        ],
        "correct": "A",
        "solution": "Bề dày quyển sách Khoa học tự nhiên 6 chỉ vào khoảng 0,5 - 1 cm. Do đó, thước kẻ có GHĐ 10 cm và ĐCNN 1 mm là dụng cụ thích hợp và chính xác nhất."
    },
    {
        "id": 15,
        "content": "Xác định giới hạn đo và độ chia nhỏ nhất của thước trong hình dưới đây:",
        "image": "khtn6_ruler_p2.png",
        "options": [
            "A. GHĐ 10 cm, ĐCNN 1 mm.",
            "B. GHĐ 20 cm, ĐCNN 1 cm.",
            "C. GHĐ 100 cm, ĐCNN 1 cm.",
            "D. GHĐ 10 cm, ĐCNN 0,5 cm."
        ],
        "correct": "D",
        "solution": "- Số lớn nhất ghi trên thước là 10 cm, vậy GHĐ = 10 cm.\n- Từ vạch 0 đến vạch 1 cm được chia làm 2 khoảng đều nhau bởi một vạch ở giữa, do đó độ dài mỗi khoảng nhỏ là: 1 cm / 2 = 0,5 cm (hay 5 mm), vậy ĐCNN = 0,5 cm."
    },
    {
        "id": 16,
        "content": "Khối lượng của một vật chỉ",
        "options": [
            "A. chiều dài của vật đó.",
            "B. sức nặng của vật đó.",
            "C. độ lớn của vật đó.",
            "D. một đơn vị thể tích của vật đó."
        ],
        "correct": "B",
        "solution": "Theo định nghĩa SGK Khoa học tự nhiên 6, khối lượng là số đo lượng chất tạo nên vật (chỉ sức nặng của vật đó)."
    },
    {
        "id": 17,
        "content": "Để cân 1 kg đường, ta chọn chiếc cân có",
        "options": [
            "A. ĐCNN 1 g và GHĐ 1 kg.",
            "B. ĐCNN 50 g và GHĐ 10 kg.",
            "C. ĐCNN 100 g và GHĐ 10 kg.",
            "D. ĐCNN 1 kg và GHĐ 100 kg."
        ],
        "correct": "B",
        "solution": "Khi cân vật nặng 1 kg, cần chọn cân có GHĐ lớn hơn khối lượng cần đo một lượng vừa phải (như GHĐ 5 kg hoặc 10 kg) và ĐCNN phù hợp (50 g) để đảm bảo không bị vượt quá giới hạn đo của cân."
    },
    {
        "id": 18,
        "content": "Trên vỏ một hộp thịt có ghi 500g. Số liệu đó chỉ",
        "options": [
            "A. thể tích của cả hộp thịt.",
            "B. thể tích của thịt trong hộp.",
            "C. khối lượng của cả hộp thịt.",
            "D. khối lượng của thịt trong hộp."
        ],
        "correct": "D",
        "solution": "Con số 500g ghi trên bao bì là khối lượng tịnh của thực phẩm chứa bên trong hộp (không bao gồm vỏ hộp)."
    },
    {
        "id": 19,
        "content": "Cân một túi hoa quả cho kết quả là 1553g. ĐCNN của cân đã dùng là",
        "options": ["A. 5 g", "B. 100 g", "C. 10 g", "D. 1 g"],
        "correct": "D",
        "solution": "Kết quả đo luôn là bội số của ĐCNN. Số 1553 chia hết cho 1 g (không chia hết cho 5 g, 10 g, 100 g), do đó ĐCNN của cân là 1 g."
    },
    {
        "id": 20,
        "content": "Để xác định thành tích của vận động viên chạy 100 m người ta phải sử dụng loại đồng hồ là",
        "options": [
            "A. đồng hồ mặt trời.",
            "B. đồng hồ điện tử.",
            "C. đồng hồ bấm giây.",
            "D. đồng hồ đeo tay."
        ],
        "correct": "C",
        "solution": "Đồng hồ bấm giây là dụng cụ chuyên dụng để đo các khoảng thời gian ngắn trong thi đấu thể thao với ĐCNN rất nhỏ (0,1s hoặc 0,01s)."
    },
    {
        "id": 21,
        "content": "Minh đi từ nhà đến bến xe mất 20 phút. Sau đó, Minh đi xe buýt đến công viên mất 45 phút. Hỏi Minh đi từ nhà đến công viên mất bao nhiêu giờ?",
        "options": [
            "A. 65 phút.",
            "B. 1 giờ 20 phút.",
            "C. 1 giờ 05 phút.",
            "D. 1 giờ 75 phút."
        ],
        "correct": "C",
        "solution": "Tổng thời gian di chuyển = 20 phút + 45 phút = 65 phút = 1 giờ 05 phút."
    },
    {
        "id": 22,
        "content": "Lúc 7 giờ 45 phút, một xe máy đi từ A đến B. Biết xe máy đi từ A đến B mất 2 giờ 30 phút. Xe máy đến B lúc",
        "options": [
            "A. 10 giờ 55 phút.",
            "B. 9 giờ 55 phút.",
            "C. 10 giờ 10 phút.",
            "D. 10 giờ 15 phút."
        ],
        "correct": "D",
        "solution": "Thời điểm xe máy đến B là: 7 giờ 45 phút + 2 giờ 30 phút = 9 giờ 75 phút = 10 giờ 15 phút."
    },
    {
        "id": 23,
        "content": "Dụng cụ đo nhiệt độ là",
        "options": ["A. đồng hồ.", "B. nhiệt kế.", "C. cân.", "D. thước kẻ."],
        "correct": "B",
        "solution": "Nhiệt kế là dụng cụ dùng để đo nhiệt độ."
    },
    {
        "id": 24,
        "content": "Đơn vị đo nhiệt độ nào sau đây thường dùng ở Việt Nam?",
        "options": ["A. độ F (°F).", "B. độ C (°C).", "C. kelvin (K).", "D. cả 3 đơn vị trên."],
        "correct": "B",
        "solution": "Ở Việt Nam, thang nhiệt độ Celsius với đơn vị là độ C (°C) được sử dụng phổ biến và chính thức trong đời sống hàng ngày."
    },
    {
        "id": 25,
        "content": "Nhiệt kế hoạt động dựa trên nguyên tắc",
        "options": [
            "A. sự nở vì nhiệt của chất rắn khác nhau.",
            "B. sự nở vì nhiệt của các chất khí khác nhau.",
            "C. sự nở vì nhiệt của các chất lỏng khác nhau.",
            "D. sự nở vì nhiệt trong môi trường chân không."
        ],
        "correct": "C",
        "solution": "Các loại nhiệt kế thông dụng (nhiệt kế thủy ngân, nhiệt kế rượu, nhiệt kế dầu) hoạt động dựa trên hiện tượng giãn nở vì nhiệt của chất lỏng."
    },
    {
        "id": 26,
        "content": "Không thể dùng nhiệt kế rượu để đo nhiệt độ của hơi nước đang sôi vì",
        "options": [
            "A. rượu sôi ở nhiệt độ cao hơn 100°C.",
            "B. rượu sôi ở nhiệt độ thấp hơn 100°C.",
            "C. rượu đông đặc ở nhiệt độ thấp hơn 100°C.",
            "D. rượu đông đặc ở nhiệt độ cao hơn 100°C."
        ],
        "correct": "B",
        "solution": "Rượu sôi ở khoảng 78,3°C (thấp hơn nhiệt độ sôi của nước là 100°C). Nếu nhúng vào nước sôi, rượu bên trong sẽ sôi, hóa hơi làm bung vỡ nhiệt kế."
    },
    {
        "id": 27,
        "content": "Nhiệt độ mùa đông tại thành phố New York (Mĩ) là 23°F. Ứng với nhiệt giai Celsius, nhiệt độ đó là",
        "options": ["A. 10°C.", "B. 5°C.", "C. -5°C.", "D. -10°C."],
        "correct": "C",
        "solution": "Công thức chuyển đổi nhiệt độ: T(°C) = (T(°F) - 32) / 1,8\nThay số: T(°C) = (23 - 32) / 1,8 = -9 / 1,8 = -5°C."
    },
    {
        "id": 28,
        "content": "Nhiệt độ của nước trong phòng theo nhiệt giai Celsius là 27°C. Ứng với nhiệt giai Fahrenheit, nhiệt độ này là",
        "options": ["A. 48,6°F.", "B. 80,6°F.", "C. 15°F.", "D. 47°F."],
        "correct": "B",
        "solution": "Công thức chuyển đổi: T(°F) = T(°C) × 1,8 + 32\nThay số: T(°F) = 27 × 1,8 + 32 = 48,6 + 32 = 80,6°F."
    }
]

# -----------------------------------------------------------------------------
# PHẦN CHUNG - 2. CÂU TRẢ LỜI ĐÚNG - SAI (8 CÂU)
# -----------------------------------------------------------------------------
PHAN_CHUNG_TF = [
    {
        "id": 1,
        "content": "Trong các phát biểu sau về một số quy định an toàn trong phòng thực hành. Phát biểu nào đúng, phát biểu nào sai?",
        "sub_items": [
            {"id": "a", "text": "Chỉ tiến hành thí nghiệm khi có người hướng dẫn.", "answer": True, "reason": "Bắt buộc phải có giáo viên hoặc người hướng dẫn để kiểm soát an toàn."},
            {"id": "b", "text": "Mặc trang phục gọn gàng, nữ buộc tóc cao, đeo găng tay, khẩu trang, kính bảo vệ mắt và các thiết bị bảo hộ khác (nếu cần thiết).", "answer": True, "reason": "Trang phục bảo hộ giúp phòng tránh hóa chất bắn vào mắt, da hay vướng víu vào dụng cụ thí nghiệm."},
            {"id": "c", "text": "Có thể nếm hoặc ngửi hóa chất để nhận biết hóa chất.", "answer": False, "reason": "Tuyệt đối không được nếm hoặc ngửi trực tiếp bất kỳ hóa chất nào vì nguy cơ nhiễm độc đường thở và tiêu hóa."},
            {"id": "d", "text": "Hóa chất còn thừa sau thí nghiệm có thể xả xuống cống thoát nước.", "answer": False, "reason": "Hóa chất thừa phải đổ vào bình thu gom chất thải riêng theo quy định, không được xả tùy tiện vào bồn rửa hay cống thoát nước."}
        ]
    },
    {
        "id": 2,
        "content": "Trong các phát biểu sau, phát biểu nào đúng, phát biểu nào sai?",
        "sub_items": [
            {"id": "a", "text": "Để đo chiều dài cuốn sách giáo khoa dùng thước kẹp (thước cặp).", "answer": False, "reason": "Chiều dài sách giáo khoa khoảng 24 - 27 cm, chỉ cần dùng thước kẻ thẳng có ĐCNN 1 mm là thuận tiện và chính xác."},
            {"id": "b", "text": "Để đo chiều rộng phòng học sử dụng thước cuộn.", "answer": True, "reason": "Phòng học có kích thước lớn (khoảng 6 - 8 m), dùng thước cuộn có GHĐ 5 m - 10 m là phù hợp nhất."},
            {"id": "c", "text": "Để đo chiều cao của tủ sách sử dụng thước thẳng.", "answer": False, "reason": "Tủ sách cao khoảng 1,8 - 2 m, thước thẳng học sinh thường chỉ dài 0,5 - 1 m khó đo chính xác, nên dùng thước dây hoặc thước cuộn."},
            {"id": "d", "text": "Để đo đường kính trong của miệng một cái cốc hình trụ sử dụng thước kẹp.", "answer": True, "reason": "Thước kẹp có mỏ kẹp trong chuyên dụng để đo chính xác đường kính trong của các ống tròn, miệng cốc."}
        ]
    },
    {
        "id": 3,
        "content": "Một học sinh muốn đo chiều dài của một vật trong phòng thí nghiệm bằng các loại thước khác nhau để chọn được dụng cụ đo phù hợp nhất. Theo lý thuyết, phát biểu nào sau đây đúng hay sai?",
        "sub_items": [
            {"id": "a", "text": "Khi cần đo các vật có kích thước nhỏ và yêu cầu độ chính xác cao, nên sử dụng thước cuộn.", "answer": False, "reason": "Thước cuộn có ĐCNN lớn (thường là 1 mm hoặc 5 mm), khi đo vật nhỏ chính xác cao nên dùng thước kẹp hoặc panme."},
            {"id": "b", "text": "Thước cuộn là lựa chọn tốt để đo chiều dài các vật lớn và có hình dạng thẳng.", "answer": True, "reason": "Thước cuộn có GHĐ lớn (3m - 10m), dây thước kim loại thẳng giúp đo các khoảng cách dài rất tốt."},
            {"id": "c", "text": "Các thông số như giới hạn đo và độ chia nhỏ nhất trên thước hỗ trợ người dùng đánh giá mức độ phù hợp của thước và giúp hạn chế sai số đo.", "answer": True, "reason": "Việc nắm rõ GHĐ và ĐCNN giúp chọn thước phù hợp với vật cần đo, giảm thiểu sai số đo."},
            {"id": "d", "text": "Thước dây có độ chính xác cao nhất vì có thể dễ dàng cuộn lại và bám sát vào bề mặt vật đo.", "answer": False, "reason": "Thước dây làm bằng vải hoặc nhựa mềm dễ bị co giãn theo thời gian và lực kéo nên độ chính xác không cao bằng thước kẹp hay thước kim loại."}
        ]
    },
    {
        "id": 4,
        "content": "Các phát biểu về việc chọn đúng loại đơn vị và thước đo sau đây đúng hay sai?",
        "sub_items": [
            {"id": "a", "text": "Để đo chiều dài của một vật có kích thước nhỏ, người ta nên chọn đơn vị đo là centimet.", "answer": False, "reason": "Vật có kích thước nhỏ (như độ dày lá cây, đường kính sợi chỉ) nên chọn đơn vị đo là milimet (mm) hoặc micromet (µm)."},
            {"id": "b", "text": "Mét (m) được xem là đơn vị đo phổ biến nhất khi đo chiều dài trong đời sống và kỹ thuật.", "answer": True, "reason": "Mét là đơn vị cơ bản đo chiều dài trong hệ đo lường quốc tế SI và được sử dụng rộng rãi nhất."},
            {"id": "c", "text": "ĐCNN của thước càng nhỏ, khả năng đo chính xác càng cao, nên người dùng thường chọn thước có ĐCNN nhỏ nhất cho tất cả các phép đo.", "answer": False, "reason": "Cần chọn thước có ĐCNN và GHĐ phù hợp với mục đích đo; không phải lúc nào cũng cần ĐCNN nhỏ nhất vì dễ gây phức tạp và mất thời gian."},
            {"id": "d", "text": "Khi đo chiều dài lớn, người ta có thể sử dụng đơn vị thay mét vì kilômét (km) để thuận tiện hơn.", "answer": True, "reason": "Khoảng cách lớn giữa các địa điểm hay chiều dài tuyến đường dùng đơn vị kilômét (km) giúp số liệu gọn gàng, dễ biểu diễn."}
        ]
    },
    {
        "id": 5,
        "content": "Trong các phát biểu sau về đo khối lượng, phát biểu nào đúng, phát biểu nào sai?",
        "sub_items": [
            {"id": "a", "text": "Đơn vị đo khối lượng trong hệ thống đo lường chính thức ở nước ta là gam.", "answer": False, "reason": "Đơn vị đo khối lượng hợp pháp trong hệ thống đo lường chính thức ở Việt Nam là kilôgam (kg)."},
            {"id": "b", "text": "Khối lượng là số đo lượng chất của vật.", "answer": True, "reason": "Đây là định nghĩa chuẩn về khối lượng trong SGK KHTN 6."},
            {"id": "c", "text": "1 tạ = 100 kg = 1000 gam.", "answer": False, "reason": "1 tạ = 100 kg = 100 000 gam (chứ không phải 1000 gam)."},
            {"id": "d", "text": "Dụng cụ thường dùng để đo khối lượng là các loại cân.", "answer": True, "reason": "Cân là dụng cụ chuyên dụng để đo khối lượng."}
        ]
    },
    {
        "id": 6,
        "content": "Một vận động viên muốn tự luyện tập đo thời gian chạy 100 mét bằng đồng hồ đeo tay điện tử. Theo lý thuyết, phát biểu sau đây đúng hay sai?",
        "sub_items": [
            {"id": "a", "text": "Để có kết quả chính xác, vận động viên nên dùng đồng hồ quả lắc.", "answer": False, "reason": "Đồng hồ quả lắc đặt cố định trong phòng, không thể mang theo để bấm thời gian chạy 100m ngoài sân."},
            {"id": "b", "text": "Đồng hồ đeo tay điện tử có thể đo thời gian của các hoạt động ngắn với độ chính xác tương đối.", "answer": True, "reason": "Chức năng bấm giờ của đồng hồ điện tử có độ chia 0,01s đáp ứng tốt nhu cầu luyện tập thể thao."},
            {"id": "c", "text": "Độ chính xác của đồng hồ đeo tay điện tử sẽ luôn cao hơn đồng hồ cát vì nó hoạt động dựa trên cơ chế điện tử.", "answer": True, "reason": "Đồng hồ cát đo thời gian thô sơ bằng dòng cát chảy, sai số rất lớn so với dao động thạch anh của đồng hồ điện tử."},
            {"id": "d", "text": "Nếu vận động viên bắt đầu bấm giờ chậm một chút sau khi xuất phát, thời gian đo được sẽ nhỏ hơn thời gian thực tế.", "answer": True, "reason": "Bấm giờ muộn làm khoảng thời gian được ghi nhận ngắn hơn khoảng thời gian chạy thực tế."}
        ]
    },
    {
        "id": 7,
        "content": "Khi nhúng một tay vào nước lạnh và tay kia vào nước nóng, sau đó cho cả hai tay vào cùng một cốc nước ấm, cảm giác nhiệt độ của hai tay có thể khác nhau. Theo lý thuyết, phát biểu sau đây đúng hay sai?",
        "sub_items": [
            {"id": "a", "text": "Cả hai tay sẽ có cảm giác giống nhau khi đặt vào cùng một cốc nước ấm.", "answer": False, "reason": "Cảm giác nhiệt độ của hai tay bị chi phối bởi trạng thái tiếp xúc trước đó nên sẽ có cảm nhận khác nhau."},
            {"id": "b", "text": "Cảm giác nhiệt độ có thể thay đổi tùy thuộc vào môi trường tiếp xúc trước đó.", "answer": True, "reason": "Thụ quan nhiệt của da nhận biết sự biến thiên nhiệt độ tương đối so với trạng thái vừa tiếp xúc."},
            {"id": "c", "text": "Khi tay vừa nhúng vào nước lạnh, cảm giác khi nhúng vào nước ấm sẽ khiến tay cảm thấy nước ấm nóng hơn thực tế.", "answer": True, "reason": "Vì có sự chênh lệch nhiệt độ lớn truyền nhiệt vào tay lạnh, tạo cảm giác rất nóng."},
            {"id": "d", "text": "Hiện tượng này chứng minh rằng cảm giác của con người có thể không đáng tin cậy để đo lường nhiệt độ chính xác.", "answer": True, "reason": "Giác quan chỉ cho cảm nhận định tính mang tính tương đối, muốn đo chính xác bắt buộc phải dùng nhiệt kế."}
        ]
    },
    {
        "id": 8,
        "content": "Trong một phòng thí nghiệm, một kỹ thuật viên đang sử dụng nhiệt kế thủy ngân để đo nhiệt độ của một dung dịch. Theo lý thuyết, phát biểu sau đây đúng hay sai?",
        "sub_items": [
            {"id": "a", "text": "Nhiệt kế thủy ngân có thể đo nhiệt độ chính xác nhất trong các thí nghiệm yêu cầu độ chính xác cao.", "answer": True, "reason": "Thủy ngân có hệ số giãn nở đều và khoảng đo nhiệt độ rộng nên cho độ chính xác cao."},
            {"id": "b", "text": "Nếu nhiệt độ của dung dịch tăng, mức thủy ngân trong nhiệt kế sẽ giãn nở và đi xuống.", "answer": False, "reason": "Khi nhiệt độ tăng, thủy ngân nở ra làm cột thủy ngân dâng lên (đi lên)."},
            {"id": "c", "text": "Nếu nhiệt kế thủy ngân bị vỡ, cần phải xử lý cẩn thận do thủy ngân là chất độc.", "answer": True, "reason": "Hơi thủy ngân rất độc cho hệ thần kinh và hô hấp, cần rắc bột lưu huỳnh để thu gom an toàn."},
            {"id": "d", "text": "Nhiệt kế thủy ngân có thể an toàn khi sử dụng với bất kỳ loại dung dịch nào trong phòng thí nghiệm.", "answer": False, "reason": "Không an toàn với dung dịch có nhiệt độ vượt ngoài GHĐ của nhiệt kế hoặc các dung dịch có khả năng ăn mòn làm vỡ vỏ thủy tinh."}
        ]
    }
]

# -----------------------------------------------------------------------------
# PHẦN CHUNG - 3. CÂU TRẢ LỜI NGẮN (8 CÂU)
# -----------------------------------------------------------------------------
PHAN_CHUNG_SHORT = [
    {
        "id": 1,
        "content": "Quan sát hình vẽ sau. Chiều dài của chiếc bút chì bằng bao nhiêu cm?",
        "image": "khtn6_pencil_p4.png",
        "answer": "6,4 cm (hoặc 6,5 cm)",
        "solution": "Đầu gọt của bút chì được đặt trùng với vạch 0 cm. Đầu nhọn bút chì gióng thẳng đứng xuống thước chỉ vào vạch 6,4 cm (hoặc 6,5 cm). Do đó, chiều dài chiếc bút chì là 6,4 cm."
    },
    {
        "id": 2,
        "content": "Ngoài các đơn vị đo thông dụng hiện nay là mét, còn một số đơn vị đo chiều dài khác như inch (in), foot (ft), mile (dặm), hải lý,....Biết 1 dặm = 1,6093440 km. Vậy cơn bão đang ở cách bờ biển 40 dặm có nghĩa là cách bờ bao nhiêu kilomet? Làm tròn đến chữ số thập phân thứ nhất.",
        "answer": "64,4 km",
        "solution": "Khoảng cách từ tâm bão đến bờ biển tính theo kilomet là:\n40 × 1,6093440 km = 64,37376 km ≈ 64,4 km."
    },
    {
        "id": 3,
        "content": "Mẹ Lan dặn Lan ra chợ mua 5 lạng thịt nạc dăm. 5 lạng có nghĩa là bao nhiêu g?",
        "answer": "500 g",
        "solution": "Theo thói quen dân gian tại Việt Nam: 1 lạng = 100 g.\nVậy 5 lạng thịt = 5 × 100 g = 500 g."
    },
    {
        "id": 4,
        "content": "Cho các phát biểu sau: Có bao nhiêu phát biểu đúng?\na) Đơn vị của khối lượng là gam.\nb) Cân dùng để đo khối lượng của vật.\nc) Cân luôn luôn có hai đĩa.\nd) Một tạ bằng 100 kg.\ne) Một tấn bằng 100 tạ.\nf) Một tạ bông có khối lượng ít hơn 1 tạ sắt.",
        "answer": "2 phát biểu đúng",
        "solution": "Phân tích từng phát biểu:\n- a) Sai, đơn vị đo khối lượng hợp pháp trong hệ SI là kilôgam (kg).\n- b) Đúng, cân là dụng cụ dùng để đo khối lượng.\n- c) Sai, có nhiều loại cân chỉ có 1 đĩa (cân đồng hồ, cân điện tử) hoặc không có đĩa (cân móc treo, cân đòn).\n- d) Đúng, 1 tạ = 100 kg.\n- e) Sai, 1 tấn = 10 tạ = 1000 kg.\n- f) Sai, 1 tạ bông và 1 tạ sắt đều có khối lượng bằng nhau là 100 kg.\nVậy có 2 phát biểu đúng (b, d)."
    },
    {
        "id": 5,
        "content": "Thân nhiệt bình thường của người là bao nhiêu °C?",
        "answer": "37 °C",
        "solution": "Thân nhiệt bình thường của một người khỏe mạnh dao động trong khoảng từ 36,5 °C đến 37 °C (trung bình là 37 °C)."
    },
    {
        "id": 6,
        "content": "Inch (đọc là inh) là một trong những đơn vị đo chiều dài của Anh. Khi mua tivi, người ta hay nói tivi 17 inh có nghĩa đường chéo của màn hình là 17 inch. Biết 1 inch = 2,54cm. Nếu bố của Bình mua một chiếc tivi 50 inch, thì có nghĩa đường chéo của màn hình có chiều dài là bao nhiêu cm?",
        "answer": "127 cm",
        "solution": "Chiều dài đường chéo màn hình tivi 50 inch tính theo cm là:\n50 × 2,54 cm = 127 cm."
    },
    {
        "id": 7,
        "content": "Một cân Roberval có đòn cân phụ được vẽ như hình sau. ĐCNN của cân này là bao nhiêu g?",
        "image": "khtn6_roberval_p4.png",
        "answer": "0,2 g",
        "solution": "Quan sát trên đòn cân phụ: Khoảng cách giữa vạch 0 và vạch 1 g được chia thành 5 khoảng nhỏ bằng nhau.\nGiá trị của mỗi khoảng chia (ĐCNN) là: 1 g / 5 = 0,2 g."
    },
    {
        "id": 8,
        "content": "Mẹ Lan nhờ Lan đi chợ mua 500 g cam; 2 kg ổi; 6 lạng nhãn và 5 kg xoài. Số kilogram trái cây bạn Lan đã mua là bao nhiêu kg?",
        "answer": "8,1 kg",
        "solution": "Đổi các khối lượng về cùng đơn vị kilôgam (kg):\n- 500 g = 0,5 kg.\n- 2 kg ổi.\n- 6 lạng nhãn = 600 g = 0,6 kg.\n- 5 kg xoài.\nTổng số kilogram trái cây bạn Lan đã mua là:\n0,5 + 2 + 0,6 + 5 = 8,1 kg."
    },
    {
        "id": 9,
        "content": "Tại một nhà máy sản xuất bánh kẹo, An có thể đóng gói 1410 viên kẹo mỗi giờ, mỗi hộp chứa 30 viên kẹo, số hộp kẹo An đóng gói được trong 1 giờ là bao nhiêu hộp?",
        "answer": "47 hộp",
        "solution": "Số hộp kẹo An đóng gói được trong 1 giờ là:\n1410 / 30 = 47 (hộp)."
    }
]

# -----------------------------------------------------------------------------
# PHẦN CHUNG - 4. TỰ LUẬN (11 CÂU)
# -----------------------------------------------------------------------------
PHAN_CHUNG_ESSAY = [
    {
        "id": 1,
        "content": "Đổi các đơn vị sau:\na) 1 km = ....... m\nb) 1 m = ....... dm\nc) 1 cm = ....... m\nd) 1 m = ....... mm",
        "solution": "a) 1 km = 1 000 m\nb) 1 m = 10 dm\nc) 1 cm = 0,01 m (hay 1/100 m)\nd) 1 m = 1 000 mm"
    },
    {
        "id": 2,
        "content": "Đổi các đơn vị đo khối lượng sau:\na) 1,5 kg = ... g\nb) 3200 g = ... kg\nc) 2,4 tạ = ... kg\nd) 2,5 tấn = .... kg",
        "solution": "a) 1,5 kg = 1,5 × 1 000 = 1 500 g\nb) 3200 g = 3200 / 1 000 = 3,2 kg\nc) 2,4 tạ = 2,4 × 100 = 240 kg\nd) 2,5 tấn = 2,5 × 1 000 = 2 500 kg"
    },
    {
        "id": 3,
        "content": "Vào đầu mỗi năm học, trường THCS A thường tổ chức thăm khám sức khỏe cho các em học sinh khối 6. Sau khi đo chiều cao, My, Thảo, Bích và Hà được kết quả như sau:\n• My: 151 cm\n• Thảo: 1,52 m\n• Bích: 1,51 m\n• Hà: 148 cm\na) So sánh chiều cao của hai bạn My và Bích.\nb) Bạn nào cao nhất trong 4 bạn kể trên?",
        "solution": "Đổi các kết quả đo về cùng đơn vị mét (m):\n- My: 151 cm = 1,51 m\n- Thảo: 1,52 m\n- Bích: 1,51 m\n- Hà: 148 cm = 1,48 m\na) So sánh chiều cao: My và Bích có chiều cao bằng nhau (151 cm = 1,51 m).\nb) So sánh: 1,48 m < 1,51 m = 1,51 m < 1,52 m (Hà < My = Bích < Thảo).\nVậy bạn Thảo là người cao nhất trong 4 bạn (1,52 m hay 152 cm)."
    },
    {
        "id": 4,
        "content": "Dùng một sợi chỉ quấn đều 20 vòng lên thân một bút chì (mỗi vòng sát nhau và không chồng chéo lên nhau). Dùng thước thẳng đo chiều dài phần được quấn (trên thân cây bút chì) ta được trị số là 0,5cm. Đường kính của sợi chỉ là bao nhiêu?",
        "solution": "Gọi d là đường kính của sợi chỉ.\nVì 20 vòng chỉ quấn sát nhau tạo thành một đoạn dài 0,5 cm, nên đường kính của sợi chỉ là:\nd = 0,5 cm / 20 = 0,025 cm = 0,25 mm."
    },
    {
        "id": 5,
        "content": "Để thực hiện đo thời gian khi đi từ cổng trường vào lớp học em dùng loại đồng hồ nào? Giải thích?",
        "solution": "- Loại đồng hồ nên dùng: Đồng hồ bấm giây (hoặc đồng hồ điện tử có chức năng bấm giờ / điện thoại thông minh).\n- Giải thích: Khoảng thời gian đi từ cổng trường vào lớp học là một khoảng thời gian tương đối ngắn (thường từ 2 đến 5 phút). Đồng hồ bấm giây có độ chia nhỏ nhất rất nhỏ (0,1s hoặc 0,01s), thao tác bấm bắt đầu (Start) và dừng (Stop) chuẩn xác giúp ghi nhận thời gian di chuyển chính xác nhất."
    },
    {
        "id": 6,
        "content": "Một đĩa cân thăng bằng khi ở đĩa cân bên trái có 4 gói kẹo, đĩa cân bên phải có một quả cân 100 g, một quả cân 50 g, một quả cân 20 g và 3 quả cân 10 g.\na) Tính khối lượng của gói kẹo trên đĩa cân.\nb) Vậy 6 gói kẹo cùng loại như trên sẽ có khối lượng là bao nhiêu?",
        "solution": "a) Tổng khối lượng các quả cân trên đĩa cân bên phải là:\nm_phải = 100 + 50 + 20 + 3 × 10 = 200 g.\nVì cân ở trạng thái thăng bằng nên khối lượng của 4 gói kẹo ở đĩa bên trái bằng 200 g.\nKhối lượng của mỗi gói kẹo là:\nm_1gói = 200 / 4 = 50 g.\nb) Khối lượng của 6 gói kẹo cùng loại là:\nm_6gói = 6 × 50 = 300 g."
    },
    {
        "id": 7,
        "content": "Đổi các đơn vị đo thời gian sau thành giây:\na) 45 phút.\nb) 1 giờ 20 phút.\nc) 10 giờ.\nd) 1 ngày.\ne) 1 tuần.",
        "solution": "a) 45 phút = 45 × 60 = 2 700 s.\nb) 1 giờ 20 phút = (1 × 60 + 20) phút = 80 phút = 80 × 60 = 4 800 s.\nc) 10 giờ = 10 × 3 600 = 36 000 s.\nd) 1 ngày = 24 giờ = 24 × 3 600 = 86 400 s.\ne) 1 tuần = 7 ngày = 7 × 86 400 = 604 800 s."
    },
    {
        "id": 8,
        "content": "Một ô tô khởi hành ở Hà Nội lúc 8 giờ 15 phút. Sau khi chạy được 1 giờ thì đến Hải Dương. Ô tô nghỉ ở Hải Dương 20 phút, sau đó chạy 50 phút nữa thì đến Hải Phòng.\na) Xác định thời điểm mà ô tô đến Hải Dương và Hải Phòng.\nb) Xác định khoảng thời gian còn lại để đi đường, nếu ô tô muốn đến Quảng Ninh lúc 1 giờ chiều.",
        "solution": "a)\n- Thời điểm ô tô đến Hải Dương: 8 giờ 15 phút + 1 giờ = 9 giờ 15 phút.\n- Thời điểm ô tô rời Hải Dương: 9 giờ 15 phút + 20 phút = 9 giờ 35 phút.\n- Thời điểm ô tô đến Hải Phòng: 9 giờ 35 phút + 50 phút = 9 giờ 85 phút = 10 giờ 25 phút.\nb)\n- Đổi: 1 giờ chiều = 13 giờ.\n- Khoảng thời gian còn lại để ô tô đi tiếp từ Hải Phòng đến Quảng Ninh lúc 13 giờ là:\n13 giờ - 10 giờ 25 phút = 12 giờ 60 phút - 10 giờ 25 phút = 2 giờ 35 phút."
    },
    {
        "id": 9,
        "content": "Hai bạn Huệ và Lan cùng hẹn nhau tại thư viện để cùng làm bài tập. Huệ bắt đầu đi vào lúc 7 giờ và đến nơi lúc 7 giờ 25 phút, Lan bắt đầu đi sau Huệ 5 phút và đến nơi lúc 7 giờ 35 phút. Hỏi thời gian di chuyển từ nhà đến thư viện của ai lâu hơn và lâu hơn bao nhiêu phút?",
        "solution": "- Thời gian di chuyển của Huệ: 7 giờ 25 phút - 7 giờ = 25 phút.\n- Thời điểm Lan xuất phát: 7 giờ + 5 phút = 7 giờ 05 phút.\n- Thời gian di chuyển của Lan: 7 giờ 35 phút - 7 giờ 05 phút = 30 phút.\n- So sánh: Vì 30 phút > 25 phút nên thời gian di chuyển từ nhà đến thư viện của bạn Lan lâu hơn của bạn Huệ.\n- Lan đi lâu hơn số phút là: 30 - 25 = 5 phút."
    },
    {
        "id": 10,
        "content": "Trường của bạn Bình có tổ chức chuyến du lịch ngoại khóa tại Củ Chi cách trường khoảng 100 km. Xe khởi hành tại trường lúc 6 giờ sáng. Trên đường đi xe có dừng lại tại điểm A khoảng 30 phút để các bạn ăn sáng. Xe rời khỏi điểm A lúc 7 giờ 30 phút và đến Củ Chi lúc 8 giờ 40 phút. Hỏi:\na) Quãng đường từ trường đến khu du lịch là bao nhiêu mét?\nb) Xe đến điểm A lúc mấy giờ?\nc) Khoảng thời gian xe chạy từ điểm A đến Củ Chi mất bao nhiêu phút?",
        "solution": "a) Đổi quãng đường từ kilômét sang mét:\n100 km = 100 × 1 000 = 100 000 m.\nb) Xe rời điểm A lúc 7 giờ 30 phút sau khi dừng nghỉ 30 phút, vậy thời điểm xe đến điểm A là:\n7 giờ 30 phút - 30 phút = 7 giờ 00 phút (7 giờ sáng).\nc) Khoảng thời gian xe chạy từ điểm A đến Củ Chi là:\n8 giờ 40 phút - 7 giờ 30 phút = 1 giờ 10 phút = 70 phút."
    },
    {
        "id": 11,
        "content": "Vì sao bảng chia độ của nhiệt kế y tế lại không có nhiệt độ dưới 34°C và trên 42°C?",
        "solution": "Giải thích:\n- Nhiệt kế y tế là dụng cụ chuyên biệt dùng để đo nhiệt độ cơ thể người.\n- Thân nhiệt của người bình thường khoảng 36,5°C - 37°C. Khi cơ thể bị hạ thân nhiệt nghiêm trọng hoặc sốt rất cao thì nhiệt độ cũng chỉ dao động trong khoảng từ 34°C đến 42°C (nếu nhiệt độ cơ thể vượt ra ngoài khoảng này, con người sẽ rơi vào tình trạng nguy kịch đến tính mạng).\n- Do đó, việc giới hạn thang đo từ 34°C đến 42°C vừa đáp ứng đầy đủ yêu cầu theo dõi y tế, vừa giúp bảng chia độ có khoảng cách rộng, chia nhỏ được từng 0,1°C để đọc kết quả chính xác và rõ ràng nhất."
    }
]
