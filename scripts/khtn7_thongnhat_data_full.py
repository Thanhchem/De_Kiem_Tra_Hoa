# -*- coding: utf-8 -*-
r"""
Dữ liệu Đề Cương Ôn Tập Giữa Học Kì I – KHTN 7
Trường THCS Thống Nhất – Năm học 2026 – 2027
Bao gồm 3 phân môn:
- Phần I: Vật lí (Chủ đề Năng lượng và sự biến đổi)
- Phần II: Hóa học (Chủ đề Chất và sự biến đổi chất)
- Phần III: Sinh học (Chủ đề Vật sống)
"""

# ==============================================================================
# PHẦN 1: VẬT LÍ (CHỦ ĐỀ: NĂNG LƯỢNG VÀ SỰ BIẾN ĐỔI)
# ==============================================================================

VATLI_THEORY = [
    ("1. Cổng quang điện và đồng hồ đo thời gian hiện số", [
        "Cổng quang điện có vai trò như công tắc điều khiển mở, đóng đồng hồ đo thời gian hiện số.",
        "Gồm một đầu phát tia hồng ngoại và một đầu thu tia hồng ngoại. Khi có vật đi qua chắn tia hồng ngoại, tín hiệu điện được gửi qua dây cáp đến đồng hồ.",
        "Đồng hồ đo thời gian hiện số hoạt động như đồng hồ bấm giây nhưng được điều khiển tự động bằng cổng quang điện."
    ]),
    ("2. Tốc độ chuyển động", [
        "Khái niệm: Thương số s/t đặc trưng cho mức độ nhanh hay chậm của chuyển động gọi là tốc độ.",
        "Công thức tính tốc độ: *v = s / t* (s: quãng đường đi được; t: thời gian đi hết quãng đường; v: tốc độ).",
        "Đơn vị đo: mét trên giây (m/s) hoặc kilômét trên giờ (km/h). Quy tắc đổi: *1 m/s = 3,6 km/h*."
    ]),
    ("3. Đo tốc độ & Thiết bị bắn tốc độ", [
        "Đo tốc độ bằng đồng hồ bấm giây hoặc cổng quang điện kết hợp đồng hồ hiện số.",
        "Thiết bị bắn tốc độ dùng để đo trực tiếp tốc độ chuyển động của các phương tiện giao thông, giúp kiểm soát an toàn giao thông đường bộ."
    ])
]

VATLI_MCQ = [
    {
        "id": 1,
        "content": "Trong bộ thiết bị thí nghiệm, cổng quang điện đóng vai trò chính là gì?",
        "options": [
            "A. Dụng cụ đo trực tiếp khối lượng của vật.",
            "B. Công tắc điều khiển mở hoặc đóng đồng hồ đo thời gian hiện số.",
            "C. Nguồn điện cung cấp năng lượng cho toàn bộ hệ thống thí nghiệm.",
            "D. Dụng cụ dùng để đo trực tiếp chiều dài và kích thước vật rắn."
        ],
        "answer": "B",
        "solution": "Cổng quang điện đóng vai trò như một công tắc tự động điều khiển việc bắt đầu (mở) và kết thúc (đóng) đếm thời gian của đồng hồ hiện số."
    },
    {
        "id": 2,
        "content": "Cổng quang điện bao gồm những bộ phận chính nào?",
        "options": [
            "A. Bộ phận phát tia hồng ngoại D1, bộ phận thu tia hồng ngoại D2 và dây cáp nối.",
            "B. Bộ phận phát sáng, màn hình tinh thể lỏng và nút điều chỉnh thang đo.",
            "C. Hộp chứa nguồn điện pin mặt trời, dây dẫn và công tắc nguồn điện xoay chiều.",
            "D. Cảm ứng từ, công tắc ngắt mạch tự động và các chốt cắm tín hiệu số."
        ],
        "answer": "A",
        "solution": "Cổng quang điện gồm đầu phát tia hồng ngoại D1, đầu thu tia hồng ngoại D2 và dây dẫn cáp nối truyền tín hiệu tới đồng hồ."
    },
    {
        "id": 3,
        "content": "Dây cáp nối giữa cổng quang điện và đồng hồ đo thời gian hiện số có những tác dụng nào sau đây?",
        "options": [
            "A. Chỉ dùng để truyền tín hiệu ánh sáng từ đèn LED đến mắt người quan sát.",
            "B. Chỉ dùng để cung cấp dòng điện một chiều cho cổng quang hoạt động.",
            "C. Cung cấp điện cho cổng quang và gửi tín hiệu điện từ cổng quang tới đồng hồ.",
            "D. Giữ cố định cổng quang điện trên giá đỡ quang học trong suốt quá trình đo."
        ],
        "answer": "C",
        "solution": "Dây cáp vừa cấp nguồn điện từ đồng hồ cho cổng quang hoạt động, vừa truyền tín hiệu điện điều khiển đóng/ngắt từ cổng quang về đồng hồ."
    },
    {
        "id": 4,
        "content": "Trên mặt trước của đồng hồ đo thời gian hiện số, nút MODE dùng để làm gì?",
        "options": [
            "A. Đưa số hiển thị trên mặt đồng hồ trở về trạng thái 0.000.",
            "B. Chọn chế độ làm việc cho đồng hồ (ví dụ: đo thời gian vật qua 1 cổng quang, 2 cổng quang...).",
            "C. Lựa chọn giới hạn đo và độ chia nhỏ nhất phù hợp với từng đại lượng.",
            "D. Bật hoặc tắt nguồn điện chính cung cấp cho toàn bộ mạch điện tử bên trong."
        ],
        "answer": "B",
        "solution": "Nút MODE dùng để chuyển đổi và cài đặt chế độ làm việc của đồng hồ đo thời gian hiện số."
    },
    {
        "id": 5,
        "content": "Nút RESET trên mặt trước của đồng hồ đo thời gian hiện số có tác dụng gì?",
        "options": [
            "A. Thay đổi chế độ đo lường từ thời gian sang vận tốc tức thời.",
            "B. Khởi động lại hệ thống nguồn điện chính của đồng hồ.",
            "C. Cho đồng hồ quay về trạng thái ban đầu, mặt đồng hồ hiện chỉ số 0.000.",
            "D. Lưu lại kết quả đo cuối cùng vào bộ nhớ của vi mạch điện tử."
        ],
        "answer": "C",
        "solution": "Nút RESET dùng để xóa kết quả đo cũ, đưa chỉ số hiển thị trở về trạng thái số 0.000 ban đầu để sẵn sàng cho lần đo tiếp theo."
    },
    {
        "id": 6,
        "content": "Những bộ phận nào sau đây nằm ở mặt sau của đồng hồ đo thời gian hiện số?",
        "options": [
            "A. Nút thang đo, nút mode và màn hình hiển thị số.",
            "B. Công tắc điện, ba ổ cắm cổng quang A, B, C và ổ cắm điện.",
            "C. Chỉ có duy nhất ổ cắm cổng quang A và công tắc nguồn phụ.",
            "D. Đèn báo tín hiệu hồng ngoại và các nút hiệu chỉnh cơ học."
        ],
        "answer": "B",
        "solution": "Mặt sau của đồng hồ hiện số bố trí công tắc nguồn điện, các ổ cắm tín hiệu vào A, B, C và ổ cắm nguồn điện lưới."
    },
    {
        "id": 7,
        "content": "Tại sao khi sử dụng cổng quang điện kết hợp đồng hồ đo thời gian hiện số, kết quả đo thời gian lại khách quan và chính xác hơn so với dùng đồng hồ bấm giây thủ công?",
        "options": [
            "A. Vì đồng hồ điện tử tự động ghi nhận tín hiệu khi vật chắn sáng, loại bỏ được độ trễ do phản xạ và thao tác bấm nút bằng tay của con người.",
            "B. Vì đồng hồ đo thời gian hiện số sử dụng nguồn điện lưới quốc gia có điện áp lớn hơn.",
            "C. Vì cổng quang điện có thể phát ra tia laser cường độ mạnh để đẩy vật chuyển động nhanh hơn.",
            "D. Vì thiết bị tự động tính toán ra kết quả vận tốc mà không cần học sinh thực hiện phép tính chia."
        ],
        "answer": "A",
        "solution": "Hệ thống điện tử tự động nhận tín hiệu cắt tia hồng ngoại, triệt tiêu sai số thời gian phản xạ chủ quan của người bấm đồng hồ."
    },
    {
        "id": 8,
        "content": "Khi ước lượng thời gian chuyển động của vật lớn hơn 10s, cần lựa chọn thang đo nào của đồng hồ hiện số?",
        "options": [
            "A. 9,999s – 0,0001s",
            "B. 99s – 1s",
            "C. 10s – 9s",
            "D. 99,99s – 0,01s"
        ],
        "answer": "D",
        "solution": "Thang đo 99,99s có giới hạn đo 99,99 giây và độ chia nhỏ nhất 0,01 giây, rất phù hợp đo các khoảng thời gian từ 10s đến 99s với độ chính xác cao."
    },
    {
        "id": 9,
        "content": "Để đo thời gian chạy của vận động viên chạy 100 m, dụng cụ nào sau đây phù hợp nhất?",
        "options": [
            "A. Đồng hồ đo thời gian dùng cổng quang điện.",
            "B. Đồng hồ bấm giây.",
            "C. Đồng hồ cát.",
            "D. Đồng hồ treo tường."
        ],
        "answer": "B",
        "solution": "Đồng hồ bấm giây là dụng cụ thông dụng, gọn nhẹ và phù hợp nhất để bấm giờ các sự kiện thể thao chạy cự li."
    },
    {
        "id": 10,
        "content": "Đại lượng nào sau đây cho biết mức độ nhanh hay chậm của chuyển động?",
        "options": [
            "A. Quãng đường.",
            "B. Khối lượng.",
            "C. Tốc độ.",
            "D. Thời gian chuyển động."
        ],
        "answer": "C",
        "solution": "Tốc độ là đại lượng đặc trưng cho mức độ nhanh chậm của chuyển động."
    },
    {
        "id": 11,
        "content": "Công thức nào sau đây dùng để tính tốc độ chuyển động?",
        "options": [
            "A. v = s/t.",
            "B. v = s.t.",
            "C. s = v/t.",
            "D. t = v.t."
        ],
        "answer": "A",
        "solution": "Công thức tính tốc độ: *v = s / t*."
    },
    {
        "id": 12,
        "content": "Trong các đơn vị sau, đơn vị nào là đơn vị đo tốc độ?",
        "options": [
            "A. km/h.",
            "B. m.s.",
            "C. km.h.",
            "D. s/m."
        ],
        "answer": "A",
        "solution": "km/h (kilômét trên giờ) là đơn vị đo tốc độ thông dụng."
    }
]

VATLI_TF = [
    {
        "id": 1,
        "prompt": "Câu 1. Ý nghĩa của tốc độ:",
        "items": [
            {"id": "a", "text": "Tốc độ cho biết mức độ nhanh hay chậm của chuyển động.", "ans": True, "reason": "Đúng theo định nghĩa tốc độ."},
            {"id": "b", "text": "Dựa vào tốc độ, ta có thể so sánh độ nhanh, chậm của các chuyển động.", "ans": True, "reason": "Chuyển động có tốc độ lớn hơn thì nhanh hơn."},
            {"id": "c", "text": "Vật có tốc độ càng lớn thì chuyển động càng chậm.", "ans": False, "reason": "Tốc độ càng lớn thì vật chuyển động càng nhanh."},
            {"id": "d", "text": "Tốc độ không có ý nghĩa trong việc so sánh chuyển động của các vật.", "ans": False, "reason": "Tốc độ là tiêu chuẩn cơ bản để so sánh chuyển động."}
        ]
    },
    {
        "id": 2,
        "prompt": "Câu 2. Khái niệm tốc độ:",
        "items": [
            {"id": "a", "text": "Tốc độ là đại lượng cho biết mức độ nhanh hay chậm của chuyển động.", "ans": True, "reason": "Đúng định nghĩa."},
            {"id": "b", "text": "Tốc độ được xác định bằng quãng đường đi được trong một đơn vị thời gian.", "ans": True, "reason": "Đúng quy chuẩn."},
            {"id": "c", "text": "Tốc độ được tính bằng thời gian chuyển động chia cho quãng đường đi được.", "ans": False, "reason": "Tốc độ bằng quãng đường chia cho thời gian (v = s/t)."},
            {"id": "d", "text": "Đơn vị đo tốc độ thường dùng là m/s và km/h.", "ans": True, "reason": "Hai đơn vị thông dụng nhất."}
        ]
    },
    {
        "id": 3,
        "prompt": "Câu 3. Đơn vị đo tốc độ:",
        "items": [
            {"id": "a", "text": "Mét trên giây (m/s) là đơn vị đo tốc độ.", "ans": True, "reason": "Đúng đơn vị SI."},
            {"id": "b", "text": "Kilômét trên giờ (km/h) có thể dùng để biểu diễn tốc độ của xe máy.", "ans": True, "reason": "Đúng thực tế đời sống."},
            {"id": "c", "text": "Đơn vị đo tốc độ chỉ có km/h và m/s, không có đơn vị nào khác.", "ans": False, "reason": "Còn có cm/s, mm/s, m/min, knot (hải lý/h)..."},
            {"id": "d", "text": "Kilômét (km) là đơn vị đo tốc độ.", "ans": False, "reason": "Kilômét là đơn vị đo độ dài, quãng đường."}
        ]
    },
    {
        "id": 4,
        "prompt": "Câu 4. Phân biệt các đơn vị đo:",
        "items": [
            {"id": "a", "text": "Mét trên giây (m/s) là đơn vị đo tốc độ.", "ans": True, "reason": "Đúng."},
            {"id": "b", "text": "Kilômét trên giờ (km/h) là đơn vị đo quãng đường.", "ans": False, "reason": "km/h là đơn vị đo tốc độ."},
            {"id": "c", "text": "Tốc độ của ô tô thường được biểu diễn bằng đơn vị km/h.", "ans": True, "reason": "Đúng thực tế."},
            {"id": "d", "text": "Đơn vị m/s cho biết số mét vật đi được trong một giây khi chuyển động với tốc độ không đổi.", "ans": True, "reason": "Đúng ý nghĩa vật lí của đơn vị m/s."}
        ]
    },
    {
        "id": 5,
        "prompt": "Câu 5. Đo tốc độ bằng đồng hồ bấm giây:",
        "items": [
            {"id": "a", "text": "Chỉ cần sử dụng đồng hồ bấm giây là có thể xác định được tốc độ của vật.", "ans": False, "reason": "Cần thêm dụng cụ đo quãng đường (thước đo)."},
            {"id": "b", "text": "Bấm đồng hồ khi vật bắt đầu chuyển động và dừng đồng hồ khi vật kết thúc chuyển động trên quãng đường đã chọn.", "ans": True, "reason": "Đúng quy trình bấm giây."},
            {"id": "c", "text": "Muốn tính tốc độ, lấy thời gian chuyển động chia cho quãng đường đi được.", "ans": False, "reason": "Lấy quãng đường chia thời gian."},
            {"id": "d", "text": "Tốc độ của vật được tính bằng quãng đường đi được chia cho thời gian chuyển động.", "ans": True, "reason": "Đúng công thức v = s / t."}
        ]
    },
    {
        "id": 6,
        "prompt": "Câu 6. Đo tốc độ bằng đồng hồ bấm giây và cổng quang điện:",
        "items": [
            {"id": "a", "text": "Khi dùng đồng hồ bấm giây, cần đo quãng đường và thời gian vật chuyển động.", "ans": True, "reason": "Đúng quy trình."},
            {"id": "b", "text": "Khi dùng cổng quang điện, thời gian được ghi nhận tự động theo cách bố trí thí nghiệm.", "ans": True, "reason": "Đúng nguyên lí tự động."},
            {"id": "c", "text": "Đo tốc độ bằng cổng quang điện không cần xác định quãng đường hoặc khoảng cách đo cần thiết.", "ans": False, "reason": "Vẫn phải đo khoảng cách giữa 2 cổng hoặc chiều rộng tấm chắn sáng."},
            {"id": "d", "text": "Muốn tính tốc độ, lấy quãng đường đi được chia cho thời gian chuyển động tương ứng.", "ans": True, "reason": "Đúng công thức."}
        ]
    },
    {
        "id": 7,
        "prompt": "Câu 7. Thiết bị đo tốc độ trong thí nghiệm:",
        "items": [
            {"id": "a", "text": "Đồng hồ bấm giây dùng để đo thời gian vật chuyển động.", "ans": True, "reason": "Đúng."},
            {"id": "b", "text": "Cổng quang điện có thể tự động ghi nhận thời gian khi vật đi qua cổng.", "ans": True, "reason": "Đúng chức năng cổng quang."},
            {"id": "c", "text": "Khi sử dụng cổng quang điện, luôn cần dùng tay bấm đồng hồ để đo thời gian.", "ans": False, "reason": "Đồng hồ hiện số tự ngắt/bật nhờ tín hiệu điện từ cổng quang."},
            {"id": "d", "text": "Muốn xác định tốc độ, cần biết quãng đường và thời gian tương ứng.", "ans": True, "reason": "Đúng."}
        ]
    },
    {
        "id": 8,
        "prompt": "Câu 8. Thiết bị bắn tốc độ:",
        "items": [
            {"id": "a", "text": "Thiết bị bắn tốc độ giúp xác định tốc độ của phương tiện giao thông.", "ans": True, "reason": "Đúng ứng dụng."},
            {"id": "b", "text": "Thiết bị bắn tốc độ chỉ dùng để đo quãng đường phương tiện đi được.", "ans": False, "reason": "Thiết bị đo trực tiếp tốc độ tức thời."},
            {"id": "c", "text": "Cảnh sát giao thông có thể sử dụng thiết bị bắn tốc độ để phát hiện phương tiện chạy quá tốc độ quy định.", "ans": True, "reason": "Đúng thực tế giám sát giao thông."},
            {"id": "d", "text": "Thiết bị bắn tốc độ góp phần nâng cao ý thức chấp hành quy định về tốc độ khi tham gia giao thông.", "ans": True, "reason": "Đúng ý nghĩa xã hội."}
        ]
    },
    {
        "id": 9,
        "prompt": "Câu 9. Một cơn bão nhiệt đới di chuyển từ ngoài khơi hướng vào đất liền, vượt qua khoảng cách 300 km trong thời gian 12 giờ. Xác định tính đúng sai của các nhận định:",
        "items": [
            {"id": "a", "text": "Tốc độ di chuyển của bão cho biết mức độ nhanh hay chậm của cơn bão.", "ans": True, "reason": "Tốc độ v = 300 / 12 = 25 km/h."},
            {"id": "b", "text": "Trong thực tế dự báo thời tiết, tốc độ bão thường được tính bằng km/h.", "ans": True, "reason": "Đúng bản tin dự báo khí tượng."},
            {"id": "c", "text": "Tốc độ di chuyển của bão càng lớn kết hợp với sức gió mạnh sẽ làm giảm mức độ tàn phá, diện tích ảnh hưởng và nguy cơ gây sạt lở, ngập lụt ở vùng đất liền.", "ans": False, "reason": "Làm tăng mức độ nguy hiểm và phạm vi tàn phá."},
            {"id": "d", "text": "Để đo tốc độ di chuyển của bão trong thực tế, người ta sử dụng đồng hồ bấm giây và thước kẻ đo trực tiếp trên bản đồ giấy khổ nhỏ.", "ans": False, "reason": "Phải dùng ảnh vệ tinh, radar khí tượng và máy tính phân tích."}
        ]
    },
    {
        "id": 10,
        "prompt": "Câu 10. Một vùng áp thấp nhiệt đới di chuyển trên biển theo đường thẳng với quãng đường 180 km trong thời gian 6 giờ. Xác định tính đúng sai của các nhận định:",
        "items": [
            {"id": "a", "text": "Tốc độ dịch chuyển của áp thấp nhiệt đới cho biết chiều dài quãng đường đi được trong một đơn vị thời gian.", "ans": True, "reason": "v = 180 / 6 = 30 km/h."},
            {"id": "b", "text": "Đơn vị đo tốc độ dịch chuyển áp thấp có thể sử dụng m/s hoặc km/h.", "ans": True, "reason": "Đúng các đơn vị đo tốc độ."},
            {"id": "c", "text": "Áp thấp di chuyển càng chậm thì lượng mưa trút xuống khu vực đó càng kéo dài, dễ gây ra tình trạng ngập lụt diện rộng và sạt lở đất nghiêm trọng.", "ans": True, "reason": "Đúng đặc điểm thời tiết ngập lụt do áp thấp di chuyển chậm."},
            {"id": "d", "text": "Trong phòng thí nghiệm, để học sinh hình dung được cách đo tốc độ chuyển động, giáo viên có thể cho học sinh dùng xe đồ chơi chạy trên mặt bàn kết hợp thước và đồng hồ bấm giây.", "ans": True, "reason": "Mô hình thí nghiệm trực quan rất tốt."}
        ]
    },
    {
        "id": 11,
        "prompt": "Câu 11. Tâm một cơn bão mạnh di chuyển trên vùng biển rộng, đi được quãng đường 120 km trong thời gian 4 giờ. Xác định tính đúng sai của các nhận định:",
        "items": [
            {"id": "a", "text": "Tốc độ chuyển động của tâm bão cho biết trong một đơn vị thời gian tâm bão dịch chuyển được bao nhiêu kilômét.", "ans": True, "reason": "Tốc độ v = 120 / 4 = 30 km/h."},
            {"id": "b", "text": "Đơn vị đo tốc độ trong hệ SI là km/h và cm/s.", "ans": False, "reason": "Đơn vị đo tốc độ chuẩn trong hệ SI duy nhất là mét trên giây (m/s)."},
            {"id": "c", "text": "Tốc độ di chuyển chậm của một cơn bão khi tiến vào vùng ven biển giúp làm giảm đáng kể sức tàn phá.", "ans": False, "reason": "Bão di chuyển chậm quần thảo gây mưa to gió lớn kéo dài, sức tàn phá dữ dội hơn."},
            {"id": "d", "text": "Đo tốc độ của bão trên diện rộng không thể dùng dụng cụ đo thủ công thông thường mà phải dựa vào công nghệ quan trắc khí tượng hiện đại.", "ans": True, "reason": "Cần ảnh mây vệ tinh và radar thời tiết."}
        ]
    }
]

VATLI_SHORT = [
    {
        "id": 1,
        "content": "Một đoàn tàu hỏa đi từ ga A đến ga B cách nhau 30 km trong 30 phút. Tốc độ của đoàn tàu là bao nhiêu km/h?",
        "answer": "60 km/h",
        "solution": "Đổi t = 30 phút = 0,5 h. Tốc độ đoàn tàu: *v = s / t = 30 / 0,5 = 60 km/h*."
    },
    {
        "id": 2,
        "content": "Một người đi bộ với tốc độ 4 km/h. Biết thời gian người đó đi từ nhà đến nơi làm việc là 30 phút. Quãng đường từ nhà đến nơi làm việc bao nhiêu km?",
        "answer": "2 km",
        "solution": "Đổi t = 30 phút = 0,5 h. Quãng đường: *s = v × t = 4 × 0,5 = 2 km*."
    },
    {
        "id": 3,
        "content": "Một người đi xe đạp với tốc độ 18 km/h trong 20 phút. Quãng đường người đó đi được bao nhiêu km?",
        "answer": "6 km",
        "solution": "Đổi t = 20 phút = 20/60 h = 1/3 h. Quãng đường: *s = v × t = 18 × (1/3) = 6 km*."
    },
    {
        "id": 4,
        "content": "Một máy bay bay với tốc độ 800 km/h từ Hà Nội đến Thành phố Hồ Chí Minh. Biết quãng đường bay dài 1400 km. Hỏi máy bay bay hết bao nhiêu phút?",
        "answer": "105 phút",
        "solution": "Thời gian bay tính theo giờ: *t = s / v = 1400 / 800 = 1,75 h*. Đổi ra phút: *t = 1,75 × 60 = 105 phút* (hoặc 1 giờ 45 phút)."
    },
    {
        "id": 5,
        "content": "Một ô tô chuyển động trên đoạn đường đầu với tốc độ 15 m/s trong 30 phút, sau đó tiếp tục chuyển động trên đoạn đường kế tiếp với tốc độ 60 km/h trong 45 phút. Tổng quãng đường ô tô đi được trong 75 phút là bao nhiêu km?",
        "answer": "72 km",
        "solution": "• Đoạn 1: v_1 = 15 m/s = 15 × 3,6 = 54 km/h; t_1 = 30 phút = 0,5 h => s_1 = 54 × 0,5 = 27 km.\n• Đoạn 2: v_2 = 60 km/h; t_2 = 45 phút = 0,75 h => s_2 = 60 × 0,75 = 45 km.\n• Tổng quãng đường: *s = s_1 + s_2 = 27 + 45 = 72 km*."
    },
    {
        "id": 6,
        "content": "Một ô tô chuyển động trên đoạn đường đầu với tốc độ 12 m/s trong 20 phút, sau đó tiếp tục chuyển động trên đoạn đường kế tiếp với tốc độ 54 km/h trong 40 phút. Tốc độ trung bình của ô tô trong suốt thời gian 60 phút là bao nhiêu m/s?",
        "answer": "14 m/s",
        "solution": "• Đoạn 1: t_1 = 20 phút = 1200 s => s_1 = 12 × 1200 = 14.400 m.\n• Đoạn 2: v_2 = 54 km/h = 15 m/s; t_2 = 40 phút = 2400 s => s_2 = 15 × 2400 = 36.000 m.\n• Tổng s = 14.400 + 36.000 = 50.400 m; tổng t = 60 phút = 3600 s.\n• Tốc độ trung bình: *v_tb = s / t = 50.400 / 3600 = 14 m/s* (tương ứng 50,4 km/h)."
    },
    {
        "id": 7,
        "content": "Một người đi bộ trong 10 phút với tốc độ 1,5 m/s, sau đó chạy trong 5 phút với tốc độ 3 m/s. Tốc độ trung bình của người đó trong suốt thời gian chuyển động là bao nhiêu km/h?",
        "answer": "7,2 km/h",
        "solution": "• s_1 = 1,5 × (10 × 60) = 900 m.\n• s_2 = 3 × (5 × 60) = 900 m.\n• Tổng s = 1800 m = 1,8 km; tổng t = 15 phút = 0,25 h.\n• Tốc độ trung bình: *v_tb = 1,8 / 0,25 = 7,2 km/h* (hoặc 2 m/s)."
    }
]

# ==============================================================================
# PHẦN 2: HÓA HỌC (CHỦ ĐỀ: CHẤT VÀ SỰ BIẾN ĐỔI CHẤT)
# ==============================================================================

HOAHOC_THEORY = [
    ("1. Cấu tạo nguyên tử & Khối lượng nguyên tử", [
        "Nguyên tử là hạt vô cùng nhỏ bé, tạo nên các chất và trung hòa về điện.",
        "Cấu tạo: Gồm hạt nhân ở tâm (chứa proton mang điện dương và neutron không mang điện) và lớp vỏ gồm các electron mang điện âm chuyển động xung quanh.",
        "Khối lượng nguyên tử: Tập trung hầu hết ở hạt nhân (vì khối lượng electron rất nhỏ, không đáng kể). Đơn vị khối lượng nguyên tử: amu (1 amu ≈ 1,6605 × 10^-24 g)."
    ]),
    ("2. Nguyên tố hóa học", [
        "Nguyên tố hóa học là tập hợp các nguyên tử cùng loại có cùng số proton trong hạt nhân.",
        "Kí hiệu hóa học gồm 1 hoặc 2 chữ cái, trong đó chữ cái đầu viết hoa, chữ cái thứ hai viết thường."
    ])
]

HOAHOC_MCQ = [
    {
        "id": 1,
        "content": "Bước đầu tiên trong phương pháp tìm hiểu tự nhiên là?",
        "options": [
            "A. Lập kế hoạch kiểm tra dự đoán.",
            "B. Đề xuất vấn đề cần tìm hiểu.",
            "C. Viết báo cáo.",
            "D. Đưa ra dự đoán để giải quyết vấn đề."
        ],
        "answer": "B",
        "solution": "Bước đầu tiên của tiến trình tìm hiểu tự nhiên luôn là Đề xuất vấn đề (quan sát và đặt câu hỏi nghiên cứu)."
    },
    {
        "id": 2,
        "content": "Con người có thể định lượng được các sự vật và hiện tượng tự nhiên dựa trên kĩ năng nào?",
        "options": [
            "A. Kĩ năng quan sát, phân loại.",
            "B. Kĩ năng liên kết tri thức.",
            "C. Kĩ năng dự báo.",
            "D. Kĩ năng đo."
        ],
        "answer": "D",
        "solution": "Kĩ năng đo cho phép xác định chính xác các giá trị số lượng, kích thước, khối lượng, nhiệt độ, thời gian (định lượng)."
    },
    {
        "id": 3,
        "content": "Một học sinh sau khi thực hiện thí nghiệm kết tinh muối từ dung dịch nước muối bão hòa, đã rút ra kết luận: \"Khi nước bay hơi (do đun nóng), chất tan là muối không bay hơi mà kết tinh lại thành các hạt rắn\". Học sinh này đã vận dụng kĩ năng tiến trình nào chủ yếu để đưa ra kết luận trên?",
        "options": [
            "A. Kĩ năng quan sát và phân loại.",
            "B. Kĩ năng liên kết.",
            "C. Kĩ năng đo.",
            "D. Kĩ năng dự báo."
        ],
        "answer": "A",
        "solution": "Học sinh sử dụng kĩ năng quan sát trực tiếp hiện tượng tinh thể muối đọng lại khi nước bay hơi để rút ra bản chất."
    },
    {
        "id": 4,
        "content": "Khẳng định nào sau đây là KHÔNG đúng?",
        "options": [
            "A. Dự báo là kĩ năng cần thiết trong nghiên cứu khoa học tự nhiên.",
            "B. Dự báo là kĩ năng không cần thiết của người làm nghiên cứu.",
            "C. Dự báo là dự đoán điều gì sẽ xảy ra dựa vào quan sát, kiến thức, suy luận của con người,... về sự vật, hiện tượng.",
            "D. Kĩ năng dự báo thường được sử dụng trong bước dự đoán của phương pháp tìm hiểu tự nhiên."
        ],
        "answer": "B",
        "solution": "Khẳng định B sai vì dự báo là một trong những kĩ năng tiến trình quan trọng hàng đầu trong nghiên cứu khoa học."
    },
    {
        "id": 5,
        "content": "Đơn vị quốc tế dùng để đo khối lượng nguyên tử là:",
        "options": [
            "A. Gam (g).",
            "B. Kilôgam (kg).",
            "C. Đơn vị khối lượng nguyên tử (amu).",
            "D. Miligram (mg)."
        ],
        "answer": "C",
        "solution": "Đơn vị khối lượng nguyên tử quốc tế quy ước là amu (atomic mass unit)."
    },
    {
        "id": 6,
        "content": "Phát biểu nào sau đây không mô tả đúng vỏ nguyên tử theo mô hình nguyên tử của Rutherford - Bohr?",
        "options": [
            "A. Electron chuyển động xung quanh hạt nhân theo từng lớp khác nhau tạo thành các lớp electron.",
            "B. Lớp electron trong cùng gần hạt nhân nhất có tối đa 2 electron, các lớp electron khác có chứa tối đa 8 electron hoặc nhiều hơn.",
            "C. Lớp electron trong cùng gần hạt nhân nhất có tối đa 8 electron, các lớp electron khác có chứa tối đa nhiều hơn 8 electron.",
            "D. Các electron sắp xếp vào từng lớp theo thứ tự từ trong ra ngoài cho đến hết."
        ],
        "answer": "C",
        "solution": "Phát biểu C sai vì lớp trong cùng (lớp 1) chỉ chứa tối đa 2 electron, không phải 8 electron."
    },
    {
        "id": 7,
        "content": "Hạt nhân của hầu hết các nguyên tử được cấu tạo từ các loại hạt nào?",
        "options": [
            "A. Electron và Proton.",
            "B. Proton và Neutron.",
            "C. Electron và Neutron.",
            "D. Chỉ chứa Proton."
        ],
        "answer": "B",
        "solution": "Hạt nhân gồm hạt proton mang điện tích dương và hạt neutron không mang điện."
    },
    {
        "id": 8,
        "content": "Theo quy ước quốc tế, cách viết kí hiệu hóa học của nguyên tố nào sau đây là đúng?",
        "options": [
            "A. Gồm 2 chữ cái viết hoa hoàn toàn.",
            "B. Gồm 2 chữ cái viết thường hoàn toàn.",
            "C. Gồm 1 hoặc 2 chữ cái, trong đó chữ cái đầu viết hoa, chữ cái sau viết thường.",
            "D. Gồm 1 chữ cái viết hoa duy nhất, không bao giờ dùng 2 chữ cái."
        ],
        "answer": "C",
        "solution": "Quy ước quốc tế: Kí hiệu hóa học gồm 1 hoặc 2 chữ cái, chữ đầu viết hoa, chữ thứ hai (nếu có) viết thường."
    },
    {
        "id": 9,
        "content": "Nhận định đúng nhất là:",
        "options": [
            "A. Các nguyên tử thuộc cùng một nguyên tố hóa học thì có tính chất giống nhau.",
            "B. Tập hợp các nguyên tử có cùng số proton đều thuộc cùng một nguyên tố hóa học.",
            "C. Nguyên tố hóa học là những nguyên tử có cùng số neutron khác nhau số proton.",
            "D. Nguyên tố hóa học là những nguyên tử có cùng số electron khác nhau số proton."
        ],
        "answer": "B",
        "solution": "Định nghĩa chuẩn: Các nguyên tử có cùng số proton trong hạt nhân đều thuộc cùng một nguyên tố hóa học."
    },
    {
        "id": 10,
        "content": "Kí hiệu hóa học nào sau đây là của nguyên tố Phosphorus?",
        "options": [
            "A. P.",
            "B. K.",
            "C. Ca.",
            "D. F."
        ],
        "answer": "A",
        "solution": "Phosphorus có kí hiệu hóa học là P (K là Potassium, Ca là Calcium, F là Fluorine)."
    }
]

HOAHOC_TF = [
    {
        "id": 1,
        "prompt": "Câu 1. Xét các phát biểu về khối lượng nguyên tử:",
        "items": [
            {"id": "a", "text": "Đơn vị khối lượng nguyên tử viết tắt là amu.", "ans": True, "reason": "Đúng quy ước quốc tế."},
            {"id": "b", "text": "Khối lượng của một proton gần đúng bằng khối lượng của một neutron và xấp xỉ bằng 1 amu.", "ans": True, "reason": "m_p ≈ m_n ≈ 1 amu."},
            {"id": "c", "text": "Một electron có khối lượng xấp xỉ bằng 5 amu.", "ans": False, "reason": "Khối lượng electron cực kì nhỏ, chỉ khoảng 0,00055 amu."},
            {"id": "d", "text": "Do khối lượng của electron rất nhỏ so với khối lượng proton và neutron nên khối lượng của nguyên tử tập trung chủ yếu ở lớp vỏ.", "ans": False, "reason": "Khối lượng nguyên tử tập trung hầu hết ở hạt nhân."}
        ]
    },
    {
        "id": 2,
        "prompt": "Câu 2. Cho mô hình một số nguyên tử như sau (Boron +5, Sodium +11, Sulfur +16):",
        "image": "mo_hinh_boron_sodium_sulfur.png",
        "items": [
            {"id": "a", "text": "Nguyên tử boron có 5 proton trong hạt nhân.", "ans": True, "reason": "Điện tích hạt nhân là +5 => có 5 proton."},
            {"id": "b", "text": "Nguyên tử sodium có 3 lớp electron.", "ans": True, "reason": "Mô hình sodium có 3 vòng tròn biểu diễn 3 lớp electron (2, 8, 1)."},
            {"id": "c", "text": "Nguyên tử sulfur có 4 electron ở lớp ngoài cùng.", "ans": False, "reason": "Lớp ngoài cùng của sulfur có 6 electron (phân bố 2, 8, 6)."},
            {"id": "d", "text": "Các nguyên tử boron, sodium và sulfur đều có số proton bằng số electron.", "ans": True, "reason": "Nguyên tử luôn trung hòa về điện (số p = số e)."}
        ]
    }
]

HOAHOC_SHORT = [
    {
        "id": 1,
        "content": "Đến nay con người đã tìm ra ...... nguyên tố hóa học.",
        "answer": "118",
        "solution": "Hiện nay bảng tuần hoàn đã ghi nhận chính thức *118 nguyên tố hóa học*."
    },
    {
        "id": 2,
        "content": "Nguyên tử Aluminum có 13 proton. Vậy lớp vỏ ngoài cùng của nguyên tử Aluminum có ...... electron.",
        "answer": "3",
        "solution": "Cấu hình 13 electron phân bố theo lớp: lớp 1 có 2e, lớp 2 có 8e, lớp 3 (ngoài cùng) có *3 electron*."
    },
    {
        "id": 3,
        "content": "Hạt nhân nguyên tử sodium có 11 proton, 11 electron và 12 neutron. Khối lượng của một nguyên tử sodium xấp xỉ bằng ...... amu.",
        "answer": "23",
        "solution": "Khối lượng xấp xỉ bằng tổng số proton và neutron: m ≈ 11 + 12 = *23 amu*."
    },
    {
        "id": 4,
        "content": "Hạt nhân một nguyên tử X có tổng số hạt là 28, trong đó số hạt mang điện nhiều hơn số hạt không mang điện là 8 hạt. Khối lượng nguyên tử X xấp xỉ bằng ...... amu.",
        "answer": "28",
        "solution": "Tổng số hạt trong hạt nhân là số nucleon (p + n) = 28. Vì khối lượng electron không đáng kể nên khối lượng nguyên tử tập trung ở hạt nhân và bằng *28 amu*."
    },
    {
        "id": 5,
        "content": "Mô tả một nguyên tử silicon (hình vẽ bên):\na) Các electron trong nguyên tử silicon được phân bố thành ...... lớp electron theo thứ tự từ trong ra ngoài là 2/8/4.\nb) Tổng số electron của nguyên tử là ......\nc) Nguyên tử silicon có số hiệu nguyên tử Z = 14. Tổng số hạt mang điện trong hạt nhân nguyên tử silicon là ......",
        "image": "mo_hinh_silicon.png",
        "answer": "a) 3 lớp;   b) 14;   c) 14 hạt (proton)",
        "solution": "a) Có *3 lớp* electron (lớp 1: 2e, lớp 2: 8e, lớp 3: 4e).\nb) Tổng số electron = 2 + 8 + 4 = *14*.\nc) Trong hạt nhân, hạt mang điện là proton: Z = 14 => có *14 hạt proton*."
    },
    {
        "id": 6,
        "content": "Cho các kí hiệu hóa học sau: Na, ZN, CA, Al, Ba, CU. Có bao nhiêu kí hiệu hóa học viết sai?",
        "answer": "3",
        "solution": "Có *3 kí hiệu viết sai* gồm: ZN (đúng là Zn), CA (đúng là Ca), CU (đúng là Cu). Na, Al, Ba viết đúng."
    },
    {
        "id": 7,
        "content": "Cho các kí hiệu hóa học sau: H, Li, NA, O, Ne, AL, CA, K, N. Có bao nhiêu kí hiệu hóa học viết đúng?",
        "answer": "6",
        "solution": "Có *6 kí hiệu viết đúng* gồm: H, Li, O, Ne, K, N. (Các kí hiệu NA, AL, CA viết sai quy ước chữ thứ hai)."
    }
]

HOAHOC_ESSAY = [
    {
        "id": 1,
        "content": "Hãy trình bày các bước cơ bản trong tiến trình tìm hiểu tự nhiên.",
        "solution": (
            "Tiến trình tìm hiểu tự nhiên gồm 5 bước cơ bản:\n"
            "• Bước 1: Đề xuất vấn đề cần tìm hiểu (Quan sát, phát hiện hiện tượng và đặt câu hỏi).\n"
            "• Bước 2: Đưa ra dự đoán khoa học để giải quyết vấn đề.\n"
            "• Bước 3: Lập kế hoạch kiểm tra dự đoán (chọn phương pháp, thiết kế thí nghiệm, dụng cụ đo).\n"
            "• Bước 4: Thực hiện kế hoạch kiểm tra dự đoán (tiến hành thí nghiệm, thu thập và xử lí số liệu).\n"
            "• Bước 5: Viết báo cáo, thảo luận và trình bày kết quả."
        )
    },
    {
        "id": 2,
        "content": "Nêu khái niệm nguyên tử, cấu tạo nguyên tử và khối lượng nguyên tử.",
        "solution": (
            "• Khái niệm: Nguyên tử là hạt vô cùng nhỏ bé, tạo nên các chất và trung hòa về điện.\n"
            "• Cấu tạo nguyên tử:\n"
            "  - Hạt nhân: Nằm ở tâm nguyên tử, gồm các hạt proton (mang điện tích +1) và hạt neutron (không mang điện).\n"
            "  - Vỏ nguyên tử: Gồm các hạt electron (mang điện tích -1) chuyển động rất nhanh xung quanh hạt nhân và sắp xếp thành từng lớp.\n"
            "  - Do nguyên tử trung hòa về điện nên: Số proton (p) = Số electron (e).\n"
            "• Khối lượng nguyên tử: Là khối lượng của một nguyên tử, được đo bằng đơn vị amu. Do khối lượng electron rất nhỏ nên khối lượng nguyên tử xấp xỉ bằng khối lượng hạt nhân (m ≈ p + n)."
        )
    },
    {
        "id": 3,
        "content": "Điền từ hoặc cụm từ thích hợp vào các chỗ trống sau:\na) Nguyên tử là những hạt ............, tạo ra chất và trung hòa về điện.\nb) Nguyên tố hóa học là tập hợp những nguyên tử cùng loại, có cùng ............ trong hạt nhân nguyên tử.\nc) Kí hiệu hóa học dùng để biểu diễn ............ và chỉ một ............\nd) Lấy 1/12 khối lượng của ............ làm đơn vị khối lượng nguyên tử, viết tắt là ............\ne) Hai nguyên tử Oxygen được biểu diễn là ............\nf) Khối lượng của một nguyên tử Chlorine là ............",
        "solution": (
            "a) cực kì nhỏ bé\n"
            "b) số proton\n"
            "c) nguyên tố hóa học; nguyên tử của nguyên tố đó\n"
            "d) một nguyên tử carbon; amu\n"
            "e) 2 O\n"
            "f) 35,5 amu"
        )
    },
    {
        "id": 4,
        "content": "Khái niệm nguyên tố hóa học. Viết tên và kí hiệu hóa học của 20 nguyên tố đầu trong bảng tuần hoàn.",
        "solution": (
            "• Khái niệm: Nguyên tố hóa học là tập hợp những nguyên tử cùng loại, có cùng số proton trong hạt nhân.\n"
            "• 20 nguyên tố đầu tiên (Tên - KHHH):\n"
            "  1. Hydrogen (H)      2. Helium (He)       3. Lithium (Li)      4. Beryllium (Be)\n"
            "  5. Boron (B)         6. Carbon (C)        7. Nitrogen (N)      8. Oxygen (O)\n"
            "  9. Fluorine (F)     10. Neon (Ne)        11. Sodium (Na)      12. Magnesium (Mg)\n"
            " 13. Aluminium (Al)   14. Silicon (Si)     15. Phosphorus (P)   16. Sulfur (S)\n"
            " 17. Chlorine (Cl)    18. Argon (Ar)       19. Potassium (K)    20. Calcium (Ca)."
        )
    },
    {
        "id": 5,
        "content": "Kí hiệu hóa học nào sau đây viết sai. Nếu sai hãy sửa lại cho đúng: mG, K, H, Ne, AL, N, O, CA, Fe, HE, Si, P, CU, BA, Zn, AR.",
        "solution": (
            "• Các kí hiệu viết sai và sửa lại:\n"
            "  1. mG  -> sửa thành: *Mg* (Magnesium)\n"
            "  2. AL  -> sửa thành: *Al* (Aluminium)\n"
            "  3. CA  -> sửa thành: *Ca* (Calcium)\n"
            "  4. HE  -> sửa thành: *He* (Helium)\n"
            "  5. CU  -> sửa thành: *Cu* (Copper)\n"
            "  6. BA  -> sửa thành: *Ba* (Barium)\n"
            "  7. AR  -> sửa thành: *Ar* (Argon)\n"
            "• Các kí hiệu đã viết đúng: K, H, Ne, N, O, Fe, Si, P, Zn."
        )
    },
    {
        "id": 6,
        "content": "(a) Các cách viết 2 C, 5 O, 3 Ca, 4 Na, 7 Mg lần lượt chỉ ý gì?\n(b) Dùng chữ số và kí hiệu hóa học diễn đạt ý sau: Ba nguyên tử nitrogen, bảy nguyên tử calcium, bốn nguyên tử potassium, sáu nguyên tử iron, mười lăm nguyên tử aluminium.",
        "solution": (
            "(a) Ý nghĩa cách viết:\n"
            "    • 2 C: 2 nguyên tử carbon.\n"
            "    • 5 O: 5 nguyên tử oxygen.\n"
            "    • 3 Ca: 3 nguyên tử calcium.\n"
            "    • 4 Na: 4 nguyên tử sodium.\n"
            "    • 7 Mg: 7 nguyên tử magnesium.\n\n"
            "(b) Diễn đạt bằng kí hiệu:\n"
            "    • Ba nguyên tử nitrogen: *3 N*\n"
            "    • Bảy nguyên tử calcium: *7 Ca*\n"
            "    • Bốn nguyên tử potassium: *4 K*\n"
            "    • Sáu nguyên tử iron: *6 Fe*\n"
            "    • Mười lăm nguyên tử aluminium: *15 Al*"
        )
    },
    {
        "id": 7,
        "content": "Vẽ sơ đồ cấu tạo nguyên tử có số hạt mang điện tích dương trong hạt nhân là 8 (X), 11 (Y), 20 (Z). Từ sơ đồ cho biết: số proton, số electron, số lớp electron và số electron lớp ngoài cùng của mỗi nguyên tử.",
        "solution": (
            "• Nguyên tử X (hạt nhân +8 - Oxygen):\n"
            "  - Số proton = 8; Số electron = 8.\n"
            "  - Phân bố: Lớp 1 có 2e; lớp 2 có 6e => *Có 2 lớp electron*, *lớp ngoài cùng có 6 electron*.\n\n"
            "• Nguyên tử Y (hạt nhân +11 - Sodium):\n"
            "  - Số proton = 11; Số electron = 11.\n"
            "  - Phân bố: Lớp 1 có 2e; lớp 2 có 8e; lớp 3 có 1e => *Có 3 lớp electron*, *lớp ngoài cùng có 1 electron*.\n\n"
            "• Nguyên tử Z (hạt nhân +20 - Calcium):\n"
            "  - Số proton = 20; Số electron = 20.\n"
            "  - Phân bố: Lớp 1 có 2e; lớp 2 có 8e; lớp 3 có 8e; lớp 4 có 2e => *Có 4 lớp electron*, *lớp ngoài cùng có 2 electron*."
        )
    }
]

# ==============================================================================
# PHẦN 3: SINH HỌC (CHỦ ĐỀ: VẬT SỐNG)
# ==============================================================================

SINHHOC_THEORY = [
    ("1. Kĩ năng học tập môn KHTN & An toàn phòng thí nghiệm", [
        "Quy tắc sử dụng hóa chất an toàn: Không nếm, không ngửi trực tiếp; dùng đúng dụng cụ lấy hóa chất (thìa thủy tinh, kẹp, ống nhỏ giọt); không đổ hóa chất thừa trở lại lọ ban đầu."
    ]),
    ("2. Trao đổi chất và chuyển hóa năng lượng ở sinh vật", [
        "Khái niệm: Trao đổi chất là quá trình sinh vật lấy các chất từ môi trường ngoài, biến đổi chúng thành các chất cần thiết cho cơ thể và thải các chất cặn bã ra môi trường.",
        "Chuyển hóa năng lượng là sự biến đổi năng lượng từ dạng này sang dạng khác (quang năng -> hóa năng -> nhiệt năng, cơ năng...).",
        "Mối quan hệ: Trao đổi chất và chuyển hóa năng lượng luôn gắn liền và diễn ra đồng thời trong tế bào sinh vật."
    ]),
    ("3. Quang hợp ở thực vật", [
        "Khái niệm: Là quá trình lá cây sử dụng chất diệp lục hấp thu năng lượng ánh sáng mặt trời để tổng hợp chất hữu cơ (glucose) từ khí carbon dioxide (CO2) và nước (H2O), đồng thời giải phóng khí oxygen (O2).",
        "Phương trình chữ: Nước + Carbon dioxide --(Ánh sáng, Diệp lục)--> Glucose + Oxygen.",
        "Các yếu tố ảnh hưởng: Ánh sáng, nồng độ CO2, nước, nhiệt độ môi trường."
    ]),
    ("4. Hô hấp tế bào & Bảo quản nông sản", [
        "Khái niệm: Hô hấp tế bào là quá trình phân giải chất hữu cơ (glucose) giải phóng năng lượng ATP cung cấp cho tế bào và cơ thể.",
        "Các yếu tố ảnh hưởng: Hàm lượng nước, nhiệt độ, nồng độ O2, nồng độ CO2.",
        "Ứng dụng: Phơi khô hạt giống, bảo quản lạnh, hút chân không để ức chế hô hấp tế bào nhằm kéo dài thời gian bảo quản."
    ])
]

SINHHOC_MCQ = [
    {
        "id": 1,
        "content": "Phương pháp tìm hiểu tự nhiên là gì?",
        "options": [
            "A. hoạt động con người chủ động tìm tòi, khám phá ra tri thức khoa học.",
            "B. tìm hiểu về thế giới con người, mối quan hệ của con người với môi trường.",
            "C. cách thức tìm hiểu các sự vật, hiện tượng trong tự nhiên và đời sống, chứng minh được các vấn đề trong thực tiễn bằng các dẫn chứng khoa học.",
            "D. cách thức tìm hiểu về thế giới tự nhiên và những ứng dụng khoa học tự nhiên trong cuộc sống thông qua các phương tiện truyền thông như sách, báo, internet,..."
        ],
        "answer": "C",
        "solution": "Phương pháp tìm hiểu tự nhiên là cách thức tìm hiểu các sự vật, hiện tượng tự nhiên và đời sống, chứng minh bằng dẫn chứng khoa học thực nghiệm."
    },
    {
        "id": 2,
        "content": "Để quan sát cấu tạo tế bào biểu bì vảy hành hoặc tế bào khí khổng của lá cây trong phòng thí nghiệm Sinh học, ta bắt buộc phải sử dụng dụng cụ nào sau đây?",
        "options": [
            "A. Kính lúp cầm tay.",
            "B. Kính hiển vi quang học.",
            "C. Kính thiên văn.",
            "D. Ống nhòm tầm xa."
        ],
        "answer": "B",
        "solution": "Tế bào có kích thước hiển vi (vài chục micromet), bắt buộc dùng kính hiển vi quang học mới quan sát được."
    },
    {
        "id": 3,
        "content": "Khi tiến hành đo nhiệt độ cơ thể người hoặc nhiệt độ trong bình thí nghiệm hạt nảy mầm (hô hấp tế bào), dụng cụ nào sau đây là phù hợp nhất?",
        "options": [
            "A. Lực kế.",
            "B. Ống đong.",
            "C. Nhiệt kế.",
            "D. Cân điện tử."
        ],
        "answer": "C",
        "solution": "Nhiệt kế là dụng cụ chuyên dụng để đo nhiệt độ."
    },
    {
        "id": 4,
        "content": "Khi làm tiêu bản tế bào thực vật (ví dụ: tế bào biểu bì vảy hành), dụng cụ nào dùng để nhỏ một giọt nước cất lên lam kính?",
        "options": [
            "A. Ống nhỏ giọt (pipette).",
            "B. Kim mũi mác.",
            "C. Lam kính.",
            "D. La-men."
        ],
        "answer": "A",
        "solution": "Ống nhỏ giọt (pipette) dùng để hút và nhỏ từng giọt chất lỏng chính xác lên lam kính."
    },
    {
        "id": 5,
        "content": "Để cắt các lát mỏng thân cây hoặc lá cây tươi phục vụ cho việc làm tiêu bản quan sát kính hiển vi, ta cần sử dụng dụng cụ sắc bén nào?",
        "options": [
            "A. Kéo cắt cành.",
            "B. Dao lam (lưỡi dao cạo).",
            "C. Kim mũi mác.",
            "D. Kẹp y tế."
        ],
        "answer": "B",
        "solution": "Dao lam có lưỡi cực mỏng và rất sắc, cắt được các lát mô mỏng không làm dập nát cấu trúc tế bào."
    },
    {
        "id": 6,
        "content": "Quá trình trao đổi chất là:",
        "options": [
            "A. Quá trình cơ thể lấy các chất từ môi trường, biến đổi chúng thành các chất cần thiết cho cơ thể, cung cấp năng lượng cho các hoạt động sống, đồng thời thải các chất thải ra môi trường.",
            "B. Quá trình cơ thể trực tiếp lấy các chất từ môi trường sử dụng các chất này cung cấp năng lượng cho các hoạt động sống, đồng thời thải các chất thải ra môi trường.",
            "C. Quá trình cơ thể lấy các chất từ môi trường, biến đổi chúng thành các chất cần thiết cho cơ thể.",
            "D. Quá trình biến đổi các chất trong cơ thể cơ thể thành năng lượng cung cấp cho các hoạt động sống, đồng thời thải các chất thải ra môi trường."
        ],
        "answer": "A",
        "solution": "Trao đổi chất bao gồm cả quá trình lấy vào, biến đổi tổng hợp chất dinh dưỡng tạo năng lượng và bài tiết chất cặn bã ra ngoài."
    },
    {
        "id": 7,
        "content": "Quá trình trao đổi chất và chuyển hóa năng lượng giúp cơ thể sinh vật:",
        "options": [
            "A. Phát triển kích thước theo thời gian.",
            "B. Tồn tại, sinh trưởng, phát triển, sinh sản, cảm ứng và vận động.",
            "C. Tích lũy năng lượng.",
            "D. Vận động tự do trong không gian."
        ],
        "answer": "B",
        "solution": "Trao đổi chất và chuyển hóa năng lượng là điều kiện cốt lõi đảm bảo mọi đặc tính sống của sinh vật."
    },
    {
        "id": 8,
        "content": "Các kĩ năng mà các nhà khoa học sử dụng trong quá trình nghiên cứu thường được gọi là:",
        "options": [
            "A. Kĩ năng thực hành.",
            "B. Kĩ năng tiến trình.",
            "C. Kĩ năng thí nghiệm.",
            "D. Kĩ năng tư duy."
        ],
        "answer": "B",
        "solution": "Các kĩ năng quan sát, đo, dự báo, phân loại, liên kết gọi chung là các kĩ năng tiến trình."
    },
    {
        "id": 9,
        "content": ".......(1).......... là kĩ năng sử dụng một hoặc nhiều giác quan để thu nhận thông tin về các đặc điểm, kích thước, hình dạng, kết cấu, vị trí của các sự vật và hiện tượng trong tự nhiên. Số (1) là:",
        "options": [
            "A. Kĩ năng đo.",
            "B. Kĩ năng quan sát.",
            "C. Kĩ năng phân loại.",
            "D. Kĩ năng dự báo."
        ],
        "answer": "B",
        "solution": "Kĩ năng quan sát là kĩ năng dùng các giác quan (hoặc qua dụng cụ hỗ trợ) để thu thập thông tin sự vật."
    },
    {
        "id": 10,
        "content": "Khi học sinh quan sát hiện tượng cháy rừng hoặc hạn hán, việc phân biệt rõ đâu là hiện tượng tự nhiên thường xảy ra và đâu là thảm họa thiên nhiên gây tác động xấu đến con người và môi trường thể hiện sự vận dụng kĩ năng tiến trình nào?",
        "options": [
            "A. Kĩ năng đo và kĩ năng dự báo.",
            "B. Kĩ năng liên kết và kĩ năng đo.",
            "C. Kĩ năng dự báo và kĩ năng liên kết.",
            "D. Kĩ năng quan sát và kĩ năng phân loại."
        ],
        "answer": "D",
        "solution": "Quan sát hiện tượng và phân chia xếp chúng vào các nhóm khác nhau dựa trên tiêu chí là kĩ năng quan sát và phân loại."
    },
    {
        "id": 11,
        "content": "Kĩ năng liên kết liên quan đến việc sử dụng các yếu tố nào để xác định các mối quan hệ mới của các sự vật và hiện tượng trong tự nhiên?",
        "options": [
            "A. Số liệu quan sát, kết quả phân tích số liệu hoặc dựa vào những điều đã biết.",
            "B. Các dụng cụ đo chiều dài và khối lượng.",
            "C. Các mô hình giả thuyết chưa được kiểm chứng.",
            "D. Các ý kiến chủ quan của cá nhân học sinh."
        ],
        "answer": "A",
        "solution": "Liên kết là kết nối các dữ liệu quan sát, phân tích số liệu và tri thức đã biết để tìm ra quy luật mới."
    },
    {
        "id": 12,
        "content": "Sự biến đổi của năng lượng từ dạng này sang dạng khác (như từ quang năng thành hóa năng, từ hóa năng thành cơ năng, nhiệt năng...) trong cơ thể sinh vật được gọi là:",
        "options": [
            "A. Quá trình trao đổi chất.",
            "B. Quá trình chuyển hóa năng lượng.",
            "C. Quá trình hô hấp tế bào.",
            "D. Quá trình quang hợp."
        ],
        "answer": "B",
        "solution": "Đó chính là khái niệm chuyển hóa năng lượng trong cơ thể sinh vật."
    },
    {
        "id": 13,
        "content": "Phát biểu nào sau đây là đúng khi nói về mối quan hệ giữa trao đổi chất và chuyển hóa năng lượng trong cơ thể sinh vật?",
        "options": [
            "A. Trao đổi chất và chuyển hóa năng lượng luôn diễn ra độc lập với nhau.",
            "B. Chuyển hóa năng lượng diễn ra trước, trao đổi chất diễn ra sau.",
            "C. Trao đổi chất và chuyển hóa năng lượng luôn gắn liền với nhau.",
            "D. Trao đổi chất chỉ diễn ra ở thực vật, chuyển hóa năng lượng chỉ diễn ra ở động vật."
        ],
        "answer": "C",
        "solution": "Trao đổi chất và chuyển hóa năng lượng là hai mặt không thể tách rời của hoạt động sống."
    },
    {
        "id": 14,
        "content": "Ở động vật, khi khí hậu trở nên lạnh hơn bình thường, tốc độ trao đổi chất thường:",
        "options": [
            "A. Tăng lên để sinh nhiệt.",
            "B. Giảm xuống để tiết kiệm năng lượng.",
            "C. Không thay đổi.",
            "D. Bị ngừng hoàn toàn."
        ],
        "answer": "A",
        "solution": "Ở động vật đẳng nhiệt, khi trời rét cơ thể tăng cường trao đổi chất và hô hấp để sinh nhiệt giữ ấm thân nhiệt."
    },
    {
        "id": 15,
        "content": "Sự gia tăng nồng độ khí carbon dioxide trong khí quyển do các hoạt động của con người là một trong những nguyên nhân chính gây biến đổi khí hậu. Ở một giới hạn nhất định, việc tăng nồng độ khí carbon dioxide ban đầu sẽ tác động như thế nào đến quá trình trao đổi chất ở thực vật?",
        "options": [
            "A. Làm ngừng trệ ngay lập tức quá trình quang hợp và gây chết cây.",
            "B. Làm giảm khả năng hấp thụ nước của rễ cây xuống mức bằng không.",
            "C. Có thể làm tăng cường độ quang hợp, thúc đẩy quá trình tổng hợp chất hữu cơ và sinh trưởng của thực vật.",
            "D. Làm cho lá cây không còn khả năng đóng mở khí khổng."
        ],
        "answer": "C",
        "solution": "CO2 là nguyên liệu của quang hợp; trong khoảng giới hạn cho phép, tăng nồng độ CO2 sẽ thúc đẩy quang hợp."
    },
    {
        "id": 16,
        "content": "Phương pháp tìm hiểu tự nhiên được thực hiện qua các bước:\n(1) Đưa ra dự đoán khoa học để giải quyết vấn đề;\n(2) Viết báo cáo;\n(3) Lập kế hoạch kiểm tra dự đoán;\n(4) Đề xuất vấn đề cần tìm hiểu;\n(5) Thực hiện kế hoạch kiểm tra dự đoán.\nThứ tự đúng của phương pháp tìm hiểu tự nhiên là:",
        "options": [
            "A. (1); (2); (3); (4); (5).",
            "B. (4); (1); (3); (5); (2).",
            "C. (5); (4); (3); (2); (1).",
            "D. (3); (4); (1); (5); (2)."
        ],
        "answer": "B",
        "solution": "Trình tự chuẩn: (4) Đề xuất vấn đề -> (1) Đưa dự đoán -> (3) Lập kế hoạch -> (5) Thực hiện kế hoạch -> (2) Viết báo cáo."
    },
    {
        "id": 17,
        "content": "Sau khi học về quang hợp, một học sinh nhận thấy cây xanh thường phát triển tốt ở nơi có đủ ánh sáng. Bạn liên hệ hiện tượng này với vai trò của ánh sáng trong quá trình quang hợp để giải thích sự sinh trưởng của cây. Học sinh đã vận dụng kĩ năng nào?",
        "options": [
            "A. Kĩ năng đo.",
            "B. Kĩ năng liên kết.",
            "C. Kĩ năng phân loại, quan sát.",
            "D. Kĩ năng dự báo."
        ],
        "answer": "B",
        "solution": "Liên hệ giữa hiện tượng cây tươi tốt với kiến thức vai trò ánh sáng trong quang hợp là vận dụng kĩ năng liên kết tri thức."
    },
    {
        "id": 18,
        "content": "Phát biểu nào sau đây không đúng khi nói về vai trò của quá trình trao đổi chất và chuyển hóa năng lượng trong cơ thể?",
        "options": [
            "A. Tạo ra nguồn nguyên liệu cấu tạo nên tế bào và cơ thể.",
            "B. Tạo ra các sản phẩm tham gia hoạt động chức năng của tế bào.",
            "C. Cung cấp năng lượng cho các hoạt động sống của tế bào.",
            "D. Sinh ra nhiệt để giải phóng ra ngoài môi trường."
        ],
        "answer": "D",
        "solution": "Nhiệt giải phóng ra ngoài môi trường là hao phí tự nhiên, không phải là mục đích/vai trò cốt lõi phục vụ sự sống của sinh vật."
    },
    {
        "id": 19,
        "content": "Trao đổi chất ở sinh vật là gì?",
        "options": [
            "A. Sự trao đổi các chất giữa cơ thể với môi trường giúp sinh vật phát triển.",
            "B. Quá trình biến đổi vật lí của các chất từ thể rắn sang thể lỏng trong cơ thể sinh vật.",
            "C. Quá trình sinh vật lấy các chất từ môi trường, biến đổi chúng thành các chất cần thiết cho cơ thể và trả lại môi trường các chất thải.",
            "D. Quá trình biến đổi năng lượng từ dạng này sang dạng khác, giúp sinh vật lớn lên, phát triển và sinh sản."
        ],
        "answer": "C",
        "solution": "Định nghĩa hoàn chỉnh của quá trình trao đổi chất ở sinh vật."
    },
    {
        "id": 20,
        "content": "Nếu nhiệt độ Trái Đất tăng lên do biến đổi khí hậu, quá trình trao đổi chất của sinh vật có thể:",
        "options": [
            "A. Không thay đổi.",
            "B. Diễn ra nhanh hơn hoặc rối loạn, ảnh hưởng đến sự sống.",
            "C. Diễn ra ổn định hơn.",
            "D. Dừng hoàn toàn."
        ],
        "answer": "B",
        "solution": "Nhiệt độ tăng bất thường làm biến đổi tốc độ phản ứng sinh hóa, enzym bị rối loạn, gây ảnh hưởng tiêu cực tới sinh vật."
    }
]

SINHHOC_TF_C21 = {
    "prompt": "Câu 21. Khi tìm hiểu về quá trình quang hợp ở thực vật và các yếu tố ảnh hưởng đến quang hợp, xác định mỗi phát biểu sau đây là Đúng hay Sai:",
    "items": [
        {"id": "a", "text": "Các yếu tố chủ yếu ảnh hưởng đến quá trình quang hợp ở thực vật bao gồm: ánh sáng, nước, hàm lượng khí oxygen và nhiệt độ môi trường.", "ans": False, "reason": "Phải là hàm lượng khí carbon dioxide (CO2), không phải oxygen."},
        {"id": "b", "text": "Khi nhiệt độ môi trường quá cao hoặc quá thấp, hoạt động quang hợp của đa số thực vật bị giảm sút.", "ans": True, "reason": "Nhiệt độ ngoài khoảng tối ưu làm giảm hoạt tính của hệ enzym quang hợp."},
        {"id": "c", "text": "Trong sơ đồ quá trình quang hợp, năng lượng ánh sáng mặt trời (quang năng) đã được diệp lục hấp thụ và chuyển hóa thành hóa năng tích lũy trong các hợp chất hữu cơ.", "ans": True, "reason": "Đúng bản chất chuyển hóa năng lượng trong quang hợp."},
        {"id": "d", "text": "Khi cường độ ánh sáng chiếu vào lá cây tăng vượt quá điểm bão hòa ánh sáng, cường độ quang hợp của cây sẽ tiếp tục tăng mạnh không ngừng.", "ans": False, "reason": "Vượt quá điểm bão hòa thì cường độ quang hợp không tăng nữa, nếu quá mạnh còn bị giảm do cháy lá."},
        {"id": "e", "text": "Ánh sáng là yếu tố bắt buộc của quang hợp, cường độ ánh sáng càng tăng vô hạn thì tốc độ quang hợp của cây càng tăng theo liên tục.", "ans": False, "reason": "Bị giới hạn bởi điểm bão hòa ánh sáng."},
        {"id": "f", "text": "Khí carbon dioxide CO2 là nguyên liệu của quang hợp, nếu nồng độ CO2 trong không khí quá thấp thì quá trình quang hợp sẽ bị giảm hiệu suất hoặc ngừng trệ.", "ans": True, "reason": "CO2 là nguyên liệu chính để cố định carbon tạo chất hữu cơ."},
        {"id": "g", "text": "Nước vừa là nguyên liệu trực tiếp tham gia quang hợp, vừa tham gia điều tiết sự đóng mở khí khổng để điều khiển quá trình trao đổi khí.", "ans": True, "reason": "Đúng hai vai trò thiết yếu của nước."},
        {"id": "h", "text": "Mỗi loài thực vật chỉ quang hợp tốt nhất ở một khoảng nhiệt độ tối ưu nhất định; nhiệt độ quá cao hoặc quá thấp đều làm giảm hoặc ngừng quang hợp.", "ans": True, "reason": "Đúng đặc tính sinh học của từng loài thực vật."},
        {"id": "i", "text": "Trong quá trình quang hợp, năng lượng ánh sáng mặt trời (quang năng) được chuyển hóa thành hóa năng tích lũy trong các chất hữu cơ.", "ans": True, "reason": "Đúng."},
        {"id": "j", "text": "Sản phẩm của quá trình quang hợp chỉ phục vụ riêng cho hoạt động sống của thực vật mà không có vai trò gì đối với động vật.", "ans": False, "reason": "Quang hợp tạo thức ăn và dưỡng khí O2 nuôi dưỡng toàn bộ sinh vật dị dưỡng (động vật, con người)."},
        {"id": "k", "text": "Lá cây là cơ quan chủ yếu thực hiện chức năng quang hợp nhờ chứa tế bào có nhiều lục lạp.", "ans": True, "reason": "Lục lạp chứa diệp lục là bào quan quang hợp."},
        {"id": "l", "text": "Trên biểu bì lá cây có các khí khổng, bộ phận này giúp thu nhận khí oxygen từ môi trường để làm nguyên liệu cho quang hợp.", "ans": False, "reason": "Khí khổng thu nhận khí carbon dioxide (CO2) làm nguyên liệu quang hợp."},
        {"id": "m", "text": "Gân lá chứa mạch gỗ và mạch rây có vai trò vận chuyển nước, muối khoáng đến tế bào quang hợp và đưa chất hữu cơ đi nuôi cây.", "ans": True, "reason": "Mạch gỗ dẫn nước, mạch rây dẫn chất hữu cơ."},
        {"id": "n", "text": "Khi thiếu khí carbon dioxide, quá trình quang hợp có thể bị hạn chế ngay cả khi cây được cung cấp đủ ánh sáng và nước.", "ans": True, "reason": "Thiếu một trong các nguyên liệu thiết yếu thì quang hợp bị ngưng trệ."},
        {"id": "p", "text": "Nhiệt độ môi trường càng tăng thì cường độ quang hợp của cây luôn tăng, kể cả khi nhiệt độ vượt quá giới hạn thích hợp.", "ans": False, "reason": "Nhiệt độ vượt quá giới hạn làm biến tính protein và enzym, làm cây ngừng quang hợp và chết."}
    ]
}

SINHHOC_ESSAY = [
    {
        "id": 1,
        "content": "Vì sao khi bảo quản các loại rau xanh, củ, quả tươi trong gia đình, chúng ta thường bọc chúng trong túi nilon có đục lỗ nhỏ hoặc màng bọc thực phẩm rồi để vào ngăn mát tủ lạnh, chứ không để ở nhiệt độ phòng hoặc để trong ngăn đá?",
        "solution": (
            "• Cơ sở khoa học:\n"
            "  1. Để trong ngăn mát tủ lạnh (khoảng 3°C - 6°C): Nhiệt độ thấp làm giảm tốc độ hô hấp tế bào và ức chế sự sinh sôi của vi khuẩn, nấm mốc, giúp giữ được chất dinh dưỡng và độ tươi ngon lâu hơn.\n"
            "  2. Bọc trong túi nilon đục lỗ hoặc màng bọc thực phẩm: Giúp giữ ẩm cho rau quả, hạn chế sự thoát hơi nước gây héo úa; các lỗ nhỏ giúp thoát khí CO2 và hơi ẩm đọng tránh gây thối nhũn.\n"
            "  3. Không để ở nhiệt độ phòng: Vì ở nhiệt độ phòng, hô hấp tế bào diễn ra mạnh mẽ làm tiêu hao nhanh chất hữu cơ, rau củ nhanh bị già, thối hỏng.\n"
            "  4. Không để trong ngăn đá: Nước trong tế bào rau quả đóng băng sẽ nở ra làm rách màng và vỡ thành tế bào. Khi rã đông, rau quả sẽ bị nhũn nhão, chảy nước và hỏng hoàn toàn."
        )
    },
    {
        "id": 2,
        "content": "Trong sản xuất nông nghiệp, sau những trận mưa lớn gây ngập úng ruộng lúa hoặc các luống rau, người nông dân phải tìm mọi cách để bơm tháo nước ra khỏi ruộng thật nhanh. Em hãy giải thích cơ sở khoa học của hành động này dưới góc độ hô hấp tế bào.",
        "solution": (
            "• Cơ sở khoa học:\n"
            "  1. Khi đất bị ngập nước kéo dài, nước đẩy hết không khí ra ngoài, làm rễ cây bị thiếu hụt khí oxygen (O2) nghiêm trọng.\n"
            "  2. Rễ cây không thể thực hiện hô hấp hiếu khí mà phải chuyển sang hô hấp kị khí (lên men). Quá trình này tạo ra rất ít năng lượng (ATP) và sinh ra các chất độc (rượu ethylic, axit lactic) tích tụ làm đầu rễ bị thối đen, lông hút bị rụng.\n"
            "  3. Khi lông hút bị chết, cây không hút được nước và muối khoáng, dẫn đến hiện tượng cây chết héo úng trên đồng ruộng.\n"
            "  4. Do đó, tháo nước kịp thời giúp rễ cây có đủ không khí (oxygen) để hô hấp bình thường, duy trì chức năng hút nước và chất dinh dưỡng."
        )
    },
    {
        "id": 3,
        "content": "Nêu khái niệm và cho ví dụ về quá trình chuyển hoá năng lượng ở sinh vật.",
        "solution": (
            "• Khái niệm: Chuyển hóa năng lượng ở sinh vật là sự biến đổi năng lượng từ dạng này sang dạng khác để phục vụ cho các hoạt động sống của cơ thể.\n"
            "• Các ví dụ điển hình:\n"
            "  - Ở thực vật: Quá trình quang hợp biến đổi năng lượng ánh sáng mặt trời (quang năng) thành năng lượng dự trữ trong các liên kết hóa học của chất hữu cơ (hóa năng).\n"
            "  - Ở động vật và người: Quá trình hô hấp tế bào biến đổi hóa năng trong thức ăn thành nhiệt năng (sưởi ấm cơ thể), cơ năng (cơ bắp co bóp khi vận động, chạy nhảy) và điện năng (dẫn truyền xung thần kinh)."
        )
    },
    {
        "id": 4,
        "content": "a) Thế nào là quá trình trao đổi chất và chuyển hoá năng lượng ở sinh vật?\nb) Nêu vai trò của quá trình trao đổi chất và chuyển hóa năng lượng đối với cơ thể sinh vật.",
        "solution": (
            "a) Khái niệm:\n"
            "   - Trao đổi chất: Là quá trình sinh vật lấy các chất từ môi trường ngoài, biến đổi chúng thành các chất cần thiết cho tế bào và cơ thể, đồng thời thải các chất cặn bã ra môi trường.\n"
            "   - Chuyển hóa năng lượng: Là sự biến đổi năng lượng từ dạng này sang dạng khác trong tế bào và cơ thể sinh vật.\n\n"
            "b) Vai trò:\n"
            "   - Cung cấp nguyên liệu: Xây dựng tế bào, mô, cơ quan giúp cơ thể lớn lên, sinh trưởng, phát triển và sửa chữa tổn thương.\n"
            "   - Cung cấp năng lượng: Dưới dạng ATP để duy trì các hoạt động sống (vận động, cảm ứng, bài tiết, sinh sản...).\n"
            "   - Bài tiết chất độc hại: Đào thải chất cặn bã sinh ra từ chuyển hóa giúp ổn định môi trường trong cơ thể."
        )
    },
    {
        "id": 5,
        "content": "Tại sao trước khi đem cất giữ (bảo quản) các loại hạt (như hạt thóc, hạt ngô, hạt đậu,...), người nông dân thường phải phơi khô? Giải thích cơ sở khoa học của việc làm này dựa trên hiểu biết về hô hấp tế bào.",
        "solution": (
            "• Cơ sở khoa học:\n"
            "  1. Nước là dung môi và là yếu tố kích hoạt mạnh mẽ hoạt động của các enzym hô hấp tế bào.\n"
            "  2. Khi hạt còn ẩm, hô hấp tế bào diễn ra mạnh mẽ sẽ tiêu hao chất hữu cơ dự trữ trong hạt, làm giảm khối lượng và chất lượng hạt giống. Đồng thời hô hấp sinh nhiệt và ẩm làm hạt dễ bị nảy mầm hoặc bị nấm mốc, vi khuẩn tấn công làm thối hỏng.\n"
            "  3. Khi phơi khô (hạ độ ẩm xuống dưới 12% - 13%), hàm lượng nước thấp làm hô hấp tế bào giảm xuống mức tối thiểu, hạt rơi vào trạng thái ngủ nghỉ, không bị nảy mầm và bảo quản được nhiều năm."
        )
    },
    {
        "id": 6,
        "content": "Việc xây dựng các công viên cây xanh trong các khu đô thị, khu công nghiệp có ý nghĩa như thế nào?",
        "solution": (
            "• Ý nghĩa sinh học và môi trường của công viên cây xanh:\n"
            "  1. Điều hòa khí quyển: Thông qua quá trình quang hợp, cây xanh hấp thụ lượng lớn khí CO2 (gây hiệu ứng nhà kính) và giải phóng khí O2 trong lành cho con người và động vật hô hấp.\n"
            "  2. Lọc sạch không khí: Tán lá cây giữ lại bụi bẩn, hấp thụ các khí độc hại (SO2, NO2, khói xe) và tiết ra các chất diệt khuẩn tự nhiên (phytoncide).\n"
            "  3. Điều hòa nhiệt độ và độ ẩm: Quá trình thoát hơi nước của lá cây giúp làm giảm nhiệt độ không khí đô thị từ 2°C - 4°C, tăng độ ẩm, giảm hiệu ứng đảo nhiệt đô thị.\n"
            "  4. Giảm ô nhiễm tiếng ồn: Tán lá rậm rạp có tác dụng phản xạ và hấp thụ âm thanh, giảm tiếng ồn giao thông và nhà máy.\n"
            "  5. Cảnh quan và tâm lí: Tạo không gian thư giãn, rèn luyện thể thao, nâng cao sức khỏe thể chất và tinh thần cho người dân."
        )
    },
    {
        "id": 7,
        "content": "Theo em, có nên bảo quản rau, quả tươi ở nhiệt độ bằng hoặc thấp hơn 0°C để kéo dài thời gian bảo quản hay không? Giải thích?",
        "solution": (
            "• Trả lời: KHÔNG NÊN.\n"
            "• Giải thích:\n"
            "  1. Ở nhiệt độ bằng hoặc thấp hơn 0°C, nước bên trong tế bào rau quả tươi sẽ bị đóng băng (kết tinh thành các tinh thể đá có thể tích nở rộng).\n"
            "  2. Các tinh thể đá này đâm rách màng sinh chất và làm vỡ cấu trúc thành tế bào.\n"
            "  3. Khi đem ra sử dụng (rã đông), tế bào bị mất nước ồ ạt, cấu trúc mô bị phá hủy khiến rau quả bị mềm nhũn, biến màu, mất vị ngọt tự nhiên và mất hết các vitamin, chất dinh dưỡng.\n"
            "  4. Do đó, rau quả tươi chỉ nên bảo quản ở nhiệt độ mát từ 3°C đến 8°C (ngăn mát tủ lạnh)."
        )
    }
]
