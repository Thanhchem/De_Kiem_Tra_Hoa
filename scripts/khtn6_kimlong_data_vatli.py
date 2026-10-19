# -*- coding: utf-8 -*-
"""
Dữ liệu Đề cương ôn tập Giữa kì I - KHTN 6: B. PHẦN RIÊNG - I. PHÂN MÔN VẬT LÍ
Trường THCS Kim Long - Năm học 2026 - 2027
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

VATLI_THEORY = [
    "I. Lí thuyết: Học sinh học từ: Bài 40. Lực là gì? đến bài 41. Biểu diễn lực.",
    "II. Bài tập: Học sinh làm tất cả các bài tập trong sách bài tập từ: Bài 40. Lực là gì? đến bài 41. Biểu diễn lực."
]

# -----------------------------------------------------------------------------
# PHÂN MÔN VẬT LÍ - 1. TRẮC NGHIỆM NHIỀU LỰA CHỌN (16 CÂU)
# -----------------------------------------------------------------------------
VATLI_MCQ = [
    {
        "id": 1,
        "content": "Để đo lực, người ta dùng",
        "options": ["A. lực kế.", "B. thước đo.", "C. cân.", "D. nhiệt kế."],
        "correct": "A",
        "solution": "Lực kế là dụng cụ dùng để đo độ lớn của lực."
    },
    {
        "id": 2,
        "content": "Đơn vị đo lực trong hệ đo lường hợp pháp của nước ta là",
        "options": ["A. newton (N).", "B. metre (m).", "C. giây (s).", "D. kilogram (kg)."],
        "correct": "A",
        "solution": "Đơn vị đo lực là newton, kí hiệu là N."
    },
    {
        "id": 3,
        "content": "Một bạn chơi trò nhảy dây. Bạn đó nhảy lên được là do",
        "options": [
            "A. lực của chân đẩy bạn đó nhảy lên.",
            "B. lực của đất tác dụng lên chân bạn đó.",
            "C. chân bạn đó tiếp xúc với đất.",
            "D. lực của đất tác dụng lên dây."
        ],
        "correct": "B",
        "solution": "Khi bạn đó dùng chân đạp mạnh xuống đất, mặt đất tác dụng trở lại chân một phản lực hướng lên, lực này đẩy bạn đó nhảy lên không trung."
    },
    {
        "id": 4,
        "content": "Treo vật vào đầu dưới của một lò xo, lò xo dãn ra. Khi đó",
        "options": [
            "A. lò xo tác dụng vào vật một lực đẩy.",
            "B. vật tác dụng vào lò xo một lực nén.",
            "C. lò xo tác dụng vào vật một lực nén.",
            "D. vật tác dụng vào lò xo một lực kéo."
        ],
        "correct": "D",
        "solution": "Do lực hút của Trái Đất kéo vật xuống dưới, vật tác dụng vào đầu dưới của lò xo một lực kéo làm lò xo bị dãn dài ra."
    },
    {
        "id": 5,
        "content": "Khi một vận động viên bắt đầu đẩy quả tạ, vận động viên đã tác dụng vào quả tạ một",
        "options": ["A. lực đẩy.", "B. lực nén.", "C. lực kéo.", "D. lực uốn."],
        "correct": "A",
        "solution": "Cánh tay của vận động viên tác dụng vào quả tạ một lực đẩy làm quả tạ chuyển động bay đi."
    },
    {
        "id": 6,
        "content": "Hoạt động nào dưới đây không cần dùng đến lực?",
        "options": [
            "A. Đọc một trang sách.",
            "B. Kéo một gàu nước.",
            "C. Nâng một tấm gỗ.",
            "D. Đẩy một chiếc xe."
        ],
        "correct": "A",
        "solution": "Đọc một trang sách là hoạt động tư duy và thị giác, không tác dụng lực cơ học làm biến đổi chuyển động hay biến dạng vật thể."
    },
    {
        "id": 7,
        "content": "Một em bé thả một quả bóng cao su xuống sàn nhà. Khi quả bóng chạm sàn nhà thì lực của sàn nhà tác dụng lên quả bóng",
        "options": [
            "A. chỉ làm cho quả bóng biến đổi chuyển động.",
            "B. chỉ làm cho quả bóng biến dạng.",
            "C. vừa làm cho quả bóng biến dạng, vừa làm cho quả bóng biến đổi chuyển động.",
            "D. không làm quả bóng biến dạng cũng không làm biến đổi chuyển động của quả bóng."
        ],
        "correct": "C",
        "solution": "Khi va chạm sàn nhà, quả bóng cao su bị móp lại (biến dạng) đồng thời bị đổi hướng chuyển động bật nảy lên trên (biến đổi chuyển động)."
    },
    {
        "id": 8,
        "content": "Khi có một lực tác dụng lên quả bóng đang chuyển động trên sân thì tốc độ của quả bóng sẽ",
        "options": [
            "A. không thay đổi.",
            "B. tăng dần.",
            "C. giảm dần.",
            "D. tăng dần hoặc giảm dần."
        ],
        "correct": "D",
        "solution": "Nếu lực tác dụng cùng hướng với hướng chuyển động thì tốc độ tăng dần; nếu lực tác dụng ngược hướng chuyển động (lực cản, lực ma sát) thì tốc độ giảm dần."
    },
    {
        "id": 9,
        "content": "Lực không gây ra tác dụng nào trong các tác dụng sau đây?",
        "options": [
            "A. Làm cho vật chuyển động nhanh lên.",
            "B. Làm cho vật chuyển động chậm lại.",
            "C. Làm cho vật biến dạng.",
            "D. Làm cho vật biến mất."
        ],
        "correct": "D",
        "solution": "Tác dụng của lực chỉ có thể làm biến đổi chuyển động của vật (nhanh lên, chậm lại, đổi hướng) hoặc làm cho vật bị biến dạng. Lực không thể làm cho vật bị biến mất."
    },
    {
        "id": 10,
        "content": "Trường hợp nào sau đây vật không bị biến dạng khi chịu tác dụng của lực?",
        "options": [
            "A. Cửa kính bị vỡ khi bị va đập mạnh.",
            "B. Đất xốp khi được cày xới cẩn thận.",
            "C. Viên bi sắt bị búng và lăn về phía trước.",
            "D. Tờ giấy bị nhàu khi ta vò nó lại."
        ],
        "correct": "C",
        "solution": "Viên bi sắt có độ cứng rất lớn, lực búng tay chỉ làm viên bi bắt đầu chuyển động lăn về phía trước (biến đổi chuyển động), không làm viên bi bị biến dạng."
    },
    {
        "id": 11,
        "content": "Khi quả bóng đập vào một bức tường, lực do tường tác dụng lên quả bóng:",
        "options": [
            "A. Chỉ làm biến đổi chuyển động của quả bóng.",
            "B. Chỉ làm biến dạng quả bóng.",
            "C. Vừa làm biến đổi chuyển động vừa làm biến dạng quả bóng.",
            "D. Không làm biến đổi chuyển động và không làm biến dạng quả bóng"
        ],
        "correct": "C",
        "solution": "Lực do tường tác dụng làm quả bóng bị móp lại (biến dạng) đồng thời bật dội ngược trở ra (biến đổi chuyển động)."
    },
    {
        "id": 12,
        "content": "Khi có một lực tác dụng lên quả bóng đang chuyển động trên sân thì tốc độ của quả bóng sẽ",
        "options": [
            "A. không thay đổi.",
            "B. tăng dần.",
            "C. giảm dần.",
            "D. tăng dần hoặc giảm dần."
        ],
        "correct": "D",
        "solution": "Tùy thuộc vào hướng và độ lớn của lực so với chuyển động ban đầu của quả bóng, tốc độ của quả bóng có thể tăng dần hoặc giảm dần."
    },
    {
        "id": 13,
        "content": "Trong các trường hợp sau đây, trường hợp nào xuất hiện lực không tiếp xúc?",
        "options": [
            "A. Em bé đẩy cho chiếc xe đồ chơi rơi xuống đất.",
            "B. Gió thổi làm thuyền chuyển động.",
            "C. Cầu thủ đá quả bóng bay vào gôn.",
            "D. Quả táo rơi từ trên cây xuống."
        ],
        "correct": "D",
        "solution": "Quả táo rơi từ trên cây xuống là do lực hút của Trái Đất (trọng lực). Trái Đất hút quả táo mà không cần có sự tiếp xúc trực tiếp giữa Trái Đất và quả táo khi nó đang ở trên cành, đây là lực không tiếp xúc."
    },
    {
        "id": 14,
        "content": "Phát biểu nào sau đây là không đúng?",
        "options": [
            "A. Lực chỉ có tác dụng làm vật biến đổi chuyển động.",
            "B. Tác dụng đẩy, kéo của vật này lên vật khác gọi là lực.",
            "C. Lực được phân thành: lực không tiếp xúc và lực tiếp xúc.",
            "D. Lực có thể vừa làm cho vật biến dạng vừa làm cho vật biến đổi chuyển động."
        ],
        "correct": "A",
        "solution": "Lực vừa có thể làm biến đổi chuyển động, vừa có thể làm biến dạng vật. Do đó khẳng định lực \"chỉ có tác dụng làm vật biến đổi chuyển động\" là không đúng."
    },
    {
        "id": 15,
        "content": "Trong các trường hợp sau đây, trường hợp nào xuất hiện lực tiếp xúc?",
        "options": [
            "A. Hai thanh nam châm hút nhau.",
            "B. Hai thanh nam châm đẩy nhau.",
            "C. Mặt Trăng quay quanh Trái Đất.",
            "D. Mẹ em ấn nút công tắc bật đèn."
        ],
        "correct": "D",
        "solution": "Ngón tay chạm và ép trực tiếp vào nút công tắc điện, đây là lực tiếp xúc. A, B là lực từ; C là lực hấp dẫn (lực không tiếp xúc)."
    },
    {
        "id": 16,
        "content": "Đơn vị nào sau đây là đơn vị của lực?",
        "options": ["A. Kilôgam (kg)", "B. Centimét (cm)", "C. Niutơn (N)", "D. Lít (L)"],
        "correct": "C",
        "solution": "Đơn vị chuẩn của lực là Niutơn, kí hiệu N."
    }
]

# -----------------------------------------------------------------------------
# PHÂN MÔN VẬT LÍ - 2. CÂU TRẢ LỜI ĐÚNG - SAI (3 CÂU)
# -----------------------------------------------------------------------------
VATLI_TF = [
    {
        "id": 17,
        "content": "Một quả bóng cao su bị nén khi búng tay vào nó. Theo lý thuyết, phát biểu sau đúng hay sai?",
        "sub_items": [
            {"id": "a", "text": "Độ biến dạng của quả bóng sẽ lớn hơn nếu lực búng tay nhẹ hơn.", "answer": False, "reason": "Lực tác dụng càng lớn thì độ biến dạng càng lớn; lực nhẹ hơn thì độ biến dạng phải nhỏ hơn."},
            {"id": "b", "text": "Quả bóng bị búng tay tạo ra lực ép lên bề mặt quả bóng.", "answer": True, "reason": "Lực từ ngón tay tác dụng ép trực tiếp lên bề mặt tiếp xúc của quả bóng."},
            {"id": "c", "text": "Độ biến dạng của quả bóng phụ thuộc hoàn toàn vào hướng của lực búng tay.", "answer": False, "reason": "Độ biến dạng phụ thuộc chủ yếu vào độ lớn của lực và tính đàn hồi của vật liệu cao su."},
            {"id": "d", "text": "Chỉ cần búng tay nhẹ, quả bóng cũng có thể bị biến dạng vĩnh viễn.", "answer": False, "reason": "Bóng cao su có tính đàn hồi cao, khi búng nhẹ bóng chỉ biến dạng tạm thời và ngay lập tức lấy lại hình dạng ban đầu."}
        ]
    },
    {
        "id": 18,
        "content": "Một người bóp tay vào một quả bóng cao su, hãy phân tích tác dụng của lực lên quả bóng.",
        "sub_items": [
            {"id": "a", "text": "Nếu người này tác dụng lực nhẹ hơn, độ biến dạng của quả bóng sẽ lớn hơn.", "answer": False, "reason": "Lực tác dụng càng nhỏ thì độ biến dạng của quả bóng càng nhỏ."},
            {"id": "b", "text": "Lực tác dụng của tay làm quả bóng cao su biến dạng.", "answer": True, "reason": "Quả bóng bị móp méo, thay đổi hình dạng dưới tác dụng của lực ép từ bàn tay."},
            {"id": "c", "text": "Trong thực tế, độ biến dạng của quả bóng còn có thể phụ thuộc vào thời gian tác dụng của lực và nhiệt độ môi trường.", "answer": True, "reason": "Nhiệt độ ảnh hưởng đến độ mềm dẻo của cao su và thời gian duy trì lực ép cũng ảnh hưởng đến biến dạng vật liệu."},
            {"id": "d", "text": "Độ biến dạng của quả bóng phụ thuộc hoàn toàn vào lực ép, không phụ thuộc vào hình dạng của quả bóng.", "answer": False, "reason": "Hình dạng, độ dày thành bóng và áp suất khí bên trong quả bóng đều ảnh hưởng đến khả năng biến dạng của nó."}
        ]
    },
    {
        "id": 19,
        "content": "Trong các phát biểu sau, phát biểu nào đúng, phát biểu nào sai?",
        "sub_items": [
            {"id": "a", "text": "Tác dụng đẩy, kéo của vật này lên vật khác gọi là lực.", "answer": True, "reason": "Đây là định nghĩa cơ bản về lực trong SGK KHTN 6."},
            {"id": "b", "text": "Khi vật A đẩy hoặc kéo vật B ta nói vật B tác dụng lực lên vật A.", "answer": False, "reason": "Khi vật A đẩy hoặc kéo vật B thì vật A tác dụng lực lên vật B."},
            {"id": "c", "text": "Lực là nguyên nhân gây ra chuyển động của vật.", "answer": False, "reason": "Lực không phải là nguyên nhân gây ra hoặc duy trì chuyển động, mà lực là nguyên nhân làm biến đổi chuyển động của vật."},
            {"id": "d", "text": "Lực được phân thành lực tiếp xúc và lực không tiếp xúc.", "answer": True, "reason": "Dựa vào đặc điểm tiếp xúc, lực được phân loại thành lực tiếp xúc và lực không tiếp xúc."}
        ]
    }
]

# -----------------------------------------------------------------------------
# PHÂN MÔN VẬT LÍ - 3. ĐIỀN TỪ VÀO CHỖ TRỐNG (CÂU 20)
# -----------------------------------------------------------------------------
VATLI_FILL = {
    "id": 20,
    "content": "Lựa chọn các từ sau: *lực kéo*, *lực nén*, *lực đẩy* và điền vào chỗ trống:\na) Bạn My đã tác dụng vào thước nhựa một ...... làm thước nhựa bị uốn cong.\nb) Để nâng tấm bê tông lên, cần cẩu đã tác dụng vào tấm bê tông một ......\nc) Đầu tàu đã tác dụng vào toa một .......\nd) Gió đã tác dụng vào dù của người nhảy dù một......",
    "solution": "a) Bạn My đã tác dụng vào thước nhựa một *lực nén* (hoặc lực đẩy) làm thước nhựa bị uốn cong.\nb) Để nâng tấm bê tông lên, cần cẩu đã tác dụng vào tấm bê tông một *lực kéo*.\nc) Đầu tàu đã tác dụng vào toa một *lực kéo*.\nd) Gió đã tác dụng vào dù của người nhảy dù một *lực đẩy*."
}

# -----------------------------------------------------------------------------
# PHÂN MÔN VẬT LÍ - 4. TỰ LUẬN (CÂU 21)
# -----------------------------------------------------------------------------
VATLI_ESSAY = [
    {
        "id": 21,
        "content": "Khi bạn dùng tay ép một chiếc lò xo bi-a hoặc bóp méo vỏ lon nước ngọt bằng nhôm, lực do tay bạn đã gây ra kết quả gì?",
        "solution": "- Khi dùng tay ép lò xo bi-a: Lực của bàn tay làm lò xo bị nén ngắn lại (gây ra sự biến dạng đàn hồi).\n- Khi dùng tay bóp mạnh vỏ lon nước ngọt bằng nhôm: Lực của bàn tay làm vỏ lon bị móp méo, gãy gập (gây ra sự biến dạng vĩnh viễn).\n=> Kết luận chung: Trong cả hai trường hợp trên, lực do tay tác dụng đều đã gây ra tác dụng làm biến dạng vật thể."
    }
]
