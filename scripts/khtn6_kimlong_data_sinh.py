# -*- coding: utf-8 -*-
"""
Dữ liệu Đề cương ôn tập Giữa kì I - KHTN 6: B. PHẦN RIÊNG - II. PHÂN MÔN SINH HỌC
Trường THCS Kim Long - Năm học 2026 - 2027
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

SINH_THEORY = [
    "Phạm vi kiến thức:",
    "• Bài 18. Tế bào – Đơn vị cơ bản của sự sống.",
    "• Bài 19. Cấu tạo và chức năng các thành phần của tế bào.",
    "• Bài 20. Sự lớn lên và sinh sản của tế bào."
]

# -----------------------------------------------------------------------------
# PHÂN MÔN SINH HỌC - PHẦN I. TRẮC NGHIỆM NHIỀU LỰA CHỌN (20 CÂU)
# -----------------------------------------------------------------------------
SINH_MCQ = [
    # Bài 18. Tế bào – Đơn vị cơ bản của sự sống (Câu 1 - 6)
    {
        "id": 1,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Đơn vị cấu trúc cơ bản của mọi cơ thể sống là",
        "options": ["A. cơ quan.", "B. tế bào.", "C. mô.", "D. hệ cơ quan."],
        "correct": "B",
        "solution": "Mọi cơ thể sinh vật (từ đơn bào đến đa bào) đều được cấu tạo từ một hoặc nhiều tế bào. Tế bào là đơn vị cấu trúc cơ bản của sự sống."
    },
    {
        "id": 2,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Hoạt động nào sau đây là một quá trình sống cơ bản của tế bào?",
        "options": ["A. Bay hơi.", "B. Nóng chảy.", "C. Hô hấp.", "D. Đông đặc."],
        "correct": "C",
        "solution": "Hô hấp tế bào là quá trình phân giải chất hữu cơ giải phóng năng lượng ATP phục vụ cho mọi hoạt động sống. Bay hơi, nóng chảy, đông đặc là hiện tượng chuyển thể vật lí."
    },
    {
        "id": 3,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Vì sao tế bào được coi là đơn vị cơ bản của sự sống?",
        "options": [
            "A. Vì tất cả tế bào đều có hình dạng giống nhau.",
            "B. Vì tế bào chỉ có ở cơ thể động vật và thực vật.",
            "C. Vì tế bào luôn có kích thước lớn và dễ quan sát.",
            "D. Vì mọi cơ thể sống đều được cấu tạo từ tế bào và tế bào thực hiện được các quá trình sống cơ bản."
        ],
        "correct": "D",
        "solution": "Tế bào vừa là đơn vị cấu trúc (tạo nên mọi sinh vật), vừa là đơn vị chức năng thực hiện đầy đủ các quá trình sống cốt lõi: trao đổi chất, lớn lên, sinh sản và cảm ứng."
    },
    {
        "id": 4,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Nhận xét nào sau đây đúng về hình dạng và kích thước của tế bào?",
        "options": [
            "A. Các loại tế bào khác nhau thường có hình dạng và kích thước khác nhau.",
            "B. Tất cả tế bào đều có hình cầu và kích thước giống nhau.",
            "C. Tế bào thực vật luôn có hình dạng giống tế bào động vật.",
            "D. Mọi tế bào đều có thể quan sát rõ bằng mắt thường."
        ],
        "correct": "A",
        "solution": "Tế bào có kích thước và hình dạng vô cùng đa dạng (hình cầu, hình đĩa, hình thoi, hình nhiều cạnh, hình sao có nhánh dài...) để thích nghi với các chức năng sinh lí khác nhau."
    },
    {
        "id": 5,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Tế bào thần kinh ở người có nhiều nhánh dài. Đặc điểm này chủ yếu phù hợp với chức năng nào sau đây?",
        "options": [
            "A. Dự trữ chất dinh dưỡng.",
            "B. Bảo vệ bề mặt cơ thể.",
            "C. Tiếp nhận và truyền tín hiệu thần kinh.",
            "D. Hấp thụ ánh sáng để quang hợp."
        ],
        "correct": "C",
        "solution": "Các sợi nhánh và sợi trục kéo dài giúp tế bào thần kinh (nơron) liên kết mạng lưới, tiếp nhận kích thích và dẫn truyền xung thần kinh cực nhanh đi khắp cơ thể."
    },
    {
        "id": 6,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Một học sinh muốn quan sát hình dạng tế bào vi khuẩn. Dụng cụ phù hợp nhất là",
        "options": ["A. kính lúp cầm tay.", "B. kính hiển vi.", "C. gương phẳng.", "D. kính râm."],
        "correct": "B",
        "solution": "Tế bào vi khuẩn có kích thước hiển vi rất nhỏ (từ 0,5 đến vài micromet), không thể nhìn thấy bằng mắt thường hay kính lúp mà bắt buộc phải dùng kính hiển vi quang học."
    },

    # Bài 19. Cấu tạo và chức năng các thành phần của tế bào (Câu 7 - 15)
    {
        "id": 7,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Thành phần nào sau đây có chức năng bao bọc tế bào chất và tham gia trao đổi chất giữa tế bào với môi trường?",
        "options": ["A. Nhân tế bào.", "B. Lục lạp.", "C. Không bào.", "D. Màng tế bào."],
        "correct": "D",
        "solution": "Màng tế bào (màng sinh chất) là lớp ranh giới bao bọc tế bào chất, có tính thấm chọn lọc giúp kiểm soát việc trao đổi các chất giữa tế bào và môi trường."
    },
    {
        "id": 8,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Phần lớn các hoạt động trao đổi chất của tế bào diễn ra ở",
        "options": ["A. tế bào chất.", "B. màng tế bào.", "C. nhân tế bào.", "D. màng nhân."],
        "correct": "A",
        "solution": "Tế bào chất là môi trường chứa dịch bào tương và các bào quan, là nơi diễn ra phần lớn các phản ứng sinh hóa và hoạt động trao đổi chất của tế bào."
    },
    {
        "id": 9,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Thành phần nào chứa vật chất di truyền và điều khiển các hoạt động sống của tế bào?",
        "options": ["A. Màng tế bào.", "B. Tế bào chất.", "C. Nhân hoặc vùng nhân.", "D. Không bào."],
        "correct": "C",
        "solution": "Nhân (ở sinh vật nhân thực) hoặc vùng nhân (ở sinh vật nhân sơ) chứa vật chất di truyền (DNA), lưu giữ thông tin di truyền và điều khiển mọi hoạt động sống của tế bào."
    },
    {
        "id": 10,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Đặc điểm nào sau đây dùng để phân biệt tế bào nhân sơ với tế bào nhân thực?",
        "options": [
            "A. Tế bào nhân sơ không có tế bào chất.",
            "B. Tế bào nhân sơ chưa có nhân hoàn chỉnh.",
            "C. Tế bào nhân thực không có màng tế bào.",
            "D. Tế bào nhân thực không chứa vật chất di truyền."
        ],
        "correct": "B",
        "solution": "Tế bào nhân sơ chưa có màng nhân bao bọc vật chất di truyền (chỉ có vùng nhân chứa phân tử DNA trần), trong khi tế bào nhân thực đã có nhân hoàn chỉnh với màng nhân."
    },
    {
        "id": 11,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Tế bào của sinh vật nào sau đây thuộc nhóm tế bào nhân sơ?",
        "options": ["A. Vi khuẩn.", "B. Cây lúa.", "C. Con mèo.", "D. Con người."],
        "correct": "A",
        "solution": "Vi khuẩn là sinh vật có cấu tạo tế bào nhân sơ điển hình. Cây lúa, con mèo, con người đều là sinh vật nhân thực."
    },
    {
        "id": 12,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Thành phần nào sau đây có ở tế bào thực vật nhưng không có ở tế bào động vật điển hình?",
        "options": ["A. Màng tế bào.", "B. Tế bào chất.", "C. Nhân tế bào.", "D. Lục lạp."],
        "correct": "D",
        "solution": "Lục lạp (chứa chất diệp lục hấp thu ánh sáng để quang hợp) chỉ có ở tế bào thực vật, không có ở tế bào động vật."
    },
    {
        "id": 13,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Thành tế bào ở thực vật có chức năng chủ yếu nào sau đây?",
        "options": [
            "A. Điều khiển các hoạt động sống của tế bào.",
            "B. Thực hiện quá trình quang hợp.",
            "C. Bảo vệ và quy định hình dạng tế bào.",
            "D. Chứa vật chất di truyền của tế bào."
        ],
        "correct": "C",
        "solution": "Thành tế bào ở thực vật được cấu tạo từ cellulose bền chắc, có chức năng bảo vệ cơ học và cố định, duy trì hình dạng ổn định cho tế bào thực vật."
    },
    {
        "id": 14,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Khi quan sát một tế bào, học sinh nhận thấy tế bào có nhân hoàn chỉnh, thành tế bào và lục lạp. Tế bào đó có thể là",
        "options": [
            "A. tế bào vi khuẩn.",
            "B. tế bào lá cây.",
            "C. tế bào da người.",
            "D. tế bào cơ ở động vật."
        ],
        "correct": "B",
        "solution": "Tế bào có nhân hoàn chỉnh, thành tế bào cellulose và lục lạp là tế bào thực vật quang hợp (ví dụ tế bào mô giậu của lá cây xanh)."
    },
    {
        "id": 15,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Nếu lục lạp trong tế bào lá cây bị tổn thương nghiêm trọng thì hoạt động nào sau đây bị ảnh hưởng trực tiếp nhất?",
        "options": [
            "A. Quang hợp.",
            "B. Phân chia nhân.",
            "C. Vận chuyển nước qua rễ.",
            "D. Trao đổi chất qua màng tế bào."
        ],
        "correct": "A",
        "solution": "Lục lạp là nơi duy nhất thực hiện quá trình quang hợp tổng hợp chất hữu cơ từ ánh sáng mặt trời, CO~2~ và nước. Khi lục lạp hỏng, quang hợp sẽ bị ngưng trệ trực tiếp."
    },

    # Bài 20. Sự lớn lên và sinh sản của tế bào (Câu 16 - 20)
    {
        "id": 16,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Tế bào lớn lên chủ yếu nhờ quá trình nào sau đây?",
        "options": ["A. Bay hơi nước.", "B. Đông đặc các chất.", "C. Trao đổi chất.", "D. Hòa tan màng tế bào."],
        "correct": "C",
        "solution": "Nhờ quá trình trao đổi chất, tế bào hấp thụ các chất dinh dưỡng từ môi trường ngoài và tổng hợp các chất hữu cơ làm gia tăng kích thước và khối lượng các thành phần nội bào."
    },
    {
        "id": 17,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Một tế bào mẹ sau một lần phân chia bình thường tạo thành bao nhiêu tế bào con?",
        "options": ["A. 1 tế bào.", "B. 2 tế bào.", "C. 3 tế bào.", "D. 4 tế bào."],
        "correct": "B",
        "solution": "Từ 1 tế bào mẹ ban đầu, sau khi nhân và tế bào chất phân chia sẽ tạo ra 2 tế bào con có cấu tạo giống nhau."
    },
    {
        "id": 18,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Vì sao tế bào không thể lớn lên mãi?",
        "options": [
            "A. Vì tế bào không thể hấp thụ chất dinh dưỡng.",
            "B. Vì tế bào luôn ngừng hoạt động sau khi được hình thành.",
            "C. Vì tế bào phải chuyển thành cơ quan khi trưởng thành.",
            "D. Vì khi đạt đến kích thước nhất định, tế bào sẽ phân chia tạo tế bào con."
        ],
        "correct": "D",
        "solution": "Khi kích thước tế bào tăng lên, tỉ số diện tích bề mặt trên thể tích (S/V) giảm làm giảm hiệu quả trao đổi chất, nhân không điều khiển kịp thể tích tế bào chất lớn. Khi đó tế bào sẽ dừng lớn và tiến hành phân chia."
    },
    {
        "id": 19,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Sự lớn lên của cây ngô từ cây non thành cây trưởng thành chủ yếu là kết quả của",
        "options": [
            "A. sự tăng kích thước và số lượng tế bào.",
            "B. sự tăng kích thước của từng tế bào mà không tăng số lượng.",
            "C. sự tăng số lượng tế bào mà không tăng kích thước.",
            "D. sự tăng lượng nước bên ngoài các tế bào."
        ],
        "correct": "A",
        "solution": "Cơ thể sinh vật đa bào lớn lên là kết quả phối hợp của hai quá trình: sự phân chia làm tăng số lượng tế bào và sự lớn lên làm tăng kích thước của từng tế bào."
    },
    {
        "id": 20,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Một tế bào ban đầu trải qua 3 lần phân chia liên tiếp. Biết rằng tất cả tế bào con đều tiếp tục phân chia ở mỗi lần, số tế bào thu được là",
        "options": ["A. 3.", "B. 4.", "C. 8.", "D. 6."],
        "correct": "C",
        "solution": "Công thức tính số tế bào con tạo thành sau n lần phân chia từ 1 tế bào: N = 2^n.\nVới n = 3 lần phân chia: N = 2^3 = 8 tế bào."
    }
]

# -----------------------------------------------------------------------------
# PHÂN MÔN SINH HỌC - PHẦN II. TRẮC NGHIỆM ĐÚNG – SAI (6 CÂU)
# -----------------------------------------------------------------------------
SINH_TF = [
    # Bài 18 (Câu 1, 2)
    {
        "id": 1,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Trong giờ học Khoa học tự nhiên, giáo viên giới thiệu hình ảnh tế bào da, tế bào thần kinh ở người, tế bào vi khuẩn và tế bào lá cây. Dựa vào kiến thức đã học, hãy xác định các nhận định sau là đúng hay sai.",
        "sub_items": [
            {"id": "a", "text": "Tế bào là đơn vị cấu trúc cơ bản của mọi cơ thể sống.", "answer": True, "reason": "Tất cả các sinh vật đều được cấu tạo từ một hoặc nhiều tế bào."},
            {"id": "b", "text": "Tất cả các loại tế bào đều có hình dạng và kích thước giống nhau.", "answer": False, "reason": "Các loại tế bào rất khác nhau về kích thước và hình dạng (hình đĩa, hình sợi, hình que, hình sao...)."},
            {"id": "c", "text": "Hình dạng khác nhau của các tế bào thường phù hợp với chức năng mà chúng đảm nhiệm.", "answer": True, "reason": "Cấu trúc hình thái luôn thích nghi với chức năng sinh lí chuyên biệt của từng loại tế bào."},
            {"id": "d", "text": "Vì tế bào vi khuẩn có kích thước nhỏ nên chúng không thể thực hiện các quá trình sống cơ bản.", "answer": False, "reason": "Dù cơ thể chỉ gồm 1 tế bào vi khuẩn đơn sơ nhưng nó vẫn thực hiện đầy đủ mọi quá trình sống độc lập."}
        ]
    },
    {
        "id": 2,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Trong một buổi thực hành, học sinh tìm hiểu kích thước của các loại tế bào và khả năng quan sát chúng bằng mắt thường hoặc kính hiển vi. Hãy xác định các nhận định sau là đúng hay sai.",
        "sub_items": [
            {"id": "a", "text": "Hầu hết các loại tế bào có kích thước rất nhỏ và cần sử dụng kính hiển vi để quan sát.", "answer": True, "reason": "Phần lớn tế bào có kích thước micromet, chỉ quan sát được dưới kính hiển vi."},
            {"id": "b", "text": "Tế bào trứng cá có thể có kích thước đủ lớn để quan sát bằng mắt thường.", "answer": True, "reason": "Tế bào trứng cá (hoặc trứng chim, tép bưởi) có kích thước milimet đến centimet, mắt thường có thể nhìn rõ."},
            {"id": "c", "text": "Tế bào có kích thước càng lớn thì luôn thực hiện các quá trình sống tốt hơn tế bào nhỏ.", "answer": False, "reason": "Tế bào kích thước quá lớn sẽ có tỉ lệ S/V nhỏ, gây khó khăn cho việc trao đổi chất qua màng."},
            {"id": "d", "text": "Nếu muốn quan sát rõ hình dạng tế bào vi khuẩn, học sinh nên sử dụng kính hiển vi thay vì chỉ quan sát bằng mắt thường.", "answer": True, "reason": "Vi khuẩn quá nhỏ, mắt thường hoàn toàn không thể nhìn thấy từng tế bào."}
        ]
    },

    # Bài 19 (Câu 3, 4)
    {
        "id": 3,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Khi tìm hiểu cấu tạo tế bào, học sinh nhận thấy tế bào có các thành phần chính như màng tế bào, tế bào chất, nhân hoặc vùng nhân. Dựa vào chức năng của các thành phần, hãy xác định các nhận định sau là đúng hay sai.",
        "sub_items": [
            {"id": "a", "text": "Màng tế bào bao bọc tế bào chất và tham gia trao đổi chất giữa tế bào với môi trường.", "answer": True, "reason": "Màng tế bào bảo vệ và kiểm soát dòng vật chất ra vào tế bào."},
            {"id": "b", "text": "Phần lớn các hoạt động trao đổi chất của tế bào diễn ra trong nhân tế bào.", "answer": False, "reason": "Hầu hết các phản ứng trao đổi chất diễn ra ở tế bào chất, nhân chỉ đóng vai trò chỉ huy."},
            {"id": "c", "text": "Nhân hoặc vùng nhân chứa vật chất di truyền và điều khiển các hoạt động sống của tế bào.", "answer": True, "reason": "Nhân chứa thông tin di truyền và điều hòa mọi hoạt động sinh hóa nội bào."},
            {"id": "d", "text": "Nếu màng tế bào bị tổn thương nghiêm trọng thì quá trình trao đổi chất giữa tế bào với môi trường vẫn diễn ra bình thường.", "answer": False, "reason": "Màng tế bào rách hoặc hỏng sẽ làm thoát dịch bào tương ra ngoài và tế bào bị chết."}
        ]
    },
    {
        "id": 4,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Một học sinh quan sát sơ đồ tế bào vi khuẩn, tế bào động vật và tế bào thực vật, sau đó đưa ra các nhận định dưới đây. Hãy xác định mỗi nhận định là đúng hay sai.",
        "sub_items": [
            {"id": "a", "text": "Tế bào vi khuẩn thuộc nhóm tế bào nhân sơ vì chưa có nhân hoàn chỉnh.", "answer": True, "reason": "Vật chất di truyền của vi khuẩn chưa có màng nhân bao bọc."},
            {"id": "b", "text": "Tế bào động vật và tế bào thực vật đều thuộc nhóm tế bào nhân thực.", "answer": True, "reason": "Cả hai đều có nhân hoàn chỉnh được bao bọc bởi màng nhân kép."},
            {"id": "c", "text": "Tế bào động vật và tế bào thực vật đều có thành tế bào và lục lạp.", "answer": False, "reason": "Chỉ có tế bào thực vật mới có thành tế bào và lục lạp, tế bào động vật không có hai thành phần này."},
            {"id": "d", "text": "Nếu một tế bào có nhân hoàn chỉnh, thành tế bào và lục lạp thì đó có thể là tế bào lá cây.", "answer": True, "reason": "Đó là các đặc điểm nhận dạng điển hình của tế bào thực vật quang hợp."}
        ]
    },

    # Bài 20 (Câu 5, 6)
    {
        "id": 5,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Khi quan sát sơ đồ sự lớn lên và phân chia tế bào, học sinh nhận thấy tế bào con mới hình thành có kích thước nhỏ, sau đó lớn dần và có thể tiếp tục phân chia. Hãy xác định các nhận định sau là đúng hay sai.",
        "sub_items": [
            {"id": "a", "text": "Tế bào con mới hình thành thường có kích thước nhỏ.", "answer": True, "reason": "Tế bào con ban đầu chỉ bằng khoảng 1/2 kích thước của tế bào mẹ lúc sắp phân chia."},
            {"id": "b", "text": "Tế bào lớn lên nhờ quá trình trao đổi chất với môi trường.", "answer": True, "reason": "Hấp thu chất dinh dưỡng và đồng hóa các chất hữu cơ giúp tế bào tăng kích thước."},
            {"id": "c", "text": "Một tế bào mẹ sau một lần phân chia bình thường tạo thành bốn tế bào con.", "answer": False, "reason": "Một lần phân chia nguyên phân thông thường chỉ tạo ra 2 tế bào con."},
            {"id": "d", "text": "Nếu một tế bào ban đầu phân chia liên tiếp hai lần và các tế bào con đều tiếp tục phân chia thì sẽ tạo thành bốn tế bào.", "answer": True, "reason": "Sau 2 lần phân chia liên tiếp: 2^2 = 4 tế bào con."}
        ]
    },
    {
        "id": 6,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Một học sinh quan sát sự lớn lên của cây ngô và tìm hiểu quá trình thay thế các tế bào già, tế bào chết trong cơ thể người. Từ đó, học sinh đưa ra các nhận định sau. Hãy xác định mỗi nhận định là đúng hay sai.",
        "sub_items": [
            {"id": "a", "text": "Sự lớn lên của cây ngô có liên quan đến sự tăng kích thước và số lượng tế bào.", "answer": True, "reason": "Cây ngô phát triển nhờ phân chia tăng tế bào và sự giãn dài, lớn lên của tế bào mô phân sinh."},
            {"id": "b", "text": "Khi cơ thể người ngừng tăng chiều cao thì tất cả các tế bào trong cơ thể cũng ngừng phân chia.", "answer": False, "reason": "Nhiều mô như biểu bì da, tế bào tủy đỏ tạo máu, niêm mạc ruột vẫn liên tục phân chia suốt đời."},
            {"id": "c", "text": "Sự phân chia tế bào giúp tạo ra các tế bào mới thay thế tế bào già, tế bào chết hoặc bị tổn thương.", "answer": True, "reason": "Đây là vai trò quan trọng trong việc tái sinh, làm lành vết thương và duy trì cân bằng nội môi."},
            {"id": "d", "text": "Nếu các tế bào trong cơ thể không thể phân chia để thay thế tế bào bị tổn thương thì quá trình phục hồi các mô bị tổn thương có thể gặp khó khăn.", "answer": True, "reason": "Không có tế bào mới thay thế thì vết thương không thể lành và cơ thể bị suy giảm chức năng."}
        ]
    }
]

# -----------------------------------------------------------------------------
# PHÂN MÔN SINH HỌC - PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN (15 CÂU)
# -----------------------------------------------------------------------------
SINH_SHORT = [
    # Bài 18 (Câu 1 - 5)
    {
        "id": 1,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Trong các loại tế bào sau: tế bào vi khuẩn, tế bào trứng cá, tế bào da ếch, tế bào lá cây; có bao nhiêu tế bào có kích thước quá nhỏ để quan sát rõ bằng mắt thường?",
        "answer": "3",
        "solution": "Các tế bào có kích thước hiển vi không thể nhìn rõ bằng mắt thường gồm: tế bào vi khuẩn, tế bào da ếch, tế bào lá cây (tổng cộng 3 loại tế bào). Tế bào trứng cá có thể nhìn rõ bằng mắt thường."
    },
    {
        "id": 2,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Một học sinh nhận xét: “Các tế bào có thể khác nhau về hình dạng và kích thước”. Nhận xét này đúng hay sai?",
        "answer": "Đúng",
        "solution": "Nhận xét hoàn toàn đúng vì thế giới tế bào vô cùng đa dạng về cả kích thước (từ vài micromet đến hàng centimet) lẫn hình dạng (hình cầu, hình đĩa, hình thoi, hình sao...)."
    },
    {
        "id": 3,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Trong các đối tượng: viên đá, cây lúa, chiếc bàn, con mèo, có bao nhiêu đối tượng là cơ thể sống được cấu tạo từ tế bào?",
        "answer": "2",
        "solution": "Các đối tượng sống cấu tạo từ tế bào là: cây lúa và con mèo (2 đối tượng). Viên đá và chiếc bàn là vật không sống."
    },
    {
        "id": 4,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Trong các hoạt động sau: trao đổi chất, sinh trưởng, sinh sản, bay hơi; có bao nhiêu hoạt động là quá trình sống cơ bản của tế bào?",
        "answer": "3",
        "solution": "3 quá trình sống cơ bản của tế bào là: trao đổi chất, sinh trưởng và sinh sản. Bay hơi là hiện tượng biến đổi vật lí của chất lỏng."
    },
    {
        "id": 5,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "Một học sinh quan sát các đối tượng gồm: con kiến, hòn đá, cây hoa hồng, giọt nước và cây nấm. Có bao nhiêu đối tượng không được cấu tạo từ tế bào?",
        "answer": "2",
        "solution": "2 đối tượng không sống không được cấu tạo từ tế bào là: hòn đá và giọt nước. Con kiến, cây hoa hồng và cây nấm là cơ thể sinh vật đều cấu tạo từ tế bào."
    },

    # Bài 19 (Câu 6 - 10)
    {
        "id": 6,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Trong tế bào nhân thực, thành phần nào chứa vật chất di truyền và điều khiển các hoạt động sống của tế bào?",
        "answer": "Nhân tế bào",
        "solution": "Nhân tế bào chứa các phân tử DNA mang hệ gen (vật chất di truyền) điều khiển mọi hoạt động sống nội bào."
    },
    {
        "id": 7,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Trong các thành phần: màng tế bào, tế bào chất, nhân, lục lạp, có bao nhiêu thành phần có ở cả tế bào động vật và tế bào thực vật?",
        "answer": "3",
        "solution": "Có 3 thành phần chung cấu tạo nên cả tế bào động vật và thực vật: màng tế bào, tế bào chất và nhân. Lục lạp chỉ có ở thực vật."
    },
    {
        "id": 8,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Một tế bào có màng tế bào, tế bào chất và nhân hoàn chỉnh. Tế bào này thuộc loại nhân gì?",
        "answer": "Tế bào nhân thực",
        "solution": "Tế bào đã có màng bao bọc nhân riêng biệt (nhân hoàn chỉnh) được gọi là tế bào nhân thực."
    },
    {
        "id": 9,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Trong các thành phần: thành tế bào, lục lạp, màng tế bào, nhân, có bao nhiêu thành phần có ở tế bào thực vật nhưng không có ở tế bào động vật điển hình?",
        "answer": "2",
        "solution": "Có 2 thành phần chỉ có ở thực vật mà không có ở động vật điển hình là: thành tế bào (bằng cellulose) và lục lạp."
    },
    {
        "id": 10,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "Tế bào vi khuẩn chưa có nhân hoàn chỉnh nên được xếp vào nhóm tế bào nhân gì?",
        "answer": "Tế bào nhân sơ",
        "solution": "Vi khuẩn chưa có màng nhân bao bọc vật chất di truyền nên được xếp vào nhóm tế bào nhân sơ."
    },

    # Bài 20 (Câu 11 - 15)
    {
        "id": 11,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Một tế bào mẹ sau một lần phân chia bình thường tạo ra bao nhiêu tế bào con?",
        "answer": "2 tế bào con",
        "solution": "Sau 1 lần phân chia, 1 tế bào mẹ tạo ra 2 tế bào con."
    },
    {
        "id": 12,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Một tế bào ban đầu phân chia liên tiếp 3 lần. Biết tất cả tế bào con đều tiếp tục phân chia, số tế bào mới được tạo thành sau 3 lần phân chia là bao nhiêu?",
        "answer": "8 tế bào con",
        "solution": "Số tế bào tạo thành là: 2^3 = 8 tế bào con."
    },
    {
        "id": 13,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Một tế bào ban đầu phân chia liên tiếp hai lần. Biết rằng tất cả tế bào con đều tiếp tục phân chia, số tế bào thu được là bao nhiêu?",
        "answer": "4 tế bào con",
        "solution": "Số tế bào tạo thành là: 2^2 = 4 tế bào con."
    },
    {
        "id": 14,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Có 2 tế bào ban đầu cùng thực hiện một lần phân chia. Số tế bào con được tạo thành là bao nhiêu?",
        "answer": "4 tế bào con",
        "solution": "Mỗi tế bào tạo ra 2 tế bào con, do đó 2 tế bào ban đầu tạo ra: 2 × 2 = 4 tế bào con."
    },
    {
        "id": 15,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "Một tế bào ban đầu phân chia liên tiếp 4 lần. Biết rằng tất cả tế bào con đều tiếp tục phân chia, tổng số tế bào thu được là bao nhiêu?",
        "answer": "16 tế bào con",
        "solution": "Tổng số tế bào thu được là: 2^4 = 16 tế bào con."
    }
]

# -----------------------------------------------------------------------------
# PHÂN MÔN SINH HỌC - B. CÂU HỎI TỰ LUẬN (6 CÂU)
# -----------------------------------------------------------------------------
SINH_ESSAY = [
    # Bài 18 (Câu 1, 2)
    {
        "id": 1,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "a) Tế bào là gì? Vì sao tế bào được coi là đơn vị cơ bản của sự sống?\nb) Nêu một số hình dạng thường gặp của tế bào. Cho ví dụ minh họa.",
        "solution": "a)\n• Khái niệm: Tế bào là đơn vị cấu trúc và chức năng cơ bản của mọi cơ thể sống.\n• Tế bào được coi là đơn vị cơ bản của sự sống vì:\n- Về mặt cấu trúc: Mọi cơ thể sống (từ sinh vật đơn bào như trùng roi, vi khuẩn đến sinh vật đa bào phức tạp như thực vật, động vật, con người) đều được cấu tạo từ một hay nhiều tế bào.\n- Về mặt chức năng: Tế bào là đơn vị nhỏ nhất có khả năng thực hiện đầy đủ các hoạt động sống cốt lõi: trao đổi chất và năng lượng, lớn lên, phân chia (sinh sản) và cảm ứng.\nb) Một số hình dạng thường gặp của tế bào:\n- Hình đĩa lõm hai mặt: tế bào hồng cầu ở người.\n- Hình nhiều nhánh dài (hình sao): tế bào thần kinh (nơron).\n- Hình thoi / hình sợi dài: tế bào cơ (cơ vân, cơ trơn, cơ tim).\n- Hình nhiều cạnh / hình hộp: tế bào biểu bì vảy hành, tế bào biểu bì lá cây.\n- Hình que: tế bào vi khuẩn đường ruột E. coli.\n- Hình xoắn: xoắn khuẩn giang mai; hình dấu phẩy: phẩy khuẩn tả."
    },
    {
        "id": 2,
        "section": "Bài 18. Tế bào – Đơn vị cơ bản của sự sống",
        "content": "a) Tế bào có những quá trình sống cơ bản nào?\nb) Tế bào thần kinh thường có nhiều nhánh dài, trong khi tế bào hồng cầu ở người có dạng đĩa lõm hai mặt. Hãy giải thích sự phù hợp giữa hình dạng và chức năng của hai loại tế bào này.",
        "solution": "a) Các quá trình sống cơ bản của tế bào gồm:\n- Trao đổi chất và chuyển hóa năng lượng (hô hấp tế bào, quang hợp...).\n- Sinh trưởng (tế bào lớn lên về kích thước và khối lượng).\n- Sinh sản (phân chia tế bào tạo ra các tế bào mới).\n- Cảm ứng (khả năng tiếp nhận và phản ứng lại với các kích thích từ môi trường ngoài).\nb) Giải thích sự phù hợp giữa hình dạng và chức năng:\n- Tế bào thần kinh có thân và nhiều sợi nhánh, sợi trục dài: Cấu trúc phân nhánh rộng và kéo dài giúp tế bào tiếp nhận tín hiệu từ các giác quan và dẫn truyền các xung thần kinh nhanh chóng, liên tục giữa các bộ phận khắp cơ thể với hệ thần kinh trung ương (não bộ và tủy sống).\n- Tế bào hồng cầu hình đĩa lõm hai mặt: Hình dạng này giúp tăng diện tích bề mặt tiếp xúc (tăng tỉ lệ S/V), tạo điều kiện cho khí oxygen và carbon dioxide khuếch tán vào - ra dễ dàng và nhanh chóng nhất. Đồng thời, cấu trúc dẹp và mềm dẻo giúp hồng cầu có thể biến dạng linh hoạt để luồn lách qua các mao mạch nhỏ hẹp nhất trong cơ thể mà không bị vỡ hay gây tắc mạch."
    },

    # Bài 19 (Câu 3, 4)
    {
        "id": 3,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "a) Kể tên các thành phần chính của tế bào và nêu chức năng của từng thành phần.\nb) Phân biệt tế bào nhân sơ và tế bào nhân thực dựa vào đặc điểm cấu tạo của nhân tế bào.",
        "solution": "a) Ba thành phần chính cấu tạo nên tế bào:\n1. Màng tế bào (màng sinh chất): Bao bọc bảo vệ tế bào chất, kiểm soát sự trao đổi chất có chọn lọc giữa tế bào với môi trường bên ngoài.\n2. Tế bào chất: Là vùng không gian nằm giữa màng tế bào và nhân/vùng nhân, chứa dịch bào tương và các bào quan, là nơi diễn ra phần lớn các hoạt động trao đổi chất của tế bào.\n3. Nhân (hoặc vùng nhân): Chứa vật chất di truyền (phân tử DNA), lưu giữ thông tin di truyền và là trung tâm chỉ huy, điều khiển mọi hoạt động sống của tế bào.\nb) Phân biệt tế bào nhân sơ và tế bào nhân thực dựa vào nhân:\n- Tế bào nhân sơ: Chưa có màng nhân bao bọc vật chất di truyền. Phân tử DNA nằm trần tập trung tại một vùng gọi là \"vùng nhân\". (Gặp ở vi khuẩn).\n- Tế bào nhân thực: Đã có nhân hoàn chỉnh, được bao bọc bởi một màng nhân kép ngăn cách với tế bào chất, bên trong chứa dịch nhân và nhân con. (Gặp ở thực vật, động vật, nấm, nguyên sinh vật)."
    },
    {
        "id": 4,
        "section": "Bài 19. Cấu tạo và chức năng các thành phần của tế bào",
        "content": "a) Nêu những điểm giống nhau và khác nhau về cấu tạo giữa tế bào thực vật và tế bào động vật.\nb) Khi quan sát một tế bào dưới kính hiển vi, học sinh nhận thấy tế bào có thành tế bào, lục lạp và nhân hoàn chỉnh. Theo em, đó có thể là tế bào động vật hay tế bào thực vật? Giải thích.",
        "solution": "a)\n• Giống nhau:\n- Đều thuộc nhóm tế bào nhân thực (có nhân hoàn chỉnh được bao bọc bởi màng nhân).\n- Đều có 3 thành phần cấu trúc nền tảng: màng tế bào, tế bào chất và nhân tế bào.\n• Khác nhau:\n- Tế bào thực vật: Có thành tế bào bằng cellulose bền vững quy định hình dạng tế bào; có bào quan lục lạp chứa chất diệp lục quang hợp; có một không bào trung tâm kích thước lớn.\n- Tế bào động vật: Không có thành tế bào cellulose (màng tế bào mềm dẻo, hình dạng linh hoạt); không có lục lạp; không bào nhỏ hoặc không có; có bào quan trung thể.\nb)\n- Tế bào mà học sinh quan sát được là tế bào thực vật (ví dụ tế bào mô giậu lá cây).\n- Giải thích: Vì tế bào động vật không bao giờ có thành tế bào và lục lạp. Chỉ có tế bào thực vật mới sở hữu đồng thời thành tế bào bằng cellulose và lục lạp giúp quang hợp."
    },

    # Bài 20 (Câu 5, 6)
    {
        "id": 5,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "a) Trình bày quá trình lớn lên của tế bào.\nb) Mô tả quá trình phân chia của tế bào. Từ một tế bào mẹ, sau một lần phân chia tạo thành bao nhiêu tế bào con?",
        "solution": "a) Quá trình lớn lên của tế bào:\n- Tế bào con mới được tạo thành sau phân chia có kích thước và thể tích nhỏ.\n- Tế bào tiến hành hấp thu các chất dinh dưỡng từ môi trường xung quanh thông qua màng tế bào, sau đó chuyển hóa và tổng hợp nên các chất hữu cơ nội bào làm gia tăng nhanh chóng khối lượng và thể tích của màng, tế bào chất và nhân.\n- Tế bào lớn dần lên cho đến khi đạt kích thước giới hạn thì dừng lớn và bước vào quá trình phân chia.\nb) Quá trình phân chia của tế bào:\n- Diễn biến gồm 2 giai đoạn chính liên tiếp:\n  + Giai đoạn 1 (Phân chia nhân): Vật chất di truyền trong nhân nhân đôi rồi tách đều về hai phía, chia nhân mẹ thành 2 nhân con.\n  + Giai đoạn 2 (Phân chia tế bào chất): Tế bào chất phân chia bằng cách thắt màng sinh chất ở giữa (đối với tế bào động vật) hoặc hình thành vách ngăn từ giữa ra ngoài (đối với tế bào thực vật) để ngăn tế bào chất thành 2 phần riêng biệt.\n- Kết quả: Từ 1 tế bào mẹ ban đầu, sau 1 lần phân chia bình thường sẽ tạo ra 2 tế bào con có cấu tạo và thông tin di truyền hoàn toàn giống nhau và giống hệt tế bào mẹ ban đầu."
    },
    {
        "id": 6,
        "section": "Bài 20. Sự lớn lên và sinh sản của tế bào",
        "content": "a) Nêu ý nghĩa của sự lớn lên và sinh sản của tế bào đối với cơ thể sinh vật.\nb) Một tế bào ban đầu trải qua 4 lần phân chia liên tiếp. Biết rằng tất cả các tế bào con đều tiếp tục phân chia ở mỗi lần. Hãy tính số tế bào thu được sau 4 lần phân chia và trình bày cách tính.",
        "solution": "a) Ý nghĩa của sự lớn lên và sinh sản của tế bào:\n- Giúp cơ thể sinh vật sinh trưởng và phát triển: Đối với sinh vật đơn bào, sự lớn lên của tế bào làm tăng kích thước cơ thể; đối với sinh vật đa bào, sự phân chia làm tăng số lượng tế bào kết hợp với sự lớn lên của tế bào giúp cơ thể tăng trưởng kích thước và khối lượng.\n- Thay thế và phục hồi: Tạo ra các tế bào mới thay thế cho những tế bào đã già cỗi, bị chết tự nhiên hoặc bị tổn thương, giúp làm lành các vết thương và tái tạo các mô cơ quan.\n- Duy trì nòi giống: Là cơ sở cho sự sinh sản vô tính ở sinh vật đơn bào và quá trình tạo giao tử sinh sản ở sinh vật đa bào.\nb) Tính số tế bào con sau 4 lần phân chia:\n- Áp dụng công thức tính số tế bào con tạo ra sau n lần phân chia liên tiếp từ a tế bào ban đầu:\n  N = a × 2^n\n- Với a = 1 tế bào và n = 4 lần phân chia:\n  N = 1 × 2^4 = 2 × 2 × 2 × 2 = 16 (tế bào con).\n- Kết luận: Sau 4 lần phân chia liên tiếp, từ 1 tế bào ban đầu sẽ thu được 16 tế bào con."
    }
]
