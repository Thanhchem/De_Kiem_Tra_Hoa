# -*- coding: utf-8 -*-
"""
Dữ liệu Đề cương ôn tập Giữa kì I - KHTN 6: B. PHẦN RIÊNG - II. PHÂN MÔN HOÁ HỌC
Trường THCS Kim Long - Năm học 2026 - 2027
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

HOA_THEORY = [
    "1. NỘI DUNG ÔN TẬP:",
    "• Bài 9. Sự đa dạng của chất",
    "• Bài 10. Các thể của chất và sự chuyển thể"
]

# -----------------------------------------------------------------------------
# PHÂN MÔN HÓA HỌC - PHẦN 1: CÂU HỎI NHIỀU LỰA CHỌN (24 CÂU)
# -----------------------------------------------------------------------------
HOA_MCQ = [
    {
        "id": 1,
        "content": "Tất cả các trường hợp nào sau đây đều là chất?",
        "options": [
            "A. Đường mía, muối ăn, con dao.",
            "B. Con dao, đôi đũa, cái thìa nhôm.",
            "C. Nhôm, muối ăn, đường.",
            "D. Con dao, đôi đũa, muối ăn."
        ],
        "correct": "C",
        "solution": "Nhôm, muối ăn, đường đều là các chất hóa học tinh khiết hoặc hợp chất cụ thể. Con dao, đôi đũa, cái thìa là các vật thể."
    },
    {
        "id": 2,
        "content": "Cho các vật thể: ngôi nhà, con gà, cây lúa, viên gạch, nước biển, xe đạp. Trong các vật thể đã cho, những vật thể do con người tạo ra là:",
        "options": [
            "A. Ngôi nhà, con gà, xe đạp.",
            "B. Con gà, nước biển, xe đạp.",
            "C. Ngôi nhà, viên gạch, xe đạp.",
            "D. Con gà, viên gạch, xe đạp."
        ],
        "correct": "C",
        "solution": "Vật thể nhân tạo là vật thể do con người chế tạo ra gồm: ngôi nhà, viên gạch, xe đạp. Con gà, cây lúa, nước biển là vật thể tự nhiên."
    },
    {
        "id": 3,
        "content": "Phát biểu nào dưới đây là đúng khi nói về những đặc điểm của chất rắn:",
        "options": [
            "A. Có khối lượng, hình dạng xác định, không có thể tích xác định.",
            "B. Có khối lượng xác định, hình dạng và thể tích không xác định.",
            "C. Có khối lượng, hình dạng, thể tích xác định.",
            "D. Có khối lượng xác định, hình dạng và thể tích không xác định."
        ],
        "correct": "C",
        "solution": "Chất rắn có các hạt liên kết rất chặt chẽ, dao động tại chỗ, nên chất rắn có khối lượng, hình dạng và thể tích xác định."
    },
    {
        "id": 4,
        "content": "Không khí xung quanh ta có đặc điểm gì?",
        "options": [
            "A. Không có hình dạng và thể tích xác định.",
            "B. Có hình dạng và thể tích xác định.",
            "C. Có hình dạng xác định, không có thể tích xác định.",
            "D. Không có hình dạng xác định, có thể tích xác định."
        ],
        "correct": "A",
        "solution": "Không khí tồn tại ở thể khí (thể hơi), có khối lượng xác định nhưng không có hình dạng và không có thể tích xác định (chiếm toàn bộ thể tích bình chứa)."
    },
    {
        "id": 5,
        "content": "Cho các mẫu chất có đặc điểm sau: có khối lượng xác định, không có thể tích xác định và không có hình dạng xác định mà mang hình dạng của vật chứa nó. Mẫu chất đó đang ở thể nào?",
        "options": ["A. Rắn.", "B. Lỏng.", "C. Khí.", "D. Không xác định."],
        "correct": "C",
        "solution": "Chất ở thể khí có khối lượng xác định, chuyển động hỗn loạn, lan tỏa khắp nơi nên không có thể tích và hình dạng xác định, lấp đầy toàn bộ không gian của vật chứa nó."
    },
    {
        "id": 6,
        "content": "Tính chất nào sau đây là tính chất hoá học của carbon dioxide?",
        "options": [
            "A. Chất khí, không màu.",
            "B. Không mùi, không vị.",
            "C. Tan rất ít trong nước.",
            "D. Làm đục nước vôi trong."
        ],
        "correct": "D",
        "solution": "Khí carbon dioxide tác dụng với dung dịch nước vôi trong tạo ra chất kết tủa trắng là calcium carbonate, làm đục nước vôi trong. Đây là phản ứng tạo chất mới, thể hiện tính chất hoá học. Các đặc điểm A, B, C là tính chất vật lí."
    },
    {
        "id": 7,
        "content": "Dãy nào sau đây đều gồm các vật sống?",
        "options": [
            "A. Cây mía, con ếch, xe đạp.",
            "B. Xe đạp, ấm đun nước, cái bút.",
            "C. Cây tre, con cá, con mèo.",
            "D. Máy vi tính, cái cặp, tivi."
        ],
        "correct": "C",
        "solution": "Cây tre, con cá, con mèo đều là sinh vật sống, có khả năng trao đổi chất, lớn lên, sinh sản và cảm ứng."
    },
    {
        "id": 8,
        "content": "Quá trình nào sau đây thể hiện tính chất hóa học?",
        "options": [
            "A. Hòa tan muối vào nước.",
            "B. Đun nóng bát đựng muối đến khi có tiếng nổ lách tách.",
            "C. Đun nóng đường ở thể rắn để chuyển sang đường ở thể lỏng.",
            "D. Đun nóng đường đến khi xuất hiện chất màu đen."
        ],
        "correct": "D",
        "solution": "Khi đun nóng đường đến nhiệt độ cao, đường bị phân hủy hóa học tạo thành than (chất rắn màu đen, carbon) và hơi nước, sinh ra chất mới nên thể hiện tính chất hóa học."
    },
    {
        "id": 9,
        "content": "Những tính chất nào sau đây là tính chất vật lí của chất?",
        "options": [
            "A. Khả năng tan trong nước, màu sắc, khả năng bị cháy.",
            "B. Tính dẫn điện, nhiệt độ sôi, khả năng tác dụng với nước.",
            "C. Khả năng tan trong nước, nhiệt độ sôi, màu sắc.",
            "D. Tính dẫn điện, khả năng tác dụng với nước, khả năng bị cháy."
        ],
        "correct": "C",
        "solution": "Màu sắc, khả năng tan, nhiệt độ sôi/nóng chảy là tính chất vật lí (không có sự tạo thành chất mới). Khả năng bị cháy và tác dụng với nước là tính chất hóa học."
    },
    {
        "id": 10,
        "content": "Đặc điểm nào sau đây không có ở thể rắn?",
        "options": [
            "A. Các hạt liên kết chặt chẽ.",
            "B. Có hình dạng và thể tích xác định.",
            "C. Rất khó bị nén.",
            "D. Có hình dạng và thể tích không xác định."
        ],
        "correct": "D",
        "solution": "Chất rắn luôn có hình dạng và thể tích xác định. Do đó, \"có hình dạng và thể tích không xác định\" là đặc điểm sai đối với thể rắn."
    },
    {
        "id": 11,
        "content": "Trường hợp nào sau đây thể hiện tính chất hóa học?",
        "options": [
            "A. Thanh sắt bị dát mỏng.",
            "B. Cồn để trong lọ không kín bị bay hơi.",
            "C. Thủy tinh nóng chảy được thổi thành bình cầu.",
            "D. Đốt cháy mẩu giấy."
        ],
        "correct": "D",
        "solution": "Đốt cháy mẩu giấy tạo ra tro than, khí CO~2~ và hơi nước, đây là quá trình biến đổi hóa học tạo chất mới."
    },
    {
        "id": 12,
        "content": "Quá trình chuyển từ thể rắn sang thể lỏng của chất gọi là",
        "options": ["A. Sự nóng chảy.", "B. Sự đông đặc.", "C. Sự bay hơi.", "D. Sự ngưng tụ."],
        "correct": "A",
        "solution": "Sự chuyển từ thể rắn sang thể lỏng gọi là sự nóng chảy."
    },
    {
        "id": 13,
        "content": "Quá trình chuyển từ thể lỏng sang thể hơi của chất gọi là",
        "options": ["A. Sự nóng chảy.", "B. Sự đông đặc.", "C. Sự bay hơi.", "D. Sự ngưng tụ."],
        "correct": "C",
        "solution": "Sự chuyển từ thể lỏng sang thể hơi (thể khí) gọi là sự bay hơi (hoặc sự hóa hơi)."
    },
    {
        "id": 14,
        "content": "Sự ngưng tụ là",
        "options": [
            "A. Quá trình chuyển từ thể rắn sang thể lỏng của chất.",
            "B. Quá trình chuyển từ thể lỏng sang thể hơi của chất.",
            "C. Quá trình chuyển từ thể rắn sang thể hơi của chất.",
            "D. Quá trình chuyển từ thể khí sang thể lỏng của chất."
        ],
        "correct": "D",
        "solution": "Sự ngưng tụ là quá trình chất chuyển từ thể khí (hơi) sang thể lỏng."
    },
    {
        "id": 15,
        "content": "Chất nào sau đây có thể tan trong nước?",
        "options": ["A. Dầu ăn.", "B. Đường.", "C. Sắt.", "D. Gỗ."],
        "correct": "B",
        "solution": "Đường tan rất tốt trong nước tạo thành dung dịch nước đường đồng nhất. Dầu ăn không tan và nổi trên mặt nước; sắt và gỗ không tan trong nước."
    },
    {
        "id": 16,
        "content": "Khi đun sôi nước, trạng thái của nước thay đổi như thế nào?",
        "options": ["A. Rắn sang lỏng.", "B. Lỏng sang khí.", "C. Khí sang lỏng.", "D. Lỏng sang rắn."],
        "correct": "B",
        "solution": "Khi đun sôi, nước ở thể lỏng chuyển thành hơi nước ở thể khí."
    },
    {
        "id": 17,
        "content": "Khi bơ tan chảy, sự thay đổi trạng thái từ gì sang gì?",
        "options": ["A. Rắn sang khí.", "B. Lỏng sang khí.", "C. Rắn sang lỏng.", "D. Lỏng sang rắn."],
        "correct": "C",
        "solution": "Miếng bơ lúc đầu ở thể rắn, khi gặp nhiệt độ ấm áp sẽ nóng chảy chuyển thành bơ lỏng (thể rắn sang thể lỏng)."
    },
    {
        "id": 18,
        "content": "Chất nào sau đây ở thể rắn ở điều kiện thường?",
        "options": ["A. Sắt.", "B. Nước.", "C. Oxygen.", "D. Thủy ngân."],
        "correct": "A",
        "solution": "Ở điều kiện nhiệt độ phòng: Sắt ở thể rắn; Nước và Thủy ngân ở thể lỏng; Oxygen ở thể khí."
    },
    {
        "id": 19,
        "content": "Chất nào dưới đây ở thể khí ở điều kiện thường?",
        "options": ["A. Muối ăn.", "B. Giấm.", "C. Carbon dioxide.", "D. Dầu ăn."],
        "correct": "C",
        "solution": "Carbon dioxide (CO~2~) là chất khí không màu. Muối ăn là chất rắn; giấm và dầu ăn là chất lỏng."
    },
    {
        "id": 20,
        "content": "Sự sôi là :",
        "options": [
            "A. Sự chuyển từ thể lỏng sang thể khí diễn ra trong lòng và trên bề mặt chất lỏng.",
            "B. Sự chuyển từ thể lỏng sang thể khí.",
            "C. Sự chuyển từ thể khí sang thể lỏng.",
            "D. Sự chuyển từ thể lỏng sang thể rắn."
        ],
        "correct": "A",
        "solution": "Sự sôi là trường hợp đặc biệt của sự bay hơi, diễn ra đồng thời cả ở bên trong lòng lẫn trên bề mặt thoáng của chất lỏng tại nhiệt độ sôi xác định."
    },
    {
        "id": 21,
        "content": "Sự chuyển thể nào sau đây xảy ra tại nhiệt độ xác định?",
        "options": ["A. Ngưng tụ.", "B. Sự bay hơi và sự sôi.", "C. Sự sôi.", "D. Bay hơi."],
        "correct": "C",
        "solution": "Mỗi chất lỏng nguyên chất sôi ở một nhiệt độ xác định gọi là nhiệt độ sôi. Hiện tượng bay hơi diễn ra ở mọi nhiệt độ."
    },
    {
        "id": 22,
        "content": "Sự chuyển thể nào sau đây không xảy ra tại nhiệt độ xác định?",
        "options": ["A. Nóng chảy.", "B. Hoá hơi.", "C. Sự sôi.", "D. Bay hơi."],
        "correct": "D",
        "solution": "Sự bay hơi của chất lỏng xảy ra ở mọi nhiệt độ trên bề mặt thoáng của chất lỏng."
    },
    {
        "id": 23,
        "content": "Khẳng định nào là sai khi nói về thể khí?",
        "options": [
            "A. Ở thể khí, chất khí có khối lượng xác định nhưng không có hình dạng và thể tích xác định.",
            "B. Ở thể khí, chất khí có thể lan tỏa theo mọi hướng và chiếm toàn bộ thể tích bất kì vật nào chứa nó.",
            "C. Ở thể khí, các “hạt” được sắp xếp chặt chẽ, khiến chúng không thể chuyển động tự do.",
            "D. Ở thể khí, các “hạt” ở cách xa nhau và có thể chuyển động tự do."
        ],
        "correct": "C",
        "solution": "Ở thể khí, khoảng cách giữa các phân tử/nguyên tử rất xa nhau và chuyển động hỗn loạn không ngừng về mọi phía. Khẳng định C là sai (đó là tính chất của thể rắn)."
    },
    {
        "id": 24,
        "content": "Chất ở thể nào thì có thể rót được và chảy tràn trên bề mặt?",
        "options": ["A. Thể dẻo.", "B. Thể rắn.", "C. Thể khí.", "D. Thể lỏng."],
        "correct": "D",
        "solution": "Chất ở thể lỏng có tính linh động cao, có thể rót được từ bình này sang bình khác và chảy tràn trên bề mặt nằm ngang."
    }
]

# -----------------------------------------------------------------------------
# PHÂN MÔN HÓA HỌC - PHẦN 2: TRẮC NGHIỆM ĐÚNG SAI (6 CÂU)
# -----------------------------------------------------------------------------
HOA_TF = [
    {
        "id": 1,
        "content": "Thế giới xung quanh chúng ta có nhiều vật thể khác nhau. Các vật thể có thể được phân loại theo nguồn gốc và đặc điểm sống.",
        "sub_items": [
            {"id": "a", "text": "Cây cối là vật thể tự nhiên.", "answer": True, "reason": "Cây cối sinh trưởng tự nhiên trong môi trường, không phải do con người chế tạo ra."},
            {"id": "b", "text": "Bàn học là vật thể tự nhiên vì được làm từ gỗ.", "answer": False, "reason": "Bàn học do con người gia công, chế tạo từ gỗ nên là vật thể nhân tạo."},
            {"id": "c", "text": "Con người là vật thể sống.", "answer": True, "reason": "Con người có đầy đủ các dấu hiệu đặc trưng của sự sống."},
            {"id": "d", "text": "Hòn đá là vật thể sống vì nó tồn tại trong tự nhiên.", "answer": False, "reason": "Hòn đá là vật thể tự nhiên nhưng là vật không sống vì không trao đổi chất, không sinh trưởng, không sinh sản."}
        ]
    },
    {
        "id": 2,
        "content": "Dựa vào các đặc điểm của vật sống và vật không sống, hãy xác định các phát biểu sau đúng hay sai.",
        "sub_items": [
            {"id": "a", "text": "Con mèo là vật sống.", "answer": True, "reason": "Con mèo là động vật, có trao đổi chất, hô hấp, sinh sản và lớn lên."},
            {"id": "b", "text": "Quyển sách là vật sống vì được làm từ giấy có nguồn gốc từ cây.", "answer": False, "reason": "Quyển sách là vật thể nhân tạo không có các quá trình sống cơ bản."},
            {"id": "c", "text": "Vật sống có các đặc trưng như trao đổi chất với môi trường, lớn lên và sinh sản.", "answer": True, "reason": "Đây là các dấu hiệu đặc trưng định nghĩa nên vật sống."},
            {"id": "d", "text": "Hòn đá có thể tự lớn lên và sinh sản như các sinh vật.", "answer": False, "reason": "Hòn đá là vật không sống, không thể tự lớn lên hay sinh sản."}
        ]
    },
    {
        "id": 3,
        "content": "Các vật thể xung quanh chúng ta được tạo nên từ một hoặc nhiều chất khác nhau.",
        "sub_items": [
            {"id": "a", "text": "Các vật thể đều được tạo nên từ chất.", "answer": True, "reason": "Chất là thành phần cấu tạo nên mọi vật thể."},
            {"id": "b", "text": "Một chất chỉ có thể tạo nên một loại vật thể duy nhất.", "answer": False, "reason": "Một chất có thể tạo nên vô số vật thể khác nhau (ví dụ thủy tinh làm cốc, bình hoa, cửa kính, bóng đèn...)."},
            {"id": "c", "text": "Một vật thể có thể được tạo nên từ nhiều chất.", "answer": True, "reason": "Ví dụ chiếc xe đạp được làm từ sắt, nhôm, cao su, nhựa..."},
            {"id": "d", "text": "Chiếc cốc thủy tinh là vật thể, thủy tinh là chất dùng để chế tạo chiếc cốc.", "answer": True, "reason": "Phân biệt đúng vật thể và chất tạo nên vật thể."}
        ]
    },
    {
        "id": 4,
        "content": "Khi tìm hiểu một chất, người ta có thể khảo sát các tính chất vật lí và tính chất hóa học của chất đó. Hãy xác định các phát biểu sau đúng hay sai.",
        "sub_items": [
            {"id": "a", "text": "Màu sắc và trạng thái là những tính chất vật lí của chất.", "answer": True, "reason": "Màu sắc, thể (rắn/lỏng/khí), mùi vị, tính tan... là tính chất vật lí."},
            {"id": "b", "text": "Sự biến đổi tạo ra chất mới là biểu hiện của sự biến đổi hóa học.", "answer": True, "reason": "Dấu hiệu cốt lõi của biến đổi hóa học là sự tạo thành chất mới."},
            {"id": "c", "text": "Khi đường tan trong nước, đường đã biến đổi thành một chất mới.", "answer": False, "reason": "Hòa tan đường chỉ là hiện tượng vật lí, các phân tử đường phân tán đều vào nước mà không bị biến đổi thành chất khác."},
            {"id": "d", "text": "Tính chất hóa học cho biết khả năng chất biến đổi thành chất khác.", "answer": True, "reason": "Tính chất hóa học mô tả khả năng biến đổi của một chất thành chất mới (cháy, phân hủy, tác dụng với chất khác...)."}
        ]
    },
    {
        "id": 5,
        "content": "Chất có thể tồn tại ở thể rắn, thể lỏng và thể khí. Hãy xác định các phát biểu sau đúng hay sai.",
        "sub_items": [
            {"id": "a", "text": "Sắt, đá và gỗ là những ví dụ về chất ở thể rắn.", "answer": True, "reason": "Ở điều kiện thường, chúng tồn tại ở thể rắn."},
            {"id": "b", "text": "Chất ở thể lỏng có hình dạng cố định.", "answer": False, "reason": "Chất lỏng không có hình dạng cố định, nó mang hình dạng của phần bình chứa nó."},
            {"id": "c", "text": "Chất ở thể khí không có hình dạng cố định và có thể lan tỏa để chiếm đầy bình chứa.", "answer": True, "reason": "Đặc trưng của chất khí là chuyển động tự do và khuếch tán chiếm toàn bộ thể tích vật chứa."},
            {"id": "d", "text": "Chất ở thể rắn thường có hình dạng xác định.", "answer": True, "reason": "Chất rắn có hình dạng và thể tích xác định."}
        ]
    },
    {
        "id": 6,
        "content": "Nước đá được lấy ra khỏi tủ lạnh và đặt trong phòng. Sau một thời gian, nước đá tan thành nước lỏng. Hãy xác định các phát biểu sau đúng hay sai.",
        "sub_items": [
            {"id": "a", "text": "Nước đá chuyển từ thể rắn sang thể lỏng, gọi là sự nóng chảy.", "answer": True, "reason": "Sự chuyển từ rắn sang lỏng là sự nóng chảy."},
            {"id": "b", "text": "Sự đông đặc là quá trình chuyển từ thể rắn sang thể lỏng.", "answer": False, "reason": "Sự đông đặc là quá trình ngược lại: chuyển từ thể lỏng sang thể rắn."},
            {"id": "c", "text": "Khi nước lỏng chuyển thành nước đá, nước xảy ra sự đông đặc.", "answer": True, "reason": "Nước lỏng hạ nhiệt độ xuống 0°C đóng băng thành đá là sự đông đặc."},
            {"id": "d", "text": "Sự nóng chảy và sự đông đặc là hai quá trình chuyển thể ngược nhau.", "answer": True, "reason": "Nóng chảy (rắn -> lỏng) và đông đặc (lỏng -> rắn) là hai quá trình thuận nghịch."}
        ]
    }
]

# -----------------------------------------------------------------------------
# PHÂN MÔN HÓA HỌC - PHẦN 3: CÂU HỎI TRẢ LỜI NGẮN (4 CÂU)
# -----------------------------------------------------------------------------
HOA_SHORT = [
    {
        "id": 1,
        "content": "Có các vật thể sau: quả chanh, máy tính, cây mít, cái chậu, lọ hoa, xe máy, cây tre. Số vật thể tự nhiên là bao nhiêu?",
        "answer": "3",
        "solution": "Các vật thể tự nhiên là: quả chanh, cây mít, cây tre (tổng cộng 3 vật thể).\nCác vật thể còn lại (máy tính, cái chậu, lọ hoa, xe máy) là vật thể nhân tạo."
    },
    {
        "id": 2,
        "content": "Có các vật thể sau: xe máy, xe đạp, con người, con suối, con trâu, bóng đèn, thước kẻ. Số vật thể nhân tạo là bao nhiêu?",
        "answer": "4",
        "solution": "Các vật thể nhân tạo do con người chế tạo ra gồm: xe máy, xe đạp, bóng đèn, thước kẻ (tổng cộng 4 vật thể).\nCon người, con suối, con trâu là vật thể tự nhiên."
    },
    {
        "id": 3,
        "content": "Có bao nhiêu thể cơ bản của chất thường được tìm hiểu ở cấp THCS?",
        "answer": "3",
        "solution": "Có 3 thể cơ bản của chất thường được tìm hiểu là: thể rắn, thể lỏng và thể khí (hơi)."
    },
    {
        "id": 4,
        "content": "Trong các quá trình sau: nóng chảy, đông đặc, bay hơi, ngưng tụ, có bao nhiêu quá trình chuyển chất từ thể lỏng sang thể khác?",
        "answer": "2",
        "solution": "Các quá trình bắt đầu từ thể lỏng chuyển sang thể khác gồm 2 quá trình:\n1. Đông đặc (chuyển từ thể lỏng sang thể rắn).\n2. Bay hơi (chuyển từ thể lỏng sang thể khí/hơi).\n(Nóng chảy là từ thể rắn sang thể lỏng; Ngưng tụ là từ thể khí sang thể lỏng)."
    }
]

# -----------------------------------------------------------------------------
# PHÂN MÔN HÓA HỌC - III. TỰ LUẬN (4 CÂU)
# -----------------------------------------------------------------------------
HOA_ESSAY = [
    {
        "id": 1,
        "content": "Điền các từ cho sẵn dưới đây vào các chỗ trống để hoàn thành thông tin: *chất*, *vật thể*, *đa dạng*, *nhân tạo*:\na. Vật thể quanh ta vô cùng .................; vật thể tự nhiên, vật thể .........., vật sống, vật không sống.\nb. Vật thể cấu tạo từ......................\nc. Ở đâu có .................. là ở đó có chất.",
        "solution": "a. Vật thể quanh ta vô cùng *đa dạng*; vật thể tự nhiên, vật thể *nhân tạo*, vật sống, vật không sống.\nb. Vật thể cấu tạo từ *chất*.\nc. Ở đâu có *vật thể* là ở đó có chất."
    },
    {
        "id": 2,
        "content": "Hãy chỉ ra đâu là vật thể, đâu là chất trong những từ (in nghiêng) sau:\na. *Cơ thể người* có 63-68% khối lượng là *nước*.\nb. *Than chì* là chất làm *lõi bút chì*.\nc. *Dây điện* làm bằng *copper* được bọc một lớp *chất dẻo*.\nd. *Áo* may bằng sợi bông (95-98% là *cellulose*) mặc thoáng mát hơn bằng *nilon* (một thứ tơ tổng hợp).\ne. *Xe đạp* được chế tạo từ: *iron*, *aluminium*, *cao su*, ...",
        "solution": "a. Vật thể: *Cơ thể người*. Chất: *nước*.\nb. Vật thể: *lõi bút chì*. Chất: *Than chì*.\nc. Vật thể: *Dây điện*. Chất: *copper* (đồng), *chất dẻo*.\nd. Vật thể: *Áo*. Chất: *cellulose*, *nilon*.\ne. Vật thể: *Xe đạp*. Chất: *iron* (sắt), *aluminium* (nhôm), *cao su*."
    },
    {
        "id": 3,
        "content": "Chọn những cụm từ thích hợp để điền vào các chỗ trống trong các câu sau:\na) Sự nóng chảy là sự chuyển từ .................... sang......................\nb) Sự đông đặc là sự chuyển từ.................... sang ....................",
        "solution": "a) Sự nóng chảy là sự chuyển từ *thể rắn* sang *thể lỏng*.\nb) Sự đông đặc là sự chuyển từ *thể lỏng* sang *thể rắn*."
    },
    {
        "id": 4,
        "content": "Hình dưới đây được chụp tại một con đường ở Ấn Độ vào mùa hè với nhiệt độ ngoài trời có lúc lên trên 50°C:\na) Theo em, hiện tượng nhựa đường như trên có thể gọi là hiện tượng gì?\nb) Qua hiện tượng trên, em có kết luận gì về nhiệt độ nóng chảy của nhựa đường?",
        "image": "khtn6_asphalt_p10.png",
        "solution": "a) Hiện tượng nhựa đường bị mềm nhũn, biến dạng và chảy tràn dính chặt vào bánh xe và dép người đi đường gọi là hiện tượng nóng chảy (sự nóng chảy của chất rắn vô định hình chuyển sang thể quánh dẻo/lỏng dưới tác dụng của nhiệt độ cao).\nb) Kết luận: Nhựa đường có nhiệt độ nóng chảy tương đối thấp (khoảng từ 45°C đến 55°C). Khi thời tiết nắng nóng gay gắt trên 50°C, mặt đường hấp thụ nhiệt lượng lớn từ ánh sáng mặt trời vượt qua nhiệt độ nóng chảy làm nhựa đường chuyển thể nóng chảy."
    }
]
