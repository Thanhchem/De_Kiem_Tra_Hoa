# -*- coding: utf-8 -*-
"""
Dữ liệu Đề cương Ôn tập Kiểm tra Giữa kì I - KHTN 6 (Năm học 2026 - 2027)
Gồm 3 Phân môn:
1. Phân môn Vật lý (Lý thuyết, Trắc nghiệm nhiều lựa chọn, Trả lời ngắn, Đúng/Sai, Tự luận)
2. Phân môn Hóa học (Lý thuyết cơ bản và các hiện tượng, tính chất)
3. Phân môn Sinh học (Chương I: Mở đầu, Chương V: Tế bào, Chương VI: Từ tế bào đến cơ thể)
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

INFO = {
    "school": "TRƯỜNG TRUNG HỌC CƠ SỞ",
    "year": "NĂM HỌC 2026 - 2027",
    "title": "ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA KÌ I",
    "subject": "MÔN: KHOA HỌC TỰ NHIÊN 6",
    "subtitle": "(Gồm 3 phân môn: Vật lý - Hoá học - Sinh học)",
    "time_str": "Thời gian làm bài: 60 phút (Không kể thời gian phát đề)"
}

# ==============================================================================
# PHẦN 1: PHÂN MÔN VẬT LÝ
# ==============================================================================

VATLI_THEORY = [
    {
        "title": "Câu 1: Đo chiều dài",
        "items": [
            "a. Đơn vị độ dài:",
            "- Đơn vị đo độ dài hợp pháp của nước Việt Nam là mét (ký hiệu: m).",
            "- Ngoài ra còn có một số đơn vị khác: km, dm, cm, mm.",
            "Chú ý: 1 km = 1000 m; 1 m = 10 dm = 100 cm = 1000 mm.",
            "b. Dụng cụ đo độ dài:",
            "- Dụng cụ đo độ dài là thước.",
            "- Có các loại thước: thước thẳng, thước dây, thước cuộn, thước kẹp...",
            "c. Giới hạn đo và ĐCNN:",
            "- Giới hạn đo (GHĐ) là độ dài lớn nhất ghi trên thước.",
            "- Độ chia nhỏ nhất (ĐCNN) là độ dài giữa 2 vạch chia liên tiếp trên thước.",
            "d. Cách đo chiều dài:",
            "- Bước 1 (Ước lượng): Ước lượng chiều dài của vật cần đo để chọn thước có GHĐ và ĐCNN phù hợp.",
            "- Bước 2 (Chọn thước): Chọn thước có Giới hạn đo (GHĐ) lớn hơn hoặc bằng chiều dài cần đo và Độ chia nhỏ nhất (ĐCNN) thích hợp với độ chính xác cần thiết.",
            "- Bước 3 (Đặt thước): Đặt thước dọc theo độ dài cần đo, sao cho một đầu của vật ngang bằng với vạch số 0 của thước.",
            "- Bước 4 (Đặt mắt): Đặt mắt nhìn vuông góc với cạnh của thước ở đầu kia của vật (tránh lỗi cận thị hay nhìn xiên làm đọc sai kết quả).",
            "- Bước 5 (Đọc và ghi kết quả): Đọc giá trị của độ dài theo vạch chia gần nhất với đầu kia của vật và ghi kết quả kèm theo đơn vị đo (theo ĐCNN của thước)."
        ]
    },
    {
        "title": "Câu 2: Đo khối lượng",
        "items": [
            "a. Khái niệm khối lượng:",
            "- Khối lượng là số đo lượng chất của một vật.",
            "b. Đơn vị đo khối lượng:",
            "- Trong hệ thống đo lường hợp pháp của Việt Nam, đơn vị đo khối lượng là kilôgam (kí hiệu: kg).",
            "- Ngoài ra còn có một số đơn vị khác: tấn, tạ, yến, hg, dag, g, mg.",
            "Chú ý: 1 tấn = 1000 kg; 1 tạ = 100 kg; 1 yến = 10 kg; 1 lạng = 1 hg = 100 g; 1 g = 1000 mg; 1 kg = 1000 g.",
            "c. Dụng cụ đo khối lượng:",
            "- Dụng cụ đo khối lượng là cân.",
            "- Có các loại cân thông dụng: cân đồng hồ, cân điện tử, cân y tế, cân đòn, cân tấn, cân tạ, cân Rô-béc-van (Roberval)..."
        ]
    },
    {
        "title": "Câu 3: Đo thời gian",
        "items": [
            "a. Đơn vị đo thời gian:",
            "- Trong hệ thống đo lường hợp pháp của Việt Nam, đơn vị cơ bản đo thời gian là giây (kí hiệu: s).",
            "- Ngoài ra còn có một số đơn vị khác: giờ (h), phút (min), ngày, tháng, năm, mili giây (ms)...",
            "Chú ý: 1 ngày = 24 giờ; 1 giờ = 60 phút = 3600 giây; 1 giây = 1000 mili giây (ms).",
            "b. Dụng cụ đo thời gian:",
            "- Dụng cụ đo thời gian là đồng hồ.",
            "- Có các loại đồng hồ: đồng hồ bấm giây, đồng hồ đeo tay, đồng hồ cát, đồng hồ quả lắc, đồng hồ điện tử..."
        ]
    }
]

VATLI_MCQ = [
    {
        "id": 1,
        "content": "Đơn vị đo độ dài hợp pháp của nước Việt Nam theo hệ đo lường quốc tế (SI) là gì?",
        "options": ["A. Centimét (cm).", "B. Kilômét (km).", "C. Mét (m).", "D. Milimét (mm)."],
        "answer": "C",
        "solution": "Theo hệ đo lường quốc tế (SI) và quy chuẩn đo lường hợp pháp tại Việt Nam, đơn vị chuẩn để đo độ dài là mét (kí hiệu: m)."
    },
    {
        "id": 2,
        "content": "GHĐ của thước là gì?",
        "options": [
            "A. Là độ dài nhỏ nhất ghi trên thước.",
            "B. Là độ dài lớn nhất ghi trên thước.",
            "C. Là khoảng cách giữa hai vạch chia liên tiếp trên thước.",
            "D. Là chiều dài thực tế của cái thước."
        ],
        "answer": "B",
        "solution": "Giới hạn đo (GHĐ) của thước là độ dài lớn nhất ghi trên thước."
    },
    {
        "id": 3,
        "content": "Dụng cụ nào sau đây không phải là dụng cụ dùng để đo chiều dài?",
        "options": ["A. Thước thẳng.", "B. Thước cuộn.", "C. Cân đồng hồ.", "D. Thước kẹp."],
        "answer": "C",
        "solution": "Cân đồng hồ là dụng cụ dùng để đo khối lượng. Thước thẳng, thước cuộn và thước kẹp đều là các dụng cụ đo chiều dài."
    },
    {
        "id": 4,
        "content": "Khi đo chiều dài của một vật bằng thước, bước đầu tiên cần thực hiện là gì?",
        "options": [
            "A. Đặt mắt vuông góc với vạch chia trên thước.",
            "B. Ước lượng chiều dài của vật cần đo.",
            "C. Đặt thước dọc theo chiều dài cần đo.",
            "D. Đọc kết quả đo theo vạch chia gần nhất."
        ],
        "answer": "B",
        "solution": "Bước đầu tiên trong quy trình đo là ước lượng chiều dài của vật cần đo để chọn thước có GHĐ và ĐCNN phù hợp."
    },
    {
        "id": 5,
        "content": "ĐCNN (độ chia nhỏ nhất) của thước là gì?",
        "options": [
            "A. Độ dài lớn nhất ghi trên thước.",
            "B. Độ dài giữa hai vạch chia liên tiếp trên thước.",
            "C. Vạch số 0 của thước.",
            "D. Chiều dài của vật cần đo."
        ],
        "answer": "B",
        "solution": "Độ chia nhỏ nhất (ĐCNN) của thước là độ dài giữa hai vạch chia liên tiếp trên thước."
    },
    {
        "id": 6,
        "content": "Đơn vị đo khối lượng hợp pháp của Việt Nam theo hệ SI là gì?",
        "options": ["A. Gam (g).", "B. Kilôgam (kg).", "C. Miligam (mg).", "D. Tấn."],
        "answer": "B",
        "solution": "Đơn vị đo khối lượng chính thức, hợp pháp của nước ta theo hệ đo lường quốc tế (SI) là kilôgam (kí hiệu: kg)."
    },
    {
        "id": 7,
        "content": "Trong các dụng cụ sau đây, dụng cụ nào dùng để đo khối lượng?",
        "options": ["A. Thước cuộn.", "B. Bình chia độ.", "C. Cân đồng hồ.", "D. Nhiệt kế."],
        "answer": "C",
        "solution": "Cân đồng hồ dùng để đo khối lượng. Thước cuộn đo chiều dài, bình chia độ đo thể tích chất lỏng, nhiệt kế đo nhiệt độ."
    },
    {
        "id": 8,
        "content": "Khi dùng cân đồng hồ để đo khối lượng của một vật, bước đầu tiên cần làm là gì?",
        "options": [
            "A. Đặt vật lên đĩa cân.",
            "B. Ước lượng khối lượng của vật cần cân.",
            "C. Đọc và ghi kết quả đo.",
            "D. Chỉnh kim cân đúng vạch số 0."
        ],
        "answer": "B",
        "solution": "Bước đầu tiên trong quy trình đo là ước lượng khối lượng của vật cần cân để lựa chọn cân có GHĐ và ĐCNN phù hợp."
    },
    {
        "id": 9,
        "content": "Đơn vị đo thời gian trong hệ thống đo lường chính thức (SI) của nước ta hiện nay là gì?",
        "options": ["A. Giờ.", "B. Phút.", "C. Giây (s).", "D. Ngày."],
        "answer": "C",
        "solution": "Trong hệ SI và hệ thống đo lường hợp pháp của nước ta, đơn vị cơ bản đo thời gian là giây (kí hiệu: s)."
    },
    {
        "id": 10,
        "content": "Dụng cụ thông dụng nhất dùng để đo thời gian là:",
        "options": ["A. Thước thẳng.", "B. Cân.", "C. Đồng hồ.", "D. Nhiệt kế."],
        "answer": "C",
        "solution": "Đồng hồ là dụng cụ thông dụng và phổ biến nhất dùng để đo khoảng thời gian và thời điểm."
    },
    {
        "id": 11,
        "content": "Trong các đồng hồ sau đây, loại đồng hồ nào thường dùng để đo thời gian trong các cuộc thi đấu thể thao (như chạy 100m, 200m)?",
        "options": [
            "A. Đồng hồ cát.",
            "B. Đồng hồ bấm giây.",
            "C. Đồng hồ treo tường.",
            "D. Đồng hồ để bàn."
        ],
        "answer": "B",
        "solution": "Đồng hồ bấm giây có độ chính xác cao và khả năng bấm bắt đầu/dừng tức thì, rất phù hợp để đo thời gian các nội dung thi đấu thể thao."
    }
]

VATLI_SHORT = [
    {
        "id": 1,
        "content": "Một thước thẳng có vạch chia bắt đầu từ số 0 cm và kết thúc ở số 30 cm. Giữa hai vạch số 0 và số 1 cm được chia thành 5 khoảng bằng nhau.",
        "sub_questions": [
            {
                "sub_id": "a",
                "text": "Độ chia nhỏ nhất (ĐCNN) của thước này là bao nhiêu mi-li-mét?",
                "answer": "2 mm",
                "solution": "Khoảng cách giữa 0 cm và 1 cm là 1 cm = 10 mm. Khoảng cách này được chia thành 5 khoảng đều nhau, do đó: ĐCNN = 10 mm / 5 = 2 mm (tương ứng 0,2 cm)."
            },
            {
                "sub_id": "b",
                "text": "Giới hạn đo (GHĐ) của cây thước này là bao nhiêu mét?",
                "answer": "0,3 m",
                "solution": "Vạch lớn nhất ghi trên thước là 30 cm. Đổi sang mét: GHĐ = 30 cm = 30 / 100 m = 0,3 m."
            },
            {
                "sub_id": "c",
                "text": "Một bạn học sinh dùng cây thước này để đo chiều dài của cây bút bi và ghi kết quả đo như sau: 22,50 cm; 22,2 cm; 22,25 cm. Kết quả nào ghi đúng quy tắc đo?",
                "answer": "22,2 cm",
                "solution": "Vì thước có ĐCNN là 0,2 cm (2 mm) nên phần chữ số thập phân phải là bội số của 0,2 (tức là 0,2; 0,4; 0,6; 0,8; 0,0) và chỉ ghi đến 1 chữ số thập phân. Do đó, kết quả ghi đúng quy tắc là 22,2 cm (kết quả 22,50 cm và 22,25 cm có hàng phần trăm không phù hợp với ĐCNN)."
            },
            {
                "sub_id": "d",
                "text": "Khi đo chiều dài của quyển sách KHTN 6 một đầu của quyển sách đặt trùng với vạch số 0, đầu kia của quyển sách nằm chính xác tại vạch số 3 sau vạch 22 cm. Chiều dài của quyển sách là bao nhiêu?",
                "answer": "22,6 cm (hoặc 226 mm)",
                "solution": "Mỗi vạch chia ứng với ĐCNN = 0,2 cm. Vạch số 3 sau vạch 22 cm có giá trị là: L = 22 + 3 × 0,2 = 22,6 cm = 226 mm."
            }
        ]
    },
    {
        "id": 2,
        "content": "Một cân đồng hồ có vạch chia bắt đầu từ số 0 và trên cân có ghi 2 kg. Giữa hai vạch số 0 và số 100 g được chia thành 10 khoảng bằng nhau.",
        "sub_questions": [
            {
                "sub_id": "a",
                "text": "Độ chia nhỏ nhất (ĐCNN) của cân này là bao nhiêu kg?",
                "answer": "0,01 kg (hoặc 10 g)",
                "solution": "Từ vạch 0 đến 100 g chia làm 10 khoảng bằng nhau, vậy mỗi khoảng có giá trị là: ĐCNN = 100 g / 10 = 10 g = 0,01 kg."
            },
            {
                "sub_id": "b",
                "text": "Giới hạn đo (GHĐ) của cân này là bao nhiêu kg?",
                "answer": "2 kg",
                "solution": "Giá trị lớn nhất ghi trên mặt cân là 2 kg, do đó GHĐ của cân là 2 kg."
            },
            {
                "sub_id": "c",
                "text": "Một bạn học sinh dùng cân này để cân khối lượng của một túi cam và ghi kết quả đo như sau: 820 g; 855 g; 870 g. Kết quả nào ghi không đúng quy tắc đo?",
                "answer": "855 g",
                "solution": "Cân có ĐCNN là 10 g nên kết quả đo phải là bội số của 10 g (chữ số hàng đơn vị phải là 0). Kết quả 855 g có chữ số tận cùng là 5 (không chia hết cho 10) nên là kết quả ghi không đúng quy tắc."
            },
            {
                "sub_id": "d",
                "text": "Khi đo khối lượng của một gói bánh, một bạn học sinh quan sát được kết quả đo là kim cân nằm chính xác tại vạch số 7 sau vạch 1,3 kg. Khối lượng của túi bánh là bao nhiêu?",
                "answer": "1,37 kg (hoặc 1370 g)",
                "solution": "1,3 kg = 1300 g. Mỗi vạch chia tương ứng 10 g. Vạch thứ 7 sau vạch 1,3 kg chỉ giá trị: m = 1300 g + 7 × 10 g = 1370 g = 1,37 kg."
            }
        ]
    }
]

VATLI_TF = [
    {
        "id": 1,
        "context": "Một học sinh dùng một chiếc thước thẳng có Giới hạn đo (GHĐ) là 1 m và Độ chia nhỏ nhất (ĐCNN) là 1 mm để đo chiều dài của một chiếc bàn học. Khi thực hiện phép đo, học sinh đó đặt vạch số 0 của thước trùng với một đầu của bàn, đầu kia của bàn nằm ở vạch 115,5 cm.",
        "items": [
            {
                "id": "a",
                "content": "Cây thước học sinh này chọn là dụng cụ phù hợp nhất để đo chiều dài chiếc bàn.",
                "answer": False,
                "reason": "Chiếc bàn dài hơn 1 m (khoảng 115,5 cm) nên thước có GHĐ 1 m nhỏ hơn chiều dài cần đo, buộc phải đặt đo nhiều lần gây sai số lớn, không phải dụng cụ phù hợp nhất."
            },
            {
                "id": "b",
                "content": "Kết quả đo chiều dài của chiếc bàn được ghi nhận là 115,5 cm là đúng quy tắc và chính xác.",
                "answer": False,
                "reason": "Thước chỉ có GHĐ là 1 m (100 cm) nên trên thước không hề có vạch 115,5 cm để đọc trong một lần đo. Ngoài ra, ĐCNN là 1 mm (0,1 cm) nên nếu đo ghép cũng không thể đọc tùy tiện 115,5 cm."
            },
            {
                "id": "c",
                "content": "Để đo chiếc bàn này một cách chính xác và thuận tiện hơn, nên thay bằng một chiếc thước cuộn hoặc thước dây có GHĐ lớn hơn 115,5 cm.",
                "answer": True,
                "reason": "Thước cuộn hoặc thước dây có GHĐ lớn hơn 115,5 cm (như 2 m hoặc 3 m) sẽ đo được trọn vẹn chiều dài bàn trong 1 lần đo, giúp thao tác nhanh và giảm thiểu sai số."
            },
            {
                "id": "d",
                "content": "Độ chia nhỏ nhất của thước là 1 mm cho phép đọc kết quả đến độ chính xác đến mi-li-mét.",
                "answer": True,
                "reason": "ĐCNN của thước là 1 mm nên độ chính xác của phép đo đạt tới từng mi-li-mét (kết quả đọc được làm tròn đến vạch mi-li-mét gần nhất)."
            }
        ]
    },
    {
        "id": 2,
        "context": "Bạn An dùng một cân đồng hồ có Giới hạn đo (GHĐ) là 2 kg và Độ chia nhỏ nhất (ĐCNN) là 10 g để cân một hộp bánh có ghi khối lượng tịnh là 750 g. Khi đặt hộp bánh lên cân, kim cân chỉ vạch 760 g. An thực hiện một số nhận định sau:",
        "items": [
            {
                "id": "a",
                "content": "Chiếc cân An đang dùng hoàn toàn phù hợp để đo khối lượng của hộp bánh này.",
                "answer": True,
                "reason": "Hộp bánh nặng khoảng 750 g - 760 g nằm trong phạm vi đo từ 0 đến 2 kg của cân, và ĐCNN 10 g đủ chi tiết để đo khối lượng vật này."
            },
            {
                "id": "b",
                "content": "Khối lượng thực tế của hộp bánh đúng 750 g như số ghi trên vỏ hộp.",
                "answer": False,
                "reason": "750 g ghi trên vỏ hộp là khối lượng tịnh (khối lượng của riêng phần bánh bên trong, chưa tính bao bì). Khi cân cả hộp bánh (gồm cả vỏ và bánh) kim cân chỉ 760 g, chứng tỏ khối lượng thực tế cả hộp là 760 g (vỏ hộp nặng 10 g)."
            },
            {
                "id": "c",
                "content": "Sai số của phép đo trong trường hợp này lớn hơn ĐCNN của cân.",
                "answer": False,
                "reason": "Kim cân chỉ đúng vạch 760 g (là bội số đúng của ĐCNN = 10 g), sai số dụng cụ tối đa chỉ bằng nửa hoặc một ĐCNN (10 g), không lớn hơn ĐCNN."
            },
            {
                "id": "d",
                "content": "Nếu thay hộp bánh bằng một bao gạo nặng 2,5 kg thì chiếc cân đồng hồ trên vẫn đo tốt và bảo đảm độ bền của cân.",
                "answer": False,
                "reason": "Bao gạo 2,5 kg vượt quá Giới hạn đo của cân (2 kg), sẽ làm biến dạng lò xo cân quá mức, làm cân mất độ chính xác và nhanh hỏng."
            }
        ]
    }
]

VATLI_ESSAY = [
    {
        "id": 1,
        "content": "Trên bao bì gói bánh COSY có ghi: Khối lượng tịnh 500 g. Ở bảng giá trị dinh dưỡng phía sau, nhà sản xuất cho biết cứ 10 g bánh sẽ cung cấp khoảng 100 cal.\na) Con số “500 g” đó cho em biết điều gì?\nb) Bạn Lan đã ăn hết toàn bộ gói bánh này. Em hãy tính xem bạn Lan đã nạp vào cơ thể bao nhiêu năng lượng từ gói bánh đó?",
        "solution": "a) Con số “500 g” ghi trên bao bì cho biết khối lượng tịnh của riêng phần bánh chứa bên trong gói (không tính khối lượng của vỏ bao bì) là 500 gam.\nb) Tính năng lượng bạn Lan đã nạp vào cơ thể:\n- Cứ 10 g bánh cung cấp 100 cal.\n- Toàn bộ gói bánh nặng 500 g bánh sẽ cung cấp lượng năng lượng là:\n  E = (500 / 10) × 100 = 50 × 100 = 5000 cal (hoặc 5 kcal).\nĐáp số: 5000 cal (hay 5 kcal)."
    },
    {
        "id": 2,
        "content": "Trên một thùng bánh COSY chứa 20 gói bánh cùng loại, nhà sản xuất có ghi: \"Khối lượng tịnh của mỗi gói: 100 g\".\na) Con số 100 g ghi trên vỏ gói bánh cho biết điều gì?\nb) Cho biết khối lượng của thùng giấy là 200 g. Tổng khối lượng của cả bánh và thùng giấy là bao nhiêu kg?",
        "solution": "a) Con số 100 g ghi trên vỏ gói bánh cho biết khối lượng tịnh của riêng phần bánh bên trong mỗi gói (không tính khối lượng vỏ gói bánh) là 100 gam.\nb) Tính tổng khối lượng của cả bánh và thùng giấy:\n- Khối lượng của 20 gói bánh là:\n  m_bánh = 20 × 100 g = 2000 g = 2 kg.\n- Khối lượng của vỏ thùng giấy là:\n  m_thùng = 200 g = 0,2 kg.\n- Tổng khối lượng của cả bánh và thùng giấy là:\n  m_tổng = m_bánh + m_thùng = 2000 g + 200 g = 2200 g = 2,2 kg.\nĐáp số: 2,2 kg (hoặc 2200 g)."
    }
]

# ==============================================================================
# PHẦN 2: PHÂN MÔN HOÁ HỌC
# ==============================================================================

HOA_THEORY = [
    {
        "title": "Câu 1: Chất tạo nên vật thể. Ở đâu có vật thể là ở đó có chất.",
        "items": [
            "- Mỗi vật thể được tạo thành từ 1 hay nhiều chất khác nhau.",
            "  + Ví dụ: Giọt nước được tạo thành từ 1 chất là nước; cây kem được tạo thành từ nhiều chất như nước, đường, hương vị, sữa...",
            "- Tùy theo cách phân loại có thể chia thành:",
            "  + Vật sống (vật hữu sinh): Có các đặc trưng sống như trao đổi chất với môi trường, lớn lên và sinh sản. Ví dụ: Con sư tử, em bé, cây mai.",
            "  + Vật không sống (vật vô sinh): Không có khả năng trao đổi chất với môi trường, không lớn lên, không sinh sản. Ví dụ: Bánh mì, Cầu Long Biên, núi đá vôi.",
            "  + Vật thể tự nhiên: Có sẵn trong tự nhiên. Ví dụ: Núi đá vôi, mủ cao su, con sư tử, nước biển...",
            "  + Vật thể nhân tạo: Do con người tạo ra từ các vật liệu tự nhiên hoặc nhân tạo. Ví dụ: Cầu Long Biên, bánh mì, cái bàn, bút chì..."
        ]
    },
    {
        "title": "Câu 2: Mỗi chất có những tính chất vật lí và tính chất hóa học nhất định, đặc trưng cho chất.",
        "items": [
            "- Tính chất vật lí: Thể hiện ở thể (rắn, lỏng, khí), màu sắc, mùi, vị, nhiệt độ sôi, nhiệt độ nóng chảy, tính dẫn điện, tính dẫn nhiệt, tính tan trong nước hoặc dung môi khác... (không tạo thành chất mới).",
            "  + Ví dụ: Nước từ thể lỏng chuyển thành thể hơi (sự bay hơi); muối ăn tan trong nước.",
            "- Tính chất hóa học: Thể hiện sự biến đổi của chất: chất này bị biến đổi thành chất khác hoàn toàn (có phản ứng hóa học xảy ra).",
            "  + Ví dụ: Đường đun cháy bị biến đổi thành than (cacbon) và hơi nước; đinh sắt bị gỉ trong không khí ẩm."
        ]
    },
    {
        "title": "Câu 3: Đặc điểm (tính chất) cơ bản 3 thể (rắn, lỏng, khí) thông qua quan sát:",
        "items": [
            "- Thể rắn: Có hình dạng và thể tích cố định (xác định), không tự chảy được, rất khó bị nén.",
            "  + Ví dụ: Cái bút có hình dạng dài cố định, không tự di chuyển, không bị nén được.",
            "- Thể lỏng: Có thể tích xác định nhưng không có hình dạng xác định (có hình dạng của phần vật chứa nó), có thể rót được và chảy tràn trên bề mặt, khó bị nén.",
            "  + Ví dụ: Nước, dầu ăn, cồn...",
            "- Thể khí (hơi): Không có hình dạng và thể tích xác định (chiếm toàn bộ thể tích của vật chứa nó), dễ dàng lan tỏa trong không gian theo mọi hướng, rất dễ bị nén.",
            "  + Ví dụ: Mùi nước hoa, dầu gió có thể bay hơi và lan tỏa trong không khí theo mọi hướng."
        ]
    },
    {
        "title": "Câu 4: Các khái niệm về sự chuyển thể của chất:",
        "items": [
            "- Sự nóng chảy: Quá trình chất ở thể rắn chuyển sang thể lỏng. Quá trình này xảy ra ở một nhiệt độ xác định gọi là nhiệt độ nóng chảy.",
            "- Sự đông đặc: Quá trình chất ở thể lỏng chuyển sang thể rắn. Quá trình này xảy ra ở một nhiệt độ xác định gọi là nhiệt độ đông đặc.",
            "  * Chú ý: Mỗi chất nóng chảy và đông đặc ở cùng một nhiệt độ. Ví dụ: Nước đá nóng chảy và nước đông đặc đều ở cùng 0 °C.",
            "- Sự hóa hơi: Quá trình chất ở thể lỏng chuyển sang thể hơi (khí). Gồm 2 hình thức: sự bay hơi và sự sôi.",
            "  + Sự bay hơi: Sự hóa hơi chỉ xảy ra trên bề mặt chất lỏng. Sự bay hơi xảy ra ở mọi nhiệt độ.",
            "  + Sự sôi: Sự hóa hơi xảy ra đồng thời cả trên bề mặt và trong lòng khối chất lỏng. Sự sôi chỉ xảy ra ở một nhiệt độ xác định gọi là nhiệt độ sôi.",
            "    Ví dụ: Nước bay hơi ở mọi nhiệt độ, nhưng sôi ở 100 °C (ở áp suất chuẩn).",
            "- Sự ngưng tụ: Quá trình chất ở thể hơi (khí) chuyển sang thể lỏng. Sự ngưng tụ xảy ra ở mọi nhiệt độ."
        ]
    },
    {
        "title": "Câu 5: Tính chất và vai trò của Oxygen:",
        "items": [
            "- Tính chất vật lí của oxygen: Ở điều kiện thường, oxygen là chất khí không màu, không mùi, không vị, ít tan trong nước và nặng hơn không khí.",
            "- Oxygen hoá lỏng ở -183 °C, hoá rắn ở -218 °C. Ở thể lỏng và thể rắn, oxygen có màu xanh nhạt.",
            "- Vai trò của oxygen trong đời sống và sản xuất:",
            "  + Cần thiết cho sự sống: Duy trì quá trình hô hấp của con người, động vật, thực vật...",
            "  + Cần thiết cho sự cháy: Duy trì quá trình đốt cháy các loại nhiên liệu (củi, than, gas, xăng dầu) để đun nấu, sưởi ấm, chạy động cơ, thắp sáng..."
        ]
    }
]

# ==============================================================================
# PHẦN 3: PHÂN MÔN SINH HỌC
# ==============================================================================

SINH_CHAPTER1_THEORY = [
    "Trình bày được cách sử dụng, bảo quản một số dụng cụ thông thường khi học tập môn Khoa học tự nhiên (kính lúp, kính hiển vi quang học...)."
]

SINH_CHAPTER1_EXERCISES = [
    {
        "id": "BT1",
        "question": "Trên kính hiển vi quang học, bộ phận nào dùng để điều chỉnh giúp hình ảnh mẫu vật hiển thị rõ nét hơn?",
        "solution": "Ốc vi cấp (núm vi chỉnh / ốc tinh chỉnh) là bộ phận dùng để điều chỉnh nâng hạ ống kính với khoảng cách rất nhỏ, giúp lấy nét chính xác và làm cho hình ảnh mẫu vật hiển thị rõ nét, chi tiết nhất sau khi đã điều chỉnh sơ bộ bằng ốc to (ốc thô)."
    },
    {
        "id": "BT2",
        "question": "Hệ thống phóng đại của kính hiển vi quang học bao gồm hai bộ phận chính nào?",
        "solution": "Hệ thống phóng đại của kính hiển vi quang học bao gồm hai bộ phận chính là:\n1. Thị kính: Thấu kính ở đầu trên của ống kính, nơi người quan sát đặt mắt nhìn vào (thường có độ phóng đại 10x, 15x, 20x).\n2. Vật kính: Thấu kính ở mâm xoay phía dưới, hướng sát vào mẫu vật trên bàn kính (thường có độ phóng đại 4x, 10x, 40x, 100x)."
    },
    {
        "id": "BT3",
        "question": "Việc làm nào là đúng khi sử dụng và bảo quản kính lúp?",
        "solution": "- Khi sử dụng: Cầm tay cầm của kính lúp, đặt mặt kính gần sát mẫu vật, từ từ dịch chuyển kính ra xa mẫu vật cho đến khi nhìn thấy hình ảnh phóng to rõ nét nhất.\n- Khi bảo quản: Lau kính bằng khăn mềm sạch chuyên dụng; không để mặt kính tiếp xúc với vật cứng, nhám tránh trầy xước; để nơi khô ráo, thoáng mát và cất vào hộp hoặc bao đựng sau khi sử dụng."
    },
    {
        "id": "BT4",
        "question": "Bộ phận nào của kính để người quan sát đặt mắt vào để nhìn mẫu vật?",
        "solution": "Bộ phận để người quan sát đặt mắt vào để nhìn mẫu vật là Thị kính (kính mắt)."
    }
]

SINH_CHAPTER5_THEORY = [
    "- Nêu được khái niệm tế bào. Chức năng của tế bào: Tế bào là đơn vị cơ bản cấu tạo nên mọi cơ thể sống và thực hiện các chức năng sống.",
    "- Nêu được hình dạng và kích thước của một số loại tế bào (hình cầu, hình đĩa, hình sợi, hình que, hình sao...).",
    "- Nêu được ý nghĩa của sự lớn lên và sinh sản của tế bào: Giúp cơ thể lớn lên, thay thế các tế bào già hoặc tế bào bị tổn thương.",
    "- Nhận biết được tế bào là đơn vị cấu trúc của sự sống.",
    "- Nhận biết được lục lạp là bào quan thực hiện chức năng quang hợp ở cây xanh.",
    "- Thông qua quan sát hình ảnh phân biệt được tế bào động vật, tế bào thực vật (tế bào thực vật có thêm thành tế bào, lục lạp, không bào lớn).",
    "- Thông qua quan sát hình ảnh phân biệt được tế bào nhân thực, tế bào nhân sơ (tế bào nhân sơ chưa có màng nhân bao bọc).",
    "- Giải thích vì sao tế bào là đơn vị cấu trúc của sự sống: Vì mọi cơ thể sống từ đơn bào đến đa bào đều được cấu tạo từ một hoặc nhiều tế bào.",
    "- Trình bày được cấu tạo tế bào và chức năng ba thành phần chính:",
    "  + Màng tế bào: Bao bọc, bảo vệ tế bào và kiểm soát sự trao đổi chất giữa tế bào với môi trường.",
    "  + Chất tế bào: Nơi diễn ra phần lớn các hoạt động trao đổi chất của tế bào, chứa các bào quan.",
    "  + Nhân tế bào (hoặc vùng nhân): Chứa vật chất di truyền, điều khiển mọi hoạt động sống của tế bào.",
    "- Dựa vào sơ đồ, nhận biết được sự lớn lên và sinh sản của tế bào (từ 1 tế bào ban đầu qua mỗi lần phân chia tạo ra 2 tế bào con: 1 → 2 → 4 → 8 → ... → 2^n tế bào)."
]

SINH_CHAPTER5_EXERCISES = [
    {
        "id": "BT1",
        "question": "Vì sao khi quan sát tế bào biểu bì vảy hành dưới kính hiển vi quang học, ta thấy chúng có hình đa giác và đứng cố định, trong khi quan sát tế bào trứng cá bằng mắt thường hoặc kính lúp lại thấy có hình tròn và tách rời nhau?",
        "solution": "- Tế bào biểu bì vảy hành là tế bào thực vật: Bên ngoài màng tế bào có lớp thành tế bào dày và cứng bằng xenlulozơ (cellulose), giúp định hình tế bào dạng đa giác cố định và gắn kết các tế bào sát nhau thành mô vững chắc.\n- Tế bào trứng cá là tế bào động vật đơn lẻ: Không có thành tế bào bao bọc, chỉ có màng sinh chất mềm dẻo, sức căng bề mặt làm tế bào co lại thành hình cầu tròn đều; đồng thời các tế bào trứng cá tách rời nhau trong xoang cơ thể nên không đứng cố định thành khối."
    },
    {
        "id": "BT2",
        "question": "Kích thước của tế bào đa dạng như thế nào? Để quan sát tế bào, chúng ta thường dùng dụng cụ gì?",
        "solution": "- Kích thước tế bào vô cùng đa dạng: Đa số tế bào có kích thước hiển vi rất nhỏ (từ vài micromet μm đến vài chục micromet), mắt thường không thể nhìn thấy được (như vi khuẩn, tế bào biểu bì thực vật, tế bào máu...); một số ít tế bào có kích thước lớn nhìn thấy được bằng mắt thường (như tế bào trứng cá, tép bưởi, lòng đỏ trứng gà/chim...).\n- Dụng cụ quan sát: Đối với các tế bào hiển vi ta dùng kính hiển vi quang học; đối với các tế bào có kích thước tương đối lớn ta có thể dùng kính lúp hoặc quan sát bằng mắt thường."
    },
    {
        "id": "BT3",
        "question": "Hình dạng của tế bào phụ thuộc vào yếu tố nào? Lấy 2 ví dụ minh họa.",
        "solution": "- Hình dạng của tế bào phụ thuộc vào và thích nghi tối ưu với chức năng mà tế bào đó đảm nhận trong cơ thể sinh vật.\n- Ví dụ minh họa:\n  1. Tế bào hồng cầu người có hình đĩa lõm hai mặt: Giúp tăng diện tích bề mặt tiếp xúc để vận chuyển khí oxygen và carbon dioxide hiệu quả hơn, đồng thời dễ dàng biến dạng uốn cong để di chuyển qua các mao mạch máu hẹp.\n  2. Tế bào thần kinh (nơ-ron) có hình sao với nhiều sợi nhánh và một sợi trục dài: Thích nghi hoàn hảo với chức năng tiếp nhận và truyền dẫn tín hiệu, xung thần kinh đi khắp các cơ quan trong cơ thể."
    },
    {
        "id": "BT4",
        "question": "Kể tên một số hình dạng tế bào tương ứng?",
        "solution": "- Hình cầu (hình tròn): Tế bào trứng động vật, vi khuẩn tụ cầu...\n- Hình đĩa lõm: Tế bào hồng cầu người.\n- Hình trụ / que: Trực khuẩn E.coli, tế bào biểu bì vảy hành...\n- Hình sợi dài: Tế bào cơ vân, tế bào sợi nấm...\n- Hình sao có nhiều nhánh: Tế bào thần kinh (nơ-ron).\n- Hình thoi: Tế bào cơ trơn."
    }
]

SINH_CHAPTER6_THEORY = [
    "- Nhận biết được cơ thể đơn bào và cơ thể đa bào thông qua hình ảnh. Lấy được ví dụ minh hoạ:",
    "  + Cơ thể đơn bào: Cấu tạo từ một tế bào duy nhất nhưng thực hiện đầy đủ các chức năng của một cơ thể sống (ví dụ: vi khuẩn E.coli, tảo lam, trùng roi xanh, trùng giày, nấm men...).",
    "  + Cơ thể đa bào: Cấu tạo từ nhiều tế bào, các tế bào phân hóa thành các mô, cơ quan, hệ cơ quan thực hiện các chức năng khác nhau (ví dụ: thực vật như cây lúa, cây hoa mai; động vật như con sư tử, con giun đất, con người...).",
    "- Quan sát và vẽ được hình cơ thể đơn bào (tảo lục đơn bào, trùng roi xanh...)."
]

SINH_CHAPTER6_EXERCISES = [
    {
        "id": "BT1",
        "question": "Vi khuẩn E.coli, Tế bào biểu bì vảy hành, tế bào trứng cá... là tế bào nhân thực hay nhân sơ?",
        "solution": "- Vi khuẩn E.coli: Là tế bào nhân sơ (chưa có màng nhân bao bọc vật chất di truyền, vùng nhân chỉ gồm phân tử DNA trần).\n- Tế bào biểu bì vảy hành: Là tế bào nhân thực (đã có màng nhân hoàn chỉnh bao bọc vật chất di truyền).\n- Tế bào trứng cá: Là tế bào nhân thực (đã có màng nhân hoàn chỉnh bao bọc nhân)."
    },
    {
        "id": "BT2",
        "question": "Quan sát hình ảnh tế bào trùng roi, chỉ ra các thành phần cơ thể trùng roi?",
        "solution": "Cơ thể trùng roi xanh là cơ thể đơn bào gồm các thành phần cấu tạo chính sau:\n1. Màng tế bào (màng sinh chất): Lớp màng mỏng bao bọc ngoài cùng bảo vệ tế bào.\n2. Chất tế bào: Chứa dịch tế bào và các bào quan thực hiện các hoạt động sống.\n3. Nhân tế bào: Mang thông tin di truyền, trung tâm điều khiển mọi hoạt động sống của cơ thể.\n4. Hạt diệp lục (lục lạp): Chứa chất diệp lục giúp trùng roi quang hợp chế tạo chất hữu cơ khi có ánh sáng.\n5. Điểm mắt (mắt đỏ): Giúp trùng roi cảm nhận ánh sáng để bơi về phía có ánh sáng.\n6. Roi bơi (tiên mao): Roi dài ở đầu cơ thể xoáy trong nước giúp trùng roi di chuyển linh hoạt.\n7. Không bào co bóp: Thu gom nước thừa và chất thải để bài tiết ra ngoài, điều hòa áp suất thẩm thấu."
    }
]
