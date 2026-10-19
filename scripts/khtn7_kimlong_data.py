# -*- coding: utf-8 -*-
"""
Dữ liệu chuẩn hóa Đề cương Ôn tập Giữa kì I - KHTN 7
Trường THCS Kim Long - Năm học 2026 - 2027
Biên soạn & Thiết kế: Thầy TRẦN VĂN THẠNH - 0777.470.803
"""

INFO = {
    "truong": "TRƯỜNG THCS KIM LONG",
    "de_muc": "ĐỀ CƯƠNG ÔN TẬP GIỮA KÌ I – KHTN 7",
    "nam_hoc": "NĂM HỌC 2026 - 2027",
    "mon": "MÔN: KHOA HỌC TỰ NHIÊN 7",
    "giao_vien": "Thầy TRẦN VĂN THẠNH",
    "sdt": "0777.470.803",
    "co_so_1": "CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế",
    "co_so_2": "CS2: 24 Đặng Thái Thân, TP Huế"
}

# ==============================================================================
# BÀI 1: PHƯƠNG PHÁP VÀ KĨ NĂNG HỌC TẬP MÔN KHOA HỌC TỰ NHIÊN
# ==============================================================================
BAI1_MCQ = [
    {
        "id": 1,
        "content": "“Trên cơ sở các số liệu và phân tích số liệu, con người có thể đưa ra các dự báo hay dự đoán tính chất của sự vật, hiện tượng, nguyên nhân của hiện tượng!” Đó là kĩ năng nào?",
        "options": [
            "A. Kĩ năng quan sát, phân loại.",
            "B. Kĩ năng liên kết tri thức.",
            "C. Kĩ năng dự báo.",
            "D. Kĩ năng đo."
        ],
        "correct": "C",
        "solution": "Kĩ năng dự báo là kĩ năng đưa ra các nhận định, dự đoán về những gì có thể xảy ra trong tương lai hoặc nguyên nhân của hiện tượng dựa trên cơ sở quan sát, thu thập và phân tích số liệu khoa học."
    },
    {
        "id": 2,
        "content": "Khẳng định nào dưới đây là không đúng?",
        "options": [
            "A. Dự báo là kĩ năng cần thiết trong nghiên cứu khoa học tự nhiên.",
            "B. Dự báo là kĩ năng không cần thiết của người làm nghiên cứu.",
            "C. Dự báo là kĩ năng dự đoán điều gì sẽ xảy ra dựa vào quan sát, kiến thức, suy luận của con người,... về các sự vật, hiện tượng.",
            "D. Kĩ năng dự báo thường được sử dụng trong bước dự đoán của phương pháp tìm hiểu tự nhiên."
        ],
        "correct": "B",
        "solution": "Khẳng định B không đúng vì dự báo là một trong những kĩ năng then chốt, vô cùng cần thiết đối với người làm nghiên cứu khoa học để định hướng nghiên cứu và đề xuất giả thuyết khoa học."
    },
    {
        "id": 3,
        "content": "Cho các bước sau:\n(1) Hình thành giả thuyết.\n(2) Quan sát và đặt câu hỏi.\n(3) Lập kế hoạch kiểm tra giả thuyết.\n(4) Thực hiện kế hoạch.\n(5) Kết luận.\nThứ tự sắp xếp đúng các bước trong phương pháp tìm hiểu tự nhiên là?",
        "options": [
            "A. (1) - (2) - (3) - (4) - (5).",
            "B. (2) - (1) - (3) - (4) - (5).",
            "C. (1) - (2) - (3) - (5) - (4).",
            "D. (2) - (1) - (3) - (5) - (4)."
        ],
        "correct": "B",
        "solution": "Quy trình chuẩn của phương pháp tìm hiểu tự nhiên gồm 5 bước:\n• Bước 1: Quan sát và đặt câu hỏi (2)\n• Bước 2: Hình thành giả thuyết (1)\n• Bước 3: Lập kế hoạch kiểm tra giả thuyết (3)\n• Bước 4: Thực hiện kế hoạch (4)\n• Bước 5: Kết luận (5)."
    },
    {
        "id": 4,
        "content": "Con người có thể định lượng được các sự vật và hiện tượng tự nhiên dựa trên kĩ năng nào?",
        "options": [
            "A. Kĩ năng quan sát, phân loại.",
            "B. Kĩ năng liên kết tri thức.",
            "C. Kĩ năng dự báo.",
            "D. Kĩ năng đo."
        ],
        "correct": "D",
        "solution": "Kĩ năng đo giúp xác định được các giá trị định lượng (kích thước, khối lượng, nhiệt độ, thời gian, thể tích,...) của các sự vật và hiện tượng tự nhiên bằng các dụng cụ đo thích hợp."
    },
    {
        "id": 5,
        "content": "Đâu không phải là kĩ năng cần vận dụng vào phương pháp tìm hiểu tự nhiên?",
        "options": [
            "A. Kĩ năng chiến đấu đặc biệt.",
            "B. Kĩ năng quan sát.",
            "C. Kĩ năng dự báo.",
            "D. Kĩ năng đo đạc."
        ],
        "correct": "A",
        "solution": "Các kĩ năng cốt lõi trong tìm hiểu tự nhiên bao gồm: quan sát, phân loại, liên kết tri thức, dự báo, đo đạc và viết báo cáo. \"Kĩ năng chiến đấu đặc biệt\" không thuộc kĩ năng học tập và nghiên cứu môn KHTN."
    },
    {
        "id": 6,
        "content": "Cho các bước sau:\n(1) Thực hiện phép đo, ghi kết quả đo và xử lí số liệu đo.\n(2) Ước lượng để lựa chọn dụng cụ/thiết bị đo phù hợp.\n(3) Phân tích kết quả và thảo luận về kết quả nghiên cứu thu được.\n(4) Đánh giá độ chính xác của kết quả đo căn cứ vào loại dụng cụ đo và cách đo.\nTrình tự các bước hình thành kĩ năng đo là:",
        "options": [
            "A. (1) -> (2) -> (3) -> (4).",
            "B. (1) -> (3) -> (2) -> (4).",
            "C. (3) -> (2) -> (4) -> (1).",
            "D. (2) -> (1) -> (4) -> (3)."
        ],
        "correct": "D",
        "solution": "Trình tự thực hiện kĩ năng đo khoa học gồm:\n1. Ước lượng để lựa chọn dụng cụ/thiết bị đo phù hợp (2)\n2. Thực hiện phép đo, ghi kết quả và xử lí số liệu (1)\n3. Đánh giá độ chính xác của kết quả đo căn cứ vào dụng cụ và cách đo (4)\n4. Phân tích kết quả và thảo luận (3)."
    },
    {
        "id": 7,
        "type": "matching",
        "content": "Hãy kết nối các thông tin ở cột A với cột B tạo thành sự liên kết giữa sự vật với hiện tượng hoặc hiện tượng với hiện tượng:",
        "table_data": [
            ("1. Nước mưa", "a. do ánh sáng từ Mặt Trời"),
            ("2. Một số loài thực vật", "b. ảnh hưởng đến sinh trưởng và phát triển của thực vật"),
            ("3. Trời nắng", "c. có khi trời mưa"),
            ("4. Phân bón", "d. rụng lá vào mùa đông")
        ],
        "correct": "1 - c, 2 - d, 3 - a, 4 - b",
        "solution": "Nối đúng các cặp thông tin:\n• 1 nối c: Nước mưa -> có khi trời mưa (sự xuất hiện của nước mưa gắn liền với hiện tượng trời mưa).\n• 2 nối d: Một số loài thực vật -> rụng lá vào mùa đông (sự thích nghi của thực vật vào mùa lạnh).\n• 3 nối a: Trời nắng -> do ánh sáng từ Mặt Trời (nguồn năng lượng chiếu sáng).\n• 4 nối b: Phân bón -> ảnh hưởng đến sinh trưởng và phát triển của thực vật (cung cấp chất dinh dưỡng)."
    },
    {
        "id": 8,
        "content": "Trong Hình 1.1, ban đầu bình a chứa nước, bình b chứa một vật rắn không thấm nước. Khi đổ hết nước từ bình a sang bình b thì mức nước trong bình b được vẽ trong hình. Thể tích của vật rắn là:",
        "image": "hinh_1_1.png",
        "options": [
            "A. 33 ml.",
            "B. 73 ml.",
            "C. 32,5 ml.",
            "D. 35,2 ml."
        ],
        "correct": "A",
        "solution": "• Quan sát bình a: Thể tích nước ban đầu là V_nước = 37 mL (mặt thoáng nằm ở vạch 37 mL).\n• Quan sát bình b sau khi đổ toàn bộ nước vào: Mực nước dâng lên mức V_tổng = 70 mL.\n• Thể tích của vật rắn không thấm nước là: V_vật = V_tổng - V_nước = 70 - 37 = 33 mL.\n-> Chọn đáp án A."
    },
    {
        "id": 9,
        "content": "Cổng quang điện có vai trò:",
        "options": [
            "A. Điều khiển mở đồng hồ đo thời gian hiện số.",
            "B. Điều khiển đóng đồng hồ đo thời gian hiện số.",
            "C. Điều khiển mở/đóng đồng hồ đo thời gian hiện số.",
            "D. Gửi tín hiệu điện tự tới đồng hồ."
        ],
        "correct": "C",
        "solution": "Cổng quang điện là thiết bị cảm biến quang học, có vai trò điều khiển việc đóng và mở đồng hồ đo thời gian hiện số khi có vật chắn chùm sáng đi qua, giúp đo thời gian chuyển động cực kì chính xác."
    },
    {
        "id": 10,
        "content": "Một bản báo cáo thực hành cần có những nội dung nào, sắp xếp lại theo thứ tự nội dung bản báo cáo:\n(1). Kết luận.\n(2). Mục đích thí nghiệm.\n(3). Kết quả.\n(4). Các bước tiến hành.\n(5). Chuẩn bị.\n(6). Thảo luận.",
        "options": [
            "A. (1) - (2) - (3) - (4) - (5) - (6).",
            "B. (2) - (5) - (4) - (3) - (6) - (1).",
            "C. (1) - (2) - (6) - (3) - (5) - (4).",
            "D. (2) - (1) - (3) - (5) - (6) - (4)."
        ],
        "correct": "B",
        "solution": "Cấu trúc chuẩn của một bản báo cáo thực hành khoa học tự nhiên:\n1. Mục đích thí nghiệm (2)\n2. Chuẩn bị dụng cụ, mẫu vật (5)\n3. Các bước tiến hành (4)\n4. Kết quả quan sát/đo đạc (3)\n5. Thảo luận (6)\n6. Kết luận (1)."
    }
]

# ==============================================================================
# PHÂN MÔN KHTN 1: HÓA HỌC
# ==============================================================================
KHTN1_THEORY = [
    {
        "title": "Bài 2. Nguyên tử",
        "items": [
            "Nguyên tử là hạt vô cùng nhỏ, trung hòa về điện, tạo nên các chất.",
            "Cấu tạo nguyên tử gồm 2 phần chính: Hạt nhân ở tâm (mang điện tích dương) và lớp vỏ gồm các electron (mang điện tích âm) chuyển động xung quanh hạt nhân.",
            "Hạt nhân nguyên tử gồm 2 loại hạt: Proton (kí hiệu p, mang điện tích dương +1) và neutron (kí hiệu n, không mang điện). Riêng nguyên tử hydrogen thông thường chỉ có 1 proton, không có neutron.",
            "Lớp vỏ nguyên tử: Tạo bởi các electron (kí hiệu e, mang điện tích âm -1). Electron chuyển động rất nhanh quanh hạt nhân theo từng quỹ đạo xác định tạo thành các lớp electron.",
            "Trong mọi nguyên tử: Số proton = Số electron. Do đó, nguyên tử luôn trung hòa về điện.",
            "Khối lượng nguyên tử (KLNT): Rất nhỏ bé, được đo bằng đơn vị amu (atomic mass unit). Do khối lượng electron rất nhỏ (không đáng kể), khối lượng nguyên tử tập trung hầu hết ở hạt nhân: m_nguyên tử ≈ m_hạt nhân = (số p + số n) amu."
        ]
    },
    {
        "title": "Bài 3. Nguyên tố hóa học",
        "items": [
            "Nguyên tố hóa học là tập hợp các nguyên tử cùng loại, có cùng số proton trong hạt nhân.",
            "Số proton là số đặc trưng của một nguyên tố hóa học. Các nguyên tử có cùng số proton đều thuộc về cùng một nguyên tố hóa học.",
            "Kí hiệu hóa học: Mỗi nguyên tố hóa học được biểu diễn bằng 1 hoặc 2 chữ cái (chữ cái đầu viết hoa, chữ thứ hai nếu có viết thường) theo quy ước IUPAC quốc tế (ví dụ: H, O, N, Na, Ca, Fe, Cu, S,...)."
        ]
    }
]

KHTN1_MCQ = [
    {
        "id": 1,
        "content": "Nguyên tử được tạo nên bởi những loại hạt nào?",
        "options": [
            "A. hạt proton, neutron, electron.",
            "B. hạt proton, neutron.",
            "C. hạt proton, electron.",
            "D. neutron, electron."
        ],
        "correct": "A",
        "solution": "Nguyên tử được cấu tạo bởi 3 loại hạt hạ nguyên tử cơ bản: proton (mang điện dương), neutron (không mang điện) trong hạt nhân và electron (mang điện âm) ở lớp vỏ."
    },
    {
        "id": 2,
        "content": "Phát biểu nào sau đây không mô tả đúng mô hình nguyên tử của Rutherford - Bohr?",
        "options": [
            "A. Nguyên tử có cấu tạo rỗng, gồm hạt nhân ở tâm nguyên tử và các electron ở vỏ nguyên tử.",
            "B. Nguyên tử có cấu tạo đặc khít, gồm hạt nhân nguyên tử và các electron.",
            "C. Electron chuyển động xung quanh hạt nhân theo những quỹ đạo xác định tạo thành các lớp electron.",
            "D. Hạt nhân nguyên tử mang điện tích dương, electron mang điện tích âm."
        ],
        "correct": "B",
        "solution": "Theo mô hình Rutherford - Bohr, nguyên tử có cấu tạo rỗng (kích thước hạt nhân rất nhỏ so với kích thước toàn bộ nguyên tử), không phải cấu tạo đặc khít. Do đó phát biểu B là không đúng."
    },
    {
        "id": 3,
        "content": "Trừ hạt nhân của nguyên tử hydrogen, hạt nhân các nguyên tử còn lại được tạo thành từ hạt:",
        "options": [
            "A. electron và proton.",
            "B. electron, proton và neutron.",
            "C. neutron và electron.",
            "D. proton và neutron."
        ],
        "correct": "D",
        "solution": "Hạt nhân của đa số các nguyên tử đều được tạo nên từ 2 loại hạt là proton (mang điện dương) và neutron (không mang điện). Riêng đồng vị phổ biến nhất của hydrogen chỉ có 1 proton."
    },
    {
        "id": 4,
        "content": "Khối lượng của nguyên tử được đo bằng đơn vị:",
        "options": [
            "A. gam.",
            "B. amu.",
            "C. mL.",
            "D. kg."
        ],
        "correct": "B",
        "solution": "Khối lượng nguyên tử vô cùng nhỏ bé nên trong khoa học người ta dùng đơn vị khối lượng nguyên tử là amu (atomic mass unit, 1 amu ≈ 1,6605 × 10^-24 g)."
    },
    {
        "id": 5,
        "content": "Nguyên tử Calcium có 20 electron ở vỏ nguyên tử. Hạt nhân nguyên tử calcium có số proton là:",
        "options": [
            "A. 2.",
            "B. 10.",
            "C. 18.",
            "D. 20."
        ],
        "correct": "D",
        "solution": "Vì nguyên tử luôn trung hòa về điện nên số proton trong hạt nhân luôn bằng số electron ở lớp vỏ: số p = số e = 20."
    },
    {
        "id": 6,
        "content": "Trong hạt nhân nguyên tử, hạt mang điện là:",
        "options": [
            "A. electron.",
            "B. proton.",
            "C. neutron.",
            "D. proton và neutron."
        ],
        "correct": "B",
        "solution": "Trong hạt nhân có hạt proton mang điện tích dương (+1) và hạt neutron không mang điện (điện tích = 0). Electron cũng là hạt mang điện âm nhưng nằm ở lớp vỏ, không nằm trong hạt nhân."
    },
    {
        "id": 7,
        "content": "Một nguyên tử có 10 proton trong hạt nhân. Theo mô hình nguyên tử của Rutherford - Bohr, số lớp electron của nguyên tử đó là:",
        "options": [
            "A. 1.",
            "B. 2.",
            "C. 3.",
            "D. 4."
        ],
        "correct": "B",
        "solution": "Nguyên tử có 10 proton => có 10 electron. Sự phân bố electron vào các lớp: Lớp thứ nhất có tối đa 2 electron; lớp thứ hai chứa 8 electron còn lại (tổng 2 + 8 = 10). Do đó nguyên tử có 2 lớp electron (nguyên tử Neon - Ne)."
    },
    {
        "id": 8,
        "content": "Trong hạt nhân nguyên tử fluorine có 9 proton. Số electron ở lớp ngoài cùng của vỏ nguyên tử fluorine là:",
        "options": [
            "A. 2.",
            "B. 5.",
            "C. 7.",
            "D. 8."
        ],
        "correct": "C",
        "solution": "Fluorine có 9 proton => có 9 electron. Phân bố electron: Lớp thứ nhất có 2 electron; lớp thứ hai (lớp ngoài cùng) có 9 - 2 = 7 electron."
    },
    {
        "id": 9,
        "content": "Số electron tối đa ở lớp electron thứ nhất là:",
        "options": [
            "A. 1.",
            "B. 2.",
            "C. 3.",
            "D. 8."
        ],
        "correct": "B",
        "solution": "Theo quy luật phân bố electron trong mô hình Rutherford - Bohr, lớp electron thứ nhất (lớp trong cùng gần hạt nhân nhất) chứa tối đa 2 electron."
    },
    {
        "id": 10,
        "content": "Nguyên tử luôn trung hoà về điện nên:",
        "options": [
            "A. số hạt proton = số hạt neutron.",
            "B. số hạt electron = số hạt neutron.",
            "C. số hạt electron = số hạt proton.",
            "D. số hạt proton = số hạt electron = số hạt neutron."
        ],
        "correct": "C",
        "solution": "Mỗi hạt proton mang 1 đơn vị điện tích dương (+1), mỗi hạt electron mang 1 đơn vị điện tích âm (-1). Để nguyên tử trung hòa về điện thì tổng điện tích dương phải triệt tiêu tổng điện tích âm, nghĩa là: số electron = số proton."
    },
    {
        "id": 11,
        "content": "Nguyên tử X có 10 proton, điện tích hạt nhân của nguyên tử X là:",
        "options": [
            "A. +1.",
            "B. +8.",
            "C. +7.",
            "D. +10."
        ],
        "correct": "D",
        "solution": "Điện tích hạt nhân bằng tổng điện tích của các hạt proton có trong hạt nhân. Vì có 10 proton nên điện tích hạt nhân là +10."
    },
    {
        "id": 12,
        "content": "Trong hạt nhân nguyên tử Lưu huỳnh (sulfur) có 16 proton. Số electron trong các lớp của vỏ nguyên tử sulfur, viết từ lớp trong ra lớp ngoài lần lượt là:",
        "options": [
            "A. 2,14.",
            "B. 2,6,8.",
            "C. 2,8,6.",
            "D. 2,9,5."
        ],
        "correct": "C",
        "solution": "Nguyên tử sulfur có 16 proton => có 16 electron. Phân bố từ trong ra ngoài:\n• Lớp 1: tối đa 2 electron\n• Lớp 2: tối đa 8 electron\n• Lớp 3: còn lại 16 - (2 + 8) = 6 electron.\nVậy thứ tự các lớp là: 2, 8, 6."
    },
    {
        "id": 13,
        "content": "Nguyên tử X có tổng số hạt là 52, trong đó số proton là 17. Số electron và số neutron của X lần lượt là?",
        "options": [
            "A. 18 và 17.",
            "B. 19 và 16.",
            "C. 16 và 19.",
            "D. 17 và 18."
        ],
        "correct": "D",
        "solution": "• Vì nguyên tử trung hòa về điện nên số e = số p = 17.\n• Tổng số hạt = p + e + n = 52 => n = 52 - 2p = 52 - 2 × 17 = 18.\nVậy số electron là 17 và số neutron là 18 (X là Chlorine - Cl)."
    },
    {
        "id": 14,
        "content": "Nguyên tố hoá học là tập hợp nguyên tử cùng loại có....",
        "options": [
            "A. cùng số neutron trong hạt nhân.",
            "B. cùng số proton trong hạt nhân.",
            "C. cùng số electron trong hạt nhân.",
            "D. cùng số proton và số neutron trong hạt nhân."
        ],
        "correct": "B",
        "solution": "Định nghĩa chuẩn: Nguyên tố hóa học là tập hợp các nguyên tử cùng loại, có cùng số proton trong hạt nhân."
    },
    {
        "id": 15,
        "content": "Điền từ phù hợp vào chỗ trống: “Số...... là số đặc trưng của một nguyên tố hoá học”",
        "options": [
            "A. proton.",
            "B. electron.",
            "C. neutron.",
            "D. neutron và electron."
        ],
        "correct": "A",
        "solution": "Số proton là đại lượng đặc trưng cơ bản nhất của một nguyên tố hóa học, quyết định tên gọi và vị trí của nguyên tố trong bảng tuần hoàn."
    },
    {
        "id": 16,
        "content": "Cho các nguyên tử được kí hiệu bởi các chữ cái và số proton trong mỗi nguyên tử như sau:\n\n| Nguyên tử | X | Y | Z | R | E | Q |\n| :--- | :---: | :---: | :---: | :---: | :---: | :---: |\n| **Số proton** | 5 | 8 | 17 | 6 | 9 | 17 |\n\nCác nguyên tử thuộc cùng một nguyên tố hoá học là:",
        "options": [
            "A. X, Y.",
            "B. Z, Q.",
            "C. R, E.",
            "D. Y, E."
        ],
        "correct": "B",
        "solution": "Các nguyên tử thuộc cùng một nguyên tố hóa học phải có cùng số proton trong hạt nhân. Tra bảng ta thấy nguyên tử Z và Q đều có 17 proton nên chúng thuộc cùng một nguyên tố hóa học (nguyên tố Chlorine)."
    }
]

KHTN1_TF = [
    {
        "id": 1,
        "content": "Cho các phát biểu về nguyên tử sau:",
        "sub_items": [
            {
                "id": "a",
                "text": "Nguyên tử trung hoà về điện.",
                "answer": True,
                "reason": "Vì trong nguyên tử số hạt proton mang điện tích dương luôn bằng số hạt electron mang điện tích âm."
            },
            {
                "id": "b",
                "text": "Khối lượng nguyên tử tập trung chủ yếu ở hạt nhân.",
                "answer": True,
                "reason": "Khối lượng của electron rất nhỏ (khoảng 0,00055 amu, coi như không đáng kể), do đó khối lượng nguyên tử tập trung gần như hoàn toàn ở khối lượng các hạt proton và neutron trong hạt nhân."
            },
            {
                "id": "c",
                "text": "Trong nguyên tử, số hạt mang điện tích dương bằng số hạt mang điện tích âm nên số hạt electron bằng số hạt neutron.",
                "answer": False,
                "reason": "Số hạt mang điện dương là proton, số hạt mang điện âm là electron => số proton = số electron. Neutron là hạt không mang điện và số neutron không nhất thiết bằng số electron."
            },
            {
                "id": "d",
                "text": "Vỏ nguyên tử gồm các lớp electron có khoảng cách khác nhau đối với hạt nhân.",
                "answer": True,
                "reason": "Theo mô hình Rutherford - Bohr, các electron sắp xếp thành từng lớp đồng tâm từ gần đến xa hạt nhân."
            }
        ]
    },
    {
        "id": 2,
        "content": "Cho các phát biểu về đặc điểm cấu tạo nguyên tử:",
        "sub_items": [
            {
                "id": "a",
                "text": "Nguyên tử có cấu tạo rỗng, gồm hạt nhân ở tâm nguyên tử và các electron ở vỏ nguyên tử.",
                "answer": True,
                "reason": "Kích thước hạt nhân nhỏ hơn kích thước toàn nguyên tử khoảng 10.000 lần, phần không gian giữa hạt nhân và electron là rỗng."
            },
            {
                "id": "b",
                "text": "Nguyên tử có cấu tạo đặc khít, gồm hạt nhân nguyên tử và các electron.",
                "answer": False,
                "reason": "Nguyên tử có cấu tạo rỗng, không phải cấu tạo đặc khít."
            },
            {
                "id": "c",
                "text": "Electron chuyển động xung quanh hạt nhân theo những quỹ đạo xác định tạo thành các lớp electron.",
                "answer": True,
                "reason": "Đây là luận điểm chính xác của mô hình nguyên tử Rutherford - Bohr giảng dạy trong SGK KHTN 7."
            },
            {
                "id": "d",
                "text": "Hạt nhân nguyên tử mang điện dương, lớp vỏ tạo bởi neutron mang điện tích âm.",
                "answer": False,
                "reason": "Lớp vỏ được tạo bởi các hạt electron (mang điện tích âm). Neutron nằm trong hạt nhân và không mang điện."
            }
        ]
    },
    {
        "id": 3,
        "content": "Nguyên tử X có 13 electron và 14 neutron. Xét tính đúng/sai của các nhận định sau:",
        "sub_items": [
            {
                "id": "a",
                "text": "X có điện tích hạt nhân là +13.",
                "answer": True,
                "reason": "Vì số p = số e = 13 nên điện tích hạt nhân là +13."
            },
            {
                "id": "b",
                "text": "Khối lượng nguyên tử KLNT_X = 40 amu.",
                "answer": False,
                "reason": "Khối lượng nguyên tử X xấp xỉ bằng tổng khối lượng hạt nhân: m_X ≈ p + n = 13 + 14 = 27 amu (chứ không phải 40 amu)."
            },
            {
                "id": "c",
                "text": "X có 3 lớp electron, 4 electron ở lớp ngoài cùng.",
                "answer": False,
                "reason": "Với 13 electron, phân bố vào các lớp: Lớp 1 có 2e, Lớp 2 có 8e, Lớp 3 có 3e. Vậy X có 3 lớp electron nhưng lớp ngoài cùng chỉ có 3 electron (chứ không phải 4)."
            },
            {
                "id": "d",
                "text": "Tổng số hạt trong hạt nhân của nguyên tử X là 27.",
                "answer": True,
                "reason": "Hạt nhân gồm proton và neutron, tổng số hạt = p + n = 13 + 14 = 27 hạt."
            }
        ]
    },
    {
        "id": 4,
        "content": "Cho các phát biểu về nguyên tố hóa học:",
        "sub_items": [
            {
                "id": "a",
                "text": "Nguyên tố hóa học là tập hợp những nguyên tử có cùng số proton trong hạt nhân.",
                "answer": True,
                "reason": "Đúng theo định nghĩa nguyên tố hóa học."
            },
            {
                "id": "b",
                "text": "Số proton trong hạt nhân được gọi là số hiệu nguyên tử.",
                "answer": True,
                "reason": "Số proton chính là số hiệu nguyên tử (kí hiệu Z)."
            },
            {
                "id": "c",
                "text": "Các nguyên tử thuộc cùng một nguyên tố hóa học luôn có cùng số neutron.",
                "answer": False,
                "reason": "Các nguyên tử cùng nguyên tố hóa học chỉ bắt buộc cùng số proton, chúng có thể có số neutron khác nhau (các đồng vị của nhau)."
            },
            {
                "id": "d",
                "text": "Nguyên tố hóa học được đặc trưng bởi số proton trong hạt nhân.",
                "answer": True,
                "reason": "Số proton quyết định điện tích hạt nhân, cấu hình electron và tính chất hóa học của nguyên tố."
            }
        ]
    }
]

KHTN1_SHORT = [
    {
        "id": 1,
        "content": "Trong hạt nhân của hầu hết nguyên tử có bao nhiêu loại hạt?",
        "answer": "2 loại hạt (proton và neutron)",
        "solution": "Hạt nhân của đa số các nguyên tử được tạo thành từ 2 loại hạt là proton (mang điện dương) và neutron (không mang điện). Ngoại lệ duy nhất là hạt nhân nguyên tử hydrogen thông thường chỉ có 1 proton."
    },
    {
        "id": 2,
        "content": "Khối lượng của nguyên tử X bằng 19 amu, số electron của nguyên tử đó là 9. Số neutron của nguyên tử X là bao nhiêu?",
        "answer": "10",
        "solution": "• Vì nguyên tử trung hòa về điện nên: số proton = số electron = 9.\n• Khối lượng nguyên tử xấp xỉ bằng tổng khối lượng proton và neutron: m_X ≈ p + n = 19 amu.\n=> Số neutron: n = 19 - 9 = 10 (X là nguyên tố Fluorine - F)."
    },
    {
        "id": 3,
        "content": "Nguyên tử carbon có điện tích hạt nhân là +6, số electron của nguyên tử carbon là bao nhiêu?",
        "answer": "6",
        "solution": "Điện tích hạt nhân là +6 nghĩa là hạt nhân carbon có 6 proton. Do nguyên tử trung hòa về điện nên số electron = số proton = 6."
    },
    {
        "id": 4,
        "content": "Nguyên tử X có 11 proton và 12 neutron. Tổng số hạt trong nguyên tử X là bao nhiêu?",
        "answer": "34",
        "solution": "• Số electron trong nguyên tử: e = p = 11.\n• Tổng số hạt trong nguyên tử = p + e + n = 11 + 11 + 12 = 34 hạt (X là Sodium - Na)."
    },
    {
        "id": 5,
        "content": "Nguyên tử Y có 4 lớp electron, 1 electron ở lớp ngoài cùng, số electron trong lớp vỏ của nguyên tử Y là bao nhiêu?",
        "answer": "19",
        "solution": "Theo mô hình vỏ electron: Lớp 1 có tối đa 2 electron; lớp 2 có tối đa 8 electron; lớp 3 bão hòa 8 electron (ở trạng thái bền); lớp 4 (lớp ngoài cùng) có 1 electron.\n=> Tổng số electron = 2 + 8 + 8 + 1 = 19 electron (Y là Potassium - K)."
    }
]

KHTN1_ESSAY = [
    {
        "id": 1,
        "content": "Điền từ/cụm từ thích hợp vào chỗ trống để hoàn thành các câu sau:\na. .................... là hạt vô cùng nhỏ, .................... về điện.\nb. Nguyên tử gồm hạt nhân mang điện tích .................... và vỏ nguyên tử tạo bởi các .................... mang điện tích ....................\nc. Khối lượng nguyên tử được coi bằng khối lượng của .................... và có đơn vị là ....................\nd. Các electron ở lớp .................... quyết định tính chất hóa học của nguyên tử.",
        "solution": "Các từ cần điền theo thứ tự:\na. **Nguyên tử** là hạt vô cùng nhỏ, **trung hoà** về điện.\nb. Nguyên tử gồm hạt nhân mang điện tích **dương** và vỏ nguyên tử tạo bởi các **electron** mang điện tích **âm**.\nc. Khối lượng nguyên tử được coi bằng khối lượng của **hạt nhân** và có đơn vị là **amu**.\nd. Các electron ở lớp **ngoài cùng** quyết định tính chất hóa học của nguyên tử."
    },
    {
        "id": 2,
        "content": "Định nghĩa nguyên tố hóa học.",
        "solution": "• **Định nghĩa:** Nguyên tố hóa học là tập hợp những nguyên tử cùng loại, có cùng số proton trong hạt nhân.\n• Các nguyên tử của cùng một nguyên tố hóa học đều có tính chất hóa học giống nhau."
    },
    {
        "id": 3,
        "content": "Nguyên tử X có 16 electron ở lớp vỏ.\na. Xác định điện tích hạt nhân của X.\nb. Vẽ sơ đồ cấu tạo nguyên tử của nguyên tố X.",
        "image": "so_do_sulfur.png",
        "solution": "a. Do nguyên tử trung hòa về điện nên số proton = số electron = 16.\n   => Điện tích hạt nhân của X là **+16** (nguyên tố Sulfur - S).\nb. Sơ đồ cấu tạo nguyên tử X:\n   • Hạt nhân ở tâm mang điện tích **16+**.\n   • Vỏ electron gồm 3 lớp:\n     - Lớp thứ nhất (trong cùng): có **2 electron**.\n     - Lớp thứ hai: có **8 electron**.\n     - Lớp thứ ba (ngoài cùng): có **6 electron**."
    },
    {
        "id": 4,
        "content": "Tổng số hạt trong nguyên tử của một nguyên tố X là 40, trong đó số hạt không mang điện là 14.\na. Xác định số hạt proton, electron, neutron trong X.\nb. Vẽ sơ đồ cấu tạo nguyên tử của nguyên tố X.\nc. Xác định khối lượng nguyên tử (KLNT) của X.",
        "image": "so_do_aluminium.png",
        "solution": "a. Theo đề bài:\n   • Hạt không mang điện là neutron: **n = 14**.\n   • Tổng số hạt: 2p + n = 40 (vì trong nguyên tử số p = số e).\n   => 2p + 14 = 40 => 2p = 26 => **p = e = 13**.\n   Vậy nguyên tử X có: **13 proton, 13 electron và 14 neutron** (X là Aluminium - Al).\nb. Sơ đồ cấu tạo nguyên tử X:\n   • Hạt nhân mang điện tích **13+**.\n   • Vỏ electron gồm 3 lớp:\n     - Lớp 1: có **2 electron**.\n     - Lớp 2: có **8 electron**.\n     - Lớp 3 (ngoài cùng): có **3 electron**.\nc. Khối lượng nguyên tử của X:\n   • KLNT_X ≈ p + n = 13 + 14 = **27 amu**."
    }
]

# ==============================================================================
# PHÂN MÔN KHTN 2: VẬT LÍ
# ==============================================================================
KHTN2_THEORY = [
    {
        "title": "1. Tốc độ chuyển động",
        "items": [
            "a. Khái niệm tốc độ: Tốc độ là đại lượng cho biết mức độ nhanh hay chậm của chuyển động và được xác định bằng quãng đường vật đi được trong một đơn vị thời gian.",
            "Công thức tính tốc độ: v = s / t. Trong đó: v là tốc độ của vật; s là quãng đường vật đi được; t là thời gian vật đi hết quãng đường đó.",
            "b. Đơn vị tốc độ: Trong hệ đơn vị đo lường chính thức ở nước ta, tốc độ được đo bằng đơn vị mét trên giây (m/s) và kilômét trên giờ (km/h). Quy đổi: 1 m/s = 3,6 km/h hay 1 km/h = 1 / 3,6 m/s ≈ 0,28 m/s.",
            "Ngoài ra còn có các đơn vị khác: m/min, cm/s, mm/s. Trên thực tế, để đo trực tiếp tốc độ của các phương tiện giao thông, người ta dùng tốc kế (đồng hồ đo tốc độ)."
        ]
    },
    {
        "title": "2. Đo tốc độ",
        "items": [
            "a. Đo tốc độ bằng đồng hồ bấm giây: Đo quãng đường s bằng thước kẻ hoặc thước cuộn; đo khoảng thời gian t đi hết quãng đường đó bằng đồng hồ bấm giây; tính tốc độ theo công thức v = s / t.",
            "b. Đo tốc độ bằng đồng hồ đo thời gian hiện số và cổng quang điện: Khi vật đi qua cổng quang điện thứ nhất, đồng hồ bắt đầu đếm thời gian; khi vật qua cổng quang điện thứ hai, đồng hồ dừng đếm. Thời gian hiển thị chính xác trên mặt đồng hồ.",
            "c. Thiết bị bắn tốc độ: Sử dụng camera và máy phát tia sáng/sóng vô tuyến để ghi lại thời gian ô tô, xe máy chạy qua giữa hai vạch mốc đã định sẵn trên đường, từ đó tự động tính toán tốc độ của phương tiện để kiểm tra xem có vượt quá tốc độ cho phép hay không."
        ]
    },
    {
        "title": "3. Đồ thị quãng đường – thời gian",
        "items": [
            "a. Đồ thị quãng đường – thời gian: Biểu diễn mối liên hệ giữa quãng đường đi được s (trên trục thẳng đứng Os) theo thời gian chuyển động t (trên trục nằm ngang Ot).",
            "b. Đặc điểm đồ thị:\n• Đường biểu diễn là đoạn thẳng dốc lên: Vật chuyển động thẳng đều với tốc độ không đổi.\n• Đường thẳng càng nghiêng dốc nhiều so với trục thời gian: Tốc độ của vật càng lớn.\n• Đường nằm ngang song song với trục thời gian: Quãng đường không đổi theo thời gian, chứng tỏ vật đứng yên (tốc độ v = 0)."
        ]
    }
]

KHTN2_MCQ = [
    {
        "id": 1,
        "content": "Tốc độ cho biết điều gì về chuyển động của vật?",
        "options": [
            "A. Mức độ nhanh hay chậm của chuyển động.",
            "B. Hướng chuyển động của vật.",
            "C. Quãng đường vật đi được.",
            "D. Thời gian chuyển động của vật."
        ],
        "correct": "A",
        "solution": "Tốc độ là đại lượng vật lí đặc trưng cho mức độ chuyển động nhanh hay chậm của vật."
    },
    {
        "id": 2,
        "content": "Công thức nào sau đây dùng để tính tốc độ của một vật?",
        "options": [
            "A. v = s.t",
            "B. v = s/t",
            "C. v = t/s",
            "D. v = s"
        ],
        "correct": "B",
        "solution": "Tốc độ v bằng quãng đường s chia cho thời gian t: v = s / t."
    },
    {
        "id": 3,
        "content": "Đơn vị nào sau đây là đơn vị đo tốc độ trong hệ đo lường chính thức của nước ta?",
        "options": [
            "A. m/s",
            "B. km/min.",
            "C. km/s",
            "D. m/phút"
        ],
        "correct": "A",
        "solution": "Trong hệ thống đo lường hợp pháp (SI) của nước ta, đơn vị chính thức của tốc độ là mét trên giây (m/s) và kilômét trên giờ (km/h)."
    },
    {
        "id": 4,
        "content": "Tốc độ 36 km/h ứng với giá trị nào sau đây?",
        "options": [
            "A. 3,6 m/s",
            "B. 12 m/s",
            "C. 10 m/s",
            "D. 1 m/s"
        ],
        "correct": "C",
        "solution": "Đổi từ km/h sang m/s ta chia cho 3,6: v = 36 / 3,6 = 10 m/s."
    },
    {
        "id": 5,
        "content": "Đơn vị của tốc độ phụ thuộc vào:",
        "options": [
            "A. đơn vị đo chiều dài.",
            "B. đơn vị đo độ dài và đơn vị đo thời gian.",
            "C. đơn vị đo thời gian.",
            "D. các yếu tố khác."
        ],
        "correct": "B",
        "solution": "Vì v = s / t nên đơn vị của tốc độ phụ thuộc đồng thời vào đơn vị đo độ dài (quãng đường) và đơn vị đo thời gian."
    },
    {
        "id": 6,
        "content": "Một người đi bộ quãng đường 100 m trong 20 giây. Tốc độ của người đó là:",
        "options": [
            "A. 5 m/s",
            "B. 20 m/s",
            "C. 100 m/s",
            "D. 0,2 m/s"
        ],
        "correct": "A",
        "solution": "Áp dụng công thức: v = s / t = 100 / 20 = 5 m/s."
    },
    {
        "id": 7,
        "content": "Đồ thị quãng đường – thời gian cho biết:",
        "options": [
            "A. Tốc độ đi được",
            "B. Thời gian đi được",
            "C. Quãng đường đi được",
            "D. Cả tốc độ, thời gian và quãng đường đi được."
        ],
        "correct": "D",
        "solution": "Dựa vào đồ thị quãng đường – thời gian, ta xác định trực tiếp được quãng đường s tại mỗi thời điểm t, từ đó tính được tốc độ chuyển động v = s / t của vật."
    },
    {
        "id": 8,
        "content": "Bạn A chạy quãng đường 60m hết 10 giây. Bạn B chạy quãng đường 100m hết 20 giây. Nhận xét nào sau đây là đúng?",
        "options": [
            "A. Bạn A đi nhanh hơn bạn B.",
            "B. Bạn B đi nhanh hơn bạn A.",
            "C. Hai bạn đi nhanh như nhau.",
            "D. Không thể so sánh được."
        ],
        "correct": "A",
        "solution": "• Tốc độ chạy của bạn A: v_A = 60 / 10 = 6 m/s.\n• Tốc độ chạy của bạn B: v_B = 100 / 20 = 5 m/s.\nVì v_A > v_B (6 m/s > 5 m/s) nên bạn A chạy nhanh hơn bạn B."
    },
    {
        "id": 9,
        "content": "Một ô tô đi 100 m đầu tiên trong 10 giây, sau đó đi tiếp 200 m trong 20 giây. Tốc độ chuyển động của ô tô trên cả quãng đường là:",
        "options": [
            "A. 10 m/s",
            "B. 15 m/s",
            "C. 30 m/s",
            "D. 100 m/s"
        ],
        "correct": "A",
        "solution": "Tốc độ trung bình trên toàn bộ quãng đường:\nv = (s1 + s2) / (t1 + t2) = (100 + 200) / (10 + 20) = 300 / 30 = 10 m/s."
    },
    {
        "id": 10,
        "content": "Trên đồ thị quãng đường – thời gian, đường thẳng càng dốc thể hiện:",
        "options": [
            "A. Chuyển động nhanh hơn.",
            "B. Chuyển động chậm hơn.",
            "C. Vật đứng yên.",
            "D. Vật chuyển động đều."
        ],
        "correct": "A",
        "solution": "Độ dốc của đường biểu diễn trên đồ thị s – t đặc trưng cho tốc độ v = s / t. Đường càng nghiêng dốc lớn thì quãng đường đi được trong cùng một khoảng thời gian càng nhiều, tức là vật chuyển động càng nhanh."
    }
]

KHTN2_TF = [
    {
        "id": 11,
        "content": "Xét tính đúng/sai của các phát biểu sau về tốc độ chuyển động:",
        "sub_items": [
            {
                "id": "a",
                "text": "Tốc độ càng lớn thì vật chuyển động càng chậm.",
                "answer": False,
                "reason": "Tốc độ càng lớn thể hiện vật chuyển động càng nhanh."
            },
            {
                "id": "b",
                "text": "Khi vật đứng yên, tốc độ của vật bằng 0.",
                "answer": True,
                "reason": "Vật đứng yên thì quãng đường đi được s = 0 trong suốt thời gian khảo sát, do đó tốc độ v = 0."
            },
            {
                "id": "c",
                "text": "Công thức tính tốc độ là v = t:s.",
                "answer": False,
                "reason": "Công thức đúng là v = s : t (quãng đường chia thời gian)."
            },
            {
                "id": "d",
                "text": "Đơn vị tốc độ có thể là km/h hoặc m/s.",
                "answer": True,
                "reason": "Đây là hai đơn vị đo tốc độ thông dụng và chuẩn mực nhất."
            }
        ]
    },
    {
        "id": 12,
        "content": "Xét tính đúng/sai của các phát biểu về chuyển động và đồ thị quãng đường – thời gian:",
        "sub_items": [
            {
                "id": "a",
                "text": "Trên cùng một quãng đường, nếu thời gian chuyển động càng ngắn thì chuyển động đó nhanh hơn.",
                "answer": True,
                "reason": "Vì v = s / t, khi s không đổi, t càng nhỏ thì v càng lớn."
            },
            {
                "id": "b",
                "text": "Đồ thị quãng đường – thời gian là đường thẳng đi qua gốc tọa độ.",
                "answer": True,
                "reason": "Đối với vật chuyển động thẳng đều xuất phát từ mốc tọa độ tại thời điểm ban đầu t = 0 (s = 0), đồ thị là đường thẳng đi qua gốc tọa độ O(0,0)."
            },
            {
                "id": "c",
                "text": "Đồ thị càng dốc thì tốc độ càng nhỏ.",
                "answer": False,
                "reason": "Đồ thị càng dốc chứng tỏ tốc độ của vật càng lớn."
            },
            {
                "id": "d",
                "text": "Người đi xe máy 60 km trong 1 giờ có tốc độ 60 km/h.",
                "answer": True,
                "reason": "v = s / t = 60 km / 1 h = 60 km/h."
            }
        ]
    }
]

KHTN2_SHORT = [
    {
        "id": 13,
        "content": "Tốc độ chuyển động được kí hiệu là.......",
        "answer": "v",
        "solution": "Tốc độ trong vật lí được kí hiệu chuẩn quốc tế bằng chữ cái v (viết tắt của từ velocity / vitesse)."
    },
    {
        "id": 14,
        "content": "1 km/h = ..... m/s ; 1 m/s = .... km/h.",
        "answer": "1/3,6 m/s (≈ 0,28 m/s) ; 3,6 km/h",
        "solution": "• 1 km/h = 1000 m / 3600 s = 1/3,6 m/s ≈ 0,28 m/s.\n• 1 m/s = (1/1000 km) / (1/3600 h) = 3600 / 1000 = 3,6 km/h."
    },
    {
        "id": 15,
        "content": "Quãng đường vật đi được đo bằng đơn vị là mét (m), thời gian đo bằng đơn vị là giây (s) thì tốc độ chuyển động có đơn vị là......",
        "answer": "mét trên giây (m/s)",
        "solution": "Đơn vị tốc độ = (đơn vị độ dài) / (đơn vị thời gian) = m/s."
    },
    {
        "id": 16,
        "content": "Nếu một vật đi cùng một quãng đường nhưng thời gian đi hết quãng đường đó tăng gấp đôi thì tốc độ chuyển động của vật giảm đi ...... lần.",
        "answer": "2",
        "solution": "Áp dụng v = s / t: Khi quãng đường s không đổi, tốc độ tỉ lệ nghịch với thời gian t. Do đó thời gian tăng 2 lần thì tốc độ giảm đi 2 lần."
    },
    {
        "id": 17,
        "content": "Trong cùng một khoảng thời gian, nếu quãng đường chuyển động ...... hơn thì chuyển động đó nhanh hơn.",
        "answer": "dài hơn (hoặc lớn hơn)",
        "solution": "Khi thời gian t như nhau, vật đi được quãng đường dài hơn sẽ có tốc độ lớn hơn (chuyển động nhanh hơn)."
    },
    {
        "id": 18,
        "content": "Một người đi xe đạp trong 3 giờ đi được quãng đường 45 km. Tốc độ của người đó là .... km/h.",
        "answer": "15 km/h",
        "solution": "v = s / t = 45 / 3 = 15 km/h."
    },
    {
        "id": 19,
        "content": "Một vận động viên chạy 400 m hết 50 giây. Tốc độ của vận động viên đó là ....... m/s.",
        "answer": "8 m/s",
        "solution": "v = s / t = 400 / 50 = 8 m/s."
    },
    {
        "id": 20,
        "content": "Công thức tính tốc độ là......",
        "answer": "v = s / t",
        "solution": "Công thức chuẩn: v = s / t (hoặc v = s : t)."
    },
    {
        "id": 21,
        "content": "Trong hệ thống đo lường chính thức ở nước ta, tốc độ được đo bằng đơn vị m/s và.....",
        "answer": "km/h (kilômét trên giờ)",
        "solution": "Hai đơn vị chính thức là mét trên giây (m/s) và kilômét trên giờ (km/h)."
    },
    {
        "id": 22,
        "content": "Đại lượng nào đặc trưng cho mức độ nhanh chậm của chuyển động là....",
        "answer": "Tốc độ",
        "solution": "Tốc độ là đại lượng đặc trưng cho mức độ chuyển động nhanh hay chậm của vật."
    },
    {
        "id": 23,
        "content": "Vật đứng yên thì tốc độ bằng ......",
        "answer": "0 (hoặc 0 m/s)",
        "solution": "Vật không chuyển động (đứng yên) thì quãng đường đi được bằng 0 nên tốc độ v = 0."
    },
    {
        "id": 24,
        "content": "Dụng cụ đo tốc độ của các phương tiện giao thông là.....",
        "answer": "Tốc kế (đồng hồ đo tốc độ)",
        "solution": "Tốc kế là dụng cụ được gắn trên bảng điều khiển xe máy, ô tô,... để đo và hiển thị trực tiếp tốc độ tức thời của phương tiện."
    }
]

KHTN2_ESSAY = [
    {
        "id": 25,
        "content": "Một xe tải chạy trên đoạn đường đầu dài 45 km trong 45 phút, sau đó xe tiếp tục chạy thêm 18 km trong 20 phút. Tính tốc độ của xe tải trên mỗi đoạn đường.",
        "solution": "• **Đoạn đường đầu:**\n  - Đổi thời gian: t1 = 45 phút = 45 / 60 h = 0,75 h.\n  - Tốc độ của xe tải trên đoạn đường đầu:\n    v1 = s1 / t1 = 45 / 0,75 = **60 km/h**.\n• **Đoạn đường sau:**\n  - Đổi thời gian: t2 = 20 phút = 20 / 60 h = 1/3 h ≈ 0,333 h.\n  - Tốc độ của xe tải trên đoạn đường thứ hai:\n    v2 = s2 / t2 = 18 / (1/3) = 18 × 3 = **54 km/h**.\n• **Đáp số:** Tốc độ đoạn đầu là 60 km/h; đoạn sau là 54 km/h."
    },
    {
        "id": 26,
        "content": "Một học sinh đi xe đạp từ nhà đến trường mất 20 phút. Đoạn đường từ nhà đến trường dài 4km.\na) Tính tốc độ của chuyển động theo đơn vị km/h?\nb) Ý nghĩa của con số tìm ở câu a?",
        "solution": "a) Đổi đơn vị thời gian: t = 20 phút = 20 / 60 h = 1/3 h.\n   Tốc độ chuyển động của học sinh là:\n   v = s / t = 4 / (1/3) = 4 × 3 = **12 km/h**.\nb) **Ý nghĩa con số 12 km/h:** Con số này cho biết nếu học sinh duy trì tốc độ đạp xe đều đặn như vậy thì cứ trong 1 giờ học sinh sẽ đi được quãng đường dài 12 km (hoặc trung bình mỗi giờ đi được 12 km)."
    },
    {
        "id": 27,
        "content": "Dựa vào bảng ghi số liệu dưới đây về quãng đường và thời gian của một người đi bộ, em hãy vẽ đồ thị quãng đường – thời gian của người này:\n\n| t (h) | 0 | 0,5 | 1,0 | 1,5 | 2,0 |\n| :--- | :---: | :---: | :---: | :---: | :---: |\n| **s (km)** | 0 | 2,5 | 5,0 | 7,5 | 10,0 |",
        "image": "do_thi_quang_duong_thoi_gian.png",
        "solution": "• **Nhận xét chuyển động:** Tốc độ của người đi bộ trên các khoảng thời gian là không đổi:\n  v = 2,5 / 0,5 = 5,0 / 1,0 = 7,5 / 1,5 = 10,0 / 2,0 = 5 km/h. Đây là chuyển động thẳng đều.\n• **Cách vẽ đồ thị:**\n  1. Vẽ hệ trục tọa độ vuông góc: Trục hoành Ot biểu diễn thời gian t (h), trục tung Os biểu diễn quãng đường s (km).\n  2. Xác định các tọa độ điểm: O(0; 0), A(0,5; 2,5), B(1,0; 5,0), C(1,5; 7,5), D(2,0; 10,0).\n  3. Nối các điểm O, A, B, C, D lại với nhau, ta được một đường thẳng dốc lên đi qua gốc tọa độ O. Đó chính là đồ thị quãng đường – thời gian cần vẽ."
    },
    {
        "id": 28,
        "content": "Camera của thiết bị “bắn tốc độ” ghi và tính được thời gian một ô tô chạy qua giữa hai vạch mốc cách nhau 10 m là 0,5 s. Nếu tốc độ giới hạn trên làn đường được quy định là 80 km/h thì ô tô này có vượt quá tốc độ cho phép không?",
        "solution": "• Tốc độ chuyển động của ô tô khi chạy qua giữa hai vạch mốc là:\n  v = s / t = 10 / 0,5 = **20 m/s**.\n• Đổi tốc độ của ô tô sang đơn vị km/h:\n  v = 20 × 3,6 = **72 km/h**.\n• So sánh với tốc độ giới hạn quy định trên làn đường:\n  Vì **72 km/h < 80 km/h**.\n• **Kết luận:** Ô tô **không** chạy vượt quá tốc độ cho phép trên làn đường đó."
    }
]

# ==============================================================================
# PHÂN MÔN KHTN 3: SINH HỌC
# ==============================================================================
KHTN3_THEORY = [
    {
        "title": "I. Khái quát về trao đổi chất và chuyển hoá năng lượng",
        "items": [
            "1. Khái niệm:\n• Trao đổi chất là quá trình cơ thể lấy các chất từ môi trường, biến đổi chúng thành các chất cần thiết cho cơ thể và thải các chất không cần thiết ra môi trường.\n• Chuyển hoá năng lượng là sự biến đổi năng lượng từ dạng này sang dạng khác, ví dụ: quang năng chuyển thành hoá năng; hoá năng chuyển thành cơ năng và nhiệt năng.",
            "2. Mối quan hệ: Trao đổi chất và chuyển hoá năng lượng gắn liền, diễn ra đồng thời trong cơ thể sinh vật.",
            "3. Vai trò: Cung cấp vật chất và năng lượng cho các hoạt động sống, giúp sinh vật duy trì sự sống, sinh trưởng, phát triển, cảm ứng, vận động và sinh sản."
        ]
    },
    {
        "title": "II. Quang hợp ở thực vật",
        "items": [
            "1. Khái niệm: Quang hợp là quá trình thực vật sử dụng nước và khí carbon dioxide, nhờ năng lượng ánh sáng được diệp lục hấp thụ, để tổng hợp chất hữu cơ và giải phóng khí oxygen.",
            "2. Phương trình tổng quát dạng chữ:\n   Nước + carbon dioxide  ──(Ánh sáng / Diệp lục)──>  chất hữu cơ + oxygen\n   Glucose được tạo ra có thể được sử dụng để tổng hợp tinh bột và các chất hữu cơ khác.",
            "3. Trao đổi chất và chuyển hoá năng lượng trong quang hợp:\n• Trao đổi chất: Cây thu nhận nước và carbon dioxide; tạo chất hữu cơ và giải phóng oxygen ra môi trường.\n• Chuyển hoá năng lượng: Quang năng chuyển thành hoá năng, được tích luỹ trong các chất hữu cơ.",
            "4. Đặc điểm cấu tạo của lá phù hợp với chức năng quang hợp:\n• Phiến lá rộng, dẹt giúp hấp thụ được nhiều ánh sáng.\n• Gân lá có mạch gỗ vận chuyển nước, muối khoáng đến lá và mạch rây vận chuyển chất hữu cơ từ lá đến các cơ quan khác.\n• Khí khổng giúp trao đổi khí giữa lá với môi trường và thoát hơi nước; khí khổng thường tập trung nhiều ở mặt dưới lá.\n• Tế bào lá có lục lạp; trong lục lạp có diệp lục hấp thụ năng lượng ánh sáng.",
            "5. Các yếu tố ảnh hưởng đến quang hợp:\n• Ánh sáng: Cường độ và thời gian chiếu sáng ảnh hưởng đến tốc độ quang hợp.\n• Nồng độ carbon dioxide (CO2): CO2 là nguyên liệu của quang hợp. Khi các điều kiện khác thích hợp, tăng nồng độ CO2 có thể làm tăng cường độ quang hợp đến một giới hạn nhất định.\n• Nhiệt độ: Ảnh hưởng đến hoạt động của các enzyme tham gia quang hợp; mỗi loài cây có khoảng nhiệt độ thích hợp.\n• Nước: Là nguyên liệu của quang hợp; thiếu nước còn có thể làm khí khổng đóng, hạn chế CO2 đi vào lá."
        ]
    },
    {
        "title": "III. Hô hấp tế bào ở sinh vật",
        "items": [
            "1. Khái niệm: Hô hấp tế bào là quá trình phân giải chất hữu cơ, chủ yếu là glucose, thành carbon dioxide và nước, đồng thời giải phóng năng lượng; một phần năng lượng được tích luỹ trong ATP để cung cấp cho hoạt động sống.",
            "2. Phương trình tổng quát dạng chữ:\n   Glucose + oxygen  ──>  carbon dioxide + nước + năng lượng (ATP và nhiệt)",
            "3. Nơi diễn ra: Diễn ra trong tế bào, bào quan thực hiện chính là bào quan ti thể.",
            "4. Mối quan hệ giữa quang hợp và hô hấp tế bào:\n• Quang hợp tổng hợp chất hữu cơ, tích luỹ năng lượng và giải phóng oxygen.\n• Hô hấp tế bào phân giải chất hữu cơ, giải phóng năng lượng cho hoạt động sống và tạo ra carbon dioxide, nước là những chất có thể được sử dụng làm nguyên liệu cho quang hợp.",
            "5. Các yếu tố ảnh hưởng đến hô hấp tế bào:\n• Nhiệt độ: Ảnh hưởng đến hoạt động của enzyme hô hấp; nhiệt độ quá thấp hoặc quá cao đều có thể làm giảm cường độ hô hấp.\n• Hàm lượng nước: Với hạt, khi hàm lượng nước tăng trong điều kiện thích hợp, cường độ hô hấp thường tăng.\n• Nồng độ oxygen (O2): O2 là nguyên liệu của hô hấp hiếu khí; khi nồng độ O2 trong môi trường xuống dưới khoảng 5%, cường độ hô hấp của thực vật có thể giảm.\n• Nồng độ carbon dioxide (CO2): Nồng độ CO2 trong môi trường từ khoảng 3% đến 5% đã có thể gây ức chế hô hấp tế bào ở thực vật."
        ]
    },
    {
        "title": "IV. Thực hành",
        "items": [
            "Ôn lại mục đích, dụng cụ, cách tiến hành, hiện tượng và kết luận của các thí nghiệm trong Bài 24 – Thực hành: Chứng minh quang hợp ở cây xanh và Bài 27 – Thực hành: Hô hấp ở thực vật."
        ]
    }
]

KHTN3_QUESTIONS = [
    # I. MỨC ĐỘ NHẬN BIẾT
    {
        "id": 1,
        "level": "I. Mức độ nhận biết",
        "content": "Phát biểu khái niệm trao đổi chất và chuyển hoá năng lượng ở sinh vật.",
        "solution": "• **Trao đổi chất:** Là quá trình cơ thể lấy các chất từ môi trường ngoài, biến đổi chúng thành các chất cần thiết cho cơ thể và thải các chất không cần thiết ra môi trường.\n• **Chuyển hoá năng lượng:** Là sự biến đổi năng lượng từ dạng này sang dạng khác (như từ quang năng thành hoá năng, từ hoá năng thành cơ năng và nhiệt năng)."
    },
    {
        "id": 2,
        "level": "I. Mức độ nhận biết",
        "content": "Nêu vai trò của trao đổi chất và chuyển hoá năng lượng đối với sinh vật.",
        "solution": "Cung cấp vật chất và năng lượng cần thiết cho toàn bộ các hoạt động sống của tế bào và cơ thể, giúp sinh vật duy trì sự sống, lớn lên, sinh trưởng, phát triển, cảm ứng, vận động và sinh sản."
    },
    {
        "id": 3,
        "level": "I. Mức độ nhận biết",
        "content": "Quang hợp là gì? Viết phương trình tổng quát của quá trình quang hợp ở thực vật dạng chữ.",
        "solution": "• **Khái niệm:** Quang hợp là quá trình thực vật sử dụng nước và khí carbon dioxide, nhờ năng lượng ánh sáng do diệp lục hấp thụ, để tổng hợp chất hữu cơ (glucose, tinh bột) và giải phóng khí oxygen.\n• **Phương trình dạng chữ:**\n  Nước + carbon dioxide  ──(Ánh sáng / Diệp lục)──>  chất hữu cơ + oxygen"
    },
    {
        "id": 4,
        "level": "I. Mức độ nhận biết",
        "content": "Hô hấp tế bào là gì? Viết phương trình tổng quát dạng chữ.",
        "solution": "• **Khái niệm:** Hô hấp tế bào là quá trình phân giải chất hữu cơ (chủ yếu là glucose) thành carbon dioxide và nước, đồng thời giải phóng năng lượng cung cấp cho các hoạt động sống.\n• **Phương trình dạng chữ:**\n  Glucose + oxygen  ──>  carbon dioxide + nước + năng lượng (ATP và nhiệt)"
    },
    {
        "id": 5,
        "level": "I. Mức độ nhận biết",
        "content": "Nêu tên bào quan thực hiện chức năng quang hợp ở lá cây và bào quan giữ vai trò chủ yếu trong hô hấp tế bào.",
        "solution": "• Bào quan thực hiện chức năng quang hợp: **Lục lạp** (chứa chất diệp lục).\n• Bào quan giữ vai trò chủ yếu trong hô hấp tế bào: **Ti thể** (nhà máy năng lượng của tế bào)."
    },
    {
        "id": 6,
        "level": "I. Mức độ nhận biết",
        "content": "Kể tên các bộ phận của lá cây phù hợp với chức năng quang hợp.",
        "solution": "Các bộ phận chính gồm:\n1. **Phiến lá:** Rộng, dẹt giúp thu nhận được tối đa ánh sáng mặt trời.\n2. **Gân lá:** Chứa mạch gỗ và mạch rây giúp vận chuyển nước, muối khoáng đến và vận chuyển chất hữu cơ đi.\n3. **Khí khổng:** Nằm nhiều ở biểu bì mặt dưới lá, giúp trao đổi khí (CO2, O2) và thoát hơi nước.\n4. **Tế bào thịt lá:** Chứa nhiều lục lạp với hệ sắc tố diệp lục hấp thụ ánh sáng."
    },
    {
        "id": 7,
        "level": "I. Mức độ nhận biết",
        "content": "Nêu vai trò của khí khổng đối với quang hợp và trao đổi khí ở thực vật.",
        "solution": "• Giúp thực hiện quá trình trao đổi khí: Lấy khí CO2 từ môi trường vào phục vụ cho quang hợp và giải phóng khí O2 ra môi trường.\n• Điều hòa quá trình thoát hơi nước: Giúp hạ nhiệt độ của lá, làm lá không bị đốt nóng dưới nắng gắt, đồng thời tạo lực hút kéo dòng nước và muối khoáng từ rễ lên lá."
    },
    {
        "id": 8,
        "level": "I. Mức độ nhận biết",
        "content": "Nêu các yếu tố môi trường chủ yếu ảnh hưởng đến quang hợp ở cây xanh.",
        "solution": "Các yếu tố môi trường chủ yếu gồm:\n1. Ánh sáng (cường độ, thời gian chiếu sáng và quang phổ ánh sáng).\n2. Nồng độ khí carbon dioxide (CO2).\n3. Nhiệt độ môi trường.\n4. Nước và độ ẩm không khí."
    },
    {
        "id": 9,
        "level": "I. Mức độ nhận biết",
        "content": "Nêu các yếu tố môi trường ảnh hưởng đến hô hấp tế bào.",
        "solution": "Các yếu tố môi trường ảnh hưởng đến hô hấp tế bào gồm:\n1. Nhiệt độ.\n2. Hàm lượng nước và độ ẩm.\n3. Nồng độ khí oxygen (O2).\n4. Nồng độ khí carbon dioxide (CO2)."
    },

    # II. MỨC ĐỘ THÔNG HIỂU
    {
        "id": 10,
        "level": "II. Mức độ thông hiểu",
        "content": "Phân tích mối quan hệ giữa trao đổi chất và chuyển hoá năng lượng trong cơ thể sinh vật.",
        "solution": "Trao đổi chất và chuyển hoá năng lượng là hai quá trình gắn liền và diễn ra đồng thời, không thể tách rời:\n• Trao đổi chất cung cấp nguyên liệu, chất hữu cơ cho quá trình chuyển hoá năng lượng.\n• Ngược lại, chuyển hoá năng lượng tạo ra năng lượng (dưới dạng ATP) để cung cấp cho hoạt động trao đổi chất, vận chuyển các chất và tổng hợp nên chất sống mới của cơ thể sinh vật."
    },
    {
        "id": 11,
        "level": "II. Mức độ thông hiểu",
        "content": "Những lá cây có màu đỏ hoặc vàng (ví dụ: cây huyết dụ, rau dền đỏ) có khả năng quang hợp không? Vì sao?",
        "solution": "• **Trả lời:** Những lá cây có màu đỏ hoặc vàng vẫn có khả năng quang hợp bình thường.\n• **Giải thích:** Trong lá của những loài cây này ngoài các sắc tố màu đỏ, vàng, cam (như anthocyanin, carotenoid) thì **vẫn có chứa chất diệp lục**. Tuy nhiên do hàm lượng các sắc tố màu đỏ, vàng quá lớn nên đã che khuất màu xanh của diệp lục. Diệp lục trong lá vẫn hấp thụ năng lượng ánh sáng để thực hiện quá trình quang hợp bình thường."
    },
    {
        "id": 12,
        "level": "II. Mức độ thông hiểu",
        "content": "Vì sao nói quang hợp và hô hấp tế bào có mối quan hệ với nhau trong hoạt động sống của cây?",
        "solution": "Quang hợp và hô hấp là hai quá trình trái ngược nhau nhưng hỗ trợ và phụ thuộc chặt chẽ vào nhau:\n• **Quang hợp:** Sử dụng CO2 và H2O để tổng hợp nên chất hữu cơ và giải phóng O2, đồng thời tích lũy năng lượng dưới dạng hoá năng.\n• **Hô hấp:** Sử dụng chính chất hữu cơ và O2 do quang hợp tạo ra để phân giải, giải phóng năng lượng (ATP) nuôi sống cây, đồng thời tạo ra sản phẩm CO2 và H2O làm nguyên liệu cung cấp lại cho quang hợp.\n• Cây không thể sống nếu thiếu một trong hai quá trình này."
    },
    {
        "id": 13,
        "level": "II. Mức độ thông hiểu",
        "content": "Vì sao trong trồng trọt, người ta thường trồng xen cây ưa sáng và cây ưa bóng?",
        "solution": "• Các loài cây có nhu cầu ánh sáng khác nhau: Cây ưa sáng cần nhiều ánh sáng mạnh để quang hợp, cây ưa bóng cần ít ánh sáng hơn và có thể phát triển dưới bóng râm.\n• Việc trồng xen cây ưa sáng (tầng trên) và cây ưa bóng (tầng dưới) giúp **tận dụng tối đa không gian dinh dưỡng, diện tích đất canh tác và nguồn ánh sáng mặt trời** ở nhiều tầng tán khác nhau, từ đó nâng cao năng suất cây trồng và hiệu quả kinh tế trên một diện tích đất canh tác."
    },
    {
        "id": 14,
        "level": "II. Mức độ thông hiểu",
        "content": "Vì sao khi làm việc nặng hoặc vận động mạnh, nhịp thở và nhịp tim tăng, cơ thể nóng lên và toát mồ hôi?",
        "solution": "• Khi vận động mạnh, các tế bào cơ co giãn liên tục đòi hỏi tiêu tốn rất nhiều năng lượng ATP, do đó **quá trình hô hấp tế bào phải tăng tốc độ**.\n• Để đáp ứng, cơ thể phải **tăng nhịp thở** để lấy thêm nhiều O2 và thải nhanh khí CO2; đồng thời **tăng nhịp tim** để đẩy nhanh máu vận chuyển O2 và chất dinh dưỡng đến các tế bào cơ.\n• Quá trình hô hấp tế bào ngoài tạo ra ATP còn giải phóng một lượng nhiệt lớn làm **cơ thể nóng lên**.\n• Cơ thể phản ứng bằng cách giãn mạch máu dưới da và **tiết mồ hôi**, mồ hôi bay hơi mang theo nhiệt lượng giúp làm mát và hạ nhiệt độ cơ thể."
    },
    {
        "id": 15,
        "level": "II. Mức độ thông hiểu",
        "content": "Vì sao không nên để quá nhiều hoa tươi hoặc cây xanh trong phòng ngủ đóng kín cửa vào ban đêm?",
        "solution": "• Vào ban đêm, khi không có ánh sáng mặt trời, cây xanh và hoa ngừng quang hợp mà chỉ diễn ra quá trình **hô hấp tế bào**.\n• Trong phòng ngủ đóng kín cửa, cây xanh và hoa sẽ liên tục hấp thu khí **oxygen (O2)** và thải ra khí **carbon dioxide (CO2)**.\n• Điều này làm cho không khí trong phòng nhanh chóng bị thiếu hụt O2 và tích tụ nhiều khí CO2, khiến người ngủ trong phòng cảm thấy ngột ngạt, khó thở, mệt mỏi, đau đầu và trong trường hợp nặng có thể dẫn đến nguy cơ ngạt thở tử vong."
    },
    {
        "id": 16,
        "level": "II. Mức độ thông hiểu",
        "content": "Vì sao khi nuôi cá cảnh trong bể kính, người ta thường thả vào bể một số cành rong và cây thuỷ sinh?",
        "solution": "• **Cung cấp oxygen cho cá:** Ban ngày (hoặc khi bật đèn chiếu sáng bể cá), rong và cây thủy sinh tiến hành quang hợp, hấp thụ khí CO2 do cá thở ra và giải phóng khí O2 hòa tan trực tiếp vào nước, giúp cá hô hấp thuận lợi và khỏe mạnh.\n• **Làm sạch môi trường nước:** Rễ và thân cây thủy sinh có khả năng hấp thụ các chất hữu cơ bài tiết, chất thải thừa của cá, giúp nước bể trong sạch hơn.\n• **Tạo cảnh quan sinh thái:** Tạo không gian sống sinh thái tự nhiên, tạo nơi ẩn nấp, vui chơi và đẻ trứng an toàn cho cá cảnh."
    },

    # III. MỨC ĐỘ VẬN DỤNG VÀ VẬN DỤNG CAO
    {
        "id": 17,
        "level": "III. Mức độ vận dụng và vận dụng cao",
        "content": "Giải thích vì sao ở khu dân cư và nhà máy, người ta thường trồng nhiều cây xanh.",
        "solution": "• **Lọc không khí:** Nhờ quá trình quang hợp, cây xanh hấp thụ một lượng lớn khí CO2 và nhả ra khí O2, giúp cân bằng khí quyển.\n• **Giảm khói bụi và tiếng ồn:** Tán lá cây có tác dụng cản bụi, hấp thụ một số khí độc (SO2, NO2) từ xe cộ và nhà máy, đồng thời làm giảm tiếng ồn đô thị.\n• **Điều hòa vi khí hậu:** Quá trình thoát hơi nước của lá cây giúp tăng độ ẩm không khí, hạ nhiệt độ môi trường xung quanh, tạo bóng mát làm dịu mát khu dân cư và khuôn viên nhà máy."
    },
    {
        "id": 18,
        "level": "III. Mức độ vận dụng và vận dụng cao",
        "content": "Một bạn cho rằng: “Ban đêm cây không quang hợp nên cây cũng không hô hấp.” Em đồng ý hay không? Giải thích.",
        "solution": "• **Ý kiến:** Em **hoàn toàn không đồng ý** với nhận định trên.\n• **Giải thích:**\n  - Quang hợp chỉ diễn ra khi có ánh sáng nên ban đêm quá trình này tạm dừng.\n  - Tuy nhiên, **hô hấp tế bào là quá trình sống còn, diễn ra liên tục suốt cả ngày lẫn đêm** ở mọi bộ phận sống của cây để liên tục phân giải chất hữu cơ tạo ra năng lượng ATP duy trì các hoạt động sống của tế bào. Nếu cây ngừng hô hấp, cây sẽ không có năng lượng và bị chết ngay lập tức."
    },
    {
        "id": 19,
        "level": "III. Mức độ vận dụng và vận dụng cao",
        "content": "Vì sao trước khi gieo hạt giống, người nông dân thường ngâm hạt trong nước ấm và ủ hạt?",
        "solution": "• **Cung cấp nước:** Hạt khô đang ở trạng thái ngủ nghỉ có cường độ hô hấp rất yếu. Ngâm nước giúp hạt no nước, kích hoạt các enzyme trao đổi chất hoạt động trở lại.\n• **Nhiệt độ thích hợp:** Nước ấm cung cấp nhiệt độ thuận lợi, thúc đẩy tốc độ hô hấp tế bào tăng cao, phân giải chất dự trữ trong hạt thành năng lượng và chất dinh dưỡng để phôi mầm phát triển nhanh.\n• **Tăng tỉ lệ nảy mầm:** Ủ ấm giúp hạt nảy mầm đồng loạt, cây con mọc nhanh, khỏe mạnh và hạn chế sự tấn công của mầm bệnh trong đất."
    },
    {
        "id": 20,
        "level": "III. Mức độ vận dụng và vận dụng cao",
        "content": "Vì sao vào ngày nắng nóng, đứng dưới bóng cây thường mát hơn đứng dưới mái che bằng tôn hoặc bạt?",
        "solution": "• **Hiện tượng thoát hơi nước của cây:** Lá cây liên tục diễn ra quá trình thoát hơi nước qua khí khổng. Khi nước bay hơi sẽ thu nhiệt lượng từ môi trường xung quanh, làm nhiệt độ vùng không khí dưới tán cây giảm xuống đáng kể (thường mát hơn nhiệt độ ngoài trời từ 2 – 4 °C).\n• **Đặc tính truyền nhiệt:** Tán lá cây hấp thụ phần lớn năng lượng ánh sáng để quang hợp, chỉ một phần nhỏ nhiệt truyền xuống dưới. Ngược lại, mái tôn hoặc mái bạt hấp thụ nhiệt rất mạnh rồi bức xạ nhiệt xuống phía dưới khiến không gian dưới mái che trở nên hầm hập, oi bức."
    },
    {
        "id": 21,
        "level": "III. Mức độ vận dụng và vận dụng cao",
        "content": "Ngập úng lâu ngày ảnh hưởng như thế nào đến rễ cây và khả năng quang hợp của lá?",
        "solution": "• **Ảnh hưởng đến rễ:** Khi đất bị ngập nước lâu ngày, nước chiếm hết các khe hở chứa không khí trong đất, dẫn đến rễ cây bị thiếu hụt trầm trọng khí oxygen. Thiếu O2, rễ không thể hô hấp hiếu khí, phải chuyển sang lên men tạo ra các chất độc (rượu, acid) làm tế bào lông hút bị chết, rễ bị thối hỏng và mất khả năng hút nước, khoáng.\n• **Ảnh hưởng đến quang hợp của lá:** Do rễ không hút được nước và muối khoáng chuyển lên lá, cây bị thiếu nước nghiêm trọng dẫn đến khí khổng đóng lại. Khi khí khổng đóng, CO2 không thể khuếch tán vào thịt lá, làm quá trình quang hợp của lá bị đình trệ, lá vàng héo và rụng dần, cuối cùng cây bị chết."
    },
    {
        "id": 22,
        "level": "III. Mức độ vận dụng và vận dụng cao",
        "content": "Biến đổi khí hậu (nhiệt độ tăng, hạn hán kéo dài) ảnh hưởng thế nào đến trao đổi chất của vật nuôi, cây trồng ở địa phương? Đề xuất biện pháp khắc phục.",
        "solution": "• **Ảnh hưởng:**\n  - *Cây trồng:* Nắng nóng, thiếu nước làm khí khổng đóng, quang hợp giảm mạnh, cây bị héo rũ, năng suất giảm sút, thậm chí chết khô.\n  - *Vật nuôi:* Nhiệt độ cao gây stress nhiệt, vật nuôi ăn uống kém, giảm sức đề kháng, rối loạn trao đổi chất, chậm lớn và dễ bùng phát dịch bệnh.\n• **Biện pháp khắc phục:**\n  - Áp dụng các hệ thống tưới nước tiết kiệm (tưới nhỏ giọt, tưới phun sương), phủ gốc giữ ẩm đất.\n  - Lựa chọn các giống cây trồng, vật nuôi có khả năng chịu hạn, chịu nhiệt tốt thích ứng với địa phương.\n  - Trồng cây che bóng, làm giàn lưới che nắng, lắp quạt thông gió và hệ thống làm mát chuồng trại chăn nuôi.\n  - Cung cấp đầy đủ thức ăn dinh dưỡng giàu năng lượng và nước uống sạch cho vật nuôi."
    },
    {
        "id": 23,
        "level": "III. Mức độ vận dụng và vận dụng cao",
        "content": "Vì sao rau củ sau thu hoạch để ở nhiệt độ phòng thường nhanh héo, mất ngọt và hư hỏng hơn so với rau củ bảo quản trong tủ lạnh?",
        "solution": "• Ở nhiệt độ phòng (khoảng 25 – 35 °C), các phản ứng sinh hóa của tế bào diễn ra mạnh mẽ: Quá trình **hô hấp tế bào diễn ra với cường độ cao** sẽ phân giải nhanh lượng đường và chất hữu cơ dự trữ trong rau củ làm rau mất đi vị ngọt; đồng thời quá trình **thoát hơi nước diễn ra nhanh** làm rau củ mau bị mất nước, khô héo.\n• Nhiệt độ phòng cũng tạo điều kiện tối ưu cho các loại vi khuẩn, nấm mốc phát triển làm thối rữa nông sản.\n• Ngược lại, **trong tủ lạnh (nhiệt độ thấp từ 4 – 8 °C)**, hoạt động của các enzyme hô hấp và vi sinh vật gây hại bị ức chế, cường độ hô hấp và thoát hơi nước giảm xuống mức tối thiểu, giúp bảo tồn chất dinh dưỡng, độ tươi ngọt và kéo dài thời gian bảo quản rau củ."
    }
]

KHTN3_NOTES = [
    "1. Học sinh cần nắm vững khái niệm, phương trình tổng quát dạng chữ, vai trò và mối quan hệ biện chứng giữa quá trình quang hợp và hô hấp tế bào.",
    "2. Khi trả lời các câu hỏi vận dụng thực tế, cần phân tích rõ bản chất sinh học dựa trên kiến thức trao đổi chất, chuyển hoá năng lượng, quang hợp và hô hấp; tránh việc học thuộc lòng máy móc.",
    "3. Phần gợi ý trả lời trong đề cương mang tính định hướng sư phạm. Học sinh hoàn toàn có thể diễn đạt theo cách hiểu của mình miễn là đảm bảo chuẩn xác về mặt khoa học."
]
