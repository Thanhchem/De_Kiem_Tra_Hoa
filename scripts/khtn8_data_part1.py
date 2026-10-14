# -*- coding: utf-8 -*-
"""
Dữ liệu Phần I (Câu 1 đến Câu 90) - Đề cương ôn tập KHTN 8 - THCS Nguyễn Thị Minh Khai
"""

PART1_QUESTIONS = [
    # Page 1
    {
        "id": 1,
        "question": "Câu 1. Trong phòng thực hành KHTN 8, dụng cụ nào sau đây được sử dụng phổ biến để đo thể tích của một lượng chất lỏng?",
        "options": [
            "A. Cân điện tử.",
            "*B. Ống đong (bình chia độ).* ",
            "C. Nhiệt kế y tế.",
            "D. Lực kế lò xo."
        ],
        "correct": "B",
        "solution": "Ống đong (bình chia độ) là dụng cụ chuyên dụng và phổ biến nhất để đo thể tích chất lỏng trong phòng thực hành."
    },
    {
        "id": 2,
        "question": "Câu 2. Thiết bị điện nào trong phòng thí nghiệm KHTN 8 dùng để đo cường độ dòng điện chạy qua đoạn mạch kín?",
        "options": [
            "A. Vôn kế.",
            "B. Biến trở con chạy.",
            "*C. Ampe kế.* ",
            "D. Joulemeter."
        ],
        "correct": "C",
        "solution": "Ampe kế là thiết bị chuyên dùng để đo cường độ dòng điện chạy qua đoạn mạch."
    },
    {
        "id": 3,
        "question": "Câu 3. Biến áp nguồn trong phòng thực hành KHTN 8 có chức năng quan trọng nào sau đây?",
        "options": [
            "A. Đo công suất dòng điện và điện năng tiêu thụ.",
            "*B. Chuyển đổi dòng điện xoay chiều 220V thành nguồn hạ thế an toàn (từ 3V đến 24V).* ",
            "C. Phát tín hiệu âm thanh báo động khi bị chập mạch.",
            "D. Tăng điện áp nguồn lên trên 380V để làm thí nghiệm."
        ],
        "correct": "B",
        "solution": "Biến áp nguồn dùng để biến đổi dòng điện xoay chiều điện áp cao (220V) thành dòng điện xoay chiều hoặc một chiều ở mức hạ thế an toàn (3V - 24V) phục vụ thực hành."
    },
    {
        "id": 4,
        "question": "Câu 4. Để đảm bảo an toàn tính mạng tuyệt đối cho học sinh khi tiến hành thực hành lắp ráp mạch điện, quy tắc an toàn điện quy định nguồn điện sử dụng không được vượt quá giá trị nào?",
        "options": [
            "*A. Dưới 24V.* ",
            "B. 110V.",
            "C. 220V.",
            "D. 380V."
        ],
        "correct": "A",
        "solution": "Trong phòng thực hành trường học, chỉ được sử dụng nguồn điện hạ thế an toàn dưới 24V (hoặc không vượt quá 40V) để đảm bảo an toàn tuyệt đối."
    },
    {
        "id": 5,
        "question": "Câu 5. Khi tiến hành làm thí nghiệm đun nóng chất lỏng trong ống nghiệm, thao tác nào sau đây là đúng quy tắc an toàn?",
        "options": [
            "A. Kẹp ống nghiệm ở vị trí sát đáy ống nghiệm và đun tập trung tại một điểm.",
            "B. Hướng miệng ống nghiệm về phía có người để dễ quan sát hiện tượng.",
            "*C. Kẹp ống nghiệm ở vị trí 1/3 từ miệng ống, làm nóng đều toàn bộ ống nghiệm rồi mới đun trực tiếp tại vị trí có chất lỏng, miệng ống nghiêng về phía không có người.* ",
            "D. Đặt đáy ống nghiệm sát vào bấc đèn cồn đang cháy."
        ],
        "correct": "C",
        "solution": "Kẹp ở 1/3 từ miệng ống, hơ nóng đều ống nghiệm tránh nứt vỡ do sốc nhiệt, đun ở vị trí có hóa chất và hướng miệng ống về nơi không có người."
    },
    {
        "id": 6,
        "question": "Câu 6. Khối lượng của một đơn vị thể tích của một chất được gọi là gì?",
        "options": [
            "A. Trọng lượng.",
            "B. Áp suất.",
            "*C. Khối lượng riêng.* ",
            "D. Moment lực."
        ],
        "correct": "C",
        "solution": "Khối lượng của một đơn vị thể tích một chất được định nghĩa là khối lượng riêng của chất đó: D = m / V."
    },
    {
        "id": 7,
        "question": "Câu 7. Trong hệ đo lường chuẩn, đơn vị đo của khối lượng riêng là gì?",
        "options": [
            "A. Niutơn trên mét vuông (N/m^2^).",
            "*B. Kilôgam trên mét khối (kg/m^3^).* ",
            "C. Gam trên mét (g/m).",
            "D. Niutơn trên mét khối (N/m^3^)."
        ],
        "correct": "B",
        "solution": "Trong hệ SI, đơn vị hợp pháp của khối lượng riêng là kilôgam trên mét khối (kg/m^3^)."
    },
    {
        "id": 8,
        "question": "Câu 8. Khi thực hành xác định khối lượng riêng của một hòn sỏi không thấm nước bằng cân điện tử và bình chia độ, học sinh thu được số liệu: khối lượng sỏi m = 27 g; thể tích nước ban đầu V~1~ = 50 cm^3^; thể tích nước sau khi thả ngập sỏi V~2~ = 60 cm^3^. Khối lượng riêng của sỏi là bao nhiêu?",
        "options": [
            "*A. 2,7 g/cm^3^ (hay 2700 kg/m^3^).* ",
            "B. 0,27 g/cm^3^.",
            "C. 270 kg/m^3^.",
            "D. 27 g/cm^3^."
        ],
        "correct": "A",
        "solution": "Thể tích hòn sỏi: V = V~2~ - V~1~ = 60 - 50 = 10 cm^3^. Khối lượng riêng D = m / V = 27 / 10 = 2,7 g/cm^3^ = 2700 kg/m^3^."
    },
    {
        "id": 9,
        "question": "Câu 9. Một học sinh nhập số liệu thực hành xác định khối lượng riêng của một khối gỗ hình hộp chữ nhật vào phần mềm bảng tính. Kích thước đo được: chiều dài a = 10 cm, chiều rộng b = 5 cm, chiều cao c = 2 cm; khối lượng m = 80 g. Công thức thuật toán trên bảng tính cho ra kết quả khối lượng riêng D là bao nhiêu?",
        "options": [
            "*A. 800 kg/m^3^.* ",
            "B. 0,8 kg/m^3^.",
            "C. 80 kg/m^3^.",
            "D. 8000 kg/m^3^."
        ],
        "correct": "A",
        "solution": "Thể tích: V = a . b . c = 10 . 5 . 2 = 100 cm^3^ = 0,0001 m^3^. Khối lượng: m = 80 g = 0,08 kg. D = m / V = 0,08 / 0,0001 = 800 kg/m^3^."
    },
    {
        "id": 10,
        "question": "Câu 10. Khi dùng phím chức năng đổi đơn vị chuẩn trên tệp dữ liệu số, giá trị khối lượng riêng D = 1 g/cm^3^ của nước nguyên chất tương ứng với giá trị nào sau đây?",
        "options": [
            "A. 10 kg/m^3^.",
            "B. 100 kg/m^3^.",
            "*C. 1000 kg/m^3^.* ",
            "D. 10000 kg/m^3^."
        ],
        "correct": "C",
        "solution": "Ta có 1 g/cm^3^ = 10^-3^ kg / 10^-6^ m^3^ = 1000 kg/m^3^."
    },
    {
        "id": 11,
        "question": "Câu 11. Thiết bị nào sau đây trong phòng thực hành KHTN 8 có chức năng chuyển đổi dòng điện xoay chiều thành dòng điện một chiều với các giá trị điện áp thích hợp?",
        "options": [
            "A. Joulemeter.",
            "*B. Biến áp nguồn.* ",
            "C. Ampe kế.",
            "D. Vôn kế."
        ],
        "correct": "B",
        "solution": "Biến áp nguồn trong phòng thực hành vừa có khả năng hạ áp vừa có bộ chỉnh lưu biến đổi dòng xoay chiều (AC) thành một chiều (DC)."
    },
    {
        "id": 12,
        "question": "Câu 12. Một thùng gỗ hình hộp chữ nhật có kích thước a = 0,2m, b = 0,3m, c = 0,5m và khối lượng 30kg. Để áp suất của thùng gỗ tác dụng lên nền nhà là nhỏ nhất, người ta phải đặt mặt đáy nào tiếp xúc với nền nhà?",
        "options": [
            "A. Mặt đáy có kích thước 0,2m x 0,3m.",
            "B. Mặt đáy có kích thước 0,2m x 0,5m.",
            "*C. Mặt đáy có kích thước 0,3m x 0,5m.* ",
            "D. Bất kì mặt nào vì áp suất không phụ thuộc diện tích tiếp xúc."
        ],
        "correct": "C",
        "solution": "Áp suất p = F / S. Với cùng một áp lực F = P = 10m = 300 N, để áp suất p nhỏ nhất thì diện tích tiếp xúc S phải lớn nhất. S_max = 0,3m x 0,5m = 0,15 m^2^."
    },

    # Page 2
    {
        "id": 13,
        "question": "Câu 13. Tác dụng làm quay của một lực lên một vật quanh một trục cố định được đặc trưng bởi đại lượng nào sau đây?",
        "options": [
            "A. Áp suất.",
            "B. Trọng lượng riêng.",
            "*C. Moment lực.* ",
            "D. Quán tính."
        ],
        "correct": "C",
        "solution": "Moment lực là đại lượng đặc trưng cho tác dụng làm quay của lực quanh một trục hoặc một điểm cố định."
    },
    {
        "id": 14,
        "question": "Câu 14. Dụng cụ nào sau đây hoạt động dựa trên nguyên tắc của đòn bẩy loại 1 (điểm tựa nằm giữa điểm đặt lực tác dụng và điểm đặt tải trọng)?",
        "options": [
            "*A. Kéo cắt giấy.* ",
            "B. Xe cút kít (xe rùa).",
            "C. Kẹp gắp thức ăn.",
            "D. Dụng cụ bấm móng tay."
        ],
        "correct": "A",
        "solution": "Kéo cắt giấy có điểm tựa (trục quay) nằm giữa tay cầm (lực tác dụng) và lưỡi kéo (điểm đặt vật/tải trọng), thuộc đòn bẩy loại 1."
    },
    {
        "id": 15,
        "question": "Câu 15. Một thanh kim loại có thể tích 0,002m^3^ và khối lượng 15,6kg. Khối lượng riêng của kim loại đó là bao nhiêu?",
        "options": [
            "A. 780kg/m^3^.",
            "B. 31200kg/m^3^.",
            "*C. 7800kg/m^3^.* ",
            "D. 3900kg/m^3^"
        ],
        "correct": "C",
        "solution": "D = m / V = 15,6 / 0,002 = 7800 kg/m^3^ (đây là khối lượng riêng của sắt/thép)."
    },
    {
        "id": 16,
        "question": "Câu 16. Một chiếc xe tải có khối lượng tổng cộng là 6 tấn đang đỗ trên mặt đường phẳng. Biết diện tích tiếp xúc của các lốp xe với mặt đường là 0,15m^2^. Lấy g = 10m/s^2^ (P = 10m). Áp suất xe tác dụng lên mặt đường là bao nhiêu?",
        "options": [
            "*A. 400 000 Pa.* ",
            "B. 40 000 Pa.",
            "C. 90 000 Pa.",
            "D. 600 000 Pa."
        ],
        "correct": "A",
        "solution": "Khối lượng m = 6 tấn = 6000 kg. Trọng lượng P = 10 . 6000 = 60 000 N. Áp suất p = P / S = 60 000 / 0,15 = 400 000 Pa."
    },
    {
        "id": 17,
        "question": "Câu 17. Trong thí nghiệm đo khối lượng riêng của một hòn sỏi bằng bình chia độ: thể tích nước ban đầu là 60cm^3^ sau khi thả chìm hoàn toàn hòn sỏi vào thì mực nước dâng lên 85cm^3^. Biết khối lượng hòn sỏi đo được là 65g. Khối lượng riêng của hòn sỏi là bao nhiêu?",
        "options": [
            "*A. 2,6g/cm^3^.* ",
            "B. 1,08g/cm^3^.",
            "C. 0,76g/cm^3^.",
            "D. 3,4g/cm^3^"
        ],
        "correct": "A",
        "solution": "V = 85 - 60 = 25 cm^3^. Khối lượng riêng D = m / V = 65 / 25 = 2,6 g/cm^3^."
    },
    {
        "id": 18,
        "question": "Câu 18. Một người dùng búa nhổ đinh (đòn bẩy loại 1). Khoảng cách từ điểm tựa đến đinh là 4cm, khoảng cách từ điểm tựa đến vị trí tay cầm tác dụng lực là 24cm. Để nhổ được chiếc đinh có lực cản nhỏ là 300N, người đó phải tác dụng một lực nhỏ nhất bằng bao nhiêu?",
        "options": [
            "A. 25N.",
            "*B. 50N.* ",
            "C. 75N.",
            "D. 1800N."
        ],
        "correct": "B",
        "solution": "Theo điều kiện cân bằng đòn bẩy: F~1~ . d~1~ = F~2~ . d~2~ => F . 24 = 300 . 4 => F = 1200 / 24 = 50 N."
    },
    {
        "id": 19,
        "question": "Câu 19. Muốn tháo một đai ốc bị rỉ sét rất chặt bằng cờ lê, cách nào sau đây giúp làm tăng tác dụng làm quay của lực hiệu quả nhất?",
        "options": [
            "*A. Dùng cờ lê có cán dài hơn và tác dụng lực vuông góc vào đầu cán.* ",
            "B. Tác dụng lực rất lớn có phương đi thẳng qua tâm ốc.",
            "C. Cầm tay sát vào vị trí đầu cờ lê sát ốc để vặn.",
            "D. Tác dụng lực song song với mặt phẳng của cờ lê."
        ],
        "correct": "A",
        "solution": "Tác dụng làm quay (moment lực M = F . d) tăng khi tăng cánh tay đòn d (dùng cán cờ lê dài hơn) và tác dụng lực vuông góc."
    },
    {
        "id": 20,
        "question": "Câu 20. Khi dùng đòn bẩy để bẩy một tảng đá, nếu dịch chuyển điểm tựa lại gần tảng đá hơn thì lực của người cần tác dụng vào đầu đòn bẩy sẽ như thế nào?",
        "options": [
            "A. tăng lên.",
            "*B. giảm đi.* ",
            "C. không thay đổi.",
            "D. bằng 0."
        ],
        "correct": "B",
        "solution": "Khi dịch điểm tựa lại gần tảng đá, cánh tay đòn của vật d~1~ giảm, cánh tay đòn của người d~2~ tăng lên => lực người tác dụng F~2~ = P . d~1~ / d~2~ sẽ giảm đi."
    },
    {
        "id": 21,
        "question": "Câu 21. Thiết bị nào sau đây là thiết bị đo điện?",
        "options": [
            "A. Pin.",
            "B. Cầu chì.",
            "*C. Ampe kế.* ",
            "D. Công tắc."
        ],
        "correct": "C",
        "solution": "Ampe kế là thiết bị dùng để đo cường độ dòng điện trong mạch."
    },
    {
        "id": 22,
        "question": "Câu 22. Để đo hiệu điện thế giữa hai đầu một thiết bị sử dụng điện thì cần sử dụng thiết bị nào sau đây?",
        "options": [
            "A. Huyết áp kế.",
            "B. Ampe kế.",
            "*C. Vôn kế.* ",
            "D. Nhiệt kế."
        ],
        "correct": "C",
        "solution": "Vôn kế là thiết bị chuyên dùng để đo hiệu điện thế giữa hai điểm trong mạch điện."
    },
    {
        "id": 23,
        "question": "Câu 23. Khi sử dụng các thiết bị đo điện (như ampe kế, vôn kế, joulemeter) trong phòng thí nghiệm, cần tuân thủ nguyên tắc nào sau đây?",
        "options": [
            "*A. Dùng đúng chức năng, đúng thang đo và mắc đúng mạch.* ",
            "B. Có thể dùng bất kỳ nguồn điện nào miễn là máy chạy được.",
            "C. Không cần chú ý tới cực âm và cực dương của thiết bị.",
            "D. Dùng thiết bị đo quá giới hạn thang đo để đo nhanh hơn."
        ],
        "correct": "A",
        "solution": "Phải sử dụng đúng chức năng, lựa chọn thang đo phù hợp với giá trị cần đo và mắc đúng cực (+/-) để tránh làm hỏng thiết bị đo."
    },
    {
        "id": 24,
        "question": "Câu 24. Khi tiến hành các thí nghiệm về điện, quy tắc an toàn nào sau đây không đúng?",
        "options": [
            "A. Giữ khoảng cách an toàn với nguồn điện.",
            "B. Lắp thiết bị đóng ngắt phù hợp.",
            "*C. Chạm tay trực tiếp vào các vị trí có điện thế nguy hiểm để kiểm tra.* ",
            "D. Sử dụng dây dẫn có vỏ bọc cách điện tốt."
        ],
        "correct": "C",
        "solution": "Tuyệt đối không được chạm tay trực tiếp vào các điểm có điện thế nguy hiểm vì có thể gây giật điện tử vong."
    },
    {
        "id": 25,
        "question": "Câu 25. Khối lượng riêng của sắt là 7 800 kg/m^3^. Hỏi 1 kg sắt sẽ có thể tích là bao nhiêu?",
        "options": [
            "A. 12,8 cm^3^.",
            "*B. 128 cm^3^.* ",
            "C. 1 280 cm^3^.",
            "D. 12 800 cm^3^."
        ],
        "correct": "B",
        "solution": "Thể tích V = m / D = 1 / 7800 m^3^ ≈ 0,0001282 m^3^ = 128,2 cm^3^ ≈ 128 cm^3^."
    },
    {
        "id": 26,
        "question": "Câu 26. Tính khối lượng của một khối đá hoa cương dạng hình hộp chữ nhật kích thước 2,0 m x 3,0 m x 1,5 m. Biết khối lượng riêng của đá hoa cương là ρ = 2 750 kg/m^3^.",
        "options": [
            "A. 2 475 kg.",
            "*B. 24 750 kg.* ",
            "C. 275 kg.",
            "D. 2 750 kg."
        ],
        "correct": "B",
        "solution": "V = 2,0 . 3,0 . 1,5 = 9 m^3^. Khối lượng m = ρ . V = 2750 . 9 = 24 750 kg."
    },

    # Page 3
    {
        "id": 27,
        "question": "Câu 27. Bức tượng phật Di Lặc tại chùa Vĩnh Tràng (Mỹ Tho, Tiền Giang) là một trong những bức tượng phật khổng lồ nổi tiếng trên thế giới. Tượng cao 20 m, nặng 250 tấn. Thể tích đồng được dùng để đúc bức tượng trên có giá trị là bao nhiêu? Biết khối lượng riêng của đồng là 8 900 kg/m^3^.",
        "options": [
            "A. 280,8 m^3^.",
            "B. 2,808 m^3^.",
            "C. 2808 m^3^.",
            "*D. 28,08 m^3^.* "
        ],
        "correct": "D",
        "solution": "m = 250 tấn = 250 000 kg. V = m / D = 250 000 / 8900 ≈ 28,08 m^3^."
    },
    {
        "id": 28,
        "question": "Câu 28. Biết 10 lít cát có khối lượng 15 kg. Tính thể tích của 1 tấn cát.",
        "options": [
            "*A. 0,667 m^3^.* ",
            "B. 0,667 m^4^.",
            "C. 0,778 m^3^.",
            "D. 0,778 m^4^."
        ],
        "correct": "A",
        "solution": "10 lít = 0,01 m^3^ có khối lượng 15 kg => Khối lượng riêng D = 15 / 0,01 = 1500 kg/m^3^. Thể tích của 1 tấn (1000 kg) cát là: V = 1000 / 1500 ≈ 0,667 m^3^."
    },
    {
        "id": 29,
        "question": "Câu 29. Hoạt động nào dưới đây không dùng vật dụng như một đòn bẩy?",
        "options": [
            "A. Dùng kéo cắt giấy.",
            "*B. Dùng búa đóng đinh.* ",
            "C. Dùng kìm cắt sắt.",
            "D. Dùng búa nhổ đinh."
        ],
        "correct": "B",
        "solution": "Dùng búa đóng đinh dựa trên tác dụng của động năng và lực va chạm trực tiếp, không sử dụng điểm tựa quay để khuếch đại lực như đòn bẩy."
    },
    {
        "id": 30,
        "question": "Câu 30. Khi hoạt động, đòn bẩy sẽ quay quanh cái gì?",
        "options": [
            "*A. điểm tựa.* ",
            "B. đầu chịu lực.",
            "C. điểm giữa của đòn.",
            "D. điểm tác dụng lực."
        ],
        "correct": "A",
        "solution": "Mọi đòn bẩy khi hoạt động đều quay quanh một trục cố định gọi là điểm tựa (thường kí hiệu là O)."
    },
    {
        "id": 31,
        "question": "Câu 31. Trong phòng thí nghiệm Khoa học tự nhiên 8, sau khi hoàn thành xong bài thực hành có sử dụng các hóa chất (dung dịch axit, kiềm, muối heavy metal/kim loại nặng...), quy tắc sử dụng và xử lý hóa chất thải nào dưới đây là đúng và an toàn nhất nhằm bảo vệ môi trường và giảm thiểu tác hại khi có thiên tai, lũ lụt xảy ra?",
        "options": [
            "A. Đổ trực tiếp tất cả các hóa chất dư thừa vào bồn rửa và xả thật nhiều nước để làm loãng hóa chất trước khi chảy ra cống thoát nước chung.",
            "*B. Thu gom, phân loại riêng biệt các hóa chất thải theo quy định, chứa trong các bình có nắp đậy kín và đặt ở nơi an toàn, đúng quy chuẩn để xử lý, tránh nguy cơ hóa chất độc hại tràn ra nguồn nước và thảm đất khi lũ lụt xảy ra.* ",
            "C. Trộn chung tất cả các chất thải hóa học rắn và lỏng vào một thùng rác thông thường của nhà trường để nhân viên vệ sinh mang đi xử lý.",
            "D. Đổ hóa chất thừa trực tiếp ra thảm cỏ hoặc đất trong khuôn viên trường vì đất có khả năng tự hấp thụ và phân hủy hóa chất an toàn."
        ],
        "correct": "B",
        "solution": "Chất thải hóa học đặc biệt là kim loại nặng, axit, kiềm phải được thu gom, phân loại riêng trong bình chứa chuyên dụng có nắp kín, bảo quản ở nơi cao ráo an toàn tránh thất thoát khi có thiên tai ngập lụt."
    },
    {
        "id": 32,
        "question": "Câu 32. Để đảm bảo an toàn trong phòng thí nghiệm Khoa học tự nhiên 8 và nâng cao khả năng ứng phó với biến đổi khí hậu (đặc biệt là tình trạng ngập lụt, thiên tai ngày càng khốc liệt), quy trình xử lý hóa chất dư thừa và chất thải sau giờ thực hành nào sau đây là đúng quy chuẩn nhất?",
        "options": [
            "A. Đổ tất cả chất thải lỏng vào bồn rửa và xả nhiều nước để làm loãng, xả trực tiếp ra hệ thống thoát nước chung.",
            "B. Gom chung hóa chất rắn và lỏng dư thừa vào thùng rác sinh hoạt của phòng thực hành để tiện cho việc dọn dẹp.",
            "*C. Thu gom, phân loại riêng biệt từng loại chất thải hóa học vào bình chứa chuyên dụng có nắp đậy kín, bảo quản ở vị trí an toàn trước khi mang đi xử lý tập trung.* ",
            "D. Đem hòa tan hóa chất thừa với nước rồi tưới ra đất trong khuôn viên trường để đất tự hấp thụ tự nhiên."
        ],
        "correct": "C",
        "solution": "Quy trình chuẩn: phân loại, chứa trong bình chuyên dụng có nắp kín và đưa đi xử lý tập trung đúng quy định môi trường."
    },
    {
        "id": 33,
        "question": "Câu 33. Trong các hiện tượng diễn ra trong đời sống dưới đây, hiện tượng nào thể hiện sự biến đổi vật lí?",
        "options": [
            "A. Cơm bị ôi thiu khi để bên ngoài thời tiết nắng nóng mùa hè.",
            "*B. Mặt đầm lầy bị đóng băng vào mùa đông lạnh giá.* ",
            "C. Thanh sắt ngoài không khí ẩm lâu ngày bị gỉ sét thành màu nâu đỏ.",
            "D. Thắp sáng nến, sáp nến cháy tạo ra khí carbon dioxide và hơi nước."
        ],
        "correct": "B",
        "solution": "Nước ở đầm lầy đóng băng chỉ là sự chuyển thể từ lỏng sang rắn của nước, không có chất mới sinh ra nên là biến đổi vật lí."
    },
    {
        "id": 34,
        "question": "Câu 34. Học sinh tiến hành thí nghiệm: Đun nóng một ít đường trắng trong ống nghiệm. Ban đầu đường chảy lỏng thành chất màu nâu, sau đó chuyển dần sang màu đen và xuất hiện các giọt nước đọng trên thành ống nghiệm. Quá trình này gồm những biến đổi nào?",
        "options": [
            "A. Chỉ xảy ra sự biến đổi vật lí.",
            "B. Chỉ xảy ra sự biến đổi hóa học.",
            "*C. Giai đoạn đường chảy lỏng là biến đổi vật lí; giai đoạn đường hóa đen và tạo giọt nước là biến đổi hóa học.* ",
            "D. Giai đoạn đường chảy lỏng là biến đổi hóa học; giai đoạn đường hóa đen là biến đổi vật lí."
        ],
        "correct": "C",
        "solution": "Giai đoạn đường nóng chảy thành chất lỏng là biến đổi vật lí (chuyển thể). Giai đoạn tiếp tục bị than hóa thành than màu đen (carbon) và giải phóng hơi nước là biến đổi hóa học."
    },
    {
        "id": 35,
        "question": "Câu 35. Trong tiết thực hành KHTN 8, nhóm học sinh thực hiện 2 thao tác thí nghiệm sau:\n- Thao tác 1: Trộn đều bột sắt (Fe) và bột lưu huỳnh (S) trong bát sứ.\n- Thao tác 2: Đưa hỗn hợp trên bếp cồn và đun nóng, thấy hỗn hợp cháy đỏ rực và chuyển thành chất rắn màu xám đen, không bị nam châm hút nữa.\nNhận xét nào sau đây về 2 thao tác trên là đúng?",
        "options": [
            "A. Thao tác 1 và Thao tác 2 đều là sự biến đổi hóa học.",
            "B. Thao tác 1 và Thao tác 2 đều là sự biến đổi vật lí.",
            "C. Thao tác 1 là sự biến đổi hóa học, Thao tác 2 là sự biến đổi vật lí.",
            "*D. Thao tác 1 là sự biến đổi vật lí, Thao tác 2 là sự biến đổi hóa học.* "
        ],
        "correct": "D",
        "solution": "Thao tác 1 chỉ trộn đều hai chất (tạo hỗn hợp vật lí, Fe vẫn bị nam châm hút). Thao tác 2 đun nóng tạo phản ứng Fe + S -> FeS (tạo chất mới không bị nam châm hút) là biến đổi hóa học."
    },

    # Page 4
    {
        "id": 36,
        "question": "Câu 36. Mol là lượng chất có chứa bao nhiêu hạt vĩ mô (nguyên tử hoặc phân tử) chất đó?",
        "options": [
            "A. 6,022 x 10^21^.",
            "B. 6,022 x 10^22^.",
            "*C. 6,022 x 10^23^.* ",
            "D. 6,022 x 10^24^."
        ],
        "correct": "C",
        "solution": "Mol là lượng chất có chứa N_A = 6,022 . 10^23^ nguyên tử hoặc phân tử chất đó (số Avogadro)."
    },
    {
        "id": 37,
        "question": "Câu 37. Thể tích mol của chất khí là gì?",
        "options": [
            "A. là thể tích chiếm bởi 1 gam chất khí đó ở nhiệt độ phòng.",
            "*B. là thể tích chiếm bởi N (6,022 x 10^23^) phân tử của chất khí đó.* ",
            "C. là khối lượng của 1 lít chất khí ở điều kiện xác định.",
            "D. là tổng thể tích của tất cả các chất khí có trong phòng thí nghiệm."
        ],
        "correct": "B",
        "solution": "Thể tích mol của chất khí là thể tích chiếm bởi 1 mol chất khí (chứa 6,022 . 10^23^ phân tử khí) ở điều kiện nhiệt độ và áp suất xác định."
    },
    {
        "id": 38,
        "question": "Câu 38. Công thức tính tỉ khối của khí A đối với không khí (d~A/kk~) được viết như thế nào?",
        "options": [
            "A. d~A/kk~ = M~A~ / 22,4.",
            "B. d~A/kk~ = M~A~ / 24,79.",
            "*C. d~A/kk~ = M~A~ / 29.* ",
            "D. d~A/kk~ = 29 / M~A~."
        ],
        "correct": "C",
        "solution": "Tỉ khối của khí A đối với không khí: d~A/kk~ = M~A~ / M~kk~ = M~A~ / 29."
    },
    {
        "id": 39,
        "question": "Câu 39. Bơm căng một quả bóng bay bằng 7,437 lít khí helium (He) ở điều kiện chuẩn (1 bar, 25°C). Số mol khí helium có trong quả bóng là bao nhiêu?",
        "options": [
            "A. 0,1 mol.",
            "B. 0,2 mol.",
            "*C. 0,3 mol.* ",
            "D. 0,4 mol."
        ],
        "correct": "C",
        "solution": "Ở đkc (1 bar, 25°C), thể tích 1 mol khí là 24,79 lít. n~He~ = V / 24,79 = 7,437 / 24,79 = 0,3 mol."
    },
    {
        "id": 40,
        "question": "Câu 40. Khi thu khí methane (CH~4~, M = 16 g/mol) vào bình bằng phương pháp dời chỗ không khí trong phòng thí nghiệm, người ta phải đặt úp ngược bình thu khí vì lý do nào sau đây?",
        "options": [
            "A. Khí CH~4~ nặng hơn không khí (d~CH4/kk~ ≈ 1,81).",
            "*B. Khí CH~4~ nhẹ hơn không khí (d~CH4/kk~ ≈ 0,55).* ",
            "C. Khí CH~4~ có thể tích mol lớn hơn 24,79 lít.",
            "D. Khí CH~4~ dễ bị phân hủy khi tiếp xúc với ánh sáng."
        ],
        "correct": "B",
        "solution": "Khí CH~4~ có M = 16 g/mol, nhẹ hơn không khí (d = 16/29 ≈ 0,55) nên bay lên trên, ta phải đặt úp ngược bình để dời không khí xuống dưới."
    },
    {
        "id": 41,
        "question": "Câu 41. Trong một trường học ở khu vực thường xảy ra mưa lũ, phòng thí nghiệm có chứa một số hóa chất. Khi dự báo có khả năng ngập lụt, giáo viên yêu cầu kiểm tra và xử lí khu vực hóa chất nhằm hạn chế hóa chất phát tán ra môi trường. Cách xử lí phù hợp nhất là gì?",
        "options": [
            "A. Đổ toàn bộ hóa chất thải xuống cống để tránh hóa chất tràn ra phòng.",
            "B. Trộn tất cả hóa chất thải vào một can chung để dễ vận chuyển.",
            "*C. Thu gom, phân loại và xử lí hóa chất thải theo đúng quy định an toàn.* ",
            "D. Để hóa chất thải ngoài sân cho nước mưa cuốn đi."
        ],
        "correct": "C",
        "solution": "Cần kiểm tra, phân loại, đậy kín nắp và di chuyển bình chứa hóa chất đến vị trí cao ráo, an toàn theo đúng quy chuẩn an toàn môi trường."
    },
    {
        "id": 42,
        "question": "Câu 42. Trong phòng thí nghiệm, khi pha dung dịch muối hoặc thực hiện thí nghiệm xác định thể tích chất lỏng, học sinh cần lựa chọn dụng cụ phù hợp. Dụng cụ nào sau đây thích hợp để đo tương đối chính xác một thể tích chất lỏng?",
        "options": [
            "A. Ống nghiệm.",
            "B. Cốc thủy tinh.",
            "*C. Ống đong.* ",
            "D. Đũa thủy tinh."
        ],
        "correct": "C",
        "solution": "Ống đong là dụng cụ có vạch chia độ chính xác dùng để đong đo thể tích chất lỏng."
    },
    {
        "id": 43,
        "question": "Câu 43. Trong đời sống hằng ngày, nhiều vật liệu có thể thay đổi dưới tác động của môi trường. Hiện tượng nào sau đây là biến đổi hóa học, làm xuất hiện chất mới?",
        "options": [
            "A. Nước đá tan thành nước khi để ngoài không khí.",
            "B. Muối ăn tan trong nước khi dùng để nấu ăn.",
            "C. Giấy bị xé thành nhiều mảnh khi làm thủ công.",
            "*D. Đinh sắt để lâu ngoài không khí ẩm bị gỉ.* "
        ],
        "correct": "D",
        "solution": "Đinh sắt bị gỉ trong không khí ẩm tạo thành gỉ sắt Fe~2~O~3~.nH~2~O (chất mới có màu nâu đỏ), đây là biến đổi hóa học."
    },
    {
        "id": 44,
        "question": "Câu 44. Trong đời sống, sodium hydrogen carbonate (NaHCO~3~, còn gọi là baking soda) được sử dụng trong chế biến thực phẩm và một số ứng dụng làm sạch. Khi cho NaHCO~3~ vào cốc giấm ăn, học sinh quan sát thấy nhiều bọt khí thoát ra. Dấu hiệu này cho thấy điều gì?",
        "options": [
            "A. Chỉ xảy ra biến đổi vật lí.",
            "*B. Có chất khí mới được tạo thành trong phản ứng hóa học.* ",
            "C. Nước bị bay hơi mạnh.",
            "D. Chất tan chỉ chuyển từ dạng rắn sang dạng lỏng."
        ],
        "correct": "B",
        "solution": "Xảy ra phản ứng: NaHCO~3~ + CH~3~COOH -> CH~3~COONa + H~2~O + CO~2~ ↑. Bọt khí CO~2~ sinh ra là dấu hiệu nhận biết phản ứng hóa học tạo chất mới."
    },
    {
        "id": 45,
        "question": "Câu 45. Nước là một chất thiết yếu đối với sự sống và là thành phần quan trọng trong nhiều dung dịch sinh học. Trong hóa học, khối lượng mol của nước được sử dụng để tính toán lượng chất trong các quá trình pha chế và phản ứng hóa học. Khối lượng mol của H~2~O là bao nhiêu?",
        "options": [
            "A. 16 g/mol.",
            "B. 17 g/mol.",
            "*C. 18 g/mol.* ",
            "D. 20 g/mol."
        ],
        "correct": "C",
        "solution": "M~H2O~ = 2 . 1 + 16 = 18 g/mol."
    },
    {
        "id": 46,
        "question": "Câu 46. Trong một số quá trình sản xuất, phòng thí nghiệm và nghiên cứu môi trường, người ta cần tính số mol khí từ thể tích khí đo được. Ở 25 °C và áp suất 1 bar, một mẫu khí có thể tích 49,58 L. Số mol khí trong mẫu là bao nhiêu?",
        "options": [
            "A. 0,5 mol.",
            "B. 1,0 mol.",
            "*C. 2,0 mol.* ",
            "D. 2,5 mol."
        ],
        "correct": "C",
        "solution": "n = V / 24,79 = 49,58 / 24,79 = 2,0 mol."
    }
]

PART1_QUESTIONS_CONTINUED = [
    # Page 5
    {
        "id": 47,
        "question": "Câu 47. Nước muối, nước đường và nhiều loại dung dịch được sử dụng phổ biến trong đời sống. Ví dụ, khi hòa tan hoàn toàn một lượng muối ăn thích hợp vào nước, ta thu được một hỗn hợp đồng nhất. Phát biểu nào sau đây đúng về dung dịch?",
        "options": [
            "A. Dung dịch luôn gồm chất rắn và chất khí.",
            "*B. Dung dịch là hỗn hợp lỏng đồng nhất của các chất đã tan trong nhau.* ",
            "C. Dung dịch luôn chỉ có một chất.",
            "D. Dung dịch là hỗn hợp không đồng nhất."
        ],
        "correct": "B",
        "solution": "Theo định nghĩa SGK KHTN 8: Dung dịch là hỗn hợp lỏng đồng nhất của chất tan và dung môi (các chất đã tan trong nhau)."
    },
    {
        "id": 48,
        "question": "Câu 48. Muối ăn (NaCl) được sử dụng hằng ngày để chế biến và bảo quản thực phẩm. Một gia đình pha dung dịch nước muối bằng cách hòa tan hoàn toàn 10 g NaCl vào 90 g nước. Nồng độ phần trăm của dung dịch thu được là bao nhiêu?",
        "options": [
            "A. 5%.",
            "B. 9%.",
            "*C. 10%.* ",
            "D. 11,1%."
        ],
        "correct": "C",
        "solution": "Khối lượng dung dịch m~dd~ = m~ct~ + m~dm~ = 10 + 90 = 100 g. C% = (10 / 100) . 100% = 10%."
    },
    {
        "id": 49,
        "question": "Câu 49. Hydrogen là một chất khí có thể cháy trong oxygen và tạo thành nước. Phản ứng này có ý nghĩa trong nghiên cứu năng lượng và công nghệ hydrogen. Phương trình 2H~2~ + O~2~ → 2H~2~O cho biết tỉ lệ số phân tử H~2~ : O~2~ : H~2~O như thế nào?",
        "options": [
            "A. 1 : 1 : 1.",
            "*B. 2 : 1 : 2.* ",
            "C. 2 : 2 : 1.",
            "D. 1 : 2 : 2."
        ],
        "correct": "B",
        "solution": "Hệ số tỉ lượng trong phương trình hóa học chính là tỉ lệ số phân tử: 2 phân tử H~2~ tác dụng 1 phân tử O~2~ tạo 2 phân tử H~2~O (tỉ lệ 2 : 1 : 2)."
    },
    {
        "id": 50,
        "question": "Câu 50. Nhôm được sử dụng rộng rãi để sản xuất vỏ lon, đồ gia dụng, phương tiện giao thông và nhiều thiết bị kỹ thuật nhờ khối lượng riêng nhỏ và khả năng chống ăn mòn tương đối tốt. Khi nhôm tác dụng với oxygen tạo thành aluminium oxide (Al~2~O~3~), phương trình hóa học nào cân bằng đúng?",
        "options": [
            "A. 2Al + O~2~ → Al~2~O~3~",
            "B. Al + O~2~ → Al~2~O~3~",
            "*C. 4Al + 3O~2~ → 2Al~2~O~3~* ",
            "D. 2Al + 3O~2~ → Al~2~O~3~"
        ],
        "correct": "C",
        "solution": "Cân bằng nguyên tử: 4 nguyên tử Al ở 2 vế và 6 nguyên tử O ở 2 vế: 4Al + 3O~2~ -> 2Al~2~O~3~."
    },
    {
        "id": 51,
        "question": "Câu 51. Trong phòng thí nghiệm, nhãn của một chai ghi “NaOH – CORROSIVE”. Cách xử lí phù hợp nhất khi lấy hóa chất là gì?",
        "options": [
            "A. dùng tay lấy trực tiếp một lượng nhỏ.",
            "*B. đọc nhãn, dùng dụng cụ thích hợp và trang bị bảo hộ.* ",
            "C. đưa chai hóa chất lại gần mũi để nhận biết.",
            "D. đổ phần hóa chất còn thừa trở lại chai gốc."
        ],
        "correct": "B",
        "solution": "Nhãn 'CORROSIVE' cảnh báo chất ăn mòn, nguy hiểm cho da và mắt. Bắt buộc phải đọc nhãn, mang găng tay, kính bảo hộ và dùng kẹp/thìa lấy hóa chất thích hợp."
    },
    {
        "id": 52,
        "question": "Câu 52. Thiết bị nào được dùng để đo pH của một dung dịch một cách thuận tiện trong phòng thí nghiệm?",
        "options": [
            "A. Ampe kế.",
            "B. Vôn kế.",
            "*C. Máy đo pH.* ",
            "D. Joulemeter."
        ],
        "correct": "C",
        "solution": "Máy đo pH (pH meter) hoặc bút đo pH là thiết bị chuyên dùng để đo chính xác và nhanh chóng giá trị pH của dung dịch."
    },
    {
        "id": 53,
        "question": "Câu 53. Sau một thí nghiệm, dung dịch hóa chất còn thừa không được phép đổ tùy tiện xuống bồn rửa. Việc làm phù hợp nhất là gì?",
        "options": [
            "A. đổ ra sân trường cho bay hơi.",
            "B. trộn với mọi hóa chất khác để giảm nồng độ.",
            "*C. thu gom, phân loại và xử lí theo quy định.* ",
            "D. đổ vào chai nước uống đã dùng."
        ],
        "correct": "C",
        "solution": "Hóa chất thừa phải được thu gom vào bình thải quy định và dán nhãn phân loại để xử lý an toàn."
    },
    {
        "id": 54,
        "question": "Câu 54. Một học sinh sử dụng bộ nguồn điện trong phòng thí nghiệm. Hành động nào sau đây an toàn?",
        "options": [
            "A. Thay dây dẫn khi nguồn đang cấp điện.",
            "B. Chạm đồng thời hai đầu dây bằng tay.",
            "*C. Kiểm tra dây dẫn, chọn đúng mức điện áp và ngắt nguồn trước khi tháo lắp.* ",
            "D. Dùng dây điện bị hở để quan sát tia lửa."
        ],
        "correct": "C",
        "solution": "Luôn ngắt nguồn điện trước khi cắm hoặc tháo lắp dây dẫn và kiểm tra chọn đúng mức điện áp cho phép."
    },
    {
        "id": 55,
        "question": "Câu 55. Một cốc nước đá để ngoài không khí, sau đó đá tan thành nước. Hiện tượng này được xếp như thế nào?",
        "options": [
            "A. biến đổi hóa học vì có nước tạo thành.",
            "*B. biến đổi vật lí vì không tạo chất mới.* ",
            "C. phản ứng tỏa nhiệt.",
            "D. phản ứng thu nhiệt tạo chất mới."
        ],
        "correct": "B",
        "solution": "Nước đá tan chảy từ trạng thái rắn sang trạng thái lỏng vẫn là chất nước (H~2~O), không có chất mới tạo thành nên là biến đổi vật lí."
    },
    {
        "id": 56,
        "question": "Câu 56. Khi cho một viên sủi vào nước, quan sát thấy có nhiều bọt khí thoát ra. Dấu hiệu này cho thấy điều gì?",
        "options": [
            "A. chắc chắn chỉ có sự thay đổi trạng thái.",
            "*B. có khả năng xuất hiện chất khí mới do phản ứng hóa học.* ",
            "C. nước bị đông đặc.",
            "D. khối lượng các chất tăng lên."
        ],
        "correct": "B",
        "solution": "Viên sủi chứa axit hữu cơ (citric acid) và muối kiềm (NaHCO~3~). Khi gặp nước, các chất hòa tan và phản ứng tạo khí CO~2~ thoát ra, đây là dấu hiệu của phản ứng hóa học."
    },
    {
        "id": 57,
        "question": "Câu 57. Khi đốt cháy cồn, nhiệt và ánh sáng được giải phóng ra môi trường. Đây là ví dụ điển hình của quá trình nào?",
        "options": [
            "A. phản ứng thu nhiệt.",
            "*B. phản ứng tỏa nhiệt.* ",
            "C. biến đổi vật lí.",
            "D. sự bay hơi."
        ],
        "correct": "B",
        "solution": "Phản ứng đốt cháy cồn giải phóng năng lượng dưới dạng nhiệt và ánh sáng ra môi trường là phản ứng tỏa nhiệt."
    },

    # Page 6
    {
        "id": 58,
        "question": "Câu 58. Phát biểu nào mô tả đúng khái niệm mol?",
        "options": [
            "A. Mol là khối lượng của một chất.",
            "*B. Mol là lượng chất chứa 6,022 x 10^23^ hạt vi mô.* ",
            "C. Mol là thể tích của một chất khí.",
            "D. Mol chỉ dùng cho chất khí."
        ],
        "correct": "B",
        "solution": "Mol là lượng chất có chứa N_A (6,022 . 10^23^) nguyên tử hoặc phân tử của chất đó."
    },
    {
        "id": 59,
        "question": "Câu 59. Khối lượng mol của CO~2~ là bao nhiêu?",
        "options": [
            "A. 28 g/mol.",
            "B. 32 g/mol.",
            "*C. 44 g/mol.* ",
            "D. 46 g/mol."
        ],
        "correct": "C",
        "solution": "M~CO2~ = 12 + 16 . 2 = 44 g/mol."
    },
    {
        "id": 60,
        "question": "Câu 60. Số mol CO~2~ có trong 8,8 g CO~2~ là bao nhiêu?",
        "options": [
            "A. 0,10 mol.",
            "*B. 0,20 mol.* ",
            "C. 0,40 mol.",
            "D. 1,00 mol."
        ],
        "correct": "B",
        "solution": "n = m / M = 8,8 / 44 = 0,20 mol."
    },
    {
        "id": 61,
        "question": "Câu 61. Phát biểu nào dưới đây trình bày đúng nhất vai trò của hệ cơ quan trong cơ thể người?",
        "options": [
            "A. Mỗi hệ cơ quan chỉ đảm nhiệm một chức năng riêng biệt và không liên hệ với hệ khác.",
            "*B. Các hệ cơ quan phối hợp hoạt động, đảm bảo cơ thể thực hiện các chức năng sống.* ",
            "C. Hệ cơ quan chỉ có chức năng bảo vệ các cơ quan bên trong cơ thể.",
            "D. Hệ cơ quan là đơn vị cấu tạo nhỏ nhất của cơ thể người."
        ],
        "correct": "B",
        "solution": "Cơ thể người là một thể thống nhất, các hệ cơ quan phối hợp nhịp nhàng dưới sự điều khiển của hệ thần kinh và thể dịch để duy trì hoạt động sống."
    },
    {
        "id": 62,
        "question": "Câu 62. Thành phần nào dưới đây giữ vai trò nâng đỡ cơ thể, bảo vệ các cơ quan và phối hợp với cơ tạo nên vận động?",
        "options": [
            "A. Hệ cơ.",
            "*B. Hệ xương.* ",
            "C. Hệ tiêu hoá.",
            "D. Hệ tuần hoàn."
        ],
        "correct": "B",
        "solution": "Hệ xương tạo nên bộ khung nâng đỡ cơ thể, làm chỗ bám cho hệ cơ và tạo khoang bảo vệ các nội quan."
    },
    {
        "id": 63,
        "question": "Câu 63. Trong hệ vận động, bộ phận nào sau đây có vai trò co và giãn để tạo lực làm xương chuyển động?",
        "options": [
            "A. Khớp.",
            "B. Xương.",
            "*C. Cơ.* ",
            "D. Dây chằng."
        ],
        "correct": "C",
        "solution": "Cơ bắp có tính chất cơ bản là co và giãn, khi co sẽ kéo xương chuyển động tạo nên cử động."
    },
    {
        "id": 64,
        "question": "Câu 64. Phát biểu nào sau đây mô tả đúng vai trò của khớp trong hệ vận động?",
        "options": [
            "*A. Là nơi nối các xương và giúp các xương có thể vận động tương đối với nhau.* ",
            "B. Là nơi sản sinh lực chủ yếu làm xương co ngắn lại.",
            "C. Là bộ phận dự trữ chất dinh dưỡng chủ yếu cho cơ thể.",
            "D. Là cơ quan trực tiếp vận chuyển oxygen đến tế bào cơ."
        ],
        "correct": "A",
        "solution": "Khớp xương là nơi tiếp giáp giữa các đầu xương, giúp liên kết các xương và cho phép chúng cử động linh hoạt tương đối với nhau."
    },
    {
        "id": 65,
        "question": "Câu 65. Cơ quan nào dưới đây là nơi diễn ra phần lớn quá trình tiêu hóa và hấp thụ chất dinh dưỡng?",
        "options": [
            "A. Khoang miệng.",
            "B. Dạ dày.",
            "*C. Ruột non.* ",
            "D. Ruột già."
        ],
        "correct": "C",
        "solution": "Ruột non có bề mặt hấp thụ cực lớn nhờ các nếp gấp, lông ruột và mạng mao mạch máu, là nơi hoàn tất tiêu hóa và hấp thụ hầu hết chất dinh dưỡng."
    },
    {
        "id": 66,
        "question": "Câu 66. Vai trò chủ yếu của dạ dày trong quá trình tiêu hóa thức ăn là gì?",
        "options": [
            "A. Hấp thụ phần lớn chất dinh dưỡng vào máu.",
            "*B. Chứa, nhào trộn thức ăn và thực hiện một phần tiêu hóa thức ăn.* ",
            "C. Hấp thụ chủ yếu nước và tạo phân.",
            "D. Tiết mật để nhũ tương hóa chất béo."
        ],
        "correct": "B",
        "solution": "Dạ dày có lớp cơ dày khỏe giúp co bóp, nhào trộn thức ăn với dịch vị và biến đổi một phần protein nhờ enzyme pepsin."
    },
    {
        "id": 67,
        "question": "Câu 67. Xét các phát biểu sau về cơ thể người:\n(1) Tế bào là đơn vị cấu tạo và chức năng của cơ thể.\n(2) Hệ vận động giúp cơ thể nâng đỡ và thực hiện vận động.\n(3) Hệ tiêu hóa có vai trò biến đổi thức ăn thành các chất cơ thể có thể hấp thụ và sử dụng.\n(4) Mọi cơ quan trong cơ thể hoạt động hoàn toàn độc lập với nhau.\nCó bao nhiêu phát biểu đúng?",
        "options": [
            "A. 1.",
            "B. 2.",
            "*C. 3.* ",
            "D. 4."
        ],
        "correct": "C",
        "solution": "Có 3 phát biểu đúng: (1), (2), (3). Phát biểu (4) sai vì các cơ quan trong cơ thể phối hợp hoạt động nhịp nhàng, thống nhất chứ không độc lập."
    },
    {
        "id": 68,
        "question": "Câu 68. Vai trò chủ yếu của hệ thần kinh và hệ nội tiết trong cơ thể người là gì?",
        "options": [
            "A. Tiếp nhận, biến đổi và hấp thụ chất dinh dưỡng.",
            "B. Vận chuyển oxygen và chất dinh dưỡng đến tế bào.",
            "*C. Điều hòa, điều khiển và phối hợp hoạt động của các cơ quan.* ",
            "D. Nâng đỡ cơ thể và bảo vệ các cơ quan nội tạng."
        ],
        "correct": "C",
        "solution": "Hệ thần kinh và hệ nội tiết đảm nhận chức năng điều khiển, điều hòa và phối hợp mọi hoạt động của các cơ quan trong cơ thể."
    },
    {
        "id": 69,
        "question": "Câu 69. Tủy xương đỏ có vai trò quan trọng nào sau đây đối với cơ thể người?",
        "options": [
            "A. Dự trữ chủ yếu chất béo cho cơ thể.",
            "B. Tiết dịch khớp giúp các đầu xương vận động.",
            "*C. Tạo ra các tế bào máu.* ",
            "D. Liên kết các đầu xương với nhau."
        ],
        "correct": "C",
        "solution": "Tủy đỏ của xương chứa các tế bào gốc sinh máu, có chức năng sinh hồng cầu, bạch cầu và tiểu cầu."
    },
    {
        "id": 70,
        "question": "Câu 70. Trong cấu tạo của hệ vận động, dây chằng có chức năng chủ yếu nào sau đây?",
        "options": [
            "A. Nối cơ với xương, truyền lực co cơ.",
            "*B. Nối các xương với nhau, góp phần giữ vững khớp.* ",
            "C. Tạo hồng cầu và bạch cầu cho cơ thể.",
            "D. Bao bọc đầu xương và trực tiếp tạo ra vận động."
        ],
        "correct": "B",
        "solution": "Dây chằng là các dải mô liên kết sợi dai chắc nối giữa các đầu xương với nhau, giúp giữ vững khớp và giới hạn chuyển động bất thường."
    },
    {
        "id": 71,
        "question": "Câu 71. Dịch mật có vai trò chủ yếu nào trong quá trình tiêu hóa thức ăn?",
        "options": [
            "A. Phân giải protein thành các amino acid.",
            "*B. Nhũ tương hóa lipid, tạo điều kiện thuận lợi cho enzyme lipase hoạt động.* ",
            "C. Phân giải tinh bột thành đường đơn.",
            "D. Phân giải cellulose thành glucose."
        ],
        "correct": "B",
        "solution": "Dịch mật do gan tiết ra không chứa enzyme tiêu hóa nhưng chứa muối mật giúp nhũ tương hóa lipid (phân tán chất béo thành các giọt nhỏ) để enzyme lipase dễ dàng thủy phân."
    },

    # Page 7
    {
        "id": 72,
        "question": "Câu 72. Đặc điểm cấu tạo nào của ruột non thích nghi với chức năng hấp thụ chất dinh dưỡng?",
        "options": [
            "A. Thành ruột dày, có nhiều lớp cơ phát triển.",
            "*B. Bề mặt bên trong có nhiều nếp gấp, lông ruột và vi nhung mao.* ",
            "C. Lòng ruột rộng, chủ yếu chứa phân.",
            "D. Thành ruột có nhiều tuyến tiết hydrochloric acid."
        ],
        "correct": "B",
        "solution": "Niêm mạc ruột non có vô số nếp gấp, lông ruột và vi nhung mao giúp tăng diện tích bề mặt hấp thụ lên gấp hàng trăm lần (khoảng 400 - 500 m^2^)."
    },
    {
        "id": 73,
        "question": "Câu 73. Cơ quan nào sau đây thuộc hệ hô hấp của cơ thể người?",
        "options": [
            "A. Thực quản.",
            "*B. Phế quản.* ",
            "C. Tĩnh mạch chủ.",
            "D. Ruột non."
        ],
        "correct": "B",
        "solution": "Hệ hô hấp gồm đường dẫn khí (mũi, họng, thanh quản, khí quản, phế quản) và hai lá phổi."
    },
    {
        "id": 74,
        "question": "Câu 74. Hệ cơ quan nào có chức năng chủ yếu là vận chuyển oxygen, chất dinh dưỡng và các chất khác đến tế bào, đồng thời vận chuyển chất thải đến cơ quan bài tiết?",
        "options": [
            "A. Hệ tiêu hóa.",
            "*B. Hệ tuần hoàn.* ",
            "C. Hệ hô hấp.",
            "D. Hệ bài tiết."
        ],
        "correct": "B",
        "solution": "Hệ tuần hoàn (gồm tim và mạch máu) vận chuyển máu mang oxygen, chất dinh dưỡng đi nuôi tế bào và chở chất thải đến phổi, thận để bài tiết."
    },
    {
        "id": 75,
        "question": "Câu 75. Bộ xương người được chia thành những phần chính nào sau đây?",
        "options": [
            "*A. Xương đầu, xương thân và xương chi.* ",
            "B. Xương đầu, xương ngực và xương chân.",
            "C. Xương sọ, xương cột sống và xương tay.",
            "D. Xương đầu, xương chậu và xương chi dưới."
        ],
        "correct": "A",
        "solution": "Bộ xương người chia làm 3 phần chính: xương đầu (khối xương sọ và xương mặt), xương thân (cột sống và lồng ngực), và xương chi (chi trên và chi dưới)."
    },
    {
        "id": 76,
        "question": "Câu 76. Cơ thể người được chia làm mấy phần? Đó là những phần nào?",
        "options": [
            "A. 3 phần: đầu, thân và chân.",
            "B. 2 phần: đầu và thân.",
            "*C. 3 phần: đầu, thân và các chi.* ",
            "D. 3 phần: đầu, cổ và thân."
        ],
        "correct": "C",
        "solution": "Cơ thể người chia làm 3 phần: đầu, thân và các chi (tay, chân)."
    },
    {
        "id": 77,
        "question": "Câu 77. Để chống vẹo cột sống, cần phải làm gì?",
        "options": [
            "*A. Khi ngồi phải ngay ngắn, không nghiêng vẹo.* ",
            "B. Mang vác về một bên liên tục.",
            "C. Mang vác quá sức chịu đựng.",
            "D. Ngồi nghiêng vẹo một bên."
        ],
        "correct": "A",
        "solution": "Cần ngồi học đúng tư thế, lưng thẳng, không tì ngực vào bàn, mang vác đều hai vai để phòng chống cong vẹo cột sống."
    },
    {
        "id": 78,
        "question": "Câu 78. Chức năng của hệ tiêu hóa của người là gì?",
        "options": [
            "A. xử lí cơ học thức ăn.",
            "B. thủy phân thức ăn thành các đơn phân tiêu hóa được.",
            "C. loại bỏ thức ăn không cần thiết.",
            "*D. Cả A, B và C đều đúng* "
        ],
        "correct": "D",
        "solution": "Hệ tiêu hóa thực hiện biến đổi cơ học (nhai, nhào trộn), biến đổi hóa học (thủy phân thức ăn thành chất dinh dưỡng đơn giản) và bài tiết cặn bã (phân)."
    },
    {
        "id": 79,
        "question": "Câu 79. Khi mất khả năng dung nạp chất dinh dưỡng, cơ thể chúng ta sẽ trở nên kiệt quệ, đồng thời khả năng vận động cũng bị ảnh hưởng nặng nề. Ví dụ trên phản ánh điều gì?",
        "options": [
            "*A. Các hệ cơ quan trong cơ thể có mối liên hệ mật thiết với nhau.* ",
            "B. Dinh dưỡng là không cần thiết của cơ và xương.",
            "C. Hệ tiêu hóa chỉ có chức năng đào thải chất cặn bã ra khỏi cơ thể.",
            "D. Hệ vận động có thể hoạt động bình thường mà không cần chất dinh dưỡng từ thức ăn."
        ],
        "correct": "A",
        "solution": "Hệ tiêu hóa cung cấp năng lượng và nguyên liệu nuôi dưỡng hệ cơ, xương; khi hệ tiêu hóa suy yếu thì hệ vận động cũng bị suy sụp, thể hiện mối quan hệ thống nhất giữa các hệ cơ quan."
    },
    {
        "id": 80,
        "question": "Câu 80. Hệ thần kinh có chức năng nào sau đây?",
        "options": [
            "A. Giúp cơ thể di chuyển, vận động.",
            "B. Trao đổi khí O~2~, CO~2~ với môi trường.",
            "C. Biến đổi thức ăn thành các chất cơ thể có thể hấp thụ.",
            "*D. Điều khiển, điều hoà và phối hợp các hoạt động của các cơ quan.* "
        ],
        "correct": "D",
        "solution": "Hệ thần kinh tiếp nhận kích thích, xử lý thông tin và điều khiển, điều hòa mọi phản ứng của cơ thể với môi trường."
    },
    {
        "id": 81,
        "question": "Câu 81. Hệ cơ quan nào có vai trò lọc các chất thải có hại cho cơ thể từ máu và thải ra môi trường?",
        "options": [
            "A. Hệ hô hấp.",
            "*B. Hệ bài tiết.* ",
            "C. Hệ tiêu hóa.",
            "D. Hệ tuần hoàn."
        ],
        "correct": "B",
        "solution": "Hệ bài tiết (chủ yếu là đôi thận) lọc máu, loại bỏ các chất cặn bã, độc hại hòa tan trong máu để tạo nước tiểu thải ra ngoài."
    },
    {
        "id": 82,
        "question": "Câu 82. Hệ vận động của người có chức năng gì?",
        "options": [
            "A. Giúp cơ thể nhận biết được các vật và thu nhận âm thanh.",
            "B. Lọc các chất thải có hại cho cơ thể từ máu và thải ra khỏi cơ thể.",
            "*C. Nâng đỡ và bảo vệ cơ thể, tạo ra hình dạng, duy trì tư thế và giúp con người vận động.* ",
            "D. Điều hòa hoạt của các cơ quan trong cơ thể thông qua viết tiết một số loại hormone tác động đến cơ quan nhất định."
        ],
        "correct": "C",
        "solution": "Hệ vận động (gồm xương, khớp, cơ) định hình cơ thể, nâng đỡ, bảo vệ nội quan và thực hiện mọi cử động, di chuyển."
    },
    {
        "id": 83,
        "question": "Câu 83. Não, tuỷ sống là cơ quan thuộc hệ nào?",
        "options": [
            "A. Hệ bài tiết.",
            "B. Hệ tiêu hoá.",
            "*C. Hệ thần kinh.* ",
            "D. Hệ vận động."
        ],
        "correct": "C",
        "solution": "Não bộ và tủy sống là hai bộ phận trung ương của hệ thần kinh."
    },
    {
        "id": 84,
        "question": "Câu 84. Hệ cơ quan nào có vai trò thực hiện quá trình sinh sản?",
        "options": [
            "*A. Hệ sinh dục.* ",
            "B. Hệ hô hấp.",
            "C. Hệ tuần hoàn.",
            "D. Hệ bài tiết."
        ],
        "correct": "A",
        "solution": "Hệ sinh dục sản sinh giao tử (tinh trùng/trứng) và hormone sinh dục, duy trì nòi giống qua sinh sản."
    },
    {
        "id": 85,
        "question": "Câu 85. Tiêu hóa là quá trình như thế nào?",
        "options": [
            "*A. biến đổi thức ăn thành các chất đơn giản mà cơ thể có thể hấp thụ.* ",
            "B. vận chuyển oxygen đến tế bào.",
            "C. loại bỏ toàn bộ thức ăn ra khỏi cơ thể.",
            "D. tạo ra máu cho cơ thể."
        ],
        "correct": "A",
        "solution": "Tiêu hóa là quá trình biến đổi thức ăn phức tạp thành các chất dinh dưỡng đơn giản mà tế bào cơ thể có thể hấp thụ qua màng ruột."
    },

    # Page 8
    {
        "id": 86,
        "question": "Câu 86. Mối quan hệ giữa tiêu hóa và dinh dưỡng được thể hiện ở nội dung nào sau đây?",
        "options": [
            "*A. Tiêu hóa giúp cơ thể hấp thụ và sử dụng các chất dinh dưỡng từ thức ăn.* ",
            "B. Tiêu hóa làm cơ thể không cần chất dinh dưỡng.",
            "C. Tiêu hóa và dinh dưỡng không có mối liên hệ với nhau.",
            "D. Dinh dưỡng chỉ diễn ra sau khi thức ăn được thải ra ngoài."
        ],
        "correct": "A",
        "solution": "Tiêu hóa là tiền đề cơ sở cung cấp các phân tử dinh dưỡng để cơ thể hấp thụ vào máu và vận chuyển đi nuôi tế bào."
    },
    {
        "id": 87,
        "question": "Câu 87. Trong quá trình tiêu hóa ở người, thức ăn được biến đổi nhờ hoạt động của các cơ quan trong hệ tiêu hóa và sự hỗ trợ của các dịch tiêu hóa. Sản phẩm của quá trình tiêu hóa thức ăn là gì?",
        "options": [
            "*A. Các chất dinh dưỡng đơn giản mà cơ thể có thể hấp thụ qua thành ruột để cung cấp cho các tế bào.* ",
            "B. Các chất dinh dưỡng phức tạp được vận chuyển trực tiếp đến các tế bào mà không cần hấp thụ.",
            "C. Khí oxygen và khí carbon dioxide được trao đổi giữa cơ thể với môi trường bên ngoài.",
            "D. Các chất thải được hình thành và đào thải ra ngoài cơ thể qua hệ bài tiết."
        ],
        "correct": "A",
        "solution": "Sản phẩm có ích của tiêu hóa là đường đơn (glucose), axit amin, axit béo, glycerol... có kích thước phân tử nhỏ hấp thụ được vào máu và bạch huyết."
    },
    {
        "id": 88,
        "question": "Câu 88. Hoạt động nào sau đây có lợi cho sự phát triển của hệ vận động?",
        "options": [
            "*A. Chạy bộ, bơi lội, tập thể dục.* ",
            "B. Ngồi lâu một chỗ.",
            "C. Nằm nghỉ cả ngày.",
            "D. Chơi điện tử liên tục."
        ],
        "correct": "A",
        "solution": "Tập luyện thể dục thể thao thường xuyên, vừa sức giúp xương chắc khỏe, cơ phát triển săn chắc, khớp linh hoạt dẻo dai."
    },
    {
        "id": 89,
        "question": "Câu 89. Khi sơ cứu cho người bị gãy xương cần chú ý điều gì?",
        "options": [
            "A. Rửa sạch vết thương, băng buộc chặt chỗ gãy.",
            "B. Bôi dầu gió.",
            "*C. Không nắn bóp bừa bãi, dùng nẹp băng cố định chỗ gãy.* ",
            "D. Chườm đá lạnh đỡ đau."
        ],
        "correct": "C",
        "solution": "Tuyệt đối không cử động hay nắn bóp chỗ gãy vì có thể làm đầu xương gãy đâm rách mạch máu, dây thần kinh; phải dùng nẹp cố định khớp trên và dưới chỗ gãy rồi đưa đến cơ sở y tế."
    },
    {
        "id": 90,
        "question": "Câu 90. Hệ nội tiết gồm các tuyến nội tiết tiết hormone vào máu và vận chuyển đến các cơ quan đích. Chức năng chủ yếu của hệ nội tiết ở người là gì?",
        "options": [
            "*A. Điều hòa các quá trình sinh lí của cơ thể thông qua hormone.* ",
            "B. Tiếp nhận và xử lí thông tin từ môi trường, điều khiển các phản xạ.",
            "C. Biến đổi thức ăn thành các chất dinh dưỡng để cơ thể hấp thụ.",
            "D. Vận chuyển oxygen và chất dinh dưỡng đến các tế bào trong cơ thể."
        ],
        "correct": "A",
        "solution": "Hệ nội tiết sản sinh hormone ngấm trực tiếp vào máu để điều hòa quá trình trao đổi chất, chuyển hóa năng lượng và sinh trưởng phát triển."
    }
]

ALL_PART1 = PART1_QUESTIONS + PART1_QUESTIONS_CONTINUED
