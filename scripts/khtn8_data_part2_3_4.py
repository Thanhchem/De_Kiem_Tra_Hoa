# -*- coding: utf-8 -*-
"""
Dữ liệu Phần II, Phần III và Phần Tự Luận - Đề cương KHTN 8 THCS Nguyễn Thị Minh Khai
"""

PART2_INFO = {
    "title": "Phần II. Câu trắc nghiệm đúng - sai",
    "subtitle": "Mỗi câu hỏi có 4 ý a, b, c, d - học sinh chọn đúng hoặc sai cho mỗi ý.",
    "req": [
        "YÊU CẦU CẦN ĐẠT:",
        "- Nêu được dung dịch là hỗn hợp lỏng đồng nhất của các chất đã tan trong nhau.",
        "- Nêu được định nghĩa độ tan của một chất trong nước, nồng độ phần trăm, nồng độ mol.",
        "- Tính được độ tan, nồng độ phần trăm, nồng độ mol theo công thức.",
        "- Tiến hành được thí nghiệm pha một dung dịch theo một nồng độ cho trước."
    ],
    "questions": [
        {
            "id": 1,
            "context": "Câu 1. Một nhóm học sinh được yêu cầu nhận biết dung dịch trong 4 cốc khác nhau: cốc 1 chứa nước muối (NaCl) đã tan hoàn toàn; cốc 2 chứa nước và đường (đã khuấy tan); cốc 3 chứa nước và cát (chưa tan); cốc 4 chứa nước và bột thạch cao (rải rác, không tan hoàn toàn). Họ thắc mắc: “Cốc nào thực sự là dung dịch?”",
            "items": [
                ("a. Dung dịch là hỗn hợp đồng nhất, nên cốc 1 (nước muối) và cốc 2 (nước đường) đều là dung dịch.", True, "Đúng. Cả nước muối và nước đường sau khi hòa tan hoàn toàn đều tạo thành hỗn hợp đồng nhất của chất tan và dung môi lỏng."),
                ("b. Cốc 3 và cốc 4 không phải là dung dịch vì cát và bột thạch cao không tan, tạo thành hỗn hợp không đồng nhất.", True, "Đúng. Cát và thạch cao không tan hoàn toàn trong nước tạo huyền phù/hỗn hợp không đồng nhất."),
                ("c. Dung dịch vẫn có thể chứa chất rắn không tan lắng ở đáy, miễn là phần trên trong suốt.", False, "Sai. Chất rắn lắng ở đáy không thuộc về dung dịch; dung dịch chỉ tính phần chất tan đã phân tán đều ở cấp độ phân tử trong dung môi."),
                ("d. Nếu không nhìn thấy “cặn”, chắc chắn cốc chứa dung dịch bão hoà.", False, "Sai. Dung dịch không có cặn có thể là dung dịch chưa bão hòa (vẫn có thể hòa tan thêm chất tan ở nhiệt độ đó).")
            ]
        },
        {
            "id": 2,
            "context": "Câu 2. Học sinh A pha dung dịch urea (CO(NH~2~)~2~) 0,02 M (dung dịch A), học sinh B pha dung dịch urea 0,1 M (dung dịch B). Sau đó trộn lẫn 2 lít dung dịch A với 3 lít dung dịch B, thu được 5 lít dung dịch C. Có bạn nói: “Cứ cộng nồng độ mol của A và B rồi chia đôi để ra nồng độ C”. Bạn khác phản đối: “Phải tính số mol thực tế chứ.”",
            "items": [
                ("a. Nồng độ mol (C~M~) là số mol chất tan trong 1 lít dung dịch.", True, "Đúng. Định nghĩa chuẩn của nồng độ mol: C~M~ = n / V (mol/L)."),
                ("b. Nếu ta pha hai dung dịch có nồng độ mol khác nhau, khi trộn lẫn, không thể chỉ “cộng và chia bình quân” nồng độ mà cần dựa trên tổng số mol tan và tổng thể tích.", True, "Đúng. Phải tính: n~tổng~ = n~A~ + n~B~ = 2 . 0,02 + 3 . 0,1 = 0,34 mol => C~M~ = 0,34 / 5 = 0,068 M (không thể lấy (0,02 + 0,1)/2 = 0,06 M)."),
                ("c. Nếu hai dung dịch A, B cùng chất tan, pha lẫn không xảy ra phản ứng hoá học, thì số mol chất tan chỉ bằng tổng số mol ban đầu.", True, "Đúng. Định luật bảo toàn lượng chất tan: n~tổng~ = n~A~ + n~B~."),
                ("d. Thể tích dung dịch tổng (5 lít) luôn bằng thể tích A (2 lít) + thể tích B (3 lít), bất kể có hoặc không có phản ứng.", False, "Sai. Về mặt lí thuyết hóa học, khi pha trộn các dung dịch có thể có sự co hoặc dãn thể tích nhẹ do lực tương tác phân tử, không phải lúc nào cũng 'luôn bằng bất kể có phản ứng'.")
            ]
        },
        {
            "id": 3,
            "context": "Câu 3. Trong một tiết thực hành Khoa học tự nhiên, nhóm học sinh tiến hành xác định khối lượng riêng của một khối hộp chữ nhật bằng cách đo kích thước và khối lượng của khối hộp. Dùng thước, nhóm đo được các cạnh của khối hộp lần lượt là 2 cm, 4 cm và 5 cm. Dùng cân, nhóm xác định được khối lượng của khối hộp là 108 g.\nDựa vào kết quả thực hành và kiến thức đã học, hãy xác định các nhận định sau là Đúng hay Sai:",
            "items": [
                ("a) Khối lượng riêng của khối hộp cho ta biết khối lượng của khối hộp đó.", False, "Sai. Khối lượng riêng cho ta biết khối lượng của 1 đơn vị thể tích (1 m^3^ hoặc 1 cm^3^) của chất đó, không phải tổng khối lượng của cả vật thể."),
                ("b) Khối lượng riêng được tính bằng công thức D = m . V.", False, "Sai. Công thức đúng là D = m / V (khối lượng chia cho thể tích)."),
                ("c) Thể tích của hộp là 40 cm^3^.", True, "Đúng. V = 2 . 4 . 5 = 40 cm^3^."),
                ("d) Khối lượng riêng của khối hộp 2700 kg/m^3^.", True, "Đúng. D = 108 / 40 = 2,7 g/cm^3^ = 2700 kg/m^3^ (đây là khối lượng riêng của nhôm).")
            ]
        }
    ]
}

PART3_INFO = {
    "title": "Phần III. Câu trắc nghiệm trả lời ngắn",
    "subtitle": "Mỗi câu hỏi, học sinh ghi câu trả lời không quá 4 kí tự",
    "questions": [
        {
            "id": 1,
            "question": "Câu 1. Cho các phát biểu sau:\n(1) PTHH dùng để biểu diễn ngắn gọn phản ứng hóa học bằng công thức hóa học.\n(2) PTHH biểu diễn các chất bằng tên gọi chữ viết đầy đủ.\n(3) Sơ đồ phản ứng chưa cân bằng số nguyên tử của các nguyên tố đã được gọi là PTHH.\n(4) PTHH hoàn chỉnh bắt buộc phải có hệ số cân bằng phù hợp ở các chất.\nHãy liệt kê các phát biểu ĐÚNG theo thứ tự từ nhỏ đến lớn.",
            "ans": "1, 4 (hoặc 14)",
            "solution": "Phát biểu đúng là (1) và (4). (2) sai vì dùng công thức hóa học; (3) sai vì sơ đồ chưa cân bằng chưa phải là phương trình hóa học."
        },
        {
            "id": 2,
            "question": "Câu 2. Khi tiến hành thí nghiệm cho dung dịch BaCl~2~ tác dụng với dung dịch Na~2~SO~4~ trên cân điện tử, các hiện tượng/kết quả quan sát được gồm:\n(1) Xuất hiện chất kết tủa màu trắng (BaSO~4~) trong cốc dung dịch.\n(2) Có bọt khí thoát ra mạnh làm dung dịch trào ra ngoài.\n(3) Kim cân/số hiển thị trên cân điện tử trước và sau phản ứng giữ nguyên không đổi.\n(4) Số hiển thị trên cân điện tử sau phản ứng giảm đi rõ rệt.\nHãy liệt kê các hiện tượng/kết quả ĐÚNG theo thứ tự từ nhỏ đến lớn.",
            "ans": "1, 3 (hoặc 13)",
            "solution": "Phản ứng: BaCl~2~ + Na~2~SO~4~ -> BaSO~4~ ↓ (trắng) + 2NaCl. Không sinh khí nên khối lượng toàn bộ hệ bảo toàn nguyên vẹn, số chỉ cân không đổi. Đúng: (1) và (3)."
        },
        {
            "id": 3,
            "question": "Câu 3. Cho phương trình hóa học: Zn + 2HCl → ZnCl~2~ + H~2~ ↑. Dựa vào ý nghĩa của phương trình hóa học, tỉ lệ số nguyên tử Zinc (Zn) tham gia phản ứng với số phân tử khí Hydrogen (H~2~) tạo thành là bao nhiêu? (Viết dạng tỉ lệ tối giản số : ......).",
            "ans": "1 : 1",
            "solution": "Hệ số của Zn là 1 và của H~2~ là 1, tỉ lệ số nguyên tử Zn : số phân tử H~2~ là 1 : 1."
        },
        {
            "id": 4,
            "question": "Câu 4. Trong phòng thí nghiệm hóa học, trước khi tiến hành một phản ứng, nhà nghiên cứu thường biểu diễn phản ứng bằng sơ đồ phản ứng rồi lập phương trình hóa học (chemical equation). Cho các bước:\n(1) Viết sơ đồ phản ứng\n(2) kiểm tra số nguyên tử của mỗi nguyên tố ở hai vế\n(3) cân bằng số nguyên tử bằng cách đặt hệ số thích hợp\n(4) hoàn chỉnh phương trình hóa học\nHãy sắp xếp theo trình tự hợp lí.",
            "ans": "1234",
            "solution": "Trình tự chuẩn các bước lập PTHH: (1) Viết sơ đồ -> (2) Đếm/kiểm tra số nguyên tử -> (3) Đặt hệ số thích hợp cân bằng -> (4) Hoàn chỉnh phương trình (thay mũi tên nét đứt bằng nét liền)."
        },
        {
            "id": 5,
            "question": "Câu 5. Trong nghiên cứu về hydrogen (H~2~) như một nguồn năng lượng, phản ứng giữa hydrogen và oxygen tạo thành nước: 2H~2~ + O~2~ → 2H~2~O. Cho:\n(1) 2 phân tử H~2~\n(2) phản ứng với\n(3) 1 phân tử O~2~\n(4) tạo thành 2 phân tử H~2~O\nHãy sắp xếp theo đúng ý nghĩa của phương trình.",
            "ans": "1234",
            "solution": "Ý nghĩa phương trình: (1) 2 phân tử H~2~ -> (2) phản ứng với -> (3) 1 phân tử O~2~ -> (4) tạo thành 2 phân tử H~2~O."
        },
        {
            "id": 6,
            "question": "Câu 6. Trong một thí nghiệm minh họa phản ứng hóa học, iron (Fe) phản ứng với sulfur (S) tạo thành iron(II) sulfide (FeS). Cho: (1) Fe; (2) + S; (3) →; (4) FeS.\nHãy sắp xếp để tạo phương trình hóa học đúng.",
            "ans": "1234",
            "solution": "Phương trình hoàn chỉnh: (1) Fe -> (2) + S -> (3) → -> (4) FeS."
        },
        {
            "id": 7,
            "question": "Câu 7. Magnesium (Mg) là kim loại được nghiên cứu trong vật liệu nhẹ. Khi magnesium cháy trong oxygen tạo magnesium oxide (MgO): 2Mg + O~2~ → 2MgO.\nHãy ghi các hệ số theo đúng thứ tự Mg, O~2~, MgO.",
            "ans": "2, 1, 2",
            "solution": "Hệ số tương ứng trước Mg là 2, trước O~2~ là 1, trước MgO là 2."
        },
        {
            "id": 8,
            "question": "Câu 8. Trong phòng thí nghiệm, học sinh cho acetic acid (CH~3~COOH) trong giấm ăn tác dụng với sodium hydrogen carbonate (NaHCO~3~). CO~2~ sinh ra làm bóng bay phồng lên.\n[HÌNH ẢNH: Bình tam giác nối bóng bay phồng to]\nChọn đáp án đúng và ghép a–b–c–d ở cột (I) với các số tương ứng ở cột (II):",
            "has_table": True,
            "table_data": [
                ("Cột I", "Cột II"),
                ("a) Trong hệ kín, tổng khối lượng trước và sau phản ứng", "(1) tăng."),
                ("b) Bóng bay phồng lên nhưng khối lượng toàn bộ hệ kín:", "(2) giảm."),
                ("c) Trong hệ hở, nếu CO~2~ thoát ra ngoài, khối lượng phần hệ được cân có thể:", "(3) không thay đổi."),
                ("d) Khối lượng dung sau phản ứng thay đổi như thế nào so với trước phản ứng", "(4) không xác định.")
            ],
            "ans": "3323",
            "solution": "a ghép (3) vì hệ kín bảo toàn khối lượng; b ghép (3) vì bóng giữ kín khí CO~2~ không thoát ra nên khối lượng cả hệ không đổi; c ghép (2) vì khí bay đi làm phần còn lại nhẹ hơn; d ghép (3) vì theo định luật bảo toàn khối lượng chất tham gia = sản phẩm."
        },
        {
            "id": 9,
            "question": "Câu 9. Dùng đòn bẩy loại 1 để nâng một vật nặng có trọng lượng P = 5000 N. Khoảng cách từ điểm tựa O đến điểm đặt vật O~1~ là OO~1~ = 0,2 m; khoảng cách từ điểm tựa O đến điểm tác dụng lực O~2~ là OO~2~ = 0,8 m. Tính độ lớn lực F (theo đơn vị Niutơn - N) tối thiểu cần tác dụng vào O~2~ để nâng vật.",
            "ans": "1250",
            "solution": "F . OO~2~ = P . OO~1~ => F = 5000 . 0,2 / 0,8 = 1250 N."
        },
        {
            "id": 10,
            "question": "Câu 10. Khi sử dụng đòn bẩy loại 1 để bẩy một tảng đá, người ta thấy tỉ số cánh tay đòn OO~1~ / OO~2~ = 1/4. Tỉ số giữa lực tác dụng F~2~ cần thiết so với trọng lượng P của tảng đá (tức F~2~ / P) dưới dạng số thập phân bằng bao nhiêu?",
            "ans": "0,25",
            "solution": "F~2~ / P = OO~1~ / OO~2~ = 1/4 = 0,25."
        },
        {
            "id": 11,
            "question": "Câu 11. Để nâng một vật nặng P = 200N ở điểm O~1~, người ta tác dụng lực F~2~ = 50N vào điểm O~2~ cách điểm tựa O một khoảng OO~2~ = 1,0 m. Khoảng cách OO~1~ (mét) từ điểm tựa đến vật bằng bao nhiêu?",
            "ans": "0,25",
            "solution": "OO~1~ = F~2~ . OO~2~ / P = 50 . 1,0 / 200 = 0,25 m."
        },
        {
            "id": 12,
            "question": "Câu 12. Để nâng một cối đá có trọng lượng P = 8000 N, công nhân tác dụng lực F = 2000 N vào đầu O~2~ của đòn bẩy. Nếu cánh tay đòn đặt vật OO~1~ = 0,5 m thì khoảng cánh tay đòn OO~2~ (mét) phải bằng bao nhiêu?",
            "ans": "2 (hoặc 2,0)",
            "solution": "OO~2~ = P . OO~1~ / F = 8000 . 0,5 / 2000 = 2,0 m."
        },
        {
            "id": 13,
            "question": "Câu 13. Cho các dụng cụ sau đây:\n1) Kéo cắt giấy.\n2) Dùng xà beng bẩy khúc gỗ.\n3) Em bé dùng chổi quét nhà.\n4) Cái khui nút chai.\n5) Cần câu cá.\n6) Đôi quang gánh.\n7) Cái kẹp vỏ hạt.\nNhững dụng cụ nào là đòn bẩy loại 1?",
            "ans": "1, 2, 6",
            "solution": "Đòn bẩy loại 1 có điểm tựa O nằm giữa điểm đặt lực tác dụng và điểm đặt vật tải: 1) Kéo cắt giấy, 2) Xà beng bẩy gỗ, 6) Đôi quang gánh (vai người là điểm tựa giữa 2 gánh)."
        },
        {
            "id": 14,
            "question": "Câu 14. Một thanh nhôm đặc có thể tích 500cm^3^. Biết khối lượng riêng của nhôm là 2700kg/m^3^. Tính khối lượng của thanh nhôm theo đơn vị kilôgam (kg). (Ghi đáp số là một số thập phân)",
            "ans": "1,35",
            "solution": "V = 500 cm^3^ = 0,0005 m^3^. Khối lượng m = D . V = 2700 . 0,0005 = 1,35 kg."
        },
        {
            "id": 15,
            "question": "Câu 15. Một chiếc cờ lê có chiều dài cán 25cm. Người thợ sửa xe tác dụng một lực 80N vuông góc với cán cờ lê tại mút cán để siết chặt bu-lông. Tính moment của lực tác dụng theo đơn vị N·m.",
            "ans": "20",
            "solution": "d = 25 cm = 0,25 m. Moment lực M = F . d = 80 . 0,25 = 20 N·m."
        },
        {
            "id": 16,
            "question": "Câu 16. Một thanh đòn bẩy có trục quay tại điểm tựa O. Ở một đầu thanh cách O một đoạn 15cm, người ta treo một quả nặng 60N. Để đòn bẩy nằm cân bằng ngang, ở đầu kia người ta tác dụng một lực vuông góc có độ lớn 20N. Hỏi điểm đặt của lực này phải cách điểm tựa O một khoảng bằng bao nhiêu xentimét (cm)?",
            "ans": "45",
            "solution": "F~1~ . d~1~ = F~2~ . d~2~ => 60 . 15 = 20 . d~2~ => d~2~ = 900 / 20 = 45 cm."
        },
        {
            "id": 17,
            "question": "Câu 17. Một khối hộp kim loại có kích thước ba cạnh lần lượt là 10cm, 20cm và 40cm. Khối lượng riêng của kim loại là 2500kg/m^3^. Lấy P = 10m. Tính áp suất nhỏ nhất mà khối hộp này có thể tác dụng lên mặt bàn nằm ngang theo đơn vị Pa.",
            "ans": "2500",
            "solution": "V = 0,1 . 0,2 . 0,4 = 0,008 m^3^. m = 2500 . 0,008 = 20 kg => P = 200 N. Áp suất nhỏ nhất khi diện tích tiếp xúc lớn nhất: S_max = 0,2 . 0,4 = 0,08 m^2^ => p_min = 200 / 0,08 = 2500 Pa."
        },
        {
            "id": 18,
            "question": "Câu 18. Khi sử dụng đòn bẩy, lực tác dụng lên vật có thể được giảm đi bao nhiêu lần nếu cánh tay đòn của lực tác dụng dài gấp 2 lần cánh tay đòn của vật?",
            "ans": "2",
            "solution": "Vì F . d~2~ = P . d~1~ => F / P = d~1~ / d~2~ = 1/2. Lực tác dụng giảm đi 2 lần."
        },
        {
            "id": 19,
            "question": "Câu 19. Một người dùng đòn bẩy để nâng vật có trọng lượng 300 N. Cánh tay đòn của lực tác dụng dài gấp 3 lần cánh tay đòn của vật. Bỏ qua ma sát, lực người cần tác dụng là bao nhiêu niutơn?",
            "ans": "100",
            "solution": "F = P / 3 = 300 / 3 = 100 N."
        },
        {
            "id": 20,
            "question": "Câu 20. Một người dùng xà beng để bẩy một tảng đá. Cánh tay đòn của lực tác dụng dài 120 cm, cánh tay đòn của tảng đá dài 30 cm. So với lực do tảng đá tác dụng, người đó cần tác dụng lực nhỏ hơn bao nhiêu lần?",
            "ans": "4",
            "solution": "Tỉ số cánh tay đòn = 120 / 30 = 4 lần. Lực tác dụng nhỏ hơn 4 lần."
        },
        {
            "id": 21,
            "question": "Câu 21. Một lực có độ lớn 50 N tác dụng vuông góc vào một cánh tay đòn dài 40 cm. Moment lực đối với trục quay có độ lớn bao nhiêu N·m?",
            "ans": "20",
            "solution": "d = 40 cm = 0,4 m. Moment lực M = F . d = 50 . 0,4 = 20 N·m."
        },
        {
            "id": 22,
            "question": "Câu 22. Một người dùng một thanh sắt dài 120 cm để nâng một vật nặng. Điểm tựa cách vật 30 cm. Người tác dụng lực ở đầu kia của thanh. Bỏ qua khối lượng thanh và ma sát. Nếu vật tác dụng lực 400 N lên đòn bẩy, người cần tác dụng lực bao nhiêu N để nâng vật?",
            "ans": "133 (hoặc 133,3)",
            "solution": "Cánh tay đòn của vật: d~1~ = 30 cm. Cánh tay đòn của người: d~2~ = 120 - 30 = 90 cm. F . 90 = 400 . 30 => F = 12000 / 90 ≈ 133,3 N."
        },
        {
            "id": 23,
            "question": "Câu 23. Một người dùng xà beng để nâng một tảng đá. Khoảng cách từ điểm tựa đến tảng đá là 20 cm, khoảng cách từ điểm tựa đến vị trí tác dụng lực của người là 100 cm. Nếu tảng đá tác dụng lực 600 N lên xà beng, người cần tác dụng lực bao nhiêu N để bắt đầu nâng tảng đá?",
            "ans": "120",
            "solution": "F . 100 = 600 . 20 => F = 12000 / 100 = 120 N."
        }
    ]
}

PART4_ESSAY = [
    {
        "id": 1,
        "question": "Câu 1. Em hãy phân biệt sự khác nhau bản chất giữa sự biến đổi vật lí và sự biến đổi hóa học. Mỗi loại hãy đưa ra 02 ví dụ thực tế trong đời sống hằng ngày để minh họa.",
        "solution": """• Bản chất khác nhau:
- Biến đổi vật lí: Là sự biến đổi về trạng thái, hình dạng, kích thước... nhưng không có sự tạo thành chất mới.
- Biến đổi hóa học: Là sự biến đổi có kèm theo sự tạo thành chất mới (liên kết giữa các nguyên tử thay đổi).
• Ví dụ minh họa:
- Biến đổi vật lí:
  + Nước đá tan chảy thành nước lỏng; nước lỏng đun sôi bốc thành hơi nước.
  + Xé nhỏ tờ giấy hoặc hòa tan đường vào nước.
- Biến đổi hóa học:
  + Đốt cháy than củi tạo thành khí carbon dioxide và tro.
  + Đinh sắt để ngoài không khí ẩm bị gỉ sét thành màu nâu đỏ."""
    },
    {
        "id": 2,
        "question": "Câu 2. Trong đời sống và trong phòng thí nghiệm, làm thế nào để chúng ta nhận biết được có phản ứng hóa học xảy ra? Hãy nêu 4 dấu hiệu nhận biết phổ biến và đưa ra 1 ví dụ thực tế tương ứng với một trong các dấu hiệu đó.",
        "solution": """• 4 dấu hiệu nhận biết phản ứng hóa học xảy ra:
1. Có sự thay đổi màu sắc của các chất.
2. Xuất hiện chất kết tủa (chất không tan).
3. Có chất khí thoát ra (sủi bọt khí).
4. Có sự tỏa nhiệt hoặc phát sáng.
• Ví dụ minh họa:
Khi cho viên C sủi hoặc NaHCO~3~ (baking soda) vào nước giấm ăn, ta quan sát thấy có rất nhiều bọt khí sủi lên mạnh mẽ (sinh ra khí carbon dioxide CO~2~)."""
    },
    {
        "id": 3,
        "question": "Câu 3. Trong giờ thực hành Khoa học tự nhiên 8, hai nhóm học sinh tiến hành hai thí nghiệm sau:\n+ Nhóm 1: Cho một ít muối ăn (NaCl) vào cốc chứa 50 mL nước cất rồi khuấy đều đến khi muối tan hết thu được dung dịch trong suốt.\n+ Nhóm 2: Nhỏ vài giọt dung dịch hydrochloric acid (HCl) vào ống nghiệm có chứa một viên kẽm (Zn).\na) Xác định thí nghiệm nào xảy ra sự biến đổi vật lí, thí nghiệm nào xảy ra sự biến đổi hóa học? Giải thích.\nb) Ở thí nghiệm có biến đổi hóa học, em hãy chỉ ra dấu hiệu giúp quan sát được bằng mắt thường để biết phản ứng đã xảy ra.",
        "solution": """a) Phân loại và giải thích:
- Thí nghiệm của Nhóm 1 là biến đổi vật lí: Muối ăn chỉ hòa tan trong nước phân tán thành các ion Na^+^ và Cl^-^ tạo dung dịch đồng nhất, khi cô cạn vẫn thu lại muối ăn NaCl, không sinh ra chất mới.
- Thí nghiệm của Nhóm 2 là biến đổi hóa học: Xảy ra phản ứng Zn + 2HCl -> ZnCl~2~ + H~2~ ↑, kim loại kẽm bị tan dần và sinh ra chất mới là muối zinc chloride và khí hydrogen.
b) Dấu hiệu quan sát bằng mắt thường:
- Viên kẽm tan dần.
- Có bọt khí không màu sủi lên và thoát ra từ bề mặt viên kẽm (khí H~2~)."""
    },
    {
        "id": 4,
        "question": "Câu 4. a) Thế nào là phản ứng tỏa nhiệt, thế nào là phản ứng thu nhiệt?\nb) Trong hai trường hợp sau, trường hợp nào là phản ứng tỏa nhiệt, trường hợp nào là phản ứng thu nhiệt?\n+ Trường hợp A: Đốt cháy than (C) hoặc khí gas để đun nấu thức ăn.\n+ Trường hợp B: Cho một viên C sủi hoặc nhiệt kế lạnh/túi chườm lạnh tự ngắt hòa tan vào nước làm nhiệt độ dung dịch giảm xuống.\nc) Hãy đưa ra thêm 01 ví dụ minh họa trong thực tế về phản ứng tỏa nhiệt có ứng dụng phục vụ đời sống.",
        "solution": """a) Định nghĩa:
- Phản ứng tỏa nhiệt: Là phản ứng hóa học giải phóng năng lượng dưới dạng nhiệt ra môi trường xung quanh (làm nhiệt độ môi trường tăng lên).
- Phản ứng thu nhiệt: Là phản ứng hóa học hấp thụ năng lượng dưới dạng nhiệt từ môi trường xung quanh (làm nhiệt độ môi trường giảm xuống).
b) Phân loại trường hợp:
- Trường hợp A (Đốt cháy than hoặc gas): Là phản ứng tỏa nhiệt (giải phóng nhiệt lớn để đun nấu).
- Trường hợp B (Hòa tan viên sủi / túi chườm lạnh làm giảm nhiệt độ): Là quá trình thu nhiệt.
c) Ví dụ thực tế về phản ứng tỏa nhiệt:
- Phản ứng đốt cháy xăng dầu trong động cơ xe máy, ô tô cung cấp năng lượng để vận hành phương tiện.
- Phản ứng tôi vôi: CaO + H~2~O -> Ca(OH)~2~ tỏa lượng nhiệt rất lớn làm nước sôi sùng sục."""
    },
    {
        "id": 5,
        "question": "Câu 5. a) Nêu định nghĩa khối lượng riêng của một chất.\nb) Liệt kê các đơn vị đo khối lượng riêng thường dùng và ghi rõ mối liên hệ giữa đơn vị g/cm^3^ và kg/m^3^.",
        "solution": """a) Định nghĩa khối lượng riêng:
Khối lượng riêng của một chất là khối lượng của một đơn vị thể tích chất đó:
D = m / V (trong đó D là khối lượng riêng, m là khối lượng, V là thể tích).
b) Đơn vị đo:
- Các đơn vị thường dùng: kilôgam trên mét khối (kg/m^3^) và gam trên xentimét khối (g/cm^3^).
- Mối liên hệ: 1 g/cm^3^ = 1 000 kg/m^3^ (hay 1 kg/m^3^ = 0,001 g/cm^3^)."""
    },
    {
        "id": 6,
        "question": "Câu 6. Một khối gang hình hộp chữ nhật có chiều dài các cạnh tương ứng là a = 2 cm, b = 3 cm, c = 5 cm và có khối lượng m = 210 g.\na) Tính thể tích V và khối lượng riêng D của khối gang theo đơn vị g/cm^3^.\nb) Đổi giá trị khối lượng riêng D của gang vừa tìm được sang đơn vị chuẩn kg/m^3^.",
        "solution": """a) Thể tích của khối gang:
V = a . b . c = 2 . 3 . 5 = 30 cm^3^.
Khối lượng riêng của khối gang theo g/cm^3^:
D = m / V = 210 / 30 = 7 g/cm^3^.
b) Đổi sang đơn vị chuẩn kg/m^3^:
D = 7 . 1000 = 7 000 kg/m^3^."""
    },
    {
        "id": 7,
        "question": "Câu 7. a) Viết công thức tính áp suất p = F/S và nêu rõ tên, đơn vị của các đại lượng.\nb) Dựa vào công thức tính áp suất, hãy trình bày các nguyên tắc cơ bản để làm tăng áp suất và làm giảm áp suất trên một bề mặt bị ép.\nc) Nêu 2 ví dụ thực tế về việc làm tăng áp suất và 2 ví dụ về làm giảm áp suất trong đời sống sinh hoạt.",
        "solution": """a) Công thức tính áp suất:
p = F / S
Trong đó:
- p là áp suất, đơn vị là paxcan (Pa) hoặc N/m^2^ (1 Pa = 1 N/m^2^).
- F là áp lực tác dụng vuông góc lên bề mặt bị ép, đơn vị là niutơn (N).
- S là diện tích bề mặt bị ép, đơn vị là mét vuông (m^2^).
b) Nguyên tắc làm tăng/giảm áp suất:
- Để làm tăng áp suất (p tăng): Tăng áp lực F, hoặc giảm diện tích tiếp xúc S, hoặc kết hợp cả hai.
- Để làm giảm áp suất (p giảm): Giảm áp lực F, hoặc tăng diện tích tiếp xúc S, hoặc kết hợp cả hai.
c) Ví dụ thực tế:
- Tăng áp suất:
  + Mài sắc lưỡi dao, lưỡi kéo để giảm diện tích tiếp xúc, giúp cắt gọt đồ vật dễ dàng hơn.
  + Đóng đinh có đầu nhọn hoắt để dễ đóng sâu vào gỗ/tường.
- Giảm áp suất:
  + Xe tăng, máy kéo dùng bánh xích bản rộng để tăng diện tích tiếp xúc, giảm áp suất lên mặt đất tránh lún lầy.
  + Đặt tà vẹt gỗ/bê tông bản rộng dưới đường ray tàu hỏa để giảm áp suất tác dụng xuống nền đất."""
    },
    {
        "id": 8,
        "question": "Câu 8. Dựa vào kiến thức áp suất trên một bề mặt bị ép, hãy giải thích cơ sở khoa học của các thiết bị kỹ thuật và vật dụng sau:\na) Tại sao mũi đinh, đầu kim khâu, mũi khoan, lưỡi dao, xẻng xúc đất lại được mài nhọn hoặc làm rất mỏng?\nb) Tại sao máy xúc, xe tăng có trọng lượng rất lớn lại được thiết kế chạy bằng hệ thống bánh xích rộng, còn đường ray tàu hỏa lại được đặt trên các thanh tà vẹt to?",
        "solution": """a) Mũi đinh, kim khâu, lưỡi dao được mài nhọn, mỏng:
Nhằm làm giảm diện tích tiếp xúc S đến mức rất nhỏ. Theo công thức p = F / S, với cùng một lực tác dụng F vừa phải, diện tích S càng nhỏ thì áp suất p sinh ra càng lớn, giúp đâm xuyên hoặc cắt gọt vật liệu dễ dàng.
b) Xe tăng dùng bánh xích rộng, đường ray đặt trên thanh tà vẹt to:
Vì xe tăng, tàu hỏa có trọng lượng P (áp lực F) cực kì lớn. Việc thiết kế bản xích rộng và lót thanh tà vẹt to có diện tích tiếp xúc S rất lớn, giúp giảm áp suất p = F / S tác dụng lên nền đất/mặt đường, ngăn không cho nền đất bị sụt lún, biến dạng hay gãy đường ray."""
    },
    {
        "id": 9,
        "question": "Câu 9. Một xe tăng có trọng lượng P~1~ = 350 000 N, tổng diện tích tiếp xúc của các bản xích với mặt đất nằm ngang là S~1~ = 1,5 m^2^. Một ô tô có trọng lượng P~2~ = 25 000 N, tổng diện tích tiếp xúc của các bánh xe với mặt đường nằm ngang là S~2~ = 250 cm^2^.\na) Tính áp suất p~1~ của xe tăng và áp suất p~2~ của ô tô tác dụng lên mặt đường.\nb) So sánh p~1~ với p~2~ và giải thích tại sao xe tăng nặng hơn ô tô rất nhiều lần nhưng lại không bị lún sâu và di chuyển dễ dàng trên nền đất mềm lầy lội?",
        "solution": """a) Tính áp suất:
- Đổi diện tích tiếp xúc của ô tô: S~2~ = 250 cm^2~ = 0,025 m^2^.
- Áp suất của xe tăng:
  p~1~ = P~1~ / S~1~ = 350 000 / 1,5 ≈ 233 333,3 Pa.
- Áp suất của ô tô:
  p~2~ = P~2~ / S~2~ = 25 000 / 0,025 = 1 000 000 Pa.
b) So sánh và giải thích:
Ta thấy p~1~ < p~2~ (khoảng hơn 4 lần). Mặc dù xe tăng có trọng lượng nặng gấp 14 lần ô tô, nhưng diện tích tiếp xúc của hệ thống bánh xích xe tăng lại lớn gấp 60 lần diện tích bánh xe ô tô, dẫn đến áp suất của xe tăng tác dụng lên nền đất nhỏ hơn nhiều so với ô tô. Do đó xe tăng không bị lún sâu và chạy tốt trên đất lầy."""
    },
    {
        "id": 10,
        "question": "Câu 10. Một khối chất rắn đồng chất có dạng hình hộp chữ nhật với kích thước các cạnh lần lượt là a = 10cm, b = 20cm và c = 30cm. Đặt khối chất rắn lên cân điện tử, số chỉ của cân là 16,2kg.\na) Nêu định nghĩa khối lượng riêng của một chất. Viết công thức tính khối lượng riêng, nêu rõ tên và đơn vị hợp pháp của các đại lượng có trong công thức.\nb) Tính khối lượng riêng của chất làm nên khối hộp trên theo đơn vị kg/m^3^\nc) Đặt khối hộp nằm yên trên mặt sàn nằm ngang. Để áp suất của khối hộp tác dụng lên mặt sàn là nhỏ nhất, ta phải đặt mặt đáy nào tiếp xúc với mặt sàn? Tính giá trị áp suất nhỏ nhất đó (lấy P = 10m).",
        "solution": """a) Định nghĩa & công thức:
Khối lượng riêng của một chất là khối lượng của một đơn vị thể tích chất đó: D = m / V.
- D: Khối lượng riêng (kg/m^3^).
- m: Khối lượng của vật (kg).
- V: Thể tích của vật (m^3^).
b) Tính khối lượng riêng:
Thể tích khối hộp: V = 0,1 . 0,2 . 0,3 = 0,006 m^3^.
Khối lượng riêng: D = m / V = 16,2 / 0,006 = 2 700 kg/m^3^ (khối nhôm).
c) Áp suất nhỏ nhất:
Trọng lượng khối hộp: P = 10 . m = 10 . 16,2 = 162 N.
Để áp suất p = P / S nhỏ nhất thì diện tích tiếp xúc S phải lớn nhất.
Mặt đáy có diện tích lớn nhất là mặt có kích thước 20cm x 30cm (S_max = 0,2 . 0,3 = 0,06 m^2^).
Áp suất nhỏ nhất là:
p_min = P / S_max = 162 / 0,06 = 2 700 Pa."""
    },
    {
        "id": 11,
        "question": "Câu 11. Đặt lực F có độ lớn bằng nhau tại các vị trí khác nhau trên cờ lê như hình dưới đây để lắp bu lông. Hình nào bu lông sẽ không quay, hình nào bu lông quay dễ dàng nhất? Vì sao?\n[HÌNH ẢNH: 3 cờ lê a), b), c)]",
        "has_image": True,
        "image_path": "San_Pham/images_khtn8/khtn8_wrenches.png",
        "solution": """• Hình bu lông sẽ không quay:
Là hình c), vì phương của lực tác dụng đi qua trục quay (tâm của bu lông), cánh tay đòn bằng 0 nên moment lực bằng 0, không có tác dụng làm quay.
• Hình bu lông quay dễ dàng nhất:
Là hình a), vì lực tác dụng vuông góc với cán cờ lê và đặt ở vị trí mút ngoài cùng của cán (cách xa trục quay nhất), làm cho cánh tay đòn d đạt giá trị lớn nhất => Moment lực M = F . d là lớn nhất, tạo tác dụng làm quay mạnh nhất."""
    },
    {
        "id": 12,
        "question": "Câu 12. Một người làm vườn cần đóng một chiếc cọc xuống đất. Hãy đề xuất phương án để có thể đóng được chiếc cọc xuống đất một cách dễ dàng. Giải thích.",
        "solution": """• Đề xuất phương án:
Vót nhọn đầu cọc tiếp xúc với mặt đất trước khi đóng, đồng thời dùng vồ hoặc búa có khối lượng thích hợp gõ thẳng góc vào đầu trên của cọc.
• Giải thích:
Vót nhọn đầu dưới làm giảm diện tích tiếp xúc S xuống rất nhỏ. Khi tác dụng lực đóng F từ trên xuống, áp suất p = F / S tác dụng lên bề mặt đất tại đầu nhọn sẽ rất lớn, dễ dàng thắng sức cản của đất và cọc ngập sâu xuống một cách nhẹ nhàng."""
    },
    {
        "id": 13,
        "question": "Câu 13. Để xe ô tô có thể vượt qua vùng đất sụt lún người ta thường làm như thế nào? Mô tả cách làm và giải thích.",
        "solution": """• Mô tả cách làm:
1. Chèn thêm các tấm ván gỗ to bản, cành cây, rơm rạ hoặc xích chống lầy lót dưới bánh xe.
2. Xả bớt một phần hơi trong lốp xe để lốp hơi xẹp xuống.
• Giải thích cơ sở khoa học:
- Cả hai cách trên đều nhằm mục đích làm tăng đáng kể diện tích tiếp xúc S giữa bánh xe với mặt đất.
- Theo công thức p = F / S, khi diện tích S tăng lên thì áp suất p do xe tác dụng lên nền đất giảm đi nhiều lần, giúp xe không bị lún sâu và tăng ma sát để vượt qua vùng sụt lún."""
    },
    {
        "id": 14,
        "question": "Câu 14. Trong quá trình vận động, cơ, xương và khớp phối hợp với nhau để giúp cơ thể thực hiện các động tác.\na) Giải thích mối quan hệ giữa cơ và xương trong quá trình vận động của cơ thể người.\nb) Một học sinh thường xuyên ngồi học sai tư thế. Em hãy nêu một hậu quả có thể xảy ra đối với hệ vận động và một biện pháp để phòng tránh.",
        "solution": """a) Mối quan hệ giữa cơ và xương:
Cơ bám vào xương thông qua các gân. Khi cơ co, nó sẽ tạo lực kéo xương chuyển động quanh khớp xương (đóng vai trò như một hệ đòn bẩy sinh học). Xương là bộ khung nâng đỡ thụ động, còn cơ là động lực chủ động tạo nên mọi vận động.
b) Hậu quả và biện pháp:
- Hậu quả: Dẫn đến tật cong vẹo cột sống (gù lưng, vẹo cột sống sang một bên), gây biến dạng lồng ngực, chèn ép tim phổi và thoái hóa xương khớp sớm.
- Biện pháp phòng tránh: Ngồi học ngay ngắn, giữ lưng thẳng, không tì ngực vào cạnh bàn, hai chân chạm sàn vuông góc, điều chỉnh bàn ghế có chiều cao phù hợp và tập thể dục giữa giờ."""
    },
    {
        "id": 15,
        "question": "Câu 15. Khi chơi bóng đá, một học sinh thực hiện động tác sút bóng. Ngay sau đó, em bị đau ở vùng khớp gối và phải dừng vận động.\nNếu học sinh trên vẫn cố tiếp tục chơi thể thao khi đang bị đau khớp gối, điều gì có thể xảy ra? Em hãy giải thích và đề xuất những việc cần làm để bảo vệ hệ vận động khi tham gia thể thao.",
        "solution": """• Hậu quả nếu cố tiếp tục chơi:
Khớp gối có thể bị tổn thương nặng hơn như giãn hoặc đứt dây chằng chéo, rách sụn chêm, tràn dịch khớp, biến dạng khớp và dẫn đến tàn phế hoặc phải phẫu thuật can thiệp kéo dài.
• Những việc cần làm để bảo vệ hệ vận động khi chơi thể thao:
1. Khởi động kĩ toàn thân và làm nóng các khớp (đặc biệt khớp cổ chân, khớp gối, cổ tay) trước khi vận động.
2. Sử dụng trang bị bảo hộ chuyên dụng (băng bảo vệ gối, giày thể thao vừa chân, có độ bám tốt).
3. Thi đấu, luyện tập vừa sức; khi có dấu hiệu chấn thương, đau khớp phải dừng ngay lập tức, chườm lạnh và đi khám bác sĩ chuyên khoa."""
    },
    {
        "id": 16,
        "question": "Câu 16. a) Trình bày nguyên nhân, hậu quả, cách phòng tránh bệnh loãng xương và tật cong vẹo cột sống.\nb) Calcium, vitamin D và nước có vai trò gì đối với hệ vận động?",
        "solution": """a) Bệnh loãng xương và tật cong vẹo cột sống:
• Bệnh loãng xương:
- Nguyên nhân: Chế độ ăn thiếu calcium, thiếu vitamin D, ít vận động ngoài trời, suy giảm hormone ở người cao tuổi.
- Hậu quả: Xương xốp, giòn, giảm khả năng chịu lực, rất dễ bị gãy dù chỉ va chạm nhẹ.
- Phòng tránh: Bổ sung đầy đủ thực phẩm giàu calcium (sữa, tôm, cua...), tắm nắng sáng sớm hấp thụ vitamin D, tập luyện thể thao thường xuyên.
• Tật cong vẹo cột sống:
- Nguyên nhân: Ngồi học sai tư thế, mang cặp sách quá nặng lệch một bên vai, lao động quá sức từ nhỏ.
- Hậu quả: Biến dạng cột sống, mất cân đối vóc dáng, ảnh hưởng đến chức năng tuần hoàn và hô hấp.
- Phòng tránh: Ngồi học đúng tư thế, đeo cặp 2 quai cân đối, không mang vác quá nặng.
b) Vai trò của calcium, vitamin D và nước:
- Calcium: Thành phần khoáng chất chính cấu tạo nên chất nền xương và răng, giúp xương chắc khỏe.
- Vitamin D: Hỗ trợ ruột hấp thụ tối đa calcium và phosphate vào máu và cố định vào mô xương.
- Nước: Thành phần quan trọng cấu tạo nên dịch khớp, giúp bôi trơn các khớp xương, giảm ma sát và nuôi dưỡng sụn khớp."""
    },
    {
        "id": 17,
        "question": "Câu 17. Em hãy lựa chọn một môn thể thao phù hợp với bản thân, xây dựng kế hoạch tập luyện và nêu những lợi ích mà hoạt động đó mang lại cho hệ vận động và sức khỏe.\nCHÚC CÁC EM THI TỐT!!!",
        "solution": """• Lựa chọn môn thể thao: Môn bơi lội (hoặc chạy bộ/cầu lông).
• Kế hoạch tập luyện:
- Tần suất: 3 - 4 buổi/tuần, mỗi buổi khoảng 45 - 60 phút vào các buổi chiều.
- Khởi động 10 phút trước khi xuống nước và thả lỏng cơ thể 5 phút sau khi bơi.
• Lợi ích đối với hệ vận động và sức khỏe:
1. Phát triển toàn diện hệ cơ và xương, kích thích xương dài ra giúp tăng chiều cao ở tuổi dậy thì.
2. Giảm tối đa áp lực trọng lượng lên các khớp xương nhờ sức nâng của nước, phòng ngừa thoái hóa và chấn thương khớp.
3. Tăng cường dung tích phổi, nâng cao sức bền tim mạch và tạo tinh thần sảng khoái sau giờ học."""
    }
]
