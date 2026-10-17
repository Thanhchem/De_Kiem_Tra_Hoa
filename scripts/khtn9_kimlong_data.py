# -*- coding: utf-8 -*-
"""
Dữ liệu Đề cương ôn tập Giữa kì I - KHTN 9
Trường THCS Kim Long - Năm học 2026 - 2027
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

INFO = {
    "school": "TRƯỜNG THCS KIM LONG",
    "year": "NĂM HỌC 2026 - 2027",
    "title": "ĐỀ CƯƠNG ÔN TẬP GIỮA KÌ I – KHTN 9",
    "subject": "MÔN: KHOA HỌC TỰ NHIÊN 9",
    "time_str": "Chương trình GDPT 2018 - Đầy đủ Phân môn Hóa học & Vật lí",
}

# -----------------------------------------------------------------------------
# BÀI 1: NHẬN BIẾT MỘT SỐ DỤNG CỤ, HÓA CHẤT. THUYẾT TRÌNH MỘT VẤN ĐỀ KHOA HỌC
# -----------------------------------------------------------------------------
BAI_1_QUESTIONS = [
    {
        "id": 1,
        "content": "Cho các dụng cụ sau: điện kế, nguồn sáng, đồng hồ đo điện đa năng, cuộn dây dẫn có hai đèn LED, bản bán trụ. Số dụng cụ thí nghiệm điện từ là",
        "options": ["A. 1", "B. 2", "C. 3", "D. 4"],
        "correct": "C",
        "solution": "Các dụng cụ dùng trong thí nghiệm điện từ gồm: điện kế, đồng hồ đo điện đa năng, cuộn dây dẫn có hai đèn LED (tổng cộng 3 dụng cụ).\nNguồn sáng và bản bán trụ là các dụng cụ dùng trong thí nghiệm quang học."
    },
    {
        "id": 2,
        "content": "Dụng cụ nào thường được sử dụng để phân tán nhiệt khi đun nóng?",
        "options": ["A. Đèn cồn", "B. Lưới tản nhiệt", "C. Giá ống nghiệm", "D. Kẹp gỗ"],
        "correct": "B",
        "solution": "Lưới tản nhiệt (lưới thép có phủ lớp amiăng/gốm chịu nhiệt) được đặt giữa ngọn lửa đèn cồn và đáy bình thủy tinh giúp nhiệt lượng phân tán đều khắp đáy bình, tránh nứt vỡ do chênh lệch nhiệt độ cục bộ."
    },
    {
        "id": 3,
        "content": "Dụng cụ nào sau đây dùng để pha chế dung dịch?",
        "options": ["A. Pipet", "B. Kẹp gỗ", "C. Phễu chiết", "D. Bình cầu"],
        "correct": "D",
        "solution": "Bình cầu (hoặc cốc thủy tinh, bình tam giác) thường được sử dụng để chứa, hòa tan, pha chế dung dịch và thực hiện các phản ứng cần đun nóng.\nPipet dùng để hút chính xác một lượng chất lỏng; Kẹp gỗ dùng kẹp ống nghiệm; Phễu chiết dùng tách các chất lỏng không trộn lẫn vào nhau."
    },
    {
        "id": 4,
        "content": "Khi thuyết trình một vấn đề khoa học, bước nào sau đây là quan trọng để thu hút người nghe?",
        "options": [
            "A. Sử dụng càng nhiều thuật ngữ chuyên ngành càng tốt",
            "B. Đọc nguyên văn tài liệu tham khảo",
            "C. Trình bày lưu loát, kết hợp hình ảnh minh họa",
            "D. Nói nhỏ và ngắn gọn để tiết kiệm thời gian"
        ],
        "correct": "C",
        "solution": "Trình bày lưu loát, tự tin, diễn đạt rõ ràng kết hợp với các hình ảnh, sơ đồ, video minh họa trực quan sinh động là yếu tố quan trọng nhất để lôi cuốn và thu hút người nghe."
    },
    {
        "id": 5,
        "content": "Khi chuẩn bị bài thuyết trình, học sinh nên làm gì?",
        "options": [
            "A. Sao chép nội dung từ sách giáo khoa",
            "B. Ghi nhớ tất cả thông tin",
            "C. Tìm hiểu thông tin từ nhiều nguồn đáng tin cậy và sắp xếp logic",
            "D. Nhờ người khác chuẩn bị nội dung"
        ],
        "correct": "C",
        "solution": "Học sinh cần chủ động tìm hiểu, thu thập tài liệu từ nhiều nguồn đáng tin cậy (SGK, tạp chí khoa học, tư liệu chính thống), sau đó tổng hợp, chắt lọc và sắp xếp theo một cấu trúc logic, mạch lạc."
    }
]

# -----------------------------------------------------------------------------
# PHÂN MÔN KHTN 1 (HÓA HỌC)
# -----------------------------------------------------------------------------
HOA_PART1_QUESTIONS = [
    {
        "id": 1,
        "content": "Hợp chất hữu cơ là",
        "options": [
            "A. hợp chất khó tan trong nước.",
            "B. hợp chất của carbon và một số nguyên tố khác trừ N, Cl, O.",
            "C. hợp chất của carbon trừ CO, CO~2~, H~2~CO~3~, muối carbonate kim loại...",
            "D. hợp chất có nhiệt độ sôi cao."
        ],
        "correct": "C",
        "solution": "Hợp chất hữu cơ là hợp chất của carbon (trừ một số hợp chất vô cơ của carbon như CO, CO~2~, H~2~CO~3~, muối carbonate kim loại, carbide kim loại, cyanide...)."
    },
    {
        "id": 2,
        "content": "Hóa học hữu cơ là:",
        "options": [
            "A. Ngành hóa học chuyên nghiên cứu các hợp chất có trong tự nhiên",
            "B. Ngành hóa học chuyên nghiên cứu các hợp chất của carbon",
            "C. Ngành hóa học chuyên nghiên cứu về các hợp chất hữu cơ",
            "D. Ngành hóa học chuyên nghiên cứu các chất trong cơ thể sống"
        ],
        "correct": "C",
        "solution": "Hóa học hữu cơ là ngành hóa học chuyên nghiên cứu về thành phần, cấu trúc, tính chất, phương pháp điều chế và ứng dụng của các hợp chất hữu cơ."
    },
    {
        "id": 3,
        "content": "Công thức cấu tạo của hợp chất hữu cơ cho biết:",
        "options": [
            "A. Thành phần nguyên tố.",
            "B. Số lượng nguyên tử của mỗi nguyên tố.",
            "C. Thành phần nguyên tố và số lượng nguyên tử của mỗi nguyên tố.",
            "D. Trật tự và cách thức liên kết giữa các nguyên tử trong phân tử."
        ],
        "correct": "D",
        "solution": "Công thức cấu tạo biểu diễn thành phần nguyên tố, số lượng nguyên tử và đặc biệt cho biết trật tự cùng cách thức liên kết (liên kết đơn, liên kết đôi,...) giữa các nguyên tử trong phân tử."
    },
    {
        "id": 4,
        "content": "Nhận xét nào sau đây không đúng?",
        "options": [
            "A. Hợp chất hữu cơ là hợp chất của carbon trừ CO, CO~2~, H~2~CO~3~, muối carbonate kim loại...",
            "B. Trong các hợp chất hữu cơ, carbon luôn có hóa trị IV.",
            "C. Mỗi hợp chất hữu cơ có một trật tự liên kết giữa các nguyên tử trong phân tử.",
            "D. Liên kết chủ yếu trong HCHC là liên kết ion."
        ],
        "correct": "D",
        "solution": "Liên kết hóa học chủ yếu trong phân tử hợp chất hữu cơ là liên kết cộng hóa trị (được tạo nên bởi các cặp electron dùng chung), không phải liên kết ion."
    },
    {
        "id": 5,
        "content": "Trong các chất sau: CH~4~, CO, C~2~H~6~, K~2~CO~3~, C~2~H~5~ONa có",
        "options": [
            "A. 1 hợp chất hữu cơ và 4 hợp chất vô cơ.",
            "B. 2 hợp chất hữu cơ và 3 hợp chất vô cơ.",
            "C. 4 hợp chất hữu cơ và 1 hợp chất vô cơ.",
            "D. 3 hợp chất hữu cơ và 2 hợp chất vô cơ."
        ],
        "correct": "D",
        "solution": "- Các hợp chất hữu cơ (3 chất): CH~4~ (methane), C~2~H~6~ (ethane), C~2~H~5~ONa (sodium ethoxide).\n- Các hợp chất vô cơ (2 chất): CO (carbon monoxide), K~2~CO~3~ (potassium carbonate)."
    },
    {
        "id": 6,
        "content": "Số công thức cấu tạo của C~4~H~10~ là",
        "options": ["A. 1", "B. 2", "C. 3", "D. 4"],
        "correct": "B",
        "solution": "C~4~H~10~ có 2 đồng phân cấu tạo mạch carbon:\n(1) Mạch không phân nhánh: CH~3~–CH~2~–CH~2~–CH~3~ (butane).\n(2) Mạch phân nhánh: CH~3~–CH(CH~3~)–CH~3~ (2-methylpropane hay isobutane)."
    },
    {
        "id": 7,
        "content": "Alkane là các hydrocarbon",
        "options": [
            "A. mạch vòng, chỉ chứa liên kết đơn trong phân tử.",
            "B. mạch thẳng, chỉ chứa liên kết đơn trong phân tử.",
            "C. mạch hở, chỉ chứa liên kết đơn trong phân tử.",
            "D. mạch nhánh, chỉ chứa liên kết đơn trong phân tử."
        ],
        "correct": "C",
        "solution": "Alkane là các hydrocarbon mạch hở (có thể là mạch không phân nhánh hoặc mạch phân nhánh), trong phân tử chỉ chứa các liên kết đơn C–C và C–H."
    },
    {
        "id": 8,
        "content": "Công thức phân tử nào sau đây không phải là công thức của một alkane?",
        "options": ["A. C~2~H~6~", "B. C~3~H~6~", "C. C~4~H~10~", "D. C~5~H~12~"],
        "correct": "B",
        "solution": "Công thức chung của alkane là C~n~H~2n+2~ (n >= 1).\nC~3~H~6~ có dạng C~n~H~2n~ (n = 3) nên là alkene hoặc cycloalkane, không phải alkane."
    },
    {
        "id": 9,
        "content": "Phát biểu nào sau đây về alkane là không đúng?",
        "options": [
            "A. Trong phân tử alkane chỉ có liên kết đơn.",
            "B. Chỉ các alkane là chất khí ở điều kiện thường được dùng làm nhiên liệu.",
            "C. Các alkane lỏng được dùng sản xuất xăng, dầu và làm dung môi.",
            "D. Các alkane rắn được dùng làm nhựa đường."
        ],
        "correct": "B",
        "solution": "Phát biểu B không đúng vì các alkane lỏng (như xăng, dầu hỏa, dầu diesel) và alkane rắn (nến sáp parafin) cũng đóng vai trò là những nguồn nhiên liệu cực kì quan trọng."
    },
    {
        "id": 10,
        "content": "Phát biểu nào sau đây đúng?",
        "options": [
            "A. Alkene có CT chung là C~n~H~2n+2~.",
            "B. Alkene là hydrocarbon mạch hở có một liên kết đôi C=C.",
            "C. Alkene không làm mất màu dung dịch Bromine.",
            "D. Alkene tác dụng với Br~2~ trong dung dịch theo tỉ lệ mol 1 : 2."
        ],
        "correct": "B",
        "solution": "Alkene là các hydrocarbon mạch hở, trong phân tử có chứa một liên kết đôi C=C, có công thức phân tử chung là C~n~H~2n~ (n >= 2)."
    },
    {
        "id": 11,
        "content": "Ethylene không có phản ứng nào sau đây?",
        "options": [
            "A. Tác dụng với NaOH.",
            "B. Trùng hợp tạo PE.",
            "C. Làm nhạt màu dung dịch Br~2~.",
            "D. Cháy trong không khí."
        ],
        "correct": "A",
        "solution": "Ethylene (C~2~H~4~) không phản ứng với dung dịch base kiềm như NaOH.\nEthylene có phản ứng cộng (làm mất màu nước Br~2~), phản ứng trùng hợp (tạo nhựa polyethylene - PE) và phản ứng cháy tỏa nhiều nhiệt."
    },
    {
        "id": 12,
        "content": "Nhận định nào sau đây là sai?",
        "options": [
            "A. Nhiên liệu là những chất cháy được, khi cháy tỏa nhiệt và phát sáng.",
            "B. Nhiên liệu đóng vai trò quan trọng trong đời sống và sản xuất.",
            "C. Nhiên liệu rắn gồm than mỏ, gỗ ...",
            "D. Nhiên liệu khí có năng suất tỏa nhiệt thấp, gây độc hại cho môi trường."
        ],
        "correct": "D",
        "solution": "Nhiên liệu khí (khí thiên nhiên, khí dầu mỏ hóa lỏng LPG) có năng suất tỏa nhiệt cao, dễ cháy hoàn toàn và thân thiện với môi trường hơn so với nhiên liệu rắn."
    },
    {
        "id": 13,
        "content": "Để sử dụng nhiên liệu cho hiệu quả cần đảm bảo yêu cầu nào sau đây?",
        "options": [
            "A. Cung cấp đủ oxygen cho quá trình cháy.",
            "B. Tăng diện tích tiếp xúc của nhiên liệu rắn với không khí.",
            "C. Điều chỉnh lượng nhiên liệu để duy trì sự cháy phù hợp với nhu cầu sử dụng.",
            "D. Cả 3 yêu cầu trên."
        ],
        "correct": "D",
        "solution": "Để sử dụng nhiên liệu an toàn, tiết kiệm và hiệu quả, cần đồng thời thực hiện cả 3 biện pháp: cấp đủ oxygen, tăng bề mặt tiếp xúc của nhiên liệu rắn với không khí và điều chỉnh ngọn lửa phù hợp nhu cầu."
    },
    {
        "id": 14,
        "content": "Cho các nhiên liệu sau: xăng, gas, dầu diesel, dầu mazut, dầu hỏa. Sắp xếp nhiệt độ sôi tăng dần các nhiên liệu trên là",
        "options": [
            "A. xăng, gas, dầu diesel, dầu hỏa, nhựa đường.",
            "B. gas, xăng, dầu diesel, nhựa đường, dầu hỏa.",
            "C. gas, xăng, dầu hỏa, dầu diesel, nhựa đường.",
            "D. gas, xăng, nhựa đường, dầu diesel, dầu hỏa."
        ],
        "correct": "C",
        "solution": "Thứ tự nhiệt độ sôi tăng dần theo chiều tăng số nguyên tử carbon và khối lượng phân tử: Khí gas < Xăng < Dầu hỏa < Dầu diesel < Dầu mazut (nhựa đường)."
    },
    {
        "id": 15,
        "content": "Tính chất vật lí của ethylic alcohol là",
        "options": [
            "A. chất lỏng không màu, nhẹ hơn nước, tan vô hạn trong nước, hòa tan được nhiều chất như iodine, benzene.",
            "B. chất lỏng màu hồng, nhẹ hơn nước, tan vô hạn trong nước, hòa tan được nhiều chất như: iodine,...",
            "C. chất lỏng không màu, không tan trong nước, hòa tan được nhiều chất như: iodine, benzene,...",
            "D. chất lỏng không màu, nặng hơn nước, tan vô hạn trong nước, hòa tan được nhiều chất: iodine, benzene."
        ],
        "correct": "A",
        "solution": "Ethylic alcohol (C~2~H~5~OH) là chất lỏng không màu, mùi thơm đặc trưng, vị cay, nhẹ hơn nước (D = 0,789 g/cm^3^), tan vô hạn trong nước và hòa tan tốt nhiều chất hữu cơ khác."
    },
    {
        "id": 16,
        "content": "Hãy cho biết dựa vào tính chất vật lí nào để làm cơ sở pha loãng ethylic alcohol thành dung dịch ethylic alcohol 45°.",
        "options": [
            "A. Tính tan vô hạn của ethylic alcohol trong nước.",
            "B. Tính nhẹ hơn nước của ethylic alcohol.",
            "C. Nhiệt độ sôi của ethylic alcohol thấp hơn nước.",
            "D. Cả A, B, C đều đúng."
        ],
        "correct": "A",
        "solution": "Nhờ tính tan vô hạn trong nước của ethylic alcohol, ta có thể trộn cồn với nước theo bất kì tỉ lệ thể tích nào để thu được dung dịch cồn có nồng độ (độ cồn) mong muốn."
    },
    {
        "id": 17,
        "content": "Công thức cấu tạo thu gọn của ethylic alcohol là",
        "options": [
            "A. CH~2~ – CH~3~ – OH.",
            "B. CH~3~ – O – CH~3~.",
            "C. CH~2~ – CH~2~ – OH~2~.",
            "D. CH~3~ – CH~2~ – OH."
        ],
        "correct": "D",
        "solution": "Công thức cấu tạo thu gọn của ethylic alcohol là CH~3~–CH~2~–OH (hoặc viết gọn là C~2~H~5~OH). Nhóm –O–CH~3~ là của dimethyl ether."
    },
    {
        "id": 18,
        "content": "Trong phân tử ethylic alcohol, nhóm nào gây nên tính chất đặc trưng của nó?",
        "options": [
            "A. Nhóm – CH~3~",
            "B. Nhóm CH~3~ – CH~2~ -",
            "C. Nhóm – OH",
            "D. Cả phân tử"
        ],
        "correct": "C",
        "solution": "Nhóm hydroxy (–OH) gắn vào nguyên tử carbon no là nhóm định chức quy định các tính chất hóa học đặc trưng của phân tử alcohol."
    },
    {
        "id": 19,
        "content": "Cồn có tác dụng diệt khuẩn tốt nên thường dùng để khử khuẩn. Hãy cho biết ý nghĩa của kí hiệu \"cồn 70°\".",
        "options": [
            "A. Cồn diệt được 70% các loại vi khuẩn",
            "B. Có 70 mL ethylic alcohol nguyên chất trong 30 mL dung dịch",
            "C. Ở 20°C, có 70 mL ethylic alcohol nguyên chất có trong 100 mL dung dịch",
            "D. Cồn có thể diệt khuẩn ở 70°C"
        ],
        "correct": "C",
        "solution": "Độ cồn là số mililít ethylic alcohol nguyên chất có trong 100 mililít dung dịch cồn ở 20°C. Kí hiệu \"cồn 70°\" nghĩa là trong 100 mL dung dịch có chứa 70 mL ethylic alcohol nguyên chất."
    },
    {
        "id": 20,
        "content": "Ethylic alcohol dễ cháy nên cần lưu ý gì khi sử dụng ethylic alcohol?",
        "options": [
            "A. Không dùng diêm, bật lửa đốt trực tiếp",
            "B. Bảo quản ở nơi thông thoáng",
            "C. Bảo quản nơi tránh xa nguồn nhiệt",
            "D. Cả A, B, C đều đúng"
        ],
        "correct": "D",
        "solution": "Để đảm bảo an toàn phòng chống cháy nổ khi dùng cồn: cần bảo quản nơi thoáng khí, đậy chặt nắp, để xa các nguồn nhiệt, nguồn điện và tuyệt đối không châm lửa trực tiếp vào chai/bình chứa cồn."
    },
    {
        "id": 21,
        "content": "Công thức phân tử của acetic acid là:",
        "options": ["A. C~2~H~4~O~2~", "B. C~2~H~6~O", "C. C~6~H~12~O~6~", "D. C~2~H~4~"],
        "correct": "A",
        "solution": "Acetic acid có công thức cấu tạo là CH~3~COOH, gồm 2 nguyên tử C, 4 nguyên tử H và 2 nguyên tử O. Do đó công thức phân tử là C~2~H~4~O~2~."
    },
    {
        "id": 22,
        "content": "Đặc điểm nào sau đây không phải là tính chất vật lý của acetic acid",
        "options": [
            "A. tan vô hạn trong nước",
            "B. nặng hơn nước (ở 25°C)",
            "C. chất lỏng không màu",
            "D. có vị ngọt"
        ],
        "correct": "D",
        "solution": "Acetic acid là chất lỏng không màu, có vị chua gắt đặc trưng của giấm ăn, tan vô hạn trong nước và hơi nặng hơn nước (khối lượng riêng D = 1,05 g/cm^3^ ở 25°C). Acetic acid không có vị ngọt."
    },
    {
        "id": 23,
        "content": "Nhỏ vài giọt dung dịch acetic acid vào mẩu quỳ tím, hiện tượng quan sát được là",
        "options": [
            "A. quỳ tím hóa xanh",
            "B. quỳ tím mất màu",
            "C. quỳ tím hóa đỏ",
            "D. quỳ tím không đổi màu"
        ],
        "correct": "C",
        "solution": "Acetic acid là một acid hữu cơ, trong dung dịch phân li tạo ion H^+^ làm giấy quỳ tím hóa đỏ (hoặc chuyển sang màu hồng đậm)."
    },
    {
        "id": 24,
        "content": "Sản phẩm của phản ứng ester hóa giữa acetic acid và ethylic alcohol là",
        "options": [
            "A. C~2~H~5~COOH",
            "B. C~2~H~5~COOCH~3~",
            "C. CH~3~COOC~3~H~7~",
            "D. CH~3~COOC~2~H~5~"
        ],
        "correct": "D",
        "solution": "Phương trình phản ứng ester hóa:\nCH~3~COOH + C~2~H~5~OH  ⇄ (H~2~SO~4~ đặc, t°)  CH~3~COOC~2~H~5~ + H~2~O.\nSản phẩm thu được là ethyl acetate (CH~3~COOC~2~H~5~), có mùi thơm hoa quả đặc trưng."
    }
]

HOA_PART2_QUESTIONS = [
    {
        "id": 1,
        "title": "Cho các phát biểu sau về đặc điểm cấu tạo của HCHC:",
        "statements": [
            {"label": "a", "text": "Trong các hợp chất hữu cơ, carbon luôn có hóa trị IV, hydrogen có hóa trị I, oxygen có hóa trị II.", "correct": True, "explain": "Đúng. Theo thuyết cấu tạo hóa học, carbon luôn có hóa trị IV, hydrogen hóa trị I, oxygen hóa trị II."},
            {"label": "b", "text": "Những nguyên tử carbon trong phân tử hợp chất hữu cơ có thể liên kết trực tiếp với nhau tạo thành mạch carbon.", "correct": True, "explain": "Đúng. Các nguyên tử carbon có khả năng liên kết trực tiếp với nhau để hình thành các loại mạch carbon."},
            {"label": "c", "text": "Tính chất của HCHC chỉ phụ thuộc vào thành phần nguyên tố và số lượng nguyên tử.", "correct": False, "explain": "Sai. Tính chất của hợp chất hữu cơ phụ thuộc rất lớn vào trật tự liên kết và cấu trúc phân tử (hiện tượng đồng phân)."},
            {"label": "d", "text": "Có 2 loại mạch carbon: mạch hở phân nhánh, mạch hở không phân nhánh.", "correct": False, "explain": "Sai. Có 3 loại mạch carbon cơ bản: mạch hở không phân nhánh, mạch hở phân nhánh và mạch vòng."}
        ]
    },
    {
        "id": 2,
        "title": "Methane là thành phần chính của khí thiên nhiên và được sử dụng làm nhiên liệu:",
        "statements": [
            {"label": "a", "text": "Methane có CTPT là C~2~H~6~.", "correct": False, "explain": "Sai. Methane có công thức phân tử là CH~4~ (C~2~H~6~ là ethane)."},
            {"label": "b", "text": "Phản ứng cháy hoàn toàn của methane tạo ra carbon dioxide và hơi nước.", "correct": True, "explain": "Đúng. PTHH: CH~4~ + 2O~2~  → (t°)  CO~2~ + 2H~2~O."},
            {"label": "c", "text": "Khi đốt cháy hoàn toàn 1 mol methane, cần vừa đủ 1 mol oxygen.", "correct": False, "explain": "Sai. Theo phương trình phản ứng cháy, 1 mol CH~4~ phản ứng vừa đủ với 2 mol O~2~."},
            {"label": "d", "text": "Khi thiếu oxygen, methane cháy không hoàn toàn, tạo ra carbon monoxide hoặc muội than.", "correct": True, "explain": "Đúng. Trong điều kiện thiếu O~2~, sản phẩm cháy có thể tạo khí độc CO hoặc carbon tự do (muội than C)."}
        ]
    },
    {
        "id": 3,
        "title": "Một học sinh dẫn khí ethylene qua dung dịch bromine. Sau một thời gian, màu của dung dịch nhạt dần rồi mất đi. Xét các phát biểu sau:",
        "statements": [
            {"label": "a", "text": "Ethylene phản ứng với bromine theo phản ứng cộng.", "correct": True, "explain": "Đúng. Liên kết kém bền trong liên kết đôi C=C bị phá vỡ để cộng nguyên tử bromine."},
            {"label": "b", "text": "Hiện tượng dung dịch bromine mất màu có thể dùng để phân biệt ethylene và các alkene khác.", "correct": False, "explain": "Sai. Các alkene khác cũng có liên kết đôi C=C nên đều làm mất màu dung dịch bromine, không dùng để phân biệt giữa các alkene với nhau."},
            {"label": "c", "text": "Sản phẩm của phản ứng giữa ethylene và bromine là C~2~H~4~Br~2~.", "correct": True, "explain": "Đúng. Phương trình: CH~2~=CH~2~ + Br~2~  →  CH~2~Br–CH~2~Br (1,2-dibromoethane: C~2~H~4~Br~2~)."},
            {"label": "d", "text": "Ethane cũng làm mất màu dung dịch bromine trong điều kiện thường tương tự ethylene.", "correct": False, "explain": "Sai. Ethane (C~2~H~6~) là alkane no, chỉ chứa liên kết đơn nên không phản ứng với dung dịch bromine ở điều kiện thường."}
        ]
    },
    {
        "id": 4,
        "title": "Ethylic alcohol có nhiều ứng dụng trong đời sống và sản xuất. Tuy nhiên, ethylic alcohol cũng có những nguy cơ nhất định khi sử dụng không đúng cách:",
        "statements": [
            {"label": "a", "text": "Ethanol có thể được dùng làm dung môi hòa tan một số chất hữu cơ.", "correct": True, "explain": "Đúng. Ethanol là dung môi hòa tan phổ biến cho chất béo, phẩm màu, tinh dầu, dược liệu..."},
            {"label": "b", "text": "Ethanol dễ cháy nên cần bảo quản tránh xa nguồn lửa.", "correct": True, "explain": "Đúng. Hơi ethanol dễ bắt lửa gây cháy nổ nguy hiểm."},
            {"label": "c", "text": "Ethanol dùng trong công nghiệp hoặc phòng thí nghiệm có thể uống được nếu đã pha loãng với nước.", "correct": False, "explain": "Sai. Cồn công nghiệp và thí nghiệm thường chứa methanol và chất phụ gia cực độc, tuyệt đối không được uống."},
            {"label": "d", "text": "Ethanol được pha với xăng để tạo một số loại nhiên liệu sinh học.", "correct": True, "explain": "Đúng. Xăng E5, E10 chứa 5% và 10% ethanol sinh học nhằm giảm thiểu khí thải độc hại."}
        ]
    }
]

HOA_PART3_QUESTIONS = [
    {
        "id": 1,
        "content": "Cho các phát biểu sau:\n(1) Đốt cháy hợp chất hữu cơ luôn thu được CO~2~ và H~2~O.\n(2) Hợp chất hữu cơ có ở xung quanh ta.\n(3) Hợp chất của carbon là hợp chất hữu cơ.\n(4) Hóa học hữu cơ có vai trò quan trọng trong sự phát triển kinh tế, xã hội.\nSố phát biểu đúng là bao nhiêu?",
        "correct": "2",
        "solution": "- (1) Sai: Các hợp chất hữu cơ chỉ chứa C, Cl (như CCl~4~) khi cháy không tạo H~2~O.\n- (2) Đúng: Hợp chất hữu cơ tồn tại phổ biến trong thức ăn, cây cối, cơ thể sống, đồ dùng...\n- (3) Sai: Các hợp chất CO, CO~2~, H~2~CO~3~, muối carbonate là hợp chất vô cơ.\n- (4) Đúng: Hóa học hữu cơ là nền tảng của công nghiệp dược phẩm, dệt may, chất dẻo, năng lượng...\nVậy có 2 phát biểu đúng là (2) và (4)."
    },
    {
        "id": 2,
        "content": "Alkane X có hàm lượng carbon là 82,76%. Khối lượng phân tử của X là bao nhiêu amu?",
        "correct": "58",
        "solution": "Đặt công thức tổng quát của alkane X là C~n~H~2n+2~ (n ≥ 1).\nTa có: %C = [12n / (14n + 2)] × 100% = 82,76%\n=> 12n = 0,8276 × (14n + 2) = 11,5864n + 1,6552\n=> 0,4136n = 1,6552 => n = 4.\nCông thức phân tử của X là C~4~H~10~.\nKhối lượng phân tử của X = 12 × 4 + 1 × 10 = 58 amu."
    },
    {
        "id": 3,
        "content": "Đốt cháy hoàn toàn 2,479 Lit khí ethylene (đkc) trong không khí thu được CO~2~ và m gam H~2~O. Giá trị m là bao nhiêu, biết hiệu suất phản ứng là 80%.",
        "correct": "2,88",
        "solution": "Số mol ethylene tham gia phản ứng:\nn(C~2~H~4~) = 2,479 / 24,79 = 0,1 mol.\nPhương trình hóa học:\nC~2~H~4~ + 3O~2~  → (t°)  2CO~2~ + 2H~2~O\nTheo PTHH: n(H~2~O lý thuyết) = 2 × n(C~2~H~4~) = 2 × 0,1 = 0,2 mol.\nm(H~2~O lý thuyết) = 0,2 × 18 = 3,6 gam.\nDo hiệu suất phản ứng H = 80%, khối lượng nước thực tế thu được là:\nm = 3,6 × (80 / 100) = 2,88 gam."
    },
    {
        "id": 4,
        "content": "Từ 250 mL dung dịch ethylic alcohol 30° có thể pha chế được bao nhiêu lít dung dịch ethylic alcohol 15°?",
        "correct": "0,5",
        "solution": "Thể tích ethylic alcohol nguyên chất trong 250 mL dung dịch cồn 30°:\nV(cồn nguyên chất) = (250 × 30) / 100 = 75 mL.\nGọi V_dd là thể tích dung dịch cồn 15° pha chế được:\nĐộ cồn = [V(cồn nguyên chất) / V_dd] × 100 = 15°\n=> V_dd = (75 × 100) / 15 = 500 mL = 0,5 lít."
    }
]

HOA_ESSAY_QUESTIONS = [
    {
        "id": 1,
        "content": "Hãy sắp xếp các hợp chất dưới đây thành hai nhóm: nhóm 1 gồm các hợp chất hữu cơ và nhóm 2 gồm các hợp chất vô cơ:\nC~6~H~6~, H~2~SO~4~, C~6~H~12~O~6~, H~2~CO~3~, CaCO~3~, KNO~3~, C~2~H~4~, NaOH, Al~2~O~3~, CH~3~Cl, CH~3~OH",
        "solution": "- *Nhóm 1 (Hợp chất hữu cơ):* C~6~H~6~, C~6~H~12~O~6~, C~2~H~4~, CH~3~Cl, CH~3~OH.\n- *Nhóm 2 (Hợp chất vô cơ):* H~2~SO~4~, H~2~CO~3~, CaCO~3~, KNO~3~, NaOH, Al~2~O~3~."
    },
    {
        "id": 2,
        "content": "Sắp xếp các chất sau đây vào một trong hai nhóm: hydrocarbon và dẫn xuất của hydrocarbon:\nCH~4~, CH~3~Cl, CH~2~=CH~2~, CH~3~CH~2~OH, CH~3~COOH, CH~3~NH~2~, CH~3~CH~2~CH~3~, CH~3~CH=CH~2~, CH~3~COOCH~2~CH~3~.",
        "solution": "- *Hydrocarbon (phân tử chỉ tạo bởi 2 nguyên tố C và H):*\n  CH~4~, CH~2~=CH~2~, CH~3~CH~2~CH~3~, CH~3~CH=CH~2~.\n- *Dẫn xuất của hydrocarbon (ngoài C và H còn có nguyên tố khác như O, Cl, N):*\n  CH~3~Cl, CH~3~CH~2~OH, CH~3~COOH, CH~3~NH~2~, CH~3~COOCH~2~CH~3~."
    },
    {
        "id": 3,
        "content": "Có hai chất khí là C~2~H~4~ và C~3~H~8~. Nêu cách tiến hành phân biệt hai khí trên. Viết phương trình hóa học xảy ra (nếu có).",
        "solution": "- *Cách tiến hành:* Dẫn lần lượt từng khí lội từ từ qua bình chứa dung dịch nước bromine (Br~2~ có màu nâu đỏ hoặc vàng cam).\n- *Hiện tượng & Kết luận:*\n  + Khí làm dung dịch bromine nhạt dần rồi mất màu là ethylene (C~2~H~4~).\n    PTHH: CH~2~=CH~2~ + Br~2~ → CH~2~Br–CH~2~Br (dung dịch mất màu).\n  + Khí không làm đổi màu dung dịch bromine là propane (C~3~H~8~)."
    },
    {
        "id": 4,
        "content": "Ethylic alcohol được sử dụng làm nhiên liệu vì khi cháy toả ra nhiều nhiệt. Biết 1 mol ethylic alcohol cháy hoàn toàn sẽ toả ra 1368 kJ. Tính nhiệt lượng tỏa ra khi đốt cháy 9,2 gam ethylic alcohol.",
        "solution": "- Khối lượng mol phân tử của ethylic alcohol (C~2~H~5~OH):\n  M = 12 × 2 + 1 × 6 + 16 = 46 (g/mol).\n- Số mol ethylic alcohol đem đốt cháy:\n  n = m / M = 9,2 / 46 = 0,2 (mol).\n- Nhiệt lượng tỏa ra khi đốt cháy hoàn toàn 9,2 gam (0,2 mol) ethylic alcohol là:\n  Q = 0,2 × 1368 = 273,6 (kJ)."
    }
]

# -----------------------------------------------------------------------------
# PHÂN MÔN KHTN 2 (VẬT LÍ)
# -----------------------------------------------------------------------------
VATLI_QUESTIONS = [
    {
        "id": 1,
        "content": "Động năng của vật phụ thuộc vào",
        "options": ["A. khối lượng.", "B. tốc độ.", "C. thể tích.", "D. A và B đều đúng."],
        "correct": "D",
        "solution": "Theo công thức W_đ = (1/2)·m·v^2, động năng của vật phụ thuộc vào cả khối lượng (m) và tốc độ (v) của vật chuyển động."
    },
    {
        "id": 2,
        "content": "Khi tốc độ của vật tăng gấp đôi thì động năng của vật sẽ",
        "options": ["A. tăng gấp 2.", "B. tăng gấp 4.", "C. tăng gấp 8.", "D. giảm 2 lần."],
        "correct": "B",
        "solution": "Vì động năng tỉ lệ thuận với bình phương tốc độ (W_đ ∝ v^2), nên khi tốc độ tăng lên 2 lần (v' = 2v), động năng của vật sẽ tăng lên 2^2 = 4 lần."
    }
]
