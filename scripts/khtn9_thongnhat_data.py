# -*- coding: utf-8 -*-
"""
Dữ liệu chuẩn hóa Đề cương ôn tập Giữa học kỳ 1 - Môn KHTN 9 (2026 - 2027)
TRƯỜNG THCS THỐNG NHẤT - TỔ KHOA HỌC TỰ NHIÊN
Thầy TRẦN VĂN THẠNH - 0777.470.803

Gồm đầy đủ 3 phân môn:
1. Chất và sự biến đổi chất (Hóa học) - Bài 1, 18, 19, 20, 21
2. Năng lượng và sự biến đổi (Vật lí) - Bài 2, 3, 4
3. Vật sống (Sinh học) - Bài 36, 37, 38
"""

INFO = {
    "truong": "TRƯỜNG THCS THỐNG NHẤT",
    "to": "TỔ: KHOA HỌC TỰ NHIÊN",
    "tieu_de": "HƯỚNG DẪN ÔN TẬP GIỮA HỌC KỲ I",
    "mon": "KHOA HỌC TỰ NHIÊN 9",
    "nam_hoc": "NĂM HỌC: 2026 - 2027",
    "giao_vien": "Thầy TRẦN VĂN THẠNH - 0777.470.803",
    "co_so_1": "CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế",
    "co_so_2": "CS2: 24 Đặng Thái Thân, TP Huế",
}

# ==============================================================================
# PHÂN MÔN 1: CHẤT VÀ SỰ BIẾN ĐỔI CHẤT (HÓA HỌC)
# ==============================================================================
HOA_THEORY = [
    {
        "title": "Bài 1: Nhận biết một số dụng cụ, hóa chất. Thuyết trình một vấn đề khoa học",
        "content": [
            "Nhận biết được một số dụng cụ và hoá chất sử dụng trong dạy học môn Khoa học tự nhiên 9 (dụng cụ đo lường, đun nóng, chứa đựng; quy tắc an toàn phòng thí nghiệm).",
            "Trình bày được các bước viết và trình bày báo cáo một vấn đề khoa học: (1) Xác định mục tiêu và đối tượng báo cáo; (2) Thu thập, phân tích dữ liệu; (3) Lập dàn ý cấu trúc báo cáo (Tiêu đề, Giới thiệu, Phương pháp, Kết quả, Thảo luận, Kết luận, Tài liệu tham khảo); (4) Viết bài báo cáo; (5) Thuyết trình báo cáo khoa học."
        ]
    },
    {
        "title": "Bài 18: Tính chất chung của kim loại",
        "content": [
            "Nêu được tính chất vật lí chung của kim loại: Tính dẻo, tính dẫn điện, tính dẫn nhiệt và tính ánh kim. (Kim loại dẫn điện tốt nhất: Ag > Cu > Au > Al > Fe; Kim loại cứng nhất: Cr; Nhiệt độ nóng chảy cao nhất: W; Thấp nhất: Hg).",
            "Trình bày được tính chất hoá học cơ bản của kim loại:",
            "+ Tác dụng với phi kim (oxygen, sulfur, chlorine): tạo oxide hoặc muối. Ví dụ: 2Mg + O2 -> 2MgO; 2Fe + 3Cl2 -> 2FeCl3; Fe + S -> FeS.",
            "+ Tác dụng với nước: Các kim loại hoạt động hóa học mạnh (K, Na, Ca, Ba) tác dụng với nước ở nhiệt độ thường tạo base và giải phóng H2: 2Na + 2H2O -> 2NaOH + H2.",
            "+ Tác dụng với dung dịch hydrochloric acid (HCl) và sulfuric acid loãng (H2SO4 loãng): Kim loại đứng trước H đẩy được H2 ra khỏi acid: Fe + 2HCl -> FeCl2 + H2; Zn + H2SO4 -> ZnSO4 + H2.",
            "+ Tác dụng với dung dịch muối: Kim loại hoạt động mạnh hơn đẩy kim loại hoạt động yếu hơn ra khỏi dung dịch muối: Fe + CuSO4 -> FeSO4 + Cu; Cu + 2AgNO3 -> Cu(NO3)2 + 2Ag.",
            "Mô tả hiện tượng các thí nghiệm liên quan tính chất hoá học của kim loại. Viết phương trình phản ứng minh hoạ."
        ]
    },
    {
        "title": "Bài 19: Dãy hoạt động hoá học của kim loại",
        "content": [
            "Dãy hoạt động hoá học của kim loại: K, Na, Ca, Mg, Al, Zn, Fe, Pb, (H), Cu, Ag, Au.",
            "Ý nghĩa của dãy hoạt động hoá học:",
            "+ Mức độ hoạt động hoá học của kim loại giảm dần từ trái sang phải.",
            "+ Kim loại đứng trước Mg (K, Na, Ca) phản ứng với nước ở nhiệt độ thường giải phóng H2.",
            "+ Kim loại đứng trước H phản ứng với dung dịch acid (HCl, H2SO4 loãng) giải phóng H2.",
            "+ Kim loại đứng trước (trừ K, Na, Ca, Ba) đẩy được kim loại đứng sau ra khỏi dung dịch muối."
        ]
    },
    {
        "title": "Bài 20: Tách kim loại và việc sử dụng hợp kim",
        "content": [
            "Phương pháp tách kim loại theo mức độ hoạt động hoá học:",
            "+ Điện phân nóng chảy: Áp dụng cho kim loại hoạt động mạnh (K, Na, Ca, Mg, Al). Ví dụ: 2Al2O3 -> 4Al + 3O2 (đpnc, xúc tác criolit); 2NaCl -> 2Na + Cl2 (đpnc).",
            "+ Nhiệt luyện: Áp dụng cho kim loại hoạt động trung bình (từ Zn đến Cu) bằng chất khử CO, C, H2. Ví dụ: Fe2O3 + 3CO -> 2Fe + 3CO2 (to); ZnO + C -> Zn + CO (to).",
            "+ Thuỷ luyện: Áp dụng cho kim loại hoạt động yếu (Cu, Ag, Au) bằng cách dùng kim loại mạnh hơn đẩy kim loại yếu ra khỏi dung dịch muối.",
            "Trình bày quá trình tách một số kim loại có nhiều ứng dụng:",
            "+ Tách sắt ra khỏi iron(III) oxide (hematite) bởi carbon oxide trong lò cao.",
            "+ Tách nhôm ra khỏi aluminium oxide (bauxite) bằng phản ứng điện phân nóng chảy.",
            "+ Tách kẽm khỏi zinc sulfide bởi oxygen và carbon: 2ZnS + 3O2 -> 2ZnO + 2SO2 (to); ZnO + C -> Zn + CO (to).",
            "Khái niệm hợp kim; thành phần, tính chất đặc trưng của một số hợp kim phổ biến (Gang: 2 - 5% C; Thép: < 2% C; Duralumin: hợp kim nhôm nhẹ, siêu bền chế tạo vỏ máy bay).",
            "Nguyên liệu và các giai đoạn cơ bản sản xuất gang và thép trong công nghiệp; tác động tiêu cực đến môi trường không khí (phát thải CO2 gây hiệu ứng nhà kính, CO, bụi, SO2 gây mưa acid)."
        ]
    },
    {
        "title": "Bài 21: Sự khác nhau cơ bản giữa phi kim và kim loại",
        "content": [
            "Ứng dụng thực tiễn của đơn chất phi kim: Than (nhiên liệu, chất khử, hấp phụ), lưu huỳnh (sản xuất H2SO4, diêm, cao su), khí chlorine (khử trùng nước, tẩy trắng).",
            "Sự khác nhau cơ bản về tính chất giữa phi kim và kim loại:",
            "+ Tính dẫn điện, dẫn nhiệt: Kim loại dẫn điện, dẫn nhiệt rất tốt; phi kim không dẫn điện (trừ than chì) và dẫn nhiệt kém.",
            "+ Nhiệt độ nóng chảy, nhiệt độ sôi, khối lượng riêng: Kim loại thường cao hơn phi kim.",
            "+ Khả năng tạo ion: Kim loại có xu hướng nhường electron tạo cation (ion dương); phi kim có xu hướng nhận electron tạo anion (ion âm).",
            "+ Phản ứng với oxygen: Kim loại tạo basic oxide / amphoteric oxide; phi kim tạo acidic oxide / neutral oxide.",
            "Tích hợp BĐKH và ứng phó PCTT:",
            "+ Ảnh hưởng của khai thác quặng kim loại đến môi trường (phá hủy thảm thực vật, xói mòn đất, ô nhiễm nguồn nước, không khí). Đề xuất giải pháp tách kim loại thân thiện môi trường, tái chế kim loại.",
            "+ Nghiên cứu ứng dụng hợp kim chịu lực, chống ăn mòn trong công trình chống thiên tai (cầu treo, kè chắn sóng, nhà chống bão).",
            "+ Ứng dụng chlorine khử trùng nước sinh hoạt và môi trường sau lũ lụt, thiên tai."
        ]
    }
]

HOA_MCQ = [
    {
        "num": 1,
        "question": "Tiêu đề của một bài báo cáo khoa học cần đáp ứng những yêu cầu sau:",
        "options": [
            "A. Chính xác và mô tả rõ ràng nội dung của báo cáo.",
            "B. Tóm tắt những phát hiện chính và đề xuất hướng nghiên cứu tiếp theo.",
            "C. Giới thiệu lý do nghiên cứu và mục tiêu của nghiên cứu.",
            "D. Nêu ngắn gọn phương pháp và kết quả nghiên cứu."
        ],
        "correct": "A",
        "solution": "Tiêu đề bài báo cáo khoa học phải ngắn gọn, chính xác, nêu bật và phản ánh rõ ràng nội dung nghiên cứu thực hiện."
    },
    {
        "num": 2,
        "question": "Phần cuối cùng của bài báo cáo một vấn đề khoa học là:",
        "options": [
            "A. Kết quả.",
            "B. Kết luận.",
            "C. Thảo luận.",
            "D. Tài liệu tham khảo"
        ],
        "correct": "D",
        "solution": "Cấu trúc chuẩn của một bài báo cáo khoa học kết thúc bằng danh mục Tài liệu tham khảo (trích dẫn nguồn tài liệu đã sử dụng)."
    },
    {
        "num": 3,
        "question": "Mục đích của mục giới thiệu trong bài báo cáo một vấn đề khoa học là gì?",
        "options": [
            "A. Tổng hợp nội dung chính của bài báo cáo.",
            "B. Mô tả vấn đề nghiên cứu và tầm quan trọng của vấn đề; mục tiêu của nghiên cứu.",
            "C. Trình bày dữ liệu thu được một cách rõ ràng, sử dụng biểu đồ, hình ảnh hoặc bảng.",
            "D. Phân tích và giải thích ý nghĩa của kết quả; so sánh với các nghiên cứu khác (nếu có)."
        ],
        "correct": "B",
        "solution": "Mục Giới thiệu nhằm đặt vấn đề, nêu rõ tính cấp thiết, tầm quan trọng và mục tiêu cụ thể của nghiên cứu."
    },
    {
        "num": 4,
        "question": "Tiêu bản nhiễm sắc thể người được sử dụng thực hành cho chủ đề:",
        "options": [
            "A. Năng lượng và sự biến đổi.",
            "B. Trái đất và bầu trời.",
            "C. Chất và sự biến đổi của chất.",
            "D. Vật sống."
        ],
        "correct": "D",
        "solution": "Nhiễm sắc thể là cấu trúc mang vật chất di truyền của sinh vật, thuộc phân môn Sinh học trong chủ đề Vật sống."
    },
    {
        "num": 5,
        "question": "Kim loại có những tính chất vật lí chung nào sau đây?",
        "options": [
            "A. Tính dẻo, tính dẫn điện, nhiệt độ nóng chảy cao.",
            "B. Tính dẻo, tính dẫn nhiệt, rất cứng.",
            "C. Tính dẻo, tính dẫn điện, tính dẫn nhiệt và tính ánh kim.",
            "D. Tính dẻo, có ánh kim, khối lượng riêng lớn."
        ],
        "correct": "C",
        "solution": "Các electron tự do trong tinh thể kim loại quyết định 4 tính chất vật lí chung: tính dẻo, tính dẫn điện, tính dẫn nhiệt và có ánh kim."
    },
    {
        "num": 6,
        "question": "Kim loại nào sau đây có độ cứng cao nhất (khoảng 8.5 theo thang độ cứng Mohs) thường được dùng để chế tạo hợp kim thép không gỉ và có độ bền cơ học cao?",
        "options": [
            "A. Copper.",
            "B. Aluminium.",
            "C. Zinc.",
            "D. Chromium."
        ],
        "correct": "D",
        "solution": "Chromium (Cr) là kim loại cứng nhất trong tất cả các kim loại, độ cứng đạt 8.5 - 9 trên thang Mohs (có thể rạch được thủy tinh)."
    },
    {
        "num": 7,
        "question": "Nhờ có tính dẫn điện mà một số kim loại được sử dụng làm dây dẫn điện. Các kim loại khác nhau có khả năng dẫn điện khác nhau. Kim loại nào sau đây có tính dẫn điện tốt nhất?",
        "options": [
            "A. Copper.",
            "B. Silver.",
            "C. Gold.",
            "D. Aluminium."
        ],
        "correct": "B",
        "solution": "Thứ tự dẫn điện giảm dần của một số kim loại thông dụng: Ag > Cu > Au > Al > Fe. Do đó Silver (Bạc) dẫn điện tốt nhất."
    },
    {
        "num": 8,
        "question": "Kim loại nào sau đây được điều chế bằng phương pháp điện phân nóng chảy?",
        "options": [
            "A. Cu và Fe.",
            "B. Ag và Hg.",
            "C. Na và Al.",
            "D. Fe và K."
        ],
        "correct": "C",
        "solution": "Các kim loại có độ hoạt động hóa học mạnh (K, Na, Ca, Mg, Al) chỉ có thể điều chế bằng phương pháp điện phân nóng chảy hợp chất của chúng (ví dụ NaCl đpnc tạo Na; Al2O3 đpnc tạo Al)."
    },
    {
        "num": 9,
        "question": "Dãy gồm các kim loại phản ứng được với dung dịch acid sulfuric loãng là",
        "options": [
            "A. Fe, Al, Zn.",
            "B. Mg, Al, Hg.",
            "C. K, Mg, Cu.",
            "D. Ba, Cu, Au."
        ],
        "correct": "A",
        "solution": "Các kim loại đứng trước H trong dãy hoạt động hóa học tác dụng được với dung dịch H2SO4 loãng giải phóng H2. Cu, Hg, Au đứng sau H nên không phản ứng. Do đó chọn dãy Fe, Al, Zn."
    },
    {
        "num": 10,
        "question": "Cho các kim loại Zn, Cu, Ag lần lượt phản ứng với dung dịch muối Copper(II) sulfate và Silver Nitrate. Số cặp chất có xảy ra phản ứng là",
        "options": [
            "A. 2",
            "B. 3",
            "C. 4",
            "D. 5."
        ],
        "correct": "B",
        "solution": "Dựa vào quy tắc kim loại đứng trước đẩy kim loại đứng sau ra khỏi dung dịch muối:\n- Zn đẩy được Cu ra khỏi CuSO4: Zn + CuSO4 -> ZnSO4 + Cu (1)\n- Zn đẩy được Ag ra khỏi AgNO3: Zn + 2AgNO3 -> Zn(NO3)2 + 2Ag (2)\n- Cu đẩy được Ag ra khỏi AgNO3: Cu + 2AgNO3 -> Cu(NO3)2 + 2Ag (3)\n- Ag không phản ứng với muối nào.\nTổng cộng có 3 cặp chất xảy ra phản ứng."
    },
    {
        "num": 11,
        "question": "Nhóm kim loại nào có thể tác dụng với nước ở nhiệt độ thường:",
        "options": [
            "A. Zinc, Potassium.",
            "B. Sodium, Lithium.",
            "C. Calcium, Magnesium.",
            "D. Iron, Zinc."
        ],
        "correct": "B",
        "solution": "Các kim loại kiềm (Li, Na, K...) và kim loại kiềm thổ nặng (Ca, Ba) phản ứng mãnh liệt với nước ở nhiệt độ thường tạo hydroxide kiềm và khí H2. Magnesium phản ứng rất chậm ở nhiệt độ thường; Zn, Fe không phản ứng với nước nguội."
    },
    {
        "num": 12,
        "question": "Dãy kim loại nào sau đây được sắp xếp theo chiều hoạt động hóa học tăng dần?",
        "options": [
            "A. K, Mg, Cu, Al, Zn.",
            "B. Fe, Cu, K, Mg, Al.",
            "C. Ag, Cu, Fe, Al, K.",
            "D. Fe, K, Mg, Cu, Na."
        ],
        "correct": "C",
        "solution": "Dãy hoạt động hoá học giảm dần là: K > Al > Fe > Cu > Ag. Do đó chiều tăng dần mức độ hoạt động là: Ag < Cu < Fe < Al < K."
    },
    {
        "num": 13,
        "question": "Gang và thép là hai hợp kim quan trọng nhất của sắt. Chúng có rất nhiều ứng dụng trong đời sống và sản xuất. Để sản xuất gang, người ta khử oxide của sắt trong quặng sắt (hematite hoặc magnetite) bằng khí CO (được tạo ra do đốt cháy than cốc). Sắt được tạo ra trong quá trình sản xuất gang bằng phương pháp nào?",
        "options": [
            "A. Điện phân nóng chảy.",
            "B. Nhiệt luyện.",
            "C. Thủy luyện.",
            "D. Điện phân dung dịch."
        ],
        "correct": "B",
        "solution": "Khử oxide kim loại ở nhiệt độ cao bằng chất khử như CO, C, H2 là bản chất của phương pháp nhiệt luyện."
    },
    {
        "num": 14,
        "question": "Hợp kim là vật liệu được tạo thành từ một kim loại cơ bản kết hợp với một hoặc nhiều kim loại hoặc phi kim khác (ví dụ: thép là hợp kim của sắt và carbon). Các nhận định chính về ưu điểm của một số hợp kim:\n(1) Tất cả các hợp kim đều không bị ăn mòn và có thể chống lại mọi tác động của thiên tai.\n(2) Độ cứng của hợp kim thường cao hơn kim loại nguyên chất.\n(3) Nhôm nguyên chất mềm, nhẹ nhưng không đủ bền để làm thân vỏ máy bay. Do đó, người ta thường sử dụng hợp kim nhôm (Duralumin) để tăng độ bền.\n(4) Thép có nhiều ưu điểm hơn sắt về độ cứng, độ đàn hồi, khả năng chịu lực nên được sử dụng rộng rãi trong các công trình xây dựng, giao thông.\nSố nhận định Đúng là",
        "options": [
            "A. 1",
            "B. 2",
            "C. 3.",
            "D. 4"
        ],
        "correct": "C",
        "solution": "- Nhận định (1) Sai: Không phải tất cả hợp kim đều không bị ăn mòn (gang, thép thường vẫn bị ăn mòn điện hóa trong tự nhiên).\n- Các nhận định (2), (3), (4) đều Đúng.\n=> Có 3 nhận định đúng."
    },
    {
        "num": 15,
        "question": "Tái chế kim loại được xem là giải pháp bền vững giúp bảo vệ môi trường và nguồn tài nguyên. Một số nhận định về ý nghĩa của giải pháp tái chế kim loại như sau:\n(1) Tái chế kim loại giảm nhu cầu khai thác quặng và giúp tiết kiệm năng lượng.\n(2) Giảm phát thải khí nhà kính so với quá trình sản xuất kim loại từ quặng.\n(3) Tiết kiệm chi phí khai thác quặng và tinh chế kim loại từ quặng.\n(4) Tái chế kim loại không làm giảm chi phí và mức độ ô nhiễm môi trường.\n(5) Việc sử dụng nhôm tái chế giúp tiết kiệm tài nguyên, giảm thiểu chất thải, và góp phần vào sự phát triển bền vững của xã hội.\nTheo em, số nhận định đúng về ý nghĩa của giải pháp tái chế kim loại là",
        "options": [
            "A. 2.",
            "B. 3.",
            "C. 4",
            "D. 5."
        ],
        "correct": "C",
        "solution": "- Nhận định (4) Sai: Tái chế kim loại LÀM GIẢM chi phí sản xuất và GIẢM mức độ ô nhiễm môi trường.\n- Các nhận định (1), (2), (3), (5) đều Đúng.\n=> Có 4 nhận định đúng."
    }
]

HOA_TF = [
    {
        "num": 16,
        "intro": "Dưới đây là một số nhận định về tính chất, ứng dụng của kim loại, hợp kim và phi kim. Hãy xác định các nhận định đúng/ sai:",
        "items": [
            {
                "label": "a",
                "text": "Kim loại nhôm (Aluminium) bền trong không khí hơn kim loại sắt (Iron) do nhôm không phản ứng với oxygen trong không khí.",
                "ans": "Sai (S)",
                "reason": "Nhôm có phản ứng với oxygen nhưng ngay lập tức tạo thành lớp màng oxide Al2O3 rất mỏng, mịn và bền vững bao bọc bề mặt, ngăn cản không cho nhôm phản ứng tiếp với oxygen và nước."
            },
            {
                "label": "b",
                "text": "Kim cương được sử dụng làm mũi khoan, dao cắt thủy tinh vì kim cương là chất có độ cứng rất lớn (đạt 10/10 theo thang độ cứng của Mohs).",
                "ans": "Đúng (Đ)",
                "reason": "Kim cương có cấu trúc tinh thể nguyên tử tứ diện cực kì bền chặt, độ cứng cao nhất trong tự nhiên."
            },
            {
                "label": "c",
                "text": "Trong dãy hoạt động hóa học kim loại, kim loại đứng trước đẩy kim loại đứng sau ra khỏi dung dịch muối, nhờ tính chất này để điều chế các kim loại có mức hoạt động yếu hơn.",
                "ans": "Đúng (Đ)",
                "reason": "Đây là nguyên tắc của phương pháp thủy luyện để điều chế kim loại hoạt động trung bình và yếu."
            },
            {
                "label": "d",
                "text": "Carbon hoạt tính dùng trong mặt nạ phòng độc, công nghiệp hóa chất và y học nhờ khả năng hấp phụ cao.",
                "ans": "Đúng (Đ)",
                "reason": "Carbon hoạt tính có diện tích bề mặt xốp rất lớn, hấp phụ tốt các chất màu, mùi và khí độc hại."
            },
            {
                "label": "e",
                "text": "Chlorine dùng để khử trùng nước sinh hoạt, xử lý nước hồ bơi, nước thải công nghiệp.",
                "ans": "Đúng (Đ)",
                "reason": "Chlorine tác dụng với nước tạo HClO có tính oxi hóa rất mạnh, tiêu diệt vi khuẩn, vi sinh vật gây bệnh."
            },
            {
                "label": "f",
                "text": "Các phi kim thường dẫn điện, dẫn nhiệt tốt hơn kim loại.",
                "ans": "Sai (S)",
                "reason": "Hầu hết phi kim dẫn điện và dẫn nhiệt rất kém (kim loại dẫn điện, dẫn nhiệt tốt hơn nhiều)."
            },
            {
                "label": "g",
                "text": "Gang và thép là hợp kim chủ yếu của sắt và carbon.",
                "ans": "Đúng (Đ)",
                "reason": "Gang chứa 2 - 5% C, thép chứa dưới 2% C cùng một lượng nhỏ nguyên tố khác."
            },
            {
                "label": "h",
                "text": "Gang có nhiều ứng dụng hơn thép.",
                "ans": "Sai (S)",
                "reason": "Thép có tính đàn hồi, dẻo dai, độ bền cơ học cao hơn gang nên có nhiều ứng dụng rộng rãi hơn trong xây dựng, giao thông, chế tạo máy."
            },
            {
                "label": "i",
                "text": "Để sản xuất thép người ta dùng than cốc để tạo chất khử.",
                "ans": "Sai (S)",
                "reason": "Than cốc dùng làm nhiên liệu và tạo chất khử CO trong quá trình sản xuất GANG. Trong luyện THÉP, người ta thổi khí oxygen để oxi hóa bớt carbon và tạp chất trong gang."
            },
            {
                "label": "j",
                "text": "Quá trình sản xuất gang thép tạo ra khí CO2 gây ô nhiễm môi trường.",
                "ans": "Đúng (Đ)",
                "reason": "Khí CO2 thải ra từ quá trình đốt cháy than cốc và khử oxide sắt là tác nhân chính gây hiệu ứng nhà kính và biến đổi khí hậu."
            }
        ]
    }
]

HOA_SHORT = [
    {
        "num": 17,
        "question": "Cho dãy gồm các kim loại sau: Cu, Zn, K, Na, Al, Fe, Ba. Dựa vào mức độ hoạt động hóa học của kim loại, có bao nhiêu kim loại được điều chế bằng phương pháp nhiệt luyện?",
        "solution": "Các kim loại hoạt động trung bình và yếu (đứng sau Al trong dãy hoạt động hóa học) được điều chế bằng phương pháp nhiệt luyện.\nTrong dãy gồm 3 kim loại: Cu, Zn, Fe.\n(Các kim loại K, Na, Ba, Al được điều chế bằng điện phân nóng chảy).\nĐáp số: 3."
    },
    {
        "num": 18,
        "question": "Cho 11,2 gam Iron tác dụng hết với khí Chlorine. Giá trị khối lượng của muối thu được là: ...... gam",
        "solution": "Số mol Fe: n_Fe = 11,2 / 56 = 0,2 mol.\nPTHH: 2Fe + 3Cl2 --to--> 2FeCl3\nTheo PTHH: n_FeCl3 = n_Fe = 0,2 mol.\nKhối lượng muối FeCl3 thu được: m_FeCl3 = 0,2 * 162,5 = 32,5 gam.\nĐáp số: 32,5."
    },
    {
        "num": 19,
        "question": "Cho 13,6 gam hỗn hợp gồm Fe và Cu vào dung dịch H2SO4 loãng dư, thu được 4,958 lít khí H2 (đkc); m (gam) chất rắn không tan X và dung dịch chứa Y. Giá trị m là: ....... gam",
        "solution": "Số mol H2 ở điều kiện chuẩn: n_H2 = 4,958 / 24,79 = 0,2 mol.\nChỉ có Fe phản ứng với H2SO4 loãng (Cu đứng sau H nên không phản ứng):\nFe + H2SO4 -> FeSO4 + H2\nTheo PTHH: n_Fe = n_H2 = 0,2 mol => m_Fe = 0,2 * 56 = 11,2 gam.\nChất rắn không tan X là Cu.\nKhối lượng Cu: m = 13,6 - 11,2 = 2,4 gam.\nĐáp số: 2,4."
    }
]

HOA_ESSAY = [
    {
        "num": 1,
        "question": "Trình bày tính chất hoá học chung của kim loại. Viết phương trình phản ứng minh hoạ.",
        "solution": "1. Tác dụng với phi kim:\n- Tác dụng với oxygen tạo oxide base: 2Mg + O2 --to--> 2MgO; 3Fe + 2O2 --to--> Fe3O4.\n- Tác dụng với phi kim khác (chlorine, sulfur) tạo muối: 2Fe + 3Cl2 --to--> 2FeCl3; Fe + S --to--> FeS.\n2. Tác dụng với nước: Các kim loại hoạt động mạnh (K, Na, Ca, Ba) tác dụng ở nhiệt độ thường giải phóng H2: 2Na + 2H2O -> 2NaOH + H2.\n3. Tác dụng với dung dịch acid (HCl, H2SO4 loãng): Kim loại đứng trước H đẩy H2 ra khỏi acid: Fe + 2HCl -> FeCl2 + H2; Zn + H2SO4 -> ZnSO4 + H2.\n4. Tác dụng với dung dịch muối: Kim loại mạnh hơn đẩy kim loại yếu hơn ra khỏi muối: Fe + CuSO4 -> FeSO4 + Cu; Cu + 2AgNO3 -> Cu(NO3)2 + 2Ag."
    },
    {
        "num": 2,
        "question": "Viết dãy hoạt động hoá học của kim loại. Nêu ý nghĩa.",
        "solution": "- Dãy hoạt động hoá học của kim loại:\n  K, Na, Ca, Mg, Al, Zn, Fe, Pb, (H), Cu, Ag, Au.\n- Ý nghĩa của dãy:\n1. Mức độ hoạt động hoá học của kim loại giảm dần từ trái sang phải.\n2. Kim loại đứng trước Mg (K, Na, Ca) tác dụng được với nước ở nhiệt độ thường giải phóng H2.\n3. Kim loại đứng trước H tác dụng với dung dịch acid (HCl, H2SO4 loãng) giải phóng khí H2.\n4. Kim loại đứng trước (trừ K, Na, Ca, Ba) đẩy được kim loại đứng sau ra khỏi dung dịch muối."
    },
    {
        "num": 3,
        "question": "Dựa vào mức độ hoạt động hóa học của kim loại, trình bày các phương pháp tách kim loại cơ bản từ các hợp chất của chúng.",
        "solution": "1. Phương pháp điện phân nóng chảy: Dùng để tách các kim loại hoạt động hoá học mạnh (K, Na, Ca, Mg, Al) khỏi oxide hoặc chloride.\nVí dụ: 2Al2O3 --đpnc, criolit--> 4Al + 3O2; 2NaCl --đpnc--> 2Na + Cl2.\n2. Phương pháp nhiệt luyện: Dùng để tách kim loại có độ hoạt động trung bình (từ Zn đến Cu) bằng cách dùng chất khử (CO, C, H2) khử oxide kim loại ở nhiệt độ cao.\nVí dụ: Fe2O3 + 3CO --to--> 2Fe + 3CO2; ZnO + C --to--> Zn + CO.\n3. Phương pháp thuỷ luyện: Dùng để tách kim loại hoạt động yếu (Cu, Ag, Au) bằng cách dùng dung dịch acid/muối hòa tan hợp chất kim loại, rồi dùng kim loại có tính khử mạnh hơn đẩy kim loại cần tách ra khỏi dung dịch."
    },
    {
        "num": 5,
        "question": "Nêu thành phần của Gang, Thép. Trình bày nguyên liệu, các công đoạn, PTHH xảy ra trong quá trình sản xuất gang, thép.",
        "solution": "a) Thành phần:\n- Gang: Là hợp kim của sắt (Iron) với carbon (trong đó hàm lượng carbon chiếm từ 2% đến 5%) và một lượng nhỏ các nguyên tố Si, Mn, S, P...\n- Thép: Là hợp kim của sắt với carbon (trong đó hàm lượng carbon dưới 2%) và một lượng rất nhỏ Si, Mn...\nb) Sản xuất Gang:\n- Nguyên liệu: Quặng sắt (hematite Fe2O3, magnetite Fe3O4), than cốc, đá vôi CaCO3, không khí giàu oxygen. Thiết bị: Lò cao.\n- Các công đoạn và PTHH chính:\n+ Tạo chất khử CO: C + O2 --to--> CO2; CO2 + C --to--> 2CO.\n+ Khử oxide sắt: Fe2O3 + 3CO --to--> 2Fe + 3CO2.\n+ Tạo xỉ loại tạp chất: CaCO3 --to--> CaO + CO2; CaO + SiO2 --to--> CaSiO3 (xỉ nhẹ nổi lên trên gang lỏng).\nc) Sản xuất Thép:\n- Nguyên liệu: Gang, sắt phế liệu, khí oxygen. Thiết bị: Lò luyện thép (lò thổi oxygen).\n- Quá trình: Thổi khí O2 vào gang lỏng để oxi hóa bớt carbon và tạp chất: C + O2 --to--> CO2; Si + O2 --to--> SiO2; 2Mn + O2 --to--> 2MnO."
    },
    {
        "num": 6,
        "question": "Viết phương trình hoá học các phản ứng sau, ghi rõ điều kiện phản ứng nếu có:\na) Iron -> Iron (II) Chloride.\nb) Iron -> Iron (III) Chloride.\nc) Copper -> Silver\nd) Aluminium -> Aluminium Sulfide.\ne) Sodium phản ứng với nước\nf) Khí Carbon monoxide + Iron(II, III) oxide.\ng) Điện phân nóng chảy Aluminium oxide.\nh) Điện phân nóng chảy Calcium Chloride.",
        "solution": "a) Fe + 2HCl -> FeCl2 + H2 (hoặc Fe + CuCl2 -> FeCl2 + Cu)\nb) 2Fe + 3Cl2 --to--> 2FeCl3\nc) Cu + 2AgNO3 -> Cu(NO3)2 + 2Ag\nd) 2Al + 3S --to--> Al2S3\ne) 2Na + 2H2O -> 2NaOH + H2\nf) Fe3O4 + 4CO --to--> 3Fe + 4CO2\ng) 2Al2O3 --đpnc, criolit--> 4Al + 3O2\nh) CaCl2 --đpnc--> Ca + Cl2"
    },
    {
        "num": 7,
        "question": "Quặng Hematite thường dùng để sản xuất gang.\na. Em hãy cho biết để tách kim loại ra khỏi quặng trên ta cần dùng phương pháp gì?\nb. Tính khối lượng quặng hematite (chứa 70% Fe2O3) cần dùng để sản xuất ra 1 tấn gang chứa 95% Fe, biết hiệu suất của toàn bộ quá trình là 80%.",
        "solution": "a. Tách kim loại sắt từ quặng hematite dùng phương pháp nhiệt luyện (dùng CO khử Fe2O3 trong lò cao ở nhiệt độ cao).\nb. Tính toán:\n- Khối lượng Fe nguyên chất trong 1 tấn gang: m_Fe = 1 * 95% = 0,95 tấn = 950 kg.\n- Sơ đồ phản ứng bảo toàn Fe: Fe2O3 -> 2Fe\n  160 tấn Fe2O3 sinh ra 2 * 56 = 112 tấn Fe.\n  Khối lượng Fe2O3 theo lí thuyết cần dùng: m_Fe2O3(LT) = (0,95 * 160) / 112 = 1,35714 tấn.\n- Vì hiệu suất quá trình H = 80%, khối lượng Fe2O3 thực tế cần dùng:\n  m_Fe2O3(TT) = 1,35714 / 80% = 1,69643 tấn.\n- Quặng hematite chứa 70% Fe2O3, nên khối lượng quặng cần dùng là:\n  m_quặng = 1,69643 / 70% ≈ 2,4235 tấn (tương đương 2423,5 kg)."
    },
    {
        "num": 8,
        "question": "Người ta sản xuất nhôm từ 2 tấn quặng bauxite (có 60% Al2O3) bằng phương pháp điện phân nóng chảy.\na. Viết PTHH xảy ra.\nb. Tính khối lượng Aluminium thu được từ quá trình sản xuất trên. Biết hiệu suất phản ứng chỉ đạt 70%.",
        "solution": "a. PTHH điện phân nóng chảy bauxite:\n  2Al2O3 --đpnc, criolit--> 4Al + 3O2\nb. Tính toán:\n- Khối lượng Al2O3 có trong 2 tấn quặng: m_Al2O3 = 2 * 60% = 1,2 tấn.\n- Theo PTHH: 2 * 102 = 204 tấn Al2O3 sinh ra 4 * 27 = 108 tấn Al.\n  Khối lượng Al thu được theo lí thuyết: m_Al(LT) = (1,2 * 108) / 204 ≈ 0,6353 tấn.\n- Vì hiệu suất phản ứng chỉ đạt 70%, nên khối lượng Al thực tế thu được là:\n  m_Al(TT) = 0,6353 * 70% ≈ 0,4447 tấn (tương đương 444,7 kg)."
    }
]

# ==============================================================================
# PHÂN MÔN 2: NĂNG LƯỢNG VÀ SỰ BIẾN ĐỔI (VẬT LÍ)
# ==============================================================================
VATLI_THEORY = [
    {
        "title": "Kiến thức trọng tâm (Bài 2. Động năng và thế năng đến Bài 4. Công và công suất)",
        "content": [
            "Vận dụng công thức tính thế năng trọng trường: Wt = P.h = m.g.h để xác định các đại lượng còn lại khi đã biết 2 đại lượng.",
            "Công thức tính động năng của vật: Wđ = 1/2.m.v^2 (với m tính bằng kg, v tính bằng m/s, Wđ tính bằng J).",
            "Cơ năng là tổng động năng và thế năng của vật: Wc = Wđ + Wt.",
            "Vận dụng khái niệm cơ năng và định luật bảo toàn cơ năng phân tích được sự chuyển hoá năng lượng: Khi một vật chuyển động chỉ dưới tác dụng của trọng lực, cơ năng của vật được bảo toàn (Wc = const); động năng giảm thì thế năng tăng và ngược lại.",
            "Cơ năng không được bảo toàn khi có tác dụng của lực cản, lực ma sát (một phần cơ năng chuyển hóa thành nhiệt năng).",
            "Đơn vị đo công: Jun (J), kJ, calo (cal), BTU (1 cal ≈ 4,186 J; 1 BTU ≈ 1055 J).",
            "Đơn vị đo công suất: Oát (W), kW, J/s, Mã lực (HP), BTU/h (1 HP ≈ 746 W; 1 W ≈ 3,412 BTU/h)."
        ]
    }
]

VATLI_MCQ = [
    {
        "num": 1,
        "question": "Một quả bưởi có khối lượng 1,5 kg đang ở trên cây. Hỏi độ cao h của quả bưởi so với mặt đất là bao nhiêu? Chọn gốc thế năng ở mặt đất, lúc này quả bưởi đang có thế năng là 45 J.",
        "options": [
            "A. h = 2 m.",
            "B. h = 2,5 m.",
            "C. h = 3 m.",
            "D. h = 4 m."
        ],
        "correct": "C",
        "solution": "Áp dụng công thức thế năng trọng trường: Wt = P.h = 10.m.h => 45 = 10 * 1,5 * h = 15.h => h = 45 / 15 = 3 m."
    },
    {
        "num": 2,
        "question": "Tại một trận đấu, quả bóng được đá ra với tốc độ v = 90 km/h. Tính khối lượng m của quả bóng, biết quả bóng có động năng là 125 J.",
        "options": [
            "A. m = 0,5 kg.",
            "B. m = 0,4 kg.",
            "C. m = 4 kg.",
            "D. m = 2 kg."
        ],
        "correct": "B",
        "solution": "Đổi đơn vị tốc độ: v = 90 km/h = 90 / 3,6 = 25 m/s.\nÁp dụng công thức động năng: Wđ = 1/2.m.v^2 => 125 = 1/2 * m * 25^2 = 312,5.m => m = 125 / 312,5 = 0,4 kg."
    },
    {
        "num": 3,
        "question": "Nhóm học sinh đang chơi đá cầu, một bạn đá rất mạnh nên quả cầu bị mắc trên cành cây. Lúc này quả cầu có thế năng là 7 J, gốc thế năng ở mặt đất, độ cao của quả cầu là 7 m. Hỏi khối lượng của quả cầu là bao nhiêu?",
        "options": [
            "A. m = 10 kg.",
            "B. m = 1 kg.",
            "C. m = 0,1 kg.",
            "D. m = 0,01 kg."
        ],
        "correct": "C",
        "solution": "Áp dụng công thức: Wt = P.h = 10.m.h => 7 = 10 * m * 7 = 70.m => m = 7 / 70 = 0,1 kg (100 gam)."
    },
    {
        "num": 4,
        "question": "Cơ năng của một vật được xác định bởi",
        "options": [
            "A. tổng nhiệt năng và động năng.",
            "B. tổng thế năng và nhiệt năng.",
            "C. tổng hóa năng và động năng.",
            "D. tổng động năng và thế năng."
        ],
        "correct": "D",
        "solution": "Theo định nghĩa SGK KHTN 9: Cơ năng của một vật bằng tổng động năng và thế năng của vật đó."
    },
    {
        "num": 5,
        "question": "Biểu thức nào sau đây mô tả đúng mối quan hệ giữa cơ năng (Wc), động năng (Wđ) và thế năng (Wt)?",
        "options": [
            "A. Wc = Wđ + Wt.",
            "B. Wc = Wđ - Wt.",
            "C. Wc = Wđ x Wt.",
            "D. Wc = Wđ = Wt."
        ],
        "correct": "A",
        "solution": "Biểu thức chuẩn của cơ năng: Wc = Wđ + Wt."
    },
    {
        "num": 6,
        "question": "Một vật có thế năng Wt = 42 J, động năng Wđ = 14 J. Thì cơ năng Wc của nó bằng bao nhiêu? Coi cơ năng của nó được bảo toàn.",
        "options": [
            "A. Wc = 28 J.",
            "B. Wc = 3 J.",
            "C. Wc = 56 J.",
            "D. Wc = 588 J."
        ],
        "correct": "C",
        "solution": "Cơ năng của vật: Wc = Wt + Wđ = 42 + 14 = 56 J."
    },
    {
        "num": 7,
        "question": "Các đơn vị nào sau đây là đơn vị của Công mà em đã được học?",
        "options": [
            "A. MJ, BTU, cal, HP.",
            "B. kJ, BTU, cal, J.",
            "C. MJ, BTU/h, cal, J.",
            "D. HP, MJ, W, BTU/h."
        ],
        "correct": "B",
        "solution": "Các đơn vị đo công / năng lượng gồm: J, kJ, MJ, cal, kcal, BTU. Còn HP, W, BTU/h là các đơn vị của công suất."
    },
    {
        "num": 8,
        "question": "Các đơn vị nào sau đây là đơn vị của Công suất mà em đã được học?",
        "options": [
            "A. MJ, BTU, cal, HP.",
            "B. kJ, BTU, cal, J.",
            "C. MJ, BTU/h, cal, J.",
            "D. HP, J/s, W, BTU/h."
        ],
        "correct": "D",
        "solution": "Công suất đo tốc độ thực hiện công, đơn vị gồm: W, kW, J/s, HP (mã lực), BTU/h."
    },
    {
        "num": 9,
        "question": "Trong các đơn vị của công có một đơn vị mà ta thường hay gặp trong đời sống đó là calo (cal). Vậy 1 calo bằng bao nhiêu Jun?",
        "options": [
            "A. 1,486 J.",
            "B. 4,186 J.",
            "C. 4186 J.",
            "D. 1486 J."
        ],
        "correct": "B",
        "solution": "Theo quy ước năng lượng chuẩn: 1 cal = 4,186 J (và 1 kcal = 4186 J)."
    }
]

VATLI_SHORT = [
    {
        "num": 10,
        "question": "Bác Thường đang lái xe ô tô đi trên đường, lúc này ô tô có động năng bằng 50 kJ. Biết khối lượng của ô tô và người là 1000 kg. Hỏi bác Thường lái xe với tốc độ bằng bao nhiêu m/s?",
        "solution": "Đổi đơn vị: Wđ = 50 kJ = 50 000 J.\nTa có: Wđ = 1/2.m.v^2 => 50 000 = 1/2 * 1000 * v^2 = 500.v^2\n=> v^2 = 50 000 / 500 = 100 => v = 10 m/s (tương đương 36 km/h).\nĐáp số: 10 m/s."
    },
    {
        "num": 11,
        "question": "Trong khi giảng bài Động năng và Thế năng, thầy Minh cầm khối gỗ trên tay ở độ cao h = 1,2 m để làm ví dụ cho bài giảng. Biết thế năng của khối gỗ lúc này bằng 2,4 J, chọn gốc thế năng ở mặt đất, hỏi khối gỗ nặng bao nhiêu kg?",
        "solution": "Áp dụng công thức thế năng trọng trường: Wt = P.h = 10.m.h\n=> 2,4 = 10 * m * 1,2 = 12.m => m = 2,4 / 12 = 0,2 kg (200 g).\nĐáp số: 0,2 kg."
    },
    {
        "num": 12,
        "question": "Chiều chủ nhật, bạn Linh cùng với gia đình đi leo núi, leo được một lúc thì Linh dừng lại nghỉ ngơi, lúc này Linh có thế năng là 250 kJ. Hỏi bạn ấy đang ở độ cao bao nhiêu mét? Biết khối lượng của Linh là 50 kg.",
        "solution": "Đổi đơn vị: Wt = 250 kJ = 250 000 J.\nTa có: Wt = P.h = 10.m.h => 250 000 = 10 * 50 * h = 500.h\n=> h = 250 000 / 500 = 500 m.\nĐáp số: 500 m."
    },
    {
        "num": 13,
        "question": "Đơn vị công suất của máy điều hòa nhiệt độ người ta thường sử dụng là Mã lực (HP). Một máy điều hòa nhiệt độ được đặt trong phòng học với công suất 2 (HP). Hỏi công suất này đổi ra đơn vị oát là bao nhiêu?",
        "solution": "Theo chuẩn quy đổi mã lực cơ học: 1 HP ≈ 746 W (hoặc hệ mét 1 HP ≈ 736 W).\nVới 1 HP ≈ 746 W: P = 2 * 746 = 1492 W.\n(Nếu tính theo hệ mét 736 W: P = 2 * 736 = 1472 W).\nĐáp số: 1492 W (hoặc 1472 W)."
    },
    {
        "num": 14,
        "question": "Bạn Nam khi đẩy xe hàng thì đã thực hiện một công là 3 BTU. Vậy công của bạn Nam được đổi ra đơn vị J là bao nhiêu?",
        "solution": "Theo hệ số quy đổi năng lượng: 1 BTU ≈ 1055 J.\nCông bạn Nam đã thực hiện đổi sang Jun là: A = 3 * 1055 = 3165 J.\nĐáp số: 3165 J."
    },
    {
        "num": 15,
        "question": "Một máy sưởi ấm có công suất là 2051 W, em hãy đổi ra đơn vị BTU/h là bao nhiêu?",
        "solution": "Ta có quy đổi giữa W và BTU/h:\n1 BTU/h = 1055 J / 3600 s ≈ 0,293 W => 1 W ≈ 3,412 BTU/h.\nCông suất máy sưởi đổi ra BTU/h: P = 2051 * 3,412 ≈ 7000 BTU/h (chính xác: 2051 * 3600 / 1055 ≈ 7000 BTU/h).\nĐáp số: 7000 BTU/h."
    }
]

VATLI_ESSAY = [
    {
        "num": 1,
        "question": "Một vật đang chuyển động có động năng 120 J và thế năng 80 J so với mặt đất. Tính cơ năng của vật.",
        "solution": "Cơ năng của vật là:\nW = Wđ + Wt = 120 + 80 = 200 J.\nĐáp số: 200 J."
    },
    {
        "num": 2,
        "question": "Một quả bóng có khối lượng 2 kg đang chuyển động với tốc độ 6 m/s ở độ cao 5 m so với mặt đất. Biết trọng lượng của quả bóng là 20 N. Chọn mặt đất làm mốc tính thế năng. Tính cơ năng của quả bóng tại độ cao 5m.",
        "solution": "- Động năng của quả bóng tại độ cao 5 m:\n  Wđ = 1/2.m.v^2 = 1/2 * 2 * 6^2 = 36 J.\n- Thế năng của quả bóng tại độ cao 5 m:\n  Wt = P.h = 20 * 5 = 100 J.\n- Cơ năng của quả bóng:\n  W = Wđ + Wt = 36 + 100 = 136 J.\nĐáp số: 136 J."
    },
    {
        "num": 3,
        "question": "Một vật có khối lượng 4 kg đang chuyển động ở độ cao 3 m so với mặt đất. Biết cơ năng của vật là 180 J và trọng lượng của vật là 40 N. Chọn mặt đất làm mốc tính thế năng. Tính tốc độ chuyển động của vật tại độ cao 3m.",
        "solution": "- Thế năng của vật ở độ cao 3 m:\n  Wt = P.h = 40 * 3 = 120 J.\n- Động năng của vật tại độ cao 3 m:\n  Wđ = W - Wt = 180 - 120 = 60 J.\n- Ta có Wđ = 1/2.m.v^2 => 60 = 1/2 * 4 * v^2 = 2.v^2\n  => v^2 = 30 => v = √30 ≈ 5,48 m/s.\nĐáp số: 5,48 m/s."
    },
    {
        "num": 4,
        "question": "Một thùng hàng được thả rơi từ độ cao h = 1,5 m xuống mặt đất. Gốc thế năng ở mặt đất. Biết khối lượng thùng hàng là 3,5 kg. Tính Cơ năng của vật? Biết cơ năng được bảo toàn.",
        "solution": "- Khi mới thả rơi, vận tốc ban đầu bằng 0 nên động năng Wđ = 0.\n- Thế năng ban đầu của thùng hàng ở độ cao 1,5 m:\n  Wt = 10.m.h = 10 * 3,5 * 1,5 = 52,5 J.\n- Vì cơ năng được bảo toàn nên cơ năng của vật là:\n  W = Wt max = 52,5 J.\nĐáp số: 52,5 J."
    },
    {
        "num": 5,
        "question": "Quả bóng rổ có khối lượng 600g được ném từ mặt đất lên cao. Gốc thế năng ở mặt đất. Biết tốc độ ném bóng là 18 m/s. Tính Cơ năng của vật? Biết cơ năng được bảo toàn.",
        "solution": "- Đổi khối lượng: m = 600 g = 0,6 kg.\n- Ở mặt đất (gốc thế năng), thế năng Wt = 0.\n- Động năng ban đầu của quả bóng:\n  Wđ = 1/2.m.v^2 = 1/2 * 0,6 * 18^2 = 0,3 * 324 = 97,2 J.\n- Vì cơ năng được bảo toàn nên cơ năng của quả bóng là:\n  W = Wđ max = 97,2 J.\nĐáp số: 97,2 J."
    },
    {
        "num": 6,
        "question": "Một vật có trọng lượng 30 N đang chuyển động với tốc độ 4 m/s ở một độ cao h cao so với mặt đất. Biết khối lượng của vật là 3 kg và cơ năng của vật tại độ cao h là 144 J. Chọn mặt đất làm mốc tính thế năng. Tính độ cao của vật so với mặt đất.",
        "solution": "- Động năng của vật tại độ cao h:\n  Wđ = 1/2.m.v^2 = 1/2 * 3 * 4^2 = 24 J.\n- Thế năng của vật tại độ cao h:\n  Wt = W - Wđ = 144 - 24 = 120 J.\n- Ta có Wt = P.h => 120 = 30 * h => h = 120 / 30 = 4 m.\nVậy độ cao của vật so với mặt đất là 4 m.\nĐáp số: 4 m."
    },
    {
        "num": 7,
        "question": "Một quả bóng có trọng lượng 10 N được thả rơi từ độ cao 12 m so với mặt đất. Bỏ qua sức cản của không khí. Chọn mặt đất làm mốc tính thế năng.\na) Tính cơ năng của quả bóng lúc bắt đầu thả. Biết cơ năng được bảo toàn.\nb) Khi quả bóng ở độ cao 5 m, tính thế năng và động năng của quả bóng.\nc) Khi quả bóng ở ngay sát mặt đất, động năng của quả bóng bằng bao nhiêu?",
        "solution": "a) Lúc bắt đầu thả ở độ cao h = 12 m, vận tốc v = 0 nên Wđ = 0.\n- Thế năng lúc thả: Wt = P.h = 10 * 12 = 120 J.\n- Cơ năng của quả bóng: W = Wt = 120 J.\nb) Khi quả bóng ở độ cao h1 = 5 m:\n- Thế năng: Wt1 = P.h1 = 10 * 5 = 50 J.\n- Vì cơ năng được bảo toàn: Wđ1 = W - Wt1 = 120 - 50 = 70 J.\nc) Khi quả bóng ở ngay sát mặt đất (h = 0 nên Wt = 0):\n- Động năng của bóng đạt cực đại và bằng cơ năng: Wđ = W = 120 J."
    },
    {
        "num": 8,
        "question": "Khi nào thì cơ năng của vật được bảo toàn?",
        "solution": "Cơ năng của vật được bảo toàn khi vật chuyển động chỉ dưới tác dụng của các lực thế (như trọng lực, lực đàn hồi) và không có sự xuất hiện của các lực cản, lực ma sát làm tiêu hao năng lượng. Khi đó có sự chuyển hóa qua lại giữa động năng và thế năng nhưng tổng của chúng luôn không đổi."
    },
    {
        "num": 9,
        "question": "Khi nào thì cơ năng của vật KHÔNG được bảo toàn?",
        "solution": "Cơ năng của vật KHÔNG được bảo toàn khi vật chuyển động trong môi trường có lực ma sát hoặc lực cản (như ma sát trượt, ma sát lăn, sức cản của không khí hay nước). Khi đó một phần cơ năng của vật đã bị chuyển hóa thành nhiệt năng hoặc dạng năng lượng khác."
    },
    {
        "num": 10,
        "question": "Một quả bóng có khối lượng 0,5 kg được ném thẳng đứng lên cao. Khi quả bóng lên đến vị trí cao nhất, nó ở độ cao 10 m so với mặt đất và có tốc độ bằng 0. Biết trọng lượng của quả bóng là 5 N. Chọn mặt đất làm mốc tính thế năng.\na) Tính động năng của quả bóng ở vị trí cao nhất.\nb) Tính thế năng và cơ năng của quả bóng ở vị trí đó.",
        "solution": "a) Tại vị trí cao nhất, vận tốc v = 0 nên động năng của quả bóng bằng 0 (Wđ = 0 J).\nb) - Thế năng của quả bóng ở vị trí cao nhất (h = 10 m):\n  Wt = P.h = 5 * 10 = 50 J.\n- Cơ năng của quả bóng ở vị trí đó:\n  W = Wđ + Wt = 0 + 50 = 50 J."
    },
    {
        "num": 11,
        "question": "Một vật có khối lượng 2 kg rơi từ độ cao 15 m xuống độ cao 5 m so với mặt đất. Tại vị trí 15m so với mặt đất, vật có tốc độ 4 m/s. Biết trọng lượng của vật là 20 N và bỏ qua sức cản của không khí. Biết cơ năng được bảo toàn.\na) Tính cơ năng của vật tại độ cao 15 m.\nb) Tính tốc độ của vật tại độ cao 5 m.",
        "solution": "a) Tại độ cao h1 = 15 m:\n- Thế năng: Wt1 = P.h1 = 20 * 15 = 300 J.\n- Động năng: Wđ1 = 1/2.m.v1^2 = 1/2 * 2 * 4^2 = 16 J.\n- Cơ năng của vật: W = Wt1 + Wđ1 = 300 + 16 = 316 J.\nb) Tại độ cao h2 = 5 m:\n- Thế năng: Wt2 = P.h2 = 20 * 5 = 100 J.\n- Vì cơ năng được bảo toàn: Wđ2 = W - Wt2 = 316 - 100 = 216 J.\n- Ta có Wđ2 = 1/2.m.v2^2 => 216 = 1/2 * 2 * v2^2 = v2^2\n  => v2 = √216 ≈ 14,7 m/s (hoặc 6√6 m/s).\nĐáp số: 316 J và 14,7 m/s."
    }
]

# ==============================================================================
# PHÂN MÔN 3: VẬT SỐNG (SINH HỌC)
# ==============================================================================
SINHHOC_THEORY = [
    {
        "title": "Bài 36: Khái quát về di truyền học",
        "content": [
            "Các thuật ngữ cơ bản: Tính trạng, nhân tố di truyền (gene), cơ thể thuần chủng (dòng thuần), cặp tính trạng tương phản, tính trạng trội, tính trạng lặn, kiểu hình, kiểu gene, allele.",
            "Kí hiệu trong nghiên cứu di truyền học: P (thế hệ bố mẹ), F1, F2 (thế hệ con lai thứ nhất, thứ hai), ♂ (cơ thể đực), ♀ (cơ thể cái), x (phép lai)."
        ]
    },
    {
        "title": "Bài 37: Các quy luật di truyền của Mendel",
        "content": [
            "Quy luật phân li: Mỗi tính trạng do một cặp nhân tố di truyền (cặp allele) quy định. Khi giảm phân hình thành giao tử, các allele trong cặp phân li đồng đều về các giao tử, nên mỗi giao tử chỉ chứa một allele của cặp.",
            "Lai phân tích: Là phép lai giữa cá thể mang tính trạng trội cần xác định kiểu gene với cá thể mang tính trạng lặn. Nếu con lai đồng tính -> cá thể mang tính trạng trội thuần chủng (AA); nếu con lai phân tính (1 : 1) -> cá thể mang tính trạng trội dị hợp (Aa).",
            "Quy luật phân li độc lập: Các cặp nhân tố di truyền (cặp allele) quy định các tính trạng khác nhau phân li độc lập và tổ hợp tự do trong quá trình phát sinh giao tử và thụ tinh.",
            "Vận dụng quy luật Mendel giải bài toán di truyền 1 cặp và 2 cặp tính trạng."
        ]
    },
    {
        "title": "Bài 38: Nucleic acid và gene",
        "content": [
            "Khái niệm nucleic acid: Đại phân tử sinh học cấu tạo theo nguyên tắc đa phân (đơn phân là nucleotide), gồm 2 loại là DNA và RNA, đảm nhận chức năng lưu trữ, bảo quản và truyền đạt thông tin di truyền.",
            "Cấu trúc hóa học của DNA: Cấu tạo từ các nguyên tố C, H, O, N, P. Đơn phân là nucleotide gồm 4 loại: adenine (A), thymine (T), guanine (G), cytosine (C).",
            "Cấu trúc không gian của DNA (Watson & Crick - 1953):",
            "+ Gồm 2 mạch polynucleotide song song, ngược chiều, xoắn đều quanh một trục tưởng tượng từ trái sang phải (xoắn phải).",
            "+ Trên mỗi mạch, các nucleotide liên kết với nhau bằng liên kết cộng hóa trị (phosphodiester) theo chiều từ 5' tới 3'.",
            "+ Giữa hai mạch, các nucleotide liên kết với nhau bằng liên kết hydrogen theo nguyên tắc bổ sung (A liên kết với T bằng 2 liên kết hydrogen; G liên kết với C bằng 3 liên kết hydrogen, và ngược lại).",
            "+ Chu kì xoắn dài 34 Å gồm 10 cặp nucleotide, đường kính vòng xoắn là 20 Å.",
            "Chức năng của DNA: Lưu giữ, bảo quản và truyền đạt thông tin di truyền qua các thế hệ tế bào và cơ thể.",
            "Khái niệm gene: Gene là một đoạn của phân tử DNA mang thông tin di truyền xác định.",
            "Tính đa dạng và đặc thù của DNA: DNA rất đa dạng và đặc thù bởi số lượng, thành phần và trật tự sắp xếp xác định của các loại nucleotide."
        ]
    }
]

SINHHOC_MCQ = [
    {
        "num": 1,
        "question": "Tập hợp tất cả các tính trạng của cơ thể sinh vật được gọi là",
        "options": [
            "A. kiểu gene.",
            "B. dòng thuần.",
            "C. kiểu hình.",
            "D. allele."
        ],
        "correct": "C",
        "solution": "Theo định nghĩa SGK: Kiểu hình là tập hợp toàn bộ các tính trạng của cơ thể sinh vật (hoặc một số tính trạng đang được xét)."
    },
    {
        "num": 2,
        "question": "Ai là người đặt nền móng cho di truyền học?",
        "options": [
            "A. Charles Darwin.",
            "B. Barbara McClintock.",
            "C. Wilmut và Campbell.",
            "D. Gregor Johann Mendel."
        ],
        "correct": "D",
        "solution": "Gregor Johann Mendel (1822 - 1884) là người đầu tiên phát hiện ra các quy luật di truyền cơ bản bằng phương pháp phân tích cơ thể lai trên cây đậu hà lan, được coi là cha đẻ của di truyền học."
    },
    {
        "num": 3,
        "question": "Tính trạng trội biểu hiện ra kiểu hình khi",
        "options": [
            "A. có kiểu gene đồng hợp lặn hoặc dị hợp.",
            "B. có kiểu gene đồng hợp trội hoặc dị hợp.",
            "C. có kiểu gene đồng hợp trội.",
            "D. có kiểu gene đồng hợp lặn."
        ],
        "correct": "B",
        "solution": "Trong trường hợp trội hoàn toàn, allele trội lấn át allele lặn nên kiểu hình trội được biểu hiện ở cả trạng thái đồng hợp trội (AA) và trạng thái dị hợp (Aa)."
    },
    {
        "num": 4,
        "question": "Kiểu hình là",
        "options": [
            "A. tổ hợp toàn bộ tính trạng của cơ thể sinh vật.",
            "B. tổ hợp một số tính trạng của cơ thể sinh vật.",
            "C. tổ hợp các tính trạng lặn của cơ thể sinh vật.",
            "D. tổ hợp tính trạng tương phản của cơ thể sinh vật."
        ],
        "correct": "A",
        "solution": "Kiểu hình là tổ hợp toàn bộ các đặc điểm, tính trạng bên ngoài và bên trong của cơ thể sinh vật."
    },
    {
        "num": 5,
        "question": "Khi đem lai các cá thể thuần chủng khác nhau về một cặp tính trạng tương phản. Mendel đã phát hiện ở thế hệ F1",
        "options": [
            "A. luôn luôn biểu hiện kiểu hình giống mẹ.",
            "B. luôn luôn biểu hiện kiểu hình giống bố.",
            "C. chỉ biểu hiện 1 trong 2 kiểu hình của bố hoặc mẹ.",
            "D. biểu hiện tính trạng trung gian giữa bố và mẹ."
        ],
        "correct": "C",
        "solution": "Khi lai hai dòng bố mẹ thuần chủng tương phản (ví dụ hạt vàng x hạt xanh), thế hệ F1 đồng tính, chỉ biểu hiện một tính trạng của bố hoặc mẹ (tính trạng trội)."
    },
    {
        "num": 6,
        "question": "Tính trạng tương phản là",
        "options": [
            "A. các đặc điểm chỉ có ở cơ thể đực mà không có ở cơ thể cái.",
            "B. các đặc điểm bị biến đổi do ảnh hưởng của môi trường.",
            "C. hai trạng thái biểu hiện trái ngược nhau của cùng một loại tính trạng.",
            "D. hai trạng thái biểu hiện khác nhau của hai loại tính trạng hoàn toàn riêng biệt."
        ],
        "correct": "C",
        "solution": "Cặp tính trạng tương phản là hai trạng thái biểu hiện trái ngược, đối lập nhau của cùng một loại tính trạng (ví dụ: hoa đỏ - hoa trắng, thân cao - thân thấp)."
    },
    {
        "num": 7,
        "question": "Theo Mendel, mỗi tính trạng của cơ thể sinh vật do",
        "options": [
            "A. một nhân tố di truyền quy định.",
            "B. một cặp nhân tố di truyền quy định.",
            "C. hai nhân tố di truyền quy định.",
            "D. hai cặp nhân tố di truyền quy định."
        ],
        "correct": "B",
        "solution": "Theo giả thuyết của Mendel, mỗi tính trạng ở cơ thể sinh vật do một cặp nhân tố di truyền (sau này gọi là cặp allele) quy định."
    },
    {
        "num": 8,
        "question": "Hiện tượng truyền đạt các đặc điểm của bố mẹ, tổ tiên cho các thế hệ con cháu gọi là",
        "options": [
            "A. biến dị.",
            "B. nguyên phân.",
            "C. thụ tinh.",
            "D. di truyền."
        ],
        "correct": "D",
        "solution": "Di truyền là hiện tượng truyền đạt các tính trạng của bố mẹ, tổ tiên cho các thế hệ con cháu tiếp theo."
    },
    {
        "num": 9,
        "question": "Trong tế bào, DNA đảm nhận các chức năng nào sau đây?",
        "options": [
            "A. Truyền đạt thông tin di truyền.",
            "B. Cấu trúc nên nhân tế bào.",
            "C. Xúc tác các phản ứng sinh hóa.",
            "D. Lưu trữ, truyền đạt thông tin di truyền."
        ],
        "correct": "D",
        "solution": "Phân tử DNA có hai chức năng sinh học cốt lõi: Lưu trữ, bảo quản thông tin di truyền và Truyền đạt thông tin di truyền qua các thế hệ tế bào và cơ thể."
    },
    {
        "num": 10,
        "question": "Khi nói về cấu trúc không gian của phân tử DNA, phát biểu nào dưới đây không đúng?",
        "options": [
            "A. Trong mỗi mạch DNA, các đơn phân liên kết với nhau bằng liên kết cộng hóa trị.",
            "B. DNA có cấu trúc xoắn kép, gồm hai mạch song song, ngược chiều, xoắn quanh một trục từ phải sang trái.",
            "C. Giữa hai mạch DNA, các đơn phân liên kết hydrogen với nhau theo nguyên tắc bổ sung.",
            "D. Đơn phân cấu tạo nên phân tử DNA gồm A, T, C, G."
        ],
        "correct": "B",
        "solution": "Phát biểu B sai vì phân tử DNA xoắn đều quanh một trục tưởng tượng theo chiều từ trái sang phải (xoắn phải), chứ không phải từ phải sang trái."
    }
]

SINHHOC_TF = [
    {
        "num": 11,
        "intro": "Cho các phát biểu sau, nhận định phát biểu nào đúng/ sai:",
        "items": [
            {
                "label": "a",
                "text": "Dòng thuần là các cơ thể đồng hợp về tất cả các cặp gene. Dòng thuần có đặc tính di truyền đồng nhất và ổn định, các thế hệ sau giống các thế hệ trước.",
                "ans": "Đúng (Đ)",
                "reason": "Dòng thuần chủng mang các cặp gene đồng hợp nên tự thụ phấn hay giao phối gần không bị phân tính ở đời sau."
            },
            {
                "label": "b",
                "text": "Allele là các trạng thái biểu hiện giống nhau của cùng một gene.",
                "ans": "Sai (S)",
                "reason": "Allele là các trạng thái biểu hiện KHÁC NHAU của cùng một gene (ví dụ allele A quy định hoa đỏ, allele a quy định hoa trắng)."
            },
            {
                "label": "c",
                "text": "Nội dung quy luật phân li: Mỗi tính trạng do một cặp nhân tố di truyền (cặp allele) quy định. Khi giảm phân hình thành giao tử, các allele trong cặp phân li đồng đều về các giao tử nên mỗi giao tử chỉ chứa một allele của cặp.",
                "ans": "Đúng (Đ)",
                "reason": "Đây là phát biểu chính xác nội dung quy luật phân li của Mendel."
            },
            {
                "label": "d",
                "text": "Kết quả thí nghiệm lai một tính trạng được Mendel giải thích bằng sự phân li của cặp nhân tố di truyền.",
                "ans": "Đúng (Đ)",
                "reason": "Sự phân li đồng đều của cặp nhân tố di truyền trong giảm phân giải thích cho tỉ lệ phân li 3 trội : 1 lặn ở F2."
            },
            {
                "label": "e",
                "text": "Ở một loài thực vật, gene A quy định hoa đỏ, allele a quy định hoa trắng. Cho cây hoa đỏ (P) tự thụ phấn, F1 thu được cả hoa đỏ với hoa trắng với tỉ lệ kiểu gene là 1Aa : 1aa.",
                "ans": "Sai (S)",
                "reason": "Cây hoa đỏ P tự thụ phấn ra hoa trắng chứng tỏ P có kiểu gene dị hợp Aa. Phép lai Aa x Aa cho F1 có tỉ lệ kiểu gene là 1AA : 2Aa : 1aa."
            },
            {
                "label": "f",
                "text": "Phép lai Aa x Aa cho tỉ lệ kiểu hình ở thế hệ con là 1 : 1.",
                "ans": "Sai (S)",
                "reason": "Phép lai Aa x Aa cho tỉ lệ kiểu hình ở thế hệ con là 3 trội : 1 lặn (tỉ lệ 1 : 1 là của phép lai phân tích Aa x aa)."
            },
            {
                "label": "g",
                "text": "Ở đậu hà lan, allele A quy định thân cao trội hoàn toàn so với allele a quy định thân thấp. Muốn xác định cây thân cao có thuần chủng hay không, ta có thể cho cây thân cao thực hiện phép lai phân tích.",
                "ans": "Đúng (Đ)",
                "reason": "Cho lai với cây thân thấp (aa). Nếu đời con 100% thân cao -> thuần chủng (AA); nếu đời con phân tính 1 cao : 1 thấp -> không thuần chủng (Aa)."
            },
            {
                "label": "h",
                "text": "Kiểu gene cơ thể mang tính trạng trội của một cặp tính trạng là AA hoặc Aa.",
                "ans": "Đúng (Đ)",
                "reason": "Tính trạng trội biểu hiện ở trạng thái đồng hợp trội (AA) hoặc dị hợp (Aa)."
            },
            {
                "label": "i",
                "text": "Lai phân tích nhằm xác định kiểu hình của cơ thể mang tính trạng trội.",
                "ans": "Sai (S)",
                "reason": "Lai phân tích nhằm xác định KIỂU GENE (đồng hợp hay dị hợp) của cá thể mang kiểu hình trội."
            },
            {
                "label": "k",
                "text": "Các kiểu gene sau thuộc kiểu gene đồng hợp: AABB, AAbb, AaBB, aaBB.",
                "ans": "Sai (S)",
                "reason": "AaBB có cặp Aa là dị hợp, do đó không phải là kiểu gene đồng hợp về tất cả các cặp."
            },
            {
                "label": "l",
                "text": "Phân tử DNA có chức năng truyền đạt thông tin di truyền qua các thế hệ tế bào và cơ thể.",
                "ans": "Đúng (Đ)",
                "reason": "Nhờ cơ chế tự nhân đôi chính xác của DNA kết hợp nguyên phân, giảm phân và thụ tinh."
            },
            {
                "label": "m",
                "text": "Trên mỗi mạch DNA, các nucleotide liên kết với nhau bằng liên kết hydrogen theo nguyên tắc bổ sung.",
                "ans": "Sai (S)",
                "reason": "Trên MỖI MẠCH, các nucleotide liên kết với nhau bằng LIÊN KẾT CỘNG HÓA TRỊ (phosphodiester). Liên kết hydrogen theo nguyên tắc bổ sung là GIỮA HAI MẠCH đối diện."
            },
            {
                "label": "n",
                "text": "Phân tử DNA được cấu tạo từ 4 loại nucleotide (A, T, G, C), sự khác nhau về số lượng, thành phần và thứ tự sắp xếp của các nucleotide này tạo nên tính đa dạng và đặc thù của phân tử DNA ở mỗi loài sinh vật.",
                "ans": "Đúng (Đ)",
                "reason": "Đây là cơ sở phân tử của tính đa dạng và tính đặc thù sinh học ở các loài sinh vật."
            }
        ]
    }
]

SINHHOC_SHORT = [
    {
        "num": 12,
        "question": "Số loại giao tử có thể được tạo ra từ cơ thể có kiểu gene AaBbDD là bao nhiêu?",
        "solution": "Cơ thể có kiểu gene AaBbDD:\n- Cặp Aa: dị hợp (2 loại giao tử: A, a)\n- Cặp Bb: dị hợp (2 loại giao tử: B, b)\n- Cặp DD: đồng hợp (1 loại giao tử: D)\nSố loại giao tử tối đa tạo ra = 2^n = 2^2 = 4 loại giao tử (gồm: ABD, AbD, aBD, abD).\nĐáp số: 4."
    },
    {
        "num": 13,
        "question": "F1 có kiểu gene AaBb tiến hành tự thụ phấn thì số tổ hợp thu được ở F2 là bao nhiêu?",
        "solution": "Cơ thể F1 (AaBb) giảm phân tạo 2^2 = 4 loại giao tử đực và 4 loại giao tử cái.\nKhi tự thụ phấn, số tổ hợp giao tử ở F2 = 4 * 4 = 16 tổ hợp giao tử.\nĐáp số: 16."
    },
    {
        "num": 14,
        "question": "Ở đậu hà lan, allele A quy định hoa tím trội hoàn toàn so với allele a quy định hoa trắng. Cho cây hoa tím dị hợp Aa tự thụ phấn thu được thế hệ con. Tính theo lý thuyết, tỉ lệ phần trăm số cây hoa tím dị hợp thu được ở đời con là bao nhiêu %?",
        "solution": "Sơ đồ lai tự thụ phấn: Aa x Aa -> 1/4 AA : 2/4 Aa : 1/4 aa.\nTỉ lệ cây hoa tím mang kiểu gene dị hợp Aa ở đời con là: 2/4 = 50%.\nĐáp số: 50%."
    },
    {
        "num": 15,
        "question": "Ở đậu hà lan, gene A quy định hạt trơn, allele a quy định hạt nhăn. Cho cây hạt trơn không thuần chủng lai với cây hạt nhăn thì ở F1 kiểu hình hạt trơn chiếm tỉ lệ bao nhiêu %?",
        "solution": "- Cây hạt trơn không thuần chủng có kiểu gene: Aa.\n- Cây hạt nhăn có kiểu gene: aa.\nSơ đồ lai phân tích: Aa x aa -> 1/2 Aa (hạt trơn) : 1/2 aa (hạt nhăn).\nTỉ lệ kiểu hình hạt trơn ở F1 = 1/2 = 50%.\nĐáp số: 50%."
    }
]

SINHHOC_ESSAY = [
    {
        "num": 1,
        "question": "Ở loài ruồi giấm, thân xám, cánh dài là tính trạng trội hoàn toàn so với thân đen, cánh cụt. Các gene nằm trên các cặp nhiễm sắc thể khác nhau di truyền phân li độc lập. Cho P thuần chủng: thân xám, cánh cụt x thân đen, cánh dài. Hãy xác định kết quả của phép lai trên.",
        "solution": "1. Quy ước gene:\n- A: thân xám, a: thân đen.\n- B: cánh dài, b: cánh cụt.\n2. Kiểu gene của P thuần chủng:\n- Thân xám, cánh cụt thuần chủng: AAbb.\n- Thân đen, cánh dài thuần chủng: aaBB.\n3. Sơ đồ lai:\nP_tc: AAbb (thân xám, cánh cụt) x aaBB (thân đen, cánh dài)\nG_P: Ab x aB\nF1: 100% AaBb (100% thân xám, cánh dài).\nNếu cho F1 giao phối với nhau (F1 x F1: AaBb x AaBb):\nG_F1: 1/4 AB : 1/4 Ab : 1/4 aB : 1/4 ab\nF2:\n- Tỉ lệ kiểu gene: 9 A-B- : 3 A-bb : 3 aaB- : 1 aabb\n  (1 AABB : 2 AaBB : 2 AABb : 4 AaBb : 1 AAbb : 2 Aabb : 1 aaBB : 2 aaBb : 1 aabb).\n- Tỉ lệ kiểu hình: 9 thân xám, cánh dài : 3 thân xám, cánh cụt : 3 thân đen, cánh dài : 1 thân đen, cánh cụt."
    },
    {
        "num": 2,
        "question": "a. Trình tự các nucleotide trên một đoạn của DNA như sau: 3' ... T-T-G-C-T-C-A-T-C-A-C-G-T ... 5'\nHãy xác định trình tự các nucleotide trên mạch bổ sung với mạch đó.\nb. Nêu một số ứng dụng của phân tích DNA trong truy tìm tội phạm, xác định huyết thống.",
        "solution": "a. Theo nguyên tắc bổ sung (A liên kết với T, G liên kết với C và ngược lại), trình tự mạch bổ sung (chiều 5' -> 3') là:\nMạch gốc:       3' ... T - T - G - C - T - C - A - T - C - A - C - G - T ... 5'\nMạch bổ sung: 5' ... A - A - C - G - A - G - T - A - G - T - G - C - A ... 3'\nb. Ứng dụng của phân tích DNA:\n- Xác định huyết thống: So sánh dấu vân tay DNA giữa con và bố mẹ nghi vấn. Do mỗi người thừa hưởng 50% vật chất di truyền từ bố và 50% từ mẹ, phân tích DNA cho phép xác định chính xác quan hệ huyết thống trực hệ đạt độ tin cậy trên 99,99%.\n- Truy tìm tội phạm: Thu thập mẫu sinh học tại hiện trường (máu, chân tóc, tế bào da, nước bọt, tinh dịch...), tách chiết DNA để lập hồ sơ ADN cá thể, so sánh với nghi phạm để truy bắt hung thủ hoặc minh oan cho người vô tội.\n- Nhận dạng nạn nhân: Giúp xác định danh tính nạn nhân trong thiên tai, hỏa hoạn, chiến tranh khi thi thể không còn nguyên vẹn."
    },
    {
        "num": 3,
        "question": "a. Nêu khái niệm nucleic acid? Nucleic acid được cấu tạo như thế nào? Gồm những loại nào?\nb. Cho các đối tượng sau: Da, tóc, tiểu cầu, ti thể, virus HIV. Đối tượng nào có chứa nucleic acid?",
        "solution": "a. Khái niệm, cấu tạo và phân loại:\n- Khái niệm: Nucleic acid là đại phân tử sinh học cấu tạo theo nguyên tắc đa phân (đơn phân là nucleotide), đóng vai trò lưu trữ, bảo quản và truyền đạt thông tin di truyền.\n- Cấu tạo: Cấu tạo từ các nguyên tố C, H, O, N, P. Mỗi nucleotide gồm 3 phần: đường pentose (ribose hoặc deoxyribose), nhóm phosphate và nitrogenous base (A, T, G, C hoặc U).\n- Phân loại: Gồm 2 loại chính là DNA (Deoxyribonucleic acid) và RNA (Ribonucleic acid).\nb. Xác định các đối tượng có chứa nucleic acid:\n- Da (tế bào da): Chứa nhân tế bào -> Có chứa DNA và RNA.\n- Tóc: Phần chân tóc (nang tóc) chứa các tế bào sống -> Có chứa DNA và RNA (thân sợi tóc chết sừng hóa có thể chứa vết DNA dính theo).\n- Tiểu cầu: Mảnh tế bào chất không có nhân nhưng chứa RNA (mRNA, miRNA) và ti thể -> Có chứa nucleic acid.\n- Ti thể: Là bào quan có DNA vòng riêng (mtDNA) và RNA -> Có chứa nucleic acid.\n- Virus HIV: Là virus mang vật chất di truyền là 2 phân tử RNA sợi đơn -> Có chứa nucleic acid.\n=> Kết luận: Cả 5 đối tượng trên đều có chứa nucleic acid."
    },
    {
        "num": 4,
        "question": "Một gene có 480 adenine và 3120 liên kết hydrogen. Tính số lượng nucleotide của đoạn gene trên.",
        "solution": "Theo nguyên tắc bổ sung:\nA = T = 480 nucleotide.\nTổng số liên kết hydrogen của gene:\nH = 2A + 3G = 3120\n=> 2 * 480 + 3G = 3120\n=> 960 + 3G = 3120\n=> 3G = 3120 - 960 = 2160\n=> G = C = 2160 / 3 = 720 nucleotide.\nTổng số nucleotide của đoạn gene trên là:\nN = 2A + 2G = 2 * 480 + 2 * 720 = 960 + 1440 = 2400 nucleotide.\nĐáp số: 2400 nucleotide."
    },
    {
        "num": 5,
        "question": "Trình bày phương pháp và công thức tính số lượng các nucleotide, số liên kết hydrogen và số liên kết cộng hóa trị của phân tử DNA.",
        "solution": "1. Số lượng các loại nucleotide và tổng số nucleotide (N):\n- Theo nguyên tắc bổ sung: A = T và G = C.\n- Tổng số nucleotide: N = A + T + G + C = 2A + 2G = 2T + 2C.\n- Tỉ lệ %: %A + %G = %T + %C = 50% N.\n- Chiều dài phân tử DNA: L = (N / 2) * 3,4 Å (1 nm = 10 Å, 1 μm = 10 000 Å).\n- Chu kì xoắn của DNA: C = N / 20.\n2. Số liên kết hydrogen (H):\n- Giữa A và T có 2 liên kết hydrogen (A = T).\n- Giữa G và C có 3 liên kết hydrogen (G ≡ C).\n- Công thức: H = 2A + 3G = 2T + 3C.\n3. Số liên kết cộng hóa trị (phosphodiester):\n- Số liên kết cộng hóa trị nối giữa các nucleotide trên 2 mạch:\n  H_giua_nu = 2 * (N/2 - 1) = N - 2 liên kết.\n- Tổng số liên kết cộng hóa trị trong toàn bộ phân tử DNA (tính cả liên kết giữa đường và phosphate trong mỗi nucleotide):\n  H_toan_bo = (N - 2) + N = 2N - 2 liên kết."
    }
]
