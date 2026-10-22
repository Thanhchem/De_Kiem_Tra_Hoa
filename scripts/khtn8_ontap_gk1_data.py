# -*- coding: utf-8 -*-
"""
Dữ liệu chuẩn hóa Đề cương Ôn tập Giữa học kì I - Môn Khoa học tự nhiên 8 (2026 - 2027)
Gồm 3 phân môn:
I. Hóa học (Chất và sự biến đổi chất): Bài 1 - Bài 5
II. Sinh học (Vật sống): Bài 30 - Bài 32
III. Vật lí (Năng lượng và sự biến đổi): Bài 13 - Bài 15
Thầy TRẦN VĂN THẠNH - 0777.470.803
"""

INFO = {
    "truong": "TRƯỜNG THCS",
    "to": "TỔ: KHOA HỌC TỰ NHIÊN",
    "tieu_de": "HƯỚNG DẪN ÔN TẬP GIỮA HỌC KỲ I",
    "mon": "KHOA HỌC TỰ NHIÊN 8",
    "nam_hoc": "NĂM HỌC: 2026 - 2027",
    "giao_vien": "Thầy TRẦN VĂN THẠNH - 0777.470.803",
    "co_so_1": "CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế",
    "co_so_2": "CS2: 24 Đặng Thái Thân, TP Huế",
}

# ==============================================================================
# I. PHÂN MÔN HÓA HỌC (CHẤT VÀ SỰ BIẾN ĐỔI CHẤT)
# ==============================================================================
HOA_THEORY = [
    {
        "title": "Bài 1. Sử dụng một số hoá chất, thiết bị cơ bản trong phòng thí nghiệm",
        "content": [
            "Nhận biết được một số dụng cụ và hoá chất sử dụng trong môn Khoa học tự nhiên 8 (dụng cụ đo thể tích, đun nóng, chứa đựng; ống nghiệm, bình tam giác, ống hút nhỏ giọt, đèn cồn...).",
            "Nêu được quy tắc sử dụng hoá chất an toàn (chủ yếu những hoá chất trong môn Khoa học tự nhiên 8; đọc kĩ nhãn cảnh báo nguy hiểm, không nếm, không ngửi trực tiếp, dùng đúng liều lượng).",
            "Nhận biết được các thiết bị điện trong môn Khoa học tự nhiên 8 (nguồn điện, công tắc, dây dẫn, ampe kế, vôn kế, điện trở, bóng đèn...).",
            "Trình bày được cách sử dụng điện an toàn trong thí nghiệm và đời sống."
        ]
    },
    {
        "title": "Bài 2. Phản ứng hoá học",
        "content": [
            "Khái niệm biến đổi vật lí và biến đổi hoá học: Biến đổi vật lí là hiện tượng chất biến đổi về trạng thái, hình dạng nhưng không sinh ra chất mới; Biến đổi hoá học là hiện tượng chất biến đổi có tạo ra chất mới.",
            "Khái niệm phản ứng hoá học, chất đầu (chất phản ứng) và sản phẩm. Chất đầu bị tiêu hao, sản phẩm được tạo thành.",
            "Dấu hiệu chứng tỏ có phản ứng hoá học xảy ra: có sự thay đổi màu sắc, xuất hiện chất kết tủa, giải phóng chất khí, tỏa nhiệt hoặc phát sáng.",
            "Khái niệm phản ứng toả nhiệt và phản ứng thu nhiệt; ứng dụng của phản ứng toả nhiệt trong đời sống và sản xuất (đốt than củi, khí gas để đun nấu, chạy động cơ...)."
        ]
    },
    {
        "title": "Bài 3. Mol và tỉ khối chất khí",
        "content": [
            "Khái niệm mol: Mol là lượng chất có chứa NA (6,022.10^23) nguyên tử hoặc phân tử của chất đó.",
            "Biểu thức tính số mol: n = m / M hoặc n = V / 24,79 (ở điều kiện chuẩn: 25 oC, 1 bar).",
            "Chuyển đổi giữa số mol, khối lượng (m = n . M) và thể tích chất khí (V = n . 24,79).",
            "Khái niệm tỉ khối chất khí: Biểu thức dA/B = MA / MB; tỉ khối của khí A đối với không khí: dA/kk = MA / 29. Dựa vào tỉ khối để so sánh khí A nặng hay nhẹ hơn khí B, nặng hay nhẹ hơn không khí."
        ]
    },
    {
        "title": "Bài 4. Dung dịch và nồng độ",
        "content": [
            "Chất tan là chất bị hòa tan trong dung môi. Dung môi là chất có khả năng hòa tan chất khác để tạo thành dung dịch. Dung dịch là hỗn hợp đồng nhất giữa dung môi và chất tan.",
            "Độ tan (S) của một chất trong nước là số gam chất đó tan được trong 100 gam nước để tạo thành dung dịch bão hòa ở nhiệt độ xác định. Biểu thức: S = (m_ct / m_dm) . 100.",
            "Nồng độ phần trăm (C%): C% = (m_ct / m_dd) . 100% (với m_dd = m_ct + m_dm).",
            "Nồng độ mol (CM): CM = n / V (với V tính bằng lít).",
            "Các bước tính toán và tiến hành thí nghiệm pha chế một dung dịch theo nồng độ cho trước."
        ]
    },
    {
        "title": "Bài 5. Định luật bảo toàn khối lượng và phương trình hoá học",
        "content": [
            "Định luật bảo toàn khối lượng: Trong một phản ứng hoá học, tổng khối lượng của các chất sản phẩm bằng tổng khối lượng của các chất tham gia phản ứng. (m_A + m_B = m_C + m_D).",
            "Giải thích định luật: Trong phản ứng hoá học, chỉ có liên kết giữa các nguyên tử thay đổi làm phân tử này biến thành phân tử khác, còn số nguyên tử của mỗi nguyên tố được giữ nguyên, nên tổng khối lượng được bảo toàn.",
            "Khái niệm và các bước lập phương trình hoá học: (1) Viết sơ đồ phản ứng; (2) Cân bằng số nguyên tử mỗi nguyên tố ở hai vế; (3) Viết phương trình hoá học.",
            "Ý nghĩa của phương trình hoá học: Cho biết tỉ lệ về số nguyên tử, số phân tử giữa các chất trong phản ứng."
        ]
    }
]

HOA_MCQ = [
    {
        "num": 1,
        "question": "Chất được tạo thành sau phản ứng hóa học là?",
        "options": [
            "A. Chất phản ứng.",
            "B. Chất lỏng.",
            "C. Chất sản phẩm.",
            "D. Chất khí."
        ],
        "correct": "C",
        "solution": "Chất ban đầu tham gia phản ứng gọi là chất phản ứng (chất tham gia); chất mới được sinh ra sau phản ứng gọi là sản phẩm."
    },
    {
        "num": 2,
        "question": "Phản ứng tỏa nhiệt là:",
        "options": [
            "A. Phản ứng có nhiệt độ lớn hơn môi trường xung quanh",
            "B. Phản ứng có nhiệt độ nhỏ hơn môi trường xung quanh",
            "C. Phản ứng có nhiệt độ bằng môi trường xung quanh",
            "D. Phản ứng không có sự thay đổi nhiệt độ"
        ],
        "correct": "A",
        "solution": "Phản ứng tỏa nhiệt giải phóng năng lượng nhiệt ra môi trường làm hệ có nhiệt độ lớn hơn môi trường xung quanh (truyền nhiệt ra môi trường)."
    },
    {
        "num": 3,
        "question": "Hòa tan đường vào nước là:",
        "options": [
            "A. Phản ứng hóa học.",
            "B. Phản ứng tỏa nhiệt.",
            "C. Phản ứng thu nhiệt.",
            "D. Sự biến đổi vật lí."
        ],
        "correct": "D",
        "solution": "Khi hòa tan đường vào nước, đường chỉ phân tán thành các phân tử đường xen kẽ phân tử nước, không có chất mới sinh ra nên là sự biến đổi vật lí."
    },
    {
        "num": 4,
        "question": "Chất mới được tạo ra từ phản ứng hóa học so với chất cũ sẽ như thế nào?",
        "options": [
            "A. Có tính chất mới, khác biệt chất ban đầu",
            "B. Giống hệt chất ban đầu",
            "C. A và B đều đúng",
            "D. A và B đều sai"
        ],
        "correct": "A",
        "solution": "Phản ứng hóa học là quá trình biến đổi chất này thành chất khác, do đó sản phẩm tạo thành có tính chất mới, khác biệt hoàn toàn với chất ban đầu."
    },
    {
        "num": 5,
        "question": "Than (thành phần chính là carbon) cháy trong không khí tạo thành khí carbon dioxide. Trong quá trình phản ứng, lượng chất nào tăng dần?",
        "options": [
            "A. Carbon dioxide tăng dần.",
            "B. Oxygen tăng dần",
            "C. Carbon tăng dần.",
            "D. Tất cả đều tăng"
        ],
        "correct": "A",
        "solution": "Trong phản ứng: C + O2 -> CO2. Các chất phản ứng (C và O2) bị tiêu hao nên lượng chất giảm dần; sản phẩm CO2 mới sinh ra nên lượng chất tăng dần."
    }
]

HOA_TF = [
    {
        "num": 1,
        "intro": "Xét tính Đúng / Sai của các khẳng định sau về Mol, khối lượng mol và thể tích chất khí:",
        "items": [
            {
                "label": "a",
                "text": "Mol là lượng chất có chứa NA nguyên tử hoặc phân tử của chất đó.",
                "ans": "Đúng (Đ)",
                "reason": "Theo định nghĩa chuẩn về mol trong SGK KHTN 8 (NA = 6,022.10^23)."
            },
            {
                "label": "b",
                "text": "Khối lượng mol của một chất là khối lượng tính bằng gam của NA nguyên tử hoặc phân tử chất đó.",
                "ans": "Đúng (Đ)",
                "reason": "Khối lượng mol (M) có trị số bằng phân tử khối hoặc nguyên tử khối, đơn vị là g/mol."
            },
            {
                "label": "c",
                "text": "Thể tích của 1 mol bất kỳ chất khí nào ở điều kiện chuẩn cũng luôn bằng 22,4 lít.",
                "ans": "Sai (S)",
                "reason": "Ở điều kiện chuẩn mới (25 oC, 1 bar), 1 mol khí bất kì chiếm thể tích 24,79 lít. (22,4 lít là ở điều kiện tiêu chuẩn cũ 0 oC, 1 atm)."
            },
            {
                "label": "d",
                "text": "Ở cùng điều kiện nhiệt độ và áp suất, 1 mol của mọi chất khí đều chiếm những thể tích bằng nhau.",
                "ans": "Đúng (Đ)",
                "reason": "Đây là nội dung định luật Avogadro."
            },
            {
                "label": "e",
                "text": "Trong 0,12 mol phân tử NH3 có chứa 0,24 mol nguyên tử N.",
                "ans": "Sai (S)",
                "reason": "Trong 1 phân tử NH3 có 1 nguyên tử N, nên trong 0,12 mol NH3 chỉ có đúng 0,12 mol nguyên tử N."
            }
        ]
    }
]

HOA_SHORT = [
    {
        "num": "a",
        "question": "Cho sơ đồ phản ứng sau: Fe + S --to--> FeS. Nếu dùng 7g sắt tác dụng vừa đủ với 4g lưu huỳnh thì khối lượng muối sắt(II) sulfide (FeS) thu được là bao nhiêu gam?",
        "solution": "Theo định luật bảo toàn khối lượng:\nm_Fe + m_S = m_FeS\n=> m_FeS = 7 + 4 = 11 gam.\nĐáp số: 11 gam."
    },
    {
        "num": "b",
        "question": "Nung nóng hoàn toàn 12g đá vôi CaCO3 thu được 6,72g khí carbon dioxide CO2 và một lượng rắn calcium oxide (CaO). Khối lượng của CaO thu được là bao nhiêu gam?",
        "solution": "PTHH: CaCO3 --to--> CaO + CO2\nTheo định luật bảo toàn khối lượng:\nm_CaCO3 = m_CaO + m_CO2\n=> m_CaO = m_CaCO3 - m_CO2 = 12 - 6,72 = 5,28 gam.\nĐáp số: 5,28 gam."
    },
    {
        "num": "c",
        "question": "Đốt cháy hoàn toàn 48g lưu huỳnh (S) trong khí oxygen thu được 96g khí sulfur dioxide (SO2). Khối lượng khí oxygen (O2) đã tham gia phản ứng là bao nhiêu gam?",
        "solution": "PTHH: S + O2 --to--> SO2\nTheo định luật bảo toàn khối lượng:\nm_S + m_O2 = m_SO2\n=> m_O2 = m_SO2 - m_S = 96 - 48 = 48 gam.\nĐáp số: 48 gam."
    }
]

HOA_ESSAY = [
    {
        "num": 1,
        "question": "Xác định độ tan của muối Na2CO3 trong nước ở 18 oC. Biết rằng ở nhiệt độ này khi hòa tan hết 53g Na2CO3 trong 250g nước thì được dung dịch bão hòa.",
        "solution": "Độ tan (S) của một chất là số gam chất đó tan được trong 100 gam nước để tạo thành dung dịch bão hòa ở nhiệt độ xác định.\nCông thức tính độ tan:\nS = (m_ct / m_H2O) * 100\nThay số vào công thức:\nS = (53 / 250) * 100 = 21,2 gam/100 gam H2O.\nVậy độ tan của muối Na2CO3 ở 18 oC là 21,2 gam."
    },
    {
        "num": 2,
        "question": "Hòa tan 8 gam đường vào nước được dung dịch có nồng độ 5%. Hãy tính:\na. Khối lượng dung dịch đường pha chế được.\nb. Khối lượng nước cần dùng cho sự pha chế.\nc. Cần thêm vào bao nhiêu gam nước để được dung dịch có nồng độ 4%.\nd. Cần thêm vào bao nhiêu gam đường để được dung dịch có nồng độ 6%.",
        "solution": "a. Khối lượng dung dịch đường pha chế được là:\nm_dd = (m_ct * 100%) / C% = (8 * 100) / 5 = 160 gam.\nb. Khối lượng nước cần dùng cho sự pha chế:\nm_H2O = m_dd - m_ct = 160 - 8 = 152 gam.\nc. Gọi khối lượng nước cần thêm vào là m1 (gam, m1 > 0).\nLượng đường không đổi là 8 gam. Khối lượng dung dịch mới là: m_dd1 = 160 + m1 (gam).\nNồng độ dung dịch mới là 4%:\nC% = [8 / (160 + m1)] * 100% = 4%\n=> 160 + m1 = (8 * 100) / 4 = 200\n=> m1 = 200 - 160 = 40 gam.\nVậy cần thêm vào 40 gam nước.\nd. Gọi khối lượng đường cần thêm vào là m2 (gam, m2 > 0).\nKhối lượng chất tan mới: m_ct2 = 8 + m2 (gam).\nKhối lượng dung dịch mới: m_dd2 = 160 + m2 (gam).\nNồng độ dung dịch mới là 6%:\nC% = [(8 + m2) / (160 + m2)] * 100% = 6%\n=> 100 * (8 + m2) = 6 * (160 + m2)\n=> 800 + 100m2 = 960 + 6m2\n=> 94m2 = 160 => m2 = 160 / 94 ≈ 1,702 gam.\nVậy cần thêm vào khoảng 1,702 gam đường."
    },
    {
        "num": 3,
        "question": "Tính số mol và số gam chất tan có trong 250 mL dung dịch CaCl2 0,3M. (Cl = 35,5; Ca = 40)",
        "solution": "- Đổi thể tích dung dịch: V = 250 mL = 0,25 lít.\n- Số mol chất tan CaCl2 có trong dung dịch là:\nn_CaCl2 = CM * V = 0,3 * 0,25 = 0,075 mol.\n- Khối lượng mol của CaCl2: M_CaCl2 = 40 + 35,5 * 2 = 111 g/mol.\n- Khối lượng chất tan CaCl2 là:\nm_CaCl2 = n * M = 0,075 * 111 = 8,325 gam.\nĐáp số: 0,075 mol và 8,325 gam."
    },
    {
        "num": 4,
        "question": "Trong 800 mL của một dung dịch có chứa 8g NaOH. (H = 1, O = 16, Na = 23)\na. Hãy tính nồng độ mol của dung dịch này.\nb. Phải thêm bao nhiêu mL nước vào 200 mL dung dịch này để được dung dịch NaOH 0,1M.",
        "solution": "a. Tính nồng độ mol dung dịch ban đầu:\n- Khối lượng mol NaOH: M_NaOH = 23 + 16 + 1 = 40 g/mol.\n- Số mol NaOH trong 800 mL dung dịch: n_NaOH = 8 / 40 = 0,2 mol.\n- Đổi V = 800 mL = 0,8 lít.\n- Nồng độ mol dung dịch là:\nCM = n / V = 0,2 / 0,8 = 0,25M.\nb. Thêm nước vào 200 mL dung dịch:\n- Đổi 200 mL = 0,2 lít.\n- Số mol NaOH có trong 200 mL dung dịch 0,25M là:\nn_NaOH = 0,25 * 0,2 = 0,05 mol.\n- Khi thêm nước, lượng NaOH không đổi. Thể tích dung dịch mới có nồng độ 0,1M là:\nV_dd mới = n / CM mới = 0,05 / 0,1 = 0,5 lít = 500 mL.\n- Thể tích nước cần thêm vào là:\nV_H2O cần thêm = 500 - 200 = 300 mL.\nVậy phải thêm 300 mL nước."
    },
    {
        "num": 5,
        "question": "Hãy chọn hệ số và công thức hóa học thích hợp đặt vào những chỗ có dấu hỏi trong các phương trình hóa học sau:\na.  ?  +  ?  --to--> 2FeCl3\nb.  FexOy  +  ?H2  --to--> ?Fe  +  ?H2O\nc.  ?Al  +  ?  --to--> 2Al2O3\nd.  ?Ca(OH)2  +  ?P2O5  --> ?Ca3(PO4)2  +  ?H2O\ne.  CnH2n+2  +  ?O2  --to--> ?CO2  +  ?H2O\ng.  ?Al  +  ?CuxOy  --to--> ?Al2O3  +  ?Cu",
        "solution": "a. 2Fe + 3Cl2 --to--> 2FeCl3\nb. FexOy + yH2 --to--> xFe + yH2O\nc. 4Al + 3O2 --to--> 2Al2O3\nd. 3Ca(OH)2 + P2O5 --> Ca3(PO4)2 + 3H2O\ne. 2CnH2n+2 + (3n+1)O2 --to--> 2nCO2 + 2(n+1)H2O  (hoặc CnH2n+2 + [(3n+1)/2]O2 --to--> nCO2 + (n+1)H2O)\ng. 2yAl + 3CuxOy --to--> yAl2O3 + 3xCu"
    },
    {
        "num": 6,
        "question": "Tính số mol và số gam chất tan có trong 200 mL dung dịch KCl 0,25M. (Cl = 35,5, K = 39)",
        "solution": "- Đổi thể tích: V = 200 mL = 0,2 lít.\n- Số mol KCl có trong dung dịch là:\nn_KCl = CM * V = 0,25 * 0,2 = 0,05 mol.\n- Khối lượng mol của KCl: M_KCl = 39 + 35,5 = 74,5 g/mol.\n- Khối lượng chất tan KCl là:\nm_KCl = n * M = 0,05 * 74,5 = 3,725 gam.\nĐáp số: 0,05 mol và 3,725 gam."
    },
    {
        "num": 7,
        "question": "Trong 400 mL của một dung dịch có chứa 16g CuSO4. (O = 16, S = 32, Cu = 64)\na. Hãy tính nồng độ mol của dung dịch này.\nb. Phải thêm bao nhiêu mL nước vào 200 mL dung dịch này để được dung dịch CuSO4 0,2M.",
        "solution": "a. Tính nồng độ mol của dung dịch ban đầu:\n- Khối lượng mol CuSO4: M = 64 + 32 + 16 * 4 = 160 g/mol.\n- Số mol CuSO4 trong 400 mL: n = 16 / 160 = 0,1 mol.\n- Đổi V = 400 mL = 0,4 lít.\n- Nồng độ mol dung dịch là:\nCM = n / V = 0,1 / 0,4 = 0,25M.\nb. Thêm nước vào 200 mL dung dịch:\n- Đổi 200 mL = 0,2 lít.\n- Số mol CuSO4 có trong 200 mL dung dịch là:\nn = 0,2 * 0,25 = 0,05 mol.\n- Thể tích dung dịch mới có nồng độ 0,2M là:\nV_dd mới = n / CM mới = 0,05 / 0,2 = 0,25 lít = 250 mL.\n- Thể tích nước cần thêm vào là:\nV_H2O cần thêm = 250 - 200 = 50 mL.\nVậy phải thêm 50 mL nước."
    },
    {
        "num": 8,
        "question": "Hãy tính toán và pha chế 50 gam dung dịch NaCl có nồng độ 2%. Dụng cụ và hoá chất coi như có đủ. (Na = 23, Cl = 35,5)",
        "solution": "1. Phần tính toán:\n- Khối lượng chất tan NaCl cần dùng là:\nm_NaCl = (m_dd * C%) / 100% = (50 * 2) / 100 = 1 gam.\n- Khối lượng nước cất cần dùng là:\nm_H2O = m_dd - m_ct = 50 - 1 = 49 gam (tương đương với 49 mL nước cất vì D_H2O = 1 g/mL).\n2. Phần tiến hành pha chế:\n- Dụng cụ: Cân điện tử, ống đong (hoặc cốc đong), cốc thủy tinh 100 mL, đũa thủy tinh.\n- Hoá chất: Muối ăn NaCl tinh khiết, nước cất.\n- Các bước tiến hành:\n+ Bước 1: Cân lấy đúng 1 gam NaCl khan cho vào cốc thủy tinh dung tích 100 mL.\n+ Bước 2: Dùng ống đong đo lấy 49 mL nước cất, rót từ từ vào cốc đựng NaCl.\n+ Bước 3: Dùng đũa thủy tinh khuấy đều cho đến khi toàn bộ muối NaCl tan hết, ta thu được đúng 50 gam dung dịch NaCl 2%."
    }
]

# ==============================================================================
# II. PHẦN SINH HỌC (VẬT SỐNG)
# ==============================================================================
SINHHOC_THEORY = [
    {
        "title": "Bài 30: Khái quát về cơ thể người",
        "content": [
            "I. Khái quát về cơ thể người:",
            "- Cơ thể người bao gồm các phần: Đầu, cổ, thân, hai tay và hai chân.",
            "- Toàn bộ cơ thể được bao bọc bên ngoài bởi một lớp da, dưới lớp da là lớp mỡ, dưới lớp mỡ là cơ và xương.",
            "II. Vai trò của các cơ quan và hệ cơ quan trong cơ thể người:",
            "- Các hệ cơ quan gồm: hệ vận động, hệ tuần hoàn, hệ hô hấp, hệ tiêu hóa, hệ bài tiết, hệ thần kinh và các giác quan, hệ nội tiết, hệ sinh dục.",
            "- Mỗi cơ quan, hệ cơ quan có một vai trò nhất định đối với cơ thể. Các cơ quan, hệ cơ quan trong cơ thể có mối liên quan chặt chẽ, phối hợp nhịp nhàng với nhau để thực hiện các chức năng sống của cơ thể thống nhất."
        ]
    },
    {
        "title": "Bài 31: Hệ vận động ở người",
        "content": [
            "I. Cấu tạo và chức năng của hệ vận động:",
            "1. Cấu tạo: Gồm bộ xương và hệ cơ.",
            "- Xương được cấu tạo từ chất hữu cơ và chất khoáng. Chất khoáng làm xương bền chắc, chất hữu cơ giúp xương có tính mềm dẻo đàn hồi.",
            "- Cấu tạo hệ cơ: Hệ cơ ở người có khoảng 600 cơ.",
            "2. Chức năng:",
            "- Hệ vận động có chức năng bảo vệ, duy trì hình dạng và vận động cơ thể:",
            "+ Bộ xương tạo nên khung cơ thể, giúp cơ thể có hình dạng nhất định và bảo vệ các nội quan bên trong.",
            "+ Cơ bám vào xương, khi cơ co hay dãn sẽ làm xương cử động, giúp cơ thể di chuyển và vận động.",
            "+ Một số khớp xương tạo kết nối kiểu đòn bẩy giữa các xương, nhờ vậy, xương có khả năng chịu tải cao khi vận động.",
            "II. Một số bệnh, tật liên quan đến hệ vận động:",
            "1. Tật cong vẹo cột sống: Biểu hiện các đốt sống bị xoay lệch về một bên, cong quá mức về phía trước/sau. Nguyên nhân do tư thế sai trong thời gian dài, mang vác nặng, tai nạn, còi xương. Phòng chống: Ngồi học ngay ngắn, đeo cặp 2 vai, tập luyện vừa sức, dinh dưỡng hợp lí.",
            "2. Bệnh loãng xương: Biểu hiện mật độ chất khoáng trong xương thưa dần. Nguyên nhân do cơ thể thiếu calcium và phosphorus làm thiếu nguyên liệu kiến tạo xương. Phòng chống: Tập TDTT đều đặn, ăn đủ đạm, canxi, vitamin D, tắm nắng hợp lí, không dùng chất kích thích.",
            "III. Ý nghĩa của tập thể dục, thể thao: Kích thích tăng chiều dài và chu vi của xương, cơ bắp nở nang rắn chắc, tăng cường sự dẻo dai của cơ thể (đi bộ, chạy bộ, đạp xe, bơi lội, cầu lông, bóng đá...).",
            "IV. Thực hành: Sơ cứu và băng bó khi người khác bị gãy xương (Đặt nẹp cố định 2 đầu khớp, đệm lót gạc êm, băng cố định nẹp nhẹ nhàng chắc chắn, đưa đến cơ sở y tế)."
        ]
    },
    {
        "title": "Bài 32: Dinh dưỡng và tiêu hóa ở người",
        "content": [
            "I. Khái niệm chất dinh dưỡng và dinh dưỡng:",
            "- Chất dinh dưỡng là các chất có trong thức ăn mà cơ thể sử dụng làm nguyên liệu cấu tạo cơ thể và cung cấp năng lượng cho các hoạt động sống (carbohydrate, protein, lipid, vitamin, chất khoáng, nước).",
            "- Dinh dưỡng là quá trình thu nhận, biến đổi và sử dụng chất dinh dưỡng để duy trì sự sống của cơ thể.",
            "II. Tiêu hóa ở người:",
            "1. Cấu tạo: Ống tiêu hóa (miệng, hầu, thực quản, dạ dày, ruột non, ruột già, trực tràng, hậu môn) và tuyến tiêu hóa (tuyến nước bọt, tuyến vị, gan, mật, tụy, tuyến ruột).",
            "2. Quá trình tiêu hóa:",
            "- Khoang miệng: Tiêu hóa cơ học (nhai, nghiền, đảo trộn) và hóa học (enzyme amylase biến đổi một phần tinh bột chín thành đường maltose).",
            "- Dạ dày: Tiêu hóa cơ học (co bóp nhuyễn thức ăn, thấm dịch vị) và hóa học (enzyme pepsin biến đổi một phần protein).",
            "- Ruột non: Biến đổi chủ yếu về mặt hóa học nhờ dịch tụy, dịch mật, dịch ruột biến đổi thức ăn thành các chất đơn giản để hấp thụ. Hấp thụ chất dinh dưỡng nhờ mạng mao mạch máu và bạch huyết dày đặc ở lông ruột.",
            "- Ruột già và trực tràng: Hấp thụ lại nước và một số muối khoáng, vi khuẩn lên men phân hủy chất bã tạo thành phân thải ra ngoài.",
            "III. Một số bệnh về đường tiêu hóa: Sâu răng (vi khuẩn phân hủy thức ăn tạo acid ăn mòn men răng); Viêm loét dạ dày - tá tràng (nhiễm khuẩn HP, ăn uống không điều độ, stress).",
            "IV. Chế độ dinh dưỡng: Xây dựng khẩu phần ăn hợp lí đảm bảo đủ lượng thức ăn, cân đối 4 nhóm dinh dưỡng và cung cấp đủ năng lượng cho cơ thể."
        ]
    }
]

# ==============================================================================
# III. PHÂN MÔN VẬT LÍ (NĂNG LƯỢNG VÀ SỰ BIẾN ĐỔI)
# ==============================================================================
VATLI_THEORY = [
    {
        "title": "Chương III: Khối lượng riêng và áp suất",
        "content": [
            "Bài 13, 14: Khối lượng riêng và thực hành xác định khối lượng riêng:",
            "- Khối lượng riêng của một chất là khối lượng của một đơn vị thể tích chất đó: D = m / V (đơn vị: kg/m^3 hoặc g/cm^3; 1 g/cm^3 = 1000 kg/m^3).",
            "- Trọng lượng riêng: d = P / V = 10D (đơn vị: N/m^3).",
            "- Cách xác định khối lượng riêng: Dùng cân đo khối lượng m; dùng thước đo kích thước (với vật hình học xác định) hoặc dùng bình chia độ / bình tràn (với vật rắn bất kì và chất lỏng) để đo thể tích V, sau đó tính D = m / V.",
            "Bài 15: Áp suất trên một bề mặt:",
            "- Áp lực là lực ép có phương vuông góc với mặt bị ép. Tác dụng của áp lực phụ thuộc vào độ lớn của áp lực và diện tích bề mặt bị ép.",
            "- Áp suất là độ lớn của áp lực trên một đơn vị diện tích bị ép: p = F / S (đơn vị: N/m^2 hoặc Pascal (Pa); 1 Pa = 1 N/m^2).",
            "- Biện pháp tăng áp suất: Tăng áp lực F hoặc giảm diện tích bị ép S (ví dụ: mài sắc mũi kim, lưỡi dao, cọc vót nhọn...).",
            "- Biện pháp giảm áp suất: Giảm áp lực F hoặc tăng diện tích bị ép S (ví dụ: bản xích xe tăng rộng, làm móng nhà to bản, ván trượt tuyết rộng...).",
            "- Lưu ý trọng lực của vật đặt trên mặt phẳng nằm ngang đóng vai trò là áp lực: F = P = 10.m."
        ]
    }
]

VATLI_MCQ = [
    {
        "num": 1,
        "question": "Đơn vị đo của khối lượng riêng trong hệ SI là:",
        "options": [
            "A. kg/m3",
            "B. N",
            "C. kg",
            "D. N/m2"
        ],
        "correct": "A",
        "solution": "Trong hệ đơn vị đo lường quốc tế SI, đơn vị của khối lượng riêng là kilôgam trên mét khối (kg/m^3)."
    },
    {
        "num": 2,
        "question": "Muốn đo khối lượng riêng của viên bi người ta dùng những dụng cụ gì?",
        "options": [
            "A. Chỉ cần dùng một cái cân",
            "B. Chỉ cần dùng một lực kế",
            "C. Cần dùng một cái cân và bình chia độ",
            "D. Chỉ cần dùng một bình chia độ"
        ],
        "correct": "C",
        "solution": "Để xác định khối lượng riêng D = m / V, cần dùng cân để đo khối lượng m của viên bi và bình chia độ để đo thể tích V của viên bi."
    },
    {
        "num": 3,
        "question": "Đơn vị của áp suất là:",
        "options": [
            "A. Pa",
            "B. N",
            "C. m",
            "D. N/m3"
        ],
        "correct": "A",
        "solution": "Đơn vị đo áp suất trong hệ SI là Pascal (Pa) hoặc Newton trên mét vuông (N/m^2). 1 Pa = 1 N/m^2."
    },
    {
        "num": 4,
        "question": "Để xác định khối lượng riêng của các viên sỏi, ba bạn Sử, Sen, Anh đưa ra ý kiến như sau:\nSử: Mình chỉ cần một cái cân là đủ\nSen: Theo mình, cần một bình chia độ mới đúng\nAnh: Sai bét, phải cần một cái cân và một bình chia độ mới xác định được chứ\nTheo em, ý kiến nào đúng?",
        "options": [
            "A. Sử đúng.",
            "B. Sen đúng.",
            "C. Anh đúng.",
            "D. Cả ba bạn cùng sai"
        ],
        "correct": "C",
        "solution": "Vì viên sỏi có hình dạng bất kì, muốn tính D = m / V ta cần một cái cân để đo khối lượng m và bình chia độ để đo thể tích V (thả chìm trong nước). Do đó bạn Anh đúng."
    },
    {
        "num": 5,
        "question": "Khối lượng riêng của sắt là 7800kg/m3. Vậy, 1kg sắt sẽ có thể tích vào khoảng",
        "options": [
            "A. 12,8cm3",
            "B. 128cm3.",
            "C. 1.280cm3.",
            "D. 12.800cm3"
        ],
        "correct": "B",
        "solution": "Áp dụng công thức: V = m / D = 1 / 7800 m^3 = 1 000 000 / 7800 cm^3 ≈ 128,2 cm^3 ≈ 128 cm^3."
    },
    {
        "num": 6,
        "question": "Phát biểu nào sau đây về khối lượng riêng là đúng?",
        "options": [
            "A. Khối lượng riêng của một chất là khối lượng của một đơn vị thể tích chất đó.",
            "B. Nói khối lượng riêng của sắt là 7800 kg/m3 có nghĩa là 1 cm3 sắt có khối lượng 7800 kg.",
            "C. Công thức tính khối lượng riêng là D = m.V.",
            "D. Khối lượng riêng bằng trọng lượng riêng."
        ],
        "correct": "A",
        "solution": "Theo định nghĩa chuẩn: Khối lượng riêng của một chất là khối lượng của một đơn vị thể tích (1 m^3 hoặc 1 cm^3) chất đó. Công thức là D = m / V."
    },
    {
        "num": 7,
        "question": "Muốn tăng áp suất tác dụng lên mặt bị ép, ta có thể:",
        "options": [
            "A. Giảm áp lực hoặc tăng diện tích mặt bị ép.",
            "B. Tăng độ lớn áp lực hoặc giảm diện tích mặt bị ép.",
            "C. Giữ nguyên áp lực và tăng diện tích mặt bị ép.",
            "D. Giảm đồng thời áp lực và diện tích mặt bị ép."
        ],
        "correct": "B",
        "solution": "Từ công thức p = F / S, muốn tăng áp suất p ta có thể tăng áp lực F hoặc giảm diện tích tiếp xúc S bị ép."
    },
    {
        "num": 8,
        "question": "Chỉ ra kết luận sai trong các kết luận sau:",
        "options": [
            "A. Áp lực là lực ép có phương vuông góc với mặt bị ép",
            "B. Đơn vị của áp suất là N/m2",
            "C. Áp suất là độ lớn của áp lực trên một diện tích bị ép",
            "D. Đơn vị của áp lực là đơn vị của lực"
        ],
        "correct": "C",
        "solution": "Kết luận C sai vì áp suất là độ lớn của áp lực trên một ĐƠN VỊ diện tích bị ép (thiếu chữ 'đơn vị')."
    }
]

VATLI_EXERCISES = [
    {
        "num": 9,
        "question": "Đặt một bao gạo 25kg lên một ghế 4 chân có khối lượng 4kg. Diện tích tiếp xúc với mặt đất của mỗi chân ghế là 8cm2. Áp suất mà gạo và ghế tác dụng lên mặt đất là bao nhiêu?",
        "solution": "- Tổng khối lượng của gạo và ghế là:\nm = m_gạo + m_ghế = 25 + 4 = 29 kg.\n- Tổng áp lực tác dụng lên mặt đất:\nF = P = 10 * m = 10 * 29 = 290 N.\n- Tổng diện tích tiếp xúc của 4 chân ghế với mặt đất:\nS = 4 * 8 cm^2 = 32 cm^2 = 32 * 10^-4 m^2 = 0,0032 m^2.\n- Áp suất tác dụng lên mặt đất là:\np = F / S = 290 / 0,0032 = 90 625 N/m^2 (hoặc 90 625 Pa).\nĐáp số: 90 625 Pa."
    },
    {
        "num": 10,
        "question": "Một vật hình lập phương đặt trên mặt bàn nằm ngang, tác dụng lên mặt bàn một áp suất p = 3 600N/m2. Biết khối lượng của vật là 14,4 kg. Hỏi độ dài một cạnh của khối lập phương ấy là bao nhiêu?",
        "solution": "- Trọng lượng của vật hình lập phương (chính là áp lực vuông góc lên mặt bàn):\nF = P = 10 * m = 10 * 14,4 = 144 N.\n- Diện tích mặt đáy tiếp xúc của khối lập phương:\nS = F / p = 144 / 3600 = 0,04 m^2.\n- Vì đáy khối lập phương là hình vuông có diện tích S = a^2 (với a là độ dài một cạnh):\na = √S = √0,04 = 0,2 m = 20 cm.\nVậy độ dài một cạnh của khối lập phương là 0,2 m (hay 20 cm)."
    },
    {
        "num": 11,
        "question": "Một cái dầm sắt có thể tích là 60 dm3, biết khối lượng riêng của sắt là 7,8 g/cm3. Tính khối lượng và trọng lượng của dầm sắt này.",
        "solution": "- Đổi thể tích: V = 60 dm^3 = 60 000 cm^3 = 0,06 m^3.\n- Đổi khối lượng riêng: D = 7,8 g/cm^3 = 7800 kg/m^3.\n- Khối lượng của dầm sắt là:\nm = D * V = 7800 * 0,06 = 468 kg (hoặc 7,8 * 60000 = 468 000 g = 468 kg).\n- Trọng lượng của dầm sắt là:\nP = 10 * m = 10 * 468 = 4680 N.\nĐáp số: Khối lượng m = 468 kg; Trọng lượng P = 4680 N."
    },
    {
        "num": 12,
        "question": "Một đồng xu có khối lượng 0,9 g, được làm từ hợp kim có khối lượng riêng là 5 g/cm3. Tính thể tích của đồng xu.",
        "solution": "Áp dụng công thức: D = m / V => V = m / D.\nThể tích của đồng xu là:\nV = 0,9 / 5 = 0,18 cm^3 (tương đương 180 mm^3).\nĐáp số: 0,18 cm^3."
    },
    {
        "num": 13,
        "question": "Một khối sắt hình hộp chữ nhật có kích thước 4cm, 5cm, 7cm và khối lượng 700g. Tính trọng lượng và khối lượng riêng của khối sắt đó.",
        "solution": "- Đổi khối lượng: m = 700 g = 0,7 kg.\n- Trọng lượng của khối sắt là:\nP = 10 * m = 10 * 0,7 = 7 N.\n- Thể tích của khối sắt hình hộp chữ nhật:\nV = 4 * 5 * 7 = 140 cm^3 = 140 * 10^-6 m^3 = 0,00014 m^3.\n- Khối lượng riêng của khối sắt:\nD = m / V = 700 / 140 = 5 g/cm^3 (tương đương 5000 kg/m^3).\nĐáp số: Trọng lượng P = 7 N; Khối lượng riêng D = 5 g/cm^3 (5000 kg/m^3)."
    },
    {
        "num": 14,
        "question": "Khối lượng riêng của dầu ăn vào khoảng 800kg/m3. Do đó, 2 lít dầu ăn sẽ có trọng lượng khoảng bao nhiêu?",
        "solution": "- Đổi thể tích: V = 2 lít = 2 dm^3 = 0,002 m^3.\n- Khối lượng của 2 lít dầu ăn là:\nm = D * V = 800 * 0,002 = 1,6 kg.\n- Trọng lượng của 2 lít dầu ăn là:\nP = 10 * m = 10 * 1,6 = 16 N.\nĐáp số: 16 N."
    }
]
