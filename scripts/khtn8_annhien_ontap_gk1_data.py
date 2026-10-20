# -*- coding: utf-8 -*-
"""
Dữ liệu chuẩn HƯỚNG DẪN ÔN TẬP GIỮA HỌC KÌ 1 NĂM HỌC 2026 - 2027
MÔN: KHOA HỌC TỰ NHIÊN - LỚP 8
Trường THCS / Lớp 8/9 (An Nhiên)
Biên soạn chuẩn theo FORM ĐỀ KIỂM TRA (Thầy Trần Văn Thạnh - 0777.470.803)
Gồm trọn vẹn:
- A. PHẦN MỞ ĐẦU (Sử dụng hóa chất, thiết bị phòng thí nghiệm)
- B. PHÂN MÔN KHTN 1 (HÓA HỌC)
- C. PHÂN MÔN KHTN 2 (VẬT LÍ)
- D. PHÂN MÔN KHTN 3 (SINH HỌC)
"""

INFO = {
    "truong": "TRƯỜNG TRUNG HỌC CƠ SỞ",
    "lop": "LỚP: 8/9",
    "de_title": "HƯỚNG DẪN ÔN TẬP GIỮA HỌC KÌ I",
    "mon": "MÔN: KHOA HỌC TỰ NHIÊN 8",
    "nam_hoc": "NĂM HỌC 2026 - 2027",
    "author": "Thầy TRẦN VĂN THẠNH - 0777.470.803",
}

# ==============================================================================
# A. PHẦN MỞ ĐẦU
# ==============================================================================
MO_DAU_THEORY = [
    "*A. PHẦN MỞ ĐẦU*",
    "*1) Sử dụng một số hóa chất, thiết bị cơ bản trong phòng thí nghiệm*",
    "- *Nhận biết hoá chất:* Các hoá chất trong phòng thí nghiệm đều được đựng trong chai hoặc lọ kín, thường được làm bằng thuỷ tinh, nhựa... và có dán nhãn ghi tên, công thức hoá học, trọng lượng hoặc thể tích, độ tinh khiết, nhà sản xuất, các kí hiệu cảnh báo, điều kiện bảo quản,...",
    "- *Quy tắc sử dụng hoá chất an toàn trong phòng thí nghiệm:*",
    "  + Không sử dụng hoá chất đựng trong đồ chứa không có nhãn hoặc nhãn mờ, mất chữ.",
    "  + Trước khi sử dụng cần đọc cẩn thận nhãn hoá chất và tìm hiểu kĩ các tính chất, các lưu ý, cảnh báo của mỗi loại hoá chất để thực hiện thí nghiệm an toàn.",
    "  + Thực hiện thí nghiệm cẩn thận, không dùng tay trực tiếp lấy hoá chất. Khi lấy hoá chất rắn hay lỏng cần lựa chọn dụng cụ phù hợp (thìa thuỷ tinh, ống hút nhỏ giọt,...).",
    "  + Khi bị hoá chất dính vào người hoặc hoá chất bị đổ, tràn ra ngoài, cần báo cáo ngay với giáo viên để được hướng dẫn xử lí.",
    "  + Các hoá chất dùng xong còn thừa, tuyệt đối không được đổ trở lại bình chứa mà cần được xử lí theo đúng hướng dẫn của giáo viên.",
    "- *Giới thiệu một số dụng cụ thí nghiệm và thiết bị:* Ống nghiệm, giá để ống nghiệm, cốc thuỷ tinh, bình tam giác, ống đong, phễu lọc, đèn cồn, cân điện tử, nhiệt kế, đồng hồ bấm giây,... (theo SGK)."
]

# ==============================================================================
# B. PHÂN MÔN KHTN 1 (HÓA HỌC)
# ==============================================================================
KHTN1_THEORY = [
    "*B. PHÂN MÔN KHTN 1 (HOÁ HỌC)*",
    "*I. KIẾN THỨC TRỌNG TÂM (Bài 2, 3, 4/SGK)*",
    "- *1. Chuyển đổi giữa số mol (n) và khối lượng chất (m):*",
    "  + Công thức: *m = n × M* (gam)  hay  *n = m / M* (mol)  hay  *M = m / n* (g/mol).",
    "- *2. Chuyển đổi giữa số mol (n) và thể tích chất khí (V) ở điều kiện chuẩn (25 °C, 1 bar):*",
    "  + Công thức: *V = n × 24,79* (lít)  hay  *n = V / 24,79* (mol).",
    "- *3. Tỉ khối của chất khí:*",
    "  + Tỉ khối của khí A đối với khí B: *d~A/B~ = M~A~ / M~B~* (cho biết khí A nặng hay nhẹ hơn khí B bao nhiêu lần).",
    "  + Tỉ khối của khí A đối với không khí: *d~A/kk~ = M~A~ / 29* (với khối lượng mol trung bình không khí *M~kk~ ≈ 29 g/mol*).",
    "- *4. Độ tan của một chất trong nước (S):*",
    "  + Là số gam chất đó tan được trong 100 gam nước để tạo thành dung dịch bão hoà ở nhiệt độ xác định.",
    "  + Công thức: *S = (m~ct~ / m~H2O~) × 100* (g/100 g nước).",
    "- *5. Nồng độ phần trăm của dung dịch (C%):*",
    "  + Cho biết số gam chất tan có trong 100 gam dung dịch.",
    "  + Công thức: *C% = (m~ct~ / m~dd~) × 100%*, với *m~dd~ = m~ct~ + m~dm~*.",
    "- *6. Nồng độ mol của dung dịch (C~M~):*",
    "  + Cho biết số mol chất tan trong 1 lít dung dịch: *C~M~ = n / V* (mol/L hoặc M)."
]

KHTN1_MCQ = [
    {
        "num": 1,
        "question": "Biến đổi vật lí là gì?",
        "options": [
            "A. Khi vật thể bị biến đổi về hình dạng, trạng thái, kích thước,...mà vẫn giữ nguyên chất ban đầu.",
            "B. Khi vật thể bị thay đổi hình dạng, kích thước nhưng vẫn giữ nguyên giá trị ban đầu.",
            "C. Khi vật thể bị thay đổi chất liệu, kích thước nhưng vẫn giữ nguyên khối lượng ban đầu.",
            "D. Khi vật thể chỉ bị thay đổi về chất liệu nhưng vẫn giữ nguyên hình dạng, kích thước ban đầu."
        ],
        "correct": "A",
        "solution": "Biến đổi vật lí là hiện tượng chất bị biến đổi về hình dạng, trạng thái (thể rắn, lỏng, khí), kích thước,... nhưng vẫn giữ nguyên là chất ban đầu, không có sự tạo thành chất mới."
    },
    {
        "num": 2,
        "question": "Phản ứng sau là phản ứng gì?\nPhản ứng phân hủy copper (II) hydroxide thành copper (II) oxide và hơi nước thì cần cung cấp năng lượng dưới dạng nhiệt bằng cách đun nóng. Khi ngừng cung cấp nhiệt, phản ứng cũng dừng lại.",
        "options": [
            "A. Phản ứng tỏa nhiệt.",
            "B. Phản ứng thu nhiệt.",
            "C. Phản ứng phân hủy.",
            "D. Phản ứng trao đổi."
        ],
        "correct": "B",
        "solution": "Phản ứng cần liên tục cung cấp năng lượng nhiệt (đun nóng) trong suốt quá trình và ngừng lại khi ngắt nguồn nhiệt chính là phản ứng thu nhiệt. PTHH: Cu(OH)~2~ --t°--> CuO + H~2~O."
    },
    {
        "num": 3,
        "question": "Quá trình nào sau đây là biến đổi hoá học?",
        "options": [
            "A. Đốt cháy cồn trong đĩa.",
            "B. Hơ nóng chiếc thìa inox.",
            "C. Hoà tan muối ăn vào nước.",
            "D. Nước hoa trong lọ mở nắp bị bay hơi."
        ],
        "correct": "A",
        "solution": "- Đốt cháy cồn là phản ứng hoá học giữa rượu ethylic với khí oxygen sinh ra khí carbon dioxide và hơi nước (tạo chất mới) -> Biến đổi hoá học.\n- Hơ nóng thìa inox, hoà tan muối ăn và nước hoa bay hơi đều không sinh ra chất mới -> Biến đổi vật lí."
    },
    {
        "num": 4,
        "question": "Ở 25 °C và 1 bar, 1,5 mol khí chiếm thể tích bao nhiêu?",
        "options": [
            "A. 31,587 L.",
            "B. 35,187 L.",
            "C. 38,175 L.",
            "D. 37,185 L."
        ],
        "correct": "D",
        "solution": "Ở điều kiện chuẩn (25 °C, 1 bar), thể tích của 1,5 mol khí là:\nV = n × 24,79 = 1,5 × 24,79 = 37,185 (L)."
    },
    {
        "num": 5,
        "question": "Có thể thu khí N~2~ bằng cách nào?",
        "options": [
            "A. Đặt đứng bình.",
            "B. Đặt úp bình.",
            "C. Đặt ngang bình.",
            "D. Cách nào cũng được."
        ],
        "correct": "B",
        "solution": "Khối lượng mol của N~2~ là 28 g/mol, nhẹ hơn không khí (M~kk~ ≈ 29 g/mol). Khi thu khí nhẹ hơn không khí bằng phương pháp đẩy không khí, ta phải đặt úp ngược bình thu."
    },
    {
        "num": 6,
        "question": "Tính số nguyên tử của 1,8 mol Fe:",
        "options": [
            "A. 10,83.10^23^ nguyên tử",
            "B. 10,84.10^23^ nguyên tử",
            "C. 11.10^25^ nguyên tử",
            "D. 1,8.10^23^ nguyên tử"
        ],
        "correct": "B",
        "solution": "Số nguyên tử Fe có trong 1,8 mol Fe là:\nN = n × N~A~ = 1,8 × 6,022 × 10^23^ = 10,8396 × 10^23^ ≈ 10,84 × 10^23^ (nguyên tử) (hoặc lấy N~A~ = 6,02 × 10^23^ thì N = 10,836 × 10^23^ ≈ 10,83 × 10^23^ nguyên tử)."
    },
    {
        "num": 7,
        "question": "Số mol chứa trong 4,5 . 10^23^ phân tử NaCl là:",
        "options": [
            "A. 0,2 mol",
            "B. 0,75 mol",
            "C. 0,25 mol",
            "D. 2,25 mol"
        ],
        "correct": "B",
        "solution": "Số mol phân tử NaCl là:\nn = N / N~A~ = (4,5 × 10^23^) / (6,022 × 10^23^) ≈ 0,75 (mol) (lấy xấp xỉ N~A~ = 6 × 10^23^ thì n = 4,5 / 6 = 0,75 mol)."
    },
    {
        "num": 8,
        "question": "Khối lượng của 0,45 mol Cu(OH)~2~ là:",
        "options": [
            "A. 36,45 gam",
            "B. 15,3 gam",
            "C. 28,8 gam",
            "D. 44,1 gam"
        ],
        "correct": "D",
        "solution": "Khối lượng mol của Cu(OH)~2~: M = 64 + (16 + 1) × 2 = 98 (g/mol).\nKhối lượng của 0,45 mol Cu(OH)~2~ là:\nm = n × M = 0,45 × 98 = 44,1 (gam)."
    },
    {
        "num": 9,
        "question": "Hòa tan 15 gam sodium chloride (NaCl) vào 55 gam nước. Nồng độ phần trăm của dung dịch là",
        "options": [
            "A. 21,43%.",
            "B. 26,12%.",
            "C. 28,10%.",
            "D. 29,18%."
        ],
        "correct": "A",
        "solution": "- Khối lượng dung dịch thu được:\nm~dd~ = m~ct~ + m~H2O~ = 15 + 55 = 70 (gam).\n- Nồng độ phần trăm của dung dịch NaCl:\nC% = (m~ct~ / m~dd~) × 100% = (15 / 70) × 100% ≈ 21,43%."
    },
    {
        "num": 10,
        "question": "Tính độ tan của K~2~CO~3~ trong nước ở 20 °C. Biết rằng ở nhiệt độ này hòa tan hết 45 gam muối trong 150 gam nước thì được dung dịch bão hòa.",
        "options": [
            "A. 20 g/100.",
            "B. 30 g/100.",
            "C. 40 g/100.",
            "D. 50 g/100."
        ],
        "correct": "B",
        "solution": "Độ tan của K~2~CO~3~ ở 20 °C là:\nS = (m~ct~ / m~H2O~) × 100 = (45 / 150) × 100 = 30 (g/100 g nước)."
    }
]

KHTN1_TF = [
    {
        "num": 1,
        "intro": "Trong giờ thực hành Khoa học tự nhiên 8, học sinh tiến hành quan sát hai quá trình: (1) Đun sôi nước lỏng trong cốc thuỷ tinh thấy xuất hiện hơi nước bốc lên và ngưng tụ lại thành các giọt nước trên đĩa kính; (2) Đốt cháy một đoạn dây magnesium (Mg) trong không khí thấy Mg cháy sáng rực và tạo ra bột màu trắng magnesium oxide (MgO). Xét các phát biểu sau:",
        "items": [
            {
                "label": "a",
                "text": "Quá trình (1) là biến đổi vật lí vì nước chỉ chuyển từ thể lỏng sang thể hơi mà không tạo thành chất mới.",
                "ans": "Đúng",
                "explain": "Quá trình bay hơi và ngưng tụ chỉ là sự thay đổi trạng thái của nước, bản chất chất vẫn là H~2~O."
            },
            {
                "label": "b",
                "text": "Quá trình (2) là biến đổi hoá học vì đã có chất mới được tạo thành là magnesium oxide (MgO).",
                "ans": "Đúng",
                "explain": "Phản ứng cháy 2Mg + O~2~ --t°--> 2MgO tạo chất mới bột màu trắng là magnesium oxide."
            },
            {
                "label": "c",
                "text": "Trong quá trình (2), dây magnesium bị biến đổi màu sắc và phát sáng là những dấu hiệu nhận biết phản ứng hoá học xảy ra.",
                "ans": "Đúng",
                "explain": "Phát sáng rực rỡ và biến đổi từ dải kim loại màu trắng bạc sang bột màu trắng là các dấu hiệu rõ rệt của phản ứng hoá học."
            },
            {
                "label": "d",
                "text": "Nếu đun nóng khối lượng bột màu trắng MgO thu được ở quá trình (2) thì MgO sẽ bị biến đổi trở lại thành dây magnesium kim loại ban đầu.",
                "ans": "Sai",
                "explain": "Magnesium oxide (MgO) là hợp chất ion rất bền nhiệt, việc đun nóng thông thường không thể phân huỷ biến đổi nó trở lại thành kim loại Mg."
            }
        ]
    },
    {
        "num": 2,
        "intro": "Ở điều kiện chuẩn (25 °C và 1 bar), thể tích mol của bất kì chất khí nào cũng bằng 24,79 lít. Xét các phát biểu sau đối với các chất khí đo ở điều kiện chuẩn:",
        "items": [
            {
                "label": "a",
                "text": "Ở điều kiện chuẩn, 1 mol khí oxygen (O~2~) và 1 mol khí hydrogen (H~2~) đều chiếm thể tích bằng nhau là 24,79 lít.",
                "ans": "Đúng",
                "explain": "Theo định luật Avogadro, ở cùng điều kiện chuẩn thì 1 mol của mọi chất khí đều chiếm thể tích bằng nhau là 24,79 L."
            },
            {
                "label": "b",
                "text": "Thể tích của 0,2 mol khí carbon dioxide (CO~2~) ở điều kiện chuẩn được tính theo công thức V = n × 24,79 = 4,958 lít.",
                "ans": "Đúng",
                "explain": "Áp dụng công thức chuẩn: V = 0,2 × 24,79 = 4,958 (lít)."
            },
            {
                "label": "c",
                "text": "Một hỗn hợp khí gồm 0,1 mol N~2~ và 0,1 mol O~2~ ở điều kiện chuẩn có tổng thể tích là 4,958 lít.",
                "ans": "Đúng",
                "explain": "Tổng số mol khí n = 0,1 + 0,1 = 0,2 mol. Thể tích hỗn hợp: V = 0,2 × 24,79 = 4,958 (lít)."
            },
            {
                "label": "d",
                "text": "Các chất khí khác nhau có khối lượng mol khác nhau thì ở cùng điều kiện nhiệt độ và áp suất chuẩn bắt buộc phải có thể tích khác nhau.",
                "ans": "Sai",
                "explain": "Thể tích chất khí chỉ phụ thuộc vào số mol phân tử khí, không phụ thuộc vào khối lượng mol M của chất khí đó."
            }
        ]
    }
]

KHTN1_SHORT = [
    {
        "num": 1,
        "question": "Tính khối lượng (tính bằng gam) của 0,4 mol phân tử khí carbon dioxide (CO~2~).",
        "ans": "17,6 gam",
        "solution": "Khối lượng của 0,4 mol CO~2~ là: m = n × M = 0,4 × 44 = 17,6 (gam)."
    },
    {
        "num": 2,
        "question": "Tính thể tích (tính bằng lít) ở điều kiện chuẩn (25 °C, 1 bar) của 0,3 mol khí nitrogen (N~2~).",
        "ans": "7,437 lít",
        "solution": "Thể tích ở điều kiện chuẩn là: V = n × 24,79 = 0,3 × 24,79 = 7,437 (lít)."
    },
    {
        "num": 3,
        "question": "Một bình kín chứa 12,395 lít khí oxygen (O~2~) ở điều kiện chuẩn. Tính số mol khí O~2~ có trong bình.",
        "ans": "0,5 mol",
        "solution": "Số mol khí O~2~ là: n = V / 24,79 = 12,395 / 24,79 = 0,5 (mol)."
    },
    {
        "num": 4,
        "question": "Tính thể tích ở điều kiện chuẩn của hỗn hợp khí gồm 0,15 mol CO~2~ và 0,25 mol O~2~.",
        "ans": "9,916 lít",
        "solution": "Tổng số mol của hỗn hợp khí: n = 0,15 + 0,25 = 0,4 (mol).\nThể tích hỗn hợp khí ở đkc: V = n × 24,79 = 0,4 × 24,79 = 9,916 (lít)."
    }
]

KHTN1_ESSAY = [
    {
        "num": 1,
        "question": "Hãy tính:\na. Khối lượng của hỗn hợp gồm 0,1 mol NaCl và 0,2 mol CuSO~4~\nb. Thể tích ở 25 °C, 1 bar của 1,5 mol khí CH~4~.\nc. Thể tích ở 25 °C, 1 bar của hỗn hợp gồm 0,05 mol CO và 0,15 mol CO~2~",
        "solution": "a. Khối lượng của hỗn hợp:\n   m~NaCl~ = 0,1 × 58,5 = 5,85 (gam)\n   m~CuSO4~ = 0,2 × 160 = 32 (gam)\n   => m~hh~ = 5,85 + 32 = 37,85 (gam).\nb. Thể tích của 1,5 mol khí CH~4~ ở đkc:\n   V = n × 24,79 = 1,5 × 24,79 = 37,185 (lít).\nc. Thể tích của hỗn hợp khí ở đkc:\n   Tổng số mol: n~hh~ = 0,05 + 0,15 = 0,2 (mol)\n   => V~hh~ = 0,2 × 24,79 = 4,958 (lít)."
    },
    {
        "num": 2,
        "question": "Tính số mol nguyên tử hoặc mol phân tử trong những lượng chất sau:\na. 8,428.10^22^ nguyên tử K.\nb. 1,505.10^24^ phân tử SO~2~.",
        "solution": "a. Số mol nguyên tử K:\n   n~K~ = N / N~A~ = (8,428 × 10^22^) / (6,022 × 10^23^) = 0,14 (mol).\nb. Số mol phân tử SO~2~:\n   n~SO2~ = N / N~A~ = (1,5055 × 10^24^) / (6,022 × 10^23^) = 2,5 (mol)."
    },
    {
        "num": 3,
        "question": "Hãy trình bày cách pha chế 250 gam dung dịch NaCl nồng độ 20%.",
        "solution": "* Bước 1: Tính toán:\n- Khối lượng chất tan NaCl cần lấy:\n  m~NaCl~ = (m~dd~ × C%) / 100% = (250 × 20%) / 100% = 50 (gam).\n- Khối lượng nước cần lấy:\n  m~H2O~ = m~dd~ - m~NaCl~ = 250 - 50 = 200 (gam) (tương đương 200 mL nước).\n\n* Bước 2: Cách pha chế:\n- Cân lấy 50 gam muối ăn NaCl khan cho vào cốc thuỷ tinh dung tích 500 mL.\n- Đong 200 mL (hoặc cân 200 gam) nước cất rồi rót từ từ vào cốc chứa muối NaCl.\n- Dùng đũa thuỷ tinh khuấy đều cho đến khi muối NaCl tan hoàn toàn. Ta thu được đúng 250 gam dung dịch NaCl 20%."
    },
    {
        "num": 4,
        "question": "Hòa tan 20 gam KCl vào 60 gam nước. Tính nồng độ phần trăm của dung dịch?",
        "solution": "- Khối lượng dung dịch thu được là:\nm~dd~ = m~ct~ + m~dm~ = 20 + 60 = 80 (gam).\n- Nồng độ phần trăm của dung dịch KCl là:\nC% = (m~ct~ / m~dd~) × 100% = (20 / 80) × 100% = 25%."
    },
    {
        "num": 5,
        "question": "Ở nhiệt độ 20 °C, khi hoà tan 15 gam muối X vào 30 gam nước, khuấy kĩ thì còn lại 3 gam muối không tan. Tính độ tan của muối X.",
        "solution": "- Khối lượng muối X đã hoà tan tối đa vào 30 gam nước là:\nm~X tan~ = 15 - 3 = 12 (gam).\n- Độ tan của muối X ở 20 °C là:\nS = (m~X tan~ / m~H2O~) × 100 = (12 / 30) × 100 = 40 (g/100 g nước)."
    },
    {
        "num": 6,
        "question": "Hoà tan hoàn toàn 14,6 gam Na~2~CO~3~ vào nước, thu được 200mL dung dịch. Tính nồng độ mol của dung dịch Na~2~CO~3~ thu được.",
        "solution": "- Đổi thể tích: V = 200 mL = 0,2 lít.\n- Khối lượng mol của Na~2~CO~3~: M = 23 × 2 + 12 + 16 × 3 = 106 (g/mol).\n- Số mol Na~2~CO~3~ là:\nn = m / M = 14,6 / 106 ≈ 0,1377 (mol).\n- Nồng độ mol của dung dịch Na~2~CO~3~ thu được:\nC~M~ = n / V = 0,1377 / 0,2 ≈ 0,69 (M) (hoặc mol/L).\n*(Ghi chú: Nếu đề gốc in là 10,6 gam thì n = 0,1 mol => C~M~ = 0,5 M; nếu là 21,2 gam thì n = 0,2 mol => C~M~ = 1,0 M)*."
    }
]

# ==============================================================================
# C. PHÂN MÔN KHTN 2 (VẬT LÍ)
# ==============================================================================
KHTN2_THEORY = [
    "*C. PHÂN MÔN KHTN 2 (VẬT LÍ)*",
    "*I) KIẾN THỨC TRỌNG TÂM*",
    "- *1. Khối lượng riêng:*",
    "  + *Định nghĩa:* Khối lượng riêng của một chất là khối lượng của một đơn vị thể tích (1 m^3^) chất đó.",
    "  + *Công thức:* *D = m / V* (trong đó: *D* là khối lượng riêng, đơn vị kg/m^3^; *m* là khối lượng của vật, đơn vị kg; *V* là thể tích của vật, đơn vị m^3^).",
    "- *2. Cách xác định khối lượng riêng:*",
    "  + *Vật hình hộp chữ nhật:* Dùng thước đo chiều dài, chiều rộng, chiều cao => tính *V = a × b × c*; dùng cân đo khối lượng *m* => tính *D = m / V*.",
    "  + *Lượng chất lỏng:* Cân khối lượng ống đong rỗng (*m~1~*); đổ chất lỏng vào đọc thể tích *V* và cân khối lượng (*m~2~*) => khối lượng chất lỏng *m = m~2~ - m~1~* => tính *D = m / V*.",
    "  + *Vật rắn hình dạng bất kì không thấm nước (viên sỏi):* Dùng cân đo khối lượng *m*; dùng bình chia độ đo thể tích nước dâng lên *V = V~2~ - V~1~* => tính *D = m / V*.",
    "- *3. Áp lực và Áp suất:*",
    "  + *Áp lực:* Là lực ép có phương vuông góc với mặt bị ép.",
    "  + *Áp suất:* Là độ lớn của áp lực trên một đơn vị diện tích bị ép.",
    "  + *Công thức:* *p = F / S* (trong đó: *p* là áp suất, đơn vị N/m^2^ hoặc Pascal - Pa; *F* là áp lực tác dụng vuông góc lên mặt bị ép, đơn vị Newton - N; *S* là diện tích bề mặt bị ép, đơn vị m^2^).",
    "- *4. Cách làm tăng và giảm áp suất:*",
    "  + *Tăng áp suất:* Tăng áp lực *F* và/hoặc giảm diện tích bị ép *S* (ví dụ: mài sắc dao, kéo, vót nhọn cọc,...).",
    "  + *Giảm áp suất:* Giảm áp lực *F* và/hoặc tăng diện tích bị ép *S* (ví dụ: làm móng nhà rộng bản, đặt tà vẹt dưới đường ray xe lửa, kê hòn đá tảng to dưới chân cột đình,...)."
]

KHTN2_MCQ = [
    {
        "num": 1,
        "question": "Đơn vị đo khối lượng riêng là",
        "options": [
            "A. N/m^3^.",
            "B. V/m^3^.",
            "C. kg/m^3^.",
            "D. kg/m^2^."
        ],
        "correct": "C",
        "solution": "Trong hệ đơn vị đo lường quốc tế SI, đơn vị của khối lượng riêng là kg/m^3^ (kilôgam trên mét khối)."
    },
    {
        "num": 2,
        "question": "Áp lực là:",
        "options": [
            "A. Lực ép có phương vuông góc với mặt bị ép.",
            "B. Lực ép có phương song song với mặt bị ép.",
            "C. Lực ép có phương tạo với mặt bị ép một góc bất kì.",
            "D. Lực ép có phương trùng với mặt bị ép."
        ],
        "correct": "A",
        "solution": "Theo định nghĩa SGK KHTN 8: Áp lực là lực ép có phương vuông góc với mặt bị ép."
    },
    {
        "num": 3,
        "question": "Trong các cách tăng, giảm áp suất sau đây, cách nào là không đúng ?",
        "options": [
            "A. Muốn tăng áp suất thì tăng áp lực, giảm diện tích bị ép.",
            "B. Muốn tăng áp suất thì giảm áp lực, tăng diện tích bị ép.",
            "C. Muốn giảm áp suất thì phải giảm áp lực, giữ nguyên diện tích bị ép.",
            "D. Muốn giảm áp suất thì phải giữ nguyên áp lực, tăng diện tích bị ép."
        ],
        "correct": "B",
        "solution": "Từ công thức p = F / S, khi giảm áp lực F và tăng diện tích bị ép S thì áp suất p sẽ giảm đi, chứ không thể tăng lên được. Vậy cách B là không đúng."
    },
    {
        "num": 4,
        "question": "Em hãy sắp xếp lại thứ tự các bước cho dưới đây để xác định được khối lượng riêng của một lượng nước?\n(1) Xác định khối lượng của ống đong.\n(2) Xác định khối lượng nước trong ống đong.\n(3) Xác định khối lượng của ống đong có đựng nước.\n(4) Sử dụng công thức tính khối lượng riêng.",
        "options": [
            "A. (1) – (2) – (3) – (4).",
            "B. (1) – (2) – (4) – (3).",
            "C. (1) – (3) – (2) – (4).",
            "D. (2) – (1) – (3) – (4)."
        ],
        "correct": "C",
        "solution": "Thứ tự tiến hành đúng: (1) Cân khối lượng ống đong rỗng (m~1~) -> (3) Đổ nước vào và cân khối lượng cả ống đong lẫn nước (m~2~) -> (2) Xác định khối lượng nước: m = m~2~ - m~1~ -> (4) Đọc thể tích V và áp dụng công thức D = m / V. Vậy thứ tự chuẩn là (1) – (3) – (2) – (4)."
    },
    {
        "num": 5,
        "question": "Chọn câu đúng.",
        "options": [
            "A. Lưỡi dao, lưỡi kéo phải mài thật sắc để tăng diện tích bề mặt nên khi cắt, thái,... được dễ dàng.",
            "B. Những cột đình làng thường kê trên những hòn đá rộng và phẳng để làm giảm áp suất gây ra lên mặt đất.",
            "C. Đường ray phải được đặt trên những thanh tà vẹt để làm tăng áp lực lên mặt đất khi tàu hỏa chạy qua.",
            "D. Đặt ván lên bùn (đất) ít bị lún hơn khi đi bằng chân không vì sẽ làm giảm áp lực của cơ thể lên bùn đất."
        ],
        "correct": "B",
        "solution": "Hòn đá rộng và phẳng làm tăng diện tích tiếp xúc S, theo công thức p = F / S sẽ làm giảm áp suất p do cột đình nén lên mặt đất, giúp chống lún sụt. (A sai vì mài sắc làm giảm diện tích; C và D sai vì thanh tà vẹt hay ván lót làm giảm áp suất nhờ tăng diện tích S, áp lực F bằng trọng lượng không đổi)."
    }
]

KHTN2_SHORT = [
    {
        "num": 1,
        "question": "Một nhóm học sinh tiến hành thí nghiệm xác định khối lượng riêng của một vật hình hộp chữ nhật. Các bước được mô tả như sau:\na) Tính thể tích của vật từ các kích thước đã đo.\nb) Dùng cân xác định khối lượng của vật.\nc) Tính khối lượng riêng của vật từ khối lượng và thể tích.\nd) Dùng thước đo chiều dài, chiều rộng và chiều cao của vật.\nThứ tự thực hiện đúng là gì?",
        "ans": "d – a – b – c (hoặc b – d – a – c)",
        "solution": "- Bước 1 (d): Dùng thước đo chiều dài, chiều rộng và chiều cao của vật.\n- Bước 2 (a): Tính thể tích của vật từ các kích thước đã đo (V = a × b × c).\n- Bước 3 (b): Dùng cân xác định khối lượng m của vật.\n- Bước 4 (c): Tính khối lượng riêng của vật theo công thức D = m / V."
    }
]

KHTN2_ESSAY = [
    {
        "num": 1,
        "question": "a) Định nghĩa khối lượng riêng? Viết công thức tính khối lượng riêng.\nb) Trình bày các bước tiến hành thí nghiệm để xác định được khối lượng riêng của một vật hình hộp chữ nhật, một lượng nước, sỏi.\nc) Một bể nước có kích thước bên trong là 80 cm x 20 cm x 25 cm. Cho biết khối lượng riêng của nước là 1,0 g/cm^3^. Tính khối lượng nước trong bể khi bể chứa đầy nước.",
        "solution": "a) Định nghĩa: Khối lượng riêng của một chất là khối lượng của một đơn vị thể tích (1 m^3^) chất đó.\n   Công thức: D = m / V (trong đó m là khối lượng, V là thể tích).\nb) Các bước tiến hành thí nghiệm:\n   * Vật hình hộp chữ nhật: Dùng thước đo 3 kích thước dài (a), rộng (b), cao (c) => tính V = a × b × c; dùng cân đo khối lượng m => tính D = m / V.\n   * Lượng nước: Cân ống đong rỗng (m~1~); rót nước vào đọc thể tích V và cân ống đong có nước (m~2~) => khối lượng nước m = m~2~ - m~1~ => tính D = m / V.\n   * Viên sỏi (vật rắn không thấm nước): Cân khối lượng viên sỏi (m); đổ nước vào bình chia độ tới mức V~1~, thả chìm sỏi nước dâng lên mức V~2~ => thể tích sỏi V = V~2~ - V~1~ => tính D = m / V.\nc) Tính khối lượng nước trong bể:\n   Thể tích của bể nước: V = 80 × 20 × 25 = 40 000 (cm^3^).\n   Khối lượng nước khi đầy bể: m = D × V = 1,0 × 40 000 = 40 000 (g) = 40 (kg)."
    },
    {
        "num": 2,
        "question": "a) Một hộp sữa Ông Thọ có ghi 397g. Biết dung tích của hộp sữa là 320 cm^3^. Hãy tính khối lượng riêng của sữa trong hộp theo đơn vị kg/m^3^.\nb) Một khối đá có thể tích 500 dm^3^ và khối lượng riêng là 2580 kg/m^3^. Khối lượng của khối đá là bao nhiêu?",
        "solution": "a) Đổi đơn vị: m = 397 g = 0,397 kg; V = 320 cm^3^ = 0,00032 m^3^.\n   Khối lượng riêng của sữa là:\n   D = m / V = 0,397 / 0,00032 = 1240,625 (kg/m^3^) ≈ 1240,6 kg/m^3^.\nb) Đổi thể tích: V = 500 dm^3^ = 0,5 m^3^.\n   Khối lượng của khối đá là:\n   m = D × V = 2580 × 0,5 = 1290 (kg)."
    },
    {
        "num": 3,
        "question": "a) Áp suất là gì? Viết công thức tính áp suất.\nb) Một người làm vườn cần đóng một chiếc cọc xuống đất. Hãy đề xuất phương án đóng chiếc cọc xuống một cách dễ dàng, giải thích cách làm đó.\nc) Hãy giải thích vì sao khi nằm trên nệm ta thấy êm ái hơn khi nằm trên phản gỗ.",
        "solution": "a) Định nghĩa: Áp suất là độ lớn của áp lực trên một đơn vị diện tích bị ép.\n   Công thức: p = F / S (trong đó F là áp lực tác dụng vuông góc lên mặt bị ép, đơn vị N; S là diện tích bị ép, đơn vị m^2^; p là áp suất, đơn vị N/m^2^ hoặc Pa).\nb) Phương án đóng cọc:\n   - Phương án: Vót nhọn một đầu chiếc cọc cắm xuống đất, dùng búa gõ vào đầu bằng phía trên.\n   - Giải thích: Vót nhọn đầu cọc làm giảm diện tích tiếp xúc S với mặt đất. Theo công thức p = F / S, với cùng một lực đóng F, diện tích S càng nhỏ thì áp suất p tạo ra càng lớn, giúp đầu cọc dễ dàng lún sâu vào lòng đất.\nc) Giải thích nằm nệm êm ái hơn phản gỗ:\n   - Nệm có tính đàn hồi và mềm, khi người nằm lên nệm bị lún xuống ôm lấy cơ thể, làm diện tích tiếp xúc S giữa cơ thể và nệm tăng lên rất lớn.\n   - Trọng lượng cơ thể (áp lực F = P) không đổi. Do p = F / S, diện tích tiếp xúc S lớn làm áp suất p tác dụng lên cơ thể giảm đi rất nhiều, giúp cơ thể không bị chèn ép đau nhức nên cảm thấy rất êm ái."
    }
]

# ==============================================================================
# D. PHÂN MÔN KHTN 3 (SINH HỌC)
# ==============================================================================
KHTN3_THEORY = [
    "*D. PHÂN MÔN KHTN 3 (SINH HỌC)*",
    "*I. KIẾN THỨC TRỌNG TÂM*",
    "- *1. Khái quát về cơ thể người:* Gồm nhiều cơ quan và hệ cơ quan (vận động, tiêu hoá, tuần hoàn, hô hấp, bài tiết, thần kinh, sinh dục, nội tiết, da). Các hệ cơ quan phối hợp nhịp nhàng tạo thành một khối thống nhất.",
    "- *2. Hệ vận động ở người:*",
    "  + Cấu tạo: Bộ xương (xương đầu, xương thân, xương chi, các khớp xương) và hệ cơ (cơ vân).",
    "  + Chức năng: Định hình cơ thể, nâng đỡ, bảo vệ nội quan, giúp cơ thể vận động và di chuyển.",
    "  + Bệnh, tật: Cong vẹo cột sống, loãng xương. Biện pháp: Ăn uống đủ chất, bổ sung vitamin D và calcium; tập thể dục đúng cách; ngồi học ngay ngắn, không mang vác lệch bên.",
    "- *3. Dinh dưỡng và tiêu hoá ở người:*",
    "  + Ống tiêu hoá: Miệng, hầu, thực quản, dạ dày, ruột non, ruột già, trực tràng, hậu môn.",
    "  + Tuyến tiêu hoá: Tuyến nước bọt, gan, tuyến tuỵ, tuyến vị, tuyến ruột.",
    "  + Chức năng: Biến đổi thức ăn thành chất dinh dưỡng đơn giản hấp thụ vào máu và đào thải chất cặn bã.",
    "  + Vệ sinh an toàn thực phẩm: Giữ thực phẩm không bị nhiễm khuẩn, nhiễm độc, biến chất; phòng chống ngộ độc thực phẩm.",
    "- *4. Máu và hệ tuần hoàn:*",
    "  + Thành phần máu: Huyết tương (duy trì trạng thái lỏng, vận chuyển dinh dưỡng, chất thải) và tế bào máu (hồng cầu vận chuyển O~2~ và CO~2~; bạch cầu bảo vệ cơ thể; tiểu cầu làm đông máu).",
    "  + Hệ nhóm máu ABO: 4 nhóm máu (A, B, AB, O). Nhóm O là nhóm máu chuyên cho; nhóm AB là nhóm máu chuyên nhận.",
    "  + Nguyên tắc truyền máu: Không để kháng thể trong huyết tương người nhận làm ngưng kết kháng nguyên trên hồng cầu người cho; xét nghiệm không truyền máu nhiễm bệnh.",
    "- *5. Hệ hô hấp ở người:*",
    "  + Cấu tạo: Đường dẫn khí (mũi, họng, thanh quản, khí quản, phế quản) và cơ quan trao đổi khí (hai lá phổi chứa hàng triệu phế nang).",
    "  + Chức năng: Cung cấp O~2~ cho tế bào cơ thể và đào thải khí CO~2~ ra ngoài qua cử động hô hấp và cơ chế khuếch tán."
]

KHTN3_MCQ = [
    {
        "num": 1,
        "question": "Lọc từ máu những chất thừa có hại cho cơ thể và đưa ra ngoài là chức năng của hệ",
        "options": [
            "A. Hệ thần kinh",
            "B. Hệ tiêu hoá",
            "C. Hệ vận động",
            "D. Hệ bài tiết"
        ],
        "correct": "D",
        "solution": "Hệ bài tiết (thận, ống dẫn nước tiểu, bóng đái, ống đái và da) có chức năng lọc từ máu các chất thải độc hại và chất dư thừa để bài tiết ra khỏi cơ thể."
    },
    {
        "num": 2,
        "question": "Cơ quan nào sau đây thuộc hệ thần kinh ?",
        "options": [
            "A. Tim",
            "B. Dạ dày",
            "C. Tủy sống",
            "D. Khí quản"
        ],
        "correct": "C",
        "solution": "Hệ thần kinh gồm bộ phận trung ương (não bộ và tuỷ sống) và bộ phận ngoại biên (các dây thần kinh, hạch thần kinh)."
    },
    {
        "num": 3,
        "question": "Các cơ quan như dạ dày, ruột non, ruột già thuộc hệ cơ quan nào dưới đây?",
        "options": [
            "A. Hệ hô hấp.",
            "B. Hệ tiêu hoá.",
            "C. Hệ bài tiết.",
            "D. Hệ sinh dục."
        ],
        "correct": "B",
        "solution": "Dạ dày, ruột non, ruột già là các đoạn quan trọng của ống tiêu hoá thuộc hệ tiêu hoá."
    },
    {
        "num": 4,
        "question": "Hệ tuần hoàn ở người được cấu tạo gồm các cơ quan nào sau đây?",
        "options": [
            "A. Tim và hệ thống mạch máu",
            "B. Phổi và đường dẫn khí",
            "C. Thận, da và phổi",
            "D. Ống tiêu hóa và các tuyến tiêu hóa"
        ],
        "correct": "A",
        "solution": "Hệ tuần hoàn ở người được cấu tạo gồm tim (co bóp đẩy máu) và hệ thống mạch máu (động mạch, mao mạch, tĩnh mạch)."
    },
    {
        "num": 5,
        "question": "Hệ cơ quan nào trong cơ thể có vai trò chính là thu nhận các kích thích từ môi trường, điều khiển, điều hòa hoạt động của các cơ quan giúp cơ thể thích nghi với môi trường?",
        "options": [
            "A. Hệ vận động",
            "B. Hệ tiêu hóa",
            "C. Hệ thần kinh",
            "D. Hệ bài tiết"
        ],
        "correct": "C",
        "solution": "Hệ thần kinh tiếp nhận kích thích từ môi trường trong và ngoài cơ thể, điều khiển, điều hoà và phối hợp hoạt động của các cơ quan giúp cơ thể thích nghi với môi trường sống."
    },
    {
        "num": 6,
        "question": "Gia đình Nam sống gần nơi có nhiều xe cộ qua lại, thường xuyên có bụi và khói. Để bảo vệ sức khỏe hô hấp của các thành viên trong gia đình, việc làm nào sau đây là phù hợp nhất?",
        "options": [
            "A. Thường xuyên mở cửa để đón không khí từ ngoài đường vào nhà.",
            "B. Đeo khẩu trang khi ra đường, vệ sinh nhà cửa thường xuyên và hạn chế tiếp xúc với khói, bụi.",
            "C. Chỉ cần uống nhiều nước mỗi ngày để tránh các bệnh về đường hô hấp.",
            "D. Tập thể dục ngoài đường vào giờ có nhiều xe cộ để tăng sức khỏe."
        ],
        "correct": "B",
        "solution": "Đeo khẩu trang giúp cản bụi bẩn và vi khuẩn, vệ sinh nhà cửa thường xuyên và hạn chế khói bụi là biện pháp bảo vệ đường hô hấp thiết thực và hiệu quả nhất."
    },
    {
        "num": 7,
        "question": "Trong hệ hô hấp, cơ quan nào đóng vai trò chính là nơi diễn ra quá trình trao đổi khí giữa cơ thể và môi trường?",
        "options": [
            "A. Mũi",
            "B. Thanh quản",
            "C. Hai lá phổi",
            "D. Khí quản"
        ],
        "correct": "C",
        "solution": "Hai lá phổi (với mạng lưới hàng trăm triệu phế nang và mao mạch máu bao quanh) là nơi diễn ra trực tiếp quá trình trao đổi khí O~2~ và CO~2~ giữa máu và môi trường."
    },
    {
        "num": 8,
        "question": "Cơ quan nào sau đây thuộc đường dẫn khí của hệ hô hấp ở người?",
        "options": [
            "A. Hai lá phổi",
            "B. Khí quản",
            "C. Dạ dày",
            "D. Tim"
        ],
        "correct": "B",
        "solution": "Đường dẫn khí của hệ hô hấp gồm: mũi, họng, thanh quản, khí quản và phế quản. Do đó khí quản thuộc đường dẫn khí."
    }
]

KHTN3_TF = [
    {
        "num": 1,
        "intro": "Xét các nhận định về máu, hệ tuần hoàn và truyền máu ở người:",
        "items": [
            {
                "label": "a",
                "text": "Máu là phần dịch lỏng trong cơ thể, gồm hai phần chính là huyết tương và các tế bào máu (hồng cầu, bạch cầu, tiểu cầu).",
                "ans": "Đúng",
                "explain": "Máu gồm huyết tương (chiếm khoảng 55%) và các tế bào máu (khoảng 45%)."
            },
            {
                "label": "b",
                "text": "Bạch cầu là thành phần của máu có chức năng chính là vận chuyển oxygen và carbon dioxide đi nuôi cơ thể.",
                "ans": "Sai",
                "explain": "Chức năng vận chuyển O~2~ và CO~2~ là của hồng cầu (nhờ huyết sắc tố hemoglobin); chức năng của bạch cầu là bảo vệ cơ thể chống tác nhân gây bệnh."
            },
            {
                "label": "c",
                "text": "Huyết tương giúp duy trì máu ở trạng thái lỏng để lưu thông dễ dàng trong mạch, đồng thời tiểu cầu tham gia bảo vệ cơ thể nhờ cơ chế làm đông máu.",
                "ans": "Đúng",
                "explain": "Huyết tương duy trì thể tích và tính lỏng của máu; tiểu cầu giải phóng enzyme đông máu tạo cục máu đông bịt kín vết thương."
            },
            {
                "label": "d",
                "text": "Người có nhóm máu AB có thể truyền máu an toàn cho người mang bất kỳ nhóm máu nào khác (A, B, O) vì hồng cầu của nhóm AB chứa cả hai loại kháng nguyên A và B.",
                "ans": "Sai",
                "explain": "Nhóm máu AB chỉ truyền được cho chính nhóm máu AB. Nếu truyền cho nhóm A, B hoặc O thì kháng nguyên A và B trên hồng cầu người cho sẽ bị kháng thể trong máu người nhận gây kết dính hồng cầu dẫn đến tử vong."
            },
            {
                "label": "e",
                "text": "Hiến máu nhân đạo theo đúng hướng dẫn y tế vừa giúp cứu sống người bệnh trong các tình huống cấp cứu, vừa không gây hại cho sức khỏe của người hiến vì cơ thể có khả năng tái tạo lại lượng máu đã mất.",
                "ans": "Đúng",
                "explain": "Tuỷ xương có khả năng sinh sản tế bào máu mới nhanh chóng để bù lại lượng máu đã hiến, đồng thời kích thích cơ thể sản sinh máu mới khỏe mạnh."
            }
        ]
    }
]

KHTN3_ESSAY = [
    {
        "num": 1,
        "question": "Để cải thiện tình trạng hay bị đầy bụng, đau dạ dày do thói quen ăn nhanh và ăn đồ chua cay, em hãy đề xuất 2 biện pháp cụ thể trong ăn uống.",
        "solution": "- Biện pháp 1: Ăn chậm, nhai thật kĩ trước khi nuốt để thức ăn được nghiền nhỏ và thấm đều enzyme trong nước bọt, giúp giảm bớt áp lực co bóp và tiết acid của dạ dày.\n- Biện pháp 2: Hạn chế tối đa các món ăn quá chua, cay nóng (ớt, tiêu, giấm, xoài chua,...), đồ ăn nhiều dầu mỡ chiên rán; ăn đúng giờ giấc, không bỏ bữa và không ăn quá no trước khi đi ngủ."
    },
    {
        "num": 2,
        "question": "Em hãy đề xuất 2 biện pháp trong sinh hoạt hằng ngày để phòng tránh bệnh sâu răng.",
        "solution": "- Biện pháp 1: Đánh răng đúng cách ít nhất 2 lần mỗi ngày (sau các bữa ăn và trước khi đi ngủ) bằng bàn chải lông mềm và kem đánh răng có chứa fluoride; kết hợp dùng chỉ nha khoa và súc miệng bằng nước muối sinh lý.\n- Biện pháp 2: Hạn chế ăn vặt đồ ngọt, bánh kẹo, uống nước ngọt có gas (đặc biệt là vào buổi tối); định kì 6 tháng/lần đi khám răng để phát hiện và điều trị kịp thời các tổn thương men răng."
    },
    {
        "num": 3,
        "question": "Kể tên 2 thói quen xấu trong sinh hoạt hằng ngày có thể làm gia tăng nguy cơ mắc tật cong vẹo cột sống ở lứa tuổi học sinh.",
        "solution": "- Thói quen 1: Ngồi học sai tư thế (cúi gập đầu quá sát mặt bàn, nằm bò ra bàn khi viết, ngồi nghiêng vẹo sang một bên hoặc ngồi bàn ghế không phù hợp với chiều cao lứa tuổi).\n- Thói quen 2: Đeo cặp sách quá nặng so với thể trạng hoặc thường xuyên đeo cặp lệch sang một bên vai khi đi học."
    },
    {
        "num": 4,
        "question": "Đề xuất các biện pháp lựa chọn, bảo quản, chế biến, chế độ ăn uống an toàn cho bản thân và gia đình.",
        "solution": "- Lựa chọn thực phẩm: Chọn thực phẩm tươi sống, nguồn gốc xuất xứ rõ ràng, có dấu kiểm định thú y; thực phẩm đóng gói phải còn nguyên hạn sử dụng, bao bì không bị rách, phồng hay móp méo.\n- Bảo quản thực phẩm: Bảo quản ở nhiệt độ thích hợp (ngăn đông/ngăn mát tủ lạnh); để riêng thực phẩm sống và thực phẩm chín; che đậy cẩn thận tránh ruồi nhặng, bụi bẩn.\n- Chế biến thực phẩm: Rửa sạch thực phẩm dưới vòi nước chảy; dùng riêng dao thớt cho đồ sống và đồ chín; nấu chín kĩ (thực hiện ăn chín, uống sôi).\n- Chế độ ăn uống: Ăn uống đủ chất, cân đối 4 nhóm dinh dưỡng; ăn đúng giờ, uống đủ 1,5 - 2 lít nước mỗi ngày; tránh thức ăn ôi thiu, biến chất hoặc nghi ngờ nhiễm độc."
    }
]
