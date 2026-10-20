# -*- coding: utf-8 -*-
"""
Dữ liệu chuẩn HƯỚNG DẪN ÔN TẬP GIỮA KÌ I - MÔN: KHOA HỌC TỰ NHIÊN 8
Trường THCS Kim Long - UBND Phường Kim Long
Biên soạn chuẩn theo FORM ĐỀ KIỂM TRA (Thầy Trần Văn Thạnh - 0777.470.803)
Gồm trọn vẹn 3 phân môn:
- PHÂN MÔN KHTN 1: Hóa học (Trang 1 - 3)
- PHÂN MÔN KHTN 2: Vật lí (Trang 4 - 6)
- PHÂN MÔN KHTN 3: Sinh học (Trang 7 - 8)
"""

INFO = {
    "ubnd": "UBND PHƯỜNG KIM LONG",
    "truong": "TRƯỜNG THCS KIM LONG",
    "de_title": "HƯỚNG DẪN ÔN TẬP GIỮA KÌ I",
    "mon": "MÔN: KHOA HỌC TỰ NHIÊN 8",
    "nam_hoc": "NĂM HỌC 2026 - 2027",
    "author": "Thầy TRẦN VĂN THẠNH - 0777.470.803",
}

# ==============================================================================
# PHÂN MÔN KHTN 1 (HÓA HỌC)
# ==============================================================================

KHTN1_THEORY = [
    "*A/ LÝ THUYẾT: Kiến thức từ bài 2 đến bài 5*",
    "*Bài 2. Phản ứng hoá học*",
    "- *Khái niệm:* Phản ứng hoá học là quá trình biến đổi từ chất này thành chất khác. Chất ban đầu gọi là chất phản ứng (chất tham gia), chất mới tạo thành gọi là sản phẩm.",
    "- *Diễn biến:* Trong phản ứng hoá học, chỉ có liên kết giữa các nguyên tử thay đổi làm cho phân tử này biến đổi thành phân tử khác. Các nguyên tử được bảo toàn (không biến đổi).",
    "- *Dấu hiệu nhận biết:* Có sự thay đổi màu sắc, xuất hiện chất kết tủa, có chất khí thoát ra, phát sáng hoặc tỏa nhiệt / thu nhiệt.",
    "- *Năng lượng trong phản ứng:*",
    "  + *Phản ứng tỏa nhiệt:* Giải phóng năng lượng nhiệt ra môi trường xung quanh (nhiệt độ môi trường tăng lên).",
    "  + *Phản ứng thu nhiệt:* Hấp thụ năng lượng nhiệt từ môi trường xung quanh (nhiệt độ môi trường hạ xuống).",
    "*Bài 3. Mol và tỉ khối chất khí*",
    "- *Mol:* Là lượng chất chứa *N* (6,022 × 10^23^) nguyên tử hoặc phân tử của chất đó.",
    "- *Khối lượng mol (M):* Là khối lượng tính bằng gam của 1 mol chất đó (đơn vị: g/mol). Công thức: *M = m / n* hay *m = n × M*.",
    "- *Thể tích mol chất khí (ở 25 °C, 1 bar):* 1 mol bất kì chất khí nào đều chiếm thể tích là 24,79 lít. Công thức: *V = n × 24,79* (lít) hay *n = V / 24,79* (mol).",
    "- *Tỉ khối của chất khí:*",
    "  + Tỉ khối của khí A đối với khí B: *d~A/B~ = M~A~ / M~B~* (cho biết khí A nặng hay nhẹ hơn khí B bao nhiêu lần).",
    "  + Tỉ khối của khí A đối với không khí: *d~A/kk~ = M~A~ / 29* (với khối lượng mol trung bình của không khí *M~kk~ ≈ 29 g/mol*).",
    "*Bài 4. Dung dịch và nồng độ*",
    "- *Dung dịch:* Là hỗn hợp đồng nhất của dung môi và chất tan.",
    "- *Độ tan (S):* Số gam chất đó tan được trong 100 g nước để tạo thành dung dịch bão hòa ở nhiệt độ xác định.",
    "- *Nồng độ phần trăm (C%):* Cho biết số gam chất tan có trong 100 gam dung dịch. Công thức: *C% = (m~ct~ / m~dd~) × 100%*, với *m~dd~ = m~ct~ + m~dm~*.",
    "- *Nồng độ mol (C~M~):* Cho biết số mol chất tan trong 1 lít dung dịch. Công thức: *C~M~ = n / V* (mol/L hoặc M).",
    "*Bài 5. Định luật bảo toàn khối lượng và phương trình hoá học*",
    "- *Định luật bảo toàn khối lượng:* Trong một phản ứng hoá học, tổng khối lượng của các chất sản phẩm bằng tổng khối lượng của các chất phản ứng. Biểu thức: *m~A~ + m~B~ = m~C~ + m~D~*.",
    "- *Phương trình hoá học:* Biểu diễn ngắn gọn phản ứng hoá học bằng công thức hoá học. Các bước lập PTHH: Viết sơ đồ phản ứng -> Cân bằng số nguyên tử của mỗi nguyên tố ở 2 vế -> Hoàn thành PTHH.",
    "- *Ý nghĩa:* Cho biết tỉ lệ về số nguyên tử, số phân tử giữa các chất trong phản ứng."
]

KHTN1_MCQ = [
    {
        "num": 1,
        "question": "Quá trình nào sau đây là sự biến đổi hóa học?",
        "options": [
            "A. Sương đọng lại trên lá cây.",
            "B. Thức ăn để lâu bị ôi thiu.",
            "C. Hoà tan muối ăn vào nước tạo thành nước muối.",
            "D. Dây sắt cắt nhỏ và tán thành đinh."
        ],
        "correct": "B",
        "solution": "- Phương án B là sự biến đổi hoá học vì thức ăn bị ôi thiu là do các chất hữu cơ trong thức ăn bị vi sinh vật phân huỷ tạo thành các chất mới có mùi chua, hôi, biến đổi màu sắc.\n- Phương án A, C, D chỉ là sự biến đổi vật lí (ngưng tụ hơi nước, hoà tan tạo hỗn hợp đồng nhất, thay đổi hình dạng kích thước), không sinh ra chất mới."
    },
    {
        "num": 2,
        "question": "Quá trình nào sau đây là sự biến đổi vật lí?",
        "options": [
            "A. Cho vôi sống CaO vào nước được vôi tôi.",
            "B. Cồn để trong lọ không kín bị bay hơi.",
            "C. Khí carbon dioxide làm đục nước vôi trong.",
            "D. Nung đá vôi thành vôi sống."
        ],
        "correct": "B",
        "solution": "- Phương án B là sự biến đổi vật lí vì cồn bay hơi chỉ là sự chuyển thể từ lỏng sang hơi, không tạo ra chất mới.\n- Phương án A, C, D đều có phản ứng hoá học xảy ra tạo chất mới:\n  + A: CaO + H~2~O -> Ca(OH)~2~\n  + C: CO~2~ + Ca(OH)~2~ -> CaCO~3~ + H~2~O\n  + D: CaCO~3~ --t°--> CaO + CO~2~"
    },
    {
        "num": 3,
        "question": "Phản ứng hóa học chỉ xảy ra khi:",
        "options": [
            "A. Đun nóng hóa chất.",
            "B. Có chất xúc tác.",
            "C. Các chất tham gia phản ứng ở gần nhau.",
            "D. Các chất tham gia phản ứng tiếp xúc nhau, có trường hợp cần đun nóng, có trường hợp cần chất xúc tác."
        ],
        "correct": "D",
        "solution": "- Điều kiện tiên quyết để phản ứng hoá học xảy ra là các chất phản ứng phải tiếp xúc trực tiếp với nhau.\n- Tuỳ từng phản ứng cụ thể, có trường hợp phản ứng xảy ra ngay ở nhiệt độ thường, có trường hợp cần đun nóng hoặc cần có mặt chất xúc tác."
    },
    {
        "num": 4,
        "question": "Hiện tượng nào là hiện tượng hoá học trong các hiện tượng thiên nhiên sau đây?",
        "options": [
            "A. Sáng sớm, khi mặt trời mọc sương mù tan dần.",
            "B. Hơi nước trong các đám mây ngưng tụ và rơi xuống tạo ra mưa.",
            "C. Nạn cháy rừng tạo khói đen dày đặc gây ô nhiễm môi trường.",
            "D. Khi mưa giông thường có sấm sét."
        ],
        "correct": "C",
        "solution": "- Cháy rừng là phản ứng cháy của gỗ, lá cây (chất hữu cơ) với khí oxygen trong không khí tạo thành khí carbon dioxide, hơi nước, tro và khói bụi (muội than C, CO...) -> có sinh ra chất mới nên là hiện tượng hoá học.\n- Sương mù tan, mây ngưng tụ tạo mưa là các quá trình chuyển thể của nước (vật lí); sấm sét là hiện tượng phóng tia lửa điện giữa các đám mây (vật lí)."
    },
    {
        "num": 5,
        "question": "Chất khí nặng hơn không khí là:",
        "options": [
            "A. Khí methane (CH~4~)",
            "B. Khí carbon dioxide (CO~2~)",
            "C. Khí Helium (He)",
            "D. Khí hydrogen (H~2~)"
        ],
        "correct": "B",
        "solution": "- Khối lượng mol trung bình của không khí là: M~kk~ ≈ 29 g/mol.\n- Xét khối lượng mol của các khí:\n  + M~CH4~ = 12 + 4 = 16 g/mol < 29 (nhẹ hơn không khí)\n  + M~CO2~ = 12 + 16 × 2 = 44 g/mol > 29 (nặng hơn không khí, d~CO2/kk~ = 44 / 29 ≈ 1,52)\n  + M~He~ = 4 g/mol < 29 (nhẹ hơn không khí)\n  + M~H2~ = 2 g/mol < 29 (nhẹ hơn không khí)\n-> Khí nặng hơn không khí là CO~2~."
    },
    {
        "num": 6,
        "question": "Khí nào nhẹ nhất trong tất cả các khí?",
        "options": [
            "A. Khí methane (CH~4~)",
            "B. Khí carbon oxide (CO)",
            "C. Khí nitrogen (N~2~)",
            "D. Khí hydrogen (H~2~)"
        ],
        "correct": "D",
        "solution": "- Khí hydrogen (H~2~) có khối lượng mol là: M~H2~ = 2 g/mol.\n- Đây là khối lượng mol nhỏ nhất trong tất cả các đơn chất và hợp chất khí đã biết, do đó H~2~ là khí nhẹ nhất."
    },
    {
        "num": 7,
        "question": "Khối lượng dung dịch đường thu được khi hòa tan 40 gam đường với nước được dung dịch đường 20%.",
        "options": [
            "A. 150 gam.",
            "B. 200 gam.",
            "C. 170 gam.",
            "D. 250 gam."
        ],
        "correct": "B",
        "solution": "Khối lượng dung dịch đường thu được là:\nm~dd~ = (m~ct~ × 100%) / C% = (40 × 100%) / 20% = 200 (gam)."
    },
    {
        "num": 8,
        "question": "Phản ứng hóa học là:",
        "options": [
            "A. Quá trình kết hợp các đơn chất thành hợp chất.",
            "B. Quá trình biến đổi chất này thành chất khác.",
            "C. Sự trao đổi của 2 hay nhiều chất ban đầu để tạo chất mới.",
            "D. Là quá trình phân hủy chất ban đầu thành nhiều chất."
        ],
        "correct": "B",
        "solution": "Theo định nghĩa SGK KHTN 8: Phản ứng hoá học là quá trình biến đổi từ chất này thành chất khác."
    },
    {
        "num": 9,
        "question": "Phản ứng tỏa nhiệt là",
        "options": [
            "A. Phản ứng có nhiệt độ lớn hơn môi trường xung quanh.",
            "B. Phản ứng có nhiệt độ nhỏ hơn môi trường xung quanh.",
            "C. Phản ứng có nhiệt độ bằng môi trường xung quanh.",
            "D. Phản ứng không có sự thay đổi nhiệt độ."
        ],
        "correct": "A",
        "solution": "- Phản ứng toả nhiệt là phản ứng hoá học giải phóng năng lượng dưới dạng nhiệt ra môi trường xung quanh.\n- Hỗn hợp phản ứng nóng lên, truyền nhiệt ra ngoài làm cho nhiệt độ của hệ phản ứng lớn hơn nhiệt độ môi trường xung quanh."
    },
    {
        "num": 10,
        "question": "Phản ứng thu nhiệt là",
        "options": [
            "A. Phản ứng có nhiệt độ lớn hơn môi trường xung quanh.",
            "B. Phản ứng có nhiệt độ nhỏ hơn môi trường xung quanh.",
            "C. Phản ứng có nhiệt độ bằng môi trường xung quanh.",
            "D. Phản ứng không có sự thay đổi nhiệt độ."
        ],
        "correct": "B",
        "solution": "- Phản ứng thu nhiệt là phản ứng hoá học hấp thụ năng lượng dưới dạng nhiệt từ môi trường xung quanh.\n- Hệ phản ứng lấy nhiệt từ môi trường nên nhiệt độ của hỗn hợp phản ứng hạ xuống thấp hơn (nhỏ hơn) nhiệt độ môi trường xung quanh."
    },
    {
        "num": 11,
        "question": "Đốt cháy xăng, dầu trong các động cơ là",
        "options": [
            "A. Phản ứng thu nhiệt.",
            "B. Phản ứng tỏa nhiệt.",
            "C. Phản ứng phân hủy.",
            "D. Phản ứng ứng thế."
        ],
        "correct": "B",
        "solution": "Quá trình đốt cháy nhiên liệu (xăng, dầu) giải phóng năng lượng nhiệt và ánh sáng rất lớn để giãn nở khí sinh công đẩy động cơ hoạt động -> Phản ứng toả nhiệt."
    },
    {
        "num": 12,
        "question": "Chất khí nào sau đây nhẹ hơn khí methane (CH~4~)?",
        "options": [
            "A. H~2~",
            "B. N~2~",
            "C. CO",
            "D. NO"
        ],
        "correct": "A",
        "solution": "- Khối lượng mol của khí methane: M~CH4~ = 12 + 4 = 16 g/mol.\n- Xét khối lượng mol các khí:\n  + M~H2~ = 2 g/mol < 16 -> H~2~ nhẹ hơn CH~4~.\n  + M~N2~ = 28 g/mol > 16.\n  + M~CO~ = 12 + 16 = 28 g/mol > 16.\n  + M~NO~ = 14 + 16 = 30 g/mol > 16.\n-> Chỉ có khí H~2~ nhẹ hơn CH~4~."
    },
    {
        "num": 13,
        "question": "Điền vào chỗ trống: \"Trong một phản ứng hóa học, tổng khối lượng của các chất sản phẩm ... tổng khối lượng của các chất phản ứng.\"",
        "options": [
            "A. lớn hơn.",
            "B. nhỏ hơn.",
            "C. bằng.",
            "D. nhỏ hơn hoặc bằng."
        ],
        "correct": "C",
        "solution": "Theo định luật bảo toàn khối lượng: \"Trong một phản ứng hoá học, tổng khối lượng của các chất sản phẩm bằng tổng khối lượng của các chất phản ứng.\""
    },
    {
        "num": 14,
        "question": "Có phương trình hóa học: 2H~2~ + O~2~ --t°--> 2H~2~O. Theo định luật bảo toàn khối lượng thì:",
        "options": [
            "A. m~H2~ + m~O2~ = m~H2O~.",
            "B. m~H2~ - m~O2~ = m~H2O~.",
            "C. m~O2~ = m~H2O~ + m~H2~.",
            "D. m~H2~ = m~H2O~ + m~O2~."
        ],
        "correct": "A",
        "solution": "Theo định luật bảo toàn khối lượng, tổng khối lượng các chất tham gia bằng tổng khối lượng các chất sản phẩm:\nm~H2~ + m~O2~ = m~H2O~ (chú ý kí hiệu khối lượng m không nhân hệ số tỉ lượng của phương trình)."
    },
    {
        "num": 15,
        "question": "Tính khối lượng FeS tạo thành trong phản ứng của Fe và S, biết khối lượng của Fe và S đã tham gia phản ứng lần lượt là 7 gam và 4 gam.",
        "options": [
            "A. 3 gam.",
            "B. 10 gam.",
            "C. 22 gam.",
            "D. 11 gam."
        ],
        "correct": "D",
        "solution": "- Phương trình phản ứng: Fe + S --t°--> FeS\n- Áp dụng định luật bảo toàn khối lượng:\nm~FeS~ = m~Fe~ + m~S~ = 7 + 4 = 11 (gam)."
    }
]

KHTN1_TF = [
    {
        "num": 1,
        "intro": "Các nhận định về phản ứng hóa học:",
        "items": [
            {
                "label": "a",
                "text": "Trong phản ứng hóa học, khi chất biến đổi làm các nguyên tử bị biến đổi.",
                "ans": "Sai",
                "explain": "Trong phản ứng hoá học, các nguyên tử được bảo toàn (không sinh ra, không mất đi và không bị biến đổi thành nguyên tử khác), chỉ có liên kết giữa các nguyên tử thay đổi."
            },
            {
                "label": "b",
                "text": "Trong phương trình hoá học, cần đặt hệ số thích hợp vào công thức của các chất sao cho số nguyên tử của mỗi nguyên tố ở 2 vế đều bằng nhau.",
                "ans": "Đúng",
                "explain": "Đây là nguyên tắc cơ bản để cân bằng phương trình hoá học, bảo đảm tuân thủ định luật bảo toàn khối lượng."
            },
            {
                "label": "c",
                "text": "Trong phản ứng hoá học, biết tổng khối lượng của các chất phản ứng ta biết được tổng khối lượng các sản phẩm.",
                "ans": "Đúng",
                "explain": "Theo định luật bảo toàn khối lượng: m~sản phẩm~ = m~chất phản ứng~."
            },
            {
                "label": "d",
                "text": "Trong phản ứng hoá học, màu sắc của các chất có thể bị thay đổi.",
                "ans": "Đúng",
                "explain": "Sự đổi màu là một trong những dấu hiệu nhận biết có phản ứng hoá học xảy ra."
            }
        ]
    },
    {
        "num": 2,
        "intro": "Cho phản ứng hóa học: 3Cu + 8HNO~3~ --t°--> 3Cu(NO~3~)~2~ + 2NO + 4H~2~O",
        "items": [
            {
                "label": "a",
                "text": "Theo định luật bảo toàn khối lượng, ta có biểu thức: 3.m~Cu~ + 8.m~HNO3~ = 3.m~Cu(NO3)2~ + 2.m~NO~ + 4.m~H2O~.",
                "ans": "Sai",
                "explain": "Biểu thức đúng là: m~Cu~ + m~HNO3~ = m~Cu(NO3)2~ + m~NO~ + m~H2O~. Khối lượng m của chất không nhân thêm các hệ số tỉ lượng."
            },
            {
                "label": "b",
                "text": "Trước phản ứng có 8 nguyên tử nitrogen; Sau phản ứng có 2 nguyên tử nitrogen.",
                "ans": "Sai",
                "explain": "Trước phản ứng có 8 nguyên tử N (trong 8HNO~3~). Sau phản ứng có: 3 × 2 (trong Cu(NO~3~)~2~) + 2 × 1 (trong NO) = 8 nguyên tử N. Số nguyên tử N được bảo toàn."
            },
            {
                "label": "c",
                "text": "Trước phản ứng có 11 nguyên tử oxygen; Sau phản ứng có 11 nguyên tử oxygen.",
                "ans": "Sai",
                "explain": "Trước phản ứng có 8 × 3 = 24 nguyên tử O. Sau phản ứng có: 3 × 6 + 2 × 1 + 4 × 1 = 24 nguyên tử O (không phải 11)."
            },
            {
                "label": "d",
                "text": "Trước phản ứng có 3 nguyên tử đồng; Sau phản ứng có 3 nguyên tử đồng.",
                "ans": "Đúng",
                "explain": "Cả hai vế của phương trình đều có 3 nguyên tử Cu."
            }
        ]
    },
    {
        "num": 3,
        "intro": "Cho phản ứng hóa học: MgCl~2~ + 2KOH ---> Mg(OH)~2~ + 2KCl",
        "items": [
            {
                "label": "a",
                "text": "Theo định luật bảo toàn khối lượng, ta có biểu thức: m~MgCl2~ + 2.m~KOH~ = m~Mg(OH)2~ + m~KCl~.",
                "ans": "Sai",
                "explain": "Biểu thức đúng theo định luật bảo toàn khối lượng là: m~MgCl2~ + m~KOH~ = m~Mg(OH)2~ + m~KCl~."
            },
            {
                "label": "b",
                "text": "Nếu tổng khối lượng sản phẩm là 103,5 gam và khối lượng KOH tham gia phản ứng là 56 gam thì khối lượng MgCl~2~ đã phản ứng là 47,5 gam.",
                "ans": "Đúng",
                "explain": "Áp dụng ĐLBTKL: m~MgCl2~ = m~sp~ - m~KOH~ = 103,5 - 56 = 47,5 gam."
            },
            {
                "label": "c",
                "text": "Nếu tổng khối lượng chất tham gia phản ứng là 51,75 kg và khối lượng KCl thu được là 37,25 kg thì khối lượng Mg(OH)~2~ thu được là 14 500 gam.",
                "ans": "Đúng",
                "explain": "Ta có: m~Mg(OH)2~ = m~tg~ - m~KCl~ = 51,75 - 37,25 = 14,5 kg = 14 500 gam."
            },
            {
                "label": "d",
                "text": "Nếu đem 9,5 gam MgCl~2~ phản ứng hoàn toàn với 11,2 gam KOH thì thu được 5,8 gam Mg(OH)~2~ và 14,9 gam KCl.",
                "ans": "Đúng",
                "explain": "Số mol: n~MgCl2~ = 9,5 / 95 = 0,1 mol; n~KOH~ = 11,2 / 56 = 0,2 mol. Tỉ lệ mol vừa đủ 1 : 2. Theo PTHH sinh ra 0,1 mol Mg(OH)~2~ (m = 0,1 × 58 = 5,8 gam) và 0,2 mol KCl (m = 0,2 × 74,5 = 14,9 gam). Thỏa mãn BTKL: 9,5 + 11,2 = 5,8 + 14,9 = 20,7 gam."
            }
        ]
    }
]

KHTN1_SHORT = [
    {
        "num": 1,
        "question": "Nến được làm bằng paraffin, khi đốt nến, xảy ra các quá trình sau:\n(1) Paraffin nóng chảy;\n(2) Paraffin lỏng chuyển thành hơi;\n(3) Hơi paraffin cháy biến đổi thành khí CO~2~ và hơi nước.\nQuá trình có sự biến đổi hoá học là: ......",
        "ans": "(3)",
        "solution": "Quá trình (1) và (2) chỉ là sự chuyển thể của paraffin (rắn -> lỏng -> hơi) nên là biến đổi vật lí. Quá trình (3) hơi paraffin phản ứng với O~2~ tạo thành chất mới là CO~2~ và H~2~O nên là sự biến đổi hoá học."
    },
    {
        "num": 2,
        "question": "Cho các hiện tượng sau:\n(1) Dưa muối lên men;\n(2) Hydrogen cháy trong không khí;\n(3) Hiệu ứng nhà kính làm Trái Đất nóng lên;\n(4) Mưa acid;\n(5) Vào mùa hè băng ở 2 cực tan chảy.\nSố hiện tượng hóa học là: ..........",
        "ans": "3",
        "solution": "Có 3 hiện tượng hoá học:\n- (1) Dưa muối lên men: Quá trình lên men lactic tạo acid lactic.\n- (2) Hydrogen cháy trong không khí: 2H~2~ + O~2~ --t°--> 2H~2~O.\n- (4) Mưa acid: Các khí SO~2~, NO~x~ tác dụng với nước và oxygen tạo acid H~2~SO~4~, HNO~3~.\n(3) và (5) chỉ là hiện tượng vật lí (hấp thụ nhiệt và nóng chảy)."
    },
    {
        "num": 3,
        "question": "Cho phản ứng hóa học: MgO + 2HCl ---> MgCl~2~ + H~2~O. Biết tổng khối lượng hai chất tham gia phản ứng là 11,3 gam và khối lượng nước thu được sau phản ứng là 1,8 gam. Khối lượng MgCl~2~ là .........",
        "ans": "9,5 gam",
        "solution": "Theo định luật bảo toàn khối lượng:\nm~MgO~ + m~HCl~ = m~MgCl2~ + m~H2O~\n=> m~MgCl2~ = (m~MgO~ + m~HCl~) - m~H2O~ = 11,3 - 1,8 = 9,5 (gam)."
    },
    {
        "num": 4,
        "question": "Muối copper sulfate (CuSO~4~) ngậm nước khi đun nóng sẽ bị tách nước. Nếu đun 25 g muối ngậm nước, thu được 16 g muối khan thì số mol nước tách ra là ............. mol.",
        "ans": "0,5 mol",
        "solution": "- Khối lượng nước bị tách ra là:\nm~H2O~ = m~muối ngậm nước~ - m~muối khan~ = 25 - 16 = 9 (gam).\n- Số mol nước tách ra là:\nn~H2O~ = m / M = 9 / 18 = 0,5 (mol)."
    }
]

KHTN1_ESSAY = [
    {
        "num": 1,
        "question": "Xác định số nguyên tử (phân tử) có trong:\na) 1,5 mol nguyên tử carbon.\nb) 0,25 mol phân tử nước (H~2~O).",
        "solution": "a) Số nguyên tử carbon có trong 1,5 mol C là:\nN~C~ = n × N = 1,5 × 6,022 × 10^23^ = 9,033 × 10^23^ (nguyên tử C).\nb) Số phân tử nước có trong 0,25 mol H~2~O là:\nN~H2O~ = n × N = 0,25 × 6,022 × 10^23^ = 1,5055 × 10^23^ (phân tử H~2~O)."
    },
    {
        "num": 2,
        "question": "Trong các khí sau: C~2~H~2~, SO~2~, N~2~, CO~2~. Khí nào nặng (nhẹ) hơn không khí và nặng (nhẹ) hơn bao nhiêu lần?",
        "solution": "Khối lượng mol của không khí: M~kk~ ≈ 29 g/mol.\n- Khí C~2~H~2~ (M = 26 g/mol < 29): Nhẹ hơn không khí, d~C2H2/kk~ = 26 / 29 ≈ 0,897 lần.\n- Khí SO~2~ (M = 64 g/mol > 29): Nặng hơn không khí, d~SO2/kk~ = 64 / 29 ≈ 2,21 lần.\n- Khí N~2~ (M = 28 g/mol < 29): Nhẹ hơn không khí, d~N2/kk~ = 28 / 29 ≈ 0,966 lần.\n- Khí CO~2~ (M = 44 g/mol > 29): Nặng hơn không khí, d~CO2/kk~ = 44 / 29 ≈ 1,52 lần."
    },
    {
        "num": 3,
        "question": "a) Hòa tan 30 gam muối ăn với nước được dung dịch muối ăn 25% (đề ghi dung dịch đường 25%). Tính khối lượng dung dịch thu được?\nb) Tính số mol KOH có trong 300 mL KOH 5M?",
        "solution": "a) Khối lượng dung dịch thu được là:\nm~dd~ = (m~ct~ × 100%) / C% = (30 × 100%) / 25% = 120 (gam).\nb) Đổi thể tích: V = 300 mL = 0,3 lít.\nSố mol KOH có trong dung dịch là:\nn~KOH~ = C~M~ × V = 5 × 0,3 = 1,5 (mol)."
    },
    {
        "num": 4,
        "question": "Hãy chỉ ra chất tham gia phản ứng và chất sản phẩm và lập phương trình hoá học của các phản ứng hóa học sau:\n(1) Phosphorus tác dụng với khí oxygen (O~2~) tạo thành phosphorus(V)oxide (P~2~O~5~).\n(2) Khí nitrogen và khí hydrogen tác dụng với nhau tạo thành khí ammonia (NH~3~).\n(3) Cho một thanh nhôm (aluminium) tác dụng với dung dịch hydrochloric acid (HCl) thu được muối aluminium chloride (AlCl~3~) và thấy có khí hydrogen thoát ra.\n(4) Magnesium (Mg) tác dụng với khí oxygen tạo thành magnesium oxide (MgO).\n(5) Cho dung dịch sodium carbonate (Na~2~CO~3~) tác dụng với dung dịch calcium hydroxide (Ca(OH)~2~) tạo thành calcium carbonate (CaCO~3~) không tan và sodium hydroxide (NaOH).\n(6) Khí hydrogen (H~2~) tác dụng với oxide sắt từ (Fe~3~O~4~) tạo thành iron (Fe) và Nước (H~2~O).",
        "solution": "(1) Chất tham gia: Phosphorus (P) và oxygen (O~2~). Sản phẩm: Phosphorus(V) oxide (P~2~O~5~).\n    PTHH: 4P + 5O~2~ --t°--> 2P~2~O~5~\n(2) Chất tham gia: Nitrogen (N~2~) và hydrogen (H~2~). Sản phẩm: Ammonia (NH~3~).\n    PTHH: N~2~ + 3H~2~ <==t°,xt,p==> 2NH~3~\n(3) Chất tham gia: Aluminium (Al) và hydrochloric acid (HCl). Sản phẩm: Aluminium chloride (AlCl~3~) và hydrogen (H~2~).\n    PTHH: 2Al + 6HCl ---> 2AlCl~3~ + 3H~2~↑\n(4) Chất tham gia: Magnesium (Mg) và oxygen (O~2~). Sản phẩm: Magnesium oxide (MgO).\n    PTHH: 2Mg + O~2~ --t°--> 2MgO\n(5) Chất tham gia: Sodium carbonate (Na~2~CO~3~) và calcium hydroxide (Ca(OH)~2~). Sản phẩm: Calcium carbonate (CaCO~3~) và sodium hydroxide (NaOH).\n    PTHH: Na~2~CO~3~ + Ca(OH)~2~ ---> CaCO~3~↓ + 2NaOH\n(6) Chất tham gia: Hydrogen (H~2~) và oxide sắt từ (Fe~3~O~4~). Sản phẩm: Iron (Fe) và nước (H~2~O).\n    PTHH: Fe~3~O~4~ + 4H~2~ --t°--> 3Fe + 4H~2~O"
    },
    {
        "num": 5,
        "question": "Lập phương trình hóa học và cho biết ý nghĩa của các phương trình hóa học sau đây:\n1. Fe + O~2~ --t°--> Fe~2~O~3~\n2. KClO~3~ --t°--> KCl + O~2~\n3. NaOH + CuSO~4~ ---> Cu(OH)~2~ + Na~2~SO~4~\n4. CuCl~2~ + KOH ---> Cu(OH)~2~ + KCl",
        "solution": "1. 4Fe + 3O~2~ --t°--> 2Fe~2~O~3~\n- Ý nghĩa: Cứ 4 nguyên tử Fe tác dụng với 3 phân tử O~2~ tạo thành 2 phân tử Fe~2~O~3~.\n  Tỉ lệ: Số nguyên tử Fe : Số phân tử O~2~ : Số phân tử Fe~2~O~3~ = 4 : 3 : 2.\n2. 2KClO~3~ --t°--> 2KCl + 3O~2~↑\n- Ý nghĩa: Cứ 2 phân tử KClO~3~ bị nhiệt phân huỷ tạo thành 2 phân tử KCl và 3 phân tử O~2~.\n  Tỉ lệ: Số phân tử KClO~3~ : Số phân tử KCl : Số phân tử O~2~ = 2 : 2 : 3.\n3. 2NaOH + CuSO~4~ ---> Cu(OH)~2~↓ + Na~2~SO~4~\n- Ý nghĩa: Cứ 2 phân tử NaOH tác dụng với 1 phân tử CuSO~4~ tạo thành 1 phân tử Cu(OH)~2~ và 1 phân tử Na~2~SO~4~.\n  Tỉ lệ: Số phân tử NaOH : Số phân tử CuSO~4~ : Số phân tử Cu(OH)~2~ : Số phân tử Na~2~SO~4~ = 2 : 1 : 1 : 1.\n4. CuCl~2~ + 2KOH ---> Cu(OH)~2~↓ + 2KCl\n- Ý nghĩa: Cứ 1 phân tử CuCl~2~ tác dụng với 2 phân tử KOH tạo thành 1 phân tử Cu(OH)~2~ và 2 phân tử KCl.\n  Tỉ lệ: Số phân tử CuCl~2~ : Số phân tử KOH : Số phân tử Cu(OH)~2~ : Số phân tử KCl = 1 : 2 : 1 : 2."
    }
]

# ==============================================================================
# PHÂN MÔN KHTN 2 (VẬT LÍ)
# ==============================================================================

KHTN2_THEORY = [
    "*A. LÝ THUYẾT: KIẾN THỨC TRỌNG TÂM*",
    "- *Sử dụng một số dụng cụ đo trong học tập và nghiên cứu khoa học:* Đo thể tích (bình chia độ, bình tràn), khối lượng (cân điện tử, cân Robecvan), thời gian (đồng hồ bấm giây), nhiệt độ (nhiệt kế), chiều dài (thước cuộn, thước kẻ, thước kẹp).",
    "- *Khối lượng riêng:*",
    "  + *Khái niệm:* Khối lượng riêng của một chất là khối lượng của một đơn vị thể tích (1 m^3^) chất đó.",
    "  + *Công thức:* *D = m / V* (trong đó *m* là khối lượng, đơn vị kg; *V* là thể tích, đơn vị m^3^).",
    "  + *Đơn vị đo chuẩn trong hệ SI:* *kg/m^3^* (hoặc g/cm^3^, với 1 g/cm^3^ = 1000 kg/m^3^).",
    "- *Áp suất:*",
    "  + *Áp lực:* Là lực ép có phương vuông góc với mặt bị ép.",
    "  + *Áp suất:* Là độ lớn của áp lực trên một đơn vị diện tích bị ép.",
    "  + *Công thức:* *p = F / S* (trong đó *F* là áp lực, đơn vị N; *S* là diện tích bị ép, đơn vị m^2^).",
    "  + *Đơn vị:* Pascal (Pa) hoặc N/m^2^ (1 Pa = 1 N/m^2^).",
    "  + *Cách làm tăng áp suất:* Tăng áp lực *F* và/hoặc giảm diện tích bị ép *S*.",
    "  + *Cách làm giảm áp suất:* Giảm áp lực *F* và/hoặc tăng diện tích bị ép *S*.",
    "- *Áp suất chất lỏng và áp suất khí quyển:*",
    "  + Chất lỏng gây áp suất theo mọi phương lên đáy bình, thành bình và các vật nằm trong lòng chất lỏng.",
    "  + Công thức tính áp suất chất lỏng: *p = d × h* (trong đó *d* là trọng lượng riêng của chất lỏng, đơn vị N/m^3^; *h* là độ sâu tính từ mặt thoáng chất lỏng, đơn vị m).",
    "  + Không khí có trọng lượng nên Trái Đất và mọi vật trên Trái Đất đều chịu áp suất khí quyển theo mọi phương.",
    "- *Lực đẩy Archimedes:*",
    "  + Tác dụng: Vật nhúng trong chất lỏng chịu lực đẩy thẳng đứng hướng từ dưới lên trên.",
    "  + Công thức tính độ lớn: *F~A~ = d × V* (trong đó *d* là trọng lượng riêng của chất lỏng, *V* là thể tích phần chất lỏng bị vật chiếm chỗ).",
    "  + *Điều kiện nổi, chìm:*",
    "    * Vật chìm xuống khi trọng lực lớn hơn lực đẩy Archimedes: *P > F~A~* (hay *d~vật~ > d~chất lỏng~*).",
    "    * Vật lơ lửng khi trọng lực bằng lực đẩy Archimedes: *P = F~A~* (hay *d~vật~ = d~chất lỏng~*).",
    "    * Vật nổi lên khi trọng lực nhỏ hơn lực đẩy Archimedes: *P < F~A~* (hay *d~vật~ < d~chất lỏng~*). Khi nổi cân bằng trên mặt thoáng: *P = F~A~'*."
]

KHTN2_MCQ = [
    {
        "num": 1,
        "question": "Để đo thể tích của một vật rắn không thấm nước và chìm hoàn toàn trong nước bằng bình tràn và bình chứa, ta thực hiện thao tác nào sau đây?",
        "options": [
            "A. Đổ nước vào bình tràn tới ngập miệng vòi tràn, thả vật vào bình chứa.",
            "B. Đổ nước tới vòi tràn, thả vật vào bình tràn, thể tích phần nước tràn ra bình chứa bằng thể tích của vật.",
            "C. Thả vật vào bình chứa không có nước rồi đổ nước vào cho ngập vật.",
            "D. Đổ nước tới vòi tràn, thả vật vào bình tràn, phần nước tràn ra bình chứa có thể tích bằng phần chìm của vật."
        ],
        "correct": "B",
        "solution": "Khi đổ nước tới miệng vòi tràn, thả chìm hoàn toàn vật vào bình tràn thì toàn bộ phần nước tràn ra bình chứa đúng bằng thể tích của vật rắn."
    },
    {
        "num": 2,
        "question": "Đơn vị đo khối lượng riêng trong hệ đơn vị SI là:",
        "options": [
            "A. kg/m^3^.",
            "B. g/cm^3^.",
            "C. g/L.",
            "D. kg/L."
        ],
        "correct": "A",
        "solution": "Trong hệ đơn vị quốc tế SI, khối lượng đo bằng kg, thể tích đo bằng m^3^ nên đơn vị của khối lượng riêng là kg/m^3^."
    },
    {
        "num": 3,
        "question": "Công thức tính khối lượng riêng của một chất là:",
        "options": [
            "A. D = m / V.",
            "B. D = m × V.",
            "C. D = V / m.",
            "D. D = P / V."
        ],
        "correct": "A",
        "solution": "Công thức định nghĩa khối lượng riêng: D = m / V."
    },
    {
        "num": 4,
        "question": "Một khối đá có thể tích 0,5 m^3^ và khối lượng riêng là 2600 kg/m^3^. Khối lượng của khối đá đó là:",
        "options": [
            "A. 5200 kg.",
            "B. 1300 kg.",
            "C. 2600 kg.",
            "D. 650 kg."
        ],
        "correct": "B",
        "solution": "Khối lượng của khối đá là: m = D × V = 2600 × 0,5 = 1300 (kg)."
    },
    {
        "num": 5,
        "question": "Áp suất được tính bằng công thức nào sau đây?",
        "options": [
            "A. p = F / S.",
            "B. p = F × S.",
            "C. p = S / F.",
            "D. p = d / V."
        ],
        "correct": "A",
        "solution": "Công thức tính áp suất: p = F / S (trong đó F là áp lực, S là diện tích bị ép)."
    },
    {
        "num": 6,
        "question": "Trong các đơn vị sau đây, đơn vị nào là đơn vị của áp suất?",
        "options": [
            "A. Newton (N).",
            "B. Pascal (Pa).",
            "C. kg/m^3^.",
            "D. Joule (J)."
        ],
        "correct": "B",
        "solution": "Pascal (Pa) hoặc N/m^2^ là đơn vị của áp suất. (N là đơn vị lực, kg/m^3^ là đơn vị khối lượng riêng, J là đơn vị công/năng lượng)."
    },
    {
        "num": 7,
        "question": "Muốn tăng áp suất tác dụng lên mặt sàn, ta có thể thực hiện cách nào sau đây?",
        "options": [
            "A. Giảm áp lực và tăng diện tích bị ép.",
            "B. Tăng áp lực và tăng diện tích bị ép.",
            "C. Tăng áp lực và giữ nguyên hoặc giảm diện tích bị ép.",
            "D. Giảm áp lực và giảm diện tích bị ép."
        ],
        "correct": "C",
        "solution": "Từ công thức p = F / S: Để tăng áp suất p, ta cần tăng lực ép F và/hoặc giảm diện tích bị ép S."
    },
    {
        "num": 8,
        "question": "Áp suất tác dụng lên một điểm ở trong lòng chất lỏng không phụ thuộc vào yếu tố nào sau đây?",
        "options": [
            "A. Khối lượng riêng của chất lỏng.",
            "B. Trọng lượng riêng của chất lỏng.",
            "C. Độ sâu của điểm tính từ thoáng chất lỏng.",
            "D. Khối lượng của toàn bộ chất lỏng trong bình chứa."
        ],
        "correct": "D",
        "solution": "Áp suất chất lỏng p = d × h chỉ phụ thuộc vào trọng lượng riêng d (hoặc khối lượng riêng D) của chất lỏng và độ sâu h tính từ mặt thoáng, không phụ thuộc vào tổng khối lượng chất lỏng hay hình dạng bình chứa."
    },
    {
        "num": 9,
        "question": "Công thức tính áp suất chất lỏng ở độ sâu h tính từ thoáng chất lỏng là:",
        "options": [
            "A. p = d × h.",
            "B. p = d / h.",
            "C. p = d × V.",
            "D. p = F / S."
        ],
        "correct": "A",
        "solution": "Công thức tính áp suất của chất lỏng đứng yên ở độ sâu h: p = d × h."
    },
    {
        "num": 10,
        "question": "Hiện tượng nào sau đây chứng minh sự tồn tại của áp suất khí quyển?",
        "options": [
            "A. Quả bóng bàn bị bẹp nhúng vào nước nóng lại phồng lên như cũ.",
            "B. Cầm một viên phấn thả rơi xuống đất.",
            "C. Có thể hút nước bằng ống hút nhựa từ cốc nước giải khát.",
            "D. Nước chảy từ trên cao xuống thấp."
        ],
        "correct": "C",
        "solution": "Khi hút không khí trong ống hút ra, áp suất trong ống giảm; áp suất khí quyển tác dụng lên mặt nước giải khát lớn hơn ép nước dâng lên trong ống hút vào miệng -> chứng minh có áp suất khí quyển."
    },
    {
        "num": 11,
        "question": "Lực đẩy Archimedes tác dụng lên một vật nhúng trong chất lỏng có chiều:",
        "options": [
            "A. Từ trên xuống dưới.",
            "B. Theo chiều của trọng lực.",
            "C. Từ dưới lên theo phương thẳng đứng.",
            "D. Theo mọi phương."
        ],
        "correct": "C",
        "solution": "Lực đẩy Archimedes luôn có phương thẳng đứng, chiều từ dưới lên trên (ngược chiều trọng lực)."
    },
    {
        "num": 12,
        "question": "Độ lớn của lực đẩy Archimedes tác dụng lên một vật nhúng hoàn toàn trong chất lỏng được tính bằng công thức:",
        "options": [
            "A. F~A~ = d × V.",
            "B. F~A~ = p × S.",
            "C. F~A~ = m × g.",
            "D. F~A~ = D × h."
        ],
        "correct": "A",
        "solution": "Công thức tính độ lớn lực đẩy Archimedes: F~A~ = d × V (trong đó d là trọng lượng riêng chất lỏng, V là thể tích phần chất lỏng bị vật chiếm chỗ)."
    },
    {
        "num": 13,
        "question": "Một vật nhúng vào trong nước, vật sẽ nổi lên khi:",
        "options": [
            "A. Trọng lực lớn hơn lực đẩy Archimedes (P > F~A~).",
            "B. Trọng lực nhỏ hơn lực đẩy Archimedes (P < F~A~).",
            "C. Trọng lực bằng lực đẩy Archimedes (P = F~A~).",
            "D. Khối lượng của vật nhỏ hơn thể tích của vật."
        ],
        "correct": "B",
        "solution": "Khi P < F~A~, hợp lực hướng lên trên làm vật chuyển động nổi lên trên mặt thoáng chất lỏng."
    },
    {
        "num": 14,
        "question": "Đặt một khối lập phương bằng gỗ có trọng lượng 20 N lên mặt bàn nằm ngang. Diện tích tiếp xúc của khối gỗ với mặt bàn là 0,04 m^2^. Áp suất do khối gỗ tác dụng lên mặt bàn là:",
        "options": [
            "A. 50 Pa.",
            "B. 500 Pa.",
            "C. 800 Pa.",
            "D. 0,8 Pa."
        ],
        "correct": "B",
        "solution": "Áp lực bằng trọng lượng khối gỗ: F = P = 20 N.\nÁp suất tác dụng lên mặt bàn là: p = F / S = 20 / 0,04 = 500 (Pa)."
    },
    {
        "num": 15,
        "question": "Một người đứng trên tuyết bằng đôi giày trượt tuyết có tổng diện tích tiếp xúc là 0,05 m^2^. Trọng lượng của người đó là 600 N. Áp suất người đó tác dụng lên mặt tuyết là:",
        "options": [
            "A. 12000 Pa.",
            "B. 30 Pa.",
            "C. 3000 Pa.",
            "D. 1200 Pa."
        ],
        "correct": "A",
        "solution": "Áp suất do người đó tác dụng lên tuyết là:\np = F / S = P / S = 600 / 0,05 = 12 000 (Pa)."
    }
]

KHTN2_TF = [
    {
        "num": 1,
        "intro": "Các nhận định về khối lượng riêng và đo lường:",
        "items": [
            {
                "label": "a",
                "text": "Khối lượng riêng của một chất cho biết khối lượng của một đơn vị thể tích (1 m^3^) chất đó.",
                "ans": "Đúng",
                "explain": "Định nghĩa: Khối lượng riêng của một chất là khối lượng của một mét khối (1 m^3^) chất đó."
            },
            {
                "label": "b",
                "text": "Sắt có khối lượng riêng khoảng 7800 kg/m^3^ nghĩa là 1 m^3^ sắt có khối lượng là 7800 kg.",
                "ans": "Đúng",
                "explain": "Ý nghĩa định lượng của khối lượng riêng: 1 m^3^ sắt nặng 7800 kg."
            },
            {
                "label": "c",
                "text": "Để đo thể tích vật rắn không thấm nước chìm trong bình chia độ, ta thả chìm vật vào bình, thể tích phần nước dâng lên chính là thể tích của vật.",
                "ans": "Đúng",
                "explain": "Vì vật chìm hoàn toàn nên thể tích nước dâng lên đúng bằng thể tích của vật rắn."
            },
            {
                "label": "d",
                "text": "Dầu hỏa có khối lượng riêng nhỏ hơn nước nên dầu hỏa luôn chìm xuống đáy khi trộn với nước.",
                "ans": "Sai",
                "explain": "Dầu hoả có khối lượng riêng nhỏ hơn nước (khoảng 800 kg/m^3^ so với 1000 kg/m^3^) và không tan trong nước nên dầu hoả nổi lên trên mặt nước."
            }
        ]
    },
    {
        "num": 2,
        "intro": "Cho một vật rắn có thể tích V được nhúng chìm hoàn toàn trong chất lỏng có trọng lượng riêng d:",
        "items": [
            {
                "label": "a",
                "text": "Lực đẩy Archimedes tác dụng lên vật có phương thẳng đứng, chiều từ dưới lên trên.",
                "ans": "Đúng",
                "explain": "Phương và chiều của lực đẩy Archimedes luôn theo phương thẳng đứng hướng từ dưới lên trên."
            },
            {
                "label": "b",
                "text": "Độ lớn lực đẩy Archimedes phụ thuộc vào độ sâu mà vật được nhúng sâu trong chất lỏng.",
                "ans": "Sai",
                "explain": "Khi vật đã chìm hoàn toàn, F~A~ = d × V chỉ phụ thuộc vào trọng lượng riêng d và thể tích V của vật, không phụ thuộc vào độ sâu."
            },
            {
                "label": "c",
                "text": "Nếu nhúng vật đó vào nước thì lực đẩy Archimedes lớn hơn khi nhúng vào dầu (với cùng thể tích phần vật chiếm chỗ).",
                "ans": "Đúng",
                "explain": "Vì trọng lượng riêng của nước (d~nước~ ≈ 10 000 N/m^3^) lớn hơn của dầu (d~dầu~ ≈ 8 000 N/m^3^) nên F~A,nước~ > F~A,dầu~."
            },
            {
                "label": "d",
                "text": "Khi vật lơ lửng trong chất lỏng, trọng lượng của vật đúng bằng độ lớn lực đẩy Archimedes.",
                "ans": "Đúng",
                "explain": "Điều kiện để vật lơ lửng trong lòng chất lỏng: Trọng lực cân bằng với lực đẩy Archimedes (P = F~A~)."
            }
        ]
    },
    {
        "num": 3,
        "intro": "Cho một bình hình trụ chứa nước đứng yên tới độ cao h. Gọi d là trọng lượng riêng của nước:",
        "items": [
            {
                "label": "a",
                "text": "Chất lỏng chỉ gây ra áp suất lên đáy của bình chứa nó.",
                "ans": "Sai",
                "explain": "Chất lỏng gây áp suất theo mọi phương lên cả đáy bình, thành bình và các vật ở trong lòng chất lỏng."
            },
            {
                "label": "b",
                "text": "Cùng một độ sâu, áp suất chất lỏng tác dụng theo mọi phương có độ lớn khác nhau.",
                "ans": "Sai",
                "explain": "Tại cùng một độ sâu trong chất lỏng đứng yên, áp suất tác dụng theo mọi phương có độ lớn bằng nhau."
            },
            {
                "label": "c",
                "text": "Các điểm nằm trên cùng một mặt phẳng nằm ngang trong lòng chất lỏng đứng yên có áp suất bằng nhau.",
                "ans": "Đúng",
                "explain": "Các điểm nằm trên cùng mặt phẳng nằm ngang có cùng độ sâu h nên p = d × h là như nhau."
            },
            {
                "label": "d",
                "text": "Nếu đổ thêm nước vào bình thì áp suất chất lỏng tác dụng lên đáy bình sẽ tăng lên.",
                "ans": "Đúng",
                "explain": "Đổ thêm nước làm tăng chiều cao h của cột nước, theo công thức p = d × h thì áp suất lên đáy bình sẽ tăng."
            }
        ]
    }
]

KHTN2_SHORT = [
    {
        "num": 1,
        "question": "Một khối đá có thể tích 200 dm^3^ (tương đương 0,2 m^3^) và khối lượng 520 kg. Khối lượng riêng của khối đá đó bằng bao nhiêu kg/m^3^?",
        "ans": "2600 kg/m^3^",
        "solution": "Khối lượng riêng của khối đá là: D = m / V = 520 / 0,2 = 2600 (kg/m^3^)."
    },
    {
        "num": 2,
        "question": "Một chiếc xe tăng có trọng lượng 350000 N. Diện tích tiếp xúc của các bản xích xe tăng với mặt đất là 2,5 m^2^. Áp suất của xe tăng tác dụng lên mặt đất là bao nhiêu Pascal?",
        "ans": "140000 Pa",
        "solution": "Áp suất xe tăng tác dụng lên mặt đất là: p = F / S = P / S = 350 000 / 2,5 = 140 000 (Pa)."
    },
    {
        "num": 3,
        "question": "Một quả cầu bằng kim loại có thể tích 0,002 m^3^ được nhúng chìm hoàn toàn trong nước. Biết trọng lượng riêng của nước là 10000 N/m^3^. Độ lớn lực đẩy Archimedes của nước tác dụng lên quả cầu là bao nhiêu Newton?",
        "ans": "20 N",
        "solution": "Độ lớn lực đẩy Archimedes tác dụng lên quả cầu là: F~A~ = d × V = 10 000 × 0,002 = 20 (N)."
    }
]

KHTN2_ESSAY = [
    {
        "num": 1,
        "question": "Khối lượng riêng của nhôm là 2700 kg/m^3^.\na) Tính khối lượng của một khối nhôm có thể tích 0,04 m^3^.\nb) Tính thể tích của một khối nhôm có khối lượng 13,5 kg.",
        "solution": "a) Khối lượng của khối nhôm thể tích 0,04 m^3^ là:\nm = D × V = 2700 × 0,04 = 108 (kg).\nb) Thể tích của khối nhôm khối lượng 13,5 kg là:\nV = m / D = 13,5 / 2700 = 0,005 (m^3^) = 5 dm^3^."
    },
    {
        "num": 2,
        "question": "Giải thích tại sao:\na) Móng nhà xây rộng bản hơn tường nhà?\nb) Kim khâu có một đầu rất nhọn, còn dao cắt giấy cần phải mài sắc?\nc) Khi lặn sâu dưới nước, thợ lặn phải mặc áo lặn chuyên dụng chịu được áp suất cao?",
        "solution": "a) Móng nhà xây rộng bản để làm tăng diện tích bị ép S với mặt đất. Theo công thức p = F / S, diện tích S lớn sẽ làm giảm áp suất p do toàn bộ ngôi nhà tác dụng lên mặt đất, giúp nhà không bị lún, sụt, nứt nền.\nb) Đầu kim khâu rất nhọn và dao mài sắc nhằm làm giảm tối đa diện tích tiếp xúc S. Khi có lực tác dụng, áp suất p = F / S tạo ra rất lớn, giúp mũi kim dễ dàng đâm xuyên qua vải và dao dễ dàng cắt đứt giấy.\nc) Càng lặn sâu dưới nước, độ sâu h càng tăng làm áp suất chất lỏng p = d × h tác dụng lên cơ thể thợ lặn càng lớn. Nếu không mặc áo lặn chuyên dụng chịu áp suất cao, áp suất nước rất lớn sẽ ép mạnh làm tổn thương màng nhĩ, lồng ngực và các cơ quan nội tạng."
    },
    {
        "num": 3,
        "question": "Một khối nhôm hình hộp chữ nhật có diện tích đáy là 0,02 m^2^, chiều cao 0,1 m được nhúng chìm hoàn toàn trong nước. Biết trọng lượng riêng của nước là 10000 N/m^3^.\na) Tính thể tích của khối nhôm.\nb) Tính lực đẩy Archimedes của nước tác dụng lên khối nhôm.",
        "solution": "a) Thể tích của khối nhôm là:\nV = S~đáy~ × h = 0,02 × 0,1 = 0,002 (m^3^).\nb) Lực đẩy Archimedes của nước tác dụng lên khối nhôm là:\nF~A~ = d × V = 10 000 × 0,002 = 20 (N)."
    },
    {
        "num": 4,
        "question": "Một vật nặng có trọng lượng P = 40 N. Khi nhúng chìm hoàn toàn vật đó vào trong nước, lực kế chỉ giá trị 25 N.\na) Tính độ lớn lực đẩy Archimedes tác dụng lên vật.\nb) Biết trọng lượng riêng của nước là 10000 N/m^3^, tính thể tích của phần nước bị vật chiếm chỗ.",
        "solution": "a) Lực đẩy Archimedes tác dụng lên vật bằng độ giảm số chỉ của lực kế:\nF~A~ = P - P' = 40 - 25 = 15 (N).\nb) Thể tích phần nước bị vật chiếm chỗ là:\nV = F~A~ / d = 15 / 10 000 = 0,0015 (m^3^) = 1,5 dm^3^ (hoặc 1500 cm^3^)."
    }
]

# ==============================================================================
# PHÂN MÔN KHTN 3 (SINH HỌC)
# ==============================================================================

KHTN3_MCQ = [
    {
        "num": 1,
        "question": "Khi bạn chạy bộ, những hệ cơ quan nào phối hợp hoạt động?",
        "options": [
            "A. Hệ vận động, hệ hô hấp, hệ tuần hoàn",
            "B. Hệ tiêu hóa, hệ sinh dục, hệ bài tiết",
            "C. Hệ thần kinh, hệ sinh dục, hệ tiêu hóa",
            "D. Hệ hô hấp, hệ bài tiết, hệ da"
        ],
        "correct": "A",
        "solution": "Khi chạy bộ, cơ quan vận động hoạt động mạnh; để cung cấp đủ oxy và năng lượng cho cơ co bóp, tim đập nhanh hơn (hệ tuần hoàn) và nhịp thở tăng lên (hệ hô hấp) phối hợp nhịp nhàng."
    },
    {
        "num": 2,
        "question": "Hệ thần kinh có vai trò gì trong hoạt động của cơ thể?",
        "options": [
            "A. Tiêu hóa thức ăn",
            "B. Điều khiển và điều hòa hoạt động của các cơ quan",
            "C. Lọc máu và thải chất độc",
            "D. Vận chuyển máu trong cơ thể"
        ],
        "correct": "B",
        "solution": "Hệ thần kinh có chức năng tiếp nhận kích thích, điều khiển, điều hoà và phối hợp mọi hoạt động của các cơ quan, hệ cơ quan trong cơ thể thành một khối thống nhất."
    },
    {
        "num": 3,
        "question": "Vì sao nói \"cơ thể người là một khối thống nhất\"?",
        "options": [
            "A. Vì các hệ cơ quan giống nhau ở mọi người",
            "B. Vì các cơ quan hoạt động đồng thời theo bản năng",
            "C. Vì các hệ cơ quan liên kết, phối hợp chặt chẽ để duy trì sự sống",
            "D. Vì cơ thể có nhiều bộ phận"
        ],
        "correct": "C",
        "solution": "Cơ thể người là một khối thống nhất vì mọi cơ quan, hệ cơ quan đều có sự liên kết mật thiết, phối hợp nhịp nhàng dưới sự điều khiển của hệ thần kinh và nội tiết nhằm duy trì sự sống và thích nghi môi trường."
    },
    {
        "num": 4,
        "question": "Hệ vận động ở người gồm những thành phần chính nào?",
        "options": [
            "A. Xương, tim, phổi",
            "B. Cơ, khớp, xương",
            "C. Xương, cơ, hệ tuần hoàn",
            "D. Cơ, khớp, hệ thần kinh"
        ],
        "correct": "B",
        "solution": "Hệ vận động ở người gồm 3 thành phần chính: bộ xương, khớp xương và hệ cơ (cơ xương)."
    },
    {
        "num": 5,
        "question": "Chức năng chính của xương trong hệ vận động là gì?",
        "options": [
            "A. Tạo ra máu và oxy",
            "B. Nâng đỡ cơ thể và bảo vệ các cơ quan",
            "C. Co giãn giúp cơ thể vận động",
            "D. Lọc chất độc trong máu"
        ],
        "correct": "B",
        "solution": "Bộ xương tạo khung nâng đỡ cơ thể, tạo hình dáng, là chỗ bám của cơ và bảo vệ các nội quan quan trọng (hộp sọ bảo vệ não, lồng ngực bảo vệ tim và phổi...)."
    },
    {
        "num": 6,
        "question": "Khớp xương có vai trò gì trong hệ vận động?",
        "options": [
            "A. Nối xương và giúp chuyển động linh hoạt",
            "B. Bảo vệ não bộ",
            "C. Tiêu hóa thức ăn",
            "D. Vận chuyển máu"
        ],
        "correct": "A",
        "solution": "Khớp xương là nơi tiếp giáp giữa các đầu xương, giúp kết nối các xương thành bộ khung hoàn chỉnh và tạo điều kiện cho các xương cử động linh hoạt."
    },
    {
        "num": 7,
        "question": "Dinh dưỡng là quá trình gì?",
        "options": [
            "A. Hấp thụ nước",
            "B. Nhận và sử dụng chất dinh dưỡng để duy trì sự sống",
            "C. Hấp thụ oxy",
            "D. Tiết mồ hôi"
        ],
        "correct": "B",
        "solution": "Dinh dưỡng là quá trình cơ thể thu nhận, biến đổi và sử dụng các chất dinh dưỡng từ thức ăn nhằm cung cấp năng lượng và nguyên liệu xây dựng tế bào, duy trì sự sống."
    },
    {
        "num": 8,
        "question": "Chức năng chính của hệ tiêu hóa là gì?",
        "options": [
            "A. Tiết ra hormone",
            "B. Vận chuyển máu",
            "C. Biến đổi thức ăn thành chất nhỏ để cơ thể hấp thụ",
            "D. Điều khiển hoạt động cơ thể"
        ],
        "correct": "C",
        "solution": "Hệ tiêu hoá có chức năng tiếp nhận, biến đổi thức ăn thành các chất dinh dưỡng đơn giản mà cơ thể có thể hấp thụ qua thành ruột non vào máu, đồng thời đào thải các chất cặn bã ra ngoài."
    },
    {
        "num": 9,
        "question": "Cơ quan nào trong hệ tiêu hóa có chức năng tiêu hóa thức ăn bằng dịch vị?",
        "options": [
            "A. Miệng",
            "B. Dạ dày",
            "C. Thực quản",
            "D. Ruột già"
        ],
        "correct": "B",
        "solution": "Dạ dày tiết dịch vị (chứa HCl và enzyme pepsin) giúp biến đổi một phần protein trong thức ăn và co bóp nhào trộn thức ăn thành dịch dưỡng trấp."
    },
    {
        "num": 10,
        "question": "Ruột già có chức năng gì trong hệ tiêu hóa?",
        "options": [
            "A. Tiêu hóa thức ăn",
            "B. Hấp thụ chất dinh dưỡng",
            "C. Hấp thụ nước và hình thành phân",
            "D. Nhai thức ăn"
        ],
        "correct": "C",
        "solution": "Ruột già có chức năng chủ yếu là tái hấp thụ nước và một số muối khoáng từ bã thức ăn, đồng thời cô đặc và hình thành phân để thải ra ngoài."
    }
]

KHTN3_TF = [
    {
        "num": 1,
        "text": "Hệ vận động gồm xương, cơ và khớp giúp cơ thể chuyển động linh hoạt.",
        "ans": "Đúng",
        "explain": "Xương đóng vai trò làm đòn bẩy, khớp làm điểm tựa và cơ co duỗi tạo lực kéo giúp cơ thể chuyển động nhịp nhàng, linh hoạt."
    },
    {
        "num": 2,
        "text": "Xương chỉ có chức năng bảo vệ cơ thể, không tham gia vào việc vận động.",
        "ans": "Sai",
        "explain": "Ngoài chức năng bảo vệ cơ thể, xương còn làm khung nâng đỡ và phối hợp với cơ bắp, khớp để thực hiện các hoạt động vận động."
    },
    {
        "num": 3,
        "text": "Khớp xương giúp các xương nối với nhau và làm cho cơ thể cử động dễ dàng hơn.",
        "ans": "Đúng",
        "explain": "Khớp xương kết nối các đầu xương và cho phép các xương cử động trượt qua nhau êm ái nhờ có sụn và dịch khớp."
    }
]

KHTN3_SHORT = [
    {
        "num": 1,
        "question": "Hệ vận động gồm những thành phần nào?",
        "ans": "Xương, cơ, khớp.",
        "solution": "Hệ vận động ở người gồm 3 thành phần chính: bộ xương, hệ cơ (cơ vân) và các khớp xương."
    },
    {
        "num": 2,
        "question": "Xương có chức năng gì trong cơ thể?",
        "ans": "Nâng đỡ, bảo vệ.",
        "solution": "Xương làm khung nâng đỡ cơ thể, bảo vệ các nội quan quan trọng (hộp sọ bảo vệ não, lồng ngực bảo vệ tim phổi...), đồng thời tham gia vận động và sinh máu."
    },
    {
        "num": 3,
        "question": "Khớp là gì?",
        "ans": "Nơi nối xương.",
        "solution": "Khớp xương là nơi tiếp giáp giữa hai hay nhiều đầu xương với nhau."
    },
    {
        "num": 4,
        "question": "Cơ có vai trò gì trong hệ vận động?",
        "ans": "Co giãn tạo chuyển động.",
        "solution": "Các bắp cơ bám vào xương, khi co và dãn sẽ kéo các xương chuyển động quanh khớp, tạo nên mọi cử động của cơ thể."
    },
    {
        "num": 5,
        "question": "Bộ xương người có khoảng bao nhiêu xương?",
        "ans": "Khoảng 206 xương.",
        "solution": "Ở người trưởng thành, bộ xương gồm khoảng 206 chiếc xương được gắn kết với nhau."
    },
    {
        "num": 6,
        "question": "Tại sao việc tập luyện thể dục thể thao lại giúp hệ vận động khỏe mạnh?",
        "ans": "Tập luyện giúp cơ phát triển, xương chắc khỏe và khớp linh hoạt.",
        "solution": "Luyện tập thể dục thể thao kích thích tuần hoàn máu nuôi cơ xương, làm tăng khối lượng và sức bền cơ bắp, tăng mật độ và độ chắc khỏe của xương, đồng thời tăng tiết dịch khớp giúp khớp linh hoạt, dẻo dai."
    },
    {
        "num": 7,
        "question": "Nếu một người bị gãy xương, tại sao cần cố định phần xương bị gãy?",
        "ans": "Để xương lành lại đúng vị trí, tránh di lệch và tổn thương thêm.",
        "solution": "Cố định phần xương gãy giúp các đầu xương không bị di lệch khỏi vị trí giải phẫu, tránh làm tổn thương thêm mạch máu, dây thần kinh và mô xung quanh, giảm đau và giúp xương nhanh lành thẳng đẹp."
    },
    {
        "num": 8,
        "question": "Vì sao ruột non lại quan trọng trong quá trình tiêu hóa?",
        "ans": "Hấp thụ chất dinh dưỡng vào máu.",
        "solution": "Ruột non có bề mặt hấp thụ cực lớn nhờ các nếp gấp, lông ruột và mạng lưới mao mạch dày đặc; đây là nơi diễn ra tiêu hoá hoá học hoàn chỉnh nhất và hấp thụ hầu hết chất dinh dưỡng vào máu nuôi cơ thể."
    },
    {
        "num": 9,
        "question": "Nước có vai trò gì trong dinh dưỡng?",
        "ans": "Tham gia phản ứng sinh học và hòa tan các chất.",
        "solution": "Nước là dung môi hoà tan các chất dinh dưỡng, là môi trường cho mọi phản ứng sinh hoá trong tế bào, tham gia vận chuyển chất và điều hoà thân nhiệt."
    }
]

KHTN3_ESSAY = [
    {
        "num": 1,
        "question": "Giải thích tại sao việc ăn uống đủ chất và cân đối các nhóm dinh dưỡng lại quan trọng đối với sức khỏe con người?",
        "solution": "- Ăn uống đủ chất và cân đối (chất bột đường, chất đạm, chất béo, vitamin, khoáng chất, chất xơ và nước) cung cấp đầy đủ năng lượng cho các hoạt động sống hàng ngày.\n- Cung cấp nguyên liệu để cấu tạo nên tế bào, mô, giúp cơ thể sinh trưởng, phát triển bình thường và hồi phục, sửa chữa các tế bào bị tổn thương.\n- Giúp tăng cường hệ miễn dịch và sức đề kháng chống lại bệnh tật.\n- Nếu thiếu chất hoặc mất cân đối: Cơ thể sẽ bị suy dinh dưỡng, mệt mỏi, suy giảm thể lực và trí lực, hoặc mắc các bệnh rối loạn chuyển hóa (béo phì, tim mạch, tiểu đường...)."
    },
    {
        "num": 2,
        "question": "Mô tả quá trình tiêu hóa thức ăn từ khi đưa vào miệng đến khi được hấp thụ ở ruột non.",
        "solution": "- Tại miệng: Thức ăn được nhai, nghiền nhỏ bằng răng và trộn đều với nước bọt; enzyme amylase trong nước bọt biến đổi một phần tinh bột chín thành đường maltose; thức ăn được tạo thành viên nuốt xuống thực quản.\n- Tại thực quản: Nhu động của cơ thực quản co bóp đẩy viên thức ăn xuống dạ dày.\n- Tại dạ dày: Dạ dày co bóp nhào trộn thức ăn với dịch vị; acid HCl và enzyme pepsin phân giải một phần protein thành các chuỗi peptide ngắn; thức ăn biến đổi thành dịch dưỡng trấp rồi đi vào tá tràng.\n- Tại ruột non: Dưỡng trấp được tiêu hoá hoàn toàn nhờ tác dụng của dịch mật, dịch tụy và dịch ruột; các chất dinh dưỡng (đường đơn, amino acid, acid béo, glycerol) được phân giải triệt để và hấp thụ qua niêm mạc ruột non vào máu và bạch huyết đi nuôi cơ thể."
    },
    {
        "num": 3,
        "question": "Nếu một người bị rối loạn tiêu hóa ở ruột non, hậu quả sẽ ảnh hưởng thế nào đến cơ thể?",
        "solution": "- Ruột non là cơ quan tiêu hoá hoá học chính và hấp thụ phần lớn chất dinh dưỡng của cơ thể.\n- Khi bị rối loạn tiêu hoá ở ruột non, quá trình phân giải thức ăn bị đình trệ, ruột non không hấp thụ được chất dinh dưỡng và nước, dẫn đến tình trạng tiêu chảy, mất nước và sụt cân nhanh chóng.\n- Hậu quả lâu dài: Cơ thể bị thiếu hụt dinh dưỡng trầm trọng dẫn đến suy dinh dưỡng, suy nhược cơ thể, thiếu máu, mệt mỏi và sức đề kháng kém, dễ mắc các bệnh nhiễm trùng."
    },
    {
        "num": 4,
        "question": "Vì sao cần uống đủ nước hàng ngày trong quá trình tiêu hóa và dinh dưỡng?",
        "solution": "- Nước làm mềm và hoá lỏng thức ăn, hỗ trợ việc nhai, nuốt và sự di chuyển của thức ăn qua đường tiêu hoá.\n- Nước là dung môi hoà tan các chất dinh dưỡng, đồng thời là môi trường cần thiết để các enzyme tiêu hoá xúc tác cho các phản ứng thuỷ phân thức ăn.\n- Nước hỗ trợ hấp thụ các chất dinh dưỡng qua thành ruột vào máu và vận chuyển đến khắp các tế bào trong cơ thể.\n- Nước giúp ruột hoạt động trơn tru, phòng chống táo bón và hỗ trợ cơ thể đào thải các chất cặn bã, độc tố qua đường bài tiết."
    }
]
