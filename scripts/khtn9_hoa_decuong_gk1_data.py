# -*- coding: utf-8 -*-
"""
Dữ liệu Đề cương Ôn tập Giữa học kì I - KHTN 9 (Phân môn Hóa học) - Năm học 2026 - 2027
Chủ đề:
- Bài 22: Giới thiệu về hợp chất hữu cơ
- Bài 23: Alkane
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

INFO = {
    "school": "TRƯỜNG TRUNG HỌC CƠ SỞ",
    "year": "NĂM HỌC 2026 - 2027",
    "title": "ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA HỌC KÌ I",
    "subject": "MÔN: KHOA HỌC TỰ NHIÊN 9 (PHÂN MÔN HÓA HỌC)",
    "subtitle": "(Trọng tâm kiến thức: Bài 22. Hợp chất hữu cơ & Bài 23. Alkane)",
    "time_str": "Thời gian làm bài: 45 phút (Không kể thời gian phát đề)"
}

# ==============================================================================
# PHẦN A: LÝ THUYẾT TRỌNG TÂM
# ==============================================================================

THEORY_SECTIONS = [
    {
        "title": "BÀI 22. GIỚI THIỆU VỀ HỢP CHẤT HỮU CƠ",
        "items": [
            "1. Khái niệm hợp chất hữu cơ và hóa học hữu cơ:",
            "- Hợp chất hữu cơ là hợp chất của carbon (trừ một số hợp chất vô cơ đơn giản như CO, CO~2~, H~2~CO~3~, các muối carbonate như CaCO~3~, NaHCO~3~, Na~2~CO~3~, các muối cyanide, carbide...).",
            "- Hóa học hữu cơ là ngành hóa học chuyên nghiên cứu về các hợp chất hữu cơ và các quá trình biến đổi của chúng.",
            "2. Phân loại hợp chất hữu cơ:",
            "Dựa vào thành phần nguyên tố, hợp chất hữu cơ được chia thành 2 loại chính:",
            "• Hydrocarbon: Phân tử chỉ chứa hai nguyên tố là Carbon (C) và Hydrogen (H). Ví dụ: CH~4~, C~2~H~6~, C~3~H~8~, C~4~H~10~, C~2~H~4~, C~2~H~2~, C~6~H~6~...",
            "• Dẫn xuất của hydrocarbon: Phân tử ngoài C và H còn chứa các nguyên tố khác như O, N, Cl, Br, S, Na... Ví dụ: C~2~H~6~O, C~2~H~5~Br, CH~3~Cl, CH~3~NO~2~, CH~3~COOH, C~2~H~3~O~2~Na...",
            "3. Đặc điểm cấu tạo phân tử hợp chất hữu cơ:",
            "- Hóa trị của các nguyên tố: Trong hợp chất hữu cơ, nguyên tử Carbon luôn có hóa trị IV, Hydrogen có hóa trị I, Oxygen có hóa trị II, Halogen (Cl, Br) có hóa trị I.",
            "- Mạch carbon: Các nguyên tử carbon có thể liên kết trực tiếp với nhau tạo thành các dạng mạch carbon phong phú:",
            "  + Mạch hở không phân nhánh (mạch thẳng).",
            "  + Mạch hở phân nhánh.",
            "  + Mạch vòng.",
            "- Trật tự liên kết: Mỗi hợp chất hữu cơ có một trật tự liên kết xác định giữa các nguyên tử trong phân tử. Nếu trật tự liên kết thay đổi sẽ tạo ra chất mới hoàn toàn (hiện tượng đồng phân).",
            "4. Công thức phân tử và công thức cấu tạo:",
            "- Công thức phân tử (CTPT): Cho biết thành phần các nguyên tố và số lượng nguyên tử của mỗi nguyên tố trong phân tử (Ví dụ: C~2~H~6~O, C~4~H~10~).",
            "- Công thức cấu tạo (CTCT): Biểu diễn trật tự và cách thức liên kết giữa các nguyên tử trong phân tử. CTCT có thể viết ở dạng khai triển đầy đủ hoặc dạng thu gọn (Ví dụ: CH~3~-CH~2~-OH và CH~3~-O-CH~3~ có cùng CTPT C~2~H~6~O nhưng có CTCT khác nhau)."
        ]
    },
    {
        "title": "BÀI 23. ALKANE",
        "items": [
            "1. Khái niệm và công thức chung của alkane:",
            "- Alkane là các hydrocarbon mạch hở, trong phân tử chỉ chứa các liên kết đơn (liên kết C-C và C-H).",
            "- Công thức chung của alkane: C~n~H~2n+2~ (với n ≥ 1, n là số nguyên dương).",
            "- Dãy đồng đẳng alkane: CH~4~ (methane), C~2~H~6~ (ethane), C~3~H~8~ (propane), C~4~H~10~ (butane), C~5~H~12~ (pentane)...",
            "2. Tính chất vật lí:",
            "- Ở điều kiện thường: Các alkane từ C1 đến C4 (CH~4~, C~2~H~6~, C~3~H~8~, C~4~H~10~) là chất khí; từ C5 đến khoảng C17 là chất lỏng; từ C18 trở lên là chất rắn.",
            "- Các alkane đều không màu, nhẹ hơn nước, hầu như không tan trong nước nhưng tan tốt trong các dung môi hữu cơ (xăng, benzen...).",
            "3. Tính chất hóa học:",
            "a) Phản ứng cháy (phản ứng oxi hóa hoàn toàn):",
            "Khi bị đốt cháy trong không khí hoặc oxygen, các alkane cháy tỏa nhiều nhiệt, tạo ra khí carbon dioxide và hơi nước:",
            "  C~n~H~2n+2~ + (3n + 1)/2 O~2~ --(t°)-> n CO~2~ + (n + 1) H~2~O",
            "  * Nhận xét: Đốt cháy alkane luôn thu được n(H~2~O) > n(CO~2~).",
            "b) Phản ứng thế halogen (phản ứng đặc trưng cho liên kết đơn):",
            "Dưới tác dụng của ánh sáng, nguyên tử hydrogen trong phân tử alkane có thể bị thay thế lần lượt bởi nguyên tử halogen (Cl, Br):",
            "  CH~4~ + Cl~2~ --(ánh sáng)-> CH~3~Cl + HCl",
            "4. Ứng dụng của alkane:",
            "- Làm nhiên liệu chủ lực: Khí đốt tự nhiên, khí sinh học biogas (chủ yếu CH~4~); khí gas đun nấu hóa lỏng LPG (hỗn hợp propane C~3~H~8~ và butane C~4~H~10~); nhiên liệu cho động cơ đốt trong (xăng, dầu diesel, dầu hỏa).",
            "- Làm nguyên liệu công nghiệp: Sản xuất hydrogen, bột than (carbon black) dùng làm mực in, xi đánh giày, gia cường cao su lốp xe; sản xuất dung môi hữu cơ, sáp nến (paraffin), dầu mỡ nhờn bôi trơn máy móc; nguyên liệu tổng hợp hóa chất hữu cơ."
        ]
    }
]

# ==============================================================================
# PHẦN B: BÀI TẬP TỰ LUẬN
# ==============================================================================

ESSAY_QUESTIONS = [
    {
        "id": 1,
        "content": "Hãy sắp xếp các chất sau: C~6~H~6~, CaCO~3~, C~4~H~10~, C~2~H~6~O, NaNO~3~, CH~3~NO~2~, NaHCO~3~, C~2~H~3~O~2~Na vào các cột thích hợp trong bảng phân loại chất dưới đây:",
        "table_header": ["Hợp chất hữu cơ (Hydrocarbon)", "Hợp chất hữu cơ (Dẫn xuất của Hydrocarbon)", "Hợp chất vô cơ"],
        "student_table": [
            ["", "", ""]
        ],
        "teacher_table": [
            ["C~6~H~6~, C~4~H~10~", "C~2~H~6~O, CH~3~NO~2~, C~2~H~3~O~2~Na", "CaCO~3~, NaNO~3~, NaHCO~3~"]
        ],
        "solution": (
            "Phân tích bản chất từng chất:\n"
            "• Hợp chất hữu cơ:\n"
            "  - Hydrocarbon (chỉ cấu tạo từ 2 nguyên tố C và H): C~6~H~6~ (benzene), C~4~H~10~ (butane).\n"
            "  - Dẫn xuất của hydrocarbon (ngoài C và H còn chứa O, N, kim loại Na): C~2~H~6~O (alcohol/ether), CH~3~NO~2~ (nitromethane), C~2~H~3~O~2~Na (sodium acetate CH~3~COONa).\n"
            "• Hợp chất vô cơ:\n"
            "  - NaNO~3~ (không chứa carbon, là muối nitrat vô cơ).\n"
            "  - CaCO~3~ (muối carbonate vô cơ của kim loại calcium).\n"
            "  - NaHCO~3~ (muối hydrogencarbonate vô cơ của kim loại sodium)."
        )
    },
    {
        "id": 2,
        "content": "Hãy viết công thức cấu tạo đầy đủ (dạng khai triển) và thu gọn có thể có của các chất có công thức phân tử sau: C~2~H~5~Br, C~4~H~10~, C~3~H~6~, C~3~H~8~O.",
        "solution": (
            "1. C~2~H~5~Br (Bromoethane - có 1 CTCT duy nhất):\n"
            "   • CTCT đầy đủ: H-C(H)(H)-C(H)(H)-Br (2 nguyên tử C liên kết đơn, 1 C liên kết với 1 Br và 2 H, C còn lại liên kết với 3 H).\n"
            "   • CTCT thu gọn: CH~3~-CH~2~Br.\n\n"
            "2. C~4~H~10~ (Butane - có 2 đồng phân cấu tạo):\n"
            "   • Đồng phân 1 (mạch không phân nhánh - butane):\n"
            "     - CTCT đầy đủ: H-C(H)(H)-C(H)(H)-C(H)(H)-C(H)(H)-H (4 nguyên tử C liên kết đơn thành chuỗi thẳng).\n"
            "     - CTCT thu gọn: CH~3~-CH~2~-CH~2~-CH~3~.\n"
            "   • Đồng phân 2 (mạch phân nhánh - isobutane / 2-methylpropane):\n"
            "     - CTCT đầy đủ: 1 nguyên tử C trung tâm liên kết với 1 H và 3 nhóm CH~3~.\n"
            "     - CTCT thu gọn: CH~3~-CH(CH~3~)-CH~3~.\n\n"
            "3. C~3~H~6~ (có 2 đồng phân cấu tạo: 1 mạch hở có liên kết đôi và 1 mạch vòng):\n"
            "   • Đồng phân 1 (propene - anken mạch hở):\n"
            "     - CTCT đầy đủ: C(H)(H)=C(H)-C(H)(H)(H) (có 1 liên kết đôi C=C).\n"
            "     - CTCT thu gọn: CH~2~=CH-CH~3~.\n"
            "   • Đồng phân 2 (cyclopropane - xicloankan mạch vòng 3 cạnh):\n"
            "     - CTCT đầy đủ: 3 nhóm -CH~2~- liên kết vòng khép kín hình tam giác.\n"
            "     - CTCT thu gọn: Vòng tam giác (CH~2~)~3~ hoặc cyclo-C~3~H~6~.\n\n"
            "4. C~3~H~8~O (có 3 đồng phân cấu tạo: 2 alcohol và 1 ether):\n"
            "   • Đồng phân 1 (propan-1-ol / alcohol bậc I):\n"
            "     - CTCT thu gọn: CH~3~-CH~2~-CH~2~-OH.\n"
            "   • Đồng phân 2 (propan-2-ol / alcohol bậc II):\n"
            "     - CTCT thu gọn: CH~3~-CH(OH)-CH~3~.\n"
            "   • Đồng phân 3 (ethyl methyl ether):\n"
            "     - CTCT thu gọn: CH~3~-O-CH~2~-CH~3~."
        )
    },
    {
        "id": 3,
        "content": "Gas dùng đun nấu có chứa alkane chủ yếu nào? Viết phương trình hoá học biểu diễn phản ứng cháy của chúng.",
        "solution": (
            "• Khí gas dân dụng dùng trong đun nấu gia đình (khí dầu mỏ hóa lỏng - LPG: Liquefied Petroleum Gas) chứa thành phần chủ yếu là 2 alkane: propane (C~3~H~8~) và butane (C~4~H~10~) được nén hóa lỏng.\n"
            "• Phương trình hoá học biểu diễn phản ứng cháy:\n"
            "  (1) Phản ứng cháy của propane:\n"
            "      C~3~H~8~ + 5 O~2~ --(t°)-> 3 CO~2~ + 4 H~2~O\n"
            "  (2) Phản ứng cháy của butane:\n"
            "      2 C~4~H~10~ + 13 O~2~ --(t°)-> 8 CO~2~ + 10 H~2~O\n"
            "      (hoặc: C~4~H~10~ + 13/2 O~2~ --(t°)-> 4 CO~2~ + 5 H~2~O)"
        )
    },
    {
        "id": 4,
        "content": "Alkane là gì? Viết công thức chung của alkane? Nêu ứng dụng của alkane trong đời sống và sản xuất.",
        "solution": (
            "• Khái niệm: Alkane là các hydrocarbon mạch hở, trong phân tử chỉ chứa các liên kết đơn (liên kết C-C và C-H).\n"
            "• Công thức chung của alkane: C~n~H~2n+2~ (với n ≥ 1, n nguyên dương).\n"
            "• Ứng dụng của alkane:\n"
            "  1. Làm nhiên liệu: Do khả năng cháy tỏa nhiều nhiệt và sinh ra ít chất độc hại, các alkane được dùng làm khí đốt gia đình (gas hóa lỏng propane, butane), khí tự nhiên/biogas (methane), nhiên liệu cho động cơ xe máy, ô tô, máy bay, tàu thủy (xăng, dầu diesel, dầu hỏa).\n"
            "  2. Làm nguyên liệu hóa chất: Methane dùng để điều chế khí hydrogen, bột than (carbon black) chế tạo mực in, lốp cao su; các alkane lỏng và rắn dùng làm dung môi hòa tan, sản xuất sáp nến (paraffin), dầu mỡ nhờn bôi trơn máy móc và làm nguyên liệu cracking tổng hợp hóa chất hữu cơ."
        )
    },
    {
        "id": 5,
        "content": "Alkane X có khối lượng phân tử 44 amu.\na) Xác định công thức phân tử của X và viết công thức cấu tạo có thể có ứng với CTPT vừa tìm được.\nb) Đốt cháy hoàn toàn X trong không khí rồi dẫn toàn bộ sản phẩm cháy qua dung dịch nước vôi trong dư. Nêu hiện tượng quan sát được và viết các phương trình hoá học xảy ra.",
        "solution": (
            "a) Xác định CTPT và CTCT của X:\n"
            "- Alkane X có công thức chung là C~n~H~2n+2~ (n ≥ 1).\n"
            "- Theo đề bài, khối lượng phân tử của X là:\n"
            "  M_X = 12n + (2n + 2) = 14n + 2 = 44 amu\n"
            "  => 14n = 42 => n = 3.\n"
            "- Vậy công thức phân tử của X là: C~3~H~8~ (propane).\n"
            "- Công thức cấu tạo có thể có của C~3~H~8~: Phân tử C~3~H~8~ chỉ có duy nhất 1 công thức cấu tạo (mạch thẳng không phân nhánh):\n"
            "  + CTCT đầy đủ: H-C(H)(H)-C(H)(H)-C(H)(H)-H.\n"
            "  + CTCT thu gọn: CH~3~-CH~2~-CH~3~.\n\n"
            "b) Hiện tượng và phương trình hoá học:\n"
            "- Khi đốt cháy hoàn toàn X (C~3~H~8~) trong không khí:\n"
            "  C~3~H~8~ + 5 O~2~ --(t°)-> 3 CO~2~ + 4 H~2~O\n"
            "- Dẫn sản phẩm cháy (gồm CO~2~ và H~2~O) qua dung dịch nước vôi trong Ca(OH)~2~ dư:\n"
            "  + Khí CO~2~ phản ứng tạo kết tủa trắng calcium carbonate (CaCO~3~):\n"
            "    CO~2~ + Ca(OH)~2~ -> CaCO~3~↓ + H~2~O\n"
            "- Hiện tượng quan sát được:\n"
            "  Dung dịch nước vôi trong bị vẩn đục, xuất hiện kết tủa màu trắng lắng xuống đáy bình/ống nghiệm."
        )
    }
]

# ==============================================================================
# PHẦN C: CÂU HỎI TRẮC NGHIỆM THAM KHẢO
# ==============================================================================

MCQ_QUESTIONS = [
    {
        "id": 1,
        "content": "Chất nào sau đây không thuộc loại chất hữu cơ?",
        "options": ["A. CH~4~.", "B. CH~3~Cl.", "C. CH~3~COONa.", "D. CO~2~."],
        "answer": "D",
        "solution": "Hợp chất hữu cơ là hợp chất của carbon (trừ CO, CO~2~, H~2~CO~3~, muối carbonate...). Khí carbon dioxide (CO~2~) là hợp chất vô cơ."
    },
    {
        "id": 2,
        "content": "Hãy cho biết công thức cấu tạo dưới đây là của hợp chất có CTPT nào?",
        "image": "khtn9_butane_structure.png",
        "image_width_cm": 9.0,
        "options": ["A. C~4~H~8~.", "B. C~5~H~10~.", "C. C~4~H~10~.", "D. C~6~H~10~."],
        "answer": "C",
        "solution": "Đếm số lượng nguyên tử trong công thức cấu tạo khai triển: có 4 nguyên tử carbon (C) liên kết đơn thành chuỗi thẳng và 10 nguyên tử hydrogen (H) bao quanh. Do đó, công thức phân tử là C~4~H~10~ (butane)."
    },
    {
        "id": 3,
        "content": "Ở điều kiện thường, alkane nào sau đây tồn tại ở thể khí?",
        "options": ["A. Pentane (C~5~H~12~).", "B. Hexane (C~6~H~14~).", "C. Methane (CH~4~).", "D. Octane (C~8~H~18~)."],
        "answer": "C",
        "solution": "Ở điều kiện thường, các alkane từ C1 đến C4 (CH~4~, C~2~H~6~, C~3~H~8~, C~4~H~10~) tồn tại ở thể khí. Từ C5 đến C17 là chất lỏng, từ C18 trở lên là chất rắn."
    },
    {
        "id": 4,
        "content": "Cho nhiệt lượng tỏa ra khi đốt cháy hoàn toàn 1 mol các chất ethane, propane, butane và pentane (C~5~H~12~) lần lượt là 1 570 kJ/mol; 2 220 kJ/mol; 2 875 kJ/mol và 3 536 kJ/mol. Khi đốt cháy hoàn toàn 1 gam chất nào sẽ thu được nhiệt lượng lớn nhất?",
        "options": ["A. Ethane.", "B. Propane.", "C. Pentane.", "D. Butane."],
        "answer": "A",
        "solution": (
            "Khối lượng mol (M) của các chất:\n"
            "• Ethane (C~2~H~6~): M = 30 g/mol => Nhiệt lượng/1g: 1570 / 30 ≈ 52,33 kJ/g.\n"
            "• Propane (C~3~H~8~): M = 44 g/mol => Nhiệt lượng/1g: 2220 / 44 ≈ 50,45 kJ/g.\n"
            "• Butane (C~4~H~10~): M = 58 g/mol => Nhiệt lượng/1g: 2875 / 58 ≈ 49,57 kJ/g.\n"
            "• Pentane (C~5~H~12~): M = 72 g/mol => Nhiệt lượng/1g: 3536 / 72 ≈ 49,11 kJ/g.\n"
            "So sánh: 52,33 > 50,45 > 49,57 > 49,11 kJ/g.\n"
            "=> Khi đốt cháy hoàn toàn 1 gam ethane thu được nhiệt lượng lớn nhất (≈ 52,33 kJ)."
        )
    },
    {
        "id": 5,
        "content": "Trong số các phát biểu sau đây, số phát biểu đúng là:\n(1) Mỗi CTPT có thể biểu diễn cho nhiều chất hữu cơ.\n(2) Hydrocarbon là hợp chất mà thành phần phân tử chỉ chứa các nguyên tố carbon, hydrogen.\n(3) Mỗi hợp chất hữu cơ có thể có nhiều CTPT.\n(4) Mỗi hợp chất hữu cơ có thể có nhiều CTCT.",
        "options": ["A. 2.", "B. 3.", "C. 4.", "D. 1."],
        "answer": "A",
        "solution": (
            "Phân tích từng phát biểu:\n"
            "• (1) Đúng: Hiện tượng đồng phân (ví dụ C~2~H~6~O biểu diễn cho cả ethyl alcohol và dimethyl ether; C~4~H~10~ biểu diễn cho cả butane và isobutane).\n"
            "• (2) Đúng: Theo định nghĩa, hydrocarbon chỉ chứa 2 nguyên tố C và H.\n"
            "• (3) Sai: Mỗi hợp chất hữu cơ xác định chỉ có duy nhất 1 công thức phân tử biểu diễn thành phần nguyên tố.\n"
            "• (4) Sai: Mỗi hợp chất hữu cơ xác định chỉ có một cấu trúc trật tự liên kết xác định (tức là chỉ có một CTCT đặc trưng).\n"
            "=> Có 2 phát biểu đúng là (1) và (2)."
        )
    }
]

# ==============================================================================
# PHẦN TRẮC NGHIỆM ĐÚNG - SAI
# ==============================================================================

TF_QUESTIONS = [
    {
        "id": 1,
        "context": "Công thức phân tử cho biết thành phần nguyên tố và số lượng nguyên tử của mỗi nguyên tố trong hợp chất hữu cơ.",
        "items": [
            {
                "id": "a",
                "content": "Công thức phân tử của ethane là CH₄.",
                "answer": False,
                "reason": "Sai. Công thức phân tử của ethane là C₂H₆. Còn CH₄ là công thức phân tử của methane."
            },
            {
                "id": "b",
                "content": "Công thức phân tử cho biết cách các nguyên tử liên kết với nhau trong phân tử.",
                "answer": False,
                "reason": "Sai. Công thức phân tử chỉ cho biết thành phần nguyên tố và số lượng nguyên tử mỗi nguyên tố, không cho biết trật tự và cách thức liên kết (đó là vai trò của công thức cấu tạo)."
            },
            {
                "id": "c",
                "content": "Công thức cấu tạo biểu diễn liên kết giữa các nguyên tử trong phân tử.",
                "answer": True,
                "reason": "Đúng. Công thức cấu tạo biểu diễn trật tự và bản chất các liên kết hóa học (liên kết đơn, liên kết bội) giữa các nguyên tử trong phân tử."
            },
            {
                "id": "d",
                "content": "Để hiểu rõ cấu trúc của một hợp chất hữu cơ, cần kết hợp cả công thức phân tử và công thức cấu tạo.",
                "answer": True,
                "reason": "Đúng. CTPT cho biết số lượng nguyên tử cấu thành phân tử, còn CTCT cho biết hình dạng không gian và cách thức liên kết; việc kết hợp cả hai giúp hiểu đầy đủ và chính xác bản chất hóa học của chất."
            }
        ]
    }
]
