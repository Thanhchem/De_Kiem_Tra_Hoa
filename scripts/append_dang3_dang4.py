# -*- coding: utf-8 -*-
"""
Append Dạng 3 and Dạng 4 to hoa12_to36_data.py
"""
import os
import sys

DATA_FILE = r"c:\Antigravity_Thanh\scripts\hoa12_to36_data.py"

dang3_and_dang4 = '''
# ==============================================================================
# DẠNG 3: ĐIỆN PHÂN NÓNG CHẢY AL2O3 (QUY TRÌNH HALL - HÉROULT)
# ==============================================================================
THEORY_DANG3 = {
    "title": "DẠNG 3: ĐIỆN PHÂN NÓNG CHẢY Al~2~O~3~ (QUY TRÌNH HALL - HÉROULT)",
    "sections": [
        {
            "heading": "I. Mục đích của quá trình",
            "items": [
                "• Tách kim loại nhôm (Al) từ alumina (Al~2~O~3~) tinh chế từ quặng bauxite.",
                "• Nhôm là kim loại có tính khử mạnh (thế điện cực E°(Al^3+^/Al) = -1,66 V), không thể điều chế bằng phương pháp nhiệt luyện (C, CO, H~2~ không khử được Al~2~O~3~ ở nhiệt độ thường) và không thể điện phân dung dịch nước (nước bị khử trước Al^3+^ sinh H~2~).",
                "• Do đó, bắt buộc phải dùng phương pháp *điện phân nóng chảy*."
            ]
        },
        {
            "heading": "II. Cấu tạo bể điện phân & Vai trò của Cryolite (Na~3~AlF~6~)",
            "image": "p4_img1_xref341_836x221.jpeg",
            "items": [
                "• *Cathode (cực âm):* Lớp lót than chì (graphite) ở đáy và thành bể điện phân.",
                "• *Anode (cực dương):* Các khối than chì (graphite) thả chìm từ trên xuống, có thể hạ thấp dần khi bị ăn mòn.",
                "• *Chất điện li nóng chảy:* Hỗn hợp Al~2~O~3~ hòa tan trong cryolite (Na~3~AlF~6~).",
                "• *3 vai trò then chốt của Cryolite (Na~3~AlF~6~):*",
                "   1. *Hạ nhiệt độ nóng chảy:* Giảm nhiệt độ nóng chảy của Al~2~O~3~ từ ~2050°C xuống khoảng 950 - 1000°C, giúp tiết kiệm năng lượng điện khổng lồ.",
                "   2. *Tăng độ dẫn điện:* Cryolite nóng chảy phân li tạo nhiều ion (Na^+^, [AlF~6~]^3-^, AlF~4~^-^, F^-^), làm tăng mạnh độ dẫn điện của khối nóng chảy.",
                "   3. *Bảo vệ nhôm lỏng:* Hỗn hợp cryolite nóng chảy có khối lượng riêng nhỏ hơn nhôm lỏng (D~cryolite~ ≈ 2,1 g/cm^3^ < D~Al lỏng~ ≈ 2,3 g/cm^3^) nên nổi lên trên bề mặt, ngăn cách Al lỏng ở đáy không bị không khí và O~2~ oxi hóa ngược lại."
            ]
        },
        {
            "heading": "III. Phản ứng hóa học tại các điện cực",
            "items": [
                "• *Tại Cathode (cực âm - đáy bể):* Cation Al^3+^ bị khử thành nhôm lỏng chìm xuống đáy:\\n      Al^3+^ + 3e → Al (lỏng)",
                "• *Tại Anode (cực dương - thanh graphite):* Anion O^2-^ bị oxi hóa thành khí O~2~:\\n      2O^2-^ → O~2~ + 4e",
                "  Khí O~2~ sinh ra ở nhiệt độ cao (~950°C) đốt cháy các khối than chì anode:\\n      C + O~2~ → CO~2~\\n      2C + O~2~ → 2CO",
                "• *Phương trình tổng quát:* 2Al~2~O~3~ → 4Al + 3O~2~\\n  (Khí thoát ra ở anode là hỗn hợp gồm CO~2~, CO và một phần O~2~ dư)."
            ]
        },
        {
            "heading": "IV. Các lưu ý công nghệ quan trọng",
            "table": {
                "headers": ["Yếu tố công nghệ", "Đặc điểm và nguyên tắc vận hành"],
                "rows": [
                    ["Tiêu hao cực Anode", "Cực than chì anode bị oxi hóa liên tục bởi khí O~2~ ở nhiệt độ cao tạo CO và CO~2~, nên cần được hạ thấp dần và định kỳ thay thế."],
                    ["Tháo nhôm định kỳ", "Nhôm lỏng nặng hơn chìm xuống đáy bể, được hút chân không tháo ra định kỳ đưa sang lò đúc thỏi."],
                    ["Điện áp & Dòng điện", "Điện áp mỗi bể rất thấp (khoảng 4 - 5 V) để tránh phân hủy Na^+^, nhưng cường độ dòng điện cực lớn (100 - 300 kA) để đạt sản lượng cao."],
                    ["Bổ sung nguyên liệu", "Al~2~O~3~ và cryolite được nạp bổ sung liên tục để duy trì thành phần tối ưu của bồn điện phân."]
                ]
            }
        }
    ]
}

EXERCISES_DANG3 = [
    {
        "id": "D3_C1",
        "num": 1,
        "type": "mcq",
        "content": "Điện phân nóng chảy hỗn hợp gồm aluminium oxide (Al~2~O~3~) và cryolite (Na~3~AlF~6~) để sản xuất nhôm (Al) trong công nghiệp. Cho các vai trò sau của cryolite:\\n(1) Tạo với Al~2~O~3~ thành hỗn hợp có nhiệt độ nóng chảy thấp hơn nhiều nhiệt độ nóng chảy Al~2~O~3~.\\n(2) Tăng khả năng dẫn điện của hỗn hợp điện phân do sự điện li tạo thành các ion.\\n(3) Hình thành lớp màng ngăn cản sự tiếp xúc giữa Al (lỏng) mới sinh với không khí.\\nHãy cho biết những vai trò nào đúng với cryolite trong quá trình sản xuất nhôm trong công nghiệp.",
        "options": [
            "A. (1), (2) và (3).",
            "B. (1) và (2).",
            "C. (1) và (3).",
            "D. (2) và (3)."
        ],
        "ans": "A",
        "solution": (
            "- Cryolite (Na~3~AlF~6~) có cả 3 vai trò cực kỳ quan trọng:\\n"
            "  (1) Hạ nhiệt độ nóng chảy từ ~2050°C xuống ~950°C (tiết kiệm năng lượng).\\n"
            "  (2) Phân li tạo nhiều ion tự do, làm tăng mạnh độ dẫn điện của khối nóng chảy.\\n"
            "  (3) Khối lượng riêng nhẹ hơn Al lỏng nên nổi lên bề mặt, ngăn cách Al lỏng với O~2~ không khí.\\n"
            "→ Đáp án đúng là *A*."
        )
    },
    {
        "id": "D3_C2",
        "num": 2,
        "type": "tf",
        "content": "Trong công nghiệp, nhôm được sản xuất bằng cách điện phân nóng chảy hỗn hợp aluminium oxide (Al~2~O~3~) và cryolite (Na~3~AlF~6~) còn gọi là quy trình Hall-Héroult. Thùng điện phân có cathode là tấm than chì ở dưới đáy thùng, anode là những tấm than chì có thể di chuyển theo phương thẳng đứng. Cho các nhận định sau:",
        "items": [
            {"id": "a", "text": "Tại anode xảy ra quá trình oxi hóa O^2-^ thành O~2~, tại cathode xảy ra quá trình khử Al^3+^ thành Al.", "ans": "Đ"},
            {"id": "b", "text": "Nhôm lỏng sinh ra có khối lượng riêng lớn hơn hỗn hợp điện li nên chìm xuống đáy và được hỗn hợp điện li bảo vệ khỏi bị oxi hóa.", "ans": "Đ"},
            {"id": "c", "text": "Lượng aluminium oxide bị điện phân với tốc độ càng lớn thì tỉ lệ khí carbonic (CO~2~) trong sản phẩm khí thoát ra ở anode càng lớn.", "ans": "Đ"},
            {"id": "d", "text": "Hỗn hợp Al~2~O~3~ và cryolite có nhiệt độ nóng chảy thấp hơn nhiều so với Al~2~O~3~ nguyên chất.", "ans": "Đ"},
            {"id": "e", "text": "Sau một thời gian phải thay thế cực dương than chì vì O~2~ tạo ra ở nhiệt độ cao đốt cháy C tạo CO và CO~2~.", "ans": "Đ"},
            {"id": "f", "text": "Công thức hoá học của cryolite là Al~2~O~3~.2SiO~2~.2H~2~O.", "ans": "S"},
            {"id": "g", "text": "Nguyên liệu chính trong tự nhiên để sản xuất nhôm là quặng dolomite.", "ans": "S"},
            {"id": "h", "text": "Nhôm nóng chảy được định kì tháo ra ở dưới đáy thùng điện phân.", "ans": "Đ"},
            {"id": "i", "text": "Cryolite được thêm vào bể điện phân giúp tiết kiệm được năng lượng và giảm chi phí sản xuất.", "ans": "Đ"},
            {"id": "j", "text": "Bên cạnh nhôm, oxygen tinh khiết cũng thu được trực tiếp như một sản phẩm thương mại từ quy trình này.", "ans": "S"},
            {"id": "k", "text": "Vì anode và cathode đều làm bằng graphite, nên nếu đổi chiều dòng điện thì quy trình điện phân vẫn diễn ra bình thường.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Anode oxi hóa O^2-^; cathode khử Al^3+^.\\n"
            "- b) *ĐÚNG:* D(Al lỏng) ≈ 2,3 g/cm^3^ > D(hỗn hợp điện li) ≈ 2,1 g/cm^3^ nên Al chìm xuống đáy.\\n"
            "- c) *ĐÚNG:* Tốc độ sinh O~2~ nhanh làm phản ứng cháy C + O~2~ → CO~2~ xảy ra hoàn toàn hơn, tăng tỉ lệ CO~2~.\\n"
            "- d) *ĐÚNG:* Giảm từ 2050°C xuống khoảng 950 - 1000°C.\\n"
            "- e) *ĐÚNG:* Than chì anode bị ăn mòn liên tục.\\n"
            "- f) *SAI:* Cryolite là Na~3~AlF~6~. Công thức Al~2~O~3~.2SiO~2~.2H~2~O là kaolin.\\n"
            "- g) *SAI:* Quặng bauxite (Al~2~O~3~.2H~2~O), còn dolomite là CaCO~3~.MgCO~3~.\\n"
            "- h) *ĐÚNG:* Nhôm lỏng tích tụ ở đáy và được hút ra định kỳ.\\n"
            "- i) *ĐÚNG:* Hạ nhiệt độ từ 2050°C xuống 950°C giúp tiết kiệm năng lượng vượt bậc.\\n"
            "- j) *SAI:* O~2~ sinh ra phản ứng ngay với anode carbon tạo CO~2~ và CO, không thu được O~2~ tự do.\\n"
            "- k) *SAI:* Đáy bể được thiết kế chuyên dụng làm cathode; nếu đổi chiều dòng điện thì đáy bể thành anode sẽ bị O~2~ đốt cháy phá hủy toàn bộ bể."
        )
    },
    {
        "id": "D3_C3",
        "num": 3,
        "type": "tf",
        "content": "Trong công nghiệp, nhôm được sản xuất bằng cách điện phân nóng chảy hỗn hợp alumina (Al~2~O~3~) và cryolite (Na~3~AlF~6~) theo quy trình Hall-Héroult: 2Al~2~O~3~(l) → 4Al(l) + 3O~2~(g). Nhiệt độ nóng chảy của hỗn hợp alumina và cryolite khoảng 950 °C, nhiệt độ nóng chảy của alumina là 2050 °C. Trong quá trình điện phân, cực dương làm bằng graphite bị ăn mòn và liên tục được nhúng xuống bể điện phân. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Nhôm kim loại được tách ra tại điện cực dương (anode).", "ans": "S"},
            {"id": "b", "text": "Khí thoát ra ở anode chủ yếu gồm CO~2~ và CO.", "ans": "Đ"},
            {"id": "c", "text": "Cryolite giúp hạ nhiệt độ nóng chảy của chất điện phân và làm tăng độ dẫn điện.", "ans": "Đ"},
            {"id": "d", "text": "Trong quá trình điện phân phải định kỳ hạ thấp thanh than chì anode.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Nhôm kim loại tách ra ở cathode (cực âm ở đáy bể).\\n"
            "- b) *ĐÚNG:* O~2~ phản ứng với cực than chì tạo hỗn hợp CO~2~ và CO.\\n"
            "- c) *ĐÚNG:* Hạ nhiệt độ từ 2050°C xuống ~950°C và tăng mạnh nồng độ ion tự do.\\n"
            "- d) *ĐÚNG:* Vì đầu thanh than chì tiếp xúc bị đốt cháy mòn dần nên phải hạ thấp liên tục để duy trì tiếp xúc."
        )
    },
    {
        "id": "D3_C4",
        "num": 4,
        "type": "tf",
        "content": "Một nhà máy luyện kim sản xuất nhôm từ 100 tấn quặng bauxite chứa 51% Al~2~O~3~ về khối lượng (phần còn lại là tạp chất không chứa nhôm). Hiệu suất toàn bộ quá trình sản xuất đạt 85%. Cho các nhận định sau:",
        "items": [
            {"id": "a", "text": "Khối lượng Al~2~O~3~ nguyên chất có trong 100 tấn quặng bauxite trên là 51 tấn.", "ans": "Đ"},
            {"id": "b", "text": "Theo lý thuyết (hiệu suất 100%), lượng nhôm thu được từ 51 tấn Al~2~O~3~ là 30 tấn.", "ans": "S"},
            {"id": "c", "text": "Khối lượng nhôm thực tế thu được với hiệu suất 85% xấp xỉ 23 tấn.", "ans": "Đ"},
            {"id": "d", "text": "Năng lượng điện tiêu thụ trong quá trình điện phân phụ thuộc trực tiếp vào khối lượng Al sản xuất và hiệu điện thế áp đặt.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* m(Al~2~O~3~) = 100 × 51% = 51 tấn.\\n"
            "- b) *SAI:* m(Al lý thuyết) = 51 × (54 / 102) = 27 tấn (không phải 30 tấn).\\n"
            "- c) *ĐÚNG:* m(Al thực tế) = 27 × 85% = 22,95 tấn ≈ 23 tấn.\\n"
            "- d) *ĐÚNG:* A = (U × m~Al~ × F) / (9 × 3,6 × 10^6) (kWh)."
        )
    },
    {
        "id": "D3_C5",
        "num": 5,
        "type": "tf",
        "content": "Xét quá trình điện phân nóng chảy Al~2~O~3~ để sản xuất nhôm. Biết ở anode thu được hỗn hợp khí X gồm CO~2~ (80% thể tích) và CO (20% thể tích). Giả sử không có khí nào khác sinh ra. Cho các phát biểu sau:",
        "items": [
            {"id": "a", "text": "Cứ điều chế được 4 mol Al tại cathode thì có 3 mol nguyên tử [O] được giải phóng ở anode.", "ans": "Đ"},
            {"id": "b", "text": "Tỉ lệ số mol giữa CO~2~ và CO trong khí X là 4 : 1.", "ans": "Đ"},
            {"id": "c", "text": "Để sản xuất được 5,4 tấn Al thì khối lượng carbon bị oxi hóa ở anode là 3,6 tấn.", "ans": "S"},
            {"id": "d", "text": "Nếu tăng cường độ dòng điện thì tốc độ giải phóng Al tại cathode sẽ tăng lên.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* 2Al~2~O~3~ → 4Al + 3O~2~ (tương ứng 6 mol O hay 3 mol O~2~ cho 4 mol Al, tỉ lệ n~Al~ : n~O~ = 2 : 3 hay 4 mol Al ứng với 6 mol nguyên tử O, xét n~e~ = 12 mol).\\n"
            "- b) *ĐÚNG:* V(CO~2~) : V(CO) = 80% : 20% = 4 : 1.\\n"
            "- c) *SAI:* n~Al~ = 5,4 × 10^6 / 27 = 0,2 Mmol → n~O~ = 0,3 Mmol.\\n"
            "  Khí X: 4x mol CO~2~ và x mol CO → n~O~ = 2(4x) + x = 9x = 0,3 → x = 1/30 Mmol.\\n"
            "  n~C~ = 5x = 1/6 Mmol → m~C~ = (1/6) × 12 = 2,0 tấn (không phải 3,6 tấn).\\n"
            "- d) *ĐÚNG:* Theo định luật Faraday, m = (A.I.t)/(n.F) tỉ lệ thuận với I."
        )
    },
    {
        "id": "D3_C6",
        "num": 6,
        "type": "tf",
        "content": "Cho sơ đồ thùng điện phân nóng chảy hỗn hợp Al~2~O~3~ và Na~3~AlF~6~. Xét các nhận định sau:",
        "image": "p5_img1_xref409_828x260.jpeg",
        "items": [
            {"id": "a", "text": "Nhiệt độ hoạt động của bình điện phân duy trì khoảng 950 - 1000 °C.", "ans": "Đ"},
            {"id": "b", "text": "Than chì lót đáy thùng đóng vai trò là cực âm (cathode).", "ans": "Đ"},
            {"id": "c", "text": "Cực than chì anode không bị hao hụt trong suốt quá trình điện phân.", "ans": "S"},
            {"id": "d", "text": "Al lỏng sinh ra nhẹ hơn chất điện li nóng chảy nên nổi lên trên mặt bồn.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Nhờ có cryolite nên nhiệt độ duy trì ở 950 - 1000°C.\\n"
            "- b) *ĐÚNG:* Lớp lót đáy là cathode than chì.\\n"
            "- c) *SAI:* Anode bị oxi hóa liên tục thành CO, CO~2~ nên tiêu hao rất nhanh.\\n"
            "- d) *SAI:* Al lỏng nặng hơn (D ≈ 2,3 g/cm^3^) nên chìm xuống đáy bồn, không nổi lên trên."
        )
    },
    {
        "id": "D3_C7",
        "num": 7,
        "type": "tf",
        "content": "Xét các phát biểu liên quan đến quá trình sản xuất nhôm bằng phương pháp điện phân nóng chảy:",
        "items": [
            {"id": "a", "text": "Hỗn hợp Al~2~O~3~ và cryolite nóng chảy ở nhiệt độ cao hơn Al~2~O~3~ tinh khiết.", "ans": "S"},
            {"id": "b", "text": "Tại cathode xảy ra quá trình khử O^2-^, tại anode xảy ra quá trình oxi hóa Al^3+^.", "ans": "S"},
            {"id": "c", "text": "Có thể thay Al~2~O~3~ bằng AlCl~3~ để điện phân nóng chảy sản xuất nhôm quy mô công nghiệp.", "ans": "S"},
            {"id": "d", "text": "Trong công nghiệp, nhôm được sản xuất từ quặng bauxite có thành phần chính là Al~2~O~3~.2H~2~O.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Thấp hơn rất nhiều (~950°C so với 2050°C).\\n"
            "- b) *SAI:* Ngược lại: tại cathode khử Al^3+^, tại anode oxi hóa O^2-^.\\n"
            "- c) *SAI:* AlCl~3~ là hợp chất cộng hóa trị có liên kết mạng tinh thể phân tử, thăng hoa ở nhiệt độ ~180°C nên không tồn tại thể lỏng dẫn điện để điện phân.\\n"
            "- d) *ĐÚNG:* Bauxite là nguồn quặng chính của nhôm."
        )
    },
    {
        "id": "D3_C8",
        "num": 8,
        "type": "tf",
        "content": "Điện phân nóng chảy hỗn hợp gồm aluminium oxide (Al~2~O~3~) và cryolite (Na~3~AlF~6~) để sản xuất nhôm trong công nghiệp. Sản phẩm ở cathode là nhôm lỏng và ở anode là hỗn hợp khí CO~2~, CO. Cho các nhận định:",
        "image": "p6_img1_xref492_431x216.png",
        "items": [
            {"id": "a", "text": "Tại anode, than chì bị tiêu hao liên tục trong quá trình điện phân.", "ans": "Đ"},
            {"id": "b", "text": "Nhôm lỏng có khối lượng riêng lớn hơn khối lượng riêng của cryolite nóng chảy.", "ans": "Đ"},
            {"id": "c", "text": "Tỉ lệ mol giữa Al~2~O~3~ và Na~3~AlF~6~ không thay đổi trong suốt quá trình điện phân.", "ans": "S"},
            {"id": "d", "text": "Khí thoát ra tại cửa thoát khí của anode có thành phần chủ yếu là CO~2~.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Than chì bị đốt cháy bởi O~2~.\\n"
            "- b) *ĐÚNG:* D(Al lỏng) ≈ 2,3 g/cm^3^ > D(cryolite) ≈ 2,1 g/cm^3^.\\n"
            "- c) *SAI:* Al~2~O~3~ bị điện phân tiêu thụ liên tục nên nồng độ giảm dần, định kỳ phải nạp thêm Al~2~O~3~.\\n"
            "- d) *ĐÚNG:* Thành phần khí anode chủ yếu là CO~2~ (chiếm 60 - 80% thể tích)."
        )
    },
    {
        "id": "D3_C9",
        "num": 9,
        "type": "essay",
        "content": "Trong công nghiệp, aluminium được sản xuất bằng phương pháp điện phân nóng chảy Al~2~O~3~ (được tinh chế từ quặng bauxite). Từ 1 tấn quặng bauxite có thể điều chế được bao nhiêu kg aluminium? Biết rằng quặng bauxite chứa 60% Al~2~O~3~.2H~2~O về khối lượng và hiệu suất chung của toàn bộ quá trình là 80%. (Kết quả làm tròn đến hàng đơn vị).",
        "ans": "188 kg",
        "solution": (
            "- Khối lượng Al~2~O~3~.2H~2~O trong 1 tấn (1000 kg) quặng: m = 1000 × 60% = 600 kg.\\n"
            "- Phân tử khối Al~2~O~3~.2H~2~O = 102 + 2 × 18 = 138 g/mol.\\n"
            "- Khối lượng Al theo lý thuyết (bảo toàn nguyên tố Al, trong 1 mol quặng có 2 mol Al):\\n"
            "  m~Al lt~ = 600 × (2 × 27 / 138) = 600 × (54 / 138) ≈ 234,78 kg.\\n"
            "- Do hiệu suất toàn bộ quá trình đạt 80%:\\n"
            "  m~Al thực tế~ = 234,78 × 80% ≈ 187,83 kg ≈ *188 kg*."
        )
    },
    {
        "id": "D3_C10",
        "num": 10,
        "type": "essay",
        "content": "Điện phân nóng chảy hỗn hợp gồm Al~2~O~3~ (10%) và cryolite (90%) với anode là than cốc và cathode là than chì. Sau một thời gian điện phân thu được 5,4 tấn Al tại cathode và hỗn hợp khí tại anode gồm CO~2~ (chiếm 80% theo thể tích) và CO (chiếm 20% theo thể tích). Giả thiết không có thêm sản phẩm nào khác sinh ra. Tính khối lượng carbon (theo tấn) đã bị oxi hoá tại anode. (Kết quả làm tròn đến hàng đơn vị).",
        "ans": "2 tấn",
        "solution": (
            "- Số mol Al thu được: n~Al~ = 5,4 × 10^6 / 27 = 200 000 mol = 0,2 Mmol.\\n"
            "- Quá trình tại cathode: Al^3+^ + 3e → Al → n~e~ = 3 × n~Al~ = 0,6 Mmol.\\n"
            "- Quá trình tại anode: 2O^2-^ → O~2~ + 4e → n~O~ = n~e~ / 2 = 0,3 Mmol (nguyên tử O).\\n"
            "- Gọi số mol của CO là x (Mmol) → số mol CO~2~ là 4x (Mmol) (vì tỉ lệ thể tích 80% : 20% = 4 : 1).\\n"
            "- Bảo toàn nguyên tố oxygen:\\n"
            "  n~O~ = 2 × n~CO2~ + 1 × n~CO~ = 2(4x) + x = 9x = 0,3 Mmol → x = 0,3 / 9 = 1/30 Mmol.\\n"
            "- Tổng số mol carbon bị đốt cháy:\\n"
            "  n~C~ = n~CO2~ + n~CO~ = 4x + x = 5x = 5 × (1/30) = 1/6 Mmol.\\n"
            "- Khối lượng carbon bị oxi hóa:\\n"
            "  m~C~ = (1/6) × 12 = 2,0 tấn ≈ *2 tấn*."
        )
    },
    {
        "id": "D3_C11",
        "num": 11,
        "type": "essay",
        "content": "Điện phân nóng chảy Al~2~O~3~ với các điện cực bằng than chì, thu được m kg Al ở cathode và 99,16 m^3^ (ở đkc) hỗn hợp khí X ở anode. Tỉ khối của X so với khí hydrogen bằng 16,7. Cho 1,2395 L X (đkc) phản ứng với dung dịch Ca(OH)~2~ dư, thu được 1,5 gam kết tủa. Biết các phản ứng xảy ra hoàn toàn. Giá trị của m là bao nhiêu? (Kết quả làm tròn đến hàng đơn vị).",
        "ans": "104 kg",
        "solution": (
            "- Thể tích khí X: V = 99,16 m^3^ = 99 160 L (ở đkc: 1 mol = 24,79 L).\\n"
            "- Tổng số mol khí X: n~X~ = 99 160 / 24,79 = 4000 mol.\\n"
            "- Khối lượng mol trung bình của X: M~X~ = 16,7 × 2 = 33,4 g/mol.\\n"
            "- Xét mẫu thử 1,2395 L X (đkc): n~mẫu~ = 1,2395 / 24,79 = 0,05 mol.\\n"
            "  Phản ứng với Ca(OH)~2~ dư: n~CaCO3~ = 1,5 / 100 = 0,015 mol → n~CO2~ = 0,015 mol.\\n"
            "  → Phần trăm thể tích của CO~2~ trong X: %V~CO2~ = (0,015 / 0,05) × 100% = 30%.\\n"
            "- Hỗn hợp X gồm CO~2~ (30%), CO (a%), O~2~ (b%):\\n"
            "  a + b = 100% - 30% = 70% = 0,70.\\n"
            "  M~X~ = 0,30 × 44 + 28a + 32b = 33,4 → 28a + 32b = 20,2.\\n"
            "  Giải hệ: a = 0,55 (55% CO); b = 0,15 (15% O~2~).\\n"
            "- Trong 4000 mol X:\\n"
            "  n~CO2~ = 4000 × 0,30 = 1200 mol; n~CO~ = 4000 × 0,55 = 2200 mol; n~O2~ = 4000 × 0,15 = 600 mol.\\n"
            "- Tổng số mol nguyên tử O sinh ra từ Al~2~O~3~:\\n"
            "  n~O~ = 2 × n~CO2~ + 1 × n~CO~ + 2 × n~O2~ = 2(1200) + 2200 + 2(600) = 5800 mol.\\n"
            "- Bảo toàn nguyên tố trong Al~2~O~3~: n~Al~ = (2/3) × n~O~ = (2/3) × 5800 = 11600 / 3 ≈ 3866,67 mol.\\n"
            "- Khối lượng nhôm thu được:\\n"
            "  m~Al~ = 3866,67 × 27 = 104 400 g = 104,4 kg ≈ *104 kg*."
        )
    }
]

# ==============================================================================
# DẠNG 4: TINH CHẾ KIM LOẠI (ĐỒNG)
# ==============================================================================
THEORY_DANG4 = {
    "title": "DẠNG 4: TINH CHẾ KIM LOẠI (ĐỒNG)",
    "sections": [
        {
            "heading": "I. Mục đích",
            "items": [
                "• Loại bỏ các tạp chất kim loại (Fe, Zn, Ni, Pb, Ag, Au...) ra khỏi đồng thô (chứa khoảng 95 - 98% Cu).",
                "• Thu được đồng tinh khiết thương phẩm có độ sạch lên đến 99,99% dùng cho ngành điện lực và điện tử."
            ]
        },
        {
            "heading": "II. Cấu tạo hệ thống điện phân tinh chế đồng",
            "image": "p7_img2_xref570_394x296.jpeg",
            "table": {
                "headers": ["Thành phần", "Bản chất và vai trò"],
                "rows": [
                    ["Anode (cực dương)", "Khối đồng thô cần tinh chế (có lẫn các tạp chất Fe, Zn, Ni, Ag, Au, Pt...)."],
                    ["Cathode (cực âm)", "Tấm đồng tinh khiết lá mỏng để kết tủa đồng nguyên chất bám vào."],
                    ["Dung dịch điện phân", "Dung dịch CuSO~4~ loãng có thêm H~2~SO~4~ để tăng khả năng dẫn điện và chống thủy phân."],
                    ["Nguồn điện một chiều (DC)", "Cung cấp dòng điện với hiệu điện thế được khống chế chính xác (~0,2 - 0,4 V)."]
                ]
            }
        },
        {
            "heading": "III. Phản ứng xảy ra tại các điện cực",
            "items": [
                "🔺 *Tại Anode (cực dương - đồng thô):*",
                "  • Kim loại Cu bị oxi hóa tan ra: Cu → Cu^2+^ + 2e",
                "  • Các kim loại hoạt động hơn Cu (Zn, Fe, Ni) cũng bị oxi hóa tan thành ion: Zn → Zn^2+^ + 2e; Fe → Fe^2+^ + 2e",
                "  • Các kim loại quý có tính khử yếu hơn Cu (Ag, Au, Pt) không bị oxi hóa, tách ra dưới dạng hạt nhỏ rơi xuống đáy bể tạo thành *bùn anode* (anode slime). Bùn anode được thu hồi để chiết tách kim loại quý.",
                "🔻 *Tại Cathode (cực âm - đồng tinh khiết):*",
                "  • Do thế điện cực chuẩn: E°(Zn^2+^/Zn) < E°(Fe^2+^/Fe) < E°(Cu^2+^/Cu) nên chỉ có ion Cu^2+^ có tính oxi hóa mạnh nhất bị khử:",
                "      Cu^2+^ + 2e → Cu (kết tinh bám chắc lên lá đồng cathode)",
                "  • Các cation tạp chất Zn^2+^, Fe^2+^, Ni^2+^ vẫn tồn tại trong dung dịch mà không bị khử.",
                "→ Kết quả: Đồng cathode có độ tinh khiết cực cao (> 99,99%)."
            ]
        },
        {
            "heading": "IV. Lưu ý quan trọng khi tinh chế đồng",
            "table": {
                "headers": ["Nội dung", "Lưu ý quan trọng"],
                "rows": [
                    ["1. Bản chất anode", "Anode là đồng thô; tan dần trong quá trình điện phân."],
                    ["2. Bản chất cathode", "Cathode là lá đồng mỏng tinh khiết; đồng tạo thành bám đều làm tăng độ dày tấm cathode."],
                    ["3. Khống chế hiệu điện thế", "Chỉ đặt hiệu điện thế nhỏ (~0,2 - 0,4 V) đủ để khử Cu^2+^, tuyệt đối không để khử tạp chất Zn^2+^, Fe^2+^."],
                    ["4. Giá trị bùn anode", "Bùn anode chứa hàm lượng cao vàng Au, bạc Ag, platin Pt... có giá trị kinh tế rất lớn."],
                    ["5. Nồng độ dung dịch CuSO~4~", "Nồng độ ion Cu^2+^ trong dung dịch giảm nhẹ do một phần dòng điện dùng để oxi hóa Zn, Fe ở anode."]
                ]
            }
        }
    ]
}

EXERCISES_DANG4 = [
    {
        "id": "D4_C1",
        "num": 1,
        "type": "tf",
        "content": "Điện phân dung dịch CuSO~4~ với các điện cực đều làm bằng Cu (hiện tượng dương cực tan). Cho các phát biểu sau:",
        "image": "p7_img1_xref569_186x190.png",
        "items": [
            {"id": "a", "text": "Tại anode xảy ra quá trình oxi hoá Cu thành ion Cu^2+^ khuếch tán vào dung dịch.", "ans": "Đ"},
            {"id": "b", "text": "Tại cathode xảy ra quá trình khử ion Cu^2+^ thành Cu bám vào điện cực.", "ans": "Đ"},
            {"id": "c", "text": "Phản ứng chung trong bình điện phân là: 2Cu^2+^(aq) + 2H~2~O(l) → 2Cu(s) + O~2~(g) + 4H^+^(aq).", "ans": "S"},
            {"id": "d", "text": "Nồng độ ion Cu^2+^ trong dung dịch không đổi trong suốt quá trình điện phân.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Anode tan: Cu → Cu^2+^ + 2e.\\n"
            "- b) *ĐÚNG:* Cathode: Cu^2+^ + 2e → Cu.\\n"
            "- c) *SAI:* Khi dùng điện cực đồng (dương cực tan), nước không bị điện phân ở anode nên không sinh ra O~2~ và H^+^.\\n"
            "- d) *ĐÚNG:* Lượng Cu tan ở anode đúng bằng lượng Cu bám ở cathode nên nồng độ Cu^2+^ và màu xanh của dung dịch không đổi."
        )
    },
    {
        "id": "D4_C2",
        "num": 2,
        "type": "tf",
        "content": "Tinh luyện đồng (Cu) bằng phương pháp điện phân được tiến hành như sau: Khối đồng có độ tinh khiết thấp gắn với cực dương; tấm đồng mỏng có độ tinh khiết cao gắn với cực âm. Dung dịch điện phân là CuSO~4~. Giả sử các tạp chất kim loại khác không bị khử ở cathode. Cho các phát biểu:",
        "items": [
            {"id": "a", "text": "Tại anode chủ yếu xảy ra quá trình oxi hoá H~2~O thành khí O~2~ và ion H^+^.", "ans": "S"},
            {"id": "b", "text": "Các tạp chất không bị điện phân sẽ bám chặt lại trên bề mặt anode.", "ans": "S"},
            {"id": "c", "text": "Nồng độ ion Cu^2+^ trong dung dịch cơ bản ổn định trong suốt quá trình điện phân.", "ans": "Đ"},
            {"id": "d", "text": "Đồng thu được ở cathode có độ tinh khiết rất cao (trên 99,9%).", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Tại anode đồng thô bị oxi hóa tan ra, nước không bị oxi hóa.\\n"
            "- b) *SAI:* Các tạp chất tan vào dung dịch (Zn, Fe) hoặc rơi xuống đáy tạo bùn anode (Ag, Au), không bám lại trên anode.\\n"
            "- c) *ĐÚNG:* Tốc độ Cu tan ở anode xấp xỉ tốc độ Cu bám ở cathode.\\n"
            "- d) *ĐÚNG:* Đồng kết tinh ở cathode có độ sạch đạt chuẩn 99,99%."
        )
    },
    {
        "id": "D4_C3",
        "num": 3,
        "type": "tf",
        "content": "Một nhóm học sinh nghiên cứu về tinh chế kim loại đồng bằng phương pháp điện phân từ một thanh đồng không tinh khiết. Nhóm học sinh đã thực hiện thí nghiệm theo các bước:\\n- Bước 1: Cân xác định khối lượng ban đầu của thanh đồng nguyên chất (3,21 gam) và thanh đồng không tinh khiết (3,30 gam).\\n- Bước 2: Nối thanh đồng nguyên chất với một điện cực và thanh đồng không tinh khiết với điện cực còn lại của nguồn điện một chiều, rồi nhúng vào bình điện phân chứa dung dịch CuSO~4~ đã acid hóa.\\n- Bước 3: Điện phân ở hiệu điện thế phù hợp. Sau thời gian điện phân t giây, sấy khô rồi cân lại thấy thanh đồng nguyên chất nặng 3,51 gam và thanh không tinh khiết nặng 2,98 gam.\\nCho các phát biểu sau:",
        "items": [
            {"id": "a", "text": "Ở cực dương xảy ra quá trình khử: Cu^2+^ + 2e → Cu.", "ans": "S"},
            {"id": "b", "text": "Màu xanh của dung dịch copper(II) sulfate gần như không thay đổi.", "ans": "Đ"},
            {"id": "c", "text": "Thanh đồng nguyên chất được nối với điện cực âm (cathode), thanh đồng không tinh khiết nối với điện cực dương (anode).", "ans": "Đ"},
            {"id": "d", "text": "Khối lượng thanh đồng nguyên chất tăng lên 0,30 gam đúng bằng khối lượng Cu nguyên chất kết tủa từ dung dịch.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Cực dương là Anode, xảy ra sự oxi hóa Cu → Cu^2+^ + 2e.\\n"
            "- b) *ĐÚNG:* Quá trình dương cực tan giữ nồng độ Cu^2+^ ổn định.\\n"
            "- c) *ĐÚNG:* Đồng tinh khiết là cathode (-), đồng thô là anode (+).\\n"
            "- d) *ĐÚNG:* Khối lượng đồng nguyên chất tăng = 3,51 - 3,21 = 0,30 gam."
        )
    },
    {
        "id": "D4_C4",
        "num": 4,
        "type": "tf",
        "content": "Một nhóm học sinh tìm hiểu quá trình thu hồi đồng từ một đồng xu làm bằng hợp kim Cu-Zn chứa khoảng 95% Cu về khối lượng. Giả thuyết của nhóm là: 'Khi điện phân, chỉ có tạp chất trong đồng xu tan hết vào trong dung dịch, còn lại sẽ là đồng tinh khiết'. Tiến hành thí nghiệm:\\n- Cân đồng xu ban đầu (2,23 gam) và thanh đồng tinh khiết (2,55 gam).\\n- Nối đồng xu với cực dương, thanh đồng tinh khiết với cực âm, nhúng vào dung dịch CuSO~4~.\\n- Điện phân một thời gian, sấy khô rồi cân lại thấy đồng xu nặng 1,94 gam và thanh đồng tinh khiết nặng m~1~ gam.\\nCho các nhận định sau:",
        "items": [
            {"id": "a", "text": "Trong thí nghiệm trên, đồng xu được nối với cực dương, thanh đồng tinh khiết nối với cực âm.", "ans": "Đ"},
            {"id": "b", "text": "Giá trị của m~1~ lớn hơn 2,55 gam.", "ans": "Đ"},
            {"id": "c", "text": "Ở cực dương xảy ra quá trình khử ion kim loại.", "ans": "S"},
            {"id": "d", "text": "Do khối lượng của đồng xu giảm nên giả thuyết ban đầu của nhóm học sinh là đúng.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Đồng xu chứa Cu cần thu hồi đóng vai trò anode (+), thanh Cu tinh khiết là cathode (-).\\n"
            "- b) *ĐÚNG:* Cu^2+^ bị khử bám vào cathode nên khối lượng thanh đồng tăng m~1~ > 2,55 gam.\\n"
            "- c) *SAI:* Cực dương (anode) xảy ra quá trình oxi hóa.\\n"
            "- d) *SAI:* Cả Zn và Cu đều tan ra ở anode, chứ không phải Cu nguyên chất ở lại trên đồng xu. Do đó giả thuyết ban đầu là sai."
        )
    },
    {
        "id": "D4_C5",
        "num": 5,
        "type": "essay",
        "content": "Một nhóm học sinh tìm hiểu quá trình thu hồi kim loại đồng từ một đồng xu làm bằng hợp kim Cu-Zn chứa khoảng 95% Cu về khối lượng. Tiến hành thí nghiệm:\\n- Cân ban đầu đồng xu nặng 2,6 gam và thanh Cu tinh khiết nặng 2,85 gam.\\n- Nối đồng xu với cực dương, thanh Cu tinh khiết với cực âm, nhúng vào bình điện phân chứa dung dịch CuSO~4~.\\n- Điện phân với cường độ dòng điện không đổi I = 5 A, hiệu suất 100%.\\n- Sau thời gian t giây, sấy khô cân lại thấy khối lượng đồng xu là 1,894 gam.\\nTính giá trị của t (kết quả làm tròn đến hàng đơn vị). Cho biết Cu = 64; Zn = 65; F = 96 500 C/mol.",
        "ans": "425 giây",
        "solution": (
            "- Khối lượng đồng xu ban đầu: 2,6 gam; chứa 95% Cu → m~Cu ban đầu~ = 2,6 × 0,95 = 2,47 g; m~Zn ban đầu~ = 2,6 - 2,47 = 0,13 g.\\n"
            "- Độ giảm khối lượng của đồng xu: Δm = 2,6 - 1,894 = 0,706 g.\\n"
            "- Vì Zn có tính khử mạnh hơn Cu nên Zn tan trước:\\n"
            "  Toàn bộ Zn tan hết: n~Zn~ = 0,13 / 65 = 0,002 mol → n~e(Zn)~ = 2 × 0,002 = 0,004 mol.\\n"
            "- Lượng Cu tan tiếp theo:\\n"
            "  m~Cu tan~ = 0,706 - 0,13 = 0,576 g → n~Cu~ = 0,576 / 64 = 0,009 mol.\\n"
            "  n~e(Cu)~ = 2 × 0,009 = 0,018 mol.\\n"
            "- Tổng số mol electron trao đổi: n~e~ = 0,004 + 0,018 = 0,022 mol.\\n"
            "- Điện lượng q = n~e~ × F = 0,022 × 96500 = 2123 C.\\n"
            "- Thời gian điện phân: t = q / I = 2123 / 5 = 424,6 giây ≈ *425 giây*."
        )
    },
    {
        "id": "D4_C6",
        "num": 6,
        "type": "tf",
        "content": "Để tách lấy kim loại Cu từ một thanh hợp kim Zn-Cu (Cu chiếm khoảng 96% về khối lượng), tiến hành thí nghiệm điện phân với anode là thanh hợp kim, cathode là tấm Cu tinh khiết trong dung dịch CuSO~4~. Toàn bộ lượng kim loại tan ra ở anode tạo thành cation trong dung dịch và lượng Cu sinh ra bám hết vào cathode. Sau một thời gian, dừng điện phân, rửa sạch sấy khô rồi cân lại điện cực. Cho các nhận định sau:",
        "image": "p8_img1_xref654_490x289.jpeg",
        "items": [
            {"id": "a", "text": "Khối lượng kim loại Cu tan ra ở anode tỉ lệ với điện lượng chạy qua bình điện phân.", "ans": "Đ"},
            {"id": "b", "text": "Ở anode, kim loại Zn có tính khử mạnh hơn nên bị oxi hóa trước Cu.", "ans": "Đ"},
            {"id": "c", "text": "Trong dung dịch sau khi kết thúc điện phân có chứa cation Zn^2+^(aq).", "ans": "Đ"},
            {"id": "d", "text": "Nồng độ mol của ion Cu^2+^(aq) trong dung dịch sau khi kết thúc điện phân thấp hơn so với ban đầu.", "ans": "Đ"},
            {"id": "đ", "text": "Ở cathode, xảy ra sự oxi hóa ion Cu^2+^(aq) thành kim loại Cu.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Theo định luật Faraday m = (A.I.t)/(n.F).\\n"
            "- b) *ĐÚNG:* E°(Zn^2+^/Zn) = -0,76 V < E°(Cu^2+^/Cu) = +0,34 V nên Zn bị oxi hóa trước.\\n"
            "- c) *ĐÚNG:* Zn tan tạo ion Zn^2+^ hòa tan trong dung dịch.\\n"
            "- d) *ĐÚNG:* Vì một phần electron ở cực dương do Zn tan cung cấp (không sinh Cu^2+^), trong khi ở cathode Cu^2+^ vẫn nhận electron liên tục để tạo Cu, do đó nồng độ Cu^2+^ trong dung dịch giảm nhẹ.\\n"
            "- đ) *SAI:* Ở cathode xảy ra sự khử Cu^2+^ + 2e → Cu, không phải sự oxi hóa."
        )
    },
    {
        "id": "D4_C7",
        "num": 7,
        "type": "tf",
        "content": "Một học sinh muốn xác định hàm lượng Cu có trong hợp kim Cu-Zn đã tiến hành thí nghiệm như sau:\\n- Cân thanh hợp kim Cu-Zn có khối lượng 2,25 gam (hàm lượng Zn trên 20%) và thanh Cu nguyên chất nặng 2,58 gam.\\n- Nối thanh hợp kim với cực dương, thanh Cu nguyên chất với cực âm của nguồn điện một chiều, rồi nhúng vào bình điện phân chứa 100 mL dung dịch CuSO~4~ 1M. Sau một thời gian điện phân, làm khô rồi cân lại thấy thanh Cu nguyên chất nặng 2,99 gam.\\nCho các nhận định sau:",
        "items": [
            {"id": "a", "text": "Quá trình điện phân xảy ra hiện tượng dương cực tan.", "ans": "Đ"},
            {"id": "b", "text": "Khối lượng thanh đồng nguyên chất tăng thêm 0,41 gam.", "ans": "Đ"},
            {"id": "c", "text": "Dung dịch sau điện phân chỉ chứa duy nhất một muối tan.", "ans": "S"},
            {"id": "d", "text": "Do anion sulfate không bị điện phân trong dung dịch nên tại anode xảy ra quá trình oxi hoá nước giải phóng O~2~.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Cực dương làm bằng hợp kim kim loại nên tan ra (dương cực tan).\\n"
            "- b) *ĐÚNG:* m(tăng) = 2,99 - 2,58 = 0,41 gam.\\n"
            "- c) *SAI:* Dung dịch chứa cả ZnSO~4~ và CuSO~4~ dư (2 muối tan).\\n"
            "- d) *SAI:* Tại anode kim loại Cu và Zn bị oxi hóa trước nước, nước không bị điện phân."
        )
    },
    {
        "id": "D4_C8",
        "num": 8,
        "type": "tf",
        "content": "Có thể tinh chế đồng không tinh khiết có lẫn một lượng nhỏ Ag, Zn, Ni bằng phương pháp điện phân. Trong quá trình điện phân phần lớn dòng điện được sử dụng để hòa tan đồng ở anode và kết tủa đồng tinh khiết trên cathode. Tuy nhiên một phần nhỏ dòng điện bị hao phí cho việc hòa tan các tạp chất ở anode. Khi cho dòng điện 20 A chạy qua bình điện phân tinh chế đồng, sau 10 giờ ngắt dòng điện thì thấy có 225 g đồng tinh khiết bám trên cathode. Cho biết E°(Cu^2+^/Cu) = +0,34 V; E°(Ag^+^/Ag) = +0,799 V; F = 96 500 C/mol; Cu = 64. Cho các nhận định sau:",
        "image": "p8_img3_xref656_525x280.png",
        "items": [
            {"id": "a", "text": "Tỉ lệ phần trăm dòng điện bị hao phí xấp xỉ 5,76%.", "ans": "Đ"},
            {"id": "b", "text": "Muối đồng sử dụng trong dung dịch điện phân luôn bắt buộc là CuCl~2~.", "ans": "S"},
            {"id": "c", "text": "Anode trong bể tinh chế được làm bằng khối đồng có độ tinh khiết cao 99,99%.", "ans": "S"},
            {"id": "d", "text": "Do Ag có tính khử yếu hơn Cu nên khi đồng tan ra thì bạc không bị tan mà lắng xuống đáy bình tạo bùn anode.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:*\\n"
            "  Điện lượng tổng cộng: q = I × t = 20 × (10 × 3600) = 720 000 C.\\n"
            "  Điện lượng có ích để kết tủa 225 g Cu: q~ci~ = (225 / 64) × 2 × 96500 ≈ 678 516 C.\\n"
            "  Hiệu suất dòng điện: H = (678 516 / 720 000) × 100% ≈ 94,24%.\\n"
            "  → Phần trăm dòng điện hao phí: 100% - 94,24% = 5,76%.\\n"
            "- b) *SAI:* Thường dùng dung dịch CuSO~4~ có thêm H~2~SO~4~, không dùng CuCl~2~ vì Cl^-^ dễ bị oxi hóa sinh khí Cl~2~ độc hại.\\n"
            "- c) *SAI:* Anode là khối đồng thô (chứa tạp chất); cathode mới là đồng tinh khiết.\\n"
            "- d) *ĐÚNG:* Ag có thế điện cực chuẩn +0,799 V > +0,34 V nên không bị oxi hóa, rơi xuống đáy tạo bùn anode."
        )
    }
]
'''

with open(DATA_FILE, "a", encoding="utf-8") as f:
    f.write(dang3_and_dang4)

print("Appended Dạng 3 and Dạng 4.")
