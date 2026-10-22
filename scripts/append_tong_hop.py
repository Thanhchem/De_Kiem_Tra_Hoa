# -*- coding: utf-8 -*-
"""
Append BÀI TẬP TỔNG HỢP (Câu 1 đến Câu 36) to hoa12_to36_data.py
"""
import os
import sys

DATA_FILE = r"c:\Antigravity_Thanh\scripts\hoa12_to36_data.py"

tong_hop_content = '''
# ==============================================================================
# BÀI TẬP TỔNG HỢP: CÂU 1 ĐẾN CÂU 36
# ==============================================================================
TONG_HOP_MCQ_TF = [
    {
        "id": "TH_C1",
        "num": 1,
        "type": "mcq",
        "content": "Điện phân dung dịch copper(II) sulfate bằng cặp điện cực đồng với cường độ dòng điện không đổi. Đồ thị biểu diễn sự thay đổi khối lượng anode và cathode theo thời gian lần lượt là:",
        "image": "p9_img2_xref738_1232x269.png",
        "options": [
            "A. (3) và (4).",
            "B. (2) và (3).",
            "C. (1) và (3).",
            "D. (3) và (2)."
        ],
        "ans": "D",
        "solution": (
            "- Khi điện phân dung dịch CuSO~4~ với cả hai điện cực bằng Cu (hiện tượng dương cực tan):\\n"
            "  + Tại Anode: Cu bị oxi hóa Cu → Cu^2+^ + 2e nên khối lượng anode giảm đều theo thời gian tuyến tính → tương ứng đồ thị *(3)* (đường thẳng đi xuống).\\n"
            "  + Tại Cathode: Cu^2+^ bị khử Cu^2+^ + 2e → Cu bám vào tấm cathode nên khối lượng cathode bắt đầu từ m~0~ ban đầu tăng đều tuyến tính theo thời gian → tương ứng đồ thị *(2)* (đường thẳng đi lên từ giá trị m~0~ > 0).\\n"
            "→ Thứ tự lần lượt của anode và cathode là *(3) và (2)*. Chọn *D*."
        )
    },
    {
        "id": "TH_C2",
        "num": 2,
        "type": "mcq",
        "content": "Trong công nghiệp, nhôm được sản xuất từ quặng bauxite qua 2 giai đoạn chính: tinh chế quặng và điện phân Al~2~O~3~ nóng chảy có cryolite Na~3~AlF~6~. Sản phẩm điện phân ở cathode là nhôm lỏng và ở anode là hỗn hợp khí CO~2~, CO. Cấu tạo bể điện phân như hình vẽ:\\nSau một thời gian điện phân thu được 5,4 tấn Al tại cathode và hỗn hợp khí tại anode gồm CO~2~ (chiếm 80% thể tích) và CO (chiếm 20% thể tích). Giả thiết không có sản phẩm phụ nào khác. Khối lượng carbon bị oxi hóa trên anode là bao nhiêu tấn?",
        "image": "p9_img3_xref739_773x202.jpeg",
        "options": [
            "A. 3,1 tấn.",
            "B. 2,0 tấn.",
            "C. 0,8 tấn.",
            "D. 1,6 tấn."
        ],
        "ans": "B",
        "solution": (
            "- Số mol Al thu được: n~Al~ = 5,4 × 10^6 / 27 = 200 000 mol = 0,2 Mmol.\\n"
            "- Số mol electron trao đổi: n~e~ = 3 × n~Al~ = 0,6 Mmol.\\n"
            "- Số mol nguyên tử O sinh ra ở anode: n~O~ = n~e~ / 2 = 0,3 Mmol.\\n"
            "- Hỗn hợp khí anode có V(CO~2~) : V(CO) = 80% : 20% = 4 : 1.\\n"
            "  Gọi n~CO~ = x (Mmol) → n~CO2~ = 4x (Mmol).\\n"
            "- Bảo toàn nguyên tố O: 2(4x) + 1(x) = 9x = 0,3 Mmol → x = 0,3 / 9 = 1/30 Mmol.\\n"
            "- Tổng số mol carbon bị đốt cháy: n~C~ = n~CO2~ + n~CO~ = 4x + x = 5x = 5 × (1/30) = 1/6 Mmol.\\n"
            "- Khối lượng carbon tiêu hao: m~C~ = (1/6) × 12 = 2,0 tấn. Chọn *B*."
        )
    },
    {
        "id": "TH_C3",
        "num": 3,
        "type": "tf",
        "content": "Một nhóm học sinh thực hiện thí nghiệm điện phân dung dịch NaCl bão hòa theo các bước:\\n• Bước 1: Chuẩn bị bình điện phân có màng ngăn và hai điện cực graphite.\\n• Bước 2: Rót 100 mL dung dịch NaCl bão hòa vào bình sao cho hai điện cực đều ngập trong dung dịch.\\n• Bước 3: Nối hai điện cực với nguồn điện một chiều và tiến hành điện phân, quan sát thấy có khí thoát ra ở hai cực.\\n• Bước 4: Lấy màng ngăn ra khỏi bình và tiếp tục điện phân một thời gian.\\nCho các phát biểu sau:",
        "items": [
            {"id": "a", "text": "Sau bước 4, dung dịch thu được trong bình điện phân có tính tẩy màu.", "ans": "Đ"},
            {"id": "b", "text": "Ở bước 3, tại cathode xảy ra quá trình khử ion Na^+^.", "ans": "S"},
            {"id": "c", "text": "Ở bước 1, sử dụng màng ngăn để các sản phẩm sinh ra ở hai điện cực (Cl~2~ và NaOH) không phản ứng với nhau.", "ans": "Đ"},
            {"id": "d", "text": "Ở bước 3, khi đưa mẫu giấy quỳ tím ẩm đến gần điện cực anode thì quỳ tím hóa đỏ rồi nhanh chóng bị mất màu.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Khi bỏ màng ngăn, Cl~2~ ở anode phản ứng với NaOH ở cathode tạo nước Javel (NaCl + NaClO + H~2~O) có tính tẩy màu mạnh.\\n"
            "- b) *SAI:* Ở cathode nước bị khử sinh H~2~ (2H~2~O + 2e → H~2~ + 2OH^-^), Na^+^ không bị khử trong dung dịch nước.\\n"
            "- c) *ĐÚNG:* Màng ngăn xốp ngăn Cl~2~ phản ứng với NaOH và H~2~.\\n"
            "- d) *ĐÚNG:* Khí Cl~2~ làm quỳ ẩm hóa đỏ do tạo axit HCl và HClO, sau đó HClO có tính oxi hóa mạnh tẩy màu quỳ tím."
        )
    },
    {
        "id": "TH_C4",
        "num": 4,
        "type": "tf",
        "content": "Một nhóm học sinh tiến hành thí nghiệm mạ đồng lên chiếc thìa kim loại. Giả thuyết ban đầu là: 'Nồng độ CuSO~4~ trước và sau quá trình mạ là không đổi'. Tiến hành theo các bước:\\n- Bước 1: Cân xác định khối lượng chiếc thìa là 10 gam và của thanh đồng nguyên chất là 15 gam.\\n- Bước 2: Nối chiếc thìa với một cực và thanh đồng với cực còn lại của nguồn điện một chiều rồi nhúng vào cốc chứa dung dịch CuSO~4~ để mạ với hiệu điện thế thích hợp.\\n- Bước 3: Sau 15 phút, lấy chiếc thìa và thanh đồng ra, sấy khô cân lại thấy chiếc thìa nặng 10,32 gam, thanh đồng nặng 14,68 gam. Hiệu suất đạt 100%.\\nCho các phát biểu sau:",
        "image": "p9_img4_xref740_307x240.png",
        "items": [
            {"id": "a", "text": "Tại anode xảy ra quá trình khử ion Cu^2+^, tại cathode xảy ra quá trình oxi hoá Cu.", "ans": "S"},
            {"id": "b", "text": "Sau khi mạ xong, độ giảm khối lượng của thanh đồng bằng độ tăng khối lượng của chiếc thìa.", "ans": "Đ"},
            {"id": "c", "text": "Do khối lượng của thanh đồng giảm nên giả thuyết ban đầu của nhóm học sinh là sai.", "ans": "S"},
            {"id": "d", "text": "Thanh đồng được nối với cực dương (anode), chiếc thìa được nối với cực âm (cathode) của nguồn điện.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Tại anode xảy ra oxi hóa Cu → Cu^2+^ + 2e; tại cathode xảy ra khử Cu^2+^ + 2e → Cu.\\n"
            "- b) *ĐÚNG:* Khối lượng đồng tan ra ở anode = 15 - 14,68 = 0,32 gam đúng bằng khối lượng Cu bám vào chiếc thìa = 10,32 - 10 = 0,32 gam.\\n"
            "- c) *SAI:* Vì lượng Cu tan ở anode bằng lượng Cu bám ở cathode nên nồng độ Cu^2+^ trong dung dịch không đổi, tức là giả thuyết ban đầu của nhóm là đúng.\\n"
            "- d) *ĐÚNG:* Thanh đồng mạ là anode (+), vật cần mạ (thìa) là cathode (-)."
        )
    },
    {
        "id": "TH_C5",
        "num": 5,
        "type": "tf",
        "content": "Bộ dụng cụ Hofmann dưới đây được sử dụng để điện phân nước với chất điện li hòa tan là H~2~SO~4~ loãng. Cho các phát biểu sau:",
        "items": [
            {"id": "a", "text": "Tại điện cực nối với cực âm xảy ra quá trình khử ion H^+^.", "ans": "Đ"},
            {"id": "b", "text": "Ống thu được thể tích khí lớn gấp đôi là ống thu khí H~2~.", "ans": "Đ"},
            {"id": "c", "text": "Nếu thay dung dịch H~2~SO~4~ bằng dung dịch HCl thì khí thu được ở hai điện cực vẫn là H~2~ và O~2~.", "ans": "S"},
            {"id": "d", "text": "Nếu ở một ống thu được 50 mL khí O~2~ thì ở ống còn lại sẽ thu được 100 mL khí H~2~ (trong cùng điều kiện nhiệt độ, áp suất).", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Ở cathode (-) ion H^+^ nhận e: 2H^+^ + 2e → H~2~.\\n"
            "- b) *ĐÚNG:* Theo phương trình 2H~2~O → 2H~2~ + O~2~, tỉ lệ V(H~2~) : V(O~2~) = 2 : 1.\\n"
            "- c) *SAI:* Nếu dùng dung dịch HCl thì ở anode ion Cl^-^ bị oxi hóa sinh ra khí Cl~2~ độc hại, không phải O~2~.\\n"
            "- d) *ĐÚNG:* Thể tích H~2~ gấp đôi thể tích O~2~: 50 mL O~2~ đi kèm 100 mL H~2~."
        )
    },
    {
        "id": "TH_C6",
        "num": 6,
        "type": "tf",
        "content": "Một học sinh làm thí nghiệm điện phân dung dịch NaCl không màng ngăn:\\n- Lắp thiết bị điện phân dung dịch NaCl với điện cực trơ graphite.\\n- Rót 80 mL dung dịch NaCl bão hoà vào cốc rồi nhúng 2 điện cực vào dung dịch.\\n- Nối với nguồn điện và điện phân trong khoảng 5 phút.\\n- Thả một cánh hoa màu hồng vào cốc chứa dung dịch sau điện phân.\\nCho các nhận xét sau:",
        "image": "p10_img4_xref846_245x262.jpeg",
        "items": [
            {"id": "a", "text": "Tại điện cực anode có khí H~2~ thoát ra.", "ans": "S"},
            {"id": "b", "text": "Tại điện cực cathode xảy ra quá trình khử nước giải phóng H~2~.", "ans": "Đ"},
            {"id": "c", "text": "Cánh hoa hồng bị mất màu do dung dịch sau điện phân có chứa NaClO.", "ans": "Đ"},
            {"id": "d", "text": "Dùng nắp đậy trong quá trình điện phân để hạn chế sự khuếch tán khí Cl~2~ ra ngoài môi trường gây độc hại.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Tại anode Cl^-^ bị oxi hóa tạo khí Cl~2~, khí H~2~ thoát ra ở cathode.\\n"
            "- b) *ĐÚNG:* Cathode: 2H~2~O + 2e → H~2~ + 2OH^-^.\\n"
            "- c) *ĐÚNG:* Không màng ngăn: Cl~2~ + 2NaOH → NaCl + NaClO + H~2~O. NaClO có tính tẩy màu mạnh làm phai màu cánh hoa.\\n"
            "- d) *ĐÚNG:* Giữ an toàn trong phòng thí nghiệm tránh ngộ độc Cl~2~."
        )
    },
    {
        "id": "TH_C7",
        "num": 7,
        "type": "tf",
        "content": "Bộ dụng cụ bình điện phân Hofmann được dùng để điện phân nước với chất điện li H~2~SO~4~ loãng. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Thêm H~2~SO~4~ nhằm mục đích tăng độ dẫn điện cho dung dịch vì nước tinh khiết dẫn điện rất kém.", "ans": "Đ"},
            {"id": "b", "text": "Thể tích khí thu được ở hai nhánh P và Q luôn bằng nhau trong cùng điều kiện.", "ans": "S"},
            {"id": "c", "text": "Khí thu được ở cực dương (anode) có khả năng làm tàn đốm que diêm bùng cháy.", "ans": "Đ"},
            {"id": "d", "text": "Khí thu được ở cực âm (cathode) cháy với ngọn lửa màu xanh nhạt phát ra tiếng nổ lách tách.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Nước nguyên chất điện li cực yếu ([H^+^] = [OH^-^] = 10^-7^ M) nên dẫn điện kém, cần thêm acid trơ làm chất dẫn điện.\\n"
            "- b) *SAI:* V(H~2~) = 2 V(O~2~).\\n"
            "- c) *ĐÚNG:* Cực dương sinh khí O~2~ duy trì sự cháy, làm bùng cháy que đóm.\\n"
            "- d) *ĐÚNG:* Cực âm sinh khí H~2~ cháy trong không khí tạo ngọn lửa xanh nhạt."
        )
    },
    {
        "id": "TH_C8",
        "num": 8,
        "type": "tf",
        "content": "Một học sinh muốn xác định hàm lượng Cu có trong hợp kim Cu-Zn:\\n- Bước 1: Cân thanh hợp kim Cu-Zn (2,25 gam) và thanh đồng nguyên chất (2,58 gam).\\n- Bước 2: Nối thanh hợp kim với cực dương, thanh Cu nguyên chất với cực âm của nguồn điện một chiều rồi nhúng vào bình điện phân chứa dung dịch CuSO~4~.\\n- Sau một thời gian, sấy khô rồi cân lại thấy thanh Cu nguyên chất nặng 2,99 gam. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Tại anode xảy ra quá trình oxi hoá H~2~O do gốc SO~4~^2-^ không tham gia điện phân.", "ans": "S"},
            {"id": "b", "text": "Khối lượng của thanh đồng nguyên chất ở cathode tăng thêm 0,41 gam.", "ans": "Đ"},
            {"id": "c", "text": "Quá trình điện phân xảy ra hiện tượng dương cực tan.", "ans": "Đ"},
            {"id": "d", "text": "Toàn bộ thanh hợp kim Cu-Zn ở cực dương đã tan hoàn toàn trong thí nghiệm trên.", "ans": "S"}
        ],
        "solution": (
            "- a) *SAI:* Cực dương làm bằng hợp kim kim loại (Zn, Cu) nên các kim loại này bị oxi hóa tan ra, nước không bị điện phân.\\n"
            "- b) *ĐÚNG:* Δm = 2,99 - 2,58 = 0,41 gam.\\n"
            "- c) *ĐÚNG:* Anode bằng kim loại tan dần.\\n"
            "- d) *SAI:* Mới điện phân một thời gian, thanh hợp kim chỉ tan một phần."
        )
    },
    {
        "id": "TH_C9",
        "num": 9,
        "type": "tf",
        "content": "Điện phân có nhiều ứng dụng trong thực tiễn như luyện kim, tinh chế kim loại, mạ điện... Cho các nhận định sau:",
        "items": [
            {"id": "a", "text": "Sau quá trình điện phân dung dịch CuSO~4~ với hai điện cực trơ than chì thì pH của dung dịch sẽ giảm.", "ans": "Đ"},
            {"id": "b", "text": "Cần mạ một lớp Ag lên một mặt đĩa tròn bán kính 12 cm, dày 0,01 mm; với I = 5 A thì chỉ cần thời gian 2 phút là mạ xong.", "ans": "S"},
            {"id": "c", "text": "Trong quá trình điện phân dung dịch các chất, tùy trường hợp mà khối lượng dung dịch có thể tăng hoặc giảm.", "ans": "Đ"},
            {"id": "d", "text": "Khi điện phân dung dịch NaCl ở cathode nước bị khử do ion Na^+^ có tính oxi hóa mạnh hơn nước.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* 2CuSO~4~ + 2H~2~O → 2Cu + O~2~ + 2H~2~SO~4~ sinh thêm H^+^ làm pH dung dịch giảm mạnh.\\n"
            "- b) *SAI:* Ở Bài tập 2 ta đã tính với đĩa r = 12 cm, h = 0,01 mm cần m~Ag~ ≈ 4,75 g; với I = 5 A cần t ≈ 14 phút (không thể 2 phút).\\n"
            "- c) *ĐÚNG:* Đa số trường hợp khối lượng dung dịch giảm (do kết tủa bám điện cực hoặc khí bay ra), nhưng trong quá trình dương cực tan kim loại nặng tan vào bù cho kim loại nhẹ có thể làm khối lượng dung dịch biến đổi.\\n"
            "- d) *SAI:* Na^+^ có tính oxi hóa rất yếu (E° = -2,71 V) nên nước (E° = -0,83 V) có tính oxi hóa mạnh hơn và bị khử trước."
        )
    },
    {
        "id": "TH_C10",
        "num": 10,
        "type": "tf",
        "content": "Do Al~2~O~3~ có nhiệt độ nóng chảy rất cao (2050°C) nên được trộn cùng với cryolite (Na~3~AlF~6~) để tạo hỗn hợp nóng chảy ở ~950°C. Sơ đồ thùng điện phân vận hành với hiệu điện thế U ≈ 5 V và cường độ dòng điện 100 - 300 kA. Biết năng lượng điện lý thuyết tiêu thụ để sản xuất khối lượng m~Al~ (gam) là A = (U × m~Al~ × F) / (9 × 3,6 × 10^6) (kWh). Cho các nhận định:",
        "image": "p10_img2_xref845_811x212.jpeg",
        "items": [
            {"id": "a", "text": "Hỗn hợp chất điện li nóng chảy có khối lượng riêng nhỏ hơn nhôm lỏng nên nổi lên bề mặt.", "ans": "Đ"},
            {"id": "b", "text": "Năng lượng điện tiêu thụ theo lý thuyết để điều chế 1 kg Al ở hiệu điện thế 5 V xấp xỉ 15 kWh.", "ans": "Đ"},
            {"id": "c", "text": "Khí thoát ra ở cực dương chỉ gồm duy nhất oxygen tinh khiết.", "ans": "S"},
            {"id": "d", "text": "Các khối than chì cực dương không bị tiêu hao trong quá trình điện phân.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* D~cryolite~ ≈ 2,1 g/cm^3^ < D~Al lỏng~ ≈ 2,3 g/cm^3^.\\n"
            "- b) *ĐÚNG:* Với m~Al~ = 1000 g: A = (5 × 1000 × 96500) / (9 × 3,6 × 10^6) ≈ 14,89 kWh ≈ 15 kWh.\\n"
            "- c) *SAI:* O~2~ đốt cháy than chì tạo hỗn hợp CO~2~ và CO.\\n"
            "- d) *SAI:* Than chì bị đốt cháy mòn liên tục."
        )
    },
    {
        "id": "TH_C11",
        "num": 11,
        "type": "tf",
        "content": "Trong quá trình sản xuất nhôm bằng phương pháp Hall-Héroult, khí giải phóng ở anode chứa 70% CO~2~ và 30% CO về thể tích. Xét quá trình sản xuất 5,4 tấn nhôm:",
        "items": [
            {"id": "a", "text": "Số mol nhôm lỏng thu được ở cathode là 200 000 mol.", "ans": "Đ"},
            {"id": "b", "text": "Số mol electron trao đổi trong quá trình điện phân trên là 600 000 mol.", "ans": "Đ"},
            {"id": "c", "text": "Khối lượng than chì anode bị oxi hóa lớn hơn 2,0 tấn.", "ans": "Đ"},
            {"id": "d", "text": "Khí CO sinh ra ở anode là khí độc gây nguy hiểm cho công nhân nếu rò rỉ.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* n~Al~ = 5,4 × 10^6 / 27 = 200 000 mol = 0,2 Mmol.\\n"
            "- b) *ĐÚNG:* n~e~ = 3 × 0,2 = 0,6 Mmol = 600 000 mol.\\n"
            "- c) *ĐÚNG:* n~O~ = 0,3 Mmol. Gọi khí: 7x mol CO~2~ và 3x mol CO → n~O~ = 2(7x) + 3x = 17x = 0,3 → x = 0,3 / 17 Mmol.\\n"
            "  n~C~ = 10x = 3 / 17 ≈ 0,1765 Mmol → m~C~ = 0,1765 × 12 ≈ 2,12 tấn > 2,0 tấn.\\n"
            "- d) *ĐÚNG:* Khí CO có độc tính rất cao."
        )
    },
    {
        "id": "TH_C12",
        "num": 12,
        "type": "tf",
        "content": "Tiến hành thí nghiệm điện phân dung dịch CuSO~4~ với các điện cực bằng đồng (Cu). Cho các phát biểu sau:",
        "image": "p12_img1_xref1081_222x163.jpeg",
        "items": [
            {"id": "a", "text": "Khối lượng tấm cathode tăng lên đúng bằng khối lượng tấm anode giảm đi.", "ans": "Đ"},
            {"id": "b", "text": "Màu xanh của dung dịch CuSO~4~ nhạt dần theo thời gian.", "ans": "S"},
            {"id": "c", "text": "Tại cathode ion Cu^2+^ bị khử thành kim loại Cu.", "ans": "Đ"},
            {"id": "d", "text": "Tại anode kim loại Cu bị oxi hóa thành ion Cu^2+^.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Lượng Cu tan ở anode bằng lượng Cu bám vào cathode.\\n"
            "- b) *SAI:* Nồng độ Cu^2+^ không đổi nên màu xanh dung dịch được giữ nguyên.\\n"
            "- c) *ĐÚNG:* Cu^2+^ + 2e → Cu.\\n"
            "- d) *ĐÚNG:* Cu → Cu^2+^ + 2e."
        )
    },
    {
        "id": "TH_C13",
        "num": 13,
        "type": "tf",
        "content": "Điện phân 200 mL dung dịch CuSO~4~ 1M với các điện cực trơ than chì bằng dòng điện I = 5 A. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Ở cực dương (anode) có bọt khí không màu O~2~ thoát ra.", "ans": "Đ"},
            {"id": "b", "text": "Ở cực âm (cathode) có lớp kim loại màu đỏ gạch bám vào.", "ans": "Đ"},
            {"id": "c", "text": "Màu xanh của dung dịch CuSO~4~ đậm dần sau quá trình điện phân.", "ans": "S"},
            {"id": "d", "text": "Dung dịch sau điện phân có khả năng làm quỳ tím hóa đỏ.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* 2H~2~O → O~2~ + 4H^+^ + 4e.\\n"
            "- b) *ĐÚNG:* Cu^2+^ + 2e → Cu (màu đỏ bám cathode).\\n"
            "- c) *SAI:* Cu^2+^ bị điện phân giảm dần nên màu xanh nhạt dần.\\n"
            "- d) *ĐÚNG:* Sinh ra H~2~SO~4~ (H^+^) làm quỳ tím chuyển đỏ."
        )
    },
    {
        "id": "TH_C14",
        "num": 14,
        "type": "tf",
        "content": "Tiến hành thí nghiệm điện phân dung dịch CuSO~4~ theo các bước:\\n• Bước 1: Rót dung dịch CuSO~4~ vào cốc thủy tinh.\\n• Bước 2: Nhúng thanh đồng vào cực dương và chìa khóa thép vào cực âm.\\n• Bước 3: Đóng mạch điện và cho dòng điện một chiều chạy qua.\\nCho các nhận định:",
        "image": "p12_img2_xref1082_997x280.jpeg",
        "items": [
            {"id": "a", "text": "Đây là mô hình ứng dụng điện phân để mạ đồng lên chìa khóa thép.", "ans": "Đ"},
            {"id": "b", "text": "Sau một thời gian chiếc chìa khóa được phủ một lớp kim loại màu đỏ bám chắc.", "ans": "Đ"},
            {"id": "c", "text": "Khối lượng thanh đồng cực dương không đổi trong suốt thí nghiệm.", "ans": "S"},
            {"id": "d", "text": "Nếu đổi chiều dòng điện thì lớp đồng trên chìa khóa sẽ bị tan ra.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Mạ đồng cho chìa khóa.\\n"
            "- b) *ĐÚNG:* Lớp Cu màu đỏ bám đều lên chìa khóa.\\n"
            "- c) *SAI:* Thanh đồng anode tan dần nên khối lượng giảm.\\n"
            "- d) *ĐÚNG:* Khi đổi chiều chìa khóa thành cực dương thì Cu trên chìa khóa bị oxi hóa tan thành Cu^2+^."
        )
    },
    {
        "id": "TH_C15",
        "num": 15,
        "type": "tf",
        "content": "Xét mô hình mạ bạc cho một chiếc nhẫn đồng trong dung dịch AgNO~3~:",
        "items": [
            {"id": "a", "text": "Chiếc nhẫn đồng phải được nối với cực âm (cathode).", "ans": "Đ"},
            {"id": "b", "text": "Thanh bạc nguyên chất được nối với cực dương (anode).", "ans": "Đ"},
            {"id": "c", "text": "Tại bề mặt chiếc nhẫn xảy ra quá trình: Ag^+^ + 1e → Ag.", "ans": "Đ"},
            {"id": "d", "text": "Trong quá trình mạ, nồng độ ion Ag^+^ trong dung dịch giảm mạnh về 0.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Vật cần mạ là cathode (-).\\n"
            "- b) *ĐÚNG:* Kim loại mạ là anode (+).\\n"
            "- c) *ĐÚNG:* Cation Ag^+^ nhận e kết tủa phủ lên nhẫn.\\n"
            "- d) *SAI:* Do dùng anode Ag tan nên nồng độ Ag^+^ duy trì ổn định."
        )
    },
    {
        "id": "TH_C16",
        "num": 16,
        "type": "tf",
        "content": "Quá trình sản xuất nhôm theo công nghệ Hall-Héroult từ quặng bauxite:\\n• Bauxite thô được làm sạch tạp chất Fe~2~O~3~, SiO~2~ bằng dung dịch NaOH theo quy trình Bayer thu được Al~2~O~3~.\\n• Điện phân Al~2~O~3~ nóng chảy trong cryolite ở ~950°C.\\nCho các nhận định:",
        "image": "p12_img3_xref1083_764x402.png",
        "items": [
            {"id": "a", "text": "Tạp chất Fe~2~O~3~ không tan trong NaOH đặc nên bị loại bỏ dưới dạng bùn đỏ.", "ans": "Đ"},
            {"id": "b", "text": "Al~2~O~3~ tan trong dung dịch NaOH tạo thành muối phức aluminate.", "ans": "Đ"},
            {"id": "c", "text": "Cryolite có công thức là Na~3~AlF~6~.", "ans": "Đ"},
            {"id": "d", "text": "Toàn bộ quặng bauxite thô có thể cho trực tiếp vào bể điện phân mà không cần tinh chế.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Fe~2~O~3~ là oxide bazo không tác dụng với NaOH nên nằm lại trong bùn đỏ.\\n"
            "- b) *ĐÚNG:* Al~2~O~3~ + 2NaOH + 3H~2~O → 2Na[Al(OH)~4~].\\n"
            "- c) *ĐÚNG:* Cryolite là trisodium hexafluoroaluminate Na~3~AlF~6~.\\n"
            "- d) *SAI:* Bắt buộc phải tinh chế trước để loại Fe, Si, nếu không sản phẩm nhôm sẽ bị lẫn tạp chất và phá hủy bể điện phân."
        )
    },
    {
        "id": "TH_C17",
        "num": 17,
        "type": "tf",
        "content": "Thí nghiệm điện phân dung dịch copper(II) sulfate với các điện cực trơ graphite được mô phỏng như hình vẽ. Sau thời gian 386 giây điện phân với dòng điện I = 5 A, người ta quan sát thấy các hiện tượng:",
        "image": "p13_img1_xref1159_831x276.png",
        "items": [
            {"id": "a", "text": "Tại cực dương xuất hiện bọt khí O~2~ không màu.", "ans": "Đ"},
            {"id": "b", "text": "Tại cực âm có lớp kim loại đồng màu đỏ bám vào.", "ans": "Đ"},
            {"id": "c", "text": "Khối lượng đồng bám trên cathode sau 386 giây xấp xỉ 0,64 gam.", "ans": "Đ"},
            {"id": "d", "text": "Dung dịch sau điện phân có giá trị pH lớn hơn 7.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Anode trơ: 2H~2~O → O~2~ + 4H^+^ + 4e.\\n"
            "- b) *ĐÚNG:* Cathode: Cu^2+^ + 2e → Cu.\\n"
            "- c) *ĐÚNG:* n~e~ = (5 × 386) / 96500 = 0,02 mol → n~Cu~ = 0,01 mol → m~Cu~ = 0,01 × 64 = 0,64 gam.\\n"
            "- d) *SAI:* Sinh ra acid H~2~SO~4~ nên dung dịch có pH < 7."
        )
    },
    {
        "id": "TH_C18",
        "num": 18,
        "type": "tf",
        "content": "Trong quá trình tinh chế đồng bằng phương pháp điện phân, xét các yếu tố sau:",
        "items": [
            {"id": "a", "text": "Kim loại sắt (Fe) và kẽm (Zn) trong đồng thô sẽ tan vào dung dịch dưới dạng Fe^2+^ và Zn^2+^.", "ans": "Đ"},
            {"id": "b", "text": "Các ion Fe^2+^ và Zn^2+^ trong dung dịch sẽ bị khử bám lên tấm cathode cùng với đồng.", "ans": "S"},
            {"id": "c", "text": "Kim loại vàng (Au) và bạc (Ag) không bị tan, lắng xuống tạo thành bùn anode.", "ans": "Đ"},
            {"id": "d", "text": "Đồng thu được ở cathode có độ tinh khiết cao lên đến 99,99%.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Fe, Zn hoạt động hơn Cu nên bị oxi hóa tan thành cation.\\n"
            "- b) *SAI:* E° của Zn^2+^ và Fe^2+^ rất âm so với Cu^2+^ nên chúng không bị khử ở cathode trong điều kiện thế khống chế.\\n"
            "- c) *ĐÚNG:* Ag, Au trơ hơn Cu nên không tan, rơi xuống đáy tạo bùn anode.\\n"
            "- d) *ĐÚNG:* Độ tinh khiết Cu thu được đạt chuẩn thương mại > 99,99%."
        )
    },
    {
        "id": "TH_C19",
        "num": 19,
        "type": "tf",
        "content": "Cho một pin điện hóa Daniel gồm cặp điện cực Zn nhúng trong ZnSO~4~ 1M và Cu nhúng trong CuSO~4~ 1M nối với nhau bằng cầu muối KCl. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Thanh kẽm Zn đóng vai trò là cực âm (anode) của pin.", "ans": "Đ"},
            {"id": "b", "text": "Thanh đồng Cu đóng vai trò là cực dương (cathode) của pin.", "ans": "Đ"},
            {"id": "c", "text": "Chiều dòng điện trong mạch ngoài đi từ thanh Zn sang thanh Cu.", "ans": "S"},
            {"id": "d", "text": "Suất điện động chuẩn của pin Daniel xấp xỉ 1,10 V.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Trong pin điện hóa Zn có E° = -0,76 V nhỏ hơn nên là cực âm (anode), xảy ra sự oxi hóa Zn → Zn^2+^ + 2e.\\n"
            "- b) *ĐÚNG:* Cu có E° = +0,34 V là cực dương (cathode), xảy ra sự khử Cu^2+^ + 2e → Cu.\\n"
            "- c) *SAI:* Electron chạy từ Zn sang Cu, do đó chiều quy ước của dòng điện ngược lại là từ Cu sang Zn.\\n"
            "- d) *ĐÚNG:* E°~pin~ = E°(Cu^2+^/Cu) - E°(Zn^2+^/Zn) = +0,34 - (-0,76) = 1,10 V."
        )
    },
    {
        "id": "TH_C20",
        "num": 20,
        "type": "tf",
        "content": "Cho các nhận định về ứng dụng của phương pháp điện phân trong công nghiệp:",
        "items": [
            {"id": "a", "text": "Dùng điện phân nóng chảy để điều chế các kim loại hoạt động hóa học mạnh như Na, Mg, Al.", "ans": "Đ"},
            {"id": "b", "text": "Mạ điện giúp bảo vệ kim loại nền khỏi bị ăn mòn và nâng cao giá trị thẩm mỹ của sản phẩm.", "ans": "Đ"},
            {"id": "c", "text": "Tinh chế điện phân cho phép tách và thu hồi các kim loại quý hiếm (vàng, bạc, platin) từ bùn anode.", "ans": "Đ"},
            {"id": "d", "text": "Có thể điều chế kim loại nhôm bằng phương pháp điện phân dung dịch AlCl~3~ trong nước.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Kim loại có thế khử âm mạnh bắt buộc điện phân nóng chảy.\\n"
            "- b) *ĐÚNG:* Mục đích cốt lõi của mạ điện.\\n"
            "- c) *ĐÚNG:* Bùn anode chứa Au, Ag, Pt là nguồn thu hồi kim loại quý giá trị cao.\\n"
            "- d) *SAI:* Điện phân dung dịch AlCl~3~ thì H~2~O bị khử trước sinh khí H~2~ và Al(OH)~3~ kết tủa, không thu được Al."
        )
    },
    {
        "id": "TH_C21",
        "num": 21,
        "type": "tf",
        "content": "Tiến hành điện phân dung dịch CuSO~4~ trong ống thủy tinh hình chữ U với hai điện cực than chì. Nhỏ thêm vài giọt dung dịch phenolphtalein hoặc quỳ tím vào ống. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Ở nhánh nối với cực dương sinh ra khí O~2~ làm dung dịch có tính acid mạnh (quỳ tím chuyển đỏ).", "ans": "Đ"},
            {"id": "b", "text": "Ở nhánh nối với cực âm xuất hiện lớp đồng bám trên bề mặt than chì.", "ans": "Đ"},
            {"id": "c", "text": "Màu xanh lam của dung dịch CuSO~4~ ở cả hai nhánh nhạt dần.", "ans": "Đ"},
            {"id": "d", "text": "Nếu tiếp tục điện phân lâu sau khi Cu^2+^ hết thì ở cathode sẽ bắt đầu thoát khí H~2~.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Anode sinh H^+^: 2H~2~O → O~2~ + 4H^+^ + 4e làm quỳ chuyển đỏ.\\n"
            "- b) *ĐÚNG:* Cathode khử Cu^2+^ thành Cu đỏ.\\n"
            "- c) *ĐÚNG:* Ion Cu^2+^ khuếch tán và bị khử dần nên màu xanh mất dần.\\n"
            "- d) *ĐÚNG:* Sau khi Cu^2+^ hết, nước bị khử ở cathode: 2H~2~O + 2e → H~2~ + 2OH^-^."
        )
    },
    {
        "id": "TH_C22",
        "num": 22,
        "type": "tf",
        "content": "Trong công nghiệp, nhôm được điều chế bằng cách điện phân nóng chảy hỗn hợp alumina (Al~2~O~3~) và cryolite (Na~3~AlF~6~) theo quy trình Hall-Héroult. Cho các phát biểu sau:",
        "image": "p14_img1_xref1297_631x232.jpeg",
        "items": [
            {"id": "a", "text": "Nhiệt độ nóng chảy của hỗn hợp điện phân khoảng 950°C, thấp hơn nhiều so với alumina nguyên chất (> 2000°C).", "ans": "Đ"},
            {"id": "b", "text": "Cryolite làm tăng độ dẫn điện của khối nóng chảy.", "ans": "Đ"},
            {"id": "c", "text": "Cực dương làm bằng than chì bị ăn mòn liên tục bởi khí oxygen sinh ra ở nhiệt độ cao.", "ans": "Đ"},
            {"id": "d", "text": "Nhôm lỏng thu được nhẹ hơn hỗn hợp chất điện li nên nổi lên trên mặt bồn điện phân.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Hạ từ 2050°C xuống 950°C.\\n"
            "- b) *ĐÚNG:* Phân li tạo nhiều ion linh động.\\n"
            "- c) *ĐÚNG:* C + O~2~ → CO~2~ và 2C + O~2~ → 2CO làm mòn cực dương.\\n"
            "- d) *SAI:* Nhôm lỏng nặng hơn chìm xuống đáy thùng."
        )
    },
    {
        "id": "TH_C23",
        "num": 23,
        "type": "tf",
        "content": "Xét sơ đồ tinh chế đồng bằng phương pháp điện phân. Anode là khối đồng thô (chứa Zn, Fe, Ag, Au), cathode là tấm đồng tinh khiết mỏng, chất điện li là CuSO~4~ có pha H~2~SO~4~. Cho các nhận định:",
        "image": "p14_img2_xref1298_897x324.png",
        "items": [
            {"id": "a", "text": "Ở cực dương luôn xảy ra quá trình oxi hóa các kim loại Cu, Zn, Fe.", "ans": "Đ"},
            {"id": "b", "text": "Thanh đồng thô anode mòn dần, thanh đồng tinh khiết cathode dày dần lên.", "ans": "Đ"},
            {"id": "c", "text": "Ion Ag^+^ và Au^3+^ sẽ bám lên cathode cùng với đồng.", "ans": "S"},
            {"id": "d", "text": "Quá trình này cho phép sản xuất đồng thương phẩm có độ sạch 99,99%.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Các kim loại có tính khử mạnh hơn hoặc bằng Cu đều bị oxi hóa.\\n"
            "- b) *ĐÚNG:* Cu tan ở anode và bám vào cathode.\\n"
            "- c) *SAI:* Ag và Au không bị oxi hóa nên không tan thành ion, chúng lắng xuống tạo bùn anode.\\n"
            "- d) *ĐÚNG:* Mục đích hàng đầu của tinh chế điện phân đồng."
        )
    },
    {
        "id": "TH_C24",
        "num": 24,
        "type": "tf",
        "content": "Tiến hành điện phân 100 mL dung dịch CuCl~2~ 1M với các điện cực trơ than chì. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Tại cathode ion Cu^2+^ bị khử thành kim loại Cu.", "ans": "Đ"},
            {"id": "b", "text": "Tại anode ion Cl^-^ bị oxi hóa giải phóng khí Cl~2~ màu vàng lục có mùi hắc.", "ans": "Đ"},
            {"id": "c", "text": "Phương trình phản ứng điện phân tổng quát là: CuCl~2~ → Cu + Cl~2~.", "ans": "Đ"},
            {"id": "d", "text": "Sau khi CuCl~2~ bị điện phân hết, nếu tiếp tục điện phân thì nước sẽ không bị điện phân.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Cu^2+^ + 2e → Cu.\\n"
            "- b) *ĐÚNG:* 2Cl^-^ → Cl~2~ + 2e.\\n"
            "- c) *ĐÚNG:* CuCl~2~ → Cu + Cl~2~.\\n"
            "- d) *SAI:* Khi hết CuCl~2~, nước sẽ tiếp tục bị điện phân tạo H~2~ ở cathode và O~2~ ở anode."
        )
    },
    {
        "id": "TH_C25",
        "num": 25,
        "type": "tf",
        "content": "Học sinh A làm thí nghiệm mạ 0,33 gam kẽm (Zn) lên chiếc chìa khóa bằng sắt bằng cách điện phân dung dịch ZnSO~4~ với dòng điện không đổi I = 1,5 A. Cho biết Zn = 65, F = 96 500 C/mol. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Chìa khóa sắt được nối với cực âm (cathode) của nguồn điện.", "ans": "Đ"},
            {"id": "b", "text": "Cực dương (anode) phải làm bằng thanh kẽm (Zn) nguyên chất.", "ans": "Đ"},
            {"id": "c", "text": "Số mol electron cần trao đổi để mạ 0,33 gam Zn là khoảng 0,01015 mol.", "ans": "Đ"},
            {"id": "d", "text": "Thời gian điện phân lý thuyết (hiệu suất 100%) là khoảng 653 giây.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Vật cần mạ là cathode (-).\\n"
            "- b) *ĐÚNG:* Dùng anode Zn tan để duy trì ion Zn^2+^.\\n"
            "- c) *ĐÚNG:* n~Zn~ = 0,33 / 65 ≈ 0,005077 mol → n~e~ = 2 × 0,005077 ≈ 0,01015 mol.\\n"
            "- d) *ĐÚNG:* q = 0,01015 × 96500 ≈ 979,5 C → t = 979,5 / 1,5 ≈ 653 giây (≈ 10,9 phút)."
        )
    },
    {
        "id": "TH_C26",
        "num": 26,
        "type": "tf",
        "content": "Quá trình điện phân để mạ đồng (Cu) lên một chiếc chìa khóa làm bằng thép không gỉ được mô tả như hình vẽ: cathode là chìa khóa, anode là thanh đồng thô, dung dịch điện phân là CuSO~4~. Cho các nhận định sau:",
        "image": "p15_img1_xref1400_389x340.jpeg",
        "items": [
            {"id": "a", "text": "Tại anode đồng thô bị oxi hóa thành ion Cu^2+^ tan vào dung dịch.", "ans": "Đ"},
            {"id": "b", "text": "Tại cathode các ion Cu^2+^ bị khử bám đều lên bề mặt chìa khóa.", "ans": "Đ"},
            {"id": "c", "text": "Nồng độ ion Cu^2+^ trong dung dịch giảm mạnh sau quá trình mạ.", "ans": "S"},
            {"id": "d", "text": "Nếu sử dụng dòng điện quá mạnh, lớp đồng mạ có thể bị xốp và dễ bong tróc.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Hiện tượng anode tan Cu → Cu^2+^ + 2e.\\n"
            "- b) *ĐÚNG:* Cu^2+^ + 2e → Cu.\\n"
            "- c) *SAI:* Dương cực tan duy trì nồng độ Cu^2+^ ổn định.\\n"
            "- d) *ĐÚNG:* Dòng điện quá lớn làm tốc độ kết tủa quá nhanh sinh tinh thể xốp, bám không chắc."
        )
    },
    {
        "id": "TH_C27",
        "num": 27,
        "type": "tf",
        "content": "Cho bộ pin điện hóa Zn - Cu được đo bằng vôn kế điện tử. Dung dịch ở hai cốc là ZnSO~4~ 1M và CuSO~4~ 1M, hai điện cực Zn và Cu nối qua cầu muối chứa KNO~3~ bão hòa. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Kim vôn kế chỉ hiệu điện thế chuẩn của pin xấp xỉ 1,10 V.", "ans": "Đ"},
            {"id": "b", "text": "Khối lượng thanh Zn giảm dần và khối lượng thanh Cu tăng dần khi pin hoạt động.", "ans": "Đ"},
            {"id": "c", "text": "Cation K^+^ từ cầu muối sẽ di chuyển về phía cốc chứa dung dịch ZnSO~4~.", "ans": "S"},
            {"id": "d", "text": "Phản ứng tự diễn biến trong pin là: Zn + Cu^2+^ → Zn^2+^ + Cu.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* E°~pin~ = +0,34 - (-0,76) = 1,10 V.\\n"
            "- b) *ĐÚNG:* Zn bị oxi hóa tan ra, Cu^2+^ khử bám vào cực Cu.\\n"
            "- c) *SAI:* Cốc Zn sinh thêm cation Zn^2+^ nên anion NO~3~^-^ di chuyển vào để trung hòa điện tích; cation K^+^ di chuyển sang cốc Cu (nơi Cu^2+^ bị giảm).\\n"
            "- d) *ĐÚNG:* Zn khử ion Cu^2+^."
        )
    },
    {
        "id": "TH_C28",
        "num": 28,
        "type": "tf",
        "content": "Điện phân dung dịch Cu(NO~3~)~2~ với điện cực trơ graphite bằng dòng điện I = 2 A. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Tại cực âm (cathode) thu được kim loại Cu màu đỏ.", "ans": "Đ"},
            {"id": "b", "text": "Tại cực dương (anode) thu được khí O~2~.", "ans": "Đ"},
            {"id": "c", "text": "Dung dịch sau điện phân có chứa acid nitric HNO~3~.", "ans": "Đ"},
            {"id": "d", "text": "Phương trình điện phân là: 2Cu(NO~3~)~2~ + 2H~2~O → 2Cu + O~2~ + 4HNO~3~.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Cathode khử Cu^2+^.\\n"
            "- b) *ĐÚNG:* Anode oxi hóa nước sinh O~2~.\\n"
            "- c) *ĐÚNG:* NO~3~^-^ kết hợp H^+^ tạo HNO~3~.\\n"
            "- d) *ĐÚNG:* Phương trình hoàn toàn chính xác."
        )
    },
    {
        "id": "TH_C29",
        "num": 29,
        "type": "tf",
        "content": "Trong một bình điện phân dung dịch CuSO~4~ có anode bằng Cu và cathode bằng thép. Để mạ một lớp Cu dày 0,02 mm lên tấm thép có diện tích 50 cm^2^. Cho biết D~Cu~ = 8,96 g/cm^3^, Cu = 64, F = 96 500 C/mol. Cho các phát biểu:",
        "items": [
            {"id": "a", "text": "Thể tích lớp đồng cần mạ là 0,10 cm^3^.", "ans": "Đ"},
            {"id": "b", "text": "Khối lượng đồng cần kết tủa là 0,896 gam.", "ans": "Đ"},
            {"id": "c", "text": "Số mol electron trao đổi là 0,028 mol.", "ans": "Đ"},
            {"id": "d", "text": "Điện lượng cần cung cấp theo lý thuyết là 2702 C.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* h = 0,02 mm = 0,002 cm → V = 50 × 0,002 = 0,10 cm^3^.\\n"
            "- b) *ĐÚNG:* m = V × D = 0,10 × 8,96 = 0,896 gam.\\n"
            "- c) *ĐÚNG:* n~Cu~ = 0,896 / 64 = 0,014 mol → n~e~ = 2 × 0,014 = 0,028 mol.\\n"
            "- d) *ĐÚNG:* q = 0,028 × 96500 = 2702 C."
        )
    },
    {
        "id": "TH_C30",
        "num": 30,
        "type": "tf",
        "content": "Một huy chương kim loại hình tròn bán kính 3 cm được mạ vàng (Au) trong dung dịch chứa phức vàng Au(CN)~2~^-^ với dòng điện I = 0,5 A. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Huy chương cần mạ được nối với cực âm (cathode).", "ans": "Đ"},
            {"id": "b", "text": "Cực dương (anode) được làm bằng vàng tinh khiết.", "ans": "Đ"},
            {"id": "c", "text": "Nếu thời gian mạ tăng gấp đôi thì độ dày lớp mạ vàng cũng tăng xấp xỉ gấp đôi.", "ans": "Đ"},
            {"id": "d", "text": "Khi mạ vàng từ phức chất xyanua, lớp mạ thu được thường mịn, bóng và bám chắc hơn mạ từ muối Au^3+^ đơn giản.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Vật cần mạ là cathode.\\n"
            "- b) *ĐÚNG:* Anode tan bằng vàng.\\n"
            "- c) *ĐÚNG:* m = (A.I.t)/(n.F) tỉ lệ thuận với thời gian t.\\n"
            "- d) *ĐÚNG:* Phức chất phóng thích ion tự do chậm và đều, giúp tinh thể kim loại kết tinh mịn và bám chặt."
        )
    },
    {
        "id": "TH_C31",
        "num": 31,
        "type": "tf",
        "content": "Tiến hành mạ đồng lên một tấm huy chương hình tròn trong dung dịch CuSO~4~ với anode bằng đồng. Cho các nhận định:",
        "image": "p16_img1_xref1506_870x239.jpeg",
        "items": [
            {"id": "a", "text": "Trong quá trình mạ điện tại điện cực anode không có khí thoát ra.", "ans": "Đ"},
            {"id": "b", "text": "Thanh đồng là cực dương, tấm huy chương được mạ đóng vai trò là cực âm.", "ans": "Đ"},
            {"id": "c", "text": "Nồng độ ion Cu^2+^ trong dung dịch điện phân không đổi trong suốt quá trình mạ.", "ans": "Đ"},
            {"id": "d", "text": "Khối lượng tấm huy chương tăng lên bao nhiêu thì khối lượng thanh đồng giảm đi bấy nhiêu.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Hiện tượng dương cực tan không sinh khí ở anode.\\n"
            "- b) *ĐÚNG:* Anode (+) là thanh Cu, cathode (-) là huy chương.\\n"
            "- c) *ĐÚNG:* Cu tan bằng Cu bám.\\n"
            "- d) *ĐÚNG:* Với hiệu suất 100%, lượng Cu tan ra bằng lượng Cu kết tủa."
        )
    },
    {
        "id": "TH_C32",
        "num": 32,
        "type": "tf",
        "content": "Bộ dụng cụ Hofmann điện phân nước sử dụng dung dịch NaOH 0,05M. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Ở cathode nước nhận electron tạo khí H~2~ và ion OH^-^.", "ans": "Đ"},
            {"id": "b", "text": "Ở anode nước nhường electron tạo khí O~2~ và ion H^+^.", "ans": "Đ"},
            {"id": "c", "text": "Nồng độ NaOH trong dung dịch tăng dần trong quá trình điện phân do nước bị mất đi.", "ans": "Đ"},
            {"id": "d", "text": "Sau một thời gian điện phân, pH của dung dịch giảm mạnh.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* 2H~2~O + 2e → H~2~ + 2OH^-^.\\n"
            "- b) *ĐÚNG:* 2H~2~O → O~2~ + 4H^+^ + 4e.\\n"
            "- c) *ĐÚNG:* Thực chất là điện phân H~2~O, lượng dung môi giảm nên nồng độ chất tan NaOH tăng lên.\\n"
            "- d) *SAI:* Nồng độ NaOH tăng làm nồng độ OH^-^ tăng nên pH của dung dịch tăng nhẹ, không giảm."
        )
    },
    {
        "id": "TH_C33",
        "num": 33,
        "type": "tf",
        "content": "Xét quá trình tinh chế đồng thô (chứa 96% Cu, 2% Zn, 1% Fe, 0,5% Ag, 0,5% Au về khối lượng) bằng phương pháp điện phân:",
        "items": [
            {"id": "a", "text": "Khi anode tan ra, các kim loại Zn và Fe tan vào dung dịch dưới dạng cation Zn^2+^ và Fe^2+^.", "ans": "Đ"},
            {"id": "b", "text": "Các kim loại Ag và Au lắng xuống đáy bể tạo bùn anode.", "ans": "Đ"},
            {"id": "c", "text": "Ở cathode chỉ có ion Cu^2+^ bị khử thành Cu bám vào tấm đồng tinh khiết.", "ans": "Đ"},
            {"id": "d", "text": "Hiệu suất sử dụng dòng điện để kết tủa Cu tại cathode luôn đạt tuyệt đối 100% so với tổng khối lượng anode tan ra.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Zn, Fe có thế khử âm hơn Cu nên bị oxi hóa trước.\\n"
            "- b) *ĐÚNG:* Ag, Au không bị oxi hóa tạo bùn anode.\\n"
            "- c) *ĐÚNG:* Ở thế khống chế chỉ Cu^2+^ bị khử.\\n"
            "- d) *SAI:* Một phần dòng điện bị tiêu hao cho việc hòa tan Zn, Fe ở anode, nên khối lượng Cu bám ở cathode không bằng 100% tổng khối lượng anode tan."
        )
    },
    {
        "id": "TH_C34",
        "num": 34,
        "type": "tf",
        "content": "Tiến hành thí nghiệm theo các bước sau:\\n• Bước 1: Rót từ từ dung dịch CuSO~4~ 0,5M vào ống thủy tinh hình chữ U đến khi mực nước dâng lên cách miệng ống 2 cm thì dừng lại.\\n• Bước 2: Đậy miệng ống bên trái bằng nút cao su có gắn điện cực than chì.\\n• Bước 3: Đậy miệng ống bên phải bằng nút cao su có gắn điện cực than chì và một ống dẫn khí.\\n• Bước 4: Nối điện cực bên trái với cực âm (cathode) và nối điện cực bên phải với cực dương (anode) của nguồn điện một chiều (hiệu điện thế 6 V, dòng điện 1 A).\\nCho các nhận định:",
        "items": [
            {"id": "a", "text": "Ở cathode, ion Cu^2+^ bị oxi hóa thành kim loại Cu.", "ans": "S"},
            {"id": "b", "text": "Dung dịch sau điện phân có giá trị pH < 7.", "ans": "Đ"},
            {"id": "c", "text": "Nếu thay dung dịch CuSO~4~ 0,5M bằng dung dịch CuCl~2~ 0,5M thì cathode và anode vẫn thu được các sản phẩm giống hệt thí nghiệm ban đầu.", "ans": "S"},
            {"id": "d", "text": "Đến khi ở cathode bắt đầu có bọt khí thoát ra thì màu xanh của dung dịch đã nhạt hẳn (Cu^2+^ gần như đã bị điện phân hết).", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Ở cathode xảy ra quá trình KHỬ: Cu^2+^ + 2e → Cu.\\n"
            "- b) *ĐÚNG:* Anode sinh H^+^: 2H~2~O → O~2~ + 4H^+^ + 4e làm pH < 7.\\n"
            "- c) *SAI:* Nếu dùng CuCl~2~ thì ở anode sinh khí Cl~2~ màu vàng lục, không phải O~2~.\\n"
            "- d) *ĐÚNG:* Khi Cu^2+^ bị khử hết, nước bắt đầu bị khử ở cathode sinh khí H~2~, lúc này dung dịch đã mất màu xanh."
        )
    },
    {
        "id": "TH_C35",
        "num": 35,
        "type": "tf",
        "content": "Vàng hồng là hợp kim của Au, Ag, Cu. Trên thị trường có các loại: 10K (41,7% Au; 20% Ag; 38,3% Cu); 14K (58,5% Au; 21,5% Ag; 20% Cu); 18K (75% Au; 10% Ag; 15% Cu). Một nhóm học sinh muốn nâng cao hàm lượng Au, Ag từ mảnh nhẫn vàng hồng 10K bằng phương pháp điện phân với giả thuyết: 'Nếu Cu tan hết thì khối lượng mảnh nhẫn không giảm nữa'.\\n- Cân ban đầu mảnh nhẫn 10K (1,125 gam) và thanh Cu tinh khiết (2,255 gam).\\n- Nối mảnh nhẫn với một điện cực và thanh Cu với cực còn lại, nhúng vào dung dịch CuSO~4~.\\n- Điện phân một thời gian, sấy khô rồi cân lại thấy mảnh nhẫn nặng 0,515 gam và thanh Cu nặng 2,740 gam.\\nCho các nhận định:",
        "items": [
            {"id": "a", "text": "Mảnh nhẫn phải được nối với cực dương (anode), thanh Cu tinh khiết nối với cực âm (cathode).", "ans": "Đ"},
            {"id": "b", "text": "Nồng độ ion Cu^2+^ trong dung dịch không đổi trong suốt quá trình điện phân.", "ans": "Đ"},
            {"id": "c", "text": "Khối lượng Cu ban đầu trong mảnh nhẫn 10K là: 1,125 × 38,3% = 0,431 gam.", "ans": "Đ"},
            {"id": "d", "text": "Độ giảm khối lượng của mảnh nhẫn là 1,125 - 0,515 = 0,610 gam lớn hơn khối lượng Cu (0,431 gam) chứng tỏ đã có một phần bột vàng, bạc bị rơi rụng xuống đáy bình tạo bùn anode.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Mảnh nhẫn là anode (+) để Cu tan ra; thanh Cu là cathode (-) để thu Cu.\\n"
            "- b) *ĐÚNG:* Cu tan ở anode bằng Cu bám ở cathode nên nồng độ Cu^2+^ không đổi.\\n"
            "- c) *ĐÚNG:* m~Cu trong nhẫn~ = 1,125 × 38,3% = 0,4309 g ≈ 0,431 gam.\\n"
            "- d) *ĐÚNG:* Khối lượng nhẫn giảm 0,610 g > 0,431 g Cu. Điều này là do khi khung mạng lưới đồng bị hòa tan, các hạt vàng, bạc siêu mịn không còn chỗ bám sẽ rụng xuống đáy bể tạo bùn anode. Do đó giả thuyết ban đầu của học sinh là sai."
        )
    },
    {
        "id": "TH_C36",
        "num": 36,
        "type": "tf",
        "content": "Vàng tây 18K chứa 75,00% Au, phần còn lại là Ni và Ag. Một nhà kim hoàn kiểm tra hàm lượng Au bằng phương pháp điện phân:\\n- Lấy 1 chỉ vàng tây 18K (3,75 gam) làm điện cực nối với cực dương.\\n- Điện cực âm là thanh than chì tinh khiết nặng 2,00 gam.\\n- Nhúng vào 100 mL dung dịch Au(NO~3~)~3~ 0,2M. Kẹp nối 2 điện cực đều làm bằng than chì trơ.\\n- Điện phân ở hiệu điện thế thích hợp cho đến khi chỉ vàng tan hết thì dừng lại, sấy khô cân lại thanh than chì thấy nặng 4,85 gam.\\nBiết trong quá trình điện phân thể tích dung dịch không đổi, toàn bộ Au sinh ra bám hết lên điện cực âm và không có cặn rơi xuống đáy. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Có thể thay kẹp điện cực than chì bằng dây đồng Cu để thực hiện thí nghiệm.", "ans": "S"},
            {"id": "b", "text": "Tại cathode xảy ra bán phản ứng khử: Au^3+^ + 3e → Au.", "ans": "Đ"},
            {"id": "c", "text": "Khối lượng Au bám lên thanh than chì là 4,85 - 2,00 = 2,85 gam.", "ans": "Đ"},
            {"id": "d", "text": "Hàm lượng vàng trong mẫu vàng tây trên đạt: (2,85 / 3,75) × 100% = 76,0% (vượt tiêu chuẩn 75%).", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Nếu dùng dây Cu thì Cu sẽ tác dụng với dung dịch Au(NO~3~)~3~ (3Cu + 2Au^3+^ → 3Cu^2+^ + 2Au) và tan ra ở cực dương, làm sai lệch kết quả thí nghiệm.\\n"
            "- b) *ĐÚNG:* Cation Au^3+^ nhận electron kết tủa bám vào cathode than chì.\\n"
            "- c) *ĐÚNG:* m~Au bám~ = 4,85 - 2,00 = 2,85 gam.\\n"
            "- d) *ĐÚNG:* %Au = (2,85 / 3,75) × 100% = 76,0% > 75,00% (vượt tiêu chuẩn 18K)."
        )
    }
]
'''

with open(DATA_FILE, "a", encoding="utf-8") as f:
    f.write(tong_hop_content)

print("Appended TONG_HOP_MCQ_TF (Câu 1 đến Câu 36).")
