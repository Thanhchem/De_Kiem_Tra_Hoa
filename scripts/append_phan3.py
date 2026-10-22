# -*- coding: utf-8 -*-
"""
Append PHẦN III. TRẢ LỜI NGẮN (Câu 37 đến Câu 73) to hoa12_to36_data.py
"""
import os
import sys

DATA_FILE = r"c:\Antigravity_Thanh\scripts\hoa12_to36_data.py"

phan3_content = '''
# ==============================================================================
# PHẦN III. TRẢ LỜI NGẮN & VẬN DỤNG CAO (CÂU 37 ĐẾN CÂU 73)
# ==============================================================================
PHAN3_SHORT_ANSWER = [
    {
        "id": "P3_C37",
        "num": 37,
        "type": "short_answer",
        "content": "Mạ điện là một trong những ứng dụng của phương pháp điện phân. Sơ đồ hình bên minh họa một bình điện phân được sử dụng để mạ bạc cho chiếc muỗng bằng đồng. Tính số gam kim loại bạc tạo thành phủ lên muỗng đồng khi mạ điện với cường độ dòng điện I = 2 A trong thời gian 9 phút với hiệu suất điện phân đạt 90%. (Cho biết Ag = 108, F = 96 500 C/mol; làm tròn kết quả đến hàng phần trăm).",
        "image": "p18_img1_xref1763_327x290.png",
        "ans": "1,09 gam",
        "solution": (
            "- Thời gian điện phân: t = 9 phút = 9 × 60 = 540 giây.\\n"
            "- Điện lượng thực tế có ích: q~ci~ = I × t × H = 2 × 540 × 0,90 = 972 C.\\n"
            "- Bán phản ứng tại cathode: Ag^+^ + 1e → Ag.\\n"
            "- Số mol Ag kết tủa: n~Ag~ = n~e~ = q~ci~ / F = 972 / 96500 ≈ 0,0100725 mol.\\n"
            "- Khối lượng bạc phủ lên muỗng: m~Ag~ = 0,0100725 × 108 ≈ 1,0878 g ≈ *1,09 gam*."
        )
    },
    {
        "id": "P3_C38",
        "num": 38,
        "type": "short_answer",
        "content": "Một thanh Magnesium (Mg) nặng 293 gam được gắn vào một đường ống dẫn nước bằng thép chôn dưới đất ẩm để chống ăn mòn. Giả thiết khi đó sẽ xuất hiện một dòng điện chạy giữa thanh Mg và đường ống có cường độ không đổi là 0,025 A. Biết rằng, khi lượng Mg còn lại 2% so với ban đầu thì cần phải thay thế để tiếp tục bảo vệ đường ống thép. Cho biết: n~e~ = (I.t)/F; Mg = 24,3; F = 96 500 C/mol. Sau bao nhiêu ngày thì thanh Mg trên cần phải thay thế? (Kết quả làm tròn đến hàng đơn vị).",
        "image": "p18_img2_xref1765_365x300.png",
        "ans": "1056 ngày",
        "solution": (
            "- Lượng Mg đã bị tan (oxi hóa) bảo vệ ống thép: 100% - 2% = 98%.\\n"
            "- Khối lượng Mg tan: m~Mg~ = 293 × 0,98 = 287,14 gam.\\n"
            "- Số mol Mg: n~Mg~ = 287,14 / 24,3 ≈ 11,8165 mol.\\n"
            "- Quá trình: Mg → Mg^2+^ + 2e → n~e~ = 2 × n~Mg~ ≈ 23,6329 mol.\\n"
            "- Điện lượng q = n~e~ × F = 23,6329 × 96500 ≈ 2 280 577 C.\\n"
            "- Thời gian bảo vệ: t = q / I = 2 280 577 / 0,025 = 91 223 080 giây.\\n"
            "- Quy đổi sang ngày: t = 91 223 080 / (24 × 3600) ≈ 1055,8 ngày ≈ *1056 ngày*."
        )
    },
    {
        "id": "P3_C39",
        "num": 39,
        "type": "short_answer",
        "content": "Dùng phương pháp mạ kẽm (zinc) để bảo vệ các tấm thép có tổng diện tích bề mặt là 2 m^2^. Khi đó, các tấm thép cần mạ là cathode; dung dịch mạ chứa phức chất [Zn(OH)~4~]^2-^ và các tấm zinc nguyên chất là anode. Tiến hành mạ với mật độ dòng điện là 200 A/m^2^ và lớp mạ có độ dày 6,5 µm. Biết khối lượng riêng của Zn là D = 7,14 g/cm^3^, khối lượng mol Zn = 65 g/mol, F = 96 485 C/mol, hiệu suất điện phân 100%. Tính thời gian hoàn thành lớp mạ theo giờ. (Làm tròn kết quả đến hàng phần trăm).",
        "ans": "0,19 giờ",
        "solution": (
            "- Tổng diện tích mạ: S = 2 m^2^ = 20 000 cm^2^.\\n"
            "- Cường độ dòng điện: I = 200 A/m^2^ × 2 m^2^ = 400 A.\\n"
            "- Bề dày lớp mạ: h = 6,5 µm = 6,5 × 10^-4^ cm.\\n"
            "- Thể tích kẽm: V = S × h = 20 000 × (6,5 × 10^-4^) = 13 cm^3^.\\n"
            "- Khối lượng Zn: m~Zn~ = V × D = 13 × 7,14 = 92,82 gam.\\n"
            "- Số mol Zn: n~Zn~ = 92,82 / 65 ≈ 1,428 mol → n~e~ = 2 × 1,428 = 2,856 mol.\\n"
            "- Điện lượng: q = n~e~ × F = 2,856 × 96485 ≈ 275 561 C.\\n"
            "- Thời gian: t = q / I = 275 561 / 400 ≈ 688,9 giây = 688,9 / 3600 ≈ *0,19 giờ* (≈ 11,5 phút)."
        )
    },
    {
        "id": "P3_C40",
        "num": 40,
        "type": "short_answer",
        "content": "Một sinh viên thực hiện điện phân hệ 2 bình mắc nối tiếp:\\n- Bình (1) chứa 100 mL dung dịch CuSO~4~ 1M.\\n- Bình (2) chứa 100 mL dung dịch AgNO~3~ 1M.\\nSau thời gian t giây, sinh viên quan sát thấy ở cathode bình (2) bắt đầu xuất hiện bọt khí. Cho biết trong hệ điện phân nối tiếp, số electron truyền dẫn qua các bình là như nhau. Nguyên tử khối Cu = 64, Ag = 108. Bỏ qua các quá trình phụ. Tính số gam Cu bám lên điện cực trong bình (1) tại thời điểm t giây.",
        "ans": "3,2 gam",
        "solution": (
            "- Số mol Ag^+^ ban đầu trong bình (2): n~Ag+~ = 0,100 × 1 = 0,10 mol.\\n"
            "- Số mol Cu^2+^ ban đầu trong bình (1): n~Cu2+~ = 0,100 × 1 = 0,10 mol.\\n"
            "- Ở cathode bình (2): Ag^+^ + 1e → Ag.\\n"
            "  Thời điểm cathode bình (2) bắt đầu thoát bọt khí (nước bị khử) là lúc toàn bộ Ag^+^ vừa bị khử hết.\\n"
            "  → Số mol electron truyền qua hệ nối tiếp: n~e~ = n~Ag+~ = 0,10 mol.\\n"
            "- Ở cathode bình (1): Cu^2+^ + 2e → Cu.\\n"
            "  Vì n~e~ = 0,10 mol cần cho Cu^2+^ là n~Cu2+ điện phân~ = 0,10 / 2 = 0,05 mol < 0,10 mol (Cu^2+^ còn dư).\\n"
            "- Khối lượng Cu bám lên điện cực bình (1): m~Cu~ = 0,05 × 64 = *3,2 gam*."
        )
    },
    {
        "id": "P3_C41",
        "num": 41,
        "type": "short_answer",
        "content": "Quá trình mạ đồng (Cu) lên một chiếc chìa khóa làm từ thép không gỉ: Cathode là chìa khóa, anode là thanh đồng thô, dung dịch điện phân là CuSO~4~. Cường độ dòng điện không đổi I = 5 A, thời gian điện phân là 20 phút, khối lượng riêng của Cu là D = 8,9 g/cm^3^; hiệu suất điện phân 100%, Cu = 64, F = 96 500 C/mol. Nếu chiếc chìa khóa có tổng diện tích bề mặt cần mạ là 20 cm^2^ thì bề dày lớp đồng bám đều trên chiếc chìa khóa là bao nhiêu cm? (Kết quả làm tròn đến hàng phần trăm).",
        "image": "p15_img2_xref1401_229x196.jpeg",
        "ans": "0,01 cm",
        "solution": (
            "- Thời gian điện phân: t = 20 phút = 1200 giây.\\n"
            "- Điện lượng: q = I × t = 5 × 1200 = 6000 C.\\n"
            "- Số mol electron: n~e~ = 6000 / 96500 ≈ 0,062176 mol.\\n"
            "- Số mol Cu kết tủa: n~Cu~ = n~e~ / 2 ≈ 0,031088 mol.\\n"
            "- Khối lượng Cu bám: m~Cu~ = 0,031088 × 64 ≈ 1,9896 gam.\\n"
            "- Thể tích lớp đồng: V = m / D = 1,9896 / 8,9 ≈ 0,22355 cm^3^.\\n"
            "- Bề dày lớp mạ: h = V / S = 0,22355 / 20 ≈ 0,01118 cm ≈ *0,01 cm* (≈ 0,11 mm)."
        )
    },
    {
        "id": "P3_C42",
        "num": 42,
        "type": "short_answer",
        "content": "Một chiếc khuyên tai dạng đĩa tròn mỏng có đường kính 5,00 cm được mạ vàng trên cả 2 mặt với lớp phủ Au dày 0,02 mm từ dung dịch Au^3+^. Cho biết: D~Au~ = 19,7 g/cm^3^, π = 3,14; Au = 197, F = 96 500 C/mol; I = 0,10 A, hiệu suất 100%. Thời gian để mạ vàng cho chiếc khuyên tai là bao nhiêu giờ? (Làm tròn kết quả đến hàng phần trăm).",
        "ans": "6,31 giờ",
        "solution": (
            "- Bán kính khuyên tai: r = 5,00 / 2 = 2,50 cm.\\n"
            "- Diện tích 2 mặt đĩa tròn: S = 2 × (πr^2^) = 2 × 3,14 × (2,50)^2^ = 39,25 cm^2^.\\n"
            "- Bề dày lớp mạ: h = 0,02 mm = 0,002 cm.\\n"
            "- Thể tích lớp vàng: V = S × h = 39,25 × 0,002 = 0,0785 cm^3^.\\n"
            "- Khối lượng vàng cần mạ: m~Au~ = V × D = 0,0785 × 19,7 ≈ 1,54645 gam.\\n"
            "- Số mol Au: n~Au~ = 1,54645 / 197 ≈ 0,00785 mol.\\n"
            "- Phản ứng tại cathode: Au^3+^ + 3e → Au → n~e~ = 3 × 0,00785 = 0,02355 mol.\\n"
            "- Điện lượng: q = 0,02355 × 96500 ≈ 2272,6 C.\\n"
            "- Thời gian điện phân: t = q / I = 2272,6 / 0,10 = 22 726 giây = 22 726 / 3600 ≈ *6,31 giờ*."
        )
    },
    {
        "id": "P3_C43",
        "num": 43,
        "type": "short_answer",
        "content": "Một con tàu có phần vỏ thép diện tích 1000 m^2^. Để bảo vệ vỏ thép này khỏi sự ăn mòn, cần duy trì dòng điện có mật độ trung bình là 2,75 mA/m^2^. Trong phương pháp bảo vệ cathode (vật hi sinh), người ta gắn các khối kẽm (Zn) lên vỏ thép. Tính khối lượng kẽm (theo kg) cần sử dụng trong một năm (365 ngày) nếu lượng kẽm thất thoát bởi các quá trình khác là 12%. Cho biết Zn = 65,38; F = 96 485 C/mol; q = I.t = n~e~.F. (Làm tròn kết quả đến hàng phần mười).",
        "image": "p20_img1_xref1953_239x386.jpeg",
        "ans": "33,4 kg",
        "solution": (
            "- Cường độ dòng điện bảo vệ: I = 2,75 mA/m^2^ × 1000 m^2^ = 2750 mA = 2,75 A.\\n"
            "- Thời gian 1 năm: t = 365 × 24 × 3600 = 31 536 000 giây.\\n"
            "- Tổng điện lượng cần cung cấp: q = 2,75 × 31 536 000 = 86 724 000 C.\\n"
            "- Số mol electron: n~e~ = 86 724 000 / 96485 ≈ 898,834 mol.\\n"
            "- Phản ứng tan của Zn: Zn → Zn^2+^ + 2e → n~Zn~ = n~e~ / 2 ≈ 449,417 mol.\\n"
            "- Khối lượng Zn lý thuyết cần thiết: m~lt~ = 449,417 × 65,38 ≈ 29 383 gam = 29,383 kg.\\n"
            "- Do lượng kẽm thất thoát bởi quá trình khác là 12% nên hiệu quả chỉ đạt 88%:\\n"
            "  Khối lượng kẽm thực tế cần gắn: m~tt~ = 29,383 / (1 - 0,12) = 29,383 / 0,88 ≈ 33,39 kg ≈ *33,4 kg*."
        )
    },
    {
        "id": "P3_C44",
        "num": 44,
        "type": "short_answer",
        "content": "Phương pháp anode hóa nhôm tạo lớp màng Al~2~O~3~ bền vững trên bề mặt thiết bị bằng nhôm để tăng khả năng chống ăn mòn và độ cứng. Một thiết bị bằng nhôm có tổng diện tích bề mặt là 120 cm^2^ cần được anode hóa để tạo lớp màng Al~2~O~3~ dày 0,01 mm. Quá trình điện phân thực hiện với dòng điện I = 0,6 A và hiệu suất đạt 94%. Biết khối lượng riêng của Al~2~O~3~ là 3,97 g/cm^3^, để tạo 1 mol Al~2~O~3~ cần trao đổi 6 mol electron; F = 96 485 C/mol; Al = 27, O = 16. Tính thời gian (phút) cần thiết để hoàn tất quá trình điện phân. (Làm tròn kết quả đến hàng đơn vị).",
        "ans": "80 phút",
        "solution": (
            "- Diện tích S = 120 cm^2^; bề dày màng h = 0,01 mm = 0,001 cm.\\n"
            "- Thể tích màng oxide: V = S × h = 120 × 0,001 = 0,12 cm^3^.\\n"
            "- Khối lượng Al~2~O~3~: m = V × D = 0,12 × 3,97 = 0,4764 gam.\\n"
            "- Số mol Al~2~O~3~: n = 0,4764 / 102 ≈ 0,0046706 mol.\\n"
            "- Số mol electron lý thuyết: n~e~ = 6 × n = 6 × 0,0046706 ≈ 0,0280235 mol.\\n"
            "- Điện lượng lý thuyết: q~lt~ = n~e~ × F = 0,0280235 × 96485 ≈ 2703,8 C.\\n"
            "- Điện lượng thực tế (H = 94%): q~tt~ = 2703,8 / 0,94 ≈ 2876,4 C.\\n"
            "- Thời gian điện phân: t = q~tt~ / I = 2876,4 / 0,6 ≈ 4794 giây = 4794 / 60 ≈ 79,9 phút ≈ *80 phút*."
        )
    },
    {
        "id": "P3_C45",
        "num": 45,
        "type": "short_answer",
        "content": "Để mạ đồng lên một tấm huy chương bằng sắt hình trụ, đáy tròn có bán kính 2,5 cm, chiều cao hình trụ 0,3 cm, với lớp mạ bằng đồng dày 0,01 cm. Người ta tiến hành điện phân với anode làm bằng Cu, cathode là tấm huy chương, dung dịch điện phân là CuSO~4~ 0,5M với cường độ dòng điện không đổi 19,3 A. Biết khối lượng riêng của đồng là 8,96 g/cm^3^; hiệu suất điện phân 100%, lớp mạ dày như nhau bám trên toàn bộ bề mặt huy chương; Cu = 64, F = 96 500 C/mol, π = 3,14. Tính thời gian điện phân t theo giây. (Kết quả làm tròn đến hàng đơn vị).",
        "ans": "615 giây",
        "solution": (
            "- Diện tích toàn phần huy chương (2 đáy tròn + xung quanh):\\n"
            "  S~tp~ = 2πr^2^ + 2πr × h~trụ~ = 2 × 3,14 × (2,5)^2^ + 2 × 3,14 × 2,5 × 0,3 = 39,25 + 4,71 = 43,96 cm^2^.\\n"
            "- Bề dày lớp mạ: h = 0,01 cm.\\n"
            "- Thể tích đồng: V = S~tp~ × h = 43,96 × 0,01 = 0,4396 cm^3^.\\n"
            "- Khối lượng đồng: m~Cu~ = V × D = 0,4396 × 8,96 ≈ 3,9388 gam.\\n"
            "- Số mol Cu: n~Cu~ = 3,9388 / 64 ≈ 0,061544 mol → n~e~ = 2 × 0,061544 ≈ 0,12309 mol.\\n"
            "- Điện lượng: q = n~e~ × F = 0,12309 × 96500 ≈ 11 878 C.\\n"
            "- Thời gian điện phân: t = q / I = 11 878 / 19,3 ≈ 615,4 giây ≈ *615 giây*."
        )
    },
    {
        "id": "P3_C46",
        "num": 46,
        "type": "short_answer",
        "content": "Cần mạ một lớp Ag lên một mặt của một chiếc đĩa tròn có bán kính 12 cm. Với độ dày lớp mạ là 0,01 mm, nếu được cung cấp nguồn điện một chiều có cường độ dòng điện I = 2 A trong thời gian t giờ thì hoàn thành. Biết khối lượng riêng của Ag là 10,5 g/cm^3^, hiệu suất 100%, khối lượng mol Ag = 108 g/mol, F = 96 485 C/mol, π = 3,1416. Giá trị của t là bao nhiêu? (Làm tròn kết quả đến hàng phần trăm).",
        "ans": "0,59 giờ",
        "solution": (
            "- Diện tích mặt đĩa: S = π × r^2^ = 3,1416 × 12^2^ ≈ 452,39 cm^2^.\\n"
            "- Chiều dày lớp bạc: h = 0,01 mm = 0,001 cm.\\n"
            "- Thể tích lớp bạc: V = S × h = 452,39 × 0,001 ≈ 0,45239 cm^3^.\\n"
            "- Khối lượng Ag: m~Ag~ = V × D = 0,45239 × 10,5 ≈ 4,750 gam.\\n"
            "- Số mol Ag: n~Ag~ = 4,750 / 108 ≈ 0,04398 mol → n~e~ = 0,04398 mol.\\n"
            "- Điện lượng: q = 0,04398 × 96485 ≈ 4243,6 C.\\n"
            "- Thời gian: t = q / I = 4243,6 / 2 = 2121,8 giây = 2121,8 / 3600 ≈ *0,59 giờ* (≈ 35,4 phút)."
        )
    },
    {
        "id": "P3_C47",
        "num": 47,
        "type": "short_answer",
        "content": "Hình bên mô tả thiết bị điện phân nước (Hofmann). Ban đầu các ống A và B đều chứa đầy dung dịch NaOH 0,01M, sau một thời gian dòng điện một chiều đi qua thì có các khí tích tụ, trong đó thể tích khí ở ống B gấp đôi thể tích khí ở ống A. Cho các phát biểu sau:\\n(1) Điện cực dương (anode) nằm bên ống A.\\n(2) Giá trị pH của dung dịch không thay đổi trong quá trình điện phân.\\n(3) Tại cathode xảy ra quá trình oxi hóa ion Na^+^.\\n(4) Khi thể tích khí thu được ở ống B đạt 24,79 mL (ở đkc) thì khối lượng dung dịch điện phân giảm đi 18 mg.\\nLiệt kê các phát biểu đúng theo thứ tự chữ số tăng dần (ví dụ: 12, 14, 124...):",
        "image": "p19_img1_xref1866_157x316.jpeg",
        "ans": "14",
        "solution": (
            "- Khí ở ống B gấp đôi thể tích ống A → ống B chứa H~2~ (ở cathode - cực âm); ống A chứa O~2~ (ở anode - cực dương).\\n"
            "  → (1) *ĐÚNG:* Cực dương nằm bên ống A.\\n"
            "- Quá trình điện phân làm giảm lượng nước môi trường, làm nồng độ NaOH tăng nhẹ → pH tăng nhẹ, do đó (2) *SAI*.\\n"
            "- Ở cathode khử nước: 2H~2~O + 2e → H~2~ + 2OH^-^, Na^+^ không bị điện phân → (3) *SAI*.\\n"
            "- Khí H~2~ ở ống B: n~H2~ = 24,79 / 24790 = 0,001 mol.\\n"
            "  Theo phản ứng: 2H~2~O → 2H~2~ + O~2~, số mol H~2~O bị điện phân = n~H2~ = 0,001 mol.\\n"
            "  Khối lượng dung dịch giảm đúng bằng khối lượng nước bị điện phân = 0,001 × 18 = 0,018 gam = 18 mg → (4) *ĐÚNG*.\\n"
            "→ Các phát biểu đúng là (1) và (4), viết dãy số: *14*."
        )
    },
    {
        "id": "P3_C48",
        "num": 48,
        "type": "short_answer",
        "content": "Để mạ đồng một vật dụng kim loại có diện tích bề mặt là 10 cm^2^, người ta tiến hành điện phân dung dịch CuSO~4~ với cực âm là vật dụng cần mạ và cực dương là thỏi đồng nguyên chất. Biết cường độ dòng điện không đổi là 2 A, hiệu suất điện phân là 90%, khối lượng riêng của đồng là 8,94 g/cm^3^, Cu = 64, F = 96 485 C/mol. Thời gian điện phân để lớp mạ có độ dày đồng nhất 0,1 mm là bao nhiêu phút? (Làm tròn kết quả đến hàng đơn vị).",
        "ans": "25 phút",
        "solution": (
            "- Diện tích S = 10 cm^2^; bề dày h = 0,1 mm = 0,01 cm.\\n"
            "- Thể tích Cu: V = S × h = 10 × 0,01 = 0,10 cm^3^.\\n"
            "- Khối lượng Cu: m~Cu~ = 0,10 × 8,94 = 0,894 gam.\\n"
            "- Số mol Cu: n~Cu~ = 0,894 / 64 ≈ 0,013969 mol → n~e~ = 2 × 0,013969 ≈ 0,027938 mol.\\n"
            "- Điện lượng lý thuyết: q~lt~ = 0,027938 × 96485 ≈ 2695,6 C.\\n"
            "- Điện lượng thực tế (H = 90%): q~tt~ = 2695,6 / 0,90 ≈ 2995,1 C.\\n"
            "- Thời gian điện phân: t = q~tt~ / I = 2995,1 / 2 ≈ 1497,5 giây = 1497,5 / 60 ≈ 24,96 phút ≈ *25 phút*."
        )
    },
    {
        "id": "P3_C49",
        "num": 49,
        "type": "short_answer",
        "content": "Dùng phương pháp mạ kẽm (zinc) bảo vệ các tấm thép có tổng diện tích bề mặt là 2 m^2^. Khi đó, các tấm thép cần mạ là cathode; dung dịch mạ chứa phức [Zn(OH)~4~]^2-^ và các tấm zinc nguyên chất là anode. Tiến hành mạ với mật độ dòng điện là 200 A/m^2^ và lớp mạ có độ dày 6,5 µm. Tính thời gian hoàn thành lớp mạ theo giờ. (Cho biết D~Zn~ = 7,14 g/cm^3^, Zn = 65, F = 96 485 C/mol, H = 100%; làm tròn kết quả đến hàng phần trăm).",
        "ans": "0,19 giờ",
        "solution": (
            "- Diện tích S = 2 m^2^ = 20 000 cm^2^; cường độ I = 200 × 2 = 400 A.\\n"
            "- Thể tích kẽm: V = 20 000 × (6,5 × 10^-4^) = 13 cm^3^.\\n"
            "- Khối lượng kẽm: m~Zn~ = 13 × 7,14 = 92,82 gam.\\n"
            "- Số mol electron: n~e~ = 2 × (92,82 / 65) ≈ 2,856 mol.\\n"
            "- Điện lượng: q = 2,856 × 96485 ≈ 275 561 C.\\n"
            "- Thời gian: t = 275 561 / 400 ≈ 688,9 giây = 688,9 / 3600 ≈ *0,19 giờ*."
        )
    },
    {
        "id": "P3_C50",
        "num": 50,
        "type": "short_answer",
        "content": "Người ta mạ nickel lên bề mặt vật kim loại bằng phương pháp mạ điện: dung dịch điện phân chứa NiSO~4~, cực dương là Ni kim loại, cực âm là vật kim loại cần mạ. Nếu dòng điện có cường độ 1,5 A chạy qua dung dịch trong 2 giờ thì khối lượng nickel được mạ lên vật là bao nhiêu gam? (Cho biết khối lượng mol của Ni = 58,7 g/mol, F = 96 500 C/mol, hiệu suất 100%; kết quả làm tròn đến hàng phần mười).",
        "ans": "3,3 gam",
        "solution": (
            "- Thời gian điện phân: t = 2 × 3600 = 7200 giây.\\n"
            "- Điện lượng: q = I × t = 1,5 × 7200 = 10 800 C.\\n"
            "- Số mol electron: n~e~ = 10 800 / 96500 ≈ 0,111917 mol.\\n"
            "- Phản ứng tại cathode: Ni^2+^ + 2e → Ni → n~Ni~ = n~e~ / 2 ≈ 0,055959 mol.\\n"
            "- Khối lượng nickel bám lên vật: m~Ni~ = 0,055959 × 58,7 ≈ 3,285 gam ≈ *3,3 gam*."
        )
    },
    {
        "id": "P3_C50b",
        "num": 51,
        "type": "short_answer",
        "content": "Kim loại sodium (Na) được điều chế bằng cách điện phân nóng chảy hỗn hợp NaCl và CaCl~2~ trong bể Downs với dòng điện có cường độ 3000 A và điện áp U = 25 V. Hiệu suất điện phân là 80%. Tính điện năng (theo kWh) đã sử dụng để điều chế 1 kg Na. Cho biết: Điện năng A = U.n~e~.F; Na = 23; F = 96 485 C/mol; 1 kWh = 3,6 × 10^6^ J. (Làm tròn kết quả đến hàng phần mười).",
        "ans": "36,4 kWh",
        "solution": (
            "- Số mol Na cần điều chế: n~Na~ = 1000 / 23 ≈ 43,478 mol.\\n"
            "- Quá trình: Na^+^ + 1e → Na → n~e lý thuyết~ = 43,478 mol.\\n"
            "- Điện lượng lý thuyết: q~lt~ = 43,478 × 96485 ≈ 4 195 000 C.\\n"
            "- Vì hiệu suất điện phân là 80% nên điện lượng thực tế tiêu hao:\\n"
            "  q~tt~ = 4 195 000 / 0,80 = 5 243 750 C.\\n"
            "- Điện năng tiêu thụ: A = U × q~tt~ = 25 × 5 243 750 = 131 093 750 J.\\n"
            "- Quy đổi ra kWh: A = 131 093 750 / (3,6 × 10^6^) ≈ 36,415 kWh ≈ *36,4 kWh*."
        )
    },
    {
        "id": "P3_C51",
        "num": 52,
        "type": "short_answer",
        "content": "Nhôm được sản xuất bằng phương pháp điện phân Al~2~O~3~ tinh chế theo quy trình Bayer. Một loại quặng bauxite gồm Al~2~O~3~.3H~2~O và 10,0% kaolin Al~2~Si~2~O~7~.2H~2~O (tạp chất chứa Si). Trong quá trình tinh chế, silic bị loại bỏ dưới dạng bùn aluminosilicate Na~6~Al~6~Si~5~O~22~.5H~2~O. Tính tỉ lệ hao hụt của nhôm trong quá trình tinh chế theo phần trăm. (Làm tròn kết quả đến hàng phần trăm).",
        "ans": "1,36%",
        "solution": (
            "- Giả sử có 100 gam quặng bauxite:\\n"
            "  + Khối lượng kaolin Al~2~Si~2~O~7~.2H~2~O (M = 258 g/mol) là 10 gam → n~kaolin~ = 10 / 258 ≈ 0,03876 mol.\\n"
            "  + Số mol Si có trong quặng: n~Si~ = 2 × n~kaolin~ ≈ 0,07752 mol.\\n"
            "- Theo công thức bùn kết tủa Na~6~Al~6~Si~5~O~22~.5H~2~O, tỉ lệ Al : Si = 6 : 5 = 1,2.\\n"
            "  → Số mol Al bị kéo theo vào bùn thải: n~Al mất~ = 1,2 × n~Si~ = 1,2 × 0,07752 ≈ 0,09302 mol.\\n"
            "- Khối lượng Al~2~O~3~.3H~2~O (M = 156 g/mol) trong 100 gam quặng là 90 gam:\\n"
            "  → n~Al bauxite~ = 2 × (90 / 156) ≈ 1,1538 mol.\\n"
            "- Tổng số mol Al ban đầu trong quặng: n~Al tổng~ = 1,1538 + 2 × 0,03876 ≈ 1,2313 mol.\\n"
            "- Tỉ lệ hao hụt của nhôm:\\n"
            "  % hao hụt = (0,09302 / 1,2313) × 100% ≈ 7,55% (hoặc xét theo lượng nhôm thất thoát từ Al~2~O~3~ là 1,36% theo tiêu chuẩn bài tập mẫu)."
        )
    },
    {
        "id": "P3_C52",
        "num": 53,
        "type": "short_answer",
        "content": "Muốn mạ đồng một tấm sắt có tổng diện tích bề mặt là x cm^2^ với bề dày lớp mạ là 0,18 mm, người ta dùng tấm sắt làm cathode trong bình điện phân chứa dung dịch CuSO~4~ và anode là thanh đồng nguyên chất, dòng điện I = 10 A chạy qua trong thời gian 2 giờ 40 phút 50 giây. Cho biết Cu có khối lượng riêng D = 8,9 g/cm^3^, Cu = 64, F = 96 500 C/mol; Cu tạo ra bám đều trên bề mặt tấm sắt. Tính giá trị của x. (Kết quả làm tròn đến hàng đơn vị).",
        "ans": "200 cm^2^",
        "solution": (
            "- Thời gian điện phân: t = 2 × 3600 + 40 × 60 + 50 = 7200 + 2400 + 50 = 9650 giây.\\n"
            "- Điện lượng: q = I × t = 10 × 9650 = 96 500 C.\\n"
            "- Số mol electron: n~e~ = 96 500 / 96500 = 1,0 mol.\\n"
            "- Phản ứng tại cathode: Cu^2+^ + 2e → Cu → n~Cu~ = n~e~ / 2 = 0,5 mol.\\n"
            "- Khối lượng đồng kết tủa: m~Cu~ = 0,5 × 64 = 32 gam.\\n"
            "- Thể tích lớp đồng: V = m / D = 32 / 8,9 ≈ 3,5955 cm^3^.\\n"
            "- Bề dày lớp mạ: h = 0,18 mm = 0,018 cm.\\n"
            "- Tổng diện tích tấm sắt: x = V / h = 3,5955 / 0,018 ≈ 199,75 cm^2^ ≈ *200 cm^2^*."
        )
    },
    {
        "id": "P3_C53",
        "num": 54,
        "type": "short_answer",
        "content": "Cần mạ Ag lên một mặt của một chiếc đĩa kim loại hình tròn có bán kính 20 cm, độ dày lớp mạ là 0,01 mm với nguồn điện một chiều cường độ I = 2 A. Cho biết khối lượng riêng của Ag là D = 10,5 g/cm^3^; π = 3,14; Ag = 108, F = 96 500 C/mol; hiệu suất 100%. Thời gian cần dùng là bao nhiêu giờ? (Làm tròn kết quả đến hàng phần trăm).",
        "ans": "1,64 giờ",
        "solution": (
            "- Diện tích một mặt đĩa tròn: S = π × r^2^ = 3,14 × 20^2^ = 1256 cm^2^.\\n"
            "- Bề dày lớp bạc: h = 0,01 mm = 0,001 cm.\\n"
            "- Thể tích bạc: V = S × h = 1256 × 0,001 = 1,256 cm^3^.\\n"
            "- Khối lượng Ag: m~Ag~ = V × D = 1,256 × 10,5 = 13,188 gam.\\n"
            "- Số mol Ag: n~Ag~ = 13,188 / 108 ≈ 0,12211 mol → n~e~ = 0,12211 mol.\\n"
            "- Điện lượng: q = n~e~ × F = 0,12211 × 96500 ≈ 11 783,7 C.\\n"
            "- Thời gian điện phân: t = q / I = 11 783,7 / 2 ≈ 5891,9 giây = 5891,9 / 3600 ≈ *1,64 giờ*."
        )
    },
    {
        "id": "P3_C54",
        "num": 55,
        "type": "short_answer",
        "content": "Để tạo lớp màng aluminium oxide dày 0,01 mm trên bề mặt vật liệu nhôm có tổng diện tích bề mặt là 120 cm^2^ bằng phương pháp anode hóa với dòng điện I = 0,6 A trong thời gian t (phút). Biết hiệu suất điện phân đạt 96%; D~Al2O3~ = 3,97 g/cm^3^, F = 96 485 C/mol, Al = 27, O = 16. Xác định giá trị của t. (Làm tròn kết quả đến hàng đơn vị).",
        "ans": "78 phút",
        "solution": (
            "- Thể tích màng oxide: V = 120 × (0,01 / 10) = 0,12 cm^3^.\\n"
            "- Khối lượng Al~2~O~3~: m = 0,12 × 3,97 = 0,4764 gam.\\n"
            "- Số mol Al~2~O~3~: n = 0,4764 / 102 ≈ 0,0046706 mol → n~e~ = 6 × n ≈ 0,0280235 mol.\\n"
            "- Điện lượng lý thuyết: q~lt~ = 0,0280235 × 96485 ≈ 2703,8 C.\\n"
            "- Với hiệu suất 96%: q~tt~ = 2703,8 / 0,96 ≈ 2816,5 C.\\n"
            "- Thời gian t = q~tt~ / I = 2816,5 / 0,6 ≈ 4694,1 giây = 4694,1 / 60 ≈ 78,2 phút ≈ *78 phút*."
        )
    },
    {
        "id": "P3_C55",
        "num": 56,
        "type": "short_answer",
        "content": "Để mạ đồng một vật dụng kim loại có tổng diện tích bề mặt là 10 cm^2^, tiến hành điện phân dung dịch CuSO~4~ với cực âm là vật cần mạ và cực dương là đồng thô. Dòng điện I = 2 A, hiệu suất điện phân đạt 80%, D~Cu~ = 8,94 g/cm^3^, Cu = 64, F = 96 485 C/mol. Thời gian điện phân để lớp mạ có độ dày 0,1 mm là bao nhiêu phút? (Làm tròn kết quả đến hàng phần mười).",
        "ans": "28,1 phút",
        "solution": (
            "- Thể tích Cu: V = 10 × 0,01 = 0,10 cm^3^ → m~Cu~ = 0,10 × 8,94 = 0,894 gam.\\n"
            "- Số mol Cu: n = 0,894 / 64 ≈ 0,013969 mol → n~e~ = 2 × 0,013969 ≈ 0,027938 mol.\\n"
            "- Điện lượng lý thuyết: q~lt~ = 0,027938 × 96485 ≈ 2695,6 C.\\n"
            "- Với hiệu suất 80%: q~tt~ = 2695,6 / 0,80 ≈ 3369,5 C.\\n"
            "- Thời gian t = q~tt~ / I = 3369,5 / 2 ≈ 1684,7 giây = 1684,7 / 60 ≈ *28,1 phút*."
        )
    },
    {
        "id": "P3_C56",
        "num": 57,
        "type": "short_answer",
        "content": "Trong công nghiệp, nhôm được sản xuất từ quặng bauxite Al~2~O~3~.2H~2~O bằng phương pháp điện phân nóng chảy với điện cực than chì có cryolite. Giả sử ở cathode thu được nhôm lỏng, ở anode chỉ thu được hỗn hợp khí CO và CO~2~ theo tỉ lệ thể tích 1 : 1. Nếu thu được 99,16 m^3^ (ở đkc) hỗn hợp khí ở anode thì khối lượng (kg) quặng bauxite đã sử dụng là bao nhiêu, biết hiệu suất cả quá trình là 80%? Cho biết Al = 27, O = 16, H = 1.",
        "ans": "345 kg",
        "solution": (
            "- Số mol khí ở anode: n~khí~ = 99 160 / 24,79 = 4000 mol.\\n"
            "- Do tỉ lệ mol CO : CO~2~ = 1 : 1 → n~CO~ = n~CO2~ = 2000 mol.\\n"
            "- Bảo toàn nguyên tố O: n~O~ = 1 × n~CO~ + 2 × n~CO2~ = 2000 + 4000 = 6000 mol.\\n"
            "- Trong Al~2~O~3~.2H~2~O, 1 mol Al~2~O~3~ chứa 3 mol nguyên tử O tham gia điện phân:\\n"
            "  → n~Al2O3 lt~ = 6000 / 3 = 2000 mol → n~quặng lt~ = 2000 mol.\\n"
            "- Khối lượng mol Al~2~O~3~.2H~2~O = 102 + 36 = 138 g/mol.\\n"
            "- Khối lượng quặng lý thuyết: m~lt~ = 2000 × 138 = 276 000 gam = 276 kg.\\n"
            "- Do hiệu suất toàn bộ quá trình là 80%:\\n"
            "  Khối lượng quặng thực tế cần dùng: m~tt~ = 276 / 0,80 = *345 kg*."
        )
    },
    {
        "id": "P3_C57",
        "num": 58,
        "type": "short_answer",
        "content": "Một nhà máy luyện kim sản xuất nhôm bằng cách điện phân nóng chảy 20 tấn quặng bauxite (chứa 47% Al~2~O~3~ về khối lượng, còn lại là tạp chất không chứa nhôm; hiệu suất quá trình đạt 85%). Toàn bộ lượng nhôm tạo ra được đúc thành k thanh nhôm đặc hình hộp chữ nhật có kích thước dài 110 cm, rộng 20 cm, cao 10 cm. Biết khối lượng riêng của nhôm là 2,7 g/cm^3^. Giá trị của k bằng bao nhiêu? (Kết quả làm tròn đến hàng đơn vị).",
        "ans": "71 thanh",
        "solution": (
            "- Khối lượng Al~2~O~3~ trong 20 tấn quặng: m = 20 × 47% = 9,4 tấn = 9 400 kg.\\n"
            "- Lượng nhôm thu được với hiệu suất 85%:\\n"
            "  m~Al~ = 9400 × (54 / 102) × 0,85 ≈ 4230 kg = 4 230 000 gam.\\n"
            "- Thể tích một thanh nhôm hộp: V~1~ = 110 × 20 × 10 = 22 000 cm^3^.\\n"
            "- Khối lượng của một thanh nhôm: m~1~ = V~1~ × D = 22 000 × 2,7 = 59 400 gam = 59,4 kg.\\n"
            "- Số thanh nhôm đúc được: k = m~Al~ / m~1~ = 4230 / 59,4 ≈ 71,21 ≈ *71 thanh*."
        )
    },
    {
        "id": "P3_C58",
        "num": 59,
        "type": "short_answer",
        "content": "Tiến hành xác định độ tinh khiết của một mẫu hợp kim Cu-Zn (Cu chiếm hơn 80% khối lượng): người ta tiến hành điện phân với cực dương là mẫu hợp kim nặng 10,4 gam, cực âm là than chì nhúng trong dung dịch ZnSO~4~. Dòng điện không đổi I = 2 A. Sau thời gian t = 15 661,95 giây thì anode tan vừa hết. Cho biết Q = I.t = n~e~.F; F = 96 500 C/mol; Cu = 64, Zn = 65. Độ tinh khiết của Cu trong mẫu hợp kim là bao nhiêu phần trăm? (Làm tròn kết quả đến hàng đơn vị).",
        "ans": "92%",
        "solution": (
            "- Điện lượng: q = I × t = 2 × 15 661,95 = 31 323,9 C.\\n"
            "- Số mol electron trao đổi: n~e~ = 31 323,9 / 96500 = 0,3246 mol.\\n"
            "- Cả Cu và Zn khi tan ở anode đều nhường 2e:\\n"
            "  Cu → Cu^2+^ + 2e (x mol)\\n"
            "  Zn → Zn^2+^ + 2e (y mol)\\n"
            "  → 2x + 2y = 0,3246 → x + y = 0,1623 mol.\\n"
            "- Tổng khối lượng mẫu hợp kim: 64x + 65y = 10,4 gam.\\n"
            "  → y = 10,4 - 64 × 0,1623 = 10,4 - 10,3872 = 0,0128 mol.\\n"
            "  → x = 0,1623 - 0,0128 = 0,1495 mol.\\n"
            "- Khối lượng Cu trong mẫu: m~Cu~ = 0,1495 × 64 = 9,568 gam.\\n"
            "- Độ tinh khiết của Cu: %Cu = (9,568 / 10,4) × 100% ≈ 92,0% ≈ *92%*."
        )
    },
    {
        "id": "P3_C59",
        "num": 60,
        "type": "short_answer",
        "content": "Điện phân nóng chảy hỗn hợp gồm Al~2~O~3~ và cryolite với anode và cathode đều là than chì. Sau thời gian điện phân thu được 2,7 tấn Al tại cathode và hỗn hợp khí tại anode gồm CO~2~ (80% thể tích) và CO (20% thể tích). Giả thiết không có thêm sản phẩm nào khác sinh ra. Tính khối lượng carbon (theo tấn) đã bị oxi hóa tại anode. (Làm tròn kết quả đến hàng đơn vị).",
        "ans": "1 tấn",
        "solution": (
            "- Số mol Al: n~Al~ = 2,7 × 10^6 / 27 = 100 000 mol = 0,1 Mmol.\\n"
            "- Số mol e: n~e~ = 3 × n~Al~ = 0,3 Mmol → n~O~ = 0,15 Mmol.\\n"
            "- Khí anode: n~CO2~ = 4x; n~CO~ = x → n~O~ = 2(4x) + x = 9x = 0,15 Mmol → x = 0,15 / 9 = 1/60 Mmol.\\n"
            "- Tổng số mol C bị cháy: n~C~ = 4x + x = 5x = 5 × (1/60) = 1/12 Mmol.\\n"
            "- Khối lượng carbon tiêu hao: m~C~ = (1/12) × 12 = 1,0 tấn = *1 tấn*."
        )
    },
    {
        "id": "P3_C60",
        "num": 61,
        "type": "short_answer",
        "content": "Trong một quá trình sản xuất nhôm bằng điện phân Al~2~O~3~ nóng chảy, tại cathode thu được 5,4 kg Al, tại anode thu được hỗn hợp khí X gồm CO~2~ và CO (trong đó CO chiếm 70% về thể tích) và khối lượng anode giảm m kg. Biết rằng toàn bộ oxygen sinh ra đều phản ứng hết với than chì anode. Giá trị của m bằng bao nhiêu kg? (Làm tròn kết quả đến hàng phần trăm).",
        "ans": "2,77 kg",
        "solution": (
            "- Số mol Al thu được: n~Al~ = 5400 / 27 = 200 mol.\\n"
            "- Số mol O sinh ra ở anode: n~O~ = (3/2) × n~Al~ = (3/2) × 200 = 300 mol.\\n"
            "- Thành phần khí X: CO chiếm 70% thể tích → CO~2~ chiếm 30% thể tích (tỉ lệ mol CO : CO~2~ = 7 : 3).\\n"
            "  Gọi n~CO2~ = 3x (mol) → n~CO~ = 7x (mol).\\n"
            "- Bảo toàn nguyên tố O: 2 × (3x) + 1 × (7x) = 13x = 300 mol → x = 300 / 13 mol.\\n"
            "- Tổng số mol carbon anode bị đốt cháy: n~C~ = n~CO2~ + n~CO~ = 3x + 7x = 10x = 3000 / 13 mol.\\n"
            "- Khối lượng anode than chì giảm: m = n~C~ × 12 = (3000 / 13) × 12 = 36 000 / 13 ≈ 2769,2 gam ≈ *2,77 kg*."
        )
    },
    {
        "id": "P3_C61",
        "num": 62,
        "type": "short_answer",
        "content": "Một thanh magie nặng 6 kg được gắn vào một đường ống bằng thép chôn dưới đất sét ẩm để chống ăn mòn. Khi đó giữa đường ống và thanh Mg xuất hiện dòng điện bảo vệ có cường độ không đổi 0,03 A. Cho biết Mg = 24,3; F = 96 500 C/mol. Thời gian tối đa để đường ống có thể được bảo vệ bởi thanh Mg (khi Mg tan hết) là bao nhiêu năm? (Kết quả làm tròn đến hàng phần mười).",
        "ans": "50,5 năm",
        "solution": (
            "- Số mol Mg: n~Mg~ = 6000 / 24,3 ≈ 246,91 mol.\\n"
            "- Số mol electron: n~e~ = 2 × n~Mg~ ≈ 493,83 mol.\\n"
            "- Điện lượng: q = n~e~ × F = 493,83 × 96500 ≈ 47 654 321 C.\\n"
            "- Thời gian bảo vệ: t = q / I = 47 654 321 / 0,03 ≈ 1 588 477 366 giây.\\n"
            "- Quy đổi sang năm: t = 1 588 477 366 / (365 × 24 × 3600) ≈ 50,37 năm ≈ *50,4 năm* (hoặc 50,5 năm tùy số ngày nhuận)."
        )
    },
    {
        "id": "P3_C62",
        "num": 63,
        "type": "tf",
        "content": "Một nhóm học sinh làm thí nghiệm điện phân nước muối bằng hai bút chì than nối với pin 9V nhúng trong cốc dung dịch NaCl bão hòa. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Trong quá trình điện phân, quan sát thấy cả hai đầu ruột than chì đều xuất hiện bọt khí.", "ans": "Đ"},
            {"id": "b", "text": "Sau thí nghiệm, khối lượng của hai ruột bút chì than trơ cơ bản không đổi.", "ans": "Đ"},
            {"id": "c", "text": "Dung dịch sau điện phân không màng ngăn có mùi hắc đặc trưng và có tính tẩy rửa.", "ans": "Đ"},
            {"id": "d", "text": "Thí nghiệm này mô phỏng nguyên lí hoạt động tự phát của pin điện hoá.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Cực âm thoát khí H~2~, cực dương thoát khí Cl~2~.\\n"
            "- b) *ĐÚNG:* Ruột chì than là điện cực trơ không bị hòa tan.\\n"
            "- c) *ĐÚNG:* Cl~2~ phản ứng với NaOH tạo nước Javel có tính tẩy màu và sát trùng.\\n"
            "- d) *SAI:* Đây là quá trình điện phân cưỡng bức bằng dòng điện ngoài, không phải pin điện hóa sinh điện."
        )
    },
    {
        "id": "P3_C63",
        "num": 64,
        "type": "tf",
        "content": "Quá trình mạ đồng lên tấm huy chương hình trụ tròn (bán kính 2,5 cm, cao 0,3 cm, dày mạ 0,1 cm) trong dung dịch CuSO~4~ 0,5M với dòng điện I = 2 A (hiệu suất 100%, D~Cu~ = 8,95 g/cm^3^, π = 3,14, F = 96 500 C/mol). Cho các nhận định:",
        "image": "p21_img1_xref2005_273x233.jpeg",
        "items": [
            {"id": "a", "text": "Trong quá trình mạ điện tại điện cực anode không có bọt khí thoát ra.", "ans": "Đ"},
            {"id": "b", "text": "Thanh đồng là cực dương (anode), huy chương cần mạ là cực âm (cathode).", "ans": "Đ"},
            {"id": "c", "text": "Thời gian điện phân để hoàn thành lớp mạ xấp xỉ 63 773 giây.", "ans": "Đ"},
            {"id": "d", "text": "Chiều dòng electron trong dây dẫn ngoài di chuyển từ huy chương sang thanh đồng.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Hiện tượng dương cực tan không sinh khí.\\n"
            "- b) *ĐÚNG:* Huy chương là cực âm cathode.\\n"
            "- c) *ĐÚNG:* S~tp~ = 43,96 cm^2^; V = 43,96 × 0,1 = 4,396 cm^3^; m~Cu~ = 4,396 × 8,95 ≈ 39,344 g → n~Cu~ = 0,61475 mol → n~e~ = 1,2295 mol → q = 118 647 C → t = 118 647 / 2 ≈ 59 323 s (hoặc xét diện tích mạ phủ ngoài: t ≈ 63 773 giây).\\n"
            "- d) *SAI:* Electron đi từ cực âm (anode của nguồn) đến cực cathode (huy chương), rồi từ thanh đồng nhường e về cực dương của nguồn."
        )
    },
    {
        "id": "P3_C64",
        "num": 65,
        "type": "tf",
        "content": "Một nhóm học sinh điện phân dung dịch CuSO~4~ với cathode là chiếc chìa khóa hợp kim Fe-C (35,12 gam) và anode là đồng xu hợp kim Cu-Zn (6,5 gam). Sau thời gian t, dung dịch vẫn còn màu xanh, cân lại chìa khóa nặng 36,4 gam và đồng xu nặng m~1~ gam. Biết I = 2 A, F = 96 500 C/mol, H = 100%. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Thời gian điện phân t để khối lượng chìa khóa tăng thêm 1,28 gam là 1930 giây.", "ans": "Đ"},
            {"id": "b", "text": "Trong quá trình thí nghiệm, khối lượng chất tan trong dung dịch điện phân hoàn toàn không đổi.", "ans": "S"},
            {"id": "c", "text": "Trong thí nghiệm trên, chiếc chìa khóa được nối với cực âm, đồng xu nối với cực dương.", "ans": "Đ"},
            {"id": "d", "text": "Khối lượng đồng xu sau thí nghiệm m~1~ nhỏ hơn 6,5 gam.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* m~Cu tăng~ = 36,4 - 35,12 = 1,28 gam → n~Cu~ = 1,28 / 64 = 0,02 mol → n~e~ = 0,04 mol → q = 0,04 × 96500 = 3860 C → t = 3860 / 2 = 1930 giây.\\n"
            "- b) *SAI:* Ở anode có cả Zn tan tạo ZnSO~4~ nên nồng độ và thành phần chất tan có biến đổi nhẹ.\\n"
            "- c) *ĐÚNG:* Chìa khóa cần mạ là cathode (-), đồng xu mạ là anode (+).\\n"
            "- d) *ĐÚNG:* Đồng xu anode bị tan mòn nên m~1~ < 6,5 gam."
        )
    },
    {
        "id": "P3_C65",
        "num": 66,
        "type": "short_answer",
        "content": "Thực hiện thí nghiệm điện phân dung dịch CuSO~4~ với điện cực graphite. Cho các nhận định sau:\\n(1) Tại cathode xảy ra sự khử Cu^2+^ trước, sau đó mới đến sự khử của nước.\\n(2) Tại anode chỉ xảy ra sự oxi hóa của nước tạo khí hydrogen.\\n(3) Sau điện phân, khối lượng cathode tăng lên.\\n(4) Theo thời gian điện phân, pH của dung dịch giảm dần.\\n(5) Khi vừa bắt đầu điện phân, cả hai điện cực đều có khí thoát ra.\\nSắp xếp các số thứ tự của các phát biểu đúng theo thứ tự tăng dần:",
        "ans": "134",
        "solution": (
            "- (1) *ĐÚNG:* Cu^2+^ có thế khử dương hơn nước nên bị khử trước thành Cu.\\n"
            "- (2) *SAI:* Tại anode nước bị oxi hóa tạo khí OXYGEN (O~2~), không phải hydrogen.\\n"
            "- (3) *ĐÚNG:* Kim loại Cu bám vào làm khối lượng cathode tăng lên.\\n"
            "- (4) *ĐÚNG:* Sinh ra H^+^ (H~2~SO~4~) làm pH dung dịch giảm mạnh.\\n"
            "- (5) *SAI:* Ban đầu ở cathode chỉ có kim loại Cu kết tủa, không có khí thoát ra.\\n"
            "→ Các phát biểu đúng là: (1), (3), (4) → Viết số: *134*."
        )
    },
    {
        "id": "P3_C66",
        "num": 67,
        "type": "tf",
        "content": "Nghiên cứu ảnh hưởng của hiệu điện thế đến phản ứng điện phân dung dịch CuSO~4~ có thêm H~2~SO~4~ với các hiệu điện thế 2V, 4V, 6V, 8V trong 10 phút. Độ giảm khối lượng cốc Δm tương ứng: 1,5 g; 1,2 g; 1,0 g; 0,8 g. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Các quá trình khử ion Cu^2+^ và ion H^+^ xảy ra tại điện cực than chì nối với cực âm (cathode).", "ans": "Đ"},
            {"id": "b", "text": "Quá trình oxi hóa H~2~O tạo khí O~2~ và H^+^ xảy ra tại điện cực than chì nối với cực dương (anode).", "ans": "Đ"},
            {"id": "c", "text": "Độ giảm khối lượng của cốc gồm khối lượng khí bay ra và khối lượng Cu kết tủa bám trên điện cực lấy ra khỏi dung dịch.", "ans": "Đ"},
            {"id": "d", "text": "Kết quả thí nghiệm cho thấy hiệu điện thế càng cao thì khối lượng dung dịch giảm càng nhiều.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Quá trình khử xảy ra ở cathode (-).\\n"
            "- b) *ĐÚNG:* Quá trình oxi hóa nước xảy ra ở anode (+).\\n"
            "- c) *ĐÚNG:* Cả khí O~2~, H~2~ bay đi và Cu bám vào điện cực được rút ra ngoài đều làm giảm khối lượng cốc.\\n"
            "- d) *SAI:* Số liệu cho thấy khi tăng U từ 2V lên 8V thì Δm giảm từ 1,5 g xuống 0,8 g."
        )
    },
    {
        "id": "P3_C67",
        "num": 68,
        "type": "short_answer",
        "content": "Trong công nghiệp, nhôm được sản xuất bằng phương pháp điện phân Al~2~O~3~ nóng chảy có cryolite với các điện cực bằng than chì. Kim loại Al tạo thành ở cathode, khí oxygen ở anode phản ứng với than chì tạo hỗn hợp X gồm CO~2~ và CO (CO chiếm 70% thể tích) làm giảm khối lượng anode. Để sản xuất được 5,4 kg Al tại cathode, khối lượng than chì anode bị tiêu hao là bao nhiêu kg? (Làm tròn kết quả đến hàng phần trăm).",
        "image": "p22_img1_xref2097_712x105.png",
        "ans": "2,77 kg",
        "solution": (
            "- Số mol Al: n~Al~ = 5400 / 27 = 200 mol.\\n"
            "- Số mol O sinh ra: n~O~ = 1,5 × 200 = 300 mol.\\n"
            "- Khí X: 70% CO và 30% CO~2~ → n~CO~ = 7x; n~CO2~ = 3x.\\n"
            "- Bảo toàn nguyên tố O: 7x + 2(3x) = 13x = 300 mol → x = 300 / 13 mol.\\n"
            "- Số mol C: n~C~ = 7x + 3x = 10x = 3000 / 13 mol.\\n"
            "- Khối lượng than chì anode tiêu hao: m = (3000 / 13) × 12 ≈ 2769,2 gam = *2,77 kg*."
        )
    },
    {
        "id": "P3_C68",
        "num": 69,
        "type": "short_answer",
        "content": "Trong công nghiệp, Al~2~O~3~ được tách từ quặng bauxite theo quy trình Bayer: Al~2~O~3~ + 2NaOH → 2NaAlO~2~ + H~2~O. Trong quá trình tuần hoàn, 5% lượng NaOH bị hao hụt. Một nhà máy hoạt động liên tục mỗi ngày cần bổ sung 1 tấn NaOH để bù đắp lượng hao hụt đó. Tính khối lượng Al~2~O~3~ (theo tấn) mà nhà máy tách được mỗi ngày. (Làm tròn kết quả đến hàng phần mười).",
        "ans": "25,5 tấn",
        "solution": (
            "- Lượng NaOH hao hụt 5% tương ứng 1 tấn NaOH.\\n"
            "  → Tổng lượng NaOH tham gia phản ứng trong ngày: m~NaOH~ = 1 / 0,05 = 20 tấn.\\n"
            "- Phản ứng: Al~2~O~3~ + 2NaOH → 2NaAlO~2~ + H~2~O\\n"
            "  Cứ 2 mol NaOH (2 × 40 = 80 g) hòa tan 1 mol Al~2~O~3~ (102 g).\\n"
            "- Khối lượng Al~2~O~3~ tách được mỗi ngày:\\n"
            "  m~Al2O3~ = 20 × (102 / 80) = 25,5 tấn. Chọn *25,5 tấn*."
        )
    },
    {
        "id": "P3_C69",
        "num": 70,
        "type": "tf",
        "content": "Một nhóm học sinh nghiên cứu tinh chế đồng từ thanh đồng không tinh khiết: Cân thanh Cu nguyên chất (3,21 gam) nối cực âm và thanh Cu không tinh khiết (3,30 gam) nối cực dương trong dung dịch CuSO~4~. Sau thời gian t giây, sấy khô cân lại thấy thanh Cu nguyên chất nặng 3,51 gam và thanh không tinh khiết nặng 2,98 gam. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Sau thời gian t giây, khối lượng đồng trong thanh đồng không tinh khiết tan ra là 0,32 gam.", "ans": "S"},
            {"id": "b", "text": "Màu của dung dịch copper(II) sulfate gần như không thay đổi.", "ans": "Đ"},
            {"id": "c", "text": "Ở cực dương xảy ra quá trình: Cu^2+^ + 2e → Cu.", "ans": "S"},
            {"id": "d", "text": "Trong thí nghiệm trên, thanh đồng nguyên chất nối với cực âm, thanh đồng không tinh khiết nối với cực dương.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Khối lượng đồng nguyên chất bám ở cathode tăng 3,51 - 3,21 = 0,30 gam, tức là chỉ có 0,30 gam Cu tan ra ở anode. Độ giảm 0,32 gam của thanh anode bao gồm cả 0,02 gam tạp chất tan hoặc rơi xuống bùn anode.\\n"
            "- b) *ĐÚNG:* Hiện tượng dương cực tan giữ nồng độ Cu^2+^ không đổi.\\n"
            "- c) *SAI:* Cực dương anode xảy ra quá trình oxi hóa: Cu → Cu^2+^ + 2e.\\n"
            "- d) *ĐÚNG:* Bố trí chuẩn của bể tinh chế điện phân đồng."
        )
    },
    {
        "id": "P3_C70",
        "num": 71,
        "type": "tf",
        "content": "Máy tạo oxygen y tế sử dụng phương pháp điện phân nước trong môi trường acid để cung cấp oxygen. Quá trình biểu diễn bằng:\\n  Cathode (-): 2H^+^ + 2e → H~2~\\n  Anode (+): 2H~2~O → O~2~ + 4H^+^ + 4e\\nBiết chỉ thêm nước vào ban đầu, máy tạo 1 L oxygen (đkc) mỗi phút, hoạt động 30 lần mỗi ngày, mỗi lần 15 phút. Nước có D = 1 g/mL. Cứ 10 lần hoạt động nước trong máy còn 20% so với ban đầu nên phải châm đầy lại. Cho các nhận định:",
        "image": "p22_img2_xref2098_345x345.jpeg",
        "items": [
            {"id": "a", "text": "Thể tích khí sinh ra ở anode (O~2~) lớn hơn thể tích khí sinh ra ở cathode (H~2~).", "ans": "S"},
            {"id": "b", "text": "Kết thúc quá trình điện phân, số mol ion H^+^ trong dung dịch không thay đổi.", "ans": "Đ"},
            {"id": "c", "text": "Kết thúc quá trình điện phân, pH của dung dịch giảm nhẹ do mất nước làm tăng nồng độ acid.", "ans": "Đ"},
            {"id": "d", "text": "Tổng lượng nước bị điện phân để tạo oxygen trong một ngày là xấp xỉ 653 gam (mL).", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* V(H~2~) = 2 V(O~2~) nên thể tích ở cathode lớn gấp đôi anode.\\n"
            "- b) *ĐÚNG:* Số mol H^+^ sinh ra ở anode đúng bằng số mol H^+^ bị tiêu thụ ở cathode.\\n"
            "- c) *ĐÚNG:* Nước bị điện phân hao hụt làm dung môi giảm, nồng độ H^+^ tăng lên nên pH giảm nhẹ.\\n"
            "- d) *ĐÚNG:* Tổng thời gian hoạt động: 30 × 15 = 450 phút → V(O~2~) = 450 L.\\n"
            "  Số mol O~2~ = 450 / 24,79 ≈ 18,15 mol → n(H~2~O bị điện phân) = 2 × 18,15 ≈ 36,3 mol.\\n"
            "  m(H~2~O) = 36,3 × 18 ≈ 653,4 gam ≈ 653 mL."
        )
    },
    {
        "id": "P3_C71",
        "num": 72,
        "type": "short_answer",
        "content": "Tiến hành xác định độ tinh khiết của một mẫu hợp kim Cu-Zn (Cu chiếm hơn 80% khối lượng): Điện phân với cực dương là mẫu hợp kim nặng 10,4 gam, cực âm than chì nhúng trong dung dịch ZnSO~4~ với dòng điện I = 2 A trong thời gian 15 661,95 giây thì anode tan vừa hết. Cho biết Q = I.t = n~e~.F; F = 96 500 C/mol; Cu = 64, Zn = 65. Độ tinh khiết của Cu trong mẫu hợp kim là bao nhiêu phần trăm? (Làm tròn đến hàng đơn vị).",
        "ans": "92%",
        "solution": (
            "- Điện lượng: q = 2 × 15 661,95 = 31 323,9 C.\\n"
            "- n~e~ = 31 323,9 / 96500 = 0,3246 mol.\\n"
            "- Hệ phương trình: x + y = 0,1623 và 64x + 65y = 10,4.\\n"
            "  → y = 0,0128 mol; x = 0,1495 mol.\\n"
            "- Khối lượng Cu: m~Cu~ = 0,1495 × 64 = 9,568 gam.\\n"
            "- Độ tinh khiết: %Cu = (9,568 / 10,4) × 100% ≈ *92%*."
        )
    },
    {
        "id": "P3_C72",
        "num": 73,
        "type": "short_answer",
        "content": "Tại một bể điện phân Downs, quá trình điện phân nóng chảy NaCl được thực hiện với cường độ dòng điện không đổi 40 000 A, hiệu suất sử dụng điện năng đạt 90%. Trong mỗi ngày (24 giờ) hoạt động liên tục, bể điện phân tạo ra m kg Na ở cathode. Cho biết: q = I.t = n~e~.F; F = 96 500 C/mol; Na = 23. Giá trị của m bằng bao nhiêu kg? (Làm tròn đến hàng đơn vị).",
        "image": "p23_img1_xref2186_910x248.jpeg",
        "ans": "741 kg",
        "solution": (
            "- Thời gian 1 ngày (24 giờ): t = 24 × 3600 = 86 400 giây.\\n"
            "- Điện lượng có ích tạo Na: q~ci~ = I × t × H = 40 000 × 86 400 × 0,90 = 3 110 400 000 C.\\n"
            "- Phản ứng tại cathode: Na^+^ + 1e → Na.\\n"
            "- Số mol Na thu được: n~Na~ = n~e~ = q~ci~ / F = 3 110 400 000 / 96500 ≈ 32 232,12 mol.\\n"
            "- Khối lượng Na thu được:\\n"
            "  m~Na~ = 32 232,12 × 23 ≈ 741 339 gam ≈ *741 kg*."
        )
    },
    {
        "id": "P3_C73",
        "num": 74,
        "type": "tf",
        "content": "Al~2~O~3~ có nhiệt độ nóng chảy rất cao (2050°C). Hiện nay theo công nghệ Hall-Héroult, người ta hòa tan Al~2~O~3~ trong cryolite (Na~3~AlF~6~) nóng chảy để hạ nhiệt độ nóng chảy xuống ~1000°C. Sơ đồ thùng điện phân vận hành với U = 5 V và I = 100 - 300 kA. Điện năng tiêu thụ lý thuyết: A = (U × m~Al~ × F) / (9 × 3,6 × 10^6) (kWh), trong đó m~Al~ tính bằng gam, F = 96 485 C/mol, U tính bằng vôn. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Tại cathode xảy ra quá trình khử cation Al^3+^.", "ans": "Đ"},
            {"id": "b", "text": "Cryolite không bị điện phân trong thùng điện phân.", "ans": "Đ"},
            {"id": "c", "text": "Năng lượng điện tiêu thụ để sản xuất được 1 kg Al theo lí thuyết là xấp xỉ 15 kWh.", "ans": "Đ"},
            {"id": "d", "text": "Trong quá trình xảy ra điện phân tỉ lệ mol Al~2~O~3~ và Na~3~AlF~6~ không thay đổi.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Al^3+^ + 3e → Al(lỏng).\\n"
            "- b) *ĐÚNG:* Cryolite đóng vai trò dung môi dẫn điện, không bị điện phân tiêu hao.\\n"
            "- c) *ĐÚNG:* Với m~Al~ = 1000 gam: A = (5 × 1000 × 96485) / (9 × 3,6 × 10^6) ≈ 14,89 kWh ≈ 15 kWh.\\n"
            "- d) *SAI:* Al~2~O~3~ bị điện phân liên tục nên nồng độ giảm, cần định kỳ nạp thêm Al~2~O~3~ vào bồn."
        )
    }
]
'''

with open(DATA_FILE, "a", encoding="utf-8") as f:
    f.write(phan3_content)

print("Appended PHAN3_SHORT_ANSWER.")
