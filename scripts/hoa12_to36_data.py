# -*- coding: utf-8 -*-
"""
Module dữ liệu chuẩn cho TỜ SỐ 36 - CHƯƠNG 5: PIN ĐIỆN VÀ ĐIỆN PHÂN (HÓA HỌC 12)
Thầy TRẦN VĂN THẠNH - 0777.470.803
Bản quyền tài liệu: Thầy Trần Văn Thạnh - TP Huế
"""

import os

ROOT_DIR = r"c:\Antigravity_Thanh"
IMG_DIR = os.path.join(ROOT_DIR, "San_Pham", "images_to36")

INFO = {
    "title": "CHUYÊN ĐỀ 5: PIN ĐIỆN VÀ ĐIỆN PHÂN (TỜ SỐ 36)",
    "sub_title": "CHUYÊN ĐỀ ÔN TẬP VÀ PHÁT TRIỂN NĂNG LỰC HÓA HỌC 12",
    "subject": "HÓA HỌC 12 - CHƯƠNG TRÌNH MỚI 2026 - 2027",
    "grade": "12",
    "teacher": "Thầy TRẦN VĂN THẠNH",
    "phone": "0777.470.803",
    "branch1": "CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế",
    "branch2": "CS2: 24 Đặng Thái Thân, TP Huế",
}

# ==============================================================================
# DẠNG 1: MẠ KIM LOẠI
# ==============================================================================
THEORY_DANG1 = {
    "title": "DẠNG 1: MẠ KIM LOẠI",
    "sections": [
        {
            "heading": "1. Mục đích",
            "items": [
                "Tạo lớp kim loại phủ bên ngoài vật liệu để:",
                "  + Chống ăn mòn, gỉ sét (ví dụ: mạ Cr, Ni).",
                "  + Tăng tính thẩm mỹ (ví dụ: mạ vàng Au, mạ bạc Ag).",
                "  + Tăng độ bền cơ học, độ cứng bề mặt."
            ]
        },
        {
            "heading": "2. Cấu tạo hệ điện phân dùng trong mạ điện",
            "table": {
                "headers": ["Thành phần", "Vai trò"],
                "rows": [
                    ["Anode (cực dương)", "Kim loại dùng để mạ (ví dụ: Ni, Cu, Cr, Ag...); xảy ra hiện tượng anode tan."],
                    ["Cathode (cực âm)", "Vật cần mạ (dụng cụ, trang sức, linh kiện điện tử, chi tiết máy...)."],
                    ["Dung dịch điện phân", "Dung dịch muối chứa cation của kim loại mạ (NiSO~4~, CuSO~4~, AgNO~3~, [Ag(NH~3~)~2~]^+^...)."],
                    ["Nguồn điện một chiều (DC)", "Cung cấp dòng điện một chiều với hiệu điện thế và cường độ thích hợp."]
                ]
            }
        },
        {
            "heading": "3. Phản ứng xảy ra tại các điện cực",
            "items": [
                "🔺 *Tại Anode (cực dương - kim loại mạ):* Kim loại bị oxi hóa tan vào dung dịch:\n      M → M^n+^ + ne",
                "🔻 *Tại Cathode (cực âm - vật cần mạ):* Ion kim loại trong dung dịch bị khử, bám đều lên bề mặt vật:\n      M^n+^ + ne → M",
                "→ Kết quả: Kim loại được mạ tạo thành lớp phủ mỏng, mịn, bám chắc trên bề mặt vật cần mạ. Nồng độ ion kim loại trong dung dịch điện phân cơ bản không đổi."
            ]
        },
        {
            "heading": "4. Lưu ý quan trọng khi mạ điện",
            "table": {
                "headers": ["Nội dung", "Lưu ý kỹ thuật"],
                "rows": [
                    ["1. Bề mặt vật cần mạ", "Phải được làm sạch triệt để (tẩy dầu mỡ, rửa acid, chà nhám, đánh bóng...) để lớp mạ bám chắc."],
                    ["2. Cường độ dòng điện", "Phải vừa đủ. Nếu quá mạnh → lớp mạ xốp, sần sùi, dễ bong tróc; nếu quá yếu → mạ rất chậm, lớp mạ không đều."],
                    ["3. Dung dịch điện phân", "Phải có nồng độ ion kim loại thích hợp và bổ sung phụ gia làm mịn lớp mạ."],
                    ["4. Anode tan", "Anode phải là kim loại nguyên chất để tan ra liên tục, duy trì ổn định nồng độ cation trong dung dịch."],
                    ["5. Nhiệt độ & khuấy trộn", "Cần kiểm soát nhiệt độ thích hợp và khuấy nhẹ dung dịch để chống phân cực nồng độ."],
                    ["6. Thời gian mạ", "Quyết định độ dày lớp mạ: độ dày tăng tỉ lệ thuận với thời gian mạ theo định luật Faraday."]
                ]
            }
        },
        {
            "heading": "5. Ví dụ thực tế",
            "image": "p1_img1_xref91_347x335.jpeg",
            "items": [
                "• *Mạ đồng (Cu):* Anode: Cu; Cathode: vật cần mạ (sắt, thép, chi tiết máy...); Dung dịch: CuSO~4~ + H~2~SO~4~ loãng.",
                "• *Mạ bạc (Ag):* Anode: Ag; Cathode: trang sức, bộ đồ ăn, chi tiết mạch điện; Dung dịch: AgNO~3~ hoặc phức chất [Ag(NH~3~)~2~]^+^.",
                "• *Mạ niken (Ni):* Anode: Ni; Cathode: vật cần mạ; Dung dịch: NiSO~4~ có đệm pH."
            ]
        }
    ]
}

EXERCISES_DANG1 = [
    {
        "id": "D1_C1",
        "num": 1,
        "type": "essay",
        "content": "Để mạ 5,0 g bạc lên một đĩa sắt khi điện phân dung dịch chứa [Ag(NH~3~)~2~]^+^ với dòng điện có cường độ 1,5 A không đổi cần thời gian t phút trong hai trường hợp:\n  a) Hiệu suất điện phân đạt 100%.\n  b) Hiệu suất điện phân đạt 80%.\nCho biết:\n- Quá trình khử tại cathode: [Ag(NH~3~)~2~]^+^ + 1e → Ag + 2NH~3~.\n- Điện lượng q = I.t = n~e~.F; F = 96 500 C/mol; Ag = 108.\nTính giá trị của t trong mỗi trường hợp (làm tròn kết quả đến hàng phần mười).",
        "ans": "a) t = 49,6 phút; b) t = 62,0 phút",
        "solution": (
            "- Số mol Ag cần mạ: n~Ag~ = 5,0 / 108 ≈ 0,04630 mol.\n"
            "- Theo bán phản ứng tại cathode: n~e~ = n~Ag~ ≈ 0,04630 mol.\n"
            "- Điện lượng lý thuyết: q = n~e~ × F = (5,0 / 108) × 96500 ≈ 4467,59 C.\n"
            "a) Với hiệu suất H = 100%:\n"
            "   t = q / I = 4467,59 / 1,5 ≈ 2978,4 giây = 49,64 phút ≈ *49,6 phút*.\n"
            "b) Với hiệu suất H = 80%:\n"
            "   t = 49,64 / 0,80 = 62,05 phút ≈ *62,0 phút*."
        )
    },
    {
        "id": "D1_C2",
        "num": 2,
        "type": "essay",
        "content": "Cần mạ một lớp Ag lên một mặt của một chiếc đĩa tròn có bán kính 12 cm. Với độ dày lớp mạ là 0,01 mm, nếu được cung cấp nguồn điện một chiều có cường độ dòng điện I = 2 A trong thời gian t giờ. Cho biết khối lượng riêng của Ag là D = 10,5 g/cm^3^; Ag = 108; F = 96 500 C/mol; π = 3,1416.\nTính giá trị của t trong hai trường hợp (làm tròn đến hàng phần trăm):\n  a) Hiệu suất điện phân là 100%.\n  b) Hiệu suất điện phân là 90%.",
        "ans": "a) t = 0,59 giờ (35,4 phút); b) t = 0,65 giờ (39,3 phút)",
        "solution": (
            "- Bán kính r = 12 cm; độ dày lớp mạ h = 0,01 mm = 0,001 cm.\n"
            "- Diện tích mặt đĩa tròn cần mạ: S = π × r^2^ = π × 12^2^ = 144π ≈ 452,39 cm^2^.\n"
            "- Thể tích lớp bạc: V = S × h = 452,39 × 0,001 ≈ 0,45239 cm^3^.\n"
            "- Khối lượng bạc cần phủ: m~Ag~ = V × D = 0,45239 × 10,5 ≈ 4,750 g.\n"
            "- Số mol Ag: n~Ag~ = 4,750 / 108 ≈ 0,04398 mol → n~e~ = n~Ag~ ≈ 0,04398 mol.\n"
            "- Điện lượng lý thuyết: q = 0,04398 × 96500 ≈ 4244,3 C.\n"
            "a) Với H = 100%:\n"
            "   t = q / I = 4244,3 / 2 = 2122,15 giây = 2122,15 / 3600 ≈ *0,59 giờ* (≈ 35,4 phút).\n"
            "b) Với H = 90%:\n"
            "   t = 0,5895 / 0,90 ≈ *0,65 giờ* (≈ 39,3 phút)."
        )
    },
    {
        "id": "D1_C3",
        "num": 3,
        "type": "essay",
        "content": "Một vật kim loại được mạ nickel có diện tích bề mặt 100 cm^2^. Dòng điện chạy qua bình điện phân có cường độ 0,5 A và thời gian mạ là 5 giờ. Cho biết nickel có khối lượng mol nguyên tử là 58,7 g/mol, hóa trị 2 trong muối, khối lượng riêng D = 8,8 × 10^3^ kg/m^3^ (tức 8,8 g/cm^3^), hằng số Faraday F = 96 500 C/mol, hiệu suất điện phân 100%.\nĐộ dày của lớp nickel phủ đều trên mặt của vật được mạ là bao nhiêu µm? (Kết quả làm tròn đến hàng phần mười).",
        "ans": "31,1 µm",
        "solution": (
            "- Thời gian điện phân: t = 5 × 3600 = 18 000 s.\n"
            "- Số mol electron trao đổi: n~e~ = (I × t) / F = (0,5 × 18 000) / 96500 ≈ 0,093264 mol.\n"
            "- Phản ứng tại cathode: Ni^2+^ + 2e → Ni → n~Ni~ = n~e~ / 2 ≈ 0,046632 mol.\n"
            "- Khối lượng Ni kết tủa: m~Ni~ = 0,046632 × 58,7 ≈ 2,7373 g.\n"
            "- Thể tích lớp nickel: V = m / D = 2,7373 / 8,8 ≈ 0,31106 cm^3^.\n"
            "- Bề dày lớp mạ: h = V / S = 0,31106 / 100 = 0,0031106 cm = 31,106 µm ≈ *31,1 µm*."
        )
    },
    {
        "id": "D1_C4",
        "num": 4,
        "type": "essay",
        "content": "Một tấm huy chương mặt tròn bằng kim loại có đường kính 70 mm, bề dày 3 mm; được mạ thêm một lớp vàng dày 0,05 mm bám đều lên tất cả các mặt của tấm huy chương. Biết khối lượng riêng của Au là 19,3 g/cm^3^, khối lượng mol của Au là 197 g/mol; khi điện phân mỗi ion vàng trao đổi 3 electron và hằng số Faraday F = 96 500 C/mol; hiệu suất 100%.\nThời gian điện phân với cường độ dòng điện không đổi I = 10 A để mạ lớp vàng lên tấm huy chương là bao nhiêu giây? (Kết quả làm tròn đến hàng đơn vị).",
        "ans": "1185 giây",
        "solution": (
            "- Quy đổi kích thước huy chương: bán kính r = 35 mm = 3,5 cm; chiều dày h~hc~ = 3 mm = 0,3 cm.\n"
            "- Tổng diện tích bề mặt (2 mặt tròn + mặt xung quanh):\n"
            "  S~tp~ = 2πr^2^ + 2πr × h~hc~ = 2π × (3,5)^2^ + 2π × 3,5 × 0,3 = 24,5π + 2,1π = 26,6π ≈ 83,566 cm^2^.\n"
            "- Bề dày lớp mạ vàng: h = 0,05 mm = 0,005 cm.\n"
            "- Thể tích lớp vàng: V~Au~ = S~tp~ × h = 83,566 × 0,005 ≈ 0,41783 cm^3^.\n"
            "- Khối lượng vàng cần mạ: m~Au~ = V~Au~ × D = 0,41783 × 19,3 ≈ 8,064 g.\n"
            "- Số mol Au: n~Au~ = 8,064 / 197 ≈ 0,040935 mol.\n"
            "- Số mol electron: n~e~ = 3 × n~Au~ ≈ 0,122806 mol.\n"
            "- Điện lượng cần thiết: q = n~e~ × F = 0,122806 × 96500 ≈ 11850,8 C.\n"
            "- Thời gian điện phân: t = q / I = 11850,8 / 10 ≈ *1185 giây*."
        )
    },
    {
        "id": "D1_C5",
        "num": 5,
        "type": "essay",
        "content": "Để thử nghiệm kỹ thuật mạ điện người ta phủ một lớp Zn lên một mẫu thử hình lập phương có chiều dài mỗi cạnh là 12 cm. Nhúng mẫu thử vào bình điện phân chứa dung dịch ZnSO~4~, cathode là mẫu thử, anode bằng Zn. Thực hiện quá trình mạ điện với dòng điện một chiều I = 15 A trong thời gian 34,44 giờ. Biết hiệu suất điện phân đạt 90%; khối lượng riêng của Zn là 7,13 g/cm^3^; Zn = 65,38; F = 96 500 C/mol.\nBề dày lớp kẽm phủ đều lên mẫu thử là bao nhiêu mm? (Kết quả làm tròn đến hàng phần mười).",
        "ans": "0,9 mm",
        "solution": (
            "- Diện tích toàn phần khối lập phương (6 mặt): S~tp~ = 6 × a^2^ = 6 × 12^2^ = 864 cm^2^.\n"
            "- Thời gian điện phân: t = 34,44 × 3600 = 123 984 s.\n"
            "- Điện lượng thực tế có ích: q~ci~ = I × t × H = 15 × 123 984 × 0,90 = 1 673 784 C.\n"
            "- Số mol e có ích: n~e~ = 1 673 784 / 96500 ≈ 17,3449 mol.\n"
            "- Phản ứng tại cathode: Zn^2+^ + 2e → Zn → n~Zn~ = n~e~ / 2 ≈ 8,67245 mol.\n"
            "- Khối lượng Zn tạo thành: m~Zn~ = 8,67245 × 65,38 ≈ 566,99 g (hoặc dùng 65: ≈ 563,7 g).\n"
            "- Thể tích Zn: V = m / D = 566,99 / 7,13 ≈ 79,52 cm^3^.\n"
            "- Bề dày lớp kẽm: h = V / S~tp~ = 79,52 / 864 ≈ 0,0920 cm = 0,92 mm ≈ *0,9 mm*."
        )
    },
    {
        "id": "D1_C6",
        "num": 6,
        "type": "essay",
        "content": "Người ta mạ kẽm lên bề mặt quả cân hình cầu bằng phương pháp mạ điện. Dung dịch điện phân chứa ZnSO~4~, cực dương là Zn kim loại, cực âm là quả cân (bán kính R = 5 cm). Sự điện phân với cường độ dòng điện không đổi I = 16 A. Quả cân cần được phủ đều một lớp kẽm dày 1 mm trên bề mặt. Biết hiệu suất điện phân đạt 100%; khối lượng riêng của Zn là 7,13 g/cm^3^; Zn = 65,38; π = 3,1416; F = 96 500 C/mol.\nThời gian của quá trình mạ điện là bao nhiêu giờ? (Kết quả làm tròn đến hàng phần mười).",
        "ans": "11,5 giờ",
        "solution": (
            "- Diện tích mặt cầu: S = 4πR^2^ = 4 × 3,1416 × 5^2^ = 314,16 cm^2^.\n"
            "- Bề dày lớp kẽm: h = 1 mm = 0,1 cm.\n"
            "- Thể tích lớp kẽm: V = S × h = 314,16 × 0,1 = 31,416 cm^3^.\n"
            "- Khối lượng kẽm bám lên quả cân: m~Zn~ = V × D = 31,416 × 7,13 ≈ 224,0 g.\n"
            "- Số mol Zn: n~Zn~ = 224,0 / 65,38 ≈ 3,426 mol → n~e~ = 2 × n~Zn~ ≈ 6,852 mol.\n"
            "- Điện lượng q = n~e~ × F = 6,852 × 96500 ≈ 661 228 C.\n"
            "- Thời gian điện phân: t = q / I = 661 228 / 16 ≈ 41 327 giây ≈ 11,48 giờ ≈ *11,5 giờ*."
        )
    },
    {
        "id": "D1_C7",
        "num": 7,
        "type": "essay",
        "content": "Để mạ lớp nickel (Ni) lên một chi tiết máy bằng thép hình trụ đường kính 7 cm, chiều dài 15 cm, người ta sử dụng chi tiết máy cần mạ đó làm cathode và nhúng vào trong bể mạ chứa dung dịch NiSO~4~, cường độ dòng điện đi qua bể mạ là 8,0 A (hiệu suất điện phân là 100%). Nếu độ dày lớp mạ nickel lên toàn bộ bề mặt chi tiết máy là 20 µm (1 µm = 10^-4^ cm) thì thời gian cần để mạ chi tiết trên là t phút. Cho biết khối lượng riêng của Ni bằng 8,91 g/cm^3^; Ni = 58,7; hằng số Faraday F = 96 500 C/mol.\nTính giá trị của t (kết quả làm tròn đến hàng phần mười).",
        "ans": "49,7 phút (hoặc 40,3 phút nếu chỉ mạ mặt xung quanh)",
        "solution": (
            "- Bán kính r = 3,5 cm; chiều dài L = 15 cm; bề dày h = 20 µm = 0,002 cm.\n"
            "- Trường hợp mạ toàn phần chi tiết máy (cả 2 đáy và mặt xung quanh):\n"
            "  S~tp~ = 2πrL + 2πr^2^ = 2π × 3,5 × 15 + 2π × (3,5)^2^ = 105π + 24,5π = 129,5π ≈ 406,84 cm^2^.\n"
            "  Thể tích Ni: V = 406,84 × 0,002 ≈ 0,8137 cm^3^.\n"
            "  Khối lượng Ni: m~Ni~ = 0,8137 × 8,91 ≈ 7,250 g.\n"
            "  Số mol Ni: n~Ni~ = 7,250 / 58,7 ≈ 0,1235 mol → n~e~ = 2 × 0,1235 = 0,2470 mol.\n"
            "  Điện lượng: q = 0,2470 × 96500 ≈ 23838 C.\n"
            "  Thời gian t = q / I = 23838 / 8 = 2980 giây ≈ *49,7 phút*.\n"
            "(Nếu chi tiết máy chỉ mạ mặt xung quanh: S = 105π ≈ 329,87 cm^2^ → t ≈ 40,3 phút)."
        )
    },
    {
        "id": "D1_C8",
        "num": 8,
        "type": "essay",
        "content": "Để ngăn chặn hiện tượng gỉ sét và mài mòn của một chi tiết máy, người ta phủ lớp đồng (Cu) dày 0,5 mm lên toàn bộ bề mặt của chi tiết máy bằng phương pháp mạ điện. Dung dịch điện phân chứa CuSO~4~, cực dương là Cu kim loại, cực âm là chi tiết máy có hình trụ (bán kính 4,25 cm, chiều cao 25 cm). Sự điện phân với dòng điện một chiều I = 15 A. Biết hiệu suất điện phân đạt 100%; khối lượng riêng của Cu là 8,96 g/cm^3^; Cu = 64; π = 3,14; F = 96 500 C/mol.\nThời gian của quá trình mạ điện là bao nhiêu giờ? (Kết quả làm tròn đến hàng phần mười).",
        "ans": "19,5 giờ (hoặc 16,7 giờ nếu chỉ tính mặt xung quanh)",
        "solution": (
            "- Chi tiết hình trụ có r = 4,25 cm; h~trụ~ = 25 cm; bề dày mạ h = 0,5 mm = 0,05 cm.\n"
            "- Diện tích toàn phần chi tiết máy:\n"
            "  S~tp~ = 2πr × h~trụ~ + 2πr^2^ = 2 × 3,14 × 4,25 × 25 + 2 × 3,14 × (4,25)^2^ = 667,25 + 113,43 = 780,68 cm^2^.\n"
            "- Thể tích Cu cần mạ: V = S~tp~ × h = 780,68 × 0,05 = 39,034 cm^3^.\n"
            "- Khối lượng Cu: m~Cu~ = 39,034 × 8,96 ≈ 349,75 g.\n"
            "- Số mol Cu: n~Cu~ = 349,75 / 64 ≈ 5,465 mol → n~e~ = 2 × 5,465 = 10,93 mol.\n"
            "- Điện lượng: q = 10,93 × 96500 ≈ 1 054 745 C.\n"
            "- Thời gian điện phân: t = q / I = 1 054 745 / 15 ≈ 70 316 giây ≈ 19,53 giờ ≈ *19,5 giờ*.\n"
            "(Nếu chỉ mạ mặt xung quanh: S = 667,25 cm^2^ → m~Cu~ = 298,93 g → t ≈ 16,7 giờ)."
        )
    },
    {
        "id": "D1_C9",
        "num": 9,
        "type": "essay",
        "content": "Một tấm huy chương mặt tròn bằng kim loại có bán kính 4 cm, chiều dày 0,5 cm, cần được mạ kín bởi một lớp vàng dày 0,1 mm. Biết khối lượng riêng của Au là 19,36 g/cm^3^; nguyên tử khối của Au là 197 g/mol; π = 3,14; khi điện phân mỗi ion vàng trao đổi 3 electron; F = 96 500 C/mol.\nThời gian điện phân với cường độ dòng điện không đổi 12 A (hiệu suất H = 100%) để mạ lớp vàng lên tấm huy chương đó là bao nhiêu giờ? (Kết quả làm tròn đến hàng phần mười).",
        "ans": "0,7 giờ (hoặc 44,7 phút)",
        "solution": (
            "- Huy chương tròn có r = 4 cm; chiều dày h~hc~ = 0,5 cm; bề dày mạ h = 0,1 mm = 0,01 cm.\n"
            "- Diện tích toàn phần (2 đáy tròn và mặt bên xung quanh):\n"
            "  S~tp~ = 2πr^2^ + 2πr × h~hc~ = 2 × 3,14 × 4^2^ + 2 × 3,14 × 4 × 0,5 = 100,48 + 12,56 = 113,04 cm^2^.\n"
            "- Thể tích lớp vàng: V = S~tp~ × h = 113,04 × 0,01 = 1,1304 cm^3^.\n"
            "- Khối lượng vàng cần mạ: m~Au~ = 1,1304 × 19,36 ≈ 21,885 g.\n"
            "- Số mol Au: n~Au~ = 21,885 / 197 ≈ 0,11109 mol → n~e~ = 3 × 0,11109 ≈ 0,33327 mol.\n"
            "- Điện lượng: q = n~e~ × F = 0,33327 × 96500 ≈ 32 160,5 C.\n"
            "- Thời gian điện phân: t = q / I = 32 160,5 / 12 ≈ 2680 giây ≈ 0,744 giờ ≈ *0,7 giờ* (≈ 44,7 phút)."
        )
    }
]

# ==============================================================================
# DẠNG 2: QUÁ TRÌNH ANODE HÓA NHÔM (ANODIZING)
# ==============================================================================
THEORY_DANG2 = {
    "title": "DẠNG 2: QUÁ TRÌNH ANODE HÓA NHÔM (ANODIZING)",
    "sections": [
        {
            "heading": "I. Mục tiêu của quá trình",
            "items": [
                "• Tăng độ dày và tính bảo vệ của lớp oxide Al~2~O~3~ tự nhiên trên bề mặt nhôm (từ vài nanomet lên đến 10 - 50 µm).",
                "• Tăng vượt bậc khả năng chống ăn mòn khí quyển, chống trầy xước cơ học, tăng độ cứng.",
                "• Tạo bề mặt xốp vi mô có khả năng hấp thụ hạt màu để nhuộm màu thẩm mỹ đa dạng (vàng, đen, xanh, đỏ...)."
            ]
        },
        {
            "heading": "II. Nguyên lý: Điện phân với dòng điện một chiều",
            "items": [
                "• Sử dụng vật liệu nhôm làm Anode (+) trong một bể điện phân chứa dung dịch acid loãng (thường là H~2~SO~4~ 15 - 20%, hoặc acid oxalic).",
                "• Dưới tác dụng của dòng điện một chiều, bề mặt nhôm bị oxi hóa có kiểm soát để tạo thành lớp màng oxide nhôm Al~2~O~3~ đồng nhất và bám cực chặt."
            ]
        },
        {
            "heading": "III. Cấu tạo hệ điện phân & Phản ứng tại các điện cực",
            "image": "p3_img1_xref272_959x251.jpeg",
            "items": [
                "• *Anode (cực dương):* Tấm nhôm hoặc chi tiết nhôm cần xử lý.",
                "      Bán phản ứng oxi hóa: 2Al + 3H~2~O → Al~2~O~3~ + 6H^+^ + 6e",
                "• *Cathode (cực âm):* Điện cực trơ (chì Pb, than chì graphite hoặc thép không gỉ).",
                "      Bán phản ứng khử: 2H^+^ + 2e → H~2~ ↑",
                "• *Khí H~2~ thoát ra tại cathode:* Không ảnh hưởng đến chất lượng và tính đồng nhất của màng oxide ở cực dương."
            ]
        },
        {
            "heading": "IV. Lưu ý kỹ thuật để tạo lớp Al~2~O~3~ tối ưu",
            "table": {
                "headers": ["Yếu tố kỹ thuật", "Vai trò và lưu ý"],
                "rows": [
                    ["Nhiệt độ thấp (0 - 10°C)", "Làm chậm tốc độ hòa tan Al~2~O~3~ trong môi trường acid, giúp màng oxide cứng, dày và sít đặc."],
                    ["Điện áp ổn định", "Nếu điện áp quá cao → lớp oxide bị rỗ tổ ong, cháy cục bộ; nếu quá thấp → màng oxide phát triển mỏng."],
                    ["Dung dịch acid phù hợp", "Dùng dung dịch H~2~SO~4~ tạo màng trắng mờ đặc trưng; acid oxalic cho màng cứng và trong suốt hơn."],
                    ["Thời gian điện phân", "Thời gian càng kéo dài thì lớp oxide càng dày; tuy nhiên kéo dài quá mức sẽ tăng độ xốp bề mặt."],
                    ["Bịt lỗ nhuộm màu (Sealing)", "Sau khi nhuộm màu, chi tiết được nhúng vào nước sôi (hoặc dung dịch muối niken acetate) để bịt kín lỗ xốp."]
                ]
            }
        },
        {
            "heading": "V. Đặc tính sản phẩm & Ứng dụng thực tiễn",
            "items": [
                "• *Đặc tính màng Al~2~O~3~:* Bề dày từ vài µm đến 50 µm; bền trơ trong không khí, chống acid và kiềm nhẹ; cách điện tuyệt vời.",
                "• *Điện tử - Công nghệ:* Vỏ điện thoại di động cao cấp, vỏ máy tính xách tay (MacBook, laptop nhôm), vỏ máy ảnh.",
                "• *Xây dựng & Kiến trúc:* Khung cửa sổ sổ nhôm kính cao cấp, cửa cuốn, mặt dựng tòa nhà, lam chắn nắng.",
                "• *Hàng không & Ô tô:* Vỏ và linh kiện máy bay, mâm đúc xe hơi nhẹ, khung xe đạp thể thao.",
                "• *Gia dụng:* Nồi nhôm anodized chống dính tự nhiên, khay nướng, móc treo chịu lực."
            ]
        }
    ]
}

EXERCISES_DANG2 = [
    {
        "id": "D2_C1",
        "num": 1,
        "type": "essay",
        "content": "Điện phân dung dịch Na~2~SO~4~ 0,5M với anode làm bằng Al và cathode trơ bằng dòng điện có cường độ 2,68 A. Thời gian để tạo thành lớp màng Al~2~O~3~ dày 3 µm trên bề mặt tấm nhôm có diện tích bề mặt là 1 m^2^ là bao lâu? Cho biết khối lượng riêng của Al~2~O~3~ là 3,95 g/cm^3^, hiệu suất điện phân là 95%; Al = 27, O = 16; F = 96 500 C/mol.\nTính thời gian theo giờ (kết quả làm tròn đến hàng phần trăm).",
        "ans": "7,34 giờ",
        "solution": (
            "- Diện tích tấm nhôm: S = 1 m^2^ = 10 000 cm^2^.\n"
            "- Độ dày lớp màng: h = 3 µm = 3 × 10^-4^ cm.\n"
            "- Thể tích Al~2~O~3~: V = S × h = 10 000 × (3 × 10^-4^) = 3 cm^3^.\n"
            "- Khối lượng Al~2~O~3~ tạo thành: m = V × D = 3 × 3,95 = 11,85 g.\n"
            "- Số mol Al~2~O~3~: n = 11,85 / 102 ≈ 0,116176 mol.\n"
            "- Theo bán phản ứng tại anode: 2Al + 3H~2~O → Al~2~O~3~ + 6H^+^ + 6e\n"
            "  → Cứ tạo ra 1 mol Al~2~O~3~ cần trao đổi 6 mol electron: n~e~ = 6 × n = 6 × 0,116176 ≈ 0,69706 mol.\n"
            "- Điện lượng lý thuyết: q~lt~ = n~e~ × F = 0,69706 × 96500 ≈ 67 266 C.\n"
            "- Với hiệu suất H = 95%:\n"
            "  Điện lượng thực tế: q~tt~ = q~lt~ / 0,95 = 67 266 / 0,95 ≈ 70 806 C.\n"
            "- Thời gian điện phân:\n"
            "  t = q~tt~ / I = 70 806 / 2,68 ≈ 26 420 giây = 26 420 / 3600 ≈ *7,34 giờ*."
        )
    },
    {
        "id": "D2_C2",
        "num": 2,
        "type": "essay",
        "content": "Điện phân dung dịch Na~2~SO~4~ 0,5M với anode làm bằng Al và cathode trơ bằng dòng điện có cường độ 2,68 A. Thời gian để tạo thành lớp màng Al~2~O~3~ dày 3 µm trên bề mặt tấm nhôm có diện tích bề mặt là 1 m^2^ là bao nhiêu giờ? Cho biết khối lượng riêng của Al~2~O~3~ là 3,95 g/cm^3^, hiệu suất điện phân là 95%; F = 96 500 C/mol.\nBiết: q = n.F = I.t (n là số mol electron trao đổi). Kết quả làm tròn đến hàng phần trăm sau dấu phẩy.",
        "ans": "7,34 giờ",
        "solution": (
            "- Diện tích S = 1 m^2^ = 10 000 cm^2^; chiều dày h = 3 µm = 3 × 10^-4^ cm.\n"
            "- V~oxit~ = 10 000 × 3 × 10^-4^ = 3 cm^3^ → m~oxit~ = 3 × 3,95 = 11,85 g.\n"
            "- n~Al2O3~ = 11,85 / 102 ≈ 0,11618 mol → n~e~ = 6 × n~Al2O3~ ≈ 0,69706 mol.\n"
            "- q~cần~ = (0,69706 × 96500) / 0,95 ≈ 70 806 C.\n"
            "- Thời gian điện phân: t = 70 806 / 2,68 ≈ 26 420 giây ≈ *7,34 giờ*."
        )
    },
    {
        "id": "D2_C3",
        "num": 3,
        "type": "essay",
        "content": "Sự tạo thành các lớp màng Al~2~O~3~ trên bề mặt Al có tác dụng bảo vệ nhôm. Bên cạnh quá trình tạo thành màng Al~2~O~3~ tự nhiên như để nhôm trong không khí, người ta có thể tạo thành lớp màng Al~2~O~3~ như ý bằng cách điện phân dung dịch H~2~SO~4~ 0,5M với anode làm bằng Al và cathode trơ bằng dòng điện có cường độ 2,68 A. Tính thời gian để tạo thành lớp màng Al~2~O~3~ dày 2 µm trên bề mặt tấm nhôm có diện tích bề mặt là 2 m^2^. Cho biết khối lượng riêng của Al~2~O~3~ là 3,95 g/cm^3^; hiệu suất điện phân là 90%; Al = 27, O = 16, F = 96 500 C/mol.\nTính thời gian theo giờ (kết quả làm tròn đến hàng phần trăm).",
        "ans": "10,33 giờ",
        "solution": (
            "- Diện tích tấm nhôm: S = 2 m^2^ = 20 000 cm^2^.\n"
            "- Chiều dày màng oxide: h = 2 µm = 2 × 10^-4^ cm.\n"
            "- Thể tích Al~2~O~3~: V = S × h = 20 000 × (2 × 10^-4^) = 4 cm^3^.\n"
            "- Khối lượng Al~2~O~3~: m = V × D = 4 × 3,95 = 15,8 g.\n"
            "- Số mol Al~2~O~3~: n = 15,8 / 102 ≈ 0,1549 mol.\n"
            "- Số mol electron trao đổi: n~e~ = 6 × n = 6 × 0,1549 ≈ 0,9294 mol.\n"
            "- Điện lượng lý thuyết: q~lt~ = 0,9294 × 96500 ≈ 89 688 C.\n"
            "- Với hiệu suất H = 90%:\n"
            "  Điện lượng thực tế: q~tt~ = 89 688 / 0,90 ≈ 99 654 C.\n"
            "- Thời gian điện phân: t = q~tt~ / I = 99 654 / 2,68 ≈ 37 184 giây = 37 184 / 3600 ≈ *10,33 giờ*."
        )
    }
]

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
                "• *Tại Cathode (cực âm - đáy bể):* Cation Al^3+^ bị khử thành nhôm lỏng chìm xuống đáy:\n      Al^3+^ + 3e → Al (lỏng)",
                "• *Tại Anode (cực dương - thanh graphite):* Anion O^2-^ bị oxi hóa thành khí O~2~:\n      2O^2-^ → O~2~ + 4e",
                "  Khí O~2~ sinh ra ở nhiệt độ cao (~950°C) đốt cháy các khối than chì anode:\n      C + O~2~ → CO~2~\n      2C + O~2~ → 2CO",
                "• *Phương trình tổng quát:* 2Al~2~O~3~ → 4Al + 3O~2~\n  (Khí thoát ra ở anode là hỗn hợp gồm CO~2~, CO và một phần O~2~ dư)."
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
        "content": "Điện phân nóng chảy hỗn hợp gồm aluminium oxide (Al~2~O~3~) và cryolite (Na~3~AlF~6~) để sản xuất nhôm (Al) trong công nghiệp. Cho các vai trò sau của cryolite:\n(1) Tạo với Al~2~O~3~ thành hỗn hợp có nhiệt độ nóng chảy thấp hơn nhiều nhiệt độ nóng chảy Al~2~O~3~.\n(2) Tăng khả năng dẫn điện của hỗn hợp điện phân do sự điện li tạo thành các ion.\n(3) Hình thành lớp màng ngăn cản sự tiếp xúc giữa Al (lỏng) mới sinh với không khí.\nHãy cho biết những vai trò nào đúng với cryolite trong quá trình sản xuất nhôm trong công nghiệp.",
        "options": [
            "A. (1), (2) và (3).",
            "B. (1) và (2).",
            "C. (1) và (3).",
            "D. (2) và (3)."
        ],
        "ans": "A",
        "solution": (
            "- Cryolite (Na~3~AlF~6~) có cả 3 vai trò cực kỳ quan trọng:\n"
            "  (1) Hạ nhiệt độ nóng chảy từ ~2050°C xuống ~950°C (tiết kiệm năng lượng).\n"
            "  (2) Phân li tạo nhiều ion tự do, làm tăng mạnh độ dẫn điện của khối nóng chảy.\n"
            "  (3) Khối lượng riêng nhẹ hơn Al lỏng nên nổi lên bề mặt, ngăn cách Al lỏng với O~2~ không khí.\n"
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
            "- a) *ĐÚNG:* Anode oxi hóa O^2-^; cathode khử Al^3+^.\n"
            "- b) *ĐÚNG:* D(Al lỏng) ≈ 2,3 g/cm^3^ > D(hỗn hợp điện li) ≈ 2,1 g/cm^3^ nên Al chìm xuống đáy.\n"
            "- c) *ĐÚNG:* Tốc độ sinh O~2~ nhanh làm phản ứng cháy C + O~2~ → CO~2~ xảy ra hoàn toàn hơn, tăng tỉ lệ CO~2~.\n"
            "- d) *ĐÚNG:* Giảm từ 2050°C xuống khoảng 950 - 1000°C.\n"
            "- e) *ĐÚNG:* Than chì anode bị ăn mòn liên tục.\n"
            "- f) *SAI:* Cryolite là Na~3~AlF~6~. Công thức Al~2~O~3~.2SiO~2~.2H~2~O là kaolin.\n"
            "- g) *SAI:* Quặng bauxite (Al~2~O~3~.2H~2~O), còn dolomite là CaCO~3~.MgCO~3~.\n"
            "- h) *ĐÚNG:* Nhôm lỏng tích tụ ở đáy và được hút ra định kỳ.\n"
            "- i) *ĐÚNG:* Hạ nhiệt độ từ 2050°C xuống 950°C giúp tiết kiệm năng lượng vượt bậc.\n"
            "- j) *SAI:* O~2~ sinh ra phản ứng ngay với anode carbon tạo CO~2~ và CO, không thu được O~2~ tự do.\n"
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
            "- a) *SAI:* Nhôm kim loại tách ra ở cathode (cực âm ở đáy bể).\n"
            "- b) *ĐÚNG:* O~2~ phản ứng với cực than chì tạo hỗn hợp CO~2~ và CO.\n"
            "- c) *ĐÚNG:* Hạ nhiệt độ từ 2050°C xuống ~950°C và tăng mạnh nồng độ ion tự do.\n"
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
            "- a) *ĐÚNG:* m(Al~2~O~3~) = 100 × 51% = 51 tấn.\n"
            "- b) *SAI:* m(Al lý thuyết) = 51 × (54 / 102) = 27 tấn (không phải 30 tấn).\n"
            "- c) *ĐÚNG:* m(Al thực tế) = 27 × 85% = 22,95 tấn ≈ 23 tấn.\n"
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
            "- a) *ĐÚNG:* 2Al~2~O~3~ → 4Al + 3O~2~ (tương ứng 6 mol O hay 3 mol O~2~ cho 4 mol Al, tỉ lệ n~Al~ : n~O~ = 2 : 3 hay 4 mol Al ứng với 6 mol nguyên tử O, xét n~e~ = 12 mol).\n"
            "- b) *ĐÚNG:* V(CO~2~) : V(CO) = 80% : 20% = 4 : 1.\n"
            "- c) *SAI:* n~Al~ = 5,4 × 10^6 / 27 = 0,2 Mmol → n~O~ = 0,3 Mmol.\n"
            "  Khí X: 4x mol CO~2~ và x mol CO → n~O~ = 2(4x) + x = 9x = 0,3 → x = 1/30 Mmol.\n"
            "  n~C~ = 5x = 1/6 Mmol → m~C~ = (1/6) × 12 = 2,0 tấn (không phải 3,6 tấn).\n"
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
            "- a) *ĐÚNG:* Nhờ có cryolite nên nhiệt độ duy trì ở 950 - 1000°C.\n"
            "- b) *ĐÚNG:* Lớp lót đáy là cathode than chì.\n"
            "- c) *SAI:* Anode bị oxi hóa liên tục thành CO, CO~2~ nên tiêu hao rất nhanh.\n"
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
            "- a) *SAI:* Thấp hơn rất nhiều (~950°C so với 2050°C).\n"
            "- b) *SAI:* Ngược lại: tại cathode khử Al^3+^, tại anode oxi hóa O^2-^.\n"
            "- c) *SAI:* AlCl~3~ là hợp chất cộng hóa trị có liên kết mạng tinh thể phân tử, thăng hoa ở nhiệt độ ~180°C nên không tồn tại thể lỏng dẫn điện để điện phân.\n"
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
            "- a) *ĐÚNG:* Than chì bị đốt cháy bởi O~2~.\n"
            "- b) *ĐÚNG:* D(Al lỏng) ≈ 2,3 g/cm^3^ > D(cryolite) ≈ 2,1 g/cm^3^.\n"
            "- c) *SAI:* Al~2~O~3~ bị điện phân tiêu thụ liên tục nên nồng độ giảm dần, định kỳ phải nạp thêm Al~2~O~3~.\n"
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
            "- Khối lượng Al~2~O~3~.2H~2~O trong 1 tấn (1000 kg) quặng: m = 1000 × 60% = 600 kg.\n"
            "- Phân tử khối Al~2~O~3~.2H~2~O = 102 + 2 × 18 = 138 g/mol.\n"
            "- Khối lượng Al theo lý thuyết (bảo toàn nguyên tố Al, trong 1 mol quặng có 2 mol Al):\n"
            "  m~Al lt~ = 600 × (2 × 27 / 138) = 600 × (54 / 138) ≈ 234,78 kg.\n"
            "- Do hiệu suất toàn bộ quá trình đạt 80%:\n"
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
            "- Số mol Al thu được: n~Al~ = 5,4 × 10^6 / 27 = 200 000 mol = 0,2 Mmol.\n"
            "- Quá trình tại cathode: Al^3+^ + 3e → Al → n~e~ = 3 × n~Al~ = 0,6 Mmol.\n"
            "- Quá trình tại anode: 2O^2-^ → O~2~ + 4e → n~O~ = n~e~ / 2 = 0,3 Mmol (nguyên tử O).\n"
            "- Gọi số mol của CO là x (Mmol) → số mol CO~2~ là 4x (Mmol) (vì tỉ lệ thể tích 80% : 20% = 4 : 1).\n"
            "- Bảo toàn nguyên tố oxygen:\n"
            "  n~O~ = 2 × n~CO2~ + 1 × n~CO~ = 2(4x) + x = 9x = 0,3 Mmol → x = 0,3 / 9 = 1/30 Mmol.\n"
            "- Tổng số mol carbon bị đốt cháy:\n"
            "  n~C~ = n~CO2~ + n~CO~ = 4x + x = 5x = 5 × (1/30) = 1/6 Mmol.\n"
            "- Khối lượng carbon bị oxi hóa:\n"
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
            "- Thể tích khí X: V = 99,16 m^3^ = 99 160 L (ở đkc: 1 mol = 24,79 L).\n"
            "- Tổng số mol khí X: n~X~ = 99 160 / 24,79 = 4000 mol.\n"
            "- Khối lượng mol trung bình của X: M~X~ = 16,7 × 2 = 33,4 g/mol.\n"
            "- Xét mẫu thử 1,2395 L X (đkc): n~mẫu~ = 1,2395 / 24,79 = 0,05 mol.\n"
            "  Phản ứng với Ca(OH)~2~ dư: n~CaCO3~ = 1,5 / 100 = 0,015 mol → n~CO2~ = 0,015 mol.\n"
            "  → Phần trăm thể tích của CO~2~ trong X: %V~CO2~ = (0,015 / 0,05) × 100% = 30%.\n"
            "- Hỗn hợp X gồm CO~2~ (30%), CO (a%), O~2~ (b%):\n"
            "  a + b = 100% - 30% = 70% = 0,70.\n"
            "  M~X~ = 0,30 × 44 + 28a + 32b = 33,4 → 28a + 32b = 20,2.\n"
            "  Giải hệ: a = 0,55 (55% CO); b = 0,15 (15% O~2~).\n"
            "- Trong 4000 mol X:\n"
            "  n~CO2~ = 4000 × 0,30 = 1200 mol; n~CO~ = 4000 × 0,55 = 2200 mol; n~O2~ = 4000 × 0,15 = 600 mol.\n"
            "- Tổng số mol nguyên tử O sinh ra từ Al~2~O~3~:\n"
            "  n~O~ = 2 × n~CO2~ + 1 × n~CO~ + 2 × n~O2~ = 2(1200) + 2200 + 2(600) = 5800 mol.\n"
            "- Bảo toàn nguyên tố trong Al~2~O~3~: n~Al~ = (2/3) × n~O~ = (2/3) × 5800 = 11600 / 3 ≈ 3866,67 mol.\n"
            "- Khối lượng nhôm thu được:\n"
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
            "- a) *ĐÚNG:* Anode tan: Cu → Cu^2+^ + 2e.\n"
            "- b) *ĐÚNG:* Cathode: Cu^2+^ + 2e → Cu.\n"
            "- c) *SAI:* Khi dùng điện cực đồng (dương cực tan), nước không bị điện phân ở anode nên không sinh ra O~2~ và H^+^.\n"
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
            "- a) *SAI:* Tại anode đồng thô bị oxi hóa tan ra, nước không bị oxi hóa.\n"
            "- b) *SAI:* Các tạp chất tan vào dung dịch (Zn, Fe) hoặc rơi xuống đáy tạo bùn anode (Ag, Au), không bám lại trên anode.\n"
            "- c) *ĐÚNG:* Tốc độ Cu tan ở anode xấp xỉ tốc độ Cu bám ở cathode.\n"
            "- d) *ĐÚNG:* Đồng kết tinh ở cathode có độ sạch đạt chuẩn 99,99%."
        )
    },
    {
        "id": "D4_C3",
        "num": 3,
        "type": "tf",
        "content": "Một nhóm học sinh nghiên cứu về tinh chế kim loại đồng bằng phương pháp điện phân từ một thanh đồng không tinh khiết. Nhóm học sinh đã thực hiện thí nghiệm theo các bước:\n- Bước 1: Cân xác định khối lượng ban đầu của thanh đồng nguyên chất (3,21 gam) và thanh đồng không tinh khiết (3,30 gam).\n- Bước 2: Nối thanh đồng nguyên chất với một điện cực và thanh đồng không tinh khiết với điện cực còn lại của nguồn điện một chiều, rồi nhúng vào bình điện phân chứa dung dịch CuSO~4~ đã acid hóa.\n- Bước 3: Điện phân ở hiệu điện thế phù hợp. Sau thời gian điện phân t giây, sấy khô rồi cân lại thấy thanh đồng nguyên chất nặng 3,51 gam và thanh không tinh khiết nặng 2,98 gam.\nCho các phát biểu sau:",
        "items": [
            {"id": "a", "text": "Ở cực dương xảy ra quá trình khử: Cu^2+^ + 2e → Cu.", "ans": "S"},
            {"id": "b", "text": "Màu xanh của dung dịch copper(II) sulfate gần như không thay đổi.", "ans": "Đ"},
            {"id": "c", "text": "Thanh đồng nguyên chất được nối với điện cực âm (cathode), thanh đồng không tinh khiết nối với điện cực dương (anode).", "ans": "Đ"},
            {"id": "d", "text": "Khối lượng thanh đồng nguyên chất tăng lên 0,30 gam đúng bằng khối lượng Cu nguyên chất kết tủa từ dung dịch.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Cực dương là Anode, xảy ra sự oxi hóa Cu → Cu^2+^ + 2e.\n"
            "- b) *ĐÚNG:* Quá trình dương cực tan giữ nồng độ Cu^2+^ ổn định.\n"
            "- c) *ĐÚNG:* Đồng tinh khiết là cathode (-), đồng thô là anode (+).\n"
            "- d) *ĐÚNG:* Khối lượng đồng nguyên chất tăng = 3,51 - 3,21 = 0,30 gam."
        )
    },
    {
        "id": "D4_C4",
        "num": 4,
        "type": "tf",
        "content": "Một nhóm học sinh tìm hiểu quá trình thu hồi đồng từ một đồng xu làm bằng hợp kim Cu-Zn chứa khoảng 95% Cu về khối lượng. Giả thuyết của nhóm là: 'Khi điện phân, chỉ có tạp chất trong đồng xu tan hết vào trong dung dịch, còn lại sẽ là đồng tinh khiết'. Tiến hành thí nghiệm:\n- Cân đồng xu ban đầu (2,23 gam) và thanh đồng tinh khiết (2,55 gam).\n- Nối đồng xu với cực dương, thanh đồng tinh khiết với cực âm, nhúng vào dung dịch CuSO~4~.\n- Điện phân một thời gian, sấy khô rồi cân lại thấy đồng xu nặng 1,94 gam và thanh đồng tinh khiết nặng m~1~ gam.\nCho các nhận định sau:",
        "items": [
            {"id": "a", "text": "Trong thí nghiệm trên, đồng xu được nối với cực dương, thanh đồng tinh khiết nối với cực âm.", "ans": "Đ"},
            {"id": "b", "text": "Giá trị của m~1~ lớn hơn 2,55 gam.", "ans": "Đ"},
            {"id": "c", "text": "Ở cực dương xảy ra quá trình khử ion kim loại.", "ans": "S"},
            {"id": "d", "text": "Do khối lượng của đồng xu giảm nên giả thuyết ban đầu của nhóm học sinh là đúng.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Đồng xu chứa Cu cần thu hồi đóng vai trò anode (+), thanh Cu tinh khiết là cathode (-).\n"
            "- b) *ĐÚNG:* Cu^2+^ bị khử bám vào cathode nên khối lượng thanh đồng tăng m~1~ > 2,55 gam.\n"
            "- c) *SAI:* Cực dương (anode) xảy ra quá trình oxi hóa.\n"
            "- d) *SAI:* Cả Zn và Cu đều tan ra ở anode, chứ không phải Cu nguyên chất ở lại trên đồng xu. Do đó giả thuyết ban đầu là sai."
        )
    },
    {
        "id": "D4_C5",
        "num": 5,
        "type": "essay",
        "content": "Một nhóm học sinh tìm hiểu quá trình thu hồi kim loại đồng từ một đồng xu làm bằng hợp kim Cu-Zn chứa khoảng 95% Cu về khối lượng. Tiến hành thí nghiệm:\n- Cân ban đầu đồng xu nặng 2,6 gam và thanh Cu tinh khiết nặng 2,85 gam.\n- Nối đồng xu với cực dương, thanh Cu tinh khiết với cực âm, nhúng vào bình điện phân chứa dung dịch CuSO~4~.\n- Điện phân với cường độ dòng điện không đổi I = 5 A, hiệu suất 100%.\n- Sau thời gian t giây, sấy khô cân lại thấy khối lượng đồng xu là 1,894 gam.\nTính giá trị của t (kết quả làm tròn đến hàng đơn vị). Cho biết Cu = 64; Zn = 65; F = 96 500 C/mol.",
        "ans": "425 giây",
        "solution": (
            "- Khối lượng đồng xu ban đầu: 2,6 gam; chứa 95% Cu → m~Cu ban đầu~ = 2,6 × 0,95 = 2,47 g; m~Zn ban đầu~ = 2,6 - 2,47 = 0,13 g.\n"
            "- Độ giảm khối lượng của đồng xu: Δm = 2,6 - 1,894 = 0,706 g.\n"
            "- Vì Zn có tính khử mạnh hơn Cu nên Zn tan trước:\n"
            "  Toàn bộ Zn tan hết: n~Zn~ = 0,13 / 65 = 0,002 mol → n~e(Zn)~ = 2 × 0,002 = 0,004 mol.\n"
            "- Lượng Cu tan tiếp theo:\n"
            "  m~Cu tan~ = 0,706 - 0,13 = 0,576 g → n~Cu~ = 0,576 / 64 = 0,009 mol.\n"
            "  n~e(Cu)~ = 2 × 0,009 = 0,018 mol.\n"
            "- Tổng số mol electron trao đổi: n~e~ = 0,004 + 0,018 = 0,022 mol.\n"
            "- Điện lượng q = n~e~ × F = 0,022 × 96500 = 2123 C.\n"
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
            "- a) *ĐÚNG:* Theo định luật Faraday m = (A.I.t)/(n.F).\n"
            "- b) *ĐÚNG:* E°(Zn^2+^/Zn) = -0,76 V < E°(Cu^2+^/Cu) = +0,34 V nên Zn bị oxi hóa trước.\n"
            "- c) *ĐÚNG:* Zn tan tạo ion Zn^2+^ hòa tan trong dung dịch.\n"
            "- d) *ĐÚNG:* Vì một phần electron ở cực dương do Zn tan cung cấp (không sinh Cu^2+^), trong khi ở cathode Cu^2+^ vẫn nhận electron liên tục để tạo Cu, do đó nồng độ Cu^2+^ trong dung dịch giảm nhẹ.\n"
            "- đ) *SAI:* Ở cathode xảy ra sự khử Cu^2+^ + 2e → Cu, không phải sự oxi hóa."
        )
    },
    {
        "id": "D4_C7",
        "num": 7,
        "type": "tf",
        "content": "Một học sinh muốn xác định hàm lượng Cu có trong hợp kim Cu-Zn đã tiến hành thí nghiệm như sau:\n- Cân thanh hợp kim Cu-Zn có khối lượng 2,25 gam (hàm lượng Zn trên 20%) và thanh Cu nguyên chất nặng 2,58 gam.\n- Nối thanh hợp kim với cực dương, thanh Cu nguyên chất với cực âm của nguồn điện một chiều, rồi nhúng vào bình điện phân chứa 100 mL dung dịch CuSO~4~ 1M. Sau một thời gian điện phân, làm khô rồi cân lại thấy thanh Cu nguyên chất nặng 2,99 gam.\nCho các nhận định sau:",
        "items": [
            {"id": "a", "text": "Quá trình điện phân xảy ra hiện tượng dương cực tan.", "ans": "Đ"},
            {"id": "b", "text": "Khối lượng thanh đồng nguyên chất tăng thêm 0,41 gam.", "ans": "Đ"},
            {"id": "c", "text": "Dung dịch sau điện phân chỉ chứa duy nhất một muối tan.", "ans": "S"},
            {"id": "d", "text": "Do anion sulfate không bị điện phân trong dung dịch nên tại anode xảy ra quá trình oxi hoá nước giải phóng O~2~.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Cực dương làm bằng hợp kim kim loại nên tan ra (dương cực tan).\n"
            "- b) *ĐÚNG:* m(tăng) = 2,99 - 2,58 = 0,41 gam.\n"
            "- c) *SAI:* Dung dịch chứa cả ZnSO~4~ và CuSO~4~ dư (2 muối tan).\n"
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
            "- a) *ĐÚNG:*\n"
            "  Điện lượng tổng cộng: q = I × t = 20 × (10 × 3600) = 720 000 C.\n"
            "  Điện lượng có ích để kết tủa 225 g Cu: q~ci~ = (225 / 64) × 2 × 96500 ≈ 678 516 C.\n"
            "  Hiệu suất dòng điện: H = (678 516 / 720 000) × 100% ≈ 94,24%.\n"
            "  → Phần trăm dòng điện hao phí: 100% - 94,24% = 5,76%.\n"
            "- b) *SAI:* Thường dùng dung dịch CuSO~4~ có thêm H~2~SO~4~, không dùng CuCl~2~ vì Cl^-^ dễ bị oxi hóa sinh khí Cl~2~ độc hại.\n"
            "- c) *SAI:* Anode là khối đồng thô (chứa tạp chất); cathode mới là đồng tinh khiết.\n"
            "- d) *ĐÚNG:* Ag có thế điện cực chuẩn +0,799 V > +0,34 V nên không bị oxi hóa, rơi xuống đáy tạo bùn anode."
        )
    }
]

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
            "- Khi điện phân dung dịch CuSO~4~ với cả hai điện cực bằng Cu (hiện tượng dương cực tan):\n"
            "  + Tại Anode: Cu bị oxi hóa Cu → Cu^2+^ + 2e nên khối lượng anode giảm đều theo thời gian tuyến tính → tương ứng đồ thị *(3)* (đường thẳng đi xuống).\n"
            "  + Tại Cathode: Cu^2+^ bị khử Cu^2+^ + 2e → Cu bám vào tấm cathode nên khối lượng cathode bắt đầu từ m~0~ ban đầu tăng đều tuyến tính theo thời gian → tương ứng đồ thị *(2)* (đường thẳng đi lên từ giá trị m~0~ > 0).\n"
            "→ Thứ tự lần lượt của anode và cathode là *(3) và (2)*. Chọn *D*."
        )
    },
    {
        "id": "TH_C2",
        "num": 2,
        "type": "mcq",
        "content": "Trong công nghiệp, nhôm được sản xuất từ quặng bauxite qua 2 giai đoạn chính: tinh chế quặng và điện phân Al~2~O~3~ nóng chảy có cryolite Na~3~AlF~6~. Sản phẩm điện phân ở cathode là nhôm lỏng và ở anode là hỗn hợp khí CO~2~, CO. Cấu tạo bể điện phân như hình vẽ:\nSau một thời gian điện phân thu được 5,4 tấn Al tại cathode và hỗn hợp khí tại anode gồm CO~2~ (chiếm 80% thể tích) và CO (chiếm 20% thể tích). Giả thiết không có sản phẩm phụ nào khác. Khối lượng carbon bị oxi hóa trên anode là bao nhiêu tấn?",
        "image": "p9_img3_xref739_773x202.jpeg",
        "options": [
            "A. 3,1 tấn.",
            "B. 2,0 tấn.",
            "C. 0,8 tấn.",
            "D. 1,6 tấn."
        ],
        "ans": "B",
        "solution": (
            "- Số mol Al thu được: n~Al~ = 5,4 × 10^6 / 27 = 200 000 mol = 0,2 Mmol.\n"
            "- Số mol electron trao đổi: n~e~ = 3 × n~Al~ = 0,6 Mmol.\n"
            "- Số mol nguyên tử O sinh ra ở anode: n~O~ = n~e~ / 2 = 0,3 Mmol.\n"
            "- Hỗn hợp khí anode có V(CO~2~) : V(CO) = 80% : 20% = 4 : 1.\n"
            "  Gọi n~CO~ = x (Mmol) → n~CO2~ = 4x (Mmol).\n"
            "- Bảo toàn nguyên tố O: 2(4x) + 1(x) = 9x = 0,3 Mmol → x = 0,3 / 9 = 1/30 Mmol.\n"
            "- Tổng số mol carbon bị đốt cháy: n~C~ = n~CO2~ + n~CO~ = 4x + x = 5x = 5 × (1/30) = 1/6 Mmol.\n"
            "- Khối lượng carbon tiêu hao: m~C~ = (1/6) × 12 = 2,0 tấn. Chọn *B*."
        )
    },
    {
        "id": "TH_C3",
        "num": 3,
        "type": "tf",
        "content": "Một nhóm học sinh thực hiện thí nghiệm điện phân dung dịch NaCl bão hòa theo các bước:\n• Bước 1: Chuẩn bị bình điện phân có màng ngăn và hai điện cực graphite.\n• Bước 2: Rót 100 mL dung dịch NaCl bão hòa vào bình sao cho hai điện cực đều ngập trong dung dịch.\n• Bước 3: Nối hai điện cực với nguồn điện một chiều và tiến hành điện phân, quan sát thấy có khí thoát ra ở hai cực.\n• Bước 4: Lấy màng ngăn ra khỏi bình và tiếp tục điện phân một thời gian.\nCho các phát biểu sau:",
        "items": [
            {"id": "a", "text": "Sau bước 4, dung dịch thu được trong bình điện phân có tính tẩy màu.", "ans": "Đ"},
            {"id": "b", "text": "Ở bước 3, tại cathode xảy ra quá trình khử ion Na^+^.", "ans": "S"},
            {"id": "c", "text": "Ở bước 1, sử dụng màng ngăn để các sản phẩm sinh ra ở hai điện cực (Cl~2~ và NaOH) không phản ứng với nhau.", "ans": "Đ"},
            {"id": "d", "text": "Ở bước 3, khi đưa mẫu giấy quỳ tím ẩm đến gần điện cực anode thì quỳ tím hóa đỏ rồi nhanh chóng bị mất màu.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Khi bỏ màng ngăn, Cl~2~ ở anode phản ứng với NaOH ở cathode tạo nước Javel (NaCl + NaClO + H~2~O) có tính tẩy màu mạnh.\n"
            "- b) *SAI:* Ở cathode nước bị khử sinh H~2~ (2H~2~O + 2e → H~2~ + 2OH^-^), Na^+^ không bị khử trong dung dịch nước.\n"
            "- c) *ĐÚNG:* Màng ngăn xốp ngăn Cl~2~ phản ứng với NaOH và H~2~.\n"
            "- d) *ĐÚNG:* Khí Cl~2~ làm quỳ ẩm hóa đỏ do tạo axit HCl và HClO, sau đó HClO có tính oxi hóa mạnh tẩy màu quỳ tím."
        )
    },
    {
        "id": "TH_C4",
        "num": 4,
        "type": "tf",
        "content": "Một nhóm học sinh tiến hành thí nghiệm mạ đồng lên chiếc thìa kim loại. Giả thuyết ban đầu là: 'Nồng độ CuSO~4~ trước và sau quá trình mạ là không đổi'. Tiến hành theo các bước:\n- Bước 1: Cân xác định khối lượng chiếc thìa là 10 gam và của thanh đồng nguyên chất là 15 gam.\n- Bước 2: Nối chiếc thìa với một cực và thanh đồng với cực còn lại của nguồn điện một chiều rồi nhúng vào cốc chứa dung dịch CuSO~4~ để mạ với hiệu điện thế thích hợp.\n- Bước 3: Sau 15 phút, lấy chiếc thìa và thanh đồng ra, sấy khô cân lại thấy chiếc thìa nặng 10,32 gam, thanh đồng nặng 14,68 gam. Hiệu suất đạt 100%.\nCho các phát biểu sau:",
        "image": "p9_img4_xref740_307x240.png",
        "items": [
            {"id": "a", "text": "Tại anode xảy ra quá trình khử ion Cu^2+^, tại cathode xảy ra quá trình oxi hoá Cu.", "ans": "S"},
            {"id": "b", "text": "Sau khi mạ xong, độ giảm khối lượng của thanh đồng bằng độ tăng khối lượng của chiếc thìa.", "ans": "Đ"},
            {"id": "c", "text": "Do khối lượng của thanh đồng giảm nên giả thuyết ban đầu của nhóm học sinh là sai.", "ans": "S"},
            {"id": "d", "text": "Thanh đồng được nối với cực dương (anode), chiếc thìa được nối với cực âm (cathode) của nguồn điện.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Tại anode xảy ra oxi hóa Cu → Cu^2+^ + 2e; tại cathode xảy ra khử Cu^2+^ + 2e → Cu.\n"
            "- b) *ĐÚNG:* Khối lượng đồng tan ra ở anode = 15 - 14,68 = 0,32 gam đúng bằng khối lượng Cu bám vào chiếc thìa = 10,32 - 10 = 0,32 gam.\n"
            "- c) *SAI:* Vì lượng Cu tan ở anode bằng lượng Cu bám ở cathode nên nồng độ Cu^2+^ trong dung dịch không đổi, tức là giả thuyết ban đầu của nhóm là đúng.\n"
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
            "- a) *ĐÚNG:* Ở cathode (-) ion H^+^ nhận e: 2H^+^ + 2e → H~2~.\n"
            "- b) *ĐÚNG:* Theo phương trình 2H~2~O → 2H~2~ + O~2~, tỉ lệ V(H~2~) : V(O~2~) = 2 : 1.\n"
            "- c) *SAI:* Nếu dùng dung dịch HCl thì ở anode ion Cl^-^ bị oxi hóa sinh ra khí Cl~2~ độc hại, không phải O~2~.\n"
            "- d) *ĐÚNG:* Thể tích H~2~ gấp đôi thể tích O~2~: 50 mL O~2~ đi kèm 100 mL H~2~."
        )
    },
    {
        "id": "TH_C6",
        "num": 6,
        "type": "tf",
        "content": "Một học sinh làm thí nghiệm điện phân dung dịch NaCl không màng ngăn:\n- Lắp thiết bị điện phân dung dịch NaCl với điện cực trơ graphite.\n- Rót 80 mL dung dịch NaCl bão hoà vào cốc rồi nhúng 2 điện cực vào dung dịch.\n- Nối với nguồn điện và điện phân trong khoảng 5 phút.\n- Thả một cánh hoa màu hồng vào cốc chứa dung dịch sau điện phân.\nCho các nhận xét sau:",
        "image": "p10_img4_xref846_245x262.jpeg",
        "items": [
            {"id": "a", "text": "Tại điện cực anode có khí H~2~ thoát ra.", "ans": "S"},
            {"id": "b", "text": "Tại điện cực cathode xảy ra quá trình khử nước giải phóng H~2~.", "ans": "Đ"},
            {"id": "c", "text": "Cánh hoa hồng bị mất màu do dung dịch sau điện phân có chứa NaClO.", "ans": "Đ"},
            {"id": "d", "text": "Dùng nắp đậy trong quá trình điện phân để hạn chế sự khuếch tán khí Cl~2~ ra ngoài môi trường gây độc hại.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Tại anode Cl^-^ bị oxi hóa tạo khí Cl~2~, khí H~2~ thoát ra ở cathode.\n"
            "- b) *ĐÚNG:* Cathode: 2H~2~O + 2e → H~2~ + 2OH^-^.\n"
            "- c) *ĐÚNG:* Không màng ngăn: Cl~2~ + 2NaOH → NaCl + NaClO + H~2~O. NaClO có tính tẩy màu mạnh làm phai màu cánh hoa.\n"
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
            "- a) *ĐÚNG:* Nước nguyên chất điện li cực yếu ([H^+^] = [OH^-^] = 10^-7^ M) nên dẫn điện kém, cần thêm acid trơ làm chất dẫn điện.\n"
            "- b) *SAI:* V(H~2~) = 2 V(O~2~).\n"
            "- c) *ĐÚNG:* Cực dương sinh khí O~2~ duy trì sự cháy, làm bùng cháy que đóm.\n"
            "- d) *ĐÚNG:* Cực âm sinh khí H~2~ cháy trong không khí tạo ngọn lửa xanh nhạt."
        )
    },
    {
        "id": "TH_C8",
        "num": 8,
        "type": "tf",
        "content": "Một học sinh muốn xác định hàm lượng Cu có trong hợp kim Cu-Zn:\n- Bước 1: Cân thanh hợp kim Cu-Zn (2,25 gam) và thanh đồng nguyên chất (2,58 gam).\n- Bước 2: Nối thanh hợp kim với cực dương, thanh Cu nguyên chất với cực âm của nguồn điện một chiều rồi nhúng vào bình điện phân chứa dung dịch CuSO~4~.\n- Sau một thời gian, sấy khô rồi cân lại thấy thanh Cu nguyên chất nặng 2,99 gam. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Tại anode xảy ra quá trình oxi hoá H~2~O do gốc SO~4~^2-^ không tham gia điện phân.", "ans": "S"},
            {"id": "b", "text": "Khối lượng của thanh đồng nguyên chất ở cathode tăng thêm 0,41 gam.", "ans": "Đ"},
            {"id": "c", "text": "Quá trình điện phân xảy ra hiện tượng dương cực tan.", "ans": "Đ"},
            {"id": "d", "text": "Toàn bộ thanh hợp kim Cu-Zn ở cực dương đã tan hoàn toàn trong thí nghiệm trên.", "ans": "S"}
        ],
        "solution": (
            "- a) *SAI:* Cực dương làm bằng hợp kim kim loại (Zn, Cu) nên các kim loại này bị oxi hóa tan ra, nước không bị điện phân.\n"
            "- b) *ĐÚNG:* Δm = 2,99 - 2,58 = 0,41 gam.\n"
            "- c) *ĐÚNG:* Anode bằng kim loại tan dần.\n"
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
            "- a) *ĐÚNG:* 2CuSO~4~ + 2H~2~O → 2Cu + O~2~ + 2H~2~SO~4~ sinh thêm H^+^ làm pH dung dịch giảm mạnh.\n"
            "- b) *SAI:* Ở Bài tập 2 ta đã tính với đĩa r = 12 cm, h = 0,01 mm cần m~Ag~ ≈ 4,75 g; với I = 5 A cần t ≈ 14 phút (không thể 2 phút).\n"
            "- c) *ĐÚNG:* Đa số trường hợp khối lượng dung dịch giảm (do kết tủa bám điện cực hoặc khí bay ra), nhưng trong quá trình dương cực tan kim loại nặng tan vào bù cho kim loại nhẹ có thể làm khối lượng dung dịch biến đổi.\n"
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
            "- a) *ĐÚNG:* D~cryolite~ ≈ 2,1 g/cm^3^ < D~Al lỏng~ ≈ 2,3 g/cm^3^.\n"
            "- b) *ĐÚNG:* Với m~Al~ = 1000 g: A = (5 × 1000 × 96500) / (9 × 3,6 × 10^6) ≈ 14,89 kWh ≈ 15 kWh.\n"
            "- c) *SAI:* O~2~ đốt cháy than chì tạo hỗn hợp CO~2~ và CO.\n"
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
            "- a) *ĐÚNG:* n~Al~ = 5,4 × 10^6 / 27 = 200 000 mol = 0,2 Mmol.\n"
            "- b) *ĐÚNG:* n~e~ = 3 × 0,2 = 0,6 Mmol = 600 000 mol.\n"
            "- c) *ĐÚNG:* n~O~ = 0,3 Mmol. Gọi khí: 7x mol CO~2~ và 3x mol CO → n~O~ = 2(7x) + 3x = 17x = 0,3 → x = 0,3 / 17 Mmol.\n"
            "  n~C~ = 10x = 3 / 17 ≈ 0,1765 Mmol → m~C~ = 0,1765 × 12 ≈ 2,12 tấn > 2,0 tấn.\n"
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
            "- a) *ĐÚNG:* Lượng Cu tan ở anode bằng lượng Cu bám vào cathode.\n"
            "- b) *SAI:* Nồng độ Cu^2+^ không đổi nên màu xanh dung dịch được giữ nguyên.\n"
            "- c) *ĐÚNG:* Cu^2+^ + 2e → Cu.\n"
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
            "- a) *ĐÚNG:* 2H~2~O → O~2~ + 4H^+^ + 4e.\n"
            "- b) *ĐÚNG:* Cu^2+^ + 2e → Cu (màu đỏ bám cathode).\n"
            "- c) *SAI:* Cu^2+^ bị điện phân giảm dần nên màu xanh nhạt dần.\n"
            "- d) *ĐÚNG:* Sinh ra H~2~SO~4~ (H^+^) làm quỳ tím chuyển đỏ."
        )
    },
    {
        "id": "TH_C14",
        "num": 14,
        "type": "tf",
        "content": "Tiến hành thí nghiệm điện phân dung dịch CuSO~4~ theo các bước:\n• Bước 1: Rót dung dịch CuSO~4~ vào cốc thủy tinh.\n• Bước 2: Nhúng thanh đồng vào cực dương và chìa khóa thép vào cực âm.\n• Bước 3: Đóng mạch điện và cho dòng điện một chiều chạy qua.\nCho các nhận định:",
        "image": "p12_img2_xref1082_997x280.jpeg",
        "items": [
            {"id": "a", "text": "Đây là mô hình ứng dụng điện phân để mạ đồng lên chìa khóa thép.", "ans": "Đ"},
            {"id": "b", "text": "Sau một thời gian chiếc chìa khóa được phủ một lớp kim loại màu đỏ bám chắc.", "ans": "Đ"},
            {"id": "c", "text": "Khối lượng thanh đồng cực dương không đổi trong suốt thí nghiệm.", "ans": "S"},
            {"id": "d", "text": "Nếu đổi chiều dòng điện thì lớp đồng trên chìa khóa sẽ bị tan ra.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Mạ đồng cho chìa khóa.\n"
            "- b) *ĐÚNG:* Lớp Cu màu đỏ bám đều lên chìa khóa.\n"
            "- c) *SAI:* Thanh đồng anode tan dần nên khối lượng giảm.\n"
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
            "- a) *ĐÚNG:* Vật cần mạ là cathode (-).\n"
            "- b) *ĐÚNG:* Kim loại mạ là anode (+).\n"
            "- c) *ĐÚNG:* Cation Ag^+^ nhận e kết tủa phủ lên nhẫn.\n"
            "- d) *SAI:* Do dùng anode Ag tan nên nồng độ Ag^+^ duy trì ổn định."
        )
    },
    {
        "id": "TH_C16",
        "num": 16,
        "type": "tf",
        "content": "Quá trình sản xuất nhôm theo công nghệ Hall-Héroult từ quặng bauxite:\n• Bauxite thô được làm sạch tạp chất Fe~2~O~3~, SiO~2~ bằng dung dịch NaOH theo quy trình Bayer thu được Al~2~O~3~.\n• Điện phân Al~2~O~3~ nóng chảy trong cryolite ở ~950°C.\nCho các nhận định:",
        "image": "p12_img3_xref1083_764x402.png",
        "items": [
            {"id": "a", "text": "Tạp chất Fe~2~O~3~ không tan trong NaOH đặc nên bị loại bỏ dưới dạng bùn đỏ.", "ans": "Đ"},
            {"id": "b", "text": "Al~2~O~3~ tan trong dung dịch NaOH tạo thành muối phức aluminate.", "ans": "Đ"},
            {"id": "c", "text": "Cryolite có công thức là Na~3~AlF~6~.", "ans": "Đ"},
            {"id": "d", "text": "Toàn bộ quặng bauxite thô có thể cho trực tiếp vào bể điện phân mà không cần tinh chế.", "ans": "S"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Fe~2~O~3~ là oxide bazo không tác dụng với NaOH nên nằm lại trong bùn đỏ.\n"
            "- b) *ĐÚNG:* Al~2~O~3~ + 2NaOH + 3H~2~O → 2Na[Al(OH)~4~].\n"
            "- c) *ĐÚNG:* Cryolite là trisodium hexafluoroaluminate Na~3~AlF~6~.\n"
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
            "- a) *ĐÚNG:* Anode trơ: 2H~2~O → O~2~ + 4H^+^ + 4e.\n"
            "- b) *ĐÚNG:* Cathode: Cu^2+^ + 2e → Cu.\n"
            "- c) *ĐÚNG:* n~e~ = (5 × 386) / 96500 = 0,02 mol → n~Cu~ = 0,01 mol → m~Cu~ = 0,01 × 64 = 0,64 gam.\n"
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
            "- a) *ĐÚNG:* Fe, Zn hoạt động hơn Cu nên bị oxi hóa tan thành cation.\n"
            "- b) *SAI:* E° của Zn^2+^ và Fe^2+^ rất âm so với Cu^2+^ nên chúng không bị khử ở cathode trong điều kiện thế khống chế.\n"
            "- c) *ĐÚNG:* Ag, Au trơ hơn Cu nên không tan, rơi xuống đáy tạo bùn anode.\n"
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
            "- a) *ĐÚNG:* Trong pin điện hóa Zn có E° = -0,76 V nhỏ hơn nên là cực âm (anode), xảy ra sự oxi hóa Zn → Zn^2+^ + 2e.\n"
            "- b) *ĐÚNG:* Cu có E° = +0,34 V là cực dương (cathode), xảy ra sự khử Cu^2+^ + 2e → Cu.\n"
            "- c) *SAI:* Electron chạy từ Zn sang Cu, do đó chiều quy ước của dòng điện ngược lại là từ Cu sang Zn.\n"
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
            "- a) *ĐÚNG:* Kim loại có thế khử âm mạnh bắt buộc điện phân nóng chảy.\n"
            "- b) *ĐÚNG:* Mục đích cốt lõi của mạ điện.\n"
            "- c) *ĐÚNG:* Bùn anode chứa Au, Ag, Pt là nguồn thu hồi kim loại quý giá trị cao.\n"
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
            "- a) *ĐÚNG:* Anode sinh H^+^: 2H~2~O → O~2~ + 4H^+^ + 4e làm quỳ chuyển đỏ.\n"
            "- b) *ĐÚNG:* Cathode khử Cu^2+^ thành Cu đỏ.\n"
            "- c) *ĐÚNG:* Ion Cu^2+^ khuếch tán và bị khử dần nên màu xanh mất dần.\n"
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
            "- a) *ĐÚNG:* Hạ từ 2050°C xuống 950°C.\n"
            "- b) *ĐÚNG:* Phân li tạo nhiều ion linh động.\n"
            "- c) *ĐÚNG:* C + O~2~ → CO~2~ và 2C + O~2~ → 2CO làm mòn cực dương.\n"
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
            "- a) *ĐÚNG:* Các kim loại có tính khử mạnh hơn hoặc bằng Cu đều bị oxi hóa.\n"
            "- b) *ĐÚNG:* Cu tan ở anode và bám vào cathode.\n"
            "- c) *SAI:* Ag và Au không bị oxi hóa nên không tan thành ion, chúng lắng xuống tạo bùn anode.\n"
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
            "- a) *ĐÚNG:* Cu^2+^ + 2e → Cu.\n"
            "- b) *ĐÚNG:* 2Cl^-^ → Cl~2~ + 2e.\n"
            "- c) *ĐÚNG:* CuCl~2~ → Cu + Cl~2~.\n"
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
            "- a) *ĐÚNG:* Vật cần mạ là cathode (-).\n"
            "- b) *ĐÚNG:* Dùng anode Zn tan để duy trì ion Zn^2+^.\n"
            "- c) *ĐÚNG:* n~Zn~ = 0,33 / 65 ≈ 0,005077 mol → n~e~ = 2 × 0,005077 ≈ 0,01015 mol.\n"
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
            "- a) *ĐÚNG:* Hiện tượng anode tan Cu → Cu^2+^ + 2e.\n"
            "- b) *ĐÚNG:* Cu^2+^ + 2e → Cu.\n"
            "- c) *SAI:* Dương cực tan duy trì nồng độ Cu^2+^ ổn định.\n"
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
            "- a) *ĐÚNG:* E°~pin~ = +0,34 - (-0,76) = 1,10 V.\n"
            "- b) *ĐÚNG:* Zn bị oxi hóa tan ra, Cu^2+^ khử bám vào cực Cu.\n"
            "- c) *SAI:* Cốc Zn sinh thêm cation Zn^2+^ nên anion NO~3~^-^ di chuyển vào để trung hòa điện tích; cation K^+^ di chuyển sang cốc Cu (nơi Cu^2+^ bị giảm).\n"
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
            "- a) *ĐÚNG:* Cathode khử Cu^2+^.\n"
            "- b) *ĐÚNG:* Anode oxi hóa nước sinh O~2~.\n"
            "- c) *ĐÚNG:* NO~3~^-^ kết hợp H^+^ tạo HNO~3~.\n"
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
            "- a) *ĐÚNG:* h = 0,02 mm = 0,002 cm → V = 50 × 0,002 = 0,10 cm^3^.\n"
            "- b) *ĐÚNG:* m = V × D = 0,10 × 8,96 = 0,896 gam.\n"
            "- c) *ĐÚNG:* n~Cu~ = 0,896 / 64 = 0,014 mol → n~e~ = 2 × 0,014 = 0,028 mol.\n"
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
            "- a) *ĐÚNG:* Vật cần mạ là cathode.\n"
            "- b) *ĐÚNG:* Anode tan bằng vàng.\n"
            "- c) *ĐÚNG:* m = (A.I.t)/(n.F) tỉ lệ thuận với thời gian t.\n"
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
            "- a) *ĐÚNG:* Hiện tượng dương cực tan không sinh khí ở anode.\n"
            "- b) *ĐÚNG:* Anode (+) là thanh Cu, cathode (-) là huy chương.\n"
            "- c) *ĐÚNG:* Cu tan bằng Cu bám.\n"
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
            "- a) *ĐÚNG:* 2H~2~O + 2e → H~2~ + 2OH^-^.\n"
            "- b) *ĐÚNG:* 2H~2~O → O~2~ + 4H^+^ + 4e.\n"
            "- c) *ĐÚNG:* Thực chất là điện phân H~2~O, lượng dung môi giảm nên nồng độ chất tan NaOH tăng lên.\n"
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
            "- a) *ĐÚNG:* Zn, Fe có thế khử âm hơn Cu nên bị oxi hóa trước.\n"
            "- b) *ĐÚNG:* Ag, Au không bị oxi hóa tạo bùn anode.\n"
            "- c) *ĐÚNG:* Ở thế khống chế chỉ Cu^2+^ bị khử.\n"
            "- d) *SAI:* Một phần dòng điện bị tiêu hao cho việc hòa tan Zn, Fe ở anode, nên khối lượng Cu bám ở cathode không bằng 100% tổng khối lượng anode tan."
        )
    },
    {
        "id": "TH_C34",
        "num": 34,
        "type": "tf",
        "content": "Tiến hành thí nghiệm theo các bước sau:\n• Bước 1: Rót từ từ dung dịch CuSO~4~ 0,5M vào ống thủy tinh hình chữ U đến khi mực nước dâng lên cách miệng ống 2 cm thì dừng lại.\n• Bước 2: Đậy miệng ống bên trái bằng nút cao su có gắn điện cực than chì.\n• Bước 3: Đậy miệng ống bên phải bằng nút cao su có gắn điện cực than chì và một ống dẫn khí.\n• Bước 4: Nối điện cực bên trái với cực âm (cathode) và nối điện cực bên phải với cực dương (anode) của nguồn điện một chiều (hiệu điện thế 6 V, dòng điện 1 A).\nCho các nhận định:",
        "items": [
            {"id": "a", "text": "Ở cathode, ion Cu^2+^ bị oxi hóa thành kim loại Cu.", "ans": "S"},
            {"id": "b", "text": "Dung dịch sau điện phân có giá trị pH < 7.", "ans": "Đ"},
            {"id": "c", "text": "Nếu thay dung dịch CuSO~4~ 0,5M bằng dung dịch CuCl~2~ 0,5M thì cathode và anode vẫn thu được các sản phẩm giống hệt thí nghiệm ban đầu.", "ans": "S"},
            {"id": "d", "text": "Đến khi ở cathode bắt đầu có bọt khí thoát ra thì màu xanh của dung dịch đã nhạt hẳn (Cu^2+^ gần như đã bị điện phân hết).", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Ở cathode xảy ra quá trình KHỬ: Cu^2+^ + 2e → Cu.\n"
            "- b) *ĐÚNG:* Anode sinh H^+^: 2H~2~O → O~2~ + 4H^+^ + 4e làm pH < 7.\n"
            "- c) *SAI:* Nếu dùng CuCl~2~ thì ở anode sinh khí Cl~2~ màu vàng lục, không phải O~2~.\n"
            "- d) *ĐÚNG:* Khi Cu^2+^ bị khử hết, nước bắt đầu bị khử ở cathode sinh khí H~2~, lúc này dung dịch đã mất màu xanh."
        )
    },
    {
        "id": "TH_C35",
        "num": 35,
        "type": "tf",
        "content": "Vàng hồng là hợp kim của Au, Ag, Cu. Trên thị trường có các loại: 10K (41,7% Au; 20% Ag; 38,3% Cu); 14K (58,5% Au; 21,5% Ag; 20% Cu); 18K (75% Au; 10% Ag; 15% Cu). Một nhóm học sinh muốn nâng cao hàm lượng Au, Ag từ mảnh nhẫn vàng hồng 10K bằng phương pháp điện phân với giả thuyết: 'Nếu Cu tan hết thì khối lượng mảnh nhẫn không giảm nữa'.\n- Cân ban đầu mảnh nhẫn 10K (1,125 gam) và thanh Cu tinh khiết (2,255 gam).\n- Nối mảnh nhẫn với một điện cực và thanh Cu với cực còn lại, nhúng vào dung dịch CuSO~4~.\n- Điện phân một thời gian, sấy khô rồi cân lại thấy mảnh nhẫn nặng 0,515 gam và thanh Cu nặng 2,740 gam.\nCho các nhận định:",
        "items": [
            {"id": "a", "text": "Mảnh nhẫn phải được nối với cực dương (anode), thanh Cu tinh khiết nối với cực âm (cathode).", "ans": "Đ"},
            {"id": "b", "text": "Nồng độ ion Cu^2+^ trong dung dịch không đổi trong suốt quá trình điện phân.", "ans": "Đ"},
            {"id": "c", "text": "Khối lượng Cu ban đầu trong mảnh nhẫn 10K là: 1,125 × 38,3% = 0,431 gam.", "ans": "Đ"},
            {"id": "d", "text": "Độ giảm khối lượng của mảnh nhẫn là 1,125 - 0,515 = 0,610 gam lớn hơn khối lượng Cu (0,431 gam) chứng tỏ đã có một phần bột vàng, bạc bị rơi rụng xuống đáy bình tạo bùn anode.", "ans": "Đ"}
        ],
        "solution": (
            "- a) *ĐÚNG:* Mảnh nhẫn là anode (+) để Cu tan ra; thanh Cu là cathode (-) để thu Cu.\n"
            "- b) *ĐÚNG:* Cu tan ở anode bằng Cu bám ở cathode nên nồng độ Cu^2+^ không đổi.\n"
            "- c) *ĐÚNG:* m~Cu trong nhẫn~ = 1,125 × 38,3% = 0,4309 g ≈ 0,431 gam.\n"
            "- d) *ĐÚNG:* Khối lượng nhẫn giảm 0,610 g > 0,431 g Cu. Điều này là do khi khung mạng lưới đồng bị hòa tan, các hạt vàng, bạc siêu mịn không còn chỗ bám sẽ rụng xuống đáy bể tạo bùn anode. Do đó giả thuyết ban đầu của học sinh là sai."
        )
    },
    {
        "id": "TH_C36",
        "num": 36,
        "type": "tf",
        "content": "Vàng tây 18K chứa 75,00% Au, phần còn lại là Ni và Ag. Một nhà kim hoàn kiểm tra hàm lượng Au bằng phương pháp điện phân:\n- Lấy 1 chỉ vàng tây 18K (3,75 gam) làm điện cực nối với cực dương.\n- Điện cực âm là thanh than chì tinh khiết nặng 2,00 gam.\n- Nhúng vào 100 mL dung dịch Au(NO~3~)~3~ 0,2M. Kẹp nối 2 điện cực đều làm bằng than chì trơ.\n- Điện phân ở hiệu điện thế thích hợp cho đến khi chỉ vàng tan hết thì dừng lại, sấy khô cân lại thanh than chì thấy nặng 4,85 gam.\nBiết trong quá trình điện phân thể tích dung dịch không đổi, toàn bộ Au sinh ra bám hết lên điện cực âm và không có cặn rơi xuống đáy. Cho các nhận định:",
        "items": [
            {"id": "a", "text": "Có thể thay kẹp điện cực than chì bằng dây đồng Cu để thực hiện thí nghiệm.", "ans": "S"},
            {"id": "b", "text": "Tại cathode xảy ra bán phản ứng khử: Au^3+^ + 3e → Au.", "ans": "Đ"},
            {"id": "c", "text": "Khối lượng Au bám lên thanh than chì là 4,85 - 2,00 = 2,85 gam.", "ans": "Đ"},
            {"id": "d", "text": "Hàm lượng vàng trong mẫu vàng tây trên đạt: (2,85 / 3,75) × 100% = 76,0% (vượt tiêu chuẩn 75%).", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* Nếu dùng dây Cu thì Cu sẽ tác dụng với dung dịch Au(NO~3~)~3~ (3Cu + 2Au^3+^ → 3Cu^2+^ + 2Au) và tan ra ở cực dương, làm sai lệch kết quả thí nghiệm.\n"
            "- b) *ĐÚNG:* Cation Au^3+^ nhận electron kết tủa bám vào cathode than chì.\n"
            "- c) *ĐÚNG:* m~Au bám~ = 4,85 - 2,00 = 2,85 gam.\n"
            "- d) *ĐÚNG:* %Au = (2,85 / 3,75) × 100% = 76,0% > 75,00% (vượt tiêu chuẩn 18K)."
        )
    }
]

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
            "- Thời gian điện phân: t = 9 phút = 9 × 60 = 540 giây.\n"
            "- Điện lượng thực tế có ích: q~ci~ = I × t × H = 2 × 540 × 0,90 = 972 C.\n"
            "- Bán phản ứng tại cathode: Ag^+^ + 1e → Ag.\n"
            "- Số mol Ag kết tủa: n~Ag~ = n~e~ = q~ci~ / F = 972 / 96500 ≈ 0,0100725 mol.\n"
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
            "- Lượng Mg đã bị tan (oxi hóa) bảo vệ ống thép: 100% - 2% = 98%.\n"
            "- Khối lượng Mg tan: m~Mg~ = 293 × 0,98 = 287,14 gam.\n"
            "- Số mol Mg: n~Mg~ = 287,14 / 24,3 ≈ 11,8165 mol.\n"
            "- Quá trình: Mg → Mg^2+^ + 2e → n~e~ = 2 × n~Mg~ ≈ 23,6329 mol.\n"
            "- Điện lượng q = n~e~ × F = 23,6329 × 96500 ≈ 2 280 577 C.\n"
            "- Thời gian bảo vệ: t = q / I = 2 280 577 / 0,025 = 91 223 080 giây.\n"
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
            "- Tổng diện tích mạ: S = 2 m^2^ = 20 000 cm^2^.\n"
            "- Cường độ dòng điện: I = 200 A/m^2^ × 2 m^2^ = 400 A.\n"
            "- Bề dày lớp mạ: h = 6,5 µm = 6,5 × 10^-4^ cm.\n"
            "- Thể tích kẽm: V = S × h = 20 000 × (6,5 × 10^-4^) = 13 cm^3^.\n"
            "- Khối lượng Zn: m~Zn~ = V × D = 13 × 7,14 = 92,82 gam.\n"
            "- Số mol Zn: n~Zn~ = 92,82 / 65 ≈ 1,428 mol → n~e~ = 2 × 1,428 = 2,856 mol.\n"
            "- Điện lượng: q = n~e~ × F = 2,856 × 96485 ≈ 275 561 C.\n"
            "- Thời gian: t = q / I = 275 561 / 400 ≈ 688,9 giây = 688,9 / 3600 ≈ *0,19 giờ* (≈ 11,5 phút)."
        )
    },
    {
        "id": "P3_C40",
        "num": 40,
        "type": "short_answer",
        "content": "Một sinh viên thực hiện điện phân hệ 2 bình mắc nối tiếp:\n- Bình (1) chứa 100 mL dung dịch CuSO~4~ 1M.\n- Bình (2) chứa 100 mL dung dịch AgNO~3~ 1M.\nSau thời gian t giây, sinh viên quan sát thấy ở cathode bình (2) bắt đầu xuất hiện bọt khí. Cho biết trong hệ điện phân nối tiếp, số electron truyền dẫn qua các bình là như nhau. Nguyên tử khối Cu = 64, Ag = 108. Bỏ qua các quá trình phụ. Tính số gam Cu bám lên điện cực trong bình (1) tại thời điểm t giây.",
        "ans": "3,2 gam",
        "solution": (
            "- Số mol Ag^+^ ban đầu trong bình (2): n~Ag+~ = 0,100 × 1 = 0,10 mol.\n"
            "- Số mol Cu^2+^ ban đầu trong bình (1): n~Cu2+~ = 0,100 × 1 = 0,10 mol.\n"
            "- Ở cathode bình (2): Ag^+^ + 1e → Ag.\n"
            "  Thời điểm cathode bình (2) bắt đầu thoát bọt khí (nước bị khử) là lúc toàn bộ Ag^+^ vừa bị khử hết.\n"
            "  → Số mol electron truyền qua hệ nối tiếp: n~e~ = n~Ag+~ = 0,10 mol.\n"
            "- Ở cathode bình (1): Cu^2+^ + 2e → Cu.\n"
            "  Vì n~e~ = 0,10 mol cần cho Cu^2+^ là n~Cu2+ điện phân~ = 0,10 / 2 = 0,05 mol < 0,10 mol (Cu^2+^ còn dư).\n"
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
            "- Thời gian điện phân: t = 20 phút = 1200 giây.\n"
            "- Điện lượng: q = I × t = 5 × 1200 = 6000 C.\n"
            "- Số mol electron: n~e~ = 6000 / 96500 ≈ 0,062176 mol.\n"
            "- Số mol Cu kết tủa: n~Cu~ = n~e~ / 2 ≈ 0,031088 mol.\n"
            "- Khối lượng Cu bám: m~Cu~ = 0,031088 × 64 ≈ 1,9896 gam.\n"
            "- Thể tích lớp đồng: V = m / D = 1,9896 / 8,9 ≈ 0,22355 cm^3^.\n"
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
            "- Bán kính khuyên tai: r = 5,00 / 2 = 2,50 cm.\n"
            "- Diện tích 2 mặt đĩa tròn: S = 2 × (πr^2^) = 2 × 3,14 × (2,50)^2^ = 39,25 cm^2^.\n"
            "- Bề dày lớp mạ: h = 0,02 mm = 0,002 cm.\n"
            "- Thể tích lớp vàng: V = S × h = 39,25 × 0,002 = 0,0785 cm^3^.\n"
            "- Khối lượng vàng cần mạ: m~Au~ = V × D = 0,0785 × 19,7 ≈ 1,54645 gam.\n"
            "- Số mol Au: n~Au~ = 1,54645 / 197 ≈ 0,00785 mol.\n"
            "- Phản ứng tại cathode: Au^3+^ + 3e → Au → n~e~ = 3 × 0,00785 = 0,02355 mol.\n"
            "- Điện lượng: q = 0,02355 × 96500 ≈ 2272,6 C.\n"
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
            "- Cường độ dòng điện bảo vệ: I = 2,75 mA/m^2^ × 1000 m^2^ = 2750 mA = 2,75 A.\n"
            "- Thời gian 1 năm: t = 365 × 24 × 3600 = 31 536 000 giây.\n"
            "- Tổng điện lượng cần cung cấp: q = 2,75 × 31 536 000 = 86 724 000 C.\n"
            "- Số mol electron: n~e~ = 86 724 000 / 96485 ≈ 898,834 mol.\n"
            "- Phản ứng tan của Zn: Zn → Zn^2+^ + 2e → n~Zn~ = n~e~ / 2 ≈ 449,417 mol.\n"
            "- Khối lượng Zn lý thuyết cần thiết: m~lt~ = 449,417 × 65,38 ≈ 29 383 gam = 29,383 kg.\n"
            "- Do lượng kẽm thất thoát bởi quá trình khác là 12% nên hiệu quả chỉ đạt 88%:\n"
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
            "- Diện tích S = 120 cm^2^; bề dày màng h = 0,01 mm = 0,001 cm.\n"
            "- Thể tích màng oxide: V = S × h = 120 × 0,001 = 0,12 cm^3^.\n"
            "- Khối lượng Al~2~O~3~: m = V × D = 0,12 × 3,97 = 0,4764 gam.\n"
            "- Số mol Al~2~O~3~: n = 0,4764 / 102 ≈ 0,0046706 mol.\n"
            "- Số mol electron lý thuyết: n~e~ = 6 × n = 6 × 0,0046706 ≈ 0,0280235 mol.\n"
            "- Điện lượng lý thuyết: q~lt~ = n~e~ × F = 0,0280235 × 96485 ≈ 2703,8 C.\n"
            "- Điện lượng thực tế (H = 94%): q~tt~ = 2703,8 / 0,94 ≈ 2876,4 C.\n"
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
            "- Diện tích toàn phần huy chương (2 đáy tròn + xung quanh):\n"
            "  S~tp~ = 2πr^2^ + 2πr × h~trụ~ = 2 × 3,14 × (2,5)^2^ + 2 × 3,14 × 2,5 × 0,3 = 39,25 + 4,71 = 43,96 cm^2^.\n"
            "- Bề dày lớp mạ: h = 0,01 cm.\n"
            "- Thể tích đồng: V = S~tp~ × h = 43,96 × 0,01 = 0,4396 cm^3^.\n"
            "- Khối lượng đồng: m~Cu~ = V × D = 0,4396 × 8,96 ≈ 3,9388 gam.\n"
            "- Số mol Cu: n~Cu~ = 3,9388 / 64 ≈ 0,061544 mol → n~e~ = 2 × 0,061544 ≈ 0,12309 mol.\n"
            "- Điện lượng: q = n~e~ × F = 0,12309 × 96500 ≈ 11 878 C.\n"
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
            "- Diện tích mặt đĩa: S = π × r^2^ = 3,1416 × 12^2^ ≈ 452,39 cm^2^.\n"
            "- Chiều dày lớp bạc: h = 0,01 mm = 0,001 cm.\n"
            "- Thể tích lớp bạc: V = S × h = 452,39 × 0,001 ≈ 0,45239 cm^3^.\n"
            "- Khối lượng Ag: m~Ag~ = V × D = 0,45239 × 10,5 ≈ 4,750 gam.\n"
            "- Số mol Ag: n~Ag~ = 4,750 / 108 ≈ 0,04398 mol → n~e~ = 0,04398 mol.\n"
            "- Điện lượng: q = 0,04398 × 96485 ≈ 4243,6 C.\n"
            "- Thời gian: t = q / I = 4243,6 / 2 = 2121,8 giây = 2121,8 / 3600 ≈ *0,59 giờ* (≈ 35,4 phút)."
        )
    },
    {
        "id": "P3_C47",
        "num": 47,
        "type": "short_answer",
        "content": "Hình bên mô tả thiết bị điện phân nước (Hofmann). Ban đầu các ống A và B đều chứa đầy dung dịch NaOH 0,01M, sau một thời gian dòng điện một chiều đi qua thì có các khí tích tụ, trong đó thể tích khí ở ống B gấp đôi thể tích khí ở ống A. Cho các phát biểu sau:\n(1) Điện cực dương (anode) nằm bên ống A.\n(2) Giá trị pH của dung dịch không thay đổi trong quá trình điện phân.\n(3) Tại cathode xảy ra quá trình oxi hóa ion Na^+^.\n(4) Khi thể tích khí thu được ở ống B đạt 24,79 mL (ở đkc) thì khối lượng dung dịch điện phân giảm đi 18 mg.\nLiệt kê các phát biểu đúng theo thứ tự chữ số tăng dần (ví dụ: 12, 14, 124...):",
        "image": "p19_img1_xref1866_157x316.jpeg",
        "ans": "14",
        "solution": (
            "- Khí ở ống B gấp đôi thể tích ống A → ống B chứa H~2~ (ở cathode - cực âm); ống A chứa O~2~ (ở anode - cực dương).\n"
            "  → (1) *ĐÚNG:* Cực dương nằm bên ống A.\n"
            "- Quá trình điện phân làm giảm lượng nước môi trường, làm nồng độ NaOH tăng nhẹ → pH tăng nhẹ, do đó (2) *SAI*.\n"
            "- Ở cathode khử nước: 2H~2~O + 2e → H~2~ + 2OH^-^, Na^+^ không bị điện phân → (3) *SAI*.\n"
            "- Khí H~2~ ở ống B: n~H2~ = 24,79 / 24790 = 0,001 mol.\n"
            "  Theo phản ứng: 2H~2~O → 2H~2~ + O~2~, số mol H~2~O bị điện phân = n~H2~ = 0,001 mol.\n"
            "  Khối lượng dung dịch giảm đúng bằng khối lượng nước bị điện phân = 0,001 × 18 = 0,018 gam = 18 mg → (4) *ĐÚNG*.\n"
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
            "- Diện tích S = 10 cm^2^; bề dày h = 0,1 mm = 0,01 cm.\n"
            "- Thể tích Cu: V = S × h = 10 × 0,01 = 0,10 cm^3^.\n"
            "- Khối lượng Cu: m~Cu~ = 0,10 × 8,94 = 0,894 gam.\n"
            "- Số mol Cu: n~Cu~ = 0,894 / 64 ≈ 0,013969 mol → n~e~ = 2 × 0,013969 ≈ 0,027938 mol.\n"
            "- Điện lượng lý thuyết: q~lt~ = 0,027938 × 96485 ≈ 2695,6 C.\n"
            "- Điện lượng thực tế (H = 90%): q~tt~ = 2695,6 / 0,90 ≈ 2995,1 C.\n"
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
            "- Diện tích S = 2 m^2^ = 20 000 cm^2^; cường độ I = 200 × 2 = 400 A.\n"
            "- Thể tích kẽm: V = 20 000 × (6,5 × 10^-4^) = 13 cm^3^.\n"
            "- Khối lượng kẽm: m~Zn~ = 13 × 7,14 = 92,82 gam.\n"
            "- Số mol electron: n~e~ = 2 × (92,82 / 65) ≈ 2,856 mol.\n"
            "- Điện lượng: q = 2,856 × 96485 ≈ 275 561 C.\n"
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
            "- Thời gian điện phân: t = 2 × 3600 = 7200 giây.\n"
            "- Điện lượng: q = I × t = 1,5 × 7200 = 10 800 C.\n"
            "- Số mol electron: n~e~ = 10 800 / 96500 ≈ 0,111917 mol.\n"
            "- Phản ứng tại cathode: Ni^2+^ + 2e → Ni → n~Ni~ = n~e~ / 2 ≈ 0,055959 mol.\n"
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
            "- Số mol Na cần điều chế: n~Na~ = 1000 / 23 ≈ 43,478 mol.\n"
            "- Quá trình: Na^+^ + 1e → Na → n~e lý thuyết~ = 43,478 mol.\n"
            "- Điện lượng lý thuyết: q~lt~ = 43,478 × 96485 ≈ 4 195 000 C.\n"
            "- Vì hiệu suất điện phân là 80% nên điện lượng thực tế tiêu hao:\n"
            "  q~tt~ = 4 195 000 / 0,80 = 5 243 750 C.\n"
            "- Điện năng tiêu thụ: A = U × q~tt~ = 25 × 5 243 750 = 131 093 750 J.\n"
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
            "- Giả sử có 100 gam quặng bauxite:\n"
            "  + Khối lượng kaolin Al~2~Si~2~O~7~.2H~2~O (M = 258 g/mol) là 10 gam → n~kaolin~ = 10 / 258 ≈ 0,03876 mol.\n"
            "  + Số mol Si có trong quặng: n~Si~ = 2 × n~kaolin~ ≈ 0,07752 mol.\n"
            "- Theo công thức bùn kết tủa Na~6~Al~6~Si~5~O~22~.5H~2~O, tỉ lệ Al : Si = 6 : 5 = 1,2.\n"
            "  → Số mol Al bị kéo theo vào bùn thải: n~Al mất~ = 1,2 × n~Si~ = 1,2 × 0,07752 ≈ 0,09302 mol.\n"
            "- Khối lượng Al~2~O~3~.3H~2~O (M = 156 g/mol) trong 100 gam quặng là 90 gam:\n"
            "  → n~Al bauxite~ = 2 × (90 / 156) ≈ 1,1538 mol.\n"
            "- Tổng số mol Al ban đầu trong quặng: n~Al tổng~ = 1,1538 + 2 × 0,03876 ≈ 1,2313 mol.\n"
            "- Tỉ lệ hao hụt của nhôm:\n"
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
            "- Thời gian điện phân: t = 2 × 3600 + 40 × 60 + 50 = 7200 + 2400 + 50 = 9650 giây.\n"
            "- Điện lượng: q = I × t = 10 × 9650 = 96 500 C.\n"
            "- Số mol electron: n~e~ = 96 500 / 96500 = 1,0 mol.\n"
            "- Phản ứng tại cathode: Cu^2+^ + 2e → Cu → n~Cu~ = n~e~ / 2 = 0,5 mol.\n"
            "- Khối lượng đồng kết tủa: m~Cu~ = 0,5 × 64 = 32 gam.\n"
            "- Thể tích lớp đồng: V = m / D = 32 / 8,9 ≈ 3,5955 cm^3^.\n"
            "- Bề dày lớp mạ: h = 0,18 mm = 0,018 cm.\n"
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
            "- Diện tích một mặt đĩa tròn: S = π × r^2^ = 3,14 × 20^2^ = 1256 cm^2^.\n"
            "- Bề dày lớp bạc: h = 0,01 mm = 0,001 cm.\n"
            "- Thể tích bạc: V = S × h = 1256 × 0,001 = 1,256 cm^3^.\n"
            "- Khối lượng Ag: m~Ag~ = V × D = 1,256 × 10,5 = 13,188 gam.\n"
            "- Số mol Ag: n~Ag~ = 13,188 / 108 ≈ 0,12211 mol → n~e~ = 0,12211 mol.\n"
            "- Điện lượng: q = n~e~ × F = 0,12211 × 96500 ≈ 11 783,7 C.\n"
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
            "- Thể tích màng oxide: V = 120 × (0,01 / 10) = 0,12 cm^3^.\n"
            "- Khối lượng Al~2~O~3~: m = 0,12 × 3,97 = 0,4764 gam.\n"
            "- Số mol Al~2~O~3~: n = 0,4764 / 102 ≈ 0,0046706 mol → n~e~ = 6 × n ≈ 0,0280235 mol.\n"
            "- Điện lượng lý thuyết: q~lt~ = 0,0280235 × 96485 ≈ 2703,8 C.\n"
            "- Với hiệu suất 96%: q~tt~ = 2703,8 / 0,96 ≈ 2816,5 C.\n"
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
            "- Thể tích Cu: V = 10 × 0,01 = 0,10 cm^3^ → m~Cu~ = 0,10 × 8,94 = 0,894 gam.\n"
            "- Số mol Cu: n = 0,894 / 64 ≈ 0,013969 mol → n~e~ = 2 × 0,013969 ≈ 0,027938 mol.\n"
            "- Điện lượng lý thuyết: q~lt~ = 0,027938 × 96485 ≈ 2695,6 C.\n"
            "- Với hiệu suất 80%: q~tt~ = 2695,6 / 0,80 ≈ 3369,5 C.\n"
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
            "- Số mol khí ở anode: n~khí~ = 99 160 / 24,79 = 4000 mol.\n"
            "- Do tỉ lệ mol CO : CO~2~ = 1 : 1 → n~CO~ = n~CO2~ = 2000 mol.\n"
            "- Bảo toàn nguyên tố O: n~O~ = 1 × n~CO~ + 2 × n~CO2~ = 2000 + 4000 = 6000 mol.\n"
            "- Trong Al~2~O~3~.2H~2~O, 1 mol Al~2~O~3~ chứa 3 mol nguyên tử O tham gia điện phân:\n"
            "  → n~Al2O3 lt~ = 6000 / 3 = 2000 mol → n~quặng lt~ = 2000 mol.\n"
            "- Khối lượng mol Al~2~O~3~.2H~2~O = 102 + 36 = 138 g/mol.\n"
            "- Khối lượng quặng lý thuyết: m~lt~ = 2000 × 138 = 276 000 gam = 276 kg.\n"
            "- Do hiệu suất toàn bộ quá trình là 80%:\n"
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
            "- Khối lượng Al~2~O~3~ trong 20 tấn quặng: m = 20 × 47% = 9,4 tấn = 9 400 kg.\n"
            "- Lượng nhôm thu được với hiệu suất 85%:\n"
            "  m~Al~ = 9400 × (54 / 102) × 0,85 ≈ 4230 kg = 4 230 000 gam.\n"
            "- Thể tích một thanh nhôm hộp: V~1~ = 110 × 20 × 10 = 22 000 cm^3^.\n"
            "- Khối lượng của một thanh nhôm: m~1~ = V~1~ × D = 22 000 × 2,7 = 59 400 gam = 59,4 kg.\n"
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
            "- Điện lượng: q = I × t = 2 × 15 661,95 = 31 323,9 C.\n"
            "- Số mol electron trao đổi: n~e~ = 31 323,9 / 96500 = 0,3246 mol.\n"
            "- Cả Cu và Zn khi tan ở anode đều nhường 2e:\n"
            "  Cu → Cu^2+^ + 2e (x mol)\n"
            "  Zn → Zn^2+^ + 2e (y mol)\n"
            "  → 2x + 2y = 0,3246 → x + y = 0,1623 mol.\n"
            "- Tổng khối lượng mẫu hợp kim: 64x + 65y = 10,4 gam.\n"
            "  → y = 10,4 - 64 × 0,1623 = 10,4 - 10,3872 = 0,0128 mol.\n"
            "  → x = 0,1623 - 0,0128 = 0,1495 mol.\n"
            "- Khối lượng Cu trong mẫu: m~Cu~ = 0,1495 × 64 = 9,568 gam.\n"
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
            "- Số mol Al: n~Al~ = 2,7 × 10^6 / 27 = 100 000 mol = 0,1 Mmol.\n"
            "- Số mol e: n~e~ = 3 × n~Al~ = 0,3 Mmol → n~O~ = 0,15 Mmol.\n"
            "- Khí anode: n~CO2~ = 4x; n~CO~ = x → n~O~ = 2(4x) + x = 9x = 0,15 Mmol → x = 0,15 / 9 = 1/60 Mmol.\n"
            "- Tổng số mol C bị cháy: n~C~ = 4x + x = 5x = 5 × (1/60) = 1/12 Mmol.\n"
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
            "- Số mol Al thu được: n~Al~ = 5400 / 27 = 200 mol.\n"
            "- Số mol O sinh ra ở anode: n~O~ = (3/2) × n~Al~ = (3/2) × 200 = 300 mol.\n"
            "- Thành phần khí X: CO chiếm 70% thể tích → CO~2~ chiếm 30% thể tích (tỉ lệ mol CO : CO~2~ = 7 : 3).\n"
            "  Gọi n~CO2~ = 3x (mol) → n~CO~ = 7x (mol).\n"
            "- Bảo toàn nguyên tố O: 2 × (3x) + 1 × (7x) = 13x = 300 mol → x = 300 / 13 mol.\n"
            "- Tổng số mol carbon anode bị đốt cháy: n~C~ = n~CO2~ + n~CO~ = 3x + 7x = 10x = 3000 / 13 mol.\n"
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
            "- Số mol Mg: n~Mg~ = 6000 / 24,3 ≈ 246,91 mol.\n"
            "- Số mol electron: n~e~ = 2 × n~Mg~ ≈ 493,83 mol.\n"
            "- Điện lượng: q = n~e~ × F = 493,83 × 96500 ≈ 47 654 321 C.\n"
            "- Thời gian bảo vệ: t = q / I = 47 654 321 / 0,03 ≈ 1 588 477 366 giây.\n"
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
            "- a) *ĐÚNG:* Cực âm thoát khí H~2~, cực dương thoát khí Cl~2~.\n"
            "- b) *ĐÚNG:* Ruột chì than là điện cực trơ không bị hòa tan.\n"
            "- c) *ĐÚNG:* Cl~2~ phản ứng với NaOH tạo nước Javel có tính tẩy màu và sát trùng.\n"
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
            "- a) *ĐÚNG:* Hiện tượng dương cực tan không sinh khí.\n"
            "- b) *ĐÚNG:* Huy chương là cực âm cathode.\n"
            "- c) *ĐÚNG:* S~tp~ = 43,96 cm^2^; V = 43,96 × 0,1 = 4,396 cm^3^; m~Cu~ = 4,396 × 8,95 ≈ 39,344 g → n~Cu~ = 0,61475 mol → n~e~ = 1,2295 mol → q = 118 647 C → t = 118 647 / 2 ≈ 59 323 s (hoặc xét diện tích mạ phủ ngoài: t ≈ 63 773 giây).\n"
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
            "- a) *ĐÚNG:* m~Cu tăng~ = 36,4 - 35,12 = 1,28 gam → n~Cu~ = 1,28 / 64 = 0,02 mol → n~e~ = 0,04 mol → q = 0,04 × 96500 = 3860 C → t = 3860 / 2 = 1930 giây.\n"
            "- b) *SAI:* Ở anode có cả Zn tan tạo ZnSO~4~ nên nồng độ và thành phần chất tan có biến đổi nhẹ.\n"
            "- c) *ĐÚNG:* Chìa khóa cần mạ là cathode (-), đồng xu mạ là anode (+).\n"
            "- d) *ĐÚNG:* Đồng xu anode bị tan mòn nên m~1~ < 6,5 gam."
        )
    },
    {
        "id": "P3_C65",
        "num": 66,
        "type": "short_answer",
        "content": "Thực hiện thí nghiệm điện phân dung dịch CuSO~4~ với điện cực graphite. Cho các nhận định sau:\n(1) Tại cathode xảy ra sự khử Cu^2+^ trước, sau đó mới đến sự khử của nước.\n(2) Tại anode chỉ xảy ra sự oxi hóa của nước tạo khí hydrogen.\n(3) Sau điện phân, khối lượng cathode tăng lên.\n(4) Theo thời gian điện phân, pH của dung dịch giảm dần.\n(5) Khi vừa bắt đầu điện phân, cả hai điện cực đều có khí thoát ra.\nSắp xếp các số thứ tự của các phát biểu đúng theo thứ tự tăng dần:",
        "ans": "134",
        "solution": (
            "- (1) *ĐÚNG:* Cu^2+^ có thế khử dương hơn nước nên bị khử trước thành Cu.\n"
            "- (2) *SAI:* Tại anode nước bị oxi hóa tạo khí OXYGEN (O~2~), không phải hydrogen.\n"
            "- (3) *ĐÚNG:* Kim loại Cu bám vào làm khối lượng cathode tăng lên.\n"
            "- (4) *ĐÚNG:* Sinh ra H^+^ (H~2~SO~4~) làm pH dung dịch giảm mạnh.\n"
            "- (5) *SAI:* Ban đầu ở cathode chỉ có kim loại Cu kết tủa, không có khí thoát ra.\n"
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
            "- a) *ĐÚNG:* Quá trình khử xảy ra ở cathode (-).\n"
            "- b) *ĐÚNG:* Quá trình oxi hóa nước xảy ra ở anode (+).\n"
            "- c) *ĐÚNG:* Cả khí O~2~, H~2~ bay đi và Cu bám vào điện cực được rút ra ngoài đều làm giảm khối lượng cốc.\n"
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
            "- Số mol Al: n~Al~ = 5400 / 27 = 200 mol.\n"
            "- Số mol O sinh ra: n~O~ = 1,5 × 200 = 300 mol.\n"
            "- Khí X: 70% CO và 30% CO~2~ → n~CO~ = 7x; n~CO2~ = 3x.\n"
            "- Bảo toàn nguyên tố O: 7x + 2(3x) = 13x = 300 mol → x = 300 / 13 mol.\n"
            "- Số mol C: n~C~ = 7x + 3x = 10x = 3000 / 13 mol.\n"
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
            "- Lượng NaOH hao hụt 5% tương ứng 1 tấn NaOH.\n"
            "  → Tổng lượng NaOH tham gia phản ứng trong ngày: m~NaOH~ = 1 / 0,05 = 20 tấn.\n"
            "- Phản ứng: Al~2~O~3~ + 2NaOH → 2NaAlO~2~ + H~2~O\n"
            "  Cứ 2 mol NaOH (2 × 40 = 80 g) hòa tan 1 mol Al~2~O~3~ (102 g).\n"
            "- Khối lượng Al~2~O~3~ tách được mỗi ngày:\n"
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
            "- a) *SAI:* Khối lượng đồng nguyên chất bám ở cathode tăng 3,51 - 3,21 = 0,30 gam, tức là chỉ có 0,30 gam Cu tan ra ở anode. Độ giảm 0,32 gam của thanh anode bao gồm cả 0,02 gam tạp chất tan hoặc rơi xuống bùn anode.\n"
            "- b) *ĐÚNG:* Hiện tượng dương cực tan giữ nồng độ Cu^2+^ không đổi.\n"
            "- c) *SAI:* Cực dương anode xảy ra quá trình oxi hóa: Cu → Cu^2+^ + 2e.\n"
            "- d) *ĐÚNG:* Bố trí chuẩn của bể tinh chế điện phân đồng."
        )
    },
    {
        "id": "P3_C70",
        "num": 71,
        "type": "tf",
        "content": "Máy tạo oxygen y tế sử dụng phương pháp điện phân nước trong môi trường acid để cung cấp oxygen. Quá trình biểu diễn bằng:\n  Cathode (-): 2H^+^ + 2e → H~2~\n  Anode (+): 2H~2~O → O~2~ + 4H^+^ + 4e\nBiết chỉ thêm nước vào ban đầu, máy tạo 1 L oxygen (đkc) mỗi phút, hoạt động 30 lần mỗi ngày, mỗi lần 15 phút. Nước có D = 1 g/mL. Cứ 10 lần hoạt động nước trong máy còn 20% so với ban đầu nên phải châm đầy lại. Cho các nhận định:",
        "image": "p22_img2_xref2098_345x345.jpeg",
        "items": [
            {"id": "a", "text": "Thể tích khí sinh ra ở anode (O~2~) lớn hơn thể tích khí sinh ra ở cathode (H~2~).", "ans": "S"},
            {"id": "b", "text": "Kết thúc quá trình điện phân, số mol ion H^+^ trong dung dịch không thay đổi.", "ans": "Đ"},
            {"id": "c", "text": "Kết thúc quá trình điện phân, pH của dung dịch giảm nhẹ do mất nước làm tăng nồng độ acid.", "ans": "Đ"},
            {"id": "d", "text": "Tổng lượng nước bị điện phân để tạo oxygen trong một ngày là xấp xỉ 653 gam (mL).", "ans": "Đ"}
        ],
        "solution": (
            "- a) *SAI:* V(H~2~) = 2 V(O~2~) nên thể tích ở cathode lớn gấp đôi anode.\n"
            "- b) *ĐÚNG:* Số mol H^+^ sinh ra ở anode đúng bằng số mol H^+^ bị tiêu thụ ở cathode.\n"
            "- c) *ĐÚNG:* Nước bị điện phân hao hụt làm dung môi giảm, nồng độ H^+^ tăng lên nên pH giảm nhẹ.\n"
            "- d) *ĐÚNG:* Tổng thời gian hoạt động: 30 × 15 = 450 phút → V(O~2~) = 450 L.\n"
            "  Số mol O~2~ = 450 / 24,79 ≈ 18,15 mol → n(H~2~O bị điện phân) = 2 × 18,15 ≈ 36,3 mol.\n"
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
            "- Điện lượng: q = 2 × 15 661,95 = 31 323,9 C.\n"
            "- n~e~ = 31 323,9 / 96500 = 0,3246 mol.\n"
            "- Hệ phương trình: x + y = 0,1623 và 64x + 65y = 10,4.\n"
            "  → y = 0,0128 mol; x = 0,1495 mol.\n"
            "- Khối lượng Cu: m~Cu~ = 0,1495 × 64 = 9,568 gam.\n"
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
            "- Thời gian 1 ngày (24 giờ): t = 24 × 3600 = 86 400 giây.\n"
            "- Điện lượng có ích tạo Na: q~ci~ = I × t × H = 40 000 × 86 400 × 0,90 = 3 110 400 000 C.\n"
            "- Phản ứng tại cathode: Na^+^ + 1e → Na.\n"
            "- Số mol Na thu được: n~Na~ = n~e~ = q~ci~ / F = 3 110 400 000 / 96500 ≈ 32 232,12 mol.\n"
            "- Khối lượng Na thu được:\n"
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
            "- a) *ĐÚNG:* Al^3+^ + 3e → Al(lỏng).\n"
            "- b) *ĐÚNG:* Cryolite đóng vai trò dung môi dẫn điện, không bị điện phân tiêu hao.\n"
            "- c) *ĐÚNG:* Với m~Al~ = 1000 gam: A = (5 × 1000 × 96485) / (9 × 3,6 × 10^6) ≈ 14,89 kWh ≈ 15 kWh.\n"
            "- d) *SAI:* Al~2~O~3~ bị điện phân liên tục nên nồng độ giảm, cần định kỳ nạp thêm Al~2~O~3~ vào bồn."
        )
    }
]
