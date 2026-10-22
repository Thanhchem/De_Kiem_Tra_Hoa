# -*- coding: utf-8 -*-
"""
Append Dạng 1 and Dạng 2 to hoa12_to36_data.py
"""
import os
import sys

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

DATA_FILE = r"c:\Antigravity_Thanh\scripts\hoa12_to36_data.py"

dang1_and_dang2 = '''
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
                "🔺 *Tại Anode (cực dương - kim loại mạ):* Kim loại bị oxi hóa tan vào dung dịch:\\n      M → M^n+^ + ne",
                "🔻 *Tại Cathode (cực âm - vật cần mạ):* Ion kim loại trong dung dịch bị khử, bám đều lên bề mặt vật:\\n      M^n+^ + ne → M",
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
        "content": "Để mạ 5,0 g bạc lên một đĩa sắt khi điện phân dung dịch chứa [Ag(NH~3~)~2~]^+^ với dòng điện có cường độ 1,5 A không đổi cần thời gian t phút trong hai trường hợp:\\n  a) Hiệu suất điện phân đạt 100%.\\n  b) Hiệu suất điện phân đạt 80%.\\nCho biết:\\n- Quá trình khử tại cathode: [Ag(NH~3~)~2~]^+^ + 1e → Ag + 2NH~3~.\\n- Điện lượng q = I.t = n~e~.F; F = 96 500 C/mol; Ag = 108.\\nTính giá trị của t trong mỗi trường hợp (làm tròn kết quả đến hàng phần mười).",
        "ans": "a) t = 49,6 phút; b) t = 62,0 phút",
        "solution": (
            "- Số mol Ag cần mạ: n~Ag~ = 5,0 / 108 ≈ 0,04630 mol.\\n"
            "- Theo bán phản ứng tại cathode: n~e~ = n~Ag~ ≈ 0,04630 mol.\\n"
            "- Điện lượng lý thuyết: q = n~e~ × F = (5,0 / 108) × 96500 ≈ 4467,59 C.\\n"
            "a) Với hiệu suất H = 100%:\\n"
            "   t = q / I = 4467,59 / 1,5 ≈ 2978,4 giây = 49,64 phút ≈ *49,6 phút*.\\n"
            "b) Với hiệu suất H = 80%:\\n"
            "   t = 49,64 / 0,80 = 62,05 phút ≈ *62,0 phút*."
        )
    },
    {
        "id": "D1_C2",
        "num": 2,
        "type": "essay",
        "content": "Cần mạ một lớp Ag lên một mặt của một chiếc đĩa tròn có bán kính 12 cm. Với độ dày lớp mạ là 0,01 mm, nếu được cung cấp nguồn điện một chiều có cường độ dòng điện I = 2 A trong thời gian t giờ. Cho biết khối lượng riêng của Ag là D = 10,5 g/cm^3^; Ag = 108; F = 96 500 C/mol; π = 3,1416.\\nTính giá trị của t trong hai trường hợp (làm tròn đến hàng phần trăm):\\n  a) Hiệu suất điện phân là 100%.\\n  b) Hiệu suất điện phân là 90%.",
        "ans": "a) t = 0,59 giờ (35,4 phút); b) t = 0,65 giờ (39,3 phút)",
        "solution": (
            "- Bán kính r = 12 cm; độ dày lớp mạ h = 0,01 mm = 0,001 cm.\\n"
            "- Diện tích mặt đĩa tròn cần mạ: S = π × r^2^ = π × 12^2^ = 144π ≈ 452,39 cm^2^.\\n"
            "- Thể tích lớp bạc: V = S × h = 452,39 × 0,001 ≈ 0,45239 cm^3^.\\n"
            "- Khối lượng bạc cần phủ: m~Ag~ = V × D = 0,45239 × 10,5 ≈ 4,750 g.\\n"
            "- Số mol Ag: n~Ag~ = 4,750 / 108 ≈ 0,04398 mol → n~e~ = n~Ag~ ≈ 0,04398 mol.\\n"
            "- Điện lượng lý thuyết: q = 0,04398 × 96500 ≈ 4244,3 C.\\n"
            "a) Với H = 100%:\\n"
            "   t = q / I = 4244,3 / 2 = 2122,15 giây = 2122,15 / 3600 ≈ *0,59 giờ* (≈ 35,4 phút).\\n"
            "b) Với H = 90%:\\n"
            "   t = 0,5895 / 0,90 ≈ *0,65 giờ* (≈ 39,3 phút)."
        )
    },
    {
        "id": "D1_C3",
        "num": 3,
        "type": "essay",
        "content": "Một vật kim loại được mạ nickel có diện tích bề mặt 100 cm^2^. Dòng điện chạy qua bình điện phân có cường độ 0,5 A và thời gian mạ là 5 giờ. Cho biết nickel có khối lượng mol nguyên tử là 58,7 g/mol, hóa trị 2 trong muối, khối lượng riêng D = 8,8 × 10^3^ kg/m^3^ (tức 8,8 g/cm^3^), hằng số Faraday F = 96 500 C/mol, hiệu suất điện phân 100%.\\nĐộ dày của lớp nickel phủ đều trên mặt của vật được mạ là bao nhiêu µm? (Kết quả làm tròn đến hàng phần mười).",
        "ans": "31,1 µm",
        "solution": (
            "- Thời gian điện phân: t = 5 × 3600 = 18 000 s.\\n"
            "- Số mol electron trao đổi: n~e~ = (I × t) / F = (0,5 × 18 000) / 96500 ≈ 0,093264 mol.\\n"
            "- Phản ứng tại cathode: Ni^2+^ + 2e → Ni → n~Ni~ = n~e~ / 2 ≈ 0,046632 mol.\\n"
            "- Khối lượng Ni kết tủa: m~Ni~ = 0,046632 × 58,7 ≈ 2,7373 g.\\n"
            "- Thể tích lớp nickel: V = m / D = 2,7373 / 8,8 ≈ 0,31106 cm^3^.\\n"
            "- Bề dày lớp mạ: h = V / S = 0,31106 / 100 = 0,0031106 cm = 31,106 µm ≈ *31,1 µm*."
        )
    },
    {
        "id": "D1_C4",
        "num": 4,
        "type": "essay",
        "content": "Một tấm huy chương mặt tròn bằng kim loại có đường kính 70 mm, bề dày 3 mm; được mạ thêm một lớp vàng dày 0,05 mm bám đều lên tất cả các mặt của tấm huy chương. Biết khối lượng riêng của Au là 19,3 g/cm^3^, khối lượng mol của Au là 197 g/mol; khi điện phân mỗi ion vàng trao đổi 3 electron và hằng số Faraday F = 96 500 C/mol; hiệu suất 100%.\\nThời gian điện phân với cường độ dòng điện không đổi I = 10 A để mạ lớp vàng lên tấm huy chương là bao nhiêu giây? (Kết quả làm tròn đến hàng đơn vị).",
        "ans": "1185 giây",
        "solution": (
            "- Quy đổi kích thước huy chương: bán kính r = 35 mm = 3,5 cm; chiều dày h~hc~ = 3 mm = 0,3 cm.\\n"
            "- Tổng diện tích bề mặt (2 mặt tròn + mặt xung quanh):\\n"
            "  S~tp~ = 2πr^2^ + 2πr × h~hc~ = 2π × (3,5)^2^ + 2π × 3,5 × 0,3 = 24,5π + 2,1π = 26,6π ≈ 83,566 cm^2^.\\n"
            "- Bề dày lớp mạ vàng: h = 0,05 mm = 0,005 cm.\\n"
            "- Thể tích lớp vàng: V~Au~ = S~tp~ × h = 83,566 × 0,005 ≈ 0,41783 cm^3^.\\n"
            "- Khối lượng vàng cần mạ: m~Au~ = V~Au~ × D = 0,41783 × 19,3 ≈ 8,064 g.\\n"
            "- Số mol Au: n~Au~ = 8,064 / 197 ≈ 0,040935 mol.\\n"
            "- Số mol electron: n~e~ = 3 × n~Au~ ≈ 0,122806 mol.\\n"
            "- Điện lượng cần thiết: q = n~e~ × F = 0,122806 × 96500 ≈ 11850,8 C.\\n"
            "- Thời gian điện phân: t = q / I = 11850,8 / 10 ≈ *1185 giây*."
        )
    },
    {
        "id": "D1_C5",
        "num": 5,
        "type": "essay",
        "content": "Để thử nghiệm kỹ thuật mạ điện người ta phủ một lớp Zn lên một mẫu thử hình lập phương có chiều dài mỗi cạnh là 12 cm. Nhúng mẫu thử vào bình điện phân chứa dung dịch ZnSO~4~, cathode là mẫu thử, anode bằng Zn. Thực hiện quá trình mạ điện với dòng điện một chiều I = 15 A trong thời gian 34,44 giờ. Biết hiệu suất điện phân đạt 90%; khối lượng riêng của Zn là 7,13 g/cm^3^; Zn = 65,38; F = 96 500 C/mol.\\nBề dày lớp kẽm phủ đều lên mẫu thử là bao nhiêu mm? (Kết quả làm tròn đến hàng phần mười).",
        "ans": "0,9 mm",
        "solution": (
            "- Diện tích toàn phần khối lập phương (6 mặt): S~tp~ = 6 × a^2^ = 6 × 12^2^ = 864 cm^2^.\\n"
            "- Thời gian điện phân: t = 34,44 × 3600 = 123 984 s.\\n"
            "- Điện lượng thực tế có ích: q~ci~ = I × t × H = 15 × 123 984 × 0,90 = 1 673 784 C.\\n"
            "- Số mol e có ích: n~e~ = 1 673 784 / 96500 ≈ 17,3449 mol.\\n"
            "- Phản ứng tại cathode: Zn^2+^ + 2e → Zn → n~Zn~ = n~e~ / 2 ≈ 8,67245 mol.\\n"
            "- Khối lượng Zn tạo thành: m~Zn~ = 8,67245 × 65,38 ≈ 566,99 g (hoặc dùng 65: ≈ 563,7 g).\\n"
            "- Thể tích Zn: V = m / D = 566,99 / 7,13 ≈ 79,52 cm^3^.\\n"
            "- Bề dày lớp kẽm: h = V / S~tp~ = 79,52 / 864 ≈ 0,0920 cm = 0,92 mm ≈ *0,9 mm*."
        )
    },
    {
        "id": "D1_C6",
        "num": 6,
        "type": "essay",
        "content": "Người ta mạ kẽm lên bề mặt quả cân hình cầu bằng phương pháp mạ điện. Dung dịch điện phân chứa ZnSO~4~, cực dương là Zn kim loại, cực âm là quả cân (bán kính R = 5 cm). Sự điện phân với cường độ dòng điện không đổi I = 16 A. Quả cân cần được phủ đều một lớp kẽm dày 1 mm trên bề mặt. Biết hiệu suất điện phân đạt 100%; khối lượng riêng của Zn là 7,13 g/cm^3^; Zn = 65,38; π = 3,1416; F = 96 500 C/mol.\\nThời gian của quá trình mạ điện là bao nhiêu giờ? (Kết quả làm tròn đến hàng phần mười).",
        "ans": "11,5 giờ",
        "solution": (
            "- Diện tích mặt cầu: S = 4πR^2^ = 4 × 3,1416 × 5^2^ = 314,16 cm^2^.\\n"
            "- Bề dày lớp kẽm: h = 1 mm = 0,1 cm.\\n"
            "- Thể tích lớp kẽm: V = S × h = 314,16 × 0,1 = 31,416 cm^3^.\\n"
            "- Khối lượng kẽm bám lên quả cân: m~Zn~ = V × D = 31,416 × 7,13 ≈ 224,0 g.\\n"
            "- Số mol Zn: n~Zn~ = 224,0 / 65,38 ≈ 3,426 mol → n~e~ = 2 × n~Zn~ ≈ 6,852 mol.\\n"
            "- Điện lượng q = n~e~ × F = 6,852 × 96500 ≈ 661 228 C.\\n"
            "- Thời gian điện phân: t = q / I = 661 228 / 16 ≈ 41 327 giây ≈ 11,48 giờ ≈ *11,5 giờ*."
        )
    },
    {
        "id": "D1_C7",
        "num": 7,
        "type": "essay",
        "content": "Để mạ lớp nickel (Ni) lên một chi tiết máy bằng thép hình trụ đường kính 7 cm, chiều dài 15 cm, người ta sử dụng chi tiết máy cần mạ đó làm cathode và nhúng vào trong bể mạ chứa dung dịch NiSO~4~, cường độ dòng điện đi qua bể mạ là 8,0 A (hiệu suất điện phân là 100%). Nếu độ dày lớp mạ nickel lên toàn bộ bề mặt chi tiết máy là 20 µm (1 µm = 10^-4^ cm) thì thời gian cần để mạ chi tiết trên là t phút. Cho biết khối lượng riêng của Ni bằng 8,91 g/cm^3^; Ni = 58,7; hằng số Faraday F = 96 500 C/mol.\\nTính giá trị của t (kết quả làm tròn đến hàng phần mười).",
        "ans": "49,7 phút (hoặc 40,3 phút nếu chỉ mạ mặt xung quanh)",
        "solution": (
            "- Bán kính r = 3,5 cm; chiều dài L = 15 cm; bề dày h = 20 µm = 0,002 cm.\\n"
            "- Trường hợp mạ toàn phần chi tiết máy (cả 2 đáy và mặt xung quanh):\\n"
            "  S~tp~ = 2πrL + 2πr^2^ = 2π × 3,5 × 15 + 2π × (3,5)^2^ = 105π + 24,5π = 129,5π ≈ 406,84 cm^2^.\\n"
            "  Thể tích Ni: V = 406,84 × 0,002 ≈ 0,8137 cm^3^.\\n"
            "  Khối lượng Ni: m~Ni~ = 0,8137 × 8,91 ≈ 7,250 g.\\n"
            "  Số mol Ni: n~Ni~ = 7,250 / 58,7 ≈ 0,1235 mol → n~e~ = 2 × 0,1235 = 0,2470 mol.\\n"
            "  Điện lượng: q = 0,2470 × 96500 ≈ 23838 C.\\n"
            "  Thời gian t = q / I = 23838 / 8 = 2980 giây ≈ *49,7 phút*.\\n"
            "(Nếu chi tiết máy chỉ mạ mặt xung quanh: S = 105π ≈ 329,87 cm^2^ → t ≈ 40,3 phút)."
        )
    },
    {
        "id": "D1_C8",
        "num": 8,
        "type": "essay",
        "content": "Để ngăn chặn hiện tượng gỉ sét và mài mòn của một chi tiết máy, người ta phủ lớp đồng (Cu) dày 0,5 mm lên toàn bộ bề mặt của chi tiết máy bằng phương pháp mạ điện. Dung dịch điện phân chứa CuSO~4~, cực dương là Cu kim loại, cực âm là chi tiết máy có hình trụ (bán kính 4,25 cm, chiều cao 25 cm). Sự điện phân với dòng điện một chiều I = 15 A. Biết hiệu suất điện phân đạt 100%; khối lượng riêng của Cu là 8,96 g/cm^3^; Cu = 64; π = 3,14; F = 96 500 C/mol.\\nThời gian của quá trình mạ điện là bao nhiêu giờ? (Kết quả làm tròn đến hàng phần mười).",
        "ans": "19,5 giờ (hoặc 16,7 giờ nếu chỉ tính mặt xung quanh)",
        "solution": (
            "- Chi tiết hình trụ có r = 4,25 cm; h~trụ~ = 25 cm; bề dày mạ h = 0,5 mm = 0,05 cm.\\n"
            "- Diện tích toàn phần chi tiết máy:\\n"
            "  S~tp~ = 2πr × h~trụ~ + 2πr^2^ = 2 × 3,14 × 4,25 × 25 + 2 × 3,14 × (4,25)^2^ = 667,25 + 113,43 = 780,68 cm^2^.\\n"
            "- Thể tích Cu cần mạ: V = S~tp~ × h = 780,68 × 0,05 = 39,034 cm^3^.\\n"
            "- Khối lượng Cu: m~Cu~ = 39,034 × 8,96 ≈ 349,75 g.\\n"
            "- Số mol Cu: n~Cu~ = 349,75 / 64 ≈ 5,465 mol → n~e~ = 2 × 5,465 = 10,93 mol.\\n"
            "- Điện lượng: q = 10,93 × 96500 ≈ 1 054 745 C.\\n"
            "- Thời gian điện phân: t = q / I = 1 054 745 / 15 ≈ 70 316 giây ≈ 19,53 giờ ≈ *19,5 giờ*.\\n"
            "(Nếu chỉ mạ mặt xung quanh: S = 667,25 cm^2^ → m~Cu~ = 298,93 g → t ≈ 16,7 giờ)."
        )
    },
    {
        "id": "D1_C9",
        "num": 9,
        "type": "essay",
        "content": "Một tấm huy chương mặt tròn bằng kim loại có bán kính 4 cm, chiều dày 0,5 cm, cần được mạ kín bởi một lớp vàng dày 0,1 mm. Biết khối lượng riêng của Au là 19,36 g/cm^3^; nguyên tử khối của Au là 197 g/mol; π = 3,14; khi điện phân mỗi ion vàng trao đổi 3 electron; F = 96 500 C/mol.\\nThời gian điện phân với cường độ dòng điện không đổi 12 A (hiệu suất H = 100%) để mạ lớp vàng lên tấm huy chương đó là bao nhiêu giờ? (Kết quả làm tròn đến hàng phần mười).",
        "ans": "0,7 giờ (hoặc 44,7 phút)",
        "solution": (
            "- Huy chương tròn có r = 4 cm; chiều dày h~hc~ = 0,5 cm; bề dày mạ h = 0,1 mm = 0,01 cm.\\n"
            "- Diện tích toàn phần (2 đáy tròn và mặt bên xung quanh):\\n"
            "  S~tp~ = 2πr^2^ + 2πr × h~hc~ = 2 × 3,14 × 4^2^ + 2 × 3,14 × 4 × 0,5 = 100,48 + 12,56 = 113,04 cm^2^.\\n"
            "- Thể tích lớp vàng: V = S~tp~ × h = 113,04 × 0,01 = 1,1304 cm^3^.\\n"
            "- Khối lượng vàng cần mạ: m~Au~ = 1,1304 × 19,36 ≈ 21,885 g.\\n"
            "- Số mol Au: n~Au~ = 21,885 / 197 ≈ 0,11109 mol → n~e~ = 3 × 0,11109 ≈ 0,33327 mol.\\n"
            "- Điện lượng: q = n~e~ × F = 0,33327 × 96500 ≈ 32 160,5 C.\\n"
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
        "content": "Điện phân dung dịch Na~2~SO~4~ 0,5M với anode làm bằng Al và cathode trơ bằng dòng điện có cường độ 2,68 A. Thời gian để tạo thành lớp màng Al~2~O~3~ dày 3 µm trên bề mặt tấm nhôm có diện tích bề mặt là 1 m^2^ là bao lâu? Cho biết khối lượng riêng của Al~2~O~3~ là 3,95 g/cm^3^, hiệu suất điện phân là 95%; Al = 27, O = 16; F = 96 500 C/mol.\\nTính thời gian theo giờ (kết quả làm tròn đến hàng phần trăm).",
        "ans": "7,34 giờ",
        "solution": (
            "- Diện tích tấm nhôm: S = 1 m^2^ = 10 000 cm^2^.\\n"
            "- Độ dày lớp màng: h = 3 µm = 3 × 10^-4^ cm.\\n"
            "- Thể tích Al~2~O~3~: V = S × h = 10 000 × (3 × 10^-4^) = 3 cm^3^.\\n"
            "- Khối lượng Al~2~O~3~ tạo thành: m = V × D = 3 × 3,95 = 11,85 g.\\n"
            "- Số mol Al~2~O~3~: n = 11,85 / 102 ≈ 0,116176 mol.\\n"
            "- Theo bán phản ứng tại anode: 2Al + 3H~2~O → Al~2~O~3~ + 6H^+^ + 6e\\n"
            "  → Cứ tạo ra 1 mol Al~2~O~3~ cần trao đổi 6 mol electron: n~e~ = 6 × n = 6 × 0,116176 ≈ 0,69706 mol.\\n"
            "- Điện lượng lý thuyết: q~lt~ = n~e~ × F = 0,69706 × 96500 ≈ 67 266 C.\\n"
            "- Với hiệu suất H = 95%:\\n"
            "  Điện lượng thực tế: q~tt~ = q~lt~ / 0,95 = 67 266 / 0,95 ≈ 70 806 C.\\n"
            "- Thời gian điện phân:\\n"
            "  t = q~tt~ / I = 70 806 / 2,68 ≈ 26 420 giây = 26 420 / 3600 ≈ *7,34 giờ*."
        )
    },
    {
        "id": "D2_C2",
        "num": 2,
        "type": "essay",
        "content": "Điện phân dung dịch Na~2~SO~4~ 0,5M với anode làm bằng Al và cathode trơ bằng dòng điện có cường độ 2,68 A. Thời gian để tạo thành lớp màng Al~2~O~3~ dày 3 µm trên bề mặt tấm nhôm có diện tích bề mặt là 1 m^2^ là bao nhiêu giờ? Cho biết khối lượng riêng của Al~2~O~3~ là 3,95 g/cm^3^, hiệu suất điện phân là 95%; F = 96 500 C/mol.\\nBiết: q = n.F = I.t (n là số mol electron trao đổi). Kết quả làm tròn đến hàng phần trăm sau dấu phẩy.",
        "ans": "7,34 giờ",
        "solution": (
            "- Diện tích S = 1 m^2^ = 10 000 cm^2^; chiều dày h = 3 µm = 3 × 10^-4^ cm.\\n"
            "- V~oxit~ = 10 000 × 3 × 10^-4^ = 3 cm^3^ → m~oxit~ = 3 × 3,95 = 11,85 g.\\n"
            "- n~Al2O3~ = 11,85 / 102 ≈ 0,11618 mol → n~e~ = 6 × n~Al2O3~ ≈ 0,69706 mol.\\n"
            "- q~cần~ = (0,69706 × 96500) / 0,95 ≈ 70 806 C.\\n"
            "- Thời gian điện phân: t = 70 806 / 2,68 ≈ 26 420 giây ≈ *7,34 giờ*."
        )
    },
    {
        "id": "D2_C3",
        "num": 3,
        "type": "essay",
        "content": "Sự tạo thành các lớp màng Al~2~O~3~ trên bề mặt Al có tác dụng bảo vệ nhôm. Bên cạnh quá trình tạo thành màng Al~2~O~3~ tự nhiên như để nhôm trong không khí, người ta có thể tạo thành lớp màng Al~2~O~3~ như ý bằng cách điện phân dung dịch H~2~SO~4~ 0,5M với anode làm bằng Al và cathode trơ bằng dòng điện có cường độ 2,68 A. Tính thời gian để tạo thành lớp màng Al~2~O~3~ dày 2 µm trên bề mặt tấm nhôm có diện tích bề mặt là 2 m^2^. Cho biết khối lượng riêng của Al~2~O~3~ là 3,95 g/cm^3^; hiệu suất điện phân là 90%; Al = 27, O = 16, F = 96 500 C/mol.\\nTính thời gian theo giờ (kết quả làm tròn đến hàng phần trăm).",
        "ans": "10,33 giờ",
        "solution": (
            "- Diện tích tấm nhôm: S = 2 m^2^ = 20 000 cm^2^.\\n"
            "- Chiều dày màng oxide: h = 2 µm = 2 × 10^-4^ cm.\\n"
            "- Thể tích Al~2~O~3~: V = S × h = 20 000 × (2 × 10^-4^) = 4 cm^3^.\\n"
            "- Khối lượng Al~2~O~3~: m = V × D = 4 × 3,95 = 15,8 g.\\n"
            "- Số mol Al~2~O~3~: n = 15,8 / 102 ≈ 0,1549 mol.\\n"
            "- Số mol electron trao đổi: n~e~ = 6 × n = 6 × 0,1549 ≈ 0,9294 mol.\\n"
            "- Điện lượng lý thuyết: q~lt~ = 0,9294 × 96500 ≈ 89 688 C.\\n"
            "- Với hiệu suất H = 90%:\\n"
            "  Điện lượng thực tế: q~tt~ = 89 688 / 0,90 ≈ 99 654 C.\\n"
            "- Thời gian điện phân: t = q~tt~ / I = 99 654 / 2,68 ≈ 37 184 giây = 37 184 / 3600 ≈ *10,33 giờ*."
        )
    }
]
'''

with open(DATA_FILE, "a", encoding="utf-8") as f:
    f.write(dang1_and_dang2)

print("Appended Dạng 1 and Dạng 2.")
