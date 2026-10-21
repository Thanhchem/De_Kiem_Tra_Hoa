# -*- coding: utf-8 -*-
"""
Dữ liệu chuẩn cho bộ 3 tài liệu Ôn tập KHTN 9 Giữa HK1 (2026 - 2027)
Trường THCS TRẦN CAO VÂN - TP Huế
Thầy TRẦN VĂN THẠNH - 0777.470.803
Bao gồm 3 phân môn:
1. Phân môn: NĂNG LƯỢNG (Vật lí)
2. Phân môn: CHẤT VÀ SỰ BIẾN ĐỔI CHẤT (Hóa học)
3. Phân môn: VẬT SỐNG (Sinh học)
"""

import os

ROOT_DIR = r"c:\Antigravity_Thanh"
OUT_DIR = os.path.join(ROOT_DIR, "San_Pham")
IMG_DIR = os.path.join(OUT_DIR, "images_khtn9_trancao_van")
IMG_CON_LAC = os.path.join(IMG_DIR, "con_lac_hd.png")

INFO = {
    "so_gddt": "UBND THÀNH PHỐ HUẾ",
    "truong": "TRƯỜNG THCS TRẦN CAO VÂN",
    "tieude": "HƯỚNG DẪN ÔN TẬP GIỮA HỌC KÌ I",
    "mon": "MÔN: KHOA HỌC TỰ NHIÊN - LỚP 9",
    "nam_hoc": "NĂM HỌC 2026 - 2027",
    "thoigian": "Thời gian làm bài: 60 phút",
    "nguyen_tu_khoi": "(Cho biết nguyên tử khối: H = 1; C = 12; O = 16; Al = 27; S = 32; K = 39; Cl = 35,5; Fe = 56; Cu = 64; Ba = 137; Ag = 108; P = 31; Zn = 65; Mg = 24; Na = 23; Ca = 40; lấy gia tốc trọng trường g = 10 m/s^2^ nếu không có ghi chú khác)"
}

# =============================================================================
# PHÂN MÔN 1: NĂNG LƯỢNG (VẬT LÍ)
# =============================================================================
VATLI_QUESTIONS = [
    {
        "id": 1,
        "type": "essay",
        "question": "Một tòa nhà cao 10 tầng, mỗi tầng cao 3,4 m, có một thang chở tối đa được 20 người, mỗi người có khối lượng trung bình 50 kg. Mỗi chuyến lên tầng 10, nếu không dừng ở các tầng khác, mất một phút. Công suất tối thiểu của động cơ thang máy phải là bao nhiêu?",
        "solution": (
            "• Tổng khối lượng tối đa của 20 người: m = 20 . 50 = 1000 kg.\n"
            "• Trọng lượng của 20 người: P = 10 . m = 10 . 1000 = 10 000 N.\n"
            "• Độ cao thang máy cần nâng lên: Thang máy đi từ tầng 1 lên tầng 10 vượt qua 9 khoảng tầng, do đó độ cao nâng là: h = (10 - 1) . 3,4 = 30,6 m.\n"
            "• Thời gian chuyển động: t = 1 phút = 60 s.\n"
            "• Công tối thiểu cần thực hiện: A = P . h = 10 000 . 30,6 = 306 000 J.\n"
            "• Công suất tối thiểu của động cơ thang máy:\n"
            "  *P* = A / t = 306 000 / 60 = *5 100 W = 5,1 kW*.\n"
            "*(Ghi chú: Nếu tính độ cao của toàn bộ 10 tầng h = 34 m thì A = 340 000 J => P ≈ 5 666,7 W ≈ 5,67 kW)*."
        )
    },
    {
        "id": 2,
        "type": "essay",
        "question": "Một máy nâng cần đưa 100 viên gạch, mỗi viên nặng 1,5 kg lên cao 5 m. Tổng công mà máy cần thực hiện là bao nhiêu?",
        "solution": (
            "• Tổng khối lượng của 100 viên gạch: m = 100 . 1,5 = 150 kg.\n"
            "• Trọng lượng của 100 viên gạch: P = 10 . m = 10 . 150 = 1 500 N.\n"
            "• Tổng công tối thiểu máy cần thực hiện để đưa gạch lên cao h = 5 m là:\n"
            "  *A = P . h = 1 500 . 5 = 7 500 J = 7,5 kJ*.\n"
            "*(Lưu ý: Đề bài gốc gõ nhầm đơn vị công thành \"W\", đơn vị chuẩn của công cơ học là Jun (J) hoặc kilôjun (kJ))*"
        )
    },
    {
        "id": 3,
        "type": "tf",
        "intro": "Một học sinh thả một vật có khối lượng 0,1 kg ở độ cao 4,5 m rơi thẳng tự do xuống mặt đất. Bỏ qua mọi lực cản (lấy g = 10 m/s^2^). Các nhận định sau là đúng hay sai?",
        "items": [
            {
                "label": "a",
                "text": "Tốc độ của vật khi vật vừa gần chạm đất là 4 m/s.",
                "ans": "Sai",
                "explain": "Cơ năng bảo toàn W = W~t0~ = m.g.h~0~ = 0,1 . 10 . 4,5 = 4,5 J. Khi vật vừa chạm đất (h = 0), thế năng bằng 0 nên động năng cực đại: W~đ~ = W => 1/2.m.v^2^ = 4,5 => v = √(2.W/m) = √(2 . 10 . 4,5) = √90 ≈ 9,49 m/s ≠ 4 m/s. Do đó ý a Sai."
            },
            {
                "label": "b",
                "text": "Khi vật ở độ cao 1,5 m thì động năng gấp 2 lần thế năng.",
                "ans": "Đúng",
                "explain": "Tại độ cao h = 1,5 m: Thế năng W~t~ = m.g.h = 0,1 . 10 . 1,5 = 1,5 J. Động năng W~đ~ = W - W~t~ = 4,5 - 1,5 = 3,0 J. Ta có W~đ~ / W~t~ = 3,0 / 1,5 = 2 => W~đ~ = 2.W~t~. Do đó ý b Đúng."
            }
        ]
    },
    {
        "id": 4,
        "type": "essay",
        "question": "Một người kéo đều một vật từ giếng sâu 12 m trong nửa phút. Người ấy phải dùng một lực F = 150 N. Công và công suất của người kéo là bao nhiêu?",
        "solution": (
            "• Quãng đường vật di chuyển: s = 12 m.\n"
            "• Thời gian thực hiện: t = 0,5 phút = 30 s.\n"
            "• Công của người kéo thực hiện: *A = F . s = 150 . 12 = 1 800 J*.\n"
            "• Công suất của người kéo: *P = A / t = 1 800 / 30 = 60 W*."
        )
    },
    {
        "id": 5,
        "type": "mcq",
        "question": "Con ngựa kéo xe chuyển động đều với tốc độ 9 km/h. Lực kéo của ngựa là 200 N. Công suất của ngựa là bao nhiêu?",
        "options": [
            "A. 6480 J",
            "B. 6550 J",
            "C. 6480 kJ",
            "D. 6550 kJ"
        ],
        "correct": "A",
        "solution": (
            "• Đổi vận tốc sang đơn vị m/s: v = 9 km/h = 9 / 3,6 = 2,5 m/s.\n"
            "• Theo đúng định nghĩa công suất cơ học: *P = F . v = 200 . 2,5 = 500 W*.\n"
            "• *Phân tích về đề bài:* Các phương án đề cho mang đơn vị của công (J, kJ). Nếu tính công trong 1 giờ (3600 s): A = 500 . 3600 = 1 800 000 J = 1 800 kJ. Ở câu hỏi này, nếu theo đáp án trắc nghiệm đề gốc trường đưa ra thường đánh dấu chọn A; tuy nhiên về bản chất vật lí chính xác thì công suất của ngựa là *500 W*."
        )
    },
    {
        "id": 6,
        "type": "tf",
        "intro": "Hai vật ở cùng một độ cao so với gốc thế năng, biết khối lượng của vật thứ nhất gấp 3 lần khối lượng của vật thứ hai (m~1~ = 3m~2~). Các nhận định sau là đúng hay sai?",
        "items": [
            {
                "label": "a",
                "text": "Thế năng trọng trường của vật thứ nhất gấp 3 lần thế năng trọng trường của vật thứ hai.",
                "ans": "Đúng",
                "explain": "Thế năng trọng trường W~t~ = m.g.h. Vì hai vật ở cùng độ cao h và cùng gia tốc trọng trường g nên thế năng tỉ lệ thuận với khối lượng: m~1~ = 3m~2~ => W~t1~ = 3W~t2~."
            },
            {
                "label": "b",
                "text": "Thế năng trọng trường của vật thứ hai gấp 3 lần thế năng trọng trường của vật thứ nhất.",
                "ans": "Sai",
                "explain": "Vật thứ nhất có khối lượng lớn hơn nên thế năng của vật thứ nhất lớn hơn, W~t2~ = 1/3 W~t1~."
            },
            {
                "label": "c",
                "text": "Cơ năng của vật thứ nhất gấp 3 lần cơ năng của vật thứ hai.",
                "ans": "Đúng",
                "explain": "Khi hai vật đứng yên ở cùng độ cao (v = 0), cơ năng bằng thế năng (W = W~t~). Do đó W~1~ = 3W~2~."
            },
            {
                "label": "d",
                "text": "Cơ năng của vật thứ hai gấp 3 lần cơ năng của vật thứ nhất.",
                "ans": "Sai",
                "explain": "Vật thứ hai có khối lượng nhỏ hơn nên cơ năng phải nhỏ hơn (bằng 1/3 cơ năng vật thứ nhất)."
            }
        ]
    },
    {
        "id": 7,
        "type": "essay",
        "question": "Một công nhân vác một bao xi măng có trọng lượng 500 N trên vai, đứng trên sân thượng toà nhà cao 20 m so với mặt đất. Độ cao của bao xi măng so với mặt sân thượng là 1,4 m. Tính thế năng trọng trường của bao xi măng trong hai trường hợp sau:\na. Chọn gốc thế năng tại mặt sân thượng toà nhà.\nb. Chọn gốc thế năng tại mặt đất.",
        "solution": (
            "a. Gốc thế năng tại mặt sân thượng toà nhà:\n"
            "   • Độ cao của bao xi măng so với mốc sân thượng: h~1~ = 1,4 m.\n"
            "   • Thế năng trọng trường: *W~t1~ = P . h~1~ = 500 . 1,4 = 700 J*.\n"
            "b. Gốc thế năng tại mặt đất:\n"
            "   • Độ cao của bao xi măng so với mặt đất: h~2~ = 20 + 1,4 = 21,4 m.\n"
            "   • Thế năng trọng trường: *W~t2~ = P . h~2~ = 500 . 21,4 = 10 700 J*."
        )
    },
    {
        "id": 8,
        "type": "pendulum_essay",
        "question": "Một con lắc gồm vật nặng có khối lượng 2 kg được treo vào đầu sợi dây nhẹ, không dãn. Từ vị trí cân bằng O ban đầu, vật được nâng lên 0,5 m đến điểm A rồi thả nhẹ (hình vẽ bên). Chọn gốc thế năng tại O. Coi cơ năng của vật không đổi (lấy g = 10 m/s^2^).\na) Tính cơ năng của vật tại A.\nb) Khi vật chuyển động từ A đến O, động năng và thế năng của nó thay đổi như thế nào? Vì sao?\nc) Tính tốc độ của vật khi đi qua điểm O.",
        "solution": (
            "a) Tại điểm A: Thả nhẹ nên vận tốc ban đầu v~A~ = 0 => động năng W~đA~ = 0.\n"
            "   • Độ cao của vật tại A so với mốc O là h~A~ = 0,5 m.\n"
            "   • Thế năng tại A: W~tA~ = m . g . h~A~ = 2 . 10 . 0,5 = 10 J.\n"
            "   • Cơ năng của vật tại A: *W~A~ = W~đA~ + W~tA~ = 0 + 10 = 10 J*.\n"
            "   *(Nếu lấy g = 9,8 m/s^2^ thì W~A~ = 2 . 9,8 . 0,5 = 9,8 J)*.\n"
            "b) Khi vật chuyển động từ A đến O:\n"
            "   • Độ cao của vật giảm dần từ 0,5 m về 0 m nên *thế năng của vật giảm dần* (từ 10 J về 0 J).\n"
            "   • Tốc độ của vật tăng dần nên *động năng của vật tăng dần* (từ 0 J lên 10 J).\n"
            "   • *Vì sao:* Bỏ qua ma sát và sức cản của không khí, cơ năng của hệ được bảo toàn; thế năng đã chuyển hóa hoàn toàn thành động năng của con lắc.\n"
            "c) Khi vật đi qua vị trí cân bằng O:\n"
            "   • Tại O có h~O~ = 0 => W~tO~ = 0; toàn bộ cơ năng chuyển thành động năng: W~O~ = W~đO~ = 1/2 . m . v~O~^2^.\n"
            "   • Theo định luật bảo toàn cơ năng: W~O~ = W~A~ => 1/2 . 2 . v~O~^2^ = 10 => v~O~^2^ = 10.\n"
            "   • Tốc độ của vật khi qua O: *v~O~ = √10 ≈ 3,16 m/s* (nếu dùng g = 9,8 thì v~O~ = √9,8 ≈ 3,13 m/s)."
        )
    }
]

# =============================================================================
# PHÂN MÔN 2: CHẤT VÀ SỰ BIẾN ĐỔI CHẤT (HÓA HỌC)
# =============================================================================
HOA_THEORY = [
    {
        "title": "1. Tính chất hóa học của kim loại và phương trình minh họa",
        "content": (
            "• *Tác dụng với phi kim:*\n"
            "  + Với oxygen: Hầu hết kim loại (trừ Au, Pt, Ag...) phản ứng với oxygen tạo basic oxide:\n"
            "    3Fe + 2O~2~ -(t°)-> Fe~3~O~4~  hoặc  2Mg + O~2~ -(t°)-> 2MgO\n"
            "  + Với phi kim khác (Cl~2~, S...): Tạo muối halogenua, sulfide:\n"
            "    2Fe + 3Cl~2~ -(t°)-> 2FeCl~3~ ;  Fe + S -(t°)-> FeS\n"
            "• *Tác dụng với dung dịch acid (HCl, H~2~SO~4~ loãng):*\n"
            "  Kim loại đứng trước Hydrogen trong dãy HĐHH tác dụng giải phóng khí H~2~:\n"
            "    Fe + 2HCl -> FeCl~2~ + H~2~↑  ;  2Al + 3H~2~SO~4~ -> Al~2~(SO~4~)~3~ + 3H~2~↑\n"
            "• *Tác dụng với dung dịch muối:*\n"
            "  Kim loại hoạt động mạnh hơn (trừ Na, K, Ba, Ca...) đẩy kim loại yếu hơn ra khỏi dung dịch muối:\n"
            "    Fe + CuSO~4~ -> FeSO~4~ + Cu↓\n"
            "• *Tác dụng với nước:*\n"
            "  Các kim loại nhóm IA và IIA (K, Na, Ca, Ba) phản ứng mãnh liệt với nước ở nhiệt độ thường:\n"
            "    2Na + 2H~2~O -> 2NaOH + H~2~↑"
        )
    },
    {
        "title": "2. Các phương pháp tách kim loại (Điện phân nóng chảy, Nhiệt luyện)",
        "content": (
            "• *Phương pháp điện phân nóng chảy (ĐPNC):*\n"
            "  Áp dụng để điều chế các kim loại hoạt động hóa học mạnh (K, Na, Ca, Mg, Al):\n"
            "  + Tách Sodium (Na) ra khỏi muối ăn (chuyển Na~2~CO~3~ + 2HCl -> 2NaCl + CO~2~↑ + H~2~O rồi đem ĐPNC):\n"
            "    2NaCl -(đpnc)-> 2Na + Cl~2~↑\n"
            "  + Tách Aluminium (Al) ra khỏi quặng bauxite (Al~2~O~3~) với chất xúc tác cryolite:\n"
            "    2Al~2~O~3~ -(đpnc / cryolite)-> 4Al + 3O~2~↑\n"
            "• *Phương pháp nhiệt luyện:*\n"
            "  Khử ion kim loại trong oxide ở nhiệt độ cao bằng các chất khử mạnh như C, CO, H~2~, Al. Dùng cho kim loại có độ hoạt động trung bình và yếu (sau Al như Zn, Fe, Pb, Cu):\n"
            "  + Tách Copper (Cu) ra khỏi CuSO~4~: Cho phản ứng với kiềm tạo Cu(OH)~2~, nung thành CuO rồi dùng CO hoặc H~2~ khử:\n"
            "    CuSO~4~ + 2NaOH -> Cu(OH)~2~↓ + Na~2~SO~4~\n"
            "    Cu(OH)~2~ -(t°)-> CuO + H~2~O\n"
            "    CuO + CO -(t°)-> Cu + CO~2~↑  (hoặc dùng sắt đẩy trực tiếp theo phương pháp thủy luyện: Fe + CuSO~4~ -> FeSO~4~ + Cu↓)."
        )
    },
    {
        "title": "3. Hợp kim: Khái niệm, tính chất và ứng dụng",
        "content": (
            "• *Khái niệm:* Hợp kim là vật liệu kim loại chứa một kim loại cơ bản và một số kim loại hoặc phi kim khác.\n"
            "• *Tính chất chung:* Hợp kim thường có độ cứng cao hơn, độ bền cơ học tốt hơn, khả năng chống ăn mòn cao hơn so với các kim loại nguyên chất cấu thành.\n"
            "• *Các hợp kim tiêu biểu:*\n"
            "  + *Gang:* Hợp kim của Fe với C (2% - 5%) cùng một lượng nhỏ Si, Mn, S, P. Rất cứng và giòn; dùng đúc bệ máy, nắp cống, đường ống nước, chi tiết máy móc.\n"
            "  + *Thép:* Hợp kim của Fe với C (dưới 2%). Độ cứng, độ dẻo và tính đàn hồi rất tốt; dùng làm vật liệu xây dựng (khung thép, cốt bê tông), chế tạo ô tô, tàu thuyền, cầu cống.\n"
            "  + *Duralumin:* Hợp kim của Al với Cu, Mg, Mn. Rất nhẹ và bền; ứng dụng quan trọng trong công nghiệp hàng không, chế tạo vỏ máy bay, tên lửa, tàu vũ trụ.\n"
            "  + *Inox (thép không gỉ):* Hợp kim của Fe với Cr (trên 10,5%), Ni, C. Khó bị gỉ, sáng bóng; dùng làm thiết bị y tế, đồ gia dụng cao cấp (nồi chảo, dao kéo, bồn rửa)."
        )
    },
    {
        "title": "4. Nguyên liệu và các công đoạn sản xuất gang thép",
        "content": (
            "• *Sản xuất Gang (trong lò cao):*\n"
            "  + *Nguyên liệu:* Quặng sắt (hematite Fe~2~O~3~, magnetite Fe~3~O~4~), than cốc (C), chất trợ dung (đá vôi CaCO~3~), không khí giàu oxygen.\n"
            "  + *Các phản ứng chính:*\n"
            "    (1) Đốt cháy than cốc: C + O~2~ -(t°)-> CO~2~ ; CO~2~ + C -(t°)-> 2CO (khí CO là chất khử chính).\n"
            "    (2) Khử quặng sắt theo từng giai đoạn:\n"
            "        3Fe~2~O~3~ + CO -(t°)-> 2Fe~3~O~4~ + CO~2~↑\n"
            "        Fe~3~O~4~ + CO -(t°)-> 3FeO + CO~2~↑\n"
            "        FeO + CO -(t°)-> Fe + CO~2~↑\n"
            "    (3) Tạo xỉ loại bỏ tạp chất SiO~2~:\n"
            "        CaCO~3~ -(t°)-> CaO + CO~2~↑ ; CaO + SiO~2~ -(t°)-> CaSiO~3~ (xỉ nhẹ nổi lên trên gang lỏng).\n"
            "• *Sản xuất Thép (trong lò luyện thép):*\n"
            "  + *Nguyên liệu:* Gang trắng hoặc gang xám, sắt thép phế liệu, khí oxygen.\n"
            "  + *Bản chất:* Thổi khí oxygen vào gang nóng chảy ở nhiệt độ cao để oxi hóa bớt carbon (giảm hàm lượng C từ 2-5% xuống dưới 2%) và các tạp chất Si, Mn, P, S."
        )
    },
    {
        "title": "5. Dãy hoạt động hóa học của kim loại và ý nghĩa",
        "content": (
            "• *Dãy hoạt động hóa học:*  *K,  Na,  Ca,  Mg,  Al,  Zn,  Fe,  Pb,  (H),  Cu,  Ag,  Au*\n"
            "• *Ý nghĩa:*\n"
            "  1. Mức độ hoạt động hóa học của kim loại giảm dần từ trái sang phải.\n"
            "  2. Kim loại đứng trước Mg (K, Na, Ca) tác dụng được với nước ở nhiệt độ thường tạo thành kiềm và giải phóng khí H~2~.\n"
            "  3. Kim loại đứng trước H tác dụng với dung dịch acid loãng (HCl, H~2~SO~4~ loãng) giải phóng khí H~2~.\n"
            "  4. Kim loại đứng trước (trừ K, Na, Ca) đẩy được kim loại đứng sau ra khỏi dung dịch muối của chúng."
        )
    },
    {
        "title": "6. Hiện tượng quan sát được và phương trình hóa học minh họa",
        "content": (
            "1. *Cho đinh sắt (Iron) vào ống nghiệm chứa dung dịch CuSO~4~:*\n"
            "   - *Hiện tượng:* Có lớp kim loại màu đỏ gạch (Cu) bám ngoài bề mặt đinh sắt; màu xanh lam của dung dịch nhạt dần.\n"
            "   - *PTHH:* Fe + CuSO~4~ -> FeSO~4~ + Cu↓\n"
            "2. *Đốt nóng đỏ một đoạn dây sắt rồi cho vào bình chứa khí oxygen:*\n"
            "   - *Hiện tượng:* Sắt cháy sáng chói, không có ngọn lửa, bắn ra các tia lửa sáng màu đỏ, sinh ra chất rắn màu nâu đen nóng chảy (oxit sắt từ Fe~3~O~4~).\n"
            "   - *PTHH:* 3Fe + 2O~2~ -(t°)-> Fe~3~O~4~\n"
            "3. *Cho mẩu Aluminium (Al) vào ống nghiệm chứa dung dịch hydrochloric acid (HCl):*\n"
            "   - *Hiện tượng:* Mẩu nhôm tan dần, có nhiều bọt khí không màu thoát ra mạnh mẽ, phản ứng tỏa nhiệt.\n"
            "   - *PTHH:* 2Al + 6HCl -> 2AlCl~3~ + 3H~2~↑\n"
            "4. *Cho mẩu Copper (Cu) vào ống nghiệm chứa dung dịch silver nitrate (AgNO~3~):*\n"
            "   - *Hiện tượng:* Mẩu đồng tan một phần, có lớp kim loại màu xám bạc (Ag) bám ngoài bề mặt mẩu đồng; dung dịch từ không màu dần chuyển sang màu xanh lam.\n"
            "   - *PTHH:* Cu + 2AgNO~3~ -> Cu(NO~3~)~2~ + 2Ag↓"
        )
    }
]

HOA_MCQ = [
    {
        "num": 1,
        "question": "Dãy kim loại sau được sắp xếp theo chiều tăng dần mức độ hoạt động hóa học:",
        "options": [
            "A. Cu, Fe, Mg, Al, K.",
            "B. K, Cu, Al, Mg, Fe.",
            "C. K, Mg, Al, Fe, Cu.",
            "D. Cu, Fe, Al, Mg, K."
        ],
        "correct": "D",
        "solution": "Theo dãy hoạt động hóa học: K > Mg > Al > Fe > Cu. Chiều tăng dần mức độ hoạt động hóa học là Cu < Fe < Al < Mg < K. Do đó chọn D."
    },
    {
        "num": 2,
        "question": "Cho các kim loại Fe, Cu, Ag, Na, Mg. Kết luận nào sau đây là sai?",
        "options": [
            "A. Kim loại tác dụng với H~2~SO~4~ loãng là Na, Fe, Mg.",
            "B. Kim loại không tác dụng với dung dịch H~2~SO~4~ loãng, HCl là Cu, Ag.",
            "C. Kim loại không tác dụng với Oxygen là Ag.",
            "D. Cả 5 kim loại không tan trong nước ở nhiệt độ thường."
        ],
        "correct": "D",
        "solution": "Kết luận D sai vì Sodium (Na) phản ứng mãnh liệt và tan hoàn toàn trong nước ở nhiệt độ thường: 2Na + 2H~2~O -> 2NaOH + H~2~↑."
    },
    {
        "num": 3,
        "question": "Ứng dụng của lưu huỳnh (sulfur) trong đời sống là",
        "options": [
            "A. Nhựa PVC.",
            "B. Lưu hóa cao su làm săm, lốp xe.",
            "C. Nilon.",
            "D. Chất tẩy rửa."
        ],
        "correct": "B",
        "solution": "Một trong những ứng dụng quan trọng hàng đầu của lưu huỳnh là dùng trong quá trình lưu hóa cao su để tạo ra cao su bền, có tính đàn hồi cao dùng chế tạo săm, lốp xe."
    },
    {
        "num": 4,
        "question": "Chlorine thường được ứng dụng để sản xuất",
        "options": [
            "A. Sulfuric acid.",
            "B. Lõi lọc nước.",
            "C. Chất tẩy rửa và nhựa PVC.",
            "D. Điện cực và ruột bút chì."
        ],
        "correct": "C",
        "solution": "Khí chlorine được sử dụng rộng rãi để sản xuất chất tẩy trắng, nước Javel, chất sát trùng và tổng hợp nhựa poly(vinyl chloride) (PVC)."
    },
    {
        "num": 5,
        "question": "Phương pháp điện phân nóng chảy được sử dụng để điều chế các kim loại nào sau đây:",
        "options": [
            "A. Al, Fe, Ag, K.",
            "B. Al, Na, Ca, K.",
            "C. Cu, Fe, Ag, K.",
            "D. Al, Cu, Ag, Na."
        ],
        "correct": "B",
        "solution": "Phương pháp điện phân nóng chảy áp dụng cho các kim loại có tính khử rất mạnh, đứng trước Zn trong dãy hoạt động hóa học như: K, Na, Ca, Ba, Mg, Al."
    },
    {
        "num": 6,
        "question": "Phương pháp nhiệt luyện được sử dụng để điều chế các kim loại",
        "options": [
            "A. Al, Fe, Ag.",
            "B. Na, Fe, K.",
            "C. Fe, Zn, Pb.",
            "D. Zn, Fe, Au."
        ],
        "correct": "C",
        "solution": "Phương pháp nhiệt luyện áp dụng cho các kim loại hoạt động trung bình và yếu (sau Al) như Fe, Zn, Pb, Sn, Cu."
    },
    {
        "num": 7,
        "question": "Đặc điểm tính chất của hợp kim inox là",
        "options": [
            "A. Dẻo và cứng.",
            "B. Khó bị gỉ.",
            "C. Có độ cứng và độ bền tương đối cao, dẫn nhiệt tốt.",
            "D. Nhẹ và bền."
        ],
        "correct": "B",
        "solution": "Inox (thép không gỉ) có đặc tính nổi trội là khó bị oxi hóa, không bị ăn mòn và gỉ sét trong không khí và môi trường nước."
    },
    {
        "num": 8,
        "question": "Đặc điểm tính chất của hợp kim duralumin là",
        "options": [
            "A. Dẻo và cứng.",
            "B. Khó bị gỉ.",
            "C. Có độ cứng và độ bền tương đối cao, dẫn nhiệt tốt.",
            "D. Nhẹ và bền."
        ],
        "correct": "D",
        "solution": "Duralumin là hợp kim của nhôm với đồng, magie, mangan với đặc tính nổi bật là nhẹ và có độ bền cơ học cao, dùng làm vật liệu chế tạo máy bay."
    },
    {
        "num": 9,
        "question": "Gang thường được sử dụng để",
        "options": [
            "A. làm chi tiết máy móc.",
            "B. làm vật liệu xây dựng.",
            "C. chế tạo vỏ máy bay.",
            "D. dao, kéo,…"
        ],
        "correct": "A",
        "solution": "Gang có tính đúc tốt, chịu lực nén cao nên thường được sử dụng để đúc bệ máy, thân máy móc, nắp cống, bánh đà..."
    },
    {
        "num": 10,
        "question": "Cho mẩu Sodium vào nước. Hiện tượng xảy ra là:",
        "options": [
            "A. Mẩu Sodium chìm trong nước và sủi bọt khí.",
            "B. Mẩu Sodium chạy vo tròn trên mặt nước, phát lửa.",
            "C. Có khí thoát ra và có kết tủa trắng.",
            "D. Có khí thoát ra và có kết tủa màu xanh lam."
        ],
        "correct": "B",
        "solution": "Do nhẹ hơn nước (d = 0,97 g/cm^3^), Sodium nổi trên mặt nước, phản ứng tỏa nhiệt mạnh làm nóng chảy thành giọt tròn chạy lăng xăng trên mặt nước và bốc cháy kèm tiếng nổ lách tách."
    },
    {
        "num": 11,
        "question": "Cho thanh Iron vào dung dịch Copper(II) chloride. Hiện tượng xảy ra là:",
        "options": [
            "A. Chất rắn đỏ gạch sinh ra, dung dịch nhạt màu dần.",
            "B. Không thấy hiện tượng gì.",
            "C. Có khí thoát ra và có kết tủa trắng.",
            "D. Có khí thoát ra và có kết tủa màu xanh lam."
        ],
        "correct": "A",
        "solution": "PTHH: Fe + CuCl~2~ -> FeCl~2~ + Cu↓. Kim loại Cu màu đỏ gạch bám ngoài thanh sắt, dung dịch CuCl~2~ màu xanh lam nhạt dần."
    },
    {
        "num": 12,
        "question": "Hoà tan hỗn hợp gồm Iron và Copper vào dung dịch hydrochloric acid (vừa đủ). Các sản phẩm thu được sau phản ứng là",
        "options": [
            "A. FeCl~2~ và khí H~2~.",
            "B. FeCl~2~, Cu và khí H~2~.",
            "C. Cu và khí H~2~.",
            "D. FeCl~2~ và Cu."
        ],
        "correct": "B",
        "solution": "Do Cu đứng sau H trong dãy HĐHH nên không phản ứng với HCl. Fe tan tạo muối FeCl~2~ và H~2~. Sau phản ứng ta thu được dung dịch FeCl~2~, khí H~2~ thoát ra và chất rắn Cu không tan còn lại."
    },
    {
        "num": 13,
        "question": "Thủy ngân dễ bay hơi và rất độc. Khi vỡ nhiệt kế thủy ngân thì dùng chất nào rắc lên để khử độc?",
        "options": [
            "A. Bột Iron.",
            "B. Bột sulfur.",
            "C. Bột than.",
            "D. Muối ăn."
        ],
        "correct": "B",
        "solution": "Bột sulfur (lưu huỳnh) phản ứng với thủy ngân ngay ở nhiệt độ thường tạo thành mercuric sulfide (HgS) dạng rắn, không bay hơi và không độc: Hg + S -> HgS."
    },
    {
        "num": 14,
        "question": "Có dung dịch muối Aluminium chloride lẫn tạp chất Copper(II) chloride. Có thể dùng kim loại nào sau đây để làm sạch:",
        "options": [
            "A. Silver.",
            "B. Magnesium.",
            "C. Aluminium.",
            "D. Copper."
        ],
        "correct": "C",
        "solution": "Dùng kim loại Aluminium dư: 2Al + 3CuCl~2~ -> 2AlCl~3~ + 3Cu↓. Kim loại Cu sinh ra và Al dư được lọc bỏ, dung dịch thu được chỉ chứa muối tinh khiết AlCl~3~."
    },
    {
        "num": 15,
        "question": "Kim loại nào có nhiệt độ nóng chảy cao nhất thường được sử dụng để làm dây tóc bóng đèn?",
        "options": [
            "A. Tungsten (W).",
            "B. Copper (Cu).",
            "C. Iron (Fe).",
            "D. Zinc (Zn)."
        ],
        "correct": "A",
        "solution": "Tungsten (Vonfram, W) có nhiệt độ nóng chảy cao nhất trong các kim loại (khoảng 3422 °C), giúp sợi đốt phát sáng mà không bị nóng chảy đứt gãy."
    },
    {
        "num": 16,
        "question": "Trong sản xuất gang người ta sử dụng than cốc để?",
        "options": [
            "A. Tạo carbon monoxide.",
            "B. Tạo hợp chất với iron.",
            "C. Tạo xỉ.",
            "D. Phản ứng với đá vôi."
        ],
        "correct": "A",
        "solution": "Than cốc cháy tạo khí CO~2~, sau đó khử tiếp than cốc tạo khí carbon monoxide (CO): C + O~2~ -> CO~2~ ; CO~2~ + C -> 2CO. Khí CO chính là chất khử để khử các oxide sắt trong quặng thành sắt tự do."
    },
    {
        "num": 17,
        "question": "Đặc điểm tính chất của gang là",
        "options": [
            "A. Dẻo và cứng.",
            "B. Khó bị gỉ.",
            "C. Có độ cứng và độ bền tương đối cao, dẫn nhiệt tốt.",
            "D. Nhẹ và bền."
        ],
        "correct": "C",
        "solution": "Gang có hàm lượng carbon từ 2-5%, đặc điểm nổi bật là có độ cứng cao, chịu tải trọng tĩnh tốt, dẫn nhiệt tốt và khả năng chống mài mòn cao."
    },
    {
        "num": 18,
        "question": "Hợp kim của sắt với một số nguyên tố khác như Cr, Ni, C... được gọi là gì?",
        "options": [
            "A. Thép thường (thép carbon).",
            "B. Gang.",
            "C. Inox.",
            "D. Duralumin."
        ],
        "correct": "C",
        "solution": "Hợp kim của sắt với chromium (từ 10,5% trở lên), niken và cacbon được gọi là thép không gỉ (Inox)."
    },
    {
        "num": 19,
        "question": "Trong phương pháp nhiệt luyện, người ta không sử dụng chất nào để phản ứng với oxide của kim loại cần tách?",
        "options": [
            "A. Al.",
            "B. C.",
            "C. CO.",
            "D. CO~2~."
        ],
        "correct": "D",
        "solution": "Phương pháp nhiệt luyện cần các chất khử như CO, C, H~2~ hoặc Al. Khí CO~2~ là hợp chất có số oxi hóa cao nhất của C (+4), không có tính khử nên không thể dùng để tách kim loại."
    },
    {
        "num": 20,
        "question": "Gang thường được sử dụng để",
        "options": [
            "A. làm chi tiết máy móc.",
            "B. làm vật liệu xây dựng.",
            "C. chế tạo vỏ máy bay.",
            "D. dao, kéo,…"
        ],
        "correct": "A",
        "solution": "Gang xám có tính đúc rất cao và chịu lực ma sát tốt nên được sử dụng rộng rãi để làm các chi tiết máy, bệ máy, nắp hố ga, đường ống dẫn."
    }
]

HOA_TF = [
    {
        "num": 1,
        "intro": "Cho các phát biểu sau về tính chất của kim loại, phát biểu nào đúng, phát biểu nào sai:",
        "items": [
            {
                "label": "a",
                "text": "Vẻ sáng lấp lánh của các kim loại ở điều kiện thường được gọi là ánh kim.",
                "ans": "Đúng",
                "explain": "Các electron tự do trong kim loại phản xạ hầu hết những tia sáng nhìn thấy được, tạo nên vẻ sáng lấp lánh gọi là ánh kim."
            },
            {
                "label": "b",
                "text": "Tungsten (W) được dùng để làm sợi đốt bóng đèn là do kim loại này dẫn điện tốt.",
                "ans": "Sai",
                "explain": "Tungsten được dùng làm sợi đốt bóng đèn là do có nhiệt độ nóng chảy cực kì cao (khoảng 3422 °C), chứ không phải do dẫn điện tốt nhất (Ag và Cu dẫn điện tốt hơn nhiều)."
            },
            {
                "label": "c",
                "text": "Kim loại thủy ngân thường được dùng trong nhiệt kế, áp kế.",
                "ans": "Đúng",
                "explain": "Thủy ngân là kim loại duy nhất ở thể lỏng tại điều kiện thường, giãn nở vì nhiệt đều đặn nên được dùng phổ biến trong nhiệt kế và áp kế."
            },
            {
                "label": "d",
                "text": "Sắt là kim loại có độ cứng cao nhất trong các kim loại.",
                "ans": "Sai",
                "explain": "Kim loại có độ cứng cao nhất là Chromium (Cr), có thể rạch được thủy tinh; kim cương là chất cứng nhất."
            }
        ]
    },
    {
        "num": 2,
        "intro": "Thực hiện thí nghiệm với một kim loại X chưa biết tên, kết quả thu được như sau:\n- X đẩy được sắt ra khỏi muối FeSO~4~ trong dung dịch.\n- X không phản ứng với nước ở điều kiện nhiệt độ thường.\nCho các phát biểu sau, phát biểu nào đúng, phát biểu nào sai:",
        "items": [
            {
                "label": "a",
                "text": "X có thể phản ứng với hơi nước ở nhiệt độ cao.",
                "ans": "Đúng",
                "explain": "Các kim loại như Mg, Zn hoạt động mạnh hơn Fe (đẩy được Fe khỏi FeSO~4~) và không phản ứng với nước lạnh nhưng phản ứng tốt với hơi nước ở nhiệt độ cao."
            },
            {
                "label": "b",
                "text": "X hoạt động hoá học yếu hơn sắt.",
                "ans": "Sai",
                "explain": "X đẩy được Fe ra khỏi dung dịch muối FeSO~4~ chứng tỏ X phải hoạt động hóa học mạnh hơn sắt."
            },
            {
                "label": "c",
                "text": "X phản ứng được với dung dịch HCl sinh ra khí H~2~.",
                "ans": "Đúng",
                "explain": "Vì X đứng trước sắt trong dãy hoạt động hóa học nên X chắc chắn đứng trước Hydrogen (H), do đó phản ứng được với dd HCl giải phóng H~2~."
            },
            {
                "label": "d",
                "text": "X là kim loại kiềm.",
                "ans": "Sai",
                "explain": "Kim loại kiềm (nhóm IA như Na, K) phản ứng mãnh liệt với nước ngay ở điều kiện thường, trong khi X không phản ứng với nước ở nhiệt độ thường."
            }
        ]
    }
]

HOA_SHORT = [
    {
        "num": 1,
        "question": "Trong gang có chứa bao nhiêu % là carbon? Thép có chứa bao nhiêu % là carbon?",
        "solution": "Trong gang hàm lượng carbon chiếm từ *2% đến 5%*. Trong thép hàm lượng carbon chiếm *dưới 2%* (thường từ 0,01% đến 2%)."
    },
    {
        "num": 2,
        "question": "Cho các chất sau: H~2~O (điều kiện thường), H~2~O (nhiệt độ cao), H~2~SO~4~, O~2~, Cl~2~, NaOH, NaCl, AlCl~3~, AgNO~3~. Có bao nhiêu chất và là những chất nào tác dụng được với Fe?",
        "solution": (
            "Có *5 chất* tác dụng được với Fe:\n"
            "1. H~2~O (ở nhiệt độ cao): 3Fe + 4H~2~O -(t° < 570°C)-> Fe~3~O~4~ + 4H~2~↑\n"
            "2. H~2~SO~4~: Fe + H~2~SO~4~ -> FeSO~4~ + H~2~↑\n"
            "3. O~2~: 3Fe + 2O~2~ -(t°)-> Fe~3~O~4~\n"
            "4. Cl~2~: 2Fe + 3Cl~2~ -(t°)-> 2FeCl~3~\n"
            "5. AgNO~3~: Fe + 2AgNO~3~ -> Fe(NO~3~)~2~ + 2Ag↓\n"
            "*(Không tác dụng: H~2~O thường, NaOH, NaCl, AlCl~3~)*."
        )
    },
    {
        "num": 3,
        "question": "Muốn sản xuất 5 tấn thép chứa 98% sắt, cần số tấn gang chứa 94,5% sắt là bao nhiêu? Biết hiệu suất của toàn bộ quá trình bằng 85%.",
        "solution": (
            "• Khối lượng Fe nguyên chất cần có trong 5 tấn thép: m~Fe(thép)~ = 5 . 98% = 4,9 tấn.\n"
            "• Do hiệu suất quá trình đạt 85%, nên lượng Fe lý thuyết cần lấy từ gang là:\n"
            "  m~Fe(cần)~ = 4,9 / 85% = 4,9 / 0,85 ≈ 5,7647 tấn.\n"
            "• Vì gang chỉ chứa 94,5% Fe nên khối lượng gang thực tế cần dùng là:\n"
            "  *m~gang~ = 5,7647 / 94,5% ≈ 6,10 tấn (khoảng 6,1 tấn)*."
        )
    },
    {
        "num": 4,
        "question": "Có các nguyên liệu: (1) Quặng sắt; (2) Quặng Cromit; (3) Quặng Boxit; (4) Than cốc; (5) Than đá; (6) CaCO~3~; (7) O~2~. Những nguyên liệu nào dùng để luyện gang?",
        "solution": "Những nguyên liệu chính dùng để luyện gang là: *(1) Quặng sắt, (4) Than cốc, (6) CaCO~3~ (chất trợ dung), (7) O~2~ (không khí giàu oxygen)*."
    },
    {
        "num": 5,
        "question": "Cho 2 ví dụ kim loại tác dụng được với khí Chlorine và dung dịch hydrochloric acid tạo ra cùng một loại muối?",
        "solution": (
            "Các kim loại chỉ thể hiện một mức hóa trị duy nhất (không đổi) khi phản ứng với Cl~2~ và HCl, ví dụ:\n"
            "• *Aluminium (Al):*\n"
            "  2Al + 3Cl~2~ -(t°)-> 2AlCl~3~  và  2Al + 6HCl -> 2AlCl~3~ + 3H~2~↑  (cùng tạo muối AlCl~3~).\n"
            "• *Zinc (Zn):*\n"
            "  Zn + Cl~2~ -(t°)-> ZnCl~2~  và  Zn + 2HCl -> ZnCl~2~ + H~2~↑  (cùng tạo muối ZnCl~2~).\n"
            "*(Có thể lấy thêm Magnesium Mg: tạo cùng muối MgCl~2~)*."
        )
    },
    {
        "num": 6,
        "question": "Cho các chất dưới đây: (1). Lưu huỳnh (sulfur); (2). Sắt (Iron); (3). Đồng (Copper); (4). Nhôm (Aluminium); (5). Phosphorus. Có bao nhiêu chất có thể dùng làm lõi dây dẫn điện?",
        "solution": (
            "Có *2 chất* được ứng dụng phổ biến làm lõi dây dẫn điện là: *(3) Đồng (Copper)* và *(4) Nhôm (Aluminium)* nhờ khả năng dẫn điện xuất sắc, tính dẻo dai và giá thành phù hợp trong công nghiệp điện năng."
        )
    }
]

# =============================================================================
# PHÂN MÔN 3: VẬT SỐNG (SINH HỌC)
# =============================================================================
SINHHOC_MCQ = [
    {
        "num": 1,
        "question": "Tên gọi đầy đủ của phân tử DNA là",
        "options": [
            "A. deoxyribonucleic acid.",
            "B. ribonucleic acid.",
            "C. nucleotide.",
            "D. nucleic acid."
        ],
        "correct": "A",
        "solution": "DNA là chữ viết tắt của deoxyribonucleic acid (axit đêoxiribônuclêic)."
    },
    {
        "num": 2,
        "question": "Mục đích của phép lai phân tích nhằm xác định",
        "options": [
            "A. kiểu gene, kiểu hình của cá thể mang tính trạng trội.",
            "B. kiểu hình của cá thể mang tính trạng trội.",
            "C. kiểu gene của tất cả các tính trạng.",
            "D. kiểu gene của cá thể mang tính trạng trội."
        ],
        "correct": "D",
        "solution": "Phép lai phân tích là phép lai giữa cá thể mang tính trạng trội cần xác định kiểu gene với cá thể mang tính trạng lặn nhằm kiểm tra kiểu gene của cá thể mang tính trạng trội đó là đồng hợp tử hay dị hợp tử."
    }
]

SINHHOC_SHORT = [
    {
        "num": 3,
        "question": (
            "Trong một thí nghiệm thực hành, người ta tiến hành giao phấn giữa hai dòng đậu Hà Lan cây hạt trơn thuần chủng và cây hạt nhăn thuần chủng. "
            "Thế hệ F~1~ thu được 100% cây hạt trơn. Cho các cây F~1~ tự thụ phấn, ở thế hệ F~2~ thu được 300 cây hạt trơn và 100 cây hạt nhăn. "
            "Trong các nhận định sau về kết quả thí nghiệm đã cho, có bao nhiêu nhận định đúng?\n"
            "(1) Hạt trơn và hạt nhăn là một cặp tính trạng tương phản.\n"
            "(2) Cây hạt trơn ở thế hệ P là cơ thể mang tính trạng lặn.\n"
            "(3) Cơ thể thuần chủng là cây hạt trơn thuần chủng hoặc cây hạt nhăn thuần chủng ở thế hệ P.\n"
            "(4) Hạt nhăn là tính trạng trội và chỉ biểu hiện ở cơ thể có kiểu gene đồng hợp lặn."
        ),
        "solution": (
            "• Phân tích từng nhận định:\n"
            "  (1) *Đúng:* Hạt trơn và hạt nhăn là hai trạng thái biểu hiện trái ngược nhau của cùng một loại tính trạng hình dạng hạt, do đó là cặp tính trạng tương phản.\n"
            "  (2) *Sai:* P thuần chủng trơn × nhăn -> F~1~ 100% hạt trơn, F~2~ tỉ lệ 3 trơn : 1 nhăn => Hạt trơn là tính trạng *trội*, hạt nhăn là tính trạng *lặn*. Vậy cây hạt trơn ở P mang tính trạng trội.\n"
            "  (3) *Đúng:* Cả hai cây P đều thuần chủng (đồng hợp tử).\n"
            "  (4) *Sai:* Hạt nhăn là tính trạng lặn, không phải tính trạng trội.\n"
            "• Vậy có tất cả *2 nhận định đúng* (nhận định (1) và (3)).\n"
            "• *Đáp số:* **2**"
        )
    }
]

SINHHOC_ESSAY = [
    {
        "num": 4,
        "question": "Ở đậu Hà Lan, tiến hành lai giữa các cá thể thuần chủng thân cao với thân thấp. F~1~ thu được 100% cây thân cao. F~2~ thu được cả cây thân cao và cây thân thấp với tỉ lệ 3 cây thân cao : 1 cây thân thấp. Sử dụng kí hiệu và thuật ngữ để mô tả thí nghiệm trên bằng sơ đồ lai.",
        "solution": (
            "• *Quy ước gene:*\n"
            "  Gene A quy định tính trạng thân cao (trội hoàn toàn).\n"
            "  Gene a quy định tính trạng thân thấp (lặn hoàn toàn).\n"
            "  Cây thân cao thuần chủng có kiểu gene AA; cây thân thấp thuần chủng có kiểu gene aa.\n"
            "• *Sơ đồ lai:*\n"
            "  P~tc~:            Thân cao (AA)       ×       Thân thấp (aa)\n"
            "  G~P~:                  A                             a\n"
            "  F~1~:                        100% Aa (Thân cao)\n"
            "  F~1~ × F~1~:       Thân cao (Aa)       ×       Thân cao (Aa)\n"
            "  G~F1~:             1/2 A : 1/2 a             1/2 A : 1/2 a\n"
            "  F~2~:               1 AA : 2 Aa : 1 aa\n"
            "• *Kết quả thế hệ F~2~:*\n"
            "  - Tỉ lệ kiểu gene: *1 AA : 2 Aa : 1 aa* (25% AA : 50% Aa : 25% aa).\n"
            "  - Tỉ lệ kiểu hình: *3 Thân cao : 1 Thân thấp* (75% thân cao : 25% thân thấp)."
        )
    },
    {
        "num": 5,
        "question": "Ở cà chua, gen A quy định hoa đỏ trội hoàn toàn so với gen a quy định hoa trắng. Cho hai cây cà chua thuần chủng khác nhau về màu sắc hoa lai với nhau.\na) Viết sơ đồ lai và xác định tỉ lệ kiểu gen, kiểu hình ở F~1~.\nb) Cho các cây F~1~ tự thụ phấn, hãy viết sơ đồ lai và xác định tỉ lệ kiểu gen, kiểu hình ở F~2~.",
        "solution": (
            "a) *Sơ đồ lai từ P đến F~1~:*\n"
            "   Cây hoa đỏ thuần chủng: AA; Cây hoa trắng thuần chủng: aa.\n"
            "   P~tc~:          Hoa đỏ (AA)         ×         Hoa trắng (aa)\n"
            "   G~P~:                A                               a\n"
            "   F~1~:                      100% Aa (Hoa đỏ)\n"
            "   - Tỉ lệ kiểu gene F~1~: *100% Aa*.\n"
            "   - Tỉ lệ kiểu hình F~1~: *100% Hoa đỏ*.\n"
            "b) *Sơ đồ lai F~1~ tự thụ phấn:*\n"
            "   F~1~ × F~1~:     Hoa đỏ (Aa)         ×         Hoa đỏ (Aa)\n"
            "   G~F1~:           1/2 A : 1/2 a               1/2 A : 1/2 a\n"
            "   F~2~:             1/4 AA : 2/4 Aa : 1/4 aa\n"
            "   - Tỉ lệ kiểu gene F~2~: *1 AA : 2 Aa : 1 aa* (1/4 AA : 1/2 Aa : 1/4 aa).\n"
            "   - Tỉ lệ kiểu hình F~2~: *3 Hoa đỏ : 1 Hoa trắng* (3/4 hoa đỏ : 1/4 hoa trắng)."
        )
    },
    {
        "num": 6,
        "question": (
            "Ở đậu Hà Lan, allele A quy định hạt vàng trội hoàn toàn so với allele a quy định hạt xanh; allele B quy định vỏ trơn trội hoàn toàn so với allele b quy định vỏ nhăn. Hai cặp gene nằm trên hai cặp nhiễm sắc thể tương đồng khác nhau và phân li độc lập.\n"
            "Cho lai giữa cây đậu Hà Lan thuần chủng hạt vàng, vỏ trơn với cây thuần chủng hạt xanh, vỏ nhăn, thu được thế hệ F~1~. Cho các cây F~1~ tự thụ phấn, thu được F~2~.\n"
            "a. Xác định kiểu gene của các cây bố mẹ (P) và kiểu gene, kiểu hình của F~1~.\n"
            "b. Lập sơ đồ lai từ P đến F~2~. Xác định tỉ lệ các loại kiểu gene và tỉ lệ các loại kiểu hình ở F~2~."
        ),
        "solution": (
            "a) *Kiểu gene của bố mẹ P và F~1~:*\n"
            "   - Cây hạt vàng, vỏ trơn thuần chủng có kiểu gene: *AABB*.\n"
            "   - Cây hạt xanh, vỏ nhăn thuần chủng có kiểu gene: *aabb*.\n"
            "   - Giao tử P: AABB cho giao tử AB; aabb cho giao tử ab.\n"
            "   - Kiểu gene F~1~: *100% AaBb*.\n"
            "   - Kiểu hình F~1~: *100% Cây hạt vàng, vỏ trơn*.\n"
            "b) *Sơ đồ lai từ P đến F~2~:*\n"
            "   P~tc~:          AABB (Vàng, trơn)      ×      aabb (Xanh, nhăn)\n"
            "   G~P~:                  AB                            ab\n"
            "   F~1~:                       100% AaBb (Vàng, trơn)\n"
            "   F~1~ × F~1~:    AaBb (Vàng, trơn)      ×      AaBb (Vàng, trơn)\n"
            "   G~F1~:          1/4 AB : 1/4 Ab : 1/4 aB : 1/4 ab (mỗi bên)\n"
            "   • *Bảng tổ hợp F~2~ (16 tổ hợp giao tử):*\n"
            "     - Tỉ lệ kiểu gene F~2~:\n"
            "       1 AABB : 2 AABb : 2 AaBB : 4 AaBb : 1 AAbb : 2 Aabb : 1 aaBB : 2 aaBb : 1 aabb\n"
            "       (Tỉ lệ: *1 : 2 : 2 : 4 : 1 : 2 : 1 : 2 : 1* gồm 9 loại kiểu gene).\n"
            "     - Tỉ lệ kiểu hình F~2~:\n"
            "       + 9 Hạt vàng, vỏ trơn (9 A-B-)\n"
            "       + 3 Hạt vàng, vỏ nhăn (3 A-bb)\n"
            "       + 3 Hạt xanh, vỏ trơn (3 aaB-)\n"
            "       + 1 Hạt xanh, vỏ nhăn (1 aabb)\n"
            "       => *Tỉ lệ phân li kiểu hình: 9 Hạt vàng, vỏ trơn : 3 Hạt vàng, vỏ nhăn : 3 Hạt xanh, vỏ trơn : 1 Hạt xanh, vỏ nhăn*."
        )
    }
]

SINHHOC_TF = [
    {
        "num": 7,
        "intro": "Trong các phát biểu sau đây về nội dung các quy luật di truyền của Mendel, mỗi phát biểu a), b), c), d) là đúng hay sai?",
        "items": [
            {
                "label": "a",
                "text": "Theo quy luật phân li, mỗi tính trạng do một cặp nhân tố di truyền (cặp allele) quy định; khi giảm phân hình thành giao tử, các allele trong cặp phân li đồng đều về các giao tử.",
                "ans": "Đúng",
                "explain": "Đây chính là nội dung cốt lõi của quy luật phân li của Menđen: mỗi cơ thể chứa cặp nhân tố di truyền, khi tạo giao tử thì mỗi nhân tố di truyền phân li về một giao tử với xác suất ngang nhau."
            },
            {
                "label": "b",
                "text": "Phép lai phân tích là phép lai giữa cơ thể mang tính trạng lặn với cơ thể mang tính trạng lặn khác để xác định độ thuần chủng của cơ thể mang tính trạng lặn.",
                "ans": "Sai",
                "explain": "Cơ thể mang tính trạng lặn luôn có kiểu gene đồng hợp lặn thuần chủng. Phép lai phân tích là lai cá thể mang *tính trạng trội* với cá thể mang *tính trạng lặn* nhằm xác định kiểu gene của cá thể mang tính trạng trội."
            },
            {
                "label": "c",
                "text": "Khi cho lai hai giống đậu Hà Lan thuần chủng khác nhau về hai cặp tính trạng tương phản (P~tc~: Hạt vàng, vỏ trơn × Hạt xanh, vỏ nhăn), ở thế hệ F~2~ thu được tỉ lệ phân li kiểu hình xấp xỉ 9: 3: 3: 1.",
                "ans": "Đúng",
                "explain": "Hai cặp gene phân li độc lập tổ hợp tự do tạo ra tỉ lệ kiểu hình ở F~2~ là (3 vàng : 1 xanh)(3 trơn : 1 nhăn) = 9 vàng, trơn : 3 vàng, nhăn : 3 xanh, trơn : 1 xanh, nhăn."
            },
            {
                "label": "d",
                "text": "Trong phép lai một cặp tính trạng của Mendel, ở thế hệ F~2~ thu được 100% cá thể có kiểu hình mang tính trạng trội.",
                "ans": "Sai",
                "explain": "Thế hệ F~1~ mới thu được 100% mang tính trạng trội; ở thế hệ F~2~ xảy ra hiện tượng phân li tính trạng theo tỉ lệ xấp xỉ 3 trội : 1 lặn."
            }
        ]
    }
]
