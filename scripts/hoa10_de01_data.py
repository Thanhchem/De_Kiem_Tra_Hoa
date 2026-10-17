# -*- coding: utf-8 -*-
"""
Dữ liệu Đề kiểm tra Giữa học kì I - Môn Hóa học 10 (Mã đề 101 / Đề thêm 01)
Năm học 2026 - 2027 - Thầy TRẦN VĂN THẠNH
Bao gồm:
- Phần I: 12 câu trắc nghiệm nhiều lựa chọn (3 điểm)
- Phần II: 2 câu trắc nghiệm đúng/sai (2 điểm)
- Phần III: 8 câu trả lời ngắn (2 điểm)
- Phần IV: 3 câu tự luận (3 điểm)
Đầy đủ đáp án, phương án, phân tích và lời giải chi tiết chuẩn mực theo chương trình GDPT 2018.
"""

EXAM_INFO = {
    "title_left_1": "SỞ GIÁO DỤC VÀ ĐÀO TẠO",
    "title_left_2": "TRƯỜNG THPT CHUYÊN / THPT",
    "sub_left": "(Đề thi gồm có 03 trang)",
    "title_right_1": "ĐỀ KIỂM TRA GIỮA HỌC KÌ I",
    "title_right_2": "MÔN: HÓA HỌC - LỚP 10",
    "time_str": "Thời gian làm bài: 45 phút (Không kể thời gian phát đề)",
    "made": "101",
    "year": "NĂM HỌC 2026 - 2027",
}

# -----------------------------------------------------------------------------
# PHẦN I: 12 CÂU TRẮC NGHIỆM NHIỀU LỰA CHỌN (3,0 ĐIỂM)
# -----------------------------------------------------------------------------
PART_1_QUESTIONS = [
    {
        "id": 1,
        "content": "Phát biểu nào sau đây *không đúng*?",
        "options": [
            "A. Số khối của hạt nhân bằng tổng số proton và số neutron.",
            "B. Nguyên tố hoá học là những nguyên tử có cùng số đơn vị điện tích hạt nhân.",
            "C. Số hiệu nguyên tử bằng số đơn vị điện tích hạt nhân nguyên tử.",
            "D. Trong nguyên tử, số đơn vị điện tích hạt nhân bằng số proton và bằng số neutron."
        ],
        "correct": "D",
        "explanation": "Trong nguyên tử trung hòa về điện, số đơn vị điện tích hạt nhân bằng số proton và bằng số electron ($Z = p = e$). Số neutron ($N$) trong hạt nhân không nhất thiết phải bằng số proton (trừ một số ít nguyên tử nhẹ như $^4_2He, ^{12}_6C, ^{14}_7N, ^{16}_8O$). Do đó phát biểu D là không đúng."
    },
    {
        "id": 2,
        "content": "Nguyên tử chứa những hạt mang điện là",
        "options": [
            "A. proton và electron.",
            "B. proton và α.",
            "C. electron và neutron.",
            "D. proton và neutron."
        ],
        "correct": "A",
        "explanation": "Nguyên tử được cấu tạo từ ba loại hạt cơ bản: proton (mang điện tích dương $+1$), electron (mang điện tích âm $-1$), và neutron (không mang điện). Do đó, các hạt mang điện trong nguyên tử gồm proton và electron."
    },
    {
        "id": 3,
        "content": "Sự phân bố electron theo orbital nào dưới đây là đúng?",
        "options_type": "images",
        "options": [
            {"label": "A", "image": "page_1_img_1_48.png"},
            {"label": "B", "image": "page_1_img_2_50.png"},
            {"label": "C", "image": "page_1_img_3_52.png"},
            {"label": "D", "image": "page_1_img_4_54.png"}
        ],
        "correct": "B",
        "explanation": "- Phương án A: Ô orbital thứ nhất có 2 electron cùng chiều spin (cùng hướng lên $\\uparrow\\uparrow$), vi phạm nguyên lí Pauli (trong một orbital, hai electron phải có chiều tự quay khác nhau $\\uparrow\\downarrow$).\n- Phương án B: Cả 3 electron phân bố vào 3 orbital riêng biệt với chiều tự quay giống nhau (cùng hướng lên $\\uparrow$), hoàn toàn thỏa mãn nguyên lí Pauli và quy tắc Hund (số electron độc thân là tối đa).\n- Phương án C: Một orbital có 2 electron cùng spin $\\uparrow\\uparrow$, vi phạm nguyên lí Pauli.\n- Phương án D: Có 3 electron nhưng 2 electron đã ghép đôi trong 1 orbital trong khi còn 1 orbital trống, vi phạm quy tắc Hund."
    },
    {
        "id": 4,
        "content": "Lớp L có số phân lớp electron bằng",
        "options": [
            "A. 1.",
            "B. 4.",
            "C. 2.",
            "D. 3."
        ],
        "correct": "C",
        "explanation": "Lớp electron thứ $n$ có $n$ phân lớp (với $n \\le 4$):\n- Lớp K ($n = 1$) có 1 phân lớp (1s).\n- Lớp L ($n = 2$) có 2 phân lớp (2s và 2p).\n- Lớp M ($n = 3$) có 3 phân lớp (3s, 3p, 3d).\n- Lớp N ($n = 4$) có 4 phân lớp (4s, 4p, 4d, 4f)."
    },
    {
        "id": 5,
        "content": "Các hạt cấu tạo nên hạt nhân của hầu hết các nguyên tử là",
        "options": [
            "A. electron và neutron.",
            "B. electron, proton và neutron.",
            "C. proton và neutron.",
            "D. neutron và electron."
        ],
        "correct": "C",
        "explanation": "Hạt nhân của hầu hết các nguyên tử được cấu tạo từ hai loại hạt là proton và neutron (thường gọi chung là nucleon). Riêng nguyên tử hydrogen thông thường ($^1_1H$) hạt nhân chỉ gồm 1 proton, không có neutron."
    },
    {
        "id": 6,
        "content": "Kí hiệu chung của mọi nguyên tử là ^A_ZX, trong đó A, Z và X lần lượt là",
        "options": [
            "A. số khối, số hiệu nguyên tử, kí hiệu nguyên tố.",
            "B. số khối, kí hiệu nguyên tố, số hiệu nguyên tử.",
            "C. số hiệu nguyên tử, kí hiệu nguyên tố, số khối.",
            "D. số hiệu nguyên tử, số khối, kí hiệu nguyên tố."
        ],
        "correct": "A",
        "explanation": "Theo quy ước quốc tế về kí hiệu nguyên tử $^A_ZX$:\n- $A$ (viết ở phía trên bên trái): Số khối ($A = Z + N$).\n- $Z$ (viết ở phía dưới bên trái): Số hiệu nguyên tử (số proton, điện tích hạt nhân).\n- $X$: Kí hiệu hoá học của nguyên tố."
    },
    {
        "id": 7,
        "content": "Nguyên tử của nguyên tố X có electron cuối cùng điền vào phân lớp 3p^1. Nguyên tử của nguyên tố Y có electron cuối cùng điền vào phân lớp 3p^3. Số proton của X, Y lần lượt là",
        "options": [
            "A. 12 và 14.",
            "B. 13 và 15.",
            "C. 12 và 15.",
            "D. 13 và 14."
        ],
        "correct": "B",
        "explanation": "- Cấu hình electron của X: $1s^2 2s^2 2p^6 3s^2 3p^1 \\Rightarrow$ Tổng số electron = số proton $= 2 + 2 + 6 + 2 + 1 = 13$ (Aluminium - Al).\n- Cấu hình electron của Y: $1s^2 2s^2 2p^6 3s^2 3p^3 \\Rightarrow$ Tổng số electron = số proton $= 2 + 2 + 6 + 2 + 3 = 15$ (Phosphorus - P).\nVậy số proton của X và Y lần lượt là 13 và 15."
    },
    {
        "id": 8,
        "content": "Orbital s có dạng",
        "options": [
            "A. hình cầu.",
            "B. hình số 8 nổi.",
            "C. hình tròn.",
            "D. hình bầu dục."
        ],
        "correct": "A",
        "explanation": "- Orbital s có dạng hình cầu trong không gian 3 chiều đối xứng qua tâm hạt nhân.\n- Orbital p có dạng hình số 8 nổi gồm hai phần đối xứng qua hạt nhân nguyên tử ($p_x, p_y, p_z$)."
    },
    {
        "id": 9,
        "content": "Dãy nào sau đây gồm các đồng vị của cùng một nguyên tố hoá học?",
        "options": [
            "A. ^14_6X, ^14_7Y, ^14_8Z.",
            "B. ^28_14X, ^29_14Y, ^30_14Z.",
            "C. ^40_18X, ^40_19Y, ^40_20Z.",
            "D. ^19_9X, ^19_10Y, ^20_10Z."
        ],
        "correct": "B",
        "explanation": "Đồng vị là những nguyên tử có cùng số proton (cùng số hiệu nguyên tử $Z$, cùng thuộc một nguyên tố hoá học) nhưng khác nhau về số neutron (dẫn đến số khối $A$ khác nhau).\n- Dãy B: Cả 3 nguyên tử đều có số hiệu nguyên tử $Z = 14$ (nguyên tố Silicon - Si) và số khối lần lượt là 28, 29, 30 $\\Rightarrow$ Là các đồng vị của nhau."
    },
    {
        "id": 10,
        "content": "Lithium (Z = 3) là một nguyên tố có nhiều công dụng, được sử dụng trong chế tạo máy bay và trong một số loại pin nhất định. Pin Lithium-Ion (Li-Ion) đang ngày càng phổ biến, nó cung cấp năng lượng cho cuộc sống của hàng triệu người mỗi ngày thông qua các thiết bị như máy tính xách tay, điện thoại di động, xe Hybrid, xe điện,… nhờ trọng lượng nhẹ, cung cấp năng lượng cao và khả năng sạc lại. Tính chất hóa học cơ bản của lithium là",
        "options": [
            "A. khí hiếm.",
            "B. vừa kim loại vừa phi kim.",
            "C. kim loại.",
            "D. phi kim."
        ],
        "correct": "C",
        "explanation": "Lithium có số hiệu nguyên tử $Z = 3$, cấu hình electron: $1s^2 2s^1$. Nguyên tử có 1 electron ở lớp ngoài cùng ($2s^1$) và có xu hướng dễ nhường 1 electron để đạt cấu hình bền vững của khí hiếm He ($1s^2$). Do đó, Lithium là một kim loại (thuộc nhóm IA - kim loại kiềm)."
    },
    {
        "id": 11,
        "content": "Cấu hình electron của nguyên tử P (Z = 15) là",
        "options": [
            "A. 1s^2 2s^2 2p^6 3s^2 3p^5.",
            "B. 1s^2 2s^2 2p^6 3s^2 3p^3.",
            "C. 1s^2 2s^2 2p^6 3s^1.",
            "D. 1s^2 2s^2 2p^6 3s^2 3p^2."
        ],
        "correct": "B",
        "explanation": "Nguyên tử P có $Z = 15$. Phân bố 15 electron theo mức năng lượng từ thấp đến cao: $1s^2 2s^2 2p^6 3s^2 3p^3$ (lớp thứ 1 có 2e, lớp thứ 2 có 8e, lớp thứ 3 có 5e)."
    },
    {
        "id": 12,
        "content": "Số proton và số neutron có trong một nguyên tử aluminium (^27_13Al) lần lượt là",
        "options": [
            "A. 12 và 14.",
            "B. 13 và 14.",
            "C. 13 và 13.",
            "D. 13 và 15."
        ],
        "correct": "B",
        "explanation": "Trong nguyên tử aluminium $^{27}_{13}Al$:\n- Số hiệu nguyên tử $Z = 13 \\Rightarrow$ Số proton $p = Z = 13$.\n- Số khối $A = 27 \\Rightarrow$ Số neutron $n = A - Z = 27 - 13 = 14$."
    }
]

# -----------------------------------------------------------------------------
# PHẦN II: 2 CÂU TRẮC NGHIỆM ĐÚNG / SAI (2,0 ĐIỂM)
# -----------------------------------------------------------------------------
PART_2_QUESTIONS = [
    {
        "id": 1,
        "intro": "Cho hình ảnh một số nguyên tố sau trong bảng tuần hoàn:",
        "image": "page_2_img_1_62.png",
        "prompt": "Chọn đúng hoặc sai cho các kết luận sau:",
        "statements": [
            {
                "sub_id": "a",
                "text": "N và P có cùng số electron lớp ngoài cùng.",
                "correct": "Đúng",
                "explanation": "N ($Z = 7$: $1s^2 2s^2 2p^3$) và P ($Z = 15$: $1s^2 2s^2 2p^6 3s^2 3p^3$) đều cùng thuộc nhóm VA trong bảng tuần hoàn nên đều có 5 electron ở lớp ngoài cùng ($ns^2 np^3$)."
            },
            {
                "sub_id": "b",
                "text": "Điện tích hạt nhân của S là +32,06.",
                "correct": "Sai",
                "explanation": "Giá trị 32,06 ghi trong ô nguyên tố là nguyên tử khối trung bình của Sulfur (amu). Sulfur có số hiệu nguyên tử $Z = 16$ nên hạt nhân có 16 proton, do đó điện tích hạt nhân của S là $+16$ (hoặc $+16e$)."
            },
            {
                "sub_id": "c",
                "text": "C và Li có cùng số lớp electron.",
                "correct": "Đúng",
                "explanation": "C ($Z = 6$: $1s^2 2s^2 2p^2$) và Li ($Z = 3$: $1s^2 2s^1$) đều thuộc chu kì 2 của bảng tuần hoàn, do đó nguyên tử của chúng đều có 2 lớp electron."
            },
            {
                "sub_id": "d",
                "text": "Có tổng 18 nguyên tố hóa học ở các chu kì nhỏ.",
                "correct": "Đúng",
                "explanation": "Các chu kì nhỏ gồm: Chu kì 1 (có 2 nguyên tố: H, He), Chu kì 2 (có 8 nguyên tố: Li đến Ne), Chu kì 3 (có 8 nguyên tố: Na đến Ar). Tổng số nguyên tố ở các chu kì nhỏ là $2 + 8 + 8 = 18$ nguyên tố."
            }
        ]
    },
    {
        "id": 2,
        "intro": "Năm 1911, Rơ-đơ-pho (E. Rutherford) và các cộng sự đã dùng các hạt α bắn phá lá vàng mỏng và dùng màn huỳnh quang đặt sau lá vàng để theo dõi đường đi của các hạt α. Đa số hạt α xuyên qua nguyên tử, một số hạt va chạm hạt nhân nguyên tử bị bật ngược trở lại hoặc bật ra hướng khác. Kết quả thí nghiệm đã rút ra các kết luận về nguyên tử như sau:",
        "image": None,
        "prompt": "Chọn đúng hoặc sai cho các kết luận sau:",
        "statements": [
            {
                "sub_id": "a",
                "text": "Nguyên tử rỗng là nguyên nhân vì sao thể tích chất khí giảm một nửa khi tăng gấp đôi áp suất trong cùng điều kiện nhiệt độ.",
                "correct": "Sai",
                "explanation": "Thể tích chất khí giảm khi tăng áp suất là do khoảng cách giữa các phân tử khí rất lớn so với kích thước bản thân phân tử, khi tăng áp suất thì các phân tử bị nén lại gần nhau hơn, chứ không phải do cấu tạo rỗng bên trong bản thân mỗi nguyên tử."
            },
            {
                "sub_id": "b",
                "text": "Hạt nhân nguyên tử có kích thước rất nhỏ so với kích thước nguyên tử.",
                "correct": "Đúng",
                "explanation": "Hầu hết các hạt $\\alpha$ truyền thẳng qua lá vàng mà không bị lệch hướng, chỉ có một tỉ lệ rất nhỏ (khoảng 1/8000) bị lệch hướng hoặc bật ngược trở lại. Điều này chứng minh hạt nhân mang điện tích dương có kích thước vô cùng nhỏ bé so với kích thước toàn bộ nguyên tử (đường kính hạt nhân nhỏ hơn đường kính nguyên tử khoảng $10^4$ đến $10^5$ lần)."
            },
            {
                "sub_id": "c",
                "text": "Hạt nhân nguyên tử tương đối đặc khít nên có một số hạt α bị bật ngược trở lại.",
                "correct": "Đúng",
                "explanation": "Mặc dù có kích thước rất nhỏ, nhưng hạt nhân lại tập trung hầu như toàn bộ khối lượng của nguyên tử (mật độ khối lượng cực kì lớn, đặc khít). Khi hạt $\\alpha$ (khối lượng xấp xỉ 4 amu) va chạm trực diện với hạt nhân vàng (khối lượng khoảng 197 amu) thì bị bật ngược trở lại."
            },
            {
                "sub_id": "d",
                "text": "Nguyên tử có cấu tạo rỗng nên hạt α xuyên qua được.",
                "correct": "Đúng",
                "explanation": "Đại đa số các hạt $\\alpha$ mang điện tích dương đi thẳng xuyên qua lá vàng dày hàng nghìn lớp nguyên tử mà không hề bị cản trở, chứng tỏ phần lớn không gian bên trong nguyên tử là khoảng trống (cấu tạo rỗng)."
            }
        ]
    }
]

# -----------------------------------------------------------------------------
# PHẦN III: 8 CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM)
# -----------------------------------------------------------------------------
PART_3_QUESTIONS = [
    {
        "id": 1,
        "content": "Kẽm là kim loại có ứng dụng quan trọng để bảo vệ sắt. Kí hiệu nguyên tử là ^65_30Zn. Tính tổng số hạt proton, neutron và electron trong nguyên tử Zn.",
        "answer": "95",
        "explanation": "Trong nguyên tử $^{65}_{30}Zn$:\n- Số proton: $p = Z = 30$.\n- Số electron: $e = p = 30$.\n- Số neutron: $n = A - Z = 65 - 30 = 35$.\nTổng số hạt trong nguyên tử Zn là: $\\sum = p + n + e = 30 + 35 + 30 = 95$ hạt."
    },
    {
        "id": 2,
        "content": "Potassium (K) là kim loại có tính khử mạnh, phân mức năng lượng cao nhất là 4s^1. Tổng số electron lớp thứ 3 của K bằng bao nhiêu?",
        "answer": "8",
        "explanation": "Phân mức năng lượng cao nhất là $4s^1$, do đó thứ tự sắp xếp electron theo phân mức năng lượng là $1s^2 2s^2 2p^6 3s^2 3p^6 4s^1$.\nCấu hình electron nguyên tử K ($Z = 19$): $1s^2 2s^2 2p^6 3s^2 3p^6 4s^1$.\nLớp thứ 3 ($n = 3$, lớp M) gồm các phân lớp $3s$ và $3p$, có tổng số electron là $2 + 6 = 8$ electron."
    },
    {
        "id": 3,
        "content": "Nguyên tố X thuộc chu kì 2. Vậy X có bao nhiêu lớp electron?",
        "answer": "2",
        "explanation": "Theo quy tắc cấu tạo bảng tuần hoàn: Số thứ tự của chu kì bằng số lớp electron trong nguyên tử của các nguyên tố thuộc chu kì đó. Do đó, nguyên tố X thuộc chu kì 2 thì nguyên tử X có 2 lớp electron."
    },
    {
        "id": 4,
        "content": "Trong phân tử MX~2~ có tổng số hạt là 164, trong phân tử số hạt mang điện nhiều hơn số hạt không mang điện là 52 hạt. Số khối của nguyên tử M lớn hơn số khối của nguyên tử X là 5. Tổng số hạt trong 1 nguyên tử M nhiều hơn trong 1 nguyên tử X là 8 hạt. Tính số khối của M.",
        "answer": "40",
        "explanation": "Gọi số proton, neutron của M là $p_M, n_M$; của X là $p_X, n_X$.\nTrong phân tử $MX_2$:\n- Tổng số hạt: $(2p_M + n_M) + 2(2p_X + n_X) = 164$  (1)\n- Số hạt mang điện nhiều hơn số hạt không mang điện là 52:\n  $(2p_M + 4p_X) - (n_M + 2n_X) = 52$  (2)\nCộng (1) và (2):\n  $2(2p_M + 4p_X) = 216 \\Rightarrow 2p_M + 4p_X = 108 \\Rightarrow p_M + 2p_X = 54$\nThay vào (1): $n_M + 2n_X = 164 - 108 = 56$.\nTổng số khối của phân tử $MX_2$ là: $A_M + 2A_X = (p_M + n_M) + 2(p_X + n_X) = 54 + 56 = 110$.\nMặt khác, số khối của M lớn hơn của X là 5: $A_M - A_X = 5 \\Rightarrow A_X = A_M - 5$.\nThay vào: $A_M + 2(A_M - 5) = 110 \\Leftrightarrow 3A_M - 10 = 110 \\Leftrightarrow 3A_M = 120 \\Rightarrow A_M = 40$.\n(Vậy M là Canxi - Ca, $A_M = 40$)."
    },
    {
        "id": 5,
        "content": "Carbon là phi kim có ứng dụng phổ biến trong đời sống như nhiên liệu, làm mặt nạ phòng độc, trang sức. Nguyên tử carbon có hai đồng vị bền: ^12_6C và ^13_6C. Nguyên tử khối trung bình của carbon là 12,01. Thành phần phần trăm số nguyên tử của đồng vị ^13_6C là bao nhiêu?",
        "answer": "1",
        "explanation": "Gọi phần trăm số nguyên tử của đồng vị $^{13}_6C$ là $x\\%$, khi đó phần trăm của đồng vị $^{12}_6C$ là $(100 - x)\\%$.\nÁp dụng công thức tính nguyên tử khối trung bình:\n$$\\bar{A} = \\frac{13 \\cdot x + 12 \\cdot (100 - x)}{100} = 12,01$$\n$$\\Leftrightarrow 13x + 1200 - 12x = 1201 \\Leftrightarrow x + 1200 = 1201 \\Rightarrow x = 1\\%.$$\nVậy phần trăm số nguyên tử của đồng vị $^{13}_6C$ là 1%."
    },
    {
        "id": 6,
        "content": "Kim loại X có ứng dụng lớn trong đời sống như làm vật liệu xây dựng, xoong, chảo. Cho ion XO~2~^-^ với O là nguyên tử oxygen (biết O có 8 proton). Tổng số electron trong ion trên bằng 30. Tổng số hạt mang điện trong nguyên tử X là bao nhiêu?",
        "answer": "26",
        "explanation": "Ion $XO_2^-$ được hình thành khi phân tử $XO_2$ nhận thêm 1 electron.\nTổng số electron trong ion $XO_2^-$ là:\n$$e_{XO_2^-} = Z_X + 2 \\cdot Z_O + 1 = 30$$\nVới $Z_O = 8$, ta có:\n$$Z_X + 2 \\cdot 8 + 1 = 30 \\Rightarrow Z_X + 17 = 30 \\Rightarrow Z_X = 13$$\nNguyên tố X có $Z = 13$ (Aluminium - Al).\nTrong nguyên tử X, các hạt mang điện gồm proton và electron:\nTổng số hạt mang điện $= 2 \\cdot Z_X = 2 \\cdot 13 = 26$ hạt."
    },
    {
        "id": 7,
        "content": "Nguyên tử X có điện tích hạt nhân bằng +9. Số electron độc thân của X ở trạng thái cơ bản là bao nhiêu?",
        "answer": "1",
        "explanation": "Điện tích hạt nhân bằng $+9 \\Rightarrow$ Số hiệu nguyên tử $Z = 9$ (nguyên tố Fluorine - F).\nCấu hình electron của X: $1s^2 2s^2 2p^5$.\n- Lớp $1s$ có 1 cặp electron ghép đôi: [$\\uparrow\\downarrow$].\n- Phân lớp $2s$ có 1 cặp electron ghép đôi: [$\\uparrow\\downarrow$].\n- Phân lớp $2p$ có 3 orbital chứa 5 electron: [$\\uparrow\\downarrow$] [$\\uparrow\\downarrow$] [$\\uparrow$] (gồm 2 cặp ghép đôi và 1 electron độc thân).\nVậy số electron độc thân của X ở trạng thái cơ bản là 1."
    },
    {
        "id": 8,
        "content": "Nguyên tử Y có phân lớp năng lượng cao nhất là 3p^3. Tổng số electron của Y là bao nhiêu?",
        "answer": "15",
        "explanation": "Phân mức năng lượng cao nhất là $3p^3$, do đó cấu hình electron đầy đủ của Y theo mức năng lượng tăng dần là: $1s^2 2s^2 2p^6 3s^2 3p^3$.\nTổng số electron của nguyên tử Y là:\n$$\\sum e = 2 + 2 + 6 + 2 + 3 = 15$$\n(Y là nguyên tố Phosphorus - P)."
    }
]

# -----------------------------------------------------------------------------
# PHẦN IV: 3 CÂU TỰ LUẬN (3,0 ĐIỂM)
# -----------------------------------------------------------------------------
PART_4_QUESTIONS = [
    {
        "id": 1,
        "points": "1,0 điểm",
        "content": "Cho nguyên tố hydrogen có ba đồng vị ^1_1H, ^2_1H, ^3_1H.",
        "sub_questions": [
            {
                "sub_id": "a",
                "points": "0,25 điểm",
                "text": "Xác định số proton, neutron và electron của đồng vị ^1_1H.",
                "solution": "Đồng vị $^1_1H$ có số hiệu nguyên tử $Z = 1$ và số khối $A = 1$:\n- Số proton: $p = Z = 1$.\n- Số electron: $e = p = 1$.\n- Số neutron: $n = A - Z = 1 - 1 = 0$."
            },
            {
                "sub_id": "b",
                "points": "0,25 điểm",
                "text": "Dựa vào đâu người ta ghi khối lượng nguyên tố hydrogen vào trong bảng tuần hoàn?",
                "solution": "Khối lượng của nguyên tố hydrogen được ghi trong bảng tuần hoàn là *nguyên tử khối trung bình* của nguyên tố đó trong tự nhiên. Giá trị này được xác định dựa trên số khối của các đồng vị bền trong tự nhiên và tỉ lệ phần trăm số nguyên tử tương ứng của từng đồng vị đó theo công thức:\n$$\\bar{A} = \\frac{A_1 \\cdot x_1 + A_2 \\cdot x_2 + ...}{100}$$"
            },
            {
                "sub_id": "c",
                "points": "0,25 điểm",
                "text": "Phân tử nước có công thức là H~2~O. Cho biết nguyên tố oxygen có 3 đồng vị bền ^16_8O, ^17_8O, ^18_8O. Em hãy cho biết phân tử nước có khối lượng nặng nhất là bao nhiêu amu (chấp nhận khối lượng nguyên tử bằng số khối)?",
                "solution": "Phân tử nước ($H_2O$) gồm 2 nguyên tử H và 1 nguyên tử O.\nĐể phân tử nước có khối lượng nặng nhất thì nó phải được tạo thành từ các đồng vị có số khối lớn nhất của H và O:\n- Đồng vị H nặng nhất là $^3_1H$ (số khối $A = 3$).\n- Đồng vị O nặng nhất là $^{18}_8O$ (số khối $A = 18$).\nKhối lượng nặng nhất của phân tử nước là:\n$$M_{max} = 2 \\times 3 + 18 = 24\\text{ amu.}$$"
            },
            {
                "sub_id": "d",
                "points": "0,25 điểm",
                "text": "Có bao nhiêu phân tử nước được hình thành từ đồng vị ^16_8O?",
                "solution": "Phân tử nước có dạng $H_{(A)} - O - H_{(B)}$. Khi cố định nguyên tử O là đồng vị $^{16}_8O$, phân tử nước được tạo nên từ việc kết hợp 2 nguyên tử H lấy từ 3 đồng vị ($^1H, ^2H, ^3H$):\n- Trường hợp 2 nguyên tử H cùng đồng vị: $^1H - ^1H - ^{16}O$, $^2H - ^2H - ^{16}O$, $^3H - ^3H - ^{16}O$ (có 3 phân tử).\n- Trường hợp 2 nguyên tử H khác đồng vị: $^1H - ^2H - ^{16}O$, $^1H - ^3H - ^{16}O$, $^2H - ^3H - ^{16}O$ (có 3 phân tử).\nTổng số phân tử nước hình thành từ đồng vị $^{16}_8O$ là: $3 + 3 = 6$ phân tử."
            }
        ]
    },
    {
        "id": 2,
        "points": "1,0 điểm",
        "content": "Sodium (Na) là nguyên tố quan trọng, có nhiều trong nước biển ở dạng muối NaCl. Na có số hiệu nguyên tử là 11.",
        "sub_questions": [
            {
                "sub_id": "a",
                "points": "0,25 điểm",
                "text": "Xác định số electron của Na.",
                "solution": "Vì nguyên tử trung hòa về điện nên số electron bằng số hiệu nguyên tử:\n$$e = Z = 11\\text{ electron.}$$"
            },
            {
                "sub_id": "b",
                "points": "0,25 điểm",
                "text": "Viết cấu hình electron nguyên tử của Na.",
                "solution": "Với $Z = 11$, cấu hình electron nguyên tử của Na là:\n$$1s^2 2s^2 2p^6 3s^1\\quad\\text{hoặc viết gọn: } [Ne]3s^1$$"
            },
            {
                "sub_id": "c",
                "points": "0,25 điểm",
                "text": "Na là kim loại, phi kim hay khí hiếm?",
                "solution": "Nguyên tử Na có 1 electron ở lớp ngoài cùng ($3s^1$) và thuộc chu kì 3, nhóm IA. Do đó Na là một *kim loại* (kim loại kiềm điển hình), có tính khử mạnh và dễ nhường 1 electron để đạt cấu hình bền vững của khí hiếm neon."
            },
            {
                "sub_id": "d",
                "points": "0,25 điểm",
                "text": "Cho nguyên tử khối trung bình của Na và Cl lần lượt là 22,99 và 35,5. Na có hai đồng vị ^23_11Na, ^22_11Na. Tính phần trăm khối lượng của ^22_11Na trong phân tử NaCl.",
                "solution": "1. Tìm phần trăm số nguyên tử của đồng vị $^{22}_{11}Na$:\nGọi $x$ là phần trăm số nguyên tử của $^{22}Na$, phần trăm của $^{23}Na$ là $(100 - x)\\%$.\nTa có: $\\bar{A}_{Na} = \\frac{22 \\cdot x + 23 \\cdot (100 - x)}{100} = 22,99$\n$\\Leftrightarrow 2300 - x = 2299 \\Rightarrow x = 1\\%$.\n\n2. Tính phần trăm khối lượng của $^{22}_{11}Na$ trong phân tử NaCl:\nKhối lượng phân tử của NaCl là: $M_{NaCl} = \\bar{A}_{Na} + \\bar{A}_{Cl} = 22,99 + 35,5 = 58,49$.\nPhần trăm khối lượng của $^{22}_{11}Na$ trong phân tử NaCl là:\n$$\\% m_{^{22}Na} = \\frac{22 \\times 1\\%}{58,49} \\times 100\\% = \\frac{0,22}{58,49} \\times 100\\% \\approx 0,376\\%.$$"
            }
        ]
    },
    {
        "id": 3,
        "points": "1,0 điểm",
        "content": "Cho nguyên tố X ở chu kì 2 và có 4 electron lớp ngoài cùng.",
        "sub_questions": [
            {
                "sub_id": "a",
                "points": "0,25 điểm",
                "text": "X có bao nhiêu lớp electron?",
                "solution": "Số thứ tự chu kì bằng số lớp electron trong nguyên tử. Vì nguyên tố X ở chu kì 2 nên nguyên tử X có *2 lớp electron*."
            },
            {
                "sub_id": "b",
                "points": "0,25 điểm",
                "text": "X ở nhóm nào trong bảng hệ thống tuần hoàn?",
                "solution": "Nguyên tử X có 2 lớp electron và 4 electron lớp ngoài cùng, cấu hình electron là $1s^2 2s^2 2p^2$.\nElectron cuối cùng điền vào phân lớp p $\\Rightarrow$ X là nguyên tố p.\nSố electron hoá trị bằng số electron lớp ngoài cùng $= 4$.\nDo đó nguyên tố X thuộc *nhóm IVA* trong bảng tuần hoàn (X là Carbon - C)."
            },
            {
                "sub_id": "c",
                "points": "0,25 điểm",
                "text": "Nguyên tố Y ở cùng chu kì với X và có nhiều hơn X hai đơn vị điện tích hạt nhân. Xác định số electron của Y.",
                "solution": "Số hiệu nguyên tử của X là: $Z_X = 2 + 4 = 6$.\nVì nguyên tố Y ở cùng chu kì 2 và có nhiều hơn X hai đơn vị điện tích hạt nhân nên:\n$$Z_Y = Z_X + 2 = 6 + 2 = 8$$\nSố electron của nguyên tử Y bằng số hiệu nguyên tử:\n$$e_Y = Z_Y = 8\\text{ electron (Y là Oxygen - O).}$$"
            },
            {
                "sub_id": "d",
                "points": "0,25 điểm",
                "text": "Hợp chất của X và Y có công thức XY~2~. Tính tổng số proton trong phân tử XY~2~.",
                "solution": "Với $Z_X = 6$ (C) và $Z_Y = 8$ (O), hợp chất $XY_2$ chính là khí carbon dioxide ($CO_2$).\nTổng số proton trong một phân tử $XY_2$ là:\n$$\\sum p = p_X + 2 \\cdot p_Y = 6 + 2 \\times 8 = 22\\text{ proton.}$$"
            }
        ]
    }
]
