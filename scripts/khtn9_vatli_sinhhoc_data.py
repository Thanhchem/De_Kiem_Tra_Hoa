# -*- coding: utf-8 -*-
"""
Dữ liệu Đề cương ôn tập KHTN 9 (Khoa học tự nhiên 9) - Học kì I
Phân môn Vật lí & Phân môn Sinh học (KHTN 3)
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

INFO = {
    "school": "TRƯỜNG THCS KIM LONG",
    "year": "NĂM HỌC 2026 - 2027",
    "title_left_1": "TRƯỜNG THCS KIM LONG",
    "title_left_2": "TỔ KHOA HỌC TỰ NHIÊN",
    "sub_left": "(Tài liệu lưu hành nội bộ)",
    "title_right_1": "ĐỀ CƯƠNG ÔN TẬP GIỮA KÌ I - KHTN 9",
    "title_right_2": "PHÂN MÔN VẬT LÍ & PHÂN MÔN SINH HỌC",
    "time_str": "Chương trình GDPT 2018 - Đầy đủ Trắc nghiệm & Tự luận có lời giải chi tiết",
    "made": "091",
}

# =============================================================================
# PHÂN MÔN VẬT LÍ (CƠ NĂNG - CÔNG VÀ CÔNG SUẤT - NĂNG LƯỢNG HÓA THẠCH & TÁI TẠO)
# =============================================================================

# 1. Trắc nghiệm nhiều lựa chọn (13 câu)
PHYSICS_CHOICE_QUESTIONS = [
    {
        "id": 1,
        "content": "Khi tốc độ của một vật tăng lên 2 lần thì động năng của vật sẽ",
        "options": [
            "A. tăng gấp 2.",
            "B. tăng gấp 4.",
            "C. tăng gấp 8.",
            "D. giảm 2 lần."
        ],
        "correct": "B",
        "solution": "Động năng của một vật được tính theo công thức: W~đ~ = (1/2)·m·v^2^.\nVì động năng tỉ lệ thuận với bình phương tốc độ (v^2^), nên khi tốc độ tăng lên 2 lần (v' = 2v), động năng mới của vật là:\nW~đ~' = (1/2)·m·(2v)^2^ = 4·[(1/2)·m·v^2^] = 4·W~đ~ (tăng gấp 4 lần)."
    },
    {
        "id": 2,
        "content": "Thế năng trọng trường của vật phụ thuộc vào",
        "options": [
            "A. khối lượng.",
            "B. tốc độ.",
            "C. thể tích.",
            "D. hình dạng."
        ],
        "correct": "A",
        "solution": "Thế năng trọng trường của vật ở độ cao h so với mốc thế năng được tính bởi: W~t~ = P·h = m·g·h.\nDo đó, thế năng trọng trường phụ thuộc vào khối lượng (m) của vật và độ cao (h) của vật so với mốc chọn tính thế năng. Trong các phương án đã cho, thế năng trọng trường phụ thuộc vào khối lượng."
    },
    {
        "id": 3,
        "content": "Khi nước chảy từ trên đập cao xuống tuabin của máy phát điện thì có sự chuyển hóa năng lượng từ",
        "options": [
            "A. động năng thành thế năng.",
            "B. động năng thành cơ năng.",
            "C. thế năng thành động năng.",
            "D. thế năng thành cơ năng."
        ],
        "correct": "C",
        "solution": "Nước ở trên đập cao có độ cao lớn so với chân đập nên tích trữ thế năng trọng trường rất lớn. Khi nước chảy từ trên đập xuống, độ cao giảm dần trong khi tốc độ dòng nước tăng dần, do đó thế năng của nước chuyển hóa thành động năng của dòng nước (sau đó động năng của dòng nước làm quay tuabin máy phát điện)."
    },
    {
        "id": 4,
        "content": "Một chiếc ô tô đang chạy thì tắt máy, xe đi thêm được một quãng đường nữa rồi dừng lại. Trong quá trình này có sự chuyển hóa năng lượng từ",
        "options": [
            "A. thế năng thành nhiệt năng.",
            "B. động năng thành nhiệt năng.",
            "C. cơ năng thành nhiệt năng.",
            "D. động năng thành thế năng."
        ],
        "correct": "B",
        "solution": "Khi ô tô đang chuyển động trên mặt đường nằm ngang, ô tô sở hữu động năng. Sau khi tắt máy, do có lực ma sát giữa bánh xe với mặt đường và lực cản không khí, xe chạy chậm dần rồi dừng hẳn. Toàn bộ động năng của ô tô đã chuyển hóa thành nhiệt năng làm nóng lốp xe, mặt đường và không khí xung quanh."
    },
    {
        "id": 5,
        "content": "Cơ năng của một vật là",
        "options": [
            "A. chỉ có động năng.",
            "B. chỉ có thế năng trọng trường.",
            "C. tổng động năng và thế năng.",
            "D. hiệu động năng và thế năng."
        ],
        "correct": "C",
        "solution": "Khi một vật vừa có khả năng sinh công do chuyển động (động năng), vừa có khả năng sinh công do vị trí hoặc độ biến dạng (thế năng), thì tổng động năng và thế năng của vật được gọi là cơ năng: W = W~đ~ + W~t~."
    },
    {
        "id": 6,
        "content": "Đơn vị nào sau đây là đơn vị đo công suất?",
        "options": [
            "A. mêgaoát (MW).",
            "B. paxcan (Pa).",
            "C. BTU.",
            "D. kilocalo (kcal)."
        ],
        "correct": "A",
        "solution": "- Mêgaoát (MW) là bội số của Oát (W), đơn vị đo công suất chuẩn trong hệ SI (1 MW = 1 000 000 W).\n- Paxcan (Pa) là đơn vị đo áp suất (1 Pa = 1 N/m^2^).\n- BTU và kilocalo (kcal) là các đơn vị đo nhiệt lượng, năng lượng."
    },
    {
        "id": 7,
        "content": "Một người kéo một chiếc hộp với lực 50 N trên quãng đường 10 m theo hướng của lực. Công mà người đó thực hiện là bao nhiêu?",
        "options": [
            "A. 500 J.",
            "B. 50 J.",
            "C. 10 J.",
            "D. 5 J."
        ],
        "correct": "A",
        "solution": "Công của lực kéo thực hiện khi vật dịch chuyển cùng hướng với lực là:\nA = F·s = 50 N × 10 m = 500 J."
    },
    {
        "id": 8,
        "content": "Một chiếc máy nâng một vật có khối lượng 100 kg lên độ cao 5 m trong thời gian 10 giây. Công suất của máy là bao nhiêu?",
        "options": [
            "A. 50 W.",
            "B. 500 W.",
            "C. 1000 W.",
            "D. 1500 W."
        ],
        "correct": "B",
        "solution": "- Trọng lượng của vật: P = 10·m = 10 × 100 = 1000 N.\n- Công cần thiết để nâng vật lên cao 5 m: A = P·h = 1000 × 5 = 5000 J.\n- Công suất của máy nâng là: P = A / t = 5000 J / 10 s = 500 W."
    },
    {
        "id": 9,
        "content": "Để đánh giá xem ai làm việc khoẻ hơn, người ta cần biết",
        "options": [
            "A. ai thực hiện công lớn hơn.",
            "B. ai dùng ít thời gian hơn.",
            "C. ai dùng lực mạnh hơn.",
            "D. trong cùng một thời gian ai thực hiện công lớn hơn."
        ],
        "correct": "D",
        "solution": "Để so sánh khả năng sinh công (ai làm việc khỏe hơn, mạnh hơn), người ta so sánh công suất. Công suất là công thực hiện được trong một đơn vị thời gian (1 giây). Do đó, người ta cần so sánh xem trong cùng một khoảng thời gian ai thực hiện được công lớn hơn."
    },
    {
        "id": 10,
        "content": "Việc đốt cháy nhiên liệu hóa thạch sẽ gây ảnh hưởng đến sức khỏe con người và gây ô nhiễm môi trường. Sản phẩm đốt cháy nào sau đây gây ra hiệu ứng nhà kính trên Trái Đất?",
        "options": [
            "A. H~2~O.",
            "B. CO~2~.",
            "C. CO.",
            "D. B và C đều đúng."
        ],
        "correct": "B",
        "solution": "Khí carbon dioxide (CO~2~) tạo ra khi đốt cháy than đá, dầu mỏ và khí đốt là chất khí nhà kính chủ yếu giữ lại bức xạ nhiệt trong bầu khí quyển, làm tăng nhiệt độ bề mặt Trái Đất và gây biến đổi khí hậu toàn cầu."
    },
    {
        "id": 11,
        "content": "Năng lượng nào dưới đây không thể tái tạo trong thời gian ngắn?",
        "options": [
            "A. Mặt Trời.",
            "B. Địa nhiệt.",
            "C. Gió.",
            "D. Dầu mỏ."
        ],
        "correct": "D",
        "solution": "Dầu mỏ là nhiên liệu hóa thạch, phải mất hàng trăm triệu năm tích tụ và biến đổi địa chất mới hình thành nên được xếp vào nhóm nguồn năng lượng không tái tạo, không thể phục hồi trong khoảng thời gian ngắn của đời sống con người."
    },
    {
        "id": 12,
        "content": "Nhiên liệu hóa thạch có nguồn gốc từ việc phân hủy xác",
        "options": [
            "A. động vật.",
            "B. thực vật.",
            "C. động vật và thực vật.",
            "D. vi khuẩn."
        ],
        "correct": "C",
        "solution": "Nhiên liệu hóa thạch (than đá, dầu mỏ, khí thiên nhiên) có nguồn gốc từ quá trình phân hủy kị khí xác sinh vật cổ đại (bao gồm cả thực vật cổ đại và động vật phù du biển) bị vùi lấp sâu trong trầm tích qua hàng trăm triệu năm dưới áp suất và nhiệt độ cao."
    },
    {
        "id": 13,
        "content": "Gió và dòng hải lưu được tạo thành từ",
        "options": [
            "A. sự chênh lệch nhiệt độ giữa các vùng trên Trái Đất.",
            "B. do lực hút của Mặt Trăng.",
            "C. do lực hút của Mặt Trời.",
            "D. do năng lượng địa nhiệt của Trái Đất."
        ],
        "correct": "A",
        "solution": "Ánh sáng Mặt Trời hun nóng bề mặt Trái Đất không đồng đều tạo nên sự chênh lệch nhiệt độ và khí áp giữa các vùng, khiến không khí di chuyển từ nơi áp cao về nơi áp thấp sinh ra gió. Gió và sự phân bố nhiệt độ không đồng đều trong nước biển lại tiếp tục thúc đẩy chuyển động tuần hoàn của các dòng hải lưu."
    }
]

# 2. Trắc nghiệm đúng sai (4 câu)
PHYSICS_TF_QUESTIONS = [
    {
        "id": 14,
        "content": "Một hòn đá được thả rơi từ độ cao 5 m so với mặt đất. Bỏ qua lực cản của không khí.",
        "statements": [
            {
                "id": "a",
                "text": "Trong khi rơi, thế năng trọng trường được chuyển hóa thành động năng.",
                "correct": True,
                "explain": "Đúng. Trong quá trình rơi, độ cao h giảm dần nên thế năng giảm, vận tốc v tăng dần nên động năng tăng; thế năng chuyển hóa thành động năng."
            },
            {
                "id": "b",
                "text": "Cơ năng của hòn đá được bảo toàn.",
                "correct": True,
                "explain": "Đúng. Bỏ qua lực cản không khí, hòn đá chỉ chịu tác dụng của trọng lực nên cơ năng của vật được bảo toàn."
            },
            {
                "id": "c",
                "text": "Trong khi rơi, thế năng tăng và động năng tăng.",
                "correct": False,
                "explain": "Sai. Vì độ cao giảm nên thế năng phải giảm, không thể cả thế năng và động năng cùng tăng."
            },
            {
                "id": "d",
                "text": "Trong khi rơi, tốc độ của hòn đá giảm dần.",
                "correct": False,
                "explain": "Sai. Dưới tác dụng của trọng lực, chuyển động rơi tự do là chuyển động nhanh dần đều nên tốc độ của hòn đá tăng dần."
            }
        ],
        "explanation": "• Ý a ĐÚNG: Độ cao giảm dần làm thế năng giảm, tốc độ tăng dần làm động năng tăng.\n• Ý b ĐÚNG: Cơ năng được bảo toàn khi vật chuyển động chỉ dưới tác dụng của trọng lực.\n• Ý c SAI: Thế năng giảm chứ không tăng.\n• Ý d SAI: Tốc độ của hòn đá tăng dần cho đến khi chạm đất."
    },
    {
        "id": 15,
        "content": "Một vật có khối lượng m = 1,5 kg được thả rơi từ độ cao h = 4 m so với mặt đất. Chọn gốc thế năng ở mặt đất. Biết toàn bộ thế năng của vật được chuyển hóa thành động năng (lấy g = 10 m/s^2^).",
        "statements": [
            {
                "id": "a",
                "text": "Trong quá trình rơi, thế năng giảm, động năng tăng.",
                "correct": True,
                "explain": "Đúng. Khi vật rơi xuống, độ cao h giảm nên thế năng giảm, vận tốc v tăng nên động năng tăng."
            },
            {
                "id": "b",
                "text": "Cơ năng của vật chỉ có thế năng trọng trường.",
                "correct": False,
                "explain": "Sai. Trong quá trình rơi, cơ năng của vật là tổng của cả thế năng trọng trường và động năng."
            },
            {
                "id": "c",
                "text": "Động năng của vật ngay trước khi chạm đất là W~đ~ = 60 (J).",
                "correct": True,
                "explain": "Đúng. Khi chạm đất, toàn bộ thế năng chuyển hóa thành động năng: W~đ~ = W~t~ = m·g·h = 1,5 × 10 × 4 = 60 J."
            },
            {
                "id": "d",
                "text": "Tốc độ của vật ngay trước khi chạm đất là v = 4√5 m/s.",
                "correct": True,
                "explain": "Đúng. Từ (1/2)·m·v^2^ = 60 => v^2^ = (2 × 60) / 1,5 = 80 => v = √80 = 4√5 m/s."
            }
        ],
        "explanation": "• Ý a ĐÚNG: Độ cao giảm nên thế năng giảm; tốc độ tăng nên động năng tăng.\n• Ý b SAI: Cơ năng gồm cả động năng và thế năng: W = W~đ~ + W~t~.\n• Ý c ĐÚNG: W~đ~ chạm đất = W~t~ ban đầu = 1,5 × 10 × 4 = 60 J.\n• Ý d ĐÚNG: v = √(2W~đ~ / m) = √(120 / 1,5) = √80 = 4√5 m/s."
    },
    {
        "id": 16,
        "content": "Một người kéo đều một vật nặng 18 (kg) từ giếng sâu 8 (m) lên trong 20 giây (lấy g = 10 m/s^2^).",
        "statements": [
            {
                "id": "a",
                "text": "Lực kéo tối thiểu bằng 180 (N).",
                "correct": True,
                "explain": "Đúng. Để kéo đều vật lên thì lực kéo tối thiểu phải cân bằng với trọng lượng: F = P = 10·m = 10 × 18 = 180 N."
            },
            {
                "id": "b",
                "text": "Công của người đó bằng 1440 (J).",
                "correct": True,
                "explain": "Đúng. Công của lực kéo là A = F·s = 180 × 8 = 1440 J."
            },
            {
                "id": "c",
                "text": "Công suất của người đó bằng 72 (W).",
                "correct": True,
                "explain": "Đúng. Công suất kéo là P = A / t = 1440 / 20 = 72 W."
            },
            {
                "id": "d",
                "text": "Tốc độ chuyển động của vật là 0,4 (m/s).",
                "correct": True,
                "explain": "Đúng. Tốc độ kéo đều là v = s / t = 8 / 20 = 0,4 m/s."
            }
        ],
        "explanation": "• Ý a ĐÚNG: Lực kéo tối thiểu F = P = 10 × 18 = 180 N.\n• Ý b ĐÚNG: Công thực hiện A = F·s = 180 × 8 = 1440 J.\n• Ý c ĐÚNG: Công suất P = A / t = 1440 / 20 = 72 W.\n• Ý d ĐÚNG: Tốc độ chuyển động v = s / t = 8 / 20 = 0,4 m/s."
    },
    {
        "id": 17,
        "content": "Xét tính đúng – sai của các phát biểu sau về năng lượng và biến đổi năng lượng:",
        "statements": [
            {
                "id": "a",
                "text": "Quạt điện có hiệu suất 55% hoạt động hiệu quả hơn quạt điện có hiệu suất 45%.",
                "correct": True,
                "explain": "Đúng. Hiệu suất cao hơn nghĩa là tỉ lệ điện năng biến đổi thành cơ năng có ích lớn hơn, quạt ít bị nóng hao phí hơn nên hoạt động hiệu quả hơn."
            },
            {
                "id": "b",
                "text": "Ưu điểm của việc khai thác và sử dụng năng lượng tái tạo là không phụ thuộc vào các yếu tố thời tiết và thiên nhiên; chi phí đầu tư ban đầu rẻ.",
                "correct": False,
                "explain": "Sai. Năng lượng tái tạo (mặt trời, gió) phụ thuộc chặt chẽ vào thời tiết và chi phí đầu tư ban đầu thường khá tốn kém."
            },
            {
                "id": "c",
                "text": "Nhiên liệu hydrogen được coi là một nguồn năng lượng xanh vì không phát thải khí nhà kính trong quá trình đốt cháy.",
                "correct": True,
                "explain": "Đúng. Khi đốt cháy hydrogen chỉ tạo thành hơi nước (2H~2~ + O~2~ → 2H~2~O), hoàn toàn không phát thải khí CO~2~ gây hiệu ứng nhà kính."
            },
            {
                "id": "d",
                "text": "Thực vật hấp thụ năng lượng Mặt Trời để chuyển khí CO~2~ và nước thành đường và tinh bột nhờ vào quá trình quang hợp.",
                "correct": True,
                "explain": "Đúng. Quang hợp là quá trình chuyển hóa quang năng Mặt Trời thành hóa năng tích trữ trong các chất hữu cơ (đường glucose, tinh bột)."
            }
        ],
        "explanation": "• Ý a ĐÚNG: Hiệu suất cao hơn đồng nghĩa với việc ít tiêu hao điện năng vô ích hơn.\n• Ý b SAI: Năng lượng tái tạo phụ thuộc lớn vào thiên nhiên và chi phí lắp đặt ban đầu cao.\n• Ý c ĐÚNG: Đốt cháy hydrogen chỉ sinh ra nước H~2~O, không tạo khí CO~2~.\n• Ý d ĐÚNG: Quang hợp biến đổi CO~2~ và H~2~O thành chất hữu cơ nhờ năng lượng ánh sáng."
    }
]

# 3. Trắc nghiệm trả lời ngắn (5 câu)
PHYSICS_SHORT_QUESTIONS = [
    {
        "id": 18,
        "content": "Một người thực hiện công 600 J trong thời gian 20 s. Công suất của người đó là bao nhiêu W?",
        "correct": "30",
        "explanation": "Công suất của người đó được tính theo công thức:\nP = A / t = 600 J / 20 s = 30 W.\nĐáp án: 30."
    },
    {
        "id": 19,
        "content": "Một lực kéo vật chuyển động cùng hướng với lực, có độ lớn 20 N. Vật dịch chuyển được 3 m. Công của lực kéo là bao nhiêu J?",
        "correct": "60",
        "explanation": "Công của lực kéo thực hiện là:\nA = F·s = 20 N × 3 m = 60 J.\nĐáp án: 60."
    },
    {
        "id": 20,
        "content": "Nếu giảm khối lượng của một vật xuống 2 lần và tăng độ cao của vật đó lên 4 lần thì thế năng trọng trường của nó tăng bao nhiêu lần?",
        "correct": "2",
        "explanation": "Thế năng trọng trường ban đầu: W~t~ = m·g·h.\nKhi m' = m / 2 và h' = 4h:\nW~t~' = m'·g·h' = (m / 2)·g·(4h) = 2·(m·g·h) = 2·W~t~.\n=> Thế năng trọng trường của vật tăng 2 lần.\nĐáp án: 2."
    },
    {
        "id": 21,
        "content": "Một vật đang chuyển động với tốc độ không đổi là v (m/s) thì có động năng là 15 J. Nếu tốc độ của vật tăng lên đến 4v thì động năng của vật bằng bao nhiêu J?",
        "correct": "240",
        "explanation": "Động năng ban đầu: W~đ~ = (1/2)·m·v^2^ = 15 J.\nKhi tốc độ tăng lên 4v, động năng mới của vật là:\nW~đ~' = (1/2)·m·(4v)^2^ = 16·[(1/2)·m·v^2^] = 16 × 15 = 240 J.\nĐáp án: 240."
    },
    {
        "id": 22,
        "content": "Một vật có khối lượng 3 kg ở độ cao 4 m so với mặt đất (lấy g = 10 m/s^2^). Hỏi thế năng trọng trường của vật là bao nhiêu J?",
        "correct": "120",
        "explanation": "Thế năng trọng trường của vật so với mặt đất là:\nW~t~ = m·g·h = 3 kg × 10 m/s^2^ × 4 m = 120 J.\nĐáp án: 120."
    }
]

# 4. Tự luận Vật lí (12 câu)
PHYSICS_ESSAY_QUESTIONS = [
    {
        "id": 1,
        "content": "Một ô tô có khối lượng 1,2 tấn đang chuyển động với tốc độ 72 km/h. Tính động năng của ô tô?",
        "solution": "- Đổi đơn vị chuẩn:\n  + Khối lượng: m = 1,2 tấn = 1 200 kg.\n  + Tốc độ: v = 72 km/h = 72 / 3,6 = 20 m/s.\n- Động năng của ô tô là:\n  W~đ~ = (1/2)·m·v^2^ = (1/2) × 1 200 × 20^2^ = 600 × 400 = 240 000 J = 240 kJ.\n*Đáp số:* 240 000 J (hoặc 240 kJ)."
    },
    {
        "id": 2,
        "content": "Thả vật nặng 100 g rơi tự do từ độ cao 45 m so với mặt đất. Bỏ qua sức cản của không khí và chọn mốc tính thế năng ở mặt đất. Biết rằng thế năng chuyển hóa hoàn toàn thành động năng (lấy g = 10 m/s^2^).\na) Tính động năng của vật ngay trước khi chạm đất.\nb) Tính tốc độ của vật ngay trước khi chạm đất.",
        "solution": "- Đổi đơn vị: m = 100 g = 0,1 kg; h = 45 m; g = 10 m/s^2^.\na) Thế năng ban đầu của vật ở độ cao 45 m là:\n   W~t~ = m·g·h = 0,1 × 10 × 45 = 45 J.\n   Vì thế năng chuyển hóa hoàn toàn thành động năng ngay trước khi chạm đất nên:\n   W~đ~ = W~t~ = 45 J.\nb) Tốc độ của vật ngay trước khi chạm đất:\n   Ta có: W~đ~ = (1/2)·m·v^2^ => v^2^ = (2·W~đ~) / m = (2 × 45) / 0,1 = 900\n   => v = √900 = 30 m/s.\n*Đáp số:* a) W~đ~ = 45 J; b) v = 30 m/s."
    },
    {
        "id": 3,
        "content": "Một hòn bi có khối lượng 20 g được ném thẳng đứng lên cao với tốc độ 4 m/s từ độ cao 1,6 m so với mặt đất. Bỏ qua lực cản của không khí và chọn mốc tính thế năng ở mặt đất (lấy g = 10 m/s^2^).\na) Tính động năng, thế năng, cơ năng của hòn bi ngay tại lúc ném.\nb) Tìm độ cao lớn nhất mà viên bi lên tới.",
        "solution": "- Đổi đơn vị: m = 20 g = 0,02 kg; v~0~ = 4 m/s; h~0~ = 1,6 m; g = 10 m/s^2^.\na) Ngay tại thời điểm ném:\n   - Động năng của hòn bi:\n     W~đ0~ = (1/2)·m·v~0~^2^ = (1/2) × 0,02 × 4^2^ = 0,16 J.\n   - Thế năng của hòn bi:\n     W~t0~ = m·g·h~0~ = 0,02 × 10 × 1,6 = 0,32 J.\n   - Cơ năng của hòn bi:\n     W = W~đ0~ + W~t0~ = 0,16 + 0,32 = 0,48 J.\nb) Khi viên bi lên đến độ cao lớn nhất (h~max~), vận tốc v = 0 nên động năng bằng 0, toàn bộ cơ năng chuyển thành thế năng cực đại:\n   W = W~t max~ = m·g·h~max~\n   => h~max~ = W / (m·g) = 0,48 / (0,02 × 10) = 0,48 / 0,2 = 2,4 m.\n*Đáp số:* a) W~đ~ = 0,16 J; W~t~ = 0,32 J; W = 0,48 J; b) h~max~ = 2,4 m."
    },
    {
        "id": 4,
        "content": "Một vật được ném thẳng đứng xuống dưới từ độ cao 20 m với tốc độ ban đầu 10 m/s. Bỏ qua mọi lực cản (lấy g = 10 m/s^2^).\na) Tìm cơ năng của vật khi ném vật (tính theo khối lượng m).\nb) Tìm thế năng của vật khi vật ở vị trí mà động năng bằng 3 lần thế năng. Từ đó suy ra độ cao của vật khi đó.",
        "solution": "Gọi khối lượng của vật là m (kg). Lấy g = 10 m/s^2^.\na) Cơ năng của vật ngay tại lúc ném là:\n   W = W~đ0~ + W~t0~ = (1/2)·m·v~0~^2^ + m·g·h~0~\n   W = (1/2)·m·10^2^ + m × 10 × 20 = 50·m + 200·m = 250·m (J).\n   (Cơ năng riêng của vật là 250 J/kg).\nb) Khi vật rơi đến vị trí có động năng bằng 3 lần thế năng: W~đ~ = 3·W~t~.\n   Theo định luật bảo toàn cơ năng: W = W~đ~ + W~t~ = 3·W~t~ + W~t~ = 4·W~t~\n   => W~t~ = W / 4 = (250·m) / 4 = 62,5·m (J).\n   Mặt khác: W~t~ = m·g·h => h = W~t~ / (m·g) = (62,5·m) / (m × 10) = 6,25 m.\n*Đáp số:* a) W = 250·m (J); b) W~t~ = 62,5·m (J), độ cao h = 6,25 m."
    },
    {
        "id": 5,
        "content": "Động cơ của một thang máy tác dụng lực kéo 20 000 N để thang máy chuyển động thẳng lên trên trong 10 s và quãng đường đi được tương ứng là 18 m.\na) Tính công thực hiện của thang máy.\nb) Tính công suất trung bình của động cơ.",
        "solution": "a) Công mà động cơ thang máy đã thực hiện là:\n   A = F·s = 20 000 N × 18 m = 360 000 J = 360 kJ.\nb) Công suất trung bình của động cơ thang máy là:\n   P = A / t = 360 000 J / 10 s = 36 000 W = 36 kW.\n*Đáp số:* a) A = 360 000 J; b) P = 36 000 W (36 kW)."
    },
    {
        "id": 6,
        "content": "Một động cơ điện cung cấp công suất 5 kW cho một cần cẩu để nâng vật 1000 kg chuyển động đều lên cao 30 m (lấy g = 10 m/s^2^).\na) Tính công của động cơ thực hiện được.\nb) Tính thời gian để thực hiện công việc đó.",
        "solution": "- Đổi đơn vị: P = 5 kW = 5 000 W; m = 1 000 kg; h = 30 m; g = 10 m/s^2^.\na) Vì cần cẩu nâng đều vật nên lực nâng tối thiểu bằng trọng lượng của vật:\n   F = P = 10·m = 10 × 1 000 = 10 000 N.\n   Công thực hiện để nâng vật lên cao 30 m là:\n   A = F·h = 10 000 N × 30 m = 300 000 J = 300 kJ.\nb) Thời gian cần thiết để cần cẩu nâng vật lên độ cao đó:\n   Từ P = A / t => t = A / P = 300 000 / 5 000 = 60 s = 1 phút.\n*Đáp số:* a) A = 300 000 J; b) t = 60 s (1 phút)."
    },
    {
        "id": 7,
        "content": "Đầu tàu hỏa kéo toa xe với lực kéo F = 50 000 N làm cho toa xe đi được quãng đường s = 1 000 m trong thời gian 4 giờ.\na) Tính công của lực kéo của đầu tàu.\nb) Tính công suất của đầu tàu.",
        "solution": "- Tóm tắt dữ kiện: F = 50 000 N; s = 1 000 m; t = 4 giờ = 4 × 3 600 = 14 400 s.\na) Công của lực kéo của đầu tàu là:\n   A = F·s = 50 000 N × 1 000 m = 50 000 000 J = 50 MJ.\nb) Công suất kéo của đầu tàu là:\n   P = A / t = 50 000 000 J / 14 400 s ≈ 3 472,22 W ≈ 3,47 kW.\n*Đáp số:* a) A = 50 000 000 J (50 MJ); b) P ≈ 3 472,22 W."
    },
    {
        "id": 8,
        "content": "Một xe máy chuyển động đều, lực kéo của động cơ là 1150 (N). Trong một phút, công sinh ra là 690 (kJ). Tốc độ trung bình của xe trong một phút đó là bao nhiêu?",
        "solution": "- Đổi đơn vị: F = 1150 N; t = 1 phút = 60 s; A = 690 kJ = 690 000 J.\n- Quãng đường xe máy đi được trong 1 phút là:\n  Từ A = F·s => s = A / F = 690 000 / 1150 = 600 m.\n- Tốc độ trung bình của xe máy trong một phút đó là:\n  v = s / t = 600 m / 60 s = 10 m/s = 10 × 3,6 = 36 km/h.\n*Đáp số:* v = 10 m/s (hoặc 36 km/h)."
    },
    {
        "id": 9,
        "content": "Liệt kê một số dạng năng lượng tái tạo và năng lượng không tái tạo mà em biết.",
        "solution": "- *Nguồn năng lượng tái tạo (Renewable energy):* Là dạng năng lượng có sẵn trong tự nhiên, liên tục được tái sinh hoặc vô tận, không bị cạn kiệt theo thời gian khai thác:\n  + Năng lượng Mặt Trời (quang năng, nhiệt năng).\n  + Năng lượng gió (phong điện).\n  + Năng lượng dòng nước (thủy điện, thế năng/động năng nước).\n  + Năng lượng sóng biển và thủy triều.\n  + Năng lượng địa nhiệt (nhiệt tỏa ra từ lòng đất).\n  + Năng lượng sinh khối (biomass từ phế phẩm nông nghiệp, gỗ rừng trồng).\n- *Nguồn năng lượng không tái tạo (Non-renewable energy):* Phải mất hàng triệu năm biến đổi địa chất mới hình thành và có nguy cơ cạn kiệt hoàn toàn nếu bị khai thác quá mức:\n  + Than đá.\n  + Dầu mỏ (xăng, dầu diesel, dầu hỏa).\n  + Khí thiên nhiên (khí tự nhiên, khí mỏ dầu).\n  + Nhiên liệu hạt nhân (Uranium, Thorium)."
    },
    {
        "id": 10,
        "content": "Dầu mỏ, than đá được hình thành như thế nào? Tại sao có thể nói “Năng lượng hóa thạch cũng có nguồn gốc từ Mặt Trời”?",
        "solution": "1. *Sự hình thành than đá và dầu mỏ:*\n   - Than đá: Hình thành từ xác thực vật cổ đại (rừng cây gỗ lớn, dương xỉ cổ) bị vùi lấp sâu trong các đầm lầy trũng qua hàng trăm triệu năm. Dưới tác dụng phân hủy kị khí của vi sinh vật cùng nhiệt độ và áp suất cao trong vỏ Trái Đất, xác thực vật bị biến đổi thành than bùn, than nâu và cuối cùng là than đá.\n   - Dầu mỏ: Hình thành từ xác sinh vật phù du (tảo, vi sinh vật biển) chết đi, lắng đọng dưới đáy đại dương và bị các tầng cát bùn trầm tích vùi lấp. Trải qua hàng triệu năm chịu nhiệt độ và áp suất cực lớn, chúng bị bẻ gãy mạch phân tử biến đổi thành hỗn hợp hydrocarbon thể lỏng và khí, tích tụ trong các tầng đá xốp gọi là mỏ dầu.\n2. *Giải thích vì sao năng lượng hóa thạch có nguồn gốc từ Mặt Trời:*\n   - Thực vật và sinh vật phù du cổ đại khi còn sống đã trực tiếp hấp thụ năng lượng ánh sáng Mặt Trời để thực hiện quá trình quang hợp, tổng hợp nên các hợp chất hữu cơ giàu năng lượng dự trữ trong tế bào.\n   - Khi sinh vật bị chôn vùi và biến đổi thành than đá, dầu mỏ thì hóa năng lưu giữ trong nhiên liệu hóa thạch chính là năng lượng Mặt Trời đã được chuyển hóa và bảo quản qua hàng trăm triệu năm."
    },
    {
        "id": 11,
        "content": "Em hãy nêu một vài biện pháp để sử dụng hiệu quả và tiết kiệm năng lượng?",
        "solution": "Một số biện pháp thiết thực để sử dụng hiệu quả và tiết kiệm năng lượng trong cuộc sống:\n1. *Trong sinh hoạt gia đình:*\n   - Tắt các thiết bị điện (bóng đèn, quạt, tivi, bình nóng lạnh) khi không có nhu cầu sử dụng hoặc khi ra khỏi phòng.\n   - Sử dụng các thiết bị điện có dán nhãn chứng nhận tiết kiệm năng lượng (công nghệ Inverter, đèn LED thay thế đèn sợi đốt).\n   - Tận dụng tối đa ánh sáng mặt trời tự nhiên và mở cửa đón gió trời vào ban ngày thay vì lạm dụng bật đèn và điều hòa.\n   - Cài đặt nhiệt độ điều hòa hợp lí từ 26°C đến 28°C kết hợp quạt đảo chiều gió nhẹ.\n2. *Trong giao thông đi lại:*\n   - Ưu tiên sử dụng phương tiện giao thông công cộng (xe buýt, tàu điện), đi xe đạp hoặc đi bộ đối với các cự li ngắn.\n   - Khuyến khích sử dụng xe điện hoặc phương tiện nhiên liệu sinh học thân thiện môi trường.\n3. *Trong học tập và sản xuất:*\n   - Tuyên truyền, nâng cao ý thức tiết kiệm điện, nước trong trường học và cộng đồng.\n   - Lắp đặt hệ thống pin năng lượng mặt trời trên mái nhà để tự cung cấp một phần điện sạch."
    },
    {
        "id": 12,
        "content": "Nêu các ưu điểm và nhược điểm của năng lượng tái tạo?",
        "solution": "1. *Ưu điểm của năng lượng tái tạo:*\n   - *Vô tận và dồi dào:* Không bị cạn kiệt theo thời gian sử dụng (Mặt Trời, gió, sóng biển, dòng nước luân chuyển liên tục).\n   - *Bảo vệ môi trường:* Là nguồn năng lượng sạch, hầu như không phát thải khí nhà kính (CO~2~) và không gây ô nhiễm khói bụi hay mưa acid, giảm biến đổi khí hậu.\n   - *An ninh năng lượng:* Giúp các quốc gia chủ động nguồn cung năng lượng tại chỗ, giảm lệ thuộc vào nhập khẩu xăng dầu than đá đắt đỏ.\n2. *Nhược điểm của năng lượng tái tạo:*\n   - *Tính phụ thuộc cao vào thời tiết và tự nhiên:* Năng lượng mặt trời không có vào ban đêm và giảm mạnh khi trời âm u, mưa bão; năng lượng gió thất thường phụ thuộc vào tốc độ gió.\n   - *Chi phí đầu tư ban đầu lớn:* Chi phí nghiên cứu, sản xuất và lắp đặt thiết bị (pin quang điện, tuabin gió, hệ thống lưu trữ điện ắc quy) còn rất cao.\n   - *Đòi hỏi diện tích mặt bằng lớn:* Cần hàng nghìn hecta đất hoặc mặt nước để bố trí các trang trại điện mặt trời, điện gió.\n   - *Khó khăn trong tích trữ:* Công nghệ pin lưu trữ dung lượng lớn chưa hoàn thiện và chi phí thay thế định kì đắt."
    }
]

# =============================================================================
# PHÂN MÔN KHTN 3 (SINH HỌC 9 - DI TRUYỀN VÀ BIẾN DỊ, GENE)
# =============================================================================

BIOLOGY_THEORY = [
    {
        "title": "A. MỘT SỐ THUẬT NGỮ CƠ BẢN",
        "items": [
            "*Tính trạng:* Là những đặc điểm về hình thái, cấu tạo, sinh lí của một cơ thể sinh vật.",
            "*Tính trạng tương phản:* Là hai trạng thái biểu hiện trái ngược nhau của cùng một loại tính trạng (ví dụ: thân cao đối lập với thân thấp; hoa tím đối lập với hoa trắng).",
            "*Nhân tố di truyền (gene):* Tồn tại thành từng cặp trong nhân tế bào, không hoà trộn vào nhau, quy định tính trạng của cơ thể sinh vật.",
            "*Kiểu hình:* Là tổ hợp toàn bộ tính trạng của cơ thể sinh vật biểu hiện ra bên ngoài.",
            "*Kiểu gene:* Là tổ hợp toàn bộ gene trong tế bào của cơ thể sinh vật.",
            "*Allele:* Là các trạng thái biểu hiện khác nhau của cùng một gene (ví dụ: gene quy định màu hoa có allele A quy định hoa tím, allele a quy định hoa trắng).",
            "*Cơ thể thuần chủng:* Là cơ thể có kiểu gene đồng hợp về một hoặc nhiều tính trạng đang xét (ví dụ: AA hoặc aa)."
        ]
    },
    {
        "title": "B. QUY LUẬT LAI MỘT CẶP TÍNH TRẠNG CỦA MENĐEN",
        "items": [
            "*1. Thí nghiệm của Menđen:*\n   P~tc~: Cây hoa tím  ×  Cây hoa trắng\n   F~1~: 100% Cây hoa tím\n   Cho F~1~ tự thụ phấn (F~1~ × F~1~)\n   F~2~: 705 cây hoa tím : 224 cây hoa trắng (tỉ lệ xấp xỉ 3 hoa tím : 1 hoa trắng).",
            "*2. Giải thích thí nghiệm:*\n   Quy ước gene: A quy định hoa tím (trội hoàn toàn), a quy định hoa trắng (lặn).\n   P~tc~: AA (hoa tím) × aa (hoa trắng)  => Giao tử P: A, a\n   F~1~: 100% Aa (kiểu hình: hoa tím)\n   F~1~ × F~1~: Aa × Aa  => Giao tử: (1/2 A : 1/2 a) × (1/2 A : 1/2 a)\n   F~2~: Tỉ lệ kiểu gene: 1 AA : 2 Aa : 1 aa\n         Tỉ lệ kiểu hình: 3 hoa tím : 1 hoa trắng.",
            "*3. Nội dung quy luật phân li:*\n   Mỗi tính trạng do một cặp nhân tố di truyền (cặp allele) quy định. Khi giảm phân hình thành giao tử, các allele trong cặp phân li đồng đều về các giao tử nên mỗi giao tử chỉ chứa một allele của cặp.",
            "*4. Lai phân tích:*\n   - Khái niệm: Là phép lai giữa cá thể mang tính trạng trội chưa biết kiểu gene với cá thể mang tính trạng lặn (đồng hợp lặn).\n   - Vai trò: Dùng để kiểm tra và xác định kiểu gene của cá thể mang tính trạng trội (nếu đời con đồng tính thì cá thể trội có kiểu gene đồng hợp AA; nếu đời con phân tính theo tỉ lệ 1 : 1 thì cá thể trội có kiểu gene dị hợp Aa)."
        ]
    },
    {
        "title": "C. QUY LUẬT PHÂN LI ĐỘC LẬP",
        "items": [
            "*1. Thí nghiệm lai hai cặp tính trạng của Menđen:*\n   P~tc~: Hạt vàng, vỏ trơn  ×  Hạt xanh, vỏ nhăn\n   F~1~: 100% Hạt vàng, vỏ trơn\n   Cho F~1~ tự thụ phấn thu được F~2~ phân li theo tỉ lệ:\n   9 Vàng, trơn : 3 Vàng, nhăn : 3 Xanh, trơn : 1 Xanh, nhăn.",
            "*2. Nội dung quy luật phân li độc lập:*\n   Các cặp nhân tố di truyền (cặp allele) quy định các cặp tính trạng khác nhau. Trong quá trình hình thành giao tử, cặp allele này phân li độc lập với cặp allele khác."
        ]
    },
    {
        "title": "IV. NUCLEIC ACID VÀ GENE",
        "items": [
            "*1. Deoxyribonucleic acid (DNA):*\n   - DNA là đại phân tử sinh học có kích thước rất dài (tới hàng trăm micromet), khối lượng phân tử đạt hàng triệu đến chục triệu amu.\n   - DNA có cấu tạo đa phân, đơn phân là nucleotide gồm 4 loại: Adenine (A), Thymine (T), Cytosine (C) và Guanine (G).\n   - Cấu trúc không gian: DNA có cấu trúc xoắn kép gồm hai chuỗi polynucleotide song song, ngược chiều, xoắn đều quanh một trục tưởng tượng từ trái sang phải (xoắn phải).\n   - Liên kết trên một mạch: Các nucleotide trên cùng một mạch liên kết với nhau bằng liên kết cộng hóa trị (phosphodiester) theo chiều từ 5' đến 3'.\n   - Liên kết giữa hai mạch: Các nucleotide giữa hai mạch đối diện liên kết với nhau bằng liên kết hydrogen theo nguyên tắc bổ sung (NTBS): A liên kết với T bằng 2 liên kết hydrogen (A = T), G liên kết với C bằng 3 liên kết hydrogen (G ≡ C) hoặc ngược lại.\n   - Thông số cấu trúc: Đường kính vòng xoắn là 20 Å (2 nm); mỗi chu kì xoắn dài 34 Å (3,4 nm) chứa đúng 10 cặp nucleotide.\n   - Chức năng: Lưu trữ, bảo quản và truyền đạt thông tin di truyền qua các thế hệ tế bào và cơ thể.",
            "*2. Ribonucleic acid (RNA):*\n   - RNA là đại phân tử nhưng kích thước và khối lượng thường nhỏ hơn nhiều so với DNA.\n   - Cấu tạo đa phân từ 4 loại đơn phân nucleotide: A (Adenine), U (Uracil), G (Guanine) và C (Cytosine).\n   - Trong tế bào, phân tử RNA thường có cấu trúc 1 mạch đơn, được tổng hợp trong nhân hoặc vùng nhân nhưng phần lớn thực hiện chức năng sinh học tại tế bào chất.\n   - Các loại RNA chính: mRNA (RNA thông tin làm khuôn dịch mã), tRNA (RNA vận chuyển amino acid), rRNA (RNA ribosome cấu tạo nên bào quan ribosome)."
        ]
    }
]

# Trắc nghiệm Sinh học nhiều lựa chọn (19 câu)
BIOLOGY_CHOICE_QUESTIONS = [
    {
        "id": 1,
        "content": "Di truyền là gì?",
        "options": [
            "A. Là quá trình truyền thông tin di động.",
            "B. Là quá trình truyền thông tin trong trên internet.",
            "C. Là quá trình truyền thông tin bằng miệng.",
            "D. Là sự truyền đạt các đặc điểm từ thế hệ này sang thế hệ khác."
        ],
        "correct": "D",
        "solution": "Di truyền là hiện tượng truyền đạt các tính trạng của bố mẹ, tổ tiên cho các thế hệ con cháu tiếp theo."
    },
    {
        "id": 2,
        "content": "Nội dung nào dưới đây không đúng khi nói về biến dị?",
        "options": [
            "A. Con cái sinh ra không giống với bố, mẹ chúng.",
            "B. Ở loài sinh sản hữu tính, có sự tổ hợp các gene của bố, mẹ tạo ra các biến dị.",
            "C. Bố mắt đen sinh ra con mắt đen là một biến dị.",
            "D. Bố, mẹ bình thường sinh con mắc bệnh Đao là một biến dị."
        ],
        "correct": "C",
        "solution": "Bố mắt đen sinh ra con mắt đen là hiện tượng di truyền (con cái giống với bố mẹ), không phải là biến dị. Biến dị là hiện tượng con sinh ra có những đặc điểm khác biệt so với bố mẹ và khác nhau giữa các cá thể trong cùng lứa."
    },
    {
        "id": 3,
        "content": "Gene là gì?",
        "options": [
            "A. Là một đoạn DNA mang thông tin di truyền mã hóa cho một sản phẩm nhất định.",
            "B. Là một loại DNA trong tế bào.",
            "C. Là một loại protein cụ thể trong tế bào.",
            "D. Là một loại acid amin."
        ],
        "correct": "A",
        "solution": "Gene là một đoạn xác định của phân tử DNA mang thông tin di truyền mã hóa cho một sản phẩm nhất định (chuỗi polypeptide hoặc một phân tử RNA)."
    },
    {
        "id": 4,
        "content": "Tính trạng là",
        "options": [
            "A. những biểu hiện của kiểu gen thành kiểu hình.",
            "B. kiểu hình bên ngoài cơ thể sinh vật.",
            "C. các đặc điểm bên trong cơ thể sinh vật.",
            "D. những đặc điểm về hình thái, cấu tạo, sinh lý của một cơ thể."
        ],
        "correct": "D",
        "solution": "Tính trạng là những đặc điểm cụ thể về hình thái (màu sắc, hình dạng), cấu tạo giải phẫu và sinh lí của một cơ thể sinh vật."
    },
    {
        "id": 5,
        "content": "Kiểu gen nào dưới đây được xem là thuần chủng?",
        "options": [
            "A. Aa.",
            "B. AA và Aa.",
            "C. AA và aa.",
            "D. AA, Aa và aa."
        ],
        "correct": "C",
        "solution": "Cơ thể thuần chủng về một tính trạng là cơ thể có kiểu gene đồng hợp (gồm 2 allele giống nhau trong cặp: đồng hợp trội AA hoặc đồng hợp lặn aa)."
    },
    {
        "id": 6,
        "content": "Khi đem lai các cá thể thuần chủng khác nhau về một cặp tính trạng tương phản, Menđen đã phát hiện được điều gì ở thế hệ con lai?",
        "options": [
            "A. Ở thế hệ con lai chỉ biểu hiện một trong hai kiểu hình của bố hoặc mẹ.",
            "B. Ở thế hệ con lai biểu hiện tính trạng trung gian giữa bố và mẹ.",
            "C. Ở thế hệ con lai luôn luôn biểu hiện kiểu hình giống bố.",
            "D. Ở thế hệ con lai luôn luôn biểu hiện kiểu hình giống mẹ."
        ],
        "correct": "A",
        "solution": "Khi lai hai dòng thuần chủng khác nhau về một cặp tính trạng tương phản, thế hệ con lai F~1~ đồng tính, tức là 100% cá thể chỉ biểu hiện kiểu hình của một bên bố hoặc mẹ (tính trạng trội)."
    },
    {
        "id": 7,
        "content": "Mục đích của phép lai phân tích nhằm xác định",
        "options": [
            "A. kiểu gen, kiểu hình của cá thể mang tính trạng trội.",
            "B. kiểu hình của cá thể mang tính trạng trội.",
            "C. kiểu gen của tất cả các tính trạng.",
            "D. kiểu gen của cá thể mang tính trạng trội."
        ],
        "correct": "D",
        "solution": "Mục đích chính của phép lai phân tích là để kiểm tra và xác định kiểu gene của cá thể mang tính trạng trội (là đồng hợp trội AA hay dị hợp tử Aa)."
    },
    {
        "id": 8,
        "content": "Phép lai nào sau đây được gọi là phép lai phân tích?",
        "options": [
            "A. Aa × Aa.",
            "B. Aa × AA.",
            "C. Aa × aa.",
            "D. AA × Aa."
        ],
        "correct": "C",
        "solution": "Lai phân tích là phép lai giữa cá thể mang tính trạng trội (ở đây là Aa) với cá thể mang tính trạng lặn (đồng hợp lặn aa)."
    },
    {
        "id": 9,
        "content": "Khi giao phấn giữa cây đậu Hà Lan thuần chủng có hạt vàng, vỏ trơn với cây có hạt xanh, vỏ nhăn thuần chủng thì kiểu hình thu được ở các cây lai F~1~ là",
        "options": [
            "A. 100% hạt vàng, vỏ trơn.",
            "B. 100% hạt vàng, vỏ nhăn.",
            "C. 100% hạt xanh, vỏ trơn.",
            "D. 100% hạt xanh, vỏ nhăn."
        ],
        "correct": "A",
        "solution": "Vì hạt vàng và vỏ trơn là hai tính trạng trội hoàn toàn so với hạt xanh và vỏ nhăn, nên phép lai P~tc~ (AABB × aabb) cho thế hệ F~1~ 100% có kiểu gene AaBb, biểu hiện kiểu hình 100% hạt vàng, vỏ trơn."
    },
    {
        "id": 10,
        "content": "Theo dõi thí nghiệm của Menđen, khi lai đậu Hà Lan thuần chủng hạt vàng, trơn và hạt xanh, nhăn với nhau thu được F~1~ đều hạt vàng, trơn. Khi cho F~1~ tự thụ phấn thì F~2~ có tỉ lệ kiểu hình là",
        "options": [
            "A. 9 vàng, nhăn : 3 vàng, trơn : 3 xanh, nhăn : 1 xanh, trơn.",
            "B. 9 vàng, trơn : 3 xanh, nhăn : 3 xanh, trơn : 1 vàng, nhăn.",
            "C. 9 vàng, nhăn : 3 xanh, nhăn : 3 vàng, trơn : 1 xanh, trơn.",
            "D. 9 vàng, trơn : 3 vàng, nhăn : 3 xanh, trơn : 1 xanh, nhăn."
        ],
        "correct": "D",
        "solution": "Khi F~1~ (AaBb) tự thụ phấn, đời F~2~ phân li kiểu hình theo quy luật phân li độc lập với tỉ lệ: 9 vàng, trơn : 3 vàng, nhăn : 3 xanh, trơn : 1 xanh, nhăn."
    },
    {
        "id": 11,
        "content": "Từ kết quả thí nghiệm lai hai cặp tính trạng, Menđen thấy rằng",
        "options": [
            "A. các tính trạng màu sắc hạt và hình dạng hạt di truyền phụ thuộc vào nhau.",
            "B. các tính trạng màu sắc hạt và hình dạng hạt di truyền không phụ thuộc vào nhau.",
            "C. các tính trạng màu sắc di truyền phụ thuộc vào nhau còn các tính trạng hình dạng hạt di truyền không phụ thuộc vào nhau.",
            "D. các tính trạng màu sắc di truyền không phụ thuộc vào nhau còn các tính trạng hình dạng hạt di truyền phụ thuộc vào nhau."
        ],
        "correct": "B",
        "solution": "Tỉ lệ phân li của từng cặp tính trạng ở F~2~ (3 vàng : 1 xanh và 3 trơn : 1 nhăn) nhân với nhau đúng bằng tỉ lệ kiểu hình chung (9 : 3 : 3 : 1), chứng tỏ các tính trạng màu sắc hạt và hình dạng hạt di truyền hoàn toàn độc lập và không phụ thuộc vào nhau."
    },
    {
        "id": 12,
        "content": "Bốn loại đơn phân cấu tạo DNA có kí hiệu là",
        "options": [
            "A. A, U, G, C.",
            "B. A, T, G, C.",
            "C. A, D, R, T.",
            "D. U, R, D, C."
        ],
        "correct": "B",
        "solution": "Bốn loại đơn phân nucleotide cấu tạo nên phân tử DNA là Adenine (A), Thymine (T), Guanine (G) và Cytosine (C)."
    },
    {
        "id": 13,
        "content": "Hai mạch polynucleotide của phân tử ADN liên kết với nhau bằng liên kết nào?",
        "options": [
            "A. Liên kết phosphodiester của nucleotide giữa mạch đơn này với nucleotide của mạch đơn kia.",
            "B. Liên kết hydrogen giữa nhóm phosphate của nucleotide mạch đơn này với phân tử đường trong nucleotide của mạch đơn kia.",
            "C. Liên kết hydrogen giữa các nitrogenous base của mạch đơn này với nitrogenous base mạch đơn kia theo nguyên tắc bổ sung (A – G, T – C).",
            "D. Liên kết hydrogen giữa các nitrogenous base của mạch đơn này với nitrogenous base mạch đơn kia theo nguyên tắc bổ sung (A – T, G – C)."
        ],
        "correct": "D",
        "solution": "Hai mạch đơn polynucleotide của phân tử DNA liên kết với nhau bằng các liên kết hydrogen giữa các base nitơ theo nguyên tắc bổ sung: A liên kết với T (bằng 2 liên kết hydrogen) và G liên kết với C (bằng 3 liên kết hydrogen)."
    },
    {
        "id": 14,
        "content": "Theo NTBS thì những trường hợp nào sau đây là đúng?\n(1) A + G = T + C\n(2) A + T = G + C\n(3) A = T; G = C\n(4) A + T + G = A + C + T\n(5) A + C + T = G + C + T",
        "options": [
            "A. 1, 2, 3.",
            "B. 1, 3, 4.",
            "C. 2, 3, 4.",
            "D. 3, 4, 5."
        ],
        "correct": "B",
        "solution": "Theo nguyên tắc bổ sung (NTBS) trong phân tử DNA dạng sợi kép:\n- (3) Đúng vì A = T và G = C.\n- (1) Đúng vì A + G = T + C (tương đương 50% tổng số nucleotide).\n- (4) Đúng vì do G = C nên vế trái A + T + G đúng bằng vế phải A + T + C.\n- (2) Sai vì A + T chỉ bằng G + C khi tỉ lệ %A = %G = 25% (không đúng tổng quát).\n- (5) Sai vì A và G nói chung khác nhau.\nVậy các hệ thức đúng là (1), (3), (4)."
    },
    {
        "id": 15,
        "content": "Điều nào dưới đây đúng khi nói về đặc điểm cấu tạo của phân tử mARN?",
        "options": [
            "A. Cấu tạo 2 mạch xoắn, song song.",
            "B. Cấu tạo gồm 1 mạch thẳng.",
            "C. Kích thước và khối lượng lớn hơn so với phân tử ADN.",
            "D. Gồm có 4 loại đơn phân là A, T, G, C."
        ],
        "correct": "B",
        "solution": "Phân tử mRNA (RNA thông tin) có cấu tạo gồm một mạch đơn dạng thẳng, được tổng hợp dựa trên một mạch khuôn của gene DNA."
    },
    {
        "id": 16,
        "content": "Loại nuclêôtit có ở ARN và không có ở ADN là",
        "options": [
            "A. Adenine.",
            "B. Thymine.",
            "C. Uracil.",
            "D. Guanine."
        ],
        "correct": "C",
        "solution": "Uracil (U) là loại nucleotide đặc trưng chỉ có trong cấu tạo phân tử RNA, trong phân tử DNA vị trí tương ứng là Thymine (T)."
    },
    {
        "id": 17,
        "content": "Đề cập đến chức năng của ARN, nội dung nào sau đây không đúng?",
        "options": [
            "A. rARN có vai trò tổng hợp các chuỗi polypeptide đặc biệt tạo thành ribosome.",
            "B. mARN mang thông tin quy định trình tự amino acid trong chuỗi polypeptide.",
            "C. tARN vận chuyển amino acid đến ribosome.",
            "D. rARN kết hợp với protein để tạo thành nhiễm sắc thể."
        ],
        "correct": "D",
        "solution": "rRNA liên kết với các phân tử protein đặc thù để cấu tạo nên bào quan ribosome (nơi tiến hành dịch mã tổng hợp protein), chứ không phải tạo thành nhiễm sắc thể (nhiễm sắc thể được cấu tạo từ DNA và protein loại histone)."
    },
    {
        "id": 18,
        "content": "Gene có chiều dài là 408 Å sẽ có bao nhiêu chu kì xoắn?",
        "options": [
            "A. 100.",
            "B. 210.",
            "C. 12.",
            "D. 120."
        ],
        "correct": "C",
        "solution": "Mỗi chu kì xoắn của phân tử DNA chuẩn dạng B có chiều dài là 34 Å.\nSố chu kì xoắn của gene là: C = L / 34 = 408 / 34 = 12 chu kì xoắn."
    },
    {
        "id": 19,
        "content": "Phát biểu nào sau đây không đúng về cấu trúc của mạch DNA?",
        "options": [
            "A. DNA có cấu trúc mạch xoắn kép gồm hai mạch song song và ngược chiều.",
            "B. Các nucleotide liên kết với nhau bằng liên kết cộng hóa trị.",
            "C. Chu kì xoắn, mỗi chu kì xoắn 34 Å tương ứng với 10 cặp nucleotide.",
            "D. Trên mạch polynucleotide theo chiều từ 3' đến 5'."
        ],
        "correct": "D",
        "solution": "Trên mỗi mạch đơn polynucleotide, các nucleotide liên kết với nhau bằng liên kết cộng hóa trị phosphodiester hình thành nên mạch theo chiều liên kết chuẩn từ 5' tới 3' (đầu 5'-phosphate liên kết với đầu 3'-OH của nucleotide kế tiếp)."
    }
]

from khtn9_sinhhoc_ontap_data import (
    TF_QUESTIONS as BIOLOGY_TF_QUESTIONS,
    SHORT_QUESTIONS_A as BIOLOGY_SHORT_QUESTIONS_A,
    SHORT_QUESTIONS_B as BIOLOGY_SHORT_QUESTIONS_B,
    SHORT_QUESTIONS_C as BIOLOGY_SHORT_QUESTIONS_C,
    ESSAY_QUESTIONS as BIOLOGY_ESSAY_QUESTIONS
)

BIOLOGY_SHORT_QUESTIONS = BIOLOGY_SHORT_QUESTIONS_A + BIOLOGY_SHORT_QUESTIONS_B + BIOLOGY_SHORT_QUESTIONS_C

