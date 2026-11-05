# -*- coding: utf-8 -*-
"""
Dữ liệu Đề kiểm tra Giữa học kì I - Môn KHTN 8 (ĐỀ CHÍNH THỨC)
Trường THCS Nguyễn Thị Minh Khai - UBND Phường Thủy Xuân - Năm học 2026 - 2027
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

EXAM_INFO = {
    "ubnd": "UBND PHƯỜNG THỦY XUÂN",
    "school": "TRƯỜNG THCS NGUYỄN THỊ MINH KHAI",
    "exam_title": "ĐỀ KIỂM TRA GIỮA HỌC KÌ I",
    "year": "NĂM HỌC: 2026 - 2027",
    "subject": "MÔN: KHOA HỌC TỰ NHIÊN - LỚP 8",
    "time_str": "Thời gian: 90 phút (không kể thời gian giao đề)",
    "official_label": "ĐỀ CHÍNH THỨC",
    "pages_label": "(Đề gồm có 04 trang)",
}

# -----------------------------------------------------------------------------
# A. TRẮC NGHIỆM (7,0 ĐIỂM)
# PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (3,0 điểm)
# Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu đúng 0,25 điểm.
# -----------------------------------------------------------------------------
PART1_QUESTIONS = [
    {
        "id": 1,
        "score": "0,25 điểm",
        "content": "Trong phòng thực hành Khoa học tự nhiên 8, dụng cụ nào sau đây được sử dụng phổ biến nhất để đo thể tích của một lượng chất lỏng?",
        "options": ["A. Cân điện tử.", "B. Bình chia độ.", "C. Nhiệt kế y tế.", "D. Lực kế lò xo."],
        "correct": "B",
        "solution": "Bình chia độ (ống đong) là dụng cụ chuyên dụng phổ biến nhất trong phòng thực hành dùng để đo thể tích chất lỏng. Cân điện tử dùng đo khối lượng; nhiệt kế y tế dùng đo nhiệt độ; lực kế dùng đo độ lớn của lực."
    },
    {
        "id": 2,
        "score": "0,25 điểm",
        "content": "Thiết bị điện nào sau đây dùng để đo cường độ dòng điện chạy qua một đoạn mạch kín trong thí nghiệm Khoa học tự nhiên 8?",
        "options": ["A. Vôn kế.", "B. Biến trở.", "C. Ampe kế.", "D. Joulemeter."],
        "correct": "C",
        "solution": "Ampe kế là thiết bị chuyên dùng để đo cường độ dòng điện trong mạch điện. Vôn kế dùng đo hiệu điện thế; biến trở dùng điều chỉnh điện trở; Joulemeter dùng đo điện năng tiêu thụ."
    },
    {
        "id": 3,
        "score": "0,25 điểm",
        "content": "Để thu gom, xử lý hóa chất thải lỏng trong phòng thí nghiệm tránh gây ô nhiễm môi trường và nguồn nước, quy tắc nào sau đây là đúng?",
        "options": [
            "A. Đổ trực tiếp hóa chất dư thừa vào bồn rửa và xả nhiều nước.",
            "B. Đem hòa tan vào nước rồi tưới cho cây xanh xung quanh phòng thí nghiệm.",
            "C. Gom hóa chất thải vào các bình chứa riêng có dán nhãn phân loại theo đúng quy chuẩn.",
            "D. Đổ tất cả các loại hóa chất thừa vào cùng một thùng chứa duy nhất."
        ],
        "correct": "C",
        "solution": "Quy tắc an toàn môi trường trong phòng thí nghiệm yêu cầu phải thu gom các hóa chất thải vào các bình chứa chuyên dụng riêng biệt, có dán nhãn phân loại rõ ràng (axit, bazơ, kim loại nặng, dung môi hữu cơ...) để xử lý theo đúng quy chuẩn."
    },
    {
        "id": 4,
        "score": "0,25 điểm",
        "content": "Để đảm bảo an toàn tuyệt đối cho học sinh khi tiến hành thực hành lắp ráp mạch điện, quy tắc an toàn điện quy định nguồn điện sử dụng trong phòng thực hành không được vượt quá giá trị nào?",
        "options": ["A. 380 V.", "B. 110 V.", "C. 220 V.", "D. Dưới 24 V."],
        "correct": "D",
        "solution": "Để đảm bảo an toàn cho học sinh phổ thông khi thực hành, nguồn điện một chiều hoặc xoay chiều sử dụng trong các bài thí nghiệm chỉ được dùng điện áp an toàn ở mức dưới 24 V (phổ biến từ 1,5 V đến 12 V)."
    },
    {
        "id": 5,
        "score": "0,25 điểm",
        "content": "Trong các trường hợp sau, trường hợp nào diễn ra sự biến đổi hóa học?",
        "options": [
            "A. Dây sắt bị uốn cong thành nhiều vòng.",
            "B. Hòa tan đường vào nước ấm thu được nước đường.",
            "C. Cồn để trong lọ không kín bị bay hơi.",
            "D. Đốt cháy ngọn nến, nến cháy tạo thành khí carbon dioxide và hơi nước."
        ],
        "correct": "D",
        "solution": "Đốt cháy nến (parafin) có phản ứng cháy với oxygen tạo thành các chất mới là khí carbon dioxide (CO~2~) và hơi nước (H~2~O), nên đây là sự biến đổi hóa học. Các trường hợp uốn dây sắt, hòa tan đường, cồn bay hơi chỉ là sự biến đổi vật lí (thay đổi hình dạng, trạng thái)."
    },
    {
        "id": 6,
        "score": "0,25 điểm",
        "content": "Mol là gì?",
        "options": [
            "A. Lượng chất có chứa 6,022 x 10^23^ nguyên tử hoặc phân tử chất đó.",
            "B. Số nguyên tử có trong 1 gam chất đó.",
            "C. Khối lượng của 6,022 x 10^23^ nguyên tử hoặc phân tử chất đó.",
            "D. Thể tích chiếm bởi 6,022 x 10^23^ nguyên tử hoặc phân tử chất đó."
        ],
        "correct": "A",
        "solution": "Mol là lượng chất có chứa 6,022 × 10^23^ nguyên tử hoặc phân tử của chất đó (con số 6,022 × 10^23^ gọi là hằng số Avogadro, kí hiệu là N~A~)."
    },
    {
        "id": 7,
        "score": "0,25 điểm",
        "content": "Ở điều kiện chuẩn (25°C và 1 bar) thể tích của 0,2 mol khí carbon dioxide (CO~2~) là bao nhiêu?",
        "options": ["A. 4,958 lít.", "B. 2,479 lít.", "C. 5,6 lít.", "D. 4,48 lít."],
        "correct": "A",
        "solution": "Ở điều kiện chuẩn (25°C, 1 bar), 1 mol chất khí bất kì đều chiếm thể tích là 24,79 lít.\nThể tích của 0,2 mol khí CO~2~ là: V = n × 24,79 = 0,2 × 24,79 = 4,958 lít."
    },
    {
        "id": 8,
        "score": "0,25 điểm",
        "content": "Thực hành xác định khối lượng riêng của một hòn sỏi không thấm nước bằng cân điện tử và bình chia độ, học sinh thu được các số liệu: khối lượng sỏi m = 27 g; thể tích nước ban đầu V~1~ = 50 cm^3^; thể tích nước dâng lên sau khi thả ngập sỏi V~2~ = 60 cm^3^. Khối lượng riêng của hòn sỏi là",
        "options": ["A. 2700 kg/m^3^.", "B. 0,27 g/cm^3^.", "C. 270 kg/m^3^.", "D. 27 g/cm^3^."],
        "correct": "A",
        "solution": "- Thể tích của hòn sỏi: V = V~2~ - V~1~ = 60 - 50 = 10 cm^3^ = 10^-5^ m^3^.\n- Khối lượng của hòn sỏi: m = 27 g = 0,027 kg.\n- Khối lượng riêng của hòn sỏi: D = m / V = 27 / 10 = 2,7 g/cm^3^ = 2700 kg/m^3^."
    },
    {
        "id": 9,
        "score": "0,25 điểm",
        "content": "Một học sinh nhập số liệu thực hành xác định khối lượng riêng của một khối gỗ hình hộp chữ nhật vào phần mềm bảng tính. Kích thước đo được: chiều dài a = 10 cm, chiều rộng b = 5 cm, chiều cao c = 2 cm; khối lượng m = 80 g. Thuật toán lập công thức trên bảng tính sẽ tính ra kết quả khối lượng riêng D theo đơn vị chuẩn kg/m^3^ là",
        "options": ["A. 800 kg/m^3^.", "B. 0,8 kg/m^3^.", "C. 80 kg/m^3^.", "D. 8000 kg/m^3^."],
        "correct": "A",
        "solution": "- Thể tích của khối gỗ: V = a × b × c = 10 × 5 × 2 = 100 cm^3^ = 10^-4^ m^3^.\n- Khối lượng khối gỗ: m = 80 g = 0,08 kg.\n- Khối lượng riêng theo đơn vị chuẩn: D = m / V = 0,08 / 10^-4^ = 800 kg/m^3^ (hoặc D = 80 / 100 = 0,8 g/cm^3^ = 800 kg/m^3^)."
    },
    {
        "id": 10,
        "score": "0,25 điểm",
        "content": "Khi nhổ một chiếc đinh nẹp gỗ bám rất chặt, người thợ mộc đã lồng thêm một đoạn ống sắt dài vào cán chiếc búa nhổ đinh. Tác dụng thực tiễn của việc làm này dựa trên nguyên lý sử dụng đòn bẩy là",
        "options": [
            "A. làm giảm khoảng cách từ điểm tựa đến đinh.",
            "B. tăng khoảng cách từ điểm tựa đến giá của lực kéo, làm tăng moment lực kéo giúp nhổ đinh dễ hơn.",
            "C. làm tăng lực bám của búa vào đầu đinh.",
            "D. đổi chiều của lực tác dụng từ vuông góc sang song song."
        ],
        "correct": "B",
        "solution": "Lồng thêm ống sắt vào cán búa giúp làm tăng khoảng cách từ điểm tựa đến điểm đặt lực kéo (tăng cánh tay đòn), do đó làm tăng moment của lực tác dụng. Điều này giúp người thợ chỉ cần tác dụng một lực nhỏ hơn mà vẫn nhổ được chiếc đinh bám chặt."
    },
    {
        "id": 11,
        "score": "0,25 điểm",
        "content": "Hệ cơ quan nào có chức năng chủ yếu là vận chuyển oxygen, chất dinh dưỡng và các chất khác đến tế bào, đồng thời vận chuyển chất thải đến cơ quan bài tiết?",
        "options": ["A. Hệ tiêu hoá.", "B. Hệ tuần hoàn.", "C. Hệ bài tiết.", "D. Hệ hô hấp."],
        "correct": "B",
        "solution": "Hệ tuần hoàn (gồm tim và hệ thống mạch máu) có chức năng chính là vận chuyển máu chứa oxygen, chất dinh dưỡng đi nuôi tất cả các tế bào trong cơ thể, đồng thời thu gom khí carbon dioxide và các chất thải chuyển đến phổi và thận để đào thải ra ngoài."
    },
    {
        "id": 12,
        "score": "0,25 điểm",
        "content": "Ở khoang miệng người, enzyme amylase trong nước bọt thực hiện quá trình biến đổi chất nào sau đây?",
        "options": [
            "A. Cellulose thành đường glucose.",
            "B. Tinh bột chín thành đường maltose.",
            "C. Protein thành chuỗi polypeptide.",
            "D. Lipid thành glycerol."
        ],
        "correct": "B",
        "solution": "Trong nước bọt ở khoang miệng có chứa enzyme amylase, enzyme này xúc tác quá trình tiêu hóa hóa học thủy phân một phần tinh bột chín thành đường maltose (đường đôi)."
    }
]

# -----------------------------------------------------------------------------
# PHẦN II. Câu trắc nghiệm đúng/sai (2,0 điểm)
# Thí sinh trả lời từ câu 13 đến câu 14. Mỗi câu 1,0 điểm; mỗi ý đúng 0,25 điểm.
# -----------------------------------------------------------------------------
PART2_QUESTIONS = [
    {
        "id": 13,
        "score": "1,0 điểm",
        "title": "Trong tiết học thực hành môn Khoa học tự nhiên về chủ đề “Dung dịch và nồng độ”, một nhóm học sinh tiến hành tìm hiểu lí thuyết và tính toán các bước pha chế dung dịch trong phòng thí nghiệm.",
        "statements": [
            {
                "label": "a",
                "text": "Dung dịch là hỗn hợp lỏng đồng nhất của dung môi và chất tan đã tan hoàn toàn vào nhau; độ tan của một chất trong nước là số gam chất đó tan tối đa trong 100 gam nước để tạo thành dung dịch bão hòa ở nhiệt độ xác định.",
                "correct": True,
                "explain": "Đúng. Đây là định nghĩa khoa học chuẩn xác về dung dịch và độ tan của một chất trong nước ở nhiệt độ xác định."
            },
            {
                "label": "b",
                "text": "Để pha chế 200mL dung dịch NaCl có nồng độ 0,5M (biết M~NaCl~ = 58,5 g/mol), khối lượng chất tan NaCl cần dùng là 5,85 gam.",
                "correct": True,
                "explain": "Đúng. Đổi V = 200 mL = 0,2 lít. Số mol NaCl: n = C_M × V = 0,5 × 0,2 = 0,1 mol. Khối lượng chất tan: m = n × M = 0,1 × 58,5 = 5,85 gam."
            },
            {
                "label": "c",
                "text": "Khi sử dụng cổng quang điện, cần đặt cổng ở cuối quãng đường chuyển động để đo tốc độ trung bình của vật trên toàn bộ quãng đường.",
                "correct": False,
                "explain": "Sai. Để đo tốc độ trung bình trên toàn bộ quãng đường cần xác định tổng thời gian chuyển động (phải dùng 2 cổng quang điện đặt ở đầu và cuối quãng đường nối với đồng hồ đo thời gian hiện số). Đặt 1 cổng quang điện ở cuối quãng đường chỉ đo được thời gian chắn quang qua cổng đó (đo tốc độ tức thời tại điểm đó)."
            },
            {
                "label": "d",
                "text": "Khi tiến hành thí nghiệm pha chế dung dịch theo nồng độ cho trước, bước đầu tiên phải cho toàn bộ lượng nước cất định mức vào bình, sau đó mới cho từ từ chất tan vào và khuấy đều cho đến khi tan hết.",
                "correct": False,
                "explain": "Sai. Quy trình chuẩn là: Cho chất tan đã cân vào bình trước, sau đó rót một lượng nước cất vừa phải vào khuấy cho tan hoàn toàn, rồi mới thêm tiếp nước cất từ từ đến vạch định mức cần pha."
            }
        ]
    },
    {
        "id": 14,
        "score": "1,0 điểm",
        "title": "Trong buổi lao động, dọn dẹp vệ sinh và chăm sóc cảnh quan khuôn viên trường THCS Nguyễn Thị Minh Khai, phường Thủy Xuân, thành phố Huế các bạn học sinh lớp 8 đã sử dụng nhiều dụng cụ như: xà beng bẩy gốc cây, tảng đá; kéo tỉa cây... dựa vào tác dụng làm quay của lực và nguyên lý đòn bẩy.",
        "statements": [
            {
                "label": "a",
                "text": "Tác dụng làm quay của lực lên một vật quanh một điểm hoặc một trục được đặc trưng bằng đại lượng gọi là moment lực. Đòn bẩy là một công cụ có thể quay quanh điểm tựa O và có tác dụng làm thay đổi hướng hoặc độ lớn của lực tác dụng lên vật.",
                "correct": True,
                "explain": "Đúng. Đây là khái niệm chuẩn xác về moment lực và tác dụng cơ học của đòn bẩy."
            },
            {
                "label": "b",
                "text": "Một lực tác dụng vào vật sẽ luôn làm vật quay quanh trục kể cả khi giá của lực đó cắt trục quay hoặc song song với trục quay.",
                "correct": False,
                "explain": "Sai. Lực không có tác dụng làm quay vật nếu giá của lực đi qua trục quay (cắt trục quay) hoặc song song với trục quay."
            },
            {
                "label": "c",
                "text": "Ứng dụng của đòn bẩy là sự vận dụng quy luật moment lực để làm giảm lực tác dụng kéo/đẩy F~2~ cần thiết nâng một vật nặng P, khoảng cách từ điểm tựa O đến giá của lực tác dụng (OO~2~) phải lớn hơn khoảng cách từ điểm tựa O đến giá của lực nâng vật (OO~1~).",
                "correct": True,
                "explain": "Đúng. Theo điều kiện cân bằng đòn bẩy: P × OO~1~ = F~2~ × OO~2~ => F~2~ = P × (OO~1~ / OO~2~). Để F~2~ < P thì cần OO~2~ > OO~1~."
            },
            {
                "label": "d",
                "text": "Trong thực tế khi bẩy hòn đá to hoặc nhổ cọc rào bám chặt trong sân trường, các bạn học sinh lồng thêm một đoạn ống thép dài vào cán xà beng là đã áp dụng nguyên lý tăng khoảng cách từ giá của lực đến trục quay để làm tăng moment lực tác dụng.",
                "correct": True,
                "explain": "Đúng. Lồng thêm đoạn ống thép dài làm tăng cánh tay đòn của lực tác dụng, do đó làm tăng moment lực tác dụng giúp nâng vật dễ dàng hơn."
            }
        ]
    }
]

# -----------------------------------------------------------------------------
# PHẦN III. Câu trắc nghiệm trả lời ngắn (2,0 điểm)
# Thí sinh trả lời từ câu 15 đến câu 22 vào giấy làm bài kiểm tra. Mỗi câu 0,25 điểm.
# -----------------------------------------------------------------------------
PART3_QUESTIONS = [
    {
        "id": 15,
        "score": "0,25 điểm",
        "content": "Cho sơ đồ phản ứng hóa học sau: P + O~2~ --t°--> P~2~O~5~. Sau khi lập phương trình hóa học với các hệ số nguyên tối giản, hãy cho biết hệ số của chất sản phẩm (P~2~O~5~) bằng bao nhiêu?",
        "correct": "2",
        "solution": "Phương trình hóa học được cân bằng như sau:\n4P + 5O~2~  → (t°)  2P~2~O~5~.\nHệ số của chất sản phẩm (P~2~O~5~) là 2."
    },
    {
        "id": 16,
        "score": "0,25 điểm",
        "content": "Cho phương trình hóa học: Zn + 2HCl → ZnCl~2~ + H~2~↑. Dựa vào ý nghĩa của phương trình hóa học, tỉ lệ số nguyên tử zinc (Zn) tham gia phản ứng với số phân tử khí hydrogen (H~2~) tạo thành là bao nhiêu? (Viết dạng tỉ lệ tối giản số, ví dụ: 1:2).",
        "correct": "1:1",
        "solution": "Theo phương trình hóa học: 1 nguyên tử Zn phản ứng với 2 phân tử HCl sinh ra 1 phân tử ZnCl~2~ và 1 phân tử H~2~.\nDo đó tỉ lệ giữa số nguyên tử Zn và số phân tử H~2~ là 1 : 1."
    },
    {
        "id": 17,
        "score": "0,25 điểm",
        "content": "Ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự nhận định đúng ở thí nghiệm sau:\nMột nhóm học sinh tiến hành thí nghiệm kiểm chứng định luật bảo toàn khối lượng bằng phản ứng giữa dung dịch hydrochloric acid (HCl) và mẩu đá vôi (CaCO~3~) trong bình kín có nút cao su:\nCaCO~3~ + 2HCl → CaCl~2~ + CO~2~↑ + H~2~O.\nHọc sinh đưa ra các nhận định sau về quá trình thí nghiệm:\n(1) Phải đậy thật chặt nút cao su của bình phản ứng trước khi cho hai chất tiếp xúc để khí CO~2~ sinh ra không thoát ra ngoài không khí.\n(2) Nếu không đậy kín nút bình, sau phản ứng số chỉ của cân sẽ nhỏ hơn số chỉ của cân ban đầu.\n(3) Khi đậy kín nút bình, kim cân hoặc số chỉ cân điện tử trước và sau phản ứng hoàn toàn không thay đổi.\n(4) Hiện tượng kim cân lệch (nếu mở nút) chứng minh định luật bảo toàn khối lượng không đúng đối với phản ứng tạo chất khí.",
        "correct": "123",
        "solution": "- (1) Đúng: để đảm bảo hệ phản ứng là hệ kín, khí CO~2~ không bay ra ngoài.\n- (2) Đúng: nếu mở nút, khí CO~2~ thoát ra làm khối lượng bình đo trên cân giảm xuống.\n- (3) Đúng: trong bình kín, tổng khối lượng các chất trước và sau phản ứng được bảo toàn, số chỉ cân không đổi.\n- (4) Sai: định luật bảo toàn khối lượng luôn đúng; cân lệch khi mở nắp là do khí thoát ra ngoài môi trường (hệ hở).\nCác nhận định đúng là (1), (2), (3) => Dãy chữ số tăng dần: 123."
    },
    {
        "id": 18,
        "score": "0,25 điểm",
        "content": "Khi đốt cháy kim loại aluminium (Al) trong bình chứa khí oxygen (O~2~), aluminium cháy sáng tạo thành bột màu trắng là aluminium oxide (Al~2~O~3~). Sơ đồ phản ứng hoá học được biểu diễn như sau:\nAl + O~2~ --t°--> Al~2~O~3~.\nSau khi lập phương trình hoá học hoàn chỉnh với các hệ số nguyên dương tối giản a, b, c theo dạng:\naAl + bO~2~ --t°--> cAl~2~O~3~\nTổng các hệ số cân bằng (a+b+c) của phương trình trên bằng bao nhiêu?",
        "correct": "9",
        "solution": "Cân bằng phương trình: 4Al + 3O~2~  → (t°)  2Al~2~O~3~.\nCác hệ số: a = 4, b = 3, c = 2.\nTổng hệ số: a + b + c = 4 + 3 + 2 = 9."
    },
    {
        "id": 19,
        "score": "0,25 điểm",
        "content": "Dùng đòn bẩy loại 1 để nâng một vật nặng có trọng lượng P = 5000 N. Khoảng cách từ điểm tựa O đến điểm đặt vật O~1~ là OO~1~ = 0,2 m; khoảng cách từ điểm tựa O đến điểm tác dụng lực O~2~ là OO~2~ = 0,8 m. Tính độ lớn lực F (theo đơn vị Niutơn - N) tối thiểu cần tác dụng vào O~2~ để nâng vật.",
        "correct": "1250",
        "solution": "Áp dụng điều kiện cân bằng của đòn bẩy:\nF × OO~2~ = P × OO~1~\n=> F = (P × OO~1~) / OO~2~ = (5000 × 0,2) / 0,8 = 1250 N."
    },
    {
        "id": 20,
        "score": "0,25 điểm",
        "content": "Khi sử dụng đòn bẩy loại 1 để bẩy một tảng đá, người ta thấy tỉ số cánh tay đòn OO~1~ / OO~2~ = 1/4. Tỉ số giữa lực tác dụng F~2~ cần thiết so với trọng lượng P của tảng đá (tức F~2~ / P) viết dưới dạng số thập phân bằng bao nhiêu?",
        "correct": "0,25",
        "solution": "Theo nguyên lí đòn bẩy: F~2~ × OO~2~ = P × OO~1~ => F~2~ / P = OO~1~ / OO~2~ = 1/4 = 0,25."
    },
    {
        "id": 21,
        "score": "0,25 điểm",
        "content": "Để nâng một vật nặng P = 200 N ở điểm O~1~, người ta tác dụng lực F~2~ = 50 N vào điểm O~2~ cách điểm tựa O một khoảng OO~2~ = 1,0 m. Khoảng cách OO~1~ (mét) từ điểm tựa O đến điểm đặt vật nặng bằng bao nhiêu?",
        "correct": "0,25",
        "solution": "Theo điều kiện cân bằng: P × OO~1~ = F~2~ × OO~2~\n=> OO~1~ = (F~2~ × OO~2~) / P = (50 × 1,0) / 200 = 0,25 m."
    },
    {
        "id": 22,
        "score": "0,25 điểm",
        "content": "Để nâng một cối đá có trọng lượng P = 8000 N, công nhân tác dụng lực F = 2000 N vào đầu O~2~ của đòn bẩy. Nếu cánh tay đòn đặt vật OO~1~ = 0,5 m thì khoảng cách cánh tay đòn OO~2~ (mét) phải bằng bao nhiêu để đòn bẩy cân bằng?",
        "correct": "2",
        "solution": "Theo điều kiện cân bằng của đòn bẩy: F × OO~2~ = P × OO~1~\n=> OO~2~ = (P × OO~1~) / F = (8000 × 0,5) / 2000 = 2 m."
    }
]

# -----------------------------------------------------------------------------
# B. TỰ LUẬN (3,0 ĐIỂM)
# Gồm 4 câu: Câu 23, 24, 25, 26 (mỗi câu 0,75 điểm)
# -----------------------------------------------------------------------------
ESSAY_QUESTIONS = [
    {
        "id": 23,
        "score": "0,75 điểm",
        "content": "Nêu điểm khác biệt cơ bản nhất giữa sự biến đổi vật lí và sự biến đổi hoá học. Dựa vào dấu hiệu đó, em hãy lấy một ví dụ về sự biến đổi vật lí và một ví dụ về sự biến đổi hoá học trong đời sống hằng ngày.",
        "solution": "- *Điểm khác biệt cơ bản nhất:*\n  + Sự biến đổi vật lí là sự biến đổi mà trong đó trạng thái, hình dạng, kích thước... của chất bị thay đổi nhưng *không tạo ra chất mới*.\n  + Sự biến đổi hoá học là sự biến đổi có *tạo ra chất mới* (chất này biến đổi thành chất khác).\n- *Ví dụ thực tế trong đời sống:*\n  + Ví dụ biến đổi vật lí: Nước đá tan chảy thành nước lỏng ở nhiệt độ phòng (hoặc bẻ cong đoạn dây nhôm, cồn bay hơi).\n  + Ví dụ biến đổi hoá học: Đốt cháy than củi sinh ra khí carbon dioxide và tro bụi (hoặc đinh sắt để ngoài không khí ẩm bị gỉ sét, sữa chua lên men)."
    },
    {
        "id": 24,
        "score": "0,75 điểm",
        "sub_questions": [
            {
                "sub_id": "a",
                "content": "Nêu định nghĩa khối lượng riêng của một chất.",
                "solution": "Khối lượng riêng của một chất là khối lượng của một đơn vị thể tích chất đó (tức là khối lượng của 1 m^3^ chất đó)."
            },
            {
                "sub_id": "b",
                "content": "Viết công thức tính khối lượng riêng. Giải thích rõ tên gọi, ký hiệu và đơn vị đo chuẩn trong hệ SI của từng đại lượng có mặt trong công thức.",
                "solution": "Công thức tính khối lượng riêng: D = m / V\nTrong đó:\n- D: Khối lượng riêng của chất, đơn vị chuẩn trong hệ SI là kilôgam trên mét khối (kg/m^3^).\n- m: Khối lượng của vật làm bằng chất đó, đơn vị chuẩn trong hệ SI là kilôgam (kg).\n- V: Thể tích của vật, đơn vị chuẩn trong hệ SI là mét khối (m^3^)."
            }
        ]
    },
    {
        "id": 25,
        "score": "0,75 điểm",
        "intro": "Dựa vào kiến thức về áp suất trên một bề mặt bị ép, hãy giải thích cơ sở khoa học của một số ứng dụng kỹ thuật và vật dụng sinh hoạt sau đây nhằm phục vụ lao động sản xuất và đời sống con người.",
        "sub_questions": [
            {
                "sub_id": "a",
                "content": "Tại sao mũi đinh, đầu kim khâu, mũi khoan, lưỡi dao, xẻng xúc đất lại được mài nhọn hoặc chế tạo rất mỏng?",
                "solution": "Theo công thức tính áp suất: p = F / S (áp suất tỉ lệ nghịch với diện tích bị ép S). Việc mài nhọn mũi đinh, kim khâu, lưỡi dao nhằm làm giảm diện tích tiếp xúc S đến mức rất nhỏ. Khi đó, cùng một lực tác dụng F sẽ tạo ra áp suất p rất lớn, giúp mũi đinh, kim khâu, lưỡi dao dễ dàng đâm thủng hoặc cắt, gọt vật thể."
            },
            {
                "sub_id": "b",
                "content": "Tại sao máy xúc, xe tăng có trọng lượng rất lớn lại được thiết kế chạy bằng hệ thống bánh xích rộng, còn đường ray tàu hỏa lại được đặt trên các thanh tà vẹt to?",
                "solution": "Máy xúc, xe tăng hay tàu hỏa đều có trọng lượng F cực kì lớn. Việc thiết kế hệ thống bánh xích bản rộng và đặt các thanh tà vẹt to dưới đường ray nhằm làm tăng diện tích tiếp xúc S. Khi diện tích S rất lớn, áp suất p = F / S tác dụng lên mặt đất/mặt đường sẽ giảm đi đáng kể, giúp xe tăng, máy xúc di chuyển không bị sa lún trên đất mềm và đường ray tàu hỏa không bị lún nứt, hư hỏng."
            }
        ]
    },
    {
        "id": 26,
        "score": "0,75 điểm",
        "intro": "Trong quá trình vận động: cơ, xương và khớp phối hợp với nhau để giúp cơ thể thực hiện các động tác.",
        "sub_questions": [
            {
                "sub_id": "a",
                "content": "Em hãy nêu 2 biện pháp để bảo vệ các cơ quan hệ vận động.",
                "solution": "Hai biện pháp bảo vệ hệ vận động:\n1. Duy trì chế độ dinh dưỡng hợp lí, cân đối, bổ sung đầy đủ canxi, phospho, vitamin D (như sữa, hải sản, trứng, rau xanh) để xương chắc khỏe.\n2. Thường xuyên rèn luyện thể dục thể thao điều độ, vừa sức; khởi động kĩ trước khi vận động và lao động, mang vác đồ vật đúng tư thế."
            },
            {
                "sub_id": "b",
                "content": "Một học sinh thường xuyên ngồi học sai tư thế. Em hãy nêu một hậu quả có thể xảy ra đối với hệ vận động và một biện pháp để phòng tránh.",
                "solution": "- *Hậu quả:* Thường xuyên ngồi học sai tư thế (ngồi vẹo người, cúi gập sát bàn, lệch vai) dễ dẫn đến cong vẹo cột sống (gù lưng, vẹo cột sống sang trái hoặc phải), lồng ngực bị thu hẹp ảnh hưởng đến tim phổi, đau mỏi cơ lưng và vai gáy.\n- *Biện pháp phòng tránh:* Ngồi học ngay ngắn, giữ lưng thẳng tự nhiên, hai vai ngang bằng, hai bàn chân đặt phẳng trên mặt sàn; không tì ngực vào cạnh bàn; giữ khoảng cách mắt đến sách vở khoảng 25 - 30 cm; sử dụng bàn ghế có kích thước phù hợp với vóc dáng."
            }
        ]
    }
]
