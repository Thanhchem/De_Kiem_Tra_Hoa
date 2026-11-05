# -*- coding: utf-8 -*-
"""
Dữ liệu Đề kiểm tra Giữa học kì I - Môn KHTN 8 (ĐỀ CHÍNH THỨC)
Trường THCS Nguyễn Du - UBND Phường Phú Xuân - Năm học 2026 - 2027
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

EXAM_INFO = {
    "ubnd": "UBND PHƯỜNG PHÚ XUÂN",
    "school": "TRƯỜNG THCS NGUYỄN DU",
    "exam_title": "ĐỀ KIỂM TRA GIỮA KÌ I",
    "year": "NĂM HỌC 2026 – 2027",
    "subject": "MÔN: KHOA HỌC TỰ NHIÊN 8",
    "time_str": "Thời gian làm bài: 60 phút",
    "official_label": "ĐỀ CHÍNH THỨC",
    "atomic_mass": "Cho biết khối lượng mol nguyên tử: C = 12; O = 16; N = 14; H = 1; Na = 23; Cl = 35,5"
}

# -----------------------------------------------------------------------------
# PHẦN A. TRẮC NGHIỆM KHÁCH QUAN (7,0 ĐIỂM)
# I. Trắc nghiệm nhiều lựa chọn (3,0 điểm; câu 1 - câu 12) - Mỗi câu đúng 0,25 điểm
# -----------------------------------------------------------------------------
PART1_QUESTIONS = [
    {
        "id": 1,
        "content": "Khi sử dụng dung dịch acid trong phòng thực hành, học sinh cần làm gì để bảo đảm an toàn?",
        "options": [
            "A. Dùng tay chạm vào dung dịch để kiểm tra.",
            "B. Đổ acid vào chai không có nhãn để dễ sử dụng.",
            "C. Pha loãng acid bằng cách đổ nước vào acid một cách nhanh chóng.",
            "D. Mang phương tiện bảo hộ phù hợp và thực hiện đúng hướng dẫn của giáo viên."
        ],
        "correct": "D",
        "solution": "Khi tiếp xúc với dung dịch acid (hóa chất có tính ăn mòn nguy hiểm), học sinh bắt buộc phải mang đầy đủ phương tiện bảo hộ (găng tay, kính bảo hộ, áo blouse) và tuân thủ tuyệt đối sự hướng dẫn của giáo viên. Tuyệt đối không dùng tay chạm trực tiếp và không được đổ nước vào acid đặc."
    },
    {
        "id": 2,
        "content": "Khi sử dụng hóa chất trong phòng thực hành Khoa học tự nhiên, học sinh cần thực hiện quy tắc nào sau đây?",
        "options": [
            "A. Tự ý lấy hóa chất để sử dụng theo ý muốn.",
            "B. Đọc kĩ nhãn hóa chất và tuân thủ hướng dẫn của giáo viên.",
            "C. Trộn các hóa chất với nhau để quan sát hiện tượng.",
            "D. Sử dụng hóa chất theo kinh nghiệm cá nhân mà không cần hướng dẫn."
        ],
        "correct": "B",
        "solution": "Quy tắc an toàn cơ bản nhất trong phòng thí nghiệm là luôn đọc kĩ nhãn hóa chất trước khi dùng và làm đúng theo hướng dẫn của giáo viên, không tự ý phối trộn hóa chất để tránh nguy cơ cháy nổ, bỏng hóa chất."
    },
    {
        "id": 3,
        "content": "Trước và sau một phản ứng hóa học, yếu tố nào sau đây thay đổi?",
        "options": [
            "A. Khối lượng các nguyên tử.",
            "B. Số lượng các nguyên tử.",
            "C. Liên kết giữa các nguyên tử.",
            "D. Thành phần các nguyên tố."
        ],
        "correct": "C",
        "solution": "Trong một phản ứng hóa học, bản chất các nguyên tử, khối lượng, số lượng nguyên tử và thành phần nguyên tố đều được giữ nguyên; chỉ có liên kết giữa các nguyên tử bị thay đổi làm cho phân tử chất này biến đổi thành phân tử chất khác."
    },
    {
        "id": 4,
        "content": "Phản ứng thu nhiệt là phản ứng hóa học trong đó có đặc điểm nào sau đây?",
        "options": [
            "A. Tỏa ra môi trường một lượng nhiệt lớn.",
            "B. Hấp thụ năng lượng dưới dạng nhiệt từ môi trường.",
            "C. Làm môi trường xung quanh nóng lên.",
            "D. Tự xảy ra ở điều kiện thường mà không cần cung cấp nhiệt."
        ],
        "correct": "B",
        "solution": "Phản ứng thu nhiệt là phản ứng hóa học hấp thụ năng lượng dưới dạng nhiệt từ môi trường xung quanh, khiến cho nhiệt độ của môi trường xung quanh bị giảm đi (mát/lạnh đi)."
    },
    {
        "id": 5,
        "content": "Đơn vị nào sau đây là đơn vị của áp suất?",
        "options": ["A. kg/m^3^.", "B. Pa.", "C. N/m^3^.", "D. N."],
        "correct": "B",
        "solution": "Đơn vị chuẩn của áp suất trong hệ đo lường SI là Pascal (kí hiệu là Pa, với 1 Pa = 1 N/m^2^). kg/m^3^ là đơn vị khối lượng riêng; N là đơn vị đo lực; N/m^3^ là đơn vị trọng lượng riêng."
    },
    {
        "id": 6,
        "content": "Hiện tượng nào dưới đây không phải là ví dụ về sự tồn tại của áp suất chất lỏng?",
        "options": [
            "A. Màng bong bóng chất lỏng bị phồng ra khi bịt kín đáy ống hình trụ hở hai đầu rồi nhúng vào chậu nước.",
            "B. Nước phun ra từ các lỗ nhỏ được đục xung quanh thành một vỏ chai nhựa chứa đầy nước.",
            "C. Quả bóng cao su căng phồng khi đổ đầy nước vào trong.",
            "D. Vỏ chai nhựa rỗng bị móp lại khi ta dùng miệng hút bớt không khí bên trong chai."
        ],
        "correct": "D",
        "solution": "Vỏ chai nhựa rỗng bị móp lại khi hút bớt không khí bên trong là do áp suất khí quyển bên ngoài lớn hơn áp suất không khí còn lại trong chai nén ép vào. Đây là ví dụ về áp suất khí quyển, không phải áp suất chất lỏng."
    },
    {
        "id": 7,
        "content": "Tật cong vẹo cột sống ở lứa tuổi học sinh chủ yếu do nguyên nhân nào sau đây gây ra?",
        "options": [
            "A. Ngồi học không đúng tư thế kéo dài.",
            "B. Thiếu hụt vitamin C kéo dài.",
            "C. Di truyền từ cha mẹ.",
            "D. Uống nhiều sữa và ăn nhiều bánh kẹo."
        ],
        "correct": "A",
        "solution": "Tật cong vẹo cột sống học đường ở lứa tuổi học sinh chủ yếu xuất phát từ thói quen ngồi học sai tư thế (cúi gập người, ngồi vẹo lưng, lệch vai) kéo dài hoặc mang cặp sách quá nặng lệch một bên."
    },
    {
        "id": 8,
        "content": "Ở người, cơ quan nào là nơi diễn ra quá trình tiêu hóa và hấp thụ phần lớn chất dinh dưỡng?",
        "options": ["A. Dạ dày.", "B. Ruột non.", "C. Ruột già.", "D. Thực quản."],
        "correct": "B",
        "solution": "Ruột non là cơ quan tiêu hóa dài nhất, có bề mặt hấp thụ cực lớn nhờ các nếp gấp, lông ruột và mạng mao mạch dày đặc, là nơi diễn ra quá trình biến đổi hóa học triệt để và hấp thụ trên 90% các chất dinh dưỡng của thức ăn."
    },
    {
        "id": 9,
        "content": "Tập thể dục thể thao thường xuyên không có ý nghĩa nào sau đây?",
        "options": ["A. Cơ bắp săn chắc.", "B. Cơ thể mệt mỏi.", "C. Cơ thể khỏe mạnh.", "D. Xương phát triển."],
        "correct": "B",
        "solution": "Tập thể dục thể thao điều độ giúp kích thích phát triển xương, cơ bắp săn chắc dẻo dai, tăng cường tuần hoàn và nâng cao sức khỏe; không gây mệt mỏi suy nhược cơ thể."
    },
    {
        "id": 10,
        "content": "Một trong các nguyên tắc lập khẩu phần ăn là",
        "options": [
            "A. ăn nhiều chất đạm và chất béo.",
            "B. ăn theo sở thích cá nhân.",
            "C. giảm tối đa năng lượng nạp vào.",
            "D. cân đối các thành phần dinh dưỡng."
        ],
        "correct": "D",
        "solution": "Nguyên tắc khoa học hàng đầu khi lập khẩu phần ăn là phải đảm bảo đáp ứng đủ nhu cầu năng lượng và cân đối hợp lí tỉ lệ giữa các nhóm chất dinh dưỡng (carbohydrate, protein, lipid, vitamin, chất khoáng và nước)."
    },
    {
        "id": 11,
        "content": "Một bạn học sinh bị ngã gãy xương cẳng tay khi chơi thể thao. Khi tiến hành sơ cứu tại chỗ, bước nào sau đây là bắt buộc phải thực hiện trước khi di chuyển nạn nhân đến cơ sở y tế?",
        "options": [
            "A. Nắn lại xương gãy giúp nạn nhân.",
            "B. Xoa bóp dầu nóng trực tiếp lên chỗ xương bị gãy.",
            "C. Dùng nẹp cố định đoạn xương bị gãy.",
            "D. Chườm nước nóng trực tiếp lên chỗ xương gãy."
        ],
        "correct": "C",
        "solution": "Khi sơ cứu gãy xương, bước quan trọng sống còn là dùng nẹp cố định chắc chắn đoạn xương bị gãy (cố định trên và dưới khớp liền kề) để chống di lệch đầu xương gãy chọc thủng mạch máu, dây thần kinh trước khi chuyển đến bệnh viện. Tuyệt đối không nắn bóp hay xoa dầu nóng."
    },
    {
        "id": 12,
        "content": "Trong giờ thực hành, một học sinh bị đứt tay và chảy máu nhẹ. Để sơ cứu đúng cách, học sinh cần làm gì?",
        "options": [
            "A. Dùng tay bẩn chạm vào vết thương để kiểm tra.",
            "B. Bôi trực tiếp hóa chất sát khuẩn mạnh vào vết thương.",
            "C. Dùng gạc sạch ép trực tiếp lên vết thương để cầm máu và báo cho giáo viên.",
            "D. Để vết thương tự chảy máu cho đến khi máu ngừng hẳn."
        ],
        "correct": "C",
        "solution": "Quy trình sơ cứu vết thương chảy máu nhẹ: Dùng một miếng gạc sạch (hoặc khăn vô trùng) ấn ép trực tiếp lên vết thương trong vài phút để cầm máu, sau đó sát trùng nhẹ bằng nước muối sinh lý hoặc cồn y tế nhẹ, băng lại và báo ngay cho giáo viên phụ trách."
    }
]

# -----------------------------------------------------------------------------
# II. Trắc nghiệm đúng - sai (2,0 điểm; câu 1 - câu 2) - Mỗi câu 1,0 điểm
# -----------------------------------------------------------------------------
PART2_QUESTIONS = [
    {
        "id": 1,
        "score": "1,0 điểm",
        "title": "Gần đây, bạn Nam thường ăn đồ ngọt, ít đánh răng và có thói quen ăn uống không đúng giờ. Lời khuyên nào sau đây đúng, lời khuyên nào sai?",
        "statements": [
            {
                "label": "a",
                "text": "Vệ sinh răng miệng không đúng cách có thể làm tăng nguy cơ mắc bệnh răng, miệng.",
                "correct": True,
                "explain": "Đúng. Thức ăn thừa và đường đọng lại tạo môi trường cho vi khuẩn phát triển, lên men tạo acid phá hủy men răng dẫn đến sâu răng và viêm nướu."
            },
            {
                "label": "b",
                "text": "Ăn uống hợp vệ sinh, đúng giờ và đủ chất góp phần phòng ngừa một số bệnh đường tiêu hóa.",
                "correct": True,
                "explain": "Đúng. Ăn đúng giờ giúp dạ dày và tuyến tiêu hóa tiết dịch ổn định theo nhịp sinh học, phòng tránh viêm loét dạ dày - tá tràng và rối loạn tiêu hóa."
            },
            {
                "label": "c",
                "text": "Khi bị đau bụng và tiêu chảy kéo dài, chỉ cần tự mua thuốc uống mà không cần đi khám.",
                "correct": False,
                "explain": "Sai. Tiêu chảy và đau bụng kéo dài có thể là dấu hiệu nhiễm khuẩn tiêu hóa nguy hiểm hoặc mất nước nghiêm trọng; cần đến cơ sở y tế để được bác sĩ thăm khám và điều trị đúng nguyên nhân, không được tự ý dùng thuốc."
            },
            {
                "label": "d",
                "text": "Để phòng bệnh tiêu hóa, Nam chỉ cần hạn chế ăn đồ ngọt.",
                "correct": False,
                "explain": "Sai. Hạn chế đồ ngọt chỉ là một phần nhỏ; phòng bệnh tiêu hóa đòi hỏi phải ăn chín uống sôi, vệ sinh răng miệng đều đặn, ăn đúng bữa, cân đối đủ các nhóm chất và đảm bảo an toàn thực phẩm."
            }
        ]
    },
    {
        "id": 2,
        "score": "1,0 điểm",
        "title": "Đại dịch cúm Châu Á vào những năm 1957 – 1958 đã cướp đi sinh mạng của ít nhất 1 triệu người. Dịch cúm thường xảy ra vào mùa lạnh với tỷ lệ mắc bệnh hàng năm dao động từ 5 – 10% ở người trưởng thành và 20 – 30% ở trẻ em.",
        "statements": [
            {
                "label": "a",
                "text": "Cúm là căn bệnh chủ yếu liên quan đến hệ tuần hoàn.",
                "correct": False,
                "explain": "Sai. Cúm là bệnh truyền nhiễm cấp tính do virus cúm tấn công trực tiếp vào đường hô hấp (mũi, họng, phế quản, phổi), liên quan chủ yếu đến hệ hô hấp."
            },
            {
                "label": "b",
                "text": "Hệ miễn dịch suy yếu và môi trường không khí ô nhiễm không ảnh hưởng đến khả năng mắc bệnh cúm.",
                "correct": False,
                "explain": "Sai. Hệ miễn dịch suy giảm làm mất khả năng chống lại virus, không khí ô nhiễm làm tổn thương niêm mạc hô hấp, khiến virus cúm dễ dàng xâm nhập và bùng phát bệnh."
            },
            {
                "label": "c",
                "text": "Biểu hiện của bệnh cúm là đau đầu, sốt, nhức mỏi cơ, chảy nước mũi, khan tiếng.",
                "correct": True,
                "explain": "Đúng. Đây là các triệu chứng lâm sàng rất điển hình của hội chứng cúm do virus."
            },
            {
                "label": "d",
                "text": "Để phòng bệnh cúm cần nâng cao sức đề kháng, giữ vệ sinh hệ hô hấp, đeo khẩu trang khi đi ngoài đường.",
                "correct": True,
                "explain": "Đúng. Giữ ấm, tăng cường dinh dưỡng, đeo khẩu trang nơi đông người, rửa tay thường xuyên và tiêm vaccine phòng cúm hằng năm là các biện pháp phòng bệnh hiệu quả nhất."
            }
        ]
    }
]

# -----------------------------------------------------------------------------
# III. Trắc nghiệm trả lời ngắn (2,0 điểm; câu 1 - câu 4) - Mỗi câu 0,5 điểm
# -----------------------------------------------------------------------------
PART3_QUESTIONS = [
    {
        "id": 1,
        "score": "0,5 điểm",
        "content": "Khối lượng riêng của dầu ăn khoảng 800 kg/m^3^. Vậy 2 lít dầu ăn sẽ có khối lượng khoảng bao nhiêu kg?",
        "correct": "1,6",
        "solution": "Đổi thể tích dầu ăn ra đơn vị m^3^:\nV = 2 lít = 2 dm^3^ = 0,002 m^3^.\nKhối lượng của 2 lít dầu ăn là:\nm = D × V = 800 × 0,002 = 1,6 kg.\n=> Đáp số: 1,6."
    },
    {
        "id": 2,
        "score": "0,5 điểm",
        "content": "Đầu mũi đinh luôn được mài rất nhọn, thiết kế này giúp áp suất tác dụng lên bề mặt khi đóng đinh thay đổi như thế nào?",
        "correct": "Tăng lên",
        "solution": "Theo công thức tính áp suất: p = F / S (áp suất tỉ lệ nghịch với diện tích bị ép S). Việc mài nhọn mũi đinh làm cho diện tích tiếp xúc S rất nhỏ, do đó cùng một lực đóng F sẽ tạo ra áp suất p tăng lên rất lớn, giúp đinh dễ dàng đâm sâu vào bề mặt vật liệu.\n=> Đáp số: Tăng lên (hoặc Tăng)."
    },
    {
        "id": 3,
        "score": "0,5 điểm",
        "content": "Cho quá trình sau:\nĐường kính --nước (1)--> Nước đường --cô cạn (2)--> Đường kính --t° (3)--> Đường nóng chảy --t° (4)--> Than\nGiai đoạn nào có biến đổi hóa học?",
        "correct": "Giai đoạn (4)",
        "solution": "- Giai đoạn (1): Hòa tan đường vào nước tạo dung dịch đường (biến đổi vật lí).\n- Giai đoạn (2): Cô cạn nước để thu hồi lại đường kính (biến đổi vật lí).\n- Giai đoạn (3): Đun nóng làm đường chuyển từ thể rắn sang thể lỏng nóng chảy (biến đổi vật lí - chuyển thể).\n- Giai đoạn (4): Đun nóng ở nhiệt độ cao làm đường bị cháy đen, phân hủy thành than (carbon) và hơi nước, có chất mới sinh ra => Đây là biến đổi hóa học.\n=> Đáp số: Giai đoạn (4) (hoặc (4))."
    },
    {
        "id": 4,
        "score": "0,5 điểm",
        "content": "Trong các hầm lò sâu hoặc dưới đáy giếng cạn thường tích tụ nhiều khí carbon dioxide (CO~2~) có thể gây ngạt thở cho người xuống thám hiểm. Hiện tượng này xảy ra do khí CO~2~ nặng hơn không khí nên có xu hướng chìm xuống dưới đáy. Em hãy cho biết khí carbon dioxide nặng hơn không khí xấp xỉ bao nhiêu lần? Làm tròn kết quả đến chữ số thập phân thứ nhất. Cho biết khối lượng mol nguyên tử: C = 12; O = 16.",
        "correct": "1,5",
        "solution": "Khối lượng mol của khí carbon dioxide (CO~2~):\nM(CO~2~) = 12 + 16 × 2 = 44 g/mol.\nKhối lượng mol trung bình của không khí:\nM(không khí) ≈ 29 g/mol.\nTỉ khối của khí CO~2~ đối với không khí:\nd(CO~2~/kk) = M(CO~2~) / M(kk) = 44 / 29 ≈ 1,517... ≈ 1,5 lần.\n=> Đáp số: 1,5."
    }
]

# -----------------------------------------------------------------------------
# PHẦN B. TỰ LUẬN (3,0 ĐIỂM)
# Gồm 3 câu: Câu 1, Câu 2, Câu 3 (mỗi câu 1,0 điểm)
# -----------------------------------------------------------------------------
ESSAY_QUESTIONS = [
    {
        "id": 1,
        "score": "1,0 điểm",
        "sub_questions": [
            {
                "sub_id": "a",
                "score": "0,5 điểm",
                "content": "Thể tích mol của chất khí là gì?",
                "solution": "Thể tích mol của chất khí là thể tích chiếm bởi 1 mol chất khí đó (tương ứng với 6,022 × 10^23^ phân tử khí) ở những điều kiện xác định về nhiệt độ và áp suất. Ở điều kiện chuẩn (25°C và 1 bar), 1 mol của mọi chất khí đều chiếm thể tích bằng nhau và bằng 24,79 lít."
            },
            {
                "sub_id": "b",
                "score": "0,5 điểm",
                "content": "Trong một buổi thực hành Hóa học xanh, học sinh thu gom được 14 gam khí nitrogen (N~2~) tinh khiết từ một phản ứng phân hủy. Hãy tính thể tích thực tế (ở điều kiện chuẩn) của lượng khí nitrogen đã thu gom được? (Cho biết khối lượng mol nguyên tử N = 14)",
                "solution": "- Khối lượng mol phân tử của khí nitrogen (N~2~): M(N~2~) = 14 × 2 = 28 (g/mol).\n- Số mol khí nitrogen thu gom được:\n  n(N~2~) = m / M = 14 / 28 = 0,5 (mol).\n- Thể tích thực tế của khí nitrogen ở điều kiện chuẩn (25°C, 1 bar):\n  V(N~2~) = n × 24,79 = 0,5 × 24,79 = 12,395 (lít)."
            }
        ]
    },
    {
        "id": 2,
        "score": "1,0 điểm",
        "content": "Trong giờ thực hành Khoa học tự nhiên 8, nhóm của bạn Nam được cô giáo giao nhiệm vụ xác định khối lượng riêng của một khối gỗ hình hộp chữ nhật nhỏ bằng các dụng cụ có sẵn trong phòng thực hành gồm: một chiếc cân điện tử và một chiếc thước kẻ. Em hãy giúp nhóm bạn Nam liệt kê ngắn gọn 3 bước cơ bản cần thực hiện để xác định khối lượng riêng của khối gỗ trên.",
        "solution": "Ba bước cơ bản để xác định khối lượng riêng của khối gỗ hình hộp chữ nhật:\n- *Bước 1 (Đo khối lượng):* Đặt khối gỗ lên đĩa cân điện tử để đo khối lượng m của khối gỗ (ghi lại giá trị theo đơn vị gam hoặc kilôgam).\n- *Bước 2 (Xác định thể tích):* Dùng thước kẻ đo 3 kích thước của khối gỗ gồm chiều dài (a), chiều rộng (b) và chiều cao (c). Sau đó tính thể tích của khối gỗ theo công thức: V = a × b × c (đổi về cùng đơn vị m^3^ hoặc cm^3^).\n- *Bước 3 (Tính khối lượng riêng):* Tính khối lượng riêng D của khối gỗ bằng công thức: D = m / V (đổi đơn vị kết quả sang kg/m^3^ hoặc g/cm^3^)."
    },
    {
        "id": 3,
        "score": "1,0 điểm",
        "content": "Sau bữa ăn tối tại gia đình bạn Trang, một số người xuất hiện triệu chứng đau bụng, buồn nôn và tiêu chảy, nghi ngờ do sử dụng thực phẩm không bảo đảm an toàn. Em hãy đề xuất các biện pháp phù hợp giúp bạn Trang xử lí tình huống trên và phòng tránh nguy cơ tương tự?",
        "solution": "- *1. Các biện pháp xử lí tình huống trước mắt (cấp cứu kịp thời):*\n  + Dừng ngay việc sử dụng các thức ăn, nước uống nghi ngờ gây ngộ độc; lưu giữ lại mẫu thức ăn thừa để phục vụ kiểm nghiệm y tế.\n  + Cho người bệnh uống nhiều nước ấm hoặc bù nước và chất điện giải bằng dung dịch oresol (pha đúng tỉ lệ hướng dẫn) để tránh kiệt sức vì mất nước do nôn và tiêu chảy.\n  + Nhanh chóng đưa người bệnh đến cơ sở y tế gần nhất để được bác sĩ khám và điều trị chuyên khoa, tuyệt đối không tự ý uống thuốc cầm tiêu chảy khi chưa có chỉ định.\n- *2. Các biện pháp phòng tránh nguy cơ tái diễn về sau:*\n  + Thực hiện ăn chín, uống sôi; lựa chọn thực phẩm tươi sống, có nguồn gốc rõ ràng, không ôi thiu nấm mốc hay quá hạn sử dụng.\n  + Bảo quản thực phẩm đúng cách (trong tủ lạnh, che đậy kín tránh ruồi muỗi, tách biệt thực phẩm sống và thực phẩm chín).\n  + Rửa sạch tay bằng xà phòng trước khi chế biến thức ăn và trước khi ăn; vệ sinh sạch sẽ dao, thớt, dụng cụ nhà bếp."
    }
]
