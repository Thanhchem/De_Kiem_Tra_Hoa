# -*- coding: utf-8 -*-
"""
Dữ liệu Đề Cương Ôn Tập Giữa Kì I KHTN 7 - Trường THCS Nguyễn Thị Minh Khai
PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG - SAI (CÂU 1 -> CÂU 2)
PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (CÂU 1 -> CÂU 38)
"""

# -----------------------------------------------------------------------------
# PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG - SAI
# -----------------------------------------------------------------------------
TF_QUESTIONS = [
    {
        "id": 1,
        "content": "Tại phòng thực hành Khoa học tự nhiên của trường THCS Nguyễn Thị Minh Khai, nhóm của bạn An đang tiến hành đo tốc độ của một chiếc ô tô đồ chơi chạy trên mặt dốc. Để tìm hiểu và so sánh ưu, nhược điểm của các phương pháp đo lường, cô giáo đã yêu cầu nhóm sử dụng luân phiên hai bộ dụng cụ: bộ cơ bản gồm thước cuộn và đồng hồ bấm giây tay; bộ hiện đại gồm cổng quang điện kết nối với đồng hồ đo thời gian hiện số.\nDựa vào kiến thức về cách đo tốc độ, hãy xác định các nhận định sau đúng hay sai?",
        "statements": [
            {
                "id": "a",
                "text": "Để đo tốc độ chuyển động bằng đồng hồ bấm giây, ta chỉ cần đo một đại lượng duy nhất là thời gian chuyển động của vật.",
                "correct": False,
                "explanation": "Sai. Để đo tốc độ theo công thức v = s / t, bắt buộc phải đo hai đại lượng: quãng đường s (bằng thước) và thời gian t (bằng đồng hồ)."
            },
            {
                "id": "b",
                "text": "Trong thí nghiệm đo tốc độ ô tô đồ chơi chạy trên mặt dốc bằng đồng hồ bấm giây, người ta thường thực hiện phép đo 3 lần nhằm mục đích tính giá trị trung bình để giảm sai số.",
                "correct": True,
                "explanation": "Đúng. Đo nhiều lần (ít nhất 3 lần) rồi tính giá trị trung bình là nguyên tắc khoa học giúp triệt tiêu sai số ngẫu nhiên do phản xạ bấm đồng hồ của người đo."
            },
            {
                "id": "c",
                "text": "Khi đo tốc độ bằng cổng quang điện và đồng hồ đo thời gian hiện số, việc mở và tắt đồng hồ đếm thời gian được điều khiển hoàn toàn tự động khi vật chắn tia hồng ngoại đi qua cổng quang.",
                "correct": True,
                "explanation": "Đúng. Chùm tia hồng ngoại nối giữa D1 và D2 đóng vai trò công tắc cảm biến tự động: khi vật chắn tia sáng, tín hiệu điện lập tức điều khiển đồng hồ bật hoặc ngắt đếm, loại bỏ hoàn toàn sai số chủ quan của con người."
            },
            {
                "id": "d",
                "text": "Khi cài đặt đồng hồ đo thời gian hiện số ở chế độ A ↔ B, đồng hồ sẽ bắt đầu đếm thời gian khi vật qua cổng quang A và ngắt đếm khi vật qua cổng quang B.",
                "correct": True,
                "explanation": "Đúng. Chế độ A ↔ B được thiết kế chuyên dụng để đo thời gian chuyển động của vật từ vị trí cổng quang A đến vị trí cổng quang B."
            }
        ],
        "explanation": """- Ý a SAI: Công thức tốc độ v = s / t yêu cầu phải đo cả quãng đường s và thời gian t.
- Ý b ĐÚNG: Đo 3 lần lấy giá trị trung bình giúp giảm thiểu sai số đo lường.
- Ý c ĐÚNG: Cổng quang điện đóng/ngắt tín hiệu tự động khi vật chắn tia hồng ngoại.
- Ý d ĐÚNG: Chế độ A ↔ B dùng để đo thời gian vật di chuyển giữa cổng quang A và cổng quang B."""
    },
    {
        "id": 2,
        "content": "Để nâng cao năng suất rau xà lách thủy canh, một trang trại nông nghiệp công nghệ cao đã áp dụng mô hình nhà kính. Người nông dân tiến hành thắp đèn LED chiếu sáng vào ban đêm, duy trì hệ thống phun sương liên tục và lắp đặt hệ thống quạt thông gió để không khí được luân chuyển. Dựa vào kiến thức về quang hợp, hãy cho biết các phát biểu sau đây là đúng hay sai?",
        "statements": [
            {
                "id": "a",
                "text": "Khí oxygen và năng lượng ánh sáng là hai nguyên liệu bắt buộc mà cây rau cần lấy từ môi trường để thực hiện quá trình quang hợp.",
                "correct": False,
                "explanation": "Sai. Nguyên liệu của quang hợp là khí carbon dioxide (CO~2~) và nước (H~2~O); năng lượng ánh sáng là điều kiện/nguồn năng lượng; còn khí oxygen là sản phẩm thải ra."
            },
            {
                "id": "b",
                "text": "Trong nhà kính, lá cây hấp thụ năng lượng ánh sáng từ hệ thống đèn LED (quang năng) để tổng hợp nên chất hữu cơ.",
                "correct": True,
                "explanation": "Đúng. Hệ sắc tố quang hợp (chất diệp lục) có thể hấp thụ ánh sáng từ đèn nhân tạo (đặc biệt là ánh sáng đỏ và xanh lam của đèn LED chuyên dụng) để thực hiện quá trình quang hợp bình thường."
            },
            {
                "id": "c",
                "text": "Hệ thống phun sương cung cấp nước - một yếu tố quan trọng vừa làm nguyên liệu tham gia trực tiếp vào quang hợp, vừa giúp điều tiết độ mở của khí khổng.",
                "correct": True,
                "explanation": "Đúng. Nước vừa là nguyên liệu trực tiếp cấu tạo nên glucose, vừa tạo áp suất trương nước để khí khổng mở ra cho CO~2~ khuếch tán vào lá, đồng thời điều hòa nhiệt độ bề mặt lá."
            },
            {
                "id": "d",
                "text": "Việc bật quạt thông gió chỉ nhằm mục đích hạ nhiệt độ trong nhà kính, hoàn toàn không ảnh hưởng đến cường độ quang hợp của cây rau.",
                "correct": False,
                "explanation": "Sai. Quạt thông gió lưu chuyển không khí giúp cung cấp và bổ sung liên tục nguồn khí carbon dioxide (CO~2~) quanh tán lá, chống cạn kiệt CO~2~ cục bộ, do đó tác động trực tiếp và gia tăng cường độ quang hợp."
            }
        ],
        "explanation": """- Ý a SAI: Nguyên liệu của quang hợp là CO~2~ và H~2~O (oxygen là sản phẩm giải phóng).
- Ý b ĐÚNG: Cây có khả năng sử dụng quang năng từ đèn LED để quang hợp ban đêm.
- Ý c ĐÚNG: Nước là nguyên liệu trực tiếp và giúp tế bào khí khổng mở để trao đổi khí.
- Ý d SAI: Quạt thông gió giúp lưu thông không khí, bổ sung CO~2~ cần thiết cho quang hợp."""
    }
]

# -----------------------------------------------------------------------------
# PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (CÂU 1 -> CÂU 38)
# -----------------------------------------------------------------------------
SHORT_QUESTIONS = [
    {
        "id": 1,
        "content": "Khi vào phòng thực hành đo khối lượng của cuốn sách Khoa học tự nhiên 7 bằng cân điện tử, các học sinh đã áp dụng “Kĩ năng đo”. Cho các thao tác sau:\n(1) Bỏ qua việc lựa chọn thang đo và cứ thế đặt cuốn sách lên bất kỳ chiếc cân nào có trong phòng.\n(2) Cầm cuốn sách lên tay để ước lượng khoảng khối lượng của nó, từ đó lựa chọn loại cân có giới hạn đo phù hợp.\n(3) Thực hiện phép đo, đọc và ghi lại số liệu hiển thị trên màn hình cân.\n(4) Tin tưởng tuyệt đối vào 1 lần đo duy nhất, bỏ qua sai số do thao tác đặt vật.\n(5) Nhận xét độ chính xác của kết quả đo căn cứ vào loại dụng cụ đo (độ chia nhỏ nhất của cân điện tử).\n(6) Tiến hành cân lặp lại 3 lần để lấy giá trị trung bình giúp kết quả chính xác hơn.\nNhững thao tác nào là thao tác đúng chuẩn của kĩ năng đo đạc trong khoa học? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các thao tác ở trên, ví dụ 123456)",
        "answer": "2356",
        "explanation": "Kĩ năng đo chuẩn mực gồm: (2) Ước lượng để chọn dụng cụ có GHĐ và ĐCNN thích hợp; (3) Thực hiện phép đo, đọc và ghi kết quả đúng quy cách; (5) Đánh giá độ chính xác theo ĐCNN; (6) Đo lặp lại 3 lần lấy giá trị trung bình. Thao tác (1) và (4) là sai lầm."
    },
    {
        "id": 2,
        "content": "Độ chia nhỏ nhất (ĐCNN) của đồng hồ đo thời gian hiện số khi chọn thang đo 99,99 s là bao nhiêu giây?",
        "answer": "0,01",
        "explanation": "Thang đo 99,99 s hiển thị thời gian chính xác đến hàng phần trăm của giây, nghĩa là giá trị nhỏ nhất thay đổi được trên màn hình là 0,01 s."
    },
    {
        "id": 3,
        "content": "Một học sinh đo chiều dài cuốn sách KHTN 7 ba lần thu được các kết quả: l1 = 26,4 cm; l2 = 26,5 cm; l3 = 26,6 cm. Giá trị trung bình chiều dài cuốn sách theo đơn vị xentimét (cm) là bao nhiêu?",
        "answer": "26,5",
        "explanation": "Giá trị trung bình: l_tb = (l1 + l2 + l3) / 3 = (26,4 + 26,5 + 26,6) / 3 = 79,5 / 3 = 26,5 cm."
    },
    {
        "id": 4,
        "content": "Học sinh thực hiện đo khối lượng cuốn sách KHTN 7 ba lần thu được số liệu: m1 = 300 g; m2 = 302 g; m3 = 298 g. Giá trị trung bình khối lượng cuốn sách tính theo đơn vị gam (g) là bao nhiêu?",
        "answer": "300",
        "explanation": "Giá trị trung bình: m_tb = (m1 + m2 + m3) / 3 = (300 + 302 + 298) / 3 = 900 / 3 = 300 g."
    },
    {
        "id": 5,
        "content": "Một viên bi sắt chuyển động từ cổng quang điện A đến cổng quang điện B cách nhau 0,6 m. Đồng hồ đo thời gian hiện số hiển thị t = 5,00 s. Tốc độ trung bình của viên bi sắt theo đơn vị m/s là bao nhiêu?",
        "answer": "0,12",
        "explanation": "Tốc độ trung bình của viên bi: v = s / t = 0,6 / 5,00 = 0,12 m/s."
    },
    {
        "id": 6,
        "content": "Bạn Nam đếm được 1212 bước đi từ nhà đến trường, mỗi bước dài 0,5 m, thời gian đi là 10 phút. Tốc độ đi bộ của bạn Nam tính theo đơn vị km/h (làm tròn đến 2 chữ số thập phân) là bao nhiêu?",
        "answer": "3,64",
        "explanation": "Quãng đường Nam đi: s = 1212 × 0,5 = 606 m = 0,606 km.\nThời gian đi: t = 10 phút = 10 / 60 h = 1 / 6 h (hoặc 600 s).\nTốc độ: v = s / t = 0,606 / (1/6) = 0,606 × 6 = 3,636 km/h ≈ 3,64 km/h."
    },
    {
        "id": 7,
        "content": "Một vận động viên bơi lội vượt qua đường bơi có chiều dài 100 m hết thời gian 50 s. Tốc độ bơi trung bình của vận động viên là bao nhiêu mét trên giây (m/s)?",
        "answer": "2",
        "explanation": "Tốc độ bơi trung bình: v = s / t = 100 / 50 = 2 m/s."
    },
    {
        "id": 8,
        "content": "Một đoàn tàu hoả chuyển động đều trên đường ray với tốc độ 72km/h.Thời gian để đoàn tàu đi hết quãng đường dài 36 km là bao nhiêu giờ (h)?",
        "answer": "0,5",
        "explanation": "Thời gian đoàn tàu đi: t = s / v = 36 / 72 = 0,5 h."
    },
    {
        "id": 9,
        "content": "Thiết bị bắn tốc độ đo được thời gian một xe máy chạy qua hai vạch mốc cách nhau 6m trên mặt đường là 0,3 s. Tốc độ của xe máy đó bằng bao nhiêu kilômét trên giờ (km/h)",
        "answer": "72",
        "explanation": "Tốc độ của xe máy tính theo m/s: v = s / t = 6 / 0,3 = 20 m/s.\nĐổi sang km/h: v = 20 × 3,6 = 72 km/h."
    },
    {
        "id": 10,
        "content": "Trong thí nghiệm đo tốc độ của xe lăn trên máng dốc bằng đồng hồ bấm giây, học sinh đo quãng đường chuyển động là s = 0,9 m. Ba lần đo thời gian xe chuyển động thu được kết quả: 1,4 s; 1,6 s; 1,5 s. Tốc độ trung bình của xe lăn tính theo số liệu trên bằng bao nhiêu mét trên giây (m/s).",
        "answer": "0,6",
        "explanation": "Thời gian trung bình: t_tb = (1,4 + 1,6 + 1,5) / 3 = 1,5 s.\nTốc độ trung bình của xe lăn: v = s / t_tb = 0,9 / 1,5 = 0,6 m/s."
    },
    {
        "id": 11,
        "content": "Trong giờ thực hành Khoa học tự nhiên 7, nhóm học sinh dùng thước đo được quãng đường một xe đồ chơi đi được là 6 m. Đồng hồ bấm giây cho biết xe đi hết quãng đường đó trong 3 s.",
        "answer": "2",
        "explanation": "Tốc độ của xe đồ chơi: v = s / t = 6 / 3 = 2 m/s."
    },
    {
        "id": 12,
        "content": "Trong buổi thực hành tìm hiểu tốc độ chuyển động, một nhóm học sinh cho xe đồ chơi chuyển động trên quãng đường dài 120 cm. Ba lần đo thời gian xe đi hết quãng đường lần lượt là 2,4 s; 2,5 s và 2,6 s. Khi làm báo cáo, nhóm lấy thời gian trung bình của ba lần đo để tính tốc độ trung bình của xe. Tính tốc độ trung bình của xe bằng bao nhiêu m/s? (Làm tròn đến hàng đơn vị).",
        "answer": "0,48",
        "explanation": "Quãng đường: s = 120 cm = 1,2 m.\nThời gian trung bình: t_tb = (2,4 + 2,5 + 2,6) / 3 = 2,5 s.\nTốc độ trung bình: v = 1,2 / 2,5 = 0,48 m/s (nếu tính theo cm/s là 48 cm/s; nếu làm tròn đến hàng đơn vị của m/s là 0 m/s)."
    },
    {
        "id": 13,
        "content": "Hãy sắp xếp thứ tự đúng cho 4 bước đầu tiên của phương pháp tìm hiểu tự nhiên:\n(1) Đề xuất giả thuyết\n(2) Thực hiện kế hoạch\n(3) Quan sát, đặt câu hỏi\n(4) Lập kế hoạch kiểm tra\n(Ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự đúng của tiến trình trên, ví dụ 123456)",
        "answer": "3142",
        "explanation": "Thứ tự tiến trình nghiên cứu khoa học: (3) Quan sát, đặt câu hỏi ──> (1) Đề xuất giả thuyết ──> (4) Lập kế hoạch kiểm tra ──> (2) Thực hiện kế hoạch. Dãy số thứ tự là 3142."
    },
    {
        "id": 14,
        "content": "Cho danh sách các kí hiệu:\n(1) Ca        (2) C        (3) Cl        (4) Cu\nHãy sắp xếp các số tương ứng với thứ tự tên các nguyên tố hóa học sau: Carbon, Chlorine, Calcium, Copper.",
        "answer": "2314",
        "explanation": "Carbon là C (2); Chlorine là Cl (3); Calcium là Ca (1); Copper là Cu (4). Dãy số tương ứng: 2314."
    },
    {
        "id": 15,
        "content": "Cho các kí hiệu/nguyên tố tương ứng:\n(1) K        (2) P        (3) Cu        (4) Ca\nHãy viết theo thứ tự chuỗi 4 ký tự viết tắt đại diện cho các nguyên tố lần lượt là Potassium, Phosphorus, Calcium, Copper.",
        "answer": "1243",
        "explanation": "Potassium: K (1); Phosphorus: P (2); Calcium: Ca (4); Copper: Cu (3). Chuỗi số đại diện thứ tự: 1243."
    },
    {
        "id": 16,
        "content": "Cho các nguyên tố phi kim và kim loại:\n(1) Oxygen        (2) Iron        (3) Nitrogen        (4) Aluminium\nHãy sắp xếp chuỗi đại diện theo thứ tự: Phi kim (O) -> Phi kim (N) -> Kim loại (Al) -> Kim loại (Fe).",
        "answer": "1342",
        "explanation": "Oxygen là (1), Nitrogen là (3), Aluminium là (4), Iron là (2). Thứ tự chuỗi số: 1342."
    },
    {
        "id": 17,
        "content": "Khi chuẩn bị báo cáo thuyết trình kết quả thực hành trước lớp, nhóm học sinh thảo luận về các giải pháp hỗ trợ sau:\n(1) Trình bày toàn bộ văn bản chi tiết lên màn chiếu để đọc lại cho lớp nghe.\n(2) Xây dựng bảng biểu, biểu đồ trực quan để tóm tắt số liệu kết quả.\n(3) Sử dụng hình ảnh/video ngắn ghi lại quá trình nhóm thao tác thí nghiệm thực tế.\n(4) Chỉ thuyết trình bằng lời nói với âm lượng to mà không cần công cụ minh họa.\n(5) Chuẩn bị một số câu hỏi ngắn tương tác với người nghe ở phần thảo luận.\n(6) Bỏ qua phần nhận xét, kết luận để tiết kiệm thời gian thuyết trình.\nNhững giải pháp nào ở trên giúp bài báo cáo thuyết trình trở nên khoa học, trực quan và hấp dẫn mà các nhóm nên ưu tiên sử dụng? (Ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải, ví dụ 123456)",
        "answer": "235",
        "explanation": "Các giải pháp khoa học và trực quan cần ưu tiên: (2) Sử dụng biểu bảng; (3) Video/hình ảnh thực tế; (5) Câu hỏi tương tác. Dãy số tăng dần: 235."
    },
    {
        "id": 18,
        "content": "Một nhóm học sinh chuẩn bị bài Báo cáo thực hành đo khối lượng một vật và lập danh sách các phần nội dung sau:\n(1) Liệt kê tên dụng cụ đo (cân điện tử) và mẫu vật cần đo.\n(2) Thực hiện phép tính khối lượng trung bình sau 3 lần đo.\n(3) Mô tả các bước đặt vật lên cân và đọc chỉ số.\n(4) Trả lời các câu hỏi thảo luận cùng cố cuối bài.\n(5) Xác định rõ mục tiêu cần đạt được của bài thực hành.\n(6) Lập bảng ghi lại kết quả số liệu thu được.\nNhững phần nội dung nào ở trên thuộc các mục cần trình bày trước mục kết quả trong mẫu báo cáo thực hành? (Ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải, ví dụ 123456)",
        "answer": "135",
        "explanation": "Cấu trúc báo cáo trước phần Kết quả gồm: (5) Mục đích thực hành, (1) Dụng cụ và mẫu vật chuẩn bị, (3) Các bước tiến hành. Dãy số tăng dần: 135."
    },
    {
        "id": 19,
        "content": "Trong mẫu Báo cáo thực hành môn Khoa học tự nhiên 7, có các mục nội dung sau:\n(1) Chuẩn bị.\n(2) Kết quả.\n(3) Các bước tiến hành.\n(4) Mục đích thí nghiệm.\nHãy sắp xếp các mục trên theo đúng thứ tự cấu trúc chuẩn của mẫu báo cáo thực hành. (Ghi đáp án dưới dạng một dãy chữ số liền nhau ứng với thứ tự sắp xếp đúng)",
        "answer": "4132",
        "explanation": "Cấu trúc chuẩn của báo cáo thực hành: (4) Mục đích thí nghiệm ──> (1) Chuẩn bị ──> (3) Các bước tiến hành ──> (2) Kết quả. Chuỗi số: 4132."
    },
    {
        "id": 20,
        "content": "Trao đổi chất là quá trình sinh vật lấy các chất từ môi trường, biến đổi chúng và trả lại môi trường các chất thải. Cho các hiện tượng sau:\n(1) Con người hít khí oxygen và thở ra khí carbon dioxide.\n(2) Hòn đá bị bào mòn bởi dòng nước chảy qua hàng trăm năm.\n(3) Cây xanh hút nước và muối khoáng từ đất.\n(4) Sự bốc hơi nước của mặt hồ dưới ánh nắng mặt trời.\n(5) Cơ thể sinh vật tiết mồ hôi, đào thải chất thải ra môi trường.\n(6) Nước đóng băng thành đá khi nhiệt độ hạ thấp dưới 0°C.\nNhững hiện tượng nào thể hiện đúng bản chất của quá trình trao đổi chất ở cơ thể sinh vật? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các hiện tượng ở trên, ví dụ 123456)",
        "answer": "135",
        "explanation": "Hiện tượng trao đổi chất ở sinh vật: (1) Hô hấp ở người; (3) Hút nước và khoáng ở cây; (5) Bài tiết mồ hôi. Các mục (2), (4), (6) là biến đổi vật lí của vật không sống. Dãy số: 135."
    },
    {
        "id": 21,
        "content": "Chuyển hóa năng lượng ở sinh vật là sự biến đổi năng lượng từ dạng này sang dạng khác. Cho các quá trình sau:\n(1) Quạt điện quay khi được cắm vào nguồn điện lưới.\n(2) Lá cây hấp thụ ánh sáng mặt trời để tổng hợp chất hữu cơ.\n(3) Thanh sắt bị gỉ sét khi để ngoài không khí ẩm lâu ngày.\n(4) Năng lượng trong thức ăn được biến đổi giúp cơ bắp co giãn khi con báo chạy săn mồi.\n(5) Nước đun sôi tỏa hơi nước trên bếp gas.\n(6) Cơ thể sinh vật phân giải chất hữu cơ tỏa ra nhiệt năng làm ấm cơ thể.\nNhững quá trình nào mô tả đúng sự chuyển hóa năng lượng diễn ra bên trong cơ thể sinh vật? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các quá trình ở trên, ví dụ 123456)",
        "answer": "246",
        "explanation": "Sự chuyển hóa năng lượng bên trong cơ thể sinh vật: (2) Quang năng thành hóa năng; (4) Hóa năng thành cơ năng; (6) Hóa năng thành nhiệt năng. Dãy số tăng dần: 246."
    },
    {
        "id": 22,
        "content": "Trao đổi chất và chuyển hóa năng lượng có vai trò đảm bảo cho sinh vật tồn tại. Cho các vai trò sau:\n(1) Cung cấp nguyên liệu để xây dựng tế bào và cơ thể.\n(2) Tạo ra năng lượng dễ sử dụng cung cấp cho các hoạt động sống (cảm ứng, vận động).\n(3) Giúp sinh vật có khả năng miễn nhiễm tự nhiên với mọi loại dịch bệnh.\n(4) Giúp cơ thể sinh vật sinh trưởng, phát triển và sinh sản.\n(5) Làm cho cơ thể sinh vật có thể tồn tại vĩnh viễn không bao giờ chết.\n(6) Giúp sinh vật giữ nguyên kích thước và khối lượng từ lúc sinh ra đến lúc già đi.\nNhững vai trò nào đúng với ý nghĩa của quá trình trao đổi chất và chuyển hóa năng lượng? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các vai trò ở trên, ví dụ 123456)",
        "answer": "124",
        "explanation": "Ba vai trò cốt lõi của trao đổi chất và chuyển hóa năng lượng: (1) Cung cấp vật chất xây dựng cơ thể; (2) Cung cấp năng lượng (ATP) cho hoạt động sống; (4) Đảm bảo sự sinh trưởng, phát triển và sinh sản. Dãy số: 124."
    },
    {
        "id": 23,
        "content": "Đặc điểm và mối quan hệ giữa trao đổi chất và chuyển hóa năng lượng luôn rất chặt chẽ. Cho các phát biểu sau:\n(1) Trao đổi chất luôn gắn liền với chuyển hoá năng lượng.\n(2) Sinh vật lấy các chất từ môi trường và biến đổi chúng thành các chất cần thiết cho cơ thể.\n(3) Khi quá trình trao đổi chất dừng lại, sinh vật vẫn có thể tiếp tục sống bình thường.\n(4) Mọi cơ thể sống đều không ngừng trao đổi chất và chuyển hoá năng lượng với môi trường.\n(5) Các chất thải sau khi được sinh vật biến đổi sẽ được tích trữ vĩnh viễn trong cơ thể mà không trả lại môi trường.\n(6) Quá trình chuyển hóa năng lượng có thể diễn ra hoàn toàn độc lập, không liên quan đến trao đổi chất.\nNhững phát biểu nào là đúng? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các phát biểu ở trên, ví dụ 123456)",
        "answer": "124",
        "explanation": "Các phát biểu đúng: (1), (2), (4). Ý (3) sai vì trao đổi chất dừng thì sinh vật chết; ý (5) sai vì chất thải phải được đào thải ra ngoài; ý (6) sai vì chuyển hóa năng lượng luôn gắn liền với trao đổi chất. Dãy số: 124."
    },
    {
        "id": 24,
        "content": "Quá trình trao đổi chất và chuyển hóa năng lượng là đặc trưng cơ bản phân biệt vật sống và vật không sống, diễn ra không ngừng để đảm bảo sự tồn tại của sinh vật. Cho các đối tượng và hiện tượng sau:\n(1) Hạt đậu hút nước, nảy mầm và phát triển thành cây non.\n(2) Tượng đá ngoài sân bị nứt vỡ, bào mòn do nhiệt độ và nước mưa.\n(3) Người chạy bộ toát nhiều mồ hôi và nhịp thở tăng lên.\n(4) Động cơ xe máy đốt cháy xăng để tạo ra lực đẩy xe đi.\n(5) Cây bàng rụng lá vào mùa đông để giảm sự thoát hơi nước.\n(6) Khối than đá cháy trong lò tỏa ra lượng nhiệt lớn.\nNhững hiện tượng nào là biểu hiện của quá trình trao đổi chất và chuyển hoá năng lượng ở sinh vật? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các hiện tượng ở trên, ví dụ 123456)",
        "answer": "135",
        "explanation": "Hiện tượng ở sinh vật sống: (1) Sự nảy mầm của hạt; (3) Hoạt động hô hấp và bài tiết ở người; (5) Sự điều hòa thoát hơi nước ở thực vật. Dãy số: 135."
    },
    {
        "id": 25,
        "content": "Khi chạy bộ, cơ thể có cảm giác nóng lên, mồ hôi ra nhiều và nhịp thở tăng. Để giải thích các hiện tượng này dựa trên khái niệm trao đổi chất và chuyển hóa năng lượng, các học sinh đưa ra các nhận định sau:\n(1) Nhịp thở tăng để cơ thể lấy thêm nhiều oxygen và thải nhanh khí carbon dioxide ra môi trường.\n(2) Cơ thể nóng lên là do một phần hoá năng chuyển hoá thành nhiệt năng.\n(3) Đổ mồ hôi là một hình thức cơ thể trả lại môi trường các chất thải và giúp ổn định nhiệt độ.\n(4) Quá trình vận động (chạy) chỉ tiêu hao năng lượng mà không liên quan gì đến sự trao đổi chất.\n(5) Việc có biểu hiện khát nước sau khi chạy chứng tỏ cơ thể đã ngừng hoàn toàn quá trình trao đổi chất.\n(6) Năng lượng dùng để chạy (cơ năng) được biến đổi từ hoá năng tích luỹ trong các chất hữu cơ.\nNhững nhận định nào giải thích đúng bản chất của các hiện tượng trên? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các nhận định ở trên, ví dụ 123456)",
        "answer": "1236",
        "explanation": "Các nhận định chính xác: (1) Tăng hô hấp lấy O~2~ và thải CO~2~; (2) Hóa năng một phần tỏa ra thành nhiệt năng; (3) Tiết mồ hôi thải nhiệt và chất thải; (6) Hóa năng từ thức ăn biến thành cơ năng. Dãy số: 1236."
    },
    {
        "id": 26,
        "content": "Khi tiến hành thí nghiệm chứng minh quang hợp tạo ra tinh bột ở lá cây, học sinh cần thực hiện các thao tác theo một quy trình chuẩn. Cho các thao tác sau:\n(1) Để chậu cây trồng (ví dụ: khoai lang, tía tô) vào chỗ tối 2 ngày.\n(2) Dùng băng giấy đen bịt kín một phần lá ở cả mặt trên và mặt dưới.\n(3) Đem chậu cây ra đặt ở chỗ có nắng gắt chỉ khoảng 10 - 15 phút.\n(4) Hái chiếc lá đã bịt băng đen, tháo giấy và đun cách thuỷ lá trong cồn 90°.\n(5) Đun lá trực tiếp trên ngọn lửa đèn cồn để rút ngắn thời gian.\n(6) Rửa sạch lá bằng nước ấm rồi nhỏ vài giọt dung dịch iodine lên mặt lá.\nHãy sắp xếp các thao tác sau theo đúng trình tự?",
        "answer": "1246",
        "explanation": "Trình tự chuẩn các thao tác đúng: (1) Để cây vào chỗ tối 2 ngày ──> (2) Bịt băng giấy đen một phần lá (sau đó đem ra nắng 4-6h) ──> (4) Đun cách thuỷ lá trong cồn 90° ──> (6) Rửa sạch lá và nhỏ thuốc thử iodine. Chuỗi số: 1246."
    },
    {
        "id": 27,
        "content": "Sau khi nhỏ dung dịch iodine lên chiếc lá đã được tẩy hết diệp lục trong thí nghiệm chứng minh quang hợp tạo ra tinh bột, học sinh sẽ quan sát thấy những hiện tượng nào? Cho các hiện tượng sau:\n(1) Phần lá không bịt băng giấy đen chuyển sang màu xanh tím.\n(2) Phần lá bị bịt băng giấy đen giữ nguyên màu vàng nâu của dung dịch iodine.\n(3) Cồn trong cốc đun cách thuỷ trước đó đã chuyển dần sang màu xanh lục.\n(4) Phần lá không bịt băng giấy đen hoàn toàn không đổi màu.\n(5) Phần lá bị bịt băng giấy đen chuyển sang màu xanh tím rất đậm.\n(6) Toàn bộ bề mặt lá đều nổi bọt khí liti khi tiếp xúc với dung dịch iodine.\nNhững hiện tượng nào quan sát được là đúng với thực tế thí nghiệm? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các hiện tượng ở trên, ví dụ 123456)",
        "answer": "123",
        "explanation": "Hiện tượng đúng thực tế: (1) Phần được chiếu sáng quang hợp tạo tinh bột nên hóa màu xanh tím; (2) Phần bịt giấy đen không có tinh bột nên giữ màu vàng nâu của iot; (3) Cồn nóng hòa tan diệp lục nên hóa màu xanh lục. Dãy số: 123."
    },
    {
        "id": 28,
        "content": "Để tiến hành thí nghiệm chứng minh quang hợp thải khí oxygen, nhóm của Minh sử dụng cành rong đuôi chó. Cho các thao tác lắp đặt sau:\n(1) Đặt cốc A ra chỗ có ánh nắng mặt trời (hoặc dưới đèn), cốc B để trong bóng tối.\n(2) Sau 6 giờ, nhẹ nhàng rút 2 cành rong ra và bịt kín ống nghiệm, lấy ra khỏi 2 cốc rồi lật ngược lại. Đưa nhanh que đóm còn tàn đỏ vào miệng mỗi ống nghiệm.\n(3) Đổ đầy nước vào 2 ống nghiệm, dùng ngón tay bịt miệng ống nghiệm rồi úp ngược mỗi ống nghiệm vào cốc nước (cốc A và cốc B).\n(4) Đặt vài cành rong đuôi chó vào 2 ống nghiệm sao cho phần ngọn rong dưới đáy ống nghiệm.\nHãy sắp xếp những thao tác của nhóm Minh theo đúng trình tự thí nghiệm?",
        "answer": "4312",
        "explanation": "Trình tự thao tác chuẩn: (4) Đặt cành rong vào ống nghiệm ──> (3) Đổ đầy nước và úp ngược vào cốc ──> (1) Đặt cốc A ra sáng, cốc B vào tối ──> (2) Thử que đóm tàn đỏ sau 6 giờ. Chuỗi số: 4312."
    },
    {
        "id": 29,
        "content": "Khi quan sát thí nghiệm với cành rong đuôi chó (cốc A để ngoài sáng, cốc B để trong tối) sau một khoảng thời gian, học sinh sẽ thấy những kết quả nào dưới đây? Cho các kết quả sau:\n(1) Ở cốc A (ngoài sáng), có nhiều bọt khí nhỏ nổi lên từ cành rong và tụ lại ở đáy ống nghiệm.\n(2) Nước trong ống nghiệm ở cốc A bị phần chất khí sinh ra đẩy dần xuống dưới.\n(3) Ở cốc B (trong tối), cành rong cũng sinh ra lượng bọt khí nhiều tương đương cốc A.\n(4) Khi đưa tàn đóm đỏ vào gần miệng ống nghiệm của cốc A (vừa lật ngửa lên), tàn đóm bùng cháy.\n(5) Mực nước trong ống nghiệm của cốc B bị đẩy xuống nhanh hơn cốc A.\n(6) Bọt khí tích tụ trong ống nghiệm ở cốc A chính là khí carbon dioxide.\nNhững hiện tượng và kết quả nào là đúng? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các kết quả ở trên, ví dụ 123456)",
        "answer": "124",
        "explanation": "Kết quả đúng thực tế: (1) Cốc A có bọt khí nổi lên tụ lại ở đáy ống nghiệm; (2) Khí sinh ra đẩy nước xuống dưới; (4) Khí oxygen làm tàn đóm bùng cháy. Các mục (3), (5), (6) là sai. Dãy số: 124."
    },
    {
        "id": 30,
        "content": "Khi tổng kết bài thực hành chứng minh quang hợp, giáo viên yêu cầu học sinh giải thích nguyên tắc khoa học và an toàn của các bước. Cho các phát biểu sau:\n(1) Phải đun cách thuỷ lá trong cồn vì cồn rất dễ bắt lửa, đun trực tiếp dễ gây hoả hoạn.\n(2) Đun trong cồn 90° giúp hoà tan chất diệp lục, làm lá mất màu để dễ quan sát sự đổi màu của thuốc thử.\n(3) Việc để cây trong tối 2 ngày giúp lá tiêu hao hết lượng tinh bột đã được tổng hợp từ trước.\n(4) Băng giấy đen có tác dụng ngăn ánh sáng, làm phần lá bị bịt không thể quang hợp tạo tinh bột để làm mẫu đối chứng.\n(5) Tàn đóm đỏ bùng cháy là đặc điểm đặc trưng để nhận biết khí oxygen (vì oxygen duy trì sự cháy).\n(6) Dung dịch iodine đổi màu xanh tím là do phản ứng đặc hiệu với khí oxygen hòa tan.\nNhững phát biểu nào giải thích đúng bản chất của các thao tác thực hành? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các phát biểu ở trên, ví dụ 123456)",
        "answer": "12345",
        "explanation": "Các phát biểu (1), (2), (3), (4), (5) đều giải thích hoàn toàn chính xác nguyên tắc khoa học và an toàn. Phát biểu (6) sai vì iodine chỉ phản ứng tạo màu xanh tím đặc trưng với tinh bột. Dãy số: 12345."
    },
    {
        "id": 31,
        "content": "Để chuẩn bị mẫu vật cho các thí nghiệm chứng minh hô hấp tế bào, học sinh cần tiến hành ủ hạt nảy mầm. Cho các thao tác sau:\n(1) Ngâm hạt (đậu xanh, đậu đen...) trong nước ấm khoảng 40°C để kích thích hạt dễ nảy mầm.\n(2) Đem luộc chín toàn bộ số hạt trước khi cho vào bình thí nghiệm để hạt phình to, dễ quan sát.\n(3) Ủ hạt trong bông ẩm hoặc khăn giấy ẩm ở nhiệt độ phòng cho đến khi hạt nhú mầm.\n(4) Chia số hạt đã nảy mầm thành 2 phần: một phần để sống, một phần đem luộc chín làm mẫu đối chứng.\n(5) Phơi khô hạt nảy mầm ngoài nắng gắt 3 ngày cho hạt cứng lại trước khi làm thí nghiệm.\n(6) Cho các phần hạt vào các bình thuỷ tinh/bình tam giác và đậy thật kín nắp nắp lại.\nNhững thao tác nào là thao tác chuẩn bị ĐÚNG trong quy trình làm thí nghiệm? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các thao tác ở trên, ví dụ 123456)",
        "answer": "1346",
        "explanation": "Thao tác chuẩn bị đúng quy trình: (1) Ngâm nước ấm kích thích nảy mầm; (3) Ủ hạt ẩm đến khi nhú mầm; (4) Chia làm 2 phần (sống làm thí nghiệm, luộc chín làm đối chứng); (6) Cho vào bình đậy kín nắp. Dãy số: 1346."
    },
    {
        "id": 32,
        "content": "Khi tiến hành thí nghiệm chứng minh hô hấp tế bào thải ra khí carbon dioxide, nhóm của An sử dụng nước vôi trong. Cho các bước tiến hành và hiện tượng sau:\n(1) Đặt một cốc nhỏ chứa nước vôi trong vào bình đựng hạt đang nảy mầm rồi đậy kín nắp.\n(2) Đặt một cốc nước vôi trong vào bình đựng hạt đã luộc chín rồi đậy thật kín nắp để làm đối chứng.\n(3) Sau khoảng 1 - 2 giờ, cốc nước vôi trong ở bình chứa hạt nảy mầm sẽ xuất hiện lớp váng đục trên bề mặt.\n(4) Nước vôi trong ở bình hạt luộc chín cũng bị vẩn đục y hệt như bình hạt nảy mầm.\n(5) Hiện tượng vẩn đục chứng tỏ hạt nảy mầm đã hút hết toàn bộ lượng nước trong cốc.\n(6) Lớp vẩn đục xuất hiện là do khí carbon dioxide từ quá trình hô hấp của hạt tác dụng với nước vôi trong.\nNhững phát biểu nào mô tả đúng các bước và hiện tượng của thí nghiệm này? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các phát biểu ở trên, ví dụ 123456)",
        "answer": "1236",
        "explanation": "Các phát biểu đúng: (1) Cốc nước vôi trong trong bình hạt sống; (2) Bình hạt chín làm đối chứng; (3) Bình hạt sống xuất hiện váng đục; (6) Váng đục CaCO~3~ do CO~2~ hô hấp phản ứng với Ca(OH)~2~. Dãy số: 1236."
    },
    {
        "id": 33,
        "content": "Để chứng minh hô hấp tế bào tiêu thụ khí oxygen, giáo viên hướng dẫn học sinh dùng que diêm (hoặc nến đang cháy). Cho các thao tác và hiện tượng sau:\n(1) Cho hạt đang nảy mầm vào bình thuỷ tinh đậy kín nắp, để yên trong khoảng 2 giờ.\n(2) Mở hé nắp bình và nhanh chóng đưa que diêm đang cháy vào miệng bình.\n(3) Que diêm đang cháy sẽ bùng lên sáng chói mãnh liệt hơn khi đưa vào bình chứa hạt nảy mầm.\n(4) Ngọn lửa của que diêm bị tắt ngay lập tức khi đưa vào bình chứa hạt nảy mầm.\n(5) Hiện tượng tắt lửa là do quá trình hô hấp của hạt đã tiêu thụ gần hết khí oxygen trong bình.\n(6) Ở bình đối chứng (chứa hạt luộc chín), que diêm vẫn tiếp tục cháy thêm một lúc vì hạt chết không hô hấp lấy mất oxygen.\nNhững thao tác và giải thích nào đúng với nguyên lí của thí nghiệm? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các phát biểu ở trên, ví dụ 123456)",
        "answer": "12456",
        "explanation": "Thao tác và giải thích đúng nguyên lí: (1) Chuẩn bị bình hạt nảy mầm đậy kín; (2) Mở hé nắp đưa que diêm vào; (4) Que diêm tắt ngay; (5) Do O~2~ bị tiêu hao; (6) Bình đối chứng diêm vẫn cháy do còn O~2~. Dãy số: 12456."
    },
    {
        "id": 34,
        "content": "Khi muốn chứng minh quá trình hô hấp ở hạt đang nảy mầm có tỏa nhiệt, nhóm của Bình sử dụng bình giữ nhiệt và nhiệt kế. Cho các phát biểu sau:\n(1) Cho hạt đang nảy mầm vào một bình giữ nhiệt (bọc xốp xung quanh) và cắm nhiệt kế vào giữa đám hạt.\n(2) Đặt bình thí nghiệm vào trong ngăn đá tủ lạnh để hạt hô hấp mạnh hơn.\n(3) Sau khoảng 2 - 3 giờ, số chỉ trên nhiệt kế ở bình chứa hạt nảy mầm tăng lên rõ rệt so với ban đầu.\n(4) Lớp xốp của bình giữ nhiệt có tác dụng ngăn lượng nhiệt do hạt sinh ra bị thất thoát ra môi trường.\n(5) Cắm nhiệt kế sâu vào giữa đám hạt nảy mầm sẽ làm hạt chết và ngừng hô hấp ngay lập tức.\n(6) Sự tăng nhiệt độ chứng tỏ hô hấp tế bào ở hạt nảy mầm đã giải phóng năng lượng dưới dạng nhiệt năng.\nNhững bước thiết kế và nhận định nào là đúng và hợp lý? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các phát biểu ở trên, ví dụ 123456)",
        "answer": "1346",
        "explanation": "Các bước và nhận định đúng: (1) Cắm nhiệt kế vào giữa đám hạt trong bình giữ nhiệt; (3) Nhiệt độ tăng sau 2-3h; (4) Bình xốp giữ nhiệt tránh thất thoát; (6) Nhiệt độ tăng chứng minh hô hấp giải phóng năng lượng dạng nhiệt. Dãy số: 1346."
    },
    {
        "id": 35,
        "content": "Khi tổng kết bài thực hành, giáo viên yêu cầu học sinh giải thích nguyên tắc chung của các thí nghiệm trên. Cho các khẳng định sau:\n(1) Hạt đang nảy mầm được ưu tiên chọn làm vật mẫu vì lúc này quá trình hô hấp tế bào diễn ra mạnh mẽ nhất.\n(2) Việc sử dụng hạt luộc chín ở các bình đối chứng nhằm mục đích tiêu diệt phôi, làm ngừng hô hấp để có cơ sở so sánh.\n(3) Quá trình hô hấp ở thực vật có cơ chế trái ngược hoàn toàn động vật: chúng hít vào khí carbon dioxide và thở ra oxygen.\n(4) Phải đậy thật kín nắp các bình thí nghiệm để ngăn sự trao đổi khí và nhiệt với môi trường bên ngoài, đảm bảo kết quả chính xác.\n(5) Hô hấp tế bào là quá trình phân giải chất hữu cơ, giải phóng năng lượng để cung cấp cho sự phát triển của mầm cây.\n(6) Trong bóng tối, hạt nảy mầm sẽ tự động chuyển sang quang hợp thay vì hô hấp tế bào.\nNhững khẳng định nào giải thích đúng bản chất khoa học của bài thực hành? (ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải ứng với thứ tự của các khẳng định ở trên, ví dụ 123456)",
        "answer": "1245",
        "explanation": "Các khẳng định khoa học chính xác: (1) Hạt nảy mầm hô hấp rất mạnh; (2) Hạt luộc chín dùng làm đối chứng; (4) Đậy kín ngăn trao đổi khí/nhiệt với bên ngoài; (5) Hô hấp phân giải chất hữu cơ tạo năng lượng. (Ý (3) và (6) hoàn toàn sai). Dãy số: 1245."
    },
    {
        "id": 36,
        "content": "Để chuẩn bị mẫu vật hạt nảy mầm phục vụ thí nghiệm chứng minh hô hấp tế bào ở thực vật, học sinh tiến hành các thao tác sau:\n(1) Ngâm hạt ngập trong nước ấm khoảng 40°C trong 2 giờ.\n(2) Đặt đĩa Petri chứa hạt trong điều kiện nhiệt độ phòng hoặc tủ ấm (30°C - 35°C) để hạt nảy mầm.\n(3) Lựa chọn các hạt chắc, khỏe, không bị vỡ hay sứt mọt.\n(4) Lót bông/giấy thấm đẫm nước vào đĩa Petri, rải đều hạt và đậy lớp bông/giấy thấm ẩm lên phía trên.\nHãy sắp xếp các thao tác trên theo đúng tiến trình chuẩn bị hạt nảy mầm trong SGK.",
        "answer": "3142",
        "explanation": "Tiến trình chuẩn bị mẫu vật trong SGK: (3) Chọn hạt chắc mẩy ──> (1) Ngâm nước ấm 40°C trong 2 giờ ──> (4) Lót bông ẩm và rải hạt vào đĩa Petri ──> (2) Đặt trong tủ ấm/nhiệt độ phòng để nảy mầm. Chuỗi số: 3142."
    },
    {
        "id": 37,
        "content": "Trong bài thực hành chứng minh hô hấp tế bào ở thực vật thông qua sự nảy mầm của hạt, học sinh cần chuẩn bị danh mục các dụng cụ, mẫu vật và hóa chất sau:\n(1) Chuông thuỷ tinh kín.\n(2) Dung dịch Iodine.\n(3) Đĩa chứa hạt đậu nảy mầm.\n(4) Cốc đựng dung dịch nước vôi trong.\n(5) Băng giấy đen.\n(6) Cốc thuỷ tinh và nhiệt kế.\nNhững dụng cụ, hóa chất và mẫu vật nào ở trên được sử dụng trực tiếp trong bài thực hành hô hấp ở thực vật? (Ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải, ví dụ 123456)",
        "answer": "1346",
        "explanation": "Dụng cụ, hóa chất dùng cho thực hành hô hấp: (1) Chuông thủy tinh kín; (3) Đĩa hạt đậu nảy mầm; (4) Cốc nước vôi trong; (6) Cốc thủy tinh và nhiệt kế. (Dung dịch iodine (2) và băng giấy đen (5) thuộc bài thực hành quang hợp). Dãy số: 1346."
    },
    {
        "id": 38,
        "content": "Để bài thực hành chứng minh hô hấp ở hạt nảy mầm đạt kết quả rõ ràng, học sinh thảo luận các lưu ý kỹ thuật sau:\n(1) Ưu tiên chọn loại hạt có vỏ mềm (như đậu xanh, đậu đỏ, hạt cải...) để hạt dễ hút nước.\n(2) Luộc chín hạt đậu trước khi ngâm để kích thích hạt nở nhanh hơn.\n(3) Đậy kín chuông thủy tinh để giữ lượng khí thải ra tập trung trong không gian chuông.\n(4) Cung cấp đủ độ ẩm và duy trì nhiệt độ ấm thích hợp (30 °C - 35 °C) cho hạt.\n(5) Dùng nước vôi trong mới pha để dễ dàng nhận biết hiện tượng đục váng kết tủa.\nNhững lưu ý kỹ thuật nào ở trên là đúng và cần thiết để thí nghiệm thành công? (Ghi đáp án dưới dạng một dãy chữ số tăng dần từ trái sang phải, ví dụ 123456)",
        "answer": "1345",
        "explanation": "Các lưu ý kĩ thuật đúng: (1), (3), (4), (5). Lưu ý (2) sai nghiêm trọng vì luộc chín hạt sẽ làm chết phôi mầm, hạt không thể hô hấp hay nảy mầm được. Dãy số: 1345."
    }
]
