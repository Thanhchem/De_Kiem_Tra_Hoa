# -*- coding: utf-8 -*-
"""
Dữ liệu Đề Cương Ôn Tập Kiểm Tra Giữa Kì I - Môn KHTN 7
Năm học 2026 - 2027
Đầy đủ 3 phân môn: Vật lý, Hóa học, Sinh học
Gồm cả nội dung câu hỏi, dữ liệu bảng, và lời giải chi tiết sư phạm.
"""

# I. VẬT LÝ - LÝ THUYẾT
VAT_LY_LY_THUYET = [
    {
        "id": "VL_LT_1",
        "question": "Câu 1. Nêu ý nghĩa của tốc độ. Tốc độ chuyển động là gì? Viết công thức tính tốc độ. Ghi chú tên và đơn vị của từng đại lượng có trong công thức.",
        "solution": """- *Ý nghĩa của tốc độ:* Tốc độ cho biết mức độ nhanh hay chậm của chuyển động. Chuyển động nào có tốc độ lớn hơn thì nhanh hơn, chuyển động nào có tốc độ nhỏ hơn thì chậm hơn.
- *Định nghĩa tốc độ chuyển động:* Tốc độ là đại lượng đặc trưng cho sự chuyển động nhanh hay chậm của một vật, được xác định bằng quãng đường vật đi được trong một đơn vị thời gian.
- *Công thức tính tốc độ:*
  *v = s / t*
- *Ghi chú tên và đơn vị của các đại lượng:*
  + *s*: Quãng đường đi được của vật (đơn vị: mét, kí hiệu ~m~; hoặc kilômét, kí hiệu ~km~).
  + *t*: Thời gian đi hết quãng đường đó (đơn vị: giây, kí hiệu ~s~; hoặc giờ, kí hiệu ~h~).
  + *v*: Tốc độ chuyển động của vật (đơn vị hợp pháp trong hệ đo lường SI là mét trên giây, kí hiệu ~m/s~; hoặc kilômét trên giờ, kí hiệu ~km/h~).
  + Mối quan hệ chuyển đổi đơn vị: 1 m/s = 3,6 km/h; 1 km/h ≈ 0,28 m/s.""",
        "blank_lines_student": 4,
        "blank_lines_screen": 8
    },
    {
        "id": "VL_LT_2",
        "question": "Câu 2. Theo em để xác định tốc độ chuyển động người ta phải đo những đại lượng nào và dùng những dụng cụ nào để đo? Tại sao?",
        "solution": """- *Các đại lượng cần đo:* Để xác định tốc độ chuyển động của một vật, người ta phải đo hai đại lượng:
  1. Quãng đường vật đi được (*s*).
  2. Thời gian vật đi hết quãng đường đó (*t*).
- *Các dụng cụ dùng để đo:*
  + Đo quãng đường (*s*): Dùng thước mét, thước dây, thước cuộn, thước kẹp hoặc cảm biến đo khoảng cách.
  + Đo thời gian (*t*): Dùng đồng hồ bấm giây (cơ học hoặc điện tử), hoặc đồng hồ đo thời gian hiện số kết hợp cổng quang điện.
- *Giải thích tại sao:*
  Vì theo công thức *v = s / t*, tốc độ phụ thuộc trực tiếp vào quãng đường và thời gian chuyển động. Do đó, muốn tính được tốc độ *v*, bắt buộc phải đo chính xác quãng đường *s* mà vật đã đi và khoảng thời gian *t* tương ứng để đi hết quãng đường đó.""",
        "blank_lines_student": 4,
        "blank_lines_screen": 8
    },
    {
        "id": "VL_LT_3",
        "question": "Câu 3. Nêu công dụng, cấu tạo và hoạt động của thiết bị bắn tốc độ.",
        "solution": """- *Công dụng:* Dùng để đo tốc độ chuyển động của các phương tiện giao thông (như ô tô, xe máy, xe tải...) đang lưu thông trên đường nhằm kiểm tra, giám sát việc chấp hành quy định về giới hạn tốc độ, góp phần bảo đảm an toàn giao thông và giảm thiểu tai nạn.
- *Cấu tạo chính:* Thiết bị bắn tốc độ gồm hai bộ phận cơ bản:
  1. Camera ghi hình: Ghi lại hình ảnh của phương tiện chuyển động và biển số xe.
  2. Bộ phận đo thời gian / phát - thu tín hiệu: Có thể là cảm biến phát - nhận xung laser, máy phát sóng vô tuyến radar, hoặc hệ thống tính toán thời gian xe chạy qua hai vạch mốc cách nhau một khoảng xác định.
  3. Bộ vi xử lý (máy tính mini tích hợp): Nhận dữ liệu, tự động tính tốc độ và hiển thị kết quả lên màn hình/ảnh chụp.
- *Hoạt động:*
  + *Cách 1 (Sử dụng hai vạch mốc và camera):* Camera chụp ảnh xe khi bắt đầu đi qua vạch mốc 1 và khi đi qua vạch mốc 2 (hai vạch cách nhau một khoảng *s* cố định). Bộ bấm giờ tự động đo khoảng thời gian *t* xe chạy giữa hai mốc. Máy tính tính tốc độ *v = s / t*. Nếu vượt quá tốc độ cho phép, hệ thống sẽ in tốc độ và lưu ảnh làm bằng chứng xử phạt.
  + *Cách 2 (Sử dụng súng bắn tốc độ laser hoặc radar):* Thiết bị phát chùm sóng laser hoặc sóng radar tới phương tiện đang chạy và thu sóng phản xạ trở lại. Dựa vào thời gian phản xạ của hai xung hoặc sự thay đổi tần số sóng (hiệu ứng Doppler), thiết bị tính toán ngay tốc độ tức thời của phương tiện cực kỳ chính xác.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 9
    },
    {
        "id": "VL_LT_4",
        "question": "Câu 4. Đồ thị quãng đường - thời gian cho biết gì? Đồ thị của chuyển động có tốc độ không đổi có dạng như thế nào?",
        "solution": """- *Đồ thị quãng đường - thời gian cho biết:*
  + Mối liên hệ định lượng giữa quãng đường đi được (*s*) và thời gian chuyển động (*t*) của một vật.
  + Vị trí, quãng đường vật đã đi được ở một thời điểm bất kì hoặc thời gian cần thiết để đi hết một quãng đường xác định mà không cần làm phép tính phức tạp.
  + Tốc độ của vật (độ dốc/nghiêng của đường đồ thị biểu thị tốc độ: đường càng dốc thì tốc độ càng lớn).
  + Tình trạng chuyển động của vật: chuyển động nhanh dần, chậm dần, chuyển động đều hay đứng yên (nếu đường đồ thị nằm ngang song song trục thời gian thì vật đứng yên).
- *Dạng của đồ thị chuyển động có tốc độ không đổi (chuyển động đều):*
  + Là một *đoạn thẳng nằm nghiêng (đường thẳng xiên góc)* đi lên so với trục nằm ngang (trục thời gian *t*).
  + Nếu chọn gốc thời gian lúc vật bắt đầu xuất phát từ vị trí mốc thì đường thẳng này xuất phát từ gốc tọa độ *O(0; 0)*.""",
        "blank_lines_student": 4,
        "blank_lines_screen": 8
    }
]

# I. VẬT LÝ - VẬN DỤNG
VAT_LY_VAN_DUNG = [
    {
        "id": "VL_VD_1",
        "question": "Câu 1. Một người đi bộ từ nhà ra công viên trên đoạn đường 3,6 km hết thời gian 40 min. Tính tốc độ của người đó ra đơn vị km/h và m/s.",
        "solution": """*Tóm tắt:*
- Quãng đường: *s = 3,6 km = 3600 m*.
- Thời gian: *t = 40 min = 40 / 60 h = 2/3 h ≈ 0,667 h* (hoặc *t = 40 × 60 = 2400 s*).
- Tính *v* theo *km/h* và *m/s*.

*Lời giải:*
- Tốc độ của người đó theo đơn vị *km/h* là:
  *v = s / t = 3,6 / (2/3) = 3,6 × 1,5 = 5,4 km/h*.
- Tốc độ của người đó theo đơn vị *m/s* là:
  *v = s / t = 3600 / 2400 = 1,5 m/s*
  *(Hoặc đổi từ km/h sang m/s: v = 5,4 / 3,6 = 1,5 m/s).*
- *Đáp số:* *5,4 km/h* và *1,5 m/s*.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 9
    },
    {
        "id": "VL_VD_2",
        "question": "Câu 2. Bạn Bình đi từ nhà đến trường hết 15 min với tốc độ 5 m/s. Tính quãng đường từ nhà đến trường của bạn Bình ra đơn vị m và km.",
        "solution": """*Tóm tắt:*
- Thời gian: *t = 15 min = 15 × 60 = 900 s*.
- Tốc độ: *v = 5 m/s* (hoặc *v = 5 × 3,6 = 18 km/h*).
- Tính *s* theo *m* và *km*.

*Lời giải:*
- Quãng đường từ nhà đến trường của bạn Bình theo đơn vị mét là:
  *s = v × t = 5 × 900 = 4500 m*.
- Đổi quãng đường sang đơn vị kilômét:
  *s = 4500 / 1000 = 4,5 km*.
- *Đáp số:* *4500 m* và *4,5 km*.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 8
    },
    {
        "id": "VL_VD_3",
        "question": "Câu 3. Bạn Tú đi bộ từ nhà đến thư viện lấy sách với tốc độ không đổi là 0,5 m/s. Biết quãng đường từ nhà đến thư viện là 0,6 km. Tính thời gian bạn Tú đi được ra đơn vị h và s.",
        "solution": """*Tóm tắt:*
- Tốc độ: *v = 0,5 m/s* (hoặc *v = 0,5 × 3,6 = 1,8 km/h*).
- Quãng đường: *s = 0,6 km = 600 m*.
- Tính *t* theo *s* và *h*.

*Lời giải:*
- Thời gian bạn Tú đi bộ theo đơn vị giây là:
  *t = s / v = 600 / 0,5 = 1200 s*.
- Đổi thời gian sang đơn vị giờ:
  *t = 1200 / 3600 = 1/3 h ≈ 0,333 h* (hoặc *t = 20 phút*).
  *(Hoặc tính: t = s / v = 0,6 / 1,8 = 1/3 h).*
- *Đáp số:* *1200 s* và *1/3 h* (khoảng *0,333 h*).""",
        "blank_lines_student": 5,
        "blank_lines_screen": 8
    },
    {
        "id": "VL_VD_4",
        "question": "Câu 4. Một chiếc ô tô và một chiếc xe máy được phát hiện bởi thiết bị bắn tốc độ, chạy từ vạch mốc 1 và vạch mốc 2 cách nhau 5 m thời gian ô tô chạy giữa hai vạch mốc là 0,2 s và thời gian xe máy chạy giữa hai vạch mốc là 0,25 s.\na. Hỏi tốc độ của ô tô và của xe máy bằng bao nhiêu?\nb. Nếu tốc độ giới hạn của cung đường là 80 km/h thì ô tô và xe máy có vượt quá tốc độ giới hạn không? Tại sao?",
        "solution": """*Tóm tắt:*
- Quãng đường giữa hai vạch mốc: *s = 5 m*.
- Thời gian ô tô chạy: *t_1 = 0,2 s*.
- Thời gian xe máy chạy: *t_2 = 0,25 s*.
- Giới hạn tốc độ: *v_gh = 80 km/h*.

*Lời giải:*
*a. Tính tốc độ của ô tô và xe máy:*
- Tốc độ của chiếc ô tô là:
  *v_ô tô = s / t_1 = 5 / 0,2 = 25 m/s*.
  Đổi sang đơn vị km/h: *v_ô tô = 25 × 3,6 = 90 km/h*.
- Tốc độ của chiếc xe máy là:
  *v_xe máy = s / t_2 = 5 / 0,25 = 20 m/s*.
  Đổi sang đơn vị km/h: *v_xe máy = 20 × 3,6 = 72 km/h*.

*b. So sánh với tốc độ giới hạn:*
- Đối với ô tô: *v_ô tô = 90 km/h > 80 km/h*, do đó *ô tô đã vượt quá tốc độ giới hạn* (vượt 10 km/h).
- Đối với xe máy: *v_xe máy = 72 km/h < 80 km/h*, do đó *xe máy không vượt quá tốc độ giới hạn* (chạy đúng quy định an toàn).""",
        "blank_lines_student": 7,
        "blank_lines_screen": 11
    },
    {
        "id": "VL_VD_5",
        "question": "Câu 5. Hình dưới đây biểu diễn đồ thị quãng đường - thời gian của một vật chuyển động trong khoảng thời gian 8 s.\na. Xác định quãng đường đi được của vật sau 8 s kể từ lúc xuất phát.\nb. Sau bao lâu kể từ lúc xuất phát vật đi hết quãng đường 15 m?\nc. Sau 2 s kể từ lúc xuất phát vật đi hết quãng đường là bao nhiêu?\nd. Từ đồ thị, hãy xác định tốc độ của vật.",
        "image_path": "San_Pham/images_khtn7/do_thi_cau5.png",
        "image_caption": "Đồ thị quãng đường – thời gian của một vật chuyển động trong 8 s",
        "solution": """*Lời giải chi tiết:*
*a. Quãng đường đi được của vật sau 8 s:*
- Từ giá trị *t = 8 s* trên trục hoành, dóng thẳng đứng lên gặp đồ thị, rồi từ điểm đó dóng nằm ngang sang trục tung, ta đọc được giá trị *s = 20 m*.
- *Kết luận:* Sau 8 s kể từ lúc xuất phát, vật đi được *20 m*.

*b. Thời gian để vật đi hết quãng đường 15 m:*
- Từ giá trị *s = 15 m* trên trục tung, dóng nằm ngang sang phải gặp đường đồ thị, rồi từ điểm cắt dóng thẳng đứng xuống trục hoành, ta đọc được giá trị *t = 6 s*.
- *Kết luận:* Sau *6 s* kể từ lúc xuất phát, vật đi hết quãng đường 15 m.

*c. Quãng đường vật đi được sau 2 s:*
- Từ giá trị *t = 2 s* trên trục hoành, dóng thẳng đứng lên gặp đường đồ thị, rồi dóng nằm ngang sang trục tung, ta đọc được giá trị *s = 5 m*.
- *Kết luận:* Sau 2 s kể từ lúc xuất phát, vật đi được quãng đường là *5 m*.

*d. Xác định tốc độ của vật:*
- Vì đồ thị quãng đường - thời gian là một đoạn thẳng đi qua gốc tọa độ nên vật chuyển động thẳng đều (tốc độ không đổi).
- Tốc độ của vật là:
  *v = s / t = 20 / 8 = 2,5 m/s*
  *(Hoặc tính: v = 5 / 2 = 10 / 4 = 15 / 6 = 2,5 m/s = 9 km/h).*
- *Đáp số:* *v = 2,5 m/s* (hay *9 km/h*).""",
        "blank_lines_student": 6,
        "blank_lines_screen": 10
    },
    {
        "id": "VL_VD_6",
        "question": "Câu 6. Bảng dưới đây ghi lại quãng đường đi được theo thời gian của một người đi bộ.\na. Dựa vào số liệu trong bảng, hãy vẽ đồ thị quãng đường - thời gian của người đi bộ.\nb. Từ đồ thị, xác định tốc độ đi bộ của người đó.",
        "table_data": {
            "headers": ["Thời gian (s)", "0", "10", "20", "30", "40", "50"],
            "rows": [
                ["Quãng đường (m)", "0", "14", "28", "42", "56", "70"]
            ]
        },
        "solution_image": "San_Pham/images_khtn7/do_thi_cau6.png",
        "solution": """*Lời giải chi tiết:*
*a. Hướng dẫn các bước vẽ đồ thị quãng đường - thời gian:*
- *Bước 1: Vẽ hệ trục tọa độ:*
  + Trục hoành nằm ngang biểu diễn thời gian *t (s)*, gốc *O(0)*, chia vạch mỗi khoảng 10 s (0, 10, 20, 30, 40, 50).
  + Trục tung thẳng đứng biểu diễn quãng đường *s (m)*, gốc *O(0)*, chia vạch mỗi khoảng 14 m (0, 14, 28, 42, 56, 70).
- *Bước 2: Xác định các điểm tọa độ từ bảng số liệu:*
  + Gốc xuất phát: *O(0; 0)*.
  + Điểm *A(10; 14)*: tại *t = 10 s, s = 14 m*.
  + Điểm *B(20; 28)*: tại *t = 20 s, s = 28 m*.
  + Điểm *C(30; 42)*: tại *t = 30 s, s = 42 m*.
  + Điểm *D(40; 56)*: tại *t = 40 s, s = 56 m*.
  + Điểm *E(50; 70)*: tại *t = 50 s, s = 70 m*.
- *Bước 3: Nối các điểm:*
  + Nối gốc tọa độ *O* và các điểm *A, B, C, D, E*, ta được một đường thẳng xiên góc đi qua gốc tọa độ. Đó là đồ thị quãng đường - thời gian của người đi bộ.

*b. Xác định tốc độ đi bộ của người đó:*
- Từ đồ thị (hoặc bảng số liệu), ta thấy sau những khoảng thời gian bằng nhau (10 s), người đó đi được những quãng đường bằng nhau (14 m). Do đó, người đi bộ chuyển động đều.
- Tốc độ đi bộ của người đó là:
  *v = s / t = 14 / 10 = 28 / 20 = 70 / 50 = 1,4 m/s*.
- Đổi sang đơn vị km/h:
  *v = 1,4 × 3,6 = 5,04 km/h*.
- *Đáp số:* *1,4 m/s* (hay *5,04 km/h*).""",
        "blank_lines_student": 5,
        "blank_lines_screen": 12,
        "screen_box": True
    }
]

# II. HÓA HỌC - LÝ THUYẾT
HOA_HOC_LY_THUYET = [
    {
        "id": "HH_LT_1",
        "question": "Câu 1. Bài 1: Phần II: Nêu và phân tích một số kĩ năng tiến trình học tập môn khoa học tự nhiên: quan sát, phân loại, liên kết, đo, dự báo.",
        "solution": """Các kĩ năng tiến trình học tập môn Khoa học tự nhiên gồm 5 kĩ năng cốt lõi:
1. *Kĩ năng quan sát:* Là việc thu thập các thông tin, hình ảnh, đặc điểm về sự vật, hiện tượng trong tự nhiên bằng các giác quan (thị giác, thính giác, khứu giác, xúc giác...) hoặc sử dụng các dụng cụ, phương tiện hỗ trợ quan sát (kính lúp, kính hiển vi, ống nhòm...).
2. *Kĩ năng phân loại:* Là việc sắp xếp, phân nhóm các đối tượng, sự vật, hiện tượng hoặc dữ liệu có những dấu hiệu, đặc điểm hoặc tiêu chí giống nhau vào từng nhóm, lớp xác định theo một trật tự logic nhất định.
3. *Kĩ năng liên kết:* Là việc kết nối, phối hợp các tri thức khoa học đã có với các kết quả, dữ liệu mới thu thập được nhằm giải thích nguyên nhân, phát hiện mối liên hệ nhân - quả và rút ra bản chất quy luật của sự vật, hiện tượng.
4. *Kĩ năng đo:* Là việc sử dụng các dụng cụ và thiết bị đo lường chuẩn xác (như thước kẻ, cân điện tử, nhiệt kế, ống đong, đồng hồ bấm giây...) để xác định kích thước, khối lượng, thể tích, nhiệt độ, thời gian... của đối tượng nghiên cứu dưới dạng định lượng số liệu cụ thể.
5. *Kĩ năng dự báo:* Là việc đưa ra các nhận định, phán đoán có cơ sở khoa học về những sự kiện, kết quả có thể xảy ra trong tương lai hoặc chiều hướng diễn biến của hiện tượng dựa trên các quan sát thực nghiệm và các quy luật khoa học đã biết.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 10
    },
    {
        "id": "HH_LT_2",
        "question": "Câu 2. Chương I. Nguyên tử. Sơ lược về bảng tuần hoàn các nguyên tố hóa học:\n- Trình bày mô hình nguyên tử của Rutherford - Bohr: mô hình sắp xếp electron trong các lớp của vỏ nguyên tử, cấu tạo nguyên tử (hạt nhân và lớp vỏ nguyên tử).\n- Nêu khối lượng của một nguyên tử theo đơn vị quốc tế amu, khối lượng amu các hạt proton, neutron, electron, từ đó nêu cách tính khối lượng nguyên tử.",
        "solution": """*1. Mô hình nguyên tử của Rutherford - Bohr:*
- *Cấu tạo nguyên tử:* Nguyên tử là hạt vô cùng nhỏ bé và trung hòa về điện, gồm hai phần chính:
  + *Hạt nhân:* Nằm ở tâm nguyên tử, mang điện tích dương (+), gồm hai loại hạt:
    • *Proton (kí hiệu p):* Mang điện tích dương (+1).
    • *Neutron (kí hiệu n):* Không mang điện tích (0).
  + *Lớp vỏ nguyên tử:* Gồm các hạt *electron (kí hiệu e)* mang điện tích âm (-1) chuyển động rất nhanh xung quanh hạt nhân.
  + Do nguyên tử trung hòa về điện nên: *Số proton (p) = Số electron (e)*.
- *Mô hình sắp xếp electron trong các lớp của vỏ nguyên tử (mô hình Bohr):*
  + Các electron chuyển động xung quanh hạt nhân theo từng quỹ đạo xác định tạo thành các lớp electron xếp từ trong ra ngoài.
  + Mỗi lớp electron có số electron tối đa xác định:
    • Lớp 1 (lớp trong cùng, sát hạt nhân nhất): Chứa tối đa *2 electron*.
    • Lớp 2: Chứa tối đa *8 electron*.
    • Lớp 3: Chứa tối đa *8 electron* (đối với 20 nguyên tố đầu bảng tuần hoàn).

*2. Khối lượng của nguyên tử theo đơn vị quốc tế amu:*
- Đơn vị đo khối lượng nguyên tử là *amu* (atomic mass unit):
  1 amu = 1/12 khối lượng của một nguyên tử Carbon-12 ≈ 1,6605 × 10^-24^ g = 1,6605 × 10^-27^ kg.
- Khối lượng xấp xỉ của các hạt cấu tạo nên nguyên tử:
  + Khối lượng của 1 proton: *m_p ≈ 1 amu*.
  + Khối lượng của 1 neutron: *m_n ≈ 1 amu*.
  + Khối lượng của 1 electron: *m_e ≈ 0,00055 amu* (rất nhỏ, chỉ bằng khoảng 1/1836 khối lượng của proton).
- *Cách tính khối lượng nguyên tử:*
  Vì khối lượng của các electron vô cùng nhỏ so với khối lượng hạt nhân nên khối lượng của nguyên tử coi như tập trung toàn bộ ở hạt nhân:
  *Khối lượng nguyên tử ≈ Khối lượng hạt nhân = m_p + m_n ≈ (Số p + Số n) amu*.""",
        "blank_lines_student": 6,
        "blank_lines_screen": 12
    }
]

# II. HÓA HỌC - VẬN DỤNG
HOA_HOC_VAN_DUNG = [
    {
        "id": "HH_VD_1",
        "question": "Câu 1. Quan sát hình 2.1 sgk và cho biết các thành phần cấu tạo của nguyên tử.",
        "solution": """Dựa vào hình 2.1 SGK (Sơ đồ cấu tạo nguyên tử), nguyên tử được cấu tạo từ hai thành phần chính:
1. *Hạt nhân nguyên tử:* Nằm ở trung tâm nguyên tử, mang điện tích dương, gồm hai loại hạt:
   - Hạt *proton (p)* mang điện tích dương (+).
   - Hạt *neutron (n)* không mang điện tích.
2. *Lớp vỏ nguyên tử:* Bao quanh hạt nhân, gồm các hạt *electron (e)* mang điện tích âm (-) chuyển động rất nhanh trong không gian xung quanh hạt nhân.""",
        "blank_lines_student": 4,
        "blank_lines_screen": 7
    },
    {
        "id": "HH_VD_2",
        "question": "Câu 2. Quan sát hình 2.2 sgk áp dụng mô hình nguyên tử của Bo, mô tả cấu tạo của nguyên tử hydrogen, và nguyên tử carbon.",
        "solution": """Áp dụng mô hình nguyên tử của Bohr (Bo):
1. *Cấu tạo nguyên tử Hydrogen (H):*
   - Hạt nhân: Có *1 proton* mang điện tích +1.
   - Vỏ nguyên tử: Có *1 electron* mang điện tích -1, chuyển động trên *1 lớp* electron duy nhất bao quanh hạt nhân.
   - Khối lượng nguyên tử: Xấp xỉ *1 amu*.
2. *Cấu tạo nguyên tử Carbon (C):*
   - Hạt nhân: Có *6 proton* mang điện tích +6 và *6 neutron* không mang điện.
   - Vỏ nguyên tử: Có *6 electron* mang điện tích -6, được phân bố vào *2 lớp* electron từ trong ra ngoài:
     + Lớp 1 (lớp trong cùng): Có *2 electron* (đã đạt số e tối đa của lớp 1).
     + Lớp 2 (lớp ngoài cùng): Có *4 electron*.
   - Khối lượng nguyên tử: Xấp xỉ *6 + 6 = 12 amu*.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 8
    },
    {
        "id": "HH_VD_3",
        "question": "Câu 3. Khi tìm hiểu về mô hình nguyên tử carbon theo Rutherford – Bohr, người ta xác định được hạt nhân nguyên tử chứa 6 proton, 6 neutron và vỏ nguyên tử chứa 6 electron.\na. Ở lớp vỏ nguyên tử carbon, 6 electron được phân bố thành 2 lớp. Vẽ mô hình cấu tạo nguyên tử Carbon.\nb. Biết rằng khối lượng nguyên tử tập trung chủ yếu ở hạt nhân và coi khối lượng mỗi proton và neutron xấp xỉ 1 amu. Tính khối lượng của nguyên tử carbon bao nhiêu amu?",
        "solution_image": "San_Pham/images_khtn7/mo_hinh_carbon.png",
        "solution": """*Lời giải chi tiết:*
*a. Phân bố electron và mô hình cấu tạo nguyên tử Carbon:*
- Vỏ nguyên tử Carbon có tổng số 6 electron, phân bố trên 2 lớp electron từ trong ra ngoài:
  + Lớp 1 (trong cùng): Chứa tối đa *2 electron*.
  + Lớp 2 (lớp ngoài cùng): Chứa *4 electron* còn lại (6 - 2 = 4).
- *Mô hình cấu tạo:* Vẽ hạt nhân ở tâm mang điện tích +6 (gồm 6p và 6n). Xung quanh hạt nhân vẽ 2 đường tròn đồng tâm: đường tròn thứ nhất chứa 2 electron đối xứng nhau, đường tròn thứ hai chứa 4 electron phân bố đều đối xứng trên 4 hướng.

*b. Tính khối lượng nguyên tử Carbon theo amu:*
- Khối lượng nguyên tử tập trung chủ yếu ở hạt nhân, bỏ qua khối lượng của 6 electron ở lớp vỏ:
  *Khối lượng nguyên tử C ≈ m_p + m_n = (6 × 1 amu) + (6 × 1 amu) = 12 amu*.
- *Đáp số:* *12 amu*.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 10,
        "screen_box": True
    },
    {
        "id": "HH_VD_4",
        "question": "Câu 4. Khi tìm hiểu cấu tạo nguyên tử nhôm (Aluminium) theo mô hình Rutherford – Bohr, học sinh xác định được trong hạt nhân nguyên tử này có 13 proton.\na. Em hãy vẽ sơ đồ mô hình nguyên tử nhôm theo Rutherford – Bohr.\nb. Biết nhôm có khối lượng nguyên tử 27 amu, khối lượng mỗi proton và neutron xấp xỉ 1 amu. Tính số hạt neutron.",
        "solution_image": "San_Pham/images_khtn7/mo_hinh_aluminium.png",
        "solution": """*Lời giải chi tiết:*
*a. Phân bố electron và sơ đồ mô hình nguyên tử Nhôm (Aluminium):*
- Vì nguyên tử trung hòa về điện nên: Số electron = Số proton = *13 electron*.
- 13 electron này được sắp xếp vào 3 lớp electron từ trong ra ngoài:
  + Lớp 1 (trong cùng): Chứa tối đa *2 electron*.
  + Lớp 2: Chứa tối đa *8 electron*.
  + Lớp 3 (lớp ngoài cùng): Chứa số electron còn lại là *13 - (2 + 8) = 3 electron*.
- *Sơ đồ mô hình:* Hạt nhân ở tâm mang điện tích +13 (gồm 13p, 14n), bao quanh bởi 3 vòng tròn đồng tâm biểu thị 3 lớp electron (lớp 1 có 2e, lớp 2 có 8e, lớp 3 có 3e).

*b. Tính số hạt neutron trong nguyên tử Nhôm:*
- Khối lượng nguyên tử tập trung ở hạt nhân:
  *Khối lượng nguyên tử Al ≈ m_p + m_n ≈ (Số p × 1 amu) + (Số n × 1 amu) = 27 amu*.
- Suy ra:
  *Số p + Số n = 27*
  *Số n = 27 - Số p = 27 - 13 = 14 hạt*.
- *Đáp số:* *14 hạt neutron*.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 10,
        "screen_box": True
    }
]

# III. SINH HỌC - TRẮC NGHIỆM / CÂU HỎI ÔN TẬP
SINH_HOC_TRAC_NGHIEM = [
    {
        "id": "SH_TN_1",
        "question": "Câu 1. Nêu khái niệm và vai trò của trao đổi chất và chuyển hóa năng lượng ở sinh vật.",
        "solution": """- *Khái niệm:*
  + *Trao đổi chất:* Là quá trình cơ thể sinh vật lấy các chất từ môi trường ngoài vào (như nước, chất dinh dưỡng, khí oxygen, ánh sáng...), biến đổi chúng thành các chất cần thiết cho cơ thể, đồng thời thải các chất cặn bã, dư thừa ra ngoài môi trường.
  + *Chuyển hóa năng lượng:* Là sự biến đổi năng lượng từ dạng này sang dạng khác trong tế bào và cơ thể (ví dụ: chuyển hóa quang năng thành hóa năng trong quang hợp, chuyển hóa hóa năng thành nhiệt năng, cơ năng, điện năng trong hoạt động sống).
  + Trao đổi chất và chuyển hóa năng lượng luôn diễn ra đồng thời và gắn bó khăng khít với nhau.
- *Vai trò:*
  + Cung cấp nguyên liệu để xây dựng, tái tạo tế bào, cơ quan và các chất sống cần thiết cho sự duy trì, sinh trưởng, phát triển và sinh sản của cơ thể sinh vật.
  + Cung cấp năng lượng (chủ yếu tích lũy dưới dạng phân tử ATP) để thực hiện mọi hoạt động sống của tế bào và cơ thể (vận động, cảm ứng, duy trì thân nhiệt, vận chuyển các chất...).
  + Đào thải các chất độc hại, chất bã ra khỏi cơ thể giúp duy trì cân bằng nội môi. Sinh vật sẽ ngừng sống nếu ngừng trao đổi chất và chuyển hóa năng lượng.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 9
    },
    {
        "id": "SH_TN_2",
        "question": "Câu 2. Các yếu tố môi trường ảnh hưởng đến quang hợp như thế nào?",
        "solution": """Các yếu tố môi trường chủ yếu ảnh hưởng đến quá trình quang hợp ở thực vật:
1. *Ánh sáng:* Là nguồn cung cấp năng lượng cho quang hợp. Khi cường độ ánh sáng tăng thì cường độ quang hợp tăng dần cho đến khi đạt điểm bão hòa ánh sáng. Thành phần quang phổ (ánh sáng đỏ và xanh tím giúp quang hợp hiệu quả nhất).
2. *Nồng độ khí Carbon dioxide (CO~2~):* CO~2~ là nguyên liệu chính để tổng hợp chất hữu cơ. Khi nồng độ CO~2~ trong không khí tăng từ mức bình thường (0,03%) thì hiệu suất quang hợp tăng; nhưng nếu nồng độ CO~2~ quá cao sẽ làm khí khổng đóng lại và gây độc cho cây.
3. *Nước:* Vừa là nguyên liệu trực tiếp tham gia phản ứng quang hợp, vừa là dung môi hòa tan vận chuyển các chất, vừa điều tiết sự đóng mở của khí khổng. Khi cây thiếu nước, khí khổng đóng, quang hợp giảm mạnh hoặc ngừng hẳn.
4. *Nhiệt độ:* Ảnh hưởng trực tiếp đến hoạt tính của các enzyme xúc tác quang hợp. Nhiệt độ thích hợp cho quang hợp của đa số thực vật là từ 25°C đến 35°C. Nhiệt độ quá thấp (< 5°C) hoặc quá cao (> 40°C) đều ức chế enzyme và làm giảm hoặc ngừng quang hợp.
5. *Chất khoáng:* Các nguyên tố khoáng (N, P, K, Mg, Fe...) tham gia cấu tạo nên phân tử diệp lục, hệ enzyme và các chất truyền điện tử trong quang hợp.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 9
    },
    {
        "id": "SH_TN_3",
        "question": "Câu 3. Các yếu tố môi trường ảnh hưởng đến hô hấp tế bào như thế nào?",
        "solution": """Quá trình hô hấp tế bào chịu ảnh hưởng sâu sắc bởi các yếu tố môi trường:
1. *Hàm lượng nước:* Nước là môi trường và tham gia trực tiếp vào các phản ứng hóa học của hô hấp tế bào. Độ ẩm và hàm lượng nước cao làm tăng cường độ hô hấp; hàm lượng nước thấp (ví dụ ở hạt khô nông sản) làm hô hấp giảm xuống mức tối thiểu.
2. *Nhiệt độ:* Ảnh hưởng trực tiếp đến hoạt tính của hệ enzyme hô hấp. Nhiệt độ tối ưu cho hô hấp tế bào thường dao động từ 30°C đến 35°C. Nhiệt độ quá thấp làm enzyme hoạt động chậm (hô hấp yếu); nhiệt độ quá cao (> 45°C - 50°C) làm biến tính enzyme, làm ngừng hô hấp dẫn đến tế bào chết.
3. *Hàm lượng khí Oxygen (O~2~):* O~2~ là chất tham gia trực tiếp oxy hóa chất hữu cơ trong hô hấp hiếu khí. Khi nồng độ O~2~ trong không khí giảm xuống dưới 5%, hô hấp hiếu khí giảm, tế bào chuyển sang hô hấp kị khí (lên men) làm tích lũy các chất gây độc (ethanol, acid lactic...).
4. *Hàm lượng khí Carbon dioxide (CO~2~):* CO~2~ là sản phẩm của quá trình hô hấp. Khi nồng độ CO~2~ trong môi trường tăng cao (> 3% - 5%), nó sẽ gây ức chế phản ứng hô hấp tế bào.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 9
    },
    {
        "id": "SH_TN_4",
        "question": "Câu 4. Nêu các đặc điểm cấu tạo của lá phù hợp với chức năng quang hợp.",
        "solution": """Lá cây có cấu tạo hình thái và giải phẫu thích nghi cao độ với chức năng quang hợp:
1. *Đặc điểm hình thái bên ngoài:*
   - *Phiến lá:* Bản mỏng, diện tích bề mặt rộng giúp hấp thụ được tối đa ánh sáng mặt trời.
   - *Cách sắp xếp lá trên thân/cành:* Các lá mọc so le hoặc xếp vòng tròn quanh thân giúp lá không che khuất ánh sáng lẫn nhau.
   - *Gân lá:* Chứa các bó mạch dẫn (mạch gỗ dẫn nước và muối khoáng từ rễ lên; mạch rây dẫn chất hữu cơ quang hợp đi khắp cơ thể), phân nhánh chằng chịt đến từng tế bào thịt lá.
2. *Đặc điểm cấu tạo giải phẫu bên trong:*
   - *Lớp biểu bì:* Có lớp cutin mỏng trong suốt để ánh sáng dễ dàng xuyên qua. Trên lớp biểu bì (chủ yếu ở mặt dưới của lá) có vô số *khí khổng* có khả năng đóng mở linh hoạt để khí CO~2~ khuếch tán vào lá và giải phóng khí O~2~, hơi nước ra ngoài.
   - *Lớp tế bào thịt lá:*
     + *Mô giậu (nằm sát biểu bì trên):* Gồm các tế bào hình trụ xếp sít nhau thẳng đứng, chứa mật độ *lục lạp (chứa diệp lục)* rất cao nhằm hấp thụ ánh sáng trực tiếp hiệu quả nhất.
     + *Mô khuyết (mô xốp nằm phía dưới):* Gồm các tế bào xếp lỏng lẻo tạo ra nhiều khoảng gian bào chứa khí giúp khí CO~2~ khuếch tán nhanh chóng đến các tế bào mô giậu.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 9
    },
    {
        "id": "SH_TN_5",
        "question": "Câu 5. Trình bày khái niệm và viết phương trình tổng quát của quang hợp và hô hấp tế bào.",
        "solution": """*1. Quang hợp:*
- *Khái niệm:* Là quá trình lục lạp ở thực vật hấp thụ và sử dụng năng lượng ánh sáng mặt trời để tổng hợp chất hữu cơ (glucose) từ các chất vô cơ đơn giản (nước và khí carbon dioxide), đồng thời giải phóng khí oxygen.
- *Phương trình tổng quát:*
  *Nước + Khí carbon dioxide ──(Ánh sáng / Diệp lục)──> Glucose + Khí oxygen*
  *(Phương trình hóa học: 6H~2~O + 6CO~2~ ──(Ánh sáng / Diệp lục)──> C~6~H~12~O~6~ + 6O~2~↑)*

*2. Hô hấp tế bào:*
- *Khái niệm:* Là quá trình tế bào phân giải các hợp chất hữu cơ (chủ yếu là glucose) thành các chất đơn giản (CO~2~ và H~2~O), đồng thời giải phóng năng lượng tích lũy trong chất hữu cơ để chuyển thành năng lượng trong các phân tử ATP cung cấp cho hoạt động sống và giải phóng nhiệt năng.
- *Phương trình tổng quát:*
  *Glucose + Khí oxygen ──(Enzyme)──> Khí carbon dioxide + Nước + Năng lượng (ATP + Nhiệt)*
  *(Phương trình hóa học: C~6~H~12~O~6~ + 6O~2~ ──(Enzyme)──> 6CO~2~ + 6H~2~O + Năng lượng (ATP + Nhiệt))*""",
        "blank_lines_student": 5,
        "blank_lines_screen": 9
    },
    {
        "id": "SH_TN_6",
        "question": "Câu 6. Trình bày mối quan hệ giữa tổng hợp và phân giải chất hữu cơ ở tế bào.",
        "solution": """Quá trình tổng hợp và phân giải chất hữu cơ ở tế bào là hai mặt của một thể thống nhất, có mối quan hệ hai chiều, gắn bó mật thiết và không thể tách rời:
- *Mặt đối lập:*
  + *Tổng hợp (đồng hóa):* Biến đổi chất vô cơ đơn giản thành chất hữu cơ phức tạp; tích lũy năng lượng dưới dạng hóa năng trong các liên kết hóa học.
  + *Phân giải (dị hóa):* Phân giải chất hữu cơ phức tạp thành các chất đơn giản; giải phóng năng lượng tích lũy trong các liên kết hóa học để tạo ATP và nhiệt năng.
- *Mặt thống nhất và phụ thuộc lẫn nhau:*
  + Sản phẩm của quá trình tổng hợp (chất hữu cơ như glucose, tinh bột, protein, lipid...) chính là nguyên liệu đầu vào cho quá trình phân giải. Nếu không có tổng hợp thì phân giải không thể diễn ra.
  + Ngược lại, năng lượng ATP và các sản phẩm trung gian do quá trình phân giải giải phóng ra lại là nguồn cung cấp trực tiếp cho các phản ứng sinh tổng hợp các chất mới trong tế bào.
- *Kết luận:* Hai quá trình này luôn diễn ra nhịp nhàng, cân bằng; nếu một trong hai quá trình ngừng trệ thì tế bào và sinh vật sẽ không thể duy trì sự sống.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 9
    },
    {
        "id": "SH_TN_7",
        "question": "Câu 7. Trình bày ý nghĩa của việc trồng và bảo vệ cây xanh.",
        "solution": """Cây xanh được ví như 'lá phổi xanh' của Trái Đất, việc trồng và bảo vệ cây xanh có ý nghĩa sống còn:
1. *Điều hòa thành phần khí quyển:* Quá trình quang hợp của cây xanh hấp thụ lượng lớn khí CO~2~ (khí gây hiệu ứng nhà kính và biến đổi khí hậu) và giải phóng khí O~2~ cần thiết cho hô hấp của con người và muôn loài sinh vật.
2. *Làm sạch và làm mát môi trường:* Tán lá cây cản bụi bẩn, giữ lại các hạt bụi mịn, hấp thụ một số khí độc hại, giảm bức xạ nhiệt mặt trời và thoát hơi nước làm hạ nhiệt độ không khí xung quanh, giảm ô nhiễm tiếng ồn ở đô thị.
3. *Bảo vệ đất và nguồn nước:* Hệ rễ cây bám sâu vào lòng đất giúp giữ đất, chống xói mòn, sạt lở đất đồi núi và bờ sông bờ biển; tán rừng cản bớt lực rơi của hạt mưa, làm tăng khả năng thấm nước vào mạch nước ngầm, giúp điều tiết nguồn nước, hạn chế lũ lụt vào mùa mưa và hạn hán vào mùa khô.
4. *Cung cấp nơi ở và nguồn sống:* Rừng là sinh cảnh, nơi cư trú của vô số loài động thực vật; cung cấp gỗ, thức ăn, dược liệu quý và các tài nguyên phục vụ đời sống con người.
5. *Nâng cao cảnh quan và sức khỏe con người:* Không gian xanh giúp tạo cảnh quan tươi đẹp cho đô thị, làng quê, trường học, mang lại cảm giác thư thái, giảm căng thẳng và tăng cường sức khỏe thể chất, tinh thần.""",
        "blank_lines_student": 5,
        "blank_lines_screen": 9
    }
]

# III. SINH HỌC - TỰ LUẬN
SINH_HOC_TU_LUAN = [
    {
        "id": "SH_TL_1",
        "question": "Câu hỏi Tự luận: Vận dụng hô hấp tế bào vào thực tiễn: Giải thích vì sao rau củ quả thường được bảo quản trong ngăn mát tủ lạnh. Vì sao một số loại thực phẩm bảo quản quá lâu thì bị giảm chất lượng?",
        "solution": """*Lời giải chi tiết:*
*1. Giải thích vì sao rau củ quả thường được bảo quản trong ngăn mát tủ lạnh:*
- Rau củ quả sau khi thu hái rời khỏi cây mẹ vẫn là các cơ quan sống, các tế bào vẫn liên tục diễn ra quá trình *hô hấp tế bào*.
- Trong quá trình hô hấp, tế bào tiêu hao chất hữu cơ dự trữ (đường, tinh bột, vitamin, khoáng chất...) và nước, đồng thời giải phóng khí CO~2~ và nhiệt. Nếu hô hấp diễn ra mạnh, rau củ quả sẽ rất nhanh bị tiêu hao dinh dưỡng, héo úa, sụt giảm khối lượng, mất vị ngọt tự nhiên và dễ bị hư hỏng.
- Nhiệt độ trong ngăn mát tủ lạnh thường duy trì ổn định từ *1°C đến 4°C*. Ở mức nhiệt độ thấp này:
  + Hoạt tính của các *enzyme xúc tác hô hấp tế bào bị ức chế mạnh*, làm cho cường độ hô hấp của rau củ giảm xuống mức tối thiểu, hạn chế tối đa việc phân giải các chất hữu cơ dự trữ.
  + Đồng thời, nhiệt độ thấp cũng kìm hãm sự phát triển và sinh sôi của các vi sinh vật, nấm mốc gây thối rữa và làm chậm quá trình thoát hơi nước.
- Nhờ đó, rau củ quả bảo quản trong ngăn mát giữ được độ tươi ngon, mọng nước, hàm lượng dinh dưỡng và kéo dài thời gian sử dụng được lâu hơn nhiều so với nhiệt độ thường.

*2. Giải thích vì sao một số loại thực phẩm bảo quản quá lâu thì bị giảm chất lượng:*
- Mặc dù ở điều kiện nhiệt độ thấp trong tủ lạnh, quá trình hô hấp tế bào bị ức chế nhưng *vẫn diễn ra chậm rãi* chứ không ngừng hẳn hoàn toàn.
- Khi bảo quản trong thời gian quá lâu:
  + Quá trình hô hấp chậm kéo dài liên tục làm tiêu hao dần lượng đường, tinh bột, protein, vitamin và các vi chất dinh dưỡng thiết yếu, làm cho thực phẩm bị xốp, nhạt nhẽo, mất hương vị thơm ngon ban đầu.
  + Sự mất nước vi mô diễn ra liên tục khiến tế bào bị co nguyên sinh, rau củ bị héo mềm, mất đi độ giòn dai tự nhiên.
  + Một số loại củ (như khoai tây, hành, tỏi) khi để lâu vẫn có thể chuyển sang trạng thái nảy mầm; quá trình nảy mầm tiêu tốn lượng lớn dinh dưỡng và đặc biệt ở khoai tây mọc mầm còn tích tụ hàm lượng cao chất độc solanine rất nguy hiểm cho sức khỏe.
  + Ngoài ra, dù ở nhiệt độ thấp, vi sinh vật kị khí hoặc nấm mốc chịu lạnh sau một thời gian dài vẫn thích nghi và phát triển, tiết ra các enzyme phân giải chất hữu cơ gây biến đổi màu sắc, mùi vị (ôi thiu, chua, biến chất) và có thể sinh ra các độc tố gây ngộ độc.""",
        "blank_lines_student": 8,
        "blank_lines_screen": 14
    }
]
