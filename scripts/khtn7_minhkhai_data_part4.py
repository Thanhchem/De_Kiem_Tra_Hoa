# -*- coding: utf-8 -*-
"""
Dữ liệu Đề Cương Ôn Tập Giữa Kì I KHTN 7 - Trường THCS Nguyễn Thị Minh Khai
PHẦN B: PHẦN TỰ LUẬN (CÂU 1 -> CÂU 18)
Đầy đủ nội dung câu hỏi, câu hỏi phụ và lời giải chi tiết sư phạm chuẩn mực.
"""

ESSAY_QUESTIONS = [
    {
        "id": 1,
        "content": "Câu 1.\na) Nêu ý nghĩa vật lí của tốc độ chuyển động.\nb) Viết công thức tính tốc độ chuyển động, giải thích rõ tên gọi và đơn vị đo của các đại lượng có trong công thức.",
        "solution": """a) Ý nghĩa vật lí của tốc độ chuyển động:
- Tốc độ chuyển động là đại lượng vật lí đặc trưng cho mức độ nhanh hay chậm của chuyển động.
- Trong cùng một khoảng thời gian, vật nào đi được quãng đường dài hơn thì chuyển động nhanh hơn (tốc độ lớn hơn). Ngược lại, để đi cùng một quãng đường, vật nào mất ít thời gian hơn thì chuyển động nhanh hơn.

b) Công thức tính tốc độ chuyển động:
*v = s / t*
Trong đó:
- *v*: Tốc độ chuyển động của vật (đơn vị: mét trên giây ~m/s~ hoặc kilômét trên giờ ~km/h~).
- *s*: Quãng đường vật đi được (đơn vị: mét ~m~ hoặc kilômét ~km~).
- *t*: Thời gian vật đi hết quãng đường đó (đơn vị: giây ~s~ hoặc giờ ~h~).""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 2,
        "content": "Câu 2.\na) Liệt kê các đơn vị đo tốc độ thường dùng trong Hệ đo lường chính thức của nước ta.\nb) Trình bày mối liên hệ và cách đổi đơn vị từ m/s sang km/h và ngược lại.",
        "solution": """a) Các đơn vị đo tốc độ thường dùng trong Hệ đo lường chính thức của nước ta:
- Đơn vị cơ bản (hợp pháp theo chuẩn SI): mét trên giây (~m/s~).
- Đơn vị thường dùng trong đời sống và giao thông: kilômét trên giờ (~km/h~).
- Ngoài ra còn có một số đơn vị khác như: xentimét trên giây (~cm/s~), mét trên phút (~m/phút~).

b) Mối liên hệ và quy tắc chuyển đổi đơn vị:
- Ta có: 1 km = 1000 m và 1 h = 3600 s.
  Do đó: 1 km/h = 1000 m / 3600 s = 1/3,6 m/s ≈ 0,28 m/s.
  Suy ra: *1 m/s = 3,6 km/h*.
- Quy tắc đổi:
  + Từ *m/s* đổi sang *km/h*: nhân với 3,6 (Ví dụ: 10 m/s = 10 × 3,6 = 36 km/h).
  + Từ *km/h* đổi sang *m/s*: chia cho 3,6 (Ví dụ: 72 km/h = 72 / 3,6 = 20 m/s).""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 3,
        "content": "Câu 3. Bạn A đi từ nhà đến trường bằng xe đạp với tốc độ không đổi. Bạn xuất phát từ nhà lúc 6 h 45 min và đến trường lúc 7 h 15 min. Biết quãng đường từ nhà bạn A đến trường dài 5 km.\na) Tính thời gian bạn A đi từ nhà đến trường theo đơn vị giờ (h).\nb) Tính tốc độ đi xe đạp của bạn A ra đơn vị km/h và m/s (làm tròn kết quả m/s đến 1 chữ số thập phân).",
        "solution": """a) Thời gian bạn A đi từ nhà đến trường:
- Thời gian đi bằng thời điểm đến trừ thời điểm xuất phát:
  *t* = 7 h 15 min - 6 h 45 min = 30 min (phút).
- Đổi đơn vị sang giờ:
  *t* = 30 / 60 = *0,5 h*.

b) Tốc độ đi xe đạp của bạn A:
- Tốc độ theo đơn vị km/h:
  *v = s / t* = 5 / 0,5 = *10 km/h*.
- Đổi tốc độ sang đơn vị m/s:
  *v* = 10 / 3,6 ≈ *2,8 m/s*.""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 4,
        "content": "Câu 4. Camera của một thiết bị bắn tốc độ ghi nhận thời gian một ô tô chạy qua hai vạch mốc cách nhau 10 m trên đường là 0,50 s.\na) Tính tốc độ chuyển động của ô tô ra đơn vị m/s và km/h.\nb) Nếu tốc độ giới hạn cho phép của cung đường là 60 km/h thì ô tô này có vượt quá tốc độ quy định không? Vì sao?",
        "solution": """a) Tốc độ chuyển động của ô tô:
- Tốc độ theo đơn vị m/s:
  *v = s / t* = 10 / 0,50 = *20 m/s*.
- Tốc độ theo đơn vị km/h:
  *v* = 20 × 3,6 = *72 km/h*.

b) Đánh giá việc chấp hành quy định tốc độ:
- Tốc độ của xe ô tô là 72 km/h, trong khi tốc độ giới hạn cho phép của cung đường là 60 km/h.
- So sánh: 72 km/h > 60 km/h.
- Kết luận: Ô tô này đã *vượt quá tốc độ quy định* (chạy quá tốc độ cho phép 12 km/h) và đã vi phạm luật giao thông đường bộ.""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 5,
        "content": "Câu 5. Một vận động viên rèn luyện thể lực bằng cách chạy bộ trên một đường thẳng gồm hai chặng:\nChặng thứ nhất: Chạy quãng đường dài 600 m hết thời gian 2 phút 30s.\nChặng thứ hai: Chạy quãng đường dài 400 m với tốc độ không đổi là 2 m/s.\na) Tính tốc độ chạy của vận động viên trong chặng thứ nhất theo đơn vị m/s?\nb) Tính thời gian vận động viên chạy hết chặng thứ hai và tính tốc độ trung bình của vận động viên trên toàn bộ quãng đường chạy.",
        "solution": """a) Tốc độ chạy của vận động viên trong chặng thứ nhất:
- Đổi thời gian chặng 1: *t~1~* = 2 phút 30 s = 2 × 60 + 30 = 150 s.
- Tốc độ chặng 1:
  *v~1~ = s~1~ / t~1~* = 600 / 150 = *4 m/s*.

b) Thời gian chạy chặng 2 và tốc độ trung bình toàn bộ quãng đường:
- Thời gian chạy hết chặng thứ hai:
  *t~2~ = s~2~ / v~2~* = 400 / 2 = *200 s*.
- Tổng quãng đường chạy:
  *s = s~1~ + s~2~* = 600 + 400 = 1000 m.
- Tổng thời gian chạy:
  *t = t~1~ + t~2~* = 150 + 200 = 350 s.
- Tốc độ trung bình trên toàn bộ quãng đường chạy:
  *v~tb~ = s / t* = 1000 / 350 ≈ *2,86 m/s*.""",
        "blank_lines_student": 7,
        "blank_lines_screen": 7
    },
    {
        "id": 6,
        "content": "Câu 6. Trong tiết thực hành môn Khoa học tự nhiên 7, một nhóm học sinh tiến hành đo tốc độ của xe đồ chơi trượt trên mặt phẳng dốc bằng thước cuộn và đồng hồ bấm giây. Quãng đường giữa vạch xuất phát và vạch đích được cố định là s = 60 cm. Nhóm tiến hành đo thời gian 3 lần và ghi lại bảng số liệu như sau:\nLần 1: t1 = 1,18 s\nLần 2: t2 = 1,22 s\nLần 3: t3 = 1,20 s\na) Tính giá trị thời gian chuyển động trung bình của xe đồ chơi.\nb) Tính tốc độ trung bình của xe đồ chơi ra đơn vị cm/s và đổi sang đơn vị m/s.",
        "solution": """a) Thời gian chuyển động trung bình của xe đồ chơi:
- Ta tính giá trị trung bình cộng của 3 lần đo:
  *t~tb~ = (t~1~ + t~2~ + t~3~) / 3* = (1,18 + 1,22 + 1,20) / 3 = 3,60 / 3 = *1,20 s*.

b) Tốc độ trung bình của xe đồ chơi:
- Quãng đường: *s* = 60 cm = 0,6 m.
- Tốc độ trung bình tính theo đơn vị cm/s:
  *v~tb~ = s / t~tb~* = 60 / 1,20 = *50 cm/s*.
- Đổi sang đơn vị m/s:
  *v~tb~* = 50 / 100 = *0,5 m/s* (hoặc tính trực tiếp: 0,6 / 1,20 = 0,5 m/s).""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 7,
        "content": "Câu 7. Trên một đoạn đường cao tốc có biển báo chỉ dẫn: Tốc độ tối đa cho phép là 100 km/h và tốc độ tối thiểu cho phép là 60 km/h. Thiết bị bắn tốc độ tự động được lắp đặt cố định bên đường ghi nhận một ô tô con đi qua hai vạch mốc cách nhau 10m trên mặt đường hết khoảng thời gian là 0,32 s.\na) Tính tốc độ chuyển động của ô tô đó ra đơn vị m/s và km/h (kết quả tính theo km/h làm tròn đến một chữ số thập phân).\nb) Căn cứ vào quy định tốc độ của tuyến đường cao tốc trên, người lái xe ô tô này có vi phạm luật giao thông không? Vì sao?",
        "solution": """a) Tốc độ chuyển động của ô tô con:
- Tốc độ theo đơn vị m/s:
  *v = s / t* = 10 / 0,32 = *31,25 m/s*.
- Tốc độ theo đơn vị km/h:
  *v* = 31,25 × 3,6 = *112,5 km/h*.

b) Đánh giá việc chấp hành quy định giao thông:
- Tuyến đường quy định: tốc độ tối thiểu là 60 km/h và tốc độ tối đa là 100 km/h (60 km/h ≤ v ≤ 100 km/h).
- Tốc độ thực tế của xe là 112,5 km/h.
- So sánh: 112,5 km/h > 100 km/h.
- Kết luận: Người lái xe ô tô này *đã vi phạm luật giao thông* vì chạy quá tốc độ tối đa cho phép trên đường cao tốc (vượt quá 12,5 km/h).""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 8,
        "content": "Câu 8. Trong một buổi hội khỏe Phù Đổng ở trường THCS Nguyễn Thị Minh Khai, một học sinh tham gia chạy cự li 100 m. Bạn An hoàn thành quãng đường này trong thời gian 20 s. Giáo viên yêu cầu học sinh phân tích kết quả chạy để đánh giá khả năng chuyển động của vận động viên.\na) Tốc độ cho biết điều gì về chuyển động của một vật?\nb) Trong hệ đơn vị đo lường, tốc độ thường được đo bằng những đơn vị nào? Hãy nêu hai đơn vị đo tốc độ thường dùng và cho biết mối liên hệ giữa chúng.\nc) Tính tốc độ chạy trung bình của An theo đơn vị m/s và km/h.",
        "solution": """a) Ý nghĩa của tốc độ:
- Tốc độ cho biết mức độ nhanh hay chậm của chuyển động của một vật. Vật có tốc độ càng lớn thì chuyển động càng nhanh.

b) Đơn vị đo tốc độ và mối liên hệ:
- Hai đơn vị thường dùng nhất: mét trên giây (~m/s~) và kilômét trên giờ (~km/h~).
- Mối liên hệ: 1 m/s = 3,6 km/h (hoặc 1 km/h ≈ 0,28 m/s).

c) Tính tốc độ chạy trung bình của bạn An:
- Tốc độ theo đơn vị m/s:
  *v = s / t* = 100 / 20 = *5 m/s*.
- Tốc độ theo đơn vị km/h:
  *v* = 5 × 3,6 = *18 km/h*.""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 9,
        "content": "Câu 9. Trong một buổi trải nghiệm an toàn giao thông, học sinh sử dụng đồng hồ bấm giây và cột mốc trên sân trường để xác định tốc độ chuyển động của xe đạp. Một bạn đi xe đạp chuyển động đều, đi được quãng đường 150 m trong thời gian 30 s.\na) Đại lượng tốc độ có ý nghĩa vật lí như thế nào đối với chuyển động của xe đạp?\nb) Một bạn ghi tốc độ của xe đạp là 18 km/h, trong khi bạn khác ghi là 5 m/s. Hai cách ghi này có thể biểu thị cùng một tốc độ hay không? Giải thích.\nc) Dựa vào số liệu thực nghiệm, hãy tính tốc độ của xe đạp theo đơn vị m/s.",
        "solution": """a) Ý nghĩa vật lí của tốc độ đối với xe đạp:
- Cho biết xe đạp đang chuyển động nhanh hay chậm, thể hiện quãng đường xe đạp đi được trong 1 giây hoặc trong 1 giờ.

b) So sánh hai cách ghi:
- Hai cách ghi này *hoàn toàn biểu thị cùng một tốc độ*.
- Giải thích: Vì 1 m/s = 3,6 km/h nên khi đổi 5 m/s sang km/h ta được: 5 × 3,6 = 18 km/h. Do đó giá trị 5 m/s và 18 km/h là tương đương nhau, chỉ khác nhau về đơn vị đo.

c) Tính tốc độ xe đạp theo số liệu thực nghiệm:
- *v = s / t* = 150 / 30 = *5 m/s*.""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 10,
        "content": "Câu 10.\na) Hãy trình bày mô hình nguyên tử theo Rutherford - Bohr (nêu rõ cấu tạo hạt nhân, vỏ nguyên tử và quy tắc sắp xếp các electron trong các lớp vỏ).\nb) Nguyên tử Nitrogen (nitơ) có điện tích hạt nhân là +7. Hãy xác định số proton và số electron trong nguyên tử Nitrogen, từ đó cho biết Nitrogen có bao nhiêu lớp electron và mấy electron ở lớp ngoài cùng?",
        "solution": """a) Mô hình nguyên tử theo Rutherford - Bohr:
- *Hạt nhân*: Nằm ở tâm nguyên tử, mang điện tích dương, gồm các hạt proton (mang điện tích dương) và neutron (không mang điện). Khối lượng nguyên tử tập trung chủ yếu ở hạt nhân.
- *Vỏ nguyên tử*: Gồm các hạt electron mang điện tích âm, chuyển động xung quanh hạt nhân trên những quỹ đạo xác định tạo thành các lớp electron.
- *Quy tắc sắp xếp electron vào các lớp vỏ*:
  + Các electron được sắp xếp lần lượt từ lớp trong cùng (gần hạt nhân nhất) ra lớp ngoài.
  + Lớp thứ nhất (trong cùng) chứa tối đa *2 electron*.
  + Lớp thứ hai chứa tối đa *8 electron*.

b) Xác định cấu tạo nguyên tử Nitrogen:
- Điện tích hạt nhân là +7 ──> Số proton trong hạt nhân là *p = 7*.
- Vì nguyên tử trung hòa về điện nên số electron ở lớp vỏ bằng số proton: *e = p = 7*.
- Sự phân bố electron vào các lớp:
  + Lớp thứ nhất chứa tối đa: *2 electron*.
  + Số electron còn lại ở lớp thứ hai: 7 - 2 = *5 electron*.
- Kết luận: Nguyên tử Nitrogen có *2 lớp electron* và có *5 electron ở lớp ngoài cùng*.""",
        "blank_lines_student": 7,
        "blank_lines_screen": 7
    },
    {
        "id": 11,
        "content": "Câu 11. Một nguyên tử X có tổng số hạt trong hạt nhân là 23 hạt, trong đó số hạt không mang điện là 12 hạt.\na) Xác định số proton và số electron của nguyên tử X.\nb) Cho biết nguyên tử X thuộc loại nguyên tử nào quen thuộc trong đời sống?",
        "solution": """a) Xác định số proton và electron của nguyên tử X:
- Trong hạt nhân chỉ gồm hai loại hạt: proton (mang điện dương) và neutron (không mang điện).
- Theo đề bài, tổng số hạt trong hạt nhân là 23 và số hạt không mang điện (neutron) là 12:
  *Số proton (p)* = Tổng số hạt trong hạt nhân - Số neutron = 23 - 12 = *11 hạt*.
- Do nguyên tử trung hòa về điện nên:
  *Số electron (e) = p = 11 hạt*.

b) Tên nguyên tử X:
- Nguyên tử có 11 proton trong hạt nhân chính là nguyên tố *Sodium* (kí hiệu hóa học: *Na*, tên tiếng Việt thường gọi là natri).
- Đây là nguyên tố kim loại quen thuộc tạo nên thành phần của muối ăn (NaCl - sodium chloride) dùng hàng ngày trong đời sống.""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 12,
        "content": "Câu 12. Nguyên tử Carbon (C) có 6 proton và nguyên tử Sodium (Na) có 11 electron.\na) Xác định số electron của nguyên tử Carbon và số proton của nguyên tử Sodium.\nb) Cho biết số lớp electron và số electron ở lớp ngoài cùng của nguyên tử Carbon và Sodium (biết lớp 1 chứa tối đa 2 electron, lớp 2 chứa tối đa 8 electron).",
        "solution": """a) Xác định số hạt còn lại:
- Do nguyên tử trung hòa về điện (số electron = số proton):
  + Nguyên tử Carbon có 6 proton ──> Số electron của Carbon là *e = 6*.
  + Nguyên tử Sodium có 11 electron ──> Số proton của Sodium là *p = 11*.

b) Sự phân bố electron vào các lớp vỏ:
- *Nguyên tử Carbon (6 electron)*:
  + Lớp 1 có: 2 electron.
  + Lớp 2 có: 6 - 2 = 4 electron.
  ──> Carbon có *2 lớp electron* và có *4 electron ở lớp ngoài cùng*.
- *Nguyên tử Sodium (11 electron)*:
  + Lớp 1 có: 2 electron.
  + Lớp 2 có: 8 electron.
  + Lớp 3 có: 11 - (2 + 8) = 1 electron.
  ──> Sodium có *3 lớp electron* và có *1 electron ở lớp ngoài cùng*.""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 13,
        "content": "Câu 13.\na) Trình bày mô hình nguyên tử Rutherford – Bohr. Electron được sắp xếp như thế nào trong các lớp vỏ nguyên tử?\nb) Nguyên tử carbon (C) có 6 proton và trung hòa về điện. Hãy xác định số electron của nguyên tử carbon và cho biết sự phân bố electron vào các lớp vỏ.",
        "solution": """a) Mô hình nguyên tử Rutherford – Bohr và quy tắc sắp xếp electron:
- Nguyên tử gồm hạt nhân mang điện tích dương ở trung tâm và các electron mang điện tích âm quay quanh hạt nhân theo các quỹ đạo tròn xác định.
- Các electron được phân bố thành từng lớp từ trong ra ngoài:
  + Lớp 1 (trong cùng) chứa tối đa *2 electron*.
  + Lớp 2 chứa tối đa *8 electron*.
  + Lớp 3 chứa tối đa *18 electron* (đối với 20 nguyên tố đầu, khi có 8 electron ở lớp 3 thì lớp này đã đạt cấu hình bền vững).

b) Cấu tạo và sự phân bố electron của nguyên tử carbon:
- Nguyên tử carbon trung hòa về điện có 6 proton ──> Số electron là *e = 6*.
- Sự phân bố electron vào các lớp vỏ:
  + Lớp thứ nhất (lớp 1): chứa *2 electron*.
  + Lớp thứ hai (lớp 2): chứa *4 electron*.
  ──> Carbon có 2 lớp electron và 4 electron lớp ngoài cùng.""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 14,
        "content": "Câu 14.\na) Mô tả sự sắp xếp electron trong các lớp vỏ nguyên tử theo mô hình Rutherford – Bohr.\nb) Nguyên tử sodium (Na) có 11 proton và trung hòa về điện. Hãy xác định số electron và sự phân bố electron của nguyên tử sodium vào các lớp vỏ.",
        "solution": """a) Mô tả sự sắp xếp electron theo mô hình Rutherford – Bohr:
- Các electron mang điện tích âm chuyển động xung quanh hạt nhân và được xếp vào các lớp vỏ theo thứ tự mức năng lượng tăng dần từ trong ra ngoài.
- Lớp 1 trong cùng chỉ chứa tối đa 2 electron. Khi lớp 1 đầy, electron tiếp theo sẽ điền vào lớp 2 (tối đa 8 electron). Khi lớp 2 đầy, electron tiếp tục điền vào lớp 3.

b) Cấu tạo và sự phân bố electron của nguyên tử sodium:
- Số proton p = 11, vì nguyên tử trung hòa về điện nên số electron là *e = 11*.
- Phân bố 11 electron vào các lớp vỏ:
  + Lớp 1: chứa *2 electron* (đã bão hòa).
  + Lớp 2: chứa *8 electron* (đã bão hòa).
  + Lớp 3: chứa *1 electron* (11 - 2 - 8 = 1).
- Kết luận: Sodium có *3 lớp electron* với cấu hình phân bố (2, 8, 1) và có *1 electron ở lớp ngoài cùng*.""",
        "blank_lines_student": 6,
        "blank_lines_screen": 6
    },
    {
        "id": 15,
        "content": "Câu 15.\na) Phát biểu khái niệm quá trình hô hấp tế bào ở sinh vật.\nb) Viết phương trình hô hấp tế bào ở dạng chữ. Dựa vào phương trình dạng chữ, hãy phân tích mối quan hệ hai chiều giữa tổng hợp và phân giải chất hữu cơ trong tế bào.",
        "solution": """a) Khái niệm hô hấp tế bào:
- Hô hấp tế bào là quá trình phân giải chất hữu cơ (chủ yếu là glucose) diễn ra bên trong tế bào với sự tham gia của khí oxygen, tạo ra sản phẩm là khí carbon dioxide, nước và giải phóng năng lượng cung cấp cho các hoạt động sống của sinh vật (dưới dạng ATP và nhiệt năng).

b) Phương trình hô hấp tế bào ở dạng chữ và phân tích mối quan hệ:
- Phương trình dạng chữ:
  *Glucose + Khí oxygen ──(Enzyme)──> Khí carbon dioxide + Nước + Năng lượng (ATP + nhiệt)*
- Mối quan hệ hai chiều giữa tổng hợp và phân giải chất hữu cơ:
  + *Tổng hợp (Quang hợp)*: Tế bào sử dụng năng lượng ánh sáng để tổng hợp chất hữu cơ (glucose) và tích lũy hóa năng từ các chất vô cơ đơn giản (CO~2~, H~2~O). Quá trình này tạo ra nguyên liệu cho hô hấp.
  + *Phân giải (Hô hấp tế bào)*: Tế bào sử dụng chất hữu cơ và giải phóng hóa năng tích lũy trong chất hữu cơ thành năng lượng ATP phục vụ cho các hoạt động sống (kể cả quá trình tổng hợp lại các chất hữu cơ đặc trưng cho tế bào).
  ──> Hai quá trình này tuy trái ngược nhau về chiều hướng biến đổi vật chất và năng lượng nhưng gắn bó mật thiết, phụ thuộc và bổ trợ lẫn nhau để duy trì sự sống.""",
        "blank_lines_student": 7,
        "blank_lines_screen": 7
    },
    {
        "id": 16,
        "content": "Câu 16.\na) Nêu các yếu tố môi trường ảnh hưởng đến hô hấp tế bào.\nb) Dựa vào kiến thức về ảnh hưởng của các yếu tố này, hãy giải thích vì sao trong thực tiễn nông nghiệp, sau khi thu hoạch các loại hạt (lúa, ngô, đỗ), người nông dân cần phải phơi khô hoặc sấy khô hạt trước khi cất giữ vào kho bảo quản.",
        "solution": """a) Các yếu tố môi trường ảnh hưởng đến hô hấp tế bào:
1. *Hàm lượng nước*: Nước vừa là dung môi vừa tham gia trực tiếp vào các phản ứng hóa sinh. Độ ẩm cao làm tăng hô hấp tế bào; độ ẩm thấp làm hô hấp giảm mạnh.
2. *Nhiệt độ*: Nhiệt độ ảnh hưởng đến hoạt tính enzyme hô hấp. Nhiệt độ tối ưu thường từ 30°C - 35°C; nhiệt độ quá thấp hoặc quá cao đều ức chế hô hấp.
3. *Nồng độ khí oxygen*: Oxygen là nguyên liệu trực tiếp của hô hấp hiếu khí. Nồng độ O~2~ giảm dưới 5% sẽ làm suy giảm hô hấp.
4. *Nồng độ khí carbon dioxide*: CO~2~ là sản phẩm của hô hấp. Nồng độ CO~2~ trong không khí cao (trên 3 - 5%) sẽ ức chế quá trình hô hấp.

b) Giải thích lí do phơi khô hoặc sấy khô hạt nông sản:
- Khi mới thu hoạch, hạt lúa, ngô, đậu có độ ẩm cao (khoảng 20 - 25%), quá trình hô hấp tế bào diễn ra rất mạnh mẽ.
- Hô hấp mạnh sẽ làm:
  + Tiêu hao nhanh lượng chất hữu cơ (tinh bột, protein) tích lũy trong hạt, làm giảm phẩm chất nông sản.
  + Tỏa nhiều nhiệt và hơi ẩm làm nhiệt độ khối hạt tăng cao, tạo môi trường thuận lợi cho nấm mốc, vi khuẩn và sâu mọt phát triển phá hoại hạt.
  + Kích thích hạt nảy mầm sớm làm hạt bị hỏng, mất khả năng làm giống.
- Vì vậy, phơi khô hoặc sấy khô giúp giảm độ ẩm của hạt xuống mức an toàn (dưới 13%), đưa hạt về trạng thái ngủ nghỉ (ức chế hô hấp tối đa), giúp bảo quản hạt được lâu dài mà không bị hư hỏng hay hao hụt khối lượng.""",
        "blank_lines_student": 8,
        "blank_lines_screen": 8
    },
    {
        "id": 17,
        "content": "Câu 17. Một nhóm học sinh tiến hành thí nghiệm: Lấy 100g hạt đậu xanh nảy mầm cho vào bình thủy tinh A và 100g hạt đậu xanh khô (chưa ngâm nước) cho vào bình thủy tinh B. Cắm vào mỗi bình một nhiệt kế, đậy kín nắp và để bọc xốp cách nhiệt bên ngoài. Sau 4 giờ, nhiệt kế ở bình A tăng thêm 3°C, trong khi nhiệt kế ở bình B gần như không thay đổi.\na) Hiện tượng tăng nhiệt độ ở bình A chứng tỏ các tế bào hạt đậu đang thực hiện quá trình sinh lí nào? Hãy nêu khái niệm của quá trình đó và viết phương trình dạng chữ thể hiện hai chiều tổng hợp - phân giải để giải thích nguồn gốc chất hữu cơ trong hạt đậu và cách hạt sử dụng chúng.\nb) Từ kết quả thí nghiệm trên, em hãy nêu tên các yếu tố môi trường có ảnh hưởng đến hô hấp tế bào. Nếu là một người nông dân cần cất giữ hàng tấn hạt giống qua mùa sau, em sẽ tác động vào những yếu tố nào để hạt không bị hỏng?",
        "solution": """a) Quá trình sinh lí và phương trình hai chiều:
- Hiện tượng nhiệt độ tăng ở bình A chứng tỏ các tế bào hạt đậu đang nảy mầm thực hiện *quá trình hô hấp tế bào*. Trong quá trình này, một phần năng lượng giải phóng từ chất hữu cơ được tỏa ra dưới dạng nhiệt năng.
- *Khái niệm*: Hô hấp tế bào là quá trình phân giải chất hữu cơ giải phóng năng lượng cung cấp cho các hoạt động sống của tế bào và cơ thể.
- *Phương trình chữ thể hiện mối quan hệ hai chiều*:
  Quang hợp (Tổng hợp ở cây mẹ): *CO~2~ + H~2~O ──(Ánh sáng / Diệp lục)──> Glucose + O~2~*
  Hô hấp (Phân giải ở hạt nảy mầm): *Glucose + O~2~ ──(Enzyme)──> CO~2~ + H~2~O + Năng lượng (ATP + nhiệt)*
  ──> Nguồn gốc chất hữu cơ trong hạt đậu do cây đậu mẹ quang hợp ở lá rồi tích lũy vào hạt. Khi nảy mầm, phôi hạt sử dụng chính chất hữu cơ đó qua hô hấp tế bào để tạo năng lượng phát triển thành cây con.

b) Các yếu tố ảnh hưởng và biện pháp bảo quản hạt giống của người nông dân:
- *Các yếu tố môi trường ảnh hưởng*: Nước (độ ẩm), nhiệt độ, nồng độ O~2~, nồng độ CO~2~.
- *Biện pháp tác động để bảo quản hạt giống*:
  1. *Giảm độ ẩm (tác động vào hàm lượng nước)*: Phơi thật khô hoặc sấy hạt đạt độ ẩm chuẩn (dưới 12 - 13%) trước khi đóng bao để giảm hô hấp xuống mức tối thiểu.
  2. *Hạ nhiệt độ*: Bảo quản ở nơi khô ráo, thoáng mát hoặc trong kho lạnh chuyên dụng (nhiệt độ thấp ức chế enzyme hô hấp).
  3. *Điều chỉnh nồng độ khí (giảm O~2~, tăng CO~2~)*: Bảo quản hạt trong chum vại đậy kín, túi hút chân không hoặc bơm khí nitrogen/CO~2~ để ngăn hô hấp hiếu khí và diệt sâu mọt.""",
        "blank_lines_student": 8,
        "blank_lines_screen": 8
    },
    {
        "id": 18,
        "content": "Câu 18. Tại Huế, những đợt mưa lớn kéo dài nhiều ngày trong mùa mưa bão thường gây ra tình trạng ngập lụt diện rộng. Vừa qua, một vườn thanh trà - đặc sản nổi tiếng ở phường Thủy Biều bị ngâm chìm trong nước lũ đục ngầu suốt một tuần liền. Dù phần tán lá phía trên vẫn được tiếp xúc với ánh sáng, nhưng toàn bộ rễ cây bám dưới đất lại bị ngập sâu. Hậu quả lá cây bắt đầu chuyển sang màu vàng, héo rũ, rễ cây thối đen và cây chết hàng loạt. Cùng lúc đó, kết quả đo lường cho thấy nồng độ khí oxygen trong đất ở vườn bưởi lúc này đã sụt giảm nghiêm trọng xuống mức gần 0%.\na) Rễ cây thanh trà dù nằm sâu dưới mặt đất hay bị ngập nước thì vẫn luôn cần năng lượng để duy trì sự sống và hút khoáng. Hãy nêu khái niệm quá trình sinh lí tạo ra năng lượng này tại tế bào rễ. Viết phương trình dạng chữ thể hiện mối quan hệ hai chiều tổng hợp vật chất (diễn ra ở lá) và phân giải vật chất (diễn ra ở rễ).\nb) Nêu một số yếu tố môi trường có ảnh hưởng trực tiếp đến quá trình hô hấp tế bào. Từ đó, em hãy giải thích nguyên nhân khoa học khiến những cây thanh trà Thủy Biều bị chết trong tình huống ngập úng kéo dài nêu trên.",
        "solution": """a) Quá trình sinh lí tạo năng lượng và phương trình hai chiều:
- Quá trình sinh lí tạo năng lượng tại tế bào rễ là *quá trình hô hấp tế bào*.
- *Phương trình chữ thể hiện mối quan hệ tổng hợp (ở lá) và phân giải (ở rễ)*:
  + Tại tế bào lá (Quang hợp - Tổng hợp):
    *Nước + Khí carbon dioxide ──(Ánh sáng / Diệp lục)──> Glucose + Khí oxygen*
  + Tại tế bào rễ (Hô hấp tế bào - Phân giải):
    *Glucose + Khí oxygen ──(Enzyme)──> Nước + Khí carbon dioxide + Năng lượng (ATP)*
  ──> Chất hữu cơ (glucose) được tạo ra từ lá qua hệ mạch rây vận chuyển xuống rễ. Tại rễ, tế bào phân giải glucose để tạo ra năng lượng ATP cung cấp cho hoạt động sinh trưởng của rễ và quá trình chủ động hấp thụ các ion khoáng từ đất.

b) Các yếu tố ảnh hưởng và giải thích nguyên nhân cây thanh trà Thủy Biều bị chết:
- *Các yếu tố môi trường ảnh hưởng trực tiếp*: Nồng độ oxygen trong đất, hàm lượng nước (độ ngập úng), nhiệt độ, nồng độ CO~2~.
- *Giải thích nguyên nhân khoa học khiến cây chết khi ngập úng*:
  1. *Thiếu hụt oxygen nghiêm trọng*: Khi đất ngập nước kéo dài suốt một tuần, các mao quản và khe hở trong đất bị nước chiếm chỗ hoàn toàn, lượng oxygen hòa tan sụt giảm về gần 0%.
  2. *Hô hấp tế bào ở rễ bị đình trệ*: Do thiếu O~2~, các tế bào rễ không thể thực hiện hô hấp hiếu khí để tạo đủ năng lượng ATP. Rễ phải chuyển sang hô hấp kị khí (lên men), tạo ra các chất độc hại như rượu ethanol, acid hữu cơ làm đầu rễ bị ngộ độc và thối đen.
  3. *Lông hút bị chết, mất khả năng hút nước*: Thiếu ATP làm các tế bào lông hút bị tổn thương và chết hàng loạt. Cây mất hoàn toàn khả năng hút nước và các chất khoáng cần thiết.
  4. *Cây chết do thiếu nước và dinh dưỡng*: Mặc dù rễ bị ngâm trong biển nước nhưng vì lông hút đã hỏng nên nước không thể vận chuyển lên tán lá (hiện tượng hạn sinh lí). Dưới tác động thoát hơi nước ở lá, cây bị mất nước nghiêm trọng, lá úa vàng, héo rũ và dẫn đến cây thanh trà chết hàng loạt.""",
        "blank_lines_student": 8,
        "blank_lines_screen": 8
    }
]
