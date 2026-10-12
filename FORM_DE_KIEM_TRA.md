# QUY CHUẨN FORM ĐỀ KIỂM TRA HÓA HỌC (CHUẨN BỘ GD&ĐT & THẦY TRẦN VĂN THẠNH)

Tài liệu này ghi nhớ và quy chuẩn hóa toàn bộ cấu trúc, quy cách định dạng, căn tab, header, footer và phong cách biên soạn cho **FORM ĐỀ KIỂM TRA**. Khi người dùng gửi hình ảnh hoặc nội dung đề thi mới, AI sẽ tự động trích xuất, tái tạo và xuất bản đề thi theo đúng chuẩn mực này.

---

## 1. THÔNG SỐ ĐỊNH DẠNG TRANG & VĂN BẢN (PAGE SETUP & STYLES)

- **Khổ giấy:** A4 chuẩn quốc tế (`21,0 cm × 29,7 cm`).
- **Căn lề (Margins):**
  - Lề trên (Top): `1,5 cm`
  - Lề dưới (Bottom): `1,5 cm`
  - Lề trái (Left): `1,5 cm`
  - Lề phải (Right): `1,5 cm`
  - *Độ rộng vùng nội dung soạn thảo khả dụng:* đúng `18,0 cm`.
- **Khoảng cách mép giấy:**
  - Header margin: `0,6 cm`
  - Footer margin: `0,6 cm`
- **Font chữ chính:** `Times New Roman`.
- **Cỡ chữ & giãn cách:**
  - Tiêu đề đề thi: `10 pt – 11 pt` (in đậm).
  - Tên phần (Phần I, II, III, IV): `11,5 pt` (in đậm, màu xanh lục đậm `#2E7D32`).
  - Lời dẫn phần thi: `10,5 pt` (in nghiêng).
  - Nội dung câu hỏi: `10,5 pt` (tiền tố `Câu X.` in đậm, thân câu in thường).
  - Phương án trắc nghiệm A, B, C, D: `10,5 pt` (nhãn `A.`, `B.`, `C.`, `D.` in đậm).
  - Lời giải chi tiết: `10 pt` (tiêu đề `► Lời giải chi tiết:` `10,5 pt` in đậm màu xanh `#1565C0`).
  - Giãn dòng: `Multiple 1.15` (hoặc `1.1`), Space Before = `1 pt`, Space After = `2 pt`.

---

## 2. QUY CÁCH HEADER & FOOTER ĐẶC TRƯNG

### 2.1. Header (Đầu trang)
- Tạo bảng **1 hàng 2 cột**, chiều rộng bảng đúng `18,0 cm` (khớp hoàn toàn lề giấy), căn giữa:
  - **Cột trái (rộng 8,0 cm):** `Tài liệu lưu hành nội bộ` (Chữ in nghiêng, cỡ `10 pt`, màu đỏ san hô Coral Red: `#D32F2F`, căn trái).
  - **Cột phải (rộng 10,0 cm):** `Thầy TRẦN VĂN THẠNH - 0777.470.803` (Cỡ `10 pt`, màu đỏ san hô `#D32F2F`, tên `TRẦN VĂN THẠNH` in đậm, căn phải).
- **Đường kẻ dưới (Border Bottom):** Đường kẻ đơn màu xanh lá Sage Green (`#8BC390`), độ dày `0,75 pt` (`sz="6"` hoặc `sz="8"` trong XML docx), cách chữ `4 pt`.

### 2.2. Footer (Chân trang)
- Tạo bảng **1 hàng 2 cột**, chiều rộng bảng đúng `18,0 cm`, căn giữa:
  - **Cột trái (rộng 14,5 cm):** `CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế   |   CS2: 24 Đặng Thái Thân, TP Huế` (Chữ in nghiêng, cỡ `9,5 pt`, màu đỏ san hô `#D32F2F`, căn trái).
  - **Cột phải (rộng 3,5 cm):** `Trang X` (Sử dụng trường số trang tự động Word `PAGE`, chữ in đậm, cỡ `10 pt`, màu đỏ san hô `#D32F2F`, căn phải).
- **Đường kẻ trên (Border Top):** Đường kẻ đơn màu xanh lá Sage Green (`#8BC390`), độ dày `0,75 pt`, cách chữ `4 pt`.

---

## 3. KHỐI TIÊU ĐỀ ĐỀ THI (TITLE BLOCK ĐẦU TRANG 1)

Khối tiêu đề được bố trí bằng **đoạn văn bản kết hợp Tab Stop phải ở vị trí 18,0 cm** (`WD_TAB_ALIGNMENT.RIGHT`) để đảm bảo không bị lỗi viền bảng:
- **Dòng 1:** `TÊN SỞ GD&ĐT / TRƯỜNG...` *(trái, 10 pt, in đậm)* `\t` `TÊN ĐỀ THI / ĐỀ ÔN TẬP SỐ...` *(phải, 11 pt, in đậm)*
- **Dòng 2:** `ĐỀ CHÍNH THỨC` *(trái, 10,5 pt, in đậm)* `\t` `MÔN: HOÁ HỌC - LỚP...` *(phải, 11 pt, in đậm)*
- **Dòng 3:** `(Đề thi có ... trang)` *(trái, 9,5 pt, in nghiêng)* `\t` `Thời gian làm bài: ... phút | MÃ ĐỀ: ...` *(phải, 10 pt, in đậm)*
- *(Nếu là Bản Giáo viên)*: Thêm dòng căn giữa màu đỏ nổi bật:
  `(BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)` *(11 pt, in đậm, màu đỏ `#C62828`)*.
- **Dòng thông tin thí sinh:**
  `Họ, tên thí sinh: ............................................................ Lớp: ............. Số báo danh: .............`
- **Dòng nguyên tử khối:**
  `_(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ba = 137)._`
- **Đường kẻ phân cách:** Dưới khối tiêu đề có đường kẻ ngang mảnh màu xám `#CCCCCC` (`sz="4"`).

---

## 4. QUY CÁCH NỘI DUNG 4 PHẦN THEO FORM ĐỀ THI CHUẨN

### 4.1. PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN
- **Tiêu đề:** `PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.` (Màu `#2E7D32`, in đậm).
- **Lời dẫn:** `Thí sinh trả lời từ câu ... đến câu ... Mỗi câu hỏi thí sinh chỉ chọn một phương án.`
- **Căn chỉnh phương án A, B, C, D (TUYỆT ĐỐI KHÔNG DÙNG BẢNG - NO TABLE):**
  - **Dạng 1 (Phương án ngắn, chiều dài tối đa ≤ 18 ký tự):** Dàn trên **1 dòng duy nhất** với 3 điểm dừng Tab lề trái:
    - Tab 1: `4,5 cm` (bắt đầu phương án B)
    - Tab 2: `9,0 cm` (bắt đầu phương án C)
    - Tab 3: `13,5 cm` (bắt đầu phương án D)
    - Cú pháp: `A. Phương án A \t B. Phương án B \t C. Phương án C \t D. Phương án D`
  - **Dạng 2 (Phương án vừa, 19 ký tự ≤ chiều dài ≤ 38 ký tự):** Dàn trên **2 dòng**, mỗi dòng 2 phương án với 1 điểm dừng Tab lề trái:
    - Tab: `9,0 cm` (bắt đầu phương án B hoặc D)
    - Dòng 1: `A. Phương án A \t B. Phương án B`
    - Dòng 2: `C. Phương án C \t D. Phương án D`
  - **Dạng 3 (Phương án dài, > 38 ký tự hoặc phương án xuống dòng):** Dàn trên **4 dòng riêng biệt**:
    - Thụt lề trái đoạn văn (`Left Indent = 0,5 cm`).
    - Dòng 1: `A. Phương án A`
    - Dòng 2: `B. Phương án B`
    - Dòng 3: `C. Phương án C`
    - Dòng 4: `D. Phương án D`
- **Quy tắc Bản Giáo viên:** Phương án đúng được **in đậm và tô màu đỏ `#C62828`** cho cả chữ cái nhãn (`*A.*`, `*B.*`, `*C.*`, `*D.*`) và nội dung chữ.

### 4.2. PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI
- **Tiêu đề:** `PHẦN II. Câu trắc nghiệm đúng sai.` (Màu `#2E7D32`, in đậm).
- **Lời dẫn:** `Thí sinh trả lời từ câu ... đến câu ... Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.`
- **Thang điểm chuẩn:**
  - Đúng 1 ý: `0,1 điểm`
  - Đúng 2 ý: `0,25 điểm`
  - Đúng 3 ý: `0,5 điểm`
  - Đúng 4 ý: `1,0 điểm`
- **Cấu trúc mỗi câu:**
  - Lời dẫn/Bối cảnh/Đoạn văn hóa học thực nghiệm: `Câu X. [Nội dung dẫn dắt hoặc hình vẽ minh họa]`
  - Bốn mệnh đề:
    - `a) [Mệnh đề a]`
    - `b) [Mệnh đề b]`
    - `c) [Mệnh đề c]`
    - `d) [Mệnh đề d]`
- **Quy tắc Bản Giáo viên:** Sau mỗi ý hiển thị mũi tên và kết luận: `➔ (Đúng)` hoặc `➔ (Sai)` in đậm màu đỏ `#C62828`.

### 4.3. PHẦN III. CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN
- **Tiêu đề:** `PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn.` (Màu `#2E7D32`, in đậm).
- **Lời dẫn:** `Thí sinh trả lời từ câu ... đến câu ...`
- **Mỗi câu:** Câu hỏi tính toán hoặc đếm số lượng phát biểu, yêu cầu kết quả số (số nguyên hoặc số thập phân làm tròn theo đề bài).
- **Quy tắc Bản Giáo viên:** Hiển thị rõ: `► Đáp án: [Giá trị số]` in đậm màu đỏ `#C62828`.

### 4.4. PHẦN IV. TỰ LUẬN (NẾU CÓ)
- **Tiêu đề:** `PHẦN IV. Tự luận.` (Màu `#2E7D32`, in đậm).
- **Mỗi câu:** Kèm thang điểm cụ thể ở đầu câu (ví dụ: `(1,0 điểm)`, `(0,5 điểm)`).
- **Bản Học sinh:** Để khoảng trống gồm các dòng chấm `............................................................` để học sinh làm bài.
- **Bản Giáo viên:** Lời giải từng bước chi tiết, viết đầy đủ phương trình hóa học cân bằng có trạng thái chất/nhiệt độ/xúc tác, tính số mol, đặt ẩn và kết luận đáp số.

### 4.5. QUY CÁCH TRÌNH BÀY LỜI GIẢI CHI TIẾT (BẢN GIÁO VIÊN)
- **TUYỆT ĐỐI KHÔNG DÙNG TABLE (BẢNG).**
- Tiêu đề lời giải: `► Lời giải chi tiết:`
  - In đậm, cỡ `10,5 pt`, màu xanh lam đậm `#1565C0`.
  - Thụt lề trái nhẹ `Left Indent = 0,4 cm`. Space Before = `3 pt`, Space After = `1 pt`.
- Nội dung các dòng lời giải:
  - Cỡ `10 pt`, thụt lề trái `Left Indent = 0,8 cm`.
  - Space Before = `1 pt`, Space After = `1 pt`.
  - Phân tích rõ từng đáp án đúng/sai, lập luận phương trình hóa học và công thức tính toán.

---

## 5. XỬ LÝ CHUẨN XÁC CÔNG THỨC HÓA HỌC & HÌNH ẢNH

1. **Chỉ số dưới (Subscript):**
   - Áp dụng triệt để cho toàn bộ công thức hóa học: H~2~O, CO~2~, N~2~, H~2~SO~4~, HNO~3~, Fe~2~O~3~, Al~2~(SO~4~)~3~, Ba(OH)~2~, NH~4~Cl, C~6~H~12~O~6~, C~2~H~5~OH, v.v.
   - Nồng độ hằng số cân bằng: K~C~.
   - Thể tích: V~dd~.
2. **Chỉ số trên (Superscript):**
   - Áp dụng cho ion, điện tích, lũy thừa, độ: H^+^, OH^-^, Fe^3+^, Al^3+^, SO~4~^2-^, NO~3~^-^, CO~3~^2-^, 10^-12^, 10^-5^, t^o^C.
3. **Mũi tên & ký hiệu phản ứng:**
   - Phản ứng thuận nghịch: dùng dấu hai chiều `⇌` (`\rightleftharpoons`).
   - Phản ứng một chiều: dùng `→`.
   - Kết tủa: `↓`, bay hơi: `↑`.
   - Biến thiên enthalpy: $\Delta_r H_{298}^o$ hoặc $\Delta H < 0$.
4. **Xử lý hình vẽ / Đồ thị:**
   - Trích xuất hoặc vẽ lại sắc nét, lưu vào `San_Pham/images_crop...`.
   - Chèn vào Word căn giữa trang (`WD_ALIGN_PARAGRAPH.CENTER`), kích thước chiều rộng từ `6,0 cm` đến `12,0 cm` tùy độ chi tiết.
   - Chú thích hình: Chữ in nghiêng/đậm nhẹ bên dưới hình ảnh.

---

## 6. QUY TRÌNH KHI TIẾP NHẬN ẢNH ĐỀ THI TỪ NGƯỜI DÙNG

Khi người dùng gửi ảnh đề thi, AI sẽ tự động thực hiện quy trình chuẩn:
1. **Trích xuất nội dung từ ảnh (OCR & Hóa học):** Đọc kỹ từng câu hỏi, công thức, bảng biểu, hình vẽ.
2. **Phân loại phần thi:** Nhận diện Phần I (MCQ 4 lựa chọn), Phần II (Đúng/Sai), Phần III (Trả lời ngắn), Phần IV (Tự luận nếu có).
3. **Biên tập lời giải chi tiết:** Giải chi tiết từng câu, lập luận khoa học, tính toán chuẩn xác từng bước.
4. **Xuất bản 2 file Word (.docx) chuẩn "Form Đề Kiểm Tra":**
   - `De_Kiem_Tra_Hoa_[Lop]_[TenDe].docx` (Bản Học sinh).
   - `De_Kiem_Tra_Hoa_[Lop]_[TenDe]_GiaoVien.docx` (Bản Giáo viên có lời giải chi tiết).
5. **(Nếu yêu cầu) Xuất bản LaTeX & PDF:** Biên dịch song song ra 2 bản PDF tương ứng.
6. **Đồng bộ tự động lên GitHub:** `git add`, `git commit` và `git push` về repository chính.
