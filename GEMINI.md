# QUY CHUẨN SOẠN THẢO VÀ XỬ LÝ ĐỀ KIỂM TRA HÓA HỌC ("FORM ĐỀ KIỂM TRA")

Tài liệu này là quy tắc bắt buộc (Project Rules) dành cho AI khi xử lý bất kỳ đề kiểm tra Hóa học nào trong workspace này, được định danh chính thức là **"form đề kiểm tra"** (chi tiết xem tại `FORM_DE_KIEM_TRA.md` và module `scripts/form_de_kiem_tra.py`). Trong mọi phiên làm việc tiếp theo, đặc biệt khi người dùng gửi ảnh đề thi, AI phải tự động áp dụng đầy đủ form chuẩn này mà không cần người dùng nhắc lại.

---

## 1. QUY CÁCH ĐỊNH DẠNG FILE WORD (.docx)

### 1.1. Trang & Lề (Page Setup)
- **Khổ giấy:** A4 (21,0 cm × 29,7 cm).
- **Căn lề (Margins):** Lề trên (Top) = 1,5 cm; Lề dưới (Bottom) = 1,5 cm; Lề trái (Left) = 1,5 cm; Lề phải (Right) = 1,5 cm.
- **Header & Footer cách mép giấy:** 0,6 cm.

### 1.2. Header (Đầu trang)
- Tạo bảng 1 hàng 2 cột, chiều rộng bảng 18,0 cm (khớp lề), căn giữa:
  - **Cột trái:** `Tài liệu lưu hành nội bộ` (Chữ in nghiêng, cỡ 10 pt, màu đỏ san hô Coral Red: `#D32F2F`).
  - **Cột phải:** `Thầy TRẦN VĂN THẠNH - 0777.470.803` (Cỡ 10 pt, màu đỏ san hô `#D32F2F`, tên `TRẦN VĂN THẠNH` in đậm).
- **Đường kẻ dưới (Border Bottom):** Màu xanh Sage Green (`#8BC390`), độ dày 0,75 pt (sz="6" hoặc sz="8" trong XML docx).

### 1.3. Footer (Chân trang)
- Tạo bảng 1 hàng 2 cột, chiều rộng bảng 18,0 cm, căn giữa:
  - **Cột trái:** `CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế   |   CS2: 24 Đặng Thái Thân, TP Huế` (Chữ in nghiêng, cỡ 9,5 pt, màu đỏ san hô `#D32F2F`).
  - **Cột phải:** `Trang X` (Sử dụng trường số trang tự động `PAGE`, chữ in đậm, cỡ 10 pt, màu đỏ san hô `#D32F2F`).
- **Đường kẻ trên (Border Top):** Màu xanh Sage Green (`#8BC390`), độ dày 0,75 pt.

### 1.4. Tiêu đề đề thi (Header Block)
- Bố trí thành bảng 2 cột không viền, gọn gàng, cân đối:
  - **Cột trái:** Tên Sở / Trường, `ĐỀ CHÍNH THỨC`, số trang.
  - **Cột phải:** `ĐỀ KIỂM TRA GIỮA / CUỐI HỌC KÌ...`, `MÔN: HOÁ HỌC - LỚP...`, thời gian làm bài, mã đề.
- Dòng thông tin học sinh: `Họ, tên thí sinh: ... Lớp: ... Số báo danh: ...` gom gọn 1 dòng.
- Dòng nguyên tử khối cho trong dấu ngoặc đơn, chữ nghiêng.

### 1.5. Công thức hóa học (Chỉ số trên/dưới)
- Bắt buộc xử lý triệt để chỉ số dưới (Subscript) và chỉ số trên (Superscript) cho toàn bộ các công thức hóa học:
  - Chỉ số dưới: H~2~O, CO~2~, C~6~H~12~O~6~, H~2~SO~4~, C~17~H~35~COONa, v.v.
  - Chỉ số trên: Fe^3+^, OH^-^, SO~4~^2-^, Cu^2+^, v.v.

### 1.6. Hai phiên bản tài liệu
Mỗi đề thi luôn tạo ra 2 file Word:
1. **Bản Học sinh** (`De_Kiem_Tra_Hoa_[Lop]_Ma[Made].docx`): Đề thi chuẩn, không hiển thị đáp án, có dòng chấm làm bài cho phần tự luận.
2. **Bản Giáo viên** (`De_Kiem_Tra_Hoa_[Lop]_Ma[Made]_GiaoVien.docx`):
   - Có dòng ghi chú đỏ `(BẢN GIÁO VIÊN -- CÓ LỜI GIẢI CHI TIẾT)`.
   - Đáp án đúng được in đậm màu đỏ/được đánh dấu rõ ràng.
   - Mỗi câu đều có phần lời giải chi tiết trình bày rõ ràng, nổi bật.

### 1.7. Tuyệt đối không dùng dạng bảng (Table) trong nội dung đề thi
- Trong toàn bộ phần nội dung đề thi (phương án lựa chọn A-B-C-D, các ý đúng/sai, và lời giải chi tiết): **Tuyệt đối KHÔNG sử dụng Table (bảng)**.
- Bố trí các phương án A, B, C, D bằng đoạn văn bản thuần túy (Paragraph) kết hợp điểm dừng Tab (Tab stops: 4.5 cm, 9.0 cm, 13.5 cm).
- Phần lời giải chi tiết trình bày dưới dạng đoạn văn bản thụt lề rõ ràng (Left Indent 0.5 - 0.8 cm), có tiêu đề in đậm màu xanh (`#1565C0`), không lồng vào ô bảng để thuận tiện copy và chỉnh sửa văn bản Word.

---

## 2. QUY CÁCH TÀI LIỆU LATEX & PDF (.tex, .pdf)

- **Gói & Tiền đề (Preamble):**
  - Khổ giấy A4, `\geometry{top=1.5cm,bottom=1.2cm,left=1.4cm,right=1.4cm,headheight=28pt,headsep=12pt,footskip=18pt}`.
  - Dùng `fancyhdr` với màu Header/Footer khớp Word: `hfRed` (`#D32F2F`) và `hfGreen` (`#8BC390`).
  - Gói hỗ trợ đề thi: `ex_test.sty`.
  - Công thức hóa học: `\usepackage[version=4]{mhchem}` (`\ce{...}`).
- **Hình ảnh:** Trích xuất, cắt nét rõ các hình vẽ cấu tạo/thí nghiệm và lưu vào `San_Pham/images_crop...`.
- **Biên dịch:** Biên dịch bằng `pdflatex` ra 2 file PDF tương ứng (`HocSinh.pdf` và `GiaoVien.pdf`).
- **Dọn dẹp:** Sau khi biên dịch, xóa toàn bộ file rác (`*.aux`, `*.log`, `*.thm`, `ans-*.tex`).

---

## 3. THƯ MỤC LƯU TRỮ & ĐỒNG BỘ GITHUB

1. **Thư mục sản phẩm:** Tất cả các file thành phẩm (`.docx`, `.tex`, `.pdf`) phải được lưu tại:
   `c:\Antigravity_Thanh\San_Pham\`
2. **Đồng bộ GitHub:** Sau khi hoàn thành bất kỳ đề nào:
   - Tự động kiểm tra `git status`.
   - `git add` tất cả sản phẩm mới và script tạo file.
   - `git commit` với thông điệp rõ ràng mô tả đề thi vừa làm.
   - `git push origin main` lên repository: `https://github.com/Thanhchem/De_Kiem_Tra_Hoa.git`.

---

## 4. THƯ VIỆN HÌNH ẢNH & MÃ VẼ HÓA HỌC (TỰ ĐỘNG TÁI SỬ DỤNG)

Khi xử lý đề thi hoặc câu hỏi có liên quan đến các cấu trúc dưới đây, AI **BẮT BUỘC TỰ ĐỘNG TÁI SỬ DỤNG** mã `chemfig` chuẩn và file ảnh độ nét cao đã dựng sẵn (chi tiết tra cứu tại `KHO_HINH_ANH_HOA_HOC.md`):

### 4.1. Sơ đồ điều chế Aniline từ Benzen
- **Đường dẫn ảnh:**
  - Bản đầy đủ chuẩn (300 DPI): `San_Pham/So_Do_Dieu_Che_Aniline/so_do_dieu_che_aniline.png`
  - Bản nền trong suốt (chèn Word không nền): `San_Pham/So_Do_Dieu_Che_Aniline/so_do_dieu_che_aniline_transparent.png`
  - Bản cực nét (600 DPI): `San_Pham/So_Do_Dieu_Che_Aniline/so_do_dieu_che_aniline_hd.png`
  - Bản vector: `San_Pham/So_Do_Dieu_Che_Aniline/so_do_dieu_che_aniline.pdf`
- **Mã LaTeX `chemfig`:**
  ```latex
  \schemestart
  \chemfig{*6(-=-(-[:30,0.75,,,draw=none]\phantom{NO_2})=-=)}
  \arrow{->[][(1)]}[,1.6]
  \chemfig{*6(-=-(-[:30,0.75]NO_2)=-=)}
  \arrow{->[Zn + HCl][(2)]}[,1.8]
  \chemfig{*6(-=-(-[:30,0.75]NH_3Cl)=-=)}
  \arrow{->[NaOH][(3)]}[,1.8]
  \chemfig{*6(-=-(-[:30,0.75]NH_2)=-=)}
  \schemestop
  ```

### 4.2. Bộ công thức ion lưỡng cực (Zwitterion) Amino Acid (Alanine)
- **Đường dẫn ảnh:**
  - Bộ 4 đáp án hàng ngang: `San_Pham/Cong_Thuc_Hoa_Hoc/de_ABCD_hang_ngang.png`
  - Bộ 4 đáp án lưới 2x2: `San_Pham/Cong_Thuc_Hoa_Hoc/de_ABCD_luoi_2x2.png`
  - Từng công thức đơn lẻ (nền trắng & transparent): `San_Pham/Cong_Thuc_Hoa_Hoc/ct_A.png`, `ct_B.png`, `ct_C.png`, `ct_D.png`.
- **Mã LaTeX `chemfig` từng chất:**
  - **A (Zwitterion $\beta$-alanine):** `\chemfig{^{+}H_3N-[:30]-[:-30]-[:30]COO^{-}}`
  - **B (Zwitterion Alanine chuẩn Hóa học):** `\chemfig{-[:-30](-[:-90]NH_3^{+})-[:30]COO^{-}}`
  - **B (Dạng in thiếu số 3 theo ảnh gốc):** `\chemfig{-[:-30](-[:-90]NH^{+})-[:30]COO^{-}}`
  - **C (Zwitterion axit 2,3-diaminopropanoic):** `\chemfig{^{+}H_3N-[:30](-[:90]NH_2)-[:-30]-[:30]COO^{-}}`
  - **D (Aminomethanol):** `\chemfig{H_2N-[:30]-[:-30]OH}`

### 4.3. Sơ đồ thí nghiệm bình tam giác nối bóng bay ($\text{CH}_3\text{COOH} + \text{NaHCO}_3 \to \text{CO}_2$)
- **Đường dẫn ảnh:**
  - Bản tiêu chuẩn (300 DPI): `San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay.png`
  - Bản trong suốt (Transparent): `San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay_transparent.png`
  - Bản cực nét (600 DPI): `San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay_hd.png`
  - Bản vector: `San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay.pdf`
  - Ảnh tài nguyên chung: `San_Pham/images_khtn8/khtn8_flask_balloon.png`
- **Mã nguồn vẽ (TikZ):** Xem tại `San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay.tex`.

