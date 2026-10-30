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

### 4.4. Sơ đồ thí nghiệm cờ lê và bu lông (Tác dụng làm quay của lực / Moment lực)
- **Đường dẫn ảnh:**
  - Bản tiêu chuẩn (300 DPI): `San_Pham/Thi_Nghiem_Co_Le_Bu_Long/thi_nghiem_co_le_bu_long.png`
  - Bản trong suốt (Transparent): `San_Pham/Thi_Nghiem_Co_Le_Bu_Long/thi_nghiem_co_le_bu_long_transparent.png`
  - Bản cực nét (600 DPI): `San_Pham/Thi_Nghiem_Co_Le_Bu_Long/thi_nghiem_co_le_bu_long_hd.png`
  - Bản vector: `San_Pham/Thi_Nghiem_Co_Le_Bu_Long/thi_nghiem_co_le_bu_long.pdf`
  - Ảnh tài nguyên chung: `San_Pham/images_khtn8/khtn8_wrenches.png`
- **Mã nguồn vẽ (TikZ):** Xem tại `San_Pham/Thi_Nghiem_Co_Le_Bu_Long/thi_nghiem_co_le_bu_long.tex`.

### 4.5. Sơ đồ phản ứng thủy phân Salicin và oxi hóa Salicylic alcohol
- **Đường dẫn ảnh:**
  - Bản tiêu chuẩn (300 DPI): `San_Pham/So_Do_Thuy_Phan_Salicin/so_do_thuy_phan_salicin.png`
  - Bản trong suốt (Transparent): `San_Pham/So_Do_Thuy_Phan_Salicin/so_do_thuy_phan_salicin_transparent.png`
  - Bản cực nét (600 DPI): `San_Pham/So_Do_Thuy_Phan_Salicin/so_do_thuy_phan_salicin_hd.png`
  - Bản vector: `San_Pham/So_Do_Thuy_Phan_Salicin/so_do_thuy_phan_salicin.pdf`
- **Mã nguồn vẽ (TikZ):** Xem tại `San_Pham/So_Do_Thuy_Phan_Salicin/so_do_thuy_phan_salicin.tex`.
- **Script sinh hình ảnh:** `scripts/ve_so_do_salicin.py`.

---

## 5. QUY CHUẨN FILE PDF TRÌNH CHIẾU TIVI 65 INCH (FORM CHIẾU TIVI 16:9)

Khi người dùng yêu cầu file PDF **trình chiếu Tivi**, AI bắt buộc áp dụng chính xác bộ quy chuẩn này:

### 5.1. Kích thước & Tỉ lệ (Khắc phục triệt để viền đen)
- **Tỉ lệ 16:9 Widescreen:** `paperwidth=29.7cm, paperheight=16.706cm`.
- **Căn lề:** `top=1.0cm, bottom=1.0cm, left=1.2cm, right=1.2cm`.
- **Loại bỏ Header/Footer:** `headheight=0pt, headsep=0pt, footskip=0pt`, `\pagestyle{empty}`.
- **Hiển thị:** Mở toàn màn hình (`Ctrl+L`) trên Tivi 65 inch ($1920\times 1080$) tràn viền $100\%$, không có bất kỳ dải viền đen nào.

### 5.2. Phong cách thiết kế
- **Tuyệt đối KHÔNG DÙNG KHUNG (No boxes / No borders / No tcolorbox).**
- **Nền trắng tinh (#FFFFFF), chữ đen tuyền (#000000), font Times New Roman.**
- **Không chèn tiêu đề rườm rà trên slide câu hỏi:** Để dành toàn bộ chiều cao cho câu hỏi và khoảng trống giảng bài.
- Nhãn câu hỏi in đậm: `\textbf{Câu X.}`.

### 5.3. Phân bổ slide & Cỡ chữ chuẩn
- **Phần I (Trắc nghiệm nhiều lựa chọn):**
  - **2 câu / 1 slide.**
  - Cỡ chữ: **30pt** (`\fontsize{30}{39}\selectfont`).
  - Dãn cách giữa 2 câu bằng `\vspace*{\fill}` để dàn đều toàn bộ chiều cao màn hình.
  - 4 phương án xếp 2 cột bằng `tabularx`, nhãn in đậm `\textbf{A.}`, khoảng cách hàng `\\[8pt]`.
- **Phần II (Trắc nghiệm Đúng / Sai):**
  - **1 câu / 1 slide.**
  - Cỡ chữ: **26pt** (`\fontsize{26}{34}\selectfont`). 4 ý a), b), c), d) rõ ràng kèm ô trống `[\quad]` bên phải.
- **Phần III (Trắc nghiệm Trả lời ngắn):**
  - **1 câu / 1 slide.**
  - Cỡ chữ: **32pt** (`\fontsize{32}{42}\selectfont`).
  - Dòng trả lời: `\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$`.
  - Phía dưới để trống $50\% - 60\%$ màn hình cho giáo viên tương tác/chữa bài trên Tivi.
- **Slide Bảng đáp án:** Đặt ở cuối mỗi đề thi, trình bày bảng đáp án chuẩn rõ ràng.

---

## 6. KHO NGÂN HÀNG CÂU HỎI & ĐỀ KIỂM TRA HÓA HỌC 12 (TỰ ĐỘNG TRA CỨU)

Khi người dùng nhắc đến **"kho đề 12"**, **"ngân hàng câu hỏi Hóa 12"**, hoặc yêu cầu rút đề kiểm tra Giữa kì / Cuối kì / Thi thử Hóa 12 theo ma trận, AI tự động truy cập tại thư mục:
- **Thư mục Kho Ngân hàng Hóa 12:** `c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\`
- **Cơ sở dữ liệu JSON trọn bộ 9 chương (1.362 câu hỏi đã phân loại Phần I/II/III và mức độ Biết `NB`, Hiểu `TH`, Vận dụng `VD`):**
  `c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Tron_Bo_9_Chuong.json`
- **18 file Word Ngân hàng từng chương (Chương 1 → Chương 8 & Chuyên đề học tập, gồm bản `_HocSinh.docx` và `_GiaoVien.docx`):**
  `c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\Ngan_Hang_Hoa_12_Chuong_*.docx`
- **Bộ 20 Đề Kiểm tra Giữa Học kì I – Hóa 12 (Đề 01 → Đề 20, Mã đề 101 → 120, 560 câu không trùng lặp):**
  - Lưu tại `c:\Antigravity_Thanh\San_Pham\De_Kiem_Tra_GK1_Hoa_12_De01..20_*.docx` và `c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_12\` (gồm đủ 3 bản cho mỗi đề: `_HocSinh.docx`, `_GiaoVien.docx` và `_Azota.docx`).
  - Script rút đề và xuất Azota: `c:\Antigravity_Thanh\scripts\generate_5_de_gk1_hoa12.py`, `c:\Antigravity_Thanh\scripts\generate_de11_to_de20_gk1_hoa12.py`, `c:\Antigravity_Thanh\scripts\generate_azota_gk1_hoa12.py`.
  - Kho hình ảnh trích xuất theo câu hỏi: `c:\Antigravity_Thanh\San_Pham\images_gk1_hoa12\`.

---

## 7. KHO NGÂN HÀNG CÂU HỎI HÓA HỌC 10 (TRỌN BỘ 7 CHƯƠNG - 1.545 CÂU HỎI)

Khi người dùng nhắc đến **"kho đề 10"**, **"ngân hàng câu hỏi Hóa 10"**, hoặc yêu cầu rút đề kiểm tra Giữa kì / Cuối kì / Ôn tập Hóa 10, AI tự động tra cứu tại thư mục:
- **Thư mục Kho Ngân hàng Hóa 10:** `c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_10\`
- **Cơ sở dữ liệu JSON trọn bộ 7 chương (1.545 câu hỏi, 100% đầy đủ đáp án & lời giải chi tiết):**
  `c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_10\Ngan_Hang_Hoa_10_Tron_Bo.json`
- **14 file Word Ngân hàng từng chương (Chương 1 → Chương 7, gồm bản `_HocSinh.docx` và `_GiaoVien.docx` chuẩn FORM ĐỀ KIỂM TRA):**
  - C1 (Cấu tạo nguyên tử): 287 câu
  - C2 (Bảng tuần hoàn & Định luật tuần hoàn): 221 câu
  - C3 (Liên kết hóa học): 261 câu
  - C4 (Phản ứng oxi hóa – khử): 158 câu
  - C5 (Năng lượng hóa học): 176 câu
  - C6 (Tốc độ phản ứng hóa học): 162 câu
  - C7 (Nguyên tố nhóm halogen): 280 câu
- **Kho hình ảnh sơ đồ/thí nghiệm trích xuất:** `c:\Antigravity_Thanh\San_Pham\images_hoa10\`

---

## 8. KHO NGÂN HÀNG CÂU HỎI HÓA HỌC 11 (TRỌN BỘ 7 CHƯƠNG - 1.996 CÂU HỎI)

Khi người dùng nhắc đến **"kho đề 11"**, **"ngân hàng câu hỏi Hóa 11"**, hoặc yêu cầu rút đề kiểm tra Giữa kì / Cuối kì / Ôn tập Hóa 11 theo ma trận, AI tự động tra cứu tại thư mục:
- **Thư mục Kho Ngân hàng Hóa 11:** `c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_11\`
- **Cơ sở dữ liệu JSON trọn bộ 7 chương (1.996 câu hỏi, 100% đầy đủ đáp án & lời giải chi tiết):**
  `c:\Antigravity_Thanh\San_Pham\Ngan_Hang_Hoa_11\Ngan_Hang_Hoa_11_Tron_Bo.json`
- **14 file Word Ngân hàng từng chương (Chương 1 → Chương 6 & Chuyên đề học tập, gồm bản `_HocSinh.docx` và `_GiaoVien.docx` chuẩn FORM ĐỀ KIỂM TRA):**
  - C1 (Cân bằng hóa học): 280 câu (P1: 171, P2: 44, P3: 65)
  - C2 (Nitrogen – Sulfur): 369 câu (P1: 229, P2: 40, P3: 100)
  - C3 (Đại cương hóa học hữu cơ): 292 câu (P1: 210, P2: 35, P3: 47)
  - C4 (Hydrocarbon): 388 câu (P1: 260, P2: 35, P3: 93)
  - C5 (Dẫn xuất halogen – Alcohol – Phenol): 355 câu (P1: 223, P2: 46, P3: 86)
  - C6 (Hợp chất carbonyl – Carboxylic acid): 281 câu (P1: 231, P2: 37, P3: 13)
  - CDHT (Chuyên đề học tập 11: Phân bón – Thực hành – Dầu mỏ): 31 câu (P1: 17, P2: 7, P3: 7)
- **Script sinh và cập nhật ngân hàng:** `c:\Antigravity_Thanh\scripts\build_ngan_hang_hoa11_full.py`.
- **Kho hình ảnh sơ đồ/thí nghiệm trích xuất:** `c:\Antigravity_Thanh\San_Pham\images_hoa11\`.






