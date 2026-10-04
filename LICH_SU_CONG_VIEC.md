# NHẬT KÝ LỊCH SỬ XỬ LÝ ĐỀ THI (PDF ➔ TEX ➔ WORD)

*Cập nhật lần cuối: 2026-10-04 10:50:00*

Hệ thống tự động theo dõi, chuyển đổi và quản lý toàn bộ các tệp đề thi từ thư mục `PDF_Goc/` sang định dạng LaTeX chuẩn `ex_test` và Microsoft Word (`.docx`).

---

## 1. BẢNG THEO DÕI TIẾN ĐỘ VÀ SẢN PHẨM HOÀN THÀNH

| STT | File PDF Gốc | Đề Học Sinh (PDF / Word) | Đề Giáo Viên Có Lời Giải (PDF / Word) | Trạng Thái |
|:---:|:---|:---|:---|:---:|
| 1 | [`De_Kiem_Tra_Hoa_11_Ma209.pdf`](file:///c:/Antigravity_Thanh/PDF_Goc/De_Kiem_Tra_Hoa_11_Ma209.pdf) | • [PDF Học sinh (3 trang A4)](file:///c:/Antigravity_Thanh/San_Pham/De_Kiem_Tra_Hoa_11_Ma209_HocSinh.pdf)<br>• [TeX Học sinh](file:///c:/Antigravity_Thanh/San_Pham/De_Kiem_Tra_Hoa_11_Ma209_HocSinh.tex)<br>• [Word Học sinh](file:///c:/Antigravity_Thanh/San_Pham/De_Kiem_Tra_Hoa_11_Ma209.docx) | • [PDF Giáo viên (8 trang chi tiết)](file:///c:/Antigravity_Thanh/San_Pham/De_Kiem_Tra_Hoa_11_Ma209_GiaoVien.pdf)<br>• [TeX Giáo viên](file:///c:/Antigravity_Thanh/San_Pham/De_Kiem_Tra_Hoa_11_Ma209_GiaoVien.tex)<br>• [Word Giáo viên](file:///c:/Antigravity_Thanh/San_Pham/De_Kiem_Tra_Hoa_11_Ma209_GiaoVien.docx) | Hoàn thành |

---

## 2. QUY CHUẨN THIẾT KẾ ĐÃ ÁP DỤNG

1. **Header & Footer bản quyền (theo file mẫu `3C2-26-27-DE.docx`)**:
   - **Header**:
     - Bên trái: `Tài liệu lưu hành nội bộ | THPT Hai Bà Trưng -- Huế`
     - Bên phải: `Thầy TRẦN VĂN THẠNH — 0777.470.803`
     - Kẻ chỉ xanh đậm thanh lịch dưới Header.
   - **Footer**:
     - Dòng 1: `CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế | CS2: 24 Đặng Thái Thân, TP Huế`
     - Dòng 2: `Lớp Hóa Thầy Thạnh — SĐT: 0777.470.803`
     - Số trang tự động bên phải (`Trang X/3` hoặc `Trang X`).
     - Kẻ chỉ xanh đậm thanh lịch phía trên Footer.
2. **Quy chuẩn Microsoft Word (.docx)**:
   - Font chữ: `Times New Roman`, cỡ 12pt, giãn dòng 1.15, giãn đoạn sau 3pt.
   - **Tuyệt đối không dùng bảng** để chia đáp án A, B, C, D: dùng Tab Stops chuẩn Word (0.5cm, 4.8cm, 9.2cm, 13.6cm) thẳng hàng tăm tắp.
   - Các chữ và từ cách nhau chuẩn, không bị dính sát nhau.
   - Hình ảnh công thức hóa học (dẫn xuất halogen, catechin, geraniol) được crop độ phân giải cao và chèn ảnh trực tiếp, không vẽ lại trong Word.
3. **Quy chuẩn LaTeX (`ex_test` + `mhchem` + `chemfig`)**:
   - Cấu trúc Catechin lập thể chính xác 100% bằng ChemFig.
   - Đề học sinh dàn trang chuẩn **chính xác 3 trang A4**.
   - Đề giáo viên giải chi tiết từng bước, có bảng tổng hợp đáp án trắc nghiệm ở cuối.

---

## 3. CẤU TRÚC THƯ MỤC LƯU TRỮ

- `PDF_Goc/`: Nơi Thầy copy các file PDF đề thi mới vào.
- `San_Pham/`: Nơi lưu trữ toàn bộ sản phẩm hoàn thiện (PDF học sinh, PDF giáo viên, Word học sinh, Word giáo viên, mã nguồn TeX, hình ảnh crop).
- `scripts/`: Thư mục chứa toàn bộ mã nguồn xử lý tự động, tạo PDF và tạo Word.
- `quan_ly_tien_trinh.py`: File script tự động quét và kiểm tra xem có file PDF mới nào cần xử lý hay không.
