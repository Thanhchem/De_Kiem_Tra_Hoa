# -*- coding: utf-8 -*-
import os

path = r"c:\Antigravity_Thanh\LICH_SU_CONG_VIEC.md"
with open(path, "r", encoding="utf-8") as f:
    content = f.read()

old_table_end = '| 50 | *2 ảnh chụp đề thi/đề cương KHTN 9 (Sinh học - Mendel, DNA, RNA - Đúng/Sai, Trả lời ngắn, Tự luận)* | • [Word Học sinh](file:///c:/Antigravity_Thanh/San_Pham/De_Kiem_Tra_KHTN_9_SinhHoc_HocSinh.docx)<br>• [Word Dạy Trên Màn Hình](file:///c:/Antigravity_Thanh/San_Pham/De_Kiem_Tra_KHTN_9_SinhHoc_DayTrenManHinh.docx) | • [Word Giáo viên (Đầy đủ đáp án & Lời giải chi tiết)](file:///c:/Antigravity_Thanh/San_Pham/De_Kiem_Tra_KHTN_9_SinhHoc_GiaoVien.docx) | Hoàn thành (Chuẩn Form Đề Kiểm Tra -- 3 phiên bản Word -- Đã sync OneDrive) |'

new_table_row = '| 51 | [`Khtn7-decuong-gk1-kimlong-2627.pdf`](file:///c:/Antigravity_Thanh/PDF_Goc/Khtn7-decuong-gk1-kimlong-2627.pdf)<br>*(Đề cương Ôn tập Giữa HK1 - Môn KHTN 7 - THCS Kim Long 2026-2027)* | • [Word Học sinh](file:///c:/Antigravity_Thanh/San_Pham/De_Cuong_KHTN_7_Giua_HK1_Kim_Long_HocSinh.docx)<br>• [Word Dạy Trên Màn Hình](file:///c:/Antigravity_Thanh/San_Pham/De_Cuong_KHTN_7_Giua_HK1_Kim_Long_DayTrenManHinh.docx) | • [Word Giáo viên (Đầy đủ đáp án & Lời giải chi tiết)](file:///c:/Antigravity_Thanh/San_Pham/De_Cuong_KHTN_7_Giua_HK1_Kim_Long_GiaoVien.docx) | Hoàn thành (Chuẩn Form Đề Kiểm Tra -- 3 phiên bản Word -- Đã sync OneDrive) |'

if old_table_end in content and new_table_row not in content:
    content = content.replace(old_table_end, old_table_end + "\n" + new_table_row)

dot49_text = r"""
---

### Đợt 49: Hoàn thành Chuyển đổi và Soạn thảo Trọn bộ Đề Cương Ôn Tập Giữa Kì I - KHTN 7 Trường THCS Kim Long (2026 - 2027)

- **Thời gian thực hiện:** 19/10/2026
- **Tài liệu nguồn:** `Khtn7-decuong-gk1-kimlong-2627.pdf` (10 trang đề cương ảnh scan trường THCS Kim Long).
- **Phạm vi kiến thức:**
  - **BÀI 1: PHƯƠNG PHÁP VÀ KĨ NĂNG HỌC TẬP MÔN KHTN:**
    * 10 câu hỏi bao gồm trắc nghiệm 4 lựa chọn, câu hỏi nối cột A - Cột B (Câu 7), bài toán bình chia độ đo thể tích vật rắn không thấm nước (Câu 8 kèm Hình 1.1), cổng quang điện (Câu 9), cấu trúc báo cáo thực hành (Câu 10).
  - **PHÂN MÔN KHTN 1 (HÓA HỌC):**
    * I. Nội dung ôn tập lý thuyết: Bài 2 (Nguyên tử) & Bài 3 (Nguyên tố hóa học).
    * II. Trắc nghiệm:
      + Phần 1 (16 câu nhiều lựa chọn): Cấu tạo nguyên tử, mô hình Rutherford - Bohr, hạt p, n, e, lớp vỏ, nguyên tố hóa học, bảng tra cứu proton (Câu 16).
      + Phần 2 (4 câu Đúng/Sai, mỗi câu 4 ý a, b, c, d = 16 lệnh hỏi): Nhận định về nguyên tử, vỏ electron, bài toán nguyên tử X (13e, 14n), đặc trưng nguyên tố hóa học.
      + Phần 3 (5 câu trả lời ngắn): Số loại hạt trong hạt nhân, tìm neutron của Fluorine (m=19, e=9), electron của Carbon, tổng số hạt của Sodium (11p, 12n), số electron của Potassium (4 lớp, lớp ngoài cùng 1e).
    * III. Tự luận: 4 câu tự luận chuyên sâu: Điền từ vào chỗ trống, định nghĩa nguyên tố hóa học, vẽ sơ đồ cấu tạo nguyên tử Sulfur Z=16 (Câu 3), bài toán tổng số hạt và vẽ sơ đồ nguyên tử Aluminium Z=13 (Câu 4).
  - **PHÂN MÔN KHTN 2 (VẬT LÍ):**
    * A. Hệ thống kiến thức trọng tâm: Tốc độ chuyển động, đơn vị tốc độ (m/s, km/h), các phương pháp đo tốc độ (bấm giây, cổng quang điện, thiết bị bắn tốc độ), đồ thị quãng đường – thời gian.
    * B. Câu hỏi ôn tập:
      + Phần I: Trắc nghiệm nhiều lựa chọn (10 câu): Ý nghĩa tốc độ, công thức v=s/t, đơn vị SI, so sánh tốc độ bạn A và B, tốc độ trung bình ô tô, độ dốc đồ thị.
      + Phần 2: Trắc nghiệm Đúng - Sai (2 câu, 8 ý): Tính chất tốc độ, đồ thị s - t.
      + Phần 3: Trắc nghiệm trả lời ngắn (12 câu): Kí hiệu, đổi đơn vị 1 km/h <-> 1 m/s, công thức, bài toán tính tốc độ xe đạp, vận động viên, tốc kế.
      + Phần II: Tự luận (4 câu): Tính tốc độ xe tải trên 2 đoạn đường (Câu 25), tốc độ đạp xe và ý nghĩa (Câu 26), vẽ đồ thị s - t từ bảng số liệu (Câu 27 kèm đồ thị vector), bài toán camera bắn tốc độ 10m trong 0,5s (Câu 28).
  - **PHÂN MÔN KHTN 3 (SINH HỌC):**
    * Phần A: Hệ thống kiến thức trọng tâm (Khái quát trao đổi chất & chuyển hóa năng lượng, Quang hợp ở thực vật, Hô hấp tế bào, Thực hành).
    * Phần B: Câu hỏi ôn tập và gợi ý trả lời (23 câu hỏi phân hóa từ Nhận biết, Thông hiểu đến Vận dụng cao):
      + I. Mức độ nhận biết (Câu 1 đến Câu 9): Khái niệm, vai trò TĐC & NL, khái niệm và phương trình quang hợp, hô hấp tế bào, bào quan lục lạp, ti thể, cấu tạo lá, khí khổng, các yếu tố ảnh hưởng.
      + II. Mức độ thông hiểu (Câu 10 đến Câu 16): Phân tích mối quan hệ TĐC & NL, quang hợp ở lá đỏ/vàng, mối liên hệ giữa quang hợp và hô hấp, trồng xen cây ưa sáng/ưa bóng, phản ứng cơ thể khi vận động mạnh, không để hoa trong phòng ngủ đóng kín ban đêm, nuôi cá cảnh thả rong thủy sinh (Câu 16 được bổ sung lời giải chi tiết sư phạm).
      + III. Mức độ vận dụng và vận dụng cao (Câu 17 đến Câu 23): Trồng cây xanh ở đô thị/nhà máy, phản biện nhận định ban đêm cây không hô hấp, ngâm ủ hạt giống nước ấm, đứng dưới bóng cây mát hơn mái tôn, ngập úng ảnh hưởng rễ và quang hợp, biến đổi khí hậu và biện pháp khắc phục, bảo quản rau củ trong tủ lạnh.
      + Lưu ý khi ôn tập (3 quy tắc).
- **Quy cách kỹ thuật:**
  - Chuẩn mực tuyệt đối theo **FORM ĐỀ KIỂM TRA** (`GEMINI.md`): Khổ A4, lề 1.5 cm; Header & Footer màu đỏ san hô Coral Red (`#D32F2F`) & viền Xanh Sage Green (`#8BC390`), thông tin Thầy TRẦN VĂN THẠNH - 0777.470.803, 2 cơ sở tại Huế, số trang tự động `PAGE`.
  - **Tuyệt đối không dùng bảng (Table) trong nội dung đề thi**: Sử dụng Tab stops chuẩn Word; Lời giải chi tiết thụt lề 0.6cm với tiêu đề in đậm màu xanh lam `#1565C0`.
  - Hình ảnh thí nghiệm, sơ đồ cấu tạo nguyên tử và đồ thị s - t được vẽ mới hoàn toàn bằng TikZ chất lượng vector siêu nét:
    + `hinh_1_1.png`: Bình chia độ đo thể tích vật rắn (cắt và lọc nhiễu nền trắng sạch).
    + `so_do_sulfur.png`: Sơ đồ nguyên tử Sulfur (16p: 2, 8, 6).
    + `so_do_aluminium.png`: Sơ đồ nguyên tử Aluminium (13p: 2, 8, 3).
    + `do_thi_quang_duong_thoi_gian.png`: Đồ thị s - t (t: 0 - 2,0 h; s: 0 - 10,0 km).
- **Sản phẩm xuất bản:**
  1. `San_Pham/De_Cuong_KHTN_7_Giua_HK1_Kim_Long_HocSinh.docx`: Bản Học sinh chuẩn mực, có dòng kẻ chấm làm bài cho phần tự luận và điền đáp án ngắn.
  2. `San_Pham/De_Cuong_KHTN_7_Giua_HK1_Kim_Long_GiaoVien.docx`: Bản Giáo viên đầy đủ đáp án in đậm màu đỏ, lời giải chi tiết sư phạm từng câu hỏi, bảng tổng hợp đáp án ở cuối tài liệu.
  3. `San_Pham/De_Cuong_KHTN_7_Giua_HK1_Kim_Long_DayTrenManHinh.docx`: Bản Dạy trên màn hình có khung chừa chỗ trống tương tác màu tím nổi bật phục vụ giảng dạy máy chiếu / màn hình cảm ứng / tivi thông minh.
- **Đồng bộ tự động:**
  - Lưu trữ tại `c:\Antigravity_Thanh\San_Pham\`
  - Sao chép trực tiếp vào thư mục OneDrive của Thầy: `c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn7\KHTN7-IN-DEKIEMTRADINHKI\New folder (2)\`, `c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn7\KHTN7-IN-DEKIEMTRADINHKI\` và `c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn7\`
  - Lưu file gốc vào `PDF_Goc/Khtn7-decuong-gk1-kimlong-2627.pdf`
"""

if "Đợt 49:" not in content:
    content += dot49_text

with open(path, "w", encoding="utf-8") as f:
    f.write(content)

print("Updated LICH_SU_CONG_VIEC.md successfully!")
