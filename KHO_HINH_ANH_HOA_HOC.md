# KHO HÌNH ẢNH & THƯ VIỆN CÔNG THỨC HÓA HỌC CHUẨN (TÁI SỬ DỤNG)
**Biên soạn & Quản lý:** Thầy TRẦN VĂN THẠNH - 0777.470.803  
**Workspace:** `c:\Antigravity_Thanh`

Tài liệu này lưu trữ danh mục toàn bộ các hình ảnh cấu trúc Hóa học, sơ đồ phản ứng và mã LaTeX `chemfig` đã được vẽ lại với độ chính xác cao nhất (300 DPI - 600 DPI, vector PDF, nền trong suốt). Trong các phiên làm việc tiếp theo, AI sẽ tự động tái sử dụng các tài nguyên này bất cứ khi nào xuất hiện câu hỏi liên quan trong đề kiểm tra.

---

## 1. SƠ ĐỒ PHẢN ỨNG ĐIỀU CHẾ ANILINE TỪ BENZEN

- **Ý nghĩa Hóa học:** Chuỗi chuyển hóa 3 giai đoạn điều chế amine thơm (aniline):
  1. $\text{C}_6\text{H}_6 \xrightarrow{\text{HNO}_3 \text{ đ} / \text{H}_2\text{SO}_4 \text{ đ}, t^\circ} \text{C}_6\text{H}_5\text{NO}_2$ *(Phản ứng Nitro hóa benzen)*
  2. $\text{C}_6\text{H}_5\text{NO}_2 \xrightarrow{\text{Zn} + \text{HCl}, t^\circ} \text{C}_6\text{H}_5\text{NH}_3\text{Cl}$ *(Phản ứng khử hợp chất nitro thành muối amoni)*
  3. $\text{C}_6\text{H}_5\text{NH}_3\text{Cl} \xrightarrow{\text{NaOH}} \text{C}_6\text{H}_5\text{NH}_2 + \text{NaCl} + \text{H}_2\text{O}$ *(Tái tạo amine tự do ít tan tách lớp)*

### 1.1. Hình ảnh sơ đồ hoàn chỉnh
- **Bản tiêu chuẩn (300 DPI, nền trắng):** `San_Pham/So_Do_Dieu_Che_Aniline/so_do_dieu_che_aniline.png`
- **Bản trong suốt (Transparent, dùng chèn Word):** `San_Pham/So_Do_Dieu_Che_Aniline/so_do_dieu_che_aniline_transparent.png`
- **Bản Ultra HD (600 DPI in ấn):** `San_Pham/So_Do_Dieu_Che_Aniline/so_do_dieu_che_aniline_hd.png`
- **Bản Vector PDF (dùng chèn LaTeX):** `San_Pham/So_Do_Dieu_Che_Aniline/so_do_dieu_che_aniline.pdf`

### 1.2. Mã nguồn LaTeX ChemFig
```latex
\usepackage{chemfig}
\usepackage[scaled=0.95]{helvet}
\renewcommand{\familydefault}{\sfdefault}

\setchemfig{
    atom sep=2.0em,
    bond offset=1.8pt,
    bond style={line width=0.9pt, line cap=round, line join=round},
    arrow style={line width=0.9pt, -{Latex[length=4.5pt, width=3.2pt]}}
}
\renewcommand*\printatom[1]{\ensuremath{\mathsf{#1}}}

% Mã vẽ sơ đồ:
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

### 1.3. Hình ảnh các chất thành phần đơn lẻ
- **Benzen:** `San_Pham/So_Do_Dieu_Che_Aniline/01_benzen.png` | `01_benzen.pdf`
- **Nitrobenzen:** `San_Pham/So_Do_Dieu_Che_Aniline/02_nitrobenzen.png` | `02_nitrobenzen.pdf`
- **Phenylamoni clorua:** `San_Pham/So_Do_Dieu_Che_Aniline/03_phenylamoni_clorua.png` | `03_phenylamoni_clorua.pdf`
- **Aniline:** `San_Pham/So_Do_Dieu_Che_Aniline/04_aniline.png` | `04_aniline.pdf`

---

## 2. BỘ CÔNG THỨC ION LƯỠNG CỰC (ZWITTERION) AMINO ACID (ALANINE)

- **Ý nghĩa Hóa học:** Dạng tồn tại chủ yếu của amino acid trong dung dịch và ở pH sinh lý ($\approx 6$). Nhóm $-\text{COOH}$ nhường proton $\text{H}^+$ cho nhóm $-\text{NH}_2$ tạo thành ion lưỡng cực có điện tích tổng bằng $0$.
  - Alanine ($\alpha$-alanine) tạo zwitterion: $\text{CH}_3-\text{CH}(\text{NH}_3^+)-\text{COO}^-$

### 2.1. Hình ảnh tổng hợp cả câu hỏi 4 phương án
- **Hàng ngang (Chuẩn đề thi):** `San_Pham/Cong_Thuc_Hoa_Hoc/de_ABCD_hang_ngang.png` (có bản `_transparent.png` và `.pdf`)
- **Lưới 2x2 (Khớp ảnh gốc đề):** `San_Pham/Cong_Thuc_Hoa_Hoc/de_ABCD_luoi_2x2.png` (có bản `_transparent.png` và `.pdf`)

### 2.2. Chi tiết từng phương án
| Phương án | Tên gọi & Bản chất | File ảnh chèn Word | File PDF Vector | Mã LaTeX ChemFig |
| :---: | :--- | :--- | :--- | :--- |
| **A** | Zwitterion của $\beta$-alanine ($\text{H}_3\text{N}^+-\text{CH}_2-\text{CH}_2-\text{COO}^-$) | `San_Pham/Cong_Thuc_Hoa_Hoc/ct_A.png` | `ct_A.pdf` | `\chemfig{^{+}H_3N-[:30]-[:-30]-[:30]COO^{-}}` |
| **B** *(Chuẩn)* | Zwitterion của Alanine ($\text{CH}_3-\text{CH}(\text{NH}_3^+)-\text{COO}^-$) | `San_Pham/Cong_Thuc_Hoa_Hoc/ct_B.png` | `ct_B.pdf` | `\chemfig{-[:-30](-[:-90]NH_3^{+})-[:30]COO^{-}}` |
| **B** *(Ảnh đề)* | Zwitterion của Alanine (bản sao đề in thiếu số 3: $\text{NH}^+$) | `San_Pham/Cong_Thuc_Hoa_Hoc/ct_B_anh_goc.png` | `ct_B_anh_goc.pdf` | `\chemfig{-[:-30](-[:-90]NH^{+})-[:30]COO^{-}}` |
| **C** | Zwitterion của axit 2,3-diaminopropanoic | `San_Pham/Cong_Thuc_Hoa_Hoc/ct_C.png` | `ct_C.pdf` | `\chemfig{^{+}H_3N-[:30](-[:90]NH_2)-[:-30]-[:30]COO^{-}}` |
| **D** | Aminomethanol ($\text{H}_2\text{N}-\text{CH}_2-\text{OH}$) | `San_Pham/Cong_Thuc_Hoa_Hoc/ct_D.png` | `ct_D.pdf` | `\chemfig{H_2N-[:30]-[:-30]OH}` |

### 2.3. Mã nguồn LaTeX khi trình bày dạng câu trắc nghiệm
```latex
% Đặt 4 phương án trên 1 dòng chuẩn đề thi:
\noindent
\textbf{A.} $\vcenter{\hbox{\chemfig{^{+}H_3N-[:30]-[:-30]-[:30]COO^{-}}}}$ \qquad\quad
\textbf{B.} $\vcenter{\hbox{\chemfig{-[:-30](-[:-90]NH_3^{+})-[:30]COO^{-}}}}$ \qquad\quad
\textbf{C.} $\vcenter{\hbox{\chemfig{^{+}H_3N-[:30](-[:90]NH_2)-[:-30]-[:30]COO^{-}}}}$ \qquad\quad
\textbf{D.} $\vcenter{\hbox{\chemfig{H_2N-[:30]-[:-30]OH}}}$
```

---

## 3. SƠ ĐỒ THÍ NGHIỆM BÌNH TAM GIÁC NỐI BÓNG BAY ($\text{CH}_3\text{COOH} + \text{NaHCO}_3 \to \text{CO}_2$)

- **Ý nghĩa Hóa học / KHTN 8:** Thí nghiệm chứng minh phản ứng hóa học xảy ra có sự tạo thành chất khí (dấu hiệu nhận biết phản ứng) và định luật bảo toàn khối lượng (khi phản ứng diễn ra trong hệ kín):
  $$\text{CH}_3\text{COOH} + \text{NaHCO}_3 \to \text{CH}_3\text{COONa} + \text{CO}_2\uparrow + \text{H}_2\text{O}$$
  Khí $\text{CO}_2$ sinh ra làm quả bóng bay gắn ở miệng bình tam giác phồng to lên.

### 3.1. Hình ảnh thí nghiệm chuẩn độ nét cao
- **Bản tiêu chuẩn (300 DPI, nền trắng):** `San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay.png`
- **Bản trong suốt (Transparent, dùng chèn Word không nền):** `San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay_transparent.png`
- **Bản Ultra HD (600 DPI in ấn):** `San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay_hd.png`
- **Bản Vector PDF (dùng chèn LaTeX):** `San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay.pdf`
- **Thư mục ảnh tài nguyên chung:** `San_Pham/images_khtn8/khtn8_flask_balloon.png`

### 3.2. Mã nguồn LaTeX TikZ
Xem chi tiết mã nguồn đầy đủ tại: `San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay.tex`.

---

## 4. CÁC SCRIPTS TỰ ĐỘNG SINH HÌNH ẢNH
- `scripts/ve_cong_thuc_alanine.py`: Tự động biên dịch và trích xuất ảnh PNG/PDF cho các công thức amino acid.
- `scripts/ve_so_do_aniline.py`: Tự động biên dịch và trích xuất sơ đồ phản ứng điều chế aniline.
- `San_Pham/Thi_Nghiem_Binh_Tam_Giac_Bong_Bay/thi_nghiem_binh_tam_giac_bong_bay.tex`: Mã nguồn TikZ vẽ thí nghiệm bình tam giác gắn bóng bay.

