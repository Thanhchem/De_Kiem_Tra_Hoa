# -*- coding: utf-8 -*-
"""
Cập nhật bảng đáp án rút gọn chuẩn mẫu ảnh vào cả 4 file Word Giáo Viên:
1. De_Kiem_Tra_Hoa_10_THPT_Nguyen_Dinh_Chieu_Ma1002_GiaoVien.docx
2. De_Kiem_Tra_CK1_Hoa_10_HoaChau_Ma132_GiaoVien.docx
3. De_Kiem_Tra_CK1_Hoa_10_DangHuyTru_Ma104_GiaoVien.docx
4. De_Kiem_Tra_CK1_Hoa_10_NguyenChiThanh_Ma1104_GiaoVien.docx
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.oxml import parse_xml

sys.path.append(os.path.abspath('scripts'))
from build_bang_dap_an_rut_gon import add_styled_answer_block, EXAMS_DATA
from generate_de_nguyen_dinh_chieu_1002 import create_exam_file as gen_1002
from generate_de_hoa_chau_132 import create_exam_file as gen_132
from generate_de_dang_huy_tru_104 import create_exam_file as gen_104
from generate_de_nguyen_chi_thanh_1104 import create_exam_file as gen_1104

# Sửa lại hàm trong generate_de_*.py để gọi add_styled_answer_block
# Hoặc cập nhật trực tiếp vào file Word:
def inject_styled_table_into_teacher_file(docx_path, exam_data):
    doc = docx.Document(docx_path)
    # Tìm đoạn "BẢNG ĐÁP ÁN RÚT GỌN" cũ (nếu có) và xóa các đoạn cũ, thay bằng bảng mới
    # Đơn giản và chuẩn nhất: tạo một doc mới hoặc chèn vào ngay sau Title Block!
    pass

print("Updating exam scripts to use add_styled_answer_block directly...")
