# -*- coding: utf-8 -*-
"""
Script biên soạn trọn bộ tài liệu chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH):
ĐỀ KIỂM TRA, ĐÁNH GIÁ GIỮA HỌC KÌ I - MÔN: KHOA HỌC TỰ NHIÊN 7
NĂM HỌC 2025 - 2026 - MÃ ĐỀ: 01 (Thời gian làm bài: 60 phút)

Sản phẩm tạo ra:
1. De_Kiem_Tra_GK1_KHTN7_Ma01_HocSinh.docx
2. De_Kiem_Tra_GK1_KHTN7_Ma01_GiaoVien.docx
3. De_Kiem_Tra_GK1_KHTN7_Ma01_HocSinh.pdf (và .tex)
4. De_Kiem_Tra_GK1_KHTN7_Ma01_GiaoVien.pdf (và .tex)
5. De_Kiem_Tra_GK1_KHTN7_Ma01_TrinhChieu_Tivi_16x9.pdf (và .tex)
"""

import os
import sys
import subprocess
import shutil
import re

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

import docx
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

sys.path.insert(0, r"c:\Antigravity_Thanh\scripts")
from form_de_kiem_tra import (
    setup_header_footer,
    build_title_block,
    add_section_header,
    render_question_prompt,
    render_options_no_table,
    render_statement,
    add_solution_no_table,
    add_student_writing_lines,
    render_rich_text,
    set_run_style,
    COLOR_PRIMARY,
    COLOR_CORRECT,
    COLOR_SOLUTION,
    COLOR_HF_RED,
    COLOR_HF_GREEN_HEX
)

OUTPUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

TARGET_ONEDRIVE_DIRS = [
    r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn7\KHTN7-IN-DEKIEMTRADINHKI",
    r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn7"
]

# ==============================================================================
# DỮ LIỆU ĐỀ THI ĐẦY ĐỦ VÀ CHÍNH XÁC
# ==============================================================================

PART1_QUESTIONS = [
    {
        "id": 1,
        "content": "Cho dãy các kí hiệu các nguyên tố sau: He, Ca, K, Fe, Cu, S. Theo thứ tự tên của các nguyên tố lần lượt là:",
        "options": [
            "A. Helium, Caesium, Potassium, Copper, Iron, Silicon.",
            "*B. Helium, Calcium, Potassium, Iron, Copper, Sulfur.*",
            "C. Helium, Calcium, Krypton, Iron, Copper, Silicon.",
            "D. Hydogen, Caesium, Krypton, Iron, Cobalt, Sulfur."
        ],
        "answer": "B",
        "solution": (
            "Kí hiệu hoá học và tên gọi tương ứng:\n"
            "- He: Helium\n"
            "- Ca: Calcium\n"
            "- K: Potassium\n"
            "- Fe: Iron (sắt)\n"
            "- Cu: Copper (đồng)\n"
            "- S: Sulfur (lưu huỳnh)\n"
            "Vậy thứ tự đúng là: Helium, Calcium, Potassium, Iron, Copper, Sulfur.\n"
            "➔ **Chọn B.**"
        )
    },
    {
        "id": 2,
        "content": "Các nguyên tố trong cùng 1 chu kì thì có cùng:",
        "options": [
            "A. Số electron lớp ngoài cùng.",
            "B. Khối lượng nguyên tử",
            "C. Số điện tích hạt nhân.",
            "*D. Số lớp electron.*"
        ],
        "answer": "D",
        "solution": (
            "Chu kì là dãy các nguyên tố mà nguyên tử của chúng có cùng số lớp electron và được xếp theo chiều điện tích hạt nhân tăng dần.\n"
            "Do đó, các nguyên tố trong cùng 1 chu kì có cùng số lớp electron.\n"
            "➔ **Chọn D.**"
        )
    },
    {
        "id": 3,
        "content": "Đơn chất là những chất được tạo nên từ bao nhiêu nguyên tố hóa học?",
        "options": [
            "A. Từ 2 nguyên tố.",
            "B. Từ 3 nguyên tố.",
            "C. Từ 4 nguyên tố trở lên.",
            "*D. Từ 1 nguyên tố.*"
        ],
        "answer": "D",
        "solution": (
            "Đơn chất là chất được tạo nên từ một nguyên tố hóa học (ví dụ: O~2~, N~2~, Fe, Cu, Al...).\n"
            "Chất tạo nên từ hai hay nhiều nguyên tố hóa học được gọi là hợp chất.\n"
            "➔ **Chọn D.**"
        )
    },
    {
        "id": 4,
        "content": "Hạt nhân một nguyên tử fluorine có 9 proton và 10 neutron. Khối lượng của nguyên tử fluorine xấp xỉ bằng:",
        "options": [
            "A. 9 amu",
            "B. 10 amu.",
            "*C. 19 amu*",
            "D. 28 amu."
        ],
        "answer": "C",
        "solution": (
            "Khối lượng của electron rất nhỏ so với proton và neutron nên khối lượng nguyên tử xấp xỉ bằng khối lượng hạt nhân:\n"
            "m ≈ m_p + m_n = 9 × 1 amu + 10 × 1 amu = 19 amu.\n"
            "➔ **Chọn C.**"
        )
    },
    {
        "id": 5,
        "content": "Một bình khí oxygen chứa:",
        "options": [
            "*A. Các phân tử O~2~.*",
            "B. Các nguyên tử oxygen riêng rẽ không liên kết với nhau.",
            "C. Một phân tử O~2~",
            "D. Một đại phân tử khổng lồ chứa rất nhiều nguyên tử oxygen."
        ],
        "answer": "A",
        "solution": (
            "Khí oxygen là đơn chất gồm các phân tử O~2~ (mỗi phân tử do 2 nguyên tử O liên kết với nhau). Trong một bình khí oxygen chứa vô số các phân tử O~2~.\n"
            "➔ **Chọn A.**"
        )
    },
    {
        "id": 6,
        "content": "Loại hạt nào sau đây đặc trưng cho 1 nguyên tố hóa học?",
        "options": [
            "A. hạt electron",
            "*B. hạt proton*",
            "C. Hạt electron và proton",
            "D. hạt electron và neutron"
        ],
        "answer": "B",
        "solution": (
            "Số hạt proton trong hạt nhân (hay số đơn vị điện tích hạt nhân) là đại lượng đặc trưng cho một nguyên tố hóa học. Các nguyên tử có cùng số proton đều thuộc cùng một nguyên tố hóa học.\n"
            "➔ **Chọn B.**"
        )
    },
    {
        "id": 7,
        "content": "Một nguyên tử carbon có khối lượng là 1,9926 × 10^-23^ gam. Khối lượng nguyên tử carbon tính theo đơn vị amu là: (biết 1 amu = 1,6605 × 10^-24^ gam).",
        "options": [
            "*A. 12 amu.*",
            "B. 24 amu.",
            "C. 6 amu.",
            "D. 48 amu."
        ],
        "answer": "A",
        "solution": (
            "Khối lượng nguyên tử carbon tính theo đơn vị amu là:\n"
            "m_C = (1,9926 × 10^-23^) / (1,6605 × 10^-24^) = 12 amu.\n"
            "➔ **Chọn A.**"
        )
    },
    {
        "id": 8,
        "content": "Thời tiết nồm ẩm vào cuối mùa xuân có bản chất là hiện tượng:",
        "options": [
            "*A. ngưng tụ hơi nước.*",
            "B. bay hơi nước ở thể lỏng.",
            "C. đông đặc nước ở thể lỏng.",
            "D. nóng chảy nước ở thể rắn."
        ],
        "answer": "A",
        "solution": (
            "Vào cuối mùa xuân, độ ẩm không khí rất cao. Khi hơi ẩm tiếp xúc với bề mặt sàn nhà hoặc tường có nhiệt độ thấp hơn nhiệt độ điểm sương, hơi nước sẽ ngưng tụ lại thành các giọt nước đọng trên bề mặt, gây ra hiện tượng nồm ẩm.\n"
            "➔ **Chọn A.**"
        )
    },
    {
        "id": 9,
        "content": "Chất nào sau đây không phải là hợp chất cộng hóa trị?",
        "options": [
            "A. SO~2~.",
            "B. CO~2~.",
            "*C. NaCl.*",
            "D. H~2~O."
        ],
        "answer": "C",
        "solution": (
            "- SO~2~, CO~2~, H~2~O là các hợp chất được tạo thành giữa các nguyên tử phi kim với nhau bằng liên kết cộng hóa trị.\n"
            "- NaCl được tạo thành từ kim loại điển hình (Na) và phi kim điển hình (Cl) bằng lực hút tĩnh điện giữa các ion Na^+^ và Cl^-^, do đó NaCl là hợp chất ion, không phải hợp chất cộng hóa trị.\n"
            "➔ **Chọn C.**"
        )
    },
    {
        "id": 10,
        "content": "Cho biết tổng số hạt trong hạt nhân nguyên tử Nitrogen là 14, số đơn vị điện tích hạt nhân là 7. Số hạt neutron và electron của nguyên tử Nitrogen:",
        "options": [
            "A. 7 và 14",
            "B. 14 và 7",
            "*C. 7 và 7.*",
            "D. 14 và 14."
        ],
        "answer": "C",
        "solution": (
            "- Số đơn vị điện tích hạt nhân là 7 ⇒ số proton p = 7.\n"
            "- Do nguyên tử trung hòa về điện ⇒ số electron e = p = 7.\n"
            "- Tổng số hạt trong hạt nhân là 14 ⇒ p + n = 14 ⇒ số neutron n = 14 - 7 = 7.\n"
            "Vậy số hạt neutron là 7 và electron là 7.\n"
            "➔ **Chọn C.**"
        )
    },
    {
        "id": 11,
        "content": "Cho các phát biểu:\n(I) Nguyên tử trung hoà về điện.\n(II) Khối lượng của nguyên tử tập trung chủ yếu ở hạt nhân.\n(III) Trong nguyên tử, số hạt mang điện tích dương bằng số hạt mang điện tích âm nên số hạt electron bằng số hạt neutron.\n(IV) Vỏ nguyên tử, gồm các lớp electron có khoảng cách khác nhau đối với hạt nhân.\nTrong các phát biểu trên, số phát biểu đúng là:",
        "options": [
            "A. 1.",
            "B. 2.",
            "*C. 3.*",
            "D. 4."
        ],
        "answer": "C",
        "solution": (
            "- (I) Đúng. Nguyên tử có số hạt mang điện dương (proton) bằng số hạt mang điện âm (electron) nên trung hòa về điện.\n"
            "- (II) Đúng. Khối lượng electron rất nhỏ bé (khoảng 0,00055 amu) nên khối lượng nguyên tử tập trung chủ yếu ở hạt nhân.\n"
            "- (III) Sai. Hạt mang điện dương là proton, mang điện âm là electron, do đó số hạt proton bằng số hạt electron, không phải neutron.\n"
            "- (IV) Đúng. Vỏ nguyên tử gồm các lớp electron có mức năng lượng và khoảng cách khác nhau đối với hạt nhân.\n"
            "Vậy có 3 phát biểu đúng là: (I), (II) và (IV).\n"
            "➔ **Chọn C.**"
        )
    },
    {
        "id": 12,
        "content": "Cho các hợp chất sau: MgCl~2~, Na~2~O, NH~3~, HCl, KCl. Hợp chất nào sau đây có liên kết cộng hóa trị?",
        "options": [
            "*A. NH~3~ và HCl*",
            "B. Na~2~O và NH~3~.",
            "C. MgCl~2~ và Na~2~O.",
            "D. HCl và KCl"
        ],
        "answer": "A",
        "solution": (
            "- Hợp chất cộng hóa trị hình thành giữa các phi kim: NH~3~ (N và H) và HCl (H và Cl).\n"
            "- Các hợp chất MgCl~2~, Na~2~O, KCl đều tạo nên từ kim loại và phi kim nên có liên kết ion.\n"
            "➔ **Chọn A.**"
        )
    },
    {
        "id": 13,
        "content": "Hợp chất Na~x~SO~4~ có khối lượng phân tử là 142 amu. Giá trị của x là:",
        "options": [
            "A. 1.",
            "*B. 2.*",
            "C. 3.",
            "D. 4."
        ],
        "answer": "B",
        "solution": (
            "Khối lượng phân tử của Na~x~SO~4~ là:\n"
            "M = 23 × x + 32 + 16 × 4 = 142\n"
            "⇒ 23x + 96 = 142\n"
            "⇒ 23x = 46 ⇒ x = 2.\n"
            "(Hợp chất là Na~2~SO~4~ - sodium sulfate).\n"
            "➔ **Chọn B.**"
        )
    },
    {
        "id": 14,
        "content": "Nguyên tố X có 12 proton, 3 lớp electron, có 2 e lớp ngoài cùng. Vị trí của nguyên tố X trong bảng tuần hoàn là:",
        "options": [
            "A. X ở ô số 12, chu kì 3, nhóm III",
            "B. X ở ô số 12, chu kì 2, nhóm I.",
            "*C. X ở ô số 12, chu kì 3, nhóm II*",
            "D. X ở ô số 12, chu kì 2, nhóm III"
        ],
        "answer": "C",
        "solution": (
            "- Số proton = 12 ⇒ X ở ô số 12 trong bảng tuần hoàn.\n"
            "- Số lớp electron = 3 ⇒ X thuộc chu kì 3.\n"
            "- Số electron lớp ngoài cùng = 2 ⇒ X thuộc nhóm II (hoặc nhóm IIA).\n"
            "➔ **Chọn C.**"
        )
    },
    {
        "id": 15,
        "content": "Trong phân tử nước, cứ 16,0 g oxygen có tương ứng 2,0 g hydrogen. Một giọt nước chứa 0,1 g hydrogen thì khối lượng của oxygen có trong giọt nước đó là:",
        "options": [
            "A. 1,6 g.",
            "B. 1,2 g.",
            "C. 0,9g.",
            "*D. 0,8 g.*"
        ],
        "answer": "D",
        "solution": (
            "Tỉ lệ khối lượng giữa oxygen và hydrogen trong phân tử nước là:\n"
            "m_O / m_H = 16,0 / 2,0 = 8.\n"
            "Khối lượng oxygen tương ứng với 0,1 g hydrogen là:\n"
            "m_O = 0,1 × 8 = 0,8 g.\n"
            "➔ **Chọn D.**"
        )
    },
    {
        "id": 16,
        "content": "Khối lượng phân tử sulfur dioxide (SO~2~) là:",
        "options": [
            "*A. 64 amu*",
            "B. 48 amu",
            "C. 16 amu",
            "D. 80 amu"
        ],
        "answer": "A",
        "solution": (
            "Khối lượng phân tử của SO~2~ là:\n"
            "M(SO~2~) = M_S + 2 × M_O = 32 + 2 × 16 = 64 amu.\n"
            "➔ **Chọn A.**"
        )
    }
]

PART2_QUESTIONS = [
    {
        "id": 17,
        "title": "Một số đặc điểm của đơn chất:",
        "statements": [
            ("a) Các nguyên tử trong đơn chất kim loại sắp xếp khít nhau theo một trật tự xác định và có vai trò như phân tử.", "Đúng"),
            ("b) Các đơn chất kim loại có tính chất vật lí chung như: Dẫn điện, dẫn nhiệt, có ánh kim,...", "Đúng"),
            ("c) Đơn chất khí hiếm tồn tại trong tự nhiên ở dạng phân tử.", "Sai"),
            ("d) Các đơn chất phi kim có tính chất khác với các đơn chất kim loại.", "Đúng")
        ],
        "solution": (
            "- **a) Đúng.** Trong đơn chất kim loại, các nguyên tử sắp xếp khít nhau theo trật tự xác định và hạt hợp thành có vai trò như phân tử.\n"
            "- **b) Đúng.** Các kim loại đều có các tính chất vật lí đặc trưng: tính dẻo, tính dẫn điện, dẫn nhiệt tốt và có ánh kim.\n"
            "- **c) Sai.** Khí hiếm tồn tại trong tự nhiên ở dạng các nguyên tử đơn độc lập (đơn nguyên tử), không liên kết tạo thành phân tử.\n"
            "- **d) Đúng.** Phi kim thường không dẫn điện, không dẫn nhiệt và không có ánh kim dẻo như kim loại."
        )
    },
    {
        "id": 18,
        "title": "Nếu một nguyên tử có 14 electron thì:",
        "statements": [
            ("a) Nguyên tử đó có 14 proton.", "Đúng"),
            ("b) Nguyên tử đó có 2 lớp electron.", "Sai"),
            ("c) Nguyên tử có số đơn vị điện tích hạt nhân là 14.", "Đúng"),
            ("d) Nguyên tử có số electron lớp ngoài cùng là 3.", "Sai")
        ],
        "solution": (
            "- **a) Đúng.** Nguyên tử trung hòa về điện nên số proton = số electron = 14.\n"
            "- **b) Sai.** Phân bố 14 electron vào các lớp: Lớp 1 có 2e, Lớp 2 có 8e, Lớp 3 có 14 - 2 - 8 = 4e. Vậy nguyên tử có 3 lớp electron.\n"
            "- **c) Đúng.** Số đơn vị điện tích hạt nhân = số proton = 14.\n"
            "- **d) Sai.** Lớp ngoài cùng (lớp 3) có 4 electron (nguyên tố Silicon)."
        )
    },
    {
        "id": 19,
        "title": "Cấu tạo nguyên tử",
        "statements": [
            ("a) Nguyên tử gồm 3 loại hạt cơ bản là proton, neutron và electron.", "Đúng"),
            ("b) Lớp vỏ nguyên tử gồm 2 loại hạt là electron và neutron.", "Sai"),
            ("c) Hạt nhân nguyên tử gồm 2 loại hạt là proton và neutron.", "Đúng"),
            ("d) Nguyên tử luôn trung hòa về điện nên số hạt proton luôn bằng số hạt neutron.", "Sai")
        ],
        "solution": (
            "- **a) Đúng.** Hầu hết các nguyên tử đều được tạo nên từ 3 loại hạt cơ bản: proton, neutron và electron.\n"
            "- **b) Sai.** Vỏ nguyên tử chỉ gồm duy nhất các hạt electron mang điện tích âm.\n"
            "- **c) Đúng.** Hạt nhân nằm ở tâm nguyên tử, gồm hạt proton (mang điện dương) và neutron (không mang điện).\n"
            "- **d) Sai.** Nguyên tử trung hòa về điện do số hạt mang điện dương (proton) bằng số hạt mang điện âm (electron): p = e, không liên quan đến số neutron."
        )
    }
]

PART3_QUESTIONS = [
    {
        "id": 20,
        "content": "Cho hợp chất P~2~O~5~. Khối lượng phân tử của hợp chất trên là bao nhiêu amu?",
        "answer": "142",
        "solution": (
            "Khối lượng phân tử của P~2~O~5~ là:\n"
            "M(P~2~O~5~) = 2 × M_P + 5 × M_O = 2 × 31 + 5 × 16 = 62 + 80 = 142 amu.\n"
            "**Đáp số: 142.**"
        )
    },
    {
        "id": 21,
        "content": "Thành phần phần trăm theo khối lượng của nguyên tố P trong hợp chất P~2~O~5~ là bao nhiêu phần trăm ? (Tính ra kết quả số thập phân và làm tròn đến 2 chữ số sau dấu phẩy ).",
        "answer": "43,66%",
        "solution": (
            "Khối lượng của nguyên tố P trong một phân tử P~2~O~5~ là: 2 × 31 = 62 amu.\n"
            "Khối lượng phân tử P~2~O~5~ là 142 amu.\n"
            "Phần trăm khối lượng của nguyên tố P là:\n"
            "%m_P = (62 / 142) × 100% ≈ 43,66%.\n"
            "**Đáp số: 43,66% (hoặc 43,66).**"
        )
    },
    {
        "id": 22,
        "content": "Nguyên tử Aluminium có điện tích hạt nhân là 13+. Trong nguyên tử, số hạt mang điện nhiều hơn số hạt không mang điện là 12 hạt. Số hạt neutron của nguyên tử Aluminium là bao nhiêu ?",
        "answer": "14",
        "solution": (
            "- Điện tích hạt nhân là 13+ ⇒ số proton p = 13.\n"
            "- Nguyên tử trung hòa về điện ⇒ số electron e = p = 13.\n"
            "- Tổng số hạt mang điện: p + e = 13 + 13 = 26.\n"
            "- Số hạt không mang điện là neutron (n).\n"
            "Theo đề bài: (p + e) - n = 12 ⇒ 26 - n = 12 ⇒ n = 26 - 12 = 14.\n"
            "Vậy số hạt neutron của nguyên tử Aluminium là 14.\n"
            "**Đáp số: 14.**"
        )
    },
    {
        "id": 23,
        "content": "Trong hạt nhân nguyên tử fluorine có 9 proton. Số electron ở lớp ngoài cùng của vỏ nguyên tử fluorine là bao nhiêu?",
        "answer": "7",
        "solution": (
            "- Nguyên tử fluorine có 9 proton ⇒ vỏ có 9 electron.\n"
            "- Sự phân bố electron vào các lớp:\n"
            "  + Lớp thứ nhất (lớp trong cùng) chứa tối đa 2 electron.\n"
            "  + Lớp thứ hai (lớp ngoài cùng) chứa: 9 - 2 = 7 electron.\n"
            "Vậy số electron ở lớp ngoài cùng của nguyên tử fluorine là 7.\n"
            "**Đáp số: 7.**"
        )
    },
    {
        "id": 24,
        "content": "Cho hợp chất K~2~S. Khối lượng phân tử của hợp chất trên là bao nhiêu amu?",
        "answer": "110",
        "solution": (
            "Khối lượng phân tử của K~2~S là:\n"
            "M(K~2~S) = 2 × M_K + 1 × M_S = 2 × 39 + 32 = 78 + 32 = 110 amu.\n"
            "**Đáp số: 110.**"
        )
    },
    {
        "id": 25,
        "content": "Một nguyên tử có 10 proton trong hạt nhân. Theo mô hình nguyên tử của Rơ- dơ- pho- Bo, nguyên tử trên có bao nhiêu lớp electron ?",
        "answer": "2",
        "solution": (
            "- Nguyên tử có 10 proton ⇒ vỏ có 10 electron.\n"
            "- Theo mô hình Rutherford – Bohr, sự phân bố electron vào các lớp như sau:\n"
            "  + Lớp thứ nhất có tối đa 2 electron.\n"
            "  + Lớp thứ hai có tối đa 8 electron (2 + 8 = 10 electron, lớp 2 bão hòa).\n"
            "Vậy vỏ nguyên tử có 2 lớp electron (đây là nguyên tố Neon).\n"
            "**Đáp số: 2.**"
        )
    }
]

# ==============================================================================
# HÀM TẠO FILE WORD (.DOCX) CHUẨN FORM ĐỀ KIỂM TRA
# ==============================================================================

def generate_word_document(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="TRƯỜNG THCS ............................",
        exam_title="ĐỀ KIỂM TRA, ĐÁNH GIÁ GIỮA HỌC KÌ I",
        subject_line="MÔN: KHOA HỌC TỰ NHIÊN 7",
        num_pages="02",
        duration="60 phút",
        exam_code="01",
        atomic_masses="H= 1; Na= 23; S= 32; P= 31; O= 16; K= 39 (amu)",
        is_teacher=is_teacher
    )

    # PHẦN I
    add_section_header(
        doc,
        title="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
        subtitle="Mỗi câu hỏi học sinh chỉ chọn 1 phương án."
    )

    for q in PART1_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(1)
        render_question_prompt(p_q, f"Câu {q['id']}. {q['content']}")
        render_options_no_table(doc, q['options'], is_teacher=is_teacher)
        if is_teacher:
            add_solution_no_table(doc, q['solution'])

    # PHẦN II
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai.",
        subtitle="Trong mỗi ý a, b, c, d học sinh chọn đúng (Đ) hoặc sai (S)."
    )

    for q in PART2_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(1)
        render_question_prompt(p_q, f"Câu {q['id']}. {q['title']}")

        for stmt, ans in q['statements']:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt, ans=ans, is_teacher=is_teacher)

        if is_teacher:
            add_solution_no_table(doc, q['solution'])

    # PHẦN III
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm trả lời ngắn.",
        subtitle="Thí sinh ghi câu trả lời ngắn vào dòng tương ứng."
    )

    for q in PART3_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(1)
        render_question_prompt(p_q, f"Câu {q['id']}. {q['content']}")

        if not is_teacher:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.5)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(3)
            render_rich_text(p_ans, "*Trả lời:* ........................................................................................................", base_size=10.5)
        else:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.5)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(1)
            r_ans_lbl = p_ans.add_run("Trả lời: ")
            set_run_style(r_ans_lbl, font_size=10.5, bold=True)
            r_ans_val = p_ans.add_run(q['answer'])
            set_run_style(r_ans_val, font_size=10.5, bold=True, color=COLOR_CORRECT)
            add_solution_no_table(doc, q['solution'])

    # Lời chúc kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(10)
    p_end.paragraph_format.space_after = Pt(6)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    render_rich_text(p_end, "------------------ *Chúc các em làm bài thi tốt !* ------------------", base_size=10.5, base_italic=True)

    filename = "De_Kiem_Tra_GK1_KHTN7_Ma01_GiaoVien.docx" if is_teacher else "De_Kiem_Tra_GK1_KHTN7_Ma01_HocSinh.docx"
    output_path = os.path.join(OUTPUT_DIR, filename)
    doc.save(output_path)
    print(f"[OK] Đã tạo file Word: {output_path}")
    return output_path

# ==============================================================================
# HÀM TẠO FILE LATEX VÀ BIÊN DỊCH PDF CHUẨN A4
# ==============================================================================

def to_latex(text):
    if not text:
        return ""
    s = text
    # Thay thế các kí tự đặc biệt toán học / hoá học
    s = s.replace("➔", r"$\rightarrow$ ")
    s = s.replace("⇒", r"$\Rightarrow$ ")
    s = s.replace("≈", r"$\approx$ ")
    s = s.replace("×", r"$\times$ ")
    s = s.replace("·", r"$\cdot$ ")
    s = s.replace("...", r"\dots ")
    
    # Kí hiệu số mũ và chỉ số dưới thông dụng
    s = s.replace("1,9926 × 10^-23^", r"$1{,}9926 \times 10^{-23}$")
    s = s.replace("1,6605 × 10^-24^", r"$1{,}6605 \times 10^{-24}$")
    s = s.replace("1,9926 × 10^-23", r"$1{,}9926 \times 10^{-23}$")
    s = s.replace("1,6605 × 10^-24", r"$1{,}6605 \times 10^{-24}$")
    
    # Biến số có chỉ số dưới
    s = re.sub(r'\bm_([pnCOHP])\b', r'$m_{\1}$', s)
    s = re.sub(r'\bM_([SOPK])\b', r'$M_{\1}$', s)
    s = re.sub(r'\bM\(([^)]+)\)', r'$M(\text{\1})$', s)
    s = re.sub(r'%m_P', r'$\%m_P$', s)
    s = re.sub(r'%m_([A-Za-z]+)', r'$\%m_{\1}$', s)
    s = re.sub(r'(?<!\\)%', r'\%', s)

    # In đậm **text**
    s = re.sub(r'\*\*(.*?)\*\*', r'\\textbf{\1}', s)

    # Subscript ~text~ và Superscript ^text^
    s = re.sub(r'~([^~]+)~', r'$_{\1}$', s)
    s = re.sub(r'\^([^^]+)\^', r'$^{\1}$', s)

    # Xuống dòng
    s = s.replace("\n", r"\\ ")
    return s

def generate_latex_a4(is_teacher=False):
    tag = "gv" if is_teacher else "hs"
    tex_filename = f"De_Kiem_Tra_GK1_KHTN7_Ma01_{'GiaoVien' if is_teacher else 'HocSinh'}.tex"
    tex_path = os.path.join(OUTPUT_DIR, tex_filename)

    lines = []
    lines.append(r"""\documentclass[11pt,a4paper]{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{geometry}
\geometry{top=1.5cm,bottom=1.2cm,left=1.4cm,right=1.4cm,headheight=28pt,headsep=12pt,footskip=18pt}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{tabularx}
\usepackage{array}
\usepackage{tikz}
\usepackage{fancyhdr}
\usepackage{enumitem}

\definecolor{hfRed}{HTML}{D32F2F}
\definecolor{hfGreen}{HTML}{8BC390}
\definecolor{mainGreen}{HTML}{2E7D32}

\usepackage{ex_test}

\pagestyle{fancy}
\fancyhf{}
\renewcommand{\headrulewidth}{0.75pt}
\renewcommand{\headrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \headrulewidth\hfill}}
\renewcommand{\footrulewidth}{0.75pt}
\renewcommand{\footrule}{\hbox to\headwidth{\color{hfGreen}\leaders\hrule height \footrulewidth\hfill}}

\fancyhead[L]{\small\itshape\color{hfRed}Tài liệu lưu hành nội bộ}
\fancyhead[R]{\small\color{hfRed}Thầy \textbf{TRẦN VĂN THẠNH} - 0777.470.803}

\fancyfoot[L]{\footnotesize\itshape\color{hfRed}CS1: 6/15 Nguyễn Hoàng, P. Kim Long, TP Huế \quad|\quad CS2: 24 Đặng Thái Thân, TP Huế}
\fancyfoot[R]{\small\textbf{\color{hfRed}Trang \thepage}}

\setlength{\parindent}{0pt}
\setlength{\parskip}{2pt}

\begin{document}

\begin{minipage}[t]{0.45\textwidth}
\textbf{TRƯỜNG THCS \dots\dots\dots\dots\dots\dots\dots\dots\dots}\\
\textbf{ĐỀ KIỂM TRA ĐÁNH GIÁ}\\
\textit{ĐỀ CHÍNH THỨC}\\
\textit{(Đề thi có 02 trang)}
\end{minipage}
\hfill
\begin{minipage}[t]{0.52\textwidth}
\raggedleft
\textbf{\color{hfRed}ĐỀ KIỂM TRA GIỮA HỌC KÌ I -- NĂM HỌC 2025 -- 2026}\\
\textbf{MÔN: KHOA HỌC TỰ NHIÊN 7}\\
\textit{Thời gian làm bài: 60 phút}\\
\textbf{MÃ ĐỀ THI: 01}
\end{minipage}

\vspace{3pt}
\noindent
\textit{Họ, tên thí sinh:} \dotfill \textit{Lớp:} \dotfill \textit{SBD:} \dotfill

\vspace{2pt}
\noindent
\textit{\footnotesize (Cho biết nguyên tử khối của các nguyên tố: H = 1; Na = 23; S = 32; P = 31; O = 16; K = 39 amu).}
""")

    if is_teacher:
        lines.append(r"""
\vspace{2pt}
\begin{center}
\textbf{\color{hfRed}\large (BẢN GIÁO VIÊN -- CÓ ĐÁP ÁN VÀ LỜI GIẢI CHI TIẾT)}
\end{center}
""")

    lines.append(r"""
\vspace{3pt}
\noindent
\textbf{\color{mainGreen}PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.}\\
\textit{Mỗi câu hỏi học sinh chỉ chọn 1 phương án.}

\Opensolutionfile{ans}[ans-khtn7-""" + tag + r"""]
""")

    for q in PART1_QUESTIONS:
        tex_opts = []
        for opt in q['options']:
            is_c = opt.startswith("*") and opt.endswith("*")
            clean_o = opt.strip("*").strip()
            clean_o = re.sub(r'^[A-D]\.\s*', '', clean_o)
            clean_o = to_latex(clean_o)
            if is_c:
                tex_opts.append(r"{\True " + clean_o + r"}")
            else:
                tex_opts.append(r"{" + clean_o + r"}")

        prompt = to_latex(q['content'])

        lines.append(r"\begin{ex}")
        lines.append(prompt)
        lines.append(r"\choice")
        for to in tex_opts:
            lines.append(to)

        if is_teacher:
            sol = to_latex(q['solution'])
            lines.append(r"\loigiai{" + sol + r"}")

        lines.append(r"\end{ex}")
        lines.append("")

    lines.append(r"\Closesolutionfile{ans}")

    # PHẦN II
    lines.append(r"""
\vspace{4pt}
\noindent
\textbf{\color{mainGreen}PHẦN II. Câu trắc nghiệm đúng sai.}\\
\textit{Trong mỗi ý a, b, c, d học sinh chọn đúng (Đ) hoặc sai (S).}

\setcounter{ex}{0}
""")

    for q in PART2_QUESTIONS:
        title = to_latex(q['title'])
        lines.append(r"\begin{ex}")
        lines.append(title)
        lines.append(r"\choiceTF")
        for stmt, ans in q['statements']:
            body = re.sub(r'^[a-d]\)\s*', '', stmt)
            body = to_latex(body)
            if ans == "Đúng":
                lines.append(r"{\True " + body + r"}")
            else:
                lines.append(r"{" + body + r"}")

        if is_teacher:
            sol = to_latex(q['solution'])
            lines.append(r"\loigiai{" + sol + r"}")

        lines.append(r"\end{ex}")
        lines.append("")

    # PHẦN III
    lines.append(r"""
\vspace{4pt}
\noindent
\textbf{\color{mainGreen}PHẦN III. Câu trắc nghiệm trả lời ngắn.}\\
\textit{Thí sinh ghi câu trả lời ngắn vào dòng tương ứng.}

\setcounter{ex}{0}
""")

    for q in PART3_QUESTIONS:
        prompt = to_latex(q['content'])

        lines.append(r"\begin{ex}")
        lines.append(prompt)
        if not is_teacher:
            lines.append(r"\par\textbf{Trả lời:} \dotfill")
        else:
            ans_clean = to_latex(q['answer'])
            lines.append(r"\par\textbf{Trả lời:} \textbf{\color{hfRed}" + ans_clean + r"}")
            sol = to_latex(q['solution'])
            lines.append(r"\loigiai{" + sol + r"}")
        lines.append(r"\end{ex}")
        lines.append("")

    lines.append(r"""
\vspace{8pt}
\begin{center}
\textit{------------------ Chúc các em làm bài thi tốt ! ------------------}
\end{center}

\end{document}
""")

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"[OK] Đã tạo file LaTeX: {tex_path}")

    # Biên dịch pdflatex
    try:
        cmd = ["pdflatex", "-interaction=nonstopmode", tex_filename]
        subprocess.run(cmd, cwd=OUTPUT_DIR, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        subprocess.run(cmd, cwd=OUTPUT_DIR, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        pdf_path = os.path.join(OUTPUT_DIR, tex_filename.replace(".tex", ".pdf"))
        if os.path.exists(pdf_path):
            print(f"[OK] Đã biên dịch PDF thành công: {pdf_path}")
        else:
            print(f"[CẢNH BÁO] Không tìm thấy file PDF sau biên dịch: {pdf_path}")
    except Exception as e:
        print(f"[LỖI] Lỗi biên dịch pdflatex: {e}")

    # Xóa file rác
    for ext in [".aux", ".log", ".thm", ".out"]:
        junk = os.path.join(OUTPUT_DIR, tex_filename.replace(".tex", ext))
        if os.path.exists(junk):
            try:
                os.remove(junk)
            except:
                pass
    junk_ans = os.path.join(OUTPUT_DIR, f"ans-khtn7-{tag}.tex")
    if os.path.exists(junk_ans):
        try:
            os.remove(junk_ans)
        except:
            pass

# ==============================================================================
# HÀM TẠO SLIDE TRÌNH CHIẾU TIVI 65 INCH (16:9 WIDESCREEN) KHÔNG ĐÁP ÁN
# ==============================================================================

def generate_latex_tivi_16x9():
    tex_filename = "De_Kiem_Tra_GK1_KHTN7_Ma01_TrinhChieu_Tivi_16x9.tex"
    tex_path = os.path.join(OUTPUT_DIR, tex_filename)

    lines = []
    lines.append(r"""\documentclass[14pt]{article}
\usepackage[utf8]{inputenc}
\usepackage[vietnamese]{babel}
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{geometry}
\geometry{paperwidth=29.7cm,paperheight=16.706cm,top=0.9cm,bottom=0.9cm,left=1.2cm,right=1.2cm,headheight=0pt,headsep=0pt,footskip=0pt}
\pagestyle{empty}
\usepackage{tabularx}
\usepackage{array}
\usepackage{tikz}
\usepackage{xcolor}

\definecolor{hfRed}{HTML}{D32F2F}
\definecolor{mainBlue}{HTML}{0D47A1}
\definecolor{mainGreen}{HTML}{1B5E20}

\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}

\begin{document}
""")

    # Slide Trang bìa / Giới thiệu
    lines.append(r"""
\begin{center}
\vspace*{\fill}
{\fontsize{34}{44}\selectfont \textbf{\color{hfRed}ĐỀ KIỂM TRA, ĐÁNH GIÁ GIỮA HỌC KÌ I}}\\[14pt]
{\fontsize{28}{36}\selectfont \textbf{\color{mainBlue}MÔN: KHOA HỌC TỰ NHIÊN 7 -- NĂM HỌC 2025 -- 2026}}\\[12pt]
{\fontsize{24}{32}\selectfont \textbf{MÃ ĐỀ: 01 \quad|\quad THỜI GIAN LÀM BÀI: 60 PHÚT}}\\[20pt]
{\fontsize{20}{26}\selectfont \textit{Giảng dạy trực tiếp trên màn hình Tivi tương tác}}\\[10pt]
{\fontsize{22}{28}\selectfont \textbf{\color{mainGreen}Thầy TRẦN VĂN THẠNH -- 0777.470.803}}
\vspace*{\fill}
\end{center}
\clearpage
""")

    # PHẦN I: 2 câu / 1 slide
    for i in range(0, len(PART1_QUESTIONS), 2):
        q1 = PART1_QUESTIONS[i]
        q2 = PART1_QUESTIONS[i+1] if i+1 < len(PART1_QUESTIONS) else None

        lines.append(r"{\fontsize{20}{26}\selectfont")
        
        # Render câu 1
        p1 = q1['content']
        p1 = re.sub(r'~([^~]+)~', r'$_{\1}$', p1)
        p1 = re.sub(r'\^([^^]+)\^', r'$^{\1}$', p1)
        p1 = p1.replace("1,9926 × 10^-23", r"$1{,}9926 \times 10^{-23}$")
        p1 = p1.replace("1,6605 × 10^-24", r"$1{,}6605 \times 10^{-24}$")
        lines.append(r"\textbf{\color{mainBlue}Câu " + str(q1['id']) + r".} " + p1 + r"\\[6pt]")

        # Options dạng 2 cột tabularx
        clean_opts1 = []
        for opt in q1['options']:
            o = opt.strip("*").strip()
            lbl = o[:2]
            body = o[2:].strip()
            body = re.sub(r'~([^~]+)~', r'$_{\1}$', body)
            body = re.sub(r'\^([^^]+)\^', r'$^{\1}$', body)
            clean_opts1.append(r"\textbf{" + lbl + r"} " + body)

        lines.append(r"\noindent\begin{tabularx}{\textwidth}{@{}X X@{}}")
        lines.append(clean_opts1[0] + r" & " + clean_opts1[1] + r"\\[4pt]")
        lines.append(clean_opts1[2] + r" & " + clean_opts1[3])
        lines.append(r"\end{tabularx}")

        if q2:
            lines.append(r"\vspace*{\fill}")
            p2 = q2['content']
            p2 = re.sub(r'~([^~]+)~', r'$_{\1}$', p2)
            p2 = re.sub(r'\^([^^]+)\^', r'$^{\1}$', p2)
            p2 = p2.replace("1,9926 × 10^-23", r"$1{,}9926 \times 10^{-23}$")
            p2 = p2.replace("1,6605 × 10^-24", r"$1{,}6605 \times 10^{-24}$")
            lines.append(r"\textbf{\color{mainBlue}Câu " + str(q2['id']) + r".} " + p2 + r"\\[6pt]")

            clean_opts2 = []
            for opt in q2['options']:
                o = opt.strip("*").strip()
                lbl = o[:2]
                body = o[2:].strip()
                body = re.sub(r'~([^~]+)~', r'$_{\1}$', body)
                body = re.sub(r'\^([^^]+)\^', r'$^{\1}$', body)
                clean_opts2.append(r"\textbf{" + lbl + r"} " + body)

            lines.append(r"\noindent\begin{tabularx}{\textwidth}{@{}X X@{}}")
            lines.append(clean_opts2[0] + r" & " + clean_opts2[1] + r"\\[4pt]")
            lines.append(clean_opts2[2] + r" & " + clean_opts2[3])
            lines.append(r"\end{tabularx}")

        lines.append(r"}")
        lines.append(r"\clearpage")

    # PHẦN II: 1 câu / 1 slide
    for q in PART2_QUESTIONS:
        lines.append(r"{\fontsize{22}{29}\selectfont")
        lines.append(r"\textbf{\color{mainBlue}Câu " + str(q['id']) + r".} " + q['title'] + r"\\[12pt]")
        for stmt, _ in q['statements']:
            body = stmt
            body = re.sub(r'~([^~]+)~', r'$_{\1}$', body)
            body = re.sub(r'\^([^^]+)\^', r'$^{\1}$', body)
            lines.append(r"\noindent\begin{tabularx}{\textwidth}{@{}X r@{}}")
            lines.append(body + r" & \textbf{[ \quad ]} \\")
            lines.append(r"\end{tabularx}\\[8pt]")
        lines.append(r"\vspace*{\fill}")
        lines.append(r"}")
        lines.append(r"\clearpage")

    # PHẦN III: 1 câu / 1 slide
    for q in PART3_QUESTIONS:
        prompt = q['content']
        prompt = re.sub(r'~([^~]+)~', r'$_{\1}$', prompt)
        prompt = re.sub(r'\^([^^]+)\^', r'$^{\1}$', prompt)

        lines.append(r"{\fontsize{24}{32}\selectfont")
        lines.append(r"\textbf{\color{mainBlue}Câu " + str(q['id']) + r".} " + prompt + r"\\[20pt]")
        lines.append(r"\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$\\[20pt]")
        lines.append(r"\vspace*{\fill}")
        lines.append(r"}")
        lines.append(r"\clearpage")

    # Slide BẢNG ĐÁP ÁN CUỐI CÙNG
    lines.append(r"""
\begin{center}
{\fontsize{26}{34}\selectfont \textbf{\color{hfRed}BẢNG ĐÁP ÁN ĐỀ KIỂM TRA GIỮA KÌ I -- MÃ ĐỀ 01}}\\[16pt]
{\fontsize{18}{24}\selectfont
\textbf{\color{mainBlue}PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN}\\[6pt]
\begin{tabular}{|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|c|}
\hline
\textbf{1} & \textbf{2} & \textbf{3} & \textbf{4} & \textbf{5} & \textbf{6} & \textbf{7} & \textbf{8} & \textbf{9} & \textbf{10} & \textbf{11} & \textbf{12} & \textbf{13} & \textbf{14} & \textbf{15} & \textbf{16} \\
\hline
\textbf{B} & \textbf{D} & \textbf{D} & \textbf{C} & \textbf{A} & \textbf{B} & \textbf{A} & \textbf{A} & \textbf{C} & \textbf{C} & \textbf{C} & \textbf{A} & \textbf{B} & \textbf{C} & \textbf{D} & \textbf{A} \\
\hline
\end{tabular}\\[16pt]

\textbf{\color{mainBlue}PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI}\\[6pt]
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Ý a} & \textbf{Ý b} & \textbf{Ý c} & \textbf{Ý d} \\
\hline
\textbf{17} & Đúng & Đúng & Sai & Đúng \\
\hline
\textbf{18} & Đúng & Sai & Đúng & Sai \\
\hline
\textbf{19} & Đúng & Sai & Đúng & Sai \\
\hline
\end{tabular}\\[16pt]

\textbf{\color{mainBlue}PHẦN III. TRẢ LỜI NGẮN}\\[6pt]
\begin{tabular}{|c|c|c|c|c|c|}
\hline
\textbf{Câu 20} & \textbf{Câu 21} & \textbf{Câu 22} & \textbf{Câu 23} & \textbf{Câu 24} & \textbf{Câu 25} \\
\hline
\textbf{142} & \textbf{43,66\%} & \textbf{14} & \textbf{7} & \textbf{11}0 & \textbf{2} \\
\hline
\end{tabular}
}
\end{center}
\clearpage
""")

    lines.append(r"\end{document}")

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"[OK] Đã tạo file LaTeX Tivi: {tex_path}")

    # Biên dịch pdflatex
    try:
        cmd = ["pdflatex", "-interaction=nonstopmode", tex_filename]
        subprocess.run(cmd, cwd=OUTPUT_DIR, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        pdf_path = os.path.join(OUTPUT_DIR, tex_filename.replace(".tex", ".pdf"))
        if os.path.exists(pdf_path):
            print(f"[OK] Đã biên dịch PDF Tivi thành công: {pdf_path}")
        else:
            print(f"[CẢNH BÁO] Không tìm thấy file PDF Tivi sau biên dịch: {pdf_path}")
    except Exception as e:
        print(f"[LỖI] Lỗi biên dịch pdflatex Tivi: {e}")

    for ext in [".aux", ".log", ".thm", ".out"]:
        junk = os.path.join(OUTPUT_DIR, tex_filename.replace(".tex", ext))
        if os.path.exists(junk):
            try:
                os.remove(junk)
            except:
                pass

# ==============================================================================
# HÀM SAO CHÉP ONEDRIVE
# ==============================================================================

def sync_to_onedrive():
    files_to_sync = [
        "De_Kiem_Tra_GK1_KHTN7_Ma01_HocSinh.docx",
        "De_Kiem_Tra_GK1_KHTN7_Ma01_GiaoVien.docx",
        "De_Kiem_Tra_GK1_KHTN7_Ma01_HocSinh.pdf",
        "De_Kiem_Tra_GK1_KHTN7_Ma01_GiaoVien.pdf",
        "De_Kiem_Tra_GK1_KHTN7_Ma01_TrinhChieu_Tivi_16x9.pdf"
    ]
    for target_dir in TARGET_ONEDRIVE_DIRS:
        if os.path.exists(target_dir):
            for fname in files_to_sync:
                src = os.path.join(OUTPUT_DIR, fname)
                if os.path.exists(src):
                    dst = os.path.join(target_dir, fname)
                    try:
                        shutil.copy2(src, dst)
                        print(f"[ONEDRIVE] Đã đồng bộ sang {dst}")
                    except Exception as e:
                        print(f"[ONEDRIVE] Lỗi đồng bộ {fname}: {e}")

# ==============================================================================
# CHƯƠNG TRÌNH CHÍNH
# ==============================================================================

if __name__ == "__main__":
    print("=== BẮT ĐẦU TẠO ĐỀ KIỂM TRA GIỮA KÌ I KHTN 7 (MÃ ĐỀ 01) ===")
    
    # 1. Tạo Word
    generate_word_document(is_teacher=False)
    generate_word_document(is_teacher=True)
    
    # 2. Tạo LaTeX & PDF A4
    generate_latex_a4(is_teacher=False)
    generate_latex_a4(is_teacher=True)
    
    # 3. Tạo Slide Tivi 16:9
    generate_latex_tivi_16x9()
    
    # 4. Đồng bộ OneDrive
    sync_to_onedrive()
    
    print("=== HOÀN THÀNH BIÊN SOẠN BỘ TÀI LIỆU ===")
