# -*- coding: utf-8 -*-
r"""
Script biên soạn trọn bộ tài liệu chuẩn FORM ĐỀ KIỂM TRA cho:
ĐỀ CƯƠNG ÔN TẬP GIỮA KÌ I – PHÂN MÔN VẬT LÍ – KHTN 7
TRƯỜNG THCS NGUYỄN HOÀNG – NĂM HỌC 2026 - 2027
1. De_Cuong_GK1_KHTN7_NguyenHoang_HocSinh.docx
2. De_Cuong_GK1_KHTN7_NguyenHoang_GiaoVien.docx
3. De_Cuong_GK1_KHTN7_NguyenHoang_DayTrenManHinh.pdf (Slide 16:9 Widescreen KHÔNG CÓ ĐÁP ÁN)
Và bản Word: De_Cuong_GK1_KHTN7_NguyenHoang_DayTrenManHinh.docx
Thầy TRẦN VĂN THẠNH - 0777.470.803
"""

import os
import sys
import subprocess
import shutil

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

import docx
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Import các hàm chuẩn từ form_de_kiem_tra
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
IMAGES_DIR = r"c:\Antigravity_Thanh\San_Pham\images_khtn7_nguyenhoang"
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)
os.makedirs(IMAGES_DIR, exist_ok=True)

TARGET_ONEDRIVE_DIRS = [
    r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn7\KHTN7-IN-DEKIEMTRADINHKI",
    r"c:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn7"
]

# ==============================================================================
# DỮ LIỆU ĐỀ CƯƠNG KHTN 7 NGUYỄN HOÀNG
# ==============================================================================

THEORY_DATA = [
    ("1. Tốc độ chuyển động", [
        "Tốc độ là đại lượng cho biết mức độ nhanh hay chậm của chuyển động.",
        "Tốc độ chuyển động của một vật được xác định bằng chiều dài quãng đường vật đi được trong một đơn vị thời gian.",
        "Công thức tính tốc độ: *v = s / t*",
        "Trong đó: *s* là quãng đường đi được (m hoặc km); *t* là thời gian đi (s hoặc h); *v* là tốc độ (m/s hoặc km/h)."
    ]),
    ("2. Đơn vị tốc độ", [
        "Trong hệ đo lường chính thức (SI) ở nước ta, đơn vị của tốc độ là mét trên giây (m/s) hoặc kilômét trên giờ (km/h).",
        "Ngoài ra có thể dùng: mét trên phút (m/min), xentimét trên giây (cm/s)...",
        "Cách đổi đơn vị: *1 m/s = 3,6 km/h*; *1 km/h = 1/3,6 m/s ≈ 0,28 m/s*."
    ]),
    ("3. Đồ thị quãng đường – thời gian", [
        "Đồ thị quãng đường – thời gian mô tả liên hệ giữa quãng đường đi được của vật với thời gian.",
        "Vận dụng đồ thị: Từ đồ thị cho trước, có thể tìm được quãng đường vật đi, thời gian chuyển động hoặc xác định tốc độ của vật (*v = s / t*).",
        "Đoạn đồ thị nằm ngang biểu diễn vật đứng yên (quãng đường không đổi theo thời gian).",
        "Đoạn đồ thị càng dốc thì tốc độ chuyển động của vật càng lớn."
    ])
]

MCQ_DATA = [
    {
        "id": 1,
        "content": "Tốc độ cho biết điều gì về chuyển động của vật?",
        "options": [
            "A. Mức độ nhanh hay chậm của chuyển động.",
            "B. Hướng chuyển động của vật.",
            "C. Quãng đường vật đi được.",
            "D. Thời gian chuyển động của vật."
        ],
        "answer": "A",
        "solution": "Tốc độ là đại lượng vật lí đặc trưng cho mức độ nhanh hay chậm của chuyển động."
    },
    {
        "id": 2,
        "content": "Công thức nào sau đây dùng để tính tốc độ của một vật?",
        "options": [
            "A. v = s . t",
            "B. v = s / t",
            "C. v = t / s",
            "D. v = s"
        ],
        "answer": "B",
        "solution": "Tốc độ bằng chiều dài quãng đường đi được chia cho thời gian đi quãng đường đó: *v = s / t*."
    },
    {
        "id": 3,
        "content": "Đơn vị nào sau đây là đơn vị đo tốc độ trong hệ đo lường chính thức của nước ta?",
        "options": [
            "A. m/s",
            "B. km/min.",
            "C. km/s",
            "D. m/phút"
        ],
        "answer": "A",
        "solution": "Trong hệ đơn vị đo lường chính thức (SI) của nước ta, đơn vị đo tốc độ là mét trên giây (*m/s*) và kilômét trên giờ (*km/h*)."
    },
    {
        "id": 4,
        "content": "Tốc độ 36 km/h ứng với giá trị nào sau đây?",
        "options": [
            "A. 3,6 m/s",
            "B. 12 m/s",
            "C. 10 m/s",
            "D. 1 m/s"
        ],
        "answer": "C",
        "solution": "Cách đổi từ km/h sang m/s: chia cho 3,6. Ta có: *v = 36 / 3,6 = 10 m/s*."
    },
    {
        "id": 5,
        "content": "Đơn vị của tốc độ phụ thuộc vào:",
        "options": [
            "A. đơn vị đo chiều dài.",
            "B. đơn vị đo độ dài và đơn vị đo thời gian.",
            "C. đơn vị đo thời gian.",
            "D. các yếu tố khác."
        ],
        "answer": "B",
        "solution": "Vì công thức tính tốc độ là *v = s / t* nên đơn vị của tốc độ phụ thuộc đồng thời vào đơn vị đo độ dài (quãng đường) và đơn vị đo thời gian."
    },
    {
        "id": 6,
        "content": "Một người đi bộ quãng đường 100 m trong 20 giây. Tốc độ của người đó là",
        "options": [
            "A. 5 m/s",
            "B. 20 m/s",
            "C. 100 m/s",
            "D. 0,2 m/s"
        ],
        "answer": "A",
        "solution": "Tốc độ của người đó: *v = s / t = 100 / 20 = 5 m/s*."
    },
    {
        "id": 7,
        "content": "Đồ thị quãng đường – thời gian cho biết:",
        "options": [
            "A. Tốc độ đi được.",
            "B. Thời gian đi được.",
            "C. Quãng đường đi được.",
            "D. Cả tốc độ, thời gian và quãng đường đi được."
        ],
        "answer": "D",
        "solution": "Đồ thị quãng đường – thời gian mô tả trực tiếp quãng đường theo thời gian, từ đó ta cũng tính được tốc độ chuyển động (*v = s / t*)."
    },
    {
        "id": 8,
        "content": "Bạn A chạy quãng đường 60 m hết 10 giây. Bạn B chạy quãng đường 100 m hết 20 giây. Nhận xét nào sau đây là đúng?",
        "options": [
            "A. Bạn A đi nhanh hơn bạn B.",
            "B. Bạn B đi nhanh hơn bạn A.",
            "C. Hai bạn đi nhanh như nhau.",
            "D. Không thể so sánh được."
        ],
        "answer": "A",
        "solution": "Tốc độ của bạn A: *v_A = 60 / 10 = 6 m/s*. Tốc độ của bạn B: *v_B = 100 / 20 = 5 m/s*. Vì *v_A > v_B* nên bạn A chạy nhanh hơn bạn B."
    },
    {
        "id": 9,
        "content": "Một ô tô đi 100 m đầu tiên trong 10 giây, sau đó đi tiếp 200 m trong 20 giây. Tốc độ chuyển động của ô tô trên cả quãng đường là",
        "options": [
            "A. 10 m/s",
            "B. 15 m/s",
            "C. 30 m/s",
            "D. 100 m/s"
        ],
        "answer": "A",
        "solution": "Tổng quãng đường: *s = 100 + 200 = 300 m*. Tổng thời gian: *t = 10 + 20 = 30 s*. Tốc độ: *v = 300 / 30 = 10 m/s*."
    },
    {
        "id": 10,
        "content": "Trên đồ thị quãng đường – thời gian, đường thẳng càng dốc thể hiện:",
        "options": [
            "A. Chuyển động nhanh hơn.",
            "B. Chuyển động chậm hơn.",
            "C. Vật đứng yên.",
            "D. Vật chuyển động đều."
        ],
        "answer": "A",
        "solution": "Độ dốc của đồ thị tỉ lệ thuận với tốc độ: đường thẳng càng dốc thì trong cùng một khoảng thời gian quãng đường đi được càng lớn, tức vật chuyển động càng nhanh."
    }
]

TF_DATA = [
    {
        "id": 11,
        "prompt": "Câu 11. Xác định mỗi phát biểu sau là đúng hay sai:",
        "items": [
            {
                "id": "a",
                "text": "Tốc độ càng lớn thì vật chuyển động càng chậm.",
                "ans": False,
                "reason": "Tốc độ càng lớn thì vật chuyển động càng nhanh."
            },
            {
                "id": "b",
                "text": "Khi vật đứng yên, tốc độ của vật bằng 0.",
                "ans": True,
                "reason": "Vật đứng yên nghĩa là không dịch chuyển (quãng đường s = 0) nên v = 0."
            },
            {
                "id": "c",
                "text": "Công thức tính tốc độ là v = t : s.",
                "ans": False,
                "reason": "Công thức đúng là v = s : t (hoặc v = s / t)."
            },
            {
                "id": "d",
                "text": "Đơn vị tốc độ có thể là km/h hoặc m/s.",
                "ans": True,
                "reason": "km/h và m/s là hai đơn vị vận tốc / tốc độ thông dụng nhất."
            }
        ]
    },
    {
        "id": 12,
        "prompt": "Câu 12. Xác định mỗi phát biểu sau là đúng hay sai:",
        "items": [
            {
                "id": "a",
                "text": "Trên cùng một quãng đường, nếu thời gian chuyển động càng ngắn thì chuyển động đó nhanh hơn.",
                "ans": True,
                "reason": "Thời gian t càng nhỏ thì thương số v = s / t càng lớn, tức chuyển động nhanh hơn."
            },
            {
                "id": "b",
                "text": "Đồ thị quãng đường – thời gian là đường thẳng đi qua gốc tọa độ.",
                "ans": True,
                "reason": "Với chuyển động thẳng đều xuất phát từ gốc thời gian và gốc tọa độ (t = 0, s = 0), đồ thị là đường thẳng qua gốc O."
            },
            {
                "id": "c",
                "text": "Đồ thị càng dốc thì tốc độ càng nhỏ.",
                "ans": False,
                "reason": "Đồ thị càng dốc thì tốc độ càng lớn."
            },
            {
                "id": "d",
                "text": "Người đi xe máy 60 km trong 1 giờ có tốc độ 60 km/h.",
                "ans": True,
                "reason": "Tốc độ v = s / t = 60 / 1 = 60 km/h."
            }
        ]
    }
]

SHORT_DATA = [
    {
        "id": 13,
        "content": "Đơn vị tốc độ chuyển động hợp pháp (trong hệ SI) được kí hiệu là .......",
        "answer": "m/s (hoặc mét trên giây)",
        "solution": "Đơn vị đo tốc độ hợp pháp trong hệ đo lường chính thức của nước ta là *m/s*."
    },
    {
        "id": 14,
        "content": "Đổi các đơn vị sau sang giá trị tương ứng:\na) 1 km/h = ...... m/s; 1 m/s = ...... km/h.\nb) 5 m/s = ...... km/h.\nc) 45 km/h = ...... m/s.\nd) 720 km/h = ...... m/s = ...... cm/s.",
        "answer": "a) ≈ 0,28 m/s (hoặc 5/18 m/s); 3,6 km/h.   b) 18 km/h.   c) 12,5 m/s.   d) 200 m/s = 20.000 cm/s.",
        "solution": "• 1 km/h = 1000 m / 3600 s ≈ *0,28 m/s* (hoặc 5/18 m/s); 1 m/s = *3,6 km/h*.\n• 5 m/s = 5 × 3,6 = *18 km/h*.\n• 45 km/h = 45 / 3,6 = *12,5 m/s*.\n• 720 km/h = 720 / 3,6 = *200 m/s* = *20.000 cm/s*."
    },
    {
        "id": 15,
        "content": "Quãng đường vật đi được đo bằng đơn vị là mét (m), thời gian đo bằng đơn vị là giây (s) thì tốc độ chuyển động có đơn vị là ......",
        "answer": "m/s (hoặc mét trên giây)",
        "solution": "Theo công thức *v = s / t*, với s tính bằng m và t tính bằng s thì đơn vị tốc độ là *m/s*."
    },
    {
        "id": 16,
        "content": "Nếu một vật đi cùng một quãng đường nhưng thời gian đi hết quãng đường đó tăng gấp đôi thì tốc độ chuyển động của vật giảm đi ...... lần.",
        "answer": "2",
        "solution": "Vì *v = s / t*, khi s không đổi và t tăng gấp 2 lần thì tốc độ giảm đi *2 lần*."
    },
    {
        "id": 17,
        "content": "Trong cùng một khoảng thời gian, nếu quãng đường chuyển động ...... hơn thì chuyển động đó nhanh hơn.",
        "answer": "dài (hoặc lớn)",
        "solution": "Với thời gian t không đổi, quãng đường s càng *dài* (lớn) thì tốc độ *v = s / t* càng lớn, tức chuyển động nhanh hơn."
    },
    {
        "id": 18,
        "content": "Một người đi xe đạp trong 3 giờ đi được quãng đường 45 km. Tốc độ của người đó là ...... km/h.",
        "answer": "15",
        "solution": "Tốc độ: *v = s / t = 45 / 3 = 15 km/h*."
    },
    {
        "id": 19,
        "content": "Một vận động viên chạy 400 m hết 50 giây. Tốc độ của vận động viên đó là ...... m/s.",
        "answer": "8",
        "solution": "Tốc độ: *v = s / t = 400 / 50 = 8 m/s*."
    },
    {
        "id": 20,
        "content": "Công thức tính tốc độ là ......",
        "answer": "v = s / t (hoặc v = s : t)",
        "solution": "Công thức chuẩn: *v = s / t*."
    },
    {
        "id": 21,
        "content": "Trong hệ thống đo lường chính thức ở nước ta, tốc độ được đo bằng đơn vị m/s và ......",
        "answer": "km/h (hoặc kilômét trên giờ)",
        "solution": "Hai đơn vị thông dụng hợp pháp là *m/s* và *km/h*."
    },
    {
        "id": 22,
        "content": "Đại lượng đặc trưng cho mức độ nhanh chậm của chuyển động là ......",
        "answer": "tốc độ (hoặc tốc độ chuyển động)",
        "solution": "Đó chính là khái niệm *tốc độ*."
    },
    {
        "id": 23,
        "content": "Vật đứng yên thì tốc độ bằng ......",
        "answer": "0 (hoặc 0 m/s)",
        "solution": "Vật không di chuyển nên quãng đường s = 0, tốc độ bằng *0*."
    },
    {
        "id": 24,
        "content": "Dụng cụ đo tốc độ của các phương tiện giao thông là ......",
        "answer": "tốc kế (hoặc đồng hồ đo tốc độ)",
        "solution": "Dụng cụ gắn trên xe để đo tốc độ là *tốc kế* (speedometer)."
    }
]

ESSAY_DATA = [
    {
        "id": 25,
        "content": "Một xe tải chạy trên đoạn đường đầu dài 45 km trong 45 phút, sau đó xe tiếp tục chạy thêm 18 km trong 20 phút. Tính tốc độ của xe tải trên mỗi đoạn đường và trên cả hai đoạn đường.",
        "solution": (
            "• Đổi đơn vị thời gian:\n"
            "  t_1 = 45 phút = 45 / 60 h = 0,75 h.\n"
            "  t_2 = 20 phút = 20 / 60 h = 1/3 h ≈ 0,333 h.\n"
            "• Tốc độ của xe tải trên đoạn đường đầu:\n"
            "  v_1 = s_1 / t_1 = 45 / 0,75 = *60 (km/h)*.\n"
            "• Tốc độ của xe tải trên đoạn đường sau:\n"
            "  v_2 = s_2 / t_2 = 18 / (1/3) = *54 (km/h)*.\n"
            "• Tốc độ của xe tải trên cả hai đoạn đường:\n"
            "  Tổng quãng đường: s = s_1 + s_2 = 45 + 18 = 63 (km).\n"
            "  Tổng thời gian: t = 45 + 20 = 65 phút = 65 / 60 h = 13/12 h.\n"
            "  v = s / t = 63 / (13/12) = 756 / 13 ≈ *58,15 (km/h)* (hoặc ≈ 16,15 m/s)."
        )
    },
    {
        "id": 26,
        "content": "Một học sinh đi xe đạp từ nhà đến trường mất 20 phút. Đoạn đường từ nhà đến trường dài 4 km.\na) Tính tốc độ của chuyển động theo đơn vị km/h?\nb) Ý nghĩa của con số tìm ở câu a?",
        "solution": (
            "a) Tính tốc độ:\n"
            "   Đổi t = 20 phút = 20 / 60 h = 1/3 h.\n"
            "   Tốc độ đi xe đạp của học sinh là:\n"
            "   v = s / t = 4 / (1/3) = *12 (km/h)*.\n"
            "b) Ý nghĩa con số 12 km/h:\n"
            "   Con số 12 km/h cho biết mức độ nhanh chậm của chuyển động, nghĩa là trung bình trong mỗi giờ học sinh đó đi được quãng đường dài 12 km (nếu cứ giữ nguyên tốc độ đó thì sau 1 giờ học sinh sẽ đi được 12 km)."
        )
    },
    {
        "id": 27,
        "content": "Dựa vào bảng ghi số liệu dưới đây về quãng đường và thời gian của một người đi bộ, em hãy vẽ đồ thị quãng đường – thời gian của người này:\n\n[BẢNG SỐ LIỆU]\nThời gian t (h): 0 | 0,5 | 1,0 | 1,5 | 2,0 | 2,5\nQuãng đường s (km): 0 | 25 | 50 | 75 | 75 | 100",
        "solution": (
            "• Các bước vẽ đồ thị quãng đường – thời gian:\n"
            "  1. Vẽ hệ trục tọa độ vuông góc: Trục hoành nằm ngang biểu diễn thời gian t (h), tỉ lệ 1 cm ứng với 0,5 h. Trục tung thẳng đứng biểu diễn quãng đường s (km), tỉ lệ 1 cm ứng với 25 km.\n"
            "  2. Xác định các điểm tọa độ từ bảng số liệu: O(0; 0), A(0,5; 25), B(1,0; 50), C(1,5; 75), D(2,0; 75), E(2,5; 100).\n"
            "  3. Nối các điểm bằng các đoạn thẳng:\n"
            "     - Từ t = 0 đến t = 1,5 h: Nối từ O đến C là đường thẳng dốc lên biểu diễn người đó đi bộ đều với tốc độ v = 75 / 1,5 = 50 km/h.\n"
            "     - Từ t = 1,5 h đến t = 2,0 h: Đoạn CD nằm ngang song song trục thời gian biểu diễn người đó dừng lại nghỉ 0,5 h (quãng đường không đổi s = 75 km).\n"
            "     - Từ t = 2,0 h đến t = 2,5 h: Đoạn DE dốc lên biểu diễn người đó tiếp tục đi bộ quãng đường 25 km trong 0,5 h với tốc độ v = 25 / 0,5 = 50 km/h."
        )
    },
    {
        "id": 28,
        "content": "Dựa vào các thông tin về quãng đường và thời gian của một người đi xe đạp trong bảng dưới đây, hãy mô tả đặc điểm chuyển động của người đó:\n\n[BẢNG SỐ LIỆU]\nThời gian t (s): 0 | 2 | 4 | 6 | 8 | 10 | 12\nQuãng đường s (m): 0 | 10 | 20 | 40 | 40 | 60 | 40",
        "solution": (
            "• Phân tích và mô tả từng giai đoạn chuyển động:\n"
            "  1. Giai đoạn 1 (từ 0 s đến 4 s): Quãng đường tăng đều 10 m sau mỗi 2 s (từ 0 đến 20 m). Người đó chuyển động đều với tốc độ v_1 = 10 / 2 = *5 (m/s)*.\n"
            "  2. Giai đoạn 2 (từ 4 s đến 6 s): Quãng đường tăng từ 20 m lên 40 m trong 2 s. Tốc độ tăng lên: v_2 = (40 - 20) / 2 = *10 (m/s)*.\n"
            "  3. Giai đoạn 3 (từ 6 s đến 8 s): Quãng đường giữ nguyên s = 40 m. Người đi xe đạp *đứng yên (dừng lại nghỉ)* trong 2 s (tốc độ v_3 = 0).\n"
            "  4. Giai đoạn 4 (từ 8 s đến 10 s): Quãng đường tăng từ 40 m lên 60 m trong 2 s. Tốc độ di chuyển: v_4 = (60 - 40) / 2 = *10 (m/s)*.\n"
            "  5. Giai đoạn 5 (từ 10 s đến 12 s): Vị trí giảm từ 60 m về 40 m. Người đó *đổi chiều chuyển động (quay lại)* với tốc độ v_5 = |40 - 60| / 2 = *10 (m/s)*."
        )
    },
    {
        "id": 29,
        "content": "Dựa vào đồ thị quãng đường – thời gian của một ô tô (hình vẽ bên) để trả lời các câu hỏi sau:\na) Trên đoạn đường nào xe chuyển động nhanh hơn? Xác định tốc độ của xe trên mỗi đoạn đường.\nb) Sau 50 giây, xe đi được bao nhiêu mét?",
        "image": "dothi_cau29_hd.png",
        "solution": (
            "a) Xác định tốc độ trên mỗi đoạn đường:\n"
            "   • Đoạn đường (1): từ t = 0 đến t = 40 s, xe đi được s_1 = 450 m.\n"
            "     Tốc độ của xe trên đoạn (1) là: v_1 = s_1 / t_1 = 450 / 40 = *11,25 (m/s)*.\n"
            "   • Đoạn đường (2): từ t = 40 s đến t = 60 s (thời gian t_2 = 60 - 40 = 20 s), xe đi thêm quãng đường s_2 = 900 - 450 = 450 m.\n"
            "     Tốc độ của xe trên đoạn (2) là: v_2 = s_2 / t_2 = 450 / 20 = *22,5 (m/s)*.\n"
            "   • Kết luận: Vì v_2 = 22,5 m/s > v_1 = 11,25 m/s (gấp 2 lần), trên đoạn đường (2) đồ thị dốc hơn nên *xe chuyển động nhanh hơn trên đoạn đường (2)*.\n\n"
            "b) Quãng đường xe đi được sau 50 giây:\n"
            "   • Thời điểm t = 50 s nằm trong giai đoạn (2). Thời gian xe chuyển động trong giai đoạn (2) là: Δt = 50 - 40 = 10 s.\n"
            "   • Quãng đường xe đi thêm được trong 10 s này là: Δs = v_2 × Δt = 22,5 × 10 = 225 (m).\n"
            "   • Vậy sau 50 s, tổng quãng đường xe đi được là: s = 450 + 225 = *675 (m)*."
        )
    },
    {
        "id": 30,
        "content": "Hình vẽ bên biểu diễn đồ thị quãng đường – thời gian của một ô tô trên đường phố vào giờ cao điểm trong hành trình dài 4 phút.\na) Mô tả các giai đoạn chuyển động của ô tô trên đồ thị.\nb) Xác định thời gian ô tô đã dừng lại trong hành trình.\nc) Tốc độ của ô tô trong giai đoạn nào là lớn nhất?",
        "image": "dothi_cau30_hd.png",
        "solution": (
            "a) Mô tả các giai đoạn chuyển động của ô tô:\n"
            "   • Giai đoạn 1 (đoạn O → A, từ t = 0 đến t = 1 min): Ô tô chuyển động đều từ gốc vị trí, sau 1 phút đi được quãng đường 100 m. Tốc độ: v_1 = 100 m / 60 s ≈ 1,67 m/s (6 km/h).\n"
            "   • Giai đoạn 2 (đoạn A → B, từ t = 1 min đến t = 3 min): Đồ thị là đoạn thẳng nằm ngang song song trục thời gian, quãng đường không đổi s = 100 m. Ô tô *dừng lại (đứng yên)* do ùn tắc hoặc dừng đèn đỏ.\n"
            "   • Giai đoạn 3 (đoạn B → C, từ t = 3 min đến t = 4 min): Ô tô tiếp tục chuyển động đều, sau 1 phút đi thêm được Δs = 300 - 100 = 200 m. Tốc độ: v_3 = 200 m / 60 s ≈ 3,33 m/s (12 km/h).\n\n"
            "b) Thời gian ô tô đã dừng lại trong hành trình:\n"
            "   Ô tô dừng lại trong giai đoạn 2 (đoạn nằm ngang A → B):\n"
            "   t_dừng = 3 min - 1 min = *2 phút* (tương ứng *120 giây*).\n\n"
            "c) Giai đoạn có tốc độ lớn nhất:\n"
            "   Tốc độ trong giai đoạn 3 (từ phút thứ 3 đến phút thứ 4) là *lớn nhất* (v_3 = 12 km/h so với v_1 = 6 km/h và v_2 = 0), biểu hiện trên đồ thị là đoạn đường B → C có độ dốc lớn nhất."
        )
    }
]

# ==============================================================================
# 1. HÀM TẠO FILE WORD
# ==============================================================================

def generate_word_document(mode="student"):
    is_teacher = (mode == "teacher")
    is_screen = (mode == "screen")

    if mode == "teacher":
        filename = "De_Cuong_GK1_KHTN7_NguyenHoang_GiaoVien.docx"
    elif mode == "screen":
        filename = "De_Cuong_GK1_KHTN7_NguyenHoang_DayTrenManHinh.docx"
    else:
        filename = "De_Cuong_GK1_KHTN7_NguyenHoang_HocSinh.docx"

    filepath = os.path.join(OUTPUT_DIR, filename)

    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="TRƯỜNG THCS NGUYỄN HOÀNG",
        exam_title="ĐỀ CƯƠNG ÔN TẬP GIỮA KÌ I – KHTN 7",
        subject_line="PHÂN MÔN: VẬT LÍ (TỐC ĐỘ & ĐỒ THỊ QUÃNG ĐƯỜNG - THỜI GIAN)",
        num_pages="04",
        duration="45 phút",
        exam_code="KHTN7-NH",
        atomic_masses=None,
        is_teacher=is_teacher
    )

    base_size = 11.5 if is_screen else 10.5

    # A. TÓM TẮT LÝ THUYẾT
    add_section_header(doc, "A. TÓM TẮT LÝ THUYẾT TRỌNG TÂM", "Học sinh ghi nhớ các định nghĩa, công thức và quy tắc đổi đơn vị:")
    for title, points in THEORY_DATA:
        p_th = doc.add_paragraph()
        p_th.paragraph_format.left_indent = Cm(0.3)
        p_th.paragraph_format.space_before = Pt(3)
        p_th.paragraph_format.space_after = Pt(1)
        render_rich_text(p_th, f"*{title}*", base_size=base_size, base_color=COLOR_PRIMARY)
        for pt in points:
            p_pt = doc.add_paragraph()
            p_pt.paragraph_format.left_indent = Cm(0.6)
            p_pt.paragraph_format.space_before = Pt(1)
            p_pt.paragraph_format.space_after = Pt(1)
            render_rich_text(p_pt, f"• {pt}", base_size=base_size)

    # B. CÂU HỎI ÔN TẬP
    # PHẦN 1: TRẮC NGHIỆM NHIỀU LỰA CHỌN
    add_section_header(doc, "B. CÂU HỎI ÔN TẬP\nI. TRẮC NGHIỆM NHIỀU LỰA CHỌN (10 CÂU)", "Chọn phương án trả lời đúng nhất trong mỗi câu sau:")
    for q in MCQ_DATA:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, f"Câu {q['id']}. {q['content']}", base_size=base_size)

        render_options_no_table(doc, q["options"], is_teacher=is_teacher)

        if is_teacher and q.get("solution"):
            add_solution_no_table(doc, f"Đáp án đúng: *{q['answer']}*.\n{q['solution']}")
        elif is_screen:
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.left_indent = Cm(0.6)
            p_sp.paragraph_format.space_before = Pt(2)
            p_sp.paragraph_format.space_after = Pt(2)
            render_rich_text(p_sp, "✎ Ghi chú / Giải thích: ........................................................................................................", base_size=base_size-1)

    # PHẦN 2: TRẮC NGHIỆM ĐÚNG - SAI
    add_section_header(doc, "II. TRẮC NGHIỆM ĐÚNG – SAI (2 CÂU)", "Trong mỗi ý ở mỗi câu, học sinh chọn Đúng hoặc Sai:")
    for q in TF_DATA:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"], base_size=base_size)

        for item in q["items"]:
            p_it = doc.add_paragraph()
            p_it.paragraph_format.left_indent = Cm(0.5)
            p_it.paragraph_format.space_before = Pt(1)
            p_it.paragraph_format.space_after = Pt(1)
            ans_str = "Đúng" if item["ans"] else "Sai"

            if is_teacher:
                render_statement(p_it, f"{item['id']}) {item['text']}", ans=ans_str, is_teacher=True, base_size=base_size)
                p_rea = doc.add_paragraph()
                p_rea.paragraph_format.left_indent = Cm(1.0)
                p_rea.paragraph_format.space_before = Pt(0)
                p_rea.paragraph_format.space_after = Pt(2)
                render_rich_text(p_rea, f"↳ _Giải thích:_ {item['reason']}", base_size=base_size-1, base_italic=True, base_color=COLOR_SOLUTION)
            else:
                render_rich_text(p_it, f"{item['id']}) {item['text']}   [   ] Đúng     [   ] Sai", base_size=base_size)

    # PHẦN 3: TRẮC NGHIỆM TRẢ LỜI NGẮN
    add_section_header(doc, "III. TRẮC NGHIỆM TRẢ LỜI NGẮN (12 CÂU)", "Học sinh tính toán và viết câu trả lời ngắn gọn vào chỗ trống:")
    for q in SHORT_DATA:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, f"Câu {q['id']}. {q['content']}", base_size=base_size)

        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.8)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(2)

        if is_teacher:
            render_rich_text(p_ans, f"➔ *Đáp án:* {q['answer']}", base_size=base_size, base_color=COLOR_CORRECT)
            add_solution_no_table(doc, q["solution"])
        elif is_screen:
            render_rich_text(p_ans, "➔ Trả lời: ...................................................................................................................", base_size=base_size)
        else:
            render_rich_text(p_ans, "Trả lời: .....................................................................................................................", base_size=base_size)

    # PHẦN 4: PHẦN TỰ LUẬN
    add_section_header(doc, "IV. PHẦN TỰ LUẬN (6 CÂU)", "Học sinh trình bày chi tiết lời giải và vẽ đồ thị vào bài làm:")
    for q in ESSAY_DATA:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, f"Câu {q['id']}. {q['content']}", base_size=base_size)

        # Chèn hình vẽ nếu có
        if q.get("image"):
            img_path = os.path.join(IMAGES_DIR, q["image"])
            if os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(3)
                p_img.paragraph_format.space_after = Pt(3)
                p_img.add_run().add_picture(img_path, width=Inches(3.2))

        if is_teacher:
            add_solution_no_table(doc, q["solution"])
        elif is_screen:
            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.left_indent = Cm(0.6)
            render_rich_text(p_sp, "✎ Bài giải chi tiết: ........................................................................................................", base_size=base_size)
        else:
            num_lines = 5 if q["id"] < 29 else 4
            add_student_writing_lines(doc, num_lines=num_lines)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(12)
    p_end.paragraph_format.space_after = Pt(6)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("--- HẾT ---")
    set_run_style(r_end, font_size=11, bold=True, italic=True, color=COLOR_PRIMARY)

    doc.save(filepath)
    print(f"[OK] Đã tạo thành công file Word: {filepath} ({os.path.getsize(filepath):,} bytes)")
    return filepath


# ==============================================================================
# 2. HÀM TẠO FILE LATEX & PDF TRÌNH CHIẾU TIVI 16:9 (KHÔNG CÓ ĐÁP ÁN)
# ==============================================================================

def generate_tivi_presentation_pdf():
    tex_filename = "De_Cuong_GK1_KHTN7_NguyenHoang_DayTrenManHinh.tex"
    pdf_filename = "De_Cuong_GK1_KHTN7_NguyenHoang_DayTrenManHinh.pdf"
    tex_path = os.path.join(SCRATCH_DIR, tex_filename)
    pdf_dest = os.path.join(OUTPUT_DIR, pdf_filename)

    tex_lines = []
    # Preamble chuẩn 16:9 widescreen
    tex_lines.append(r"""\documentclass{article}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{tabularx}
\usepackage{array}
\usepackage{booktabs}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage[
  paperwidth=29.7cm,
  paperheight=16.706cm,
  top=0.9cm,
  bottom=0.9cm,
  left=1.2cm,
  right=1.2cm,
  footskip=0pt,
  headheight=0pt,
  headsep=0pt
]{geometry}

\pagestyle{empty}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}

\definecolor{brandRed}{RGB}{211, 47, 47}
\definecolor{brandGreen}{RGB}{46, 125, 50}
\definecolor{brandBlue}{RGB}{21, 101, 192}
\definecolor{textColor}{RGB}{20, 20, 20}

\color{textColor}

\begin{document}
""")

    # --------------------------------------------------------------------------
    # SLIDE 0: TIÊU ĐỀ
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 0: TIÊU ĐỀ ĐỀ CƯƠNG
% ==============================================================================
\begin{center}
\vspace*{0.2cm}
{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}TRƯỜNG THCS NGUYỄN HOÀNG -- NĂM HỌC 2026 -- 2027}}\\[10pt]
{\fontsize{34}{42}\selectfont \textbf{\color{brandBlue}ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA KÌ I}}\\[8pt]
{\fontsize{28}{36}\selectfont \textbf{\color{brandGreen}MÔN: KHOA HỌC TỰ NHIÊN 7 -- PHÂN MÔN VẬT LÍ}}\\[8pt]
{\fontsize{22}{28}\selectfont \textit{(Chủ đề: Tốc độ chuyển động -- Đồ thị quãng đường -- thời gian)}}\\[16pt]
\rule{16cm}{1.5pt}\\[20pt]
{\fontsize{26}{34}\selectfont \textbf{Giáo viên hướng dẫn: Thầy TRẦN VĂN THẠNH -- 0777.470.803}}\\[8pt]
{\fontsize{22}{28}\selectfont \textit{Cơ sở 1: 6/15 Nguyễn Hoàng, Kim Long | Cơ sở 2: 24 Đặng Thái Thân, TP Huế}}\\[16pt]
{\fontsize{24}{30}\selectfont \textbf{CẤU TRÚC: LÝ THUYẾT -- TRẮC NGHIỆM ABCD -- ĐÚNG/SAI -- TRẢ LỜI NGẮN -- TỰ LUẬN}}
\end{center}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 1: TÓM TẮT LÝ THUYẾT (TỐC ĐỘ & ĐƠN VỊ)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 1: TÓM TẮT LÝ THUYẾT (TỐC ĐỘ & ĐƠN VỊ)
% ==============================================================================
{\fontsize{25}{32}\selectfont \textbf{\color{brandBlue}A. TÓM TẮT LÝ THUYẾT TRỌNG TÂM (PHẦN 1)}}
\vspace{8pt}

\fontsize{21}{28}\selectfont
\textbf{1. Tốc độ chuyển động:}
\begin{itemize}
  \setlength{\itemsep}{2pt}
  \item Tốc độ là đại lượng đặc trưng cho \textbf{mức độ nhanh hay chậm} của chuyển động.
  \item Xác định bằng chiều dài quãng đường vật đi được trong một đơn vị thời gian.
  \item \textbf{Công thức tính tốc độ:} $v = \dfrac{s}{t}$
  \item \textit{Trong đó:} $s$ là quãng đường đi được ($\text{m}$ hoặc $\text{km}$); $t$ là thời gian đi ($\text{s}$ hoặc $\text{h}$); $v$ là tốc độ ($\text{m/s}$ hoặc $\text{km/h}$).
\end{itemize}

\vspace{6pt}
\textbf{2. Đơn vị tốc độ:}
\begin{itemize}
  \setlength{\itemsep}{2pt}
  \item Đơn vị hợp pháp trong hệ SI của nước ta là: \textbf{mét trên giây ($\text{m/s}$)} và \textbf{kilômét trên giờ ($\text{km/h}$)}.
  \item Quy tắc đổi đơn vị: \textbf{$1\text{ m/s} = 3,6\text{ km/h}$} \quad $\Longleftrightarrow$ \quad \textbf{$1\text{ km/h} = \dfrac{1}{3,6}\text{ m/s} \approx 0,28\text{ m/s}$}.
\end{itemize}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 2: TÓM TẮT LÝ THUYẾT (ĐỒ THỊ QUÃNG ĐƯỜNG - THỜI GIAN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 2: TÓM TẮT LÝ THUYẾT (ĐỒ THỊ QUÃNG ĐƯỜNG - THỜI GIAN)
% ==============================================================================
{\fontsize{25}{32}\selectfont \textbf{\color{brandBlue}A. TÓM TẮT LÝ THUYẾT TRỌNG TÂM (PHẦN 2)}}
\vspace{10pt}

\fontsize{22}{30}\selectfont
\textbf{3. Đồ thị quãng đường – thời gian ($s - t$):}
\begin{itemize}
  \setlength{\itemsep}{6pt}
  \item Mô tả mối quan hệ giữa \textbf{quãng đường đi được ($s$)} và \textbf{thời gian chuyển động ($t$)}.
  \item Trục hoành nằm ngang biểu diễn thời gian $t$; trục tung thẳng đứng biểu diễn quãng đường $s$.
\end{itemize}

\vspace{10pt}
\textbf{4. Vận dụng đồ thị quãng đường – thời gian:}
\begin{itemize}
  \setlength{\itemsep}{6pt}
  \item Từ đồ thị có thể tìm được: quãng đường vật đi, thời gian chuyển động hoặc tính tốc độ: $v = \dfrac{s}{t}$.
  \item \textbf{Đoạn đồ thị nằm ngang:} Vật đang \textbf{đứng yên} (quãng đường không đổi theo thời gian).
  \item \textbf{Độ dốc của đồ thị:} Đồ thị \textbf{càng dốc} thì tốc độ chuyển động \textbf{càng lớn} (vật đi càng nhanh).
\end{itemize}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 3: TRẮC NGHIỆM CÂU 1 & CÂU 2
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 3: TRẮC NGHIỆM CÂU 1 & CÂU 2
% ==============================================================================
\fontsize{26}{35}\selectfont
\textbf{Câu 1.} Tốc độ cho biết điều gì về chuyển động của vật?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Mức độ nhanh hay chậm của chuyển động. & \textbf{B.} Hướng chuyển động của vật.\\[6pt]
\textbf{C.} Quãng đường vật đi được. & \textbf{D.} Thời gian chuyển động của vật.
\end{tabularx}

\vspace{18pt}

\textbf{Câu 2.} Công thức nào sau đây dùng để tính tốc độ của một vật?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $v = s \cdot t$ & \textbf{B.} $v = s / t$\\[6pt]
\textbf{C.} $v = t / s$ & \textbf{D.} $v = s$
\end{tabularx}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 4: TRẮC NGHIỆM CÂU 3 & CÂU 4
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 4: TRẮC NGHIỆM CÂU 3 & CÂU 4
% ==============================================================================
\fontsize{26}{35}\selectfont
\textbf{Câu 3.} Đơn vị nào sau đây là đơn vị đo tốc độ trong hệ đo lường chính thức của nước ta?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\text{m/s}$ & \textbf{B.} $\text{km/min.}$\\[6pt]
\textbf{C.} $\text{km/s}$ & \textbf{D.} $\text{m/phút}$
\end{tabularx}

\vspace{18pt}

\textbf{Câu 4.} Tốc độ $36\text{ km/h}$ ứng với giá trị nào sau đây?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $3,6\text{ m/s}$ & \textbf{B.} $12\text{ m/s}$\\[6pt]
\textbf{C.} $10\text{ m/s}$ & \textbf{D.} $1\text{ m/s}$
\end{tabularx}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 5: TRẮC NGHIỆM CÂU 5 & CÂU 6
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 5: TRẮC NGHIỆM CÂU 5 & CÂU 6
% ==============================================================================
\fontsize{26}{35}\selectfont
\textbf{Câu 5.} Đơn vị của tốc độ phụ thuộc vào:
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} đơn vị đo chiều dài. & \textbf{B.} đơn vị đo độ dài và đơn vị đo thời gian.\\[6pt]
\textbf{C.} đơn vị đo thời gian. & \textbf{D.} các yếu tố khác.
\end{tabularx}

\vspace{18pt}

\textbf{Câu 6.} Một người đi bộ quãng đường $100\text{ m}$ trong $20\text{ giây}$. Tốc độ của người đó là:
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $5\text{ m/s}$ & \textbf{B.} $20\text{ m/s}$\\[6pt]
\textbf{C.} $100\text{ m/s}$ & \textbf{D.} $0,2\text{ m/s}$
\end{tabularx}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 6: TRẮC NGHIỆM CÂU 7 & CÂU 8
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 6: TRẮC NGHIỆM CÂU 7 & CÂU 8
% ==============================================================================
\fontsize{25}{34}\selectfont
\textbf{Câu 7.} Đồ thị quãng đường – thời gian cho biết:
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Tốc độ đi được. & \textbf{B.} Thời gian đi được.\\[6pt]
\textbf{C.} Quãng đường đi được. & \textbf{D.} Cả tốc độ, thời gian và quãng đường.
\end{tabularx}

\vspace{14pt}

\textbf{Câu 8.} Bạn A chạy quãng đường $60\text{ m}$ hết $10\text{ s}$. Bạn B chạy quãng đường $100\text{ m}$ hết $20\text{ s}$. Nhận xét nào sau đây là đúng?
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Bạn A đi nhanh hơn bạn B. & \textbf{B.} Bạn B đi nhanh hơn bạn A.\\[6pt]
\textbf{C.} Hai bạn đi nhanh như nhau. & \textbf{D.} Không thể so sánh được.
\end{tabularx}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 7: TRẮC NGHIỆM CÂU 9 & CÂU 10
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 7: TRẮC NGHIỆM CÂU 9 & CÂU 10
% ==============================================================================
\fontsize{25}{34}\selectfont
\textbf{Câu 9.} Một ô tô đi $100\text{ m}$ đầu tiên trong $10\text{ s}$, sau đó đi tiếp $200\text{ m}$ trong $20\text{ s}$. Tốc độ chuyển động của ô tô trên cả quãng đường là:
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $10\text{ m/s}$ & \textbf{B.} $15\text{ m/s}$\\[6pt]
\textbf{C.} $30\text{ m/s}$ & \textbf{D.} $100\text{ m/s}$
\end{tabularx}

\vspace{16pt}

\textbf{Câu 10.} Trên đồ thị quãng đường – thời gian, đường thẳng càng dốc thể hiện:
\vspace{4pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Chuyển động nhanh hơn. & \textbf{B.} Chuyển động chậm hơn.\\[6pt]
\textbf{C.} Vật đứng yên. & \textbf{D.} Vật chuyển động đều.
\end{tabularx}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 8: ĐÚNG - SAI CÂU 11 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 8: ĐÚNG - SAI CÂU 11
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}II. TRẮC NGHIỆM ĐÚNG SAI -- CÂU 11}}
\vspace{8pt}

\fontsize{22}{30}\selectfont
\textbf{Câu 11.} Xác định mỗi phát biểu sau là đúng hay sai:

\vspace{12pt}
\textbf{a)} Tốc độ càng lớn thì vật chuyển động càng chậm.\hfill $[\quad]$\\[10pt]
\textbf{b)} Khi vật đứng yên, tốc độ của vật bằng 0.\hfill $[\quad]$\\[10pt]
\textbf{c)} Công thức tính tốc độ là $v = t : s$.\hfill $[\quad]$\\[10pt]
\textbf{d)} Đơn vị tốc độ có thể là $\text{km/h}$ hoặc $\text{m/s}$.\hfill $[\quad]$
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 9: ĐÚNG - SAI CÂU 12 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 9: ĐÚNG - SAI CÂU 12
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}II. TRẮC NGHIỆM ĐÚNG SAI -- CÂU 12}}
\vspace{8pt}

\fontsize{22}{30}\selectfont
\textbf{Câu 12.} Xác định mỗi phát biểu sau là đúng hay sai:

\vspace{12pt}
\textbf{a)} Trên cùng một quãng đường, nếu thời gian chuyển động càng ngắn thì chuyển động đó nhanh hơn.\hfill $[\quad]$\\[10pt]
\textbf{b)} Đồ thị quãng đường – thời gian là đường thẳng đi qua gốc tọa độ.\hfill $[\quad]$\\[10pt]
\textbf{c)} Đồ thị càng dốc thì tốc độ càng nhỏ.\hfill $[\quad]$\\[10pt]
\textbf{d)} Người đi xe máy $60\text{ km}$ trong $1\text{ giờ}$ có tốc độ $60\text{ km/h}$.\hfill $[\quad]$
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 10: TRẢ LỜI NGẮN CÂU 13 & CÂU 14 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 10: TRẢ LỜI NGẮN CÂU 13 & CÂU 14
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}III. TRẮC NGHIỆM TRẢ LỜI NGẮN -- CÂU 13 \& CÂU 14}}
\vspace{8pt}

\fontsize{21}{29}\selectfont
\textbf{Câu 13.} Đơn vị tốc độ chuyển động hợp pháp (trong hệ SI) được kí hiệu là $\dots\dots\dots\dots\dots$

\vspace{14pt}
\textbf{Câu 14.} Đổi các đơn vị sau sang giá trị tương ứng:
\begin{itemize}
  \setlength{\itemsep}{4pt}
  \item a) $1\text{ km/h} = \dots\dots\dots\dots\dots\text{m/s}$; \quad $1\text{ m/s} = \dots\dots\dots\dots\dots\text{km/h}$.
  \item b) $5\text{ m/s} = \dots\dots\dots\dots\dots\text{km/h}$.
  \item c) $45\text{ km/h} = \dots\dots\dots\dots\dots\text{m/s}$.
  \item d) $720\text{ km/h} = \dots\dots\dots\dots\dots\text{m/s} = \dots\dots\dots\dots\dots\text{cm/s}$.
\end{itemize}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 11: TRẢ LỜI NGẮN CÂU 15, 16 & 17 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 11: TRẢ LỜI NGẮN CÂU 15, 16 & 17
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}III. TRẮC NGHIỆM TRẢ LỜI NGẮN -- CÂU 15, 16 \& 17}}
\vspace{10pt}

\fontsize{22}{30}\selectfont
\textbf{Câu 15.} Quãng đường vật đi được đo bằng mét ($\text{m}$), thời gian đo bằng giây ($\text{s}$) thì tốc độ có đơn vị là: $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$

\vspace{16pt}
\textbf{Câu 16.} Nếu vật đi cùng quãng đường nhưng thời gian tăng gấp đôi thì tốc độ chuyển động giảm đi: $\dots\dots\dots\dots\dots$ lần.

\vspace{16pt}
\textbf{Câu 17.} Trong cùng một khoảng thời gian, nếu quãng đường chuyển động $\dots\dots\dots\dots\dots$ hơn thì chuyển động đó nhanh hơn.
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 12: TRẢ LỜI NGẮN CÂU 18, 19 & 20 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 12: TRẢ LỜI NGẮN CÂU 18, 19 & 20
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}III. TRẮC NGHIỆM TRẢ LỜI NGẮN -- CÂU 18, 19 \& 20}}
\vspace{10pt}

\fontsize{22}{30}\selectfont
\textbf{Câu 18.} Một người đi xe đạp trong $3\text{ giờ}$ đi được quãng đường $45\text{ km}$. Tốc độ của người đó là: $\dots\dots\dots\dots\dots\text{km/h}$.

\vspace{16pt}
\textbf{Câu 19.} Một vận động viên chạy $400\text{ m}$ hết $50\text{ giây}$. Tốc độ của vận động viên đó là: $\dots\dots\dots\dots\dots\text{m/s}$.

\vspace{16pt}
\textbf{Câu 20.} Viết công thức tính tốc độ của chuyển động: $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 13: TRẢ LỜI NGẮN CÂU 21, 22, 23 & 24 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 13: TRẢ LỜI NGẮN CÂU 21, 22, 23 & 24
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}III. TRẮC NGHIỆM TRẢ LỜI NGẮN -- CÂU 21, 22, 23 \& 24}}
\vspace{8pt}

\fontsize{21}{29}\selectfont
\textbf{Câu 21.} Trong hệ đo lường chính thức ở nước ta, tốc độ được đo bằng $\text{m/s}$ và $\dots\dots\dots\dots\dots$

\vspace{12pt}
\textbf{Câu 22.} Đại lượng đặc trưng cho mức độ nhanh chậm của chuyển động là $\dots\dots\dots\dots\dots$

\vspace{12pt}
\textbf{Câu 23.} Khi một vật đứng yên thì tốc độ của vật bằng $\dots\dots\dots\dots\dots$

\vspace{12pt}
\textbf{Câu 24.} Dụng cụ đo tốc độ gắn trên các phương tiện giao thông gọi là $\dots\dots\dots\dots\dots$
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 14: TỰ LUẬN CÂU 25 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 14: TỰ LUẬN CÂU 25
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}IV. PHẦN TỰ LUẬN -- CÂU 25}}
\vspace{8pt}

\fontsize{22}{30}\selectfont
\textbf{Đề bài:} Một xe tải chạy trên đoạn đường đầu dài $45\text{ km}$ trong $45\text{ phút}$, sau đó xe tiếp tục chạy thêm $18\text{ km}$ trong $20\text{ phút}$.\\
Tính tốc độ của xe tải trên mỗi đoạn đường và trên cả hai đoạn đường.

\vspace*{2.5cm}
\begin{center}
{\fontsize{20}{26}\selectfont \textit{\color{gray}(Khoảng trống trình chiếu tương tác / Giáo viên và học sinh chữa bài trực tiếp trên bảng)}}
\end{center}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 15: TỰ LUẬN CÂU 26 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 15: TỰ LUẬN CÂU 26
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}IV. PHẦN TỰ LUẬN -- CÂU 26}}
\vspace{8pt}

\fontsize{22}{30}\selectfont
\textbf{Đề bài:} Một học sinh đi xe đạp từ nhà đến trường mất $20\text{ phút}$. Đoạn đường từ nhà đến trường dài $4\text{ km}$.
\begin{enumerate}
  \setlength{\itemsep}{6pt}
  \item[a)] Tính tốc độ của chuyển động theo đơn vị $\text{km/h}$?
  \item[b)] Nêu ý nghĩa của con số tìm được ở câu a?
\end{enumerate}

\vspace*{2.5cm}
\begin{center}
{\fontsize{20}{26}\selectfont \textit{\color{gray}(Khoảng trống trình chiếu tương tác / Giáo viên và học sinh chữa bài trực tiếp trên bảng)}}
\end{center}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 16: TỰ LUẬN CÂU 27 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 16: TỰ LUẬN CÂU 27
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}IV. PHẦN TỰ LUẬN -- CÂU 27}}
\vspace{8pt}

\fontsize{21}{29}\selectfont
\textbf{Đề bài:} Dựa vào bảng ghi số liệu dưới đây về quãng đường và thời gian của một người đi bộ, em hãy vẽ đồ thị quãng đường – thời gian của người này:

\vspace{8pt}
\begin{center}
\begin{tabular}{|l|c|c|c|c|c|c|}
\hline
\textbf{Thời gian $t$ (h)} & 0 & 0,5 & 1,0 & 1,5 & 2,0 & 2,5 \\
\hline
\textbf{Quãng đường $s$ (km)} & 0 & 25 & 50 & 75 & 75 & 100 \\
\hline
\end{tabular}
\end{center}

\vspace*{2.0cm}
\begin{center}
{\fontsize{20}{26}\selectfont \textit{\color{gray}(Khoảng trống trình chiếu tương tác / Vẽ đồ thị trực tiếp trên màn hình)}}
\end{center}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 17: TỰ LUẬN CÂU 28 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 17: TỰ LUẬN CÂU 28
% ==============================================================================
{\fontsize{24}{30}\selectfont \textbf{\color{brandBlue}IV. PHẦN TỰ LUẬN -- CÂU 28}}
\vspace{8pt}

\fontsize{21}{29}\selectfont
\textbf{Đề bài:} Dựa vào các thông tin về quãng đường và thời gian của một người đi xe đạp trong bảng dưới đây, hãy mô tả đặc điểm chuyển động của người đó:

\vspace{8pt}
\begin{center}
\begin{tabular}{|l|c|c|c|c|c|c|c|}
\hline
\textbf{Thời gian $t$ (s)} & 0 & 2 & 4 & 6 & 8 & 10 & 12 \\
\hline
\textbf{Quãng đường $s$ (m)} & 0 & 10 & 20 & 40 & 40 & 60 & 40 \\
\hline
\end{tabular}
\end{center}

\vspace*{2.0cm}
\begin{center}
{\fontsize{20}{26}\selectfont \textit{\color{gray}(Khoảng trống trình chiếu tương tác / Phân tích từng giai đoạn trên màn hình)}}
\end{center}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 18: TỰ LUẬN CÂU 29 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 18: TỰ LUẬN CÂU 29
% ==============================================================================
{\fontsize{23}{28}\selectfont \textbf{\color{brandBlue}IV. PHẦN TỰ LUẬN -- CÂU 29}}
\vspace{4pt}

\begin{minipage}[c]{0.54\linewidth}
\fontsize{19}{26}\selectfont
\textbf{Đề bài:} Dựa vào đồ thị quãng đường – thời gian của một ô tô (hình bên) để trả lời các câu hỏi sau:
\begin{enumerate}
  \setlength{\itemsep}{4pt}
  \item[a)] Trên đoạn đường nào xe chuyển động nhanh hơn? Xác định tốc độ của xe trên mỗi đoạn đường.
  \item[b)] Sau $50\text{ giây}$, xe đi được bao nhiêu mét?
\end{enumerate}

\vspace{14pt}
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\end{minipage}
\hfill
\begin{minipage}[c]{0.44\linewidth}
\centering
\includegraphics[height=11.2cm,keepaspectratio]{San_Pham/images_khtn7_nguyenhoang/dothi_cau29_hd.png}
\end{minipage}
\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE 19: TỰ LUẬN CÂU 30 (KHÔNG CÓ ĐÁP ÁN)
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 19: TỰ LUẬN CÂU 30
% ==============================================================================
{\fontsize{23}{28}\selectfont \textbf{\color{brandBlue}IV. PHẦN TỰ LUẬN -- CÂU 30}}
\vspace{4pt}

\begin{minipage}[c]{0.52\linewidth}
\fontsize{18}{25}\selectfont
\textbf{Đề bài:} Hình bên biểu diễn đồ thị quãng đường – thời gian của một ô tô trên đường phố vào giờ cao điểm trong hành trình dài $4\text{ phút}$.
\begin{enumerate}
  \setlength{\itemsep}{3pt}
  \item[a)] Mô tả các giai đoạn chuyển động của ô tô trên đồ thị.
  \item[b)] Xác định thời gian ô tô đã dừng lại trong hành trình.
  \item[c)] Tốc độ của ô tô trong giai đoạn nào là lớn nhất?
\end{enumerate}

\vspace{10pt}
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\end{minipage}
\hfill
\begin{minipage}[c]{0.46\linewidth}
\centering
\includegraphics[height=11.2cm,keepaspectratio]{San_Pham/images_khtn7_nguyenhoang/dothi_cau30_hd.png}
\end{minipage}

\end{document}
""")

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write("".join(tex_lines))
    print(f"[OK] Đã ghi file LaTeX: {tex_path}")

    # Biên dịch pdflatex 2 lần để ổn định layout
    cmd = ["pdflatex", "-interaction=nonstopmode", "-output-directory", SCRATCH_DIR, tex_path]
    res1 = subprocess.run(cmd, capture_output=True, text=True)
    res2 = subprocess.run(cmd, capture_output=True, text=True)

    pdf_scratch = os.path.join(SCRATCH_DIR, pdf_filename)
    if os.path.exists(pdf_scratch):
        shutil.copy2(pdf_scratch, pdf_dest)
        print(f"[OK] Đã biên dịch và copy PDF thành công: {pdf_dest} ({os.path.getsize(pdf_dest):,} bytes)")
        return pdf_dest
    else:
        print("[ERROR] Không tìm thấy file PDF được tạo. Log:")
        print(res2.stdout[-1500:])
        return None


# ==============================================================================
# HÀM MAIN
# ==============================================================================

def main():
    print("=" * 75)
    print("BẮT ĐẦU TẠO 3 FILE CHO ĐỀ CƯƠNG ÔN TẬP GIỮA KÌ I KHTN 7 (THCS NGUYỄN HOÀNG)")
    print("1. FILE HỌC SINH (.docx)")
    print("2. FILE GIÁO VIÊN (.docx có lời giải chi tiết)")
    print("3. FILE DẠY TRÊN MÀN HÌNH (Slide PDF 16:9 widescreen cho Tivi 65 inch)")
    print("=" * 75)

    f_hs = generate_word_document(mode="student")
    f_gv = generate_word_document(mode="teacher")
    f_screen_docx = generate_word_document(mode="screen")
    f_screen_pdf = generate_tivi_presentation_pdf()

    all_files = [f_hs, f_gv, f_screen_docx, f_screen_pdf]

    # Đồng bộ sang các thư mục OneDrive
    print("\n[+] Đồng bộ các file sang các thư mục OneDrive:")
    for od in TARGET_ONEDRIVE_DIRS:
        if os.path.exists(od):
            for src_path in all_files:
                if src_path and os.path.exists(src_path):
                    dest = os.path.join(od, os.path.basename(src_path))
                    try:
                        shutil.copy2(src_path, dest)
                        print(f"    -> Đã copy {os.path.basename(src_path)} vào: {od}")
                    except Exception as e:
                        print(f"    [!] Cảnh báo file đang mở hoặc bận tại {od}: {e}")

    print("\n" + "=" * 75)
    print("HOÀN THÀNH TẤT CẢ 3 FILE THEO ĐÚNG CHUẨN FORM ĐỀ KIỂM TRA!")
    print("=" * 75)

if __name__ == "__main__":
    main()
