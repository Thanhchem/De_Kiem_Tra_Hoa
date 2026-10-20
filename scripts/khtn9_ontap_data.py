# -*- coding: utf-8 -*-
"""
Dữ liệu và mã nguồn tạo bộ 3 file Word (.docx) Đề Ôn tập Kiểm tra Giữa kì I - KHTN 9 (2026 - 2027)
Theo chuẩn "FORM ĐỀ KIỂM TRA" của Thầy TRẦN VĂN THẠNH:
1. Bản Học sinh (De_Kiem_Tra_KHTN_9_Giua_HK1_HocSinh.docx)
2. Bản Giáo viên (De_Kiem_Tra_KHTN_9_Giua_HK1_GiaoVien.docx)
3. Bản Dạy trên màn hình (De_Kiem_Tra_KHTN_9_Giua_HK1_DayTrenManHinh.docx)
"""

import os
import sys
import re
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

ROOT_DIR = r"c:\Antigravity_Thanh"
OUT_DIR = os.path.join(ROOT_DIR, "San_Pham")
IMG_DNA = os.path.join(OUT_DIR, "images_khtn9", "dna_cau3_part2_hd.png")

ONEDRIVE_DIRS = [
    r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9\KHTN9-DEKIEMTRADINHKI",
    r"C:\Users\Admin\OneDrive\Tệp đính kèm\2026-2027\khtn9"
]

# Bảng màu quy chuẩn
COLOR_HF_RED = RGBColor(211, 47, 47)        # Coral Red (#D32F2F) cho Header & Footer
COLOR_HF_GREEN_HEX = "8BC390"               # Sage Green (#8BC390) cho viền Header & Footer
COLOR_PRIMARY = RGBColor(46, 125, 50)       # Xanh lá đậm (#2E7D32) cho Tiêu đề Phần
COLOR_CORRECT = RGBColor(198, 40, 40)       # Đỏ (#C62828) cho đáp án đúng Bản Giáo Viên
COLOR_SOLUTION = RGBColor(21, 101, 192)     # Xanh lam (#1565C0) cho tiêu đề Lời giải
COLOR_SCREEN_TITLE = RGBColor(106, 27, 154) # Tím đậm (#6A1B9A) cho tiêu đề bản Dạy Màn Hình
COLOR_GRAY = RGBColor(110, 110, 110)
COLOR_LINE_DOT = RGBColor(180, 180, 180)    # Màu nét chấm mờ đẹp
COLOR_BORDER_GREY_HEX = "CCCCCC"            # Xám viền phân cách tiêu đề

# Dữ liệu 24 câu trắc nghiệm nhiều lựa chọn Phần 1
PART_1_QUESTIONS = [
    {
        "id": 1,
        "content": "Hai chiếc máy cẩu cùng nâng một lượng hàng hóa có trọng lượng bằng nhau lên cùng một độ cao, nhưng máy A hoàn thành trong 1 phút, máy B hoàn thành trong 2 phút. Để so sánh khả năng làm việc của hai máy này, học sinh cần đề xuất câu hỏi vấn đề nào sau đây?",
        "options": [
            "A. Máy cẩu nào tiêu thụ nhiều nhiên liệu hơn trong quá trình nâng?",
            "B. Máy cẩu nào thực hiện công việc với tốc độ nhanh hơn?",
            "C. Trọng lượng tổng cộng của hàng hóa là bao nhiêu?",
            "D. Lực kéo của động cơ máy cẩu nào lớn hơn?"
        ],
        "correct": "B",
        "solution": "Vì hai máy cẩu nâng cùng một lượng hàng hóa có trọng lượng bằng nhau lên cùng độ cao nên công thực hiện của hai máy cẩu là như nhau (A~A~ = A~B~ = P.h). Tuy nhiên, máy A hoàn thành công việc trong thời gian ngắn hơn (t~A~ = 1 phút < t~B~ = 2 phút). Đại lượng đặc trưng cho khả năng thực hiện công nhanh hay chậm là công suất (tốc độ thực hiện công P = A/t). Do đó, để so sánh khả năng làm việc của hai máy, học sinh cần đề xuất câu hỏi: *\"Máy cẩu nào thực hiện công việc với tốc độ nhanh hơn?\"* (hoặc so sánh công suất của hai máy cẩu)."
    },
    {
        "id": 2,
        "content": "Khi khảo sát quá trình đẩy một chiếc xe đẩy hàng siêu thị, một học sinh đặt giả thuyết về sự thay đổi của công cơ học. Nếu giữ nguyên quãng đường dịch chuyển nhưng tăng lực đẩy lên gấp đôi (xe chở nặng hơn), phán đoán nào sau đây là chính xác về công cơ học được thực hiện?",
        "options": [
            "A. Công thực hiện sẽ giảm đi một nửa.",
            "B. Công thực hiện không thay đổi vì quãng đường không đổi.",
            "C. Công thực hiện sẽ tăng lên gấp đôi.",
            "D. Công thực hiện sẽ tăng lên gấp bốn lần."
        ],
        "correct": "C",
        "solution": "Công cơ học được xác định theo công thức: A = F.s. Khi quãng đường dịch chuyển s không đổi, công cơ học A tỉ lệ thuận với lực tác dụng F. Do đó, khi tăng lực đẩy lên gấp đôi (F' = 2F) thì công cơ học thực hiện cũng tăng lên gấp đôi (A' = 2A)."
    },
    {
        "id": 3,
        "content": "Hiện tượng con sinh ra có đặc điểm khác nhau và khác với bố mẹ được gọi là",
        "options": [
            "A. di truyền.",
            "B. biến dị.",
            "C. nguyên phân.",
            "D. thụ tinh."
        ],
        "correct": "B",
        "solution": "Theo định nghĩa sinh học (SGK KHTN 9): Biến dị là hiện tượng con sinh ra mang các đặc điểm sai khác nhau và khác với bố mẹ, tổ tiên. (Còn di truyền là hiện tượng truyền đạt các tính trạng của bố mẹ cho con cái)."
    },
    {
        "id": 4,
        "content": "Mendel cho rằng đơn vị quy định sự di truyền của một tính trạng tồn tại thành từng cặp trong nhân tế bào và chúng không pha trộn vào nhau, gọi là cặp",
        "options": [
            "A. gene.",
            "B. allele.",
            "C. nhân tố di truyền.",
            "D. tính trạng."
        ],
        "correct": "C",
        "solution": "Theo Menđen: quy định tính trạng của cơ thể là các \"nhân tố di truyền\" (sau này di truyền học hiện đại gọi là gene/allele). Các nhân tố di truyền tồn tại thành từng cặp trong nhân tế bào, không pha trộn vào nhau và phân li đồng đều về các giao tử."
    },
    {
        "id": 5,
        "content": "Dụng cụ nào sau đây dùng để phân tán nhiệt khi đốt?",
        "options": [
            "A. Lưới tản nhiệt.",
            "B. Bát sứ.",
            "C. Bình cầu.",
            "D. Phễu chiết."
        ],
        "correct": "A",
        "solution": "Lưới tản nhiệt (lưới kim loại có phủ amiăng hoặc gốm chịu nhiệt) được đặt giữa ngọn lửa đèn cồn và đáy dụng cụ thuỷ tinh (cốc thuỷ tinh, bình tam giác, bình cầu) nhằm giúp phân tán nhiệt lượng đồng đều, tránh hiện tượng nứt vỡ do sốc nhiệt cục bộ."
    },
    {
        "id": 6,
        "content": "Hợp chất hữu cơ là",
        "options": [
            "A. hợp chất của oxygen với một nguyên tố hóa học khác.",
            "B. các hợp chất chứa carbon, hydrogen và oxygen.",
            "C. những hợp chất chứa carbon và hydrogen.",
            "D. hợp chất của carbon (trừ CO, CO~2~, H~2~CO~3~, muối carbonate kim loại…)."
        ],
        "correct": "D",
        "solution": "Hợp chất hữu cơ là hợp chất của carbon (trừ một số hợp chất vô cơ đơn giản như carbon monoxide CO, carbon dioxide CO~2~, acid carbonic H~2~CO~3~, các muối carbonate kim loại, các muối cyanide, carbide...)."
    },
    {
        "id": 7,
        "content": "Phát biểu nào sau đây đúng?",
        "options": [
            "A. Alkane là các hydrocarbon mà phân tử chỉ có liên kết đơn.",
            "B. Trong CTPT của alkane, khi tăng thêm một nguyên tử C thì số nguyên tử H tăng lên 2.",
            "C. Phân tử alkane chỉ có mạch carbon không nhánh.",
            "D. Khi alkane cháy, số mol H~2~O sinh ra nhỏ hơn hoặc bằng số mol CO~2~."
        ],
        "correct": "B",
        "solution": "• A sai: Cycloalkane (như xiclopropan, xiclobutan) cũng là hydrocarbon chỉ có liên kết đơn nhưng là hydrocarbon mạch vòng. Alkane là hydrocarbon no mạch hở.\n• B đúng: Alkane có CTTQ là C~n~H~2n+2~ (n ≥ 1). Khi tăng thêm 1 nguyên tử C (n → n+1) thì số nguyên tử H là 2(n+1)+2 = 2n+4, tức tăng thêm 2 nguyên tử H (đặc trưng của dãy đồng đẳng, hơn kém nhau nhóm -CH~2~-).\n• C sai: Alkane từ 4 nguyên tử C trở lên (C~4~H~10~) có đồng phân mạch phân nhánh.\n• D sai: Khi alkane cháy, số mol H~2~O luôn lớn hơn số mol CO~2~ (n~H2O~ = n+1 > n~CO2~ = n)."
    },
    {
        "id": 8,
        "content": "Hợp chất nào sau đây là hydrocarbon?",
        "options": [
            "A. Saccharose.",
            "B. Glucose.",
            "C. Methane.",
            "D. Ethylic alcohol."
        ],
        "correct": "C",
        "solution": "Hydrocarbon là hợp chất hữu cơ mà trong phân tử chỉ chứa 2 nguyên tố là Carbon (C) và Hydrogen (H). Methane có công thức phân tử là CH~4~, thuộc loại alkane (hydrocarbon). Các chất còn lại (Saccharose C~12~H~22~O~11~, Glucose C~6~H~12~O~6~, Ethylic alcohol C~2~H~5~OH) là dẫn xuất của hydrocarbon do chứa thêm nguyên tố oxygen."
    },
    {
        "id": 9,
        "content": "Dụng cụ nào sau đây được sử dụng để đo cường độ dòng điện trong mạch điện?",
        "options": [
            "A. Vôn kế.",
            "B. Ampe kế.",
            "C. Nhiệt kế.",
            "D. Lực kế."
        ],
        "correct": "B",
        "solution": "Ampe kế là dụng cụ chuyên dụng dùng để đo cường độ dòng điện trong mạch điện (đơn vị là Ampe - A). Vôn kế dùng đo hiệu điện thế; Nhiệt kế đo nhiệt độ; Lực kế đo lực."
    },
    {
        "id": 10,
        "content": "Động năng là năng lượng mà vật có được do chuyển động. Biểu thức nào sau đây thể hiện đúng công thức tính động năng của một vật?",
        "options": [
            "A. W~đ~ = mv^2^.",
            "B. W~đ~ = 1/2mv^2^.",
            "C. W~đ~ = 1/2mv.",
            "D. W~đ~ = Ph."
        ],
        "correct": "B",
        "solution": "Động năng của một vật có khối lượng m chuyển động với tốc độ v được xác định theo công thức: W~đ~ = 1/2mv^2^."
    },
    {
        "id": 11,
        "content": "Công suất là đại lượng đặc trưng cho tốc độ thực hiện công và được tính bằng công thực hiện được trong một đơn vị thời gian (P = A/t). Dựa vào hình thức biểu đạt qua công thức này, đơn vị của công suất được viết dưới dạng tương đương nào sau đây?",
        "options": [
            "A. N.m.",
            "B. J.s.",
            "C. J/s.",
            "D. N/s."
        ],
        "correct": "C",
        "solution": "Từ công thức tính công suất P = A/t, trong đó công A có đơn vị là Joule (J) và thời gian t có đơn vị là giây (s), đơn vị của công suất là Watt (W), tương đương với Joule trên giây (J/s)."
    },
    {
        "id": 12,
        "content": "Chiết suất của một môi trường trong suốt được xác định bằng biểu thức nào sau đây? (Với c là tốc độ ánh sáng trong chân không hoặc không khí, v là tốc độ ánh sáng trong môi trường đó).",
        "options": [
            "A. n = c/v.",
            "B. n = c - v.",
            "C. n = cv.",
            "D. n = v/c."
        ],
        "correct": "A",
        "solution": "Chiết suất tuyệt đối của một môi trường trong suốt được xác định bằng tỉ số giữa tốc độ ánh sáng trong chân không c với tốc độ ánh sáng trong môi trường đó v: n = c/v (n ≥ 1)."
    },
    {
        "id": 13,
        "content": "Cần cẩu A nâng một thùng hàng nặng 1,5 tấn lên cao 3m trong thời gian 15 giây. Cần cẩu B nâng một thùng hàng nặng 2 tấn lên cao 2m trong thời gian 20 giây. Nhận xét nào sau đây là đúng về công suất của hai cần cẩu?",
        "options": [
            "A. Hai cần cẩu có công suất như nhau.",
            "B. Cần cẩu A có công suất nhỏ hơn.",
            "C. Cần cẩu A có công suất lớn hơn.",
            "D. Không thể so sánh được."
        ],
        "correct": "C",
        "solution": "• Cần cẩu A: m~A~ = 1,5 tấn = 1500 kg ⇒ P~A~ = 10.m~A~ = 15000 N.\n  Công thực hiện: A~A~ = P~A~.h~A~ = 15000 . 3 = 45000 J.\n  Công suất: P~1~ = A~A~/t~A~ = 45000 / 15 = 3000 W.\n• Cần cẩu B: m~B~ = 2 tấn = 2000 kg ⇒ P~B~ = 10.m~B~ = 20000 N.\n  Công thực hiện: A~B~ = P~B~.h~B~ = 20000 . 2 = 40000 J.\n  Công suất: P~2~ = A~B~/t~B~ = 40000 / 20 = 2000 W.\nVì 3000 W > 2000 W nên Cần cẩu A có công suất lớn hơn cần cẩu B."
    },
    {
        "id": 14,
        "content": "Hiện tượng con sinh ra mang những đặc điểm giống bố mẹ được gọi là",
        "options": [
            "A. biến dị.",
            "B. di truyền.",
            "C. đột biến.",
            "D. tiến hóa."
        ],
        "correct": "B",
        "solution": "Di truyền là hiện tượng truyền đạt các tính trạng, đặc điểm của bố mẹ, tổ tiên cho các thế hệ con cháu."
    },
    {
        "id": 15,
        "content": "Theo quan điểm của Mendel, mỗi tính trạng của cơ thể do yếu tố nào quy định?",
        "options": [
            "A. Một cặp nhân tố di truyền.",
            "B. Một nhân tố di truyền duy nhất.",
            "C. Các yếu tố môi trường.",
            "D. Sự hòa trộn các tính trạng của bố mẹ."
        ],
        "correct": "A",
        "solution": "Theo quan điểm của Menđen, trong tế bào sinh dưỡng, mỗi tính trạng của cơ thể do một cặp nhân tố di truyền (gene) quy định (ví dụ cặp gen AA, Aa hoặc aa)."
    },
    {
        "id": 16,
        "content": "Các hóa chất cần được bảo quản như thế nào?",
        "options": [
            "A. Trong chai hoặc lọ nhựa, để ở chỗ tối.",
            "B. Trong chai hoặc lọ, được dán nhãn ghi thông tin về hóa chất.",
            "C. Trong chai hoặc lọ, có nắp đậy.",
            "D. Trong chai hoặc lọ, có nắp đậy và được dán nhãn ghi thông tin về hóa chất."
        ],
        "correct": "D",
        "solution": "Để đảm bảo an toàn phòng thí nghiệm và tránh hóa chất bị biến chất, bay hơi hay gây nhầm lẫn nguy hiểm, hóa chất cần được bảo quản trong chai hoặc lọ có nắp đậy kín và được dán nhãn ghi đầy đủ thông tin (tên hóa chất, công thức phân tử, độ tinh khiết, cảnh báo an toàn)."
    },
    {
        "id": 17,
        "content": "Hợp chất nào sau đây thuộc loại dẫn xuất của hydrocarbon?",
        "options": [
            "A. C~2~H~4~O~2~.",
            "B. CaCO~3~.",
            "C. NaHCO~3~.",
            "D. C~3~H~4~."
        ],
        "correct": "A",
        "solution": "• C~2~H~4~O~2~ là hợp chất hữu cơ chứa các nguyên tố C, H và O nên thuộc loại dẫn xuất của hydrocarbon.\n• C~3~H~4~ chỉ chứa C và H nên là hydrocarbon.\n• CaCO~3~ và NaHCO~3~ là các hợp chất vô cơ (muối carbonate)."
    },
    {
        "id": 18,
        "content": "Công thức phân tử nào sau đây không phải là của một alkane?",
        "options": [
            "A. C~3~H~8~.",
            "B. C~2~H~6~.",
            "C. C~5~H~10~.",
            "D. C~6~H~14~."
        ],
        "correct": "C",
        "solution": "Alkane có công thức phân tử chung là C~n~H~2n+2~ (n ≥ 1).\n• Với n = 3: C~3~H~8~ (propane - alkane).\n• Với n = 2: C~2~H~6~ (ethane - alkane).\n• Với n = 5: C~5~H~12~ mới là alkane; do đó C~5~H~10~ không phải là alkane (nó thuộc loại alkene hoặc cycloalkane).\n• Với n = 6: C~6~H~14~ (hexane - alkane)."
    },
    {
        "id": 19,
        "content": "Khi các alkane cháy hoàn toàn trong khí oxygen, các sản phẩm tạo ra thường là gì?",
        "options": [
            "A. CO~2~ và H~2~O.",
            "B. CO và H~2~O.",
            "C. CO~2~ và H~2~.",
            "D. CH~4~ và O~2~."
        ],
        "correct": "A",
        "solution": "Khi các alkane bị đốt cháy hoàn toàn trong khí oxygen dư, phản ứng tạo thành hai sản phẩm chính là carbon dioxide (CO~2~) và nước (H~2~O):\nC~n~H~2n+2~ + (3n+1)/2 O~2~ -(t°)-> n CO~2~ + (n+1) H~2~O."
    },
    {
        "id": 20,
        "content": "Đồng hồ đo điện đa năng cho phép đo được các đại lượng khác nhau như",
        "options": [
            "A. cường độ dòng điện.",
            "B. cường độ ánh sáng.",
            "C. hiệu điện thế.",
            "D. cường độ dòng điện, hiệu điện thế, điện trở, ..."
        ],
        "correct": "D",
        "solution": "Đồng hồ đo điện đa năng (đồng hồ vạn năng VOM) là thiết bị đo điện tử tích hợp nhiều chức năng đo: cường độ dòng điện (I), hiệu điện thế (U), điện trở (R), kiểm tra thông mạch, điện dung..."
    },
    {
        "id": 21,
        "content": "Công thức tính thế năng của một vật là",
        "options": [
            "A. W~t~ = 10.m",
            "B. W~t~ = 1/2 .m.v^2^",
            "C. W~t~ = P.h",
            "D. W~t~ = F.s"
        ],
        "correct": "C",
        "solution": "Thế năng trọng trường của một vật có trọng lượng P ở độ cao h so với mốc chọn tính thế năng được tính theo công thức: W~t~ = P.h = 10.m.h."
    },
    {
        "id": 22,
        "content": "Cơ năng của một vật được tính bằng",
        "options": [
            "A. tích của động năng và thế năng của vật.",
            "B. hiệu giữa động năng và thế năng của vật.",
            "C. thương của động năng và thế năng của vật.",
            "D. tổng động năng và thế năng của vật."
        ],
        "correct": "D",
        "solution": "Cơ năng của một vật bằng tổng động năng và thế năng của vật đó: W = W~đ~ + W~t~."
    },
    {
        "id": 23,
        "content": "Chiết suất của một môi trường trong suốt có giá trị bằng",
        "options": [
            "A. tỉ số giữa tốc độ ánh sáng trong môi trường đó với tốc độ ánh sáng trong chân không.",
            "B. tỉ số giữa tốc độ ánh sáng trong chân không với tốc độ ánh sáng trong môi trường đó.",
            "C. tích của tốc độ ánh sáng trong không khí với tốc độ ánh sáng trong môi trường đó.",
            "D. tổng của tốc độ ánh sáng trong không khí và tốc độ ánh sáng trong môi trường đó."
        ],
        "correct": "B",
        "solution": "Chiết suất tuyệt đối n của một môi trường trong suốt bằng tỉ số giữa tốc độ ánh sáng trong chân không (c) với tốc độ ánh sáng trong môi trường đó (v): n = c/v."
    },
    {
        "id": 24,
        "content": "Một lực sĩ nâng một quả tạ khối lượng 120 kg lên cao 2m theo phương thẳng đứng. Công của lực nâng trong trường hợp này bằng",
        "options": [
            "A. 240 J.",
            "B. 2400 J.",
            "C. 60J.",
            "D. 600J."
        ],
        "correct": "B",
        "solution": "Trọng lượng của quả tạ là: P = 10.m = 10 . 120 = 1200 N.\nĐể nâng quả tạ lên đều theo phương thẳng đứng, lực nâng tối thiểu cần tác dụng là: F = P = 1200 N.\nCông của lực nâng thực hiện là: A = F.h = 1200 . 2 = 2400 J."
    }
]

# Dữ liệu Phần 2: Trắc nghiệm Đúng - Sai (4 câu)
PART_2_QUESTIONS = [
    {
        "id": 1,
        "content": "Trong tiết học thực hành, giáo viên yêu cầu học sinh tìm hiểu về tính chất ảnh tạo bởi thấu kính thông qua việc bố trí thí nghiệm với ngọn nến, thấu kính và màn chắn.",
        "statements": [
            {
                "id": "a",
                "text": "Ảnh thật là ảnh có thể hứng được trên màn chắn đặt phía sau thấu kính.",
                "correct": True,
                "explain": "Đúng. Ảnh thật tạo bởi chùm tia ló hội tụ và có thể hứng được rõ nét trên màn chắn đặt phía sau thấu kính."
            },
            {
                "id": "b",
                "text": "Tiêu cự của thấu kính được quy ước là khoảng cách từ quang tâm đến tiêu điểm chính của thấu kính.",
                "correct": True,
                "explain": "Đúng. Tiêu cự f là khoảng cách từ quang tâm O đến tiêu điểm chính F hoặc F' (f = OF = OF')."
            },
            {
                "id": "c",
                "text": "Khi chiếu tia sáng qua thấu kính phân kì, thấy chùm tia ló loe rộng ra, học sinh cho rằng: \"Thấu kính phân kì có thể ứng dụng để làm kính chiếu hậu cho ô tô nhằm mở rộng vùng quan sát.\"",
                "correct": False,
                "explain": "Sai. Kính chiếu hậu ô tô phải dùng GƯƠNG CẦU LỒI (gương phản xạ ánh sáng) để người lái quan sát phía sau. Thấu kính là dụng cụ quang học cho ánh sáng khúc xạ truyền qua, không dùng làm gương chiếu hậu ô tô."
            },
            {
                "id": "d",
                "text": "Thấu kính hội tụ lại làm chùm tia tới song song hội tụ tại một điểm, ta có thể phân tích vấn đề bằng cách đề xuất mô hình coi thấu kính này như một tập hợp các lăng kính nhỏ có đáy hướng xa trục chính.",
                "correct": False,
                "explain": "Sai. Lăng kính luôn làm tia sáng lệch về phía đáy của nó. Vì thấu kính hội tụ bẻ gãy tia sáng về phía trục chính (hội tụ tại một điểm) nên mô hình tương đương phải coi thấu kính hội tụ như tập hợp các lăng kính nhỏ có ĐÁY HƯỚNG VỀ PHÍA TRỤC CHÍNH (chứ không phải hướng xa trục chính)."
            }
        ],
        "explanation": "• Ý a ĐÚNG: Ảnh thật luôn hứng được trên màn chắn đặt sau thấu kính.\n• Ý b ĐÚNG: Tiêu cự f = OF = OF' là khoảng cách từ quang tâm đến tiêu điểm chính.\n• Ý c SAI: Kính chiếu hậu ô tô bắt buộc là gương cầu lồi để phản xạ ánh sáng từ phía sau vào mắt người lái; thấu kính phân kì cho ánh sáng truyền qua nên không thể dùng làm kính chiếu hậu xe ô tô.\n• Ý d SAI: Mô hình tương đương của thấu kính hội tụ là tập hợp các lăng kính nhỏ có đáy hướng về phía trục chính (lệch tia sáng vào trong)."
    },
    {
        "id": 2,
        "content": "Khi tìm hiểu về thấu kính, em hãy cho biết trong các nhận định sau, nhận định nào đúng, nhận định nào sai.",
        "statements": [
            {
                "id": "a",
                "text": "Điểm chính giữa thấu kính mà mọi tia sáng đi qua đều truyền thẳng gọi là quang tâm của thấu kính.",
                "correct": True,
                "explain": "Đúng. Đây là định nghĩa chuẩn về quang tâm O của thấu kính."
            },
            {
                "id": "b",
                "text": "Khoảng cách giữa hai tiêu điểm của thấu kính gọi là tiêu cự.",
                "correct": False,
                "explain": "Sai. Tiêu cự là khoảng cách từ quang tâm đến tiêu điểm (f = OF = OF'). Khoảng cách giữa hai tiêu điểm chính F và F' bằng 2f (bằng 2 lần tiêu cự)."
            },
            {
                "id": "c",
                "text": "Tia tới đi qua tiêu điểm của thấu kính hội tụ thì tia ló song song với trục chính.",
                "correct": True,
                "explain": "Đúng. Đây là một trong ba đường truyền đặc biệt của tia sáng qua thấu kính hội tụ."
            },
            {
                "id": "d",
                "text": "Các tia sáng đi qua tiêu điểm đều song song với trục chính.",
                "correct": False,
                "explain": "Sai. Chỉ có tia ló sau khi truyền qua thấu kính hội tụ (ứng với tia tới đi qua tiêu điểm vật) mới song song với trục chính, bản thân tia sáng tự nó đi qua tiêu điểm không thể song song với trục chính."
            }
        ],
        "explanation": "• Ý a ĐÚNG: Quang tâm O là điểm chính giữa thấu kính mà mọi tia sáng đi qua đều truyền thẳng.\n• Ý b SAI: Tiêu cự là khoảng cách từ quang tâm đến tiêu điểm chính (f = OF); khoảng cách giữa hai tiêu điểm chính là FF' = 2f.\n• Ý c ĐÚNG: Tính chất tia sáng đặc biệt: tia tới qua tiêu điểm thì tia ló song song với trục chính.\n• Ý d SAI: Khẳng định sai về mặt quang hình học; chỉ có tia ló sau thấu kính hội tụ mới song song trục chính."
    },
    {
        "id": 3,
        "content": "Cho hình ảnh về cấu trúc một đoạn phân tử DNA. Hãy xác định mỗi nhận định dưới đây về DNA và gene là Đúng hay Sai.",
        "has_image": True,
        "image_path": IMG_DNA,
        "statements": [
            {
                "id": "a",
                "text": "DNA có chức năng lưu giữ, bảo quản và truyền đạt thông tin di truyền.",
                "correct": True,
                "explain": "Đúng. DNA là vật chất mang thông tin di truyền ở cấp độ phân tử."
            },
            {
                "id": "b",
                "text": "Gene mang thông tin di truyền quy định việc tổng hợp một sản phẩm xác định, qua đó góp phần quy định tính trạng của cơ thể.",
                "correct": True,
                "explain": "Đúng. Gene là một đoạn phân tử DNA mang thông tin mã hóa một chuỗi polypeptide hoặc phân tử RNA."
            },
            {
                "id": "c",
                "text": "Một đoạn DNA có 20 cặp nucleotide, trong đó có 8 cặp A–T thì số cặp G–C của đoạn DNA đó là 12 cặp.",
                "correct": True,
                "explain": "Đúng. Số cặp nucleotide = số cặp A-T + số cặp G-C ⇒ Số cặp G-C = 20 - 8 = 12 cặp."
            },
            {
                "id": "d",
                "text": "Một đoạn DNA có trình tự nucleotide trên một mạch là A–T–G–C thì trình tự nucleotide trên mạch bổ sung tương ứng là T–A–G–C.",
                "correct": False,
                "explain": "Sai. Theo NTBS: A liên kết T, T liên kết A, G liên kết C, C liên kết G. Do đó, mạch bổ sung phải có trình tự tương ứng là T–A–C–G (chứ không phải T–A–G–C)."
            }
        ],
        "explanation": "• Ý a ĐÚNG: Chức năng chuẩn xác của DNA là lưu giữ, bảo quản và truyền đạt thông tin di truyền.\n• Ý b ĐÚNG: Gene mang thông tin tổng hợp sản phẩm xác định quy định tính trạng của cơ thể.\n• Ý c ĐÚNG: Đoạn DNA có 20 cặp nucleotide, gồm 8 cặp A–T thì có 20 - 8 = 12 cặp G–C.\n• Ý d SAI: Trình tự mạch bổ sung tương ứng với A–T–G–C phải là T–A–C–G."
    },
    {
        "id": 4,
        "content": "Hai nhà khoa học James Watson (nhà sinh vật học phân tử Hoa Kỳ) và Francis Crick (nhà sinh vật học, vật lý học phân tử người Anh) đã tìm ra mô hình của DNA vào năm 1953. Đây được xem là công trình quan trọng nhất trong kỷ nguyên hiện đại. Năm 1962, hai nhà khoa học trên được trao giải Nobel cho việc tìm ra cấu trúc DNA. Trong các nhận định về phân tử DNA sau, nhận định nào đúng(Đ), nhận định nào sai (S)?",
        "statements": [
            {
                "id": "a",
                "text": "DNA là nơi lưu giữ, bảo quản thông tin di truyền.",
                "correct": True,
                "explain": "Đúng. Cấu trúc chuỗi xoắn kép bền vững của DNA giúp lưu giữ và bảo quản thông tin di truyền."
            },
            {
                "id": "b",
                "text": "Phân tử DNA không có khả năng tự nhân đôi (tái bản).",
                "correct": False,
                "explain": "Sai. DNA có khả năng tự nhân đôi theo nguyên tắc bán bảo toàn và nguyên tắc bổ sung trong pha S kì trung gian."
            },
            {
                "id": "c",
                "text": "DNA là đại phân tử cấu trúc từ các đơn phân adenine (A), thymine (T), cytosine (C) và guanine (G).",
                "correct": True,
                "explain": "Đúng. 4 loại đơn phân cấu tạo nên DNA là A, T, C, G."
            },
            {
                "id": "d",
                "text": "Giữa hai mạch đơn, các nucleotide liên kết với nhau bằng liên kết hóa trị theo nguyên tắc bổ sung.",
                "correct": False,
                "explain": "Sai. Giữa hai mạch đơn, các base liên kết với nhau bằng LIÊN KẾT HYDROGEN (A-T bằng 2 liên kết, G-C bằng 3 liên kết) theo nguyên tắc bổ sung. Liên kết hoá trị (phosphodiester) liên kết các nucleotide trên cùng 1 mạch đơn."
            }
        ],
        "explanation": "• Ý a ĐÚNG: DNA lưu giữ và bảo quản thông tin di truyền.\n• Ý b SAI: DNA có đặc tính tự nhân đôi (tái bản) chính xác.\n• Ý c ĐÚNG: DNA cấu tạo từ 4 loại nucleotide đơn phân là A, T, C, G.\n• Ý d SAI: Giữa hai mạch đơn, các nucleotide liên kết với nhau bằng liên kết hydrogen theo NTBS, không phải liên kết hóa trị."
    }
]

# Dữ liệu Phần 3: Trắc nghiệm trả lời ngắn (8 câu)
PART_3_QUESTIONS = [
    {
        "id": 1,
        "content": "Để hạn chế việc chặt phá rừng lấy củi và hướng tới phát triển bền vững, một trang trại đã lắp đặt hệ thống hầm biogas xử lý chất thải chăn nuôi để lấy khí đốt. Thành phần chính của khí biogas là methane (CH~4~ - một loại alkane). Biết rằng khi đốt cháy hoàn toàn 1 mol khí methane sẽ toả ra lượng nhiệt là 890kJ. Trong một ngày, trang trại này cần sử dụng lượng nhiệt là 44500kJ để đun nấu và sưởi ấm cho vật nuôi. Khối lượng khí methane tối thiểu mà hầm biogas cần sinh ra trong một ngày để cung cấp đủ lượng nhiệt trên là … gam. (Cho nguyên tử khối: C = 12, H = 1)",
        "correct": "800",
        "solution": "Số mol khí methane (CH~4~) cần đốt cháy để cung cấp đủ 44500 kJ nhiệt lượng:\n  n~CH4~ = 44500 / 890 = 50 (mol).\nKhối lượng mol của methane: M~CH4~ = 12 + 4.1 = 16 (g/mol).\nKhối lượng khí methane tối thiểu hầm biogas cần sinh ra trong một ngày:\n  m~CH4~ = n . M = 50 . 16 = 800 (gam)."
    },
    {
        "id": 2,
        "content": "Một quả bóng khối lượng 0,2kg được ném thẳng đứng lên cao từ mặt đất với động năng ban đầu là 40J. Bỏ qua lực cản của không khí. Do tác dụng của trọng lực, toàn bộ động năng ban đầu sẽ chuyển hóa hoàn toàn thành thế năng tại điểm cao nhất. Độ cao cực đại mà quả bóng đạt được là………m.",
        "correct": "20",
        "solution": "Chọn mốc tính thế năng tại mặt đất.\nCơ năng ban đầu của quả bóng: W = W~đ0~ = 40 J.\nTại điểm cao nhất (v = 0), động năng bằng 0, toàn bộ động năng chuyển thành thế năng cực đại:\n  W~t,max~ = W~đ0~ = 40 J.\nMặt khác: W~t,max~ = P . h~max~ = 10.m . h~max~.\n⇒ h~max~ = W~t,max~ / (10.m) = 40 / (10 . 0,2) = 40 / 2 = 20 (m)."
    },
    {
        "id": 3,
        "content": "Một người thợ xách một xô vữa có trọng lượng 150N di chuyển đều theo phương ngang một quãng đường 10m. Một học sinh nhận định: \"Công của người thợ thực hiện để xô vữa di chuyển ngang là A=F.s=150.10=1500J\". Dựa vào kiến thức về điều kiện sinh công cơ học, công cơ học do trọng lực thực hiện trong trường hợp này bằng….J.",
        "correct": "0",
        "solution": "Trọng lực có phương thẳng đứng, chiều hướng từ trên xuống dưới.\nQuãng đường dịch chuyển của xô vữa theo phương nằm ngang (vuông góc với phương của trọng lực).\nKhi lực tác dụng vuông góc với phương dịch chuyển của vật thì công của lực đó bằng 0.\nDo đó, công cơ học do trọng lực thực hiện bằng 0 J."
    },
    {
        "id": 4,
        "content": "Propane (C~3~H~8~) là một trong những thành phần của khí hoá lỏng (LPG) được sử dụng phổ biến làm nhiên liệu đun nấu. Biết nhiệt lượng toả ra khi đốt cháy hoàn toàn 1 mol propane là 2220kJ. Nếu đốt cháy hoàn toàn 7,437 lít khí propane ở đkc thì nhiệt lượng toả ra khi quá trình đốt cháy trên là…. kJ.",
        "correct": "666",
        "solution": "Ở điều kiện chuẩn (đkc: 25 °C, 1 bar), thể tích của 1 mol khí là 24,79 lít.\nSố mol propane (C~3~H~8~) đốt cháy:\n  n~C3H8~ = V / 24,79 = 7,437 / 24,79 = 0,3 (mol).\nNhiệt lượng tỏa ra khi đốt cháy hoàn toàn 0,3 mol propane:\n  Q = 0,3 . 2220 = 666 (kJ)."
    },
    {
        "id": 5,
        "content": "Một vật có khối lượng 0,5 kg được thả rơi tự do ở độ cao 6,25 m. Nếu bỏ qua sức cản của không khí thì động năng ngay khi vật vừa chạm đất (làm tròn đến phần mười) là … J.",
        "correct": "31,3",
        "solution": "Chọn mốc thế năng tại mặt đất. Bỏ qua lực cản không khí, cơ năng của vật bảo toàn.\nThế năng ban đầu ở độ cao h = 6,25 m:\n  W~t~ = P . h = 10.m . h = 10 . 0,5 . 6,25 = 31,25 J.\nNgay khi vừa chạm đất, thế năng bằng 0, toàn bộ thế năng chuyển hóa thành động năng:\n  W~đ~ = W~t~ = 31,25 J ≈ 31,3 J.\n(Ghi chú: Nếu tính theo gia tốc trọng trường g = 9,8 m/s^2^: W~đ~ = mgh = 0,5 . 9,8 . 6,25 = 30,625 J ≈ 30,6 J)."
    },
    {
        "id": 6,
        "content": "Đầu tàu xe lửa tác dụng vào các toa tàu một lực 2500N làm cho đoàn tàu chuyển động với tốc độ 65 km/h. Công suất của đầu tàu tính theo đơn vị kW (làm tròn đến phần mười) là …",
        "correct": "45,1",
        "solution": "Đổi tốc độ đoàn tàu sang m/s:\n  v = 65 km/h = 65 / 3,6 m/s ≈ 18,056 m/s.\nCông suất của đầu tàu:\n  P = F . v = 2500 . (65 / 3,6) ≈ 45138,89 W ≈ 45,139 kW.\nLàm tròn đến chữ số phần mười: 45,1 kW."
    },
    {
        "id": 7,
        "content": "Cho các đặc điểm về đường đi của tia sáng qua thấu kính được quy ước đánh số như sau: Tia tới đi qua quang tâm(1); tia ló đi qua tiêu điểm(2); tia tới song song với trục chính(3); tia ló truyền thẳng theo phương của tia tới(4). Để xác định vị trí ảnh của một điểm sáng qua thấu kính bằng cách vẽ hình, ta sử dụng giao điểm của các tia sáng nào? (Điền số thứ tự, ví dụ 1,2)",
        "correct": "2, 4",
        "solution": "Để xác định vị trí ảnh của một điểm sáng qua thấu kính, ta vẽ hai tia sáng đặc biệt xuất phát từ điểm sáng truyền qua thấu kính:\n- Tia tới qua quang tâm (1) cho tia ló truyền thẳng theo phương tia tới (4).\n- Tia tới song song với trục chính (3) cho tia ló đi qua tiêu điểm (2).\nẢnh của điểm sáng là giao điểm của hai TIA LÓ sau khi qua thấu kính, tức là giao điểm của tia ló (2) và tia ló (4).\nDo đó điền: 2, 4 (hoặc 2,4)."
    },
    {
        "id": 8,
        "content": "Trong một thí nghiệm với thấu kính hội tụ có tiêu cự lý thuyết f = 15cm, học sinh bố trí đặt vật cách thấu kính d = 30cm để thu ảnh thật trên màn. Do đo đạc sai sót, học sinh ghi nhận khoảng cách từ ảnh đến thấu kính là d' = 16 cm. Dựa vào công thức thấu kính 1/d + 1/d' = 1/f, hãy tính lại giá trị d' chính xác (theo cm) để chỉ ra điểm sai trong số liệu của học sinh.",
        "correct": "30",
        "solution": "Áp dụng công thức thấu kính: 1/d + 1/d' = 1/f\nThay d = 30 cm, f = 15 cm:\n  1/d' = 1/f - 1/d = 1/15 - 1/30 = (2 - 1)/30 = 1/30\n⇒ d' = 30 (cm).\n(Nhận xét: Khi vật đặt tại d = 2f = 30 cm, thấu kính hội tụ cho ảnh thật ngược chiều bằng vật ở khoảng cách d' = 2f = 30 cm. Số liệu đo 16 cm của học sinh là sai sót lớn khi lấy nét trên màn chắn)."
    }
]

# Dữ liệu Phần II: Tự luận (8 câu)
ESSAY_QUESTIONS = [
    {
        "id": 1,
        "content": "Viết công thức cấu tạo và gọi tên các alkane từ C~1~ đến C~4~.",
        "solution": "• C~1~ (CH~4~): Methane.\n  - CTCT thu gọn: CH~4~\n• C~2~ (C~2~H~6~): Ethane.\n  - CTCT thu gọn: CH~3~–CH~3~\n• C~3~ (C~3~H~8~): Propane.\n  - CTCT thu gọn: CH~3~–CH~2~–CH~3~\n• C~4~ (C~4~H~10~): Có 2 đồng phân cấu tạo:\n  - Đồng phân mạch không phân nhánh: CH~3~–CH~2~–CH~2~–CH~3~ → Tên gọi: Butane.\n  - Đồng phân mạch phân nhánh: CH~3~–CH(CH~3~)–CH~3~ → Tên gọi: 2-methylpropane (hoặc Isobutane)."
    },
    {
        "id": 2,
        "content": "Đốt cháy hoàn toàn 7,437 lít khí methane trong không khí.\na) Viết PTHH xảy ra.\nb) Tính thể tích khí carbon dioxide, khối lượng nước tạo thành. Biết thể tích các khí được đo ở điều kiện chuẩn.",
        "solution": "a) Phương trình hóa học của phản ứng cháy:\n   CH~4~ + 2O~2~ -(t°)-> CO~2~ + 2H~2~O\nb) Số mol khí methane (CH~4~):\n   n~CH4~ = V / 24,79 = 7,437 / 24,79 = 0,3 (mol).\nTheo phương trình hóa học:\n- n~CO2~ = n~CH4~ = 0,3 (mol).\n  Thể tích khí CO~2~ sinh ra ở điều kiện chuẩn là:\n  V~CO2~ = 0,3 . 24,79 = 7,437 (lít).\n- n~H2O~ = 2 . n~CH4~ = 2 . 0,3 = 0,6 (mol).\n  Khối lượng nước (H~2~O) tạo thành là:\n  m~H2O~ = n . M = 0,6 . 18 = 10,8 (gam)."
    },
    {
        "id": 3,
        "content": "Đốt cháy hoàn toàn 5,8 gam khí butane C~4~H~10~ (có trong thành phần của khí gas) cần dùng bao nhiêu lít oxygen? Bao nhiêu lít không khí chứa 20% thể tích khí oxygen? Biết các khí đo ở điều kiện chuẩn. (Cho biết: C = 12; O = 16; H =1)",
        "solution": "Khối lượng mol của butane (C~4~H~10~): M = 12 . 4 + 1 . 10 = 58 (g/mol).\nSố mol khí butane C~4~H~10~:\n  n~C4H10~ = m / M = 5,8 / 58 = 0,1 (mol).\nPhương trình hóa học của phản ứng cháy:\n  2C~4~H~10~ + 13O~2~ -(t°)-> 8CO~2~ + 10H~2~O\nTheo phương trình:\n  n~O2~ = 13/2 . n~C4H10~ = 6,5 . 0,1 = 0,65 (mol).\n- Thể tích khí oxygen (O~2~) cần dùng ở đkc:\n  V~O2~ = 0,65 . 24,79 = 16,1135 (lít).\n- Thể tích không khí cần dùng (chứa 20% thể tích O~2~):\n  V~kk~ = V~O2~ . 100 / 20 = 16,1135 . 5 = 80,5675 (lít)."
    },
    {
        "id": 4,
        "content": "Hai người cùng dịch chuyển ghế từ tầng 1 lên tầng 2. Người thứ nhất dịch chuyển 20 cái trong 5 phút, người thứ hai dịch chuyển 25 cái trong 7 phút. Hỏi người nào có công suất lớn hơn. Biết mỗi cái ghế nặng 5 kg.",
        "solution": "Gọi h (m) là độ cao từ tầng 1 lên tầng 2 (h > 0).\nTrọng lượng của mỗi cái ghế là: P~0~ = 10.m = 10 . 5 = 50 N.\n• Người thứ nhất:\n  Tổng trọng lượng của 20 cái ghế: P~1~ = 20 . 50 = 1000 N.\n  Công thực hiện: A~1~ = P~1~ . h = 1000h (J).\n  Thời gian thực hiện: t~1~ = 5 phút = 5 . 60 = 300 s.\n  Công suất của người thứ nhất:\n  P~1~ = A~1~ / t~1~ = 1000h / 300 = 10/3 h ≈ 3,333h (W).\n• Người thứ hai:\n  Tổng trọng lượng của 25 cái ghế: P~2~ = 25 . 50 = 1250 N.\n  Công thực hiện: A~2~ = P~2~ . h = 1250h (J).\n  Thời gian thực hiện: t~2~ = 7 phút = 7 . 60 = 420 s.\n  Công suất của người thứ hai:\n  P~2~ = A~2~ / t~2~ = 1250h / 420 = 125/42 h ≈ 2,976h (W).\nSo sánh: Vì 10/3 h ≈ 3,333h > 125/42 h ≈ 2,976h nên P~1~ > P~2~.\nVậy người thứ nhất có công suất lớn hơn người thứ hai."
    },
    {
        "id": 5,
        "content": "a) Một xe lăn có khối lượng 250g được thả từ một máng nghiêng, người ta đo được tốc độ của xe ở chân máng nghiêng là 54 km/h. Tính độ cao của máng nghiêng? Giả sử toàn bộ thế năng của xe chuyển hóa hoàn toàn thành động năng.\nb) Một vật sáng AB đặt trước thấu kính hội tụ (vật đặt vuông góc với thấu kính) thu được một ảnh ảo, gấp 3 lần vật. Hãy dựng ảnh của vật sáng và xác định quang tâm, tiêu điểm của thấu kính.",
        "solution": "a) Đổi đơn vị:\n   Khối lượng xe lăn: m = 250 g = 0,25 kg.\n   Tốc độ ở chân máng: v = 54 km/h = 54 / 3,6 = 15 m/s.\n   Lấy gia tốc trọng trường g = 10 m/s^2^.\n   Động năng của xe lăn ở chân máng nghiêng:\n   W~đ~ = 1/2 . m . v^2^ = 1/2 . 0,25 . 15^2^ = 28,125 (J).\n   Theo định luật bảo toàn cơ năng, toàn bộ thế năng ở đỉnh chuyển hóa hoàn toàn thành động năng ở chân máng:\n   W~t~ = W~đ~ ⇔ m . g . h = 28,125\n   ⇒ h = 28,125 / (0,25 . 10) = 28,125 / 2,5 = 11,25 (m).\n   (Nếu lấy g = 9,8 m/s^2^ thì h = 15^2^ / (2 . 9,8) ≈ 11,48 m).\n\nb) Dựng ảnh ảo của AB qua thấu kính hội tụ (A'B' = 3AB) và xác định quang tâm, tiêu điểm:\n   - Vì ảnh thu được là ảnh ảo qua thấu kính hội tụ nên ảnh cùng chiều với vật, nằm cùng phía với vật so với thấu kính (vật nằm trong khoảng tiêu cự: d < f).\n   - Các bước dựng hình và xác định các điểm đặc biệt:\n     1. Vẽ trục chính Δ của thấu kính. Dựng vật sáng AB ⊥ Δ tại A. Dựng ảnh ảo A'B' ⊥ Δ tại A' cùng phía với AB, cùng chiều và có độ cao A'B' = 3AB.\n     2. Nối B' với B, đường thẳng B'B cắt trục chính Δ tại quang tâm O (vì tia sáng tới qua quang tâm thì truyền thẳng: B', B, O thẳng hàng).\n     3. Dựng thấu kính hội tụ vuông góc với trục chính Δ tại quang tâm O.\n     4. Từ B kẻ tia tới BI // Δ cắt thấu kính tại I. Nối B' với I kéo dài cắt trục chính Δ tại tiêu điểm ảnh chính F' (nằm sau thấu kính).\n     5. Lấy tiêu điểm vật chính F trên trục chính đối xứng với F' qua O (OF = OF')."
    },
    {
        "id": 6,
        "content": "a) Một quả dừa có khối lượng 1,5 kg rơi từ độ cao 6m. Tính tốc độ của quả dừa khi vừa chạm đất, giả sử bỏ qua sức cản của không khí tác dụng lên quả dừa.\nb) Một vật sáng AB đặt trước thấu kính hội tụ (vật đặt vuông góc với thấu kính) thu được một ảnh thật, gấp 2 lần vật. Hãy dựng ảnh của vật sáng và xác định quang tâm, tiêu điểm của thấu kính.",
        "solution": "a) Chọn mốc tính thế năng tại mặt đất. Lấy g = 10 m/s^2^.\n   Thế năng ban đầu của quả dừa ở độ cao h = 6 m:\n   W~t~ = P . h = 10.m . h = 10 . 1,5 . 6 = 90 (J).\n   Bỏ qua sức cản của không khí, cơ năng được bảo toàn. Khi vừa chạm đất, thế năng bằng 0, toàn bộ chuyển thành động năng:\n   W~đ~ = W~t~ ⇔ 1/2 . m . v^2^ = 90\n   ⇒ v^2^ = 2 . 90 / 1,5 = 120\n   ⇒ v = √120 ≈ 10,95 (m/s).\n   (Nếu lấy g = 9,8 m/s^2^ thì v = √(2gh) = √(2 . 9,8 . 6) = √117,6 ≈ 10,84 m/s).\n\nb) Dựng ảnh thật của AB qua thấu kính hội tụ (A'B' = 2AB) và xác định quang tâm, tiêu điểm:\n   - Vì ảnh thu được là ảnh thật nên ảnh ngược chiều với vật và nằm ở phía bên kia thấu kính so với vật (d > f).\n   - Các bước dựng hình và xác định các điểm đặc biệt:\n     1. Vẽ trục chính Δ của thấu kính. Dựng vật sáng AB ⊥ Δ tại A (nằm phía trên trục chính). Dựng ảnh thật A'B' ⊥ Δ tại A' (nằm phía dưới trục chính, ngược chiều với AB) sao cho A'B' = 2AB.\n     2. Nối đỉnh vật B với đỉnh ảnh B', đoạn thẳng BB' cắt trục chính Δ tại điểm O, đó chính là quang tâm của thấu kính (vì tia sáng qua quang tâm truyền thẳng nên B, O, B' thẳng hàng).\n     3. Dựng thấu kính hội tụ vuông góc với trục chính Δ tại quang tâm O.\n     4. Từ B kẻ tia tới BI // Δ cắt thấu kính tại I. Tia tới // trục chính cho tia ló đi qua tiêu điểm ảnh chính F' và đi qua đỉnh ảnh thật B'. Do đó, nối I với B', đoạn thẳng IB' cắt trục chính Δ tại tiêu điểm ảnh chính F' (nằm sau thấu kính).\n     5. Lấy tiêu điểm vật chính F trên trục chính đối xứng với F' qua quang tâm O (OF = OF')."
    },
    {
        "id": 7,
        "content": "Ở đậu Hà Lan, tính trạng hoa tím là trội hoàn toàn so với hoa trắng. Kiểu gen AA quy định hoa tím thuần chủng và Aa quy định hoa tím không thuần chủng. Em hãy nêu phương pháp xác định kiểu gen của hai cây hoa tím nói trên.",
        "solution": "Để xác định kiểu gen của cá thể mang tính trạng trội (cây hoa tím AA hoặc Aa), ta có thể sử dụng một trong hai phương pháp:\n\n* Phương pháp 1: Thực hiện phép lai phân tích (Đem lai với cây mang tính trạng lặn - hoa trắng aa)\n- Cho từng cây hoa tím lai với cây hoa trắng (có kiểu gen đồng hợp lặn aa):\n  + Nếu đời con F~1~ đồng loạt 100% hoa tím:\n    Sơ đồ lai: P: AA (hoa tím) × aa (hoa trắng) → F~1~: 100% Aa (hoa tím).\n    ⇒ Cây hoa tím đem lai có kiểu gen thuần chủng (AA).\n  + Nếu đời con F~1~ phân tính theo tỉ lệ 1 hoa tím : 1 hoa trắng (50% tím : 50% trắng):\n    Sơ đồ lai: P: Aa (hoa tím) × aa (hoa trắng) → F~1~: 1 Aa (hoa tím) : 1 aa (hoa trắng).\n    ⇒ Cây hoa tím đem lai có kiểu gen không thuần chủng (Aa).\n\n* Phương pháp 2: Cho từng cây hoa tím tự thụ phấn nghiêm ngặt\n- Cho từng cây hoa tím tự thụ phấn riêng rẽ qua 1 thế hệ và quan sát kiểu hình đời con:\n  + Nếu đời con đồng tính 100% hoa tím:\n    Sơ đồ lai: P: AA × AA → F~1~: 100% AA (hoa tím).\n    ⇒ Cây hoa tím là thuần chủng (AA).\n  + Nếu đời con phân tính theo tỉ lệ 3 hoa tím : 1 hoa trắng (75% tím : 25% trắng):\n    Sơ đồ lai: P: Aa × Aa → F~1~: 1 AA : 2 Aa : 1 aa (3 hoa tím : 1 hoa trắng).\n    ⇒ Cây hoa tím là không thuần chủng (Aa)."
    },
    {
        "id": 8,
        "content": "Qua khảo sát tại một vườn đậu Hà Lan, người ta nhận thấy có những cây kháng bệnh tốt, giúp giảm sử dụng thuốc bảo vệ thực vật. Giả sử tính trạng kháng bệnh do một cặp gene quy định, trong đó allele A quy định kháng bệnh trội hoàn toàn so với allele a quy định không kháng bệnh. Em hãy đề xuất cách chọn được cây kháng bệnh thuần chủng nhằm nhân giống, góp phần bảo vệ tự nhiên.",
        "solution": "• Cơ sở di truyền:\n  - Allele A: kháng bệnh (trội hoàn toàn).\n  - Allele a: không kháng bệnh (lặn).\n  - Cây kháng bệnh có thể có kiểu gen thuần chủng (AA) hoặc không thuần chủng (Aa). Mục tiêu là chọn được cây có kiểu gen AA để làm giống lâu dài không bị phân li tính trạng.\n\n• Quy trình đề xuất các bước thực hiện:\n- Bước 1: Khảo sát và chọn lọc các cá thể cây có kiểu hình kháng bệnh tốt, sinh trưởng khỏe mạnh trong vườn đậu Hà Lan.\n- Bước 2: Kiểm tra độ thuần chủng của các cây đã chọn bằng một trong hai cách:\n  + Cách 1 (Tự thụ phấn qua các thế hệ - tối ưu nhất đối với đậu Hà Lan vì có hoa lưỡng tính tự thụ phấn nghiêm ngặt):\n    Cho từng cây kháng bệnh tự thụ phấn riêng rẽ, thu hoạch hạt của từng cây gieo thành từng luống/dòng riêng biệt để theo dõi đời con:\n    * Dòng nào cho toàn bộ thế hệ con 100% cây kháng bệnh qua các thế hệ ⇒ Cây ban đầu mang kiểu gen thuần chủng (AA).\n    * Dòng nào xuất hiện cây bị bệnh (kiểu hình lặn aa chiếm 25%) ⇒ Cây ban đầu là dị hợp (Aa), loại bỏ dòng này.\n  + Cách 2 (Lai phân tích):\n    Đem từng cây kháng bệnh lai với cây không kháng bệnh (kiểu gen lặn aa).\n    * Nếu đời con 100% kháng bệnh ⇒ Cây đem lai thuần chủng (AA).\n    * Nếu đời con xuất hiện cây không kháng bệnh (tỉ lệ 1:1) ⇒ Cây đem lai không thuần chủng (Aa), loại bỏ.\n- Bước 3: Nhân giống dòng thuần chủng:\n  Lấy hạt của các cây thuần chủng (AA) đã xác định ở Bước 2 tiến hành nhân giống trên diện rộng bằng tự thụ phấn. Giống thuần chủng duy trì ổn định đặc tính kháng bệnh cho các mùa vụ sau, giúp hạn chế sâu bệnh và giảm thiểu tối đa việc sử dụng thuốc bảo vệ thực vật hóa học, góp phần bảo vệ môi trường và hệ sinh thái tự nhiên."
    }
]

print("Loaded all questions data successfully!")
