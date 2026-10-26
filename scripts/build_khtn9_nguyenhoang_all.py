# -*- coding: utf-8 -*-
"""
Script biên soạn 3 tài liệu chuẩn theo FORM ĐỀ KIỂM TRA:
1. De_Cuong_GK1_KHTN9_NguyenHoang_HocSinh.docx (Bản Học Sinh)
2. De_Cuong_GK1_KHTN9_NguyenHoang_GiaoVien.docx (Bản Giáo Viên có lời giải chi tiết)
3. De_Cuong_GK1_KHTN9_NguyenHoang_DayTrenManHinh.pdf (Slide PDF tỉ lệ 16:9 Full HD dạy trên Tivi 65 inch)
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
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

# ==============================================================================
# DỮ LIỆU ĐỀ CƯƠNG ÔN TẬP KHTN 9 - THCS NGUYỄN HOÀNG
# ==============================================================================

LY_THUYET_DATA = [
    {
        "title": "1. Khái niệm hợp chất hữu cơ và Hóa học hữu cơ",
        "items": [
            "• *Khái niệm:* Hợp chất hữu cơ là hợp chất của carbon (trừ các chất vô cơ như CO, CO~2~, H~2~CO~3~, các muối carbonate, carbide, cyanide,...). Hóa học hữu cơ là ngành Hóa học chuyên nghiên cứu về hợp chất hữu cơ và những biến đổi của chúng.",
            "• *Phân loại hợp chất hữu cơ:* Gồm 2 loại chính:\n  + *Hydrocarbon:* Phân tử chỉ chứa 2 nguyên tố carbon và hydrogen (ví dụ: CH~4~, C~2~H~4~, C~2~H~2~, C~6~H~6~,...).\n  + *Dẫn xuất của hydrocarbon:* Phân tử ngoài carbon, hydrogen còn có các nguyên tố khác như oxygen, nitrogen, halogen, sulfur,... (ví dụ: C~2~H~5~OH, CH~3~COOH, C~6~H~12~O~6~, CH~3~Cl,...).",
            "• *Đặc điểm cấu tạo:* Nguyên tử carbon luôn có hóa trị IV. Các nguyên tử carbon có thể liên kết trực tiếp với nhau tạo thành mạch carbon: mạch hở không phân nhánh, mạch hở phân nhánh hoặc mạch vòng. Liên kết giữa các nguyên tử carbon có thể là liên kết đơn, liên kết đôi hoặc liên kết ba.",
            "• *Công thức phân tử và công thức cấu tạo:* CTPT biểu thị số lượng từng nguyên tử trong phân tử. CTCT biểu diễn trật tự và cách thức liên kết của các nguyên tử trong phân tử."
        ]
    },
    {
        "title": "2. Alkane",
        "items": [
            "• *Khái niệm:* Alkane là hydrocarbon no, mạch hở, trong phân tử chỉ chứa các liên kết đơn C-C và C-H.",
            "• *Công thức chung:* C~n~H~2n+2~ (với n ≥ 1). Alkane đơn giản nhất là methane (CH~4~).",
            "• *Phản ứng cháy (oxi hóa hoàn toàn):* C~n~H~2n+2~ + ((3n+1)/2)O~2~ --t°--> nCO~2~ + (n+1)H~2~O. Phản ứng tỏa nhiều nhiệt, tỉ lệ n~H2O~ > n~CO2~.",
            "• *Ứng dụng:* Dùng làm nhiên liệu (khí gas LPG, khí thiên nhiên, xăng, dầu diesel,...); làm nguyên liệu quan trọng trong công nghiệp hóa chất."
        ]
    },
    {
        "title": "3. Alkene",
        "items": [
            "• *Khái niệm:* Alkene là hydrocarbon không no, mạch hở, trong phân tử có chứa một liên kết đôi C=C.",
            "• *Công thức chung:* C~n~H~2n~ (với n ≥ 2). Alkene đơn giản nhất là ethylene (C~2~H~4~)."
        ]
    },
    {
        "title": "4. Ethylene",
        "items": [
            "• *Đặc điểm cấu tạo:* CTPT: C~2~H~4~ (M = 28); CTCT: CH~2~=CH~2~. Trong liên kết đôi C=C có một liên kết kém bền, dễ bị đứt ra trong các phản ứng hóa học.",
            "• *Tính chất vật lý:* Là chất khí không màu, không mùi, nhẹ hơn không khí (d = 28/29 ≈ 0,97), rất ít tan trong nước.",
            "• *Tính chất hóa học:*\n  + *Phản ứng cháy:* C~2~H~4~ + 3O~2~ --t°--> 2CO~2~ + 2H~2~O (tỏa nhiều nhiệt, n~CO2~ = n~H2O~).\n  + *Phản ứng cộng:* Làm mất màu vàng nâu của dung dịch nước bromine: CH~2~=CH~2~ + Br~2~ ➔ CH~2~Br-CH~2~Br (1,2-dibromoethane). Có phản ứng cộng với H~2~, HCl, H~2~O (xúc tác thích hợp).\n  + *Phản ứng trùng hợp:* nCH~2~=CH~2~ --xt, t°, p--> (-CH~2~-CH~2~-)~n~ (polyethylene - nhựa PE).",
            "• *Ứng dụng:* Sản xuất nhựa PE, PVC, dung môi, rượu etylic, acid acetic; dùng để kích thích trái cây mau chín."
        ]
    },
    {
        "title": "5. Nguồn nhiên liệu",
        "items": [
            "• *Dầu mỏ:* Là chất lỏng sánh màu nâu đen, không tan trong nước và nhẹ hơn nước. Dầu mỏ là hỗn hợp tự nhiên phức tạp của nhiều loại hydrocarbon. Chưng cất dầu mỏ thu được: khí hóa lỏng (LPG), xăng, dầu hỏa, dầu diesel, dầu bôi trơn, sáp paraffin, nhựa đường.",
            "• *Khí thiên nhiên và khí mỏ dầu:* Có thành phần chủ yếu là khí methane (CH~4~ chiếm từ 85% - 95%).",
            "• *Phân loại nhiên liệu:* Nhiên liệu rắn (than đá, than củi, gỗ...), nhiên liệu lỏng (xăng, cồn, dầu hỏa, dầu diesel...), nhiên liệu khí (khí gas LPG, khí thiên nhiên, khí sinh học biogas...).",
            "• *Sử dụng an toàn, tiết kiệm và hiệu quả:* Cung cấp đủ không khí/oxygen; tăng diện tích tiếp xúc giữa nhiên liệu và không khí; tận dụng triệt để nhiệt lượng thải ra."
        ]
    },
    {
        "title": "6. Ethylic alcohol (Ethanol)",
        "items": [
            "• *Công thức:* CTPT: C~2~H~6~O (M = 46); CTCT: CH~3~-CH~2~-OH (viết gọn C~2~H~5~OH, có nhóm -OH linh động quyết định tính chất hóa học).",
            "• *Tính chất vật lý và độ cồn:* Chất lỏng không màu, mùi thơm nhẹ, vị cay, nhẹ hơn nước, tan vô hạn trong nước. Độ cồn (độ rượu) là số mL ethylic alcohol nguyên chất có trong 100 mL dung dịch rượu ở 20°C.",
            "• *Tính chất hóa học:*\n  + *Phản ứng cháy:* C~2~H~5~OH + 3O~2~ --t°--> 2CO~2~ + 3H~2~O (ngọn lửa xanh mờ, tỏa nhiệt mạnh).\n  + *Tác dụng với kim loại kiềm (Na):* 2C~2~H~5~OH + 2Na ➔ 2C~2~H~5~ONa + H~2~↑ (sủi bọt khí hydrogen).\n  + *Phản ứng lên men giấm:* C~2~H~5~OH + O~2~ --men giấm--> CH~3~COOH + H~2~O.",
            "• *Ứng dụng:* Nhiên liệu sinh học (xăng E5); sát trùng y tế (cồn 70°); dung môi hòa tan hương liệu, dược phẩm; nguyên liệu sản xuất giấm ăn, ester, cao su tổng hợp."
        ]
    },
    {
        "title": "7. Acetic acid",
        "items": [
            "• *Công thức:* CTPT: C~2~H~4~O~2~ (M = 60); CTCT: CH~3~-COOH (có nhóm carboxyl -COOH mang tính acid).",
            "• *Tính chất vật lý:* Chất lỏng không màu, mùi xốc, vị chua đặc trưng, tan vô hạn trong nước. Dung dịch acetic acid có nồng độ 2% - 5% được dùng làm *giấm ăn*.",
            "• *Tính chất hóa học:* Mang tính acid của một carboxylic acid (acid yếu nhưng mạnh hơn carbonic acid H~2~CO~3~):\n  + Làm đổi màu giấy quỳ tím sang màu đỏ.\n  + Tác dụng với kim loại đứng trước H: 2CH~3~COOH + 2Na ➔ 2CH~3~COONa + H~2~↑; 2CH~3~COOH + Mg ➔ (CH~3~COO)~2~Mg + H~2~↑.\n  + Tác dụng với basic oxide và base: 2CH~3~COOH + CuO ➔ (CH~3~COO)~2~Cu + H~2~O; CH~3~COOH + NaOH ➔ CH~3~COONa + H~2~O.\n  + Tác dụng với muối carbonate: 2CH~3~COOH + CaCO~3~ ➔ (CH~3~COO)~2~Ca + CO~2~↑ + H~2~O.",
            "• *Ứng dụng:* Dùng làm gia vị (giấm ăn), chất bảo quản thực phẩm; sản xuất tơ nhân tạo, chất dẻo, phẩm nhuộm, dược phẩm (aspirin,...)."
        ]
    }
]

# Phần I: Trắc nghiệm nhiều lựa chọn (8 câu)
PART1_QUESTIONS = [
    {
        "num": 1,
        "prompt": "Câu 1. Dãy chất nào sau đây đều là hydrocarbon?",
        "options": [
            "A. FeCl~2~, C~2~H~6~O, CH~4~, NaHCO~3~.",
            "B. C~6~H~5~Na, CH~4~O, HNO~3~, C~6~H~6~.",
            "*C. CH~4~, C~2~H~4~, C~2~H~2~, C~6~H~6~.*",
            "D. CH~3~NO~2~, CH~3~Br, NaOH, C~2~H~5~OH."
        ],
        "ans": "C",
        "sol": "Hydrocarbon là hợp chất hữu cơ mà trong phân tử chỉ chứa 2 nguyên tố carbon và hydrogen.\n• Dãy C (CH~4~, C~2~H~4~, C~2~H~2~, C~6~H~6~) đều chỉ cấu tạo từ C và H nên đều là hydrocarbon.\n• Ở các phương án khác: FeCl~2~, NaHCO~3~, HNO~3~, NaOH là các hợp chất vô cơ; C~2~H~6~O, C~6~H~5~Na, CH~4~O, CH~3~NO~2~, CH~3~Br, C~2~H~5~OH là các dẫn xuất của hydrocarbon."
    },
    {
        "num": 2,
        "prompt": "Câu 2. Hợp chất hữu cơ là",
        "options": [
            "A. Hợp chất của oxygen với một nguyên tố hóa học khác.",
            "B. Hợp chất của nitrogen và một nguyên tố phi kim khác.",
            "C. Hợp chất của carbon và hydrogen ngoài ra không chứa các nguyên tố khác.",
            "*D. Hợp chất của carbon (trừ CO, CO~2~, H~2~CO~3~ và các muối carbonate, …).*"
        ],
        "ans": "D",
        "sol": "Theo định nghĩa chuẩn: Hợp chất hữu cơ là hợp chất của carbon (trừ các hợp chất vô cơ đơn giản của carbon như CO, CO~2~, H~2~CO~3~, các muối carbonate kim loại, carbide, cyanide,...)."
    },
    {
        "num": 3,
        "prompt": "Câu 3. Cho phương trình hóa học: 2X + 7O~2~ --t°--> 4CO~2~ + 6H~2~O. Công thức phân tử của X là",
        "options": [
            "A. C~2~H~2~.",
            "B. C~2~H~4~.",
            "*C. C~2~H~6~.*",
            "D. C~5~H~12~."
        ],
        "ans": "C",
        "sol": "Áp dụng định luật bảo toàn nguyên tố cho phản ứng:\n• Vế phải có: 4 nguyên tử C; 6 × 2 = 12 nguyên tử H; 4 × 2 + 6 = 14 nguyên tử O.\n• Vế trái có: 7 phân tử O~2~ (chứa 14 nguyên tử O).\n➔ Trong 2 phân tử X phải chứa 4 nguyên tử C và 12 nguyên tử H (không chứa O).\nDo đó mỗi phân tử X có: 4 / 2 = 2 nguyên tử C và 12 / 2 = 6 nguyên tử H. Vậy công thức phân tử của X là C~2~H~6~ (ethane)."
    },
    {
        "num": 4,
        "prompt": "Câu 4. Khi đốt cháy hoàn toàn 1,2395 L alkane X thu được 6,1975 L khí CO~2~ (đo ở đkc). Công thức phân tử của X là",
        "options": [
            "A. C~3~H~8~.",
            "B. C~5~H~10~.",
            "C. C~4~H~10~.",
            "*D. C~5~H~12~.*"
        ],
        "ans": "D",
        "sol": "Ở cùng điều kiện nhiệt độ và áp suất (đkc), tỉ lệ thể tích khí bằng tỉ lệ số mol:\n• Số nguyên tử carbon trong phân tử alkane X là: n = V~CO2~ / V~X~ = 6,1975 / 1,2395 = 5.\n• Alkane có công thức phân tử chung là C~n~H~2n+2~ (n ≥ 1). Với n = 5, ta được C~5~H~12~ (pentane)."
    },
    {
        "num": 5,
        "prompt": "Câu 5. Trong các dãy chất sau, dãy chất nào là dẫn xuất của hydrocarbon?",
        "options": [
            "A. CH~4~, C~2~H~4~O, C~2~H~2~.",
            "B. C~2~H~2~, C~2~H~4~, C~6~H~6~.",
            "*C. CH~3~Cl, C~2~H~6~O, C~6~H~5~NO~2~.*",
            "D. C~2~H~4~O~2~, C~6~H~6~Cl~6~, C~6~H~6~."
        ],
        "ans": "C",
        "sol": "Dẫn xuất của hydrocarbon là những hợp chất hữu cơ mà trong phân tử ngoài C và H còn có các nguyên tố khác như O, Cl, N, S,...\n• Dãy C gồm CH~3~Cl (chứa Cl), C~2~H~6~O (chứa O), C~6~H~5~NO~2~ (chứa N và O) đều là dẫn xuất của hydrocarbon.\n• Các phương án A, B, D đều có chứa CH~4~, C~2~H~2~, C~2~H~4~, C~6~H~6~ là các hydrocarbon thuần túy."
    },
    {
        "num": 6,
        "prompt": "Câu 6. Công thức phân tử của chất hữu cơ nào cũng nhất thiết phải chứa nguyên tố",
        "options": [
            "*A. carbon.*",
            "B. oxygen.",
            "C. nitrogen.",
            "D. chlorine."
        ],
        "ans": "A",
        "sol": "Theo định nghĩa, hợp chất hữu cơ là hợp chất của nguyên tố carbon (trừ một số hợp chất vô cơ đơn giản). Do đó, trong phân tử của bất kì hợp chất hữu cơ nào cũng nhất thiết phải chứa nguyên tố carbon."
    },
    {
        "num": 7,
        "prompt": "Câu 7. Đốt cháy 0,5 mol hydrocarbon X thu được H~2~O và 22 g CO~2~. Công thức phân tử của X là",
        "options": [
            "*A. CH~4~.*",
            "B. C~2~H~2~.",
            "C. C~2~H~4~.",
            "D. C~3~H~8~."
        ],
        "ans": "A",
        "sol": "• Số mol khí CO~2~ thu được: n~CO2~ = m / M = 22 / 44 = 0,5 mol.\n• Số nguyên tử carbon trong phân tử hydrocarbon X là: Số C = n~CO2~ / n~X~ = 0,5 / 0,5 = 1.\n• Hydrocarbon là hợp chất chỉ chứa C và H. Hydrocarbon duy nhất có 1 nguyên tử carbon trong phân tử là methane (CH~4~)."
    },
    {
        "num": 8,
        "prompt": "Câu 8. Đốt cháy hoàn toàn 6,1975 L khí methane (đkc). Thể tích khí oxygen và thể tích không khí (đkc) cần dùng để đốt cháy hết lượng methane trên lần lượt là (biết oxygen chiếm 20% thể tích không khí)",
        "options": [
            "A. 6,1975 L và 30,9875 L.",
            "*B. 12,395 L và 61,975 L.*",
            "C. 24,79 L và 123,95 L.",
            "D. 37,185 L và 185,925 L."
        ],
        "ans": "B",
        "sol": "Phương trình hóa học của phản ứng cháy:\nCH~4~ + 2O~2~ --t°--> CO~2~ + 2H~2~O\n• Theo phương trình: V~O2~ = 2 × V~CH4~ = 2 × 6,1975 = 12,395 L.\n• Vì thể tích khí oxygen chiếm 20% (1/5) thể tích của không khí nên:\nV~không khí~ = 5 × V~O2~ = 5 × 12,395 = 61,975 L."
    }
]

# Phần II: Trắc nghiệm Đúng/Sai (4 câu: Câu 9, 10, 11, 12)
PART2_QUESTIONS = [
    {
        "num": 9,
        "prompt": "Câu 9. Cho phát biểu: Công thức chung của alkene là C~n~H~2n~ và alkene đơn giản nhất là C~2~H~4~.",
        "stmts": [
            ("a) Với n = 1 ta có một alkene với công thức CH~2~.", "Sai"),
            ("b) Tất cả những hợp chất hữu cơ thuộc loại alkene đều làm mất màu vàng nâu của nước bromine. Phản ứng đặc trưng là phản ứng cộng.", "Đúng"),
            ("c) Với alkane và alkene đều tham gia phản ứng trùng hợp.", "Sai"),
            ("d) Khi đốt cháy một alkene ta thu được hai sản phẩm là hơi nước và khí carbon dioxide.", "Đúng")
        ],
        "sol": "• a) Sai. Alkene là hydrocarbon không no mạch hở có 1 liên kết đôi C=C, do đó cần tối thiểu 2 nguyên tử C trong phân tử (điều kiện n ≥ 2). Không tồn tại phân tử alkene CH~2~.\n• b) Đúng. Mọi alkene đều có liên kết đôi C=C gồm một liên kết kém bền, dễ dàng tham gia phản ứng cộng làm mất màu vàng nâu của dung dịch nước bromine.\n• c) Sai. Chỉ có alkene (chứa liên kết đôi C=C) mới có khả năng tham gia phản ứng trùng hợp tạo polymer (như ethylene trùng hợp thành PE). Alkane là hợp chất no, chỉ có liên kết đơn bền vững nên không tham gia phản ứng trùng hợp.\n• d) Đúng. Khi đốt cháy hoàn toàn alkene (hay bất kì hydrocarbon nào), các nguyên tố C và H kết hợp với oxygen tạo thành khí CO~2~ và hơi nước: C~n~H~2n~ + 1,5nO~2~ --t°--> nCO~2~ + nH~2~O."
    },
    {
        "num": 10,
        "prompt": "Câu 10. Xét tính đúng – sai của các phát biểu sau về dầu mỏ và khí thiên nhiên:",
        "stmts": [
            ("a) Dầu mỏ và khí thiên nhiên là nhiên liệu hóa thạch có thành phần chủ yếu là khí methane.", "Sai"),
            ("b) Khi chưng cất dầu mỏ ta thu được hai sản phẩm là xăng và dầu diesel.", "Sai"),
            ("c) Dầu mỏ là hỗn hợp của dẫn xuất của hydrocarbon.", "Sai"),
            ("d) Mỏ dầu gồm ba lớp theo thứ tự từ trên xuống dưới là khí đồng hành, dầu lỏng và nước mặn.", "Đúng")
        ],
        "sol": "• a) Sai. Khí thiên nhiên có thành phần chủ yếu là khí methane (khoảng 85% - 95%), nhưng dầu mỏ là hỗn hợp phức tạp của hàng trăm hydrocarbon khác nhau ở các thể lỏng, rắn và khí tan, không phải chủ yếu là methane.\n• b) Sai. Quá trình chưng cất dầu mỏ thu được nhiều phân đoạn sản phẩm khác nhau ở các khoảng nhiệt độ sôi khác nhau: khí dầu mỏ hóa lỏng (LPG), dầu nhẹ, xăng, dầu hỏa, dầu diesel, dầu nhờn bôi trơn, sáp paraffin, nhựa đường... chứ không chỉ thu được 2 sản phẩm.\n• c) Sai. Dầu mỏ là hỗn hợp tự nhiên của nhiều loại hydrocarbon (chủ yếu là alkane, cycloalkane và arene), không phải là dẫn xuất của hydrocarbon.\n• d) Đúng. Cấu tạo địa chất của mỏ dầu thông thường gồm 3 lớp rõ rệt từ trên xuống dưới: lớp khí mỏ dầu (khí đồng hành), lớp dầu lỏng (dầu thô) và lớp nước mặn ở đáy mỏ."
    },
    {
        "num": 11,
        "prompt": "Câu 11. Xét tính đúng – sai của các phát biểu sau về cấu tạo và tính chất của ethylene:",
        "stmts": [
            ("a) Trong công thức cấu tạo của alkene có hai liên kết đôi C=C.", "Sai"),
            ("b) Trong liên kết đôi C=C có một liên kết kém bền, dễ bị đứt ra trong các phản ứng hóa học.", "Đúng"),
            ("c) Phản ứng giữa ethylene C~2~H~4~ và nước bromine được gọi là phản ứng cộng. Hiện tượng xảy ra là nước bromine màu vàng nâu bị mất màu.", "Đúng"),
            ("d) Ethylene là chất khí không màu, không mùi, hầu như không tan trong nước và nặng hơn không khí.", "Sai")
        ],
        "sol": "• a) Sai. Phân tử của một alkene mạch hở chỉ chứa một liên kết đôi C=C (chứ không phải hai liên kết đôi).\n• b) Đúng. Trong liên kết đôi C=C có một liên kết σ bền và một liên kết π kém bền hơn, liên kết kém bền này dễ bị bẻ gãy khi tham gia phản ứng hóa học.\n• c) Đúng. Phương trình: CH~2~=CH~2~ + Br~2~ ➔ CH~2~Br-CH~2~Br. Đây là phản ứng cộng đặc trưng, làm dung dịch bromine màu vàng nâu chuyển sang không màu.\n• d) Sai. Ethylene (C~2~H~4~) có phân tử khối M = 28 g/mol, nhẹ hơn không khí (M~không khí~ ≈ 29 g/mol, d ≈ 28/29 ≈ 0,97)."
    },
    {
        "num": 12,
        "prompt": "Câu 12. Khi chưng cất dầu thô ta thu được các sản phẩm: khí hóa lỏng, dầu nhẹ, xăng, dầu hỏa, dầu diesel, dầu bôi trơn, sáp paraffin, nhựa đường. Xét tính đúng – sai của các ứng dụng sau:",
        "stmts": [
            ("a) Xăng và dầu diesel là nhiên liệu dùng cho động cơ đốt trong (xe máy, ô tô,…).", "Đúng"),
            ("b) Nhựa đường dùng trong ngành xây dựng cầu, đường bộ,…", "Đúng"),
            ("c) Dầu hỏa chủ yếu được sử dụng để nấu đồ ăn, sưởi ấm.", "Đúng"),
            ("d) Khí hóa lỏng được sử dụng trong động cơ máy bay.", "Sai")
        ],
        "sol": "• a) Đúng. Xăng là nhiên liệu cho động cơ đánh lửa (xe máy, ô tô con), còn dầu diesel dùng cho động cơ nén cháy của xe tải, tàu hỏa, tàu thủy, máy phát điện.\n• b) Đúng. Nhựa đường (bitum) là phần cặn còn lại sau chưng cất, có độ bám dính và chống thấm cao, dùng chủ yếu để rải thảm mặt đường bộ và làm vật liệu xây dựng.\n• c) Đúng. Dầu hỏa là chất đốt gia dụng thông dụng dùng để đun nấu, thắp đèn, sưởi ấm trong đời sống.\n• d) Sai. Động cơ máy bay phản lực sử dụng nhiên liệu đặc biệt là dầu phản lực tinh chế từ phân đoạn dầu hỏa (kerosene / Jet A-1), không dùng khí hóa lỏng (LPG)."
    }
]

# Phần III: Trắc nghiệm Trả lời ngắn (Câu 13 gồm các ý a, b, c, d, e, f, g)
PART3_QUESTIONS = [
    {
        "id": "a",
        "q": "a) Acetic acid thuộc loại hợp chất hóa học nào có khả năng làm quỳ tím hóa đỏ?",
        "ans": "Acid (Axit)",
        "sol": "Acetic acid (CH~3~COOH) thuộc loại carboxylic acid (acid hữu cơ), mang tính acid yếu nhưng vẫn làm dung dịch quỳ tím chuyển sang màu đỏ."
    },
    {
        "id": "b",
        "q": "b) Tên gọi thông thường của dung dịch acetic acid nồng độ 2% - 5% là gì?",
        "ans": "Giấm ăn (Giấm)",
        "sol": "Dung dịch acetic acid có nồng độ từ 2% đến 5% thường được gọi là giấm ăn, dùng làm gia vị chua trong chế biến thực phẩm."
    },
    {
        "id": "c",
        "q": "c) Ngoài nước và khí CO~2~, phản ứng giữa acetic acid với đá vôi (CaCO~3~) tạo ra loại hợp chất gì?",
        "ans": "Muối (Calcium acetate)",
        "sol": "Phương trình: 2CH~3~COOH + CaCO~3~ ➔ (CH~3~COO)~2~Ca + CO~2~↑ + H~2~O. Hợp chất sinh ra cùng với CO~2~ và H~2~O là muối (muối calcium acetate)."
    },
    {
        "id": "d",
        "q": "d) Tên gọi thông thường của ethylic alcohol dùng để lên men giấm là gì?",
        "ans": "Rượu (Cồn)",
        "sol": "Ethylic alcohol (C~2~H~5~OH) trong đời sống hàng ngày thường được gọi là rượu hoặc cồn. Phương trình lên men giấm: C~2~H~5~OH + O~2~ --men giấm--> CH~3~COOH + H~2~O."
    },
    {
        "id": "e",
        "q": "e) Vị đặc trưng của giấm ăn (acetic acid) khi dùng trong chế biến món ăn là gì?",
        "ans": "Vị chua",
        "sol": "Vị chua thanh mát là vị đặc trưng của các phân tử acid hữu cơ (chứa nhóm chức -COOH), cụ thể là acetic acid trong giấm ăn."
    },
    {
        "id": "f",
        "q": "f) Ngoài khí CO~2~ và muối, phản ứng giữa giấm ăn và cặn đá vôi ở ấm đun nước sinh ra chất gì?",
        "ans": "Nước (H~2~O)",
        "sol": "Cặn bám trong ấm đun nước thành phần chủ yếu là CaCO~3~. Phản ứng: 2CH~3~COOH + CaCO~3~ ➔ (CH~3~COO)~2~Ca + CO~2~↑ + H~2~O. Chất sinh ra ngoài khí CO~2~ và muối là nước (H~2~O)."
    },
    {
        "id": "g",
        "q": "g) Tên gọi thông thường của ethylic alcohol dùng để lên men giấm là gì?",
        "ans": "Rượu (Cồn)",
        "sol": "Ethylic alcohol được gọi là rượu (rượu etylic) hoặc cồn."
    }
]

# Phần Tự luận (Câu 14 - 1 điểm)
ESSAY_QUESTION = {
    "title": "Câu 14. (1,0 điểm)",
    "context": "Trong đời sống hằng ngày, ethylic alcohol có rất nhiều ứng dụng quan trọng. Xăng sinh học E5 là nhiên liệu chứa 5% thể tích ethylic alcohol nguyên chất và 95% thể tích xăng truyền thống, giúp động cơ vận hành hiệu quả và giảm thiểu khí thải gây ô nhiễm môi trường. Bên cạnh đó, trong y tế, dung dịch cồn 70° được sử dụng phổ biến để diệt khuẩn, sát trùng vết thương và làm nước rửa tay khô.",
    "sub_qs": [
        "a) Viết công thức cấu tạo của ethylic alcohol.",
        "b) Hãy cho biết ý nghĩa cồn 60°.",
        "c) Viết phương trình hóa học xảy ra khi cho viên Na vào dung dịch C~2~H~5~OH 60°."
    ],
    "sol": (
        "• a) Công thức cấu tạo của ethylic alcohol (C~2~H~6~O):\n"
        "   - Công thức cấu tạo thu gọn: CH~3~-CH~2~-OH (hoặc C~2~H~5~OH).\n"
        "   - Công thức cấu tạo đầy đủ biểu diễn liên kết: H~3~C–CH~2~–O–H (trong phân tử có một nguyên tử H liên kết với nguyên tử O tạo thành nhóm chức alcohol –OH linh động).\n"
        "• b) Ý nghĩa của cồn 60°:\n"
        "   - Cồn 60° (độ rượu = 60°) có nghĩa là: Trong 100 mL dung dịch cồn (rượu) ở 20°C có chứa 60 mL ethylic alcohol nguyên chất và 40 mL nước cất.\n"
        "• c) Phương trình hóa học khi cho viên sodium (Na) vào dung dịch C~2~H~5~OH 60°:\n"
        "   - Dung dịch cồn 60° là hỗn hợp gồm ethylic alcohol (C~2~H~5~OH) và nước (H~2~O). Cả hai chất đều tác dụng với sodium giải phóng khí hydrogen (H~2~).\n"
        "   - Các phương trình hóa học xảy ra:\n"
        "     (1) 2Na + 2H~2~O ➔ 2NaOH + H~2~↑\n"
        "     (2) 2Na + 2C~2~H~5~OH ➔ 2C~2~H~5~ONa + H~2~↑ (sodium ethoxide)\n"
        "   - Do liên kết O-H trong nước phân cực hơn liên kết O-H trong alcohol nên phản ứng (1) xảy ra mãnh liệt hơn so với phản ứng (2)."
    )
}

# ==============================================================================
# HÀM TẠO FILE WORD THEO FORM ĐỀ KIỂM TRA
# ==============================================================================

def generate_word_document(is_teacher=False):
    filename = "De_Cuong_GK1_KHTN9_NguyenHoang_GiaoVien.docx" if is_teacher else "De_Cuong_GK1_KHTN9_NguyenHoang_HocSinh.docx"
    filepath = os.path.join(OUTPUT_DIR, filename)

    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="TRƯỜNG THCS NGUYỄN HOÀNG",
        exam_title="ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA KÌ I",
        subject_line="MÔN: KHTN 9 (PHÂN MÔN HÓA HỌC)",
        num_pages="03",
        duration="45 phút",
        exam_code="KHTN9-NH",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Ca = 40; Br = 80",
        is_teacher=is_teacher
    )

    # A. PHẦN LÝ THUYẾT
    add_section_header(doc, "A. LÝ THUYẾT TRỌNG TÂM", "Học sinh ôn tập kĩ các nội dung kiến thức cơ bản sau:")
    for section in LY_THUYET_DATA:
        p_sec = doc.add_paragraph()
        p_sec.paragraph_format.space_before = Pt(4)
        p_sec.paragraph_format.space_after = Pt(2)
        render_rich_text(p_sec, f"*{section['title']}*", base_size=11, base_bold=True, base_color=COLOR_SOLUTION)

        for item in section["items"]:
            lines = item.split("\n")
            for line in lines:
                p_it = doc.add_paragraph()
                p_it.paragraph_format.left_indent = Cm(0.4)
                p_it.paragraph_format.space_before = Pt(1)
                p_it.paragraph_format.space_after = Pt(1)
                render_rich_text(p_it, line, base_size=10.5)

    # B. PHẦN BÀI TẬP
    add_section_header(doc, "B. BÀI TẬP ÔN TẬP VÀ LUYỆN ĐỀ")

    # 1. PHẦN TRẮC NGHIỆM NHIỀU LỰA CHỌN
    add_section_header(doc, "1. PHẦN TRẮC NGHIỆM NHIỀU LỰA CHỌN", "Chọn một phương án trả lời đúng nhất trong mỗi câu hỏi sau:")
    for q in PART1_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"], base_size=10.5)

        render_options_no_table(doc, q["options"], is_teacher=is_teacher)

        if is_teacher and q["sol"]:
            add_solution_no_table(doc, q["sol"])

    # 2. PHẦN TRẮC NGHIỆM ĐÚNG SAI
    add_section_header(doc, "2. PHẦN TRẮC NGHIỆM ĐÚNG SAI", "Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn Đúng hoặc Sai:")
    for q in PART2_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"], base_size=10.5)

        for stmt, ans_str in q["stmts"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            if is_teacher:
                render_statement(p_s, stmt, ans=ans_str, is_teacher=True, base_size=10.5)
            else:
                render_rich_text(p_s, f"{stmt}  [   ] Đúng     [   ] Sai", base_size=10.5)

        if is_teacher and q["sol"]:
            add_solution_no_table(doc, q["sol"])

    # 3. PHẦN TRẮC NGHIỆM TRẢ LỜI NGẮN
    add_section_header(doc, "3. PHẦN TRẮC NGHIỆM TRẢ LỜI NGẮN", "Học sinh viết câu trả lời ngắn gọn vào chỗ trống:")
    p_c13 = doc.add_paragraph()
    p_c13.paragraph_format.space_before = Pt(3)
    p_c13.paragraph_format.space_after = Pt(2)
    render_rich_text(p_c13, "*Câu 13. Trả lời các câu hỏi sau:*", base_size=10.5, base_bold=True)

    for item in PART3_QUESTIONS:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.left_indent = Cm(0.5)
        p_item.paragraph_format.space_before = Pt(2)
        p_item.paragraph_format.space_after = Pt(1)
        render_rich_text(p_item, item["q"], base_size=10.5)

        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.8)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(2)
        if is_teacher:
            render_rich_text(p_ans, f"➔ *Đáp án:* {item['ans']}", base_size=10.5, base_color=COLOR_CORRECT)
            if item["sol"]:
                add_solution_no_table(doc, item["sol"])
        else:
            render_rich_text(p_ans, "Đáp án: ..................................................................................................", base_size=10.5)

    # II. PHẦN TỰ LUẬN
    add_section_header(doc, "II. PHẦN TỰ LUẬN (1,0 ĐIỂM)", "Trình bày chi tiết bài giải vào bài thi:")
    p_tl_head = doc.add_paragraph()
    p_tl_head.paragraph_format.space_before = Pt(4)
    p_tl_head.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_tl_head, f"*{ESSAY_QUESTION['title']}*", base_size=10.5)

    p_ctx = doc.add_paragraph()
    p_ctx.paragraph_format.left_indent = Cm(0.4)
    p_ctx.paragraph_format.space_before = Pt(1)
    p_ctx.paragraph_format.space_after = Pt(2)
    render_rich_text(p_ctx, ESSAY_QUESTION["context"], base_size=10.5)

    for sub_q in ESSAY_QUESTION["sub_qs"]:
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.left_indent = Cm(0.8)
        p_sub.paragraph_format.space_before = Pt(1)
        p_sub.paragraph_format.space_after = Pt(1)
        render_rich_text(p_sub, sub_q, base_size=10.5)

    if is_teacher:
        add_solution_no_table(doc, ESSAY_QUESTION["sol"])
    else:
        doc.add_paragraph()
        add_student_writing_lines(doc, num_lines=8)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(10)
    p_end.paragraph_format.space_after = Pt(4)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("--- HẾT ---")
    set_run_style(r_end, font_size=11, bold=True, italic=True, color=COLOR_PRIMARY)

    doc.save(filepath)
    print(f"[OK] Đã tạo thành công file Word: {filepath}")
    return filepath

# ==============================================================================
# TẠO FILE LATEX VÀ BIÊN DỊCH PDF TRÌNH CHIẾU TIVI 65 INCH (16:9)
# ==============================================================================

def generate_tivi_presentation_pdf():
    tex_filename = "De_Cuong_GK1_KHTN9_NguyenHoang_DayTrenManHinh.tex"
    pdf_filename = "De_Cuong_GK1_KHTN9_NguyenHoang_DayTrenManHinh.pdf"
    tex_path = os.path.join(SCRATCH_DIR, tex_filename)
    pdf_dest = os.path.join(OUTPUT_DIR, pdf_filename)

    tex_lines = []
    # Preamble chuẩn 16:9 widescreen không viền đen
    tex_lines.append(r"""\documentclass{article}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{tabularx}
\usepackage{array}
\usepackage{booktabs}
\usepackage{xcolor}
\usepackage[version=4]{mhchem}
\usepackage[
  paperwidth=29.7cm,
  paperheight=16.706cm,
  top=1.0cm,
  bottom=1.0cm,
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

    # SLIDE 1: TRANG TIÊU ĐỀ
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 0: TIÊU ĐỀ ĐỀ CƯƠNG
% ==============================================================================
\begin{center}
\vspace*{0.3cm}
{\fontsize{24}{30}\selectfont \textbf{\color{brandRed}TRƯỜNG THCS NGUYỄN HOÀNG -- NĂM HỌC 2026 -- 2027}}\\[12pt]
{\fontsize{34}{42}\selectfont \textbf{\color{brandBlue}ĐỀ CƯƠNG ÔN TẬP KIỂM TRA GIỮA KÌ I}}\\[10pt]
{\fontsize{28}{36}\selectfont \textbf{\color{brandGreen}MÔN: KHOA HỌC TỰ NHIÊN 9 (PHÂN MÔN HÓA HỌC)}}\\[20pt]
\rule{16cm}{1.5pt}\\[25pt]
{\fontsize{26}{34}\selectfont \textbf{Giáo viên hướng dẫn: Thầy TRẦN VĂN THẠNH -- 0777.470.803}}\\[10pt]
{\fontsize{22}{28}\selectfont \textit{Cơ sở 1: 6/15 Nguyễn Hoàng, Kim Long | Cơ sở 2: 24 Đặng Thái Thân, TP Huế}}\\[20pt]
{\fontsize{24}{30}\selectfont \textbf{CẤU TRÚC: TRẮC NGHIỆM ABCD -- ĐÚNG/SAI -- TRẢ LỜI NGẮN -- TỰ LUẬN}}
\end{center}
\newpage
""")

    # PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (2 câu / slide, font 28-30pt)
    # Slide 1: Câu 1 & 2
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 1: CÂU 1 & CÂU 2
% ==============================================================================
\fontsize{27}{36}\selectfont
\textbf{Câu 1.} Dãy chất nào sau đây đều là hydrocarbon?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\ce{FeCl2}$, $\ce{C2H6O}$, $\ce{CH4}$, $\ce{NaHCO3}$. & \textbf{B.} $\ce{C6H5Na}$, $\ce{CH4O}$, $\ce{HNO3}$, $\ce{C6H6}$.\\[8pt]
\textbf{C.} $\ce{CH4}$, $\ce{C2H4}$, $\ce{C2H2}$, $\ce{C6H6}$. & \textbf{D.} $\ce{CH3NO2}$, $\ce{CH3Br}$, $\ce{NaOH}$, $\ce{C2H5OH}$.
\end{tabularx}

\vspace{20pt}

\textbf{Câu 2.} Hợp chất hữu cơ là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} Hợp chất của oxygen với nguyên tố hóa học khác. & \textbf{B.} Hợp chất của nitrogen và phi kim khác.\\[8pt]
\textbf{C.} Hợp chất chỉ của carbon và hydrogen. & \textbf{D.} Hợp chất của carbon (trừ $\ce{CO}$, $\ce{CO2}$, $\ce{H2CO3}$,\dots).
\end{tabularx}
\newpage
""")

    # Slide 2: Câu 3 & 4
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 2: CÂU 3 & CÂU 4
% ==============================================================================
\fontsize{30}{39}\selectfont
\textbf{Câu 3.} Cho phương trình hóa học: $2\text{X} + 7\ce{O2} \xrightarrow{t^\circ} 4\ce{CO2} + 6\ce{H2O}$. Công thức phân tử của X là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X X X@{}}
\textbf{A.} $\ce{C2H2}$. & \textbf{B.} $\ce{C2H4}$. & \textbf{C.} $\ce{C2H6}$. & \textbf{D.} $\ce{C5H12}$.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 4.} Khi đốt cháy hoàn toàn $1{,}2395\text{ L}$ alkane X thu được $6{,}1975\text{ L}$ khí $\ce{CO2}$ (đkc). Công thức phân tử của X là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X X X@{}}
\textbf{A.} $\ce{C3H8}$. & \textbf{B.} $\ce{C5H10}$. & \textbf{C.} $\ce{C4H10}$. & \textbf{D.} $\ce{C5H12}$.
\end{tabularx}
\newpage
""")

    # Slide 3: Câu 5 & 6
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 3: CÂU 5 & CÂU 6
% ==============================================================================
\fontsize{30}{39}\selectfont
\textbf{Câu 5.} Trong các dãy chất sau, dãy chất nào là dẫn xuất của hydrocarbon?
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $\ce{CH4}$, $\ce{C2H4O}$, $\ce{C2H2}$. & \textbf{B.} $\ce{C2H2}$, $\ce{C2H4}$, $\ce{C6H6}$.\\[8pt]
\textbf{C.} $\ce{CH3Cl}$, $\ce{C2H6O}$, $\ce{C6H5NO2}$. & \textbf{D.} $\ce{C2H4O2}$, $\ce{C6H6Cl6}$, $\ce{C6H6}$.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 6.} Công thức phân tử của chất hữu cơ nào cũng nhất thiết phải chứa nguyên tố
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X X X@{}}
\textbf{A.} carbon. & \textbf{B.} oxygen. & \textbf{C.} nitrogen. & \textbf{D.} chlorine.
\end{tabularx}
\newpage
""")

    # Slide 4: Câu 7 & 8
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 4: CÂU 7 & CÂU 8
% ==============================================================================
\fontsize{30}{39}\selectfont
\textbf{Câu 7.} Đốt cháy $0{,}5\text{ mol}$ hydrocarbon X thu được $\ce{H2O}$ và $22\text{ g } \ce{CO2}$. Công thức phân tử của X là
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X X X@{}}
\textbf{A.} $\ce{CH4}$. & \textbf{B.} $\ce{C2H2}$. & \textbf{C.} $\ce{C2H4}$. & \textbf{D.} $\ce{C3H8}$.
\end{tabularx}

\vspace*{\fill}

\textbf{Câu 8.} Đốt cháy hoàn toàn $6{,}1975\text{ L}$ khí methane (đkc). Thể tích khí $\ce{O2}$ và không khí (đkc) cần dùng là (biết $\ce{O2}$ chiếm $20\%$ thể tích không khí)
\vspace{6pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
\textbf{A.} $6{,}1975\text{ L}$ và $30{,}9875\text{ L}$. & \textbf{B.} $12{,}395\text{ L}$ và $61{,}975\text{ L}$.\\[8pt]
\textbf{C.} $24{,}79\text{ L}$ và $123{,}95\text{ L}$. & \textbf{D.} $37{,}185\text{ L}$ và $185{,}925\text{ L}$.
\end{tabularx}
\newpage
""")

    # PHẦN II: TRẮC NGHIỆM ĐÚNG/SAI (1 câu / slide, font 26pt)
    # Slide 5: Câu 9
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 5: PHẦN II - CÂU 9
% ==============================================================================
\fontsize{26}{34}\selectfont
\textbf{Câu 9 (Đúng/Sai).} Công thức chung của alkene là $\ce{C_nH_{2n}}$ và alkene đơn giản nhất là $\ce{C2H4}$:
\vspace{14pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} Với $n = 1$ ta có một alkene với công thức $\ce{CH2}$. & $[\quad]$\\[12pt]
\textbf{b)} Tất cả các alkene đều làm mất màu vàng nâu của nước bromine. Phản ứng đặc trưng là phản ứng cộng. & $[\quad]$\\[12pt]
\textbf{c)} Cả alkane và alkene đều tham gia phản ứng trùng hợp. & $[\quad]$\\[12pt]
\textbf{d)} Khi đốt cháy một alkene ta thu được hai sản phẩm là hơi nước và khí carbon dioxide. & $[\quad]$
\end{tabularx}
\newpage
""")

    # Slide 6: Câu 10
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 6: PHẦN II - CÂU 10
% ==============================================================================
\fontsize{26}{34}\selectfont
\textbf{Câu 10 (Đúng/Sai).} Xét tính đúng -- sai của các phát biểu sau về dầu mỏ và khí thiên nhiên:
\vspace{14pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} Dầu mỏ và khí thiên nhiên là nhiên liệu hóa thạch có thành phần chủ yếu là khí methane. & $[\quad]$\\[12pt]
\textbf{b)} Khi chưng cất dầu mỏ ta thu được hai sản phẩm là xăng và dầu diesel. & $[\quad]$\\[12pt]
\textbf{c)} Dầu mỏ là hỗn hợp của các dẫn xuất của hydrocarbon. & $[\quad]$\\[12pt]
\textbf{d)} Mỏ dầu gồm ba lớp theo thứ tự từ trên xuống dưới là khí đồng hành, dầu lỏng và nước mặn. & $[\quad]$
\end{tabularx}
\newpage
""")

    # Slide 7: Câu 11
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 7: PHẦN II - CÂU 11
% ==============================================================================
\fontsize{26}{34}\selectfont
\textbf{Câu 11 (Đúng/Sai).} Xét tính đúng -- sai của các phát biểu sau về ethylene ($\ce{C2H4}$):
\vspace{14pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} Trong công thức cấu tạo của alkene có hai liên kết đôi $\ce{C=C}$. & $[\quad]$\\[12pt]
\textbf{b)} Trong liên kết đôi $\ce{C=C}$ có một liên kết kém bền, dễ bị đứt ra trong các phản ứng hóa học. & $[\quad]$\\[12pt]
\textbf{c)} Phản ứng giữa ethylene $\ce{C2H4}$ và nước bromine được gọi là phản ứng cộng, làm mất màu nước bromine. & $[\quad]$\\[12pt]
\textbf{d)} Ethylene là chất khí không màu, không mùi, hầu như không tan trong nước và nặng hơn không khí. & $[\quad]$
\end{tabularx}
\newpage
""")

    # Slide 8: Câu 12
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 8: PHẦN II - CÂU 12
% ==============================================================================
\fontsize{26}{34}\selectfont
\textbf{Câu 12 (Đúng/Sai).} Khi chưng cất dầu thô ta thu được các sản phẩm: khí hóa lỏng, dầu nhẹ, xăng, dầu hỏa, dầu diesel, dầu bôi trơn, sáp paraffin, nhựa đường:
\vspace{14pt}

\begin{tabularx}{\linewidth}{@{}p{0.88\linewidth} c@{}}
\textbf{a)} Xăng và dầu diesel là nhiên liệu dùng cho động cơ đốt trong (xe máy, ô tô,…). & $[\quad]$\\[12pt]
\textbf{b)} Nhựa đường dùng trong ngành xây dựng cầu, đường bộ,… & $[\quad]$\\[12pt]
\textbf{c)} Dầu hỏa chủ yếu được sử dụng để nấu đồ ăn, sưởi ấm. & $[\quad]$\\[12pt]
\textbf{d)} Khí hóa lỏng được sử dụng trong động cơ máy bay. & $[\quad]$
\end{tabularx}
\newpage
""")

    # PHẦN III: TRẢ LỜI NGẮN (Câu 13: chia 2 slide trình bày rõ ràng)
    # Slide 9: Câu 13 a, b, c, d
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 9: PHẦN III - CÂU 13 (Ý a, b, c, d)
% ==============================================================================
\fontsize{24}{32}\selectfont
\textbf{Câu 13. (Trả lời ngắn -- Phần 1)}
\vspace{8pt}

\textbf{a)} Acetic acid thuộc loại hợp chất hóa học nào có khả năng làm quỳ tím hóa đỏ?\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace{8pt}

\textbf{b)} Tên gọi thông thường của dung dịch acetic acid nồng độ $2\% - 5\%$ là gì?\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace{8pt}

\textbf{c)} Ngoài nước và khí $\ce{CO2}$, phản ứng giữa acetic acid với đá vôi ($\ce{CaCO3}$) tạo ra loại hợp chất gì?\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace{8pt}

\textbf{d)} Tên gọi thông thường của ethylic alcohol dùng để lên men giấm là gì?\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\newpage
""")

    # Slide 10: Câu 13 e, f, g
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 10: PHẦN III - CÂU 13 (Ý e, f, g)
% ==============================================================================
\fontsize{25}{34}\selectfont
\textbf{Câu 13. (Trả lời ngắn -- Phần 2)}
\vspace{10pt}

\textbf{e)} Vị đặc trưng của giấm ăn (acetic acid) khi dùng trong chế biến món ăn là gì?\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace{14pt}

\textbf{f)} Ngoài khí $\ce{CO2}$ và muối, phản ứng giữa giấm ăn và cặn đá vôi ở ấm đun nước sinh ra chất gì?\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\vspace{14pt}

\textbf{g)} Tên gọi thông thường của ethylic alcohol dùng để lên men giấm là gì?\\
\textbf{Trả lời:} $\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots\dots$
\newpage
""")

    # PHẦN II: TỰ LUẬN (Câu 14 - font 25pt, để trống phía dưới cho GV tương tác)
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 11: PHẦN TỰ LUẬN - CÂU 14
% ==============================================================================
\fontsize{25}{34}\selectfont
\textbf{Câu 14. (Tự luận -- 1,0 điểm)}

\textit{Trong đời sống hằng ngày, ethylic alcohol có rất nhiều ứng dụng quan trọng. Xăng sinh học E5 chứa $5\%$ thể tích cồn nguyên chất và $95\%$ xăng truyền thống. Trong y tế, dung dịch cồn $70^\circ$ được dùng để diệt khuẩn, sát trùng vết thương.}
\vspace{10pt}

\textbf{a)} Viết công thức cấu tạo của ethylic alcohol.\\[6pt]
\textbf{b)} Hãy cho biết ý nghĩa của cồn $60^\circ$.\\[6pt]
\textbf{c)} Viết phương trình hóa học xảy ra khi cho viên Na vào dung dịch $\ce{C2H5OH}\ 60^\circ$.
\vspace*{\fill}

\begin{center}
\textit{\color{gray}(Khoảng trống dành cho Thầy / Cô ghi chú và hướng dẫn học sinh giải bài)}
\end{center}
\vspace*{1.5cm}
\newpage
""")

    # SLIDE 12: BẢNG ĐÁP ÁN CHÍNH THỨC
    tex_lines.append(r"""
% ==============================================================================
% SLIDE 12: BẢNG ĐÁP ÁN CHÍNH THỨC
% ==============================================================================
\fontsize{18}{23}\selectfont
\begin{center}
\textbf{\color{brandRed}BẢNG ĐÁP ÁN CHÍNH THỨC ĐỀ CƯƠNG ÔN TẬP GIỮA KÌ I}
\end{center}
\vspace{2pt}

\textbf{\color{brandBlue}1. PHẦN TRẮC NGHIỆM NHIỀU LỰA CHỌN (Mỗi câu đúng 0,25 điểm):}
\vspace{1pt}

\begin{center}
\begin{tabular}{|c|c|c|c|c|c|c|c|}
\hline
\textbf{Câu 1} & \textbf{Câu 2} & \textbf{Câu 3} & \textbf{Câu 4} & \textbf{Câu 5} & \textbf{Câu 6} & \textbf{Câu 7} & \textbf{Câu 8} \\
\hline
\textbf{\color{brandRed}C} & \textbf{\color{brandRed}D} & \textbf{\color{brandRed}C} & \textbf{\color{brandRed}D} & \textbf{\color{brandRed}C} & \textbf{\color{brandRed}A} & \textbf{\color{brandRed}A} & \textbf{\color{brandRed}B} \\
\hline
\end{tabular}
\end{center}
\vspace{3pt}

\textbf{\color{brandBlue}2. PHẦN TRẮC NGHIỆM ĐÚNG -- SAI:}
\vspace{1pt}

\begin{center}
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu hỏi} & \textbf{Ý a} & \textbf{Ý b} & \textbf{Ý c} & \textbf{Ý d} \\
\hline
\textbf{Câu 9} & \textbf{\color{brandRed}Sai} & \textbf{\color{brandGreen}Đúng} & \textbf{\color{brandRed}Sai} & \textbf{\color{brandGreen}Đúng} \\
\hline
\textbf{Câu 10} & \textbf{\color{brandRed}Sai} & \textbf{\color{brandRed}Sai} & \textbf{\color{brandRed}Sai} & \textbf{\color{brandGreen}Đúng} \\
\hline
\textbf{Câu 11} & \textbf{\color{brandRed}Sai} & \textbf{\color{brandGreen}Đúng} & \textbf{\color{brandGreen}Đúng} & \textbf{\color{brandRed}Sai} \\
\hline
\textbf{Câu 12} & \textbf{\color{brandGreen}Đúng} & \textbf{\color{brandGreen}Đúng} & \textbf{\color{brandGreen}Đúng} & \textbf{\color{brandRed}Sai} \\
\hline
\end{tabular}
\end{center}
\vspace{3pt}

\textbf{\color{brandBlue}3. ĐÁP ÁN PHẦN TRẢ LỜI NGẮN (CÂU 13):}
\vspace{1pt}

\begin{tabularx}{\linewidth}{@{}X X@{}}
• a) \textbf{\color{brandRed}Acid (Axit)} & • b) \textbf{\color{brandRed}Giấm ăn (Giấm)}\\
• c) \textbf{\color{brandRed}Muối (Calcium acetate)} & • d) \textbf{\color{brandRed}Rượu (Cồn)}\\
• e) \textbf{\color{brandRed}Vị chua} & • f) \textbf{\color{brandRed}Nước ($\ce{H2O}$)}\\
• g) \textbf{\color{brandRed}Rượu (Cồn)} &
\end{tabularx}

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
        print(f"[OK] Đã biên dịch và copy PDF thành công: {pdf_dest}")
    else:
        print("[ERROR] Không tìm thấy file PDF được tạo. Log:")
        print(res2.stdout[-1500:])

if __name__ == "__main__":
    print("--- BẮT ĐẦU TẠO 3 FILE THEO QUY CHUẨN ---")
    f_hs = generate_word_document(is_teacher=False)
    f_gv = generate_word_document(is_teacher=True)
    generate_tivi_presentation_pdf()
    print("--- HOÀN THÀNH TẤT CẢ 3 FILE ---")
