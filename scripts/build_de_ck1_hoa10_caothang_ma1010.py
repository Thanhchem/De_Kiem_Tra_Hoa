# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Cuối Học Kì I - Hóa học 10 - Trường THPT Cao Thắng (TP Huế)
Mã đề: 1010 - Năm học 2025 - 2026
Chuẩn theo Quy chuẩn Form Đề Kiểm Tra của Thầy TRẦN VĂN THẠNH
Tạo 2 phiên bản: Bản Học Sinh & Bản Giáo Viên (có đáp án đậm đỏ & lời giải chi tiết)
"""

import os
import sys
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Import module form_de_kiem_tra
sys.path.append(os.path.join(os.path.dirname(__file__)))
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
    COLOR_CORRECT,
    COLOR_SOLUTION
)

DIR_SAN_PHAM = os.path.abspath("San_Pham")
os.makedirs(DIR_SAN_PHAM, exist_ok=True)

# Dữ liệu Đề thi mã 1010
DATA_EXAM = {
    "school_line": "SỞ GIÁO DỤC VÀ ĐÀO TẠO - TRƯỜNG THPT CAO THẮNG",
    "exam_title": "ĐỀ KIỂM TRA CUỐI KÌ 1 - NĂM HỌC 2025 - 2026",
    "subject_line": "MÔN: HÓA HỌC 10",
    "duration": "45 phút (không kể thời gian phát đề)",
    "num_pages": "02",
    "exam_code": "1010",
    "atomic_masses": "Mg = 24; P = 31; O = 16; S = 32; H = 1; Na = 23; K = 39 (Số hiệu nguyên tử: Na=11; Cl=17; H=1; C=6; O=8; Mg=12; P=15; Be=4; Al=13; S=16; K=19; N=14; F=9)",

    # I. TRẮC NGHIỆM (7,0 ĐIỂM)
    # PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. (12 câu - 3,0 điểm)
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1. Để đạt quy tắc Octet, nguyên tử của nguyên tố sodium (Z=11) có xu hướng",
            "options": [
                "*A. nhường 1 electron.*",
                "B. nhận 1 electron.",
                "C. nhận 7 electron.",
                "D. nhường 3 electron."
            ],
            "solution": "Cấu hình electron của nguyên tử sodium (Z=11) là 1s~2~ 2s~2~ 2p~6~ 3s~1~ ([Ne]3s~1~), có 1 electron ở lớp ngoài cùng. Để đạt cấu hình electron bền vững của khí hiếm neon (8 electron ở lớp ngoài cùng) theo quy tắc Octet, nguyên tử sodium có xu hướng nhường 1 electron ở phân lớp 3s~1~ để tạo thành cation Na^+^ (1s~2~ 2s~2~ 2p~6~).\n*Đáp án đúng: A.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Nguyên tố X thuộc nhóm IVA trong Bảng tuần hoàn. Công thức oxide cao nhất của X là",
            "options": [
                "A. X~2~O~4~",
                "B. X~2~O",
                "C. XO",
                "*D. XO~2~*"
            ],
            "solution": "Nguyên tố X thuộc nhóm IVA nên có hóa trị cao nhất với oxygen bằng số thứ tự nhóm là IV. Do đó, công thức oxide cao nhất của X là XO~2~ (dạng tối giản từ X~2~O~4~).\n*Đáp án đúng: D.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Nguyên tử có cấu tạo gồm hạt nhân và vỏ nguyên tử. Vỏ nguyên tử chứa loại hạt nào sau đây?",
            "options": [
                "A. Neutron",
                "*B. Electron*",
                "C. Proton",
                "D. Proton và neutron"
            ],
            "solution": "Nguyên tử có cấu tạo gồm:\n- Hạt nhân nằm ở tâm nguyên tử, gồm các hạt proton (mang điện tích dương) và neutron (không mang điện).\n- Vỏ nguyên tử được tạo nên bởi các hạt electron (mang điện tích âm) chuyển động xung quanh hạt nhân.\n*Đáp án đúng: B.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Đại lượng nào sau đây trong nguyên tử của các nguyên tố nhóm A biến đổi tuần hoàn theo chiều tăng của điện tích hạt nhân nguyên tử?",
            "options": [
                "A. Số electron trong nguyên tử",
                "B. Nguyên tử khối",
                "*C. Số electron ở lớp ngoài cùng*",
                "D. Số lớp electron"
            ],
            "solution": "- Số electron trong nguyên tử và nguyên tử khối tăng dần liên tục theo chiều tăng của điện tích hạt nhân, không biến đổi tuần hoàn.\n- Số lớp electron giữ nguyên trong một chu kì và tăng khi sang chu kì mới, không biến đổi tuần hoàn.\n- Số electron ở lớp ngoài cùng của các nguyên tố nhóm A biến đổi tuần hoàn lặp đi lặp lại từ 1 đến 8 (riêng chu kì 1 từ 1 đến 2) theo chiều tăng của điện tích hạt nhân nguyên tử.\n*Đáp án đúng: C.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5. Nguyên tố X thuộc nhóm VA trong Bảng tuần hoàn. X thuộc khối nguyên tố",
            "options": [
                "A. s",
                "B. d",
                "*C. p*",
                "D. f"
            ],
            "solution": "Các nguyên tố nhóm A gồm các nguyên tố khối s (nhóm IA, IIA) và khối p (nhóm IIIA đến VIIIA). Nguyên tố X thuộc nhóm VA có cấu hình electron lớp ngoài cùng dạng ns~2~ np~3~, electron cuối cùng được điền vào phân lớp p nên X thuộc khối nguyên tố p.\n*Đáp án đúng: C.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6. Lớp L có bao nhiêu phân lớp electron?",
            "options": [
                "A. 4",
                "*B. 2*",
                "C. 1",
                "D. 3"
            ],
            "solution": "Lớp L là lớp electron thứ hai (n = 2). Lớp thứ n có n phân lớp electron (với n ≤ 4), do đó lớp L có 2 phân lớp electron là phân lớp 2s và phân lớp 2p.\n*Đáp án đúng: B.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7. Trong một chu kì, theo chiều tăng của điện tích hạt nhân nguyên tử,",
            "options": [
                "A. bán kính nguyên tử và độ âm điện đều tăng.",
                "B. bán kính nguyên tử tăng, độ âm điện giảm.",
                "C. bán kính nguyên tử và độ âm điện đều giảm.",
                "*D. bán kính nguyên tử giảm, độ âm điện tăng.*"
            ],
            "solution": "Trong một chu kì, đi từ trái sang phải theo chiều tăng của điện tích hạt nhân:\n- Số lớp electron không đổi nhưng điện tích hạt nhân tăng làm tăng lực hút tĩnh điện giữa hạt nhân với các electron lớp ngoài cùng, kéo các electron lại gần hạt nhân hơn ⇒ Bán kính nguyên tử giảm dần.\n- Khả năng hút electron liên kết về phía nguyên tử tăng dần ⇒ Độ âm điện tăng dần.\n*Đáp án đúng: D.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8. Cho độ âm điện của các nguyên tố sau: O: 3,44; Cl: 3,16; Al: 1,61; F: 3,98. Nguyên tố nào có tính phi kim mạnh nhất?",
            "options": [
                "A. O",
                "B. Cl",
                "*C. F*",
                "D. Al"
            ],
            "solution": "Độ âm điện đặc trưng cho khả năng hút electron của nguyên tử trong phân tử. Độ âm điện càng lớn thì tính phi kim của nguyên tố càng mạnh.\nSo sánh độ âm điện: χ(F) = 3,98 > χ(O) = 3,44 > χ(Cl) = 3,16 > χ(Al) = 1,61.\nDo đó, Fluorine (F) là nguyên tố có tính phi kim mạnh nhất.\n*Đáp án đúng: C.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9. Phân tử O~2~ có công thức Lewis là :Ö = Ö: . Giữa 2 nguyên tử O có bao nhiêu cặp electron dùng chung?",
            "options": [
                "A. 6",
                "B. 1",
                "C. 4",
                "*D. 2*"
            ],
            "solution": "Mỗi nguyên tử O (Z=8, nhóm VIA) có 6 electron lớp ngoài cùng, cần thêm 2 electron để đạt cấu hình bát tử (8 electron). Do đó, hai nguyên tử O góp chung 2 cặp electron hình thành liên kết đôi O=O. Vậy giữa 2 nguyên tử O có 2 cặp electron dùng chung (tương ứng với 4 electron liên kết).\n*Đáp án đúng: D.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10. Hợp chất nào sau đây là hợp chất ion?",
            "options": [
                "A. Cl~2~",
                "B. HCl",
                "C. CO~2~",
                "*D. NaCl*"
            ],
            "solution": "- Cl~2~ là đơn chất phi kim, có liên kết cộng hóa trị không phân cực.\n- HCl và CO~2~ tạo bởi các phi kim, có liên kết cộng hóa trị phân cực.\n- NaCl được tạo thành bởi kim loại điển hình (Na) và phi kim điển hình (Cl), có hiệu độ âm điện Δχ = 3,16 - 0,93 = 2,23 > 1,7 nên liên kết trong NaCl là liên kết ion và NaCl là hợp chất ion.\n*Đáp án đúng: D.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11. Nguyên tố hóa học là tập hợp những nguyên tử có cùng",
            "options": [
                "*A. điện tích hạt nhân*",
                "B. nguyên tử khối",
                "C. số neutron",
                "D. số khối"
            ],
            "solution": "Theo định nghĩa chuẩn: Nguyên tố hóa học là tập hợp những nguyên tử có cùng điện tích hạt nhân (có cùng số proton Z trong hạt nhân).\n*Đáp án đúng: A.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12. Số nguyên tố có trong chu kì 3 của Bảng tuần hoàn là",
            "options": [
                "*A. 8*",
                "B. 32",
                "C. 18",
                "D. 2"
            ],
            "solution": "Chu kì 3 là chu kì nhỏ, bắt đầu từ nguyên tố sodium (~11~Na: [Ne]3s~1~) đến nguyên tố argon (~18~Ar: [Ne]3s~2~ 3p~6~), bao gồm 8 nguyên tố: Na, Mg, Al, Si, P, S, Cl, Ar.\n*Đáp án đúng: A.*"
        }
    ],

    # PHẦN II. Câu trắc nghiệm đúng sai. (2 câu - 2,0 điểm)
    "part2": [
        {
            "id": 1,
            "prompt": "Câu 1. Cho C (Z=6) và O (Z=8); Độ âm điện của C là 2,55 và O là 3,44.",
            "statements": [
                ("a) Liên kết hóa học trong phân tử CO là liên kết cộng hóa trị phân cực, trong đó 3 cặp electron dùng chung đều do 2 nguyên tử C và O cùng đóng góp bị lệch về phía nguyên tử O.", "S"),
                ("b) C và O đều là các nguyên tố p.", "Đ"),
                ("c) Liên kết hóa học trong phân tử O~2~ là liên kết cộng hóa trị không phân cực.", "Đ"),
                ("d) Phân tử CO~2~ là phân tử phân cực do liên kết hóa học trong phân tử CO~2~ là liên kết cộng hóa trị phân cực.", "S")
            ],
            "solution": "• a) Sai. Phân tử CO có liên kết ba C≡O gồm 2 liên kết cộng hóa trị bình thường (mỗi nguyên tử góp 1e tạo thành cặp e dùng chung) và 1 liên kết cho - nhận (phối trí) trong đó cặp electron dùng chung do một mình nguyên tử O đóng góp vào orbital trống của C (C ← O). Do đó nói \"3 cặp electron dùng chung đều do 2 nguyên tử C và O cùng đóng góp\" là sai.\n• b) Đúng. Cấu hình electron: C (Z=6): 1s~2~ 2s~2~ 2p~2~; O (Z=8): 1s~2~ 2s~2~ 2p~4~. Cả C và O đều có electron cuối cùng điền vào phân lớp p nên đều là các nguyên tố p.\n• c) Đúng. Phân tử O~2~ tạo bởi 2 nguyên tử oxygen giống nhau, có hiệu độ âm điện Δχ = 0 nên liên kết O=O là liên kết cộng hóa trị không phân cực.\n• d) Sai. Trong phân tử CO~2~ (O=C=O), liên kết C=O là liên kết cộng hóa trị phân cực (Δχ = 3,44 - 2,55 = 0,89 > 0,4), tuy nhiên phân tử CO~2~ có cấu tạo dạng đường thẳng đối xứng nên hai vectơ mômen lưỡng cực của 2 liên kết C=O cùng phương, ngược chiều triệt tiêu lẫn nhau (μ = 0), do đó phân tử CO~2~ là phân tử không phân cực."
        },
        {
            "id": 2,
            "prompt": "Câu 2. Cho X (Z=12); Y (Z=15)",
            "statements": [
                ("a) Hydroxide ứng với hóa trị cao nhất của X có dạng X(OH)~2~ và có tính base. Hydroxide ứng với hóa trị cao nhất của Y có dạng HYO~4~ và có tính acid.", "S"),
                ("b) X là kim loại, Y là phi kim.", "Đ"),
                ("c) X và Y là 2 nguyên tố thuộc cùng 1 chu kì.", "Đ"),
                ("d) Trong công thức oxide cao nhất của Y có chứa thành phần phần trăm theo khối lượng của Y là 46,36%.", "S")
            ],
            "solution": "Cấu hình electron nguyên tử:\n- X (Z=12, Mg): 1s~2~ 2s~2~ 2p~6~ 3s~2~ (ô số 12, chu kì 3, nhóm IIA).\n- Y (Z=15, P): 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~3~ (ô số 15, chu kì 3, nhóm VA).\n• a) Sai. Hydroxide cao nhất của X (Mg) là Mg(OH)~2~ có dạng X(OH)~2~ mang tính base. Còn hydroxide cao nhất của Y (P, nhóm VA) là phosphoric acid H~3~PO~4~ có dạng H~3~YO~4~ (không phải HYO~4~, dạng HYO~4~ ứng với nhóm VIIA như HClO~4~).\n• b) Đúng. X có 2 electron ở lớp ngoài cùng là kim loại; Y có 5 electron ở lớp ngoài cùng là phi kim.\n• c) Đúng. Cả X và Y đều có 3 lớp electron nên đều thuộc cùng chu kì 3 trong bảng tuần hoàn.\n• d) Sai. Công thức oxide cao nhất của Y (P) là Y~2~O~5~ (P~2~O~5~). Phần trăm theo khối lượng của Y trong oxide cao nhất là:\n  %m(Y) = [2 × 31 / (2 × 31 + 5 × 16)] × 100% = (62 / 142) × 100% ≈ 43,66% (khác 46,36%)."
        }
    ],

    # PHẦN III. Câu trắc nghiệm trả lời ngắn. (4 câu - 2,0 điểm)
    "part3": [
        {
            "id": 1,
            "prompt": "Câu 1. Cho các nguyên tố được đánh số thứ tự (1), (2), (3), (4) cùng số hiệu nguyên tử của các nguyên tố đó như sau: (1) (Z=11); (2) (Z=19); (3) (Z=13); (4) (Z=17). Hãy sắp xếp các nguyên tố trên theo chiều tăng dần tính kim loại. Biểu diễn kết quả bằng số thứ tự ứng với các nguyên tố (ví dụ: 1234, 3241,...).",
            "answer": "4312",
            "solution": "Xác định vị trí các nguyên tố trong bảng tuần hoàn:\n- (1) Na (Z=11): [Ne]3s~1~ (chu kì 3, nhóm IA)\n- (2) K (Z=19): [Ar]4s~1~ (chu kì 4, nhóm IA)\n- (3) Al (Z=13): [Ne]3s~2~ 3p~1~ (chu kì 3, nhóm IIIA)\n- (4) Cl (Z=17): [Ne]3s~2~ 3p~5~ (chu kì 3, nhóm VIIA)\nQuy luật biến đổi tính kim loại:\n- Trong cùng chu kì 3: theo chiều tăng của điện tích hạt nhân, tính kim loại giảm dần: Na > Al > Cl (Cl là phi kim có tính kim loại yếu nhất).\n- Trong cùng nhóm IA: theo chiều tăng của điện tích hạt nhân, tính kim loại tăng dần: K > Na.\n⇒ Chiều tăng dần tính kim loại: Cl < Al < Na < K.\nBiểu diễn theo số thứ tự: (4) < (3) < (1) < (2) ➔ Kết quả: 4312.\n*Đáp án: 4312*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Cho 2 nguyên tố: oxygen (Z = 8) và phosphorus (Z = 15). Số electron có trong anion PO~4~^3-^ là bao nhiêu?",
            "answer": "50",
            "solution": "- Trong mỗi nguyên tử trung hòa: số electron bằng số hiệu nguyên tử Z.\n  + Nguyên tử P (Z=15) có 15 electron.\n  + Mỗi nguyên tử O (Z=8) có 8 electron, trong nhóm PO~4~ có 4 nguyên tử O chứa 4 × 8 = 32 electron.\n- Anion PO~4~^3-^ mang điện tích 3- tức là phân tử đã nhận thêm 3 electron.\n⇒ Tổng số electron trong anion PO~4~^3-^ là:\n  N(e) = Z(P) + 4 × Z(O) + 3 = 15 + 4 × 8 + 3 = 50 electron.\n*Đáp án: 50*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Trong tự nhiên, magnesium có 3 đồng vị bền là ^24^Mg, ^25^Mg và ^26^Mg. Phương pháp phổ khối lượng xác nhận đồng vị ^26^Mg chiếm tỉ lệ phần trăm số nguyên tử là 11%. Biết rằng nguyên tử khối trung bình của Mg là 24,32. Phần trăm số nguyên tử của đồng vị ^25^Mg là bao nhiêu?",
            "answer": "10",
            "solution": "Gọi x là phần trăm số nguyên tử của đồng vị ^25^Mg (0 < x < 89).\nVì tổng phần trăm của 3 đồng vị bằng 100% nên phần trăm của đồng vị ^24^Mg là:\n  %^24^Mg = 100% - 11% - x = (89 - x)%.\nÁp dụng công thức tính nguyên tử khối trung bình của Mg:\n  Ā = [24 × (89 - x) + 25 × x + 26 × 11] / 100 = 24,32\n  ⇔ 2136 - 24x + 25x + 286 = 2432\n  ⇔ 2422 + x = 2432\n  ⇔ x = 10 (%).\nVậy phần trăm số nguyên tử của đồng vị ^25^Mg là 10%.\n*Đáp án: 10*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Cho các phân tử sau: BeH~2~, AlCl~3~, PCl~5~, HCl. Có bao nhiêu phân tử có các nguyên tử đều đạt cấu hình electron bão hòa theo quy tắc Octet?",
            "answer": "1",
            "solution": "Xét cấu hình electron hóa trị xung quanh các nguyên tử trong từng phân tử:\n- BeH~2~: Nguyên tử Be ở tâm chỉ có 2 cặp electron liên kết xung quanh (tổng 4 electron hóa trị), chưa đạt bát tử (trường hợp thiếu octet).\n- AlCl~3~: Nguyên tử Al ở tâm chỉ có 3 cặp electron liên kết xung quanh (tổng 6 electron hóa trị), chưa đạt bát tử (trường hợp thiếu octet).\n- PCl~5~: Nguyên tử P ở tâm có 5 cặp electron liên kết xung quanh (tổng 10 electron hóa trị), vượt quá bát tử (trường hợp mở rộng octet).\n- HCl: Nguyên tử Cl có 1 cặp electron liên kết và 3 cặp electron riêng (tổng 8 electron lớp ngoài cùng, thỏa mãn quy tắc octet). Nguyên tử H dùng chung 1 cặp electron đạt cấu hình bền vững của khí hiếm He (2 electron, bão hòa lớp vỏ 1 theo quy tắc đuplet/octet đối với chu kì 1).\nDo đó, chỉ có 1 phân tử (HCl) có các nguyên tử đều đạt cấu hình electron bão hòa theo quy tắc octet.\n*Đáp án: 1*"
        }
    ],

    # II. TỰ LUẬN (3,0 ĐIỂM)
    "part4": [
        {
            "id": 1,
            "prompt": "Câu 1. (1,0 điểm)\nX là nguyên tố phổ biến thứ 17 trong vỏ Trái Đất, tồn tại ở bốn dạng đồng vị bền. Trong cơ thể người, X chiếm 0,2% khối lượng và là thành phần quan trọng của nhiều protein và enzyme. Nguyên tử của nguyên tố X có cấu hình electron lớp ngoài cùng là [Ne]3s~2~3p~4~. Hãy xác định vị trí của nguyên tố này trong bảng tuần hoàn (có giải thích) và cho biết tính chất của nó.",
            "num_lines": 8,
            "solution": "1. Xác định vị trí của nguyên tố X trong bảng tuần hoàn:\n- Cấu hình electron nguyên tử của X là [Ne]3s~2~3p~4~ hay viết đầy đủ: 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~4~.\n- Tổng số electron của X là 16 ⇒ Số hiệu nguyên tử Z = 16 (nguyên tố Sulfur, S).\n- Vị trí của X trong bảng tuần hoàn:\n  + Ô số 16: Vì số hiệu nguyên tử Z = 16 (nguyên tử có 16 proton và 16 electron).\n  + Chu kì 3: Vì nguyên tử có 3 lớp electron (lớp 1, 2, 3).\n  + Nhóm VIA: Vì X là nguyên tố khối p (electron cuối cùng điền vào phân lớp p) và có 6 electron ở lớp ngoài cùng (3s~2~ 3p~4~).\n\n2. Tính chất cơ bản của nguyên tố X:\n- X là một phi kim điển hình.\n- Hóa trị cao nhất với oxygen là VI; công thức oxide cao nhất là SO~3~ (sulfur trioxide), là một acidic oxide (oxide acid).\n- Hydroxide ứng với hóa trị cao nhất là H~2~SO~4~ (sulfuric acid), là một acid mạnh.\n- Hợp chất khí với hydrogen là H~2~S (hydrogen sulfide)."
        },
        {
            "id": 2,
            "prompt": "Câu 2. (1,0 điểm)\nNguyên tố X tích lũy trong các tế bào thực vật nên rau và trái cây tươi là nguồn cung cấp tốt nguyên tố X cho cơ thể. Các nghiên cứu chỉ ra khẩu phần ăn chứa nhiều X có thể giảm nguy cơ cao huyết áp và đột quỵ. Y là nguyên tố hóa học có trong thành phần của chất có tác dụng oxi hóa và sát khuẩn cực mạnh, thường được sử dụng với mục đích khử trùng và tẩy trắng trong lĩnh vực thủy sản, dệt nhuộm, xử lí nước cấp, nước thải, nước bể bơi. Nguyên tử của nguyên tố X có 7 electron trên các phân lớp s; còn nguyên tử của nguyên tố Y có 3 lớp electron và có 7 electron ở lớp ngoài cùng.\na. Phân tử tạo thành bởi X và Y có công thức là gì?\nb. Xác định loại liên kết và biểu diễn sự hình thành liên kết trong phân tử được tạo thành bởi X và Y.",
            "num_lines": 10,
            "solution": "1. Xác định nguyên tố X và Y:\n- Nguyên tố X: Nguyên tử X có 7 electron trên các phân lớp s, do đó cấu hình electron ở các phân lớp s lần lượt là 1s~2~ 2s~2~ 3s~2~ 4s~1~ (2 + 2 + 2 + 1 = 7 electron s).\n  ⇒ Cấu hình electron đầy đủ của X: 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~6~ 4s~1~ ([Ar]4s~1~).\n  ⇒ Z = 19, X là Potassium (kí hiệu: K).\n- Nguyên tố Y: Nguyên tử Y có 3 lớp electron và có 7 electron ở lớp ngoài cùng (lớp 3).\n  ⇒ Cấu hình electron lớp ngoài cùng là 3s~2~ 3p~5~.\n  ⇒ Cấu hình electron đầy đủ của Y: 1s~2~ 2s~2~ 2p~6~ 3s~2~ 3p~5~ ([Ne]3s~2~ 3p~5~).\n  ⇒ Z = 17, Y là Chlorine (kí hiệu: Cl).\n\na. Công thức phân tử tạo thành bởi X và Y:\n- Potassium (K) thuộc nhóm IA có hóa trị I, Chlorine (Cl) kết hợp với kim loại có hóa trị I.\n- Công thức hóa học của phân tử tạo bởi X và Y là KCl (Potassium chloride).\n\nb. Loại liên kết và biểu diễn sự hình thành liên kết:\n- Loại liên kết: Liên kết ion (do chênh lệch độ âm điện: Δχ = χ(Cl) - χ(K) = 3,16 - 0,82 = 2,34 > 1,7; liên kết được hình thành giữa kim loại điển hình và phi kim điển hình).\n- Biểu diễn sự hình thành liên kết:\n  + Nguyên tử K nhường 1 electron ở lớp ngoài cùng để đạt cấu hình electron bền vững của khí hiếm Ar, tạo thành cation potassium:\n    K → K^+^ + 1e   (hoặc [Ar]4s~1~ → [Ar] + 1e)\n  + Nguyên tử Cl nhận 1 electron vào lớp ngoài cùng để đạt cấu hình electron bền vững của khí hiếm Ar, tạo thành anion chloride:\n    Cl + 1e → Cl^-^   (hoặc [Ne]3s~2~ 3p~5~ + 1e → [Ar])\n  + Cation K^+^ và anion Cl^-^ mang điện tích trái dấu hút nhau bằng lực hút tĩnh điện hình thành liên kết ion trong hợp chất KCl:\n    K^+^ + Cl^-^ → KCl."
        },
        {
            "id": 3,
            "prompt": "Câu 3. (1,0 điểm)\nX và Y là hai nguyên tố thuộc chu kì nhỏ, ở hai nhóm A kế tiếp nhau trong bảng tuần hoàn. Ở trạng thái đơn chất, X và Y phản ứng được với nhau. Tổng số proton trong hạt nhân nguyên tử của X và Y là 23. Biết rằng X đứng sau Y trong bảng tuần hoàn. Xác định tên nguyên tố X và Y.",
            "num_lines": 10,
            "solution": "Chu kì nhỏ trong bảng tuần hoàn gồm chu kì 1, chu kì 2 và chu kì 3 (số hiệu nguyên tử Z từ 1 đến 18).\nTheo giả thiết:\n- X đứng sau Y trong bảng tuần hoàn: Z_X > Z_Y.\n- Tổng số proton: Z_X + Z_Y = 23 (1).\n- X và Y thuộc 2 nhóm A kế tiếp nhau và ở trạng thái đơn chất chúng phản ứng được với nhau.\n\nBiện luận các trường hợp:\n• Trường hợp 1: X và Y cùng thuộc một chu kì nhỏ.\nVì X và Y ở hai nhóm A kế tiếp nhau và Z_X > Z_Y nên:\nZ_X - Z_Y = 1 (2).\nTừ (1) và (2) giải hệ phương trình được: Z_X = 12 (Magnesium, Mg) và Z_Y = 11 (Sodium, Na).\nKiểm tra: Na (nhóm IA, chu kì 3) và Mg (nhóm IIA, chu kì 3). Tuy nhiên, ở trạng thái đơn chất, hai kim loại Na và Mg không phản ứng hóa học với nhau. Do đó trường hợp này loại.\n\n• Trường hợp 2: X và Y thuộc hai chu kì nhỏ khác nhau (chu kì 2 và chu kì 3).\nVì Z_X + Z_Y = 23 và Z_X > Z_Y nên Z_Y < 11,5 (Y thuộc chu kì 2) và Z_X > 11,5 (X thuộc chu kì 3).\nDo X và Y thuộc 2 nhóm A kế tiếp nhau, có 2 khả năng:\n  - Khả năng 2a: Nhóm của X kế tiếp sau nhóm của Y (nhóm X = nhóm Y + 1):\n    Khoảng cách số hiệu: Z_X = Z_Y + 8 + 1 = Z_Y + 9 ⇒ Z_X - Z_Y = 9.\n    Kết hợp Z_X + Z_Y = 23 ⇒ Z_X = 16 (Sulfur, S) và Z_Y = 7 (Nitrogen, N).\n    (Đơn chất N~2~ và S rất bền, không phản ứng trực tiếp với nhau ở điều kiện thông thường).\n  - Khả năng 2b: Nhóm của X kế tiếp trước nhóm của Y (nhóm X = nhóm Y - 1):\n    Khoảng cách số hiệu: Z_X = Z_Y + 8 - 1 = Z_Y + 7 ⇒ Z_X - Z_Y = 7.\n    Kết hợp Z_X + Z_Y = 23 ⇒ Z_X = 15 (Phosphorus, P) và Z_Y = 8 (Oxygen, O).\n    Kiểm tra:\n    + Y là Oxygen (O, Z = 8): chu kì 2, nhóm VIA.\n    + X là Phosphorus (P, Z = 15): chu kì 3, nhóm VA.\n    + O và P thuộc 2 nhóm A kế tiếp nhau (nhóm VA và nhóm VIA) và đều thuộc chu kì nhỏ.\n    + Z_X (15) > Z_Y (8) nên X đứng sau Y trong bảng tuần hoàn.\n    + Ở trạng thái đơn chất, phosphorus và oxygen phản ứng mãnh liệt với nhau tạo thành oxide:\n      4P + 3O~2~ (thiếu) --t°--> 2P~2~O~3~\n      4P + 5O~2~ (dư) --t°--> 2P~2~O~5~\n    ⇒ Hoàn toàn thỏa mãn tất cả các điều kiện của đề bài!\n\n* Kết luận:\n- Nguyên tố X là Phosphorus (kí hiệu: P, Z = 15).\n- Nguyên tố Y là Oxygen (kí hiệu: O, Z = 8).\n\n(Lưu ý: Nếu học sinh giải theo số học thuần túy giả định cùng chu kì ra Na và Mg nhưng bỏ qua dữ kiện đơn chất phản ứng với nhau thì giáo viên xem xét cho điểm linh hoạt theo thang điểm)."
        }
    ]
}

def generate_docx(is_teacher=False):
    doc = Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line=DATA_EXAM["school_line"],
        exam_title=DATA_EXAM["exam_title"],
        subject_line=DATA_EXAM["subject_line"],
        num_pages=DATA_EXAM["num_pages"],
        duration=DATA_EXAM["duration"],
        exam_code=DATA_EXAM["exam_code"],
        atomic_masses=DATA_EXAM["atomic_masses"],
        is_teacher=is_teacher
    )

    # I. TRẮC NGHIỆM (7,0 ĐIỂM)
    add_section_header(
        doc,
        title="I. TRẮC NGHIỆM (7,0 ĐIỂM)",
        subtitle="PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    for item in DATA_EXAM["part1"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)
        render_options_no_table(doc, item["options"], is_teacher=is_teacher)
        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN II. Câu trắc nghiệm đúng sai.
    add_section_header(
        doc,
        title="PHẦN II. Câu trắc nghiệm đúng sai.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    for item in DATA_EXAM["part2"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        for stmt, ans in item["statements"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt, ans=ans, is_teacher=is_teacher, base_size=10.5)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # PHẦN III. Câu trắc nghiệm trả lời ngắn.
    add_section_header(
        doc,
        title="PHẦN III. Câu trắc nghiệm trả lời ngắn.",
        subtitle="Thí sinh trả lời từ câu 1 đến câu 4. Thí sinh điền kết quả vào ô trả lời tương ứng."
    )

    for item in DATA_EXAM["part3"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(2)
        if is_teacher:
            r_ans_lbl = p_ans.add_run("Trả lời: ")
            set_run_style(r_ans_lbl, font_size=10.5, bold=True)
            r_ans_val = p_ans.add_run(item["answer"])
            set_run_style(r_ans_val, font_size=10.5, bold=True, color=COLOR_CORRECT)
        else:
            r_ans = p_ans.add_run("Trả lời: ............................................................ ")
            set_run_style(r_ans, font_size=10.5, italic=True)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])

    # II. TỰ LUẬN (3,0 ĐIỂM)
    add_section_header(
        doc,
        title="II. TỰ LUẬN (3,0 ĐIỂM)",
        subtitle="Thí sinh trình bày chi tiết lời giải vào giấy làm bài thi."
    )

    for item in DATA_EXAM["part4"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, item["prompt"], base_size=10.5)

        if is_teacher and "solution" in item:
            add_solution_no_table(doc, item["solution"])
        elif not is_teacher:
            num_lines = item.get("num_lines", 6)
            add_student_writing_lines(doc, num_lines=num_lines)

    # Phần kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(0)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(r_end, font_size=11, bold=True)

    suffix = "_GiaoVien.docx" if is_teacher else "_HocSinh.docx"
    filename = f"De_Kiem_Tra_CK1_Hoa_10_CaoThang_Ma1010{suffix}"
    out_path = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(out_path)
    print(f"Đã tạo file Word: {out_path}")
    return out_path

if __name__ == "__main__":
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
