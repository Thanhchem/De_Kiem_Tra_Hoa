# -*- coding: utf-8 -*-
"""
Script tạo Đề Kiểm Tra Giữa Học Kì II - Hóa học 11 - Trường THPT Gia Hội (TP Huế)
Mã đề thi: 2608 - Năm học 2025 - 2026
Chuẩn theo Quy chuẩn Form Đề Kiểm Tra của Thầy TRẦN VĂN THẠNH
Tạo 2 phiên bản:
1. De_Kiem_Tra_GK2_Hoa_11_GiaHoi_Ma2608_HocSinh.docx
2. De_Kiem_Tra_GK2_Hoa_11_GiaHoi_Ma2608_GiaoVien.docx
"""

import os
import sys
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

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
DIR_IMAGES = os.path.join(DIR_SAN_PHAM, "images_gk2_hoa11_giahoi_2608")
os.makedirs(DIR_SAN_PHAM, exist_ok=True)
os.makedirs(DIR_IMAGES, exist_ok=True)

DATA_EXAM = {
    "school_line": "SỞ GD & ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT GIA HỘI",
    "exam_title": "KIỂM TRA GIỮA KÌ II - NĂM HỌC 2025 - 2026",
    "subject_line": "MÔN: HÓA HỌC - LỚP 11",
    "duration": "45 phút (không kể thời gian phát đề)",
    "num_pages": "03",
    "exam_code": "2608",
    "atomic_masses": "H: 1; C: 12; O: 16; Cl: 35,5; Ag: 108; Br: 80; N: 14.",

    # PHẦN I (12 câu - 3,0 điểm)
    "part1": [
        {
            "id": 1,
            "prompt": "Câu 1. Phản ứng thế giữa 2,3-dimethylbutane với Cl~2~ (tỉ lệ mol 1:1) cho mấy sản phẩm thế?",
            "options": [
                "*A. 2.*",
                "B. 5.",
                "C. 3.",
                "D. 4."
            ],
            "solution": "Công thức cấu tạo của 2,3-dimethylbutane:\n  CH~3~–CH(CH~3~)–CH(CH~3~)–CH~3~\nPhân tử có tính đối xứng cao:\n- 4 nhóm –CH~3~ ở hai đầu tương đương nhau (12 nguyên tử H) $\\rightarrow$ thế 1 nguyên tử Cl vào nhóm –CH~3~ cho 1 sản phẩm: 1-chloro-2,3-dimethylbutane.\n- 2 nhóm –CH– tương đương nhau (2 nguyên tử H) $\\rightarrow$ thế 1 nguyên tử Cl vào nhóm –CH– cho 1 sản phẩm: 2-chloro-2,3-dimethylbutane.\nVậy số sản phẩm thế monoclo (cấu tạo) thu được là 2.\n*Đáp án đúng: A.*"
        },
        {
            "id": 2,
            "prompt": "Câu 2. Cho 5 chất: CH~3~CH~2~CH~2~Cl (1); CH~2~=CHCH~2~Cl (2); C~6~H~5~Cl (3); CH~2~=CHCl (4); C~6~H~5~CH~2~Cl (5). Đun từng chất với dung dịch NaOH loãng, dư, sau đó gạn lấy lớp nước và acid hoá bằng dung dịch HNO~3~, sau đó nhỏ vào đó dung dịch AgNO~3~ thì các chất có xuất hiện kết tủa trắng là",
            "options": [
                "A. (1), (2), (3), (5).",
                "*B. (1), (2), (5).*",
                "C. (2), (3), (5).",
                "D. (1), (3), (5)."
            ],
            "solution": "Phản ứng thủy phân dẫn xuất halogen bởi dung dịch kiềm loãng đun nóng:\n- Các dẫn xuất alkyl (1: CH~3~CH~2~CH~2~Cl), allyl (2: CH~2~=CHCH~2~Cl) và benzyl (5: C~6~H~5~CH~2~Cl) có liên kết C–Cl linh động, dễ bị thủy phân giải phóng ion Cl^–^: R–Cl + NaOH $\\xrightarrow{t^o}$ R–OH + NaCl. Sau khi acid hóa bằng HNO~3~ rồi thêm AgNO~3~ sẽ xuất hiện kết tủa trắng AgCl (Ag^+^ + Cl^–^ $\\rightarrow$ AgCl↓).\n- Dẫn xuất phenyl (3: C~6~H~5~Cl) và vinyl (4: CH~2~=CHCl) có liên kết C–Cl rất bền do hiệu ứng liên hợp p–$\\pi$, không bị thủy phân bởi dung dịch NaOH loãng ở điều kiện thường.\nVậy các chất tạo kết tủa trắng là (1), (2), (5).\n*Đáp án đúng: B.*"
        },
        {
            "id": 3,
            "prompt": "Câu 3. Cho dãy các chất: (a) toluene; (b) o-xylene; (c) ethylbenzene; (d) m-dimethylbenzene; (e) styrene. Những chất đồng đẳng của benzene là",
            "options": [
                "A. (a), (b), (c), (e).",
                "B. (a), (d).",
                "*C. (a), (b), (c), (d).*",
                "D. (a), (e)."
            ],
            "solution": "Đồng đẳng của benzene là các alkylbenzene có công thức chung C~n~H~2n-6~ (n ≥ 6):\n- (a) toluene (C~7~H~8~): C~6~H~5~–CH~3~ (thuộc dãy đồng đẳng).\n- (b) o-xylene (C~8~H~10~): 1,2-dimethylbenzene (thuộc dãy đồng đẳng).\n- (c) ethylbenzene (C~8~H~10~): C~6~H~5~–CH~2~CH~3~ (thuộc dãy đồng đẳng).\n- (d) m-dimethylbenzene (C~8~H~10~): 1,3-dimethylbenzene (thuộc dãy đồng đẳng).\n- (e) styrene (C~8~H~8~): C~6~H~5~–CH=CH~2~ có liên kết đôi ở mạch nhánh, không no nên không thuộc dãy đồng đẳng của benzene.\n*Đáp án đúng: C.*"
        },
        {
            "id": 4,
            "prompt": "Câu 4. Ethyl chloride được dùng làm thuốc xịt có tác dụng giảm đau tạm thời khi chơi thể thao. Công thức phân tử của ethyl chloride là",
            "options": [
                "A. C~2~H~5~F.",
                "B. C~2~H~3~Cl.",
                "C. CH~3~Cl.",
                "*D. C~2~H~5~Cl.*"
            ],
            "solution": "Gốc ethyl là C~2~H~5~–, chloride là –Cl $\\rightarrow$ Công thức phân tử của ethyl chloride là C~2~H~5~Cl.\n*Đáp án đúng: D.*"
        },
        {
            "id": 5,
            "prompt": "Câu 5. Dẫn xuất halogen của hydrocarbon không chứa nguyên tố nào sau đây?",
            "options": [
                "A. Chlorine.",
                "B. Bromine.",
                "*C. Oxygen.*",
                "D. Fluorine."
            ],
            "solution": "Dẫn xuất halogen của hydrocarbon được tạo thành khi thay thế một hay nhiều nguyên tử hydrogen trong phân tử hydrocarbon bằng một hay nhiều nguyên tử halogen (F, Cl, Br, I). Phân tử chỉ chứa các nguyên tố C, H và Halogen; không chứa Oxygen.\n*Đáp án đúng: C.*"
        },
        {
            "id": 6,
            "prompt": "Câu 6. Cho X có công thức cấu tạo: CH~3~–CH=C(CH~3~)–CH~3~. Danh pháp thay thế của X là",
            "options": [
                "A. 3-methylbut-3-ene.",
                "*B. 2-methylbut-2-ene.*",
                "C. 3-methylbut-2-ene.",
                "D. 2-methylbut-3-ene."
            ],
            "solution": "Quy tắc gọi tên theo danh pháp IUPAC:\n- Chọn mạch chính là mạch dài nhất chứa liên kết đôi: mạch 4 carbon (but).\n- Đánh số mạch chính từ đầu gần liên kết đôi hơn: liên kết đôi ở vị trí giữa mạch (C2=C3), ta ưu tiên đánh số từ đầu gần nhánh methyl hơn (từ phải sang trái):\n  $^{4}$CH~3~–$^{3}$CH=$^{2}$C(CH~3~)–$^{1}$CH~3~.\n- Tên gọi: 2-methylbut-2-ene.\n*Đáp án đúng: B.*"
        },
        {
            "id": 7,
            "prompt": "Câu 7. Đặc điểm nào sau đây là của hydrocarbon no?",
            "options": [
                "A. Chỉ có liên kết đôi.",
                "B. Có ít nhất một vòng no.",
                "C. Có ít nhất một liên kết đôi.",
                "*D. Chỉ có liên kết đơn.*"
            ],
            "solution": "Hydrocarbon no là hydrocarbon mà trong phân tử chỉ chứa các liên kết đơn (liên kết $\\sigma$).\n*Đáp án đúng: D.*"
        },
        {
            "id": 8,
            "prompt": "Câu 8. Cho các dẫn xuất halogen của hydrocarbon sau: C~2~H~5~F (1); C~2~H~5~Br (2); C~2~H~5~I (3); C~2~H~5~Cl (4). Thứ tự giảm dần nhiệt độ sôi là",
            "options": [
                "A. (1) > (4) > (2) > (3).",
                "*B. (3) > (2) > (4) > (1).*",
                "C. (1) > (2) > (3) > (4).",
                "D. (3) > (2) > (1) > (4)."
            ],
            "solution": "Với các dẫn xuất halogen có cùng gốc alkyl (C~2~H~5~–), khi khối lượng phân tử và bán kính nguyên tử halogen tăng dần (F < Cl < Br < I), kích thước đám mây electron tăng dẫn đến tương tác van der Waals giữa các phân tử tăng $\\rightarrow$ nhiệt độ sôi tăng dần: C~2~H~5~F < C~2~H~5~Cl < C~2~H~5~Br < C~2~H~5~I.\nDo đó, thứ tự GIẢM DẦN nhiệt độ sôi là: C~2~H~5~I (3) > C~2~H~5~Br (2) > C~2~H~5~Cl (4) > C~2~H~5~F (1).\n*Đáp án đúng: B.*"
        },
        {
            "id": 9,
            "prompt": "Câu 9. Alkyne là các hydrocarbon không no, mạch hở, có công thức chung là",
            "options": [
                "*A. C~n~H~2n-2~ (n ≥ 2).*",
                "B. C~n~H~2n~ (n ≥ 2).",
                "C. C~n~H~2n+2~ (n ≥ 1).",
                "D. C~n~H~2n~ (n ≥ 3)."
            ],
            "solution": "Alkyne là hydrocarbon không no, mạch hở, trong phân tử có chứa một liên kết ba C≡C. Công thức chung của alkyne là C~n~H~2n-2~ (n ≥ 2).\n*Đáp án đúng: A.*"
        },
        {
            "id": 10,
            "prompt": "Câu 10. Benzene tác dụng được với chất nào sau đây ở điều kiện thích hợp?",
            "options": [
                "A. Dung dịch Br~2~.",
                "B. Dung dịch NaOH.",
                "*C. Br~2~/FeBr~3~, t^o^.*",
                "D. NaCl khan."
            ],
            "solution": "Vòng benzene rất bền vững, không phản ứng với dung dịch bromine ở điều kiện thường. Trong điều kiện có xúc tác FeBr~3~ (hoặc bột Fe) và đun nóng, benzene tham gia phản ứng thế với bromine khan (nguyên chất):\n  C~6~H~6~ + Br~2~ $\\xrightarrow{FeBr_3, t^o}$ C~6~H~5~Br + HBr.\n*Đáp án đúng: C.*"
        },
        {
            "id": 11,
            "prompt": "Câu 11. Cho sơ đồ phản ứng sau: CH~3~–C≡CH + AgNO~3~/NH~3~ $\\rightarrow$ X + NH~4~NO~3~. X có công thức cấu tạo là",
            "options": [
                "*A. CH~3~–C≡C–Ag.*",
                "B. CH~3~–C–Ag≡C–Ag.",
                "C. CAg≡CAg.",
                "D. Ag–CH~2~–C≡C–Ag."
            ],
            "solution": "Propyne (CH~3~–C≡CH) là một alk-1-yne, nguyên tử hydrogen liên kết trực tiếp với nguyên tử carbon của liên kết ba mang tính linh động nên bị thay thế bởi ion Ag^+^ trong phức chất [Ag(NH~3~)~2~]OH $\\rightarrow$ tạo kết tủa màu vàng nhạt bạc propynide: CH~3~–C≡C–Ag.\n*Đáp án đúng: A.*"
        },
        {
            "id": 12,
            "prompt": "Câu 12. Ở điều kiện thường alkane nào sau đây ở thể lỏng?",
            "options": [
                "A. C~4~H~10~.",
                "B. C~3~H~8~.",
                "C. C~2~H~6~.",
                "*D. C~6~H~14~.*"
            ],
            "solution": "Ở điều kiện thường, các alkane từ C~1~ đến C~4~ (CH~4~, C~2~H~6~, C~3~H~8~, C~4~H~10~) tồn tại ở thể khí. Các alkane từ C~5~ đến khoảng C~17~ tồn tại ở thể lỏng (hexane C~6~H~14~ là chất lỏng). Từ C~18~ trở lên tồn tại ở thể rắn.\n*Đáp án đúng: D.*"
        }
    ],

    # PHẦN II (2 câu - 2,0 điểm)
    'part2': [
        {
            'id': 1,
            'prompt': 'Câu 1. Tiến hành thí nghiệm điều chế và thử tính chất của ethylene, người ta tiến hành các bước sau đây:\nBước 1: Cho vào bình cầu khoảng 10 mL cồn 90^o^, thêm tiếp 5 mL dung dịch H~2~SO~4~ đậm đặc và một ít đá bọt, lắc đều.\nBước 2: Chuẩn bị 2 bình thủy tinh, cho vào khoảng 5 mL dung dịch NaOH đặc vào bình (1) và khoảng 2 mL nước bromine vào bình (2).\nBước 3: Dùng đèn cồn đun nóng đều bình cầu; dẫn khí thu được lần lượt vào bình (1) sau đó vào bình (2).',
            'image': os.path.join(DIR_IMAGES, 'thi_nghiem_dieu_che_ethylene.png'),
            'image_width': 7.5,
            'statements': [
                ('a) Vai trò của đá bọt là làm chất lỏng sôi đều và không bị trào ra.', True),
                ('b) Có thể thay dung dịch H~2~SO~4~ đặc bằng dung dịch H~2~SO~4~ 10% vẫn thu được ethylene.', False),
                ('c) Khi ethylene sinh ra sẽ làm nhạt màu ống nghiệm đựng dung dịch Br~2~, phản ứng xảy ra là phản ứng oxi hóa khử.', True),
                ('d) Trong phòng thí nghiệm, để thu khí ethylene, người ta dùng phương pháp đẩy nước.', True)
            ],
            'solution': 'a) ĐÚNG. Đá bọt (chứa các mao quản xốp) đóng vai trò điều hòa quá trình sôi, phân tán nhiệt đều giúp chất lỏng sôi êm và tránh hiện tượng quá sôi cục bộ làm trào chất lỏng ra ngoài.\nb) SAI. H~2~SO~4~ đặc đóng vai trò vừa là chất xúc tác vừa là chất hút nước mạnh ở nhiệt độ cao (~170^o^C) để thực hiện phản ứng tách nước alcohol thành alkene. H~2~SO~4~ loãng 10% không có khả năng hút nước nên không thể tách nước tạo ethylene.\nc) ĐÚNG. Khí ethylene sinh ra phản ứng cộng với bromine làm nhạt màu/mất màu nước bromine:\n  CH~2~=CH~2~ + Br~2~ $\\rightarrow$ CH~2~Br–CH~2~Br.\nTrong phản ứng này, số oxi hóa của C tăng từ -2 lên -1 (chất khử), của Br giảm từ 0 xuống -1 (chất oxi hóa) $\\rightarrow$ đây là phản ứng oxi hóa – khử.\nd) ĐÚNG. Do khí ethylene rất ít tan trong nước và nhẹ hơn không khí nên phương pháp tốt nhất để thu khí sạch trong phòng thí nghiệm là phương pháp đẩy nước.'
        },
        {
            'id': 2,
            'prompt': 'Câu 2. Benzene, toluene và styrene đều thuộc hợp chất hydrocarbon thơm (Arene).',
            'statements': [
                ('a) Benzene, toluene và styrene đều có công thức chung là C~n~H~2n-6~ (n ≥ 6).', False),
                ('b) Styrene còn có tên gọi khác là ethylbenzene.', False),
                ('c) Để phân biệt benzene, toluene và styrene có thể dùng dung dịch KMnO~4~.', True),
                ('d) So với benzene thì toluene phản ứng thế nguyên tử hydrogen ở vòng benzene dễ dàng hơn so với benzene và sản phẩm chính ưu tiên thế vào vị trí meta (-m).', False)
            ],
            'solution': 'a) SAI. Benzene (C~6~H~6~) và toluene (C~7~H~8~) thuộc dãy đồng đẳng alkylbenzene có công thức chung C~n~H~2n-6~ (n ≥ 6). Styrene có công thức C~8~H~8~ (có 1 vòng benzene và 1 nối đôi ở nhánh, độ bất bão hòa k = 5) nên có công thức tổng quát là C~n~H~2n-8~.\nb) SAI. Styrene là vinylbenzene (C~6~H~5~–CH=CH~2~), còn ethylbenzene là hợp chất no ở mạch nhánh (C~6~H~5~–CH~2~CH~3~).\nc) ĐÚNG. Sử dụng dung dịch KMnO~4~:\n- Styrene chứa liên kết đôi ở mạch nhánh nên làm mất màu dung dịch KMnO~4~ ngay ở nhiệt độ thường.\n- Toluene chỉ làm mất màu dung dịch KMnO~4~ khi đun nóng.\n- Benzene không làm mất màu dung dịch KMnO~4~ kể cả khi đun nóng.\nd) SAI. Nhóm methyl (–CH~3~) là nhóm đẩy electron, làm tăng mật độ electron trong vòng thơm (đặc biệt tại các vị trí ortho và para), do đó phản ứng thế electrophile vào nhân thơm của toluene xảy ra dễ dàng hơn benzene và ưu tiên thế vào các vị trí ortho (-o) và para (-p), không phải meta (-m).'
        }
    ],

    # PHẦN III (4 câu - 2,0 điểm)
    'part3': [
        {
            'id': 1,
            'prompt': 'Câu 1. Trong số các hydrocarbon mạch hở sau: C~2~H~2~, C~4~H~10~, C~4~H~6~, C~4~H~8~, C~3~H~4~, C~3~H~6~. Có bao nhiêu hydrocarbon có thể tạo kết tủa với dung dịch AgNO~3~/NH~3~?',
            'ans': '3',
            'solution': 'Hydrocarbon mạch hở tác dụng với dung dịch AgNO~3~/NH~3~ tạo kết tủa vàng là các alk-1-yne (có liên kết ba đầu mạch –C≡CH):\n- C~2~H~2~: ethyne (CH≡CH) $\\rightarrow$ có phản ứng.\n- C~4~H~10~: alkane no $\\rightarrow$ không phản ứng.\n- C~4~H~6~: mạch hở có đồng phân but-1-yne (CH≡C–CH~2~CH~3~) $\\rightarrow$ có thể tạo kết tủa.\n- C~4~H~8~: alkene/cycloalkane $\\rightarrow$ không phản ứng.\n- C~3~H~4~: propyne (CH≡C–CH~3~) $\\rightarrow$ có phản ứng.\n- C~3~H~6~: propene $\\rightarrow$ không phản ứng.\nVậy có 3 hydrocarbon thỏa mãn: C~2~H~2~, C~4~H~6~, C~3~H~4~.\n*Đáp án: 3*'
        },
        {
            'id': 2,
            'prompt': 'Câu 2. Phần trăm khối lượng của carbon (%C) có trong phân tử ethylbenzene là bao nhiêu? (làm tròn đến hàng phần mười).',
            'ans': '90,6%',
            'solution': 'Ethylbenzene có công thức phân tử là C~8~H~10~ (C~6~H~5~–C~2~H~5~).\nKhối lượng mol: M = 8 × 12 + 10 × 1 = 96 + 10 = 106 g/mol.\nPhần trăm khối lượng của carbon trong phân tử:\n  %C = (96 / 106) × 100% ≈ 90,566% $\\rightarrow$ làm tròn đến hàng phần mười là 90,6%.\n*Đáp án: 90,6%*'
        },
        {
            'id': 3,
            'prompt': 'Câu 3. Trong các chất: CH~3~Cl; CH~3~NH~2~; CH~4~O; CH~2~Br~2~; C~2~H~4~Cl~2~, HCl, C~6~H~5~F có bao nhiêu chất thuộc loại dẫn xuất halogen của hydrocarbon?',
            'ans': '4',
            'solution': 'Dẫn xuất halogen của hydrocarbon được tạo thành khi thay thế một hay nhiều nguyên tử hydrogen trong phân tử hydrocarbon bằng một hay nhiều nguyên tử halogen (F, Cl, Br, I).\n- CH~3~Cl (chloromethane): dẫn xuất halogen.\n- CH~3~NH~2~ (methanamine): amine (chứa N), không phải dẫn xuất halogen.\n- CH~4~O (methanol): alcohol (chứa O), không phải dẫn xuất halogen.\n- CH~2~Br~2~ (dibromomethane): dẫn xuất halogen.\n- C~2~H~4~Cl~2~ (dichloroethane): dẫn xuất halogen.\n- HCl (hydrogen chloride): hợp chất vô cơ, không phải dẫn xuất của hydrocarbon.\n- C~6~H~5~F (fluorobenzene): dẫn xuất halogen.\nVậy có 4 chất thuộc loại dẫn xuất halogen của hydrocarbon: CH~3~Cl, CH~2~Br~2~, C~2~H~4~Cl~2~, C~6~H~5~F.\n*Đáp án: 4*'
        },
        {
            'id': 4,
            'prompt': 'Câu 4. Da nhân tạo (PVC) được điều chế từ khí thiên nhiên theo sơ đồ:\nCH~4~ $\\rightarrow$ C~2~H~2~ $\\rightarrow$ CH~2~=CH–Cl $\\rightarrow$ (–CH~2~–CHCl–)~n~.\nNếu hiệu suất của toàn bộ quá trình điều chế là 20%, muốn điều chế được 1,0 tấn PVC thì thể tích (m^3^) khí thiên nhiên (chứa 80% methane) ở điều kiện chuẩn cần dùng là bao nhiêu? (làm tròn kết quả đến hàng đơn vị).',
            'ans': '4958',
            'solution': 'Khối lượng PVC cần sản xuất: m~PVC~ = 1,0 tấn = 1000 kg = 10^6^ gam.\nKhối lượng mol của 1 mắt xích PVC (–CH~2~–CHCl–) là: M = 2 × 12 + 3 × 1 + 35,5 = 62,5 g/mol.\nSố mol mắt xích PVC cần tạo thành:\n  n~PVC~ = 10^6^ / 62,5 = 16 000 mol = 16 kmol.\nTheo sơ đồ bảo toàn nguyên tố carbon: 2 CH~4~ $\\rightarrow$ 1 mắt xích PVC.\nSố mol CH~4~ lí thuyết cần dùng:\n  n~CH4 (LT)~ = 2 × n~PVC~ = 2 × 16 = 32 kmol.\nVì hiệu suất của toàn bộ quá trình là 20% (H = 0,2) nên số mol CH~4~ thực tế cần dùng là:\n  n~CH4 (TT)~ = n~CH4 (LT)~ / 0,2 = 32 / 0,2 = 160 kmol = 160 000 mol.\nKhí thiên nhiên chứa 80% methane theo thể tích (tương ứng 80% theo số mol) nên số mol khí thiên nhiên cần dùng là:\n  n~khí TN~ = n~CH4 (TT)~ / 0,8 = 160 / 0,8 = 200 kmol = 200 000 mol.\nỞ điều kiện chuẩn (25^o^C, 1 bar), thể tích khí thiên nhiên cần dùng là:\n  V~khí TN~ = 200 × 24,79 = 4958 m^3^.\n(Nếu tính theo điều kiện tiêu chuẩn cũ 22,4 L/mol: V = 200 × 22,4 = 4480 m^3^).\n*Đáp án: 4958*'
        }
    ],

    # PHẦN IV (3 câu - 3,0 điểm)
    'part4': [
        {
            'id': 1,
            'prompt': 'Câu 1. (1,0 điểm) Khí hóa lỏng hay khí gas (LPG - Liquefied Petroleum Gas) được sử dụng làm nhiên liệu trong các hộ gia đình là propane hóa lỏng. Một bình khí gas hóa lỏng chứa 12,0 kg propane có thể cung cấp bao nhiêu Lít khí propane ở điều kiện chuẩn (25^o^C, 1 bar)?',
            'num_lines': 8,
            'solution': 'Khối lượng propane trong bình gas: m = 12,0 kg = 12 000 gam.\nKhối lượng mol của propane (C~3~H~8~): M = 3 × 12 + 8 × 1 = 44 g/mol.\nSố mol propane chứa trong bình:\n  n~C3H8~ = 12 000 / 44 ≈ 272,727 mol (hoặc 3000/11 mol).\nỞ điều kiện chuẩn (25^o^C, 1 bar), 1 mol khí chiếm thể tích là 24,79 Lít.\nThể tích khí propane ở điều kiện chuẩn mà bình gas có thể cung cấp là:\n  V = n × 24,79 = (3000 / 11) × 24,79 ≈ 6760,91 Lít.\n(Nếu tính theo điều kiện tiêu chuẩn cũ 22,4 L/mol: V = (12000 / 44) × 22,4 ≈ 6109,09 Lít).'
        },
        {
            'id': 2,
            'prompt': 'Câu 2. (1,0 điểm) Nhựa PE (Polyethylene) là loại nhựa nhiệt dẻo phổ biến nhất thế giới, nổi bật với đặc tính dẻo dai, chống thấm nước, cách điện tốt và độ bền cao. Một trong những cách điều chế PE là theo sơ đồ: C~2~H~2~ $\\rightarrow$ C~2~H~4~ $\\rightarrow$ polyethylene (PE).\nViết các phương trình phản ứng thực hiện sơ đồ trên (ghi rõ điều kiện phản ứng).',
            'num_lines': 8,
            'solution': '(1) Phản ứng hydrogen hóa có chọn lọc chuyển acetylene thành ethylene:\n  CH≡CH + H~2~ $\\xrightarrow{Pd/PbCO_3, t^o}$ CH~2~=CH~2~\\n\n(2) Phản ứng trùng hợp ethylene điều chế polyethylene (PE):\n  n CH~2~=CH~2~ $\\xrightarrow{t^o, p, xt}$ (–CH~2~–CH~2~–)~n~'
        },
        {
            'id': 3,
            'prompt': 'Câu 3. (1,0 điểm) Thuốc nổ TNT (2,4,6-trinitrotoluene) được điều chế bằng phản ứng của toluene với hỗn hợp gồm HNO~3~ đặc và H~2~SO~4~ đặc, trong điều kiện đun nóng. Biết hiệu suất của toàn bộ quá trình tổng hợp là 60%. Lượng TNT (2,4,6-trinitrotoluene) tạo thành từ 552,0 gam toluene là bao nhiêu gam?',
            'num_lines': 8,
            'solution': 'Phương trình hóa học điều chế 2,4,6-trinitrotoluene (TNT):\n  C~6~H~5~CH~3~ + 3 HNO~3~ $\\xrightarrow{H_2SO_4 đ, t^o}$ C~6~H~2~(NO~2~)~3~CH~3~ + 3 H~2~O\n  (Toluene: M = 92 g/mol)                    (TNT: M = 227 g/mol)\n\nSố mol toluene ban đầu:\n  n~toluene~ = 552,0 / 92 = 6,0 mol.\nTheo phương trình phản ứng, số mol TNT lí thuyết thu được:\n  n~TNT (LT)~ = n~toluene~ = 6,0 mol.\nDo hiệu suất của toàn bộ quá trình là 60% nên số mol TNT thực tế thu được là:\n  n~TNT (TT)~ = n~TNT (LT)~ × 60% = 6,0 × 0,6 = 3,6 mol.\nKhối lượng 2,4,6-trinitrotoluene (TNT) tạo thành:\n  m~TNT~ = n~TNT (TT)~ × M~TNT~ = 3,6 × 227 = 817,2 gam.'
        }
    ]
}

def generate_docx(is_teacher=False):
    doc = Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line=DATA_EXAM['school_line'],
        exam_title=DATA_EXAM['exam_title'],
        subject_line=DATA_EXAM['subject_line'],
        num_pages=DATA_EXAM['num_pages'],
        duration=DATA_EXAM['duration'],
        exam_code=DATA_EXAM['exam_code'],
        atomic_masses=DATA_EXAM['atomic_masses'],
        is_teacher=is_teacher
    )

    # I. TRẮC NGHIỆM NHIỀU LỰA CHỌN
    add_section_header(
        doc,
        title='I. TRẮC NGHIỆM NHIỀU LỰA CHỌN: (12 câu – 3,0 điểm)',
        subtitle='(Học sinh chọn đáp án đúng nhất. Mỗi câu trả lời đúng được 0,25 điểm)'
    )

    for item in DATA_EXAM['part1']:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item['prompt'], base_size=10.5)

        render_options_no_table(doc, item['options'], is_teacher=is_teacher)

        if is_teacher and 'solution' in item:
            add_solution_no_table(doc, item['solution'])

    # II. TRẮC NGHIỆM ĐÚNG-SAI
    add_section_header(
        doc,
        title='II. TRẮC NGHIỆM ĐÚNG-SAI: (2 câu – 2,0 điểm)',
        subtitle='(Trong mỗi ý (a), (b), (c), (d) ở mỗi câu, học sinh chọn đúng hoặc sai: Đ - S)'
    )

    for item in DATA_EXAM['part2']:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item['prompt'], base_size=10.5)

        if 'image' in item and os.path.exists(item['image']):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(2)
            r_img = p_img.add_run()
            r_img.add_picture(item['image'], width=Cm(item.get('image_width', 7.5)))

        for stmt, ans in item['statements']:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt, ans=ans, is_teacher=is_teacher, base_size=10.5)

        if is_teacher and 'solution' in item:
            add_solution_no_table(doc, item['solution'])

    # III. TRẮC NGHIỆM TRẢ LỜI NGẮN
    add_section_header(
        doc,
        title='III. TRẮC NGHIỆM TRẢ LỜI NGẮN: (4 câu – 2,0 điểm)',
        subtitle='(Thí sinh ghi đáp án ngắn gọn vào ô/dòng trả lời)'
    )

    for item in DATA_EXAM['part3']:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item['prompt'], base_size=10.5)

        if not is_teacher:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = Cm(0.5)
            p_ans.paragraph_format.space_before = Pt(1)
            p_ans.paragraph_format.space_after = Pt(2)
            r_ans_lbl = p_ans.add_run('Trả lời: ')
            set_run_style(r_ans_lbl, bold=True, font_size=10.5)
            r_dots = p_ans.add_run('.' * 40)
            set_run_style(r_dots, color=RGBColor(180, 180, 180), font_size=10.5)
        else:
            add_solution_no_table(doc, item['solution'])

    # IV. BÀI TẬP TỰ LUẬN
    add_section_header(
        doc,
        title='IV. BÀI TẬP TỰ LUẬN: (3 câu – 3,0 điểm)',
        subtitle='(Học sinh trình bày chi tiết lời giải vào bài làm)'
    )

    for item in DATA_EXAM['part4']:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, item['prompt'], base_size=10.5)

        if not is_teacher:
            add_student_writing_lines(doc, num_lines=item.get('num_lines', 6))
        else:
            add_solution_no_table(doc, item['solution'])

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(10)
    p_end.paragraph_format.space_after = Pt(6)
    r_end = p_end.add_run('--------- HẾT ---------')
    set_run_style(r_end, bold=True, font_size=11, color=RGBColor(100, 100, 100))

    suffix = '_GiaoVien' if is_teacher else '_HocSinh'
    filename = f'De_Kiem_Tra_GK2_Hoa_11_GiaHoi_Ma2608{suffix}.docx'
    filepath = os.path.join(DIR_SAN_PHAM, filename)
    doc.save(filepath)
    print(f'Saved: {filepath}')

if __name__ == '__main__':
    generate_docx(is_teacher=False)
    generate_docx(is_teacher=True)
