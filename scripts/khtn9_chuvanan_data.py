# -*- coding: utf-8 -*-
"""
Dữ liệu Đề kiểm tra Cuối Học kỳ I - Môn KHTN 9
Trường THCS Chu Văn An - Năm học 2024 - 2025
Thời gian làm bài: 60 phút (ĐỀ CHÍNH THỨC)
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

EXAM_INFO = {
    "school": "TRƯỜNG THCS CHU VĂN AN",
    "exam_title": "ĐỀ KIỂM TRA CUỐI HỌC KỲ I",
    "year": "NĂM HỌC 2024 – 2025",
    "subject": "MÔN: KHOA HỌC TỰ NHIÊN LỚP 9",
    "time_str": "Thời gian làm bài: 60 phút (Không kể thời gian giao đề)",
    "official_label": "ĐỀ CHÍNH THỨC",
    "atomic_mass": "Biết: H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; Al = 27; S = 32; Cl = 35,5; Ca = 40; Fe = 56; Cu = 64; Zn = 65; Br = 80; Ag = 108"
}

# -----------------------------------------------------------------------------
# I. PHẦN TRẮC NGHIỆM (4,0 điểm) Hãy chọn một phương án trả lời đúng nhất cho
# mỗi câu sau và ghi vào giấy bài làm, mỗi câu trả lời đúng được 0,25 điểm
# -----------------------------------------------------------------------------
MCQ_QUESTIONS = [
    {
        "id": 1,
        "content": "Bốn loại đơn phân cấu tạo nên DNA có kí hiệu là:",
        "options": [
            "A. A, U, G, C",
            "B. A, T, G, C",
            "C. A, D, R, T",
            "D. U, R, D, C"
        ],
        "correct": "B",
        "solution": "DNA (deoxyribonucleic acid) được cấu tạo theo nguyên tắc đa phân từ 4 loại đơn phân nucleotide có tên gọi viết tắt là: A (adenine), T (thymine), G (guanine) và C (cytosine). (Uracil - U là đơn phân của RNA)."
    },
    {
        "id": 2,
        "content": "Trường hợp gene cấu trúc bị đột biến thay thế một cặp A - T bằng một cặp G - C thì số liên kết hydrogen trong gene sẽ là:",
        "options": [
            "A. giảm 1 liên kết",
            "B. giảm 2 liên kết",
            "C. tăng 1 liên kết",
            "D. tăng 2 liên kết"
        ],
        "correct": "C",
        "solution": "Cặp nucleotide A - T liên kết với nhau bằng 2 liên kết hydrogen, cặp G - C liên kết với nhau bằng 3 liên kết hydrogen. Khi thay thế một cặp A - T bằng một cặp G - C thì số liên kết hydrogen của gene tăng lên: 3 - 2 = 1 liên kết."
    },
    {
        "id": 3,
        "content": "Kết thúc quá trình tái bản, một phân tử DNA ban đầu sẽ tạo ra:",
        "options": [
            "A. hai phân tử DNA mới có cấu tạo khác một phần so với DNA mẹ ban đầu",
            "B. hai phân tử DNA mới có cấu tạo giống một phần so với DNA mẹ ban đầu",
            "C. hai phân tử DNA mới có cấu tạo khác hoàn toàn so với DNA mẹ ban đầu",
            "D. hai phân tử DNA mới có cấu tạo giống hoàn toàn so với DNA mẹ ban đầu"
        ],
        "correct": "D",
        "solution": "Theo nguyên tắc bổ sung và nguyên tắc bán bảo toàn (giữ lại một nửa mạch gốc), kết thúc quá trình tái bản (nhân đôi), một phân tử DNA mẹ ban đầu tạo ra 2 phân tử DNA con có cấu tạo giống hệt nhau và giống hoàn toàn với phân tử DNA mẹ ban đầu."
    },
    {
        "id": 4,
        "content": "Loại enzyme nào sau đây tham gia vào quá trình phiên mã?",
        "options": [
            "A. Enzyme cắt giới hạn.",
            "B. Enzyme RNA polymerase.",
            "C. Enzyme DNA polymerase.",
            "D. Enzyme peptidase."
        ],
        "correct": "B",
        "solution": "Quá trình phiên mã tổng hợp phân tử RNA từ mạch khuôn của gene (DNA) được xúc tác bởi enzyme RNA polymerase. (Enzyme cắt giới hạn dùng trong kĩ thuật di truyền; DNA polymerase tham gia tái bản DNA; peptidase thủy phân liên kết peptide)."
    },
    {
        "id": 5,
        "content": "Nếu trình tự nucleotide trên mạch gốc của DNA là: 5'- ATGCGGATTTAA - 3' thì trình tự trên mạch bổ sung sẽ như thế nào?",
        "options": [
            "A. 3'-TACGCCTAAATT-5'",
            "B. 3'-TTAAATCCGCAT-5'",
            "C. 5'-TTAAATCCGCAT-3'",
            "D. 5'-AUGUGGAUUUAA-3'"
        ],
        "correct": "A",
        "solution": "Theo nguyên tắc bổ sung (A liên kết với T, T liên kết với A, G liên kết với C, C liên kết với G) và chiều hai mạch đối song song (5' -> 3' bổ sung với 3' -> 5'):\nMạch gốc: 5'- A - T - G - C - G - G - A - T - T - T - A - A - 3'\nMạch bổ sung: 3'- T - A - C - G - C - C - T - A - A - A - T - T - 5'."
    },
    {
        "id": 6,
        "content": "Trình tự các nucleotide trên gene quy định trình tự các nucleotide trên phân tử mRNA thông qua quá trình:",
        "options": [
            "A. dịch mã",
            "B. tái bản",
            "C. phiên mã",
            "D. biểu hiện"
        ],
        "correct": "C",
        "solution": "Phiên mã là quá trình truyền thông tin di truyền từ trình tự các nucleotide trên mạch khuôn của gene (DNA) sang trình tự các ribonucleotide trên phân tử mRNA theo nguyên tắc bổ sung."
    },
    {
        "id": 7,
        "content": "Sơ đồ mối quan hệ giữa gene và protein trong việc biểu hiện các tính trạng ở sinh vật là:",
        "options": [
            "A. gene (DNA) -> mRNA -> protein -> tính trạng",
            "B. gene (DNA) -> protein -> mRNA -> tính trạng",
            "C. mRNA -> gene (DNA) -> protein -> tính trạng",
            "D. protein -> gene (DNA) -> mRNA -> tính trạng"
        ],
        "correct": "A",
        "solution": "Cơ chế di truyền ở cấp độ phân tử biểu hiện tính trạng: Trình tự nucleotide trong gene (DNA) qua phiên mã quy định trình tự ribonucleotide trong mRNA; trình tự các ribonucleotide trong mRNA qua dịch mã quy định trình tự amino acid trong protein; protein trực tiếp tham gia cấu trúc và chức năng biểu hiện thành tính trạng của cơ thể."
    },
    {
        "id": 8,
        "content": "Phát biểu nào sau đây là đúng khi nói về cặp nhiễm sắc thể tương đồng?",
        "options": [
            "A. Mỗi cặp nhiễm sắc thể tương đồng gồm hai nhiễm sắc thể giống nhau về hình thái và kích thước, có nguồn gốc từ bố hoặc từ mẹ.",
            "B. Mỗi cặp nhiễm sắc thể tương đồng gồm hai nhiễm sắc thể khác nhau về hình thái và kích thước, trong đó, một nhiễm sắc thể có nguồn gốc từ bố, một nhiễm sắc thể có nguồn gốc từ mẹ.",
            "C. Mỗi cặp nhiễm sắc thể tương đồng gồm hai nhiễm sắc thể khác nhau về hình thái và kích thước, có nguồn gốc từ bố hoặc từ mẹ.",
            "D. Mỗi cặp nhiễm sắc thể tương đồng gồm hai nhiễm sắc thể giống nhau về hình thái và kích thước, trong đó, một nhiễm sắc thể có nguồn gốc từ bố, một nhiễm sắc thể có nguồn gốc từ mẹ."
        ],
        "correct": "D",
        "solution": "Cặp nhiễm sắc thể tương đồng gồm hai nhiễm sắc thể giống nhau về hình thái và kích thước, trong đó một chiếc có nguồn gốc từ bố và một chiếc có nguồn gốc từ mẹ."
    },
    {
        "id": 9,
        "content": "Phân tử nào sau đây biểu hiện thành tính trạng của cơ thể?",
        "options": [
            "A. mRNA",
            "B. DNA",
            "C. Protein",
            "D. Ribosome"
        ],
        "correct": "C",
        "solution": "Protein là phân tử trực tiếp tham gia vào cấu trúc tế bào và thực hiện các chức năng sinh lí (xúc tác, vận chuyển, miễn dịch, điều hòa...), từ đó biểu hiện thành tính trạng của cơ thể."
    },
    {
        "id": 10,
        "content": "Cấu trúc nhiễm sắc thể gồm:",
        "options": [
            "A. protein histone amino acid",
            "B. phân tử DNA và protein histone",
            "C. phân tử DNA và nucleotide",
            "D. amino acid và protein histone"
        ],
        "correct": "B",
        "solution": "Ở sinh vật nhân thực, cấu trúc của nhiễm sắc thể được tạo thành từ phân tử DNA quấn quanh các khối protein loại histone tạo nên chuỗi nucleosome và các mức xoắn tiếp theo."
    },
    {
        "id": 11,
        "content": "Các chất sau: C~6~H~12~O~6~, H~2~CO~3~, CH~4~, CCl~4~ đều là:",
        "options": [
            "A. hợp chất vô cơ",
            "B. hợp chất hữu cơ",
            "C. hydrocarbon",
            "D. hợp chất chứa carbon"
        ],
        "correct": "D",
        "solution": "Phân tích các chất:\n- C~6~H~12~O~6~, CH~4~, CCl~4~ là các hợp chất hữu cơ.\n- H~2~CO~3~ là hợp chất vô cơ chứa carbon.\nDo đó, điểm chung chính xác nhất của tất cả 4 chất trên là chúng đều là những *hợp chất chứa carbon*."
    },
    {
        "id": 12,
        "content": "Ngâm một viên kẽm sạch trong dung dịch CuSO~4~. Hiện tượng quan sát được là:",
        "options": [
            "A. không có hiện tượng nào xảy ra",
            "B. có chất rắn màu đỏ bám ngoài viên kẽm và màu xanh lam của dung dịch nhạt dần",
            "C. không có chất mới nào sinh ra, chỉ có một phần viên kẽm bị hòa tan",
            "D. kim loại đồng màu đỏ bám ngoài viên kẽm, viên kẽm không bị hòa tan"
        ],
        "correct": "B",
        "solution": "Phương trình hóa học: Zn + CuSO~4~ -> ZnSO~4~ + Cu v\nKẽm hoạt động hóa học mạnh hơn đồng nên đẩy đồng ra khỏi dung dịch muối: kim loại đồng (Cu) màu đỏ sinh ra bám ngoài bề mặt viên kẽm, kẽm tan dần vào dung dịch, đồng thời nồng độ ion Cu^2+^ giảm làm màu xanh lam của dung dịch CuSO~4~ nhạt dần."
    },
    {
        "id": 13,
        "content": "Dãy kim loại đều phản ứng được với nước ở nhiệt độ thường là :",
        "options": [
            "A. Na, Fe, K",
            "B. Na, Li, Mg",
            "C. Na, K, Ca",
            "D. Na, Li, Fe"
        ],
        "correct": "C",
        "solution": "Các kim loại kiềm và kiềm thổ như K, Na, Ca, Ba, Li phản ứng mãnh liệt với nước ở nhiệt độ thường tạo thành dung dịch kiềm và giải phóng khí hydrogen (H~2~). Fe và Mg không phản ứng với nước ở nhiệt độ thường."
    },
    {
        "id": 14,
        "content": "Các chất đều phản ứng được với dung dịch H~2~SO~4~ loãng là :",
        "options": [
            "A. Al, Fe, Cu",
            "B. Al, Zn, Fe",
            "C. Zn, Cu, Ag",
            "D. Zn, Al, Cu"
        ],
        "correct": "B",
        "solution": "Kim loại tác dụng được với dung dịch acid loãng (HCl, H~2~SO~4~ loãng) giải phóng H~2~ phải đứng trước hydrogen trong dãy hoạt động hóa học. Cả 3 kim loại Al, Zn, Fe đều đứng trước H nên đều phản ứng được. (Cu và Ag đứng sau H nên không phản ứng)."
    },
    {
        "id": 15,
        "content": "Dựa vào trạng thái, người ta chia nhiên liệu thành bao nhiêu loại?",
        "options": [
            "A. 2",
            "B. 3",
            "C. 4",
            "D. 5"
        ],
        "correct": "B",
        "solution": "Dựa vào trạng thái tồn tại (thể), nhiên liệu được chia thành 3 loại chính: nhiên liệu rắn (than đá, than củi, gỗ...), nhiên liệu lỏng (xăng, dầu hỏa, cồn...), và nhiên liệu khí (khí thiên nhiên, khí mỏ dầu, biogas...)."
    },
    {
        "id": 16,
        "content": "Dãy kim loại nào được sắp xếp theo chiều hoạt động hóa học giảm dần?",
        "options": [
            "A. Na, Mg, Zn",
            "B. Al, Zn, Na",
            "C. Mg, Al, Na",
            "D. Pb, Al, Mg"
        ],
        "correct": "A",
        "solution": "Theo trật tự dãy hoạt động hóa học của kim loại: K > Na > Ca > Mg > Al > Zn > Fe > Ni > Sn > Pb > (H) > Cu > Hg > Ag > Pt > Au.\nChiều hoạt động hóa học giảm dần là: Na > Mg > Zn."
    }
]

# -----------------------------------------------------------------------------
# II. PHẦN TỰ LUẬN (6,0 điểm)
# -----------------------------------------------------------------------------
ESSAY_QUESTIONS = [
    {
        "id": 17,
        "score_str": "(1,5 điểm)",
        "sub_items": [
            {
                "id": "17.1",
                "content": "Hợp chất hữu cơ là gì? Nhiên liệu là gì?",
                "solution": "• *Hợp chất hữu cơ:* Là hợp chất của carbon (trừ một số hợp chất vô cơ của carbon như carbon monoxide CO, carbon dioxide CO~2~, carbonic acid H~2~CO~3~, các muối carbonate, muối cyanide, carbide...).\n• *Nhiên liệu (chất đốt):* Là những chất cháy được, khi cháy tỏa nhiều nhiệt và phát sáng, được con người sử dụng làm nguồn cung cấp năng lượng cho sinh hoạt và sản xuất công nghiệp."
            },
            {
                "id": "17.2",
                "content": "Trình bày phương pháp hóa học để phân biệt các khí sau: CH~4~, H~2~, C~2~H~4~ khi chúng đựng trong các bình riêng biệt bị mất nhãn, viết phương trình hóa học xảy ra (nếu có).",
                "solution": "• *Bước 1: Nhận biết khí ethylene (C~2~H~4~):*\n- Dẫn lần lượt từng mẫu khí qua bình chứa dung dịch bromine (Br~2~) màu da cam.\n- Khí làm dung dịch bromine nhạt màu rồi mất màu hoàn toàn là ethylene (C~2~H~4~).\n  Phương trình hóa học:\n  CH~2~=CH~2~ + Br~2~ -> CH~2~Br-CH~2~Br  (hoặc C~2~H~4~ + Br~2~ -> C~2~H~4~Br~2~)\n- Hai khí không làm mất màu dung dịch bromine là methane (CH~4~) và hydrogen (H~2~).\n\n• *Bước 2: Phân biệt methane (CH~4~) và hydrogen (H~2~):*\n- Đốt cháy hoàn toàn 2 mẫu khí còn lại trong bình chứa khí oxygen, sau đó dẫn toàn bộ sản phẩm cháy lần lượt qua cốc đựng dung dịch nước vôi trong Ca(OH)~2~ dư:\n  + Sản phẩm cháy làm nước vôi trong vẩn đục (xuất hiện kết tủa trắng CaCO~3~) -> khí ban đầu là methane (CH~4~).\n    PTHH đốt cháy: CH~4~ + 2O~2~ --(t°)--> CO~2~ + 2H~2~O\n    PTHH tạo kết tủa: CO~2~ + Ca(OH)~2~ -> CaCO~3~ v + H~2~O\n  + Sản phẩm cháy chỉ ngưng tụ thành các giọt nước, không làm đục nước vôi trong -> khí ban đầu là hydrogen (H~2~).\n    PTHH: 2H~2~ + O~2~ --(t°)--> 2H~2~O"
            }
        ]
    },
    {
        "id": 18,
        "score_str": "(2,0 điểm)",
        "intro": "Đốt cháy hoàn toàn 14,4 gam hỗn hợp X gồm C~2~H~6~ và C~4~H~8~ phải dùng hết 39,664 lít khí oxygen (đo ở điều kiện chuẩn).",
        "sub_items": [
            {
                "id": "18.1",
                "content": "Tính phần trăm về khối lượng mỗi chất có trong hỗn hợp X?",
                "solution": "• Số mol khí oxygen phản ứng (ở điều kiện chuẩn, 1 mol = 24,79 L):\n  n~O2~ = 39,664 / 24,79 = 1,6 (mol).\n• Gọi số mol của C~2~H~6~ và C~4~H~8~ trong 14,4 gam hỗn hợp X lần lượt là x và y (x, y > 0).\n  Khối lượng mol: M(C~2~H~6~) = 30 g/mol; M(C~4~H~8~) = 56 g/mol.\n• Ta có phương trình khối lượng hỗn hợp X:\n  30x + 56y = 14,4    (1)\n• Các phương trình hóa học đốt cháy:\n  C~2~H~6~ + 3,5O~2~ --(t°)--> 2CO~2~ + 3H~2~O\n  x          3,5x           2x       3x          (mol)\n  C~4~H~8~ + 6O~2~ --(t°)--> 4CO~2~ + 4H~2~O\n  y          6y             4y       4y          (mol)\n• Tổng số mol O~2~ tham gia phản ứng:\n  3,5x + 6y = 1,6    (2)\n• Giải hệ phương trình (1) và (2):\n  { 30x + 56y = 14,4\n  { 3,5x + 6y = 1,6\n  => x = 0,2 (mol) và y = 0,15 (mol).\n• Khối lượng từng chất trong hỗn hợp X:\n  m(C~2~H~6~) = 0,2 * 30 = 6,0 (g)\n  m(C~4~H~8~) = 0,15 * 56 = 8,4 (g)\n• Phần trăm về khối lượng mỗi chất:\n  %m(C~2~H~6~) = (6,0 / 14,4) * 100% = 41,67%\n  %m(C~4~H~8~) = 100% - 41,67% = 58,33%"
            },
            {
                "id": "18.2",
                "content": "Dẫn toàn bộ sản phẩm qua bình đựng nước vôi trong dư, thì khối lượng bình nước vôi trong tăng lên thêm a gam và thu được b gam kết tủa trắng. Tính a, b?\n(Biết: H = 1; O = 16; C = 12; Ca = 40; N = 14)",
                "solution": "• Tổng số mol CO~2~ sinh ra:\n  n~CO2~ = 2x + 4y = 2*0,2 + 4*0,15 = 0,4 + 0,6 = 1,0 (mol).\n• Tổng số mol H~2~O sinh ra:\n  n~H2O~ = 3x + 4y = 3*0,2 + 4*0,15 = 0,6 + 0,6 = 1,2 (mol).\n• Dẫn toàn bộ sản phẩm cháy qua bình đựng dung dịch nước vôi trong Ca(OH)~2~ dư:\n  - Cả CO~2~ và H~2~O đều bị giữ lại trong bình.\n  - Khối lượng bình tăng thêm a gam chính bằng tổng khối lượng của CO~2~ và H~2~O:\n    a = m~CO2~ + m~H2O~ = 1,0 * 44 + 1,2 * 18 = 44 + 21,6 = *65,6 gam*.\n  - Khí CO~2~ phản ứng hoàn toàn với dung dịch Ca(OH)~2~ dư tạo kết tủa trắng CaCO~3~:\n    CO~2~ + Ca(OH)~2~ -> CaCO~3~ v + H~2~O\n    Theo phương trình: n~CaCO3~ = n~CO2~ = 1,0 (mol).\n  - Khối lượng kết tủa trắng thu được là:\n    b = m~CaCO3~ = 1,0 * 100 = *100 gam*.\n• Kết luận: a = 65,6 gam; b = 100 gam."
            }
        ]
    },
    {
        "id": 19,
        "score_str": "(0,75 điểm)",
        "content": "Thế năng là gì? Viết biểu thức tính thế năng của vật.",
        "solution": "• *Khái niệm thế năng:* Thế năng là dạng năng lượng mà một vật có được do vị trí tương đối của nó so với các vật khác (hoặc so với mốc chọn thế năng), hoặc do vật bị biến dạng đàn hồi.\n  Trong phạm vi cơ học, thế năng trọng trường là dạng năng lượng tương tác giữa Trái Đất và vật, phụ thuộc vào độ cao của vật so với mốc chọn tính thế năng.\n• *Biểu thức tính thế năng trọng trường:*\n  W~t~ = P·h = m·g·h\n  Trong đó:\n  - W~t~: Thế năng trọng trường của vật (đơn vị: Jun, kí hiệu J).\n  - P: Trọng lượng của vật (đơn vị: Niutơn, kí hiệu N).\n  - m: Khối lượng của vật (đơn vị: kilôgam, kí hiệu kg).\n  - g: Gia tốc trọng trường (m/s^2^).\n  - h: Độ cao của vật so với mốc chọn tính thế năng (đơn vị: mét, kí hiệu m)."
    },
    {
        "id": 20,
        "score_str": "(1,75 điểm)",
        "sub_items": [
            {
                "id": "20.1",
                "content": "Hiện tượng khúc xạ ánh sáng là gì?",
                "solution": "• *Khái niệm hiện tượng khúc xạ ánh sáng:*\nHiện tượng khúc xạ ánh sáng là hiện tượng tia sáng bị gãy khúc (đổi phương truyền) tại mặt phân cách giữa hai môi trường trong suốt khác nhau khi truyền xiên góc từ môi trường trong suốt này sang môi trường trong suốt khác."
            },
            {
                "id": "20.2",
                "content": "Tia sáng truyền từ không khí vào chất lỏng với góc tới i = 60° thì góc khúc xạ r = 30°. Khi chiếu tia sáng đó từ chất lỏng ra không khí, muốn có phản xạ toàn phần thì góc tới i phải thoả mãn điều kiện nào?",
                "solution": "• *Bước 1: Tính chiết suất của chất lỏng:*\n  - Áp dụng định luật khúc xạ ánh sáng khi tia sáng truyền từ không khí (chiết suất n~1~ = 1) vào chất lỏng (chiết suất n~2~ = n):\n    n~1~·sini = n~2~·sinr\n    <=> 1·sin(60°) = n·sin(30°)\n    => n = sin(60°) / sin(30°) = (\\sqrt{3} / 2) / (1 / 2) = \\sqrt{3} \\approx 1,732.\n\n• *Bước 2: Tìm điều kiện để xảy ra phản xạ toàn phần:*\n  - Khi chiếu tia sáng từ chất lỏng (chiết suất n~tới~ = \\sqrt{3}) ra không khí (chiết suất n~kx~ = 1):\n    + Điều kiện 1: Ánh sáng truyền từ môi trường chiết quang hơn sang môi trường chiết quang kém hơn (n~tới~ > n~kx~, thỏa mãn vì \\sqrt{3} > 1).\n    + Điều kiện 2: Góc tới i phải lớn hơn hoặc bằng góc giới hạn phản xạ toàn phần (i~gh~):\n      i >= i~gh~\n  - Tính góc giới hạn phản xạ toàn phần i~gh~:\n    sin(i~gh~) = n~kx~ / n~tới~ = 1 / \\sqrt{3} \\approx 0,5774\n    => i~gh~ = arcsin(1 / \\sqrt{3}) \\approx 35,26°  (hay 35°16').\n\n• *Kết luận:* Muốn xảy ra hiện tượng phản xạ toàn phần thì góc tới i trong chất lỏng phải thỏa mãn:\n  *i >= 35,3°* (hoặc *sin(i) >= 1 / \\sqrt{3}*)."
            }
        ]
    }
]
