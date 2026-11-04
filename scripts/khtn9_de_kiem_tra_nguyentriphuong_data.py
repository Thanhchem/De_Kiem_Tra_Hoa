# -*- coding: utf-8 -*-
"""
Dữ liệu Đề kiểm tra Giữa học kì I - Môn KHTN 9 (MÃ ĐỀ 101)
Trường THCS Nguyễn Tri Phương - Sở GD&ĐT Thành phố Huế
Năm học 2025 - 2026 (Thời gian làm bài: 90 phút)
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

EXAM_INFO = {
    "so_gd": "SỞ GD&ĐT THÀNH PHỐ HUẾ",
    "school": "TRƯỜNG THCS NGUYỄN TRI PHƯƠNG",
    "exam_title": "ĐỀ KIỂM TRA GIỮA HỌC KÌ I",
    "year": "NĂM HỌC 2025 – 2026",
    "subject": "MÔN: KHOA HỌC TỰ NHIÊN – LỚP 9",
    "time_str": "Thời gian làm bài: 90 phút (Không kể thời gian phát đề)",
    "official_label": "ĐỀ CHÍNH THỨC",
    "code": "MÃ ĐỀ 101",
    "footer_note": "Học sinh được sử dụng bảng tuần hoàn các nguyên tố hoá học, bảng hoá trị của các nguyên tố và nhóm nguyên tử, bảng tính tan của các chất trong nước."
}

# -----------------------------------------------------------------------------
# PHẦN I. (3,0 điểm) Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi chỉ chọn một phương án.
# (Mỗi câu đúng 0,25 điểm)
# -----------------------------------------------------------------------------
PART1_QUESTIONS = [
    {
        "id": 1,
        "content": "Thế năng trọng trường của một vật được xác định bởi biểu thức nào?",
        "options": [
            "A. W~t~ = Ph.",
            "B. W~t~ = mv^2^.",
            "C. W~t~ = mh.",
            "D. W~t~ = 10Ph."
        ],
        "correct": "A",
        "solution": "Thế năng trọng trường của một vật ở độ cao h so với mốc chọn thế năng được xác định theo công thức: W~t~ = P·h = m·g·h (trong đó P là trọng lượng của vật, h là độ cao so với mốc thế năng)."
    },
    {
        "id": 2,
        "content": "Phát biểu nào sau đây là đúng khi nói về công suất?",
        "options": [
            "A. Công suất được xác định bằng công thực hiện khi vật dịch chuyển được 1 mét.",
            "B. Công suất là đại lượng đặc trưng cho tốc độ thực hiện công.",
            "C. Công suất được xác định bằng công thực hiện được trong một thời gian nhất định.",
            "D. Công suất được xác định bằng lực tác dụng trong 1 giây."
        ],
        "correct": "B",
        "solution": "Công suất là đại lượng đặc trưng cho tốc độ thực hiện công của người hoặc máy móc thiết bị, được xác định bằng công thực hiện được trong một đơn vị thời gian: \\mathcal{P} = A / t."
    },
    {
        "id": 3,
        "content": "Khi tia sáng truyền từ môi trường (1) có chiết suất n~1~ sang môi trường (2) có chiết suất n~2~ với góc tới i thì góc khúc xạ là r. Biểu thức nào sau đây là biểu thức của định luật khúc xạ ánh sáng?",
        "options": [
            "A. n~1~sinr = n~2~sini.",
            "B. n~1~tanr = n~2~tani.",
            "C. n~1~sini = n~2~sinr.",
            "D. n~1~cosr = n~2~cosi."
        ],
        "correct": "C",
        "solution": "Định luật khúc xạ ánh sáng được biểu diễn ở dạng đối xứng thuận tiện là: n~1~·sini = n~2~·sinr (trong đó n~1~ là chiết suất của môi trường chứa tia tới, n~2~ là chiết suất của môi trường chứa tia khúc xạ)."
    },
    {
        "id": 4,
        "content": "Trong các vật sau, vật nào không có động năng?",
        "options": [
            "A. Máy bay đang bay.",
            "B. Hòn bi lăn trên sàn nhà.",
            "C. Quả bóng nằm yên trên mặt sàn.",
            "D. Ô tô đang chạy trên đường."
        ],
        "correct": "C",
        "solution": "Động năng là dạng năng lượng mà một vật có được do chuyển động (W~đ~ = 1/2·m·v^2^). Quả bóng nằm yên trên mặt sàn có vận tốc v = 0 nên động năng bằng 0."
    },
    {
        "id": 5,
        "content": "Dù ở điều kiện thường hay điều kiện nhiệt độ cao thì kim loại nào sau đây đều không tác dụng với khí oxygen và với nước?",
        "options": [
            "A. Aluminium.",
            "B. Gold.",
            "C. Iron.",
            "D. Copper."
        ],
        "correct": "B",
        "solution": "Vàng (Gold - Au) là kim loại quý có tính khử rất yếu (đứng cuối dãy hoạt động hóa học), không bị oxi hóa bởi oxygen ở bất kỳ nhiệt độ nào và hoàn toàn không tác dụng với nước."
    },
    {
        "id": 6,
        "content": "Dãy gồm các kim loại nào sau đây đẩy được nguyên tử H trong phân tử acid để tạo muối và giải phóng khí hydrogen?",
        "options": [
            "A. Sodium, potassium, aluminium, silver.",
            "B. Iron, copper, zinc, potassium.",
            "C. Sodium, iron, zinc, magnesium.",
            "D. Iron, copper, aluminium, silver."
        ],
        "correct": "C",
        "solution": "Kim loại tác dụng được với dung dịch acid (HCl, H~2~SO~4~ loãng) giải phóng H~2~ phải đứng trước hydrogen (H) trong dãy hoạt động hóa học. Dãy C gồm Na, Fe, Zn, Mg đều đứng trước H.\n(Loại A và D vì có Ag; loại B vì có Cu đứng sau H)."
    },
    {
        "id": 7,
        "content": "Để tách các kim loại có mức độ hoạt động hoá học mạnh như Na, K, Al thì phương pháp được sử dụng là",
        "options": [
            "A. Điện phân nóng chảy.",
            "B. Điện phân dung dịch.",
            "C. Thuỷ luyện.",
            "D. Nhiệt luyện."
        ],
        "correct": "A",
        "solution": "Các kim loại có tính khử mạnh (từ đầu dãy đến Al: K, Na, Ca, Mg, Al) có ion rất khó bị khử, phương pháp duy nhất để điều chế chúng là điện phân nóng chảy các hợp chất của chúng (như muối halide hoặc oxide Al~2~O~3~)."
    },
    {
        "id": 8,
        "content": "Phản ứng hóa học nào sau đây dùng để tách sắt ra khỏi hợp chất bằng phương pháp nhiệt luyện?",
        "options": [
            "A. Mg + FeCl~2~ -> MgCl~2~ + Fe.",
            "B. Fe + 2HCl -> FeCl~2~ + H~2~.",
            "C. 2FeCl~3~ --(đpnc)--> 2Fe + 3Cl~2~.",
            "D. Fe~2~O~3~ + 3CO --(t°)--> 2Fe + 3CO~2~."
        ],
        "correct": "D",
        "solution": "Phương pháp nhiệt luyện là phương pháp dùng chất khử (như CO, C, H~2~, Al) để khử ion kim loại trong oxide ở nhiệt độ cao. Phản ứng Fe~2~O~3~ + 3CO --(t°)--> 2Fe + 3CO~2~ là phản ứng cốt lõi trong lò cao luyện gang thép."
    },
    {
        "id": 9,
        "content": "Lựa chọn tính trạng tương phản trong các ví dụ sau ở đậu hà lan:",
        "options": [
            "A. Hạt trơn và hạt vàng.",
            "B. Hoa tím và hoa ở ngọn.",
            "C. Quả có ngấn và quả xanh.",
            "D. Hạt trơn và hạt nhăn."
        ],
        "correct": "D",
        "solution": "Cặp tính trạng tương phản là hai trạng thái biểu hiện trái ngược nhau của cùng một loại tính trạng. Ở đậu hà lan, hạt trơn và hạt nhăn là hai trạng thái đối lập về hình dạng vỏ hạt."
    },
    {
        "id": 10,
        "content": "Theo Mendel thì kiểu gene AA sẽ tạo ra:",
        "options": [
            "A. Tỉ lệ giao tử A sẽ nhiều hơn giao tử a.",
            "B. 50% giao tử A, 50% giao tử a.",
            "C. Chỉ tạo 100% giao tử A hoặc 100% giao tử a.",
            "D. Chỉ tạo 100% giao tử A."
        ],
        "correct": "D",
        "solution": "Cá thể có kiểu gene đồng hợp trội AA chỉ chứa một loại allele A nên qua giảm phân bình thường chỉ tạo ra 1 loại giao tử duy nhất mang allele A với tỉ lệ 100%."
    },
    {
        "id": 11,
        "content": "Nucleic acid bao gồm:",
        "options": [
            "A. DNA, RNA.",
            "B. DNA, protein.",
            "C. RNA, protein.",
            "D. DNA, RNA, amino acid."
        ],
        "correct": "A",
        "solution": "Nucleic acid là nhóm đại phân tử sinh học mang thông tin di truyền, bao gồm hai loại chính là DNA (deoxyribonucleic acid) và RNA (ribonucleic acid)."
    },
    {
        "id": 12,
        "content": "Gene bình thường đã bị đột biến một cặp nucleotide tạo ra gene đột biến như sau:\n   Gene bình thường:\n   3'- T - A - C - G - C - A - T - 5'\n   5'- A - T - G - C - G - T - A - 3'\n   Gene đột biến:\n   3'- T - A - C - C - A - T - 5'\n   5'- A - T - G - G - T - A - 3'\nDạng đột biến gene đã diễn ra ở trên là:",
        "options": [
            "A. Mất 1 cặp nucleotide AT.",
            "B. Mất 1 cặp nucleotide GC.",
            "C. Thêm 1 cặp nucleotide AT.",
            "D. Thêm 1 cặp nucleotide GC."
        ],
        "correct": "B",
        "solution": "Gene ban đầu có 7 cặp nucleotide. Gene đột biến chỉ còn 6 cặp nucleotide do bị mất đi cặp nucleotide G - C ở vị trí thứ 4. Do đó đây là dạng đột biến mất 1 cặp nucleotide GC."
    }
]

# -----------------------------------------------------------------------------
# PHẦN II. (2,0 điểm) Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d)
# ở mỗi câu, thí sinh chọn đúng hoặc sai.
# -----------------------------------------------------------------------------
PART2_QUESTIONS = [
    {
        "id": 1,
        "content": "Cho các thí nghiệm sau:\n- Thí nghiệm 1: Đốt dây sắt cuộn thành hình lò xo rồi đưa nhanh vào bình chứa khí chlorine.\n- Thí nghiệm 2: Dùng dao cắt một mẩu nhỏ sodium, lau sạch mẩu sodium sau đó cho sodium vào cốc thuỷ tinh chứa dung dịch copper (II) nitrate.\n- Thí nghiệm 3: Ngâm một phoi đồng vào dung dịch silver nitrate trong một thời gian.\n- Thí nghiệm 4: Cắt mẩu nhỏ dây magnesium rồi cho vào dung dịch sulfuric acid loãng.",
        "sub_items": [
            {
                "id": "a",
                "text": "Những tính chất hoá học mà kim loại nói chung thể hiện trong các thí nghiệm trên là: tác dụng với phi kim, tác dụng với nước, tác dụng với dung dịch acid và tác dụng với dung dịch muối.",
                "answer": True,
                "reason": "TN1: Fe + Cl~2~ (tác dụng với phi kim); TN2: Na + H~2~O trong dung dịch (tác dụng với nước); TN3: Cu + AgNO~3~ (tác dụng với dung dịch muối); TN4: Mg + H~2~SO~4~ loãng (tác dụng với dung dịch acid)."
            },
            {
                "id": "b",
                "text": "Phương trình hoá học xảy ra ở các thí nghiệm là:\n   Thí nghiệm 1: 2Fe + 3Cl~2~ --(t°)--> 2FeCl~3~\n   Thí nghiệm 2: 2Na + CuSO~4~ -> Na~2~SO~4~ + Cu\n   Thí nghiệm 3: Cu + 2AgNO~3~ -> Cu(NO~3~)~2~ + 2Ag\n   Thí nghiệm 4: Mg + H~2~SO~4~ -> MgSO~4~ + H~2~",
                "answer": False,
                "reason": "Ở thí nghiệm 2, dung dịch đề cho là copper (II) nitrate Cu(NO~3~)~2~ (chứ không phải CuSO~4~). Mặt khác, Na phản ứng mãnh liệt với nước trước: 2Na + 2H~2~O -> 2NaOH + H~2~, sau đó 2NaOH + Cu(NO~3~)~2~ -> Cu(OH)~2~ + 2NaNO~3~, Na không trực tiếp đẩy Cu ra khỏi muối."
            },
            {
                "id": "c",
                "text": "Ở thí nghiệm 4, magnesium tan ra trong dung dịch acid, có các bọt khí không màu mùi hắc thoát ra.",
                "answer": False,
                "reason": "Khí thoát ra là hydrogen (H~2~), đây là chất khí không màu, không mùi, không vị. Khí mùi hắc là SO~2~ chỉ sinh ra khi dùng H~2~SO~4~ đặc nóng."
            },
            {
                "id": "d",
                "text": "Ở thí nghiệm 3, sau một thời gian xuất hiện kim loại màu đỏ và màu xanh lam của dung dịch ban đầu càng ngày càng nhạt dần.",
                "answer": False,
                "reason": "Dung dịch AgNO~3~ ban đầu không màu. Khi Cu tan vào dung dịch tạo Cu(NO~3~)~2~ làm dung dịch chuyển sang màu xanh lam, và kim loại bạc (Ag) màu trắng bạc sinh ra bám trên thanh đồng (không phải kim loại màu đỏ)."
            }
        ]
    },
    {
        "id": 2,
        "content": "Cho các hình sau mô tả các loại nucleic acid và quá trình phiên mã:",
        "image": "khtn9_ntp_cau2_dna_rna.png",
        "sub_items": [
            {
                "id": "a",
                "text": "Hình 1 là loại deoxyribonucleic acid và hình 2 là loại ribonucleic acid.",
                "answer": True,
                "reason": "Hình 1 mô tả cấu trúc chuỗi xoắn kép của DNA (deoxyribonucleic acid) và Hình 2 mô tả cấu trúc một chuỗi xoắn đơn của RNA (ribonucleic acid)."
            },
            {
                "id": "b",
                "text": "Deoxyribonucleic acid có 4 loại đơn phân là adenine, thymine, cytosine và guanine.",
                "answer": True,
                "reason": "Bốn loại đơn phân nucleotide cấu tạo nên phân tử DNA là A (adenine), T (thymine), C (cytosine) và G (guanine)."
            },
            {
                "id": "c",
                "text": "Hình 3 mô tả quá trình phiên mã, diễn ra trên 2 mạch khuôn của gene.",
                "answer": False,
                "reason": "Quá trình phiên mã tổng hợp phân tử RNA chỉ sử dụng một mạch khuôn duy nhất của gene có chiều 3' -> 5' (không diễn ra trên cả 2 mạch)."
            },
            {
                "id": "d",
                "text": "Một lần phiên mã tạo ra 2 phân tử RNA khác nhau.",
                "answer": False,
                "reason": "Một lần phiên mã từ mạch khuôn của gene chỉ tạo ra 1 phân tử RNA đơn lẻ."
            }
        ]
    }
]

# -----------------------------------------------------------------------------
# PHẦN III. (2,0 điểm) Thí sinh trả lời từ câu 1 đến câu 4. (Mỗi câu đúng 0,5 điểm)
# -----------------------------------------------------------------------------
PART3_QUESTIONS = [
    {
        "id": 1,
        "content": "Trong một công trình xây dựng, một cần cẩu nâng một kiện hàng có trọng lượng 27000N từ mặt đất lên độ cao 5 m trong 1 phút. Cho lực nâng bằng trọng lượng của kiện hàng. Công suất của cần cẩu khi thực hiện công việc này là bao nhiêu kW?",
        "answer": "2,25 kW",
        "solution": "- Đổi thời gian: t = 1 phút = 60 giây.\n- Công của cần cẩu thực hiện để nâng kiện hàng lên cao 5 m là:\n  A = F·s = P·h = 27000 · 5 = 135 000 (J).\n- Công suất của cần cẩu là:\n  \\mathcal{P} = A / t = 135 000 / 60 = 2250 (W) = 2,25 (kW).\n- Đáp số: 2,25 kW (hoặc 2,25)."
    },
    {
        "id": 2,
        "content": "Một tia sáng truyền từ không khí (nước có chiết suất n~1~ = 1) vào nước (chiết suất n~2~ = 1,33). Góc tới trong không khí là i = 40°. Góc khúc xạ r trong nước là bao nhiêu độ? Làm tròn đến 1 chữ số thập phân.",
        "answer": "28,9°",
        "solution": "- Áp dụng định luật khúc xạ ánh sáng: n~1~·sini = n~2~·sinr\n- Thay số: 1 · sin(40°) = 1,33 · sinr\n  => sinr = sin(40°) / 1,33 = 0,642788 / 1,33 \\approx 0,4833\n  => r = arcsin(0,4833) \\approx 28,9°.\n- Đáp số: 28,9° (hoặc 28,9)."
    },
    {
        "id": 3,
        "content": "Một phân tử DNA tái bản liên tiếp 2 lần sẽ tạo thành bao nhiêu phân tử DNA con?",
        "answer": "4 phân tử DNA con",
        "solution": "- Số phân tử DNA con tạo thành sau k lần nhân đôi (tái bản) từ 1 phân tử DNA ban đầu được tính theo công thức: N = 2^k.\n- Với k = 2 lần: N = 2^2 = 4 (phân tử DNA con).\n- Đáp số: 4."
    },
    {
        "id": 4,
        "content": "Cho một mạch khuôn của một gene có trình tự nucleotide là:\n   3'- T - C - C - A - T - T - G - A - A - T - A - T - A - 5'.\nHỏi trên cả 2 mạch của gene sẽ có số nucleotide loại Thymine là bao nhiêu?",
        "answer": "10 nucleotide",
        "solution": "- Đếm số nucleotide trên mạch khuôn (mạch 1):\n  + Số nucleotide loại Thymine: T~1~ = 5 (tại các vị trí 1, 5, 6, 10, 12).\n  + Số nucleotide loại Adenine: A~1~ = 5 (tại các vị trí 4, 8, 9, 11, 13).\n- Theo nguyên tắc bổ sung giữa 2 mạch của phân tử DNA kép:\n  A~1~ = T~2~ = 5 và T~1~ = A~2~ = 5.\n- Tổng số nucleotide loại Thymine trên cả 2 mạch của gene là:\n  T = T~1~ + T~2~ = T~1~ + A~1~ = 5 + 5 = 10 (nucleotide).\n- Đáp số: 10."
    }
]

# -----------------------------------------------------------------------------
# PHẦN IV: (3,0 điểm) Thí sinh trả lời từ câu 1 đến câu 3. Thí sinh viết câu trả lời dưới dạng tự luận.
# -----------------------------------------------------------------------------
ESSAY_QUESTIONS = [
    {
        "id": 1,
        "score_str": "(0,5 điểm)",
        "content": "Viết biểu thức tính cơ năng của một vật. Giải thích các đại lượng trong công thức.",
        "solution": "• Biểu thức tính cơ năng của vật:\n  W = W~đ~ + W~t~ = 1/2·m·v^2^ + P·h  (hoặc W = 1/2·m·v^2^ + m·g·h)\n• Giải thích các đại lượng và đơn vị:\n  - W: Cơ năng của vật (đơn vị: Jun, kí hiệu J).\n  - W~đ~: Động năng của vật (J), trong đó m là khối lượng của vật (kg), v là tốc độ của vật (m/s).\n  - W~t~: Thế năng trọng trường của vật (J), trong đó P là trọng lượng của vật (N) hoặc g là gia tốc trọng trường (m/s^2^), h là độ cao của vật so với mốc chọn tính thế năng (m)."
    },
    {
        "id": 2,
        "score_str": "(0,5 điểm)",
        "content": "Dự đoán hiện tượng hoá học sẽ xảy ra khi cho kim loại barium vào dung dịch copper (II) sulfate. Viết phương trình hoá học minh hoạ.",
        "solution": "• Dự đoán hiện tượng:\n  - Mẩu kim loại barium (Ba) tan dần, có bọt khí không màu (H~2~) thoát ra mạnh và tỏa nhiều nhiệt.\n  - Xuất hiện đồng thời kết tủa màu trắng (BaSO~4~) và kết tủa màu xanh lam (Cu(OH)~2~).\n  - Màu xanh lam của dung dịch CuSO~4~ ban đầu nhạt dần.\n• Phương trình hoá học minh hoạ:\n  (1) Barium phản ứng với nước trong dung dịch trước:\n      Ba + 2H~2~O -> Ba(OH)~2~ + H~2~ ^\n  (2) Dung dịch kiềm Ba(OH)~2~ sinh ra phản ứng trao đổi với muối CuSO~4~:\n      Ba(OH)~2~ + CuSO~4~ -> BaSO~4~ v (trắng) + Cu(OH)~2~ v (xanh lam)"
    },
    {
        "id": 3,
        "score_str": "(2,0 điểm)",
        "content": "a) (1,0 điểm) Ở đậu hà lan, allele A quy định tính trạng hạt trơn trội hoàn toàn so với allele a quy định tính trạng hạt nhăn, biết không xảy ra đột biến. Để biết cây hạt trơn thuần chủng hay không, các nhà khoa học đã lai phân tích cây hạt trơn với cây hạt nhăn. Em hãy viết sơ đồ lai minh họa cho phép lai phân tích trên.\nb) (1,0 điểm) Cho đoạn mạch khuôn của một gene ở vi khuẩn có trình tự nucleotide như sau:\n   3'- T - C - C - C - C - A - T - T - G - A - A - T - A - A - A - 5'.\n- Xác định trình tự nucleotide của mRNA được phiên mã từ gene này.\n- Xác định trình tự đoạn amino acid được dịch mã từ gene này.\nBiết trên mRNA các bộ ba mã hóa các amino acid như sau:\nUUU: Phenylalanine, GGU: Glycine, AAC: Asparagine, AGG: Arginine, UUA: Leucine.",
        "solution": "a) Phép lai phân tích cây hạt trơn:\n• Khái niệm: Lai phân tích là phép lai giữa cá thể mang tính trạng trội cần xác định kiểu gene với cá thể mang tính trạng lặn (cây hạt nhăn, kiểu gene aa).\n• Cây hạt trơn có thể có kiểu gene đồng hợp trội (thuần chủng: AA) hoặc dị hợp (không thuần chủng: Aa).\n- Trường hợp 1: Cây hạt trơn thuần chủng (AA)\n  + Sơ đồ lai:\n    P:  Hạt trơn (AA)   x   Hạt nhăn (aa)\n    G~P~:     A                     a\n    F~a~:               100% Aa (100% hạt trơn)\n  + Kết luận: Đời con F~a~ đồng tính 100% hạt trơn thì cây hạt trơn đem lai là thuần chủng (AA).\n- Trường hợp 2: Cây hạt trơn không thuần chủng (Aa)\n  + Sơ đồ lai:\n    P:  Hạt trơn (Aa)   x   Hạt nhăn (aa)\n    G~P~: 1/2 A : 1/2 a             a\n    F~a~:       1/2 Aa (hạt trơn) : 1/2 aa (hạt nhăn)\n  + Kết luận: Đời con F~a~ phân tính theo tỉ lệ 1 hạt trơn : 1 hạt nhăn (50% trơn : 50% nhăn) thì cây hạt trơn đem lai là không thuần chủng (Aa).\n\nb) Xác định trình tự nucleotide và chuỗi amino acid:\n• Trình tự nucleotide của phân tử mRNA:\n- Nguyên tắc bổ sung trong phiên mã (chiều mRNA ngược với mạch khuôn: 5' -> 3'):\n  A~khuôn~ liên kết với U~mRNA~; T~khuôn~ liên kết với A~mRNA~;\n  G~khuôn~ liên kết với C~mRNA~; C~khuôn~ liên kết với G~mRNA~.\n- Mạch khuôn của gene (3' -> 5'):\n  3'- T - C - C - C - C - A - T - T - G - A - A - T - A - A - A - 5'\n- Phân tử mRNA được tổng hợp (5' -> 3'):\n  5'- A - G - G - G - G - U - A - A - C - U - U - A - U - U - U - 3'\n• Trình tự đoạn amino acid dịch mã:\n- Chia mạch mRNA thành từng bộ ba (codon) từ đầu 5' sang 3':\n  + Codon 1: 5'-AGG-3' mã hóa Arginine.\n  + Codon 2: 5'-GGU-3' mã hóa Glycine.\n  + Codon 3: 5'-AAC-3' mã hóa Asparagine.\n  + Codon 4: 5'-UUA-3' mã hóa Leucine.\n  + Codon 5: 5'-UUU-3' mã hóa Phenylalanine.\n- Trình tự chuỗi amino acid được dịch mã hoàn chỉnh là:\n  Arginine - Glycine - Asparagine - Leucine - Phenylalanine\n  (viết tắt: Arg - Gly - Asn - Leu - Phe)."
    }
]
