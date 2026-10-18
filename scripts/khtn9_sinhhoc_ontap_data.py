# -*- coding: utf-8 -*-
"""
Dữ liệu Đề kiểm tra / Ôn tập môn KHTN 9 - Phân môn Sinh học
Trường THCS Kim Long - Thầy TRẦN VĂN THẠNH - 0777.470.803
Biên soạn chuẩn theo "FORM ĐỀ KIỂM TRA" từ ảnh chụp đề thi thực tế.
Gồm:
- Phần I: Trắc nghiệm đúng sai (5 câu)
- Phần II: Câu hỏi trả lời ngắn (15 câu)
- Phần III: Tự luận (4 câu có hình vẽ cấu trúc DNA)
"""

import os

ROOT_DIR = r"c:\Antigravity_Thanh"
IMG_DNA_PATH = os.path.join(ROOT_DIR, "San_Pham", "images_khtn9", "hinh38_1_dna_hd.png")

INFO = {
    "school": "TRƯỜNG THCS KIM LONG",
    "year": "NĂM HỌC 2026 - 2027",
    "title": "ĐỀ CƯƠNG ÔN TẬP KIỂM TRA ĐỊNH KÌ – KHTN 9",
    "subject": "MÔN: KHOA HỌC TỰ NHIÊN 9 - PHÂN MÔN SINH HỌC",
    "time_str": "Chương trình GDPT 2018 - Di truyền Mendel, DNA & RNA",
}

# =============================================================================
# PHẦN I: TRẮC NGHIỆM ĐÚNG SAI (5 CÂU - 20 Ý)
# =============================================================================
TF_QUESTIONS = [
    {
        "id": 1,
        "content": "Khi tìm hiểu các thuật ngữ cơ bản trong di truyền học Mendel. Hãy xác định các nhận định sau là đúng hay sai.",
        "statements": [
            {
                "id": "a",
                "text": "Tính trạng là những đặc điểm về hình thái, cấu tạo và sinh lí của một cơ thể.",
                "correct": True,
                "explain": "Đúng. Đây là định nghĩa khoa học chuẩn xác về tính trạng của cơ thể sinh vật theo SGK KHTN 9."
            },
            {
                "id": "b",
                "text": "Cặp tính trạng tương phản là hai trạng thái biểu hiện giống nhau của cùng một loại tính trạng.",
                "correct": False,
                "explain": "Sai. Cặp tính trạng tương phản là hai trạng thái biểu hiện trái ngược nhau (đối lập nhau) của cùng một loại tính trạng (ví dụ: thân cao đối lập với thân thấp, hạt vàng đối lập với hạt xanh)."
            },
            {
                "id": "c",
                "text": "Kiểu hình là tổ hợp toàn bộ các tính trạng của cơ thể sinh vật.",
                "correct": True,
                "explain": "Đúng. Kiểu hình là tập hợp tất cả các tính trạng của cơ thể (trên thực tế thường xét một hoặc vài tính trạng cần nghiên cứu)."
            },
            {
                "id": "d",
                "text": "Allele là các trạng thái biểu hiện khác nhau của cùng một gene.",
                "correct": True,
                "explain": "Đúng. Allele là các trạng thái tồn tại khác nhau của cùng một gene xác định (ví dụ allele A quy định hoa tím, allele a quy định hoa trắng)."
            }
        ],
        "explanation": "• Ý a ĐÚNG: Tính trạng phản ánh các đặc tính hình thái, giải phẫu và sinh lí của sinh vật.\n• Ý b SAI: Cặp tính trạng tương phản phải là hai trạng thái biểu hiện trái ngược nhau.\n• Ý c ĐÚNG: Định nghĩa chuẩn về kiểu hình trong di truyền học.\n• Ý d ĐÚNG: Allele là các trạng thái biểu hiện khác nhau của cùng một locus gene."
    },
    {
        "id": 2,
        "content": "Ở đậu Hà Lan, Menđen tiến hành lai cây hoa tím thuần chủng với cây hoa trắng thuần chủng. Biết hoa tím trội hoàn toàn so với hoa trắng. Hãy xác định các nhận định sau là đúng hay sai.",
        "statements": [
            {
                "id": "a",
                "text": "Cây hoa tím thuần chủng có kiểu gene đồng hợp về gene quy định màu hoa.",
                "correct": True,
                "explain": "Đúng. Dòng thuần chủng mang cặp gene đồng hợp trội (AA)."
            },
            {
                "id": "b",
                "text": "F~1~ có 100% cây hoa tím.",
                "correct": True,
                "explain": "Đúng. Phép lai P~tc~: AA × aa → F~1~: 100% Aa (100% kiểu hình hoa tím)."
            },
            {
                "id": "c",
                "text": "Cho các cây F~1~ tự thụ phấn, F~2~ có tỉ lệ kiểu hình 3 hoa tím : 1 hoa trắng.",
                "correct": True,
                "explain": "Đúng. F~1~ × F~1~: Aa × Aa → F~2~: 1 AA : 2 Aa : 1 aa (3 cây hoa tím : 1 cây hoa trắng)."
            },
            {
                "id": "d",
                "text": "Ở F~2~, tất cả các cây hoa tím đều có kiểu gene đồng hợp.",
                "correct": False,
                "explain": "Sai. Ở F~2~, các cây hoa tím gồm 2 loại kiểu gene: đồng hợp trội AA (chiếm 1/3 số cây hoa tím) và dị hợp Aa (chiếm 2/3 số cây hoa tím)."
            }
        ],
        "explanation": "• Ý a ĐÚNG: Cây hoa tím thuần chủng mang kiểu gene đồng hợp AA.\n• Ý b ĐÚNG: Đời con F~1~ 100% mang kiểu gene Aa và biểu hiện kiểu hình hoa tím.\n• Ý c ĐÚNG: Tỉ lệ phân li kiểu hình ở F~2~ là 3 trội (hoa tím) : 1 lặn (hoa trắng).\n• Ý d SAI: Trong số các cây hoa tím ở F~2~, có 1/3 là đồng hợp (AA) và 2/3 là dị hợp (Aa)."
    },
    {
        "id": 3,
        "content": "Về phép lai phân tích, hãy xác định các nhận định sau là đúng hay sai.",
        "statements": [
            {
                "id": "a",
                "text": "Lai phân tích là phép lai giữa cơ thể mang tính trạng trội chưa biết kiểu gene với cơ thể mang tính trạng lặn.",
                "correct": True,
                "explain": "Đúng. Đây là định nghĩa chuẩn của phép lai phân tích trong di truyền học Mendel."
            },
            {
                "id": "b",
                "text": "Lai phân tích giúp xác định kiểu gene của cơ thể mang tính trạng trội.",
                "correct": True,
                "explain": "Đúng. Mục đích chính của lai phân tích là xác định cơ thể mang tính trạng trội là đồng hợp tử hay dị hợp tử."
            },
            {
                "id": "c",
                "text": "Nếu phép lai phân tích cho đời con có 100% kiểu hình trội thì cơ thể cần kiểm tra chắc chắn có kiểu gene đồng hợp trội, trong điều kiện di truyền Mendel đơn gene và trội hoàn toàn.",
                "correct": True,
                "explain": "Đúng. Phép lai: AA × aa → 100% Aa (100% kiểu hình trội). Khi đời con đồng tính trội thì cá thể trội cần kiểm tra có kiểu gene đồng hợp trội AA."
            },
            {
                "id": "d",
                "text": "Nếu cơ thể cần kiểm tra có kiểu gene dị hợp thì phép lai phân tích cho đời con có tỉ lệ kiểu hình 3 trội : 1 lặn.",
                "correct": False,
                "explain": "Sai. Cơ thể dị hợp lai phân tích: Aa × aa → đời con phân li theo tỉ lệ 1 trội : 1 lặn (1 Aa : 1 aa), không phải 3 trội : 1 lặn."
            }
        ],
        "explanation": "• Ý a ĐÚNG: Định nghĩa chuẩn xác của phép lai phân tích.\n• Ý b ĐÚNG: Ứng dụng chính của lai phân tích là xác định kiểu gene của cá thể trội.\n• Ý c ĐÚNG: Đời con 100% trội chứng tỏ cá thể mang tính trạng trội chỉ tạo 1 loại giao tử A → kiểu gene AA.\n• Ý d SAI: Cá thể dị hợp Aa lai phân tích tạo đời con có tỉ lệ phân li 1 trội : 1 lặn (50% trội : 50% lặn)."
    },
    {
        "id": 4,
        "content": "Về cấu tạo hóa học của DNA, hãy xác định các nhận định sau là đúng hay sai.",
        "statements": [
            {
                "id": "a",
                "text": "DNA là đại phân tử có cấu tạo theo nguyên tắc đa phân.",
                "correct": True,
                "explain": "Đúng. Phân tử DNA có kích thước và khối lượng lớn, cấu tạo từ nhiều đơn phân hợp lại theo nguyên tắc đa phân."
            },
            {
                "id": "b",
                "text": "Đơn phân cấu tạo nên DNA là nucleotide.",
                "correct": True,
                "explain": "Đúng. Mỗi đơn phân của DNA là một nucleotide (gồm gốc phosphate, đường deoxyribose và một base nitrogen)."
            },
            {
                "id": "c",
                "text": "DNA được cấu tạo từ bốn loại nucleotide là A, T, G và C.",
                "correct": True,
                "explain": "Đúng. Bốn loại nucleotide cấu tạo nên DNA là Adenine (A), Thymine (T), Guanine (G) và Cytosine (C)."
            },
            {
                "id": "d",
                "text": "Trong DNA, nucleotide A liên kết bổ sung với nucleotide G, còn T liên kết với C.",
                "correct": False,
                "explain": "Sai. Theo nguyên tắc bổ sung, A liên kết với T bằng 2 liên kết hydrogen, còn G liên kết với C bằng 3 liên kết hydrogen."
            }
        ],
        "explanation": "• Ý a ĐÚNG: DNA là đại phân tử sinh học cấu tạo theo nguyên tắc đa phân.\n• Ý b ĐÚNG: Đơn phân của phân tử DNA là các nucleotide.\n• Ý c ĐÚNG: Bốn loại nucleotide của DNA là Adenine (A), Thymine (T), Guanine (G), Cytosine (C).\n• Ý d SAI: Nguyên tắc bổ sung chuẩn xác là A liên kết với T và G liên kết với C."
    },
    {
        "id": 5,
        "content": "Về cấu tạo của RNA, hãy xác định các nhận định sau là đúng hay sai.",
        "statements": [
            {
                "id": "a",
                "text": "RNA được cấu tạo theo nguyên tắc đa phân, đơn phân là nucleotide.",
                "correct": True,
                "explain": "Đúng. Phân tử RNA cũng là một đại phân tử đa phân mà đơn phân là các ribonucleotide."
            },
            {
                "id": "b",
                "text": "RNA có bốn loại nucleotide là A, T, G và C.",
                "correct": False,
                "explain": "Sai. Bốn loại nucleotide cấu tạo nên RNA là Adenine (A), Uracil (U), Guanine (G) và Cytosine (C); RNA không chứa nucleotide loại T."
            },
            {
                "id": "c",
                "text": "RNA có nucleotide U thay cho nucleotide T của DNA.",
                "correct": True,
                "explain": "Đúng. Điểm khác biệt quan trọng giữa RNA và DNA là RNA chứa base Uracil (U) thay cho Thymine (T)."
            },
            {
                "id": "d",
                "text": "Các nucleotide trên mạch RNA liên kết với nhau bằng liên kết cộng hóa trị.",
                "correct": True,
                "explain": "Đúng. Trên một chuỗi ribonucleotide, các nucleotide liên kết với nhau bằng liên kết cộng hóa trị phosphodiester bền vững."
            }
        ],
        "explanation": "• Ý a ĐÚNG: RNA cấu tạo theo nguyên tắc đa phân, đơn phân là ribonucleotide.\n• Ý b SAI: Bốn loại đơn phân của RNA là A, U, G, C (không có nucleotide loại T).\n• Ý c ĐÚNG: Nucleotide loại U trong RNA thay thế vị trí của nucleotide loại T trong DNA.\n• Ý d ĐÚNG: Các nucleotide dọc theo mạch đơn liên kết bằng liên kết cộng hóa trị phosphodiester."
    }
]

# =============================================================================
# PHẦN II: CÂU HỎI TRẢ LỜI NGẮN (15 CÂU)
# =============================================================================
SHORT_QUESTIONS_A = [
    {
        "id": 1,
        "content": "Ở đậu Hà Lan, lai cây hoa tím thuần chủng với cây hoa trắng thuần chủng, biết hoa tím trội hoàn toàn so với hoa trắng. F~1~ có 100% hoa tím. Hỏi tỉ lệ cây hoa trắng ở F~2~ là bao nhiêu phần trăm?",
        "correct": "25%",
        "explanation": "P~tc~: AA (hoa tím) × aa (hoa trắng) → F~1~: Aa (100% hoa tím).\nF~1~ × F~1~: Aa × Aa → F~2~: 1 AA : 2 Aa : 1 aa (3 hoa tím : 1 hoa trắng).\nTỉ lệ cây hoa trắng (aa) ở F~2~ chiếm 1/4 = 25%."
    },
    {
        "id": 2,
        "content": "Cho hai cây F~1~ dị hợp về một cặp gene lai với nhau (Aa × Aa). Hỏi đời con có bao nhiêu kiểu gene khác nhau?",
        "correct": "3",
        "explanation": "Phép lai Aa × Aa tạo ra đời con có 3 loại kiểu gene khác nhau gồm: AA, Aa và aa với tỉ lệ tương ứng là 1 : 2 : 1."
    },
    {
        "id": 3,
        "content": "Phép lai Aa × Aa tạo ra đời con có tỉ lệ kiểu hình 3 trội : 1 lặn. Hỏi tỉ lệ kiểu hình trội là bao nhiêu phần trăm?",
        "correct": "75%",
        "explanation": "Kiểu hình trội chiếm 3 phần trên tổng số 4 phần (3/4).\nTỉ lệ phần trăm kiểu hình trội là: (3 / 4) × 100% = 75%."
    },
    {
        "id": 4,
        "content": "Một cơ thể có kiểu gene Aa giảm phân bình thường. Hỏi cơ thể đó tạo ra bao nhiêu loại giao tử về gene đang xét?",
        "correct": "2",
        "explanation": "Cơ thể dị hợp 1 cặp gene Aa giảm phân bình thường phân li tạo ra 2 loại giao tử mang allele đơn là: giao tử mang allele A và giao tử mang allele a (tỉ lệ 1 : 1)."
    },
    {
        "id": 5,
        "content": "Một cơ thể có kiểu gene AaBb, hai cặp gene phân li độc lập và trội hoàn toàn. Hỏi cơ thể đó tạo ra bao nhiêu loại giao tử?",
        "correct": "4",
        "explanation": "Số loại giao tử tạo ra được tính theo công thức 2^n (với n là số cặp gene dị hợp).\nỞ đây n = 2 → Số loại giao tử = 2^2 = 4 loại giao tử gồm: AB, Ab, aB, ab (mỗi loại chiếm 25%)."
    },
    {
        "id": 6,
        "content": "Một phép lai phân tích giữa cơ thể Aa và cơ thể aa tạo ra đời con có tỉ lệ kiểu hình trội : kiểu hình lặn là 1 : 1. Hỏi tỉ lệ phần trăm đời con mang kiểu hình lặn là bao nhiêu?",
        "correct": "50%",
        "explanation": "Phép lai phân tích: Aa × aa → Đời con: 1 Aa (trội) : 1 aa (lặn).\nTỉ lệ phần trăm cá thể mang kiểu hình lặn (aa) là: 1/2 × 100% = 50%."
    },
    {
        "id": 7,
        "content": "Một cơ thể có kiểu gene AaBb, hai cặp gene phân li độc lập. Hỏi cơ thể đó có bao nhiêu loại kiểu gene trong các giao tử được tạo ra?",
        "correct": "4",
        "explanation": "Giao tử của cơ thể mang kiểu gene đơn bội gồm 1 allele của mỗi cặp gene. Cơ thể AaBb tạo ra 4 loại tổ hợp allele khác nhau trong giao tử: AB, Ab, aB, ab (tương ứng với 4 loại kiểu gene giao tử)."
    }
]

SHORT_QUESTIONS_B = [
    {
        "id": 8,
        "content": "Một phân tử DNA có 1.000 cặp nucleotide. Hỏi phân tử DNA đó có tổng cộng bao nhiêu nucleotide?",
        "correct": "2.000",
        "explanation": "Mỗi cặp nucleotide gồm 2 nucleotide liên kết với nhau giữa hai mạch.\nTổng số nucleotide của phân tử DNA là: N = 1.000 × 2 = 2.000 nucleotide."
    },
    {
        "id": 9,
        "content": "Một phân tử DNA có 2.000 nucleotide, trong đó có 400 nucleotide loại A. Hỏi phân tử DNA đó có bao nhiêu nucleotide loại T?",
        "correct": "400",
        "explanation": "Theo nguyên tắc bổ sung trong phân tử DNA, nucleotide loại A luôn liên kết bổ sung với nucleotide loại T, do đó số lượng nucleotide loại T bằng số lượng loại A: T = A = 400 nucleotide."
    },
    {
        "id": 10,
        "content": "Một phân tử DNA có 2.000 nucleotide, trong đó có 400 nucleotide loại A. Hỏi tổng số nucleotide loại G và C là bao nhiêu?",
        "correct": "1.200",
        "explanation": "Ta có: A = T = 400 nucleotide → A + T = 400 + 400 = 800 nucleotide.\nTổng số nucleotide loại G và C là: G + C = N - (A + T) = 2.000 - 800 = 1.200 nucleotide (trong đó G = C = 600 nucleotide)."
    },
    {
        "id": 11,
        "content": "Một phân tử DNA có 1.500 cặp nucleotide. Theo mô hình cấu trúc phổ biến, cứ 10 cặp nucleotide tương ứng với một chu kì xoắn. Hỏi phân tử DNA đó có bao nhiêu chu kì xoắn?",
        "correct": "150",
        "explanation": "Số chu kì xoắn của phân tử DNA là: C = (Số cặp nucleotide) / 10 = 1.500 / 10 = 150 chu kì xoắn."
    },
    {
        "id": 12,
        "content": "Một phân tử DNA có 20 chu kì xoắn. Biết mỗi chu kì xoắn gồm 10 cặp nucleotide. Hỏi phân tử DNA đó có tổng cộng bao nhiêu nucleotide?",
        "correct": "400",
        "explanation": "Mỗi chu kì xoắn gồm 10 cặp nucleotide = 20 nucleotide.\nTổng số nucleotide của phân tử DNA là: N = 20 chu kì × 20 nucleotide/chu kì = 400 nucleotide."
    },
    {
        "id": 13,
        "content": "Một phân tử DNA có 1.000 cặp nucleotide, trong đó có 200 nucleotide loại G. Hỏi phân tử DNA đó có bao nhiêu nucleotide loại A?",
        "correct": "800",
        "explanation": "Tổng số nucleotide của DNA là: N = 1.000 × 2 = 2.000 nucleotide.\nTheo nguyên tắc bổ sung: G = C = 200 nucleotide → G + C = 400 nucleotide.\nTa có: A + T = N - (G + C) = 2.000 - 400 = 1.600 nucleotide.\nVì A = T nên số nucleotide loại A là: A = 1.600 / 2 = 800 nucleotide."
    }
]

SHORT_QUESTIONS_C = [
    {
        "id": 14,
        "content": "RNA có 4 loại nucleotide khác nhau. Hỏi có bao nhiêu loại nucleotide của RNA không có trong DNA?",
        "correct": "1",
        "explanation": "Phân tử RNA cấu tạo từ 4 loại nucleotide là: A, U, G, C. Phân tử DNA cấu tạo từ 4 loại nucleotide là: A, T, G, C.\nChỉ có duy nhất 1 loại nucleotide của RNA không có trong DNA là nucleotide loại U (Uracil thay thế cho Thymine)."
    },
    {
        "id": 15,
        "content": "Một phân tử DNA có 800 nucleotide trên một mạch. Hỏi mạch còn lại có bao nhiêu nucleotide?",
        "correct": "800",
        "explanation": "Phân tử DNA gồm hai mạch đơn song song liên kết bổ sung với nhau từng đôi một, do đó số lượng nucleotide trên mạch 1 luôn bằng đúng số lượng nucleotide trên mạch 2: N~mạch 2~ = N~mạch 1~ = 800 nucleotide."
    }
]

# =============================================================================
# PHẦN III: TỰ LUẬN (4 CÂU)
# =============================================================================
ESSAY_QUESTIONS = [
    {
        "id": 1,
        "content": "Ở bí, tính trạng quả tròn trội hoàn toàn so với tính trạng quả dài. Cho giao phấn giữa cây có quả tròn với cây có quả dài thu được F~1~ rồi tiếp tục cho F~1~ giao phấn với nhau.\na. Lập sơ đồ lai từ P → F~2~.\nb. Cho F~1~ lai phân tích thì kết quả tạo ra sẽ như thế nào về kiểu gene và kiểu hình?",
        "has_image": False,
        "solution": (
            "► Quy ước gene:\n"
            "- Theo đề bài, tính trạng quả tròn trội hoàn toàn so với tính trạng quả dài:\n"
            "  + Allele A quy định tính trạng quả tròn (trội hoàn toàn).\n"
            "  + Allele a quy định tính trạng quả dài (lặn).\n"
            "- Cây quả dài mang tính trạng lặn nên có kiểu gene đồng hợp lặn: aa.\n"
            "- Cây quả tròn giao phấn với cây quả dài thu được F~1~ đồng tính rồi tiếp tục F~1~ giao phấn thu được F~2~ phân li kiểu hình theo quy luật Mendel → Cây quả tròn ở P phải thuần chủng mang kiểu gene: AA.\n\n"
            "a. Lập sơ đồ lai từ P → F~2~:\n"
            "- Sơ đồ lai từ P đến F~1~:\n"
            "  P~tc~:       Cây quả tròn (AA)     ×     Cây quả dài (aa)\n"
            "  G~P~:               A                           a\n"
            "  F~1~:       100% Aa (100% Cây quả tròn)\n\n"
            "- Cho F~1~ giao phấn với nhau (F~1~ × F~1~):\n"
            "  F~1~ × F~1~:   Cây quả tròn (Aa)     ×     Cây quả tròn (Aa)\n"
            "  G~F1~:       1/2 A : 1/2 a                1/2 A : 1/2 a\n"
            "  F~2~:\n"
            "  + Tỉ lệ kiểu gene: 1 AA : 2 Aa : 1 aa (25% AA : 50% Aa : 25% aa).\n"
            "  + Tỉ lệ kiểu hình: 3 quả tròn : 1 quả dài (75% Cây quả tròn : 25% Cây quả dài).\n\n"
            "b. Cho F~1~ lai phân tích:\n"
            "- Lai phân tích là phép lai giữa cá thể mang tính trạng trội F~1~ (Aa) với cá thể mang tính trạng lặn (aa):\n"
            "  P~a~:        Cây F~1~ quả tròn (Aa)   ×   Cây quả dài (aa)\n"
            "  G~a~:             1/2 A : 1/2 a                  a\n"
            "  F~a~:\n"
            "  + Tỉ lệ kiểu gene: 1 Aa : 1 aa (50% Aa : 50% aa).\n"
            "  + Tỉ lệ kiểu hình: 1 quả tròn : 1 quả dài (50% Quả tròn : 50% Quả dài)."
        )
    },
    {
        "id": 2,
        "content": "Ở bí, quả tròn, hoa vàng là hai tính trạng trội hoàn toàn so với quả dài, hoa trắng. Sự di truyền của hai tính trạng trên tuân theo quy luật phân li độc lập của Mendel. Cho cây bí quả tròn, hoa vàng thuần chủng lai với cây bí quả dài, hoa trắng. Xác định kiểu gene, kiểu hình của P~tc~ và lập sơ đồ lai từ P~tc~ đến F~2~.",
        "has_image": False,
        "solution": (
            "► Quy ước gene:\n"
            "- Cặp tính trạng dạng quả: Allele A quy định quả tròn >> allele a quy định quả dài.\n"
            "- Cặp tính trạng màu hoa: Allele B quy định hoa vàng >> allele b quy định hoa trắng.\n\n"
            "► Xác định kiểu gene, kiểu hình của P~tc~:\n"
            "- Cây bí quả tròn, hoa vàng thuần chủng có kiểu gene là: AABB.\n"
            "- Cây bí quả dài, hoa trắng mang hai tính trạng lặn nên có kiểu gene là: aabb.\n\n"
            "► Sơ đồ lai từ P~tc~ đến F~2~:\n"
            "- Sơ đồ lai P~tc~ → F~1~:\n"
            "  P~tc~:   Quả tròn, hoa vàng (AABB)   ×   Quả dài, hoa trắng (aabb)\n"
            "  G~P~:                 AB                              ab\n"
            "  F~1~:   100% AaBb (100% Cây quả tròn, hoa vàng)\n\n"
            "- Cho F~1~ tự thụ phấn (hoặc giao phấn với nhau):\n"
            "  F~1~ × F~1~:       AaBb (Quả tròn, hoa vàng)   ×   AaBb (Quả tròn, hoa vàng)\n"
            "  G~F1~:       1/4 AB : 1/4 Ab : 1/4 aB : 1/4 ab       1/4 AB : 1/4 Ab : 1/4 aB : 1/4 ab\n\n"
            "- Kết quả ở F~2~:\n"
            "  + Tỉ lệ kiểu gene: Gồm 9 loại kiểu gene phân li theo tỉ lệ:\n"
            "    1 AABB : 2 AABb : 1 AAbb : 2 AaBB : 4 AaBb : 2 Aabb : 1 aaBB : 2 aaBb : 1 aabb.\n"
            "  + Tỉ lệ kiểu hình: Phân li theo tỉ lệ 9 : 3 : 3 : 1, cụ thể:\n"
            "    * 9/16 Cây quả tròn, hoa vàng (A-B-)\n"
            "    * 3/16 Cây quả tròn, hoa trắng (A-bb)\n"
            "    * 3/16 Cây quả dài, hoa vàng (aaB-)\n"
            "    * 1/16 Cây quả dài, hoa trắng (aabb)."
        )
    },
    {
        "id": 3,
        "content": "Dựa vào hình ảnh bên dưới (Hình 38.1 - Mô hình cấu trúc một đoạn phân tử DNA):\na) Cho biết phân tử DNA có cấu trúc như thế nào?\nb) Gọi tên các đơn phân cấu tạo nên phân tử DNA và cho biết chúng liên kết với nhau bằng cách nào?",
        "has_image": True,
        "image_path": IMG_DNA_PATH,
        "image_caption": "Hình 38.1. Mô hình cấu trúc không gian một đoạn phân tử DNA (SGK KHTN 9)",
        "solution": (
            "a) Cấu trúc không gian của phân tử DNA:\n"
            "- DNA là một đại phân tử sinh học có cấu trúc xoắn kép, gồm 2 mạch polynucleotide song song và ngược chiều nhau (một mạch chạy theo chiều 5' → 3', mạch kia chạy theo chiều 3' → 5'), cùng xoắn đều quanh một trục tưởng tượng theo chiều từ trái sang phải (xoắn phải).\n"
            "- Các kích thước hình học chuẩn của vòng xoắn:\n"
            "  + Đường kính vòng xoắn là 20 Å (Angstrom).\n"
            "  + Chiều dài mỗi chu kì xoắn là 34 Å, gồm đúng 10 cặp nucleotide (trung bình mỗi cặp nucleotide dài 3,4 Å).\n\n"
            "b) Các đơn phân và cách thức liên kết trong phân tử DNA:\n"
            "- Tên các đơn phân: Đơn phân cấu tạo nên phân tử DNA là các nucleotide, gồm 4 loại:\n"
            "  1. Adenine (kí hiệu là A)\n"
            "  2. Thymine (kí hiệu là T)\n"
            "  3. Guanine (kí hiệu là G)\n"
            "  4. Cytosine (kí hiệu là C)\n"
            "- Cách thức liên kết giữa các đơn phân:\n"
            "  + Dọc trên mỗi mạch đơn: Các nucleotide liên kết với nhau bằng liên kết cộng hóa trị phosphodiester bền vững giữa nhóm phosphate của nucleotide này với đường deoxyribose của nucleotide kế tiếp tạo thành chuỗi polynucleotide.\n"
            "  + Giữa hai mạch đơn: Các base nitrogen đối diện liên kết với nhau bằng liên kết hydrogen theo nguyên tắc bổ sung (NTBS):\n"
            "    * Nucleotide loại A liên kết với nucleotide loại T bằng 2 liên kết hydrogen (A = T).\n"
            "    * Nucleotide loại G liên kết với nucleotide loại C bằng 3 liên kết hydrogen (G ≡ C).\n"
            "  => Nguyên tắc bổ sung giúp đảm bảo đường kính vòng xoắn DNA luôn giữ ổn định ở mức 20 Å."
        )
    },
    {
        "id": 4,
        "content": "Trình tự các nucleotide trên một đoạn của DNA như sau:\n- A - G - T - T - C - C - A - T - T - G - C - G - A - A - G - T - C -\nHãy xác định trình tự các nucleotide trên mạch bổ sung với mạch đó.",
        "has_image": False,
        "solution": (
            "► Nguyên tắc bổ sung (NTBS):\n"
            "- Trong cấu trúc phân tử DNA, các nucleotide giữa hai mạch đơn liên kết bổ sung theo nguyên tắc: A liên kết với T (và ngược lại T liên kết với A); G liên kết với C (và ngược lại C liên kết với G).\n\n"
            "► Xác định trình tự mạch bổ sung:\n"
            "- Mạch ban đầu (mạch 1):\n"
            "  - A - G - T - T - C - C - A - T - T - G - C - G - A - A - G - T - C -\n"
            "- Mạch bổ sung (mạch 2):\n"
            "  - T - C - A - A - G - G - T - A - A - C - G - C - T - T - C - A - G -\n\n"
            "► Kết luận: Đoạn phân tử DNA xoắn kép hoàn chỉnh gồm 2 mạch bổ sung:\n"
            "  Mạch 1:  - A - G - T - T - C - C - A - T - T - G - C - G - A - A - G - T - C -\n"
            "           |   |   |   |   |   |   |   |   |   |   |   |   |   |   |   |   |\n"
            "  Mạch 2:  - T - C - A - A - G - G - T - A - A - C - G - C - T - T - C - A - G -"
        )
    }
]
