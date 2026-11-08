# -*- coding: utf-8 -*-
"""
Script biên soạn bộ tài liệu chuẩn:
CHUYÊN ĐỀ: DỊCH MÃ VÀ MỐI QUAN HỆ TỪ GENE ĐẾN TÍNH TRẠNG (SINH HỌC 12)
Bao gồm:
1. Bản Học sinh (.docx) - Chuẩn Form Đề Kiểm Tra, có dòng chấm làm bài.
2. Bản Giáo viên (.docx) - Có lời giải chi tiết và đáp án in đậm màu đỏ.
3. Bản Dạy trên màn hình (.docx & .pdf slide 16:9 widescreen cho Tivi 65 inch).
"""

import os
import sys
import subprocess

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Import các hàm chuẩn từ form_de_kiem_tra
sys.path.append(r"c:\Antigravity_Thanh\scripts")
from form_de_kiem_tra import (
    COLOR_HF_RED, COLOR_HF_GREEN_HEX, COLOR_PRIMARY, COLOR_CORRECT, COLOR_SOLUTION,
    set_run_style, add_p_border_bottom, add_p_border_top, add_page_number_field,
    render_rich_text, setup_header_footer, build_title_block, add_section_header,
    render_question_prompt, render_options_no_table, render_statement,
    add_solution_no_table, add_student_writing_lines
)

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham"
SCRATCH_DIR = r"c:\Antigravity_Thanh\scratch"
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

# ==============================================================================
# DỮ LIỆU CÂU HỎI
# ==============================================================================

PART1_QUESTIONS = [
    {
        "num": 1,
        "text": "Câu 1. Phân tử nào sau đây được dùng làm khuôn cho quá trình dịch mã?",
        "options": ["A. rRNA.", "*B. mRNA.*", "C. tRNA.", "D. Gene."],
        "ans": "B",
        "solution": "mRNA (RNA thông tin) đóng vai trò là mạch khuôn trực tiếp chứa các codon để tổng hợp chuỗi polypeptide."
    },
    {
        "num": 2,
        "text": "Câu 2. Khi nói về quá trình dịch mã, kết luận nào sau đây không đúng?",
        "options": [
            "*A. Bộ ba kết thúc quy định tổng hợp amino acid cuối cùng trên chuỗi polypeptide.*",
            "B. Trình tự các bộ ba trên mRNA quy định trình tự amino acid trong chuỗi polypeptide.",
            "C. Liên kết hydrogen được hình thành trước liên kết peptide.",
            "D. Chiều chuyển dịch của ribosome trên mRNA là 5’ → 3’."
        ],
        "ans": "A",
        "solution": "Bộ ba kết thúc (stop codon: UAA, UAG, UGA) không mã hóa amino acid nào cả, mà chỉ đóng vai trò báo hiệu kết thúc quá trình dịch mã. Do đó nó không quy định amino acid cuối cùng."
    },
    {
        "num": 3,
        "text": "Câu 3. Đối mã đặc hiệu trên phân tử tRNA được gọi là",
        "options": ["A. codon.", "B. amino acid.", "*C. anticodon.*", "D. triplet."],
        "ans": "C",
        "solution": "Anticodon (bộ ba đối mã) là bộ ba nucleotide nằm trên phân tử tRNA, có khả năng nhận diện và khớp bổ sung với codon tương ứng trên mRNA."
    },
    {
        "num": 4,
        "text": "Câu 4. Dịch mã thông tin di truyền trên bản mã sao thành trình tự amino acid trong chuỗi polypeptide là chức năng của",
        "options": ["A. rRNA.", "B. mRNA.", "*C. tRNA.*", "D. RNA."],
        "ans": "C",
        "solution": "tRNA (RNA vận chuyển) mang amino acid tương ứng và nhận biết codon trên mRNA thông qua anticodon để lắp ráp đúng trình tự amino acid vào chuỗi polypeptide."
    },
    {
        "num": 5,
        "text": "Câu 5. Các chuỗi polypeptide được tổng hợp trong tế bào nhân thực đều",
        "options": [
            "A. kết thúc bằng Met.",
            "*B. bắt đầu bằng amino acid Met.*",
            "C. bắt đầu bằng formyl-Met.",
            "D. bắt đầu từ một phức hợp aa-tRNA."
        ],
        "ans": "B",
        "solution": "Ở sinh vật nhân thực, quá trình dịch mã bắt đầu với amino acid methionine (Met). (Khác với sinh vật nhân sơ bắt đầu bằng formyl-methionine, fMet)."
    },
    {
        "num": 6,
        "text": "Câu 6. Dịch mã là",
        "options": [
            "A. quá trình tổng hợp RNA dựa trên trình tự nucleotide của DNA.",
            "*B. quá trình tổng hợp chuỗi polypeptide dựa trên trình tự nucleotide trên bản phiên mã của gene (mRNA).*",
            "C. quá trình tạo 2 phân tử DNA mới giống như DNA ban đầu.",
            "D. quá trình tổng hợp mRNA dựa trên trình tự nucleotide của chuỗi polypeptide."
        ],
        "ans": "B",
        "solution": "Dịch mã (translation) là giai đoạn sử dụng thông tin di truyền mã hóa trên phân tử mRNA làm khuôn để sinh tổng hợp chuỗi polypeptide."
    },
    {
        "num": 7,
        "text": "Câu 7. Mã di truyền là",
        "options": [
            "A. Trình tự nucleotide trên gene (DNA) quy định thành phần và trình tự amino acid trên phân tử protein và mRNA.",
            "B. Trình tự nucleotide trên RNA quy định thành phần và trình tự amino acid trên phân tử protein, qua phân tử protein và mRNA.",
            "C. Trình tự nucleotide trên chuỗi polypeptide quy định thành phần và trình tự amino acid trên phân tử protein và mRNA.",
            "*D. Trình tự nucleotide trên phân tử mRNA quy định thành phần và trình tự amino acid trên phân tử protein.*"
        ],
        "ans": "D",
        "solution": "Mã di truyền gồm các bộ ba nucleotide (codon) trên phân tử mRNA, mỗi bộ ba quy định một amino acid (hoặc mang tín hiệu kết thúc) trong chuỗi polypeptide của protein."
    },
    {
        "num": 8,
        "text": "Câu 8. Đâu là sơ đồ thể hiện mối quan hệ giữa gene và tính trạng?",
        "options": [
            "*A. Gene → mRNA → protein → tính trạng.*",
            "B. Protein → mRNA → gene → tính trạng.",
            "C. Tính trạng → protein → mRNA → gene.",
            "D. Gene → protein → mRNA → tính trạng."
        ],
        "ans": "A",
        "solution": "Mối quan hệ biểu hiện thông tin di truyền: Trình tự nucleotide trên gene (DNA) làm khuôn phiên mã thành mRNA; mRNA dịch mã thành protein; protein trực tiếp quy định và biểu hiện thành tính trạng."
    },
    {
        "num": 9,
        "text": "Câu 9. Trong quá trình dịch mã, bộ ba mã sao 3’AUC5’ của mRNA khớp bổ sung với bộ ba đối mã nào sau đây?",
        "options": ["*A. 5’UAG3’.*", "B. 3’AUG5’.", "C. 3’UAG5’.", "D. 3’UAC5’."],
        "ans": "A",
        "solution": "Nguyên tắc bổ sung giữa codon trên mRNA và anticodon trên tRNA (ngược chiều, A - U, U - A, C - G):\nCodon mRNA: 3’–AUC–5’\nAnticodon tRNA: 5’–UAG–3’."
    },
    {
        "num": 10,
        "text": "Câu 10. Từ 4 loại nucleotide, có thể tạo ra bao nhiêu loại codon?",
        "options": ["A. 4.", "B. 8.", "C. 32.", "*D. 64.*"],
        "ans": "D",
        "solution": "Mã di truyền là mã bộ ba (mỗi codon gồm 3 nucleotide). Với 4 loại base (A, U, G, C), số loại codon có thể tạo ra là: 4^3 = 64 loại codon khác nhau."
    },
    {
        "num": 11,
        "text": "Câu 11. Quá trình dịch mã dừng lại khi",
        "options": [
            "A. Ribosome tiếp xúc với vùng kết thúc nằm ở đầu 5’ của mạch mã sao.",
            "B. Ribosome tiếp xúc với vùng kết thúc nằm ở đầu 3’ của mạch mã gốc.",
            "*C. Ribosome tiếp xúc với bộ ba kết thúc trên mRNA.*",
            "D. Ribosome tiếp xúc với vùng kết thúc nằm ở đầu của mạch mã gốc."
        ],
        "ans": "C",
        "solution": "Dịch mã dừng lại khi ribosome dịch chuyển tiếp xúc với một trong 3 bộ ba kết thúc (UAA, UAG, UGA) trên phân tử mRNA."
    },
    {
        "num": 12,
        "text": "Câu 12. Ngoài chức năng vận chuyển amino acid, RNA vận chuyển còn có chức năng quan trọng là",
        "options": [
            "A. nhân tố trung gian vận chuyển thông tin di truyền từ nhân ra tế bào chất.",
            "B. truyền thông tin di truyền qua các thế hệ cơ thể và thế hệ tế bào.",
            "C. cấu tạo nên ribosome là nơi xảy ra quá trình sinh tổng hợp protein.",
            "*D. nhận ra bộ ba mã sao tương ứng trên RNA thông tin theo nguyên tắc bổ sung.*"
        ],
        "ans": "D",
        "solution": "Phân tử tRNA có một đầu mang bộ ba đối mã (anticodon) giúp nhận biết và bắt cặp bổ sung với codon tương ứng trên mRNA để đặt đúng amino acid vào vị trí."
    },
    {
        "num": 13,
        "text": "Câu 13. Một phân tử mRNA chỉ chứa 3 loại nucleotide là A, U, C. Nhóm các bộ ba nào dưới đây có thể có trên mạch bổ sung của gene đã phiên mã ra mRNA nói trên?",
        "options": [
            "A. TAG, GAA, ATA, ATG.",
            "*B. AAA, CCA, TAA, TCC.*",
            "C. AAG, GTT, TCC, CAA.",
            "D. ATC, TAG, GAC, GAA."
        ],
        "ans": "B",
        "solution": "Mạch bổ sung của gene có trình tự giống hệt mRNA (chỉ thay U bằng T). Do phân tử mRNA chỉ chứa 3 loại A, U, C nên mạch bổ sung của gene chỉ chứa 3 loại A, T, C (hoàn toàn không có nucleotide loại G). Nhóm bộ ba thỏa mãn là: AAA, CCA, TAA, TCC."
    },
    {
        "num": 14,
        "text": "Câu 14. Cho biết các codon mã hóa các amino acid tương ứng: GGG – Gly; CCC – Pro; GCU – Ala; CGA – Arg; UCG – Ser; AGC – Ser. Một mạch gốc của một gene ở vi khuẩn có trình tự các nucleotide là 5’AGCCGACCCGGG3’. Nếu đoạn mạch gốc này mang thông tin mã hóa cho đoạn polypeptide có 4 amino acid thì trình tự của 4 amino acid đó là",
        "options": [
            "A. Ser-Arg-Pro-Gly.",
            "B. Ser-Ala-Gly-Pro.",
            "*C. Pro-Gly-Ser-Ala.*",
            "D. Gly-Pro-Ser-Arg."
        ],
        "ans": "C",
        "solution": "Mạch gốc theo đề bài: 5’–AGCCGACCCGGG–3’\nViết lại theo đúng chiều khuôn mẫu 3’ → 5’: 3’–GGGCCAGCCGA–5’\nPhiên mã ra mRNA (5’ → 3’): 5’–CCC–GGG–UCG–GCU–3’\nTra mã di truyền:\n- 5’CCC3’ → Pro\n- 5’GGG3’ → Gly\n- 5’UCG3’ → Ser\n- 5’GCU3’ → Ala\nVậy trình tự 4 amino acid là: Pro – Gly – Ser – Ala."
    },
    {
        "num": 15,
        "text": "Câu 15. Khẳng định nào dưới đây không đúng khi nói về mã di truyền?",
        "options": [
            "A. Mã di truyền là mã bộ ba.",
            "B. Có nhiều mã di truyền cùng mang tín hiệu kết thúc dịch mã.",
            "*C. Mỗi loại amino acid chỉ được xác định bởi một mã di truyền.*",
            "D. Một amino acid có thể được mã hóa bởi nhiều bộ ba."
        ],
        "ans": "C",
        "solution": "Mã di truyền có tính thoái hóa (nhiều bộ ba khác nhau cùng mã hóa cho một loại amino acid, trừ Met và Trp). Vì vậy, khẳng định 'Mỗi loại amino acid chỉ được xác định bởi một mã di truyền' là sai."
    },
    {
        "num": 16,
        "text": "Câu 16. Các codon nào dưới đây không mã hoá amino acid?",
        "options": [
            "*A. UAA, UAG, UGA.*",
            "B. AAU, GAU, UAC.",
            "C. AUA, UAA, UGC.",
            "D. CUG, ACC, GUA."
        ],
        "ans": "A",
        "solution": "Trong 64 codon, có 3 codon kết thúc là UAA, UAG và UGA không mã hóa cho bất kì amino acid nào, đóng vai trò báo hiệu kết thúc dịch mã."
    },
    {
        "num": 17,
        "text": "Câu 17. Các mã di truyền khác nhau ở",
        "options": [
            "A. số lượng và thành phần các nucleotide.",
            "*B. trình tự và thành phần nucleotide.*",
            "C. thành phần các nucleotide.",
            "D. số lượng và trình tự nucleotide."
        ],
        "ans": "B",
        "solution": "Tất cả các mã di truyền đều là mã bộ ba (có cùng số lượng là 3 nucleotide), các mã khác nhau ở thành phần và trình tự sắp xếp của các nucleotide."
    },
    {
        "num": 18,
        "text": "Câu 18. Thành phần nào sau đây không tham gia quá trình dịch mã?",
        "options": ["A. RNA.", "B. Ribosome.", "*C. DNA.*", "D. Amino acid."],
        "ans": "C",
        "solution": "Dịch mã diễn ra ở tế bào chất trên phân tử mRNA. DNA nằm trong nhân tế bào (hoặc vùng nhân), không tham gia trực tiếp vào quá trình dịch mã."
    },
    {
        "num": 19,
        "text": "Câu 19. Trong các bộ ba, có bao nhiêu bộ ba mã hóa cho các amino acid?",
        "options": ["A. 64.", "B. 63.", "C. 62.", "*D. 61.*"],
        "ans": "D",
        "solution": "Trong 64 mã bộ ba, có 3 bộ ba mang tín hiệu kết thúc (UAA, UAG, UGA) không mã hóa amino acid; còn lại 64 - 3 = 61 bộ ba mã hóa cho 20 loại amino acid."
    },
    {
        "num": 20,
        "text": "Câu 20. Trong các bộ ba sau, bộ ba nào mã hóa cho amino acid Methionine?",
        "options": ["A. UAA.", "*B. AUG.*", "C. GGU.", "D. CAA."],
        "ans": "B",
        "solution": "Bộ ba mã hóa cho amino acid Methionine là 5’AUG3’. Ngoài ra, AUG còn đóng vai trò là mã mở đầu cho quá trình dịch mã."
    },
    {
        "num": 21,
        "text": "Câu 21. Cho các phát biểu sau:\n(1) Mã di truyền là thông tin về trình tự các amino acid được mã hóa dưới dạng trình tự các nucleotide trên mRNA.\n(2) Có 64 bộ ba được hình thành từ bốn loại nucleotide.\n(3) Số lượng, thành phần và trình tự sắp xếp các mã di truyền trên mRNA quy định số lượng, thành phần và trình tự sắp xếp các amino acid trên phân tử protein.\n(4) Ba bộ ba UAA, UGG và UCC đóng vai trò kết thúc dịch mã.\nTrong các phát biểu trên, có bao nhiêu phát biểu đúng?",
        "options": ["A. 1.", "*B. 2.*", "C. 3.", "D. 4."],
        "ans": "B",
        "solution": "Các phát biểu đúng là: (1), (2).\n- (1) Đúng: Định nghĩa mã di truyền.\n- (2) Đúng: Có 4^3 = 64 bộ ba.\n- (3) Sai: mRNA quy định trình tự trên chuỗi polypeptide, sau đó polypeptide được hoàn thiện cấu trúc để tạo nên phân tử protein.\n- (4) Sai: Ba bộ ba kết thúc là UAA, UAG và UGA."
    },
    {
        "num": 22,
        "text": "Câu 22. Dịch mã là",
        "options": [
            "*A. quá trình tổng hợp chuỗi polypeptide thông qua thông tin di truyền được mã hóa trên phân tử mRNA.*",
            "B. quá trình tổng hợp chuỗi polynucleotide thông qua thông tin di truyền được mã hóa trên phân tử mRNA.",
            "C. quá trình tổng hợp chuỗi polypeptide thông qua thông tin di truyền được mã hóa trên phân tử DNA.",
            "D. quá trình tổng hợp chuỗi polynucleotide thông qua thông tin di truyền được mã hóa trên phân tử DNA."
        ],
        "ans": "A",
        "solution": "Dịch mã là quá trình tổng hợp chuỗi polypeptide dựa trên thông tin di truyền được mã hóa trên phân tử mRNA."
    },
    {
        "num": 23,
        "text": "Câu 23. Trong các thành phần sau, có bao nhiêu thành phần trực tiếp tham gia vào quá trình tổng hợp chuỗi polypeptide?\n(1) Gene; (2) mRNA; (3) Amino acid; (4) tRNA; (5) Ribosome; (6) Enzyme; (7) rRNA; (8) RNA mồi.",
        "options": ["*A. 6.*", "B. 2.", "C. 3.", "D. 4."],
        "ans": "A",
        "solution": "Có 6 thành phần trực tiếp tham gia: (2) mRNA, (3) Amino acid, (4) tRNA, (5) Ribosome, (6) Enzyme, (7) rRNA.\n(Gene chỉ tham gia phiên mã, RNA mồi tham gia tái bản DNA)."
    },
    {
        "num": 24,
        "text": "Câu 24. Cho các bộ ba mã hóa cho các amino acid tương ứng như sau: GGG – Gly; CCC – Pro; GCU – Ala; CGA – Arg; UCG – Ser; AGC – Ser. Đoạn mạch bổ sung của một gene ở vi khuẩn có trình tự các nucleotide: 3’-TCGGCTGGGGCCC-5’. Nếu đoạn mạch khuôn của gene này mang thông tin mã hóa cho đoạn polypeptide có 4 amino acid thì trình tự của 4 amino acid đó là",
        "options": [
            "A. Ser - Ala - Gly - Pro.",
            "B. Pro - Gly - Ser - Ala.",
            "C. Ser - Arg - Pro - Gly.",
            "*D. Gly - Pro - Ser - Arg.*"
        ],
        "ans": "D",
        "solution": "Mạch bổ sung: 3’–TCG GCT GGG CCC–5’ ↔ 5’–CCC GGG TCG GCT–3’\nMạch khuôn (3’ → 5’): 3’–GGG CCC AGC CGA–5’\nmRNA (5’ → 3’): 5’–CCC GGG UCG GCU–3’ → Đọc mạch khuôn 3'->5':\nCodon 1: 5'GGG3' (từ 3'CCC5') → Gly\nCodon 2: 5'CCC3' (từ 3'GGG5') → Pro\nCodon 3: 5'AGC3' (từ 3'UCG5') → Ser\nCodon 4: 5'CGA3' (từ 3'GCU5') → Arg\nChuỗi polypeptide: Gly – Pro – Ser – Arg."
    },
    {
        "num": 25,
        "text": "Câu 25. Quá trình dịch mã dừng lại khi",
        "options": [
            "A. ribosome tiếp xúc với vùng kết thúc nằm ở đầu 5’ của mạch mã sao.",
            "B. ribosome tiếp xúc với vùng kết thúc nằm ở đầu 3’ của mạch mã gốc.",
            "*C. ribosome tiếp xúc với bộ ba kết thúc trên mRNA.*",
            "D. ribosome tiếp xúc với vùng kết thúc nằm ở đầu 5’ của mạch mã gốc."
        ],
        "ans": "C",
        "solution": "Quá trình dịch mã dừng lại khi ribosome dịch chuyển tiếp xúc với một trong 3 bộ ba kết thúc (UAA, UAG, UGA) trên phân tử mRNA."
    },
    {
        "num": 26,
        "text": "Câu 26. Chọn từ/cụm từ thích hợp điền vào chỗ trống trong câu sau: Trình tự các nucleotide trên phân tử mRNA được ......... thành trình tự các amino acid trên phân tử protein.",
        "options": ["*A. dịch mã.*", "B. tái bản.", "C. phiên mã.", "D. biểu hiện."],
        "ans": "A",
        "solution": "Quá trình chuyển đổi trình tự ribonucleotide trên mRNA thành trình tự amino acid trên protein được gọi là dịch mã."
    },
    {
        "num": 27,
        "text": "Câu 27. Các gene khác nhau có thể quy định các tính trạng khác nhau là do",
        "options": [
            "A. các gene có kích thước khác nhau.",
            "B. các gene có thể bị đột biến.",
            "C. các gene có vị trí khác nhau trong các tế bào khác nhau.",
            "*D. các gene có trình tự các nucleotide khác nhau.*"
        ],
        "ans": "D",
        "solution": "Trình tự nucleotide trong gene quy định trình tự codon trên mRNA, từ đó quy định trình tự amino acid trong chuỗi polypeptide tạo nên protein biểu hiện thành tính trạng. Do đó các gene khác nhau có trình tự nucleotide khác nhau."
    },
    {
        "num": 28,
        "text": "Câu 28. Khi nói về dịch mã ở sinh vật nhân thực, nhận định nào sau đây là không đúng?",
        "options": [
            "A. Trong cùng một thời điểm có thể có nhiều ribosome tham gia dịch mã trên một phân tử mRNA.",
            "B. Amino acid mở đầu trong quá trình dịch mã là metionine.",
            "C. Khi dịch mã, ribosome dịch chuyển theo chiều 5’ → 3’ trên phân tử mRNA.",
            "*D. Khi dịch mã, ribosome dịch chuyển theo chiều 3’ → 5’ trên phân tử mRNA.*"
        ],
        "ans": "D",
        "solution": "Khi dịch mã, ribosome luôn dịch chuyển theo chiều 5’ → 3’ trên phân tử mRNA. Khẳng định ribosome dịch chuyển theo chiều 3’ → 5’ là sai."
    },
    {
        "num": 29,
        "text": "Câu 29. Khi nói về cơ chế dịch mã, có bao nhiêu nhận định không đúng trong các nhận định sau?\n(1) Trên một phân tử mRNA hoạt động của polysome giúp tạo ra các chuỗi polypeptide loại khác nhau.\n(2) Khi dịch mã, ribosome dịch chuyển theo chiều 3’ – 5’ trên phân tử mRNA.\n(3) Với bộ ba UAG trên mRNA thì tRNA mang đối mã là AUC.\n(4) Các chuỗi polypeptide sau dịch mã được cắt bỏ amino acid mở đầu và tiếp tục hình thành các bậc cấu trúc cao hơn để trở thành protein có hoạt tính sinh học.\n(5) Sau khi hoàn tất quá trình dịch mã, ribosome tách khỏi mRNA và giữ nguyên cấu trúc để chuẩn bị cho quá trình dịch mã tiếp theo.\nSố nhận định không đúng là",
        "options": ["A. 2.", "*B. 4.*", "C. 1.", "D. 3."],
        "ans": "B",
        "solution": "Có 4 nhận định không đúng là: (1), (2), (3), (5).\n- (1) Sai: Polysome trượt trên cùng 1 mRNA nên tạo các chuỗi polypeptide cùng loại.\n- (2) Sai: Ribosome trượt theo chiều 5’ → 3’.\n- (3) Sai: UAG là bộ ba kết thúc, không có tRNA mang anticodon tương ứng.\n- (4) Đúng.\n- (5) Sai: Khi kết thúc dịch mã, ribosome tách thành 2 tiểu phần tự do, không giữ nguyên cấu trúc gắn kết."
    },
    {
        "num": 30,
        "text": "Câu 30. Phân tử nào sau đây mang bộ ba đối mã?",
        "options": ["A. mRNA.", "B. DNA.", "*C. tRNA.*", "D. rRNA."],
        "ans": "C",
        "solution": "Phân tử tRNA mang bộ ba đối mã (anticodon) đặc hiệu để nhận biết codon trên mRNA."
    },
    {
        "num": 31,
        "text": "Câu 31. Ở tế bào nhân thực, quá trình dịch mã được bắt đầu bằng bộ ba nào trên mRNA và amino acid nào mở đầu chuỗi polypeptide?",
        "options": [
            "A. AGU và formyl-Met.",
            "B. AUG và formyl-Met.",
            "*C. AUG và amino acid Met.*",
            "D. AGU và amino acid Met."
        ],
        "ans": "C",
        "solution": "Ở sinh vật nhân thực, bộ ba mở đầu là 5’AUG3’ và mã hóa cho amino acid methionine (Met)."
    },
    {
        "num": 32,
        "text": "Câu 32. Trong giờ thực hành, giáo viên trình chiếu sơ đồ: DNA → (phiên mã) → mRNA → (dịch mã) → chuỗi polypeptide → protein → tính trạng. Ý nào sau đây mô tả chính xác vai trò của gene trong sơ đồ trên?",
        "options": [
            "A. Gene trực tiếp tạo nên tính trạng mà không cần trung gian.",
            "*B. Gene quy định trình tự nucleotide trên mRNA; mRNA quy định trình tự amino acid trong protein; protein góp phần tạo nên tính trạng.*",
            "C. Gene chỉ là vật chất lưu trữ, không tham gia quyết định tính trạng.",
            "D. Gene chỉ tác động lên tính trạng khi môi trường thay đổi."
        ],
        "ans": "B",
        "solution": "Sơ đồ thể hiện luồng truyền đạt thông tin di truyền: Gene quy định trình tự nucleotide trên mRNA; mRNA quy định trình tự amino acid trong protein; protein trực tiếp tham gia tạo nên tính trạng."
    },
    {
        "num": 33,
        "text": "Câu 33. Quá trình dịch mã kết thúc khi",
        "options": [
            "A. ribosome tiếp xúc với bộ ba AUG trên mRNA.",
            "B. ribosome rời khỏi mRNA và trở về trạng thái tự do.",
            "*C. ribosome tiếp xúc với một trong các bộ ba: UAA, UAG, UGA.*",
            "D. ribosome gắn methionine vào vị trí cuối cùng của chuỗi polypeptide."
        ],
        "ans": "C",
        "solution": "Quá trình dịch mã kết thúc khi ribosome dịch chuyển gặp một trong ba bộ ba kết thúc (UAA, UAG, UGA) trên phân tử mRNA."
    }
]

PART2_QUESTIONS = [
    {
        "num": 1,
        "text": "Câu 1. Khi nói về đặc điểm của mã di truyền:",
        "items": [
            ("a) Mã di truyền là thông tin về trình tự các axit amin được mã hoá dưới dạng trình tự các nuclêôtit trên mARN.", "Đúng"),
            ("b) Có 61 bộ ba mã hoá cho các axit amin và 3 bộ ba UAA, UAG, UGA đóng vai trò là mã kết thúc dịch mã.", "Đúng"),
            ("c) Bộ ba AUG chỉ mã hoá axit amin methionin và không đóng vai trò là mã mở đầu.", "Sai"),
            ("d) Số lượng, thành phần và trật tự sắp xếp các bộ ba trên mARN quy định số lượng, thành phần và trật tự sắp xếp các amino acid trong chuỗi polypeptide.", "Đúng")
        ],
        "solution": "a) Đúng: Định nghĩa mã di truyền.\nb) Đúng: Trong 64 codon có 61 codon mã hóa cho 20 loại amino acid và 3 codon kết thúc.\nc) Sai: AUG là codon mở đầu quy định điểm khởi đầu dịch mã và mã hóa Met ở nhân thực.\nd) Đúng: Trình tự các bộ ba trên mRNA quyết định trực tiếp số lượng, thành phần, trật tự sắp xếp các amino acid trong chuỗi polypeptide."
    },
    {
        "num": 2,
        "text": "Câu 2. Khi xét về cơ sở phân tử và di truyền học:",
        "items": [
            ("a) Axit nucleic là hợp chất hữu cơ có vai trò lưu giữ và truyền đạt thông tin di truyền.", "Đúng"),
            ("b) Nhiều mã di truyền cùng mã hoá một loại amino acid.", "Đúng"),
            ("c) Kiểu gen AaBb là kết quả của phép lai sau: P: AaBB × AaBB.", "Sai"),
            ("d) Dòng thuần là tập hợp các cơ thể có kiểu hình giống nhau qua nhiều thế hệ tự thụ phấn hoặc giao phối cận huyết.", "Sai")
        ],
        "solution": "a) Đúng: Nucleic acid (DNA và RNA) lưu trữ, bảo quản và truyền đạt thông tin di truyền.\nb) Đúng: Đây là tính thoái hóa của mã di truyền.\nc) Sai: Phép lai P: AaBB × AaBB không có alen b nên không thể sinh ra con có kiểu gen mang alen b (AaBb).\nd) Sai: Dòng thuần chủng (pure-breeding line) phải có kiểu gen đồng hợp tử về tất cả các cặp gen, không chỉ đơn thuần là có cùng kiểu hình."
    },
    {
        "num": 3,
        "text": "Câu 3. Một đoạn phân tử DNA ở sinh vật nhân thực có trình tự nucleotide trên mạch gốc là: 3’...AAACAATGGGGA...5’.",
        "items": [
            ("a) Trình tự nucleotide trên mạch bổ sung của đoạn DNA này là 5’...AAAGTTACCGGT...3’.", "Sai"),
            ("b) Trình tự nucleotide trên mạch mRNA do gene trên phiên mã là 5’...UUUGUUACCCCU...3’.", "Đúng"),
            ("c) Chiều tổng hợp mạch mRNA theo chiều 3’......5’.", "Sai"),
            ("d) Có 5 codon được tạo ra từ đoạn gene trên.", "Sai")
        ],
        "solution": "a) Sai: Mạch bổ sung theo NTBS (A - T, G - C) phải là: 5’...TTTGTTACCCCT...3’.\nb) Đúng: mRNA phiên mã theo NTBS (A - U, T - A, G - C, C - G) theo chiều 5’ → 3’ là 5’...UUUGUUACCCCU...3’.\nc) Sai: Mạch RNA luôn được tổng hợp theo chiều 5’ → 3’.\nd) Sai: Đoạn gene gồm 12 nucleotide → tổng hợp mRNA có 12 ribonucleotide → tạo ra 12 : 3 = 4 codon (không phải 5 codon)."
    },
    {
        "num": 4,
        "text": "Câu 4. Trong các vụ án hình sự hoặc xác định quan hệ huyết thống, các nhà khoa học thường sử dụng kỹ thuật \"vân tay DNA\". Chỉ cần một mẫu nhỏ như sợi tóc, giọt máu hoặc tế bào biểu mô, người ta có thể giải mã trình tự các nucleotide để xác định danh tính cá nhân với độ chính xác gần như tuyệt đối. Điều này dựa trên tính đặc trưng của hệ gene và các cơ sở phân tử của sự di truyền.\nDựa vào thông tin trên và kiến thức đã học, hãy xác định các nhận định sau là đúng hay sai?",
        "items": [
            ("a) Kỹ thuật xác định huyết thống dựa trên nguyên tắc mỗi cá nhân có một hệ gene đặc trưng, trừ trường hợp sinh đôi cùng trứng.", "Đúng"),
            ("b) Mã di truyền là trình tự sắp xếp các amino acid trong phân tử protein, quy định các đặc điểm tính trạng của cơ thể.", "Sai"),
            ("c) Chỉ từ 4 loại nucleotide (A, T, G, C), sự thay đổi trình tự sắp xếp và số lượng của chúng trên DNA đã tạo ra sự đa dạng vô tận của mã di truyền giữa các loài.", "Đúng"),
            ("d) Quá trình tái bản DNA diễn ra theo nguyên tắc bổ sung, giúp thông tin di truyền được truyền đạt nguyên vẹn từ tế bào mẹ sang các tế bào con.", "Đúng")
        ],
        "solution": "a) Đúng: Hệ gene của mỗi cá nhân là duy nhất và đặc thù, trừ trường hợp sinh đôi cùng trứng.\nb) Sai: Mã di truyền là trình tự sắp xếp nucleotide trên nucleic acid (mRNA/DNA), không phải trình tự amino acid trên protein.\nc) Đúng: Tính đa dạng và đặc thù của DNA được quy định bởi số lượng, thành phần và trật tự sắp xếp của 4 loại nucleotide.\nd) Đúng: Tái bản bán bảo tồn và bổ sung giúp thông tin di truyền truyền đạt chính xác qua các thế hệ tế bào."
    }
]

PART3_QUESTIONS = [
    {
        "num": 1,
        "text": "Câu 1. Trên mạch mang mã gốc của gene có một bộ ba 3’AGC5’. Bộ ba tương ứng trên phân tử mRNA được phiên mã từ gene này là gì?",
        "ans": "5’UCG3’",
        "solution": "Theo NTBS trong phiên mã: A liên kết với U, G liên kết với C, C liên kết với G và ngược chiều. Mạch khuôn 3’AGC5’ phiên mã thành mRNA là 5’UCG3’."
    },
    {
        "num": 2,
        "text": "Câu 2. Phân tử nào trực tiếp làm khuôn cho quá trình dịch mã?",
        "ans": "mRNA",
        "solution": "Phân tử mRNA mang thông tin di truyền trực tiếp làm mạch khuôn cho ribosome trượt qua dịch mã."
    },
    {
        "num": 3,
        "text": "Câu 3. Ở sinh vật nhân thực, codon 5’AUG3’ mã hóa loại amino acid nào?",
        "ans": "Methionine",
        "solution": "Ở sinh vật nhân thực, codon mở đầu 5’AUG3’ mã hóa amino acid Methionine (Met)."
    },
    {
        "num": 4,
        "text": "Câu 4. Mạch khuôn của gene có đoạn 3’ TATGGGCATGTA 5’ thì mRNA được phiên mã từ mạch khuôn này có trình tự nucleotide như thế nào?",
        "ans": "5’AUACCCGUACAU3’",
        "solution": "Phiên mã theo NTBS (T - A, A - U, G - C, C - G): Mạch gốc 3’TATGGGCATGTA5’ phiên mã thành 5’AUACCCGUACAU3’."
    },
    {
        "num": 5,
        "text": "Câu 5. Giả sử có 3 loại nucleotide A, U, G thì phân tử mRNA có tối đa bao nhiêu loại mã di truyền mã hóa amino acid?",
        "ans": "24",
        "solution": "Tổng số codon tạo được từ 3 loại A, U, G là: 3^3 = 27 codon. Trong đó có 3 codon kết thúc tạo từ A, U, G là: UAA, UAG, UGA. Số codon mã hóa amino acid là: 27 - 3 = 24."
    },
    {
        "num": 6,
        "text": "Câu 6. Một phân tử mRNA có 1200 nucleotide nhưng tính từ bộ ba mở đầu đến bộ ba kết thúc thì có tổng số 303 nucleotide. Phân tử mRNA này dịch mã có 6 ribosome trượt qua 1 lần đã cần môi trường cung cấp bao nhiêu amino acid?",
        "ans": "600",
        "solution": "Số bộ ba trong vùng mã hóa = 303 : 3 = 101 bộ ba. Trừ 1 bộ ba kết thúc không mã hóa, mỗi chuỗi polypeptide cần môi trường cung cấp: 101 - 1 = 100 amino acid. Với 6 ribosome trượt qua 1 lần, số amino acid cần cung cấp là: 6 × 100 = 600 amino acid."
    },
    {
        "num": 7,
        "text": "Câu 7. Một phân tử mRNA có 1200 nucleotide nhưng tính từ bộ ba mở đầu đến bộ ba kết thúc thì có tổng số 606 nucleotide. Phân tử mRNA này dịch mã tạo ra chuỗi polypeptide. Tính số amino acid của mỗi chuỗi polypeptide hoàn chỉnh.",
        "ans": "200",
        "solution": "Số bộ ba trong vùng mã hóa = 606 : 3 = 202 bộ ba. Chuỗi polypeptide sơ khai có: 202 - 1 = 201 amino acid. Chuỗi polypeptide hoàn chỉnh sau khi cắt bỏ amino acid mở đầu có: 201 - 1 = 200 amino acid."
    },
    {
        "num": 8,
        "text": "Câu 8. Quá trình dịch mã diễn ra ở đâu?",
        "ans": "Tế bào chất",
        "solution": "Quá trình dịch mã sinh tổng hợp protein diễn ra tại tế bào chất (trên ribosome)."
    },
    {
        "num": 9,
        "text": "Câu 9. tRNA là gì?",
        "ans": "RNA vận chuyển",
        "solution": "tRNA là RNA vận chuyển (transfer RNA), có nhiệm vụ vận chuyển amino acid đặc hiệu đến ribosome và nhận diện codon tương ứng trên mRNA."
    },
    {
        "num": 10,
        "text": "Câu 10. Phân tử mRNA ở tế bào nhân sơ được phiên mã từ một gene có 3000 nucleotide sau đó tham gia vào quá trình dịch mã. Quá trình tổng hợp protein có 5 ribosome cùng trượt trên mRNA đó. Số amino acid môi trường cần cung cấp để hoàn tất quá trình dịch mã trên là bao nhiêu?",
        "ans": "2495",
        "solution": "Số ribonucleotide của mRNA: rN = N / 2 = 3000 / 2 = 1500 nucleotide. Số bộ ba = 1500 : 3 = 500 bộ ba. Mỗi ribosome cần cung cấp: 500 - 1 = 499 amino acid. Tổng số amino acid 5 ribosome cần là: 5 × 499 = 2495 amino acid."
    },
    {
        "num": 11,
        "text": "Câu 11. Các bộ ba trên mRNA mã hóa các amino acid tương ứng sau: 5'UUG3' – Leu, 5'AUC3' – Ile, 5'ACC3' – Thr, 5'AAU3' – Asn, 5'AUG3' – Met (amino acid mở đầu), 5'GUA3' – Val, 5'UGC3' – Cys, 5'UAG 3' – Kết thúc, 5'GGC3' – Gly. Một phân tử mRNA ngắn của vi khuẩn có trình tự các nucleotide như sau: 5'AUG AUC UUG ACC AUC AAU UGC GGC UGC UGA3'. Xác định trình tự nucleotide trên mạch bổ sung của gene.",
        "ans": "5’ATG ATC TTG ACC ATC AAT TGC GGC TGC TGA3’",
        "solution": "Mạch bổ sung của gene có trình tự nucleotide giống hệt mRNA theo chiều 5’ → 3’, chỉ thay thế ribonucleotide U bằng nucleotide T. Vậy mạch bổ sung là: 5’ATG ATC TTG ACC ATC AAT TGC GGC TGC TGA3’."
    },
    {
        "num": 12,
        "text": "Câu 12. Cho biết các codon mã hóa các amino acid tương ứng như sau: 5'GGG3' – Gly; 5'CCC3' – Pro; 5'GCU3' – Ala; 5'CGA3' – Arg; 5'UCG3' – Ser; 5'AGC3' – Ser; 5'UAC3' – Tyr. Một đoạn mạch gốc của một gene ở vi khuẩn mang thông tin mã hóa cho đoạn polypeptide có 5 amino acid có trình tự các nucleotide là 3'CCC-AGC-ATG-CGA-GGG5'. Xác định trình tự của 5 amino acid do đoạn gene này quy định tổng hợp.",
        "ans": "Gly - Ser - Tyr - Ala - Pro",
        "solution": "Mạch gốc (3’ → 5’): 3’–CCC–AGC–ATG–CGA–GGG–5’\n→ mRNA (5’ → 3’): 5’–GGG–UCG–UAC–GCU–CCC–3’\nTra bảng mã:\n- 5’GGG3’ → Gly\n- 5’UCG3’ → Ser\n- 5’UAC3’ → Tyr\n- 5’GCU3’ → Ala\n- 5’CCC3’ → Pro\nVậy trình tự 5 amino acid là: Gly – Ser – Tyr – Ala – Pro."
    },
    {
        "num": 13,
        "text": "Câu 13. Một bộ ba ribonucleotide liền kề trên mRNA quy định một amino acid (hoặc mang tín hiệu kết thúc) được gọi là gì?",
        "ans": "Codon",
        "solution": "Bộ ba ribonucleotide liền kề trên phân tử mRNA được gọi là codon (hoặc bộ ba mã sao)."
    },
    {
        "num": 14,
        "text": "Câu 14. Từ 4 loại ribonucleotide (A, U, G, C), có thể hình thành tất cả bao nhiêu loại codon?",
        "ans": "64",
        "solution": "Số loại codon tối đa hình thành từ 4 loại base là: 4^3 = 64 loại codon."
    },
    {
        "num": 15,
        "text": "Câu 15. Trong tổng số 64 bộ ba mã di truyền, có bao nhiêu bộ ba mã hóa cho các amino acid?",
        "ans": "61",
        "solution": "Có 3 bộ ba mang tín hiệu kết thúc (UAA, UAG, UGA) không mã hóa amino acid, do đó có 64 - 3 = 61 bộ ba mã hóa amino acid."
    },
    {
        "num": 16,
        "text": "Câu 16. Bộ ba làm nhiệm vụ khởi đầu quá trình dịch mã trên mRNA là bộ ba nào?",
        "ans": "5’AUG3’",
        "solution": "Bộ ba làm nhiệm vụ khởi đầu dịch mã trên mRNA là 5’AUG3’."
    },
    {
        "num": 17,
        "text": "Câu 17. Khi gặp một trong ba bộ ba nào (viết theo chiều 5’ → 3’) trên phân tử mRNA thì quá trình dịch mã sẽ dừng lại?",
        "ans": "UAA, UAG, UGA",
        "solution": "Ba bộ ba kết thúc dịch mã là: 5’UAA3’, 5’UAG3’, 5’UGA3’."
    },
    {
        "num": 18,
        "text": "Câu 18. Bào quan nào trong tế bào đóng vai trò là nơi trực tiếp diễn ra quá trình dịch mã tổng hợp chuỗi polypeptide?",
        "ans": "Ribosome",
        "solution": "Bào quan trực tiếp thực hiện quá trình dịch mã là ribosome."
    },
    {
        "num": 19,
        "text": "Câu 19. Phân tử RNA nào giữ vai trò mang bộ ba đối mã (anticodon) và trực tiếp vận chuyển amino acid tương ứng đến ribosome?",
        "ans": "tRNA",
        "solution": "Phân tử tRNA mang bộ ba đối mã (anticodon) và vận chuyển amino acid đặc hiệu."
    },
    {
        "num": 20,
        "text": "Câu 20. Quá trình dịch mã tổng hợp chuỗi polypeptide được chia làm mấy giai đoạn chính?",
        "ans": "3 giai đoạn",
        "solution": "Quá trình dịch mã gồm 3 giai đoạn chính: Khởi đầu (Initiation), Kéo dài chuỗi polypeptide (Elongation), và Kết thúc (Termination)."
    }
]

PART4_QUESTIONS = [
    {
        "num": 1,
        "text": "Câu 1. Một phân tử mRNA có tổng cộng 600 nucleotide (kể cả bộ ba kết thúc). Chuỗi polypeptide hoàn chỉnh được tổng hợp có bao nhiêu amino acid?",
        "solution": "- Tổng số bộ ba (codon) trên mRNA: 600 : 3 = 200 bộ ba.\n- Số amino acid được tổng hợp trong chuỗi polypeptide sơ khai: Naa = Ncodon - 1 (trừ bộ ba kết thúc) = 200 - 1 = 199 amino acid.\n- Chuỗi polypeptide hoàn chỉnh sau khi cắt bỏ amino acid mở đầu có số amino acid là: 199 - 1 = 198 amino acid (theo tài liệu gốc: Naa = Ncodon - 1 = 199 aa)."
    },
    {
        "num": 2,
        "text": "Câu 2. Để tổng hợp một chuỗi polypeptide gồm 399 amino acid, số lượng nucleotide tối thiểu cần có trên phân tử mRNA (bao gồm codon kết thúc) là bao nhiêu nucleotide?",
        "solution": "- Số amino acid của chuỗi polypeptide sơ khai = 399.\n- Số bộ ba mã hóa (codon) cần thiết (kể cả bộ ba kết thúc): 399 + 1 = 400 bộ ba.\n- Số nucleotide tối thiểu cần có trên mRNA: N = 400 × 3 = 1200 nucleotide."
    },
    {
        "num": 3,
        "text": "Câu 3. Một đoạn phân tử mRNA có trình tự như sau:\n5’ – AUGGCGAAGCCCUGCGAGUAA – 3’\na. Đoạn phân tử mRNA trên chứa bao nhiêu mã bộ ba?\nb. Chuỗi polypeptide được dịch mã từ đoạn phân tử mRNA trên có bao nhiêu amino acid?",
        "solution": "a. Đoạn phân tử mRNA trên gồm 21 nucleotide.\nSố mã bộ ba chứa trong đoạn phân tử mRNA trên là: 21 : 3 = 7 bộ ba (AUG – GCG – AAG – CCC – UGC – GAG – UAA).\nb. Trong 7 bộ ba trên, bộ ba cuối cùng 5’UAA3’ là bộ ba kết thúc không mã hóa amino acid.\n- Chuỗi polypeptide mới sinh tổng hợp có: 7 - 1 = 6 amino acid.\n- Chuỗi polypeptide hoàn chỉnh (sau khi cắt bỏ amino acid mở đầu Met) có: 6 - 1 = 5 amino acid."
    },
    {
        "num": 4,
        "text": "Câu 4. Một nhà khoa học tổng hợp phân tử mRNA nhân tạo với vùng mã hoá protein chứa 1200 nucleotide (bao gồm cả mã mở đầu và mã kết thúc). Nhà khoa học thực hiện phản ứng dịch mã phân tử mRNA mới tổng hợp trong tế bào vi khuẩn E. coli. Hãy cho biết chuỗi polypeptide được dịch mã có bao nhiêu amino acid?",
        "solution": "- Số bộ ba trên vùng mã hóa của phân tử mRNA là: 1200 : 3 = 400 bộ ba (trong đó có 1 bộ ba mở đầu và 1 bộ ba kết thúc).\n- Chuỗi polypeptide sơ khai mới dịch mã có: 400 - 1 = 399 amino acid.\n- Chuỗi polypeptide hoàn chỉnh sau khi cắt bỏ formyl-methionine (fMet) có hoạt tính sinh học chứa: 399 - 1 = 398 amino acid."
    },
    {
        "num": 5,
        "text": "Câu 5. Cho sơ đồ sau: Gene (DNA) ——(1)——> mRNA ——(2)——> Protein ——> Tính trạng\na) Nêu tên và sản phẩm của quá trình 1, quá trình 2.\nb) Giải thích mối quan hệ giữa gene (DNA), mRNA, protein và tính trạng.",
        "solution": "a) Tên và sản phẩm của các quá trình:\n- Quá trình (1) là Phiên mã (Transcription), sản phẩm là phân tử mRNA (RNA thông tin).\n- Quá trình (2) là Dịch mã (Translation), sản phẩm là chuỗi polypeptide (Protein).\nb) Giải thích mối quan hệ:\n- Gene (DNA) là nơi lưu trữ thông tin di truyền dưới dạng trình tự sắp xếp các nucleotide.\n- Thông qua quá trình phiên mã theo NTBS, thông tin di truyền trên mạch khuôn của gene được chuyển giao thành trình tự các codon trên mRNA.\n- Tại tế bào chất, phân tử mRNA làm khuôn cho quá trình dịch mã để sắp xếp các amino acid thành chuỗi polypeptide theo đúng trình tự mã hóa.\n- Phân tử protein sau khi hình thành cấu trúc không gian đặc thù sẽ tham gia các hoạt động chức năng sinh học, từ đó trực tiếp biểu hiện thành các tính trạng của cơ thể sinh vật."
    },
    {
        "num": 6,
        "text": "Câu 6. Cho đoạn phân tử mRNA có trình tự sắp xếp các nucleotide như sau:\n5’...A–C–C–U–C–A–U–U–A–G–C–A–G–C–G...3’\nViết trình tự nucleotide của cả đoạn gene phiên mã đoạn mRNA trên?",
        "solution": "Theo nguyên tắc bổ sung trong phiên mã và cấu tạo xoắn kép DNA (A liên kết với U/T, G liên kết với C):\n- Mạch khuôn của gene (chiều 3’ → 5’):\n3’...T–G–G–A–G–T–A–A–T–C–G–T–C–G–C...5’\n- Mạch bổ sung của gene (chiều 5’ → 3’):\n5’...A–C–C–T–C–A–T–T–A–G–C–A–G–C–G...3’\nCấu trúc cả đoạn gene:\nMạch bổ sung (5’ → 3’): 5’...A–C–C–T–C–A–T–T–A–G–C–A–G–C–G...3’\nMạch khuôn   (3’ → 5’): 3’...T–G–G–A–G–T–A–A–T–C–G–T–C–G–C...5’"
    },
    {
        "num": 7,
        "text": "Câu 7. Một đoạn mạch của gene có cấu trúc như sau:\nMạch 1: – A – T – G – C – T – C – G –\nMạch 2: – T – A – C – G – A – G – C –\nXác định trình tự các đơn phân của đoạn mạch RNA được tổng hợp từ mạch 2.",
        "solution": "Mạch 2 đóng vai trò là mạch khuôn: – T – A – C – G – A – G – C –\nTheo nguyên tắc bổ sung trong phiên mã (T – A, A – U, C – G, G – C), đoạn mạch RNA được tổng hợp từ mạch 2 có trình tự đơn phân là:\n– A – U – G – C – U – C – G –."
    },
    {
        "num": 8,
        "text": "Câu 8. Cho biết các bộ ba trên mRNA mã hóa cho các amino acid tương ứng là: 5’CUG3’ - Leu, 5’GUC3’ - Val, 5’ACG3’ - Thr, 5’GCA3’ - Ala. Từ đoạn mạch gốc chứa 4 mã di truyền của một gene có trình tự đơn phân 5’CAGCGTGACCAG3’ phiên mã tổng hợp đoạn mRNA. Theo cơ chế dịch mã, hãy viết đoạn polypeptide được tạo ra từ đoạn mRNA trên.",
        "solution": "- Đoạn mạch gốc của gene theo đề bài: 5’–CAGCGTGACCAG–3’\n→ Viết lại theo đúng chiều khuôn mẫu 3’ → 5’: 3’–GACCAGTCGGAC–5’ (hay 3’–GAC–CAG–TGC–GAC–5’)\n- Phiên mã tạo mRNA theo NTBS và chiều 5’ → 3’:\n3’–GAC – CAG – TGC – GAC–5’\n→ mRNA (5’ → 3’): 5’–CUG – GUC – ACG – CUG–3’\n- Đối chiếu bảng mã di truyền:\n+ 5’CUG3’ → Leu\n+ 5’GUC3’ → Val\n+ 5’ACG3’ → Thr\n+ 5’CUG3’ → Leu\nĐoạn polypeptide được tạo ra là: Leu – Val – Thr – Leu."
    }
]

# ==============================================================================
# HÀM TẠO FILE WORD
# ==============================================================================

def create_word_document(mode="student"):
    """
    Tạo file Word theo Form Đề Kiểm Tra:
    mode: 'student', 'teacher', 'screen'
    """
    is_teacher = (mode == "teacher")
    is_screen = (mode == "screen")

    if mode == "teacher":
        filename = "Chuyen_De_Dich_Ma_Gene_Den_Tinh_Trang_GiaoVien.docx"
    elif mode == "screen":
        filename = "Chuyen_De_Dich_Ma_Gene_Den_Tinh_Trang_DayTrenManHinh.docx"
    else:
        filename = "Chuyen_De_Dich_Ma_Gene_Den_Tinh_Trang_HocSinh.docx"

    filepath = os.path.join(OUT_DIR, filename)

    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="LỚP KHTN THẦY THẠNH CÔ CHÂU",
        exam_title="CHUYÊN ĐỀ ÔN TẬP SINH HỌC 12",
        subject_line="DỊCH MÃ VÀ MỐI QUAN HỆ TỪ GENE ĐẾN TÍNH TRẠNG",
        num_pages="06",
        duration="90 phút",
        exam_code="DM-12",
        atomic_masses=None,
        is_teacher=is_teacher
    )

    base_size = 11.5 if is_screen else 10.5

    # --------------------------------------------------------------------------
    # PHẦN I. TRẮC NGHIỆM NHIỀU LỰA CHỌN
    # --------------------------------------------------------------------------
    add_section_header(
        doc,
        "PHẦN I. CÂU HỎI TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN",
        "Thí sinh trả lời từ câu 1 đến câu 33. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
    )

    for q in PART1_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["text"], base_size=base_size)

        render_options_no_table(doc, q["options"], is_teacher=is_teacher)

        if is_teacher and q.get("solution"):
            add_solution_no_table(doc, q["solution"])

    # --------------------------------------------------------------------------
    # PHẦN II. TRẮC NGHIỆM ĐÚNG/SAI
    # --------------------------------------------------------------------------
    add_section_header(
        doc,
        "PHẦN II. CÂU HỎI TRẮC NGHIỆM ĐÚNG/SAI",
        "Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
    )

    for q in PART2_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["text"], base_size=base_size)

        for stmt, ans in q["items"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt, ans=ans, is_teacher=is_teacher, base_size=base_size)

        if is_teacher and q.get("solution"):
            add_solution_no_table(doc, q["solution"])

    # --------------------------------------------------------------------------
    # PHẦN III. TRẢ LỜI NGẮN
    # --------------------------------------------------------------------------
    add_section_header(
        doc,
        "PHẦN III. CÂU HỎI TRẮC NGHIỆM TRẢ LỜI NGẮN",
        "Thí sinh trả lời từ câu 1 đến câu 20. Viết câu trả lời hoặc số thích hợp vào chỗ trống."
    )

    for q in PART3_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["text"], base_size=base_size)

        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(2)

        if is_teacher:
            r_lbl = p_ans.add_run("Đáp án: ")
            set_run_style(r_lbl, font_size=base_size, bold=True, color=COLOR_CORRECT)
            r_val = p_ans.add_run(q["ans"])
            set_run_style(r_val, font_size=base_size, bold=True, color=COLOR_CORRECT)
            if q.get("solution"):
                add_solution_no_table(doc, q["solution"])
        else:
            r_lbl = p_ans.add_run("Trả lời: ....................................................................................................................")
            set_run_style(r_lbl, font_size=base_size, color=RGBColor(120, 120, 120))

    # --------------------------------------------------------------------------
    # PHẦN IV. TỰ LUẬN
    # --------------------------------------------------------------------------
    add_section_header(
        doc,
        "PHẦN IV. CÂU HỎI TỰ LUẬN",
        "Thí sinh trình bày chi tiết lời giải từ câu 1 đến câu 8."
    )

    for q in PART4_QUESTIONS:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["text"], base_size=base_size)

        if is_teacher:
            add_solution_no_table(doc, q["solution"])
        else:
            add_student_writing_lines(doc, num_lines=4)

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
# HÀM TẠO FILE LATEX VÀ BIÊN DỊCH PDF TRÌNH CHIẾU TIVI 16:9
# ==============================================================================

def latex_clean(text):
    """Làm sạch văn bản để đưa vào LaTeX không bị lỗi kí tự đặc biệt và hiển thị đẹp."""
    t = text.replace("’", "'").replace("‘", "'")
    t = t.replace("→", r"$\rightarrow$").replace("——>", r"$\longrightarrow$").replace("—>", r"$\longrightarrow$").replace("->", r"$\rightarrow$")
    t = t.replace("↔", r"$\longleftrightarrow$").replace("<->", r"$\longleftrightarrow$")
    t = t.replace("&", r"\&").replace("%", r"\%")
    return t

def render_q_latex(q):
    lines = []
    num = q["num"]
    txt = latex_clean(q["text"])
    if txt.startswith(f"Câu {num}."):
        txt = r"\textbf{Câu " + str(num) + r".} " + txt[len(f"Câu {num}."):].strip()
    
    txt = txt.replace("\n", r"\\[3pt]")
    lines.append(txt + r"\\[4pt]")

    opts = [latex_clean(o.replace("*", "")) for o in q["options"]]
    clean_opts = []
    for o in opts:
        lbl = o[:3]
        body = o[3:].strip()
        clean_opts.append(r"\textbf{" + lbl + r"} " + body)
        
    lines.append(r"\begin{tabularx}{\linewidth}{@{}X X@{}}")
    lines.append(f"{clean_opts[0]} & {clean_opts[1]}\\\\[4pt]")
    lines.append(f"{clean_opts[2]} & {clean_opts[3]}")
    lines.append(r"\end{tabularx}")
    return "\n".join(lines)

def generate_tivi_presentation_pdf():
    tex_filename = "Chuyen_De_Dich_Ma_Gene_Den_Tinh_Trang_DayTrenManHinh.tex"
    pdf_filename = "Chuyen_De_Dich_Ma_Gene_Den_Tinh_Trang_DayTrenManHinh.pdf"
    tex_path = os.path.join(SCRATCH_DIR, tex_filename)
    pdf_dest = os.path.join(OUT_DIR, pdf_filename)

    tex_lines = [r"""\documentclass{article}
\usepackage[utf8]{vietnam}
\usepackage{times}
\usepackage{amsmath,amssymb}
\usepackage{tabularx}
\usepackage{array}
\usepackage{booktabs}
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
\raggedbottom

\definecolor{brandRed}{RGB}{211, 47, 47}
\definecolor{brandGreen}{RGB}{46, 125, 50}
\definecolor{brandBlue}{RGB}{21, 101, 192}

\begin{document}\noindent\begin{minipage}[t][14.4cm][c]{\linewidth}
\begin{center}
{\fontsize{34}{44}\selectfont \textbf{\color{brandRed}LỚP KHTN THẦY THẠNH CÔ CHÂU}}\\[16pt]
{\fontsize{24}{32}\selectfont \textbf{CHUYÊN ĐỀ SINH HỌC 12}}\\[22pt]
{\fontsize{38}{48}\selectfont \textbf{\color{brandGreen}DỊCH MÃ VÀ MỐI QUAN HỆ TỪ GENE ĐẾN TÍNH TRẠNG}}\\[26pt]
{\fontsize{22}{30}\selectfont \textit{Hệ thống câu hỏi Trắc nghiệm nhiều lựa chọn, Đúng/Sai, Trả lời ngắn \& Tự luận}}\\[14pt]
{\fontsize{20}{28}\selectfont \textit{Tài liệu trình chiếu màn hình tương tác / Tivi 65 inch}}
\end{center}
\end{minipage}
\newpage
"""]

    # --------------------------------------------------------------------------
    # PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (2 câu / slide)
    # --------------------------------------------------------------------------
    p1 = PART1_QUESTIONS
    for i in range(0, len(p1), 2):
        q1 = p1[i]
        q2 = p1[i+1] if i+1 < len(p1) else None
        
        tot_len = len(q1["text"]) + sum(len(o) for o in q1["options"])
        if q2:
            tot_len += len(q2["text"]) + sum(len(o) for o in q2["options"])
            
        if tot_len < 320:
            fsize, fbase = 25, 33
        elif tot_len < 500:
            fsize, fbase = 22, 29
        elif tot_len < 700:
            fsize, fbase = 20, 26
        else:
            fsize, fbase = 18, 23
            
        tex_lines.append(f"""
\\begin{{minipage}}[t][14.4cm][t]{{\\linewidth}}
\\fontsize{{{fsize}}}{{{fbase}}}\\selectfont
""")
        tex_lines.append(render_q_latex(q1))
        if q2:
            tex_lines.append(r"\vspace*{\fill}")
            tex_lines.append(render_q_latex(q2))
        else:
            tex_lines.append(r"\vspace*{\fill}")
            
        tex_lines.append(r"""
\end{minipage}
\newpage
""")

    # --------------------------------------------------------------------------
    # PHẦN II: TRẮC NGHIỆM ĐÚNG/SAI (1 câu / slide)
    # --------------------------------------------------------------------------
    for q in PART2_QUESTIONS:
        tot_len = len(q["text"]) + sum(len(s[0]) for s in q["items"])
        if tot_len < 400:
            fsize, fbase = 25, 33
        elif tot_len < 650:
            fsize, fbase = 22, 29
        else:
            fsize, fbase = 19, 25
            
        q_body = latex_clean(q["text"][len(f"Câu {q['num']}."):].strip()).replace("\n", r"\\[4pt]")
        tex_lines.append(f"""
\\begin{{minipage}}[t][14.4cm][t]{{\\linewidth}}
\\fontsize{{{fsize}}}{{{fbase}}}\\selectfont
{{\\color{{brandGreen}}\\textbf{{PHẦN II: CÂU TRẮC NGHIỆM ĐÚNG / SAI}}\\\\[8pt]}}
\\textbf{{Câu {q['num']}.}} {q_body}
\\vspace{{8pt}}

\\begin{{tabularx}}{{\\linewidth}}{{@{{}}X r@{{}}}}
""")
        for stmt, _ in q["items"]:
            clean_stmt = latex_clean(stmt)
            lbl = clean_stmt[:3]
            body = clean_stmt[3:].strip()
            tex_lines.append(r"\textbf{" + lbl + r"} " + body + r" & $[\quad]$\\[8pt]")
            
        tex_lines.append(r"""\end{tabularx}
\vfill
\end{minipage}
\newpage
""")

    # --------------------------------------------------------------------------
    # PHẦN III: TRẢ LỜI NGẮN (1 câu / slide)
    # --------------------------------------------------------------------------
    for q in PART3_QUESTIONS:
        tot_len = len(q["text"])
        if tot_len < 150:
            fsize, fbase = 28, 38
        elif tot_len < 300:
            fsize, fbase = 25, 34
        else:
            fsize, fbase = 22, 30
            
        q_body = latex_clean(q["text"][len(f"Câu {q['num']}."):].strip()).replace("\n", r"\\[6pt]")
        tex_lines.append(f"""
\\begin{{minipage}}[t][14.4cm][t]{{\\linewidth}}
\\fontsize{{{fsize}}}{{{fbase}}}\\selectfont
{{\\color{{brandGreen}}\\textbf{{PHẦN III: CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN}}\\\\[10pt]}}
\\textbf{{Câu {q['num']}.}} {q_body}
\\vspace{{20pt}}

\\textbf{{Trả lời:}} $\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots\\dots$
\\vfill
\\end{{minipage}}
\\newpage
""")

    # --------------------------------------------------------------------------
    # PHẦN IV: TỰ LUẬN (1 câu / slide)
    # --------------------------------------------------------------------------
    for q in PART4_QUESTIONS:
        tot_len = len(q["text"])
        if tot_len < 200:
            fsize, fbase = 26, 35
        elif tot_len < 400:
            fsize, fbase = 23, 31
        else:
            fsize, fbase = 20, 27
            
        q_body = latex_clean(q["text"][len(f"Câu {q['num']}."):].strip()).replace("\n", r"\\[6pt]")
        tex_lines.append(f"""
\\begin{{minipage}}[t][14.4cm][t]{{\\linewidth}}
\\fontsize{{{fsize}}}{{{fbase}}}\\selectfont
{{\\color{{brandGreen}}\\textbf{{PHẦN IV: CÂU HỎI TỰ LUẬN}}\\\\[10pt]}}
\\textbf{{Câu {q['num']}.}} {q_body}
\\vspace{{14pt}}

\\textbf{{Bài làm:}}
\\vfill
\\end{{minipage}}
\\newpage
""")

    # --------------------------------------------------------------------------
    # SLIDE BẢNG ĐÁP ÁN CHÍNH THỨC
    # --------------------------------------------------------------------------
    tex_lines.append(r"""
\begin{minipage}[t][14.4cm][t]{\linewidth}
\fontsize{19}{25}\selectfont
\begin{center}
{\fontsize{26}{34}\selectfont \textbf{\color{brandRed}BẢNG ĐÁP ÁN THAM KHẢO}}\\[10pt]
\textbf{\color{brandBlue}PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN}\\[4pt]
\begin{tabular}{|c|c||c|c||c|c||c|c||c|c||c|c|}
\hline
\textbf{Câu} & \textbf{ĐA} & \textbf{Câu} & \textbf{ĐA} & \textbf{Câu} & \textbf{ĐA} & \textbf{Câu} & \textbf{ĐA} & \textbf{Câu} & \textbf{ĐA} & \textbf{Câu} & \textbf{ĐA}\\
\hline
1 & B & 7 & D & 13 & B & 19 & D & 25 & C & 31 & C\\
2 & A & 8 & A & 14 & C & 20 & B & 26 & A & 32 & B\\
3 & C & 9 & A & 15 & C & 21 & B & 27 & D & 33 & C\\
4 & C & 10 & D & 16 & A & 22 & A & 28 & D & & \\
5 & B & 11 & C & 17 & B & 23 & A & 29 & B & & \\
6 & B & 12 & D & 18 & C & 24 & D & 30 & C & & \\
\hline
\end{tabular}

\vspace{10pt}
\textbf{\color{brandBlue}PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI}\\[4pt]
\begin{tabular}{|c|c|c|c|c|}
\hline
\textbf{Câu} & \textbf{Ý a} & \textbf{Ý b} & \textbf{Ý c} & \textbf{Ý d}\\
\hline
\textbf{1} & Đúng & Đúng & Sai & Đúng\\
\textbf{2} & Đúng & Đúng & Sai & Sai\\
\textbf{3} & Sai & Đúng & Sai & Sai\\
\textbf{4} & Đúng & Sai & Đúng & Đúng\\
\hline
\end{tabular}

\vspace{10pt}
\textbf{\color{brandBlue}PHẦN III: TRẢ LỜI NGẮN}\\[4pt]
\begin{tabular}{|c|c||c|c||c|c||c|c|}
\hline
\textbf{Câu} & \textbf{Đáp án} & \textbf{Câu} & \textbf{Đáp án} & \textbf{Câu} & \textbf{Đáp án} & \textbf{Câu} & \textbf{Đáp án}\\
\hline
1 & 5'UCG3' & 6 & 600 & 11 & 5'ATG...TGA3' & 16 & 5'AUG3'\\
2 & mRNA & 7 & 200 & 12 & Gly-Ser-Tyr-Ala-Pro & 17 & UAA, UAG, UGA\\
3 & Met & 8 & Tế bào chất & 13 & Codon & 18 & Ribosome\\
4 & 5'AUA...U3' & 9 & tRNA & 14 & 64 & 19 & tRNA\\
5 & 24 & 10 & 2495 & 15 & 61 & 20 & 3 giai đoạn\\
\hline
\end{tabular}
\end{center}
\end{minipage}
\newpage
""")

    tex_lines.append(r"\end{document}")

    with open(tex_path, "w", encoding="utf-8") as f:
        f.write("\n".join(tex_lines))
    print(f"[OK] Đã ghi file LaTeX: {tex_path}")

    cmd = ["pdflatex", "-interaction=nonstopmode", "-output-directory", SCRATCH_DIR, tex_path]
    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    pdf_tmp = os.path.join(SCRATCH_DIR, pdf_filename)
    if os.path.exists(pdf_tmp):
        import shutil
        import pymupdf
        shutil.copy2(pdf_tmp, pdf_dest)
        doc = pymupdf.open(pdf_dest)
        print(f"[OK] Đã biên dịch thành công PDF Tivi 16:9: {pdf_dest} ({os.path.getsize(pdf_dest):,} bytes) với {len(doc)} slide chuẩn!")
    else:
        print("[ERROR] Lỗi biên dịch pdflatex:")
        print(res.stdout[-1500:])


# ==============================================================================
# HÀM CHÍNH
# ==============================================================================

if __name__ == "__main__":
    print("--- BẮT ĐẦU TẠO BỘ TÀI LIỆU DỊCH MÃ VÀ GENE ĐẾN TÍNH TRẠNG ---")
    
    # 1. Tạo bản Học Sinh
    create_word_document(mode="student")

    # 2. Tạo bản Giáo Viên
    create_word_document(mode="teacher")

    # 3. Tạo bản Dạy Trên Màn Hình (Word)
    create_word_document(mode="screen")

    # 4. Tạo bản Dạy Trên Màn Hình (PDF Tivi 16:9)
    generate_tivi_presentation_pdf()

    print("--- HOÀN THÀNH TẤT CẢ TÀI LIỆU! ---")
