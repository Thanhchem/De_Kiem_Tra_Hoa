# -*- coding: utf-8 -*-
"""
Script sinh 2 Đề Kiểm Tra Giữa Học Kì I Hóa Học 12 (Năm học 2026-2027)
Chuẩn xác 100% theo Ma Trận Đề Kiểm Tra do người dùng cung cấp:
- Đề 01 (Mã 101) & Đề 02 (Mã 102)
- Mỗi đề gồm 4 phần (27 câu / lệnh hỏi, thang điểm 10):
  + PHẦN I: 12 câu TN nhiều lựa chọn (3,0 điểm)
  + PHẦN II: 2 câu TN Đúng - Sai (8 ý a,b,c,d - 2,0 điểm)
  + PHẦN III: 4 câu TN Trả lời ngắn (2,0 điểm)
  + PHẦN IV: 3 câu Tự luận Vận dụng (3,0 điểm)
- Chuẩn FORM ĐỀ KIỂM TRA (Thầy TRẦN VĂN THẠNH):
  + Xuất bản Học Sinh & Giáo Viên có giải chi tiết cho mỗi đề.
"""

import os
import sys
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Import module form đề kiểm tra chuẩn
sys.path.append(os.path.abspath(r"C:\Antigravity_Thanh\scripts"))
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
    COLOR_SOLUTION
)

OUT_DIR = r"C:\Antigravity_Thanh\San_Pham"

# ==============================================================================
# DỮ LIỆU ĐỀ 01 (MÃ ĐỀ 101)
# ==============================================================================
EXAM_01 = {
    "de_num": 1,
    "made": "101",
    "title": "ĐỀ KIỂM TRA GIỮA HỌC KỲ I - ĐỀ SỐ 01",
    # PHẦN I: 12 CÂU
    "part1": [
        {
            "idx": 1,
            "topic": "Bài 1: Ester - Lipit (Biết)",
            "prompt": "Chất béo tristearin có công thức phân tử là",
            "opts": [
                "A. (C~17~H~35~COO)~3~C~3~H~5~.",
                "B. (C~17~H~33~COO)~3~C~3~H~5~.",
                "C. (C~15~H~31~COO)~3~C~3~H~5~.",
                "D. (C~17~H~31~COO)~3~C~3~H~5~."
            ],
            "correct": "A",
            "sol": "Tristearin là triester của glycerol với stearic acid (C~17~H~35~COOH), có công thức là (C~17~H~35~COO)~3~C~3~H~5~ hay C~57~H~110~O~6~ (chất béo no, trạng thái rắn ở nhiệt độ phòng)."
        },
        {
            "idx": 2,
            "topic": "Bài 4: Carbohydrat. Glucose và fructose (Biết)",
            "prompt": "Trong điều kiện thường, carbohydrate nào sau đây là chất rắn kết tinh, không màu, có vị ngọt, dễ tan trong nước và có nhiều trong quả nho chín?",
            "opts": [
                "A. Fructose.",
                "B. Glucose.",
                "C. Saccharose.",
                "D. Cellulose."
            ],
            "correct": "B",
            "sol": "Glucose (đường nho) là chất rắn kết tinh, không màu, dễ tan trong nước, có vị ngọt (chiếm khoảng 20% trong quả nho chín)."
        },
        {
            "idx": 3,
            "topic": "Bài 4: Carbohydrat. Glucose và fructose (Hiểu)",
            "prompt": "Để chứng minh trong phân tử glucose có nhiều nhóm hydroxy (-OH) kề nhau, người ta cho dung dịch glucose tác dụng với thuốc thử nào sau đây ở nhiệt độ thường?",
            "opts": [
                "A. Cu(OH)~2~ trong môi trường kiềm.",
                "B. Thuốc thử Tollens ([Ag(NH~3~)~2~]OH), đun nóng.",
                "C. Dung dịch nước bromine (Br~2~).",
                "D. Cu(OH)~2~ đun nóng tạo kết tủa đỏ gạch."
            ],
            "correct": "A",
            "sol": "Glucose hòa tan Cu(OH)~2~ ở nhiệt độ thường tạo phức đồng dung dịch màu xanh lam thẫm, chứng minh phân tử có nhiều nhóm -OH liền kề (tính chất của polyalcohol)."
        },
        {
            "idx": 4,
            "topic": "Bài 5: Saccharose và maltose (Biết)",
            "prompt": "Phân tử saccharose được tạo nên từ liên kết glycosidic giữa",
            "opts": [
                "A. một gốc α-glucose và một gốc β-fructose qua liên kết α-1,2-glycosidic.",
                "B. hai gốc α-glucose liên kết qua liên kết α-1,4-glycosidic.",
                "C. một gốc α-glucose và một gốc α-galactose qua liên kết α-1,4-glycosidic.",
                "D. hai gốc β-fructose liên kết qua liên kết β-1,2-glycosidic."
            ],
            "correct": "A",
            "sol": "Saccharose là disaccharide tạo bởi 1 gốc α-glucose và 1 gốc β-fructose qua liên kết α-1,2-glycosidic. Do không còn nhóm -OH hemiacetal tự do nên saccharose không có tính khử (không tráng bạc)."
        },
        {
            "idx": 5,
            "topic": "Bài 6: Tinh bột và cellulose (Biết)",
            "prompt": "Thành phần nào của tinh bột có cấu trúc mạch phân nhánh do các đơn vị glucose liên kết với nhau bằng cả liên kết α-1,4-glycosidic và liên kết α-1,6-glycosidic?",
            "opts": [
                "A. Amylose.",
                "B. Amylopectin.",
                "C. Cellulose.",
                "D. Maltose."
            ],
            "correct": "B",
            "sol": "Tinh bột gồm 2 thành phần: Amylose có cấu trúc mạch không phân nhánh (chỉ chứa liên kết α-1,4-glycosidic) xoắn lò xo; Amylopectin có cấu trúc mạch phân nhánh (chứa liên kết α-1,4 và α-1,6-glycosidic)."
        },
        {
            "idx": 6,
            "topic": "Bài 6: Tinh bột và cellulose (Biết)",
            "prompt": "Phân tử cellulose được cấu tạo từ nhiều đơn vị nào liên kết với nhau bằng liên kết β-1,4-glycosidic tạo thành chuỗi polymer không phân nhánh?",
            "opts": [
                "A. α-glucose.",
                "B. β-glucose.",
                "C. β-fructose.",
                "D. α-galactose."
            ],
            "correct": "B",
            "sol": "Cellulose là polysaccharide gồm nhiều mắt xích β-glucose liên kết với nhau qua liên kết β-1,4-glycosidic tạo thành mạch kéo dài, không phân nhánh, không xoắn."
        },
        {
            "idx": 7,
            "topic": "Bài 6: Tinh bột và cellulose (Biết)",
            "prompt": "Vật liệu tơ nhân tạo (bán tổng hợp) nào sau đây được sản xuất từ nguyên liệu cellulose?",
            "opts": [
                "A. Tơ tằm.",
                "B. Tơ capron.",
                "C. Tơ visco.",
                "D. Tơ nylon-6,6."
            ],
            "correct": "C",
            "sol": "Tơ visco và tơ acetate là các loại tơ bán tổng hợp (nhân tạo) được chế tạo từ cellulose tự nhiên qua quá trình biến đổi hóa học thích hợp."
        },
        {
            "idx": 8,
            "topic": "Bài 6: Tinh bột và cellulose (Hiểu)",
            "prompt": "Cho cellulose phản ứng với dung dịch HNO~3~ đặc (xúc tác H~2~SO~4~ đặc, đun nóng), thu được cellulose trinitrate dùng làm thuốc súng không khói. Phân tử khối của một mắt xích cellulose trinitrate [C~6~H~7~O~2~(ONO~2~)~3~] là",
            "opts": [
                "A. 162.",
                "B. 297.",
                "C. 252.",
                "D. 207."
            ],
            "correct": "B",
            "sol": "Mỗi mắt xích cellulose [C~6~H~7~O~2~(OH)~3~] có phân tử khối 162. Khi phản ứng thế 3 nhóm -OH bằng 3 nhóm -ONO~2~: M = 162 + 3 × (62 - 17) = 162 + 135 = 297."
        },
        {
            "idx": 9,
            "topic": "Bài 8: Amine (Hiểu)",
            "prompt": "Dãy các chất nào sau đây được sắp xếp theo thứ tự tính base tăng dần?",
            "opts": [
                "A. C~6~H~5~NH~2~ < NH~3~ < CH~3~NH~2~ < (CH~3~)~2~NH.",
                "B. NH~3~ < C~6~H~5~NH~2~ < CH~3~NH~2~ < (CH~3~)~2~NH.",
                "C. (CH~3~)~2~NH < CH~3~NH~2~ < NH~3~ < C~6~H~5~NH~2~.",
                "D. C~6~H~5~NH~2~ < (CH~3~)~2~NH < CH~3~NH~2~ < NH~3~."
            ],
            "correct": "A",
            "sol": "Gốc phenyl (C~6~H~5~-) hút electron làm giảm mật độ electron trên nguyên tử N → tính base của aniline yếu hơn NH~3~. Gốc alkyl (-CH~3~) đẩy electron làm tăng mật độ electron trên N → tính base tăng: C~6~H~5~NH~2~ < NH~3~ < CH~3~NH~2~ < (CH~3~)~2~NH."
        },
        {
            "idx": 10,
            "topic": "Bài 9: Amino acid và peptide (Biết)",
            "prompt": "Chất nào sau đây là một α-amino acid có mạch carbon phân nhánh và chứa 5 nguyên tử carbon trong phân tử?",
            "opts": [
                "A. Glycine.",
                "B. Alanine.",
                "C. Valine.",
                "D. Glutamic acid."
            ],
            "correct": "C",
            "sol": "Valine là α-amino acid có công thức (CH~3~)~2~CH-CH(NH~2~)-COOH, chứa 5 carbon, mạch phân nhánh. Glycine có 2C, Alanine có 3C, Glutamic acid có 5C nhưng mạch thẳng và có 2 nhóm -COOH."
        },
        {
            "idx": 11,
            "topic": "Bài 9: Amino acid và peptide (Hiểu)",
            "prompt": "Chất nào sau đây phản ứng với Cu(OH)~2~ trong môi trường kiềm tạo thành dung dịch có màu tím đặc trưng (phản ứng màu biuret)?",
            "opts": [
                "A. Gly-Ala.",
                "B. Gly-Ala-Val.",
                "C. Glucose.",
                "D. Alanine."
            ],
            "correct": "B",
            "sol": "Điều kiện tham gia phản ứng màu biuret là phân tử phải chứa từ 2 liên kết peptide trở lên (tripeptide trở lên hoặc protein). Gly-Ala là dipeptide (chỉ có 1 liên kết peptide) nên không có phản ứng biuret. Gly-Ala-Val là tripeptide nên cho màu tím đặc trưng."
        },
        {
            "idx": 12,
            "topic": "Bài 10: Protein và enzyme (Biết)",
            "prompt": "Hiện tượng sữa đậu nành bị vón cục (đông tụ) khi nhỏ nước cốt chanh vào là do sự đông tụ protein gây ra bởi",
            "opts": [
                "A. nhiệt độ cao.",
                "B. môi trường acid (pH thấp).",
                "C. ion kim loại nặng.",
                "D. tác dụng xúc tác của enzyme."
            ],
            "correct": "B",
            "sol": "Nước cốt chanh chứa citric acid (môi trường acid, pH thấp). Khi thay đổi pH đến gần điểm đẳng điện, protein bị mất lớp vỏ hydrate và trung hòa điện tích dẫn tới đông tụ, vón cục."
        }
    ],

    # PHẦN II: 2 CÂU ĐÚNG - SAI
    "part2": [
        {
            "idx": 1,
            "topic": "Chủ đề 1: Ester - Lipid. Xà phòng và chất giặt rửa",
            "intro": "Tiến hành thí nghiệm xà phòng hóa chất béo trong phòng thí nghiệm: Cho vào bát sứ nhỏ khoảng 2 mL dầu dừa (hoặc mỡ lợn) và 5 mL dung dịch NaOH 40%. Đun sôi nhẹ hỗn hợp trong khoảng 30 phút và liên tục khuấy bằng đũa thủy tinh, thỉnh thoảng thêm vài giọt nước cất để giữ thể tích không đổi. Sau đó, rót thêm vào hỗn hợp 8 mL dung dịch NaCl bão hòa nóng, khuấy nhẹ rồi để nguội.",
            "items": [
                ("a) Mục đích chính của việc thêm dung dịch NaCl bão hòa là để làm tăng khối lượng riêng của dung dịch, giúp muối sodium của acid béo (xà phòng) nổi lên trên.", "Đúng"),
                ("b) Sau khi để nguội, khối chất rắn màu trắng nổi lên trên bề mặt bát sứ chứa thành phần chính là muối sodium của các acid béo.", "Đúng"),
                ("c) Trong quá trình đun sôi, nếu không bổ sung nước cất thì hỗn hợp có thể bị cạn khô, dẫn tới phản ứng thủy phân không hoàn toàn và xà phòng bị cháy khét.", "Đúng"),
                ("d) Phần dung dịch lỏng phía dưới bát sứ sau khi vớt tách xà phòng có chứa glycerol, dung dịch này có khả năng hòa tan kết tủa Cu(OH)~2~ tạo phức chất màu xanh lam thẫm.", "Đúng")
            ],
            "sol": "• a) ĐÚNG: NaCl bão hòa làm giảm độ tan của muối acid béo (hiệu ứng muối kết) và làm tăng khối lượng riêng dung dịch lỏng phía dưới giúp xà phòng nhẹ hơn nổi lên.\n• b) ĐÚNG: Xà phòng là muối sodium hoặc potassium của các acid béo, tách ra ở dạng rắn xốp màu trắng.\n• c) ĐÚNG: Nước vừa là dung môi, vừa tham gia vào quá trình hòa tan dung dịch NaOH; nếu cạn thì nhiệt độ tăng cao làm than hóa/cháy khét hỗn hợp.\n• d) ĐÚNG: Dung dịch dưới chứa glycerol và NaCl dư. Glycerol là polyalcohol liền kề tạo phức màu xanh lam đặc trưng với Cu(OH)~2~."
        },
        {
            "idx": 2,
            "topic": "Chủ đề 3: Hợp chất chứa nitrogen",
            "intro": "Cho các hợp chất hữu cơ sau: methylamine (CH~3~NH~2~), aniline (C~6~H~5~NH~2~), glycine (H~2~N-CH~2~-COOH), lysine (H~2~N-(CH~2~)~4~-CH(NH~2~)-COOH) và dung dịch lòng trắng trứng (albumin).",
            "items": [
                ("a) Methylamine ở điều kiện thường là chất khí, mùi khai khó chịu, độc và tan rất nhiều trong nước do tạo được liên kết hydrogen liên phân tử với nước.", "Đúng"),
                ("b) Glycine là chất rắn kết tinh, có nhiệt độ nóng chảy cao và tồn tại chủ yếu ở dạng ion lưỡng cực (+H~3~N-CH~2~-COO^-) trong cả tinh thể và dung dịch nước.", "Đúng"),
                ("c) Khi nhúng giấy quỳ tím vào dung dịch lysine 0,1 M thì quỳ tím sẽ chuyển sang màu xanh do phân tử chứa 2 nhóm -NH~2~ và 1 nhóm -COOH nên dung dịch có tính base.", "Đúng"),
                ("d) Nhỏ vài giọt dung dịch HNO~3~ đặc vào ống nghiệm chứa dung dịch albumin, đun nóng nhẹ thấy xuất hiện kết tủa màu vàng (phản ứng xanthoproteic).", "Đúng")
            ],
            "sol": "• a) ĐÚNG: CH~3~NH~2~ là chất khí ở nhiệt độ thường, tan vô hạn trong nước nhờ tạo liên kết hydrogen mạnh với nước.\n• b) ĐÚNG: Amino acid có cấu trúc ion lưỡng cực nên có tính chất giống hợp chất ion: thể rắn, nhiệt độ nóng chảy cao, tan tốt trong nước.\n• c) ĐÚNG: Phân tử lysine có số nhóm -NH~2~ (2) nhiều hơn số nhóm -COOH (1), môi trường base làm quỳ tím hóa xanh.\n• d) ĐÚNG: Albumin chứa các gốc amino acid có vòng benzene (phenylalanine, tyrosine), phản ứng nitro hóa tạo hợp chất nitro thơm màu vàng."
        }
    ],

    # PHẦN III: 4 CÂU TRẢ LỜI NGẮN
    "part3": [
        {
            "idx": 1,
            "topic": "Bài 5: Saccharose và maltose (Biết)",
            "prompt": "Trong một phân tử saccharose (C~12~H~22~O~11~), có bao nhiêu nhóm hydroxy (-OH) tự do?",
            "ans": "8",
            "sol": "Phân tử saccharose tạo từ 1 gốc α-glucose (còn 4 nhóm -OH tự do) và 1 gốc β-fructose (còn 4 nhóm -OH tự do). Liên kết glycosidic tạo giữa C1 của glucose và C2 của fructose làm mất 2 nhóm -OH hemiacetal/hemiketal. Tổng số nhóm -OH còn lại là 4 + 4 = 8 nhóm."
        },
        {
            "idx": 2,
            "topic": "Bài 6: Tinh bột và cellulose (Hiểu)",
            "prompt": "Từ 16,2 tấn mùn cưa chứa 50% cellulose, người ta tiến hành sản xuất ethylic alcohol (C~2~H~5~OH) qua hai giai đoạn thủy phân và lên men với hiệu suất toàn bộ quá trình đạt 80%. Tính khối lượng ethylic alcohol thu được (theo đơn vị tấn, làm tròn đến chữ số thập phân thứ nhất).",
            "ans": "3,7",
            "sol": "Khối lượng cellulose nguyên chất: m = 16,2 × 50% = 8,1 tấn.\nSơ đồ: (C~6~H~10~O~5~)~n~ → n C~6~H~12~O~6~ → 2n C~2~H~5~OH\nCứ 162 tấn cellulose tạo ra 2 × 46 = 92 tấn C~2~H~5~OH.\nKhối lượng C~2~H~5~OH theo lí thuyết: m_lt = (8,1 × 92) / 162 = 4,6 tấn.\nVì hiệu suất toàn bộ quá trình là 80%: m_tt = 4,6 × 80% = 3,68 tấn ≈ 3,7 tấn."
        },
        {
            "idx": 3,
            "topic": "Bài 9: Amino acid và peptide (Hiểu)",
            "prompt": "Cho một oligopeptide mạch hở X có công thức cấu tạo: Gly-Ala-Val-Ala-Gly. Để thủy phân hoàn toàn 0,1 mol X cần dùng vừa đủ bao nhiêu mol NaOH trong dung dịch khi đun nóng?",
            "ans": "0,5",
            "sol": "X là một pentapeptide tạo bởi các amino acid no chứa 1 nhóm -NH~2~ và 1 nhóm -COOH (Gly, Ala, Val). Phương trình thủy phân:\nX + 5NaOH → Muối + H~2~O\nn_NaOH = 5 × n_X = 5 × 0,1 = 0,5 mol."
        },
        {
            "idx": 4,
            "topic": "Bài 10: Protein và enzyme (Biết)",
            "prompt": "Cho các chất sau: glycerol, saccharose, lòng trắng trứng (albumin), glycylalanine (Gly-Ala), tripeptide Gly-Ala-Val. Có bao nhiêu chất tham gia phản ứng tạo dung dịch màu tím đặc trưng với thuốc thử biuret (Cu(OH)~2~ trong môi trường kiềm)?",
            "ans": "2",
            "sol": "Chỉ các peptide có từ 2 liên kết peptide trở lên (tripeptide trở lên) và protein mới có phản ứng màu biuret.\n- Glycerol hòa tan Cu(OH)~2~ tạo dung dịch màu xanh lam (không phải màu tím biuret).\n- Saccharose tạo dung dịch màu xanh lam.\n- Gly-Ala là dipeptide (chỉ 1 liên kết peptide) → không phản ứng màu biuret.\n- Lòng trắng trứng (albumin) và tripeptide Gly-Ala-Val có phản ứng màu biuret → Tổng có 2 chất."
        }
    ],

    # PHẦN IV: 3 CÂU TỰ LUẬN
    "part4": [
        {
            "idx": 1,
            "topic": "Bài 1: Ester - Lipit (Vận dụng)",
            "prompt": "Thủy phân hoàn toàn 17,24 gam một loại chất béo X (triglyceride) cần dùng vừa đủ 0,06 mol NaOH trong dung dịch, đun nóng, thu được glycerol và m gam hỗn hợp muối của các acid béo (xà phòng).\na) Viết phương trình hóa học tổng quát của phản ứng xà phòng hóa chất béo X.\nb) Tính khối lượng glycerol thu được (theo gam).\nc) Tính khối lượng xà phòng m (theo gam) thu được sau phản ứng.",
            "sol": "a) Phương trình hóa học tổng quát:\n(RCOO)~3~C~3~H~5~ + 3NaOH → 3RCOONa + C~3~H~5~(OH)~3~\n\nb) Theo phương trình phản ứng:\nn_glycerol = (1/3) × n_NaOH = (1/3) × 0,06 = 0,02 mol.\nKhối lượng glycerol thu được:\nm_glycerol = 0,02 × 92 = 1,84 gam.\n\nc) Áp dụng định luật bảo toàn khối lượng:\nm_chất béo + m_NaOH = m_muối + m_glycerol\n⇒ m_muối = 17,24 + (0,06 × 40) - 1,84 = 17,24 + 2,40 - 1,84 = 17,80 gam.\nVậy khối lượng xà phòng m = 17,80 gam."
        },
        {
            "idx": 2,
            "topic": "Bài 4: Glucose và fructose (Vận dụng)",
            "prompt": "Lên men m gam glucose để sản xuất ethylic alcohol (rượu vang) với hiệu suất của quá trình đạt 75%. Toàn bộ lượng khí CO~2~ sinh ra được dẫn hấp thụ hoàn toàn vào dung dịch nước vôi trong dư, sau phản ứng thu được 50 gam kết tủa trắng.\na) Viết phương trình hóa học của phản ứng lên men glucose và phản ứng tạo kết tủa.\nb) Tính giá trị của m.",
            "sol": "a) Các phương trình hóa học:\n(1) C~6~H~12~O~6~ ──(men rượu, 30-35°C)──> 2C~2~H~5~OH + 2CO~2~↑\n(2) CO~2~ + Ca(OH)~2~ dư → CaCO~3~↓ + H~2~O\n\nb) Số mol kết tủa CaCO~3~: n_CaCO3 = 50 / 100 = 0,5 mol.\nTheo phương trình (2): n_CO2 = n_CaCO3 = 0,5 mol.\nTheo phương trình (1): n_glucose(phản ứng) = (1/2) × n_CO2 = 0,25 mol.\nKhối lượng glucose phản ứng trên lí thuyết:\nm_lí thuyết = 0,25 × 180 = 45 gam.\nDo hiệu suất lên men chỉ đạt H = 75%, khối lượng glucose m thực tế cần lấy là:\nm = m_lí thuyết / H% = 45 / 75% = 60 gam."
        },
        {
            "idx": 3,
            "topic": "Bài 8: Amine (Vận dụng)",
            "prompt": "Cho 9,0 gam một amine đơn chức no, mạch hở X tác dụng vừa đủ với dung dịch chứa 0,2 mol HCl.\na) Xác định công thức phân tử của amine X.\nb) Viết các công thức cấu tạo có thể có của X và gọi tên theo danh pháp thay thế.\nc) Cho chất X tác dụng với dung dịch nitrous acid (HNO~2~) ở nhiệt độ phòng thấy có bọt khí N~2~ thoát ra. Xác định chính xác công thức cấu tạo của X.",
            "sol": "a) Đặt công thức của amine đơn chức là RNH~2~ (hoặc R~1~R~2~NH, R~1~R~2~R~3~N).\nPhương trình phản ứng: Amine + HCl → Muối amonium chloride\nn_amine = n_HCl = 0,2 mol.\nKhối lượng mol của X: M_X = 9,0 / 0,2 = 45 g/mol.\nAmine no, đơn chức, mạch hở có công thức chung là C~n~H~2n+3~N:\n14n + 17 = 45 ⇒ 14n = 28 ⇒ n = 2.\nCông thức phân tử của X là C~2~H~7~N.\n\nb) Các công thức cấu tạo của C~2~H~7~N:\n- Bậc 1: CH~3~-CH~2~-NH~2~ (ethanamine).\n- Bậc 2: CH~3~-NH-CH~3~ (N-methylmethanamine).\n\nc) Amine tác dụng với HNO~2~ ở nhiệt độ thường giải phóng khí N~2~ là amine bậc một mạch hở:\nCH~3~CH~2~NH~2~ + HNO~2~ → CH~3~CH~2~OH + N~2~↑ + H~2~O.\nVậy công thức cấu tạo chính xác của X là CH~3~-CH~2~-NH~2~ (ethanamine)."
        }
    ]
}

# ==============================================================================
# DỮ LIỆU ĐỀ 02 (MÃ ĐỀ 102)
# ==============================================================================
EXAM_02 = {
    "de_num": 2,
    "made": "102",
    "title": "ĐỀ KIỂM TRA GIỮA HỌC KỲ I - ĐỀ SỐ 02",
    # PHẦN I: 12 CÂU
    "part1": [
        {
            "idx": 1,
            "topic": "Bài 1: Ester - Lipit (Biết)",
            "prompt": "Ester benzyl acetate có mùi thơm đặc trưng của hoa nhài. Công thức cấu tạo của benzyl acetate là",
            "opts": [
                "A. CH~3~COOCH~2~C~6~H~5~.",
                "B. C~6~H~5~COOCH~3~.",
                "C. CH~3~COOC~6~H~5~.",
                "D. C~6~H~5~CH~2~COOCH~3~."
            ],
            "correct": "A",
            "sol": "Benzyl acetate là ester tạo bởi acetic acid (CH~3~COOH) và benzyl alcohol (C~6~H~5~CH~2~OH), có công thức là CH~3~COOCH~2~C~6~H~5~, mùi hoa nhài."
        },
        {
            "idx": 2,
            "topic": "Bài 4: Carbohydrat. Glucose và fructose (Biết)",
            "prompt": "Carbohydrate nào sau đây có vị ngọt sắc nhất trong các loại đường tự nhiên và chiếm khoảng 40% khối lượng của mật ong?",
            "opts": [
                "A. Glucose.",
                "B. Fructose.",
                "C. Saccharose.",
                "D. Maltose."
            ],
            "correct": "B",
            "sol": "Fructose có độ ngọt cao nhất trong các loại đường tự nhiên (ngọt hơn saccharose và glucose), chiếm khoảng 40% trong mật ong."
        },
        {
            "idx": 3,
            "topic": "Bài 4: Carbohydrat. Glucose và fructose (Hiểu)",
            "prompt": "Mặc dù phân tử không chứa nhóm chức aldehyde (-CHO) nhưng fructose vẫn tham gia phản ứng tráng bạc với thuốc thử Tollens là do",
            "opts": [
                "A. trong môi trường kiềm, fructose chuyển hóa thuận nghịch thành glucose.",
                "B. fructose có nhóm chức ketone bị khử bởi ion Ag^+.",
                "C. fructose có nhiều nhóm hydroxy kề nhau phản ứng trực tiếp với Ag^+.",
                "D. phân tử fructose tự phân cắt thành hai phân tử aldehyde đơn giản."
            ],
            "correct": "A",
            "sol": "Trong môi trường kiềm của thuốc thử Tollens ([Ag(NH~3~)~2~]OH), fructose chuyển hóa thuận nghịch thành glucose nhờ cân bằng đồng phân hóa enediol, do đó fructose vẫn cho phản ứng tráng bạc."
        },
        {
            "idx": 4,
            "topic": "Bài 5: Saccharose và maltose (Biết)",
            "prompt": "Phân tử maltose (đường mạch nha) được cấu tạo từ hai gốc glucose liên kết với nhau qua liên kết",
            "opts": [
                "A. α-1,4-glycosidic.",
                "B. α-1,2-glycosidic.",
                "C. β-1,4-glycosidic.",
                "D. α-1,6-glycosidic."
            ],
            "correct": "A",
            "sol": "Maltose gồm hai gốc α-glucose liên kết với nhau qua liên kết α-1,4-glycosidic, còn một nhóm -OH hemiacetal tự do nên maltose có tính khử (tráng bạc được)."
        },
        {
            "idx": 5,
            "topic": "Bài 6: Tinh bột và cellulose (Biết)",
            "prompt": "Nhỏ vài giọt dung dịch iodine vào mặt cắt củ khoai lang tươi thấy xuất hiện màu xanh tím. Hiện tượng này là do iodine tạo phức hấp phụ với",
            "opts": [
                "A. cấu trúc xoắn của tinh bột.",
                "B. chuỗi mạch kéo dài của cellulose.",
                "C. các nhóm -OH của glucose.",
                "D. protein có trong khoai lang."
            ],
            "correct": "A",
            "sol": "Phân tử amylose trong tinh bột có cấu trúc xoắn dạng lò xo, các phân tử iodine lọt vào trong lòng ống xoắn tạo phức chất hấp phụ có màu xanh tím đặc trưng."
        },
        {
            "idx": 6,
            "topic": "Bài 6: Tinh bột và cellulose (Biết)",
            "prompt": "Đặc điểm cấu tạo nào sau đây của phân tử cellulose làm cho cellulose có độ bền cơ học cao và tạo nên sợi dẻo dai?",
            "opts": [
                "A. Các chuỗi mạch cellulose dài, không phân nhánh, xếp song song và liên kết chặt với nhau bằng liên kết hydrogen liên phân tử.",
                "B. Mạch polymer phân nhánh mạnh với các liên kết α-1,6-glycosidic.",
                "C. Cấu trúc cuộn tròn hình cầu nhờ liên kết peptide.",
                "D. Chứa nhiều nhóm aldehyde hoạt động liên kết đồng hóa trị với nhau."
            ],
            "correct": "A",
            "sol": "Các mạch cellulose không phân nhánh, duỗi thẳng, xếp song song tạo thành các vi sợi liên kết rất chặt chẽ với nhau bởi mạng lưới liên kết hydrogen liên phân tử."
        },
        {
            "idx": 7,
            "topic": "Bài 6: Tinh bột và cellulose (Biết)",
            "prompt": "Gỗ, tre, nứa và bông nõn là những nguyên liệu tự nhiên giàu chất nào sau đây dùng để sản xuất bột giấy và tơ sợi tổng hợp/nhân tạo?",
            "opts": [
                "A. Tinh bột.",
                "B. Saccharose.",
                "C. Cellulose.",
                "D. Fructose."
            ],
            "correct": "C",
            "sol": "Cellulose chiếm tới 95-98% trong bông nõn và khoảng 50% trong gỗ tre nứa, là nguyên liệu cốt lõi sản xuất giấy, tơ visco, tơ acetate."
        },
        {
            "idx": 8,
            "topic": "Bài 6: Tinh bột và cellulose (Hiểu)",
            "prompt": "Khi thủy phân hoàn toàn tinh bột hoặc cellulose trong môi trường acid vô cơ loãng, đun nóng, sản phẩm cuối cùng thu được duy nhất là",
            "opts": [
                "A. Fructose.",
                "B. Glucose.",
                "C. Saccharose.",
                "D. Maltose."
            ],
            "correct": "B",
            "sol": "Cả tinh bột (amylose, amylopectin) và cellulose đều là polysaccharide cấu tạo từ các mắt xích glucose, nên khi thủy phân hoàn toàn đều chỉ tạo ra monosaccharide duy nhất là glucose."
        },
        {
            "idx": 9,
            "topic": "Bài 8: Amine (Hiểu)",
            "prompt": "Khi nhỏ vài giọt dung dịch aniline (C~6~H~5~NH~2~) vào ống nghiệm chứa nước bromine, hiện tượng quan sát được là",
            "opts": [
                "A. nước bromine bị mất màu và xuất hiện kết tủa trắng.",
                "B. xuất hiện kết tủa màu vàng và sủi bọt khí.",
                "C. dung dịch chuyển sang màu xanh thẫm.",
                "D. không có hiện tượng gì do aniline không phản ứng."
            ],
            "correct": "A",
            "sol": "Nhóm -NH~2~ đẩy electron làm tăng mật độ electron ở vị trí ortho và para trên vòng benzene, làm mất màu nước bromine và sinh ra kết tủa trắng 2,4,6-tribromoaniline."
        },
        {
            "idx": 10,
            "topic": "Bài 9: Amino acid và peptide (Biết)",
            "prompt": "Tên gọi thay thế theo danh pháp IUPAC của alanine (CH~3~-CH(NH~2~)-COOH) là",
            "opts": [
                "A. 2-aminopropanoic acid.",
                "B. α-aminopropionic acid.",
                "C. 2-aminoethanoic acid.",
                "D. 3-aminopropanoic acid."
            ],
            "correct": "A",
            "sol": "Mạch chính có 3 carbon (propanoic acid), nhóm -NH~2~ ở vị trí carbon số 2 → Tên thay thế IUPAC là 2-aminopropanoic acid. (α-aminopropionic acid là tên bán hệ thống)."
        },
        {
            "idx": 11,
            "topic": "Bài 9: Amino acid và peptide (Hiểu)",
            "prompt": "Chất nào sau đây vừa có khả năng tác dụng với dung dịch HCl vừa có khả năng tác dụng với dung dịch NaOH (thể hiện tính chất lưỡng tính)?",
            "opts": [
                "A. Ethylamine (CH~3~CH~2~NH~2~).",
                "B. Acetic acid (CH~3~COOH).",
                "C. Glycine (H~2~N-CH~2~-COOH).",
                "D. Ethyl acetate (CH~3~COOC~2~H~5~)."
            ],
            "correct": "C",
            "sol": "Glycine chứa đồng thời nhóm basic (-NH~2~) tác dụng với acid mạnh (HCl) và nhóm acidic (-COOH) tác dụng với base mạnh (NaOH) → thể hiện tính chất lưỡng tính."
        },
        {
            "idx": 12,
            "topic": "Bài 10: Protein và enzyme (Biết)",
            "prompt": "Protein nào sau đây thuộc loại protein hình cầu, tan tốt trong nước tạo thành dung dịch keo nhớt?",
            "opts": [
                "A. Albumin (lòng trắng trứng).",
                "B. Keratin (protein ở tóc, móng).",
                "C. Collagen (protein ở da, sụn).",
                "D. Fibroin (protein ở tơ tằm)."
            ],
            "correct": "A",
            "sol": "Albumin (lòng trắng trứng) và hemoglobin là các protein hình cầu, tan được trong nước tạo dung dịch keo. Keratin, collagen, fibroin là protein hình sợi không tan trong nước."
        }
    ],

    # PHẦN II: 2 CÂU ĐÚNG - SAI
    "part2": [
        {
            "idx": 1,
            "topic": "Chủ đề 1: Ester - Lipid. Xà phòng và chất giặt rửa",
            "intro": "Cho các nhận định về chất béo, xà phòng và chất giặt rửa tổng hợp trong đời sống và sản xuất:",
            "items": [
                ("a) Xà phòng là hỗn hợp muối sodium hoặc potassium của các acid béo kết hợp với một số chất phụ gia.", "Đúng"),
                ("b) Ưu điểm nổi bật của chất giặt rửa tổng hợp so với xà phòng là có thể dùng để giặt rửa tốt trong môi trường nước cứng mà không tạo kết tủa cản trở.", "Đúng"),
                ("c) Phân tử của cả xà phòng và chất giặt rửa tổng hợp đều có cấu trúc gồm hai phần: phần ưa nước (mang điện hoặc phân cực) và phần kị nước (gốc hydrocarbon dài không phân cực).", "Đúng"),
                ("d) Hydrogen hóa hoàn toàn triolein (C~17~H~33~COO)~3~C~3~H~5~ xúc tác Ni, đun nóng sẽ chuyển chất béo lỏng thành chất béo rắn (tristearin), thuận tiện cho bảo quản và làm bơ nhân tạo.", "Đúng")
            ],
            "sol": "• a) ĐÚNG: Định nghĩa chuẩn về xà phòng: hỗn hợp muối Na/K của acid béo và chất độn, tạo hương.\n• b) ĐÚNG: Muối của chất giặt rửa tổng hợp với ion Ca^2+, Mg^2+ (trong nước cứng) dễ tan, không bị kết tủa như muối xà phòng (RCOO)~2~Ca/Mg.\n• c) ĐÚNG: Cơ chế tẩy rửa dựa vào cấu trúc lưỡng ái: đuôi kị nước hòa tan vết dầu mỡ, đầu ưa nước quay ra ngoài kéo hạt dầu vào nước.\n• d) ĐÚNG: Phản ứng hydrogen hóa làm no các liên kết C=C trong gốc acid béo chưa no, chuyển dầu lỏng thành mỡ rắn (bơ thực vật, margarin)."
        },
        {
            "idx": 2,
            "topic": "Chủ đề 3: Hợp chất chứa nitrogen",
            "intro": "Về cấu trúc và tính chất của các hợp chất chứa nitrogen trong chương trình Hóa học 12:",
            "items": [
                ("a) Aniline (C~6~H~5~NH~2~) là chất lỏng không màu, ít tan trong nước, nặng hơn nước và có tính base yếu hơn ammonia (không làm đổi màu quỳ tím).", "Đúng"),
                ("b) Trong cơ thể sinh vật, các amino acid tồn tại chủ yếu ở dạng ion lưỡng cực và tham gia cấu tạo nên protein là các α-amino acid thiên nhiên.", "Đúng"),
                ("c) Cho peptide mạch hở Gly-Ala phản ứng với lượng dư dung dịch NaOH đun nóng thì tỉ lệ mol giữa peptide và NaOH phản ứng là 1 : 2.", "Đúng"),
                ("d) Khi nhỏ cồn (ethanol 70°) hoặc đun nóng lòng trắng trứng thì protein bị đông tụ và biến tính, làm mất hoạt tính sinh học ban đầu.", "Đúng")
            ],
            "sol": "• a) ĐÚNG: Aniline là chất lỏng ít tan trong nước, nặng hơn nước, dung dịch không làm đổi màu quỳ tím do tính base rất yếu.\n• b) ĐÚNG: Đa số amino acid thiên nhiên là α-amino acid và tồn tại dạng ion lưỡng cực ở pH sinh lí.\n• c) ĐÚNG: Dipeptide Gly-Ala có 1 liên kết peptide và 1 nhóm -COOH đầu C: Gly-Ala + 2NaOH → Gly-Na + Ala-Na + H~2~O (tỉ lệ 1 : 2).\n• d) ĐÚNG: Nhiệt độ, dung môi hữu cơ (cồn), acid, kiềm đều phá vỡ cấu trúc bậc 2, 3, 4 của protein gây hiện tượng đông tụ và biến tính."
        }
    ],

    # PHẦN III: 4 CÂU TRẢ LỜI NGẮN
    "part3": [
        {
            "idx": 1,
            "topic": "Bài 5: Saccharose và maltose (Biết)",
            "prompt": "Khối lượng mol phân tử (amu) của disaccharide maltose (C~12~H~22~O~11~) là bao nhiêu?",
            "ans": "342",
            "sol": "Maltose có công thức phân tử C~12~H~22~O~11~. M = 12 × 12 + 22 × 1 + 11 × 16 = 144 + 22 + 176 = 342 g/mol (hoặc amu)."
        },
        {
            "idx": 2,
            "topic": "Bài 6: Tinh bột và cellulose (Hiểu)",
            "prompt": "Để sản xuất 460 kg ethylic alcohol (C~2~H~5~OH) bằng phương pháp lên men từ tinh bột với hiệu suất toàn bộ quá trình đạt 80%, cần dùng bao nhiêu kilogam tinh bột (chứa 10% tạp chất trơ)?",
            "ans": "1125",
            "sol": "Sơ đồ: (C~6~H~10~O~5~)~n~ → 2n C~2~H~5~OH\nn_C2H5OH = 460 / 46 = 10 kmol.\nTheo lí thuyết, số mol tinh bột cần dùng: n_tinh bột = 10 / 2 = 5 kmol.\nKhối lượng tinh bột nguyên chất lí thuyết: m_tinh bột = 5 × 162 = 810 kg.\nDo hiệu suất toàn bộ quá trình là 80%: m_tinh bột thực tế = 810 / 80% = 1012,5 kg.\nVì tinh bột chứa 10% tạp chất trơ (độ tinh khiết 90%): m_nguyên liệu = 1012,5 / 90% = 1125 kg."
        },
        {
            "idx": 3,
            "topic": "Bài 9: Amino acid và peptide (Hiểu)",
            "prompt": "Từ hai amino acid là glycine (Gly) và alanine (Ala), có thể tạo ra tối đa bao nhiêu đồng phân tripeptide mạch hở chứa đồng thời cả Gly và Ala trong phân tử?",
            "ans": "6",
            "sol": "Tripeptide tạo bởi Gly và Ala có dạng: X-Y-Z.\nTổng số tripeptide tạo từ 2 amino acid là 2^3 = 8 tripeptide.\nTrong đó có 2 tripeptide chỉ chứa 1 loại amino acid là: Gly-Gly-Gly và Ala-Ala-Ala.\nSố tripeptide chứa đồng thời cả Gly và Ala là: 8 - 2 = 6 (gồm: Gly-Gly-Ala, Gly-Ala-Gly, Ala-Gly-Gly, Ala-Ala-Gly, Ala-Gly-Ala, Gly-Ala-Ala)."
        },
        {
            "idx": 4,
            "topic": "Bài 10: Protein và enzyme (Biết)",
            "prompt": "Một phân tử pentapeptide mạch hở X có cấu tạo từ các gốc α-amino acid. Trong phân tử chất X có bao nhiêu liên kết peptide?",
            "ans": "4",
            "sol": "Pentapeptide là peptide được cấu tạo từ 5 gốc α-amino acid. Số liên kết peptide trong oligopeptide mạch hở bằng (số gốc amino acid - 1) = 5 - 1 = 4 liên kết peptide."
        }
    ],

    # PHẦN IV: 3 CÂU TỰ LUẬN
    "part4": [
        {
            "idx": 1,
            "topic": "Bài 1: Ester - Lipit (Vận dụng)",
            "prompt": "Đun sôi 80,6 gam một loại triglyceride X (chất béo) với dung dịch KOH vừa đủ, thu được 9,2 gam glycerol và hỗn hợp muối potassium của các acid béo dùng để sản xuất xà phòng mềm.\na) Viết phương trình hóa học tổng quát của phản ứng.\nb) Tính số mol KOH cần dùng cho phản ứng xà phòng hóa.\nc) Tính khối lượng hỗn hợp muối xà phòng thu được sau phản ứng.",
            "sol": "a) Phương trình hóa học tổng quát:\n(RCOO)~3~C~3~H~5~ + 3KOH → 3RCOOK + C~3~H~5~(OH)~3~\n\nb) Số mol glycerol thu được: n_glycerol = 9,2 / 92 = 0,1 mol.\nTheo phương trình phản ứng: n_KOH = 3 × n_glycerol = 3 × 0,1 = 0,3 mol.\n\nc) Áp dụng định luật bảo toàn khối lượng:\nm_chất béo + m_KOH = m_muối + m_glycerol\n⇒ m_muối = 80,6 + (0,3 × 56) - 9,2 = 80,6 + 16,8 - 9,2 = 88,2 gam.\nVậy khối lượng xà phòng mềm thu được là 88,2 gam."
        },
        {
            "idx": 2,
            "topic": "Bài 4: Glucose và fructose (Vận dụng)",
            "prompt": "Để tráng một lớp bạc mỏng lên bề mặt của 100 chiếc ruột phích nước, người ta đun nóng dung dịch chứa 36 gam glucose với lượng dư dung dịch AgNO~3~ trong NH~3~. Biết hiệu suất của phản ứng tráng bạc đạt 80% và toàn bộ lượng bạc sinh ra bám đều lên các ruột phích.\na) Viết phương trình hóa học của phản ứng tráng bạc của glucose.\nb) Tính tổng khối lượng bạc kim loại (Ag) sinh ra bám trên các ruột phích.\nc) Tính khối lượng bạc trung bình bám trên mỗi chiếc ruột phích (theo gam).",
            "sol": "a) Phương trình hóa học:\nCH~2~OH[CHOH]~4~CHO + 2[Ag(NH~3~)~2~]OH ──(t°)──> CH~2~OH[CHOH]~4~COONH~4~ + 2Ag↓ + 3NH~3~ + H~2~O\n\nb) Số mol glucose ban đầu: n_glucose = 36 / 180 = 0,2 mol.\nTheo phương trình: n_Ag(lí thuyết) = 2 × n_glucose = 0,4 mol.\nVì hiệu suất phản ứng đạt 80%: n_Ag(thực tế) = 0,4 × 80% = 0,32 mol.\nTổng khối lượng bạc kim loại bám trên ruột phích:\nm_Ag = 0,32 × 108 = 34,56 gam.\n\nc) Khối lượng bạc trung bình bám trên mỗi chiếc ruột phích:\nm_trung bình = 34,56 / 100 = 0,3456 gam."
        },
        {
            "idx": 3,
            "topic": "Bài 8: Amine (Vận dụng)",
            "prompt": "Cho 18,6 gam aniline (C~6~H~5~NH~2~) phản ứng hoàn toàn với lượng dư dung dịch nước bromine (Br~2~).\na) Nêu hiện tượng quan sát được và viết phương trình hóa học của phản ứng xảy ra.\nb) Tính khối lượng kết tủa trắng thu được sau phản ứng.\nc) Giải thích vì sao aniline dễ tham gia phản ứng thế nguyên tử H trên vòng benzene hơn benzene.",
            "sol": "a) Hiện tượng: Dung dịch nước bromine bị mất màu da cam và xuất hiện kết tủa trắng đục.\nPhương trình hóa học:\nC~6~H~5~NH~2~ + 3Br~2~ → C~6~H~2~Br~3~NH~2~↓ + 3HBr\n(Sản phẩm kết tủa là 2,4,6-tribromoaniline).\n\nb) Số mol aniline: n_aniline = 18,6 / 93 = 0,2 mol.\nTheo phương trình phản ứng: n_kết tủa = n_aniline = 0,2 mol.\nKhối lượng mol của 2,4,6-tribromoaniline (C~6~H~4~Br~3~N): M = 12×6 + 4×1 + 80×3 + 14 = 330 g/mol.\nKhối lượng kết tủa trắng thu được:\nm_kết tủa = 0,2 × 330 = 66,0 gam.\n\nc) Giải thích: Do đôi electron tự do trên nguyên tử nitrogen của nhóm -NH~2~ liên hợp với hệ electron π của vòng benzene (+C), làm tăng mật độ electron trong vòng, đặc biệt ở các vị trí ortho và para. Do đó, aniline phản ứng thế halogen dễ dàng hơn nhiều so với benzene."
        }
    ]
}

# ==============================================================================
# HÀM XUẤT FILE WORD CHO MỖI ĐỀ
# ==============================================================================
def render_exam_doc(exam, is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    de_num = exam["de_num"]
    made = exam["made"]

    # 1. Title Block
    build_title_block(
        doc,
        school_line="SỞ GD&ĐT THỪA THIÊN HUẾ\nTRƯỜNG THPT CHUYÊN QUỐC HỌC",
        exam_title="ĐỀ KIỂM TRA GIỮA HỌC KỲ I - NĂM HỌC 2026 - 2027",
        subject_line="MÔN: HOÁ HỌC - LỚP 12",
        num_pages="03",
        duration="50 phút (không kể thời gian phát đề)",
        exam_code=made,
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Br = 80; Ag = 108",
        is_teacher=is_teacher
    )

    # 2. PHẦN I
    add_section_header(
        doc,
        "PHẦN I. Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.",
        subtitle="Mỗi câu trả lời đúng thí sinh được 0,25 điểm (Tổng: 3,0 điểm)"
    )

    for q in exam["part1"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, f"Câu {q['idx']}. {q['prompt']}")

        formatted_opts = []
        for opt in q["opts"]:
            opt_letter = opt[:1]
            if is_teacher and opt_letter == q["correct"]:
                formatted_opts.append(f"*{opt}*")
            else:
                formatted_opts.append(opt)

        render_options_no_table(doc, formatted_opts, is_teacher=is_teacher)

        if is_teacher:
            sol_text = f"✔ Đáp án đúng: {q['correct']}.\n{q['sol']}"
            add_solution_no_table(doc, sol_text)

    # 3. PHẦN II
    add_section_header(
        doc,
        "PHẦN II. Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.",
        subtitle="Mỗi ý trả lời đúng được 0,25 điểm (Tối đa 1,0 điểm / câu, Tổng: 2,0 điểm)"
    )

    for q in exam["part2"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, f"Câu {q['idx']}. {q['intro']}")

        for stmt_text, ans_val in q["items"]:
            if not stmt_text:
                continue
            p_st = doc.add_paragraph()
            p_st.paragraph_format.left_indent = Cm(0.5)
            p_st.paragraph_format.space_before = Pt(1)
            p_st.paragraph_format.space_after = Pt(1)
            render_statement(p_st, stmt_text, ans=ans_val, is_teacher=is_teacher)

        if is_teacher:
            add_solution_no_table(doc, q["sol"])

    # 4. PHẦN III
    add_section_header(
        doc,
        "PHẦN III. Thí sinh trả lời từ câu 1 đến câu 4. Viết câu trả lời vào chỗ trống tương ứng.",
        subtitle="Mỗi câu trả lời đúng thí sinh được 0,5 điểm (Tổng: 2,0 điểm)"
    )

    for q in exam["part3"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, f"Câu {q['idx']}. {q['prompt']}")

        if not is_teacher:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.space_before = Pt(2)
            p_ans.paragraph_format.space_after = Pt(2)
            render_rich_text(p_ans, "*Trả lời:* ...................................................................................................................................................")
        else:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.space_before = Pt(2)
            p_ans.paragraph_format.space_after = Pt(2)
            render_rich_text(p_ans, f"*Trả lời:* *{q['ans']}*", base_color=COLOR_CORRECT)
            add_solution_no_table(doc, q["sol"])

    # 5. PHẦN IV - TỰ LUẬN
    add_section_header(
        doc,
        "PHẦN IV. TỰ LUẬN (3 câu – 3,0 điểm). Thí sinh trình bày chi tiết lời giải vào giấy làm bài.",
        subtitle="Mỗi câu đúng hoàn toàn được 1,0 điểm"
    )

    for q in exam["part4"]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, f"Câu {q['idx']} (1,0 điểm). {q['prompt']}")

        if not is_teacher:
            add_student_writing_lines(doc, num_lines=7)
        else:
            add_solution_no_table(doc, q["sol"])

    # Kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(8)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run(f"-------------------------- HẾT ĐỀ THI MÃ {made} --------------------------")
    set_run_style(r_end, font_size=10.5, bold=True)

    return doc

def main():
    print("=== ĐANG KHỞI TẠO 2 ĐỀ KIỂM TRA THEO MA TRẬN GIỮA KÌ I HÓA 12 ===")
    
    exams = [EXAM_01, EXAM_02]
    
    for ex in exams:
        num = ex["de_num"]
        made = ex["made"]
        print(f"-> Tạo Đề {num:02d} (Mã {made})...")

        # 1. Bản Học Sinh
        doc_hs = render_exam_doc(ex, is_teacher=False)
        hs_name = f"De_Kiem_Tra_GK1_Hoa_12_De{num:02d}_Ma{made}_HocSinh.docx"
        hs_path = os.path.join(OUT_DIR, hs_name)
        doc_hs.save(hs_path)
        print(f"   + Đã lưu Học Sinh: {hs_path}")

        # 2. Bản Giáo Viên
        doc_gv = render_exam_doc(ex, is_teacher=True)
        gv_name = f"De_Kiem_Tra_GK1_Hoa_12_De{num:02d}_Ma{made}_GiaoVien.docx"
        gv_path = os.path.join(OUT_DIR, gv_name)
        doc_gv.save(gv_path)
        print(f"   + Đã lưu Giáo Viên: {gv_path}")

    print("=== HOÀN TẤT TẠO 4 FILE WORD CHUẨN FORM ĐỀ KIỂM TRA! ===")

if __name__ == "__main__":
    main()
