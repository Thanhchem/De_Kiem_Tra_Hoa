# -*- coding: utf-8 -*-
"""
Script sinh Đề kiểm tra Giữa Kì 1 - Hóa học 12 - THPT Hai Bà Trưng Huế (Mã đề 121)
Tuân thủ 100% quy chuẩn FORM ĐỀ KIỂM TRA (RULE GEMINI.md & form_de_kiem_tra.py)
Sinh cả 2 phiên bản Word (.docx) và 2 phiên bản LaTeX (.tex / .pdf).
"""

import os
import sys
import subprocess
import docx
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Import các hàm chuẩn từ scripts/form_de_kiem_tra.py
sys.path.append(r"c:\Antigravity_Thanh\scripts")
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
    COLOR_SOLUTION,
    COLOR_PRIMARY
)

DATA_DE = {
    "truong": "SỞ GD&ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT HAI BÀ TRƯNG",
    "ky_thi": "KIỂM TRA GK1 NĂM HỌC 2026 - 2027",
    "mon": "MÔN: HOÁ HỌC - LỚP 12",
    "ma_de": "121",
    "thoi_gian": "45 phút (không kể thời gian phát đề)",
    "so_trang": "02",
    "nguyen_tu_khoi": "H = 1; C = 12; N = 14; O = 16; Na = 23; Cl = 35,5",
    
    # PHẦN 1
    "phan_1": [
        {
            "num": 1,
            "prompt": "Câu 1. Ester là hợp chất hữu cơ được tạo thành khi thay thế nhóm -OH ở nhóm carboxyl (-COOH) của carboxylic acid bằng nhóm nào sau đây?",
            "options": ["*A. -OR'.*", "B. -NH~2~.", "C. -CH~3~.", "D. -Cl."],
            "solution": "Theo khái niệm về ester: Khi thay thế nhóm -OH ở nhóm carboxyl (-COOH) của carboxylic acid bằng nhóm -OR' (với R' là gốc hydrocarbon) thì thu được ester.\n► *Chọn A.*"
        },
        {
            "num": 2,
            "prompt": "Câu 2. Trong thực tế, cá tươi thường có mùi tanh đặc trưng do hợp chất trimethylamine (CH~3~)~3~N gây ra. Khi sơ chế cá, người ta thường dùng giấm ăn hoặc chanh tươi chà xát lên thân cá để khử mùi tanh. Cơ sở hóa học của việc làm này là",
            "options": [
                "*A. acid trong chanh/giấm phản ứng với trimethylamine tạo thành muối không có mùi.*",
                "B. chanh/giấm bay hơi mang theo khí trimethylamine.",
                "C. chanh/giấm làm đông tụ protein của cá.",
                "D. chanh/giấm làm hòa tan hoàn toàn trimethylamine bám trên da cá."
            ],
            "solution": "Trimethylamine (CH~3~)~3~N là một amine có tính base yếu, dễ bay hơi gây mùi tanh khó chịu. Giấm ăn (chứa acetic acid CH~3~COOH) hoặc chanh (chứa citric acid) có tính acid sẽ phản ứng với amine tạo thành các muối ammonium tan tốt trong nước và không bay hơi, do đó khử sạch mùi tanh:\n(CH~3~)~3~N + CH~3~COOH -> (CH~3~)~3~NH^+^ + CH~3~COO^-^\n► *Chọn A.*"
        },
        {
            "num": 3,
            "prompt": "Câu 3. Phân tử maltose được hợp thành từ",
            "options": [
                "A. hai đơn vị beta-glucose.",
                "B. một đơn vị beta-glucose và một đơn vị alpha-fructose.",
                "*C. hai đơn vị alpha-glucose.*",
                "D. một đơn vị alpha-glucose và một đơn vị beta-fructose."
            ],
            "solution": "Phân tử maltose được hợp thành từ hai đơn vị alpha-glucose liên kết với nhau qua nguyên tử oxygen giữa C1 của gốc này với C4 của gốc kia (liên kết alpha-1,4-glycoside).\n► *Chọn C.*"
        },
        {
            "num": 4,
            "prompt": "Câu 4. Chất nào sau đây là một amine bậc 1?",
            "options": [
                "*A. C~2~H~5~NH~2~.*",
                "B. (CH~3~)~3~N.",
                "C. (CH~3~)~2~NH.",
                "D. CH~3~-NH-C~2~H~5~."
            ],
            "solution": "Bậc của amine bằng số nguyên tử H trong phân tử NH~3~ bị thay thế bởi gốc hydrocarbon:\n- C~2~H~5~NH~2~ (ethylamine): thay thế 1 nguyên tử H -> Amine bậc 1.\n- (CH~3~)~2~NH và CH~3~-NH-C~2~H~5~: Amine bậc 2.\n- (CH~3~)~3~N: Amine bậc 3.\n► *Chọn A.*"
        },
        {
            "num": 5,
            "prompt": "Câu 5. Carbohydrate nào sau đây thuộc loại disaccharide?",
            "options": [
                "A. Cellulose.",
                "B. Glucose.",
                "C. Fructose.",
                "*D. Saccharose.*"
            ],
            "solution": "Phân loại carbohydrate:\n- Monosaccharide: glucose, fructose (đường đơn).\n- Disaccharide: saccharose, maltose (đường đôi).\n- Polysaccharide: tinh bột, cellulose (đường đa).\n► *Chọn D.*"
        },
        {
            "num": 6,
            "prompt": "Câu 6. Ở điều kiện thường, các amino acid là những chất tinh thể không màu, dễ tan trong nước và có nhiệt độ nóng chảy cao là do trong phân tử tồn tại chủ yếu ở dạng",
            "options": [
                "A. phân tử trung hòa.",
                "*B. ion lưỡng cực.*",
                "C. anion.",
                "D. cation."
            ],
            "solution": "Phân tử amino acid có cả nhóm amino mang tính base (-NH~2~) và nhóm carboxyl mang tính acid (-COOH). Trong tinh thể cũng như trong dung dịch, nhóm -COOH chuyển proton H^+^ sang nhóm -NH~2~ tạo thành ion lưỡng cực H~3~N^+^-CHR-COO^-^. Lực hút tĩnh điện mạnh giữa các ion lưỡng cực làm cho amino acid là chất tinh thể, nhiệt độ nóng chảy cao và dễ tan trong nước.\n► *Chọn B.*"
        },
        {
            "num": 7,
            "prompt": "Câu 7. Xà phòng là hỗn hợp muối sodium hoặc potassium của acid béo cùng các chất phụ gia. Công thức nào sau đây đại diện cho thành phần chính của xà phòng?",
            "options": [
                "A. C~12~H~25~-C~6~H~4~-SO~3~Na.",
                "B. CH~3~[CH~2~]~10~CH~2~OSO~3~Na.",
                "*C. C~17~H~35~COONa.*",
                "D. C~3~H~5~(OH)~3~."
            ],
            "solution": "Thành phần chính của xà phòng là muối sodium hoặc potassium của các acid béo mạch dài (thường có từ 12 đến 18 C như stearic acid, palmitic acid, oleic acid,...). Do đó, C~17~H~35~COONa (sodium stearate) đại diện cho thành phần chính của xà phòng. Các chất ở A và B là chất giặt rửa tổng hợp; D là glycerol.\n► *Chọn C.*"
        },
        {
            "num": 8,
            "prompt": "Câu 8. Chất nào sau đây không có khả năng tham gia phản ứng thủy phân trong môi trường acid?",
            "options": [
                "A. Tinh bột.",
                "B. Cellulose.",
                "*C. Glucose.*",
                "D. Saccharose."
            ],
            "solution": "Glucose là monosaccharide, là đơn vị cấu trúc cơ bản nhất của carbohydrate nên không bị thủy phân tiếp. Tinh bột, cellulose và saccharose đều bị thủy phân trong dung dịch acid đun nóng tạo monosaccharide.\n► *Chọn C.*"
        },
        {
            "num": 9,
            "prompt": "Câu 9. Cho các dung dịch riêng biệt sau: glucose, saccharose, ethyl alcohol. Thuốc thử duy nhất có thể dùng để phân biệt cả 3 dung dịch trên ở nhiệt độ phòng và khi đun nóng là",
            "options": [
                "A. Nước bromine.",
                "*B. Cu(OH)~2~ trong môi trường kiềm.*",
                "C. Quỳ tím.",
                "D. Dung dịch AgNO~3~/NH~3~."
            ],
            "solution": "Sử dụng Cu(OH)~2~ trong môi trường kiềm:\n- Ở nhiệt độ phòng:\n  + Ethyl alcohol không phản ứng, kết tủa Cu(OH)~2~ không tan.\n  + Glucose và saccharose là các polyalcohol có các nhóm -OH kề nhau nên hòa tan Cu(OH)~2~ tạo dung dịch phức chất màu xanh lam thẫm.\n- Đun nóng 2 dung dịch màu xanh lam:\n  + Glucose có nhóm aldehyde (-CH=O) bị oxi hóa tạo kết tủa Cu~2~O màu đỏ gạch.\n  + Saccharose không có nhóm aldehyde tự do nên không có kết tủa đỏ gạch, giữ nguyên màu xanh lam.\nNhận biết hoàn toàn cả 3 chất.\n► *Chọn B.*"
        },
        {
            "num": 10,
            "prompt": "Câu 10. Nguyên nhân chính khiến dung dịch methylamine (CH~3~NH~2~) làm quỳ tím hóa xanh là do",
            "options": [
                "A. CH~3~NH~2~ bị phân hủy tạo ra NaOH.",
                "B. CH~3~NH~2~ có tính acid mạnh.",
                "C. phân tử CH~3~NH~2~ chứa nhóm -OH.",
                "*D. nguyên tử N còn cặp electron chưa liên kết có khả năng nhận H^+^ từ nước.*"
            ],
            "solution": "Nguyên tử nitrogen trong nhóm amine còn một cặp electron chưa liên kết, có khả năng nhận ion H^+^ từ phân tử nước theo phản ứng thuận nghịch:\nCH~3~NH~2~ + H~2~O <=> CH~3~NH~3~^+^ + OH^-^\nSự tạo thành ion OH^-^ làm tăng nồng độ hydroxide trong dung dịch ([OH^-^] > [H^+^]), khiến môi trường có tính base và làm quỳ tím hóa xanh.\n► *Chọn D.*"
        },
        {
            "num": 11,
            "prompt": "Câu 11. Số nhóm amino (-NH~2~) và nhóm carboxyl (-COOH) có trong phân tử Glycine lần lượt là",
            "options": [
                "A. 1 và 2.",
                "*B. 1 và 1.*",
                "C. 2 và 2.",
                "D. 2 và 1."
            ],
            "solution": "Glycine là amino acid đơn giản nhất, có công thức cấu tạo là H~2~N-CH~2~-COOH. Phân tử glycine chứa đúng 1 nhóm amino (-NH~2~) và 1 nhóm carboxyl (-COOH).\n► *Chọn B.*"
        },
        {
            "num": 12,
            "prompt": "Câu 12. Trong điều kiện thường, chất nào sau đây ở trạng thái khí và có mùi khai/tanh?",
            "options": [
                "A. Glutamic acid.",
                "B. Glycine.",
                "*C. Methylamine.*",
                "D. Aniline."
            ],
            "solution": "Methylamine (CH~3~NH~2~) ở điều kiện thường là chất khí không màu, có mùi khai/tanh nồng tương tự ammonia, độc và tan rất nhiều trong nước. Glutamic acid và glycine là chất rắn kết tinh; aniline là chất lỏng không màu, ít tan trong nước.\n► *Chọn C.*"
        }
    ],

    # PHẦN 2
    "phan_2": [
        {
            "num": 1,
            "prompt": "Câu 1. Khi để lâu ngoài không khí, dầu mỡ động thực vật thường xảy ra hiện tượng bị ôi và có mùi hôi khó chịu. Nguyên nhân chính gây ra hiện tượng này là do",
            "statements": [
                ("a) Phản ứng thủy phân chất béo tạo ra glycerol có vị đắng.", "Sai"),
                ("b) Các gốc acid béo không no bị oxi hóa chậm bởi oxygen không khí.", "Đúng"),
                ("c) Đã xảy ra phản ứng hydrogen hóa vào gốc acid béo không no.", "Sai"),
                ("d) Sự nhân lên của vi khuẩn phân hủy gốc acid béo no.", "Sai")
            ],
            "solution": "- a) Sai: Hiện tượng ôi mỡ chủ yếu do quá trình tự oxi hóa liên kết C=C trong gốc acid béo không no. Phản ứng thủy phân chất béo sinh ra glycerol (có vị ngọt, không đắng) và acid béo tự do, không phải nguyên nhân tạo mùi hôi ôi thiu đặc trưng.\n- b) Đúng: Gốc acid béo không no có nối đôi C=C bị oxi hóa chậm bởi oxygen không khí tạo hợp chất hydroperoxide, sau đó bị phân cắt thành các aldehyde, ketone có mùi hôi ôi khó chịu.\n- c) Sai: Phản ứng hydrogen hóa là phản ứng no hóa chất béo để bảo quản chất béo tốt hơn (tránh bị ôi mỡ), cần có xúc tác Ni và nhiệt độ, không tự xảy ra ngoài không khí.\n- d) Sai: Vi khuẩn không phân hủy chọn lọc gốc acid béo no tạo mùi ôi thiu; nguyên nhân chính của ôi mỡ là do phản ứng tự oxi hóa hóa học của liên kết C=C không no bởi oxygen."
        },
        {
            "num": 2,
            "prompt": "Câu 2. Tinh bột và cellulose là hai polysaccharide phổ biến trong tự nhiên, có cấu tạo và ứng dụng rất khác nhau. Xét các phát biểu sau:",
            "statements": [
                ("a) Phân tử amylose trong tinh bột được cấu tạo từ nhiều đơn vị alpha-glucose liên kết với nhau qua liên kết alpha-1,2-glycoside tạo thành chuỗi xoắn.", "Sai"),
                ("b) Con người không thể sử dụng cellulose (như rơm, gỗ) làm thực phẩm chính cung cấp năng lượng vì trong hệ tiêu hóa của con người không có enzyme cellulase xúc tác thủy phân cellulose.", "Đúng"),
                ("c) Phản ứng màu giữa hồ tinh bột và dung dịch iodine xuất hiện màu xanh tím do phân tử iodine đã phản ứng với các phân tử amylose có màu xanh tím. Màu sẽ biến mất khi đun nóng, để nguội lại xuất hiện.", "Sai"),
                ("d) Từ 1,0 tấn mùn cưa chứa 50% cellulose về khối lượng, qua quá trình thủy phân và lên men thu được ethanol với hiệu suất toàn bộ quá trình là 80%, khối lượng ethanol thu được là 237,2 kg.", "Sai")
            ],
            "solution": "- a) Sai: Amylose được cấu tạo từ các đơn vị alpha-glucose liên kết với nhau qua liên kết alpha-1,4-glycoside (không phải alpha-1,2).\n- b) Đúng: Hệ tiêu hóa của con người không tiết ra enzyme cellulase nên không thể thủy phân được liên kết beta-1,4-glycoside trong cellulose thành glucose để cung cấp năng lượng.\n- c) Sai: Hiện tượng xuất hiện màu xanh tím là tương tác vật lý hấp phụ (phân tử iodine chui vào khoang xoắn của amylose tạo hợp chất bao màu xanh tím), không phải phản ứng hóa học tạo chất mới và bản thân phân tử amylose không có màu xanh tím.\n- d) Sai: Tính toán khối lượng ethanol:\n  + m(cellulose) = 1,0 tấn x 50% = 500 kg.\n  + Sơ đồ: (C~6~H~10~O~5~)~n~ -> n C~6~H~12~O~6~ -> 2n C~2~H~5~OH\n  + 162 kg cellulose lý thuyết cho 2 x 46 = 92 kg ethanol.\n  + Khối lượng ethanol lý thuyết: m_lt = 500 x 92 / 162 ≈ 283,95 kg.\n  + Với hiệu suất H = 80%: m_tt = 283,95 x 80% ≈ 227,16 kg (khác 237,2 kg)."
        }
    ],

    # PHẦN 3
    "phan_3": [
        {
            "num": 1,
            "prompt": "Câu 1. Một học sinh lớp 12 cần cung cấp nhu cầu năng lượng mỗi ngày là 9690 kJ, trong đó chất béo đóng góp 20% tổng năng lượng này. Biết rằng trung bình mỗi gram chất béo khi chuyển hóa trong cơ thể cung cấp 38 kJ năng lượng. Khối lượng chất béo mà học sinh này cần tiêu thụ trong một ngày để đảm bảo dinh dưỡng là bao nhiêu gam?",
            "ans": "51",
            "solution": "Nhu cầu năng lượng từ chất béo mà học sinh cần trong 1 ngày là:\nE(chất béo) = 9690 kJ x 20% = 1938 kJ.\nKhối lượng chất béo cần tiêu thụ là:\nm = 1938 / 38 = 51 gam.\n► *Đáp số:* 51"
        },
        {
            "num": 2,
            "prompt": "Câu 2. Tính phân tử khối (theo amu) của tripeptide Ala-Gly-Val. (Biết Alanine = 89; Glycine = 75; Valine = 117).",
            "ans": "245",
            "solution": "Tripeptide Ala-Gly-Val được cấu tạo từ 1 phân tử Alanine, 1 phân tử Glycine và 1 phân tử Valine, loại đi 2 phân tử H~2~O:\nM(tripeptide) = M(Ala) + M(Gly) + M(Val) - 2 x M(H~2~O)\nM(tripeptide) = 89 + 75 + 117 - 2 x 18 = 245 amu.\n► *Đáp số:* 245"
        },
        {
            "num": 3,
            "prompt": "Câu 3. Cho 8,9 gam Alanine tác dụng hoàn toàn với dung dịch NaOH dư. Sau phản ứng thu được m gam muối khan. Giá trị của m là bao nhiêu? (Cho H = 1, C = 12, N = 14, O = 16, Na = 23, Cl = 35,5).",
            "ans": "11,1",
            "solution": "Số mol Alanine: n(Ala) = 8,9 / 89 = 0,1 mol.\nPhương trình hóa học:\nH~2~N-CH(CH~3~)-COOH + NaOH -> H~2~N-CH(CH~3~)-COONa + H~2~O\n0,1 mol ----------------------------------> 0,1 mol\nMuối thu được là sodium alaninate H~2~N-CH(CH~3~)-COONa có phân tử khối:\nM = 89 - 1 + 23 = 111 g/mol.\nKhối lượng muối khan m = 0,1 x 111 = 11,1 gam.\n► *Đáp số:* 11,1"
        },
        {
            "num": 4,
            "prompt": "Câu 4. Cho các chất/dung dịch sau: methylamine, aniline, glycine, glutamic acid, lysine. Có bao nhiêu dung dịch chất làm quỳ tím chuyển sang màu đỏ?",
            "ans": "1",
            "solution": "Xét môi trường dung dịch của từng chất đối với quỳ tím:\n- Methylamine (CH~3~NH~2~): có tính base mạnh hơn NH~3~ -> làm quỳ tím hóa xanh.\n- Aniline (C~6~H~5~NH~2~): có tính base rất yếu không làm đổi màu quỳ tím.\n- Glycine (H~2~N-CH~2~-COOH): có 1 nhóm -NH~2~ và 1 nhóm -COOH -> môi trường trung tính, không đổi màu quỳ tím.\n- Glutamic acid (HOOC-CH~2~-CH~2~-CH(NH~2~)-COOH): có 1 nhóm -NH~2~ và 2 nhóm -COOH -> dung dịch có tính acid, làm quỳ tím chuyển sang màu đỏ (hồng).\n- Lysine (H~2~N-(CH~2~)~4~-CH(NH~2~)-COOH): có 2 nhóm -NH~2~ và 1 nhóm -COOH -> dung dịch có tính base, làm quỳ tím hóa xanh.\nVậy chỉ có 1 dung dịch làm quỳ tím hóa đỏ (glutamic acid).\n► *Đáp số:* 1"
        }
    ],

    # PHẦN 4
    "phan_4": [
        {
            "num": 1,
            "prompt": "Câu 1. Tiến hành phản ứng ester hóa điều chế dầu chuối (isoamyl acetate) bằng cách đun nóng hỗn hợp gồm 16,2 g acetic acid (CH~3~COOH) và 15,2 g isoamyl alcohol ((CH~3~)~2~CHCH~2~CH~2~OH) với xúc tác H~2~SO~4~ đặc. Sau phản ứng thu được 14,16 g isoamyl acetate. Tính hiệu suất của phản ứng điều chế trên.",
            "solution": "1. Tính số mol các chất ban đầu:\n- n(CH~3~COOH) = 16,2 / 60 = 0,27 mol.\n- M(isoamyl alcohol, C~5~H~12~O) = 12 x 5 + 12 + 16 = 88 g/mol -> n(isoamyl alcohol) = 15,2 / 88 = 19/110 ≈ 0,1727 mol.\n2. Phương trình phản ứng ester hóa:\nCH~3~COOH + (CH~3~)~2~CHCH~2~CH~2~OH <=H~2~SO~4~ đ, t°=> CH~3~COOCH~2~CH~2~CH(CH~3~)~2~ + H~2~O\n3. So sánh tỉ lệ phản ứng:\nTheo phương trình phản ứng với tỉ lệ mol 1 : 1, do n(isoamyl alcohol) = 0,1727 mol < n(acid) = 0,27 mol nên hiệu suất phản ứng được tính theo lượng isoamyl alcohol (chất thiếu).\n4. Tính lượng ester theo lý thuyết:\nM(isoamyl acetate, C~7~H~14~O~2~) = 130 g/mol.\nn(ester lý thuyết) = n(isoamyl alcohol) = 19/110 mol.\nm(ester lý thuyết) = (19/110) x 130 = 247/11 ≈ 22,4545 gam.\n5. Hiệu suất của phản ứng điều chế:\nH = (m(ester thực tế) / m(ester lý thuyết)) x 100%\nH = (14,16 / 22,4545) x 100% ≈ 63,06% (hay xấp xỉ 63,1%)."
        },
        {
            "num": 2,
            "prompt": "Câu 2.\na) Trình bày cấu tạo của phân tử saccharose, biểu diễn công thức cấu tạo dạng mạch vòng của phân tử đó (ghi rõ số chỉ vị trí carbon mỗi gốc monosaccharide).\nb) Phân tử saccharose có thể mở vòng không? Giải thích.",
            "image": r"c:\Antigravity_Thanh\San_Pham\images_crop_de06\p2_cau2_saccharose.png",
            "solution": "a) Cấu tạo phân tử saccharose:\n- Saccharose là một disaccharide có công thức phân tử C~12~H~22~O~11~, được cấu tạo từ một gốc alpha-glucose và một gốc beta-fructose liên kết với nhau qua nguyên tử oxygen giữa C1 của gốc alpha-glucose và C2 của gốc beta-fructose (liên kết alpha-1,2-glycoside).\n- Công thức cấu tạo mạch vòng (công thức phối cảnh Haworth):\n  + Gốc alpha-glucose tồn tại ở dạng vòng pyranose (vòng 6 cạnh), carbon số 1 (C1) tham gia liên kết glycoside; các carbon còn lại đánh số tuần tự C2, C3, C4, C5 trên vòng và C6 ở nhóm -CH~2~OH.\n  + Gốc beta-fructose tồn tại ở dạng vòng furanose (vòng 5 cạnh), carbon số 2 (C2) tham gia liên kết glycoside; nhóm C1H~2~OH ở vị trí C2 ngoài vòng, các carbon C3, C4, C5 trên vòng và C6 ở nhóm -CH~2~OH.\nb) Khả năng mở vòng:\n- Phân tử saccharose KHÔNG THỂ mở vòng được.\n- Giải thích: Vì nhóm -OH hemiacetal ở carbon C1 của gốc alpha-glucose và nhóm -OH hemiketal ở carbon C2 của gốc beta-fructose đều đã tham gia tạo liên kết alpha-1,2-glycoside. Do đó, trong phân tử saccharose không còn nhóm -OH hemiacetal hoặc hemiketal tự do nào để có thể mở vòng tái tạo nhóm chức aldehyde (-CH=O) hay ketone (>C=O). Phân tử chỉ tồn tại cố định ở dạng mạch vòng khép kín."
        },
        {
            "num": 3,
            "prompt": "Câu 3. Thủy phân hết m gam tetrapeptide Ala-Ala-Ala-Ala thu được hỗn hợp gồm 44,5 gam Ala, 24 gam Ala-Ala và 46,2 gam Ala-Ala-Ala. Xác định giá trị m.",
            "solution": "1. Khối lượng mol các peptide và amino acid:\n- Ala (C~3~H~7~O~2~N): M = 89 g/mol.\n- Ala-Ala (Ala~2~): M = 89 x 2 - 18 = 160 g/mol.\n- Ala-Ala-Ala (Ala~3~): M = 89 x 3 - 36 = 231 g/mol.\n- Tetrapeptide Ala~4~: M = 89 x 4 - 54 = 302 g/mol.\n2. Tính số mol các sản phẩm thu được:\n- n(Ala) = 44,5 / 89 = 0,5 mol.\n- n(Ala-Ala) = 24 / 160 = 0,15 mol.\n- n(Ala-Ala-Ala) = 46,2 / 231 = 0,2 mol.\n3. Áp dụng định luật bảo toàn gốc (mắt xích) Ala:\nTổng số mol mắt xích Ala có trong hỗn hợp sản phẩm:\nn(Ala mắt xích) = 1 x n(Ala) + 2 x n(Ala-Ala) + 3 x n(Ala-Ala-Ala)\nn(Ala mắt xích) = 1 x 0,5 + 2 x 0,15 + 3 x 0,2 = 0,5 + 0,3 + 0,6 = 1,4 mol.\n4. Số mol tetrapeptide Ala~4~ ban đầu:\nDo mỗi phân tử Ala~4~ chứa 4 mắt xích Ala nên:\nn(Ala~4~) = 1,4 / 4 = 0,35 mol.\n5. Khối lượng tetrapeptide m ban đầu:\nm = n(Ala~4~) x M(Ala~4~) = 0,35 x 302 = 105,7 gam.\n► *Đáp số:* 105,7 gam"
        }
    ]
}

def export_docx(is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)
    
    # Title Block
    build_title_block(
        doc,
        school_line="SỞ GD&ĐT THÀNH PHỐ HUẾ\nTRƯỜNG THPT HAI BÀ TRƯNG",
        exam_title="KIỂM TRA GK1 NĂM HỌC 2026 - 2027",
        subject_line="MÔN: HOÁ HỌC - LỚP 12",
        num_pages="02",
        duration="45 phút",
        exam_code="121",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Cl = 35,5",
        is_teacher=is_teacher
    )

    # PHẦN 1
    add_section_header(doc, "PHẦN 1. Trắc nghiệm nhiều lựa chọn", "Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.")
    for q in DATA_DE["phan_1"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(3)
        p_q.paragraph_format.space_after = Pt(1)
        render_question_prompt(p_q, q["prompt"], base_size=10.5)
        
        # Render 4 options
        render_options_no_table(doc, q["options"], is_teacher=is_teacher)
        
        if is_teacher and "solution" in q:
            add_solution_no_table(doc, q["solution"])

    # PHẦN 2
    add_section_header(doc, "PHẦN 2. Trắc nghiệm đúng-sai", "Thí sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.")
    for q in DATA_DE["phan_2"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"], base_size=10.5)
        
        for stmt_txt, ans in q["statements"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt_txt, ans=ans, is_teacher=is_teacher, base_size=10.5)
            
        if is_teacher and "solution" in q:
            add_solution_no_table(doc, q["solution"])

    # PHẦN 3
    add_section_header(doc, "PHẦN 3. Trắc nghiệm trả lời ngắn", "Thí sinh trả lời từ câu 1 đến câu 4. Viết câu trả lời vào ô tương ứng.")
    for q in DATA_DE["phan_3"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"], base_size=10.5)
        
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(2)
        if is_teacher:
            render_rich_text(p_ans, f"*Trả lời:* *{q['ans']}*", base_size=10.5, base_color=COLOR_CORRECT)
        else:
            render_rich_text(p_ans, "*Trả lời:* ............................................................", base_size=10.5)

        if is_teacher and "solution" in q:
            add_solution_no_table(doc, q["solution"])

    # PHẦN 4
    add_section_header(doc, "PHẦN 4. Tự luận", "Thí sinh trình bày chi tiết lời giải từ câu 1 đến câu 3.")
    for q in DATA_DE["phan_4"]:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(5)
        p_q.paragraph_format.space_after = Pt(2)
        render_question_prompt(p_q, q["prompt"], base_size=10.5)
        
        if is_teacher and "image" in q and os.path.exists(q["image"]):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(2)
            r_img = p_img.add_run()
            r_img.add_picture(q["image"], width=Inches(4.2))

        if is_teacher and "solution" in q:
            add_solution_no_table(doc, q["solution"])
        else:
            add_student_writing_lines(doc, num_lines=5)

    # Dòng kết thúc
    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(10)
    render_rich_text(p_end, "---------- *HẾT* ----------", base_size=11, base_bold=True)

    suffix = "_GiaoVien.docx" if is_teacher else "_HocSinh.docx"
    out_path = os.path.join(r"c:\Antigravity_Thanh\San_Pham", f"De_Kiem_Tra_GK1_Hoa_12_HaiBaTrung_Ma121{suffix}")
    doc.save(out_path)
    print(f"Đã lưu thành công: {out_path}")
    return out_path

if __name__ == "__main__":
    import sys
    if sys.platform == "win32":
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    os.makedirs(r"c:\Antigravity_Thanh\San_Pham", exist_ok=True)
    f_hs = export_docx(is_teacher=False)
    f_gv = export_docx(is_teacher=True)
