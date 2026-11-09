# -*- coding: utf-8 -*-
"""
Script sinh file Word (.docx) Đề Kiểm Tra Giữa Kỳ I - Hóa Học 11 - Mã Đề 1202
Trường THPT Số 2 Quang Trung - Gia Lai.
Chuẩn Form Đề Kiểm Tra:
- 2 phiên bản: Bản Học Sinh (_HocSinh.docx) và Bản Giáo Viên (_GiaoVien.docx)
- Tuyệt đối KHÔNG DÙNG TABLE trong nội dung câu hỏi, phương án, lời giải.
- Lề trang, Header, Footer màu Coral Red (#D32F2F) và viền Sage Green (#8BC390).
- Đầy đủ lời giải chi tiết, công thức hóa học chuẩn subscript/superscript.
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# Import các hàm chuẩn hóa từ module form_de_kiem_tra
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from scripts.form_de_kiem_tra import (
    setup_header_footer, build_title_block, add_section_header,
    render_question_prompt, render_options_no_table, render_statement,
    add_solution_no_table, add_student_writing_lines, render_rich_text,
    set_run_style, COLOR_SOLUTION, COLOR_CORRECT
)

def build_exam_document(output_path, is_teacher=False):
    doc = docx.Document()
    
    # Cài đặt trang A4, lề 1.5 cm, Header & Footer
    setup_header_footer(doc)
    
    # Khối tiêu đề
    build_title_block(
        doc,
        school_line="SỞ GD & ĐT GIA LAI\nTRƯỜNG THPT SỐ 2 QUANG TRUNG",
        exam_title="ĐỀ KIỂM TRA GIỮA KỲ I, NĂM HỌC 2024 - 2025",
        subject_line="MÔN: HOÁ HỌC; KHỐI 11",
        num_pages="04",
        duration="45 phút",
        exam_code="1202",
        atomic_masses="C = 12; H = 1; O = 16; S = 32; Fe = 56",
        is_teacher=is_teacher
    )
    
    # -------------------------------------------------------------
    # PHẦN I. TRẮC NGHIỆM NHIỀU LỰA CHỌN (4,0 ĐIỂM)
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (4,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 16. Mỗi câu hỏi học sinh chỉ chọn một phương án đúng."
    )
    
    questions_p1 = [
        {
            "id": 1,
            "prompt": "Câu 1: Cho phản ứng hoá học: CO~2~(g) + H~2~(g) ⇌ CO(g) + H~2~O(g). Biểu thức hằng số cân bằng (K~C~) của phản ứng trên là",
            "options": [
                "A. K~C~ = [H~2~][CO] / [CO~2~][H~2~O]",
                "*B.* K~C~ = [H~2~O][CO] / [CO~2~][H~2~]",
                "C. K~C~ = [CO~2~][CO] / [H~2~O][H~2~]",
                "D. K~C~ = [CO~2~][H~2~] / [CO][H~2~O]"
            ],
            "solution": "Theo định luật tác dụng khối lượng, đối với phản ứng thuận nghịch: aA + bB ⇌ cC + dD.\nBiểu thức hằng số cân bằng nồng độ: K~C~ = ([C]^c^ × [D]^d^) / ([A]^a^ × [B]^b^).\nVới phản ứng: CO~2~(g) + H~2~(g) ⇌ CO(g) + H~2~O(g), tất cả các chất đều ở thể khí, ta có:\nK~C~ = ([CO] × [H~2~O]) / ([CO~2~] × [H~2~]).\n➔ Chọn đáp án B."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Phản ứng thuận nghịch là phản ứng:",
            "options": [
                "*A.* xảy ra theo hai chiều ngược nhau trong cùng điều kiện.",
                "B. xảy ra theo một chiều duy nhất trong điều kiện xác định.",
                "C. có phương trình hoá học được biểu diễn bằng mũi tên một chiều.",
                "D. xảy ra theo hai chiều giống nhau trong cùng điều kiện xác định."
            ],
            "solution": "Theo định nghĩa trong SGK Hóa học 11 (GDPT 2018):\nPhản ứng thuận nghịch là phản ứng trong đó ở cùng điều kiện, sự chuyển hóa chất đầu thành chất sản phẩm đồng thời diễn ra với sự chuyển hóa chất sản phẩm thành chất đầu (xảy ra theo hai chiều ngược nhau trong cùng điều kiện).\n➔ Chọn đáp án A."
        },
        {
            "id": 3,
            "prompt": "Câu 3: Sulfur là nguyên tố phổ biến thứ 17 trên vỏ Trái Đất, tồn tại ở bốn dạng đồng vị bền. Đồng vị chiếm % về khối lượng cao nhất là:",
            "options": [
                "*A.* ^32_16^S.",
                "B. ^34_16^S.",
                "C. ^36_16^S.",
                "D. ^33_16^S."
            ],
            "solution": "Trong tự nhiên, nguyên tố sulfur có 4 đồng vị bền với hàm lượng phần trăm tương ứng:\n- ^32_16^S: chiếm khoảng 94,98% (gần 95%).\n- ^34_16^S: chiếm khoảng 4,21%.\n- ^33_16^S: chiếm khoảng 0,75%.\n- ^36_16^S: chiếm khoảng 0,02%.\nDo đó, đồng vị ^32_16^S chiếm tỉ lệ phần trăm cao nhất.\n➔ Chọn đáp án A."
        },
        {
            "id": 4,
            "prompt": "Câu 4: Biểu thức tính pH đúng là",
            "options": [
                "A. pH = +10 lg[H^+].",
                "B. pH = -lg[OH^-].",
                "C. pH = lg[H^+].",
                "*D.* pH = -lg[H^+]."
            ],
            "solution": "Theo định nghĩa của nhà hóa học Sørensen:\npH = -lg[H^+] (hoặc [H^+] = 10^-pH^ M).\n➔ Chọn đáp án D."
        },
        {
            "id": 5,
            "prompt": "Câu 5: Cân bằng hoá học nào sau đây bị chuyển dịch khi thay đổi áp suất?",
            "options": [
                "A. 3Fe(s) + 4H~2~O(g) ⇌ Fe~3~O~4~(s) + 4H~2~(g)",
                "*B.* C(s) + H~2~O(g) ⇌ CO(g) + H~2~(g)",
                "C. C(s) + O~2~(g) ⇌ CO~2~(g)",
                "D. H~2~(g) + I~2~(g) ⇌ 2HI(g)"
            ],
            "solution": "Sự thay đổi áp suất chỉ làm chuyển dịch cân bằng đối với các hệ có chất khí và có sự biến thiên tổng số mol khí trước và sau phản ứng (Δn~khí~ ≠ 0):\n- Phản ứng A: n~khí trước~ = 4 (H~2~O) = n~khí sau~ = 4 (H~2~) ➔ Δn = 0 (Fe và Fe~3~O~4~ là chất rắn) ➔ Áp suất không ảnh hưởng.\n- Phản ứng B: vế trái có 1 mol khí (H~2~O), vế phải có 2 mol khí (1 CO + 1 H~2~) ➔ Δn = 2 - 1 = +1 ≠ 0 ➔ Cân bằng bị chuyển dịch khi thay đổi áp suất.\n- Phản ứng C: vế trái có 1 mol O~2~, vế phải có 1 mol CO~2~ ➔ Δn = 0 ➔ Áp suất không ảnh hưởng.\n- Phản ứng D: vế trái 2 mol khí (1 H~2~ + 1 I~2~), vế phải 2 mol HI ➔ Δn = 0 ➔ Áp suất không ảnh hưởng.\n➔ Chọn đáp án B."
        },
        {
            "id": 6,
            "prompt": "Câu 6: Trong tự nhiên, nguyên tố sulfur",
            "options": [
                "A. chỉ tồn tại ở dạng hợp chất.",
                "*B.* tồn tại ở cả dạng đơn chất và dạng hợp chất.",
                "C. chỉ tồn tại ở dạng đơn chất.",
                "D. chỉ sinh ra khi núi lửa hoạt động ở dạng khí SO~2~."
            ],
            "solution": "Trong tự nhiên, sulfur là nguyên tố khá dồi dào, tồn tại ở cả:\n- Dạng đơn chất: trong các mỏ lưu huỳnh lớn ở vỏ Trái Đất (đặc biệt các vùng có nguồn gốc núi lửa).\n- Dạng hợp chất: trong các khoáng vật sulfide (pyrite FeS~2~, chalcopyrite CuFeS~2~, galena PbS...) và sulfate (thạch cao CaSO~4~·2H~2~O, muối Epsom MgSO~4~·7H~2~O...), cũng như trong cơ thể sinh vật (protein, amino acid methionine, cysteine).\n➔ Chọn đáp án B."
        },
        {
            "id": 7,
            "prompt": "Câu 7: Ứng dụng không phải của ammonia là",
            "options": [
                "*A.* Sản xuất sulfuric acid.",
                "B. Sử dụng làm dung môi.",
                "C. Sản xuất nitric acid.",
                "D. Sản xuất phân bón như: đạm, ammophos."
            ],
            "solution": "- Ammonia (NH~3~) được ứng dụng rộng rãi để: sản xuất các loại phân bón (đạm urê, đạm amoni, ammonium sulfate, ammophos), sản xuất nitric acid (HNO~3~), làm chất sinh hàn (gas lạnh công nghiệp), làm dung môi hòa tan một số chất vô cơ và hữu cơ.\n- Sulfuric acid (H~2~SO~4~) được sản xuất từ nguyên liệu sulfur (S) hoặc quặng pyrite (FeS~2~), không phải từ ammonia.\n➔ Chọn đáp án A."
        },
        {
            "id": 8,
            "prompt": "Câu 8: Sulfur dioxide (SO~2~) là một trong những chất gây ô nhiễm khí quyển phổ biến. Biện pháp nào sau đây không thể giảm phát thải SO~2~?",
            "options": [
                "A. Có biện pháp tái chế các sản phẩm phụ có chứa sulfur.",
                "B. Phát triển các nguồn năng lượng sạch, thân thiện với môi trường.",
                "*C.* Sử dụng các nguồn nhiên liệu như than đá, dầu mỏ.",
                "D. Xây dựng hệ thống xử lí khí thải tại các nhà máy, khu công nghiệp."
            ],
            "solution": "Nhiên liệu hóa thạch (than đá, dầu mỏ) chứa một hàm lượng tạp chất lưu huỳnh nhất định. Khi đốt cháy, sulfur bị oxi hóa thành khí SO~2~ thải vào khí quyển gây ô nhiễm và mưa acid. Do đó việc sử dụng than đá, dầu mỏ sẽ làm tăng phát thải khí SO~2~ chứ không thể giảm.\n➔ Chọn đáp án C."
        },
        {
            "id": 9,
            "prompt": "Câu 9: Học sinh X thực hiện thí nghiệm chuẩn độ dung dịch NaOH chưa biết chính xác nồng độ bằng dung dịch chuẩn HCl 0,1 M với chỉ thị phenolphthalein. Lấy 10,00 mL dung dịch HCl 0,1 M cho vào bình tam giác cùng với 1-2 giọt phenolphthalein.\nKết quả ghi lại thể tích NaOH qua 3 lần chuẩn độ như sau:\n  • Lần 1: 19,85 mL\n  • Lần 2: 19,95 mL\n  • Lần 3: 20,20 mL\nNồng độ (mol/L) của dung dịch NaOH trong thí nghiệm trên là:",
            "options": [
                "A. 0,5.",
                "*B.* 0,05.",
                "C. 0,1.",
                "D. 0,06."
            ],
            "image": "San_Pham/images_de_1202/chuan_do_clean.png",
            "image_width": 4.5,
            "solution": "1. Tính thể tích trung bình của dung dịch NaOH qua 3 lần chuẩn độ:\n   V_NaOH trung bình = (19,85 + 19,95 + 20,20) / 3 = 60,00 / 3 = 20,00 mL = 0,020 L.\n2. Phương trình phản ứng chuẩn độ:\n   NaOH + HCl ➔ NaCl + H~2~O\n3. Tại điểm tương đương, số mol NaOH phản ứng bằng số mol HCl:\n   n_NaOH = n_HCl = C_HCl × V_HCl = 0,1 M × 0,010 L = 10^-3^ mol.\n4. Nồng độ mol của dung dịch NaOH:\n   C_NaOH = n_NaOH / V_NaOH = 10^-3^ / 0,020 = 0,05 mol/L (M).\n➔ Chọn đáp án B."
        },
        {
            "id": 10,
            "prompt": "Câu 10: Thực hiện thí nghiệm: Sulfur tác dụng với oxygen\nChuẩn bị: bột sulfur, bình khí oxygen, muôi sắt, đèn cồn, nút cao su.\nTiến hành:\n- Lấy một ít bột sulfur vào muôi sắt (đã cắm xuyên qua nút cao su).\n- Hơ nóng muôi sắt trên ngọn lửa đèn cồn đến khi sulfur nóng chảy và cháy một phần trong không khí.\n- Đưa nhanh muôi sắt vào bình khí oxygen.\nKết luận nào sau là không phù hợp?",
            "options": [
                "A. Nếu đốt cháy hoàn toàn 6,4 gam sulfur trong oxygen dư thì thể tích SO~2~ (ở đkc) thu được là 4,958 L.",
                "B. Sulfur tác dụng với oxygen tạo thành hợp chất oxide acid (acidic oxide).",
                "*C.* Phản ứng chứng minh sulfur bị khử bởi oxygen.",
                "D. Sulfur cháy mạnh hơn khi đưa vào bình khí oxygen."
            ],
            "solution": "- Phương trình hóa học: S^0^ + O~2~^0^ ➔ S^+4^O~2~^-2^.\n- Trong phản ứng trên, số oxi hóa của sulfur tăng từ 0 lên +4 nên sulfur đóng vai trò là chất khử, tức là sulfur bị oxi hóa bởi oxygen (chứ không phải bị khử). Do đó kết luận C sai.\n- Kiểm tra ý A: n_S = 6,4 / 32 = 0,2 mol ➔ n_SO~2~ = 0,2 mol. Ở đkc (25°C, 1 bar): V_SO~2~ = 0,2 × 24,79 = 4,958 L (đúng).\n- SO~2~ là acidic oxide tương ứng của H~2~SO~3~ (đúng).\n- Trong bình chứa oxygen nguyên chất nồng độ O~2~ cao hơn nhiều so với không khí (khoảng 21%) nên sulfur cháy sáng mạnh hơn (đúng).\n➔ Chọn đáp án C."
        },
        {
            "id": 11,
            "prompt": "Câu 11: Nitrogen monoxide được tạo thành khi mưa dông kèm theo sấm sét do phản ứng giữa nitrogen và oxygen trong không khí được gọi là:",
            "options": [
                "A. NO~x~ nhiên liệu.",
                "B. NO~x~ tự nhiên.",
                "*C.* NO~x~ nhiệt.",
                "D. NO~x~ tức thời."
            ],
            "solution": "Theo cơ chế hóa học môi trường và SGK Hóa học 11:\n- NO~x~ sinh ra do phản ứng trực tiếp giữa N~2~ và O~2~ ở nhiệt độ rất cao (trên 2000°C) hoặc dưới tác dụng của tia lửa điện (sấm sét) trong các cơn mưa dông được gọi là NO~x~ nhiệt (thermal NO~x~).\n- Phản ứng: N~2~(g) + O~2~(g) ⇌ 2NO(g).\n➔ Chọn đáp án C."
        },
        {
            "id": 12,
            "prompt": "Câu 12: Sự điện li là quá trình:",
            "options": [
                "A. phân li các chất trong nước thành các chất đơn giản.",
                "*B.* phân li các chất trong nước tạo thành ion.",
                "C. phân li các chất trong nước thành phần tử nhỏ hơn.",
                "D. hòa tan các chất trong nước."
            ],
            "solution": "Theo định nghĩa SGK Hóa học 11:\nSự điện li là quá trình phân li các chất trong nước (hoặc khi nóng chảy) tạo thành các ion dương (cation) và ion âm (anion).\n➔ Chọn đáp án B."
        },
        {
            "id": 13,
            "prompt": "Câu 13: Đo pH của một cốc nước chanh được giá trị pH bằng 2,4. Nhận định nào sau đây là đúng?",
            "options": [
                "A. Nước chanh có môi trường base.",
                "B. Nước chanh làm phenolphthalein hóa màu hồng.",
                "C. Nước chanh làm giấy quỳ tím hóa xanh.",
                "*D.* Nước chanh làm giấy quỳ tím hóa đỏ."
            ],
            "solution": "- Ta có pH = 2,4 < 7 nên nước chanh có môi trường acid (chứa citric acid và ascorbic acid).\n- Dung dịch acid làm giấy quỳ tím chuyển sang màu đỏ (hoặc hồng nhạt).\n- Chỉ thị phenolphthalein chỉ chuyển màu hồng trong môi trường base (pH ≥ 8,3), trong môi trường acid phenolphthalein không có màu.\n➔ Chọn đáp án D."
        },
        {
            "id": 14,
            "prompt": "Câu 14: Cặp kim loại nào sau đều không phản ứng được với nitric acid?",
            "options": [
                "A. Fe, Mg.",
                "B. Ag, Al.",
                "C. Cu, Zn.",
                "*D.* Au, Pt."
            ],
            "solution": "Nitric acid (HNO~3~) là một acid có tính oxi hóa rất mạnh, hòa tan được hầu hết các kim loại (kể cả kim loại đứng sau H trong dãy hoạt động hóa học như Cu, Ag), trừ hai kim loại quý trơ là vàng (Au) và platinum (Pt).\n(Vàng và bạch kim chỉ tan trong dung dịch nước cường toan / cường thủy gồm hỗn hợp HCl đặc và HNO~3~ đặc theo tỉ lệ thể tích 3 : 1).\n➔ Chọn đáp án D."
        },
        {
            "id": 15,
            "prompt": "Câu 15: Dựa vào công thức cấu tạo của nitric acid dưới đây:\nCho biết phát biểu nào sau đây không đúng?",
            "options": [
                "*A.* Liên kết O-H phân cực mạnh về phía nguyên tử hydrogen.",
                "B. Liên kết N ➔ O là liên kết cho nhận.",
                "C. Công thức phân tử nitric acid là HNO~3~.",
                "D. Nguyên tử N có số oxi hóa +5, là số oxi hóa cao nhất của nitrogen."
            ],
            "image": "San_Pham/images_de_1202/hno3_hd.png",
            "image_width": 3.2,
            "solution": "- Do độ âm điện của Oxygen (χ_O ≈ 3,44) lớn hơn nhiều so với Hydrogen (χ_H ≈ 2,20) nên cặp electron dùng chung trong liên kết O-H bị lệch mạnh về phía nguyên tử oxygen (chứ không phải hydrogen). Vì vậy H dễ tách ra dưới dạng H^+ thể hiện tính acid. Phát biểu A sai.\n- Liên kết N ➔ O là liên kết cho - nhận (cặp electron do nguyên tử N đóng góp hoàn toàn): đúng.\n- Công thức phân tử là HNO~3~: đúng.\n- Nitrogen có số oxi hóa cao nhất là +5: đúng.\n➔ Chọn đáp án A."
        },
        {
            "id": 16,
            "prompt": "Câu 16: Ở trong tự nhiên, nguyên tố nitrogen tồn tại ở dạng:",
            "options": [
                "A. chỉ ở dạng đơn chất.",
                "B. chỉ ở dạng hợp chất.",
                "C. chỉ ở dạng ion nitrate.",
                "*D.* cả đơn chất và hợp chất."
            ],
            "solution": "Trong tự nhiên, nitrogen tồn tại ở cả hai dạng:\n- Đơn chất: khí N~2~ chiếm khoảng 78% thể tích không khí quyển Trái Đất.\n- Hợp chất: khoáng diêm tiêu (NaNO~3~), các ion nitrate (NO~3~^-^), amoni (NH~4~^+^) trong đất, và đặc biệt là thành phần cấu tạo nên các hợp chất hữu cơ quan trọng của sự sống như protein, amino acid, acid nucleic...\n➔ Chọn đáp án D."
        }
    ]
    
    for q in questions_p1:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, q["prompt"])
        
        if "image" in q and os.path.exists(q["image"]):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(2)
            p_img.add_run().add_picture(q["image"], width=Cm(q.get("image_width", 5.0)))
        
        render_options_no_table(doc, q["options"], is_teacher=is_teacher)
        
        if is_teacher and "solution" in q:
            add_solution_no_table(doc, q["solution"])

    # -------------------------------------------------------------
    # PHẦN II. TRẮC NGHIỆM ĐÚNG SAI (2,0 ĐIỂM)
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI (2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Trong mỗi ý a), b), c), d) ở mỗi câu, học sinh chọn đúng hoặc sai."
    )
    
    questions_p2 = [
        {
            "id": 1,
            "prompt": "Câu 1: Dung dịch X chứa H~2~SO~4~ 0,05 M ở 25°C. Xét từng phát biểu sau đây là đúng hay sai?",
            "statements": [
                ("a) H~2~SO~4~ là chất điện li mạnh.", "Đúng"),
                ("b) Theo thuyết Brønsted - Lowry, H~2~SO~4~ đóng vai trò acid vì có khả năng nhận proton (H^+).", "Sai"),
                ("c) Dung dịch X làm phenolphthalein chuyển màu hồng.", "Sai"),
                ("d) Trộn 300 mL dung dịch chứa NaOH 0,15 M và Ba(OH)~2~ 0,05 M với 200 mL dung dịch X trên thu được 500 mL dung dịch có pH = 13 (Kết quả đã làm tròn số đến phần nguyên).", "Đúng")
            ],
            "solution": "a) ĐÚNG: H~2~SO~4~ là acid mạnh, nấc 1 phân li hoàn toàn: H~2~SO~4~ ➔ H^+ + HSO~4~^-; nấc 2 phân li gần như hoàn toàn ở nồng độ loãng. Do đó H~2~SO~4~ là chất điện li mạnh.\nb) SAI: Theo thuyết Brønsted - Lowry, acid là chất có khả năng cho (nhường) proton (H^+), còn base là chất nhận proton (H^+). Do đó H~2~SO~4~ đóng vai trò acid vì có khả năng CHO proton.\nc) SAI: Dung dịch X chứa H~2~SO~4~ có môi trường acid (pH < 7). Chỉ thị phenolphthalein không đổi màu (không màu) trong môi trường acid, chỉ chuyển màu hồng trong môi trường kiềm (pH ≥ 8,3).\nd) ĐÚNG:\n- Số mol OH^- trong 300 mL dung dịch base:\n  n_NaOH = 0,3 × 0,15 = 0,045 mol;  n_Ba(OH)~2~ = 0,3 × 0,05 = 0,015 mol.\n  ➔ Tổng n_OH^- = n_NaOH + 2 × n_Ba(OH)~2~ = 0,045 + 2 × 0,015 = 0,075 mol.\n- Số mol H^+ trong 200 mL dung dịch X (H~2~SO~4~ 0,05 M):\n  n_H~2~SO~4~ = 0,2 × 0,05 = 0,01 mol ➔ n_H^+ = 2 × 0,01 = 0,02 mol.\n- Phản ứng trung hòa: H^+ + OH^- ➔ H~2~O.\n  Ta thấy n_OH^- > n_H^+ ➔ OH^- còn dư:\n  n_OH^- dư = 0,075 - 0,02 = 0,055 mol.\n- Tổng thể tích sau trộn: V = 300 + 200 = 500 mL = 0,5 L.\n- Nồng độ OH^- dư:\n  [OH^-] = 0,055 / 0,5 = 0,11 M.\n  pOH = -lg(0,11) ≈ 0,9586.\n  ➔ pH = 14 - pOH = 14 - 0,9586 = 13,0414 ≈ 13 (làm tròn số nguyên)."
        },
        {
            "id": 2,
            "prompt": "Câu 2: Hình 5.1 dưới đây mô tả công thức Lewis và dạng hình học của phân tử ammonia.",
            "image": "San_Pham/images_de_1202/nh3_hd.png",
            "image_width": 7.0,
            "caption": "Hình 5.1. Công thức Lewis (a) và dạng hình học (b) của phân tử ammonia",
            "statements": [
                ("a) Trong công nghiệp, ammonia thường được sử dụng đối với vai trò tác nhân làm lạnh (chất sinh hàn).", "Đúng"),
                ("b) Trong công nghiệp, phản ứng giữa NH~3~ với acid được dùng để sản xuất phân bón.", "Đúng"),
                ("c) NH~3~ tan nhiều trong nước chủ yếu do phân tử NH~3~ phân cực, có khả năng tạo liên kết hydrogen mạnh với nước.", "Đúng"),
                ("d) Phương trình hoá học của phản ứng tổng hợp ammonia từ nitrogen và hydrogen bằng quá trình Haber như sau:\n   N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)   Δ~r~H°~298~ = -91,8 kJ.\n   Nếu giảm áp suất của hệ thì cân bằng sẽ chuyển dịch theo chiều thuận, làm tăng hiệu suất tổng hợp NH~3~.", "Sai")
            ],
            "solution": "a) ĐÚNG: Ammonia dễ hóa lỏng, nhiệt hóa hơi lớn nên được ứng dụng phổ biến làm chất sinh hàn trong tủ lạnh và hệ thống lạnh công nghiệp.\nb) ĐÚNG: NH~3~ phản ứng với các acid (H~2~SO~4~, HNO~3~, H~3~PO~4~) tạo các loại phân đạm và phân phức hợp quan trọng: (NH~4~)~2~SO~4~, NH~4~NO~3~, NH~4~H~2~PO~4~...\nc) ĐÚNG: Phân tử NH~3~ có cấu tạo hình chóp tam giác phân cực mạnh và nguyên tử N có độ âm điện lớn, mang cặp electron tự do nên tạo được liên kết hydrogen liên phân tử rất bền với nước, dẫn đến độ tan cực lớn (1 lít nước hòa tan khoảng 700 lít NH~3~ ở 20°C).\nd) SAI: Vế trái có 4 mol khí (1 N~2~ + 3 H~2~), vế phải có 2 mol khí (2 NH~3~). Theo nguyên lí Le Chatelier, khi giảm áp suất cân bằng sẽ chuyển dịch theo chiều làm tăng số mol khí, tức là chiều nghịch (từ 2 mol sang 4 mol), làm giảm hiệu suất tổng hợp NH~3~."
        }
    ]
    
    for q in questions_p2:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        render_question_prompt(p, q["prompt"])
        
        if "image" in q and os.path.exists(q["image"]):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(1)
            p_img.add_run().add_picture(q["image"], width=Cm(q.get("image_width", 6.0)))
            if "caption" in q:
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cap.paragraph_format.space_before = Pt(0)
                p_cap.paragraph_format.space_after = Pt(3)
                render_rich_text(p_cap, f"_{q['caption']}_", base_size=9.5, base_italic=True)
                
        for stmt, ans in q["statements"]:
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Cm(0.5)
            p_s.paragraph_format.space_before = Pt(1)
            p_s.paragraph_format.space_after = Pt(1)
            render_statement(p_s, stmt, ans=ans, is_teacher=is_teacher)
            
        if is_teacher and "solution" in q:
            add_solution_no_table(doc, q["solution"])

    # -------------------------------------------------------------
    # PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 ĐIỂM)
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="PHẦN III. CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN (2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 4."
    )
    
    questions_p3 = [
        {
            "id": 1,
            "prompt": "Câu 1: Xét các hệ cân bằng sau:\n(1) N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g)\n(2) 2NO~2~(g) ⇌ N~2~O~4~(g)\n(3) 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g)\n(4) 3O~2~(g) ⇌ 2O~3~(g)\nKhi tăng áp suất, có bao nhiêu cân bằng hoá học chuyển dịch theo chiều thuận?",
            "ans": "4",
            "solution": "Theo nguyên lí Le Chatelier, khi tăng áp suất, cân bằng sẽ chuyển dịch theo chiều làm giảm số mol chất khí:\n- Hệ (1): n~trước~ = 1 + 3 = 4 mol > n~sau~ = 2 mol ➔ Cân bằng chuyển dịch theo chiều thuận.\n- Hệ (2): n~trước~ = 2 mol > n~sau~ = 1 mol ➔ Cân bằng chuyển dịch theo chiều thuận.\n- Hệ (3): n~trước~ = 2 + 1 = 3 mol > n~sau~ = 2 mol ➔ Cân bằng chuyển dịch theo chiều thuận.\n- Hệ (4): n~trước~ = 3 mol > n~sau~ = 2 mol ➔ Cân bằng chuyển dịch theo chiều thuận.\nNhư vậy, cả 4 cân bằng đều chuyển dịch theo chiều thuận khi tăng áp suất.\n➔ Trả lời: 4"
        },
        {
            "id": 2,
            "prompt": "Câu 2: Sulfur có thể tham gia những phản ứng hoá học sau (Cho biết điều kiện có đủ):\n(1) S + O~2~ ➔ SO~2~\n(2) S + H~2~ ➔ H~2~S\n(3) 3F~2~ + S ➔ SF~6~\n(4) Hg + S ➔ HgS\n(5) 2Al + 3S ➔ Al~2~S~3~\nCó bao nhiêu phản ứng trong đó sulfur thể hiện tính khử?",
            "ans": "2",
            "solution": "Sulfur đơn chất có số oxi hóa bằng 0 (S^0^). Khi thể hiện tính khử, số oxi hóa của sulfur tăng lên (> 0):\n- (1) S^0^ + O~2~ ➔ S^+4^O~2~ (Số oxi hóa tăng từ 0 lên +4 ➔ thể hiện tính khử) (1).\n- (2) S^0^ + H~2~ ➔ H~2~S^-2^ (Số oxi hóa giảm từ 0 xuống -2 ➔ thể hiện tính oxi hóa).\n- (3) 3F~2~ + S^0^ ➔ S^+6^F~6~ (Số oxi hóa tăng từ 0 lên +6 ➔ thể hiện tính khử) (2).\n- (4) Hg + S^0^ ➔ HgS^-2^ (Số oxi hóa giảm từ 0 xuống -2 ➔ thể hiện tính oxi hóa).\n- (5) 2Al + 3S^0^ ➔ Al~2~S~3~^-2^ (Số oxi hóa giảm từ 0 xuống -2 ➔ thể hiện tính oxi hóa).\nCó 2 phản ứng sulfur thể hiện tính khử là (1) và (3).\n➔ Trả lời: 2"
        },
        {
            "id": 3,
            "prompt": "Câu 3: Số oxi hóa của nguyên tử nitrogen trong phân tử ammonia là bao nhiêu?",
            "ans": "-3",
            "solution": "Trong phân tử NH~3~, liên kết giữa N và H là liên kết cộng hóa trị có cực, độ âm điện của H (2,20) bé hơn của N (3,04) nên số oxi hóa của H là +1.\nGọi số oxi hóa của N là x: x + 3 × (+1) = 0 ➔ x = -3.\n➔ Trả lời: -3"
        },
        {
            "id": 4,
            "prompt": "Câu 4: Hiện tượng phú dưỡng là hiện tượng nguồn nước dư quá nhiều chất dinh dưỡng (nitrogen, phosphorus).\na) Có thể quan sát được hiện tượng phú dưỡng thông qua sự xuất hiện dày đặc của rong, tảo xanh trong nước.\nb) Nguồn dinh dưỡng ở ao, hồ thường có nguồn gốc từ nước mưa đọng và lưu giữ lại ở ao, hồ.\nc) Khi hàm lượng trong nước của nitrogen đạt 300 µg/L và phosphorus đạt 20 µg/L sẽ gây ra hiện tượng phú dưỡng.\nd) Nước thải, hay các đầm nuôi trồng thủy sản, sự dư thừa thức ăn chăn nuôi, phân bón hóa học... cũng gây hiện tượng phú dưỡng.\nCó bao nhiêu nhận định trên là đúng?",
            "ans": "3",
            "solution": "- a) ĐÚNG: Ao hồ phú dưỡng làm tảo, rong bèo phát triển bùng phát tạo thành váng xanh phủ kín mặt nước (\"tảo nở hoa\").\n- b) SAI: Nước mưa tự nhiên rất nghèo dinh dưỡng. Dinh dưỡng dư thừa (N, P) ở ao hồ chủ yếu từ nước thải sinh hoạt, công nghiệp và dòng chảy tràn phân bón từ đồng ruộng.\n- c) ĐÚNG: Theo quy chuẩn SGK Hóa học 11, ao hồ bị coi là phú dưỡng khi nồng độ tổng nitrogen đạt từ 300 µg/L trở lên và tổng phosphorus đạt từ 20 µg/L trở lên.\n- d) ĐÚNG: Nước thải sinh hoạt, ao đầm nuôi thủy sản, thức ăn chăn nuôi dư thừa, phân hóa học rửa trôi là những nguồn chính gây phú dưỡng.\nCó 3 nhận định đúng là (a, c, d).\n➔ Trả lời: 3"
        }
    ]
    
    for q in questions_p3:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        render_question_prompt(p, q["prompt"])
        
        p_ans = doc.add_paragraph()
        p_ans.paragraph_format.left_indent = Cm(0.5)
        p_ans.paragraph_format.space_before = Pt(1)
        p_ans.paragraph_format.space_after = Pt(1)
        if is_teacher:
            render_rich_text(p_ans, f"*Trả lời:* *{q['ans']}*", base_size=10.5, base_bold=True, base_color=COLOR_CORRECT)
        else:
            render_rich_text(p_ans, "*Trả lời:* .....................................................................", base_size=10.5)
            
        if is_teacher and "solution" in q:
            add_solution_no_table(doc, q["solution"])

    # -------------------------------------------------------------
    # PHẦN IV. TỰ LUẬN (2,0 ĐIỂM)
    # -------------------------------------------------------------
    add_section_header(
        doc,
        title="PHẦN IV. TỰ LUẬN (2,0 điểm)",
        subtitle="Học sinh trả lời từ câu 1 đến câu 2. Học sinh trình bày lời giải trên giấy bài thi."
    )
    
    # Câu 1 Tự luận
    p_tl1 = doc.add_paragraph()
    p_tl1.paragraph_format.space_before = Pt(4)
    p_tl1.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_tl1, "Câu 1 (1,0 điểm): Cho cân bằng hóa học sau: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g)   Δ~r~H°~298~ = -198,4 kJ.\na) Vận dụng nguyên lí Le Chatelier, hãy cho biết cần tác động các yếu tố như: nhiệt độ, nồng độ, áp suất như thế nào để cân bằng trên chuyển dịch theo chiều thuận?\nb) Cho hỗn hợp X ban đầu gồm SO~2~ và O~2~ có tỉ khối so với H~2~ bằng 24. Nung nóng X trong bình kín ở 450°C, có xúc tác V~2~O~5~, áp suất 1 bar. Ở trạng thái cân bằng, thu được hỗn hợp khí Y có tỉ khối so với H~2~ bằng 30. Tính hằng số cân bằng (K~C~) cho phản ứng trên tại 450°C?")
    
    if is_teacher:
        sol_tl1 = """a) Vận dụng nguyên lí Le Chatelier để cân bằng chuyển dịch theo chiều thuận:
- Về nhiệt độ: Phản ứng thuận tỏa nhiệt (Δ_r H°_298 = -198,4 kJ < 0). Do đó, khi GIẢM NHIỆT ĐỘ, cân bằng sẽ chuyển dịch theo chiều thuận (tỏa nhiệt). Trong thực tế công nghiệp, nhiệt độ được duy trì ở khoảng 450°C để đảm bảo tốc độ phản ứng và hoạt tính của chất xúc tác V_2O_5.
- Về áp suất: Phản ứng thuận làm giảm số mol khí (từ 3 mol khí vế trái xuống 2 mol khí vế phải). Do đó, khi TĂNG ÁP SUẤT của hệ, cân bằng sẽ chuyển dịch theo chiều thuận (chiều giảm số mol khí).
- Về nồng độ: Cần TĂNG NỒNG ĐỘ các chất phản ứng (thêm SO_2 hoặc thêm O_2/không khí) hoặc GIẢM NỒNG ĐỘ sản phẩm (tách liên tục SO_3 ra khỏi hỗn hợp phản ứng bằng cách hấp thụ vào H_2SO~4~ đặc).

b) Tính hằng số cân bằng K_C:
- Bước 1: Xét hỗn hợp X ban đầu:
  M_X = 24 × 2 = 48 g/mol.
  Gọi số mol SO~2~ là a, số mol O~2~ là b:
  (64a + 32b) / (a + b) = 48 ➔ 16a = 16b ➔ a = b.
  Giả sử ban đầu có 1 mol SO~2~ và 1 mol O~2~ ➔ Tổng số mol khí n_X = 2 mol. Khối lượng m_X = 1×64 + 1×32 = 96 gam.
- Bước 2: Xét hỗn hợp Y tại cân bằng:
  M_Y = 30 × 2 = 60 g/mol.
  Theo định luật bảo toàn khối lượng: m_Y = m_X = 96 gam.
  Số mol hỗn hợp khí tại cân bằng: n_Y = m_Y / M_Y = 96 / 60 = 1,6 mol.
- Bước 3: Xác định lượng chất tại cân bằng:
  Phương trình:        2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g)
  Ban đầu (mol):          1          1          0
  Phản ứng (mol):        2x          x         2x
  Cân bằng (mol):      1 - 2x      1 - x       2x
  Tổng số mol khí tại cân bằng: n_Y = (1 - 2x) + (1 - x) + 2x = 2 - x (mol).
  ➔ 2 - x = 1,6 ➔ x = 0,4 mol.
  Tại cân bằng:
  n_SO~2~ = 1 - 2(0,4) = 0,2 mol;  n_O~2~ = 1 - 0,4 = 0,6 mol;  n_SO~3~ = 2 × 0,4 = 0,8 mol.
  (Hiệu suất phản ứng tính theo SO~2~: H = (0,8 / 1) × 100% = 80%).
- Bước 4: Tính hằng số cân bằng K_C:
  Gọi thể tích bình kín là V (lít). Nồng độ mol các chất tại cân bằng:
  [SO~2~] = 0,2 / V;   [O~2~] = 0,6 / V;   [SO~3~] = 0,8 / V.
  Biểu thức hằng số cân bằng:
  K_C = [SO~3~]^2 / ([SO~2~]^2 × [O~2~]) = (0,8 / V)^2 / [ (0,2 / V)^2 × (0,6 / V) ] = (0,64 / V^2) / [ (0,04 / V^2) × (0,6 / V) ]
  ➔ K_C = 16 / (0,6 / V) = (80 / 3) × V ≈ 26,67 V.
  * Nếu tính thể tích V từ điều kiện trạng thái ở cân bằng (áp suất P = 1 bar, T = 450°C = 723,15 K):
    V = (n_Y × R × T) / P = (1,6 × 0,08314 × 723,15) / 1 ≈ 96,20 lít.
    Khi đó: K_C = (80 / 3) × 96,20 ≈ 2565,3 (M^-1^).
  * (Hằng số cân bằng theo áp suất: K_P = P_SO~3~^2 / (P_SO~2~^2 × P_O~2~) = (0,5)^2 / [(0,125)^2 × 0,375] = 42,67 bar^-1^)."""
        add_solution_no_table(doc, sol_tl1)
    else:
        add_student_writing_lines(doc, num_lines=7)
        
    # Câu 2 Tự luận
    p_tl2 = doc.add_paragraph()
    p_tl2.paragraph_format.space_before = Pt(4)
    p_tl2.paragraph_format.space_after = Pt(2)
    render_question_prompt(p_tl2, "Câu 2 (1,0 điểm): Khi đốt nóng hỗn hợp bột sắt (iron) với sulfur (lưu huỳnh) được trộn đều theo tỉ lệ khối lượng khoảng 1,5 : 1 trong ống nghiệm khô chịu nhiệt. Kết thúc phản ứng thu được hợp chất có màu đen.\na) Hãy viết phương trình hóa học và xác định chất oxi hóa, chất khử của phản ứng trên.\nb) Nếu đun nóng hỗn hợp gồm 5,6 g bột Fe và 3,8 g bột S (không có không khí) đến phản ứng hoàn toàn. Khối lượng muối sulfide tạo thành là bao nhiêu gam?")
    
    if is_teacher:
        sol_tl2 = """a) Phương trình hóa học và vai trò các chất:
- Phương trình phản ứng: Fe^0^ + S^0^ ➔ Fe^+2^S^-2^ (sắt(II) sulfide, chất rắn màu đen).
- Xác định chất oxi hóa và chất khử:
  + Nguyên tử Fe có số oxi hóa tăng từ 0 lên +2: Fe là chất khử (chất bị oxi hóa).
  + Nguyên tử S có số oxi hóa giảm từ 0 xuống -2: S là chất oxi hóa (chất bị khử).

b) Tính khối lượng muối sulfide tạo thành:
- Số mol các chất ban đầu:
  n_Fe = 5,6 / 56 = 0,1 mol.
  n_S = 3,8 / 32 = 0,11875 mol.
- So sánh tỉ lệ phản ứng:
  Phương trình: Fe + S ➔ FeS.
  Tỉ lệ mol phản ứng là 1 : 1.
  Ta thấy: n_Fe / 1 = 0,1 < n_S / 1 = 0,11875 ➔ Fe phản ứng hết, S còn dư (n_S dư = 0,11875 - 0,1 = 0,01875 mol).
- Tính theo chất phản ứng hết (Fe):
  n_FeS = n_Fe = 0,1 mol.
- Khối lượng muối iron(II) sulfide (FeS) tạo thành là:
  m_FeS = n_FeS × M_FeS = 0,1 × (56 + 32) = 0,1 × 88 = 8,8 gam.
- Kết luận: Khối lượng muối sulfide tạo thành là 8,8 gam."""
        add_solution_no_table(doc, sol_tl2)
    else:
        add_student_writing_lines(doc, num_lines=7)
        
    # Kết thúc
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(4)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run("---------- HẾT ----------")
    set_run_style(r_end, font_size=10.5, bold=True)
    
    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(0)
    p_note.paragraph_format.space_after = Pt(0)
    p_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_n = p_note.add_run("Thí sinh không được sử dụng tài liệu. Cán bộ coi thi không giải thích gì thêm.")
    set_run_style(r_n, font_size=9.5, italic=True)
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    print(f"Saved: {output_path}")

if __name__ == "__main__":
    out_dir = "c:/Antigravity_Thanh/San_Pham"
    student_file = os.path.join(out_dir, "De_Kiem_Tra_GK1_Hoa_11_Ma1202_HocSinh.docx")
    teacher_file = os.path.join(out_dir, "De_Kiem_Tra_GK1_Hoa_11_Ma1202_GiaoVien.docx")
    
    print("Generating Student Word document...")
    build_exam_document(student_file, is_teacher=False)
    
    print("Generating Teacher Word document...")
    build_exam_document(teacher_file, is_teacher=True)
    
    print("Completed generation of both Word documents successfully!")
