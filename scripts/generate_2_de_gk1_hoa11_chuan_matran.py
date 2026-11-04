# -*- coding: utf-8 -*-
"""
Script soan 2 De Kiem Tra Giua Hoc Ky I - Hoa Hoc 11 (Nam hoc 2026-2027)
Chuan xac 100% theo Ma Tran De Kiem Tra 45 phut nguoi dung cung cap:
- De 01 (Ma 201) & De 02 (Ma 202)
- Thoi gian lam bai: 45 phut
- Cau truc: 27 cau / lenh hoi - Thang diem 10 (40% Biet - 30% Hieu - 30% Van dung):
  + PHAN I: 12 cau TN 4 lua chon (3,0 diem, moi cau 0,25d)
  + PHAN II: 2 cau TN Dung/Sai (8 y a,b,c,d - 2,0 diem, moi y 0,25d)
  + PHAN III: 4 cau TN Tra loi ngan (2,0 diem, moi cau 0,5d)
  + PHAN IV: 3 cau Tu luan Van dung (3,0 diem, moi cau 1,0d)
- Form De Kiem Tra chuan Thay TRAN VAN THANH:
  + Xuat Ban Hoc Sinh (co dong ke cham lam bai)
  + Xuat Ban Giao Vien (co loi giai chi tiet 100% va Bang dap an & Thang diem tu luan o cuoi de)
  + Xuat File rieng: Bang_Dap_An_Va_Thang_Diem_Tu_Luan_2_De_GK1_Hoa11.docx
"""

import os
import sys
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Import module form chuan
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
SUB_DIR = os.path.join(OUT_DIR, "2_De_Kiem_Tra_GK1_Hoa_11_Ma201_202")
os.makedirs(SUB_DIR, exist_ok=True)

# Helper styling bang
def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="90A4AE", sz="4"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'<w:tblBorders {nsdecls("w")}><w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/></w:tblBorders>')
    tblPr.append(borders)

# ==============================================================================
# DỮ LIỆU ĐỀ 01 (MÃ ĐỀ 201)
# ==============================================================================
EXAM_01 = {
    "de_num": 1,
    "made": "201",
    "title": "ĐỀ KIỂM TRA GIỮA HỌC KỲ I - ĐỀ SỐ 01",
    "part1": [
        {
            "idx": 1,
            "topic": "Bài 1: Khái niệm về cân bằng hóa học (Biết)",
            "prompt": "Phản ứng thuận nghịch là phản ứng hóa học",
            "opts": [
                "A. chỉ xảy ra theo một chiều từ chất phản ứng thành sản phẩm.",
                "B. trong cùng điều kiện, xảy ra đồng thời theo hai chiều trái ngược nhau.",
                "C. có chất phản ứng chuyển hóa hoàn toàn thành sản phẩm.",
                "D. chỉ xảy ra khi có mặt chất xúc tác thích hợp."
            ],
            "correct": "B",
            "sol": "Phản ứng thuận nghịch là phản ứng trong cùng điều kiện, xảy ra đồng thời theo hai chiều trái ngược nhau (chiều thuận và chiều nghịch)."
        },
        {
            "idx": 2,
            "topic": "Bài 1: Khái niệm về cân bằng hóa học (Biết)",
            "prompt": "Biểu thức hằng số cân bằng nồng độ (K~C~) của phản ứng: N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g) là",
            "opts": [
                "A. K~C~ = [NH~3~]^2^ / ([N~2~].[H~2~]^3^).",
                "B. K~C~ = [NH~3~] / ([N~2~].[H~2~]).",
                "C. K~C~ = ([N~2~].[H~2~]^3^) / [NH~3~]^2^.",
                "D. K~C~ = 2[NH~3~] / ([N~2~] + 3[H~2~])."
            ],
            "correct": "A",
            "sol": "Biểu thức hằng số cân bằng: K~C~ = [sản phẩm]^[hệ số] / [chất đầu]^[hệ số] = [NH~3~]^2^ / ([N~2~].[H~2~]^3^)."
        },
        {
            "idx": 3,
            "topic": "Bài 1: Khái niệm về cân bằng hóa học (Biết)",
            "prompt": "Ở trạng thái cân bằng hóa học, đại lượng nào sau đây của hệ không thay đổi theo thời gian khi nhiệt độ và áp suất giữ nguyên?",
            "opts": [
                "A. Khối lượng riêng của từng chất khí.",
                "B. Nồng độ của các chất trong hệ phản ứng.",
                "C. Thể tích của toàn bộ bình phản ứng.",
                "D. Tốc độ của phản ứng hóa học thuận và nghịch bằng 0."
            ],
            "correct": "B",
            "sol": "Khi hệ đạt cân bằng động, tốc độ phản ứng thuận bằng tốc độ phản ứng nghịch (v~t~ = v~n~ ≠ 0), nồng độ các chất trong hệ phản ứng không thay đổi theo thời gian."
        },
        {
            "idx": 4,
            "topic": "Bài 1: Khái niệm về cân bằng hóa học (Hiểu)",
            "prompt": "Cho cân bằng hóa học: CO(g) + H~2~O(g) ⇌ CO~2~(g) + H~2~(g)   Δ~r~H°~298~ = -41 kJ. Cân bằng chuyển dịch theo chiều thuận khi",
            "opts": [
                "A. tăng nhiệt độ của hệ phản ứng.",
                "B. giảm nhiệt độ của hệ phản ứng.",
                "C. tăng áp suất chung của hệ.",
                "D. giảm nồng độ của khí CO."
            ],
            "correct": "B",
            "sol": "Phản ứng có Δ~r~H°~298~ < 0 (tỏa nhiệt). Theo nguyên lý Le Chatelier, khi giảm nhiệt độ, cân bằng sẽ chuyển dịch theo chiều tỏa nhiệt để chống lại sự giảm nhiệt độ, tức là chuyển dịch theo chiều thuận."
        },
        {
            "idx": 5,
            "topic": "Bài 2: Cân bằng trong dung dịch nước (Biết)",
            "prompt": "Theo thuyết acid - base của Brønsted - Lowry, acid là chất",
            "opts": [
                "A. có khả năng cho proton (H^+).",
                "B. có khả năng nhận proton (H^+).",
                "C. phân li ra anion OH^- trong nước.",
                "D. phân li hoàn toàn thành các ion âm."
            ],
            "correct": "A",
            "sol": "Theo Brønsted - Lowry: Acid là chất có khả năng cho proton (H^+); Base là chất có khả năng nhận proton (H^+)."
        },
        {
            "idx": 6,
            "topic": "Bài 2: Cân bằng trong dung dịch nước (Hiểu)",
            "prompt": "Ở 25 °C, dung dịch nào sau đây có môi trường base (làm quỳ tím chuyển sang màu xanh)?",
            "opts": [
                "A. Dung dịch có [H^+] = 1,0 × 10^-4^ M.",
                "B. Dung dịch có pH = 5,0.",
                "C. Dung dịch có [OH^-] = 1,0 × 10^-3^ M.",
                "D. Dung dịch có [H^+] = [OH^-] = 1,0 × 10^-7^ M."
            ],
            "correct": "C",
            "sol": "Ở 25 °C, dung dịch có [OH^-] = 1,0 × 10^-3^ M > 1,0 × 10^-7^ M ⇒ [H^+] = 10^-11^ M ⇒ pH = 11 > 7 (môi trường base, quỳ tím hóa xanh)."
        },
        {
            "idx": 7,
            "topic": "Bài 3: Đơn chất nitrogen (Biết)",
            "prompt": "Khí nitrogen (N~2~) chiếm khoảng bao nhiêu phần trăm thể tích của không khí?",
            "opts": [
                "A. 21%.",
                "B. 78%.",
                "C. 0,04%.",
                "D. 1%."
            ],
            "correct": "B",
            "sol": "Trong không khí, khí nitrogen (N~2~) chiếm khoảng 78% về thể tích (khoảng 4/5 thể tích không khí)."
        },
        {
            "idx": 8,
            "topic": "Bài 3: Đơn chất nitrogen (Biết)",
            "prompt": "Ở nhiệt độ thường, đơn chất nitrogen khá trơ về mặt hóa học là do",
            "opts": [
                "A. phân tử N~2~ có liên kết ba N≡N rất bền với năng lượng liên kết lớn.",
                "B. nguyên tử nitrogen có bán kính nguyên tử rất nhỏ.",
                "C. nguyên tử nitrogen có độ âm điện lớn.",
                "D. phân tử nitrogen hoàn toàn không tan trong nước."
            ],
            "correct": "A",
            "sol": "Phân tử N~2~ chứa liên kết ba rất bền vững (năng lượng liên kết E~b~ = 945 kJ/mol), nên ở điều kiện thường rất khó bị phá vỡ để tham gia phản ứng hóa học."
        },
        {
            "idx": 9,
            "topic": "Bài 5: Ammonia. Muối ammonium (Biết)",
            "prompt": "Dung dịch ammonia (NH~3~) trong nước có tính chất nào sau đây?",
            "opts": [
                "A. Có tính acid yếu, làm quỳ tím hóa đỏ.",
                "B. Có tính base yếu, làm quỳ tím hóa xanh và phenolphthalein hóa hồng.",
                "C. Có tính trung tính, không đổi màu chất chỉ thị.",
                "D. Có tính base mạnh, phân li hoàn toàn thành ion."
            ],
            "correct": "B",
            "sol": "Phân tử NH~3~ nhận H^+ từ nước: NH~3~ + H~2~O ⇌ NH~4~^+ + OH^-. Dung dịch có tính base yếu, làm quỳ tím hóa xanh và phenolphthalein hóa hồng."
        },
        {
            "idx": 10,
            "topic": "Bài 5: Ammonia. Muối ammonium (Hiểu)",
            "prompt": "Cho dung dịch NH~4~Cl tác dụng với lượng dư dung dịch NaOH, đun nóng nhẹ. Hiện tượng quan sát được là",
            "opts": [
                "A. xuất hiện kết tủa trắng không tan.",
                "B. thoát ra chất khí có mùi khai, làm xanh giấy quỳ tím ẩm.",
                "C. xuất hiện kết tủa màu nâu đỏ.",
                "D. không có hiện tượng gì xảy ra."
            ],
            "correct": "B",
            "sol": "Phương trình ion: NH~4~^+ + OH^- ──(t°)──> NH~3~↑ + H~2~O. Khí NH~3~ sinh ra có mùi khai đặc trưng, làm xanh mẩu giấy quỳ tím ẩm đặt ở miệng ống nghiệm."
        },
        {
            "idx": 11,
            "topic": "Bài 6: Một số hợp chất với oxygen của nitrogen (Biết)",
            "prompt": "Số oxi hóa của nguyên tử nitrogen trong phân tử nitric acid (HNO~3~) là",
            "opts": [
                "A. +3.",
                "B. +4.",
                "C. +5.",
                "D. -3."
            ],
            "correct": "C",
            "sol": "Trong HNO~3~: H có số oxi hóa +1, O có số oxi hóa -2 ⇒ (+1) + x + 3 × (-2) = 0 ⇒ x = +5 (mức oxi hóa cao nhất của nitrogen)."
        },
        {
            "idx": 12,
            "topic": "Bài 6: Một số hợp chất với oxygen của nitrogen (Hiểu)",
            "prompt": "Hiện tượng mưa acid chủ yếu do sự hòa tan các khí nào sau đây vào nước mưa trong khí quyển?",
            "opts": [
                "A. CO và CO~2~.",
                "B. SO~2~ và NO~x~ (NO, NO~2~).",
                "C. CH~4~ và NH~3~.",
                "D. O~3~ và N~2~."
            ],
            "correct": "B",
            "sol": "Mưa acid sinh ra chủ yếu do khí SO~2~ (từ đốt than đá, dầu mỏ) và NO~x~ (từ khí thải động cơ, sấm sét) phản ứng với nước và oxygen trong không khí tạo thành H~2~SO~4~ và HNO~3~."
        }
    ],
    "part2": [
        {
            "idx": 1,
            "topic": "Chủ đề 1: Cân bằng hóa học",
            "intro": "Cho các nhận định về cân bằng hóa học và cân bằng trong dung dịch nước:",
            "items": [
                ("a) Khi một phản ứng thuận nghịch ở trạng thái cân bằng, nếu thêm chất xúc tác thích hợp thì cân bằng sẽ chuyển dịch theo chiều thuận để tăng tốc độ phản ứng.", "Sai"),
                ("b) Cho cân bằng: N~2~O~4~(g) [không màu] ⇌ 2NO~2~(g) [nâu đỏ] (Δ~r~H°~298~ > 0). Khi tăng nhiệt độ, cân bằng chuyển dịch theo chiều thuận làm màu nâu đỏ của hỗn hợp khí đậm dần lên.", "Đúng"),
                ("c) Trong dung dịch nước, các chất điện li mạnh (như HCl, HNO~3~, NaOH, NaCl) phân li hoàn toàn thành các ion.", "Đúng"),
                ("d) Nước cất là chất điện li rất yếu; ở 25 °C, tích số ion của nước là K~w~ = [H^+].[OH^-] = 1,0 × 10^-14^.", "Đúng")
            ],
            "sol": "• a) SAI: Chất xúc tác làm tăng đồng thời tốc độ phản ứng thuận và nghịch như nhau nên không làm chuyển dịch cân bằng hóa học, chỉ giúp hệ nhanh đạt trạng thái cân bằng.\n• b) ĐÚNG: Phản ứng có Δ~r~H° > 0 (thu nhiệt). Khi tăng nhiệt độ, cân bằng chuyển dịch theo chiều thu nhiệt (chiều thuận) tạo thêm khí NO~2~ có màu nâu đỏ đậm hơn.\n• c) ĐÚNG: Định nghĩa chất điện li mạnh phân li hoàn toàn thành ion trong dung dịch nước.\n• d) ĐÚNG: Tích số ion của nước ở 25 °C là hằng số: K~w~ = [H^+].[OH^-] = 1,0 × 10^-14^."
        },
        {
            "idx": 2,
            "topic": "Chủ đề 2: Nitrogen và sulfur",
            "intro": "Về tính chất và ứng dụng của ammonia, muối ammonium và các hợp chất oxygen của nitrogen:",
            "items": [
                ("a) Khí ammonia dễ bị hóa lỏng, tan rất nhiều trong nước do phân tử phân cực và tạo được liên kết hydrogen liên phân tử mạnh với nước.", "Đúng"),
                ("b) Tất cả các muối ammonium (như NH~4~Cl, NH~4~NO~3~, (NH~4~)~2~SO~4~) đều dễ tan trong nước và kém bền nhiệt (dễ bị nhiệt phân hủy khi đun nóng).", "Đúng"),
                ("c) Dung dịch HNO~3~ đặc, nguội làm thụ động hóa các kim loại nhôm (Al) và sắt (Fe) do tạo màng oxide bảo vệ bề mặt kim loại.", "Đúng"),
                ("d) Khí NO là khí không màu, không phản ứng với oxygen ở nhiệt độ phòng và hoàn toàn không gây độc hại cho đường hô hấp con người.", "Sai")
            ],
            "sol": "• a) ĐÚNG: NH~3~ có cấu trúc chóp tam giác, phân cực mạnh và tạo liên kết hydrogen với nước nên tan rất nhiều (1 lít nước hòa tan khoảng 800 lít NH~3~ ở 20 °C).\n• b) ĐÚNG: Các muối ammonium đều tan tốt và dễ bị phân hủy bởi nhiệt độ (NH~4~Cl ⇌ NH~3~ + HCl; NH~4~HCO~3~ ──> NH~3~ + CO~2~ + H~2~O).\n• c) ĐÚNG: HNO~3~ đặc, nguội làm Al, Fe bị thụ động hóa (không tan trong dung dịch acid).\n• d) SAI: Khí NO phản ứng rất nhanh với O~2~ ngay ở điều kiện thường tạo khí NO~2~ màu nâu đỏ (2NO + O~2~ → 2NO~2~) và cả NO, NO~2~ đều là khí độc hại gây viêm đường hô hấp."
        }
    ],
    "part3": [
        {
            "idx": 1,
            "topic": "Bài 2: Cân bằng trong dung dịch nước (Hiểu)",
            "prompt": "Hòa tan hoàn toàn 0,04 gam NaOH rắn vào nước cất, thu được 100 mL dung dịch X. Tính giá trị pH của dung dịch X ở 25 °C.",
            "ans": "12",
            "sol": "Số mol NaOH: n = 0,04 / 40 = 0,001 mol.\nNồng độ OH^-: [OH^-] = 0,001 / 0,1 = 0,01 M = 10^-2^ M.\npOH = -log(10^-2^) = 2 ⇒ pH = 14 - pOH = 14 - 2 = 12."
        },
        {
            "idx": 2,
            "topic": "Bài 3: Đơn chất nitrogen (Hiểu)",
            "prompt": "Ở điều kiện chuẩn (25 °C và 1 bar), một bình kín bằng thép dung tích 24,79 lít chứa đầy khí nitrogen (N~2~). Tính khối lượng khí nitrogen có trong bình (theo đơn vị gam).",
            "ans": "28",
            "sol": "Ở điều kiện chuẩn (1 bar, 25 °C), 1 mol khí chiếm thể tích 24,79 lít.\nSố mol khí N~2~: n = 24,79 / 24,79 = 1,0 mol.\nKhối lượng khí N~2~ trong bình: m = 1,0 × 28 = 28 gam."
        },
        {
            "idx": 3,
            "topic": "Bài 5: Ammonia. Muối ammonium (Biết)",
            "prompt": "Khối lượng mol phân tử (amu hoặc g/mol) của phân bón ammonium sulfate ((NH~4~)~2~SO~4~) là bao nhiêu?",
            "ans": "132",
            "sol": "Phân tử khối của (NH~4~)~2~SO~4~: M = 2 × (14 + 4 × 1) + 32 + 4 × 16 = 2 × 18 + 32 + 64 = 132 g/mol."
        },
        {
            "idx": 4,
            "topic": "Bài 6: Một số hợp chất với oxygen của nitrogen (Biết)",
            "prompt": "Cho dãy các oxide của nitrogen sau: N~2~O, NO, N~2~O~3~, NO~2~, N~2~O~5~. Có bao nhiêu oxide trong dãy mà nguyên tử nitrogen có số oxi hóa chẵn (+2 hoặc +4)?",
            "ans": "2",
            "sol": "Số oxi hóa của N trong từng chất:\nN~2~O (+1); NO (+2); N~2~O~3~ (+3); NO~2~ (+4); N~2~O~5~ (+5).\nCác oxide có số oxi hóa chẵn (+2 hoặc +4) là NO và NO~2~ (gồm 2 chất)."
        }
    ],
    "part4": [
        {
            "idx": 1,
            "topic": "Bài 1: Khái niệm về cân bằng hóa học (Vận dụng)",
            "prompt": "Trong một bình kín dung tích 1,0 lít ở nhiệt độ không đổi T, người ta nạp vào 1,0 mol khí H~2~ và 1,0 mol hơi I~2~ để thực hiện phản ứng thuận nghịch:\nH~2~(g) + I~2~(g) ⇌ 2HI(g)\nKhi phản ứng đạt tới trạng thái cân bằng ở nhiệt độ T, nồng độ của HI đo được là 1,6 M.\na) Xác định nồng độ mol của H~2~ và I~2~ ở trạng thái cân bằng.\nb) Tính giá trị hằng số cân bằng nồng độ (K~C~) của phản ứng ở nhiệt độ T.",
            "rubric": [
                ("Ý a (bước 1)", "Nêu phản ứng: H~2~(g) + I~2~(g) ⇌ 2HI(g)\nTheo tỉ lệ phản ứng: n_H2(pư) = n_I2(pư) = (1/2) × n_HI = (1/2) × 1,6 = 0,8 mol.", "0,25 đ"),
                ("Ý a (bước 2)", "Vì dung tích bình là 1,0 lít nên nồng độ ở trạng thái cân bằng:\n[H~2~] = [I~2~] = (1,0 - 0,8) / 1,0 = 0,2 M.", "0,25 đ"),
                ("Ý b (bước 1)", "Viết đúng biểu thức hằng số cân bằng:\nK~C~ = [HI]^2^ / ([H~2~].[I~2~])", "0,25 đ"),
                ("Ý b (bước 2)", "Thay số và tính đúng giá trị K~C~:\nK~C~ = (1,6)^2^ / (0,2 × 0,2) = 2,56 / 0,04 = 64.", "0,25 đ")
            ]
        },
        {
            "idx": 2,
            "topic": "Bài 2: Cân bằng trong dung dịch nước (Vận dụng)",
            "prompt": "Trộn 100 mL dung dịch HCl 0,08 M với 100 mL dung dịch Ba(OH)~2~ 0,05 M, thu được 200 mL dung dịch X ở 25 °C.\na) Viết phương trình ion rút gọn của phản ứng trung hòa xảy ra.\nb) Tính nồng độ mol của ion dư và xác định giá trị pH của dung dịch X.",
            "rubric": [
                ("Ý a", "Viết đúng phương trình ion rút gọn:\nH^+ + OH^- → H~2~O", "0,25 đ"),
                ("Ý b (bước 1)", "Tính số mol ban đầu của các ion:\nn_H+ = 0,1 × 0,08 = 0,008 mol.\nn_OH- = 2 × n_Ba(OH)2 = 2 × (0,1 × 0,05) = 0,010 mol.", "0,25 đ"),
                ("Ý b (bước 2)", "So sánh thấy n_OH- > n_H+ ⇒ OH^- dư:\nn_OH-(dư) = 0,010 - 0,008 = 0,002 mol.\nNồng độ ion OH^- dư trong 200 mL dung dịch X:\n[OH^-] = 0,002 / 0,2 = 0,01 M = 10^-2^ M.", "0,25 đ"),
                ("Ý b (bước 3)", "Tính đúng giá trị pH:\npOH = -log[OH^-] = -log(10^-2^) = 2 ⇒ pH = 14 - 2 = 12.", "0,25 đ")
            ]
        },
        {
            "idx": 3,
            "topic": "Bài 5: Ammonia. Muối ammonium (Vận dụng)",
            "prompt": "Đun nóng 10,7 gam ammonium chloride (NH~4~Cl) rắn với lượng dư dung dịch Ca(OH)~2~. Toàn bộ lượng khí NH~3~ thoát ra được dẫn hấp thụ hoàn toàn vào 150 mL dung dịch H~2~SO~4~ 1 M.\na) Viết các phương trình hóa học của phản ứng xảy ra.\nb) Tính thể tích khí NH~3~ thoát ra (đo ở điều kiện chuẩn: 25 °C, 1 bar).\nc) Tính khối lượng muối ammonium sulfate ((NH~4~)~2~SO~4~) thu được sau phản ứng.",
            "rubric": [
                ("Ý a", "Viết đúng 2 PTHH:\n(1) 2NH~4~Cl + Ca(OH)~2~ ──(t°)──> CaCl~2~ + 2NH~3~↑ + 2H~2~O\n(2) 2NH~3~ + H~2~SO~4~ → (NH~4~)~2~SO~4~", "0,25 đ"),
                ("Ý b", "Số mol NH~4~Cl: n = 10,7 / 53,5 = 0,2 mol.\nTheo (1): n_NH3 = n_NH4Cl = 0,2 mol.\nThể tích khí NH~3~ ở đkc: V_NH3 = 0,2 × 24,79 = 4,958 lít.", "0,25 đ"),
                ("Ý c (bước 1)", "Số mol H~2~SO~4~: n_H2SO4 = 0,15 × 1 = 0,15 mol.\nTheo phản ứng (2): n_H2SO4(pư) = (1/2) × n_NH3 = 0,1 mol < 0,15 mol ⇒ H~2~SO~4~ dư, tính muối theo NH~3~.", "0,25 đ"),
                ("Ý c (bước 2)", "Số mol muối tạo thành: n_(NH4)2SO4 = (1/2) × n_NH3 = 0,1 mol.\nKhối lượng muối ammonium sulfate:\nm = 0,1 × 132 = 13,2 gam.", "0,25 đ")
            ]
        }
    ]
}

# ==============================================================================
# DỮ LIỆU ĐỀ 02 (MÃ ĐỀ 202)
# ==============================================================================
EXAM_02 = {
    "de_num": 2,
    "made": "202",
    "title": "ĐỀ KIỂM TRA GIỮA HỌC KỲ I - ĐỀ SỐ 02",
    "part1": [
        {
            "idx": 1,
            "topic": "Bài 1: Khái niệm về cân bằng hóa học (Biết)",
            "prompt": "Cân bằng hóa học là một cân bằng động vì tại thời điểm cân bằng,",
            "opts": [
                "A. phản ứng thuận và phản ứng nghịch đều dừng lại hoàn toàn.",
                "B. phản ứng thuận và phản ứng nghịch vẫn tiếp tục diễn ra với tốc độ bằng nhau.",
                "C. chỉ có phản ứng thuận tiếp tục xảy ra với tốc độ không đổi.",
                "D. nồng độ chất phản ứng và nồng độ chất sản phẩm luôn luôn bằng nhau."
            ],
            "correct": "B",
            "sol": "Cân bằng hóa học là cân bằng động vì các phản ứng thuận và nghịch vẫn tiếp tục xảy ra với tốc độ bằng nhau (v~t~ = v~n~ ≠ 0)."
        },
        {
            "idx": 2,
            "topic": "Bài 1: Khái niệm về cân bằng hóa học (Biết)",
            "prompt": "Yếu tố nào sau đây KHÔNG làm chuyển dịch cân bằng của một hệ phản ứng thuận nghịch?",
            "opts": [
                "A. Nhiệt độ.",
                "B. Nồng độ các chất.",
                "C. Chất xúc tác.",
                "D. Áp suất (đối với phản ứng có chất khí)."
            ],
            "correct": "C",
            "sol": "Chất xúc tác làm tăng đồng thời tốc độ phản ứng thuận và nghịch như nhau nên không làm chuyển dịch cân bằng hóa học."
        },
        {
            "idx": 3,
            "topic": "Bài 1: Khái niệm về cân bằng hóa học (Biết)",
            "prompt": "Theo nguyên lý chuyển dịch cân bằng Le Chatelier, một hệ đang ở trạng thái cân bằng nếu chịu một tác động từ bên ngoài (như biến đổi nồng độ, nhiệt độ, áp suất) thì cân bằng sẽ",
            "opts": [
                "A. chuyển dịch theo chiều làm giảm tác động đó.",
                "B. chuyển dịch theo chiều làm tăng tác động đó.",
                "C. bị phá hủy hoàn toàn và không thể thiết lập cân bằng mới.",
                "D. tự động dừng lại ở trạng thái cân bằng ban đầu."
            ],
            "correct": "A",
            "sol": "Nguyên lý Le Chatelier: Khi thay đổi một trong các yếu tố nồng độ, nhiệt độ, áp suất thì cân bằng sẽ chuyển dịch theo chiều chống lại (làm giảm) tác động bên ngoài đó."
        },
        {
            "idx": 4,
            "topic": "Bài 1: Khái niệm về cân bằng hóa học (Hiểu)",
            "prompt": "Cho cân bằng hóa học trong bình kín: 2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g)   Δ~r~H°~298~ = -198 kJ. Để tăng hiệu suất chuyển hóa SO~2~ thành SO~3~, biện pháp nào sau đây là phù hợp?",
            "opts": [
                "A. Giảm áp suất chung của hệ.",
                "B. Tăng nhiệt độ của hệ phản ứng.",
                "C. Tăng áp suất chung và giảm nhiệt độ thích hợp.",
                "D. Giảm nồng độ khí O~2~ trong bình phản ứng."
            ],
            "correct": "C",
            "sol": "Chiều thuận là chiều tỏa nhiệt (ΔH < 0) và giảm số mol khí (3 mol → 2 mol). Do đó, tăng áp suất và giảm nhiệt độ thích hợp sẽ làm cân bằng chuyển dịch theo chiều thuận, làm tăng hiệu suất tạo SO~3~."
        },
        {
            "idx": 5,
            "topic": "Bài 2: Cân bằng trong dung dịch nước (Biết)",
            "prompt": "Theo thuyết Brønsted - Lowry, chất nào sau đây là một base khi hòa tan vào nước?",
            "opts": [
                "A. HCl.",
                "B. CH~3~COOH.",
                "C. NH~3~.",
                "D. HNO~3~."
            ],
            "correct": "C",
            "sol": "NH~3~ nhận proton (H^+) từ phân tử nước (NH~3~ + H~2~O ⇌ NH~4~^+ + OH^-) nên đóng vai trò là base."
        },
        {
            "idx": 6,
            "topic": "Bài 2: Cân bằng trong dung dịch nước (Hiểu)",
            "prompt": "Ion nào sau đây thể hiện tính chất lưỡng tính (vừa có khả năng cho proton, vừa có khả năng nhận proton)?",
            "opts": [
                "A. SO~4~^2-^.",
                "B. Cl^-.",
                "C. HCO~3~^-.",
                "D. Na^+."
            ],
            "correct": "C",
            "sol": "Ion HCO~3~^- là chất lưỡng tính: HCO~3~^- + H~2~O ⇌ H~3~O^+ + CO~3~^2-^ (cho H^+); HCO~3~^- + H~2~O ⇌ H~2~CO~3~ + OH^- (nhận H^+)."
        },
        {
            "idx": 7,
            "topic": "Bài 3: Đơn chất nitrogen (Biết)",
            "prompt": "Trong phản ứng tổng hợp ammonia: N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g), đơn chất nitrogen (N~2~) thể hiện",
            "opts": [
                "A. tính khử.",
                "B. tính oxi hóa.",
                "C. tính base.",
                "D. tính acid."
            ],
            "correct": "B",
            "sol": "Số oxi hóa của N giảm từ 0 (trong N~2~) xuống -3 (trong NH~3~) nên nitrogen đóng vai trò là chất oxi hóa."
        },
        {
            "idx": 8,
            "topic": "Bài 3: Đơn chất nitrogen (Biết)",
            "prompt": "Ứng dụng nào sau đây KHÔNG phải là ứng dụng phổ biến của đơn chất khí nitrogen?",
            "opts": [
                "A. Làm môi trường trơ bảo quản thực phẩm đóng gói.",
                "B. Dùng nitrogen lỏng để bảo quản mẫu máu, trứng, tinh trùng trong y học.",
                "C. Sản xuất nitric acid và phân đạm.",
                "D. Dùng làm nhiên liệu đốt trực tiếp cho động cơ ô tô."
            ],
            "correct": "D",
            "sol": "Khí nitrogen rất trơ, không cháy trong không khí nên không thể dùng làm nhiên liệu đốt cho động cơ."
        },
        {
            "idx": 9,
            "topic": "Bài 5: Ammonia. Muối ammonium (Biết)",
            "prompt": "Phân tử ammonia (NH~3~) có dạng hình học nào sau đây?",
            "opts": [
                "A. Đường thẳng.",
                "B. Chóp tam giác (tam diện).",
                "C. Tứ diện đều.",
                "D. Tam giác đều phẳng."
            ],
            "correct": "B",
            "sol": "Phân tử NH~3~ có dạng hình học chóp tam giác, nguyên tử N ở đỉnh chóp và 3 nguyên tử H ở đáy tam giác, trên nguyên tử N còn một cặp electron tự do."
        },
        {
            "idx": 10,
            "topic": "Bài 5: Ammonia. Muối ammonium (Hiểu)",
            "prompt": "Đốt cháy khí ammonia trong khí oxygen ở nhiệt độ cao không có chất xúc tác, sản phẩm thu được là",
            "opts": [
                "A. N~2~ và H~2~O.",
                "B. NO và H~2~O.",
                "C. NO~2~ và H~2~O.",
                "D. N~2~O và H~2~O."
            ],
            "correct": "A",
            "sol": "Phương trình đốt cháy không xúc tác: 4NH~3~ + 3O~2~ ──(t°)──> 2N~2~ + 6H~2~O. (Nếu có xúc tác Pt ở 850 °C thì mới tạo ra khí NO)."
        },
        {
            "idx": 11,
            "topic": "Bài 6: Một số hợp chất với oxygen của nitrogen (Biết)",
            "prompt": "Khí nào sau đây có màu nâu đỏ đặc trưng và là một trong những tác nhân chính gây ô nhiễm không khí?",
            "opts": [
                "A. N~2~O.",
                "B. NO.",
                "C. NO~2~.",
                "D. NH~3~."
            ],
            "correct": "C",
            "sol": "Nitrogen dioxide (NO~2~) là chất khí có màu nâu đỏ, mùi hắc độc hại, dễ gây viêm đường hô hấp."
        },
        {
            "idx": 12,
            "topic": "Bài 6: Một số hợp chất với oxygen của nitrogen (Hiểu)",
            "prompt": "Cho mảnh kim loại đồng (Cu) vào ống nghiệm chứa dung dịch HNO~3~ loãng, hiện tượng quan sát được là",
            "opts": [
                "A. dung dịch chuyển sang màu xanh lam và thoát ra chất khí không màu hóa nâu trong không khí.",
                "B. xuất hiện kết tủa trắng và có khí mùi khai bay lên.",
                "C. kim loại đồng không tan vì Cu đứng sau hydrogen trong dãy hoạt động hóa học.",
                "D. có khí màu nâu đỏ thoát ra ngay lập tức và dung dịch không đổi màu."
            ],
            "correct": "A",
            "sol": "Phương trình: 3Cu + 8HNO~3~(loãng) → 3Cu(NO~3~)~2~ (xanh lam) + 2NO↑ (không màu, hóa nâu ngoài không khí) + 4H~2~O."
        }
    ],
    "part2": [
        {
            "idx": 1,
            "topic": "Chủ đề 1: Cân bằng hóa học",
            "intro": "Về quá trình tổng hợp ammonia trong công nghiệp và sự phân li của chất điện li yếu trong dung dịch nước:",
            "items": [
                ("a) Trong quá trình Haber - Bosch: N~2~(g) + 3H~2~(g) ⇌ 2NH~3~(g) (Δ~r~H°~298~ = -92 kJ), để tăng hiệu suất tạo NH~3~ trong công nghiệp, người ta tiến hành ở áp suất cao (khoảng 150 - 200 bar).", "Đúng"),
                ("b) Trong quá trình Haber - Bosch, nếu giảm nhiệt độ xuống quá thấp (dưới 100 °C) thì cân bằng chuyển dịch theo chiều thuận nhưng tốc độ phản ứng lại quá chậm, không mang lại hiệu quả kinh tế.", "Đúng"),
                ("c) Acetic acid (CH~3~COOH) là chất điện li yếu, trong dung dịch nước chỉ có một phần nhỏ phân tử phân li thành ion H^+ và CH~3~COO^-.", "Đúng"),
                ("d) Khi pha loãng dung dịch CH~3~COOH 0,1 M bằng nước cất thì độ điện li α của acetic acid sẽ giảm xuống.", "Sai")
            ],
            "sol": "• a) ĐÚNG: Phản ứng tạo NH~3~ làm giảm số mol khí (4 mol → 2 mol), tăng áp suất giúp cân bằng chuyển dịch theo chiều thuận.\n• b) ĐÚNG: Nhiệt độ thấp cân bằng chuyển dịch thuận nhưng tốc độ phản ứng v quá bé, do đó trong công nghiệp phải chọn nhiệt độ tối ưu khoảng 400 - 450 °C kèm xúc tác Fe.\n• c) ĐÚNG: Acid hữu cơ CH~3~COOH là chất điện li yếu, quá trình phân li là thuận nghịch: CH~3~COOH ⇌ CH~3~COO^- + H^+.\n• d) SAI: Theo định luật pha loãng Ostwald, khi pha loãng dung dịch chất điện li yếu thì độ điện li α tăng lên."
        },
        {
            "idx": 2,
            "topic": "Chủ đề 2: Nitrogen và sulfur",
            "intro": "Cho các nhận định về ứng dụng và tính chất của các hợp chất chứa nitrogen:",
            "items": [
                ("a) Muối ammonium hydrogencarbonate (NH~4~HCO~3~) dễ bị phân hủy bởi nhiệt sinh ra các khí NH~3~, CO~2~ và hơi nước nên được ứng dụng làm bột nở trong làm bánh bao.", "Đúng"),
                ("b) Trong nông nghiệp, không nên bón phân đạm ammonium (chứa ion NH~4~^+) cùng lúc với vôi sống (CaO) vì sẽ gây thất thoát đạm dưới dạng khí NH~3~ bay ra.", "Đúng"),
                ("c) Nitric acid (HNO~3~) là một acid mạnh và là chất oxi hóa rất mạnh, có thể hòa tan hầu hết các kim loại (trừ Au, Pt) và phá hủy nhiều hợp chất hữu cơ.", "Đúng"),
                ("d) Khí nitrogen dioxide (NO~2~) tác dụng với nước khi có mặt oxygen sẽ tạo ra hỗn hợp gồm hai acid là HNO~2~ và HNO~3~.", "Sai")
            ],
            "sol": "• a) ĐÚNG: NH~4~HCO~3~ ──(t°)──> NH~3~↑ + CO~2~↑ + H~2~O↑ tạo các lỗ xốp làm bánh nở mềm.\n• b) ĐÚNG: CaO tác dụng với nước tạo Ca(OH)~2~, kiềm phản ứng với NH~4~^+ sinh ra NH~3~ bay đi mất: NH~4~^+ + OH^- → NH~3~↑ + H~2~O.\n• c) ĐÚNG: HNO~3~ có tính oxi hóa mạnh do N có số oxi hóa cao nhất (+5), hòa tan được Cu, Ag, Fe,...\n• d) SAI: Phản ứng có mặt O~2~ tạo ra duy nhất nitric acid: 4NO~2~ + O~2~ + 2H~2~O → 4HNO~3~ (chỉ khi không có O~2~ mới tạo HNO~2~ và HNO~3~)."
        }
    ],
    "part3": [
        {
            "idx": 1,
            "topic": "Bài 2: Cân bằng trong dung dịch nước (Hiểu)",
            "prompt": "Hòa tan hoàn toàn một lượng Ba(OH)~2~ vào nước thu được dung dịch có nồng độ Ba(OH)~2~ là 0,005 M. Tính giá trị pH của dung dịch này ở 25 °C.",
            "ans": "12",
            "sol": "Ba(OH)~2~ phân li hoàn toàn: Ba(OH)~2~ → Ba^2+^ + 2OH^-.\n[OH^-] = 2 × 0,005 = 0,01 M = 10^-2^ M.\npOH = -log(10^-2^) = 2 ⇒ pH = 14 - pOH = 14 - 2 = 12."
        },
        {
            "idx": 2,
            "topic": "Bài 3: Đơn chất nitrogen (Hiểu)",
            "prompt": "Cần lấy bao nhiêu lít khí nitrogen (N~2~, đo ở điều kiện chuẩn 25 °C, 1 bar) để phản ứng với lượng dư H~2~ tạo thành 0,4 mol khí NH~3~, biết hiệu suất phản ứng tổng hợp đạt 25%? (Làm tròn kết quả đến hai chữ số thập phân).",
            "ans": "19,83",
            "sol": "Phương trình: N~2~ + 3H~2~ ⇌ 2NH~3~\nSố mol N~2~ phản ứng theo lí thuyết: n_N2(lt) = (1/2) × n_NH3 = 0,2 mol.\nVì hiệu suất đạt H = 25%: n_N2(thực tế) = 0,2 / 25% = 0,8 mol.\nThể tích khí N~2~ ở đkc: V = 0,8 × 24,79 = 19,832 lít ≈ 19,83 lít."
        },
        {
            "idx": 3,
            "topic": "Bài 5: Ammonia. Muối ammonium (Biết)",
            "prompt": "Một loại phân đạm urê tinh khiết có công thức (NH~2~)~2~CO. Tính hàm lượng phần trăm (%) theo khối lượng của nguyên tố nitrogen trong phân đạm urê (làm tròn kết quả đến một chữ số thập phân).",
            "ans": "46,7",
            "sol": "Phân tử khối của (NH~2~)~2~CO: M = 2 × (14 + 2) + 12 + 16 = 32 + 12 + 16 = 60 g/mol.\n%N = (2 × 14 / 60) × 100% = 46,67% ≈ 46,7%."
        },
        {
            "idx": 4,
            "topic": "Bài 6: Một số hợp chất với oxygen của nitrogen (Biết)",
            "prompt": "Tổng số nguyên tử của tất cả các nguyên tố có trong một phân tử dinitrogen pentoxide (N~2~O~5~) là bao nhiêu?",
            "ans": "7",
            "sol": "Phân tử N~2~O~5~ gồm 2 nguyên tử nitrogen và 5 nguyên tử oxygen ⇒ Tổng số nguyên tử là 2 + 5 = 7."
        }
    ],
    "part4": [
        {
            "idx": 1,
            "topic": "Bài 1: Khái niệm về cân bằng hóa học (Vận dụng)",
            "prompt": "Xét cân bằng hóa học thực hiện trong bình kín dung tích 2,0 lít ở nhiệt độ t °C:\n2SO~2~(g) + O~2~(g) ⇌ 2SO~3~(g)\nBan đầu người ta nạp vào bình 4,0 mol SO~2~ và 2,0 mol O~2~. Khi hệ đạt trạng thái cân bằng ở nhiệt độ t °C, lượng khí SO~3~ thu được là 3,2 mol.\na) Tính nồng độ mol của các chất SO~2~, O~2~ và SO~3~ ở trạng thái cân bằng.\nb) Tính giá trị hằng số cân bằng nồng độ (K~C~) của phản ứng ở nhiệt độ t °C.",
            "rubric": [
                ("Ý a (bước 1)", "Số mol SO~3~ sinh ra: n_SO3 = 3,2 mol.\nTheo PTHH: n_SO2(pư) = n_SO3 = 3,2 mol; n_O2(pư) = (1/2) × 3,2 = 1,6 mol.", "0,25 đ"),
                ("Ý a (bước 2)", "Số mol các chất ở trạng thái cân bằng:\nn_SO2 = 4,0 - 3,2 = 0,8 mol; n_O2 = 2,0 - 1,6 = 0,4 mol; n_SO3 = 3,2 mol.\nVì dung tích bình V = 2,0 lít nên nồng độ mol ở cân bằng:\n[SO~2~] = 0,8 / 2 = 0,4 M; [O~2~] = 0,4 / 2 = 0,2 M; [SO~3~] = 3,2 / 2 = 1,6 M.", "0,25 đ"),
                ("Ý b (bước 1)", "Viết đúng biểu thức hằng số cân bằng K~C~:\nK~C~ = [SO~3~]^2^ / ([SO~2~]^2^.[O~2~])", "0,25 đ"),
                ("Ý b (bước 2)", "Thay số và tính đúng giá trị K~C~:\nK~C~ = (1,6)^2^ / ((0,4)^2^ × 0,2) = 2,56 / (0,16 × 0,2) = 2,56 / 0,032 = 80.", "0,25 đ")
            ]
        },
        {
            "idx": 2,
            "topic": "Bài 2: Cân bằng trong dung dịch nước (Vận dụng)",
            "prompt": "Để xác định nồng độ của một dung dịch HCl chưa biết nồng độ, người ta tiến hành phép chuẩn độ: Lấy chính xác 20,0 mL dung dịch HCl này cho vào bình tam giác, nhỏ thêm 2 giọt chỉ thị phenolphthalein. Sau đó chuẩn độ bằng dung dịch chuẩn NaOH 0,10 M từ buret.\na) Nêu hiện tượng quan sát được tại điểm tương đương của phép chuẩn độ.\nb) Biết rằng để làm đổi màu chất chỉ thị cần dùng vừa đúng 24,0 mL dung dịch NaOH 0,10 M. Tính nồng độ mol của dung dịch HCl ban đầu.",
            "rubric": [
                ("Ý a", "Tại điểm tương đương: Dung dịch trong bình tam giác chuyển từ không màu sang màu hồng nhạt bền vững trong khoảng 30 giây.", "0,25 đ"),
                ("Ý b (bước 1)", "Viết đúng phương trình hóa học của phản ứng chuẩn độ:\nHCl + NaOH → NaCl + H~2~O", "0,25 đ"),
                ("Ý b (bước 2)", "Tính số mol NaOH tiêu tốn:\nn_NaOH = 0,024 × 0,10 = 0,0024 mol.\nTheo PTHH: n_HCl = n_NaOH = 0,0024 mol.", "0,25 đ"),
                ("Ý b (bước 3)", "Tính nồng độ mol của dung dịch HCl:\nC_M(HCl) = n_HCl / V_dd = 0,0024 / 0,020 = 0,12 M.", "0,25 đ")
            ]
        },
        {
            "idx": 3,
            "topic": "Bài 5: Ammonia. Muối ammonium (Vận dụng)",
            "prompt": "Nhiệt phân hoàn toàn 39,5 gam muối ammonium hydrogencarbonate (NH~4~HCO~3~) trong bình kín không có không khí.\na) Viết phương trình hóa học của phản ứng nhiệt phân muối NH~4~HCO~3~.\nb) Tính tổng thể tích hỗn hợp các khí và hơi thu được sau phản ứng (đo ở điều kiện chuẩn 25 °C, 1 bar).\nc) Nếu dẫn toàn bộ lượng khí và hơi thu được đi qua bình đựng dung dịch H~2~SO~4~ đặc dư thì thể tích khí thoát ra khỏi bình là bao nhiêu lít (đo ở đkc)?",
            "rubric": [
                ("Ý a", "Viết đúng phương trình nhiệt phân:\nNH~4~HCO~3~ ──(t°)──> NH~3~↑ + CO~2~↑ + H~2~O↑", "0,25 đ"),
                ("Ý b (bước 1)", "Số mol NH~4~HCO~3~: n = 39,5 / 79 = 0,5 mol.\nTheo PTHH: n_NH3 = n_CO2 = n_H2O = 0,5 mol.\nTổng số mol khí và hơi thu được: n_khí = 0,5 × 3 = 1,5 mol.", "0,25 đ"),
                ("Ý b (bước 2)", "Tính tổng thể tích khí và hơi ở đkc:\nV_tổng = 1,5 × 24,79 = 37,185 lít.", "0,25 đ"),
                ("Ý c", "Khi dẫn qua bình H~2~SO~4~ đặc dư, NH~3~ và H~2~O bị giữ lại hoàn toàn (do H~2~SO~4~ hút ẩm và phản ứng tạo muối (NH~4~)~2~SO~4~).\nKhí thoát ra duy nhất là CO~2~: n_CO2 = 0,5 mol.\nThể tích khí CO~2~ thoát ra: V_CO2 = 0,5 × 24,79 = 12,395 lít.", "0,25 đ")
            ]
        }
    ]
}

# ==============================================================================
# HÀM THÊM BẢNG ĐÁP ÁN & BIỂU ĐIỂM CHẤM VÀO ĐỀ
# ==============================================================================
def add_rubric_table_to_doc(doc, exam):
    made = exam["made"]
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(14)
    p_title.paragraph_format.space_after = Pt(4)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(f"HƯỚNG DẪN CHẤM VÀ THANG ĐIỂM CHI TIẾT - MÃ ĐỀ {made}")
    set_run_style(r_title, font_size=12, bold=True, color=COLOR_PRIMARY)

    # 1. BẢNG TỔNG HỢP TRẮC NGHIỆM (PHẦN I, II, III)
    p_sub1 = doc.add_paragraph()
    p_sub1.paragraph_format.space_before = Pt(6)
    p_sub1.paragraph_format.space_after = Pt(2)
    r_sub1 = p_sub1.add_run("A. ĐÁP ÁN PHẦN TRẮC NGHIỆM (7,0 ĐIỂM)")
    set_run_style(r_sub1, font_size=11, bold=True, color=COLOR_SOLUTION)

    # Bảng Phần I: 12 câu
    p_p1_note = doc.add_paragraph()
    p_p1_note.paragraph_format.space_before = Pt(2)
    p_p1_note.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p1_note, "*1. Phần I (12 câu - mỗi câu đúng 0,25 điểm = 3,0 điểm):*")

    table_p1 = doc.add_table(rows=2, cols=12)
    table_p1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table_p1, color="90A4AE", sz="4")

    for col_idx in range(12):
        cell = table_p1.cell(0, col_idx)
        cell.width = Cm(1.5)
        set_cell_background(cell, "E3F2FD")
        set_cell_margins(cell, top=60, bottom=60, left=40, right=40)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"C{col_idx+1}")
        set_run_style(r, font_size=9.5, bold=True, color=COLOR_PRIMARY)

    for col_idx, q in enumerate(exam["part1"]):
        cell = table_p1.cell(1, col_idx)
        cell.width = Cm(1.5)
        set_cell_margins(cell, top=60, bottom=60, left=40, right=40)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(q["correct"])
        set_run_style(r, font_size=10, bold=True, color=COLOR_CORRECT)

    # Bảng Phần II: 2 câu
    p_p2_note = doc.add_paragraph()
    p_p2_note.paragraph_format.space_before = Pt(6)
    p_p2_note.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p2_note, "*2. Phần II (Trắc nghiệm Đúng/Sai - mỗi ý đúng 0,25 điểm = 2,0 điểm):*")

    table_p2 = doc.add_table(rows=3, cols=5)
    table_p2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table_p2, color="90A4AE", sz="4")

    headers_p2 = ["Câu", "Lệnh a)", "Lệnh b)", "Lệnh c)", "Lệnh d)"]
    for j, h in enumerate(headers_p2):
        cell = table_p2.cell(0, j)
        set_cell_background(cell, "E3F2FD")
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_run_style(r, font_size=9.5, bold=True, color=COLOR_PRIMARY)

    for i, q in enumerate(exam["part2"]):
        row_idx = i + 1
        cell_c = table_p2.cell(row_idx, 0)
        set_cell_margins(cell_c, top=60, bottom=60, left=60, right=60)
        p_c = cell_c.paragraphs[0]
        p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_c = p_c.add_run(f"Câu {q['idx']}")
        set_run_style(r_c, font_size=9.5, bold=True)

        for col_j, item in enumerate(q["items"]):
            cell_ans = table_p2.cell(row_idx, col_j + 1)
            set_cell_margins(cell_ans, top=60, bottom=60, left=60, right=60)
            p_ans = cell_ans.paragraphs[0]
            p_ans.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_ans = p_ans.add_run(item[1])
            set_run_style(r_ans, font_size=9.5, bold=True, color=COLOR_CORRECT)

    # Bảng Phần III: 4 câu
    p_p3_note = doc.add_paragraph()
    p_p3_note.paragraph_format.space_before = Pt(6)
    p_p3_note.paragraph_format.space_after = Pt(2)
    render_rich_text(p_p3_note, "*3. Phần III (Trả lời ngắn - mỗi câu đúng 0,5 điểm = 2,0 điểm):*")

    table_p3 = doc.add_table(rows=2, cols=4)
    table_p3.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table_p3, color="90A4AE", sz="4")

    for col_j, q in enumerate(exam["part3"]):
        cell_h = table_p3.cell(0, col_j)
        cell_h.width = Cm(4.5)
        set_cell_background(cell_h, "E3F2FD")
        set_cell_margins(cell_h, top=60, bottom=60, left=60, right=60)
        p_h = cell_h.paragraphs[0]
        p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_h = p_h.add_run(f"Câu {q['idx']}")
        set_run_style(r_h, font_size=9.5, bold=True, color=COLOR_PRIMARY)

        cell_a = table_p3.cell(1, col_j)
        cell_a.width = Cm(4.5)
        set_cell_margins(cell_a, top=60, bottom=60, left=60, right=60)
        p_a = cell_a.paragraphs[0]
        p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_a = p_a.add_run(q["ans"])
        set_run_style(r_a, font_size=10.5, bold=True, color=COLOR_CORRECT)

    # 2. BẢNG THANG ĐIỂM CHI TIẾT TỰ LUẬN (PHẦN IV - 3,0 ĐIỂM)
    p_sub2 = doc.add_paragraph()
    p_sub2.paragraph_format.space_before = Pt(12)
    p_sub2.paragraph_format.space_after = Pt(4)
    r_sub2 = p_sub2.add_run("B. THANG ĐIỂM VÀ BIỂU ĐIỂM CHẤM PHẦN TỰ LUẬN (3,0 ĐIỂM)")
    set_run_style(r_sub2, font_size=11, bold=True, color=COLOR_SOLUTION)

    total_rubric_rows = sum(len(q["rubric"]) for q in exam["part4"])
    table_tl = doc.add_table(rows=total_rubric_rows + 1, cols=3)
    table_tl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table_tl, color="78909C", sz="6")

    col_widths = [Cm(2.2), Cm(13.6), Cm(2.2)]
    headers_tl = ["Câu / Ý", "Nội dung đáp án & Yêu cầu cần đạt", "Điểm"]
    for j in range(3):
        cell = table_tl.cell(0, j)
        cell.width = col_widths[j]
        set_cell_background(cell, "DCE775" if j == 2 else "CFD8DC")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(headers_tl[j])
        set_run_style(r, font_size=10, bold=True, color=RGBColor(0x26, 0x32, 0x38))

    current_row = 1
    for q in exam["part4"]:
        q_idx = q["idx"]
        rubric_items = q["rubric"]
        
        for k, (sub_name, content_text, score_text) in enumerate(rubric_items):
            cell_c = table_tl.cell(current_row, 0)
            cell_cnt = table_tl.cell(current_row, 1)
            cell_sc = table_tl.cell(current_row, 2)

            cell_c.width = col_widths[0]
            cell_cnt.width = col_widths[1]
            cell_sc.width = col_widths[2]

            set_cell_margins(cell_c, top=80, bottom=80, left=80, right=80)
            set_cell_margins(cell_cnt, top=80, bottom=80, left=100, right=100)
            set_cell_margins(cell_sc, top=80, bottom=80, left=80, right=80)

            p_c = cell_c.paragraphs[0]
            p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if k == 0:
                r_c1 = p_c.add_run(f"Câu {q_idx}\n")
                set_run_style(r_c1, font_size=10, bold=True, color=COLOR_PRIMARY)
            r_sub = p_c.add_run(sub_name)
            set_run_style(r_sub, font_size=9.5, bold=True)

            p_cnt = cell_cnt.paragraphs[0]
            p_cnt.alignment = WD_ALIGN_PARAGRAPH.LEFT
            render_rich_text(p_cnt, content_text)

            p_sc = cell_sc.paragraphs[0]
            p_sc.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_sc = p_sc.add_run(score_text)
            set_run_style(r_sc, font_size=10, bold=True, color=COLOR_CORRECT)

            current_row += 1

    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(6)
    p_note.paragraph_format.space_after = Pt(4)
    r_note = p_note.add_run("• Ghi chú chấm bài: Học sinh giải theo cách khác đúng bản chất hóa học và ra kết quả chính xác vẫn được trọn vẹn điểm theo từng ý tương ứng. Thiếu hoặc sai điều kiện, cân bằng phản ứng trừ 0,125đ/lần.")
    set_run_style(r_note, font_size=9.5, italic=True, color=RGBColor(0x55, 0x55, 0x55))

# ==============================================================================
# HÀM XUẤT FILE WORD CHO MỖI ĐỀ
# ==============================================================================
def render_exam_doc(exam, is_teacher=False):
    doc = docx.Document()
    setup_header_footer(doc)

    de_num = exam["de_num"]
    made = exam["made"]

    # 1. Khối Tiêu đề
    build_title_block(
        doc,
        school_line="SỞ GD&ĐT THỪA THIÊN HUẾ\nTRƯỜNG THPT CHUYÊN QUỐC HỌC",
        exam_title="ĐỀ KIỂM TRA GIỮA HỌC KỲ I - NĂM HỌC 2026 - 2027",
        subject_line="MÔN: HOÁ HỌC - LỚP 11",
        num_pages="03" if not is_teacher else "04",
        duration="45 phút (không kể thời gian phát đề)",
        exam_code=made,
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; Mg = 24; S = 32; Cl = 35,5; K = 39; Ca = 40; Fe = 56; Cu = 64; Ba = 137",
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
            rubric_full = "\n".join([f"• {item[0]}: {item[1]} ({item[2]})" for item in q["rubric"]])
            add_solution_no_table(doc, rubric_full)

    # Kết thúc đề
    p_end = doc.add_paragraph()
    p_end.paragraph_format.space_before = Pt(8)
    p_end.paragraph_format.space_after = Pt(8)
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_end = p_end.add_run(f"-------------------------- HẾT ĐỀ THI MÃ {made} --------------------------")
    set_run_style(r_end, font_size=10.5, bold=True)

    # Nếu là bản Giáo Viên: Thêm bảng đáp án tổng hợp & thang điểm tự luận
    if is_teacher:
        add_rubric_table_to_doc(doc, exam)

    return doc

# ==============================================================================
# TẠO FILE RIÊNG BẢNG ĐIỂM ĐÁP ÁN TỰ LUẬN CẢ 2 ĐỀ
# ==============================================================================
def create_standalone_rubric_file():
    doc = docx.Document()
    setup_header_footer(doc)

    build_title_block(
        doc,
        school_line="SỞ GD&ĐT THỪA THIÊN HUẾ\nTRƯỜNG THPT CHUYÊN QUỐC HỌC",
        exam_title="HƯỚNG DẪN CHẤM VÀ THANG ĐIỂM CHI TIẾT ĐỀ KIỂM TRA GIỮA HỌC KỲ I",
        subject_line="MÔN: HOÁ HỌC - LỚP 11 (NĂM HỌC 2026 - 2027)",
        num_pages="02",
        duration="Áp dụng chấm thi Đề 01 (Mã 201) & Đề 02 (Mã 202)",
        exam_code="HD-CHẤM-11",
        atomic_masses="H = 1; C = 12; N = 14; O = 16; Na = 23; S = 32; Cl = 35,5; Ba = 137",
        is_teacher=True
    )

    # Đề 1
    add_rubric_table_to_doc(doc, EXAM_01)

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(16)
    p_div.paragraph_format.space_after = Pt(16)
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_div = p_div.add_run("==================================================================================")
    set_run_style(r_div, font_size=10, bold=True, color=COLOR_PRIMARY)

    # Đề 2
    add_rubric_table_to_doc(doc, EXAM_02)

    rubric_path1 = os.path.join(SUB_DIR, "Bang_Dap_An_Va_Thang_Diem_Tu_Luan_2_De_GK1_Hoa11.docx")
    rubric_path2 = os.path.join(OUT_DIR, "Bang_Dap_An_Va_Thang_Diem_Tu_Luan_2_De_GK1_Hoa11.docx")
    
    def safe_save(d, path):
        try:
            d.save(path)
            print(f"   + Đã lưu: {path}")
        except PermissionError:
            base, ext = os.path.splitext(path)
            alt = base + "_CapNhat" + ext
            d.save(alt)
            print(f"   [!] File đang mở, đã lưu sang: {alt}")

    safe_save(doc, rubric_path1)
    safe_save(doc, rubric_path2)
    print(f"Đã tạo file riêng Bảng đáp án & Thang điểm: {rubric_path1}")

# ==============================================================================
# HÀM CHÍNH
# ==============================================================================
def main():
    print("=== ĐANG KHỞI TẠO 2 ĐỀ KIỂM TRA GIỮA KÌ I HÓA 11 THEO MA TRẬN 45 PHÚT ===")
    
    exams = [EXAM_01, EXAM_02]
    
    for ex in exams:
        num = ex["de_num"]
        made = ex["made"]
        print(f"-> Tạo Đề {num:02d} (Mã {made})...")

        def safe_save(d, path):
            try:
                d.save(path)
                print(f"   + Đã lưu: {path}")
            except PermissionError:
                base, ext = os.path.splitext(path)
                alt = base + "_CapNhat" + ext
                d.save(alt)
                print(f"   [!] File đang mở, đã lưu sang: {alt}")

        # 1. Bản Học Sinh
        doc_hs = render_exam_doc(ex, is_teacher=False)
        hs_name = f"De_Kiem_Tra_GK1_Hoa_11_De{num:02d}_Ma{made}_HocSinh.docx"
        safe_save(doc_hs, os.path.join(SUB_DIR, hs_name))
        safe_save(doc_hs, os.path.join(OUT_DIR, hs_name))

        # 2. Bản Giáo Viên
        doc_gv = render_exam_doc(ex, is_teacher=True)
        gv_name = f"De_Kiem_Tra_GK1_Hoa_11_De{num:02d}_Ma{made}_GiaoVien.docx"
        safe_save(doc_gv, os.path.join(SUB_DIR, gv_name))
        safe_save(doc_gv, os.path.join(OUT_DIR, gv_name))

    # 3. Tạo file riêng Bảng điểm đáp án tự luận 2 đề
    create_standalone_rubric_file()

    print("=== HOÀN TẤT TẠO TOÀN BỘ FILE WORD HÓA 11 THEO MA TRẬN! ===")

if __name__ == "__main__":
    main()
