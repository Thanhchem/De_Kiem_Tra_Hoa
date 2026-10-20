# -*- coding: utf-8 -*-
"""
Dữ liệu Đề kiểm tra định kỳ Giữa học kỳ I - Môn Hoá học 12
Năm học 2025 - 2026 - Mã đề: 638
Biên soạn chuẩn hóa theo "FORM ĐỀ KIỂM TRA" - Thầy TRẦN VĂN THẠNH
"""

INFO = {
    "school": "TRƯỜNG TRUNG HỌC PHỔ THÔNG",
    "exam_name": "KIỂM TRA ĐỊNH KỲ GIỮA HỌC KỲ I",
    "year": "NĂM HỌC 2025 - 2026",
    "subject": "MÔN: HOÁ HỌC - KHỐI 12",
    "time_str": "Thời gian làm bài: 50 phút (Không kể thời gian phát đề)",
    "code": "638",
    "pages_str": "Đề thi gồm có 04 trang",
    "atomic_weights": "(Cho biết nguyên tử khối: H = 1; C = 12; N = 14; O = 16; Na = 23; K = 39)."
}

# ==============================================================================
# PHẦN I: TRẮC NGHIỆM NHIỀU LỰA CHỌN (18 CÂU)
# ==============================================================================

PART1_QUESTIONS = [
    {
        "id": 1,
        "content": "Dựa vào phổ khối (MS) người ta xác định được peak ứng với phân tử khối của chất hữu cơ X là 60. Chất X có thể là chất nào sau đây?",
        "options": [
            "A. Methyl acetate.",
            "B. Dimethylamine.",
            "C. Methyl formate.",
            "D. Propan-1-amine."
        ],
        "answer": "C",
        "solution": (
            "Phân tích phân tử khối (M) của các chất:\n"
            "• Methyl acetate (CH~3~COOCH~3~): M = 74.\n"
            "• Dimethylamine ((CH~3~)~2~NH): M = 45.\n"
            "• Methyl formate (HCOOCH~3~): M = 1 + 12 + 16×2 + 15 = 60.\n"
            "• Propan-1-amine (CH~3~CH~2~CH~2~NH~2~): M = 59.\n"
            "=> Chất X có peak ion phân tử m/z = 60 là methyl formate."
        )
    },
    {
        "id": 2,
        "content": "Chất ứng với công thức cấu tạo nào sau đây là amine bậc hai?",
        "options": [
            "A. CH~3~CH(NH~2~)CH~3~.",
            "B. (CH~3~)~2~NCH~3~.",
            "C. CH~3~CH~2~CH~2~NH~2~.",
            "D. CH~3~NHCH~2~CH~3~."
        ],
        "answer": "D",
        "solution": (
            "Bậc của amine được xác định bằng số gốc hydrocarbon liên kết trực tiếp với nguyên tử nitrogen:\n"
            "• CH~3~CH(NH~2~)CH~3~: Nhóm -NH~2~ liên kết với 1 gốc hydrocarbon => Amine bậc một.\n"
            "• (CH~3~)~2~NCH~3~: Nguyên tử N liên kết với 3 gốc methyl => Amine bậc ba.\n"
            "• CH~3~CH~2~CH~2~NH~2~: Amine bậc một.\n"
            "• CH~3~NHCH~2~CH~3~: Nhóm -NH- liên kết với 2 gốc hydrocarbon (methyl và ethyl) => Amine bậc hai."
        )
    },
    {
        "id": 3,
        "content": "Nhỏ dung dịch iodine (I~2~ trong dung dịch KI) vào hồ tinh bột, thu được hỗn hợp có màu",
        "options": [
            "A. vàng nhạt.",
            "B. trắng hồng.",
            "C. xanh tím.",
            "D. xanh lam."
        ],
        "answer": "C",
        "solution": (
            "Các phân tử amylose trong tinh bột có cấu trúc dạng chuỗi xoắn ốc tạo thành các khoang rỗng. "
            "Khi cho dung dịch iodine vào, các phân tử I~2~ bị giữ lại trong các khoang xoắn này tạo thành phức chất bọc có màu xanh tím đặc trưng. "
            "Khi đun nóng, cấu trúc xoắn duỗi ra làm màu xanh tím biến mất; để nguội cấu trúc xoắn tái lập và màu xanh tím xuất hiện trở lại."
        )
    },
    {
        "id": 4,
        "content": "Amine nào sau đây khi tác dụng với HNO~2~, thu được ethanol?",
        "options": [
            "A. (CH~3~)~2~NH.",
            "B. CH~3~NH~2~.",
            "C. (CH~3~)~2~CHNH~2~.",
            "D. CH~3~CH~2~NH~2~."
        ],
        "answer": "D",
        "solution": (
            "Amine no bậc một tác dụng với nitrous acid (HNO~2~) ở nhiệt độ phòng tạo thành alcohol tương ứng và giải phóng khí nitrogen:\n"
            "CH~3~CH~2~NH~2~ + HNO~2~ -> CH~3~CH~2~OH + N~2~↑ + H~2~O.\n"
            "=> Chất cần tìm là ethylamine (CH~3~CH~2~NH~2~)."
        )
    },
    {
        "id": 5,
        "content": "Sodium stearate là thành phần chính của nhiều loại xà phòng. Công thức hoá học của sodium stearate là",
        "options": [
            "A. C~17~H~33~COONa.",
            "B. C~17~H~35~COOK.",
            "C. C~15~H~31~COONa.",
            "D. C~17~H~35~COONa."
        ],
        "answer": "D",
        "solution": (
            "• Sodium stearate là muối sodium của stearic acid (C~17~H~35~COOH), có công thức là C~17~H~35~COONa.\n"
            "• A là sodium oleate (C~17~H~33~COONa).\n"
            "• B là potassium stearate (C~17~H~35~COOK).\n"
            "• C là sodium palmitate (C~15~H~31~COONa)."
        )
    },
    {
        "id": 6,
        "content": "Ethyl acetate là một dung môi hữu cơ thông dụng. Ethyl acetate có thể điều chế bằng phản ứng ester hoá giữa",
        "options": [
            "A. acetic acid và methyl alcohol.",
            "B. acetic acid và ethyl alcohol.",
            "C. formic acid và methyl alcohol.",
            "D. formic acid và ethyl alcohol."
        ],
        "answer": "B",
        "solution": (
            "Phản ứng điều chế ethyl acetate (CH~3~COOC~2~H~5~):\n"
            "CH~3~COOH + C~2~H~5~OH <=(H~2~SO~4~ đặc, t°)=> CH~3~COOC~2~H~5~ + H~2~O.\n"
            "=> Phản ứng ester hoá giữa acetic acid và ethyl alcohol."
        )
    },
    {
        "id": 7,
        "content": "Chất hữu cơ X có vị ngọt, không mùi, tồn tại ở dạng lỏng màu trắng, tan trong nước và được dùng làm chất tạo ngọt, giữ ẩm và tạo độ bóng trong thực phẩm như kẹo, bánh. X còn có nhiều ứng dụng khác trong y học, mỹ phẩm và công nghiệp nhờ tác dụng nhuận tràng, làm thuốc tẩy và phụ gia. X được điều chế bằng phản ứng hydrogen hoá glucose. X là",
        "options": [
            "A. ethanol.",
            "B. sorbitol.",
            "C. lactic acid.",
            "D. gluconic acid."
        ],
        "answer": "B",
        "solution": (
            "Khi hydrogen hoá glucose (khử bằng H~2~ có xúc tác Ni, t°), nhóm aldehyde (-CHO) bị khử thành nhóm alcohol bậc một (-CH~2~OH), "
            "thu được polyalcohol sáu chức là sorbitol:\n"
            "CH~2~OH[CHOH]~4~CHO + H~2~ -(Ni, t°)-> CH~2~OH[CHOH]~4~CH~2~OH (sorbitol)."
        )
    },
    {
        "id": 8,
        "content": "Chất nào sau đây thuộc loại disaccharide?",
        "options": [
            "A. Glucose.",
            "B. Fructose.",
            "C. Maltose.",
            "D. Cellulose."
        ],
        "answer": "C",
        "solution": (
            "Phân loại carbohydrate:\n"
            "• Monosaccharide: Glucose, Fructose (không bị thuỷ phân).\n"
            "• Disaccharide: Saccharose, Maltose (thuỷ phân tạo 2 phân tử monosaccharide).\n"
            "• Polysaccharide: Tinh bột, Cellulose."
        )
    },
    {
        "id": 9,
        "content": "Sản phẩm của quá trình khử acetaldehyde (CH~3~CHO) bởi LiAlH~4~ là",
        "options": [
            "A. CH~3~COOH.",
            "B. CH~3~CH~3~.",
            "C. HOCH~2~CH~2~OH.",
            "D. CH~3~CH~2~OH."
        ],
        "answer": "D",
        "solution": (
            "LiAlH~4~ (Lithium aluminium hydride) là tác nhân khử mạnh, khử hợp chất aldehyde thành alcohol bậc một tương ứng:\n"
            "CH~3~CHO + 2[H] -(LiAlH~4~)-> CH~3~CH~2~OH (ethyl alcohol)."
        )
    },
    {
        "id": 10,
        "content": "Nhúng quỳ tím vào dung dịch nào sau đây, quỳ tím chuyển sang màu xanh?",
        "options": [
            "A. Glycine.",
            "B. Lysine.",
            "C. Alanine.",
            "D. Valine."
        ],
        "answer": "B",
        "solution": (
            "Xét số lượng nhóm amino (-NH~2~) và carboxyl (-COOH) trong phân tử amino acid:\n"
            "• Glycine, Alanine, Valine: Có 1 nhóm -NH~2~ và 1 nhóm -COOH => Dung dịch có môi trường gần như trung tính, không làm đổi màu quỳ tím.\n"
            "• Lysine (H~2~N-[CH~2~]~4~-CH(NH~2~)-COOH): Có 2 nhóm -NH~2~ và 1 nhóm -COOH => Dung dịch có tính base (pH > 7), làm quỳ tím chuyển sang màu xanh."
        )
    },
    {
        "id": 11,
        "content": "Giấm ăn là một gia vị thường gặp trong gia đình. Giấm ăn là dung dịch chứa chất X có nồng độ từ 4–5%. Chất X trong giấm ăn là",
        "options": [
            "A. CH~3~COOH.",
            "B. H~2~NCH~2~COOH.",
            "C. CH~3~CH(OH)COOH.",
            "D. HCOOH."
        ],
        "answer": "A",
        "solution": (
            "Giấm ăn là dung dịch acetic acid (CH~3~COOH) có nồng độ từ 4% đến 5%, được sản xuất phổ biến bằng phương pháp lên men giấm từ dung dịch rượu ethylic loãng."
        )
    },
    {
        "id": 12,
        "content": "Phản ứng hoá học đặc trưng cho ester là",
        "options": [
            "A. phản ứng trùng hợp.",
            "B. phản ứng oxi hoá.",
            "C. phản ứng thuỷ phân.",
            "D. phản ứng cộng."
        ],
        "answer": "C",
        "solution": (
            "Liên kết ester C-O trong nhóm chức -COO- phân cực và dễ bị cắt đứt dưới tác dụng của nước trong môi trường acid (thuận nghịch) "
            "hoặc base (phản ứng xà phòng hoá một chiều), do đó phản ứng hoá học đặc trưng nhất của ester là phản ứng thuỷ phân."
        )
    },
    {
        "id": 13,
        "content": "Cho các công thức cấu tạo sau:\n(Cấu tạo I, II, III, IV trong hình vẽ)\nCấu tạo nào đúng với α-glucose?",
        "image": "image1.png",
        "image_width_cm": 15.5,
        "options": [
            "A. (IV).",
            "B. (III).",
            "C. (II).",
            "D. (I)."
        ],
        "answer": "C",
        "solution": (
            "Phân tích cấu trúc vòng Haworth:\n"
            "• Vòng glucopyranose (6 cạnh): C1 là carbon hemiacetal (ở bên phải nguyên tử O của vòng).\n"
            "  - Ở α-D-glucopyranose: Nhóm -OH hemiacetal tại C1 hướng xuống dưới mặt phẳng vòng (cùng phía với -OH ở C2 và C4).\n"
            "    => Cấu trúc (II) có -OH tại C1 hướng xuống dưới mặt phẳng vòng chính là α-glucose.\n"
            "  - Ở β-D-glucopyranose (I): Nhóm -OH hemiacetal tại C1 hướng lên trên mặt phẳng vòng.\n"
            "• Cấu trúc (III) và (IV) là dạng vòng furanose (5 cạnh) của fructose.\n"
            "=> Đáp án đúng là C. (II)."
        )
    },
    {
        "id": 14,
        "content": "Chất ứng với công thức hoá học nào sau đây là ester?",
        "options": [
            "A. CH~3~COOCH~3~.",
            "B. CH~3~-O-CH~2~-CHO.",
            "C. CH~3~-COONH~3~-CH~3~.",
            "D. CH~3~COOH."
        ],
        "answer": "A",
        "solution": (
            "• CH~3~COOCH~3~ (methyl acetate) là ester tạo bởi acetic acid và methyl alcohol.\n"
            "• B là ether - aldehyde.\n"
            "• C là muối methylammonium acetate.\n"
            "• D là acetic acid (carboxylic acid)."
        )
    },
    {
        "id": 15,
        "content": "Thuốc thử Tollens là hoá chất thường sử dụng để nhận biết nhóm chức nào sau đây?",
        "options": [
            "A. nhóm -CHO (aldehyde).",
            "B. nhóm -COOH (carboxyl).",
            "C. nhóm -OH (phenol).",
            "D. nhóm -OH (alcohol)."
        ],
        "answer": "A",
        "solution": (
            "Thuốc thử Tollens là phức chất [Ag(NH~3~)~2~]OH trong môi trường kiềm ammonia, có tính oxi hoá chọn lọc đối với nhóm chức aldehyde (-CHO), "
            "khử ion Ag^+ thành Ag kim loại bám sáng bóng vào thành ống nghiệm (phản ứng tráng bạc)."
        )
    },
    {
        "id": 16,
        "content": "Bơ nhân tạo có thành phần chính là ...(1).... được sản xuất từ quá trình ... (2)... các chất béo lỏng. Nội dung phù hợp điền vào chỗ trống (1), (2) là",
        "options": [
            "A. (1)-triglyceride không no; (2)-hydrogen hoá.",
            "B. (1)-triglyceride không no; (2)-oxi hoá.",
            "C. (1)-triglyceride no; (2)-hydrogen hoá.",
            "D. (1)-triglyceride no; (2)-oxi hoá."
        ],
        "answer": "C",
        "solution": (
            "Quá trình hydrogen hoá chất béo lỏng (dầu thực vật chứa triglyceride không no như triolein) với xúc tác Ni, đun nóng "
            "sẽ chuyển hoá thành chất béo rắn (chứa triglyceride no như tristearin), được dùng làm thành phần chính của bơ nhân tạo và shortening:\n"
            "=> (1)-triglyceride no; (2)-hydrogen hoá."
        )
    },
    {
        "id": 17,
        "content": "Glycerol là chất hữu cơ có nhiều ứng dụng như làm dung môi, điều chế thuốc nổ,... Công thức hoá học của glycerol là",
        "options": [
            "A. C~3~H~5~OH.",
            "B. C~2~H~5~OH.",
            "C. C~2~H~4~(OH)~2~.",
            "D. C~3~H~5~(OH)~3~."
        ],
        "answer": "D",
        "solution": (
            "Glycerol (propane-1,2,3-triol) là alcohol no ba chức có công thức hoá học là C~3~H~5~(OH)~3~ (hay C~3~H~8~O~3~)."
        )
    },
    {
        "id": 18,
        "content": "Khi sử dụng các chất giặt rửa tổng hợp cùng với nước để loại bỏ vết bẩn là dầu mỡ bám trên bề mặt vật rắn, chất giặt rửa đưa ...(1)... thâm nhập vào vết dầu mỡ, thiết lập liên kết với dầu mỡ và phân cắt mảng dầu mỡ thành các phần nhỏ để ...(2)... nước. Nội dung phù hợp điền vào chỗ trống (1), (2) là",
        "options": [
            "A. (1)-nhóm ưa nước; (2)-phân tán vào.",
            "B. (1)-gốc hydrocarbon; (2)-phân tán vào.",
            "C. (1)-gốc hydrocarbon; (2)-bay ra khỏi.",
            "D. (1)-nhóm ưa nước; (2)-bay ra khỏi."
        ],
        "answer": "B",
        "solution": (
            "Cơ chế tẩy rửa của xà phòng và chất giặt rửa tổng hợp: Phân tử gồm đuôi kị nước (gốc hydrocarbon dài) ưa dầu mỡ và đầu ưa nước phân cực. "
            "Khi cọ rửa, gốc hydrocarbon kị nước thâm nhập gắn vào vết dầu mỡ, còn đầu ưa nước quay ra ngoài tương tác với nước. "
            "Nhờ đó vết dầu mỡ bị chia cắt thành các hạt micelle kích thước siêu nhỏ để phân tán lơ lửng vào nước rồi bị dòng nước cuốn trôi đi."
        )
    }
]

# ==============================================================================
# PHẦN II: TRẮC NGHIỆM ĐÚNG - SAI (4 CÂU)
# ==============================================================================

PART2_QUESTIONS = [
    {
        "id": 1,
        "context": "Fructose có nhiều trong mật ong và là chất tạo nên vị ngọt đậm đà của mật ong. Trong dung dịch nước, fructose tồn tại chủ yếu ở dạng mạch vòng (5 hoặc 6 cạnh) và một lượng nhỏ tồn tại ở dạng mạch hở. Các dạng đó chuyển hoá lẫn nhau và thiết lập các trạng thái cân bằng hoá học.",
        "items": [
            {
                "id": "a",
                "content": "Fructose thuộc loại monosaccharide.",
                "answer": True,
                "reason": "Đúng. Fructose là một monosaccharide (polyhydroxy ketone C~6~H~12~O~6~) không bị thuỷ phân."
            },
            {
                "id": "b",
                "content": "Các cấu tạo của fructose (mạch vòng và mạch hở) có cùng số nhóm hydroxy (-OH).",
                "answer": True,
                "reason": "Đúng. Dạng mạch hở có 5 nhóm -OH (tại C1, C3, C4, C5, C6) và 1 nhóm C=O (tại C2). Khi đóng vòng, liên kết giữa -OH (ở C5 hoặc C6) với C=O tạo ra nhóm -OH hemiketal mới tại C2, do đó dạng mạch vòng vẫn có đúng 5 nhóm -OH."
            },
            {
                "id": "c",
                "content": "Khi nhóm hydroxy (-OH) ở carbon số 5 kết hợp với nhóm C=O dạng mạch hở sẽ hình thành fructose có vòng 5 cạnh.",
                "answer": True,
                "reason": "Đúng. Nhóm -OH ở C5 phản ứng nội phân tử với nhóm C=O ở C2 khép thành vòng furanose 5 cạnh gồm 4 nguyên tử carbon (C2, C3, C4, C5) và 1 nguyên tử oxygen."
            },
            {
                "id": "d",
                "content": "Khi phản ứng với thuốc thử Tollens, các dạng mạch hở của fructose sẽ chuyển dịch cân bằng theo hướng tạo mạch vòng.",
                "answer": False,
                "reason": "Sai. Trong môi trường kiềm của thuốc thử Tollens, fructose đồng phân hoá thành glucose (dạng mạch hở có nhóm -CHO). Dạng mạch hở này bị tiêu hao do phản ứng oxi hoá tráng bạc, do đó cân bằng chuyển hoá giữa dạng mạch vòng và mạch hở phải chuyển dịch theo hướng mở vòng tạo thêm dạng mạch hở (theo nguyên lí Le Chatelier)."
            }
        ]
    },
    {
        "id": 2,
        "context": "Diethylamine thường dùng làm chất chống ăn mòn và dùng trong công nghiệp sản xuất cao su, nhựa, chất nhuộm và trong ngành dược. Một đề xuất tổng hợp diethylamine từ ethylene được thực hiện theo sơ đồ sau:",
        "image": "image2.png",
        "image_width_cm": 15.0,
        "items": [
            {
                "id": "a",
                "content": "Các phản ứng (1), (2), (3) đều là các phản ứng oxi hoá - khử.",
                "answer": False,
                "reason": "Sai. Phản ứng (1) là phản ứng cộng electrophile (CH~2~=CH~2~ + HCl -> CH~3~CH~2~Cl) có thay đổi số oxi hoá của các nguyên tử carbon nên là phản ứng oxi hoá - khử. Tuy nhiên, phản ứng (2) và (3) là các phản ứng thế nucleophile của dẫn xuất halogen bằng ammonia và amine, số oxi hoá của tất cả các nguyên tố C, H, N, Cl đều không đổi, không phải là phản ứng oxi hoá - khử."
            },
            {
                "id": "b",
                "content": "Trong phản ứng (2), để sản phẩm hữu cơ chủ yếu là CH~3~CH~2~NH~2~ thì CH~3~CH~2~Cl phải lấy dư so với NH~3~.",
                "answer": False,
                "reason": "Sai. Nếu CH~3~CH~2~Cl dư thì amine bậc một CH~3~CH~2~NH~2~ sinh ra sẽ tiếp tục phản ứng với CH~3~CH~2~Cl dư tạo amine bậc hai, bậc ba và muối ammonium bậc bốn. Để thu được sản phẩm chủ yếu là amine bậc một, bắt buộc phải dùng NH~3~ rất dư so với ethyl chloride."
            },
            {
                "id": "c",
                "content": "Nếu tỉ lệ mol tham gia phản ứng của CH~3~CH~2~Cl với NH~3~ là 1 : 2 thì NH~4~Cl là sản phẩm còn lại của phản ứng (2).",
                "answer": True,
                "reason": "Đúng. Phương trình hoá học: CH~3~CH~2~Cl + 2NH~3~ -> CH~3~CH~2~NH~2~ + NH~4~Cl. Với tỉ lệ mol phản ứng 1 : 2, phân tử NH~3~ thứ hai đóng vai trò base thu nhận proton HCl sinh ra tạo muối NH~4~Cl."
            },
            {
                "id": "d",
                "content": "Các phản ứng (2) và (3) đều có sự thay thế nguyên tử H liên kết với N bằng gốc hydrocarbon.",
                "answer": True,
                "reason": "Đúng. Ở phản ứng (2), một nguyên tử H trong NH~3~ được thay thế bằng gốc ethyl (CH~3~CH~2~-). Ở phản ứng (3), một nguyên tử H trên nhóm -NH~2~ của ethylamine tiếp tục được thay thế bằng gốc ethyl để tạo amine bậc hai (CH~3~CH~2~)~2~NH."
            }
        ]
    },
    {
        "id": 3,
        "context": "Phản ứng tổng hợp ethyl acetate từ ethyl alcohol và acetic acid theo phương trình hoá học sau:",
        "image": "image3.png",
        "image_width_cm": 14.5,
        "items": [
            {
                "id": "a",
                "content": "Phản ứng tổng hợp ester từ carboxylic acid và alcohol là phản ứng ester hoá.",
                "answer": True,
                "reason": "Đúng. Theo định nghĩa, phản ứng tạo ester từ carboxylic acid và alcohol có xúc tác acid được gọi là phản ứng ester hoá."
            },
            {
                "id": "b",
                "content": "Trong phản ứng trên, nguyên tử oxygen trong phân tử H~2~O có nguồn gốc từ ethyl alcohol.",
                "answer": False,
                "reason": "Sai. Nghiên cứu bằng phương pháp đồng vị phóng xạ (^18^O) chứng minh nhóm -OH bị cắt ra từ phân tử carboxylic acid, còn nguyên tử H bị tách ra từ nhóm -OH của alcohol. Do đó, nguyên tử oxygen trong phân tử nước H~2~O có nguồn gốc từ acetic acid, không phải từ alcohol."
            },
            {
                "id": "c",
                "content": "Với cùng tỉ lệ chất phản ứng, hiệu suất phản ứng tổng hợp ethyl acetate từ nguyên liệu là giấm ăn (dung dịch CH~3~COOH 5%) và rượu trắng (cồn ethanol khoảng 30 - 40°) cao hơn so với từ nguyên liệu là acetic acid và ethyl alcohol nguyên chất.",
                "answer": False,
                "reason": "Sai. Phản ứng ester hoá là phản ứng thuận nghịch. Khi dùng giấm ăn và rượu trắng, lượng nước có sẵn trong hệ phản ứng là rất lớn (> 80%), nồng độ chất phản ứng rất loãng, làm cân bằng chuyển dịch mạnh theo chiều nghịch (chiều thuỷ phân), do đó hiệu suất este hoá sẽ rất thấp so với khi dùng acid và alcohol nguyên chất kết hợp H~2~SO~4~ đặc hút nước."
            },
            {
                "id": "d",
                "content": "Các hợp chất ester đều tổng hợp được từ nguyên liệu ban đầu là carboxylic acid và alcohol (xúc tác H~2~SO~4~ đặc, t°).",
                "answer": False,
                "reason": "Sai. Một số ester không thể điều chế trực tiếp từ carboxylic acid và alcohol tương ứng, ví dụ: phenyl acetate (CH~3~COOC~6~H~5~) phải điều chế từ phenol và acetic anhydride; vinyl acetate (CH~3~COOCH=CH~2~) điều chế từ acetic acid tác dụng với acetylene."
            }
        ]
    },
    {
        "id": 4,
        "context": "Trong dung dịch nước, amine tác dụng với nước theo cân bằng hoá học sau:\nRNH~2~(aq) + H~2~O(l) <=> RNH~3~^+(aq) + OH^-(aq)\nVới cùng nồng độ mol/L, dung dịch methylamine làm quỳ tím chuyển sang màu xanh và dung dịch aniline không làm quỳ tím chuyển màu.",
        "items": [
            {
                "id": "a",
                "content": "Methylamine và aniline đều là các amine bậc một.",
                "answer": True,
                "reason": "Đúng. Methylamine (CH~3~NH~2~) và aniline (C~6~H~5~NH~2~) đều có 1 gốc hydrocarbon liên kết trực tiếp với nhóm -NH~2~, đều là amine bậc một."
            },
            {
                "id": "b",
                "content": "Do các amine tác dụng được với nước nên các amine có độ tan trong nước lớn.",
                "answer": False,
                "reason": "Sai. Độ tan của amine trong nước phụ thuộc chủ yếu vào kích thước và bản chất của gốc hydrocarbon. Các amine đầu dãy có gốc alkyl nhỏ tan tốt nhờ tạo liên kết hydrogen với nước. Các amine có gốc hydrocarbon lớn hoặc amine thơm như aniline (C~6~H~5~NH~2~) rất ít tan trong nước do gốc phenyl kị nước chiếm ưu thế."
            },
            {
                "id": "c",
                "content": "Tính base của amine là do trên nguyên tử nitrogen có cặp electron hoá trị riêng có khả năng nhận H^+.",
                "answer": True,
                "reason": "Đúng. Theo thuyết Brønsted - Lowry, nguyên tử N trong phân tử amine còn 1 cặp electron tự do chưa liên kết, có khả năng dùng cặp electron này để liên kết với proton H^+ của acid hoặc nước, thể hiện tính base."
            },
            {
                "id": "d",
                "content": "Dựa vào sự đổi màu của quỳ tím trong điều kiện trên, do đó kết luận rằng lực base của methylamine lớn hơn lực base của aniline.",
                "answer": True,
                "reason": "Đúng. Ở cùng nồng độ mol/L, dung dịch methylamine làm quỳ tím chuyển màu xanh chứng tỏ nồng độ OH^- sinh ra lớn (tính base mạnh); còn dung dịch aniline không làm quỳ tím đổi màu chứng tỏ nồng độ OH^- rất nhỏ (tính base rất yếu do hiệu ứng hút electron của vòng benzene). Vậy lực base của methylamine lớn hơn aniline."
            }
        ]
    }
]

# ==============================================================================
# PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (6 CÂU)
# ==============================================================================

PART3_QUESTIONS = [
    {
        "id": 1,
        "content": (
            "Chỉ số xà phòng hóa (Saponification value, viết tắt là SAP) là lượng KOH cần thiết (tính bằng milligam) "
            "để phản ứng hết với 1,00 gam dầu, mỡ theo phản ứng xà phòng hóa và trung hòa hết carboxylic acid tự do có trong dầu, mỡ. "
            "Một loại dầu dừa có giá trị SAP là 255. Tính khối lượng của NaOH (theo gam) cần dùng để xà phòng hóa hết 200 gam dầu dừa. "
            "(Không làm tròn các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng phần mười)."
        ),
        "answer": "36,4",
        "solution": (
            "• Cứ 1,00 gam dầu dừa cần 255 mg KOH = 0,255 g KOH.\n"
            "• Khối lượng KOH cần để xà phòng hoá hoàn toàn 200 g dầu dừa là:\n"
            "  m(KOH) = 200 × 0,255 = 51 gam.\n"
            "• Số mol KOH phản ứng:\n"
            "  n(KOH) = 51 / 56 mol.\n"
            "• Do NaOH và KOH đều có bản chất phản ứng tương đương (1 mol nhóm este hoặc acid tự do phản ứng với 1 mol OH^-):\n"
            "  n(NaOH) = n(KOH) = 51 / 56 mol.\n"
            "• Khối lượng NaOH cần dùng là:\n"
            "  m(NaOH) = n(NaOH) × 40 = (51 / 56) × 40 = 2040 / 56 ≈ 36,42857 gam.\n"
            "=> Làm tròn đến hàng phần mười: 36,4."
        )
    },
    {
        "id": 2,
        "content": (
            "Điểm đẳng điện (pI) phụ thuộc vào cấu tạo của amino acid, peptide. Các amino acid, peptide có nhiều nhóm -COOH thì có điểm đẳng điện thấp và các amino acid có nhiều nhóm -NH~2~ thì có điểm đẳng điện cao. Với các amino acid có số nhóm -NH~2~ bằng số nhóm -COOH thì điểm đẳng điện gần với giá trị pH bằng 6,0.\n"
            "(Bảng điểm đẳng điện pI: Glycine: 5,97; Alanine: 6,00; Lysine: 9,74; Glutamic acid: 3,22; Aspartic acid: 2,77; Phenylalanine: 5,48)\n"
            "Trong một dung dịch có pH = 4,5 và giữ ổn định. Một sinh viên nhỏ một dung dịch mẫu chứa hỗn hợp 6 amino acid là glycine, alanine, phenylalanine, lysine, aspartic acid và glutamic acid. Có bao nhiêu amino acid điện di về phía cực dương?"
        ),
        "table_data": [
            ["Amino acid", "Công thức", "Kí hiệu", "pI"],
            ["Glycine", "H~2~NCH~2~COOH", "Gly", "5,97"],
            ["Alanine", "H~2~NC~2~H~4~COOH", "Ala", "6,00"],
            ["Lysine", "(H~2~N)~2~C~5~H~9~COOH", "Lys", "9,74"],
            ["Glutamic acid", "H~2~NC~3~H~5~(COOH)~2~", "Glu", "3,22"],
            ["Aspartic acid", "H~2~NC~2~H~3~(COOH)~2~", "Asp", "2,77"],
            ["Phenylalanine", "H~2~NC~8~H~9~COOH", "Phe", "5,48"]
        ],
        "answer": "2",
        "solution": (
            "Quy tắc điện di của amino acid:\n"
            "• Cực dương hút các ion mang điện tích âm (anion).\n"
            "• Amino acid tồn tại ở dạng anion (tổng điện tích âm) khi môi trường có pH > pI (môi trường kiềm hơn điểm đẳng điện, nhóm -COOH phân li H^+ thành -COO^-).\n"
            "• Ngược lại, khi pH < pI, amino acid nhận proton H^+ tạo cation mang điện tích dương và di chuyển về cực âm.\n"
            "Xét dung dịch đệm có pH = 4,5:\n"
            "• Glycine (pI = 5,97 > 4,5): pH < pI => Dạng cation (+), điện di về cực âm.\n"
            "• Alanine (pI = 6,00 > 4,5): pH < pI => Dạng cation (+), điện di về cực âm.\n"
            "• Lysine (pI = 9,74 > 4,5): pH < pI => Dạng cation (+), điện di về cực âm.\n"
            "• Phenylalanine (pI = 5,48 > 4,5): pH < pI => Dạng cation (+), điện di về cực âm.\n"
            "• Glutamic acid (pI = 3,22 < 4,5): pH > pI => Dạng anion (-), điện di về cực dương.\n"
            "• Aspartic acid (pI = 2,77 < 4,5): pH > pI => Dạng anion (-), điện di về cực dương.\n"
            "=> Có đúng 2 amino acid (Glutamic acid và Aspartic acid) điện di về phía cực dương."
        )
    },
    {
        "id": 3,
        "content": "Chuẩn độ 10 mL dung dịch methylamine (CH~3~NH~2~) nồng độ a mol/L bằng dung dịch H~2~SO~4~ 0,10 M với chất chỉ thị phù hợp, khi phản ứng đạt đến điểm tương đương thấy hết 12,5 mL dung dịch H~2~SO~4~. Xác định giá trị của a.",
        "answer": "0,25",
        "solution": (
            "• Phương trình hoá học phản ứng chuẩn độ:\n"
            "  2CH~3~NH~2~ + H~2~SO~4~ -> (CH~3~NH~3~)~2~SO~4~\n"
            "• Số mol H~2~SO~4~ đã phản ứng:\n"
            "  n(H~2~SO~4~) = C~M~ × V = 0,10 × 0,0125 = 0,00125 mol.\n"
            "• Theo phương trình tỉ lệ mol:\n"
            "  n(CH~3~NH~2~) = 2 × n(H~2~SO~4~) = 2 × 0,00125 = 0,0025 mol.\n"
            "• Nồng độ mol của dung dịch methylamine:\n"
            "  a = n / V = 0,0025 / 0,010 = 0,25 M.\n"
            "=> Giá trị của a là 0,25."
        )
    },
    {
        "id": 4,
        "content": (
            "Cho các công thức cấu tạo (cột I) và tên gọi của các chất (cột II) sau:\n"
            "Hãy thiết lập bộ ba số liên tiếp nhau (được số hoá ở cột II) ứng với tên thay thế của các chất (a), (b) và (c) (ở cột I)."
        ),
        "table_data": [
            ["Cột I", "", "Cột II", ""],
            ["(a)", "CH~3~CH(NH~2~)CH~3~", "(1)", "Acetic acid"],
            ["(b)", "CH~3~-CH(NH~2~)-COOH", "(2)", "2-aminopropanoic acid"],
            ["(c)", "CH~3~-COOH", "(3)", "Ethanoic acid"],
            ["", "", "(4)", "α-aminopropionic acid"],
            ["", "", "(5)", "Isopropylamine"],
            ["", "", "(6)", "Propan-2-amine"]
        ],
        "answer": "623",
        "solution": (
            "Xác định tên thay thế (IUPAC) của từng chất:\n"
            "• Chất (a): CH~3~-CH(NH~2~)-CH~3~ có mạch chính 3 carbon, nhóm amino ở C2 => Tên thay thế là propan-2-amine ứng với số (6). (Lưu ý: Isopropylamine (5) là tên gốc - chức).\n"
            "• Chất (b): CH~3~-CH(NH~2~)-COOH có mạch chính 3 carbon, C1 là carbon của nhóm -COOH, nhóm amino ở C2 => Tên thay thế là 2-aminopropanoic acid ứng với số (2). (Lưu ý: α-aminopropionic acid (4) là tên bán hệ thống).\n"
            "• Chất (c): CH~3~-COOH có 2 nguyên tử carbon => Tên thay thế là ethanoic acid ứng với số (3). (Lưu ý: Acetic acid (1) là tên thông thường).\n"
            "=> Bộ ba số liên tiếp ghép lại theo thứ tự (a) - (b) - (c) là: 623."
        )
    },
    {
        "id": 5,
        "content": (
            "Xăng sinh học (biogasoline) là hỗn hợp xăng thông thường và cồn sinh học (bioethanol). "
            "Bioethanol là ethanol được sản xuất bằng phương pháp sinh hoá thông qua sự lên men các sản phẩm hữu cơ như tinh bột, cellulose. "
            "Giả thiết quá trình sản xuất bioethanol từ cellulose gồm hai giai đoạn với hiệu suất như sơ đồ:\n"
            "Hiệu suất giai đoạn (1) là 69% và giai đoạn (2) là 55%.\n"
            "Từ 1,0 tấn bã mía (chứa 46% cellulose) sản xuất được V m^3^ xăng E5 (chứa 5% ethanol về thể tích). "
            "Giá trị của V bằng bao nhiêu? (Không làm tròn các phép tính trung gian, chỉ làm tròn kết quả cuối cùng đến hàng phần trăm)."
        ),
        "image": "image4.png",
        "image_width_cm": 15.0,
        "answer": "2,51",
        "solution": (
            "• Khối lượng cellulose nguyên chất trong 1,0 tấn bã mía:\n"
            "  m(cellulose) = 1,0 × 46% = 0,46 tấn = 460 kg.\n"
            "• Sơ đồ chuyển hoá từ cellulose sang ethanol:\n"
            "  (C~6~H~10~O~5~)~n~ -(1)-> n C~6~H~12~O~6~ -(2)-> 2n C~2~H~5~OH\n"
            "  162 kg mắt xích --------------------------> 2 × 46 = 92 kg ethanol.\n"
            "• Hiệu suất tổng cộng của cả hai giai đoạn:\n"
            "  H = H~1~ × H~2~ = 69% × 55% = 0,69 × 0,55 = 0,3795 (tức 37,95%).\n"
            "• Khối lượng ethanol thực tế thu được:\n"
            "  m(ethanol) = (460 / 162) × 92 × 0,3795 = 99,138518... kg.\n"
            "• Thể tích ethanol thực tế (D = 0,79 kg/L):\n"
            "  V(ethanol) = m / D = 99,138518... / 0,79 = 125,491795... L = 0,1254918... m^3^.\n"
            "• Xăng E5 chứa 5% thể tích bioethanol, do đó thể tích xăng E5 sản xuất được là:\n"
            "  V(xăng E5) = V(ethanol) / 5% = 125,491795... / 0,05 = 2509,8359... L ≈ 2,5098... m^3^.\n"
            "=> Làm tròn đến hàng phần trăm: 2,51."
        )
    },
    {
        "id": 6,
        "content": "Amino acid X có công thức phân tử là C~4~H~9~O~2~N. Hãy cho biết X có bao nhiêu công thức cấu tạo?",
        "answer": "5",
        "solution": (
            "Amino acid no, đơn chức, mạch hở chứa 1 nhóm -NH~2~ và 1 nhóm -COOH có công thức tổng quát là C~n~H~2n+1~NO~2~.\n"
            "Với CTPT C~4~H~9~O~2~N, amino acid có dạng C~3~H~6~(NH~2~)COOH.\n"
            "Xét các khung carbon của gốc carboxylic acid 4 carbon:\n"
            "1. Khung mạch thẳng (butanoic acid: C-C-C-COOH, C1 là nhóm -COOH):\n"
            "   • Đính nhóm -NH~2~ ở C2: CH~3~-CH~2~-CH(NH~2~)-COOH (2-aminobutanoic acid).\n"
            "   • Đính nhóm -NH~2~ ở C3: CH~3~-CH(NH~2~)-CH~2~-COOH (3-aminobutanoic acid).\n"
            "   • Đính nhóm -NH~2~ ở C4: H~2~N-CH~2~-CH~2~-CH~2~-COOH (4-aminobutanoic acid / GABA).\n"
            "2. Khung mạch nhánh (2-methylpropanoic acid: C-CH(CH~3~)-COOH):\n"
            "   • Đính nhóm -NH~2~ ở C2: (CH~3~)~2~C(NH~2~)-COOH (2-amino-2-methylpropanoic acid).\n"
            "   • Đính nhóm -NH~2~ ở C3: H~2~N-CH~2~-CH(CH~3~)-COOH (3-amino-2-methylpropanoic acid).\n"
            "=> Tổng cộng có 5 công thức cấu tạo amino acid phù hợp."
        )
    }
]
