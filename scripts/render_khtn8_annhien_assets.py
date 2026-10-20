# -*- coding: utf-8 -*-
"""
Vẽ sơ đồ truyền máu ở người (ABO) bằng PIL/Pillow chất lượng cao.
"""
import os
import math
from PIL import Image, ImageDraw, ImageFont

OUT_DIR = r"c:\Antigravity_Thanh\San_Pham\images_khtn8"
os.makedirs(OUT_DIR, exist_ok=True)
IMG_PATH = os.path.join(OUT_DIR, "so_do_truyen_mau.png")

# Kích thước ảnh 1000 x 700 px (độ nét cao)
W, H = 1000, 700
img = Image.new("RGBA", (W, H), (255, 255, 255, 255))
draw = ImageDraw.Draw(img)

# Tọa độ 4 nút (hình thoi)
# O ở bên trái, AB ở bên phải, A ở trên, B ở dưới
# Hoặc O ở giữa trên dưới? Trong ảnh trang 5:
#       A
#       ^
#       |
# O ---------> AB
#       |
#       v
#       B
# Hãy xem chính xác ảnh trang 5:
#   O ở bên trái
#   A ở trên
#   B ở dưới
#   AB ở bên phải
# Mũi tên:
#   O -> O (mũi tên vòng)
#   O -> A
#   O -> B
#   O -> AB (mũi tên ngang qua giữa)
#   A -> A (vòng)
#   A -> AB
#   B -> B (vòng)
#   B -> AB
#   AB -> AB (vòng)

# Tọa độ các nút
nodes = {
    'O': (200, 320),
    'A': (500, 160),
    'B': (500, 480),
    'AB': (800, 320)
}
R = 45 # Bán kính nút

# Load font
try:
    font_node = ImageFont.truetype("arialbd.ttf", 36)
    font_label = ImageFont.truetype("ariali.ttf", 26)
except:
    font_node = ImageFont.load_default()
    font_label = ImageFont.load_default()

def draw_arrow(draw, start, end, color=(211, 47, 47, 255), width=3, arrow_size=16):
    """Vẽ đường thẳng có mũi tên ở điểm cuối."""
    x1, y1 = start
    x2, y2 = end
    dx = x2 - x1
    dy = y2 - y1
    dist = math.hypot(dx, dy)
    if dist == 0:
        return
    # Rút ngắn đoạn thẳng để không chạm tâm mà chạm mép vòng tròn
    ux = dx / dist
    uy = dy / dist
    sx = x1 + ux * (R + 4)
    sy = y1 + uy * (R + 4)
    ex = x2 - ux * (R + 6)
    ey = y2 - uy * (R + 6)
    
    draw.line([(sx, sy), (ex, ey)], fill=color, width=width)
    
    # Vẽ đầu mũi tên
    angle = math.atan2(ey - sy, ex - sx)
    p1 = (ex - arrow_size * math.cos(angle - math.pi / 6),
          ey - arrow_size * math.sin(angle - math.pi / 6))
    p2 = (ex - arrow_size * math.cos(angle + math.pi / 6),
          ey - arrow_size * math.sin(angle + math.pi / 6))
    draw.polygon([(ex, ey), p1, p2], fill=color)

def draw_self_loop(draw, center, direction="left", color=(211, 47, 47, 255), width=3):
    """Vẽ mũi tên tự truyền (vòng lặp)."""
    cx, cy = center
    r_loop = 28
    if direction == "left":
        lx = cx - R - r_loop
        ly = cy
        draw.arc([lx - r_loop, ly - r_loop, lx + r_loop, ly + r_loop], start=45, end=315, fill=color, width=width)
        # Mũi tên hướng vào nút
        tip_x = cx - R - 2
        tip_y = cy - 8
        draw.polygon([(tip_x, tip_y), (tip_x - 12, tip_y - 8), (tip_x - 12, tip_y + 6)], fill=color)
    elif direction == "right":
        lx = cx + R + r_loop
        ly = cy
        draw.arc([lx - r_loop, ly - r_loop, lx + r_loop, ly + r_loop], start=225, end=135, fill=color, width=width)
        tip_x = cx + R + 2
        tip_y = cy - 8
        draw.polygon([(tip_x, tip_y), (tip_x + 12, tip_y - 8), (tip_x + 12, tip_y + 6)], fill=color)
    elif direction == "top":
        lx = cx
        ly = cy - R - r_loop
        draw.arc([lx - r_loop, ly - r_loop, lx + r_loop, ly + r_loop], start=135, end=45, fill=color, width=width)
        tip_x = cx + 8
        tip_y = cy - R - 2
        draw.polygon([(tip_x, tip_y), (tip_x + 8, tip_y - 12), (tip_x - 6, tip_y - 12)], fill=color)
    elif direction == "bottom":
        lx = cx
        ly = cy + R + r_loop
        draw.arc([lx - r_loop, ly - r_loop, lx + r_loop, ly + r_loop], start=315, end=225, fill=color, width=width)
        tip_x = cx + 8
        tip_y = cy + R + 2
        draw.polygon([(tip_x, tip_y), (tip_x + 8, tip_y + 12), (tip_x - 6, tip_y + 12)], fill=color)

# Vẽ các mũi tên liên kết
# 1. O -> A
draw_arrow(draw, nodes['O'], nodes['A'], color=(183, 28, 28, 255), width=3)
# 2. O -> B
draw_arrow(draw, nodes['O'], nodes['B'], color=(183, 28, 28, 255), width=3)
# 3. O -> AB
draw_arrow(draw, nodes['O'], nodes['AB'], color=(183, 28, 28, 255), width=3)
# 4. A -> AB
draw_arrow(draw, nodes['A'], nodes['AB'], color=(183, 28, 28, 255), width=3)
# 5. B -> AB
draw_arrow(draw, nodes['B'], nodes['AB'], color=(183, 28, 28, 255), width=3)

# Vẽ các vòng tự truyền
draw_self_loop(draw, nodes['O'], direction="left", color=(183, 28, 28, 255), width=3)
draw_self_loop(draw, nodes['AB'], direction="right", color=(183, 28, 28, 255), width=3)
draw_self_loop(draw, nodes['A'], direction="top", color=(183, 28, 28, 255), width=3)
draw_self_loop(draw, nodes['B'], direction="bottom", color=(183, 28, 28, 255), width=3)

# Vẽ các vòng tròn nút và chữ bên trong
for name, (cx, cy) in nodes.items():
    # Vòng tròn nền đỏ nhạt viền đỏ đậm
    draw.ellipse([cx - R, cy - R, cx + R, cy + R], fill=(255, 235, 238, 255), outline=(198, 40, 40, 255), width=3)
    # Text
    bbox = draw.textbbox((0, 0), name, font=font_node)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    draw.text((cx - tw / 2, cy - th / 2 - 2), name, fill=(183, 28, 28, 255), font=font_node)

# Tiêu đề dưới hình
caption = "Sơ đồ truyền máu ở người"
bbox_c = draw.textbbox((0, 0), caption, font=font_label)
cw = bbox_c[2] - bbox_c[0]
draw.text((W / 2 - cw / 2, 620), caption, fill=(60, 60, 60, 255), font=font_label)

img.save(IMG_PATH, "PNG")
print(f"Đã tạo ảnh sơ đồ truyền máu: {IMG_PATH} ({os.path.getsize(IMG_PATH)} bytes)")
