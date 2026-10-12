import sys
sys.stdout.reconfigure(encoding='utf-8')
import fitz # PyMuPDF
from PIL import Image
import numpy as np

# Page 99:
# Let's inspect where text is around Câu 14, 15, 16, 17
p99 = Image.open('scratch/de23_pages/page_99.png')
w99, h99 = p99.size

# Let's check Câu 14:
# Text: "Phân tử phức chất cis-dichloro(ethylendiammine)platinium(II) có cấu tạo như hình sau:"
# Image of complex cis-dichloro(ethylenediamine)platinum(II)
# Text below image: "Có bao nhiêu loại phối tử có trong phân tử phức chất đó?"
# Then options A. 1. B. 2. C. 3. D. 4.

# Let's crop vertical region from y=2800 to 3800 and check where the drawing is
# Let's write a helper to save multiple slices for visualization
for y in range(2600, 4000, 200):
    p99.crop((400, y, 4600, y+200)).save(f'scratch/c14_y_{y}.png')

for y in range(5300, 6500, 200):
    p99.crop((400, y, 4600, y+200)).save(f'scratch/c16_y_{y}.png')

p100 = Image.open('scratch/de23_pages/page_100.png')
for y in range(300, 1600, 200):
    p100.crop((400, y, 4600, y+200)).save(f'scratch/c17_y_{y}.png')

print("Saved test slices")
