import pymupdf
import re
import sys
sys.stdout.reconfigure(encoding='utf-8')

# Let's inspect text blocks on page 105, 106, 107
with open('scratch/decompressed_press.xml', 'r', encoding='utf-8', errors='ignore') as f:
    text = f.read()

# Pages are divided by <page ... >
pages = text.split('<page ')

# Find page 105 (index 104), page 106 (index 105), page 107 (index 106)
for p_idx in [104, 105, 106]:
    p_content = pages[p_idx]
    blocks = re.findall(r'<block\s+l="(\d+)"\s+t="(\d+)"\s+w="(\d+)"\s+h="(\d+)"\s*>(.*?)</block>', p_content, re.DOTALL)
    print(f"\n================ PAGE {p_idx+1} ================")
    for l, t, w, h, content in blocks:
        t_int = int(t)
        texts = re.findall(r'<text[^>]*>(.*?)</text>', content, re.DOTALL)
        combined = " ".join([t_str.strip() for t_str in texts if t_str.strip()])
        # Print blocks near the image areas
        if p_idx == 104 and 400 <= t_int <= 560:
            print(f"[{l},{t},{w},{h}] {combined}")
        elif p_idx == 105 and 130 <= t_int <= 420:
            print(f"[{l},{t},{w},{h}] {combined}")
        elif p_idx == 106 and (70 <= t_int <= 280 or 340 <= t_int <= 500):
            print(f"[{l},{t},{w},{h}] {combined}")
