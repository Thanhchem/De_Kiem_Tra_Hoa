import sys
sys.stdout.reconfigure(encoding='utf-8')
import re

with open('scratch/decompressed_press.xml', 'r', encoding='utf-8', errors='ignore') as f:
    text = f.read()

idx = 1161819
de24_xml = text[idx:]

# Extract all blocks and texts
# In decompressed_press.xml: <block ...><text ...>content</text></block>
blocks = re.findall(r'<block\s+l="(\d+)"\s+t="(\d+)"\s+w="(\d+)"\s+h="(\d+)"\s*>(.*?)</block>', de24_xml, re.DOTALL)
print(f'Total blocks found in De 24: {len(blocks)}')

clean_lines = []
for l, t, w, h, content in blocks:
    texts = re.findall(r'<text[^>]*>(.*?)</text>', content, re.DOTALL)
    combined = " ".join([t_str.strip() for t_str in texts if t_str.strip()])
    if combined:
        clean_lines.append(f"[{l},{t},{w},{h}] {combined}")

with open('scratch/de24_raw_text.txt', 'w', encoding='utf-8') as f:
    f.write("\n".join(clean_lines))

print("Saved raw text to scratch/de24_raw_text.txt")
