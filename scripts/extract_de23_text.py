import sys
import re

with open('scratch/decompressed_press.xml', 'r', encoding='utf-8') as f:
    full_text = f.read()

pos_23 = 1108716
pos_24 = 1161819

chunk = full_text[pos_23:pos_24]

# Find all text elements
raw_texts = re.findall(r'<text[^>]*>(.*?)</text>', chunk)

clean_texts = []
for t in raw_texts:
    # remove inner xml tags if any
    clean = re.sub(r'<[^>]+>', '', t)
    clean = clean.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').strip()
    if clean:
        clean_texts.append(clean)

with open('scratch/de23_clean_text.txt', 'w', encoding='utf-8') as f:
    for t in clean_texts:
        f.write(t + '\n')

print(f"Total lines: {len(clean_texts)}")
for i, t in enumerate(clean_texts[:35]):
    print(f"{i+1}: {t}")
