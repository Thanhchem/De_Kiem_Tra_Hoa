import re
import sys
sys.stdout.reconfigure(encoding='utf-8')

with open('scratch/de23_raw_chunk.txt', 'r', encoding='utf-8') as f:
    raw = f.read()

# Also let's check which <page> tags are present
pages = re.findall(r'<page l="(\d+)" t="(\d+)" w="(\d+)" h="(\d+)" >', raw)
print(f"Pages found in raw chunk: {len(pages)}")

# Split raw chunk by <page
page_chunks = raw.split('<page ')
for idx, p_chunk in enumerate(page_chunks[1:]):
    print(f"\n--- Page Chunk {idx+1} ---")
    matches = re.findall(r'<text l="(\d+)" t="(\d+)" w="(\d+)" h="(\d+)"[^>]*>(.*?)</text>', p_chunk)
    for l, t, w, h, txt in matches:
        clean_txt = re.sub(r'<[^>]+>', '', txt).strip()
        if any(k in clean_txt for k in ['Câu 14', 'Câu 15', 'Câu 16', 'Câu 17', 'Câu 18', 'PHẦN II', 'PHẦN III', 'ĐỀ SỐ 23', 'Trang']):
            print(f"  {clean_txt[:50]}: l={l}, t={t}, w={w}, h={h}")
