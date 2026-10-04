import os
import json
import hashlib
from datetime import datetime

ROOT_DIR = r"c:\Antigravity_Thanh"
PDF_GOC_DIR = os.path.join(ROOT_DIR, "PDF_Goc")
HISTORY_JSON = os.path.join(ROOT_DIR, "lich_su_cong_viec.json")
HISTORY_MD = os.path.join(ROOT_DIR, "LICH_SU_CONG_VIEC.md")

def get_file_hash(filepath):
    hasher = hashlib.md5()
    with open(filepath, 'rb') as f:
        buf = f.read(65536)
        while len(buf) > 0:
            hasher.update(buf)
            buf = f.read(65536)
    return hasher.hexdigest()

def load_history():
    if os.path.exists(HISTORY_JSON):
        try:
            with open(HISTORY_JSON, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_history(history):
    with open(HISTORY_JSON, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)
    
    # Update Markdown log
    with open(HISTORY_MD, 'w', encoding='utf-8') as f:
        f.write("# NHẬT KÝ LỊCH SỬ XỬ LÝ ĐỀ THI (PDF -> TEX -> WORD)\n\n")
        f.write(f"*Cập nhật lần cuối: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n\n")
        f.write("| STT | Tên file PDF gốc | Ngày xử lý | File TeX (ex_test) | File PDF TeX | File Word (.docx) | Trạng thái |\n")
        f.write("|:---:|:---|:---:|:---:|:---:|:---:|:---:|\n")
        
        idx = 1
        for filename, info in history.items():
            status_icon = " Hoàn thành" if info.get("status") == "completed" else "⏳ Đang xử lý"
            f.write(f"| {idx} | `{filename}` | {info.get('processed_at', '')} | `{info.get('tex_file', '')}` | `{info.get('pdf_file', '')}` | `{info.get('docx_file', '')}` | {status_icon} |\n")
            idx += 1
        f.write("\n\n---\n")
        f.write("### Hướng dẫn sử dụng:\n")
        f.write("1. Copy bất kỳ file PDF mới nào cần xử lý vào thư mục `PDF_Goc/`.\n")
        f.write("2. Trợ lý sẽ tự động kiểm tra nhật ký, phân loại file mới chưa xử lý và tiến hành TeX hóa theo chuẩn `ex_test` rồi xuất Word.\n")

def check_pending_files():
    history = load_history()
    if not os.path.exists(PDF_GOC_DIR):
        os.makedirs(PDF_GOC_DIR, exist_ok=True)
        return []
    
    all_pdfs = [f for f in os.listdir(PDF_GOC_DIR) if f.lower().endswith('.pdf')]
    unprocessed = []
    for pdf in all_pdfs:
        if pdf not in history or history[pdf].get("status") != "completed":
            unprocessed.append(pdf)
    return unprocessed

if __name__ == "__main__":
    import sys
    unprocessed = check_pending_files()
    print("Files chưa xử lý:", unprocessed)
