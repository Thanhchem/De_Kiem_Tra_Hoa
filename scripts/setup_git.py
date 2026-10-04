import os
import urllib.request
import zipfile
import shutil

TARGET_DIR = r"C:\Users\Admin\AppData\Local\Programs\Git"
ZIP_PATH = r"C:\Users\Admin\AppData\Local\Temp\mingit.zip"
URL = "https://github.com/git-for-windows/git/releases/download/v2.56.0.windows.1/MinGit-2.56.0-64-bit.zip"

print(f"Downloading MinGit from {URL}...")
req = urllib.request.Request(URL, headers={"User-Agent": "Mozilla/5.0"})
with urllib.request.urlopen(req) as resp, open(ZIP_PATH, "wb") as out_file:
    shutil.copyfileobj(resp, out_file)
print("Download complete. Size:", os.path.getsize(ZIP_PATH))

print(f"Extracting to {TARGET_DIR}...")
os.makedirs(TARGET_DIR, exist_ok=True)
with zipfile.ZipFile(ZIP_PATH, 'r') as z:
    z.extractall(TARGET_DIR)

git_cmd_dir = os.path.join(TARGET_DIR, "cmd")
git_exe = os.path.join(git_cmd_dir, "git.exe")
print("Git exe exists:", os.path.exists(git_exe))

# Clean temp zip
try:
    os.remove(ZIP_PATH)
except Exception:
    pass

print("Git installed successfully at:", git_exe)
