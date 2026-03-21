import os
import re
import subprocess

# 1. Bóc tách phiên bản từ file main.py
with open('main.py', 'r', encoding='utf-8') as f:
    content = f.read()
    match = re.search(r'CURRENT_VERSION\s*=\s*["\']([^"\']+)["\']', content)
    version = match.group(1) if match else "Unknown"

print(f"🚀 Bắt đầu Build PDFScan2Word phiên bản v{version} cho macOS...")

# 2. Dọn dẹp chiến trường cũ cho sạch sẽ
os.system("rm -rf build dist PDFScan2Word.spec")

# 3. Chạy lệnh PyInstaller
subprocess.run([
    "pyinstaller", 
    "--noconsole", 
    "--windowed", 
    "--collect-all", "customtkinter", 
    "--name", "PDFScan2Word", 
    "main.py"
])

# 4. Nén thành file zip với tên tự động
zip_name = f"PDFScan2Word-v{version}-macOS.zip"
print(f"\n📦 Đang nén file thành: {zip_name}...")

# Di chuyển vào thư mục dist và thực hiện lệnh zip của macOS
os.chdir("dist")
subprocess.run(["zip", "-r", zip_name, "PDFScan2Word.app"])

print(f"\n✅ HOÀN TẤT! File đã sẵn sàng tại: dist/{zip_name}")