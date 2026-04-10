<div align="center">

# 📄 PDFScan2Word (v2.0.0)

**[🇺🇸 English](#-english) • [🇻🇳 Tiếng Việt](#-tiếng-việt)**

![GitHub release](https://img.shields.io/github/v/release/tozn607/pdfscan2word?color=success)
![Python](https://img.shields.io/badge/python-3.8%2B-blue)
![AI](https://img.shields.io/badge/AI-Google_Gemini_3.1-orange)
![License](https://img.shields.io/badge/license-MIT-green)

A powerful, AI-driven Desktop application to digitize scanned PDFs & Images into fully formatted Word (`.docx`) documents.  
*Một ứng dụng Desktop sức mạnh AI giúp số hóa PDF & Ảnh quét thành tài liệu Word (.docx) giữ nguyên định dạng.*

<img src="readme/1.png" alt="Main Interface" width="700"/>

*(Note: Screenshots are located in the `readme/` folder. Feel free to update them to feature the new v2.0 UI!)*

</div>

---

## 🇺🇸 English

### 🚀 What's New in v2.0.0?
This major update brings a complete overhaul to the application:
- **No System Dependencies Required:** We completely replaced `pdf2image` with `PyMuPDF`, meaning **Poppler is no longer required**! Pandoc is also automatically fetched if not detected.
- **Modernized UI:** A brand new, sleek, card-based interface with native Dark Mode support and responsive design.
- **Bilingual Support:** Full English & Vietnamese UI, featuring a first-time setup language selector.
- **AI Exercise Solver:** A brand new option that automatically detects exercises in the scan and appends fully mathematical/worked-out solutions!
- **Cross-Platform Pre-built Binaries:** Automatically built via GitHub Actions. Now available natively for Windows, macOS (Apple Silicon M1/M2/M3), and macOS (Intel).

### ✨ Key Features
- **Format Preservation:** Maintains bullet points, numbered lists, bold/italic text, and tables natively in Word using PyPandoc.
- **Smart Text Recovery:** Uses Google's Gemini 3.1 Flash to infer and fill in text cut-off at the page edges—perfect for thick university textbooks or legal docs.
- **Batch Processing:** Convert entire folders of PDFs with a single click.
- **Image to PDF Maker:** Includes a built-in utility to merge, compress, and enhance `.heic`, `.jpg`, `.png` into a single PDF.
- **Left-Right Page Splitting:** Scans 2 pages (like an A5 book spread) at once to speed up processing.

### ⚙️ Quick Installation (Recommended)
You don't need to know Python to use this app. Just download the pre-built application:

1. Go to the [Releases](https://github.com/tozn607/pdfscan2word/releases) page.
2. Download the version for your OS:
   - **Windows:** Download the `.exe` or `.zip` for Windows.
   - **macOS:** Download the `.zip` for `Apple Silicon (arm64)` or `Intel (x86_64)`.
3. Extract and run!
> **🍎 macOS Quarantine Note:** If macOS prevents the app from opening by saying "app is damaged", open Terminal and run: `xattr -cr /path/to/extracted/PDFScan2Word.app`

### 💻 Running from Source
If you are a developer and want to run it from the code:
```bash
git clone https://github.com/tozn607/pdfscan2word.git
cd pdfscan2word
pip install -r requirements.txt
python main.py
```

### 💡 How to Use
1. Get a free API Key from [Google AI Studio](https://aistudio.google.com/).
2. Open the app, select your preferred language (English).
3. Paste the API Key in the field.
4. Select "Single Mode" (1 file) or "Batch Mode" (folder).
5. Add your inputs and click **START PROCESSING**. It will auto-save the `.docx` file!

---

## 🇻🇳 Tiếng Việt

### 🚀 Có gì mới trong bản v2.0.0?
Phiên bản này mang đến đợt nâng cấp toàn diện nhất từ trước tới nay:
- **Không Cần Cài Thư Viện Hệ Thống:** Loại bỏ hoàn toàn sự phụ thuộc vào Poppler nhờ chuyển sang dùng `PyMuPDF`. Pandoc cũng được tự động tải về. Người dùng không cực nhọc cài đặt nữa!
- **Giao Diện Hiện Đại:** Lột xác hoàn toàn với giao diện thẻ (card-based), phông chữ to rõ trang nhã và Native Dark Mode.
- **Đa Ngôn Ngữ:** Hỗ trợ song ngữ Anh/Việt từ đầu tới cuối ứng dụng.
- **AI Giải Bài Tập:** Tính năng mới cho phép AI tự động nhận diện bài tập trong sách và làm bài giải chi tiết đính kèm cuối file Word!
- **Đóng Gói Đa Nền Tảng:** Hỗ trợ file chạy trực tiếp không cần cài Python trên Windows, macOS (Apple Silicon M1/M2/M3) và macOS (Intel).

### ✨ Tính năng Nổi bật
- **Bảo toàn Định dạng:** Nhận diện và giữ nguyên cấu trúc bảng biểu, danh sách list, in đậm/nghiêng trực tiếp.
- **Phục hồi Văn bản:** Có khả năng điền bù chữ bị tối hoặc gấp nếp sát mép gáy quyển giáo trình thật.
- **Xử lý Hàng loạt:** Quét hàng chục file PDF trong thư mục cùng lúc.
- **Gộp & Tối ưu Ảnh:** Tiện ích nhỏ giúp bạn ghép các file hình rời rạc (.heic, .jpg) thành 1 file PDF để chuẩn bị xử lý.
- **Xử lý Trang Đôi:** Nhận diện sách A5 được scan với 2 trang mỗi khung hình.

### ⚙️ Cài đặt Siêu Dễ (Khuyên dùng)
Bạn không cần biết code để xài, chỉ cần tải về nhấp đúp là chạy:

1. Tới trang [Releases](https://github.com/tozn607/pdfscan2word/releases).
2. Tải phiên bản phù hợp với hệ điều hành:
   - **Lưu ý Windows:** Tải file `.exe` hoạt động dạng Portable hoặc giải nén file `.zip`.
   - **Lưu ý macOS:** Tải bản `.zip` cho tương ứng dòng máy Apple Silicon (arm64) hoặc chip Intel (x86_64).
3. Giải nén và tận hưởng!
> **🍎 Khắc phục lỗi macOS:** Với lần đầu, nếu hệ thống báo "Ứng dụng bị hỏng (App is damaged)", hãy mở ứng dụng Terminal và gõ: `xattr -cr /đường-dẫn-tới/PDFScan2Word.app`

### 💻 Chạy trên Môi trường Python (Source)
```bash
git clone https://github.com/tozn607/pdfscan2word.git
cd pdfscan2word
pip install -r requirements.txt
python main.py
```

### 💡 Hướng dẫn Sử dụng Thao tác
1. Tạo một API Key miễn phí từ [Google AI Studio](https://aistudio.google.com/).
2. Mở ứng dụng lên, cài đặt ngôn ngữ ban đầu là Tiếng Việt.
3. Dán API key vào mục yêu cầu (Key sẽ tự lưu cho các lần sau).
4. Chọn chế độ **1 File PDF** hoặc **Thư Mục Hàng Loạt**.
5. Chọn đường dẫn file và bấm **BẮT ĐẦU XỬ LÝ**. Thưởng thức thành quả lưu ở đầu ra `.docx`.

---

### ❤️ Credits
- Cốt lõi giao diện: [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
- Hệ thần kinh số: [Google Generative AI (Gemini)](https://ai.google.dev/)
- Xử lý biên dịch Format: [Pandoc](https://pandoc.org/)
- Nền tảng PDF: [PyMuPDF](https://pymupdf.readthedocs.io/)

### 📜 Giấy phép (License)
Dự án open-source hoàn toàn 100% dưới giấy phép [MIT License](LICENSE).