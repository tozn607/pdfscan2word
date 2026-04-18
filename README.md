<div align="center">

# 📄 PDFScan2Word (v2.2.0)

**[🇻🇳 Tiếng Việt](#-tiếng-việt) • [🇺🇸 English](#-english)**

![GitHub release](https://img.shields.io/github/v/release/tozn607/pdfscan2word?color=success)
![Python](https://img.shields.io/badge/python-3.8%2B-blue)
![AI](https://img.shields.io/badge/AI-Google_Gemini_3.1-orange)
![License](https://img.shields.io/badge/license-MIT-green)

A powerful, AI-driven Desktop application to digitize scanned PDFs & Images into fully formatted Word (`.docx`) documents.  
*Ứng dụng Desktop tích hợp AI giúp chuyển đổi PDF và ảnh quét sang tài liệu Word (.docx) giữ nguyên định dạng.*

<img src="readme/1.png" alt="Main Interface" width="700"/>

</div>

---

## 🇻🇳 Tiếng Việt

### 🚀 Có gì mới trong bản v2.2.0?
Phiên bản này tập trung vào hiệu suất và sự ổn định:
- **Xử lý Song song:** Tăng siêu tốc độ quét (scan) bằng cách xử lý đồng thời nhiều trang PDF, thay vì xử lý tuần tự từng trang.
- **Tùy chỉnh Tốc độ:** Thêm lựa chọn "Tốc độ xử lý" với 3 mức (Eco, Balanced, Turbo) giúp tối ưu hóa quota API cho cả tài khoản Free và Paid.
- **Nâng cấp SDK Gemini:** Chuyển đổi sang thư viện `google-genai` mới nhất của Google để đảm bảo hiệu năng và hỗ trợ lâu dài.
- **Cải thiện độ ổn định:** Hệ thống tự động lắp ráp trang theo đúng thứ tự bản gốc kể cả khi xử lý song song, đồng thời tối ưu chiến lược backoff khi gặp lỗi giới hạn API (429).
- **Trải nghiệm Native:** Sử dụng font chữ hệ thống trên macOS và khắc phục triệt để các cảnh báo console liên quan đến accessibility.

### ✨ Tính năng nổi bật
- **Bảo toàn định dạng:** Nhận diện và giữ nguyên cấu trúc bảng biểu, danh sách, định dạng chữ đậm/nghiêng trực tiếp qua PyPandoc.
- **Khôi phục văn bản thông minh:** Sử dụng Gemini 3.1 Flash để suy luận và điền chính xác các phần văn bản bị mất do lỗi quét hoặc mép giấy bị che khuất.
- **Giải bài tập bằng AI:** Tự động nhận diện các câu hỏi/bài tập trong tài liệu scan và cung cấp lời giải chi tiết đính kèm cuối văn bản.
- **Xử lý hàng loạt:** Chuyển đổi toàn bộ thư mục PDF chỉ với một thao tác duy nhất.
- **Công cụ tạo PDF từ ảnh:** Tích hợp tiện ích gộp, nén và tối ưu hóa các định dạng `.heic`, `.jpg`, `.png` thành một file PDF duy nhất.
- **Xử lý trang đôi:** Tự động nhận diện và cắt đôi các bản quét sách định dạng A5 (hai trang trên một mặt scan).

### ⚙️ Hướng dẫn cài đặt (Khuyên dùng)
Người dùng thông thường có thể sử dụng ứng dụng ngay mà không cần cài đặt môi trường lập trình:

1. Truy cập trang [Releases](https://github.com/tozn607/pdfscan2word/releases).
2. Tải về phiên bản tương ứng với hệ điều hành:
   - **Windows:** Tải file `.exe` hoặc bản nén `.zip`.
   - **macOS:** Tải bản `.zip` dành cho `Apple Silicon (arm64)` hoặc `Intel (x86_64)`.
3. Giải nén và khởi chạy ứng dụng.
> **Lưu ý cho macOS:** Trong trường hợp hệ thống báo lỗi "app is damaged", vui lòng mở Terminal và chạy lệnh: `xattr -cr /đường-dẫn-đến/PDFScan2Word.app`

### 💻 Chạy từ mã nguồn
Dành cho lập trình viên muốn tùy chỉnh hoặc phát triển thêm:
```bash
git clone https://github.com/tozn607/pdfscan2word.git
cd pdfscan2word
pip install -r requirements.txt
python main.py
```

### 💡 Hướng dẫn sử dụng
1. Đăng ký API Key miễn phí tại [Google AI Studio](https://aistudio.google.com/).
2. Khởi động ứng dụng và chọn ngôn ngữ ưu tiên.
3. Nhập API Key vào ô cấu hình (Key sẽ được lưu tự động cho các lần sau).
4. Chọn chế độ **Single Mode** (xử lý 1 file) hoặc **Batch Mode** (xử lý cả thư mục).
5. Chọn file/thư mục đầu vào và nhấn **START PROCESSING**. Kết quả sẽ được tự động lưu dưới định dạng `.docx`.

---

## 🇺🇸 English

### 🚀 What's New in v2.2.0?
This update focuses on performance, speed, and reliability:
- **Parallel Execution:** Scanning speed is multiplied. The core engine now processes multiple PDF pages concurrently rather than one by one.
- **Processing Speed Controls:** New "Processing Speed" card with 3 selectable tiers (Eco, Balanced, Turbo) to optimize API quota usage for both Free and Paid accounts.
- **Modernized Gemini SDK:** Fully migrated to the new `google-genai` SDK for increased stability and future-proofing.
- **Perfect Page Assembly:** Guaranteed reassembly of pages in the exact original sequence even during high-speed parallel processing.
- **Native UI & Stability:** Native macOS system font support and improved rate-limit protection with smarter exponential backoff strategies.

### ✨ Key Features
- **Format Preservation:** Maintains bullet points, numbered lists, bold/italic text, and tables natively in Word using PyPandoc.
- **Smart Text Recovery:** Uses Google's Gemini 3.1 Flash to infer and fill in text cut-off at the page edges—perfect for thick university textbooks or legal docs.
- **AI Exercise Solver:** Automatically detects exercises in the scan and appends fully mathematical/worked-out solutions!
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
> **macOS Quarantine Note:** If macOS prevents the app from opening by saying "app is damaged", open Terminal and run: `xattr -cr /path/to/extracted/PDFScan2Word.app`

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

### ❤️ Credits
- UI Framework: [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
- AI Engine: [Google Generative AI (Gemini)](https://ai.google.dev/)
- Formatting Engine: [Pandoc](https://pandoc.org/)
- PDF Backend: [PyMuPDF](https://pymupdf.readthedocs.io/)

### 📜 Giấy phép (License)
Dự án được phát hành dưới giấy phép mã nguồn mở [MIT License](LICENSE).