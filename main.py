import os
import urllib.request
import webbrowser
import sys
import time
import threading
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Pt, Cm
from datetime import datetime
import json
from PIL import Image, ImageEnhance
import ssl

# --- MACOS CRASH FIX ---
if sys.platform == "darwin":
    os.environ["PATH"] += os.pathsep + "/usr/local/bin" + os.pathsep + "/opt/homebrew/bin"

# --- MACOS SSL CERT FIX ---
try:
    _create_unverified_https_context = ssl._create_unverified_context
except AttributeError:
    pass
else:
    ssl._create_default_https_context = _create_unverified_https_context
# --------------------------------------------------------

import fitz 
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import pypandoc

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except ImportError:
    print("[!] CẢNH BÁO: Chưa cài đặt pillow-heif. Không thể đọc file HEIC.")

CURRENT_VERSION = "2.0.1"
GITHUB_API_URL = "https://api.github.com/repos/tozn607/pdfscan2word/releases/latest"
RELEASES_URL = "https://github.com/tozn607/pdfscan2word/releases"

def get_build_date():
    if getattr(sys, 'frozen', False):
        try:
            mtime = os.path.getmtime(sys.executable)
            return datetime.fromtimestamp(mtime)
        except: pass
    return datetime.now()

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".pdfscan2word")
API_KEY_FILE = os.path.join(CONFIG_DIR, "api_key.txt")
CONFIG_JSON_FILE = os.path.join(CONFIG_DIR, "config.json")


safety_config = {
    HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
}

PROMPT_VN = r"""
Bạn là một chuyên gia số hóa và phục hồi tài liệu chuyên nghiệp. Dưới đây là hình ảnh scan của một trang tài liệu/giáo trình. 
Nhiệm vụ của bạn là trích xuất và phục hồi, làm sạch văn bản theo các quy tắc NGHIÊM NGẶT sau đây:
1. CƠ CHẾ XUỐNG DÒNG: 
   - BẮT BUỘC chèn MỘT DÒNG TRẮNG (Enter 2 lần) giữa các đoạn văn, tiêu đề, mục danh sách để tránh bị dồn chữ.
   - Chỉ nối dòng nếu dòng dưới là câu đứt đoạn của dòng trên.
2. ĐẶC BIỆT XỬ LÝ MỤC LỤC (QUAN TRỌNG): 
   - Vẫn giữ nguyên một dòng trắng (Enter 2 lần) giữa các mục.
   - BẮT BUỘC khôi phục sự Thụt lề (Indentation) bằng cách chèn `&emsp;` vào đầu dòng (Mục lớn không thụt, Cấp 1 thụt 1 `&emsp;`, Cấp 2 thụt 2 `&emsp;`...).
   - TUYỆT ĐỐI KHÔNG GÕ dải dấu chấm (`......`) nối giữa tên mục và số trang. 
   - BẮT BUỘC phải thay thế toàn bộ dải dấu chấm đó bằng MỘT KÝ TỰ TAB DUY NHẤT với mã HTML là `&#9;`. 
   - Cú pháp chuẩn: `&emsp;1.1. Các phong cách học tập&#9;4`
3. NGĂN TỰ ĐỘNG TẠO BULLET POINT: 
   - NẾU bản gốc có dấu gạch ngang (-) ở đầu dòng, BẮT BUỘC phải dùng dấu gạch chéo ngược để thoát (escape): Viết là `\- ` thay vì `- `.
   - Các mục đánh số (1., 2.) hoặc chữ cái (a., b.) thì giữ nguyên, KHÔNG chèn thêm gạch ngang.
4. ĐỊNH DẠNG: **In đậm** và *In nghiêng* đúng bản gốc.
5. ĐIỀN CHỮ THIẾU: Dựa vào ngữ cảnh chung để điền bù chữ khuất mép giấy. Xóa bỏ hoàn toàn ký tự rác.
6. LOẠI BỎ SỐ TRANG: TUYỆT ĐỐI KHÔNG ghi lại số trang ở lề trên/dưới cùng của trang.
7. XỬ LÝ CHÚ THÍCH (FOOTNOTE):
   - Đặt mốc `[^1]`, `[^2]`... sát ngay sau từ/câu được chú thích.
   - Ghi nội dung chú thích ở tận cùng của văn bản theo cú pháp: `[^1]: Nội dung chú thích...`
8. THỤT ĐẦU DÒNG ĐOẠN VĂN: Nếu đoạn văn trong bản gốc có lùi vào ở dòng đầu tiên, chèn `&emsp;&emsp;` vào đầu đoạn.
9. BẢNG BIỂU (TABLES) - KHÔNG DÙNG HTML:
   - BẮT BUỘC sử dụng cú pháp bảng Markdown tiêu chuẩn (dùng các dấu gạch đứng `|`). TUYỆT ĐỐI KHÔNG DÙNG mã HTML.
   - XỬ LÝ Ô GỘP (MERGED CELLS): Điền nội dung vào hàng đầu tiên của nhóm ô gộp. Các hàng bên dưới thuộc cùng nhóm thì ĐỂ TRỐNG (ví dụ: `| | Toán | 105 |`).
Chỉ trả về văn bản bằng Markdown, không giải thích gì thêm.
"""

PROMPT_EN = r"""
You are a professional document digitization and restoration expert. Here is a scanned image of a document/textbook page. 
Your task is to extract, restore, and clean the text following these STRICT rules:
1. LINE BREAK MECHANISM: 
   - YOU MUST insert ONE BLANK LINE (Enter twice) between paragraphs, headings, and lists to prevent clumping.
   - Only join lines if the bottom line is a continuation of a broken sentence.
2. TABLE OF CONTENTS (CRUCIAL): 
   - Keep one blank line between items.
   - YOU MUST restore Indentation by inserting `&emsp;` at the beginning of the line.
   - ABSOLUTELY DO NOT type a dot strip (`......`) connecting the section name and page number.
   - YOU MUST replace the entire dot strip with a SINGLE TAB character using the HTML code `&#9;`. 
   - Standard syntax: `&emsp;1.1. Learning styles&#9;4`
3. PREVENT AUTO BULLET POINTS: 
   - If the original has a dash (-) at the start of a line, YOU MUST escape it: Write `\- ` instead of `- `.
   - Numbered (1., 2.) or lettered (a., b.) items remain unchanged, DO NOT insert dashes.
4. FORMATTING: Preserve **Bold** and *Italics* exactly as the original.
5. FILL MISSING TEXT: Use context to fill in text cut off at the edges. Remove all garbage characters.
6. REMOVE PAGE NUMBERS: ABSOLUTELY DO NOT transcribe page numbers at the top/bottom margins.
7. FOOTNOTES:
   - Place markers `[^1]`, `[^2]`... immediately after the annotated word/sentence.
   - Write the footnote content at the very end of the text using the syntax: `[^1]: Footnote content...`
8. PARAGRAPH INDENTATION: If a paragraph is indented on the first line in the original, insert `&emsp;&emsp;` at the beginning.
9. TABLES - NO HTML:
   - YOU MUST use standard Markdown table syntax (using vertical pipes `|`). ABSOLUTELY NO HTML code.
   - MERGED CELLS: Fill the content in the first row of the merged group. Leave the rows below in the same group BLANK (e.g., `| | Math | 105 |`).
Only return the text in Markdown, provide no additional explanations.
"""

STRINGS = {
    "VN": {
        "title": "Chuyển Ảnh sang Word",
        "toolbar": "Thanh công cụ:",
        "lang_switch": "🌐 Ngôn ngữ",
        "merge_pdf": "🖼️ GỘP ẢNH THÀNH PDF",
        "mode_single": "Chế độ Đơn (1 File PDF)",
        "mode_batch": "Chế độ Hàng loạt (Thư mục)",
        "api_key": "Google API Key:",
        "load_api": "Tải từ file .txt",
        "api_placeholder": "Nhập API Key...",
        "select_pdf": "Chọn File PDF",
        "input_pdf_ph": "Đường dẫn đến 1 file PDF...",
        "input_dir_btn": "Thư mục Input",
        "input_dir_ph": "Thư mục chứa nhiều file PDF...",
        "output_dir": "Thư mục Output",
        "output_ph": "Trống = Tự động lưu cùng nơi với Input",
        "solve_opt": "🤖 AI Giải bài tập",
        "cover_opt": "🖼️ Lưu riêng trang bìa",
        "merge_opt": "📖 Gộp 2 trang làm 1 (Sách A5)",
        "start": "▶ BẮT ĐẦU XỬ LÝ",
        "stop": "⏹ DỪNG LẠI",
        "timer": "Thời gian xử lý: {0:02d}:{1:02d}",
        "timer_init": "Thời gian: 00:00",
        "log_ready": "[*] Ứng dụng đã sẵn sàng. Hãy chọn file/thư mục để bắt đầu.",
        "log_stop_cmd": "\n[!] NHẬN LỆNH DỪNG... Đang hủy bỏ (Vui lòng đợi 1-2 giây).",
        "err_api": "[!] LỖI: Vui lòng nhập API Key.",
        "err_input": "[!] LỖI: Đường dẫn Input không tồn tại.",
        "err_output": "[!] LỖI: Thư mục Output không tồn tại.",
        "msg_auto_out": "[*] Không chọn Output, tự động lưu cùng thư mục Input: {0}",
        "log_no_pdf": "\n[!] Không tìm thấy file PDF nào để xử lý!",
        "log_start_batch": "\n[>>>] BẮT ĐẦU: Sẽ xử lý {0} file PDF.",
        "log_split": "\n[FILE {0}/{1}] Đang tách: {2}...",
        "log_pandoc": "  [*] Đang tải trình biên dịch Word (chỉ chạy 1 lần)...",
        "log_err_read": "  [X] LỖI ĐỌC PDF: {0}",
        "log_save_cover": "  [+] Đã lưu ảnh bìa: {0}",
        "log_err_cover": "  [!] Lỗi khi lưu ảnh bìa: {0}",
        "log_merge_pages": "  [*] Đang gộp cặp trang (trái-phải) để tăng tốc độ quét...",
        "log_read_block": "    [>] Đang đọc khối ảnh {0}/{1}...",
        "log_reject": "      [!] LỖI BẢN QUYỀN: Google từ chối đọc trang {0}.",
        "log_err_attempt": "      [!] Lỗi (Thử {0}/{1}): {2}",
        "log_skip": "      [X] Bỏ qua trang {0}.",
        "log_format_err": "  [!] Lỗi khi định dạng file Word: {0}",
        "log_save_draft": "  [BẢN NHÁP] Đã lưu kết quả tại: {0}",
        "log_save_ok": "  [***] Đã lưu kết quả tại: {0}",
        "log_rescue": "  [+] Đã cứu dữ liệu dưới dạng Markdown tại: {0}",
        "log_cancel": "\n[⏹] TIẾN TRÌNH ĐÃ BỊ HỦY BỞI NGƯỜI DÙNG.",
        "log_done": "\n[✓] ĐÃ HOÀN TẤT!",
        "dev_footer": "Developed by @tozn607 | Version v{0} Build {1} | © {2}",
        "merge_title": "Tiện ích: Gộp ảnh thành PDF",
        "add_img": "Thêm ảnh",
        "up": "▲ Lên",
        "down": "▼ Xuống",
        "clear": "Xóa hết",
        "compress": "🗜️ Nén giảm dung lượng",
        "enhance": "✨ Làm sáng & Rõ chữ",
        "export_pdf": "XUẤT RA PDF",
        "exporting": "Đang xử lý ảnh, vui lòng đợi...",
        "merge_success": "--- THÀNH CÔNG ---\nĐã lưu PDF: {0}",
        "merge_err": "[!] LỖI: {0}",
        "update_title": "Thông báo Cập nhật",
        "update_msg": "Có phiên bản mới: v{0}\nPhiên bản hiện tại: v{1}\n\nBạn có muốn tải bản cập nhật về không?",
        "btn_yes": "Tải ngay",
        "btn_no": "Bỏ qua",
        "prompt_solve_ext": "\n\n6. YÊU CẦU ĐẶC BIỆT: Tài liệu này chứa các bài tập/câu hỏi. Bạn BẮT BUỘC phải đọc và TRẢ LỜI/GIẢI CHI TIẾT các bài tập đó. Hãy tạo một phần 'ĐÁP ÁN' riêng biệt và rõ ràng ở cuối tài liệu và viết đáp án ở đó."
    },
    "EN": {
        "title": "Image to Word Converter",
        "toolbar": "Toolbar:",
        "lang_switch": "🌐 Language",
        "merge_pdf": "🖼️ MERGE IMAGES TO PDF",
        "mode_single": "Single Mode (1 PDF file)",
        "mode_batch": "Batch Mode (Folder)",
        "api_key": "Google API Key:",
        "load_api": "Load from .txt",
        "api_placeholder": "Enter API Key...",
        "select_pdf": "Select PDF",
        "input_pdf_ph": "Path to a single PDF file...",
        "input_dir_btn": "Input Folder",
        "input_dir_ph": "Folder containing multiple PDFs...",
        "output_dir": "Output Folder",
        "output_ph": "Empty = Auto save next to Input",
        "solve_opt": "🤖 AI Solves Exercises",
        "cover_opt": "🖼️ Save cover separately",
        "merge_opt": "📖 Merge 2 pages (A5 books)",
        "start": "▶ START PROCESSING",
        "stop": "⏹ STOP",
        "timer": "Processing time: {0:02d}:{1:02d}",
        "timer_init": "Time: 00:00",
        "log_ready": "[*] Application ready. Please select a file/folder to start.",
        "log_stop_cmd": "\n[!] STOP COMMAND RECEIVED... Canceling tasks (Please wait 1-2s).",
        "err_api": "[!] ERROR: Please enter API Key.",
        "err_input": "[!] ERROR: Input path does not exist.",
        "err_output": "[!] ERROR: Output folder does not exist.",
        "msg_auto_out": "[*] Output not selected, auto saving to Input folder: {0}",
        "log_no_pdf": "\n[!] No PDF files found to process!",
        "log_start_batch": "\n[>>>] START: Will process {0} PDF files.",
        "log_split": "\n[FILE {0}/{1}] Extracting: {2}...",
        "log_pandoc": "  [*] Downloading Word compiler (runs only once)...",
        "log_err_read": "  [X] PDF READ ERROR: {0}",
        "log_save_cover": "  [+] Saved cover image: {0}",
        "log_err_cover": "  [!] Error saving cover: {0}",
        "log_merge_pages": "  [*] Merging page pairs (left-right) to speed up scanning...",
        "log_read_block": "    [>] Reading image block {0}/{1}...",
        "log_reject": "      [!] COPYRIGHT ERROR: Google refused to read page {0}.",
        "log_err_attempt": "      [!] Error (Attempt {0}/{1}): {2}",
        "log_skip": "      [X] Skipping page {0}.",
        "log_format_err": "  [!] Error formatting Word file: {0}",
        "log_save_draft": "  [DRAFT] Result saved at: {0}",
        "log_save_ok": "  [***] Result saved at: {0}",
        "log_rescue": "  [+] Rescued data as Markdown at: {0}",
        "log_cancel": "\n[⏹] PROCESS CANCELLED BY USER.",
        "log_done": "\n[✓] COMPLETED!",
        "dev_footer": "Developed by @tozn607 | Version v{0} Build {1} | © {2}",
        "merge_title": "Utility: Merge Images to PDF",
        "add_img": "Add Images",
        "up": "▲ Up",
        "down": "▼ Down",
        "clear": "Clear All",
        "compress": "🗜️ Compress size",
        "enhance": "✨ Enhance & Brighten",
        "export_pdf": "EXPORT TO PDF",
        "exporting": "Processing images, please wait...",
        "merge_success": "--- SUCCESS ---\nSaved PDF: {0}",
        "merge_err": "[!] ERROR: {0}",
        "update_title": "Update Notification",
        "update_msg": "New version available: v{0}\nCurrent version: v{1}\n\nDo you want to download the update?",
        "btn_yes": "Download Now",
        "btn_no": "Skip",
        "prompt_solve_ext": "\n\n6. SPECIAL REQUIREMENT: This document contains exercises/questions. You MUST read and ANSWER/SOLVE them in detail. Create a clear, separate 'ANSWERS' section at the end of the document and write your answers there."
    }
}


from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QCheckBox, QComboBox, QTextEdit, QFileDialog, 
                             QMessageBox, QDialog, QButtonGroup, QListWidget, QFrame,
                             QSizePolicy, QScrollArea, QAbstractItemView)
from PyQt6.QtCore import Qt, pyqtSignal, QThread, QTimer, QSize
from PyQt6.QtGui import QFont, QIcon, QColor, QPalette, QCursor

QSS = """
    QWidget {
        font-family: 'Segoe UI', Arial, sans-serif;
        font-size: 14px;
    }
    QFrame#Card {
        background-color: rgba(128, 128, 128, 0.08);
        border-radius: 12px;
        border: 1px solid rgba(128, 128, 128, 0.2);
    }
    QPushButton {
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: bold;
        background-color: rgba(128, 128, 128, 0.15);
        border: 1px solid rgba(128, 128, 128, 0.2);
    }
    QPushButton:hover {
        background-color: rgba(128, 128, 128, 0.25);
    }
    QPushButton:checked {
        background-color: #1f538d;
        color: white;
        border: none;
    }
    QPushButton#Primary {
        background-color: #1f538d;
        color: white;
        border: none;
    }
    QPushButton#Primary:hover { background-color: #14375e; }
    QPushButton#Primary:disabled { background-color: #566c85; color: #d0d0d0; }
    
    QPushButton#Danger {
        background-color: #a83232;
        color: white;
        border: none;
    }
    QPushButton#Danger:hover { background-color: #7a2121; }
    QPushButton#Danger:disabled { background-color: #633f3f; color: #d0d0d0; }
    
    QPushButton#Success {
        background-color: #2b7a4b;
        color: white;
        border: none;
    }
    QPushButton#Success:hover { background-color: #1e5c37; }
    
    QLineEdit, QComboBox, QTextEdit, QListWidget {
        background-color: rgba(128, 128, 128, 0.05);
        border: 1px solid rgba(128, 128, 128, 0.3);
        border-radius: 6px;
        padding: 6px;
    }
    QLineEdit:focus, QComboBox:focus, QTextEdit:focus, QListWidget:focus {
        border: 1px solid #1f538d;
    }
"""


class LanguageSelectorPopup(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Language / Chọn ngôn ngữ")
        self.setFixedSize(450, 300)
        self.selected_lang = "EN"
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        
        lbl = QLabel("Welcome! Please choose your preferred language.\nChào mừng! Vui lòng chọn ngôn ngữ của bạn.")
        font = QFont("Segoe UI", 13, QFont.Weight.Bold)
        lbl.setFont(font)
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl)
        
        layout.addSpacing(20)
        
        self.btn_en = QPushButton("🇺🇸 English")
        self.btn_en.setCheckable(True)
        self.btn_en.setChecked(True)
        self.btn_en.setFont(QFont("Segoe UI", 12))
        
        self.btn_vn = QPushButton("🇻🇳 Tiếng Việt")
        self.btn_vn.setCheckable(True)
        self.btn_vn.setFont(QFont("Segoe UI", 12))
        
        self.group = QButtonGroup(self)
        self.group.addButton(self.btn_en)
        self.group.addButton(self.btn_vn)
        
        layout.addWidget(self.btn_en)
        layout.addWidget(self.btn_vn)
        
        layout.addStretch()
        
        self.btn_save = QPushButton("Continue / Tiếp tục")
        self.btn_save.setObjectName("Success")
        self.btn_save.setMinimumHeight(45)
        self.btn_save.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        self.btn_save.clicked.connect(self.on_save)
        layout.addWidget(self.btn_save)
        
    def on_save(self):
        if self.btn_vn.isChecked():
            self.selected_lang = "VN"
        else:
            self.selected_lang = "EN"
        self.accept()

class WorkerThread(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, app_instance, input_path, output_dir, mode):
        super().__init__()
        self.app = app_instance
        self.input_path = input_path
        self.output_dir = output_dir
        self.mode = mode
        
    def run(self):
        import time 
        try:
            self.process_documents(self.input_path, self.output_dir, self.mode)
        except Exception as e:
            self.log_signal.emit(f"[!] ERROR: {str(e)}")
        finally:
            self.finished_signal.emit()

    def write_log(self, msg):
        self.log_signal.emit(msg)

    def process_documents(self, input_path, output_dir, mode):
        if mode == self.app.t("mode_single"):
            pdf_files = [input_path]
        else:
            pdf_files = [os.path.join(input_path, f) for f in os.listdir(input_path) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            self.write_log(self.app.t("log_no_pdf"))
            return

        total_files = len(pdf_files)
        self.write_log(self.app.t("log_start_batch", total_files))
        model = genai.GenerativeModel('gemini-3.1-flash-lite-preview')

        active_prompt = PROMPT_EN if self.app.current_lang == "EN" else PROMPT_VN
        if self.app.solve_var:
            active_prompt += self.app.t("prompt_solve_ext")

        for file_idx, pdf_path in enumerate(pdf_files):
            if self.app.stop_event.is_set(): break

            pdf_filename = os.path.basename(pdf_path)
            base_name = os.path.splitext(pdf_filename)[0]
            suffix = "_hoanthien" if self.app.current_lang == "VN" else "_completed"
            output_docx_path = os.path.join(output_dir, f"{base_name}{suffix}.docx")

            self.write_log(self.app.t("log_split", file_idx + 1, total_files, pdf_filename))
            
            images = []
            try: 
                pandoc_exe = os.path.join(CONFIG_DIR, "pandoc" + (".exe" if sys.platform == "win32" else ""))
                if os.path.exists(pandoc_exe):
                    os.environ['PYPANDOC_PANDOC'] = pandoc_exe

                try:
                    pypandoc.get_pandoc_version()
                except OSError:
                    self.write_log(self.app.t("log_pandoc"))
                    pypandoc.download_pandoc(targetfolder=CONFIG_DIR, download_folder=CONFIG_DIR)
                    os.environ['PYPANDOC_PANDOC'] = pandoc_exe

                doc = fitz.open(pdf_path)
                for page in doc:
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    images.append(img)
                    
            except Exception as e:
                self.write_log(self.app.t("log_err_read", e))
                continue 

            if not images:
                continue

            start_idx = 0
            if self.app.cover_var:
                cover_path = os.path.join(output_dir, f"{base_name}_cover.jpg")
                try:
                    images[0].save(cover_path, "JPEG")
                    self.write_log(self.app.t("log_save_cover", f"{base_name}_cover.jpg"))
                except Exception as e:
                    self.write_log(self.app.t("log_err_cover", e))
                start_idx = 1 

            processed_images = []
            if self.app.merge_pages_var:
                self.write_log(self.app.t("log_merge_pages"))
                for i in range(start_idx, len(images), 2):
                    img_left = images[i]
                    if i + 1 < len(images):
                        img_right = images[i+1]
                        total_width = img_left.width + img_right.width
                        max_height = max(img_left.height, img_right.height)
                        new_img = Image.new('RGB', (total_width, max_height))
                        new_img.paste(img_left, (0, 0))
                        new_img.paste(img_right, (img_left.width, 0))
                        processed_images.append(new_img)
                    else:
                        processed_images.append(img_left)
            else:
                processed_images = images[start_idx:]

            full_markdown_content = ""
            max_retries = 5 

            for i, image in enumerate(processed_images):
                if self.app.stop_event.is_set():
                    break

                self.write_log(self.app.t("log_read_block", i+1, len(processed_images)))
                
                for attempt in range(max_retries):
                    if self.app.stop_event.is_set(): break
                    try:
                        response = model.generate_content([active_prompt, image], safety_settings=safety_config)
                        try:
                            text_result = response.text
                        except ValueError:
                            self.write_log(self.app.t("log_reject", i+1))
                            text_result = f"> **[COPYRIGHT BLOCK: PAGE {i+1}]**"
                            full_markdown_content += f"\n\n\n\n{text_result}\n\n"
                            break
                        
                        text_result = text_result.replace("[^", f"[^p{i}_")
                        full_markdown_content += f"\n\n\n\n{text_result}\n\n"
                        
                        if i < len(images) - 1 and not self.app.stop_event.is_set():
                            time.sleep(3) 
                        break 
                        
                    except Exception as e:
                        self.write_log(self.app.t("log_err_attempt", attempt+1, max_retries, e))
                        if attempt < max_retries - 1:
                            time.sleep(60) 
                        else:
                            self.write_log(self.app.t("log_skip", i+1))

            if full_markdown_content.strip():
                try:
                    pypandoc.convert_text(full_markdown_content, 'docx', format='md', outputfile=output_docx_path)
                    try:
                        doc_docx = docx.Document(output_docx_path)
                        for paragraph in doc_docx.paragraphs:
                            if '\t' in paragraph.text:
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                paragraph.paragraph_format.space_after = Pt(0)
                                paragraph.paragraph_format.space_before = Pt(0)
                                paragraph.paragraph_format.tab_stops.clear_all()
                                paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
                            else:
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

                        for table in doc_docx.tables:
                            tblPr = table._tbl.tblPr
                            tblBorders = OxmlElement('w:tblBorders')
                            for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                                border = OxmlElement(f'w:{border_name}')
                                border.set(qn('w:val'), 'single') 
                                border.set(qn('w:sz'), '4')       
                                border.set(qn('w:space'), '0')    
                                border.set(qn('w:color'), '000000') 
                                tblBorders.append(border)
                            tblPr.append(tblBorders)
                            
                        doc_docx.save(output_docx_path)
                    except Exception as e_docx:
                        self.write_log(self.app.t("log_format_err", e_docx))

                    if self.app.stop_event.is_set():
                        self.write_log(self.app.t("log_save_draft", output_docx_path))
                    else:
                        self.write_log(self.app.t("log_save_ok", output_docx_path))
                except Exception as e:
                    md_path = output_docx_path.replace('.docx', '.md')
                    with open(md_path, 'w', encoding='utf-8') as f: 
                        f.write(full_markdown_content)
                    self.write_log(self.app.t("log_rescue", md_path))
            
            if self.app.stop_event.is_set():
                break

        if self.app.stop_event.is_set():
            self.write_log(self.app.t("log_cancel"))
        else:
            self.write_log(self.app.t("log_done"))

class MergeWindow(QDialog):
    def __init__(self, app_instance):
        super().__init__(app_instance)
        self.app = app_instance
        self.setWindowTitle(self.app.t("merge_title"))
        self.setMinimumSize(660, 520)
        self.setStyleSheet(QSS)
        
        self.selected_images_list = []
        layout = QVBoxLayout(self)
        
        controls = QHBoxLayout()
        self.btn_add = QPushButton(self.app.t("add_img"))
        self.btn_add.setObjectName("Success")
        self.btn_add.clicked.connect(self.add_images)
        
        self.btn_up = QPushButton(self.app.t("up"))
        self.btn_up.clicked.connect(self.move_up)
        
        self.btn_down = QPushButton(self.app.t("down"))
        self.btn_down.clicked.connect(self.move_down)
        
        self.btn_clear = QPushButton(self.app.t("clear"))
        self.btn_clear.setObjectName("Danger")
        self.btn_clear.clicked.connect(self.clear_images)
        
        controls.addWidget(self.btn_add)
        controls.addWidget(self.btn_up)
        controls.addWidget(self.btn_down)
        controls.addStretch()
        controls.addWidget(self.btn_clear)
        layout.addLayout(controls)
        
        opt_layout = QHBoxLayout()
        self.chk_compress = QCheckBox(self.app.t("compress"))
        self.chk_compress.setChecked(True)
        self.chk_enhance = QCheckBox(self.app.t("enhance"))
        opt_layout.addWidget(self.chk_compress)
        opt_layout.addWidget(self.chk_enhance)
        opt_layout.addStretch()
        layout.addLayout(opt_layout)
        
        self.listbox = QListWidget()
        self.listbox.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        layout.addWidget(self.listbox)
        
        self.btn_export = QPushButton(self.app.t("export_pdf"))
        self.btn_export.setObjectName("Success")
        self.btn_export.setMinimumHeight(50)
        self.btn_export.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        self.btn_export.clicked.connect(self.export_to_pdf)
        layout.addWidget(self.btn_export)
        
    def add_images(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select Images", "", "Image Files (*.png *.jpg *.jpeg *.bmp *.heic *.heif)")
        if paths:
            self.selected_images_list.extend(paths)
            self.update_listbox()
            
    def move_up(self):
        row = self.listbox.currentRow()
        if row > 0:
            self.selected_images_list[row-1], self.selected_images_list[row] = self.selected_images_list[row], self.selected_images_list[row-1]
            self.update_listbox()
            self.listbox.setCurrentRow(row-1)

    def move_down(self):
        row = self.listbox.currentRow()
        if 0 <= row < len(self.selected_images_list) - 1:
            self.selected_images_list[row+1], self.selected_images_list[row] = self.selected_images_list[row], self.selected_images_list[row+1]
            self.update_listbox()
            self.listbox.setCurrentRow(row+1)

    def clear_images(self):
        self.selected_images_list.clear()
        self.update_listbox()
        
    def update_listbox(self):
        self.listbox.clear()
        for p in self.selected_images_list:
            self.listbox.addItem(os.path.basename(p))

    def export_to_pdf(self):
        if not self.selected_images_list: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF (*.pdf)")
        if save_path:
            self.listbox.clear()
            self.listbox.addItem(self.app.t("exporting"))
            QApplication.processEvents()
            
            try:
                processed = []
                for p in self.selected_images_list:
                    img = Image.open(p).convert('RGB')
                    if self.chk_enhance.isChecked():
                        img = ImageEnhance.Contrast(img).enhance(1.5)
                        img = ImageEnhance.Sharpness(img).enhance(2.0)
                        img = ImageEnhance.Brightness(img).enhance(1.1)
                    if self.chk_compress.isChecked():
                        max_w = 1200
                        if img.width > max_w:
                            ratio = max_w / img.width
                            img = img.resize((max_w, int(img.height * ratio)), Image.Resampling.LANCZOS)
                    processed.append(img)
                processed[0].save(save_path, save_all=True, append_images=processed[1:])
                self.listbox.clear()
                self.listbox.addItem(self.app.t("merge_success", save_path))
            except Exception as e:
                self.listbox.clear()
                self.listbox.addItem(self.app.t("merge_err", str(e)))

class PDFOCRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_lang = "VN" # Default
        self.load_config()
        self.solve_var = False
        self.cover_var = False
        self.merge_pages_var = False
        self.stop_event = threading.Event()
        self.start_time = None
        self.timer_running = False

        self.setWindowTitle(self.t("title") + " v" + CURRENT_VERSION)
        self.setMinimumSize(840, 880)
        self.setStyleSheet(QSS)
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_timer_ui)
        
        self.worker = None

        self.build_ui()
        self.update_ui_texts()
        
        if not os.path.exists(CONFIG_JSON_FILE):
            QTimer.singleShot(200, self.show_language_popup)

        # Update check threaded
        threading.Thread(target=self.check_for_updates, daemon=True).start()

    def t(self, key, *args):
        text = STRINGS.get(self.current_lang, STRINGS["VN"]).get(key, key)
        if args: return text.format(*args)
        return text

    def load_config(self):
        if os.path.exists(CONFIG_JSON_FILE):
            try:
                with open(CONFIG_JSON_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.current_lang = data.get("lang", "VN")
            except: pass

    def save_config(self):
        if not os.path.exists(CONFIG_DIR): os.makedirs(CONFIG_DIR)
        try:
            with open(CONFIG_JSON_FILE, "w", encoding="utf-8") as f:
                json.dump({"lang": self.current_lang}, f)
        except: pass

    def show_language_popup(self):
        dlg = LanguageSelectorPopup(self)
        if dlg.exec():
            self.current_lang = dlg.selected_lang
            self.save_config()
            idx = 0 if self.current_lang == "EN" else 1
            self.lang_menu.setCurrentIndex(idx)
            self.update_ui_texts()

    def build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Menu frame
        menu_frame = QFrame()
        menu_frame.setStyleSheet("background-color: rgba(128, 128, 128, 0.2);")
        menu_layout = QHBoxLayout(menu_frame)
        self.lbl_toolbar = QLabel("")
        self.lbl_toolbar.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        self.btn_merge = QPushButton("")
        self.btn_merge.setObjectName("Success")
        self.btn_merge.setMinimumHeight(40)
        self.btn_merge.clicked.connect(self.open_merge_popup)
        menu_layout.addWidget(self.lbl_toolbar)
        menu_layout.addWidget(self.btn_merge)
        menu_layout.addStretch()
        
        self.lbl_lang = QLabel("🌐")
        self.lang_menu = QComboBox()
        self.lang_menu.addItems(["EN (English)", "VN (Tiếng Việt)"])
        idx = 0 if self.current_lang == "EN" else 1
        self.lang_menu.setCurrentIndex(idx)
        self.lang_menu.currentIndexChanged.connect(self.change_language)
        menu_layout.addWidget(self.lbl_lang)
        menu_layout.addWidget(self.lang_menu)
        main_layout.addWidget(menu_frame)

        # Content 
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(25, 20, 25, 20)
        main_layout.addLayout(content_layout)

        # Segmented Button (Mode)
        mode_layout = QHBoxLayout()
        self.btn_mode_single = QPushButton("")
        self.btn_mode_single.setCheckable(True)
        self.btn_mode_single.setChecked(True)
        self.btn_mode_single.setMinimumHeight(45)
        self.btn_mode_batch = QPushButton("")
        self.btn_mode_batch.setCheckable(True)
        self.btn_mode_batch.setMinimumHeight(45)
        
        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.btn_mode_single)
        self.mode_group.addButton(self.btn_mode_batch)
        self.mode_group.buttonClicked.connect(self.change_mode)
        
        mode_layout.addWidget(self.btn_mode_single)
        mode_layout.addWidget(self.btn_mode_batch)
        content_layout.addLayout(mode_layout)

        # API KEY CARD
        api_card = QFrame(); api_card.setObjectName("Card")
        api_layout = QHBoxLayout(api_card)
        self.lbl_api = QLabel("")
        self.lbl_api.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        self.entry_api = QLineEdit()
        self.entry_api.setEchoMode(QLineEdit.EchoMode.Password)
        self.entry_api.setMinimumHeight(40)
        self.btn_load_api = QPushButton("")
        self.btn_load_api.setMinimumHeight(40)
        self.btn_load_api.clicked.connect(self.load_api_from_file)
        
        api_layout.addWidget(self.lbl_api)
        api_layout.addWidget(self.entry_api)
        api_layout.addWidget(self.btn_load_api)
        content_layout.addWidget(api_card)

        # I/O CARD
        io_card = QFrame(); io_card.setObjectName("Card")
        io_layout = QVBoxLayout(io_card)
        
        i_layout = QHBoxLayout()
        self.btn_input = QPushButton("")
        self.btn_input.setMinimumHeight(40)
        self.btn_input.setMinimumWidth(160)
        self.btn_input.clicked.connect(self.browse_input)
        self.entry_input = QLineEdit()
        self.entry_input.setMinimumHeight(40)
        i_layout.addWidget(self.btn_input)
        i_layout.addWidget(self.entry_input)
        
        o_layout = QHBoxLayout()
        self.btn_output = QPushButton("")
        self.btn_output.setMinimumHeight(40)
        self.btn_output.setMinimumWidth(160)
        self.btn_output.clicked.connect(self.browse_output)
        self.entry_output = QLineEdit()
        self.entry_output.setMinimumHeight(40)
        o_layout.addWidget(self.btn_output)
        o_layout.addWidget(self.entry_output)
        
        io_layout.addLayout(i_layout)
        io_layout.addLayout(o_layout)
        content_layout.addWidget(io_card)
        
        # OPTIONS
        opt_layout = QHBoxLayout()
        self.chk_solve = QCheckBox("")
        self.chk_solve.toggled.connect(self.update_options)
        self.chk_cover = QCheckBox("")
        self.chk_cover.toggled.connect(self.update_options)
        self.chk_merge_pages = QCheckBox("")
        self.chk_merge_pages.toggled.connect(self.update_options)
        
        opt_layout.addWidget(self.chk_solve)
        opt_layout.addWidget(self.chk_cover)
        opt_layout.addWidget(self.chk_merge_pages)
        opt_layout.addStretch()
        content_layout.addLayout(opt_layout)
        
        # BUTTONS
        btn_layout = QHBoxLayout()
        self.btn_start = QPushButton("")
        self.btn_start.setObjectName("Primary")
        self.btn_start.setMinimumHeight(60)
        self.btn_start.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        self.btn_start.clicked.connect(self.start_processing)
        
        self.btn_stop = QPushButton("")
        self.btn_stop.setObjectName("Danger")
        self.btn_stop.setMinimumHeight(60)
        self.btn_stop.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_processing)
        
        btn_layout.addWidget(self.btn_start)
        btn_layout.addWidget(self.btn_stop)
        content_layout.addLayout(btn_layout)

        # TIMER
        self.lbl_timer = QLabel("")
        self.lbl_timer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_timer.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        self.lbl_timer.setStyleSheet("color: #3a7ebf;")
        content_layout.addWidget(self.lbl_timer)

        # LOG
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Consolas", 13))
        self.log_box.setObjectName("Card")
        content_layout.addWidget(self.log_box)

        # FOOTER
        build_date = get_build_date()
        self.build_str = build_date.strftime("%Y%m%d")
        self.year_str = build_date.strftime("%Y")
        
        self.lbl_footer = QLabel("")
        self.lbl_footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_footer.setStyleSheet("color: gray;")
        content_layout.addWidget(self.lbl_footer)
        
        self.load_saved_api_key()

    def update_options(self):
        self.solve_var = self.chk_solve.isChecked()
        self.cover_var = self.chk_cover.isChecked()
        self.merge_pages_var = self.chk_merge_pages.isChecked()

    def update_ui_texts(self):
        self.setWindowTitle(self.t("title") + " v" + CURRENT_VERSION)
        self.lbl_toolbar.setText(self.t("toolbar"))
        self.btn_merge.setText(self.t("merge_pdf"))
        
        self.btn_mode_single.setText(self.t("mode_single"))
        self.btn_mode_batch.setText(self.t("mode_batch"))
        
        self.lbl_api.setText(self.t("api_key"))
        self.entry_api.setPlaceholderText(self.t("api_placeholder"))
        self.btn_load_api.setText(self.t("load_api"))
        
        self.chk_solve.setText(self.t("solve_opt"))
        self.chk_cover.setText(self.t("cover_opt"))
        self.chk_merge_pages.setText(self.t("merge_opt"))
        
        self.btn_start.setText(self.t("start"))
        self.btn_stop.setText(self.t("stop"))
        
        self.btn_output.setText(self.t("output_dir"))
        self.entry_output.setPlaceholderText(self.t("output_ph"))
        
        self.change_mode(self.mode_group.checkedButton())
        
        self.lbl_timer.setText(self.t("timer_init"))
        self.lbl_footer.setText(self.t("dev_footer", CURRENT_VERSION, self.build_str, self.year_str))
        
        if self.log_box.toPlainText() == "":
            self.write_log(self.t("log_ready"))

    def change_language(self, index):
        new_lang = "VN" if index == 1 else "EN"
        if new_lang != self.current_lang:
            self.current_lang = new_lang
            self.save_config()
            self.update_ui_texts()

    def change_mode(self, btn):
        if not btn: return
        self.entry_input.clear()
        if btn == self.btn_mode_single:
            self.btn_input.setText(self.t("select_pdf"))
            self.entry_input.setPlaceholderText(self.t("input_pdf_ph"))
        else:
            self.btn_input.setText(self.t("input_dir_btn"))
            self.entry_input.setPlaceholderText(self.t("input_dir_ph"))

    def closeEvent(self, event):
        key = self.entry_api.text().strip()
        if key: self.save_api_key(key)
        self.stop_event.set()
        event.accept()

    def save_api_key(self, key):
        try:
            if not os.path.exists(CONFIG_DIR): os.makedirs(CONFIG_DIR)
            with open(API_KEY_FILE, "w", encoding="utf-8") as f: f.write(key)
        except: pass

    def load_saved_api_key(self):
        if os.path.exists(API_KEY_FILE):
            try:
                with open(API_KEY_FILE, "r", encoding="utf-8") as f:
                    key = f.read().strip()
                    if key: self.entry_api.setText(key)
            except: pass

    def load_api_from_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load API Key", "", "Text Files (*.txt)")
        if path:
            try:
                with open(path, "r") as f: self.entry_api.setText(f.read().strip())
                self.write_log("[+] API Key loaded.")
            except: pass

    def write_log(self, text):
        self.log_box.append(text)
        scroll = self.log_box.verticalScrollBar()
        scroll.setValue(scroll.maximum())

    def browse_input(self):
        if self.btn_mode_single.isChecked():
            path, _ = QFileDialog.getOpenFileName(self, self.t("select_pdf"), "", "PDF (*.pdf)")
        else:
            path = QFileDialog.getExistingDirectory(self, self.t("input_dir_btn"))
        if path: self.entry_input.setText(path)

    def browse_output(self):
        path = QFileDialog.getExistingDirectory(self, self.t("output_dir"))
        if path: self.entry_output.setText(path)

    def stop_processing(self):
        self.write_log(self.t("log_stop_cmd"))
        self.stop_event.set()
        self.btn_stop.setEnabled(False)

    def start_processing(self):
        api_key = self.entry_api.text().strip()
        input_path = self.entry_input.text().strip()
        output_dir = self.entry_output.text().strip()
        
        if not api_key:
            self.write_log(self.t("err_api"))
            return
        if not input_path or not os.path.exists(input_path):
            self.write_log(self.t("err_input"))
            return

        if not output_dir:
            if os.path.isdir(input_path):
                output_dir = input_path
            else:
                output_dir = os.path.dirname(input_path) 
            self.write_log(self.t("msg_auto_out", output_dir))
        elif not os.path.exists(output_dir):
            self.write_log(self.t("err_output"))
            return

        self.save_api_key(api_key)
        self.start_time = datetime.now()
        self.timer_running = True
        self.timer.start(1000)

        self.btn_start.setEnabled(False)
        self.btn_input.setEnabled(False)
        self.btn_output.setEnabled(False)
        self.btn_mode_single.setEnabled(False)
        self.btn_mode_batch.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.stop_event.clear()
        
        genai.configure(api_key=api_key)
        
        mode = self.t("mode_single") if self.btn_mode_single.isChecked() else self.t("mode_batch")
        self.worker = WorkerThread(self, input_path, output_dir, mode)
        self.worker.log_signal.connect(self.write_log)
        self.worker.finished_signal.connect(self.reset_ui)
        self.worker.start()

    def reset_ui(self):
        self.timer_running = False
        self.timer.stop()
        self.btn_start.setEnabled(True)
        self.btn_input.setEnabled(True)
        self.btn_output.setEnabled(True)
        self.btn_mode_single.setEnabled(True)
        self.btn_mode_batch.setEnabled(True)
        self.btn_stop.setEnabled(False)

    def update_timer_ui(self):
        if self.timer_running and self.start_time:
            delta = datetime.now() - self.start_time
            m, s = divmod(int(delta.total_seconds()), 60)
            self.lbl_timer.setText(self.t("timer", m, s))

    def check_for_updates(self):
        try:
            req = urllib.request.Request(GITHUB_API_URL, headers={'User-Agent': 'PDFScan2Word-App'})
            with urllib.request.urlopen(req, timeout=5) as r:
                data = json.loads(r.read().decode('utf-8'))
                latest = data.get('tag_name', '').lstrip('v')
            if latest > CURRENT_VERSION:
                self.write_log(f"\n[*] Cập nhật mới có sẵn: v{latest} - Vui lòng tải tại Github.")
        except: pass

    def open_merge_popup(self):
        self.merge_window = MergeWindow(self)
        self.merge_window.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFOCRApp()
    window.show()
    sys.exit(app.exec())
