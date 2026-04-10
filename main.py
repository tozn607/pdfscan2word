import os
import urllib.request
import webbrowser
import sys
import time
import threading
import tkinter as tk
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Pt, Cm
from datetime import datetime
import json
from PIL import Image, ImageEnhance

# --- MACOS CRASH FIX ---
if sys.platform == "darwin":
    os.environ["PATH"] += os.pathsep + "/usr/local/bin" + os.pathsep + "/opt/homebrew/bin"
# --------------------------------------------------------

import customtkinter as ctk
from tkinter import filedialog
import fitz 
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import pypandoc

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except ImportError:
    print("[!] CẢNH BÁO: Chưa cài đặt pillow-heif. Không thể đọc file HEIC.")

CURRENT_VERSION = "2.0.0"
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

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

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

class LanguageSelectorPopup(ctk.CTkToplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.title("Select Language / Chọn ngôn ngữ")
        self.geometry("450x300")
        self.resizable(False, False)
        
        # Center the window relative to parent
        self.update_idletasks()
        try:
            x = parent.winfo_x() + (parent.winfo_width() // 2) - (450 // 2)
            y = parent.winfo_y() + (parent.winfo_height() // 2) - (300 // 2)
            self.geometry(f"+{x}+{y}")
        except:
            pass
        
        self.attributes("-topmost", True)
        self.grab_set()
        self.callback = callback
        
        lbl = ctk.CTkLabel(self, text="Welcome! Please choose your preferred language.\nChào mừng! Vui lòng chọn ngôn ngữ của bạn.", font=("Segoe UI", 15, "bold"))
        lbl.pack(pady=(35, 20))
        
        self.lang_var = ctk.StringVar(value="EN")
        
        r_en = ctk.CTkRadioButton(self, text="🇺🇸 English", variable=self.lang_var, value="EN", font=("Segoe UI", 16))
        r_en.pack(pady=10)
        
        r_vn = ctk.CTkRadioButton(self, text="🇻🇳 Tiếng Việt", variable=self.lang_var, value="VN", font=("Segoe UI", 16))
        r_vn.pack(pady=10)
        
        btn = ctk.CTkButton(self, text="Continue / Tiếp tục", font=("Segoe UI", 15, "bold"), height=45, corner_radius=10, fg_color="#2b7a4b", hover_color="#1e5c37", command=self.on_save)
        btn.pack(pady=25)
        
        self.protocol("WM_DELETE_WINDOW", self.on_save_default)
        
    def on_save(self):
        self.callback(self.lang_var.get())
        self.destroy()
        
    def on_save_default(self):
        self.callback("EN") # default fallback
        self.destroy()

class PDFOCRApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.current_lang = "VN" # Default
        self.load_config()

        self.title(self.t("title") + " v" + CURRENT_VERSION)
        self.geometry("840x880")
        self.minsize(840, 880)

        # Base font configs
        self.font_main = ("Segoe UI", 14)
        self.font_bold = ("Segoe UI", 14, "bold")
        self.font_large = ("Segoe UI", 16, "bold")
        self.card_kwargs = {"fg_color": ("gray90", "gray13"), "corner_radius": 12}

        self.stop_event = threading.Event()
        self.start_time = None
        self.timer_running = False
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.build_ui()
        self.update_ui_texts() 
        
        if not os.path.exists(CONFIG_JSON_FILE):
            self.after(200, self.show_language_popup_first_time)
            
        threading.Thread(target=self.check_for_updates, daemon=True).start()

    def t(self, key, *args):
        text = STRINGS.get(self.current_lang, STRINGS["VN"]).get(key, key)
        if args:
            return text.format(*args)
        return text

    def load_config(self):
        if os.path.exists(CONFIG_JSON_FILE):
            try:
                with open(CONFIG_JSON_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.current_lang = data.get("lang", "VN")
            except:
                pass

    def save_config(self):
        if not os.path.exists(CONFIG_DIR):
            os.makedirs(CONFIG_DIR)
        try:
            with open(CONFIG_JSON_FILE, "w", encoding="utf-8") as f:
                json.dump({"lang": self.current_lang}, f)
        except Exception as e:
            print(f"Error saving config: {e}")

    def show_language_popup_first_time(self):
        LanguageSelectorPopup(self, self.on_language_selected_first_time)
        
    def on_language_selected_first_time(self, lang):
        self.current_lang = lang
        self.save_config()
        self.lang_var.set("EN (English)" if lang == "EN" else "VN (Tiếng Việt)")
        self.update_ui_texts()

    def build_ui(self):
        # --- 0. MENU BAR TRÊN CÙNG ---
        self.frame_menu = ctk.CTkFrame(self, height=65, corner_radius=0, fg_color=("gray85", "gray20"))
        self.frame_menu.pack(side="top", fill="x")
        self.frame_menu.pack_propagate(False)
        
        self.lbl_toolbar = ctk.CTkLabel(self.frame_menu, text="", font=self.font_bold)
        self.lbl_toolbar.pack(side="left", padx=(20, 10), pady=15)

        self.btn_merge = ctk.CTkButton(self.frame_menu, text="", font=self.font_bold, height=45, corner_radius=10,
                                       fg_color="#2b7a4b", hover_color="#1e5c37", text_color="white", 
                                       command=self.open_merge_popup)
        self.btn_merge.pack(side="left", padx=5, pady=10)

        # Dropdown Language Selector
        self.lbl_lang_icon = ctk.CTkLabel(self.frame_menu, text="🌐", font=("Segoe UI", 18))
        self.lbl_lang_icon.pack(side="right", padx=(5, 20), pady=15)
        
        self.lang_var = ctk.StringVar(value="EN (English)" if self.current_lang == "EN" else "VN (Tiếng Việt)")
        self.lang_menu = ctk.CTkOptionMenu(self.frame_menu, values=["EN (English)", "VN (Tiếng Việt)"],
                                           variable=self.lang_var, command=self.change_language,
                                           font=self.font_main, width=150, height=38)
        self.lang_menu.pack(side="right", padx=5, pady=12)

        # Container for main content
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=25, pady=10)

        # --- 1. CHỌN CHẾ ĐỘ (MODE SWITCHER) ---
        self.frame_mode = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.frame_mode.pack(pady=(10, 15), fill="x")
        
        self.mode_var = ctk.StringVar(value="1")
        self.mode_selector = ctk.CTkSegmentedButton(self.frame_mode, 
                                                    values=["1", "2"], 
                                                    variable=self.mode_var, 
                                                    command=self.change_mode,
                                                    height=45, font=self.font_bold)
        self.mode_selector.pack(expand=True, fill="x", padx=10)

        # --- 2. API KEY CARD ---
        self.frame_api = ctk.CTkFrame(self.main_container, **self.card_kwargs)
        self.frame_api.pack(pady=10, fill="x", ipady=10)
        
        self.lbl_api = ctk.CTkLabel(self.frame_api, text="", width=120, anchor="w", font=self.font_bold)
        self.lbl_api.pack(side="left", padx=15, pady=10)
        
        self.entry_api = ctk.CTkEntry(self.frame_api, show="*", height=45, placeholder_text="", font=self.font_main, corner_radius=10)
        self.entry_api.pack(side="left", padx=10, pady=10, expand=True, fill="x")
        
        self.btn_load_api = ctk.CTkButton(self.frame_api, text="", width=140, height=45, font=self.font_bold, corner_radius=10, command=self.load_api_from_file)
        self.btn_load_api.pack(side="right", padx=15, pady=10)
        self.load_saved_api_key()

        # --- 3. INPUT / OUTPUT CARD ---
        self.frame_io = ctk.CTkFrame(self.main_container, **self.card_kwargs)
        self.frame_io.pack(pady=10, fill="x", ipady=10)
        
        # Input
        self.io_inner1 = ctk.CTkFrame(self.frame_io, fg_color="transparent")
        self.io_inner1.pack(fill="x", padx=10, pady=(10, 5))
        self.btn_input = ctk.CTkButton(self.io_inner1, text="", width=160, height=45, font=self.font_bold, corner_radius=10, command=self.browse_input)
        self.btn_input.pack(side="left", padx=5)
        self.entry_input = ctk.CTkEntry(self.io_inner1, height=45, placeholder_text="", font=self.font_main, corner_radius=10)
        self.entry_input.pack(side="left", padx=10, expand=True, fill="x")

        # Output
        self.io_inner2 = ctk.CTkFrame(self.frame_io, fg_color="transparent")
        self.io_inner2.pack(fill="x", padx=10, pady=(5, 10))
        self.btn_output = ctk.CTkButton(self.io_inner2, text="", width=160, height=45, font=self.font_bold, corner_radius=10, command=self.browse_output)
        self.btn_output.pack(side="left", padx=5)
        self.entry_output = ctk.CTkEntry(self.io_inner2, height=45, placeholder_text="", font=self.font_main, corner_radius=10)
        self.entry_output.pack(side="left", padx=10, expand=True, fill="x")

        # --- CÁC TÙY CHỌN MỞ RỘNG ---
        self.frame_options = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.frame_options.pack(pady=15, fill="x")
        
        self.solve_var = ctk.BooleanVar(value=False)
        self.chk_solve = ctk.CTkCheckBox(self.frame_options, text="", variable=self.solve_var, font=self.font_bold)
        self.chk_solve.pack(side="left", padx=10)

        self.cover_var = ctk.BooleanVar(value=False)
        self.chk_cover = ctk.CTkCheckBox(self.frame_options, text="", variable=self.cover_var, font=self.font_bold)
        self.chk_cover.pack(side="left", padx=20)

        self.merge_pages_var = ctk.BooleanVar(value=False)
        self.chk_merge_pages = ctk.CTkCheckBox(self.frame_options, text="", variable=self.merge_pages_var, font=self.font_bold)
        self.chk_merge_pages.pack(side="left", padx=10)

        # --- 4. NÚT ĐIỀU KHIỂN CHIẾM VỊ TRÍ LỚN ---
        self.frame_buttons = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.frame_buttons.pack(pady=15, fill="x")
        
        self.btn_start = ctk.CTkButton(self.frame_buttons, text="", font=("Segoe UI", 18, "bold"), height=60, corner_radius=12,
                                       fg_color="#1f538d", hover_color="#14375e", command=self.start_processing)
        self.btn_start.pack(side="left", expand=True, fill="x", padx=10)

        self.btn_stop = ctk.CTkButton(self.frame_buttons, text="", font=("Segoe UI", 18, "bold"), height=60, corner_radius=12,
                                      fg_color="#a83232", hover_color="#7a2121", state="disabled", command=self.stop_processing)
        self.btn_stop.pack(side="right", expand=True, fill="x", padx=10)

        # --- ĐỒNG HỒ ĐẾM GIỜ ---
        self.lbl_timer = ctk.CTkLabel(self.main_container, text="", font=("Segoe UI", 15, "bold"), text_color="#3a7ebf")
        self.lbl_timer.pack(pady=(0, 10))

        # --- 5. LOG TRUNG TÂM ---
        self.log_box = ctk.CTkTextbox(self.main_container, height=220, font=("Consolas", 14), **self.card_kwargs)
        self.log_box.pack(pady=5, fill="both", expand=True)
        self.log_box.insert("0.0", self.t("log_ready") + "\n")
        self.log_box.configure(state="disabled")

        # --- 6. FOOTER ---
        build_date = get_build_date()
        self.build_str = build_date.strftime("%Y%m%d")
        self.year_str = build_date.strftime("%Y")

        self.lbl_footer = ctk.CTkLabel(self.main_container, text="", text_color="gray50", font=("Segoe UI", 13), cursor="hand2")
        self.lbl_footer.pack(side="bottom", pady=(10, 0))
        self.lbl_footer.bind("<Button-1>", self.show_about_popup)


    def update_ui_texts(self):
        self.title(self.t("title") + " v" + CURRENT_VERSION)
        self.lbl_toolbar.configure(text=self.t("toolbar"))
        self.btn_merge.configure(text=self.t("merge_pdf"))
        self.lbl_api.configure(text=self.t("api_key"))
        self.btn_load_api.configure(text=self.t("load_api"))
        self.entry_api.configure(placeholder_text=self.t("api_placeholder"))
        
        old_mode_idx = 0 if self.mode_var.get() in ("1", self.t("mode_single")) else 1
        new_values = [self.t("mode_single"), self.t("mode_batch")]
        self.mode_selector.configure(values=new_values)
        self.mode_var.set(new_values[old_mode_idx])
        
        self.chk_solve.configure(text=self.t("solve_opt"))
        self.chk_cover.configure(text=self.t("cover_opt"))
        self.chk_merge_pages.configure(text=self.t("merge_opt"))
        self.btn_start.configure(text=self.t("start"))
        self.btn_stop.configure(text=self.t("stop"))
        
        self.change_mode(self.mode_var.get())
        self.btn_output.configure(text=self.t("output_dir"))
        self.entry_output.configure(placeholder_text=self.t("output_ph"))
        self.lbl_timer.configure(text=self.t("timer_init"))
        
        footer_text = self.t("dev_footer", CURRENT_VERSION, self.build_str, self.year_str)
        self.lbl_footer.configure(text=footer_text)
        
    def change_language(self, choice):
        new_lang = "VN" if "VN" in choice else "EN"
        if new_lang != self.current_lang:
            self.current_lang = new_lang
            self.save_config()
            self.update_ui_texts()

    def change_mode(self, selected_mode):
        self.entry_input.delete(0, "end")
        if selected_mode == self.t("mode_single"):
            self.btn_input.configure(text=self.t("select_pdf"))
            self.entry_input.configure(placeholder_text=self.t("input_pdf_ph"))
        else:
            self.btn_input.configure(text=self.t("input_dir_btn"))
            self.entry_input.configure(placeholder_text=self.t("input_dir_ph"))

    def on_closing(self):
        current_key = self.entry_api.get().strip()
        if current_key:
            self.save_api_key(current_key)
        self.stop_event.set() 
        self.destroy() 
        os._exit(0)

    def save_api_key(self, key):
        try:
            if not os.path.exists(CONFIG_DIR):
                os.makedirs(CONFIG_DIR)
            with open(API_KEY_FILE, "w", encoding="utf-8") as f: 
                f.write(key)
        except Exception as e:
            print(f"Lỗi khi lưu key: {e}")

    def load_saved_api_key(self):
        if os.path.exists(API_KEY_FILE):
            try:
                with open(API_KEY_FILE, "r", encoding="utf-8") as f:
                    key = f.read().strip()
                    if key: 
                        self.entry_api.delete(0, "end")
                        self.entry_api.insert(0, key)
            except Exception as e:
                print(f"Lỗi khi load key: {e}")

    def load_api_from_file(self):
        filepath = filedialog.askopenfilename(title="Chọn file txt chứa API Key", filetypes=[("Text Files", "*.txt")])
        if filepath:
            with open(filepath, "r") as f:
                key = f.read().strip()
                self.entry_api.delete(0, "end")
                self.entry_api.insert(0, key)
                self.write_log("[+] API Key loaded.")

    def write_log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def browse_input(self):
        current_mode = self.mode_var.get()
        if current_mode == self.t("mode_single"):
            path = filedialog.askopenfilename(title=self.t("select_pdf"), filetypes=[("PDF", "*.pdf")])
        else:
            path = filedialog.askdirectory(title=self.t("input_dir_btn"))
            
        if path:
            self.entry_input.delete(0, "end")
            self.entry_input.insert(0, path)

    def browse_output(self):
        folder_path = filedialog.askdirectory(title=self.t("output_dir"))
        if folder_path:
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, folder_path)

    def stop_processing(self):
        self.write_log(self.t("log_stop_cmd"))
        self.stop_event.set()

    def start_processing(self):
        api_key = self.entry_api.get().strip()
        input_path = self.entry_input.get().strip()
        output_dir = self.entry_output.get().strip()
        current_mode = self.mode_var.get()

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
        self.update_timer_ui()

        self.btn_start.configure(state="disabled")
        self.btn_input.configure(state="disabled")
        self.btn_output.configure(state="disabled")
        self.mode_selector.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        
        self.stop_event.clear()
        genai.configure(api_key=api_key)

        thread = threading.Thread(target=self.process_documents, args=(input_path, output_dir, current_mode))
        thread.daemon = True
        thread.start()

    def process_documents(self, input_path, output_dir, mode):
        if mode == self.t("mode_single"):
            pdf_files = [input_path]
        else:
            pdf_files = [os.path.join(input_path, f) for f in os.listdir(input_path) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            self.write_log(self.t("log_no_pdf"))
            self.reset_ui()
            return

        total_files = len(pdf_files)
        self.write_log(self.t("log_start_batch", total_files))
        model = genai.GenerativeModel('gemini-3.1-flash-lite-preview')

        active_prompt = PROMPT_EN if self.current_lang == "EN" else PROMPT_VN
        if self.solve_var.get():
            active_prompt += self.t("prompt_solve_ext")

        for file_idx, pdf_path in enumerate(pdf_files):
            if self.stop_event.is_set(): break

            pdf_filename = os.path.basename(pdf_path)
            base_name = os.path.splitext(pdf_filename)[0]
            suffix = "_hoanthien" if self.current_lang == "VN" else "_completed"
            output_docx_path = os.path.join(output_dir, f"{base_name}{suffix}.docx")

            self.write_log(self.t("log_split", file_idx + 1, total_files, pdf_filename))
            
            images = []
            try: 
                pandoc_exe = os.path.join(CONFIG_DIR, "pandoc" + (".exe" if sys.platform == "win32" else ""))
                if os.path.exists(pandoc_exe):
                    os.environ['PYPANDOC_PANDOC'] = pandoc_exe

                try:
                    pypandoc.get_pandoc_version()
                except OSError:
                    self.write_log(self.t("log_pandoc"))
                    pypandoc.download_pandoc(targetfolder=CONFIG_DIR, download_folder=CONFIG_DIR)
                    os.environ['PYPANDOC_PANDOC'] = pandoc_exe

                doc = fitz.open(pdf_path)
                for page in doc:
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    images.append(img)
                    
            except Exception as e:
                self.write_log(self.t("log_err_read", e))
                continue 

            if not images:
                continue

            start_idx = 0
            if self.cover_var.get():
                cover_path = os.path.join(output_dir, f"{base_name}_cover.jpg")
                try:
                    images[0].save(cover_path, "JPEG")
                    self.write_log(self.t("log_save_cover", f"{base_name}_cover.jpg"))
                except Exception as e:
                    self.write_log(self.t("log_err_cover", e))
                start_idx = 1 

            processed_images = []
            if self.merge_pages_var.get():
                self.write_log(self.t("log_merge_pages"))
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
                if self.stop_event.is_set():
                    break

                self.write_log(self.t("log_read_block", i+1, len(processed_images)))
                
                for attempt in range(max_retries):
                    if self.stop_event.is_set(): break
                    try:
                        response = model.generate_content([active_prompt, image], safety_settings=safety_config)
                        try:
                            text_result = response.text
                        except ValueError:
                            self.write_log(self.t("log_reject", i+1))
                            text_result = f"> **[COPYRIGHT BLOCK: PAGE {i+1}]**"
                            full_markdown_content += f"\n\n\n\n{text_result}\n\n"
                            break
                        
                        text_result = text_result.replace("[^", f"[^p{i}_")
                        full_markdown_content += f"\n\n\n\n{text_result}\n\n"
                        
                        if i < len(images) - 1 and not self.stop_event.is_set():
                            time.sleep(3) 
                        break 
                        
                    except Exception as e:
                        self.write_log(self.t("log_err_attempt", attempt+1, max_retries, e))
                        if attempt < max_retries - 1:
                            time.sleep(60) 
                        else:
                            self.write_log(self.t("log_skip", i+1))

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
                        self.write_log(self.t("log_format_err", e_docx))

                    if self.stop_event.is_set():
                        self.write_log(self.t("log_save_draft", output_docx_path))
                    else:
                        self.write_log(self.t("log_save_ok", output_docx_path))
                except Exception as e:
                    md_path = output_docx_path.replace('.docx', '.md')
                    with open(md_path, 'w', encoding='utf-8') as f: 
                        f.write(full_markdown_content)
                    self.write_log(self.t("log_rescue", md_path))
            
            if self.stop_event.is_set():
                break

        if self.stop_event.is_set():
            self.write_log(self.t("log_cancel"))
        else:
            self.write_log(self.t("log_done"))
        
        self.reset_ui()

    def reset_ui(self):
        self.timer_running = False 
        self.btn_start.configure(state="normal")
        self.btn_input.configure(state="normal")
        self.btn_output.configure(state="normal")
        self.mode_selector.configure(state="normal")
        self.btn_stop.configure(state="disabled")

    def open_merge_popup(self):
        self.merge_window = ctk.CTkToplevel(self)
        self.merge_window.title(self.t("merge_title"))
        self.merge_window.geometry("660x520")
        self.merge_window.attributes("-topmost", True)
        self.merge_window.grab_set()

        self.selected_images_list = []

        frame_controls = ctk.CTkFrame(self.merge_window, fg_color="transparent")
        frame_controls.pack(pady=10, padx=20, fill="x")

        btn_add = ctk.CTkButton(frame_controls, text=self.t("add_img"), width=120, height=40, corner_radius=10, font=self.font_bold, command=self.add_images)
        btn_add.pack(side="left", padx=5)

        btn_up = ctk.CTkButton(frame_controls, text=self.t("up"), width=80, height=40, corner_radius=10, font=self.font_bold, fg_color="#454545", hover_color="#5a5a5a", command=self.move_up)
        btn_up.pack(side="left", padx=5)
        
        btn_down = ctk.CTkButton(frame_controls, text=self.t("down"), width=80, height=40, corner_radius=10, font=self.font_bold, fg_color="#454545", hover_color="#5a5a5a", command=self.move_down)
        btn_down.pack(side="left", padx=5)

        btn_clear = ctk.CTkButton(frame_controls, text=self.t("clear"), width=90, height=40, corner_radius=10, font=self.font_bold, fg_color="#a83232", hover_color="#7a2121", command=self.clear_images)
        btn_clear.pack(side="right", padx=5)

        self.compress_var = ctk.BooleanVar(value=True) 
        self.enhance_var = ctk.BooleanVar(value=False)
        
        frame_options = ctk.CTkFrame(self.merge_window, fg_color="transparent")
        frame_options.pack(pady=(0, 5), padx=20, fill="x")
        
        chk_compress = ctk.CTkCheckBox(frame_options, text=self.t("compress"), variable=self.compress_var, font=self.font_bold)
        chk_compress.pack(side="left", padx=5)
        
        chk_enhance = ctk.CTkCheckBox(frame_options, text=self.t("enhance"), variable=self.enhance_var, font=self.font_bold)
        chk_enhance.pack(side="left", padx=20)

        list_frame = tk.Frame(self.merge_window, bg="#343638", bd=0, highlightthickness=0)
        list_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.img_listbox = tk.Listbox(list_frame, font=("Consolas", 14), bg="#2b2b2b", fg="#ffffff", 
                                      selectbackground="#1f538d", highlightthickness=0, borderwidth=0,
                                      yscrollcommand=scrollbar.set)
        self.img_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.img_listbox.yview)

        btn_export = ctk.CTkButton(self.merge_window, text=self.t("export_pdf"), font=self.font_large, height=50, corner_radius=12, fg_color="#2b7a4b", hover_color="#1e5c37", command=self.export_to_pdf)
        btn_export.pack(pady=15)

    def add_images(self):
        paths = filedialog.askopenfilenames(filetypes=[("Image Files", "*.png *.jpg *.jpeg *.bmp *.heic *.heif")])
        if paths:
            self.selected_images_list.extend(paths)
            self.update_image_listbox()

    def move_up(self):
        try:
            idx = self.img_listbox.curselection()[0]
            if idx > 0:
                self.selected_images_list[idx-1], self.selected_images_list[idx] = self.selected_images_list[idx], self.selected_images_list[idx-1]
                self.update_image_listbox()
                self.img_listbox.select_set(idx-1) 
        except IndexError:
            pass 

    def move_down(self):
        try:
            idx = self.img_listbox.curselection()[0]
            if idx < len(self.selected_images_list) - 1:
                self.selected_images_list[idx+1], self.selected_images_list[idx] = self.selected_images_list[idx], self.selected_images_list[idx+1]
                self.update_image_listbox()
                self.img_listbox.select_set(idx+1) 
        except IndexError:
            pass

    def clear_images(self):
        self.selected_images_list.clear()
        self.update_image_listbox()

    def update_image_listbox(self):
        self.img_listbox.delete(0, 'end')
        for path in self.selected_images_list:
            self.img_listbox.insert('end', os.path.basename(path))

    def export_to_pdf(self):
        if not self.selected_images_list: return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        
        if save_path:
            self.img_listbox.delete(0, 'end')
            self.img_listbox.insert('end', self.t("exporting"))
            self.merge_window.update() 
            
            try:
                processed_imgs = []
                for p in self.selected_images_list:
                    img = Image.open(p).convert('RGB')
                    if self.enhance_var.get():
                        enhancer_contrast = ImageEnhance.Contrast(img)
                        img = enhancer_contrast.enhance(1.5)
                        enhancer_sharpness = ImageEnhance.Sharpness(img)
                        img = enhancer_sharpness.enhance(2.0)
                        enhancer_brightness = ImageEnhance.Brightness(img)
                        img = enhancer_brightness.enhance(1.1)

                    if self.compress_var.get():
                        max_width = 1200
                        if img.width > max_width:
                            ratio = max_width / img.width
                            new_size = (max_width, int(img.height * ratio))
                            img = img.resize(new_size, Image.Resampling.LANCZOS)
                            
                    processed_imgs.append(img)
                
                processed_imgs[0].save(save_path, save_all=True, append_images=processed_imgs[1:])
                
                self.img_listbox.delete(0, 'end')
                self.img_listbox.insert('end', self.t("merge_success", save_path))
                
            except Exception as e: 
                self.img_listbox.delete(0, 'end')
                self.img_listbox.insert('end', self.t("merge_err", e))

    def update_timer_ui(self):
        if self.timer_running and self.start_time:
            delta = datetime.now() - self.start_time
            minutes, seconds = divmod(int(delta.total_seconds()), 60)
            self.lbl_timer.configure(text=self.t("timer", minutes, seconds))
            self.after(1000, self.update_timer_ui)

    def check_for_updates(self):
        try:
            req = urllib.request.Request(GITHUB_API_URL, headers={'User-Agent': 'PDFScan2Word-App'})
            with urllib.request.urlopen(req, timeout=5) as response:
                data = json.loads(response.read().decode('utf-8'))
                latest_tag = data.get('tag_name', '')
                latest_version = latest_tag.lstrip('v')

            if latest_version > CURRENT_VERSION:
                self.after(500, self.show_update_popup, latest_version)
        except Exception as e:
            print(f"[!] Không thể kiểm tra cập nhật: {e}")

    def show_update_popup(self, latest_version):
        dialog = ctk.CTkToplevel(self)
        dialog.title(self.t("update_title"))
        dialog.geometry("420x220")
        dialog.attributes("-topmost", True)
        dialog.grab_set()
        
        lbl = ctk.CTkLabel(dialog, text=self.t("update_msg", latest_version, CURRENT_VERSION), font=self.font_main)
        lbl.pack(pady=30)
        
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        btn_yes = ctk.CTkButton(btn_frame, text=self.t("btn_yes"), fg_color="#2b7a4b", hover_color="#1e5c37", 
                                font=self.font_bold, height=40, corner_radius=8,
                                command=lambda: [webbrowser.open(RELEASES_URL), dialog.destroy()])
        btn_yes.pack(side="left", padx=10)
        
        btn_no = ctk.CTkButton(btn_frame, text=self.t("btn_no"), fg_color="gray", hover_color="darkgray", 
                               font=self.font_bold, height=40, corner_radius=8, command=dialog.destroy)
        btn_no.pack(side="left", padx=10)

    def show_about_popup(self, event=None):
        about_window = ctk.CTkToplevel(self)
        about_window.title("")
        about_window.geometry("340x360")
        about_window.resizable(False, False)
        about_window.attributes("-topmost", True)
        about_window.grab_set()

        lbl_logo = ctk.CTkLabel(about_window, text="📄", font=("Segoe UI", 60))
        lbl_logo.pack(pady=(20, 10))

        lbl_name = ctk.CTkLabel(about_window, text=self.t("title"), font=("Segoe UI", 20, "bold"))
        lbl_name.pack(pady=(0, 5))

        lbl_ver = ctk.CTkLabel(about_window, text=f"Version {CURRENT_VERSION} ({self.build_str})", font=("Segoe UI", 12), text_color="gray50")
        lbl_ver.pack(pady=(0, 15))

        lbl_copyright = ctk.CTkLabel(about_window, text="Thank you for using this app! ❤️\nCảm ơn bạn hiền đã sử dụng ứng dụng này!", font=("Segoe UI", 13))
        lbl_copyright.pack(pady=(0, 20))

        lbl_link = ctk.CTkLabel(about_window, text="Developed by @tozn607", text_color="#1f6aa5", font=("Segoe UI", 13, "bold", "underline"), cursor="hand2")
        lbl_link.pack(pady=(0, 20))
        lbl_link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/tozn607"))

if __name__ == "__main__":
    app = PDFOCRApp()
    app.mainloop()