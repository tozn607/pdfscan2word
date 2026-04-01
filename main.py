import os
import urllib.request
import webbrowser
import sys
import time
import threading
import tkinter as tk
from datetime import datetime

# --- MACOS CRASH FIX ---
if sys.platform == "darwin":
    os.environ["PATH"] += os.pathsep + "/usr/local/bin" + os.pathsep + "/opt/homebrew/bin"
# --------------------------------------------------------

import customtkinter as ctk
from tkinter import filedialog
from pdf2image import convert_from_path
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import pypandoc
from PIL import Image

# --- HỖ TRỢ ĐỊNH DẠNG HEIF/HEIC CỦA APPLE ---
try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except ImportError:
    print("[!] CẢNH BÁO: Chưa cài đặt pillow-heif. Không thể đọc file HEIC.")

# --- CẤU HÌNH PHIÊN BẢN VÀ CẬP NHẬT ---
CURRENT_VERSION = "1.1.1"
UPDATE_RAW_URL = "https://raw.githubusercontent.com/tozn607/pdfscan2word/main/version.txt"
RELEASES_URL = "https://github.com/tozn607/pdfscan2word/releases"

# Tự động tìm thư mục Home của máy tính và tạo folder ẩn .pdfscan2word
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".pdfscan2word")
API_KEY_FILE = os.path.join(CONFIG_DIR, "api_key.txt")

# Cấu hình UI
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

safety_config = {
    HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
}

prompt_template = r"""
Bạn là một chuyên gia số hóa và phục hồi tài liệu chuyên nghiệp. Dưới đây là hình ảnh scan của một trang tài liệu/giáo trình. 
Nhiệm vụ của bạn là trích xuất và phục hồi, làm sạch văn bản theo các quy tắc NGHIÊM NGẶT sau đây:
1. ƯU TIÊN SỐ 1 - BẢO TOÀN DANH SÁCH: Mọi mục bắt đầu bằng số (1., 2.), chữ cái (A., a.) BẮT BUỘC nằm ở dòng riêng. 
2. NGĂN TỰ ĐỘNG TẠO BULLET POINT (QUAN TRỌNG): 
   - Trình biên dịch sẽ tự động biến dấu gạch ngang đầu dòng thành dấu chấm đen. Để ngăn chặn điều này: NẾU bản gốc có dấu gạch ngang (-) ở đầu dòng, BẮT BUỘC phải dùng dấu gạch chéo ngược để thoát (escape): Viết là `\- ` thay vì `- `.
   - Vẫn giữ nguyên định dạng in nghiêng hoặc in đậm phía sau dấu gạch ngang nếu có (ví dụ: `\- *Chủ nghĩa duy vật lịch sử*`).
   - Các mục đánh số (1., 2.) hoặc chữ cái (a., b.) thì giữ nguyên, KHÔNG được tự ý chèn thêm dấu gạch ngang vào trước chúng.
3. NỐI DÒNG THÔNG MINH: Chỉ nối nếu dòng dưới là phần đứt đoạn của câu trên. 
4. ĐỊNH DẠNG: **In đậm** và *In nghiêng* đúng bản gốc. Tiêu đề lớn in đậm và đứng riêng.
5. ĐIỀN CHỮ THIẾU: Dựa vào ngữ cảnh chung để điền bù chữ khuất mép giấy. Xóa bỏ hoàn toàn ký tự rác.
6. LOẠI BỎ SỐ TRANG: TUYỆT ĐỐI KHÔNG ghi lại số trang. Hãy chủ động bỏ qua chúng.
7. XỬ LÝ CHÚ THÍCH (FOOTNOTE): NẾU trang tài liệu có chú thích ở dưới cùng:
   - Đặt mốc `[^1]`, `[^2]`... sát ngay sau từ/câu được chú thích.
   - Ghi nội dung chú thích ở tận cùng của văn bản theo cú pháp: `[^1]: Nội dung chú thích...`
8. THỤT ĐẦU DÒNG (INDENTATION): Nếu đoạn văn trong bản gốc có lùi vào ở dòng đầu tiên, BẮT BUỘC chèn cụm ký tự `&emsp;&emsp;` vào ngay vị trí bắt đầu của đoạn văn đó. TUYỆT ĐỐI KHÔNG dùng phím Space hoặc Tab.
Chỉ trả về văn bản bằng Markdown, không giải thích gì thêm.
"""

class PDFOCRApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Scanned Images to Word v" + CURRENT_VERSION)
        self.geometry("780x800")
        self.stop_event = threading.Event()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # --- 0. MENU BAR TRÊN CÙNG ---
        self.frame_menu = ctk.CTkFrame(self, height=45, corner_radius=0, fg_color=("gray85", "gray20"))
        self.frame_menu.pack(side="top", fill="x")
        
        self.lbl_toolbar = ctk.CTkLabel(self.frame_menu, text="Thanh công cụ:", font=("Arial", 13, "bold"))
        self.lbl_toolbar.pack(side="left", padx=(20, 10), pady=10)

        # Nút gộp ảnh với nền xanh lá nổi bật
        self.btn_merge = ctk.CTkButton(self.frame_menu, text="GỘP ẢNH THÀNH PDF", font=("Arial", 12, "bold"), 
                                       fg_color="#2b7a4b", hover_color="#1e5c37", text_color="white", 
                                       command=self.open_merge_popup)
        self.btn_merge.pack(side="left", padx=5, pady=10)

        # --- 1. CHỌN CHẾ ĐỘ (MODE SWITCHER) ---
        self.frame_mode = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_mode.pack(pady=(15, 5), padx=20, fill="x")
        
        self.mode_var = ctk.StringVar(value="Chế độ Đơn (1 File PDF)")
        self.mode_selector = ctk.CTkSegmentedButton(self.frame_mode, 
                                                    values=["Chế độ Đơn (1 File PDF)", "Chế độ Hàng loạt (Thư mục)"], 
                                                    variable=self.mode_var, 
                                                    command=self.change_mode)
        self.mode_selector.pack(side="left")

        # --- 2. API KEY ---
        self.frame_api = ctk.CTkFrame(self)
        self.frame_api.pack(pady=5, padx=20, fill="x")
        
        self.lbl_api = ctk.CTkLabel(self.frame_api, text="Google API Key:", width=100, anchor="w")
        self.lbl_api.pack(side="left", padx=10, pady=10)
        
        self.entry_api = ctk.CTkEntry(self.frame_api, show="*", width=380, placeholder_text="Nhập API Key...")
        self.entry_api.pack(side="left", padx=10, pady=10)
        
        self.btn_load_api = ctk.CTkButton(self.frame_api, text="Tải từ file .txt", width=120, command=self.load_api_from_file)
        self.btn_load_api.pack(side="left", padx=10, pady=10)
        self.load_saved_api_key()

        # --- 3. INPUT / OUTPUT ---
        self.frame_input = ctk.CTkFrame(self)
        self.frame_input.pack(pady=5, padx=20, fill="x")
        
        # Nút Input sẽ thay đổi text tùy theo Mode
        self.btn_input = ctk.CTkButton(self.frame_input, text="Chọn File PDF", width=120, command=self.browse_input)
        self.btn_input.pack(side="left", padx=10, pady=10)
        self.entry_input = ctk.CTkEntry(self.frame_input, width=500, placeholder_text="Đường dẫn đến 1 file PDF...")
        self.entry_input.pack(side="left", padx=10, pady=10)

        self.frame_output = ctk.CTkFrame(self)
        self.frame_output.pack(pady=5, padx=20, fill="x")
        self.btn_output = ctk.CTkButton(self.frame_output, text="Thư mục Output", width=120, command=self.browse_output)
        self.btn_output.pack(side="left", padx=10, pady=10)
        self.entry_output = ctk.CTkEntry(self.frame_output, width=500, placeholder_text="Trống = Tự động lưu cùng nơi với File/Thư mục Input")
        self.entry_output.pack(side="left", padx=10, pady=10)

        # --- CÁC TÙY CHỌN MỞ RỘNG ---
        self.frame_options = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_options.pack(pady=5, padx=20, fill="x")
        
        self.solve_var = ctk.BooleanVar(value=False)
        self.chk_solve = ctk.CTkCheckBox(self.frame_options, text="🤖 Yêu cầu AI Giải bài tập", variable=self.solve_var, font=("Arial", 12, "bold"))
        self.chk_solve.pack(side="left", padx=10)

        # Checkbox mới: Lưu riêng trang bìa
        self.cover_var = ctk.BooleanVar(value=False)
        self.chk_cover = ctk.CTkCheckBox(self.frame_options, text="🖼️ Lưu riêng trang bìa (Bỏ qua Scan trang 1)", variable=self.cover_var, font=("Arial", 12, "bold"))
        self.chk_cover.pack(side="left", padx=15)

        # --- 4. NÚT ĐIỀU KHIỂN ---
        self.frame_buttons = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_buttons.pack(pady=10)
        
        self.btn_start = ctk.CTkButton(self.frame_buttons, text="▶ BẮT ĐẦU XỬ LÝ", font=("Arial", 14, "bold"), command=self.start_processing)
        self.btn_start.pack(side="left", padx=10)

        self.btn_stop = ctk.CTkButton(self.frame_buttons, text="⏹ DỪNG LẠI", font=("Arial", 14, "bold"), fg_color="#a83232", hover_color="#7a2121", state="disabled", command=self.stop_processing)
        self.btn_stop.pack(side="left", padx=10)

        # --- 5. LOG TRUNG TÂM ---
        self.log_box = ctk.CTkTextbox(self, width=740, height=350, font=("Consolas", 13))
        self.log_box.pack(pady=10, padx=20)
        self.log_box.insert("0.0", "[*] Ứng dụng đã sẵn sàng. Hãy chọn file/thư mục để bắt đầu.\n")
        self.log_box.configure(state="disabled")

        # --- Khởi chạy luồng kiểm tra cập nhật ngầm ---
        threading.Thread(target=self.check_for_updates, daemon=True).start()

        # --- 6. FOOTER (CHỮ KÝ & PHIÊN BẢN) ---
        self.frame_footer = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_footer.pack(side="bottom", pady=(0, 10))

        today = datetime.now()
        build_str = today.strftime("%Y%m%d")
        year_str = today.strftime("%Y")

        lbl_dev_prefix = ctk.CTkLabel(self.frame_footer, text="Developed by ", text_color="gray50", font=("Arial", 12))
        lbl_dev_prefix.pack(side="left")

        lbl_link = ctk.CTkLabel(self.frame_footer, text="@tozn607", text_color="#1f6aa5", font=("Arial", 12, "bold", "underline"), cursor="hand2")
        lbl_link.pack(side="left")
        # Gắn sự kiện: Bấm chuột trái (<Button-1>) sẽ mở trình duyệt tới link GitHub
        lbl_link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/tozn607"))

        lbl_version_suffix = ctk.CTkLabel(self.frame_footer, text=f" | Version v{CURRENT_VERSION} Build {build_str} | © {year_str}", text_color="gray50", font=("Arial", 12))
        lbl_version_suffix.pack(side="left")

    # --- HÀM THAY ĐỔI GIAO DIỆN THEO CHẾ ĐỘ ---
    def change_mode(self, selected_mode):
        self.entry_input.delete(0, "end")
        if selected_mode == "Chế độ Đơn (1 File PDF)":
            self.btn_input.configure(text="Chọn File PDF")
            self.entry_input.configure(placeholder_text="Đường dẫn đến 1 file PDF...")
        else:
            self.btn_input.configure(text="Thư mục Input")
            self.entry_input.configure(placeholder_text="Thư mục chứa nhiều file PDF...")

    # --- HÀM ĐIỀU KHIỂN HỆ THỐNG VÀ API KEY ---
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
                self.write_log("[+] Đã tải API Key từ file.")

    # --- HÀM TƯƠNG TÁC NGƯỜI DÙNG ---
    def write_log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def browse_input(self):
        current_mode = self.mode_var.get()
        if current_mode == "Chế độ Đơn (1 File PDF)":
            path = filedialog.askopenfilename(title="Chọn 1 file PDF", filetypes=[("PDF", "*.pdf")])
        else:
            path = filedialog.askdirectory(title="Chọn thư mục chứa PDF")
            
        if path:
            self.entry_input.delete(0, "end")
            self.entry_input.insert(0, path)

    def browse_output(self):
        folder_path = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
        if folder_path:
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, folder_path)

    def stop_processing(self):
        self.write_log("\n[!] NHẬN LỆNH DỪNG... Đang hủy bỏ các tiến trình (Vui lòng đợi 1-2 giây).")
        self.stop_event.set()

    def start_processing(self):
        api_key = self.entry_api.get().strip()
        input_path = self.entry_input.get().strip()
        output_dir = self.entry_output.get().strip()
        current_mode = self.mode_var.get()

        if not api_key:
            self.write_log("[!] LỖI: Vui lòng nhập API Key.")
            return
        if not input_path or not os.path.exists(input_path):
            self.write_log(f"[!] LỖI: Đường dẫn Input không tồn tại.")
            return

        # TỰ ĐỘNG XỬ LÝ NẾU BỎ TRỐNG OUTPUT THƯ MỤC
        if not output_dir:
            # Kiểm tra xem Input là thư mục (Chế độ hàng loạt) hay file (Chế độ đơn)
            if os.path.isdir(input_path):
                output_dir = input_path
            else:
                output_dir = os.path.dirname(input_path) # Cắt lấy đường dẫn thư mục chứa file
                
            self.write_log(f"[*] Không chọn Output, tự động lưu cùng thư mục Input: {output_dir}")
        elif not os.path.exists(output_dir):
            self.write_log("[!] LỖI: Thư mục Output không tồn tại.")
            return

        self.save_api_key(api_key)

        self.btn_start.configure(state="disabled")
        self.btn_input.configure(state="disabled")
        self.btn_output.configure(state="disabled")
        self.mode_selector.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        
        self.stop_event.clear()
        genai.configure(api_key=api_key)

        # Chạy luồng xử lý chung cho cả 2 mode
        thread = threading.Thread(target=self.process_documents, args=(input_path, output_dir, current_mode))
        thread.daemon = True
        thread.start()

    # --- HÀM XỬ LÝ LÕI TỔNG HỢP ---
    def process_documents(self, input_path, output_dir, mode):
        # Lập danh sách các file cần chạy dựa vào Mode
        if mode == "Chế độ Đơn (1 File PDF)":
            pdf_files = [input_path]
        else:
            pdf_files = [os.path.join(input_path, f) for f in os.listdir(input_path) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            self.write_log("\n[!] Không tìm thấy file PDF nào để xử lý!")
            self.reset_ui()
            return

        total_files = len(pdf_files)
        self.write_log(f"\n[>>>] BẮT ĐẦU: Sẽ xử lý {total_files} file PDF.")
        model = genai.GenerativeModel('gemini-3.1-flash-lite-preview')

        # DYNAMIC PROMPT: Nối thêm lệnh giải bài tập nếu được tick
        active_prompt = prompt_template
        if self.solve_var.get():
            active_prompt += "\n\n6. YÊU CẦU ĐẶC BIỆT: Tài liệu này chứa các bài tập/câu hỏi. Bạn BẮT BUỘC phải đọc và TRẢ LỜI/GIẢI CHI TIẾT các bài tập đó. Hãy tạo một phần 'ĐÁP ÁN' riêng biệt và rõ ràng ở cuối tài liệu và viết đáp án ở đó."

        for file_idx, pdf_path in enumerate(pdf_files):
            if self.stop_event.is_set(): break

            pdf_filename = os.path.basename(pdf_path)
            base_name = os.path.splitext(pdf_filename)[0]
            output_docx_path = os.path.join(output_dir, f"{base_name}_hoanthien.docx")

            self.write_log(f"\n[FILE {file_idx + 1}/{total_files}] Đang tách: {pdf_filename}...")
            
            try: images = convert_from_path(pdf_path)
            except Exception as e:
                self.write_log(f"  [X] LỖI ĐỌC PDF: {e}")
                continue 

            if not images:
                continue

            # --- XỬ LÝ TRANG BÌA ---
            start_idx = 0
            if self.cover_var.get():
                # Tạo tên file ảnh bìa: "TenFilePDF_cover.jpg"
                cover_path = os.path.join(output_dir, f"{base_name}_cover.jpg")
                try:
                    images[0].save(cover_path, "JPEG")
                    self.write_log(f"  [+] Đã lưu ảnh bìa: {base_name}_cover.jpg")
                except Exception as e:
                    self.write_log(f"  [!] Lỗi khi lưu ảnh bìa: {e}")
                
                # Bắt đầu vòng lặp đọc chữ từ trang số 2 (index 1)
                start_idx = 1 

            full_markdown_content = ""
            max_retries = 5 

            # Thay vì dùng enumerate(images), dùng range để kiểm soát điểm bắt đầu
            for i in range(start_idx, len(images)):
                image = images[i]
                if self.stop_event.is_set():
                    self.write_log("  [-] Bỏ dở tài liệu này do lệnh Dừng.")
                    break

                self.write_log(f"    [>] Trang {i+1}/{len(images)}...")
                
                for attempt in range(max_retries):
                    if self.stop_event.is_set(): break
                    try:
                        response = model.generate_content([active_prompt, image], safety_settings=safety_config)
                        
                        try:
                            text_result = response.text
                        except ValueError:
                            self.write_log(f"      [!] LỖI BẢN QUYỀN: Google từ chối đọc trang {i+1}.")
                            text_result = f"> **[LỖI: TỪ CHỐI NHẬN DIỆN TRANG {i+1} DO VƯỚNG BẢN QUYỀN]**"
                            full_markdown_content += f"\n\n\n\n{text_result}\n\n"
                            break
                        
                        # [TRICK XỬ LÝ FOOTNOTE] 
                        # Biến [^1] thành [^p1_1], [^p2_1] để tránh các trang bị trùng ID footnote của nhau khi gộp lại.
                        # Pandoc sẽ tự động sắp xếp và đánh số lại từ 1, 2, 3... trong Word một cách hoàn hảo.
                        text_result = text_result.replace("[^", f"[^p{i}_")
                        
                        full_markdown_content += f"\n\n\n\n{text_result}\n\n"
                        
                        if i < len(images) - 1 and not self.stop_event.is_set():
                            time.sleep(3) 
                        break 
                        
                    except Exception as e:
                        self.write_log(f"      [!] Lỗi (Thử {attempt+1}/{max_retries}): {e}")
                        if attempt < max_retries - 1:
                            time.sleep(60) 
                        else:
                            self.write_log(f"      [X] Bỏ qua trang {i+1}.")

            if not self.stop_event.is_set():
                try:
                    pypandoc.convert_text(full_markdown_content, 'docx', format='md', outputfile=output_docx_path)
                    self.write_log(f"  [***] Đã lưu Word: {output_docx_path}")
                except:
                    md_path = output_docx_path.replace('.docx', '.md')
                    with open(md_path, 'w', encoding='utf-8') as f: f.write(full_markdown_content)
                    self.write_log(f"  [+] Đã lưu Markdown tại: {md_path}")

        if self.stop_event.is_set():
            self.write_log("\n[⏹] TIẾN TRÌNH ĐÃ BỊ HỦY BỞI NGƯỜI DÙNG.")
        else:
            self.write_log("\n[✓] ĐÃ HOÀN TẤT!")
        
        self.reset_ui()

    def reset_ui(self):
        self.btn_start.configure(state="normal")
        self.btn_input.configure(state="normal")
        self.btn_output.configure(state="normal")
        self.mode_selector.configure(state="normal")
        self.btn_stop.configure(state="disabled")

    # --- TÍNH NĂNG GỘP ẢNH ---
    def open_merge_popup(self):
        self.merge_window = ctk.CTkToplevel(self)
        self.merge_window.title("Tiện ích: Gộp ảnh thành PDF")
        self.merge_window.geometry("600x450")
        self.merge_window.attributes("-topmost", True)
        self.merge_window.grab_set()

        self.selected_images_list = []

        frame_controls = ctk.CTkFrame(self.merge_window, fg_color="transparent")
        frame_controls.pack(pady=10, padx=20, fill="x")

        # Nút Thêm
        btn_add = ctk.CTkButton(frame_controls, text="Thêm ảnh", width=100, command=self.add_images)
        btn_add.pack(side="left", padx=5)

        # Nút Đảo vị trí
        btn_up = ctk.CTkButton(frame_controls, text="▲ Lên", width=70, fg_color="#454545", hover_color="#5a5a5a", command=self.move_up)
        btn_up.pack(side="left", padx=5)
        
        btn_down = ctk.CTkButton(frame_controls, text="▼ Xuống", width=70, fg_color="#454545", hover_color="#5a5a5a", command=self.move_down)
        btn_down.pack(side="left", padx=5)

        # Nút Xóa
        btn_clear = ctk.CTkButton(frame_controls, text="Xóa hết", width=80, fg_color="#a83232", hover_color="#7a2121", command=self.clear_images)
        btn_clear.pack(side="left", padx=5)

        # Sử dụng Listbox của tkinter thay vì Textbox để có thể click chọn dòng
        list_frame = tk.Frame(self.merge_window, bg="#343638")
        list_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        # Style Listbox cho giống Dark Mode
        self.img_listbox = tk.Listbox(list_frame, font=("Consolas", 12), bg="#2b2b2b", fg="#ffffff", 
                                      selectbackground="#1f538d", highlightthickness=0, borderwidth=0,
                                      yscrollcommand=scrollbar.set)
        self.img_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.img_listbox.yview)

        btn_export = ctk.CTkButton(self.merge_window, text="XUẤT RA PDF", font=("Arial", 14, "bold"), fg_color="#2b7a4b", command=self.export_to_pdf)
        btn_export.pack(pady=10)

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
                self.img_listbox.select_set(idx-1) # Giữ trạng thái bôi đen
        except IndexError:
            pass # Chưa chọn dòng nào thì bỏ qua

    def move_down(self):
        try:
            idx = self.img_listbox.curselection()[0]
            if idx < len(self.selected_images_list) - 1:
                self.selected_images_list[idx+1], self.selected_images_list[idx] = self.selected_images_list[idx], self.selected_images_list[idx+1]
                self.update_image_listbox()
                self.img_listbox.select_set(idx+1) # Giữ trạng thái bôi đen
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
            try:
                # Pillow có thể mở HEIC nhờ pillow-heif đã register ở trên
                imgs = [Image.open(p).convert('RGB') for p in self.selected_images_list]
                imgs[0].save(save_path, save_all=True, append_images=imgs[1:])
                self.img_listbox.delete(0, 'end')
                self.img_listbox.insert('end', f"--- THÀNH CÔNG ---")
                self.img_listbox.insert('end', f"Đã lưu PDF: {save_path}")
            except Exception as e: 
                self.img_listbox.delete(0, 'end')
                self.img_listbox.insert('end', f"[!] LỖI: {e}")

    # --- TÍNH NĂNG KIỂM TRA CẬP NHẬT TỪ GITHUB ---
    def check_for_updates(self):
        """Hàm chạy ngầm kiểm tra version.txt từ GitHub"""
        try:
            req = urllib.request.Request(UPDATE_RAW_URL, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=5) as response:
                latest_version = response.read().decode('utf-8').strip()

            # Nếu phiên bản trên mạng lớn hơn phiên bản hiện tại
            if latest_version > CURRENT_VERSION:
                # Gọi hàm hiển thị Popup thông qua UI thread chính
                self.after(500, self.show_update_popup, latest_version)
        except Exception as e:
            print(f"[!] Không thể kiểm tra cập nhật: {e}")

    def show_update_popup(self, latest_version):
        """Hiển thị cửa sổ thông báo khi có bản mới"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Thông báo Cập nhật")
        dialog.geometry("400x200")
        dialog.attributes("-topmost", True)
        dialog.grab_set()
        
        lbl = ctk.CTkLabel(dialog, text=f"Có phiên bản mới: v{latest_version}\nPhiên bản hiện tại: v{CURRENT_VERSION}\n\nBạn có muốn tải bản cập nhật về không?", font=("Arial", 14))
        lbl.pack(pady=30)
        
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        # Nút Tải ngay sẽ mở trình duyệt và đóng popup
        btn_yes = ctk.CTkButton(btn_frame, text="Tải ngay", fg_color="#2b7a4b", hover_color="#1e5c37", 
                                command=lambda: [webbrowser.open(RELEASES_URL), dialog.destroy()])
        btn_yes.pack(side="left", padx=10)
        
        btn_no = ctk.CTkButton(btn_frame, text="Bỏ qua", fg_color="gray", hover_color="darkgray", command=dialog.destroy)
        btn_no.pack(side="left", padx=10)

if __name__ == "__main__":
    app = PDFOCRApp()
    app.mainloop()