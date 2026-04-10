import re
import os

with open("main.py", "r", encoding="utf-8") as f:
    orig = f.read()

constants_end = orig.find("class LanguageSelectorPopup")
constants_part = orig[:constants_end]

# remove customtkinter and darkdetect imports if any
constants_part = re.sub(r'import customtkinter as ctk\n', '', constants_part)
constants_part = re.sub(r'ctk\.set_appearance_mode\(.*?\)\n', '', constants_part)
constants_part = re.sub(r'ctk\.set_default_color_theme\(.*?\)\n', '', constants_part)
constants_part = re.sub(r'from tkinter import filedialog\n', '', constants_part)
constants_part = re.sub(r'import tkinter as tk\n', '', constants_part)

pyqt_imports = """
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QCheckBox, QComboBox, QTextEdit, QFileDialog, 
                             QMessageBox, QDialog, QButtonGroup, QListWidget, QFrame,
                             QSizePolicy, QScrollArea, QAbstractItemView)
from PyQt6.QtCore import Qt, pyqtSignal, QThread, QTimer, QSize
from PyQt6.QtGui import QFont, QIcon, QColor, QPalette, QCursor

QSS = \"\"\"
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
\"\"\"
"""

pyqt_classes = """

class LanguageSelectorPopup(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Language / Chọn ngôn ngữ")
        self.setFixedSize(450, 300)
        self.selected_lang = "EN"
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        
        lbl = QLabel("Welcome! Please choose your preferred language.\\nChào mừng! Vui lòng chọn ngôn ngữ của bạn.")
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
        self.group.addWidget(self.btn_en)
        self.group.addWidget(self.btn_vn)
        
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
                            full_markdown_content += f"\\n\\n\\n\\n{text_result}\\n\\n"
                            break
                        
                        text_result = text_result.replace("[^", f"[^p{i}_")
                        full_markdown_content += f"\\n\\n\\n\\n{text_result}\\n\\n"
                        
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
                            if '\\t' in paragraph.text:
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
        self.mode_group.addWidget(self.btn_mode_single)
        self.mode_group.addWidget(self.btn_mode_batch)
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
                self.write_log(f"\\n[*] Cập nhật mới có sẵn: v{latest} - Vui lòng tải tại Github.")
        except: pass

    def open_merge_popup(self):
        self.merge_window = MergeWindow(self)
        self.merge_window.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFOCRApp()
    window.show()
    sys.exit(app.exec())
"""

with open("rewrite_main.py", "w", encoding="utf-8") as f:
    f.write(constants_part + pyqt_imports + pyqt_classes)

print("Generated rewrite_main.py")
