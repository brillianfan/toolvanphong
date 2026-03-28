import os
import sys
import ctypes

import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image
try:
    import pillow_heif
    # Đăng ký opener để Pillow có thể đọc được file HEIC
    pillow_heif.register_heif_opener()
except ImportError:
    # pillow_heif không khả dụng, ứng dụng vẫn có thể chạy
    pass
from pdf2image import convert_from_path
from pypdf import PdfReader, PdfWriter

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller tạo một thư mục tạm và lưu đường dẫn trong _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Sử dụng đường dẫn của chính file script làm gốc thay vì working directory (.)
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

def open_folder(path):
    """Mở thư mục chứa file hoặc thư mục chỉ định"""
    if os.path.isfile(path):
        path = os.path.dirname(path)
    if sys.platform == 'win32':
        os.startfile(path)
    elif sys.platform == 'darwin':
        import subprocess
        subprocess.Popen(['open', path])
    else:
        import subprocess
        subprocess.Popen(['xdg-open', path])

class MergeDialog:
    def __init__(self, parent, files, callback):
        self.top = tk.Toplevel(parent)
        self.top.title("Xử lý gộp PDF - Sắp xếp thứ tự")
        self.top.geometry("600x450")
        self.top.configure(bg="white")
        self.top.grab_set()  # Make dialog modal
        
        self.files = list(files)
        self.callback = callback
        
        tk.Label(self.top, text="Kéo thả hoặc sử dụng nút để đổi thứ tự", font=("Arial", 11, "bold"), bg="white").pack(pady=10)
        
        frame = tk.Frame(self.top, bg="white")
        frame.pack(expand=True, fill="both", padx=20, pady=5)
        
        self.listbox = tk.Listbox(frame, selectmode=tk.SINGLE, font=("Arial", 10), height=15)
        self.listbox.pack(side="left", expand=True, fill="both")
        
        # Đăng ký sự kiện kéo thả để đổi thứ tự
        self.listbox.bind("<Button-1>", self.on_drag_start)
        self.listbox.bind("<B1-Motion>", self.on_drag_motion)
        self.listbox.bind("<ButtonRelease-1>", self.on_drag_drop)
        self.drag_index = None
        
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=scrollbar.set)
        
        self.update_listbox()
        
        btn_frame = tk.Frame(self.top, bg="white")
        btn_frame.pack(fill="x", pady=15, padx=20)
        
        tk.Button(btn_frame, text="↑ Lên", command=self.move_up, width=10).pack(side="left", padx=5)
        tk.Button(btn_frame, text="↓ Xuống", command=self.move_down, width=10).pack(side="left", padx=5)
        tk.Button(btn_frame, text="❌ Xóa", command=self.remove_item, fg="red", width=10).pack(side="left", padx=5)
        
        tk.Button(btn_frame, text="Tiến hành Gộp", command=self.on_merge, bg="#1a73e8", fg="white", font=("Arial", 10, "bold"), width=15).pack(side="right", padx=5)

    def on_drag_start(self, event):
        self.drag_index = self.listbox.nearest(event.y)

    def on_drag_motion(self, event):
        i = self.listbox.nearest(event.y)
        if i != self.drag_index:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(i)
            
        # Tự động cuộn khi kéo gần biên
        if event.y < 10:
            self.listbox.yview_scroll(-1, "units")
        elif event.y > self.listbox.winfo_height() - 10:
            self.listbox.yview_scroll(1, "units")

    def on_drag_drop(self, event):
        if self.drag_index is None:
            return
        
        drop_index = self.listbox.nearest(event.y)
        if drop_index != self.drag_index:
            # Di chuyển item trong danh sách files
            item = self.files.pop(self.drag_index)
            self.files.insert(drop_index, item)
            self.update_listbox()
            self.listbox.selection_set(drop_index)
        
        self.drag_index = None

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for f in self.files:
            self.listbox.insert(tk.END, os.path.basename(f))

    def move_up(self):
        idx = self.listbox.curselection()
        if not idx or idx[0] == 0:
            return
        
        i = idx[0]
        self.files[i], self.files[i-1] = self.files[i-1], self.files[i]
        self.update_listbox()
        self.listbox.selection_set(i-1)

    def move_down(self):
        idx = self.listbox.curselection()
        if not idx or idx[0] == len(self.files) - 1:
            return
        
        i = idx[0]
        self.files[i], self.files[i+1] = self.files[i+1], self.files[i]
        self.update_listbox()
        self.listbox.selection_set(i+1)

    def remove_item(self):
        idx = self.listbox.curselection()
        if not idx:
            return
        
        if len(self.files) <= 2:
            messagebox.showwarning("Cảnh báo", "Cần ít nhất 2 file để gộp!")
            return
            
        self.files.pop(idx[0])
        self.update_listbox()

    def on_merge(self):
        self.top.destroy()
        self.callback(self.files)

class ExtractDialog:
    def __init__(self, parent, input_path, num_pages, callback):
        self.top = tk.Toplevel(parent)
        self.top.title("Trích xuất trang PDF")
        self.top.geometry("400x250")
        self.top.configure(bg="white")
        self.top.resizable(False, False)
        self.top.grab_set()
        
        self.input_path = input_path
        self.num_pages = num_pages
        self.callback = callback
        
        tk.Label(self.top, text="Trích xuất trang PDF", font=("Arial", 12, "bold"), bg="white").pack(pady=10)
        tk.Label(self.top, text=f"File: {os.path.basename(input_path)}", font=("Arial", 9), bg="white", fg="#666").pack()
        tk.Label(self.top, text=f"Tổng số trang: {num_pages}", font=("Arial", 9, "bold"), bg="white", fg="#1a73e8").pack(pady=5)
        
        tk.Label(self.top, text="Nhập các trang cần lấy (VD: 1, 3, 5-8)", font=("Arial", 10), bg="white").pack(pady=(10, 0))
        
        self.entry = tk.Entry(self.top, font=("Arial", 11), width=30, justify="center")
        self.entry.pack(pady=10)
        self.entry.insert(0, f"1-{num_pages}")
        self.entry.focus_set()
        
        btn_frame = tk.Frame(self.top, bg="white")
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="Hủy", command=self.top.destroy, width=10).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Trích xuất", command=self.on_submit, bg="#1a73e8", fg="white", font=("Arial", 10, "bold"), width=15).pack(side="left", padx=10)

    def on_submit(self):
        pages_str = self.entry.get().strip()
        if not pages_str:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập số trang!")
            return
        self.top.destroy()
        self.callback(self.input_path, pages_str)

class DeletePagesDialog:
    def __init__(self, parent, input_path, num_pages, callback):
        self.top = tk.Toplevel(parent)
        self.top.title("Xóa trang PDF")
        self.top.geometry("400x250")
        self.top.configure(bg="white")
        self.top.resizable(False, False)
        self.top.grab_set()
        
        self.input_path = input_path
        self.num_pages = num_pages
        self.callback = callback
        
        tk.Label(self.top, text="Xóa trang PDF", font=("Arial", 12, "bold"), bg="white").pack(pady=10)
        tk.Label(self.top, text=f"File: {os.path.basename(input_path)}", font=("Arial", 9), bg="white", fg="#666").pack()
        tk.Label(self.top, text=f"Tổng số trang: {num_pages}", font=("Arial", 9, "bold"), bg="white", fg="#d93025").pack(pady=5)
        
        tk.Label(self.top, text="Nhập các trang cần XÓA (VD: 1, 3, 5-8)", font=("Arial", 10), bg="white").pack(pady=(10, 0))
        
        self.entry = tk.Entry(self.top, font=("Arial", 11), width=30, justify="center")
        self.entry.pack(pady=10)
        self.entry.focus_set()
        
        btn_frame = tk.Frame(self.top, bg="white")
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="Hủy", command=self.top.destroy, width=10).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Xóa trang", command=self.on_submit, bg="#d93025", fg="white", font=("Arial", 10, "bold"), width=15).pack(side="left", padx=10)

    def on_submit(self):
        pages_str = self.entry.get().strip()
        if not pages_str:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập số trang!")
            return
        self.top.destroy()
        self.callback(self.input_path, pages_str)

class ImageFormatDialog:
    def __init__(self, parent, callback):
        self.top = tk.Toplevel(parent)
        self.top.title("Chọn định dạng xuất")
        self.top.geometry("350x200")
        self.top.configure(bg="white")
        self.top.resizable(False, False)
        self.top.grab_set()
        
        self.callback = callback
        
        tk.Label(self.top, text="Chọn định dạng ảnh đầu ra", font=("Arial", 12, "bold"), bg="white").pack(pady=20)
        
        btn_frame = tk.Frame(self.top, bg="white")
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="JPEG (JPG)", command=lambda: self.on_select("JPEG"), 
                  bg="#1a73e8", fg="white", font=("Arial", 10, "bold"), width=12, height=2).pack(side="left", padx=15)
        tk.Button(btn_frame, text="PNG", command=lambda: self.on_select("PNG"), 
                  bg="#34a853", fg="white", font=("Arial", 10, "bold"), width=12, height=2).pack(side="left", padx=15)
        
        tk.Button(self.top, text="Hủy", command=self.top.destroy, width=10).pack(pady=15)

    def on_select(self, fmt):
        self.top.destroy()
        self.callback(fmt)

class ProToolboxApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tool Van Phong - v.2.0 - Brillian Pham")
        self.root.geometry("900x600")
        self.root.configure(bg="white")
        
        # Thiết lập Icon cho ứng dụng
        try:
            icon_path = resource_path("app_icon.ico")
            if os.path.exists(icon_path):
                # 1. Icon cho thanh tiêu đề (Thường là 16x16 hoặc 32x32)
                self.root.iconbitmap(icon_path)
                
                # 2. Icon cho Taskbar và Alt+Tab (Hỗ trợ đa độ phân giải)
                icon_image = Image.open(icon_path)
                from PIL import ImageTk
                self.icon_photo = ImageTk.PhotoImage(icon_image)
                self.root.wm_iconphoto(True, self.icon_photo)
            
            # Đặt AppUserModelID để Windows hiển thị icon trên taskbar đúng cách
            myappid = 'mycompany.myproduct.subproduct.version' # arbitrary string
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception as e:
            print(f"Không thể tải icon: {e}")
        
        self.quality = tk.IntVar(value=85)
        self.quality.trace_add("write", self.update_quality_label)
        
        self.dpi = tk.IntVar(value=300)
        self.dpi.trace_add("write", self.update_dpi_label)

        self.setup_menu()
        self.setup_header()
        self.setup_grid_ui()
        
        self.status = tk.Label(self.root, text="Sẵn sàng!", bg="white", fg="green", font=("Arial", 10))
        self.status.pack(side="bottom", pady=15)

    def reset_tool(self):
        """Đưa tool về trạng thái mặc định"""
        self.quality.set(85)
        self.dpi.set(300)
        self.status.config(text="Đã làm mới! Sẵn sàng!", fg="green")

    def setup_menu(self):
        menubar = tk.Menu(self.root)
        
        # Menu Chất lượng
        q_menu = tk.Menu(menubar, tearoff=0)
        for q in [100, 90, 80, 70, 60, 50, 40, 30, 20, 10]:
            q_menu.add_radiobutton(label=f"Chất lượng: {q}%", variable=self.quality, value=q)
        menubar.add_cascade(label="Chất lượng nén", menu=q_menu)
        
        # Menu DPI
        dpi_menu = tk.Menu(menubar, tearoff=0)
        for d in [600, 300, 150, 96, 72]:
            dpi_menu.add_radiobutton(label=f"DPI: {d}", variable=self.dpi, value=d)
        menubar.add_cascade(label="Độ phân giải (DPI)", menu=dpi_menu)

        about_menu = tk.Menu(menubar, tearoff=0)
        about_menu.add_command(label="Thông tin", command=lambda: messagebox.showinfo("Tác giả", "Tool by Brillian Pham\nPhiên bản 2.0"))
        menubar.add_cascade(label="Hỗ trợ", menu=about_menu)
        self.root.config(menu=menubar)

    def setup_header(self):
        header_frame = tk.Frame(self.root, bg="white")
        header_frame.pack(fill="x", pady=20)
        tk.Label(header_frame, text="Tool Văn Phòng", font=("Arial", 26, "bold"), bg="white", fg="#333").pack()
        tk.Label(header_frame, text="Chọn chất lượng nén, DPI và kéo thả tệp để xử lý", font=("Arial", 11), bg="white", fg="#666").pack(pady=5)
        
        settings_frame = tk.Frame(header_frame, bg="white")
        settings_frame.pack(pady=10)

        self.quality_label = tk.Label(settings_frame, text=f"Chất lượng nén: {self.quality.get()}%", 
                                     font=("Arial", 12, "bold"), bg="#e8f0fe", fg="#1a73e8", 
                                     padx=15, pady=5)
        self.quality_label.pack(side="left", padx=10)

        self.dpi_label = tk.Label(settings_frame, text=f"DPI: {self.dpi.get()}", 
                                  font=("Arial", 12, "bold"), bg="#f1f3f5", fg="#495057", 
                                  padx=15, pady=5)
        self.dpi_label.pack(side="left", padx=10)

        # Nút Refresh
        self.refresh_btn = tk.Button(header_frame, text="🔄 Làm mới (Refresh)", 
                                     command=self.reset_tool,
                                     bg="#f8f9fa", fg="#333", font=("Arial", 10),
                                     relief="flat", highlightthickness=1,
                                     padx=10, pady=2)
        self.refresh_btn.pack(pady=5)
        self.refresh_btn.bind("<Enter>", lambda e: self.refresh_btn.config(bg="#e9ecef"))
        self.refresh_btn.bind("<Leave>", lambda e: self.refresh_btn.config(bg="#f8f9fa"))

    def update_quality_label(self, *args):
        if hasattr(self, 'quality_label'):
            self.quality_label.config(text=f"Chất lượng nén: {self.quality.get()}%")

    def update_dpi_label(self, *args):
        if hasattr(self, 'dpi_label'):
            self.dpi_label.config(text=f"DPI: {self.dpi.get()}")

    def create_tool_card(self, parent, title, action_type, row, col, icon="📄"):
        card = tk.Frame(parent, bg="#f8f9fa", highlightbackground="#e0e0e0", highlightthickness=1, cursor="hand2")
        card.grid(row=row, column=col, padx=12, pady=12, sticky="nsew")
        
        tk.Label(card, text=icon, font=("Arial", 24), bg="#f8f9fa").pack(pady=(20, 0))
        tk.Label(card, text=title, font=("Arial", 10, "bold"), bg="#f8f9fa", wraplength=130, justify="center").pack(pady=20, padx=10)

        card.drop_target_register(DND_FILES)
        card.dnd_bind('<<Drop>>', lambda e, act=action_type: self.handle_drop(e, act))
        
        card.bind("<Enter>", lambda e: card.config(bg="#f1f3f5", highlightbackground="#1a73e8"))
        card.bind("<Leave>", lambda e: card.config(bg="#f8f9fa", highlightbackground="#e0e0e0"))

    def setup_grid_ui(self):
        container = tk.Frame(self.root, bg="white")
        container.pack(expand=True, fill="both", padx=50)

        for i in range(4): container.grid_columnconfigure(i, weight=1)
        for i in range(2): container.grid_rowconfigure(i, weight=1)

        # Danh sách các công cụ (Title, Action_ID, Row, Col, Icon)
        tools = [
            ("Nén và chuyển sang PDF", "TO_PDF", 0, 0, "📑"),
            ("PDF sang Ảnh (JPG)", "PDF_TO_IMG", 0, 1, "🖼️"),
            ("Trích xuất trang PDF", "EXTRACT_PDF", 0, 2, "📄"),
            ("Nén tệp PDF", "COMPRESS_PDF", 0, 3, "🗜️"),
            ("Chuyển sang Ảnh (JPG/PNG)", "CONVERT_IMG", 1, 0, "🖼️"),
            ("Gộp nhiều PDF", "MERGE_PDF", 1, 1, "➕"),
            ("Tách 1 PDF thành nhiều trang", "SPLIT_PDF", 1, 2, "✂️"),
            ("Xóa trang PDF", "DELETE_PDF_PAGES", 1, 3, "🗑️"),
        ]

        for title, act, r, c, ico in tools:
            self.create_tool_card(container, title, act, r, c, ico)

    def get_unique_path(self, folder, base_name, ext):
        """Tạo đường dẫn file duy nhất để tránh ghi đè"""
        output_path = os.path.join(folder, f"{base_name}.{ext}")
        counter = 1
        while os.path.exists(output_path):
            output_path = os.path.join(folder, f"{base_name}_{counter}.{ext}")
            counter += 1
        return output_path

    def handle_drop(self, event, action):
        files = self.root.tk.splitlist(event.data)
        
        # Chỉ lấy các đường dẫn là file thực sự (loại bỏ thư mục)
        valid_files = [f for f in files if os.path.isfile(f)]
        
        if not valid_files:
            self.status.config(text="⚠ Lỗi: Không tìm thấy file hợp lệ (vui lòng không kéo thư mục)!", fg="orange")
            return

        if action == "MERGE_PDF":
            # Lọc chỉ lấy file PDF
            pdf_files = [f for f in valid_files if f.lower().endswith(".pdf")]
            if not pdf_files:
                self.status.config(text="⚠ Lỗi: Vui lòng chọn các file PDF để gộp!", fg="red")
                return
            self.process_pdf_merge(pdf_files)
            return

        for f in valid_files:
            f_ext = os.path.splitext(f)[1].lower()
            
            if action == "PDF_TO_IMG":
                if f_ext == ".pdf":
                    self.process_pdf_to_img(f)
                else:
                    self.status.config(text=f"⚠ Bỏ qua: {os.path.basename(f)} không phải PDF", fg="orange")
            elif action == "SPLIT_PDF":
                if f_ext == ".pdf":
                    self.process_pdf_split(f)
                else:
                    self.status.config(text=f"⚠ Bỏ qua: {os.path.basename(f)} không phải PDF", fg="orange")
            elif action == "COMPRESS_PDF":
                if f_ext == ".pdf":
                    self.process_pdf_compress(f)
                else:
                    self.status.config(text=f"⚠ Bỏ qua: {os.path.basename(f)} không phải PDF", fg="orange")
            elif action == "EXTRACT_PDF":
                if f_ext == ".pdf":
                    self.process_pdf_extract(f)
                else:
                    self.status.config(text=f"⚠ Bỏ qua: {os.path.basename(f)} không phải PDF", fg="orange")
            elif action == "DELETE_PDF_PAGES":
                if f_ext == ".pdf":
                    self.process_pdf_delete(f)
                else:
                    self.status.config(text=f"⚠ Bỏ qua: {os.path.basename(f)} không phải PDF", fg="orange")
            elif action == "CONVERT_IMG":
                # Lọc các file ảnh hỗ trợ bao gồm mới: jfif, heic
                img_files = [f for f in valid_files if os.path.splitext(f)[1].lower() in [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".jfif", ".heic"]]
                if not img_files:
                    self.status.config(text="⚠ Lỗi: Vui lòng chọn các file ảnh hỗ trợ!", fg="red")
                    return
                # Nếu chỉ có 1 file, có thể vẫn dùng list để đồng nhất logic callback
                ImageFormatDialog(self.root, lambda fmt: [self.process_image_convert(img_f, fmt) for img_f in img_files])
                return
            else:
                # Các action chuyển đổi ảnh: TO_PDF (có thể thêm các định dạng khác sau này)
                if f_ext in [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".jfif", ".heic"]:
                    self.process_image_convert(f, action)
                else:
                    self.status.config(text=f"⚠ Bỏ qua: {os.path.basename(f)} không phải định dạng ảnh hỗ trợ", fg="orange")

    def process_pdf_to_img(self, input_path):
        """Chức năng mới: Chuyển PDF thành các file ảnh"""
        if not input_path.lower().endswith(".pdf"):
            self.status.config(text="Lỗi: Vui lòng kéo file PDF vào ô này!", fg="red")
            return
        
        try:
            self.status.config(text="🔄 Đang xử lý PDF sang ảnh...", fg="blue")
            self.root.update_idletasks()
            
            # Cố gắng xác định đường dẫn bin của Poppler một cách linh hoạt nhất
            # 1. Thử đường dẫn tương đối so với script (cho dev)
            script_dir = os.path.dirname(os.path.abspath(__file__))
            pp_path = os.path.join(script_dir, 'poppler-windows', 'Library', 'bin')
            
            # 2. Thử qua resource_path (cho PyInstaller)
            if not os.path.exists(pp_path):
                pp_path = resource_path(os.path.join('poppler-windows', 'Library', 'bin'))
            
            pp_path = os.path.abspath(pp_path)
            
            if not os.path.exists(pp_path):
                raise Exception(f"Không thể định vị thư mục Poppler/bin tại: {pp_path}")

            # Thiết lập PATH để subprocess có thể tìm thấy các DLL phụ thuộc
            original_path = os.environ.get("PATH", "")
            if pp_path not in original_path:
                os.environ["PATH"] = pp_path + os.pathsep + original_path
            
            # Thực hiện chuyển đổi
            pages = convert_from_path(input_path, dpi=self.dpi.get(), poppler_path=pp_path)
            
            folder = os.path.dirname(input_path)
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            
            num_pages = len(pages)
            padding = len(str(num_pages))
            
            q = self.quality.get()
            d = self.dpi.get()
            for i, page in enumerate(pages):
                page_num = str(i + 1).zfill(padding)
                output = self.get_unique_path(folder, f"{base_name}_page_{page_num}", "jpg")
                page.save(output, "JPEG", quality=q, dpi=(d, d))
            
            self.status.config(text=f"✅ Thành công! Đã lưu {len(pages)} ảnh vào: {os.path.basename(folder)}", fg="green")
            open_folder(folder)
        except Exception as e:
            self.status.config(text=f"❌ Lỗi PDF sang Ảnh: {str(e)}", fg="red")
            print(f"DEBUG: {str(e)}")

    def process_pdf_merge(self, files):
        """Gộp nhiều file PDF thành 1 (có bước xem lại thứ tự)"""
        pdf_files = [f for f in files if f.lower().endswith(".pdf")]
        if len(pdf_files) < 2:
            self.status.config(text="Lỗi: Cần ít nhất 2 file PDF để gộp!", fg="red")
            return
        
        # Mở dialog cho phép sắp xếp
        MergeDialog(self.root, pdf_files, self.execute_pdf_merge)

    def execute_pdf_merge(self, pdf_files):
        """Thực hiện gộp PDF sau khi đã xác nhận thứ tự"""
        try:
            writer = PdfWriter()
            for pdf in pdf_files:
                reader = PdfReader(pdf)
                for page in reader.pages:
                    writer.add_page(page)
            
            output_path = self.get_unique_path(os.path.dirname(pdf_files[0]), "Merged_Output", "pdf")
                
            with open(output_path, "wb") as f:
                writer.write(f)
            
            self.status.config(text=f"✅ Đã gộp {len(pdf_files)} file thành: {os.path.basename(output_path)}", fg="green")
            open_folder(output_path)
        except Exception as e:
            self.status.config(text=f"Lỗi Gộp PDF: {str(e)}", fg="red")

    def process_pdf_extract(self, input_path):
        """Mở dialog để chọn trang cần trích xuất"""
        try:
            reader = PdfReader(input_path)
            num_pages = len(reader.pages)
            ExtractDialog(self.root, input_path, num_pages, self.execute_pdf_extract)
        except Exception as e:
            self.status.config(text=f"Lỗi đọc PDF: {str(e)}", fg="red")

    def execute_pdf_extract(self, input_path, pages_str):
        """Thực hiện trích xuất các trang đã chọn"""
        try:
            reader = PdfReader(input_path)
            total_pages = len(reader.pages)
            writer = PdfWriter()
            
            # Phân tích chuỗi trang (vd: 1, 3, 5-8)
            selected_pages = set()
            parts = pages_str.replace(" ", "").split(",")
            for part in parts:
                if "-" in part:
                    start, end = map(int, part.split("-"))
                    selected_pages.update(range(start, end + 1))
                else:
                    selected_pages.add(int(part))
            
            # Lọc các trang hợp lệ và chuyển về 0-index
            valid_pages = sorted([p-1 for p in selected_pages if 1 <= p <= total_pages])
            
            if not valid_pages:
                messagebox.showerror("Lỗi", "Không có trang nào hợp lệ để trích xuất!")
                return
            
            for p_idx in valid_pages:
                writer.add_page(reader.pages[p_idx])
            
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = self.get_unique_path(os.path.dirname(input_path), f"{base_name}_extracted", "pdf")
            
            with open(output_path, "wb") as f:
                writer.write(f)
                
            self.status.config(text=f"✅ Đã trích {len(valid_pages)} trang thành: {os.path.basename(output_path)}", fg="green")
            open_folder(output_path)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể trích xuất trang: {str(e)}")
            self.status.config(text=f"Lỗi Trích xuất PDF: {str(e)}", fg="red")

    def process_pdf_split(self, input_path):
        """Tách 1 file PDF thành nhiều file PDF từng trang"""
        if not input_path.lower().endswith(".pdf"):
            self.status.config(text="Lỗi: Vui lòng chọn file PDF!", fg="red")
            return
            
        try:
            reader = PdfReader(input_path)
            num_pages = len(reader.pages)
            
            # Kiểm tra nếu file chỉ có 1 trang
            if num_pages <= 1:
                self.status.config(text="⚠ File chỉ có 1 trang, không thể tách!", fg="orange")
                return

            base_name = os.path.splitext(os.path.basename(input_path))[0]
            parent_folder = os.path.dirname(input_path)
            
            # Tạo thư mục con để chứa các trang đã tách
            output_folder = self.get_unique_path(parent_folder, f"{base_name}_Split", "")
            if output_folder.endswith("."): output_folder = output_folder[:-1] # Remove trailing dot if any from get_unique_path
            os.makedirs(output_folder, exist_ok=True)
            
            padding = len(str(num_pages))
            
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                page_num = str(i + 1).zfill(padding)
                output_file = os.path.join(output_folder, f"{base_name}_page_{page_num}.pdf")
                with open(output_file, "wb") as f:
                    writer.write(f)
                    
            self.status.config(text=f"✅ Đã tách {num_pages} trang vào folder: {os.path.basename(output_folder)}", fg="green")
            open_folder(output_folder)
        except Exception as e:
            self.status.config(text=f"Lỗi Tách PDF: {str(e)}", fg="red")

    def process_pdf_delete(self, input_path):
        """Mở dialog để chọn trang cần xóa"""
        try:
            reader = PdfReader(input_path)
            num_pages = len(reader.pages)
            DeletePagesDialog(self.root, input_path, num_pages, self.execute_pdf_delete)
        except Exception as e:
            self.status.config(text=f"Lỗi đọc PDF: {str(e)}", fg="red")

    def execute_pdf_delete(self, input_path, pages_str):
        """Thực hiện xóa các trang đã chọn"""
        try:
            reader = PdfReader(input_path)
            total_pages = len(reader.pages)
            writer = PdfWriter()
            
            # Phân tích chuỗi trang (vd: 1, 3, 5-8)
            removed_pages = set()
            parts = pages_str.replace(" ", "").split(",")
            for part in parts:
                if "-" in part:
                    try:
                        start, end = map(int, part.split("-"))
                        removed_pages.update(range(start, end + 1))
                    except: continue
                else:
                    try:
                        removed_pages.add(int(part))
                    except: continue
            
            # Các trang giữ lại (chuyển về 0-index)
            pages_to_keep = [i for i in range(total_pages) if (i + 1) not in removed_pages]
            
            if not pages_to_keep:
                messagebox.showerror("Lỗi", "Bạn không thể xóa tất cả các trang!")
                return
            
            for p_idx in pages_to_keep:
                writer.add_page(reader.pages[p_idx])
            
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = self.get_unique_path(os.path.dirname(input_path), f"{base_name}_removed_pages", "pdf")
            
            with open(output_path, "wb") as f:
                writer.write(f)
                
            self.status.config(text=f"✅ Đã xóa {total_pages - len(pages_to_keep)} trang. Còn lại {len(pages_to_keep)} trang.", fg="green")
            open_folder(output_path)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xóa trang: {str(e)}")
            self.status.config(text=f"Lỗi Xóa trang PDF: {str(e)}", fg="red")

    def process_pdf_compress(self, input_path):
        """Nén file PDF bằng cách chuyển sang ảnh và gộp lại (Chất lượng gốc, nén theo chất lượng người dùng chọn)"""
        if not input_path.lower().endswith(".pdf"):
            self.status.config(text="Lỗi: Vui lòng chọn file PDF!", fg="red")
            return
            
        try:
            self.status.config(text="🔄 Đang nén PDF (Chuyển sang ảnh & nén)...", fg="blue")
            self.root.update_idletasks()
            
            # Định vị Poppler
            script_dir = os.path.dirname(os.path.abspath(__file__))
            pp_path = os.path.join(script_dir, 'poppler-windows', 'Library', 'bin')
            if not os.path.exists(pp_path):
                pp_path = resource_path(os.path.join('poppler-windows', 'Library', 'bin'))
            pp_path = os.path.abspath(pp_path)
            
            if not os.path.exists(pp_path):
                raise Exception("Không tìm thấy bộ công cụ Poppler để xử lý PDF.")

            # Thiết lập PATH
            original_path = os.environ.get("PATH", "")
            if pp_path not in original_path:
                os.environ["PATH"] = pp_path + os.pathsep + original_path

            q = self.quality.get()
            d = self.dpi.get()
            
            # 1. Chuyển PDF sang ảnh
            pages = convert_from_path(input_path, dpi=d, poppler_path=pp_path)
            
            if not pages:
                raise Exception("Không thể đọc được các trang từ file PDF.")

            # 2. Chuẩn bị danh sách ảnh
            output_pages = []
            for page in pages:
                if page.mode != "RGB":
                    page = page.convert("RGB")
                output_pages.append(page)
            
            # 3. Lưu thành 1 file PDF duy nhất
            base_name_str = os.path.splitext(os.path.basename(input_path))[0]
            output_path = self.get_unique_path(os.path.dirname(input_path), f"{base_name_str}_compressed", "pdf")
            
            output_pages[0].save(
                output_path, 
                "PDF", 
                save_all=True, 
                append_images=output_pages[1:], 
                quality=q,
                optimize=True,
                resolution=float(d)
            )
                
            orig_size = os.path.getsize(input_path)
            new_size = os.path.getsize(output_path)
            reduction = (orig_size - new_size) / orig_size * 100
            
            self.status.config(text=f"✅ Đã nén {reduction:.1f}%: {os.path.basename(output_path)}", fg="green")
            open_folder(output_path)
            
        except Exception as e:
            self.status.config(text=f"❌ Lỗi Nén PDF: {str(e)}", fg="red")
            print(f"DEBUG Nén PDF: {str(e)}")

    def process_image_convert(self, input_path, fmt):
        """Xử lý nén và chuyển đổi ảnh (giữ nguyên logic cũ)"""
        try:
            folder = os.path.dirname(input_path)
            name = os.path.splitext(os.path.basename(input_path))[0]
            ext = "pdf" if fmt == "TO_PDF" else fmt.lower()
            
            output_path = self.get_unique_path(folder, f"{name}_converted", ext)

            with Image.open(input_path) as img:
                # Chuyển đổi sang RGB nếu cần (cho PDF và JPEG không hỗ trợ Alpha)
                if (fmt in ["TO_PDF", "JPEG"]) and img.mode in ("RGBA", "LA", "P"):
                    img = img.convert("RGB")
                
                q = self.quality.get()
                d = self.dpi.get()
                if fmt == "TO_PDF":
                    img.save(output_path, "PDF", resolution=float(d), quality=q)
                elif fmt == "PNG":
                    # PNG sử dụng compress_level (0-9) thay vì quality
                    cl = max(0, min(9, (100-q)//10))
                    img.save(output_path, "PNG", compress_level=cl, dpi=(d, d))
                else: # JPEG
                    img.save(output_path, "JPEG", optimize=True, quality=q, dpi=(d, d))
            
            self.status.config(text=f"✅ Đã lưu: {os.path.basename(output_path)}", fg="green")
            open_folder(output_path)
        except Exception as e:
            self.status.config(text=f"❌ Lỗi xử lý ảnh: {str(e)}", fg="red")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ProToolboxApp(root)
    root.mainloop()
