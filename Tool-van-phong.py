import os
import sys
import ctypes
import threading

import tkinter as tk
from tkinter import messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image
import pillow_heif
# Đăng ký opener để Pillow có thể đọc được file HEIC
pillow_heif.register_heif_opener()
from pdf2image import convert_from_path
from pypdf import PdfReader, PdfWriter
import comtypes.client

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
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

class ProgressDialog:
    def __init__(self, parent, title="Đang xử lý"):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("350x120")
        self.top.configure(bg="white")
        self.top.resizable(False, False)
        self.top.grab_set()
        
        # Center in parent
        x = parent.winfo_x() + (parent.winfo_width() - 350) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 120) // 2
        self.top.geometry(f"+{x}+{y}")
        
        self.label = tk.Label(self.top, text="Vui lòng đợi trong giây lát...", font=("Arial", 10), bg="white")
        self.label.pack(pady=15)
        
        self.progress = ttk.Progressbar(self.top, mode="indeterminate", length=250)
        self.progress.pack(pady=5)
        self.progress.start(10)
        
    def set_text(self, text):
        self.label.config(text=text)
        
    def close(self):
        self.top.grab_release()
        self.top.destroy()

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

class ChangeDpiDialog:
    def __init__(self, parent, callback):
        self.top = tk.Toplevel(parent)
        self.top.title("Thay đổi DPI")
        self.top.geometry("350x200")
        self.top.configure(bg="white")
        self.top.resizable(False, False)
        self.top.grab_set()
        
        self.callback = callback
        
        tk.Label(self.top, text="Chọn DPI mục tiêu", font=("Arial", 12, "bold"), bg="white").pack(pady=20)
        
        self.dpi_var = tk.IntVar(value=300)
        dpi_frame = tk.Frame(self.top, bg="white")
        dpi_frame.pack(pady=5)
        
        for d in [96, 150, 300, 600]:
            tk.Radiobutton(dpi_frame, text=str(d), variable=self.dpi_var, value=d, bg="white", font=("Arial", 10)).pack(side="left", padx=10)
            
        btn_frame = tk.Frame(self.top, bg="white")
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="Hủy", command=self.top.destroy, width=10).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Xác nhận", command=self.on_submit, bg="#1a73e8", fg="white", font=("Arial", 10, "bold"), width=10).pack(side="left", padx=10)

    def on_submit(self):
        dpi = self.dpi_var.get()
        self.top.destroy()
        self.callback(dpi)

class ProToolboxApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tool Văn Phòng Pro - v.3.0 - Brillian Pham")
        self.root.geometry("1100x700")
        self.root.configure(bg="white")
        
        # Thiết lập Icon cho ứng dụng
        try:
            icon_path = resource_path("app_icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                icon_image = Image.open(icon_path)
                from PIL import ImageTk
                self.icon_photo = ImageTk.PhotoImage(icon_image)
                self.root.wm_iconphoto(True, self.icon_photo)
            
            myappid = 'mycompany.myproduct.subproduct.version'
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
        
        self.status = tk.Label(self.root, text="Sẵn sàng! Kéo thả tệp vào ô tương ứng để xử lý.", bg="white", fg="#4b5563", font=("Arial", 11, "bold"))
        self.status.pack(side="bottom", pady=15)

    def update_status(self, text, type_status="info"):
        colors = {
            "info": "#1e40af",    # Blue
            "success": "#166534", # Green
            "error": "#991b1b",   # Red
            "warning": "#854d0e"  # Yellow/Orange
        }
        self.status.config(text=text, fg=colors.get(type_status, "black"))

    def run_in_thread(self, target, *args, **kwargs):
        t = threading.Thread(target=target, args=args, kwargs=kwargs)
        t.daemon = True
        t.start()

    def reset_tool(self):
        self.quality.set(85)
        self.dpi.set(300)
        self.update_status("Đã làm mới! Sẵn sàng!", "success")

    def setup_menu(self):
        menubar = tk.Menu(self.root)
        
        q_menu = tk.Menu(menubar, tearoff=0)
        for q in [100, 90, 80, 70, 60, 50, 40, 30, 20, 10]:
            q_menu.add_radiobutton(label=f"Chất lượng: {q}%", variable=self.quality, value=q)
        menubar.add_cascade(label="Chất lượng nén", menu=q_menu)
        
        dpi_menu = tk.Menu(menubar, tearoff=0)
        for d in [600, 300, 150, 96, 72]:
            dpi_menu.add_radiobutton(label=f"DPI: {d}", variable=self.dpi, value=d)
        menubar.add_cascade(label="Độ phân giải (DPI)", menu=dpi_menu)

        about_menu = tk.Menu(menubar, tearoff=0)
        about_menu.add_command(label="Thông tin", command=lambda: messagebox.showinfo("Tác giả", "Tool by Brillian Pham\nPhiên bản 3.0"))
        menubar.add_cascade(label="Hỗ trợ", menu=about_menu)
        self.root.config(menu=menubar)

    def setup_header(self):
        header_frame = tk.Frame(self.root, bg="white")
        header_frame.pack(fill="x", pady=15)
        tk.Label(header_frame, text="Tool Văn Phòng Pro", font=("Arial", 28, "bold"), bg="white", fg="#1f2937").pack()
        tk.Label(header_frame, text="Chọn cấu hình chung và kéo thả tệp vào công cụ bên dưới để bắt đầu", font=("Arial", 11), bg="white", fg="#4b5563").pack(pady=3)
        
        settings_frame = tk.Frame(header_frame, bg="white")
        settings_frame.pack(pady=8)

        self.quality_label = tk.Label(settings_frame, text=f"Chất lượng nén: {self.quality.get()}%", 
                                     font=("Arial", 11, "bold"), bg="#eff6ff", fg="#1d4ed8", 
                                     padx=15, pady=5)
        self.quality_label.pack(side="left", padx=10)

        self.dpi_label = tk.Label(settings_frame, text=f"DPI: {self.dpi.get()}", 
                                  font=("Arial", 11, "bold"), bg="#f3f4f6", fg="#374151", 
                                  padx=15, pady=5)
        self.dpi_label.pack(side="left", padx=10)

        self.refresh_btn = tk.Button(header_frame, text="🔄 Làm mới (Refresh)", 
                                     command=self.reset_tool,
                                     bg="#f9fafb", fg="#374151", font=("Arial", 9, "bold"),
                                     relief="flat", highlightthickness=1,
                                     padx=12, pady=3)
        self.refresh_btn.pack(pady=3)
        self.refresh_btn.bind("<Enter>", lambda e: self.refresh_btn.config(bg="#f3f4f6"))
        self.refresh_btn.bind("<Leave>", lambda e: self.refresh_btn.config(bg="#f9fafb"))

    def update_quality_label(self, *args):
        if hasattr(self, 'quality_label'):
            self.quality_label.config(text=f"Chất lượng nén: {self.quality.get()}%")

    def update_dpi_label(self, *args):
        if hasattr(self, 'dpi_label'):
            self.dpi_label.config(text=f"DPI: {self.dpi.get()}")

    def create_tool_card(self, parent, title, description, action_type, row, col, icon="📄"):
        card = tk.Frame(parent, bg="#f9fafb", highlightbackground="#e5e7eb", highlightthickness=1, cursor="hand2")
        card.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        
        tk.Label(card, text=icon, font=("Arial", 26), bg="#f9fafb").pack(pady=(12, 0))
        tk.Label(card, text=title, font=("Arial", 10, "bold"), bg="#f9fafb", wraplength=180, justify="center", fg="#111827").pack(pady=(8, 2), padx=5)
        tk.Label(card, text=description, font=("Arial", 8), bg="#f9fafb", wraplength=180, justify="center", fg="#6b7280").pack(pady=(0, 12), padx=8)

        card.drop_target_register(DND_FILES)
        card.dnd_bind('<<Drop>>', lambda e, act=action_type: self.handle_drop(e, act))
        
        card.bind("<Enter>", lambda e: card.config(bg="#f3f4f6", highlightbackground="#3b82f6"))
        card.bind("<Leave>", lambda e: card.config(bg="#f9fafb", highlightbackground="#e5e7eb"))

    def setup_grid_ui(self):
        container = tk.Frame(self.root, bg="white")
        container.pack(expand=True, fill="both", padx=30)

        for i in range(4): container.grid_columnconfigure(i, weight=1)
        for i in range(3): container.grid_rowconfigure(i, weight=1)

        tools = [
            ("Ảnh sang PDF", "Gộp & Chuyển đổi ảnh sang tệp PDF", "TO_PDF", 0, 0, "📑"),
            ("PDF sang Ảnh (JPG)", "Trích xuất toàn bộ các trang PDF thành ảnh", "PDF_TO_IMG", 0, 1, "🖼️"),
            ("Trích xuất trang PDF", "Trích xuất các trang cụ thể từ tệp PDF", "EXTRACT_PDF", 0, 2, "📄"),
            ("Nén tệp PDF", "Giảm dung lượng PDF tối ưu nhất", "COMPRESS_PDF", 0, 3, "🗜️"),
            
            ("Đổi định dạng Ảnh", "Chuyển đổi qua lại giữa JPG và PNG", "CONVERT_IMG", 1, 0, "🖼️"),
            ("Gộp nhiều PDF", "Ghép nhiều file PDF thành một duy nhất", "MERGE_PDF", 1, 1, "➕"),
            ("Tách PDF làm nhiều file", "Tách PDF thành từng trang riêng biệt", "SPLIT_PDF", 1, 2, "✂️"),
            ("Xóa trang PDF", "Loại bỏ trang không mong muốn khỏi PDF", "DELETE_PDF_PAGES", 1, 3, "🗑️"),
            
            ("Chuyển PDF sang Word", "Chuyển tài liệu PDF sang Word (.doc)", "PDF_TO_WORD", 2, 0, "📘"),
            ("Chuyển Word sang PDF", "Chuyển tài liệu Word (.docx) sang PDF", "WORD_TO_PDF", 2, 1, "📕"),
            ("Thay đổi DPI tệp", "Đổi DPI của PDF hoặc Hình ảnh tùy chọn", "CHANGE_DPI", 2, 2, "⚙️"),
        ]

        for title, desc, act, r, c, ico in tools:
            self.create_tool_card(container, title, desc, act, r, c, ico)

    def get_unique_path(self, folder, base_name, ext):
        output_path = os.path.join(folder, f"{base_name}.{ext}")
        counter = 1
        while os.path.exists(output_path):
            output_path = os.path.join(folder, f"{base_name}_{counter}.{ext}")
            counter += 1
        return output_path

    def handle_drop(self, event, action):
        files = self.root.tk.splitlist(event.data)
        valid_files = [f for f in files if os.path.isfile(f)]
        
        if not valid_files:
            self.update_status("⚠ Không tìm thấy file hợp lệ để xử lý!", "warning")
            return

        if action == "TO_PDF":
            img_files = [f for f in valid_files if os.path.splitext(f)[1].lower() in [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".jfif", ".heic"]]
            if not img_files:
                self.update_status("⚠ Vui lòng chọn định dạng ảnh hỗ trợ để tạo PDF!", "warning")
                return
            self.run_in_thread(self.process_image_to_pdf, img_files)
            
        elif action == "PDF_TO_IMG":
            pdf_files = [f for f in valid_files if f.lower().endswith(".pdf")]
            if not pdf_files:
                self.update_status("⚠ Vui lòng chọn tệp PDF!", "warning")
                return
            self.run_in_thread(self.process_pdf_to_img_multi, pdf_files)

        elif action == "MERGE_PDF":
            pdf_files = [f for f in valid_files if f.lower().endswith(".pdf")]
            if len(pdf_files) < 2:
                self.update_status("⚠ Vui lòng kéo ít nhất 2 file PDF để gộp!", "warning")
                return
            MergeDialog(self.root, pdf_files, lambda sorted_files: self.run_in_thread(self.execute_pdf_merge, sorted_files))

        elif action == "SPLIT_PDF":
            pdf_files = [f for f in valid_files if f.lower().endswith(".pdf")]
            if not pdf_files:
                self.update_status("⚠ Vui lòng chọn tệp PDF!", "warning")
                return
            for f in pdf_files:
                self.run_in_thread(self.process_pdf_split, f)

        elif action == "COMPRESS_PDF":
            pdf_files = [f for f in valid_files if f.lower().endswith(".pdf")]
            if not pdf_files:
                self.update_status("⚠ Vui lòng chọn tệp PDF!", "warning")
                return
            for f in pdf_files:
                self.run_in_thread(self.process_pdf_compress, f)

        elif action == "EXTRACT_PDF":
            pdf_file = next((f for f in valid_files if f.lower().endswith(".pdf")), None)
            if not pdf_file:
                self.update_status("⚠ Vui lòng chọn tệp PDF!", "warning")
                return
            try:
                reader = PdfReader(pdf_file)
                num_pages = len(reader.pages)
                ExtractDialog(self.root, pdf_file, num_pages, lambda path, p_str: self.run_in_thread(self.execute_pdf_extract, path, p_str))
            except Exception as e:
                self.update_status(f"❌ Lỗi đọc PDF: {str(e)}", "error")

        elif action == "DELETE_PDF_PAGES":
            pdf_file = next((f for f in valid_files if f.lower().endswith(".pdf")), None)
            if not pdf_file:
                self.update_status("⚠ Vui lòng chọn tệp PDF!", "warning")
                return
            try:
                reader = PdfReader(pdf_file)
                num_pages = len(reader.pages)
                DeletePagesDialog(self.root, pdf_file, num_pages, lambda path, p_str: self.run_in_thread(self.execute_pdf_delete, path, p_str))
            except Exception as e:
                self.update_status(f"❌ Lỗi đọc PDF: {str(e)}", "error")

        elif action == "CONVERT_IMG":
            img_files = [f for f in valid_files if os.path.splitext(f)[1].lower() in [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".jfif", ".heic"]]
            if not img_files:
                self.update_status("⚠ Vui lòng kéo các tệp ảnh hợp lệ!", "warning")
                return
            ImageFormatDialog(self.root, lambda fmt: self.run_in_thread(self.execute_image_convert_multi, img_files, fmt))

        elif action == "PDF_TO_WORD":
            pdf_files = [f for f in valid_files if f.lower().endswith(".pdf")]
            if not pdf_files:
                self.update_status("⚠ Vui lòng chọn file PDF để chuyển sang Word!", "warning")
                return
            for f in pdf_files:
                self.run_in_thread(self.process_pdf_to_word, f)

        elif action == "WORD_TO_PDF":
            docx_files = [f for f in valid_files if f.lower().endswith(".docx") or f.lower().endswith(".doc")]
            if not docx_files:
                self.update_status("⚠ Vui lòng chọn file Word (.docx hoặc .doc)!", "warning")
                return
            for f in docx_files:
                self.run_in_thread(self.process_word_to_pdf, f)

        elif action == "CHANGE_DPI":
            valid_targets = [f for f in valid_files if f.lower().endswith(".pdf") or os.path.splitext(f)[1].lower() in [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".jfif", ".heic"]]
            if not valid_targets:
                self.update_status("⚠ Chỉ hỗ trợ đổi DPI cho PDF hoặc Ảnh!", "warning")
                return
            ChangeDpiDialog(self.root, lambda dpi_val: self.run_in_thread(self.execute_change_dpi_multi, valid_targets, dpi_val))

    def process_image_to_pdf(self, img_files):
        dialog = ProgressDialog(self.root, "Đang tạo PDF từ ảnh")
        try:
            converted_images = []
            for i, f in enumerate(img_files):
                dialog.set_text(f"Đang xử lý ảnh {i+1}/{len(img_files)}...")
                img = Image.open(f)
                if img.mode in ("RGBA", "LA", "P"):
                    img = img.convert("RGB")
                converted_images.append(img)
            
            if not converted_images:
                return
            
            folder = os.path.dirname(img_files[0])
            output_path = self.get_unique_path(folder, "Images_Combined", "pdf")
            
            dialog.set_text("Đang đóng gói tệp PDF...")
            d_val = self.dpi.get()
            q_val = self.quality.get()
            converted_images[0].save(
                output_path,
                "PDF",
                save_all=True,
                append_images=converted_images[1:],
                resolution=float(d_val),
                quality=q_val
            )
            
            self.update_status(f"✅ Đã tạo PDF gộp từ {len(img_files)} ảnh tại: {os.path.basename(output_path)}", "success")
            open_folder(output_path)
        except Exception as e:
            self.update_status(f"❌ Lỗi tạo PDF: {str(e)}", "error")
        finally:
            dialog.close()

    def process_pdf_to_img_multi(self, pdf_files):
        dialog = ProgressDialog(self.root, "Đang chuyển PDF sang ảnh")
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            pp_path = os.path.join(script_dir, 'poppler-windows', 'Library', 'bin')
            if not os.path.exists(pp_path):
                pp_path = resource_path(os.path.join('poppler-windows', 'Library', 'bin'))
            pp_path = os.path.abspath(pp_path)
            
            if not os.path.exists(pp_path):
                raise Exception("Không tìm thấy bộ công cụ Poppler để trích xuất PDF.")

            original_path = os.environ.get("PATH", "")
            if pp_path not in original_path:
                os.environ["PATH"] = pp_path + os.pathsep + original_path

            for idx, input_path in enumerate(pdf_files):
                dialog.set_text(f"Đang xử lý file {idx+1}/{len(pdf_files)}...")
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
                
            self.update_status(f"✅ Đã chuyển đổi hoàn tất {len(pdf_files)} file PDF sang ảnh!", "success")
            open_folder(os.path.dirname(pdf_files[0]))
        except Exception as e:
            self.update_status(f"❌ Lỗi PDF sang Ảnh: {str(e)}", "error")
        finally:
            dialog.close()

    def execute_pdf_merge(self, pdf_files):
        dialog = ProgressDialog(self.root, "Đang gộp PDF")
        try:
            writer = PdfWriter()
            for pdf in pdf_files:
                reader = PdfReader(pdf)
                for page in reader.pages:
                    writer.add_page(page)
            
            output_path = self.get_unique_path(os.path.dirname(pdf_files[0]), "Merged_Output", "pdf")
            with open(output_path, "wb") as f:
                writer.write(f)
            
            self.update_status(f"✅ Đã gộp {len(pdf_files)} file thành: {os.path.basename(output_path)}", "success")
            open_folder(output_path)
        except Exception as e:
            self.update_status(f"❌ Lỗi gộp PDF: {str(e)}", "error")
        finally:
            dialog.close()

    def process_pdf_split(self, input_path):
        dialog = ProgressDialog(self.root, "Đang tách PDF")
        try:
            reader = PdfReader(input_path)
            num_pages = len(reader.pages)
            if num_pages <= 1:
                self.update_status("⚠ File chỉ có 1 trang, không cần tách!", "warning")
                return

            base_name = os.path.splitext(os.path.basename(input_path))[0]
            parent_folder = os.path.dirname(input_path)
            output_folder = self.get_unique_path(parent_folder, f"{base_name}_Split", "")
            os.makedirs(output_folder, exist_ok=True)
            
            padding = len(str(num_pages))
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                page_num = str(i + 1).zfill(padding)
                output_file = os.path.join(output_folder, f"{base_name}_page_{page_num}.pdf")
                with open(output_file, "wb") as f:
                    writer.write(f)
                    
            self.update_status(f"✅ Đã tách {num_pages} trang vào folder: {os.path.basename(output_folder)}", "success")
            open_folder(output_folder)
        except Exception as e:
            self.update_status(f"❌ Lỗi tách PDF: {str(e)}", "error")
        finally:
            dialog.close()

    def process_pdf_compress(self, input_path):
        dialog = ProgressDialog(self.root, "Đang nén PDF")
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            pp_path = os.path.join(script_dir, 'poppler-windows', 'Library', 'bin')
            if not os.path.exists(pp_path):
                pp_path = resource_path(os.path.join('poppler-windows', 'Library', 'bin'))
            pp_path = os.path.abspath(pp_path)
            
            if not os.path.exists(pp_path):
                raise Exception("Không tìm thấy bộ công cụ Poppler để nén PDF.")

            original_path = os.environ.get("PATH", "")
            if pp_path not in original_path:
                os.environ["PATH"] = pp_path + os.pathsep + original_path

            q = self.quality.get()
            d = self.dpi.get()
            
            pages = convert_from_path(input_path, dpi=d, poppler_path=pp_path)
            if not pages:
                raise Exception("Không đọc được nội dung các trang.")

            output_pages = []
            for page in pages:
                if page.mode != "RGB":
                    page = page.convert("RGB")
                output_pages.append(page)
            
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
            
            self.update_status(f"✅ Đã nén thành công {reduction:.1f}% tại: {os.path.basename(output_path)}", "success")
            open_folder(output_path)
        except Exception as e:
            self.update_status(f"❌ Lỗi nén PDF: {str(e)}", "error")
        finally:
            dialog.close()

    def execute_pdf_extract(self, input_path, pages_str):
        dialog = ProgressDialog(self.root, "Đang trích xuất trang")
        try:
            reader = PdfReader(input_path)
            total_pages = len(reader.pages)
            writer = PdfWriter()
            
            selected_pages = set()
            parts = pages_str.replace(" ", "").split(",")
            for part in parts:
                if "-" in part:
                    start, end = map(int, part.split("-"))
                    selected_pages.update(range(start, end + 1))
                else:
                    selected_pages.add(int(part))
            
            valid_pages = sorted([p-1 for p in selected_pages if 1 <= p <= total_pages])
            if not valid_pages:
                messagebox.showerror("Lỗi", "Không có trang nào hợp lệ!")
                return
            
            for p_idx in valid_pages:
                writer.add_page(reader.pages[p_idx])
            
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = self.get_unique_path(os.path.dirname(input_path), f"{base_name}_extracted", "pdf")
            
            with open(output_path, "wb") as f:
                writer.write(f)
                
            self.update_status(f"✅ Đã trích {len(valid_pages)} trang thành công tại: {os.path.basename(output_path)}", "success")
            open_folder(output_path)
        except Exception as e:
            self.update_status(f"❌ Lỗi trích xuất: {str(e)}", "error")
        finally:
            dialog.close()

    def execute_pdf_delete(self, input_path, pages_str):
        dialog = ProgressDialog(self.root, "Đang xóa trang")
        try:
            reader = PdfReader(input_path)
            total_pages = len(reader.pages)
            writer = PdfWriter()
            
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
            
            pages_to_keep = [i for i in range(total_pages) if (i + 1) not in removed_pages]
            if not pages_to_keep:
                messagebox.showerror("Lỗi", "Bạn không thể xóa toàn bộ tất cả các trang!")
                return
            
            for p_idx in pages_to_keep:
                writer.add_page(reader.pages[p_idx])
            
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = self.get_unique_path(os.path.dirname(input_path), f"{base_name}_removed_pages", "pdf")
            
            with open(output_path, "wb") as f:
                writer.write(f)
                
            self.update_status(f"✅ Đã xóa {total_pages - len(pages_to_keep)} trang. Còn lại {len(pages_to_keep)} trang.", "success")
            open_folder(output_path)
        except Exception as e:
            self.update_status(f"❌ Lỗi xóa trang: {str(e)}", "error")
        finally:
            dialog.close()

    def execute_image_convert_multi(self, img_files, fmt):
        dialog = ProgressDialog(self.root, "Đang chuyển định dạng ảnh")
        try:
            for idx, input_path in enumerate(img_files):
                dialog.set_text(f"Đang xử lý ảnh {idx+1}/{len(img_files)}...")
                folder = os.path.dirname(input_path)
                name = os.path.splitext(os.path.basename(input_path))[0]
                ext = fmt.lower()
                output_path = self.get_unique_path(folder, f"{name}_converted", ext)

                with Image.open(input_path) as img:
                    if fmt == "JPEG" and img.mode in ("RGBA", "LA", "P"):
                        img = img.convert("RGB")
                    
                    q = self.quality.get()
                    d = self.dpi.get()
                    if fmt == "PNG":
                        img.save(output_path, "PNG", compress_level=9, optimize=True, dpi=(d, d))
                    else:
                        img.save(output_path, "JPEG", optimize=True, quality=q, dpi=(d, d))
            
            self.update_status(f"✅ Đã chuyển đổi hoàn tất {len(img_files)} ảnh sang {fmt}!", "success")
            open_folder(os.path.dirname(img_files[0]))
        except Exception as e:
            self.update_status(f"❌ Lỗi chuyển đổi ảnh: {str(e)}", "error")
        finally:
            dialog.close()

    def process_pdf_to_word(self, input_path):
        dialog = ProgressDialog(self.root, "Đang chuyển PDF sang Word")
        try:
            reader = PdfReader(input_path)
            
            doc_content = ""
            for i, page in enumerate(reader.pages):
                dialog.set_text(f"Đang trích xuất văn bản trang {i+1}/{len(reader.pages)}...")
                text = page.extract_text()
                if text:
                    doc_content += text + "\n\n"

            # Create an HTML-based document mimicking the web version behavior
            html = f"""<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
            <head><meta charset="utf-8"/><title>Converted Document</title>
            <style>
            body {{ font-family: "Times New Roman", Times, serif; font-size: 12pt; line-height: 1.5; }}
            p {{ margin: 0 0 12pt 0; white-space: pre-wrap; }}
            </style>
            </head>
            <body>
            <div>
            {doc_content.replace("\n", "<br/>")}
            </div>
            </body>
            </html>"""

            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = self.get_unique_path(os.path.dirname(input_path), base_name, "doc")
            
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(html)

            self.update_status(f"✅ Đã chuyển đổi PDF sang Word thành công: {os.path.basename(output_path)}", "success")
            open_folder(output_path)
        except Exception as e:
            self.update_status(f"❌ Lỗi chuyển PDF sang Word: {str(e)}", "error")
        finally:
            dialog.close()

    def process_word_to_pdf(self, input_path):
        dialog = ProgressDialog(self.root, "Đang chuyển Word sang PDF")
        try:
            comtypes.CoInitialize()
            
            abs_input_path = os.path.abspath(input_path)
            base_name = os.path.splitext(os.path.basename(abs_input_path))[0]
            output_path = self.get_unique_path(os.path.dirname(abs_input_path), base_name, "pdf")
            
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            
            doc = word.Documents.Open(abs_input_path)
            doc.SaveAs(output_path, FileFormat=17)
            doc.Close()
            word.Quit()
            
            self.update_status(f"✅ Đã chuyển Word sang PDF thành công: {os.path.basename(output_path)}", "success")
            open_folder(output_path)
        except Exception as e:
            self.update_status(f"❌ Lỗi chuyển Word sang PDF: {str(e)}", "error")
        finally:
            comtypes.CoUninitialize()
            dialog.close()

    def execute_change_dpi_multi(self, files, target_dpi):
        dialog = ProgressDialog(self.root, f"Đang thay đổi DPI về {target_dpi}")
        try:
            for idx, input_path in enumerate(files):
                dialog.set_text(f"Đang xử lý tệp {idx+1}/{len(files)}...")
                f_ext = os.path.splitext(input_path)[1].lower()
                folder = os.path.dirname(input_path)
                name = os.path.splitext(os.path.basename(input_path))[0]
                output_path = self.get_unique_path(folder, f"{name}_{target_dpi}dpi", f_ext[1:])

                if f_ext == ".pdf":
                    script_dir = os.path.dirname(os.path.abspath(__file__))
                    pp_path = os.path.join(script_dir, 'poppler-windows', 'Library', 'bin')
                    if not os.path.exists(pp_path):
                        pp_path = resource_path(os.path.join('poppler-windows', 'Library', 'bin'))
                    pp_path = os.path.abspath(pp_path)
                    
                    if not os.path.exists(pp_path):
                        raise Exception("Không tìm thấy bộ công cụ Poppler.")
                    
                    pages = convert_from_path(input_path, dpi=target_dpi, poppler_path=pp_path)
                    output_pages = []
                    for page in pages:
                        if page.mode != "RGB":
                            page = page.convert("RGB")
                        output_pages.append(page)
                        
                    output_pages[0].save(
                        output_path, 
                        "PDF", 
                        save_all=True, 
                        append_images=output_pages[1:], 
                        resolution=float(target_dpi)
                    )
                else:
                    with Image.open(input_path) as img:
                        img.save(output_path, dpi=(target_dpi, target_dpi))

            self.update_status(f"✅ Đã đổi DPI thành công {len(files)} tệp về {target_dpi}!", "success")
            open_folder(os.path.dirname(files[0]))
        except Exception as e:
            self.update_status(f"❌ Lỗi đổi DPI: {str(e)}", "error")
        finally:
            dialog.close()

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ProToolboxApp(root)
    root.mainloop()
