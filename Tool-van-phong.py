import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image
import os
import sys
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

class ProToolboxApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Tool Van Phong - v.2.0 - Brillian Pham")
        self.root.geometry("900x600")
        self.root.configure(bg="white")
        
        self.quality = tk.IntVar(value=85)
        self.setup_menu()
        self.setup_header()
        self.setup_grid_ui()
        
        self.status = tk.Label(self.root, text="Sẵn sàng! Chỉ cần kéo thả", bg="white", fg="green", font=("Arial", 10))
        self.status.pack(side="bottom", pady=15)

    def setup_menu(self):
        menubar = tk.Menu(self.root)
        q_menu = tk.Menu(menubar, tearoff=0)
        for q in [100, 90, 80, 70, 60, 50, 40, 30, 20, 10]:
            q_menu.add_radiobutton(label=f"Chất lượng: {q}%", variable=self.quality, value=q)
        menubar.add_cascade(label="Chất lượng nén", menu=q_menu)
        
        about_menu = tk.Menu(menubar, tearoff=0)
        about_menu.add_command(label="Thông tin", command=lambda: messagebox.showinfo("Tác giả", "Tool by Brillian Pham\nPhiên bản 2.0"))
        menubar.add_cascade(label="Hỗ trợ", menu=about_menu)
        self.root.config(menu=menubar)

    def setup_header(self):
        header_frame = tk.Frame(self.root, bg="white")
        header_frame.pack(fill="x", pady=25)
        tk.Label(header_frame, text="Tool Văn Phòng", font=("Arial", 26, "bold"), bg="white", fg="#333").pack()
        tk.Label(header_frame, text="Chọn chất lượng nén và kéo thả tệp vào ô tương ứng để xử lý nhanh", font=("Arial", 11), bg="white", fg="#666").pack(pady=5)

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
            ("Chuyển sang WEBP", "WEBP", 0, 2, "🌐"),
            ("Nén tệp PDF", "COMPRESS_PDF", 0, 3, "🗜️"),
            ("Nén và chuyển sang JPEG", "JPEG", 1, 0, "📉"),
            ("Nén và chuyển sang PNG", "PNG", 1, 1, "📦"),
            ("Gộp nhiều PDF", "MERGE_PDF", 1, 2, "➕"),
            ("Tách 1 PDF thành nhiều trang", "SPLIT_PDF", 1, 3, "✂️"),
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
            else:
                # Các action chuyển đổi ảnh: TO_PDF, WEBP, JPEG, PNG
                if f_ext in [".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff"]:
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
            pages = convert_from_path(input_path, poppler_path=pp_path)
            
            folder = os.path.dirname(input_path)
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            
            for i, page in enumerate(pages):
                output = self.get_unique_path(folder, f"{base_name}_page_{i+1}", "jpg")
                page.save(output, "JPEG", quality=self.quality.get())
            
            self.status.config(text=f"✅ Thành công! Đã lưu {len(pages)} ảnh vào: {os.path.basename(folder)}", fg="green")
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
        except Exception as e:
            self.status.config(text=f"Lỗi Gộp PDF: {str(e)}", fg="red")

    def process_pdf_split(self, input_path):
        """Tách 1 file PDF thành nhiều file PDF từng trang"""
        if not input_path.lower().endswith(".pdf"):
            self.status.config(text="Lỗi: Vui lòng chọn file PDF!", fg="red")
            return
            
        try:
            reader = PdfReader(input_path)
            base_name = os.path.splitext(input_path)[0]
            
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                output = self.get_unique_path(os.path.dirname(input_path), f"{os.path.basename(base_name)}_page_{i+1}", "pdf")
                with open(output, "wb") as f:
                    writer.write(f)
                    
            self.status.config(text=f"✅ Đã tách thành {len(reader.pages)} trang PDF!", fg="green")
        except Exception as e:
            self.status.config(text=f"Lỗi Tách PDF: {str(e)}", fg="red")

    def process_pdf_compress(self, input_path):
        """Nén file PDF (giảm kích thước tệp)"""
        if not input_path.lower().endswith(".pdf"):
            self.status.config(text="Lỗi: Vui lòng chọn file PDF!", fg="red")
            return
            
        try:
            self.status.config(text="Đang nén PDF (vui lòng đợi)...", fg="blue")
            self.root.update_idletasks()
            
            reader = PdfReader(input_path)
            writer = PdfWriter()
            
            for page in reader.pages:
                writer.add_page(page)
            
            # 1. Nén luồng nội dung (text, vectors)
            for page in writer.pages:
                page.compress_content_streams()
            
            # 2. Loại bỏ metadata không cần thiết
            writer.add_metadata({})
            
            # 3. Sử dụng tính năng giảm chất lượng ảnh nếu pypdf phiên bản mới hỗ trợ
            # (Ở mức độ cơ bản, pypdf chủ yếu nén stream, để nén mạnh cần giảm DPI ảnh)
            # Thử nén ảnh nếu có thể (pypdf 3.0+)
            for page in writer.pages:
                if "/Resources" in page and "/XObject" in page["/Resources"]:
                    xobjects = page["/Resources"]["/XObject"]
                    for obj in xobjects:
                        if xobjects[obj]["/Subtype"] == "/Image":
                            # Đây là nơi có thể can thiệp sâu hơn, nhưng pypdf giới hạn
                            pass

            base_name_str = os.path.splitext(os.path.basename(input_path))[0]
            output_path = self.get_unique_path(os.path.dirname(input_path), f"{base_name_str}_compressed", "pdf")
            
            with open(output_path, "wb") as f:
                # 4. Ghi với các tùy chọn tối ưu hóa
                writer.write(f)
                
            # Kiểm tra hiệu quả
            orig_size = os.path.getsize(input_path)
            new_size = os.path.getsize(output_path)
            reduction = (orig_size - new_size) / orig_size * 100
            
            if reduction > 0.1:
                self.status.config(text=f"✅ Đã nén {reduction:.1f}%: {os.path.basename(output_path)}", fg="green")
            else:
                self.status.config(text=f"⚠ Nén xong (giảm {reduction:.1f}%). Thử giảm chất lượng ở menu trên.", fg="orange")
        except Exception as e:
            self.status.config(text=f"Lỗi Nén PDF: {str(e)}", fg="red")

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
                if fmt == "TO_PDF":
                    img.save(output_path, "PDF", resolution=100.0, quality=q)
                elif fmt == "PNG":
                    # PNG sử dụng compress_level (0-9) thay vì quality
                    cl = max(0, min(9, (100-q)//10))
                    img.save(output_path, "PNG", compress_level=cl)
                elif fmt == "WEBP":
                    img.save(output_path, "WEBP", quality=q, lossless=(q==100))
                else: # JPEG
                    img.save(output_path, "JPEG", optimize=True, quality=q)
            
            self.status.config(text=f"✅ Đã lưu: {os.path.basename(output_path)}", fg="green")
        except Exception as e:
            self.status.config(text=f"❌ Lỗi xử lý ảnh: {str(e)}", fg="red")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ProToolboxApp(root)
    root.mainloop()
