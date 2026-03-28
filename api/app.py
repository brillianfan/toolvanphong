import os
import io
import uuid
import zipfile
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
from PIL import Image
try:
    import pillow_heif
    # Register HEIC opener
    pillow_heif.register_heif_opener()
except ImportError:
    pillow_heif = None

app = Flask(__name__, template_folder='../templates')
app.config['UPLOAD_FOLDER'] = '/tmp'
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB limit for server processing

def parse_pages(pages_str, total_pages):
    """Parse page range string (e.g., '1, 3, 5-8') into a list of 0-indexed integers."""
    selected_pages = set()
    if not pages_str:
        return []
    parts = pages_str.replace(" ", "").split(",")
    for part in parts:
        try:
            if "-" in part:
                start, end = map(int, part.split("-"))
                selected_pages.update(range(start, end + 1))
            else:
                selected_pages.add(int(part))
        except ValueError:
            continue
    
    # Filter valid pages and convert to 0-indexed
    return sorted([p - 1 for p in selected_pages if 1 <= p <= total_pages])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_file():
    action = request.form.get('action')
    quality = int(request.form.get('quality', 85))
    dpi = int(request.form.get('dpi', 300))
    pages_str = request.form.get('pages', '')
    files = request.files.getlist('files')

    if not files:
        return jsonify({"error": "Không có tệp nào được tải lên"}), 400

    try:
        if action == "MERGE_PDF":
            writer = PdfWriter()
            # Order is determined by the order in the 'files' list from the form
            for f in files:
                reader = PdfReader(f)
                for page in reader.pages:
                    writer.add_page(page)
            
            output = io.BytesIO()
            writer.write(output)
            output.seek(0)
            return send_file(output, mimetype='application/pdf', as_attachment=True, download_name='merged.pdf')

        elif action == "EXTRACT_PDF":
            f = files[0]
            reader = PdfReader(f)
            total = len(reader.pages)
            valid_pages = parse_pages(pages_str, total)
            
            if not valid_pages:
                return jsonify({"error": "Không có trang hợp lệ để trích xuất"}), 400
                
            writer = PdfWriter()
            for p_idx in valid_pages:
                writer.add_page(reader.pages[p_idx])
            
            output = io.BytesIO()
            writer.write(output)
            output.seek(0)
            return send_file(output, mimetype='application/pdf', as_attachment=True, download_name='extracted.pdf')

        elif action == "DELETE_PDF_PAGES":
            f = files[0]
            reader = PdfReader(f)
            total = len(reader.pages)
            removed_pages = set(parse_pages(pages_str, total))
            
            pages_to_keep = [i for i in range(total) if i not in removed_pages]
            if not pages_to_keep:
                return jsonify({"error": "Bạn không thể xóa tất cả các trang"}), 400
                
            writer = PdfWriter()
            for p_idx in pages_to_keep:
                writer.add_page(reader.pages[p_idx])
            
            output = io.BytesIO()
            writer.write(output)
            output.seek(0)
            return send_file(output, mimetype='application/pdf', as_attachment=True, download_name='removed_pages.pdf')

        elif action == "SPLIT_PDF":
            f = files[0]
            reader = PdfReader(f)
            total = len(reader.pages)
            
            if total <= 1:
                return jsonify({"error": "Tệp chỉ có 1 trang, không thể tách"}), 400
                
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for i in range(total):
                    writer = PdfWriter()
                    writer.add_page(reader.pages[i])
                    page_io = io.BytesIO()
                    writer.write(page_io)
                    zip_file.writestr(f"page_{i+1}.pdf", page_io.getvalue())
            
            zip_buffer.seek(0)
            return send_file(zip_buffer, mimetype='application/zip', as_attachment=True, download_name='split_pages.zip')

        elif action == "COMPRESS_PDF":
            # Better compression: Convert to Image then back to PDF
            f = files[0]
            doc = fitz.open(stream=f.read(), filetype="pdf")
            
            output_pdf_io = io.BytesIO()
            img_list = []
            
            for i in range(len(doc)):
                page = doc.load_page(i)
                # Apply DPI
                zoom = dpi / 72
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                img_list.append(img)
            
            if img_list:
                img_list[0].save(
                    output_pdf_io, 
                    "PDF", 
                    save_all=True, 
                    append_images=img_list[1:], 
                    quality=quality,
                    optimize=True,
                    resolution=float(dpi)
                )
            
            output_pdf_io.seek(0)
            return send_file(output_pdf_io, mimetype='application/pdf', as_attachment=True, download_name='compressed.pdf')

        elif action == "PDF_TO_IMG":
            f = files[0]
            doc = fitz.open(stream=f.read(), filetype="pdf")
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for i in range(len(doc)):
                    page = doc.load_page(i)
                    zoom = dpi / 72
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    
                    img_io = io.BytesIO()
                    img.save(img_io, format="JPEG", quality=quality, dpi=(dpi, dpi))
                    zip_file.writestr(f"page_{i+1}.jpg", img_io.getvalue())
            
            zip_buffer.seek(0)
            return send_file(zip_buffer, mimetype='application/zip', as_attachment=True, download_name='pdf_images.zip')

        elif action in ["TO_PDF", "WEBP", "JPEG", "PNG", "CONVERT_IMG"]:
            # Default to first file for single conversion, or loop if multi
            actual_fmt = request.form.get('format', action)
            if actual_fmt == "CONVERT_IMG": actual_fmt = "JPEG" # Fallback
            
            if len(files) > 1:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for idx, f in enumerate(files):
                        img = Image.open(f)
                        output_io = io.BytesIO()
                        
                        ext = actual_fmt.lower()
                        if actual_fmt == "TO_PDF":
                            if img.mode != 'RGB': img = img.convert('RGB')
                            img.save(output_io, format="PDF", resolution=float(dpi), quality=quality)
                            ext = "pdf"
                        elif actual_fmt == "PNG":
                            cl = max(0, min(3, (100 - quality) // 10))
                            img.save(output_io, format="PNG", compress_level=cl, dpi=(dpi, dpi))
                        else: # JPEG, WEBP
                            if actual_fmt == "JPEG" and img.mode != 'RGB': img = img.convert('RGB')
                            img.save(output_io, format=actual_fmt, quality=quality, dpi=(dpi, dpi))
                            
                        zip_file.writestr(f"converted_{idx+1}.{ext}", output_io.getvalue())
                zip_buffer.seek(0)
                return send_file(zip_buffer, mimetype='application/zip', as_attachment=True, download_name='converted_images.zip')
            else:
                f = files[0]
                img = Image.open(f)
                output_io = io.BytesIO()
                
                if actual_fmt == "TO_PDF":
                    if img.mode != 'RGB': img = img.convert('RGB')
                    img.save(output_io, format="PDF", resolution=float(dpi), quality=quality)
                    output_io.seek(0)
                    return send_file(output_io, mimetype='application/pdf', as_attachment=True, download_name='converted.pdf')
                
                elif actual_fmt == "PNG":
                    cl = max(0, min(3, (100 - quality) // 10))
                    img.save(output_io, format="PNG", compress_level=cl, dpi=(dpi, dpi))
                    output_io.seek(0)
                    return send_file(output_io, mimetype='image/png', as_attachment=True, download_name='converted.png')
                
                else: # JPEG, WEBP
                    if actual_fmt == "JPEG" and img.mode != 'RGB': img = img.convert('RGB')
                    img.save(output_io, format=actual_fmt, quality=quality, dpi=(dpi, dpi))
                    output_io.seek(0)
                    return send_file(output_io, mimetype=f'image/{actual_fmt.lower()}', as_attachment=True, download_name=f'converted.{actual_fmt.lower()}')

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
