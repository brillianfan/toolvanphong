import os
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
from PIL import Image
import io
import uuid

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = '/tmp'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

def get_unique_filename(extension):
    return f"{uuid.uuid4()}.{extension}"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_file():
    action = request.form.get('action')
    quality = int(request.form.get('quality', 85))
    files = request.files.getlist('files')

    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    try:
        if action == "MERGE_PDF":
            writer = PdfWriter()
            for f in files:
                reader = PdfReader(f)
                for page in reader.pages:
                    writer.add_page(page)
            
            output = io.BytesIO()
            writer.write(output)
            output.seek(0)
            return send_file(output, mimetype='application/pdf', as_attachment=True, download_name='merged.pdf')

        elif action == "SPLIT_PDF":
            # For split, we take the first file and return a zip or just the first page for simplicity in this MVP
            # But let's try to be helpful. If it's one file, we split it.
            f = files[0]
            reader = PdfReader(f)
            # In a real app we'd zip these, but for a single page split:
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            output = io.BytesIO()
            writer.write(output)
            output.seek(0)
            return send_file(output, mimetype='application/pdf', as_attachment=True, download_name='page_1.pdf')

        elif action == "COMPRESS_PDF":
            f = files[0]
            reader = PdfReader(f)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
                page.compress_content_streams()
            
            output = io.BytesIO()
            writer.write(output)
            output.seek(0)
            return send_file(output, mimetype='application/pdf', as_attachment=True, download_name='compressed.pdf')

        elif action == "PDF_TO_IMG":
            f = files[0]
            doc = fitz.open(stream=f.read(), filetype="pdf")
            page = doc.load_page(0)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            output = io.BytesIO()
            img.save(output, format="JPEG", quality=quality)
            output.seek(0)
            return send_file(output, mimetype='image/jpeg', as_attachment=True, download_name='page_1.jpg')

        elif action in ["TO_PDF", "WEBP", "JPEG", "PNG"]:
            f = files[0]
            img = Image.open(f)
            
            output = io.BytesIO()
            if action == "TO_PDF":
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                img.save(output, format="PDF")
                output.seek(0)
                return send_file(output, mimetype='application/pdf', as_attachment=True, download_name='converted.pdf')
            
            elif action == "WEBP":
                img.save(output, format="WEBP", quality=quality)
                output.seek(0)
                return send_file(output, mimetype='image/webp', as_attachment=True, download_name='converted.webp')
            
            elif action == "JPEG":
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                img.save(output, format="JPEG", quality=quality)
                output.seek(0)
                return send_file(output, mimetype='image/jpeg', as_attachment=True, download_name='converted.jpg')
            
            elif action == "PNG":
                # PNG use compress_level (0-9)
                cl = max(0, min(9, (100 - quality) // 10))
                img.save(output, format="PNG", compress_level=cl)
                output.seek(0)
                return send_file(output, mimetype='image/png', as_attachment=True, download_name='converted.png')

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
