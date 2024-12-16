from flask import Flask, request, send_from_directory, render_template
import os
from markitdown import MarkItDown

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
ALLOWED_EXTENSIONS = {'docx', 'xlsx', 'pptx', 'doc', 'xls', 'ppt'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file part'
        file = request.files['file']
        if file.filename == '':
            return 'No selected file'
        if file and allowed_file(file.filename):
            try:
                upload_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
                file.save(upload_path)

                markitdown_converter = MarkItDown()
                markdown_output = markitdown_converter.convert(upload_path)
                markdown_text = markdown_output.text_content

                download_filename = file.filename.rsplit('.', 1)[0] + ".md"
                download_path = os.path.join(app.config['DOWNLOAD_FOLDER'], download_filename)

                with open(download_path, "w", encoding="utf-8") as md_file:
                    md_file.write(markdown_text)

                return send_from_directory(app.config['DOWNLOAD_FOLDER'], download_filename, as_attachment=True)

            except Exception as e:
                return f"An error occurred during processing: {e}"
        else:
            return "Allowed file types are docx, xlsx, pptx, doc, xls, ppt"
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, port=5555)