from flask import Flask, request, render_template, send_file
import os, zipfile
from datetime import datetime
from marks_code import generate_individual_pdfs  # your PDF generation code

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    # Clear old PDFs in output folder
    for old_file in os.listdir(OUTPUT_FOLDER):
        if old_file.endswith(".pdf"):
            os.remove(os.path.join(OUTPUT_FOLDER, old_file))

    # Run your marks code (generate PDFs for each student)
    generate_individual_pdfs(filepath, OUTPUT_FOLDER)

    # Create a unique ZIP file
    zip_filename = f"marks_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)

    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for pdf in os.listdir(OUTPUT_FOLDER):
            if pdf.endswith(".pdf") and pdf != zip_filename:  # avoid adding zip into itself
                zipf.write(
                    os.path.join(OUTPUT_FOLDER, pdf),
                    arcname=pdf
                )

    return send_file(zip_path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
