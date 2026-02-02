from flask import Flask, request, render_template, redirect, url_for, send_from_directory, flash, jsonify
from werkzeug.utils import secure_filename
import os
import pandas as pd
from convert_to_xls import pdf_to_excel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'converted')
ALLOWED_EXTENSIONS = {'pdf'}

for d in (UPLOAD_FOLDER, OUTPUT_FOLDER):
    os.makedirs(d, exist_ok=True)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.secret_key = 'change-me'


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    files = sorted(os.listdir(app.config['OUTPUT_FOLDER']))
    return render_template('index.html', files=files)


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(url_for('index'))
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        saved_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(saved_path)
        # Build output path with same base name and .xls extension
        base = os.path.splitext(filename)[0]
        out_name = base + '.xls'
        out_path = os.path.join(app.config['OUTPUT_FOLDER'], out_name)
        try:
            pdf_to_excel(saved_path, out_path)
        except Exception as e:
            flash(f'Conversion failed: {e}')
            return redirect(url_for('index'))
        # Après conversion, afficher un message et revenir à l'accueil
        flash(f'Conversion terminée: {out_name}')
        return redirect(url_for('index'))
    else:
        flash('Only PDF files are allowed')
        return redirect(url_for('index'))


@app.route('/download/<path:filename>')
def download(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)


@app.route('/preview/<path:filename>')
def preview(filename):
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404
    try:
        excel_file = pd.ExcelFile(file_path)
        sheets = {}
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            sheets[sheet_name] = df.to_html(classes='preview-table', max_rows=100)
        return jsonify({'sheets': sheets, 'filename': filename})
    except Exception as e:
        return jsonify({'error': str(e)}), 400


if __name__ == '__main__':
    app.run(debug=True)
