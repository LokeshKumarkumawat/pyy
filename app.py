from flask import Flask, render_template, request, redirect, url_for
from werkzeug.utils import secure_filename
from spire.presentation import *
from spire.presentation.common import *

app = Flask(__name__)

# Set the upload folder and allowed extensions
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'pptx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)

    file = request.files['file']

    if file.filename == '':
        return redirect(request.url)

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        convert_to_pdf(file_path)

        return redirect(url_for('index'))

    return redirect(request.url)

def convert_to_pdf(input_path):
    presentation = Presentation()
    presentation.LoadFromFile(input_path)

    pdf_output_path = os.path.splitext(input_path)[0] + '.pdf'

    presentation.SaveToFile(pdf_output_path, FileFormat.PDF)
    presentation.Dispose()

if __name__ == '__main__':
    app.run(debug=True)
