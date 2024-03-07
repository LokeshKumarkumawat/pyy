import os
from flask import Flask, render_template, request, send_file, redirect, url_for
from werkzeug.utils import secure_filename
from spire.presentation import *
from spire.presentation.common import *
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime, timedelta  # Add this import

app = Flask(__name__)

# Set the upload folder and allowed extensions
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'pptx', 'pdf'}

scheduler = BackgroundScheduler(daemon=True)
scheduler.start()

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

        pdf_filename = os.path.splitext(filename)[0] + '.pdf'  # Define pdf_filename here

        # Schedule cleanup job after one minute
        scheduler.add_job(delete_files, 'date', run_date=datetime.now() + timedelta(minutes=1),
                          args=[file_path, os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)])

        return redirect(url_for('download_pdf', filename=filename))

    return redirect(request.url)

@app.route('/download/<filename>')
def download_pdf(filename):
    pdf_filename = os.path.splitext(filename)[0] + '.pdf'
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)

    return send_file(pdf_path, as_attachment=True, download_name=pdf_filename)

def convert_to_pdf(input_path):
    presentation = Presentation()
    presentation.LoadFromFile(input_path)

    pdf_output_path = os.path.splitext(input_path)[0] + '.pdf'

    presentation.SaveToFile(pdf_output_path, FileFormat.PDF)
    presentation.Dispose()

def delete_files(input_path, pdf_path):
    os.remove(input_path)
    os.remove(pdf_path)

if __name__ == '__main__':
    # app.run(debug=False)
    http_server = WSGIServer(('0.0.0.0', 5000), app)
    http_server.serve_forever()
