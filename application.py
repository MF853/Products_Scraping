from flask import Flask, render_template, request, send_file
import os
from werkzeug.utils import secure_filename
import subprocess
import psutil
from threading import Thread
export FLASK_APP=application.py


application = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
MODIFIED_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

application.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Store reference to the running process
running_process = None

@application.route('/')
def index():
    print("Página Inicial")
    return render_template('index.html')

@application.route('/upload', methods=['POST'])
def upload_file():
    global running_process
    if 'file' not in request.files:
        return render_template('index.html', message='No file part')
    
    file = request.files['file']

    if file.filename == '':
        return render_template('index.html', message='No file selected')

    if file and allowed_file(file.filename):
        if not os.path.exists(application.config['UPLOAD_FOLDER']):
            os.makedirs(application.config['UPLOAD_FOLDER'])
        if not os.path.exists(MODIFIED_FOLDER):
            os.makedirs(MODIFIED_FOLDER)
        filename = secure_filename(file.filename)
        file.save(os.path.join(application.config['UPLOAD_FOLDER'], filename))
        # Start processing the excel file
        script_path = os.path.join(os.path.dirname(__file__), 'script.py')
        running_process = subprocess.Popen(['python', script_path])
        return render_template('index.html', message='File successfully uploaded', modified_file=filename)

@application.route('/download-modified-excel/<filename>')
def download_modified_excel(filename):
    modified_filepath = os.path.join(MODIFIED_FOLDER, filename)
    return send_file(modified_filepath, as_attachment=True)

@application.route('/stop-script')
def stop_script():
    global running_process
    if running_process is not None:
        # Terminate the running process
        running_process.terminate()
        running_process = None
        return render_template('index.html', message='Script stopped')
    else:
        # If no running process found
        return render_template('index.html', message='No running script found')


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

if __name__ == '__main__':
    application.run(debug=True)
