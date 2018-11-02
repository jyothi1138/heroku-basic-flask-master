import comtypes
import os
from os import listdir
import img2pdf
from PIL import Image
from flask import Flask, render_template, request, flash, redirect, send_from_directory
from fpdf import FPDF
from werkzeug.utils import secure_filename


UPLOAD_FOLDER = '\\upload\\'
ALLOWED_EXTENSIONS = set(['txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif'])

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/')
def index():
    return render_template('index.html')


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/upload_file', methods=['POST', 'GET'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # if user does not select file, browser also
        # submit a empty part without filename
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            sourcepath = os.path.join('upload', filename)
            destinationpath = os.path.join('upload', "file.pdf")
            file.save(sourcepath)
            # storing image path
            img_path = sourcepath
            # storing pdf path
            pdf_path = destinationpath
            # opening image
            image = Image.open(img_path)
            # converting into chunks using img2pdf
            pdf_bytes = img2pdf.convert(image.filename)
            # opening or creating pdf file
            file = open(pdf_path, "wb")
            # writing pdf files with chunks
            file.write(pdf_bytes)
            # closing image file
            image.close()
            # closing pdf file
            file.close()
            # output
            print("Successfully made pdf file")
            return send_from_directory(directory='upload',
                                       filename='file.pdf',
                                       mimetype='application/pdf')


@app.route('/upload_multipe_file', methods=['POST', 'GET'])
def upload_multipe_file():
    if request.method == 'POST':
        uploaded_files = request.files.getlist("images")
        for i in range(len(uploaded_files)):
            sourcepath = os.path.join('upload', uploaded_files[i].filename)
            uploaded_files[i].save(sourcepath)

        pdf = FPDF('P', 'mm', 'A4')  # create an A4-size pdf document
        x, y, w, h = 0, 0, 200, 250
        mypath = "upload"  # path to your Image directory
        for each_file in listdir(mypath):
            pdf.add_page()
            path1 = os.path.join(mypath, each_file)
            pdf.image(path1, x, y, w, h)

        pdf.output("images.pdf", "F")
        return ""


@app.route('/upload_doc_file', methods=['POST', 'GET'])
def upload_doc_file():
    if request.method == 'POST':
        file = request.files['file']
        filename = secure_filename(file.filename)
        new_name = filename.replace(".docx", r".pdf")
        sourcepath = os.path.join('upload', filename)
        destinationpath = os.path.join('upload', "file.pdf")
        file.save(sourcepath)
        wdFormatPDF = 17
        word = comtypes.client.Dispatch('Word.Application')
        word.Visible = True
        word.Documents.Open(os.path.abspath(sourcepath))
        word.Documents[0].SaveAs(os.path.abspath(destinationpath), 17)
        word.Documents[0].Close()
        word.Quit()
        return filename


if __name__ == '__main__':
    app.run()
