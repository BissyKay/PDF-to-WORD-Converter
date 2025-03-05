import os
import logging
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from pdf2docx import Converter
from werkzeug.utils import secure_filename

# Initialize Flask app
app = Flask(__name__)
app.secret_key = "supersecretkey"  # Used for flash messages

# File upload settings
UPLOAD_FOLDER = "uploads"
CONVERTED_FOLDER = "converted"
ALLOWED_EXTENSIONS = {"pdf"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["CONVERTED_FOLDER"] = CONVERTED_FOLDER

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

# Configure logging
logging.basicConfig(filename="conversion.log", level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")


def allowed_file(filename):
    """Check if the uploaded file has a valid extension."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET", "POST"])
def upload_file():
    """Handle file uploads and conversion requests."""
    if request.method == "POST":
        if "file" not in request.files:
            flash("No file uploaded. Please select a PDF file.", "error")
            return redirect(url_for("upload_file"))

        file = request.files["file"]

        if file.filename == "":
            flash("No selected file. Please choose a PDF file to convert.", "error")
            return redirect(url_for("upload_file"))

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(pdf_path)

            # Convert PDF to Word
            docx_filename = filename.rsplit(".", 1)[0] + ".docx"
            docx_path = os.path.join(app.config["CONVERTED_FOLDER"], docx_filename)

            try:
                logging.info(f"Converting file: {filename}")
                cv = Converter(pdf_path)
                cv.convert(docx_path, start=0, end=None)
                cv.close()
                logging.info(f"Conversion successful: {docx_filename}")

                flash("Conversion successful! Download your Word file below.", "success")
                return redirect(url_for("download_file", filename=docx_filename))

            except Exception as e:
                logging.error(f"Error during conversion: {str(e)}")
                flash("Error converting file. Please try again later.", "error")
                return redirect(url_for("upload_file"))

        else:
            flash("Invalid file type. Please upload a valid PDF file.", "error")
            return redirect(url_for("upload_file"))

    return render_template("index.html")


@app.route("/download/<filename>")
def download_file(filename):
    """Allow users to download the converted Word file."""
    file_path = os.path.join(app.config["CONVERTED_FOLDER"], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        flash("File not found. Please convert a PDF first.", "error")
        return redirect(url_for("upload_file"))


if __name__ == "__main__":
    app.run(debug=True)
