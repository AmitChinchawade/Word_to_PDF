import os
import pythoncom
from flask import Flask, request, render_template, send_file
from win32com import client

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def convert_to_pdf(word_file, output_file):
    """ Convert Word file to PDF using Microsoft Word. """
    pythoncom.CoInitialize()  # Ensure COM is initialized

    if not os.path.exists(word_file):
        return "File not found"

    try:
        word = client.Dispatch("Word.Application")
        word.Visible = False  # Run Word in the background
        doc = word.Documents.Open(os.path.abspath(word_file))  # Open file
        doc.SaveAs(os.path.abspath(output_file), FileFormat=17)  # Save as PDF
        doc.Close()
        word.Quit()
        return output_file  # Return converted PDF file path

    except Exception as e:
        return f"Error: {str(e)}"

    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM


@app.route("/")
def upload_form():
    """ Render the file upload form. """
    return render_template("pdf_web.html")


@app.route("/upload", methods=["POST"])
def upload_file():
    """ Handle file upload and conversion. """
    if "file" not in request.files:
        return "No file part"

    file = request.files["file"]
    if file.filename == "":
        return "No selected file"

    if file and file.filename.endswith(".docx"):
        word_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
        pdf_path = os.path.join(OUTPUT_FOLDER, file.filename.replace(".docx", ".pdf"))

        file.save(word_path)  # Save uploaded Word file

        # Convert Word to PDF
        result = convert_to_pdf(word_path, pdf_path)

        if os.path.exists(pdf_path):
            return send_file(pdf_path, as_attachment=True)  # Return converted PDF
        else:
            return f"Conversion failed: {result}"

    return "Invalid file format. Only .docx is allowed."


if __name__ == "__main__":
    app.run(debug=True, threaded=False)  # Run Flask in single-threaded mode
