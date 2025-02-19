# Word_to_PDF
Word file to PDF convertor
This is a simple web-based application that allows users to upload a Word (.docx) file and convert it to PDF using Python and Flask.

**Features**
1) Upload a .docx file
2) Convert it to PDF using Microsoft Word
3) Download the converted PDF file

**Download**
1) Create a folder on local drive e.g.- "Project".
2) Download i) main.py  ii) model.pkl in "Project" folder.
3) Create another folder "templates" inside "Project" folder.
4) Download the pdf_web.html file in  "templates" folder. 

**How to Run the Application**
1️⃣ Start the Flask Server
    Open Cammand Prompt from "Project" folder.
    Run code "python main.py"
    The application will start on http://127.0.0.1:5000/.

2️⃣ Open the Web Interface
    Open a browser and go to http://127.0.0.1:5000/.
    Click Browse, select a .docx file, and click Convert.
    The converted PDF file will be downloaded automatically.

3️⃣ File Structure

    Word-To-PDF-Converter/
    │-- main.py                # Main Flask Application
    │-- templates/
    │   ├── pdf_web.html        # Web Interface
    │-- static/
    │-- uploads/              # Uploaded Word Files
    │-- outputs/              # Converted PDF Files
    
