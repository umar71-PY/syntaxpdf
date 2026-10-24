# SyntaxPDF - The Ultimate Online PDF Toolkit

SyntaxPDF is a fast, secure, and professional web application designed for editing and processing PDF files with ease. The project is currently live at [umerlabs.tech](https://umerlabs.tech). It features a comprehensive suite of 22 different PDF tools to streamline your daily document workflows.

## 🚀 Key Features

* **22 Powerful Tools:** Includes capabilities to Merge, Split, Compress, Crop, Watermark, Protect, and Unlock PDFs, all accessible from a clean frontend and processed via robust Django backend views.
* **Format Conversion:** Seamlessly convert PDFs to Word, Excel, PowerPoint, JPG, Text, and Speech (MP3). It also supports reverse conversions (Word, Excel, HTML, and CSV to PDF).
* **Sejda-Style Smart Editor:** Click directly on existing PDF text to edit it in place. The editor automatically matches font styles, sizes, text colors, and background colors using PyMuPDF (`fitz`).
* **Tesseract OCR:** Extract text from scanned documents and generate fully searchable PDFs utilizing the Tesseract OCR engine.
* **Privacy & Security (Zero-Retention):** Files are never stored permanently. Once a download is complete, `SelfCleaningFileResponse` automatically securely wipes the files from the server's temporary directories.
* **Clerk Authentication:** Guests can process up to 3 documents for free. For unlimited access, a secure Clerk-based authentication system is integrated, protecting backend API endpoints with the `@clerk_authenticated` decorator.

## 🛠️ Tech Stack

* **Frontend:** HTML5, Tailwind CSS, PDF.js, pdf-lib
* **Backend:** Python, Django, PyMuPDF, pypdf, Tesseract OCR, Ghostscript, ReportLab, pdfplumber
* **Authentication:** Clerk SDK 

## ⚙️ Local Setup Instructions

**1. Clone the Repository:**
```bash
git clone https://github.com/umar71-PY/syntaxpdf.git
cd syntaxpdf
```

**2. Create and Activate a Virtual Environment:**
```bash
python -m venv env
source env/bin/activate  # For Linux/macOS
# For Windows: env\Scripts\activate
```

**3. Install Python Dependencies:**
```bash
pip install -r requirements.txt
```

**4. Install System Dependencies (Ubuntu / Zorin OS):**
Because the backend relies on system-level tools for compression, OCR, and document conversion, you must install the following packages:
```bash
sudo apt update
sudo apt install tesseract-ocr ghostscript poppler-utils libreoffice
```

**5. Run the Django Development Server:**
```bash
python manage.py migrate
python manage.py runserver
```
Once the server is running, open `http://127.0.0.1:8000` in your browser to access the SyntaxPDF interface.
