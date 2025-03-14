# CyberGen Document Formatter

A Streamlit web application for generating properly formatted documents using the CyberGen template.

## Features

- Generate formatted documents with consistent styling
- Direct text input option
- Import existing Word (.docx/.doc) or PDF files
- Automatic heading detection and formatting
- Proper paragraph spacing and pagination
- Document download functionality

## Requirements

- Python 3.6+
- Streamlit
- python-docx
- PyPDF2
- base64
- tempfile
- os

## Installation

1. Clone the repository:
```
git clone <repository-url>
cd document-formatting
```

2. Install the required dependencies:
```
pip install -r requirements.txt
```

3. Make sure you have the template file `cybergen-template.docx` in the same directory as the app.py file.

## Usage

1. Run the Streamlit app:
```
streamlit run app.py
```

2. Open your browser and go to the URL displayed in the terminal (usually http://localhost:8501)

3. Use the sidebar to choose one of the following options:
   - **Enter Text Directly**: Type your document content in the text area
   - **Import Document**: Upload an existing Word or PDF document

4. Specify an output filename (optional)

5. Click "Generate Document" to process and create the formatted document

6. Download the generated document using the provided download link

## How It Works

The app uses the following components:

- `cybergen_template.py`: Contains the core document processing logic
- `app.py`: Streamlit interface for the document formatter
- `cybergen-template.docx`: Template file for document generation

The formatting follows these rules:
- Text detected as headings is formatted with center alignment, bold, underline, and 14pt font
- Regular paragraphs are formatted with justified alignment and 12.5pt font
- All paragraphs have proper spacing after them
- Headings are kept with the following text to prevent page breaks between them

## Notes

- The template file `cybergen-template.docx` must be present in the same directory as the app
- For PDF imports, text extraction may not preserve all formatting from the original document
- Word documents (.docx/.doc) will better preserve original formatting during import 