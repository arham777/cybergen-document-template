# CyberGen Document Formatter - Streamlit App

This repository contains a professional document formatting application built with Streamlit. The app automatically detects headings, applies consistent formatting, and enhances document readability.

## Features

- **Smart Heading Detection**: Automatically identifies headings based on patterns like ALL CAPS, colons, bullets, or numbers
- **Professional Formatting**: Applies consistent styling with center alignment, bold, and underline for headings
- **Proper Spacing**: Adds appropriate spacing after paragraphs and headings
- **Pagination Controls**: Ensures headings stay with their following content
- **Real-time Date**: Adds the current date at the top right of each document
- **Multiple Input Options**: Support for direct text input, Word documents, and PDFs

## Live Demo

The application is deployed on Streamlit Cloud. You can access it here:
[CyberGen Document Formatter](https://cybergen-document-formatter.streamlit.app)

## Deployment Instructions

### 1. Clone this repository:
```bash
git clone <your-repository-url>
cd <repository-folder>
```

### 2. Create and activate a virtual environment (optional but recommended):
```bash
# On Windows
python -m venv venv
venv\Scripts\activate

# On macOS/Linux
python -m venv venv
source venv/bin/activate
```

### 3. Install the required packages:
```bash
pip install -r requirements.txt
```

### 4. Run the app locally:
```bash
streamlit run streamlit_app.py
```

### 5. Deploy to Streamlit Cloud:

1. Push your code to a GitHub repository
2. Sign up at [Streamlit Cloud](https://streamlit.io/cloud)
3. Click "New app"
4. Select your repository, branch, and the `streamlit_app.py` file
5. Click "Deploy"

## Files in this Repository

- `streamlit_app.py`: The main Streamlit application
- `cybergen_template.py`: Core document processing functionality
- `cybergen-template.docx`: Default template document
- `requirements.txt`: Required Python packages

## How to Use

1. Open the application in your web browser
2. Choose between entering text directly or uploading a document
3. For text input:
   - Type or paste your document content
   - Click "Generate Formatted Document"
4. For file upload:
   - Upload a Word document (.docx) or PDF file
   - Click "Generate Formatted Document"
5. Download the professionally formatted document

## Sample Text for Testing

```
DOCUMENT TITLE

This is an example document to demonstrate the formatting capabilities
of the CyberGen Document Formatter.

First Section:

This text follows a heading with a colon and demonstrates how headings
are kept with their content across page breaks.

• Bulleted item as heading

This text demonstrates that bulleted items are detected as headings
and formatted accordingly.

Regular paragraph with more content. This shows the standard paragraph
formatting with justified text and appropriate spacing after paragraphs.
```

## License

[MIT License](LICENSE) 