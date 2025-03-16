# CyberGen Document Formatter

A Streamlit web application for generating properly formatted documents using the CyberGen template.

## Features

- Generate formatted documents with consistent styling
- Direct text input option
- Import existing Word (.docx/.doc) or PDF files
- Automatic heading and subheading detection and formatting
- Proper paragraph spacing and pagination
- Intelligent date detection and formatting
- Preservation of tables and images from Word documents
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

### Formatting Rules

The document formatting follows these rules:

#### Text Formatting
- Main headings are formatted with center alignment, bold, underline, and 14pt font
- Subheadings are formatted with center alignment, bold, underline, and 13pt font (1pt smaller than main headings)
- Regular paragraphs are formatted with left alignment (not justified) and 12.5pt font in pure black color (#000000)
- All paragraphs have proper spacing after them
- Headings and subheadings are kept with the following text to prevent page breaks between them

#### Heading & Subheading Detection
The system intelligently identifies:
- **Main headings** based on text characteristics (all caps, ending with colon, numbered, etc.)
- **Subheadings** based on:
  - Position (appearing shortly after a main heading)
  - Format (indented or using hierarchical numbering like "1.1", "1.2", etc.)
  - Starting with common subheading prefixes like "subsection", "part", etc.
  - Using bullet points or specific formatting

#### Date Handling
- The system detects dates in various formats including:
  - DD/MM/YY (e.g., "16/06/24")
  - DD/MM/YYYY (e.g., "16/06/2024")
  - "Date: DD/MM/YY" or "Date: DD/MM/YYYY"
  - Month DD, YYYY (e.g., "June 16, 2024" or "Jun 16, 2024")
  - DD Month YYYY (e.g., "16 June 2024" or "16 Jun 2024")
  - YYYY-MM-DD (e.g., "2024-06-16")
- When a date is detected, it is:
  1. Extracted from the document
  2. Formatted to "MMM DD, YYYY" format (e.g., "Jun 16, 2024")
  3. Placed at the top right of the first page
  4. The original date line is completely removed from the document content
- If no date is found in the document, no date will be added

#### Tables and Images
- Tables from Word documents (.docx/.doc) are preserved in the formatted document:
  - Table structure (rows and columns) is maintained
  - Cell content is formatted with the same rules as regular text
  - Table styles are preserved when possible
- Images from Word documents are also preserved in the generated document
- For PDF files, tables and images cannot be preserved in the same way due to format limitations

## Notes

- The template file `cybergen-template.docx` must be present in the same directory as the app
- For PDF imports, text extraction may not preserve all formatting from the original document
- Word documents (.docx/.doc) will better preserve original formatting, tables, and images during import 