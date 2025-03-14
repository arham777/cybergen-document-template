import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
import os
from datetime import datetime
import PyPDF2  # For PDF text extraction

def extract_text_from_pdf(file_path):
    """
    Extract text content from a PDF file.
    
    Args:
        file_path (str): Path to the PDF file
        
    Returns:
        str: Extracted text content
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        if not file_path.lower().endswith('.pdf'):
            raise ValueError("File must be a PDF")
        
        text = ""
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text += page.extract_text() + "\n\n"
        
        return text.strip()
    
    except Exception as e:
        print(f"Error extracting text from PDF: {str(e)}")
        return None

def extract_text_from_docx(file_path):
    """
    Extract text content from a Word document.
    
    Args:
        file_path (str): Path to the Word document
        
    Returns:
        str: Extracted text content
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        if not file_path.lower().endswith(('.docx', '.doc')):
            raise ValueError("File must be a Word document (.doc or .docx)")
        
        doc = docx.Document(file_path)
        full_text = []
        
        for para in doc.paragraphs:
            if para.text.strip():  # Only add non-empty paragraphs
                full_text.append(para.text)
        
        return '\n'.join(full_text)
    
    except Exception as e:
        print(f"Error extracting text from document: {str(e)}")
        return None

def parse_document(file_path):
    """
    Parse an existing document and return its text content.
    
    Args:
        file_path (str): Path to the document file
    
    Returns:
        str: The text content of the document
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        # Check file extension to determine parsing method
        file_ext = os.path.splitext(file_path.lower())[1]
        
        if file_ext == '.pdf':
            # Parse PDF file
            return extract_text_from_pdf(file_path)
        elif file_ext in ('.docx', '.doc'):
            # Parse Word document
            return extract_text_from_docx(file_path)
        else:
            raise ValueError("File must be a Word document (.doc or .docx) or a PDF (.pdf)")
    
    except Exception as e:
        print(f"Error parsing document: {str(e)}")
        return None

def set_document_margins(doc, top=1.5, bottom=1.5, left=1.0, right=1.0):
    """
    Set margins for all sections in a document.
    
    Args:
        doc: The document to modify
        top: Top margin in inches
        bottom: Bottom margin in inches
        left: Left margin in inches
        right: Right margin in inches
    """
    for section in doc.sections:
        section.top_margin = Inches(top)
        section.bottom_margin = Inches(bottom)
        section.left_margin = Inches(left)
        section.right_margin = Inches(right)

def add_current_date(doc):
    """
    Add the current date at the top right of the document.
    
    Args:
        doc: The document to modify
    
    Returns:
        The added paragraph
    """
    # Format the current date as "Nov 28, 2024"
    current_date = datetime.now().strftime("%b %d, %Y")
    
    # Add a paragraph for the date at the top
    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    date_run = date_para.add_run(current_date)
    date_run.font.size = Pt(12)
    date_run.bold = True
    
    # Add space after paragraph using Word's standard formatting
    date_para.paragraph_format.space_after = Pt(12)
    
    return date_para

def is_heading(text):
    """
    Determine if the text is likely a heading.
    
    Args:
        text (str): The text to check
        
    Returns:
        bool: True if the text is likely a heading, False otherwise
    """
    if not text.strip():
        return False
    
    # Check if text is short (typical of headings)
    if len(text.strip()) < 100:
        # Check for patterns typical of headings
        if text.isupper() or text.endswith(':'):
            return True
        # Check if text is numbered or bulleted
        if (text.strip().startswith(('•', '-', '*')) or 
            (text[0].isdigit() and '.' in text[:3])):
            return True
    
    return False

def format_paragraph(paragraph, is_heading_text=False):
    """
    Apply formatting to a paragraph based on whether it's a heading or not.
    
    Args:
        paragraph: The paragraph to format
        is_heading_text: Whether the paragraph is a heading
        
    Returns:
        The formatted paragraph
    """
    if is_heading_text:
        # Heading formatting: center alignment, 14pt font size, bold, underlined
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Use Word's standard "Add Space After Paragraph" for headings
        paragraph.paragraph_format.space_after = Pt(18)
        # Keep headings with the following paragraph text to prevent orphaned headings
        paragraph.paragraph_format.keep_with_next = True
        for run in paragraph.runs:
            run.font.size = Pt(14)
            run.font.bold = True
            run.underline = WD_UNDERLINE.SINGLE
    else:
        # Regular paragraph formatting: justify alignment, 12.5pt font size
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # Use Word's standard "Add Space After Paragraph" for normal paragraphs
        paragraph.paragraph_format.space_after = Pt(12)
        for run in paragraph.runs:
            run.font.size = Pt(12.5)
    
    return paragraph

def add_space_after_paragraph(paragraph, is_heading=False):
    """
    Adds proper spacing after a paragraph according to Word standards.
    Also applies pagination controls to prevent orphaned headings.
    
    Args:
        paragraph: The paragraph to modify
        is_heading: Whether the paragraph is a heading
        
    Returns:
        The modified paragraph
    """
    # Use Word's standard spacing
    if is_heading:
        paragraph.paragraph_format.space_after = Pt(18)  # More space after headings (18pt)
        # Keep heading with next paragraph to prevent orphaned headings
        paragraph.paragraph_format.keep_with_next = True
    else:
        paragraph.paragraph_format.space_after = Pt(12)  # Standard space after paragraphs (12pt)
    
    return paragraph

def copy_document_to_template(source_file, template_path="cybergen-template.docx", output_filename="generated_document.docx"):
    """
    Copies content from a source document to a template, preserving formatting.
    
    Args:
        source_file (str): Path to the source document (Word or PDF)
        template_path (str): Path to the template document
        output_filename (str): Name for the output document
    
    Returns:
        str: Path to the created document
    """
    try:
        # Check if files exist
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")
        if not os.path.exists(source_file):
            raise FileNotFoundError(f"Source file not found: {source_file}")
        
        # Load the template document
        template_doc = docx.Document(template_path)
        
        # Set margins to ensure spacing on every page
        set_document_margins(template_doc, top=1.5, bottom=1.5)
        
        # Add current date to the first page
        add_current_date(template_doc)
        
        # Determine file type
        file_ext = os.path.splitext(source_file.lower())[1]
        
        if file_ext in ('.docx', '.doc'):
            # For Word documents, copy content preserving formatting
            source_doc = docx.Document(source_file)
            
            # Copy each paragraph
            for para in source_doc.paragraphs:
                if para.text.strip():  # Skip empty paragraphs
                    # Check if this paragraph is a heading
                    heading_status = is_heading(para.text)
                    is_bold = any(run.bold for run in para.runs) if para.runs else False
                    
                    # Add paragraph to template
                    new_para = template_doc.add_paragraph()
                    
                    # Copy text with formatting
                    for run in para.runs:
                        new_run = new_para.add_run(run.text)
                        # Copy run formatting
                        new_run.bold = run.bold if not heading_status else True
                        new_run.italic = run.italic
                        new_run.underline = run.underline if not heading_status else WD_UNDERLINE.SINGLE
                        # Set font size based on heading status
                        new_run.font.size = Pt(14) if heading_status else Pt(12.5)
                    
                    # Set alignment based on heading status
                    new_para.alignment = WD_ALIGN_PARAGRAPH.CENTER if heading_status else WD_ALIGN_PARAGRAPH.JUSTIFY
                    
                    # Add proper spacing after paragraph using Word's standard
                    add_space_after_paragraph(new_para, is_heading=heading_status)
                    
                    # If there are no runs (plain paragraph), add text with appropriate formatting
                    if not para.runs and para.text.strip():
                        new_run = new_para.add_run(para.text)
                        new_run.font.size = Pt(14) if heading_status else Pt(12.5)
                        new_run.bold = True if heading_status else False
                        new_run.underline = WD_UNDERLINE.SINGLE if heading_status else None
                    
        elif file_ext == '.pdf':
            # For PDFs, we extract text and maintain paragraph structure
            text_content = extract_text_from_pdf(source_file)
            if text_content:
                paragraphs = text_content.split('\n\n')
                last_para = None
                
                for paragraph in paragraphs:
                    if paragraph.strip():
                        # Check if this paragraph is a heading
                        heading_status = is_heading(paragraph)
                        
                        # Add paragraph with appropriate formatting
                        p = template_doc.add_paragraph()
                        run = p.add_run(paragraph)
                        
                        # Apply formatting based on heading status
                        run.font.size = Pt(14) if heading_status else Pt(12.5)
                        run.bold = True if heading_status else False
                        run.underline = WD_UNDERLINE.SINGLE if heading_status else None
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if heading_status else WD_ALIGN_PARAGRAPH.JUSTIFY
                        
                        # Add proper spacing after paragraph using Word's standard
                        add_space_after_paragraph(p, is_heading=heading_status)
                        
                        # Store this paragraph to check if it needs to be kept with the next
                        last_para = p
        
        # Set widow/orphan control for the whole document to prevent single lines
        for paragraph in template_doc.paragraphs:
            paragraph.paragraph_format.widow_control = True
        
        # Save the document
        template_doc.save(output_filename)
        return os.path.abspath(output_filename)
    
    except Exception as e:
        print(f"Error copying document: {str(e)}")
        return None

def insert_text_into_template(input_text, template_path="cybergen-template.docx", output_filename="generated_document.docx"):
    """
    Inserts the user's text into the template document.
    
    Args:
        input_text (str): The text content to be inserted
        template_path (str): Path to the template document
        output_filename (str): Name for the output document
    
    Returns:
        str: Path to the created document
    """
    try:
        # Check if template exists
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")
        
        # Load the template document
        doc = docx.Document(template_path)
        
        # Set margins to ensure spacing on every page
        set_document_margins(doc, top=1.5, bottom=1.5)
        
        # Add current date to the first page
        add_current_date(doc)
        
        # Process the input text
        paragraphs = input_text.strip().split('\n')
        
        # Simply append each paragraph to the document
        for i, paragraph in enumerate(paragraphs):
            if paragraph.strip():  # Skip empty paragraphs
                # Check if this paragraph is a heading
                heading_status = is_heading(paragraph)
                
                # Add paragraph with appropriate formatting
                p = doc.add_paragraph()
                if len(doc.paragraphs) > 1:
                    # Copy style from an existing paragraph if available
                    p.style = doc.paragraphs[1].style
                
                # Add run with appropriate formatting
                run = p.add_run(paragraph)
                run.font.size = Pt(14) if heading_status else Pt(12.5)
                run.bold = True if heading_status else False
                run.underline = WD_UNDERLINE.SINGLE if heading_status else None
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if heading_status else WD_ALIGN_PARAGRAPH.JUSTIFY
                
                # Add proper spacing after paragraph using Word's standard
                add_space_after_paragraph(p, is_heading=heading_status)
        
        # Set widow/orphan control for the whole document
        for paragraph in doc.paragraphs:
            paragraph.paragraph_format.widow_control = True
        
        # Save the document
        doc.save(output_filename)
        return os.path.abspath(output_filename)
    
    except Exception as e:
        print(f"Error creating document: {str(e)}")
        return None

def main():
    """
    Main function to handle user interaction and document processing.
    """
    print("CyberGen Document Formatter")
    print("==========================")
    
    # Default template path
    template_path = "cybergen-template.docx"
    
    while True:
        print("\nOptions:")
        print("1. Enter text directly")
        print("2. Import text from a document (Word or PDF)")
        print("3. Exit")
        
        choice = input("\nEnter your choice (1-3): ")
        
        if choice == '1':
            print("\nEnter your document text (type 'END' on a new line when finished):")
            print("Note: Text detected as headings will be automatically formatted with center alignment, bold, underline, and 14pt font.")
            print("      All paragraphs will have proper spacing after them.")
            print("      Headings will be kept with the following text to prevent page breaks between them.")
            lines = []
            while True:
                line = input()
                if line == 'END':
                    break
                lines.append(line)
            
            input_text = '\n'.join(lines)
            
            output_name = input("\nEnter output filename (leave blank for default 'generated_document.docx'): ")
            if not output_name:
                output_name = "generated_document.docx"
            elif not output_name.lower().endswith('.docx'):
                output_name += '.docx'
                
            document_path = insert_text_into_template(input_text, template_path=template_path, output_filename=output_name)
            if document_path:
                print(f"\nDocument successfully created at: {document_path}")
                print("Note: Text has been formatted according to heading detection rules.")
                print("      All paragraphs have standard spacing after them.")
                print("      Headings are kept with their following paragraphs across page breaks.")
        
        elif choice == '2':
            file_path = input("\nEnter the path to the document (Word or PDF): ")
            
            if not os.path.exists(file_path):
                print(f"Error: File not found at {file_path}")
                continue
                
            output_name = input("\nEnter output filename (leave blank for default 'generated_document.docx'): ")
            if not output_name:
                output_name = "generated_document.docx"
            elif not output_name.lower().endswith('.docx'):
                output_name += '.docx'
            
            # Use the new function to copy content preserving formatting
            document_path = copy_document_to_template(file_path, template_path=template_path, output_filename=output_name)
            if document_path:
                print(f"\nDocument successfully created at: {document_path}")
                print("Note: Text has been formatted according to heading detection rules.")
                print("      All paragraphs have standard spacing after them.")
                print("      Headings are kept with their following paragraphs across page breaks.")
        
        elif choice == '3':
            print("\nExiting program. Goodbye!")
            break
        
        else:
            print("\nInvalid choice. Please try again.")

if __name__ == "__main__":
    main() 