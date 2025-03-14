# CyberGen Document Formatter POC

This Proof of Concept (POC) demonstrates the core functionality of the CyberGen Document Formatter, focusing on essential features for document generation and formatting.

## Features Included in POC

1. **Template-Based Document Creation**
   - Uses a template file as the base for new documents
   - Applies consistent formatting

2. **Smart Heading Detection**
   - Identifies text likely to be headings based on patterns:
     - Text in ALL CAPITAL LETTERS
     - Text ending with a colon
     - Numbered or bulleted items
     - Short text (under 100 characters)

3. **Professional Formatting**
   - Headings: centered, bold, underlined, 14pt font size, 18pt spacing after
   - Normal paragraphs: justified, 12.5pt font size, 12pt spacing after

4. **Pagination Controls**
   - Headings stay with the following paragraph (no orphaned headings)
   - Widow/orphan control to prevent single lines at page boundaries

5. **Automatic Date Insertion**
   - Adds current date at the top right of the document
   - Format: "Nov 28, 2024"

## Requirements

- Python 3.7+
- python-docx library: `pip install python-docx`
- A template document (default: "cybergen-template.docx")

## Usage

1. Run the script:
   ```
   python cybergen_poc.py
   ```

2. Enter the path to your template (or press Enter for default)

3. Enter your document text, including some example headings:
   ```
   EXAMPLE HEADING

   This is a normal paragraph that will be formatted with justified text.
   It demonstrates the paragraph formatting functionality.

   Another heading with a colon:

   This text follows a heading with a colon and should stay with it
   across page breaks.

   1. Numbered heading

   Text under a numbered heading shows the heading detection feature.
   ```

4. Type `END` on a new line when finished

5. Enter a name for your output file (or press Enter for default)

6. The formatted document will be created and the path displayed

## Limitations in POC

- No PDF support (simplified for POC)
- Limited file format compatibility (Word documents only)
- Basic error handling
- No complex document structure support

## Next Steps

After validating this POC, the full implementation will include:
- PDF support
- Enhanced error handling
- More advanced formatting options
- Better document structure preservation

## Example Test Text

```
DOCUMENT TITLE

This is an example document to demonstrate the formatting capabilities
of the CyberGen Document Formatter POC.

First Section:

This text follows a heading with a colon and demonstrates how headings
are kept with their content across page breaks.

• Bulleted item as heading

This text demonstrates that bulleted items are detected as headings
and formatted accordingly.

Regular paragraph with more content. This shows the standard paragraph
formatting with justified text and appropriate spacing after paragraphs.
``` 