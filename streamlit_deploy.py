import streamlit as st
import os
import tempfile
import base64
from datetime import datetime
from cybergen_template import insert_text_into_template, copy_document_to_template, parse_document

# Set page configuration
st.set_page_config(
    page_title="CyberGen Document Formatter",
    page_icon="📝",
    layout="wide"
)

# Function to create a download link for a file
def get_download_link(file_path, link_text="Download Document"):
    with open(file_path, "rb") as file:
        data = file.read()
    b64 = base64.b64encode(data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{os.path.basename(file_path)}">{link_text}</a>'
    return href

def main():
    # Header
    st.title("CyberGen Document Formatter")
    st.markdown("---")
    
    # Sidebar navigation
    option = st.sidebar.radio(
        "Choose an option:",
        ["Enter Text Directly", "Import Document"]
    )
    
    # Default template path
    template_path = "cybergen-template.docx"
    
    # Check if template exists
    if not os.path.exists(template_path):
        st.error(f"Template file not found: {template_path}")
        st.info("Please make sure 'cybergen-template.docx' is in the same directory as this app.")
        return
    
    if option == "Enter Text Directly":
        st.header("Enter Document Text")
        
        # Instructions
        st.info("""
        Note: 
        - Text detected as headings will be automatically formatted with center alignment, bold, underline, and 14pt font.
        - All paragraphs will have proper spacing after them.
        - Headings will be kept with the following text to prevent page breaks between them.
        """)
        
        # Text input area
        user_text = st.text_area("Enter your document text here:", height=300)
        
        # Output filename
        output_filename = st.text_input("Output filename (leave blank for default):")
        if not output_filename:
            output_filename = "generated_document.docx"
        elif not output_filename.lower().endswith('.docx'):
            output_filename += '.docx'
        
        # Process button
        if st.button("Generate Document"):
            if user_text:
                with st.spinner("Generating document..."):
                    # Create a temporary file
                    temp_dir = tempfile.mkdtemp()
                    output_path = os.path.join(temp_dir, output_filename)
                    
                    # Use the function from cybergen_template.py
                    document_path = insert_text_into_template(user_text, template_path=template_path, output_filename=output_path)
                    
                    if document_path:
                        st.success(f"Document successfully created!")
                        st.markdown(get_download_link(document_path, "Download Document"), unsafe_allow_html=True)
                        st.info("""
                        Note: 
                        - Text has been formatted according to heading detection rules.
                        - All paragraphs have standard spacing after them.
                        - Headings are kept with their following paragraphs across page breaks.
                        """)
                    else:
                        st.error("Error creating document. Please try again.")
            else:
                st.warning("Please enter some text first.")
    
    elif option == "Import Document":
        st.header("Import Document")
        
        # File uploader
        uploaded_file = st.file_uploader("Choose a Word or PDF document", type=["docx", "doc", "pdf"])
        
        if uploaded_file is not None:
            # Display file details
            file_details = {"Filename": uploaded_file.name, "File size": f"{uploaded_file.size} bytes"}
            st.write(file_details)
            
            # Save the uploaded file temporarily
            temp_dir = tempfile.mkdtemp()
            temp_file_path = os.path.join(temp_dir, uploaded_file.name)
            with open(temp_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Output filename
            output_filename = st.text_input("Output filename (leave blank for default):")
            if not output_filename:
                output_filename = "generated_document.docx"
            elif not output_filename.lower().endswith('.docx'):
                output_filename += '.docx'
            
            # Optional: Show preview of document content
            if st.checkbox("Show document content preview"):
                document_text = parse_document(temp_file_path)
                if document_text:
                    st.text_area("Document content:", document_text, height=200, disabled=True)
                else:
                    st.warning("Could not extract text from the document.")
            
            # Process button
            if st.button("Generate Document"):
                with st.spinner("Processing document..."):
                    output_path = os.path.join(temp_dir, output_filename)
                    
                    # Use the function from cybergen_template.py
                    document_path = copy_document_to_template(temp_file_path, template_path=template_path, output_filename=output_path)
                    
                    if document_path:
                        st.success(f"Document successfully created!")
                        st.markdown(get_download_link(document_path, "Download Document"), unsafe_allow_html=True)
                        st.info("""
                        Note: 
                        - Text has been formatted according to heading detection rules.
                        - All paragraphs have standard spacing after them.
                        - Headings are kept with their following paragraphs across page breaks.
                        """)
                    else:
                        st.error("Error creating document. Please try again.")

if __name__ == "__main__":
    main() 