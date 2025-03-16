import streamlit as st
import os
import tempfile
import base64
from datetime import datetime
from cybergen_template import insert_text_into_template, copy_document_to_template, parse_document

# Set page configuration
st.set_page_config(
    page_title="CyberGen Document Formatter",
    page_icon="��",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #4CAF50;
        color: white;
        padding: 0.5rem;
        border-radius: 5px;
        border: none;
        margin-top: 1rem;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .download-btn {
        display: inline-block;
        padding: 0.5rem 1rem;
        background-color: #008CBA;
        color: white;
        text-decoration: none;
        border-radius: 5px;
        margin-top: 1rem;
        text-align: center;
    }
    .download-btn:hover {
        background-color: #007B9E;
    }
    .stTextArea>div>div>textarea {
        background-color: #f8f9fa;
    }
    .sidebar .sidebar-content {
        background-color: #f8f9fa;
    }
    h1 {
        color: #2C3E50;
        margin-bottom: 2rem;
    }
    h2 {
        color: #34495E;
        margin: 1.5rem 0;
    }
    .info-box {
        background-color: #E3F2FD;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #E8F5E9;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #FFF3E0;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #FFEBEE;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Function to create a download link for a file
def get_download_link(file_path, link_text="Download Document"):
    with open(file_path, "rb") as file:
        data = file.read()
    b64 = base64.b64encode(data).decode()
    href = f'<a class="download-btn" href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{os.path.basename(file_path)}">{link_text} 📥</a>'
    return href

def main():
    # Create a two-column layout for the header
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.title("📝 CyberGen Document Formatter")
    with col2:
        st.markdown("""
        <div style='text-align: right; padding-top: 1rem;'>
            <span style='color: #666; font-size: 0.8rem;'>Version 1.0</span>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<hr style='margin: 2rem 0;'>", unsafe_allow_html=True)
    
    # Sidebar with improved styling
    with st.sidebar:
        st.markdown("### 🛠️ Options")
        option = st.radio(
            "Choose your input method:",
            ["Enter Text Directly", "Import Document"],
            format_func=lambda x: "✏️ " + x if x == "Enter Text Directly" else "📁 " + x
        )
        
        st.markdown("---")
        st.markdown("### ℹ️ About")
        st.markdown("""
        CyberGen Document Formatter helps you create professionally formatted documents with consistent styling and layout.
        
        **Features:**
        - Automatic heading detection
        - Smart date formatting
        - Table and image preservation
        - Professional spacing and alignment
        """)
    
    # Default template path
    template_path = "cybergen-template.docx"
    
    # Check if template exists
    if not os.path.exists(template_path):
        st.error("⚠️ Template file not found!")
        st.info("💡 Please make sure 'cybergen-template.docx' is in the same directory as this app.")
        return
    
    if option == "Enter Text Directly":
        st.header("✏️ Enter Document Text")
        
        # Create tabs for input and formatting guide
        tab1, tab2 = st.tabs(["📝 Input", "📖 Formatting Guide"])
        
        with tab1:
            # Text input area with improved styling
            user_text = st.text_area(
                "Enter your document text here:",
                height=300,
                placeholder="Start typing your document content here..."
            )
            
            col1, col2 = st.columns([2, 1])
            with col1:
                output_filename = st.text_input(
                    "Output filename:",
                    placeholder="generated_document.docx",
                    help="Leave blank for default name"
                )
            
            if not output_filename:
                output_filename = "generated_document.docx"
            elif not output_filename.lower().endswith('.docx'):
                output_filename += '.docx'
            
            # Process button
            if st.button("🔄 Generate Document"):
                if user_text:
                    with st.spinner("🔄 Processing your document..."):
                        temp_dir = tempfile.mkdtemp()
                        output_path = os.path.join(temp_dir, output_filename)
                        document_path = insert_text_into_template(user_text, template_path=template_path, output_filename=output_path)
                        
                        if document_path:
                            st.markdown('<div class="success-box">✅ Document successfully created!</div>', unsafe_allow_html=True)
                            st.markdown(get_download_link(document_path, "📥 Download Your Document"), unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="error-box">❌ Error creating document. Please try again.</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="warning-box">⚠️ Please enter some text first.</div>', unsafe_allow_html=True)
        
        with tab2:
            st.markdown("""
            ### 📋 Formatting Guidelines
            
            Your text will be automatically formatted according to these rules:
            
            #### 📌 Headings
            - Center aligned
            - Bold and underlined
            - 14pt font size
            
            #### 📌 Subheadings
            - Center aligned
            - Bold and underlined
            - 13pt font size
            
            #### 📌 Regular Text
            - Left aligned
            - 12.5pt font size
            - Pure black color
            
            #### 📅 Date Handling
            - Dates are automatically detected
            - Formatted as "MMM DD, YYYY"
            - Placed at top right
            """)
    
    elif option == "Import Document":
        st.header("📁 Import Document")
        
        # Create tabs for upload and guide
        tab1, tab2 = st.tabs(["📤 Upload", "📖 Import Guide"])
        
        with tab1:
            # File uploader with improved styling
            uploaded_file = st.file_uploader(
                "Choose a Word or PDF document",
                type=["docx", "doc", "pdf"],
                help="Supported formats: .docx, .doc, .pdf"
            )
            
            if uploaded_file is not None:
                st.markdown(f"""
                <div class="info-box">
                    📄 <b>File Details</b><br>
                    Name: {uploaded_file.name}<br>
                    Size: {uploaded_file.size/1024:.1f} KB
                </div>
                """, unsafe_allow_html=True)
                
                # Save the uploaded file
                temp_dir = tempfile.mkdtemp()
                temp_file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                col1, col2 = st.columns([2, 1])
                with col1:
                    output_filename = st.text_input(
                        "Output filename:",
                        placeholder="generated_document.docx",
                        help="Leave blank for default name"
                    )
                
                if not output_filename:
                    output_filename = "generated_document.docx"
                elif not output_filename.lower().endswith('.docx'):
                    output_filename += '.docx'
                
                # Preview option
                with st.expander("👀 Preview Document Content"):
                    document_text = parse_document(temp_file_path)
                    if document_text:
                        st.text_area("Content preview:", document_text, height=200, disabled=True)
                    else:
                        st.warning("⚠️ Could not extract text preview from the document.")
                
                # Process button
                if st.button("🔄 Generate Document"):
                    with st.spinner("🔄 Processing your document..."):
                        output_path = os.path.join(temp_dir, output_filename)
                        document_path = copy_document_to_template(temp_file_path, template_path=template_path, output_filename=output_path)
                        
                        if document_path:
                            st.markdown('<div class="success-box">✅ Document successfully created!</div>', unsafe_allow_html=True)
                            st.markdown(get_download_link(document_path, "📥 Download Your Document"), unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="error-box">❌ Error creating document. Please try again.</div>', unsafe_allow_html=True)
        
        with tab2:
            st.markdown("""
            ### 📋 Import Guidelines
            
            #### 📄 Supported Formats
            - Microsoft Word (.docx, .doc)
            - PDF (.pdf)
            
            #### 🔄 Processing Features
            - **Tables**: Preserved from Word documents
            - **Images**: Preserved from Word documents
            - **Formatting**: Converted to consistent styling
            - **Dates**: Automatically detected and reformatted
            
            #### ⚠️ PDF Limitations
            - Tables may not be preserved
            - Images may not be preserved
            - Complex formatting may be simplified
            """)

if __name__ == "__main__":
    main() 