import streamlit as st
import os
import tempfile
import base64
from datetime import datetime
from cybergen_template import insert_text_into_template, copy_document_to_template, parse_document

# Set page configuration
st.set_page_config(
    page_title="Cybergen-Doc",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for modern styling
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    /* Main container styling */
    .main {
        background-color: white;
        padding: 0;
    }
    
    /* Header styling */
    .header-container {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 1.5rem 0;
        border-bottom: 1px solid #f0f0f0;
        margin-bottom: 2rem;
    }
    
    .app-title {
        font-size: 1.6rem;
        font-weight: 600;
        color: #0066cc;
        margin: 0;
    }
    
    .version-badge {
        font-size: 0.7rem;
        color: #0066cc;
    }
    
    /* Animation for text typing effect */
    @keyframes typing {
        from { width: 0 }
        to { width: 100% }
    }
    
    .typing-animation {
        overflow: hidden;
        white-space: nowrap;
        margin: 0 auto;
        letter-spacing: .15em;
        animation: typing 3.5s steps(40, end);
    }
    
    /* Button styling */
    .stButton>button {
        width: 100%;
        background-color: #0066cc;
        color: white;
        padding: 0.75rem;
        border-radius: 6px;
        border: none;
        margin-top: 1rem;
        font-weight: 500;
        transition: all 0.3s ease;
    }
    
    .stButton>button:hover {
        background-color: #004c99;
        transform: translateY(-2px);
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    .stButton>button:active {
        transform: translateY(0);
    }
    
    /* Download button */
    .download-btn {
        display: inline-block;
        padding: 0.75rem 1.5rem;
        background-color: #0066cc;
        color: white !important;
        text-decoration: none;
        border-radius: 6px;
        margin-top: 1rem;
        text-align: center;
        font-weight: 500;
        transition: all 0.3s ease;
    }
    
    .download-btn:hover {
        background-color: #004c99;
        transform: translateY(-2px);
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        color: white !important;
    }
    
    /* Input field styling */
    .stTextInput>div>div>input {
        border-radius: 6px;
        border: 1px solid #e0e0e0;
        padding: 0.75rem;
    }
    
    .stTextArea>div>div>textarea {
        border-radius: 6px;
        border: 1px solid #e0e0e0;
        background-color: white;
        padding: 0.75rem;
    }
    
    /* File uploader */
    .stFileUploader>div>button {
        background-color: transparent;
        color: #0066cc;
        border: 1px dashed #0066cc;
        border-radius: 6px;
    }
    
    /* Status boxes - all using blue tones */
    .success-box {
        background-color: #f0f7ff;
        padding: 1rem;
        border-radius: 6px;
        margin: 1.5rem 0;
        border-left: 4px solid #0066cc;
        display: flex;
        align-items: center;
        animation: fadeIn 0.5s ease-in;
    }
    
    .error-box {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 6px;
        margin: 1.5rem 0;
        border-left: 4px solid #0066cc;
        display: flex;
        align-items: center;
        animation: fadeIn 0.5s ease-in;
    }
    
    .warning-box {
        background-color: #f0f7ff;
        padding: 1rem;
        border-radius: 6px;
        margin: 1.5rem 0;
        border-left: 4px solid #004c99;
        display: flex;
        align-items: center;
        animation: fadeIn 0.5s ease-in;
    }
    
    .info-box {
        background-color: #f0f7ff;
        padding: 1rem;
        border-radius: 6px;
        margin: 1.5rem 0;
        border-left: 4px solid #0066cc;
        animation: fadeIn 0.5s ease-in;
    }
    
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    /* Feature card */
    .feature-card {
        background-color: #f0f7ff;
        border-radius: 8px;
        padding: 1.5rem;
        
        height: 220px;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        border: 1px solid rgba(0, 102, 204, 0.1);
        margin-bottom: 4rem;
        display: flex;
        flex-direction: column;
    }
    
    .feature-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 10px 20px rgba(0, 102, 204, 0.1);
        background-color: #e6f0ff;
    }
    
    .feature-title {
        color: #0066cc;
        font-weight: 600;
        margin-bottom: 0.8rem;
        font-size: 1.2rem;
    }
    
    .feature-card p {
        margin: 0;
        color: #333;
    }
    
    /* Remove default Streamlit padding and margins */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 0;
        max-width: 90%;
    }
    
    /* Hide hamburger menu */
    #MainMenu {visibility: hidden;}
    
    /* Hide footer */
    footer {visibility: hidden;}
    
    /* Expander styling */
    .streamlit-expanderHeader {
        font-weight: 500;
        color: #0066cc;
    }
    
    /* Dropdown styling */
    div[data-baseweb="select"] {
        border-radius: 6px;
    }
    
    
    /* Animations */
    .fade-in {
        animation: fadeIn 0.5s ease;
    }
    
    /* Radio buttons */
    .stRadio > div {
        display: flex;
        gap: 20px;
    }
    
    .stRadio [data-testid="stRadio"] > div {
        flex-direction: row;
    }
    
    /* Section headers */
    .section-header {
        color: #333;
        font-weight: 600;
        margin-bottom: 1.5rem;
    }
    
</style>
""", unsafe_allow_html=True)

# Function to create a download link for a file
def get_download_link(file_path, link_text="Download Document"):
    with open(file_path, "rb") as file:
        data = file.read()
    b64 = base64.b64encode(data).decode()
    href = f'<a class="download-btn" href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{os.path.basename(file_path)}" style="color: white !important;">{link_text}</a>'
    return href

def main():
    # Modern header
    st.markdown("""
    <div class="header-container">
        <h1 class="app-title">Cybergen-Doc</h1>
        <span class="version-badge">Version 1.0</span>
    </div>
    """, unsafe_allow_html=True)
    
    # Hero section with typing animation effect
    st.markdown("""
    <div style="text-align: center; margin-bottom: 3rem;">
        <p class="typing-animation" style="font-size: 1.6rem; color: #333; font-weight: 500;">
            Transform your documents with professional formatting
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Options in a horizontal layout instead of sidebar
    col1, col2 = st.columns([1, 3])
    
    with col1:
        st.markdown('<p class="section-header">Document Format</p>', unsafe_allow_html=True)
        option = st.radio(
            "",
            ["Import Document", "Enter Text Directly"],
            index=0
        )
    
    with col2:
        # Default template path
        template_path = "cybergen-template.docx"
        
        # Check if template exists
        if not os.path.exists(template_path):
            st.markdown(
                '<div class="error-box">⚠️ Template file not found! Please make sure \'cybergen-template.docx\' is in the same directory as this app.</div>',
                unsafe_allow_html=True
            )
            return
        
        if option == "Enter Text Directly":
            # Text tab
            st.markdown('<p class="section-header">✏️ Enter Document Text</p>', unsafe_allow_html=True)
            
            # Text input area with improved styling
            user_text = st.text_area(
                "",
                height=300,
                placeholder="Start typing your document content here..."
            )
            
            col1, col2 = st.columns([2, 1])
            with col1:
                output_filename = st.text_input(
                    "Output filename:",
                    placeholder="generated_document.docx"
                )
            
            if not output_filename:
                output_filename = "generated_document.docx"
            elif not output_filename.lower().endswith('.docx'):
                output_filename += '.docx'
            
            # Process button
            if st.button("Generate Document"):
                if user_text:
                    with st.spinner("Processing your document..."):
                        temp_dir = tempfile.mkdtemp()
                        output_path = os.path.join(temp_dir, output_filename)
                        document_path = insert_text_into_template(user_text, template_path=template_path, output_filename=output_path)
                        
                        if document_path:
                            st.markdown('<div class="success-box">✅ Document successfully created!</div>', unsafe_allow_html=True)
                            st.markdown(get_download_link(document_path, "Download Your Document"), unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="error-box">❌ Error creating document. Please try again.</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="warning-box">⚠️ Please enter some text first.</div>', unsafe_allow_html=True)
            
            # Formatting guidelines in an expander
            with st.expander("📋 Formatting Guidelines"):
                st.markdown("""
                Your text will be automatically formatted according to these rules:
                
                - **Headings**: Center aligned, bold, underlined, 14pt font
                - **Subheadings**: Center aligned, bold, underlined, 13pt font
                - **Regular Text**: Left aligned, 12.5pt font
                - **Dates**: Automatically detected and formatted as "MMM DD, YYYY"
                """)
        
        elif option == "Import Document":
            # Import tab
            st.markdown('<p class="section-header">📁 Import Document</p>', unsafe_allow_html=True)
            
            # File uploader with custom styling
            st.markdown('<div class="upload-dropzone">', unsafe_allow_html=True)
            uploaded_file = st.file_uploader(
                "",
                type=["docx", "doc", "pdf"],
                label_visibility="collapsed"
            )
            st.markdown('<div style="margin-top: 10px; font-size: 0.8rem;">Supported formats: .docx, .doc, .pdf</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
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
                        placeholder="generated_document.docx"
                    )
                
                if not output_filename:
                    output_filename = "generated_document.docx"
                elif not output_filename.lower().endswith('.docx'):
                    output_filename += '.docx'
                
                # Preview option
                with st.expander("👀 Preview Document Content"):
                    document_text = parse_document(temp_file_path)
                    if document_text:
                        st.text_area("", document_text, height=200, disabled=True)
                    else:
                        st.markdown('<div class="warning-box">⚠️ Could not extract text preview from the document.</div>', unsafe_allow_html=True)
                
                # Process button
                if st.button("Generate Document"):
                    with st.spinner("Processing your document..."):
                        output_path = os.path.join(temp_dir, output_filename)
                        document_path = copy_document_to_template(temp_file_path, template_path=template_path, output_filename=output_path)
                        
                        if document_path:
                            st.markdown('<div class="success-box">✅ Document successfully created!</div>', unsafe_allow_html=True)
                            st.markdown(get_download_link(document_path, "Download Your Document"), unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="error-box">❌ Error creating document. Please try again.</div>', unsafe_allow_html=True)
            
            # Import guidelines in an expander
            with st.expander("📋 Import Guidelines"):
                st.markdown("""
                - **Formats**: Microsoft Word (.docx, .doc) and PDF (.pdf)
                - **Features**: Tables and images are preserved from Word documents
                - **Formatting**: All content is converted to consistent styling
                - **Dates**: Automatically detected and standardized
                
                > PDF files have limitations: tables, images, and complex formatting may not be fully preserved
                """)
    
    # Features section
    st.markdown('<p class="section-header" style="margin-top: 2rem;">Key Features</p>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div class="feature-card">
            <h4 class="feature-title">🎨 Smart Formatting</h4>
            <p>Automatic detection of headings, subheadings, and paragraphs with professional styling</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="feature-card">
            <h4 class="feature-title">📅 Date Detection</h4>
            <p>Intelligent date detection and standardization to a consistent professional format</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="feature-card">
            <h4 class="feature-title">📊 Content Preservation</h4>
            <p>Maintains tables, images, and complex formatting from source documents</p>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main() 