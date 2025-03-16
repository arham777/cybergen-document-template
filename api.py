from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional
import tempfile
import os
import base64
from cybergen_template import insert_text_into_template, copy_document_to_template

app = FastAPI()

# Add CORS middleware to allow requests from the frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods
    allow_headers=["*"],  # Allows all headers
)

# Get the template path - check both the current directory and the frontend/public directory
def get_template_path():
    template_paths = [
        "cybergen-template.docx",  # Current directory
        os.path.join("frontend", "public", "cybergen-template.docx"),  # Frontend public directory
    ]
    
    for path in template_paths:
        if os.path.exists(path):
            return path
    
    raise FileNotFoundError("Template file not found")

class TextInput(BaseModel):
    text: str
    filename: Optional[str] = "generated_document.docx"

@app.post("/api/generate-from-text")
async def generate_from_text(input_data: TextInput):
    try:
        # Create a temporary directory
        temp_dir = tempfile.mkdtemp()
        output_path = os.path.join(temp_dir, input_data.filename)
        template_path = get_template_path()
        
        # Generate the document
        document_path = insert_text_into_template(
            input_data.text,
            template_path=template_path,
            output_filename=output_path
        )
        
        if not document_path:
            raise HTTPException(status_code=500, detail="Failed to generate document")
        
        # Read the generated document and convert to base64
        with open(document_path, "rb") as f:
            document_bytes = f.read()
        
        base64_document = base64.b64encode(document_bytes).decode()
        
        # Clean up
        try:
            os.remove(document_path)
            os.rmdir(temp_dir)
        except:
            pass
        
        return {"success": True, "document": base64_document}
    
    except Exception as e:
        print(f"Error generating document: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/generate-from-file")
async def generate_from_file(
    file: UploadFile = File(...),
    filename: str = Form("generated_document.docx")
):
    try:
        # Create a temporary directory
        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, file.filename)
        output_path = os.path.join(temp_dir, filename)
        template_path = get_template_path()
        
        # Save the uploaded file
        with open(input_path, "wb") as f:
            content = await file.read()
            f.write(content)
        
        # Generate the document
        document_path = copy_document_to_template(
            input_path,
            template_path=template_path,
            output_filename=output_path
        )
        
        if not document_path:
            raise HTTPException(status_code=500, detail="Failed to generate document")
        
        # Read the generated document and convert to base64
        with open(document_path, "rb") as f:
            document_bytes = f.read()
        
        base64_document = base64.b64encode(document_bytes).decode()
        
        # Clean up
        try:
            os.remove(input_path)
            os.remove(document_path)
            os.rmdir(temp_dir)
        except:
            pass
        
        return {"success": True, "document": base64_document}
    
    except Exception as e:
        print(f"Error generating document: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000) 