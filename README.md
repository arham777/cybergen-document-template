# CyberGen Document Formatter

A professional document formatting tool that automatically formats documents with consistent styling, detects headings, and handles dates.

## Features

- **Smart Formatting**: Automatic detection and formatting of headings, subheadings, and paragraphs
- **Date Handling**: Intelligent date detection and standardization to a professional format
- **Content Preservation**: Maintains tables, images, and complex formatting from source documents

## Project Structure

- `frontend/`: Next.js frontend application
- `api.py`: FastAPI backend for document processing
- `cybergen_template.py`: Core document formatting logic
- `cybergen-template.docx`: Template document for formatting

## Development Setup

### Prerequisites

- Python 3.8+
- Node.js 16+
- npm or yarn

### Installation

1. Install Python dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Install frontend dependencies:
   ```
   cd frontend
   npm install
   ```

### Running Locally

For Windows:
```
dev.bat
```

For Unix/Mac:
```
# Terminal 1
python api.py

# Terminal 2
cd frontend
npm run dev
```

The application will be available at:
- Frontend: http://localhost:3000
- Backend API: http://localhost:8000

## Deployment to Vercel

1. Install Vercel CLI:
   ```
   npm install -g vercel
   ```

2. Login to Vercel:
   ```
   vercel login
   ```

3. Deploy:
   ```
   vercel
   ```

4. Set environment variables in Vercel dashboard:
   - `NEXT_PUBLIC_API_URL`: Leave empty for production (API and frontend are on the same domain)

## API Endpoints

- `POST /api/generate-from-text`: Generate a document from text input
- `POST /api/generate-from-file`: Generate a document from an uploaded file (DOCX or PDF)

## License

MIT 