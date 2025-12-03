"""
FastAPI Web Server for PDF/PPTX Speaker Notes Generator
Upload PDF or PPTX files and download them with AI-generated speaker notes.
"""

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import os
import shutil
from pathlib import Path
import uuid
from add_speaker_notes import pdf_to_pptx_with_notes, add_notes_to_pptx
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

app = FastAPI(title="Speaker Notes Generator")

# Create directories for uploads and outputs
UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


@app.get("/", response_class=HTMLResponse)
async def home():
    """Serve the main HTML page."""
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Handle file upload and process it."""
    
    # Validate file type
    file_ext = Path(file.filename).suffix.lower()
    if file_ext not in ['.pdf', '.pptx']:
        raise HTTPException(status_code=400, detail="Only PDF and PPTX files are supported")
    
    # Get API key
    api_key = os.environ.get('GOOGLE_API_KEY')
    if not api_key:
        raise HTTPException(status_code=500, detail="GOOGLE_API_KEY not configured")
    
    # Generate unique filename
    unique_id = str(uuid.uuid4())
    input_filename = f"{unique_id}{file_ext}"
    input_path = UPLOAD_DIR / input_filename
    
    # Save uploaded file
    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save file: {str(e)}")
    
    # Generate output filename using original file name
    original_stem = Path(file.filename).stem
    output_filename = f"{original_stem}_with_notes.pptx"
    output_path = OUTPUT_DIR / output_filename
    
    # Process the file
    try:
        if file_ext == '.pdf':
            pdf_to_pptx_with_notes(str(input_path), str(output_path), dpi=200, api_key=api_key)
        else:  # .pptx
            add_notes_to_pptx(str(input_path), str(output_path), api_key)
        
        # Clean up input file
        os.remove(input_path)
        
        # Return success with download link
        original_name = Path(file.filename).stem + "_with_notes.pptx"
        return {
            "success": True,
            "filename": output_filename,
            "original_name": original_name
        }
        
    except Exception as e:
        # Clean up on error
        if input_path.exists():
            os.remove(input_path)
        if output_path.exists():
            os.remove(output_path)
        
        raise HTTPException(status_code=500, detail=f"Processing failed: {str(e)}")


@app.get("/download/{filename}")
async def download_file(filename: str):
    """Download the processed file."""
    
    file_path = OUTPUT_DIR / filename
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    
    # Get original name from filename parameter in request
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )


@app.on_event("startup")
async def startup_event():
    """Check if API key is configured on startup."""
    api_key = os.environ.get('GOOGLE_API_KEY')
    if not api_key:
        print("⚠️  WARNING: GOOGLE_API_KEY not found in environment variables!")
        print("   Set it in the .env file before using the application.")
    else:
        print("✓ GOOGLE_API_KEY configured")
    
    print("\n" + "="*60)
    print("🎤 Speaker Notes Generator Server")
    print("="*60)
    print("\n🌐 Open your browser and go to: http://localhost:8000")
    print("\n✨ Upload a PDF or PPTX to add AI-generated speaker notes!\n")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
