"""
FastAPI Web Server for PDF/PPTX Speaker Notes Generator
Upload PDF or PPTX files and download them with AI-generated speaker notes.
"""

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import os
import shutil
from pathlib import Path
import uuid
import asyncio
import json
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

# Store mapping of processing_id to original filename
filename_mapping = {}


@app.get("/", response_class=HTMLResponse)
async def home():
    """Serve the main HTML page."""
    with open("index.html", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Handle file upload and return a processing ID."""
    
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
    
    # Store the mapping for later retrieval
    filename_mapping[unique_id] = output_filename
    
    # Return the processing ID for SSE endpoint
    return {
        "processing_id": unique_id,
        "file_type": file_ext,
        "original_name": Path(file.filename).stem + "_with_notes.pptx",
        "output_filename": output_filename
    }


@app.get("/process/{processing_id}")
async def process_file(processing_id: str, style: str = "standard", tone: str = "professional"):
    """Stream processing progress using Server-Sent Events."""
    
    async def event_generator():
        try:
            # Find the input file
            input_files = list(UPLOAD_DIR.glob(f"{processing_id}.*"))
            if not input_files:
                yield f"data: {json.dumps({'error': 'File not found'})}\n\n"
                return
            
            input_path = input_files[0]
            file_ext = input_path.suffix.lower()
            
            # Get original filename from mapping, fallback to UUID if not found
            output_filename = filename_mapping.get(processing_id, f"{processing_id}_with_notes.pptx")
            output_path = OUTPUT_DIR / output_filename
            
            # Get API key
            api_key = os.environ.get('GOOGLE_API_KEY')
            
            # Process the file with progress updates
            try:
                if file_ext == '.pdf':
                    # Process PDF
                    async for progress in pdf_to_pptx_with_notes_streaming(
                        str(input_path), str(output_path), dpi=200, api_key=api_key,
                        note_style=style, note_tone=tone
                    ):
                        yield f"data: {json.dumps(progress)}\n\n"
                        await asyncio.sleep(0.1)  # Small delay for streaming
                else:
                    # Process PPTX
                    async for progress in add_notes_to_pptx_streaming(
                        str(input_path), str(output_path), api_key,
                        note_style=style, note_tone=tone
                    ):
                        yield f"data: {json.dumps(progress)}\n\n"
                        await asyncio.sleep(0.1)
                
                # Clean up input file
                if input_path.exists():
                    os.remove(input_path)
                
                # Clean up the filename mapping
                if processing_id in filename_mapping:
                    del filename_mapping[processing_id]
                
                # Send completion
                yield f"data: {json.dumps({'status': 'complete', 'filename': output_path.name})}\n\n"
                
            except Exception as e:
                # Clean up on error
                if input_path.exists():
                    os.remove(input_path)
                if output_path.exists():
                    os.remove(output_path)
                # Clean up the filename mapping
                if processing_id in filename_mapping:
                    del filename_mapping[processing_id]
                yield f"data: {json.dumps({'error': str(e)})}\n\n"
        
        except Exception as e:
            yield f"data: {json.dumps({'error': str(e)})}\n\n"
    
    return StreamingResponse(event_generator(), media_type="text/event-stream")


async def pdf_to_pptx_with_notes_streaming(pdf_path, output_pptx, dpi, api_key, note_style="standard", note_tone="professional"):
    """Process PDF with progress streaming."""
    from add_speaker_notes import process_pdf_with_progress
    
    for progress in process_pdf_with_progress(pdf_path, output_pptx, dpi, api_key, note_style, note_tone):
        yield progress
        await asyncio.sleep(0)


async def add_notes_to_pptx_streaming(input_pptx, output_pptx, api_key, note_style="standard", note_tone="professional"):
    """Process PPTX with progress streaming."""
    from add_speaker_notes import process_pptx_with_progress
    
    for progress in process_pptx_with_progress(input_pptx, output_pptx, api_key, note_style, note_tone):
        yield progress
        await asyncio.sleep(0)


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
