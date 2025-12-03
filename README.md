# 🎤 AI Speaker Notes Generator

Automatically add AI-generated speaker notes to your PDF or PowerPoint presentations using Google Gemini AI. Perfect for presenters who need natural, conversational speaking scripts.

## ✨ Features

- 🤖 **AI-Powered**: Uses Google Gemini 2.5 Flash to generate natural speaker notes
- 📄 **PDF Support**: Converts PDF presentations to PPTX with speaker notes
- 📊 **PPTX Support**: Adds notes to existing PowerPoint presentations
- 🌐 **Web Interface**: Beautiful, modern UI for easy file upload and download
- 💬 **Natural Scripts**: Generates conversational text, not bullet points
- ⚡ **Fast Processing**: Efficient processing with progress updates
- 🎯 **Drag & Drop**: Simple drag-and-drop file upload
- 📥 **Auto Download**: Your file downloads automatically when ready

## 🚀 Quick Start

### Prerequisites

- Python 3.8 or higher
- Google AI API Key ([Get one here](https://aistudio.google.com/app/apikey))

### Installation

1. **Clone or download this repository**

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up your API key:**
   
   Edit the `.env` file and add your Google API key:
   ```
   GOOGLE_API_KEY=your-api-key-here
   ```

### Usage

#### Option 1: Web Interface (Recommended)

1. **Start the server:**
   ```bash
   python server.py
   ```

2. **Open your browser:**
   ```
   http://localhost:8000
   ```

3. **Upload your file:**
   - Drag and drop a PDF or PPTX file
   - Or click to browse and select a file

4. **Download:**
   - Your presentation with speaker notes downloads automatically!

#### Option 2: Command Line

**For PDF files:**
```bash
python add_speaker_notes.py presentation.pdf
```

**For PPTX files:**
```bash
python add_speaker_notes.py presentation.pptx
```

**Custom output filename:**
```bash
python add_speaker_notes.py input.pdf output.pptx
```

**Adjust image quality (PDF only):**
```bash
python add_speaker_notes.py input.pdf output.pptx 300
```
Higher DPI = better quality (default: 200)

## 📝 What You Get

### Speaker Notes Format

The AI generates **plain text speaking scripts** that:
- Sound natural when spoken aloud
- Are conversational and engaging
- Take 30-60 seconds to present per slide
- Are written in first person
- Contain no bullets, markdown, or special formatting
- Can be read directly during your presentation

### Example

Instead of bullet points like:
```
• Introduce product
• Highlight key features
• Mention pricing
```

You get natural speech like:
```
Today I'm excited to introduce our new product that's going to revolutionize 
how you work. We've packed it with features that our customers have been 
asking for, and I think you'll be really impressed with what we've built. 
Let me walk you through what makes this special...
```

## 📁 Project Structure

```
.
├── server.py                  # FastAPI web server
├── add_speaker_notes.py       # Main CLI script
├── pdf_to_pptx.py            # PDF to editable PPTX converter
├── requirements.txt           # Python dependencies
├── .env                       # API key configuration
└── README.md                  # This file
```

## 🛠️ Advanced Usage

### pdf_to_pptx.py - PDF to Editable PPTX

Convert PDFs to fully editable PowerPoint presentations with text recognition:

```bash
python pdf_to_pptx.py presentation.pdf
```

This script:
- Extracts text from PDF with positions and fonts
- Removes text from images
- Creates editable text boxes with original formatting
- Preserves colors, fonts, and styling

**Note:** Requires `GOOGLE_API_KEY` for text extraction via Gemini Vision API.

## 🔧 Configuration

### Environment Variables

Create or edit `.env` file:
```
GOOGLE_API_KEY=your-google-ai-api-key
```

### Server Configuration

Edit `server.py` to change:
- Port (default: 8000)
- Upload directory
- Output directory
- DPI settings

## 📋 Requirements

- python-pptx==0.6.23
- PyMuPDF==1.23.8
- Pillow==10.1.0
- google-genai
- python-dotenv
- fastapi
- uvicorn
- python-multipart

## 🎯 Use Cases

- **Business Presentations**: Add professional speaker notes to sales decks
- **Academic Lectures**: Generate talking points for educational slides
- **Conference Talks**: Prepare speaking scripts for technical presentations
- **Training Materials**: Create instructor notes for training presentations
- **Pitch Decks**: Add compelling narratives to investor presentations

## 🤝 Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.

## 📄 License

This project is open source and available under the MIT License.

## 🙏 Acknowledgments

- Powered by [Google Gemini AI](https://ai.google.dev/)
- Built with [FastAPI](https://fastapi.tiangolo.com/)
- PDF processing with [PyMuPDF](https://pymupdf.readthedocs.io/)
- PowerPoint manipulation with [python-pptx](https://python-pptx.readthedocs.io/)

## 💡 Tips

- **PDF Quality**: Higher DPI = better text extraction but slower processing
- **PPTX Input**: Works best with slides that have text content
- **Speaker Notes**: View notes in PowerPoint's Presenter View (Slide Show → Presenter View)
- **API Limits**: Google Gemini API has rate limits - if processing fails, wait a moment and retry

## 🐛 Troubleshooting

### "GOOGLE_API_KEY not found"
- Check that `.env` file exists in the project directory
- Verify your API key is correctly set in `.env`

### "Processing failed"
- Ensure your PDF/PPTX file is not corrupted
- Check that you have internet connection (for AI processing)
- Verify your API key is valid and has remaining quota

### Text not extracted properly
- Try increasing DPI for PDF files: `python add_speaker_notes.py file.pdf output.pptx 300`
- Ensure the PDF has actual text (not scanned images)

## 📞 Support

For issues and questions, please open an issue on the repository.

