# Standalone Hybrid PDF to Excel Converter

A simple, standalone Python script that converts PDF files to Excel format using a hybrid approach - it automatically detects whether the PDF contains extractable text or needs OCR processing.

## Features

- 🔍 **Automatic Detection**: Automatically detects if PDF is text-based or image-based
- 📄 **Text Extraction**: Uses pdfplumber for text-based PDFs
- 🖼️ **OCR Processing**: Uses Tesseract OCR for scanned/image-based PDFs
- 📊 **Table Support**: Extracts tables when present
- 🔄 **Fallback Mechanism**: Falls back to OCR if text extraction fails
- 📁 **Flexible Output**: Saves to Excel format with configurable settings

## Requirements

### Python Dependencies
```bash
pip install pdfplumber pytesseract pdf2image pandas openpyxl
```

### External Dependencies

#### Tesseract OCR
- **Windows**: Download from [UB-Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
- **Linux**: `sudo apt-get install tesseract-ocr`
- **macOS**: `brew install tesseract`

#### Poppler (for PDF to image conversion)
- **Windows**: Download from [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases)
- **Linux**: `sudo apt-get install poppler-utils`
- **macOS**: `brew install poppler`

## Quick Start

### 1. Configuration
Edit `config_standalone.py` to set your paths:

```python
# Update these paths for your system
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
POPPLER_PATH = r'C:\path\to\poppler\bin'  # Update this path
```

### 2. Usage

#### Command Line Usage
```bash
# Interactive mode
python standalone_hybrid_converter.py

# Direct mode
python standalone_hybrid_converter.py "input.pdf" "output.xlsx"
```

#### Python Script Usage
```python
from standalone_hybrid_converter import hybrid_pdf_to_excel

# Convert a PDF file
success = hybrid_pdf_to_excel("input.pdf", "output.xlsx")
if success:
    print("Conversion successful!")
else:
    print("Conversion failed!")
```

## Configuration Options

Edit `config_standalone.py` to customize:

- **OCR Settings**: DPI, page segmentation mode
- **Text Extraction**: Minimum text length thresholds
- **Output Settings**: Default directories, headers inclusion

## How It Works

1. **PDF Analysis**: Checks if PDF contains extractable text
2. **Text-Based PDFs**: Uses pdfplumber to extract text and tables
3. **Image-Based PDFs**: Converts to images and uses Tesseract OCR
4. **Fallback**: If text extraction fails, automatically tries OCR
5. **Output**: Saves extracted content to Excel format

## Troubleshooting

### Tesseract Not Found
```
❌ Tesseract not found. OCR functionality will be disabled.
```
**Solution**: Install Tesseract OCR and update the path in `config_standalone.py`

### Poppler Not Found
```
❌ Poppler not found.
```
**Solution**: Install Poppler and update the path in `config_standalone.py`

### No Content Extracted
```
❌ No content could be extracted from the PDF
```
**Possible Causes**:
- PDF is password protected
- PDF contains only images with poor quality
- PDF is corrupted

**Solutions**:
- Remove password protection
- Improve image quality
- Try a different PDF

## Examples

### Basic Conversion
```python
from standalone_hybrid_converter import hybrid_pdf_to_excel

# Convert a syllabus PDF
success = hybrid_pdf_to_excel("syllabus.pdf", "syllabus_converted.xlsx")
```

### Batch Processing
```python
import os
from standalone_hybrid_converter import hybrid_pdf_to_excel

pdf_folder = "pdfs/"
output_folder = "excel_output/"

for filename in os.listdir(pdf_folder):
    if filename.endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, filename)
        excel_path = os.path.join(output_folder, filename.replace(".pdf", ".xlsx"))
        hybrid_pdf_to_excel(pdf_path, excel_path)
```

## Comparison with Other Converters

| Feature | Standalone Converter | Streamlit App | Syllabus Converter |
|---------|---------------------|---------------|-------------------|
| **Ease of Use** | Simple script | Web interface | Specialized |
| **OCR Support** | ✅ | ✅ | ✅ |
| **Table Extraction** | ✅ | ✅ | ✅ |
| **Syllabus Structure** | Basic | Advanced | Advanced |
| **Batch Processing** | Manual | No | Manual |
| **Configuration** | File-based | UI-based | File-based |

## License

This project is part of the PDF to Excel Converter suite. See the main README for license information.
