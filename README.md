# Enhanced PDF to Excel Converter

A powerful Python tool for converting PDF files to Excel format with **complete data preservation** and **no truncation**. This enhanced converter automatically detects PDF types and uses the best extraction method for optimal results.

## 🚀 Key Features

- **🔍 Smart Detection**: Automatically detects text-based vs image-based PDFs
- **📄 Full Data Preservation**: No text truncation or data loss
- **📊 Table Structure**: Preserves original table formatting and columns
- **🖼️ Advanced OCR**: Multiple OCR configurations for best accuracy
- **📋 Multiple Output Sheets**: Organized Excel output with separate sheets
- **🔧 Easy Configuration**: Simple setup with configurable settings
- **📁 Batch Processing**: Handle multiple files efficiently

## 📁 Project Structure

```
pdf-to-excel-converter/
├── 🚀 enhanced_pdf_converter.py    # Main enhanced converter
├── ⚙️ config_standalone.py         # Configuration settings
├── 📖 README_standalone.md         # Detailed documentation
├── 📋 example_usage.py             # Usage examples
├── 📦 requirements.txt             # Python dependencies
├── 📁 input_folder/                # Input PDF files
├── 📁 output_folder/               # Output Excel files
└── 📁 temp/                        # Temporary files
```

## 🛠️ Installation

### Prerequisites

1. **Python 3.7+**
2. **Tesseract OCR** (for OCR functionality)
3. **Poppler** (for PDF to image conversion)

### Install Dependencies

```bash
# Clone the repository
git clone https://github.com/GeedikantiSunitha/pdf-to-excel-converter.git
cd pdf-to-excel-converter

# Install Python dependencies
pip install -r requirements.txt
```

### External Dependencies

#### Tesseract OCR
- **Windows**: Download from [UB-Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
- **Linux**: `sudo apt-get install tesseract-ocr`
- **macOS**: `brew install tesseract`

#### Poppler
- **Windows**: Download from [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases)
- **Linux**: `sudo apt-get install poppler-utils`
- **macOS**: `brew install poppler`

## 🚀 Quick Start

### 1. Basic Usage

```bash
# Interactive mode
python enhanced_pdf_converter.py

# Direct mode
python enhanced_pdf_converter.py "input.pdf" "output.xlsx"
```

### 2. Python Script Usage

```python
from enhanced_pdf_converter import enhanced_pdf_to_excel

success = enhanced_pdf_to_excel("input.pdf", "output.xlsx")
if success:
    print("✅ Conversion successful!")
```

### 3. Batch Processing

```python
import os
from enhanced_pdf_converter import enhanced_pdf_to_excel

for filename in os.listdir("pdfs/"):
    if filename.endswith(".pdf"):
        pdf_path = f"pdfs/{filename}"
        excel_path = f"output/{filename.replace('.pdf', '.xlsx')}"
        enhanced_pdf_to_excel(pdf_path, excel_path)
```

## 📊 Excel Output Structure

The enhanced converter creates an Excel file with multiple sheets:

1. **All_Text_Content**: Complete text extraction with page and line numbers
2. **Tables**: Individual sheets for each table found (preserved structure)
3. **Page_Summary**: Overview of content found on each page
4. **Raw_Data**: Detailed extraction data for debugging

## 🔧 Configuration

Edit `config_standalone.py` to customize settings:

```python
# OCR Settings
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
POPPLER_PATH = r'C:\path\to\poppler\bin'
OCR_DPI = 300  # Higher DPI = better quality but slower
OCR_CONFIG = "--psm 6"

# Text Extraction
MIN_TEXT_LENGTH = 50  # Minimum characters to consider as text-based PDF
MIN_LINE_LENGTH = 2   # Minimum characters for a line to be included

# Output Settings
DEFAULT_OUTPUT_DIR = "output"
INCLUDE_HEADERS = False
```

## 🎯 Advanced Features

### 1. **Enhanced Text Extraction**
- Word-level precision extraction
- Preserves full content without truncation
- Multiple fallback methods for reliability

### 2. **Smart Table Detection**
- Preserves original table structure
- Maintains column alignment
- Handles complex table layouts

### 3. **Advanced OCR**
- Multiple OCR configurations for best results
- Higher DPI processing for accuracy
- Automatic configuration selection

### 4. **Comprehensive Error Handling**
- Graceful fallbacks between methods
- Detailed logging and error reporting
- Robust file validation

## 📋 Usage Examples

### Check Configuration
```bash
python config_standalone.py
```

### Run Examples
```bash
python example_usage.py
```

### Convert with Error Handling
```python
from enhanced_pdf_converter import enhanced_pdf_to_excel
import os

def safe_convert(pdf_path, excel_path):
    if not os.path.exists(pdf_path):
        print(f"❌ File not found: {pdf_path}")
        return False
    
    success = enhanced_pdf_to_excel(pdf_path, excel_path)
    if success and os.path.exists(excel_path):
        file_size = os.path.getsize(excel_path)
        print(f"✅ Success! File size: {file_size} bytes")
        return True
    return False

# Usage
safe_convert("document.pdf", "output.xlsx")
```

## 🐛 Troubleshooting

### Common Issues

1. **Tesseract Not Found**
   ```
   ❌ Tesseract not available. Cannot perform OCR.
   ```
   **Solution**: Install Tesseract OCR and update path in `config_standalone.py`

2. **No Content Extracted**
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

3. **Table Structure Lost**
   ```
   Tables appear as single lines
   ```
   **Solution**: The enhanced converter automatically preserves table structure

### Performance Tips

- **For Speed**: Use lower OCR_DPI (200-300)
- **For Accuracy**: Use higher OCR_DPI (400-600)
- **For Large Files**: Process in batches
- **For Tables**: The enhanced converter automatically handles table preservation

## 📄 Requirements

```
pdfplumber>=0.7.0
pytesseract>=0.3.8
pdf2image>=1.16.0
pandas>=1.3.0
openpyxl>=3.0.0
Pillow>=8.0.0
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License.

## 🙏 Acknowledgments

- [pdfplumber](https://github.com/jsvine/pdfplumber) for PDF text extraction
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) for OCR functionality
- [pandas](https://pandas.pydata.org/) for data manipulation

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/GeedikantiSunitha/pdf-to-excel-converter/issues)
- **Documentation**: See `README_standalone.md` for detailed documentation
- **Examples**: Check `example_usage.py` for usage examples

---

**Made with ❤️ for efficient PDF to Excel conversion with complete data preservation** 