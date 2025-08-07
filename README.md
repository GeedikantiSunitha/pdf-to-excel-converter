# 📊 PDF to Excel Converter

A powerful and user-friendly application that converts PDF files to Excel format with advanced text and table extraction capabilities. Built with Streamlit, this application supports both text-based PDFs and scanned documents using OCR technology.

## ✨ Features

- **Hybrid Extraction**: Combines text-based extraction (pdfplumber) with OCR (Tesseract) for maximum compatibility
- **Table Detection**: Automatically detects and extracts tables from PDF documents
- **Modern UI**: Beautiful, responsive interface built with Streamlit
- **Formatting Options**: Configurable Excel output with styling and formatting
- **Progress Tracking**: Real-time progress indicators during conversion
- **Preview Functionality**: Preview extracted data before downloading
- **Batch Processing**: Support for multiple file formats and sizes
- **Error Handling**: Comprehensive error handling and user feedback

## 🚀 Quick Start

### Prerequisites

1. **Python 3.8+** installed on your system
2. **Tesseract OCR** installed (for scanned PDF support)

#### Installing Tesseract OCR

**Windows:**
```bash
# Download and install from: https://github.com/UB-Mannheim/tesseract/wiki
# Add to PATH: C:\Program Files\Tesseract-OCR
```

**macOS:**
```bash
brew install tesseract
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install tesseract-ocr
```

### Installation

1. **Clone the repository:**
```bash
git clone <repository-url>
cd pdf-to-excel-converter
```

2. **Install Python dependencies:**
```bash
pip install -r requirements.txt
```

3. **Run the application:**
```bash
streamlit run app.py
```

4. **Open your browser** and navigate to `http://localhost:8501`

## 📖 Usage

### Basic Usage

1. **Upload PDF**: Click the upload area and select your PDF file
2. **Configure Settings**: Adjust extraction options in the sidebar
3. **Convert**: Click the "Convert to Excel" button
4. **Download**: Download your converted Excel file

### Advanced Settings

#### Extraction Options
- **Enable OCR**: Toggle OCR for scanned PDFs
- **OCR DPI**: Adjust image quality for OCR processing (150-600 DPI)

#### Output Options
- **Include Formatting**: Add colors, borders, and styling to Excel
- **Auto-adjust Columns**: Automatically resize columns based on content

#### Advanced Options
- **Minimum Text Threshold**: Minimum characters to consider PDF as text-based
- **Maximum Column Width**: Maximum width for auto-adjusted columns

## 🏗️ Project Structure

```
pdf-to-excel-converter/
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── README.md             # Project documentation
├── core/                 # Core processing modules
│   ├── __init__.py
│   ├── extractor.py      # PDF extraction logic
│   └── converter.py      # Excel conversion logic
├── output/               # Generated Excel files
├── assets/               # Static assets
└── utils/                # Utility functions
```

## 🔧 Core Modules

### PDFExtractor (`core/extractor.py`)
- Handles text and table extraction from PDFs
- Supports both text-based and image-based PDFs
- Uses pdfplumber for text extraction
- Falls back to Tesseract OCR for scanned documents

### ExcelConverter (`core/converter.py`)
- Converts extracted data to Excel format
- Applies formatting and styling
- Handles multiple tables and sheets
- Auto-adjusts column widths

## 🎯 Supported PDF Types

### Text-based PDFs
- Documents with selectable text
- Tables with clear structure
- Forms and reports
- Invoices and receipts

### Scanned PDFs
- Image-based documents
- Handwritten text (with OCR)
- Printed documents
- Historical documents

## 📊 Output Format

The application generates Excel files with:
- **Multiple Sheets**: One sheet per table found
- **Formatted Headers**: Styled headers with colors and borders
- **Auto-sized Columns**: Optimized column widths
- **Data Validation**: Clean, structured data
- **Metadata**: Processing information and timestamps

## 🛠️ Troubleshooting

### Common Issues

1. **Tesseract not found**
   - Ensure Tesseract is installed and in your PATH
   - Restart your terminal after installation

2. **PDF extraction fails**
   - Check if the PDF contains readable text
   - Try enabling OCR for scanned documents
   - Verify PDF file integrity

3. **Memory issues with large files**
   - Reduce OCR DPI setting
   - Process smaller files
   - Increase system memory

### Performance Tips

- **For text-based PDFs**: Disable OCR for faster processing
- **For scanned PDFs**: Use higher DPI for better accuracy
- **For large files**: Process in smaller batches

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **Streamlit** for the amazing web framework
- **pdfplumber** for PDF text extraction
- **Tesseract** for OCR capabilities
- **pandas** for data manipulation
- **openpyxl** for Excel file generation

## 📞 Support

If you encounter any issues or have questions:
1. Check the troubleshooting section
2. Search existing issues
3. Create a new issue with detailed information

---

**Made with ❤️ for the data community** 