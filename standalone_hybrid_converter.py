#!/usr/bin/env python3
"""
Standalone Hybrid PDF to Excel Converter
A simple, standalone script that can handle both text-based and image-based PDFs.
"""

import os
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import sys

# Import configuration
try:
    from config_standalone import (
        TESSERACT_PATH, POPPLER_PATH, OCR_DPI, OCR_CONFIG,
        MIN_TEXT_LENGTH, MIN_LINE_LENGTH, validate_tesseract_path
    )
except ImportError:
    # Fallback configuration if config file is not available
    TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    POPPLER_PATH = r'C:\path\to\poppler\bin'
    OCR_DPI = 300
    OCR_CONFIG = "--psm 6"
    MIN_TEXT_LENGTH = 50
    MIN_LINE_LENGTH = 2
    
    def validate_tesseract_path():
        """Fallback validation function"""
        possible_paths = [
            TESSERACT_PATH,
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            r".\tesseract\tesseract.exe",
            "/usr/bin/tesseract",  # Linux
            "/usr/local/bin/tesseract"  # macOS
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return True, path
        return False, None

def setup_tesseract():
    """Setup Tesseract OCR path"""
    tesseract_ok, tesseract_path = validate_tesseract_path()
    if tesseract_ok:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
        print(f"✅ Tesseract found at: {tesseract_path}")
        return True
    
    print("⚠️ Tesseract not found. OCR functionality will be disabled.")
    print("Please install Tesseract OCR or update the TESSERACT_PATH in config_standalone.py")
    return False
    
def is_text_based(pdf_path):
    """Check if PDF contains extractable text."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text and len(text.strip()) > MIN_TEXT_LENGTH:
                    return True
        return False
    except Exception as e:
        print(f"Error checking PDF type: {e}")
        return False

def extract_text_pdfplumber(pdf_path):
    """Extract structured text from a PDF using pdfplumber."""
    print("📄 Extracting text using pdfplumber...")
    rows = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                print(f"🔄 Processing page {page_num}/{len(pdf.pages)}...")
                
                # First try to extract tables
                tables = page.extract_tables()
                if tables:
                    print(f"📊 Found {len(tables)} table(s) on page {page_num}")
                    for table in tables:
                        rows.extend(table)
                else:
                    # Extract text if no tables found
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        for line in lines:
                            line = line.strip()
                            if line and len(line) >= MIN_LINE_LENGTH:  # Only add non-empty lines
                                rows.append([line])
        
        if not rows:
            print("⚠️ No content extracted with pdfplumber")
            return pd.DataFrame()
        
        df = pd.DataFrame(rows)
        print(f"✅ Extracted {len(df)} rows using pdfplumber")
        return df
        
    except Exception as e:
        print(f"❌ Error extracting with pdfplumber: {e}")
        return pd.DataFrame()

def extract_text_ocr(pdf_path):
    """Extract text from scanned/image PDFs using OCR."""
    if not setup_tesseract():
        print("❌ Tesseract not available. Cannot perform OCR.")
        return pd.DataFrame()
    
    print("🖼️ Extracting text using OCR...")
    rows = []
    
    try:
        # Convert PDF to images
        print("🔄 Converting PDF to images...")
        images = convert_from_path(pdf_path, dpi=OCR_DPI)
        
        for i, img in enumerate(images, 1):
            print(f"🔄 Processing page {i}/{len(images)} with OCR...")
            try:
                # Convert to grayscale for better OCR
                gray = img.convert("L")
                text = pytesseract.image_to_string(gray, config=OCR_CONFIG)
                
                for line in text.split('\n'):
                    line = line.strip()
                    if line and len(line) >= MIN_LINE_LENGTH:  # Only add non-empty lines
                        rows.append([line])
                        
            except Exception as e:
                print(f"⚠️ OCR failed on page {i}: {e}")
                continue
        
        if not rows:
            print("⚠️ No content extracted with OCR")
            return pd.DataFrame()
        
        df = pd.DataFrame(rows)
        print(f"✅ Extracted {len(df)} rows using OCR")
        return df
        
    except Exception as e:
        print(f"❌ Error extracting with OCR: {e}")
        return pd.DataFrame()

def hybrid_pdf_to_excel(pdf_path, excel_output_path):
    """Main function to convert PDF to Excel using hybrid approach."""
    print(f"🔍 Processing file: {pdf_path}")
    
    # Check if input file exists
    if not os.path.exists(pdf_path):
        print(f"❌ Error: File not found - {pdf_path}")
        return False
    
    # Check if output directory exists
    output_dir = os.path.dirname(excel_output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"📁 Created output directory: {output_dir}")
    
    try:
        # Determine PDF type and extract content
        if is_text_based(pdf_path):
            print("📄 Text-based PDF detected. Using pdfplumber...")
            df = extract_text_pdfplumber(pdf_path)
            
            # If pdfplumber fails, try OCR as fallback
            if df.empty:
                print("⚠️ pdfplumber failed, trying OCR as fallback...")
                df = extract_text_ocr(pdf_path)
        else:
            print("🖼️ Image-based PDF detected. Using OCR...")
            df = extract_text_ocr(pdf_path)
        
        # Check if we got any data
        if df.empty:
            print("❌ No content could be extracted from the PDF")
            return False
        
        # Save to Excel
        print(f"💾 Saving to Excel: {excel_output_path}")
        df.to_excel(excel_output_path, index=False, header=False, engine='openpyxl')
        
        print(f"✅ Successfully converted PDF to Excel!")
        print(f"📊 Extracted {len(df)} rows")
        print(f"📁 Output file: {excel_output_path}")
        
        return True
        
    except Exception as e:
        print(f"❌ Error during conversion: {e}")
        return False

def main():
    """Main function with command line interface."""
    print("=" * 60)
    print("🔧 Standalone Hybrid PDF to Excel Converter")
    print("=" * 60)
    
    # Check if file paths are provided as command line arguments
    if len(sys.argv) == 3:
        pdf_path = sys.argv[1]
        excel_output_path = sys.argv[2]
    else:
        # Interactive mode
        print("\n📁 Enter file paths:")
        pdf_path = input("PDF file path: ").strip().strip('"')
        excel_output_path = input("Excel output path: ").strip().strip('"')
    
    # Validate input
    if not pdf_path:
        print("❌ Please provide a PDF file path")
        return
    
    # Validate PDF file extension
    if not pdf_path.lower().endswith('.pdf'):
        print("❌ Input file must be a PDF file (.pdf extension)")
        return
    
    if not excel_output_path:
        # Generate default output path
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        excel_output_path = f"{base_name}_converted.xlsx"
        print(f"📁 Using default output path: {excel_output_path}")
    else:
        # Ensure Excel output has .xlsx extension
        if not excel_output_path.lower().endswith('.xlsx'):
            excel_output_path = excel_output_path + '.xlsx'
            print(f"📁 Added .xlsx extension: {excel_output_path}")
    
    # Run conversion
    success = hybrid_pdf_to_excel(pdf_path, excel_output_path)
    
    if success:
        print("\n🎉 Conversion completed successfully!")
    else:
        print("\n❌ Conversion failed. Please check the error messages above.")
        sys.exit(1)

if __name__ == "__main__":
    main()
