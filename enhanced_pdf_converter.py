#!/usr/bin/env python3
"""
Enhanced PDF to Excel Converter
Advanced converter that preserves all data without truncation and handles tables properly.
"""

import os
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import sys
import re
from typing import List, Dict, Any, Tuple
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Import configuration
try:
    from config_standalone import (
        TESSERACT_PATH, POPPLER_PATH, OCR_DPI, OCR_CONFIG,
        MIN_TEXT_LENGTH, MIN_LINE_LENGTH, validate_tesseract_path
    )
except ImportError:
    # Fallback configuration
    TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    POPPLER_PATH = r'C:\path\to\poppler\bin'
    OCR_DPI = 300
    OCR_CONFIG = "--psm 6"
    MIN_TEXT_LENGTH = 50
    MIN_LINE_LENGTH = 2
    
    def validate_tesseract_path():
        possible_paths = [
            TESSERACT_PATH,
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            r".\tesseract\tesseract.exe",
            "/usr/bin/tesseract",
            "/usr/local/bin/tesseract"
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return True, path
        return False, None

def setup_tesseract():
    """Setup Tesseract OCR path with enhanced configuration"""
    tesseract_ok, tesseract_path = validate_tesseract_path()
    if tesseract_ok:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path
        logger.info(f"✅ Tesseract found at: {tesseract_path}")
        return True
    
    logger.warning("⚠️ Tesseract not found. OCR functionality will be disabled.")
    return False

def is_text_based(pdf_path: str) -> bool:
    """Enhanced check if PDF contains extractable text"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_text_length = 0
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    total_text_length += len(text.strip())
            
            # More lenient threshold for text detection
            return total_text_length > MIN_TEXT_LENGTH
    except Exception as e:
        logger.error(f"Error checking PDF type: {e}")
        return False

def extract_tables_enhanced(page) -> List[List[List[str]]]:
    """Enhanced table extraction with better structure preservation"""
    tables = []
    try:
        # Extract tables with enhanced settings
        extracted_tables = page.extract_tables({
            'vertical_strategy': 'text',
            'horizontal_strategy': 'text',
            'intersection_x_tolerance': 10,
            'intersection_y_tolerance': 10
        })
        
        for table in extracted_tables:
            if table and len(table) > 0:
                # Clean and validate table data
                cleaned_table = []
                for row in table:
                    cleaned_row = []
                    for cell in row:
                        if cell is not None:
                            # Preserve full cell content without truncation
                            cell_text = str(cell).strip()
                            if cell_text:
                                cleaned_row.append(cell_text)
                            else:
                                cleaned_row.append("")
                        else:
                            cleaned_row.append("")
                    if any(cell for cell in cleaned_row):  # Only add non-empty rows
                        cleaned_table.append(cleaned_row)
                
                if cleaned_table:
                    tables.append(cleaned_table)
                    
    except Exception as e:
        logger.warning(f"Error extracting tables: {e}")
    
    return tables

def extract_text_enhanced(page) -> List[str]:
    """Enhanced text extraction that preserves full content"""
    lines = []
    try:
        # Extract text with word-level precision
        words = page.extract_words(
            x_tolerance=3,
            y_tolerance=3,
            keep_blank_chars=False,
            use_text_flow=True
        )
        
        if words:
            # Group words by line based on y-position
            current_line = []
            current_y = None
            
            for word in words:
                word_text = word.get('text', '').strip()
                word_y = word.get('top', 0)
                
                if not word_text:
                    continue
                
                # Check if this word belongs to the same line
                if current_y is None or abs(word_y - current_y) < 10:
                    current_line.append(word_text)
                    current_y = word_y
                else:
                    # New line detected
                    if current_line:
                        full_line = ' '.join(current_line)
                        if len(full_line.strip()) >= MIN_LINE_LENGTH:
                            lines.append(full_line.strip())
                    current_line = [word_text]
                    current_y = word_y
            
            # Don't forget the last line
            if current_line:
                full_line = ' '.join(current_line)
                if len(full_line.strip()) >= MIN_LINE_LENGTH:
                    lines.append(full_line.strip())
        
        # Fallback to simple text extraction if word extraction fails
        if not lines:
            text = page.extract_text()
            if text:
                lines = [line.strip() for line in text.split('\n') if line.strip() and len(line.strip()) >= MIN_LINE_LENGTH]
                
    except Exception as e:
        logger.warning(f"Error extracting text: {e}")
        # Final fallback
        try:
            text = page.extract_text()
            if text:
                lines = [line.strip() for line in text.split('\n') if line.strip() and len(line.strip()) >= MIN_LINE_LENGTH]
        except:
            pass
    
    return lines

def extract_content_pdfplumber_enhanced(pdf_path: str) -> Dict[str, Any]:
    """Enhanced content extraction using pdfplumber"""
    logger.info("📄 Extracting content using enhanced pdfplumber...")
    
    all_tables = []
    all_text = []
    page_info = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                logger.info(f"🔄 Processing page {page_num}/{len(pdf.pages)}...")
                
                page_data = {
                    'page_number': page_num,
                    'tables': [],
                    'text_lines': [],
                    'has_content': False
                }
                
                # Extract tables first
                tables = extract_tables_enhanced(page)
                if tables:
                    logger.info(f"📊 Found {len(tables)} table(s) on page {page_num}")
                    for table_idx, table in enumerate(tables):
                        table_info = {
                            'page': page_num,
                            'table_index': table_idx,
                            'rows': len(table),
                            'columns': len(table[0]) if table else 0,
                            'data': table
                        }
                        all_tables.append(table_info)
                        page_data['tables'].append(table_info)
                        page_data['has_content'] = True
                
                # Extract text if no tables or as supplementary content
                text_lines = extract_text_enhanced(page)
                if text_lines:
                    logger.info(f"📝 Extracted {len(text_lines)} text lines from page {page_num}")
                    for line_idx, line in enumerate(text_lines, 1):
                        text_info = {
                            'page': page_num,
                            'line_number': line_idx,
                            'content': line,
                            'type': 'text'
                        }
                        all_text.append(text_info)
                        page_data['text_lines'].append(text_info)
                        page_data['has_content'] = True
                
                page_info.append(page_data)
        
        if not all_tables and not all_text:
            logger.warning("⚠️ No content extracted with pdfplumber")
            return {'tables': [], 'text': [], 'pages': page_info}
        
        logger.info(f"✅ Extracted {len(all_tables)} tables and {len(all_text)} text lines")
        return {
            'tables': all_tables,
            'text': all_text,
            'pages': page_info
        }
        
    except Exception as e:
        logger.error(f"❌ Error extracting with pdfplumber: {e}")
        return {'tables': [], 'text': [], 'pages': []}

def extract_text_ocr_enhanced(pdf_path: str) -> Dict[str, Any]:
    """Enhanced OCR extraction with better accuracy"""
    if not setup_tesseract():
        logger.error("❌ Tesseract not available. Cannot perform OCR.")
        return {'tables': [], 'text': [], 'pages': []}
    
    logger.info("🖼️ Extracting text using enhanced OCR...")
    
    all_text = []
    page_info = []
    
    try:
        # Convert PDF to images with higher DPI for better accuracy
        logger.info("🔄 Converting PDF to images...")
        images = convert_from_path(pdf_path, dpi=OCR_DPI + 100)  # Higher DPI for better accuracy
        
        for page_num, img in enumerate(images, 1):
            logger.info(f"🔄 Processing page {page_num}/{len(images)} with OCR...")
            
            page_data = {
                'page_number': page_num,
                'tables': [],
                'text_lines': [],
                'has_content': False
            }
            
            try:
                # Preprocess image for better OCR
                gray = img.convert("L")
                
                # Use multiple OCR configurations for better results
                ocr_configs = [
                    "--psm 6",  # Uniform block of text
                    "--psm 3",  # Fully automatic page segmentation
                    "--psm 4"   # Assume a single column of text
                ]
                
                best_text = ""
                for config in ocr_configs:
                    try:
                        text = pytesseract.image_to_string(gray, config=config)
                        if len(text.strip()) > len(best_text.strip()):
                            best_text = text
                    except:
                        continue
                
                if best_text:
                    lines = best_text.split('\n')
                    for line_idx, line in enumerate(lines, 1):
                        line = line.strip()
                        if line and len(line) >= MIN_LINE_LENGTH:
                            text_info = {
                                'page': page_num,
                                'line_number': line_idx,
                                'content': line,
                                'type': 'ocr_text'
                            }
                            all_text.append(text_info)
                            page_data['text_lines'].append(text_info)
                            page_data['has_content'] = True
                    
                    logger.info(f"✅ Extracted {len([l for l in lines if l.strip()])} lines from page {page_num}")
                else:
                    logger.warning(f"⚠️ No text extracted from page {page_num}")
                    
            except Exception as e:
                logger.warning(f"⚠️ OCR failed on page {page_num}: {e}")
                continue
            
            page_info.append(page_data)
        
        if not all_text:
            logger.warning("⚠️ No content extracted with OCR")
            return {'tables': [], 'text': [], 'pages': page_info}
        
        logger.info(f"✅ Extracted {len(all_text)} text lines using OCR")
        return {
            'tables': [],
            'text': all_text,
            'pages': page_info
        }
        
    except Exception as e:
        logger.error(f"❌ Error extracting with OCR: {e}")
        return {'tables': [], 'text': [], 'pages': []}

def create_enhanced_excel_output(content_data: Dict[str, Any], excel_path: str) -> bool:
    """Create enhanced Excel output with multiple sheets and preserved structure"""
    try:
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            
            # Sheet 1: All Text Content (Preserved)
            if content_data['text']:
                text_df = pd.DataFrame(content_data['text'])
                text_df.to_excel(writer, sheet_name='All_Text_Content', index=False)
                logger.info(f"📝 Saved {len(text_df)} text lines to 'All_Text_Content' sheet")
            
            # Sheet 2: Tables (Preserved Structure)
            if content_data['tables']:
                for table_idx, table_info in enumerate(content_data['tables']):
                    table_df = pd.DataFrame(table_info['data'])
                    sheet_name = f'Table_{table_info["page"]}_{table_idx + 1}'
                    # Truncate sheet name if too long
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:31]
                    table_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                    logger.info(f"📊 Saved table {table_idx + 1} from page {table_info['page']}")
            
            # Sheet 3: Page Summary
            page_summary = []
            for page_data in content_data['pages']:
                if page_data['has_content']:
                    page_summary.append({
                        'Page_Number': page_data['page_number'],
                        'Tables_Found': len(page_data['tables']),
                        'Text_Lines': len(page_data['text_lines']),
                        'Total_Content_Items': len(page_data['tables']) + len(page_data['text_lines'])
                    })
            
            if page_summary:
                summary_df = pd.DataFrame(page_summary)
                summary_df.to_excel(writer, sheet_name='Page_Summary', index=False)
                logger.info(f"📋 Saved page summary for {len(page_summary)} pages")
            
            # Sheet 4: Raw Data (for debugging)
            raw_data = []
            for text_item in content_data['text']:
                raw_data.append({
                    'Page': text_item['page'],
                    'Line_Number': text_item['line_number'],
                    'Content': text_item['content'],
                    'Type': text_item['type']
                })
            
            if raw_data:
                raw_df = pd.DataFrame(raw_data)
                raw_df.to_excel(writer, sheet_name='Raw_Data', index=False)
                logger.info(f"🔍 Saved {len(raw_df)} raw data items")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ Error creating Excel file: {e}")
        return False

def enhanced_pdf_to_excel(pdf_path: str, excel_output_path: str) -> bool:
    """Enhanced main function for PDF to Excel conversion"""
    logger.info(f"🔍 Processing file: {pdf_path}")
    
    # Check if input file exists
    if not os.path.exists(pdf_path):
        logger.error(f"❌ Error: File not found - {pdf_path}")
        return False
    
    # Check if output directory exists
    output_dir = os.path.dirname(excel_output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
        logger.info(f"📁 Created output directory: {output_dir}")
    
    try:
        # Determine PDF type and extract content
        if is_text_based(pdf_path):
            logger.info("📄 Text-based PDF detected. Using enhanced pdfplumber...")
            content_data = extract_content_pdfplumber_enhanced(pdf_path)
            
            # If pdfplumber fails, try OCR as fallback
            if not content_data['text'] and not content_data['tables']:
                logger.warning("⚠️ pdfplumber failed, trying OCR as fallback...")
                content_data = extract_text_ocr_enhanced(pdf_path)
        else:
            logger.info("🖼️ Image-based PDF detected. Using enhanced OCR...")
            content_data = extract_text_ocr_enhanced(pdf_path)
        
        # Check if we got any data
        total_items = len(content_data['text']) + len(content_data['tables'])
        if total_items == 0:
            logger.error("❌ No content could be extracted from the PDF")
            return False
        
        # Create enhanced Excel output
        logger.info(f"💾 Saving to Excel: {excel_output_path}")
        success = create_enhanced_excel_output(content_data, excel_output_path)
        
        if success:
            logger.info(f"✅ Successfully converted PDF to Excel!")
            logger.info(f"📊 Extracted {len(content_data['text'])} text lines and {len(content_data['tables'])} tables")
            logger.info(f"📁 Output file: {excel_output_path}")
            return True
        else:
            logger.error("❌ Failed to create Excel file")
            return False
        
    except Exception as e:
        logger.error(f"❌ Error during conversion: {e}")
        return False

def main():
    """Main function with command line interface"""
    print("=" * 70)
    print("🔧 Enhanced PDF to Excel Converter")
    print("📊 Preserves all data without truncation")
    print("=" * 70)
    
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
        excel_output_path = f"{base_name}_enhanced_converted.xlsx"
        print(f"📁 Using default output path: {excel_output_path}")
    else:
        # Ensure Excel output has .xlsx extension
        if not excel_output_path.lower().endswith('.xlsx'):
            excel_output_path = excel_output_path + '.xlsx'
            print(f"📁 Added .xlsx extension: {excel_output_path}")
    
    # Run conversion
    success = enhanced_pdf_to_excel(pdf_path, excel_output_path)
    
    if success:
        print("\n🎉 Enhanced conversion completed successfully!")
        print("📋 Check the multiple sheets in the Excel file for complete data")
    else:
        print("\n❌ Conversion failed. Please check the error messages above.")
        sys.exit(1)

if __name__ == "__main__":
    main()
