#!/usr/bin/env python3
"""
Syllabus PDF to Excel Converter
Specialized converter for educational syllabus PDFs that organizes content by units and topics.
"""

import fitz  # PyMuPDF
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import os
import tempfile
import re
from PIL import Image

# Configure tesseract path
def setup_tesseract():
    possible_paths = [
        r"C:\tesseract\tesseract.exe",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r".\tesseract\tesseract.exe"
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return True
    
    return False

# Setup tesseract
tesseract_available = setup_tesseract()

def is_text_based(pdf_path):
    """Determine if PDF has extractable text using PyMuPDF"""
    try:
        with fitz.open(pdf_path) as doc:
            for page in doc:
                if len(page.get_text("text").strip()) > 50:
                    return True
        return False
    except Exception as e:
        print(f"Error detecting text: {e}")
        return False

def extract_syllabus_structure(pdf_path):
    """Extract syllabus content and organize by units and topics"""
    data = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            current_unit = ""
            current_topic = ""
            current_subtopic = ""
            
            for page_num, page in enumerate(pdf.pages, 1):
                print(f"🔄 Processing page {page_num}...")
                try:
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        
                        for line in lines:
                            line = line.strip()
                            if not line:
                                continue
                            
                            # Detect unit headers (usually numbers like "1.", "2.", etc.)
                            unit_match = re.match(r'^(\d+)\.?\s*(.+)$', line)
                            if unit_match:
                                current_unit = f"Unit {unit_match.group(1)}: {unit_match.group(2).strip()}"
                                current_topic = ""
                                current_subtopic = ""
                                data.append({
                                    'Page': page_num,
                                    'Unit': current_unit,
                                    'Topic': '',
                                    'Subtopic': '',
                                    'Content': line,
                                    'Level': 'Unit'
                                })
                                continue
                            
                            # Detect topic headers (usually letters like "a)", "b)", etc. or bullet points)
                            topic_match = re.match(r'^([a-z]\)|•|\*|\-)\s*(.+)$', line, re.IGNORECASE)
                            if topic_match:
                                current_topic = topic_match.group(2).strip()
                                current_subtopic = ""
                                data.append({
                                    'Page': page_num,
                                    'Unit': current_unit,
                                    'Topic': current_topic,
                                    'Subtopic': '',
                                    'Content': line,
                                    'Level': 'Topic'
                                })
                                continue
                            
                            # Detect subtopics (usually indented or with different markers)
                            subtopic_match = re.match(r'^(\s+)(.+)$', line)
                            if subtopic_match and current_topic:
                                current_subtopic = subtopic_match.group(2).strip()
                                data.append({
                                    'Page': page_num,
                                    'Unit': current_unit,
                                    'Topic': current_topic,
                                    'Subtopic': current_subtopic,
                                    'Content': line,
                                    'Level': 'Subtopic'
                                })
                                continue
                            
                            # Regular content line
                            if current_unit:
                                data.append({
                                    'Page': page_num,
                                    'Unit': current_unit,
                                    'Topic': current_topic,
                                    'Subtopic': current_subtopic,
                                    'Content': line,
                                    'Level': 'Content'
                                })
                            else:
                                # If no unit detected yet, treat as general content
                                data.append({
                                    'Page': page_num,
                                    'Unit': '',
                                    'Topic': '',
                                    'Subtopic': '',
                                    'Content': line,
                                    'Level': 'General'
                                })
                                
                except Exception as e:
                    print(f"⚠️ Error processing page {page_num}: {e}")
                    continue
        
        if not data:
            raise ValueError("No data extracted from any page")
            
        df = pd.DataFrame(data)
        print(f"✅ Extracted {len(df)} items from {df['Page'].nunique()} pages")
        return df
    except Exception as e:
        print(f"Error extracting syllabus: {e}")
        return pd.DataFrame()

def extract_text_ocr(pdf_path):
    """Extract text using OCR for image-based PDFs"""
    if not tesseract_available:
        print("❌ Tesseract OCR is not installed. Please install Tesseract.")
        return pd.DataFrame()
    
    try:
        temp_dir = tempfile.mkdtemp()
        print("🖼️ Converting PDF to images for OCR...")
        
        images = convert_from_path(pdf_path, dpi=300, output_folder=temp_dir)
        all_text = []
        
        for page_num, img in enumerate(images, start=1):
            print(f"🔄 Processing page {page_num}/{len(images)} with OCR...")
            try:
                gray = img.convert("L")
                text = pytesseract.image_to_string(gray, config="--psm 6")
                lines = text.split("\n")
                for line in lines:
                    line = line.strip()
                    if line:
                        all_text.append((f"Page {page_num}", line))
            except Exception as e:
                print(f"⚠️ OCR failed on page {page_num}: {e}")
                continue
        
        df = pd.DataFrame(all_text, columns=["Page", "Content"])
        print(f"✅ Extracted {len(df)} lines from {df['Page'].nunique()} pages using OCR")
        return df
    except Exception as e:
        print(f"OCR Error: {e}")
        return pd.DataFrame()

def organize_syllabus_data(df):
    """Organize the extracted data into a proper syllabus structure"""
    if df.empty:
        return df
    
    # If we have structured data (with Unit, Topic columns)
    if 'Unit' in df.columns:
        # Clean up the data
        df = df[df['Content'].str.len() > 2]  # Remove very short lines
        
        # Fill forward unit and topic information
        df['Unit'] = df['Unit'].fillna(method='ffill')
        df['Topic'] = df['Topic'].fillna(method='ffill')
        
        # Create a clean syllabus structure
        syllabus_data = []
        
        for _, row in df.iterrows():
            if row['Level'] in ['Unit', 'Topic', 'Subtopic']:
                syllabus_data.append({
                    'Page': row['Page'],
                    'Unit': row['Unit'],
                    'Topic': row['Topic'],
                    'Subtopic': row['Subtopic'],
                    'Content': row['Content'],
                    'Type': row['Level']
                })
        
        return pd.DataFrame(syllabus_data)
    
    # If we have raw text data, try to organize it
    else:
        # Try to detect structure from raw text
        organized_data = []
        current_unit = ""
        current_topic = ""
        
        for _, row in df.iterrows():
            content = row['Content']
            
            # Detect units (numbers followed by text)
            unit_match = re.match(r'^(\d+)\.?\s*(.+)$', content)
            if unit_match:
                current_unit = f"Unit {unit_match.group(1)}: {unit_match.group(2).strip()}"
                current_topic = ""
                organized_data.append({
                    'Page': row['Page'],
                    'Unit': current_unit,
                    'Topic': '',
                    'Subtopic': '',
                    'Content': content,
                    'Type': 'Unit'
                })
                continue
            
            # Detect topics (letters or bullet points)
            topic_match = re.match(r'^([a-z]\)|•|\*|\-)\s*(.+)$', content, re.IGNORECASE)
            if topic_match:
                current_topic = topic_match.group(2).strip()
                organized_data.append({
                    'Page': row['Page'],
                    'Unit': current_unit,
                    'Topic': current_topic,
                    'Subtopic': '',
                    'Content': content,
                    'Type': 'Topic'
                })
                continue
            
            # Regular content
            organized_data.append({
                'Page': row['Page'],
                'Unit': current_unit,
                'Topic': current_topic,
                'Subtopic': '',
                'Content': content,
                'Type': 'Content'
            })
        
        return pd.DataFrame(organized_data)

def syllabus_pdf_to_excel(pdf_path, excel_path):
    """Main function to convert syllabus PDF to Excel"""
    try:
        print("🔍 Analyzing PDF type...")
        
        if is_text_based(pdf_path):
            print("📄 Text-based PDF detected. Extracting syllabus structure...")
            try:
                df = extract_syllabus_structure(pdf_path)
                if df.empty:
                    raise ValueError("No syllabus structure found")
            except Exception as e:
                print(f"❌ Syllabus extraction failed, trying OCR. Reason: {e}")
                if tesseract_available:
                    df = extract_text_ocr(pdf_path)
                    df = organize_syllabus_data(df)
                else:
                    print("❌ OCR not available and syllabus extraction failed")
                    return pd.DataFrame()
        else:
            print("🖼️ Image-based PDF detected. Using OCR...")
            df = extract_text_ocr(pdf_path)
            df = organize_syllabus_data(df)

        if df.empty:
            raise ValueError("No content extracted from PDF")
        
        # Organize the data
        df = organize_syllabus_data(df)
        
        # Save to Excel with multiple sheets
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # Main organized data
            df.to_excel(writer, sheet_name='Syllabus_Organized', index=False)
            
            # Summary by units
            if 'Unit' in df.columns:
                unit_summary = df[df['Unit'] != ''].groupby('Unit').agg({
                    'Topic': 'count',
                    'Content': lambda x: '; '.join(x.unique())
                }).reset_index()
                unit_summary.columns = ['Unit', 'Topic_Count', 'Content_Summary']
                unit_summary.to_excel(writer, sheet_name='Unit_Summary', index=False)
            
            # Topics only
            topics_df = df[df['Type'].isin(['Topic', 'Subtopic'])].copy()
            if not topics_df.empty:
                topics_df.to_excel(writer, sheet_name='Topics_Only', index=False)
        
        print(f"✅ Excel saved to: {excel_path}")
        return df
    except Exception as e:
        print(f"❌ Failed to extract syllabus: {e}")
        return pd.DataFrame()

def main():
    """Main function to run the syllabus converter"""
    print("📚 Syllabus PDF to Excel Converter")
    print("=" * 50)
    
    # Show OCR status
    if tesseract_available:
        print("✅ OCR (Tesseract) is available")
    else:
        print("⚠️ OCR (Tesseract) is not available. Only text-based PDFs can be processed.")
    
    print()
    
    # Get input file path
    while True:
        pdf_path = input("Enter the path to your syllabus PDF file: ").strip()
        if os.path.exists(pdf_path):
            break
        else:
            print("❌ File not found. Please check the path and try again.")
    
    # Generate output file path
    base_name = os.path.splitext(pdf_path)[0]
    excel_path = f"{base_name}_syllabus_organized.xlsx"
    
    print(f"📁 Input: {pdf_path}")
    print(f"📁 Output: {excel_path}")
    print()
    
    # Convert the PDF
    print("🔍 Starting syllabus conversion...")
    df = syllabus_pdf_to_excel(pdf_path, excel_path)
    
    if not df.empty:
        print()
        print("✅ Conversion complete!")
        print(f"📊 Total rows: {len(df)}")
        print(f"📊 Total columns: {len(df.columns)}")
        print(f"📊 Pages processed: {df['Page'].nunique()}")
        
        if 'Unit' in df.columns:
            units = df[df['Unit'] != '']['Unit'].nunique()
            topics = df[df['Type'] == 'Topic'].shape[0]
            print(f"📊 Units detected: {units}")
            print(f"📊 Topics detected: {topics}")
        
        print()
        print("📋 First 10 rows of organized syllabus:")
        print(df.head(10).to_string(index=False))
        print()
        print(f"📥 Excel file saved to: {excel_path}")
        print("📋 Excel contains multiple sheets:")
        print("   - Syllabus_Organized: Complete organized syllabus")
        print("   - Unit_Summary: Summary by units")
        print("   - Topics_Only: Topics and subtopics only")
    else:
        print("❌ Conversion failed. No data extracted.")

if __name__ == "__main__":
    main() 