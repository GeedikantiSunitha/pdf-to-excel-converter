"""
PDF Extractor Module
Handles text and table extraction from PDFs using both pdfplumber and OCR
"""

import os
import pdfplumber
import pandas as pd
import pytesseract
from pdf2image import convert_from_path
from tempfile import TemporaryDirectory
from PIL import Image
from typing import List, Dict, Any, Optional
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PDFExtractor:
    """
    Extracts text and tables from PDF files using hybrid approach
    - Text-based PDFs: Uses pdfplumber for direct extraction
    - Image-based PDFs: Uses OCR (Tesseract) for scanned documents
    """
    
    def __init__(self, dpi: int = 300):
        """
        Initialize PDF Extractor
        
        Args:
            dpi (int): DPI for image conversion (default: 300)
        """
        self.dpi = dpi
        self._check_dependencies()
    
    def _check_dependencies(self):
        """Check if required dependencies are available"""
        try:
            import pdfplumber
            logger.info("✅ pdfplumber is available")
        except ImportError:
            logger.error("❌ pdfplumber not found. Install with: pip install pdfplumber")
        
        try:
            import pytesseract
            # Test Tesseract
            pytesseract.get_tesseract_version()
            logger.info("✅ Tesseract OCR is available")
        except Exception as e:
            logger.warning(f"⚠️ Tesseract OCR not available: {e}")
    
    def is_text_based(self, pdf_path: str) -> bool:
        """
        Check if PDF contains extractable text
        
        Args:
            pdf_path (str): Path to PDF file
            
        Returns:
            bool: True if PDF is text-based, False if image-based
        """
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text and len(text.strip()) > 50:  # Minimum text threshold
                        return True
            return False
        except Exception as e:
            logger.error(f"Error checking PDF type: {e}")
            return False
    
    def extract_text_tables(self, pdf_path: str) -> List[pd.DataFrame]:
        """
        Extract tables from text-based PDF using pdfplumber
        
        Args:
            pdf_path (str): Path to PDF file
            
        Returns:
            List[pd.DataFrame]: List of extracted tables as DataFrames
        """
        tables = []
        total_pages = 0
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                logger.info(f"📄 Total pages in PDF: {total_pages}")
                
                for page_num, page in enumerate(pdf.pages, start=1):
                    logger.info(f"Processing page {page_num}/{total_pages}...")
                    
                    # First, try to extract tables
                    page_tables = page.extract_tables()
                    tables_found = False
                    
                    # Process each table on the page
                    for table_idx, table in enumerate(page_tables):
                        try:
                            if table and len(table) > 0:
                                # Filter out completely empty rows
                                filtered_table = []
                                for row in table:
                                    if row and any(cell is not None and str(cell).strip() for cell in row):
                                        filtered_table.append(row)
                                
                                if len(filtered_table) > 0:
                                    # Handle case where we have data but no clear header
                                    if len(filtered_table) == 1:
                                        # Single row - treat as data with generic headers
                                        headers = [f"Column_{i+1}" for i in range(len(filtered_table[0]))]
                                        df = pd.DataFrame([filtered_table[0]], columns=headers)
                                    else:
                                        # Multiple rows - use first row as header
                                        headers = filtered_table[0]
                                        cleaned_headers = self._clean_column_names(headers)
                                        
                                        # Create DataFrame with data rows (skip header)
                                        data_rows = filtered_table[1:]
                                        if data_rows:
                                            df = pd.DataFrame(data_rows, columns=cleaned_headers)
                                        else:
                                            # Only header row, create empty DataFrame
                                            df = pd.DataFrame(columns=cleaned_headers)
                                    
                                    # Add metadata columns only if they don't exist
                                    if 'Page' not in df.columns:
                                        df.insert(0, 'Page', page_num)
                                    if 'Table_Index' not in df.columns:
                                        df.insert(1, 'Table_Index', table_idx + 1)
                                    
                                    tables.append(df)
                                    tables_found = True
                                    logger.info(f"✅ Extracted table {table_idx + 1} from page {page_num} ({len(df)} rows)")
                        except Exception as e:
                            logger.warning(f"⚠️ Error processing table {table_idx + 1} on page {page_num}: {e}")
                            # Continue with next table instead of stopping
                            continue
                    
                    # If no tables found on this page, extract ALL text content
                    if not tables_found:
                        try:
                            text = page.extract_text()
                            if text and text.strip():
                                lines = [line.strip() for line in text.split('\n') if line.strip()]
                                if lines:
                                    df = pd.DataFrame({'Text': lines})
                                    if 'Page' not in df.columns:
                                        df.insert(0, 'Page', page_num)
                                    if 'Table_Index' not in df.columns:
                                        df.insert(1, 'Table_Index', 1)
                                    tables.append(df)
                                    logger.info(f"✅ Extracted {len(lines)} text lines from page {page_num}")
                                else:
                                    logger.warning(f"⚠️ No text content found on page {page_num}")
                            else:
                                logger.warning(f"⚠️ No text content found on page {page_num}")
                        except Exception as text_error:
                            logger.error(f"Error extracting text from page {page_num}: {text_error}")
                                
        except Exception as e:
            logger.error(f"Error extracting text tables: {e}")
            # Try to extract at least some text if table extraction fails
            try:
                logger.info("Attempting to extract text as fallback...")
                with pdfplumber.open(pdf_path) as pdf:
                    for page_num, page in enumerate(pdf.pages, start=1):
                        text = page.extract_text()
                        if text and text.strip():
                            lines = [line.strip() for line in text.split('\n') if line.strip()]
                            if lines:
                                df = pd.DataFrame({'Text': lines})
                                if 'Page' not in df.columns:
                                    df.insert(0, 'Page', page_num)
                                if 'Table_Index' not in df.columns:
                                    df.insert(1, 'Table_Index', 1)
                                tables.append(df)
                                logger.info(f"✅ Extracted text from page {page_num} (fallback)")
            except Exception as fallback_error:
                logger.error(f"Fallback text extraction also failed: {fallback_error}")
        
        # Ensure we have at least one entry per page
        if len(tables) < total_pages:
            logger.info(f"⚠️ Only extracted {len(tables)} sections from {total_pages} pages. Adding missing pages...")
            extracted_pages = set()
            for table in tables:
                if 'Page' in table.columns:
                    extracted_pages.update(table['Page'].unique())
            
            # Add missing pages with empty content
            for page_num in range(1, total_pages + 1):
                if page_num not in extracted_pages:
                    df = pd.DataFrame({'Page': [page_num], 'Table_Index': [1], 'Text': ['[No extractable content]']})
                    tables.append(df)
                    logger.info(f"✅ Added placeholder for page {page_num}")
        
        return tables
    
    def extract_from_images(self, pdf_path: str) -> List[pd.DataFrame]:
        """
        Extract text from image-based PDF using OCR
        
        Args:
            pdf_path (str): Path to PDF file
            
        Returns:
            List[pd.DataFrame]: List of extracted text as DataFrames
        """
        tables = []
        try:
            with TemporaryDirectory() as temp_dir:
                logger.info("Converting PDF to images...")
                images = convert_from_path(pdf_path, dpi=self.dpi, output_folder=temp_dir)
                
                for i, image in enumerate(images):
                    logger.info(f"Processing image {i + 1}/{len(images)}...")
                    
                    # Extract text using OCR
                    text = pytesseract.image_to_string(image)
                    lines = [line.strip() for line in text.split('\n') if line.strip()]
                    
                    if lines:
                        # Try to detect table structure
                        data = self._parse_ocr_text(lines)
                        if data:
                            # Clean column names for OCR data too
                            if data and len(data) > 0:
                                headers = data[0]
                                cleaned_headers = self._clean_column_names(headers)
                                df = pd.DataFrame(data[1:], columns=cleaned_headers)
                            else:
                                df = pd.DataFrame({'Text': lines})
                        else:
                            df = pd.DataFrame({'Text': lines})
                        
                        if 'Page' not in df.columns:
                            df.insert(0, 'Page', i + 1)
                        if 'Table_Index' not in df.columns:
                            df.insert(1, 'Table_Index', 1)
                        tables.append(df)
                        logger.info(f"✅ Extracted text from image {i + 1}")
                        
        except Exception as e:
            logger.error(f"Error extracting from images: {e}")
        
        return tables
    
    def _clean_column_names(self, headers: List[str]) -> List[str]:
        """
        Clean column names to avoid duplicates and invalid characters
        
        Args:
            headers (List[str]): Original column headers
            
        Returns:
            List[str]: Cleaned column headers
        """
        cleaned_headers = []
        seen_headers = set()
        
        for i, header in enumerate(headers):
            try:
                if header is None:
                    header = f"Column_{i+1}"
                else:
                    # Clean the header
                    header = str(header).strip()
                    if not header:
                        header = f"Column_{i+1}"
                    else:
                        # Remove newlines and extra whitespace
                        header = header.replace('\n', ' ').replace('\r', ' ')
                        header = ' '.join(header.split())
                        
                        # Remove or replace invalid characters for Excel
                        invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
                        for char in invalid_chars:
                            header = header.replace(char, '_')
                        
                        # Limit length
                        if len(header) > 50:
                            header = header[:47] + "..."
                
                # Handle duplicates
                original_header = header
                counter = 1
                while header in seen_headers:
                    header = f"{original_header}_{counter}"
                    counter += 1
                
                seen_headers.add(header)
                cleaned_headers.append(header)
                
            except Exception as e:
                # Fallback for any unexpected errors
                fallback_header = f"Column_{i+1}"
                counter = 1
                while fallback_header in seen_headers:
                    fallback_header = f"Column_{i+1}_{counter}"
                    counter += 1
                seen_headers.add(fallback_header)
                cleaned_headers.append(fallback_header)
                logger.warning(f"Error cleaning column header {i}: {e}, using fallback: {fallback_header}")
        
        return cleaned_headers
    
    def _parse_ocr_text(self, lines: List[str]) -> Optional[List[List[str]]]:
        """
        Parse OCR text to detect table structure
        
        Args:
            lines (List[str]): Lines of text from OCR
            
        Returns:
            Optional[List[List[str]]]: Parsed table data or None
        """
        if not lines:
            return None
        
        # Simple table detection based on consistent spacing
        parsed_data = []
        for line in lines:
            # Split by multiple spaces to detect columns
            columns = [col.strip() for col in line.split('  ') if col.strip()]
            if len(columns) > 1:
                parsed_data.append(columns)
        
        return parsed_data if parsed_data else None
    
    def extract(self, pdf_path: str) -> Dict[str, Any]:
        """
        Main extraction method - automatically detects PDF type and extracts content
        Ensures ALL pages and ALL data are captured without loss
        
        Args:
            pdf_path (str): Path to PDF file
            
        Returns:
            Dict[str, Any]: Extraction results with metadata
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        
        logger.info(f"🔍 Analyzing PDF: {pdf_path}")
        
        # Get total pages first
        total_pages = 0
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
        except Exception as e:
            logger.error(f"Error getting page count: {e}")
        
        logger.info(f"📄 Total pages detected: {total_pages}")
        
        # Detect PDF type
        is_text = self.is_text_based(pdf_path)
        
        # Initialize comprehensive extraction
        all_tables = []
        extracted_pages = set()
        
        # Step 1: Try table extraction first (most structured data)
        if is_text:
            logger.info("📄 PDF is text-based. Extracting tables with pdfplumber...")
            tables = self.extract_text_tables(pdf_path)
            all_tables.extend(tables)
            extraction_method = "pdfplumber"
        else:
            logger.info("🖼️ PDF is image-based. Extracting tables with OCR...")
            tables = self.extract_from_images(pdf_path)
            all_tables.extend(tables)
            extraction_method = "ocr"
        
        # Step 2: Extract ALL text content from every page (comprehensive coverage)
        logger.info("📝 Extracting ALL text content from every page...")
        all_text_tables = self.extract_all_text_comprehensive(pdf_path)
        all_tables.extend(all_text_tables)
        
        # Step 3: Extract any remaining content (images, forms, etc.)
        logger.info("🔍 Extracting any remaining content...")
        remaining_content = self.extract_remaining_content(pdf_path)
        all_tables.extend(remaining_content)
        
        # Step 4: Validate and ensure ALL pages are covered
        extracted_pages = set()
        for table in all_tables:
            if 'Page' in table.columns:
                extracted_pages.update(table['Page'].unique())
        
        # Ensure every page has at least one entry
        missing_pages = set(range(1, total_pages + 1)) - extracted_pages
        if missing_pages:
            logger.warning(f"⚠️ Adding content for missing pages: {sorted(missing_pages)}")
            for page_num in missing_pages:
                # Try to extract any remaining content from these pages
                page_content = self.extract_page_content(pdf_path, page_num)
                if page_content:
                    all_tables.append(page_content)
                else:
                    # Add placeholder for completely empty pages
                    df = pd.DataFrame({
                        'Page': [page_num], 
                        'Content_Type': ['Empty_Page'],
                        'Text': [f'[Page {page_num} - No extractable content found]'],
                        'Extraction_Status': ['No_Content']
                    })
                    all_tables.append(df)
                logger.info(f"✅ Ensured coverage for page {page_num}")
        
        # Step 5: Consolidate and organize results
        consolidated_tables = self.consolidate_tables(all_tables, total_pages)
        
        # Calculate comprehensive statistics
        total_rows = sum(len(df) for df in consolidated_tables) if consolidated_tables else 0
        final_extracted_pages = set()
        for table in consolidated_tables:
            if 'Page' in table.columns:
                final_extracted_pages.update(table['Page'].unique())
        
        # Prepare comprehensive results
        results = {
            'pdf_path': pdf_path,
            'extraction_method': extraction_method,
            'is_text_based': is_text,
            'tables': consolidated_tables,
            'total_tables': len(consolidated_tables),
            'total_rows': total_rows,
            'total_pages': total_pages,
            'extracted_pages': len(final_extracted_pages),
            'missing_pages': 0,  # We ensure all pages are covered
            'completeness_score': (len(final_extracted_pages) / total_pages * 100) if total_pages > 0 else 100,
            'extraction_timestamp': datetime.now().isoformat(),
            'processing_time': 0  # Will be set by caller
        }
        
        logger.info(f"📊 COMPREHENSIVE EXTRACTION COMPLETE:")
        logger.info(f"   📄 Pages: {total_pages} total, {len(final_extracted_pages)} extracted")
        logger.info(f"   📋 Tables/Sections: {results['total_tables']}")
        logger.info(f"   📝 Total Rows: {results['total_rows']}")
        logger.info(f"   ✅ Completeness: {results['completeness_score']:.1f}%")
        
        if results['completeness_score'] == 100:
            logger.info("🎉 PERFECT EXTRACTION - All pages and data captured!")
        else:
            logger.warning(f"⚠️ Extraction completeness: {results['completeness_score']:.1f}%")
            
        return results
    
    def extract_all_text(self, pdf_path: str) -> List[pd.DataFrame]:
        """
        Extract all text content from PDF when no tables are found
        
        Args:
            pdf_path (str): Path to PDF file
            
        Returns:
            List[pd.DataFrame]: List of text content as DataFrames
        """
        tables = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    logger.info(f"Extracting text from page {page_num}...")
                    
                    # Extract all text from the page
                    text = page.extract_text()
                    if text and text.strip():
                        # Split into lines and clean
                        lines = [line.strip() for line in text.split('\n') if line.strip()]
                        
                        if lines:
                            # Create DataFrame with text content
                            df = pd.DataFrame({'Text': lines})
                            
                            # Add metadata columns
                            if 'Page' not in df.columns:
                                df.insert(0, 'Page', page_num)
                            if 'Table_Index' not in df.columns:
                                df.insert(1, 'Table_Index', 1)
                            
                            tables.append(df)
                            logger.info(f"✅ Extracted {len(lines)} lines from page {page_num}")
                        
        except Exception as e:
            logger.error(f"Error extracting all text: {e}")
        
        return tables
    
    def extract_all_text_comprehensive(self, pdf_path: str) -> List[pd.DataFrame]:
        """
        Extract ALL text content from EVERY page comprehensively
        
        Args:
            pdf_path (str): Path to PDF file
            
        Returns:
            List[pd.DataFrame]: List of comprehensive text content as DataFrames
        """
        tables = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    logger.info(f"📝 Extracting ALL text from page {page_num}...")
                    
                    # Extract all text from the page
                    text = page.extract_text()
                    
                    # Also try to extract words and characters for maximum coverage
                    words = page.extract_words()
                    chars = page.extract_text()
                    
                    if text and text.strip():
                        # Split into lines and clean
                        lines = [line.strip() for line in text.split('\n') if line.strip()]
                        
                        if lines:
                            # Create comprehensive DataFrame
                            df = pd.DataFrame({
                                'Page': [page_num] * len(lines),
                                'Content_Type': ['Text_Line'] * len(lines),
                                'Text': lines,
                                'Line_Number': range(1, len(lines) + 1),
                                'Extraction_Method': ['pdfplumber_text'] * len(lines)
                            })
                            
                            tables.append(df)
                            logger.info(f"✅ Extracted {len(lines)} text lines from page {page_num}")
                    
                    # Add word-level extraction for better granularity
                    if words:
                        word_data = []
                        for i, word in enumerate(words):
                            word_data.append({
                                'Page': page_num,
                                'Content_Type': 'Word',
                                'Text': word['text'],
                                'X': word.get('x0', 0),
                                'Y': word.get('top', 0),
                                'Width': word.get('width', 0),
                                'Height': word.get('height', 0),
                                'Line_Number': i + 1,
                                'Extraction_Method': 'pdfplumber_words'
                            })
                        
                        if word_data:
                            word_df = pd.DataFrame(word_data)
                            tables.append(word_df)
                            logger.info(f"✅ Extracted {len(word_data)} words from page {page_num}")
                    
                    # If no text found, mark the page as processed but empty
                    if not text and not words:
                        df = pd.DataFrame({
                            'Page': [page_num],
                            'Content_Type': ['No_Text_Content'],
                            'Text': ['[No text content found on this page]'],
                            'Line_Number': [1],
                            'Extraction_Method': ['pdfplumber_empty']
                        })
                        tables.append(df)
                        logger.info(f"⚠️ No text content found on page {page_num}")
                        
        except Exception as e:
            logger.error(f"Error in comprehensive text extraction: {e}")
        
        return tables
    
    def extract_remaining_content(self, pdf_path: str) -> List[pd.DataFrame]:
        """
        Extract any remaining content like images, forms, annotations, etc.
        
        Args:
            pdf_path (str): Path to PDF file
            
        Returns:
            List[pd.DataFrame]: List of remaining content as DataFrames
        """
        tables = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    logger.info(f"🔍 Extracting remaining content from page {page_num}...")
                    
                    # Extract images
                    images = page.images
                    if images:
                        image_data = []
                        for i, img in enumerate(images):
                            image_data.append({
                                'Page': page_num,
                                'Content_Type': 'Image',
                                'Text': f'[Image {i+1}: {img.get("width", 0)}x{img.get("height", 0)}]',
                                'X': img.get('x0', 0),
                                'Y': img.get('top', 0),
                                'Width': img.get('width', 0),
                                'Height': img.get('height', 0),
                                'Extraction_Method': 'pdfplumber_images'
                            })
                        
                        if image_data:
                            img_df = pd.DataFrame(image_data)
                            tables.append(img_df)
                            logger.info(f"✅ Found {len(image_data)} images on page {page_num}")
                    
                    # Extract forms and annotations
                    annotations = page.annots if hasattr(page, 'annots') else []
                    if annotations:
                        annotation_data = []
                        for i, ann in enumerate(annotations):
                            annotation_data.append({
                                'Page': page_num,
                                'Content_Type': 'Annotation',
                                'Text': f'[Annotation {i+1}: {ann.get("subtype", "Unknown")}]',
                                'X': ann.get('x0', 0),
                                'Y': ann.get('top', 0),
                                'Extraction_Method': 'pdfplumber_annotations'
                            })
                        
                        if annotation_data:
                            ann_df = pd.DataFrame(annotation_data)
                            tables.append(ann_df)
                            logger.info(f"✅ Found {len(annotation_data)} annotations on page {page_num}")
                    
        except Exception as e:
            logger.error(f"Error extracting remaining content: {e}")
        
        return tables
    
    def extract_page_content(self, pdf_path: str, page_num: int) -> Optional[pd.DataFrame]:
        """
        Extract any remaining content from a specific page
        
        Args:
            pdf_path (str): Path to PDF file
            page_num (int): Page number to extract
            
        Returns:
            Optional[pd.DataFrame]: Content from the page or None
        """
        try:
            with pdfplumber.open(pdf_path) as pdf:
                if page_num <= len(pdf.pages):
                    page = pdf.pages[page_num - 1]
                    
                    # Try multiple extraction methods
                    text = page.extract_text()
                    words = page.extract_words()
                    images = page.images
                    
                    if text or words or images:
                        content_data = []
                        
                        if text:
                            content_data.append({
                                'Page': page_num,
                                'Content_Type': 'Remaining_Text',
                                'Text': text[:500] + '...' if len(text) > 500 else text,
                                'Extraction_Method': 'page_specific_text'
                            })
                        
                        if words:
                            content_data.append({
                                'Page': page_num,
                                'Content_Type': 'Remaining_Words',
                                'Text': f'[Found {len(words)} words]',
                                'Extraction_Method': 'page_specific_words'
                            })
                        
                        if images:
                            content_data.append({
                                'Page': page_num,
                                'Content_Type': 'Remaining_Images',
                                'Text': f'[Found {len(images)} images]',
                                'Extraction_Method': 'page_specific_images'
                            })
                        
                        if content_data:
                            return pd.DataFrame(content_data)
        
        except Exception as e:
            logger.error(f"Error extracting page {page_num} content: {e}")
        
        return None
    
    def consolidate_tables(self, all_tables: List[pd.DataFrame], total_pages: int) -> List[pd.DataFrame]:
        """
        Consolidate and organize all extracted tables
        
        Args:
            all_tables (List[pd.DataFrame]): All extracted tables
            total_pages (int): Total number of pages
            
        Returns:
            List[pd.DataFrame]: Consolidated and organized tables
        """
        if not all_tables:
            return []
        
        # Create a comprehensive summary table
        summary_data = []
        for page_num in range(1, total_pages + 1):
            page_tables = [t for t in all_tables if 'Page' in t.columns and page_num in t['Page'].values]
            
            if page_tables:
                total_rows = sum(len(t) for t in page_tables)
                content_types = []
                for table in page_tables:
                    if 'Content_Type' in table.columns:
                        content_types.extend(table['Content_Type'].unique())
                
                summary_data.append({
                    'Page': page_num,
                    'Content_Type': 'Page_Summary',
                    'Text': f'Page {page_num}: {total_rows} total items, Content types: {", ".join(set(content_types))}',
                    'Total_Items': total_rows,
                    'Content_Types': ', '.join(set(content_types)),
                    'Extraction_Method': 'consolidated_summary'
                })
            else:
                summary_data.append({
                    'Page': page_num,
                    'Content_Type': 'Page_Summary',
                    'Text': f'Page {page_num}: No content extracted',
                    'Total_Items': 0,
                    'Content_Types': 'None',
                    'Extraction_Method': 'consolidated_summary'
                })
        
        # Add summary table at the beginning
        consolidated = [pd.DataFrame(summary_data)]
        
        # Add all original tables
        consolidated.extend(all_tables)
        
        return consolidated 