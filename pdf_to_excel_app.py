import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import os
import tempfile
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

def extract_text_pdfplumber(pdf_path):
    """Extract text using pdfplumber with improved syllabus structure preservation"""
    data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                print(f"🔄 Processing page {page_num}...")
                try:
                    # First try to extract complete text to preserve structure
                    text = page.extract_text()
                    if text and text.strip():
                        lines = text.split('\n')
                        for line_num, line in enumerate(lines, 1):
                            line = line.strip()
                            if line and len(line) > 2:  # Skip very short lines
                                data.append((f"Page {page_num}", line, line_num))
                        print(f"✅ Extracted {len([l for l in lines if l.strip()])} lines from page {page_num}")
                    else:
                        # Fallback to word extraction if no text found
                        words = page.extract_words()
                        if words:
                            # Group words into lines based on y-position
                            word_groups = {}
                            for word in words:
                                y_pos = round(word.get('top', 0), 2)
                                if y_pos not in word_groups:
                                    word_groups[y_pos] = []
                                word_groups[y_pos].append(word['text'])
                            
                            # Create lines from word groups
                            for line_num, (y_pos, word_list) in enumerate(sorted(word_groups.items()), 1):
                                line = ' '.join(word_list)
                                if line.strip():
                                    data.append((f"Page {page_num}", line, line_num))
                            print(f"✅ Extracted {len(word_groups)} lines from page {page_num} (word grouping)")
                        else:
                            raise ValueError("No content found")
                except Exception as e:
                    print(f"⚠️ pdfplumber failed on page {page_num}: {e}")
                    # Try alternative extraction method for this page
                    try:
                        text = page.extract_text()
                        if text:
                            lines = text.split('\n')
                            for line_num, line in enumerate(lines, 1):
                                line = line.strip()
                                if line and len(line) > 2:
                                    data.append((f"Page {page_num}", line, line_num))
                            print(f"✅ Alternative extraction: {len([l for l in lines if l.strip()])} lines from page {page_num}")
                        else:
                            raise ValueError("No text extracted")
                    except Exception as e2:
                        print(f"❌ Alternative extraction also failed on page {page_num}: {e2}")
                        continue
        
        if not data:
            raise ValueError("No data extracted from any page")
            
        df = pd.DataFrame(data, columns=["Page", "Content", "Line_Number"])
        print(f"✅ Extracted {len(df)} lines from {df['Page'].nunique()} pages")
        return df
    except Exception as e:
        print(f"Error extracting with pdfplumber: {e}")
        return pd.DataFrame()

def reconstruct_words_to_sentences(df):
    """Reconstruct individual words back into complete sentences and paragraphs"""
    if df.empty:
        return df
    
    # Check if content looks like individual words
    avg_content_length = df['Content'].str.len().mean()
    if avg_content_length > 10:  # If average content length is reasonable, return as is
        return df
    
    reconstructed_data = []
    current_page = ""
    current_line_words = []
    
    for _, row in df.iterrows():
        content = row['Content'].strip()
        page = row['Page']
        
        # If we're on a new page, start fresh
        if page != current_page:
            if current_line_words:
                # Join the previous line
                line = ' '.join(current_line_words)
                if line.strip():
                    reconstructed_data.append({
                        'Page': current_page,
                        'Content': line,
                        'Line_Number': len(reconstructed_data) + 1
                    })
            current_page = page
            current_line_words = []
        
        # Check if this word might be the start of a new sentence/paragraph
        # (capitalized words, numbers, or special markers)
        is_new_sentence = (
            content and len(content) > 0 and
            (content[0].isupper() or 
             content[0].isdigit() or 
             content in ['PAPER', 'UNIT', 'Subject:', 'Code', 'The', 'This', 'It'] or
             content.startswith('(') or
             content.startswith('•'))
        )
        
        if is_new_sentence and current_line_words:
            # Join the previous line
            line = ' '.join(current_line_words)
            if line.strip():
                reconstructed_data.append({
                    'Page': current_page,
                    'Content': line,
                    'Line_Number': len(reconstructed_data) + 1
                })
            current_line_words = []
        
        current_line_words.append(content)
    
    # Don't forget the last line
    if current_line_words:
        line = ' '.join(current_line_words)
        if line.strip():
            reconstructed_data.append({
                'Page': current_page,
                'Content': line,
                'Line_Number': len(reconstructed_data) + 1
            })
    
    return pd.DataFrame(reconstructed_data)

def organize_syllabus_structure(df):
    """Organize extracted content into proper syllabus structure with units and topics"""
    if df.empty:
        return df
    
    # First, try to reconstruct words into sentences if needed
    df = reconstruct_words_to_sentences(df)
    
    organized_data = []
    current_unit = ""
    current_topic = ""
    current_subtopic = ""
    
    import re
    
    for _, row in df.iterrows():
        content = row['Content']
        page = row['Page']
        line_num = row.get('Line_Number', 1)
        
        # Detect unit headers (patterns like "1.", "Unit 1:", "I.", etc.)
        unit_patterns = [
            r'^(\d+)\.?\s*(.+)$',  # "1. Unit Title" or "1 Unit Title"
            r'^Unit\s+(\d+)[:\.]?\s*(.+)$',  # "Unit 1: Title" or "Unit 1. Title"
            r'^([IVX]+)\.?\s*(.+)$',  # "I. Title" or "II. Title"
            r'^Chapter\s+(\d+)[:\.]?\s*(.+)$',  # "Chapter 1: Title"
            r'^PAPER[-\s]?(\d+)[:\.]?\s*(.+)$',  # "PAPER-I" or "PAPER 1: Title"
            r'^Subject:\s*(.+)$',  # "Subject: General Paper on Teaching & Research Aptitude"
            r'^Code\s+No\.\s*:\s*(\d+)\s*(.+)$',  # "Code No. : 00 PAPER-I"
        ]
        
        unit_detected = False
        for pattern in unit_patterns:
            match = re.match(pattern, content, re.IGNORECASE)
            if match:
                # Handle different patterns with different group structures
                if 'Subject:' in pattern:
                    # Subject pattern has only one group
                    unit_title = match.group(1).strip()
                    current_unit = f"Subject: {unit_title}"
                elif 'Code' in pattern:
                    # Code pattern has two groups
                    unit_num = match.group(1)
                    unit_title = match.group(2).strip()
                    current_unit = f"Code {unit_num}: {unit_title}"
                else:
                    # Standard patterns with two groups
                    unit_num = match.group(1)
                    unit_title = match.group(2).strip()
                    current_unit = f"Unit {unit_num}: {unit_title}"
                current_topic = ""
                current_subtopic = ""
                organized_data.append({
                    'Page': page,
                    'Line_Number': line_num,
                    'Unit': current_unit,
                    'Topic': '',
                    'Subtopic': '',
                    'Content': content,
                    'Type': 'Unit_Header',
                    'Level': 1
                })
                unit_detected = True
                break
        
        if unit_detected:
            continue
        
        # Detect topic headers (patterns like "a)", "•", "1.1", etc.)
        topic_patterns = [
            r'^([a-z]\)|•|\*|\-)\s*(.+)$',  # "a) Topic" or "• Topic"
            r'^(\d+\.\d+)\s*(.+)$',  # "1.1 Topic"
            r'^\((\d+)\)\s*(.+)$',  # "(a) Topic"
            r'^([A-Z])\.\s*(.+)$',  # "A. Topic"
            r'^The\s+(.+)$',  # "The main objective is to assess..."
            r'^This\s+(.+)$',  # "This test aims at..."
            r'^It\s+(.+)$',  # "It will assess..."
        ]
        
        topic_detected = False
        for pattern in topic_patterns:
            match = re.match(pattern, content, re.IGNORECASE)
            if match:
                # Handle different patterns with different group structures
                if pattern in [r'^The\s+(.+)$', r'^This\s+(.+)$', r'^It\s+(.+)$']:
                    # These patterns have only one group
                    topic_title = match.group(1).strip()
                    current_topic = f"Objective: {topic_title}"
                else:
                    # Standard patterns with two groups
                    topic_marker = match.group(1)
                    topic_title = match.group(2).strip()
                    current_topic = f"{topic_marker}. {topic_title}"
                current_subtopic = ""
                organized_data.append({
                    'Page': page,
                    'Line_Number': line_num,
                    'Unit': current_unit,
                    'Topic': current_topic,
                    'Subtopic': '',
                    'Content': content,
                    'Type': 'Topic_Header',
                    'Level': 2
                })
                topic_detected = True
                break
        
        if topic_detected:
            continue
        
        # Detect subtopics (indented or with different markers)
        subtopic_patterns = [
            r'^(\s{2,})(.+)$',  # Indented content
            r'^(\d+\.\d+\.\d+)\s*(.+)$',  # "1.1.1 Subtopic"
            r'^\((\d+)\)\s*(.+)$',  # "(1) Subtopic"
        ]
        
        subtopic_detected = False
        for pattern in subtopic_patterns:
            match = re.match(pattern, content, re.IGNORECASE)
            if match:
                subtopic_marker = match.group(1)
                subtopic_title = match.group(2).strip()
                current_subtopic = f"{subtopic_marker}. {subtopic_title}"
                organized_data.append({
                    'Page': page,
                    'Line_Number': line_num,
                    'Unit': current_unit,
                    'Topic': current_topic,
                    'Subtopic': current_subtopic,
                    'Content': content,
                    'Type': 'Subtopic_Header',
                    'Level': 3
                })
                subtopic_detected = True
                break
        
        if subtopic_detected:
            continue
        
        # Regular content
        if current_unit:
            organized_data.append({
                'Page': page,
                'Line_Number': line_num,
                'Unit': current_unit,
                'Topic': current_topic,
                'Subtopic': current_subtopic,
                'Content': content,
                'Type': 'Content',
                'Level': 4 if current_subtopic else (3 if current_topic else 2)
            })
        else:
            # If no unit detected yet, treat as general content
            organized_data.append({
                'Page': page,
                'Line_Number': line_num,
                'Unit': '',
                'Topic': '',
                'Subtopic': '',
                'Content': content,
                'Type': 'General_Content',
                'Level': 1
            })
    
    organized_df = pd.DataFrame(organized_data)
    
    # Fill forward unit and topic information
    organized_df['Unit'] = organized_df['Unit'].ffill()
    organized_df['Topic'] = organized_df['Topic'].ffill()
    organized_df['Subtopic'] = organized_df['Subtopic'].ffill()
    
    return organized_df

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

def hybrid_pdf_to_excel(pdf_path, excel_path):
    """Main function to convert PDF to Excel using hybrid approach with improved syllabus structure"""
    try:
        print("🔍 Analyzing PDF type...")
        
        if is_text_based(pdf_path):
            print("📄 Text-based PDF detected. Trying pdfplumber...")
            try:
                df = extract_text_pdfplumber(pdf_path)
                if df.empty:
                    raise ValueError("pdfplumber returned empty dataframe")
                
                # Organize into syllabus structure
                print("📚 Organizing content into syllabus structure...")
                df = organize_syllabus_structure(df)
                
            except Exception as e:
                print(f"❌ pdfplumber failed, switching to OCR. Reason: {e}")
                if tesseract_available:
                    df = extract_text_ocr(pdf_path)
                    df = organize_syllabus_structure(df)
                else:
                    print("❌ OCR not available and pdfplumber failed")
                    return pd.DataFrame()
        else:
            print("🖼️ Image-based PDF detected. Using OCR...")
            df = extract_text_ocr(pdf_path)
            df = organize_syllabus_structure(df)

        if df.empty:
            raise ValueError("No content extracted from PDF")
        
        # Save to Excel with multiple sheets for better organization
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # Main organized data - exclude Subtopic and Content columns
            columns_to_include = ['Page', 'Line_Number', 'Unit', 'Topic', 'Type', 'Level']
            available_columns = [col for col in columns_to_include if col in df.columns]
            main_df = df[available_columns].copy()
            main_df.to_excel(writer, sheet_name='Syllabus_Organized', index=False)
            
            # Summary by units
            if 'Unit' in df.columns and df['Unit'].str.len().sum() > 0:
                unit_summary = df[df['Unit'] != ''].groupby('Unit').agg({
                    'Topic': lambda x: x[x != ''].nunique(),
                    'Type': 'count'
                }).reset_index()
                unit_summary.columns = ['Unit', 'Topic_Count', 'Total_Items']
                unit_summary.to_excel(writer, sheet_name='Unit_Summary', index=False)
            
            # Topics only - exclude Subtopic and Content columns
            topics_df = df[df['Type'].isin(['Topic_Header', 'Subtopic_Header'])].copy()
            if not topics_df.empty:
                topic_columns = ['Page', 'Line_Number', 'Unit', 'Topic', 'Type', 'Level']
                available_topic_columns = [col for col in topic_columns if col in topics_df.columns]
                topics_df = topics_df[available_topic_columns]
                topics_df.to_excel(writer, sheet_name='Topics_Only', index=False)
            
            # Raw extracted data (for debugging) - exclude Content column
            raw_columns = ['Page', 'Line_Number']
            available_raw_columns = [col for col in raw_columns if col in df.columns]
            raw_df = df[available_raw_columns].copy()
            raw_df.to_excel(writer, sheet_name='Raw_Data', index=False)
        
        print(f"✅ Excel saved to: {excel_path}")
        return df
    except Exception as e:
        print(f"❌ Failed to extract PDF: {e}")
        return pd.DataFrame()

# --- Streamlit UI ---
st.set_page_config(page_title="Hybrid PDF to Excel Converter", layout="centered")
st.title("📄 Hybrid PDF to Excel Converter")
st.markdown("**Enhanced with Syllabus Structure Recognition** - Automatically organizes content by units and topics")

# Show OCR status
if tesseract_available:
    st.success("✅ OCR (Tesseract) is available")
else:
    st.warning("⚠️ OCR (Tesseract) is not available. Only text-based PDFs can be processed.")

# Add file upload with better error handling
uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"], help="Upload a PDF file to convert to Excel")

if uploaded_file:
    try:
        # Display file info
        file_details = {
            "Filename": uploaded_file.name,
            "File size": f"{uploaded_file.size / 1024:.1f} KB",
            "File type": uploaded_file.type
        }
        
        st.subheader("📋 File Information")
        for key, value in file_details.items():
            st.write(f"**{key}:** {value}")
        
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.read())
            tmp_pdf_path = tmp.name
        
        # Generate output filename
        base_name = os.path.splitext(uploaded_file.name)[0]
        output_excel_path = f"{base_name}_converted.xlsx"
        
        with st.spinner("🔍 Converting PDF to Excel..."):
            df = hybrid_pdf_to_excel(tmp_pdf_path, output_excel_path)
        
        if not df.empty:
            st.success("✅ Conversion complete!")
            
            # Display extraction stats
            st.subheader("📊 Extraction Statistics")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Rows", len(df))
            with col2:
                st.metric("Total Columns", len(df.columns))
            with col3:
                st.metric("Pages Processed", df['Page'].nunique())
            with col4:
                if 'Unit' in df.columns:
                    units = df[df['Unit'] != '']['Unit'].nunique()
                    st.metric("Units Detected", units)
                else:
                    st.metric("Units Detected", 0)
            
            # Show syllabus structure if available
            if 'Unit' in df.columns and 'Type' in df.columns:
                st.subheader("📚 Syllabus Structure")
                col1, col2 = st.columns(2)
                
                with col1:
                    # Unit summary
                    unit_counts = df[df['Type'] == 'Unit_Header']['Unit'].value_counts()
                    if not unit_counts.empty:
                        st.write("**Units Found:**")
                        for unit, count in unit_counts.head(5).items():
                            st.write(f"• {unit}")
                        if len(unit_counts) > 5:
                            st.write(f"... and {len(unit_counts) - 5} more")
                
                with col2:
                    # Topic summary
                    topic_counts = df[df['Type'] == 'Topic_Header']['Topic'].value_counts()
                    if not topic_counts.empty:
                        st.write("**Topics Found:**")
                        for topic, count in topic_counts.head(5).items():
                            st.write(f"• {topic}")
                        if len(topic_counts) > 5:
                            st.write(f"... and {len(topic_counts) - 5} more")
                
                # Show simplified column structure
                st.info("📋 **Simplified Output:** Columns C (Subtopic) and D (Content) have been removed for cleaner structure")
            
            st.subheader("📋 Extracted Data Preview:")
            st.dataframe(df.head(20), use_container_width=True)
            
            # Download button
            with open(output_excel_path, "rb") as f:
                st.download_button(
                    "📥 Download Excel File",
                    f,
                    file_name=output_excel_path,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # Clean up
            os.remove(output_excel_path)
        else:
            st.error("❌ Could not extract any content from this PDF.")
        
        # Clean up temporary PDF file
        os.remove(tmp_pdf_path)
        
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.info("💡 Try using the standalone script below if the web interface doesn't work.")

# Add standalone script section
st.markdown("---")
st.subheader("🔧 Standalone Script (Alternative)")
st.markdown("""
If the web interface doesn't work, you can use this standalone script directly:

```python
# Save this as standalone_pdf_converter.py and run it
import fitz
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import os

# Set your file paths here:
pdf_path = "your_pdf_file.pdf"  # Replace with your PDF path
excel_path = "output.xlsx"      # Replace with desired output path

# Run the conversion
df = hybrid_pdf_to_excel(pdf_path, excel_path)
if not df.empty:
    print(f"✅ Successfully converted {len(df)} rows to {excel_path}")
else:
    print("❌ Conversion failed")
```
""")

# Add direct file path input as alternative
st.subheader("📁 Direct File Path (Alternative)")
pdf_path_input = st.text_input("Enter PDF file path:", placeholder="C:/path/to/your/file.pdf")
if pdf_path_input and os.path.exists(pdf_path_input):
    if st.button("Convert using file path"):
        try:
            output_path = pdf_path_input.replace('.pdf', '_converted.xlsx')
            with st.spinner("🔍 Converting PDF to Excel..."):
                df = hybrid_pdf_to_excel(pdf_path_input, output_path)
            
            if not df.empty:
                st.success(f"✅ Conversion complete! File saved to: {output_path}")
                st.dataframe(df.head(10), use_container_width=True)
            else:
                st.error("❌ Could not extract any content from this PDF.")
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
elif pdf_path_input and not os.path.exists(pdf_path_input):
    st.error("❌ File not found. Please check the path.")
