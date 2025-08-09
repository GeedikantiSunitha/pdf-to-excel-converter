#!/usr/bin/env python3
"""
Web-based PDF to Excel Converter
Streamlit interface for uploading PDF files and downloading Excel output.
"""

import streamlit as st
import os
import tempfile
from enhanced_pdf_converter import enhanced_pdf_to_excel
import pandas as pd
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Enhanced PDF to Excel Converter",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .feature-box {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.markdown('<h1 class="main-header">🔧 Enhanced PDF to Excel Converter</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; font-size: 1.2rem; color: #666;">Upload PDF files and get Excel output with complete data preservation</p>', unsafe_allow_html=True)
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # OCR Settings
        st.subheader("OCR Settings")
        ocr_dpi = st.slider("OCR DPI", min_value=200, max_value=600, value=300, step=50,
                           help="Higher DPI = better quality but slower processing")
        
        # Processing Options
        st.subheader("Processing Options")
        preserve_tables = st.checkbox("Preserve Table Structure", value=True,
                                    help="Maintain original table formatting")
        
        include_summary = st.checkbox("Include Page Summary", value=True,
                                    help="Add summary sheet with page statistics")
        
        # File Upload Settings
        st.subheader("File Settings")
        max_file_size = st.number_input("Max File Size (MB)", min_value=1, max_value=100, value=50)
        
        st.markdown("---")
        st.markdown("### 📊 Current Status")
        
        # Check dependencies
        try:
            from config_standalone import print_configuration_status
            with st.expander("🔧 System Configuration"):
                st.code("Configuration check will be shown here")
        except:
            st.warning("⚠️ Configuration file not found")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 Upload PDF File")
        
        # File uploader
        uploaded_file = st.file_uploader(
            "Choose a PDF file",
            type=['pdf'],
            help="Upload a PDF file to convert to Excel format"
        )
        
        if uploaded_file is not None:
            # Display file information
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.write(f"**File Name:** {uploaded_file.name}")
            st.write(f"**File Size:** {uploaded_file.size / 1024 / 1024:.2f} MB")
            st.write(f"**File Type:** {uploaded_file.type}")
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Conversion options
            st.subheader("🔄 Conversion Options")
            
            col_a, col_b = st.columns(2)
            with col_a:
                output_filename = st.text_input(
                    "Output Filename",
                    value=f"{os.path.splitext(uploaded_file.name)[0]}_converted.xlsx",
                    help="Name for the output Excel file"
                )
            
            with col_b:
                include_headers = st.checkbox("Include Headers", value=False,
                                            help="Add column headers to Excel output")
            
            # Convert button
            if st.button("🚀 Convert to Excel", type="primary", use_container_width=True):
                with st.spinner("🔄 Converting PDF to Excel..."):
                    try:
                        # Create temporary file
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                            tmp_file.write(uploaded_file.getvalue())
                            tmp_pdf_path = tmp_file.name
                        
                        # Create temporary output path
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_output:
                            tmp_excel_path = tmp_output.name
                        
                        # Perform conversion
                        success = enhanced_pdf_to_excel(tmp_pdf_path, tmp_excel_path)
                        
                        if success:
                            # Read the Excel file for download
                            with open(tmp_excel_path, 'rb') as f:
                                excel_data = f.read()
                            
                            # Clean up temporary files
                            os.unlink(tmp_pdf_path)
                            os.unlink(tmp_excel_path)
                            
                            # Success message
                            st.markdown('<div class="success-box">', unsafe_allow_html=True)
                            st.success("✅ Conversion completed successfully!")
                            st.markdown('</div>', unsafe_allow_html=True)
                            
                            # Download button
                            st.download_button(
                                label="📥 Download Excel File",
                                data=excel_data,
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            
                            # Show file info
                            st.info(f"📁 File saved as: {output_filename}")
                            
                        else:
                            st.error("❌ Conversion failed. Please check the file and try again.")
                            # Clean up temporary files
                            if os.path.exists(tmp_pdf_path):
                                os.unlink(tmp_pdf_path)
                            if os.path.exists(tmp_excel_path):
                                os.unlink(tmp_excel_path)
                    
                    except Exception as e:
                        st.error(f"❌ Error during conversion: {str(e)}")
                        # Clean up temporary files
                        if 'tmp_pdf_path' in locals() and os.path.exists(tmp_pdf_path):
                            os.unlink(tmp_pdf_path)
                        if 'tmp_excel_path' in locals() and os.path.exists(tmp_excel_path):
                            os.unlink(tmp_excel_path)
    
    with col2:
        st.header("📋 Features")
        
        st.markdown('<div class="feature-box">', unsafe_allow_html=True)
        st.write("🔍 **Smart Detection**")
        st.write("Automatically detects text-based vs image-based PDFs")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="feature-box">', unsafe_allow_html=True)
        st.write("📄 **Full Data Preservation**")
        st.write("No text truncation or data loss")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="feature-box">', unsafe_allow_html=True)
        st.write("📊 **Table Structure**")
        st.write("Preserves original table formatting and columns")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="feature-box">', unsafe_allow_html=True)
        st.write("🖼️ **Advanced OCR**")
        st.write("Multiple OCR configurations for best accuracy")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="feature-box">', unsafe_allow_html=True)
        st.write("📋 **Multiple Sheets**")
        st.write("Organized Excel output with separate sheets")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>🔧 Enhanced PDF to Excel Converter | Made with ❤️ for efficient data conversion</p>
        <p>Supports text-based PDFs, scanned documents, and complex tables</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
