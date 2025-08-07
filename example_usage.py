#!/usr/bin/env python3
"""
Example usage of the Standalone Hybrid PDF to Excel Converter
This script demonstrates various ways to use the converter.
"""

import os
import sys
from standalone_hybrid_converter import hybrid_pdf_to_excel

def example_single_conversion():
    """Example: Convert a single PDF file"""
    print("=" * 50)
    print("📄 Example 1: Single PDF Conversion")
    print("=" * 50)
    
    # Example file paths (update these to your actual files)
    pdf_path = "input_folder/sample.pdf"  # Update this path
    excel_path = "output_folder/sample_converted.xlsx"
    
    if os.path.exists(pdf_path):
        print(f"Converting: {pdf_path}")
        success = hybrid_pdf_to_excel(pdf_path, excel_path)
        
        if success:
            print(f"✅ Success! Output saved to: {excel_path}")
        else:
            print("❌ Conversion failed")
    else:
        print(f"⚠️ File not found: {pdf_path}")
        print("Please update the pdf_path variable with a valid PDF file")

def example_batch_conversion():
    """Example: Convert multiple PDF files in a folder"""
    print("\n" + "=" * 50)
    print("📁 Example 2: Batch PDF Conversion")
    print("=" * 50)
    
    input_folder = "input_folder"
    output_folder = "output_folder"
    
    # Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"📁 Created output folder: {output_folder}")
    
    # Find all PDF files in input folder
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"⚠️ No PDF files found in {input_folder}")
        return
    
    print(f"Found {len(pdf_files)} PDF file(s) to convert:")
    for pdf_file in pdf_files:
        print(f"  - {pdf_file}")
    
    # Convert each PDF
    successful_conversions = 0
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file)
        excel_file = pdf_file.replace('.pdf', '_converted.xlsx')
        excel_path = os.path.join(output_folder, excel_file)
        
        print(f"\n🔄 Converting: {pdf_file}")
        success = hybrid_pdf_to_excel(pdf_path, excel_path)
        
        if success:
            print(f"✅ Success: {excel_file}")
            successful_conversions += 1
        else:
            print(f"❌ Failed: {pdf_file}")
    
    print(f"\n📊 Batch conversion complete: {successful_conversions}/{len(pdf_files)} successful")

def example_with_error_handling():
    """Example: Convert with comprehensive error handling"""
    print("\n" + "=" * 50)
    print("🛡️ Example 3: Conversion with Error Handling")
    print("=" * 50)
    
    def safe_convert(pdf_path, excel_path):
        """Safely convert a PDF with error handling"""
        try:
            # Check if input file exists
            if not os.path.exists(pdf_path):
                print(f"❌ Input file not found: {pdf_path}")
                return False
            
            # Check if output directory exists
            output_dir = os.path.dirname(excel_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
                print(f"📁 Created output directory: {output_dir}")
            
            # Perform conversion
            success = hybrid_pdf_to_excel(pdf_path, excel_path)
            
            if success:
                # Verify output file was created
                if os.path.exists(excel_path):
                    file_size = os.path.getsize(excel_path)
                    print(f"✅ Conversion successful! File size: {file_size} bytes")
                    return True
                else:
                    print("❌ Output file was not created")
                    return False
            else:
                print("❌ Conversion failed")
                return False
                
        except Exception as e:
            print(f"❌ Unexpected error: {e}")
            return False
    
    # Example usage
    pdf_path = "input_folder/test.pdf"  # Update this path
    excel_path = "output_folder/test_converted.xlsx"
    
    success = safe_convert(pdf_path, excel_path)
    if success:
        print("🎉 All checks passed!")
    else:
        print("💡 Try checking the file paths and dependencies")

def example_configuration_check():
    """Example: Check configuration before conversion"""
    print("\n" + "=" * 50)
    print("⚙️ Example 4: Configuration Check")
    print("=" * 50)
    
    try:
        from config_standalone import print_configuration_status
        print_configuration_status()
    except ImportError:
        print("⚠️ Configuration file not found. Using default settings.")
        print("💡 Create config_standalone.py for better configuration management.")

def main():
    """Run all examples"""
    print("🔧 Standalone PDF to Excel Converter - Examples")
    print("=" * 60)
    
    # Check if standalone converter is available
    try:
        from standalone_hybrid_converter import hybrid_pdf_to_excel
    except ImportError as e:
        print(f"❌ Error importing converter: {e}")
        print("💡 Make sure standalone_hybrid_converter.py is in the same directory")
        return
    
    # Run examples
    example_configuration_check()
    example_single_conversion()
    example_batch_conversion()
    example_with_error_handling()
    
    print("\n" + "=" * 60)
    print("🎉 Examples completed!")
    print("💡 Update the file paths in this script to test with your own PDFs")
    print("=" * 60)

if __name__ == "__main__":
    main()
