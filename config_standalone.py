#!/usr/bin/env python3
"""
Configuration file for Standalone PDF to Excel Converter
Update these paths according to your system setup.
"""

import os

# =============================================================================
# OCR CONFIGURATION
# =============================================================================

# Tesseract OCR Path
# Windows: Usually "C:\Program Files\Tesseract-OCR\tesseract.exe"
# Linux: Usually "/usr/bin/tesseract"
# macOS: Usually "/usr/local/bin/tesseract"
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Poppler Path (for PDF to image conversion)
# Windows: Download from https://github.com/oschwartz10612/poppler-windows/releases
# Linux: Install with "sudo apt-get install poppler-utils"
# macOS: Install with "brew install poppler"
POPPLER_PATH = r'C:\path\to\poppler\bin'  # Update this path

# =============================================================================
# PROCESSING CONFIGURATION
# =============================================================================

# OCR Settings
OCR_DPI = 300  # Higher DPI = better quality but slower processing
OCR_CONFIG = "--psm 6"  # Page segmentation mode

# Text extraction settings
MIN_TEXT_LENGTH = 50  # Minimum characters to consider as text-based PDF
MIN_LINE_LENGTH = 2   # Minimum characters for a line to be included

# =============================================================================
# OUTPUT CONFIGURATION
# =============================================================================

# Default output settings
DEFAULT_OUTPUT_DIR = "output"
INCLUDE_HEADERS = False  # Set to True to include column headers in Excel

# =============================================================================
# VALIDATION FUNCTIONS
# =============================================================================

def validate_tesseract_path():
    """Check if Tesseract is available at the configured path."""
    if os.path.exists(TESSERACT_PATH):
        return True, TESSERACT_PATH
    
    # Try common alternative paths
    alternative_paths = [
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r".\tesseract\tesseract.exe",
        "/usr/bin/tesseract",
        "/usr/local/bin/tesseract"
    ]
    
    for path in alternative_paths:
        if os.path.exists(path):
            return True, path
    
    return False, None

def validate_poppler_path():
    """Check if Poppler is available at the configured path."""
    if os.path.exists(POPPLER_PATH):
        return True, POPPLER_PATH
    
    # On Linux/macOS, poppler-utils might be in PATH
    if os.name != 'nt':  # Not Windows
        return True, None  # Assume it's in PATH
    
    return False, None

def print_configuration_status():
    """Print the current configuration status."""
    print("=" * 50)
    print("🔧 Configuration Status")
    print("=" * 50)
    
    # Check Tesseract
    tesseract_ok, tesseract_path = validate_tesseract_path()
    if tesseract_ok:
        print(f"✅ Tesseract OCR: {tesseract_path or 'Found in PATH'}")
    else:
        print("❌ Tesseract OCR: Not found")
        print("   Please install Tesseract OCR and update TESSERACT_PATH")
    
    # Check Poppler
    poppler_ok, poppler_path = validate_poppler_path()
    if poppler_ok:
        print(f"✅ Poppler: {poppler_path or 'Found in PATH'}")
    else:
        print("❌ Poppler: Not found")
        print("   Please install Poppler and update POPPLER_PATH")
    
    print(f"📊 OCR DPI: {OCR_DPI}")
    print(f"📝 Min text length: {MIN_TEXT_LENGTH}")
    print(f"📁 Default output: {DEFAULT_OUTPUT_DIR}")
    print("=" * 50)

if __name__ == "__main__":
    print_configuration_status()
