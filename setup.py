#!/usr/bin/env python3
"""
Setup script for PDF to Excel Converter
Installs dependencies and performs initial configuration
"""

import subprocess
import sys
import os
import platform

def check_python_version():
    """Check if Python version is compatible"""
    if sys.version_info < (3, 8):
        print("❌ Python 3.8 or higher is required")
        print(f"Current version: {sys.version}")
        return False
    print(f"✅ Python {sys.version.split()[0]} detected")
    return True

def install_requirements():
    """Install Python dependencies"""
    print("📦 Installing Python dependencies...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ Dependencies installed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install dependencies: {e}")
        return False

def check_tesseract():
    """Check if Tesseract is installed"""
    print("🔍 Checking Tesseract installation...")
    try:
        import pytesseract
        version = pytesseract.get_tesseract_version()
        print(f"✅ Tesseract {version} found")
        return True
    except Exception as e:
        print("⚠️ Tesseract not found or not properly configured")
        print("Please install Tesseract OCR:")
        
        system = platform.system().lower()
        if system == "windows":
            print("Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki")
            print("Add to PATH: C:\\Program Files\\Tesseract-OCR")
        elif system == "darwin":
            print("macOS: brew install tesseract")
        elif system == "linux":
            print("Ubuntu/Debian: sudo apt-get install tesseract-ocr")
        else:
            print("Please install Tesseract OCR for your system")
        
        return False

def create_directories():
    """Create necessary directories"""
    print("📁 Creating directories...")
    directories = ["output", "assets", "utils"]
    
    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"✅ Created {directory}/ directory")
        else:
            print(f"✅ {directory}/ directory already exists")

def test_imports():
    """Test if all modules can be imported"""
    print("🧪 Testing imports...")
    try:
        import streamlit
        import pdfplumber
        import pandas
        import openpyxl
        import pytesseract
        import pdf2image
        from PIL import Image
        print("✅ All Python modules imported successfully")
        return True
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False

def main():
    """Main setup function"""
    print("🚀 Setting up PDF to Excel Converter...")
    print("=" * 50)
    
    # Check Python version
    if not check_python_version():
        sys.exit(1)
    
    # Install requirements
    if not install_requirements():
        sys.exit(1)
    
    # Test imports
    if not test_imports():
        sys.exit(1)
    
    # Check Tesseract
    check_tesseract()
    
    # Create directories
    create_directories()
    
    print("\n" + "=" * 50)
    print("🎉 Setup completed successfully!")
    print("\nTo run the application:")
    print("streamlit run app.py")
    print("\nThen open your browser to: http://localhost:8501")

if __name__ == "__main__":
    main() 