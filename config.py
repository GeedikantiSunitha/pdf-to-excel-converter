"""
Configuration settings for PDF to Excel Converter
"""

import os
from pathlib import Path

# Application settings
APP_NAME = "PDF to Excel Converter"
APP_VERSION = "1.0.0"
APP_DESCRIPTION = "Convert PDF files to Excel with advanced text and table extraction"

# File settings
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100 MB
SUPPORTED_FORMATS = ['.pdf']
TEMP_DIR = Path("temp")
OUTPUT_DIR = Path("output")

# OCR settings
DEFAULT_DPI = 300
MIN_DPI = 150
MAX_DPI = 600
OCR_TIMEOUT = 300  # 5 minutes

# Extraction settings
MIN_TEXT_THRESHOLD = 50
MAX_COLUMN_WIDTH = 50
DEFAULT_PAGE_LIMIT = 100

# Excel settings
DEFAULT_SHEET_NAME = "Extracted_Data"
MAX_SHEET_NAME_LENGTH = 31
DEFAULT_HEADER_STYLE = {
    'font_color': 'FFFFFF',
    'background_color': '366092',
    'bold': True,
    'align': 'center'
}

# UI settings
PAGE_TITLE = "PDF to Excel Converter"
PAGE_ICON = "📊"
LAYOUT = "wide"
INITIAL_SIDEBAR_STATE = "expanded"

# Logging settings
LOG_LEVEL = "INFO"
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_FILE = "app.log"

# Performance settings
CHUNK_SIZE = 4096
MAX_WORKERS = 4
CACHE_TTL = 3600  # 1 hour

# Security settings
ALLOWED_EXTENSIONS = {'pdf'}
MAX_CONTENT_LENGTH = MAX_FILE_SIZE

# Development settings
DEBUG = os.getenv('DEBUG', 'False').lower() == 'true'
TESTING = os.getenv('TESTING', 'False').lower() == 'true'

# Create necessary directories
TEMP_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True) 