"""
Utility functions for PDF to Excel Converter
"""

import os
import hashlib
import mimetypes
from pathlib import Path
from typing import Optional, Tuple
import logging

logger = logging.getLogger(__name__)

def validate_pdf_file(file_path: str) -> Tuple[bool, str]:
    """
    Validate if a file is a valid PDF
    
    Args:
        file_path (str): Path to the file to validate
        
    Returns:
        Tuple[bool, str]: (is_valid, error_message)
    """
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            return False, "File does not exist"
        
        # Check file size
        file_size = os.path.getsize(file_path)
        if file_size == 0:
            return False, "File is empty"
        
        # Check file type
        mime_type, _ = mimetypes.guess_type(file_path)
        if mime_type != 'application/pdf':
            return False, "File is not a PDF"
        
        # Check file extension
        if not file_path.lower().endswith('.pdf'):
            return False, "File does not have .pdf extension"
        
        # Try to open with pdfplumber to validate PDF structure
        try:
            import pdfplumber
            with pdfplumber.open(file_path) as pdf:
                if len(pdf.pages) == 0:
                    return False, "PDF has no pages"
        except Exception as e:
            return False, f"Invalid PDF structure: {str(e)}"
        
        return True, "Valid PDF file"
        
    except Exception as e:
        return False, f"Validation error: {str(e)}"

def generate_file_hash(file_path: str) -> str:
    """
    Generate SHA-256 hash of a file
    
    Args:
        file_path (str): Path to the file
        
    Returns:
        str: SHA-256 hash of the file
    """
    hash_sha256 = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_sha256.update(chunk)
    return hash_sha256.hexdigest()

def sanitize_filename(filename: str) -> str:
    """
    Sanitize filename for safe file operations
    
    Args:
        filename (str): Original filename
        
    Returns:
        str: Sanitized filename
    """
    # Remove or replace invalid characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    
    # Remove leading/trailing spaces and dots
    filename = filename.strip('. ')
    
    # Limit length
    if len(filename) > 200:
        name, ext = os.path.splitext(filename)
        filename = name[:200-len(ext)] + ext
    
    return filename

def get_file_info(file_path: str) -> dict:
    """
    Get comprehensive file information
    
    Args:
        file_path (str): Path to the file
        
    Returns:
        dict: File information
    """
    try:
        stat = os.stat(file_path)
        return {
            'name': os.path.basename(file_path),
            'size': stat.st_size,
            'size_mb': round(stat.st_size / (1024 * 1024), 2),
            'created': stat.st_ctime,
            'modified': stat.st_mtime,
            'hash': generate_file_hash(file_path)
        }
    except Exception as e:
        logger.error(f"Error getting file info: {e}")
        return {}

def ensure_directory(directory_path: str) -> bool:
    """
    Ensure a directory exists, create if it doesn't
    
    Args:
        directory_path (str): Path to the directory
        
    Returns:
        bool: True if directory exists or was created successfully
    """
    try:
        os.makedirs(directory_path, exist_ok=True)
        return True
    except Exception as e:
        logger.error(f"Error creating directory {directory_path}: {e}")
        return False

def cleanup_temp_files(temp_files: list) -> None:
    """
    Clean up temporary files
    
    Args:
        temp_files (list): List of temporary file paths to delete
    """
    for file_path in temp_files:
        try:
            if os.path.exists(file_path):
                os.unlink(file_path)
                logger.info(f"Cleaned up temporary file: {file_path}")
        except Exception as e:
            logger.warning(f"Failed to clean up {file_path}: {e}")

def format_file_size(size_bytes: int) -> str:
    """
    Format file size in human-readable format
    
    Args:
        size_bytes (int): Size in bytes
        
    Returns:
        str: Formatted size string
    """
    if size_bytes == 0:
        return "0 B"
    
    size_names = ["B", "KB", "MB", "GB", "TB"]
    import math
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return f"{s} {size_names[i]}"

def is_supported_pdf(file_path: str) -> bool:
    """
    Check if PDF is supported (not corrupted, readable)
    
    Args:
        file_path (str): Path to PDF file
        
    Returns:
        bool: True if PDF is supported
    """
    try:
        import pdfplumber
        with pdfplumber.open(file_path) as pdf:
            # Check if we can access at least one page
            if len(pdf.pages) > 0:
                # Try to extract some text from first page
                first_page = pdf.pages[0]
                text = first_page.extract_text()
                return True
        return False
    except Exception:
        return False 