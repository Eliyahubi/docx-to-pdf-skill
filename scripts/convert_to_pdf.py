#!/usr/bin/env python3
"""
DOCX to PDF Converter
Converts Word documents to PDF using LibreOffice
"""

import sys
import subprocess
import os
from pathlib import Path

OUTPUT_DIR = "/mnt/user-data/outputs"

def convert_docx_to_pdf(docx_path):
    """
    Convert a single DOCX file to PDF using LibreOffice
    
    Args:
        docx_path: Path to the DOCX file
        
    Returns:
        tuple: (success: bool, pdf_path: str or None, error: str or None)
    """
    try:
        docx_file = Path(docx_path)
        
        # Validate input
        if not docx_file.exists():
            return False, None, f"File not found: {docx_path}"
        
        if not docx_file.suffix.lower() == '.docx':
            return False, None, f"Not a DOCX file: {docx_path}"
        
        # Ensure output directory exists
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        
        # Convert using LibreOffice
        # --headless: run without GUI
        # --convert-to pdf: output format
        # --outdir: where to save the PDF
        cmd = [
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', OUTPUT_DIR,
            str(docx_file)
        ]
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60  # 60 second timeout per file
        )
        
        if result.returncode != 0:
            return False, None, f"LibreOffice error: {result.stderr}"
        
        # Construct expected PDF path
        pdf_name = docx_file.stem + '.pdf'
        pdf_path = Path(OUTPUT_DIR) / pdf_name
        
        if not pdf_path.exists():
            return False, None, "PDF was not created"
        
        return True, str(pdf_path), None
        
    except subprocess.TimeoutExpired:
        return False, None, "Conversion timeout (file too large or complex)"
    except Exception as e:
        return False, None, f"Unexpected error: {str(e)}"

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 convert_to_pdf.py <file1.docx> [file2.docx ...]")
        print("Example: python3 convert_to_pdf.py /mnt/user-data/uploads/*.docx")
        sys.exit(1)
    
    input_files = sys.argv[1:]
    
    print(f"Starting conversion of {len(input_files)} file(s)...\n")
    
    successful = []
    failed = []
    
    for docx_path in input_files:
        filename = os.path.basename(docx_path)
        print(f"Converting: {filename}...", end=" ")
        
        success, pdf_path, error = convert_docx_to_pdf(docx_path)
        
        if success:
            print("✓ Success")
            successful.append((filename, pdf_path))
        else:
            print(f"✗ Failed")
            failed.append((filename, error))
    
    # Print summary
    print("\n" + "="*60)
    print("CONVERSION SUMMARY")
    print("="*60)
    print(f"Total files: {len(input_files)}")
    print(f"Successful: {len(successful)}")
    print(f"Failed: {len(failed)}")
    
    if successful:
        print("\n✓ Successfully converted:")
        for filename, pdf_path in successful:
            pdf_name = os.path.basename(pdf_path)
            print(f"  • {filename} → {pdf_name}")
    
    if failed:
        print("\n✗ Failed conversions:")
        for filename, error in failed:
            print(f"  • {filename}")
            print(f"    Reason: {error}")
    
    print("\nAll PDF files are saved in: /mnt/user-data/outputs/")
    
    # Exit with error code if any conversions failed
    sys.exit(0 if not failed else 1)

if __name__ == "__main__":
    main()
