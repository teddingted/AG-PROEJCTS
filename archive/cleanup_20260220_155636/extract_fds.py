
# -*- coding: utf-8 -*-
import sys
import os

# UTF-8 Fix for Windows cp949
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr.reconfigure(encoding='utf-8')
    os.environ['PYTHONIOENCODING'] = 'utf-8'


import PyPDF2
import os

pdf_path = r"c:\Users\Admin\Desktop\AG-BEGINNING\8196 FDS_Final.pdf"
output_path = r"c:\Users\Admin\Desktop\AG-BEGINNING\fds_content.txt"

def extract_text(pdf_path, output_path):
    print(f"Extracting text from {pdf_path}...")
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                text += page.extract_text() + "\n"
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        
        print(f"Extracted {len(text)} characters to {output_path}")
        return True
    except Exception as e:
        print(f"Error extracting PDF: {e}")
        return False

if __name__ == "__main__":
    extract_text(pdf_path, output_path)
