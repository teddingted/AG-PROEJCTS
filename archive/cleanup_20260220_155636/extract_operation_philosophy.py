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

PDF_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\RL-KM60002-A-03_Operation and Control Philosophy_Finished Plan.pdf"
OUTPUT_PATH = r"c:\Users\Admin\Desktop\AG-BEGINNING\operation_philosophy_content.txt"

def extract_operation_philosophy():
    """Extract text from Operation Philosophy PDF"""
    print("Extracting Operation and Control Philosophy document...")
    
    try:
        with open(PDF_PATH, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            num_pages = len(pdf_reader.pages)
            
            print(f"Document has {num_pages} pages")
            
            full_text = ""
            for page_num in range(num_pages):
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                full_text += text + "\n"
            
            # Save to text file
            with open(OUTPUT_PATH, 'w', encoding='utf-8') as out_file:
                out_file.write(full_text)
            
            char_count = len(full_text)
            print(f"Extracted {char_count} characters to {OUTPUT_PATH}")
            
            return full_text
            
    except Exception as e:
        print(f"Error: {e}")
        return None

if __name__ == "__main__":
    content = extract_operation_philosophy()
    
    if content:
        # Preview first 500 characters
        print("\n=== DOCUMENT PREVIEW ===")
        print(content[:500])
        print("...")
