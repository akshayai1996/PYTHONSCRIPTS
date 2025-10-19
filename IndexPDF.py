"""
===============================================================================
 Project Name:   PDF Indexing Tool with Progress
 File Name:      IndexPDF.py
 Description:    Extracts all text from a single PDF and saves it as a TXT
                 file with page markers for easy searching later. Displays
                 a progress bar using tqdm for large PDFs.

 Author:         Akshay Solanki
 Created on:     19-Oct-2025
 Version:        1.1.0
 License:        Unlicense (Public Domain - Free to use without attribution)

 Dependencies:   fitz (PyMuPDF), os, tqdm

 Usage:
     1. Update the USER CONFIGURATION section with your PDF path.
     2. Install tqdm if not already: pip install tqdm
     3. Run:
        python IndexPDF.py

 Notes:
     - Generates a TXT file containing the full text of the PDF.
     - Each page is marked with "======= PAGE <n> =======" for easy reference.
     - Shows a progress bar for page indexing.
===============================================================================
"""

import fitz  # PyMuPDF
import os
from tqdm import tqdm

# ==============================
# USER CONFIGURATION
# ==============================
PDF_PATH = r"C:\Users\Akshay\Documents\MyBigDocument.pdf"  # Path to PDF to index
# ==============================

def create_named_pdf_index(pdf_path):
    """
    Extract all text from the PDF and save it to a TXT file with page markers.
    Returns the path to the TXT index file.
    """
    if not os.path.exists(pdf_path):
        print(f"🚨 Error: PDF file not found at {pdf_path}")
        return None

    base_name = os.path.splitext(pdf_path)[0]
    index_path = base_name + ".txt"
    print(f"🔍 Starting indexing of: {os.path.basename(pdf_path)}")

    try:
        with fitz.open(pdf_path) as pdf, open(index_path, 'w', encoding='utf-8') as outfile:
            outfile.write(f"--- INDEX FOR: {os.path.basename(pdf_path)} ---\n")
            outfile.write(f"--- TOTAL PAGES: {len(pdf)} ---\n\n")

            # tqdm progress bar for pages
            for page_num in tqdm(range(len(pdf)), desc="Indexing Pages", unit="page"):
                page = pdf.load_page(page_num)
                text_content = page.get_text("text")
                outfile.write(f"======= PAGE {page_num + 1} =======\n")
                outfile.write(text_content)
                outfile.write("\n\n")

    except Exception as e:
        print(f"❌ Failed to create index: {e}")
        return None

    print(f"\n🎉 Indexing complete! Output index file saved to: {index_path}")
    return index_path

if __name__ == "__main__":
    create_named_pdf_index(PDF_PATH)
