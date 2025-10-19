"""
===============================================================================
 Project Name:    PDF Page Extractor from Excel (Fixed PDF Path Version)
 File Name:       ExtractPagesFromExcel.py
 Description:     Iterates through subfolders, reads user-specified Excel files,
                  extracts pages from a fixed PDF path, and saves them in the
                  same folder as each Excel file. Logs all errors to a single
                  'log_errors.txt' in the root folder.

 Features:
 - Fixed PDF path (any location)
 - Incremental extraction (skips already extracted pages)
 - Robust page number parsing (handles numbers stored as text)
 - Progress bar per folder
 - Fully user-customizable: root folder, Excel filename, PDF path, column name
 Author:          Akshay Solanki
 Dependencies:    pandas, PyPDF2, openpyxl, tqdm, os
===============================================================================
"""

import pandas as pd
import PyPDF2
import os
from tqdm import tqdm

# ==============================
# USER CONFIGURATION
# ==============================
root_folder = r"Z:\01. P&IDs. IFC-1 + IFC 2 rev 4"       # Root folder to iterate
excel_filename = "output.xlsx"                           # Excel file name to look for
pdf_path = r"Z:\01. P&IDs. IFC-1 + IFC 2 rev 4\PID_4.pdf"  # Full path to the PDF
excel_column_name = "PageNumbers"                        # Column header in Excel containing page numbers
log_file_name = "log_errors.txt"                         # Name of consolidated log file

# ==============================
# INITIALIZATION
# ==============================
log_file = os.path.join(root_folder, log_file_name)
open(log_file, 'w').close()  # Clear previous log file

def log_error(msg):
    """Log errors to console and main log file"""
    print(f"❌ {msg}")
    with open(log_file, 'a') as f:
        f.write(msg + "\n")

def extract_pages(pdf_path, page_numbers, folder_path):
    """Extract pages from PDF and save in the specified folder"""
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_path)
    except FileNotFoundError:
        log_error(f"PDF file not found: {pdf_path}")
        return
    except Exception as e:
        log_error(f"Failed to load PDF '{pdf_path}': {e}")
        return

    total_pages = len(pdf_reader.pages)

    # Already extracted pages in the folder
    already_extracted = set(
        int(f.split('.')[0]) for f in os.listdir(folder_path)
        if f.endswith('.pdf') and f.split('.')[0].isdigit()
    )
    pages_to_extract = [p for p in page_numbers if p not in already_extracted]

    if not pages_to_extract:
        print(f"All pages already extracted for '{folder_path}'. Skipping.")
        return

    for page_num in pages_to_extract:
        try:
            if page_num < 1 or page_num > total_pages:
                log_error(f"Page number {page_num} out of range in PDF '{pdf_path}'")
                continue

            pdf_writer = PyPDF2.PdfWriter()
            page = pdf_reader.pages[page_num - 1]
            pdf_writer.add_page(page)

            # Optional: attempt A3 landscape
            page.mediabox.upper_right = (1191, 842)

            output_pdf_path = os.path.join(folder_path, f"{page_num}.pdf")
            with open(output_pdf_path, 'wb') as f:
                pdf_writer.write(f)

        except Exception as e:
            log_error(f"Failed to extract page {page_num} from '{pdf_path}': {e}")

# ==============================
# ITERATE THROUGH SUBFOLDERS
# ==============================
for subdir, dirs, files in os.walk(root_folder):
    if excel_filename not in files:
        continue  # skip folders without Excel

    excel_path = os.path.join(subdir, excel_filename)
    folder_path = subdir  # save extracted pages in the same folder as Excel

    # Load Excel
    try:
        df = pd.read_excel(excel_path, engine='openpyxl')
    except Exception as e:
        log_error(f"Failed to read Excel '{excel_path}': {e}")
        continue

    # ==============================
    # Robust page number parsing
    # ==============================
    page_numbers = set()
    if excel_column_name not in df.columns:
        log_error(f"Column '{excel_column_name}' not found in Excel '{excel_path}'. Skipping folder.")
        continue

    for pages in df[excel_column_name]:
        if pd.isna(pages):
            continue
        pages_str = str(pages)
        for p in pages_str.split(','):
            p = p.strip()
            if not p:
                continue
            try:
                page_numbers.add(int(p))
            except ValueError:
                log_error(f"Invalid page number '{p}' in Excel '{excel_path}'. Skipping.")

    if not page_numbers:
        log_error(f"No valid pages found in Excel '{excel_path}'. Skipping folder.")
        continue

    # Extract pages with progress bar
    for _ in tqdm([pdf_path], desc=f"Processing PDF for folder '{subdir}'"):
        extract_pages(pdf_path, page_numbers, folder_path)

print("\n🎉 Iterative incremental extraction complete. All errors logged in 'log_errors.txt'.")
