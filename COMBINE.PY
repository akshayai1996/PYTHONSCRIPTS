"""
===============================================================================
 Project Name:   PDF Automation Tools
 File Name:      combine.py
 Description:    Combines all PDF files in the current folder into a single PDF,
                 filename is fixed as folder name. Duplicate file/page checking
                 is optional. Skipped PDFs are logged in 'skippedpdf.txt' in
                 the same folder.

 Author:         Akshay Solanki
 Version:        1.7.0
 Dependencies:   fitz (PyMuPDF), hashlib, os
===============================================================================
"""

import fitz
import os
import hashlib

# ==============================
# USER CONFIGURATION
# ==============================
MERGE_ALPHABETICALLY = True      # True: Alphabetical order, False: folder order
SKIP_DUPLICATE_FILE_NAMES = True # True: skip files with duplicate names
SKIP_DUPLICATE_PAGES = True      # True: skip duplicate pages based on content hash

def page_hash(page):
    """Compute SHA256 hash of page text"""
    return hashlib.sha256(page.get_text("rawdict").__repr__().encode('utf-8')).hexdigest()

def combine_pdfs():
    current_dir = os.getcwd()
    folder_name = os.path.basename(current_dir)
    output_filename = f"{folder_name}.pdf"  # Fixed output file name
    skipped_log = os.path.join(current_dir, "skippedpdf.txt")  # Skipped PDFs log

    # Clear previous skipped log
    open(skipped_log, 'w').close()

    combined_pdf = fitz.open()

    # List all PDFs in folder, excluding the output file itself
    pdf_files = [f for f in os.listdir(current_dir)
                 if f.lower().endswith(".pdf") and not f.startswith(".") and f != output_filename]

    if not pdf_files:
        print("⚠️ No PDF files found in this directory.")
        return

    if MERGE_ALPHABETICALLY:
        pdf_files.sort()

    print(f"🔍 Found {len(pdf_files)} PDF files. Combining now...\n")
    print(f"📄 Merge order: {'Alphabetical' if MERGE_ALPHABETICALLY else 'Folder order'}\n")

    seen_file_names = set()
    seen_page_hashes = set()

    for filename in pdf_files:
        file_skipped = False

        # Check for duplicate file name
        if SKIP_DUPLICATE_FILE_NAMES and filename in seen_file_names:
            print(f"⚠️ Skipping duplicate file name: {filename}")
            file_skipped = True
        else:
            seen_file_names.add(filename)

        file_path = os.path.join(current_dir, filename)

        try:
            with fitz.open(file_path) as pdf:
                for page in pdf:
                    if SKIP_DUPLICATE_PAGES:
                        ph = page_hash(page)
                        if ph in seen_page_hashes:
                            file_skipped = True
                            continue
                        seen_page_hashes.add(ph)
                    combined_pdf.insert_pdf(fitz.open(file_path), from_page=page.number, to_page=page.number)
            if not file_skipped:
                print(f"✅ Added: {filename}")
        except Exception as e:
            print(f"❌ Error processing {filename}: {e}")
            file_skipped = True

        # Log skipped PDF if any reason to skip
        if file_skipped:
            with open(skipped_log, 'a') as f:
                f.write(filename + "\n")

    try:
        combined_pdf.save(output_filename)
        print(f"\n🎉 All PDFs combined successfully into: {output_filename}")
        print(f"ℹ️ Skipped PDFs (if any) are logged in '{skipped_log}'")
    except Exception as e:
        print(f"🚨 Failed to save combined file: {e}")
    finally:
        combined_pdf.close()

if __name__ == "__main__":
    try:
        combine_pdfs()
    except Exception as e:
        print(f"\n🚨 Fatal error occurred: {e}")

