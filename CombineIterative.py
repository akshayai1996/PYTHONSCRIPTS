"""
===============================================================================
 Project Name:   Incremental PDF Automation Tools
 File Name:      CombineIterative.py
 Description:    Iterates through all subfolders of a root folder, combines all
                 PDFs in each subfolder into a single PDF named after the folder,
                 supports incremental merging, optional duplicate file/page checks,
                 logs skipped PDFs, and displays a progress bar with tqdm.

 Author:         Akshay Solanki
 Version:        1.2.0
 Dependencies:   fitz (PyMuPDF), hashlib, os, tqdm
===============================================================================
"""

import fitz
import os
import hashlib
import sys
from tqdm import tqdm

# ==============================
# USER CONFIGURATION
# ==============================
ROOT_FOLDER = r"Z:\01. P&IDs. IFC-1 + IFC 2 rev 4"  # Root folder to iterate
MERGE_ALPHABETICALLY = True      # True: Alphabetical order, False: folder order
SKIP_DUPLICATE_FILE_NAMES = True # True: skip files with duplicate names
SKIP_DUPLICATE_PAGES = True      # True: skip duplicate pages based on content hash

# ==============================
# HELPER FUNCTIONS
# ==============================
def page_hash(page):
    """Compute SHA256 hash of page text using XML content for stability"""
    return hashlib.sha256(page.get_text("xml").encode('utf-8')).hexdigest()

def log_skipped(log_path, filename, reason):
    """Append skipped file info to skippedpdf.txt"""
    with open(log_path, 'a', encoding='utf-8') as f:
        f.write(f"{filename} | Reason: {reason}\n")

def load_manifest(manifest_path):
    """Load previously combined PDF filenames"""
    if os.path.exists(manifest_path):
        with open(manifest_path, 'r', encoding='utf-8') as f:
            return set(line.strip() for line in f.readlines())
    return set()

def update_manifest(manifest_path, new_files):
    """Append newly combined PDF filenames to manifest"""
    with open(manifest_path, 'a', encoding='utf-8') as f:
        for file in new_files:
            f.write(file + '\n')

# ==============================
# MAIN FUNCTION
# ==============================
def combine_pdfs_in_folder(folder_path):
    folder_name = os.path.basename(folder_path)
    output_pdf_path = os.path.join(folder_path, f"{folder_name}.pdf")
    skipped_log_path = os.path.join(folder_path, "skippedpdf.txt")
    manifest_path = os.path.join(folder_path, "combined_manifest.txt")

    # Clear skipped log at the start
    open(skipped_log_path, 'w', encoding='utf-8').close()

    combined_pdf = fitz.open()
    pdf_files = [f for f in os.listdir(folder_path)
                 if f.lower().endswith(".pdf") and f != f"{folder_name}.pdf" and not f.startswith(".")]

    if MERGE_ALPHABETICALLY:
        pdf_files.sort()

    previously_combined = load_manifest(manifest_path)
    new_files_to_add = []

    seen_file_names = set()
    seen_page_hashes = set()

    # Hash pages of existing combined PDF if duplicate page check is active
    if os.path.exists(output_pdf_path) and SKIP_DUPLICATE_PAGES:
        try:
            with fitz.open(output_pdf_path) as existing_pdf:
                for page in existing_pdf:
                    seen_page_hashes.add(page_hash(page))
        except Exception as e:
            print(f"⚠️ Could not read existing combined PDF '{output_pdf_path}': {e}")

    print(f"\n📂 Processing folder: {folder_path} ({len(pdf_files)} PDF(s))")

    # Use tqdm progress bar for PDFs in this subfolder
    for filename in tqdm(pdf_files, desc=f"Processing PDFs in {folder_name}", unit="file"):
        if filename in previously_combined:
            continue

        file_path = os.path.join(folder_path, filename)

        # Duplicate file name check
        if SKIP_DUPLICATE_FILE_NAMES and filename in seen_file_names:
            log_skipped(skipped_log_path, filename, "Duplicate file name")
            continue
        seen_file_names.add(filename)

        temp_pdf = fitz.open()
        pages_added = 0

        try:
            with fitz.open(file_path) as pdf:
                for page in pdf:
                    if SKIP_DUPLICATE_PAGES:
                        ph = page_hash(page)
                        if ph in seen_page_hashes:
                            continue
                        seen_page_hashes.add(ph)
                    temp_pdf.insert_pdf(pdf, from_page=page.number, to_page=page.number)
                    pages_added += 1

            if pages_added > 0:
                combined_pdf.insert_pdf(temp_pdf)
                new_files_to_add.append(filename)
            else:
                log_skipped(skipped_log_path, filename, "All pages duplicates")

        except Exception as e:
            log_skipped(skipped_log_path, filename, f"Error processing: {e}")
        finally:
            temp_pdf.close()

    # Save combined PDF if any new pages were added
    if new_files_to_add:
        try:
            combined_pdf.save(output_pdf_path, garbage=4, deflate=True)
            update_manifest(manifest_path, new_files_to_add)
            print(f"🎉 Combined PDF saved: {output_pdf_path}")
            print(f"ℹ️ Skipped PDFs logged in: {skipped_log_path}")
        except Exception as e:
            print(f"🚨 Failed to save combined PDF '{output_pdf_path}': {e}")
    else:
        print(f"ℹ️ No new pages added for folder '{folder_name}'.")

    combined_pdf.close()

# ==============================
# ITERATE THROUGH ALL SUBFOLDERS
# ==============================
def process_all_subfolders(root_folder):
    all_dirs = [os.path.join(root, d) for root, dirs, _ in os.walk(root_folder) for d in dirs]
    # tqdm progress bar for subfolders
    for subfolder_path in tqdm(all_dirs, desc="Processing all subfolders", unit="folder"):
        combine_pdfs_in_folder(subfolder_path)

# ==============================
# SCRIPT ENTRY POINT
# ==============================
if __name__ == "__main__":
    if not os.path.exists(ROOT_FOLDER):
        print(f"🚨 Root folder does not exist: {ROOT_FOLDER}")
        sys.exit(1)

    try:
        process_all_subfolders(ROOT_FOLDER)
        print("\n✅ All subfolders processed successfully.")
    except Exception as e:
        print(f"\n🚨 Fatal error: {e}")
        sys.exit(1)
