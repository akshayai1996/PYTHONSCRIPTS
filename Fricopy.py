"""
===============================================================================
 Project Name:    FRI PDF Copier
 File Name:       Fricopy.py
 Description:     Copies LINEWISE PDFs to matching backup folders based on 
                  names inside parentheses in backup PDFs. Supports:
                  - Subfolder iteration
                  - Overwrite option
                  - Progress bar with tqdm
                  - Comprehensive logging

 Author:          Akshay Solanki
 Created on:      19-Oct-2025
 Version:         2.0.0
 License:         Unlicense (Public Domain)

 Usage:
     python Fricopy.py

 User Configuration:
     backup_root         -> Root folder containing backup PDFs (and subfolders)
     linewise_root       -> Source folder containing LINEWISE PDFs
     OVERWRITE_EXISTING  -> True: overwrite existing copied files
                            False: skip existing files
     LOG_FILE            -> File to log all skipped or not copied files

 Notes for Non-Technical Users:
     - The script automatically scans all subfolders of backup_root.
     - It searches for backup PDF names containing parentheses.
       Example: "P&ID (Valve1).pdf" -> Extracts "Valve1"
     - It matches "Valve1" against LINEWISE PDFs (case-insensitive).
     - Matching files are copied to the backup folder with "_FRI.pdf" appended.
===============================================================================
"""

import os
import shutil
import re
from tqdm import tqdm
from collections import defaultdict

# ==============================
# USER CONFIGURATION
# ==============================
backup_root = r"Z:\ALL_LOOPS_SOFT_COPY_BACKUP"      # Root folder to iterate
linewise_root = r"Z:\PIPING FRI\LINEWISE"          # Source folder
OVERWRITE_EXISTING = False                           # True = overwrite existing PDFs
LOG_FILE = "skipped_or_not_overwritten_log.txt"     # Log file for skipped PDFs
# ==============================

def extract_name_in_parentheses(filename):
    """
    Extract text inside parentheses from a filename.
    Example: "P&ID (Valve1).pdf" -> "Valve1"
    """
    match = re.search(r"\(([^)]+)\)", filename)
    return match.group(1).strip() if match else None

def create_linewise_index(root_path):
    """
    Scan linewise_root folder (and subfolders) and create a fast lookup dictionary.
    Key: Base filename in lowercase
    Value: List of tuples: (full path, actual filename)
    """
    linewise_index = defaultdict(list)
    print("🧠 Creating LINEWISE file index (including subfolders)...")
    
    for subdir, _, files in os.walk(root_path):
        for linewise_file in files:
            if linewise_file.lower().endswith(".pdf"):
                search_key = os.path.splitext(linewise_file)[0].lower()
                full_path = os.path.join(subdir, linewise_file)
                linewise_index[search_key].append((full_path, linewise_file))
    
    print(f"✅ Indexed {sum(len(v) for v in linewise_index.values())} LINEWISE PDFs.\n")
    return linewise_index

def fricopy():
    """
    Main function to copy LINEWISE PDFs to backup folders based on
    names in parentheses with optional overwrite and logging.
    """
    # Clear previous log file
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"--- Log for FRI PDF Copier (Version 2.0.0) ---\n")
        f.write(f"Overwrite Existing Files: {OVERWRITE_EXISTING}\n\n")

    def write_log(message):
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(message + "\n")
    
    # Step 1: Index LINEWISE folder (fast lookup)
    linewise_index = create_linewise_index(linewise_root)
    
    # Step 2: Gather all backup PDFs including subfolders
    all_backup_files = []
    for subdir, _, files in os.walk(backup_root):
        for file in files:
            if file.lower().endswith(".pdf"):
                all_backup_files.append((subdir, file))
    
    print(f"🔍 Found {len(all_backup_files)} backup PDFs to process.\n")

    copied_count = 0
    skipped_count = 0

    # Step 3: Iterate over backup PDFs with progress bar
    for subdir, file in tqdm(all_backup_files, desc="Processing Backup PDFs", unit="file"):
        base_name = extract_name_in_parentheses(file)
        
        if not base_name:
            skipped_count += 1
            write_log(f"SKIPPED (No Parentheses): {os.path.join(subdir, file)}")
            continue

        base_name_lower = base_name.lower()
        found_match = False

        # Step 4: Search LINEWISE index for matching PDFs
        for linewise_key, linewise_entries in linewise_index.items():
            if base_name_lower in linewise_key:
                for full_path, lw_file in linewise_entries:
                    # Target filename: add _FRI.pdf
                    target_filename = os.path.splitext(lw_file)[0] + "_FRI.pdf"
                    target_path = os.path.join(subdir, target_filename)

                    if not os.path.exists(target_path) or OVERWRITE_EXISTING:
                        shutil.copy2(full_path, target_path)
                        copied_count += 1
                        tqdm.write(f"✅ Copied: {lw_file} → {os.path.relpath(target_path, backup_root)}")
                    else:
                        skipped_count += 1
                        tqdm.write(f"⚠️ Skipped (already exists): {os.path.relpath(target_path, backup_root)}")
                        write_log(f"SKIPPED (Exists/No Overwrite): {os.path.join(subdir, file)} | Target: {target_filename}")
                    
                    found_match = True
        
        if not found_match:
            skipped_count += 1
            write_log(f"SKIPPED (No Match Found in LINEWISE): {os.path.join(subdir, file)} (Extracted Key: {base_name})")

    # Summary
    print("\n📊 Summary:")
    print(f"Total files copied: {copied_count}")
    print(f"Total files skipped/not overwritten: {skipped_count}")
    print(f"All skip details logged in '{LOG_FILE}' (saved in script folder).")

# ==============================
# Entry Point
# ==============================
if __name__ == "__main__":
    try:
        fricopy()
    except Exception as e:
        print(f"\n🚨 A fatal error occurred: {e}")
