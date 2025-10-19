"""
==============================================================================
 Project Name:   PDF ISO to Excel Mapper (Incremental Version)
 File Name:      Generate_excel.py
 Description:    Reads an index file (PID_4_index.txt) to map page numbers
                 for 'pid_4.pdf', traverses subfolders in a root directory,
                 extracts ISO codes from PDF filenames, and generates Excel
                 files listing ISO codes and their pages.

                 This incremental version:
                 - Checks existing Excel files per folder
                 - Compares ISO list in Excel vs current PDFs
                 - Only processes missing ISOs
                 - Preserves previously processed data

 Author:         Akshay Solanki
 Created on:     19-Oct-2025
 Version:        1.3.0
 License:        Unlicense (Public Domain)

 Dependencies:   pandas, tqdm, openpyxl
 Python Version: 3.8+

 Important Notes:
 - PDFs must be indexed in PID_4_index.txt for fast page number extraction.
   If not indexed, processing may take several minutes per folder.
 - Excel column "PDF PAGE" is forced to numeric format.
 - PDF filenames must follow the pattern: "anything(ISO1-ISO2-...).pdf"

 ISO Extraction Example:
 - Filename: "Drawing(AB-12-34).pdf"
 - Extracted ISO: "AB-12"  (first two segments inside parentheses)

 Example Excel Table:

 | ISO LIST | PDF PAGE |
 |----------|----------|
 | AB-12    | 1,2,3    |
 | XY-99    | 4,5      |

 Usage:
     python Generate_excel.py

 Customization:
     - root_dir: Folder containing PDF subfolders
     - index_file_path: PID index text file path
     - output_excel_name: Excel file name to generate in each subfolder
==============================================================================
"""

# ==============================
# Import Required Libraries
# ==============================
import os
import re
import pandas as pd
from tqdm import tqdm

# ==============================
# User Configuration
# ==============================
root_dir = r"Z:\ALL_LOOPS_SOFT_COPY_BACKUP"
index_file_path = r"Z:\01. P&IDs. IFC-1 + IFC 2 rev 4\PID_4_index.txt"
output_excel_name = "output.xlsx"

# ==============================
# Step 1: Load PID index file
# ==============================
iso_page_map = {}

try:
    with open(index_file_path, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            parts = line.strip().split()
            if len(parts) >= 2 and parts[0].lower() == "pid_4.pdf":
                try:
                    page_number = int(parts[1])
                except ValueError:
                    continue
                iso_page_map.setdefault("pid_4.pdf", []).append(page_number)
except FileNotFoundError:
    print(f"❌ Index file not found at: {index_file_path}")
    exit(1)

if not iso_page_map.get("pid_4.pdf"):
    print("⚠️ No page numbers found in index file. Ensure PDFs are indexed for fast processing.")

# ==============================
# Step 2: Regex to extract ISO from PDF filename
# ==============================
pdf_pattern = re.compile(r'\(([^)]+)\)\.pdf$', re.IGNORECASE)

# ==============================
# Step 3: Collect all subdirectories
# ==============================
subdirs = [os.path.join(dp, f) for dp, dn, _ in os.walk(root_dir) for f in dn]

if not subdirs:
    print(f"⚠️ No subfolders found under '{root_dir}'. Exiting.")
    exit(1)

print(f"🔍 Found {len(subdirs)} subfolders. Starting processing...\n")

# ==============================
# Step 4: Traverse subfolders and generate Excel
# ==============================
for idx, subdir in enumerate(tqdm(subdirs, desc="Processing folders")):
    print(f"\n📁 [{idx+1}/{len(subdirs)}] Processing folder: {subdir}")
    iso_set = set()

    # List all PDFs in folder
    pdf_files = [f for f in os.listdir(subdir) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print(f"⚠️ No PDF files found in this folder. Skipping...")
        continue

    # Extract ISO codes from filenames
    for file in pdf_files:
        match = pdf_pattern.search(file)
        if match:
            parenthesis_content = match.group(1)
            segments = parenthesis_content.split('-')
            if len(segments) >= 2:
                iso = f"{segments[0]}-{segments[1]}"
                iso_set.add(iso)

    if not iso_set:
        print("⚠️ No ISO codes extracted from PDFs in this folder. Skipping...")
        continue

    iso_list = sorted(iso_set)
    excel_path = os.path.join(subdir, output_excel_name)

    # ==============================
    # Step 4a: Check existing Excel (incremental)
    # ==============================
    existing_iso_list = set()
    if os.path.exists(excel_path):
        try:
            df_existing = pd.read_excel(excel_path, engine='openpyxl')
            existing_iso_list = set(df_existing["ISO LIST"].dropna().astype(str))
        except Exception as e:
            print(f"⚠️ Could not read existing Excel file: {e}")

    # Determine new ISOs to process
    new_isos = [iso for iso in iso_list if iso not in existing_iso_list]

    if not new_isos:
        print(f"⏭️ All ISOs already processed for this folder. Skipping...")
        continue

    print(f"📝 New ISOs to process: {len(new_isos)}")

    # ==============================
    # Step 4b: Map PDF pages
    # ==============================
    pdf_pages = []
    for iso in new_isos:
        page_numbers = iso_page_map.get("pid_4.pdf", [])
        page_str = ",".join(str(p) for p in sorted(set(page_numbers))) if page_numbers else ""
        pdf_pages.append(page_str)

    # Create DataFrame for new ISOs
    df_new = pd.DataFrame({
        "ISO LIST": new_isos,
        "PDF PAGE": pdf_pages
    })

    # Force numeric format for PDF PAGE
    def format_numeric(cell):
        if not cell:
            return None
        nums = [int(n) for n in cell.split(",") if n.isdigit()]
        return ",".join(str(n) for n in nums)

    df_new["PDF PAGE"] = df_new["PDF PAGE"].apply(format_numeric)

    # Merge with existing Excel if present
    if existing_iso_list:
        df_final = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_final = df_new

    # Save Excel
    try:
        df_final.to_excel(excel_path, index=False, engine='openpyxl')
        print(f"✅ Excel updated: {excel_path} ({len(df_final)} total rows)")
    except Exception as e:
        print(f"❌ Failed to write Excel in {subdir}: {e}")

print("\n🎉 All processing completed successfully.")
print("⚠️ Reminder: PDFs must be indexed for fast page number extraction.")
