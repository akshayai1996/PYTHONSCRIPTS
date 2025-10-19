
🤖 Complete Master Orchestrator: Full 7-Process Workflow (v3.1)
📌 Code Explanation for Non-Technical Users
This script is a project manager in code that automates the organization and assembly of project documents.
 * It asks you (via pop-up windows) for five main locations: the central Excel sheet (the task list), the folder where all source documents (ISOs) are stored, a Master Index File (which pages to extract), the Master PDF (the source book for pages), and the Destination Folder (where the final organized project structure will be created).
 * It reads the Excel and creates a dedicated, clearly named folder for every item in your list.
 * It automatically fetches the original ISO document from your server and copies it into its new folder (P1).
 * It uses the Master Index to find the specific pages needed for that ISO and extracts them from the Master PDF, saving them as individual files (P2, P3).
 * It makes a secure backup of every file in the main folder by creating a copy with a _FRI suffix (P4).
 * It cleans up any unnecessary original ISO files that didn't have pages mapped (P5).
 * Finally, it merges all the remaining necessary documents and extracted pages in that folder into one comprehensive PDF file named Combined.pdf (P6).
This process ensures every project folder is complete, backed up, and consolidated for easy review.
"""
===============================================================================
 Project Name:    Complete Master Orchestrator - Full 7-Process Workflow
 File Name:       Master_Orchestrator_Final_v3.1.py
 Description:     Runs 7 interconnected processes:
                  P1: ISO Manager
                  P2: Generate Excel with ISO->Page Mapping
                  P3: Extract Pages (duplicate-safe)
                  P4: FRI Copies (Modified: uses _FRI suffix in source folder)
                  P5: Cleanup Redundancy
                  P6: Combine PDFs
                  P7: Final Cleanup + Verification

 Author:          Akshay Solanki (Original), Gemini (Updates)
 Created on:      19-Oct-2025
 Version:         3.1
 Dependencies:    pandas, openpyxl, tkinter, tqdm, PyPDF2, fitz, re
 Notes:           GUI-based path selection; incremental updates supported.
===============================================================================
"""
import os, shutil, re, time
from collections import defaultdict
from datetime import datetime
import pandas as pd
from tqdm import tqdm
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import PyPDF2

# ===============================
# Global Log and Error Paths
# ===============================
LOG_FILE = os.path.join(os.getcwd(), "orchestrator_log.txt")
ERROR_REPORT = os.path.join(os.getcwd(), "error_report.txt")

# ===============================
# Logging Functions
# ===============================
def log_msg(msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def log_error(msg):
    with open(ERROR_REPORT, "a", encoding="utf-8") as f:
        f.write(msg + "\n")

# ===============================
# GUI File/Folder Selectors
# ===============================
def select_folder(title):
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title=title)
    root.destroy()
    return folder

def select_file(title, filetypes=[("All files", "*.*")]):
    root = tk.Tk()
    root.withdraw()
    file = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return file

# ===============================
# Time Formatter
# ===============================
def format_time(seconds):
    hrs = int(seconds // 3600)
    mins = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    return f"{hrs:02d}:{mins:02d}:{secs:02d}"

# ===============================
# Safe Copy for ISO/PDF
# ===============================
def safe_copy(src, dest):
    if not os.path.exists(src):
        raise FileNotFoundError(src)
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    base, ext = os.path.splitext(dest)
    candidate = dest
    i = 1
    while os.path.exists(candidate):
        if os.path.getsize(candidate) == os.path.getsize(src):
            # File already exists and is identical, no copy needed.
            return candidate
        # File exists but is different, create a duplicate with a suffix.
        candidate = f"{base}_dup{i}{ext}"
        i += 1
    shutil.copy2(src, candidate)
    return candidate

# ===============================
# Folder & ISO Helpers
# ===============================
def make_folder_name(loop_no: str, system_no: str) -> str:
    return f"{str(loop_no).strip()}_{str(system_no).strip()}"

def find_iso_on_server(iso_no: str, server_path: str) -> str:
    if not iso_no:
        return ""
    iso_no = iso_no.strip().lower()
    try:
        for item in os.listdir(server_path):
            if item.lower().endswith(".pdf") and f"({iso_no})" in item.lower():
                full_path = os.path.join(server_path, item)
                if os.path.isfile(full_path):
                    return full_path
    except PermissionError:
        pass
    return ""

# ===============================
# Excel Utilities (Pre-flight Check)
# ===============================
def create_or_update_excel(excel_file):
    headers = ["Iso no", "loop no", "system no", "folder name", "history folder name", "ISO Status"]
    if not os.path.exists(excel_file):
        df = pd.DataFrame(columns=headers)
        df.to_excel(excel_file, index=False)
        messagebox.showinfo("Excel Created", f"Pre-flight check failed. Please fill first three columns and save: {excel_file}")
        return False
    
    # Pre-flight check passed: file exists. Now, ensure data structure is correct.
    df = pd.read_excel(excel_file, dtype=str).fillna("")
    for col in headers:
        if col not in df.columns:
            df[col] = ""
    df = df[headers]
    
    # Ensure folder names are calculated and history is initialized.
    for idx, row in df.iterrows():
        df.at[idx, "folder name"] = make_folder_name(row["loop no"], row["system no"])
        if not row["history folder name"]:
            df.at[idx, "history folder name"] = row["folder name"]
    
    # Save the updated structure back before P1 runs.
    df.to_excel(excel_file, index=False)
    return True

def highlight_missing_iso(excel_file):
    try:
        wb = load_workbook(excel_file)
        ws = wb.active
        red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        clear_fill = PatternFill(fill_type=None)
        header = [cell.value for cell in ws[1]]
        if "ISO Status" not in header:
            return
        col_idx = header.index("ISO Status") + 1
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            status_val = str(cell.value).strip().upper() if cell.value else ""
            fill = red_fill if status_val == "MISSING" else clear_fill
            # Apply fill to the entire row if ISO is missing
            for c in range(1, ws.max_column + 1):
                ws.cell(row=row, column=c).fill = fill
        wb.save(excel_file)
    except Exception as e:
        log_msg(f"ERROR highlighting: {e}")

# ===============================
# P1: ISO MANAGER
# ===============================
def iso_manager(excel_file, server_path, dest_root):
    process_start = time.time()
    log_msg("=== P1: ISO MANAGER START ===")
    
    df = pd.read_excel(excel_file, dtype=str).fillna("")
    for idx, row in df.iterrows():
        df.at[idx, "folder name"] = make_folder_name(row["loop no"], row["system no"])

    processed_folders = {}
    for idx, row in df.iterrows():
        desired = row["folder name"].strip()
        history = row["history folder name"].strip()
        if not desired or not history:
            continue
        key = (history, desired)
        if key in processed_folders:
            df.at[idx, "history folder name"] = desired
            continue
        desired_path = os.path.join(dest_root, desired)
        history_path = os.path.join(dest_root, history)
        try:
            if history != desired and os.path.exists(history_path):
                if not os.path.exists(desired_path):
                    os.rename(history_path, desired_path)
                else:
                    for f_name in os.listdir(history_path):
                        srcf = os.path.join(history_path, f_name)
                        dstf = os.path.join(desired_path, f_name)
                        safe_copy(srcf, dstf)
                    try: os.rmdir(history_path)
                    except: pass
                df.loc[df["history folder name"] == history, "history folder name"] = desired
                processed_folders[key] = True
            os.makedirs(desired_path, exist_ok=True)
            df.at[idx, "history folder name"] = desired
        except Exception as e:
            log_msg(f"ERROR row {idx}: {e}")
            log_error(f"P1 ERROR row {idx}: {e}")

    for idx, row in tqdm(df.iterrows(), total=len(df), desc="[P1] Copying ISOs", ncols=80):
        iso_no = row["Iso no"].strip()
        folder = row["folder name"].strip()
        dest_folder = os.path.join(dest_root, folder)
        if not folder or not iso_no:
            continue
        src_iso = find_iso_on_server(iso_no, server_path)
        if not src_iso:
            os.makedirs(dest_folder, exist_ok=True)
            df.at[idx, "ISO Status"] = "MISSING"
            log_error(f"P1: MISSING ISO {iso_no}")
            continue
        dest_iso = os.path.join(dest_folder, os.path.basename(src_iso))
        try:
            safe_copy(src_iso, dest_iso)
            df.at[idx, "ISO Status"] = "OK"
        except Exception as e:
            log_msg(f"ERROR copying {iso_no}: {e}")
            log_error(f"P1: ERROR copying {iso_no}: {e}")
            df.at[idx, "ISO Status"] = "MISSING"

    df.to_excel(excel_file, index=False)
    highlight_missing_iso(excel_file)
    log_msg(f"=== P1: ISO MANAGER END ({format_time(time.time() - process_start)}) ===\n")

# ===============================
# P2: GENERATE EXCEL
# ===============================
def generate_excel(dest_root, index_file_path, output_excel_name="output.xlsx"):
    process_start = time.time()
    log_msg("=== P2: GENERATE EXCEL START ===")

    iso_page_map = defaultdict(list)
    try:
        with open(index_file_path, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                parts = line.strip().split()
                if len(parts) >= 2:
                    pdf_file = parts[0].lower()
                    page = parts[1]
                    if page.isdigit():
                        iso_page_map[pdf_file].append(int(page))
    except FileNotFoundError:
        log_msg(f"❌ Index file not found: {index_file_path}")
        return

    pdf_pattern = re.compile(r'\(([^)]+)\)\.pdf$', re.IGNORECASE)
    processed = 0
    for root, dirs, _ in os.walk(dest_root):
        for dir_name in tqdm(dirs, desc="[P2] Generating Excel", ncols=80):
            subdir = os.path.join(root, dir_name)
            excel_path = os.path.join(subdir, output_excel_name)
            iso_set = set()
            
            # Look for ISOs in the folder (ignoring the Combined.pdf if present)
            pdf_files = [f for f in os.listdir(subdir) if f.lower().endswith('.pdf') and "combined" not in f.lower()]
            for file in pdf_files:
                match = pdf_pattern.search(file)
                if match:
                    # Extract the ISO part (e.g., from (ISO-1234-A).pdf -> ISO-1234)
                    seg = match.group(1).split('-')
                    if len(seg) >= 2:
                        iso_set.add(f"{seg[0]}-{seg[1]}")
                        
            if not iso_set:
                continue
                
            existing_iso = set()
            df_existing = pd.DataFrame()
            if os.path.exists(excel_path):
                try:
                    df_existing = pd.read_excel(excel_path)
                    if "ISO LIST" in df_existing.columns:
                        existing_iso = set(df_existing["ISO LIST"].dropna().astype(str))
                except Exception:
                    # Handle corrupted or unreadable Excel, treat as new
                    df_existing = pd.DataFrame()
            
            new_isos = [iso for iso in sorted(iso_set) if iso not in existing_iso]
            if not new_isos:
                continue
                
            pdf_pages = []
            for iso in new_isos:
                pages = []
                # Check for ISO pages in the index map
                for pdf_file, pnums in iso_page_map.items():
                    if iso.lower() in pdf_file:
                        pages.extend(pnums)
                pdf_pages.append(",".join(str(p) for p in sorted(set(pages))))
                
            df_new = pd.DataFrame({"ISO LIST": new_isos, "PDF PAGE": pdf_pages})
            
            if not df_existing.empty:
                df_final = pd.concat([df_existing, df_new], ignore_index=True)
            else:
                df_final = df_new
                
            # Ensure "ISO Status" column exists for P7 verification
            if "ISO Status" not in df_final.columns:
                df_final["ISO Status"] = ""
                
            df_final.to_excel(excel_path, index=False)
            processed += 1
            log_msg(f"Updated Excel: {subdir}")

    log_msg(f"=== P2: GENERATE EXCEL END ({format_time(time.time() - process_start)}) - {processed} folders ===\n")

# ===============================
# P3: EXTRACT PAGES
# ===============================
def extract_pages(dest_root, master_pdf, excel_filename="output.xlsx"):
    process_start = time.time()
    log_msg("=== P3: EXTRACT PAGES START ===")

    if not os.path.exists(master_pdf):
        log_msg(f"ERROR: Master PDF not found: {master_pdf}")
        return

    try:
        pdf_reader = PyPDF2.PdfReader(master_pdf)
        total_pages = len(pdf_reader.pages)
    except Exception as e:
        log_msg(f"ERROR reading master PDF: {e}")
        return

    for root, dirs, _ in os.walk(dest_root):
        for dir_name in tqdm(dirs, desc="[P3] Extracting Pages", ncols=80):
            subdir = os.path.join(root, dir_name)
            excel_path = os.path.join(subdir, excel_filename)
            if not os.path.exists(excel_path):
                continue
            try:
                df = pd.read_excel(excel_path)
                if "PDF PAGE" not in df.columns:
                    continue
                for idx, row in df.iterrows():
                    pages_str = str(row["PDF PAGE"])
                    if not pages_str or pages_str.lower() == "nan" or not pages_str.strip():
                        continue
                    
                    # Process page numbers from the Excel cell
                    for p in pages_str.split(','):
                        p = p.strip()
                        if p.isdigit():
                            page_num = int(p)
                            # PyPDF2 is 0-indexed, so page 1 is index 0.
                            if 1 <= page_num <= total_pages:
                                out_pdf = os.path.join(subdir, f"{page_num}.pdf")
                                if os.path.exists(out_pdf):
                                    continue # Skip if page already extracted
                                    
                                pdf_writer = PyPDF2.PdfWriter()
                                pdf_writer.add_page(pdf_reader.pages[page_num-1])
                                
                                with open(out_pdf, 'wb') as f:
                                    pdf_writer.write(f)
            except Exception as e:
                log_error(f"P3: Error processing {subdir}: {e}")

    log_msg(f"=== P3: EXTRACT PAGES END ({format_time(time.time() - process_start)}) ===\n")
    
# ===============================
# P4: FRI COPIES (MODIFIED)
# Saves copy with _FRI suffix in the same folder.
# ===============================
def fri_copies(dest_root):
    process_start = time.time()
    log_msg("=== P4: FRI COPIES START (Direct Copy with _FRI Suffix) ===")
    
    # Files to exclude from copying (these are typically generated later or are temporary)
    EXCLUDE_FILES = ["combined.pdf", "output.xlsx"]
    
    for root, dirs, _ in os.walk(dest_root):
        for dir_name in tqdm(dirs, desc="[P4] Creating FRI Copies", ncols=80):
            subdir = os.path.join(root, dir_name)
            
            # Find all PDF files that are not explicitly excluded
            pdf_files = [
                f for f in os.listdir(subdir) 
                if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(subdir, f)) and f.lower() not in EXCLUDE_FILES
            ]
            
            for pdf in pdf_files:
                src = os.path.join(subdir, pdf)
                
                # Create the new destination name with _FRI suffix
                base, ext = os.path.splitext(pdf)
                # Check if it already has _FRI suffix (to prevent multiple copies in one run)
                if base.lower().endswith("_fri"):
                    continue
                
                dst_name = f"{base}_FRI{ext}"
                dst = os.path.join(subdir, dst_name)
                
                try:
                    safe_copy(src, dst)
                except Exception as e:
                    log_error(f"P4: Error copying {src} -> {dst}: {e}")
                    
    log_msg(f"=== P4: FRI COPIES END ({format_time(time.time() - process_start)}) ===\n")


# ===============================
# P5: CLEANUP REDUNDANT FILES
# ===============================
def cleanup_redundancy(dest_root, excel_filename="output.xlsx"):
    process_start = time.time()
    log_msg("=== P5: CLEANUP REDUNDANCY START ===")
    
    # Pattern to match original ISO files, excluding page extracts and FRI copies
    ISO_PATTERN = re.compile(r'\(([^)]+)\)\.pdf$', re.IGNORECASE)
    
    for root, dirs, _ in os.walk(dest_root):
        for dir_name in tqdm(dirs, desc="[P5] Cleaning Redundant PDFs", ncols=80):
            subdir = os.path.join(root, dir_name)
            excel_path = os.path.join(subdir, excel_filename)
            if not os.path.exists(excel_path):
                continue
                
            try:
                df = pd.read_excel(excel_path)
                if "ISO LIST" not in df.columns:
                    continue
                    
                # Get the set of valid ISOs that MUST be kept
                valid_isos = set(df["ISO LIST"].dropna().astype(str))
                
                pdf_files = [f for f in os.listdir(subdir) if f.lower().endswith(".pdf")]
                
                for pdf in pdf_files:
                    # Only check files that look like original ISOs (e.g., have the parenthesized number)
                    match = ISO_PATTERN.search(pdf)
                    
                    # Exclude extracted pages (e.g., 10.pdf) and FRI copies (handled by checking for _FRI suffix)
                    if match and "_FRI" not in pdf.upper():
                        
                        # Extract the base ISO number (e.g., ISO-1234)
                        iso_match_part = match.group(1).split('-')
                        iso = '-'.join(iso_match_part[:2])
                        
                        if iso not in valid_isos:
                            # This ISO is redundant, delete it.
                            file_to_delete = os.path.join(subdir, pdf)
                            try: 
                                os.remove(file_to_delete)
                                log_msg(f"P5: Deleted redundant ISO {pdf} in {dir_name}")
                            except Exception as e:
                                log_error(f"P5: Could not delete {pdf} in {subdir}: {e}")
                                
            except Exception as e:
                log_error(f"P5: Error in {subdir}: {e}")
                
    log_msg(f"=== P5: CLEANUP REDUNDANCY END ({format_time(time.time() - process_start)}) ===\n")


# ===============================
# P6: COMBINE PDFs
# ===============================
def combine_pdfs(dest_root, combined_name="Combined.pdf"):
    process_start = time.time()
    log_msg("=== P6: COMBINE PDFs START ===")
    
    for root, dirs, _ in os.walk(dest_root):
        for dir_name in tqdm(dirs, desc="[P6] Combining PDFs", ncols=80):
            subdir = os.path.join(root, dir_name)
            
            # Get files, prioritizing extracted pages (numeric names) first, then ISOs, excluding FRI copies.
            all_pdfs = [f for f in os.listdir(subdir) if f.lower().endswith(".pdf") and "_FRI" not in f.upper() and "COMBINED" not in f.upper()]
            
            # Sort: Numeric files (pages) come first, then ISO files alphabetically.
            def sort_key(f):
                try:
                    # Pages (e.g., 10.pdf) get numeric priority
                    return (0, int(os.path.splitext(f)[0]))
                except ValueError:
                    # ISOs and others come after pages
                    return (1, f.lower())

            pdf_files = sorted(all_pdfs, key=sort_key)
            
            if not pdf_files:
                continue
                
            combined_path = os.path.join(subdir, combined_name)
            pdf_writer = PyPDF2.PdfWriter()
            
            # Set to track pages added to prevent duplicates from the same source PDF
            added_pages = set() 
            
            for pdf in pdf_files:
                pdf_path = os.path.join(subdir, pdf)
                try:
                    reader = PyPDF2.PdfReader(pdf_path)
                    for p in range(len(reader.pages)):
                        key = (pdf, p)
                        if key not in added_pages:
                            pdf_writer.add_page(reader.pages[p])
                            added_pages.add(key)
                except Exception as e:
                    log_error(f"P6: Could not read {pdf_path}: {e}")
            
            # Only write if there's content
            if pdf_writer.pages:
                try:
                    with open(combined_path, 'wb') as f:
                        pdf_writer.write(f)
                except Exception as e:
                    log_error(f"P6: Could not write combined PDF in {subdir}: {e}")
                    
    log_msg(f"=== P6: COMBINE PDFs END ({format_time(time.time() - process_start)}) ===\n")


# ===============================
# P7: FINAL CLEANUP & VERIFICATION
# ===============================
def final_cleanup(dest_root, excel_filename="output.xlsx"):
    process_start = time.time()
    log_msg("=== P7: FINAL CLEANUP & VERIFICATION START ===")
    
    # Pattern to match original ISO files
    ISO_PATTERN = re.compile(r'\(([^)]+)\)\.pdf$', re.IGNORECASE)
    
    for root, dirs, _ in os.walk(dest_root):
        for dir_name in tqdm(dirs, desc="[P7] Verifying Excel vs PDFs", ncols=80):
            subdir = os.path.join(root, dir_name)
            excel_path = os.path.join(subdir, excel_filename)
            if not os.path.exists(excel_path):
                continue
                
            try:
                df = pd.read_excel(excel_path)
                if "ISO LIST" not in df.columns:
                    continue
                    
                pdf_files = [f for f in os.listdir(subdir) if f.lower().endswith(".pdf")]
                pdf_isos = set()
                
                # Identify which ISOs are actually present in the folder after cleanup
                for pdf in pdf_files:
                    match = ISO_PATTERN.search(pdf)
                    if match and "_FRI" not in pdf.upper():
                        # Extract the base ISO number
                        iso_match_part = match.group(1).split('-')
                        pdf_isos.add('-'.join(iso_match_part[:2]))
                        
                # Update the ISO Status in the Excel
                df["ISO Status"] = df["ISO LIST"].apply(lambda x: "OK" if str(x) in pdf_isos else "MISSING")
                
                df.to_excel(excel_path, index=False)
                highlight_missing_iso(excel_path) # Highlight rows in the local output.xlsx
                
            except Exception as e:
                log_error(f"P7: Error in {subdir}: {e}")
                
    log_msg(f"=== P7: FINAL CLEANUP & VERIFICATION END ({format_time(time.time() - process_start)}) ===\n")

# ===============================
# Main Orchestrator Execution Logic
# ===============================

def main():
    log_msg("Starting Master Orchestrator Workflow (v3.1)...")
    
    # 1. Select Primary Paths via GUI
    
    # Path 1: Input Excel (the task list)
    excel_file = select_file("1. Select Input Excel (must be saved first)", [("Excel files", "*.xlsx")])
    if not excel_file: 
        log_msg("Workflow aborted: Input Excel not selected.")
        return
    
    # Path 2: ISO Server (source documents)
    server_path = select_folder("2. Select ISO Server Root Folder")
    if not server_path: 
        log_msg("Workflow aborted: ISO Server path not selected.")
        return

    # Path 3: Master Index File (New requirement)
    index_file_path = select_file("3. Select Master Index File (text file with ISO/Page mapping)", [("Text files", "*.txt"), ("All files", "*.*")])
    if not index_file_path: 
        log_msg("Workflow aborted: Master Index File not selected.")
        return

    # Path 4: Master PDF (source for extracted pages)
    master_pdf = select_file("4. Select Master PDF Document", [("PDF files", "*.pdf")])
    if not master_pdf: 
        log_msg("Workflow aborted: Master PDF not selected.")
        return
    
    # Path 5: Destination Root Folder (where all work will be done)
    dest_root = select_folder("5. Select Destination Root Folder")
    if not dest_root: 
        log_msg("Workflow aborted: Destination Root Folder not selected.")
        return
    
    log_msg(f"Paths selected:\n  Excel: {excel_file}\n  Server: {server_path}\n  Index: {index_file_path}\n  Master PDF: {master_pdf}\n  Destination: {dest_root}")

    # Run Pre-flight Check (Excel validation and structure update)
    log_msg("Running Pre-flight Check (Excel validation)...")
    if not create_or_update_excel(excel_file):
        # Function prints a message and shows a Tkinter box if file is missing/created.
        log_msg("Pre-flight check failed. Please fill the Excel and re-run.")
        return

    # 2. Execute Processes P1-P7
    
    try:
        # P1: ISO Manager - Creates folders, renames old ones, copies source ISOs
        iso_manager(excel_file, server_path, dest_root)
        
        # P2: Generate Excel - Reads Index File, creates local output.xlsx with page numbers
        generate_excel(dest_root, index_file_path)
        
        # P3: Extract Pages - Uses output.xlsx to extract pages from Master PDF
        extract_pages(dest_root, master_pdf)
        
        # P4: FRI Copies - Creates backup copies of all main PDFs with _FRI suffix
        fri_copies(dest_root)
        
        # P5: Cleanup Redundancy - Deletes unneeded original ISO files
        cleanup_redundancy(dest_root)
        
        # P6: Combine PDFs - Merges all remaining PDFs into Combined.pdf
        combine_pdfs(dest_root)
        
        # P7: Final Cleanup & Verification - Final status update and highlight
        final_cleanup(dest_root)
        
        log_msg("--- ✅ Workflow Complete Successfully ---")
        messagebox.showinfo("Success", "Master Orchestrator Workflow Completed Successfully. Check orchestrator_log.txt for details.")

    except Exception as e:
        log_msg(f"--- ❌ FATAL ERROR IN WORKFLOW EXECUTION ---: {e}")
        log_error(f"FATAL ERROR: {e}")
        messagebox.showerror("Error", f"A fatal error occurred: {e}. Check error_report.txt for details.")


if __name__ == "__main__":
    main()

