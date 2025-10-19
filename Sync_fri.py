"""
===============================================================================
 Project Name:    FRI Folder Synchronizer
 File Name:       Sync_fri.py
 Description:     Synchronizes files from a source folder to a destination
                  folder. Only copies new or updated files to save time.
                  Supports subfolder iteration and timestamp checks.

 Author:          Akshay Solanki
 Created on:      19-Oct-2025
 Version:         1.1.0
 License:         Unlicense (Public Domain)

 Usage:
     python Sync_fri.py

 Notes for Non-Technical Users:
     - Only new files or files modified in the source folder will be copied.
     - Folder structure is preserved.
     - Shows real-time progress using a progress bar.
===============================================================================
"""

import os
import shutil
from datetime import datetime
from tqdm import tqdm

# ==============================
# USER CONFIGURATION
# ==============================
source_dir = r"\\in-nayra-fs\Nayara\Piping\PIPING FRI"
destination_dir = r"Z:\PIPING FRI"
# ==============================

def sync_directories(source, destination):
    """
    Synchronizes all files from source to destination.
    Only copies files that are new or updated based on modification time.
    """
    print(f"\n🔄 Sync started at {datetime.now()}\n")
    
    all_files = []
    for root, dirs, files in os.walk(source):
        for file in files:
            all_files.append((root, file))
    
    print(f"📦 Found {len(all_files)} files to check in source folder.\n")
    
    for root, file in tqdm(all_files, desc="Syncing files", unit="file"):
        source_file = os.path.join(root, file)
        relative_path = os.path.relpath(source_file, source)
        destination_file = os.path.join(destination, relative_path)
        
        # Ensure destination subfolder exists
        destination_subdir = os.path.dirname(destination_file)
        if not os.path.exists(destination_subdir):
            os.makedirs(destination_subdir)
        
        # Copy if new or updated
        if (not os.path.exists(destination_file) or
            os.path.getmtime(source_file) > os.path.getmtime(destination_file)):
            shutil.copy2(source_file, destination_file)
            tqdm.write(f"✅ Copied: {relative_path}")
        else:
            tqdm.write(f"⚠️ Skipped (up-to-date): {relative_path}")
    
    print(f"\n✅ Sync completed at {datetime.now()}\n")

if __name__ == "__main__":
    try:
        sync_directories(source_dir, destination_dir)
    except Exception as e:
        print(f"\n🚨 A fatal error occurred during sync: {e}")
