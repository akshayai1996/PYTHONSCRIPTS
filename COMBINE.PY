"""
==============================================================================
 Project Name:   PDF Automation Tools
 File Name:      combine.py
 Description:    Combines all PDF files in the current working directory into
                 a single PDF file named after the folder itself. Allows
                 merging either alphabetically or in filesystem order.

 Author:         Akshay Solanki
 Created on:     19-Oct-2025
 Version:        1.2.0
 License:        Unlicense (Public Domain - Free to use without attribution)

 Dependencies:   fitz (PyMuPDF)
 Python Version: 3.8+

 Usage:
     python combine.py

 Options:
     MERGE_ALPHABETICALLY = True  → Merge files in alphabetical order (A → Z)
     MERGE_ALPHABETICALLY = False → Merge files in folder order (as listed by OS)

 Customization Guide:
     - Place this script inside the folder containing your PDF files.
     - Open a terminal or command prompt in that folder.
     - Run: python combine.py
     - To change output name, edit: output_filename = f"{folder_name}.pdf"
==============================================================================
"""

# ==============================
# Import Required Libraries
# ==============================
import fitz
import os
import sys

# ==============================
# USER CONFIGURATION
# ==============================
# Choose whether to merge PDFs alphabetically (A → Z) or in folder order
MERGE_ALPHABETICALLY = True  # Set False for folder order merge

def combine_pdfs_in_current_folder():
    """
    Combines all PDF files in the current working directory into one output file,
    with optional alphabetical or folder-order merging.
    """

    # Step 1: Identify the current working directory and folder name
    current_dir = os.getcwd()
    folder_name = os.path.basename(current_dir)

    # Step 2: Define the output file name
    output_filename = f"{folder_name}.pdf"

    # Step 3: Create an empty PDF to store all combined pages
    combined_pdf = fitz.open()

    # Step 4: Get a list of all PDF files in the folder
    pdf_files = [
        f for f in os.listdir(current_dir)
        if f.lower().endswith(".pdf")
        and not f.startswith(".")
        and f != output_filename
    ]

    # Step 5: Handle the case when no PDFs are found
    if not pdf_files:
        print("⚠️ No PDF files found in this directory. Please place some .pdf files here.")
        combined_pdf.close()
        return

    # Step 6: Sort files if user chose alphabetical merge
    if MERGE_ALPHABETICALLY:
        pdf_files.sort()

    print(f"🔍 Found {len(pdf_files)} PDF files. Combining now...\n")
    print(f"📄 Merge order: {'Alphabetical' if MERGE_ALPHABETICALLY else 'Folder order'}\n")

    # Step 7: Loop through each file and merge it
    for filename in pdf_files:
        file_path = os.path.join(current_dir, filename)
        try:
            with fitz.open(file_path) as pdf:
                combined_pdf.insert_pdf(pdf)
            print(f"✅ Added: {filename}")
        except Exception as e:
            print(f"❌ Error processing {filename}: {e}")

    # Step 8: Save the combined PDF
    try:
        combined_pdf.save(output_filename)
        print(f"\n🎉 All PDFs combined successfully into: {output_filename}")
    except Exception as e:
        print(f"🚨 Failed to save combined file: {e}")
    finally:
        combined_pdf.close()


# ==============================
# Script Entry Point
# ==============================
if __name__ == "__main__":
    try:
        combine_pdfs_in_current_folder()
    except Exception as e:
        print(f"\n🚨 Fatal error occurred: {e}")
        sys.exit(1)
