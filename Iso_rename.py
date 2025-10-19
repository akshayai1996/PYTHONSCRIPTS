"""
===============================================================================
 Project Name:    ISO Filename Shortener & Smart Renamer (Safe v2.2)
 File Name:       Iso_rename.py
 Description:     Phase-based automation for shortening and renaming ISO-style
                  filenames based on a structured pattern.
                  
                  ➤ Phase 1: Scans a folder and generates/updates an Excel
                    summary (summary.xlsx) with columns:
                      - Current Filename
                      - Shortened Name
                      - Last Processed Name
                    
                  ➤ Phase 2: Reads the reviewed Excel file and renames only
                    those files whose 'Shortened Name' has changed, appending
                    it inside parentheses (e.g., "aa-bb-uvw-mno-pqr.pdf" →
                    "aa-bb-uvw-mno-pqr(aa-bb-pqr).pdf").
 
 Author:          Akshay Solanki
 Dependencies:    pandas, openpyxl, os
 Created On:      19-Oct-2025
===============================================================================
"""

import os
import pandas as pd

# === CONFIGURATION ===
FOLDER_PATH = r"D:\ISO_Files"  # 🔹 Change this to your working folder path
EXCEL_FILE = os.path.join(FOLDER_PATH, "summary.xlsx")
VALID_EXT = ".pdf"  # 🔸 Change to ".iso" if you’re processing ISO drawings


# ---------------------------------------------------------------------------
# Helper: Extract shortened name
# ---------------------------------------------------------------------------
def get_short_name(filename: str) -> str:
    """Extracts short name (first two + last part) from a hyphen-separated filename."""
    name = os.path.splitext(filename)[0]
    parts = name.split('-')
    if len(parts) < 3:
        return name
    return '-'.join([parts[0], parts[1], parts[-1]])


# ---------------------------------------------------------------------------
# PHASE 1 — Create or update Excel summary only
# ---------------------------------------------------------------------------
def generate_or_update_summary() -> pd.DataFrame:
    """Scans folder, creates or updates summary Excel with new/deleted entries."""
    files = [
        f for f in os.listdir(FOLDER_PATH)
        if os.path.isfile(os.path.join(FOLDER_PATH, f)) and f.lower().endswith(VALID_EXT)
    ]

    new_data = [{"Current Filename": f, "Shortened Name": get_short_name(f), "Last Processed Name": ""} for f in files]
    new_df = pd.DataFrame(new_data)

    if not os.path.exists(EXCEL_FILE):
        new_df.to_excel(EXCEL_FILE, index=False)
        print(f"✅ Created new summary file: {EXCEL_FILE}")
        print("👉 Review and edit 'Shortened Name' column if needed, then rerun for rename.")
        return pd.DataFrame()

    old_df = pd.read_excel(EXCEL_FILE)
    merged = pd.merge(new_df, old_df, on="Current Filename", how="left", suffixes=("", "_old"))

    merged["Shortened Name"] = merged["Shortened Name_old"].combine_first(merged["Shortened Name"])
    merged["Last Processed Name"] = merged["Last Processed Name_old"].combine_first(merged["Last Processed Name"])
    merged = merged[["Current Filename", "Shortened Name", "Last Processed Name"]]

    # Remove deleted files
    missing = set(old_df["Current Filename"]) - set(files)
    if missing:
        print(f"⚠️ Removing {len(missing)} entries for deleted files.")
        merged = merged[~merged["Current Filename"].isin(missing)]

    merged.to_excel(EXCEL_FILE, index=False)
    print("✅ Excel summary updated with all files.")
    print("👉 Please review and correct 'Shortened Name' values before Phase 2.")
    return merged


# ---------------------------------------------------------------------------
# PHASE 2 — Rename changed files (only after user confirmation)
# ---------------------------------------------------------------------------
def rename_changed_files():
    """Renames only files whose 'Shortened Name' differs from 'Last Processed Name'."""
    if not os.path.exists(EXCEL_FILE):
        print("⚠️ No summary found. Run script once to generate it first.")
        return

    df = pd.read_excel(EXCEL_FILE)
    changed = df[(df["Shortened Name"] != df["Last Processed Name"]) | (df["Last Processed Name"].isna())]

    if changed.empty:
        print("✅ All files already up to date. No renames needed.")
        return

    print(f"\n⚙️ {len(changed)} file(s) require renaming based on updated Excel.")
    confirm = input("Proceed with renaming? (y/n): ").strip().lower()
    if confirm != "y":
        print("🚫 Rename operation cancelled by user.")
        return

    for idx, row in changed.iterrows():
        current_filename = str(row["Current Filename"]).strip()
        short = str(row["Shortened Name"]).strip()
        old_path = os.path.join(FOLDER_PATH, current_filename)

        if not os.path.exists(old_path):
            print(f"⚠️ File not found: {current_filename}. Removing entry.")
            df.drop(idx, inplace=True)
            continue

        name_no_ext, ext = os.path.splitext(current_filename)
        if "(" in name_no_ext and name_no_ext.endswith(")"):
            base_name = name_no_ext[:name_no_ext.rfind("(")].rstrip()
        else:
            base_name = name_no_ext

        new_name = f"{base_name}({short}){ext}"
        new_path = os.path.join(FOLDER_PATH, new_name)

        if new_name == current_filename:
            df.at[idx, "Last Processed Name"] = short
            continue

        try:
            os.rename(old_path, new_path)
            print(f"✅ Renamed: {current_filename} → {new_name}")
            df.at[idx, "Current Filename"] = new_name
            df.at[idx, "Last Processed Name"] = short
        except Exception as e:
            print(f"❌ Error renaming {current_filename}: {e}")

    df.to_excel(EXCEL_FILE, index=False)
    print("\n📝 Rename operation completed and Excel updated.")


# ---------------------------------------------------------------------------
# MAIN EXECUTION FLOW
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("=== ISO Smart Filename Shortener v2.2 ===")
    print("Choose operation mode:")
    print("1️⃣  Phase 1 — Create/Update Excel Summary")
    print("2️⃣  Phase 2 — Rename Files from Excel")

    choice = input("\nEnter 1 or 2: ").strip()

    if choice == "1":
        generate_or_update_summary()
    elif choice == "2":
        rename_changed_files()
    else:
        print("❌ Invalid input. Please enter 1 or 2.")