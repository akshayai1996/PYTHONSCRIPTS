<#
==============================================================================
 Project Name:   Interactive Bulk File Copier
 File Name:      CPP.ps1
 Description:    Copies selected files from a chosen source folder into all
                 subfolders recursively. Fully interactive for non-technical users.

 Author:         Akshay Solanki
 Created on:     19-Oct-2025
 Version:        1.0.0
 License:        Unlicense (Public Domain - Free to use without attribution)

 Features:
     - Select source folder interactively
     - Choose which files to copy from the folder
     - Automatically finds all subfolders
     - Optionally skip overwriting existing files
     - Friendly messages and prompts

 Usage:
     1. Run PowerShell
     2. Navigate to this script's folder
     3. Execute:  .\CPP.ps1
==============================================================================
#>

# ==============================
# Ask user for the source folder
# ==============================
$sourcePath = Read-Host "📂 Enter the full path of the source folder containing files to copy"

if (-not (Test-Path $sourcePath)) {
    Write-Host "🚨 Folder not found. Exiting."
    exit
}

# ==============================
# List all files in the source folder
# ==============================
$allFiles = Get-ChildItem -Path $sourcePath -File
if ($allFiles.Count -eq 0) {
    Write-Host "⚠️ No files found in the source folder. Exiting."
    exit
}

Write-Host "`nFound the following files in the folder:"
for ($i = 0; $i -lt $allFiles.Count; $i++) {
    Write-Host "$($i+1). $($allFiles[$i].Name)"
}

# ==============================
# Ask user which files to copy
# ==============================
$selection = Read-Host "`nEnter the numbers of files to copy, separated by commas (e.g., 1,3,5)"
$selectedIndices = $selection -split "," | ForEach-Object { ($_ -as [int]) - 1 }

$filesToCopy = @()
foreach ($index in $selectedIndices) {
    if ($index -ge 0 -and $index -lt $allFiles.Count) {
        $filesToCopy += $allFiles[$index].Name
    }
}

if ($filesToCopy.Count -eq 0) {
    Write-Host "🚨 No valid files selected. Exiting."
    exit
}

# ==============================
# Ask if user wants to overwrite existing files
# ==============================
$overwriteChoice = Read-Host "`nDo you want to overwrite existing files in subfolders? (Y/N)"
$overwrite = $false
if ($overwriteChoice.ToUpper() -eq "Y") { $overwrite = $true }

# ==============================
# Get all subfolders recursively
# ==============================
$subfolders = Get-ChildItem -Path $sourcePath -Directory -Recurse
if ($subfolders.Count -eq 0) {
    Write-Host "⚠️ No subfolders found under '$sourcePath'. Exiting."
    exit
}

Write-Host "`n✅ Starting copy of $($filesToCopy.Count) file(s) to $($subfolders.Count) subfolders..."

# ==============================
# Copy files
# ==============================
foreach ($folder in $subfolders) {
    foreach ($file in $filesToCopy) {
        $sourceFile = Join-Path $sourcePath $file
        $destinationFile = Join-Path $folder.FullName $file

        if ($overwrite -or -not (Test-Path $destinationFile)) {
            Copy-Item -Path $sourceFile -Destination $destinationFile -Force
            Write-Host "✅ Copied '$file' to $($folder.FullName)"
        } else {
            Write-Host "⚠️ '$file' already exists in $($folder.FullName). Skipping."
        }
    }
}

Write-Host "`n🎉 All copy operations completed successfully."
