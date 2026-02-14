# CSV_folder-indexer
batch and powershell script that indexes all directories, archives and files within a folder and its subfolders.
Made with love by Riccardo Maria Polidoro.
This repository contains a small Windows-friendly tool to create an Excel-ready index of all files and subfolders inside a given folder.

Files:
- `index_folder.bat` — double-clickable wrapper you can keep with the folder (or run from a command prompt). Calls the PowerShell script.
- `index_folder.ps1` — the PowerShell engine that enumerates files and writes a CSV. Optionally converts the CSV to `.xlsx` using Excel COM if Excel is installed.

What the index contains (columns):
- FullName (absolute path)
- RelativePath (path relative to the root folder)
- Name
- ItemType (File or Directory)
- Size (bytes, files only)
- Extension (files only)
- LastWriteTime
- CreationTime
- Attributes

Usage examples (Windows):

1) Double-click `index_folder.bat` — it will index the current folder (`.`) and save a CSV in the current working folder with a timestamped filename.

2) From Command Prompt or PowerShell, specify a target folder:

```powershell
# Index a specific folder and let the script generate the output filename
index_folder.bat "D:\MyArchive\ProjectFolder"

# Index and specify output CSV path
index_folder.bat "D:\MyArchive\ProjectFolder" "D:\Indexes\ProjectFolder_index.csv"

# Index and request XLSX conversion (requires Excel installed)
index_folder.bat "D:\MyArchive\ProjectFolder" "D:\Indexes\ProjectFolder_index.csv" /xlsx
```

Notes and limitations:
- The script is written to be simple and portable. It uses PowerShell v5.1+ features. On most Windows 10/11 systems PowerShell v5.1 is available.
- CSV is the default output because it's universally readable by Excel. The optional XLSX conversion depends on Excel being installed and available via COM automation.
- Directory `Size` is left blank. Computing recursive directory sizes can be slow on large trees; if you want folder sizes, I can add an optional (slower) mode that aggregates file sizes per folder.
- The script attempts to handle paths with spaces by quoting arguments; keep the `.bat` and `.ps1` together in the same folder.

If you'd like any changes (add file checksums, include owner/permissions, compute directory sizes, or output to a single `.xlsx` natively without requiring Excel), tell me which feature to add and I'll update the scripts.

Bulk CSV → XLSX conversion
--------------------------

If you want to convert many CSV indexes (for example the index for a folder plus CSVs inside subfolders) to `.xlsx`, there are two good options:

1) Use the built-in `/xlsx` flag when calling `index_folder.bat` — this converts the single CSV the script produces to `.xlsx` (requires Excel installed).

2) Convert many CSV files at once with the included `convert_csvs_to_xlsx.ps1` script. Example usages:

```powershell
# Convert all CSVs in a folder (non-recursive)
.\convert_csvs_to_xlsx.ps1 -Root "E:\Indexes"

# Convert all CSVs in a folder and subfolders (recursive)
.\convert_csvs_to_xlsx.ps1 -Root "E:\Indexes" -Recursive

# Force overwrite of existing .xlsx files
.\convert_csvs_to_xlsx.ps1 -Root "E:\Indexes" -Recursive -Force
```

Notes:
- The converter uses Excel COM automation, so Excel must be installed on the machine where you run it.
