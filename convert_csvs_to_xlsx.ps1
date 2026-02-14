<#
.SYNOPSIS
    Convert CSV files to XLSX using Excel COM automation.

.DESCRIPTION
    Finds CSV files under a given folder (optionally recursively) and converts each to an .xlsx file using Excel COM.
    Skips files where an .xlsx already exists unless -Force is provided.

.PARAMETER Root
    The root folder to scan for CSV files. Defaults to current directory.

.PARAMETER Recursive
    If specified, searches subfolders recursively.

.PARAMETER Force
    Overwrite existing .xlsx files if present.

.PARAMETER WhatIf
    PowerShell's built-in WhatIf support (no-op preview).

.EXAMPLE
    .\convert_csvs_to_xlsx.ps1 -Root 'E:\Indexes' -Recursive

    Converts all CSV files under E:\Indexes and subfolders to XLSX.

.NOTES
    Requires Excel to be installed and available via COM. If Excel isn't available, the script will report and skip conversion.
#>

param(
    [Parameter(Position=0)]
    [string]$Root = ".",
    [switch]$Recursive,
    [switch]$Force
)

Set-StrictMode -Version Latest

try {
    $rootResolved = Resolve-Path -LiteralPath $Root -ErrorAction Stop
    $rootPath = $rootResolved.Path
} catch {
    Write-Error "Cannot resolve root path '$Root': $_"
    exit 2
}

$searchOption = if ($Recursive) { '-Recurse' } else { '' }
Write-Host "Scanning for CSV files in: $rootPath (Recursive: $($Recursive.IsPresent))"

# Build file list
if ($Recursive) {
    $csvFiles = Get-ChildItem -Path $rootPath -Filter '*.csv' -File -Recurse -ErrorAction SilentlyContinue
} else {
    $csvFiles = Get-ChildItem -Path $rootPath -Filter '*.csv' -File -ErrorAction SilentlyContinue
}

if (-not $csvFiles -or $csvFiles.Count -eq 0) {
    Write-Host "No CSV files found under $rootPath"
    exit 0
}

# Try to create Excel COM object once
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.DisplayAlerts = $false
    $excel.Visible = $false
} catch {
    Write-Warning "Excel COM object not available. Conversion requires Excel installed. Error: $_"
    exit 3
}

$converted = 0
$skipped = 0
$failed = 0

foreach ($file in $csvFiles) {
    $csvPath = $file.FullName
    $xlsxPath = [System.IO.Path]::ChangeExtension($csvPath, '.xlsx')

    if ((-not $Force) -and (Test-Path $xlsxPath)) {
        Write-Host "Skipping (xlsx exists): $csvPath"
        $skipped++
        continue
    }

    Write-Host "Converting: $csvPath -> $xlsxPath"
    try {
        # Workbooks.Open handles CSV; the locale/delimiter may affect parsing but in most cases Windows CSV opens fine.
        $workbook = $excel.Workbooks.Open($csvPath)
        # 51 = xlOpenXMLWorkbook (xlsx)
        $workbook.SaveAs($xlsxPath, 51)
        $workbook.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) > $null
        $converted++
    } catch {
        Write-Warning "Failed to convert '$csvPath': $_"
        $failed++
    }
}

# Clean up Excel COM object
try {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) > $null
} catch {
    # ignore
}

Write-Host "Conversion completed. Converted: $converted; Skipped: $skipped; Failed: $failed"

if ($failed -gt 0) { exit 4 } else { exit 0 }
