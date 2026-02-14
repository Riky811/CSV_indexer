param(
    [string]$TargetFolder = ".",
    [string]$OutputPath = "",
    [switch]$ToXlsx
)

# Resolve root folder
$rootResolved = Resolve-Path -LiteralPath $TargetFolder -ErrorAction Stop
$root = $rootResolved.Path.TrimEnd('\')

# Prepare output filename if not provided
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$folderName = Split-Path -LiteralPath $root -Leaf
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path -Path (Get-Location) -ChildPath ("{0}_index_{1}.csv" -f $folderName, $timestamp)
}

Write-Host "Indexing: $root"
Write-Host "Output CSV: $OutputPath"

# Gather items recursively. For directories, Size will be blank. For files, show Length.
# Capture errors silently to avoid stopping on permissions issues.
$items = Get-ChildItem -LiteralPath $root -Force -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    $full = $_.FullName
    $rel = if ($full.Length -gt ($root.Length + 1)) { $full.Substring($root.Length + 1) } else { $_.Name }
    [PSCustomObject]@{
        FullName = $full
        RelativePath = $rel
        Name = $_.Name
        ItemType = if ($_.PSIsContainer) { 'Directory' } else { 'File' }
        Size = if ($_.PSIsContainer) { $null } else { $_.Length }
        Extension = if ($_.PSIsContainer) { $null } else { $_.Extension }
        LastWriteTime = $_.LastWriteTime
        CreationTime = $_.CreationTime
        Attributes = $_.Attributes.ToString()
    }
}

# Export to CSV (UTF8). In Windows PowerShell 5.1, -Encoding UTF8 writes BOM which helps Excel detect UTF-8.
$items | Sort-Object ItemType, RelativePath | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV saved to $OutputPath"

if ($ToXlsx) {
    Write-Host "Attempting to convert CSV to XLSX (requires Excel)."
    try {
        $fullCsvPath = (Resolve-Path -LiteralPath $OutputPath).ProviderPath
        $excel = New-Object -ComObject Excel.Application
        $excel.DisplayAlerts = $false
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($fullCsvPath)
        $xlsxPath = [System.IO.Path]::ChangeExtension($fullCsvPath, ".xlsx")
        # 51 = xlOpenXMLWorkbook (xlsx)
        $workbook.SaveAs($xlsxPath, 51)
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) > $null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) > $null
        Write-Host "Saved XLSX at: $xlsxPath"
    } catch {
        Write-Warning "Failed to convert to XLSX: $_"
    }
}
