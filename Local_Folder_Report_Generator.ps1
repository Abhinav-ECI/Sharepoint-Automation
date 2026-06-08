param(
    [Parameter(Mandatory=$false, HelpMessage="Local source folder path to scan")]
    [string]$SourcePath = ""
)

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = Read-Host "Enter local source folder path to scan"
}

if (-not (Test-Path $SourcePath)) {
    Write-Host "Source path not found: $SourcePath" -ForegroundColor Red
    exit 1
}

Write-Host "Scanning local source path: $SourcePath" -ForegroundColor Cyan

$AllFolders = @()
$AllFiles = @()
$TotalFolders = 0
$TotalFiles = 0
$TotalSize = 0
$MaxDepth = 0

# Enumerate files and folders
try {
    $files = Get-ChildItem -Path $SourcePath -Recurse -File -Force -ErrorAction Stop
    $folders = Get-ChildItem -Path $SourcePath -Recurse -Directory -Force -ErrorAction Stop
} catch {
    Write-Host "ERROR enumerating files/folders: $_" -ForegroundColor Red
    exit 1
}

foreach ($f in $files) {
    $full = $f.FullName
    $rel = $full.Substring($SourcePath.Length).TrimStart('\','/')

    if ($rel -ne '') { $segments = $rel -split '[\\/]'; $depth = ($segments.Length - 1) } else { $depth = 0 }
    if ($depth -gt $MaxDepth) { $MaxDepth = $depth }

    $folderFull = Split-Path $full -Parent

    $folderObj = $AllFolders | Where-Object { $_.Path -eq $folderFull }
    if (-not $folderObj) {
        # Count direct children (files + subfolders) for this folder
        $directChildren = @()
        $directChildren += @($files | Where-Object { (Split-Path $_.FullName -Parent) -eq $folderFull })
        $directChildren += @($folders | Where-Object { $_.Parent.FullName -eq $folderFull })
        $itemCount = $directChildren.Count
        
        $folderObj = [PSCustomObject]@{
            Library = 'Local'
            FolderName = Split-Path $folderFull -Leaf
            FileCount = 0
            ItemCount = $itemCount
            Depth = $depth
            SizeBytes = 0
            SizeMB = 0
            SizeGB = 0
            Path = $folderFull
        }
        $AllFolders += $folderObj
    }

    # update the folder stats
    $folderObj.FileCount = $folderObj.FileCount + 1
    $folderObj.SizeBytes = $folderObj.SizeBytes + $f.Length
    $folderObj.SizeMB = [math]::Round($folderObj.SizeBytes / 1MB, 4)
    $folderObj.SizeGB = [math]::Round($folderObj.SizeBytes / 1GB, 6)

    # add file entry
    $AllFiles += [PSCustomObject]@{
        Library = 'Local'
        FileName = $f.Name
        SizeBytes = $f.Length
        SizeMB = [math]::Round($f.Length / 1MB, 4)
        Path = $full
    }

    $TotalFiles++
    $TotalSize += $f.Length
}

$TotalFolders = $AllFolders.Count

Write-Host "Local scan complete. Folders: $TotalFolders | Files: $TotalFiles | Size: $([math]::Round($TotalSize/1MB,4)) MB" -ForegroundColor Green
Write-Host "Maximum folder depth found: $MaxDepth" -ForegroundColor Green

# Export to Excel
$desktop = [Environment]::GetFolderPath('Desktop')
$ExportPath = Join-Path $desktop ("Local_Folder_Report_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

try {
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Workbook = $Excel.Workbooks.Add()
    while ($Workbook.Sheets.Count -lt 3) { $Workbook.Sheets.Add() | Out-Null }

    $desired = @("Summary", "Folders", "Root Files")
    for ($i = 0; $i -lt $desired.Count; $i++) {
        $idx = $i + 1
        $targetName = $desired[$i]
        $conflict = $null
        for ($j = 1; $j -le $Workbook.Sheets.Count; $j++) {
            $s = $Workbook.Sheets.Item($j)
            if ($s.Name -eq $targetName -and $j -ne $idx) { $conflict = $s; break }
        }
        if ($conflict) { $conflict.Name = "$($targetName)_tmp_$(Get-Random)" }
        $Workbook.Sheets.Item($idx).Name = $targetName
    }

    $SheetSummary = $Workbook.Sheets.Item("Summary")
    $SheetFolders = $Workbook.Sheets.Item("Folders")
    $SheetRoot = $Workbook.Sheets.Item("Root Files")

    # Summary
    $SheetSummary.Cells.Item(1,1) = "Metric"
    $SheetSummary.Cells.Item(1,2) = "Value"
    $HeaderRange1 = $SheetSummary.Range("A1:B1")
    $HeaderRange1.Font.Bold = $true
    $HeaderRange1.Interior.ColorIndex = 15

    $SheetSummary.Cells.Item(2,1) = "Total Folders"
    $SheetSummary.Cells.Item(2,2) = $TotalFolders
    $SheetSummary.Cells.Item(3,1) = "Total Files"
    $SheetSummary.Cells.Item(3,2) = $TotalFiles
    $SheetSummary.Cells.Item(4,1) = "Total Size (GB)"
    $SheetSummary.Cells.Item(4,2) = [math]::Round($TotalSize / 1GB, 4)
    $SheetSummary.Cells.Item(5,1) = "Total Size (MB)"
    $SheetSummary.Cells.Item(5,2) = [math]::Round($TotalSize / 1MB, 4)
    $SheetSummary.Cells.Item(6,1) = "Max Depth"
    $SheetSummary.Cells.Item(6,2) = $MaxDepth
    $SheetSummary.Columns.Item(1).AutoFit() | Out-Null
    $SheetSummary.Columns.Item(2).AutoFit() | Out-Null

    # Folders sheet
    $SheetFolders.Cells.Item(1,1) = "Library"
    $SheetFolders.Cells.Item(1,2) = "Folder Name"
    $SheetFolders.Cells.Item(1,3) = "File Count"
    $SheetFolders.Cells.Item(1,4) = "Item Count"
    $SheetFolders.Cells.Item(1,5) = "Depth"
    $SheetFolders.Cells.Item(1,6) = "Size (MB)"
    $SheetFolders.Cells.Item(1,7) = "Size (GB)"
    $SheetFolders.Cells.Item(1,8) = "Path"
    $HeaderRange2 = $SheetFolders.Range("A1:H1")
    $HeaderRange2.Font.Bold = $true
    $HeaderRange2.Interior.ColorIndex = 15
    $r = 2
    foreach ($Folder in $AllFolders) {
        $SheetFolders.Cells.Item($r,1) = $Folder.Library
        $SheetFolders.Cells.Item($r,2) = $Folder.FolderName
        $SheetFolders.Cells.Item($r,3) = $Folder.FileCount
        $SheetFolders.Cells.Item($r,4) = $Folder.ItemCount
        $SheetFolders.Cells.Item($r,5) = $Folder.Depth
        $SheetFolders.Cells.Item($r,6) = $Folder.SizeMB
        $SheetFolders.Cells.Item($r,7) = $Folder.SizeGB
        $SheetFolders.Cells.Item($r,8) = $Folder.Path
        $r++
    }
    for ($i = 1; $i -le 8; $i++) { $SheetFolders.Columns.Item($i).AutoFit() | Out-Null }

    # Root Files sheet
    if ($AllFiles.Count -gt 0) {
        $SheetRoot.Cells.Item(1,1) = "Library"
        $SheetRoot.Cells.Item(1,2) = "File Name"
        $SheetRoot.Cells.Item(1,3) = "Size (MB)"
        $SheetRoot.Cells.Item(1,4) = "Path"
        $HeaderRange3 = $SheetRoot.Range("A1:D1")
        $HeaderRange3.Font.Bold = $true
        $HeaderRange3.Interior.ColorIndex = 15
        $r = 2
        foreach ($File in $AllFiles) {
            $SheetRoot.Cells.Item($r,1) = $File.Library
            $SheetRoot.Cells.Item($r,2) = $File.FileName
            $SheetRoot.Cells.Item($r,3) = $File.SizeMB
            $SheetRoot.Cells.Item($r,4) = $File.Path
            $r++
        }
        for ($i = 1; $i -le 4; $i++) { $SheetRoot.Columns.Item($i).AutoFit() | Out-Null }
    }

    $Workbook.SaveAs($ExportPath)
    $Excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null

    Write-Host "`n[SUCCESS] Excel file created successfully!`n  Location: $ExportPath" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to create Excel file: $_" -ForegroundColor Red
}
