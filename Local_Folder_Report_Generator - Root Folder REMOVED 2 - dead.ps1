param(
    [Parameter(Mandatory=$false, HelpMessage="Local source folder path to scan")]
    [string]$SourcePath = "",
    [Parameter(Mandatory=$false, HelpMessage="Show textual progress bar during scan")]
    [bool]$ShowProgressBar = $true,
    [Parameter(Mandatory=$false, HelpMessage="Pre-count files to enable accurate progress; this still enumerates files once")]
    [bool]$PreCountFiles = $false
    ,
    [Parameter(Mandatory=$false, HelpMessage="Max rows per CSV file (splits when reached)")]
    [int]$MaxRowsPerFile = 1000000
    ,
    [Parameter(Mandatory=$false, HelpMessage="Write parts to local temp dir and copy to final destination (faster for network targets)")]
    [bool]$UseLocalTemp = $true
    ,
    [Parameter(Mandatory=$false, HelpMessage="Local temp directory to use when -UseLocalTemp is set")]
    [string]$LocalTempDir = $env:TEMP
    ,
    [Parameter(Mandatory=$false, HelpMessage="Number of rows between automatic flushes to disk (0 = never flush except on close)")]
    [int]$FlushEvery = 1000
    ,
    [Parameter(Mandatory=$false, HelpMessage="Copy completed parts in background using Start-Job (true) or synchronously (false)")]
    [bool]$CopyInBackground = $true
    ,
    [Parameter(Mandatory=$false, HelpMessage="Write a checkpoint every N rows so the run can resume if interrupted (0 = disabled)")]
    [int]$CheckpointEvery = 1000
    ,
    [Parameter(Mandatory=$false, HelpMessage="Automatically resume from existing checkpoint if found")]
    [bool]$ResumeFromCheckpoint = $true
)

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = Read-Host "Enter local source folder path to scan"
}

if (-not (Test-Path $SourcePath)) {
    Write-Host "Source path not found: $SourcePath" -ForegroundColor Red
    exit 1
}

Write-Host "Scanning local source path: $SourcePath" -ForegroundColor Cyan

# Optional: prompt user for an age-in-years filter (counts files last-modified on or before Dec 31 of (currentYear - N))
$ageInput = Read-Host "Optional: enter number of years (e.g. 5) to count files older than (press Enter to skip)"
$AgeYears = $null
$AgeCutoff = $null
if (-not [string]::IsNullOrWhiteSpace($ageInput)) {
    if ($ageInput -match '^\d+$') {
        $AgeYears = [int]$ageInput
        $currentYear = (Get-Date).Year
        $cutoffYear = $currentYear - $AgeYears
        $AgeCutoff = Get-Date -Year $cutoffYear -Month 12 -Day 31 -Hour 23 -Minute 59 -Second 59
        Write-Host "Filtering: counting files last modified on or before $($AgeCutoff.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
    } else {
        Write-Host "Invalid input for years; skipping age filter." -ForegroundColor Yellow
    }
}

# Streaming scan: avoid loading all files/folders into memory and avoid += on large arrays
$TotalFolders = 0
$TotalFiles = 0
$TotalSize = 0
$MaxDepth = 0
$OldFilesCount = 0
$OldFilesSize = 0

# Helper to repeat a character N times (used for text progress bar)
function Repeat-Char($char, $count) {
    if ($count -le 0) { return '' }
    return -join (1..$count | ForEach-Object { $char })
}

# Helper to escape CSV fields (returns a string)
function EscapeCSV([object]$val) {
    if ($null -eq $val) { return '""' }
    $s = [string]$val
    $s = $s.Replace('"', '""')
    return '"' + $s + '"'
}

$desktop = [Environment]::GetFolderPath('Desktop')
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$base = "Local_Folder_Report_$ts"
$summaryPath = Join-Path $desktop ("${base}_Summary.csv")
$foldersPath = Join-Path $desktop ("${base}_Folders.csv")
 $filesPath = Join-Path $desktop ("${base}_Files.csv")

# Checkpoint and manifest paths
$checkpointPath = Join-Path $desktop ("${base}_checkpoint.json")
$manifestPath = Join-Path $desktop ("${base}_manifest.json")

$EnableFolderReport = $false  # set to $true to enable per-folder stats and folder CSV export
$folderStats = @{}
$processed = 0
$totalEstimated = 0

try {
    # Use .NET EnumerateFiles for robust, lower-overhead enumeration and better long-path handling.
    try {
        $fileEnumerable = [System.IO.Directory]::EnumerateFiles($SourcePath, '*', [System.IO.SearchOption]::AllDirectories)
    } catch {
        # If the path is too long or system rejects it, retry with the long-path (\\?\) prefix.
        try {
            $lp = $SourcePath
            if ($lp.StartsWith('\\')) { $lp = '\\?\UNC\' + $lp.TrimStart('\') } else { $lp = '\\?\' + $lp }
            $fileEnumerable = [System.IO.Directory]::EnumerateFiles($lp, '*', [System.IO.SearchOption]::AllDirectories)
        } catch {
            Write-Host "Warning: Could not enumerate files with .NET EnumerateFiles: $_" -ForegroundColor Yellow
            $fileEnumerable = @()
        }
    }

    if (-not $fileEnumerable) {
        Write-Host "Warning: file enumeration returned no results; proceeding with empty set." -ForegroundColor Yellow
        $fileEnumerable = @()
    }

    # Prepare streaming split CSV writer (creates part files to avoid large single-file memory/size issues)
    $FileReportCount = 0
    $fileIndex = 1
    $fileRowCount = 0
    $generatedFiles = @()
    $filesDir = [System.IO.Path]::GetDirectoryName($filesPath)
    $filesBaseName = [System.IO.Path]::GetFileNameWithoutExtension($filesPath)
    $filesBaseNoExt = [System.IO.Path]::Combine($filesDir, $filesBaseName)

    # Load checkpoint and manifest if requested; prompt the user to resume when a checkpoint exists
    $resuming = $false
    $resumeLastPath = $null
    if (Test-Path $checkpointPath) {
        $answer = Read-Host "Checkpoint found at $checkpointPath. Resume from checkpoint? (Y/N) [Y]"
        if ([string]::IsNullOrWhiteSpace($answer) -or $answer.Trim().ToUpper().StartsWith('Y')) {
            try {
                $ck = Get-Content -Path $checkpointPath -Raw | ConvertFrom-Json
                if ($ck -and $ck.LastProcessedPath) {
                    $resumeLastPath = $ck.LastProcessedPath
                    $fileIndex = [int]$ck.CurrentPartIndex
                    $fileRowCount = [int]$ck.FileRowCount
                    $resuming = $true
                    Write-Host "Resuming from checkpoint. Last processed: $resumeLastPath (part $fileIndex, rows $fileRowCount)" -ForegroundColor Cyan
                }
            } catch {
                Write-Host "Warning: failed to read checkpoint: $_" -ForegroundColor Yellow
            }
        } else {
            Write-Host "Skipping resume from checkpoint as requested by user." -ForegroundColor Cyan
        }
    }

    if (Test-Path $manifestPath) {
        try { $generatedFiles = Get-Content -Path $manifestPath -Raw | ConvertFrom-Json } catch { $generatedFiles = @() }
    }

    function Open-NewFileWriter {
        param([int]$Index)
        try {
            if ($script:fileWriter) { try { $script:fileWriter.Flush(); $script:fileWriter.Close(); $script:fileStream.Close() } catch { } }
        } catch { }

        $destPath = "$filesBaseNoExt-part$Index.csv"
        if ($UseLocalTemp) {
            $localFileName = ([System.IO.Path]::GetFileName($destPath)) + ".writing"
            $localPath = [System.IO.Path]::Combine($LocalTempDir, $localFileName)
        } else {
            $localPath = $destPath
        }

        try {
            # If local file exists and we're resuming, append rather than truncate
            if (Test-Path $localPath) {
                $fileMode = [System.IO.FileMode]::Append
            } else {
                $fileMode = [System.IO.FileMode]::Create
            }
            $fs = New-Object System.IO.FileStream($localPath, $fileMode, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None, 65536, [System.IO.FileOptions]::SequentialScan)
            $script:fileStream = $fs
            $script:fileWriter = New-Object System.IO.StreamWriter($fs, [System.Text.Encoding]::UTF8, 65536)
            $script:fileWriter.AutoFlush = $false
            $script:currentLocalPath = $localPath
            $script:currentDestPath = $destPath
            $script:fileWriter.WriteLine('"Name","Modified","Path"')
        } catch {
            Write-Host (("Warning: could not create files CSV writer for {0}: {1}" -f $localPath, $_)) -ForegroundColor Yellow
            $script:fileWriter = $null
            $script:fileStream = $null
            $script:currentLocalPath = $null
            $script:currentDestPath = $null
        }
        return $destPath
    }

    function Finalize-Part {
        param(
            [string]$localPath,
            [string]$destPath
        )
        try {
            if ($script:fileWriter) { $script:fileWriter.Flush() }
            if ($script:fileStream) {
                try { $script:fileStream.Flush($true) } catch { try { $script:fileStream.Flush() } catch { } }
            }
        } catch { }

        try { if ($script:fileWriter) { $script:fileWriter.Close() } } catch { }
        try { if ($script:fileStream) { $script:fileStream.Close() } } catch { }

        if ($UseLocalTemp -and $localPath -and (Test-Path $localPath)) {
            $destWriting = $destPath + '.writing'
            if ($CopyInBackground) {
                Start-Job -ScriptBlock {
                    param($src,$dest,$destWriting)
                    try {
                        Copy-Item -Path $src -Destination $destWriting -Force
                        Move-Item -Path $destWriting -Destination $dest -Force
                        Remove-Item -Path $src -Force
                    } catch {
                        Write-Output "Copy job failed for $src -> $dest : $($_)"
                    }
                } -ArgumentList $localPath,$destPath,$destWriting | Out-Null
            } else {
                try {
                    Copy-Item -Path $localPath -Destination $destWriting -Force
                    Move-Item -Path $destWriting -Destination $destPath -Force
                    Remove-Item -Path $localPath -Force
                } catch {
                    Write-Host (("Failed copying part {0} to network: {1}" -f $localPath, $_)) -ForegroundColor Yellow
                }
            }
        }

        # null out script writer/stream pointers
        $script:fileWriter = $null
        $script:fileStream = $null
        $script:currentLocalPath = $null
        $script:currentDestPath = $null
    }

    $currentFilesPath = Open-NewFileWriter -Index $fileIndex

    foreach ($path in $fileEnumerable) {
        $processed++
        if ($ShowProgressBar -and $totalEstimated -gt 0) {
            $percent = [int](($processed / $totalEstimated) * 100)
            $barLen = 20
            $filled = [int]([math]::Floor($percent/100 * $barLen))
            $bar = '[' + (Repeat-Char '#' $filled) + (Repeat-Char '-' ($barLen - $filled)) + "] $percent% ($processed/$totalEstimated)"
            Write-Host -NoNewline ("`r$bar")
        } elseif ($ShowProgressBar -and ($processed % 1000 -eq 0)) {
            Write-Host "Files seen: $processed" -ForegroundColor DarkGray
        }

        try {
            $f = New-Object System.IO.FileInfo($path)
        } catch {
            # Skip files we cannot access or construct FileInfo for
            continue
        }

        $full = $f.FullName

        # If resuming from checkpoint, skip until we find the last processed path
        if ($resuming) {
            if ($null -ne $resumeLastPath) {
                if ($full -eq $resumeLastPath) {
                    # found the last processed file; stop skipping and continue after it
                    $resuming = $false
                    continue
                } else {
                    continue
                }
            } else {
                $resuming = $false
            }
        }
        $folderFull = Split-Path $full -Parent
        if ($EnableFolderReport) {
            if (-not $folderStats.ContainsKey($folderFull)) {
                $folderStats[$folderFull] = [PSCustomObject]@{
                        Library = 'Local'
                        FolderName = Split-Path $folderFull -Leaf
                        FileCount = 0
                        ItemCount = 0
                        Depth = 0
                        SizeBytes = 0
                        SizeMB = 0
                        SizeGB = 0
                        Path = $folderFull
                    }
            }

            $folderObj = $folderStats[$folderFull]
            $folderObj.FileCount = $folderObj.FileCount + 1
            $folderObj.SizeBytes = $folderObj.SizeBytes + $f.Length
            $folderObj.SizeMB = [math]::Round($folderObj.SizeBytes / 1MB, 4)
            $folderObj.SizeGB = [math]::Round($folderObj.SizeBytes / 1GB, 6)
            $folderStats[$folderFull] = $folderObj
        }

        $TotalFiles++
        $TotalSize += $f.Length

        if ($AgeCutoff) {
            try {
                if ($f.LastWriteTime -le $AgeCutoff) {
                    $OldFilesCount++
                    $OldFilesSize += $f.Length
                }
            } catch {
                # ignore any file timestamp access errors
            }
        }

        # File-wise report: include file if no age filter or if it meets the cutoff
        $includeFile = $true
        if ($AgeCutoff) { $includeFile = ($f.LastWriteTime -le $AgeCutoff) }
        if ($includeFile) {
            if ($fileWriter) {
                $formatted = "{0},{1},{2}" -f (EscapeCSV($f.Name)), (EscapeCSV($f.LastWriteTime.ToString('o'))), (EscapeCSV($full))
                $fileWriter.WriteLine($formatted)
                $fileRowCount++
                $FileReportCount++

                if ($FlushEvery -gt 0 -and ($fileRowCount % $FlushEvery -eq 0)) {
                    try { $script:fileWriter.Flush() } catch { }
                    try { if ($script:fileStream) { try { $script:fileStream.Flush($true) } catch { $script:fileStream.Flush() } } } catch { }
                }

                if ($CheckpointEvery -gt 0 -and ($FileReportCount % $CheckpointEvery -eq 0)) {
                    try {
                        $ckObj = @{ LastProcessedPath = $full; CurrentPartIndex = $fileIndex; FileRowCount = $fileRowCount; Timestamp = (Get-Date).ToString('o') }
                        $ckObj | ConvertTo-Json -Depth 5 | Out-File -FilePath $checkpointPath -Encoding UTF8
                    } catch { }
                }

                if ($fileRowCount -ge $MaxRowsPerFile) {
                    # finalize current part (flush/close and copy to dest)
                    Finalize-Part -localPath $script:currentLocalPath -destPath $script:currentDestPath
                    $generatedFiles += [PSCustomObject]@{ Path = $script:currentDestPath; Rows = $fileRowCount }
                    try { $generatedFiles | ConvertTo-Json -Depth 5 | Out-File -FilePath $manifestPath -Encoding UTF8 } catch { }
                    $fileIndex++
                    $currentFilesPath = Open-NewFileWriter -Index $fileIndex
                    $fileRowCount = 0
                }
            } else {
                # fallback: append via Export-Csv to the original single-file path
                $fileRow = [PSCustomObject]@{
                    Name = $f.Name
                    Modified = $f.LastWriteTime
                    Path = $full
                }
                $fileRow | Export-Csv -Path $filesPath -Append -NoTypeInformation -Encoding UTF8
                $FileReportCount++
            }
        }

        # Periodic user update: give a concise status message so the user knows the system is working
        if ($processed % 5000 -eq 0) {
            $sizeMB = [math]::Round($TotalSize / 1MB, 2)
            Write-Host "Scanning in progress: $processed files processed, approx size: $sizeMB MB" -ForegroundColor Cyan
        }
    }
} catch {
    Write-Host "ERROR enumerating files: $_" -ForegroundColor Red
    exit 1
}
if ($EnableFolderReport) {
    # Compute direct ItemCount and folder depths for each folder (done after scan to avoid per-file overhead)
    $MaxDepth = 0
    $folderKeys = $folderStats.Keys
    $folderTotal = $folderKeys.Count
    $folderProcessed = 0
    Write-Host "Computing folder item counts and depths..." -ForegroundColor Cyan
    foreach ($k in $folderKeys) {
    # depth relative to source path
    try {
        $normSource = $SourcePath.TrimEnd('\','/')
        if ([string]::IsNullOrEmpty($normSource)) {
            $segments = $k -split '[\\/]'
            $depth = [math]::Max(0, $segments.Length - 1)
        } elseif ($k.StartsWith($normSource, [System.StringComparison]::OrdinalIgnoreCase)) {
            $relFolder = $k.Substring($normSource.Length).TrimStart('\','/')
            if ($relFolder -ne '') { $segments = $relFolder -split '[\\/]'; $depth = $segments.Length } else { $depth = 0 }
        } else {
            $segments = $k -split '[\\/]'
            $depth = [math]::Max(0, $segments.Length - 1)
        }
        $folderStats[$k].Depth = $depth
        if ($depth -gt $MaxDepth) { $MaxDepth = $depth }
    } catch {
        $folderStats[$k].Depth = 0
    }

    try {
        $children = Get-ChildItem -Path $k -Force -ErrorAction SilentlyContinue
        $folderStats[$k].ItemCount = ($children | Measure-Object).Count
    } catch {
        $folderStats[$k].ItemCount = 0
    }

    $folderProcessed++
    if ($folderProcessed % 100 -eq 0 -or $folderProcessed -eq $folderTotal) {
        $pct = [int]((100 * $folderProcessed) / [math]::Max(1,$folderTotal))
        Write-Progress -Activity 'Processing folders' -Status "Folders processed: $folderProcessed of $folderTotal" -PercentComplete $pct
    }
    }

    $TotalFolders = ($folderStats.Keys).Count
} else {
    # Folder processing disabled
    $TotalFolders = 0
}

Write-Host "Local scan complete. Folders: $TotalFolders | Files: $TotalFiles | Size: $([math]::Round($TotalSize/1MB,4)) MB" -ForegroundColor Green
Write-Host "Maximum folder depth found: $MaxDepth" -ForegroundColor Green

# Export summary and folders
try {
    $summary = @()
    $summary += [PSCustomObject]@{ Metric = 'Total Folders'; Value = $TotalFolders }
    $summary += [PSCustomObject]@{ Metric = 'Total Files'; Value = $TotalFiles }
    $summary += [PSCustomObject]@{ Metric = 'Total Size (MB)'; Value = [math]::Round($TotalSize / 1MB, 4) }
    $summary += [PSCustomObject]@{ Metric = 'Total Size (GB)'; Value = [math]::Round($TotalSize / 1GB, 4) }
    $summary += [PSCustomObject]@{ Metric = 'Max Depth'; Value = $MaxDepth }

    if ($AgeCutoff) {
        $summary += [PSCustomObject]@{ Metric = "Files older than $AgeYears years (<= $($AgeCutoff.ToString('yyyy-MM-dd')))"; Value = $OldFilesCount }
        $summary += [PSCustomObject]@{ Metric = "Size of files older than $AgeYears years (MB)"; Value = [math]::Round($OldFilesSize / 1MB, 4) }
        $summary += [PSCustomObject]@{ Metric = "Size of files older than $AgeYears years (GB)"; Value = [math]::Round($OldFilesSize / 1GB, 4) }
    }
    # Include file-wise report count in summary
    $summary += [PSCustomObject]@{ Metric = 'File Report Rows'; Value = $FileReportCount }

    $summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8

    # Folder CSV export temporarily disabled (commented out).
    # $folderList = $folderStats.Values | Select-Object Library, FolderName, FileCount, ItemCount, Depth, SizeMB, SizeGB, Path
    # Write-Host "Writing folders CSV..." -ForegroundColor Cyan
    # # write header
    # $folderList | Select-Object -First 0 | Export-Csv -Path $foldersPath -NoTypeInformation -Encoding UTF8
    # $totalFoldersToWrite = $folderList.Count
    # $written = 0
    # foreach ($row in $folderList) {
    #     $row | Export-Csv -Path $foldersPath -Append -NoTypeInformation -Encoding UTF8
    #     $written++
    #     if ($written % 100 -eq 0 -or $written -eq $totalFoldersToWrite) {
    #         $pct = [int](100 * $written / [math]::Max(1,$totalFoldersToWrite))
    #         Write-Progress -Activity 'Writing folders CSV' -Status "Written $written/$totalFoldersToWrite" -PercentComplete $pct
    #     }
    # }
    # Write-Progress -Activity 'Writing folders CSV' -Completed

    # Close file writer if opened and record final part
    if ($script:fileWriter -or $script:fileStream) {
        # finalize last open part
        Finalize-Part -localPath $script:currentLocalPath -destPath $script:currentDestPath
        $generatedFiles += [PSCustomObject]@{ Path = $script:currentDestPath; Rows = $fileRowCount }
        try { $generatedFiles | ConvertTo-Json -Depth 5 | Out-File -FilePath $manifestPath -Encoding UTF8 } catch { }
    }

    # If we never used split writer (fallback), ensure original path appears
    if (-not $generatedFiles -or $generatedFiles.Count -eq 0) {
        $generatedFiles += [PSCustomObject]@{ Path = $filesPath; Rows = $FileReportCount }
        try { $generatedFiles | ConvertTo-Json -Depth 5 | Out-File -FilePath $manifestPath -Encoding UTF8 } catch { }
    }

    $fileListString = ($generatedFiles | ForEach-Object { $_.Path + ' (' + $_.Rows + ' rows)' }) -join "`n  "

    Write-Host "`n[SUCCESS] CSV files created successfully!`n  Summary: $summaryPath`n  Files:`n  $fileListString" -ForegroundColor Green
    # remove checkpoint on successful completion
    try { if (Test-Path $checkpointPath) { Remove-Item -Path $checkpointPath -Force } } catch { }
} catch {
    Write-Host "ERROR: Failed to export CSV files: $_" -ForegroundColor Red
}
