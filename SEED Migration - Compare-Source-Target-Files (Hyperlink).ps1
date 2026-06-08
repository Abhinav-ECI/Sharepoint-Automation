# ============================================================
# Compare Source (Local) Files to Target (SharePoint) Files
# ============================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SourcePath = "",
    [Parameter(Mandatory=$false)]
    [string]$SharePointUrl = ""
)

# Auto-switch to Windows PowerShell if needed
$PSVersion = $PSVersionTable.PSVersion.Major

if ($PSVersion -gt 5) {
    Write-Host "PowerShell 7+ detected. Switching to Windows PowerShell..." -ForegroundColor Yellow
    $ArgumentList = @()
    foreach ($Key in $PSBoundParameters.Keys) {
        $ArgumentList += "-$Key"
        $ArgumentList += "`"$($PSBoundParameters[$Key])`""
    }
    & powershell.exe -NoProfile -File $PSCommandPath $ArgumentList
    exit
}

Write-Host "PowerShell version: $PSVersion (Windows PowerShell)" -ForegroundColor Green

# ============================================================
# HELPER - Extract path from SharePoint hyperlink
# ============================================================
function Get-SharePointPathFromHyperlink {
    param([string]$Url)
    
    try {
        $uri = [uri]$Url
        
        # Check if this is a SharePoint Forms/view URL
        if ($uri.AbsolutePath -like "*/Forms/*" -or $uri.Query -match "id=") {
            # Try to extract from 'id' query parameter
            if ($uri.Query -match "id=([^&]+)") {
                $idParam = $matches[1]
                $decodedPath = [System.Uri]::UnescapeDataString($idParam)
                
                # The decoded path should be a server-relative path like /sites/DSG_QualityAssurance/Shared Documents/zz - QA Automation
                # Return as full URL
                $cleanPath = $decodedPath.TrimEnd('/')
                return "https://$($uri.Host)$cleanPath"
            }
        }
        
        # If not a forms URL, return original
        return $Url
    } catch {
        return $Url
    }
}

# STEP 1 - Get Source Path
# ============================================================
if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = Read-Host "Enter local source folder path"
}
$SourcePath = $SourcePath.Trim('"').Trim("'").Trim()

if (-not (Test-Path $SourcePath)) {
    Write-Host "Source path not found: $SourcePath" -ForegroundColor Red
    exit 1
}

Write-Host "Source folder: $SourcePath" -ForegroundColor Green

# ============================================================
# STEP 2 - Get SharePoint URL (with optional target folder)
# ============================================================
if ([string]::IsNullOrWhiteSpace($SharePointUrl)) {
    $SharePointUrl = Read-Host "Enter SharePoint site URL (or paste a SharePoint folder hyperlink)"
}
$SharePointUrl = $SharePointUrl.Trim('"').Trim("'").Trim()

# If user pasted a hyperlink, extract the actual path from it
if ($SharePointUrl -like "*/Forms/*" -or $SharePointUrl -match "id=") {
    Write-Host "SharePoint hyperlink detected. Extracting folder path..." -ForegroundColor Yellow
    $SharePointUrl = Get-SharePointPathFromHyperlink -Url $SharePointUrl
}

# Parse the input to extract site URL and target folder
$SiteUrl = $null
$TargetServerRelativeUrl = $null

if ($SharePointUrl -match '^(https?://)') {
    $uri = [uri]$SharePointUrl
    $absPath = [System.Uri]::UnescapeDataString($uri.AbsolutePath).TrimEnd('/')
    $parts = $absPath -split '/' | Where-Object { $_ }
    
    # Find /sites/ or /teams/ segment
    $siteIndex = -1
    for ($i = 0; $i -lt $parts.Length; $i++) {
        if ($parts[$i] -eq 'sites' -or $parts[$i] -eq 'teams') {
            $siteIndex = $i
            break
        }
    }
    
    if ($siteIndex -ge 0 -and $siteIndex + 1 -lt $parts.Length) {
        # Build site URL
        $siteParts = $parts[0..($siteIndex + 1)]
        $SiteUrl = "$($uri.Scheme)://$($uri.Host)/$($siteParts -join '/')"
        
        # If there are more path segments, those are the target folder
        if ($siteIndex + 2 -lt $parts.Length) {
            $targetParts = $parts[($siteIndex + 2)..($parts.Length - 1)]
            $TargetServerRelativeUrl = "/$($siteParts -join '/')/$($targetParts -join '/')"
        }
    } else {
        # No /sites/ structure found; treat whole URL as site
        $SiteUrl = "$($uri.Scheme)://$($uri.Host)$absPath"
    }
}

if (-not $SiteUrl) {
    Write-Host "ERROR: Could not parse SharePoint URL" -ForegroundColor Red
    exit 1
}

Write-Host "Site URL: $SiteUrl" -ForegroundColor Green
if ($TargetServerRelativeUrl) {
    Write-Host "Target folder: $TargetServerRelativeUrl" -ForegroundColor Green
} else {
    Write-Host "Target folder: (not specified, will scan entire site)" -ForegroundColor Yellow
}

# ============================================================
# STEP 3 - Connect to SharePoint
# ============================================================
Write-Host "Importing SharePointPnPPowerShellOnline module..." -ForegroundColor Cyan

$ModuleExists = Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline
if (-not $ModuleExists) {
    Write-Host "ERROR: Module not found. Install with: Install-Module SharePointPnPPowerShellOnline -Force" -ForegroundColor Red
    exit 1
}

try {
    Import-Module SharePointPnPPowerShellOnline -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Host "Module imported successfully." -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to import module: $_" -ForegroundColor Red
    exit 1
}

Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
try {
    $Conn = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop
    Write-Host "Connected OK" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to connect: $_" -ForegroundColor Red
    exit 1
}

# ============================================================
# STEP 4 - Verify Target Folder (if specified)
# ============================================================
if ($TargetServerRelativeUrl) {
    Write-Host "Verifying target folder..." -ForegroundColor Cyan
    try {
        $folderObj = Get-PnPFolder -Url $TargetServerRelativeUrl -Connection $Conn -ErrorAction Stop
        Write-Host "Target folder verified" -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Target folder not found: $TargetServerRelativeUrl" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "No target folder specified; will scan entire site" -ForegroundColor Yellow
}

# ============================================================
# STEP 5 - Scan Source Folder
# ============================================================
Write-Host ""
Write-Host "Scanning source folder..." -ForegroundColor Cyan

$SourceFiles = @()
try {
    $allSourceFiles = Get-ChildItem -Path $SourcePath -Recurse -File -Force -ErrorAction Stop
    
    foreach ($file in $allSourceFiles) {
        $relPath = $file.FullName.Substring($SourcePath.Length).TrimStart('\', '/')
        $relPath = $relPath -replace '\\', '/'
        
        $SourceFiles += [PSCustomObject]@{
            FileName = $file.Name
            RelativePath = $relPath
            FullPath = $file.FullName
            SizeBytes = $file.Length
            SizeMB = [math]::Round($file.Length / 1MB, 2)
            Modified = $file.LastWriteTime
            Exists_InTarget = $null
            TargetSize = $null
            SizeMatch = $null
            TargetModified = $null
            DateMatch = $null
        }
    }
    
    Write-Host "Found $($SourceFiles.Count) file(s) in source" -ForegroundColor Green
} catch {
    Write-Host "ERROR scanning source: $_" -ForegroundColor Red
    exit 1
}

# ============================================================
# STEP 6 - Scan Target and Match Files
# ============================================================
Write-Host ""
Write-Host "Scanning target folder and matching files..." -ForegroundColor Cyan

$MatchedCount = 0

try {
    $Lists = Get-PnPList -Connection $Conn -Includes RootFolder | Where-Object { $_.BaseTemplate -eq "101" -or $_.BaseTemplate -eq "109" }
    
    $TargetList = $null
    
    # If a target folder was specified, find the list containing it
    if ($TargetServerRelativeUrl) {
        foreach ($List in $Lists) {
            if ($TargetServerRelativeUrl.StartsWith($List.RootFolder.ServerRelativeUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
                $TargetList = $List
                break
            }
        }
    } else {
        # If no target folder, scan all lists (use the first one as a fallback, or all)
        if ($Lists.Count -gt 0) {
            $TargetList = $Lists[0]
        }
    }
    
    if (-not $TargetList) {
        Write-Host "ERROR: Could not find list" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Found target list: $($TargetList.Title)" -ForegroundColor Green
    
    # Get items from the target list
    $Items = $null
    if ($TargetServerRelativeUrl) {
        try {
            $Items = Get-PnPListItem -List $TargetList.Id -FolderServerRelativeUrl $TargetServerRelativeUrl -PageSize 5000 -Connection $Conn -ErrorAction Stop
        } catch {
            Write-Host "Using full library fetch..." -ForegroundColor Yellow
            $Items = Get-PnPListItem -List $TargetList.Id -PageSize 5000 -Connection $Conn
        }
    } else {
        $Items = Get-PnPListItem -List $TargetList.Id -PageSize 5000 -Connection $Conn
    }
    
    # Filter to only files (not folders)
    $TargetFileItems = $Items | Where-Object {
        $_["FileRef"] -and 
        $_["ContentTypeId"].StringValue -notlike "0x0120*"
    }
    
    # If target folder was specified, filter to only files under that folder
    if ($TargetServerRelativeUrl) {
        $TargetFileItems = $TargetFileItems | Where-Object {
            $_["FileRef"] -like "$TargetServerRelativeUrl/*" -or $_["FileRef"] -eq $TargetServerRelativeUrl
        }
    }
    
    Write-Host "Found $($TargetFileItems.Count) file(s) in target" -ForegroundColor Green
    
    # Build target file map
    $TargetFileMap = @{}
    foreach ($item in $TargetFileItems) {
        $fileRef = $item["FileRef"]
        
        # Extract relative path from target folder (or list root if no target)
        if ($TargetServerRelativeUrl) {
            $relFromTarget = $fileRef.Substring($TargetServerRelativeUrl.Length).TrimStart('/')
        } else {
            # Use relative to library root
            $relFromTarget = $fileRef.Substring($TargetList.RootFolder.ServerRelativeUrl.Length).TrimStart('/')
        }
        
        $relFromTarget = $relFromTarget -replace '\\', '/'
        
        if (-not $TargetFileMap.ContainsKey($relFromTarget)) {
            $TargetFileMap[$relFromTarget] = @{
                Item = $item
                Name = $item["FileLeafRef"]
                Size = [int64]($item["File_x0020_Size"] -as [int64])
                Modified = $item["Modified"] -as [datetime]
                Path = $fileRef
            }
        }
    }
    
    # Match source to target
    foreach ($srcFile in $SourceFiles) {
        if ($TargetFileMap.ContainsKey($srcFile.RelativePath)) {
            $targetInfo = $TargetFileMap[$srcFile.RelativePath]
            $srcFile.Exists_InTarget = $true
            $srcFile.TargetSize = $targetInfo.Size
            $srcFile.SizeMatch = ($srcFile.SizeBytes -eq $targetInfo.Size)
            $srcFile.TargetModified = $targetInfo.Modified
            
            if ($targetInfo.Modified) {
                $timeDiff = [Math]::Abs(($srcFile.Modified - $targetInfo.Modified).TotalSeconds)
                $srcFile.DateMatch = ($timeDiff -le 2)
            }
            
            $MatchedCount++
        } else {
            $srcFile.Exists_InTarget = $false
        }
    }
    
    Write-Host "Matched $MatchedCount file(s)" -ForegroundColor Green
    
} catch {
    Write-Host "ERROR scanning target: $_" -ForegroundColor Red
    exit 1
}

# ============================================================
# STEP 7 - Summary Report
# ============================================================
Write-Host ""
Write-Host "====== COMPARISON SUMMARY ======" -ForegroundColor Yellow

$FilesInSourceOnly = @($SourceFiles | Where-Object { $_.Exists_InTarget -eq $false })
$FilesMatched = @($SourceFiles | Where-Object { $_.Exists_InTarget -eq $true })
$FilesSizeMatch = @($FilesMatched | Where-Object { $_.SizeMatch -eq $true })
$FilesSizeMismatch = @($FilesMatched | Where-Object { $_.SizeMatch -eq $false })
$FilesDateMatch = @($FilesMatched | Where-Object { $_.DateMatch -eq $true })
$FilesDateMismatch = @($FilesMatched | Where-Object { $_.DateMatch -eq $false -or $null -eq $_.DateMatch })

Write-Host "Total source files:      $($SourceFiles.Count)" -ForegroundColor White
Write-Host "  Matched in target:     $($FilesMatched.Count)" -ForegroundColor Green
Write-Host "  Missing from target:   $($FilesInSourceOnly.Count)" -ForegroundColor Red

if ($FilesMatched.Count -gt 0) {
    Write-Host ""
    Write-Host "Matched file details:" -ForegroundColor Cyan
    Write-Host "  Size matches:        $($FilesSizeMatch.Count)" -ForegroundColor Green
    Write-Host "  Size differs:        $($FilesSizeMismatch.Count)" -ForegroundColor Yellow
    Write-Host "  Date matches:        $($FilesDateMatch.Count)" -ForegroundColor Green
    Write-Host "  Date differs:        $($FilesDateMismatch.Count)" -ForegroundColor Yellow
}

# ============================================================
# STEP 8 - Export CSV
# ============================================================
$desktop = [Environment]::GetFolderPath('Desktop')
$ExportPath = Join-Path $desktop ("Compare_Source_Target_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

Write-Host ""
Write-Host "Exporting to CSV..." -ForegroundColor Cyan

try {
    $SourceFiles | Select-Object `
        FileName,
        RelativePath,
        @{Name='SourceSize_Bytes'; Expression={$_.SizeBytes}},
        @{Name='SourceSize_MB'; Expression={$_.SizeMB}},
        @{Name='SourceModified'; Expression={$_.Modified}},
        @{Name='Exists_InTarget'; Expression={if ($_.Exists_InTarget) { 'YES' } else { 'NO' }}},
        @{Name='TargetSize_Bytes'; Expression={$_.TargetSize}},
        @{Name='SizeMatch'; Expression={if ($_.SizeMatch -eq $true) { 'YES' } elseif ($_.SizeMatch -eq $false) { 'NO' } else { 'N/A' }}},
        @{Name='TargetModified'; Expression={$_.TargetModified}},
        @{Name='DateMatch'; Expression={if ($_.DateMatch -eq $true) { 'YES' } elseif ($_.DateMatch -eq $false) { 'NO' } else { 'N/A' }}} |
        Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "SUCCESS: Comparison report created!" -ForegroundColor Green
    Write-Host "Location: $ExportPath" -ForegroundColor Green
    
} catch {
    Write-Host "ERROR: Failed to export CSV: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
