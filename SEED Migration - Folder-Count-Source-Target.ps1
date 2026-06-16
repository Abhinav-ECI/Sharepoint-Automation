# ============================================================
# Folder Count Compare: Source (Local/Network) vs Target (SharePoint)
# Exports a single CSV with per-folder counts
# ============================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SourcePath = "",
    [Parameter(Mandatory=$false)]
    [string]$SharePointInput = "",
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = ""
)

# Auto-switch to Windows PowerShell if needed (PnP classic module compatibility)
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

function Normalize-RelPath {
    param([string]$PathValue)
    if (-not $PathValue) { return "" }
    $s = $PathValue -replace '\\', '/'
    return $s.Trim('/').ToLower()
}

function Get-SharePointPathFromHyperlink {
    param([string]$Url)

    try {
        $uri = [uri]$Url
        if ($uri.AbsolutePath -like "*/Forms/*" -or $uri.Query -match "id=") {
            if ($uri.Query -match "id=([^&]+)") {
                $idParam = $matches[1]
                $decodedPath = [System.Uri]::UnescapeDataString($idParam)
                $cleanPath = $decodedPath.TrimEnd('/')
                return "https://$($uri.Host)$cleanPath"
            }
        }
        return $Url
    } catch {
        return $Url
    }
}

function Get-UrlFromInput {
    param([string]$InputValue)

    if ([string]::IsNullOrWhiteSpace($InputValue)) { return $InputValue }

    $v = $InputValue.Trim('"').Trim("'").Trim()

    # Support Markdown links like:
    # [Title](https://tenant.sharepoint.com/sites/...)
    if ($v -match '^\[[^\]]+\]\((https?://[^\)]+)\)$') {
        return $matches[1]
    }

    return $v
}

function Resolve-SharePointInput {
    param(
        [string]$InputValue,
        [string]$ProvidedSiteUrl
    )

    $resolved = [PSCustomObject]@{
        SiteUrl = $null
        TargetServerRelativeUrl = $null
    }

    if (-not $InputValue) { return $resolved }

    $inputTrimmed = Get-UrlFromInput -InputValue $InputValue

    if ($inputTrimmed -like "*/Forms/*" -or $inputTrimmed -match "id=") {
        Write-Host "SharePoint hyperlink detected. Extracting folder path..." -ForegroundColor Yellow
        $inputTrimmed = Get-SharePointPathFromHyperlink -Url $inputTrimmed
    }

    # Case 1: Full URL provided
    if ($inputTrimmed -match '^(https?://)') {
        $uri = [uri]$inputTrimmed
        $absPath = [System.Uri]::UnescapeDataString($uri.AbsolutePath).TrimEnd('/')
        $parts = $absPath -split '/' | Where-Object { $_ }

        $siteIndex = -1
        for ($i = 0; $i -lt $parts.Length; $i++) {
            if ($parts[$i] -eq 'sites' -or $parts[$i] -eq 'teams') {
                $siteIndex = $i
                break
            }
        }

        if ($siteIndex -ge 0 -and $siteIndex + 1 -lt $parts.Length) {
            $siteParts = $parts[0..($siteIndex + 1)]
            $resolved.SiteUrl = "$($uri.Scheme)://$($uri.Host)/$($siteParts -join '/')"

            if ($siteIndex + 2 -lt $parts.Length) {
                $targetParts = $parts[($siteIndex + 2)..($parts.Length - 1)]
                $resolved.TargetServerRelativeUrl = "/$($siteParts -join '/')/$($targetParts -join '/')"
            }
        } else {
            $resolved.SiteUrl = "$($uri.Scheme)://$($uri.Host)$absPath"
        }

        return $resolved
    }

    # Case 2: Server-relative SharePoint path provided
    if ($inputTrimmed.StartsWith('/')) {
        $resolved.TargetServerRelativeUrl = $inputTrimmed.TrimEnd('/')
        if ($ProvidedSiteUrl) {
            $resolved.SiteUrl = $ProvidedSiteUrl.Trim('"').Trim("'").Trim().TrimEnd('/')
        }
        return $resolved
    }

    return $resolved
}

# ============================================================
# STEP 1 - Get Source Path
# ============================================================
if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Select local source folder"
        $folderDialog.ShowNewFolderButton = $true

        $dialogResult = $null
        $ownerForm = New-Object System.Windows.Forms.Form
        $ownerForm.TopMost = $true
        $ownerForm.ShowInTaskbar = $false
        $ownerForm.FormBorderStyle = 'None'
        $ownerForm.StartPosition = 'Manual'
        $ownerForm.Size = New-Object System.Drawing.Size(0,0)
        $ownerForm.Location = New-Object System.Drawing.Point(-2000,-2000)
        $ownerForm.Opacity = 0
        $ownerForm.Show()
        try {
            $dialogResult = $folderDialog.ShowDialog($ownerForm)
        } finally {
            try { $ownerForm.Close(); $ownerForm.Dispose() } catch {}
        }

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $SourcePath = $folderDialog.SelectedPath
        } else {
            $SourcePath = Read-Host "Enter local source folder path"
        }
    } catch {
        $SourcePath = Read-Host "Enter local source folder path"
    }
}

$SourcePath = $SourcePath.Trim('"').Trim("'").Trim()
if (-not (Test-Path $SourcePath)) {
    Write-Host "Source path not found: $SourcePath" -ForegroundColor Red
    exit 1
}

Write-Host "Source folder: $SourcePath" -ForegroundColor Green

# ============================================================
# STEP 2 - Get SharePoint input
# ============================================================
if ([string]::IsNullOrWhiteSpace($SharePointInput)) {
    $SharePointInput = Read-Host "Enter SharePoint folder URL or server-relative path"
}

$resolvedInput = Resolve-SharePointInput -InputValue $SharePointInput -ProvidedSiteUrl $SiteUrl
$SiteUrl = $resolvedInput.SiteUrl
$TargetServerRelativeUrl = $resolvedInput.TargetServerRelativeUrl

if (-not $SiteUrl) {
    $SiteUrl = Read-Host "Enter SharePoint site URL (required when providing only server-relative path)"
    $SiteUrl = $SiteUrl.Trim('"').Trim("'").Trim().TrimEnd('/')
}

if (-not $TargetServerRelativeUrl) {
    $TargetServerRelativeUrl = Read-Host "Enter SharePoint target folder server-relative path (example: /sites/YourSite/Shared Documents/Folder)"
    $TargetServerRelativeUrl = $TargetServerRelativeUrl.Trim('"').Trim("'").Trim().TrimEnd('/')
}

if (-not $SiteUrl -or -not $TargetServerRelativeUrl) {
    Write-Host "ERROR: Site URL and target folder path are required." -ForegroundColor Red
    exit 1
}

Write-Host "Site URL: $SiteUrl" -ForegroundColor Green
Write-Host "Target folder: $TargetServerRelativeUrl" -ForegroundColor Green

# ============================================================
# STEP 3 - Connect to SharePoint
# ============================================================
$ModuleExists = Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline
if (-not $ModuleExists) {
    Write-Host "ERROR: Module not found. Install with: Install-Module SharePointPnPPowerShellOnline -Force" -ForegroundColor Red
    exit 1
}

try {
    Import-Module SharePointPnPPowerShellOnline -ErrorAction Stop -WarningAction SilentlyContinue
} catch {
    Write-Host "ERROR: Failed to import module: $_" -ForegroundColor Red
    exit 1
}

try {
    $Conn = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ReturnConnection -WarningAction SilentlyContinue -ErrorAction Stop
    Write-Host "Connected to SharePoint." -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to connect: $_" -ForegroundColor Red
    exit 1
}

try {
    $null = Get-PnPFolder -Url $TargetServerRelativeUrl -Connection $Conn -ErrorAction Stop
} catch {
    Write-Host "ERROR: Target folder not found: $TargetServerRelativeUrl" -ForegroundColor Red
    exit 1
}

# ============================================================
# STEP 4 - Build source folder stats
# ============================================================
Write-Host "Building source folder statistics..." -ForegroundColor Cyan

$sourceStats = @{}

function Ensure-SourceFolderStat {
    param(
        [string]$relKey,
        [string]$folderName,
        [string]$fullPath
    )

    if (-not $sourceStats.ContainsKey($relKey)) {
        $sourceStats[$relKey] = [PSCustomObject]@{
            FolderName = $folderName
            RelativePath = $relKey
            SourcePath = $fullPath
            SourceFileCount = 0
            SourceItemCount = 0
            SourceSizeBytes = [int64]0
        }
    }
}

$sourceRootName = Split-Path -Path $SourcePath -Leaf
$sourceRootKey = ""
Ensure-SourceFolderStat -relKey $sourceRootKey -folderName $sourceRootName -fullPath $SourcePath

$sourceFolders = Get-ChildItem -Path $SourcePath -Recurse -Directory -Force -ErrorAction Stop
foreach ($folder in $sourceFolders) {
    $rel = $folder.FullName.Substring($SourcePath.Length).TrimStart('\','/')
    $key = Normalize-RelPath $rel
    Ensure-SourceFolderStat -relKey $key -folderName $folder.Name -fullPath $folder.FullName

    $parent = Split-Path -Path $key -Parent
    if ($parent -eq '.' -or $parent -eq '') { $parent = "" }
    if ($sourceStats.ContainsKey($parent)) {
        $sourceStats[$parent].SourceItemCount++
    }
}

$sourceFiles = Get-ChildItem -Path $SourcePath -Recurse -File -Force -ErrorAction Stop
foreach ($file in $sourceFiles) {
    $rel = $file.FullName.Substring($SourcePath.Length).TrimStart('\','/') -replace '\\','/'
    $parentRel = Normalize-RelPath (Split-Path -Path $rel -Parent)
    if ($parentRel -eq '.') { $parentRel = "" }

    if (-not $sourceStats.ContainsKey($parentRel)) {
        $folderName = if ($parentRel) { Split-Path -Path $parentRel -Leaf } else { $sourceRootName }
        $folderPath = if ($parentRel) { Join-Path $SourcePath ($parentRel -replace '/','\\') } else { $SourcePath }
        Ensure-SourceFolderStat -relKey $parentRel -folderName $folderName -fullPath $folderPath
    }

    $sourceStats[$parentRel].SourceFileCount++
    $sourceStats[$parentRel].SourceItemCount++
    $sourceStats[$parentRel].SourceSizeBytes += [int64]$file.Length
}

$sourceTotalFiles = @($sourceFiles).Count
$sourceTotalFolders = @($sourceFolders).Count
$sourceTotalItems = $sourceTotalFiles + $sourceTotalFolders
$sourceTotalSizeBytes = [int64]0
if ($sourceTotalFiles -gt 0) {
    $sum = ($sourceFiles | Measure-Object -Property Length -Sum).Sum
    if ($sum) { $sourceTotalSizeBytes = [int64]$sum }
}

# ============================================================
# STEP 5 - Build target folder stats
# ============================================================
Write-Host "Building target folder statistics..." -ForegroundColor Cyan

$targetStats = @{}

function Ensure-TargetFolderStat {
    param(
        [string]$relKey,
        [string]$folderName,
        [string]$targetPath
    )

    if (-not $targetStats.ContainsKey($relKey)) {
        $targetStats[$relKey] = [PSCustomObject]@{
            FolderName = $folderName
            RelativePath = $relKey
            TargetPath = $targetPath
            TargetFileCount = 0
            TargetItemCount = 0
            TargetSizeBytes = [int64]0
        }
    }
}

$targetRootName = Split-Path -Path $TargetServerRelativeUrl -Leaf
Ensure-TargetFolderStat -relKey "" -folderName $targetRootName -targetPath $TargetServerRelativeUrl

$targetTotalFiles = 0
$targetTotalFolders = 0
$targetTotalItems = 0
$targetTotalSizeBytes = [int64]0

$Lists = Get-PnPList -Connection $Conn -Includes RootFolder | Where-Object { $_.BaseTemplate -eq "101" -or $_.BaseTemplate -eq "109" }
$TargetList = $null
foreach ($List in $Lists) {
    if ($TargetServerRelativeUrl.StartsWith($List.RootFolder.ServerRelativeUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
        $TargetList = $List
        break
    }
}
if (-not $TargetList) {
    Write-Host "ERROR: Could not find document library for target path." -ForegroundColor Red
    exit 1
}

$targetQuery = @{
    List = $TargetList.Id
    Fields = @('FileRef','FileLeafRef','File_x0020_Size','ContentTypeId','FSObjType')
    PageSize = 5000
    Connection = $Conn
    FolderServerRelativeUrl = $TargetServerRelativeUrl
}

Get-PnPListItem @targetQuery | ForEach-Object {
    $item = $_
    $fileRef = $item['FileRef']
    if (-not $fileRef) { return }

    if (-not ($fileRef -like "$TargetServerRelativeUrl/*" -or $fileRef -eq $TargetServerRelativeUrl)) { return }

    $isFolder = $false
    try {
        if ($item['ContentTypeId'] -and $item['ContentTypeId'].StringValue -like '0x0120*') { $isFolder = $true }
    } catch {}
    if (-not $isFolder) {
        try { if ([int]$item['FSObjType'] -eq 1) { $isFolder = $true } } catch {}
    }

    $relFromTarget = $fileRef.Substring($TargetServerRelativeUrl.Length).TrimStart('/')
    $relNorm = Normalize-RelPath $relFromTarget

    if ($isFolder) {
        # Skip counting root as child of itself
        if ($fileRef -eq $TargetServerRelativeUrl) { return }

        $targetTotalFolders++
        $targetTotalItems++

        $folderName = Split-Path -Path $fileRef -Leaf
        Ensure-TargetFolderStat -relKey $relNorm -folderName $folderName -targetPath $fileRef

        $parentRel = Normalize-RelPath (Split-Path -Path $relNorm -Parent)
        if ($parentRel -eq '.') { $parentRel = "" }
        if (-not $targetStats.ContainsKey($parentRel)) {
            $parentName = if ($parentRel) { Split-Path -Path $parentRel -Leaf } else { $targetRootName }
            $parentPath = if ($parentRel) { "$TargetServerRelativeUrl/$parentRel" } else { $TargetServerRelativeUrl }
            Ensure-TargetFolderStat -relKey $parentRel -folderName $parentName -targetPath $parentPath
        }
        $targetStats[$parentRel].TargetItemCount++
    } else {
        $targetTotalFiles++
        $targetTotalItems++

        $parentRel = Normalize-RelPath (Split-Path -Path $relNorm -Parent)
        if ($parentRel -eq '.') { $parentRel = "" }

        if (-not $targetStats.ContainsKey($parentRel)) {
            $folderName = if ($parentRel) { Split-Path -Path $parentRel -Leaf } else { $targetRootName }
            $folderPath = if ($parentRel) { "$TargetServerRelativeUrl/$parentRel" } else { $TargetServerRelativeUrl }
            Ensure-TargetFolderStat -relKey $parentRel -folderName $folderName -targetPath $folderPath
        }

        $targetStats[$parentRel].TargetFileCount++
        $targetStats[$parentRel].TargetItemCount++

        $fileSize = [int64]0
        if ($item['File_x0020_Size']) {
            try { $fileSize = [int64]$item['File_x0020_Size'] } catch { $fileSize = [int64]0 }
        }
        $targetStats[$parentRel].TargetSizeBytes += $fileSize
        $targetTotalSizeBytes += $fileSize
    }
}

# ============================================================
# STEP 6 - Merge and export CSV
# ============================================================
Write-Host "Merging source and target folder stats..." -ForegroundColor Cyan

$allKeys = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($k in $sourceStats.Keys) { [void]$allKeys.Add($k) }
foreach ($k in $targetStats.Keys) { [void]$allKeys.Add($k) }

$rows = @()

# First row: totals for selected root folder input
$rows += [PSCustomObject]@{
    'Folder Name' = $sourceRootName
    'Source File Count' = $sourceTotalFiles
    'Target File Count' = $targetTotalFiles
    'Source Item Count' = $sourceTotalItems
    'Target Item Count' = $targetTotalItems
    'Source Size Bytes' = $sourceTotalSizeBytes
    'Target Size Bytes' = $targetTotalSizeBytes
    'Path' = $SourcePath
}

foreach ($k in (($allKeys | Where-Object { $_ -ne "" }) | Sort-Object)) {
    $src = $null
    $tgt = $null
    if ($sourceStats.ContainsKey($k)) { $src = $sourceStats[$k] }
    if ($targetStats.ContainsKey($k)) { $tgt = $targetStats[$k] }

    $folderName = $null
    if ($src -and $src.FolderName) { $folderName = $src.FolderName }
    elseif ($tgt -and $tgt.FolderName) { $folderName = $tgt.FolderName }
    else { $folderName = "(root)" }

    $pathValue = $k

    $rows += [PSCustomObject]@{
        'Folder Name' = $folderName
        'Source File Count' = if ($src) { $src.SourceFileCount } else { 0 }
        'Target File Count' = if ($tgt) { $tgt.TargetFileCount } else { 0 }
        'Source Item Count' = if ($src) { $src.SourceItemCount } else { 0 }
        'Target Item Count' = if ($tgt) { $tgt.TargetItemCount } else { 0 }
        'Source Size Bytes' = if ($src) { [int64]$src.SourceSizeBytes } else { [int64]0 }
        'Target Size Bytes' = if ($tgt) { [int64]$tgt.TargetSizeBytes } else { [int64]0 }
        'Path' = $pathValue
    }
}

$desktop = [Environment]::GetFolderPath('Desktop')
$ExportPath = Join-Path $desktop ("Folder_Count_Compare_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

try {
    $rows | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "SUCCESS: Folder comparison CSV created." -ForegroundColor Green
    Write-Host "Location: $ExportPath" -ForegroundColor Green
    Write-Host "Total folders in report: $($rows.Count)" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to export CSV: $_" -ForegroundColor Red
    exit 1
}

Write-Host "Done." -ForegroundColor Green
