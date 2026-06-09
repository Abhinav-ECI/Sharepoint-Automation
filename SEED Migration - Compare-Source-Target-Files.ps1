# ============================================================
# Compare Source (Local) Files to Target (SharePoint) Files
# ============================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SourcePath = "",
    [Parameter(Mandatory=$false)]
    [string]$SharePointUrl = "",
    [Parameter(Mandatory=$false)]
    [int]$PartitionBatchSize = 50000,
    [Parameter(Mandatory=$false)]
    [int]$MaxCachedPartitions = 8,
    [Parameter(Mandatory=$false)]
    [int]$PartitionFlushSize = 2000,
    [Parameter(Mandatory=$false)]
    [switch]$KeepPartitionFiles
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

# Record script start time for elapsed-time reporting
$ScriptStartTime = Get-Date
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

# Helper: Compare Windows owner (domain\user) to SharePoint owner (display name/email)
function Test-OwnerMatch {
    param(
        [string]$SourceOwner,
        [string]$TargetOwner
    )

    if (-not $SourceOwner -or -not $TargetOwner) { return $false }

    $s = $SourceOwner.ToLower()
    $t = $TargetOwner.ToLower()

    # Extract username from domain\user (or keep full if no backslash)
    $winUser = ($s -split '\\')[-1]

    # Build tokens from target owner: remove punctuation, split on spaces, dots, underscores, hyphens
    $tClean = $t -replace '[(),]',' '
    $tokens = ($tClean -split '\s+|[._-]') | Where-Object { $_ -ne '' }

    # If target looks like an email, also include local-part tokens
    if ($t -match '@') {
        $local = ($t -split '@')[0]
        $tokens += ($local -split '[._-]') | Where-Object { $_ -ne '' }
    }

    # Exact matches
    if ($s -eq $t -or $winUser -eq $t) { return $true }

    # If target tokens contain both first and last and winUser contains both, it's a match
    $tokenMatches = 0
    foreach ($tok in $tokens) {
        if ($winUser -and $winUser -like "*$tok*") { $tokenMatches++ }
    }
    if ($tokens.Count -gt 1 -and $tokenMatches -ge 2) { return $true }

    # If any token equals the winUser or winUser contains a token, accept
    foreach ($tok in $tokens) {
        if ($tok -and ($winUser -eq $tok -or $winUser -like "*$tok*" -or $tok -like "*$winUser*")) { return $true }
    }

    return $false
}

# STEP 1 - Get Source Path
# ============================================================
function Normalize-RelPath {
    param([string]$p)
    if (-not $p) { return "" }
    $s = $p -replace '\\','/'
    $s = $s.Trim('/')
    return $s.ToLower()
}

function Get-PartitionId {
    param(
        [string]$key,
        [int]$numPartitions,
        $md5
    )
    if (-not $key) { return 0 }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($key)
    $hash = $md5.ComputeHash($bytes)
    $val = [System.BitConverter]::ToUInt32($hash,0)
    return [int]($val % [uint32]$numPartitions)
}

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    # Try to use a Windows Forms folder-picker dialog; fall back to Read-Host if unavailable.
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Select local source folder"
        $folderDialog.ShowNewFolderButton = $true

        # First try to show the dialog owned by the host main window (works in PowerShell ISE)
        $dialogResult = $null
        try {
            $mainHandle = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle
        } catch {
            $mainHandle = [IntPtr]::Zero
        }

        if ($mainHandle -and $mainHandle -ne [IntPtr]::Zero) {
            try {
                $code = @"
using System;
using System.Windows.Forms;
public class PSWindowWrapper : IWin32Window {
    private IntPtr _hwnd;
    public PSWindowWrapper(IntPtr handle) { _hwnd = handle; }
    public IntPtr Handle { get { return _hwnd; } }
}
"@
                Add-Type -TypeDefinition $code -Language CSharp -ErrorAction Stop
                $wrapper = New-Object PSWindowWrapper($mainHandle)
                $dialogResult = $folderDialog.ShowDialog($wrapper)
            } catch {
                # Fallback: invisible top-most owner form
                $ownerForm = New-Object System.Windows.Forms.Form
                $ownerForm.TopMost = $true
                $ownerForm.ShowInTaskbar = $false
                $ownerForm.FormBorderStyle = 'None'
                $ownerForm.StartPosition = 'Manual'
                $ownerForm.Size = New-Object System.Drawing.Size(0,0)
                $ownerForm.Location = New-Object System.Drawing.Point(-2000,-2000)
                $ownerForm.Opacity = 0
                $ownerForm.Show()
                try { $dialogResult = $folderDialog.ShowDialog($ownerForm) } finally { try { $ownerForm.Close(); $ownerForm.Dispose() } catch {} }
            }
        } else {
            # No main window handle; use invisible owner form
            $ownerForm = New-Object System.Windows.Forms.Form
            $ownerForm.TopMost = $true
            $ownerForm.ShowInTaskbar = $false
            $ownerForm.FormBorderStyle = 'None'
            $ownerForm.StartPosition = 'Manual'
            $ownerForm.Size = New-Object System.Drawing.Size(0,0)
            $ownerForm.Location = New-Object System.Drawing.Point(-2000,-2000)
            $ownerForm.Opacity = 0
            $ownerForm.Show()
            try { $dialogResult = $folderDialog.ShowDialog($ownerForm) } finally { try { $ownerForm.Close(); $ownerForm.Dispose() } catch {} }
        }

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $SourcePath = $folderDialog.SelectedPath
        } else {
            # User cancelled the dialog; fall back to prompt
            $SourcePath = Read-Host "Enter local source folder path"
        }
    } catch {
        # Non-GUI host or other error — fall back to text prompt
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
        $null = Get-PnPFolder -Url $TargetServerRelativeUrl -Connection $Conn -ErrorAction Stop
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
        
        # Get Windows file owner (domain\user or localmachine\user)
        $owner = $null
        try {
            $owner = (Get-Acl $file.FullName -ErrorAction Stop).Owner
        } catch {
            $owner = $null
        }

        $SourceFiles += [PSCustomObject]@{
            FileName = $file.Name
            Type = 'file'
            RelativePath = $relPath
            SourcePath = $file.FullName
            FullPath = $file.FullName
            SizeBytes = $file.Length
            SizeMB = [math]::Round($file.Length / 1MB, 2)
            Modified = $file.LastWriteTime
            SourceOwner = $owner
            Exists_InTarget = $null
            TargetSize = $null
            SizeMatch = $null
            SizeMatchStatus = $null
            TargetModified = $null
            DateMatch = $null
            TargetOwner = $null
            OwnerMatch = $null
            DestinationPath = $null
        }
    }
    
    Write-Host "Found $($SourceFiles.Count) file(s) in source" -ForegroundColor Green
} catch {
    Write-Host "ERROR scanning source: $_" -ForegroundColor Red
    exit 1
}

# ============================================================
# STEP 6 - Partitioned target index + match (disk-backed)
# ============================================================
Write-Host ""
Write-Host "Partitioning target items and matching..." -ForegroundColor Cyan

$MatchedCount = 0

try {
    $Lists = Get-PnPList -Connection $Conn -Includes RootFolder | Where-Object { $_.BaseTemplate -eq "101" -or $_.BaseTemplate -eq "109" }
    $TargetList = $null
    if ($TargetServerRelativeUrl) {
        foreach ($List in $Lists) {
            if ($TargetServerRelativeUrl.StartsWith($List.RootFolder.ServerRelativeUrl, [System.StringComparison]::OrdinalIgnoreCase)) { $TargetList = $List; break }
        }
    } else {
        if ($Lists.Count -gt 0) { $TargetList = $Lists[0] }
    }
    if (-not $TargetList) { Write-Host "ERROR: Could not find list" -ForegroundColor Red; exit 1 }
    Write-Host "Found target list: $($TargetList.Title)" -ForegroundColor Green

    $targetQuery = @{
        List = $TargetList.Id
        Fields = @('FileRef','FileLeafRef','File_x0020_Size','Modified','Author')
        PageSize = 5000
        Connection = $Conn
    }
    if ($TargetServerRelativeUrl) { $targetQuery['FolderServerRelativeUrl'] = $TargetServerRelativeUrl }

    # Pass 1: count target files (streamed)
    Write-Host "Counting target files..." -ForegroundColor Cyan
    $TotalTargetFiles = 0
    Get-PnPListItem @targetQuery | ForEach-Object {
        if ($_["FileRef"] -and $_["ContentTypeId"].StringValue -notlike '0x0120*') {
            if ($TargetServerRelativeUrl) {
                $fr = $_["FileRef"]
                if ($fr -like "$TargetServerRelativeUrl/*" -or $fr -eq $TargetServerRelativeUrl) { $TotalTargetFiles++ }
            } else { $TotalTargetFiles++ }
        }
    }
    Write-Host "Total target files: $TotalTargetFiles" -ForegroundColor Green

    $numPartitions = [int][Math]::Max(1, [Math]::Ceiling($TotalTargetFiles / $PartitionBatchSize))
    Write-Host "Using $numPartitions partition(s) (batch size $PartitionBatchSize)" -ForegroundColor Cyan

    $PartitionDir = Join-Path $env:TEMP ("SPCompare_Parts_{0}" -f ([System.Guid]::NewGuid().ToString()))
    New-Item -Path $PartitionDir -ItemType Directory -Force | Out-Null

    # Partition target items (pass 2) with buffered writes
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $partitionBuffers = @{}
    for ($i=0; $i -lt $numPartitions; $i++) { $partitionBuffers[$i] = @() }
    $flushSize = $PartitionFlushSize

    Write-Host "Partitioning target items to disk..." -ForegroundColor Cyan
    Get-PnPListItem @targetQuery | ForEach-Object {
        $item = $_
        if ($item["FileRef"] -and $item["ContentTypeId"].StringValue -notlike '0x0120*') {
            if ($TargetServerRelativeUrl) {
                $fref = $item["FileRef"]
                if (-not ($fref -like "$TargetServerRelativeUrl/*" -or $fref -eq $TargetServerRelativeUrl)) { return }
            }
            if ($TargetServerRelativeUrl) {
                $relFromTarget = $item["FileRef"].Substring($TargetServerRelativeUrl.Length).TrimStart('/')
            } else {
                $relFromTarget = $item["FileRef"].Substring($TargetList.RootFolder.ServerRelativeUrl.Length).TrimStart('/')
            }
            $relNorm = Normalize-RelPath $relFromTarget

            $ownerName = $null
            $author = $null
            try { $author = $item["Author"] } catch { $author = $null }
            if ($author) {
                try {
                    if ($author.PSObject.Properties.Name -contains 'LookupValue') { $ownerName = $author.LookupValue }
                    elseif ($author.PSObject.Properties.Name -contains 'Email') { $ownerName = $author.Email }
                    elseif ($author.PSObject.Properties.Name -contains 'Title') { $ownerName = $author.Title }
                    else { $ownerName = $author.ToString() }
                } catch { $ownerName = $author.ToString() }
            }

            $rec = [PSCustomObject]@{
                RelativePath = $relNorm
                Name = $item["FileLeafRef"]
                Size = ([int64]($item["File_x0020_Size"] -as [int64]))
                Modified = ($item["Modified"] -as [datetime]).ToString('o')
                Path = $item["FileRef"]
                Owner = $ownerName
            }

            $partId = Get-PartitionId -key $relNorm -numPartitions $numPartitions -md5 $md5
            $partitionBuffers[$partId] += $rec
            if ($partitionBuffers[$partId].Count -ge $flushSize) {
                $file = Join-Path $PartitionDir ("part_$partId.csv")
                $partitionBuffers[$partId] | Export-Csv -Path $file -Append -NoTypeInformation -Encoding UTF8
                $partitionBuffers[$partId] = @()
            }
        }
    }

    $partKeys = @($partitionBuffers.Keys)
    foreach ($partId in $partKeys) {
        $buf = $partitionBuffers[$partId]
        if ($buf -and $buf.Count -gt 0) {
            $file = Join-Path $PartitionDir ("part_$partId.csv")
            $buf | Export-Csv -Path $file -Append -NoTypeInformation -Encoding UTF8
            $partitionBuffers[$partId] = @()
        }
    }

    Write-Host "Partitioning complete. Partition files in: $PartitionDir" -ForegroundColor Green

    Write-Host "Matching source files against partitioned index..." -ForegroundColor Cyan
    $PartitionCache = @{}
    $PartitionCacheLRU = New-Object System.Collections.ArrayList

    function Save-PartitionMap {
        param($partId, $map)
        $file = Join-Path $PartitionDir ("part_$partId.csv")
        $outList = @()
        foreach ($entry in $map.GetEnumerator()) {
            $k = $entry.Key
            $v = $entry.Value
            $modVal = ''
            if ($v.Modified) {
                try { $modVal = $v.Modified.ToString('o') } catch { $modVal = '' }
            }
            $o = [PSCustomObject]@{
                RelativePath = $k
                Name = $v.Name
                Size = $v.Size
                Modified = $modVal
                Path = $v.Path
                Owner = $v.Owner
            }
            $outList += $o
        }
        if ($outList.Count -gt 0) {
            $outList | Export-Csv -Path $file -NoTypeInformation -Force -Encoding UTF8
        } else {
            if (Test-Path $file) { Remove-Item $file -Force -ErrorAction SilentlyContinue }
        }
    }

    function Load-PartitionMap {
        param($partId)
        if ($PartitionCache.ContainsKey($partId)) { return $PartitionCache[$partId] }
        $file = Join-Path $PartitionDir ("part_$partId.csv")
        $map = @{}
        if (Test-Path $file) {
            $rows = Import-Csv -Path $file -Encoding UTF8
            foreach ($r in $rows) {
                if ([string]::IsNullOrWhiteSpace($r.RelativePath)) { continue }
                $k = $r.RelativePath.ToLower().Trim('/')
                $modVal = $null
                if ($r.Modified) {
                    try { $modVal = [datetime]::Parse($r.Modified) } catch { $modVal = $null }
                }
                $map[$k] = [PSCustomObject]@{
                    Name = $r.Name
                    Size = ([int64]$r.Size)
                    Modified = $modVal
                    Path = $r.Path
                    Owner = $r.Owner
                }
            }
        }
        $PartitionCache[$partId] = $map
        $PartitionCacheLRU.Add($partId) | Out-Null

        while ($PartitionCache.Count -gt $MaxCachedPartitions) {
            $old = $PartitionCacheLRU[0]
            $oldMap = $PartitionCache[$old]
            Save-PartitionMap $old $oldMap
            $PartitionCache.Remove($old)
            $PartitionCacheLRU.RemoveAt(0)
        }
        return $PartitionCache[$partId]
    }

    $siteUri = [uri]$SiteUrl
    foreach ($srcFile in $SourceFiles) {
        $srcRelNorm = Normalize-RelPath $srcFile.RelativePath
        $partId = Get-PartitionId -key $srcRelNorm -numPartitions $numPartitions -md5 $md5
        $map = Load-PartitionMap $partId
        if ($map.ContainsKey($srcRelNorm)) {
            $targetInfo = $map[$srcRelNorm]
            $srcFile.Exists_InTarget = $true
            $srcFile.TargetSize = $targetInfo.Size
            $srcFile.SizeMatch = ($srcFile.SizeBytes -eq $targetInfo.Size)
            # Determine human-friendly size match status: YES / CHECK (target larger) / CORRUPT (target smaller)
            try {
                if ($null -ne $srcFile.SizeBytes -and $null -ne $targetInfo.Size) {
                    if ($srcFile.SizeBytes -eq $targetInfo.Size) { $srcFile.SizeMatchStatus = 'YES' }
                    elseif ($targetInfo.Size -gt $srcFile.SizeBytes) { $srcFile.SizeMatchStatus = 'CHECK' }
                    else { $srcFile.SizeMatchStatus = 'CORRUPT' }
                } else { $srcFile.SizeMatchStatus = $null }
            } catch { $srcFile.SizeMatchStatus = $null }
            $srcFile.TargetModified = $targetInfo.Modified
            $srcFile.TargetOwner = $targetInfo.Owner
            if ($targetInfo.Path) {
                if ($targetInfo.Path -match '^https?://') { $srcFile.DestinationPath = $targetInfo.Path }
                else { $srcFile.DestinationPath = "$($siteUri.Scheme)://$($siteUri.Host)$($targetInfo.Path)" }
            } else { $srcFile.DestinationPath = $null }
            $srcFile.OwnerMatch = Test-OwnerMatch -SourceOwner $srcFile.SourceOwner -TargetOwner $targetInfo.Owner
            if ($targetInfo.Modified) {
                $timeDiff = [Math]::Abs(($srcFile.Modified - $targetInfo.Modified).TotalSeconds)
                $srcFile.DateMatch = ($timeDiff -le 2)
            }
            $map.Remove($srcRelNorm) | Out-Null
            $MatchedCount++
        } else {
            $srcFile.Exists_InTarget = $false
            $srcFile.SizeMatchStatus = $null
        }
    }

    $cachedIds = @($PartitionCache.Keys)
    foreach ($partId in $cachedIds) { Save-PartitionMap $partId $PartitionCache[$partId] }

    Write-Host "Matched $MatchedCount file(s)" -ForegroundColor Green

    # Build per-folder summary rows (Type = 'folder')
    try {
        $baseServer = $null
        if ($TargetServerRelativeUrl) { $baseServer = $TargetServerRelativeUrl } else { $baseServer = $TargetList.RootFolder.ServerRelativeUrl }
        $baseServer = $baseServer.TrimEnd('/')

        $sourceDirs = Get-ChildItem -Path $SourcePath -Recurse -Directory -Force -ErrorAction SilentlyContinue
        foreach ($dir in $sourceDirs) {
            $rel = $dir.FullName.Substring($SourcePath.Length).TrimStart('\','/') -replace '\\','/'
            $relNorm = Normalize-RelPath $rel

                # Only consider files directly inside this folder (do not include files in subfolders)
                $filesInFolder = @($SourceFiles | Where-Object { $_.RelativePath } | Where-Object {
                    $fr = $_.RelativePath.ToLower()
                    if ([string]::IsNullOrWhiteSpace($relNorm)) {
                        # If folder relative path is empty (shouldn't happen for subfolders), treat files without any subpath
                        return -not ($fr.Contains('/'))
                    }
                    if ($fr.StartsWith("$relNorm/")) {
                        $suffix = $fr.Substring($relNorm.Length).TrimStart('/')
                        return -not ($suffix.Contains('/'))
                    }
                    return $false
                }
                )
                $totalFiles = $filesInFolder.Count
                if ($totalFiles -eq 0) { continue }
                $missing = @($filesInFolder | Where-Object { ($_.Type -eq 'file') -and ($_.Exists_InTarget -ne $true) }).Count

            if ($missing -eq 0) { $status = 'Migrated' } else { $status = "Failed - $missing" }

            $destUrl = "$($siteUri.Scheme)://$($siteUri.Host)$baseServer"
            if ($relNorm) { $destUrl = "$destUrl/$relNorm" }

            $folderRec = [PSCustomObject]@{
                FileName = $dir.Name
                Type = 'folder'
                RelativePath = $rel
                SourcePath = $dir.FullName
                FullPath = $dir.FullName
                DestinationPath = $destUrl
                Exists_InTarget = $status
                SizeBytes = $null
                TargetSize = $null
                SizeMatch = $null
                SizeMatchStatus = $null
                Modified = $null
                TargetModified = $null
                DateMatch = $null
                SourceOwner = $null
                TargetOwner = $null
                OwnerMatch = $null
            }

            $SourceFiles += $folderRec
        }
    } catch {
        # Non-fatal: continue without folder rows
    }

} catch {
    Write-Host "ERROR scanning target:" -ForegroundColor Red
    try {
        if ($_.Exception) {
            Write-Host "Exception.Message: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Exception.Type: $($($_.Exception.GetType().FullName))" -ForegroundColor Red
        } else {
            Write-Host $_.ToString() -ForegroundColor Red
        }
        if ($PSBoundParameters) { Write-Host "(Debug) PSBoundParameters present" -ForegroundColor DarkYellow }
        if ($_.ScriptStackTrace) { Write-Host "ScriptStackTrace:" -ForegroundColor Red; Write-Host $_.ScriptStackTrace -ForegroundColor Red }
        Write-Host "Full error object:" -ForegroundColor Red
        $_ | Format-List * -Force | Out-String | Write-Host -ForegroundColor Red
    } catch {
        Write-Host "(Additionally failed to format exception) $_" -ForegroundColor Red
    }
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

# Show total elapsed time for the comparison (terminal only)
try {
    $elapsed = (Get-Date) - $ScriptStartTime
    if ($elapsed.TotalHours -ge 1) {
        $elapsedStr = "{0}h {1}m {2}s" -f [int]$elapsed.TotalHours, $elapsed.Minutes, $elapsed.Seconds
    } elseif ($elapsed.TotalMinutes -ge 1) {
        $elapsedStr = "{0}m {1}s" -f [int]$elapsed.TotalMinutes, $elapsed.Seconds
    } else {
        $elapsedStr = "{0}s" -f $elapsed.Seconds
    }
    Write-Host ""
    Write-Host "Time elapsed: $elapsedStr" -ForegroundColor Cyan
} catch {
    # Non-fatal: ignore if timing fails
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
        @{Name='Type'; Expression={$_.Type}},
        @{Name='SourcePath'; Expression={$_.SourcePath}},
        @{Name='DestinationPath'; Expression={$_.DestinationPath}},
        RelativePath,
        @{Name='Exists_InTarget'; Expression={ if ($_.Type -eq 'folder') { $_.Exists_InTarget } else { if ($_.Exists_InTarget) { 'YES' } else { 'NO' } } }},
        @{Name='SourceSize_Bytes'; Expression={$_.SizeBytes}},
        @{Name='TargetSize_Bytes'; Expression={$_.TargetSize}},
        @{Name='SizeMatch'; Expression={
            if ($_.Type -eq 'folder') { 'N/A' }
            elseif ($_.SizeMatchStatus) { $_.SizeMatchStatus }
            elseif ($_.SizeMatch -eq $true) { 'YES' }
            elseif ($_.SizeMatch -eq $false) { 'NO' }
            else { 'N/A' }
        }},
        @{Name='SourceModified'; Expression={$_.Modified}},
        @{Name='TargetModified'; Expression={$_.TargetModified}},
        @{Name='DateMatch'; Expression={if ($_.DateMatch -eq $true) { 'YES' } elseif ($_.DateMatch -eq $false) { 'NO' } else { 'N/A' }}},
        @{Name='SourceOwner'; Expression={$_.SourceOwner}},
        @{Name='TargetOwner'; Expression={$_.TargetOwner}},
        @{Name='OwnerMatch'; Expression={if ($_.OwnerMatch -eq $true) { 'YES' } elseif ($_.OwnerMatch -eq $false) { 'NO' } else { 'N/A' }}} |
        Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "SUCCESS: Comparison report created!" -ForegroundColor Green
    Write-Host "Location: $ExportPath" -ForegroundColor Green
    
} catch {
    Write-Host "ERROR: Failed to export CSV: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
