<#
.SYNOPSIS
  Read a migration report (CSV or Excel) and verify whether each "Item name" exists
  in the destination SharePoint location. Appends an `ExistsInTarget` column to the
  exported CSV report.

USAGE
  .\Compare-Report-Source-Target.ps1 -ReportPath "C:\path\to\report.csv" -SharePointUrl "https://tenant/sites/sitename"

NOTES
  - Uses `Connect-PnPOnline -UseWebLogin` to authenticate. Requires the
    SharePointPnPPowerShellOnline module to be installed.
  - Supports CSV and Excel (.xls/.xlsx) input. Excel is converted to CSV via COM.
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$ReportPath = "",
    [Parameter(Mandatory=$false)]
    [string]$SharePointUrl = "",
    [Parameter(Mandatory=$false)]
    [string]$OutFile = "",
    [Parameter(Mandatory=$false)]
    [switch]$ShowProgress,
    [Parameter(Mandatory=$false)]
    [string]$LogPath = "",
    [Parameter(Mandatory=$false)]
    [switch]$AppendLog,
    [Parameter(Mandatory=$false)]
    [switch]$KeepTempCsv,
    [Parameter(Mandatory=$false)]
    [int]$ThrottleDelayMs = 80
)

# Auto-switch to Windows PowerShell if running in PowerShell 7+
$PSVersion = $PSVersionTable.PSVersion.Major
if ($PSVersion -gt 5) {
    Write-Host "PowerShell 7+ detected. Switching to Windows PowerShell..." -ForegroundColor Yellow
    $ArgumentList = @()
    foreach ($Key in $PSBoundParameters.Keys) {
        $ArgumentList += "-$Key"
        $val = $PSBoundParameters[$Key]
        if ($val -is [System.Boolean]) {
            if ($val) { } else { }
        } else {
            $ArgumentList += "`"$val`""
        }
    }
    & powershell.exe -NoProfile -File $PSCommandPath $ArgumentList
    exit
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

function Convert-ExcelToCsv {
    param([string]$Path)
    $tempCsv = Join-Path $env:TEMP ("compare_report_{0}.csv" -f ([guid]::NewGuid().ToString()))
    $excel = $null
    $wb = $null
    $ws = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($Path)
        try { $ws = $wb.Worksheets.Item(1) } catch { }
        $wb.SaveAs($tempCsv, 6) # xlCSV = 6
        $wb.Close($false)
        if ($excel) { $excel.Quit() }
        return $tempCsv
    } catch {
        throw "Failed to convert Excel to CSV: $_"
    } finally {
        if ($ws -ne $null) {
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null } catch { }
            $ws = $null
        }
        if ($wb -ne $null) {
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null } catch { }
            $wb = $null
        }
        if ($excel -ne $null) {
            try { $excel.Quit() } catch { }
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
            $excel = $null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function NormalizeKey {
    param([string]$s)
    if (-not $s) { return '' }
    return ($s -replace '\s+','' -replace '[^a-z0-9]','').ToLower()
}

function Find-Header {
    param([array]$candidates, [hashtable]$normHeaders)
    foreach ($p in $candidates) {
        # prefer exact normalized key match first
        foreach ($k in $normHeaders.Keys) {
            if ($k -eq $p) { return $normHeaders[$k] }
        }
        # then prefer keys that start with the candidate
        foreach ($k in $normHeaders.Keys) {
            if ($k -like "$p*") { return $normHeaders[$k] }
        }
        # finally, fall back to any key that contains the candidate
        foreach ($k in $normHeaders.Keys) {
            if ($k -like "*$p*") { return $normHeaders[$k] }
        }
    }
    return $null
}

function Parse-SiteFromUrl {
    param([string]$Url)
    $result = @{ SiteUrl = $null; SitePrefix = $null; TargetFolder = $null }
    try {
        if (-not $Url) { return $result }
        $uStr = $Url.Trim()
        if ($uStr -match '^://') { $uStr = 'https' + $uStr }
        elseif ($uStr -match '^//') { $uStr = 'https:' + $uStr }
        elseif ($uStr -notmatch '^[a-zA-Z][a-zA-Z0-9+.-]*://') {
            # If it's a server-relative path (starts with '/') we can't know the host here.
            # Leave it as-is and let caller handle server-relative values when possible.
            if ($uStr.StartsWith('/')) {
                # keep as server-relative
            } else {
                $uStr = 'https://' + $uStr.TrimStart('/')
            }
        }
        $uri = [uri]$uStr
        $absPath = [System.Uri]::UnescapeDataString($uri.AbsolutePath).TrimEnd('/')
        $parts = $absPath -split '/' | Where-Object { $_ }
        $siteIndex = -1
        for ($i = 0; $i -lt $parts.Length; $i++) {
            if ($parts[$i] -eq 'sites' -or $parts[$i] -eq 'teams') { $siteIndex = $i; break }
        }
        if ($siteIndex -ge 0 -and $siteIndex + 1 -lt $parts.Length) {
            $siteParts = $parts[0..($siteIndex + 1)]
            $result.SiteUrl = "$($uri.Scheme)://$($uri.Host)/$($siteParts -join '/')"
            $result.SitePrefix = "/$($siteParts -join '/')"
            if ($siteIndex + 2 -lt $parts.Length) {
                $targetParts = $parts[($siteIndex + 2)..($parts.Length - 1)]
                $result.TargetFolder = "/$($siteParts -join '/')/$($targetParts -join '/')"
            }
        } else {
            $result.SiteUrl = "$($uri.Scheme)://$($uri.Host)$absPath"
            $result.SitePrefix = $absPath
        }
    } catch {
        $result.SiteUrl = $null
    }
    return $result
}

# Simple logging helper. Writes timestamped lines to $LogPath if set.
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO',
        [switch]$Console
    )
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    if ($LogPath) {
        try {
            Add-Content -Path $LogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch { }
    }
    if ($Console) { Write-Host $line }
}

if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $ofd.Filter = "CSV files (*.csv)|*.csv|Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*"
        $ofd.Multiselect = $false
        $ofd.Title = "Select migration report (CSV or Excel)"
        $dlg = $ofd.ShowDialog()
        if ($dlg -eq [System.Windows.Forms.DialogResult]::OK) {
            $ReportPath = $ofd.FileName
        } else {
            $ReportPath = Read-Host "Enter path to migration report (CSV or Excel)"
        }
    } catch {
        # Fallback to text prompt if GUI isn't available
        $ReportPath = Read-Host "Enter path to migration report (CSV or Excel)"
    }
}
$ReportPath = $ReportPath.Trim('"').Trim("'")
if (-not (Test-Path $ReportPath)) { Write-Host "Report not found: $ReportPath" -ForegroundColor Red; exit 1 }

$ext = [IO.Path]::GetExtension($ReportPath).ToLower()
if ($ext -in @('.xls','.xlsx')) {
    Write-Host "Converting Excel to CSV..." -ForegroundColor Cyan
    try { $csvPath = Convert-ExcelToCsv -Path $ReportPath } catch { Write-Host $_ -ForegroundColor Red; exit 1 }
} else {
    $csvPath = $ReportPath
}

Write-Host "Loading report..." -ForegroundColor Cyan
try {
    $rows = Import-Csv -Path $csvPath -Encoding UTF8
} catch {
    Write-Host "Failed to read CSV: $_" -ForegroundColor Red
    exit 1
}
if (-not $rows -or $rows.Count -eq 0) { Write-Host "No rows found in the report" -ForegroundColor Yellow; exit 0 }

# By default logging is disabled. To enable, pass -LogPath "C:\path\to\logfile.log"
Write-Host "Report loaded: $csvPath ($($rows.Count) rows)" -ForegroundColor Cyan

# Build normalized header map
$headers = $rows[0].psobject.properties | Select-Object -ExpandProperty Name
$normHeaders = @{}
foreach ($h in $headers) { $k = NormalizeKey $h; if (-not $normHeaders.ContainsKey($k)) { $normHeaders[$k] = $h } }


$colItem = Find-Header -candidates @('itemname','filename','item','name') -normHeaders $normHeaders
$colSource = Find-Header -candidates @('sourcepath','source','sourcefolder','localpath') -normHeaders $normHeaders
$colDest = Find-Header -candidates @('destination','dest','target','destinationpath') -normHeaders $normHeaders
$colStatus = Find-Header -candidates @('status','migrationstatus','result','message') -normHeaders $normHeaders

# Fallback: try looser matching if header detection missed the source column
if (-not $colSource) {
    # Prefer headers that indicate a path or folder (avoid ID columns)
    foreach ($h in $headers) {
        if ($h -match '(?i)source' -and $h -match '(?i)path|folder|file|location|filepath') { $colSource = $h; break }
    }
    # If none, pick a source-like header that isn't an ID column
    if (-not $colSource) {
        foreach ($h in $headers) {
            if ($h -match '(?i)source' -and -not ($h -match '(?i)\bid\b')) { $colSource = $h; break }
        }
    }
    # Last resort: any header with 'local' or 'src'
    if (-not $colSource) {
        foreach ($h in $headers) {
            if ($h -match '(?i)local' -or $h -match '(?i)src') { $colSource = $h; break }
        }
    }
}

# Resolved headers (kept for internal use)

if (-not $colItem -or -not $colDest) {
    Write-Host "Could not find required columns. Need an 'Item name' and a 'Destination' column." -ForegroundColor Red
    exit 1
}

if (-not $SharePointUrl) {
    # Prefer a destination value that contains the SharePoint host or an absolute URL
    $firstDest = $rows | Where-Object { $_.$colDest -and ($_."$colDest" -match 'sharepoint\.com' -or $_."$colDest" -match '^https?://' -or $_."$colDest" -match '^//') } | Select-Object -First 1 -ExpandProperty $colDest
    if (-not $firstDest) {
        # Fallback to any non-empty destination
        $firstDest = $rows | Where-Object { $_.$colDest } | Select-Object -First 1 -ExpandProperty $colDest
    }
    if ($firstDest) {
        $firstDest = Get-SharePointPathFromHyperlink -Url $firstDest
        $parse = Parse-SiteFromUrl -Url $firstDest
        $SiteUrl = $parse.SiteUrl
        $SitePrefix = $parse.SitePrefix
    } else { $SiteUrl = $null }
} else {
    $SharePointUrl = $SharePointUrl.Trim('"').Trim("'")
    $SharePointUrl = Get-SharePointPathFromHyperlink -Url $SharePointUrl
    $parse = Parse-SiteFromUrl -Url $SharePointUrl
    $SiteUrl = $parse.SiteUrl
    $SitePrefix = $parse.SitePrefix
}

if (-not $SiteUrl) {
    Write-Host "Could not determine SharePoint site URL. Provide -SharePointUrl or ensure Destination values are full SharePoint links." -ForegroundColor Red
    exit 1
}

Write-Host "Site URL: $SiteUrl" -ForegroundColor Green

Write-Host "Checking for SharePointPnPPowerShellOnline module..." -ForegroundColor Cyan
$ModuleExists = Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline
if (-not $ModuleExists) { Write-Host "Module not found. Install with: Install-Module SharePointPnPPowerShellOnline -Force" -ForegroundColor Red; exit 1 }

try {
    Import-Module SharePointPnPPowerShellOnline -ErrorAction Stop -WarningAction SilentlyContinue
} catch {
    Write-Host "Failed to import PnP module: $_" -ForegroundColor Red
    exit 1
}

Write-Host "Connecting to SharePoint (web login)..." -ForegroundColor Cyan
try {
    $Conn = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ReturnConnection -WarningAction Ignore -ErrorAction Stop
    Write-Host "Connected." -ForegroundColor Green
} catch {
    Write-Host "Failed to connect: $_" -ForegroundColor Red
    exit 1
}

$existsCache = @{}
function Test-FileInSharePoint {
    param([string]$ServerRelativeFile)
    if ($existsCache.ContainsKey($ServerRelativeFile)) { return $existsCache[$ServerRelativeFile] }
    try {
        # Try to get the file. This will error if the file does not exist.
        Get-PnPFile -Url $ServerRelativeFile -Connection $Conn -ErrorAction Stop | Out-Null
        $existsCache[$ServerRelativeFile] = $true
        return $true
    } catch {
        $existsCache[$ServerRelativeFile] = $false
        return $false
    }
}

$results = @()
$count = 0
foreach ($r in $rows) {
    $count++
    $itemName = ($r."$colItem") -as [string]
    if (-not $itemName) { $itemName = ($r.$colItem) -as [string] }

    # Robustly retrieve source value from detected header (or fallback to property lookup)
    $srcVal = $null
    if ($colSource) {
        $prop = $r.PSObject.Properties | Where-Object { $_.Name -eq $colSource } | Select-Object -First 1
        if ($prop) { $srcVal = [string]$prop.Value } else { $srcVal = ($r."$colSource") -as [string] }
    }

    $dstVal = ($r."$colDest") -as [string]
    $statusVal = $null; if ($colStatus) { $statusVal = ($r."$colStatus") -as [string] }

    # Keep original source path (including filename) and prepare a normalized path for existence checks
    $sourceOriginal = $null
    $sourceForTest = $null
    if ($srcVal) {
        $sourceOriginal = $srcVal.Trim()
        $sourceForTest = $sourceOriginal
        if ($sourceForTest -match '(?i)^file:') {
            try { $u = [uri]$sourceForTest; $sourceForTest = $u.LocalPath } catch { $sourceForTest = $sourceForTest -replace '^file:///?','' }
        }
    }

    # Determine source folder for internal use if needed
    $srcFolder = $null
    if ($sourceForTest) {
        try { $srcFolder = Split-Path -Path $sourceForTest -Parent } catch { $srcFolder = $sourceForTest }
        if (-not $srcFolder) { $srcFolder = $sourceForTest }
    }

    # Check existence in source (file)
    $existsSource = $false
    if ($sourceForTest) {
        try { $existsSource = Test-Path -Path $sourceForTest -PathType Leaf } catch { $existsSource = $false }
    }

    # Destination folder / path
    $dstNormalized = $dstVal
    if ($dstNormalized -and ($dstNormalized -match '^(https?://)' -or $dstNormalized -match 'id=')) { $dstNormalized = Get-SharePointPathFromHyperlink -Url $dstNormalized }

    $serverRelativeFile = $null
    if ($dstNormalized) {
        if ($dstNormalized -match '^(https?://)') {
            try {
                $u = [uri]$dstNormalized
                $absPath = [System.Uri]::UnescapeDataString($u.AbsolutePath).TrimEnd('/')
                $basename = [IO.Path]::GetFileName($absPath)
                if ($basename -and $itemName -and ($basename -ieq $itemName)) {
                    $serverRelativeFile = $absPath
                } else {
                    $folder = Split-Path -Path $absPath -Parent
                    $serverRelativeFile = "$folder/$itemName"
                }
            } catch {
                # not a full uri; fall back
                $dstNormalized = $dstNormalized.TrimEnd('/')
                $serverRelativeFile = "$dstNormalized/$itemName"
            }
        } elseif ($dstNormalized.StartsWith('/')) {
            $dstNormalized = $dstNormalized.TrimEnd('/')
            $serverRelativeFile = "$dstNormalized/$itemName"
        } else {
            # site-relative folder without leading slash; prefix with site base
            $dstNormalized = $dstNormalized.Trim('/')
            $pref = $SitePrefix.TrimEnd('/')
            if (-not $pref.StartsWith('/')) { $pref = "/$pref" }
            $serverRelativeFile = "$pref/$dstNormalized/$itemName"
        }
    } else {
        # No destination provided; construct from site prefix only
        $pref = $SitePrefix.TrimEnd('/')
        if (-not $pref.StartsWith('/')) { $pref = "/$pref" }
        $serverRelativeFile = "$pref/$itemName"
    }

    # Normalize double slashes
    $serverRelativeFile = $serverRelativeFile -replace '//+','/'

    # Test existence in target
    $exists = $false
    try {
        $exists = Test-FileInSharePoint -ServerRelativeFile $serverRelativeFile
    } catch {
        $exists = $false
    }

    $rowResult = [PSCustomObject]@{
        Source = $sourceOriginal
        Destination = $dstVal
        ItemName = $itemName
        OriginalStatus = $statusVal
        ExistsInSource = if ($existsSource) { 'YES' } else { 'NO' }
        ServerRelativePath = $serverRelativeFile
        ExistsInTarget = if ($exists) { 'YES' } else { 'NO' }
    }
    $results += $rowResult

    # Logging per-row
    Write-Log "Row ${count}: Item='$itemName' SourceExists=$($rowResult.ExistsInSource) TargetExists=$($rowResult.ExistsInTarget) TargetPath=$($rowResult.ServerRelativePath)" -Level INFO
    if ($ShowProgress) {
        $percent = 0
        if ($rows.Count -gt 0) { $percent = [int](($count / $rows.Count) * 100) }
        Write-Progress -Activity "Checking files" -Status "Row $count of $($rows.Count): $itemName" -PercentComplete $percent
        # Fallback single-line console status (overwrites same line)
        try {
            $statusLine = "Row $count/$($rows.Count): $itemName ($percent`%)"
            Write-Host ("`r" + $statusLine) -NoNewline -ForegroundColor DarkCyan
        } catch { }
    }

    if ($ThrottleDelayMs -gt 0) { Start-Sleep -Milliseconds $ThrottleDelayMs }
}

if (-not $OutFile) {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $OutFile = Join-Path $desktop ("Compare_Report_Checked_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
if ($ShowProgress) { Write-Host "" }

try {
    $results | Select-Object Source, Destination, ItemName, OriginalStatus, ExistsInSource, ServerRelativePath, ExistsInTarget |
        Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
    Write-Host "Results exported: $OutFile" -ForegroundColor Green
    Write-Log "Results exported to $OutFile" -Level INFO -Console
} catch {
    Write-Log "Failed to export results: $_" -Level ERROR -Console
    Write-Host "Failed to export results: $_" -ForegroundColor Red
    exit 1
}

if (-not $KeepTempCsv -and $csvPath -and ($csvPath -ne $ReportPath)) {
    try {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Start-Sleep -Milliseconds 200
        Remove-Item $csvPath -ErrorAction SilentlyContinue
    } catch { }
}

Write-Host "Done. Processed $($rows.Count) row(s)." -ForegroundColor Green
