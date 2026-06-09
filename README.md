# Sharepoint-Automation
This repository contains automation helpers for QA, primarily a SEED migration comparison tool.

## SEED Migration Tool — Compare Source to SharePoint

The primary script is [SEED Migration - Compare-Source-Target-Files (Hyperlink).ps1](SEED%20Migration%20-%20Compare-Source-Target-Files%20%28Hyperlink%29.ps1).
It compares a local source folder tree to a SharePoint site (or a specific target folder) and exports a CSV report summarizing which files exist in the target, and which have size, date, or owner mismatches.

Key features:
- Scans local source recursively and collects file metadata (including hidden and system files) using `Get-ChildItem -Recurse -File -Force`.
- Accepts either a SharePoint site URL or a pasted SharePoint folder hyperlink — the script extracts server-relative paths from hyperlinks automatically.
- Connects to SharePoint using `SharePointPnPPowerShellOnline` and streams list items.
- Efficient, disk-backed partitioning of target index to handle very large libraries; configurable with `-PartitionBatchSize`, `-PartitionFlushSize`, and `-MaxCachedPartitions`.
- Matching is performed on a normalized relative path (lowercased, forward slashes) produced by the `Normalize-RelPath` function — this includes the filename and extension (case-insensitive).
- Compares file size, last-modified time (2‑second tolerance), and owner (heuristic `Test-OwnerMatch`), and reports results in a CSV saved to the Desktop.
- Auto-switches to Windows PowerShell when run under PowerShell Core for improved compatibility.

Why this helps SEED migrations:
- Quickly identifies files that are missing on the target SharePoint site.
- Highlights size, date, or owner mismatches so you can prioritize remediation.
- Produces an exportable, auditable CSV for reporting and follow-up.
- Scales to large libraries by partitioning the target index to disk to avoid high memory use.

Usage examples
Run interactively (folder picker or prompts):

```
& 'SEED Migration - Compare-Source-Target-Files (Hyperlink).ps1'
```

Run with parameters:

```
& 'SEED Migration - Compare-Source-Target-Files (Hyperlink).ps1' -SourcePath 'C:\MySource' -SharePointUrl 'https://contoso/sites/site/Shared Documents/folder' -PartitionBatchSize 50000
```

Notes and customization
- Hidden files: the script includes hidden/system files by default because it uses `-Force`. To exclude hidden files, filter the results:

```
Get-ChildItem -Path $SourcePath -Recurse -File -Force |
	Where-Object { -not ($_.Attributes -band [System.IO.FileAttributes]::Hidden) }
```

- Filename-only matching: to match only by filename (ignore relative path) change the partition key and matching logic to use `$_Name` or strip extensions with `[System.IO.Path]::GetFileNameWithoutExtension()`.

- Requirements: the `SharePointPnPPowerShellOnline` module must be installed. Install with:

```
Install-Module SharePointPnPPowerShellOnline -Force
```

Output
- A CSV named like `Compare_Source_Target_YYYYMMDD_HHmmss.csv` is written to your Desktop. The CSV includes columns for source/target size, modified date, owner, and match results.

If you'd like, I can:
- add a short example that shows how to exclude hidden files,
- or patch the script to add a `-ExcludeHidden` switch.

