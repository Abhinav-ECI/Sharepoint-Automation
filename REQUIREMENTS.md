# REQUIREMENTS — SEED Migration Tool

Copy and paste the following into an elevated Windows PowerShell (Run as Administrator) to install the module required by the SEED migration comparison script:

```powershell
Install-Module -Name SharePointPnPPowerShellOnline -Scope CurrentUser -Force
```

If you encounter policy errors, run:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
Install-Module -Name SharePointPnPPowerShellOnline -Scope CurrentUser -Force
```

Optional (modern cross-platform PnP module — not required by this script):

```powershell
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
```

Notes:
- The script auto-switches to Windows PowerShell when executed under PowerShell 7+. The `SharePointPnPPowerShellOnline` module is intended for Windows PowerShell (5.1).
- After installing the module, re-run the comparison script: `& 'SEED Migration - Compare-Source-Target-Files (Hyperlink).ps1'`.
