#### Parameters – Get-PCInfo-Advanced.ps1

- **ComputerName** – target computer name. If omitted, the script will ask interactively.
- **Credential** – optional `Get-Credential` object for remote connections (different domain/user). If omitted, current user context is used.
- **Export** – when present, results are also exported to CSV and JSON files.
- **ExportDir** – directory where export files will be created (default: current directory).  
  Files are named like: `PCInfo_HOSTNAME_YYYYMMDD_HHMMSS.csv/json`.

#### Examples – Get-PCInfo-Advanced.ps1

- Run locally, show info in console only:  
  `.\Get-PCInfo-Advanced.ps1`

- Run for a remote computer with current credentials:  
  `.\Get-PCInfo-Advanced.ps1 -ComputerName PC1234`

- Run for a remote computer with alternate credentials:  
  `$cred = Get-Credential`  
  `.\Get-PCInfo-Advanced.ps1 -ComputerName LAPTOP-42 -Credential $cred`

- Run for a remote computer and export results to CSV + JSON:  
  `.\Get-PCInfo-Advanced.ps1 -ComputerName PC1234 -Export -ExportDir "C:\Temp\PCReports"`

> Tip: For best results, run from an elevated PowerShell session (Run as administrator) and make sure WinRM or DCOM access to the target host is allowed by your environment.

---

### Install-MsiSilent.ps1

Generic helper for **silent installation of any MSI** with full logging.

This script standardizes how you run `msiexec.exe` in quiet mode and automatically:
- validates that the path points to an `.msi` file
- creates a log directory (if it doesn’t exist)
- generates a unique log file name with timestamp
- runs `msiexec /i ... /qn /norestart /L*v <logfile>`
- returns the exit code and prints a success / warning message

#### Parameters – Install-MsiSilent.ps1

- **Path** – full path to the MSI file (local path or UNC). Must exist.
- **LogDirectory** – folder where the MSI log will be created (default: `.\Logs` next to the script).
- **AdditionalArgs** – optional extra arguments passed to `msiexec.exe`  
  (e.g. `TRANSFORMS=app.mst`, `ALLUSERS=1`, custom properties, etc.).

#### Examples – Install-MsiSilent.ps1

- Basic silent install with log:  
  `.\Install-MsiSilent.ps1 -Path .\setup.msi`

- Install from network share with additional msiexec arguments:  
  `.\Install-MsiSilent.ps1 -Path \\SERVER\Share\App.msi -AdditionalArgs "TRANSFORMS=app.mst"`

- Custom log directory:  
  `.\Install-MsiSilent.ps1 -Path .\client.msi -LogDirectory 'C:\Logs\MSI'`

---

All scripts in this repository are designed to be reusable in different corporate environments.  
Feel free to clone, adapt and extend them for your own tooling.
