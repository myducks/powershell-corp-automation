### Get-PCInfo-Advanced.ps1

Collects a **detailed inventory** of a workstation or notebook using CIM/WMI and (optionally) Active Directory.

This script is intended as a “corporate Swiss-army knife” for IT support and endpoint diagnostics.

**What it collects:**

- Basic system info (hostname, model, manufacturer, logged-on user, domain/workgroup)
- OS details (edition, architecture, build, version, install date, last boot, uptime)
- BIOS info (serial number, version, release date)
- CPU, GPU, RAM modules and total memory
- Local disks: size, free space and filesystem per drive
- Network adapters: IPs, DNS servers, DHCP status
- Battery status, estimated runtime, wear level, design/full capacity and charge cycles (if available)
- Docking station summary (USB-C/Thunderbolt docks, DisplayLink, Dell/Lenovo/HP, etc.)
- Monitors: manufacturer, model, serial, size (approx. diagonal), connection type
- Active Directory computer object (OU and groups) – if RSAT AD module is installed
- Network ports: top listening TCP ports and a sample of established connections
- System/Application error events from the last 7 days
- Installed updates summary (Win32_QuickFixEngineering + recent Windows Update successes)
- Security & VPN:
  - CrowdStrike (presence + version, if detected)
  - Cisco Secure Client / AnyConnect (presence + version, if detected)


#### Parameters

```powershell
[CmdletBinding()]
param(
    [string]$ComputerName = $(Read-Host "Computer name"),
    [pscredential]$Credential,
    [switch]$Export,
    [string]$ExportDir = "."
)
