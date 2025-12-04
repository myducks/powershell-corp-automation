[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ComputerName = $(Read-Host "Computer name"),
    [pscredential]$Credential,
    [switch]$Export,
    [string]$ExportDir = "."
)

function Write-Info { param($msg,$color="Cyan") Write-Host $msg -ForegroundColor $color }

function Add-Row {
    param([ref]$list,[string]$section,[string]$property,[string]$value)
    if ($null -eq $value -or ($value -is [string] -and [string]::IsNullOrWhiteSpace($value))) { $value = "No data" }
    $list.Value.Add([pscustomobject]@{ Section=$section; Property=$property; Value=$value })
}

function Convert-WmiDateSafe {
    param($value)
    try {
        if ($null -eq $value) { return $null }
        if ($value -is [datetime]) { return $value }
        $s = [string]$value
        if ([string]::IsNullOrWhiteSpace($s)) { return $null }
        return [Management.ManagementDateTimeConverter]::ToDateTime($s)
    } catch { return $null }
}

function Convert-EdidString {
    param([System.Array]$Data)
    try {
        if ($null -eq $Data) { return "No data" }
        ($Data | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ''
    } catch { "No data" }
}

function Get-RegStringRemote {
    param(
        [Microsoft.Management.Infrastructure.CimSession]$Cim,
        [uint32]$Hive,
        [string]$Path,
        [string]$Name
    )
    try {
        $o = Invoke-CimMethod -CimSession $Cim -ClassName StdRegProv -MethodName GetStringValue -Namespace root\cimv2 `
             -Arguments @{ hDefKey=$Hive; sSubKeyName=$Path; sValueName=$Name } -ErrorAction Stop
        return $o.sValue
    } catch { return $null }
}

function Get-FileVersionRemote {
    param(
        [Microsoft.Management.Infrastructure.CimSession]$Cim,
        [string]$FullPath
    )
    try {
        if ([string]::IsNullOrWhiteSpace($FullPath)) { return $null }
        $p = $FullPath.Trim()
        if ($p.StartsWith('"')) {
            if ($p -match '^"(.+?)"') { $p = $matches[1] }
        } else {
            $p = $p.Split(' ')[0]
        }
        if ([string]::IsNullOrWhiteSpace($p)) { return $null }
        $wql = $p -replace '\\','\\\\'
        $f = Get-CimInstance -Class CIM_DataFile -Filter ("Name='{0}'" -f $wql) -CimSession $Cim -ErrorAction SilentlyContinue
        if ($f) { return $f.Version } else { return $null }
    } catch { return $null }
}

function Get-ExeVersionFromFolders {
    [CmdletBinding()]
    param(
        [Microsoft.Management.Infrastructure.CimSession]$Cim,
        [string[]]$Folders,
        [string[]]$Preferred = @()
    )

    foreach ($folder in $Folders) {
        if ([string]::IsNullOrWhiteSpace($folder)) { continue }
        $dir = $folder.TrimEnd('\')
        if ($dir -notmatch '^[A-Za-z]:\\') { continue }

        $drive = $dir.Substring(0,1) + ':'
        $path  = $dir.Substring(2).Replace('\','\\') + '\\'

        foreach ($exe in $Preferred) {
            try {
                $f = Get-CimInstance -Class CIM_DataFile -CimSession $Cim `
                     -Filter ("Drive='{0}' AND Path='{1}' AND FileName='{2}' AND Extension='exe'" -f $drive,$path,([IO.Path]::GetFileNameWithoutExtension($exe))) `
                     -ErrorAction SilentlyContinue
                if ($f -and $f.Version) { return $f.Version }
            } catch {}
        }

        try {
            $files = Get-CimInstance -Class CIM_DataFile -CimSession $Cim `
                     -Filter ("Drive='{0}' AND Path='{1}' AND Extension='exe'" -f $drive,$path) -ErrorAction SilentlyContinue
            $v = $files | Where-Object { $_.Version } | Sort-Object {[version]$_.Version} -Descending | Select-Object -First 1
            if ($v) { return $v.Version }
        } catch {}
    }
    return $null
}

function Get-InstalledAppsSummary {
    param([Parameter(Mandatory)][Microsoft.Management.Infrastructure.CimSession]$Cim)

    $result = @{}
    $hklm = 2147483650
    $paths = @(
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    $targets = @(
        @{ Key='CrowdStrike'; Pattern='crowdstrike|falcon' }
        @{ Key='Cisco';       Pattern='cisco' }
        @{ Key='Office';      Pattern='office|microsoft 365' }
    )

    foreach ($path in $paths) {
        try {
            $enum = Invoke-CimMethod -CimSession $Cim -ClassName StdRegProv -MethodName EnumKey -Namespace root\cimv2 `
                    -Arguments @{ hDefKey=$hklm; sSubKeyName=$path } -ErrorAction SilentlyContinue
        } catch { continue }

        if ($enum.sNames) {
            foreach ($sub in $enum.sNames) {
                $subPath = "$path\$sub"

                try {
                    $dnObj = Invoke-CimMethod -CimSession $Cim -ClassName StdRegProv -MethodName GetStringValue -Namespace root\cimv2 `
                              -Arguments @{ hDefKey=$hklm; sSubKeyName=$subPath; sValueName='DisplayName' } -ErrorAction SilentlyContinue
                    $dn = $dnObj.sValue
                } catch { continue }
                if ([string]::IsNullOrWhiteSpace($dn)) { continue }

                $dv = $null
                try {
                    $dvObj = Invoke-CimMethod -CimSession $Cim -ClassName StdRegProv -MethodName GetStringValue -Namespace root\cimv2 `
                              -Arguments @{ hDefKey=$hklm; sSubKeyName=$subPath; sValueName='DisplayVersion' } -ErrorAction SilentlyContinue
                    $dv = $dvObj.sValue
                } catch {}

                foreach ($t in $targets) {
                    if (-not $result.ContainsKey($t.Key) -and $dn -match $t.Pattern) {
                        $val = if ($dv) { "$dn (version $dv)" } else { $dn }
                        $result[$t.Key] = $val
                    }
                }
            }
        }
    }
    return $result
}

function New-CimSessionSmart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [pscredential]$Credential
    )
    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
        throw "Host '$ComputerName' is not responding (ping failed)."
    }

    $wsmanOpt = New-CimSessionOption -Protocol Wsman
    try {
        $p = @{ ComputerName=$ComputerName; SessionOption=$wsmanOpt; ErrorAction='Stop' }
        if ($Credential) { $p.Credential = $Credential }
        return (New-CimSession @p)
    } catch {
        Write-Info "WSMan is not available ($($_.Exception.Message)). Trying DCOM..." "Yellow"
        $dcomOpt = New-CimSessionOption -Protocol Dcom
        $p = @{ ComputerName=$ComputerName; SessionOption=$dcomOpt; ErrorAction='Stop' }
        if ($Credential) { $p.Credential = $Credential }
        return (New-CimSession @p)
    }
}

function Get-DockInfoSummary {
    param(
        [Parameter(Mandatory)][Microsoft.Management.Infrastructure.CimSession]$Cim,
        [switch]$Detailed
    )

    $out = @()
    $pnp = $null
    try { $pnp = Get-CimInstance -Class Win32_PnPEntity -CimSession $Cim -ErrorAction SilentlyContinue } catch {}

    $vidMap = @{
        'VID_17E9'='DisplayLink'; 'VID_413C'='Dell'; 'VID_17EF'='Lenovo';
        'VID_03F0'='HP'; 'VID_103C'='HP'
    }

    $modelMap = @(
        @{ Pattern='(40ay|40aj|40an)'; Model='Lenovo ThinkPad Universal USB-C Dock' }
        @{ Pattern='thinkpad\s+universal.*usb[-\s]*c.*dock'; Model='Lenovo ThinkPad Universal USB-C Dock' }
        @{ Pattern='universal.*usb[-\s]*c.*dock'; Model='Lenovo ThinkPad Universal USB-C Dock' }
        @{ Pattern='dock\s*gen\s*2'; Model='Lenovo ThinkPad USB-C Dock Gen 2' }

        @{ Pattern='wd19|wd19s|wd22|wd19tb'; Model='Dell WD19 / WD22' }
        @{ Pattern='d6000'; Model='Dell D6000' }
        @{ Pattern='thinkpad.*(usb[-\s]*c|thunderbolt).*dock'; Model='Lenovo ThinkPad USB-C / Thunderbolt Dock' }
        @{ Pattern='lenovo.*(usb[-\s]*c|thunderbolt).*dock'; Model='Lenovo USB-C / Thunderbolt Dock' }
        @{ Pattern='hp.*(usb[-\s]*c|thunderbolt).*dock'; Model='HP USB-C / Thunderbolt Dock' }
        @{ Pattern='displaylink.*dock'; Model='DisplayLink-based Dock' }
        @{ Pattern='ultra\s*dock'; Model='Ultra Dock' }
        @{ Pattern='billboard device'; Model='USB-C Alt-Mode Billboard (fallback)' }
    )

    $best = $null
    $brandHit = $null
    $dlHint = $false

    if ($pnp) {
        foreach ($dev in $pnp) {
            $name = [string]$dev.Name
            $desc = [string]$dev.Description
            $mfr = [string]$dev.Manufacturer
            $id = [string]$dev.PNPDeviceID

            if ($id -like '*VID_17E9*' -or $name -match 'DisplayLink' -or $mfr -match 'DisplayLink') { $dlHint = $true }
            foreach ($k in $vidMap.Keys) { if ($id -like "*$k*") { $brandHit = $vidMap[$k] } }

            foreach ($rule in $modelMap) {
                if ($name -match $rule.Pattern -or $desc -match $rule.Pattern -or $mfr -match $rule.Pattern -or $id -match $rule.Pattern) {
                    $best = $rule.Model; break
                }
            }
            if ($best) { break }
        }
    }

    if (-not $best) {
        try {
            $roots = @('HKLM:\SYSTEM\CurrentControlSet\Enum\USB','HKLM:\SYSTEM\CurrentControlSet\Enum\PCI')
            foreach ($root in $roots) {
                $keys = Invoke-CimMethod -CimSession $Cim -ClassName StdRegProv -MethodName EnumKey -Namespace root\cimv2 `
                        -Arguments @{ hDefKey=2147483650; sSubKeyName=$root.Substring(6) } -ErrorAction SilentlyContinue
                if ($keys.sNames) {
                    foreach ($k in $keys.sNames) {
                        if ($k -match '40AY|UNIVERSAL.*USB.?C.*DOCK|DOCK.?GEN.?2') { $best = 'Lenovo ThinkPad Universal USB-C Dock'; break }
                    }
                }
                if ($best) { break }
            }
        } catch {}
    }

    if ($best) {
        $out += [pscustomobject]@{ Section='Dock'; Property='Model'; Value=$best }
    } elseif ($brandHit) {
        $out += [pscustomobject]@{ Section='Dock'; Property='Model'; Value=("$brandHit dock/replicator (detected by VID/PID)") }
    } elseif ($dlHint) {
        $out += [pscustomobject]@{ Section='Dock'; Property='Model'; Value='DisplayLink-based dock' }
    } else {
        $out += [pscustomobject]@{ Section='Dock'; Property='Model'; Value='No data' }
    }

    return $out
}

function Probe-CrowdStrike {
    param(
        [Microsoft.Management.Infrastructure.CimSession]$Cim,
        [hashtable]$Apps
    )

    $out = [ordered]@{ Product='CrowdStrike'; Present=$false; Version=$null; Source=$null }

    $svc = Get-CimInstance Win32_Service -CimSession $Cim `
           -Filter "Name='CSAgent' OR Name='CSFalconService' OR DisplayName LIKE '%CrowdStrike%'" `
           -ErrorAction SilentlyContinue

    if ($svc) {
        $out.Present = $true

        $ver = Get-RegStringRemote -Cim $Cim -Hive 2147483650 -Path 'SOFTWARE\CrowdStrike\Products\Sensor' -Name 'Version'
        if (-not $ver) { $ver = Get-RegStringRemote -Cim $Cim -Hive 2147483650 -Path 'SOFTWARE\CrowdStrike' -Name 'ProductVersion' }
        if (-not $ver) { $ver = Get-RegStringRemote -Cim $Cim -Hive 2147483650 -Path 'SOFTWARE\CrowdStrike' -Name 'SensorVersion' }
        if (-not $ver) { $ver = Get-RegStringRemote -Cim $Cim -Hive 2147483650 -Path 'SOFTWARE\WOW6432Node\CrowdStrike\Products\Sensor' -Name 'Version' }
        if (-not $ver) { $ver = Get-RegStringRemote -Cim $Cim -Hive 2147483650 -Path 'SOFTWARE\WOW6432Node\CrowdStrike' -Name 'ProductVersion' }

        if ($ver) {
            $out.Version = $ver
            $out.Source  = 'Registry'
            return [pscustomobject]$out
        }

        $ver = Get-FileVersionRemote -Cim $Cim -FullPath ($svc | Select-Object -First 1).PathName
        if ($ver) {
            $out.Version = $ver
            $out.Source  = 'ServicePath'
            return [pscustomobject]$out
        }
    }
    else {
        $ver = Get-RegStringRemote -Cim $Cim -Hive 2147483650 -Path 'SOFTWARE\CrowdStrike\Products\Sensor' -Name 'Version'
        if ($ver) {
            $out.Present = $true
            $out.Version = $ver
            $out.Source  = 'RegistryOnly'
            return [pscustomobject]$out
        }
    }

    $ver = Get-ExeVersionFromFolders -Cim $Cim `
           -Folders @('C:\Program Files\CrowdStrike\','C:\Program Files (x86)\CrowdStrike\') `
           -Preferred @('CSAgent.exe','CSFalconService.exe')
    if ($ver) {
        $out.Present = $true
        $out.Version = $ver
        $out.Source  = 'FolderScan'
    }

    if (-not $out.Version -and $Apps -and $Apps.ContainsKey('CrowdStrike')) {
        $out.Present = $true
        if ($Apps['CrowdStrike'] -match 'version\s+([0-9\.,]+)') {
            $out.Version = $matches[1]
            if (-not $out.Source) { $out.Source = 'UninstallDisplayName' }
        }
    }

    return [pscustomobject]$out
}

function Probe-CiscoSecureClient {
    param(
        [Microsoft.Management.Infrastructure.CimSession]$Cim,
        [hashtable]$Apps
    )

    $out = [ordered]@{ Product='Cisco Secure Client'; Present=$false; Version=$null; Source=$null }

    $svc = Get-CimInstance Win32_Service -CimSession $Cim `
           -Filter "Name LIKE 'vpnagent%' OR Name LIKE 'cisco%' OR DisplayName LIKE '%Cisco%'" `
           -ErrorAction SilentlyContinue

    if ($svc) {
        $out.Present = $true

        $ver = Get-RegStringRemote -Cim $Cim -Hive 2147483650 -Path 'SOFTWARE\Cisco\Cisco Secure Client' -Name 'DisplayVersion'
        if (-not $ver) { $ver = Get-RegStringRemote -Cim $Cim -Hive 2147483650 -Path 'SOFTWARE\Cisco\Cisco Secure Client' -Name 'ProductVersion' }
        if (-not $ver) { $ver = Get-RegStringRemote -Cim $Cim -Hive 2147483650 -Path 'SOFTWARE\WOW6432Node\Cisco\Cisco Secure Client' -Name 'DisplayVersion' }
        if (-not $ver) { $ver = Get-RegStringRemote -Cim $Cim -Hive 2147483650 -Path 'SOFTWARE\WOW6432Node\Cisco\Cisco Secure Client' -Name 'ProductVersion' }

        if ($ver) {
            $out.Version = $ver
            $out.Source  = 'Registry'
            return [pscustomobject]$out
        }

        $s   = $svc | Sort-Object DisplayName | Select-Object -First 1
        $ver = Get-FileVersionRemote -Cim $Cim -FullPath $s.PathName
        if ($ver) {
            $out.Version = $ver
            $out.Source  = 'ServicePath'
            return [pscustomobject]$out
        }
    }
    else {
        $ver = Get-RegStringRemote -Cim $Cim -Hive 2147483650 -Path 'SOFTWARE\Cisco\Cisco Secure Client' -Name 'DisplayVersion'
        if ($ver) {
            $out.Present = $true
            $out.Version = $ver
            $out.Source  = 'RegistryOnly'
            return [pscustomobject]$out
        }
    }

    $ver = Get-ExeVersionFromFolders -Cim $Cim `
           -Folders @('C:\Program Files\Cisco\Cisco Secure Client\',
                      'C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\') `
           -Preferred @('vpnagent.exe','ciscod_tunnel.exe','vpnui.exe')
    if ($ver) {
        $out.Present = $true
        $out.Version = $ver
        $out.Source  = 'FolderScan'
    }

    if (-not $out.Version -and $Apps -and $Apps.ContainsKey('Cisco')) {
        $out.Present = $true
        if ($Apps['Cisco'] -match 'version\s+([0-9\.,]+)') {
            $out.Version = $matches[1]
            if (-not $out.Source) { $out.Source = 'UninstallDisplayName' }
        }
    }

    return [pscustomobject]$out
}

function Get-PCInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [pscredential]$Credential
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $session = $null

    try {
        Write-Info "Connecting to $ComputerName ..."
        $session = New-CimSessionSmart -ComputerName $ComputerName -Credential $Credential
        Write-Info "Connected." "Green"

        $apps = Get-InstalledAppsSummary -Cim $session

        $cs  = Get-CimInstance -Class Win32_ComputerSystem -CimSession $session -ErrorAction SilentlyContinue
        $os  = Get-CimInstance -Class Win32_OperatingSystem -CimSession $session -ErrorAction SilentlyContinue
        $bios= Get-CimInstance -Class Win32_BIOS -CimSession $session -ErrorAction SilentlyContinue
        $cpu = Get-CimInstance -Class Win32_Processor -CimSession $session -ErrorAction SilentlyContinue
        $gpu = Get-CimInstance -Class Win32_VideoController -CimSession $session -ErrorAction SilentlyContinue
        $mem = Get-CimInstance -Class Win32_PhysicalMemory -CimSession $session -ErrorAction SilentlyContinue
        $ld  = Get-CimInstance -Class Win32_LogicalDisk -Filter "DriveType=3" -CimSession $session -ErrorAction SilentlyContinue
        $net = Get-CimInstance -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled=TRUE" -CimSession $session -ErrorAction SilentlyContinue
        $bat = Get-CimInstance -Class Win32_Battery -CimSession $session -ErrorAction SilentlyContinue

        if ($cs) {
            Add-Row ([ref]$rows) "System" "Name" $cs.Name
            Add-Row ([ref]$rows) "System" "Model" $cs.Model
            Add-Row ([ref]$rows) "System" "Manufacturer" $cs.Manufacturer
            Add-Row ([ref]$rows) "System" "Logged-on user" $cs.UserName
            Add-Row ([ref]$rows) "System" "Domain/Workgroup" $cs.Domain
            Add-Row ([ref]$rows) "System" "RAM (GB)" ([math]::Round(($cs.TotalPhysicalMemory/1GB),2))
            try {
                if ($cs.PCSystemType) {
                    $map = @{
                        1='Desktop'; 2='Mobile/Laptop'; 3='Workstation'; 4='Enterprise Server';
                        5='SOHO Server'; 6='Appliance PC'; 7='Performance Server'; 8='Maximum'
                    }
                    $val = $map[[int]$cs.PCSystemType]
                    if (-not $val) { $val = $cs.PCSystemType }
                    Add-Row ([ref]$rows) "System" "System type" $val
                }
            } catch {}
        }

        if ($os) {
            Add-Row ([ref]$rows) "OS" "Name" ("{0} {1}" -f $os.Caption,$os.OSArchitecture)
            Add-Row ([ref]$rows) "OS" "Version" $os.Version
            Add-Row ([ref]$rows) "OS" "Build" $os.BuildNumber

            $lastBoot = Convert-WmiDateSafe $os.LastBootUpTime
            if ($lastBoot) {
                Add-Row ([ref]$rows) "OS" "Last boot" $lastBoot
                Add-Row ([ref]$rows) "OS" "Uptime (h)" ([math]::Round(((Get-Date) - $lastBoot).TotalHours,1))
            } else {
                Add-Row ([ref]$rows) "OS" "Last boot" "No data"
                Add-Row ([ref]$rows) "OS" "Uptime (h)" "No data"
            }

            $instDate = Convert-WmiDateSafe $os.InstallDate
            if ($instDate) {
                Add-Row ([ref]$rows) "OS" "Install date" $instDate
            } else {
                Add-Row ([ref]$rows) "OS" "Install date" "No data"
            }
        }

        if ($bios) {
            Add-Row ([ref]$rows) "BIOS" "Serial number" $bios.SerialNumber
            Add-Row ([ref]$rows) "BIOS" "Version" $bios.SMBIOSBIOSVersion
            $biosDate = Convert-WmiDateSafe $bios.ReleaseDate
            if ($biosDate) { Add-Row ([ref]$rows) "BIOS" "Release date" $biosDate } else { Add-Row ([ref]$rows) "BIOS" "Release date" "No data" }
        }

        if ($cpu) {
            Add-Row ([ref]$rows) "CPU" "Processor" $cpu.Name
            try { Add-Row ([ref]$rows) "CPU" "Cores/Logical" ("{0}/{1}" -f $cpu.NumberOfCores,$cpu.NumberOfLogicalProcessors) } catch {}
        }

        if ($gpu) { $gpu | ForEach-Object { Add-Row ([ref]$rows) "GPU" "Adapter" $_.Name } }

        if ($mem) {
            $memCount = ($mem | Measure-Object).Count
            $memGB = [math]::Round(($mem | Measure-Object -Property Capacity -Sum).Sum/1GB,2)
            Add-Row ([ref]$rows) "Memory" "Modules" $memCount
            Add-Row ([ref]$rows) "Memory" "Total (GB)" $memGB
        }

        if ($ld) {
            foreach ($d in $ld) {
                $free = if ($d.FreeSpace) { [math]::Round($d.FreeSpace/1GB,1) } else { "No data" }
                $size = if ($d.Size) { [math]::Round($d.Size/1GB,1) } else { "No data" }
                $line = ("{0} GB free of {1} GB (FS: {2})" -f $free,$size,$d.FileSystem)
                Add-Row ([ref]$rows) "Disks" "[$($d.DeviceID)] Label=$($d.VolumeName)" $line
            }
        }

        if ($net) {
            foreach ($n in $net) {
                $ips = ($n.IPAddress | Where-Object {$_ -match '^\d{1,3}(\.\d{1,3}){3}$'}) -join ', '
                $dns = ($n.DNSServerSearchOrder -join ', ')
                $line = ("IP: {0} | DNS: {1} | DHCP: {2}" -f $ips,$dns,$n.DHCPEnabled)
                Add-Row ([ref]$rows) "Network" "Adapter $($n.Description)" $line
            }
        }

        if ($bat) {
            foreach ($b in $bat) {
                Add-Row ([ref]$rows) "Battery" "Status" $b.BatteryStatus
                Add-Row ([ref]$rows) "Battery" "Chemistry" $b.Chemistry
                try {
                    if ($b.EstimatedChargeRemaining -ne $null) { Add-Row ([ref]$rows) "Battery" "Charge (%)" $b.EstimatedChargeRemaining }
                    if ($b.EstimatedRunTime -ne $null -and $b.EstimatedRunTime -gt 0 -and $b.EstimatedRunTime -lt 10000) {
                        $h = [math]::Round($b.EstimatedRunTime / 60,1)
                        Add-Row ([ref]$rows) "Battery" "Estimated runtime (h)" $h
                    }
                } catch {}
            }
        }

        try {
            $battStatic = Get-CimInstance -Namespace root\wmi -Class BatteryStaticData -CimSession $session -ErrorAction SilentlyContinue
            $battFull   = Get-CimInstance -Namespace root\wmi -Class BatteryFullChargedCapacity -CimSession $session -ErrorAction SilentlyContinue
            if ($battStatic -and $battFull) {
                $design = ($battStatic | Select-Object -First 1).DesignedCapacity
                $full   = ($battFull   | Select-Object -First 1).FullChargedCapacity
                if ($design -gt 0 -and $full -gt 0) {
                    $wear   = [math]::Round((1 - ($full / $design)) * 100,1)
                    $health = [math]::Round(($full / $design) * 100,1)
                    Add-Row ([ref]$rows) "Battery" "Design capacity (mWh)" $design
                    Add-Row ([ref]$rows) "Battery" "Full charge capacity (mWh)" $full
                    Add-Row ([ref]$rows) "Battery" "Wear (%)" $wear
                    Add-Row ([ref]$rows) "Battery" "Health vs new (%)" $health
                }
            }
            $battCycle = Get-CimInstance -Namespace root\wmi -Class BatteryCycleCount -CimSession $session -ErrorAction SilentlyContinue
            if ($battCycle) {
                $cycles = ($battCycle | Select-Object -First 1).CycleCount
                if ($cycles -ne $null) { Add-Row ([ref]$rows) "Battery" "Charge cycles" $cycles }
            }
        } catch {
            Add-Row ([ref]$rows) "Battery" "Error" $_.Exception.Message
        }

        try {
            $dockRows = Get-DockInfoSummary -Cim $session
            foreach ($r in $dockRows) { $rows.Add($r) | Out-Null }
        } catch { Add-Row ([ref]$rows) "Dock" "Error" $_.Exception.Message }

        try {
            $monID    = Get-CimInstance -Namespace root\wmi -Class WmiMonitorID                 -CimSession $session -ErrorAction SilentlyContinue
            $monBasic = Get-CimInstance -Namespace root\wmi -Class WmiMonitorBasicDisplayParams -CimSession $session -ErrorAction SilentlyContinue
            $monConn  = Get-CimInstance -Namespace root\wmi -Class WmiMonitorConnectionParams   -CimSession $session -ErrorAction SilentlyContinue

            if ($monID) {
                foreach ($m in $monID) {
                    $inst = $m.InstanceName
                    $mfr  = Convert-EdidString $m.ManufacturerName
                    $mdl  = Convert-EdidString $m.ProductCodeID
                    $ser  = Convert-EdidString $m.SerialNumberID
                    $name = Convert-EdidString $m.UserFriendlyName

                    $basic = $null; if ($monBasic) { $basic = $monBasic | Where-Object { $_.InstanceName -eq $inst } }
                    $conn  = $null; if ($monConn)  { $conn  = $monConn  | Where-Object { $_.InstanceName -eq $inst } }

                    $sizeText = "No data"
                    if ($basic) {
                        $h = [int]$basic.MaxVerticalImageSize
                        $w = [int]$basic.MaxHorizontalImageSize
                        if ($h -gt 0 -and $w -gt 0) {
                            $diag = [math]::Round(( [math]::Sqrt($h*$h + $w*$w) / 2.54 ), 1)
                            $sizeText = ("{0}x{1} cm (~{2}'')" -f $w,$h,$diag)
                        }
                    }

                    $connText = "No data"
                    if ($conn) {
                        $map = @{ 0='HD15'; 1='DVI'; 2='HDMI'; 3='LVDS/eDP'; 4='DisplayPort'; 5='Composite'; 6='SVideo'; 7='Component'; 8='Internal' }
                        $ct = $conn.VideoOutputTechnology
                        if ($map.ContainsKey($ct)) { $connText = $map[$ct] } else { $connText = [string]$ct }
                    }

                    $model = $name
                    if ([string]::IsNullOrWhiteSpace($model) -or $model -eq 'No data') { $model = $mdl }
                    if ([string]::IsNullOrWhiteSpace($model)) { $model = "No data" }

                    $line = ("Mfr: {0} | Model: {1} | S/N: {2} | Size: {3} | Connection: {4}" -f $mfr,$model,$ser,$sizeText,$connText)
                    Add-Row ([ref]$rows) "Monitors" "Monitor" $line
                }
            } else {
                $wm = Get-CimInstance -Class Win32_DesktopMonitor -CimSession $session -ErrorAction SilentlyContinue
                if ($wm) {
                    foreach ($mm in $wm) {
                        $line = ("Name: {0} | PNPDeviceID: {1} | Status: {2}" -f $mm.Name,$mm.PNPDeviceID,$mm.Status)
                        Add-Row ([ref]$rows) "Monitors" "Monitor" $line
                    }
               } else {
                    Add-Row ([ref]$rows) "Monitors" "Monitor" "No data"
                }
            }
        } catch { Add-Row ([ref]$rows) "Monitors" "Error" $_.Exception.Message }

        try {
            if (Get-Module -ListAvailable -Name ActiveDirectory) {
                Import-Module ActiveDirectory -ErrorAction Stop
                $adObject = Get-ADComputer -Identity $ComputerName -Properties DistinguishedName,MemberOf -ErrorAction Stop
                Add-Row ([ref]$rows) "Active Directory" "DistinguishedName" $adObject.DistinguishedName
                if ($adObject.MemberOf) {
                    $adObject.MemberOf |
                        ForEach-Object { $_.Split(',')[0] -replace '^CN=', '' } |
                        Sort-Object |
                        ForEach-Object { Add-Row ([ref]$rows) "Active Directory" "Group" $_ }
                } else {
                    Add-Row ([ref]$rows) "Active Directory" "Groups" "No data"
                }
            } else {
                Add-Row ([ref]$rows) "Active Directory" "Info" "RSAT-AD module is not installed"
            }
        } catch { Add-Row ([ref]$rows) "Active Directory" "Error" $_.Exception.Message }

        try {
            $tcp = Get-CimInstance -Namespace root/StandardCimv2 -ClassName MSFT_NetTCPConnection -CimSession $session -ErrorAction SilentlyContinue
            if ($tcp) {
                $listening = $tcp | Where-Object { $_.State -eq 2 -or $_.State -eq 'Listen' }
                $topListen = $listening | Group-Object LocalPort | Sort-Object Count -Descending | Select-Object -First 10
                foreach ($p in $topListen) { Add-Row ([ref]$rows) "Network ports" "Listening" ("Port {0}/TCP - {1} endpoints" -f $p.Name,$p.Count) }

                $established = $tcp | Where-Object { $_.State -eq 5 -or $_.State -eq 'Established' } |
                               Select-Object -Property LocalAddress,LocalPort,RemoteAddress,RemotePort -First 10
                foreach ($c in $established) {
                    Add-Row ([ref]$rows) "Network ports" "Connection" ("{0}:{1} -> {2}:{3}" -f $c.LocalAddress,$c.LocalPort,$c.RemoteAddress,$c.RemotePort)
                }
            } else {
                Add-Row ([ref]$rows) "Network ports" "Info" "No data (MSFT_NetTCPConnection)"
            }
        } catch { Add-Row ([ref]$rows) "Network ports" "Error" $_.Exception.Message }

        try {
            $since = (Get-Date).AddDays(-7)
            $sysErr = Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{LogName='System'; Level=2; StartTime=$since} -MaxEvents 5 -ErrorAction SilentlyContinue
            $appErr = Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{LogName='Application'; Level=2; StartTime=$since} -MaxEvents 5 -ErrorAction SilentlyContinue

            if ($sysErr) {
                Add-Row ([ref]$rows) "Events" "System errors (7 days)" ("Count: {0}" -f ($sysErr | Measure-Object).Count)
                foreach ($e in $sysErr) { Add-Row ([ref]$rows) "Events" "System" ("{0} | ID={1} | {2}" -f $e.TimeCreated,$e.Id,$e.ProviderName) }
            } else { Add-Row ([ref]$rows) "Events" "System" "No critical System errors (last 7 days)" }

            if ($appErr) {
                Add-Row ([ref]$rows) "Events" "Application errors (7 days)" ("Count: {0}" -f ($appErr | Measure-Object).Count)
                foreach ($e in $appErr) { Add-Row ([ref]$rows) "Events" "Application" ("{0} | ID={1} | {2}" -f $e.TimeCreated,$e.Id,$e.ProviderName) }
            } else { Add-Row ([ref]$rows) "Events" "Application" "No critical Application errors (last 7 days)" }
        } catch { Add-Row ([ref]$rows) "Events" "Error" $_.Exception.Message }

        try {
            $qfe = Get-CimInstance -Class Win32_QuickFixEngineering -CimSession $session -ErrorAction SilentlyContinue
            if ($qfe) {
                $last = $qfe | Sort-Object -Property InstalledOn -Descending | Select-Object -First 1
                if ($last) {
                    Add-Row ([ref]$rows) "Updates" "Last hotfix (QFE)" ("{0} | KB: {1} | Description: {2}" -f $last.InstalledOn,$last.HotFixID,$last.Description)
                    Add-Row ([ref]$rows) "Updates" "QFE entries count" (($qfe | Measure-Object).Count)
                }
            } else {
                Add-Row ([ref]$rows) "Updates" "Info" "No hotfix info (Win32_QuickFixEngineering)"
            }
        } catch { Add-Row ([ref]$rows) "Updates" "Error" $_.Exception.Message }

        try {
            $wu = Get-WinEvent -ComputerName $ComputerName -LogName 'Microsoft-Windows-WindowsUpdateClient/Operational' `
                 -FilterXPath "*[System[TimeCreated[timediff(@SystemTime) <= 2592000000]] and System[EventID=19]]" `
                 -MaxEvents 20 -ErrorAction SilentlyContinue
            if ($wu) {
                $i = 0
                foreach ($e in $wu) {
                    $xml = [xml]$e.ToXml()
                    $titleNode = $xml.Event.EventData.Data | Where-Object { $_.Name -eq 'UpdateTitle' }
                    $title = if ($titleNode) { $titleNode.'#text' } else { $e.Message.Split("`n")[0] }
                    Add-Row ([ref]$rows) "Updates" "WU (successful)" ("{0} | {1}" -f $e.TimeCreated,$title)
                    $i++; if ($i -ge 5) { break }
                }
            } else {
                Add-Row ([ref]$rows) "Updates" "WU (successful)" "No entries in last 30 days or no access to log"
            }
        } catch { Add-Row ([ref]$rows) "Updates" "WU (successful)" $_.Exception.Message }

        try {
            $csInfo    = Probe-CrowdStrike -Cim $session -Apps $apps
            $ciscoInfo = Probe-CiscoSecureClient -Cim $session -Apps $apps

            if ($csInfo.Present) {
                $txt = if ($csInfo.Version) {
                    "Installed (version $($csInfo.Version), source: $($csInfo.Source))"
                } else {
                    "Installed (version: no data, source: $($csInfo.Source))"
                }
                Add-Row ([ref]$rows) "Security" "CrowdStrike" $txt
            } else {
                Add-Row ([ref]$rows) "Security" "CrowdStrike" "NOT detected (no CrowdStrike services/keys/files)"
            }

            if ($ciscoInfo.Present) {
                $txt = if ($ciscoInfo.Version) {
                    "Installed (version $($ciscoInfo.Version), source: $($ciscoInfo.Source))"
                } else {
                    "Installed (version: no data, source: $($ciscoInfo.Source))"
                }
                Add-Row ([ref]$rows) "Security" "Cisco Secure Client" $txt
            } else {
                Add-Row ([ref]$rows) "Security" "Cisco Secure Client" "NOT detected (no Cisco Secure Client services/keys/files)"
            }
        }
        catch {
            Add-Row ([ref]$rows) "Security" "Error" $_.Exception.Message
        }

        try {
            $officeInfo = $null
            if ($apps -and $apps.ContainsKey('Office')) { $officeInfo = $apps['Office'] }

            $excelPath = Get-RegStringRemote -Cim $session -Hive 2147483650 `
                         -Path 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE' -Name 'Path'
            if ($excelPath) {
                $exe = Join-Path $excelPath 'EXCEL.EXE'
                $excelVer = Get-FileVersionRemote -Cim $session -FullPath $exe
                if ($excelVer) {
                    $line = "Excel version {0} ({1})" -f $excelVer,$exe
                    if ($officeInfo) { $officeInfo = "$officeInfo; $line" } else { $officeInfo = $line }
                }
            }
            $ctrVer = Get-RegStringRemote -Cim $session -Hive 2147483650 -Path 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name 'ClientVersionToReport'
            $ctrIds = Get-RegStringRemote -Cim $session -Hive 2147483650 -Path 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name 'ProductReleaseIds'
            if ($ctrVer) {
                $line = if ($ctrIds) { "ClickToRun: {0} ({1})" -f $ctrVer,$ctrIds } else { "ClickToRun: {0}" -f $ctrVer }
                if ($officeInfo) { $officeInfo = "$officeInfo; $line" } else { $officeInfo = $line }
            }
            if ($officeInfo) {
                Add-Row ([ref]$rows) "Software" "Microsoft Office / Excel" $officeInfo
            } else {
                Add-Row ([ref]$rows) "Software" "Microsoft Office / Excel" "Not found (no Office, AppPaths or ClickToRun entries)"
            }
        } catch { Add-Row ([ref]$rows) "Software" "Error" $_.Exception.Message }

    } catch {
        Add-Row ([ref]$rows) "Error" "Computer" "$ComputerName is not reachable"
        Add-Row ([ref]$rows) "Error" "Details" ($_.Exception.Message)
    }

    if ($session) { $session | Remove-CimSession }
    return $rows
}

if ([string]::IsNullOrWhiteSpace($ComputerName)) {
    Write-Host "Computer name is not specified!" -ForegroundColor Red
    return
}

$info = Get-PCInfo -ComputerName $ComputerName -Credential $Credential

Write-Host "`nInformation summary:`n" -ForegroundColor Cyan
$info | Group-Object Section | ForEach-Object {
    Write-Host ($_.Name + ":") -ForegroundColor Yellow
    $_.Group | ForEach-Object {
        Write-Host (" {0}: {1}" -f $_.Property,$_.Value) -ForegroundColor White
    }
}

if ($Export) {
    if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir | Out-Null }
    $ts = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $csv = Join-Path $ExportDir ("PCInfo_{0}_{1}.csv" -f $ComputerName,$ts)
    $json= Join-Path $ExportDir ("PCInfo_{0}_{1}.json" -f $ComputerName,$ts)
    $info | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    $info | ConvertTo-Json -Depth 5 | Out-File -FilePath $json -Encoding UTF8
    Write-Info "Saved: $csv`nSaved: $json" "Green"
}

Write-Host "`nPress Enter to exit..." -ForegroundColor Green
[void](Read-Host)
