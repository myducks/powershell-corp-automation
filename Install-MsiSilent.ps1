[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({ Test-Path $_ })]
    [string]$Path,

    [string]$LogDirectory = "$PSScriptRoot\Logs",

    [string]$AdditionalArgs = ""
)

try {
    if (-not (Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    $msiFullPath = (Resolve-Path $Path).Path

    if ([IO.Path]::GetExtension($msiFullPath) -ne ".msi") {
        throw "File '$msiFullPath' is not an MSI."
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logFile = Join-Path $LogDirectory ("{0}_{1}.log" -f [IO.Path]::GetFileNameWithoutExtension($msiFullPath), $timestamp)

    $arguments = "/i `"$msiFullPath`" /qn /norestart /L*v `"$logFile`" $AdditionalArgs"

    Write-Host "Running: msiexec.exe $arguments"
    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -PassThru

    if ($process.ExitCode -eq 0) {
        Write-Host "Installation completed successfully." -ForegroundColor Green
    }
    else {
        Write-Warning "Installer exited with code $($process.ExitCode). Check log: $logFile"
    }
}
catch {
    Write-Error "Installation failed: $_"
}
