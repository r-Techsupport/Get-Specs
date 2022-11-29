
# ------------------ #
promptStart
$timer = [diagnostics.stopwatch]::StartNew()
# ------------------ #

# ------------------ #
$today = Get-Date
$hostsSum = $(Get-FileHash $hostsFile).hash
$hostsContent = Get-Content $hostsFile
# ------------------ #

# Bulk data gathering
Write-Host 'Gathering main data...'
## CIM sources
$cimOs = Get-CimInstance -ClassName Win32_OperatingSystem
$cimStart = Get-CimInstance Win32_StartupCommand
$taskStart = Get-ScheduledTask -TaskName * | ? { $_.Triggers -Like "MSFT_TaskBootTrigger" -Or $_.Triggers -Like "MSFT_TaskLogonTrigger" } | ? { $_.State -eq "Ready" -Or $_.State -eq "Running" } | ? { $_.Author -NotLike "Microsof*" } | Select TaskName
$startUps = $cimStart.caption + $taskStart.TaskName
$cimAudio= Get-CimInstance win32_sounddevice | Select Name,ProductName
$cimLics = Get-CimInstance -ClassName SoftwareLicensingProduct | ? { $_.PartialProductKey -ne $null } | Select Name,ProductKeyChannel,LicenseFamily,LicenseStatus,PartialProductKey
$av = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct
$fw = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName FirewallProduct
$tpm = Get-CimInstance -Namespace root/cimv2/Security/MicrosoftTpm -ClassName win32_tpm
$ramSticks = Get-WmiObject win32_physicalmemory
$cimBios = Get-CimInstance Win32_bios
$powerProfiles = Get-CimInstance -N root\cimv2\power -Class win32_PowerPlan 
$monitors = Get-WmiObject WmiMonitorID -Namespace root\wmi -ErrorAction SilentlyContinue
$cimVids = Get-CimInstance -Class win32_PnpEntity -Filter "PNPClass = 'Display'"

## Other
$bootupState = $(gwmi win32_computersystem -Property BootupState).BootupState
$key = $(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' | Select-Object -ExpandProperty 'DigitalProductId')
$installedBase0 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | ? {$_.DisplayName -notlike $null}
$installedBase1 = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | ? {$_.DisplayName -notlike $null}
$installedBase = $installedBase0 + $installedBase1
$services = $(Get-Service)
$runningProcesses = Get-Process
$volumes = Get-Volume
$dns = Get-DnsClientGlobalSetting
$netAdapters = Get-NetADapter
$pnpDevices = Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue
$issueDevices = $pnpDevices | ? { $_.Status -Match "Error|Warning|Degraded|Unknown" }

# janky check for msconfig core setting
$bcdedit = bcdedit | Select-String numproc

# secureboot
$secureBoot = Confirm-SecureBootUEFI -ErrorAction SilentlyContinue

# legacy/uefi
$bootMode = $env:firmware_type

## Build info
$i = 0
$eol = $false
$eolON = "Unknown"
$osBuild = 'Version is unknown. Build: ' + $cimOS.BuildNumber
foreach ($b in $builds) {
    if ($cimOS.BuildNumber -eq $builds[$i]) {
        $osBuild = 'Version: ' + $versions[$i]
        if ((Get-Date $eolDates[$i]) -lt $today) {
            $eol = $true
            $eolOn = $eolDates[$i]
        }
    }
    $i = $i + 1
}
Write-Host 'Got main data' -ForegroundColor Green

# get smart now so we can parse it for bad things
Write-Host 'Getting SMART data...'
$smart = .\Get-Smart\Get-Smart.ps1 -cdiPath '.\DiskInfo64.exe'
Write-Host 'Got SMART' -ForegroundColor Green

# chrome extensions
$chromeExtensions = @()
$rawChromeExt = Get-ChildItem -Recurse "$env:USERPROFILE\AppData\Local\Google\Chrome\User Data\Default\Extensions" | ? {$_.Name -Like "manifest.json" }
ForEach ($ext in $rawChromeext) { 
    $chromeExtensions += Get-Content $ext.FullName | ConvertFrom-Json 
}
