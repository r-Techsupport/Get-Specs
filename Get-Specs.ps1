<#
.SYNOPSIS
 Gather and upload specifications of a Windows host to rTechsupport
.DESCRIPTION
  Use various native powershell and wmic functions to gather verbose information on a system to assist troubleshooting
.OUTPUTS Specs
  '.\TechSupport_Specs.html'
#>

# Run the basic checks before doing anything
function promptFileIssue {
    If (Test-Path ./files) {
        } Else {
            msg.exe * "The files directory is missing. This most likely means you did not extract the zip file properly. 
            
Right click the zip file and choose 'Extract All' then run the new exe."
    }
}
promptFileIssue

# set execution policy to allow all our children to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# source our other ps1 files
. files\wpf.ps1

# Declarations
$file = 'TechSupport_Specs.html'
$badSoftware = @(
    'Driver Booster*',
    'iTop*',
    'Driver Easy*',
    'Roblox*',
    'ccleaner',
    'Malwarebytes'
)
$badStartup = @(
    'AutoKMS',
    'kmspico',
    'McAfee Remediation',
    'KMS_VL_ALL'
)
$badProcesses = @(
    'MBAMService',
    'McAfee WebAdvisor',
    'Norton Security',
    'Wallpaper Engine Service',
    'Service_KMS.exe',
    'iTopVPN'
)
# YOU MUST MATCH THE KEY AND VALUE BELOW TO THE SAME ARRAY VALUE
$badKeys = @(
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PreviewBuilds\',
    'HKLM:\SYSTEM\Setup\LabConfig\',
    'HKLM:\SYSTEM\Setup\LabConfig\',
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\',
    'HKLM:\SYSTEM\Setup\Status\'
)
$badValues = @(
    'AllowBuildPreview',
    'BypassTPMCheck',
    'BypassSecureBootCheck',
    'NoAutoUpdate',
    'AuditBoot'
)
$badData = @(
    '1',
    '1',
    '1',
    '1',
    '1'
)
$badRegExp = @(
    'Insider builds are set to: ',
    'Windows 11 TPM bypass set to: ',
    'Windows 11 SecureBoot bypass set to: ',
    'Windows auto update is set to: ',
    'Audit Boot is set to: '
)
$badHostnames = @(
    'ATLASOS-DESKTOP',
    'Revision-PC'
)
$builds = @(
    '10240',
    '10586',
    '14393',
    '15063',
    '16299',
    '17017',
    '17134',
    '17763',
    '18362',
    '18363',
    '19041',
    '19042',
    '19043',
    '19044'
)
$versions = @(
    '1507',
    '1511',
    '1607',
    '1703',
    '1709',
    '1803',
    '1809',
    '1903',
    '1909',
    '2004',
    '20H2',
    '21H1',
    '21H2'
)
$eolDates = @(
    '2017-05-17',
    '2017-10-10',
    '2018-04-10',
    '2018-10-09',
    '2019-04-19',
    '2019-11-12',
    '2020-11-10',
    '2020-10-08',
    '2021-05-11',
    '2021-12-14',
    '2022-05-10',
    '2022-12-13',
    '2023-10-10'
)

$masKeys = Get-Content .\files\keys

### Functions

function header {
$1 = "<!DOCTYPE html>
<html>
<head>
<style>
body {
    background-color: #383c4a;
    color: White;
    margin-left: 30px;
}
h2 {
    color: #87ab63;
}
a:link {
    color: White;
}
a:visited {
    color: White;
}
a:hover{
    color: #87ab63;
}
table {
  font-family: arial, sans-serif;
  font-size: 12px;
  border-collapse: collapse;
}

td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}

tr:nth-child(even) {
  background-color: #2A2E3A;
}
</style>
<body>
<pre>
"
Return $1
}
function table {
$1 = '<h2>Sections</h2>
<div style="line-height:0">
<p><a href="#Lics">Licensing</a></p>
<p><a href="#SecInfo">Security Information</a></p>
<p><a href="#hw">Hardware Basics</a></p>
<p><a href="#SysVar">System Variables</a></p>
<p><a href="#UserVar">User Variables</a></p>
<p><a href="#hotfixes">Installed updates</a></p>
<p><a href="#StartupTasks">Startup Tasks</a></p>
<p><a href="#Power">Powerprofiles</a></p>
<p><a href="#RunningProcs">Running Processes</a></p>
<p><a href="#Services">Services</a></p>
<p><a href="#InstalledApps">Installed Applications</a></p>
<p><a href="#NetConfig">Network Configuration</a></p>
<p><a href="#Drivers">Drivers and device versions</a></p>
<p><a href="#Audio">Audio Devices</a></p>
<p><a href="#Disks">Disk Layouts</a></p>
<p><a href="#SMART">SMART</a></p>
</div>'
Return $1
}

function getDate {
    Get-Date
}
function getbasicInfo {
    Write-Host 'Getting basic information...'
    $1 = '<h2>System Information</h2>'
    $bootuptime = $cimOs.LastBootUpTime
    $uptime = $(Get-Date) - $bootuptime
    $2 = 'Edition: ' + $cimOs.Caption 
    $3 = $osBuild
    $4 = 'Install date: ' + $cimOs.InstallDate
    $5 = 'Uptime: ' + $uptime.Days + " Days " + $uptime.Hours + " Hours " +  $uptime.Minutes + " Minutes"
    $6 = 'Hostname: ' + $cimOs.CSName
    $7 = 'Domain: ' + $env:USERDOMAIN
    $8 = 'Boot mode: ' + $env:firmware_type
    $9 = 'Boot state: ' + $bootupState
    Write-Host 'Got basic information' -ForegroundColor Green
    Return $1,$2,$3,$4,$5,$6,$7,$8,$9
}
function getFullKey {
    $keyOutput=""
    $keyOffset = 52

    $IsWin10 = ([System.Math]::Truncate($key[66] / 6)) -band 1
    $key[66] = ($key[66] -band 0xF7) -bor (($IsWin10 -band 2) * 4)
    $i = 24
    $maps = "BCDFGHJKMPQRTVWXY2346789"
    do {
        $current= 0
        $j = 14
        do {
            $current = $current* 256
            $current = $key[$j + $keyOffset] + $current
            $key[$j + $keyOffset] = [System.Math]::Truncate($current / 24 )
            $current=$current % 24
            $j--
        }
        while ($j -ge 0)
            $i--
            $keyOutput = $maps.Substring($current, 1) + $keyOutput
            $last = $current
    } while ($i -ge 0)

    if ($IsWin10 -eq 1) {
        $keypart1 = $keyOutput.Substring(1, $last)
        $keypart2 = $keyOutput.Substring($last + 1, $keyOutput.Length - ($last + 1));
        $keyOutput = $keypart1 + "N" + $keypart2;
    }

    if ($keyOutput.Length -eq 25) {
        $1 = [String]::Format("{0}-{1}-{2}-{3}-{4}",
            $keyOutput.Substring(0, 5),
            $keyOutput.Substring(5, 5),
            $keyOutput.Substring(10,5),
            $keyOutput.Substring(15,5),
            $keyOutput.Substring(20,5))
    } else {
        $keyOutput
    }
    return $1
}
function getBadThings {
    Write-Host 'Checking for issues...'
    $1 = '<h2>Visible issues</h2>'
    $2 = @()
    $3 = @()
    $4 = @()
    $5 = @()
    # bad software 
    $i = 0
    $t = $badSoftware.Count
    while ($i -lt $t) {
        $2 += $(@($installedBase.DisplayName) -like $badSoftware[$i])
        $i = $i + 1
    }
    # bad startups
    foreach ($start in $badStartup) { 
        if ($cimStart.Caption -contains $start) { 
            $3 += $start 
        }
    }
    # bad processes
    foreach ($running in $badProcesses) {
        if ($runningProcesses.Name -contains $running) {
            $4 += "Process: " + $running 
        } 
    }
    # bad registry values
    $i = 0
    foreach ($reg in $badKeys) {
        $value = $null
        If (Test-Path -Path $badKeys[$i]) {
            $data = Get-ItemProperty -Path $badKeys[$i] -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $badValues[$i] -ErrorAction SilentlyContinue
        If ($data -eq $badData[$i]) {
                $5 += $badRegExp[$i] + $data
            }
        }
        $i = $i + 1
    }
    # bad hostnames
    if ($badHostnames -contains $cimOS.CSName) {
        $6 = "Modified OS: " + $cimOS.CSName
    }
    # bad diskspace
    $c = $volumes | ? { $_.DriveLetter -eq 'C' }
    $cAllowable = $c.Size - $c.Size * .20
    $cConsumed = $c.Size - $c.SizeRemaining
    If ($cConsumed -gt $cAllowable) {
        $7 = "Less than 20% left on C"
    }
    # bad productKeys
    # If ($masKeys -contains $(getFullKey)) {
        # $8 = "MAS detected"
    # }
    # bad dns suffixes
    If ('utopia.net' -in $dns.SuffixSearchList) {
        $9 = "Utopia malware, router is infected"
    }
    If ($eol) {
        $10 = "OS was EOL on " + $eolOn + " Build is " + $osBuild
    }
    # bad smart things
    $11 = @()
    Foreach ($disk in $smart) {
        If ($disk.'Reallocated Sectors Count' -gt 0) {
            $11 += "Reallocated sector on" + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Reallocated·Sectors·Count' 
        }
        If ($disk.'Current Pending Sector Count' -gt 0) {
            $11 += "Current Pending Sector Count on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Current Pending Sector Count'
        }
        If ($disk.'Uncorrectable Sector Count' -gt 0) {
            $11 += "Uncorrectable·Sector·Count on" + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'UncorrectableSector·Count'
        }
        If ($disk.'Command Timeout' -gt 0) {
            $11 += "Command Timeout on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Command Timeout'
        }
        If ($disk.'Reported Uncorrectable Errors' -gt 0) {
            $11 += "Reported Uncorrectable Errors on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Reported Uncorrectable Errors'
        }
    }
    Write-Host 'Checked for issues' -ForegroundColor Green
    Return $1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11
}
function getLicensing {
    Write-Host 'Getting license information...'
    $1 = "<h2 id='Lics'> Licensing</h2>"
    $2 = $cimLics | ConvertTo-Html -Fragment
    Write-Host 'Got license information' -ForegroundColor Green
    Return $1,$2
}
function getSecureInfo {
    Write-Host 'Getting security information...'
    $1 = "<h2 id='SecInfo'>Security Information</h2>"
    $2 = 'AV: ' + $av.DisplayName 
    $3 = 'Firewall: ' + $fw.DisplayName 
    $4 = "UAC: " + $(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System).EnableLUA 
    $5 = "Secureboot: " + $(Confirm-SecureBootUEFI) 
    $6 = "TPM: "
    $7 = $tpm | Select IsActivated_InitialValue,IsEnabled_InitialValue,IsOwned_InitialValue,PhysicalPresenceVersionInfo,SpecVersion | ConvertTo-Html -Fragment
    Write-Host 'Got security information' -ForegroundColor Green
    Return $1,$2,$3,$4,$5,$6,$7
}
function getTemps {
    Add-Type -Path .\files\OpenHardwareMonitorLib.dll
    $ohm = New-Object -TypeName OpenHardwareMonitor.Hardware.Computer
    $ohm.CPUEnabled= 1;
    $ohm.GPUEnabled = 1;
    $ohm.Open();
    foreach ($comp in $ohm.Hardware) {
        if($comp.HardwareType -eq [OpenHardwareMonitor.Hardware.HardwareType]::CPU){
        $comp.Update()
            foreach ($sens in $comp.Sensors) {
                 if ($sens.SensorType -eq [OpenHardwareMonitor.Hardware.SensorType]::Temperature) {
                    $1 = $sens.Value.ToString()
                    #$sens.Identifier for a better name
                 }
            }
        }
    }
    foreach ($comp in $ohm.Hardware) {
        if($comp.HardwareType -ne [OpenHardwareMonitor.Hardware.HardwareType]::CPU){
        $comp.Update()
            foreach ($sens in $comp.Sensors) {
                 if ($sens.SensorType -eq [OpenHardwareMonitor.Hardware.SensorType]::Temperature) {
                    $2 = $sens.Value.ToString() 
                    # $sens.Identifier for a better name
                 }
            }
        }
    }
    $ohm.Close();
    Return $1,$2
}
$temps = getTemps
function getCPU{
    Write-Host 'Getting hardware information...'
    $1 = "<h2 id='hw'>Hardware Basics</h2>"
    $cpuInfo = Get-WmiObject Win32_Processor
    $cpu = $cpuInfo.Name
    $2 = 'CPU: ' + $cpu + $temps[0] + 'C' 
    Return $1,$2
}
function getMobo{
    $moboBase = Get-WmiObject Win32_BaseBoard
    $moboMan = $moboBase.manufacturer
    $moboMod = $moboBase.product
    $mobo = $moboMan + " | " + $moboMod
    $1 = "Motherboard: " + $mobo 
    Return $1
}
function getGPU {
    $GPUbase = Get-WmiObject Win32_VideoController
    $1 = "Graphics Card: " + $GPUbase.Name + " " + $temps[1]+ 'C' 
    Return $1
}
function getRAM {
    $1 = "RAM: " + $(Get-WMIObject -class Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)}) + 'GB' 
    $2 = $(Get-WmiObject win32_physicalmemory | Select Manufacturer,Configuredclockspeed,Devicelocator,Capacity,Serialnumber | ConvertTo-Html -Fragment)
    Write-Host 'Got hardware information' -ForegroundColor Green
    Return $1,$2
}
function getVars {
    Write-Host 'Getting variables...'
    $1 = "<h2 id='SysVar'>System Variables</h2>"
    $2 = [Environment]::GetEnvironmentVariables("Machine")
    $3 = "<h2 id='UserVar'>User Variables</h2>"
    $4 = [Environment]::GetEnvironmentVariables("User")
    Write-Host 'Got variables' -ForegroundColor Green
    Return $1,$2,$3,$4
}
function getUpdates {
    Write-Host 'Getting applied hotfixes...'
    $1 = "<h2 id='hotfixes'>Installed updates</h2>"
    $2 = Get-HotFix | Sort-Object -Property InstalledOn -Descending | Select Description,HotFixID,InstalledOn | ConvertTo-Html -Fragment
    Write-Host 'Got applied hotfixes' -ForegroundColor Green
    Return $1,$2
}
function getStartup {
    Write-Host 'Getting startup tasks...'
    $1 = "<h2 id='StartupTasks'>Startup Tasks for user</h2>"
    $2 = $cimStart.Caption
    Write-Host 'Got startup tasks' -ForegroundColor Green
    Return $1,$2
}
function getPower {
    Write-Host 'Getting power profiles...'
    $1 = "<h2 id='Power'>Powerprofiles</h2>"
    $2 = powercfg /l
    Write-Host 'Got power profiles' -ForegroundColor Green
    Return $1,$2
}
function getRamUsage {
    Write-Host 'Getting running processes...'
    $1 = "<h2 id='RunningProcs'>Running Processes</h2>"
    $mem =  Get-WmiObject -Class WIN32_OperatingSystem
    $memUsed = [Math]::Round($($mem.TotalVisibleMemorySize - $mem.FreePhysicalMemory)/1048576,2)
    $memTotal = Get-WMIObject -class Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)}
    $2 = "Total RAM usage: " + $memUsed + "/" + $memTotal + " GB" 
    Return $1,$2
}
function getProcesses {
    $properties=@(
        @{Name="Name"; 
            Expression = {$_.name}},
        @{Name="Count"; 
            Expression = {(Get-Process -Name $_.Name | Group-Object -Property ProcessName).Count}
        },
        @{Name="NPM (M)"; 
            Expression = {[Math]::Round(($_.NPM / 1MB), 3)}
        },
        @{Name="Mem (M)"; 
            Expression = {
                $total = 0
                ForEach ($proc in $(Get-Process -Name $_.Name)) {
                    $total = $total + [Math]::Round(($proc.WS / 1MB), 2)
                }
                $total
            }
        },
        @{Name = "CPU"; 
            Expression = {
                $total = 0
                ForEach ($proc in $(Get-Process -Name $_.Name)) {
                    $TotalSec = (New-TimeSpan -Start $proc.StartTime).TotalSeconds
                    $total = $total + [Math]::Round( ($proc.CPU * 100 / $TotalSec), 2)
                }
                $total
            }
        },
        @{Name="ProductVersion ";
            Expression = {$_.ProductVersion}
        },
        @{Name="Path";
            Expression = {$_.Path}
        }
    )
    $1 = "`n"
    $2 = $runningProcesses | Select -Unique | Select $properties | Sort-Object "Mem (M)" -desc | ConvertTo-Html -Fragment
    Write-Host 'Got processes' -ForegroundColor Green
    return $1,$2
}
function getServices {
    Write-Host 'Getting services...'
    $1 = "<h2 id='Services'>Services</h2>"
    $2 = $services | Select Status,DisplayName | ConvertTo-Html -Fragment
    Write-Host 'Got services' -ForegroundColor Green
    Return $1,$2
}
function getInstalledApps {
    Write-Host 'Getting installed apps...'
    $apps = $installedBase | Select InstallDate,DisplayName | Sort-Object InstallDate -desc
    $1 = "<h2 id='InstalledApps'>Installed Apps</h2>"
    $2 = $apps | ConvertTo-Html -Fragment
    Write-Host 'Got installed apps' -ForegroundColor Green
    Return $1,$2
}
function getNets {
    Write-Host 'Getting network configurations...'
    $1 = "<h2 id='NetConfig'>Network Configuration</h2>"
    $2 = $(Get-NetAdapter|Select Name,InterfaceDescription,Status,LinkSpeed | ConvertTo-Html -Fragment) 
    $3 = $(Get-NetIPAddress|Select IpAddress,InterfaceAlias,PrefixOrigin | ConvertTo-Html -Fragment)
    Write-Host 'Got network configurations' -ForegroundColor Green
    Return $1,$2,$3
}
function getDrivers {
    Write-Host 'Getting driver information...'
    $1 = "<h2 id='Drivers'>Drivers and device versions</h2>"
    $2 = $(gwmi Win32_PnPSignedDriver | Select devicename,driverversion | ConvertTo-Html -Fragment)
    Write-Host 'Got driver information' -ForegroundColor Green
    Return $1,$2
}
function getAudio {
    Write-Host 'Getting audio devices...'
    $1 = "<h2 id='Audio'>Audio devices</h2>"
    $2 = $cimAudio | ConvertTo-Html -Fragment
    Write-Host 'Got audio devices' -ForegroundColor Green
    Return $1,$2
}
function getDisks {
    Write-Host 'Getting disk layouts...'
    $1 = "<h2 id='Disks'>Disk layouts</h2>"
    $2 = $(Get-Partition| Select OperationalStatus,DiskNumber,PartitionNumber,Size,IsActive,IsBoot,IsReadOnly | ConvertTo-Html -Fragment) 
    $3 = $volumes | Select HealthStatus,DriveType,FileSystem,FileSystemLabel,DedupMode,AllocationUnitSize,DriveLetter,SizeRemaining,Size |ConvertTo-Html -Fragment
    Write-Host 'Got disk layouts' -ForegroundColor Green
    Return $1,$2,$3
}
function getSmart {
    $1 = "<h2 id='SMART'>SMART</h2>"
    $2 = @()
    ForEach ($i in $smart) {
        $2 += $i | ConvertTo-Html -Fragment -As List
    }
    Return $1,$2
}
function getTimer {
    $timer.Stop()
    $1 = 'Runtime'
    $2 = $timer.Elapsed | Select Minutes,Seconds | ConvertTo-Html -Fragment
    Return $1,$2
}
function uploadFile {
    $link = Invoke-WebRequest -ContentType 'text/plain' -Method 'PUT' -InFile $file -Uri "https://paste.rtech.support/upload/$null.html" -UseBasicParsing
    $linkProper = $link.Content -Replace "support","support/selif"
    set-clipboard $linkProper
}
function promptStart {
        $Params = @{
        Content = "&#10;
This tool will gather specifications and configurations from your machine. After running this application will ask if you want to upload your results for sharing.
        &#10; 
        &#10;
Would you like to continue?
        &#10;
        &#10;
        &#10;
        &#10;
The source code for this application can be found at https://github.com/PipeItToDevNull/Get-Specs"
        Title = "rTechsupport Specs Tool"
        TitleBackground = "DodgerBlue"
        TitleFontSize = 16
        TitleFontWeight = "Bold"
        TitleTextForeground = "White"
        ContentFontSize = 12
        ContentFontWeight = "Medium"
        ContentTextForeground = "Black"
        ContentBackground = "White"
        ButtonType = "none"
        CustomButtons = "Start","Exit"
    }
    New-WPFMessageBox @Params
    If ($WPFMessageBoxOutput -eq "Exit") {
        Exit
    }
}
function promptUpload {
    $Params = @{
        Content = "Do you want to view or upload the specs?
        &#10;
Uploaded results are deleted after 24 hours"
        Title = "rTechsupport Specs Tool"
        TitleBackground = "DodgerBlue"
        TitleFontSize = 16
        TitleFontWeight = "Bold"
        TitleTextForeground = "White"
        ContentFontSize = 12
        ContentFontWeight = "Medium"
        ContentTextForeground = "Black"
        ContentBackground = "White"
        ButtonType = "none"
        CustomButtons = "View","Upload"
    }
    New-WPFMessageBox @Params
    If ($WPFMessageBoxOutput -eq "View") {
        Invoke-Item $file
    }
    ElseIf ($WPFMessageBoxOutput -eq "Upload") {
        uploadFile
        New-WPFMessageBox -Content "Link has been copied to your clipboard. Paste into chat to share." -Title "Upload Success" -TitleBackground Coral -WindowHost $Window
    }
}

# ------------------ #
promptStart
$timer = [diagnostics.stopwatch]::StartNew()
# ------------------ #

# Bulk data gathering
Write-Host 'Gathering main data...'
## CIM sources
$cimOs = Get-CimInstance -ClassName Win32_OperatingSystem
$cimStart = Get-CimInstance Win32_StartupCommand
$cimAudio= Get-CimInstance win32_sounddevice | Select Name,ProductName
$cimLics = Get-CimInstance -ClassName SoftwareLicensingProduct | ? { $_.PartialProductKey -ne $null } | Select Name,ProductKeyChannel,LicenseFamily,LicenseStatus,PartialProductKey
$av = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct
$fw = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName FirewallProduct
$tpm = Get-CimInstance -Namespace root/cimv2/Security/MicrosoftTpm -ClassName win32_tpm

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

## Build info
$i = 0
$eol = $false
$eolON = "Unknown"
$osBuild = 'Version is unknown. Build: ' + $cimOS.BuildNumber
foreach ($b in $builds) {
    if ($cimOS.BuildNumber -eq $builds[$i]) {
        $osBuild = 'Version: ' + $versions[$i]
        if ((Get-Date $eolDates[$i]) -lt (Get-Date)) {
            $eol = $true
            $eolOn = $eolDates[$i]
        }
    }
    $i = $i + 1
}
Write-Host 'Got main data' -ForegroundColor Green

# get smart now so we can parse it for bad things
Write-Host 'Getting SMART data...'
$smart = .\files\Get-Smart\Get-Smart.ps1 -cdiPath '.\files\DiskInfo64.exe'
cd ..
Write-Host 'Got SMART' -ForegroundColor Green

# ------------------ #
# Write da file
header | Out-File -Encoding ascii $file
getDate | Out-File -Append -Encoding ascii $file
getBasicInfo | Out-File -Append -Encoding ascii $file
getBadThings | Out-File -Append -Encoding ascii $file
table | Out-File -Append -Encoding ascii $file
getLicensing | Out-File -Append -Encoding ascii $file
getSecureInfo | Out-File -Append -Encoding ascii $file
getCPU | Out-File -Append -Encoding ascii $file
getMobo | Out-File -Append -Encoding ascii $file
getGPU | Out-File -Append -Encoding ascii $file
getRAM | Out-File -Append -Encoding ascii $file
getVars | Out-File -Append -Encoding ascii $file
getUpdates | Out-File -Append -Encoding ascii $file
getStartup | Out-File -Append -Encoding ascii $file
getPower | Out-File -Append -Encoding ascii $file
getRamUsage | Out-File -Append -Encoding ascii $file
getProcesses | Out-File -Append -Encoding ascii $file
getServices | Out-File -Append -Encoding ascii $file
getInstalledApps | Out-File -Append -Encoding ascii $file
getNets | Out-File -Append -Encoding ascii $file
getDrivers | Out-File -Append -Encoding ascii $file
getAudio | Out-File -Append -Encoding ascii $file
getDisks | Out-File -Append -Encoding ascii $file
getSmart | Out-File -Append -Encoding ascii $file
getTimer | Out-File -Append -Encoding ascii $file
promptUpload
# ------------------ #
