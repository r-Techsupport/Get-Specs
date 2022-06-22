<#
.SYNOPSIS
 Gather and upload specifications of a Windows host to rTechsupport
.DESCRIPTION
 Use various native powershell and wmic functions to gather verbose information on a system to assist troubleshooting
.OUTPUTS Specs
  '.\TechSupport_Specs.html'
#>
# VERSION
$version = '1.5.2'

# source our other ps1 files
. files\wpf.ps1

# Declarations
## files we use
$file = 'TechSupport_Specs.html'
$today = Get-Date

## hashes for executables
$hashPaths = @(
    './files/DiskInfo64.exe',
    './files/OpenHardwareMonitorLib.dll',
    './files/wpf.ps1',
    './files/Get-Smart/Get-SMART.ps1'
)
$hashSums = @(
    '58778361E2DCDA5DEB94D52C071FCEDA48FE7850D2D992939DE2A4210C347246',
    'D657E857466FE517CB91B4C3742B40DE4566D051D72C7D836921A7D47947EB70',
    '7D214F840DFB47C1BC900D1EF6700DAF7C7697676452866D765264C0D95927F6',
    'CB9EAA8903975F23196918EF7DFC1CAE9D4F28046A0955C842D410206DFB7B59'
)

## hosts related
$hostsFile = 'C:\Windows\System32\drivers\etc\hosts'
$hostsHash = '2D6BDFB341BE3A6234B24742377F93AA7C7CFB0D9FD64EFA9282C87852E57085'
$hostsSum = $(Get-FileHash $hostsFile).hash
$hostsContent = Get-Content $hostsFile

## bad things
$badSoftware = @(
    'Driver Booster*',
    'iTop*',
    'Driver Easy*',
    'Roblox*',
    'ccleaner',
    'Malwarebytes',
    'Wallpaper Engine',
    'Voxal Voice Changer',
    'Clownfish Voice Changer',
    'Voicemod',
    'Microsoft Office Enterprise 2007',
    'Memory Cleaner*',
    'System Mechanic',
    'MyCleanPC',
    'DriverFix*',
    'Reimage Repair',
    'cFosSpeed*',
    'Browser Assistant',
    'KMS*'
)
$badStartup = @(
    'AutoKMS',
    'kmspico',
    'McAfee Remediation',
    'KMS_VL_ALL',
    'WallpaperEngine'
)
$badProcesses = @(
    'MBAMService',
    'McAfee WebAdvisor',
    'Norton Security',
    'Wallpaper Engine Service',
    'Service_KMS.exe',
    'iTopVPN',
    'wallpaper32',
    'TaskbarX'
)
# YOU MUST MATCH THE KEY AND VALUE BELOW TO THE SAME ARRAY VALUE
$badKeys = @(
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PreviewBuilds\',
    'HKLM:\SYSTEM\Setup\LabConfig\',
    'HKLM:\SYSTEM\Setup\LabConfig\',
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\',
    'HKLM:\SYSTEM\Setup\Status\',
    'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers',
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
)
$badValues = @(
    'AllowBuildPreview',
    'BypassTPMCheck',
    'BypassSecureBootCheck',
    'NoAutoUpdate',
    'AuditBoot',
    'HwSchMode',
    'UseWUServer'
)
$badData = @(
    '1',
    '1',
    '1',
    '1',
    '1',
    '2',
    '1'
)
$badRegExp = @(
    'Insider builds are set to: ',
    'Windows 11 TPM bypass set to: ',
    'Windows 11 SecureBoot bypass set to: ',
    'Windows auto update is set to: ',
    'Audit Boot is set to: ',
    '(HAGS) HwSchMode is set to: '
)
$badHostnames = @(
    'ATLASOS-DESKTOP',
    'Revision-PC'
)
$badAdapters = @(
    '*TAP*',
    '*TUN*',
    '*VPN*',
    '*Hamachi*',
    '*Tunnel*',
    '*Nord*',
    '*SurfShark*',
    'TunnelBear Adapter V9',
    'Private Internet Access Network Adapter',
    'ZeroTier Virtual Port',
    'Kaspersky Security Data Escort Adapter'
)
$badFiles = @(
    'C:\Windows\system32\SppExtComObjHook.dll',
    'HKCU:\Software\azurite'
)
$microCode = @(
    'C:\Windows\System32\mcupdate_genuineintel.dll',
    'C:\WindowsSystem32\mcupdate_authenticamd.dll'
)
$builds = @(
    '10240',
    '10586',
    '14393',
    '15063',
    '16299',
    '17134',
    '17763',
    '18362',
    '18363',
    '19041',
    '19042',
    '19043',
    '19044',
    '22000'
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
    '21H2',
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
    '2023-10-10',
    '2023-10-10'
)

### Functions

function header {
$1 = '<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />

<style>
* {
    font-family: verdana !important;
    font-size: 12px;
}
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
#topbutton{
  opacity: 80%;
  width: 5%;
  padding-top: -3%;
  background-color: #ccc;
  position: fixed;
  bottom: 0;
  right: 0;
  border-radius: 20px;
  text-align: center;
  font-size: 24px;
  color: #87ab63;
}
</style>
<body>
<pre>
<a href="#top"><div id="topbutton">
TOP
</div></a>
'
Return $1
}
function table {
$1 = '<a name="top"></a>
<h2>Sections</h2>
<div style="line-height:0">
<p><a href="#hw">Hardware Basics</a></p>
<p><a href="#Lics">Licensing</a></p>
<p><a href="#SecInfo">Security Information</a></p>
<p><a href="#bios">BIOS</a></p>
<p><a href="#SysVar">System Variables</a></p>
<p><a href="#UserVar">User Variables</a></p>
<p><a href="#hotfixes">Installed updates</a></p>
<p><a href="#StartupTasks">Startup Tasks</a></p>
<p><a href="#Power">Powerprofiles</a></p>
<p><a href="#RunningProcs">Running Processes</a></p>
<p><a href="#Services">Services</a></p>
<p><a href="#InstalledApps">Installed Applications</a></p>
<p><a href="#NetConfig">Network Configuration</a></p>
<p><a href="#NetConnections">Network Connections</a></p>
<p><a href="#Drivers">Drivers and device versions</a></p>
<p><a href="#issueDevices">Devices with issues</a></p>
<p><a href="#Audio">Audio Devices</a></p>
<p><a href="#Disks">Disk Layouts</a></p>
<p><a href="#SMART">SMART</a></p>
</div>'
Return $1
}

function checkExecutableHashes {
    $i = 0
    ForEach ($file in $hashPaths) {
        $currentHash = $(Get-FileHash $file).Hash
        If ($currentHash -ne $hashSums[$i] {
            Throw "$file sum mismatch"
        }
        $i++
    }
}

function getDate {
    $1 = "Local: " + $today
    $2 = "UTC: " + $([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId(($today), 'Greenwich Standard Time'))
    $3 = "Version: " + $version
    Return $1,$2,$3
}
function getbasicInfo {
    Write-Host 'Getting basic information...'
    $1 = '<h2>System Information</h2>'
    $bootuptime = $cimOs.LastBootUpTime
    $uptime = $today - $bootuptime
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
function getNotes {
    Write-Host 'Checking for notes...'
    $1 = '<h2>Notes</h2>'
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
        if ($startUps -contains $start) { 
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
        $6 = "This may be a malicious distribution of Windows: " + $cimOS.CSName
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
        $9 = "Utopia malware found, router is infected. This adds a 'utopia.net' suffix to your DNS and can cause issues with shortnames."
    }
    If ($eol) {
        $10 = "OS was EOL on " + $eolOn + " Build is " + $osBuild
    }
    # bad smart things
    If (-Not $smart) {
       $11 = "SMART timeout. Manually run Crystal Disk Info to determine your failed disk(s)" 
    }
    $12 = @()
    Foreach ($disk in $smart) {
        If ($disk.'Reallocated Sectors Count') {
            If ($disk.'Reallocated Sectors Count' -ne '000000000000') {
                $12 += "Reallocated sector on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Reallocated Sectors Count'
            }
        }
        If ($disk.'Current Pending Sector Count') {
             If ($disk.'Current Pending Sector Count' -ne '000000000000') {
                 $12 += "Current Pending Sector Count on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Current Pending Sector Count'
            }
        }
        If ($disk.'Uncorrectable Sector Count') {
            If ($disk.'Uncorrectable Sector Count' -ne '000000000000') {
                $12 += "Uncorrectable Sector Count on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Uncorrectable Sector Count'
            }
        }
        If ($disk.'Command Timeout') {
            If ($disk.'Command Timeout' -ne '000000000000') {
                $12 += "Command Timeout on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Command Timeout'
            }
        }
        If ($disk.'Reported Uncorrectable Errors') {
            If ($disk.'Reported Uncorrectable Errors' -ne '000000000000') {
                $12 += "Reported Uncorrectable Errors on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Reported Uncorrectable Errors'
            }
        }
        If ($disk.'CRC Error Count') {
            If ($disk.'CRC Error Count' -ne '000000000000') {
                $12 += "CRC Error Count on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'CRC Error Count'
            }
        }
        If ($disk.'Rotation Rate' -NotLike '---- (SSD)' -And $disk.'Drive Letter' -eq 'C:') {
            $13 += "C: is on an HDD"
        }
    }
    # check for VPNs being connected
    $cAdapters = $netAdapters | Where {$_.MediaConnectionState -eq 'Connected'}
    ForEach ($adapter in $badAdapters) { 
        If ($cAdapters.IfDesc -Like $adapter) { 
            $14 = "VPN is connected"
        }
    }
    # check for specific files in the FS
    $15 = @()
    ForEach ($file in $badFiles) {
        If (Test-Path $file) {
            $14 += "Bad file found $file"
        }
    }
    # check if the user configured a static number of cores in msconfig
    If ($bcdedit -ne $NULL) {
        $16 = "Static core number is set in msconfig"
    }
    # ram checks
    If ($ramSticks.count -eq 2) {
        If ($ramSticks[0].Devicelocator -eq 'DIMM1' -And $ramSticks[1].Devicelocator -eq 'DIMM2') {
            $17 = "RAM is in DIMM1 and DIMM2"
        }
        If ($ramSticks[0].Devicelocator -eq 'DIMM3' -And $ramSticks[1].Devicelocator -eq 'DIMM4') {
            $17 = "RAM is in DIMM3 and DIMM4"
        }
    }
    # hosts checks 
    If ($hostsSum -ne $hostsHash) {
        $18 = "Hosts file has been modified from stock, it has been appended to bottom of this page"
        If ($hostsContent -Like "*license.piriform.com*") {
            $19 = "piriform license server redirection"
        }
        If ($hostsContent -Like "*Spybot*") {
            $20 = "spybot has edited hosts file"
        }
    }
    # check for dumps
    $dmpFiles = Get-ChildItem 'C:\Windows\Minidump'
    $lastWeek = $today.AddDays(-7)
    $count = 0
    ForEach ($dmp in $dmpFiles) {
        If ($dmp.LastWriteTime -gt $lastWeek) {
            $count++
        }
    }
    If ($count -gt 0) {
        $21 = "Found $count dumps"
    }
    # check for missing files in FS
    If (!(Test-Path $microCode[0]) -And !(Test-Path $microCode[1])) {
        $22 += "Microcode fixes are missing or were removed by malicious 'fixers'"
    }
    # count and report issue devices
    If ($issueDevices.Status.Count -gt 0) {
        $23 = $issueDevices.Status.Count + " devices have issues, see 'Devices with issues' section"
    }
    # Check for TPM and secure boot if on Windows 11
    If ([System.Environment]::OSVersion.Version.Build -ge 22000) {
        # Why must SpecVersion be a string :(
        If ($tpm -eq $NULL) {
            $24 = "No TPM"
        }
        ElseIf ([int][string]($tpm.SpecVersion[0]) -l 2) {
            $24 = "Windows 11 TPM version not satisfied"
        }

        If (!(Confirm-SecureBootUEFI)) {
            $25 = "Secure boot not enabled"
        }
    }

    Write-Host 'Checked for notes' -ForegroundColor Green
    Return $1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$22,$23,$24,$25
}
function getTemps {
    Try {
        Add-Type -Path .\files\OpenHardwareMonitorLib.dll -ErrorAction SilentlyContinue
    }
    catch {
        $1 = ' OHMF'
        $2 = ' OHMF'
        return $1,$2
    }
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
function getHardware {
    Write-Host 'Getting harwdware information...'
    $1 = "<h2 id='hw'>Hardware Basics</h2>"
    $cpu = Get-WmiObject Win32_Processor
    $mobo = Get-WmiObject Win32_BaseBoard
    $gpu = Get-WmiObject Win32_VideoController

    $hwArray = @()

    $cpuObject = New-Object PSobject
    Add-Member -InputObject $cpuObject -MemberType NoteProperty -Name "Part" -Value "CPU"
    Add-Member -InputObject $cpuObject -MemberType NoteProperty -Name "Manufacturer" -Value $cpu.Manufacturer
    Add-Member -InputObject $cpuObject -MemberType NoteProperty -Name "Product" -Value $cpu.Name
    Add-Member -InputObject $cpuObject -MemberType NoteProperty -Name "Temperature" -Value $temps[0]
    Add-Member -InputObject $cpuObject -MemberType NoteProperty -Name "Driver" -Value ""
    $hwArray += $cpuObject

    $moboObject = New-Object PSObject
    Add-Member -InputObject $moboObject -MemberType NoteProperty -Name "Part" -Value "Motherboard"
    Add-Member -InputObject $moboObject -MemberType NoteProperty -Name "Manufacturer" -Value $mobo.Manufacturer
    Add-Member -InputObject $moboObject -MemberType NoteProperty -Name "Product" -Value $mobo.Product
    Add-Member -InputObject $moboObject -MemberType NoteProperty -Name "Driver" -Value ""
    $hwArray += $moboObject

    $gpuDriver = $(gwmi Win32_PnPSignedDriver | ? { $_.deviceName -eq $gpu.Name }).driverversion
    If (Get-Member -inputobject $gpu -name "Count" -Membertype Properties) {
        $i = 0
        foreach ($g in $gpu) {
            $gpuObject = New-Object PSObject
            Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Part" -Value "Video Card"
            Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Manufacturer" -Value $gpu[$i].AdapterCompatibility
            Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Product" -Value $gpu[$i].Name
            Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Temperature" -Value $temps[1]
            Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Driver" -Value $gpuDriver
            $hwArray += $gpuObject
            $i = $i + 1
        }
    } Else {
        $gpuObject = New-Object PSObject
        Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Part" -Value "Video Card"
        Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Manufacturer" -Value $gpu.AdapterCompatibility
        Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Product" -Value $gpu.Name
        Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Temperature" -Value $temps[1]
        Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Driver" -Value $gpuDriver
        $hwArray += $gpuObject
    }

    $2 = $hwArray | ConvertTo-Html -Fragment

    Return $1,$2
}
function getRAM {
    $totalRam = $(Get-WMIObject -class Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)})
    $ramObject = New-Object PSObject
    Add-Member -InputObject $ramObject -MemberType NoteProperty -Name "Total" -Value $totalRam
    Add-Member -InputObject $ramObject -MemberType NoteProperty -Name "RAM" -Value "GB"
    $1 = $ramObject | ConvertTo-Html -Fragment

    $2 = $($ramSticks | Select Manufacturer,Configuredclockspeed,Devicelocator,Capacity,Serialnumber,PartNumber | ConvertTo-Html -Fragment)
    Write-Host 'Got hardware information' -ForegroundColor Green
    Return $1,$2
}
function getLicensing {
    Write-Host 'Getting license information...'
    $1 = "<h2 id='Lics'>Licensing</h2>"
    $2 = $cimLics | ConvertTo-Html -Fragment
    Write-Host 'Got license information' -ForegroundColor Green
    Return $1,$2
}
function getSecureInfo {
    Write-Host 'Getting security information...'
    $1 = "<h2 id='SecInfo'>Security Information</h2>"

    $secObject = New-Object PSObject
    # Add AVs
    $i = 0
    ForEach ($a in $av) {
        Add-Member -InputObject $secObject -MemberType NoteProperty -Name "Antivirus$i" -Value $a.DisplayName
        $i++
    }
    # Add FW and assume its defender if there is no entry (default)
    If ($fw.DisplayName -eq $NULL) {
        Add-Member -InputObject $secObject -MemberType NoteProperty -Name 'Firewall' -Value "Assume Defender"
        } Else {
        Add-Member -InputObject $secObject -MemberType NoteProperty -Name 'Firewall' -Value $fw.DisplayName
    }
    # Add UAC
    Add-Member -InputObject $secObject -MemberType NoteProperty -Name 'UAC' -Value $(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System).EnableLUA
    # Add Secureboot
    Add-Member -InputObject $secObject -MemberType NoteProperty -Name 'SecureBoot' -Value $(Confirm-SecureBootUEFI -ErrorAction SilentlyContinue)
    $2 = $secObject | ConvertTo-Html -Fragment -As List

    $6 = "<h2 id='TPM'>TPM</h2>"
    If ($tpm -eq $NULL) {
        $7 = 'TPM not detected'
        } Else {
        $7 = $tpm | Select IsActivated_InitialValue,IsEnabled_InitialValue,IsOwned_InitialValue,PhysicalPresenceVersionInfo,SpecVersion | ConvertTo-Html -Fragment
    }
    Write-Host 'Got security information' -ForegroundColor Green
    Return $1,$2,$6,$7
}
function getBIOS {
    $1 = "<h2 id='bios'>BIOS</h2>"
    $2 = $cimBios | Select Manufacturer,SMBIOSBIOSVersion,Name,Version | ConvertTo-Html -Fragment
    Return $1,$2
}
function getVars {
    Write-Host 'Getting variables...'
    $1 = "<h2 id='SysVar'>System Variables</h2>"
    $varsMachine = [Environment]::GetEnvironmentVariables("Machine")
    $varObjMachine = New-Object PSObject
    ForEach ($var in $varsMachine.Keys) {
        Add-Member -InputObject $varObjMachine -MemberType NoteProperty -Name $var -Value $varsMachine.$var
    }
    $2 = $varObjMachine | ConvertTo-Html -Fragment -As List
    $3 = "<h2 id='UserVar'>User Variables</h2>"
    $varsUser = [Environment]::GetEnvironmentVariables("User")
    $varObjUser = New-Object PSObject
    ForEach ($var in $varsUser.Keys) {
        Add-Member -InputObject $varObjUser -MemberType NoteProperty -Name $var -Value $varsUser.$var
    }
    $4 = $varObjUser | ConvertTo-Html -Fragment -As List
    Write-Host 'Got variables' -ForegroundColor Green
    Return $1,$2,$3,$4
}
function getUpdates {
    Write-Host 'Getting applied hotfixes...'
    $1 = "<h2 id='hotfixes'>Installed updates</h2>"
    $2 = Get-HotFix | Sort-Object -Property InstalledOn -Descending -ErrorAction SilentlyContinue | Select Description,HotFixID,InstalledOn | ConvertTo-Html -Fragment
    Write-Host 'Got applied hotfixes' -ForegroundColor Green
    Return $1,$2
}
function getStartup {
    Write-Host 'Getting startup tasks...'
    $1 = "<h2 id='StartupTasks'>Startup Tasks for user</h2>"
    $2 = $startUps
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
            Expression = {(Get-Process -Name $_.Name -ErrorAction SilentlyContinue | Group-Object -Property ProcessName).Count}
        },
        @{Name="NPM (M)"; 
            Expression = {[Math]::Round(($_.NPM / 1MB), 3)}
        },
        @{Name="Mem (M)"; 
            Expression = {
                $totar = 0
                ForEach ($proc in $(Get-Process -Name $_.Name -ErrorAction SilentlyContinue)) {
                    $total = $total + [Math]::Round(($proc.WS / 1MB), 2)
                }
                $total
            }
        },
        @{Name = "CPU"; 
            Expression = {
                $total = 0
                ForEach ($proc in $(Get-Process -Name $_.Name -ErrorAction SilentlyContinue)) {
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
    $servicesTable = $services | Select Status,DisplayName | Sort -Property DisplayName | ConvertTo-Html -Fragment
    $2 = $servicesTable -replace "<td>Stopped</td>", "<td style='color:#ab6387'>Stopped</td>" -replace "<td>Running</td>", "<td style='color:#87ab63'>Running</td>"
    Write-Host 'Got services' -ForegroundColor Green
    Return $1,$2
}
function getInstalledApps {
    Write-Host 'Getting installed apps...'
    $apps = $installedBase | Select InstallDate,DisplayName | Sort-Object DisplayName
    $1 = "<h2 id='InstalledApps'>Installed Apps</h2>"
    $2 = $apps | ConvertTo-Html -Fragment
    Write-Host 'Got installed apps' -ForegroundColor Green
    Return $1,$2
}
function getNets {
    Write-Host 'Getting network configurations...'
    $1 = "<h2 id='NetConfig'>Network Configuration</h2>"
    # make an object for each adatper and then add them to an array
    $netArray = @()
    ForEach ($int in $netAdapters) {
        $intObject = New-Object PSObject 
        Add-Member -InputObject $intObject -MemberType NoteProperty -Name "Name" -Value $int.Name -ErrorAction SilentlyContinue
        Add-Member -InputObject $intObject -MemberType NoteProperty -Name "State" -Value $int.MediaConnectionState -ErrorAction SilentlyContinue
        Add-Member -InputObject $intObject -MemberType NoteProperty -Name "Mac" -Value $int.MacAddress -ErrorAction SilentlyContinue
        Add-Member -InputObject $intObject -MemberType NoteProperty -Name "Description" -Value $int.ifDesc -ErrorAction SilentlyContinue

        $i = 0
        $ips = $(Get-NetIPAddress -InterfaceIndex $int.IfIndex)
        ForEach ($ip in $ips) {
            Add-Member -InputObject $intObject -MemberType NoteProperty -Name $ip.AddressFamily -Value $ip.IPAddress -ErrorAction SilentlyContinue
            Add-Member -InputObject $intObject -MemberType NoteProperty -Name "Lease$i" -Value $ip.PrefixOrigin -ErrorAction SilentlyContinue
            $i++
        }

        $dnsS = $(Get-DnsClientServerAddress -InterfaceIndex $int.IfIndex)
        ForEach ($dns in $dnsS) {
            $i = 0
            ForEach ($d in $dns.ServerAddresses) {
                Add-Member -InputObject $intObject -MemberType NoteProperty -Name "DNS$i" -Value $d -ErrorAction SilentlyContinue
                $i++
            }
        }
        $netArray += $intObject
    }
    $2 = $netArray | ConvertTo-Html -Fragment -As List
    Write-Host 'Got network configurations' -ForegroundColor Green
    Return $1,$2
}
function getNetConnections {
    Write-Host 'Getting network connections...'
    $1 = "<h2 id='NetConnections'>Network Connections</h2>"
    $2 = Get-NetTCPConnection | Select local*,remote*,state,@{Name="Process";Expression={(Get-Process -Id $_.OwningProcess).ProcessName}},@{Name="Path";Expression={(Get-Process -Id $_.OwningProcess).Path}} | ConvertTo-Html -Fragment
    Write-Host 'Got network connections' -ForegroundColor Green
    Return $1,$2
}
function getDrivers {
    Write-Host 'Getting driver information...'
    $1 = "<h2 id='Drivers'>Drivers and device versions</h2>"
    $2 = $(gwmi Win32_PnPSignedDriver | Select devicename,driverversion | ConvertTo-Html -Fragment)
    $3 = "<h2 id='issueDevices'>Devices with issues</h2>"
    $4 = $issueDevices | Select Status,Name,InstanceID | ConvertTo-HTML -Fragment
    Write-Host 'Got driver information' -ForegroundColor Green
    Return $1,$2,$3,$4
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
    $disks = Get-Partition | Select DiskNumber -Unique
    $2 = $($disks | % { Get-Partition -DiskNumber $_.DiskNumber | Select OperationalStatus,DiskNumber,PartitionNumber,@{Name="Size (GB)";Expression={[math]::round($_.Size/1GB,4)}},IsActive,IsBoot,IsReadOnly | ConvertTo-Html -Fragment })
    $3 = $volumes | Select HealthStatus,DriveType,FileSystem,FileSystemLabel,DedupMode,AllocationUnitSize,DriveLetter,@{Name="Size Remaining (GB)";Expression={[math]::round($_.SizeRemaining/1GB,4)}}, @{Name="Size Total (GB)";Expression={[math]::round($_.Size/1GB,4)}} |ConvertTo-Html -Fragment
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
function getHosts {
    If ($hostsSum -ne $hostsHash) {
        $hostsSum
        $hostsContent
        Write-Output ""
    }
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
    Start-Process $linkProper
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
    Write-Host "Completed. Check for another prompt offering to view or upload the results."
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
$taskStart = Get-ScheduledTask -TaskName * | ? { $_.Triggers -Like "MSFT_TaskBootTrigger" -Or $_.Triggers -Like "MSFT_TaskLogonTrigger" } | ? { $_.State -eq "Ready" -Or $_.State -eq "Running" } | ? { $_.Author -NotLike "Microsof*" } | Select TaskName
$startUps = $cimStart.caption + $taskStart.TaskName
$cimAudio= Get-CimInstance win32_sounddevice | Select Name,ProductName
$cimLics = Get-CimInstance -ClassName SoftwareLicensingProduct | ? { $_.PartialProductKey -ne $null } | Select Name,ProductKeyChannel,LicenseFamily,LicenseStatus,PartialProductKey
$av = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct
$fw = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName FirewallProduct
$tpm = Get-CimInstance -Namespace root/cimv2/Security/MicrosoftTpm -ClassName win32_tpm
$ramSticks = Get-WmiObject win32_physicalmemory
$cimBios = Get-CimInstance Win32_bios

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
$issueDevices = Get-PnpDevice -PresentOnly -Status ERROR,DEGRADED,UNKNOWN -ErrorAction SilentlyContinue

# janky check for msconfig core setting
$bcdedit = bcdedit | Select-String numproc

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
$smart = .\files\Get-Smart\Get-Smart.ps1 -cdiPath '.\files\DiskInfo64.exe'
cd ..
Write-Host 'Got SMART' -ForegroundColor Green

# ------------------ #
# Write da file
header | Out-File -Encoding ascii $file
getDate | Out-File -Append -Encoding ascii $file
getBasicInfo | Out-File -Append -Encoding ascii $file
getNotes | Out-File -Append -Encoding ascii $file
table | Out-File -Append -Encoding ascii $file
getHardware | Out-File -Append -Encoding ascii $file
getRAM | Out-File -Append -Encoding ascii $file
getLicensing | Out-File -Append -Encoding ascii $file
getSecureInfo | Out-File -Append -Encoding ascii $file
getBIOS | Out-File -Append -Encoding ascii $file
getVars | Out-File -Append -Encoding ascii $file
getUpdates | Out-File -Append -Encoding ascii $file
getStartup | Out-File -Append -Encoding ascii $file
getPower | Out-File -Append -Encoding ascii $file
getRamUsage | Out-File -Append -Encoding ascii $file
getProcesses | Out-File -Append -Encoding ascii $file
getServices | Out-File -Append -Encoding ascii $file
getInstalledApps | Out-File -Append -Encoding ascii $file
getNets | Out-File -Append -Encoding ascii $file
getNetConnections | Out-File -Append -Encoding ascii $file
getDrivers | Out-File -Append -Encoding ascii $file
getAudio | Out-File -Append -Encoding ascii $file
getDisks | Out-File -Append -Encoding ascii $file
getSmart | Out-File -Append -Encoding ascii $file
getHosts | Out-FIle -Append -Encoding ascii $file
getTimer | Out-File -Append -Encoding ascii $file
promptUpload
# ------------------ #

# SIG # Begin signature block
# MIIVogYJKoZIhvcNAQcCoIIVkzCCFY8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUZ1CUQ4RUoGGWlj+WojXQSyPv
# FSKgghICMIIFbzCCBFegAwIBAgIQSPyTtGBVlI02p8mKidaUFjANBgkqhkiG9w0B
# AQwFADB7MQswCQYDVQQGEwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHDAdTYWxmb3JkMRowGAYDVQQKDBFDb21vZG8gQ0EgTGltaXRlZDEh
# MB8GA1UEAwwYQUFBIENlcnRpZmljYXRlIFNlcnZpY2VzMB4XDTIxMDUyNTAwMDAw
# MFoXDTI4MTIzMTIzNTk1OVowVjELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEtMCsGA1UEAxMkU2VjdGlnbyBQdWJsaWMgQ29kZSBTaWduaW5n
# IFJvb3QgUjQ2MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAjeeUEiIE
# JHQu/xYjApKKtq42haxH1CORKz7cfeIxoFFvrISR41KKteKW3tCHYySJiv/vEpM7
# fbu2ir29BX8nm2tl06UMabG8STma8W1uquSggyfamg0rUOlLW7O4ZDakfko9qXGr
# YbNzszwLDO/bM1flvjQ345cbXf0fEj2CA3bm+z9m0pQxafptszSswXp43JJQ8mTH
# qi0Eq8Nq6uAvp6fcbtfo/9ohq0C/ue4NnsbZnpnvxt4fqQx2sycgoda6/YDnAdLv
# 64IplXCN/7sVz/7RDzaiLk8ykHRGa0c1E3cFM09jLrgt4b9lpwRrGNhx+swI8m2J
# mRCxrds+LOSqGLDGBwF1Z95t6WNjHjZ/aYm+qkU+blpfj6Fby50whjDoA7NAxg0P
# OM1nqFOI+rgwZfpvx+cdsYN0aT6sxGg7seZnM5q2COCABUhA7vaCZEao9XOwBpXy
# bGWfv1VbHJxXGsd4RnxwqpQbghesh+m2yQ6BHEDWFhcp/FycGCvqRfXvvdVnTyhe
# Be6QTHrnxvTQ/PrNPjJGEyA2igTqt6oHRpwNkzoJZplYXCmjuQymMDg80EY2NXyc
# uu7D1fkKdvp+BRtAypI16dV60bV/AK6pkKrFfwGcELEW/MxuGNxvYv6mUKe4e7id
# FT/+IAx1yCJaE5UZkADpGtXChvHjjuxf9OUCAwEAAaOCARIwggEOMB8GA1UdIwQY
# MBaAFKARCiM+lvEH7OKvKe+CpX/QMKS0MB0GA1UdDgQWBBQy65Ka/zWWSC8oQEJw
# IDaRXBeF5jAOBgNVHQ8BAf8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zATBgNVHSUE
# DDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEMGA1Ud
# HwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29tb2RvY2EuY29tL0FBQUNlcnRpZmlj
# YXRlU2VydmljZXMuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3DQEBDAUAA4IBAQASv6Hvi3Sa
# mES4aUa1qyQKDKSKZ7g6gb9Fin1SB6iNH04hhTmja14tIIa/ELiueTtTzbT72ES+
# BtlcY2fUQBaHRIZyKtYyFfUSg8L54V0RQGf2QidyxSPiAjgaTCDi2wH3zUZPJqJ8
# ZsBRNraJAlTH/Fj7bADu/pimLpWhDFMpH2/YGaZPnvesCepdgsaLr4CnvYFIUoQx
# 2jLsFeSmTD1sOXPUC4U5IOCFGmjhp0g4qdE2JXfBjRkWxYhMZn0vY86Y6GnfrDyo
# XZ3JHFuu2PMvdM+4fvbXg50RlmKarkUT2n/cR/vfw1Kf5gZV6Z2M8jpiUbzsJA8p
# 1FiAhORFe1rYMIIGGjCCBAKgAwIBAgIQYh1tDFIBnjuQeRUgiSEcCjANBgkqhkiG
# 9w0BAQwFADBWMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVk
# MS0wKwYDVQQDEyRTZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYw
# HhcNMjEwMzIyMDAwMDAwWhcNMzYwMzIxMjM1OTU5WjBUMQswCQYDVQQGEwJHQjEY
# MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1Ymxp
# YyBDb2RlIFNpZ25pbmcgQ0EgUjM2MIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIB
# igKCAYEAmyudU/o1P45gBkNqwM/1f/bIU1MYyM7TbH78WAeVF3llMwsRHgBGRmxD
# eEDIArCS2VCoVk4Y/8j6stIkmYV5Gej4NgNjVQ4BYoDjGMwdjioXan1hlaGFt4Wk
# 9vT0k2oWJMJjL9G//N523hAm4jF4UjrW2pvv9+hdPX8tbbAfI3v0VdJiJPFy/7Xw
# iunD7mBxNtecM6ytIdUlh08T2z7mJEXZD9OWcJkZk5wDuf2q52PN43jc4T9OkoXZ
# 0arWZVeffvMr/iiIROSCzKoDmWABDRzV/UiQ5vqsaeFaqQdzFf4ed8peNWh1OaZX
# nYvZQgWx/SXiJDRSAolRzZEZquE6cbcH747FHncs/Kzcn0Ccv2jrOW+LPmnOyB+t
# AfiWu01TPhCr9VrkxsHC5qFNxaThTG5j4/Kc+ODD2dX/fmBECELcvzUHf9shoFvr
# n35XGf2RPaNTO2uSZ6n9otv7jElspkfK9qEATHZcodp+R4q2OIypxR//YEb3fkDn
# 3UayWW9bAgMBAAGjggFkMIIBYDAfBgNVHSMEGDAWgBQy65Ka/zWWSC8oQEJwIDaR
# XBeF5jAdBgNVHQ4EFgQUDyrLIIcouOxvSK4rVKYpqhekzQwwDgYDVR0PAQH/BAQD
# AgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwGwYD
# VR0gBBQwEjAGBgRVHSAAMAgGBmeBDAEEATBLBgNVHR8ERDBCMECgPqA8hjpodHRw
# Oi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RS
# NDYuY3JsMHsGCCsGAQUFBwEBBG8wbTBGBggrBgEFBQcwAoY6aHR0cDovL2NydC5z
# ZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdSb290UjQ2LnA3YzAj
# BggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEM
# BQADggIBAAb/guF3YzZue6EVIJsT/wT+mHVEYcNWlXHRkT+FoetAQLHI1uBy/YXK
# ZDk8+Y1LoNqHrp22AKMGxQtgCivnDHFyAQ9GXTmlk7MjcgQbDCx6mn7yIawsppWk
# vfPkKaAQsiqaT9DnMWBHVNIabGqgQSGTrQWo43MOfsPynhbz2Hyxf5XWKZpRvr3d
# MapandPfYgoZ8iDL2OR3sYztgJrbG6VZ9DoTXFm1g0Rf97Aaen1l4c+w3DC+IkwF
# kvjFV3jS49ZSc4lShKK6BrPTJYs4NG1DGzmpToTnwoqZ8fAmi2XlZnuchC4NPSZa
# PATHvNIzt+z1PHo35D/f7j2pO1S8BCysQDHCbM5Mnomnq5aYcKCsdbh0czchOm8b
# kinLrYrKpii+Tk7pwL7TjRKLXkomm5D1Umds++pip8wH2cQpf93at3VDcOK4N7Ew
# oIJB0kak6pSzEu4I64U6gZs7tS/dGNSljf2OSSnRr7KWzq03zl8l75jy+hOds9TW
# SenLbjBQUGR96cFr6lEUfAIEHVC1L68Y1GGxx4/eRI82ut83axHMViw1+sVpbPxg
# 51Tbnio1lB93079WPFnYaOvfGAA0e0zcfF/M9gXr+korwQTh2Prqooq2bYNMvUoU
# KD85gnJ+t0smrWrb8dee2CvYZXD5laGtaAxOfy/VKNmwuWuAh9kcMIIGbTCCBNWg
# AwIBAgIRAIHZB+Gthrz4Gk95ksYTJmkwDQYJKoZIhvcNAQEMBQAwVDELMAkGA1UE
# BhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDErMCkGA1UEAxMiU2VjdGln
# byBQdWJsaWMgQ29kZSBTaWduaW5nIENBIFIzNjAeFw0yMTExMjkwMDAwMDBaFw0y
# MjExMjkyMzU5NTlaMFYxCzAJBgNVBAYTAlVTMREwDwYDVQQIDAhNYXJ5bGFuZDEZ
# MBcGA1UECgwQSHVudGVyIEtpbWJyb3VnaDEZMBcGA1UEAwwQSHVudGVyIEtpbWJy
# b3VnaDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBANOrEDZASM3m7QBr
# A7oFzMrdtbLL/hfJrutOVa9y9P7nEGXv5PRJJQZnj1DRTnS2p/nwYqeQHZX8N0dr
# CZUXnxyeOpQhuTcBMKnFL7R5v3SOQxCYJQXFoIlIC1WHQ2n9N55JC0+4NYO+Tp8C
# ozdNK/SOGB7vvF2EZYj0t0fV2EpxoqleKVfBv4AbJim/C6RUUjMRyzDVwklXi0Fr
# db8qrDMpEn0W38DZvqMUo/QUeZB0zntgGZ0ess1Riq9i6eBtbzgC5HVxwsKUTqAe
# CiQ3E59BfQgDzZ3iLv+EaQOxxLrT4V34v9AxLfeatNAjiTqceaAx/wsDDIxgbpGS
# D7vVZEFn3j1xdRa2U+AvXE6+T3vtcZWGugjoRXHCDZAXcd8Vzjm9Yn4FCRQ/Pccn
# 3YILU+cesimuCbdelDy0nEbzbgmIIAWWoAqcKftvXjhVWSB/O25mD7nZ1pIgwn07
# uXgNbvfAMr2nckRmWuvwzbYbFP+iMnfg31iyoPUG54TjO2uYvPfeqZHh3WkNwcv4
# vGau7cuCPKaxDj+WQk+HpOz8znqWt3Ek73JVNjvP+Rr8BcZSOrOhuT30f6ScOouZ
# FY6PV0PTBVmGAkvg4rpN7LWiBEBiC/y8eCUFMYuA1ITtrSiYaOabl+WZZ72KQxYq
# 0wguynjwHdj5qqiRGZdce38VuTnzAgMBAAGjggG2MIIBsjAfBgNVHSMEGDAWgBQP
# Kssghyi47G9IritUpimqF6TNDDAdBgNVHQ4EFgQUu0ZxoA6rDE2JeSEFZdIoYeLc
# jTcwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwEwYDVR0lBAwwCgYIKwYB
# BQUHAwMwEQYJYIZIAYb4QgEBBAQDAgQQMEoGA1UdIARDMEEwNQYMKwYBBAGyMQEC
# AQMCMCUwIwYIKwYBBQUHAgEWF2h0dHBzOi8vc2VjdGlnby5jb20vQ1BTMAgGBmeB
# DAEEATBJBgNVHR8EQjBAMD6gPKA6hjhodHRwOi8vY3JsLnNlY3RpZ28uY29tL1Nl
# Y3RpZ29QdWJsaWNDb2RlU2lnbmluZ0NBUjM2LmNybDB5BggrBgEFBQcBAQRtMGsw
# RAYIKwYBBQUHMAKGOGh0dHA6Ly9jcnQuc2VjdGlnby5jb20vU2VjdGlnb1B1Ymxp
# Y0NvZGVTaWduaW5nQ0FSMzYuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5z
# ZWN0aWdvLmNvbTAYBgNVHREEETAPgQ1hZG1pbkBkZXYwLnNoMA0GCSqGSIb3DQEB
# DAUAA4IBgQAKXhJLtbPJ3JlRNAAnPXhsK7ZIbMQfCCLKdGFqHn9xiQGJivhWM0sm
# n+fD1QTPjrmxDeaiFIUs2e+KSqRKXOafM1163kYrF4Hd+HVF+aISRSVYb/Tr4wPj
# 40LLGCIzfI9a8j5wmxE7NDRMAa4Y3zaVnkTLE3iOc0bKh0aX2dzyUA882msrg8Pc
# EmJ+ojILl3Ur0B4Zt+z6lzbLCzN3ONvwpKC8GJmzvx5RHjcl40+HX5AXyeMF4F/j
# 3RP/zNolNMa9t1BOQbqPFx9EcvgWZSriVuQrzyyMhfUeDteY5IGdTxL7CTA1Di1G
# CqwYW++rfP+KbDMZfbVNpHqNMuBOlS5J7hCQ/QZW/1ueE1ms0vsgQgxnKz8fE4NB
# RhWxUN8qkTsXQnaF8afc/dIfDQXhGlUMNQCiolfYOIpmgw4B7AFdMYkKI3rMxKKA
# tEVioBQ4lByxdeANL/LedsGXDlD2NlqNZ1tpS8/7MiSLQGNNOOa0+PYauhRbTHyG
# xL9lpId+xrExggMKMIIDBgIBATBpMFQxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9T
# ZWN0aWdvIExpbWl0ZWQxKzApBgNVBAMTIlNlY3RpZ28gUHVibGljIENvZGUgU2ln
# bmluZyBDQSBSMzYCEQCB2QfhrYa8+BpPeZLGEyZpMAkGBSsOAwIaBQCgeDAYBgor
# BgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEE
# MBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBTA
# fQotLkhNRdbshkImiGxU55kxAjANBgkqhkiG9w0BAQEFAASCAgCC/0g9xd3Nxas7
# CyUW2O3XZgk+oQhlWbn3oghsTowOo1FbScoIKkpOU7b9SOzr4OU7QGm1rXDk5zdA
# qni0HFF/Ints/wRKq+bF1jmDJ9Nj7yjRTzItU6MsrF/75s7hYcQhE5z2n5tjBr50
# a40A7k2K/DJi+Io7rIsCwg9vlIzCfInCMlTMgUYSggM5T1gjpOpa5wel/nqVqv6F
# DBAgt+xMtKd1Y7w/GrWn3Byyh/ct27DiW+4GbjKAEwZTgGkpGbcT6jT9xlciJIjP
# 7rRO+Ly/wm6JYx8b7QtbkuD2Z6GNRNshyaCA6P2wLwKAquV8RcWVlORkeL5U53p6
# +YH+MqxXvjsHdZ5h8cyatzozSsJkFIX2QJQtL4PDqxR/ugio709QHHAdvdWbckuA
# IidW4k+1cmQj3qQI9H5zEEknevXYX3mTK66320nl259mANmxSnc+CZck7hVo0htT
# 7/d+s4edZ2G6C7eFfFF0pPxr+Zq8Qd5OS0Kw6h5WbaP5LV0xnzS+iyP2g+bUL0pu
# G5ME3hZ8AnCeklc0lKQp3Vca92MDducKbn6FqLolf/z1q2ec1ctIzyCkYq1Vh2XL
# Oic/esFY1sUhkFbT3XJYrYlRMQX16ettVq59iSPV6qBQabR0EWanjtt5BUvwoSwZ
# Pfji2b/6AxDPcD1Ud37Jxu84YIVQlA==
# SIG # End signature block
