<#
.SYNOPSIS
 Gather and upload specifications of a Windows host to rTechsupport
.DESCRIPTION
 Use various native powershell and wmic functions to gather verbose information on a system to assist troubleshooting
.OUTPUTS Specs
  '..\TechSupport_Specs.html'
#>

param (
    [Switch]$run,
    [Switch]$view,
    [Switch]$upload
)

# Working directory should be \files
. .\src\config.ps1
. .\src\outfunctions.ps1
. .\src\notes.ps1
. .\src\functions.ps1
. .\src\data.ps1

# ------------------ #
# Write da file
header | Out-File -Encoding ascii $file
getDate | Out-File -Append -Encoding ascii $file
getBasicInfo | Out-File -Append -Encoding ascii $file
getNotes | Out-File -Append -Encoding ascii $file
table | Out-File -Append -Encoding ascii $file
getHardware | Out-File -Append -Encoding ascii $file
getMonitors | Out-File -Append -Encoding ascii $file
getBattery | Out-File -Append -Encoding ascii $file
getRAM | Out-File -Append -Encoding ascii $file
#getLicensing | Out-File -Append -Encoding ascii $file
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
getBrowserExtensions | Out-File -Append -Encoding ascii $file
getNets | Out-File -Append -Encoding ascii $file
getTcpSettings | Out-File -Append -Encoding ascii $file
getNetConnections | Out-File -Append -Encoding ascii $file
getDrivers | Out-File -Append -Encoding ascii $file
getAudio | Out-File -Append -Encoding ascii $file
getDisks | Out-File -Append -Encoding ascii $file
getSmart | Out-File -Append -Encoding ascii $file
getHosts | Out-FIle -Append -Encoding ascii $file
getTimer | Out-File -Append -Encoding ascii $file
promptUpload
# ------------------ #

