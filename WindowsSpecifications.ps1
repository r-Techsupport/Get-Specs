<#
.SYNOPSIS
  Gather and upload specifications of a Windows host to rTechsupport
.DESCRIPTION
  Use various native powershell and wmic functions to gather verbose information on a system to assist troubleshooting
.OUTPUTS Specs
  '.\TechSupport_Specs.txt'
.NOTES
  Version:        .1
  Author:         PipeItToDevNull
  Creation Date:  12/29/2020
  Purpose/Change: Initial script development
#>

# ------------------ #
$file = 'TechSupport_Specs.txt'
# ------------------ #

function startItOut {
	Get-Date
}
function basicInfo {
	$1 = 'Edition: ' + $(Get-Item "HKLM:Software\Microsoft\Windows NT\CurrentVersion").GetValue("ProductName")
	$2 = 'Build: ' + $(Get-Item "HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion").GetValue("ReleaseID")
	Return $1,$2
}
function getCPU{
    $CPUInfo = Get-WmiObject Win32_Processor
    $CPU = $CPUInfo.Name
	$1 = "`n" + 'CPU: ' + $CPU
	Return $1
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
    $GPUname = $GPUbase.Name
    $GPU= $GPUname + " at " + $GPUbase.CurrentHorizontalResolution + "x" + $GPUbase.CurrentVerticalResolution
    $1 = "Graphics Card: " + $GPU
	Return $1
}
function getRAM {
	$1 = "`n" + "RAM: " + $(Get-WMIObject -class Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)}) + 'GB'
	$2 = $(Get-WmiObject win32_physicalmemory | Format-Table Manufacturer,Configuredclockspeed,Devicelocator,Capacity,Serialnumber -autosize)
	Return $1,$2
}
function getUpdates {
	$1 = "`n" + "Installed updates:"
	$2 = Get-HotFix |format-table -auto Description,HotFixID,InstalledOn
	Return $1,$2
}
function getStartup {
    $startBase = Get-CimInstance Win32_StartupCommand
    $1 = "Startup Tasks for user: "
	$2 = $startBase.Caption
	Return $1,$2
}
function getProcesses {
    $procBase = Get-Process
    $procTrash = $procBase.ProcessName
    $1 = "`n" + "Running processes: "
	$2 = $($procTrash | select -Unique)
	return $1,$2
}
function getServices {
	$1 = "`n" + "Services: "
	$2 = $(Get-Service | Format-Table -auto)
	Return $1,$2
}
function getInstalledApps {
    $installedBase = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*
    $1 = "Installed Apps: "
	$2 = $installedBase.DisplayName
	Return $1,$2
}
function getDisks {
	$1 = "`n" + "Disk layouts: "
	$2 = $(Get-Partition|format-table -auto)
	$3 = $(Get-Volume|format-table -auto)
	Return $1,$2,$3
}
function getNets {
	$1 = "`n" + "Network adapters:"
	$2 = $(Get-NetAdapter|format-list Name,InterfaceDescription,Status,LinkSpeed)
	$3 = $(Get-NetIPAddress|format-table -auto IpAddress,InterfaceAlias,PrefixOrigin)
	Return $1,$2,$3
}
function getDrivers {
	$1 = "`n" + "Drivers and device versions: "
	$2 = $(gwmi Win32_PnPSignedDriver | format-table -auto devicename,driverversion)
	Return $1,$2
}
function uploadFile {
	$link = Invoke-WebRequest -ContentType 'text/plain' -Method 'PUT' -InFile $file -Uri "https://share.dev0.sh/upload/$null.txt" -UseBasicParsing
	set-clipboard $link.Content
}
function promptUser {
	. .\new-wpfMessageBox.ps1
	$Params = @{
		Content = "Do you want to view or upload the specs?"
		Title = "rTechsupport "
		TitleBackground = "DeepSkyBlue"
		TitleFontSize = 28
		TitleFontWeight = "Bold"
		TitleTextForeground = "White"
		ContentFontSize = 12
		ContentFontWeight = "Medium"
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
startItOut > $file
basicInfo >> $file
getCPU >> $file
getMobo >> $file
getGPU >> $file
getRAM >> $file
getUpdates >> $file
getStartup >> $file
getProcesses >> $file
getServices >> $file
getInstalledApps >> $file
getDisks >> $file
getNets >> $file
getDrivers >> $file
# ------------------ #

# ------------------ #
promptUser
# ------------------ #