$admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

Start-Transcript "TechSupport_Speccy.txt"
Get-Date
Write-Host "`n" -NoNewline

function Get-OS{
    $OSInfo = Get-WmiObject Win32_OperatingSystem
    $OS = $OSInfo.Version
    return $OS
}
Write-Host "OS: " -NoNewline
Get-OS
Write-Host "`n" -NoNewline

function Get-CPU{
    $CPUInfo = Get-WmiObject Win32_Processor
    $CPU = $CPUInfo.Name
    return $CPU
}
Write-Host "CPU: " -NoNewline
Get-CPU
Write-Host "`n" -NoNewline

function Get-Temperature {
    $t = Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi" -ErrorAction SilentlyContinue
    $returntemp = @()
    if ($t){
        foreach ($temp in $t.CurrentTemperature) {
            $currentTempKelvin = $temp / 10
            $currentTempCelsius = $currentTempKelvin - 273.15

            $currentTempFahrenheit = (9/5) * $currentTempCelsius + 32

            $returntemp += $currentTempCelsius.ToString() + " C : " + $currentTempFahrenheit.ToString() + " F : " + $currentTempKelvin + "K"  
        }
    }
    else {
        $returntemp = "Not supported"
        }
    return $returntemp
}
Write-Host "CPU Temperature: " -NoNewline
Get-Temperature
Write-Host "`n" -NoNewline

function Get-Mobo{
    $moboBase = Get-WmiObject Win32_BaseBoard
    $moboMan = $moboBase.manufacturer
    $moboMod = $moboBase.product
    $mobo = $moboMan + " | " + $moboMod
    return $mobo
}
Write-Host "Motherboard: " -NoNewline
Get-Mobo
Write-Host "`n" -NoNewline

function Get-GPU {
    $GPUbase = Get-WmiObject Win32_VideoController
    $GPUname = $GPUbase.Name
    $GPU= $GPUname + " at " + $GPUbase.CurrentHorizontalResolution + "x" + $GPUbase.CurrentVerticalResolution
    return $GPU
}
Write-Host "Graphics Card: " -NoNewline
Get-GPU
Write-Host "`n" -NoNewline

function Get-Startup {
    $startBase = Get-CimInstance Win32_StartupCommand
    $startNames = $startBase.Caption
    return $startNames
}
Write-Host "Startup Tasks for user: "
Get-Startup
Write-Host "`n" -NoNewline

function Get-Processes {
    $procBase = Get-Process
    $procTrash = $procBase.ProcessName
    $procClean = $procTrash | select -Unique
    return $procClean
}
Write-Host "Running processes: "
Get-Processes
Write-Host "`n" -NoNewline

Write-Host "Services: " -NoNewline
Get-Service | Format-Table

if ($admin -eq $true) {
function Get-SMART {
    $smartBase = gwmi -namespace root\wmi -class MSStorageDriver_FailurePredictStatus
    $smartValue = $smartBase | Select InstanceName, PredictFailure | Format-Table
    return $smartValue
}
Write-Host "Basic SMART: " -NoNewline
Get-SMART
}
Stop-transcript

$FilePath = '.\TechSupport_Speccy.txt'
Invoke-WebRequest -ContentType 'text/plain' -Method 'PUT' -InFile $FilePath -Uri 'https://share.dev0.sh/upload' -UseBasicParsing