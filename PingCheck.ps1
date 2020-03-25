#! /bin/powerhshell

<#

.SYNOPSIS
This is a powerhsell script for logging speed and latency over an indefinite period of time when ran in the foreground. It uses a remote server (1.1.1.1 by default) and an interval (2 seconds) by default

.DESCRIPTION
-remote <remote server IP or name>
-interval <interval in seconds>
-speedtest ($true|$false)

.EXAMPLE
./PingCheck.ps1 -remote dev0.sh -interval 5

.NOTES
No notes at the moment

.LINK
https://git.dev0.sh/piper/techsupport_scripts

#>

#get things
param ([string]$remote='1.1.1.1', [decimal]$interval='2', [bool]$speedtest=$false)
$gate = $($(Get-NetIPConfiguration).IPv4DefaultGateway).NextHop

# make our file
Add-Content -Path $env:USERPROFILE\Desktop\PingCheck_$interval.csv  -Value "Time,$gate,$remote,Speed"

Write-Host "All output is being made in $env:USERPROFILE\Desktop\PingCheck_$interval.csv. Do not open the file until you have terminated this script." -Foreground Green

While ($true) {
	$time = Get-Date -Format "%M/%d %H:%m:%s"
	$timeRemote = $(Test-Connection -ComputerName $remote -Count 1).ResponseTime
	$timeGate = $(Test-Connection -ComputerName $gate -Count 1).ResponseTime
	If ($speedtest){
			$speed = $($a=Get-Date; Invoke-WebRequest https://dev0.sh/1MiB |Out-Null; "$((10/((Get-Date)-$a).TotalSeconds)*8) Mbps")
		}Else{
			$speed = "skipped"
	}
	Add-Content -Path $env:USERPROFILE\Desktop\PingCheck_$interval.csv -Value "$time,$timeGate,$timeRemote,$speed"
	Start-Sleep -Seconds $interval
}
