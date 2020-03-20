#! /bin/powerhshell

<#

.SYNOPSIS
This is a powerhsell script for logging speed and latenvy over an indefinite period of time when ran in the foreground. It uses a remote server (1.1.1.1 by default) and an interval (10 seconds) by default

.DESCRIPTION
-remote <remote server IP or name>
-interval <interval in seconds>

.EXAMPLE
./PingCheck.ps1 -remote dev0.sh -interval 5

.NOTES
No notes at the moment

.LINK
https://git.dev0.sh/piper/techsupport_scripts

#>

param ([string]$remote='1.1.1.1', [decimal]$interval='10')

While ($true) {
	Write-Output "--------------------"
	Get-Date -Format "%M/%d %H:%m:%s"
	Test-Connection -ComputerName $remote -Count 1 | Format-Table Address,ResponseTime
	Write-Output "Speed is $($a=Get-Date; Invoke-WebRequest https://dev0.sh/1MiB |Out-Null; "$((10/((Get-Date)-$a).TotalSeconds)*8) Mbps")"
	Start-Sleep -Seconds $interval
}
