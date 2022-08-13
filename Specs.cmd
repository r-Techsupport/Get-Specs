@echo off

IF NOT EXIST ".\files" (
    msg.exe * "The files directory is missing. This most likely means you did not extract the zip file properly. Right click the zip file and choose 'Extract All' then run the new exe."
    exit
)

powershell Start-Process PowerShell -Verb RunAs -ArgumentList ('Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass; ',('cd ''{0}'';' -f $pwd.ProviderPath),'.\Get-Specs.ps1')

