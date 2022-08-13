powershell Start-Process PowerShell -Verb RunAs -ArgumentList ('Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass; ',('cd ''{0}'';' -f $pwd.ProviderPath),'.\Get-Specs.ps1')

