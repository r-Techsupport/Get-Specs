powershell -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass; Start-Process PowerShell -Verb RunAs -ArgumentList ('-NoExit',('cd ''{0}'';' -f $pwd.ProviderPath),'.\Get-Specs.ps1')"

