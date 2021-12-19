# set execution policy to allow all our children to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Chainload our primary script
./Get-Specs.ps1
