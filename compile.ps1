ps2exe -inputFile .\Get-Specs.ps1 -outputFile .\Get-Specs.exe -x64 -title 'rTechsupport Spec Tool' -company 'Dev0' -description 'Developed by PipeItToDevNull for rTechsupport' -requireAdmin -iconFile SNOO_256.ico

& 'C:\Program Files (x86)\Windows Kits\10\bin\10.0.22000.0\x64\signtool.exe' sign /fd certHash /sha1 ACD87F7650858831D787D734D635CC32705166D2 /t http://timestamp.digicert.com .\Get-Specs.exe
