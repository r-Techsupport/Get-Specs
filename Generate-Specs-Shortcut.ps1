<#
.SYNOPSIS
 Generates windows shortcut files (.lnk) with appropriate target and argument details 
.DESCRIPTION
 This will generate a windows shortcut in the form of FILENAME.lnk at the PATH you specify. Additionally, this shortcut will pass along the current working directory to the newly spawned elevated powershell session, and switches the default directory of the Administrative session (C:\Windows\System32) to the current working directory.
 
 It will execute the target script which is currently defined as ".\powershell.ps1" with administrative privilges using the RunAs Verb parameter
 
 NOTE: Once the shortcut has been created, you cannot edit the shortcut Target with Windows Explorer due to an ancient explorer path length limitation of 260 characters. This does not affect how the shortcut gets executed however.
#>

# Set the PowerShell target path.
$poshPath = '%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe'

# Set the Command Arguments for the above PowerShell Target command. It's golf-y because my initial plan was to get this shortened to less than 260 characters before I discovered a method to generate shortcuts without the 260 character limitation.
$argCmdVars = '-Comm "$nvg = @{Sco = ''Global''; Opt = ''AllScope''}; nv -Name teDir -val $pwd @nvg; nv -Name teArg -val ""-NoPro -Exec Bypass -NoEx -Comm "cd $teDir; .\powershell.ps1"""; start PowerShell -Verb RunAs -Work $teDir -Arg $teArg"'

# Defines a new WScript Object and stores in $WObj
$WObj = New-Object -ComObject WScript.Shell

# Creates a new shortcut named $poshShortcut and stores it in the path of the current working directory using $pwd, and will be named accordingly. Adjust as needed
$poshShortcut = $WObj.CreateShortcut("$pwd\posh-shortcut.lnk")
$poshShortcut.TargetPath = $poshPath
$poshShortcut.Arguments = $argCmdVars
$poshShortcut.Save()
