# set execution policy to allow all our children to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Run the basic checks before doing anything
function promptFileIssue {
    If (Test-Path ./files) {
        } Else {
            msg.exe * "The files directory is missing. This most likely means you did not extract the zip file properly.

Right click the zip file and choose 'Extract All' then run the new exe."
            Throw
    }
}
promptFileIssue

# Chainload our primary script
./Get-Specs.ps1
