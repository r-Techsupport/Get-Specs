# set execution policy to allow all our children to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

# Run the basic checks before doing anything
function promptFileIssue {
    If (Test-Path ./files) {
        } Else {
            msg.exe * "The files directory is missing. This most likely means you did not extract the zip file properly.

Right click the zip file and choose 'Extract All' then run the new exe."
            Throw
    }
}
function loadMain {
    $validation = Get-AuthenticodeSignature .\Get-Specs.ps1
    $cert = $validation | Select -ExpandProperty SignerCertificate
    If ($validation.Status -ne 'Valid') {
        msg.exe * "Get-Specs.ps1 has an invalid signature"
        Throw
    }
    If ($cert.Thumbprint -ne 'ACD87F7650858831D787D734D635CC32705166D2') {
        msg.exe * "Get-Specs.ps1 has an invalid signature"
        Throw
    } Else {
        ./Get-Specs.ps1
    }
}
promptFileIssue
loadMain
