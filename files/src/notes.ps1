# Notes is giant, so it probably deserves its own section

function getNotes {
    Write-Host 'Checking for notes...'
    $1 = '<h2>Notes</h2>'
    $2 = @()
    $3 = @()
    $4 = @()
    $5 = @()
    # bad software 
   foreach ($software in $badSoftware) { 
        if ($installedBase.DisplayName -Like $software) { 
            $software = $software.Replace('*','')
            $2 += "Installed: " + $software 
        }
    }
    # bad startups
    foreach ($start in $badStartup) { 
        if ($startUps -Like $start) { 
            $3 += "Startup: " + $start 
        }
    }
    # bad processes
    foreach ($running in $badProcesses) {
        if ($runningProcesses.Name -Like $running) {
            $4 += "Process: " + $running 
        } 
    }
    # bad registry values
    $i = 0
    foreach ($reg in $badKeys) {
        $value = $null
        If (Test-Path -Path $badKeys[$i]) {
            $data = Get-ItemProperty -Path $badKeys[$i] -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $badValues[$i] -ErrorAction SilentlyContinue
        If ($data -eq $badData[$i]) {
                $5 += $badRegExp[$i] + $data
            }
        }
        $i = $i + 1
    }
    # bad hostnames
    if ($badHostnames -contains $cimOS.CSName) {
        $6 = "This may be a malicious distribution of Windows: " + $cimOS.CSName
    }
    # bad diskspace
    $c = $volumes | ? { $_.DriveLetter -eq 'C' }
    $cAllowable = $c.Size - $c.Size * .20
    $cConsumed = $c.Size - $c.SizeRemaining
    If ($cConsumed -gt $cAllowable) {
        $7 = "Less than 20% left on C"
    }
    # bad productKeys
    # If ($masKeys -contains $(getFullKey)) {
        # $8 = "MAS detected"
    # }
    # bad dns suffixes
    If ('utopia.net' -in $dns.SuffixSearchList) {
        $9 = "Utopia malware found, router is infected. This adds a 'utopia.net' suffix to your DNS and can cause issues with shortnames."
    }
    If ($eol) {
        $10 = "OS was EOL on " + $eolOn + " Build is " + $osBuild
    }
    # bad smart things
    If (-Not $smart) {
       $11 = "SMART timeout. Manually run Crystal Disk Info to determine your failed disk(s)" 
    }
    $12 = @()
    Foreach ($disk in $smart) {
        If ($disk.'Reallocated Sectors Count') {
            If ($disk.'Reallocated Sectors Count' -ne '000000000000') {
                $12 += "Reallocated sector on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Reallocated Sectors Count'
            }
        }
        If ($disk.'Current Pending Sector Count') {
             If ($disk.'Current Pending Sector Count' -ne '000000000000') {
                 $12 += "Current Pending Sector Count on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Current Pending Sector Count'
            }
        }
        If ($disk.'Uncorrectable Sector Count') {
            If ($disk.'Uncorrectable Sector Count' -ne '000000000000') {
                $12 += "Uncorrectable Sector Count on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Uncorrectable Sector Count'
            }
        }
        If ($disk.'Command Timeout') {
            If ($disk.'Command Timeout' -ne '000000000000') {
                $12 += "Command Timeout on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Command Timeout'
            }
        }
        If ($disk.'Reported Uncorrectable Errors') {
            If ($disk.'Reported Uncorrectable Errors' -ne '000000000000') {
                $12 += "Reported Uncorrectable Errors on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'Reported Uncorrectable Errors'
            }
        }
        If ($disk.'CRC Error Count') {
            If ($disk.'CRC Error Count' -ne '000000000000') {
                $12 += "CRC Error Count on " + $disk.'Drive Letter' + " " + $disk.Model + " is " + $disk.'CRC Error Count'
            }
        }
        If ($disk.'Rotation Rate' -NotLike '---- (SSD)' -And $disk.'Drive Letter' -eq 'C:') {
            $13 += "C: is on an HDD"
        }
    }
    # check for VPNs being connected
    $cAdapters = $netAdapters | Where {$_.MediaConnectionState -eq 'Connected'}
    ForEach ($adapter in $badAdapters) { 
        If ($cAdapters.IfDesc -Like $adapter) { 
            $14 = "VPN is connected"
        }
    }
    # check for specific files in the FS
    $15 = @()
    ForEach ($file in $badFiles) {
        If (Test-Path $file) {
            $14 += "Bad file found $file"
        }
    }
    # check if the user configured a static number of cores in msconfig
    If ($bcdedit -ne $NULL) {
        $16 = "Static core number is set in msconfig"
    }
    # ram checks
    If ($ramSticks.count -eq 2) {
        If ($ramSticks[0].Devicelocator -eq 'DIMM1' -And $ramSticks[1].Devicelocator -eq 'DIMM2') {
            $17 = "RAM is in DIMM1 and DIMM2"
        }
        If ($ramSticks[0].Devicelocator -eq 'DIMM3' -And $ramSticks[1].Devicelocator -eq 'DIMM4') {
            $17 = "RAM is in DIMM3 and DIMM4"
        }
    }
    # hosts checks 
    If ($hostsSum -ne $hostsHash) {
        $18 = "Hosts file has been modified from stock, it has been appended to bottom of this page"
        If ($hostsContent -Like "*license.piriform.com*") {
            $19 = "piriform license server redirection"
        }
        If ($hostsContent -Like "*Spybot*") {
            $20 = "spybot has edited hosts file"
        }
    }
    # check for dumps
    $dmpFiles = Get-ChildItem 'C:\Windows\Minidump' -ErrorAction SilentlyContinue
    $lastWeek = $today.AddDays(-7)
    $count = 0
    ForEach ($dmp in $dmpFiles) {
        If ($dmp.LastWriteTime -gt $lastWeek) {
            $count++
        }
    }
    If ($count -gt 0) {
        $21 = "Found $count dumps"
    }
    # check for missing files in FS
    If (!(Test-Path $microCode[0]) -And !(Test-Path $microCode[1])) {
        $22 += "Microcode fixes are missing or were removed by malicious 'fixers'"
    }
    # count and report issue devices
    If ($issueDevices.Status.Count -gt 0) {
        $issueDeviceCount = $issueDevices.Status.Count
        $23 = "$issueDeviceCount devices have issues, see 'Devices with issues' section"
    }
    # Check for TPM and secure boot if on Windows 11
    If ($cimOS.Caption -Like "Microsoft Windows 11*") {
        # Why must SpecVersion be a string :(
        If ($tpm -eq $NULL) {
            $24 = "Windows 11 with no TPM"
        }
        ElseIf ([int][string]($tpm.SpecVersion[0]) -lt 2) {
            $24 = "Windows 11 TPM version not satisfied"
        }

        If ($secureBoot -eq $NULL) {
            $25 = "Windows 11 with secure boot not enabled"
        }
    }
    # Verify C and the ESP are on the same disk
    If ($bootMode -eq "UEFI" -and ($NULL -eq $(Get-Partition -DiskNumber $(Get-Partition | ? {$_.DriveLetter -eq "C" }).DiskNumber | ? {$_.GptType -eq "{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}"}))) {
        $26 = "C and the ESP are not on the same disk"
    }

    # Check for MBR disk, for Arc
    $27 = @()
    Get-Disk | Select-Object PartitionStyle, FriendlyName | ? PartitionStyle -eq "MBR" | %{ $27 += "$($_.FriendlyName) is MBR" }


    # check for power profiles that indicate 'custom' OS
    $28 = @()
    foreach ($profile in $badPower) { 
        if ($powerProfiles.ElementName -Like $profile) { 
            $28 += "Power Profile: " + $profile
        }
    }

    Write-Host 'Checked for notes' -ForegroundColor Green
    Return $1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27,$28
}

