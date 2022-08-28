function getDate {
    $1 = "Local: " + $today
    $2 = "UTC: " + $([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId(($today), 'Greenwich Standard Time'))
    $3 = "Version: " + $version
    Return $1,$2,$3
}
function getbasicInfo {
    Write-Host 'Getting basic information...'
    $1 = '<h2>System Information</h2>'
    $bootuptime = $cimOs.LastBootUpTime
    $uptime = $today - $bootuptime
    $2 = 'Edition: ' + $cimOs.Caption 
    $3 = $osBuild
    $4 = 'Install date: ' + $cimOs.InstallDate
    $5 = 'Uptime: ' + $uptime.Days + " Days " + $uptime.Hours + " Hours " +  $uptime.Minutes + " Minutes"
    $6 = 'Hostname: ' + $cimOs.CSName
    $7 = 'Domain: ' + $env:USERDOMAIN
    $8 = 'Boot mode: ' + $bootMode
    $9 = 'Boot state: ' + $bootupState
    Write-Host 'Got basic information' -ForegroundColor Green
    Return $1,$2,$3,$4,$5,$6,$7,$8,$9
}
function getFullKey {
    $keyOutput=""
    $keyOffset = 52

    $IsWin10 = ([System.Math]::Truncate($key[66] / 6)) -band 1
    $key[66] = ($key[66] -band 0xF7) -bor (($IsWin10 -band 2) * 4)
    $i = 24
    $maps = "BCDFGHJKMPQRTVWXY2346789"
    do {
        $current= 0
        $j = 14
        do {
            $current = $current* 256
            $current = $key[$j + $keyOffset] + $current
            $key[$j + $keyOffset] = [System.Math]::Truncate($current / 24 )
            $current=$current % 24
            $j--
        }
        while ($j -ge 0)
            $i--
            $keyOutput = $maps.Substring($current, 1) + $keyOutput
            $last = $current
    } while ($i -ge 0)

    if ($IsWin10 -eq 1) {
        $keypart1 = $keyOutput.Substring(1, $last)
        $keypart2 = $keyOutput.Substring($last + 1, $keyOutput.Length - ($last + 1));
        $keyOutput = $keypart1 + "N" + $keypart2;
    }

    if ($keyOutput.Length -eq 25) {
        $1 = [String]::Format("{0}-{1}-{2}-{3}-{4}",
            $keyOutput.Substring(0, 5),
            $keyOutput.Substring(5, 5),
            $keyOutput.Substring(10,5),
            $keyOutput.Substring(15,5),
            $keyOutput.Substring(20,5))
    } else {
        $keyOutput
    }
    return $1
}
function getTemps {
    Try {
        Add-Type -Path .\OpenHardwareMonitorLib.dll -ErrorAction SilentlyContinue
    }
    catch {
        $1 = ' OHMF'
        $2 = ' OHMF'
        return $1,$2
    }
    $ohm = New-Object -TypeName OpenHardwareMonitor.Hardware.Computer
    $ohm.CPUEnabled= 1;
    $ohm.GPUEnabled = 1;
    $ohm.Open();
    foreach ($comp in $ohm.Hardware) {
        if($comp.HardwareType -eq [OpenHardwareMonitor.Hardware.HardwareType]::CPU){
        $comp.Update()
            foreach ($sens in $comp.Sensors) {
                 if ($sens.SensorType -eq [OpenHardwareMonitor.Hardware.SensorType]::Temperature) {
                    $1 = $sens.Value.ToString()
                    #$sens.Identifier for a better name
                 }
            }
        }
    }
    foreach ($comp in $ohm.Hardware) {
        if($comp.HardwareType -ne [OpenHardwareMonitor.Hardware.HardwareType]::CPU){
        $comp.Update()
            foreach ($sens in $comp.Sensors) {
                 if ($sens.SensorType -eq [OpenHardwareMonitor.Hardware.SensorType]::Temperature) {
                    $2 = $sens.Value.ToString() 
                    # $sens.Identifier for a better name
                 }
            }
        }
    }
    $ohm.Close();
    Return $1,$2
}
$temps = getTemps
function getHardware {
    Write-Host 'Getting harwdware information...'
    $1 = "<h2 id='hw'>Hardware Basics</h2>"
    $cpu = Get-WmiObject Win32_Processor
    $mobo = Get-WmiObject Win32_BaseBoard
    $gpu = Get-WmiObject Win32_VideoController

    $hwArray = @()

    $cpuObject = New-Object PSobject
    Add-Member -InputObject $cpuObject -MemberType NoteProperty -Name "Part" -Value "CPU"
    Add-Member -InputObject $cpuObject -MemberType NoteProperty -Name "Manufacturer" -Value $cpu.Manufacturer
    Add-Member -InputObject $cpuObject -MemberType NoteProperty -Name "Product" -Value $cpu.Name
    Add-Member -InputObject $cpuObject -MemberType NoteProperty -Name "Temperature" -Value $temps[0]
    Add-Member -InputObject $cpuObject -MemberType NoteProperty -Name "Driver" -Value ""
    $hwArray += $cpuObject

    $moboObject = New-Object PSObject
    Add-Member -InputObject $moboObject -MemberType NoteProperty -Name "Part" -Value "Motherboard"
    Add-Member -InputObject $moboObject -MemberType NoteProperty -Name "Manufacturer" -Value $mobo.Manufacturer
    Add-Member -InputObject $moboObject -MemberType NoteProperty -Name "Product" -Value $mobo.Product
    Add-Member -InputObject $moboObject -MemberType NoteProperty -Name "Driver" -Value ""
    $hwArray += $moboObject
 
    If (Get-Member -inputobject $gpu -name "Count" -Membertype Properties) {
        $i = 0
        foreach ($g in $gpu) {
            $gpuDriver = $(gwmi Win32_PnPSignedDriver | ? { $_.deviceName -eq $gpu[$i].Name }).driverversion
            $gpuObject = New-Object PSObject
            Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Part" -Value "Video Card"
            Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Manufacturer" -Value $gpu[$i].AdapterCompatibility
            Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Product" -Value $gpu[$i].Name
            Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Temperature" -Value $temps[1]
            Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Driver" -Value $gpuDriver
            $hwArray += $gpuObject
            $i = $i + 1
        }
    } Else {
        $gpuDriver = $(gwmi Win32_PnPSignedDriver | ? { $_.deviceName -eq $gpu.Name }).driverversion
        $gpuObject = New-Object PSObject
        Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Part" -Value "Video Card"
        Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Manufacturer" -Value $gpu.AdapterCompatibility
        Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Product" -Value $gpu.Name
        Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Temperature" -Value $temps[1]
        Add-Member -InputObject $gpuObject -MemberType NoteProperty -Name "Driver" -Value $gpuDriver
        $hwArray += $gpuObject
    }

    $2 = $hwArray | ConvertTo-Html -Fragment

    Return $1,$2
}
function getRAM {
    $totalRam = $(Get-WMIObject -class Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)})
    $ramObject = New-Object PSObject
    Add-Member -InputObject $ramObject -MemberType NoteProperty -Name "Total" -Value $totalRam
    Add-Member -InputObject $ramObject -MemberType NoteProperty -Name "RAM" -Value "GB"
    $1 = $ramObject | ConvertTo-Html -Fragment

    $2 = $($ramSticks | Select Manufacturer,Configuredclockspeed,Devicelocator,Capacity,Serialnumber,PartNumber | ConvertTo-Html -Fragment)
    Write-Host 'Got hardware information' -ForegroundColor Green
    Return $1,$2
}
function getLicensing {
    Write-Host 'Getting license information...'
    $1 = "<h2 id='Lics'>Licensing</h2>"
    $2 = $cimLics | ConvertTo-Html -Fragment
    Write-Host 'Got license information' -ForegroundColor Green
    Return $1,$2
}
function getSecureInfo {
    Write-Host 'Getting security information...'
    $1 = "<h2 id='SecInfo'>Security Information</h2>"

    $secObject = New-Object PSObject
    # Add AVs
    $i = 0
    ForEach ($a in $av) {
        Add-Member -InputObject $secObject -MemberType NoteProperty -Name "Antivirus$i" -Value $a.DisplayName
        $i++
    }
    # Add FW and assume its defender if there is no entry (default)
    If ($fw.DisplayName -eq $NULL) {
        Add-Member -InputObject $secObject -MemberType NoteProperty -Name 'Firewall' -Value "Assume Defender"
        } Else {
        Add-Member -InputObject $secObject -MemberType NoteProperty -Name 'Firewall' -Value $fw.DisplayName
    }
    # Add UAC
    Add-Member -InputObject $secObject -MemberType NoteProperty -Name 'UAC' -Value $(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System).EnableLUA
    # Add Secureboot
    Add-Member -InputObject $secObject -MemberType NoteProperty -Name 'SecureBoot' -Value $secureBoot
    $2 = $secObject | ConvertTo-Html -Fragment -As List

    $6 = "<h2 id='TPM'>TPM</h2>"
    If ($tpm -eq $NULL) {
        $7 = 'TPM not detected'
        } Else {
        $7 = $tpm | Select IsActivated_InitialValue,IsEnabled_InitialValue,IsOwned_InitialValue,PhysicalPresenceVersionInfo,SpecVersion | ConvertTo-Html -Fragment
    }
    Write-Host 'Got security information' -ForegroundColor Green
    Return $1,$2,$6,$7
}
function getBIOS {
    $1 = "<h2 id='bios'>BIOS</h2>"
    $2 = $cimBios | Select Manufacturer,SMBIOSBIOSVersion,Name,Version | ConvertTo-Html -Fragment
    Return $1,$2
}
function getVars {
    Write-Host 'Getting variables...'
    $1 = "<h2 id='SysVar'>System Variables</h2>"
    $varsMachine = [Environment]::GetEnvironmentVariables("Machine")
    $varObjMachine = New-Object PSObject
    ForEach ($var in $varsMachine.Keys) {
        Add-Member -InputObject $varObjMachine -MemberType NoteProperty -Name $var -Value $varsMachine.$var
    }
    $2 = $varObjMachine | ConvertTo-Html -Fragment -As List
    $3 = "<h2 id='UserVar'>User Variables</h2>"
    $varsUser = [Environment]::GetEnvironmentVariables("User")
    $varObjUser = New-Object PSObject
    ForEach ($var in $varsUser.Keys) {
        Add-Member -InputObject $varObjUser -MemberType NoteProperty -Name $var -Value $varsUser.$var
    }
    $4 = $varObjUser | ConvertTo-Html -Fragment -As List
    Write-Host 'Got variables' -ForegroundColor Green
    Return $1,$2,$3,$4
}
function getUpdates {
    Write-Host 'Getting applied hotfixes...'
    $1 = "<h2 id='hotfixes'>Installed updates</h2>"
    $2 = Get-HotFix | Sort-Object -Property InstalledOn -Descending -ErrorAction SilentlyContinue | Select Description,HotFixID,InstalledOn | ConvertTo-Html -Fragment
    Write-Host 'Got applied hotfixes' -ForegroundColor Green
    Return $1,$2
}
function getStartup {
    Write-Host 'Getting startup tasks...'
    $1 = "<h2 id='StartupTasks'>Startup Tasks for user</h2>"
    $2 = $startUps
    Write-Host 'Got startup tasks' -ForegroundColor Green
    Return $1,$2
}
function getPower {
    Write-Host 'Getting power profiles...'
    $1 = "<h2 id='Power'>Powerprofiles</h2>"
    $2 = $powerProfiles | select @{Name="Profile";Expression={$_.ElementName}},IsActive | ConvertTo-Html -Fragment
    Write-Host 'Got power profiles' -ForegroundColor Green
    Return $1,$2
}
function getRamUsage {
    Write-Host 'Getting running processes...'
    $1 = "<h2 id='RunningProcs'>Running Processes</h2>"
    $mem =  Get-WmiObject -Class WIN32_OperatingSystem
    $memUsed = [Math]::Round($($mem.TotalVisibleMemorySize - $mem.FreePhysicalMemory)/1048576,2)
    $memTotal = Get-WMIObject -class Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)}
    $2 = "Total RAM usage: " + $memUsed + "/" + $memTotal + " GB" 
    Return $1,$2
}
function getProcesses {
    $properties=@(
        @{Name="Name"; 
            Expression = {$_.name}},
        @{Name="Count"; 
            Expression = {(Get-Process -Name $_.Name -ErrorAction SilentlyContinue | Group-Object -Property ProcessName).Count}
        },
        @{Name="NPM (M)"; 
            Expression = {[Math]::Round(($_.NPM / 1MB), 3)}
        },
        @{Name="Mem (M)"; 
            Expression = {
                $totar = 0
                ForEach ($proc in $(Get-Process -Name $_.Name -ErrorAction SilentlyContinue)) {
                    $total = $total + [Math]::Round(($proc.WS / 1MB), 2)
                }
                $total
            }
        },
        @{Name = "CPU"; 
            Expression = {
                $total = 0
                ForEach ($proc in $(Get-Process -Name $_.Name -ErrorAction SilentlyContinue)) {
                    $TotalSec = (New-TimeSpan -Start $proc.StartTime).TotalSeconds
                    $total = $total + [Math]::Round( ($proc.CPU * 100 / $TotalSec), 2)
                }
                $total
            }
        },
        @{Name="ProductVersion ";
            Expression = {$_.ProductVersion}
        },
        @{Name="Path";
            Expression = {$_.Path}
        }
    )
    $1 = "`n"
    $2 = $runningProcesses | Select -Unique | Select $properties | Sort-Object -Property name | ConvertTo-Html -Fragment
    Write-Host 'Got processes' -ForegroundColor Green
    return $1,$2
}
function getServices {
    Write-Host 'Getting services...'
    $1 = "<h2 id='Services'>Services</h2>"
    $servicesTable = $services | Select Status,DisplayName | Sort -Property DisplayName | ConvertTo-Html -Fragment
    $2 = $servicesTable -replace "<td>Stopped</td>", "<td style='color:#ab6387'>Stopped</td>" -replace "<td>Running</td>", "<td style='color:#87ab63'>Running</td>"
    Write-Host 'Got services' -ForegroundColor Green
    Return $1,$2
}
function getInstalledApps {
    Write-Host 'Getting installed apps...'
    $apps = $installedBase | Select InstallDate,DisplayName | Sort-Object DisplayName
    $1 = "<h2 id='InstalledApps'>Installed Apps</h2>"
    $2 = $apps | ConvertTo-Html -Fragment
    Write-Host 'Got installed apps' -ForegroundColor Green
    Return $1,$2
}
function getNets {
    Write-Host 'Getting network configurations...'
    $1 = "<h2 id='NetConfig'>Network Configuration</h2>"
    # make an object for each adatper and then add them to an array
    $netArray = @()
    ForEach ($int in $netAdapters) {
        $intObject = New-Object PSObject 
        Add-Member -InputObject $intObject -MemberType NoteProperty -Name "Name" -Value $int.Name -ErrorAction SilentlyContinue
        Add-Member -InputObject $intObject -MemberType NoteProperty -Name "State" -Value $int.MediaConnectionState -ErrorAction SilentlyContinue
        Add-Member -InputObject $intObject -MemberType NoteProperty -Name "Mac" -Value $int.MacAddress -ErrorAction SilentlyContinue
        Add-Member -InputObject $intObject -MemberType NoteProperty -Name "LinkSpeed" -Value $int.LinkSpeed -ErrorAction SilentlyContinue
        Add-Member -InputObject $intObject -MemberType NoteProperty -Name "Description" -Value $int.ifDesc -ErrorAction SilentlyContinue

        $i = 0
        $ips = $(Get-NetIPAddress -InterfaceIndex $int.IfIndex)
        ForEach ($ip in $ips) {
            Add-Member -InputObject $intObject -MemberType NoteProperty -Name $ip.AddressFamily -Value $ip.IPAddress -ErrorAction SilentlyContinue
            Add-Member -InputObject $intObject -MemberType NoteProperty -Name "Lease$i" -Value $ip.PrefixOrigin -ErrorAction SilentlyContinue
            $i++
        }

        $dnsS = $(Get-DnsClientServerAddress -InterfaceIndex $int.IfIndex)
        ForEach ($dns in $dnsS) {
            $i = 0
            ForEach ($d in $dns.ServerAddresses) {
                Add-Member -InputObject $intObject -MemberType NoteProperty -Name "DNS$i" -Value $d -ErrorAction SilentlyContinue
                $i++
            }
        }
        $netArray += $intObject
    }
    $2 = $netArray | ConvertTo-Html -Fragment -As List
    Write-Host 'Got network configurations' -ForegroundColor Green
    Return $1,$2
}
function getTcpSettings {
    $1 = Get-NetOffloadGlobalSetting | ConvertTo-Html ReceiveSideScaling -Fragment
    $2 = Get-NetTCPSetting | Select SettingName,Autotuninglevellocal | ConvertTo-Html -Fragment
    Return $1,$2
}
function getNetConnections {
    Write-Host 'Getting network connections...'
    $1 = "<h2 id='NetConnections'>Network Connections</h2>"
    $2 = "<h3>TCP</h3>"
    $3 = Get-NetTCPConnection | Select local*,remote*,state,@{Name="Process";Expression={(Get-Process -Id $_.OwningProcess).ProcessName}},@{Name="Path";Expression={(Get-Process -Id $_.OwningProcess).Path}} | Sort-Object -Property Process,RemotePort | ConvertTo-Html -Fragment
    $4 = "<h3>UDP</h3>"
    $5 = Get-NetUDPEndpoint | Select LocalAddress,LocalPort,@{Name="Process";Expression={(Get-Process -Id $_.OwningProcess).ProcessName}},@{Name="Path";Expression={(Get-Process -Id $_.OwningProcess).Path}} | Sort-Object -Property Process,LocalPort | ConvertTo-Html -Fragment

    Write-Host 'Got network connections' -ForegroundColor Green
    Return $1,$2,$3,$4,$5
}
function getDrivers {
    Write-Host 'Getting driver information...'
    $1 = "<h2 id='Drivers'>Drivers and device versions</h2>"
    $2 = $(gwmi Win32_PnPSignedDriver | Select devicename,driverversion | ConvertTo-Html -Fragment)
    $3 = "<h2 id='issueDevices'>Devices with issues</h2>"
    $4 = $issueDevices | Select Status,Name,InstanceID | ConvertTo-HTML -Fragment
    Write-Host 'Got driver information' -ForegroundColor Green
    Return $1,$2,$3,$4
}
function getAudio {
    Write-Host 'Getting audio devices...'
    $1 = "<h2 id='Audio'>Audio devices</h2>"
    $2 = $cimAudio | ConvertTo-Html -Fragment
    Write-Host 'Got audio devices' -ForegroundColor Green
    Return $1,$2
}
function getDisks {
    Write-Host 'Getting disk layouts...'
    $1 = "<h2 id='Disks'>Disk layouts</h2>"
    $disks = Get-Partition | Select DiskNumber -Unique
    $2 = $($disks | % { Get-Partition -DiskNumber $_.DiskNumber | Select OperationalStatus,DiskNumber,PartitionNumber,@{Name="Size (GB)";Expression={[math]::round($_.Size/1GB,4)}},IsActive,IsBoot,IsReadOnly | ConvertTo-Html -Fragment })
    $3 = $volumes | Select HealthStatus,DriveType,FileSystem,FileSystemLabel,DedupMode,AllocationUnitSize,DriveLetter,@{Name="Size Remaining (GB)";Expression={[math]::round($_.SizeRemaining/1GB,4)}}, @{Name="Size Total (GB)";Expression={[math]::round($_.Size/1GB,4)}} |ConvertTo-Html -Fragment
    Write-Host 'Got disk layouts' -ForegroundColor Green
    Return $1,$2,$3
}
function getSmart {
    $1 = "<h2 id='SMART'>SMART</h2>"
    $2 = @()
    ForEach ($i in $smart) {
        $2 += $i | ConvertTo-Html -Fragment -As List
    }
    Return $1,$2
}
function getHosts {
    If ($hostsSum -ne $hostsHash) {
        $hostsSum
        $hostsContent
        Write-Output ""
    }
}
function getTimer {
    $timer.Stop()
    $1 = 'Runtime'
    $2 = $timer.Elapsed | Select Minutes,Seconds | ConvertTo-Html -Fragment
    Return $1,$2
}

