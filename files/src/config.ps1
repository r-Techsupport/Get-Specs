# VERSION
$version = '2.3.0'

# Declarations
## from the perspective of Get-Specs.ps1, so we use ..
$file = '..\TechSupport_Specs.html'

## hosts related
$hostsFile = 'C:\Windows\System32\drivers\etc\hosts'
$hostsHash = '2D6BDFB341BE3A6234B24742377F93AA7C7CFB0D9FD64EFA9282C87852E57085'

## bad things
$badSoftware = @(
    'Driver Booster*',
    'iTop*',
    'Driver Easy*',
    'Roblox*',
    'ccleaner*',
    'Malwarebytes*',
    'Wallpaper Engine*',
    'Voxal Voice Changer*',
    'Clownfish Voice Changer*',
    'Voicemod*',
    'Microsoft Office Enterprise 2007*',
    'System Mechanic*',
    'MyCleanPC*',
    'DriverFix*',
    'Reimage Repair*',
    'cFosSpeed*',
    'Browser Assistant*',
    'KMS*',
    'Advanced SystemCare',
    'AVG*',
    'Avast*',
    'salad*',
    'McAfee*',
    'Citrix*',
    'Norton*',
    '*Cleaner*',
    'Kaspersky*'
)
$badStartup = @(
    'AutoKMS',
    'kmspico',
    'McAfee Remediation',
    'KMS_VL_ALL',
    'WallpaperEngine'
)
$badProcesses = @(
    'MBAMService',
    'McAfee WebAdvisor',
    'Norton Security',
    'Wallpaper Engine Service',
    'Service_KMS.exe',
    'iTopVPN',
    'wallpaper32',
    'TaskbarX',
    'Process Lasso Core (Process Governor)'
)
# YOU MUST MATCH THE KEY AND VALUE BELOW TO THE SAME ARRAY VALUE
$badKeys = @(
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PreviewBuilds\',
    'HKLM:\SYSTEM\Setup\LabConfig\',
    'HKLM:\SYSTEM\Setup\LabConfig\',
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\',
    'HKLM:\SYSTEM\Setup\Status\',
    'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers',
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU',
    'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power'
)
$badValues = @(
    'AllowBuildPreview',
    'BypassTPMCheck',
    'BypassSecureBootCheck',
    'NoAutoUpdate',
    'AuditBoot',
    'HwSchMode',
    'UseWUServer',
    'HiberbootEnabled'
)
$badData = @(
    '1',
    '1',
    '1',
    '1',
    '1',
    '2',
    '1',
    '1'
)
$badRegExp = @(
    'Insider builds are set to: ',
    'Windows 11 TPM bypass set to: ',
    'Windows 11 SecureBoot bypass set to: ',
    'Windows auto update is set to: ',
    'Audit Boot is set to: ',
    '(HAGS) HwSchMode is set to: ',
    'Use WSUS is: ',
    'Fastboot is set to: '
)
$badHostnames = @(
    'ATLASOS-DESKTOP',
    'Revision-PC',
    'NEXUSLITE-PC'
)
$badAdapters = @(
    '*TAP*',
    '*TUN*',
    '*VPN*',
    '*Hamachi*',
    '*Tunnel*',
    '*Nord*',
    '*SurfShark*',
    'TunnelBear Adapter V9',
    'Private Internet Access Network Adapter',
    'ZeroTier Virtual Port',
    'Kaspersky Security Data Escort Adapter'
)
$badFiles = @(
    'C:\Windows\system32\SppExtComObjHook.dll',
    'HKCU:\Software\azurite'
)
$badPower = @(
    'KernelOS*',
    'NixOS*',
    'amit*'
)
$microCode = @(
    'C:\Windows\System32\mcupdate_genuineintel.dll',
    'C:\WindowsSystem32\mcupdate_authenticamd.dll'
)
$builds = @(
    '10240',
    '10586',
    '14393',
    '15063',
    '16299',
    '17134',
    '17763',
    '18362',
    '18363',
    '19041',
    '19042',
    '19043',
    '19044',
    '22000',
    '22621'
)
$versions = @(
    '1507',
    '1511',
    '1607',
    '1703',
    '1709',
    '1803',
    '1809',
    '1903',
    '1909',
    '2004',
    '20H2',
    '21H1',
    '21H2',
    '21H2',
    '22H2'
)
$eolDates = @(
    '2017-05-17',
    '2017-10-10',
    '2018-04-10',
    '2018-10-09',
    '2019-04-19',
    '2019-11-12',
    '2020-11-10',
    '2020-10-08',
    '2021-05-11',
    '2021-12-14',
    '2022-05-10',
    '2022-12-13',
    '2023-10-10',
    '2023-10-10',
    '2024-10-14'
)
