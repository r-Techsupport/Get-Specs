<#
.SYNOPSIS
  Gather and upload specifications of a Windows host to rTechsupport
.DESCRIPTION
  Use various native powershell and wmic functions to gather verbose information on a system to assist troubleshooting
.OUTPUTS Specs
  '.\TechSupport_Specs.txt'
#>

# Real script starts at line 498

Function New-WPFMessageBox {

    # CHANGES
    # 2017-09-11 - Added some required assemblies in the dynamic parameters to avoid errors when run from the PS console host.
    
    # Define Parameters
    [CmdletBinding()]
    Param
    (
        # The popup Content
        [Parameter(Mandatory=$True,Position=0)]
        [Object]$Content,

        # The window title
        [Parameter(Mandatory=$false,Position=1)]
        [string]$Title,

        # The buttons to add
        [Parameter(Mandatory=$false,Position=2)]
        [ValidateSet('OK','OK-Cancel','Abort-Retry-Ignore','Yes-No-Cancel','Yes-No','Retry-Cancel','Cancel-TryAgain-Continue','None')]
        [array]$ButtonType = 'OK',

        # The buttons to add
        [Parameter(Mandatory=$false,Position=3)]
        [array]$CustomButtons,

        # Content font size
        [Parameter(Mandatory=$false,Position=4)]
        [int]$ContentFontSize = 14,

        # Title font size
        [Parameter(Mandatory=$false,Position=5)]
        [int]$TitleFontSize = 14,

        # BorderThickness
        [Parameter(Mandatory=$false,Position=6)]
        [int]$BorderThickness = 0,

        # CornerRadius
        [Parameter(Mandatory=$false,Position=7)]
        [int]$CornerRadius = 8,

        # ShadowDepth
        [Parameter(Mandatory=$false,Position=8)]
        [int]$ShadowDepth = 3,

        # BlurRadius
        [Parameter(Mandatory=$false,Position=9)]
        [int]$BlurRadius = 20,

        # WindowHost
        [Parameter(Mandatory=$false,Position=10)]
        [object]$WindowHost,

        # Timeout in seconds,
        [Parameter(Mandatory=$false,Position=11)]
        [int]$Timeout,

        # Code for Window Loaded event,
        [Parameter(Mandatory=$false,Position=12)]
        [scriptblock]$OnLoaded,

        # Code for Window Closed event,
        [Parameter(Mandatory=$false,Position=13)]
        [scriptblock]$OnClosed

    )

    # Dynamically Populated parameters
    DynamicParam {
        
        # Add assemblies for use in PS Console 
        Add-Type -AssemblyName System.Drawing, PresentationCore
        
        # ContentBackground
        $ContentBackground = 'ContentBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentBackground, $RuntimeParameter)
        

        # FontFamily
        $FontFamily = 'FontFamily'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute)  
        $arrSet = [System.Drawing.FontFamily]::Families.Name | Select -Skip 1 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FontFamily, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($FontFamily, $RuntimeParameter)
        $PSBoundParameters.FontFamily = "Segoe UI"

        # TitleFontWeight
        $TitleFontWeight = 'TitleFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleFontWeight, $RuntimeParameter)

        # ContentFontWeight
        $ContentFontWeight = 'ContentFontWeight'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Windows.FontWeights] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentFontWeight = "Normal"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentFontWeight, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentFontWeight, $RuntimeParameter)
        

        # ContentTextForeground
        $ContentTextForeground = 'ContentTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ContentTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ContentTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ContentTextForeground, $RuntimeParameter)

        # TitleTextForeground
        $TitleTextForeground = 'TitleTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleTextForeground, $RuntimeParameter)

        # BorderBrush
        $BorderBrush = 'BorderBrush'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.BorderBrush = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($BorderBrush, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($BorderBrush, $RuntimeParameter)


        # TitleBackground
        $TitleBackground = 'TitleBackground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.TitleBackground = "White"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($TitleBackground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($TitleBackground, $RuntimeParameter)

        # ButtonTextForeground
        $ButtonTextForeground = 'ButtonTextForeground'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = [System.Drawing.Brushes] | Get-Member -Static -MemberType Property | Select -ExpandProperty Name 
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $PSBoundParameters.ButtonTextForeground = "Black"
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ButtonTextForeground, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ButtonTextForeground, $RuntimeParameter)

        # Sound
        $Sound = 'Sound'
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttribute.Mandatory = $False
        #$ParameterAttribute.Position = 14
        $AttributeCollection.Add($ParameterAttribute) 
        $arrSet = (Get-ChildItem "$env:SystemDrive\Windows\Media" -Filter Windows* | Select -ExpandProperty Name).Replace('.wav','')
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)    
        $AttributeCollection.Add($ValidateSetAttribute)
        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($Sound, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($Sound, $RuntimeParameter)

        return $RuntimeParameterDictionary
    }

    Begin {
        Add-Type -AssemblyName PresentationFramework
    }
    
    Process {

# Define the XAML markup
[XML]$Xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="" SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen" WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True" Background="Transparent" Opacity="1">
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border>
                            <Grid Background="{TemplateBinding Background}">
                                <ContentPresenter />
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border x:Name="MainBorder" Margin="10" CornerRadius="$CornerRadius" BorderThickness="$BorderThickness" BorderBrush="$($PSBoundParameters.BorderBrush)" Padding="0" >
        <Border.Effect>
            <DropShadowEffect x:Name="DSE" Color="Black" Direction="270" BlurRadius="$BlurRadius" ShadowDepth="$ShadowDepth" Opacity="0.6" />
        </Border.Effect>
        <Border.Triggers>
            <EventTrigger RoutedEvent="Window.Loaded">
                <BeginStoryboard>
                    <Storyboard>
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="ShadowDepth" From="0" To="$ShadowDepth" Duration="0:0:1" AutoReverse="False" />
                        <DoubleAnimation Storyboard.TargetName="DSE" Storyboard.TargetProperty="BlurRadius" From="0" To="$BlurRadius" Duration="0:0:1" AutoReverse="False" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </Border.Triggers>
        <Grid >
            <Border Name="Mask" CornerRadius="$CornerRadius" Background="$($PSBoundParameters.ContentBackground)" />
            <Grid x:Name="Grid" Background="$($PSBoundParameters.ContentBackground)">
                <Grid.OpacityMask>
                    <VisualBrush Visual="{Binding ElementName=Mask}"/>
                </Grid.OpacityMask>
                <StackPanel Name="StackPanel" >                   
                    <TextBox Name="TitleBar" IsReadOnly="True" IsHitTestVisible="False" Text="$Title" Padding="10" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$TitleFontSize" Foreground="$($PSBoundParameters.TitleTextForeground)" FontWeight="$($PSBoundParameters.TitleFontWeight)" Background="$($PSBoundParameters.TitleBackground)" HorizontalAlignment="Stretch" VerticalAlignment="Center" Width="Auto" HorizontalContentAlignment="Center" BorderThickness="0"/>
                    <DockPanel Name="ContentHost" Margin="0,10,0,10"  >
                    </DockPanel>
                    <DockPanel Name="ButtonHost" LastChildFill="False" HorizontalAlignment="Center" >
                    </DockPanel>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

[XML]$ButtonXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="Auto" Height="30" FontFamily="Segui" FontSize="16" Background="Transparent" Foreground="White" BorderThickness="1" Margin="10" Padding="20,0,20,0" HorizontalAlignment="Right" Cursor="Hand"/>
"@

[XML]$ButtonTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="16" Background="Transparent" Foreground="$($PSBoundParameters.ButtonTextForeground)" Padding="20,5,20,5" HorizontalAlignment="Center" VerticalAlignment="Center"/>
"@

[XML]$ContentTextXaml = @"
<TextBlock xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Text="$Content" Foreground="$($PSBoundParameters.ContentTextForeground)" DockPanel.Dock="Right" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="$($PSBoundParameters.FontFamily)" FontSize="$ContentFontSize" FontWeight="$($PSBoundParameters.ContentFontWeight)" TextWrapping="Wrap" Height="Auto" MaxWidth="500" MinWidth="50" Padding="10"/>
"@

    # Load the window from XAML
    $Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml))

    # Custom function to add a button
    Function Add-Button {
        Param($Content)
        $Button = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonXaml))
        $ButtonText = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ButtonTextXaml))
        $ButtonText.Text = "$Content"
        $Button.Content = $ButtonText
        $Button.Add_MouseEnter({
            $This.Content.FontSize = "17"
        })
        $Button.Add_MouseLeave({
            $This.Content.FontSize = "16"
        })
        $Button.Add_Click({
            New-Variable -Name WPFMessageBoxOutput -Value $($This.Content.Text) -Option ReadOnly -Scope Script -Force
            $Window.Close()
        })
        $Window.FindName('ButtonHost').AddChild($Button)
    }

    # Add buttons
    If ($ButtonType -eq "OK")
    {
        Add-Button -Content "OK"
    }

    If ($ButtonType -eq "OK-Cancel")
    {
        Add-Button -Content "OK"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Abort-Retry-Ignore")
    {
        Add-Button -Content "Abort"
        Add-Button -Content "Retry"
        Add-Button -Content "Ignore"
    }

    If ($ButtonType -eq "Yes-No-Cancel")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Yes-No")
    {
        Add-Button -Content "Yes"
        Add-Button -Content "No"
    }

    If ($ButtonType -eq "Retry-Cancel")
    {
        Add-Button -Content "Retry"
        Add-Button -Content "Cancel"
    }

    If ($ButtonType -eq "Cancel-TryAgain-Continue")
    {
        Add-Button -Content "Cancel"
        Add-Button -Content "TryAgain"
        Add-Button -Content "Continue"
    }

    If ($ButtonType -eq "None" -and $CustomButtons)
    {
        Foreach ($CustomButton in $CustomButtons)
        {
            Add-Button -Content "$CustomButton"
        }
    }

    # Remove the title bar if no title is provided
    If ($Title -eq "")
    {
        $TitleBar = $Window.FindName('TitleBar')
        $Window.FindName('StackPanel').Children.Remove($TitleBar)
    }

    # Add the Content
    If ($Content -is [String])
    {
        # Replace double quotes with single to avoid quote issues in strings
        If ($Content -match '"')
        {
            $Content = $Content.Replace('"',"'")
        }
        
        # Use a text box for a string value...
        $ContentTextBox = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $ContentTextXaml))
        $Window.FindName('ContentHost').AddChild($ContentTextBox)
    }
    Else
    {
        # ...or add a WPF element as a child
        Try
        {
            $Window.FindName('ContentHost').AddChild($Content) 
        }
        Catch
        {
            $_
        }        
    }

    # Enable window to move when dragged
    $Window.FindName('Grid').Add_MouseLeftButtonDown({
        $Window.DragMove()
    })

    # Activate the window on loading
    If ($OnLoaded)
    {
        $Window.Add_Loaded({
            $This.Activate()
            Invoke-Command $OnLoaded
        })
    }
    Else
    {
        $Window.Add_Loaded({
            $This.Activate()
        })
    }
    

    # Stop the dispatcher timer if exists
    If ($OnClosed)
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
            Invoke-Command $OnClosed
        })
    }
    Else
    {
        $Window.Add_Closed({
            If ($DispatcherTimer)
            {
                $DispatcherTimer.Stop()
            }
        })
    }
    

    # If a window host is provided assign it as the owner
    If ($WindowHost)
    {
        $Window.Owner = $WindowHost
        $Window.WindowStartupLocation = "CenterOwner"
    }

    # If a timeout value is provided, use a dispatcher timer to close the window when timeout is reached
    If ($Timeout)
    {
        $Stopwatch = New-object System.Diagnostics.Stopwatch
        $TimerCode = {
            If ($Stopwatch.Elapsed.TotalSeconds -ge $Timeout)
            {
                $Stopwatch.Stop()
                $Window.Close()
            }
        }
        $DispatcherTimer = New-Object -TypeName System.Windows.Threading.DispatcherTimer
        $DispatcherTimer.Interval = [TimeSpan]::FromSeconds(1)
        $DispatcherTimer.Add_Tick($TimerCode)
        $Stopwatch.Start()
        $DispatcherTimer.Start()
    }

    # Play a sound
    If ($($PSBoundParameters.Sound))
    {
        $SoundFile = "$env:SystemDrive\Windows\Media\$($PSBoundParameters.Sound).wav"
        $SoundPlayer = New-Object System.Media.SoundPlayer -ArgumentList $SoundFile
        $SoundPlayer.Add_LoadCompleted({
            $This.Play()
            $This.Dispose()
        })
        $SoundPlayer.LoadAsync()
    }

    # Display the window
    $null = $window.Dispatcher.InvokeAsync{$window.ShowDialog()}.Wait()

    # SOURCE OF WPF BELOW
    # https://smsagent.wordpress.com/2017/08/24/a-customisable-wpf-messagebox-for-powershell/
    }
}

# Run the basic checks before doing anything
function promptFileIssue {
    If (Test-Path ./files) {
        } Else {
        $Params = @{
            Content = "&#10;
The files directory is missing. This most likely means you did not extract the zip file properly.
            &#10; 
            &#10;
Right click the zip file and choose 'Extract All' then run the new exe.
            &#10;"
            Title = "rTechsupport WindowsSpecifications"
            TitleBackground = "Crimson"
            TitleFontSize = 16
            TitleFontWeight = "Bold"
            TitleTextForeground = "White"
            ContentFontSize = 12
            ContentFontWeight = "Medium"
            ContentTextForeground = "Black"
            ContentBackground = "White"
            ButtonType = "none"
            CustomButtons = "Exit"
        }
        New-WPFMessageBox @Params
        If ($WPFMessageBoxOutput -eq "Exit") {
            Exit
        }
    }
}
promptFileIssue

# Declarations
$file = 'TechSupport_Specs.html'
$badSoftware = @(
    'Driver Booster'
)
$badStartup = @(
    'AutoKMS',
    'kmspico',
    'CCleanerSkipUAC',
    'McAfee Remediation',
    'IObit Uninstaller Service',
    'Driver Booster Scheduler',
    'Driver Easy Scheduled Scan',
    'KMS_VL_ALL'
)
$badProcesses = @(
    'Malwarebytes',
    'McAfee WebAdvisor',
    'Norton Security',
    'Wallpaper Engine Service',
    'Service_KMS.exe'
)
# YOU MUST MATCH THE KEY AND VALUE BELOW TO THE SAME ARRAY VALUE
$badKeys = @(
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PreviewBuilds\',
    'HKLM:\SYSTEM\Setup\LabConfig\',
    'HKLM:\SYSTEM\Setup\LabConfig\',
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\'
)
$badValues = @(
    'AllowBuildPreview',
    'BypassTPMCheck',
    'BypassSecureBootCheck',
    'NoAutoUpdate'
)
$badRegExp = @(
    'Insider builds are set to: ',
    'Windows 11 TPM bypass set to: ',
    'Windows 11 SecureBoot bypass set to: ',
    'Windows auto update is set to: '
)
$badHostnames = @(
    'ATLASOS-DESKTOP',
    'Revision-PC'
)
### Functions

function header {
$1 = "<!DOCTYPE html>
<html>
<head>
<style>
body {
    background-color: #383c4a;
    color: White;
    margin-left: 30px;
}
h2 {
    color: #87ab63;
}
a:link {
    color: White;
}
a:visited {
    color: White;
}
a:hover{
    color: #87ab63;
}
table {
  font-family: arial, sans-serif;
  font-size: 12px;
  border-collapse: collapse;
}

td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}

tr:nth-child(even) {
  background-color: #2A2E3A;
}
</style>
<body>
<pre>
"
Return $1
}
function table {
$1 = '<h2>Sections</h2>
<div style="line-height:0">
<p><a href="#Lics">Licensing</a></p>
<p><a href="#SecInfo">Security Information</a></p>
<p><a href="#hw">Hardware Basics</a></p>
<p><a href="#SysVar">System Variables</a></p>
<p><a href="#UserVar">User Variables</a></p>
<p><a href="#hotfixes">Installed updates</a></p>
<p><a href="#StartupTasks">Startup Tasks</a></p>
<p><a href="#Power">Powerprofiles</a></p>
<p><a href="#RunningProcs">Running Processes</a></p>
<p><a href="#Services">Services</a></p>
<p><a href="#InstalledApps">Installed Applications</a></p>
<p><a href="#NetConfig">Network Configuration</a></p>
<p><a href="#Drivers">Drivers and device versions</a></p>
<p><a href="#Audio">Audio Devices</a></p>
<p><a href="#Disks">Disk Layouts</a></p>
<p><a href="#SMART">SMART</a></p>
</div>'
Return $1
}

function getDate {
    Get-Date
}
function getbasicInfo {
    Write-Host 'Getting basic information...'
    $1 = '<h2>System Information</h2>'
    $bootuptime = $cimOs.LastBootUpTime
    $uptime = $(Get-Date) - $bootuptime
    $2 = 'Edition: ' + $cimOs.Caption
    $3 = 'Build: ' + $cimOs.BuildNumber
    $4 = 'Install date: ' + $cimOs.InstallDate
    $5 = 'Uptime: ' + $uptime.Days + " Days " + $uptime.Hours + " Hours " +  $uptime.Minutes + " Minutes"
    $6 = 'Hostname: ' + $cimOs.CSName
    $7 = 'Domain: ' + $env:USERDOMAIN
    $8 = 'Boot mode: ' + $env:firmware_type
    Write-Host 'Got basic information'
    Return $1,$2,$3,$4,$5,$6,$7,$8
}
function getBadThings {
    Write-Host 'Checking for issues...'
    $1 = '<h2>Visible issues</h2>'
    $2 = @()
    $3 = @()
    $4 = @()
    $5 = @()
    foreach ($soft in $badSoftware) { 
        if ($installedBase.DisplayName -contains $soft) { 
            $2 += $soft 
        } 
    }
    foreach ($start in $badStartup) { 
        if ($cimStart.Caption -contains $start) { 
            $3 += $start 
        } 
    }
    foreach ($running in $badProcesses) {
        if ($runningProcesses.Name -contains $running) {
            $4 += $running 
        } 
    }
    $i = 0
    foreach ($reg in $badKeys) {
        If (Test-Path -Path $badKeys[$i]) {
            $value = Get-ItemProperty -Path $badKeys[$i] -ErrorAction SilentlyContinue | Select-Object -ExpandProperty $badValues[$i] -ErrorAction SilentlyContinue
            $5 += $badRegExp[$i] + $value
        }
        # } Else {
            # Write-Host $badKeys[$i] $badValues[$i] "does not exist"
        # }[
        $i = $i + 1
    }
    foreach ($name in $badHostnames) {
        if ($cimOs.CSName -contains $name) {
            $6 += "Modified OS: " + $name
        } 
    }
    $c = $volumes | ? { $_.DriveLetter -eq 'C' }
    $cAllowable = $c.Size - $c.Size * .20
    $cConsumed = $c.Size - $c.SizeRemaining
    If ($cConsumed -gt $cAllowable) {
        $7 = "Less than 20% left on C"
    }
    Write-Host 'Checked for issues'
    Return $1,$2,$3,$4,$5,$6,$7
}
function getLicensing {
    Write-Host 'Getting license information...'
    $1 = "<h2 id='Lics'> Licensing</h2>"
    $2 = $cimLics | ConvertTo-Html -Fragment
    Write-Host 'Got license information'
    Return $1,$2
}
function getSecureInfo {
    Write-Host 'Getting security information...'
    $1 = "<h2 id='SecInfo'>Security Information</h2>"
    $2 = 'AV: ' + $av.DisplayName 
    $3 = 'Firewall: ' + $fw.DisplayName 
    $4 = "UAC: " + $(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System).EnableLUA 
    $5 = "Secureboot: " + $(Confirm-SecureBootUEFI) 
    $6 = $tpm | Select IsActivated_InitialValue,IsEnabled_InitialValue,IsOwned_InitialValue,PhysicalPresenceVersionInfo,SpecVersion | ConvertTo-Html -Fragment
    Write-Host 'Got security information'
    Return $1,$2,$3,$4,$5,$6
}
function getTemps {
    Add-Type -Path .\files\OpenHardwareMonitorLib.dll
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
function getCPU{
    Write-Host 'Getting hardware information...'
    $1 = "<h2 id='hw'>Hardware Basics</h2>"
    $cpuInfo = Get-WmiObject Win32_Processor
    $cpu = $cpuInfo.Name
    $2 = 'CPU: ' + $cpu + $temps[0] + 'C' 
    Return $1,$2
}
function getMobo{
    $moboBase = Get-WmiObject Win32_BaseBoard
    $moboMan = $moboBase.manufacturer
    $moboMod = $moboBase.product
    $mobo = $moboMan + " | " + $moboMod
    $1 = "Motherboard: " + $mobo 
    Return $1
}
function getGPU {
    $GPUbase = Get-WmiObject Win32_VideoController
    $1 = "Graphics Card: " + $GPUbase.Name + " " + $temps[1]+ 'C' 
    Return $1
}
function getRAM {
    $1 = "RAM: " + $(Get-WMIObject -class Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)}) + 'GB' 
    $2 = $(Get-WmiObject win32_physicalmemory | Select Manufacturer,Configuredclockspeed,Devicelocator,Capacity,Serialnumber | ConvertTo-Html -Fragment)
    Write-Host 'Got hardware information'
    Return $1,$2
}
function getVars {
    Write-Host 'Getting variables...'
    $1 = "<h2 id='SysVar'>System Variables</h2>"
    $2 = [Environment]::GetEnvironmentVariables("Machine")
    $3 = "<h2 id='UserVar'>User Variables</h2>"
    $4 = [Environment]::GetEnvironmentVariables("User")
    Write-Host 'Got variables'
    Return $1,$2,$3,$4
}
function getUpdates {
    Write-Host 'Getting applied hotfixes...'
    $1 = "<h2 id='hotfixes'>Installed updates</h2>"
    $2 = Get-HotFix | Sort-Object -Property InstalledOn -Descending | Select Description,HotFixID,InstalledOn | ConvertTo-Html -Fragment
    Write-Host 'Got applied hotfixes'
    Return $1,$2
}
function getStartup {
    Write-Host 'Getting startup tasks...'
    $1 = "<h2 id='StartupTasks'>Startup Tasks for user</h2>"
    $2 = $cimStart.Caption
    Write-Host 'Got startup tasks'
    Return $1,$2
}
function getPower {
    Write-Host 'Getting power profiles...'
    $1 = "<h2 id='Power'>Powerprofiles</h2>"
    $2 = powercfg /l
    Write-Host 'Got power profiles'
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
            Expression = {(Get-Process -Name $_.Name | Group-Object -Property ProcessName).Count}
        },
        @{Name="NPM (M)"; 
            Expression = {[Math]::Round(($_.NPM / 1MB), 3)}
        },
        @{Name="Mem (M)"; 
            Expression = {
                $total = 0
                ForEach ($proc in $(Get-Process -Name $_.Name)) {
                    $total = $total + [Math]::Round(($proc.WS / 1MB), 2)
                }
                $total
            }
        },
        @{Name = "CPU"; 
            Expression = {
                $total = 0
                ForEach ($proc in $(Get-Process -Name $_.Name)) {
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
    $2 = $(Get-Process | Select -Unique | Select $properties | Sort-Object "Mem (M)" -desc | ConvertTo-Html -Fragment)
    Write-Host 'Got processes'
    return $1,$2
}
function getServices {
    Write-Host 'Getting services...'
    $1 = "<h2 id='Services'>Services</h2>"
    $2 = $services | Select Status,DisplayName | ConvertTo-Html -Fragment
    Write-Host 'Got services'
    Return $1,$2
}
function getInstalledApps {
    Write-Host 'Getting installed apps...'
    $apps = $installedBase | Select InstallDate,DisplayName | Sort-Object InstallDate -desc
    $1 = "<h2 id='InstalledApps'>Installed Apps</h2>"
    $2 = $apps | ConvertTo-Html -Fragment
    Write-Host 'Got installed apps'
    Return $1,$2
}
function getNets {
    Write-Host 'Getting network configurations...'
    $1 = "<h2 id='NetConfig'>Network Configuration</h2>"
    $2 = $(Get-NetAdapter|Select Name,InterfaceDescription,Status,LinkSpeed | ConvertTo-Html -Fragment) 
    $3 = $(Get-NetIPAddress|Select IpAddress,InterfaceAlias,PrefixOrigin | ConvertTo-Html -Fragment)
    Write-Host 'Got network configurations'
    Return $1,$2,$3
}
function getDrivers {
    Write-Host 'Getting driver information...'
    $1 = "<h2 id='Drivers'>Drivers and device versions</h2>"
    $2 = $(gwmi Win32_PnPSignedDriver | Select devicename,driverversion | ConvertTo-Html -Fragment)
    Write-Host 'Got driver information'
    Return $1,$2
}
function getAudio {
    Write-Host 'Getting audio devices...'
    $1 = "<h2 id='Audio'>Audio devices</h2>"
    $2 = $cimAudio | ConvertTo-Html -Fragment
    Write-Host 'Got audio devices'
    Return $1,$2
}
function getDisks {
    Write-Host 'Getting disk layouts...'
    $1 = "<h2 id='Disks'>Disk layouts</h2>"
    $2 = $(Get-Partition| Select OperationalStatus,DiskNumber,PartitionNumber,Size,IsActive,IsBoot,IsReadOnly | ConvertTo-Html -Fragment) 
    $3 = $volumes | Select HealthStatus,DriveType,FileSystem,FileSystemLabel,DedupMode,AllocationUnitSize,DriveLetter,SizeRemaining,Size |ConvertTo-Html -Fragment
    Write-Host 'Got disk layouts'
    Return $1,$2,$3
}
function getSmart {
    Write-Host 'Getting SMART data...'
    $(files\DiskInfo64.exe /CopyExit)
    while (!(Test-Path 'files\DiskInfo.txt')) { 
        Start-Sleep 1
    }
    $1 = "<h2 id='SMART'>SMART</h2>"
    $2 = $(Get-Content files\DiskInfo.txt)
    Remove-Item -Force -Recurse 'files\Smart','files\DiskInfo.txt','files\DiskInfo.ini'
    Write-Host 'Got SMART'
    Return $1,$2
}
function uploadFile {
    $link = Invoke-WebRequest -ContentType 'text/plain' -Method 'PUT' -InFile $file -Uri "https://paste.rtech.support/upload/$null.html" -UseBasicParsing
    $linkProper = $link.Content -Replace "support","support/selif"
    set-clipboard $linkProper
}
function promptStart {
        $Params = @{
        Content = "&#10;
This tool will gather specifications and configurations from your machine. After running this application will ask if you want to upload your results for sharing.
        &#10; 
        &#10;
Would you like to continue?
        &#10;
        &#10;
        &#10;
        &#10;
The source code for this application can be found at https://git.dev0.sh/piper/WindowsSpecifications"
        Title = "rTechsupport Specs Tool"
        TitleBackground = "DodgerBlue"
        TitleFontSize = 16
        TitleFontWeight = "Bold"
        TitleTextForeground = "White"
        ContentFontSize = 12
        ContentFontWeight = "Medium"
        ContentTextForeground = "Black"
        ContentBackground = "White"
        ButtonType = "none"
        CustomButtons = "Start","Exit"
    }
    New-WPFMessageBox @Params
    If ($WPFMessageBoxOutput -eq "Exit") {
        Exit
    }
}
function promptUpload {
    $Params = @{
        Content = "Do you want to view or upload the specs?
        &#10;
Uploaded results are deleted after 24 hours"
        Title = "rTechsupport Specs Tool"
        TitleBackground = "DodgerBlue"
        TitleFontSize = 16
        TitleFontWeight = "Bold"
        TitleTextForeground = "White"
        ContentFontSize = 12
        ContentFontWeight = "Medium"
        ContentTextForeground = "Black"
        ContentBackground = "White"
        ButtonType = "none"
        CustomButtons = "View","Upload"
    }
    New-WPFMessageBox @Params
    If ($WPFMessageBoxOutput -eq "View") {
        Invoke-Item $file
    }
    ElseIf ($WPFMessageBoxOutput -eq "Upload") {
        uploadFile
        New-WPFMessageBox -Content "Link has been copied to your clipboard. Paste into chat to share." -Title "Upload Success" -TitleBackground Coral -WindowHost $Window
    }
}

# ------------------ #
promptStart

# Bulk data gathering
Write-Host 'Gathering main data...'
## CIM sources
$cimOs = Get-CimInstance -ClassName Win32_OperatingSystem
$cimStart = Get-CimInstance Win32_StartupCommand
$cimAudio= Get-CimInstance win32_sounddevice | Select Name,ProductName
$cimLics = Get-CimInstance -ClassName SoftwareLicensingProduct | ? { $_.PartialProductKey -ne $null } | Select Name,ProductKeyChannel,LicenseFamily,LicenseStatus,PartialProductKey
$av = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct
$fw = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName FirewallProduct
$tpm = Get-CimInstance -Namespace root/cimv2/Security/MicrosoftTpm -ClassName win32_tpm

## Other
$installedBase = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | ? {$_.DisplayName -notlike $null}
$services = $(Get-Service)
$runningProcesses = Get-Process
$volumes = Get-Volume
Write-Host 'Got main data'

# Write da file
header | Out-File -Encoding ascii $file
getDate | Out-File -Append -Encoding ascii $file
getBasicInfo | Out-File -Append -Encoding ascii $file
getBadThings | Out-File -Append -Encoding ascii $file
table | Out-File -Append -Encoding ascii $file
getLicensing | Out-File -Append -Encoding ascii $file
getSecureInfo | Out-File -Append -Encoding ascii $file
getCPU | Out-File -Append -Encoding ascii $file
getMobo | Out-File -Append -Encoding ascii $file
getGPU | Out-File -Append -Encoding ascii $file
getRAM | Out-File -Append -Encoding ascii $file
getVars | Out-File -Append -Encoding ascii $file
getUpdates | Out-File -Append -Encoding ascii $file
getStartup | Out-File -Append -Encoding ascii $file
getPower | Out-File -Append -Encoding ascii $file
getRamUsage | Out-File -Append -Encoding ascii $file
getProcesses | Out-File -Append -Encoding ascii $file
getServices | Out-File -Append -Encoding ascii $file
getInstalledApps | Out-File -Append -Encoding ascii $file
getNets | Out-File -Append -Encoding ascii $file
getDrivers | Out-File -Append -Encoding ascii $file
getAudio | Out-File -Append -Encoding ascii $file
getDisks | Out-File -Append -Encoding ascii $file
getSmart | Out-File -Append -Encoding ascii $file
promptUpload
# ------------------ #
