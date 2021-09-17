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
    'Driver Booster*',
    'iTop*',
    'Driver Easy*',
    'Roblox*',
    'ccleaner',
    'Malwarebytes'
)
$badStartup = @(
    'AutoKMS',
    'kmspico',
    'McAfee Remediation',
    'KMS_VL_ALL'
)
$badProcesses = @(
    'MBAMService',
    'McAfee WebAdvisor',
    'Norton Security',
    'Wallpaper Engine Service',
    'Service_KMS.exe',
    'iTopVPN'
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
$builds = @(
    '10240',
    '10586',
    '14393',
    '15063',
    '16299',
    '17017',
    '17134',
    '17763',
    '18362',
    '18363',
    '19041',
    '19042',
    '19043'
)
$versions = @(
    '1507',
    '1511',
    '1607',
    '1703',
    '1709',
    '1803',
    '1803',
    '1809',
    '1903',
    '1909',
    '2004',
    '20H2',
    '21H1'
)
$masKeys = @(
    'YTMG3-N6DKC-DKB77-7M9GH-8HVX7',
    '4CPRK-NM3K3-X6XXQ-RXX86-WXCHW',
    'N2434-X9D7W-8PF6X-8DV9T-8TYMD',
    'BT79Q-G7N6G-PGBYW-4YWX6-6F4BT',
    'YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY',
    '84NGF-MHBT6-FXBX8-QWJK7-DRR8H',
    'XGVPP-NMH47-7TTHJ-W3FW7-8HV2C',
    '3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT',
    'FWN7H-PF93Q-4GGP8-M8RF3-MDWWW',
    '8V8WN-3GXBH-2TCMG-XHRX3-9766K',
    'NK96Y-D9CD8-W44CQ-R8YTK-DYJWX',
    '2DBW3-N2PJG-MVHW3-G7TDK-9HKR4',
    '43TBQ-NH92J-XKTM7-KT3KK-P39PB',
    'VK7JG-NPHTM-C97JM-9MPGT-3V66T',
    '2B87N-8KFHP-DKV6R-Y2C8J-PKCKT',
    '8PTT6-RNW4C-6V7J2-C2D3X-MHBPB',
    'GJTYN-HDMQY-FRR76-HVGC7-QPF8P',
    'DXG7C-N36C4-C4HTG-X4T3X-2YV77',
    'WYPNQ-8C467-V2W6J-TX4WX-WT2RQ',
    'NJCF7-PW8QT-3324D-688JX-2YV66',
    'XQQYW-NFFMW-XJPBH-K8732-CKFFD',
    'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99',
    'PVMJN-6DFY6-9CCP6-7BKTT-D3WVR',
    '3KHY7-WNT83-DGQKR-F7HPR-844BM',
    '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH',
    'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2',
    '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ',
    'NPPR9-FWDCX-D2C8J-H872K-2YT43',
    'YYVX9-NTFWV-6MDM3-9PT4T-4M68B',
    '44RPN-FTY23-9VTTB-MP9BX-T84FV',
    'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4',
    'DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ',
    'M7XTQ-FN8P6-TTKYV-9D4CC-J462D',
    'QFFDN-GRT3P-VKWWX-X7T3R-8B639',
    '92NFX-8DJQP-P6BBQ-THF9C-7CG2H',
    'W269N-WFGWX-YVC9B-4J6C9-T83GX',
    '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y',
    'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC',
    'MH37W-N47XK-V7XM9-C7227-GCQG9',
    'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J',
    '9FNHH-K3HBT-3W4TD-6383H-6XYWF',
    '7NBT4-WGBQX-MP4H7-QXFF8-YP3KX',
    'CPWHC-NT2C7-VYW78-DHDB2-PG3GK',
    'QN4C6-GBJD2-FB422-GHWJK-GJG2R',
    'CB7KF-BWN84-R7R2Y-793K2-8XDDG',
    'WMDGN-G9PQG-XVVXX-R3X43-63DFG',
    'JCKRF-N37P4-C2D82-9YXRT-4M63B',
    'WVDHN-86M7X-466P6-VHXV7-YY726',
    'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY',
    'N69G4-B89J2-4G8F4-WWYCC-J464C',
    'VP34G-4NPPG-79JTQ-864T4-R3MQX',
    'FDNH6-VW9RW-BXPJ7-4XTYG-239TB',
    '6Y6KB-N82V8-D8CQV-23MJW-BWTG6',
    'DPCNP-XQFKJ-BJF7R-FRC8D-GF6G4',
    'WNMTR-4C88C-JK8YV-HQ7T2-76DF9',
    '2F77B-TNFGY-69QQF-B8YKP-D69TJ',
    'NBTWJ-3DR69-3C4V8-C26MC-GQ9M6',
    'M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK',
    '7B9N3-D94CG-YTVHR-QBPX3-RJP64',
    'BB6NG-PQ82V-VRDPW-8XVD2-V8P66',
    'NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3',
    'XYTND-K6QKT-K2MRH-66RTM-43JKP',
    'GCRJD-8NW9H-F2CDX-CCM8D-9D6T9',
    'HMCNV-VVBFX-7HMBH-CTY9B-B4FXY',
    '789NJ-TQK6T-6XTH8-J39CJ-J8D3P',
    'MHF9N-XY6XB-WVXMC-BTDCT-MKKG7',
    'TT4HM-HN7YT-62K67-RGRQJ-JFFXW',
    'NMMPB-38DD4-R2823-62W8D-VXKJB',
    'FNFKF-PWTVT-9RC8H-32HB2-JB34X',
    'VHXM3-NR6FT-RY6RT-CK882-KW2CJ',
    '3PY8R-QHNP9-W7XQD-G6DPH-3J2C9',
    'Q6HTR-N24GM-PMJFP-69CD8-2GXKR',
    'KF37N-VDV38-GRRTV-XH8X6-6F3BB',
    'R962J-37N87-9VVK2-WJ74P-XTMHR',
    'MX3RK-9HNGX-K3QKC-6PJ3F-W8D7B',
    'TNFGH-2R6PB-8XM3K-QYHX2-J4296',
    'BN3D2-R7TKB-3YPBD-8DRP2-27GG4',
    '8N2M2-HWPGY-7PGT9-HGDD8-GVGGY',
    '2WN2H-YGCQR-KFX6K-CD6TF-84YXQ',
    '4K36P-JN4VD-GDC6V-KDT89-DYFKP',
    'DXHJF-N9KQX-MFPVR-GHGQK-Y7RKV',
    'NG4HW-VH26C-733KW-K6F98-J8CK4',
    'XCVCF-2NXM9-723PB-MHCB7-2RYQQ',
    'GNBB8-YVD74-QJHX6-27H4K-8QHDG',
    '32JNW-9KQ84-P47T8-D8GGY-CWCK7',
    'JMNMF-RHW7P-DMY6X-RF3DR-X2BQT',
    'RYXVT-BNQG7-VD29F-DBMRY-HT73M',
    'NKB3R-R2F8T-3XCDP-7Q2KW-XWYQ2',
    'FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4',
    'MRPKT-YTG23-K7D7T-X2JMM-QY7MG',
    'W82YF-2Q76Y-63HXB-FGJG9-GF7QX',
    '33PXH-7Y6KF-2VJC9-XBBR8-HVTHH',
    'YDRBP-3D83W-TY26F-D46B2-XCKRJ',
    'C29WB-22CC8-VJ326-GHFJW-H9DH4',
    'YBYF6-BHCR3-JPKRB-CDW7B-F9BK4',
    'XGY72-BRBBT-FF8MH-2GG8H-W7KCW',
    '73KQT-CD9G6-K7TQG-66MRP-CQ22C',
    'N2KJX-J94YW-TQVFB-DG9YT-724CC',
    '6NMRW-2C8FM-D24W7-TQWMY-CWH2D',
    'GRFBW-QNDC4-6QBHG-CCK3B-2PR88',
    'PTXN8-JFHJM-4WC78-MPCBR-9W4KR',
    '2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG',
    'K9FYF-G6NCK-73M32-XMVPY-F9DRR',
    'D2N9P-3P6X9-2R39C-7RTCD-MDVJX',
    'W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9',
    'KNC87-3J2TX-XB4WP-VCPJV-M4FWM',
    '3NPTF-33KPT-GGBPR-YX76B-39KDD',
    'XC9B7-NBPP2-83J2H-RHMBY-92BT4',
    '48HP8-DN98B-MYWDG-T2DCC-8W83P',
    'HM7DN-YVMH3-46JC3-XYTG7-CYQJJ',
    'XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G',
    '6TPJF-RBVHG-WBW2R-86QPH-6RTM4',
    'TT8MH-CG224-D3D7Q-498W2-9QCTX',
    'YC6KT-GKW9T-YTKYR-T4X34-R7VHC',
    '74YFP-3QFB3-KQT8W-PMXWJ-7M648',
    '489J6-VHDMP-X63PK-3K798-CPX3Y',
    'GT63C-RJFQ3-4GMB6-BRFB9-CB83V',
    '736RG-XDKJK-V34PF-BHK87-J6X3K',
    'M33WV-NHY3C-R7FPM-BQGPT-239PG',
    'DRNV7-VGMM2-B3G9T-4BF84-VMFTK',
    'NCHRJ-3VPGW-X73DM-6B36K-3RQ6B',
    '3FBRX-NFP7C-6JWVK-F2YGK-H499R',
    '9FNY8-PWWTY-8RY4F-GJMTV-KHGM9',
    '8843N-BCXXD-Q84H8-R4Q37-T3CPT',
    'MCPBN-CPY7X-3PK9R-P6GTT-H8P8Y',
    'VQ9DP-NVHPH-T9HJC-J9PDT-KTQRG',
    'XM2V9-DN9HH-QB449-XDGKC-W2RMW',
    'N2CG9-YD3YK-936X4-3WR82-Q3X4H',
    'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP',
    '6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK',
    'B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B',
    'C4F7P-NCP8C-6CQPT-MQHV9-JXD2M',
    '9BGNQ-K37YR-RQHF2-38RQ3-7VCBB',
    '7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2',
    '9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT',
    'TMJWT-YYNMB-3BKTF-644FC-RVXBD',
    '7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK',
    'RRNCX-C64HY-W2MM7-MCH9G-TJHMQ',
    'G2KWX-3NW6P-PY93R-JXK2T-C9Y9V',
    'NCJ33-JHBBY-HTK98-MYCV8-HMKHJ',
    'PBX3G-NWMT6-Q7XBW-PYJGG-WXD33',
    'WGT24-HCNMF-FQ7XH-6M8K7-DRTW9',
    'D8NRQ-JTYM3-7J2DX-646CT-6836M',
    '69WXN-MBYV6-22PQG-3WGHK-RM6XC',
    'NY48V-PPYYH-3F4PX-XJRKJ-W4423',
    'DMTCJ-KNRKX-26982-JYCKT-P7KB6',
    'HFTND-W9MK4-8B7MJ-B6C4G-XQBR2',
    'XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99',
    'JNRGM-WHDWX-FJJG3-K47QV-DRTFM',
    'YG9NW-3K39V-2T3HJ-93F3Q-G83KT',
    'GNFHQ-F6YQM-KQDGJ-327XX-KQBVC',
    'PD3PC-RHNGV-FXJ29-8JK7D-RJRJK',
    '7WHWN-4T7MP-G96JF-G33KR-W8GF4',
    'GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW',
    '9C2PK-NWTVB-JMPW8-BFT28-7FTBF',
    'DR92N-9HTF2-97XKM-XW2WJ-XW3J6',
    'R69KK-NTPKF-7M3Q4-QYBHW-6MT9B',
    'J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6',
    'F47MM-N3XJP-TQXJ9-BP99D-8K837',
    '869NQ-FJ69K-466HW-QYCP2-DDBV6',
    'WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6',
    '42QTK-RN8M7-J3C4G-BBGYM-88CYV',
    'YC7DK-G2NP3-2QQC3-J6H88-GVGXT',
    'KBKQT-2NMXY-JJWGP-M62JB-92CD4',
    'FN8TT-7WMH6-2D4X9-M337T-2342K',
    '6NTH3-CW976-3G3Y2-JK3TX-8QHTT',
    'C2FG9-N6J68-H8BTJ-BW3QX-RM3B3',
    'J484Y-4NKBF-W2HMG-DBMJC-PGWR7',
    'NG2JY-H4JBT-HQXYP-78QH9-4JM2D',
    'VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB',
    'H7R7V-WPNXQ-WCYYC-76BGV-VT7GH',
    'DKT8B-N7VXH-D963P-Q4PHY-F8894',
    '2MG3G-3BNTT-3MFW9-KDQW3-TCK7R',
    'TGN6P-8MMBC-37P2F-XHXXK-P34VW',
    'QPN8Q-BJBTJ-334K3-93TGY-2PMBT',
    '4NT99-8RJFH-Q2VDH-KYG2C-4RD4F',
    'PN2WF-29XG2-T9HJ7-JQPJR-FCXK4',
    '6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7',
    'YBJTT-JG6MD-V9Q7P-DBKXJ-38W9R',
    '7TC2V-WXF6P-TD7RT-BQRXR-B8K32',
    'VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB',
    'V7QKV-4XVVR-XYV4D-F7DFM-8R6BM',
    'YGX6F-PGV49-PGW3J-9BTGG-VHKC6',
    '4HP3K-88W3F-W2K3D-6677X-F9PGB',
    'D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ',
    '7MCW8-VRQVK-G677T-PDJCM-Q8TCP',
    '767HD-QGMWX-8QTDB-9G3R2-KHFGJ',
    'V7Y44-9T38C-R2VJK-666HK-T7DDX',
    'H62QG-HXVKF-PP4HP-66KMR-CW9BM',
    'QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4',
    'K96W8-67RPQ-62T9Y-J8FQJ-BT37T',
    'Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX',
    '7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ',
    'RC8FX-88JRY-3PF7C-X8P67-P4VTT',
    'BFK7F-9MYHM-V68C7-DRQ66-83YTP',
    'HVHB3-C6FV7-KQX9W-YQG79-CRY7T',
    'D6QFG-VBYP2-XQHM7-J97RH-VVRCK'
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
    $i = 0
    $3 = 'Version is unknown. Build: ' + $cimOS.BuildNumber
    foreach ($b in $builds) {
        if ($cimOS.BuildNumber -eq $builds[$i]) {
            $3 = 'Version: ' + $versions[$i]
        }
        $i = $i + 1
    }
    $4 = 'Install date: ' + $cimOs.InstallDate
    $5 = 'Uptime: ' + $uptime.Days + " Days " + $uptime.Hours + " Hours " +  $uptime.Minutes + " Minutes"
    $6 = 'Hostname: ' + $cimOs.CSName
    $7 = 'Domain: ' + $env:USERDOMAIN
    $8 = 'Boot mode: ' + $env:firmware_type
    Write-Host 'Got basic information' -ForegroundColor Green
    Return $1,$2,$3,$4,$5,$6,$7,$8
}
Function getFullKey {
    $KeyOutput=""
    $KeyOffset = 52
     
    $IsWin10 = ([System.Math]::Truncate($key[66] / 6)) -band 1
    $key[66] = ($Key[66] -band 0xF7) -bor (($IsWin10 -band 2) * 4)
    $i = 24 
    $maps = "BCDFGHJKMPQRTVWXY2346789"
    Do {
        $current= 0
        $j = 14
        Do {
           $current = $current* 256
           $current = $Key[$j + $KeyOffset] + $Current
           $Key[$j + $KeyOffset] = [System.Math]::Truncate($Current / 24 )
           $Current=$Current % 24
           $j--
        } 
        while ($j -ge 0)
            $i-- 
            $KeyOutput = $Maps.Substring($Current, 1) + $KeyOutput
            $last = $current
    } while ($i -ge 0)
      
    If ($IsWin10 -eq 1) {
        $keypart1 = $KeyOutput.Substring(1,$last)
        $insert = "N"
        $KeyOutput = $KeyOutput.Replace($keypart1, $keypart1 + $insert)
        if ($Last -eq 0) {  $KeyOutput = $insert + $KeyOutput }
    }
  
    if ($keyOutput.Length -eq 26) {
        $1 = [String]::Format("{0}-{1}-{2}-{3}-{4}",
            $KeyOutput.Substring(1, 5),
            $KeyOutput.Substring(6, 5),
            $KeyOutput.Substring(11,5),
            $KeyOutput.Substring(16,5),
            $KeyOutput.Substring(21,5))
    } else {
        $KeyOutput
    }
    return $1
}
function getBadThings {
    Write-Host 'Checking for issues...'
    $1 = '<h2>Visible issues</h2>'
    $2 = @()
    $3 = @()
    $4 = @()
    $5 = @()
    # bad software 
    $i = 0
    $t = $badSoftware.Count
    while ($i -lt $t) {
        $2 += $(@($installedBase.DisplayName) -like $badSoftware[$i])
        $i = $i + 1
    }
    # bad startups
    foreach ($start in $badStartup) { 
        if ($cimStart.Caption -contains $start) { 
            $3 += $start 
        }
    }
    # bad processes
    foreach ($running in $badProcesses) {
        if ($runningProcesses.Name -contains $running) {
            $4 += "Process: " + $running 
        } 
    }
    # bad registry values
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
    # bad hostnames
    if ($badHostnames -contains $cimOS.CSName) {
        $6 = "Modified OS: " + $cimOS.CSName
    }
    # bad diskspace
    $c = $volumes | ? { $_.DriveLetter -eq 'C' }
    $cAllowable = $c.Size - $c.Size * .20
    $cConsumed = $c.Size - $c.SizeRemaining
    If ($cConsumed -gt $cAllowable) {
        $7 = "Less than 20% left on C"
    }
    # bad productKeys
    If ($masKeys -contains $(getFullKey)) {
        $8 = "MAS detected"
    }
    # bad dns suffixes
    If ('utopia.net' -in $dns.SuffixSearchList) {
        $9 = "Utopia malware"
    }
    Write-Host 'Checked for issues' -ForegroundColor Green
    Return $1,$2,$3,$4,$5,$6,$7,$8,$9
}
function getLicensing {
    Write-Host 'Getting license information...'
    $1 = "<h2 id='Lics'> Licensing</h2>"
    $2 = $cimLics | ConvertTo-Html -Fragment
    Write-Host 'Got license information' -ForegroundColor Green
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
    Write-Host 'Got security information' -ForegroundColor Green
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
    Write-Host 'Got hardware information' -ForegroundColor Green
    Return $1,$2
}
function getVars {
    Write-Host 'Getting variables...'
    $1 = "<h2 id='SysVar'>System Variables</h2>"
    $2 = [Environment]::GetEnvironmentVariables("Machine")
    $3 = "<h2 id='UserVar'>User Variables</h2>"
    $4 = [Environment]::GetEnvironmentVariables("User")
    Write-Host 'Got variables' -ForegroundColor Green
    Return $1,$2,$3,$4
}
function getUpdates {
    Write-Host 'Getting applied hotfixes...'
    $1 = "<h2 id='hotfixes'>Installed updates</h2>"
    $2 = Get-HotFix | Sort-Object -Property InstalledOn -Descending | Select Description,HotFixID,InstalledOn | ConvertTo-Html -Fragment
    Write-Host 'Got applied hotfixes' -ForegroundColor Green
    Return $1,$2
}
function getStartup {
    Write-Host 'Getting startup tasks...'
    $1 = "<h2 id='StartupTasks'>Startup Tasks for user</h2>"
    $2 = $cimStart.Caption
    Write-Host 'Got startup tasks' -ForegroundColor Green
    Return $1,$2
}
function getPower {
    Write-Host 'Getting power profiles...'
    $1 = "<h2 id='Power'>Powerprofiles</h2>"
    $2 = powercfg /l
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
    Write-Host 'Got processes' -ForegroundColor Green
    return $1,$2
}
function getServices {
    Write-Host 'Getting services...'
    $1 = "<h2 id='Services'>Services</h2>"
    $2 = $services | Select Status,DisplayName | ConvertTo-Html -Fragment
    Write-Host 'Got services' -ForegroundColor Green
    Return $1,$2
}
function getInstalledApps {
    Write-Host 'Getting installed apps...'
    $apps = $installedBase | Select InstallDate,DisplayName | Sort-Object InstallDate -desc
    $1 = "<h2 id='InstalledApps'>Installed Apps</h2>"
    $2 = $apps | ConvertTo-Html -Fragment
    Write-Host 'Got installed apps' -ForegroundColor Green
    Return $1,$2
}
function getNets {
    Write-Host 'Getting network configurations...'
    $1 = "<h2 id='NetConfig'>Network Configuration</h2>"
    $2 = $(Get-NetAdapter|Select Name,InterfaceDescription,Status,LinkSpeed | ConvertTo-Html -Fragment) 
    $3 = $(Get-NetIPAddress|Select IpAddress,InterfaceAlias,PrefixOrigin | ConvertTo-Html -Fragment)
    Write-Host 'Got network configurations' -ForegroundColor Green
    Return $1,$2,$3
}
function getDrivers {
    Write-Host 'Getting driver information...'
    $1 = "<h2 id='Drivers'>Drivers and device versions</h2>"
    $2 = $(gwmi Win32_PnPSignedDriver | Select devicename,driverversion | ConvertTo-Html -Fragment)
    Write-Host 'Got driver information' -ForegroundColor Green
    Return $1,$2
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
    $2 = $(Get-Partition| Select OperationalStatus,DiskNumber,PartitionNumber,Size,IsActive,IsBoot,IsReadOnly | ConvertTo-Html -Fragment) 
    $3 = $volumes | Select HealthStatus,DriveType,FileSystem,FileSystemLabel,DedupMode,AllocationUnitSize,DriveLetter,SizeRemaining,Size |ConvertTo-Html -Fragment
    Write-Host 'Got disk layouts' -ForegroundColor Green
    Return $1,$2,$3
}
function getSmart {
    Write-Host 'Getting SMART data...'
    $1 = "<h2 id='SMART'>SMART</h2>"
    $to = new-timespan -Seconds 30
    $sw = [diagnostics.stopwatch]::StartNew()
    $(files\DiskInfo64.exe /CopyExit)
    While ($sw.elapsed -lt $to){
        if (Test-Path 'files\DiskInfo.txt'){
            $2 = $(Get-Content files\DiskInfo.txt)
        } Else {
            Start-Sleep -Seconds 1
            $2 = "Timed out after 30 seconds. Verify disk health manually with CDI."
        }
    }
    Remove-Item -Force -Recurse 'files\Smart','files\DiskInfo.txt','files\DiskInfo.ini' -ErrorAction SilentlyContinue
    Write-Host 'Got SMART' -ForegroundColor Green
    Return $1,$2
}
function getTimer {
    $timer.Stop()
    $1 = 'Runtime'
    $2 = $timer.Elapsed | Select Minutes,Seconds | ConvertTo-Html -Fragment
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
$timer = [diagnostics.stopwatch]::StartNew()
# ------------------ #

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
$key = $(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' | Select-Object -ExpandProperty 'DigitalProductId')
$installedBase0 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | ? {$_.DisplayName -notlike $null}
$installedBase1 = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | ? {$_.DisplayName -notlike $null}
$installedBase = $installedBase0 + $installedBase1
$services = $(Get-Service)
$runningProcesses = Get-Process
$volumes = Get-Volume
$dns = Get-DnsClientGlobalSetting
Write-Host 'Got main data' -ForegroundColor Green

# ------------------ #
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
getTimer | Out-File -Append -Encoding ascii $file
promptUpload
# ------------------ #
