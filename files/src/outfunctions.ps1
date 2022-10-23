function header {
$1 = '<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />

<style>
* {
    font-family: verdana !important;
    font-size: 12px;
}
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
#topbutton{
  opacity: 80%;
  width: 5%;
  padding-top: -3%;
  background-color: #ccc;
  position: fixed;
  bottom: 0;
  right: 0;
  border-radius: 20px;
  text-align: center;
  font-size: 24px;
  color: #87ab63;
}
</style>
<body>
<pre>
<a href="#top"><div id="topbutton">
TOP
</div></a>
'
Return $1
}

function table {
$1 = '<a name="top"></a>
<h2>Sections</h2>
<div style="line-height:0">
<p><a href="#hw">Hardware Basics</a></p>
<p><a href="#SecInfo">Security Information</a></p>
<p><a href="#bios">BIOS</a></p>
<p><a href="#SysVar">System Variables</a></p>
<p><a href="#UserVar">User Variables</a></p>
<p><a href="#hotfixes">Installed updates</a></p>
<p><a href="#StartupTasks">Startup Tasks</a></p>
<p><a href="#Power">Powerprofiles</a></p>
<p><a href="#RunningProcs">Running Processes</a></p>
<p><a href="#Services">Services</a></p>
<p><a href="#InstalledApps">Installed Applications</a></p>
<p><a href="#NetConfig">Network Configuration</a></p>
<p><a href="#NetConnections">Network Connections</a></p>
<p><a href="#Drivers">Drivers and device versions</a></p>
<p><a href="#usbDevices">USB Devices</a></p>
<p><a href="#issueDevices">Devices with issues</a></p>
<p><a href="#Audio">Audio Devices</a></p>
<p><a href="#Disks">Disk Layouts</a></p>
<p><a href="#SMART">SMART</a></p>
</div>'
Return $1
}

function uploadFile {
    $link = Invoke-WebRequest -ContentType 'text/plain' -Method 'PUT' -InFile $file -Uri "https://paste.rtech.support/upload/$null.html" -UseBasicParsing
    $linkProper = $link.Content -Replace "support","support/selif"
    set-clipboard $linkProper
    Start-Process $linkProper
}

function promptStart {  
    If (!($run.IsPresent)) {
        . .\wpf.ps1
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
    The source code for this application can be found at https://github.com/r-techsupport/Get-Specs"
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
            Break
        }
    }
}

function promptUpload {
    If (!($upload.IsPresent) -and !($view.IsPresent)) {
        . .\wpf.ps1
        Write-Host "Completed. Check for another prompt offering to view or upload the results."
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
    } ElseIf ($upload.IsPresent) {
        uploadFile
        Write-Host "Link has been copied to your clipboard."
    } ElseIf ($view.IsPresent) {
        Invoke-Item $file
    }
}

