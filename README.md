# Get-Specs Script

## This repo has been archived. It has been superseeded by [Specify](https://github.com/Spec-ify), another r/TS affiliated project.

This was used for 5 years in the r/Techsupport live chat and during that time I learned a great deal about PowerShell and how it can be ~~used~~ abused. I leave this up for anyone looking to reference the methods used to gather Windows specifications and the creation of valid HTML reports in PowerShell.

## Get Windows Specifications

This repo contains the PS1 and files needed to run the Get-Specs application made for rTechsupport

Running `Specs.cmd` (the `.cmd` extension may not show up on all systems) will execute the script interactively as admin.

## CLI
You can also execute `Get-Specs.ps1` from the command line. `./get-specs.ps1 -run [ -view | -upload ]`. 

## What is reported
### Basics

* Windows edition
* Build #
* OS install date
* Uptime
* Hostname
* Domain

### Security Information
* AV product
* Firewall product
* UAC Status
* SecureBoot Status

### Hardware
* CPU model and temperature
* Motherboard brand and model
* Graphics card model and temperature
* Amount of RAM
* Ram model and capacities

### System Information
* System variables
* User variables
* Hotfix list
* User startup tasks
* Running processes with statistics
* Services
* Installed applications
* Installed Chrome extensions
* Integrity of Hosts file

### Networking
* Network adapters
* IPconfig
* Active network connections

### Devices
* Device manager
* Audio devices

### Disks
* Disk layout information
* SMART report from CDI using [Get-Smart](https://github.com/r-Techsupport/Get-SMART)
