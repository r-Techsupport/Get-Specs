# Get-Specs Script
## Get Windows Specifications

This repo contains the PS1 and files needed to run the Get-Specs application made for rTechsupport

`Loader.ps1` is used to call `Get-Specs.ps1` and is compiled into an .exe using `compile.ps1`

The PS1 or EXE must be run as admin to gather SMART, and to gather temperatures.

## CLI
You can execute `Get-Specs.ps1` from the command line. `./get-specs.ps1 -run [ -view | -upload ]`. 

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
