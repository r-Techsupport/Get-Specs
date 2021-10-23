# Get-Specs Script
## Get Windows Specifications

This repo contains the PS1 and files needed to run the Get-Specs application made for rTechsupport

The PS1 or EXE must be run as admin to gather SMART, and to gather temperatures. 

This script gathers:


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

### Networking
* Network adapters
* IPconfig

### Devices
* Device manager
* Audio devices

### Disks
* Disk layout information
* SMART report from CDI using [Get-Smart](https://git.dev0.sh/piper/Get-Smart)
