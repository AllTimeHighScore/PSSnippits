#############
# - Check BIOS and Audio Drivers
# - 
# - Created by: Kevin Van Bogart
#############
#Name specific of the driver
$AudioDriverName = 'Conexant'

#Surface Audio Drivers
#Audio Endpoint
#Realtek High Definition Audio
#High Definition Audio Controller

#Check Audio Driver
$AudioDriver = Get-CimInstance -ClassName Win32_PnPSignedDriver -ErrorAction SilentlyContinue | 
    Select-Object devicename, driverversion |
        Where-Object {$_.DeviceName -match $AudioDriverName} 

#Get Bios version
$BIOS = Get-CimInstance -ClassName WIN32_BIOS -ErrorAction SilentlyContinue

$CompSystem = Get-CimInstance -ClassName Win32_ComputerSystem
$OS = Get-CimInstance -ClassName Win32_OperatingSystem

[pscustomobject]@{
    SystemType = $CompSystem.Model
    SystemName = $CompSystem.Name
    OperatingSystem = $OS.Version
    BiosVersion = $BIOS.SMBIOSBIOSVersion
    AudioDriver = $AudioDriver.driverversion
    AudioDriverName = $AudioDriver.devicename
}