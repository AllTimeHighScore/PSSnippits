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
##$AudioDriver = Get-CimInstance -ClassName Win32_PnPSignedDriver -ErrorAction SilentlyContinue | 
##    Select-Object devicename, driverversion |
##        Where-Object {$_.DeviceName -match $AudioDriverName} 
$AudioDriver = Get-CimInstance -ClassName Win32_SoundDevice -PipelineVariable AudioDriver | 
    Where-Object {($AudioDriver.PNPDeviceID -notmatch "USB") -and ($AudioDriver.Name -match "Conexant|Realtek")} | 
        ForEach-Object {Get-CimInstance -ClassName Win32_PnPSignedDriver -ErrorAction SilentlyContinue | 
            Select-Object devicename, driverversion |
                Where-Object {$_.DeviceName -match $AudioDriver.Name}
        }


#Get Bios version
$BIOS = Get-CimInstance -ClassName WIN32_BIOS -ErrorAction SilentlyContinue

$CompSystem = Get-CimInstance -ClassName Win32_ComputerSystem
$OS = Get-CimInstance -ClassName Win32_OperatingSystem

$ConexantHandles = Get-CimInstance -ClassName win32_Process -ErrorAction Stop | ? {$_.Name -match 'MicTray64'} | Select-Object 'HandleCount'

#$PnPAudio = Get-CimInstance -ClassName Win32_PnPSignedDriver | Where-Object {($_.Manufacturer -match "Conexant|Realtek") -and ($_.HardWareID -notmatch 'USB')}
#$Proc = (Get-Item -Path "C:\Windows\inf\$($PnPAudio.InfName)" | Get-Content | Where-Object {$_ -match '\%AUTORUN\%'}).split(',').split(' ').split('\').where({$_ -match ".exe"})
#$Process = (get-process -Name $Proc.split('.')[0])

$TopHandles = Get-Process | Sort-Object {$_.HandleCount} | Select-Object -Property ProcessName,Handles | Select-Object -Last 10

$DeviceInspection = [pscustomobject]@{
    SystemType = $CompSystem.Model
    SystemName = $CompSystem.Name
    OperatingSystem = $OS.Version
    BiosVersion = $BIOS.SMBIOSBIOSVersion
    AudioDriver = $AudioDriver.driverversion
    AudioDriverName = $AudioDriver.devicename
    TopHandles = $TopHandles
    Mictray = $ConexantHandles.HandleCount
    LastBootTime = $OS.LastBootUpTime
    CurrentLocalTime = $OS.LocalDateTime
}

$DeviceInspection

$DeviceInspection.TopHandles | format-table *