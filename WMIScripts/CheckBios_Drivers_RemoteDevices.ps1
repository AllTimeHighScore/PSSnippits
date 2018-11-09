

$YourMom = ((New-Guid).Guid.split('-')[0])
$TargetLaptops = Get-Content -path "C:\TEMP\Win10Laptops.csv"

$TargetLaptops | % {Invoke-Command -ComputerName $_ -SessionOption (New-PSSessionOption -NoMachineProfile) -ScriptBlock {

        #############
        # - Check BIOS and Audio Drivers
        # - 
        # - Created by: Kevin Van Bogart
        #############
        Try {
            #Name specific of the driver
            $AudioDriverName = 'Conexant'
            
            #Surface Audio Drivers
            #Audio Endpoint
            #Realtek High Definition Audio
            #High Definition Audio Controller
            
            $Handles = Get-CimInstance -ClassName win32_Process -ErrorAction Stop | ? {$_.Name -match 'MicTray64'} | Select-Object 'HandleCount'
            
            #Check Audio Driver
            $AudioDriver = Get-CimInstance  -ClassName Win32_PnPSignedDriver -ErrorAction Stop | 
                Select-Object devicename, driverversion |
                    Where-Object {$_.DeviceName -match $AudioDriverName} 
            
            #Get Bios version
            $BIOS = Get-CimInstance -ClassName WIN32_BIOS -ErrorAction SilentlyContinue
            
            $CompSystem = Get-CimInstance -ClassName Win32_ComputerSystem
            $OS = Get-CimInstance -ClassName Win32_OperatingSystem
            
             [pscustomobject]@{
               SystemName = $CompSystem.Name
               SystemType = $CompSystem.Model
               OperatingSystem = $OS.Version
               BiosVersion = $BIOS.SMBIOSBIOSVersion
               AudioDriverversion = $AudioDriver.driverversion
               AudioDriverName = $AudioDriver.devicename
               HandleCount = $Handles.HandleCount
           }# | ConvertTo-Csv -Delimiter "`t" | Out-File -FilePath C:\temp\Laptop_Check_11082018.csv -Append -Force
    
        }
        Catch {
            Write-Warning -Message "Problem connecting to $_"
        }
    } -ErrorAction SilentlyContinue -AsJob -JobName $YourMom
}

$WorkDone 

