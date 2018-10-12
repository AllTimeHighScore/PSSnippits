#############
# - Check drive space
# - I need to revisit this script I'm not a terribly big fan of it
#############

#Check all drives 1 by 1 for required disk space
Get-CimInstance -ClassName Win32_logicalDisk -filter "DriveType = 3" -ErrorAction SilentlyContinue | ForEach-Object {
    $FreeSpace = $($_.FreeSpace)/1MB

    #As solidification only happen to system drive due to cluster service, check if appropreate space (min 180 MB) in system drive exist
    if ((Get-Service -Name "ClusSvc" -ErrorAction SilentlyContinue) -and (($_.DeviceID -eq $env:SystemDrive) -and ($FreeSpace -ge 180))){
        "Required space found in system drive $($_.DeviceID). Installation will continue"
    }
    #If Current drive is system drive, check if min 180 MB disk space available. else exit the process
    elseif ((($_.DeviceID -eq $env:SystemDrive) -and ($FreeSpace -ge 180)) -or (($_.DeviceID -ne $env:SystemDrive) -and ($FreeSpace -ge 80))){
        "Required space found in drive $($_.DeviceID). Installation will continue"
    }
    else {
        Write-Warning -Message "Required disk space not found in $($_.DeviceID). Installation will exit with return code 33004"
        "Required disk space not found $($_.DeviceID). Installation will exit with return code 33004"
        Exit 1 #Add custom code here
    }
} #End foreach