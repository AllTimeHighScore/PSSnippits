################################
# - An attempt to grab items that are being copied locally from the web for troubleshooting purposes 
# - This could be changed to initiate the web request and by a tad more dynamic
# - Origin - This was an attempt to catch what Adobe was allowing for downloads in China before we realized there was more going on.
# - Author: Kevin Van Bogart
################################

$Location2Watch = ##############Download Directory##################
$BetaLocal = ############Copy 2 Location############
$LocalLogFile = ############Log FileLocation############
$Timeout = 2400
$SourceExists = $false
$Times = 0


if ((Test-Path -LiteralPath $BetaLocal) -eq $false){New-Item -ItemType 'Directory' -Path $BetaLocal}

do {
    #Get contents of the backup directory
    $Backup = Get-ChildItem -Path $BetaLocal -Recurse -ErrorAction SilentlyContinue
 
    #Only go into the actions if the folder exists
    if (Test-Path -Path $Location2Watch){
        
        #Obtain the contents of the download folder so we can inspect and copy them 
        Get-ChildItem -Path $Location2Watch -Recurse -PipelineVariable DlItem -ErrorAction SilentlyContinue | ForEach-Object {
            
            #Indicate the folder existed
            $SourceExists = $true
            
            #See if the item exists yet.
            if ($BetaLocal.Name -notcontains $DlItem.name){
                try {
                    Copy-Item -Path $DlItem.FullName -Destination $BetaLocal -Recurse -Force -PassThru -ErrorAction Stop | Out-File -FilePath $LocalLogFile -Force -Append
                    "Succesfully copied item $($DlItem.Name) from $Location2Watch to $BetaLocal" | Out-File -FilePath $LocalLogFile -Force -Append
                }
                catch {
                    Write-Warning -Message "Error: Could not copy item $($DlItem.Name) from $Location2Watch to $BetaLocal"
                }
            }
            #If it's a file and it was already copied, check to see if the filesize matches the one that was downloaded
            elseif (((Get-Item -Path $DlItem.FullName).Length -lt (($Backup | Where-Object {$_.name -eq $DlItem.Name}).length)) -and ((Get-Item -Path $DlItem.FullName).Attributes -ne 'Directory')){
                "Item $($DlItem.Name) that was copied to $Betalocal is not the same size as the original in $Location2Watch. Will attempt to grab a fresh copy and overwrite."

                try {
                    Copy-Item -Path $DlItem.FullName -Destination $BetaLocal -Container -Force -ErrorAction Stop | Out-File -FilePath $LocalLogFile -Force -Append 
                    "Succesfully copied item $($DlItem.Name) from $Location2Watch to $BetaLocal" | Out-File -FilePath $LocalLogFile -Force -Append
                }
                catch {
                    Write-Warning -Message "Error: Could not update item $($DlItem.Name) from $Location2Watch to $BetaLocal" | Out-File -FilePath $LocalLogFile -Force -Append
                }
            }#End Elseif
        }#end Get-Childitem loop
    }#end test path

    #Set up to dump out of the loop should the source folder be deleted
    if (($SourceExists -eq $true) -and (Test-Path -Path $Location2Watch) -eq $false){
        $Times = $Timeout
        "Warning: The source folder has been removed. Exiting Script" | Out-File -FilePath $LocalLogFile -Force -Append
    }
    #if the directory exists, let's slow down the loop speed and let the files copy in
    elseif ($SourceExists -eq $true){
        Start-Sleep -Milliseconds 25
        $Times++
    }
    #The directory hasn't been created yet; increase the frequency of checks.
    else {
        Start-Sleep -Milliseconds 1
        $Times++
    }

} While ($Times -lt $Timeout)